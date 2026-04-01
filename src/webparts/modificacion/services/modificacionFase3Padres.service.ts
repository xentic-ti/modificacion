/* eslint-disable */
// @ts-nocheck
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { fillAndAttachFromFolder } from './documentFillAndAttach.service';
import { listFilesRecursive } from './spFolderExplorer.service';
import { addListItem, getAllItems, getAttachmentFiles, spGetJson, updateListItem } from './sharepointRest.service';
import { descargarReporteFase1Word, IFase1WordReportRow } from '../utils/fase1WordReportExcel';
import { openExcelRevisionSession } from '../utils/modificacionExcelHelper';

type LogFn = (s: string) => void;

export interface IFase3RollbackEntry {
  solicitudOrigenId: number;
  solicitudNuevaId: number;
  nombreDocumento: string;
  hijosIds: number[];
  diagramasIds: number[];
}

type IExcelRowData = {
  clasificacion: string;
  macroproceso: string;
  proceso: string;
  subproceso: string;
  areaDuena: string;
  duenoDocumento: string;
  categoriaDocumento: string;
  tipoDocumento: string;
  nombreArchivo: string;
  nombreDocumento: string;
  documentoPadre: string;
  versionDocumento: string;
  fechaAprobacion: string;
  fechaAprobacionRaw: any;
  fechaVigencia: string;
  fechaVigenciaRaw: any;
  instanciaAprobacion: string;
  flagConducta: string;
  flagExperiencia: string;
  areasImpactadasTxt: string;
  resumen: string;
};

type IRelacionadoRow = { codigo: string; nombre: string; enlace: string; };
type IDiagramaRow = { id: number; codigo: string; nombre: string; enlace: string; };

function cleanPart(value: string): string {
  const text = String(value || '').trim();
  if (!text) return '';
  const lower = text.toLowerCase();
  if (text === '-' || text === '—' || lower === 'na' || lower === 'n/a' || lower === 'null') return '';
  if (lower === 'sin subproceso') return '';
  return text;
}

function sanitizeFolderPart(value: string): string {
  return String(value || '')
    .trim()
    .replace(/[~#%&*{}\\:<>?/+"|]/g, '')
    .replace(/\s+/g, ' ');
}

function buildDestinoWordTemp(baseFolder: string, a: string, b: string, c: string, d: string): string {
  const parts = [a, b, c, d].map(cleanPart).filter(Boolean).map(sanitizeFolderPart);
  const root = (baseFolder || '').replace(/\/$/, '');
  return parts.length ? `${root}/${parts.join('/')}` : root;
}

function isEmptyLike(value: any): boolean {
  const text = String(value ?? '').trim().toLowerCase();
  return !text || text === '-' || text === '—' || text === 'na' || text === 'n/a' || text === 'null';
}

function isSi(value: any): boolean {
  const text = String(value ?? '').trim().toLowerCase();
  return text === 'si' || text === 'sí' || text === 's';
}

function incrementVersion(value: any): string {
  const text = String(value ?? '').trim();
  if (!text) return '1.1';
  const match = text.match(/^(\d+)(?:\.(\d+))?$/);
  if (!match) return text;
  const major = parseInt(match[1], 10);
  const minor = parseInt(match[2] || '0', 10) + 1;
  return `${major}.${minor}`;
}

function normKey(value: any): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function compactKey(value: any): string {
  return normKey(value).replace(/[^a-z0-9]/g, '');
}

function addImpactNameUnique(set: Set<string>, name: string): void {
  const key = normKey(name);
  if (key) set.add(key);
}

function findColumnIndex(headers: any[], expected: string): number {
  const normalized = String(expected || '').trim().toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim().toLowerCase() === normalized) {
      return i;
    }
  }
  return -1;
}

async function buildLookupMapByField(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldInternalName: string
): Promise<Map<string, number>> {
  const field = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${listTitle}')/fields/getbyinternalnameortitle('${fieldInternalName}')?$select=LookupList`
  );

  const lookupListId = field.LookupList;
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists(guid'${lookupListId}')/items?$select=Id,Title&$top=5000`
  );

  const map = new Map<string, number>();
  for (let i = 0; i < items.length; i++) {
    const title = items[i].Title;
    const key = normKey(title);
    const compact = compactKey(title);
    if (key && !map.has(key)) map.set(key, items[i].Id);
    if (compact && !map.has(compact)) map.set(compact, items[i].Id);
  }

  return map;
}

async function getAllowMultipleValues(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldInternalName: string
): Promise<boolean> {
  const field = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${listTitle}')/fields/getbyinternalnameortitle('${fieldInternalName}')?$select=AllowMultipleValues`
  );
  return !!field.AllowMultipleValues;
}

async function getFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldTitleOrInternalName: string
): Promise<string> {
  const field = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${listTitle}')/fields/getbyinternalnameortitle('${fieldTitleOrInternalName}')?$select=InternalName`
  );
  return String(field?.InternalName || fieldTitleOrInternalName);
}

async function tryGetFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldTitleOrInternalName: string
): Promise<string | null> {
  try {
    return await getFieldInternalName(context, webUrl, listTitle, fieldTitleOrInternalName);
  } catch (_error) {
    return null;
  }
}

async function resolveFirstExistingFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  candidates: string[]
): Promise<string> {
  for (let i = 0; i < candidates.length; i++) {
    const resolved = await tryGetFieldInternalName(context, webUrl, listTitle, candidates[i]);
    if (resolved) return resolved;
  }
  throw new Error(`No se encontró el campo esperado en "${listTitle}": ${candidates.join(', ')}`);
}

async function buildProcesoCorporativoMap(context: WebPartContext, webUrl: string): Promise<Map<string, number>> {
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Procesos Corporativos')/items?$select=Id,Title,field_1,field_2,field_3&$top=5000`
  );
  const map = new Map<string, number>();
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const key = normKey(
      [cleanPart(item.Title), cleanPart(item.field_1), cleanPart(item.field_2), cleanPart(item.field_3)]
        .filter(Boolean)
        .join('/')
    );
    if (key) map.set(key, item.Id);
  }
  return map;
}

function resolveLookupId(map: Map<string, number>, value: any): number | undefined {
  const normalized = normKey(value);
  if (normalized && map.has(normalized)) return map.get(normalized);
  const compact = compactKey(value);
  if (compact && map.has(compact)) return map.get(compact);
  return undefined;
}

function parseDatePartsToIso(year: number, month: number, day: number): string {
  const parsed = new Date(Date.UTC(year, month - 1, day, 0, 0, 0));
  if (
    isNaN(parsed.getTime()) ||
    parsed.getUTCFullYear() !== year ||
    parsed.getUTCMonth() !== month - 1 ||
    parsed.getUTCDate() !== day
  ) {
    throw new Error('Fecha inválida.');
  }
  return parsed.toISOString();
}

function parseLooseSlashDateToIso(raw: string): string | null {
  const match = raw.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (!match) return null;
  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3].length === 2 ? `20${match[3]}` : match[3]);
  return parseDatePartsToIso(year, month, day);
}

function excelSerialToIso(serialValue: number): string {
  const wholeDays = Number(String(serialValue).split('.')[0] || serialValue);
  const parsed = new Date(Date.UTC(1899, 11, 30 + wholeDays, 0, 0, 0));
  if (isNaN(parsed.getTime())) throw new Error('Fecha serial de Excel inválida.');
  return parsed.toISOString();
}

function normalizeExcelDateForSharePoint(value: any, fieldLabel: string, documentName: string): string {
  if (value === null || value === undefined) throw new Error(`El Excel no tiene ${fieldLabel} para "${documentName}".`);
  if (typeof value === 'number' && !isNaN(value)) return excelSerialToIso(value);
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0)).toISOString();
  }
  const raw = String(value).trim();
  if (!raw) throw new Error(`El Excel no tiene ${fieldLabel} para "${documentName}".`);
  const normalizedRaw = raw.replace(',', '.');
  const looseSlashDate = parseLooseSlashDateToIso(normalizedRaw);
  if (looseSlashDate) return looseSlashDate;
  const ymd = normalizedRaw.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[T\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (ymd) return parseDatePartsToIso(Number(ymd[1]), Number(ymd[2]), Number(ymd[3]));
  const serial = Number(normalizedRaw);
  if (!isNaN(serial) && serial > 0 && /^\d+(\.\d+)?$/.test(normalizedRaw)) return excelSerialToIso(serial);
  const parsed = new Date(normalizedRaw);
  if (!isNaN(parsed.getTime())) {
    return new Date(Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate(), 0, 0, 0)).toISOString();
  }
  throw new Error(`Formato inválido en ${fieldLabel} para "${documentName}": "${raw}".`);
}

function formatLogValue(value: any): string {
  if (value === null) return 'null';
  if (value === undefined) return 'undefined';
  if (value instanceof Date) return isNaN(value.getTime()) ? 'Invalid Date' : value.toISOString();
  return String(value);
}

function parseExcelRow(row: any[], rawRow?: any[]): IExcelRowData {
  return {
    clasificacion: String(row[0] || '').trim(),
    macroproceso: String(row[1] || '').trim(),
    proceso: String(row[2] || '').trim(),
    subproceso: String(row[3] || '').trim(),
    areaDuena: String(row[4] || '').trim(),
    duenoDocumento: String(row[5] || '').trim(),
    categoriaDocumento: String(row[6] || '').trim(),
    tipoDocumento: String(row[7] || '').trim(),
    nombreArchivo: String(row[8] || '').trim(),
    nombreDocumento: String(row[9] || '').trim(),
    documentoPadre: String(row[10] || '').trim(),
    versionDocumento: String(row[11] || '').trim(),
    fechaAprobacion: String(row[12] || '').trim(),
    fechaAprobacionRaw: rawRow ? rawRow[12] : row[12],
    fechaVigencia: String(row[13] || '').trim(),
    fechaVigenciaRaw: rawRow ? rawRow[13] : row[13],
    instanciaAprobacion: String(row[14] || '').trim(),
    flagConducta: String(row[15] || '').trim(),
    flagExperiencia: String(row[16] || '').trim(),
    areasImpactadasTxt: String(row[17] || '').trim(),
    resumen: String(row[18] || '').trim()
  };
}

function buildImpactNames(excelRow: IExcelRowData): string[] {
  const impactNamesRaw = String(excelRow.areasImpactadasTxt || '')
    .split(/[\n;,|/]+/g)
    .map((item) => item.trim())
    .filter(Boolean);
  const impactSet = new Set<string>();
  for (let i = 0; i < impactNamesRaw.length; i++) addImpactNameUnique(impactSet, impactNamesRaw[i]);
  if (isSi(excelRow.flagConducta)) addImpactNameUnique(impactSet, 'Impacto en Conducta de Mercado');
  if (isSi(excelRow.flagExperiencia)) addImpactNameUnique(impactSet, 'Impacto en Experiencia del Cliente');
  return Array.from(impactSet);
}

function parseSlashIds(value: any): number[] {
  return String(value || '')
    .split('/')
    .map((part) => Number(String(part || '').trim()))
    .filter((num) => Number.isFinite(num) && num > 0);
}

async function getSolicitudById(context: WebPartContext, webUrl: string, solicitudId: number): Promise<any> {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items(${solicitudId})` +
    `?$select=Id,Title,NombreDocumento,CodigoDocumento,CategoriadeDocumento,ResumenDocumento,` +
    `FechaDeAprobacionSolicitud,FechadeVigencia,VersionDocumento,TipoDocumentoId,TipoDocumento/Title,` +
    `ProcesoDeNegocioId,ProcesoDeNegocio/Title,ProcesoDeNegocio/field_1,ProcesoDeNegocio/field_2,ProcesoDeNegocio/field_3,` +
    `AreaDuenaId,AreaDuena/Title,EstadoId,InstanciasdeaprobacionId,Instanciasdeaprobacion/Title,` +
    `AreasImpactadas/Id,AreasImpactadas/Title,Accion,DocumentosApoyo,EsVersionActualDocumento` +
    `&$expand=TipoDocumento,ProcesoDeNegocio,AreaDuena,Instanciasdeaprobacion,AreasImpactadas`;
  return spGetJson<any>(context, url);
}

function buildNewSolicitudPayload(
  oldSolicitud: any,
  excelRow: IExcelRowData,
  versionDocumento: string,
  lookups: {
    tipoDocumentoId?: number;
    procesoDeNegocioId?: number;
    areaDuenaId?: number;
    instanciaAprobacionId?: number;
    impactAreaIds: number[];
    impactIsMulti: boolean;
  }
): any {
  if (!lookups.tipoDocumentoId) {
    throw new Error(`No se encontró TipoDocumento en lookup para "${excelRow.nombreDocumento}". Valor Excel="${excelRow.tipoDocumento}".`);
  }
  if (!lookups.procesoDeNegocioId) throw new Error(`No se encontró ProcesoDeNegocio en lookup para "${excelRow.nombreDocumento}".`);
  if (!lookups.areaDuenaId) throw new Error(`No se encontró AreaDuena en lookup para "${excelRow.nombreDocumento}".`);

  const payload: any = {
    Title: excelRow.nombreDocumento,
    Accion: 'Actualización de documento',
    NombreDocumento: excelRow.nombreDocumento,
    CategoriadeDocumento: excelRow.categoriaDocumento,
    ResumenDocumento: excelRow.resumen,
    FechaDeAprobacionSolicitud: normalizeExcelDateForSharePoint(excelRow.fechaAprobacionRaw, 'FechaDeAprobacion', excelRow.nombreDocumento),
    FechadeVigencia: normalizeExcelDateForSharePoint(excelRow.fechaVigenciaRaw, 'FechaDeVigencia', excelRow.nombreDocumento),
    FechaDePublicacionSolicitud: new Date().toISOString(),
    FechadeEnvio: new Date().toISOString(),
    VersionDocumento: versionDocumento,
    EsVersionActualDocumento: true,
    DocumentosApoyo: false,
    CodigoDocumento: oldSolicitud.CodigoDocumento || ''
  };

  payload.TipoDocumentoId = lookups.tipoDocumentoId;
  payload.ProcesoDeNegocioId = lookups.procesoDeNegocioId;
  payload.AreaDuenaId = lookups.areaDuenaId;
  if (oldSolicitud.EstadoId) payload.EstadoId = oldSolicitud.EstadoId;
  if (lookups.instanciaAprobacionId) payload.InstanciasdeaprobacionId = lookups.instanciaAprobacionId;
  if (lookups.impactAreaIds.length) {
    payload.AreasImpactadasId = lookups.impactIsMulti ? lookups.impactAreaIds : lookups.impactAreaIds[0];
  }

  return payload;
}

async function getSolicitudesRelacionadas(
  context: WebPartContext,
  webUrl: string,
  ids: number[]
): Promise<IRelacionadoRow[]> {
  const uniqueIds = Array.from(new Set(ids.filter((id) => id > 0)));
  const rows: IRelacionadoRow[] = [];
  for (let i = 0; i < uniqueIds.length; i++) {
    const id = uniqueIds[i];
    try {
      const item = await spGetJson<any>(
        context,
        `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items(${id})?$select=Id,Title,NombreDocumento,CodigoDocumento`
      );
      const codigo = item.CodigoDocumento || '';
      const nombre = item.NombreDocumento || item.Title || '';
      rows.push({
        codigo,
        nombre,
        enlace: codigo
          ? `${new URL(webUrl).origin}/sites/SistemadeGestionDocumental/SitePages/verdocumento-vigente.aspx?codigodocumentosgd=${encodeURIComponent(codigo)}`
          : ''
      });
    } catch (_error) {
      // omitir hijo inválido sin romper la fila
    }
  }
  return rows;
}

async function getDiagramasFlujoRows(
  context: WebPartContext,
  webUrl: string,
  ids: number[]
): Promise<IDiagramaRow[]> {
  const uniqueIds = Array.from(new Set(ids.filter((id) => id > 0)));
  const listRoot = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Diagramas de Flujo')?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`
  );
  const rootUrl = listRoot?.RootFolder?.ServerRelativeUrl || '';

  const rows: IDiagramaRow[] = [];
  for (let i = 0; i < uniqueIds.length; i++) {
    const id = uniqueIds[i];
    try {
      const item = await spGetJson<any>(
        context,
        `${webUrl}/_api/web/lists/getbytitle('Diagramas de Flujo')/items(${id})?$select=Id,Title,Codigo`
      );
      const attachments = await getAttachmentFiles(context, webUrl, 'Diagramas de Flujo', id);
      const attachmentName = attachments.length ? attachments[0].FileName : '';
      const enlace = attachmentName
        ? `${new URL(webUrl).origin}${rootUrl}/Attachments/${id}/${attachmentName.split('/').map((p: string) => encodeURIComponent(p)).join('/')}`
        : '';
      rows.push({
        id,
        codigo: item.Codigo || '',
        nombre: item.Title || '',
        enlace
      });
    } catch (_error) {
      // omitir diagrama inválido sin romper la fila
    }
  }
  return rows;
}

async function updateRelacionesDocumentosPadre(
  context: WebPartContext,
  webUrl: string,
  oldParentSolicitudId: number,
  newParentSolicitudId: number,
  childIds: number[],
  log: LogFn
): Promise<number> {
  const wanted = new Set(childIds.filter((id) => id > 0));
  if (!wanted.size) return 0;

  const parentField = await getFieldInternalName(context, webUrl, 'Relaciones Documentos', 'DocumentoPadre');
  const childField = await getFieldInternalName(context, webUrl, 'Relaciones Documentos', 'DocumentoHijo');
  const parentFieldId = `${parentField}Id`;
  const childFieldId = `${childField}Id`;

  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Relaciones Documentos')/items?$select=Id,${parentFieldId},${childFieldId}&$top=5000&$filter=${parentFieldId} eq ${oldParentSolicitudId}`
  );

  let updated = 0;
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (!wanted.has(Number(item[childFieldId] || 0))) continue;
    await updateListItem(context, webUrl, 'Relaciones Documentos', item.Id, {
      [parentFieldId]: newParentSolicitudId
    });
    updated++;
  }

  log(`🔗 Relaciones actualizadas | PadreAnterior=${oldParentSolicitudId} | PadreNuevo=${newParentSolicitudId} | Registros=${updated}`);
  return updated;
}

async function updateDiagramasSolicitud(
  context: WebPartContext,
  webUrl: string,
  diagramIds: number[],
  newParentSolicitudId: number,
  log: LogFn
): Promise<number> {
  const ids = Array.from(new Set(diagramIds.filter((id) => id > 0)));
  const solicitudField = await getFieldInternalName(context, webUrl, 'Diagramas de Flujo', 'Solicitud');
  const solicitudFieldId = `${solicitudField}Id`;
  const solicitudIsMulti = await getAllowMultipleValues(context, webUrl, 'Diagramas de Flujo', solicitudField);
  let updated = 0;
  for (let i = 0; i < ids.length; i++) {
    await updateListItem(context, webUrl, 'Diagramas de Flujo', ids[i], {
      [solicitudFieldId]: solicitudIsMulti ? [newParentSolicitudId] : newParentSolicitudId
    });
    updated++;
  }
  log(`🧭 Diagramas actualizados | SolicitudNueva=${newParentSolicitudId} | Registros=${updated}`);
  return updated;
}

function toLookupIdArray(value: any): number[] {
  if (Array.isArray(value)) {
    return value.map((x) => Number(x)).filter((x) => Number.isFinite(x) && x > 0);
  }

  if (value && Array.isArray(value.results)) {
    return value.results.map((x: any) => Number(x)).filter((x: number) => Number.isFinite(x) && x > 0);
  }

  const single = Number(value);
  return Number.isFinite(single) && single > 0 ? [single] : [];
}

async function updateChildSolicitudesDocPadres(
  context: WebPartContext,
  webUrl: string,
  childSolicitudIds: number[],
  oldParentSolicitudId: number,
  newParentSolicitudId: number,
  log: LogFn
): Promise<number> {
  const childIds = Array.from(new Set((childSolicitudIds || []).filter((id) => Number(id) > 0)));
  if (!childIds.length) return 0;

  const docPadresField = await resolveFirstExistingFieldInternalName(
    context,
    webUrl,
    'Solicitudes',
    ['docpadres', 'DocPadres', 'DocumentoPadre']
  );
  const docPadresFieldId = `${docPadresField}Id`;
  const docPadresIsMulti = await getAllowMultipleValues(context, webUrl, 'Solicitudes', docPadresField);

  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items?$select=Id,${docPadresFieldId}&$top=5000`
  );
  const itemById = new Map<number, any>();
  for (let i = 0; i < items.length; i++) {
    itemById.set(Number(items[i].Id || 0), items[i]);
  }

  let updated = 0;
  for (let i = 0; i < childIds.length; i++) {
    const item = itemById.get(childIds[i]);
    if (!item) continue;

    const currentIds = toLookupIdArray(item[docPadresFieldId]);
    if (!currentIds.length) continue;

    let changed = false;
    const nextIds = currentIds.map((id) => {
      if (id === oldParentSolicitudId) {
        changed = true;
        return newParentSolicitudId;
      }
      return id;
    });

    if (!changed) continue;

    const deduped = Array.from(new Set(nextIds.filter((id) => id > 0)));
    await updateListItem(context, webUrl, 'Solicitudes', childIds[i], {
      [docPadresFieldId]: docPadresIsMulti ? deduped : (deduped[0] || null)
    });
    updated++;
  }

  log(`👨‍👧 Fase 3 | DocPadres actualizados en hijos | PadreAnterior=${oldParentSolicitudId} | PadreNuevo=${newParentSolicitudId} | Solicitudes=${updated}`);
  return updated;
}

function buildReportFileName(): string {
  const now = new Date();
  const pad = (value: number): string => String(value).padStart(2, '0');
  return (
    `Reporte_Fase3_WORD_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_` +
    `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.xlsx`
  );
}

export async function ejecutarFase3PadresConHijosYFlujos(params: {
  context: WebPartContext;
  excelFile: IFilePickerResult;
  sourceFolderServerRelativeUrl: string;
  tempWordBaseFolderServerRelativeUrl: string;
  log?: LogFn;
}): Promise<{
  rollbackEntries: IFase3RollbackEntry[];
  createdSolicitudIds: number[];
  oldSolicitudIds: number[];
  tempFileUrls: string[];
  reportRows: IFase1WordReportRow[];
  processed: number;
  ok: number;
  skipped: number;
  error: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const session = await openExcelRevisionSession(params.excelFile);
  const grid = session.grid || [];
  const rawGrid = session.rawGrid || [];
  if (!grid.length) throw new Error('El Excel está vacío.');

  const headers = grid[0] || [];
  const idxSolicitud = findColumnIndex(headers, 'ID Solicitud');
  const idxHijos = findColumnIndex(headers, 'Documentos hijos');
  const idxFlujos = findColumnIndex(headers, 'Diagrama de Flujos');
  if (idxSolicitud < 0 || idxHijos < 0 || idxFlujos < 0) {
    throw new Error('El Excel no contiene las columnas ID Solicitud, Documentos hijos y Diagrama de Flujos.');
  }

  const files = await listFilesRecursive(params.context, webUrl, params.sourceFolderServerRelativeUrl, log);
  const mapTipoDoc = await buildLookupMapByField(params.context, webUrl, 'Solicitudes', 'TipoDocumento');
  const mapArea = await buildLookupMapByField(params.context, webUrl, 'Solicitudes', 'AreaDuena');
  const mapInst = await buildLookupMapByField(params.context, webUrl, 'Solicitudes', 'Instanciasdeaprobacion');
  const mapImpact = await buildLookupMapByField(params.context, webUrl, 'Solicitudes', 'AreasImpactadas');
  const mapProceso = await buildProcesoCorporativoMap(params.context, webUrl);
  const impactIsMulti = await getAllowMultipleValues(params.context, webUrl, 'Solicitudes', 'AreasImpactadas');

  const fileByName = new Map<string, string>();
  for (let i = 0; i < files.length; i++) {
    fileByName.set(String(files[i].Name || '').toLowerCase(), files[i].ServerRelativeUrl);
  }

  const createdSolicitudIds: number[] = [];
  const oldSolicitudIds: number[] = [];
  const tempFileUrls: string[] = [];
  const rollbackEntries: IFase3RollbackEntry[] = [];
  const reportRows: IFase1WordReportRow[] = [];

  for (let rowIndex = 1; rowIndex < grid.length; rowIndex++) {
    const row = grid[rowIndex] || [];
    const rawRow = rawGrid[rowIndex] || row;
    const excelRow = parseExcelRow(row, rawRow);
    const solicitudOrigenId = Number(row[idxSolicitud] || 0);
    const hijosIds = parseSlashIds(row[idxHijos]);
    const flujosIds = parseSlashIds(row[idxFlujos]);
    const tieneHijos = hijosIds.length > 0;
    const tieneFlujos = flujosIds.length > 0;
    const tienePadre = !isEmptyLike(excelRow.documentoPadre);

    if (!solicitudOrigenId || !excelRow.nombreDocumento) {
      continue;
    }

    if (tienePadre || (!tieneHijos && !tieneFlujos)) {
      reportRows.push({
        SolicitudOrigenID: solicitudOrigenId,
        SolicitudID: '',
        NombreDocumento: excelRow.nombreDocumento,
        NombreArchivo: excelRow.nombreArchivo,
        CodigoDocumento: '',
        VersionDocumento: excelRow.versionDocumento,
        TieneDocumentoPadre: tienePadre ? 'Sí' : 'No',
        DocumentoPadreNombre: tienePadre ? excelRow.documentoPadre : '',
        DocumentoPadreSolicitudID: '',
        PadreRegeneradoConLinks: 'No',
        RutaTemporalWord: '',
        EstadoFase1: 'SKIP_NO_APLICA',
        Error: 'Documento omitido por no ser padre procesable en Fase 3.',
        DocumentosHijosIDs: hijosIds.join('/'),
        DocumentoPadreSolicitudAnteriorID: solicitudOrigenId,
        DiagramasFlujoNombres: ''
      });
      continue;
    }

    let newSolicitudId = 0;
    let oldSolicitudMarcadaNoVigente = false;
    let relacionesMovidas = false;
    let diagramasMovidos = false;
    let docPadresMovidos = false;

    try {
      const oldSolicitud = await getSolicitudById(params.context, webUrl, solicitudOrigenId);
      const versionNueva = incrementVersion(excelRow.versionDocumento || oldSolicitud.VersionDocumento || '1.0');
      const procesoDeNegocioKey = normKey(
        [cleanPart(excelRow.clasificacion), cleanPart(excelRow.macroproceso), cleanPart(excelRow.proceso), cleanPart(excelRow.subproceso)]
          .filter(Boolean)
          .join('/')
      );
      const impactNames = buildImpactNames(excelRow);
      const impactAreaIds = impactNames.map((name) => mapImpact.get(name)).filter((value): value is number => !!value);
      const tipoDocumentoId = resolveLookupId(mapTipoDoc, excelRow.tipoDocumento);
      const procesoDeNegocioId = mapProceso.get(procesoDeNegocioKey);
      const areaDuenaId = resolveLookupId(mapArea, excelRow.areaDuena);
      const instanciaAprobacionId = resolveLookupId(mapInst, excelRow.instanciaAprobacion);
      const instanciaAprobacionDoc = instanciaAprobacionId ? excelRow.instanciaAprobacion : 'Gerencia de Área';

      const relacionados = await getSolicitudesRelacionadas(params.context, webUrl, hijosIds);
      const diagramas = await getDiagramasFlujoRows(params.context, webUrl, flujosIds);

      if (oldSolicitudIds.indexOf(solicitudOrigenId) === -1) oldSolicitudIds.push(solicitudOrigenId);
      if (!impactIsMulti && impactAreaIds.length > 1) {
        throw new Error(`El campo AreasImpactadas no admite múltiples valores para "${excelRow.nombreDocumento}".`);
      }

      log(
        `🗓️ Fechas Fase 3 | SolicitudOrigen=${solicitudOrigenId} | ` +
        `FechaAprobacionRaw="${formatLogValue(excelRow.fechaAprobacionRaw)}" | ` +
        `FechaVigenciaRaw="${formatLogValue(excelRow.fechaVigenciaRaw)}"`
      );

      newSolicitudId = await addListItem(
        params.context,
        webUrl,
        'Solicitudes',
        buildNewSolicitudPayload(oldSolicitud, excelRow, versionNueva, {
          tipoDocumentoId,
          procesoDeNegocioId,
          areaDuenaId,
          instanciaAprobacionId,
          impactAreaIds,
          impactIsMulti
        })
      );

      createdSolicitudIds.push(newSolicitudId);

      const proceso = oldSolicitud.ProcesoDeNegocio || {};
      const tempDestino = buildDestinoWordTemp(
        params.tempWordBaseFolderServerRelativeUrl,
        excelRow.clasificacion || proceso.Title || '',
        excelRow.macroproceso || proceso.field_1 || '',
        excelRow.proceso || proceso.field_2 || '',
        excelRow.subproceso || proceso.field_3 || ''
      );

      log(`📂 TEMP destino calculado Fase 3 | SolicitudOrigen=${solicitudOrigenId} | ${tempDestino}`);

      const attachResult = await fillAndAttachFromFolder({
        context: params.context,
        webUrl,
        listTitle: 'Solicitudes',
        itemId: newSolicitudId,
        originalFileName: excelRow.nombreArchivo,
        fileByName,
        titulo: excelRow.nombreDocumento,
        instanciaRaw: instanciaAprobacionDoc,
        impactAreaIds,
        dueno: excelRow.duenoDocumento,
        fechaVigencia: excelRow.fechaVigencia,
        fechaAprobacion: excelRow.fechaAprobacion,
        resumen: excelRow.resumen,
        version: versionNueva,
        codigoDocumento: oldSolicitud.CodigoDocumento || '',
        categoriaDoc: excelRow.categoriaDocumento,
        tipoDocExcel: excelRow.tipoDocumento,
        esDocumentoApoyo: false,
        tempDestinoFolderServerRelativeUrl: tempDestino,
        replaceIfExists: true,
        relacionados,
        diagramasFlujo: diagramas.map((row) => ({
          codigo: row.codigo,
          nombre: row.nombre,
          enlace: row.enlace
        })),
        log
      });

      await updateListItem(params.context, webUrl, 'Solicitudes', solicitudOrigenId, {
        EsVersionActualDocumento: false
      });
      oldSolicitudMarcadaNoVigente = true;

      await updateRelacionesDocumentosPadre(params.context, webUrl, solicitudOrigenId, newSolicitudId, hijosIds, log);
      relacionesMovidas = true;
      await updateDiagramasSolicitud(params.context, webUrl, flujosIds, newSolicitudId, log);
      diagramasMovidos = true;
      await updateChildSolicitudesDocPadres(params.context, webUrl, hijosIds, solicitudOrigenId, newSolicitudId, log);
      docPadresMovidos = true;

      if (attachResult.tempFileServerRelativeUrl) tempFileUrls.push(attachResult.tempFileServerRelativeUrl);

      rollbackEntries.push({
        solicitudOrigenId,
        solicitudNuevaId: newSolicitudId,
        nombreDocumento: excelRow.nombreDocumento,
        hijosIds,
        diagramasIds: flujosIds
      });

      reportRows.push({
        SolicitudOrigenID: solicitudOrigenId,
        SolicitudID: newSolicitudId,
        NombreDocumento: excelRow.nombreDocumento,
        NombreArchivo: excelRow.nombreArchivo,
        CodigoDocumento: oldSolicitud.CodigoDocumento || '',
        VersionDocumento: versionNueva,
        TieneDocumentoPadre: 'No',
        DocumentoPadreNombre: '',
        DocumentoPadreSolicitudID: '',
        PadreRegeneradoConLinks: 'Sí',
        RutaTemporalWord: attachResult.tempFileServerRelativeUrl || '',
        EstadoFase1: attachResult.ok ? 'OK' : 'ERROR',
        Error: attachResult.error || '',
        TipoDocumento: excelRow.tipoDocumento,
        CategoriaDocumento: excelRow.categoriaDocumento,
        Clasificaciondeproceso: excelRow.clasificacion,
        Macroproceso: excelRow.macroproceso,
        Proceso: excelRow.proceso,
        Subproceso: excelRow.subproceso,
        AreaDuena: excelRow.areaDuena,
        AreaImpactada: impactNames.join(' / '),
        Resumen: excelRow.resumen,
        FechaDeAprobacion: excelRow.fechaAprobacion,
        FechaDeVigencia: excelRow.fechaVigencia,
        InstanciaDeAprobacionId: instanciaAprobacionId || '',
        MetadataPendiente: 'Sí',
        DocumentosHijosIDs: hijosIds.join('/'),
        DocumentosHijosNombres: relacionados.map((row) => row.nombre || '').filter(Boolean).join(' / '),
        DocumentoPadreSolicitudAnteriorID: solicitudOrigenId,
        DocumentoPadreSolicitudNuevaID: newSolicitudId,
        DiagramasFlujoNombres: diagramas.map((row) => row.nombre || '').filter(Boolean).join(' / ')
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);

      if (diagramasMovidos && newSolicitudId) {
        try {
          await updateDiagramasSolicitud(params.context, webUrl, flujosIds, solicitudOrigenId, log);
          log(`↩️ Rollback local Fase 3 | Diagramas restaurados a solicitud antigua: ${solicitudOrigenId}`);
        } catch (_e) {
          // sin acción
        }
      }

      if (docPadresMovidos && newSolicitudId) {
        try {
          await updateChildSolicitudesDocPadres(params.context, webUrl, hijosIds, newSolicitudId, solicitudOrigenId, log);
          log(`↩️ Rollback local Fase 3 | DocPadres restaurados a padre antiguo: ${solicitudOrigenId}`);
        } catch (_e) {
          // sin acción
        }
      }

      if (relacionesMovidas && newSolicitudId) {
        try {
          await updateRelacionesDocumentosPadre(params.context, webUrl, newSolicitudId, solicitudOrigenId, hijosIds, log);
          log(`↩️ Rollback local Fase 3 | Relaciones restauradas a padre antiguo: ${solicitudOrigenId}`);
        } catch (_e) {
          // sin acción
        }
      }

      if (oldSolicitudMarcadaNoVigente) {
        try {
          await updateListItem(params.context, webUrl, 'Solicitudes', solicitudOrigenId, {
            EsVersionActualDocumento: true
          });
          if (newSolicitudId) {
            await updateListItem(params.context, webUrl, 'Solicitudes', newSolicitudId, {
              EsVersionActualDocumento: false
            });
          }
          log(`↩️ Rollback local Fase 3 | Vigencia restaurada | Antigua=${solicitudOrigenId} | Nueva=${newSolicitudId || ''}`);
        } catch (_e) {
          // sin acción
        }
      }
      else if (newSolicitudId) {
        try {
          await updateListItem(params.context, webUrl, 'Solicitudes', newSolicitudId, {
            EsVersionActualDocumento: false
          });
          log(`↩️ Rollback local Fase 3 | Nueva solicitud desactivada: ${newSolicitudId}`);
        } catch (_e) {
          // sin acción
        }
      }

      reportRows.push({
        SolicitudOrigenID: solicitudOrigenId,
        SolicitudID: '',
        NombreDocumento: excelRow.nombreDocumento,
        NombreArchivo: excelRow.nombreArchivo,
        CodigoDocumento: '',
        VersionDocumento: excelRow.versionDocumento,
        TieneDocumentoPadre: 'No',
        DocumentoPadreNombre: '',
        DocumentoPadreSolicitudID: '',
        PadreRegeneradoConLinks: 'No',
        RutaTemporalWord: '',
        EstadoFase1: 'ERROR',
        Error: message,
        DocumentosHijosIDs: hijosIds.join('/'),
        DocumentoPadreSolicitudAnteriorID: solicitudOrigenId,
        DiagramasFlujoNombres: ''
      });
      log(
        `❌ Error Fase 3 | SolicitudOrigen=${solicitudOrigenId} | ` +
        `FechaAprobacionRaw="${formatLogValue(excelRow.fechaAprobacionRaw)}" | ` +
        `FechaVigenciaRaw="${formatLogValue(excelRow.fechaVigenciaRaw)}" | ${message}`
      );
    }
  }

  descargarReporteFase1Word(reportRows, buildReportFileName());

  return {
    rollbackEntries,
    createdSolicitudIds,
    oldSolicitudIds,
    tempFileUrls,
    reportRows,
    processed: reportRows.length,
    ok: reportRows.filter((row) => row.EstadoFase1 === 'OK').length,
    skipped: reportRows.filter((row) => row.EstadoFase1.indexOf('SKIP') === 0).length,
    error: reportRows.filter((row) => row.EstadoFase1 === 'ERROR').length
  };
}
