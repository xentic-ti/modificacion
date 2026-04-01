/* eslint-disable */
// @ts-nocheck
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { fillAndAttachFromFolder } from './documentFillAndAttach.service';
import { listFilesRecursive } from './spFolderExplorer.service';
import { addListItem, getAllItems, spGetJson, updateListItem } from './sharepointRest.service';
import { descargarReporteFase1Word, IFase1WordReportRow } from '../utils/fase1WordReportExcel';
import { openExcelRevisionSession } from '../utils/modificacionExcelHelper';

type LogFn = (s: string) => void;

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

function cleanPart(value: string): string {
  const text = String(value || '').trim();
  if (!text) return '';
  const lower = text.toLowerCase();
  if (text === '-' || text === '—' || lower === 'na' || lower === 'n/a' || lower === 'null') {
    return '';
  }
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
  const parts = [a, b, c, d]
    .map(cleanPart)
    .filter(Boolean)
    .map(sanitizeFolderPart);

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
  const minorText = match[2] || '0';
  const minor = parseInt(minorText, 10) + 1;
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
  if (key) {
    set.add(key);
  }
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
    if (key && !map.has(key)) {
      map.set(key, items[i].Id);
    }
    if (compact && !map.has(compact)) {
      map.set(compact, items[i].Id);
    }
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

async function buildProcesoCorporativoMap(
  context: WebPartContext,
  webUrl: string
): Promise<Map<string, number>> {
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
    if (key) {
      map.set(key, item.Id);
    }
  }

  return map;
}

function resolveLookupId(map: Map<string, number>, value: any): number | undefined {
  const normalized = normKey(value);
  if (normalized && map.has(normalized)) {
    return map.get(normalized);
  }

  const compact = compactKey(value);
  if (compact && map.has(compact)) {
    return map.get(compact);
  }

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
  const match = raw.match(
    /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
  );
  if (!match) {
    return null;
  }

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3].length === 2 ? `20${match[3]}` : match[3]);
  return parseDatePartsToIso(year, month, day);
}

function excelSerialToIso(serialValue: number): string {
  const serialText = String(serialValue);
  const wholeDays = Number(serialText.split('.')[0] || serialText);
  const parsed = new Date(Date.UTC(1899, 11, 30 + wholeDays, 0, 0, 0));
  if (isNaN(parsed.getTime())) {
    throw new Error('Fecha serial de Excel inválida.');
  }

  return parsed.toISOString();
}

function normalizeExcelDateForSharePoint(value: any, fieldLabel: string, documentName: string): string {
  if (value === null || value === undefined) {
    throw new Error(`El Excel no tiene ${fieldLabel} para "${documentName}".`);
  }

  if (typeof value === 'number' && !isNaN(value)) {
    return excelSerialToIso(value);
  }

  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0)).toISOString();
  }

  const raw = String(value).trim();
  if (!raw) {
    throw new Error(`El Excel no tiene ${fieldLabel} para "${documentName}".`);
  }

  const normalizedRaw = raw.replace(',', '.');

  const looseSlashDate = parseLooseSlashDateToIso(normalizedRaw);
  if (looseSlashDate) {
    return looseSlashDate;
  }

  const ymd = normalizedRaw.match(
    /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[T\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
  );
  if (ymd) {
    return parseDatePartsToIso(Number(ymd[1]), Number(ymd[2]), Number(ymd[3]));
  }

  const serial = Number(normalizedRaw);
  if (!isNaN(serial) && serial > 0 && /^\d+(\.\d+)?$/.test(normalizedRaw)) {
    return excelSerialToIso(serial);
  }

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
  const impactNamesRaw = (excelRow.areasImpactadasTxt || '')
    .split(/[\n;,|/]+/g)
    .map((item) => item.trim())
    .filter(Boolean);

  const impactSet = new Set<string>();
  for (let i = 0; i < impactNamesRaw.length; i++) {
    addImpactNameUnique(impactSet, impactNamesRaw[i]);
  }

  if (isSi(excelRow.flagConducta)) {
    addImpactNameUnique(impactSet, 'Impacto en Conducta de Mercado');
  }

  if (isSi(excelRow.flagExperiencia)) {
    addImpactNameUnique(impactSet, 'Impacto en Experiencia del Cliente');
  }

  return Array.from(impactSet);
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
  if (!excelRow.nombreDocumento) {
    throw new Error('El Excel no tiene NombreDocumento para crear la nueva solicitud.');
  }

  if (!excelRow.categoriaDocumento) {
    throw new Error(`El Excel no tiene CategoriaDocumento para "${excelRow.nombreDocumento}".`);
  }

  if (!excelRow.resumen) {
    throw new Error(`El Excel no tiene Resumen para "${excelRow.nombreDocumento}".`);
  }

  if (!excelRow.fechaAprobacion) {
    throw new Error(`El Excel no tiene FechaDeAprobacion para "${excelRow.nombreDocumento}".`);
  }

  if (!excelRow.fechaVigencia) {
    throw new Error(`El Excel no tiene FechaDeVigencia para "${excelRow.nombreDocumento}".`);
  }

  if (!lookups.tipoDocumentoId) {
    throw new Error(
      `No se encontró TipoDocumento en lookup para "${excelRow.nombreDocumento}". Valor Excel="${excelRow.tipoDocumento}".`
    );
  }

  if (!lookups.procesoDeNegocioId) {
    throw new Error(`No se encontró ProcesoDeNegocio en lookup para "${excelRow.nombreDocumento}".`);
  }

  if (!lookups.areaDuenaId) {
    throw new Error(`No se encontró AreaDuena en lookup para "${excelRow.nombreDocumento}".`);
  }

  const payload: any = {
    Title: excelRow.nombreDocumento,
    Accion: 'Actualización de documento',
    NombreDocumento: excelRow.nombreDocumento,
    CategoriadeDocumento: excelRow.categoriaDocumento,
    ResumenDocumento: excelRow.resumen,
    FechaDeAprobacionSolicitud: normalizeExcelDateForSharePoint(
      excelRow.fechaAprobacionRaw,
      'FechaDeAprobacion',
      excelRow.nombreDocumento
    ),
    FechadeVigencia: normalizeExcelDateForSharePoint(
      excelRow.fechaVigenciaRaw,
      'FechaDeVigencia',
      excelRow.nombreDocumento
    ),
    FechaDePublicacionSolicitud: new Date().toISOString(),
    FechadeEnvio: new Date().toISOString(),
    VersionDocumento: versionDocumento,
    EsVersionActualDocumento: true,
    DocumentosApoyo: false,
    CodigoDocumento: oldSolicitud.CodigoDocumento || ''
  };

  if (lookups.tipoDocumentoId) payload.TipoDocumentoId = lookups.tipoDocumentoId;
  if (lookups.procesoDeNegocioId) payload.ProcesoDeNegocioId = lookups.procesoDeNegocioId;
  if (lookups.areaDuenaId) payload.AreaDuenaId = lookups.areaDuenaId;
  if (oldSolicitud.EstadoId) payload.EstadoId = oldSolicitud.EstadoId;
  if (lookups.instanciaAprobacionId) payload.InstanciasdeaprobacionId = lookups.instanciaAprobacionId;
  if (lookups.impactAreaIds.length) {
    payload.AreasImpactadasId = lookups.impactIsMulti
      ? lookups.impactAreaIds
      : lookups.impactAreaIds[0];
  }

  return payload;
}

export async function ejecutarFase1DocumentosSinHijosNiFlujos(params: {
  context: WebPartContext;
  excelFile: IFilePickerResult;
  sourceFolderServerRelativeUrl: string;
  tempWordBaseFolderServerRelativeUrl: string;
  log?: LogFn;
}): Promise<{
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
  if (!grid.length) {
    throw new Error('El Excel está vacío.');
  }

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
  const reportRows: IFase1WordReportRow[] = [];

  for (let rowIndex = 1; rowIndex < grid.length; rowIndex++) {
    const row = grid[rowIndex] || [];
    const rawRow = rawGrid[rowIndex] || row;
    const excelRow = parseExcelRow(row, rawRow);
    const solicitudOrigenId = Number(row[idxSolicitud] || 0);
    const nombreArchivo = excelRow.nombreArchivo;
    const nombreDocumento = excelRow.nombreDocumento;
    const documentoPadre = excelRow.documentoPadre;
    const versionExcel = excelRow.versionDocumento;
    const duenoDocumento = excelRow.duenoDocumento;

    const tieneHijos = !isEmptyLike(row[idxHijos]);
    const tieneFlujos = !isEmptyLike(row[idxFlujos]);
    const tienePadre = !isEmptyLike(documentoPadre);

    if (!solicitudOrigenId || !nombreDocumento) {
      continue;
    }

    if (tieneHijos || tieneFlujos || tienePadre) {
      reportRows.push({
        SolicitudOrigenID: solicitudOrigenId,
        SolicitudID: '',
        NombreDocumento: nombreDocumento,
        NombreArchivo: nombreArchivo,
        CodigoDocumento: '',
        VersionDocumento: versionExcel,
        TieneDocumentoPadre: tienePadre ? 'Sí' : 'No',
        DocumentoPadreNombre: tienePadre ? documentoPadre : '',
        DocumentoPadreSolicitudID: '',
        PadreRegeneradoConLinks: 'No',
        RutaTemporalWord: '',
        EstadoFase1: 'SKIP_NO_APLICA',
        Error: 'Documento omitido por tener hijos, flujos o documento padre.'
      });
      continue;
    }

    try {
      const oldSolicitud = await getSolicitudById(params.context, webUrl, solicitudOrigenId);
      const versionNueva = incrementVersion(versionExcel || oldSolicitud.VersionDocumento || '1.0');
      const procesoDeNegocioKey = normKey(
        [cleanPart(excelRow.clasificacion), cleanPart(excelRow.macroproceso), cleanPart(excelRow.proceso), cleanPart(excelRow.subproceso)]
          .filter(Boolean)
          .join('/')
      );
      const impactNames = buildImpactNames(excelRow);
      const impactAreaIds = impactNames
        .map((name) => mapImpact.get(name))
        .filter((value): value is number => !!value);
      const tipoDocumentoId = resolveLookupId(mapTipoDoc, excelRow.tipoDocumento);
      const procesoDeNegocioId = mapProceso.get(procesoDeNegocioKey);
      const areaDuenaId = resolveLookupId(mapArea, excelRow.areaDuena);
      const instanciaAprobacionId = resolveLookupId(mapInst, excelRow.instanciaAprobacion);
      const instanciaAprobacionDoc = instanciaAprobacionId ? excelRow.instanciaAprobacion : 'Gerencia de Área';

      if (oldSolicitudIds.indexOf(solicitudOrigenId) === -1) {
        oldSolicitudIds.push(solicitudOrigenId);
      }

      if (!impactIsMulti && impactAreaIds.length > 1) {
        throw new Error(`El campo AreasImpactadas no admite múltiples valores para "${excelRow.nombreDocumento}".`);
      }

      log(
        `🗓️ Fechas Fase 1 | SolicitudOrigen=${solicitudOrigenId} | ` +
        `FechaAprobacionRaw="${formatLogValue(excelRow.fechaAprobacionRaw)}" | ` +
        `FechaVigenciaRaw="${formatLogValue(excelRow.fechaVigenciaRaw)}"`
      );

      const newSolicitudId = await addListItem(
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

      log(`📂 TEMP destino calculado | SolicitudOrigen=${solicitudOrigenId} | ${tempDestino}`);

      const attachResult = await fillAndAttachFromFolder({
        context: params.context,
        webUrl,
        listTitle: 'Solicitudes',
        itemId: newSolicitudId,
        originalFileName: nombreArchivo,
        fileByName,
        titulo: excelRow.nombreDocumento,
        instanciaRaw: instanciaAprobacionDoc,
        impactAreaIds,
        dueno: duenoDocumento,
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
        log
      });

      await updateListItem(params.context, webUrl, 'Solicitudes', solicitudOrigenId, {
        EsVersionActualDocumento: false
      });

      if (attachResult.tempFileServerRelativeUrl) {
        tempFileUrls.push(attachResult.tempFileServerRelativeUrl);
      }

      reportRows.push({
        SolicitudOrigenID: solicitudOrigenId,
        SolicitudID: newSolicitudId,
        NombreDocumento: excelRow.nombreDocumento,
        NombreArchivo: nombreArchivo,
        CodigoDocumento: oldSolicitud.CodigoDocumento || '',
        VersionDocumento: versionNueva,
        TieneDocumentoPadre: 'No',
        DocumentoPadreNombre: '',
        DocumentoPadreSolicitudID: '',
        PadreRegeneradoConLinks: 'No',
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
        MetadataPendiente: 'Sí'
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      reportRows.push({
        SolicitudOrigenID: solicitudOrigenId,
        SolicitudID: '',
        NombreDocumento: nombreDocumento,
        NombreArchivo: nombreArchivo,
        CodigoDocumento: '',
        VersionDocumento: versionExcel,
        TieneDocumentoPadre: 'No',
        DocumentoPadreNombre: '',
        DocumentoPadreSolicitudID: '',
        PadreRegeneradoConLinks: 'No',
        RutaTemporalWord: '',
        EstadoFase1: 'ERROR',
        Error: message
      });
      log(
        `❌ Error Fase 1 | SolicitudOrigen=${solicitudOrigenId} | ` +
        `FechaAprobacionRaw="${formatLogValue(excelRow.fechaAprobacionRaw)}" | ` +
        `FechaVigenciaRaw="${formatLogValue(excelRow.fechaVigenciaRaw)}" | ${message}`
      );
    }
  }

  descargarReporteFase1Word(reportRows);

  return {
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
