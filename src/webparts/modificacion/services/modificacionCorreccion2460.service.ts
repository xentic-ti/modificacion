/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';
import { AadHttpClient } from '@microsoft/sp-http';
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { fillAndAttachFromServerRelativeUrl } from './documentFillAndAttachFromUrl.service';
import { listFilesRecursive } from './spFolderExplorer.service';
import {
  addAttachment,
  addListItem,
  ensureFolderPath,
  escapeODataValue,
  getAllItems,
  getAttachmentFiles,
  recycleFile,
  spGetJson,
  spPostJson,
  updateListItem,
  uploadFileToFolder
} from './sharepointRest.service';
import { descargarReporteFase2Publicacion, IFase2PublicacionReportRow } from '../utils/fase2PublicacionReportExcel';
import { IFase8RollbackEntry, rollbackModificacionFase8 } from './modificacionFase8Rollback.service';

type LogFn = (s: string) => void;
type IRelacionadoRow = { solicitudId: number; codigo: string; nombre: string; enlace: string; };
type IDiagramaRow = { id: number; codigo: string; nombre: string; enlace: string; };

const PROCESOS_ROOT = '/sites/SistemadeGestionDocumental/Procesos';
const HISTORICOS_ROOT = '/sites/SistemadeGestionDocumental/Documentos Histricos';
const TEMP_WORD_ROOT = '/sites/SistemadeGestionDocumental/Procesos/TEMP_MIGRACION_WORD';

type IFase8ExcelRow = {
  nombreArchivo: string;
  nombreDocumento: string;
  documentoPadre: string;
};

function normalizeHeader(value: any): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '')
    .trim()
    .toLowerCase();
}

function normalizeLooseText(value: any): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function normalizeLooseFileKey(value: any): string {
  const raw = String(value ?? '').trim();
  const extensionMatch = raw.match(/(\.[^.]+)$/);
  const extension = extensionMatch ? extensionMatch[1].toLowerCase() : '';
  const baseName = extension ? raw.slice(0, -extension.length) : raw;
  const normalized = normalizeLooseText(baseName).replace(/ /g, '');
  return `${normalized}${extension}`;
}

function trimSlash(value: string): string {
  return String(value || '').replace(/\/+$/, '');
}

function joinFolder(base: string, relative: string): string {
  const cleanBase = trimSlash(base);
  const cleanRelative = String(relative || '').replace(/^\/+/, '').replace(/\/+$/, '');
  return cleanRelative ? `${cleanBase}/${cleanRelative}` : cleanBase;
}

function sanitizeVersion(value: any): string {
  return String(value || '').trim().replace(/^v/i, '') || '1.0';
}

function incrementVersion(value: any): string {
  const text = sanitizeVersion(value);
  const match = text.match(/^(\d+)(?:\.(\d+))?$/);
  if (!match) return text;
  return `${parseInt(match[1], 10)}.${parseInt(match[2] || '0', 10) + 1}`;
}

function buildTodayStamp(now: Date): string {
  const pad = (value: number): string => String(value).padStart(2, '0');
  return `${pad(now.getDate())}${pad(now.getMonth() + 1)}${now.getFullYear()}`;
}

function buildTodayDdMmYyyy(now: Date): string {
  const pad = (value: number): string => String(value).padStart(2, '0');
  return `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()}`;
}

function parseDdMmYyyyToDate(value: any): Date | null {
  if (value === null || value === undefined || value === '') return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === 'number' && !isNaN(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d, 0, 0, 0);
  }

  const raw = String(value).trim();
  const match = raw.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]);
    const year = Number(match[3].length === 2 ? `20${match[3]}` : match[3]);
    const parsed = new Date(year, month - 1, day, 0, 0, 0);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  const parsed = new Date(raw);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function toDateOnlyIso(value: Date | null): string | null {
  if (!value || isNaN(value.getTime())) return null;
  return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0)).toISOString();
}

function parseAreaImpactada(value: any): string[] {
  if (Array.isArray(value)) {
    return value.map((item) => String(item || '').trim()).filter(Boolean);
  }

  return String(value || '')
    .split('/')
    .map((part) => part.trim())
    .filter(Boolean);
}

function normalizeLookupIds(value: any): number[] {
  if (Array.isArray(value)) {
    return value.map((item) => Number(item)).filter((item) => Number.isFinite(item) && item > 0);
  }

  if (value && Array.isArray(value.results)) {
    return value.results.map((item: any) => Number(item)).filter((item: number) => Number.isFinite(item) && item > 0);
  }

  const single = Number(value);
  return Number.isFinite(single) && single > 0 ? [single] : [];
}

function getRelativeFolderWithinProcesos(fileUrl: string): string {
  const full = trimSlash(fileUrl);
  const root = trimSlash(PROCESOS_ROOT);
  const fileDir = full.substring(0, full.lastIndexOf('/'));
  return fileDir.indexOf(root) === 0 ? fileDir.substring(root.length).replace(/^\/+/, '') : '';
}

function buildRenamedHistoricalFileName(originalName: string, oldVersion: string, todayStamp: string): string {
  const baseName = String(originalName || '').replace(/\.[^.]+$/, '');
  const extension = (String(originalName || '').match(/\.[^.]+$/) || [''])[0];
  return `${baseName}_V${sanitizeVersion(oldVersion)}_${todayStamp}${extension}`;
}

function replaceExtension(name: string, extensionWithDot: string): string {
  const baseName = String(name || '').replace(/\.[^.]+$/, '');
  return `${baseName}${extensionWithDot}`;
}

function base64UrlEncode(str: string): string {
  const b64 = btoa(unescape(encodeURIComponent(str)));
  return b64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function buildGraphShareIdFromUrl(absoluteUrl: string): string {
  return `u!${base64UrlEncode(encodeURI(absoluteUrl))}`;
}

function obtenerPrimerNombreYApellido(displayName: string): string {
  const limpio = (displayName || '').replace(/\s+/g, ' ').trim();
  if (!limpio) return '';
  const partes = limpio.split(' ').filter(Boolean);
  if (partes.length === 4) return `${partes[0]} ${partes[2]}`;
  if (partes.length === 3) return `${partes[0]} ${partes[1]}`;
  if (partes.length >= 2) return `${partes[0]} ${partes[1]}`;
  return partes[0] || '';
}

function pickBestWordAttachment(files: Array<{ FileName: string; ServerRelativeUrl: string; }>): { FileName: string; ServerRelativeUrl: string; } | null {
  const candidates = (files || []).filter((file) => /\.docx$/i.test(String(file?.FileName || '')));
  return candidates.length ? candidates[0] : null;
}

function buildMoveCopyBody(webUrl: string, srcFileUrl: string, destFileUrl: string, overwrite: boolean): any {
  const origin = new URL(webUrl).origin;
  const toAbsolute = (value: string): string => `${origin}${value.startsWith('/') ? '' : '/'}${value}`;

  return {
    srcPath: { DecodedUrl: toAbsolute(srcFileUrl) },
    destPath: { DecodedUrl: toAbsolute(destFileUrl) },
    overwrite,
    options: {
      KeepBoth: false,
      ResetAuthorAndCreatedOnCopy: false,
      ShouldBypassSharedLocks: true
    }
  };
}

async function moveFileByPath(
  context: WebPartContext,
  webUrl: string,
  srcFileUrl: string,
  destFileUrl: string,
  overwrite: boolean
): Promise<void> {
  await spPostJson(
    context,
    webUrl,
    `${webUrl}/_api/SP.MoveCopyUtil.MoveFileByPath()`,
    buildMoveCopyBody(webUrl, srcFileUrl, destFileUrl, overwrite),
    'POST'
  );
}

async function copyFileByPath(
  context: WebPartContext,
  webUrl: string,
  srcFileUrl: string,
  destFileUrl: string,
  overwrite: boolean
): Promise<void> {
  await spPostJson(
    context,
    webUrl,
    `${webUrl}/_api/SP.MoveCopyUtil.CopyFileByPath()`,
    buildMoveCopyBody(webUrl, srcFileUrl, destFileUrl, overwrite),
    'POST'
  );
}

async function updateFileMetadataByPath(context: WebPartContext, webUrl: string, fileUrl: string, payload: any): Promise<void> {
  await spPostJson(
    context,
    webUrl,
    `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapeODataValue(fileUrl)}')/ListItemAllFields`,
    payload,
    'MERGE'
  );
}

async function getFieldInfoByListPath(
  context: WebPartContext,
  webUrl: string,
  listPath: string,
  fieldInternalName: string
): Promise<any> {
  const url =
    `${webUrl}/_api/web/GetList('${escapeODataValue(listPath)}')/fields/getbyinternalnameortitle('${escapeODataValue(fieldInternalName)}')` +
    `?$select=Choices,AllowMultipleValues,TypeAsString`;

  return spGetJson<any>(context, url);
}

async function ensureChoiceOptionByListPath(
  context: WebPartContext,
  webUrl: string,
  listPath: string,
  fieldInternalName: string,
  valueToEnsure: string
): Promise<void> {
  const normalized = String(valueToEnsure || '').trim();
  if (!normalized) return;

  const field = await getFieldInfoByListPath(context, webUrl, listPath, fieldInternalName);
  const choices = Array.isArray(field?.Choices) ? field.Choices : [];
  const exists = choices.some((choice: string) => String(choice || '').trim().toLowerCase() === normalized.toLowerCase());
  if (exists) return;

  await spPostJson(
    context,
    webUrl,
    `${webUrl}/_api/web/GetList('${escapeODataValue(listPath)}')/fields/getbyinternalnameortitle('${escapeODataValue(fieldInternalName)}')`,
    { Choices: [...choices, normalized] },
    'MERGE'
  );
}

async function getAllowMultipleValues(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldInternalName: string
): Promise<boolean> {
  const field = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${listTitle.replace(/'/g, "''")}')/fields/getbyinternalnameortitle('${fieldInternalName.replace(/'/g, "''")}')?$select=AllowMultipleValues,TypeAsString`
  );
  return !!field?.AllowMultipleValues || String(field?.TypeAsString || '').toLowerCase().indexOf('multi') !== -1;
}

async function getAllowMultipleValuesByListPath(
  context: WebPartContext,
  webUrl: string,
  listPath: string,
  fieldInternalName: string
): Promise<boolean> {
  const field = await getFieldInfoByListPath(context, webUrl, listPath, fieldInternalName);
  return !!field?.AllowMultipleValues || String(field?.TypeAsString || '').toLowerCase().indexOf('multi') !== -1;
}

async function getFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldTitleOrInternalName: string
): Promise<string> {
  const field = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${listTitle.replace(/'/g, "''")}')/fields/getbyinternalnameortitle('${fieldTitleOrInternalName.replace(/'/g, "''")}')?$select=InternalName`
  );
  return String(field?.InternalName || fieldTitleOrInternalName);
}

async function resolveFirstExistingFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  candidates: string[]
): Promise<string> {
  for (let i = 0; i < candidates.length; i++) {
    try {
      const resolved = await getFieldInternalName(context, webUrl, listTitle, candidates[i]);
      if (resolved) return resolved;
    } catch (_error) {
      // continuar
    }
  }
  throw new Error(`No se encontró el campo esperado en "${listTitle}": ${candidates.join(', ')}`);
}

async function getFileItemMetadata(context: WebPartContext, webUrl: string, fileUrl: string): Promise<any> {
  const url =
    `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapeODataValue(fileUrl)}')/ListItemAllFields` +
    `?$select=Id,Title,FileLeafRef,FileRef,NombreDocumento,Tipodedocumento,CategoriaDocumento,Codigodedocumento,AreaDuena,AreaImpactada,` +
    `SolicitudId,Clasificaciondeproceso,Macroproceso,Proceso,Subproceso,Resumen,FechaDeAprobacion,FechaDeVigencia,` +
    `InstanciaDeAprobacionId,VersionDocumento,Accion,Aprobadores,Descripcion,DocumentoPadreId,FechaDePublicacion`;

  return spGetJson<any>(context, url);
}

async function getCurrentProcessFileBySolicitudId(
  context: WebPartContext,
  webUrl: string,
  solicitudId: number
): Promise<{ FileRef: string; FileLeafRef: string; Id: number; } | null> {
  const items = await spGetJson<{ value?: any[] }>(
    context,
    `${webUrl}/_api/web/GetList('${escapeODataValue(PROCESOS_ROOT)}')/items?$select=Id,FileRef,FileLeafRef,SolicitudId&$filter=SolicitudId eq ${solicitudId}&$top=5`
  );
  const row = (items.value || [])[0];
  if (!row) return null;

  return {
    Id: Number(row.Id || 0),
    FileRef: String(row.FileRef || ''),
    FileLeafRef: String(row.FileLeafRef || '')
  };
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

async function buscarSolicitudVigentePorNombre(context: WebPartContext, webUrl: string, documentName: string): Promise<any | null> {
  const escaped = String(documentName || '').replace(/'/g, "''").trim();
  if (!escaped) return null;

  const filter = `(Title eq '${escaped}' or NombreDocumento eq '${escaped}') and EsVersionActualDocumento eq 1`;
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items?$select=Id,Title,NombreDocumento,CodigoDocumento,EsVersionActualDocumento&$top=5&$filter=${encodeURIComponent(filter)}`
  );

  if (items.length > 1) {
    throw new Error(`Se encontraron múltiples solicitudes vigentes para "${documentName}".`);
  }

  return items.length ? items[0] : null;
}

async function getAreaGerenteNombre(context: WebPartContext, webUrl: string, areaId: number): Promise<string> {
  if (!areaId) return '';
  const item = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Áreas de Negocio')/items(${areaId})?$select=Id,Gerente/Title&$expand=Gerente`
  );
  return obtenerPrimerNombreYApellido(item?.Gerente?.Title || '');
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
        solicitudId: id,
        codigo,
        nombre,
        enlace: codigo
          ? `${new URL(webUrl).origin}/sites/SistemadeGestionDocumental/SitePages/verdocumento-vigente.aspx?codigodocumentosgd=${encodeURIComponent(codigo)}`
          : ''
      });
    } catch (_error) {
      // omitir hijo inválido
    }
  }
  return rows;
}

async function getChildSolicitudIdsByParent(
  context: WebPartContext,
  webUrl: string,
  parentSolicitudId: number
): Promise<number[]> {
  const parentField = await getFieldInternalName(context, webUrl, 'Relaciones Documentos', 'DocumentoPadre');
  const childField = await getFieldInternalName(context, webUrl, 'Relaciones Documentos', 'DocumentoHijo');
  const parentFieldId = `${parentField}Id`;
  const childFieldId = `${childField}Id`;
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Relaciones Documentos')/items?$select=Id,${parentFieldId},${childFieldId}&$top=5000&$filter=${parentFieldId} eq ${parentSolicitudId}`
  );

  return items.map((item) => Number(item[childFieldId] || 0)).filter((id) => id > 0);
}

async function getDiagramasFlujoRowsBySolicitud(
  context: WebPartContext,
  webUrl: string,
  solicitudId: number
): Promise<IDiagramaRow[]> {
  const listRoot = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Diagramas de Flujo')?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`
  );
  const rootUrl = listRoot?.RootFolder?.ServerRelativeUrl || '';
  const solicitudField = await getFieldInternalName(context, webUrl, 'Diagramas de Flujo', 'Solicitud');
  const solicitudFieldId = `${solicitudField}Id`;
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Diagramas de Flujo')/items?$select=Id,Title,Codigo,${solicitudFieldId}&$top=5000&$filter=${solicitudFieldId} eq ${solicitudId}`
  );

  const rows: IDiagramaRow[] = [];
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const attachments = await getAttachmentFiles(context, webUrl, 'Diagramas de Flujo', item.Id);
    const attachmentName = attachments.length ? attachments[0].FileName : '';
    rows.push({
      id: Number(item.Id || 0),
      codigo: item.Codigo || '',
      nombre: item.Title || '',
      enlace: attachmentName
        ? `${new URL(webUrl).origin}${rootUrl}/Attachments/${item.Id}/${attachmentName.split('/').map((p: string) => encodeURIComponent(p)).join('/')}`
        : ''
    });
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
    await updateListItem(context, webUrl, 'Relaciones Documentos', item.Id, { [parentFieldId]: newParentSolicitudId });
    updated++;
  }

  log(`🔗 Fase 8 | Relaciones actualizadas | PadreAnterior=${oldParentSolicitudId} | PadreNuevo=${newParentSolicitudId} | Registros=${updated}`);
  return updated;
}

async function updateExistingDiagramasSolicitud(
  context: WebPartContext,
  webUrl: string,
  oldParentSolicitudId: number,
  newParentSolicitudId: number,
  keepDiagramItemId: number,
  log: LogFn
): Promise<number> {
  const solicitudField = await getFieldInternalName(context, webUrl, 'Diagramas de Flujo', 'Solicitud');
  const solicitudFieldId = `${solicitudField}Id`;
  const solicitudIsMulti = await getAllowMultipleValues(context, webUrl, 'Diagramas de Flujo', solicitudField);
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Diagramas de Flujo')/items?$select=Id,Title,Codigo,${solicitudFieldId}&$top=5000&$filter=${solicitudFieldId} eq ${oldParentSolicitudId}`
  );

  let updated = 0;
  for (let i = 0; i < items.length; i++) {
    const itemId = Number(items[i].Id || 0);
    if (!itemId || itemId === keepDiagramItemId) continue;
    await updateListItem(context, webUrl, 'Diagramas de Flujo', itemId, {
      [solicitudFieldId]: solicitudIsMulti ? [newParentSolicitudId] : newParentSolicitudId
    });
    updated++;
  }

  log(`🧭 Fase 8 | Diagramas vigentes reasignados | SolicitudNueva=${newParentSolicitudId} | Registros=${updated}`);
  return updated;
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

  const docPadresField = await resolveFirstExistingFieldInternalName(context, webUrl, 'Solicitudes', ['docpadres', 'DocPadres', 'DocumentoPadre']);
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

    const currentIds = normalizeLookupIds(item[docPadresFieldId]);
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

  log(`👨‍👧 Fase 8 | DocPadres actualizados | PadreAnterior=${oldParentSolicitudId} | PadreNuevo=${newParentSolicitudId} | Solicitudes=${updated}`);
  return updated;
}

async function updateChildProcessParentReferences(params: {
  context: WebPartContext;
  webUrl: string;
  childSolicitudIds: number[];
  oldParentProcessItemId: number;
  newParentProcessItemId: number;
  log: LogFn;
}): Promise<number> {
  const childIds = Array.from(new Set((params.childSolicitudIds || []).filter((id) => Number(id) > 0)));
  if (!childIds.length || !params.oldParentProcessItemId || !params.newParentProcessItemId) return 0;

  const documentoPadreIsMulti = await getAllowMultipleValuesByListPath(params.context, params.webUrl, PROCESOS_ROOT, 'DocumentoPadre');
  let updated = 0;

  for (let i = 0; i < childIds.length; i++) {
    const childFile = await getCurrentProcessFileBySolicitudId(params.context, params.webUrl, childIds[i]);
    if (!childFile?.FileRef) continue;

    const childMetadata = await getFileItemMetadata(params.context, params.webUrl, childFile.FileRef);
    const currentIds = normalizeLookupIds(childMetadata?.DocumentoPadreId);
    if (!currentIds.length) continue;

    let changed = false;
    const nextIds = currentIds.map((id) => {
      if (id === params.oldParentProcessItemId) {
        changed = true;
        return params.newParentProcessItemId;
      }
      return id;
    });
    if (!changed) continue;

    const deduped = Array.from(new Set(nextIds.filter((id) => id > 0)));
    await updateFileMetadataByPath(params.context, params.webUrl, childFile.FileRef, {
      DocumentoPadreId: documentoPadreIsMulti ? deduped : (deduped[0] || null)
    });
    updated++;
  }

  params.log(`👨‍👧 Fase 8 | Referencias en Procesos actualizadas | PadreArchivoAnterior=${params.oldParentProcessItemId} | PadreArchivoNuevo=${params.newParentProcessItemId} | Archivos=${updated}`);
  return updated;
}

function buildNewSolicitudPayload(oldSolicitud: any, versionDocumento: string): any {
  const impactAreaIds = Array.isArray(oldSolicitud?.AreasImpactadas)
    ? oldSolicitud.AreasImpactadas.map((item: any) => Number(item?.Id || 0)).filter((id: number) => id > 0)
    : [];

  const payload: any = {
    Title: oldSolicitud?.NombreDocumento || oldSolicitud?.Title || '',
    Accion: 'Actualización de documento',
    NombreDocumento: oldSolicitud?.NombreDocumento || oldSolicitud?.Title || '',
    CategoriadeDocumento: oldSolicitud?.CategoriadeDocumento || '',
    ResumenDocumento: oldSolicitud?.ResumenDocumento || '',
    FechaDeAprobacionSolicitud: oldSolicitud?.FechaDeAprobacionSolicitud || new Date().toISOString(),
    FechadeVigencia: oldSolicitud?.FechadeVigencia || null,
    FechaDePublicacionSolicitud: new Date().toISOString(),
    FechadeEnvio: new Date().toISOString(),
    VersionDocumento: versionDocumento,
    EsVersionActualDocumento: true,
    DocumentosApoyo: !!oldSolicitud?.DocumentosApoyo,
    CodigoDocumento: oldSolicitud?.CodigoDocumento || ''
  };

  if (oldSolicitud?.TipoDocumentoId) payload.TipoDocumentoId = oldSolicitud.TipoDocumentoId;
  if (oldSolicitud?.ProcesoDeNegocioId) payload.ProcesoDeNegocioId = oldSolicitud.ProcesoDeNegocioId;
  if (oldSolicitud?.AreaDuenaId) payload.AreaDuenaId = oldSolicitud.AreaDuenaId;
  if (oldSolicitud?.EstadoId) payload.EstadoId = oldSolicitud.EstadoId;
  if (oldSolicitud?.InstanciasdeaprobacionId) payload.InstanciasdeaprobacionId = oldSolicitud.InstanciasdeaprobacionId;
  if (impactAreaIds.length) payload.AreasImpactadasId = impactAreaIds;

  return payload;
}

async function convertOfficeFileToPdfAndUpload(params: {
  context: WebPartContext;
  webUrl: string;
  sourceServerRelativeUrl: string;
  destinoFolderServerRelativeUrl: string;
  outputPdfName: string;
  log?: LogFn;
}): Promise<void> {
  const origin = new URL(params.webUrl).origin;
  const absoluteUrl = `${origin}${params.sourceServerRelativeUrl.startsWith('/') ? '' : '/'}${params.sourceServerRelativeUrl}`;
  const shareId = buildGraphShareIdFromUrl(absoluteUrl);
  const maxRetries = 3;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const client = await params.context.aadHttpClientFactory.getClient('https://graph.microsoft.com');
    const response = await client.get(
      `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem/content?format=pdf`,
      AadHttpClient.configurations.v1
    );

    if (response.ok) {
      const pdfBuffer = await response.arrayBuffer();
      const pdfBlob = new Blob([pdfBuffer], { type: 'application/pdf' });
      await uploadFileToFolder(params.context, params.webUrl, params.destinoFolderServerRelativeUrl, params.outputPdfName, pdfBlob);
      params.log?.(`📄✅ Fase 8 | Documento publicado en PDF | ${params.outputPdfName}`);
      return;
    }

    const body = await response.text();
    const normalized = String(body || '').toLowerCase();
    const retryable =
      [401, 403, 408, 409, 423, 429, 500, 502, 503, 504].indexOf(response.status) !== -1 ||
      normalized.indexOf('general_timeout') !== -1 ||
      normalized.indexOf('timeout') !== -1 ||
      normalized.indexOf('temporarily unavailable') !== -1;

    if (retryable && attempt < maxRetries) {
      params.log?.(`⚠️ Fase 8 | Reintentando conversión PDF (${attempt}/${maxRetries}) | ${params.outputPdfName} | HTTP ${response.status}`);
      await new Promise((resolve) => setTimeout(resolve, attempt * 2000));
      continue;
    }

    throw new Error(`Graph PDF failed (${response.status}): ${body}`);
  }
}

async function publishParentFile(params: {
  context: WebPartContext;
  webUrl: string;
  sourceFileUrl: string;
  targetFolderUrl: string;
  outputFileName: string;
  log?: LogFn;
}): Promise<string> {
  const destinationFileUrl = `${trimSlash(params.targetFolderUrl)}/${params.outputFileName}`;
  await convertOfficeFileToPdfAndUpload({
    context: params.context,
    webUrl: params.webUrl,
    sourceServerRelativeUrl: params.sourceFileUrl,
    destinoFolderServerRelativeUrl: params.targetFolderUrl,
    outputPdfName: params.outputFileName,
    log: params.log
  });
  return destinationFileUrl;
}

async function updateHistoricoMetadata(params: {
  context: WebPartContext;
  webUrl: string;
  historicoFileUrl: string;
  oldMetadata: any;
  fechaAprobacion: any;
  today: Date;
}): Promise<void> {
  const hisAreaImpactada = parseAreaImpactada(params.oldMetadata?.AreaImpactada).join(' / ');
  if (hisAreaImpactada) {
    await ensureChoiceOptionByListPath(params.context, params.webUrl, HISTORICOS_ROOT, 'HisAreaImpactada', hisAreaImpactada);
  }

  const payload: any = {
    HisAreaDuena: params.oldMetadata?.AreaDuena || '',
    HisAreaImpactada: hisAreaImpactada,
    HisClasificaciondeproceso: params.oldMetadata?.Clasificaciondeproceso || '',
    HisMacroproceso: params.oldMetadata?.Macroproceso || '',
    HisProceso: params.oldMetadata?.Proceso || '',
    HisSubproceso: params.oldMetadata?.Subproceso || '',
    HisTipodedocumento: params.oldMetadata?.Tipodedocumento || '',
    HisCodigodedocumento: params.oldMetadata?.Codigodedocumento || '',
    HisResumen: params.oldMetadata?.Resumen || '',
    HisVersionDocumento: params.oldMetadata?.VersionDocumento || '',
    HisAprobadores: params.oldMetadata?.Aprobadores || '',
    HisFechaDeBaja: toDateOnlyIso(params.today),
    HisCategoriaDocumento: params.oldMetadata?.CategoriaDocumento || '',
    InstanciaDeAprobacionId: Number(params.oldMetadata?.InstanciaDeAprobacionId || 0) || null,
    Accion: 'Actualización de documento',
    HisFechaAprobacionBaja: toDateOnlyIso(parseDdMmYyyyToDate(params.fechaAprobacion) || new Date(params.fechaAprobacion))
  };

  await updateFileMetadataByPath(params.context, params.webUrl, params.historicoFileUrl, payload);
}

async function stageDocumentoAHistorico(params: {
  context: WebPartContext;
  webUrl: string;
  solicitudId: number;
  fechaAprobacion: any;
  now: Date;
  todayStamp: string;
  log: LogFn;
}): Promise<{
  rutaProcesoOriginal: string;
  rutaProcesoRenombrada: string;
  rutaHistorico: string;
  nombreArchivo: string;
  versionDocumentoAnterior: string;
  codigoDocumento: string;
  nombreDocumento: string;
  processItemId: number;
  oldMetadata: any;
}> {
  const currentFile = await getCurrentProcessFileBySolicitudId(params.context, params.webUrl, params.solicitudId);
  if (!currentFile?.FileRef) {
    throw new Error(`No se encontró el archivo actual en Procesos para la solicitud ${params.solicitudId}.`);
  }

  const oldOriginalUrl = currentFile.FileRef;
  const oldMetadata = await getFileItemMetadata(params.context, params.webUrl, oldOriginalUrl);
  const relativeFolder = getRelativeFolderWithinProcesos(oldOriginalUrl);
  const procesosFolder = joinFolder(PROCESOS_ROOT, relativeFolder);
  const historicosFolder = joinFolder(HISTORICOS_ROOT, relativeFolder);
  const oldFileName = currentFile.FileLeafRef;
  const oldVersion = sanitizeVersion(oldMetadata?.VersionDocumento || '');
  const renamedFileName = buildRenamedHistoricalFileName(oldFileName, oldVersion, params.todayStamp);
  const oldRenamedUrl = `${procesosFolder}/${renamedFileName}`;
  const historicoUrl = `${historicosFolder}/${renamedFileName}`;

  params.log(`📁 Fase 8 | Documento original: ${oldOriginalUrl}`);
  params.log(`✏️ Fase 8 | Documento renombrado temporal: ${oldRenamedUrl}`);
  params.log(`📚 Fase 8 | Histórico destino: ${historicoUrl}`);

  await moveFileByPath(params.context, params.webUrl, oldOriginalUrl, oldRenamedUrl, false);
  await ensureFolderPath(params.context, params.webUrl, historicosFolder);
  await copyFileByPath(params.context, params.webUrl, oldRenamedUrl, historicoUrl, false);
  await updateHistoricoMetadata({
    context: params.context,
    webUrl: params.webUrl,
    historicoFileUrl: historicoUrl,
    oldMetadata,
    fechaAprobacion: params.fechaAprobacion,
    today: params.now
  });

  return {
    rutaProcesoOriginal: oldOriginalUrl,
    rutaProcesoRenombrada: oldRenamedUrl,
    rutaHistorico: historicoUrl,
    nombreArchivo: oldFileName,
    versionDocumentoAnterior: oldMetadata?.VersionDocumento || '',
    codigoDocumento: oldMetadata?.Codigodedocumento || '',
    nombreDocumento: oldMetadata?.NombreDocumento || oldMetadata?.Title || '',
    processItemId: currentFile.Id,
    oldMetadata
  };
}

async function updateParentProcesosMetadataAfterPublish(params: {
  context: WebPartContext;
  webUrl: string;
  targetFileUrl: string;
  oldMetadata: any;
  parentSolicitud: any;
  newSolicitudId: number;
  newVersion: string;
}): Promise<void> {
  const fechaAprobacion = params.parentSolicitud?.FechaDeAprobacionSolicitud || params.oldMetadata?.FechaDeAprobacion || null;
  const fechaVigencia = params.parentSolicitud?.FechadeVigencia || params.oldMetadata?.FechaDeVigencia || null;
  const areaImpactada = parseAreaImpactada(params.oldMetadata?.AreaImpactada);
  const areaImpactadaIsMulti = await getAllowMultipleValuesByListPath(params.context, params.webUrl, PROCESOS_ROOT, 'AreaImpactada');

  const payload: any = {
    Clasificaciondeproceso: params.oldMetadata?.Clasificaciondeproceso || '',
    AreaDuena: params.oldMetadata?.AreaDuena || '',
    VersionDocumento: params.newVersion || '',
    AreaImpactada: areaImpactadaIsMulti ? areaImpactada : (areaImpactada[0] || ''),
    Macroproceso: params.oldMetadata?.Macroproceso || '',
    Proceso: params.oldMetadata?.Proceso || '',
    Subproceso: params.oldMetadata?.Subproceso || '',
    Tipodedocumento: params.oldMetadata?.Tipodedocumento || '',
    SolicitudId: Number(params.newSolicitudId || 0) || null,
    Codigodedocumento: params.oldMetadata?.Codigodedocumento || params.parentSolicitud?.CodigoDocumento || '',
    Resumen: params.parentSolicitud?.ResumenDocumento || params.oldMetadata?.Resumen || '',
    CategoriaDocumento: params.oldMetadata?.CategoriaDocumento || '',
    FechaDeAprobacion: fechaAprobacion,
    FechaDePublicacion: new Date().toISOString(),
    FechaDeVigencia: fechaVigencia,
    InstanciaDeAprobacionId: Number(params.parentSolicitud?.InstanciasdeaprobacionId || params.oldMetadata?.InstanciaDeAprobacionId || 0) || null,
    Accion: 'Actualización de documento',
    NombreDocumento: params.parentSolicitud?.NombreDocumento || params.parentSolicitud?.Title || params.oldMetadata?.NombreDocumento || ''
  };

  await updateFileMetadataByPath(params.context, params.webUrl, params.targetFileUrl, payload);
}

async function readArrayBufferFromFilePicker(file: IFilePickerResult): Promise<ArrayBuffer> {
  if (!file) throw new Error('Archivo Excel no recibido.');

  if (typeof file.downloadFileContent === 'function') {
    const blob = await file.downloadFileContent();
    return blob.arrayBuffer();
  }

  const url = (file as any).fileAbsoluteUrl || '';
  if (!url) throw new Error('No se pudo obtener el contenido del Excel de Fase 8.');

  const response = await fetch(url, { credentials: 'same-origin' });
  if (!response.ok) {
    throw new Error(`No se pudo descargar el Excel de Fase 8. HTTP ${response.status}`);
  }

  return response.arrayBuffer();
}

async function readFase8Excel(file: IFilePickerResult): Promise<IFase8ExcelRow[]> {
  const buffer = await readArrayBufferFromFilePicker(file);
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error('No se encontró la hoja del Excel de Fase 8.');

  const worksheet = workbook.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false }) as any[][];
  if (!aoa.length) return [];

  const headers = aoa[0] || [];
  const headerMap = new Map<string, number>();
  for (let i = 0; i < headers.length; i++) {
    headerMap.set(normalizeHeader(headers[i]), i);
  }

  const findIndex = (aliases: string[]): number => {
    for (let i = 0; i < aliases.length; i++) {
      const index = headerMap.get(normalizeHeader(aliases[i]));
      if (index !== undefined) return index;
    }
    return -1;
  };

  const idxArchivo = findIndex(['Nombre del Archivo', 'NombreArchivo', 'Archivo']);
  const idxDocumento = findIndex(['Nombre del Documento', 'NombreDocumento', 'Documento']);
  const idxPadre = findIndex(['Documento Padre', 'DocumentoPadre', 'Padre']);

  if (idxArchivo === -1 || idxDocumento === -1 || idxPadre === -1) {
    throw new Error('El Excel de Fase 8 debe contener las columnas "Nombre del Archivo", "Nombre del Documento" y "Documento Padre".');
  }

  const rows: IFase8ExcelRow[] = [];
  for (let i = 1; i < aoa.length; i++) {
    const row = aoa[i] || [];
    const nombreArchivo = String(row[idxArchivo] || '').trim();
    const nombreDocumento = String(row[idxDocumento] || '').trim();
    const documentoPadre = String(row[idxPadre] || '').trim();

    if (!nombreArchivo && !nombreDocumento && !documentoPadre) continue;
    rows.push({ nombreArchivo, nombreDocumento, documentoPadre });
  }

  return rows;
}

async function buildSourceFileMaps(params: {
  context: WebPartContext;
  webUrl: string;
  folderServerRelativeUrl: string;
  log: LogFn;
}): Promise<{
  exactMap: Map<string, { name: string; url: string; }>;
  looseMap: Map<string, { name: string; url: string; }>;
}> {
  const files = await listFilesRecursive(params.context, params.webUrl, params.folderServerRelativeUrl, params.log);
  const exactMap = new Map<string, { name: string; url: string; }>();
  const looseMap = new Map<string, { name: string; url: string; }>();

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const exactKey = String(file.Name || '').trim().toLowerCase();
    const looseKey = normalizeLooseFileKey(file.Name || '');
    if (exactKey && !exactMap.has(exactKey)) exactMap.set(exactKey, { name: file.Name, url: file.ServerRelativeUrl });
    if (looseKey && !looseMap.has(looseKey)) looseMap.set(looseKey, { name: file.Name, url: file.ServerRelativeUrl });
  }

  return { exactMap, looseMap };
}

function resolveSourceFile(
  fileName: string,
  exactMap: Map<string, { name: string; url: string; }>,
  looseMap: Map<string, { name: string; url: string; }>
): { name: string; url: string; } | null {
  const exact = exactMap.get(String(fileName || '').trim().toLowerCase());
  if (exact) return exact;
  return looseMap.get(normalizeLooseFileKey(fileName || '')) || null;
}

async function createDiagramItemWithAttachment(params: {
  context: WebPartContext;
  webUrl: string;
  solicitudId: number;
  nombreDocumento: string;
  sourceFileName: string;
  sourceFileUrl: string;
  log: LogFn;
}): Promise<number> {
  const solicitudField = await getFieldInternalName(params.context, params.webUrl, 'Diagramas de Flujo', 'Solicitud');
  const solicitudFieldId = `${solicitudField}Id`;
  const solicitudIsMulti = await getAllowMultipleValues(params.context, params.webUrl, 'Diagramas de Flujo', solicitudField);
  const payload: any = {
    Title: params.nombreDocumento || params.sourceFileName || '',
    Codigo: ''
  };
  payload[solicitudFieldId] = solicitudIsMulti ? [params.solicitudId] : params.solicitudId;

  const itemId = await addListItem(params.context, params.webUrl, 'Diagramas de Flujo', payload);
  const response = await fetch(`${new URL(params.webUrl).origin}${params.sourceFileUrl}`, { credentials: 'same-origin' });
  if (!response.ok) {
    throw new Error(`No se pudo descargar el archivo BPM "${params.sourceFileName}". HTTP ${response.status}`);
  }

  await addAttachment(params.context, params.webUrl, 'Diagramas de Flujo', itemId, params.sourceFileName, await response.blob());
  params.log(`🧩 Fase 8 | Nuevo diagrama creado | ID=${itemId} | Solicitud=${params.solicitudId} | Archivo=${params.sourceFileName}`);
  return itemId;
}

export async function ejecutarFase8AltaDiagramaSolicitud(params: {
  context: WebPartContext;
  excelFile: IFilePickerResult;
  sourceFolderServerRelativeUrl: string;
  log?: LogFn;
}): Promise<{
  reportRows: IFase2PublicacionReportRow[];
  rollbackEntries: IFase8RollbackEntry[];
  reportFileName: string;
  processed: number;
  ok: number;
  skipped: number;
  error: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const rows = await readFase8Excel(params.excelFile);
  const fileMaps = await buildSourceFileMaps({
    context: params.context,
    webUrl,
    folderServerRelativeUrl: params.sourceFolderServerRelativeUrl,
    log
  });

  const reportRows: IFase2PublicacionReportRow[] = [];
  const rollbackEntries: IFase8RollbackEntry[] = [];
  const now = new Date();
  const todayStamp = buildTodayStamp(now);
  const todayText = buildTodayDdMmYyyy(now);

  let ok = 0;
  let skipped = 0;
  let error = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowRollbackEntries: IFase8RollbackEntry[] = [];
    let newSolicitudId: number | '' = '';
    let publishedUrl = '';

    if (!row.documentoPadre || !row.nombreArchivo || !row.nombreDocumento) {
      skipped++;
      reportRows.push({
        EstadoFase2: 'SKIP',
        SolicitudOrigenID: '',
        SolicitudID: '',
        NombreDocumento: row.documentoPadre || '',
        NombreArchivo: row.nombreArchivo || '',
        CodigoDocumento: '',
        ArchivoProcesoOriginal: '',
        RutaProcesoOriginal: '',
        ArchivoProcesoRenombrado: '',
        RutaProcesoRenombrada: '',
        RutaHistorico: '',
        RutaNuevoPublicado: '',
        VersionDocumentoAnterior: '',
        VersionDocumentoNueva: '',
        FechaBajaHistorico: todayText,
        FechaAprobacionBaja: '',
        Error: 'Fila incompleta. Se requieren Nombre del Archivo, Nombre del Documento y Documento Padre.'
      });
      continue;
    }

    try {
      const parentSolicitudRef = await buscarSolicitudVigentePorNombre(params.context, webUrl, row.documentoPadre);
      if (!parentSolicitudRef?.Id) {
        throw new Error(`No se encontró una solicitud vigente para el Documento Padre "${row.documentoPadre}".`);
      }

      const parentSolicitudId = Number(parentSolicitudRef.Id || 0);
      const parentSolicitud = await getSolicitudById(params.context, webUrl, parentSolicitudId);
      const currentParentFile = await getCurrentProcessFileBySolicitudId(params.context, webUrl, parentSolicitudId);
      if (!currentParentFile?.FileRef) {
        throw new Error(`No se encontró el archivo vigente del padre ${parentSolicitudId}.`);
      }

      const parentAttachments = await getAttachmentFiles(params.context, webUrl, 'Solicitudes', parentSolicitudId);
      const sourceWord = pickBestWordAttachment(parentAttachments as any);
      if (!sourceWord?.ServerRelativeUrl) {
        throw new Error(`No se encontró un Word adjunto en la solicitud ${parentSolicitudId} para regenerarlo.`);
      }

      const sourceDiagramFile = resolveSourceFile(row.nombreArchivo, fileMaps.exactMap, fileMaps.looseMap);
      if (!sourceDiagramFile?.url) {
        throw new Error(`No se encontró el archivo origen "${row.nombreArchivo}" dentro de la carpeta seleccionada.`);
      }

      const childIds = await getChildSolicitudIdsByParent(params.context, webUrl, parentSolicitudId);
      const relacionados = await getSolicitudesRelacionadas(params.context, webUrl, childIds);
      const existingDiagramas = await getDiagramasFlujoRowsBySolicitud(params.context, webUrl, parentSolicitudId);
      const versionNueva = incrementVersion(parentSolicitud?.VersionDocumento || '1.0');
      const duenoDocumento = await getAreaGerenteNombre(params.context, webUrl, Number(parentSolicitud?.AreaDuenaId || 0));
      const impactAreaIds = Array.isArray(parentSolicitud?.AreasImpactadas)
        ? parentSolicitud.AreasImpactadas.map((item: any) => Number(item?.Id || 0)).filter((id: number) => id > 0)
        : [];

      newSolicitudId = await addListItem(
        params.context,
        webUrl,
        'Solicitudes',
        buildNewSolicitudPayload(parentSolicitud, versionNueva)
      );

      const rollbackEntry: IFase8RollbackEntry = {
        solicitudId: parentSolicitudId,
        nombreDocumento: parentSolicitud?.NombreDocumento || parentSolicitud?.Title || row.documentoPadre || '',
        oldOriginalUrl: currentParentFile.FileRef,
        oldRenamedUrl: '',
        historicoUrl: '',
        oldOriginalMetadata: undefined,
        replacementSolicitudId: Number(newSolicitudId),
        updatedChildSolicitudIds: childIds.slice(),
        reassignedExistingDiagramIds: existingDiagramas.map((item) => item.id).filter((id) => id > 0)
      };
      rowRollbackEntries.push(rollbackEntry);

      const newDiagramItemId = await createDiagramItemWithAttachment({
        context: params.context,
        webUrl,
        solicitudId: Number(newSolicitudId),
        nombreDocumento: row.nombreDocumento,
        sourceFileName: row.nombreArchivo,
        sourceFileUrl: sourceDiagramFile.url,
        log
      });

      rollbackEntry.newDiagramItemId = newDiagramItemId;

      const newDiagramas = await getDiagramasFlujoRowsBySolicitud(params.context, webUrl, Number(newSolicitudId));
      const diagramas = [...existingDiagramas, ...newDiagramas.filter((item) => item.id === newDiagramItemId)];

      const tempDestino = joinFolder(TEMP_WORD_ROOT, getRelativeFolderWithinProcesos(currentParentFile.FileRef));
      log(`📂 Fase 8 | TEMP destino padre ${parentSolicitudId}: ${tempDestino}`);

      const attachResult = await fillAndAttachFromServerRelativeUrl({
        context: params.context,
        webUrl,
        listTitle: 'Solicitudes',
        itemId: Number(newSolicitudId),
        originalFileName: sourceWord.FileName,
        sourceFileServerRelativeUrl: sourceWord.ServerRelativeUrl,
        titulo: parentSolicitud?.NombreDocumento || parentSolicitud?.Title || '',
        instanciaRaw: parentSolicitud?.Instanciasdeaprobacion?.Title || 'Gerencia de Área',
        impactAreaIds,
        dueno: duenoDocumento,
        fechaVigencia: parentSolicitud?.FechadeVigencia || '',
        fechaAprobacion: parentSolicitud?.FechaDeAprobacionSolicitud || '',
        resumen: parentSolicitud?.ResumenDocumento || '',
        version: versionNueva,
        codigoDocumento: parentSolicitud?.CodigoDocumento || '',
        relacionados: relacionados.map((item) => ({ codigo: item.codigo, nombre: item.nombre, enlace: item.enlace })),
        diagramasFlujo: diagramas.map((item) => ({ codigo: item.codigo, nombre: item.nombre, enlace: item.enlace })),
        tempDestinoFolderServerRelativeUrl: tempDestino,
        replaceIfExists: true,
        log
      });
      if (!attachResult.ok || !attachResult.tempFileServerRelativeUrl) {
        throw new Error(attachResult.error || `No se pudo regenerar el adjunto del padre ${parentSolicitudId}.`);
      }
      rollbackEntry.tempFileUrl = attachResult.tempFileServerRelativeUrl;

      const movedParent = await stageDocumentoAHistorico({
        context: params.context,
        webUrl,
        solicitudId: parentSolicitudId,
        fechaAprobacion: parentSolicitud?.FechaDeAprobacionSolicitud || '',
        now,
        todayStamp,
        log
      });

      rollbackEntry.oldOriginalUrl = movedParent.rutaProcesoOriginal;
      rollbackEntry.oldRenamedUrl = movedParent.rutaProcesoRenombrada;
      rollbackEntry.historicoUrl = movedParent.rutaHistorico;
      rollbackEntry.oldOriginalMetadata = movedParent.oldMetadata;

      const procesosFolder = joinFolder(PROCESOS_ROOT, getRelativeFolderWithinProcesos(movedParent.rutaProcesoOriginal));
      const outputFileName = replaceExtension(currentParentFile.FileLeafRef, '.pdf');
      publishedUrl = await publishParentFile({
        context: params.context,
        webUrl,
        sourceFileUrl: attachResult.tempFileServerRelativeUrl,
        targetFolderUrl: procesosFolder,
        outputFileName,
        log
      });

      rollbackEntry.replacementPublishedUrl = publishedUrl;

      await updateParentProcesosMetadataAfterPublish({
        context: params.context,
        webUrl,
        targetFileUrl: publishedUrl,
        oldMetadata: movedParent.oldMetadata,
        parentSolicitud,
        newSolicitudId: Number(newSolicitudId),
        newVersion: versionNueva
      });

      await updateExistingDiagramasSolicitud(params.context, webUrl, parentSolicitudId, Number(newSolicitudId), newDiagramItemId, log);
      await updateRelacionesDocumentosPadre(params.context, webUrl, parentSolicitudId, Number(newSolicitudId), childIds, log);
      await updateChildSolicitudesDocPadres(params.context, webUrl, childIds, parentSolicitudId, Number(newSolicitudId), log);
      await updateListItem(params.context, webUrl, 'Solicitudes', parentSolicitudId, { EsVersionActualDocumento: false });
      await recycleFile(params.context, webUrl, movedParent.rutaProcesoRenombrada);

      const newParentProcess = await getCurrentProcessFileBySolicitudId(params.context, webUrl, Number(newSolicitudId));
      rollbackEntry.replacementProcessItemId = Number(newParentProcess?.Id || 0) || undefined;
      if (movedParent.processItemId && Number(newParentProcess?.Id || 0)) {
        await updateChildProcessParentReferences({
          context: params.context,
          webUrl,
          childSolicitudIds: childIds,
          oldParentProcessItemId: movedParent.processItemId,
          newParentProcessItemId: Number(newParentProcess?.Id || 0),
          log
        });
      }

      rollbackEntries.push(...rowRollbackEntries);
      ok++;
      reportRows.push({
        EstadoFase2: 'OK',
        SolicitudOrigenID: parentSolicitudId,
        SolicitudID: Number(newSolicitudId),
        NombreDocumento: parentSolicitud?.NombreDocumento || parentSolicitud?.Title || row.documentoPadre,
        NombreArchivo: movedParent.nombreArchivo || '',
        CodigoDocumento: movedParent.codigoDocumento || '',
        ArchivoProcesoOriginal: movedParent.nombreArchivo || '',
        RutaProcesoOriginal: movedParent.rutaProcesoOriginal || '',
        ArchivoProcesoRenombrado: movedParent.rutaProcesoRenombrada.split('/').pop() || '',
        RutaProcesoRenombrada: movedParent.rutaProcesoRenombrada || '',
        RutaHistorico: movedParent.rutaHistorico || '',
        RutaNuevoPublicado: publishedUrl || '',
        VersionDocumentoAnterior: movedParent.versionDocumentoAnterior || '',
        VersionDocumentoNueva: versionNueva,
        FechaBajaHistorico: todayText,
        FechaAprobacionBaja: String(parentSolicitud?.FechaDeAprobacionSolicitud || ''),
        Error: ''
      });
      log(`✅ Fase 8 | Solicitud versionada ${parentSolicitudId} -> ${newSolicitudId} | Diagrama nuevo=${newDiagramItemId}`);
    } catch (fase8Error) {
      const message = fase8Error instanceof Error ? fase8Error.message : String(fase8Error);

      if (rowRollbackEntries.length) {
        try {
          await rollbackModificacionFase8({
            context: params.context,
            webUrl,
            entries: rowRollbackEntries,
            log
          });
          log(`↩️ Fase 8 | Rollback local ejecutado para el documento padre ${row.documentoPadre}`);
        } catch (rollbackError) {
          const rollbackMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
          log(`⚠️ Fase 8 | Falló el rollback local del documento padre ${row.documentoPadre}: ${rollbackMessage}`);
        }
      }

      error++;
      reportRows.push({
        EstadoFase2: 'ERROR',
        SolicitudOrigenID: '',
        SolicitudID: newSolicitudId || '',
        NombreDocumento: row.documentoPadre || '',
        NombreArchivo: row.nombreArchivo || '',
        CodigoDocumento: '',
        ArchivoProcesoOriginal: '',
        RutaProcesoOriginal: '',
        ArchivoProcesoRenombrado: '',
        RutaProcesoRenombrada: '',
        RutaHistorico: '',
        RutaNuevoPublicado: publishedUrl || '',
        VersionDocumentoAnterior: '',
        VersionDocumentoNueva: '',
        FechaBajaHistorico: todayText,
        FechaAprobacionBaja: '',
        Error: message
      });
      log(`❌ Error Fase 8 | DocumentoPadre="${row.documentoPadre}" | Diagrama="${row.nombreDocumento}" | ${message}`);
    }
  }

  const reportFileName = `Reporte_Fase8_ALTA_DIAGRAMA_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}${String(now.getSeconds()).padStart(2, '0')}.xlsx`;
  descargarReporteFase2Publicacion(reportRows, reportFileName);

  return {
    reportRows,
    rollbackEntries,
    reportFileName,
    processed: rows.length,
    ok,
    skipped,
    error
  };
}


export async function ejecutarCorreccionSolicitud2460(params: {
  context: WebPartContext;
  log?: LogFn;
}): Promise<{
  rollbackEntries: IFase8RollbackEntry[];
  processed: number;
  ok: number;
  error: number;
  solicitudOrigenId: number;
  solicitudNuevaId: number;
  rutaNuevoPublicado: string;
}> {
  const SOLICITUD_ID_OBJETIVO = 2460;
  const NOMBRE_SOLICITUD_OBJETIVO = 'Procedimiento de Registro Manual de Asientos Contables';
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const now = new Date();
  const todayStamp = buildTodayStamp(now);
  const rollbackEntries: IFase8RollbackEntry[] = [];
  let newSolicitudId = 0;
  let publishedUrl = '';

  try {
    const parentSolicitud = await getSolicitudById(params.context, webUrl, SOLICITUD_ID_OBJETIVO);
    if (!parentSolicitud?.Id) {
      throw new Error(`No se encontró la solicitud ${SOLICITUD_ID_OBJETIVO}.`);
    }

    const nombreActual = String(parentSolicitud?.NombreDocumento || parentSolicitud?.Title || '').trim();
    if (nombreActual !== NOMBRE_SOLICITUD_OBJETIVO) {
      throw new Error(`La solicitud ${SOLICITUD_ID_OBJETIVO} no corresponde a "${NOMBRE_SOLICITUD_OBJETIVO}". Se encontró "${nombreActual}".`);
    }

    if (!parentSolicitud?.EsVersionActualDocumento) {
      throw new Error(`La solicitud ${SOLICITUD_ID_OBJETIVO} ya no está vigente.`);
    }

    const currentParentFile = await getCurrentProcessFileBySolicitudId(params.context, webUrl, SOLICITUD_ID_OBJETIVO);
    if (!currentParentFile?.FileRef) {
      throw new Error(`No se encontró el archivo vigente del padre ${SOLICITUD_ID_OBJETIVO}.`);
    }

    const parentAttachments = await getAttachmentFiles(params.context, webUrl, 'Solicitudes', SOLICITUD_ID_OBJETIVO);
    const sourceWord = pickBestWordAttachment(parentAttachments as any);
    if (!sourceWord?.ServerRelativeUrl) {
      throw new Error(`No se encontró un Word adjunto en la solicitud ${SOLICITUD_ID_OBJETIVO} para regenerarlo.`);
    }

    const childIds = await getChildSolicitudIdsByParent(params.context, webUrl, SOLICITUD_ID_OBJETIVO);
    const relacionados = await getSolicitudesRelacionadas(params.context, webUrl, childIds);
    const existingDiagramas = await getDiagramasFlujoRowsBySolicitud(params.context, webUrl, SOLICITUD_ID_OBJETIVO);
    const versionNueva = incrementVersion(parentSolicitud?.VersionDocumento || '1.0');
    const duenoDocumento = await getAreaGerenteNombre(params.context, webUrl, Number(parentSolicitud?.AreaDuenaId || 0));
    const impactAreaIds = Array.isArray(parentSolicitud?.AreasImpactadas)
      ? parentSolicitud.AreasImpactadas.map((item: any) => Number(item?.Id || 0)).filter((id: number) => id > 0)
      : [];

    log(`🛠️ Corrección 2460 | Regenerando solicitud vigente ${SOLICITUD_ID_OBJETIVO} | "${NOMBRE_SOLICITUD_OBJETIVO}"`);
    log(`🔎 Corrección 2460 | Relacionados encontrados=${relacionados.length} | Diagramas encontrados=${existingDiagramas.length}`);

    newSolicitudId = await addListItem(
      params.context,
      webUrl,
      'Solicitudes',
      buildNewSolicitudPayload(parentSolicitud, versionNueva)
    );

    const rollbackEntry: IFase8RollbackEntry = {
      solicitudId: SOLICITUD_ID_OBJETIVO,
      nombreDocumento: nombreActual,
      oldOriginalUrl: currentParentFile.FileRef,
      oldRenamedUrl: '',
      historicoUrl: '',
      oldOriginalMetadata: undefined,
      replacementSolicitudId: Number(newSolicitudId),
      updatedChildSolicitudIds: childIds.slice(),
      reassignedExistingDiagramIds: existingDiagramas.map((item) => item.id).filter((id) => id > 0)
    };
    rollbackEntries.push(rollbackEntry);

    const tempDestino = joinFolder(TEMP_WORD_ROOT, getRelativeFolderWithinProcesos(currentParentFile.FileRef));
    log(`📂 Corrección 2460 | TEMP destino padre ${SOLICITUD_ID_OBJETIVO}: ${tempDestino}`);

    const attachResult = await fillAndAttachFromServerRelativeUrl({
      context: params.context,
      webUrl,
      listTitle: 'Solicitudes',
      itemId: Number(newSolicitudId),
      originalFileName: sourceWord.FileName,
      sourceFileServerRelativeUrl: sourceWord.ServerRelativeUrl,
      titulo: nombreActual,
      instanciaRaw: parentSolicitud?.Instanciasdeaprobacion?.Title || 'Gerencia de Área',
      impactAreaIds,
      dueno: duenoDocumento,
      fechaVigencia: parentSolicitud?.FechadeVigencia || '',
      fechaAprobacion: parentSolicitud?.FechaDeAprobacionSolicitud || '',
      resumen: parentSolicitud?.ResumenDocumento || '',
      version: versionNueva,
      codigoDocumento: parentSolicitud?.CodigoDocumento || '',
      relacionados: relacionados.map((item) => ({ codigo: item.codigo, nombre: item.nombre, enlace: item.enlace })),
      diagramasFlujo: existingDiagramas.map((item) => ({ codigo: item.codigo, nombre: item.nombre, enlace: item.enlace })),
      tempDestinoFolderServerRelativeUrl: tempDestino,
      replaceIfExists: true,
      log
    });
    if (!attachResult.ok || !attachResult.tempFileServerRelativeUrl) {
      throw new Error(attachResult.error || `No se pudo regenerar el adjunto del padre ${SOLICITUD_ID_OBJETIVO}.`);
    }
    rollbackEntry.tempFileUrl = attachResult.tempFileServerRelativeUrl;

    const movedParent = await stageDocumentoAHistorico({
      context: params.context,
      webUrl,
      solicitudId: SOLICITUD_ID_OBJETIVO,
      fechaAprobacion: parentSolicitud?.FechaDeAprobacionSolicitud || '',
      now,
      todayStamp,
      log
    });

    rollbackEntry.oldOriginalUrl = movedParent.rutaProcesoOriginal;
    rollbackEntry.oldRenamedUrl = movedParent.rutaProcesoRenombrada;
    rollbackEntry.historicoUrl = movedParent.rutaHistorico;
    rollbackEntry.oldOriginalMetadata = movedParent.oldMetadata;

    const procesosFolder = joinFolder(PROCESOS_ROOT, getRelativeFolderWithinProcesos(movedParent.rutaProcesoOriginal));
    const outputFileName = replaceExtension(currentParentFile.FileLeafRef, '.pdf');
    publishedUrl = await publishParentFile({
      context: params.context,
      webUrl,
      sourceFileUrl: attachResult.tempFileServerRelativeUrl,
      targetFolderUrl: procesosFolder,
      outputFileName,
      log
    });

    rollbackEntry.replacementPublishedUrl = publishedUrl;

    await updateParentProcesosMetadataAfterPublish({
      context: params.context,
      webUrl,
      targetFileUrl: publishedUrl,
      oldMetadata: movedParent.oldMetadata,
      parentSolicitud,
      newSolicitudId: Number(newSolicitudId),
      newVersion: versionNueva
    });

    await updateExistingDiagramasSolicitud(params.context, webUrl, SOLICITUD_ID_OBJETIVO, Number(newSolicitudId), 0, log);
    await updateRelacionesDocumentosPadre(params.context, webUrl, SOLICITUD_ID_OBJETIVO, Number(newSolicitudId), childIds, log);
    await updateChildSolicitudesDocPadres(params.context, webUrl, childIds, SOLICITUD_ID_OBJETIVO, Number(newSolicitudId), log);
    await updateListItem(params.context, webUrl, 'Solicitudes', SOLICITUD_ID_OBJETIVO, { EsVersionActualDocumento: false });
    await recycleFile(params.context, webUrl, movedParent.rutaProcesoRenombrada);

    const newParentProcess = await getCurrentProcessFileBySolicitudId(params.context, webUrl, Number(newSolicitudId));
    rollbackEntry.replacementProcessItemId = Number(newParentProcess?.Id || 0) || undefined;
    if (movedParent.processItemId && Number(newParentProcess?.Id || 0)) {
      await updateChildProcessParentReferences({
        context: params.context,
        webUrl,
        childSolicitudIds: childIds,
        oldParentProcessItemId: movedParent.processItemId,
        newParentProcessItemId: Number(newParentProcess?.Id || 0),
        log
      });
    }

    log(`✅ Corrección 2460 | Solicitud versionada ${SOLICITUD_ID_OBJETIVO} -> ${newSolicitudId}`);

    return {
      rollbackEntries,
      processed: 1,
      ok: 1,
      error: 0,
      solicitudOrigenId: SOLICITUD_ID_OBJETIVO,
      solicitudNuevaId: Number(newSolicitudId),
      rutaNuevoPublicado: publishedUrl || ''
    };
  } catch (correccionError) {
    const message = correccionError instanceof Error ? correccionError.message : String(correccionError);

    if (rollbackEntries.length) {
      try {
        await rollbackModificacionFase8({
          context: params.context,
          webUrl,
          entries: rollbackEntries,
          log
        });
        log(`↩️ Corrección 2460 | Rollback local ejecutado para la solicitud ${SOLICITUD_ID_OBJETIVO}`);
      } catch (rollbackError) {
        const rollbackMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
        log(`⚠️ Corrección 2460 | Falló el rollback local de la solicitud ${SOLICITUD_ID_OBJETIVO}: ${rollbackMessage}`);
      }
    }

    log(`❌ Error Corrección 2460 | Solicitud=${SOLICITUD_ID_OBJETIVO} | ${message}`);
    throw correccionError;
  }
}
