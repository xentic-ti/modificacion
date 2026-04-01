/* eslint-disable */
// @ts-nocheck
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { fillAndAttachFromFolder } from './documentFillAndAttach.service';
import { listFilesRecursive } from './spFolderExplorer.service';
import { addListItem, ensureFolderPath, escapeODataValue, getAllItems, recycleFile, spGetJson, spPostJson, updateListItem } from './sharepointRest.service';
import { descargarReporteFase1Word, IFase1WordReportRow } from '../utils/fase1WordReportExcel';
import { openExcelRevisionSession } from '../utils/modificacionExcelHelper';

type LogFn = (s: string) => void;
const PROCESOS_ROOT = '/sites/SistemadeGestionDocumental/Procesos';
const HISTORICOS_ROOT = '/sites/SistemadeGestionDocumental/Documentos Histricos';

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

function trimSlash(value: string): string {
  return String(value || '').replace(/\/+$/, '');
}

function joinFolder(base: string, relative: string): string {
  const cleanBase = trimSlash(base);
  const cleanRelative = String(relative || '').replace(/^\/+/, '').replace(/\/+$/, '');
  return cleanRelative ? `${cleanBase}/${cleanRelative}` : cleanBase;
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
    if (String(headers[i] || '').trim().toLowerCase() === normalized) return i;
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

  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists(guid'${field.LookupList}')/items?$select=Id,Title&$top=5000`
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

async function getAllowMultipleValuesByListPath(
  context: WebPartContext,
  webUrl: string,
  listPath: string,
  fieldInternalName: string
): Promise<boolean> {
  const field = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/GetList('${escapeODataValue(listPath)}')/fields/getbyinternalnameortitle('${escapeODataValue(fieldInternalName)}')?$select=AllowMultipleValues,TypeAsString`
  );
  return !!field.AllowMultipleValues || String(field?.TypeAsString || '').toLowerCase().indexOf('multi') !== -1;
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
      [cleanPart(item.Title), cleanPart(item.field_1), cleanPart(item.field_2), cleanPart(item.field_3)].filter(Boolean).join('/')
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
  if (isNaN(parsed.getTime()) || parsed.getUTCFullYear() !== year || parsed.getUTCMonth() !== month - 1 || parsed.getUTCDate() !== day) {
    throw new Error('Fecha inválida.');
  }
  return parsed.toISOString();
}

function parseLooseSlashDateToIso(raw: string): string | null {
  const match = raw.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (!match) return null;
  return parseDatePartsToIso(
    Number(match[3].length === 2 ? `20${match[3]}` : match[3]),
    Number(match[2]),
    Number(match[1])
  );
}

function excelSerialToIso(serialValue: number): string {
  const wholeDays = Number(String(serialValue).split('.')[0] || serialValue);
  const parsed = new Date(Date.UTC(1899, 11, 30 + wholeDays, 0, 0, 0));
  if (isNaN(parsed.getTime())) throw new Error('Fecha serial de Excel inválida.');
  return parsed.toISOString();
}

function buildTodayStamp(now: Date): string {
  const pad = (value: number): string => String(value).padStart(2, '0');
  return `${pad(now.getDate())}${pad(now.getMonth() + 1)}${now.getFullYear()}`;
}

function sanitizeVersion(value: any): string {
  return String(value || '').trim().replace(/^v/i, '') || '1.0';
}

function replaceExtension(name: string, extensionWithDot: string): string {
  const baseName = String(name || '').replace(/\.[^.]+$/, '');
  return `${baseName}${extensionWithDot}`;
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

function toDateOnlyIso(value: Date | null): string | null {
  if (!value || isNaN(value.getTime())) return null;
  return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0)).toISOString();
}

function parseDdMmYyyyToDate(value: any): Date | null {
  if (value === null || value === undefined || value === '') return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === 'number' && !isNaN(value)) {
    return new Date(excelSerialToIso(value));
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
  const impactNamesRaw = (excelRow.areasImpactadasTxt || '')
    .split(/[\n;,|/]+/g)
    .map((item) => item.trim())
    .filter(Boolean);

  const impactSet = new Set<string>();
  for (let i = 0; i < impactNamesRaw.length; i++) addImpactNameUnique(impactSet, impactNamesRaw[i]);
  if (isSi(excelRow.flagConducta)) addImpactNameUnique(impactSet, 'Impacto en Conducta de Mercado');
  if (isSi(excelRow.flagExperiencia)) addImpactNameUnique(impactSet, 'Impacto en Experiencia del Cliente');
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

async function buscarSolicitudPorNombre(context: WebPartContext, webUrl: string, documentName: string): Promise<any | null> {
  const filter = `(Title eq '${String(documentName || '').replace(/'/g, `''`)}' or NombreDocumento eq '${String(documentName || '').replace(/'/g, `''`)}')`;
  const url =
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items` +
    `?$select=Id,Title,NombreDocumento,CodigoDocumento` +
    `&$top=2&$filter=${encodeURIComponent(filter)}`;
  const items = await getAllItems<any>(context, url);
  return items.length ? items[0] : null;
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

async function getParentProcessFileItemIdBySolicitudId(
  context: WebPartContext,
  webUrl: string,
  solicitudId: number
): Promise<number | null> {
  const row = await getCurrentProcessFileBySolicitudId(context, webUrl, solicitudId);
  return row ? row.Id : null;
}

async function getFileItemMetadata(context: WebPartContext, webUrl: string, fileUrl: string): Promise<any> {
  const url =
    `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapeODataValue(fileUrl)}')/ListItemAllFields` +
    `?$select=Id,Title,FileLeafRef,FileRef,NombreDocumento,Tipodedocumento,CategoriaDocumento,Codigodedocumento,AreaDuena,AreaImpactada,` +
    `SolicitudId,Clasificaciondeproceso,Macroproceso,Proceso,Subproceso,Resumen,FechaDeAprobacion,FechaDeVigencia,` +
    `InstanciaDeAprobacionId,VersionDocumento,Accion,Aprobadores,Descripcion,DocumentoPadreId,FechaDePublicacion`;
  return spGetJson<any>(context, url);
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

async function moveFileByPath(context: WebPartContext, webUrl: string, srcFileUrl: string, destFileUrl: string, overwrite: boolean): Promise<void> {
  await spPostJson(context, webUrl, `${webUrl}/_api/SP.MoveCopyUtil.MoveFileByPath()`, buildMoveCopyBody(webUrl, srcFileUrl, destFileUrl, overwrite), 'POST');
}

async function copyFileByPath(context: WebPartContext, webUrl: string, srcFileUrl: string, destFileUrl: string, overwrite: boolean): Promise<void> {
  await spPostJson(context, webUrl, `${webUrl}/_api/SP.MoveCopyUtil.CopyFileByPath()`, buildMoveCopyBody(webUrl, srcFileUrl, destFileUrl, overwrite), 'POST');
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

function base64UrlEncode(str: string): string {
  const b64 = btoa(unescape(encodeURIComponent(str)));
  return b64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function buildGraphShareIdFromUrl(absoluteUrl: string): string {
  const safe = encodeURI(absoluteUrl);
  return 'u!' + base64UrlEncode(safe);
}

async function fileExistsByServerRelativeUrl(context: WebPartContext, webUrl: string, fileUrl: string): Promise<boolean> {
  try {
    await spGetJson<any>(context, `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapeODataValue(fileUrl)}')?$select=Exists`);
    return true;
  } catch (error: any) {
    const message = String(error?.message || '');
    if (message.indexOf('(404)') !== -1 || message.toLowerCase().indexOf('not found') !== -1 || message.toLowerCase().indexOf('no se encuentra') !== -1) {
      return false;
    }
    throw error;
  }
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
  const client = await params.context.aadHttpClientFactory.getClient('https://graph.microsoft.com');
  const response = await client.get(`https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem/content?format=pdf`, 1 as any);
  if (!response.ok) {
    throw new Error(`Graph PDF failed (${response.status}): ${await response.text()}`);
  }
  const pdfBuffer = await response.arrayBuffer();
  const pdfBlob = new Blob([pdfBuffer], { type: 'application/pdf' });
  await spPostJson(params.context, params.webUrl, `${params.webUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${escapeODataValue(params.destinoFolderServerRelativeUrl)}')`, undefined, 'POST');
  await (await fetch(
    `${params.webUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${escapeODataValue(params.destinoFolderServerRelativeUrl)}')/Files/addUsingPath(decodedurl='${escapeODataValue(params.outputPdfName)}',overwrite=false)`,
    {
      method: 'POST',
      credentials: 'same-origin',
      headers: {
        Accept: 'application/json;odata=nometadata'
      },
      body: pdfBlob
    }
  ));
  params.log?.(`📄✅ Hijo publicado en PDF | ${params.outputPdfName}`);
}

async function publishNewFile(params: {
  context: WebPartContext;
  webUrl: string;
  sourceFileUrl: string;
  targetFolderUrl: string;
  outputFileName: string;
  log?: LogFn;
}): Promise<string> {
  const destinationFileUrl = `${trimSlash(params.targetFolderUrl)}/${params.outputFileName}`;
  if (/\.docx$/i.test(params.sourceFileUrl)) {
    const siblingPdfUrl = replaceExtension(params.sourceFileUrl, '.pdf');
    const siblingPdfExists = await fileExistsByServerRelativeUrl(params.context, params.webUrl, siblingPdfUrl);
    if (siblingPdfExists) {
      await copyFileByPath(params.context, params.webUrl, siblingPdfUrl, destinationFileUrl, false);
      params.log?.(`📄✅ Hijo publicado usando PDF existente | ${params.outputFileName}`);
    } else {
      throw new Error(`No se encontró el PDF temporal del hijo. Fuente="${siblingPdfUrl}"`);
    }
  } else {
    await copyFileByPath(params.context, params.webUrl, params.sourceFileUrl, destinationFileUrl, false);
    params.log?.(`📄✅ Hijo publicado sin conversión | ${params.outputFileName}`);
  }
  return destinationFileUrl;
}

async function updateHistoricoMetadata(params: {
  context: WebPartContext;
  webUrl: string;
  historicoFileUrl: string;
  oldMetadata: any;
  excelRow: IExcelRowData;
  today: Date;
}): Promise<void> {
  const payload: any = {
    HisAreaDuena: params.oldMetadata?.AreaDuena || '',
    HisAreaImpactada: parseAreaImpactada(params.oldMetadata?.AreaImpactada).join(' / '),
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
    HisFechaAprobacionBaja: toDateOnlyIso(parseDdMmYyyyToDate(params.excelRow.fechaAprobacion))
  };
  await updateFileMetadataByPath(params.context, params.webUrl, params.historicoFileUrl, payload);
}

async function updateChildProcesosMetadataAfterPublish(params: {
  context: WebPartContext;
  webUrl: string;
  targetFileUrl: string;
  excelRow: IExcelRowData;
  solicitudNuevaId: number;
  codigoDocumento: string;
  areaImpactadaIsMulti: boolean;
  parentProcessItemId: number | null;
  documentoPadreIsMulti: boolean;
  instanciaDeAprobacionId: number | '';
}): Promise<void> {
  const fechaAprobacion = toDateOnlyIso(parseDdMmYyyyToDate(params.excelRow.fechaAprobacion));
  const fechaVigencia = toDateOnlyIso(parseDdMmYyyyToDate(params.excelRow.fechaVigencia));
  const areaImpactada = buildImpactNames(params.excelRow);
  const payload: any = {
    Clasificaciondeproceso: params.excelRow.clasificacion || '',
    AreaDuena: params.excelRow.areaDuena || '',
    VersionDocumento: params.excelRow.versionDocumento || '',
    AreaImpactada: params.areaImpactadaIsMulti ? areaImpactada : (areaImpactada[0] || ''),
    Macroproceso: params.excelRow.macroproceso || '',
    Proceso: params.excelRow.proceso || '',
    Subproceso: params.excelRow.subproceso || '',
    Tipodedocumento: params.excelRow.tipoDocumento || '',
    SolicitudId: Number(params.solicitudNuevaId || 0) || null,
    Codigodedocumento: params.codigoDocumento || '',
    Resumen: params.excelRow.resumen || '',
    CategoriaDocumento: params.excelRow.categoriaDocumento || '',
    FechaDeAprobacion: fechaAprobacion,
    FechaDePublicacion: new Date().toISOString(),
    FechaDeVigencia: fechaVigencia,
    InstanciaDeAprobacionId: Number(params.instanciaDeAprobacionId || 0) || null,
    Accion: 'Actualización de documento',
    NombreDocumento: params.excelRow.nombreDocumento || ''
  };
  if (params.parentProcessItemId) {
    payload.DocumentoPadreId = params.documentoPadreIsMulti ? [params.parentProcessItemId] : params.parentProcessItemId;
  }
  await updateFileMetadataByPath(params.context, params.webUrl, params.targetFileUrl, payload);
}

function parseAreaImpactada(value: any): string[] {
  if (Array.isArray(value)) return value.map((item) => String(item || '').trim()).filter(Boolean);
  return String(value || '').split('/').map((part) => part.trim()).filter(Boolean);
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
    docPadresFieldId: string;
    docPadresIsMulti: boolean;
    parentSolicitudIds: number[];
  }
): any {
  if (!excelRow.nombreDocumento) throw new Error('El Excel no tiene NombreDocumento para crear la nueva solicitud.');
  if (!excelRow.categoriaDocumento) throw new Error(`El Excel no tiene CategoriaDocumento para "${excelRow.nombreDocumento}".`);
  if (!excelRow.resumen) throw new Error(`El Excel no tiene Resumen para "${excelRow.nombreDocumento}".`);
  if (!excelRow.fechaAprobacion) throw new Error(`El Excel no tiene FechaDeAprobacion para "${excelRow.nombreDocumento}".`);
  if (!excelRow.fechaVigencia) throw new Error(`El Excel no tiene FechaDeVigencia para "${excelRow.nombreDocumento}".`);
  if (!lookups.tipoDocumentoId) throw new Error(`No se encontró TipoDocumento en lookup para "${excelRow.nombreDocumento}". Valor Excel="${excelRow.tipoDocumento}".`);
  if (!lookups.procesoDeNegocioId) throw new Error(`No se encontró ProcesoDeNegocio en lookup para "${excelRow.nombreDocumento}".`);
  if (!lookups.areaDuenaId) throw new Error(`No se encontró AreaDuena en lookup para "${excelRow.nombreDocumento}".`);
  if (!lookups.parentSolicitudIds.length) throw new Error(`No se encontró la solicitud del padre "${excelRow.documentoPadre}".`);

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
    DocumentosApoyo: true,
    CodigoDocumento: oldSolicitud.CodigoDocumento || ''
  };

  if (lookups.tipoDocumentoId) payload.TipoDocumentoId = lookups.tipoDocumentoId;
  if (lookups.procesoDeNegocioId) payload.ProcesoDeNegocioId = lookups.procesoDeNegocioId;
  if (lookups.areaDuenaId) payload.AreaDuenaId = lookups.areaDuenaId;
  if (oldSolicitud.EstadoId) payload.EstadoId = oldSolicitud.EstadoId;
  if (lookups.instanciaAprobacionId) payload.InstanciasdeaprobacionId = lookups.instanciaAprobacionId;
  if (lookups.impactAreaIds.length) {
    payload.AreasImpactadasId = lookups.impactIsMulti ? lookups.impactAreaIds : lookups.impactAreaIds[0];
  }

  payload[lookups.docPadresFieldId] = lookups.docPadresIsMulti ? lookups.parentSolicitudIds : lookups.parentSolicitudIds[0];
  return payload;
}

export async function ejecutarFase5HijosConPadre(params: {
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
  const docPadresField = await resolveFirstExistingFieldInternalName(params.context, webUrl, 'Solicitudes', ['docpadres', 'DocPadres', 'DocumentoPadre']);
  const docPadresFieldId = `${docPadresField}Id`;
  const docPadresIsMulti = await getAllowMultipleValues(params.context, webUrl, 'Solicitudes', docPadresField);
  const procesosAreaImpactadaIsMulti = await getAllowMultipleValuesByListPath(params.context, webUrl, PROCESOS_ROOT, 'AreaImpactada');
  const procesosDocumentoPadreIsMulti = await getAllowMultipleValuesByListPath(params.context, webUrl, PROCESOS_ROOT, 'DocumentoPadre');

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

    if (!solicitudOrigenId || !nombreDocumento) continue;

    if (!tienePadre || tieneHijos || tieneFlujos) {
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
        Error: 'Documento omitido por no ser hijo con padre o por tener hijos/flujos.'
      });
      continue;
    }

    try {
      const oldSolicitud = await getSolicitudById(params.context, webUrl, solicitudOrigenId);
      const parentSolicitud = await buscarSolicitudPorNombre(params.context, webUrl, documentoPadre);
      if (!parentSolicitud?.Id) {
        throw new Error(`No se encontró la solicitud del padre por nombre "${documentoPadre}".`);
      }

      const versionNueva = incrementVersion(versionExcel || oldSolicitud.VersionDocumento || '1.0');
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

      if (oldSolicitudIds.indexOf(solicitudOrigenId) === -1) oldSolicitudIds.push(solicitudOrigenId);
      if (!impactIsMulti && impactAreaIds.length > 1) {
        throw new Error(`El campo AreasImpactadas no admite múltiples valores para "${excelRow.nombreDocumento}".`);
      }

      log(
        `🗓️ Fechas Fase 5 | SolicitudOrigen=${solicitudOrigenId} | ` +
        `FechaAprobacionRaw="${formatLogValue(excelRow.fechaAprobacionRaw)}" | ` +
        `FechaVigenciaRaw="${formatLogValue(excelRow.fechaVigenciaRaw)}"`
      );
      log(`👨‍👧 Fase 5 | Padre resuelto por nombre | Nombre="${documentoPadre}" | SolicitudPadre=${parentSolicitud.Id}`);

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
          impactIsMulti,
          docPadresFieldId,
          docPadresIsMulti,
          parentSolicitudIds: [Number(parentSolicitud.Id)]
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

      log(`📂 TEMP destino calculado Fase 5 | SolicitudOrigen=${solicitudOrigenId} | ${tempDestino}`);

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
        esDocumentoApoyo: true,
        tempDestinoFolderServerRelativeUrl: tempDestino,
        replaceIfExists: true,
        log
      });

      if (!attachResult.tempFileServerRelativeUrl) {
        throw new Error(`No se generó RutaTemporalWord para "${excelRow.nombreDocumento}".`);
      }
      tempFileUrls.push(attachResult.tempFileServerRelativeUrl);

      const currentFile = await getCurrentProcessFileBySolicitudId(params.context, webUrl, solicitudOrigenId);
      if (!currentFile?.FileRef) {
        throw new Error(`No se encontró el archivo actual en Procesos para la solicitud hija ${solicitudOrigenId}.`);
      }

      const oldOriginalUrl = currentFile.FileRef;
      const oldMetadata = await getFileItemMetadata(params.context, webUrl, oldOriginalUrl);
      const relativeFolder = getRelativeFolderWithinProcesos(oldOriginalUrl);
      const procesosFolder = joinFolder(PROCESOS_ROOT, relativeFolder);
      const historicosFolder = joinFolder(HISTORICOS_ROOT, relativeFolder);
      const oldFileName = currentFile.FileLeafRef;
      const oldVersion = sanitizeVersion(oldMetadata?.VersionDocumento || '');
      const renamedFileName = buildRenamedHistoricalFileName(oldFileName, oldVersion, buildTodayStamp(new Date()));
      const oldRenamedUrl = `${procesosFolder}/${renamedFileName}`;
      const historicoUrl = `${historicosFolder}/${renamedFileName}`;
      const outputFileName = /\.docx$/i.test(attachResult.tempFileServerRelativeUrl) ? replaceExtension(oldFileName, '.pdf') : oldFileName;
      const newPublishedUrl = `${procesosFolder}/${outputFileName}`;
      const parentProcessItemId = await getParentProcessFileItemIdBySolicitudId(params.context, webUrl, Number(parentSolicitud.Id || 0));

      log(`📁 Fase 5 | Hijo original: ${oldOriginalUrl}`);
      log(`✏️ Fase 5 | Hijo renombrado: ${oldRenamedUrl}`);
      log(`📚 Fase 5 | Histórico destino: ${historicoUrl}`);
      log(`🚀 Fase 5 | Nuevo destino: ${newPublishedUrl}`);
      log(`👨‍👧 Fase 5 | Padre en Procesos=${parentProcessItemId || ''} | SolicitudPadre=${parentSolicitud.Id}`);

      await moveFileByPath(params.context, webUrl, oldOriginalUrl, oldRenamedUrl, false);
      await ensureFolderPath(params.context, webUrl, historicosFolder);
      await copyFileByPath(params.context, webUrl, oldRenamedUrl, historicoUrl, false);
      await updateHistoricoMetadata({
        context: params.context,
        webUrl,
        historicoFileUrl: historicoUrl,
        oldMetadata,
        excelRow,
        today: new Date()
      });

      await publishNewFile({
        context: params.context,
        webUrl,
        sourceFileUrl: attachResult.tempFileServerRelativeUrl,
        targetFolderUrl: procesosFolder,
        outputFileName,
        log
      });

      await updateListItem(params.context, webUrl, 'Solicitudes', solicitudOrigenId, {
        EsVersionActualDocumento: false
      });

      await updateChildProcesosMetadataAfterPublish({
        context: params.context,
        webUrl,
        targetFileUrl: newPublishedUrl,
        excelRow: {
          ...excelRow,
          versionDocumento: versionNueva
        },
        solicitudNuevaId: newSolicitudId,
        codigoDocumento: oldSolicitud.CodigoDocumento || '',
        areaImpactadaIsMulti: procesosAreaImpactadaIsMulti,
        parentProcessItemId,
        documentoPadreIsMulti: procesosDocumentoPadreIsMulti,
        instanciaDeAprobacionId: instanciaAprobacionId || ''
      });

      await recycleFile(params.context, webUrl, oldRenamedUrl);

      reportRows.push({
        SolicitudOrigenID: solicitudOrigenId,
        SolicitudID: newSolicitudId,
        NombreDocumento: excelRow.nombreDocumento,
        NombreArchivo: nombreArchivo,
        CodigoDocumento: oldSolicitud.CodigoDocumento || '',
        VersionDocumento: versionNueva,
        TieneDocumentoPadre: 'Sí',
        DocumentoPadreNombre: documentoPadre,
        DocumentoPadreSolicitudID: Number(parentSolicitud.Id || 0) || '',
        CodigoDocumentoPadre: parentSolicitud.CodigoDocumento || '',
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
        TieneDocumentoPadre: 'Sí',
        DocumentoPadreNombre: documentoPadre,
        DocumentoPadreSolicitudID: '',
        PadreRegeneradoConLinks: 'No',
        RutaTemporalWord: '',
        EstadoFase1: 'ERROR',
        Error: message
      });
      log(
        `❌ Error Fase 5 | SolicitudOrigen=${solicitudOrigenId} | ` +
        `FechaAprobacionRaw="${formatLogValue(excelRow.fechaAprobacionRaw)}" | ` +
        `FechaVigenciaRaw="${formatLogValue(excelRow.fechaVigenciaRaw)}" | ${message}`
      );
    }
  }

  descargarReporteFase1Word(
    reportRows,
    `Reporte_Fase5_WORD_${new Date().getFullYear()}${String(new Date().getMonth() + 1).padStart(2, '0')}${String(new Date().getDate()).padStart(2, '0')}_${String(new Date().getHours()).padStart(2, '0')}${String(new Date().getMinutes()).padStart(2, '0')}${String(new Date().getSeconds()).padStart(2, '0')}.xlsx`
  );

  return {
    createdSolicitudIds,
    oldSolicitudIds,
    tempFileUrls,
    reportRows,
    processed: reportRows.length,
    ok: reportRows.filter((reportRow) => reportRow.EstadoFase1 === 'OK').length,
    skipped: reportRows.filter((reportRow) => reportRow.EstadoFase1.indexOf('SKIP') === 0).length,
    error: reportRows.filter((reportRow) => reportRow.EstadoFase1 === 'ERROR').length
  };
}
