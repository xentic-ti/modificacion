/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';
import { AadHttpClient } from '@microsoft/sp-http';
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { fillAndAttachFromServerRelativeUrl } from './documentFillAndAttachFromUrl.service';
import {
  addAttachment,
  deleteAttachment,
  deleteListItem,
  ensureAbsoluteUrl,
  escapeODataValue,
  getAllItems,
  getAttachmentFiles,
  recycleFile,
  spGetJson,
  spPostJson,
  updateListItem,
  uploadFileToFolder
} from './sharepointRest.service';
import { openExcelRevisionSession } from '../utils/modificacionExcelHelper';

type LogFn = (message: string) => void;
type IRelacionadoRow = { solicitudId: number; codigo: string; nombre: string; enlace: string; };
type IDiagramaRow = { id: number; codigo: string; nombre: string; enlace: string; };

const PROCESOS_ROOT = '/sites/SistemadeGestionDocumental/Procesos';
const HISTORICOS_ROOT = '/sites/SistemadeGestionDocumental/Documentos Histricos';
const TEMP_WORD_ROOT = '/sites/SistemadeGestionDocumental/Procesos/TEMP_MIGRACION_WORD';
const CODIGOS_DOCUMENTOS_LIST = 'Códigos Documentos';

type IInputRow = {
  rowIndex: number;
  row: any[];
  id: number;
  codigoDocumento: string;
  versionDocumento: string;
  title: string;
  nombreDocumento: string;
  created: string;
  procesar: boolean;
  procesarRaw: string;
  procesarMode: 'si' | 'renovar' | '';
};

type ISolicitudRuntimeInfo = {
  rowIndex: number;
  id: number;
  codigoDocumento: string;
  versionDocumento: string;
  created: string;
  title: string;
  nombreDocumento: string;
  procesar: boolean;
  categoriaDocumento: string;
  solicitud: any;
  childIds: number[];
  diagramas: IDiagramaRow[];
  related: IRelacionadoRow[];
  processFile: { Id: number; FileRef: string; FileLeafRef: string; } | null;
  attachment: { FileName: string; ServerRelativeUrl: string; } | null;
  hasDependencias: boolean;
};

type IExecutionTrace = {
  pasoFallido: string;
  codigoRegistroId: number | '';
  codigoReservado: string;
  adjuntoRegenerado: string;
  procesoReemplazado: string;
  solicitudActualizada: string;
  metadataProcesosActualizada: string;
  tempFileUrl: string;
  rollbackOmitido: string;
  rollbackCodigoRegistro: string;
  rollbackAdjunto: string;
  rollbackProceso: string;
  rollbackSolicitud: string;
  rollbackMetadataProcesos: string;
};

function normalizeHeader(value: any): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '')
    .trim()
    .toLowerCase();
}

function trimSlash(value: string): string {
  return String(value || '').replace(/\/+$/, '');
}

function joinFolder(base: string, relative: string): string {
  const cleanBase = trimSlash(base);
  const cleanRelative = String(relative || '').replace(/^\/+/, '').replace(/\/+$/, '');
  return cleanRelative ? `${cleanBase}/${cleanRelative}` : cleanBase;
}

function base64UrlEncode(str: string): string {
  const b64 = btoa(unescape(encodeURIComponent(str)));
  return b64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function buildGraphShareIdFromUrl(absoluteUrl: string): string {
  const safe = encodeURI(absoluteUrl);
  return 'u!' + base64UrlEncode(safe);
}

function splitFolderAndFile(fileUrl: string): { folder: string; fileName: string; } {
  const clean = String(fileUrl || '').trim();
  const lastSlash = clean.lastIndexOf('/');
  if (lastSlash === -1) {
    return { folder: '', fileName: clean };
  }

  return {
    folder: clean.substring(0, lastSlash),
    fileName: clean.substring(lastSlash + 1)
  };
}

function replaceExtension(name: string, extensionWithDot: string): string {
  const clean = String(name || '').trim();
  if (!clean) return clean;
  return clean.replace(/\.[^.]+$/, '') + extensionWithDot;
}

function getRelativeFolderWithinHistoricos(fileUrl: string): string {
  const full = trimSlash(fileUrl);
  const root = trimSlash(HISTORICOS_ROOT);
  const fileDir = full.substring(0, full.lastIndexOf('/'));
  return fileDir.indexOf(root) === 0 ? fileDir.substring(root.length).replace(/^\/+/, '') : '';
}

function removeHistoricalSuffix(fileName: string): string {
  const clean = String(fileName || '').trim();
  const extension = (clean.match(/\.[^.]+$/) || [''])[0];
  const baseName = extension ? clean.slice(0, -extension.length) : clean;
  const normalizedBase = baseName
    .replace(/_V[^_]+_\d{8}$/i, '')
    .replace(/\s+v\d+(?:\.\d+)?\s+\d{8}$/i, '')
    .replace(/\s+v\d+(?:\.\d+)?[_-]\d{8}$/i, '')
    .replace(/[_-]v\d+(?:\.\d+)?[_-]?\d{8}$/i, '')
    .trim();
  return `${normalizedBase}${extension}`;
}

function getRelativeFolderWithinProcesos(fileUrl: string): string {
  const full = trimSlash(fileUrl);
  const root = trimSlash(PROCESOS_ROOT);
  const fileDir = full.substring(0, full.lastIndexOf('/'));
  return fileDir.indexOf(root) === 0 ? fileDir.substring(root.length).replace(/^\/+/, '') : '';
}

function isVersionUno(value: any): boolean {
  return String(value || '').trim() === '1.0';
}

function isProcesarSi(value: any): boolean {
  return getProcesarMode(value) === 'si';
}

function getProcesarMode(value: any): 'si' | 'renovar' | '' {
  const normalized = String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
  if (normalized === 'si') return 'si';
  if (normalized === 'renovar') return 'renovar';
  return '';
}

function obtenerPrimerNombreYApellido(displayName: string): string {
  const limpio = (displayName || "").replace(/\s+/g, " ").trim();
  if (!limpio) return "";
  const partes = limpio.split(" ").filter(Boolean);

  if (partes.length === 3) return `${partes[0]} ${partes[1]}`;
  if (partes.length === 4) return `${partes[0]} ${partes[2]}`;
  if (partes.length >= 2) return `${partes[0]} ${partes[1]}`;
  return partes[0] || "";
}

function pickEditableAttachment(files: Array<{ FileName: string; ServerRelativeUrl: string; }>): { FileName: string; ServerRelativeUrl: string; } | null {
  const list = Array.isArray(files) ? files : [];
  const priorities = [/\.docx$/i, /\.xlsx$/i, /\.xlsm$/i, /\.xls$/i, /\.pptx$/i];

  for (let i = 0; i < priorities.length; i++) {
    const found = list.find((file) => priorities[i].test(String(file?.FileName || '')));
    if (found) {
      return found;
    }
  }

  return null;
}

function buildHeaderMap(headers: any[]): Map<string, number> {
  const map = new Map<string, number>();
  for (let i = 0; i < headers.length; i++) {
    map.set(normalizeHeader(headers[i]), i);
  }
  return map;
}

function getCellByCandidates(row: any[], headerMap: Map<string, number>, candidates: string[]): any {
  for (let i = 0; i < candidates.length; i++) {
    const index = headerMap.get(normalizeHeader(candidates[i]));
    if (index !== undefined) {
      return row[index];
    }
  }
  return '';
}

function cloneGrid(grid: any[][]): any[][] {
  const result: any[][] = [];
  for (let i = 0; i < grid.length; i++) {
    result.push(Array.isArray(grid[i]) ? [...grid[i]] : []);
  }
  return result;
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

async function getSolicitudById(context: WebPartContext, webUrl: string, solicitudId: number): Promise<any> {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items(${solicitudId})` +
    `?$select=Id,Title,NombreDocumento,CodigoDocumento,CategoriadeDocumento,ResumenDocumento,VersionDocumento,TipoDocumentoId,TipoDocumento/Title,ProcesoDeNegocioId,` +
    `ProcesoDeNegocio/Title,ProcesoDeNegocio/field_1,ProcesoDeNegocio/field_2,ProcesoDeNegocio/field_3,` +
    `FechaDeAprobacionSolicitud,FechadeVigencia,AreaDuenaId,InstanciasdeaprobacionId,Instanciasdeaprobacion/Title,` +
    `AreasImpactadas/Id,AreasImpactadas/Title,EsVersionActualDocumento` +
    `&$expand=TipoDocumento,ProcesoDeNegocio,Instanciasdeaprobacion,AreasImpactadas`;
  return spGetJson<any>(context, url);
}

async function getCurrentProcessFileBySolicitudId(
  context: WebPartContext,
  webUrl: string,
  solicitudId: number
): Promise<{ Id: number; FileRef: string; FileLeafRef: string; } | null> {
  const json = await spGetJson<{ value?: any[] }>(
    context,
    `${webUrl}/_api/web/GetList('${escapeODataValue(PROCESOS_ROOT)}')/items?$select=Id,FileRef,FileLeafRef,SolicitudId&$filter=SolicitudId eq ${solicitudId}&$top=5`
  );

  const item = (json.value || [])[0];
  if (!item) {
    return null;
  }

  return {
    Id: Number(item.Id || 0),
    FileRef: String(item.FileRef || ''),
    FileLeafRef: String(item.FileLeafRef || '')
  };
}

async function getAreaGerenteNombre(context: WebPartContext, webUrl: string, areaId: number): Promise<string> {
  if (!areaId) return '';
  const item = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Áreas de Negocio')/items(${areaId})?$select=Id,Gerente/Title&$expand=Gerente`
  );
  return obtenerPrimerNombreYApellido(String(item?.Gerente?.Title || '').trim());
}

async function getSolicitudesRelacionadas(
  context: WebPartContext,
  webUrl: string,
  ids: number[]
): Promise<IRelacionadoRow[]> {
  const uniqueIds = Array.from(new Set((ids || []).filter((id) => id > 0)));
  const rows: IRelacionadoRow[] = [];

  for (let i = 0; i < uniqueIds.length; i++) {
    try {
      const item = await spGetJson<any>(
        context,
        `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items(${uniqueIds[i]})?$select=Id,Title,NombreDocumento,CodigoDocumento`
      );
      const codigo = String(item?.CodigoDocumento || '').trim();
      rows.push({
        solicitudId: uniqueIds[i],
        codigo,
        nombre: String(item?.NombreDocumento || item?.Title || '').trim(),
        enlace: codigo
          ? `${new URL(webUrl).origin}/sites/SistemadeGestionDocumental/SitePages/verdocumento-vigente.aspx?codigodocumentosgd=${encodeURIComponent(codigo)}`
          : ''
      });
    } catch (_error) {
      // omitir relacionado inválido
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
      codigo: String(item?.Codigo || '').trim(),
      nombre: String(item?.Title || '').trim(),
      enlace: attachmentName
        ? `${new URL(webUrl).origin}${rootUrl}/Attachments/${item.Id}/${attachmentName.split('/').map((part: string) => encodeURIComponent(part)).join('/')}`
        : ''
    });
  }

  return rows;
}

async function getFileItemMetadata(context: WebPartContext, webUrl: string, fileUrl: string): Promise<any> {
  return spGetJson<any>(
    context,
    `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapeODataValue(fileUrl)}')/ListItemAllFields?$select=Id,Title,FileLeafRef,NombreDocumento,Tipodedocumento,CategoriaDocumento,Codigodedocumento,AreaDuena,AreaImpactada,SolicitudId,Clasificaciondeproceso,Macroproceso,Proceso,Subproceso,Resumen,FechaDeAprobacion,FechaDeVigencia,InstanciaDeAprobacionId,VersionDocumento,Accion,Aprobadores,Descripcion,DocumentoPadreId`
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

async function getHistoricoFileForRenew(
  context: WebPartContext,
  webUrl: string,
  solicitud: any,
  attachmentFileName?: string,
  log?: LogFn
): Promise<{ Id: number; FileRef: string; FileLeafRef: string; Title: string; } | null> {
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/GetList('${escapeODataValue(HISTORICOS_ROOT)}')/items?$select=Id,FileRef,FileLeafRef,Title&$top=5000`
  );

  const rows = (items || [])
    .map((item) => ({
      Id: Number(item?.Id || 0),
      FileRef: String(item?.FileRef || ''),
      FileLeafRef: String(item?.FileLeafRef || ''),
      Title: String(item?.Title || '').trim()
    }))
    .filter((item) => item.Id > 0 && item.FileRef);

  if (!rows.length) {
    return null;
  }

  const nombreDocumento = normalizeHeader(String(solicitud?.NombreDocumento || solicitud?.Title || '').trim());
  const attachmentName = normalizeHeader(String(attachmentFileName || '').trim());
  const rawNombreDocumento = String(solicitud?.NombreDocumento || solicitud?.Title || '').trim();
  const rawAttachmentName = String(attachmentFileName || '').trim();
  const historicalBaseName = removeHistoricalSuffix(rawAttachmentName);
  const historicalPdfName = /\.docx$/i.test(historicalBaseName)
    ? replaceExtension(historicalBaseName, '.pdf')
    : historicalBaseName;
  const expectedHistoricalNames = Array.from(
    new Set(
      [historicalBaseName, historicalPdfName]
        .map((value) => String(value || '').trim())
        .filter(Boolean)
    )
  );
  const expectedHistoricalNamesNormalized = new Set(
    expectedHistoricalNames.map((value) => normalizeHeader(value))
  );

  log?.(
    `🔎 Renovar | Buscando en históricos | Solicitud=${solicitud?.Id || ''} | ` +
    `Title="${rawNombreDocumento}" | Archivo="${rawAttachmentName}" | ` +
    `HistoricoEsperado="${expectedHistoricalNames.join(' | ')}"`
  );

  let candidatos = nombreDocumento
    ? rows.filter((item) => normalizeHeader(String(item.Title || '').trim()) === nombreDocumento)
    : [];

  if (candidatos.length) {
    log?.(`🔎 Renovar | Coincidencias por Title en históricos: ${candidatos.length}`);
  }

  if (!candidatos.length && expectedHistoricalNamesNormalized.size) {
    candidatos = rows.filter((item) => {
      const fileLeafBase = normalizeHeader(removeHistoricalSuffix(item.FileLeafRef).trim());
      return expectedHistoricalNamesNormalized.has(fileLeafBase);
    });
    if (candidatos.length) {
      log?.(`🔎 Renovar | Coincidencias por archivo base en históricos: ${candidatos.length}`);
    }
  }

  if (!candidatos.length) {
    const sampleFiles = rows
      .slice(0, 5)
      .map((item) => item.FileLeafRef)
      .filter(Boolean)
      .join(' | ');
    if (sampleFiles) {
      log?.(`🔎 Renovar | Muestra archivos en históricos: ${sampleFiles}`);
    }
    log?.(`⚠️ Renovar | No se encontraron coincidencias en históricos para Solicitud=${solicitud?.Id || ''}`);
    return null;
  }

  candidatos.sort((a, b) => a.Id - b.Id);
  return candidatos[candidatos.length - 1];
}

async function updateProcesosMetadataForRenew(params: {
  context: WebPartContext;
  webUrl: string;
  targetFileUrl: string;
  solicitud: any;
  baseMetadata: any;
}): Promise<void> {
  const areaImpactadaRaw = params.baseMetadata?.AreaImpactada || '';
  const areaImpactada = Array.isArray(areaImpactadaRaw)
    ? areaImpactadaRaw
    : String(areaImpactadaRaw || '').split(/[;/]/).map((item) => String(item || '').trim()).filter(Boolean);

  const payload: any = {
    Clasificaciondeproceso: params.baseMetadata?.Clasificaciondeproceso || params.solicitud?.ProcesoDeNegocio?.Title || '',
    AreaDuena: params.baseMetadata?.AreaDuena || '',
    VersionDocumento: params.solicitud?.VersionDocumento || params.baseMetadata?.VersionDocumento || '',
    AreaImpactada: areaImpactada,
    Macroproceso: params.baseMetadata?.Macroproceso || params.solicitud?.ProcesoDeNegocio?.Title || '',
    Proceso: params.baseMetadata?.Proceso || params.solicitud?.ProcesoDeNegocio?.field_1 || '',
    Subproceso: params.baseMetadata?.Subproceso || params.solicitud?.ProcesoDeNegocio?.field_2 || '',
    Tipodedocumento: params.baseMetadata?.Tipodedocumento || params.solicitud?.TipoDocumento?.Title || '',
    SolicitudId: Number(params.solicitud?.Id || 0) || null,
    Codigodedocumento: params.solicitud?.CodigoDocumento || params.baseMetadata?.Codigodedocumento || '',
    Resumen: params.solicitud?.ResumenDocumento || params.baseMetadata?.Resumen || '',
    CategoriaDocumento: params.baseMetadata?.CategoriaDocumento || params.solicitud?.CategoriadeDocumento || '',
    FechaDeAprobacion: params.solicitud?.FechaDeAprobacionSolicitud || params.baseMetadata?.FechaDeAprobacion || null,
    FechaDePublicacion: new Date().toISOString(),
    FechaDeVigencia: params.solicitud?.FechadeVigencia || params.baseMetadata?.FechaDeVigencia || null,
    InstanciaDeAprobacionId: Number(params.solicitud?.InstanciasdeaprobacionId || params.baseMetadata?.InstanciaDeAprobacionId || 0) || null,
    Accion: 'Actualización de documento',
    NombreDocumento: params.solicitud?.NombreDocumento || params.solicitud?.Title || params.baseMetadata?.NombreDocumento || ''
  };

  await updateFileMetadataByPath(params.context, params.webUrl, params.targetFileUrl, payload);
}

async function getLookupListIdFromField(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldInternalName: string
): Promise<string> {
  const field = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(listTitle)}')/fields/getbyinternalnameortitle('${escapeODataValue(fieldInternalName)}')?$select=LookupList`
  );
  return String(field?.LookupList || '').trim();
}

async function getCodigosTipoDocumentoByTipoDocumentoId(
  context: WebPartContext,
  webUrl: string,
  solicitudesListTitle: string,
  tipoDocumentoId: number
): Promise<{ codigoCategoria: 'DO' | 'DE' | 'DS'; codigoTipoDocumento: string; }> {
  if (!tipoDocumentoId) {
    throw new Error('La solicitud no tiene TipoDocumentoId.');
  }

  const lookupListId = await getLookupListIdFromField(context, webUrl, solicitudesListTitle, 'TipoDocumento');
  const item = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists(guid'${lookupListId}')/items(${tipoDocumentoId})?$select=CodigoCategoria,CodigoTipoDocumento`
  );

  return {
    codigoCategoria: String(item?.CodigoCategoria || '').trim() as 'DO' | 'DE' | 'DS',
    codigoTipoDocumento: String(item?.CodigoTipoDocumento || '').trim()
  };
}

async function getCodigoProcesoByProcesoCorporativoId(
  context: WebPartContext,
  webUrl: string,
  procesoCorporativoId: number
): Promise<string> {
  if (!procesoCorporativoId) {
    return '';
  }

  const item = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Procesos Corporativos')/items(${procesoCorporativoId})?$select=CodigoProceso`
  );

  return String(item?.CodigoProceso || '').trim();
}

async function getCodigosDocumentosByTitle(
  context: WebPartContext,
  webUrl: string,
  title: string
): Promise<any[]> {
  const filter = `Title eq '${String(title || '').replace(/'/g, `''`)}'`;
  return getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(CODIGOS_DOCUMENTOS_LIST)}')/items?$select=Id,Title,CodigoBase,CategoriaDocumento,CorrelativoPadre,TipoSoporte,CorrelativoHijo,DocumentoPadreId&$filter=${encodeURIComponent(filter)}&$top=50`
  );
}

async function obtenerSiguienteCorrelativoPadre(
  context: WebPartContext,
  webUrl: string,
  codigoBase: string,
  categoriaDocumento: string
): Promise<number> {
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(CODIGOS_DOCUMENTOS_LIST)}')/items?$select=Id,CorrelativoPadre,CodigoBase,CategoriaDocumento,DocumentoPadreId&$top=5000`
  );

  let maxCorrelativo = 0;
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (String(item?.CodigoBase || '').trim() !== codigoBase) {
      continue;
    }
    if (String(item?.CategoriaDocumento || '').trim() !== categoriaDocumento) {
      continue;
    }
    if (item?.DocumentoPadreId) {
      continue;
    }

    maxCorrelativo = Math.max(maxCorrelativo, Number(item?.CorrelativoPadre || 0));
  }

  return maxCorrelativo + 1;
}

async function obtenerSiguienteCorrelativoHijo(
  context: WebPartContext,
  webUrl: string,
  codigoPadreId: number,
  tipoSoporte: string
): Promise<number> {
  const filter = `DocumentoPadreId eq ${codigoPadreId} and TipoSoporte eq '${String(tipoSoporte || '').replace(/'/g, `''`)}'`;
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(CODIGOS_DOCUMENTOS_LIST)}')/items?$select=Id,CorrelativoHijo&$filter=${encodeURIComponent(filter)}&$orderby=CorrelativoHijo desc&$top=1`
  );

  return Number(items?.[0]?.CorrelativoHijo || 0) + 1;
}

async function getCodigoDocumentoRegistroById(
  context: WebPartContext,
  webUrl: string,
  itemId: number
): Promise<{ Id: number; Title: string; CodigoBase: string; CorrelativoPadre: number; }> {
  const item = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(CODIGOS_DOCUMENTOS_LIST)}')/items(${itemId})?$select=Id,Title,CodigoBase,CorrelativoPadre`
  );

  return {
    Id: Number(item?.Id || 0),
    Title: String(item?.Title || '').trim(),
    CodigoBase: String(item?.CodigoBase || '').trim(),
    CorrelativoPadre: Number(item?.CorrelativoPadre || 0)
  };
}

async function obtenerSiguienteCorrelativoDF(
  context: WebPartContext,
  webUrl: string,
  codigoDocumentoPadre: string
): Promise<number> {
  const filter = `CodigoBase eq '${String(codigoDocumentoPadre || '').replace(/'/g, `''`)}' and TipoSoporte eq 'DF'`;
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(CODIGOS_DOCUMENTOS_LIST)}')/items?$select=Id,CorrelativoHijo&$filter=${encodeURIComponent(filter)}&$orderby=CorrelativoHijo desc&$top=1`
  );

  return Number(items?.[0]?.CorrelativoHijo || 0) + 1;
}

async function registerCodigoDiagramaWithRetry(params: {
  context: WebPartContext;
  webUrl: string;
  parentCodigoRegistroId: number;
  parentCodigoCompleto: string;
  parentCodigoBase: string;
  correlativoPadre: number;
  log?: LogFn;
}): Promise<{ codigo: string; registroId: number; correlativoHijo: number; }> {
  const log = params.log || (() => undefined);
  const maxRetries = 8;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const correlativoHijo = await obtenerSiguienteCorrelativoDF(
      params.context,
      params.webUrl,
      params.parentCodigoCompleto
    );
    const codigo =
      `DS-${params.parentCodigoBase}-${String(params.correlativoPadre).padStart(3, '0')}-DF-${String(correlativoHijo).padStart(2, '0')}`;

    const existentesAntes = await getCodigosDocumentosByTitle(params.context, params.webUrl, codigo);
    if (existentesAntes.length) {
      log(`⚠️ Código de diagrama ya existente en Códigos Documentos, reintentando | ${codigo}`);
      continue;
    }

    const creado = await spPostJson<any>(
      params.context,
      params.webUrl,
      `${params.webUrl}/_api/web/lists/getbytitle('${escapeODataValue(CODIGOS_DOCUMENTOS_LIST)}')/items`,
      {
        Title: codigo,
        CodigoBase: params.parentCodigoCompleto,
        CategoriaDocumento: 'DS',
        CorrelativoPadre: params.correlativoPadre,
        TipoSoporte: 'DF',
        CorrelativoHijo: correlativoHijo,
        DocumentoPadreId: params.parentCodigoRegistroId
      },
      'POST'
    );
    const registroId = Number(creado?.Id || 0);
    if (!registroId) {
      throw new Error(`No se pudo registrar el código de diagrama "${codigo}" en Códigos Documentos.`);
    }

    const existentesDespues = await getCodigosDocumentosByTitle(params.context, params.webUrl, codigo);
    existentesDespues.sort((a, b) => Number(a?.Id || 0) - Number(b?.Id || 0));
    const ganadorId = Number(existentesDespues[0]?.Id || 0);

    if (ganadorId === registroId) {
      return {
        codigo,
        registroId,
        correlativoHijo
      };
    }

    await deleteListItem(params.context, params.webUrl, CODIGOS_DOCUMENTOS_LIST, registroId);
    log(`⚠️ Colisión concurrente al reservar código de diagrama, se liberó el registro ${registroId} y se reintenta.`);
  }

  throw new Error(`No se pudo reservar un código único de diagrama para el padre ${params.parentCodigoCompleto}.`);
}

async function resolveFirstExistingFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  candidates: string[]
): Promise<string | null> {
  for (let i = 0; i < candidates.length; i++) {
    try {
      const resolved = await getFieldInternalName(context, webUrl, listTitle, candidates[i]);
      if (resolved) {
        return resolved;
      }
    } catch (_error) {
      // continuar
    }
  }

  return null;
}

function normalizeLookupIds(value: any): number[] {
  if (Array.isArray(value)) {
    const result: number[] = [];
    for (let i = 0; i < value.length; i++) {
      const entry = value[i];
      const id = Number(entry?.Id || entry || 0);
      if (id > 0 && result.indexOf(id) === -1) {
        result.push(id);
      }
    }
    return result;
  }

  const single = Number(value);
  return Number.isFinite(single) && single > 0 ? [single] : [];
}

async function resolveParentSolicitudIdForDs(
  context: WebPartContext,
  webUrl: string,
  solicitudId: number,
  solicitud: any
): Promise<number> {
  const docPadresField = await resolveFirstExistingFieldInternalName(
    context,
    webUrl,
    'Solicitudes',
    ['docpadres', 'DocPadres', 'DocumentoPadre']
  );

  let localCandidates: number[] = [];
  if (docPadresField) {
    try {
      const fieldValue = await spGetJson<any>(
        context,
        `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items(${solicitudId})?$select=${escapeODataValue(docPadresField)}Id`
      );
      localCandidates = normalizeLookupIds(fieldValue?.[`${docPadresField}Id`]);
    } catch (_error) {
      localCandidates = [];
    }
  }

  if (localCandidates.length === 1) {
    return localCandidates[0];
  }
  if (localCandidates.length > 1) {
    throw new Error(`La solicitud ${solicitudId} tiene múltiples padres asociados.`);
  }

  const parentField = await getFieldInternalName(context, webUrl, 'Relaciones Documentos', 'DocumentoPadre');
  const childField = await getFieldInternalName(context, webUrl, 'Relaciones Documentos', 'DocumentoHijo');
  const parentFieldId = `${parentField}Id`;
  const childFieldId = `${childField}Id`;
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Relaciones Documentos')/items?$select=Id,${parentFieldId},${childFieldId}&$filter=${encodeURIComponent(`${childFieldId} eq ${solicitudId}`)}&$top=50`
  );

  const parentIds = Array.from(new Set(items.map((item) => Number(item?.[parentFieldId] || 0)).filter((id) => id > 0)));
  if (parentIds.length !== 1) {
    throw new Error(`No se pudo determinar un único padre para la solicitud DS ${solicitudId}.`);
  }

  return parentIds[0];
}

async function resolveCodigoDocumentoRegistroPadreId(
  context: WebPartContext,
  webUrl: string,
  parentSolicitudId: number
): Promise<number> {
  const parentSolicitud = await getSolicitudById(context, webUrl, parentSolicitudId);
  const parentCode = String(parentSolicitud?.CodigoDocumento || '').trim();
  if (!parentCode) {
    throw new Error(`La solicitud padre ${parentSolicitudId} no tiene CodigoDocumento.`);
  }

  const registros = await getCodigosDocumentosByTitle(context, webUrl, parentCode);
  const padres = registros.filter((item) => !item?.DocumentoPadreId);
  if (!padres.length) {
    throw new Error(`No se encontró el registro padre del código "${parentCode}" en Códigos Documentos.`);
  }

  padres.sort((a, b) => Number(a?.Id || 0) - Number(b?.Id || 0));
  return Number(padres[0].Id || 0);
}

function buildCodigoPadre(categoriaDocumento: string, codigoBase: string, correlativoPadre: number): string {
  return `${categoriaDocumento}-${codigoBase}-${String(correlativoPadre).padStart(3, '0')}`;
}

function buildCodigoHijo(codigoBase: string, correlativoPadre: number, codigoTipoDocumento: string, correlativoHijo: number): string {
  return `DS-${codigoBase}-${String(correlativoPadre).padStart(3, '0')}-${codigoTipoDocumento}-${String(correlativoHijo).padStart(2, '0')}`;
}

async function registerCodigoDocumentoWithRetry(params: {
  context: WebPartContext;
  webUrl: string;
  solicitud: any;
  solicitudId: number;
  log?: LogFn;
}): Promise<{ codigo: string; registroId: number; }> {
  const log = params.log || (() => undefined);
  const tipoDocumentoData = await getCodigosTipoDocumentoByTipoDocumentoId(
    params.context,
    params.webUrl,
    'Solicitudes',
    Number(params.solicitud?.TipoDocumentoId || 0)
  );
  const categoriaDocumento = tipoDocumentoData.codigoCategoria;
  const codigoTipoDocumento = tipoDocumentoData.codigoTipoDocumento;

  if (!['DO', 'DE', 'DS'].includes(categoriaDocumento)) {
    throw new Error(`La categoría "${categoriaDocumento}" no es válida para generar el código.`);
  }

  const codigoProceso = categoriaDocumento === 'DS'
    ? ''
    : await getCodigoProcesoByProcesoCorporativoId(params.context, params.webUrl, Number(params.solicitud?.ProcesoDeNegocioId || 0));

  if (categoriaDocumento !== 'DS' && !codigoProceso) {
    throw new Error(`La solicitud ${params.solicitudId} no tiene CodigoProceso resoluble.`);
  }

  const maxRetries = 8;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    let codigo = '';
    let payload: any = {};

    if (categoriaDocumento === 'DO' || categoriaDocumento === 'DE') {
      const codigoBase = `${codigoTipoDocumento}-${codigoProceso}`;
      const correlativoPadre = await obtenerSiguienteCorrelativoPadre(params.context, params.webUrl, codigoBase, categoriaDocumento);
      codigo = buildCodigoPadre(categoriaDocumento, codigoBase, correlativoPadre);
      payload = {
        Title: codigo,
        CodigoBase: codigoBase,
        CategoriaDocumento: categoriaDocumento,
        CorrelativoPadre: correlativoPadre,
        TipoSoporte: null,
        CorrelativoHijo: null
      };
    } else {
      const parentSolicitudId = await resolveParentSolicitudIdForDs(params.context, params.webUrl, params.solicitudId, params.solicitud);
      const codigoDocumentoPadreId = await resolveCodigoDocumentoRegistroPadreId(params.context, params.webUrl, parentSolicitudId);
      const codigoPadre = await spGetJson<any>(
        params.context,
        `${params.webUrl}/_api/web/lists/getbytitle('${escapeODataValue(CODIGOS_DOCUMENTOS_LIST)}')/items(${codigoDocumentoPadreId})?$select=Id,CodigoBase,CorrelativoPadre`
      );
      const correlativoHijo = await obtenerSiguienteCorrelativoHijo(params.context, params.webUrl, codigoDocumentoPadreId, codigoTipoDocumento);
      codigo = buildCodigoHijo(String(codigoPadre?.CodigoBase || '').trim(), Number(codigoPadre?.CorrelativoPadre || 0), codigoTipoDocumento, correlativoHijo);
      payload = {
        Title: codigo,
        CodigoBase: String(codigoPadre?.CodigoBase || '').trim(),
        CategoriaDocumento: categoriaDocumento,
        CorrelativoPadre: Number(codigoPadre?.CorrelativoPadre || 0),
        TipoSoporte: codigoTipoDocumento,
        CorrelativoHijo: correlativoHijo,
        DocumentoPadreId: codigoDocumentoPadreId
      };
    }

    const existentesAntes = await getCodigosDocumentosByTitle(params.context, params.webUrl, codigo);
    if (existentesAntes.length) {
      log(`⚠️ Código ya existente en Códigos Documentos, reintentando | ${codigo}`);
      continue;
    }

    const creado = await spPostJson<any>(
      params.context,
      params.webUrl,
      `${params.webUrl}/_api/web/lists/getbytitle('${escapeODataValue(CODIGOS_DOCUMENTOS_LIST)}')/items`,
      payload,
      'POST'
    );
    const registroId = Number(creado?.Id || 0);
    if (!registroId) {
      throw new Error(`No se pudo registrar el código "${codigo}" en Códigos Documentos.`);
    }

    const existentesDespues = await getCodigosDocumentosByTitle(params.context, params.webUrl, codigo);
    existentesDespues.sort((a, b) => Number(a?.Id || 0) - Number(b?.Id || 0));
    const ganadorId = Number(existentesDespues[0]?.Id || 0);

    if (ganadorId === registroId) {
      return {
        codigo,
        registroId
      };
    }

    await deleteListItem(params.context, params.webUrl, CODIGOS_DOCUMENTOS_LIST, registroId);
    log(`⚠️ Colisión concurrente al reservar código, se liberó el registro ${registroId} y se reintenta.`);
  }

  throw new Error(`No se pudo reservar un código único en Códigos Documentos para la solicitud ${params.solicitudId} tras varios reintentos.`);
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

async function downloadBlobByServerRelativeUrl(webUrl: string, fileUrl: string): Promise<Blob> {
  const response = await fetch(ensureAbsoluteUrl(webUrl, fileUrl), {
    credentials: 'same-origin'
  });

  if (!response.ok) {
    throw new Error(`No se pudo descargar el archivo "${fileUrl}". HTTP ${response.status}`);
  }

  return response.blob();
}

async function convertOfficeFileToPdfAndUpload(params: {
  context: WebPartContext;
  webUrl: string;
  sourceServerRelativeUrl: string;
  destinoFolderServerRelativeUrl: string;
  outputPdfName: string;
  log?: LogFn;
}): Promise<void> {
  const log = params.log || (() => undefined);
  const origin = new URL(params.webUrl).origin;
  const absoluteUrl = `${origin}${params.sourceServerRelativeUrl.startsWith('/') ? '' : '/'}${params.sourceServerRelativeUrl}`;
  const shareId = buildGraphShareIdFromUrl(absoluteUrl);
  const client = await params.context.aadHttpClientFactory.getClient('https://graph.microsoft.com');
  const response = await client.get(
    `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem/content?format=pdf`,
    AadHttpClient.configurations.v1
  );

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`No se pudo convertir a PDF (${response.status}): ${text}`);
  }

  const pdfBlob = new Blob([await response.arrayBuffer()], { type: 'application/pdf' });
  await uploadFileToFolder(
    params.context,
    params.webUrl,
    params.destinoFolderServerRelativeUrl,
    params.outputPdfName,
    pdfBlob
  );
  log(`📄 Archivo en Procesos reemplazado en PDF | ${params.outputPdfName}`);
}

async function publishReplacementToProcesos(params: {
  context: WebPartContext;
  webUrl: string;
  tempFileUrl: string;
  targetProcessFileUrl: string;
  log?: LogFn;
}): Promise<void> {
  const { folder, fileName } = splitFolderAndFile(params.targetProcessFileUrl);
  const targetIsPdf = /\.pdf$/i.test(fileName);

  if (!folder || !fileName) {
    throw new Error(`Ruta inválida en Procesos: ${params.targetProcessFileUrl}`);
  }

  if (targetIsPdf) {
    await convertOfficeFileToPdfAndUpload({
      context: params.context,
      webUrl: params.webUrl,
      sourceServerRelativeUrl: params.tempFileUrl,
      destinoFolderServerRelativeUrl: folder,
      outputPdfName: fileName,
      log: params.log
    });
    return;
  }

  const tempBlob = await downloadBlobByServerRelativeUrl(params.webUrl, params.tempFileUrl);
  await uploadFileToFolder(params.context, params.webUrl, folder, fileName, tempBlob);
  params.log?.(`📄 Archivo en Procesos reemplazado | ${fileName}`);
}

function isPdfConversionFailure(error: any): boolean {
  const message = error instanceof Error ? error.message : String(error || '');
  const trace = error?.executionTrace as IExecutionTrace | undefined;
  return (trace?.pasoFallido || '') === 'ReemplazarProceso' && /No se pudo convertir a PDF/i.test(message);
}

function shouldPreserveTempWithoutRollback(trace: IExecutionTrace | undefined, error: any): boolean {
  return !!trace?.tempFileUrl && isPdfConversionFailure(error);
}

async function enrichRowInfo(params: {
  context: WebPartContext;
  webUrl: string;
  input: IInputRow;
  rowLabel?: string;
  log?: LogFn;
}): Promise<ISolicitudRuntimeInfo> {
  const log = params.log || (() => undefined);
  const rowLabel = params.rowLabel || `Fila ${params.input.rowIndex}`;

  log(`🔎 ${rowLabel} | Leyendo solicitud ${params.input.id}...`);
  const solicitud = await getSolicitudById(params.context, params.webUrl, params.input.id);

  log(`🧩 ${rowLabel} | Resolviendo tipo documental...`);
  const tipoDocumentoData = await getCodigosTipoDocumentoByTipoDocumentoId(
    params.context,
    params.webUrl,
    'Solicitudes',
    Number(solicitud?.TipoDocumentoId || 0)
  );

  log(`👨‍👩‍👧 ${rowLabel} | Leyendo relaciones de hijos...`);
  const childIds = await getChildSolicitudIdsByParent(params.context, params.webUrl, params.input.id);

  log(`🔗 ${rowLabel} | Armando relacionados (${childIds.length})...`);
  const related = await getSolicitudesRelacionadas(params.context, params.webUrl, childIds);

  log(`🗺️ ${rowLabel} | Leyendo diagramas de flujo...`);
  const diagramas = await getDiagramasFlujoRowsBySolicitud(params.context, params.webUrl, params.input.id);

  log(`📄 ${rowLabel} | Buscando archivo vigente en Procesos...`);
  const processFile = await getCurrentProcessFileBySolicitudId(params.context, params.webUrl, params.input.id);

  log(`📎 ${rowLabel} | Leyendo adjuntos de la solicitud...`);
  const attachments = await getAttachmentFiles(params.context, params.webUrl, 'Solicitudes', params.input.id);
  const attachment = pickEditableAttachment(attachments as any);

  log(
    `✅ ${rowLabel} | Enriquecida | Categoria=${tipoDocumentoData.codigoCategoria || ''} | ` +
    `Hijos=${childIds.length} | Diagramas=${diagramas.length} | ` +
    `Proceso=${processFile?.FileRef ? 'Sí' : 'No'} | AdjuntoEditable=${attachment?.FileName ? 'Sí' : 'No'}`
  );

  return {
    rowIndex: params.input.rowIndex,
    id: params.input.id,
    codigoDocumento: params.input.codigoDocumento,
    versionDocumento: params.input.versionDocumento,
    created: params.input.created,
    title: params.input.title,
    nombreDocumento: params.input.nombreDocumento,
    procesar: params.input.procesar,
    categoriaDocumento: tipoDocumentoData.codigoCategoria || '',
    solicitud,
    childIds,
    diagramas,
    related,
    processFile,
    attachment,
    hasDependencias: childIds.length > 0 || diagramas.length > 0
  };
}

async function applyCodigoChange(params: {
  context: WebPartContext;
  webUrl: string;
  info: ISolicitudRuntimeInfo;
  rowLabel?: string;
  log?: LogFn;
}): Promise<{ attachmentFileName: string; processFileUrl: string; tempFileUrl: string; nuevoCodigo: string; codigoRegistroId: number; trace: IExecutionTrace; }> {
  const log = params.log || (() => undefined);
  const rowLabel = params.rowLabel || `Fila ${params.info.rowIndex}`;
  const solicitudId = params.info.id;
  const solicitud = params.info.solicitud;
  const processFile = params.info.processFile;
  const attachment = params.info.attachment;

  if (!solicitud?.Id) {
    throw new Error(`No se encontró la solicitud ${solicitudId}.`);
  }

  if (!solicitud?.EsVersionActualDocumento) {
    throw new Error(`La solicitud ${solicitudId} no está vigente.`);
  }

  if (!attachment?.ServerRelativeUrl || !attachment?.FileName) {
    throw new Error(`La solicitud ${solicitudId} no tiene adjunto editable compatible.`);
  }

  if (!processFile?.FileRef) {
    throw new Error(`La solicitud ${solicitudId} no tiene archivo vigente en Procesos.`);
  }

  const oldCodigo = String(solicitud.CodigoDocumento || '').trim();
  const attachmentBlobOriginal = await downloadBlobByServerRelativeUrl(params.webUrl, attachment.ServerRelativeUrl);
  const processBlobOriginal = await downloadBlobByServerRelativeUrl(params.webUrl, processFile.FileRef);
  const processMetadataOriginal = await getFileItemMetadata(params.context, params.webUrl, processFile.FileRef);
  const tempDestino = joinFolder(TEMP_WORD_ROOT, getRelativeFolderWithinProcesos(processFile.FileRef));
  const impactAreaIds = Array.isArray(solicitud?.AreasImpactadas)
    ? solicitud.AreasImpactadas.map((item: any) => Number(item?.Id || 0)).filter((id: number) => id > 0)
    : [];
  const duenoDocumento = await getAreaGerenteNombre(params.context, params.webUrl, Number(solicitud?.AreaDuenaId || 0));

  let solicitudUpdated = false;
  let attachmentUpdated = false;
  let processUpdated = false;
  let processMetadataUpdated = false;
  let tempFileUrl = '';
  let nuevoCodigo = '';
  let codigoRegistroId = 0;
  let diagramasParaDocumento = params.info.diagramas.map((item) => ({ ...item }));
  const diagramRollbackEntries: Array<{ itemId: number; oldCodigo: string; }> = [];
  const diagramCodigoRegistroIds: number[] = [];
  const trace: IExecutionTrace = {
    pasoFallido: '',
    codigoRegistroId: '',
    codigoReservado: 'No',
    adjuntoRegenerado: 'No',
    procesoReemplazado: 'No',
    solicitudActualizada: 'No',
    metadataProcesosActualizada: 'No',
    tempFileUrl: '',
    rollbackOmitido: 'No',
    rollbackCodigoRegistro: 'No',
    rollbackAdjunto: 'No',
    rollbackProceso: 'No',
    rollbackSolicitud: 'No',
    rollbackMetadataProcesos: 'No'
  };

  try {
    trace.pasoFallido = 'ReservarCodigo';
    log(`🏷️ ${rowLabel} | Reservando nuevo código...`);
    const registroCodigo = await registerCodigoDocumentoWithRetry({
      context: params.context,
      webUrl: params.webUrl,
      solicitud,
      solicitudId,
      log
    });
    nuevoCodigo = registroCodigo.codigo;
    codigoRegistroId = registroCodigo.registroId;
    trace.codigoRegistroId = codigoRegistroId;
    trace.codigoReservado = 'Si';

    log(`🛠️ ${rowLabel} | Solicitud ${solicitudId} | Código actual="${oldCodigo}" | Nuevo="${nuevoCodigo}" | RegistroCodigo=${codigoRegistroId}`);

    if (params.info.diagramas.length) {
      trace.pasoFallido = 'ActualizarDiagramasFlujo';
      log(`🧭 ${rowLabel} | Actualizando códigos de ${params.info.diagramas.length} diagrama(s) con base en el nuevo código padre...`);
      const parentCodigoRegistro = await getCodigoDocumentoRegistroById(params.context, params.webUrl, codigoRegistroId);
      const sortedDiagramas = params.info.diagramas.slice().sort((a, b) => a.id - b.id);
      diagramasParaDocumento = [];

      for (let i = 0; i < sortedDiagramas.length; i++) {
        const diagrama = sortedDiagramas[i];
        const codigoDiagrama = await registerCodigoDiagramaWithRetry({
          context: params.context,
          webUrl: params.webUrl,
          parentCodigoRegistroId: parentCodigoRegistro.Id,
          parentCodigoCompleto: nuevoCodigo,
          parentCodigoBase: parentCodigoRegistro.CodigoBase,
          correlativoPadre: parentCodigoRegistro.CorrelativoPadre,
          log
        });

        diagramCodigoRegistroIds.push(codigoDiagrama.registroId);
        diagramRollbackEntries.push({
          itemId: diagrama.id,
          oldCodigo: diagrama.codigo
        });

        await updateListItem(params.context, params.webUrl, 'Diagramas de Flujo', diagrama.id, {
          Codigo: codigoDiagrama.codigo
        });

        diagramasParaDocumento.push({
          ...diagrama,
          codigo: codigoDiagrama.codigo
        });

        log(`🧭 ${rowLabel} | Diagrama ${diagrama.id} actualizado | Código="${codigoDiagrama.codigo}"`);
      }

    }

    trace.pasoFallido = 'RegenerarAdjunto';
    log(`🧾 ${rowLabel} | Regenerando adjunto...`);
    const attachResult = await fillAndAttachFromServerRelativeUrl({
      context: params.context,
      webUrl: params.webUrl,
      listTitle: 'Solicitudes',
      itemId: solicitudId,
      originalFileName: attachment.FileName,
      sourceFileServerRelativeUrl: attachment.ServerRelativeUrl,
      titulo: solicitud?.NombreDocumento || solicitud?.Title || '',
      instanciaRaw: solicitud?.Instanciasdeaprobacion?.Title || 'Gerencia de Área',
      impactAreaIds,
      dueno: duenoDocumento,
      fechaVigencia: solicitud?.FechadeVigencia || '',
      fechaAprobacion: solicitud?.FechaDeAprobacionSolicitud || '',
      resumen: solicitud?.ResumenDocumento || '',
      version: solicitud?.VersionDocumento || '1.0',
      codigoDocumento: nuevoCodigo,
      relacionados: params.info.related.map((item) => ({ codigo: item.codigo, nombre: item.nombre, enlace: item.enlace })),
      diagramasFlujo: diagramasParaDocumento.map((item) => ({ codigo: item.codigo, nombre: item.nombre, enlace: item.enlace })),
      tempDestinoFolderServerRelativeUrl: tempDestino,
      replaceIfExists: true,
      log
    });

    if (!attachResult.ok || !attachResult.tempFileServerRelativeUrl) {
      throw new Error(attachResult.error || `No se pudo regenerar el adjunto de la solicitud ${solicitudId}.`);
    }

    attachmentUpdated = true;
    tempFileUrl = attachResult.tempFileServerRelativeUrl;
    trace.adjuntoRegenerado = 'Si';
    trace.tempFileUrl = tempFileUrl;

    trace.pasoFallido = 'ReemplazarProceso';
    log(`📚 ${rowLabel} | Reemplazando archivo en Procesos...`);
    await publishReplacementToProcesos({
      context: params.context,
      webUrl: params.webUrl,
      tempFileUrl,
      targetProcessFileUrl: processFile.FileRef,
      log
    });
    processUpdated = true;
    trace.procesoReemplazado = 'Si';

    trace.pasoFallido = 'ActualizarSolicitud';
    log(`📝 ${rowLabel} | Actualizando Solicitudes...`);
    await updateListItem(params.context, params.webUrl, 'Solicitudes', solicitudId, {
      CodigoDocumento: nuevoCodigo
    });
    solicitudUpdated = true;
    trace.solicitudActualizada = 'Si';

    trace.pasoFallido = 'ActualizarMetadataProcesos';
    log(`🗂️ ${rowLabel} | Actualizando metadata en Procesos...`);
    await updateFileMetadataByPath(params.context, params.webUrl, processFile.FileRef, {
      Codigodedocumento: nuevoCodigo
    });
    processMetadataUpdated = true;
    trace.metadataProcesosActualizada = 'Si';
    trace.pasoFallido = '';
    log(`✅ ${rowLabel} | Solicitud ${solicitudId} actualizada correctamente.`);

    return {
      attachmentFileName: attachment.FileName,
      processFileUrl: processFile.FileRef,
      tempFileUrl,
      nuevoCodigo,
      codigoRegistroId,
      trace
    };
  } catch (error) {
    if (shouldPreserveTempWithoutRollback(trace, error)) {
      trace.rollbackOmitido = 'Si';
      log(`⚠️ ${rowLabel} | Falló la conversión a PDF. Se conserva el documento generado en TEMP y no se ejecuta rollback.`);
      (error as any).executionTrace = trace;
      throw error;
    }

    if (diagramRollbackEntries.length) {
      for (let i = 0; i < diagramRollbackEntries.length; i++) {
        try {
          await updateListItem(params.context, params.webUrl, 'Diagramas de Flujo', diagramRollbackEntries[i].itemId, {
            Codigo: diagramRollbackEntries[i].oldCodigo
          });
        } catch (_rollbackError) {
          // sin acción
        }
      }
    }

    if (processMetadataUpdated) {
      try {
        await updateFileMetadataByPath(params.context, params.webUrl, processFile.FileRef, {
          Codigodedocumento: processMetadataOriginal?.Codigodedocumento || oldCodigo
        });
        trace.rollbackMetadataProcesos = 'Si';
      } catch (_rollbackError) {
        // sin acción
      }
    }

    if (solicitudUpdated) {
      try {
        await updateListItem(params.context, params.webUrl, 'Solicitudes', solicitudId, {
          CodigoDocumento: oldCodigo
        });
        trace.rollbackSolicitud = 'Si';
      } catch (_rollbackError) {
        // sin acción
      }
    }

    if (processUpdated) {
      try {
        const target = splitFolderAndFile(processFile.FileRef);
        await uploadFileToFolder(params.context, params.webUrl, target.folder, target.fileName, processBlobOriginal);
        trace.rollbackProceso = 'Si';
      } catch (_rollbackError) {
        // sin acción
      }
    }

    if (attachmentUpdated) {
      try {
        await deleteAttachment(params.context, params.webUrl, 'Solicitudes', solicitudId, attachment.FileName);
        await addAttachment(params.context, params.webUrl, 'Solicitudes', solicitudId, attachment.FileName, attachmentBlobOriginal);
        trace.rollbackAdjunto = 'Si';
      } catch (_rollbackError) {
        // sin acción
      }
    }

    if (codigoRegistroId) {
      try {
        await deleteListItem(params.context, params.webUrl, CODIGOS_DOCUMENTOS_LIST, codigoRegistroId);
        trace.rollbackCodigoRegistro = 'Si';
      } catch (_rollbackError) {
        // sin acción
      }
    }

    for (let i = 0; i < diagramCodigoRegistroIds.length; i++) {
      try {
        await deleteListItem(params.context, params.webUrl, CODIGOS_DOCUMENTOS_LIST, diagramCodigoRegistroIds[i]);
      } catch (_rollbackError) {
        // sin acción
      }
    }

    (error as any).executionTrace = trace;
    throw error;
  }
}

async function applyRenovarSolicitud(params: {
  context: WebPartContext;
  webUrl: string;
  info: ISolicitudRuntimeInfo;
  rowLabel?: string;
  log?: LogFn;
}): Promise<{ attachmentFileName: string; processFileUrl: string; tempFileUrl: string; trace: IExecutionTrace; }> {
  const log = params.log || (() => undefined);
  const rowLabel = params.rowLabel || `Fila ${params.info.rowIndex}`;
  const solicitud = params.info.solicitud;
  const solicitudId = params.info.id;
  const attachment = params.info.attachment;

  if (!solicitud?.Id) {
    throw new Error(`No se encontró la solicitud ${solicitudId}.`);
  }

  if (!attachment?.ServerRelativeUrl || !attachment?.FileName) {
    throw new Error(`La solicitud ${solicitudId} no tiene adjunto compatible para renovar.`);
  }

  const historicoFile = await getHistoricoFileForRenew(
    params.context,
    params.webUrl,
    solicitud,
    attachment.FileName,
    log
  );
  if (!historicoFile?.FileRef) {
    throw new Error(`No se encontró archivo en históricos para la solicitud ${solicitudId}.`);
  }

  const historicoMetadata = await getFileItemMetadata(params.context, params.webUrl, historicoFile.FileRef);
  const relativeFolder = getRelativeFolderWithinHistoricos(historicoFile.FileRef);
  const targetFolder = joinFolder(PROCESOS_ROOT, relativeFolder);
  const restoredHistoricalName = removeHistoricalSuffix(historicoFile.FileLeafRef);
  const outputFileName = /\.docx$/i.test(attachment.FileName)
    ? replaceExtension(restoredHistoricalName || attachment.FileName, '.pdf')
    : (restoredHistoricalName || attachment.FileName);
  const processFileUrl = `${targetFolder}/${outputFileName}`;
  const tempDestino = joinFolder(TEMP_WORD_ROOT, relativeFolder);
  const attachmentBlob = await downloadBlobByServerRelativeUrl(params.webUrl, attachment.ServerRelativeUrl);
  const tempFileUrl = await uploadFileToFolder(
    params.context,
    params.webUrl,
    tempDestino,
    attachment.FileName,
    attachmentBlob
  );

  const trace: IExecutionTrace = {
    pasoFallido: '',
    codigoRegistroId: '',
    codigoReservado: 'No',
    adjuntoRegenerado: 'No',
    procesoReemplazado: 'No',
    solicitudActualizada: 'No',
    metadataProcesosActualizada: 'No',
    tempFileUrl,
    rollbackOmitido: 'No',
    rollbackCodigoRegistro: 'No',
    rollbackAdjunto: 'No',
    rollbackProceso: 'No',
    rollbackSolicitud: 'No',
    rollbackMetadataProcesos: 'No'
  };

  log(`📂 ${rowLabel} | Archivo renovado copiado a TEMP | ${tempFileUrl}`);

  trace.pasoFallido = 'ActualizarSolicitud';
  await updateListItem(params.context, params.webUrl, 'Solicitudes', solicitudId, {
    EsVersionActualDocumento: true
  });
  trace.solicitudActualizada = 'Si';
  log(`📝 ${rowLabel} | Solicitud ${solicitudId} marcada como versión actual.`);

  try {
    trace.pasoFallido = 'ReemplazarProceso';
    log(`📚 ${rowLabel} | Publicando documento renovado en Procesos...`);
    await publishReplacementToProcesos({
      context: params.context,
      webUrl: params.webUrl,
      tempFileUrl,
      targetProcessFileUrl: processFileUrl,
      log
    });
    trace.procesoReemplazado = 'Si';

    trace.pasoFallido = 'ActualizarMetadataProcesos';
    await updateProcesosMetadataForRenew({
      context: params.context,
      webUrl: params.webUrl,
      targetFileUrl: processFileUrl,
      solicitud,
      baseMetadata: historicoMetadata
    });
    trace.metadataProcesosActualizada = 'Si';
    log(`🗂️ ${rowLabel} | Metadata de Procesos actualizada para renovación.`);

    trace.pasoFallido = 'EliminarHistorico';
    await recycleFile(params.context, params.webUrl, historicoFile.FileRef);
    trace.pasoFallido = '';
    log(`🗑️ ${rowLabel} | Histórico eliminado | ${historicoFile.FileRef}`);

    return {
      attachmentFileName: attachment.FileName,
      processFileUrl,
      tempFileUrl,
      trace
    };
  } catch (error) {
    if (shouldPreserveTempWithoutRollback(trace, error)) {
      trace.rollbackOmitido = 'Si';
      log(`⚠️ ${rowLabel} | Falló la conversión a PDF en renovación. Se conserva el archivo en TEMP y no se ejecuta rollback.`);
    }

    (error as any).executionTrace = trace;
    throw error;
  }
}

function buildOutputFileName(originalFileName: string): string {
  const dotIndex = String(originalFileName || '').lastIndexOf('.');
  if (dotIndex === -1) {
    return `${originalFileName}_correccion_codigos_duplicados.xlsx`;
  }

  return `${originalFileName.substring(0, dotIndex)}_correccion_codigos_duplicados${originalFileName.substring(dotIndex)}`;
}

export async function ejecutarCorreccionCodigosDuplicadosDesdeExcel(params: {
  context: WebPartContext;
  excelFile: IFilePickerResult;
  log?: LogFn;
}): Promise<{
  blob: Blob;
  fileName: string;
  processed: number;
  updated: number;
  skipped: number;
  error: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const session = await openExcelRevisionSession(params.excelFile);
  const grid = session.grid || [];

  if (!grid.length) {
    throw new Error('El Excel está vacío.');
  }

  const outputGrid = cloneGrid(grid);
  const headers = outputGrid[0] || [];
  const headerMap = buildHeaderMap(headers);

  const idxEstado = headers.length;
  const idxDetalle = headers.length + 1;
  const idxNuevoCodigo = headers.length + 2;
  const idxSolicitud = headers.length + 3;
  const idxProceso = headers.length + 4;
  const idxAdjunto = headers.length + 5;
  const idxPasoFallido = headers.length + 6;
  const idxArchivoTemp = headers.length + 7;
  const idxCodigoRegistroId = headers.length + 8;
  const idxCodigoReservado = headers.length + 9;
  const idxAdjuntoRegenerado = headers.length + 10;
  const idxProcesoReemplazado = headers.length + 11;
  const idxSolicitudActualizada = headers.length + 12;
  const idxMetadataProcesosActualizada = headers.length + 13;
  const idxRollbackOmitido = headers.length + 14;
  const idxRollbackCodigoRegistro = headers.length + 15;
  const idxRollbackAdjunto = headers.length + 16;
  const idxRollbackProceso = headers.length + 17;
  const idxRollbackSolicitud = headers.length + 18;
  const idxRollbackMetadataProcesos = headers.length + 19;

  headers[idxEstado] = 'EstadoCorreccionCodigoDuplicado';
  headers[idxDetalle] = 'DetalleCorreccionCodigoDuplicado';
  headers[idxNuevoCodigo] = 'NuevoCodigoDocumento';
  headers[idxSolicitud] = 'SolicitudIdCorregida';
  headers[idxProceso] = 'ArchivoProcesoCorregido';
  headers[idxAdjunto] = 'AdjuntoCorregido';
  headers[idxPasoFallido] = 'PasoFallido';
  headers[idxArchivoTemp] = 'ArchivoTempGenerado';
  headers[idxCodigoRegistroId] = 'CodigoRegistroId';
  headers[idxCodigoReservado] = 'CodigoReservado';
  headers[idxAdjuntoRegenerado] = 'AdjuntoRegenerado';
  headers[idxProcesoReemplazado] = 'ProcesoReemplazado';
  headers[idxSolicitudActualizada] = 'SolicitudActualizada';
  headers[idxMetadataProcesosActualizada] = 'MetadataProcesosActualizada';
  headers[idxRollbackOmitido] = 'RollbackOmitido';
  headers[idxRollbackCodigoRegistro] = 'RollbackCodigoRegistro';
  headers[idxRollbackAdjunto] = 'RollbackAdjunto';
  headers[idxRollbackProceso] = 'RollbackProceso';
  headers[idxRollbackSolicitud] = 'RollbackSolicitud';
  headers[idxRollbackMetadataProcesos] = 'RollbackMetadataProcesos';
  outputGrid[0] = headers;

  const inputRows: IInputRow[] = [];
  for (let rowIndex = 1; rowIndex < grid.length; rowIndex++) {
    const row = grid[rowIndex] || [];
    const id = Number(getCellByCandidates(row, headerMap, ['Id']) || 0);
    const codigoDocumento = String(getCellByCandidates(row, headerMap, ['CodigoDocumento', 'ClaveCodigoDocumento']) || '').trim();
    const versionDocumento = String(getCellByCandidates(row, headerMap, ['VersionDocumento']) || '').trim();
    const title = String(getCellByCandidates(row, headerMap, ['Title']) || '').trim();
    const nombreDocumento = String(getCellByCandidates(row, headerMap, ['NombreDocumento']) || '').trim();
    const created = String(getCellByCandidates(row, headerMap, ['Created']) || '').trim();
    const procesarRaw = String(getCellByCandidates(row, headerMap, ['Procesar']) || '').trim();
    const procesarMode = getProcesarMode(procesarRaw);

    if (!id && !codigoDocumento && !title && !nombreDocumento) {
      continue;
    }

    inputRows.push({
      rowIndex,
      row,
      id,
      codigoDocumento,
      versionDocumento,
      title,
      nombreDocumento,
      created,
      procesar: procesarMode !== '',
      procesarRaw,
      procesarMode
    });
  }

  if (!inputRows.length) {
    throw new Error('El Excel no tiene filas utilizables.');
  }

  log(`📄 Excel cargado: ${session.fileName}`);
  log(`📋 Filas detectadas: ${inputRows.length}`);

  let processed = 0;
  let updated = 0;
  let skipped = 0;
  let error = 0;

  for (let i = 0; i < inputRows.length; i++) {
    const row = outputGrid[inputRows[i].rowIndex] || [];
    const rowLabel = `Fila ${i + 1}/${inputRows.length}`;

    if (!inputRows[i].procesar) {
      row[idxEstado] = 'SKIP';
      row[idxDetalle] = `La columna Procesar está en "${inputRows[i].procesarRaw || '(vacía)'}". Solo se procesa cuando vale SI o Renovar.`;
      outputGrid[inputRows[i].rowIndex] = row;
      skipped++;
      continue;
    }

    if (!inputRows[i].id) {
      row[idxEstado] = 'ERROR';
      row[idxDetalle] = 'La fila no tiene Id de solicitud.';
      outputGrid[inputRows[i].rowIndex] = row;
      error++;
      continue;
    }

    if (!inputRows[i].codigoDocumento) {
      row[idxEstado] = 'ERROR';
      row[idxDetalle] = 'La fila no tiene CodigoDocumento.';
      outputGrid[inputRows[i].rowIndex] = row;
      error++;
      continue;
    }

    if (!isVersionUno(inputRows[i].versionDocumento)) {
      row[idxEstado] = 'SKIP';
      row[idxDetalle] = `La versión ${inputRows[i].versionDocumento || '(vacía)'} no es 1.0.`;
      outputGrid[inputRows[i].rowIndex] = row;
      skipped++;
      continue;
    }

    try {
      log(`⏳ ${rowLabel} | Preparando solicitud ${inputRows[i].id || ''}...`);
      const info = await enrichRowInfo({
        context: params.context,
        webUrl,
        input: inputRows[i],
        rowLabel,
        log
      });

      if (String(info.categoriaDocumento || '').trim().toUpperCase() === 'DS') {
        row[idxEstado] = 'SKIP';
        row[idxDetalle] = 'Documento hijo (DS) omitido por ahora.';
        row[idxSolicitud] = info.id;
        outputGrid[inputRows[i].rowIndex] = row;
        skipped++;
        continue;
      }

      if (inputRows[i].procesarMode === 'si' && info.childIds.length > 0) {
        row[idxEstado] = 'SKIP';
        row[idxDetalle] = 'La solicitud tiene documentos hijos; se omite por ahora.';
        row[idxSolicitud] = info.id;
        outputGrid[inputRows[i].rowIndex] = row;
        skipped++;
        continue;
      }

      processed++;
      const applyRowLabel = `FilaExcel ${info.rowIndex}`;

      try {
        log(`🚀 ${applyRowLabel} | Aplicando cambio de código a solicitud ${info.id}...`);
        const result = inputRows[i].procesarMode === 'renovar'
          ? await applyRenovarSolicitud({
            context: params.context,
            webUrl,
            info,
            rowLabel: applyRowLabel,
            log
          })
          : await applyCodigoChange({
            context: params.context,
            webUrl,
            info,
            rowLabel: applyRowLabel,
            log
          });

        row[idxEstado] = 'OK';
        row[idxDetalle] = inputRows[i].procesarMode === 'renovar'
          ? 'Se renovó la solicitud: se marcó como versión actual, se publicó nuevamente en Procesos y se eliminó el histórico asociado.'
          : info.diagramas.length
            ? 'Se cambió el código del documento y también los códigos de sus diagramas de flujo asociados.'
            : 'Se cambió el código del documento seleccionado.';
        row[idxNuevoCodigo] = (result as any).nuevoCodigo || info.codigoDocumento;
        row[idxSolicitud] = info.id;
        row[idxProceso] = result.processFileUrl;
        row[idxAdjunto] = result.attachmentFileName;
        row[idxPasoFallido] = '';
        row[idxArchivoTemp] = result.trace.tempFileUrl;
        row[idxCodigoRegistroId] = result.trace.codigoRegistroId;
        row[idxCodigoReservado] = result.trace.codigoReservado;
        row[idxAdjuntoRegenerado] = result.trace.adjuntoRegenerado;
        row[idxProcesoReemplazado] = result.trace.procesoReemplazado;
        row[idxSolicitudActualizada] = result.trace.solicitudActualizada;
        row[idxMetadataProcesosActualizada] = result.trace.metadataProcesosActualizada;
        row[idxRollbackOmitido] = result.trace.rollbackOmitido;
        row[idxRollbackCodigoRegistro] = result.trace.rollbackCodigoRegistro;
        row[idxRollbackAdjunto] = result.trace.rollbackAdjunto;
        row[idxRollbackProceso] = result.trace.rollbackProceso;
        row[idxRollbackSolicitud] = result.trace.rollbackSolicitud;
        row[idxRollbackMetadataProcesos] = result.trace.rollbackMetadataProcesos;
        outputGrid[inputRows[i].rowIndex] = row;
        updated++;
      } catch (rowError) {
        const message = rowError instanceof Error ? rowError.message : String(rowError);
        const trace = (rowError as any)?.executionTrace as IExecutionTrace | undefined;
        row[idxEstado] = 'ERROR';
        row[idxDetalle] = shouldPreserveTempWithoutRollback(trace, rowError)
          ? `No se pudo convertir a PDF. Se conserva el documento generado en TEMP y no se ejecutó rollback. Revisar archivo temporal para gestión manual. Error: ${message}`
          : message;
        row[idxSolicitud] = info.id;
        row[idxPasoFallido] = trace?.pasoFallido || '';
        row[idxArchivoTemp] = trace?.tempFileUrl || '';
        row[idxCodigoRegistroId] = trace?.codigoRegistroId || '';
        row[idxCodigoReservado] = trace?.codigoReservado || 'No';
        row[idxAdjuntoRegenerado] = trace?.adjuntoRegenerado || 'No';
        row[idxProcesoReemplazado] = trace?.procesoReemplazado || 'No';
        row[idxSolicitudActualizada] = trace?.solicitudActualizada || 'No';
        row[idxMetadataProcesosActualizada] = trace?.metadataProcesosActualizada || 'No';
        row[idxRollbackOmitido] = trace?.rollbackOmitido || 'No';
        row[idxRollbackCodigoRegistro] = trace?.rollbackCodigoRegistro || 'No';
        row[idxRollbackAdjunto] = trace?.rollbackAdjunto || 'No';
        row[idxRollbackProceso] = trace?.rollbackProceso || 'No';
        row[idxRollbackSolicitud] = trace?.rollbackSolicitud || 'No';
        row[idxRollbackMetadataProcesos] = trace?.rollbackMetadataProcesos || 'No';
        outputGrid[inputRows[i].rowIndex] = row;
        error++;
        log(
          shouldPreserveTempWithoutRollback(trace, rowError)
            ? `❌ Solicitud ${info.id} | No se pudo convertir a PDF. Se dejó el documento en TEMP sin rollback | ${trace?.tempFileUrl || ''}`
            : `❌ Solicitud ${info.id} | ${message}`
        );
      }
    } catch (rowError) {
      const message = rowError instanceof Error ? rowError.message : String(rowError);
      row[idxEstado] = 'ERROR';
      row[idxDetalle] = message;
      outputGrid[inputRows[i].rowIndex] = row;
      error++;
    }
  }

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(outputGrid);
  XLSX.utils.book_append_sheet(workbook, worksheet, session.sheetName || 'Resultado');

  const output = XLSX.write(workbook, {
    bookType: String(session.fileName || '').toLowerCase().endsWith('.xlsm') ? 'xlsm' : 'xlsx',
    type: 'array'
  });

  return {
    blob: new Blob([output], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }),
    fileName: buildOutputFileName(session.fileName),
    processed,
    updated,
    skipped,
    error
  };
}
