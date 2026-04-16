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
};

type ISolicitudRuntimeInfo = {
  rowIndex: number;
  id: number;
  codigoDocumento: string;
  versionDocumento: string;
  created: string;
  title: string;
  nombreDocumento: string;
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

function normKey(value: any): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
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

function getRelativeFolderWithinProcesos(fileUrl: string): string {
  const full = trimSlash(fileUrl);
  const root = trimSlash(PROCESOS_ROOT);
  const fileDir = full.substring(0, full.lastIndexOf('/'));
  return fileDir.indexOf(root) === 0 ? fileDir.substring(root.length).replace(/^\/+/, '') : '';
}

function isVersionUno(value: any): boolean {
  return String(value || '').trim() === '1.0';
}

function parseDateSafe(value: any): number {
  const text = String(value || '').trim();
  if (!text) {
    return 0;
  }

  const ddmmyyyy = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (ddmmyyyy) {
    const day = Number(ddmmyyyy[1]);
    const month = Number(ddmmyyyy[2]) - 1;
    const year = Number(ddmmyyyy[3].length === 2 ? `20${ddmmyyyy[3]}` : ddmmyyyy[3]);
    const hour = Number(ddmmyyyy[4] || 0);
    const minute = Number(ddmmyyyy[5] || 0);
    const second = Number(ddmmyyyy[6] || 0);
    return new Date(year, month, day, hour, minute, second).getTime() || 0;
  }

  const iso = Date.parse(text);
  return Number.isFinite(iso) ? iso : 0;
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
    `?$select=Id,Title,NombreDocumento,CodigoDocumento,ResumenDocumento,VersionDocumento,TipoDocumentoId,ProcesoDeNegocioId,` +
    `FechaDeAprobacionSolicitud,FechadeVigencia,AreaDuenaId,Instanciasdeaprobacion/Title,` +
    `AreasImpactadas/Id,AreasImpactadas/Title,EsVersionActualDocumento` +
    `&$expand=Instanciasdeaprobacion,AreasImpactadas`;
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
    `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapeODataValue(fileUrl)}')/ListItemAllFields?$select=Id,Codigodedocumento`
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

function chooseGroupTargets(group: ISolicitudRuntimeInfo[]): {
  keeperId?: number;
  changeIds: number[];
  skipReason?: string;
} {
  const sorted = group.slice().sort((a, b) => {
    const createdDiff = parseDateSafe(a.created) - parseDateSafe(b.created);
    if (createdDiff !== 0) {
      return createdDiff;
    }

    return a.id - b.id;
  });

  const withDependencias = sorted.filter((item) => item.hasDependencias);
  if (withDependencias.length > 1) {
    return {
      changeIds: [],
      skipReason: 'El grupo tiene más de una solicitud con hijos o diagramas; se omite por ahora.'
    };
  }

  const keeper = withDependencias.length === 1 ? withDependencias[0] : sorted[0];
  return {
    keeperId: keeper.id,
    changeIds: sorted.filter((item) => item.id !== keeper.id).map((item) => item.id)
  };
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
      diagramasFlujo: params.info.diagramas.map((item) => ({ codigo: item.codigo, nombre: item.nombre, enlace: item.enlace })),
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
      created
    });
  }

  if (!inputRows.length) {
    throw new Error('El Excel no tiene filas utilizables.');
  }

  const liveInfos: ISolicitudRuntimeInfo[] = [];

  log(`📄 Excel cargado: ${session.fileName}`);
  log(`📋 Filas detectadas: ${inputRows.length}`);

  let processed = 0;
  let updated = 0;
  let skipped = 0;
  let error = 0;

  for (let i = 0; i < inputRows.length; i++) {
    const row = outputGrid[inputRows[i].rowIndex] || [];
    const rowLabel = `Fila ${i + 1}/${inputRows.length}`;

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

      liveInfos.push(info);
    } catch (rowError) {
      const message = rowError instanceof Error ? rowError.message : String(rowError);
      row[idxEstado] = 'ERROR';
      row[idxDetalle] = message;
      outputGrid[inputRows[i].rowIndex] = row;
      error++;
    }
  }

  const groups = new Map<string, ISolicitudRuntimeInfo[]>();
  for (let i = 0; i < liveInfos.length; i++) {
    const key = normKey(liveInfos[i].codigoDocumento);
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key)!.push(liveInfos[i]);
  }

  for (const group of Array.from(groups.values())) {
    if (group.length < 2) {
      for (let i = 0; i < group.length; i++) {
        const row = outputGrid[group[i].rowIndex] || [];
        row[idxEstado] = 'SKIP';
        row[idxDetalle] = 'El código ya no aparece duplicado en las filas válidas del Excel.';
        outputGrid[group[i].rowIndex] = row;
        skipped++;
      }
      continue;
    }

    const decision = chooseGroupTargets(group);
    if (decision.skipReason) {
      for (let i = 0; i < group.length; i++) {
        const row = outputGrid[group[i].rowIndex] || [];
        row[idxEstado] = 'SKIP';
        row[idxDetalle] = decision.skipReason;
        outputGrid[group[i].rowIndex] = row;
        skipped++;
      }
      continue;
    }

    let currentKeeperId = decision.keeperId;
    let keeperMoved = false;
    const swappableKeeper = !!group.find((item) => item.id === decision.keeperId && !item.hasDependencias);

    for (let i = 0; i < group.length; i++) {
      const info = group[i];
      if (decision.changeIds.indexOf(info.id) === -1) {
        continue;
      }

      const row = outputGrid[info.rowIndex] || [];
      processed++;
      const rowLabel = `FilaExcel ${info.rowIndex}`;

      try {
        log(`🚀 ${rowLabel} | Aplicando cambio de código a solicitud ${info.id}...`);
        const result = await applyCodigoChange({
          context: params.context,
          webUrl,
          info,
          rowLabel,
          log
        });

        row[idxEstado] = 'OK';
        row[idxDetalle] = info.hasDependencias
          ? 'Se cambió el código porque en el grupo solo una solicitud podía conservar el código original.'
          : 'Se cambió el código por ser una solicitud posterior dentro del grupo duplicado.';
        row[idxNuevoCodigo] = result.nuevoCodigo;
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
        outputGrid[info.rowIndex] = row;
        updated++;
      } catch (rowError) {
        if (!shouldPreserveTempWithoutRollback((rowError as any)?.executionTrace as IExecutionTrace | undefined, rowError)
          && isPdfConversionFailure(rowError) && swappableKeeper && !keeperMoved && currentKeeperId) {
          const keeperInfo = group.find((item) => item.id === currentKeeperId);
          if (keeperInfo) {
            const keeperRow = outputGrid[keeperInfo.rowIndex] || [];
            const failedTrace = (rowError as any)?.executionTrace as IExecutionTrace | undefined;
            const failedMessage = rowError instanceof Error ? rowError.message : String(rowError);

            log(
              `⚠️ Solicitud ${info.id} | Falló la conversión a PDF. ` +
              `Se intentará cambiar la solicitud ${keeperInfo.id}, dejando ${info.id} con el código original.`
            );

            processed++;
            try {
              const fallbackLabel = `${rowLabel} | Fallback`;
              const fallbackResult = await applyCodigoChange({
                context: params.context,
                webUrl,
                info: keeperInfo,
                rowLabel: fallbackLabel,
                log
              });

              keeperMoved = true;
              currentKeeperId = info.id;

              keeperRow[idxEstado] = 'OK';
              keeperRow[idxDetalle] = 'Se cambió el código como alternativa porque la otra solicitud duplicada no pudo convertirse a PDF.';
              keeperRow[idxNuevoCodigo] = fallbackResult.nuevoCodigo;
              keeperRow[idxSolicitud] = keeperInfo.id;
              keeperRow[idxProceso] = fallbackResult.processFileUrl;
              keeperRow[idxAdjunto] = fallbackResult.attachmentFileName;
              keeperRow[idxPasoFallido] = '';
              keeperRow[idxArchivoTemp] = fallbackResult.trace.tempFileUrl;
              keeperRow[idxCodigoRegistroId] = fallbackResult.trace.codigoRegistroId;
              keeperRow[idxCodigoReservado] = fallbackResult.trace.codigoReservado;
              keeperRow[idxAdjuntoRegenerado] = fallbackResult.trace.adjuntoRegenerado;
              keeperRow[idxProcesoReemplazado] = fallbackResult.trace.procesoReemplazado;
              keeperRow[idxSolicitudActualizada] = fallbackResult.trace.solicitudActualizada;
              keeperRow[idxMetadataProcesosActualizada] = fallbackResult.trace.metadataProcesosActualizada;
              keeperRow[idxRollbackOmitido] = fallbackResult.trace.rollbackOmitido;
              keeperRow[idxRollbackCodigoRegistro] = fallbackResult.trace.rollbackCodigoRegistro;
              keeperRow[idxRollbackAdjunto] = fallbackResult.trace.rollbackAdjunto;
              keeperRow[idxRollbackProceso] = fallbackResult.trace.rollbackProceso;
              keeperRow[idxRollbackSolicitud] = fallbackResult.trace.rollbackSolicitud;
              keeperRow[idxRollbackMetadataProcesos] = fallbackResult.trace.rollbackMetadataProcesos;
              outputGrid[keeperInfo.rowIndex] = keeperRow;
              updated++;

              row[idxEstado] = 'KEEP';
              row[idxDetalle] = `Conserva el código porque la conversión a PDF falló al intentar cambiar esta solicitud, y se cambió la otra del grupo. Error original: ${failedMessage}`;
              row[idxNuevoCodigo] = info.codigoDocumento;
              row[idxSolicitud] = info.id;
              row[idxPasoFallido] = failedTrace?.pasoFallido || '';
              row[idxArchivoTemp] = failedTrace?.tempFileUrl || '';
              row[idxCodigoRegistroId] = failedTrace?.codigoRegistroId || '';
              row[idxCodigoReservado] = failedTrace?.codigoReservado || 'No';
              row[idxAdjuntoRegenerado] = failedTrace?.adjuntoRegenerado || 'No';
              row[idxProcesoReemplazado] = failedTrace?.procesoReemplazado || 'No';
              row[idxSolicitudActualizada] = failedTrace?.solicitudActualizada || 'No';
              row[idxMetadataProcesosActualizada] = failedTrace?.metadataProcesosActualizada || 'No';
              row[idxRollbackOmitido] = failedTrace?.rollbackOmitido || 'No';
              row[idxRollbackCodigoRegistro] = failedTrace?.rollbackCodigoRegistro || 'No';
              row[idxRollbackAdjunto] = failedTrace?.rollbackAdjunto || 'No';
              row[idxRollbackProceso] = failedTrace?.rollbackProceso || 'No';
              row[idxRollbackSolicitud] = failedTrace?.rollbackSolicitud || 'No';
              row[idxRollbackMetadataProcesos] = failedTrace?.rollbackMetadataProcesos || 'No';
              outputGrid[info.rowIndex] = row;
              skipped++;
              continue;
            } catch (fallbackError) {
              const fallbackMessage = fallbackError instanceof Error ? fallbackError.message : String(fallbackError);
              log(`❌ Solicitud ${keeperInfo.id} | También falló el fallback del grupo: ${fallbackMessage}`);
            }
          }
        }

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
        outputGrid[info.rowIndex] = row;
        error++;
        log(
          shouldPreserveTempWithoutRollback(trace, rowError)
            ? `❌ Solicitud ${info.id} | No se pudo convertir a PDF. Se dejó el documento en TEMP sin rollback | ${trace?.tempFileUrl || ''}`
            : `❌ Solicitud ${info.id} | ${message}`
        );
      }
    }

    if (currentKeeperId) {
      const keeperInfo = group.find((item) => item.id === currentKeeperId);
      if (keeperInfo) {
        const keeperRow = outputGrid[keeperInfo.rowIndex] || [];
        if (!String(keeperRow[idxEstado] || '').trim()) {
          keeperRow[idxEstado] = 'KEEP';
          keeperRow[idxDetalle] = keeperInfo.hasDependencias
            ? 'Conserva el código porque es la única solicitud con hijos o diagramas.'
            : keeperMoved
              ? 'Conserva el código original porque la otra solicitud del grupo no pudo convertirse a PDF.'
              : 'Conserva el código por ser la solicitud más antigua del grupo.';
          keeperRow[idxNuevoCodigo] = keeperInfo.codigoDocumento;
          keeperRow[idxSolicitud] = keeperInfo.id;
          outputGrid[keeperInfo.rowIndex] = keeperRow;
          skipped++;
        }
      }
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
