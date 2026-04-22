/* eslint-disable */
// @ts-nocheck
import { AadHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { fillAndAttachFromServerRelativeUrl } from './documentFillAndAttachFromUrl.service';
import {
  escapeODataValue,
  getAllItems,
  getAttachmentFiles,
  spGetJson,
  spPostJson,
  updateListItem,
  uploadFileToFolder
} from './sharepointRest.service';

type LogFn = (message: string) => void;
type IRelacionadoRow = { solicitudId: number; codigo: string; nombre: string; enlace: string; };
type IDiagramaRow = { id: number; codigo: string; nombre: string; enlace: string; };

const PROCESOS_ROOT = '/sites/SistemadeGestionDocumental/Procesos';
const TEMP_WORD_ROOT = '/sites/SistemadeGestionDocumental/Procesos/TEMP_MIGRACION_WORD';
const CODIGOS_DOCUMENTOS_LIST = 'Códigos Documentos';
const TARGET_DIAGRAMA_ID = 1085;

function trimSlash(value: string): string {
  return String(value || '').replace(/\/+$/, '');
}

function joinFolder(base: string, relative: string): string {
  const cleanBase = trimSlash(base);
  const cleanRelative = String(relative || '').replace(/^\/+/, '').replace(/\/+$/, '');
  return cleanRelative ? `${cleanBase}/${cleanRelative}` : cleanBase;
}

function getRelativeFolderWithinProcesos(fileUrl: string): string {
  const full = trimSlash(fileUrl);
  const root = trimSlash(PROCESOS_ROOT);
  const fileDir = full.substring(0, full.lastIndexOf('/'));
  return fileDir.indexOf(root) === 0 ? fileDir.substring(root.length).replace(/^\/+/, '') : '';
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

function base64UrlEncode(str: string): string {
  const b64 = btoa(unescape(encodeURIComponent(str)));
  return b64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function buildGraphShareIdFromUrl(absoluteUrl: string): string {
  const safe = encodeURI(absoluteUrl);
  return 'u!' + base64UrlEncode(safe);
}

function buildSolicitudDocumentLink(webUrl: string, codigo: string): string {
  return codigo
    ? `${new URL(webUrl).origin}/sites/SistemadeGestionDocumental/SitePages/verdocumento-vigente.aspx?codigodocumentosgd=${encodeURIComponent(codigo)}`
    : '';
}

function obtenerPrimerNombreYApellido(displayName: string): string {
  const limpio = (displayName || '').replace(/\s+/g, ' ').trim();
  if (!limpio) return '';
  const partes = limpio.split(' ').filter(Boolean);
  if (partes.length === 3) return `${partes[0]} ${partes[1]}`;
  if (partes.length === 4) return `${partes[0]} ${partes[2]}`;
  if (partes.length >= 2) return `${partes[0]} ${partes[1]}`;
  return partes[0] || '';
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

  if (value && Array.isArray(value.results)) {
    const result: number[] = [];
    for (let i = 0; i < value.results.length; i++) {
      const id = Number(value.results[i] || 0);
      if (id > 0 && result.indexOf(id) === -1) {
        result.push(id);
      }
    }
    return result;
  }

  const single = Number(value);
  return Number.isFinite(single) && single > 0 ? [single] : [];
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

async function getAreaGerenteNombre(context: WebPartContext, webUrl: string, areaId: number): Promise<string> {
  if (!areaId) return '';
  const item = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Áreas de Negocio')/items(${areaId})?$select=Id,Gerente/Title&$expand=Gerente`
  );
  return obtenerPrimerNombreYApellido(String(item?.Gerente?.Title || '').trim());
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
        enlace: buildSolicitudDocumentLink(webUrl, codigo)
      });
    } catch (_error) {
      // omitir relacionado invalido
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
      log(`⚠️ Codigo de diagrama ya existente en Codigos Documentos, reintentando | ${codigo}`);
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
      throw new Error(`No se pudo registrar el codigo de diagrama "${codigo}" en Codigos Documentos.`);
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
  }

  throw new Error(`No se pudo reservar un codigo unico de diagrama para el padre ${params.parentCodigoCompleto}.`);
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
    throw new Error(`No se encontro el registro padre del codigo "${parentCode}" en Codigos Documentos.`);
  }

  padres.sort((a, b) => Number(a?.Id || 0) - Number(b?.Id || 0));
  return Number(padres[0].Id || 0);
}

async function downloadBlobByServerRelativeUrl(webUrl: string, fileUrl: string): Promise<Blob> {
  const response = await fetch(`${new URL(webUrl).origin}${fileUrl.startsWith('/') ? '' : '/'}${fileUrl}`, {
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
    throw new Error(`Ruta invalida en Procesos: ${params.targetProcessFileUrl}`);
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

async function updateProcesosMetadataForSolicitud(params: {
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

type IGroupedDiagramas = {
  solicitudId: number;
  items: Array<{ id: number; title: string; codigo: string; solicitudIds: number[]; }>;
};

async function getDiagramasSinCodigoAgrupados(
  context: WebPartContext,
  webUrl: string
): Promise<{ groups: IGroupedDiagramas[]; totalSinCodigo: number; skippedAmbiguous: number; }> {
  const solicitudField = await getFieldInternalName(context, webUrl, 'Diagramas de Flujo', 'Solicitud');
  const solicitudFieldId = `${solicitudField}Id`;
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Diagramas de Flujo')/items?$select=Id,Title,Codigo,${solicitudFieldId}&$top=5000`
  );

  const bySolicitud = new Map<number, IGroupedDiagramas>();
  let totalSinCodigo = 0;
  let skippedAmbiguous = 0;

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const itemId = Number(item?.Id || 0);
    if (itemId !== TARGET_DIAGRAMA_ID) {
      continue;
    }

    const codigo = String(item?.Codigo || '').trim();
    if (codigo) {
      continue;
    }

    totalSinCodigo++;
    const solicitudIds = normalizeLookupIds(item?.[solicitudFieldId]);
    if (solicitudIds.length !== 1) {
      skippedAmbiguous++;
      continue;
    }

    const solicitudId = solicitudIds[0];
    if (!bySolicitud.has(solicitudId)) {
      bySolicitud.set(solicitudId, {
        solicitudId,
        items: []
      });
    }

    bySolicitud.get(solicitudId)!.items.push({
      id: itemId,
      title: String(item?.Title || '').trim(),
      codigo,
      solicitudIds
    });
  }

  return {
    groups: Array.from(bySolicitud.values()).sort((a, b) => a.solicitudId - b.solicitudId),
    totalSinCodigo,
    skippedAmbiguous
  };
}

export async function ejecutarBuscarDiagramaFlujoSinCodigo(params: {
  context: WebPartContext;
  log?: LogFn;
}): Promise<{
  totalDiagramasSinCodigo: number;
  solicitudesProcesadas: number;
  diagramasCorregidos: number;
  solicitudesRegeneradas: number;
  skipped: number;
  error: number;
  tempFileUrls: string[];
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const tempFileUrls: string[] = [];

  let solicitudesProcesadas = 0;
  let diagramasCorregidos = 0;
  let solicitudesRegeneradas = 0;
  let skipped = 0;
  let error = 0;

  log('🔎 Buscando diagramas de flujo sin codigo...');
  const agrupados = await getDiagramasSinCodigoAgrupados(params.context, webUrl);
  skipped += agrupados.skippedAmbiguous;

  log(`📊 Diagramas sin codigo encontrados: ${agrupados.totalSinCodigo}`);
  log(`📦 Solicitudes con diagramas faltantes: ${agrupados.groups.length}`);
  if (agrupados.skippedAmbiguous) {
    log(`⚠️ Diagramas omitidos por Solicitud vacia o multiple: ${agrupados.skippedAmbiguous}`);
  }

  for (let i = 0; i < agrupados.groups.length; i++) {
    const group = agrupados.groups[i];
    const rowLabel = `Solicitud ${group.solicitudId}`;
    solicitudesProcesadas++;

    try {
      log(`🔧 ${rowLabel} | Iniciando correccion de ${group.items.length} diagrama(s) sin codigo...`);
      const solicitud = await getSolicitudById(params.context, webUrl, group.solicitudId);
      const parentCode = String(solicitud?.CodigoDocumento || '').trim();
      if (!parentCode) {
        skipped++;
        log(`⏭️ ${rowLabel} | Omitida porque la solicitud no tiene CodigoDocumento.`);
        continue;
      }

      const processFile = await getCurrentProcessFileBySolicitudId(params.context, webUrl, group.solicitudId);
      if (!processFile?.FileRef) {
        throw new Error(`La solicitud ${group.solicitudId} no tiene archivo vigente en Procesos.`);
      }

      const attachments = await getAttachmentFiles(params.context, webUrl, 'Solicitudes', group.solicitudId);
      const attachment = pickEditableAttachment(attachments as any);
      if (!attachment?.ServerRelativeUrl || !attachment?.FileName) {
        throw new Error(`La solicitud ${group.solicitudId} no tiene adjunto editable compatible.`);
      }

      const parentCodigoRegistroId = await resolveCodigoDocumentoRegistroPadreId(params.context, webUrl, group.solicitudId);
      const parentCodigoRegistro = await getCodigoDocumentoRegistroById(params.context, webUrl, parentCodigoRegistroId);
      const tipoDocumento = await getCodigosTipoDocumentoByTipoDocumentoId(
        params.context,
        webUrl,
        'Solicitudes',
        Number(solicitud?.TipoDocumentoId || 0)
      );

      if (String(tipoDocumento.codigoCategoria || '').trim().toUpperCase() === 'DS') {
        log(`⚠️ ${rowLabel} | La solicitud es DS. Se usara igualmente su codigo padre "${parentCode}" para generar los codigos DF.`);
      }

      const sortedMissing = group.items.slice().sort((a, b) => a.id - b.id);
      for (let j = 0; j < sortedMissing.length; j++) {
        const item = sortedMissing[j];
        const reservado = await registerCodigoDiagramaWithRetry({
          context: params.context,
          webUrl,
          parentCodigoRegistroId,
          parentCodigoCompleto: parentCode,
          parentCodigoBase: parentCodigoRegistro.CodigoBase,
          correlativoPadre: parentCodigoRegistro.CorrelativoPadre,
          log
        });

        await updateListItem(params.context, webUrl, 'Diagramas de Flujo', item.id, {
          Codigo: reservado.codigo
        });
        diagramasCorregidos++;
        log(`🧭 ${rowLabel} | Diagrama ${item.id} corregido | Codigo="${reservado.codigo}" | RegistroCodigo=${reservado.registroId}`);
      }

      const childIds = await getChildSolicitudIdsByParent(params.context, webUrl, group.solicitudId);
      const hasHijos = childIds.length > 0;
      log(`👨‍👧 ${rowLabel} | Validacion previa de regeneracion | TieneHijos=${hasHijos ? 'Si' : 'No'} | TotalHijos=${childIds.length}`);
      const relacionados = await getSolicitudesRelacionadas(params.context, webUrl, childIds);
      const diagramasActualizados = await getDiagramasFlujoRowsBySolicitud(params.context, webUrl, group.solicitudId);
      const processMetadata = await getFileItemMetadata(params.context, webUrl, processFile.FileRef);
      const impactAreaIds = Array.isArray(solicitud?.AreasImpactadas)
        ? solicitud.AreasImpactadas.map((item: any) => Number(item?.Id || 0)).filter((id: number) => id > 0)
        : [];
      const duenoDocumento = await getAreaGerenteNombre(params.context, webUrl, Number(solicitud?.AreaDuenaId || 0));
      const tempDestino = joinFolder(TEMP_WORD_ROOT, getRelativeFolderWithinProcesos(processFile.FileRef));

      log(`🧾 ${rowLabel} | Regenerando adjunto del padre con diagramas corregidos...`);
      const attachResult = await fillAndAttachFromServerRelativeUrl({
        context: params.context,
        webUrl,
        listTitle: 'Solicitudes',
        itemId: group.solicitudId,
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
        codigoDocumento: parentCode,
        relacionados: hasHijos
          ? relacionados.map((item) => ({ codigo: item.codigo, nombre: item.nombre, enlace: item.enlace }))
          : [],
        diagramasFlujo: diagramasActualizados.map((item) => ({ codigo: item.codigo, nombre: item.nombre, enlace: item.enlace })),
        tempDestinoFolderServerRelativeUrl: tempDestino,
        replaceIfExists: true,
        log
      });

      if (!attachResult.ok || !attachResult.tempFileServerRelativeUrl) {
        throw new Error(attachResult.error || `No se pudo regenerar el adjunto de la solicitud ${group.solicitudId}.`);
      }

      tempFileUrls.push(attachResult.tempFileServerRelativeUrl);
      log(`📚 ${rowLabel} | Republicando documento en Procesos...`);
      await publishReplacementToProcesos({
        context: params.context,
        webUrl,
        tempFileUrl: attachResult.tempFileServerRelativeUrl,
        targetProcessFileUrl: processFile.FileRef,
        log
      });

      await updateProcesosMetadataForSolicitud({
        context: params.context,
        webUrl,
        targetFileUrl: processFile.FileRef,
        solicitud,
        baseMetadata: processMetadata
      });

      solicitudesRegeneradas++;
      log(`✅ ${rowLabel} | Correccion completa | DiagramasCorregidos=${group.items.length} | TEMP=${attachResult.tempFileServerRelativeUrl}`);
    } catch (groupError) {
      error++;
      const message = groupError instanceof Error ? groupError.message : String(groupError || '');
      log(`❌ ${rowLabel} | Error corrigiendo diagramas sin codigo: ${message}`);
    }
  }

  log(
    `🏁 Buscar Diagrama de Flujo sin codigo finalizado | ` +
    `SolicitudesProcesadas=${solicitudesProcesadas} | ` +
    `DiagramasCorregidos=${diagramasCorregidos} | ` +
    `SolicitudesRegeneradas=${solicitudesRegeneradas} | ` +
    `SKIP=${skipped} | ERROR=${error}`
  );

  return {
    totalDiagramasSinCodigo: agrupados.totalSinCodigo,
    solicitudesProcesadas,
    diagramasCorregidos,
    solicitudesRegeneradas,
    skipped,
    error,
    tempFileUrls
  };
}
