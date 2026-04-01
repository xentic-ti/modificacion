/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';
import { AadHttpClient } from '@microsoft/sp-http';
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ensureFolderPath, escapeODataValue, recycleFile, spGetJson, spPostJson, uploadFileToFolder, getAttachmentFiles } from './sharepointRest.service';
import { listFilesRecursive } from './spFolderExplorer.service';
import { IFase2RollbackEntry } from './modificacionFase2Publicacion.service';
import { IFase2PublicacionReportRow } from '../utils/fase2PublicacionReportExcel';

type LogFn = (s: string) => void;

const PROCESOS_ROOT = '/sites/SistemadeGestionDocumental/Procesos';
const TEMP_WORD_ROOT = '/sites/SistemadeGestionDocumental/Procesos/TEMP_MIGRACION_WORD';
const HISTORICOS_ROOT = '/sites/SistemadeGestionDocumental/Documentos Histricos';

type IParentBridgeRow = {
  SolicitudOrigenID: number;
  SolicitudID: number;
  NombreDocumento: string;
  NombreArchivo: string;
  CodigoDocumento: string;
  VersionDocumento: string;
  RutaTemporalWord: string;
  EstadoFase1: string;
  DocumentosHijosIDs: string;
  DocumentoPadreSolicitudAnteriorID: number;
  DocumentoPadreSolicitudNuevaID: number;
  Clasificaciondeproceso: string;
  Macroproceso: string;
  Proceso: string;
  Subproceso: string;
  AreaDuena: string;
  AreaImpactada: string;
  Resumen: string;
  CategoriaDocumento: string;
  TipoDocumento: string;
  FechaDeAprobacion: string;
  FechaDeVigencia: string;
  InstanciaDeAprobacionId: number | '';
};

type IChildTask = {
  childSolicitudId: number;
  parentMappings: Array<{
    oldParentSolicitudId: number;
    newParentSolicitudId: number;
    parentDocumentName: string;
  }>;
};

function normalizeHeader(value: any): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '')
    .trim()
    .toLowerCase();
}

function normalizeLooseFileKey(value: any): string {
  const raw = String(value ?? '').trim();
  const extensionMatch = raw.match(/(\.[^.]+)$/);
  const extension = extensionMatch ? extensionMatch[1].toLowerCase() : '';
  const baseName = extension ? raw.slice(0, -extension.length) : raw;

  const normalizedTokens = baseName
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map((token) => (token.length > 4 && token.endsWith('s') ? token.slice(0, -1) : token));

  return `${normalizedTokens.join('')}${extension}`;
}

function trimSlash(value: string): string {
  return String(value || '').replace(/\/+$/, '');
}

function joinFolder(base: string, relative: string): string {
  const cleanBase = trimSlash(base);
  const cleanRelative = String(relative || '').replace(/^\/+/, '').replace(/\/+$/, '');
  return cleanRelative ? `${cleanBase}/${cleanRelative}` : cleanBase;
}

function replaceExtension(name: string, extensionWithDot: string): string {
  const baseName = String(name || '').replace(/\.[^.]+$/, '');
  return `${baseName}${extensionWithDot}`;
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

function sanitizeVersion(value: any): string {
  return String(value || '').trim().replace(/^v/i, '') || '1.0';
}

function parseSlashIds(value: any): number[] {
  return String(value || '')
    .split('/')
    .map((part) => Number(String(part || '').trim()))
    .filter((id) => Number.isFinite(id) && id > 0);
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

function toDateOnlyIso(value: Date | null): string | null {
  if (!value || isNaN(value.getTime())) return null;
  return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0)).toISOString();
}

function base64UrlEncode(str: string): string {
  const b64 = btoa(unescape(encodeURIComponent(str)));
  return b64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
}

function buildGraphShareIdFromUrl(absoluteUrl: string): string {
  const safe = encodeURI(absoluteUrl);
  return 'u!' + base64UrlEncode(safe);
}

async function readArrayBufferFromFilePicker(file: IFilePickerResult): Promise<ArrayBuffer> {
  if (!file) throw new Error('Archivo Excel no recibido.');
  if (typeof file.downloadFileContent === 'function') {
    const blob = await file.downloadFileContent();
    return blob.arrayBuffer();
  }

  const url = (file as any).fileAbsoluteUrl || '';
  if (!url) throw new Error('No se pudo obtener el contenido del Excel.');

  const response = await fetch(url, { credentials: 'same-origin' });
  if (!response.ok) throw new Error(`No se pudo descargar el Excel. HTTP ${response.status}`);
  return response.arrayBuffer();
}

async function readFase3BridgeExcel(file: IFilePickerResult): Promise<IParentBridgeRow[]> {
  const buffer = await readArrayBufferFromFilePicker(file);
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new Error('No se encontró la hoja del Excel de Fase 3.');

  const worksheet = workbook.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false }) as any[][];
  if (!aoa.length) return [];

  const headers = aoa[0] || [];
  const headerMap = new Map<string, number>();
  for (let i = 0; i < headers.length; i++) headerMap.set(normalizeHeader(headers[i]), i);

  const getValue = (row: any[], header: string): any => {
    const index = headerMap.get(normalizeHeader(header));
    return index === undefined ? '' : row[index];
  };

  const rows: IParentBridgeRow[] = [];
  for (let i = 1; i < aoa.length; i++) {
    const row = aoa[i] || [];
    rows.push({
      SolicitudOrigenID: Number(getValue(row, 'SolicitudOrigenID') || 0),
      SolicitudID: Number(getValue(row, 'SolicitudID') || 0),
      NombreDocumento: String(getValue(row, 'NombreDocumento') || '').trim(),
      NombreArchivo: String(getValue(row, 'NombreArchivo') || '').trim(),
      CodigoDocumento: String(getValue(row, 'CodigoDocumento') || '').trim(),
      VersionDocumento: String(getValue(row, 'VersionDocumento') || '').trim(),
      RutaTemporalWord: String(getValue(row, 'RutaTemporalWord') || '').trim(),
      EstadoFase1: String(getValue(row, 'EstadoFase1') || '').trim(),
      DocumentosHijosIDs: String(getValue(row, 'DocumentosHijosIDs') || '').trim(),
      DocumentoPadreSolicitudAnteriorID: Number(getValue(row, 'DocumentoPadreSolicitudAnteriorID') || 0),
      DocumentoPadreSolicitudNuevaID: Number(getValue(row, 'DocumentoPadreSolicitudNuevaID') || 0),
      Clasificaciondeproceso: String(getValue(row, 'Clasificaciondeproceso') || '').trim(),
      Macroproceso: String(getValue(row, 'Macroproceso') || '').trim(),
      Proceso: String(getValue(row, 'Proceso') || '').trim(),
      Subproceso: String(getValue(row, 'Subproceso') || '').trim(),
      AreaDuena: String(getValue(row, 'AreaDuena') || '').trim(),
      AreaImpactada: String(getValue(row, 'AreaImpactada') || '').trim(),
      Resumen: String(getValue(row, 'Resumen') || '').trim(),
      CategoriaDocumento: String(getValue(row, 'CategoriaDocumento') || '').trim(),
      TipoDocumento: String(getValue(row, 'TipoDocumento') || '').trim(),
      FechaDeAprobacion: String(getValue(row, 'FechaDeAprobacion') || '').trim(),
      FechaDeVigencia: String(getValue(row, 'FechaDeVigencia') || '').trim(),
      InstanciaDeAprobacionId: Number(getValue(row, 'InstanciaDeAprobacionId') || 0) || ''
    });
  }

  return rows;
}

function buildChildTasks(rows: IParentBridgeRow[]): IChildTask[] {
  const byChild = new Map<number, IChildTask>();

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (String(row.EstadoFase1 || '').trim().toUpperCase() !== 'OK') continue;
    const childIds = parseSlashIds(row.DocumentosHijosIDs);
    if (!childIds.length) continue;

    for (let j = 0; j < childIds.length; j++) {
      const childId = childIds[j];
      if (!byChild.has(childId)) {
        byChild.set(childId, { childSolicitudId: childId, parentMappings: [] });
      }

      const task = byChild.get(childId)!;
      const exists = task.parentMappings.some(
        (item) =>
          item.oldParentSolicitudId === row.DocumentoPadreSolicitudAnteriorID &&
          item.newParentSolicitudId === row.DocumentoPadreSolicitudNuevaID
      );
      if (!exists) {
        task.parentMappings.push({
          oldParentSolicitudId: row.DocumentoPadreSolicitudAnteriorID,
          newParentSolicitudId: row.DocumentoPadreSolicitudNuevaID,
          parentDocumentName: row.NombreDocumento || ''
        });
      }
    }
  }

  return Array.from(byChild.values());
}

function getRelativeFolderWithinProcesos(fileUrl: string): string {
  const full = trimSlash(fileUrl);
  const root = trimSlash(PROCESOS_ROOT);
  const fileDir = full.substring(0, full.lastIndexOf('/'));
  return fileDir.indexOf(root) === 0 ? fileDir.substring(root.length).replace(/^\/+/, '') : '';
}

async function getFieldInfoByListPath(
  context: WebPartContext,
  webUrl: string,
  listPath: string,
  fieldInternalName: string
): Promise<any> {
  const url =
    `${webUrl}/_api/web/GetList('${escapeODataValue(listPath)}')/fields/getbyinternalnameortitle('${escapeODataValue(fieldInternalName)}')` +
    `?$select=TypeAsString,AllowMultipleValues,Choices`;
  return spGetJson<any>(context, url);
}

async function getFieldTypeFlagsForProcesos(context: WebPartContext, webUrl: string): Promise<{
  areaImpactadaIsMulti: boolean;
  documentoPadreIsMulti: boolean;
}> {
  const areaField = await getFieldInfoByListPath(context, webUrl, PROCESOS_ROOT, 'AreaImpactada');
  const documentoPadreField = await getFieldInfoByListPath(context, webUrl, PROCESOS_ROOT, 'DocumentoPadre');
  return {
    areaImpactadaIsMulti: String(areaField?.TypeAsString || '').toLowerCase().indexOf('multi') !== -1,
    documentoPadreIsMulti: !!documentoPadreField?.AllowMultipleValues
  };
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

async function updateFileMetadataByPath(context: WebPartContext, webUrl: string, fileUrl: string, payload: any): Promise<void> {
  await spPostJson(
    context,
    webUrl,
    `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapeODataValue(fileUrl)}')/ListItemAllFields`,
    payload,
    'MERGE'
  );
}

async function fileExistsByServerRelativeUrl(context: WebPartContext, webUrl: string, fileUrl: string): Promise<boolean> {
  try {
    await spGetJson<any>(
      context,
      `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapeODataValue(fileUrl)}')?$select=Exists`
    );
    return true;
  } catch (error: any) {
    const message = String(error?.message || '');
    if (message.indexOf('(404)') !== -1 || message.toLowerCase().indexOf('not found') !== -1 || message.toLowerCase().indexOf('no se encuentra') !== -1) {
      return false;
    }
    throw error;
  }
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
  const maxRetries = 3;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const client = await params.context.aadHttpClientFactory.getClient('https://graph.microsoft.com');
    const response = await client.get(`https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem/content?format=pdf`, AadHttpClient.configurations.v1);

    if (response.ok) {
      const pdfBuffer = await response.arrayBuffer();
      const pdfBlob = new Blob([pdfBuffer], { type: 'application/pdf' });
      await uploadFileToFolder(params.context, params.webUrl, params.destinoFolderServerRelativeUrl, params.outputPdfName, pdfBlob);
      log(`📄✅ Hijo publicado en PDF | ${params.outputPdfName}`);
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
      log(`⚠️ Reintentando conversión PDF hijo (${attempt}/${maxRetries}) | ${params.outputPdfName} | HTTP ${response.status}`);
      await new Promise((resolve) => setTimeout(resolve, attempt * 2000));
      continue;
    }

    throw new Error(`Graph PDF failed (${response.status}): ${body}`);
  }
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
      try {
        await copyFileByPath(params.context, params.webUrl, siblingPdfUrl, destinationFileUrl, false);
        params.log?.(`📄✅ Hijo publicado usando PDF existente | ${params.outputFileName}`);
      } catch (_error) {
        params.log?.(`⚠️ PDF hermano hijo no disponible al copiar, se intentará convertir el Word | ${params.outputFileName}`);
        await convertOfficeFileToPdfAndUpload({
          context: params.context,
          webUrl: params.webUrl,
          sourceServerRelativeUrl: params.sourceFileUrl,
          destinoFolderServerRelativeUrl: params.targetFolderUrl,
          outputPdfName: params.outputFileName,
          log: params.log
        });
      }
    } else {
      await convertOfficeFileToPdfAndUpload({
        context: params.context,
        webUrl: params.webUrl,
        sourceServerRelativeUrl: params.sourceFileUrl,
        destinoFolderServerRelativeUrl: params.targetFolderUrl,
        outputPdfName: params.outputFileName,
        log: params.log
      });
    }
  } else {
    await copyFileByPath(params.context, params.webUrl, params.sourceFileUrl, destinationFileUrl, false);
    params.log?.(`📄✅ Hijo publicado sin conversión | ${params.outputFileName}`);
  }
  return destinationFileUrl;
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

async function getChildSolicitudItem(context: WebPartContext, webUrl: string, solicitudId: number): Promise<any> {
  return spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items(${solicitudId})?$select=Id,Title,NombreDocumento,CodigoDocumento,VersionDocumento`
  );
}

async function getParentProcessFileItemIdBySolicitudId(
  context: WebPartContext,
  webUrl: string,
  solicitudId: number
): Promise<number | null> {
  const row = await getCurrentProcessFileBySolicitudId(context, webUrl, solicitudId);
  return row ? row.Id : null;
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

function replaceLookupIds(currentIds: number[], replacements: Map<number, number>): number[] {
  const replaced = currentIds.map((id) => (replacements.has(id) ? Number(replacements.get(id)) : id));
  return Array.from(new Set(replaced.filter((id) => Number.isFinite(id) && id > 0)));
}

function buildRenamedHistoricalFileName(originalName: string, oldVersion: string, todayStamp: string): string {
  const baseName = String(originalName || '').replace(/\.[^.]+$/, '');
  const extension = (String(originalName || '').match(/\.[^.]+$/) || [''])[0];
  return `${baseName}_V${sanitizeVersion(oldVersion)}_${todayStamp}${extension}`;
}

async function updateHistoricoMetadata(params: {
  context: WebPartContext;
  webUrl: string;
  historicoFileUrl: string;
  oldMetadata: any;
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
    HisFechaAprobacionBaja: params.oldMetadata?.FechaDeAprobacion || null
  };

  await updateFileMetadataByPath(params.context, params.webUrl, params.historicoFileUrl, payload);
}

async function updateChildProcesosMetadataAfterPublish(params: {
  context: WebPartContext;
  webUrl: string;
  targetFileUrl: string;
  oldMetadata: any;
  replacementParentIds: number[];
  areaImpactadaIsMulti: boolean;
  documentoPadreIsMulti: boolean;
}): Promise<void> {
  const areaImpactada = parseAreaImpactada(params.oldMetadata?.AreaImpactada);
  const payload: any = {
    Clasificaciondeproceso: params.oldMetadata?.Clasificaciondeproceso || '',
    AreaDuena: params.oldMetadata?.AreaDuena || '',
    VersionDocumento: params.oldMetadata?.VersionDocumento || '',
    AreaImpactada: params.areaImpactadaIsMulti ? areaImpactada : (areaImpactada[0] || ''),
    Macroproceso: params.oldMetadata?.Macroproceso || '',
    Proceso: params.oldMetadata?.Proceso || '',
    Subproceso: params.oldMetadata?.Subproceso || '',
    Tipodedocumento: params.oldMetadata?.Tipodedocumento || '',
    SolicitudId: Number(params.oldMetadata?.SolicitudId || 0) || null,
    Codigodedocumento: params.oldMetadata?.Codigodedocumento || '',
    Resumen: params.oldMetadata?.Resumen || '',
    CategoriaDocumento: params.oldMetadata?.CategoriaDocumento || '',
    FechaDeAprobacion: params.oldMetadata?.FechaDeAprobacion || null,
    FechaDePublicacion: new Date().toISOString(),
    FechaDeVigencia: params.oldMetadata?.FechaDeVigencia || null,
    InstanciaDeAprobacionId: Number(params.oldMetadata?.InstanciaDeAprobacionId || 0) || null,
    Accion: params.oldMetadata?.Accion || 'Actualización de documento',
    NombreDocumento: params.oldMetadata?.NombreDocumento || ''
  };

  if (params.replacementParentIds.length) {
    payload.DocumentoPadreId = params.documentoPadreIsMulti ? params.replacementParentIds : params.replacementParentIds[0];
  }

  await updateFileMetadataByPath(params.context, params.webUrl, params.targetFileUrl, payload);
}

async function updateParentProcesosMetadataAfterPublish(params: {
  context: WebPartContext;
  webUrl: string;
  targetFileUrl: string;
  row: IParentBridgeRow;
  areaImpactadaIsMulti: boolean;
}): Promise<void> {
  const fechaAprobacion = toDateOnlyIso(parseDdMmYyyyToDate(params.row.FechaDeAprobacion));
  const fechaVigencia = toDateOnlyIso(parseDdMmYyyyToDate(params.row.FechaDeVigencia));
  const areaImpactada = parseAreaImpactada(params.row.AreaImpactada);

  const payload: any = {
    Clasificaciondeproceso: params.row.Clasificaciondeproceso || '',
    AreaDuena: params.row.AreaDuena || '',
    VersionDocumento: params.row.VersionDocumento || '',
    AreaImpactada: params.areaImpactadaIsMulti ? areaImpactada : (areaImpactada[0] || ''),
    Macroproceso: params.row.Macroproceso || '',
    Proceso: params.row.Proceso || '',
    Subproceso: params.row.Subproceso || '',
    Tipodedocumento: params.row.TipoDocumento || '',
    SolicitudId: Number(params.row.SolicitudID || 0) || null,
    Codigodedocumento: params.row.CodigoDocumento || '',
    Resumen: params.row.Resumen || '',
    CategoriaDocumento: params.row.CategoriaDocumento || '',
    FechaDeAprobacion: fechaAprobacion,
    FechaDePublicacion: new Date().toISOString(),
    FechaDeVigencia: fechaVigencia,
    InstanciaDeAprobacionId: Number(params.row.InstanciaDeAprobacionId || 0) || null,
    Accion: 'Actualización de documento',
    NombreDocumento: params.row.NombreDocumento || ''
  };

  await updateFileMetadataByPath(params.context, params.webUrl, params.targetFileUrl, payload);
}

async function updateExistingChildProcessParentReferences(params: {
  context: WebPartContext;
  webUrl: string;
  childSolicitudId: number;
  oldParentSolicitudId: number;
  newParentSolicitudId: number;
  documentoPadreIsMulti: boolean;
  log?: LogFn;
}): Promise<boolean> {
  const currentFile = await getCurrentProcessFileBySolicitudId(params.context, params.webUrl, params.childSolicitudId);
  if (!currentFile?.FileRef) {
    params.log?.(`⚠️ Fase 4 | No se encontró el archivo vigente del hijo ${params.childSolicitudId} para actualizar referencia de padre.`);
    return false;
  }

  const oldParentFileItemId = await getParentProcessFileItemIdBySolicitudId(params.context, params.webUrl, params.oldParentSolicitudId);
  const newParentFileItemId = await getParentProcessFileItemIdBySolicitudId(params.context, params.webUrl, params.newParentSolicitudId);
  if (!oldParentFileItemId || !newParentFileItemId) {
    throw new Error(`No se pudieron resolver los archivos de padre para actualizar el hijo ${params.childSolicitudId}. PadreAntiguo=${params.oldParentSolicitudId} | PadreNuevo=${params.newParentSolicitudId}`);
  }

  const oldMetadata = await getFileItemMetadata(params.context, params.webUrl, currentFile.FileRef);
  const currentParentIds = normalizeLookupIds(oldMetadata?.DocumentoPadreId);
  const replacementParentIds = replaceLookupIds(currentParentIds, new Map<number, number>([[oldParentFileItemId, newParentFileItemId]]));

  const changed =
    replacementParentIds.length !== currentParentIds.length ||
    replacementParentIds.some((id, index) => id !== currentParentIds[index]);

  if (!changed) return false;

  await updateFileMetadataByPath(params.context, params.webUrl, currentFile.FileRef, {
    DocumentoPadreId: params.documentoPadreIsMulti ? replacementParentIds : (replacementParentIds[0] || null)
  });

  params.log?.(`👨‍👧 Fase 4 | Referencia de padre actualizada en hijo ${params.childSolicitudId} | PadreAntiguo=${params.oldParentSolicitudId} | PadreNuevo=${params.newParentSolicitudId}`);
  return true;
}

function getCandidateNamesFromSolicitud(childSolicitud: any, attachmentFiles: Array<{ FileName: string; ServerRelativeUrl: string; }>): string[] {
  const names: string[] = [];
  for (let i = 0; i < attachmentFiles.length; i++) {
    const fileName = String(attachmentFiles[i]?.FileName || '').trim();
    if (fileName) names.push(fileName);
  }

  const nombreDocumento = String(childSolicitud?.NombreDocumento || childSolicitud?.Title || '').trim();
  if (nombreDocumento) {
    names.push(`${nombreDocumento}.docx`);
    names.push(`${nombreDocumento}.xlsx`);
    names.push(`${nombreDocumento}.xlsm`);
    names.push(`${nombreDocumento}.pptx`);
    names.push(`${nombreDocumento}.pdf`);
  }

  return Array.from(new Set(names));
}

function pickBestSourceFile(
  candidateNames: string[],
  currentRelativeFolder: string,
  sourceIndex: Map<string, Array<{ Name: string; ServerRelativeUrl: string; }>>
): { Name: string; ServerRelativeUrl: string; } | null {
  const folderKey = normalizeLooseFileKey(currentRelativeFolder.replace(/\//g, ' '));
  const candidates: Array<{ Name: string; ServerRelativeUrl: string; score: number; }> = [];

  for (let i = 0; i < candidateNames.length; i++) {
    const key = normalizeLooseFileKey(candidateNames[i]);
    const matches = sourceIndex.get(key) || [];
    for (let j = 0; j < matches.length; j++) {
      const match = matches[j];
      let score = 0;
      if (String(match.Name || '').toLowerCase() === String(candidateNames[i] || '').toLowerCase()) score += 10;
      if (folderKey && normalizeLooseFileKey(match.ServerRelativeUrl).indexOf(folderKey) !== -1) score += 5;
      candidates.push({ ...match, score });
    }
  }

  if (!candidates.length) return null;
  candidates.sort((a, b) => b.score - a.score || a.ServerRelativeUrl.localeCompare(b.ServerRelativeUrl));
  return { Name: candidates[0].Name, ServerRelativeUrl: candidates[0].ServerRelativeUrl };
}

function buildRollbackEntry(
  childSolicitudId: number,
  oldOriginalUrl: string,
  oldRenamedUrl: string,
  historicoUrl: string,
  newPublishedUrl: string,
  oldOriginalMetadata: any
): IFase2RollbackEntry {
  return {
    solicitudOrigenId: childSolicitudId,
    solicitudId: childSolicitudId,
    nombreDocumento: oldOriginalMetadata?.NombreDocumento || oldOriginalMetadata?.Title || '',
    oldOriginalUrl,
    oldRenamedUrl,
    historicoUrl,
    newPublishedUrl,
    oldOriginalMetadata
  };
}

export async function ejecutarFase4PublicacionHijos(params: {
  context: WebPartContext;
  excelFile: IFilePickerResult;
  sourceFolderServerRelativeUrl?: string;
  log?: LogFn;
}): Promise<{
  rollbackEntries: IFase2RollbackEntry[];
  reportRows: IFase2PublicacionReportRow[];
  processed: number;
  ok: number;
  skipped: number;
  error: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const parentRows = await readFase3BridgeExcel(params.excelFile);
  const reportRows: IFase2PublicacionReportRow[] = [];
  const rollbackEntries: IFase2RollbackEntry[] = [];
  const fieldFlags = await getFieldTypeFlagsForProcesos(params.context, webUrl);
  const now = new Date();
  const todayStamp = buildTodayStamp(now);
  const todayText = buildTodayDdMmYyyy(now);

  let ok = 0;
  let skipped = 0;
  let error = 0;

  for (let i = 0; i < parentRows.length; i++) {
    const row = parentRows[i];
    const childIds = parseSlashIds(row.DocumentosHijosIDs);
    let oldOriginalUrl = '';
    let oldRenamedUrl = '';
    let historicoUrl = '';
    let newPublishedUrl = '';
    let oldMetadata: any = null;
    let renamedOld = false;
    let historicoCopied = false;
    let newPublished = false;

    try {
      if (!row.RutaTemporalWord || String(row.EstadoFase1 || '').trim().toUpperCase() !== 'OK') {
        skipped++;
        reportRows.push({
          EstadoFase2: 'SKIP',
          SolicitudOrigenID: row.SolicitudOrigenID || '',
          SolicitudID: row.SolicitudID || '',
          NombreDocumento: row.NombreDocumento || '',
          NombreArchivo: row.NombreArchivo || '',
          CodigoDocumento: row.CodigoDocumento || '',
          ArchivoProcesoOriginal: '',
          RutaProcesoOriginal: '',
          ArchivoProcesoRenombrado: '',
          RutaProcesoRenombrada: '',
          RutaHistorico: '',
          RutaNuevoPublicado: '',
          VersionDocumentoAnterior: '',
          VersionDocumentoNueva: row.VersionDocumento || '',
          FechaBajaHistorico: todayText,
          FechaAprobacionBaja: row.FechaDeAprobacion || '',
          Error: 'Fila omitida por no estar en estado OK de Fase 3 o no tener RutaTemporalWord.'
        });
        continue;
      }

      const currentFile = await getCurrentProcessFileBySolicitudId(params.context, webUrl, row.SolicitudOrigenID);
      if (!currentFile?.FileRef) {
        throw new Error(`No se encontró el archivo actual en Procesos para la solicitud padre ${row.SolicitudOrigenID}.`);
      }

      oldOriginalUrl = currentFile.FileRef;
      oldMetadata = await getFileItemMetadata(params.context, webUrl, oldOriginalUrl);
      const relativeFolder = getRelativeFolderWithinProcesos(oldOriginalUrl);
      const procesosFolder = joinFolder(PROCESOS_ROOT, relativeFolder);
      const historicosFolder = joinFolder(HISTORICOS_ROOT, relativeFolder);
      const oldFileName = currentFile.FileLeafRef;
      const oldVersion = sanitizeVersion(oldMetadata?.VersionDocumento || '');
      const renamedFileName = buildRenamedHistoricalFileName(oldFileName, oldVersion, todayStamp);
      oldRenamedUrl = `${procesosFolder}/${renamedFileName}`;
      historicoUrl = `${historicosFolder}/${renamedFileName}`;

      const outputFileName = /\.docx$/i.test(row.RutaTemporalWord) ? replaceExtension(oldFileName, '.pdf') : oldFileName;
      newPublishedUrl = `${procesosFolder}/${outputFileName}`;

      log(`📁 Fase 4 | Padre original: ${oldOriginalUrl}`);
      log(`✏️ Fase 4 | Padre renombrado: ${oldRenamedUrl}`);
      log(`📚 Fase 4 | Histórico destino: ${historicoUrl}`);
      log(`🚀 Fase 4 | Nuevo destino: ${newPublishedUrl}`);
      log(`📂 Fase 4 | Fuente TEMP padre: ${row.RutaTemporalWord}`);

      await moveFileByPath(params.context, webUrl, oldOriginalUrl, oldRenamedUrl, false);
      renamedOld = true;
      log(`✏️ Archivo padre renombrado | ${oldRenamedUrl}`);

      await ensureFolderPath(params.context, webUrl, historicosFolder);
      log(`📁 Carpeta histórico asegurada | ${historicosFolder}`);

      await copyFileByPath(params.context, webUrl, oldRenamedUrl, historicoUrl, false);
      historicoCopied = true;
      log(`📚 Archivo histórico padre copiado | ${historicoUrl}`);

      await updateHistoricoMetadata({
        context: params.context,
        webUrl,
        historicoFileUrl: historicoUrl,
        oldMetadata,
        today: now
      });
      log(`🧾 Metadata histórico padre aplicada | ${historicoUrl}`);

      await publishNewFile({
        context: params.context,
        webUrl,
        sourceFileUrl: row.RutaTemporalWord,
        targetFolderUrl: procesosFolder,
        outputFileName,
        log
      });
      newPublished = true;

      await updateParentProcesosMetadataAfterPublish({
        context: params.context,
        webUrl,
        targetFileUrl: newPublishedUrl,
        row,
        areaImpactadaIsMulti: fieldFlags.areaImpactadaIsMulti
      });
      log(`🧾 Metadata nuevo padre aplicada | ${newPublishedUrl}`);

      for (let j = 0; j < childIds.length; j++) {
        await updateExistingChildProcessParentReferences({
          context: params.context,
          webUrl,
          childSolicitudId: childIds[j],
          oldParentSolicitudId: row.DocumentoPadreSolicitudAnteriorID || row.SolicitudOrigenID,
          newParentSolicitudId: row.DocumentoPadreSolicitudNuevaID || row.SolicitudID,
          documentoPadreIsMulti: fieldFlags.documentoPadreIsMulti,
          log
        });
      }

      if (oldRenamedUrl) {
        await recycleFile(params.context, webUrl, oldRenamedUrl);
        log(`🗑️ Archivo padre viejo renombrado eliminado de Procesos | ${oldRenamedUrl}`);
      }

      rollbackEntries.push(buildRollbackEntry(row.SolicitudOrigenID, oldOriginalUrl, oldRenamedUrl, historicoUrl, newPublishedUrl, oldMetadata));
      ok++;
      reportRows.push({
        EstadoFase2: 'OK',
        SolicitudOrigenID: row.SolicitudOrigenID,
        SolicitudID: row.SolicitudID,
        NombreDocumento: row.NombreDocumento || oldMetadata?.NombreDocumento || '',
        NombreArchivo: row.NombreArchivo || '',
        CodigoDocumento: row.CodigoDocumento || oldMetadata?.Codigodedocumento || '',
        ArchivoProcesoOriginal: oldOriginalUrl.split('/').pop() || '',
        RutaProcesoOriginal: oldOriginalUrl,
        ArchivoProcesoRenombrado: oldRenamedUrl.split('/').pop() || '',
        RutaProcesoRenombrada: oldRenamedUrl,
        RutaHistorico: historicoUrl,
        RutaNuevoPublicado: newPublishedUrl,
        VersionDocumentoAnterior: oldMetadata?.VersionDocumento || '',
        VersionDocumentoNueva: row.VersionDocumento || '',
        FechaBajaHistorico: todayText,
        FechaAprobacionBaja: row.FechaDeAprobacion || '',
        Error: ''
      });
    } catch (e: any) {
      const message = e?.message || String(e);

      if (newPublished && newPublishedUrl) {
        try {
          await recycleFile(params.context, webUrl, newPublishedUrl);
          log(`↩️ Rollback local Fase 4 | Nuevo padre eliminado: ${newPublishedUrl}`);
        } catch (_error) {}
      }

      if (historicoCopied && historicoUrl) {
        try {
          await recycleFile(params.context, webUrl, historicoUrl);
          log(`↩️ Rollback local Fase 4 | Histórico padre eliminado: ${historicoUrl}`);
        } catch (_error) {}
      }

      if (renamedOld && oldRenamedUrl && oldOriginalUrl) {
        try {
          await moveFileByPath(params.context, webUrl, oldRenamedUrl, oldOriginalUrl, true);
          log(`↩️ Rollback local Fase 4 | Padre viejo restaurado: ${oldOriginalUrl}`);
        } catch (_error) {}
      }

      error++;
      reportRows.push({
        EstadoFase2: 'ERROR',
        SolicitudOrigenID: row.SolicitudOrigenID,
        SolicitudID: row.SolicitudID,
        NombreDocumento: row.NombreDocumento || oldMetadata?.NombreDocumento || '',
        NombreArchivo: '',
        CodigoDocumento: row.CodigoDocumento || oldMetadata?.Codigodedocumento || '',
        ArchivoProcesoOriginal: oldOriginalUrl ? oldOriginalUrl.split('/').pop() || '' : '',
        RutaProcesoOriginal: oldOriginalUrl,
        ArchivoProcesoRenombrado: oldRenamedUrl ? oldRenamedUrl.split('/').pop() || '' : '',
        RutaProcesoRenombrada: oldRenamedUrl,
        RutaHistorico: historicoUrl,
        RutaNuevoPublicado: newPublishedUrl,
        VersionDocumentoAnterior: oldMetadata?.VersionDocumento || '',
        VersionDocumentoNueva: row.VersionDocumento || '',
        FechaBajaHistorico: todayText,
        FechaAprobacionBaja: row.FechaDeAprobacion || '',
        Error: message
      });
      log(`❌ Error Fase 4 | SolicitudPadre=${row.SolicitudOrigenID} | ${message}`);
    }
  }

  return {
    rollbackEntries,
    reportRows,
    processed: parentRows.length,
    ok,
    skipped,
    error
  };
}
