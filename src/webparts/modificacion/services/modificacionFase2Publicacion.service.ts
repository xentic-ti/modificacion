/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';
import { AadHttpClient } from '@microsoft/sp-http';
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ensureFolderPath, escapeODataValue, recycleFile, spGetJson, spPostJson, uploadFileToFolder } from './sharepointRest.service';
import { IFase2PublicacionReportRow, descargarReporteFase2Publicacion } from '../utils/fase2PublicacionReportExcel';

type LogFn = (s: string) => void;

const PROCESOS_ROOT = '/sites/SistemadeGestionDocumental/Procesos';
const TEMP_WORD_ROOT = '/sites/SistemadeGestionDocumental/Procesos/TEMP_MIGRACION_WORD';
const HISTORICOS_ROOT = '/sites/SistemadeGestionDocumental/Documentos Histricos';

export interface IFase2RollbackEntry {
  solicitudOrigenId: number;
  solicitudId: number;
  nombreDocumento: string;
  oldOriginalUrl: string;
  oldRenamedUrl: string;
  historicoUrl: string;
  newPublishedUrl: string;
  oldOriginalMetadata?: any;
}

type IBridgeRow = {
  SolicitudOrigenID: number;
  SolicitudID: number;
  NombreDocumento: string;
  NombreArchivo: string;
  CodigoDocumento: string;
  VersionDocumento: string;
  RutaTemporalWord: string;
  EstadoFase1: string;
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
    .map((token) => {
      if (token.length > 4 && token.endsWith('s')) {
        return token.slice(0, -1);
      }

      return token;
    });

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

function sanitizeVersion(value: any): string {
  return String(value || '').trim().replace(/^v/i, '') || '1.0';
}

function parseDdMmYyyyToDate(value: any): Date | null {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === 'number' && !isNaN(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) {
      return null;
    }

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
  if (!value || isNaN(value.getTime())) {
    return null;
  }

  return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0)).toISOString();
}

function parseAreaImpactada(value: any): string[] {
  return String(value || '')
    .split('/')
    .map((part) => part.trim())
    .filter(Boolean);
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
  if (!file) {
    throw new Error('Archivo Excel no recibido.');
  }

  if (typeof file.downloadFileContent === 'function') {
    const blob = await file.downloadFileContent();
    return blob.arrayBuffer();
  }

  const url = (file as any).fileAbsoluteUrl || '';
  if (!url) {
    throw new Error('No se pudo obtener el contenido del Excel de Fase 1.');
  }

  const response = await fetch(url, { credentials: 'same-origin' });
  if (!response.ok) {
    throw new Error(`No se pudo descargar el Excel de Fase 1. HTTP ${response.status}`);
  }

  return response.arrayBuffer();
}

async function readBridgeExcel(file: IFilePickerResult): Promise<IBridgeRow[]> {
  const buffer = await readArrayBufferFromFilePicker(file);
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error('No se encontró la hoja del Excel de Fase 1.');
  }

  const worksheet = workbook.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false }) as any[][];
  if (!aoa.length) {
    return [];
  }

  const headers = aoa[0] || [];
  const headerMap = new Map<string, number>();
  for (let i = 0; i < headers.length; i++) {
    headerMap.set(normalizeHeader(headers[i]), i);
  }

  const getValue = (row: any[], header: string): any => {
    const index = headerMap.get(normalizeHeader(header));
    return index === undefined ? '' : row[index];
  };

  const rows: IBridgeRow[] = [];

  for (let i = 1; i < aoa.length; i++) {
    const row = aoa[i] || [];
    const estadoFase1 = String(getValue(row, 'EstadoFase1') || '').trim();
    const rutaTemporalWord = String(getValue(row, 'RutaTemporalWord') || '').trim();
    const nombreDocumento = String(getValue(row, 'NombreDocumento') || '').trim();

    if (!nombreDocumento && !rutaTemporalWord) {
      continue;
    }

    rows.push({
      SolicitudOrigenID: Number(getValue(row, 'SolicitudOrigenID') || 0),
      SolicitudID: Number(getValue(row, 'SolicitudID') || 0),
      NombreDocumento: nombreDocumento,
      NombreArchivo: String(getValue(row, 'NombreArchivo') || '').trim(),
      CodigoDocumento: String(getValue(row, 'CodigoDocumento') || '').trim(),
      VersionDocumento: String(getValue(row, 'VersionDocumento') || '').trim(),
      RutaTemporalWord: rutaTemporalWord,
      EstadoFase1: estadoFase1,
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

function getRelativeFolderBetween(fullFileUrl: string, rootFolder: string): string {
  const full = trimSlash(fullFileUrl);
  const root = trimSlash(rootFolder);
  const fileDir = full.substring(0, full.lastIndexOf('/'));

  if (fileDir.indexOf(root) !== 0) {
    return '';
  }

  return fileDir.substring(root.length).replace(/^\/+/, '');
}

function getOldProcessFileNameFromTemp(tempFileName: string): string {
  return /\.docx$/i.test(tempFileName) ? replaceExtension(tempFileName, '.pdf') : tempFileName;
}

function getNewPublishedFileNameFromTemp(tempFileName: string): string {
  return /\.docx$/i.test(tempFileName) ? replaceExtension(tempFileName, '.pdf') : tempFileName;
}

function buildRenamedHistoricalFileName(originalName: string, oldVersion: string, todayStamp: string): string {
  const baseName = String(originalName || '').replace(/\.[^.]+$/, '');
  const extension = (String(originalName || '').match(/\.[^.]+$/) || [''])[0];
  return `${baseName}_V${sanitizeVersion(oldVersion)}_${todayStamp}${extension}`;
}

async function getFileItemMetadata(context: WebPartContext, webUrl: string, fileUrl: string): Promise<any> {
  const url =
    `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapeODataValue(fileUrl)}')/ListItemAllFields` +
    `?$select=Id,Title,FileLeafRef,NombreDocumento,Tipodedocumento,CategoriaDocumento,Codigodedocumento,AreaDuena,AreaImpactada,` +
    `SolicitudId,Clasificaciondeproceso,Macroproceso,Proceso,Subproceso,Resumen,FechaDeAprobacion,FechaDeVigencia,` +
    `InstanciaDeAprobacionId,VersionDocumento,Accion,Aprobadores,Descripcion,DocumentoPadreId`;

  return spGetJson<any>(context, url);
}

async function fileExists(context: WebPartContext, webUrl: string, fileUrl: string): Promise<boolean> {
  const url =
    `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapeODataValue(fileUrl)}')?$select=Name,ServerRelativeUrl`;

  try {
    await spGetJson<any>(context, url);
    return true;
  } catch (_error) {
    return false;
  }
}

async function listFolderFiles(context: WebPartContext, webUrl: string, folderUrl: string): Promise<Array<{ Name: string; ServerRelativeUrl: string; }>> {
  const url =
    `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${escapeODataValue(folderUrl)}')/Files` +
    `?$select=Name,ServerRelativeUrl&$top=5000`;

  const response = await spGetJson<{ value?: Array<{ Name: string; ServerRelativeUrl: string; }>; }>(context, url);
  return response.value || [];
}

async function resolveExistingProcessFileUrl(
  context: WebPartContext,
  webUrl: string,
  folderUrl: string,
  expectedFileName: string,
  log?: LogFn
): Promise<string | null> {
  const exactUrl = `${folderUrl}/${expectedFileName}`;
  if (await fileExists(context, webUrl, exactUrl)) {
    return exactUrl;
  }

  const files = await listFolderFiles(context, webUrl, folderUrl);
  const expectedKey = normalizeLooseFileKey(expectedFileName);
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    if (normalizeLooseFileKey(file.Name) === expectedKey) {
      log?.(`🔎 Fase 2 | Coincidencia flexible archivo viejo | Esperado="${expectedFileName}" | Encontrado="${file.Name}"`);
      return file.ServerRelativeUrl;
    }
  }

  return null;
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

async function ensureChoiceOptionByListPath(
  context: WebPartContext,
  webUrl: string,
  listPath: string,
  fieldInternalName: string,
  valueToEnsure: string
): Promise<void> {
  const normalized = String(valueToEnsure || '').trim();
  if (!normalized) {
    return;
  }

  const field = await getFieldInfoByListPath(context, webUrl, listPath, fieldInternalName);
  const choices = Array.isArray(field?.Choices) ? field.Choices : [];
  const exists = choices.some((choice: string) => String(choice || '').trim().toLowerCase() === normalized.toLowerCase());
  if (exists) {
    return;
  }

  await spPostJson(
    context,
    webUrl,
    `${webUrl}/_api/web/GetList('${escapeODataValue(listPath)}')/fields/getbyinternalnameortitle('${escapeODataValue(fieldInternalName)}')`,
    {
      Choices: [...choices, normalized]
    },
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

async function fileExistsByServerRelativeUrl(
  context: WebPartContext,
  webUrl: string,
  fileUrl: string
): Promise<boolean> {
  try {
    await spGetJson(
      context,
      webUrl,
      `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapeODataValue(fileUrl)}')?$select=Exists`
    );
    return true;
  } catch (error: any) {
    const message = String(error?.message || '');
    if (
      message.indexOf('(404)') !== -1 ||
      message.toLowerCase().indexOf('not found') !== -1 ||
      message.toLowerCase().indexOf('no se encuentra') !== -1
    ) {
      return false;
    }

    throw error;
  }
}

async function getFieldTypeFlagsForProcesos(context: WebPartContext, webUrl: string): Promise<{
  areaImpactadaIsMulti: boolean;
  categoriaDocumentoIsMulti: boolean;
}> {
  const areaField = await getFieldInfoByListPath(context, webUrl, PROCESOS_ROOT, 'AreaImpactada');
  const categoriaField = await getFieldInfoByListPath(context, webUrl, PROCESOS_ROOT, 'CategoriaDocumento');

  return {
    areaImpactadaIsMulti: String(areaField?.TypeAsString || '').toLowerCase().indexOf('multi') !== -1,
    categoriaDocumentoIsMulti: String(categoriaField?.TypeAsString || '').toLowerCase().indexOf('multi') !== -1
  };
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
    const response = await client.get(
      `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem/content?format=pdf`,
      AadHttpClient.configurations.v1
    );

    if (response.ok) {
      const pdfBuffer = await response.arrayBuffer();
      const pdfBlob = new Blob([pdfBuffer], { type: 'application/pdf' });
      await uploadFileToFolder(
        params.context,
        params.webUrl,
        params.destinoFolderServerRelativeUrl,
        params.outputPdfName,
        pdfBlob
      );
      log(`📄✅ Nuevo documento publicado en PDF | ${params.outputPdfName}`);
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
      log(`⚠️ Reintentando conversión PDF (${attempt}/${maxRetries}) | ${params.outputPdfName} | HTTP ${response.status}`);
      await new Promise((resolve) => setTimeout(resolve, attempt * 2000));
      continue;
    }

    throw new Error(`Graph PDF failed (${response.status}): ${body}`);
  }
}

async function publishNewFile(params: {
  context: WebPartContext;
  webUrl: string;
  tempFileUrl: string;
  targetFolderUrl: string;
  outputFileName: string;
  log?: LogFn;
}): Promise<string> {
  const destinationFileUrl = `${trimSlash(params.targetFolderUrl)}/${params.outputFileName}`;
  if (/\.docx$/i.test(params.tempFileUrl)) {
    const siblingPdfUrl = replaceExtension(params.tempFileUrl, '.pdf');
    const siblingPdfExists = await fileExistsByServerRelativeUrl(params.context, params.webUrl, siblingPdfUrl);

    if (siblingPdfExists) {
      try {
        await copyFileByPath(params.context, params.webUrl, siblingPdfUrl, destinationFileUrl, false);
        params.log?.(`📄✅ Nuevo documento publicado usando PDF existente | ${params.outputFileName}`);
      } catch (error: any) {
        const message = error instanceof Error ? error.message : String(error);
        const isNotFound = message.indexOf('(404)') !== -1 || message.toLowerCase().indexOf('archivo no encontrado') !== -1;
        if (!isNotFound) {
          throw new Error(`No se pudo usar el archivo temporal PDF existente. Fuente="${siblingPdfUrl}" | Destino="${destinationFileUrl}" | ${message}`);
        }

        params.log?.(`⚠️ PDF hermano no disponible al copiar, se intentará convertir el Word | ${params.outputFileName}`);
        try {
          await convertOfficeFileToPdfAndUpload({
            context: params.context,
            webUrl: params.webUrl,
            sourceServerRelativeUrl: params.tempFileUrl,
            destinoFolderServerRelativeUrl: params.targetFolderUrl,
            outputPdfName: params.outputFileName,
            log: params.log
          });
        } catch (conversionError: any) {
          const conversionMessage = conversionError instanceof Error ? conversionError.message : String(conversionError);
          if (conversionMessage.indexOf('(404)') !== -1 || conversionMessage.toLowerCase().indexOf('archivo no encontrado') !== -1) {
            throw new Error(`No se encontró el archivo temporal Word para convertir. Fuente="${params.tempFileUrl}" | Destino="${destinationFileUrl}" | ${conversionMessage}`);
          }

          throw conversionError;
        }
      }
    } else {
      try {
        await convertOfficeFileToPdfAndUpload({
          context: params.context,
          webUrl: params.webUrl,
          sourceServerRelativeUrl: params.tempFileUrl,
          destinoFolderServerRelativeUrl: params.targetFolderUrl,
          outputPdfName: params.outputFileName,
          log: params.log
        });
      } catch (error: any) {
        const message = error instanceof Error ? error.message : String(error);
        if (message.indexOf('(404)') !== -1 || message.toLowerCase().indexOf('archivo no encontrado') !== -1) {
          throw new Error(`No se encontró el archivo temporal Word para convertir. Fuente="${params.tempFileUrl}" | Destino="${destinationFileUrl}" | ${message}`);
        }

        throw error;
      }
    }
  } else {
    try {
      await copyFileByPath(params.context, params.webUrl, params.tempFileUrl, destinationFileUrl, false);
      params.log?.(`📄✅ Nuevo documento publicado sin conversión | ${params.outputFileName}`);
    } catch (error: any) {
      const message = error instanceof Error ? error.message : String(error);
      throw new Error(`No se encontró el archivo temporal para publicar. Fuente="${params.tempFileUrl}" | Destino="${destinationFileUrl}" | ${message}`);
    }
  }

  return destinationFileUrl;
}

async function updateProcesosMetadataAfterPublish(params: {
  context: WebPartContext;
  webUrl: string;
  targetFileUrl: string;
  row: IBridgeRow;
  areaImpactadaIsMulti: boolean;
  categoriaDocumentoIsMulti: boolean;
}): Promise<void> {
  const fechaAprobacion = toDateOnlyIso(parseDdMmYyyyToDate(params.row.FechaDeAprobacion));
  const fechaVigencia = toDateOnlyIso(parseDdMmYyyyToDate(params.row.FechaDeVigencia));
  const areaImpactada = parseAreaImpactada(params.row.AreaImpactada);
  const categoriaDocumento = String(params.row.CategoriaDocumento || '').trim();

  const payload: any = {
    Clasificaciondeproceso: params.row.Clasificaciondeproceso || '',
    AreaDuena: params.row.AreaDuena || '',
    VersionDocumento: params.row.VersionDocumento || '',
    AreaImpactada: params.areaImpactadaIsMulti
      ? areaImpactada
      : (areaImpactada[0] || ''),
    Macroproceso: params.row.Macroproceso || '',
    Proceso: params.row.Proceso || '',
    Subproceso: params.row.Subproceso || '',
    Tipodedocumento: params.row.TipoDocumento || '',
    SolicitudId: Number(params.row.SolicitudID || 0) || null,
    Codigodedocumento: params.row.CodigoDocumento || '',
    Resumen: params.row.Resumen || '',
    CategoriaDocumento: categoriaDocumento,
    FechaDeAprobacion: fechaAprobacion,
    FechaDePublicacion: new Date().toISOString(),
    FechaDeVigencia: fechaVigencia,
    InstanciaDeAprobacionId: Number(params.row.InstanciaDeAprobacionId || 0) || null,
    Accion: 'Actualización de documento',
    NombreDocumento: params.row.NombreDocumento || ''
  };

  await updateFileMetadataByPath(params.context, params.webUrl, params.targetFileUrl, payload);
}

async function updateHistoricoMetadata(params: {
  context: WebPartContext;
  webUrl: string;
  historicoFileUrl: string;
  oldMetadata: any;
  row: IBridgeRow;
  today: Date;
}): Promise<void> {
  const hisAreaImpactada = parseAreaImpactada(params.oldMetadata?.AreaImpactada).join(' / ');
  if (hisAreaImpactada) {
    await ensureChoiceOptionByListPath(
      params.context,
      params.webUrl,
      HISTORICOS_ROOT,
      'HisAreaImpactada',
      hisAreaImpactada
    );
  }

  const payload: any = {
    HisAreaDuena: params.oldMetadata?.AreaDuena || '',
    HisAreaImpactada: hisAreaImpactada,
    HisClasificaciondeproceso: params.oldMetadata?.Clasificaciondeproceso || params.row.Clasificaciondeproceso || '',
    HisMacroproceso: params.oldMetadata?.Macroproceso || params.row.Macroproceso || '',
    HisProceso: params.oldMetadata?.Proceso || params.row.Proceso || '',
    HisSubproceso: params.oldMetadata?.Subproceso || params.row.Subproceso || '',
    HisTipodedocumento: params.oldMetadata?.Tipodedocumento || params.row.TipoDocumento || '',
    HisCodigodedocumento: params.oldMetadata?.Codigodedocumento || params.row.CodigoDocumento || '',
    HisResumen: params.oldMetadata?.Resumen || params.row.Resumen || '',
    HisVersionDocumento: params.oldMetadata?.VersionDocumento || '',
    HisAprobadores: params.oldMetadata?.Aprobadores || '',
    HisFechaDeBaja: toDateOnlyIso(params.today),
    HisCategoriaDocumento: params.oldMetadata?.CategoriaDocumento || params.row.CategoriaDocumento || '',
    InstanciaDeAprobacionId: Number(params.oldMetadata?.InstanciaDeAprobacionId || 0) || null,
    Accion: 'Actualización de documento',
    HisFechaAprobacionBaja: toDateOnlyIso(parseDdMmYyyyToDate(params.row.FechaDeAprobacion))
  };

  await updateFileMetadataByPath(params.context, params.webUrl, params.historicoFileUrl, payload);
}

function buildRollbackEntry(
  row: IBridgeRow,
  oldOriginalUrl: string,
  oldRenamedUrl: string,
  historicoUrl: string,
  newPublishedUrl: string,
  oldOriginalMetadata: any
): IFase2RollbackEntry {
  return {
    solicitudOrigenId: Number(row.SolicitudOrigenID || 0),
    solicitudId: Number(row.SolicitudID || 0),
    nombreDocumento: row.NombreDocumento || '',
    oldOriginalUrl,
    oldRenamedUrl,
    historicoUrl,
    newPublishedUrl,
    oldOriginalMetadata
  };
}

export async function ejecutarFase2PublicacionDocumentosSinHijos(params: {
  context: WebPartContext;
  excelFile: IFilePickerResult;
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
  const rows = await readBridgeExcel(params.excelFile);
  const reportRows: IFase2PublicacionReportRow[] = [];
  const rollbackEntries: IFase2RollbackEntry[] = [];
  const fieldFlags = await getFieldTypeFlagsForProcesos(params.context, webUrl);
  const now = new Date();
  const todayStamp = buildTodayStamp(now);
  const todayText = buildTodayDdMmYyyy(now);

  let ok = 0;
  let skipped = 0;
  let error = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const tempFileUrl = String(row.RutaTemporalWord || '').trim();
    const tempFileName = String(row.NombreArchivo || '').trim();

    if (!tempFileUrl || String(row.EstadoFase1 || '').trim().toUpperCase() !== 'OK') {
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
        Error: 'Fila omitida por no estar en estado OK de Fase 1 o no tener RutaTemporalWord.'
      });
      continue;
    }

    const relativeFolder = getRelativeFolderBetween(tempFileUrl, TEMP_WORD_ROOT);
    const procesosFolder = joinFolder(PROCESOS_ROOT, relativeFolder);
    const historicosFolder = joinFolder(HISTORICOS_ROOT, relativeFolder);
    const oldProcessFileName = getOldProcessFileNameFromTemp(tempFileName);
    const newPublishedFileName = getNewPublishedFileNameFromTemp(tempFileName);
    const oldOriginalUrlExpected = `${procesosFolder}/${oldProcessFileName}`;

    let oldMetadata: any = null;
    let oldOriginalUrl = '';
    let oldRenamedUrl = '';
    let historicoUrl = '';
    let newPublishedUrl = '';
    let renamedOld = false;
    let historicoCopied = false;
    let newPublished = false;

    try {
      oldOriginalUrl = await resolveExistingProcessFileUrl(
        params.context,
        webUrl,
        procesosFolder,
        oldProcessFileName,
        log
      ) || '';

      if (!oldOriginalUrl) {
        throw new Error(`No se encontró el archivo viejo en Procesos: ${oldOriginalUrlExpected}`);
      }

      oldMetadata = await getFileItemMetadata(params.context, webUrl, oldOriginalUrl);
      const oldVersion = sanitizeVersion(oldMetadata?.VersionDocumento || '');
      const renamedFileName = buildRenamedHistoricalFileName(oldProcessFileName, oldVersion, todayStamp);
      oldRenamedUrl = `${procesosFolder}/${renamedFileName}`;
      historicoUrl = `${historicosFolder}/${renamedFileName}`;
      newPublishedUrl = `${procesosFolder}/${newPublishedFileName}`;

      log(`📁 Fase 2 | Viejo original: ${oldOriginalUrl}`);
      log(`✏️ Fase 2 | Viejo renombrado: ${oldRenamedUrl}`);
      log(`📚 Fase 2 | Histórico destino: ${historicoUrl}`);
      log(`🚀 Fase 2 | Nuevo destino: ${newPublishedUrl}`);

      await moveFileByPath(params.context, webUrl, oldOriginalUrl, oldRenamedUrl, false);
      renamedOld = true;
      log(`✏️ Archivo viejo renombrado | ${oldRenamedUrl}`);

      await ensureFolderPath(params.context, webUrl, historicosFolder);
      log(`📁 Carpeta histórico asegurada | ${historicosFolder}`);

      await copyFileByPath(params.context, webUrl, oldRenamedUrl, historicoUrl, false);
      historicoCopied = true;
      log(`📚 Archivo histórico copiado | ${historicoUrl}`);

      await updateHistoricoMetadata({
        context: params.context,
        webUrl,
        historicoFileUrl: historicoUrl,
        oldMetadata,
        row,
        today: now
      });
      log(`🧾 Metadata histórico aplicada | ${historicoUrl}`);

      await publishNewFile({
        context: params.context,
        webUrl,
        tempFileUrl,
        targetFolderUrl: procesosFolder,
        outputFileName: newPublishedFileName,
        log
      });
      newPublished = true;

      await updateProcesosMetadataAfterPublish({
        context: params.context,
        webUrl,
        targetFileUrl: newPublishedUrl,
        row,
        areaImpactadaIsMulti: fieldFlags.areaImpactadaIsMulti,
        categoriaDocumentoIsMulti: fieldFlags.categoriaDocumentoIsMulti
      });
      log(`🧾 Metadata nuevo documento aplicada | ${newPublishedUrl}`);

      if (oldRenamedUrl) {
        await recycleFile(params.context, webUrl, oldRenamedUrl);
        log(`🗑️ Archivo viejo renombrado eliminado de Procesos | ${oldRenamedUrl}`);
        renamedOld = false;
      }

      rollbackEntries.push(buildRollbackEntry(row, oldOriginalUrl, oldRenamedUrl, historicoUrl, newPublishedUrl, oldMetadata));

      ok++;
      reportRows.push({
        EstadoFase2: 'OK',
        SolicitudOrigenID: row.SolicitudOrigenID || '',
        SolicitudID: row.SolicitudID || '',
        NombreDocumento: row.NombreDocumento || '',
        NombreArchivo: row.NombreArchivo || '',
        CodigoDocumento: row.CodigoDocumento || '',
        ArchivoProcesoOriginal: oldProcessFileName,
        RutaProcesoOriginal: oldOriginalUrl,
        ArchivoProcesoRenombrado: renamedFileName,
        RutaProcesoRenombrada: oldRenamedUrl,
        RutaHistorico: historicoUrl,
        RutaNuevoPublicado: newPublishedUrl,
        VersionDocumentoAnterior: oldVersion,
        VersionDocumentoNueva: row.VersionDocumento || '',
        FechaBajaHistorico: todayText,
        FechaAprobacionBaja: row.FechaDeAprobacion || '',
        Error: ''
      });
    } catch (phase2Error) {
      const message = phase2Error instanceof Error ? phase2Error.message : String(phase2Error);

      if (newPublished && newPublishedUrl) {
        try {
          await recycleFile(params.context, webUrl, newPublishedUrl);
          log(`↩️ Rollback local Fase 2 | Nuevo eliminado: ${newPublishedUrl}`);
        } catch (_cleanupError) {
          // sin acción
        }
      }

      if (historicoCopied && historicoUrl) {
        try {
          await recycleFile(params.context, webUrl, historicoUrl);
          log(`↩️ Rollback local Fase 2 | Histórico eliminado: ${historicoUrl}`);
        } catch (_cleanupError) {
          // sin acción
        }
      }

      if (renamedOld && oldRenamedUrl && oldOriginalUrl) {
        try {
          await moveFileByPath(params.context, webUrl, oldRenamedUrl, oldOriginalUrl, true);
          log(`↩️ Rollback local Fase 2 | Viejo restaurado: ${oldOriginalUrl}`);
        } catch (_cleanupError) {
          // sin acción
        }
      }

      error++;
      reportRows.push({
        EstadoFase2: 'ERROR',
        SolicitudOrigenID: row.SolicitudOrigenID || '',
        SolicitudID: row.SolicitudID || '',
        NombreDocumento: row.NombreDocumento || '',
        NombreArchivo: row.NombreArchivo || '',
        CodigoDocumento: row.CodigoDocumento || '',
        ArchivoProcesoOriginal: oldProcessFileName,
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
      log(`❌ Error Fase 2 | SolicitudOrigen=${row.SolicitudOrigenID} | Documento="${row.NombreDocumento}" | ${message}`);
    }
  }

  descargarReporteFase2Publicacion(reportRows);

  return {
    rollbackEntries,
    reportRows,
    processed: rows.length,
    ok,
    skipped,
    error
  };
}
