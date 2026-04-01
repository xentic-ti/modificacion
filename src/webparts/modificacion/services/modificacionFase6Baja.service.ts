/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  deleteListItem,
  ensureFolderPath,
  escapeODataValue,
  recycleFile,
  spGetJson,
  spPostJson,
  updateListItem
} from './sharepointRest.service';
import { IFase2PublicacionReportRow, descargarReporteFase2Publicacion } from '../utils/fase2PublicacionReportExcel';

type LogFn = (s: string) => void;

const PROCESOS_ROOT = '/sites/SistemadeGestionDocumental/Procesos';
const HISTORICOS_ROOT = '/sites/SistemadeGestionDocumental/Documentos Histricos';

type IFase6ExcelRow = {
  solicitudId: number;
  nombreDocumento: string;
  fechaAprobacion: string;
  documentosHijosIds: number[];
  diagramasFlujoIds: number[];
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

function sanitizeVersion(value: any): string {
  return String(value || '').trim().replace(/^v/i, '') || '1.0';
}

function buildTodayStamp(now: Date): string {
  const pad = (value: number): string => String(value).padStart(2, '0');
  return `${pad(now.getDate())}${pad(now.getMonth() + 1)}${now.getFullYear()}`;
}

function buildTodayDdMmYyyy(now: Date): string {
  const pad = (value: number): string => String(value).padStart(2, '0');
  return `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()}`;
}

function parseSlashIds(value: any): number[] {
  return String(value || '')
    .split('/')
    .map((part) => Number(String(part || '').trim()))
    .filter((id) => Number.isFinite(id) && id > 0);
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
  if (Array.isArray(value)) {
    return value.map((item) => String(item || '').trim()).filter(Boolean);
  }

  return String(value || '')
    .split('/')
    .map((part) => part.trim())
    .filter(Boolean);
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
    `?$select=Choices`;

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
  if (!row) {
    return null;
  }

  return {
    Id: Number(row.Id || 0),
    FileRef: String(row.FileRef || ''),
    FileLeafRef: String(row.FileLeafRef || '')
  };
}

async function readFase6Excel(file: IFilePickerResult): Promise<IFase6ExcelRow[]> {
  const buffer = typeof file.downloadFileContent === 'function'
    ? await (await file.downloadFileContent()).arrayBuffer()
    : await (await fetch((file as any).fileAbsoluteUrl || '', { credentials: 'same-origin' })).arrayBuffer();

  const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error('No se encontró la hoja del Excel de revisión.');
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

  const findIndex = (aliases: string[]): number => {
    for (let i = 0; i < aliases.length; i++) {
      const index = headerMap.get(normalizeHeader(aliases[i]));
      if (index !== undefined) {
        return index;
      }
    }

    return -1;
  };

  const idxSolicitud = findIndex(['ID Solicitud', 'SolicitudID', 'Solicitud Id']);
  const idxNombre = findIndex(['Nombre Documento', 'NombreDocumento']);
  const idxFechaAprobacion = findIndex(['Fecha Aprobación', 'FechaDeAprobacion', 'Fecha de Aprobacion', 'Fecha Aprobacion']);
  const idxHijos = findIndex(['Documentos hijos', 'DocumentosHijosIDs', 'DocumentosHijos']);
  const idxFlujos = findIndex(['Diagrama de Flujos', 'Diagramas de Flujo', 'DiagramasFlujo']);

  if (idxSolicitud === -1) {
    throw new Error('El Excel no contiene la columna "ID Solicitud".');
  }

  const rows: IFase6ExcelRow[] = [];
  for (let i = 1; i < aoa.length; i++) {
    const row = aoa[i] || [];
    const solicitudId = Number(row[idxSolicitud] || 0);
    const nombreDocumento = idxNombre === -1 ? '' : String(row[idxNombre] || '').trim();
    const fechaAprobacion = idxFechaAprobacion === -1 ? '' : String(row[idxFechaAprobacion] || '').trim();
    const documentosHijosIds = idxHijos === -1 ? [] : parseSlashIds(row[idxHijos]);
    const diagramasFlujoIds = idxFlujos === -1 ? [] : parseSlashIds(row[idxFlujos]);

    if (!solicitudId && !nombreDocumento) {
      continue;
    }

    rows.push({
      solicitudId,
      nombreDocumento,
      fechaAprobacion,
      documentosHijosIds,
      diagramasFlujoIds
    });
  }

  return rows;
}

async function updateHistoricoMetadataForBaja(params: {
  context: WebPartContext;
  webUrl: string;
  historicoFileUrl: string;
  oldMetadata: any;
  fechaAprobacionExcel: string;
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
    Accion: 'Baja de documento',
    HisFechaAprobacionBaja: toDateOnlyIso(parseDdMmYyyyToDate(params.fechaAprobacionExcel))
  };

  await updateFileMetadataByPath(params.context, params.webUrl, params.historicoFileUrl, payload);
}

async function moverDocumentoAHistoricoPorSolicitud(params: {
  context: WebPartContext;
  webUrl: string;
  solicitudId: number;
  fechaAprobacionExcel: string;
  now: Date;
  todayStamp: string;
  todayText: string;
  log: LogFn;
}): Promise<{
  rutaProcesoOriginal: string;
  rutaProcesoRenombrada: string;
  rutaHistorico: string;
  nombreArchivo: string;
  versionDocumentoAnterior: string;
  codigoDocumento: string;
  nombreDocumento: string;
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

  params.log(`📁 Fase 6 | Documento original: ${oldOriginalUrl}`);
  params.log(`✏️ Fase 6 | Documento renombrado: ${oldRenamedUrl}`);
  params.log(`📚 Fase 6 | Histórico destino: ${historicoUrl}`);

  await moveFileByPath(params.context, params.webUrl, oldOriginalUrl, oldRenamedUrl, false);
  await ensureFolderPath(params.context, params.webUrl, historicosFolder);
  await copyFileByPath(params.context, params.webUrl, oldRenamedUrl, historicoUrl, false);

  await updateHistoricoMetadataForBaja({
    context: params.context,
    webUrl: params.webUrl,
    historicoFileUrl: historicoUrl,
    oldMetadata,
    fechaAprobacionExcel: params.fechaAprobacionExcel,
    today: params.now
  });

  await recycleFile(params.context, params.webUrl, oldRenamedUrl);
  await updateListItem(params.context, params.webUrl, 'Solicitudes', params.solicitudId, {
    EsVersionActualDocumento: false
  });

  return {
    rutaProcesoOriginal: oldOriginalUrl,
    rutaProcesoRenombrada: oldRenamedUrl,
    rutaHistorico: historicoUrl,
    nombreArchivo: oldFileName,
    versionDocumentoAnterior: oldMetadata?.VersionDocumento || '',
    codigoDocumento: oldMetadata?.Codigodedocumento || '',
    nombreDocumento: oldMetadata?.NombreDocumento || oldMetadata?.Title || ''
  };
}

export async function ejecutarFase6BajaDocumentos(params: {
  context: WebPartContext;
  excelFile: IFilePickerResult;
  log?: LogFn;
}): Promise<{
  reportRows: IFase2PublicacionReportRow[];
  processed: number;
  ok: number;
  skipped: number;
  error: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const rows = await readFase6Excel(params.excelFile);
  const reportRows: IFase2PublicacionReportRow[] = [];
  const now = new Date();
  const todayStamp = buildTodayStamp(now);
  const todayText = buildTodayDdMmYyyy(now);
  const processedSolicitudes = new Set<number>();
  const processedChildren = new Set<number>();

  let ok = 0;
  let skipped = 0;
  let error = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    if (!row.solicitudId) {
      skipped++;
      reportRows.push({
        EstadoFase2: 'SKIP',
        SolicitudOrigenID: '',
        SolicitudID: '',
        NombreDocumento: row.nombreDocumento || '',
        NombreArchivo: '',
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
        FechaAprobacionBaja: row.fechaAprobacion || '',
        Error: 'Fila omitida por no tener ID Solicitud.'
      });
      continue;
    }

    if (processedSolicitudes.has(row.solicitudId)) {
      skipped++;
      log(`ℹ️ Fase 6 | Solicitud repetida omitida: ${row.solicitudId}`);
      continue;
    }

    try {
      const principal = await moverDocumentoAHistoricoPorSolicitud({
        context: params.context,
        webUrl,
        solicitudId: row.solicitudId,
        fechaAprobacionExcel: row.fechaAprobacion,
        now,
        todayStamp,
        todayText,
        log
      });
      processedSolicitudes.add(row.solicitudId);

      for (let j = 0; j < row.documentosHijosIds.length; j++) {
        const childSolicitudId = row.documentosHijosIds[j];
        if (processedChildren.has(childSolicitudId) || processedSolicitudes.has(childSolicitudId)) {
          log(`ℹ️ Fase 6 | Hijo ya procesado, se omite: ${childSolicitudId}`);
          continue;
        }

        await moverDocumentoAHistoricoPorSolicitud({
          context: params.context,
          webUrl,
          solicitudId: childSolicitudId,
          fechaAprobacionExcel: row.fechaAprobacion,
          now,
          todayStamp,
          todayText,
          log
        });
        processedChildren.add(childSolicitudId);
        processedSolicitudes.add(childSolicitudId);
        log(`👶 Fase 6 | Hijo enviado a histórico: ${childSolicitudId}`);
      }

      for (let j = 0; j < row.diagramasFlujoIds.length; j++) {
        await deleteListItem(params.context, webUrl, 'Diagramas de Flujo', row.diagramasFlujoIds[j]);
      }
      if (row.diagramasFlujoIds.length) {
        log(`🧭 Fase 6 | Diagramas eliminados: ${row.diagramasFlujoIds.join('/')}`);
      }

      ok++;
      reportRows.push({
        EstadoFase2: 'OK',
        SolicitudOrigenID: row.solicitudId,
        SolicitudID: row.solicitudId,
        NombreDocumento: principal.nombreDocumento || row.nombreDocumento || '',
        NombreArchivo: principal.nombreArchivo || '',
        CodigoDocumento: principal.codigoDocumento || '',
        ArchivoProcesoOriginal: principal.nombreArchivo || '',
        RutaProcesoOriginal: principal.rutaProcesoOriginal || '',
        ArchivoProcesoRenombrado: principal.rutaProcesoRenombrada.split('/').pop() || '',
        RutaProcesoRenombrada: principal.rutaProcesoRenombrada || '',
        RutaHistorico: principal.rutaHistorico || '',
        RutaNuevoPublicado: '',
        VersionDocumentoAnterior: principal.versionDocumentoAnterior || '',
        VersionDocumentoNueva: '',
        FechaBajaHistorico: todayText,
        FechaAprobacionBaja: row.fechaAprobacion || '',
        Error: ''
      });
    } catch (fase6Error) {
      const message = fase6Error instanceof Error ? fase6Error.message : String(fase6Error);
      error++;
      reportRows.push({
        EstadoFase2: 'ERROR',
        SolicitudOrigenID: row.solicitudId || '',
        SolicitudID: row.solicitudId || '',
        NombreDocumento: row.nombreDocumento || '',
        NombreArchivo: '',
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
        FechaAprobacionBaja: row.fechaAprobacion || '',
        Error: message
      });
      log(`❌ Error Fase 6 | Solicitud=${row.solicitudId} | Documento="${row.nombreDocumento}" | ${message}`);
    }
  }

  descargarReporteFase2Publicacion(reportRows, `Reporte_Fase6_BAJA_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}${String(now.getSeconds()).padStart(2, '0')}.xlsx`);

  return {
    reportRows,
    processed: rows.length,
    ok,
    skipped,
    error
  };
}
