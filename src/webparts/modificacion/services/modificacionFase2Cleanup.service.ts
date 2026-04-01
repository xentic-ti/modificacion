/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { escapeODataValue, recycleFile, spGetJson } from './sharepointRest.service';

type LogFn = (s: string) => void;

const PROCESOS_ROOT = '/sites/SistemadeGestionDocumental/Procesos';
const TEMP_WORD_ROOT = '/sites/SistemadeGestionDocumental/Procesos/TEMP_MIGRACION_WORD';

type IBridgeRow = {
  SolicitudOrigenID: number;
  NombreDocumento: string;
  NombreArchivo: string;
  RutaTemporalWord: string;
  EstadoFase1: string;
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

function replaceExtension(name: string, extensionWithDot: string): string {
  const baseName = String(name || '').replace(/\.[^.]+$/, '');
  return `${baseName}${extensionWithDot}`;
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

function buildTodayStamp(now: Date): string {
  const pad = (value: number): string => String(value).padStart(2, '0');
  return `${pad(now.getDate())}${pad(now.getMonth() + 1)}${now.getFullYear()}`;
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
    const rutaTemporalWord = String(getValue(row, 'RutaTemporalWord') || '').trim();
    const nombreDocumento = String(getValue(row, 'NombreDocumento') || '').trim();

    if (!nombreDocumento && !rutaTemporalWord) {
      continue;
    }

    rows.push({
      SolicitudOrigenID: Number(getValue(row, 'SolicitudOrigenID') || 0),
      NombreDocumento: nombreDocumento,
      NombreArchivo: String(getValue(row, 'NombreArchivo') || '').trim(),
      RutaTemporalWord: rutaTemporalWord,
      EstadoFase1: String(getValue(row, 'EstadoFase1') || '').trim()
    });
  }

  return rows;
}

async function listFolderFiles(context: WebPartContext, webUrl: string, folderUrl: string): Promise<Array<{ Name: string; ServerRelativeUrl: string }>> {
  const url =
    `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${escapeODataValue(folderUrl)}')/Files` +
    `?$select=Name,ServerRelativeUrl&$top=5000`;

  const response = await spGetJson<{ value?: Array<{ Name: string; ServerRelativeUrl: string }> }>(context, url);
  return response.value || [];
}

function escapeRegex(value: string): string {
  return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export async function cleanupArchivosViejosRenombradosFase2(params: {
  context: WebPartContext;
  excelFile: IFilePickerResult;
  log?: LogFn;
}): Promise<{ deleted: number; skipped: number; }> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const rows = await readBridgeExcel(params.excelFile);
  const todayStamp = buildTodayStamp(new Date());
  const processedFolders = new Set<string>();
  let deleted = 0;
  let skipped = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const tempFileUrl = String(row.RutaTemporalWord || '').trim();
    const tempFileName = String(row.NombreArchivo || '').trim();

    if (!tempFileUrl || String(row.EstadoFase1 || '').trim().toUpperCase() !== 'OK') {
      skipped++;
      continue;
    }

    const relativeFolder = getRelativeFolderBetween(tempFileUrl, TEMP_WORD_ROOT);
    const procesosFolder = joinFolder(PROCESOS_ROOT, relativeFolder);
    const oldProcessFileName = getOldProcessFileNameFromTemp(tempFileName);
    const baseName = String(oldProcessFileName || '').replace(/\.[^.]+$/, '');
    const extension = (String(oldProcessFileName || '').match(/\.[^.]+$/) || [''])[0];
    const cacheKey = `${procesosFolder}|${baseName}|${extension}`;

    if (processedFolders.has(cacheKey)) {
      continue;
    }
    processedFolders.add(cacheKey);

    let files: Array<{ Name: string; ServerRelativeUrl: string }> = [];
    try {
      files = await listFolderFiles(params.context, webUrl, procesosFolder);
    } catch (error: any) {
      log(`⚠️ Cleanup Fase 2 | No se pudo listar carpeta ${procesosFolder} -> ${error instanceof Error ? error.message : String(error)}`);
      skipped++;
      continue;
    }

    const pattern = new RegExp(`^${escapeRegex(baseName)}_V[^/]+_${todayStamp}${escapeRegex(extension)}$`, 'i');
    const candidates = files.filter((file) => pattern.test(String(file.Name || '')));

    if (!candidates.length) {
      log(`ℹ️ Cleanup Fase 2 | No se encontró viejo renombrado para "${row.NombreDocumento}" en ${procesosFolder}`);
      skipped++;
      continue;
    }

    for (let j = 0; j < candidates.length; j++) {
      const candidate = candidates[j];
      await recycleFile(params.context, webUrl, candidate.ServerRelativeUrl);
      deleted++;
      log(`🗑️ Cleanup Fase 2 | Viejo renombrado eliminado: ${candidate.ServerRelativeUrl}`);
    }
  }

  return { deleted, skipped };
}
