/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { openExcelRevisionSession } from '../utils/modificacionExcelHelper';
import { getAllItems, spGetJson, spPostJson, updateListItem, escapeODataValue } from './sharepointRest.service';

type LogFn = (message: string) => void;

const PROCESOS_ROOT = '/sites/SistemadeGestionDocumental/Procesos';

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

function findColumnIndex(headers: any[], candidates: string[]): number {
  const normalizedCandidates = candidates.map((item) => normKey(item));

  for (let i = 0; i < headers.length; i++) {
    const current = normKey(headers[i]);
    if (normalizedCandidates.indexOf(current) !== -1) {
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

  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists(guid'${field.LookupList}')/items?$select=Id,Title&$top=5000`
  );

  const map = new Map<string, number>();
  for (let i = 0; i < items.length; i++) {
    const title = String(items[i]?.Title || '').trim();
    const key = normKey(title);
    const compact = compactKey(title);
    if (key && !map.has(key)) map.set(key, Number(items[i].Id || 0));
    if (compact && !map.has(compact)) map.set(compact, Number(items[i].Id || 0));
  }

  return map;
}

function resolveLookupId(map: Map<string, number>, value: any): number | undefined {
  const key = normKey(value);
  if (key && map.has(key)) {
    return map.get(key);
  }

  const compact = compactKey(value);
  if (compact && map.has(compact)) {
    return map.get(compact);
  }

  return undefined;
}

async function buscarSolicitudVigentePorNombre(
  context: WebPartContext,
  webUrl: string,
  documentName: string
): Promise<any | null> {
  const escaped = String(documentName || '').replace(/'/g, `''`);
  const filter = `(Title eq '${escaped}' or NombreDocumento eq '${escaped}') and EsVersionActualDocumento eq 1`;
  const url =
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items` +
    `?$select=Id,Title,NombreDocumento,InstanciasdeaprobacionId,EsVersionActualDocumento` +
    `&$top=5&$filter=${encodeURIComponent(filter)}`;

  const items = await getAllItems<any>(context, url);
  return items.length ? items[0] : null;
}

async function getCurrentProcessFileBySolicitudId(
  context: WebPartContext,
  webUrl: string,
  solicitudId: number
): Promise<{ Id: number; FileRef: string; FileLeafRef: string; SolicitudId: number; } | null> {
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
    FileLeafRef: String(item.FileLeafRef || ''),
    SolicitudId: Number(item.SolicitudId || 0)
  };
}

async function updateFileMetadataByPath(
  context: WebPartContext,
  webUrl: string,
  fileUrl: string,
  payload: any
): Promise<void> {
  await spPostJson(
    context,
    webUrl,
    `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapeODataValue(fileUrl)}')/ListItemAllFields`,
    payload,
    'MERGE'
  );
}

function buildOutputFileName(originalFileName: string): string {
  const dotIndex = String(originalFileName || '').lastIndexOf('.');
  if (dotIndex === -1) {
    return `${originalFileName}_cambio_instancia.xlsx`;
  }

  return `${originalFileName.substring(0, dotIndex)}_cambio_instancia${originalFileName.substring(dotIndex)}`;
}

function cloneGrid(grid: any[][]): any[][] {
  const result: any[][] = [];
  for (let i = 0; i < grid.length; i++) {
    result.push(Array.isArray(grid[i]) ? [...grid[i]] : []);
  }
  return result;
}

export async function ejecutarCambioInstanciaDesdeExcel(params: {
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

  const headers = grid[0] || [];
  const idxNombre = findColumnIndex(headers, [
    'NombreDocumento',
    'Nombre Documento',
    'Nombre del Documento',
    'Nombre de la solicitud',
    'Solicitud',
    'Title',
    'Titulo'
  ]);
  const idxCambio = findColumnIndex(headers, [
    'Cambio',
    'Cambio Instancia'
  ]);

  if (idxNombre < 0 || idxCambio < 0) {
    throw new Error('El Excel debe contener las columnas de nombre de solicitud y Cambio.');
  }

  const mapInstancias = await buildLookupMapByField(params.context, webUrl, 'Solicitudes', 'Instanciasdeaprobacion');
  const outputGrid = cloneGrid(grid);
  const outputHeaders = outputGrid[0] || [];
  const idxEstado = outputHeaders.length;
  const idxDetalle = outputHeaders.length + 1;
  const idxSolicitudId = outputHeaders.length + 2;
  const idxProceso = outputHeaders.length + 3;

  outputHeaders[idxEstado] = 'ResultadoCambioInstancia';
  outputHeaders[idxDetalle] = 'DetalleCambioInstancia';
  outputHeaders[idxSolicitudId] = 'SolicitudIdCambioInstancia';
  outputHeaders[idxProceso] = 'ArchivoProcesoCambioInstancia';
  outputGrid[0] = outputHeaders;

  let processed = 0;
  let updated = 0;
  let skipped = 0;
  let error = 0;

  log(`📄 Excel cargado: ${session.fileName}`);
  log(`📋 Filas detectadas: ${Math.max(0, grid.length - 1)}`);

  for (let rowIndex = 1; rowIndex < grid.length; rowIndex++) {
    const row = grid[rowIndex] || [];
    const nombreSolicitud = String(row[idxNombre] || '').trim();
    const cambio = String(row[idxCambio] || '').trim();
    const outputRow = outputGrid[rowIndex] || [];

    if (!nombreSolicitud) {
      outputRow[idxEstado] = 'SKIP';
      outputRow[idxDetalle] = 'Fila sin nombre de solicitud.';
      outputGrid[rowIndex] = outputRow;
      skipped++;
      continue;
    }

    if (!cambio) {
      outputRow[idxEstado] = 'ERROR';
      outputRow[idxDetalle] = 'La columna Cambio está vacía.';
      outputGrid[rowIndex] = outputRow;
      error++;
      continue;
    }

    processed++;
    log(`🔎 Buscando solicitud vigente para "${nombreSolicitud}"...`);

    try {
      const solicitud = await buscarSolicitudVigentePorNombre(params.context, webUrl, nombreSolicitud);
      if (!solicitud?.Id) {
        throw new Error('No se encontró una solicitud vigente con ese nombre.');
      }

      const instanciaId = resolveLookupId(mapInstancias, cambio);
      if (!instanciaId) {
        throw new Error(`No se encontró la instancia "${cambio}" en el lookup Instanciasdeaprobacion.`);
      }

      const processFile = await getCurrentProcessFileBySolicitudId(params.context, webUrl, Number(solicitud.Id));
      if (!processFile?.FileRef) {
        throw new Error(`No se encontró el archivo correspondiente en Procesos para la solicitud ${solicitud.Id}.`);
      }

      await updateListItem(params.context, webUrl, 'Solicitudes', Number(solicitud.Id), {
        InstanciasdeaprobacionId: instanciaId
      });

      await updateFileMetadataByPath(params.context, webUrl, processFile.FileRef, {
        InstanciaDeAprobacionId: instanciaId
      });

      outputRow[idxEstado] = 'OK';
      outputRow[idxDetalle] = `Instancia actualizada a "${cambio}".`;
      outputRow[idxSolicitudId] = Number(solicitud.Id);
      outputRow[idxProceso] = processFile.FileRef;
      outputGrid[rowIndex] = outputRow;

      updated++;
      log(`✅ Solicitud ${solicitud.Id} y archivo en Procesos actualizados a "${cambio}".`);
    } catch (rowError) {
      const message = rowError instanceof Error ? rowError.message : String(rowError);
      outputRow[idxEstado] = 'ERROR';
      outputRow[idxDetalle] = message;
      outputGrid[rowIndex] = outputRow;
      error++;
      log(`❌ ${nombreSolicitud}: ${message}`);
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
