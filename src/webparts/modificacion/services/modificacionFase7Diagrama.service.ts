/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { addAttachment, deleteAttachment, getAllItems, getAttachmentFiles } from './sharepointRest.service';
import { listFilesRecursive } from './spFolderExplorer.service';
import { descargarReporteFase7Diagrama, IFase7DiagramaReportRow } from '../utils/fase7DiagramaReportExcel';
import { IFase7RollbackEntry, rollbackModificacionFase7, snapshotDiagramAttachments } from './modificacionFase7Rollback.service';

type LogFn = (s: string) => void;

type IFase7ExcelRow = {
  nombreArchivo: string;
  nombreDocumento: string;
  documentoPadre: string;
};

type IDiagramaItem = {
  Id: number;
  Title?: string;
  Codigo?: string;
  [key: string]: any;
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
    throw new Error('No se pudo obtener el contenido del Excel de Fase 7.');
  }

  const response = await fetch(url, { credentials: 'same-origin' });
  if (!response.ok) {
    throw new Error(`No se pudo descargar el Excel de Fase 7. HTTP ${response.status}`);
  }

  return response.arrayBuffer();
}

async function readFase7Excel(file: IFilePickerResult): Promise<IFase7ExcelRow[]> {
  const buffer = await readArrayBufferFromFilePicker(file);
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error('No se encontró la hoja del Excel de Fase 7.');
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

  const idxArchivo = findIndex(['Nombre del Archivo', 'NombreArchivo', 'Archivo']);
  const idxDocumento = findIndex(['Nombre del Documento', 'NombreDocumento', 'Documento']);
  const idxPadre = findIndex(['Documento Padre', 'DocumentoPadre', 'Padre']);

  if (idxArchivo === -1 || idxDocumento === -1) {
    throw new Error('El Excel de Fase 7 debe contener las columnas "Nombre del Archivo" y "Nombre del Documento".');
  }

  const rows: IFase7ExcelRow[] = [];
  for (let i = 1; i < aoa.length; i++) {
    const row = aoa[i] || [];
    const nombreArchivo = String(row[idxArchivo] || '').trim();
    const nombreDocumento = String(row[idxDocumento] || '').trim();
    const documentoPadre = idxPadre === -1 ? '' : String(row[idxPadre] || '').trim();

    if (!nombreArchivo && !nombreDocumento && !documentoPadre) {
      continue;
    }

    rows.push({ nombreArchivo, nombreDocumento, documentoPadre });
  }

  return rows;
}

async function buscarDiagramasPorTitulo(context: WebPartContext, webUrl: string, nombreDocumento: string): Promise<IDiagramaItem[]> {
  const wanted = normalizeLooseText(nombreDocumento);
  if (!wanted) {
    return [];
  }

  const url =
    `${webUrl}/_api/web/lists/getbytitle('Diagramas de Flujo')/items` +
    `?$select=Id,Title,Codigo` +
    `&$top=5000`;

  const items = await getAllItems<IDiagramaItem>(context, url);
  return items.filter((item) => normalizeLooseText(item.Title || '') === wanted);
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
    if (exactKey && !exactMap.has(exactKey)) {
      exactMap.set(exactKey, { name: file.Name, url: file.ServerRelativeUrl });
    }
    if (looseKey && !looseMap.has(looseKey)) {
      looseMap.set(looseKey, { name: file.Name, url: file.ServerRelativeUrl });
    }
  }

  return { exactMap, looseMap };
}

function resolveSourceFile(
  fileName: string,
  exactMap: Map<string, { name: string; url: string; }>,
  looseMap: Map<string, { name: string; url: string; }>
): { name: string; url: string; } | null {
  const exact = exactMap.get(String(fileName || '').trim().toLowerCase());
  if (exact) {
    return exact;
  }

  const loose = looseMap.get(normalizeLooseFileKey(fileName || ''));
  return loose || null;
}

async function replaceDiagramAttachment(params: {
  context: WebPartContext;
  webUrl: string;
  diagramItemId: number;
  fileName: string;
  fileUrl: string;
  expectedPreviousAttachmentName: string;
  log: LogFn;
}): Promise<void> {
  const absoluteUrl = `${new URL(params.webUrl).origin}${params.fileUrl}`;
  const response = await fetch(absoluteUrl, { credentials: 'same-origin' });
  if (!response.ok) {
    throw new Error(`No se pudo descargar el archivo origen ${params.fileName}. HTTP ${response.status}`);
  }

  const blob = await response.blob();
  await deleteAttachment(
    params.context,
    params.webUrl,
    'Diagramas de Flujo',
    params.diagramItemId,
    params.expectedPreviousAttachmentName
  );
  await addAttachment(params.context, params.webUrl, 'Diagramas de Flujo', params.diagramItemId, params.fileName, blob);
  params.log(`🧩 Fase 7 | Diagrama actualizado | ID=${params.diagramItemId} | Archivo=${params.fileName}`);
}

export async function ejecutarFase7ModificacionDiagramas(params: {
  context: WebPartContext;
  excelFile: IFilePickerResult;
  sourceFolderServerRelativeUrl: string;
  log?: LogFn;
}): Promise<{
  reportRows: IFase7DiagramaReportRow[];
  rollbackEntries: IFase7RollbackEntry[];
  reportFileName: string;
  processed: number;
  ok: number;
  skipped: number;
  error: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const rows = await readFase7Excel(params.excelFile);
  const reportRows: IFase7DiagramaReportRow[] = [];
  const rollbackEntries: IFase7RollbackEntry[] = [];
  const fileMaps = await buildSourceFileMaps({
    context: params.context,
    webUrl,
    folderServerRelativeUrl: params.sourceFolderServerRelativeUrl,
    log
  });

  let ok = 0;
  let skipped = 0;
  let error = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    let localRollbackEntry: IFase7RollbackEntry | null = null;
    let diagramaId: number | '' = '';
    let sourceFileUrl = '';

    if (!row.nombreArchivo || !row.nombreDocumento) {
      skipped++;
      reportRows.push({
        EstadoFase7: 'SKIP',
        DocumentoPadre: row.documentoPadre || '',
        SolicitudPadreID: '',
        NombreDocumento: row.nombreDocumento || '',
        DiagramaID: '',
        NombreArchivoOrigen: row.nombreArchivo || '',
        RutaArchivoOrigen: '',
        AdjuntosPrevios: '',
        ArchivoCargado: '',
        Error: 'Fila incompleta. Se requieren Nombre del Archivo y Nombre del Documento.'
      });
      continue;
    }

    try {
      const diagramas = await buscarDiagramasPorTitulo(params.context, webUrl, row.nombreDocumento);
      log(`🧭 Fase 7 | Diagramas encontrados por Title para "${row.nombreDocumento}": ${diagramas.length}`);

      if (!diagramas.length) {
        throw new Error(`No se encontró ningún diagrama de flujo con Title "${row.nombreDocumento}".`);
      }

      if (diagramas.length > 1) {
        throw new Error(`Se encontraron ${diagramas.length} diagramas de flujo con el mismo Title "${row.nombreDocumento}". No se puede reemplazar el archivo.`);
      }

      const diagrama = diagramas[0];
      diagramaId = Number(diagrama.Id || 0) || '';

      const sourceFile = resolveSourceFile(row.nombreArchivo, fileMaps.exactMap, fileMaps.looseMap);
      if (!sourceFile?.url) {
        throw new Error(`No se encontró el archivo origen "${row.nombreArchivo}" dentro de la carpeta seleccionada.`);
      }
      sourceFileUrl = sourceFile.url;

      const currentAttachments = await getAttachmentFiles(params.context, webUrl, 'Diagramas de Flujo', diagrama.Id);
      if (currentAttachments.length !== 1) {
        throw new Error(`El diagrama con Title "${row.nombreDocumento}" debe tener exactamente 1 adjunto actual y tiene ${currentAttachments.length}.`);
      }

      const previousAttachments = await snapshotDiagramAttachments({
        context: params.context,
        webUrl,
        itemId: diagrama.Id
      });

      localRollbackEntry = {
        diagramItemId: diagrama.Id,
        solicitudPadreId: 0,
        documentoPadre: row.documentoPadre,
        nombreDocumento: row.nombreDocumento,
        uploadedAttachmentFileName: row.nombreArchivo,
        previousAttachments
      };

      await replaceDiagramAttachment({
        context: params.context,
        webUrl,
        diagramItemId: diagrama.Id,
        fileName: row.nombreArchivo,
        fileUrl: sourceFile.url,
        expectedPreviousAttachmentName: currentAttachments[0].FileName,
        log
      });

      rollbackEntries.push(localRollbackEntry);
      ok++;
      reportRows.push({
        EstadoFase7: 'OK',
        DocumentoPadre: row.documentoPadre,
        SolicitudPadreID: '',
        NombreDocumento: row.nombreDocumento,
        DiagramaID: diagramaId,
        NombreArchivoOrigen: row.nombreArchivo,
        RutaArchivoOrigen: sourceFile.url,
        AdjuntosPrevios: previousAttachments.map((item) => item.fileName).join(' / '),
        ArchivoCargado: row.nombreArchivo,
        Error: ''
      });
    } catch (fase7Error) {
      const message = fase7Error instanceof Error ? fase7Error.message : String(fase7Error);

      if (localRollbackEntry) {
        try {
          await rollbackModificacionFase7({
            context: params.context,
            webUrl,
            entries: [localRollbackEntry],
            log
          });
          log(`↩️ Fase 7 | Rollback local ejecutado para el diagrama ${localRollbackEntry.diagramItemId}`);
        } catch (rollbackError) {
          const rollbackMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
          log(`⚠️ Fase 7 | Falló el rollback local del diagrama ${localRollbackEntry.diagramItemId}: ${rollbackMessage}`);
        }
      }

      error++;
      reportRows.push({
        EstadoFase7: 'ERROR',
        DocumentoPadre: row.documentoPadre || '',
        SolicitudPadreID: '',
        NombreDocumento: row.nombreDocumento || '',
        DiagramaID: diagramaId,
        NombreArchivoOrigen: row.nombreArchivo || '',
        RutaArchivoOrigen: sourceFileUrl,
        AdjuntosPrevios: localRollbackEntry ? localRollbackEntry.previousAttachments.map((item) => item.fileName).join(' / ') : '',
        ArchivoCargado: '',
        Error: message
      });
      log(`❌ Error Fase 7 | Title="${row.nombreDocumento}" | ${message}`);
    }
  }

  const now = new Date();
  const pad = (value: number): string => String(value).padStart(2, '0');
  const reportFileName =
    `Reporte_Fase7_DIAGRAMAS_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_` +
    `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.xlsx`;
  descargarReporteFase7Diagrama(reportRows, reportFileName);

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
