/* eslint-disable */
import * as XLSX from 'xlsx';
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';

export interface IExcelRevisionRow {
  rowNumber: number;
  documentName: string;
}

export interface IExcelRevisionSession {
  fileName: string;
  workbook: XLSX.WorkBook;
  sheetName: string;
  grid: any[][];
  rows: IExcelRevisionRow[];
}

export async function openExcelRevisionSession(file: IFilePickerResult): Promise<IExcelRevisionSession> {
  const fileName = file.fileName || '';
  ensureSupportedExcel(fileName);

  const arrayBuffer = await downloadExcelAsArrayBuffer(file);
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error('El Excel no tiene hojas.');
  }

  const worksheet = workbook.Sheets[sheetName];
  const grid = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
    raw: false
  }) as any[][];
  const rows = buildRevisionRows(grid);

  return {
    fileName,
    workbook,
    sheetName,
    grid,
    rows
  };
}

export async function buildReviewedWorkbook(
  session: IExcelRevisionSession,
  valuesByRowNumber: Map<number, { solicitudId: string; documentosHijos: string; diagramasFlujo: string; }>
): Promise<Blob> {
  const grid = cloneGrid(session.grid);

  if (!grid.length) {
    grid.push([]);
  }

  const headerRowIndex = 0;
  const headerRow = grid[headerRowIndex] || [];
  const solicitudIdColumnIndex = headerRow.length;
  const documentosHijosColumnIndex = headerRow.length + 1;
  const diagramasFlujoColumnIndex = headerRow.length + 2;

  headerRow[solicitudIdColumnIndex] = 'ID Solicitud';
  headerRow[documentosHijosColumnIndex] = 'Documentos hijos';
  headerRow[diagramasFlujoColumnIndex] = 'Diagrama de Flujos';
  grid[headerRowIndex] = headerRow;

  valuesByRowNumber.forEach((value, rowNumber) => {
    const gridRowIndex = rowNumber - 1;

    while (grid.length <= gridRowIndex) {
      grid.push([]);
    }

    const row = grid[gridRowIndex] || [];
    row[solicitudIdColumnIndex] = value.solicitudId || '';
    row[documentosHijosColumnIndex] = value.documentosHijos || '';
    row[diagramasFlujoColumnIndex] = value.diagramasFlujo || '';
    grid[gridRowIndex] = row;
  });

  const newWorksheet = XLSX.utils.aoa_to_sheet(grid);
  session.workbook.Sheets[session.sheetName] = newWorksheet;

  const output = XLSX.write(session.workbook, {
    bookType: getBookType(session.fileName),
    type: 'array'
  });

  return new Blob([output], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}

export function buildReviewedFileName(fileName: string): string {
  const lastDot = fileName.lastIndexOf('.');
  if (lastDot === -1) {
    return `${fileName}_revisado.xlsx`;
  }

  const baseName = fileName.substring(0, lastDot);
  const extension = fileName.substring(lastDot);
  return `${baseName}_revisado${extension}`;
}

async function downloadExcelAsArrayBuffer(file: IFilePickerResult): Promise<ArrayBuffer> {
  const url = file.fileAbsoluteUrl || '';
  if (!url) {
    throw new Error('No se obtuvo fileAbsoluteUrl del Excel.');
  }

  const response = await fetch(url, { credentials: 'same-origin' });
  if (!response.ok) {
    throw new Error(`No se pudo descargar Excel. HTTP ${response.status}`);
  }

  return response.arrayBuffer();
}

function buildRevisionRows(grid: any[][]): IExcelRevisionRow[] {
  const rows: IExcelRevisionRow[] = [];

  for (let i = 1; i < grid.length; i++) {
    const row = grid[i] || [];
    rows.push({
      rowNumber: i + 1,
      documentName: String(row[9] || '').trim()
    });
  }

  return rows;
}

function cloneGrid(grid: any[][]): any[][] {
  const result: any[][] = [];

  for (let i = 0; i < grid.length; i++) {
    result.push(Array.isArray(grid[i]) ? [...grid[i]] : []);
  }

  return result;
}

function getBookType(fileName: string): XLSX.BookType {
  const lowerName = String(fileName || '').toLowerCase();
  return lowerName.indexOf('.xlsm') !== -1 ? 'xlsm' : 'xlsx';
}

function ensureSupportedExcel(fileName: string): void {
  const lowerName = String(fileName || '').toLowerCase();

  if (lowerName.indexOf('.xlsx') !== -1 || lowerName.indexOf('.xlsm') !== -1) {
    return;
  }

  throw new Error('Selecciona un archivo Excel válido (.xlsx o .xlsm).');
}
