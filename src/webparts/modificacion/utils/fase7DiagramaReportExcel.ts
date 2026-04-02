/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';

export interface IFase7DiagramaReportRow {
  EstadoFase7: string;
  DocumentoPadre: string;
  SolicitudPadreID?: number | '';
  NombreDocumento: string;
  DiagramaID?: number | '';
  NombreArchivoOrigen: string;
  RutaArchivoOrigen: string;
  AdjuntosPrevios: string;
  ArchivoCargado: string;
  Error: string;
}

const headers = [
  'EstadoFase7',
  'DocumentoPadre',
  'SolicitudPadreID',
  'NombreDocumento',
  'DiagramaID',
  'NombreArchivoOrigen',
  'RutaArchivoOrigen',
  'AdjuntosPrevios',
  'ArchivoCargado',
  'Error'
];

function autoFitColumns(rows: IFase7DiagramaReportRow[]): Array<{ wch: number; }> {
  const widths = headers.map((header) => ({ wch: header.length + 2 }));

  rows.forEach((row) => {
    headers.forEach((header, index) => {
      const value = String((row as any)[header] ?? '');
      widths[index].wch = Math.min(Math.max(widths[index].wch, value.length + 2), 90);
    });
  });

  return widths;
}

export function descargarReporteFase7Diagrama(rows: IFase7DiagramaReportRow[], fileName?: string): void {
  const safeRows = Array.isArray(rows) ? rows : [];
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(safeRows, { header: headers });

  worksheet['!cols'] = autoFitColumns(safeRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Fase7_DIAGRAMAS');

  const now = new Date();
  const pad = (value: number): string => String(value).padStart(2, '0');
  const defaultName =
    `Reporte_Fase7_DIAGRAMAS_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_` +
    `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.xlsx`;

  XLSX.writeFile(workbook, fileName || defaultName);
}
