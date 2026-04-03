/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';

export interface IDocPadresFixReportRow {
  SolicitudHijaId: number;
  SolicitudHijaNombre: string;
  RelacionesIds: string;
  SolicitudesPadreEsperadas: string;
  DocPadresAntes: string;
  DocPadresDespues: string;
  Estado: string;
  Observaciones: string;
}

const headers = [
  'SolicitudHijaId',
  'SolicitudHijaNombre',
  'RelacionesIds',
  'SolicitudesPadreEsperadas',
  'DocPadresAntes',
  'DocPadresDespues',
  'Estado',
  'Observaciones'
];

function autoFitColumns(rows: IDocPadresFixReportRow[]): Array<{ wch: number; }> {
  const widths = headers.map((header) => ({ wch: header.length + 2 }));

  rows.forEach((row) => {
    headers.forEach((header, index) => {
      const value = String((row as any)[header] ?? '');
      widths[index].wch = Math.min(Math.max(widths[index].wch, value.length + 2), 60);
    });
  });

  return widths;
}

export function buildDocPadresFixWorkbook(
  rows: IDocPadresFixReportRow[],
  fileName?: string
): { blob: Blob; fileName: string; } {
  const safeRows = Array.isArray(rows) ? rows : [];
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(safeRows, { header: headers });

  worksheet['!cols'] = autoFitColumns(safeRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'DocPadresCorregidos');

  const output = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array'
  });

  const now = new Date();
  const pad = (value: number): string => String(value).padStart(2, '0');
  const defaultName =
    `Correccion_DocPadres_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_` +
    `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.xlsx`;

  return {
    blob: new Blob([output], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }),
    fileName: fileName || defaultName
  };
}
