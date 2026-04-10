/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';

export interface ICopiarHijosReportRow {
  PadreOrigenId: number;
  PadreOrigenNombre: string;
  PadreDestinoId: number;
  PadreDestinoNombre: string;
  HijoId: number | '';
  HijoNombre: string;
  RelacionOrigenId: number | '';
  RelacionNuevaId: number | '';
  RelacionNuevaTitle: string;
  DocPadresAntes: string;
  DocPadresDespues: string;
  Estado: string;
  Observaciones: string;
}

const headers = [
  'PadreOrigenId',
  'PadreOrigenNombre',
  'PadreDestinoId',
  'PadreDestinoNombre',
  'HijoId',
  'HijoNombre',
  'RelacionOrigenId',
  'RelacionNuevaId',
  'RelacionNuevaTitle',
  'DocPadresAntes',
  'DocPadresDespues',
  'Estado',
  'Observaciones'
];

function autoFitColumns(rows: ICopiarHijosReportRow[]): Array<{ wch: number; }> {
  const widths = headers.map((header) => ({ wch: header.length + 2 }));

  rows.forEach((row) => {
    headers.forEach((header, index) => {
      const value = String((row as any)[header] ?? '');
      widths[index].wch = Math.min(Math.max(widths[index].wch, value.length + 2), 60);
    });
  });

  return widths;
}

export function buildCopiarHijosWorkbook(
  rows: ICopiarHijosReportRow[],
  fileName?: string
): { blob: Blob; fileName: string; } {
  const safeRows = Array.isArray(rows) ? rows : [];
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(safeRows, { header: headers });

  worksheet['!cols'] = autoFitColumns(safeRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'CopiarHijos');

  const output = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array'
  });

  const now = new Date();
  const pad = (value: number): string => String(value).padStart(2, '0');
  const defaultName =
    `Copiar_Hijos_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_` +
    `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.xlsx`;

  return {
    blob: new Blob([output], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }),
    fileName: fileName || defaultName
  };
}
