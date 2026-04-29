/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';

export interface IProcesoSinSolicitudReportRow {
  Id: number;
  NombreArchivo: string;
  RutaCarpeta: string;
  UrlArchivo: string;
  Title: string;
  NombreDocumento: string;
  CodigoDocumento: string;
  VersionDocumento: string;
  TipoDocumento: string;
  CategoriaDocumento: string;
  SolicitudId: string;
  Created: string;
  Modified: string;
  CreadoPor: string;
  ModificadoPor: string;
}

const headers = [
  'Id',
  'NombreArchivo',
  'RutaCarpeta',
  'UrlArchivo',
  'Title',
  'NombreDocumento',
  'CodigoDocumento',
  'VersionDocumento',
  'TipoDocumento',
  'CategoriaDocumento',
  'SolicitudId',
  'Created',
  'Modified',
  'CreadoPor',
  'ModificadoPor'
];

function autoFitColumns(rows: IProcesoSinSolicitudReportRow[]): Array<{ wch: number; }> {
  const widths = headers.map((header) => ({ wch: header.length + 2 }));

  rows.forEach((row) => {
    headers.forEach((header, index) => {
      const value = String((row as any)[header] ?? '');
      widths[index].wch = Math.min(Math.max(widths[index].wch, value.length + 2), 70);
    });
  });

  return widths;
}

function formatDateValue(value: any): string {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  if (value instanceof Date && !isNaN(value.getTime())) {
    return buildDateString(value);
  }

  const raw = String(value).trim();
  if (!raw) {
    return '';
  }

  const isoMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s](\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (isoMatch) {
    const year = isoMatch[1];
    const month = isoMatch[2];
    const day = isoMatch[3];
    const hour = isoMatch[4];
    const minute = isoMatch[5];
    const second = isoMatch[6];
    const datePart = `${day}/${month}/${year}`;

    if (hour && minute) {
      return `${datePart} ${hour}:${minute}${second ? `:${second}` : ''}`;
    }

    return datePart;
  }

  return raw;
}

function buildDateString(value: Date): string {
  const pad = (input: number): string => String(input).padStart(2, '0');
  return `${pad(value.getDate())}/${pad(value.getMonth() + 1)}/${value.getFullYear()} ${pad(value.getHours())}:${pad(value.getMinutes())}:${pad(value.getSeconds())}`;
}

function normalizeRows(rows: IProcesoSinSolicitudReportRow[]): IProcesoSinSolicitudReportRow[] {
  return rows.map((row) => ({
    ...row,
    Created: formatDateValue(row.Created),
    Modified: formatDateValue(row.Modified)
  }));
}

export function buildProcesosSinSolicitudWorkbook(
  rows: IProcesoSinSolicitudReportRow[],
  fileName?: string
): { blob: Blob; fileName: string; } {
  const safeRows = normalizeRows(Array.isArray(rows) ? rows : []);
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(safeRows, { header: headers });

  worksheet['!cols'] = autoFitColumns(safeRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'ProcesosSinSolicitud');

  const output = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array'
  });

  const now = new Date();
  const pad = (value: number): string => String(value).padStart(2, '0');
  const defaultName =
    `Procesos_Sin_Solicitud_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_` +
    `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.xlsx`;

  return {
    blob: new Blob([output], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }),
    fileName: fileName || defaultName
  };
}
