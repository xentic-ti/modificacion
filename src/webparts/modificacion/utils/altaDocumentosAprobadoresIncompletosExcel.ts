/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';

export interface IAltaDocumentoAprobadorIncompletoRow {
  SolicitudId: number;
  SolicitudTitulo: string;
  NombreDocumento: string;
  CodigoDocumento: string;
  VersionDocumento: string;
  TipoDocumento: string;
  MotivoSolicitud: string;
  AreasImpactadas: string;
  AccionSolicitud: string;
  EstadoSolicitud: string;
  TipoAprobadorEsperado: string;
  OrigenAprobadorEsperado: string;
  AprobadorEsperadoId: number | '';
  AprobadorEsperadoNombre: string;
  AprobadorEsperadoEmail: string;
  RegistroAprobadorId: number | '';
  Rol: string;
  AprobadorId: number | '';
  AprobadorNombre: string;
  AprobadorEmail: string;
  ImpactadoPorArea: string;
  ImpactadoPorMotivo: string;
  ImpactadoPorAccion: string;
  MotivoIncompleto: string;
  CreatedSolicitud: string;
  ModifiedSolicitud: string;
  CreatedAprobador: string;
  ModifiedAprobador: string;
}

const headers = [
  'SolicitudId',
  'SolicitudTitulo',
  'NombreDocumento',
  'CodigoDocumento',
  'VersionDocumento',
  'TipoDocumento',
  'MotivoSolicitud',
  'AreasImpactadas',
  'AccionSolicitud',
  'EstadoSolicitud',
  'TipoAprobadorEsperado',
  'OrigenAprobadorEsperado',
  'AprobadorEsperadoId',
  'AprobadorEsperadoNombre',
  'AprobadorEsperadoEmail',
  'RegistroAprobadorId',
  'Rol',
  'AprobadorId',
  'AprobadorNombre',
  'AprobadorEmail',
  'ImpactadoPorArea',
  'ImpactadoPorMotivo',
  'ImpactadoPorAccion',
  'MotivoIncompleto',
  'CreatedSolicitud',
  'ModifiedSolicitud',
  'CreatedAprobador',
  'ModifiedAprobador'
];

function autoFitColumns(rows: IAltaDocumentoAprobadorIncompletoRow[]): Array<{ wch: number; }> {
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

function normalizeRows(rows: IAltaDocumentoAprobadorIncompletoRow[]): IAltaDocumentoAprobadorIncompletoRow[] {
  return rows.map((row) => ({
    ...row,
    CreatedSolicitud: formatDateValue(row.CreatedSolicitud),
    ModifiedSolicitud: formatDateValue(row.ModifiedSolicitud),
    CreatedAprobador: formatDateValue(row.CreatedAprobador),
    ModifiedAprobador: formatDateValue(row.ModifiedAprobador)
  }));
}

export function buildAltaDocumentosAprobadoresIncompletosWorkbook(
  rows: IAltaDocumentoAprobadorIncompletoRow[],
  fileName?: string
): { blob: Blob; fileName: string; } {
  const safeRows = normalizeRows(Array.isArray(rows) ? rows : []);
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(safeRows, { header: headers });

  worksheet['!cols'] = autoFitColumns(safeRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'AprobadoresIncompletos');

  const output = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array'
  });

  const now = new Date();
  const pad = (value: number): string => String(value).padStart(2, '0');
  const defaultName =
    `Alta_Documentos_Aprobadores_Incompletos_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_` +
    `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.xlsx`;

  return {
    blob: new Blob([output], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }),
    fileName: fileName || defaultName
  };
}
