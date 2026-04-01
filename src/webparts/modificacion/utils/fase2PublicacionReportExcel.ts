/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';

export interface IFase2PublicacionReportRow {
  EstadoFase2: string;
  SolicitudOrigenID?: number | '';
  SolicitudID?: number | '';
  NombreDocumento: string;
  NombreArchivo: string;
  CodigoDocumento: string;
  ArchivoProcesoOriginal: string;
  RutaProcesoOriginal: string;
  ArchivoProcesoRenombrado: string;
  RutaProcesoRenombrada: string;
  RutaHistorico: string;
  RutaNuevoPublicado: string;
  VersionDocumentoAnterior: string;
  VersionDocumentoNueva: string;
  FechaBajaHistorico: string;
  FechaAprobacionBaja: string;
  Error: string;
}

const headers = [
  'EstadoFase2',
  'SolicitudOrigenID',
  'SolicitudID',
  'NombreDocumento',
  'NombreArchivo',
  'CodigoDocumento',
  'ArchivoProcesoOriginal',
  'RutaProcesoOriginal',
  'ArchivoProcesoRenombrado',
  'RutaProcesoRenombrada',
  'RutaHistorico',
  'RutaNuevoPublicado',
  'VersionDocumentoAnterior',
  'VersionDocumentoNueva',
  'FechaBajaHistorico',
  'FechaAprobacionBaja',
  'Error'
];

function autoFitColumns(rows: IFase2PublicacionReportRow[]): Array<{ wch: number; }> {
  const widths = headers.map((header) => ({ wch: header.length + 2 }));

  rows.forEach((row) => {
    headers.forEach((header, index) => {
      const value = String((row as any)[header] ?? '');
      widths[index].wch = Math.min(Math.max(widths[index].wch, value.length + 2), 90);
    });
  });

  return widths;
}

export function descargarReporteFase2Publicacion(rows: IFase2PublicacionReportRow[], fileName?: string): void {
  const safeRows = Array.isArray(rows) ? rows : [];
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(safeRows, { header: headers });

  worksheet['!cols'] = autoFitColumns(safeRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Fase2_PUBLICACION');

  const now = new Date();
  const pad = (value: number): string => String(value).padStart(2, '0');
  const defaultName =
    `Reporte_Fase2_PUBLICACION_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_` +
    `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.xlsx`;

  XLSX.writeFile(workbook, fileName || defaultName);
}
