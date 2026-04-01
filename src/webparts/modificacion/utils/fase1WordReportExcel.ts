/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';

export interface IFase1WordReportRow {
  SolicitudOrigenID?: number | '';
  SolicitudID: number | '';
  NombreDocumento: string;
  NombreArchivo: string;
  CodigoDocumento: string;
  VersionDocumento?: string;
  TieneDocumentoPadre: string;
  DocumentoPadreNombre: string;
  DocumentoPadreSolicitudID: number | '';
  CodigoDocumentoPadre?: string;
  PadreRegeneradoConLinks: string;
  RutaTemporalWord: string;
  EstadoFase1: string;
  Error: string;
  TipoDocumento?: string;
  CategoriaDocumento?: string;
  Clasificaciondeproceso?: string;
  Macroproceso?: string;
  Proceso?: string;
  Subproceso?: string;
  AreaDuena?: string;
  AreaImpactada?: string;
  Resumen?: string;
  FechaDeAprobacion?: string;
  FechaDeVigencia?: string;
  InstanciaDeAprobacionId?: number | '';
  MetadataPendiente?: string;
  DocumentosHijosIDs?: string;
  DocumentosHijosNombres?: string;
  DocumentoPadreSolicitudAnteriorID?: number | '';
  DocumentoPadreSolicitudNuevaID?: number | '';
  DiagramasFlujoNombres?: string;
}

const headers = [
  'SolicitudOrigenID',
  'SolicitudID',
  'NombreDocumento',
  'NombreArchivo',
  'CodigoDocumento',
  'VersionDocumento',
  'TieneDocumentoPadre',
  'DocumentoPadreNombre',
  'DocumentoPadreSolicitudID',
  'CodigoDocumentoPadre',
  'PadreRegeneradoConLinks',
  'RutaTemporalWord',
  'EstadoFase1',
  'Error',
  'TipoDocumento',
  'CategoriaDocumento',
  'Clasificaciondeproceso',
  'Macroproceso',
  'Proceso',
  'Subproceso',
  'AreaDuena',
  'AreaImpactada',
  'Resumen',
  'FechaDeAprobacion',
  'FechaDeVigencia',
  'InstanciaDeAprobacionId',
  'MetadataPendiente',
  'DocumentosHijosIDs',
  'DocumentosHijosNombres',
  'DocumentoPadreSolicitudAnteriorID',
  'DocumentoPadreSolicitudNuevaID',
  'DiagramasFlujoNombres'
];

function autoFitColumns(rows: IFase1WordReportRow[]): Array<{ wch: number; }> {
  const widths = headers.map((header) => ({ wch: header.length + 2 }));

  rows.forEach((row) => {
    headers.forEach((header, index) => {
      const value = String((row as any)[header] ?? '');
      widths[index].wch = Math.min(Math.max(widths[index].wch, value.length + 2), 80);
    });
  });

  return widths;
}

export function descargarReporteFase1Word(rows: IFase1WordReportRow[], fileName?: string): void {
  const safeRows = Array.isArray(rows) ? rows : [];
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(safeRows, { header: headers });

  worksheet['!cols'] = autoFitColumns(safeRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Fase1_WORD');

  const now = new Date();
  const pad = (value: number): string => String(value).padStart(2, '0');
  const defaultName =
    `Reporte_Fase1_WORD_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_` +
    `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.xlsx`;

  XLSX.writeFile(workbook, fileName || defaultName);
}
