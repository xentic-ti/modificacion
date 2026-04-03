/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';

export interface IRelacionesDocumentosAuditRow {
  RelacionId: number | '';
  RelacionTitle: string;
  DocumentoPadreId: number | '';
  DocumentoPadreNombre: string;
  DocumentoHijoId: number | '';
  DocumentoHijoNombre: string;
  EstadoRelacion: string;
  Observaciones: string;
  HijoDocPadresActual: string;
}

export interface ISolicitudHijaSinDocPadreRow {
  RelacionId: number | '';
  SolicitudHijaId: number;
  SolicitudHijaNombre: string;
  SolicitudPadreEsperadoId: number | '';
  SolicitudPadreEsperadoNombre: string;
  HijoDocPadresActual: string;
  EstadoDocPadres: string;
  Observaciones: string;
}

const auditHeaders = [
  'RelacionId',
  'RelacionTitle',
  'DocumentoPadreId',
  'DocumentoPadreNombre',
  'DocumentoHijoId',
  'DocumentoHijoNombre',
  'EstadoRelacion',
  'Observaciones',
  'HijoDocPadresActual'
];

const missingDocPadreHeaders = [
  'RelacionId',
  'SolicitudHijaId',
  'SolicitudHijaNombre',
  'SolicitudPadreEsperadoId',
  'SolicitudPadreEsperadoNombre',
  'HijoDocPadresActual',
  'EstadoDocPadres',
  'Observaciones'
];

function autoFitColumns(rows: any[], headers: string[]): Array<{ wch: number; }> {
  const widths = headers.map((header) => ({ wch: header.length + 2 }));

  rows.forEach((row) => {
    headers.forEach((header, index) => {
      const value = String((row as any)[header] ?? '');
      widths[index].wch = Math.min(Math.max(widths[index].wch, value.length + 2), 60);
    });
  });

  return widths;
}

export function buildRelacionesDocumentosAuditWorkbook(params: {
  auditRows: IRelacionesDocumentosAuditRow[];
  missingDocPadreRows: ISolicitudHijaSinDocPadreRow[];
  fileName?: string;
}): { blob: Blob; fileName: string; } {
  const auditRows = Array.isArray(params.auditRows) ? params.auditRows : [];
  const missingDocPadreRows = Array.isArray(params.missingDocPadreRows) ? params.missingDocPadreRows : [];

  const workbook = XLSX.utils.book_new();
  const auditSheet = XLSX.utils.json_to_sheet(auditRows, { header: auditHeaders });
  const missingSheet = XLSX.utils.json_to_sheet(missingDocPadreRows, { header: missingDocPadreHeaders });

  auditSheet['!cols'] = autoFitColumns(auditRows, auditHeaders);
  missingSheet['!cols'] = autoFitColumns(missingDocPadreRows, missingDocPadreHeaders);

  XLSX.utils.book_append_sheet(workbook, auditSheet, 'RelacionesAudit');
  XLSX.utils.book_append_sheet(workbook, missingSheet, 'HijosSinDocPadre');

  const output = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array'
  });

  const now = new Date();
  const pad = (value: number): string => String(value).padStart(2, '0');
  const defaultName =
    `Auditoria_Relaciones_Documentos_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_` +
    `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.xlsx`;

  return {
    blob: new Blob([output], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }),
    fileName: params.fileName || defaultName
  };
}
