/* eslint-disable */
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { escapeODataValue, getAllItems } from './sharepointRest.service';
import {
  buildProcesosSinSolicitudWorkbook,
  IProcesoSinSolicitudReportRow
} from '../utils/procesosSinSolicitudExcel';

type LogFn = (message: string) => void;

const PROCESOS_ROOT = '/sites/SistemadeGestionDocumental/Procesos';

interface IProcesoFileItem {
  Id: number;
  Title?: string;
  FileLeafRef?: string;
  FileRef?: string;
  FileDirRef?: string;
  FileSystemObjectType?: number;
  SolicitudId?: number | string | null;
  NombreDocumento?: string;
  Codigodedocumento?: string;
  VersionDocumento?: string;
  Tipodedocumento?: string;
  CategoriaDocumento?: string;
  Created?: string;
  Modified?: string;
  Author?: {
    Title?: string;
  };
  Editor?: {
    Title?: string;
  };
}

export async function exportarProcesosSinSolicitud(params: {
  context: WebPartContext;
  log?: LogFn;
}): Promise<{
  blob: Blob;
  fileName: string;
  totalDocumentos: number;
  totalSinSolicitud: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;

  log('🔎 Consultando documentos de la biblioteca Procesos...');
  const items = await obtenerDocumentosProcesos(params.context, webUrl);
  log(`📄 Documentos leidos en Procesos: ${items.length}`);

  const rows: IProcesoSinSolicitudReportRow[] = [];
  const origin = new URL(webUrl).origin;

  for (let i = 0; i < items.length; i++) {
    const item = items[i];

    if (hasSolicitud(item.SolicitudId)) {
      continue;
    }

    rows.push({
      Id: Number(item.Id || 0),
      NombreArchivo: String(item.FileLeafRef || '').trim(),
      RutaCarpeta: String(item.FileDirRef || getFolderFromFileRef(item.FileRef || '')).trim(),
      UrlArchivo: buildAbsoluteUrl(origin, item.FileRef || ''),
      Title: String(item.Title || '').trim(),
      NombreDocumento: String(item.NombreDocumento || '').trim(),
      CodigoDocumento: String(item.Codigodedocumento || '').trim(),
      VersionDocumento: String(item.VersionDocumento || '').trim(),
      TipoDocumento: String(item.Tipodedocumento || '').trim(),
      CategoriaDocumento: String(item.CategoriaDocumento || '').trim(),
      SolicitudId: '',
      Created: String(item.Created || '').trim(),
      Modified: String(item.Modified || '').trim(),
      CreadoPor: String(item.Author?.Title || '').trim(),
      ModificadoPor: String(item.Editor?.Title || '').trim()
    });
  }

  rows.sort(compareRows);

  if (rows.length) {
    log(`⚠️ Documentos sin Solicitud: ${rows.length}`);
  } else {
    log('✅ Todos los documentos de Procesos tienen Solicitud informada.');
  }

  const report = buildProcesosSinSolicitudWorkbook(rows);

  return {
    blob: report.blob,
    fileName: report.fileName,
    totalDocumentos: items.length,
    totalSinSolicitud: rows.length
  };
}

async function obtenerDocumentosProcesos(context: WebPartContext, webUrl: string): Promise<IProcesoFileItem[]> {
  const url =
    `${webUrl}/_api/web/GetList('${escapeODataValue(PROCESOS_ROOT)}')/items` +
    `?$select=Id,Title,FileLeafRef,FileRef,FileDirRef,FileSystemObjectType,SolicitudId,NombreDocumento,` +
    `Codigodedocumento,VersionDocumento,Tipodedocumento,CategoriaDocumento,Created,Modified,Author/Title,Editor/Title` +
    `&$expand=Author,Editor` +
    `&$top=5000`;

  const items = await getAllItems<IProcesoFileItem>(context, url);
  return items.filter((item) => Number(item.FileSystemObjectType || 0) === 0 && !isSystemPath(item.FileRef || ''));
}

function hasSolicitud(value: any): boolean {
  if (value === null || value === undefined || value === '') {
    return false;
  }

  const numeric = Number(value);
  if (Number.isFinite(numeric)) {
    return numeric > 0;
  }

  return String(value || '').trim() !== '';
}

function isSystemPath(fileRef: string): boolean {
  const normalized = String(fileRef || '').toLowerCase();
  return normalized.indexOf('/forms/') !== -1;
}

function getFolderFromFileRef(fileRef: string): string {
  const clean = String(fileRef || '').trim();
  const index = clean.lastIndexOf('/');
  return index > -1 ? clean.substring(0, index) : '';
}

function buildAbsoluteUrl(origin: string, fileRef: string): string {
  const clean = String(fileRef || '').trim();
  if (!clean) {
    return '';
  }

  if (/^https?:\/\//i.test(clean)) {
    return clean;
  }

  return `${origin}${clean}`;
}

function compareRows(a: IProcesoSinSolicitudReportRow, b: IProcesoSinSolicitudReportRow): number {
  const folder = String(a.RutaCarpeta || '').localeCompare(String(b.RutaCarpeta || ''), 'es');
  if (folder !== 0) {
    return folder;
  }

  return String(a.NombreArchivo || '').localeCompare(String(b.NombreArchivo || ''), 'es');
}
