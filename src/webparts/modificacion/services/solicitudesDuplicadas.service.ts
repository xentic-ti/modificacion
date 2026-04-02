/* eslint-disable */
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { getAllItems } from './sharepointRest.service';
import {
  buildSolicitudesDuplicadasWorkbook,
  ISolicitudDuplicadaReportRow
} from '../utils/solicitudesDuplicadasExcel';

type LogFn = (message: string) => void;

interface ISolicitudItem {
  Id: number;
  Title?: string;
  NombreDocumento?: string;
  CodigoDocumento?: string;
  VersionDocumento?: string;
  EsVersionActualDocumento?: boolean | string | number;
  DocumentosApoyo?: boolean | string | number;
  EstadoId?: number;
  FechaDeAprobacionSolicitud?: string;
  FechadeVigencia?: string;
  FechaDePublicacionSolicitud?: string;
  Created?: string;
  Modified?: string;
}

export async function buscarSolicitudesDuplicadas(params: {
  context: WebPartContext;
  log?: LogFn;
}): Promise<{
  blob: Blob;
  fileName: string;
  totalSolicitudes: number;
  duplicatedGroups: number;
  duplicatedRows: number;
  nonCurrentRows: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;

  log('🔎 Consultando la lista Solicitudes para detectar versiones duplicadas...');
  const solicitudes = await obtenerSolicitudes(params.context, webUrl);
  log(`📋 Solicitudes leidas: ${solicitudes.length}`);

  const grouped = new Map<string, ISolicitudItem[]>();

  for (let i = 0; i < solicitudes.length; i++) {
    const item = solicitudes[i];
    const key = buildDuplicateKey(item);

    if (!key) {
      continue;
    }

    if (!grouped.has(key)) {
      grouped.set(key, []);
    }

    grouped.get(key)!.push(item);
  }

  const duplicateRows: ISolicitudDuplicadaReportRow[] = [];
  let duplicatedGroups = 0;
  let nonCurrentRows = 0;

  grouped.forEach((items) => {
    if (items.length < 2) {
      return;
    }

    duplicatedGroups++;

    const sortedItems = items.slice().sort((a, b) => Number(a.Id || 0) - Number(b.Id || 0));

    for (let i = 0; i < sortedItems.length; i++) {
      const item = sortedItems[i];
      if (!isTruthyField(item.EsVersionActualDocumento)) {
        nonCurrentRows++;
      }

      duplicateRows.push(buildReportRow(item, sortedItems.length));
    }
  });

  duplicateRows.sort(compareReportRows);

  if (duplicateRows.length) {
    log(`⚠️ Grupos duplicados detectados: ${duplicatedGroups}`);
    log(`⚠️ Solicitudes duplicadas detectadas: ${duplicateRows.length}`);
    log(`ℹ️ Solicitudes no vigentes dentro de duplicados: ${nonCurrentRows}`);
  } else {
    log('✅ No se encontraron solicitudes con la misma version exacta para el mismo documento.');
  }

  const report = buildSolicitudesDuplicadasWorkbook(duplicateRows);

  return {
    blob: report.blob,
    fileName: report.fileName,
    totalSolicitudes: solicitudes.length,
    duplicatedGroups,
    duplicatedRows: duplicateRows.length,
    nonCurrentRows
  };
}

async function obtenerSolicitudes(context: WebPartContext, webUrl: string): Promise<ISolicitudItem[]> {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items` +
    `?$select=Id,Title,NombreDocumento,CodigoDocumento,VersionDocumento,EsVersionActualDocumento,DocumentosApoyo,` +
    `EstadoId,FechaDeAprobacionSolicitud,FechadeVigencia,FechaDePublicacionSolicitud,Created,Modified` +
    `&$top=5000`;

  return getAllItems<ISolicitudItem>(context, url);
}

function buildDuplicateKey(item: ISolicitudItem): string {
  const documentKey = normalizeKeyPart(item.Title);
  const versionKey = String(item.VersionDocumento || '').trim();

  if (!documentKey || !versionKey) {
    return '';
  }

  return `${documentKey}||${versionKey}`;
}

function normalizeKeyPart(value: any): string {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function buildReportRow(item: ISolicitudItem, totalCoincidencias: number): ISolicitudDuplicadaReportRow {
  return {
    ClaveTitle: String(item.Title || '').trim(),
    TotalCoincidencias: totalCoincidencias,
    Id: Number(item.Id || 0),
    Title: String(item.Title || '').trim(),
    NombreDocumento: String(item.NombreDocumento || '').trim(),
    CodigoDocumento: String(item.CodigoDocumento || '').trim(),
    VersionDocumento: String(item.VersionDocumento || '').trim(),
    EsVersionActualDocumento: formatBooleanField(item.EsVersionActualDocumento),
    DocumentosApoyo: formatBooleanField(item.DocumentosApoyo),
    EstadoId: Number(item.EstadoId || 0) || '',
    FechaDeAprobacionSolicitud: String(item.FechaDeAprobacionSolicitud || '').trim(),
    FechadeVigencia: String(item.FechadeVigencia || '').trim(),
    FechaDePublicacionSolicitud: String(item.FechaDePublicacionSolicitud || '').trim(),
    Created: String(item.Created || '').trim(),
    Modified: String(item.Modified || '').trim()
  };
}

function formatBooleanField(value: any): string {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  return isTruthyField(value) ? 'Si' : 'No';
}

function isTruthyField(value: any): boolean {
  if (typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'number') {
    return value === 1;
  }

  const normalized = String(value || '').trim().toLowerCase();
  return normalized === '1' || normalized === 'true' || normalized === 'si' || normalized === 'sí' || normalized === 'yes';
}

function compareReportRows(a: ISolicitudDuplicadaReportRow, b: ISolicitudDuplicadaReportRow): number {
  const byDocument = String(a.ClaveTitle || '').localeCompare(String(b.ClaveTitle || ''), 'es', { sensitivity: 'base' });
  if (byDocument !== 0) {
    return byDocument;
  }

  const byVersion = String(a.VersionDocumento || '').localeCompare(String(b.VersionDocumento || ''), 'es', { sensitivity: 'base' });
  if (byVersion !== 0) {
    return byVersion;
  }

  return Number(a.Id || 0) - Number(b.Id || 0);
}
