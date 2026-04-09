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
  Estado?: {
    Title?: string;
  };
  Author?: {
    Id?: number;
    Title?: string;
    EMail?: string;
  };
}

interface IRelacionDocumentoItem {
  DocumentoPadreId?: number;
  DocumentoHijoId?: number;
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

  log('🔎 Consultando la lista Solicitudes para detectar CodigoDocumento duplicado...');
  const solicitudes = await obtenerSolicitudes(params.context, webUrl);
  log(`📋 Solicitudes leidas: ${solicitudes.length}`);

  log('🔎 Consultando Relaciones Documentos para armar padres e hijos...');
  const relaciones = await obtenerRelaciones(params.context, webUrl);
  log(`🔗 Relaciones leidas: ${relaciones.length}`);

  const grouped = new Map<string, ISolicitudItem[]>();
  const relationMaps = buildRelationMaps(relaciones);

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

    const distinctTitles = buildDistinctTitles(items);
    if (distinctTitles.length < 2) {
      return;
    }

    duplicatedGroups++;

    const sortedItems = items.slice().sort((a, b) => Number(a.Id || 0) - Number(b.Id || 0));

    for (let i = 0; i < sortedItems.length; i++) {
      const item = sortedItems[i];
      if (!isTruthyField(item.EsVersionActualDocumento)) {
        nonCurrentRows++;
      }

      duplicateRows.push(buildReportRow(item, sortedItems.length, distinctTitles.length, relationMaps));
    }
  });

  duplicateRows.sort(compareReportRows);

  if (duplicateRows.length) {
    log(`⚠️ Grupos duplicados detectados: ${duplicatedGroups}`);
    log(`⚠️ Solicitudes duplicadas detectadas: ${duplicateRows.length}`);
    log(`ℹ️ Solicitudes no vigentes dentro de duplicados: ${nonCurrentRows}`);
  } else {
    log('✅ No se encontraron solicitudes con CodigoDocumento duplicado y titulos distintos.');
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
    `EstadoId,Estado/Title,FechaDeAprobacionSolicitud,FechadeVigencia,FechaDePublicacionSolicitud,Created,Modified,` +
    `Author/Id,Author/Title,Author/EMail` +
    `&$expand=Estado,Author` +
    `&$top=5000`;

  return getAllItems<ISolicitudItem>(context, url);
}

async function obtenerRelaciones(context: WebPartContext, webUrl: string): Promise<IRelacionDocumentoItem[]> {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('Relaciones Documentos')/items` +
    `?$select=DocumentoPadreId,DocumentoHijoId` +
    `&$top=5000`;

  return getAllItems<IRelacionDocumentoItem>(context, url);
}

function buildDuplicateKey(item: ISolicitudItem): string {
  return normalizeKeyPart(item.CodigoDocumento);
}

function normalizeKeyPart(value: any): string {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function buildDistinctTitles(items: ISolicitudItem[]): string[] {
  const map = new Map<string, string>();

  for (let i = 0; i < items.length; i++) {
    const title = String(items[i].Title || '').trim();
    const key = normalizeKeyPart(title);
    if (!key || map.has(key)) {
      continue;
    }

    map.set(key, title);
  }

  return Array.from(map.values());
}

function buildRelationMaps(relaciones: IRelacionDocumentoItem[]): {
  parentIdsByChildId: Map<number, number[]>;
  childIdsByParentId: Map<number, number[]>;
} {
  const parentIdsByChildId = new Map<number, number[]>();
  const childIdsByParentId = new Map<number, number[]>();

  for (let i = 0; i < relaciones.length; i++) {
    const relacion = relaciones[i];
    const parentId = Number(relacion.DocumentoPadreId || 0);
    const childId = Number(relacion.DocumentoHijoId || 0);

    if (parentId && childId) {
      pushUniqueNumber(parentIdsByChildId, childId, parentId);
      pushUniqueNumber(childIdsByParentId, parentId, childId);
    }
  }

  sortMapValues(parentIdsByChildId);
  sortMapValues(childIdsByParentId);

  return {
    parentIdsByChildId,
    childIdsByParentId
  };
}

function pushUniqueNumber(map: Map<number, number[]>, key: number, value: number): void {
  if (!map.has(key)) {
    map.set(key, []);
  }

  const values = map.get(key)!;
  if (values.indexOf(value) === -1) {
    values.push(value);
  }
}

function sortMapValues(map: Map<number, number[]>): void {
  map.forEach((values) => values.sort((a, b) => a - b));
}

function formatIdList(values: number[]): string {
  return values.length ? values.join(' | ') : '';
}

function buildReportRow(
  item: ISolicitudItem,
  totalCoincidencias: number,
  totalTitulosDistintos: number,
  relationMaps: {
    parentIdsByChildId: Map<number, number[]>;
    childIdsByParentId: Map<number, number[]>;
  }
): ISolicitudDuplicadaReportRow {
  const itemId = Number(item.Id || 0);
  return {
    ClaveCodigoDocumento: String(item.CodigoDocumento || '').trim(),
    TotalCoincidencias: totalCoincidencias,
    TotalTitulosDistintos: totalTitulosDistintos,
    Id: itemId,
    Title: String(item.Title || '').trim(),
    NombreDocumento: String(item.NombreDocumento || '').trim(),
    CodigoDocumento: String(item.CodigoDocumento || '').trim(),
    VersionDocumento: String(item.VersionDocumento || '').trim(),
    EsVersionActualDocumento: formatBooleanField(item.EsVersionActualDocumento),
    DocumentosApoyo: formatBooleanField(item.DocumentosApoyo),
    Estado: String(item.Estado?.Title || '').trim(),
    EstadoId: Number(item.EstadoId || 0) || '',
    CreadoPor: String(item.Author?.Title || '').trim(),
    CreadoPorEmail: String(item.Author?.EMail || '').trim(),
    DocumentosPadreIds: formatIdList(relationMaps.parentIdsByChildId.get(itemId) || []),
    DocumentosHijosIds: formatIdList(relationMaps.childIdsByParentId.get(itemId) || []),
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
  const byDocument = String(a.ClaveCodigoDocumento || '').localeCompare(String(b.ClaveCodigoDocumento || ''), 'es', { sensitivity: 'base' });
  if (byDocument !== 0) {
    return byDocument;
  }

  const byTitle = String(a.Title || '').localeCompare(String(b.Title || ''), 'es', { sensitivity: 'base' });
  if (byTitle !== 0) {
    return byTitle;
  }

  const byVersion = String(a.VersionDocumento || '').localeCompare(String(b.VersionDocumento || ''), 'es', { sensitivity: 'base' });
  if (byVersion !== 0) {
    return byVersion;
  }

  return Number(a.Id || 0) - Number(b.Id || 0);
}
