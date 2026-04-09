/* eslint-disable */
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { getAllItems } from './sharepointRest.service';
import {
  buildSolicitudesNoAntonioWorkbook,
  ISolicitudNoAntonioReportRow
} from '../utils/solicitudesNoAntonioExcel';

type LogFn = (message: string) => void;

const ANTONIO_NAME = 'Antonio Sánchez Panta';

interface ISolicitudItem {
  Id: number;
  Title?: string;
  NombreDocumento?: string;
  CodigoDocumento?: string;
  VersionDocumento?: string;
  EsVersionActualDocumento?: boolean | string | number;
  EstadoId?: number;
  Created?: string;
  Modified?: string;
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

export async function exportarSolicitudesNoCreadasPorAntonio(params: {
  context: WebPartContext;
  log?: LogFn;
}): Promise<{
  blob: Blob;
  fileName: string;
  totalSolicitudes: number;
  totalFiltradas: number;
  totalExcluidasAntonio: number;
  conHijos: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;

  log('🔎 Consultando todas las Solicitudes...');
  const solicitudes = await obtenerSolicitudes(params.context, webUrl);
  log(`📋 Solicitudes leidas: ${solicitudes.length}`);

  log('🔎 Consultando Relaciones Documentos para identificar hijos...');
  const relaciones = await obtenerRelaciones(params.context, webUrl);
  log(`🔗 Relaciones leidas: ${relaciones.length}`);

  const childrenByParentId = buildChildrenMap(relaciones);
  const rows: ISolicitudNoAntonioReportRow[] = [];
  let totalExcluidasAntonio = 0;
  let conHijos = 0;

  for (let i = 0; i < solicitudes.length; i++) {
    const solicitud = solicitudes[i];
    const creadoPor = String(solicitud.Author?.Title || '').trim();

    if (isAntonioName(creadoPor)) {
      totalExcluidasAntonio++;
      continue;
    }

    const childIds = childrenByParentId.get(Number(solicitud.Id || 0)) || [];
    if (childIds.length) {
      conHijos++;
    }

    rows.push({
      Id: Number(solicitud.Id || 0),
      Title: String(solicitud.Title || '').trim(),
      NombreDocumento: String(solicitud.NombreDocumento || '').trim(),
      CodigoDocumento: String(solicitud.CodigoDocumento || '').trim(),
      VersionDocumento: String(solicitud.VersionDocumento || '').trim(),
      EsVersionActualDocumento: formatBooleanField(solicitud.EsVersionActualDocumento),
      EstadoId: Number(solicitud.EstadoId || 0) || '',
      CreadoPor: creadoPor,
      CreadoPorEmail: String(solicitud.Author?.EMail || '').trim(),
      Created: String(solicitud.Created || '').trim(),
      Modified: String(solicitud.Modified || '').trim(),
      TieneDocumentosHijos: childIds.length ? 'Si' : 'No',
      TotalDocumentosHijos: childIds.length,
      DocumentosHijosIds: childIds.join(' | ')
    });
  }

  rows.sort((a, b) => Number(a.Id || 0) - Number(b.Id || 0));

  log(`🚫 Solicitudes excluidas por creador "${ANTONIO_NAME}": ${totalExcluidasAntonio}`);
  log(`✅ Solicitudes incluidas en el reporte: ${rows.length}`);
  log(`📎 Solicitudes incluidas con hijos: ${conHijos}`);

  const report = buildSolicitudesNoAntonioWorkbook(rows);

  return {
    blob: report.blob,
    fileName: report.fileName,
    totalSolicitudes: solicitudes.length,
    totalFiltradas: rows.length,
    totalExcluidasAntonio,
    conHijos
  };
}

async function obtenerSolicitudes(context: WebPartContext, webUrl: string): Promise<ISolicitudItem[]> {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items` +
    `?$select=Id,Title,NombreDocumento,CodigoDocumento,VersionDocumento,EsVersionActualDocumento,EstadoId,Created,Modified,Author/Id,Author/Title,Author/EMail` +
    `&$expand=Author` +
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

function buildChildrenMap(relaciones: IRelacionDocumentoItem[]): Map<number, number[]> {
  const map = new Map<number, number[]>();

  for (let i = 0; i < relaciones.length; i++) {
    const relacion = relaciones[i];
    const parentId = Number(relacion.DocumentoPadreId || 0);
    const childId = Number(relacion.DocumentoHijoId || 0);

    if (!parentId || !childId) {
      continue;
    }

    if (!map.has(parentId)) {
      map.set(parentId, []);
    }

    const current = map.get(parentId)!;
    if (current.indexOf(childId) === -1) {
      current.push(childId);
    }
  }

  map.forEach((values) => values.sort((a, b) => a - b));
  return map;
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

function isAntonioName(value: string): boolean {
  return normalizeText(value) === normalizeText(ANTONIO_NAME);
}

function normalizeText(value: string): string {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}
