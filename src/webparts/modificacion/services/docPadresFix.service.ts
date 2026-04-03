/* eslint-disable */
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { getAllItems, spGetJson, updateListItem } from './sharepointRest.service';
import { buildDocPadresFixWorkbook, IDocPadresFixReportRow } from '../utils/docPadresFixExcel';
import { openExcelRevisionSession } from '../utils/modificacionExcelHelper';

type LogFn = (message: string) => void;

interface ISolicitudItem {
  Id: number;
  Title?: string;
  NombreDocumento?: string;
  [key: string]: any;
}

interface IExcelInputRow {
  relacionId: number;
  documentoPadreId: number;
  documentoPadreNombre: string;
  documentoHijoId: number;
  documentoHijoNombre: string;
  observaciones: string;
}

function escapeODataValue(value: string): string {
  return String(value || '').replace(/'/g, `''`);
}

function normalizeLookupIds(value: any): number[] {
  if (Array.isArray(value)) {
    return value.map((item) => Number(item)).filter((item) => Number.isFinite(item) && item > 0);
  }

  if (value && Array.isArray(value.results)) {
    return value.results.map((item: any) => Number(item)).filter((item: number) => Number.isFinite(item) && item > 0);
  }

  const single = Number(value);
  return Number.isFinite(single) && single > 0 ? [single] : [];
}

function formatSolicitudName(item: ISolicitudItem | undefined): string {
  if (!item) {
    return '';
  }

  return String(item.Title || item.NombreDocumento || '').trim();
}

function formatSolicitudRefs(ids: number[], solicitudById: Map<number, ISolicitudItem>): string {
  return ids
    .map((id) => {
      const name = formatSolicitudName(solicitudById.get(id));
      return name ? `${id} - ${name}` : String(id);
    })
    .join(' | ');
}

function findColumnIndex(headers: any[], candidates: string[]): number {
  const normalizedCandidates = candidates.map((value) => String(value || '').trim().toLowerCase());
  for (let i = 0; i < headers.length; i++) {
    const current = String(headers[i] || '').trim().toLowerCase();
    if (normalizedCandidates.indexOf(current) !== -1) {
      return i;
    }
  }
  return -1;
}

function parseExcelRows(grid: any[][]): IExcelInputRow[] {
  if (!grid.length) {
    throw new Error('El Excel está vacío.');
  }

  const headers = grid[0] || [];
  const idxRelacionId = findColumnIndex(headers, ['RelacionId']);
  const idxPadreId = findColumnIndex(headers, ['DocumentoPadreId']);
  const idxPadreNombre = findColumnIndex(headers, ['DocumentoPadreNombre']);
  const idxHijoId = findColumnIndex(headers, ['DocumentoHijoId']);
  const idxHijoNombre = findColumnIndex(headers, ['DocumentoHijoNombre']);
  const idxObservaciones = findColumnIndex(headers, ['Observaciones']);

  if (idxRelacionId < 0 || idxPadreId < 0 || idxHijoId < 0) {
    throw new Error('El Excel debe contener las columnas RelacionId, DocumentoPadreId y DocumentoHijoId.');
  }

  const rows: IExcelInputRow[] = [];
  for (let i = 1; i < grid.length; i++) {
    const row = grid[i] || [];
    const relacionId = Number(row[idxRelacionId] || 0);
    const documentoPadreId = Number(row[idxPadreId] || 0);
    const documentoHijoId = Number(row[idxHijoId] || 0);

    if (!relacionId || !documentoPadreId || !documentoHijoId) {
      continue;
    }

    rows.push({
      relacionId,
      documentoPadreId,
      documentoPadreNombre: idxPadreNombre >= 0 ? String(row[idxPadreNombre] || '').trim() : '',
      documentoHijoId,
      documentoHijoNombre: idxHijoNombre >= 0 ? String(row[idxHijoNombre] || '').trim() : '',
      observaciones: idxObservaciones >= 0 ? String(row[idxObservaciones] || '').trim() : ''
    });
  }

  return rows;
}

async function tryGetFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldTitleOrInternalName: string
): Promise<string | null> {
  try {
    const field = await spGetJson<any>(
      context,
      `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(listTitle)}')/fields/getbyinternalnameortitle('${escapeODataValue(fieldTitleOrInternalName)}')?$select=InternalName`
    );
    return String(field?.InternalName || fieldTitleOrInternalName);
  } catch (_error) {
    return null;
  }
}

async function resolveFirstExistingFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  candidates: string[]
): Promise<string> {
  for (let i = 0; i < candidates.length; i++) {
    const found = await tryGetFieldInternalName(context, webUrl, listTitle, candidates[i]);
    if (found) {
      return found;
    }
  }

  throw new Error(`No se encontró el campo esperado en "${listTitle}": ${candidates.join(', ')}`);
}

async function getAllowMultipleValues(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldInternalName: string
): Promise<boolean> {
  const field = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(listTitle)}')/fields/getbyinternalnameortitle('${escapeODataValue(fieldInternalName)}')?$select=AllowMultipleValues,TypeAsString`
  );

  return !!field.AllowMultipleValues || String(field?.TypeAsString || '').toLowerCase().indexOf('multi') !== -1;
}

export async function corregirDocPadresDesdeRelaciones(params: {
  context: WebPartContext;
  excelFile: IFilePickerResult;
  log?: LogFn;
}): Promise<{
  blob: Blob;
  fileName: string;
  processed: number;
  updated: number;
  skipped: number;
  error: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const session = await openExcelRevisionSession(params.excelFile);
  const inputRows = parseExcelRows(session.grid);

  log(`📄 Excel de entrada cargado: ${session.fileName}`);
  log(`📋 Filas válidas detectadas para corrección: ${inputRows.length}`);

  const docPadresField = await resolveFirstExistingFieldInternalName(params.context, webUrl, 'Solicitudes', ['docpadres', 'DocPadres', 'DocumentoPadre']);
  const docPadresFieldId = `${docPadresField}Id`;
  const docPadresIsMulti = await getAllowMultipleValues(params.context, webUrl, 'Solicitudes', docPadresField);

  log(`🔎 Campo DocPadres resuelto: ${docPadresFieldId} | Multiple=${docPadresIsMulti ? 'Si' : 'No'}`);

  const solicitudes = await getAllItems<ISolicitudItem>(
    params.context,
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items?$select=Id,Title,NombreDocumento,${docPadresFieldId}&$top=5000`
  );

  const solicitudById = new Map<number, ISolicitudItem>();
  for (let i = 0; i < solicitudes.length; i++) {
    solicitudById.set(Number(solicitudes[i].Id || 0), solicitudes[i]);
  }

  const grouped = new Map<number, { parentIds: Set<number>; relationIds: number[]; childName: string; parentNames: string[]; observations: string[]; }>();
  for (let i = 0; i < inputRows.length; i++) {
    const row = inputRows[i];
    if (!grouped.has(row.documentoHijoId)) {
      grouped.set(row.documentoHijoId, {
        parentIds: new Set<number>(),
        relationIds: [],
        childName: row.documentoHijoNombre || '',
        parentNames: [],
        observations: []
      });
    }

    const bucket = grouped.get(row.documentoHijoId)!;
    bucket.parentIds.add(row.documentoPadreId);
    bucket.relationIds.push(row.relacionId);
    if (row.documentoPadreNombre) bucket.parentNames.push(row.documentoPadreNombre);
    if (!bucket.childName && row.documentoHijoNombre) bucket.childName = row.documentoHijoNombre;
    if (row.observaciones) bucket.observations.push(row.observaciones);
  }

  const reportRows: IDocPadresFixReportRow[] = [];
  let updated = 0;
  let skipped = 0;
  let error = 0;

  const childIds = Array.from(grouped.keys()).sort((a, b) => a - b);
  for (let i = 0; i < childIds.length; i++) {
    const childId = childIds[i];
    const bucket = grouped.get(childId)!;
    const childSolicitud = solicitudById.get(childId);
    const expectedParentIds = Array.from(bucket.parentIds).filter((id) => solicitudById.has(id)).sort((a, b) => a - b);
    const expectedParentsText = expectedParentIds.length
      ? formatSolicitudRefs(expectedParentIds, solicitudById)
      : bucket.parentNames.join(' | ');
    const relationIdsText = bucket.relationIds.filter((id) => id > 0).join('/');

    if (!childSolicitud) {
      error++;
      reportRows.push({
        SolicitudHijaId: childId,
        SolicitudHijaNombre: bucket.childName,
        RelacionesIds: relationIdsText,
        SolicitudesPadreEsperadas: expectedParentsText,
        DocPadresAntes: '',
        DocPadresDespues: '',
        Estado: 'ERROR',
        Observaciones: 'La solicitud hija no existe en la lista Solicitudes.'
      });
      continue;
    }

    if (!expectedParentIds.length) {
      error++;
      reportRows.push({
        SolicitudHijaId: childId,
        SolicitudHijaNombre: formatSolicitudName(childSolicitud),
        RelacionesIds: relationIdsText,
        SolicitudesPadreEsperadas: expectedParentsText,
        DocPadresAntes: formatSolicitudRefs(normalizeLookupIds(childSolicitud[docPadresFieldId]), solicitudById),
        DocPadresDespues: '',
        Estado: 'ERROR',
        Observaciones: 'No se encontraron solicitudes padre válidas para aplicar con base en el Excel.'
      });
      continue;
    }

    const currentIds = normalizeLookupIds(childSolicitud[docPadresFieldId]);
    const nextIds = Array.from(new Set(currentIds.concat(expectedParentIds))).sort((a, b) => a - b);
    const beforeText = formatSolicitudRefs(currentIds, solicitudById);
    const afterText = formatSolicitudRefs(nextIds, solicitudById);

    if (!docPadresIsMulti && nextIds.length > 1) {
      error++;
      reportRows.push({
        SolicitudHijaId: childId,
        SolicitudHijaNombre: formatSolicitudName(childSolicitud),
        RelacionesIds: relationIdsText,
        SolicitudesPadreEsperadas: expectedParentsText,
        DocPadresAntes: beforeText,
        DocPadresDespues: afterText,
        Estado: 'ERROR',
        Observaciones: 'El campo DocPadres no admite múltiples valores y el Excel indica más de un padre.'
      });
      continue;
    }

    const changed = currentIds.length !== nextIds.length || currentIds.some((id, index) => id !== nextIds[index]);
    if (!changed) {
      skipped++;
      reportRows.push({
        SolicitudHijaId: childId,
        SolicitudHijaNombre: formatSolicitudName(childSolicitud),
        RelacionesIds: relationIdsText,
        SolicitudesPadreEsperadas: expectedParentsText,
        DocPadresAntes: beforeText,
        DocPadresDespues: afterText,
        Estado: 'SKIP',
        Observaciones: 'La solicitud ya tenía los DocPadres informados según el Excel.'
      });
      continue;
    }

    await updateListItem(params.context, webUrl, 'Solicitudes', childId, {
      [docPadresFieldId]: docPadresIsMulti ? nextIds : (nextIds[0] || null)
    });

    updated++;
    log(`✅ DocPadres actualizado | Hijo=${childId} | Antes="${beforeText}" | Despues="${afterText}"`);
    reportRows.push({
      SolicitudHijaId: childId,
      SolicitudHijaNombre: formatSolicitudName(childSolicitud),
      RelacionesIds: relationIdsText,
      SolicitudesPadreEsperadas: expectedParentsText,
      DocPadresAntes: beforeText,
      DocPadresDespues: afterText,
      Estado: 'ACTUALIZADO',
      Observaciones: 'DocPadres completado usando las filas del Excel cargado.'
    });
  }

  const report = buildDocPadresFixWorkbook(reportRows);
  return {
    blob: report.blob,
    fileName: report.fileName,
    processed: childIds.length,
    updated,
    skipped,
    error
  };
}
