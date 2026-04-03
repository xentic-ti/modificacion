/* eslint-disable */
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { getAllItems, spGetJson } from './sharepointRest.service';
import {
  buildRelacionesDocumentosAuditWorkbook,
  IRelacionesDocumentosAuditRow,
  ISolicitudHijaSinDocPadreRow
} from '../utils/relacionesDocumentosAuditExcel';

type LogFn = (message: string) => void;

interface ISolicitudItem {
  Id: number;
  Title?: string;
  NombreDocumento?: string;
  [key: string]: any;
}

interface IRelacionDocumentoItem {
  Id: number;
  Title?: string;
  DocumentoPadreId?: number;
  DocumentoHijoId?: number;
  DocumentoPadre?: {
    Id?: number;
    Title?: string;
  };
  DocumentoHijo?: {
    Id?: number;
    Title?: string;
  };
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

function buildSolicitudName(item: ISolicitudItem | undefined): string {
  if (!item) {
    return '';
  }

  return String(item.Title || item.NombreDocumento || '').trim();
}

function formatSolicitudRefs(ids: number[], solicitudById: Map<number, ISolicitudItem>): string {
  if (!ids.length) {
    return '';
  }

  return ids
    .map((id) => {
      const name = buildSolicitudName(solicitudById.get(id));
      return name ? `${id} - ${name}` : String(id);
    })
    .join(' | ');
}

export async function auditarRelacionesDocumentos(params: {
  context: WebPartContext;
  log?: LogFn;
}): Promise<{
  blob: Blob;
  fileName: string;
  totalRelaciones: number;
  relacionesConFaltantes: number;
  hijosSinDocPadre: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;

  log('🔎 Consultando Relaciones Documentos...');
  const relaciones = await getAllItems<IRelacionDocumentoItem>(
    params.context,
    `${webUrl}/_api/web/lists/getbytitle('Relaciones Documentos')/items` +
      `?$select=Id,Title,DocumentoPadreId,DocumentoHijoId,DocumentoPadre/Id,DocumentoPadre/Title,DocumentoHijo/Id,DocumentoHijo/Title` +
      `&$expand=DocumentoPadre,DocumentoHijo` +
      `&$top=5000`
  );
  log(`🔗 Relaciones leidas: ${relaciones.length}`);

  const docPadresField = await resolveFirstExistingFieldInternalName(params.context, webUrl, 'Solicitudes', ['docpadres', 'DocPadres', 'DocumentoPadre']);
  const docPadresFieldId = `${docPadresField}Id`;

  log(`🔎 Consultando Solicitudes y campo ${docPadresFieldId}...`);
  const solicitudes = await getAllItems<ISolicitudItem>(
    params.context,
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items` +
      `?$select=Id,Title,NombreDocumento,${docPadresFieldId}` +
      `&$top=5000`
  );
  log(`📋 Solicitudes leidas: ${solicitudes.length}`);

  const solicitudById = new Map<number, ISolicitudItem>();
  for (let i = 0; i < solicitudes.length; i++) {
    const solicitud = solicitudes[i];
    solicitudById.set(Number(solicitud.Id || 0), solicitud);
  }

  const auditRows: IRelacionesDocumentosAuditRow[] = [];
  const missingDocPadreRows: ISolicitudHijaSinDocPadreRow[] = [];
  let relacionesConFaltantes = 0;

  for (let i = 0; i < relaciones.length; i++) {
    const relacion = relaciones[i];
    const parentId = Number(relacion.DocumentoPadreId || relacion.DocumentoPadre?.Id || 0);
    const childId = Number(relacion.DocumentoHijoId || relacion.DocumentoHijo?.Id || 0);
    const parentSolicitud = solicitudById.get(parentId);
    const childSolicitud = solicitudById.get(childId);
    const parentName = String(relacion.DocumentoPadre?.Title || buildSolicitudName(parentSolicitud) || '').trim();
    const childName = String(relacion.DocumentoHijo?.Title || buildSolicitudName(childSolicitud) || '').trim();
    const observations: string[] = [];

    if (!parentId) {
      observations.push('Falta DocumentoPadreId en la relacion.');
    } else if (!parentSolicitud) {
      observations.push(`No existe la solicitud padre ${parentId} en la lista Solicitudes.`);
    }

    if (!childId) {
      observations.push('Falta DocumentoHijoId en la relacion.');
    } else if (!childSolicitud) {
      observations.push(`No existe la solicitud hija ${childId} en la lista Solicitudes.`);
    }

    const childDocPadresIds = childSolicitud ? normalizeLookupIds(childSolicitud[docPadresFieldId]) : [];
    const childDocPadresText = formatSolicitudRefs(childDocPadresIds, solicitudById);

    if (parentId && childSolicitud) {
      if (!childDocPadresIds.length) {
        observations.push('La solicitud hija no tiene DocPadres.');
        missingDocPadreRows.push({
          RelacionId: Number(relacion.Id || 0) || '',
          SolicitudHijaId: childId,
          SolicitudHijaNombre: buildSolicitudName(childSolicitud),
          SolicitudPadreEsperadoId: parentId || '',
          SolicitudPadreEsperadoNombre: parentName,
          HijoDocPadresActual: childDocPadresText,
          EstadoDocPadres: 'VACIO',
          Observaciones: 'Existe relacion padre/hijo pero la solicitud hija no tiene DocPadres informado.'
        });
      } else if (childDocPadresIds.indexOf(parentId) === -1) {
        observations.push('La solicitud hija no referencia al padre esperado en DocPadres.');
        missingDocPadreRows.push({
          RelacionId: Number(relacion.Id || 0) || '',
          SolicitudHijaId: childId,
          SolicitudHijaNombre: buildSolicitudName(childSolicitud),
          SolicitudPadreEsperadoId: parentId || '',
          SolicitudPadreEsperadoNombre: parentName,
          HijoDocPadresActual: childDocPadresText,
          EstadoDocPadres: 'INCONSISTENTE',
          Observaciones: 'La relacion apunta a un padre, pero DocPadres de la solicitud hija contiene otros valores.'
        });
      }
    }

    if (observations.length) {
      relacionesConFaltantes++;
    }

    auditRows.push({
      RelacionId: Number(relacion.Id || 0) || '',
      RelacionTitle: String(relacion.Title || '').trim(),
      DocumentoPadreId: parentId || '',
      DocumentoPadreNombre: parentName,
      DocumentoHijoId: childId || '',
      DocumentoHijoNombre: childName,
      EstadoRelacion: observations.length ? 'OBSERVADA' : 'OK',
      Observaciones: observations.join(' '),
      HijoDocPadresActual: childDocPadresText
    });
  }

  auditRows.sort((a, b) => Number(a.RelacionId || 0) - Number(b.RelacionId || 0));
  missingDocPadreRows.sort((a, b) => Number(a.SolicitudHijaId || 0) - Number(b.SolicitudHijaId || 0));

  log(`⚠️ Relaciones con observaciones: ${relacionesConFaltantes}`);
  log(`⚠️ Solicitudes hijas con DocPadres faltante o inconsistente: ${missingDocPadreRows.length}`);

  const report = buildRelacionesDocumentosAuditWorkbook({
    auditRows,
    missingDocPadreRows
  });

  return {
    blob: report.blob,
    fileName: report.fileName,
    totalRelaciones: relaciones.length,
    relacionesConFaltantes,
    hijosSinDocPadre: missingDocPadreRows.length
  };
}
