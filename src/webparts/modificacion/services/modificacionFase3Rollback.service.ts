/* eslint-disable */
// @ts-nocheck
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { deleteListItem, getAllItems, spGetJson, updateListItem } from './sharepointRest.service';
import { IFase3RollbackEntry } from './modificacionFase3Padres.service';

type LogFn = (s: string) => void;

async function getFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldTitleOrInternalName: string
): Promise<string> {
  const field = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${listTitle}')/fields/getbyinternalnameortitle('${fieldTitleOrInternalName}')?$select=InternalName`
  );
  return String(field?.InternalName || fieldTitleOrInternalName);
}

async function tryGetFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldTitleOrInternalName: string
): Promise<string | null> {
  try {
    return await getFieldInternalName(context, webUrl, listTitle, fieldTitleOrInternalName);
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
    const resolved = await tryGetFieldInternalName(context, webUrl, listTitle, candidates[i]);
    if (resolved) return resolved;
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
    `${webUrl}/_api/web/lists/getbytitle('${listTitle}')/fields/getbyinternalnameortitle('${fieldInternalName}')?$select=AllowMultipleValues`
  );
  return !!field.AllowMultipleValues;
}

function toLookupIdArray(value: any): number[] {
  if (Array.isArray(value)) {
    return value.map((x) => Number(x)).filter((x) => Number.isFinite(x) && x > 0);
  }

  if (value && Array.isArray(value.results)) {
    return value.results.map((x: any) => Number(x)).filter((x: number) => Number.isFinite(x) && x > 0);
  }

  const single = Number(value);
  return Number.isFinite(single) && single > 0 ? [single] : [];
}

async function rollbackRelacionesDocumentosPadre(
  context: WebPartContext,
  webUrl: string,
  newParentSolicitudId: number,
  oldParentSolicitudId: number,
  childIds: number[],
  log: LogFn
): Promise<number> {
  const wanted = new Set((childIds || []).filter((id) => Number(id) > 0));
  if (!wanted.size) return 0;

  const parentField = await getFieldInternalName(context, webUrl, 'Relaciones Documentos', 'DocumentoPadre');
  const childField = await getFieldInternalName(context, webUrl, 'Relaciones Documentos', 'DocumentoHijo');
  const parentFieldId = `${parentField}Id`;
  const childFieldId = `${childField}Id`;

  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Relaciones Documentos')/items?$select=Id,${parentFieldId},${childFieldId}&$top=5000&$filter=${parentFieldId} eq ${newParentSolicitudId}`
  );

  let updated = 0;
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (!wanted.has(Number(item[childFieldId] || 0))) continue;
    await updateListItem(context, webUrl, 'Relaciones Documentos', item.Id, {
      [parentFieldId]: oldParentSolicitudId
    });
    updated++;
  }

  log(`🔁 Fase 3 rollback | Relaciones restauradas=${updated} | PadreAnterior=${oldParentSolicitudId}`);
  return updated;
}

async function rollbackDiagramasSolicitud(
  context: WebPartContext,
  webUrl: string,
  diagramIds: number[],
  oldParentSolicitudId: number,
  log: LogFn
): Promise<number> {
  const ids = Array.from(new Set((diagramIds || []).filter((id) => Number(id) > 0)));
  const solicitudField = await getFieldInternalName(context, webUrl, 'Diagramas de Flujo', 'Solicitud');
  const solicitudFieldId = `${solicitudField}Id`;
  const solicitudIsMulti = await getAllowMultipleValues(context, webUrl, 'Diagramas de Flujo', solicitudField);
  let updated = 0;

  for (let i = 0; i < ids.length; i++) {
    await updateListItem(context, webUrl, 'Diagramas de Flujo', ids[i], {
      [solicitudFieldId]: solicitudIsMulti ? [oldParentSolicitudId] : oldParentSolicitudId
    });
    updated++;
  }

  log(`🔁 Fase 3 rollback | Diagramas restaurados=${updated} | SolicitudAnterior=${oldParentSolicitudId}`);
  return updated;
}

async function rollbackChildSolicitudesDocPadres(
  context: WebPartContext,
  webUrl: string,
  childSolicitudIds: number[],
  newParentSolicitudId: number,
  oldParentSolicitudId: number,
  log: LogFn
): Promise<number> {
  const childIds = Array.from(new Set((childSolicitudIds || []).filter((id) => Number(id) > 0)));
  if (!childIds.length) return 0;

  const docPadresField = await resolveFirstExistingFieldInternalName(
    context,
    webUrl,
    'Solicitudes',
    ['docpadres', 'DocPadres', 'DocumentoPadre']
  );
  const docPadresFieldId = `${docPadresField}Id`;
  const docPadresIsMulti = await getAllowMultipleValues(context, webUrl, 'Solicitudes', docPadresField);

  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items?$select=Id,${docPadresFieldId}&$top=5000`
  );
  const itemById = new Map<number, any>();
  for (let i = 0; i < items.length; i++) {
    itemById.set(Number(items[i].Id || 0), items[i]);
  }

  let updated = 0;
  for (let i = 0; i < childIds.length; i++) {
    const item = itemById.get(childIds[i]);
    if (!item) continue;

    const currentIds = toLookupIdArray(item[docPadresFieldId]);
    if (!currentIds.length) continue;

    let changed = false;
    const nextIds = currentIds.map((id) => {
      if (id === newParentSolicitudId) {
        changed = true;
        return oldParentSolicitudId;
      }
      return id;
    });

    if (!changed) continue;

    const deduped = Array.from(new Set(nextIds.filter((id) => id > 0)));
    await updateListItem(context, webUrl, 'Solicitudes', childIds[i], {
      [docPadresFieldId]: docPadresIsMulti ? deduped : (deduped[0] || null)
    });
    updated++;
  }

  log(`🔁 Fase 3 rollback | DocPadres restaurados=${updated} | PadreAnterior=${oldParentSolicitudId}`);
  return updated;
}

export async function rollbackModificacionFase3(params: {
  context: WebPartContext;
  webUrl: string;
  entries: IFase3RollbackEntry[];
  log?: LogFn;
}): Promise<void> {
  const log = params.log || (() => undefined);
  const entries = Array.isArray(params.entries) ? params.entries : [];

  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];

    try {
      await rollbackRelacionesDocumentosPadre(
        params.context,
        params.webUrl,
        entry.solicitudNuevaId,
        entry.solicitudOrigenId,
        entry.hijosIds,
        log
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 3 rollback | No se pudieron restaurar relaciones para "${entry.nombreDocumento}" -> ${message}`);
    }

    try {
      await rollbackDiagramasSolicitud(
        params.context,
        params.webUrl,
        entry.diagramasIds,
        entry.solicitudOrigenId,
        log
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 3 rollback | No se pudieron restaurar diagramas para "${entry.nombreDocumento}" -> ${message}`);
    }

    try {
      await rollbackChildSolicitudesDocPadres(
        params.context,
        params.webUrl,
        entry.hijosIds,
        entry.solicitudNuevaId,
        entry.solicitudOrigenId,
        log
      );
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 3 rollback | No se pudo restaurar DocPadres para "${entry.nombreDocumento}" -> ${message}`);
    }

    try {
      await updateListItem(params.context, params.webUrl, 'Solicitudes', entry.solicitudOrigenId, {
        EsVersionActualDocumento: true
      });
      log(`♻️ Fase 3 rollback | Vigencia restaurada | Antigua=${entry.solicitudOrigenId}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 3 rollback | No se pudo restaurar vigencia para "${entry.nombreDocumento}" -> ${message}`);
    }

    try {
      await deleteListItem(params.context, params.webUrl, 'Solicitudes', entry.solicitudNuevaId);
      log(`🗑️ Fase 3 rollback | Nueva solicitud eliminada: ${entry.solicitudNuevaId}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 3 rollback | No se pudo eliminar la nueva solicitud ${entry.solicitudNuevaId} -> ${message}`);
    }
  }
}
