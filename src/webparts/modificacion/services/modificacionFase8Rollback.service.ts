/* eslint-disable */
// @ts-nocheck
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { deleteListItem, getAllItems, recycleFile, spGetJson, spPostJson, updateListItem } from './sharepointRest.service';

type LogFn = (s: string) => void;

const PROCESOS_ROOT = '/sites/SistemadeGestionDocumental/Procesos';

export interface IFase8RollbackEntry {
  solicitudId: number;
  nombreDocumento: string;
  oldOriginalUrl: string;
  oldRenamedUrl: string;
  historicoUrl: string;
  oldOriginalMetadata?: any;
  replacementSolicitudId?: number;
  replacementPublishedUrl?: string;
  replacementProcessItemId?: number;
  tempFileUrl?: string;
  updatedChildSolicitudIds?: number[];
  reassignedExistingDiagramIds?: number[];
  newDiagramItemId?: number;
}

function buildMoveCopyBody(webUrl: string, srcFileUrl: string, destFileUrl: string, overwrite: boolean): any {
  const origin = new URL(webUrl).origin;
  const toAbsolute = (value: string): string => `${origin}${value.startsWith('/') ? '' : '/'}${value}`;

  return {
    srcPath: { DecodedUrl: toAbsolute(srcFileUrl) },
    destPath: { DecodedUrl: toAbsolute(destFileUrl) },
    overwrite,
    options: {
      KeepBoth: false,
      ResetAuthorAndCreatedOnCopy: false,
      ShouldBypassSharedLocks: true
    }
  };
}

async function moveFileByPath(
  context: WebPartContext,
  webUrl: string,
  srcFileUrl: string,
  destFileUrl: string,
  overwrite: boolean
): Promise<void> {
  await spPostJson(
    context,
    webUrl,
    `${webUrl}/_api/SP.MoveCopyUtil.MoveFileByPath()`,
    buildMoveCopyBody(webUrl, srcFileUrl, destFileUrl, overwrite),
    'POST'
  );
}

async function updateFileMetadataByPath(context: WebPartContext, webUrl: string, fileUrl: string, payload: any): Promise<void> {
  await spPostJson(
    context,
    webUrl,
    `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${fileUrl.replace(/'/g, "''")}')/ListItemAllFields`,
    payload,
    'MERGE'
  );
}

function normalizeAreaImpactadaForProcesos(value: any): string[] {
  if (Array.isArray(value)) {
    return value.map((item) => String(item || '').trim()).filter(Boolean);
  }

  return String(value || '')
    .split('/')
    .map((part) => part.trim())
    .filter(Boolean);
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

async function getAllowMultipleValues(
  context: WebPartContext,
  webUrl: string,
  listPath: string,
  fieldInternalName: string
): Promise<boolean> {
  const field = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/GetList('${listPath.replace(/'/g, "''")}')/fields/getbyinternalnameortitle('${fieldInternalName.replace(/'/g, "''")}')?$select=AllowMultipleValues`
  );
  return !!field?.AllowMultipleValues;
}

async function getFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldTitleOrInternalName: string
): Promise<string> {
  const field = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${listTitle.replace(/'/g, "''")}')/fields/getbyinternalnameortitle('${fieldTitleOrInternalName.replace(/'/g, "''")}')?$select=InternalName`
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

async function buildProcesosRestorePayload(context: WebPartContext, webUrl: string, oldMetadata: any): Promise<any> {
  const areaImpactada = normalizeAreaImpactadaForProcesos(oldMetadata?.AreaImpactada);
  const documentoPadreIds = normalizeLookupIds(oldMetadata?.DocumentoPadreId);
  const documentoPadreIsMulti = documentoPadreIds.length
    ? await getAllowMultipleValues(context, webUrl, PROCESOS_ROOT, 'DocumentoPadre')
    : false;

  const payload: any = {
    Title: oldMetadata?.Title || '',
    NombreDocumento: oldMetadata?.NombreDocumento || '',
    Tipodedocumento: oldMetadata?.Tipodedocumento || '',
    CategoriaDocumento: oldMetadata?.CategoriaDocumento || '',
    Codigodedocumento: oldMetadata?.Codigodedocumento || '',
    AreaDuena: oldMetadata?.AreaDuena || '',
    AreaImpactada: areaImpactada,
    SolicitudId: Number(oldMetadata?.SolicitudId || 0) || null,
    Clasificaciondeproceso: oldMetadata?.Clasificaciondeproceso || '',
    Macroproceso: oldMetadata?.Macroproceso || '',
    Proceso: oldMetadata?.Proceso || '',
    Subproceso: oldMetadata?.Subproceso || '',
    Resumen: oldMetadata?.Resumen || '',
    FechaDeAprobacion: oldMetadata?.FechaDeAprobacion || null,
    FechaDeVigencia: oldMetadata?.FechaDeVigencia || null,
    InstanciaDeAprobacionId: Number(oldMetadata?.InstanciaDeAprobacionId || 0) || null,
    VersionDocumento: oldMetadata?.VersionDocumento || '',
    Accion: oldMetadata?.Accion || '',
    FechaDePublicacion: oldMetadata?.FechaDePublicacion || null
  };

  if (documentoPadreIds.length) {
    payload.DocumentoPadreId = documentoPadreIsMulti ? documentoPadreIds : documentoPadreIds[0];
  }

  return payload;
}

async function getCurrentProcessFileBySolicitudId(
  context: WebPartContext,
  webUrl: string,
  solicitudId: number
): Promise<{ FileRef: string; FileLeafRef: string; Id: number; } | null> {
  const items = await spGetJson<{ value?: any[] }>(
    context,
    `${webUrl}/_api/web/GetList('${PROCESOS_ROOT.replace(/'/g, "''")}')/items?$select=Id,FileRef,FileLeafRef,SolicitudId&$filter=SolicitudId eq ${solicitudId}&$top=5`
  );
  const row = (items.value || [])[0];
  if (!row) return null;
  return {
    Id: Number(row.Id || 0),
    FileRef: String(row.FileRef || ''),
    FileLeafRef: String(row.FileLeafRef || '')
  };
}

async function restoreRelacionesDocumentosPadre(params: {
  context: WebPartContext;
  webUrl: string;
  oldParentSolicitudId: number;
  newParentSolicitudId: number;
  childSolicitudIds: number[];
  log: LogFn;
}): Promise<void> {
  const childIds = Array.from(new Set((params.childSolicitudIds || []).filter((id) => Number(id) > 0)));
  if (!childIds.length) return;

  const parentField = await getFieldInternalName(params.context, params.webUrl, 'Relaciones Documentos', 'DocumentoPadre');
  const childField = await getFieldInternalName(params.context, params.webUrl, 'Relaciones Documentos', 'DocumentoHijo');
  const parentFieldId = `${parentField}Id`;
  const childFieldId = `${childField}Id`;
  const items = await getAllItems<any>(
    params.context,
    `${params.webUrl}/_api/web/lists/getbytitle('Relaciones Documentos')/items?$select=Id,${parentFieldId},${childFieldId}&$top=5000&$filter=${parentFieldId} eq ${params.newParentSolicitudId}`
  );

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (childIds.indexOf(Number(item[childFieldId] || 0)) === -1) continue;
    await updateListItem(params.context, params.webUrl, 'Relaciones Documentos', item.Id, {
      [parentFieldId]: params.oldParentSolicitudId
    });
  }

  params.log(`🔗 Fase 8 rollback | Relaciones restauradas al padre ${params.oldParentSolicitudId}`);
}

async function restoreChildSolicitudesDocPadres(params: {
  context: WebPartContext;
  webUrl: string;
  oldParentSolicitudId: number;
  newParentSolicitudId: number;
  childSolicitudIds: number[];
  log: LogFn;
}): Promise<void> {
  const childIds = Array.from(new Set((params.childSolicitudIds || []).filter((id) => Number(id) > 0)));
  if (!childIds.length) return;

  const docPadresField = await resolveFirstExistingFieldInternalName(params.context, params.webUrl, 'Solicitudes', ['docpadres', 'DocPadres', 'DocumentoPadre']);
  const docPadresFieldId = `${docPadresField}Id`;
  const docPadresIsMulti = await getAllowMultipleValues(params.context, params.webUrl, 'Solicitudes', docPadresField);
  const items = await getAllItems<any>(
    params.context,
    `${params.webUrl}/_api/web/lists/getbytitle('Solicitudes')/items?$select=Id,${docPadresFieldId}&$top=5000`
  );
  const itemById = new Map<number, any>();
  for (let i = 0; i < items.length; i++) {
    itemById.set(Number(items[i].Id || 0), items[i]);
  }

  for (let i = 0; i < childIds.length; i++) {
    const childId = childIds[i];
    const item = itemById.get(childId);
    if (!item) continue;

    const currentIds = normalizeLookupIds(item[docPadresFieldId]);
    if (!currentIds.length) continue;

    const nextIds = currentIds.map((id) => id === params.newParentSolicitudId ? params.oldParentSolicitudId : id);
    const deduped = Array.from(new Set(nextIds.filter((id) => id > 0)));
    await updateListItem(params.context, params.webUrl, 'Solicitudes', childId, {
      [docPadresFieldId]: docPadresIsMulti ? deduped : (deduped[0] || null)
    });
  }

  params.log(`👨‍👧 Fase 8 rollback | DocPadres restaurados al padre ${params.oldParentSolicitudId}`);
}

async function restoreChildProcessParentReferences(params: {
  context: WebPartContext;
  webUrl: string;
  restoredParentProcessItemId: number;
  replacementParentProcessItemId: number;
  childSolicitudIds: number[];
  log: LogFn;
}): Promise<void> {
  const childIds = Array.from(new Set((params.childSolicitudIds || []).filter((id) => Number(id) > 0)));
  if (!childIds.length) return;

  const documentoPadreIsMulti = await getAllowMultipleValues(params.context, params.webUrl, PROCESOS_ROOT, 'DocumentoPadre');

  for (let i = 0; i < childIds.length; i++) {
    const childFile = await getCurrentProcessFileBySolicitudId(params.context, params.webUrl, childIds[i]);
    if (!childFile?.FileRef) continue;

    const metadata = await spGetJson<any>(
      params.context,
      `${params.webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${childFile.FileRef.replace(/'/g, "''")}')/ListItemAllFields?$select=DocumentoPadreId`
    );
    const currentIds = normalizeLookupIds(metadata?.DocumentoPadreId);
    if (!currentIds.length) continue;

    const replaced = currentIds.map((id) => id === params.replacementParentProcessItemId ? params.restoredParentProcessItemId : id);
    const deduped = Array.from(new Set(replaced.filter((id) => id > 0)));
    await updateFileMetadataByPath(params.context, params.webUrl, childFile.FileRef, {
      DocumentoPadreId: documentoPadreIsMulti ? deduped : (deduped[0] || null)
    });
  }

  params.log('👨‍👧 Fase 8 rollback | Referencias en Procesos restauradas al padre vigente anterior');
}

async function restoreExistingDiagramasSolicitud(params: {
  context: WebPartContext;
  webUrl: string;
  oldParentSolicitudId: number;
  diagramIds: number[];
  log: LogFn;
}): Promise<void> {
  const ids = Array.from(new Set((params.diagramIds || []).filter((id) => Number(id) > 0)));
  if (!ids.length) return;

  const solicitudField = await getFieldInternalName(params.context, params.webUrl, 'Diagramas de Flujo', 'Solicitud');
  const solicitudFieldId = `${solicitudField}Id`;
  const solicitudIsMulti = await getAllowMultipleValues(params.context, params.webUrl, 'Diagramas de Flujo', solicitudField);

  for (let i = 0; i < ids.length; i++) {
    await updateListItem(params.context, params.webUrl, 'Diagramas de Flujo', ids[i], {
      [solicitudFieldId]: solicitudIsMulti ? [params.oldParentSolicitudId] : params.oldParentSolicitudId
    });
  }

  params.log(`🧭 Fase 8 rollback | Diagramas restaurados a la solicitud ${params.oldParentSolicitudId}`);
}

export async function rollbackModificacionFase8(params: {
  context: WebPartContext;
  webUrl: string;
  entries: IFase8RollbackEntry[];
  log?: LogFn;
}): Promise<void> {
  const log = params.log || (() => undefined);
  const entries = Array.isArray(params.entries) ? params.entries : [];

  for (let i = entries.length - 1; i >= 0; i--) {
    const entry = entries[i];

    try {
      if (entry.historicoUrl && entry.oldOriginalUrl) {
        await moveFileByPath(params.context, params.webUrl, entry.historicoUrl, entry.oldOriginalUrl, true);
        log(`♻️ Fase 8 rollback | Documento restaurado desde Históricos: ${entry.oldOriginalUrl}`);

        if (entry.oldOriginalMetadata) {
          await updateFileMetadataByPath(
            params.context,
            params.webUrl,
            entry.oldOriginalUrl,
            await buildProcesosRestorePayload(params.context, params.webUrl, entry.oldOriginalMetadata)
          );
          log(`🧾 Fase 8 rollback | Metadata restaurada: ${entry.oldOriginalUrl}`);
        }
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 8 rollback | No se pudo restaurar el documento ${entry.oldOriginalUrl} -> ${message}`);
    }

    try {
      const restoredParentFile = await getCurrentProcessFileBySolicitudId(params.context, params.webUrl, entry.solicitudId);
      if (restoredParentFile?.Id && entry.replacementProcessItemId && (entry.updatedChildSolicitudIds || []).length) {
        await restoreChildProcessParentReferences({
          context: params.context,
          webUrl: params.webUrl,
          restoredParentProcessItemId: restoredParentFile.Id,
          replacementParentProcessItemId: entry.replacementProcessItemId,
          childSolicitudIds: entry.updatedChildSolicitudIds || [],
          log
        });
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 8 rollback | No se pudieron restaurar las referencias en Procesos -> ${message}`);
    }

    try {
      if (entry.replacementSolicitudId && (entry.updatedChildSolicitudIds || []).length) {
        await restoreChildSolicitudesDocPadres({
          context: params.context,
          webUrl: params.webUrl,
          oldParentSolicitudId: entry.solicitudId,
          newParentSolicitudId: entry.replacementSolicitudId,
          childSolicitudIds: entry.updatedChildSolicitudIds || [],
          log
        });

        await restoreRelacionesDocumentosPadre({
          context: params.context,
          webUrl: params.webUrl,
          oldParentSolicitudId: entry.solicitudId,
          newParentSolicitudId: entry.replacementSolicitudId,
          childSolicitudIds: entry.updatedChildSolicitudIds || [],
          log
        });
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 8 rollback | No se pudieron restaurar relaciones/hijos -> ${message}`);
    }

    try {
      if ((entry.reassignedExistingDiagramIds || []).length) {
        await restoreExistingDiagramasSolicitud({
          context: params.context,
          webUrl: params.webUrl,
          oldParentSolicitudId: entry.solicitudId,
          diagramIds: entry.reassignedExistingDiagramIds || [],
          log
        });
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 8 rollback | No se pudieron restaurar los diagramas existentes -> ${message}`);
    }

    if (entry.newDiagramItemId) {
      try {
        await deleteListItem(params.context, params.webUrl, 'Diagramas de Flujo', entry.newDiagramItemId);
        log(`🗑️ Fase 8 rollback | Nuevo diagrama eliminado: ${entry.newDiagramItemId}`);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        log(`⚠️ Fase 8 rollback | No se pudo eliminar el nuevo diagrama ${entry.newDiagramItemId} -> ${message}`);
      }
    }

    if (entry.replacementPublishedUrl) {
      const sameAsRestored =
        String(entry.replacementPublishedUrl || '').trim().toLowerCase() === String(entry.oldOriginalUrl || '').trim().toLowerCase();

      if (sameAsRestored) {
        log(`ℹ️ Fase 8 rollback | Publicación nueva no reciclada porque coincide con la ruta restaurada: ${entry.replacementPublishedUrl}`);
      } else {
        try {
          await recycleFile(params.context, params.webUrl, entry.replacementPublishedUrl);
          log(`🗑️ Fase 8 rollback | Publicación nueva reciclada: ${entry.replacementPublishedUrl}`);
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error);
          log(`⚠️ Fase 8 rollback | No se pudo reciclar la publicación nueva -> ${message}`);
        }
      }
    }

    if (entry.tempFileUrl) {
      try {
        await recycleFile(params.context, params.webUrl, entry.tempFileUrl);
        log(`🗑️ Fase 8 rollback | TEMP reciclado: ${entry.tempFileUrl}`);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        log(`⚠️ Fase 8 rollback | No se pudo reciclar el TEMP -> ${message}`);
      }
    }

    if (entry.replacementSolicitudId) {
      try {
        await deleteListItem(params.context, params.webUrl, 'Solicitudes', entry.replacementSolicitudId);
        log(`🗑️ Fase 8 rollback | Nueva solicitud eliminada: ${entry.replacementSolicitudId}`);
      } catch (_deleteError) {
        try {
          await updateListItem(params.context, params.webUrl, 'Solicitudes', entry.replacementSolicitudId, {
            EsVersionActualDocumento: false
          });
        } catch (_updateError) {
          // noop
        }
      }
    }

    try {
      await updateListItem(params.context, params.webUrl, 'Solicitudes', entry.solicitudId, {
        EsVersionActualDocumento: true
      });
      log('♻️ Fase 8 rollback | Solicitud restaurada: ' + entry.solicitudId + (entry.replacementSolicitudId ? ' | Nueva solicitud eliminada: ' + entry.replacementSolicitudId : ''));
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 8 rollback | No se pudo restaurar la solicitud ${entry.solicitudId} -> ${message}`);
    }
  }
}
