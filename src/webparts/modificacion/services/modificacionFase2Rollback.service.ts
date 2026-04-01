/* eslint-disable */
// @ts-nocheck
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { recycleFile, spGetJson, spPostJson } from './sharepointRest.service';
import { IFase2RollbackEntry } from './modificacionFase2Publicacion.service';

type LogFn = (s: string) => void;
const PROCESOS_ROOT = '/sites/SistemadeGestionDocumental/Procesos';

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

async function buildProcesosRestorePayload(context: WebPartContext, webUrl: string, oldMetadata: any): Promise<any> {
  const areaImpactada = normalizeAreaImpactadaForProcesos(oldMetadata?.AreaImpactada);
  const documentoPadreIds = normalizeLookupIds(oldMetadata?.DocumentoPadreId);
  const documentoPadreIsMulti = documentoPadreIds.length
    ? await getAllowMultipleValues(context, webUrl, PROCESOS_ROOT, 'DocumentoPadre')
    : false;

  const payload: any = {
    Clasificaciondeproceso: oldMetadata?.Clasificaciondeproceso || '',
    AreaDuena: oldMetadata?.AreaDuena || '',
    VersionDocumento: oldMetadata?.VersionDocumento || '',
    AreaImpactada: areaImpactada,
    Macroproceso: oldMetadata?.Macroproceso || '',
    Proceso: oldMetadata?.Proceso || '',
    Subproceso: oldMetadata?.Subproceso || '',
    Tipodedocumento: oldMetadata?.Tipodedocumento || '',
    SolicitudId: Number(oldMetadata?.SolicitudId || 0) || null,
    Codigodedocumento: oldMetadata?.Codigodedocumento || '',
    Resumen: oldMetadata?.Resumen || '',
    CategoriaDocumento: oldMetadata?.CategoriaDocumento || '',
    FechaDeAprobacion: oldMetadata?.FechaDeAprobacion || null,
    FechaDePublicacion: oldMetadata?.FechaDePublicacion || null,
    FechaDeVigencia: oldMetadata?.FechaDeVigencia || null,
    InstanciaDeAprobacionId: Number(oldMetadata?.InstanciaDeAprobacionId || 0) || null,
    Accion: oldMetadata?.Accion || '',
    NombreDocumento: oldMetadata?.NombreDocumento || ''
  };

  if (documentoPadreIds.length) {
    payload.DocumentoPadreId = documentoPadreIsMulti ? documentoPadreIds : documentoPadreIds[0];
  }

  return payload;
}

export async function rollbackModificacionFase2(params: {
  context: WebPartContext;
  webUrl: string;
  entries: IFase2RollbackEntry[];
  log?: LogFn;
}): Promise<void> {
  const log = params.log || (() => undefined);
  const entries = Array.isArray(params.entries) ? params.entries : [];

  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];

    try {
      if (entry.newPublishedUrl) {
        await recycleFile(params.context, params.webUrl, entry.newPublishedUrl);
        log(`🗑️ Fase 2 rollback | Nuevo eliminado: ${entry.newPublishedUrl}`);
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 2 rollback | No se pudo eliminar nuevo publicado ${entry.newPublishedUrl} -> ${message}`);
    }

    try {
      if (entry.historicoUrl && entry.oldOriginalUrl) {
        await moveFileByPath(params.context, params.webUrl, entry.historicoUrl, entry.oldOriginalUrl, true);
        log(`♻️ Fase 2 rollback | Viejo restaurado desde Históricos: ${entry.oldOriginalUrl}`);

        if (entry.oldOriginalMetadata) {
          await updateFileMetadataByPath(
            params.context,
            params.webUrl,
            entry.oldOriginalUrl,
            await buildProcesosRestorePayload(params.context, params.webUrl, entry.oldOriginalMetadata)
          );
          log(`🧾 Fase 2 rollback | Metadata original restaurada: ${entry.oldOriginalUrl}`);
        }
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 2 rollback | No se pudo restaurar desde Históricos ${entry.oldOriginalUrl} -> ${message}`);
    }
  }
}
