/* eslint-disable */
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { rollbackTempFiles } from './rollbackTempFiles.service';
import { deleteListItem, updateListItem } from './sharepointRest.service';

type LogFn = (s: string) => void;

export async function rollbackModificacionFase1(params: {
  context: WebPartContext;
  webUrl: string;
  createdSolicitudIds: number[];
  oldSolicitudIds: number[];
  tempFileUrls: string[];
  log?: LogFn;
}): Promise<void> {
  const log = params.log || (() => undefined);
  const createdIds = (params.createdSolicitudIds || []).filter((id) => Number(id) > 0);
  const oldIds = (params.oldSolicitudIds || []).filter((id) => Number(id) > 0);

  for (let i = 0; i < createdIds.length; i++) {
    const id = createdIds[i];
    try {
      await deleteListItem(params.context, params.webUrl, 'Solicitudes', id);
      log(`🗑️ Nueva solicitud eliminada: ${id}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ No se pudo eliminar la nueva solicitud ${id} -> ${message}`);
    }
  }

  for (let i = 0; i < oldIds.length; i++) {
    const id = oldIds[i];
    try {
      await updateListItem(params.context, params.webUrl, 'Solicitudes', id, {
        EsVersionActualDocumento: true
      });
      log(`♻️ Solicitud antigua restaurada: ${id}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ No se pudo restaurar la solicitud antigua ${id} -> ${message}`);
    }
  }

  await rollbackTempFiles({
    context: params.context,
    webUrl: params.webUrl,
    fileUrls: params.tempFileUrls || [],
    log
  });
}
