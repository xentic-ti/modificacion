/* eslint-disable */
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { recycleFile } from './sharepointRest.service';

type LogFn = (s: string) => void;

export async function rollbackTempFiles(params: {
  context: WebPartContext;
  webUrl: string;
  fileUrls: string[];
  log?: LogFn;
}): Promise<void> {
  const log = params.log || (() => undefined);
  const fileUrls = params.fileUrls || [];

  if (!fileUrls.length) {
    log('⚠️ No hay archivos temporales para eliminar.');
    return;
  }

  log(`🧹 Eliminando ${fileUrls.length} archivos temporales...`);

  for (let i = 0; i < fileUrls.length; i++) {
    const fileUrl = fileUrls[i];
    if (!fileUrl) {
      continue;
    }

    try {
      await recycleFile(params.context, params.webUrl, fileUrl);
      log(`🗑️ Archivo temporal eliminado: ${fileUrl}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ No se pudo eliminar: ${fileUrl} -> ${message}`);
    }
  }

  log('✅ Limpieza de TEMP_MIGRACION_WORD finalizada.');
}
