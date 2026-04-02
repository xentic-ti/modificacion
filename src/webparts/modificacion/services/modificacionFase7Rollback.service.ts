/* eslint-disable */
// @ts-nocheck
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { addAttachment, deleteAttachment, getAttachmentFiles } from './sharepointRest.service';

type LogFn = (s: string) => void;

export interface IFase7RollbackAttachment {
  fileName: string;
  content: Blob;
}

export interface IFase7RollbackEntry {
  diagramItemId: number;
  solicitudPadreId: number;
  documentoPadre: string;
  nombreDocumento: string;
  uploadedAttachmentFileName: string;
  previousAttachments: IFase7RollbackAttachment[];
}

export async function snapshotDiagramAttachments(params: {
  context: WebPartContext;
  webUrl: string;
  itemId: number;
}): Promise<IFase7RollbackAttachment[]> {
  const attachmentsInfo = await getAttachmentFiles(params.context, params.webUrl, 'Diagramas de Flujo', params.itemId);
  const attachments: IFase7RollbackAttachment[] = [];

  for (let i = 0; i < attachmentsInfo.length; i++) {
    const file = attachmentsInfo[i];
    const response = await fetch(`${new URL(params.webUrl).origin}${file.ServerRelativeUrl}`, {
      credentials: 'same-origin'
    });

    if (!response.ok) {
      throw new Error(`No se pudo descargar el adjunto ${file.FileName}. HTTP ${response.status}`);
    }

    attachments.push({
      fileName: file.FileName,
      content: await response.blob()
    });
  }

  return attachments;
}

export async function rollbackModificacionFase7(params: {
  context: WebPartContext;
  webUrl: string;
  entries: IFase7RollbackEntry[];
  log?: LogFn;
}): Promise<void> {
  const log = params.log || (() => undefined);
  const entries = Array.isArray(params.entries) ? params.entries : [];

  for (let i = entries.length - 1; i >= 0; i--) {
    const entry = entries[i];

    try {
      const currentAttachments = await getAttachmentFiles(params.context, params.webUrl, 'Diagramas de Flujo', entry.diagramItemId);
      for (let j = 0; j < currentAttachments.length; j++) {
        await deleteAttachment(params.context, params.webUrl, 'Diagramas de Flujo', entry.diagramItemId, currentAttachments[j].FileName);
      }
      log(`🗑️ Fase 7 rollback | Adjuntos actuales eliminados: ${entry.diagramItemId}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 7 rollback | No se pudieron limpiar adjuntos actuales del diagrama ${entry.diagramItemId} -> ${message}`);
    }

    try {
      for (let j = 0; j < entry.previousAttachments.length; j++) {
        const attachment = entry.previousAttachments[j];
        await addAttachment(params.context, params.webUrl, 'Diagramas de Flujo', entry.diagramItemId, attachment.fileName, attachment.content);
      }
      log(`♻️ Fase 7 rollback | Adjuntos restaurados: ${entry.diagramItemId}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      log(`⚠️ Fase 7 rollback | No se pudieron restaurar adjuntos del diagrama ${entry.diagramItemId} -> ${message}`);
    }
  }
}
