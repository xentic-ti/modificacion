/* eslint-disable */
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { escapeODataValue, getAllItems } from './sharepointRest.service';

type LogFn = (s: string) => void;

export async function listFilesRecursive(
  context: WebPartContext,
  webUrl: string,
  folderServerRelativeUrl: string,
  log?: LogFn
): Promise<Array<{ Name: string; ServerRelativeUrl: string; }>> {
  const writeLog = log || (() => undefined);
  const files: Array<{ Name: string; ServerRelativeUrl: string; }> = [];

  async function walk(currentFolderUrl: string): Promise<void> {
    writeLog(`📂 Explorando carpeta: ${currentFolderUrl}`);

    const filesUrl =
      `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${escapeODataValue(currentFolderUrl)}')/Files` +
      `?$select=Name,ServerRelativeUrl&$top=5000`;
    const foldersUrl =
      `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${escapeODataValue(currentFolderUrl)}')/Folders` +
      `?$select=Name,ServerRelativeUrl&$top=5000`;

    const folderFiles = await getAllItems<{ Name: string; ServerRelativeUrl: string; }>(context, filesUrl);
    const folders = await getAllItems<{ Name: string; ServerRelativeUrl: string; }>(context, foldersUrl);

    for (let i = 0; i < folderFiles.length; i++) {
      files.push(folderFiles[i]);
    }

    for (let i = 0; i < folders.length; i++) {
      const folder = folders[i];
      const name = String(folder.Name || '').toLowerCase();
      if (name === 'forms') {
        continue;
      }

      await walk(folder.ServerRelativeUrl);
    }
  }

  await walk(folderServerRelativeUrl);
  writeLog(`📁 Archivos indexados: ${files.length}`);
  return files;
}
