/* eslint-disable */
// @ts-nocheck
import { WebPartContext } from '@microsoft/sp-webpart-base';

let digestValue: string | null = null;
let digestExpiresAt = 0;

function ensureAbsoluteUrl(webUrl: string, relativeOrAbsoluteUrl: string): string {
  if (/^https?:\/\//i.test(relativeOrAbsoluteUrl)) {
    return relativeOrAbsoluteUrl;
  }

  const origin = new URL(webUrl).origin;
  return `${origin}${relativeOrAbsoluteUrl}`;
}

function escapeODataValue(value: string): string {
  return String(value || '').replace(/'/g, `''`);
}

function trimTrailingSlash(value: string): string {
  return String(value || '').replace(/\/$/, '');
}

function resolveWebUrlForServerRelativePath(
  context: WebPartContext,
  webUrl: string,
  serverRelativeUrl: string
): string {
  const currentWebUrl = trimTrailingSlash(webUrl);
  const currentWebPath = trimTrailingSlash(new URL(currentWebUrl).pathname);
  const siteUrl = trimTrailingSlash(context.pageContext.site.absoluteUrl || currentWebUrl);
  const sitePath = trimTrailingSlash(new URL(siteUrl).pathname);
  const targetPath = trimTrailingSlash(serverRelativeUrl);

  if (targetPath && (targetPath === currentWebPath || targetPath.indexOf(`${currentWebPath}/`) === 0)) {
    return currentWebUrl;
  }

  if (targetPath && (targetPath === sitePath || targetPath.indexOf(`${sitePath}/`) === 0)) {
    return siteUrl;
  }

  return currentWebUrl;
}

function isAlreadyExistsFolderResponse(status: number, text: string): boolean {
  if (status === 409) {
    return true;
  }

  const normalized = String(text || '').toLowerCase();
  return status === 400 && normalized.indexOf('ya existe un archivo o una carpeta con el nombre') !== -1;
}

async function getRequestDigest(context: WebPartContext, webUrl: string): Promise<string> {
  const now = Date.now();
  if (digestValue && now < digestExpiresAt) {
    return digestValue;
  }

  const response = await fetch(`${webUrl}/_api/contextinfo`, {
    method: 'POST',
    credentials: 'same-origin',
    headers: {
      Accept: 'application/json;odata=nometadata'
    }
  });

  if (!response.ok) {
    throw new Error(`No se pudo obtener FormDigest. HTTP ${response.status}`);
  }

  const json = await response.json();
  const timeoutSeconds = Number(json.FormDigestTimeoutSeconds || 1200);

  digestValue = json.FormDigestValue;
  digestExpiresAt = now + Math.max(30000, (timeoutSeconds - 60) * 1000);

  return digestValue || '';
}

export async function spGetJson<T>(context: WebPartContext, url: string): Promise<T> {
  const response = await fetch(url, {
    method: 'GET',
    credentials: 'same-origin',
    headers: {
      Accept: 'application/json;odata=nometadata'
    }
  });

  if (!response.ok) {
    const body = await safeReadText(response);
    throw new Error(`Error consultando SharePoint (${response.status}): ${body || response.statusText}`);
  }

  return response.json();
}

export async function spPostJson<T>(
  context: WebPartContext,
  webUrl: string,
  url: string,
  body: any,
  method?: 'POST' | 'MERGE' | 'DELETE'
): Promise<T | null> {
  const digest = await getRequestDigest(context, webUrl);
  const headers: Record<string, string> = {
    Accept: 'application/json;odata=nometadata',
    'Content-Type': 'application/json;odata=nometadata',
    'X-RequestDigest': digest
  };

  let httpMethod = 'POST';
  if (method === 'MERGE' || method === 'DELETE') {
    headers['X-HTTP-Method'] = method;
    headers['IF-MATCH'] = '*';
  }

  const response = await fetch(url, {
    method: httpMethod,
    credentials: 'same-origin',
    headers,
    body: body !== undefined && body !== null ? JSON.stringify(body) : undefined
  });

  if (!response.ok) {
    const text = await safeReadText(response);
    throw new Error(`Error enviando datos a SharePoint (${response.status}): ${text || response.statusText}`);
  }

  if (response.status === 204) {
    return null;
  }

  const text = await response.text();
  if (!text) {
    return null;
  }

  return JSON.parse(text);
}

export async function getAllItems<T>(context: WebPartContext, initialUrl: string): Promise<T[]> {
  const items: T[] = [];
  let nextUrl: string | undefined = initialUrl;

  while (nextUrl) {
    const json = await spGetJson<{ value?: T[]; ['@odata.nextLink']?: string; }>(context, nextUrl);
    const pageItems = (json.value || []) as T[];

    for (let i = 0; i < pageItems.length; i++) {
      items.push(pageItems[i]);
    }

    nextUrl = json['@odata.nextLink'];
  }

  return items;
}

export async function addListItem(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  payload: any
): Promise<number> {
  const url = `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(listTitle)}')/items`;
  const result = await spPostJson<any>(context, webUrl, url, payload, 'POST');
  return Number(result && result.Id);
}

export async function updateListItem(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  itemId: number,
  payload: any
): Promise<void> {
  const url = `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(listTitle)}')/items(${itemId})`;
  await spPostJson(context, webUrl, url, payload, 'MERGE');
}

export async function deleteListItem(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  itemId: number
): Promise<void> {
  const url = `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(listTitle)}')/items(${itemId})`;
  await spPostJson(context, webUrl, url, undefined, 'DELETE');
}

export async function getAttachmentFiles(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  itemId: number
): Promise<Array<{ FileName: string; ServerRelativeUrl: string; }>> {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(listTitle)}')/items(${itemId})/AttachmentFiles` +
    `?$select=FileName,ServerRelativeUrl`;

  return getAllItems(context, url);
}

export async function deleteAttachment(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  itemId: number,
  fileName: string
): Promise<void> {
  const digest = await getRequestDigest(context, webUrl);
  const url =
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(listTitle)}')/items(${itemId})` +
    `/AttachmentFiles/getByFileName('${escapeODataValue(fileName)}')`;

  const response = await fetch(url, {
    method: 'POST',
    credentials: 'same-origin',
    headers: {
      Accept: 'application/json;odata=nometadata',
      'X-RequestDigest': digest,
      'X-HTTP-Method': 'DELETE',
      'IF-MATCH': '*'
    }
  });

  if (!response.ok) {
    const text = await safeReadText(response);
    throw new Error(`No se pudo borrar el adjunto (${response.status}): ${text || response.statusText}`);
  }
}

export async function addAttachment(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  itemId: number,
  fileName: string,
  body: Blob
): Promise<void> {
  const digest = await getRequestDigest(context, webUrl);
  const url =
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(listTitle)}')/items(${itemId})` +
    `/AttachmentFiles/add(FileName='${escapeODataValue(fileName)}')`;

  const response = await fetch(url, {
    method: 'POST',
    credentials: 'same-origin',
    headers: {
      Accept: 'application/json;odata=nometadata',
      'X-RequestDigest': digest
    },
    body
  });

  if (!response.ok) {
    const text = await safeReadText(response);
    throw new Error(`No se pudo adjuntar el archivo (${response.status}): ${text || response.statusText}`);
  }
}

export async function ensureFolderPath(
  context: WebPartContext,
  webUrl: string,
  folderServerRelativeUrl: string
): Promise<void> {
  const targetWebUrl = resolveWebUrlForServerRelativePath(context, webUrl, folderServerRelativeUrl);
  const digest = await getRequestDigest(context, targetWebUrl);
  const normalizedFolderPath = trimTrailingSlash(folderServerRelativeUrl);
  const targetWebPath = trimTrailingSlash(new URL(targetWebUrl).pathname);
  const parts = normalizedFolderPath.split('/').filter(Boolean);
  const targetWebParts = targetWebPath.split('/').filter(Boolean);
  if (!parts.length) {
    return;
  }

  let currentPath = targetWebPath || '';
  for (let i = targetWebParts.length; i < parts.length; i++) {
    currentPath += `/${parts[i]}`;
    const url =
      `${targetWebUrl}/_api/web/folders/addUsingPath(decodedurl='${escapeODataValue(currentPath)}')`;

    const response = await fetch(url, {
      method: 'POST',
      credentials: 'same-origin',
      headers: {
        Accept: 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      }
    });

    if (!response.ok) {
      const text = await safeReadText(response);
      if (isAlreadyExistsFolderResponse(response.status, text)) {
        continue;
      }
      throw new Error(`No se pudo asegurar la carpeta "${currentPath}" (${response.status}): ${text || response.statusText}`);
    }
  }
}

export async function uploadFileToFolder(
  context: WebPartContext,
  webUrl: string,
  folderServerRelativeUrl: string,
  fileName: string,
  body: Blob
): Promise<string> {
  await ensureFolderPath(context, webUrl, folderServerRelativeUrl);
  const targetWebUrl = resolveWebUrlForServerRelativePath(context, webUrl, folderServerRelativeUrl);
  const digest = await getRequestDigest(context, targetWebUrl);
  const url =
    `${targetWebUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${escapeODataValue(folderServerRelativeUrl)}')` +
    `/Files/addUsingPath(decodedurl='${escapeODataValue(fileName)}',overwrite=true)`;

  const response = await fetch(url, {
    method: 'POST',
    credentials: 'same-origin',
    headers: {
      Accept: 'application/json;odata=nometadata',
      'X-RequestDigest': digest
    },
    body
  });

  if (!response.ok) {
    const text = await safeReadText(response);
    throw new Error(`No se pudo subir el archivo a TEMP (${response.status}): ${text || response.statusText}`);
  }

  return `${folderServerRelativeUrl.replace(/\/$/, '')}/${fileName}`;
}

export async function recycleFile(
  context: WebPartContext,
  webUrl: string,
  fileServerRelativeUrl: string
): Promise<void> {
  const digest = await getRequestDigest(context, webUrl);
  const url =
    `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${escapeODataValue(fileServerRelativeUrl)}')/recycle()`;

  const response = await fetch(url, {
    method: 'POST',
    credentials: 'same-origin',
    headers: {
      Accept: 'application/json;odata=nometadata',
      'X-RequestDigest': digest
    }
  });

  if (!response.ok) {
    const text = await safeReadText(response);
    throw new Error(`No se pudo reciclar el archivo (${response.status}): ${text || response.statusText}`);
  }
}

export async function safeReadText(response: Response): Promise<string> {
  try {
    return await response.text();
  } catch (_error) {
    return '';
  }
}

export {
  ensureAbsoluteUrl,
  escapeODataValue
};
