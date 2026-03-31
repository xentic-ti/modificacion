/* eslint-disable */
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  buildReviewedFileName,
  buildReviewedWorkbook,
  openExcelRevisionSession
} from '../utils/modificacionExcelHelper';

type LogFn = (message: string) => void;

interface ISolicitudItem {
  Id: number;
  Title?: string;
  NombreDocumento?: string;
}

interface IRelacionDocumentoItem {
  DocumentoHijoId?: number;
}

interface IDiagramaFlujoItem {
  Id: number;
}

export async function revisarExcelModificacion(params: {
  context: WebPartContext;
  file: IFilePickerResult;
  log?: LogFn;
}): Promise<{
  blob: Blob;
  fileName: string;
  processed: number;
  found: number;
  notFound: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const session = await openExcelRevisionSession(params.file);
  const valuesByRowNumber = new Map<number, { solicitudId: string; documentosHijos: string; diagramasFlujo: string; }>();
  const cache: { [key: string]: { solicitudId: string; documentosHijos: string; diagramasFlujo: string; found: boolean; } } = {};

  let processed = 0;
  let found = 0;
  let notFound = 0;

  log(`📄 Excel cargado: ${session.fileName}`);
  log(`📋 Filas detectadas para revisión: ${session.rows.length}`);

  for (let i = 0; i < session.rows.length; i++) {
    const row = session.rows[i];
    const documentName = String(row.documentName || '').trim();

    if (!documentName) {
      valuesByRowNumber.set(row.rowNumber, {
        solicitudId: '',
        documentosHijos: '',
        diagramasFlujo: ''
      });
      continue;
    }

    processed++;

    if (!cache[documentName]) {
      log(`🔎 Buscando solicitud para "${documentName}"...`);
      cache[documentName] = await resolveDocumentRelations(
        params.context,
        webUrl,
        documentName,
        log
      );
    }

    const result = cache[documentName];
    valuesByRowNumber.set(row.rowNumber, {
      solicitudId: result.solicitudId,
      documentosHijos: result.documentosHijos,
      diagramasFlujo: result.diagramasFlujo
    });

    if (result.found) {
      found++;
    } else {
      notFound++;
    }
  }

  const blob = await buildReviewedWorkbook(session, valuesByRowNumber);

  return {
    blob,
    fileName: buildReviewedFileName(session.fileName),
    processed,
    found,
    notFound
  };
}

async function resolveDocumentRelations(
  context: WebPartContext,
  webUrl: string,
  documentName: string,
  log: LogFn
): Promise<{ solicitudId: string; documentosHijos: string; diagramasFlujo: string; found: boolean; }> {
  const solicitud = await buscarSolicitudPorNombre(context, webUrl, documentName);

  if (!solicitud) {
    log(`⚠️ No se encontró la solicitud para "${documentName}".`);
    return {
      solicitudId: '',
      documentosHijos: '',
      diagramasFlujo: '',
      found: false
    };
  }

  log(`✅ Solicitud encontrada | Doc="${documentName}" | ID=${solicitud.Id}`);

  const hijosIds = await buscarDocumentosHijos(context, webUrl, solicitud.Id);
  const diagramasIds = await buscarDiagramasFlujo(context, webUrl, solicitud.Id);

  log(
    `📎 Resultado | SolicitudID=${solicitud.Id} | Hijos=${hijosIds.length} | Diagramas=${diagramasIds.length}`
  );

  return {
    solicitudId: String(solicitud.Id),
    documentosHijos: hijosIds.join('/'),
    diagramasFlujo: diagramasIds.join('/'),
    found: true
  };
}

async function buscarSolicitudPorNombre(
  context: WebPartContext,
  webUrl: string,
  documentName: string
): Promise<ISolicitudItem | null> {
  const filter = `(Title eq '${escapeODataValue(documentName)}' or NombreDocumento eq '${escapeODataValue(documentName)}')`;
  const url =
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items` +
    `?$select=Id,Title,NombreDocumento` +
    `&$top=2` +
    `&$filter=${encodeURIComponent(filter)}`;

  const items = await getAllItems<ISolicitudItem>(context, url);

  if (!items.length) {
    return null;
  }

  return items[0];
}

async function buscarDocumentosHijos(
  context: WebPartContext,
  webUrl: string,
  solicitudId: number
): Promise<number[]> {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('Relaciones Documentos')/items` +
    `?$select=DocumentoHijoId` +
    `&$top=5000` +
    `&$filter=DocumentoPadreId eq ${solicitudId}`;

  const items = await getAllItems<IRelacionDocumentoItem>(context, url);
  return uniqueSortedNumbers(items.map((item) => item.DocumentoHijoId));
}

async function buscarDiagramasFlujo(
  context: WebPartContext,
  webUrl: string,
  solicitudId: number
): Promise<number[]> {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('Diagramas de Flujo')/items` +
    `?$select=Id` +
    `&$top=5000` +
    `&$filter=SolicitudId eq ${solicitudId}`;

  const items = await getAllItems<IDiagramaFlujoItem>(context, url);
  return uniqueSortedNumbers(items.map((item) => item.Id));
}

async function getAllItems<T>(context: WebPartContext, initialUrl: string): Promise<T[]> {
  void context;
  const items: T[] = [];
  let nextUrl: string | undefined = initialUrl;

  while (nextUrl) {
    const response = await fetch(nextUrl, {
      method: 'GET',
      credentials: 'same-origin',
      headers: {
        Accept: 'application/json'
      }
    });
    await ensureSuccessfulResponse(response);

    const json: {
      value?: T[];
      ['@odata.nextLink']?: string;
      d?: {
        results?: T[];
        __next?: string;
      };
    } = await response.json();
    const pageItems = ((json.value || (json.d && json.d.results) || []) as T[]);

    for (let i = 0; i < pageItems.length; i++) {
      items.push(pageItems[i]);
    }

    nextUrl = json['@odata.nextLink'] || (json.d && json.d.__next) || undefined;
  }

  return items;
}

async function ensureSuccessfulResponse(response: Response): Promise<void> {
  if (response.ok) {
    return;
  }

  let errorText = response.statusText;

  try {
    const body = await response.text();
    if (body) {
      errorText = body;
    }
  } catch (_error) {
    // sin acción
  }

  throw new Error(`Error consultando SharePoint (${response.status}): ${errorText}`);
}

function uniqueSortedNumbers(values: Array<number | undefined>): number[] {
  const map: { [key: string]: boolean } = {};
  const result: number[] = [];

  for (let i = 0; i < values.length; i++) {
    const value = values[i];
    if (!value || map[`${value}`]) {
      continue;
    }

    map[`${value}`] = true;
    result.push(value);
  }

  result.sort((a, b) => a - b);
  return result;
}

function escapeODataValue(value: string): string {
  return String(value || '').replace(/'/g, `''`);
}
