/* eslint-disable */
// @ts-nocheck
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { generateDocxWithContentControls } from '../utils/wordContentControls';
import { generateXlsxWithPlaceholders } from '../utils/excelPlaceholders';
import { generatePptxWithPlaceholders } from '../utils/pptPlaceholders';
import { getWordTagsFromArrayBuffer } from '../utils/wordTagsHelper';
import { getExcelPlaceholdersFromArrayBuffer } from '../utils/excelPlaceholdersHelper';
import { getPptPlaceholdersFromArrayBuffer } from '../utils/pptPlaceholdersHelper';
import {
  addAttachment,
  deleteAttachment,
  ensureAbsoluteUrl,
  getAllItems,
  getAttachmentFiles,
  uploadFileToFolder
} from './sharepointRest.service';

type LogFn = (s: string) => void;

function obtenerPrimerNombreYApellido(displayName: string): string {
  const limpio = (displayName || '').replace(/\s+/g, ' ').trim();
  if (!limpio) return '';
  const partes = limpio.split(' ').filter(Boolean);

  if (partes.length === 3) return `${partes[0]} ${partes[1]}`;
  if (partes.length === 4) return `${partes[0]} ${partes[2]}`;
  if (partes.length >= 2) return `${partes[0]} ${partes[1]}`;
  return partes[0] || '';
}

function uniqueByPuestoYNombre(
  rows: Array<{ puesto: string; nombre: string; }>
): Array<{ puesto: string; nombre: string; }> {
  const seen = new Set<string>();
  const result: Array<{ puesto: string; nombre: string; }> = [];

  for (let i = 0; i < rows.length; i++) {
    const puesto = (rows[i].puesto || '').trim();
    const nombre = (rows[i].nombre || '').trim();
    const key = `${puesto.toLowerCase()}||${nombre.toLowerCase()}`;
    if (seen.has(key)) continue;
    seen.add(key);
    result.push({ puesto, nombre });
  }

  return result;
}

function excelToDDMMYYYY(value: any): string {
  if (value === null || value === undefined) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return `${String(value.getDate()).padStart(2, '0')}/${String(value.getMonth() + 1).padStart(2, '0')}/${value.getFullYear()}`;
  }

  const raw = String(value).trim();
  if (!raw) return '';

  const match = raw.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/);
  if (match) {
    const day = match[1].padStart(2, '0');
    const month = match[2].padStart(2, '0');
    const year = match[3].length === 2 ? `20${match[3]}` : match[3];
    return `${day}/${month}/${year}`;
  }

  const isoMatch = raw.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[T\s]+\d{1,2}:\d{2}(?::\d{2})?)?$/);
  if (isoMatch) {
    const year = isoMatch[1];
    const month = isoMatch[2].padStart(2, '0');
    const day = isoMatch[3].padStart(2, '0');
    return `${day}/${month}/${year}`;
  }

  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return `${String(parsed.getDate()).padStart(2, '0')}/${String(parsed.getMonth() + 1).padStart(2, '0')}/${parsed.getFullYear()}`;
  }

  return raw;
}

function getExtFromName(name: string): 'docx' | 'xlsx' | 'pptx' | 'other' {
  const lowerName = (name || '').toLowerCase().trim();
  if (lowerName.endsWith('.docx')) return 'docx';
  if (lowerName.endsWith('.xlsx') || lowerName.endsWith('.xlsm') || lowerName.endsWith('.xls')) return 'xlsx';
  if (lowerName.endsWith('.pptx')) return 'pptx';
  return 'other';
}

async function getGerentesFromAreas(params: {
  context: WebPartContext;
  webUrl: string;
  areaIds: number[];
}): Promise<Array<{ nombre: string; puesto: string; }>> {
  const ids = (params.areaIds || []).filter(Boolean);
  if (!ids.length) return [];

  const filter = ids.map((id) => `Id eq ${id}`).join(' or ');
  const url =
    `${params.webUrl}/_api/web/lists/getbytitle('Áreas de Negocio')/items` +
    `?$select=Id,Gerente/Title,Gerente/JobTitle&$expand=Gerente&$filter=${encodeURIComponent(filter)}`;

  const items = await getAllItems<any>(params.context, url);
  const map = new Map<number, { nombre: string; puesto: string; }>();

  for (let i = 0; i < items.length; i++) {
    map.set(items[i].Id, {
      nombre: items[i]?.Gerente?.Title || '',
      puesto: items[i]?.Gerente?.JobTitle || ''
    });
  }

  return ids.map((id) => map.get(id) || { nombre: '', puesto: '' });
}

export async function fillAndAttachFromFolder(params: {
  context: WebPartContext;
  webUrl: string;
  listTitle: string;
  itemId: number;
  originalFileName: string;
  fileByName: Map<string, string>;
  titulo: string;
  instanciaRaw: string;
  impactAreaIds: number[];
  dueno: string;
  fechaVigencia: string;
  fechaAprobacion: string;
  resumen: string;
  version: string;
  codigoDocumento?: string;
  relacionados?: Array<{ codigo: string; nombre: string; enlace: string; }>;
  diagramasFlujo?: Array<{ codigo: string; nombre: string; enlace: string; }>;
  replaceIfExists?: boolean;
  categoriaDoc?: string;
  tipoDocExcel?: string;
  esDocumentoApoyo?: boolean;
  gerenciaAprobadora?: string;
  tempDestinoFolderServerRelativeUrl?: string;
  log?: LogFn;
}): Promise<{ ok: boolean; tempFileServerRelativeUrl?: string; attachmentFileName?: string; error?: string; }> {
  const log = params.log || (() => undefined);
  const cleanName = (params.originalFileName || '').trim();
  const targetName = cleanName;

  if (!cleanName) {
    return { ok: false, error: 'No se recibió nombre de archivo.' };
  }

  const relativeUrl = params.fileByName.get(cleanName.toLowerCase());
  if (!relativeUrl) {
    return { ok: false, error: `No se encontró el archivo origen "${cleanName}".` };
  }

  const ext = getExtFromName(cleanName);
  if (ext === 'other') {
    return { ok: false, error: `Tipo no soportado para relleno: ${cleanName}` };
  }

  const fileResponse = await fetch(ensureAbsoluteUrl(params.webUrl, relativeUrl), {
    credentials: 'same-origin'
  });
  if (!fileResponse.ok) {
    return { ok: false, error: `No se pudo descargar el archivo origen. HTTP ${fileResponse.status}` };
  }

  const buffer = await fileResponse.arrayBuffer();
  let tagsEncontrados: string[] = [];

  if (ext === 'docx') {
    tagsEncontrados = await getWordTagsFromArrayBuffer(buffer);
  } else if (ext === 'xlsx') {
    tagsEncontrados = getExcelPlaceholdersFromArrayBuffer(buffer);
  } else if (ext === 'pptx') {
    tagsEncontrados = await getPptPlaceholdersFromArrayBuffer(buffer);
  }

  const gerentes = await getGerentesFromAreas({
    context: params.context,
    webUrl: params.webUrl,
    areaIds: params.impactAreaIds || []
  });

  const revisores = uniqueByPuestoYNombre(
    gerentes
      .map((gerente) => ({
        puesto: (gerente.puesto || '').trim(),
        nombre: obtenerPrimerNombreYApellido(gerente.nombre || '')
      }))
      .filter((row) => row.puesto !== '' || row.nombre !== '')
  );

  const puestos = revisores.map((row) => row.puesto);
  const nombres = revisores.map((row) => row.nombre);
  const fechaVigenciaTxt = excelToDDMMYYYY(params.fechaVigencia);
  const fechaAprobacionTxt = excelToDDMMYYYY(params.fechaAprobacion);
  const relacionados = params.relacionados || [];
  const diagramas = params.diagramasFlujo || [];

  const replacementsWordForDocx: Record<string, any> = {
    TituloDocumento: params.titulo || '',
    DuenoDocumento: params.dueno || '',
    FechaVigencia: fechaVigenciaTxt || '',
    FechaAprobacion: fechaAprobacionTxt || '',
    Version: params.version || '',
    CodigoDocumento: params.codigoDocumento || '',
    Resumen: params.resumen || '',
    PuestoRevisor: puestos.filter(Boolean),
    NombreRevisor: nombres.filter(Boolean)
  };

  const replacementsPlaceholders: Record<string, string> = {
    '{TituloDocumento}': params.titulo || '',
    '{DuenoDocumento}': params.dueno || '',
    '{FechaVigencia}': fechaVigenciaTxt || '',
    '{FechaAprobacion}': fechaAprobacionTxt || '',
    '{PuestoRevisor}': puestos.filter(Boolean).join('\n'),
    '{NombreRevisor}': nombres.filter(Boolean).join('\n'),
    '{Version}': params.version || '',
    '{CodigoDocumento}': params.codigoDocumento || '',
    '{Resumen}': params.resumen || '',
    '{CodigoDocumentoRelacionado}': relacionados.map((row) => row.codigo || '').filter(Boolean).join('\n'),
    '{NombreDocumentoRelacionado}': relacionados.map((row) => row.nombre || '').filter(Boolean).join('\n'),
    '{EnlaceDocumentoRelacionado}': relacionados.map((row) => row.enlace || '').filter(Boolean).join('\n'),
    '{CodigoDiagramaFlujo}': diagramas.map((row) => row.codigo || '').filter(Boolean).join('\n'),
    '{NombreDiagramaFlujo}': diagramas.map((row) => row.nombre || '').filter(Boolean).join('\n'),
    '{EnlaceDiagramaFlujo}': diagramas.map((row) => row.enlace || '').filter(Boolean).join('\n')
  };

  if (params.esDocumentoApoyo) {
    replacementsWordForDocx.GerenciaAprobadora = params.gerenciaAprobadora || '';
    replacementsPlaceholders['{GerenciaAprobadora}'] = params.gerenciaAprobadora || '';
  } else {
    replacementsWordForDocx.InstanciaAprobacion = params.instanciaRaw || '';
    replacementsPlaceholders['{InstanciaAprobacion}'] = params.instanciaRaw || '';
  }

  log(`🧾 ID=${params.itemId} | ${cleanName}`);
  log(`   Tags encontrados en archivo: ${tagsEncontrados.join(', ') || 'NINGUNO'}`);

  let outputBlob: Blob;
  if (ext === 'docx') {
    outputBlob = generateDocxWithContentControls(buffer, replacementsWordForDocx, {
      revisores,
      documentosRelacionados: relacionados.map((row) => ({
        codigoDocumento: row.codigo || '',
        nombreDocumento: row.nombre || '',
        enlace: row.enlace || ''
      })),
      flujosProceso: diagramas.map((row) => ({
        codigoDocumento: row.codigo || '',
        nombreDocumento: row.nombre || '',
        enlace: row.enlace || ''
      }))
    });
  } else if (ext === 'xlsx') {
    outputBlob = generateXlsxWithPlaceholders(buffer, replacementsPlaceholders);
  } else {
    outputBlob = generatePptxWithPlaceholders(buffer, replacementsPlaceholders);
  }

  if (params.replaceIfExists) {
    const attachments = await getAttachmentFiles(params.context, params.webUrl, params.listTitle, params.itemId);
    const exists = attachments.some((attachment) => String(attachment.FileName || '').toLowerCase() === targetName.toLowerCase());
    if (exists) {
      await deleteAttachment(params.context, params.webUrl, params.listTitle, params.itemId, targetName);
      log(`🗑️ Attachment reemplazado (borrado previo) | ID=${params.itemId} | ${targetName}`);
    }
  }

  await addAttachment(params.context, params.webUrl, params.listTitle, params.itemId, targetName, outputBlob);
  log(`🧩📎 Rellenado+Adjuntado OK | ID=${params.itemId} | ${targetName}`);

  let tempFileServerRelativeUrl: string | undefined;
  if (params.tempDestinoFolderServerRelativeUrl) {
    tempFileServerRelativeUrl = await uploadFileToFolder(
      params.context,
      params.webUrl,
      params.tempDestinoFolderServerRelativeUrl,
      targetName,
      outputBlob
    );
    log(`📂 Copiado a TEMP | ${tempFileServerRelativeUrl}`);
  }

  return {
    ok: true,
    tempFileServerRelativeUrl,
    attachmentFileName: targetName
  };
}
