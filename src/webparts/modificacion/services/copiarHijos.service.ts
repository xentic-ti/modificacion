/* eslint-disable */
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { addListItem, getAllItems, spGetJson, updateListItem } from './sharepointRest.service';
import { buildCopiarHijosWorkbook, ICopiarHijosReportRow } from '../utils/copiarHijosExcel';

type LogFn = (message: string) => void;

interface ISolicitudItem {
  Id: number;
  Title?: string;
  NombreDocumento?: string;
  [key: string]: any;
}

interface IRelacionDocumentoItem {
  Id: number;
  Title?: string;
  [key: string]: any;
}

function escapeODataValue(value: string): string {
  return String(value || '').replace(/'/g, `''`);
}

function parseSolicitudId(value: number | string, label: string): number {
  const id = Number(String(value || '').trim());
  if (!Number.isFinite(id) || id <= 0 || Math.floor(id) !== id) {
    throw new Error(`${label} debe ser un ID numerico valido.`);
  }

  return id;
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

function formatSolicitudName(item: ISolicitudItem | undefined): string {
  if (!item) {
    return '';
  }

  return String(item.Title || item.NombreDocumento || '').trim();
}

function formatSolicitudRefs(ids: number[], solicitudById: Map<number, ISolicitudItem>): string {
  return ids
    .map((id) => {
      const name = formatSolicitudName(solicitudById.get(id));
      return name ? `${id} - ${name}` : String(id);
    })
    .join(' | ');
}

async function tryGetFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldTitleOrInternalName: string
): Promise<string | null> {
  try {
    const field = await spGetJson<any>(
      context,
      `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(listTitle)}')/fields/getbyinternalnameortitle('${escapeODataValue(fieldTitleOrInternalName)}')?$select=InternalName`
    );
    return String(field?.InternalName || fieldTitleOrInternalName);
  } catch (_error) {
    return null;
  }
}

async function getFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldTitleOrInternalName: string
): Promise<string> {
  const found = await tryGetFieldInternalName(context, webUrl, listTitle, fieldTitleOrInternalName);
  if (!found) {
    throw new Error(`No se encontro el campo "${fieldTitleOrInternalName}" en "${listTitle}".`);
  }

  return found;
}

async function resolveFirstExistingFieldInternalName(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  candidates: string[]
): Promise<string> {
  for (let i = 0; i < candidates.length; i++) {
    const found = await tryGetFieldInternalName(context, webUrl, listTitle, candidates[i]);
    if (found) {
      return found;
    }
  }

  throw new Error(`No se encontro el campo esperado en "${listTitle}": ${candidates.join(', ')}`);
}

async function getAllowMultipleValues(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldInternalName: string
): Promise<boolean> {
  const field = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(listTitle)}')/fields/getbyinternalnameortitle('${escapeODataValue(fieldInternalName)}')?$select=AllowMultipleValues,TypeAsString`
  );

  return !!field.AllowMultipleValues || String(field?.TypeAsString || '').toLowerCase().indexOf('multi') !== -1;
}

async function getSolicitudById(
  context: WebPartContext,
  webUrl: string,
  solicitudId: number
): Promise<ISolicitudItem | null> {
  try {
    return await spGetJson<ISolicitudItem>(
      context,
      `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items(${solicitudId})?$select=Id,Title,NombreDocumento`
    );
  } catch (_error) {
    return null;
  }
}

export async function copiarHijosRelacionesDocumentos(params: {
  context: WebPartContext;
  padreOrigenId: number | string;
  padreDestinoId: number | string;
  log?: LogFn;
}): Promise<{
  blob: Blob;
  fileName: string;
  totalHijos: number;
  creadas: number;
  omitidas: number;
  errores: number;
  docPadresActualizados: number;
  docPadresOmitidos: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const padreOrigenId = parseSolicitudId(params.padreOrigenId, 'El ID padre origen');
  const padreDestinoId = parseSolicitudId(params.padreDestinoId, 'El ID padre destino');

  if (padreOrigenId === padreDestinoId) {
    throw new Error('El ID padre origen y el ID padre destino deben ser distintos.');
  }

  log(`🔎 Validando solicitudes origen=${padreOrigenId} y destino=${padreDestinoId}...`);
  const padreOrigen = await getSolicitudById(params.context, webUrl, padreOrigenId);
  const padreDestino = await getSolicitudById(params.context, webUrl, padreDestinoId);

  if (!padreOrigen) {
    throw new Error(`No existe la solicitud padre origen ${padreOrigenId}.`);
  }

  if (!padreDestino) {
    throw new Error(`No existe la solicitud padre destino ${padreDestinoId}.`);
  }

  const padreOrigenNombre = formatSolicitudName(padreOrigen);
  const padreDestinoNombre = formatSolicitudName(padreDestino);

  const parentField = await getFieldInternalName(params.context, webUrl, 'Relaciones Documentos', 'DocumentoPadre');
  const childField = await getFieldInternalName(params.context, webUrl, 'Relaciones Documentos', 'DocumentoHijo');
  const parentFieldId = `${parentField}Id`;
  const childFieldId = `${childField}Id`;

  const docPadresField = await resolveFirstExistingFieldInternalName(params.context, webUrl, 'Solicitudes', ['docpadres', 'DocPadres', 'DocumentoPadre']);
  const docPadresFieldId = `${docPadresField}Id`;
  const docPadresIsMulti = await getAllowMultipleValues(params.context, webUrl, 'Solicitudes', docPadresField);
  if (!docPadresIsMulti) {
    throw new Error(`El campo ${docPadresField} de Solicitudes no permite multiples valores.`);
  }
  log(`🔎 Campo de padres en Solicitudes resuelto: ${docPadresFieldId}`);

  log('🔎 Consultando hijos del padre origen en Relaciones Documentos...');
  const origenRelaciones = await getAllItems<IRelacionDocumentoItem>(
    params.context,
    `${webUrl}/_api/web/lists/getbytitle('Relaciones Documentos')/items` +
      `?$select=Id,Title,${parentFieldId},${childFieldId},${childField}/Id,${childField}/Title` +
      `&$expand=${childField}` +
      `&$top=5000&$filter=${parentFieldId} eq ${padreOrigenId}`
  );
  log(`🔗 Relaciones origen encontradas: ${origenRelaciones.length}`);

  log('🔎 Consultando relaciones ya existentes del padre destino...');
  const destinoRelaciones = await getAllItems<IRelacionDocumentoItem>(
    params.context,
    `${webUrl}/_api/web/lists/getbytitle('Relaciones Documentos')/items` +
      `?$select=Id,${parentFieldId},${childFieldId}&$top=5000&$filter=${parentFieldId} eq ${padreDestinoId}`
  );

  log('🔎 Consultando Solicitudes para actualizar DocPadres de hijos...');
  const solicitudes = await getAllItems<ISolicitudItem>(
    params.context,
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items?$select=Id,Title,NombreDocumento,${docPadresFieldId}&$top=5000`
  );
  const solicitudById = new Map<number, ISolicitudItem>();
  for (let i = 0; i < solicitudes.length; i++) {
    solicitudById.set(Number(solicitudes[i].Id || 0), solicitudes[i]);
  }

  const hijosDestinoExistentes = new Set<number>();
  for (let i = 0; i < destinoRelaciones.length; i++) {
    const hijoId = Number(destinoRelaciones[i][childFieldId] || 0);
    if (hijoId > 0) {
      hijosDestinoExistentes.add(hijoId);
    }
  }

  const hijosProcesados = new Set<number>();
  const reportRows: ICopiarHijosReportRow[] = [];
  let creadas = 0;
  let omitidas = 0;
  let errores = 0;
  let docPadresActualizados = 0;
  let docPadresOmitidos = 0;

  if (!origenRelaciones.length) {
    reportRows.push({
      PadreOrigenId: padreOrigenId,
      PadreOrigenNombre: padreOrigenNombre,
      PadreDestinoId: padreDestinoId,
      PadreDestinoNombre: padreDestinoNombre,
      HijoId: '',
      HijoNombre: '',
      RelacionOrigenId: '',
      RelacionNuevaId: '',
      RelacionNuevaTitle: '',
      DocPadresAntes: '',
      DocPadresDespues: '',
      Estado: 'SIN_HIJOS',
      Observaciones: 'No se encontraron relaciones hijas para el padre origen.'
    });
  }

  for (let i = 0; i < origenRelaciones.length; i++) {
    const relacion = origenRelaciones[i];
    const hijoId = Number(relacion[childFieldId] || relacion[childField]?.Id || 0);
    const childSolicitud = solicitudById.get(hijoId);
    const hijoNombre = String(relacion[childField]?.Title || formatSolicitudName(childSolicitud) || '').trim();
    const relacionNuevaTitle = hijoId ? `P${padreDestinoId}-H${hijoId}` : '';
    let relacionNuevaId: number | '' = '';
    let estadoRelacion = '';
    let observacionesRelacion = '';

    if (!hijoId) {
      errores++;
      reportRows.push({
        PadreOrigenId: padreOrigenId,
        PadreOrigenNombre: padreOrigenNombre,
        PadreDestinoId: padreDestinoId,
        PadreDestinoNombre: padreDestinoNombre,
        HijoId: '',
        HijoNombre: hijoNombre,
        RelacionOrigenId: Number(relacion.Id || 0) || '',
        RelacionNuevaId: '',
        RelacionNuevaTitle: '',
        DocPadresAntes: '',
        DocPadresDespues: '',
        Estado: 'ERROR',
        Observaciones: 'La relacion origen no tiene DocumentoHijoId valido.'
      });
      continue;
    }

    if (hijosProcesados.has(hijoId)) {
      omitidas++;
      reportRows.push({
        PadreOrigenId: padreOrigenId,
        PadreOrigenNombre: padreOrigenNombre,
        PadreDestinoId: padreDestinoId,
        PadreDestinoNombre: padreDestinoNombre,
        HijoId: hijoId,
        HijoNombre: hijoNombre,
        RelacionOrigenId: Number(relacion.Id || 0) || '',
        RelacionNuevaId: '',
        RelacionNuevaTitle: relacionNuevaTitle,
        DocPadresAntes: '',
        DocPadresDespues: '',
        Estado: 'SKIP',
        Observaciones: 'El hijo aparece repetido en las relaciones del padre origen; no se genera duplicado ni se vuelve a tocar DocPadres.'
      });
      continue;
    }

    hijosProcesados.add(hijoId);

    if (hijosDestinoExistentes.has(hijoId)) {
      omitidas++;
      estadoRelacion = 'SKIP_RELACION';
      observacionesRelacion = 'Ya existe una relacion para el padre destino con este mismo hijo.';
    } else {
      try {
        relacionNuevaId = await addListItem(params.context, webUrl, 'Relaciones Documentos', {
          Title: relacionNuevaTitle,
          [parentFieldId]: padreDestinoId,
          [childFieldId]: hijoId
        });

        creadas++;
        hijosDestinoExistentes.add(hijoId);
        estadoRelacion = 'CREADA';
        observacionesRelacion = 'Relacion creada con el nuevo padre. No se modifico ni elimino la relacion origen.';
        log(`✅ Relacion creada | Title=${relacionNuevaTitle} | PadreDestino=${padreDestinoId} | Hijo=${hijoId} | Relacion=${relacionNuevaId}`);
      } catch (createError) {
        errores++;
        estadoRelacion = 'ERROR_RELACION';
        observacionesRelacion = createError instanceof Error ? createError.message : String(createError);
      }
    }

    const currentIds = childSolicitud ? normalizeLookupIds(childSolicitud[docPadresFieldId]) : [];
    const beforeText = childSolicitud ? formatSolicitudRefs(currentIds, solicitudById) : '';
    const nextIds = Array.from(new Set(currentIds.concat([padreDestinoId]))).sort((a, b) => a - b);
    const afterText = childSolicitud ? formatSolicitudRefs(nextIds, solicitudById) : '';
    let estadoDocPadres = '';
    let observacionesDocPadres = '';

    if (estadoRelacion === 'ERROR_RELACION') {
      estadoDocPadres = 'SKIP_DOCPADRES';
      observacionesDocPadres = 'No se actualizo DocPadres porque no se pudo crear la relacion nueva.';
    } else if (!childSolicitud) {
      errores++;
      estadoDocPadres = 'ERROR_DOCPADRES';
      observacionesDocPadres = `No existe la solicitud hija ${hijoId} en Solicitudes.`;
    } else if (currentIds.indexOf(padreDestinoId) !== -1) {
      docPadresOmitidos++;
      estadoDocPadres = 'SKIP_DOCPADRES';
      observacionesDocPadres = 'La solicitud hija ya tenia el nuevo padre en DocPadres.';
    } else {
      await updateListItem(params.context, webUrl, 'Solicitudes', hijoId, {
        [docPadresFieldId]: nextIds
      });

      childSolicitud[docPadresFieldId] = nextIds;
      docPadresActualizados++;
      estadoDocPadres = 'DOCPADRES_ACTUALIZADO';
      observacionesDocPadres = 'Se agrego el nuevo padre en DocPadres sin eliminar valores anteriores.';
      log(`✅ DocPadres actualizado | Hijo=${hijoId} | Antes="${beforeText}" | Despues="${afterText}"`);
    }

    reportRows.push({
      PadreOrigenId: padreOrigenId,
      PadreOrigenNombre: padreOrigenNombre,
      PadreDestinoId: padreDestinoId,
      PadreDestinoNombre: padreDestinoNombre,
      HijoId: hijoId,
      HijoNombre: hijoNombre,
      RelacionOrigenId: Number(relacion.Id || 0) || '',
      RelacionNuevaId: relacionNuevaId,
      RelacionNuevaTitle: relacionNuevaTitle,
      DocPadresAntes: beforeText,
      DocPadresDespues: afterText,
      Estado: estadoRelacion === 'CREADA' && estadoDocPadres === 'DOCPADRES_ACTUALIZADO' ? 'CREADA_ACTUALIZADA' : `${estadoRelacion}_${estadoDocPadres}`,
      Observaciones: `${observacionesRelacion} ${observacionesDocPadres}`.trim()
    });
  }

  const report = buildCopiarHijosWorkbook(reportRows);

  return {
    blob: report.blob,
    fileName: report.fileName,
    totalHijos: hijosProcesados.size,
    creadas,
    omitidas,
    errores,
    docPadresActualizados,
    docPadresOmitidos
  };
}
