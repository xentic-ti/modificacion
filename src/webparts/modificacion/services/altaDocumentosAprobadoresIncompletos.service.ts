/* eslint-disable */
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { escapeODataValue, getAllItems } from './sharepointRest.service';
import {
  buildAltaDocumentosAprobadoresIncompletosWorkbook,
  IAltaDocumentoAprobadorIncompletoRow
} from '../utils/altaDocumentosAprobadoresIncompletosExcel';

type LogFn = (message: string) => void;

const LIST_SOLICITUDES = 'Solicitudes';
const LIST_APROBADORES = 'Aprobadores por Solicitudes';
const LIST_MOTIVOS = 'Motivos';
const LIST_ACCIONES = 'Acciones';
const LIST_AREAS_NEGOCIO = 'Áreas de Negocio';
const ROL_REVISOR_IMPACTADO = 'Revisor Impactado';

interface ISolicitudItem {
  Id: number;
  Title?: string;
  NombreDocumento?: string;
  CodigoDocumento?: string;
  VersionDocumento?: string;
  MotivoId?: number;
  Motivo?: {
    Title?: string;
  };
  Accion?: string;
  Estado?: {
    Title?: string;
  };
  TipoDocumento?: {
    Id?: number;
    Title?: string;
  };
  AreasImpactadas?: Array<{
    Id?: number;
    Title?: string;
  }>;
  Created?: string;
  Modified?: string;
}

interface IAprobadorSolicitudItem {
  Id: number;
  SolicitudId?: number;
  Rol?: string;
  AprobadorId?: number;
  Aprobador?: {
    Id?: number;
    Title?: string;
    EMail?: string;
  };
  ImpactadoPorArea?: boolean | string | number;
  ImpactadoPorMotivo?: boolean | string | number;
  ImpactadoPorAccion?: boolean | string | number;
  Created?: string;
  Modified?: string;
}

interface IExpectedApprover {
  id: number;
  title: string;
  email: string;
  tipo: 'Motivo' | 'Area' | 'Accion';
  origen: string;
  flag: 'ImpactadoPorMotivo' | 'ImpactadoPorArea' | 'ImpactadoPorAccion';
}

export async function exportarAltaDocumentosAprobadoresIncompletos(params: {
  context: WebPartContext;
  log?: LogFn;
}): Promise<{
  blob: Blob;
  fileName: string;
  totalSolicitudes: number;
  totalSolicitudesObjetivo: number;
  totalSolicitudesConIncompletos: number;
  totalFilas: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;

  log('🔎 Consultando Solicitudes...');
  const solicitudes = await obtenerSolicitudes(params.context, webUrl);
  log(`📋 Solicitudes leidas: ${solicitudes.length}`);

  const solicitudesObjetivo = solicitudes.filter(esAltaDocumentoNoVigente);
  log(`📌 Solicitudes Alta de documentos con estado distinto de Vigente: ${solicitudesObjetivo.length}`);

  log('🔎 Consultando Aprobadores por Solicitudes...');
  const aprobadores = await obtenerAprobadores(params.context, webUrl);
  log(`👥 Registros de aprobadores leidos: ${aprobadores.length}`);

  const aprobadoresBySolicitud = buildAprobadoresBySolicitud(aprobadores);
  const rows: IAltaDocumentoAprobadorIncompletoRow[] = [];
  const solicitudesConIncompletos = new Set<number>();

  for (let i = 0; i < solicitudesObjetivo.length; i++) {
    const solicitud = solicitudesObjetivo[i];
    const solicitudId = Number(solicitud.Id || 0);
    const registros = aprobadoresBySolicitud.get(solicitudId) || [];
    const expected = await obtenerAprobadoresEsperados(params.context, webUrl, solicitud);

    if (!expected.length) {
      solicitudesConIncompletos.add(solicitudId);
      rows.push(buildReportRow(
        solicitud,
        null,
        null,
        'No se encontraron aprobadores esperados configurados por motivo, área ni acción para esta solicitud.'
      ));
      continue;
    }

    for (let j = 0; j < expected.length; j++) {
      const expectedApprover = expected[j];
      const registrosMismoAprobador = registros.filter((registro) => getAprobadorId(registro) === expectedApprover.id);
      const registroCorrecto = registrosMismoAprobador.find((registro) =>
        normalizeKey(registro.Rol || '') === normalizeKey(ROL_REVISOR_IMPACTADO) &&
        isTruthyField((registro as any)[expectedApprover.flag])
      );

      if (registroCorrecto) {
        continue;
      }

      const bestRegistro = registrosMismoAprobador[0] || null;
      solicitudesConIncompletos.add(solicitudId);
      rows.push(buildReportRow(
        solicitud,
        bestRegistro,
        expectedApprover,
        buildMotivoFaltante(expectedApprover, bestRegistro, registros.length)
      ));
    }
  }

  rows.sort(compareRows);

  if (rows.length) {
    log(`⚠️ Solicitudes con aprobadores incompletos: ${solicitudesConIncompletos.size}`);
    log(`⚠️ Filas incluidas en Excel: ${rows.length}`);
  } else {
    log('✅ No se encontraron aprobadores incompletos para las solicitudes objetivo.');
  }

  const report = buildAltaDocumentosAprobadoresIncompletosWorkbook(rows);

  return {
    blob: report.blob,
    fileName: report.fileName,
    totalSolicitudes: solicitudes.length,
    totalSolicitudesObjetivo: solicitudesObjetivo.length,
    totalSolicitudesConIncompletos: solicitudesConIncompletos.size,
    totalFilas: rows.length
  };
}

async function obtenerSolicitudes(context: WebPartContext, webUrl: string): Promise<ISolicitudItem[]> {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(LIST_SOLICITUDES)}')/items` +
    `?$select=Id,Title,NombreDocumento,CodigoDocumento,VersionDocumento,MotivoId,Motivo/Title,Accion,Estado/Title,` +
    `TipoDocumento/Id,TipoDocumento/Title,AreasImpactadas/Id,AreasImpactadas/Title,Created,Modified` +
    `&$expand=Motivo,Estado,TipoDocumento,AreasImpactadas` +
    `&$top=5000`;

  return getAllItems<ISolicitudItem>(context, url);
}

async function obtenerAprobadores(context: WebPartContext, webUrl: string): Promise<IAprobadorSolicitudItem[]> {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(LIST_APROBADORES)}')/items` +
    `?$select=Id,SolicitudId,Rol,AprobadorId,Aprobador/Id,Aprobador/Title,Aprobador/EMail,` +
    `ImpactadoPorArea,ImpactadoPorMotivo,ImpactadoPorAccion,Created,Modified` +
    `&$expand=Aprobador` +
    `&$top=5000`;

  return getAllItems<IAprobadorSolicitudItem>(context, url);
}

async function obtenerAprobadoresEsperados(
  context: WebPartContext,
  webUrl: string,
  solicitud: ISolicitudItem
): Promise<IExpectedApprover[]> {
  const result: IExpectedApprover[] = [];

  result.push(...await obtenerAprobadoresPorMotivo(context, webUrl, solicitud));
  result.push(...await obtenerAprobadoresPorArea(context, webUrl, solicitud));

  if (normalizeKey(solicitud.TipoDocumento?.Title || '') === 'procedimientos') {
    result.push(...await obtenerAprobadoresPorAccion(context, webUrl, solicitud));
  }

  return dedupeExpectedApprovers(result);
}

async function obtenerAprobadoresPorMotivo(
  context: WebPartContext,
  webUrl: string,
  solicitud: ISolicitudItem
): Promise<IExpectedApprover[]> {
  const motivoId = Number(solicitud.MotivoId || 0);
  if (!motivoId) {
    return [];
  }

  const motivo = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(LIST_MOTIVOS)}')/items` +
    `?$select=Id,Title,Aprobadores/Id,Aprobadores/Title,Aprobadores/EMail` +
    `&$expand=Aprobadores` +
    `&$filter=Id eq ${motivoId}` +
    `&$top=1`
  );

  const item = motivo[0];
  const personas = getPeople(item?.Aprobadores);
  return personas.map((persona) => ({
    ...persona,
    tipo: 'Motivo' as 'Motivo',
    origen: String(item?.Title || solicitud.Motivo?.Title || '').trim(),
    flag: 'ImpactadoPorMotivo' as 'ImpactadoPorMotivo'
  }));
}

async function obtenerAprobadoresPorArea(
  context: WebPartContext,
  webUrl: string,
  solicitud: ISolicitudItem
): Promise<IExpectedApprover[]> {
  const areasIds = (solicitud.AreasImpactadas || [])
    .map((area) => Number(area?.Id || 0))
    .filter((id) => Number.isFinite(id) && id > 0);

  if (!areasIds.length) {
    return [];
  }

  const filter = areasIds.map((id) => `Id eq ${id}`).join(' or ');
  const areas = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(LIST_AREAS_NEGOCIO)}')/items` +
    `?$select=Id,Title,Gerente/Id,Gerente/Title,Gerente/EMail` +
    `&$expand=Gerente` +
    `&$filter=${encodeURIComponent(filter)}` +
    `&$top=5000`
  );

  const result: IExpectedApprover[] = [];
  for (let i = 0; i < areas.length; i++) {
    const personas = getPeople(areas[i]?.Gerente);
    for (let j = 0; j < personas.length; j++) {
      result.push({
        ...personas[j],
        tipo: 'Area',
        origen: String(areas[i]?.Title || '').trim(),
        flag: 'ImpactadoPorArea'
      });
    }
  }

  return result;
}

async function obtenerAprobadoresPorAccion(
  context: WebPartContext,
  webUrl: string,
  solicitud: ISolicitudItem
): Promise<IExpectedApprover[]> {
  const candidates = buildAccionTitleCandidates(solicitud);
  if (!candidates.length) {
    return [];
  }

  const acciones = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(LIST_ACCIONES)}')/items` +
    `?$select=Id,Title,Revisor/Id,Revisor/Title,Revisor/EMail` +
    `&$expand=Revisor` +
    `&$top=5000`
  );

  const result: IExpectedApprover[] = [];
  for (let i = 0; i < acciones.length; i++) {
    if (!accionCoincide(String(acciones[i]?.Title || ''), candidates, solicitud)) {
      continue;
    }

    const personas = getPeople(acciones[i]?.Revisor);
    for (let j = 0; j < personas.length; j++) {
      result.push({
        ...personas[j],
        tipo: 'Accion',
        origen: String(acciones[i]?.Title || '').trim(),
        flag: 'ImpactadoPorAccion'
      });
    }
  }

  return result;
}

function esAltaDocumentoNoVigente(solicitud: ISolicitudItem): boolean {
  const accion = normalizeKey(solicitud?.Accion || '');
  const estado = normalizeKey(solicitud?.Estado?.Title || '');

  return isAltaDocumentos(accion) && estado !== 'vigente';
}

function isAltaDocumentos(accionKey: string): boolean {
  return accionKey === 'alta de documentos' || (accionKey.indexOf('alta') !== -1 && accionKey.indexOf('documento') !== -1);
}

function buildAccionTitleCandidates(solicitud: ISolicitudItem): string[] {
  const accion = String(solicitud.Accion || '').trim();
  const accionKey = normalizeKey(accion);
  const candidates: string[] = [];

  if (accion) {
    candidates.push(accion);
  }

  if (normalizeKey(solicitud.TipoDocumento?.Title || '') === 'procedimientos') {
    if (accionKey.indexOf('alta') !== -1) {
      candidates.push('Alta de Nuevos Procedimientos');
      candidates.push('Alta de Procedimientos');
    }
  }

  return candidates.filter((value, index, arr) => value && arr.indexOf(value) === index);
}

function accionCoincide(itemTitle: string, candidates: string[], solicitud: ISolicitudItem): boolean {
  const itemKey = normalizeKey(itemTitle);
  const accionKey = normalizeKey(solicitud.Accion || '');

  for (let i = 0; i < candidates.length; i++) {
    if (itemKey === normalizeKey(candidates[i])) {
      return true;
    }
  }

  return itemKey.indexOf('procedimiento') !== -1 && accionKey.indexOf('alta') !== -1 && itemKey.indexOf('alta') !== -1;
}

function buildAprobadoresBySolicitud(aprobadores: IAprobadorSolicitudItem[]): Map<number, IAprobadorSolicitudItem[]> {
  const result = new Map<number, IAprobadorSolicitudItem[]>();

  for (let i = 0; i < aprobadores.length; i++) {
    const solicitudId = Number(aprobadores[i]?.SolicitudId || 0);
    if (!solicitudId) {
      continue;
    }

    if (!result.has(solicitudId)) {
      result.set(solicitudId, []);
    }

    result.get(solicitudId)!.push(aprobadores[i]);
  }

  return result;
}

function getPeople(raw: any): Array<{ id: number; title: string; email: string; }> {
  if (!raw) {
    return [];
  }

  const values = Array.isArray(raw)
    ? raw
    : (raw && Array.isArray(raw.results) ? raw.results : [raw]);
  const seen = new Set<number>();
  const result: Array<{ id: number; title: string; email: string; }> = [];

  for (let i = 0; i < values.length; i++) {
    const id = Number(values[i]?.Id || values[i] || 0);
    if (!id || seen.has(id)) {
      continue;
    }

    seen.add(id);
    result.push({
      id,
      title: String(values[i]?.Title || '').trim(),
      email: String(values[i]?.EMail || '').trim()
    });
  }

  return result;
}

function dedupeExpectedApprovers(items: IExpectedApprover[]): IExpectedApprover[] {
  const seen = new Set<string>();
  const result: IExpectedApprover[] = [];

  for (let i = 0; i < items.length; i++) {
    const key = `${items[i].tipo}|${items[i].id}|${items[i].flag}|${normalizeKey(items[i].origen)}`;
    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    result.push(items[i]);
  }

  return result;
}

function getAprobadorId(registro: IAprobadorSolicitudItem): number {
  return Number(registro?.AprobadorId || registro?.Aprobador?.Id || 0);
}

function buildMotivoFaltante(expected: IExpectedApprover, registro: IAprobadorSolicitudItem | null, totalRegistrosSolicitud: number): string {
  if (!totalRegistrosSolicitud) {
    return `Falta aprobador por ${expected.tipo}. La solicitud no tiene registros en Aprobadores por Solicitudes.`;
  }

  if (!registro) {
    return `Falta aprobador por ${expected.tipo}. No existe registro para el aprobador esperado.`;
  }

  const problemas: string[] = [];
  if (normalizeKey(registro.Rol || '') !== normalizeKey(ROL_REVISOR_IMPACTADO)) {
    problemas.push(`el Rol actual no es "${ROL_REVISOR_IMPACTADO}"`);
  }

  if (!isTruthyField((registro as any)[expected.flag])) {
    problemas.push(`no tiene marcada la columna ${expected.flag}`);
  }

  return `Falta aprobador por ${expected.tipo}. Existe registro para el aprobador esperado, pero ${problemas.join(' y ')}.`;
}

function buildReportRow(
  solicitud: ISolicitudItem,
  aprobador: IAprobadorSolicitudItem | null,
  expected: IExpectedApprover | null,
  motivo: string
): IAltaDocumentoAprobadorIncompletoRow {
  return {
    SolicitudId: Number(solicitud.Id || 0),
    SolicitudTitulo: String(solicitud.Title || '').trim(),
    NombreDocumento: String(solicitud.NombreDocumento || '').trim(),
    CodigoDocumento: String(solicitud.CodigoDocumento || '').trim(),
    VersionDocumento: String(solicitud.VersionDocumento || '').trim(),
    TipoDocumento: String(solicitud.TipoDocumento?.Title || '').trim(),
    MotivoSolicitud: String(solicitud.Motivo?.Title || '').trim(),
    AreasImpactadas: formatAreasImpactadas(solicitud),
    AccionSolicitud: String(solicitud.Accion || '').trim(),
    EstadoSolicitud: String(solicitud.Estado?.Title || '').trim(),
    TipoAprobadorEsperado: String(expected?.tipo || '').trim(),
    OrigenAprobadorEsperado: String(expected?.origen || '').trim(),
    AprobadorEsperadoId: Number(expected?.id || 0) || '',
    AprobadorEsperadoNombre: String(expected?.title || '').trim(),
    AprobadorEsperadoEmail: String(expected?.email || '').trim(),
    RegistroAprobadorId: Number(aprobador?.Id || 0) || '',
    Rol: String(aprobador?.Rol || '').trim(),
    AprobadorId: Number(aprobador?.AprobadorId || aprobador?.Aprobador?.Id || 0) || '',
    AprobadorNombre: String(aprobador?.Aprobador?.Title || '').trim(),
    AprobadorEmail: String(aprobador?.Aprobador?.EMail || '').trim(),
    ImpactadoPorArea: formatBooleanField(aprobador?.ImpactadoPorArea),
    ImpactadoPorMotivo: formatBooleanField(aprobador?.ImpactadoPorMotivo),
    ImpactadoPorAccion: formatBooleanField(aprobador?.ImpactadoPorAccion),
    MotivoIncompleto: motivo,
    CreatedSolicitud: String(solicitud.Created || '').trim(),
    ModifiedSolicitud: String(solicitud.Modified || '').trim(),
    CreatedAprobador: String(aprobador?.Created || '').trim(),
    ModifiedAprobador: String(aprobador?.Modified || '').trim()
  };
}

function formatAreasImpactadas(solicitud: ISolicitudItem): string {
  return (solicitud.AreasImpactadas || [])
    .map((area) => String(area?.Title || '').trim())
    .filter(Boolean)
    .join(' | ');
}

function formatBooleanField(value: any): string {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  return isTruthyField(value) ? 'Si' : 'No';
}

function isTruthyField(value: any): boolean {
  if (typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'number') {
    return value === 1;
  }

  const normalized = normalizeKey(value);
  return normalized === '1' || normalized === 'true' || normalized === 'si' || normalized === 'sí' || normalized === 'yes';
}

function normalizeKey(value: any): string {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function compareRows(a: IAltaDocumentoAprobadorIncompletoRow, b: IAltaDocumentoAprobadorIncompletoRow): number {
  const solicitud = Number(a.SolicitudId || 0) - Number(b.SolicitudId || 0);
  if (solicitud !== 0) {
    return solicitud;
  }

  return Number(a.RegistroAprobadorId || 0) - Number(b.RegistroAprobadorId || 0);
}
