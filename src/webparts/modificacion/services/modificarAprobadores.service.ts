/* eslint-disable */
// @ts-nocheck
import * as XLSX from 'xlsx';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';

import { getAllItems, spGetJson, escapeODataValue, updateListItem } from './sharepointRest.service';
import { openExcelRevisionSession } from '../utils/modificacionExcelHelper';

type LogFn = (message: string) => void;

const LIST_SOLICITUDES = 'Solicitudes';
const LIST_MOTIVOS = 'Motivos';
const LIST_APROBADORES = 'Aprobadores por Solicitudes';
const LIST_ACCIONES = 'Acciones';
const LIST_AREAS_NEGOCIO = 'Áreas de Negocio';
const ROL_REVISOR_IMPACTADO = 'Revisor Impactado';
const ESTADOS_BAJA_MOTIVO_PERMITIDOS = new Set<string>([
  'observado',
  'en edicion',
  'rechazado',
  'enviado a revision de calidad',
  'en revision de calidad'
]);

interface ISolicitudItem {
  Id: number;
  Title?: string;
  MotivoId?: number;
  Motivo?: {
    Title?: string;
  };
  Accion?: string;
  EstadoId?: number;
  EsVersionActualDocumento?: boolean | string | number;
  Estado?: {
    Title?: string;
  };
  EstadoAnteriorVencimiento?: string | {
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
}

interface IAprobadorSolicitudItem {
  Id: number;
  SolicitudId?: number;
  Rol?: string;
  ImpactadoPorArea?: boolean | string | number;
  ImpactadoPorMotivo?: boolean | string | number;
  ImpactadoPorAccion?: boolean | string | number;
  AprobadorId?: number;
  Aprobador?: {
    Id?: number;
    Title?: string;
    EMail?: string;
  };
}

interface IImpactoCalculado {
  porArea: Set<number>;
  porMotivo: Set<number>;
  porAccion: Set<number>;
}

export interface IModificarAprobadoresResultado {
  blob: Blob;
  fileName: string;
  totalEncontrados: number;
  totalCambiarian: number;
  totalActualizados: number;
  totalOmitidos: number;
  totalError: number;
  totalSolicitudesEvaluadas: number;
  totalSolicitudesVigentesOmitidas: number;
}

export interface IModificarAprobadoresRollbackResultado {
  blob: Blob;
  fileName: string;
  totalFilas: number;
  totalRestaurados: number;
  totalOmitidos: number;
  totalError: number;
}

export interface IRevisoresBajaMotivoResultado {
  blob: Blob;
  fileName: string;
  totalSolicitudes: number;
  totalAprobadoresMotivoEsperados: number;
  totalCrearRegistro: number;
  totalMarcarMotivo: number;
  totalSinCambios: number;
  totalDuplicadosExistentes: number;
  totalError: number;
}

interface IModificarAprobadoresReportRow {
  RegistroAprobadorId: number | '';
  SolicitudId: number | '';
  SolicitudTitulo: string;
  SolicitudEstado: string;
  Rol: string;
  AprobadorId: number | '';
  AprobadorNombre: string;
  AprobadorEmail: string;
  AccionSolicitud: string;
  MotivoSolicitud: string;
  TipoDocumento: string;
  AreasImpactadas: string;
  ImpactadoPorAreaActual: string;
  ImpactadoPorAreaCalculado: string;
  ImpactadoPorMotivoActual: string;
  ImpactadoPorMotivoCalculado: string;
  ImpactadoPorAccionActual: string;
  ImpactadoPorAccionCalculado: string;
  CambiariaImpactadoPorArea: string;
  CambiariaImpactadoPorMotivo: string;
  CambiariaImpactadoPorAccion: string;
  ResultadoEnsayo: string;
  Motivo: string;
}

interface IRevisoresBajaMotivoReportRow {
  SolicitudId: number | '';
  SolicitudTitulo: string;
  SolicitudEstado: string;
  EstadoAnteriorVencimiento: string;
  AccionSolicitud: string;
  MotivoId: number | '';
  MotivoSolicitud: string;
  AprobadorId: number | '';
  AprobadorNombre: string;
  AprobadorEmail: string;
  RegistroAprobadorId: number | '';
  RegistrosMismoAprobador: string;
  Rol: string;
  ImpactadoPorAreaActual: string;
  ImpactadoPorMotivoActual: string;
  ImpactadoPorAccionActual: string;
  AccionSugerida: string;
  ImpactadoPorMotivoSugerido: string;
  CrearRegistroSugerido: string;
  EvitaDuplicado: string;
  Observacion: string;
}

const reportHeaders = [
  'RegistroAprobadorId',
  'SolicitudId',
  'SolicitudTitulo',
  'SolicitudEstado',
  'Rol',
  'AprobadorId',
  'AprobadorNombre',
  'AprobadorEmail',
  'AccionSolicitud',
  'MotivoSolicitud',
  'TipoDocumento',
  'AreasImpactadas',
  'ImpactadoPorAreaActual',
  'ImpactadoPorAreaCalculado',
  'ImpactadoPorMotivoActual',
  'ImpactadoPorMotivoCalculado',
  'ImpactadoPorAccionActual',
  'ImpactadoPorAccionCalculado',
  'CambiariaImpactadoPorArea',
  'CambiariaImpactadoPorMotivo',
  'CambiariaImpactadoPorAccion',
  'ResultadoEnsayo',
  'Motivo'
];

const revisoresBajaMotivoHeaders = [
  'SolicitudId',
  'SolicitudTitulo',
  'SolicitudEstado',
  'EstadoAnteriorVencimiento',
  'AccionSolicitud',
  'MotivoId',
  'MotivoSolicitud',
  'AprobadorId',
  'AprobadorNombre',
  'AprobadorEmail',
  'RegistroAprobadorId',
  'RegistrosMismoAprobador',
  'Rol',
  'ImpactadoPorAreaActual',
  'ImpactadoPorMotivoActual',
  'ImpactadoPorAccionActual',
  'AccionSugerida',
  'ImpactadoPorMotivoSugerido',
  'CrearRegistroSugerido',
  'EvitaDuplicado',
  'Observacion'
];

function normalizeKey(value: any): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function isTruthyField(value: any): boolean {
  if (value === true || value === 1) {
    return true;
  }

  const normalized = normalizeKey(value);
  return normalized === 'true' || normalized === '1' || normalized === 'si' || normalized === 'yes';
}

function formatBoolean(value: boolean): string {
  return value ? 'TRUE' : 'FALSE';
}

function formatAreasImpactadas(solicitud: ISolicitudItem | null | undefined): string {
  return (solicitud?.AreasImpactadas || [])
    .map((area) => String(area?.Title || '').trim())
    .filter(Boolean)
    .join('/');
}

function getEstadoAnteriorVencimientoTitle(solicitud: ISolicitudItem | null | undefined): string {
  const estadoAnterior = solicitud?.EstadoAnteriorVencimiento;

  if (typeof estadoAnterior === 'string') {
    return estadoAnterior;
  }

  return String(estadoAnterior?.Title || '');
}

function debeProcesarBajaMotivo(solicitud: ISolicitudItem): boolean {
  const accionKey = normalizeKey(solicitud?.Accion || '');
  if (accionKey !== 'baja de documentos') {
    return false;
  }

  const estadoKey = normalizeKey(solicitud?.Estado?.Title || '');
  if (ESTADOS_BAJA_MOTIVO_PERMITIDOS.has(estadoKey)) {
    return true;
  }

  if (estadoKey !== 'vencido') {
    return false;
  }

  const estadoAnteriorKey = normalizeKey(getEstadoAnteriorVencimientoTitle(solicitud));
  return ESTADOS_BAJA_MOTIVO_PERMITIDOS.has(estadoAnteriorKey);
}

function autoFitColumns(rows: IModificarAprobadoresReportRow[]): Array<{ wch: number; }> {
  const widths = reportHeaders.map((header) => ({ wch: header.length + 2 }));

  rows.forEach((row) => {
    reportHeaders.forEach((header, index) => {
      const value = String((row as any)[header] ?? '');
      widths[index].wch = Math.min(Math.max(widths[index].wch, value.length + 2), 60);
    });
  });

  return widths;
}

function buildModificarAprobadoresWorkbook(rows: IModificarAprobadoresReportRow[], prefix?: string): { blob: Blob; fileName: string; } {
  const safeRows = Array.isArray(rows) ? rows : [];
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(safeRows, { header: reportHeaders });

  worksheet['!cols'] = autoFitColumns(safeRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'ModificarAprobadores');

  const output = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array'
  });

  const now = new Date();
  const pad = (value: number): string => String(value).padStart(2, '0');
  const fileName =
    `${prefix || 'Resultado_ModificarAprobadores'}_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_` +
    `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.xlsx`;

  return {
    blob: new Blob([output], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }),
    fileName
  };
}

function autoFitRevisoresBajaMotivoColumns(rows: IRevisoresBajaMotivoReportRow[]): Array<{ wch: number; }> {
  const widths = revisoresBajaMotivoHeaders.map((header) => ({ wch: header.length + 2 }));

  rows.forEach((row) => {
    revisoresBajaMotivoHeaders.forEach((header, index) => {
      const value = String((row as any)[header] ?? '');
      widths[index].wch = Math.min(Math.max(widths[index].wch, value.length + 2), 70);
    });
  });

  return widths;
}

function buildRevisoresBajaMotivoWorkbook(rows: IRevisoresBajaMotivoReportRow[]): { blob: Blob; fileName: string; } {
  const safeRows = Array.isArray(rows) ? rows : [];
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(safeRows, { header: revisoresBajaMotivoHeaders });

  worksheet['!cols'] = autoFitRevisoresBajaMotivoColumns(safeRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, 'RevisoresBajaMotivo');

  const output = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array'
  });

  const now = new Date();
  const pad = (value: number): string => String(value).padStart(2, '0');
  const fileName =
    `Revisores_Impactados_Baja_Motivo_DryRun_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_` +
    `${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}.xlsx`;

  return {
    blob: new Blob([output], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }),
    fileName
  };
}

function findHeaderIndex(headers: any[], headerName: string): number {
  const expected = normalizeKey(headerName);
  for (let i = 0; i < headers.length; i++) {
    if (normalizeKey(headers[i]) === expected) {
      return i;
    }
  }

  return -1;
}

function getCellValue(row: any[], headerMap: Map<string, number>, headerName: string): any {
  const index = headerMap.get(headerName);
  return index === undefined || index < 0 ? '' : row[index];
}

function buildHeaderMap(headers: any[]): Map<string, number> {
  const result = new Map<string, number>();
  for (let i = 0; i < reportHeaders.length; i++) {
    result.set(reportHeaders[i], findHeaderIndex(headers, reportHeaders[i]));
  }

  return result;
}

function parseReportRow(row: any[], headerMap: Map<string, number>): IModificarAprobadoresReportRow {
  return {
    RegistroAprobadorId: Number(getCellValue(row, headerMap, 'RegistroAprobadorId') || 0) || '',
    SolicitudId: Number(getCellValue(row, headerMap, 'SolicitudId') || 0) || '',
    SolicitudTitulo: String(getCellValue(row, headerMap, 'SolicitudTitulo') || ''),
    SolicitudEstado: String(getCellValue(row, headerMap, 'SolicitudEstado') || ''),
    Rol: String(getCellValue(row, headerMap, 'Rol') || ''),
    AprobadorId: Number(getCellValue(row, headerMap, 'AprobadorId') || 0) || '',
    AprobadorNombre: String(getCellValue(row, headerMap, 'AprobadorNombre') || ''),
    AprobadorEmail: String(getCellValue(row, headerMap, 'AprobadorEmail') || ''),
    AccionSolicitud: String(getCellValue(row, headerMap, 'AccionSolicitud') || ''),
    MotivoSolicitud: String(getCellValue(row, headerMap, 'MotivoSolicitud') || ''),
    TipoDocumento: String(getCellValue(row, headerMap, 'TipoDocumento') || ''),
    AreasImpactadas: String(getCellValue(row, headerMap, 'AreasImpactadas') || ''),
    ImpactadoPorAreaActual: String(getCellValue(row, headerMap, 'ImpactadoPorAreaActual') || ''),
    ImpactadoPorAreaCalculado: String(getCellValue(row, headerMap, 'ImpactadoPorAreaCalculado') || ''),
    ImpactadoPorMotivoActual: String(getCellValue(row, headerMap, 'ImpactadoPorMotivoActual') || ''),
    ImpactadoPorMotivoCalculado: String(getCellValue(row, headerMap, 'ImpactadoPorMotivoCalculado') || ''),
    ImpactadoPorAccionActual: String(getCellValue(row, headerMap, 'ImpactadoPorAccionActual') || ''),
    ImpactadoPorAccionCalculado: String(getCellValue(row, headerMap, 'ImpactadoPorAccionCalculado') || ''),
    CambiariaImpactadoPorArea: String(getCellValue(row, headerMap, 'CambiariaImpactadoPorArea') || ''),
    CambiariaImpactadoPorMotivo: String(getCellValue(row, headerMap, 'CambiariaImpactadoPorMotivo') || ''),
    CambiariaImpactadoPorAccion: String(getCellValue(row, headerMap, 'CambiariaImpactadoPorAccion') || ''),
    ResultadoEnsayo: String(getCellValue(row, headerMap, 'ResultadoEnsayo') || ''),
    Motivo: String(getCellValue(row, headerMap, 'Motivo') || '')
  };
}

async function actualizarChecksAprobador(
  context: WebPartContext,
  webUrl: string,
  registroAprobadorId: number,
  values: {
    ImpactadoPorArea: boolean;
    ImpactadoPorMotivo: boolean;
    ImpactadoPorAccion: boolean;
  }
): Promise<void> {
  await updateListItem(context, webUrl, LIST_APROBADORES, registroAprobadorId, values);
}

function getPersonIds(raw: any): number[] {
  if (!raw) {
    return [];
  }

  if (Array.isArray(raw)) {
    return raw.map((item: any) => Number(item?.Id || item)).filter((id: number) => Number.isFinite(id) && id > 0);
  }

  if (raw && Array.isArray(raw.results)) {
    return raw.results.map((item: any) => Number(item?.Id || item)).filter((id: number) => Number.isFinite(id) && id > 0);
  }

  const id = Number(raw?.Id || raw);
  return Number.isFinite(id) && id > 0 ? [id] : [];
}

function addToSet(target: Set<number>, ids: number[]): void {
  for (let i = 0; i < ids.length; i++) {
    target.add(ids[i]);
  }
}

function writeInfo(log: LogFn, message: string): void {
  console.log(message);
  log(message);
}

function writeError(log: LogFn, message: string): void {
  console.error(message);
  log(message);
}

async function obtenerSolicitud(
  context: WebPartContext,
  webUrl: string,
  solicitudId: number
): Promise<ISolicitudItem> {
  return spGetJson<ISolicitudItem>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(LIST_SOLICITUDES)}')/items(${solicitudId})` +
    `?$select=Id,Title,MotivoId,Motivo/Title,Accion,Estado/Title,TipoDocumento/Id,TipoDocumento/Title,AreasImpactadas/Id,AreasImpactadas/Title` +
    `&$expand=Motivo,Estado,TipoDocumento,AreasImpactadas`
  );
}

async function obtenerTodosAprobadoresImpactados(
  context: WebPartContext,
  webUrl: string
): Promise<IAprobadorSolicitudItem[]> {
  const filter = `Rol eq '${ROL_REVISOR_IMPACTADO.replace(/'/g, "''")}'`;
  const url =
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(LIST_APROBADORES)}')/items` +
    `?$select=Id,SolicitudId,Rol,AprobadorId,Aprobador/Id,Aprobador/Title,Aprobador/EMail,ImpactadoPorArea,ImpactadoPorMotivo,ImpactadoPorAccion` +
    `&$expand=Aprobador` +
    `&$filter=${encodeURIComponent(filter)}` +
    `&$top=5000`;

  return getAllItems<IAprobadorSolicitudItem>(context, url);
}

async function obtenerSolicitudesBajaDocumentosEnProceso(
  context: WebPartContext,
  webUrl: string
): Promise<ISolicitudItem[]> {
  const solicitudes = await getAllItems<ISolicitudItem>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(LIST_SOLICITUDES)}')/items` +
    `?$select=Id,Title,MotivoId,Motivo/Title,Accion,EstadoId,Estado/Title,EstadoAnteriorVencimiento,EsVersionActualDocumento` +
    `&$expand=Motivo,Estado` +
    `&$top=5000`
  );

  return solicitudes.filter(debeProcesarBajaMotivo);
}

function buildBajaDocumentosStats(solicitudes: ISolicitudItem[]): {
  total: number;
  totalBaja: number;
  totalBajaPermitidas: number;
  totalBajaOmitidasPorEstado: number;
  estadosBaja: string;
} {
  const estados = new Map<string, number>();
  let totalBaja = 0;
  let totalBajaPermitidas = 0;
  let totalBajaOmitidasPorEstado = 0;

  for (let i = 0; i < solicitudes.length; i++) {
    const accionKey = normalizeKey(solicitudes[i]?.Accion || '');
    if (accionKey !== 'baja de documentos') {
      continue;
    }

    totalBaja++;
    const estadoTitle = String(solicitudes[i]?.Estado?.Title || '(sin estado)').trim() || '(sin estado)';
    estados.set(estadoTitle, (estados.get(estadoTitle) || 0) + 1);

    if (debeProcesarBajaMotivo(solicitudes[i])) {
      totalBajaPermitidas++;
    } else {
      totalBajaOmitidasPorEstado++;
    }
  }

  const estadosBaja = Array.from(estados.entries())
    .sort((a, b) => b[1] - a[1])
    .map(([estado, count]) => `${estado}: ${count}`)
    .join(' | ');

  return {
    total: solicitudes.length,
    totalBaja,
    totalBajaPermitidas,
    totalBajaOmitidasPorEstado,
    estadosBaja
  };
}

function buildAprobadoresBySolicitud(
  aprobadores: IAprobadorSolicitudItem[]
): Map<number, IAprobadorSolicitudItem[]> {
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

function getBestExistingAprobador(items: IAprobadorSolicitudItem[]): IAprobadorSolicitudItem | null {
  if (!items.length) {
    return null;
  }

  const withMotivo = items.filter((item) => isTruthyField(item.ImpactadoPorMotivo));
  if (withMotivo.length) {
    return withMotivo[0];
  }

  const withArea = items.filter((item) => isTruthyField(item.ImpactadoPorArea));
  if (withArea.length) {
    return withArea[0];
  }

  return items[0];
}

function buildRevisoresBajaMotivoRow(params: {
  solicitud: ISolicitudItem;
  aprobadorId: number;
  aprobadorNombre: string;
  aprobadorEmail: string;
  existingItems: IAprobadorSolicitudItem[];
  accionSugerida: string;
  crearRegistro: boolean;
  evitaDuplicado: boolean;
  observacion: string;
}): IRevisoresBajaMotivoReportRow {
  const best = getBestExistingAprobador(params.existingItems);

  return {
    SolicitudId: Number(params.solicitud.Id || 0) || '',
    SolicitudTitulo: String(params.solicitud.Title || ''),
    SolicitudEstado: String(params.solicitud.Estado?.Title || ''),
    EstadoAnteriorVencimiento: getEstadoAnteriorVencimientoTitle(params.solicitud),
    AccionSolicitud: String(params.solicitud.Accion || ''),
    MotivoId: Number(params.solicitud.MotivoId || 0) || '',
    MotivoSolicitud: String(params.solicitud.Motivo?.Title || ''),
    AprobadorId: params.aprobadorId || '',
    AprobadorNombre: params.aprobadorNombre || String(best?.Aprobador?.Title || ''),
    AprobadorEmail: params.aprobadorEmail || String(best?.Aprobador?.EMail || ''),
    RegistroAprobadorId: Number(best?.Id || 0) || '',
    RegistrosMismoAprobador: params.existingItems.map((item) => String(item.Id || '')).filter(Boolean).join(', '),
    Rol: ROL_REVISOR_IMPACTADO,
    ImpactadoPorAreaActual: formatBoolean(isTruthyField(best?.ImpactadoPorArea)),
    ImpactadoPorMotivoActual: formatBoolean(isTruthyField(best?.ImpactadoPorMotivo)),
    ImpactadoPorAccionActual: formatBoolean(isTruthyField(best?.ImpactadoPorAccion)),
    AccionSugerida: params.accionSugerida,
    ImpactadoPorMotivoSugerido: formatBoolean(true),
    CrearRegistroSugerido: params.crearRegistro ? 'Si' : 'No',
    EvitaDuplicado: params.evitaDuplicado ? 'Si' : 'No',
    Observacion: params.observacion
  };
}

async function obtenerImpactadosPorMotivo(
  context: WebPartContext,
  webUrl: string,
  motivoId: number | undefined
): Promise<Set<number>> {
  const result = new Set<number>();
  if (!motivoId) {
    return result;
  }

  const motivo = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(LIST_MOTIVOS)}')/items(${motivoId})` +
    `?$select=Id,Aprobadores/Id,Aprobadores/Title,Aprobadores/EMail&$expand=Aprobadores`
  );

  addToSet(result, getPersonIds(motivo?.Aprobadores));
  return result;
}

async function obtenerAprobadoresMotivoDetalle(
  context: WebPartContext,
  webUrl: string,
  motivoId: number | undefined
): Promise<Array<{ id: number; title: string; email: string; }>> {
  if (!motivoId) {
    return [];
  }

  const motivo = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(LIST_MOTIVOS)}')/items(${motivoId})` +
    `?$select=Id,Aprobadores/Id,Aprobadores/Title,Aprobadores/EMail&$expand=Aprobadores`
  );

  const raw = motivo?.Aprobadores;
  const values = Array.isArray(raw)
    ? raw
    : (raw && Array.isArray(raw.results) ? raw.results : []);
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
      title: String(values[i]?.Title || ''),
      email: String(values[i]?.EMail || '')
    });
  }

  return result;
}

async function obtenerImpactadosPorArea(
  context: WebPartContext,
  webUrl: string,
  solicitud: ISolicitudItem
): Promise<Set<number>> {
  const result = new Set<number>();
  const areasIds = (solicitud.AreasImpactadas || [])
    .map((area) => Number(area?.Id || 0))
    .filter((id) => Number.isFinite(id) && id > 0);

  if (!areasIds.length) {
    return result;
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

  for (let i = 0; i < areas.length; i++) {
    addToSet(result, getPersonIds(areas[i]?.Gerente));
  }

  return result;
}

function buildAccionTitleCandidates(solicitud: ISolicitudItem): string[] {
  const accion = String(solicitud.Accion || '').trim();
  const accionKey = normalizeKey(accion);
  const tipoDocumentoKey = normalizeKey(solicitud.TipoDocumento?.Title || '');
  const candidates: string[] = [];

  if (accion) {
    candidates.push(accion);
  }

  if (tipoDocumentoKey === 'procedimientos') {
    if (accionKey.indexOf('alta') !== -1) {
      candidates.push('Alta de Nuevos Procedimientos');
      candidates.push('Alta de Procedimientos');
    }

    if (accionKey.indexOf('actualizacion') !== -1 || accionKey.indexOf('actualización') !== -1) {
      candidates.push('Actualización de Procedimientos');
      candidates.push('Actualizacion de Procedimientos');
    }

    if (!candidates.length) {
      candidates.push('Alta de Nuevos Procedimientos');
    }
  }

  return candidates.filter((value, index, arr) => value && arr.indexOf(value) === index);
}

function accionCoincide(itemTitle: string, candidates: string[], solicitud: ISolicitudItem): boolean {
  const itemKey = normalizeKey(itemTitle);
  const accionKey = normalizeKey(solicitud.Accion || '');
  const tipoDocumentoKey = normalizeKey(solicitud.TipoDocumento?.Title || '');

  for (let i = 0; i < candidates.length; i++) {
    if (itemKey === normalizeKey(candidates[i])) {
      return true;
    }
  }

  if (tipoDocumentoKey === 'procedimientos' && itemKey.indexOf('procedimiento') !== -1) {
    if (accionKey.indexOf('alta') !== -1 && itemKey.indexOf('alta') !== -1) {
      return true;
    }

    if (accionKey.indexOf('actualizacion') !== -1 && itemKey.indexOf('actualizacion') !== -1) {
      return true;
    }
  }

  return false;
}

async function obtenerImpactadosPorAccion(
  context: WebPartContext,
  webUrl: string,
  solicitud: ISolicitudItem
): Promise<Set<number>> {
  const result = new Set<number>();
  const candidates = buildAccionTitleCandidates(solicitud);

  if (!candidates.length) {
    return result;
  }

  const acciones = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(LIST_ACCIONES)}')/items` +
    `?$select=Id,Title,Revisor/Id,Revisor/Title,Revisor/EMail` +
    `&$expand=Revisor` +
    `&$top=5000`
  );

  for (let i = 0; i < acciones.length; i++) {
    if (accionCoincide(String(acciones[i]?.Title || ''), candidates, solicitud)) {
      addToSet(result, getPersonIds(acciones[i]?.Revisor));
    }
  }

  return result;
}

async function calcularImpactos(
  context: WebPartContext,
  webUrl: string,
  solicitud: ISolicitudItem
): Promise<IImpactoCalculado> {
  const porMotivo = await obtenerImpactadosPorMotivo(context, webUrl, Number(solicitud.MotivoId || 0) || undefined);
  const porArea = await obtenerImpactadosPorArea(context, webUrl, solicitud);
  const porAccion = await obtenerImpactadosPorAccion(context, webUrl, solicitud);

  return {
    porArea,
    porMotivo,
    porAccion
  };
}

function getAprobadorId(item: IAprobadorSolicitudItem): number {
  return Number(item.AprobadorId || item.Aprobador?.Id || 0);
}

export async function generarReporteRevisoresImpactadosBajaMotivo(params: {
  context: WebPartContext;
  log?: LogFn;
}): Promise<IRevisoresBajaMotivoResultado> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;

  writeInfo(log, '🔎 Revisores Baja por Motivo | Consultando solicitudes con Acción="Baja de documentos" y estados permitidos...');
  const todasSolicitudes = await getAllItems<ISolicitudItem>(
    params.context,
    `${webUrl}/_api/web/lists/getbytitle('${escapeODataValue(LIST_SOLICITUDES)}')/items` +
    `?$select=Id,Title,MotivoId,Motivo/Title,Accion,EstadoId,Estado/Title,EstadoAnteriorVencimiento,EsVersionActualDocumento` +
    `&$expand=Motivo,Estado` +
    `&$top=5000`
  );
  const stats = buildBajaDocumentosStats(todasSolicitudes);
  writeInfo(
    log,
    `📊 Revisores Baja por Motivo | Solicitudes leídas=${stats.total} | ` +
    `Baja de documentos=${stats.totalBaja} | Baja permitidas=${stats.totalBajaPermitidas} | ` +
    `Baja omitidas por estado=${stats.totalBajaOmitidasPorEstado}`
  );
  if (stats.estadosBaja) {
    writeInfo(log, `📊 Revisores Baja por Motivo | Estados encontrados en Baja de documentos: ${stats.estadosBaja}`);
  }

  const solicitudes = todasSolicitudes.filter(debeProcesarBajaMotivo);
  writeInfo(log, `📋 Revisores Baja por Motivo | Solicitudes candidatas: ${solicitudes.length}`);

  writeInfo(log, `👥 Revisores Baja por Motivo | Consultando registros con Rol="${ROL_REVISOR_IMPACTADO}"...`);
  const aprobadores = await obtenerTodosAprobadoresImpactados(params.context, webUrl);
  const aprobadoresBySolicitud = buildAprobadoresBySolicitud(aprobadores);

  let totalAprobadoresMotivoEsperados = 0;
  let totalCrearRegistro = 0;
  let totalMarcarMotivo = 0;
  let totalSinCambios = 0;
  let totalDuplicadosExistentes = 0;
  let totalError = 0;
  const reportRows: IRevisoresBajaMotivoReportRow[] = [];

  for (let i = 0; i < solicitudes.length; i++) {
    const solicitud = solicitudes[i];
    const solicitudId = Number(solicitud.Id || 0);
    const actuales = aprobadoresBySolicitud.get(solicitudId) || [];

    if (!Number(solicitud.MotivoId || 0)) {
      totalError++;
      reportRows.push(buildRevisoresBajaMotivoRow({
        solicitud,
        aprobadorId: 0,
        aprobadorNombre: '',
        aprobadorEmail: '',
        existingItems: [],
        accionSugerida: 'ERROR',
        crearRegistro: false,
        evitaDuplicado: false,
        observacion: 'La solicitud no tiene MotivoId; no se puede consultar aprobadores asociados al motivo.'
      }));
      continue;
    }

    let aprobadoresMotivo: Array<{ id: number; title: string; email: string; }> = [];
    try {
      aprobadoresMotivo = await obtenerAprobadoresMotivoDetalle(params.context, webUrl, Number(solicitud.MotivoId || 0));
    } catch (error) {
      totalError++;
      reportRows.push(buildRevisoresBajaMotivoRow({
        solicitud,
        aprobadorId: 0,
        aprobadorNombre: '',
        aprobadorEmail: '',
        existingItems: [],
        accionSugerida: 'ERROR',
        crearRegistro: false,
        evitaDuplicado: false,
        observacion: `No se pudo consultar la lista Motivos: ${error instanceof Error ? error.message : String(error)}`
      }));
      continue;
    }

    if (!aprobadoresMotivo.length) {
      totalError++;
      reportRows.push(buildRevisoresBajaMotivoRow({
        solicitud,
        aprobadorId: 0,
        aprobadorNombre: '',
        aprobadorEmail: '',
        existingItems: [],
        accionSugerida: 'ERROR',
        crearRegistro: false,
        evitaDuplicado: false,
        observacion: 'El motivo asociado a la solicitud no tiene aprobadores configurados.'
      }));
      continue;
    }

    const aprobadoresMotivoIds = new Set<number>();
    for (let j = 0; j < aprobadoresMotivo.length; j++) {
      aprobadoresMotivoIds.add(aprobadoresMotivo[j].id);
    }

    const actualesPorMotivo = actuales.filter((item) =>
      isTruthyField(item.ImpactadoPorMotivo) && aprobadoresMotivoIds.has(getAprobadorId(item))
    );

    if (actualesPorMotivo.length) {
      totalSinCambios++;
      reportRows.push(buildRevisoresBajaMotivoRow({
        solicitud,
        aprobadorId: getAprobadorId(actualesPorMotivo[0]),
        aprobadorNombre: String(actualesPorMotivo[0]?.Aprobador?.Title || ''),
        aprobadorEmail: String(actualesPorMotivo[0]?.Aprobador?.EMail || ''),
        existingItems: actualesPorMotivo,
        accionSugerida: actualesPorMotivo.length > 1 ? 'SIN_CAMBIOS_CON_DUPLICADOS_EXISTENTES' : 'SIN_CAMBIOS',
        crearRegistro: false,
        evitaDuplicado: true,
        observacion: actualesPorMotivo.length > 1
          ? 'Ya existe al menos un aprobador configurado en el motivo con ImpactadoPorMotivo marcado. Se detectaron varios registros; el dry run no modifica ni elimina duplicados.'
          : 'Ya existe al menos un aprobador configurado en el motivo con ImpactadoPorMotivo marcado para la solicitud.'
      }));

      if (actualesPorMotivo.length > 1) {
        totalDuplicadosExistentes++;
      }

      continue;
    }

    totalAprobadoresMotivoEsperados += aprobadoresMotivo.length;
    for (let j = 0; j < aprobadoresMotivo.length; j++) {
      const aprobadorMotivo = aprobadoresMotivo[j];
      const existentesMismoAprobador = actuales.filter((item) => getAprobadorId(item) === aprobadorMotivo.id);
      const existeConMotivo = existentesMismoAprobador.some((item) => isTruthyField(item.ImpactadoPorMotivo));
      const existePorArea = existentesMismoAprobador.some((item) => isTruthyField(item.ImpactadoPorArea));

      if (existentesMismoAprobador.length > 1) {
        totalDuplicadosExistentes++;
      }

      if (existeConMotivo) {
        totalSinCambios++;
        reportRows.push(buildRevisoresBajaMotivoRow({
          solicitud,
          aprobadorId: aprobadorMotivo.id,
          aprobadorNombre: aprobadorMotivo.title,
          aprobadorEmail: aprobadorMotivo.email,
          existingItems: existentesMismoAprobador,
          accionSugerida: 'SIN_CAMBIOS',
          crearRegistro: false,
          evitaDuplicado: true,
          observacion: 'El aprobador del motivo ya está registrado con ImpactadoPorMotivo marcado.'
        }));
        continue;
      }

      if (existentesMismoAprobador.length) {
        totalMarcarMotivo++;
        reportRows.push(buildRevisoresBajaMotivoRow({
          solicitud,
          aprobadorId: aprobadorMotivo.id,
          aprobadorNombre: aprobadorMotivo.title,
          aprobadorEmail: aprobadorMotivo.email,
          existingItems: existentesMismoAprobador,
          accionSugerida: 'MARCAR_MOTIVO',
          crearRegistro: false,
          evitaDuplicado: true,
          observacion: existePorArea
            ? 'El aprobador ya existe por área; corresponde marcar ImpactadoPorMotivo en el mismo registro, sin crear duplicado.'
            : 'El aprobador ya existe para la solicitud; corresponde marcar ImpactadoPorMotivo en el registro existente, sin crear duplicado.'
        }));
        continue;
      }

      totalCrearRegistro++;
      reportRows.push(buildRevisoresBajaMotivoRow({
        solicitud,
        aprobadorId: aprobadorMotivo.id,
        aprobadorNombre: aprobadorMotivo.title,
        aprobadorEmail: aprobadorMotivo.email,
        existingItems: [],
        accionSugerida: 'CREAR_REGISTRO',
        crearRegistro: true,
        evitaDuplicado: false,
        observacion: 'No existe registro para este aprobador en la solicitud; se sugiere crear Revisor Impactado con ImpactadoPorMotivo marcado.'
      }));
    }
  }

  const workbook = buildRevisoresBajaMotivoWorkbook(reportRows);
  writeInfo(
    log,
    `📌 Revisores Baja por Motivo | Dry run final | Solicitudes=${solicitudes.length} | ` +
    `EsperadosPorMotivo=${totalAprobadoresMotivoEsperados} | Crear=${totalCrearRegistro} | ` +
    `MarcarMotivo=${totalMarcarMotivo} | SinCambios=${totalSinCambios} | DuplicadosExistentes=${totalDuplicadosExistentes} | Error=${totalError}`
  );
  writeInfo(log, `📥 Revisores Baja por Motivo | Excel generado: ${workbook.fileName}`);

  return {
    blob: workbook.blob,
    fileName: workbook.fileName,
    totalSolicitudes: solicitudes.length,
    totalAprobadoresMotivoEsperados,
    totalCrearRegistro,
    totalMarcarMotivo,
    totalSinCambios,
    totalDuplicadosExistentes,
    totalError
  };
}

function debeProcesarSolicitud(solicitud: ISolicitudItem): boolean {
  return normalizeKey(solicitud?.Estado?.Title || '') !== 'vigente';
}

async function getSolicitudFromCache(
  context: WebPartContext,
  webUrl: string,
  solicitudId: number,
  cache: Map<number, ISolicitudItem | null>
): Promise<ISolicitudItem | null> {
  if (cache.has(solicitudId)) {
    return cache.get(solicitudId) || null;
  }

  try {
    const solicitud = await obtenerSolicitud(context, webUrl, solicitudId);
    cache.set(solicitudId, solicitud);
    return solicitud;
  } catch (error) {
    cache.set(solicitudId, null);
    return null;
  }
}

async function getImpactosFromCache(
  context: WebPartContext,
  webUrl: string,
  solicitud: ISolicitudItem,
  cache: Map<number, IImpactoCalculado>
): Promise<IImpactoCalculado> {
  const solicitudId = Number(solicitud.Id || 0);
  if (cache.has(solicitudId)) {
    return cache.get(solicitudId)!;
  }

  const impactos = await calcularImpactos(context, webUrl, solicitud);
  cache.set(solicitudId, impactos);
  return impactos;
}

export async function modificarAprobadores(params: {
  context: WebPartContext;
  log?: LogFn;
}): Promise<IModificarAprobadoresResultado> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;

  writeInfo(log, `🛠️ ModificarAprobadores | Inicio del proceso`);
  writeInfo(log, `🔎 ModificarAprobadores | Consultando registros con Rol="${ROL_REVISOR_IMPACTADO}"...`);
  const aprobadores = await obtenerTodosAprobadoresImpactados(params.context, webUrl);
  writeInfo(log, `👥 ModificarAprobadores | Revisores impactados encontrados: ${aprobadores.length}`);

  let totalCambiarian = 0;
  let totalActualizados = 0;
  let totalOmitidos = 0;
  let totalError = 0;
  let totalSolicitudesVigentesOmitidas = 0;
  const solicitudesEvaluadas = new Set<number>();
  const solicitudesVigentesOmitidas = new Set<number>();
  const solicitudCache = new Map<number, ISolicitudItem | null>();
  const impactosCache = new Map<number, IImpactoCalculado>();
  const reportRows: IModificarAprobadoresReportRow[] = [];

  for (let i = 0; i < aprobadores.length; i++) {
    const item = aprobadores[i];
    const solicitudId = Number(item.SolicitudId || 0);
    const aprobadorId = getAprobadorId(item);
    const currentArea = isTruthyField(item.ImpactadoPorArea);
    const currentMotivo = isTruthyField(item.ImpactadoPorMotivo);
    const currentAccion = isTruthyField(item.ImpactadoPorAccion);

    if (!solicitudId || !aprobadorId) {
      totalOmitidos++;
      reportRows.push({
        RegistroAprobadorId: Number(item.Id || 0) || '',
        SolicitudId: solicitudId || '',
        SolicitudTitulo: '',
        SolicitudEstado: '',
        Rol: String(item.Rol || ''),
        AprobadorId: aprobadorId || '',
        AprobadorNombre: String(item.Aprobador?.Title || ''),
        AprobadorEmail: String(item.Aprobador?.EMail || ''),
        AccionSolicitud: '',
        MotivoSolicitud: '',
        TipoDocumento: '',
        AreasImpactadas: '',
        ImpactadoPorAreaActual: formatBoolean(currentArea),
        ImpactadoPorAreaCalculado: '',
        ImpactadoPorMotivoActual: formatBoolean(currentMotivo),
        ImpactadoPorMotivoCalculado: '',
        ImpactadoPorAccionActual: formatBoolean(currentAccion),
        ImpactadoPorAccionCalculado: '',
        CambiariaImpactadoPorArea: 'No',
        CambiariaImpactadoPorMotivo: 'No',
        CambiariaImpactadoPorAccion: 'No',
        ResultadoEnsayo: 'OMITIDO',
        Motivo: 'Referencia incompleta: falta SolicitudId o AprobadorId.'
      });
      writeInfo(log, `⏭️ ModificarAprobadores | Omitido por referencia incompleta | Registro=${item.Id} | Solicitud=${solicitudId} | Aprobador=${aprobadorId}`);
      continue;
    }

    const solicitud = await getSolicitudFromCache(params.context, webUrl, solicitudId, solicitudCache);
    if (!solicitud) {
      totalError++;
      reportRows.push({
        RegistroAprobadorId: Number(item.Id || 0) || '',
        SolicitudId: solicitudId,
        SolicitudTitulo: '',
        SolicitudEstado: '',
        Rol: String(item.Rol || ''),
        AprobadorId: aprobadorId,
        AprobadorNombre: String(item.Aprobador?.Title || ''),
        AprobadorEmail: String(item.Aprobador?.EMail || ''),
        AccionSolicitud: '',
        MotivoSolicitud: '',
        TipoDocumento: '',
        AreasImpactadas: '',
        ImpactadoPorAreaActual: formatBoolean(currentArea),
        ImpactadoPorAreaCalculado: '',
        ImpactadoPorMotivoActual: formatBoolean(currentMotivo),
        ImpactadoPorMotivoCalculado: '',
        ImpactadoPorAccionActual: formatBoolean(currentAccion),
        ImpactadoPorAccionCalculado: '',
        CambiariaImpactadoPorArea: 'No',
        CambiariaImpactadoPorMotivo: 'No',
        CambiariaImpactadoPorAccion: 'No',
        ResultadoEnsayo: 'ERROR',
        Motivo: 'No se encontró la solicitud asociada.'
      });
      writeError(log, `❌ ModificarAprobadores | No se encontró la solicitud asociada | Registro=${item.Id} | Solicitud=${solicitudId}`);
      continue;
    }

    solicitudesEvaluadas.add(solicitudId);
    if (!debeProcesarSolicitud(solicitud)) {
      totalOmitidos++;
      reportRows.push({
        RegistroAprobadorId: Number(item.Id || 0) || '',
        SolicitudId: solicitudId,
        SolicitudTitulo: String(solicitud.Title || ''),
        SolicitudEstado: String(solicitud.Estado?.Title || ''),
        Rol: String(item.Rol || ''),
        AprobadorId: aprobadorId,
        AprobadorNombre: String(item.Aprobador?.Title || ''),
        AprobadorEmail: String(item.Aprobador?.EMail || ''),
        AccionSolicitud: String(solicitud.Accion || ''),
        MotivoSolicitud: String(solicitud.Motivo?.Title || ''),
        TipoDocumento: String(solicitud.TipoDocumento?.Title || ''),
        AreasImpactadas: formatAreasImpactadas(solicitud),
        ImpactadoPorAreaActual: formatBoolean(currentArea),
        ImpactadoPorAreaCalculado: '',
        ImpactadoPorMotivoActual: formatBoolean(currentMotivo),
        ImpactadoPorMotivoCalculado: '',
        ImpactadoPorAccionActual: formatBoolean(currentAccion),
        ImpactadoPorAccionCalculado: '',
        CambiariaImpactadoPorArea: 'No',
        CambiariaImpactadoPorMotivo: 'No',
        CambiariaImpactadoPorAccion: 'No',
        ResultadoEnsayo: 'OMITIDO',
        Motivo: 'La solicitud tiene Estado = Vigente; el utilitario solo procesa Estado distinto de Vigente.'
      });
      if (!solicitudesVigentesOmitidas.has(solicitudId)) {
        solicitudesVigentesOmitidas.add(solicitudId);
        totalSolicitudesVigentesOmitidas++;
        writeInfo(log, `⏭️ ModificarAprobadores | Solicitud vigente omitida | Solicitud=${solicitudId} | Estado=${solicitud.Estado?.Title || ''}`);
      }
      continue;
    }

    const impactos = await getImpactosFromCache(params.context, webUrl, solicitud, impactosCache);
    const calculadoPorArea = impactos.porArea.has(aprobadorId);
    const correcto = {
      ImpactadoPorArea: currentArea || calculadoPorArea,
      ImpactadoPorMotivo: impactos.porMotivo.has(aprobadorId),
      ImpactadoPorAccion: impactos.porAccion.has(aprobadorId)
    };

    const cambiaArea = currentArea !== correcto.ImpactadoPorArea;
    const cambiaMotivo = currentMotivo !== correcto.ImpactadoPorMotivo;
    const cambiaAccion = currentAccion !== correcto.ImpactadoPorAccion;
    const tieneCambios = cambiaArea || cambiaMotivo || cambiaAccion;

    const reportRow: IModificarAprobadoresReportRow = {
      RegistroAprobadorId: Number(item.Id || 0) || '',
      SolicitudId: solicitudId,
      SolicitudTitulo: String(solicitud.Title || ''),
      SolicitudEstado: String(solicitud.Estado?.Title || ''),
      Rol: String(item.Rol || ''),
      AprobadorId: aprobadorId,
      AprobadorNombre: String(item.Aprobador?.Title || ''),
      AprobadorEmail: String(item.Aprobador?.EMail || ''),
      AccionSolicitud: String(solicitud.Accion || ''),
      MotivoSolicitud: String(solicitud.Motivo?.Title || ''),
      TipoDocumento: String(solicitud.TipoDocumento?.Title || ''),
      AreasImpactadas: formatAreasImpactadas(solicitud),
      ImpactadoPorAreaActual: formatBoolean(currentArea),
      ImpactadoPorAreaCalculado: formatBoolean(correcto.ImpactadoPorArea),
      ImpactadoPorMotivoActual: formatBoolean(currentMotivo),
      ImpactadoPorMotivoCalculado: formatBoolean(correcto.ImpactadoPorMotivo),
      ImpactadoPorAccionActual: formatBoolean(currentAccion),
      ImpactadoPorAccionCalculado: formatBoolean(correcto.ImpactadoPorAccion),
      CambiariaImpactadoPorArea: cambiaArea ? 'Si' : 'No',
      CambiariaImpactadoPorMotivo: cambiaMotivo ? 'Si' : 'No',
      CambiariaImpactadoPorAccion: cambiaAccion ? 'Si' : 'No',
      ResultadoEnsayo: tieneCambios ? 'CAMBIARIA' : 'SIN_CAMBIOS',
      Motivo: tieneCambios
        ? 'Los checks actuales no coinciden con los valores calculados por Área, Motivo o Acción. Si ImpactadoPorArea ya estaba marcado, se conserva aunque no se ubique por área.'
        : 'Los checks actuales ya coinciden con los valores calculados. Si ImpactadoPorArea ya estaba marcado, se conserva aunque no se ubique por área.'
    };

    if (!tieneCambios) {
      reportRows.push(reportRow);
      totalOmitidos++;
      writeInfo(log, `⏭️ ModificarAprobadores | Sin cambios | Registro=${item.Id} | Aprobador=${aprobadorId}`);
      continue;
    }

    totalCambiarian++;
    try {
      await actualizarChecksAprobador(params.context, webUrl, Number(item.Id || 0), correcto);
      totalActualizados++;
      reportRow.ResultadoEnsayo = 'ACTUALIZADO';
      reportRow.Motivo = 'Checks actualizados con los valores calculados por Área, Motivo y Acción. Si ImpactadoPorArea ya estaba marcado, se conserva aunque no se ubique por área.';
      writeInfo(
        log,
        `✅ ModificarAprobadores | Actualizado | Registro=${item.Id} | Aprobador=${aprobadorId} | ` +
        `Area=${correcto.ImpactadoPorArea} | Motivo=${correcto.ImpactadoPorMotivo} | Accion=${correcto.ImpactadoPorAccion}`
      );
    } catch (error) {
      totalError++;
      reportRow.ResultadoEnsayo = 'ERROR';
      reportRow.Motivo = `No se pudo actualizar el registro: ${error instanceof Error ? error.message : String(error)}`;
      writeError(log, `❌ ModificarAprobadores | Error actualizando | Registro=${item.Id} | ${reportRow.Motivo}`);
    }

    reportRows.push(reportRow);
  }

  const workbook = buildModificarAprobadoresWorkbook(reportRows);

  writeInfo(
    log,
    `📌 ModificarAprobadores | Resumen final | Encontrados=${aprobadores.length} | ` +
    `SolicitudesEvaluadas=${solicitudesEvaluadas.size} | SolicitudesVigentesOmitidas=${totalSolicitudesVigentesOmitidas} | ` +
    `Cambiarian=${totalCambiarian} | Actualizados=${totalActualizados} | Omitidos=${totalOmitidos} | Error=${totalError}`
  );
  writeInfo(log, `📥 ModificarAprobadores | Excel de resultado generado: ${workbook.fileName}`);

  return {
    blob: workbook.blob,
    fileName: workbook.fileName,
    totalEncontrados: aprobadores.length,
    totalCambiarian,
    totalActualizados,
    totalOmitidos,
    totalError,
    totalSolicitudesEvaluadas: solicitudesEvaluadas.size,
    totalSolicitudesVigentesOmitidas
  };
}

export async function rollbackModificarAprobadores(params: {
  context: WebPartContext;
  excelFile: IFilePickerResult;
  log?: LogFn;
}): Promise<IModificarAprobadoresRollbackResultado> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const session = await openExcelRevisionSession(params.excelFile);
  const grid = session.grid || [];

  if (!grid.length) {
    throw new Error('El Excel de rollback está vacío.');
  }

  const headerMap = buildHeaderMap(grid[0] || []);
  const requiredHeaders = [
    'RegistroAprobadorId',
    'ImpactadoPorAreaActual',
    'ImpactadoPorMotivoActual',
    'ImpactadoPorAccionActual'
  ];

  for (let i = 0; i < requiredHeaders.length; i++) {
    if ((headerMap.get(requiredHeaders[i]) ?? -1) < 0) {
      throw new Error(`El Excel de rollback debe contener la columna "${requiredHeaders[i]}".`);
    }
  }

  let totalRestaurados = 0;
  let totalOmitidos = 0;
  let totalError = 0;
  const reportRows: IModificarAprobadoresReportRow[] = [];

  writeInfo(log, `↩️ Rollback ModificarAprobadores | Archivo cargado: ${session.fileName}`);

  for (let rowIndex = 1; rowIndex < grid.length; rowIndex++) {
    const parsed = parseReportRow(grid[rowIndex] || [], headerMap);
    const registroAprobadorId = Number(parsed.RegistroAprobadorId || 0);

    if (!registroAprobadorId) {
      totalOmitidos++;
      parsed.ResultadoEnsayo = 'OMITIDO';
      parsed.Motivo = 'Fila sin RegistroAprobadorId.';
      reportRows.push(parsed);
      continue;
    }

    const rollbackValues = {
      ImpactadoPorArea: isTruthyField(parsed.ImpactadoPorAreaActual),
      ImpactadoPorMotivo: isTruthyField(parsed.ImpactadoPorMotivoActual),
      ImpactadoPorAccion: isTruthyField(parsed.ImpactadoPorAccionActual)
    };

    try {
      await actualizarChecksAprobador(params.context, webUrl, registroAprobadorId, rollbackValues);
      totalRestaurados++;
      parsed.ResultadoEnsayo = 'RESTAURADO';
      parsed.Motivo = 'Checks restaurados usando los valores Actual del Excel.';
      writeInfo(log, `✅ Rollback ModificarAprobadores | Restaurado Registro=${registroAprobadorId}`);
    } catch (error) {
      totalError++;
      parsed.ResultadoEnsayo = 'ERROR';
      parsed.Motivo = `No se pudo restaurar el registro: ${error instanceof Error ? error.message : String(error)}`;
      writeError(log, `❌ Rollback ModificarAprobadores | Error Registro=${registroAprobadorId} | ${parsed.Motivo}`);
    }

    reportRows.push(parsed);
  }

  const workbook = buildModificarAprobadoresWorkbook(reportRows, 'Rollback_ModificarAprobadores');
  writeInfo(
    log,
    `📌 Rollback ModificarAprobadores | Resumen final | Filas=${Math.max(grid.length - 1, 0)} | ` +
    `Restaurados=${totalRestaurados} | Omitidos=${totalOmitidos} | Error=${totalError}`
  );
  writeInfo(log, `📥 Rollback ModificarAprobadores | Excel de resultado generado: ${workbook.fileName}`);

  return {
    blob: workbook.blob,
    fileName: workbook.fileName,
    totalFilas: Math.max(grid.length - 1, 0),
    totalRestaurados,
    totalOmitidos,
    totalError
  };
}
