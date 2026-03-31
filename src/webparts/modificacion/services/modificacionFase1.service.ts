/* eslint-disable */
// @ts-nocheck
import { IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { fillAndAttachFromFolder } from './documentFillAndAttach.service';
import { listFilesRecursive } from './spFolderExplorer.service';
import { addListItem, getAllItems, spGetJson, updateListItem } from './sharepointRest.service';
import { descargarReporteFase1Word, IFase1WordReportRow } from '../utils/fase1WordReportExcel';
import { openExcelRevisionSession } from '../utils/modificacionExcelHelper';

type LogFn = (s: string) => void;

type IExcelRowData = {
  clasificacion: string;
  macroproceso: string;
  proceso: string;
  subproceso: string;
  areaDuena: string;
  duenoDocumento: string;
  categoriaDocumento: string;
  tipoDocumento: string;
  nombreArchivo: string;
  nombreDocumento: string;
  documentoPadre: string;
  versionDocumento: string;
  fechaAprobacion: string;
  fechaVigencia: string;
  instanciaAprobacion: string;
  flagConducta: string;
  flagExperiencia: string;
  areasImpactadasTxt: string;
  resumen: string;
};

function cleanPart(value: string): string {
  const text = String(value || '').trim();
  if (!text) return '';
  const lower = text.toLowerCase();
  if (text === '-' || text === '—' || lower === 'na' || lower === 'n/a' || lower === 'null') {
    return '';
  }
  if (lower === 'sin subproceso') return '';
  return text;
}

function sanitizeFolderPart(value: string): string {
  return String(value || '')
    .trim()
    .replace(/[~#%&*{}\\:<>?/+"|]/g, '')
    .replace(/\s+/g, ' ');
}

function buildDestinoWordTemp(baseFolder: string, a: string, b: string, c: string, d: string): string {
  const parts = [a, b, c, d]
    .map(cleanPart)
    .filter(Boolean)
    .map(sanitizeFolderPart);

  const root = (baseFolder || '').replace(/\/$/, '');
  return parts.length ? `${root}/${parts.join('/')}` : root;
}

function isEmptyLike(value: any): boolean {
  const text = String(value ?? '').trim().toLowerCase();
  return !text || text === '-' || text === '—' || text === 'na' || text === 'n/a' || text === 'null';
}

function isSi(value: any): boolean {
  const text = String(value ?? '').trim().toLowerCase();
  return text === 'si' || text === 'sí' || text === 's';
}

function incrementVersion(value: any): string {
  const text = String(value ?? '').trim();
  if (!text) return '1.1';

  const match = text.match(/^(\d+)(?:\.(\d+))?$/);
  if (!match) return text;

  const major = parseInt(match[1], 10);
  const minorText = match[2] || '0';
  const minor = parseInt(minorText, 10) + 1;
  return `${major}.${minor}`;
}

function normKey(value: any): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function addImpactNameUnique(set: Set<string>, name: string): void {
  const key = normKey(name);
  if (key) {
    set.add(key);
  }
}

function findColumnIndex(headers: any[], expected: string): number {
  const normalized = String(expected || '').trim().toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim().toLowerCase() === normalized) {
      return i;
    }
  }
  return -1;
}

async function buildLookupMapByField(
  context: WebPartContext,
  webUrl: string,
  listTitle: string,
  fieldInternalName: string
): Promise<Map<string, number>> {
  const field = await spGetJson<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('${listTitle}')/fields/getbyinternalnameortitle('${fieldInternalName}')?$select=LookupList`
  );

  const lookupListId = field.LookupList;
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists(guid'${lookupListId}')/items?$select=Id,Title&$top=5000`
  );

  const map = new Map<string, number>();
  for (let i = 0; i < items.length; i++) {
    const key = normKey(items[i].Title);
    if (key) {
      map.set(key, items[i].Id);
    }
  }

  return map;
}

async function buildProcesoCorporativoMap(
  context: WebPartContext,
  webUrl: string
): Promise<Map<string, number>> {
  const items = await getAllItems<any>(
    context,
    `${webUrl}/_api/web/lists/getbytitle('Procesos Corporativos')/items?$select=Id,Title,field_1,field_2,field_3&$top=5000`
  );

  const map = new Map<string, number>();
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const key = normKey(
      [cleanPart(item.Title), cleanPart(item.field_1), cleanPart(item.field_2), cleanPart(item.field_3)]
        .filter(Boolean)
        .join('/')
    );
    if (key) {
      map.set(key, item.Id);
    }
  }

  return map;
}

function parseExcelRow(row: any[]): IExcelRowData {
  return {
    clasificacion: String(row[0] || '').trim(),
    macroproceso: String(row[1] || '').trim(),
    proceso: String(row[2] || '').trim(),
    subproceso: String(row[3] || '').trim(),
    areaDuena: String(row[4] || '').trim(),
    duenoDocumento: String(row[5] || '').trim(),
    categoriaDocumento: String(row[6] || '').trim(),
    tipoDocumento: String(row[7] || '').trim(),
    nombreArchivo: String(row[8] || '').trim(),
    nombreDocumento: String(row[9] || '').trim(),
    documentoPadre: String(row[10] || '').trim(),
    versionDocumento: String(row[11] || '').trim(),
    fechaAprobacion: String(row[12] || '').trim(),
    fechaVigencia: String(row[13] || '').trim(),
    instanciaAprobacion: String(row[14] || '').trim(),
    flagConducta: String(row[15] || '').trim(),
    flagExperiencia: String(row[16] || '').trim(),
    areasImpactadasTxt: String(row[17] || '').trim(),
    resumen: String(row[18] || '').trim()
  };
}

function buildImpactNames(excelRow: IExcelRowData): string[] {
  const impactNamesRaw = (excelRow.areasImpactadasTxt || '')
    .split(/[\n;,|/]+/g)
    .map((item) => item.trim())
    .filter(Boolean);

  const impactSet = new Set<string>();
  for (let i = 0; i < impactNamesRaw.length; i++) {
    addImpactNameUnique(impactSet, impactNamesRaw[i]);
  }

  if (isSi(excelRow.flagConducta)) {
    addImpactNameUnique(impactSet, 'Conducta de Mercado');
  }

  if (isSi(excelRow.flagExperiencia)) {
    addImpactNameUnique(impactSet, 'Experiencia del Cliente');
  }

  return Array.from(impactSet);
}

async function getSolicitudById(context: WebPartContext, webUrl: string, solicitudId: number): Promise<any> {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('Solicitudes')/items(${solicitudId})` +
    `?$select=Id,Title,NombreDocumento,CodigoDocumento,CategoriadeDocumento,ResumenDocumento,` +
    `FechaDeAprobacionSolicitud,FechadeVigencia,VersionDocumento,TipoDocumentoId,TipoDocumento/Title,` +
    `ProcesoDeNegocioId,ProcesoDeNegocio/Title,ProcesoDeNegocio/field_1,ProcesoDeNegocio/field_2,ProcesoDeNegocio/field_3,` +
    `AreaDuenaId,AreaDuena/Title,EstadoId,InstanciasdeaprobacionId,Instanciasdeaprobacion/Title,` +
    `AreasImpactadas/Id,AreasImpactadas/Title,Accion,DocumentosApoyo,EsDocumentoVigente,EsVersionActualDocumento` +
    `&$expand=TipoDocumento,ProcesoDeNegocio,AreaDuena,Instanciasdeaprobacion,AreasImpactadas`;

  return spGetJson<any>(context, url);
}

function buildNewSolicitudPayload(
  oldSolicitud: any,
  excelRow: IExcelRowData,
  versionDocumento: string,
  lookups: {
    tipoDocumentoId?: number;
    procesoDeNegocioId?: number;
    areaDuenaId?: number;
    instanciaAprobacionId?: number;
    impactAreaIds: number[];
  }
): any {
  const payload: any = {
    Title: excelRow.nombreDocumento || oldSolicitud.Title || oldSolicitud.NombreDocumento || '',
    Accion: 'Actualización de documentos',
    NombreDocumento: excelRow.nombreDocumento || oldSolicitud.NombreDocumento || oldSolicitud.Title || '',
    CategoriadeDocumento: excelRow.categoriaDocumento || oldSolicitud.CategoriadeDocumento || '',
    ResumenDocumento: excelRow.resumen || oldSolicitud.ResumenDocumento || '',
    FechaDeAprobacionSolicitud: excelRow.fechaAprobacion || oldSolicitud.FechaDeAprobacionSolicitud || null,
    FechadeVigencia: excelRow.fechaVigencia || oldSolicitud.FechadeVigencia || null,
    FechaDePublicacionSolicitud: new Date().toISOString(),
    FechadeEnvio: new Date().toISOString(),
    VersionDocumento: versionDocumento,
    EsVersionActualDocumento: true,
    EsDocumentoVigente: oldSolicitud.EsDocumentoVigente,
    DocumentosApoyo: false,
    CodigoDocumento: oldSolicitud.CodigoDocumento || ''
  };

  if (lookups.tipoDocumentoId) payload.TipoDocumentoId = lookups.tipoDocumentoId;
  if (lookups.procesoDeNegocioId) payload.ProcesoDeNegocioId = lookups.procesoDeNegocioId;
  if (lookups.areaDuenaId) payload.AreaDuenaId = lookups.areaDuenaId;
  if (oldSolicitud.EstadoId) payload.EstadoId = oldSolicitud.EstadoId;
  if (lookups.instanciaAprobacionId) payload.InstanciasdeaprobacionId = lookups.instanciaAprobacionId;
  if (lookups.impactAreaIds.length) payload.AreasImpactadasId = { results: lookups.impactAreaIds };

  return payload;
}

export async function ejecutarFase1DocumentosSinHijosNiFlujos(params: {
  context: WebPartContext;
  excelFile: IFilePickerResult;
  sourceFolderServerRelativeUrl: string;
  tempWordBaseFolderServerRelativeUrl: string;
  log?: LogFn;
}): Promise<{
  createdSolicitudIds: number[];
  oldSolicitudIds: number[];
  tempFileUrls: string[];
  reportRows: IFase1WordReportRow[];
  processed: number;
  ok: number;
  skipped: number;
  error: number;
}> {
  const log = params.log || (() => undefined);
  const webUrl = params.context.pageContext.web.absoluteUrl;
  const session = await openExcelRevisionSession(params.excelFile);
  const grid = session.grid || [];
  if (!grid.length) {
    throw new Error('El Excel está vacío.');
  }

  const headers = grid[0] || [];
  const idxSolicitud = findColumnIndex(headers, 'ID Solicitud');
  const idxHijos = findColumnIndex(headers, 'Documentos hijos');
  const idxFlujos = findColumnIndex(headers, 'Diagrama de Flujos');

  if (idxSolicitud < 0 || idxHijos < 0 || idxFlujos < 0) {
    throw new Error('El Excel no contiene las columnas ID Solicitud, Documentos hijos y Diagrama de Flujos.');
  }

  const files = await listFilesRecursive(params.context, webUrl, params.sourceFolderServerRelativeUrl, log);
  const mapTipoDoc = await buildLookupMapByField(params.context, webUrl, 'Solicitudes', 'TipoDocumento');
  const mapArea = await buildLookupMapByField(params.context, webUrl, 'Solicitudes', 'AreaDuena');
  const mapInst = await buildLookupMapByField(params.context, webUrl, 'Solicitudes', 'Instanciasdeaprobacion');
  const mapImpact = await buildLookupMapByField(params.context, webUrl, 'Solicitudes', 'AreasImpactadas');
  const mapProceso = await buildProcesoCorporativoMap(params.context, webUrl);
  const fileByName = new Map<string, string>();
  for (let i = 0; i < files.length; i++) {
    fileByName.set(String(files[i].Name || '').toLowerCase(), files[i].ServerRelativeUrl);
  }

  const createdSolicitudIds: number[] = [];
  const oldSolicitudIds: number[] = [];
  const tempFileUrls: string[] = [];
  const reportRows: IFase1WordReportRow[] = [];

  for (let rowIndex = 1; rowIndex < grid.length; rowIndex++) {
    const row = grid[rowIndex] || [];
    const excelRow = parseExcelRow(row);
    const solicitudOrigenId = Number(row[idxSolicitud] || 0);
    const nombreArchivo = excelRow.nombreArchivo;
    const nombreDocumento = excelRow.nombreDocumento;
    const documentoPadre = excelRow.documentoPadre;
    const versionExcel = excelRow.versionDocumento;
    const duenoDocumento = excelRow.duenoDocumento;

    const tieneHijos = !isEmptyLike(row[idxHijos]);
    const tieneFlujos = !isEmptyLike(row[idxFlujos]);
    const tienePadre = !isEmptyLike(documentoPadre);

    if (!solicitudOrigenId || !nombreDocumento) {
      continue;
    }

    if (tieneHijos || tieneFlujos || tienePadre) {
      reportRows.push({
        SolicitudOrigenID: solicitudOrigenId,
        SolicitudID: '',
        NombreDocumento: nombreDocumento,
        NombreArchivo: nombreArchivo,
        CodigoDocumento: '',
        VersionDocumento: versionExcel,
        TieneDocumentoPadre: tienePadre ? 'Sí' : 'No',
        DocumentoPadreNombre: tienePadre ? documentoPadre : '',
        DocumentoPadreSolicitudID: '',
        PadreRegeneradoConLinks: 'No',
        RutaTemporalWord: '',
        EstadoFase1: 'SKIP_NO_APLICA',
        Error: 'Documento omitido por tener hijos, flujos o documento padre.'
      });
      continue;
    }

    try {
      const oldSolicitud = await getSolicitudById(params.context, webUrl, solicitudOrigenId);
      const versionNueva = incrementVersion(versionExcel || oldSolicitud.VersionDocumento || '1.0');
      const procesoDeNegocioKey = normKey(
        [cleanPart(excelRow.clasificacion), cleanPart(excelRow.macroproceso), cleanPart(excelRow.proceso), cleanPart(excelRow.subproceso)]
          .filter(Boolean)
          .join('/')
      );
      const impactNames = buildImpactNames(excelRow);
      const impactAreaIds = impactNames
        .map((name) => mapImpact.get(name))
        .filter((value): value is number => !!value);
      const tipoDocumentoId = mapTipoDoc.get(normKey(excelRow.tipoDocumento)) || oldSolicitud.TipoDocumentoId;
      const procesoDeNegocioId = mapProceso.get(procesoDeNegocioKey) || oldSolicitud.ProcesoDeNegocioId;
      const areaDuenaId = mapArea.get(normKey(excelRow.areaDuena)) || oldSolicitud.AreaDuenaId;
      const instanciaAprobacionId = mapInst.get(normKey(excelRow.instanciaAprobacion)) || oldSolicitud.InstanciasdeaprobacionId;

      await updateListItem(params.context, webUrl, 'Solicitudes', solicitudOrigenId, {
        EsVersionActualDocumento: false
      });

      if (oldSolicitudIds.indexOf(solicitudOrigenId) === -1) {
        oldSolicitudIds.push(solicitudOrigenId);
      }

      const newSolicitudId = await addListItem(
        params.context,
        webUrl,
        'Solicitudes',
        buildNewSolicitudPayload(oldSolicitud, excelRow, versionNueva, {
          tipoDocumentoId,
          procesoDeNegocioId,
          areaDuenaId,
          instanciaAprobacionId,
          impactAreaIds
        })
      );

      createdSolicitudIds.push(newSolicitudId);

      const proceso = oldSolicitud.ProcesoDeNegocio || {};
      const tempDestino = buildDestinoWordTemp(
        params.tempWordBaseFolderServerRelativeUrl,
        excelRow.clasificacion || proceso.Title || '',
        excelRow.macroproceso || proceso.field_1 || '',
        excelRow.proceso || proceso.field_2 || '',
        excelRow.subproceso || proceso.field_3 || ''
      );

      const attachResult = await fillAndAttachFromFolder({
        context: params.context,
        webUrl,
        listTitle: 'Solicitudes',
        itemId: newSolicitudId,
        originalFileName: nombreArchivo,
        fileByName,
        titulo: excelRow.nombreDocumento || oldSolicitud.NombreDocumento || oldSolicitud.Title || nombreDocumento,
        instanciaRaw: excelRow.instanciaAprobacion || oldSolicitud?.Instanciasdeaprobacion?.Title || '',
        impactAreaIds,
        dueno: duenoDocumento,
        fechaVigencia: excelRow.fechaVigencia || oldSolicitud.FechadeVigencia || '',
        fechaAprobacion: excelRow.fechaAprobacion || oldSolicitud.FechaDeAprobacionSolicitud || '',
        resumen: excelRow.resumen || oldSolicitud.ResumenDocumento || '',
        version: versionNueva,
        codigoDocumento: oldSolicitud.CodigoDocumento || '',
        categoriaDoc: excelRow.categoriaDocumento || oldSolicitud.CategoriadeDocumento || '',
        tipoDocExcel: excelRow.tipoDocumento || oldSolicitud?.TipoDocumento?.Title || '',
        esDocumentoApoyo: false,
        tempDestinoFolderServerRelativeUrl: tempDestino,
        replaceIfExists: true,
        log
      });

      if (attachResult.tempFileServerRelativeUrl) {
        tempFileUrls.push(attachResult.tempFileServerRelativeUrl);
      }

      reportRows.push({
        SolicitudOrigenID: solicitudOrigenId,
        SolicitudID: newSolicitudId,
        NombreDocumento: oldSolicitud.NombreDocumento || oldSolicitud.Title || nombreDocumento,
        NombreArchivo: nombreArchivo,
        CodigoDocumento: oldSolicitud.CodigoDocumento || '',
        VersionDocumento: versionNueva,
        TieneDocumentoPadre: 'No',
        DocumentoPadreNombre: '',
        DocumentoPadreSolicitudID: '',
        PadreRegeneradoConLinks: 'No',
        RutaTemporalWord: attachResult.tempFileServerRelativeUrl || '',
        EstadoFase1: attachResult.ok ? 'OK' : 'ERROR',
        Error: attachResult.error || '',
        TipoDocumento: excelRow.tipoDocumento || oldSolicitud?.TipoDocumento?.Title || '',
        CategoriaDocumento: excelRow.categoriaDocumento || oldSolicitud.CategoriadeDocumento || '',
        Clasificaciondeproceso: excelRow.clasificacion || proceso.Title || '',
        Macroproceso: excelRow.macroproceso || proceso.field_1 || '',
        Proceso: excelRow.proceso || proceso.field_2 || '',
        Subproceso: excelRow.subproceso || proceso.field_3 || '',
        AreaDuena: excelRow.areaDuena || oldSolicitud?.AreaDuena?.Title || '',
        AreaImpactada: impactNames.join(' / '),
        Resumen: excelRow.resumen || oldSolicitud.ResumenDocumento || '',
        FechaDeAprobacion: excelRow.fechaAprobacion || oldSolicitud.FechaDeAprobacionSolicitud || '',
        FechaDeVigencia: excelRow.fechaVigencia || oldSolicitud.FechadeVigencia || '',
        InstanciaDeAprobacionId: instanciaAprobacionId || oldSolicitud.InstanciasdeaprobacionId || '',
        MetadataPendiente: 'Sí'
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      reportRows.push({
        SolicitudOrigenID: solicitudOrigenId,
        SolicitudID: '',
        NombreDocumento: nombreDocumento,
        NombreArchivo: nombreArchivo,
        CodigoDocumento: '',
        VersionDocumento: versionExcel,
        TieneDocumentoPadre: 'No',
        DocumentoPadreNombre: '',
        DocumentoPadreSolicitudID: '',
        PadreRegeneradoConLinks: 'No',
        RutaTemporalWord: '',
        EstadoFase1: 'ERROR',
        Error: message
      });
      log(`❌ Error Fase 1 | SolicitudOrigen=${solicitudOrigenId} | ${message}`);
    }
  }

  descargarReporteFase1Word(reportRows);

  return {
    createdSolicitudIds,
    oldSolicitudIds,
    tempFileUrls,
    reportRows,
    processed: reportRows.length,
    ok: reportRows.filter((row) => row.EstadoFase1 === 'OK').length,
    skipped: reportRows.filter((row) => row.EstadoFase1.indexOf('SKIP') === 0).length,
    error: reportRows.filter((row) => row.EstadoFase1 === 'ERROR').length
  };
}
