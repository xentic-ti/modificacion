/* eslint-disable */
import * as React from 'react';
import {
  DefaultButton,
  Icon,
  MessageBar,
  MessageBarType,
  Stack,
  Text,
  TextField
} from '@fluentui/react';
import {
  FilePicker,
  IFilePickerResult
} from '@pnp/spfx-controls-react/lib/FilePicker';
import { FolderPicker } from '@pnp/spfx-controls-react/lib/FolderPicker';

import styles from './Modificacion.module.scss';
import type { IModificacionProps } from './IModificacionProps';
import { revisarExcelModificacion } from '../services/modificacionRevision.service';
import { ejecutarFase1DocumentosSinHijosNiFlujos } from '../services/modificacionFase1.service';
import { rollbackModificacionFase1 } from '../services/modificacionFase1Rollback.service';
import { IFase2RollbackEntry, ejecutarFase2PublicacionDocumentosSinHijos } from '../services/modificacionFase2Publicacion.service';
import { rollbackModificacionFase2 } from '../services/modificacionFase2Rollback.service';
import { ejecutarFase3PadresConHijosYFlujos, IFase3RollbackEntry } from '../services/modificacionFase3Padres.service';
import { rollbackModificacionFase3 } from '../services/modificacionFase3Rollback.service';
import { ejecutarFase4PublicacionHijos } from '../services/modificacionFase4Hijos.service';
import { ejecutarFase5HijosConPadre } from '../services/modificacionFase5HijosConPadre.service';
import { ejecutarFase6BajaDocumentos } from '../services/modificacionFase6Baja.service';
import { IFase6RollbackDiagramEntry, IFase6RollbackEntry, rollbackModificacionFase6 } from '../services/modificacionFase6Rollback.service';
import { ejecutarFase7ModificacionDiagramas } from '../services/modificacionFase7Diagrama.service';
import { IFase7RollbackEntry, rollbackModificacionFase7 } from '../services/modificacionFase7Rollback.service';
import { ejecutarFase8AltaDiagramaSolicitud } from '../services/modificacionFase8AltaDiagrama.service';
import { ejecutarFase9BajaDiagramaSolicitud } from '../services/modificacionFase9BajaDiagrama.service';
import { ejecutarCorreccionSolicitud2460 } from '../services/modificacionCorreccion2460.service';
import { IFase8RollbackEntry, rollbackModificacionFase8 } from '../services/modificacionFase8Rollback.service';
import { corregirDocPadresDesdeRelaciones } from '../services/docPadresFix.service';
import { copiarHijosRelacionesDocumentos } from '../services/copiarHijos.service';
import { exportarSolicitudesNoCreadasPorAntonio } from '../services/solicitudesNoAntonio.service';
import { buscarSolicitudesDuplicadas } from '../services/solicitudesDuplicadas.service';
import { ejecutarCambioInstanciaDesdeExcel } from '../services/modificacionCambioInstancia.service';
import { ejecutarCorreccionCodigosDuplicadosDesdeExcel } from '../services/modificacionCodigoDuplicado.service';
import { ejecutarBuscarDiagramaFlujoSinCodigo } from '../services/buscarDiagramaFlujoSinCodigo.service';
import { modificarAprobadores, rollbackModificarAprobadores } from '../services/modificarAprobadores.service';

const Modificacion: React.FC<IModificacionProps> = ({ context, hasTeamsContext, isDarkTheme }) => {
  const [excelFile, setExcelFile] = React.useState<IFilePickerResult | null>(null);
  const [copiarHijosPadreOrigenId, setCopiarHijosPadreOrigenId] = React.useState<string>('');
  const [copiarHijosPadreDestinoId, setCopiarHijosPadreDestinoId] = React.useState<string>('');
  const [sourceFolderUrl, setSourceFolderUrl] = React.useState<string>('');
  const [error, setError] = React.useState<string | null>(null);
  const [isRunning, setIsRunning] = React.useState<boolean>(false);
  const [createdSolicitudIds, setCreatedSolicitudIds] = React.useState<number[]>([]);
  const [oldSolicitudIds, setOldSolicitudIds] = React.useState<number[]>([]);
  const [tempFileUrls, setTempFileUrls] = React.useState<string[]>([]);
  const [fase2RollbackEntries, setFase2RollbackEntries] = React.useState<IFase2RollbackEntry[]>([]);
  const [fase3RollbackEntries, setFase3RollbackEntries] = React.useState<IFase3RollbackEntry[]>([]);
  const [fase4RollbackEntries, setFase4RollbackEntries] = React.useState<IFase2RollbackEntry[]>([]);
  const [fase6RollbackEntries, setFase6RollbackEntries] = React.useState<IFase6RollbackEntry[]>([]);
  const [fase6DiagramRollbackEntries, setFase6DiagramRollbackEntries] = React.useState<IFase6RollbackDiagramEntry[]>([]);
  const [fase7RollbackEntries, setFase7RollbackEntries] = React.useState<IFase7RollbackEntry[]>([]);
  const [fase8RollbackEntries, setFase8RollbackEntries] = React.useState<IFase8RollbackEntry[]>([]);
  const [fase9RollbackEntries, setFase9RollbackEntries] = React.useState<IFase8RollbackEntry[]>([]);
  const [lastExecutedPhase, setLastExecutedPhase] = React.useState<'fase1' | 'fase2' | 'fase3' | 'fase4' | 'fase6' | 'fase7' | 'fase8' | 'fase9' | 'correccion2460' | null>(null);
  const [logRevision, setLogRevision] = React.useState<string>(
    'Panel de revisión listo.\nEsperando la selección de un archivo Excel de modificación.'
  );

  const appendLog = React.useCallback((linea: string) => {
    setLogRevision((prev) => `${prev}\n${linea}`);
  }, []);

  const onSaveExcel = React.useCallback((files: IFilePickerResult[]): void => {
    const selectedFile = files && files.length ? files[0] : null;
    if (!selectedFile) {
      return;
    }

    const extensionValida = /\.(xlsx|xlsm)$/i.test(selectedFile.fileName || '');
    if (!extensionValida) {
      setExcelFile(null);
      setError('Selecciona un archivo Excel valido (.xlsx o .xlsm).');
      setLogRevision('Panel de revisión listo.\nEsperando la selección de un archivo Excel de modificación.');
      return;
    }

    setExcelFile(selectedFile);
    setFase2RollbackEntries([]);
    setFase3RollbackEntries([]);
    setFase4RollbackEntries([]);
    setFase6RollbackEntries([]);
    setFase6DiagramRollbackEntries([]);
    setFase7RollbackEntries([]);
    setFase8RollbackEntries([]);
    setFase9RollbackEntries([]);
    setLastExecutedPhase(null);
    setError(null);
    setLogRevision([
      'Panel de revisión listo.',
      `Archivo cargado desde SharePoint: ${selectedFile.fileName}`,
      'La revisión ya puede ejecutarse con ese archivo.'
    ].join('\n'));
  }, []);

  const descargarArchivo = React.useCallback((blob: Blob, fileName: string): void => {
    const blobUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = blobUrl;
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(blobUrl);
  }, []);

  const ejecutarRevision = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar un Excel antes de ejecutar la revisión.');
      appendLog('No se pudo iniciar la revisión porque no hay archivo seleccionado.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision(`Iniciando revisión del archivo: ${excelFile.fileName}`);

    try {
      const resultado = await revisarExcelModificacion({
        context,
        file: excelFile,
        log: appendLog
      });

      descargarArchivo(resultado.blob, resultado.fileName);
      appendLog(`✅ Revisión terminada. Filas procesadas: ${resultado.processed}`);
      appendLog(`✅ Solicitudes encontradas: ${resultado.found}`);
      appendLog(`ℹ️ Solicitudes no encontradas: ${resultado.notFound}`);
      appendLog(`📥 Archivo generado: ${resultado.fileName}`);
    } catch (revisionError) {
      const errorMessage = revisionError instanceof Error ? revisionError.message : String(revisionError);
      setError(errorMessage);
      appendLog(`❌ Error durante la revisión: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, descargarArchivo, excelFile]);

  const revisarArchivo = React.useCallback((): void => {
    void ejecutarRevision();
  }, [ejecutarRevision]);





  const ejecutarCorreccionDocPadres = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar el Excel de la corrida anterior antes de corregir DocPadres.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision('Iniciando correccion de DocPadres desde Relaciones Documentos...');

    try {
      const resultado = await corregirDocPadresDesdeRelaciones({
        context,
        excelFile,
        log: appendLog
      });

      descargarArchivo(resultado.blob, resultado.fileName);
      appendLog('✅ Correccion de DocPadres terminada. Hijos procesados: ' + resultado.processed);
      appendLog('✅ Actualizados: ' + resultado.updated + ' | SKIP: ' + resultado.skipped + ' | ERROR: ' + resultado.error);
      appendLog('📥 Archivo generado: ' + resultado.fileName);
    } catch (fixError) {
      const errorMessage = fixError instanceof Error ? fixError.message : String(fixError);
      setError(errorMessage);
      appendLog('❌ Error en correccion de DocPadres: ' + errorMessage);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, descargarArchivo, excelFile]);

  const ejecutarCopiarHijos = React.useCallback(async (): Promise<void> => {
    if (!copiarHijosPadreOrigenId.trim() || !copiarHijosPadreDestinoId.trim()) {
      setError('Debes ingresar el ID padre origen y el ID padre destino para copiar hijos.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision('Iniciando copia de hijos en Relaciones Documentos...');

    try {
      const resultado = await copiarHijosRelacionesDocumentos({
        context,
        padreOrigenId: copiarHijosPadreOrigenId,
        padreDestinoId: copiarHijosPadreDestinoId,
        log: appendLog
      });

      descargarArchivo(resultado.blob, resultado.fileName);
      appendLog('✅ Copia de hijos terminada. Hijos origen: ' + resultado.totalHijos);
      appendLog('✅ Relaciones creadas: ' + resultado.creadas + ' | SKIP: ' + resultado.omitidas + ' | ERROR: ' + resultado.errores);
      appendLog('✅ DocPadres actualizados: ' + resultado.docPadresActualizados + ' | SKIP: ' + resultado.docPadresOmitidos);
      appendLog('📥 Archivo generado: ' + resultado.fileName);
    } catch (copyError) {
      const errorMessage = copyError instanceof Error ? copyError.message : String(copyError);
      setError(errorMessage);
      appendLog('❌ Error copiando hijos: ' + errorMessage);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, copiarHijosPadreDestinoId, copiarHijosPadreOrigenId, descargarArchivo]);

  const ejecutarModificarAprobadores = React.useCallback(async (): Promise<void> => {
    setError(null);
    setIsRunning(true);
    setLogRevision('Iniciando actualización de ModificarAprobadores para todos los revisores impactados...');

    try {
      const resultado = await modificarAprobadores({
        context,
        log: appendLog
      });

      descargarArchivo(resultado.blob, resultado.fileName);
      appendLog(
        `✅ ModificarAprobadores terminado. ` +
        `Encontrados: ${resultado.totalEncontrados} | ` +
        `Solicitudes evaluadas: ${resultado.totalSolicitudesEvaluadas} | ` +
        `Solicitudes vigentes omitidas: ${resultado.totalSolicitudesVigentesOmitidas} | ` +
        `Con cambios: ${resultado.totalCambiarian} | ` +
        `Actualizados: ${resultado.totalActualizados} | ` +
        `Omitidos: ${resultado.totalOmitidos} | ` +
        `Error: ${resultado.totalError}`
      );
      appendLog(`📥 Archivo generado: ${resultado.fileName}`);
    } catch (modifyError) {
      const errorMessage = modifyError instanceof Error ? modifyError.message : String(modifyError);
      setError(errorMessage);
      appendLog('❌ Error en ModificarAprobadores: ' + errorMessage);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, descargarArchivo]);

  const ejecutarRollbackModificarAprobadores = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar el Excel de resultado de ModificarAprobadores antes de ejecutar el rollback.');
      appendLog('No se pudo iniciar el rollback de ModificarAprobadores porque no hay Excel seleccionado.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision(`Iniciando rollback de ModificarAprobadores desde el archivo: ${excelFile.fileName}`);

    try {
      const resultado = await rollbackModificarAprobadores({
        context,
        excelFile,
        log: appendLog
      });

      descargarArchivo(resultado.blob, resultado.fileName);
      appendLog(
        `✅ Rollback ModificarAprobadores terminado. ` +
        `Filas: ${resultado.totalFilas} | ` +
        `Restaurados: ${resultado.totalRestaurados} | ` +
        `Omitidos: ${resultado.totalOmitidos} | ` +
        `Error: ${resultado.totalError}`
      );
      appendLog(`📥 Archivo generado: ${resultado.fileName}`);
    } catch (rollbackError) {
      const errorMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
      setError(errorMessage);
      appendLog('❌ Error en rollback ModificarAprobadores: ' + errorMessage);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, descargarArchivo, excelFile]);

  const ejecutarReporteSolicitudesNoAntonio = React.useCallback(async (): Promise<void> => {
    setError(null);
    setIsRunning(true);
    setLogRevision('Iniciando reporte de solicitudes no creadas por Antonio Sánchez Panta...');

    try {
      const resultado = await exportarSolicitudesNoCreadasPorAntonio({
        context,
        log: appendLog
      });

      descargarArchivo(resultado.blob, resultado.fileName);
      appendLog(`✅ Reporte generado. Total solicitudes leidas: ${resultado.totalSolicitudes}`);
      appendLog(`✅ Incluidas en Excel: ${resultado.totalFiltradas}`);
      appendLog(`ℹ️ Excluidas por Antonio Sánchez Panta: ${resultado.totalExcluidasAntonio}`);
      appendLog(`📎 Incluidas con hijos: ${resultado.conHijos}`);
      appendLog(`📥 Archivo generado: ${resultado.fileName}`);
    } catch (reportError) {
      const errorMessage = reportError instanceof Error ? reportError.message : String(reportError);
      setError(errorMessage);
      appendLog(`❌ Error generando reporte de solicitudes: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, descargarArchivo]);

  const ejecutarReporteSolicitudesDuplicadas = React.useCallback(async (): Promise<void> => {
    setError(null);
    setIsRunning(true);
    setLogRevision('Iniciando reporte de solicitudes con CodigoDocumento duplicado y distinto nombre...');

    try {
      const resultado = await buscarSolicitudesDuplicadas({
        context,
        log: appendLog
      });

      descargarArchivo(resultado.blob, resultado.fileName);
      appendLog(`✅ Reporte generado. Total solicitudes leidas: ${resultado.totalSolicitudes}`);
      appendLog(`⚠️ Grupos duplicados con distinto nombre: ${resultado.duplicatedGroups}`);
      appendLog(`⚠️ Filas incluidas en Excel: ${resultado.duplicatedRows}`);
      appendLog(`ℹ️ Filas no vigentes dentro del reporte: ${resultado.nonCurrentRows}`);
      appendLog(`📥 Archivo generado: ${resultado.fileName}`);
    } catch (reportError) {
      const errorMessage = reportError instanceof Error ? reportError.message : String(reportError);
      setError(errorMessage);
      appendLog(`❌ Error generando reporte de duplicados: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, descargarArchivo]);

  const ejecutarCambioInstancia = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar un Excel antes de ejecutar el cambio de instancia.');
      appendLog('No se pudo iniciar el cambio de instancia porque no hay Excel seleccionado.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision('Iniciando cambio masivo de instancia desde Excel...');

    try {
      const resultado = await ejecutarCambioInstanciaDesdeExcel({
        context,
        excelFile,
        log: appendLog
      });

      descargarArchivo(resultado.blob, resultado.fileName);
      appendLog(`✅ Cambio de instancia terminado. Filas elegibles: ${resultado.processed}`);
      appendLog(`✅ Actualizadas: ${resultado.updated} | SKIP: ${resultado.skipped} | ERROR: ${resultado.error}`);
      appendLog(`📥 Archivo generado: ${resultado.fileName}`);
    } catch (changeError) {
      const errorMessage = changeError instanceof Error ? changeError.message : String(changeError);
      setError(errorMessage);
      appendLog(`❌ Error en cambio de instancia: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, descargarArchivo, excelFile]);

  const ejecutarCorreccionCodigosDuplicados = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar un Excel antes de ejecutar la corrección de códigos duplicados.');
      appendLog('No se pudo iniciar la corrección de códigos duplicados porque no hay Excel seleccionado.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision('Iniciando corrección de códigos de documento duplicados desde Excel...');

    try {
      const resultado = await ejecutarCorreccionCodigosDuplicadosDesdeExcel({
        context,
        excelFile,
        log: appendLog
      });

      descargarArchivo(resultado.blob, resultado.fileName);
      appendLog(`✅ Corrección de códigos duplicados terminada. Filas objetivo: ${resultado.processed}`);
      appendLog(`✅ Actualizadas: ${resultado.updated} | SKIP/KEEP: ${resultado.skipped} | ERROR: ${resultado.error}`);
      appendLog(`📥 Archivo generado: ${resultado.fileName}`);
    } catch (changeError) {
      const errorMessage = changeError instanceof Error ? changeError.message : String(changeError);
      setError(errorMessage);
      appendLog(`❌ Error en corrección de códigos duplicados: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, descargarArchivo, excelFile]);

  const ejecutarBuscarDiagramasSinCodigo = React.useCallback(async (): Promise<void> => {
    setError(null);
    setIsRunning(true);
    setLogRevision('Iniciando utilitario Buscar Diagrama de Flujo sin codigo...');

    try {
      const resultado = await ejecutarBuscarDiagramaFlujoSinCodigo({
        context,
        log: appendLog
      });

      appendLog(`✅ Utilitario terminado. Diagramas sin codigo detectados: ${resultado.totalDiagramasSinCodigo}`);
      appendLog(`✅ Solicitudes procesadas: ${resultado.solicitudesProcesadas}`);
      appendLog(`✅ Diagramas corregidos: ${resultado.diagramasCorregidos}`);
      appendLog(`✅ Solicitudes regeneradas/publicadas: ${resultado.solicitudesRegeneradas}`);
      appendLog(`ℹ️ SKIP: ${resultado.skipped} | ERROR: ${resultado.error}`);
      appendLog(`📂 Archivos TEMP generados: ${resultado.tempFileUrls.length}`);
    } catch (runError) {
      const errorMessage = runError instanceof Error ? runError.message : String(runError);
      setError(errorMessage);
      appendLog(`❌ Error en Buscar Diagrama de Flujo sin codigo: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context]);

  const ejecutarFase1 = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar el Excel revisado antes de ejecutar la Fase 1.');
      return;
    }

    if (!sourceFolderUrl) {
      setError('Debes seleccionar la carpeta SharePoint donde están los archivos origen.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision(`Iniciando Fase 1 para el archivo: ${excelFile.fileName}`);

    try {
      const resultado = await ejecutarFase1DocumentosSinHijosNiFlujos({
        context,
        excelFile,
        sourceFolderServerRelativeUrl: sourceFolderUrl,
        tempWordBaseFolderServerRelativeUrl: '/sites/SistemadeGestionDocumental/Procesos/TEMP_MIGRACION_WORD',
        log: appendLog
      });

      setCreatedSolicitudIds(resultado.createdSolicitudIds);
      setOldSolicitudIds(resultado.oldSolicitudIds);
      setTempFileUrls(resultado.tempFileUrls);
      setLastExecutedPhase('fase1');

      appendLog(`✅ Fase 1 terminada. Procesados=${resultado.processed}`);
      appendLog(`✅ OK=${resultado.ok} | SKIP=${resultado.skipped} | ERROR=${resultado.error}`);
      appendLog(`📄 Nuevas solicitudes creadas=${resultado.createdSolicitudIds.length}`);
      appendLog(`📂 Archivos TEMP generados=${resultado.tempFileUrls.length}`);
    } catch (fase1Error) {
      const errorMessage = fase1Error instanceof Error ? fase1Error.message : String(fase1Error);
      setError(errorMessage);
      appendLog(`❌ Error en Fase 1: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, excelFile, sourceFolderUrl]);

  const ejecutarRollback = React.useCallback(async (): Promise<void> => {
    if (lastExecutedPhase === 'fase2') {
      if (!fase2RollbackEntries.length) {
        setError('No hay resultados exitosos de Fase 2 para revertir.');
        appendLog('ℹ️ Rollback de Fase 2 omitido: no hubo publicaciones exitosas para revertir.');
        return;
      }

      setError(null);
      setIsRunning(true);
      appendLog('🧨 Iniciando rollback de Fase 2...');

      try {
        await rollbackModificacionFase2({
          context,
          webUrl: context.pageContext.web.absoluteUrl,
          entries: fase2RollbackEntries,
          log: appendLog
        });

        setFase2RollbackEntries([]);
        setLastExecutedPhase(null);
        appendLog('✅ Rollback de Fase 2 finalizado.');
      } catch (rollbackError) {
        const errorMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
        setError(errorMessage);
        appendLog(`❌ Error durante el rollback de Fase 2: ${errorMessage}`);
      } finally {
        setIsRunning(false);
      }
      return;
    }

    if (lastExecutedPhase === 'fase4') {
      if (!fase4RollbackEntries.length) {
        setError('No hay resultados exitosos de Fase 4 para revertir.');
        appendLog('ℹ️ Rollback de Fase 4 omitido: no hubo publicaciones exitosas para revertir.');
        return;
      }

      setError(null);
      setIsRunning(true);
      appendLog('🧨 Iniciando rollback de Fase 4...');

      try {
        await rollbackModificacionFase2({
          context,
          webUrl: context.pageContext.web.absoluteUrl,
          entries: fase4RollbackEntries,
          log: appendLog
        });

        setFase4RollbackEntries([]);
        setLastExecutedPhase(null);
        appendLog('✅ Rollback de Fase 4 finalizado.');
      } catch (rollbackError) {
        const errorMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
        setError(errorMessage);
        appendLog(`❌ Error durante el rollback de Fase 4: ${errorMessage}`);
      } finally {
        setIsRunning(false);
      }
      return;
    }

    if (lastExecutedPhase === 'fase3') {
      if (!fase3RollbackEntries.length) {
        setError('No hay resultados exitosos de Fase 3 para revertir.');
        appendLog('ℹ️ Rollback de Fase 3 omitido: no hubo cambios exitosos para revertir.');
        return;
      }

      setError(null);
      setIsRunning(true);
      appendLog('🧨 Iniciando rollback de Fase 3...');

      try {
        await rollbackModificacionFase3({
          context,
          webUrl: context.pageContext.web.absoluteUrl,
          entries: fase3RollbackEntries,
          log: appendLog
        });

        setFase3RollbackEntries([]);
        setLastExecutedPhase(null);
        appendLog('✅ Rollback de Fase 3 finalizado.');
      } catch (rollbackError) {
        const errorMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
        setError(errorMessage);
        appendLog(`❌ Error durante el rollback de Fase 3: ${errorMessage}`);
      } finally {
        setIsRunning(false);
      }
      return;
    }

    if (lastExecutedPhase === 'fase6') {
      if (!fase6RollbackEntries.length && !fase6DiagramRollbackEntries.length) {
        setError('No hay resultados exitosos de Fase 6 para revertir.');
        appendLog('ℹ️ Rollback de Fase 6 omitido: no hubo bajas exitosas para revertir.');
        return;
      }

      setError(null);
      setIsRunning(true);
      appendLog('🧨 Iniciando rollback de Fase 6...');

      try {
        await rollbackModificacionFase6({
          context,
          webUrl: context.pageContext.web.absoluteUrl,
          entries: fase6RollbackEntries,
          diagramEntries: fase6DiagramRollbackEntries,
          log: appendLog
        });

        setFase6RollbackEntries([]);
        setFase6DiagramRollbackEntries([]);
        setLastExecutedPhase(null);
        appendLog('✅ Rollback de Fase 6 finalizado.');
      } catch (rollbackError) {
        const errorMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
        setError(errorMessage);
        appendLog(`❌ Error durante el rollback de Fase 6: ${errorMessage}`);
      } finally {
        setIsRunning(false);
      }
      return;
    }

    if (lastExecutedPhase === 'fase7') {
      if (!fase7RollbackEntries.length) {
        setError('No hay resultados exitosos de Fase 7 para revertir.');
        appendLog('ℹ️ Rollback de Fase 7 omitido: no hubo diagramas reemplazados para revertir.');
        return;
      }

      setError(null);
      setIsRunning(true);
      appendLog('🧨 Iniciando rollback de Fase 7...');

      try {
        await rollbackModificacionFase7({
          context,
          webUrl: context.pageContext.web.absoluteUrl,
          entries: fase7RollbackEntries,
          log: appendLog
        });

        setFase7RollbackEntries([]);
        setLastExecutedPhase(null);
        appendLog('✅ Rollback de Fase 7 finalizado.');
      } catch (rollbackError) {
        const errorMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
        setError(errorMessage);
        appendLog(`❌ Error durante el rollback de Fase 7: ${errorMessage}`);
      } finally {
        setIsRunning(false);
      }
      return;
    }

    if (lastExecutedPhase === 'fase8') {
      if (!fase8RollbackEntries.length) {
        setError('No hay resultados exitosos de Fase 8 para revertir.');
        appendLog('ℹ️ Rollback de Fase 8 omitido: no hubo altas exitosas para revertir.');
        return;
      }

      setError(null);
      setIsRunning(true);
      appendLog('🧨 Iniciando rollback de Fase 8...');

      try {
        await rollbackModificacionFase8({
          context,
          webUrl: context.pageContext.web.absoluteUrl,
          entries: fase8RollbackEntries,
          log: appendLog
        });

        setFase8RollbackEntries([]);
        setLastExecutedPhase(null);
        appendLog('✅ Rollback de Fase 8 finalizado.');
      } catch (rollbackError) {
        const errorMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
        setError(errorMessage);
        appendLog(`❌ Error durante el rollback de Fase 8: ${errorMessage}`);
      } finally {
        setIsRunning(false);
      }
      return;
    }

    if (lastExecutedPhase === 'correccion2460') {
      if (!fase8RollbackEntries.length) {
        setError('No hay resultados exitosos de la corrección 2460 para revertir.');
        appendLog('ℹ️ Rollback de corrección 2460 omitido: no hubo cambios exitosos para revertir.');
        return;
      }

      setError(null);
      setIsRunning(true);
      appendLog('🧨 Iniciando rollback de corrección 2460...');

      try {
        await rollbackModificacionFase8({
          context,
          webUrl: context.pageContext.web.absoluteUrl,
          entries: fase8RollbackEntries,
          log: appendLog
        });

        setFase8RollbackEntries([]);
        setLastExecutedPhase(null);
        appendLog('✅ Rollback de corrección 2460 finalizado.');
      } catch (rollbackError) {
        const errorMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
        setError(errorMessage);
        appendLog(`❌ Error durante el rollback de corrección 2460: ${errorMessage}`);
      } finally {
        setIsRunning(false);
      }
      return;
    }

    if (lastExecutedPhase === 'fase9') {
      if (!fase9RollbackEntries.length) {
        setError('No hay resultados exitosos de Fase 9 para revertir.');
        appendLog('ℹ️ Rollback de Fase 9 omitido: no hubo bajas exitosas para revertir.');
        return;
      }

      setError(null);
      setIsRunning(true);
      appendLog('🧨 Iniciando rollback de Fase 9...');

      try {
        await rollbackModificacionFase8({
          context,
          webUrl: context.pageContext.web.absoluteUrl,
          entries: fase9RollbackEntries,
          log: appendLog
        });

        setFase9RollbackEntries([]);
        setLastExecutedPhase(null);
        appendLog('✅ Rollback de Fase 9 finalizado.');
      } catch (rollbackError) {
        const errorMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
        setError(errorMessage);
        appendLog(`❌ Error durante el rollback de Fase 9: ${errorMessage}`);
      } finally {
        setIsRunning(false);
      }
      return;
    }

    if (!createdSolicitudIds.length && !oldSolicitudIds.length && !tempFileUrls.length) {
      setError('No hay resultados de Fase 1 para revertir.');
      return;
    }

    setError(null);
    setIsRunning(true);
    appendLog('🧨 Iniciando rollback de Fase 1...');

    try {
      await rollbackModificacionFase1({
        context,
        webUrl: context.pageContext.web.absoluteUrl,
        createdSolicitudIds,
        oldSolicitudIds,
        tempFileUrls,
        log: appendLog
      });

      setCreatedSolicitudIds([]);
      setOldSolicitudIds([]);
      setTempFileUrls([]);
      setLastExecutedPhase(null);
      appendLog('✅ Rollback de Fase 1 finalizado.');
    } catch (rollbackError) {
      const errorMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
      setError(errorMessage);
      appendLog(`❌ Error durante el rollback: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, createdSolicitudIds, fase2RollbackEntries, fase3RollbackEntries, fase4RollbackEntries, fase6DiagramRollbackEntries, fase6RollbackEntries, fase7RollbackEntries, fase8RollbackEntries, fase9RollbackEntries, lastExecutedPhase, oldSolicitudIds, tempFileUrls]);

  const ejecutarFase2 = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar el Excel FASE1_WORD antes de ejecutar la Fase 2.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision(`Iniciando Fase 2 para el archivo: ${excelFile.fileName}`);

    try {
      const resultado = await ejecutarFase2PublicacionDocumentosSinHijos({
        context,
        excelFile,
        log: appendLog
      });

      setFase2RollbackEntries(resultado.rollbackEntries);
      setFase3RollbackEntries([]);
      setFase4RollbackEntries([]);
      setFase6RollbackEntries([]);
      setFase6DiagramRollbackEntries([]);
      setFase7RollbackEntries([]);
      setFase8RollbackEntries([]);
      setLastExecutedPhase('fase2');
      appendLog(`✅ Fase 2 terminada. Procesados=${resultado.processed}`);
      appendLog(`✅ OK=${resultado.ok} | SKIP=${resultado.skipped} | ERROR=${resultado.error}`);
      appendLog(`📄 Registros rollback Fase 2=${resultado.rollbackEntries.length}`);
    } catch (fase2Error) {
      const errorMessage = fase2Error instanceof Error ? fase2Error.message : String(fase2Error);
      setError(errorMessage);
      appendLog(`❌ Error en Fase 2: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, excelFile]);

  const ejecutarFase3 = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar el Excel revisado antes de ejecutar la Fase 3.');
      return;
    }

    if (!sourceFolderUrl) {
      setError('Debes seleccionar la carpeta SharePoint donde están los archivos origen.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision(`Iniciando Fase 3 para el archivo: ${excelFile.fileName}`);

    try {
      const resultado = await ejecutarFase3PadresConHijosYFlujos({
        context,
        excelFile,
        sourceFolderServerRelativeUrl: sourceFolderUrl,
        tempWordBaseFolderServerRelativeUrl: '/sites/SistemadeGestionDocumental/Procesos/TEMP_MIGRACION_WORD',
        log: appendLog
      });

      setFase3RollbackEntries(resultado.rollbackEntries);
      setFase2RollbackEntries([]);
      setFase4RollbackEntries([]);
      setFase6RollbackEntries([]);
      setFase6DiagramRollbackEntries([]);
      setFase7RollbackEntries([]);
      setFase8RollbackEntries([]);
      setLastExecutedPhase('fase3');
      appendLog(`✅ Fase 3 terminada. Procesados=${resultado.processed}`);
      appendLog(`✅ OK=${resultado.ok} | SKIP=${resultado.skipped} | ERROR=${resultado.error}`);
      appendLog(`📄 Nuevas solicitudes creadas Fase 3=${resultado.createdSolicitudIds.length}`);
      appendLog(`📂 Archivos TEMP generados Fase 3=${resultado.tempFileUrls.length}`);
      appendLog(`📄 Registros rollback Fase 3=${resultado.rollbackEntries.length}`);
    } catch (fase3Error) {
      const errorMessage = fase3Error instanceof Error ? fase3Error.message : String(fase3Error);
      setError(errorMessage);
      appendLog(`❌ Error en Fase 3: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, excelFile, sourceFolderUrl]);

  const ejecutarFase4 = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar el Excel FASE3_WORD antes de ejecutar la Fase 4.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision(`Iniciando Fase 4 para el archivo: ${excelFile.fileName}`);

    try {
      const resultado = await ejecutarFase4PublicacionHijos({
        context,
        excelFile,
        log: appendLog
      });

      setFase4RollbackEntries(resultado.rollbackEntries);
      setFase2RollbackEntries([]);
      setFase3RollbackEntries([]);
      setFase6RollbackEntries([]);
      setFase6DiagramRollbackEntries([]);
      setFase7RollbackEntries([]);
      setFase8RollbackEntries([]);
      setLastExecutedPhase('fase4');
      appendLog(`✅ Fase 4 terminada. Procesados=${resultado.processed}`);
      appendLog(`✅ OK=${resultado.ok} | SKIP=${resultado.skipped} | ERROR=${resultado.error}`);
      appendLog(`📄 Registros rollback Fase 4=${resultado.rollbackEntries.length}`);
      appendLog('ℹ️ Fase 4 publica padres desde TEMP y actualiza la referencia del padre en los hijos vigentes.');
    } catch (fase4Error) {
      const errorMessage = fase4Error instanceof Error ? fase4Error.message : String(fase4Error);
      setError(errorMessage);
      appendLog(`❌ Error en Fase 4: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, excelFile]);

  const ejecutarFase5 = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar el Excel revisado antes de ejecutar la Fase 5.');
      return;
    }

    if (!sourceFolderUrl) {
      setError('Debes seleccionar la carpeta SharePoint donde están los archivos origen.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision(`Iniciando Fase 5 para el archivo: ${excelFile.fileName}`);

    try {
      const resultado = await ejecutarFase5HijosConPadre({
        context,
        excelFile,
        sourceFolderServerRelativeUrl: sourceFolderUrl,
        tempWordBaseFolderServerRelativeUrl: '/sites/SistemadeGestionDocumental/Procesos/TEMP_MIGRACION_WORD',
        log: appendLog
      });

      setCreatedSolicitudIds([]);
      setOldSolicitudIds([]);
      setTempFileUrls([]);
      setFase2RollbackEntries([]);
      setFase3RollbackEntries([]);
      setFase4RollbackEntries([]);
      setFase6RollbackEntries([]);
      setFase6DiagramRollbackEntries([]);
      setFase7RollbackEntries([]);
      setLastExecutedPhase(null);

      appendLog(`✅ Fase 5 terminada. Procesados=${resultado.processed}`);
      appendLog(`✅ OK=${resultado.ok} | SKIP=${resultado.skipped} | ERROR=${resultado.error}`);
      appendLog(`📄 Nuevas solicitudes creadas Fase 5=${resultado.createdSolicitudIds.length}`);
      appendLog(`📂 Archivos TEMP generados Fase 5=${resultado.tempFileUrls.length}`);
      appendLog('ℹ️ Fase 5 crea el hijo nuevo, publica en Procesos, envía el anterior a Históricos y mantiene el mismo docPadres.');
      appendLog('ℹ️ Fase 5 no deja rollback automático habilitado.');
    } catch (fase5Error) {
      const errorMessage = fase5Error instanceof Error ? fase5Error.message : String(fase5Error);
      setError(errorMessage);
      appendLog(`❌ Error en Fase 5: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, excelFile, sourceFolderUrl]);

  const ejecutarFase6 = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar el Excel revisado antes de ejecutar la Fase 6.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision(`Iniciando Fase 6 para el archivo: ${excelFile.fileName}`);

    try {
      const resultado = await ejecutarFase6BajaDocumentos({
        context,
        excelFile,
        log: appendLog
      });

      setCreatedSolicitudIds([]);
      setOldSolicitudIds([]);
      setTempFileUrls([]);
      setFase2RollbackEntries([]);
      setFase3RollbackEntries([]);
      setFase4RollbackEntries([]);
      setFase7RollbackEntries([]);
      setFase8RollbackEntries([]);
      setFase6RollbackEntries(resultado.rollbackEntries);
      setFase6DiagramRollbackEntries(resultado.diagramRollbackEntries);
      setLastExecutedPhase('fase6');
      appendLog(`✅ Fase 6 terminada. Procesados=${resultado.processed}`);
      appendLog(`✅ OK=${resultado.ok} | SKIP=${resultado.skipped} | ERROR=${resultado.error}`);
      appendLog(`📄 Registros rollback Fase 6=${resultado.rollbackEntries.length}`);
      appendLog(`🧭 Diagramas respaldados Fase 6=${resultado.diagramRollbackEntries.length}`);
      appendLog(`📥 Excel final generado: ${resultado.reportFileName}`);
    } catch (fase6Error) {
      const errorMessage = fase6Error instanceof Error ? fase6Error.message : String(fase6Error);
      setError(errorMessage);
      appendLog(`❌ Error en Fase 6: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, excelFile]);

  const ejecutarFase7 = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar el Excel revisado antes de ejecutar la Fase 7.');
      return;
    }

    if (!sourceFolderUrl) {
      setError('Debes seleccionar la carpeta SharePoint donde están los archivos BPM origen.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision(`Iniciando Fase 7 para el archivo: ${excelFile.fileName}`);

    try {
      const resultado = await ejecutarFase7ModificacionDiagramas({
        context,
        excelFile,
        sourceFolderServerRelativeUrl: sourceFolderUrl,
        log: appendLog
      });

      setCreatedSolicitudIds([]);
      setOldSolicitudIds([]);
      setTempFileUrls([]);
      setFase2RollbackEntries([]);
      setFase3RollbackEntries([]);
      setFase4RollbackEntries([]);
      setFase6RollbackEntries([]);
      setFase6DiagramRollbackEntries([]);
      setFase8RollbackEntries([]);
      setFase9RollbackEntries([]);
      setFase7RollbackEntries(resultado.rollbackEntries);
      setLastExecutedPhase('fase7');
      appendLog(`✅ Fase 7 terminada. Procesados=${resultado.processed}`);
      appendLog(`✅ OK=${resultado.ok} | SKIP=${resultado.skipped} | ERROR=${resultado.error}`);
      appendLog(`📄 Registros rollback Fase 7=${resultado.rollbackEntries.length}`);
      appendLog(`📥 Excel final generado: ${resultado.reportFileName}`);
    } catch (fase7Error) {
      const errorMessage = fase7Error instanceof Error ? fase7Error.message : String(fase7Error);
      setError(errorMessage);
      appendLog(`❌ Error en Fase 7: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, excelFile, sourceFolderUrl]);

  const ejecutarFase8 = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar el Excel revisado antes de ejecutar la Fase 8.');
      return;
    }

    if (!sourceFolderUrl) {
      setError('Debes seleccionar la carpeta SharePoint donde están los archivos BPM origen.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision('Iniciando Fase 8 para el archivo: ' + excelFile.fileName);

    try {
      const resultado = await ejecutarFase8AltaDiagramaSolicitud({
        context,
        excelFile,
        sourceFolderServerRelativeUrl: sourceFolderUrl,
        log: appendLog
      });

      setCreatedSolicitudIds([]);
      setOldSolicitudIds([]);
      setTempFileUrls([]);
      setFase2RollbackEntries([]);
      setFase3RollbackEntries([]);
      setFase4RollbackEntries([]);
      setFase6RollbackEntries([]);
      setFase6DiagramRollbackEntries([]);
      setFase7RollbackEntries([]);
      setFase8RollbackEntries(resultado.rollbackEntries);
      setFase9RollbackEntries([]);
      setLastExecutedPhase('fase8');

      appendLog('✅ Fase 8 terminada. Procesados=' + resultado.processed);
      appendLog('✅ OK=' + resultado.ok + ' | SKIP=' + resultado.skipped + ' | ERROR=' + resultado.error);
      appendLog('📄 Registros rollback Fase 8=' + resultado.rollbackEntries.length);
      appendLog('📥 Excel final generado: ' + resultado.reportFileName);
      appendLog('ℹ️ Fase 8 versiona la solicitud padre vigente, crea el nuevo diagrama, regenera el Word y publica nuevamente en Procesos.');
      appendLog('ℹ️ El rollback de Fase 8 restaura solicitud vigente, relaciones, diagramas y publicación en Procesos.');
    } catch (fase8Error) {
      const errorMessage = fase8Error instanceof Error ? fase8Error.message : String(fase8Error);
      setError(errorMessage);
      appendLog('❌ Error en Fase 8: ' + errorMessage);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, excelFile, sourceFolderUrl]);

  const ejecutarCorreccion2460 = React.useCallback(async (): Promise<void> => {
    setError(null);
    setIsRunning(true);
    setLogRevision('Iniciando corrección puntual para la solicitud 2460: Procedimiento de Registro Manual de Asientos Contables');

    try {
      const resultado = await ejecutarCorreccionSolicitud2460({
        context,
        log: appendLog
      });

      setCreatedSolicitudIds([]);
      setOldSolicitudIds([]);
      setTempFileUrls([]);
      setFase2RollbackEntries([]);
      setFase3RollbackEntries([]);
      setFase4RollbackEntries([]);
      setFase6RollbackEntries([]);
      setFase6DiagramRollbackEntries([]);
      setFase7RollbackEntries([]);
      setFase8RollbackEntries(resultado.rollbackEntries);
      setFase9RollbackEntries([]);
      setLastExecutedPhase('correccion2460');

      appendLog('✅ Corrección 2460 terminada. Procesados=' + resultado.processed);
      appendLog('✅ OK=' + resultado.ok + ' | ERROR=' + resultado.error);
      appendLog('🆔 Solicitud origen=' + resultado.solicitudOrigenId + ' | Solicitud nueva=' + resultado.solicitudNuevaId);
      appendLog('📄 Publicación nueva=' + resultado.rutaNuevoPublicado);
      appendLog('ℹ️ La corrección 2460 regenera la solicitud vigente usando Relación documentos y Diagramas de Flujo, publica nuevamente en Procesos y reasigna hijos al nuevo padre.');
      appendLog('ℹ️ El rollback de esta corrección reutiliza la restauración de Fase 8.');
    } catch (correccionError) {
      const errorMessage = correccionError instanceof Error ? correccionError.message : String(correccionError);
      setError(errorMessage);
      appendLog('❌ Error en corrección 2460: ' + errorMessage);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context]);

  const ejecutarFase9 = React.useCallback(async (): Promise<void> => {
    if (!excelFile) {
      setError('Debes seleccionar el Excel revisado antes de ejecutar la Fase 9.');
      return;
    }

    setError(null);
    setIsRunning(true);
    setLogRevision('Iniciando Fase 9 para el archivo: ' + excelFile.fileName);

    try {
      const resultado = await ejecutarFase9BajaDiagramaSolicitud({
        context,
        excelFile,
        log: appendLog
      });

      setCreatedSolicitudIds([]);
      setOldSolicitudIds([]);
      setTempFileUrls([]);
      setFase2RollbackEntries([]);
      setFase3RollbackEntries([]);
      setFase4RollbackEntries([]);
      setFase6RollbackEntries([]);
      setFase6DiagramRollbackEntries([]);
      setFase7RollbackEntries([]);
      setFase8RollbackEntries([]);
      setFase9RollbackEntries(resultado.rollbackEntries);
      setLastExecutedPhase('fase9');

      appendLog('✅ Fase 9 terminada. Procesados=' + resultado.processed);
      appendLog('✅ OK=' + resultado.ok + ' | SKIP=' + resultado.skipped + ' | ERROR=' + resultado.error);
      appendLog('📄 Registros rollback Fase 9=' + resultado.rollbackEntries.length);
      appendLog('📥 Excel final generado: ' + resultado.reportFileName);
      appendLog('ℹ️ Fase 9 versiona la solicitud padre vigente, excluye el diagrama indicado, regenera el Word y publica nuevamente en Procesos.');
      appendLog('ℹ️ El rollback de Fase 9 restaura solicitud vigente, relaciones, diagramas y publicación en Procesos.');
    } catch (fase9Error) {
      const errorMessage = fase9Error instanceof Error ? fase9Error.message : String(fase9Error);
      setError(errorMessage);
      appendLog('❌ Error en Fase 9: ' + errorMessage);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, excelFile]);

  const copiarLog = React.useCallback(() => {
    void navigator.clipboard.writeText(logRevision);
  }, [logRevision]);

  const rollbackSummary = React.useMemo(() => {
    if (lastExecutedPhase === 'fase2') {
      return fase2RollbackEntries.length
        ? `Fase 2 lista para rollback | Registros ${fase2RollbackEntries.length}`
        : 'Fase 2 sin publicaciones exitosas para rollback.';
    }

    if (lastExecutedPhase === 'fase4') {
      return fase4RollbackEntries.length
        ? `Fase 4 lista para rollback | Registros ${fase4RollbackEntries.length}`
        : 'Fase 4 sin publicaciones exitosas para rollback.';
    }

    if (lastExecutedPhase === 'fase3') {
      return fase3RollbackEntries.length
        ? `Fase 3 lista para rollback | Registros ${fase3RollbackEntries.length}`
        : 'Fase 3 sin cambios exitosos para rollback.';
    }

    if (lastExecutedPhase === 'fase6') {
      return fase6RollbackEntries.length || fase6DiagramRollbackEntries.length
        ? `Fase 6 lista para rollback | Documentos ${fase6RollbackEntries.length} | Diagramas ${fase6DiagramRollbackEntries.length}`
        : 'Fase 6 sin cambios exitosos para rollback.';
    }

    if (lastExecutedPhase === 'fase7') {
      return fase7RollbackEntries.length
        ? `Fase 7 lista para rollback | Diagramas ${fase7RollbackEntries.length}`
        : 'Fase 7 sin cambios exitosos para rollback.';
    }

    if (lastExecutedPhase === 'fase8') {
      return fase8RollbackEntries.length
        ? `Fase 8 lista para rollback | Registros ${fase8RollbackEntries.length}`
        : 'Fase 8 sin cambios exitosos para rollback.';
    }

    if (lastExecutedPhase === 'fase9') {
      return fase9RollbackEntries.length
        ? `Fase 9 lista para rollback | Registros ${fase9RollbackEntries.length}`
        : 'Fase 9 sin cambios exitosos para rollback.';
    }

    if (lastExecutedPhase === 'correccion2460') {
      return fase8RollbackEntries.length
        ? `Corrección 2460 lista para rollback | Registros ${fase8RollbackEntries.length}`
        : 'Corrección 2460 sin cambios exitosos para rollback.';
    }

    if (!createdSolicitudIds.length && !oldSolicitudIds.length && !tempFileUrls.length) {
      return 'Aun no hay resultados para revertir.';
    }

    return `ID modificados ${oldSolicitudIds.length} | ID nuevos ${createdSolicitudIds.length} | TEMP ${tempFileUrls.length}`;
  }, [createdSolicitudIds.length, fase2RollbackEntries.length, fase3RollbackEntries.length, fase4RollbackEntries.length, fase6DiagramRollbackEntries.length, fase6RollbackEntries.length, fase7RollbackEntries.length, fase8RollbackEntries.length, fase9RollbackEntries.length, lastExecutedPhase, oldSolicitudIds.length, tempFileUrls.length]);

  return (
    <section className={`${styles.modificacion} ${hasTeamsContext ? styles.teams : ''} ${isDarkTheme ? styles.dark : ''}`}>
      <Stack tokens={{ childrenGap: 24 }}>
        <div className={styles.hero}>
          <div>
            <Text variant="xxLarge" className={styles.title}>Modificacion masiva</Text>
            <Text variant="large" className={styles.subtitle}>
              Prepara el Excel de modificación y ejecuta una revision inicial antes de conectar el flujo completo.
            </Text>
          </div>
          <div className={styles.heroBadge}>
            <Icon iconName="PageEdit" className={styles.heroBadgeIcon} />
            <Text variant="medium">WebPart de modificacion</Text>
          </div>
        </div>

        <div className={styles.layout}>
          <div className={styles.mainCard}>
            <Stack tokens={{ childrenGap: 18 }}>
              <div>
                <Text variant="xLarge" className={styles.sectionTitle}>Archivo de entrada</Text>
                <Text className={styles.sectionDescription}>
                  Selecciona el Excel que se usara para la revision de modificaciones.
                </Text>
              </div>

              <div className={styles.uploadPanel}>
                <div className={styles.uploadIconWrap}>
                  <Icon iconName="ExcelDocument" className={styles.uploadIcon} />
                </div>

                <div className={styles.uploadContent}>
                  <Text variant="large" className={styles.uploadTitle}>Excel de modificación</Text>
                  <Text className={styles.uploadHint}>
                    Selecciona el Excel desde la misma biblioteca de SharePoint.
                  </Text>

                  <FilePicker
                    context={context as any}
                    buttonLabel={excelFile ? 'Cambiar Excel' : 'Seleccionar Excel'}
                    buttonIcon="ExcelDocument"
                    onSave={onSaveExcel}
                    accepts={['.xlsx', '.xlsm']}
                    hideLinkUploadTab={true}
                    hideLocalUploadTab={true}
                    hideWebSearchTab={true}
                    hideStockImages={true}
                    hideRecentTab={false}
                    hideOrganisationalAssetTab={false}
                    hideOneDriveTab={true}
                    disabled={isRunning}
                  />

                  <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                    <DefaultButton text={isRunning ? 'Revisando...' : 'Revisar archivo'} onClick={revisarArchivo} disabled={!excelFile || isRunning} />

                    <DefaultButton
                      text={isRunning ? 'Cambiando instancia...' : 'Cambiar instancia desde Excel'}
                      onClick={() => { void ejecutarCambioInstancia(); }}
                      disabled={!excelFile || isRunning}
                    />

                    <DefaultButton
                      text={isRunning ? 'Corrigiendo DocPadres...' : 'Corregir DocPadres'}
                      onClick={() => { void ejecutarCorreccionDocPadres(); }}
                      disabled={!excelFile || isRunning}
                    />

                    <DefaultButton
                      text={isRunning ? 'Generando reporte...' : 'Reporte: no creadas por Antonio'}
                      onClick={() => { void ejecutarReporteSolicitudesNoAntonio(); }}
                      disabled={isRunning}
                    />

                    <DefaultButton
                      text={isRunning ? 'Generando duplicados...' : 'Reporte: duplicadas por codigo'}
                      onClick={() => { void ejecutarReporteSolicitudesDuplicadas(); }}
                      disabled={isRunning}
                    />

                    <DefaultButton
                      text={isRunning ? 'Corrigiendo codigos...' : 'Corregir codigos duplicados'}
                      onClick={() => { void ejecutarCorreccionCodigosDuplicados(); }}
                      disabled={!excelFile || isRunning}
                    />

                    <DefaultButton
                      text={isRunning ? 'Buscando diagramas...' : 'Buscar Diagrama de Flujo sin codigo'}
                      onClick={() => { void ejecutarBuscarDiagramasSinCodigo(); }}
                      disabled={isRunning}
                    />
                  </Stack>
                </div>
              </div>

              <div className={styles.statusCard}>
                <Text variant="mediumPlus" className={styles.statusTitle}>Estado actual</Text>
                <Text className={styles.statusValue}>
                  {excelFile ? excelFile.fileName : 'Aun no se ha seleccionado ningun archivo.'}
                </Text>
              </div>

              <div className={styles.utilityCard}>
                <Stack tokens={{ childrenGap: 12 }}>
                  <div>
                    <Text variant="mediumPlus" className={styles.statusTitle}>Copiar hijos</Text>
                    <Text className={styles.statusValue}>
                      Genera relaciones nuevas con los mismos hijos de un padre origen, usando otro ID como padre destino.
                    </Text>
                  </div>

                  <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                    <TextField
                      label="ID padre origen"
                      value={copiarHijosPadreOrigenId}
                      onChange={(_event, value) => setCopiarHijosPadreOrigenId(value || '')}
                      disabled={isRunning}
                      styles={{ root: { width: 180 } }}
                    />
                    <TextField
                      label="ID padre destino"
                      value={copiarHijosPadreDestinoId}
                      onChange={(_event, value) => setCopiarHijosPadreDestinoId(value || '')}
                      disabled={isRunning}
                      styles={{ root: { width: 180 } }}
                    />
                    <Stack verticalAlign="end">
                      <DefaultButton
                        text={isRunning ? 'Copiando hijos...' : 'Copiar hijos'}
                        onClick={() => { void ejecutarCopiarHijos(); }}
                        disabled={!copiarHijosPadreOrigenId.trim() || !copiarHijosPadreDestinoId.trim() || isRunning}
                      />
                    </Stack>
                  </Stack>
                </Stack>
              </div>

              <div className={styles.utilityCard}>
                <Stack tokens={{ childrenGap: 12 }}>
                  <div>
                    <Text variant="mediumPlus" className={styles.statusTitle}>ModificarAprobadores</Text>
                    <Text className={styles.statusValue}>
                      Recorre Aprobadores por Solicitudes, actualiza los checks de Revisor Impactado y genera un Excel con los valores anteriores para rollback.
                    </Text>
                  </div>

                  <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
                    <Stack verticalAlign="end">
                      <DefaultButton
                        text={isRunning ? 'Actualizando...' : 'Aplicar ModificarAprobadores'}
                        onClick={() => { void ejecutarModificarAprobadores(); }}
                        disabled={isRunning}
                      />
                    </Stack>
                    <Stack verticalAlign="end">
                      <DefaultButton
                        text={isRunning ? 'Restaurando...' : 'Rollback ModificarAprobadores'}
                        onClick={() => { void ejecutarRollbackModificarAprobadores(); }}
                        disabled={!excelFile || isRunning}
                      />
                    </Stack>
                  </Stack>
                </Stack>
              </div>

              {error && (
                <MessageBar messageBarType={MessageBarType.error}>
                  {error}
                </MessageBar>
              )}
            </Stack>
          </div>

          <div className={styles.sideCard}>
            <Stack tokens={{ childrenGap: 14 }}>
              <Text variant="large" className={styles.sideTitle}>Fase 1 operativa</Text>

              <FolderPicker
                context={context as any}
                label="Carpeta de archivos origen 2"
                required={false}
                canCreateFolders={false}
                rootFolder={{
                  Name: 'Sistema de Gestión Documental',
                  ServerRelativeUrl: '/sites/SistemadeGestionDocumental'
                }}
                onSelect={(folder: any) => {
                  const url = folder?.ServerRelativeUrl || '';
                  setSourceFolderUrl(url);
                  if (url) {
                    appendLog(`📁 Carpeta origen seleccionada: ${url}`);
                  }
                }}
              />

              <Text className={styles.sideDescription}>
                {sourceFolderUrl || 'Selecciona la carpeta SharePoint donde están los archivos Word, Excel o PPT.'}
              </Text>

              <DefaultButton
                text={isRunning ? 'Procesando Fase 1...' : 'Fase 1: Documentos sin hijos ni flujos'}
                onClick={() => { void ejecutarFase1(); }}
                disabled={!excelFile || !sourceFolderUrl || isRunning}
              />

              <DefaultButton
                text={isRunning ? 'Procesando Fase 2...' : 'Fase 2: Publicación Documentos sin hijos'}
                onClick={() => { void ejecutarFase2(); }}
                disabled={!excelFile || isRunning}
              />

              <DefaultButton
                text={isRunning ? 'Procesando Fase 3...' : 'Fase 3: Padres con hijos y/o flujos'}
                onClick={() => { void ejecutarFase3(); }}
                disabled={!excelFile || !sourceFolderUrl || isRunning}
              />

              <DefaultButton
                text={isRunning ? 'Procesando Fase 4...' : 'Fase 4: Publicar padres y actualizar hijos'}
                onClick={() => { void ejecutarFase4(); }}
                disabled={!excelFile || isRunning}
              />

              <DefaultButton
                text={isRunning ? 'Procesando Fase 5...' : 'Fase 5: Hijos con padre y publicación directa'}
                onClick={() => { void ejecutarFase5(); }}
                disabled={!excelFile || !sourceFolderUrl || isRunning}
              />

              <DefaultButton
                text={isRunning ? 'Procesando Fase 6...' : 'Fase 6: Documentos para baja'}
                onClick={() => { void ejecutarFase6(); }}
                disabled={!excelFile || isRunning}
              />


              <DefaultButton
                text={isRunning ? 'Procesando Fase 7...' : 'Fase 7: Modificación diagrama de flujo'}
                onClick={() => { void ejecutarFase7(); }}
                disabled={!excelFile || !sourceFolderUrl || isRunning}
              />

              <DefaultButton
                text={isRunning ? 'Procesando Fase 8...' : 'Fase 8: Alta diagrama nuevo en solicitud'}
                onClick={() => { void ejecutarFase8(); }}
                disabled={!excelFile || !sourceFolderUrl || isRunning}
              />
              <DefaultButton
                text={isRunning ? 'Procesando Fase 9...' : 'Fase 9: Baja diagrama de flujo en solicitud'}
                onClick={() => { void ejecutarFase9(); }}
                disabled={!excelFile || isRunning}
              />

              {false && (
                <DefaultButton
                  text={isRunning ? 'Corrigiendo 2460...' : 'Corregir ID 2460'}
                  onClick={() => { void ejecutarCorreccion2460(); }}
                  disabled={isRunning}
                />
              )}

              <Text className={styles.rollbackInfo}>
                {rollbackSummary}
              </Text>

              <DefaultButton
                text="Rollback"
                onClick={() => { void ejecutarRollback(); }}
                disabled={
                  isRunning ||
                  (
                    lastExecutedPhase === 'fase2'
                      ? !fase2RollbackEntries.length
                      : lastExecutedPhase === 'fase4'
                        ? !fase4RollbackEntries.length
                      : lastExecutedPhase === 'fase3'
                        ? !fase3RollbackEntries.length
                        : lastExecutedPhase === 'fase6'
                          ? (!fase6RollbackEntries.length && !fase6DiagramRollbackEntries.length)
                          : lastExecutedPhase === 'fase7'
                            ? !fase7RollbackEntries.length
                            : lastExecutedPhase === 'fase8'
                              ? !fase8RollbackEntries.length
                              : lastExecutedPhase === 'fase9'
                                ? !fase9RollbackEntries.length
                                : lastExecutedPhase === 'correccion2460'
                                  ? !fase8RollbackEntries.length
                                  : (!createdSolicitudIds.length && !oldSolicitudIds.length && !tempFileUrls.length)
                  )
                }
              />
            </Stack>
          </div>
        </div>

        <div className={styles.logCard}>
          <Stack tokens={{ childrenGap: 14 }}>
            <div className={styles.logHeader}>
              <div>
                <Text variant="large" className={styles.logTitle}>Resultado de revision</Text>
                <Text className={styles.logDescription}>
                  Se mantendra este panel para mostrar el detalle del proceso, igual que en migracion.
                </Text>
              </div>
              <DefaultButton text="Copiar log" onClick={copiarLog} />
            </div>

            <TextField
              multiline
              readOnly
              value={logRevision}
              className={styles.logField}
              styles={{
                field: {
                  minHeight: 260,
                  whiteSpace: 'pre-wrap',
                  fontFamily: 'Consolas, Monaco, monospace'
                }
              }}
            />
          </Stack>
        </div>
      </Stack>
    </section>
  );
};

export default Modificacion;
