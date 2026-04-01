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

const Modificacion: React.FC<IModificacionProps> = ({ context, hasTeamsContext, isDarkTheme }) => {
  const [excelFile, setExcelFile] = React.useState<IFilePickerResult | null>(null);
  const [sourceFolderUrl, setSourceFolderUrl] = React.useState<string>('');
  const [error, setError] = React.useState<string | null>(null);
  const [isRunning, setIsRunning] = React.useState<boolean>(false);
  const [createdSolicitudIds, setCreatedSolicitudIds] = React.useState<number[]>([]);
  const [oldSolicitudIds, setOldSolicitudIds] = React.useState<number[]>([]);
  const [tempFileUrls, setTempFileUrls] = React.useState<string[]>([]);
  const [fase2RollbackEntries, setFase2RollbackEntries] = React.useState<IFase2RollbackEntry[]>([]);
  const [fase3RollbackEntries, setFase3RollbackEntries] = React.useState<IFase3RollbackEntry[]>([]);
  const [fase4RollbackEntries, setFase4RollbackEntries] = React.useState<IFase2RollbackEntry[]>([]);
  const [lastExecutedPhase, setLastExecutedPhase] = React.useState<'fase1' | 'fase2' | 'fase3' | 'fase4' | null>(null);
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
  }, [appendLog, context, createdSolicitudIds, fase2RollbackEntries, fase3RollbackEntries, fase4RollbackEntries, lastExecutedPhase, oldSolicitudIds, tempFileUrls]);

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
      setLastExecutedPhase(null);
      appendLog(`✅ Fase 6 terminada. Procesados=${resultado.processed}`);
      appendLog(`✅ OK=${resultado.ok} | SKIP=${resultado.skipped} | ERROR=${resultado.error}`);
      appendLog('ℹ️ Fase 6 no genera nuevas solicitudes ni rollback automático.');
    } catch (fase6Error) {
      const errorMessage = fase6Error instanceof Error ? fase6Error.message : String(fase6Error);
      setError(errorMessage);
      appendLog(`❌ Error en Fase 6: ${errorMessage}`);
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

    if (!createdSolicitudIds.length && !oldSolicitudIds.length && !tempFileUrls.length) {
      return 'Aun no hay resultados para revertir.';
    }

    return `ID modificados ${oldSolicitudIds.length} | ID nuevos ${createdSolicitudIds.length} | TEMP ${tempFileUrls.length}`;
  }, [createdSolicitudIds.length, fase2RollbackEntries.length, fase3RollbackEntries.length, fase4RollbackEntries.length, lastExecutedPhase, oldSolicitudIds.length, tempFileUrls.length]);

  return (
    <section className={`${styles.modificacion} ${hasTeamsContext ? styles.teams : ''} ${isDarkTheme ? styles.dark : ''}`}>
      <Stack tokens={{ childrenGap: 24 }}>
        <div className={styles.hero}>
          <div>
            <Text variant="xxLarge" className={styles.title}>Modificacion masiva</Text>
            <Text variant="large" className={styles.subtitle}>
              Prepara el Excel de modificacion y ejecuta una revision inicial antes de conectar el flujo completo.
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
                  <Text variant="large" className={styles.uploadTitle}>Excel de modificacion</Text>
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
                  </Stack>
                </div>
              </div>

              <div className={styles.statusCard}>
                <Text variant="mediumPlus" className={styles.statusTitle}>Estado actual</Text>
                <Text className={styles.statusValue}>
                  {excelFile ? excelFile.fileName : 'Aun no se ha seleccionado ningun archivo.'}
                </Text>
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
                label="Carpeta de archivos origen"
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
