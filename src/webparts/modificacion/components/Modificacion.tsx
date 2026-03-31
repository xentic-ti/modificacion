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

const Modificacion: React.FC<IModificacionProps> = ({ context, hasTeamsContext, isDarkTheme }) => {
  const [excelFile, setExcelFile] = React.useState<IFilePickerResult | null>(null);
  const [sourceFolderUrl, setSourceFolderUrl] = React.useState<string>('');
  const [error, setError] = React.useState<string | null>(null);
  const [isRunning, setIsRunning] = React.useState<boolean>(false);
  const [createdSolicitudIds, setCreatedSolicitudIds] = React.useState<number[]>([]);
  const [oldSolicitudIds, setOldSolicitudIds] = React.useState<number[]>([]);
  const [tempFileUrls, setTempFileUrls] = React.useState<string[]>([]);
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
      appendLog('✅ Rollback de Fase 1 finalizado.');
    } catch (rollbackError) {
      const errorMessage = rollbackError instanceof Error ? rollbackError.message : String(rollbackError);
      setError(errorMessage);
      appendLog(`❌ Error durante el rollback: ${errorMessage}`);
    } finally {
      setIsRunning(false);
    }
  }, [appendLog, context, createdSolicitudIds, oldSolicitudIds, tempFileUrls]);

  const copiarLog = React.useCallback(() => {
    void navigator.clipboard.writeText(logRevision);
  }, [logRevision]);

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
                text="Rollback"
                onClick={() => { void ejecutarRollback(); }}
                disabled={isRunning || (!createdSolicitudIds.length && !oldSolicitudIds.length && !tempFileUrls.length)}
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
