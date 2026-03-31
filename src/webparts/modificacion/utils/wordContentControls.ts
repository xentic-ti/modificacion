/* eslint-disable */
// @ts-nocheck
import * as PizZip from 'pizzip';

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const HYPERLINK_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

function rowHasTag(row: Element, tag: string): boolean {
  const sdts = Array.from(row.getElementsByTagNameNS(WORD_NS, 'sdt'));
  for (let i = 0; i < sdts.length; i++) {
    const sdtPr = sdts[i].getElementsByTagNameNS(WORD_NS, 'sdtPr')[0];
    if (!sdtPr) continue;
    const tagEl = sdtPr.getElementsByTagNameNS(WORD_NS, 'tag')[0];
    const tagVal =
      tagEl?.getAttributeNS(WORD_NS, 'val') ??
      tagEl?.getAttribute('w:val') ??
      tagEl?.getAttribute('val');
    if (tagVal === tag) return true;
  }
  return false;
}

function findTableWithRevisorTags(doc: Document): Element | null {
  const allTables = doc.getElementsByTagNameNS(WORD_NS, 'tbl');
  for (let i = 0; i < allTables.length; i++) {
    const table = allTables[i];
    const rows = Array.from(table.getElementsByTagNameNS(WORD_NS, 'tr'));
    const hasReviewerRow = rows.some(
      (row) => rowHasTag(row, 'PuestoRevisor') || rowHasTag(row, 'NombreRevisor')
    );
    if (hasReviewerRow) return table;
  }
  return null;
}

function ajustarFilasRevisoresPorTags(zip: any, cantidad: number): void {
  if (cantidad <= 1) return;

  const path = 'word/document.xml';
  const xml = zip.file(path)?.asText();
  if (!xml) return;

  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) return;

  const table = findTableWithRevisorTags(doc);
  if (!table) return;

  const rows = Array.from(table.getElementsByTagNameNS(WORD_NS, 'tr'));
  const reviewerRows = rows.filter(
    (row) => rowHasTag(row, 'PuestoRevisor') || rowHasTag(row, 'NombreRevisor')
  );

  if (!reviewerRows.length) return;

  const templateRow = reviewerRows[reviewerRows.length - 1];
  for (let i = 0; i < reviewerRows.length - 1; i++) {
    table.removeChild(reviewerRows[i]);
  }

  const insertBeforeNode = templateRow.nextSibling;
  for (let i = 0; i < cantidad - 1; i++) {
    const clone = templateRow.cloneNode(true) as Element;
    table.insertBefore(clone, insertBeforeNode);
  }

  zip.file(path, new XMLSerializer().serializeToString(doc));
}

function obtenerSiguienteRId(zip: any): number {
  const relsPath = 'word/_rels/document.xml.rels';
  const relsXml = zip.file(relsPath)?.asText();
  if (!relsXml) return 100;

  const regex = /<Relationship[^>]+Id="rId(\d+)"/g;
  const ids: number[] = [];
  let match: RegExpExecArray | null;

  while ((match = regex.exec(relsXml)) !== null) {
    ids.push(parseInt(match[1], 10));
  }

  return ids.length ? Math.max(...ids) + 1 : 100;
}

function agregarRelacionesHipervinculos(
  zip: any,
  relaciones: Array<{ id: string; url: string; }>
): void {
  if (!relaciones.length) return;

  const relsPath = 'word/_rels/document.xml.rels';
  const relsXml = zip.file(relsPath)?.asText();
  if (!relsXml) throw new Error('No se pudo leer document.xml.rels');

  const parser = new DOMParser();
  const doc = parser.parseFromString(relsXml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) {
    throw new Error('Error parseando document.xml.rels');
  }

  const relationships = doc.getElementsByTagNameNS(RELS_NS, 'Relationships')[0];
  if (!relationships) {
    throw new Error('No se encontró Relationships en .rels');
  }

  for (let i = 0; i < relaciones.length; i++) {
    const relation = relaciones[i];
    const rel = doc.createElementNS(RELS_NS, 'Relationship');
    rel.setAttribute('Id', relation.id);
    rel.setAttribute('Type', HYPERLINK_TYPE);
    rel.setAttribute('Target', relation.url);
    rel.setAttribute('TargetMode', 'External');
    relationships.appendChild(rel);
  }

  zip.file(relsPath, new XMLSerializer().serializeToString(doc));
}

function limpiarYLlenarCeldaConHipervinculo(
  cell: Element,
  textToShow: string,
  doc: Document,
  relationId: string
): void {
  const paragraphs = Array.from(cell.childNodes).filter(
    (node): node is Element => node.nodeType === 1 && (node as Element).nodeName === 'w:p'
  );

  for (let i = 1; i < paragraphs.length; i++) {
    cell.removeChild(paragraphs[i]);
  }

  const paragraph = paragraphs[0] ?? doc.createElementNS(WORD_NS, 'w:p');
  if (!paragraphs[0]) {
    cell.appendChild(paragraph);
  }

  const runs = Array.from(paragraph.childNodes).filter(
    (node): node is Element => node.nodeType === 1 && (node as Element).nodeName === 'w:r'
  );
  const firstRun = runs[0] || null;
  const originalRunProps = firstRun
    ? Array.from(firstRun.childNodes).find(
        (node) => node.nodeType === 1 && (node as Element).nodeName === 'w:rPr'
      )
    : null;

  const children = Array.from(paragraph.childNodes);
  for (let i = 0; i < children.length; i++) {
    if (children[i].nodeType === 1 && (children[i] as Element).nodeName === 'w:pPr') continue;
    paragraph.removeChild(children[i]);
  }

  const hyperlink = doc.createElementNS(WORD_NS, 'w:hyperlink');
  hyperlink.setAttributeNS(
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'r:id',
    relationId
  );

  const run = doc.createElementNS(WORD_NS, 'w:r');
  const runProps = doc.createElementNS(WORD_NS, 'w:rPr');

  if (originalRunProps) {
    for (let i = 0; i < originalRunProps.childNodes.length; i++) {
      const child = originalRunProps.childNodes[i] as Element;
      const nodeName = child?.nodeName || '';
      if (nodeName === 'w:rStyle' || nodeName === 'w:color' || nodeName === 'w:u') {
        continue;
      }
      runProps.appendChild(child.cloneNode(true));
    }
  }

  const style = doc.createElementNS(WORD_NS, 'w:rStyle');
  style.setAttribute('w:val', 'Hyperlink');
  runProps.appendChild(style);

  const color = doc.createElementNS(WORD_NS, 'w:color');
  color.setAttribute('w:val', '0563C1');
  runProps.appendChild(color);

  const underline = doc.createElementNS(WORD_NS, 'w:u');
  underline.setAttribute('w:val', 'single');
  runProps.appendChild(underline);

  run.appendChild(runProps);

  const text = doc.createElementNS(WORD_NS, 'w:t');
  if (/^\s|\s$/.test(textToShow || '') || (textToShow || '').indexOf('  ') !== -1) {
    text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  }
  text.textContent = textToShow || '';
  run.appendChild(text);

  hyperlink.appendChild(run);
  paragraph.appendChild(hyperlink);
}

function limpiarYLlenarCeldaConTexto(cell: Element, textValue: string, doc: Document): void {
  const paragraphs = Array.from(cell.childNodes).filter(
    (node): node is Element => node.nodeType === 1 && (node as Element).nodeName === 'w:p'
  );

  for (let i = 1; i < paragraphs.length; i++) {
    cell.removeChild(paragraphs[i]);
  }

  const paragraph = paragraphs[0] ?? doc.createElementNS(WORD_NS, 'w:p');
  if (!paragraphs[0]) {
    cell.appendChild(paragraph);
  }

  const runs = Array.from(paragraph.childNodes).filter(
    (node): node is Element => node.nodeType === 1 && (node as Element).nodeName === 'w:r'
  );
  const firstRun = runs[0] || null;
  const originalRunProps = firstRun
    ? Array.from(firstRun.childNodes).find(
        (node) => node.nodeType === 1 && (node as Element).nodeName === 'w:rPr'
      )
    : null;

  const children = Array.from(paragraph.childNodes);
  for (let i = 0; i < children.length; i++) {
    if (children[i].nodeType === 1 && (children[i] as Element).nodeName === 'w:pPr') continue;
    paragraph.removeChild(children[i]);
  }

  const run = doc.createElementNS(WORD_NS, 'w:r');
  if (originalRunProps) {
    run.appendChild(originalRunProps.cloneNode(true));
  }

  const text = doc.createElementNS(WORD_NS, 'w:t');
  if (/^\s|\s$/.test(textValue || '') || (textValue || '').indexOf('  ') !== -1) {
    text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  }
  text.textContent = textValue || '';
  run.appendChild(text);

  paragraph.appendChild(run);
}

function buscarTablaPorContentControlTag(doc: Document, contentControlTag: string): Element | null {
  const allTables = doc.getElementsByTagNameNS(WORD_NS, 'tbl');
  for (let i = 0; i < allTables.length; i++) {
    const table = allTables[i];
    const sdts = table.getElementsByTagNameNS(WORD_NS, 'sdt');

    for (let j = 0; j < sdts.length; j++) {
      const sdtPr = sdts[j].getElementsByTagNameNS(WORD_NS, 'sdtPr')[0];
      if (!sdtPr) continue;
      const tagEl = sdtPr.getElementsByTagNameNS(WORD_NS, 'tag')[0];
      const tagVal =
        tagEl?.getAttributeNS(WORD_NS, 'val') ||
        tagEl?.getAttribute('w:val') ||
        tagEl?.getAttribute('val');

      if (tagVal === contentControlTag) return table;
    }
  }
  return null;
}

function procesarTablaDinamicaConHipervinculo(
  zip: any,
  contentControlTag: string,
  items: Array<{ codigoDocumento: string; nombreDocumento: string; enlace: string; }>,
  linkText: string
): void {
  const documentXml = zip.file('word/document.xml')?.asText();
  if (!documentXml) {
    throw new Error('No se encontró word/document.xml');
  }

  const parser = new DOMParser();
  const doc = parser.parseFromString(documentXml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) {
    throw new Error('Error parseando document.xml');
  }

  const table = buscarTablaPorContentControlTag(doc, contentControlTag);
  if (!table) return;

  const rows = Array.from(table.getElementsByTagNameNS(WORD_NS, 'tr'));
  if (rows.length < 2) return;

  const dataRows = rows.slice(1);
  for (let i = 0; i < dataRows.length - 1; i++) {
    table.removeChild(dataRows[i]);
  }

  const templateRow = dataRows[dataRows.length - 1];
  let nextRid = obtenerSiguienteRId(zip);
  const relations: Array<{ id: string; url: string; }> = [];

  items.forEach((item, index) => {
    const row = index === 0 ? templateRow : (templateRow.cloneNode(true) as Element);
    if (index !== 0) table.appendChild(row);

    const cells = Array.from(row.getElementsByTagNameNS(WORD_NS, 'tc'));
    if (cells[0]) limpiarYLlenarCeldaConTexto(cells[0], item.codigoDocumento || '', doc);
    if (cells[1]) limpiarYLlenarCeldaConTexto(cells[1], item.nombreDocumento || '', doc);

    const relationId = `rId${nextRid + index}`;
    if (cells[2] && item.enlace) {
      limpiarYLlenarCeldaConHipervinculo(cells[2], linkText, doc, relationId);
      relations.push({ id: relationId, url: item.enlace });
    } else if (cells[2]) {
      limpiarYLlenarCeldaConTexto(cells[2], '', doc);
    }
  });

  if (relations.length) {
    agregarRelacionesHipervinculos(zip, relations);
  }

  zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
}

function procesarTablaDocumentosRelacionados(
  zip: any,
  documents: Array<{ codigoDocumento: string; nombreDocumento: string; enlace: string; }>
): void {
  procesarTablaDinamicaConHipervinculo(zip, 'NombreDocumentoRelacionado', documents, 'Enlace del documento');
}

function procesarTablaFlujosProceso(
  zip: any,
  flows: Array<{ codigoDocumento: string; nombreDocumento: string; enlace: string; }>
): void {
  procesarTablaDinamicaConHipervinculo(zip, 'NombreDiagramaFlujo', flows, 'Enlace del flujo');
}

type HyperlinkValue = { url: string; text?: string; };
type ReplacementValue = string | string[] | HyperlinkValue | HyperlinkValue[];

const isHyperlinkObj = (value: any): value is HyperlinkValue =>
  !!value && typeof value === 'object' && typeof value.url === 'string';

type ZipLike = {
  files?: Record<string, any>;
  file: (path: string, content?: string) => { asText: () => string; } | null;
  generate: (opts: { type: 'blob'; mimeType: string; }) => Blob;
};

function getZipFiles(zip: any): Record<string, any> {
  return (zip && (zip.files as Record<string, any>)) || {};
}

function getXmlTargets(zip: ZipLike): string[] {
  return Object.keys(getZipFiles(zip)).filter((path) =>
    /^word\/(document|header\d+|footer\d+)\.xml$/.test(path)
  );
}

function toStr(value: any): string {
  if (value === null || value === undefined) return '';
  return String(value);
}

function getValueForOccurrence(value: ReplacementValue, occurrenceIndex: number): string | HyperlinkValue {
  if (Array.isArray(value)) {
    const picked = (value as any[])[occurrenceIndex];
    return isHyperlinkObj(picked) ? picked : toStr(picked ?? '');
  }

  return isHyperlinkObj(value) ? value : toStr(value);
}

function buildHyperlinkFieldRuns(doc: Document, url: string, textValue: string): Element[] {
  const makeElement = (name: string): Element => doc.createElementNS(WORD_NS, `w:${name}`);

  const runBegin = makeElement('r');
  const fieldBegin = makeElement('fldChar');
  fieldBegin.setAttributeNS(WORD_NS, 'w:fldCharType', 'begin');
  runBegin.appendChild(fieldBegin);

  const runInstr = makeElement('r');
  const instrText = makeElement('instrText');
  instrText.setAttribute('xml:space', 'preserve');
  instrText.textContent = ` HYPERLINK "${url}" `;
  runInstr.appendChild(instrText);

  const runSep = makeElement('r');
  const fieldSep = makeElement('fldChar');
  fieldSep.setAttributeNS(WORD_NS, 'w:fldCharType', 'separate');
  runSep.appendChild(fieldSep);

  const runText = makeElement('r');
  const runProps = makeElement('rPr');
  const runStyle = makeElement('rStyle');
  runStyle.setAttributeNS(WORD_NS, 'w:val', 'Hyperlink');
  runProps.appendChild(runStyle);
  runText.appendChild(runProps);

  const text = makeElement('t');
  text.textContent = textValue || url;
  runText.appendChild(text);

  const runEnd = makeElement('r');
  const fieldEnd = makeElement('fldChar');
  fieldEnd.setAttributeNS(WORD_NS, 'w:fldCharType', 'end');
  runEnd.appendChild(fieldEnd);

  return [runBegin, runInstr, runSep, runText, runEnd];
}

function fillContentControlsInXml(xml: string, replacements: Record<string, ReplacementValue>): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');

  if (doc.getElementsByTagName('parsererror').length > 0) {
    return xml;
  }

  const sdts = Array.from(doc.getElementsByTagNameNS(WORD_NS, 'sdt'));
  const occurrences: Record<string, number> = {};
  let touched = false;

  for (let i = 0; i < sdts.length; i++) {
    const sdt = sdts[i];
    const sdtPr = sdt.getElementsByTagNameNS(WORD_NS, 'sdtPr')[0];
    if (!sdtPr) continue;

    const tagEl = sdtPr.getElementsByTagNameNS(WORD_NS, 'tag')[0];
    const tagVal =
      tagEl?.getAttributeNS(WORD_NS, 'val') ||
      tagEl?.getAttribute('w:val') ||
      tagEl?.getAttribute('val');

    if (!tagVal || !(tagVal in replacements)) continue;

    const sdtContent = sdt.getElementsByTagNameNS(WORD_NS, 'sdtContent')[0];
    if (!sdtContent) continue;

    const textNodes = Array.from(sdtContent.getElementsByTagNameNS(WORD_NS, 't'));
    if (!textNodes.length) continue;

    const index = occurrences[tagVal] ?? 0;
    const value = getValueForOccurrence(replacements[tagVal], index);

    if (isHyperlinkObj(value)) {
      while (sdtContent.firstChild) {
        sdtContent.removeChild(sdtContent.firstChild);
      }

      const runs = buildHyperlinkFieldRuns(doc, value.url, value.text || value.url);
      for (let j = 0; j < runs.length; j++) {
        sdtContent.appendChild(runs[j]);
      }
    } else {
      textNodes[0].textContent = toStr(value);
      for (let j = 1; j < textNodes.length; j++) {
        textNodes[j].textContent = '';
      }

      if (tagVal === 'TituloDocumento') {
        const run = textNodes[0].parentNode as Element;
        if (run) {
          let runProps = run.getElementsByTagNameNS(WORD_NS, 'rPr')[0];
          if (!runProps) {
            runProps = doc.createElementNS(WORD_NS, 'w:rPr');
            run.insertBefore(runProps, run.firstChild);
          }

          let bold = runProps.getElementsByTagNameNS(WORD_NS, 'b')[0];
          if (!bold) {
            bold = doc.createElementNS(WORD_NS, 'w:b');
            runProps.appendChild(bold);
          }

          let size = runProps.getElementsByTagNameNS(WORD_NS, 'sz')[0];
          if (!size) {
            size = doc.createElementNS(WORD_NS, 'w:sz');
            runProps.appendChild(size);
          }

          size.setAttribute('w:val', '48');
        }
      }
    }

    occurrences[tagVal] = index + 1;
    touched = true;
  }

  if (!touched) return xml;
  return new XMLSerializer().serializeToString(doc);
}

function fillContentControls(zip: ZipLike, replacements: Record<string, ReplacementValue>): void {
  const targets = getXmlTargets(zip);
  for (let i = 0; i < targets.length; i++) {
    const target = targets[i];
    const xml = zip.file(target)?.asText();
    if (!xml) continue;
    const newXml = fillContentControlsInXml(xml, replacements);
    if (newXml !== xml) {
      (zip as any).file(target, newXml);
    }
  }
}

export function generateDocxWithContentControls(
  templateArrayBuffer: ArrayBuffer,
  replacements: Record<string, any>,
  dynamicData?: {
    revisores?: Array<{ puesto: string; nombre: string; }>;
    documentosRelacionados?: Array<{ codigoDocumento: string; nombreDocumento: string; enlace: string; }>;
    flujosProceso?: Array<{ codigoDocumento: string; nombreDocumento: string; enlace: string; }>;
  }
): Blob {
  const zip = new (PizZip as any)(templateArrayBuffer);

  if (dynamicData?.revisores?.length && dynamicData.revisores.length > 1) {
    ajustarFilasRevisoresPorTags(zip, dynamicData.revisores.length);
  }

  fillContentControls(zip as any, replacements as any);

  if (dynamicData?.documentosRelacionados?.length) {
    procesarTablaDocumentosRelacionados(zip, dynamicData.documentosRelacionados);
  }

  if (dynamicData?.flujosProceso?.length) {
    procesarTablaFlujosProceso(zip, dynamicData.flujosProceso);
  }

  return zip.generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });
}
