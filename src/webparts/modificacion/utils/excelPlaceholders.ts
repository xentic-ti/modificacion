/* eslint-disable */
// @ts-nocheck
import * as PizZip from 'pizzip';

const XLS_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

type ZipLike = {
  files?: Record<string, any>;
  file: (path: string) => { asText: () => string; } | null;
  generate: (opts: { type: 'blob'; mimeType: string; }) => Blob;
};

function getZipFiles(zip: any): Record<string, any> {
  return (zip && (zip.files as Record<string, any>)) || {};
}

function getExcelTargets(zip: ZipLike): string[] {
  const files = Object.keys(getZipFiles(zip));
  const targets: string[] = [];

  if (files.includes('xl/sharedStrings.xml')) {
    targets.push('xl/sharedStrings.xml');
  }

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    if (/^xl\/worksheets\/sheet\d+\.xml$/.test(file)) {
      targets.push(file);
    }
  }

  return targets;
}

function applyReplacementMap(value: string, map: Record<string, string>): string {
  let next = value;

  Object.keys(map || {}).forEach((key) => {
    if (key && next.indexOf(key) !== -1) {
      next = next.split(key).join(map[key] || '');
    }
  });

  return next;
}

function mustPreserveSpace(value: string): boolean {
  return /^\s|\s$/.test(value) || value.indexOf('\n') !== -1 || value.indexOf('\r') !== -1 || /\s{2,}/.test(value);
}

function writePlainTextElement(doc: XMLDocument, parent: Element, value: string): void {
  while (parent.firstChild) {
    parent.removeChild(parent.firstChild);
  }

  const tNode = doc.createElementNS(XLS_NS, 't');
  if (mustPreserveSpace(value)) {
    tNode.setAttribute('xml:space', 'preserve');
  }
  tNode.textContent = value;
  parent.appendChild(tNode);
}

function replaceGroupedTextNodes(
  doc: XMLDocument,
  parentTagName: 'si' | 'is',
  map: Record<string, string>
): boolean {
  const parents = Array.from(doc.getElementsByTagNameNS(XLS_NS, parentTagName));
  let changed = false;

  for (let i = 0; i < parents.length; i++) {
    const parent = parents[i];
    const tNodes = Array.from(parent.getElementsByTagNameNS(XLS_NS, 't'));
    if (!tNodes.length) {
      continue;
    }

    const original = tNodes.map((node) => node.textContent || '').join('');
    const next = applyReplacementMap(original, map);
    if (next === original) {
      continue;
    }

    writePlainTextElement(doc, parent, next);
    changed = true;
  }

  return changed;
}

function replaceTextNodesInSpreadsheetXml(xml: string, map: Record<string, string>): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');

  if (doc.getElementsByTagName('parsererror').length > 0) {
    return xml;
  }

  let changed = false;
  changed = replaceGroupedTextNodes(doc, 'si', map) || changed;
  changed = replaceGroupedTextNodes(doc, 'is', map) || changed;

  const tNodes = Array.from(doc.getElementsByTagNameNS(XLS_NS, 't'));

  for (let i = 0; i < tNodes.length; i++) {
    const node = tNodes[i];
    const original = node.textContent || '';
    const next = applyReplacementMap(original, map);

    if (next !== original) {
      node.textContent = next;
      if (mustPreserveSpace(next)) {
        node.setAttribute('xml:space', 'preserve');
      } else {
        node.removeAttribute('xml:space');
      }
      changed = true;
    }
  }

  if (!changed) {
    return xml;
  }

  return new XMLSerializer().serializeToString(doc);
}

export function applyExcelReplacements(zip: ZipLike, map: Record<string, string>): void {
  const targets = getExcelTargets(zip);

  for (let i = 0; i < targets.length; i++) {
    const target = targets[i];
    const xml = zip.file(target)?.asText();
    if (!xml) {
      continue;
    }

    const newXml = replaceTextNodesInSpreadsheetXml(xml, map);
    if (newXml !== xml) {
      (zip as any).file(target, newXml);
    }
  }
}

export function generateXlsxWithPlaceholders(
  templateArrayBuffer: ArrayBuffer,
  replacements: Record<string, string>
): Blob {
  const zip = new (PizZip as any)(templateArrayBuffer) as ZipLike;
  applyExcelReplacements(zip, replacements);

  return zip.generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
}
