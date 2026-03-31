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

function replaceTextNodesInSpreadsheetXml(xml: string, map: Record<string, string>): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');

  if (doc.getElementsByTagName('parsererror').length > 0) {
    return xml;
  }

  const tNodes = Array.from(doc.getElementsByTagNameNS(XLS_NS, 't'));
  let changed = false;

  for (let i = 0; i < tNodes.length; i++) {
    const node = tNodes[i];
    const original = node.textContent || '';
    let next = original;

    Object.keys(map || {}).forEach((key) => {
      if (key && next.indexOf(key) !== -1) {
        next = next.split(key).join(map[key] || '');
      }
    });

    if (next !== original) {
      node.textContent = next;
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
