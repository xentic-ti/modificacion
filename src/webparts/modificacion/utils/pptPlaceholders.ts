/* eslint-disable */
// @ts-nocheck
import * as PizZip from 'pizzip';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

type ZipLike = {
  files?: Record<string, any>;
  file: (path: string) => { asText: () => string; } | null;
  generate: (opts: { type: 'blob'; mimeType: string; }) => Blob;
};

function getZipFiles(zip: any): Record<string, any> {
  return (zip && (zip.files as Record<string, any>)) || {};
}

function getPptTargets(zip: ZipLike): string[] {
  const files = Object.keys(getZipFiles(zip));
  const targets: string[] = [];

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    if (
      /^ppt\/slides\/slide\d+\.xml$/.test(file) ||
      /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(file) ||
      /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(file) ||
      /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(file)
    ) {
      targets.push(file);
    }
  }

  return targets;
}

function normalizeMap(map: Record<string, string>): Record<string, string> {
  const result: Record<string, string> = {};

  Object.keys(map || {}).forEach((key) => {
    const normalizedKey = String(key || '').trim();
    if (normalizedKey) {
      result[normalizedKey] = map[key] || '';
    }
  });

  return result;
}

function replaceInPptXml(xml: string, rawMap: Record<string, string>): string {
  const map = normalizeMap(rawMap);
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'application/xml');

  if (doc.getElementsByTagName('parsererror').length > 0) {
    return xml;
  }

  const paragraphs = Array.from(doc.getElementsByTagNameNS(A_NS, 'p'));
  let changed = false;

  for (let i = 0; i < paragraphs.length; i++) {
    const paragraph = paragraphs[i];
    const textNodes = Array.from(paragraph.getElementsByTagNameNS(A_NS, 't'));
    if (!textNodes.length) {
      continue;
    }

    const original = textNodes.map((node) => node.textContent || '').join('');
    let next = original;

    Object.keys(map).forEach((key) => {
      if (next.indexOf(key) !== -1) {
        next = next.split(key).join(map[key] || '');
      }
    });

    if (next !== original) {
      textNodes[0].textContent = next;
      for (let j = 1; j < textNodes.length; j++) {
        textNodes[j].textContent = '';
      }
      changed = true;
    }
  }

  if (!changed) {
    return xml;
  }

  return new XMLSerializer().serializeToString(doc);
}

export function applyPptReplacements(zip: ZipLike, map: Record<string, string>): void {
  const targets = getPptTargets(zip);

  for (let i = 0; i < targets.length; i++) {
    const target = targets[i];
    const xml = zip.file(target)?.asText();
    if (!xml) {
      continue;
    }

    const newXml = replaceInPptXml(xml, map);
    if (newXml !== xml) {
      (zip as any).file(target, newXml);
    }
  }
}

export function generatePptxWithPlaceholders(
  templateArrayBuffer: ArrayBuffer,
  replacements: Record<string, string>
): Blob {
  const zip = new (PizZip as any)(templateArrayBuffer) as ZipLike;
  applyPptReplacements(zip, replacements);

  return zip.generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  });
}
