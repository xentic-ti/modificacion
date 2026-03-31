/* eslint-disable */
// @ts-nocheck
import * as JSZip from 'jszip';

function extractPlaceholders(text: string): string[] {
  const tags: string[] = [];
  const regex = /\{([A-Za-z][A-Za-z0-9_]{1,50})\}/g;
  let match: RegExpExecArray | null;

  while ((match = regex.exec(String(text || ''))) !== null) {
    if (match[1]) {
      tags.push(match[1]);
    }
  }

  return tags;
}

export async function getPptPlaceholdersFromArrayBuffer(buf: ArrayBuffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(buf);
  const parts = Object.keys(zip.files).filter((path) =>
    /^ppt\/slides\/slide\d+\.xml$/i.test(path) || /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(path)
  );

  const tags: string[] = [];
  for (let i = 0; i < parts.length; i++) {
    const xml = await zip.files[parts[i]].async('text');
    const texts = xml.match(/<a:t[^>]*>[\s\S]*?<\/a:t>/gi) || [];

    for (let j = 0; j < texts.length; j++) {
      const inner = texts[j]
        .replace(/<a:t[^>]*>/i, '')
        .replace(/<\/a:t>/i, '');

      tags.push(...extractPlaceholders(inner));
    }
  }

  const guidRegex = /^[0-9A-F]{8}-[0-9A-F\-]{27}$/i;
  const clean = tags.filter((tag) => !guidRegex.test(tag));

  const seen = new Set<string>();
  const result: string[] = [];
  for (let i = 0; i < clean.length; i++) {
    const tag = clean[i];
    const key = tag.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      result.push(tag);
    }
  }

  return result;
}
