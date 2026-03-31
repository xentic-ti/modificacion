/* eslint-disable */
// @ts-nocheck
import * as JSZip from 'jszip';

function getAttrVal(node: string): string | undefined {
  const match = node.match(/w:val="([^"]+)"/i) || node.match(/val="([^"]+)"/i);
  return match ? match[1] : undefined;
}

function extractTagsFromXml(xml: string): string[] {
  const tags: string[] = [];
  const tagNodes = xml.match(/<w:tag\b[^>]*\/?>/gi) || [];

  for (let i = 0; i < tagNodes.length; i++) {
    const value = getAttrVal(tagNodes[i]);
    if (value) {
      tags.push(String(value).trim());
    }
  }

  const seen = new Set<string>();
  const result: string[] = [];

  for (let i = 0; i < tags.length; i++) {
    const tag = tags[i];
    const key = tag.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      result.push(tag);
    }
  }

  return result;
}

export async function getWordTagsFromArrayBuffer(buf: ArrayBuffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(buf);
  const parts = Object.keys(zip.files).filter((path) =>
    path === 'word/document.xml' ||
    /^word\/header\d+\.xml$/i.test(path) ||
    /^word\/footer\d+\.xml$/i.test(path)
  );

  const tags: string[] = [];
  for (let i = 0; i < parts.length; i++) {
    const xml = await zip.files[parts[i]].async('text');
    tags.push(...extractTagsFromXml(xml));
  }

  const seen = new Set<string>();
  const result: string[] = [];
  for (let i = 0; i < tags.length; i++) {
    const tag = tags[i];
    const key = tag.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      result.push(tag);
    }
  }

  return result;
}
