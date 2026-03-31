/* eslint-disable */
import * as XLSX from 'xlsx';

export function getExcelPlaceholdersFromArrayBuffer(buf: ArrayBuffer): string[] {
  const workbook = XLSX.read(buf, { type: 'array' });
  const tags: string[] = [];
  const regex = /\{\s*([A-Za-z0-9_.\-]+)\s*\}/g;

  for (let i = 0; i < workbook.SheetNames.length; i++) {
    const sheet = workbook.Sheets[workbook.SheetNames[i]];
    const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) as any[][];

    for (let r = 0; r < grid.length; r++) {
      for (let c = 0; c < grid[r].length; c++) {
        const cell = grid[r][c];
        if (typeof cell !== 'string') {
          continue;
        }

        let match: RegExpExecArray | null;
        while ((match = regex.exec(cell)) !== null) {
          if (match[1]) {
            tags.push(match[1].trim());
          }
        }
      }
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
