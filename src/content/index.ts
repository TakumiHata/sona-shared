import type { EmbeddedTable } from '../types';

const MERMAID_BLOCK_RE = /```mermaid\n([\s\S]*?)```/g;
const IMAGE_RE = /!\[([^\]]*)\]\(([^)]+)\)/g;

/** Markdownからmermaidコードブロックを抽出 */
export function extractMermaidBlocks(markdown: string): string[] {
  const results: string[] = [];
  let match;
  const re = new RegExp(MERMAID_BLOCK_RE.source, 'g');
  while ((match = re.exec(markdown)) !== null) {
    const code = match[1].trim();
    if (code) results.push(code);
  }
  return results;
}

/** Markdownから画像URLを抽出（data:URIは除外） */
export function extractImageUrls(markdown: string): string[] {
  const results: string[] = [];
  let match;
  const re = new RegExp(IMAGE_RE.source, 'g');
  while ((match = re.exec(markdown)) !== null) {
    const url = match[2];
    if (!url.startsWith('data:')) {
      results.push(url);
    }
  }
  return results;
}

/** Markdownからテーブルを抽出 */
export function extractMarkdownTables(markdown: string): EmbeddedTable[] {
  const tables: EmbeddedTable[] = [];
  const lines = markdown.split('\n');
  let i = 0;

  while (i < lines.length) {
    const line = lines[i].trim();
    if (line.startsWith('|') && line.endsWith('|')) {
      const headerLine = line;
      const separatorLine = (i + 1 < lines.length) ? lines[i + 1].trim() : '';

      if (separatorLine.startsWith('|') && /^[\s|:-]+$/.test(separatorLine)) {
        const headers = headerLine.split('|').slice(1, -1).map(h => h.trim());
        const rows: string[][] = [];

        i += 2;
        while (i < lines.length) {
          const rowLine = lines[i].trim();
          if (!rowLine.startsWith('|') || !rowLine.endsWith('|')) break;
          rows.push(rowLine.split('|').slice(1, -1).map(c => c.trim()));
          i++;
        }

        if (headers.length > 0) {
          tables.push({ headers, rows });
        }
        continue;
      }
    }
    i++;
  }

  return tables;
}

/** 全テキストフィールドからコンテンツを抽出するヘルパー */
export function extractAllFromMarkdown(markdown: string) {
  return {
    mermaidBlocks: extractMermaidBlocks(markdown),
    imageUrls: extractImageUrls(markdown),
    tables: extractMarkdownTables(markdown),
  };
}
