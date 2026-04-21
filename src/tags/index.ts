import type { AgendaItem } from '../types';

const AGENDA_RE = /<agenda>([\s\S]*?)<\/agenda>/g;
const DETAILS_RE = /<details>[\s\S]*?<summary>[\s\S]*?<\/summary>([\s\S]*?)<\/details>/g;

/** <agenda>タグを除去して中身だけ残す */
export function stripAgendaTags(text: string): string {
  return text.replace(AGENDA_RE, '$1').trim();
}

/** <details>/<summary>タグを除去して中身だけ残す */
export function stripDetailsTags(text: string): string {
  return text.replace(DETAILS_RE, '$1').trim();
}

/** Sona 固有カスタムタグ (<agenda>) が含まれているか */
export function hasCustomTags(text: string): boolean {
  return AGENDA_RE.test(text);
}

/** Excel 出力用に Agenda ツリーをサニタイズ（<agenda> / <details> のタグを除去し中身は保持） */
export function sanitizeAgendasForExport(agendas: AgendaItem[]): AgendaItem[] {
  return agendas.map(item => {
    const sanitized: AgendaItem = { ...item };

    if (sanitized.description) {
      sanitized.description = stripDetailsTags(stripAgendaTags(sanitized.description));
    }
    if (sanitized.rawTranscript) {
      sanitized.rawTranscript = stripDetailsTags(stripAgendaTags(sanitized.rawTranscript));
    }
    if (sanitized.refinedTranscript) {
      sanitized.refinedTranscript = stripDetailsTags(stripAgendaTags(sanitized.refinedTranscript));
    }
    if (sanitized.summaryText) {
      sanitized.summaryText = stripDetailsTags(stripAgendaTags(sanitized.summaryText));
    }

    if (sanitized.children) {
      sanitized.children = sanitizeAgendasForExport(sanitized.children);
    }

    return sanitized;
  });
}
