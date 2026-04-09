import type { AgendaItem } from '../types';

const PREVIEW_ONLY_RE = /<preview-only>([\s\S]*?)<\/preview-only>/g;
const FIXED_RE = /<fixed>([\s\S]*?)<\/fixed>/g;
const AGENDA_RE = /<agenda>([\s\S]*?)<\/agenda>/g;
const DETAILS_RE = /<details>[\s\S]*?<summary>[\s\S]*?<\/summary>([\s\S]*?)<\/details>/g;
const MEETING_SUMMARY_RE = /<meeting-summary>([\s\S]*?)<\/meeting-summary>/g;

/** <preview-only>...</preview-only> を除去（タグと中身の両方を削除） */
export function stripPreviewOnly(text: string): string {
  return text.replace(PREVIEW_ONLY_RE, '').trim();
}

/** <fixed>...</fixed> の中身だけを抽出して配列で返す */
export function extractFixedContent(text: string): string[] {
  const results: string[] = [];
  let match;
  const re = new RegExp(FIXED_RE.source, 'g');
  while ((match = re.exec(text)) !== null) {
    results.push(match[1].trim());
  }
  return results;
}

/** <preview-only>...</preview-only> の中身だけを抽出して配列で返す */
export function extractPreviewOnlyContent(text: string): string[] {
  const results: string[] = [];
  let match;
  const re = new RegExp(PREVIEW_ONLY_RE.source, 'g');
  while ((match = re.exec(text)) !== null) {
    results.push(match[1].trim());
  }
  return results;
}

/** <fixed>タグを除去して中身だけ残す */
export function stripFixedTags(text: string): string {
  return text.replace(FIXED_RE, '$1').trim();
}

/** <agenda>タグを除去して中身だけ残す */
export function stripAgendaTags(text: string): string {
  return text.replace(AGENDA_RE, '$1').trim();
}

/** <details>/<summary>タグを除去して中身だけ残す */
export function stripDetailsTags(text: string): string {
  return text.replace(DETAILS_RE, '$1').trim();
}

/** <meeting-summary>タグと中身の両方を除去（インポート時に読み飛ばす） */
export function stripMeetingSummaryTags(text: string): string {
  return text.replace(MEETING_SUMMARY_RE, '').trim();
}

/** カスタムタグが含まれているか */
export function hasCustomTags(text: string): boolean {
  return PREVIEW_ONLY_RE.test(text) || FIXED_RE.test(text) || AGENDA_RE.test(text) || MEETING_SUMMARY_RE.test(text);
}

/** Excel出力用にAgendaツリーをサニタイズ（preview-onlyを除去、fixedタグのみ除去して中身は保持） */
export function sanitizeAgendasForExport(agendas: AgendaItem[]): AgendaItem[] {
  return agendas.map(item => {
    const sanitized: AgendaItem = { ...item };

    if (sanitized.description) {
      sanitized.description = stripDetailsTags(stripAgendaTags(stripFixedTags(stripPreviewOnly(sanitized.description))));
    }
    if (sanitized.rawTranscript) {
      sanitized.rawTranscript = stripDetailsTags(stripAgendaTags(stripFixedTags(stripPreviewOnly(sanitized.rawTranscript))));
    }
    if (sanitized.refinedTranscript) {
      sanitized.refinedTranscript = stripDetailsTags(stripAgendaTags(stripFixedTags(stripPreviewOnly(sanitized.refinedTranscript))));
    }
    if (sanitized.summaryText) {
      sanitized.summaryText = stripDetailsTags(stripAgendaTags(stripFixedTags(stripPreviewOnly(sanitized.summaryText))));
    }

    if (sanitized.children) {
      sanitized.children = sanitizeAgendasForExport(sanitized.children);
    }

    return sanitized;
  });
}
