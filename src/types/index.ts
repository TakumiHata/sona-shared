export interface EmbeddedImage {
    base64: string;
    width: number;
    height: number;
    source: 'mermaid' | 'markdown-image';
    alt?: string;
}

export interface EmbeddedTable {
    headers: string[];
    rows: string[][];
}

export interface EnrichedContent {
    images: EmbeddedImage[];
    tables: EmbeddedTable[];
    imageUrls: string[];
}

export interface AgendaItem {
    id: string;
    title: string;
    description?: string;
    originalDescription?: string;
    rawTranscript?: string;
    refinedTranscript?: string;
    summaryText?: string;
    durationMinutes?: number;
    speaker?: string | null;
    children?: AgendaItem[];
    lastProcessedTranscript?: string;
    isProcessing?: boolean;
    enrichedContent?: EnrichedContent;
}

export interface MeetingInfo {
    title: string;
    date: string;
    participants: string[];
    location: string;
}

export interface AgendaJSON {
    version: string;
    meetingInfo: MeetingInfo;
    agendaItems: AgendaItem[];
}

export interface TranscribeSegment {
    start: number;
    end: number;
    text: string;
    speaker: string | null;
    confidence: number;
}

export interface TranscribeResponse {
    segments: TranscribeSegment[];
    full_text: string;
    duration_seconds: number;
    processing_time_seconds: number;
}

export interface FlatAgendaWithDepth extends AgendaItem {
    depth: number;
}

// ── mapping_json スキーマ ──

export interface ColumnRegion {
    tag: string;
    col_start: string;
    col_end: string;
    label: string;
    /** テンプレート内での行オフセット（0始まり）。複数行パターン時に使用。 */
    row_offset?: number;
}

export interface PrintArea {
    data_start_row: number;
    data_end_row: number;
    repeat_header: boolean;
    footer_rows: number;
}

export interface GridDetection {
    is_hougan: boolean;
    base_cell_size: number | null;
}

export interface MappingJsonV3 {
    version: 3;
    detected_tags: string[];
    print_area: PrintArea;
    column_regions: ColumnRegion[];
    grid_detection: GridDetection;
}

export interface MappingJsonV2 {
    version: 2;
    detected_tags: string[];
    rows_per_page: number;
}

export type MappingJson = MappingJsonV2 | MappingJsonV3;

export const isMappingV3 = (json: unknown): json is MappingJsonV3 =>
    json != null && typeof json === 'object' && (json as Record<string, unknown>).version === 3;
