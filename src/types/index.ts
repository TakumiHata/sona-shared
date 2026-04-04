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
