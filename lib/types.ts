// トランスクリプトの1エントリを表す型
export interface TranscriptEntry {
  id: string;
  text: string;
  speaker: string;
  isFinal: boolean;
  timestamp: string;
}

// セッション情報の型
export interface SessionInfo {
  id: string;
  name: string;
  status: string;
  description: string | null;
}
