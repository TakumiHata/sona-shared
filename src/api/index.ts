export interface RegisterSpeakerResponse {
    embedding: number[];
}

export interface VoiceVerifierConfig {
    baseUrl: string;
}

/**
 * voice-verifier APIクライアント。
 * baseUrlを外部から注入することで、環境変数への依存を排除。
 */
export function createVoiceVerifierClient(config: VoiceVerifierConfig) {
    const { baseUrl } = config;

    return {
        /** 音声データを送信して Embedding (声紋��ータ) を取得する */
        async getEmbedding(audioBlob: Blob): Promise<number[]> {
            const formData = new FormData();
            formData.append('audio', audioBlob, 'recording.wav');

            const response = await fetch(`${baseUrl}/register_speaker`, {
                method: 'POST',
                body: formData,
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.detail || 'Failed to get embedding');
            }

            const data: RegisterSpeakerResponse = await response.json();
            return data.embedding;
        },

        /** ヘルスチェック */
        async healthCheck(): Promise<boolean> {
            try {
                const response = await fetch(`${baseUrl}/health`);
                return response.ok;
            } catch {
                return false;
            }
        },
    };
}
