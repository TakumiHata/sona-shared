/**
 * AnalyserNodeからボリューム平均値を算出する。
 * Web Audio APIを利用した録音UIのボリュームメーター表示用。
 */
export function analyzeVolume(analyser: AnalyserNode): number {
    const dataArray = new Uint8Array(analyser.frequencyBinCount);
    analyser.getByteFrequencyData(dataArray);
    return dataArray.reduce((a, b) => a + b) / dataArray.length;
}

/**
 * Web Audio APIを使って録音用のAudioContextとAnalyserNodeをセットアップする。
 */
export function createAudioAnalyser(stream: MediaStream): {
    audioContext: AudioContext;
    analyser: AnalyserNode;
} {
    const audioContext = new AudioContext();
    const source = audioContext.createMediaStreamSource(stream);
    const analyser = audioContext.createAnalyser();
    analyser.fftSize = 256;
    source.connect(analyser);
    return { audioContext, analyser };
}
