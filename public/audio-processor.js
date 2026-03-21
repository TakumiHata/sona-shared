// AudioWorklet プロセッサ: Float32 PCM を Int16 PCM に変換して送信する
class PcmProcessor extends AudioWorkletProcessor {
  process(inputs) {
    const input = inputs[0];
    if (input && input[0]) {
      const float32 = input[0];
      const int16 = new Int16Array(float32.length);
      for (let i = 0; i < float32.length; i++) {
        // Float32 (-1.0 〜 1.0) を Int16 (-32768 〜 32767) にクランプして変換
        int16[i] = Math.max(-32768, Math.min(32767, float32[i] * 32768));
      }
      // バッファを転送（ゼロコピー）
      this.port.postMessage(int16.buffer, [int16.buffer]);
    }
    return true;
  }
}

registerProcessor('pcm-processor', PcmProcessor);
