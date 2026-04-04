type ImageExt = 'png' | 'jpeg' | 'gif';

/** 外部URLから画像をダウンロードしてBufferとして返す */
export async function fetchImageAsBuffer(
  url: string,
  timeoutMs = 5000
): Promise<{ buffer: Buffer; ext: ImageExt } | null> {
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);

    const resp = await fetch(url, { signal: controller.signal });
    clearTimeout(timer);

    if (!resp.ok) return null;

    const contentType = resp.headers.get('content-type') || '';
    let ext: ImageExt = 'png';
    if (contentType.includes('jpeg') || contentType.includes('jpg')) ext = 'jpeg';
    else if (contentType.includes('gif')) ext = 'gif';

    const arrayBuffer = await resp.arrayBuffer();
    return { buffer: Buffer.from(arrayBuffer), ext };
  } catch (e) {
    console.warn(`[imageUtils] Failed to fetch image: ${url}`, e);
    return null;
  }
}
