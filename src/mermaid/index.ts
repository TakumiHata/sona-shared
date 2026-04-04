import type { EmbeddedImage } from '../types/index.js';

let renderCounter = 0;

/**
 * MermaidコードをSVG文字列に変換する。
 * mermaidライブラリはpeerDependencyとして外部から注入する。
 */
export async function renderMermaidToSvg(
  mermaidInstance: { render: (id: string, code: string) => Promise<{ svg: string }> },
  code: string
): Promise<string> {
  try {
    const id = `mermaid-${renderCounter++}`;
    const { svg } = await mermaidInstance.render(id, code);
    return svg;
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    return `<pre style="color:#ff6b6b;font-size:12px;">Mermaid エラー: ${msg}</pre>`;
  }
}

const MAX_WIDTH = 800;
const MAX_HEIGHT = 600;
const SCALE = 2;

/**
 * MermaidコードをPNG（base64）に変換する。
 * ブラウザ環境でのみ動作（DOMParser, Canvas, Image を使用）。
 */
export async function renderMermaidToPng(
  mermaidInstance: { render: (id: string, code: string) => Promise<{ svg: string }> },
  code: string
): Promise<EmbeddedImage | null> {
  try {
    const svg = await renderMermaidToSvg(mermaidInstance, code);
    if (svg.includes('Mermaid エラー')) return null;

    const parser = new DOMParser();
    const doc = parser.parseFromString(svg, 'image/svg+xml');
    const svgEl = doc.querySelector('svg');
    if (!svgEl) return null;

    let width = parseFloat(svgEl.getAttribute('width') || '400');
    let height = parseFloat(svgEl.getAttribute('height') || '300');

    if (width > MAX_WIDTH) {
      height = height * (MAX_WIDTH / width);
      width = MAX_WIDTH;
    }
    if (height > MAX_HEIGHT) {
      width = width * (MAX_HEIGHT / height);
      height = MAX_HEIGHT;
    }

    const svgBase64 = btoa(unescape(encodeURIComponent(svg)));
    const svgDataUrl = `data:image/svg+xml;base64,${svgBase64}`;

    const img = new Image();
    const loadPromise = new Promise<void>((resolve, reject) => {
      img.onload = () => resolve();
      img.onerror = () => reject(new Error('SVG image load failed'));
    });
    img.src = svgDataUrl;
    await loadPromise;

    const canvas = document.createElement('canvas');
    canvas.width = Math.round(width * SCALE);
    canvas.height = Math.round(height * SCALE);
    const ctx = canvas.getContext('2d');
    if (!ctx) return null;

    ctx.scale(SCALE, SCALE);
    ctx.drawImage(img, 0, 0, width, height);

    const pngDataUrl = canvas.toDataURL('image/png');
    const base64 = pngDataUrl.replace(/^data:image\/png;base64,/, '');

    return {
      base64,
      width: Math.round(width),
      height: Math.round(height),
      source: 'mermaid',
    };
  } catch (e) {
    console.error('[mermaidToPng] Failed:', e);
    return null;
  }
}
