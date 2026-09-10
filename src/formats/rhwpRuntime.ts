import init, { HwpDocument } from '@rhwp/core';

export interface RhwpPageDef {
  width: number;
  height: number;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  marginHeader: number;
  marginFooter: number;
  marginGutter: number;
  landscape: boolean;
  binding: number;
}

export interface RhwpImportOptions {
  pageDef?: Partial<RhwpPageDef>;
  headerText?: string;
  footerText?: string;
}

let initialization: Promise<void> | null = null;

function installTextMeasurement() {
  if ('measureTextWidth' in globalThis) return;

  let context: CanvasRenderingContext2D | null = null;
  let previousFont = '';
  Object.assign(globalThis, {
    measureTextWidth(font: string, text: string) {
      if (!context) context = document.createElement('canvas').getContext('2d');
      if (!context) return text.length * 8;
      if (previousFont !== font) {
        context.font = font;
        previousFont = font;
      }
      return context.measureText(text).width;
    },
  });
}

export async function ensureRhwp(): Promise<void> {
  if (!initialization) {
    installTextMeasurement();
    initialization = init({ module_or_path: '/vendor/rhwp/rhwp_bg.wasm' }).then(() => undefined);
  }
  await initialization;
}

export async function openRhwp(bytes: Uint8Array): Promise<HwpDocument> {
  await ensureRhwp();
  return new HwpDocument(bytes);
}

export async function createRhwpFromHtml(html: string, options: RhwpImportOptions = {}): Promise<HwpDocument> {
  await ensureRhwp();
  const document = HwpDocument.createEmpty();
  let paraIdx = 0;
  let charOffset = 0;

  try {
    // rHWP 0.8.6 can panic when a table follows multiple block elements in one
    // pasteHtml call. Insert Mammoth's top-level semantic blocks sequentially.
    for (const block of splitTopLevelHtmlBlocks(html)) {
      const result = JSON.parse(document.pasteHtml(0, paraIdx, charOffset, block)) as {
        ok?: boolean;
        error?: string;
        paraIdx?: number;
        charOffset?: number;
      };
      if (result.ok === false) {
        throw new Error(result.error ?? 'DOCX 내용을 OmniDocs 문서 모델로 변환하지 못했습니다.');
      }
      paraIdx = result.paraIdx ?? paraIdx;
      charOffset = result.charOffset ?? charOffset;
    }
    applyImportLayout(document, options);
    return document;
  } catch (error) {
    // A Rust panic can poison the borrowed WASM value; free only for ordinary JS errors.
    if (!(error instanceof WebAssembly.RuntimeError)) document.free();
    throw error;
  }
}

export function applyImportLayout(document: HwpDocument, options: RhwpImportOptions): void {
  if (options.pageDef && Object.keys(options.pageDef).length > 0) {
    const current = JSON.parse(document.getPageDef(0)) as RhwpPageDef;
    const next = { ...current, ...options.pageDef };
    const result = JSON.parse(document.setPageDef(0, JSON.stringify(next))) as { ok?: boolean; error?: string };
    if (result.ok === false) throw new Error(result.error ?? 'DOCX 페이지 설정을 OMDX에 적용하지 못했습니다.');
  }

  if (options.headerText) {
    document.createHeaderFooter(0, true, 0);
    document.insertTextInHeaderFooter(0, true, 0, 0, 0, options.headerText);
  }
  if (options.footerText) {
    document.createHeaderFooter(0, false, 0);
    document.insertTextInHeaderFooter(0, false, 0, 0, 0, options.footerText);
  }
}

export function splitTopLevelHtmlBlocks(html: string): string[] {
  if (!html.trim()) return [];
  if (typeof DOMParser === 'undefined') return [html];

  const parsed = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  const body = parsed.body;
  const blocks = Array.from(body.childNodes)
    .map((node) => {
      if (node.nodeType === Node.ELEMENT_NODE) return (node as Element).outerHTML;
      const text = node.textContent?.trim();
      return text ? `<p>${escapeHtml(text)}</p>` : '';
    })
    .filter(Boolean);
  return blocks.length ? blocks : [html];
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
