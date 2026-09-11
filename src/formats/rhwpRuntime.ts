import init, { HwpDocument } from '@rhwp/core';
import {
  decodeOmdxFormat,
  OMDX_CHAR_FORMAT_ATTR,
  OMDX_LIST_KIND_ATTR,
  OMDX_LIST_LEVEL_ATTR,
  OMDX_PARA_FORMAT_ATTR,
  OMDX_SECTION_BREAK_ATTR,
  type RhwpCharacterFormat,
  type RhwpParagraphFormat,
} from './docxOoxml';

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
  headers?: {
    default?: string;
    first?: string;
    even?: string;
  };
  footers?: {
    default?: string;
    first?: string;
    even?: string;
  };
}

type HtmlPasteDocument = Pick<
  HwpDocument,
  'pasteHtml' | 'insertParagraph' | 'applyParaFormat' | 'applyCharFormat' | 'findOrCreateFontId' | 'breakAtCursor'
  | 'ensureDefaultNumbering' | 'ensureDefaultBullet'
>;

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

export async function createBlankHwpx(): Promise<Uint8Array> {
  await ensureRhwp();
  const document = HwpDocument.createEmpty();
  try {
    return document.exportHwpx();
  } finally {
    document.free();
  }
}

export async function createRhwpFromHtml(html: string, options: RhwpImportOptions = {}): Promise<HwpDocument> {
  await ensureRhwp();
  const document = HwpDocument.createEmpty();

  try {
    pasteTopLevelHtmlBlocks(document, html);
    applyImportLayout(document, options);
    return document;
  } catch (error) {
    // A Rust panic can poison the borrowed WASM value; free only for ordinary JS errors.
    if (!(error instanceof WebAssembly.RuntimeError)) document.free();
    throw error;
  }
}

/**
 * Inserts Mammoth HTML block-by-block without collapsing adjacent Word
 * paragraphs. DOCX-only layout metadata attached by docxOoxml is applied with
 * rHWP's native formatting APIs immediately after each block is inserted.
 */
export function pasteTopLevelHtmlBlocks(document: HtmlPasteDocument, html: string): void {
  let paraIdx = 0;
  let charOffset = 0;
  const blocks = splitTopLevelHtmlBlocks(html);

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    const formatting = readBlockFormatting(block);
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
    applyNativeBlockFormatting(document, paraIdx, formatting);

    if (index < blocks.length - 1) {
      const nextParaIdx = paraIdx + 1;
      const inserted = JSON.parse(document.insertParagraph(0, nextParaIdx)) as {
        ok?: boolean;
        error?: string;
        paraIdx?: number;
      };
      if (inserted.ok === false) {
        throw new Error(inserted.error ?? 'DOCX 문단 경계를 OMDX에 생성하지 못했습니다.');
      }
      paraIdx = inserted.paraIdx ?? nextParaIdx;
      charOffset = 0;

      if (formatting.sectionBreakAfter) {
        const broken = JSON.parse(document.breakAtCursor(0, paraIdx, 0, 'section')) as {
          ok?: boolean;
          reason?: string;
          para?: number;
          pos?: number;
        };
        if (broken.ok === false) {
          throw new Error(broken.reason ?? 'DOCX 구역 나누기를 OMDX에 적용하지 못했습니다.');
        }
        paraIdx = broken.para ?? paraIdx;
        charOffset = broken.pos ?? 0;
      }
    }
  }
}

interface BlockFormatting {
  paragraph?: RhwpParagraphFormat;
  characterRanges: Array<{ start: number; end: number; format: RhwpCharacterFormat }>;
  sectionBreakAfter: boolean;
  list?: { kind: 'number' | 'bullet'; level: number };
}

function readBlockFormatting(block: string): BlockFormatting {
  const empty: BlockFormatting = { characterRanges: [], sectionBreakAfter: false };
  if (typeof DOMParser === 'undefined') return empty;
  const parsed = new DOMParser().parseFromString(`<body>${block}</body>`, 'text/html');
  const root = parsed.body.firstElementChild as HTMLElement | null;
  if (!root) return empty;

  const sectionBreakAfter = root.hasAttribute(OMDX_SECTION_BREAK_ATTR);
  const isSingleParagraph = /^(P|H[1-6])$/.test(root.tagName);
  if (!isSingleParagraph) return { characterRanges: [], sectionBreakAfter };

  const paragraph = decodeOmdxFormat<RhwpParagraphFormat>(root.getAttribute(OMDX_PARA_FORMAT_ATTR));
  const listKind = root.getAttribute(OMDX_LIST_KIND_ATTR);
  const listLevel = Number(root.getAttribute(OMDX_LIST_LEVEL_ATTR) ?? '0');
  const list: BlockFormatting['list'] = listKind === 'number' || listKind === 'bullet'
    ? { kind: listKind, level: Number.isFinite(listLevel) ? Math.max(0, Math.min(6, Math.trunc(listLevel))) : 0 }
    : undefined;
  const characterRanges: BlockFormatting['characterRanges'] = [];
  for (const element of Array.from(root.querySelectorAll(`[${OMDX_CHAR_FORMAT_ATTR}]`))) {
    const span = element as HTMLElement;
    const format = decodeOmdxFormat<RhwpCharacterFormat>(span.getAttribute(OMDX_CHAR_FORMAT_ATTR));
    if (!format || !span.textContent) continue;
    const range = parsed.createRange();
    range.selectNodeContents(root);
    range.setEndBefore(span);
    const start = range.toString().length;
    const end = start + span.textContent.length;
    if (end > start) characterRanges.push({ start, end, format });
  }

  return { paragraph, characterRanges, sectionBreakAfter, list };
}

function applyNativeBlockFormatting(
  document: HtmlPasteDocument,
  paraIdx: number,
  formatting: BlockFormatting,
) {
  const paragraphProps: RhwpParagraphFormat & { headType?: string; paraLevel?: number; numberingId?: number } = {
    ...(formatting.paragraph ?? {}),
  };
  if (formatting.list) {
    paragraphProps.headType = formatting.list.kind === 'number' ? 'Number' : 'Bullet';
    paragraphProps.paraLevel = formatting.list.level;
    paragraphProps.numberingId = formatting.list.kind === 'number'
      ? document.ensureDefaultNumbering()
      : document.ensureDefaultBullet(['●', '○', '■', '□', '◆', '◇', '•'][formatting.list.level] ?? '•');
  }

  if (Object.keys(paragraphProps).length > 0) {
    const result = JSON.parse(document.applyParaFormat(0, paraIdx, JSON.stringify(paragraphProps))) as {
      ok?: boolean;
      error?: string;
    };
    if (result.ok === false) throw new Error(result.error ?? 'DOCX 문단 서식을 OMDX에 적용하지 못했습니다.');
  }

  for (const range of formatting.characterRanges) {
    const props: Record<string, unknown> = { ...range.format };
    const fontName = typeof props.fontName === 'string' ? props.fontName : undefined;
    delete props.fontName;
    if (fontName) {
      const fontId = document.findOrCreateFontId(fontName);
      if (fontId >= 0) props.fontId = fontId;
    }
    const result = JSON.parse(document.applyCharFormat(
      0,
      paraIdx,
      range.start,
      range.end,
      JSON.stringify(props),
    )) as { ok?: boolean; error?: string };
    if (result.ok === false) throw new Error(result.error ?? 'DOCX 글자 서식을 OMDX에 적용하지 못했습니다.');
  }
}

export function applyImportLayout(document: HwpDocument, options: RhwpImportOptions): void {
  if (options.pageDef && Object.keys(options.pageDef).length > 0) {
    const current = JSON.parse(document.getPageDef(0)) as RhwpPageDef;
    const next = { ...current, ...options.pageDef };
    const result = JSON.parse(document.setPageDef(0, JSON.stringify(next))) as { ok?: boolean; error?: string };
    if (result.ok === false) throw new Error(result.error ?? 'DOCX 페이지 설정을 OMDX에 적용하지 못했습니다.');
  }

  applyHeaderFooterGroup(document, true, options.headers, options.headerText);
  applyHeaderFooterGroup(document, false, options.footers, options.footerText);
}

function applyHeaderFooterGroup(
  document: HwpDocument,
  isHeader: boolean,
  values: RhwpImportOptions['headers'] | RhwpImportOptions['footers'],
  fallback?: string,
) {
  const defaultText = values?.default ?? fallback ?? values?.first;
  const evenText = values?.even;

  if (defaultText) {
    // Word's default reference is effectively the odd-page value when an even
    // reference is also present. rHWP exposes Both/Even/Odd rather than Word's
    // Default/Even split, so use Odd + Even in that case and Both otherwise.
    const applyTo = evenText ? 2 : 0;
    document.createHeaderFooter(0, isHeader, applyTo);
    document.insertTextInHeaderFooter(0, isHeader, applyTo, 0, 0, defaultText);
  }
  if (evenText) {
    document.createHeaderFooter(0, isHeader, 1);
    document.insertTextInHeaderFooter(0, isHeader, 1, 0, 0, evenText);
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
