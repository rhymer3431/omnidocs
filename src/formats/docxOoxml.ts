const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export const OMDX_PARA_FORMAT_ATTR = 'data-omdx-para-format';
export const OMDX_CHAR_FORMAT_ATTR = 'data-omdx-char-format';
export const OMDX_SECTION_BREAK_ATTR = 'data-omdx-section-break-after';
export const OMDX_LIST_KIND_ATTR = 'data-omdx-list-kind';
export const OMDX_LIST_LEVEL_ATTR = 'data-omdx-list-level';

export interface RhwpParagraphFormat {
  alignment?: 'left' | 'right' | 'center' | 'justify' | 'distribute';
  lineSpacing?: number;
  lineSpacingType?: 'Percent' | 'Fixed' | 'Minimum';
  indent?: number;
  marginLeft?: number;
  marginRight?: number;
  spacingBefore?: number;
  spacingAfter?: number;
  widowOrphan?: boolean;
  keepWithNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  autoSpaceKrEn?: boolean;
  autoSpaceKrNum?: boolean;
}

export interface RhwpCharacterFormat {
  fontName?: string;
  fontSize?: number;
  textColor?: string;
  shadeColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  emboss?: boolean;
  engrave?: boolean;
}

export interface DocxRunProfile {
  text: string;
  format: RhwpCharacterFormat;
}

export interface DocxParagraphProfile {
  text: string;
  hasDrawing: boolean;
  sectionBreakAfter: boolean;
  paragraphFormat: RhwpParagraphFormat;
  runs: DocxRunProfile[];
}

interface StyleDefinition {
  basedOn?: string;
  paragraphFormat: RhwpParagraphFormat;
  characterFormat: RhwpCharacterFormat;
}

interface StyleContext {
  defaultParagraphFormat: RhwpParagraphFormat;
  defaultCharacterFormat: RhwpCharacterFormat;
  defaultParagraphStyleId?: string;
  defaultCharacterStyleId?: string;
  paragraphStyles: Map<string, StyleDefinition>;
  characterStyles: Map<string, StyleDefinition>;
}

export interface HtmlFormattingEnrichment {
  matchedParagraphs: number;
  totalParagraphs: number;
  sectionBreakApproximation: boolean;
}

export function encodeOmdxFormat(value: object): string {
  return encodeURIComponent(JSON.stringify(value));
}

export function decodeOmdxFormat<T>(value: string | null): T | undefined {
  if (!value) return undefined;
  try {
    return JSON.parse(decodeURIComponent(value)) as T;
  } catch {
    return undefined;
  }
}

export function parseDocxParagraphProfiles(
  documentXml: string,
  stylesXml?: string,
): DocxParagraphProfile[] {
  if (typeof DOMParser === 'undefined') return [];
  const document = parseXml(documentXml);
  if (!document) return [];

  const styles = parseStyleContext(stylesXml);
  const paragraphElements = Array.from(document.getElementsByTagNameNS(WORD_NS, 'p'));

  return paragraphElements.map((paragraph) => {
    const pPr = childWordElement(paragraph, 'pPr');
    const paragraphStyleId = wordAttr(childWordElement(pPr, 'pStyle'), 'val')
      ?? styles.defaultParagraphStyleId;
    const paragraphStyle = resolveStyle(paragraphStyleId, styles.paragraphStyles);
    const paragraphFormat = mergeDefined<RhwpParagraphFormat>(
      styles.defaultParagraphFormat,
      paragraphStyle.paragraphFormat,
      parseParagraphFormat(pPr),
    );
    const paragraphCharacterFormat = mergeDefined<RhwpCharacterFormat>(
      styles.defaultCharacterFormat,
      paragraphStyle.characterFormat,
    );

    const runs = runsOwnedByParagraph(paragraph).map((run) => {
      const rPr = childWordElement(run, 'rPr');
      const characterStyleId = wordAttr(childWordElement(rPr, 'rStyle'), 'val')
        ?? styles.defaultCharacterStyleId;
      const characterStyle = resolveStyle(characterStyleId, styles.characterStyles);
      return {
        text: readRunText(run),
        format: mergeDefined<RhwpCharacterFormat>(
          paragraphCharacterFormat,
          characterStyle.characterFormat,
          parseCharacterFormat(rPr),
        ),
      };
    });

    return {
      text: runs.map((run) => run.text).join(''),
      hasDrawing: paragraph.getElementsByTagNameNS(WORD_NS, 'drawing').length > 0
        || paragraph.getElementsByTagNameNS(WORD_NS, 'pict').length > 0,
      sectionBreakAfter: Boolean(childWordElement(pPr, 'sectPr')),
      paragraphFormat,
      runs,
    };
  });
}

export function enrichHtmlWithDocxFormatting(
  body: HTMLElement,
  profiles: DocxParagraphProfile[],
): HtmlFormattingEnrichment {
  const htmlParagraphs = Array.from(body.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li')) as HTMLElement[];
  let profileCursor = 0;
  let matchedParagraphs = 0;
  let lastMatched: HTMLElement | null = null;
  let sectionBreakApproximation = false;

  for (const element of htmlParagraphs) {
    const targetText = normalizeMatchText(paragraphElementText(element));
    const targetHasImage = Boolean(element.querySelector('img'));
    const matchIndex = findProfileMatch(profiles, profileCursor, targetText, targetHasImage);
    if (matchIndex < 0) continue;

    for (let skipped = profileCursor; skipped < matchIndex; skipped += 1) {
      if (profiles[skipped].sectionBreakAfter && lastMatched) {
        const topLevel = topLevelBlock(lastMatched, body);
        topLevel.setAttribute(OMDX_SECTION_BREAK_ATTR, '1');
        if (topLevel !== lastMatched) sectionBreakApproximation = true;
      }
    }

    const profile = profiles[matchIndex];
    applyParagraphMetadata(element, profile.paragraphFormat);
    applyRunMetadata(element, profile.runs);

    if (profile.sectionBreakAfter) {
      const topLevel = topLevelBlock(element, body);
      topLevel.setAttribute(OMDX_SECTION_BREAK_ATTR, '1');
      if (topLevel !== element) sectionBreakApproximation = true;
    }

    matchedParagraphs += 1;
    lastMatched = element;
    profileCursor = matchIndex + 1;
  }

  for (let skipped = profileCursor; skipped < profiles.length; skipped += 1) {
    if (profiles[skipped].sectionBreakAfter && lastMatched) {
      const topLevel = topLevelBlock(lastMatched, body);
      topLevel.setAttribute(OMDX_SECTION_BREAK_ATTR, '1');
      if (topLevel !== lastMatched) sectionBreakApproximation = true;
    }
  }

  return { matchedParagraphs, totalParagraphs: profiles.length, sectionBreakApproximation };
}

/**
 * rHWP 0.8.6 imports both <ol> and <ul> as literal bullet-prefixed text. Flatten
 * top-level Word lists into annotated paragraphs so rhwpRuntime can apply native
 * Number/Bullet paragraph heads instead of turning numbered lists into bullets.
 */
export function flattenTopLevelLists(body: HTMLElement): void {
  const lists = Array.from(body.children).filter((element) => element.tagName === 'OL' || element.tagName === 'UL');
  for (const list of lists) {
    const fragment = body.ownerDocument.createDocumentFragment();
    appendListParagraphs(list as HTMLOListElement | HTMLUListElement, 0, fragment);
    list.replaceWith(fragment);
  }
}

function parseStyleContext(stylesXml?: string): StyleContext {
  const context: StyleContext = {
    defaultParagraphFormat: {},
    defaultCharacterFormat: {},
    paragraphStyles: new Map(),
    characterStyles: new Map(),
  };
  if (!stylesXml || typeof DOMParser === 'undefined') return context;
  const document = parseXml(stylesXml);
  if (!document) return context;

  const defaults = firstWordElement(document, 'docDefaults');
  const pPrDefault = childWordElement(childWordElement(defaults, 'pPrDefault'), 'pPr');
  const rPrDefault = childWordElement(childWordElement(defaults, 'rPrDefault'), 'rPr');
  context.defaultParagraphFormat = parseParagraphFormat(pPrDefault);
  context.defaultCharacterFormat = parseCharacterFormat(rPrDefault);

  for (const style of Array.from(document.getElementsByTagNameNS(WORD_NS, 'style'))) {
    const id = wordAttr(style, 'styleId');
    const type = wordAttr(style, 'type');
    if (!id || (type !== 'paragraph' && type !== 'character')) continue;
    const definition: StyleDefinition = {
      basedOn: wordAttr(childWordElement(style, 'basedOn'), 'val'),
      paragraphFormat: parseParagraphFormat(childWordElement(style, 'pPr')),
      characterFormat: parseCharacterFormat(childWordElement(style, 'rPr')),
    };
    if (type === 'paragraph') {
      context.paragraphStyles.set(id, definition);
      if (wordOnOffAttribute(style, 'default') === true) context.defaultParagraphStyleId = id;
    } else {
      context.characterStyles.set(id, definition);
      if (wordOnOffAttribute(style, 'default') === true) context.defaultCharacterStyleId = id;
    }
  }

  return context;
}

function resolveStyle(
  styleId: string | undefined,
  styles: Map<string, StyleDefinition>,
  seen = new Set<string>(),
): StyleDefinition {
  if (!styleId || seen.has(styleId)) return { paragraphFormat: {}, characterFormat: {} };
  const current = styles.get(styleId);
  if (!current) return { paragraphFormat: {}, characterFormat: {} };
  seen.add(styleId);
  const base = resolveStyle(current.basedOn, styles, seen);
  return {
    paragraphFormat: mergeDefined(base.paragraphFormat, current.paragraphFormat),
    characterFormat: mergeDefined(base.characterFormat, current.characterFormat),
  };
}

function parseParagraphFormat(pPr?: Element): RhwpParagraphFormat {
  if (!pPr) return {};
  const result: RhwpParagraphFormat = {};

  const alignment = wordAttr(childWordElement(pPr, 'jc'), 'val');
  if (alignment) {
    if (alignment === 'both') result.alignment = 'justify';
    else if (alignment === 'distribute') result.alignment = 'distribute';
    else if (alignment === 'left' || alignment === 'right' || alignment === 'center') result.alignment = alignment;
  }

  const indent = childWordElement(pPr, 'ind');
  if (indent) {
    const left = twipsToHwpUnit(wordAttr(indent, 'start') ?? wordAttr(indent, 'left'));
    const right = twipsToHwpUnit(wordAttr(indent, 'end') ?? wordAttr(indent, 'right'));
    const firstLine = twipsToHwpUnit(wordAttr(indent, 'firstLine'));
    const hanging = twipsToHwpUnit(wordAttr(indent, 'hanging'));
    if (left !== undefined) result.marginLeft = left;
    if (right !== undefined) result.marginRight = right;
    if (firstLine !== undefined) result.indent = firstLine;
    else if (hanging !== undefined) result.indent = -hanging;
  }

  const spacing = childWordElement(pPr, 'spacing');
  if (spacing) {
    const before = twipsToHwpUnit(wordAttr(spacing, 'before'));
    const after = twipsToHwpUnit(wordAttr(spacing, 'after'));
    if (before !== undefined) result.spacingBefore = before;
    if (after !== undefined) result.spacingAfter = after;

    const line = finiteNumber(wordAttr(spacing, 'line'));
    if (line !== undefined) {
      const lineRule = wordAttr(spacing, 'lineRule') ?? 'auto';
      if (lineRule === 'exact') {
        result.lineSpacing = Math.round(line * 5);
        result.lineSpacingType = 'Fixed';
      } else if (lineRule === 'atLeast') {
        result.lineSpacing = Math.round(line * 5);
        result.lineSpacingType = 'Minimum';
      } else {
        result.lineSpacing = Math.round((line / 240) * 100);
        result.lineSpacingType = 'Percent';
      }
    }
  }

  assignOnOff(result, 'keepWithNext', wordElementOnOff(pPr, 'keepNext'));
  assignOnOff(result, 'keepLines', wordElementOnOff(pPr, 'keepLines'));
  assignOnOff(result, 'pageBreakBefore', wordElementOnOff(pPr, 'pageBreakBefore'));
  assignOnOff(result, 'widowOrphan', wordElementOnOff(pPr, 'widowControl'));
  assignOnOff(result, 'autoSpaceKrEn', wordElementOnOff(pPr, 'autoSpaceDE'));
  assignOnOff(result, 'autoSpaceKrNum', wordElementOnOff(pPr, 'autoSpaceDN'));
  return result;
}

function parseCharacterFormat(rPr?: Element): RhwpCharacterFormat {
  if (!rPr) return {};
  const result: RhwpCharacterFormat = {};

  const fonts = childWordElement(rPr, 'rFonts');
  const fontName = wordAttr(fonts, 'eastAsia')
    ?? wordAttr(fonts, 'hAnsi')
    ?? wordAttr(fonts, 'ascii')
    ?? wordAttr(fonts, 'cs');
  if (fontName) result.fontName = fontName;

  const halfPoints = finiteNumber(wordAttr(childWordElement(rPr, 'sz'), 'val'));
  if (halfPoints !== undefined) result.fontSize = Math.round(halfPoints * 50);

  const color = normalizeHexColor(wordAttr(childWordElement(rPr, 'color'), 'val'));
  if (color) result.textColor = color;

  const shading = normalizeHexColor(wordAttr(childWordElement(rPr, 'shd'), 'fill'));
  const highlight = wordHighlightColor(wordAttr(childWordElement(rPr, 'highlight'), 'val'));
  if (highlight ?? shading) result.shadeColor = highlight ?? shading;

  assignOnOff(result, 'bold', wordElementOnOff(rPr, 'b'));
  assignOnOff(result, 'italic', wordElementOnOff(rPr, 'i'));
  assignOnOff(result, 'strikethrough', wordElementOnOff(rPr, 'strike'));
  assignOnOff(result, 'emboss', wordElementOnOff(rPr, 'emboss'));
  assignOnOff(result, 'engrave', wordElementOnOff(rPr, 'imprint'));

  const underline = childWordElement(rPr, 'u');
  if (underline) {
    const value = wordAttr(underline, 'val');
    result.underline = value !== 'none' && value !== '0' && value !== 'false';
  }

  const verticalAlign = wordAttr(childWordElement(rPr, 'vertAlign'), 'val');
  if (verticalAlign === 'superscript') result.superscript = true;
  if (verticalAlign === 'subscript') result.subscript = true;
  return result;
}

function applyParagraphMetadata(element: HTMLElement, format: RhwpParagraphFormat) {
  if (!hasOwnValues(format)) return;
  element.setAttribute(OMDX_PARA_FORMAT_ATTR, encodeOmdxFormat(format));

  if (format.alignment) element.style.textAlign = format.alignment === 'distribute' ? 'justify' : format.alignment;
  if (format.marginLeft !== undefined) element.style.marginLeft = `${formatNumber(format.marginLeft / 100)}pt`;
  if (format.marginRight !== undefined) element.style.marginRight = `${formatNumber(format.marginRight / 100)}pt`;
  if (format.indent !== undefined) element.style.textIndent = `${formatNumber(format.indent / 100)}pt`;
  if (format.spacingBefore !== undefined) element.style.marginTop = `${formatNumber(format.spacingBefore / 100)}pt`;
  if (format.spacingAfter !== undefined) element.style.marginBottom = `${formatNumber(format.spacingAfter / 100)}pt`;
  if (format.lineSpacing !== undefined) {
    if (format.lineSpacingType === 'Percent') {
      element.style.lineHeight = `${format.lineSpacing}%`;
    } else {
      const points = format.lineSpacing / 100;
      element.style.lineHeight = `${formatNumber(points * 96 / 72)}px`;
    }
  }
}

function applyRunMetadata(element: HTMLElement, runs: DocxRunProfile[]) {
  const runText = runs.map((run) => run.text).join('');
  const textNodes = paragraphTextNodes(element);
  const htmlText = textNodes.map((node) => node.data).join('');
  if (runText !== htmlText || !runText) return;

  const ranges: Array<{ start: number; end: number; format: RhwpCharacterFormat }> = [];
  let cursor = 0;
  for (const run of runs) {
    const start = cursor;
    cursor += run.text.length;
    if (run.text && hasOwnValues(run.format)) ranges.push({ start, end: cursor, format: run.format });
  }
  if (!ranges.length) return;

  let nodeStart = 0;
  for (const node of textNodes) {
    const text = node.data;
    const nodeEnd = nodeStart + text.length;
    const boundaries = new Set<number>([0, text.length]);
    for (const range of ranges) {
      const overlapStart = Math.max(nodeStart, range.start);
      const overlapEnd = Math.min(nodeEnd, range.end);
      if (overlapStart < overlapEnd) {
        boundaries.add(overlapStart - nodeStart);
        boundaries.add(overlapEnd - nodeStart);
      }
    }

    const ordered = [...boundaries].sort((a, b) => a - b);
    const fragment = node.ownerDocument.createDocumentFragment();
    for (let index = 0; index < ordered.length - 1; index += 1) {
      const localStart = ordered[index];
      const localEnd = ordered[index + 1];
      const value = text.slice(localStart, localEnd);
      if (!value) continue;
      const globalStart = nodeStart + localStart;
      const range = ranges.find((candidate) => candidate.start <= globalStart && candidate.end > globalStart);
      if (!range) {
        fragment.append(node.ownerDocument.createTextNode(value));
        continue;
      }
      const span = node.ownerDocument.createElement('span');
      span.setAttribute(OMDX_CHAR_FORMAT_ATTR, encodeOmdxFormat(range.format));
      applyCharacterCss(span, range.format);
      span.textContent = value;
      fragment.append(span);
    }
    node.replaceWith(fragment);
    nodeStart = nodeEnd;
  }
}

function applyCharacterCss(span: HTMLElement, format: RhwpCharacterFormat) {
  if (format.fontName) span.style.fontFamily = format.fontName;
  if (format.fontSize !== undefined) span.style.fontSize = `${formatNumber(format.fontSize / 100)}pt`;
  if (format.textColor) span.style.color = format.textColor;
  if (format.shadeColor) span.style.backgroundColor = format.shadeColor;
  if (format.bold) span.style.fontWeight = 'bold';
  if (format.italic) span.style.fontStyle = 'italic';
  const decoration: string[] = [];
  if (format.underline) decoration.push('underline');
  if (format.strikethrough) decoration.push('line-through');
  if (decoration.length) span.style.textDecoration = decoration.join(' ');
  if (format.superscript) span.style.verticalAlign = 'super';
  if (format.subscript) span.style.verticalAlign = 'sub';
}

function findProfileMatch(
  profiles: DocxParagraphProfile[],
  start: number,
  targetText: string,
  targetHasImage: boolean,
): number {
  const limit = Math.min(profiles.length, start + 32);
  for (let index = start; index < limit; index += 1) {
    const profile = profiles[index];
    const profileText = normalizeMatchText(profile.text);
    if (profileText === targetText && (targetText !== '' || profile.hasDrawing === targetHasImage)) return index;
  }
  return -1;
}

function paragraphElementText(element: HTMLElement): string {
  if (element.tagName !== 'LI') return element.textContent ?? '';
  const clone = element.cloneNode(true) as HTMLElement;
  clone.querySelectorAll('ol,ul').forEach((list) => list.remove());
  return clone.textContent ?? '';
}

function paragraphTextNodes(element: HTMLElement): Text[] {
  const view = element.ownerDocument.defaultView;
  const walker = element.ownerDocument.createTreeWalker(element, view?.NodeFilter.SHOW_TEXT ?? 4);
  const result: Text[] = [];
  let current = walker.nextNode();
  while (current) {
    const text = current as Text;
    let ancestor = text.parentElement;
    let excluded = false;
    while (ancestor && ancestor !== element) {
      if (ancestor.tagName === 'OL' || ancestor.tagName === 'UL' || ancestor.tagName === 'TABLE') {
        excluded = true;
        break;
      }
      ancestor = ancestor.parentElement;
    }
    if (!excluded) result.push(text);
    current = walker.nextNode();
  }
  return result;
}

function topLevelBlock(element: HTMLElement, body: HTMLElement): HTMLElement {
  let current = element;
  while (current.parentElement && current.parentElement !== body) current = current.parentElement;
  return current;
}

function appendListParagraphs(
  list: HTMLOListElement | HTMLUListElement,
  level: number,
  target: DocumentFragment,
) {
  const kind = list.tagName === 'OL' ? 'number' : 'bullet';
  for (const child of Array.from(list.children)) {
    if (child.tagName !== 'LI') continue;
    const item = child as HTMLLIElement;
    const paragraph = item.ownerDocument.createElement('p');
    for (const attribute of Array.from(item.attributes)) {
      if (attribute.name !== 'value') paragraph.setAttribute(attribute.name, attribute.value);
    }
    paragraph.setAttribute(OMDX_LIST_KIND_ATTR, kind);
    paragraph.setAttribute(OMDX_LIST_LEVEL_ATTR, String(Math.min(level, 6)));

    const nestedLists: Array<HTMLOListElement | HTMLUListElement> = [];
    for (const node of Array.from(item.childNodes)) {
      if (node.nodeType === 1 && ((node as Element).tagName === 'OL' || (node as Element).tagName === 'UL')) {
        nestedLists.push(node as HTMLOListElement | HTMLUListElement);
      } else {
        paragraph.append(node.cloneNode(true));
      }
    }
    if (paragraph.textContent?.trim() || paragraph.querySelector('img,br')) target.append(paragraph);
    for (const nested of nestedLists) appendListParagraphs(nested, level + 1, target);
  }
}

function runsOwnedByParagraph(paragraph: Element): Element[] {
  return Array.from(paragraph.getElementsByTagNameNS(WORD_NS, 'r')).filter((run) => {
    let current: Element | null = run.parentElement;
    while (current) {
      if (current.namespaceURI === WORD_NS && current.localName === 'p') return current === paragraph;
      current = current.parentElement;
    }
    return false;
  });
}

function readRunText(run: Element): string {
  let result = '';
  const walk = (node: Node) => {
    if (node.nodeType === 1) {
      const element = node as Element;
      if (element.namespaceURI === WORD_NS) {
        switch (element.localName) {
          case 't':
          case 'delText':
            result += element.textContent ?? '';
            return;
          case 'tab':
          case 'ptab':
            result += '\t';
            return;
          case 'br':
          case 'cr':
            result += '\n';
            return;
          case 'noBreakHyphen':
            result += '\u2011';
            return;
          case 'softHyphen':
            result += '\u00ad';
            return;
          case 'instrText':
            return;
          default:
            break;
        }
      }
    }
    node.childNodes.forEach(walk);
  };
  run.childNodes.forEach(walk);
  return result;
}

function parseXml(xml: string): XMLDocument | undefined {
  const parsed = new DOMParser().parseFromString(xml, 'application/xml');
  if (parsed.getElementsByTagName('parsererror').length > 0) return undefined;
  return parsed;
}

function firstWordElement(parent: Document | Element, localName: string): Element | undefined {
  return parent.getElementsByTagNameNS(WORD_NS, localName)[0] ?? undefined;
}

function childWordElement(parent: Element | undefined, localName: string): Element | undefined {
  if (!parent) return undefined;
  return Array.from(parent.children).find(
    (child) => child.namespaceURI === WORD_NS && child.localName === localName,
  );
}

function wordAttr(element: Element | undefined, name: string): string | undefined {
  if (!element) return undefined;
  return element.getAttributeNS(WORD_NS, name)
    ?? element.getAttribute(`w:${name}`)
    ?? element.getAttribute(name)
    ?? undefined;
}

function wordElementOnOff(parent: Element, childName: string): boolean | undefined {
  const child = childWordElement(parent, childName);
  if (!child) return undefined;
  const value = wordAttr(child, 'val');
  return value === undefined || !['0', 'false', 'off', 'none'].includes(value.toLowerCase());
}

function wordOnOffAttribute(element: Element, name: string): boolean | undefined {
  const value = wordAttr(element, name);
  if (value === undefined) return undefined;
  return !['0', 'false', 'off', 'none'].includes(value.toLowerCase());
}

function twipsToHwpUnit(value?: string): number | undefined {
  const number = finiteNumber(value);
  return number === undefined ? undefined : Math.round(number * 5);
}

function finiteNumber(value?: string): number | undefined {
  if (value === undefined || value === '') return undefined;
  const number = Number(value);
  return Number.isFinite(number) ? number : undefined;
}

function normalizeHexColor(value?: string): string | undefined {
  if (!value || value === 'auto' || value === 'none') return undefined;
  const normalized = value.replace(/^#/, '');
  return /^[0-9a-f]{6}$/i.test(normalized) ? `#${normalized.toUpperCase()}` : undefined;
}

function wordHighlightColor(value?: string): string | undefined {
  const colors: Record<string, string> = {
    black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00',
    magenta: '#FF00FF', red: '#FF0000', yellow: '#FFFF00', white: '#FFFFFF',
    darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080',
    darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0',
  };
  return value ? colors[value] : undefined;
}

function normalizeMatchText(value: string): string {
  return value.replaceAll('\u00a0', ' ').replace(/\s+/g, ' ').trim();
}

function assignOnOff<T extends object, K extends keyof T>(target: T, key: K, value: boolean | undefined) {
  if (value !== undefined) target[key] = value as T[K];
}

function mergeDefined<T extends object>(...values: Array<Partial<T> | undefined>): T {
  const result: Record<string, unknown> = {};
  for (const value of values) {
    if (!value) continue;
    for (const [key, item] of Object.entries(value)) {
      if (item !== undefined) result[key] = item;
    }
  }
  return result as T;
}

function hasOwnValues(value: object): boolean {
  return Object.values(value).some((item) => item !== undefined);
}

function formatNumber(value: number): string {
  return String(Math.round(value * 100) / 100);
}
