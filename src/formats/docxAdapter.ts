import mammoth from 'mammoth/mammoth.browser';
import JSZip from 'jszip';
import type { DocumentFeatureInventory } from './canonical';
import { analyzeDocxFeatures } from './docxFeatures';
import { createRhwpFromHtml } from './rhwpRuntime';
import { extractDocxLayout, type DocxLayoutProfile } from './docxLayout';
import { enrichHtmlWithDocxFormatting, flattenTopLevelLists, parseDocxParagraphProfiles } from './docxOoxml';
import { patchHwpxSectionPageDefs } from './hwpxSections';

export interface DocxImportResult {
  hwpxBytes: Uint8Array;
  warnings: string[];
  layout: DocxLayoutProfile;
  features: DocumentFeatureInventory;
}

interface DocxTableStyle {
  widthPt?: number;
  border?: string;
  cells: Array<{
    widthPt?: number;
    padding?: string;
    verticalAlign?: string;
    background?: string;
    border?: string;
  }>;
}

interface DocxImageLayout {
  widthPx?: number;
  heightPx?: number;
}

export async function importDocx(bytes: Uint8Array): Promise<DocxImportResult> {
  const [layout, features] = await Promise.all([
    extractDocxLayout(bytes),
    analyzeDocxFeatures(bytes),
  ]);
  const arrayBuffer = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
  const converted = await convertDocxToLayoutHtml(arrayBuffer);
  const firstSection = layout.sections[0];
  const document = await createRhwpFromHtml(converted.html, {
    pageDef: layout.pageDef,
    headerText: layout.headerText,
    footerText: layout.footerText,
    headers: firstSection?.headers,
    footers: firstSection?.footers,
  });
  try {
    const patchedSections = await patchHwpxSectionPageDefs(
      document.exportHwpx(),
      layout.sections.map((section) => section.pageDef),
    );
    return {
      hwpxBytes: patchedSections.bytes,
      warnings: [
        ...converted.warnings,
        ...layout.warnings,
        ...patchedSections.warnings,
      ],
      layout,
      features,
    };
  } finally {
    document.free();
  }
}

export async function convertDocxToLayoutHtml(arrayBuffer: ArrayBuffer): Promise<{ html: string; warnings: string[] }> {
  const styleMap = ['u => u'];
  const classStyles = new Map<string, string>();
  let paragraphStyleIndex = 0;
  let runStyleIndex = 0;

  const paragraphTransform = mammoth.transforms.paragraph((paragraph) => {
    if (paragraph.numbering) return paragraph;
    const css = paragraphCss(paragraph.alignment, paragraph.indent);
    if (!css) return paragraph;

    const className = `omdx-p-${paragraphStyleIndex++}`;
    const styleName = `OMDX Paragraph ${paragraphStyleIndex}`;
    const tag = headingTag(paragraph.styleName);
    classStyles.set(className, css);
    styleMap.push(`p[style-name='${styleName}'] => ${tag}.${className}:fresh`);
    return { ...paragraph, styleId: className, styleName };
  });

  const runTransform = mammoth.transforms.run((run) => {
    const declarations: string[] = [];
    if (run.font) declarations.push(`font-family:${quoteCssFont(run.font)}`);
    if (run.fontSize) declarations.push(`font-size:${run.fontSize}pt`);
    if (run.highlight && run.highlight !== 'none') declarations.push(`background-color:${run.highlight}`);
    // Mammoth nests <strong>/<em>/<u> inside our generated style span. rHWP
    // treats a <span> as one atomic inline run and does not recursively inspect
    // those nested tags, so mirror the semantic flags onto the span CSS as well.
    if (run.isBold) declarations.push('font-weight:bold');
    if (run.isItalic) declarations.push('font-style:italic');
    const decorations: string[] = [];
    if (run.isUnderline) decorations.push('underline');
    if (run.isStrikethrough) decorations.push('line-through');
    if (decorations.length) declarations.push(`text-decoration:${decorations.join(' ')}`);
    if (run.verticalAlignment === 'superscript') declarations.push('vertical-align:super');
    if (run.verticalAlignment === 'subscript') declarations.push('vertical-align:sub');
    if (!declarations.length) return run;

    const className = `omdx-r-${runStyleIndex++}`;
    const styleName = `OMDX Run ${runStyleIndex}`;
    classStyles.set(className, declarations.join(';'));
    styleMap.push(`r[style-name='${styleName}'] => span.${className}`);
    return { ...run, styleId: className, styleName };
  });

  const result = await mammoth.convertToHtml(
    { arrayBuffer },
    {
      styleMap,
      transformDocument(document) {
        return runTransform(paragraphTransform(document));
      },
    },
  );

  const generated = inlineGeneratedStyles(result.value, classStyles);
  const enriched = await enrichHtmlFromDocxOoxml(arrayBuffer, generated);
  return {
    html: enriched.html,
    warnings: [
      ...result.messages.map((message: { message: string }) => message.message),
      ...enriched.warnings,
    ],
  };
}

/**
 * Mammoth intentionally targets semantic HTML and omits much of Word's visual
 * layout metadata. rHWP's HTML importer, on the other hand, can preserve table
 * geometry/borders/padding and image size when those values are present as CSS
 * or HTML attributes. Re-attach the relevant OOXML values before handing the
 * HTML to rHWP.
 */
export async function enrichHtmlFromDocxOoxml(
  arrayBuffer: ArrayBuffer,
  html: string,
): Promise<{ html: string; warnings: string[] }> {
  if (typeof DOMParser === 'undefined') return { html, warnings: [] };

  const zip = await JSZip.loadAsync(arrayBuffer);
  const documentXml = await zip.file('word/document.xml')?.async('text');
  if (!documentXml) return { html, warnings: [] };
  const stylesXml = await zip.file('word/styles.xml')?.async('text');

  const parsed = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  const body = parsed.body;
  const paragraphProfiles = parseDocxParagraphProfiles(documentXml, stylesXml);
  const formatting = enrichHtmlWithDocxFormatting(body, paragraphProfiles);
  const tableStyles = parseDocxTables(documentXml);
  const imageLayouts = parseDocxImages(documentXml);
  const htmlTables = Array.from(body.querySelectorAll('table'));
  const htmlImages = Array.from(body.querySelectorAll('img'));

  for (let index = 0; index < Math.min(htmlTables.length, tableStyles.length); index += 1) {
    applyTableStyle(htmlTables[index] as HTMLTableElement, tableStyles[index]);
  }

  for (let index = 0; index < Math.min(htmlImages.length, imageLayouts.length); index += 1) {
    const image = htmlImages[index] as HTMLImageElement;
    const layout = imageLayouts[index];
    if (layout.widthPx) image.setAttribute('width', formatNumber(layout.widthPx));
    if (layout.heightPx) image.setAttribute('height', formatNumber(layout.heightPx));
  }

  flattenTopLevelLists(body);

  // rHWP 0.8.6 only recognises <img> at the block parser level. Mammoth emits
  // inline images inside <p>, and parse_inline_content() ignores <img>, so the
  // image silently disappears. Split image-bearing paragraphs around each image
  // and promote the image to a top-level block. This approximates inline flow,
  // but preserves the actual image and its Word dimensions instead of dropping it.
  const hadInlineImages = htmlImages.some((image) => image.closest('p'));
  promoteParagraphImages(body);

  const warnings: string[] = [];
  if (hadInlineImages) {
    warnings.push('DOCX 인라인 그림은 현재 OMDX에서 그림 크기를 보존한 블록 배치로 변환됩니다.');
  }
  if (tableStyles.length > 0 && htmlTables.length !== tableStyles.length) {
    warnings.push('DOCX 표 일부의 레이아웃 메타데이터를 정확히 대응하지 못해 일부 표 서식이 단순화될 수 있습니다.');
  }
  if (formatting.totalParagraphs > 0 && formatting.matchedParagraphs / formatting.totalParagraphs < 0.7) {
    warnings.push('DOCX 문단 일부를 OOXML 서식과 정확히 대응하지 못해 복합 필드/텍스트박스 서식이 단순화될 수 있습니다.');
  }
  if (formatting.sectionBreakApproximation) {
    warnings.push('표/목록 내부와 인접한 DOCX 구역 나누기는 가장 가까운 블록 경계로 정규화됩니다.');
  }

  return { html: body.innerHTML, warnings };
}

function paragraphCss(
  alignment: string | null,
  indent: { start?: string | null; end?: string | null; firstLine?: string | null; hanging?: string | null },
): string {
  const declarations: string[] = [];
  const mappedAlignment = alignment === 'both' || alignment === 'distribute' ? 'justify' : alignment;
  if (mappedAlignment && ['left', 'center', 'right', 'justify'].includes(mappedAlignment)) {
    declarations.push(`text-align:${mappedAlignment}`);
  }
  const start = twipsToPoints(indent.start);
  const end = twipsToPoints(indent.end);
  const firstLine = twipsToPoints(indent.firstLine);
  const hanging = twipsToPoints(indent.hanging);
  if (start) declarations.push(`margin-left:${start}pt`);
  if (end) declarations.push(`margin-right:${end}pt`);
  if (firstLine) declarations.push(`text-indent:${firstLine}pt`);
  else if (hanging) declarations.push(`text-indent:-${hanging}pt`);
  return declarations.join(';');
}

function twipsToPoints(value?: string | null): number | undefined {
  if (!value) return undefined;
  const number = Number(value);
  if (!Number.isFinite(number) || number === 0) return undefined;
  return Math.round((number / 20) * 100) / 100;
}

function headingTag(styleName: string | null): string {
  const match = styleName?.match(/^heading\s*([1-6])$/i);
  return match ? `h${match[1]}` : 'p';
}

function quoteCssFont(font: string): string {
  return /\s/.test(font) ? `'${font.replaceAll("'", "\\'")}'` : font;
}

function inlineGeneratedStyles(html: string, styles: Map<string, string>): string {
  let result = html;
  for (const [className, css] of styles) {
    result = result.replaceAll(`class="${className}"`, `style="${css}"`);
  }
  return result;
}

function parseDocxTables(documentXml: string): DocxTableStyle[] {
  return matchElements(documentXml, 'w:tbl').map((tableXml) => {
    const tblPr = firstElement(tableXml, 'w:tblPr') ?? '';
    const width = firstStartTag(tblPr, 'w:tblW');
    const widthPt = width?.includes('w:type="dxa"') ? twipsAttributeToPt(width, 'w') : undefined;
    const tableBorders = firstElement(tblPr, 'w:tblBorders');
    const border = tableBorders ? borderCssFromContainer(tableBorders) : undefined;
    const cells = matchElements(tableXml, 'w:tc').map((cellXml) => {
      const tcPr = firstElement(cellXml, 'w:tcPr') ?? '';
      const tcW = firstStartTag(tcPr, 'w:tcW');
      const tcMar = firstElement(tcPr, 'w:tcMar');
      const tcBorders = firstElement(tcPr, 'w:tcBorders');
      const shd = firstStartTag(tcPr, 'w:shd');
      const vAlign = firstStartTag(tcPr, 'w:vAlign');
      return {
        widthPt: tcW?.includes('w:type="dxa"') ? twipsAttributeToPt(tcW, 'w') : undefined,
        padding: tcMar ? paddingCss(tcMar) : undefined,
        verticalAlign: vAlign ? mapVerticalAlign(attribute(vAlign, 'val')) : undefined,
        background: shd ? normalizeWordColor(attribute(shd, 'fill')) : undefined,
        border: tcBorders ? borderCssFromContainer(tcBorders) : border,
      };
    });
    return { widthPt, border, cells };
  });
}

function parseDocxImages(documentXml: string): DocxImageLayout[] {
  const layouts: DocxImageLayout[] = [];
  const drawingRegex = /<w:drawing\b[\s\S]*?<\/w:drawing>/g;
  for (const match of documentXml.matchAll(drawingRegex)) {
    const drawing = match[0];
    const extent = drawing.match(/<wp:extent\b([^>]*)\/?\s*>/)?.[1];
    if (!extent) {
      layouts.push({});
      continue;
    }
    const cx = Number(attribute(extent, 'cx'));
    const cy = Number(attribute(extent, 'cy'));
    layouts.push({
      widthPx: Number.isFinite(cx) && cx > 0 ? cx / 9525 : undefined,
      heightPx: Number.isFinite(cy) && cy > 0 ? cy / 9525 : undefined,
    });
  }
  return layouts;
}

function applyTableStyle(table: HTMLTableElement, style: DocxTableStyle) {
  table.style.borderCollapse = 'collapse';
  table.style.tableLayout = 'fixed';
  if (style.widthPt) table.style.width = `${formatNumber(style.widthPt)}pt`;
  if (style.border) applyCssDeclarations(table, style.border);

  const cells = Array.from(table.querySelectorAll('td,th')) as HTMLTableCellElement[];
  for (let index = 0; index < Math.min(cells.length, style.cells.length); index += 1) {
    const cell = cells[index];
    const cellStyle = style.cells[index];
    if (cellStyle.widthPt) cell.style.width = `${formatNumber(cellStyle.widthPt)}pt`;
    if (cellStyle.padding) applyCssDeclarations(cell, cellStyle.padding);
    if (cellStyle.verticalAlign) cell.style.verticalAlign = cellStyle.verticalAlign;
    if (cellStyle.background) cell.style.backgroundColor = cellStyle.background;
    if (cellStyle.border) applyCssDeclarations(cell, cellStyle.border);
  }
}

function promoteParagraphImages(body: HTMLElement) {
  let paragraph = Array.from(body.querySelectorAll('p')).find((item) => item.querySelector('img'));
  while (paragraph) {
    const image = paragraph.querySelector('img');
    if (!image) break;
    const owner = paragraph.ownerDocument;
    const beforeRange = owner.createRange();
    beforeRange.selectNodeContents(paragraph);
    beforeRange.setEndBefore(image);
    const afterRange = owner.createRange();
    afterRange.selectNodeContents(paragraph);
    afterRange.setStartAfter(image);

    const replacements: Node[] = [];
    const before = cloneParagraphShell(paragraph);
    before.append(beforeRange.cloneContents());
    if (hasMeaningfulContent(before)) replacements.push(before);
    replacements.push(image.cloneNode(true));
    const after = cloneParagraphShell(paragraph);
    after.append(afterRange.cloneContents());
    if (hasMeaningfulContent(after)) replacements.push(after);

    paragraph.replaceWith(...replacements);
    paragraph = Array.from(body.querySelectorAll('p')).find((item) => item.querySelector('img'));
  }
}

function cloneParagraphShell(paragraph: HTMLParagraphElement): HTMLParagraphElement {
  const clone = paragraph.cloneNode(false) as HTMLParagraphElement;
  clone.removeAttribute('id');
  return clone;
}

function hasMeaningfulContent(element: HTMLElement): boolean {
  return Boolean(element.textContent?.trim() || element.querySelector('br,img'));
}

function matchElements(xml: string, tag: string): string[] {
  const escaped = tag.replace(':', '\\:');
  return [...xml.matchAll(new RegExp(`<${escaped}\\b[\\s\\S]*?<\\/${escaped}>`, 'g'))].map((match) => match[0]);
}

function firstElement(xml: string, tag: string): string | undefined {
  const escaped = tag.replace(':', '\\:');
  return xml.match(new RegExp(`<${escaped}\\b[\\s\\S]*?<\\/${escaped}>`))?.[0];
}

function firstStartTag(xml: string, tag: string): string | undefined {
  const escaped = tag.replace(':', '\\:');
  return xml.match(new RegExp(`<${escaped}\\b([^>]*)\\/?\\s*>`))?.[1];
}

function attribute(attributes: string, name: string): string | undefined {
  const escaped = name.replace(':', '\\:');
  return attributes.match(new RegExp(`(?:w:)?${escaped}="([^"]+)"`))?.[1];
}

function twipsAttributeToPt(attributes: string, name: string): number | undefined {
  const value = Number(attribute(attributes, name));
  return Number.isFinite(value) && value > 0 ? value / 20 : undefined;
}

function paddingCss(container: string): string | undefined {
  const declarations: string[] = [];
  for (const [wordSide, cssSide] of [['top', 'top'], ['right', 'right'], ['bottom', 'bottom'], ['left', 'left']] as const) {
    const tag = firstStartTag(container, `w:${wordSide}`);
    const pt = tag?.includes('w:type="dxa"') ? twipsAttributeToPt(tag, 'w') : undefined;
    if (pt !== undefined) declarations.push(`padding-${cssSide}:${formatNumber(pt)}pt`);
  }
  return declarations.length ? declarations.join(';') : undefined;
}

function borderCssFromContainer(container: string): string | undefined {
  const declarations: string[] = [];
  for (const [wordSide, cssSide] of [['top', 'top'], ['right', 'right'], ['bottom', 'bottom'], ['left', 'left']] as const) {
    const tag = firstStartTag(container, `w:${wordSide}`);
    if (!tag) continue;
    const value = attribute(tag, 'val');
    const style = mapBorderStyle(value);
    if (style === 'none') {
      declarations.push(`border-${cssSide}:none`);
      continue;
    }
    const size = Number(attribute(tag, 'sz'));
    const widthPt = Number.isFinite(size) && size > 0 ? size / 8 : 0.5;
    const color = normalizeWordColor(attribute(tag, 'color')) ?? '#000000';
    declarations.push(`border-${cssSide}:${formatNumber(widthPt)}pt ${style} ${color}`);
  }
  return declarations.length ? declarations.join(';') : undefined;
}

function mapBorderStyle(value?: string): string {
  switch (value) {
    case 'nil':
    case 'none': return 'none';
    case 'dashed':
    case 'dashSmallGap': return 'dashed';
    case 'dotted': return 'dotted';
    case 'double': return 'double';
    default: return 'solid';
  }
}

function mapVerticalAlign(value?: string): string | undefined {
  if (value === 'center') return 'middle';
  if (value === 'bottom') return 'bottom';
  if (value === 'top') return 'top';
  return undefined;
}

function normalizeWordColor(value?: string): string | undefined {
  if (!value || value === 'auto' || value === 'none') return undefined;
  const normalized = value.replace(/^#/, '');
  return /^[0-9a-f]{6}$/i.test(normalized) ? `#${normalized}` : undefined;
}

function applyCssDeclarations(element: HTMLElement, css: string) {
  for (const declaration of css.split(';')) {
    const colon = declaration.indexOf(':');
    if (colon <= 0) continue;
    const property = declaration.slice(0, colon).trim();
    const value = declaration.slice(colon + 1).trim();
    if (property && value) element.style.setProperty(property, value);
  }
}

function formatNumber(value: number): string {
  return String(Math.round(value * 100) / 100);
}
