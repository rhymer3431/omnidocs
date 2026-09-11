import JSZip from 'jszip';

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

export interface HwpxDocxExportResult {
  bytes: Uint8Array;
  warnings: string[];
}

interface FontCatalog {
  hangul: Map<string, string>;
  latin: Map<string, string>;
  hanja: Map<string, string>;
  japanese: Map<string, string>;
  other: Map<string, string>;
  symbol: Map<string, string>;
  user: Map<string, string>;
}

interface CharacterStyle {
  id: string;
  xml: string;
}

interface ParagraphStyle {
  id: string;
  xml: string;
  numbering?: { kind: 'number' | 'bullet'; key: string; level: number };
}

interface BorderStyle {
  id: string;
  tableBorders: string;
  cellBorders: string;
  shading?: string;
}

interface StyleCatalog {
  fonts: FontCatalog;
  characters: Map<string, CharacterStyle>;
  paragraphs: Map<string, ParagraphStyle>;
  borders: Map<string, BorderStyle>;
}

interface ImageResource {
  id: string;
  sourcePath: string;
  targetName: string;
  mediaType: string;
  relationshipId: string;
}

interface HeaderFooterPart {
  kind: 'header' | 'footer';
  type: 'default' | 'even' | 'first';
  relationshipId: string;
  targetName: string;
  element: Element;
}

interface SectionState {
  sectionProperties: string;
  headerFooters: Map<string, HeaderFooterPart>;
}

interface NotePart {
  kind: 'footnote' | 'endnote';
  id: number;
  element: Element;
}

interface RenderContext {
  catalog: StyleCatalog;
  images: Map<string, ImageResource>;
  warnings: Set<string>;
  nextDrawingId: number;
  numberingIds: Map<string, number>;
  footnotes: NotePart[];
  endnotes: NotePart[];
}

interface PackageContext extends RenderContext {
  zip: JSZip;
  output: JSZip;
  documentRelationships: string[];
  headerFooterParts: HeaderFooterPart[];
}

/**
 * Direct canonical HWPX -> WordprocessingML writer.
 *
 * This deliberately reads the editable HWPX package instead of rHWP's rendered
 * page HTML. That keeps semantic paragraphs, runs, tables, images and section
 * geometry available to Word and avoids baking pagination into HTML snapshots.
 */
export async function exportHwpxToDocx(hwpxBytes: Uint8Array): Promise<HwpxDocxExportResult> {
  if (typeof DOMParser === 'undefined') {
    throw new Error('DOCX 직접 저장에는 XML DOMParser를 사용할 수 있는 환경이 필요합니다.');
  }

  const zip = await JSZip.loadAsync(hwpxBytes);
  const headerXml = await readRequiredText(zip, ['Contents/header.xml', 'contents/header.xml']);
  const header = parseXml(headerXml, 'HWPX header.xml');
  const catalog = parseStyleCatalog(header);
  const images = await collectImages(zip);
  const output = new JSZip();
  const warnings = new Set<string>();

  const ctx: PackageContext = {
    zip,
    output,
    catalog,
    images,
    warnings,
    nextDrawingId: 1,
    numberingIds: new Map(),
    footnotes: [],
    endnotes: [],
    documentRelationships: [],
    headerFooterParts: [],
  };

  for (const image of images.values()) {
    if (image.mediaType === 'application/octet-stream') {
      warnings.add(`DOCX가 직접 표시하지 못할 수 있는 HWPX 이미지 형식(${image.sourcePath})을 원본 리소스로 포함했습니다.`);
    }
    const bytes = await zip.file(image.sourcePath)?.async('uint8array');
    if (!bytes) {
      warnings.add(`HWPX 이미지 리소스 ${image.sourcePath}를 찾지 못해 DOCX에서 제외했습니다.`);
      continue;
    }
    output.file(`word/media/${image.targetName}`, bytes);
    ctx.documentRelationships.push(relationshipXml(
      image.relationshipId,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      `media/${image.targetName}`,
    ));
  }

  const sectionPaths = Object.keys(zip.files)
    .filter((path) => /^contents\/section\d+\.xml$/i.test(path))
    .sort((a, b) => sectionIndex(a) - sectionIndex(b));
  if (!sectionPaths.length) throw new Error('HWPX에 section XML이 없어 DOCX로 저장할 수 없습니다.');

  const bodyBlocks: string[] = [];
  let currentSection: SectionState | undefined;
  let renderedAnyContent = false;

  for (const sectionPath of sectionPaths) {
    const sectionXml = await zip.file(sectionPath)!.async('text');
    const sectionDocument = parseXml(sectionXml, sectionPath);
    const root = sectionDocument.documentElement;
    detectUnsupportedHwpxFeatures(root, warnings);

    for (const paragraph of directChildren(root, 'p')) {
      const sectionElement = firstDescendant(paragraph, 'secPr');
      if (sectionElement) {
        const nextSection = parseSectionState(sectionElement, paragraph);
        if (currentSection && renderedAnyContent) {
          bodyBlocks.push(`<w:p><w:pPr>${renderSectionProperties(currentSection, ctx)}</w:pPr></w:p>`);
        }
        currentSection = nextSection;
      }

      if (!currentSection) currentSection = defaultSectionState();
      collectHeaderFooters(paragraph, currentSection, ctx);

      const blocks = renderTopLevelParagraph(paragraph, ctx);
      if (blocks.length) {
        bodyBlocks.push(...blocks);
        renderedAnyContent = true;
      }
    }
  }

  currentSection ??= defaultSectionState();
  const finalSectionXml = renderSectionProperties(currentSection, ctx);
  const bodyXml = bodyBlocks.join('') || '<w:p/>';

  ctx.documentRelationships.unshift(
    relationshipXml('rIdStyles', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'styles.xml'),
    relationshipXml('rIdSettings', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml'),
  );

  if (ctx.numberingIds.size) {
    ctx.documentRelationships.push(
      relationshipXml('rIdNumbering', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering', 'numbering.xml'),
    );
    output.file('word/numbering.xml', renderNumbering(ctx.numberingIds));
  }

  if (ctx.footnotes.length) {
    ctx.documentRelationships.push(relationshipXml(
      'rIdFootnotes',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
      'footnotes.xml',
    ));
    output.file('word/footnotes.xml', renderNotesPart('footnote', ctx.footnotes, ctx));
  }
  if (ctx.endnotes.length) {
    ctx.documentRelationships.push(relationshipXml(
      'rIdEndnotes',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
      'endnotes.xml',
    ));
    output.file('word/endnotes.xml', renderNotesPart('endnote', ctx.endnotes, ctx));
  }

  output.file('[Content_Types].xml', renderContentTypes(ctx));
  output.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG_REL_NS}">${relationshipXml(
    'rId1',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
    'word/document.xml',
  )}${relationshipXml(
    'rId2',
    'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
    'docProps/core.xml',
  )}${relationshipXml(
    'rId3',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
    'docProps/app.xml',
  )}</Relationships>`);
  output.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG_REL_NS}">${ctx.documentRelationships.join('')}</Relationships>`);
  output.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${WORD_NS}" xmlns:r="${REL_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:m="${MATH_NS}" xmlns:wps="${WPS_NS}"><w:body>${bodyXml}${finalSectionXml}</w:body></w:document>`);
  output.file('word/styles.xml', renderStyles(catalog));
  output.file('word/settings.xml', renderSettings(ctx));
  output.file('docProps/core.xml', renderCoreProperties());
  output.file('docProps/app.xml', renderAppProperties());

  for (const part of ctx.headerFooterParts) {
    const tag = part.kind === 'header' ? 'hdr' : 'ftr';
    const paragraphs = renderHeaderFooterParagraphs(part.element, ctx);
    output.file(`word/${part.targetName}`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${tag} xmlns:w="${WORD_NS}" xmlns:r="${REL_NS}" xmlns:m="${MATH_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="${WPS_NS}">${paragraphs || '<w:p/>'}</w:${tag}>`);
  }

  const bytes = await output.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  return { bytes, warnings: [...warnings] };
}

function parseStyleCatalog(header: XMLDocument): StyleCatalog {
  const fonts: FontCatalog = {
    hangul: new Map(), latin: new Map(), hanja: new Map(), japanese: new Map(),
    other: new Map(), symbol: new Map(), user: new Map(),
  };
  for (const face of descendants(header.documentElement, 'fontface')) {
    const lang = (face.getAttribute('lang') ?? '').toLowerCase() as keyof FontCatalog;
    const target = fonts[lang];
    if (!(target instanceof Map)) continue;
    for (const font of directChildren(face, 'font')) {
      const id = font.getAttribute('id');
      const name = font.getAttribute('face');
      if (id !== null && name) target.set(id, name);
    }
  }

  const characters = new Map<string, CharacterStyle>();
  for (const charPr of descendants(header.documentElement, 'charPr')) {
    const id = charPr.getAttribute('id');
    if (id === null) continue;
    characters.set(id, { id, xml: renderCharacterProperties(charPr, fonts) });
  }

  const paragraphs = new Map<string, ParagraphStyle>();
  for (const paraPr of descendants(header.documentElement, 'paraPr')) {
    const id = paraPr.getAttribute('id');
    if (id === null) continue;
    paragraphs.set(id, parseParagraphProperties(paraPr));
  }

  const borders = new Map<string, BorderStyle>();
  for (const borderFill of descendants(header.documentElement, 'borderFill')) {
    const id = borderFill.getAttribute('id');
    if (id === null) continue;
    borders.set(id, parseBorderFill(borderFill));
  }

  return { fonts, characters, paragraphs, borders };
}

function renderCharacterProperties(charPr: Element, fonts: FontCatalog): string {
  const chunks: string[] = [];
  const fontRef = firstDirectChild(charPr, 'fontRef');
  if (fontRef) {
    const eastAsia = fontName(fonts.hangul, fontRef.getAttribute('hangul'))
      ?? fontName(fonts.hanja, fontRef.getAttribute('hanja'));
    const latin = fontName(fonts.latin, fontRef.getAttribute('latin'))
      ?? eastAsia
      ?? fontName(fonts.other, fontRef.getAttribute('other'));
    const complex = fontName(fonts.other, fontRef.getAttribute('other')) ?? latin;
    if (eastAsia || latin || complex) {
      chunks.push(`<w:rFonts${latin ? ` w:ascii="${escapeXml(latin)}" w:hAnsi="${escapeXml(latin)}"` : ''}${eastAsia ? ` w:eastAsia="${escapeXml(eastAsia)}"` : ''}${complex ? ` w:cs="${escapeXml(complex)}"` : ''}/>`);
    }
  }

  const height = finiteNumber(charPr.getAttribute('height'));
  if (height && height > 0) {
    const halfPoints = Math.max(1, Math.round(height / 50));
    chunks.push(`<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/>`);
  }
  const textColor = normalizeHex(charPr.getAttribute('textColor'));
  if (textColor) chunks.push(`<w:color w:val="${textColor}"/>`);
  const shadeColor = normalizeHex(charPr.getAttribute('shadeColor'));
  if (shadeColor) chunks.push(`<w:shd w:val="clear" w:fill="${shadeColor}"/>`);
  if (firstDirectChild(charPr, 'bold')) chunks.push('<w:b/>');
  if (firstDirectChild(charPr, 'italic')) chunks.push('<w:i/>');

  const underline = firstDirectChild(charPr, 'underline');
  if (underline && underline.getAttribute('type') !== 'NONE') {
    const color = normalizeHex(underline.getAttribute('color'));
    chunks.push(`<w:u w:val="${mapUnderline(underline.getAttribute('shape'))}"${color ? ` w:color="${color}"` : ''}/>`);
  }
  const strike = firstDirectChild(charPr, 'strikeout');
  if (strike && strike.getAttribute('shape') !== 'NONE') chunks.push('<w:strike/>');
  if (firstDirectChild(charPr, 'supscript')) chunks.push('<w:vertAlign w:val="superscript"/>');
  if (firstDirectChild(charPr, 'subscript')) chunks.push('<w:vertAlign w:val="subscript"/>');
  if (firstDirectChild(charPr, 'emboss')) chunks.push('<w:emboss/>');
  if (firstDirectChild(charPr, 'engrave')) chunks.push('<w:imprint/>');
  const outline = firstDirectChild(charPr, 'outline');
  if (outline && outline.getAttribute('type') !== 'NONE') chunks.push('<w:outline/>');

  const ratio = firstDirectChild(charPr, 'ratio');
  const widthPercent = finiteNumber(ratio?.getAttribute('hangul') ?? ratio?.getAttribute('latin'));
  if (widthPercent && widthPercent !== 100) chunks.push(`<w:w w:val="${Math.max(1, Math.min(600, Math.round(widthPercent)))}"/>`);
  return chunks.join('');
}

function parseParagraphProperties(paraPr: Element): ParagraphStyle {
  const chunks: string[] = [];
  const align = firstDirectChild(paraPr, 'align')?.getAttribute('horizontal');
  if (align) chunks.push(`<w:jc w:val="${mapParagraphAlignment(align)}"/>`);

  const breakSetting = firstDirectChild(paraPr, 'breakSetting');
  if (breakSetting) {
    if (isOne(breakSetting.getAttribute('keepWithNext'))) chunks.push('<w:keepNext/>');
    if (isOne(breakSetting.getAttribute('keepLines'))) chunks.push('<w:keepLines/>');
    if (isOne(breakSetting.getAttribute('pageBreakBefore'))) chunks.push('<w:pageBreakBefore/>');
    const widow = breakSetting.getAttribute('widowOrphan');
    if (widow !== null) chunks.push(`<w:widowControl w:val="${isOne(widow) ? '1' : '0'}"/>`);
  }

  const switchElement = firstDirectChild(paraPr, 'switch');
  const metrics = switchElement
    ? firstDirectChild(switchElement, 'default') ?? firstDirectChild(switchElement, 'case')
    : paraPr;
  const margin = metrics ? firstDescendant(metrics, 'margin') : undefined;
  if (margin) {
    const left = hwpUnitValue(firstDirectChild(margin, 'left'));
    const right = hwpUnitValue(firstDirectChild(margin, 'right'));
    const indent = hwpUnitValue(firstDirectChild(margin, 'intent'));
    const attrs: string[] = [];
    if (left !== undefined) attrs.push(`w:left="${hwpUnitToTwip(left)}"`);
    if (right !== undefined) attrs.push(`w:right="${hwpUnitToTwip(right)}"`);
    if (indent !== undefined && indent !== 0) {
      attrs.push(indent > 0 ? `w:firstLine="${hwpUnitToTwip(indent)}"` : `w:hanging="${hwpUnitToTwip(-indent)}"`);
    }
    if (attrs.length) chunks.push(`<w:ind ${attrs.join(' ')}/>`);

    const before = hwpUnitValue(firstDirectChild(margin, 'prev'));
    const after = hwpUnitValue(firstDirectChild(margin, 'next'));
    const line = metrics ? firstDescendant(metrics, 'lineSpacing') : undefined;
    const spacingAttrs: string[] = [];
    if (before !== undefined) spacingAttrs.push(`w:before="${hwpUnitToTwip(before)}"`);
    if (after !== undefined) spacingAttrs.push(`w:after="${hwpUnitToTwip(after)}"`);
    if (line) {
      const value = finiteNumber(line.getAttribute('value'));
      const type = (line.getAttribute('type') ?? '').toUpperCase();
      if (value !== undefined && value > 0) {
        if (type === 'PERCENT') {
          spacingAttrs.push(`w:line="${Math.round(value * 2.4)}"`, 'w:lineRule="auto"');
        } else {
          spacingAttrs.push(`w:line="${hwpUnitToTwip(value)}"`, `w:lineRule="${type === 'MINIMUM' ? 'atLeast' : 'exact'}"`);
        }
      }
    }
    if (spacingAttrs.length) chunks.push(`<w:spacing ${spacingAttrs.join(' ')}/>`);
  }

  const heading = firstDirectChild(paraPr, 'heading');
  let numbering: ParagraphStyle['numbering'];
  if (heading) {
    const type = (heading.getAttribute('type') ?? '').toUpperCase();
    if (type === 'NUMBER' || type === 'BULLET') {
      const kind = type === 'NUMBER' ? 'number' : 'bullet';
      const id = heading.getAttribute('idRef') ?? '0';
      const level = Math.max(0, Math.min(8, Math.trunc(finiteNumber(heading.getAttribute('level')) ?? 0)));
      numbering = { kind, key: `${kind}:${id}`, level };
    }
  }
  return { id: paraPr.getAttribute('id') ?? '', xml: chunks.join(''), numbering };
}

function parseBorderFill(borderFill: Element): BorderStyle {
  const sides = [
    ['topBorder', 'top'], ['leftBorder', 'left'], ['bottomBorder', 'bottom'], ['rightBorder', 'right'],
  ] as const;
  const rendered: string[] = [];
  for (const [hwpName, wordName] of sides) {
    const border = firstDirectChild(borderFill, hwpName);
    if (!border) continue;
    rendered.push(renderWordBorder(wordName, border));
  }
  const bordersXml = rendered.length ? rendered.join('') : '';
  const fill = descendants(borderFill, 'winBrush')[0];
  const fillColor = normalizeHex(fill?.getAttribute('faceColor') ?? fill?.getAttribute('color'));
  return {
    id: borderFill.getAttribute('id') ?? '',
    tableBorders: bordersXml ? `<w:tblBorders>${bordersXml}</w:tblBorders>` : '',
    cellBorders: bordersXml ? `<w:tcBorders>${bordersXml}</w:tcBorders>` : '',
    shading: fillColor ? `<w:shd w:val="clear" w:fill="${fillColor}"/>` : undefined,
  };
}

function renderWordBorder(name: string, border: Element): string {
  const hwpType = (border.getAttribute('type') ?? 'NONE').toUpperCase();
  const color = normalizeHex(border.getAttribute('color')) ?? '000000';
  const width = borderWidthEighthPoints(border.getAttribute('width'));
  return `<w:${name} w:val="${mapBorderType(hwpType)}" w:sz="${width}" w:space="0" w:color="${color}"/>`;
}

function parseSectionState(secPr: Element, paragraph: Element): SectionState {
  const pagePr = firstDescendant(secPr, 'pagePr');
  const margin = pagePr ? firstDirectChild(pagePr, 'margin') : undefined;
  const chunks: string[] = [];
  if (pagePr) {
    const width = finiteNumber(pagePr.getAttribute('width'));
    const height = finiteNumber(pagePr.getAttribute('height'));
    const orientation = pagePr.getAttribute('landscape') === 'NARROWLY' ? 'landscape' : undefined;
    if (width && height) chunks.push(`<w:pgSz w:w="${hwpUnitToTwip(width)}" w:h="${hwpUnitToTwip(height)}"${orientation ? ` w:orient="${orientation}"` : ''}/>`);
  }
  if (margin) {
    chunks.push(`<w:pgMar w:top="${hwpUnitToTwip(finiteNumber(margin.getAttribute('top')) ?? 0)}" w:right="${hwpUnitToTwip(finiteNumber(margin.getAttribute('right')) ?? 0)}" w:bottom="${hwpUnitToTwip(finiteNumber(margin.getAttribute('bottom')) ?? 0)}" w:left="${hwpUnitToTwip(finiteNumber(margin.getAttribute('left')) ?? 0)}" w:header="${hwpUnitToTwip(finiteNumber(margin.getAttribute('header')) ?? 0)}" w:footer="${hwpUnitToTwip(finiteNumber(margin.getAttribute('footer')) ?? 0)}" w:gutter="${hwpUnitToTwip(finiteNumber(margin.getAttribute('gutter')) ?? 0)}"/>`);
  }
  const colPr = firstDescendant(paragraph, 'colPr');
  if (colPr) {
    const count = Math.max(1, Math.trunc(finiteNumber(colPr.getAttribute('colCount')) ?? 1));
    const gap = finiteNumber(colPr.getAttribute('sameGap'));
    chunks.push(`<w:cols w:num="${count}"${gap !== undefined ? ` w:space="${hwpUnitToTwip(gap)}"` : ''}/>`);
  }
  return { sectionProperties: chunks.join(''), headerFooters: new Map() };
}

function defaultSectionState(): SectionState {
  return {
    sectionProperties: '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1701" w:bottom="850" w:left="1701" w:header="850" w:footer="850" w:gutter="0"/>',
    headerFooters: new Map(),
  };
}

function collectHeaderFooters(paragraph: Element, section: SectionState, ctx: PackageContext): void {
  for (const kind of ['header', 'footer'] as const) {
    for (const element of descendants(paragraph, kind)) {
      const type = mapHeaderFooterType(element.getAttribute('applyPageType'));
      const key = `${kind}:${type}`;
      if (section.headerFooters.has(key)) continue;
      const index = ctx.headerFooterParts.filter((part) => part.kind === kind).length + 1;
      const targetName = `${kind}${index}.xml`;
      const relationshipId = `rId${kind === 'header' ? 'Header' : 'Footer'}${index}`;
      const part: HeaderFooterPart = { kind, type, relationshipId, targetName, element };
      section.headerFooters.set(key, part);
      ctx.headerFooterParts.push(part);
      ctx.documentRelationships.push(relationshipXml(
        relationshipId,
        `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}`,
        targetName,
      ));
    }
  }
}

function renderSectionProperties(section: SectionState, ctx: RenderContext): string {
  const refs = [...section.headerFooters.values()].map((part) => (
    `<w:${part.kind}Reference w:type="${part.type}" r:id="${part.relationshipId}"/>`
  )).join('');
  if ([...section.headerFooters.values()].some((part) => part.type === 'first')) {
    return `<w:sectPr>${refs}<w:titlePg/>${section.sectionProperties}</w:sectPr>`;
  }
  return `<w:sectPr>${refs}${section.sectionProperties}</w:sectPr>`;
}

function renderTopLevelParagraph(paragraph: Element, ctx: RenderContext): string[] {
  const tables = directRunChildren(paragraph, 'tbl');
  if (!tables.length) return [renderParagraph(paragraph, ctx)];

  const blocks: string[] = [];
  const hasNonTableContent = directRuns(paragraph).some((run) =>
    directChildren(run, 't').some((text) => meaningfulText(text.textContent ?? ''))
    || directChildren(run, 'pic').length > 0
    || directChildren(run, 'equation').length > 0
    || Array.from(run.children).some((child) => isSupportedShapeName(child.localName))
    || directChildren(run, 'ctrl').some((control) => Boolean(firstDirectChild(control, 'footNote') || firstDirectChild(control, 'endNote')))
  );
  if (hasNonTableContent) blocks.push(renderParagraph(paragraph, ctx, true));
  for (const table of tables) blocks.push(renderTable(table, ctx));
  return blocks.length ? blocks : ['<w:p/>'];
}

function renderParagraph(paragraph: Element, ctx: RenderContext, skipTables = false, prefixRuns = ''): string {
  const pPrParts: string[] = [];
  const paraStyle = ctx.catalog.paragraphs.get(paragraph.getAttribute('paraPrIDRef') ?? '');
  if (paraStyle?.xml) pPrParts.push(paraStyle.xml);
  if (paraStyle?.numbering) {
    const numId = ensureNumberingId(ctx, paraStyle.numbering.key);
    pPrParts.push(`<w:numPr><w:ilvl w:val="${paraStyle.numbering.level}"/><w:numId w:val="${numId}"/></w:numPr>`);
  }
  if (paragraph.getAttribute('pageBreak') === '1') pPrParts.push('<w:pageBreakBefore/>');

  const runs: string[] = prefixRuns ? [prefixRuns] : [];
  for (const run of directRuns(paragraph)) {
    const charStyle = ctx.catalog.characters.get(run.getAttribute('charPrIDRef') ?? '');
    const rPr = charStyle?.xml ? `<w:rPr>${charStyle.xml}</w:rPr>` : '';
    const hasPicture = directChildren(run, 'pic').length > 0;
    for (const child of Array.from(run.children)) {
      switch (child.localName) {
        case 't': {
          const value = child.textContent ?? '';
          if (hasPicture && /^\[이미지\]$/.test(value.trim())) break;
          runs.push(renderTextRun(child, rPr));
          break;
        }
        case 'pic':
          runs.push(renderPictureRun(child, rPr, ctx));
          break;
        case 'equation':
          runs.push(renderEquation(child, ctx));
          break;
        case 'ctrl': {
          const rendered = renderRunControl(child, ctx);
          if (rendered) runs.push(rendered);
          break;
        }
        case 'tbl':
          if (!skipTables) ctx.warnings.add('한 문단 안의 표는 DOCX 블록 표로 정규화했습니다.');
          break;
        default:
          if (isSupportedShapeName(child.localName)) runs.push(renderShapeRun(child, rPr, ctx));
          break;
      }
    }
  }

  if (paragraph.getAttribute('columnBreak') === '1') runs.push('<w:r><w:br w:type="column"/></w:r>');
  return `<w:p>${pPrParts.length ? `<w:pPr>${pPrParts.join('')}</w:pPr>` : ''}${runs.join('')}</w:p>`;
}

function renderRunControl(control: Element, ctx: RenderContext): string {
  const footnote = firstDirectChild(control, 'footNote');
  if (footnote) return renderNoteReference('footnote', footnote, ctx);
  const endnote = firstDirectChild(control, 'endNote');
  if (endnote) return renderNoteReference('endnote', endnote, ctx);
  return '';
}

function renderNoteReference(kind: NotePart['kind'], element: Element, ctx: RenderContext): string {
  const collection = kind === 'footnote' ? ctx.footnotes : ctx.endnotes;
  const id = collection.length + 1;
  collection.push({ kind, id, element });
  const style = kind === 'footnote' ? 'FootnoteReference' : 'EndnoteReference';
  return `<w:r><w:rPr><w:rStyle w:val="${style}"/></w:rPr><w:${kind}Reference w:id="${id}"/></w:r>`;
}

function renderTextRun(textElement: Element, rPr: string): string {
  const pieces: string[] = [];
  for (const node of Array.from(textElement.childNodes)) {
    if (node.nodeType === 3) {
      if (node.nodeValue) pieces.push(`<w:t xml:space="preserve">${escapeXml(node.nodeValue)}</w:t>`);
      continue;
    }
    if (node.nodeType !== 1) continue;
    const element = node as Element;
    const name = element.localName.toLowerCase();
    if (name.includes('tab')) pieces.push('<w:tab/>');
    else if (name.includes('linebreak') || name === 'br') pieces.push('<w:br/>');
    else if (element.textContent) pieces.push(`<w:t xml:space="preserve">${escapeXml(element.textContent)}</w:t>`);
  }
  if (!pieces.length && textElement.textContent) pieces.push(`<w:t xml:space="preserve">${escapeXml(textElement.textContent)}</w:t>`);
  return pieces.length ? `<w:r>${rPr}${pieces.join('')}</w:r>` : '';
}

function renderPictureRun(picture: Element, rPr: string, ctx: RenderContext): string {
  const image = firstDescendant(picture, 'img');
  const binaryId = image?.getAttribute('binaryItemIDRef');
  const resource = binaryId ? ctx.images.get(binaryId) : undefined;
  if (!resource) {
    ctx.warnings.add('일부 HWPX 그림의 binary resource를 찾지 못해 DOCX에서 제외했습니다.');
    return '';
  }

  const currentSize = firstDescendant(picture, 'curSz') ?? firstDescendant(picture, 'sz');
  const widthHu = finiteNumber(currentSize?.getAttribute('width')) ?? 7200;
  const heightHu = finiteNumber(currentSize?.getAttribute('height')) ?? 7200;
  const cx = Math.max(1, Math.round(widthHu * 127));
  const cy = Math.max(1, Math.round(heightHu * 127));
  const drawingId = ctx.nextDrawingId++;
  const position = firstDescendant(picture, 'pos');
  const graphic = `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${escapeXml(resource.targetName)}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${resource.relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
  const drawing = position && position.getAttribute('treatAsChar') !== '1'
    ? renderAnchoredDrawing(picture, position, cx, cy, drawingId, graphic, ctx)
    : renderInlineDrawing(picture, cx, cy, drawingId, graphic);
  return `<w:r>${rPr}${drawing}</w:r>`;
}

function renderShapeRun(shape: Element, rPr: string, ctx: RenderContext): string {
  const currentSize = firstDescendant(shape, 'curSz') ?? firstDescendant(shape, 'sz');
  const widthHu = finiteNumber(currentSize?.getAttribute('width')) ?? 7200;
  const heightHu = finiteNumber(currentSize?.getAttribute('height')) ?? 3600;
  const cx = Math.max(1, Math.round(widthHu * 127));
  const cy = Math.max(1, Math.round(heightHu * 127));
  const drawingId = ctx.nextDrawingId++;
  const preset = mapShapePreset(shape);
  const graphic = renderShapeGraphic(shape, preset, cx, cy, ctx);
  const position = firstDescendant(shape, 'pos');
  const description = firstDirectChild(shape, 'shapeComment')?.textContent?.trim() || undefined;
  const meta = { name: `Shape ${drawingId}`, description };
  const drawing = position && position.getAttribute('treatAsChar') !== '1'
    ? renderAnchoredDrawing(shape, position, cx, cy, drawingId, graphic, ctx, meta)
    : renderInlineDrawing(shape, cx, cy, drawingId, graphic, meta);
  return `<w:r>${rPr}${drawing}</w:r>`;
}

function renderShapeGraphic(
  shape: Element,
  preset: string,
  cx: number,
  cy: number,
  ctx: RenderContext,
): string {
  const rotation = firstDirectChild(shape, 'rotationInfo');
  const flip = firstDirectChild(shape, 'flip');
  const transformAttrs: string[] = [];
  const angle = finiteNumber(rotation?.getAttribute('angle'));
  if (angle) transformAttrs.push(`rot="${Math.round(angle * 60000)}"`);
  if (isOne(flip?.getAttribute('horizontal') ?? null)) transformAttrs.push('flipH="1"');
  if (isOne(flip?.getAttribute('vertical') ?? null)) transformAttrs.push('flipV="1"');

  const fill = renderShapeFill(shape, preset);
  const line = renderShapeLine(shape);
  const subList = firstDirectChild(shape, 'subList');
  const textBox = subList
    ? `<wps:txbx><w:txbxContent>${directChildren(subList, 'p').map((paragraph) => renderParagraph(paragraph, ctx)).join('') || '<w:p/>'}</w:txbxContent></wps:txbx>`
    : '';
  const bodyPr = subList ? renderShapeBodyProperties(shape, subList) : '<wps:bodyPr/>';
  return `<a:graphic><a:graphicData uri="${WPS_NS}"><wps:wsp><wps:cNvSpPr${subList ? ' txBox="1"' : ''}/><wps:spPr><a:xfrm${transformAttrs.length ? ` ${transformAttrs.join(' ')}` : ''}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>${fill}${line}</wps:spPr>${textBox}${bodyPr}</wps:wsp></a:graphicData></a:graphic>`;
}

function renderShapeFill(shape: Element, preset: string): string {
  if (preset === 'line') return '<a:noFill/>';
  const fillBrush = firstDirectChild(shape, 'fillBrush');
  const winBrush = fillBrush ? firstDescendant(fillBrush, 'winBrush') : undefined;
  const color = normalizeHex(winBrush?.getAttribute('faceColor') ?? winBrush?.getAttribute('color'));
  return color ? `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` : '<a:noFill/>';
}

function renderShapeLine(shape: Element): string {
  const line = firstDirectChild(shape, 'lineShape');
  if (!line) return '<a:ln><a:noFill/></a:ln>';
  const width = Math.max(1, Math.round((finiteNumber(line.getAttribute('width')) ?? 33) * 127));
  const color = normalizeHex(line.getAttribute('color')) ?? '000000';
  const dash = mapDrawingLineDash(line.getAttribute('style'));
  const head = mapDrawingArrow(line.getAttribute('headStyle'));
  const tail = mapDrawingArrow(line.getAttribute('tailStyle'));
  return `<a:ln w="${width}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:prstDash val="${dash}"/>${head ? `<a:headEnd type="${head}"/>` : ''}${tail ? `<a:tailEnd type="${tail}"/>` : ''}</a:ln>`;
}

function renderShapeBodyProperties(shape: Element, subList: Element): string {
  const margin = firstDirectChild(shape, 'inMargin');
  const vertical = (subList.getAttribute('vertAlign') ?? 'TOP').toUpperCase();
  const anchor = vertical === 'CENTER' ? 'ctr' : vertical === 'BOTTOM' ? 'b' : 't';
  const attrs = [
    `anchor="${anchor}"`,
    `lIns="${Math.max(0, Math.round((finiteNumber(margin?.getAttribute('left')) ?? 0) * 127))}"`,
    `rIns="${Math.max(0, Math.round((finiteNumber(margin?.getAttribute('right')) ?? 0) * 127))}"`,
    `tIns="${Math.max(0, Math.round((finiteNumber(margin?.getAttribute('top')) ?? 0) * 127))}"`,
    `bIns="${Math.max(0, Math.round((finiteNumber(margin?.getAttribute('bottom')) ?? 0) * 127))}"`,
  ];
  return `<wps:bodyPr ${attrs.join(' ')}/>`;
}

function mapShapePreset(shape: Element): string {
  switch (shape.localName) {
    case 'ellipse': return 'ellipse';
    case 'line':
    case 'connectLine': return 'line';
    case 'rect': {
      const roundRate = finiteNumber(shape.getAttribute('ratio')) ?? 0;
      return roundRate > 0 ? 'roundRect' : 'rect';
    }
    default: return 'rect';
  }
}

function isSupportedShapeName(localName: string): boolean {
  return localName === 'rect' || localName === 'ellipse' || localName === 'line' || localName === 'connectLine';
}

function mapDrawingLineDash(value: string | null): string {
  const normalized = (value ?? 'SOLID').toUpperCase();
  if (normalized.includes('DASH_DOT_DOT')) return 'lgDashDotDot';
  if (normalized.includes('DASH_DOT')) return 'lgDashDot';
  if (normalized.includes('DASH')) return 'dash';
  if (normalized.includes('DOT')) return 'dot';
  return 'solid';
}

function mapDrawingArrow(value: string | null): string | undefined {
  const normalized = (value ?? 'NORMAL').toUpperCase();
  if (normalized === 'NORMAL' || normalized === 'NONE') return undefined;
  if (normalized.includes('DIAMOND')) return 'diamond';
  if (normalized.includes('CIRCLE') || normalized.includes('OVAL')) return 'oval';
  if (normalized.includes('STEALTH')) return 'stealth';
  return 'triangle';
}

function renderInlineDrawing(
  object: Element,
  cx: number,
  cy: number,
  drawingId: number,
  graphic: string,
  metadata: { name: string; description?: string } = { name: `Picture ${drawingId}` },
): string {
  const margin = objectMarginEmu(object);
  return `<w:drawing><wp:inline distT="${margin.top}" distB="${margin.bottom}" distL="${margin.left}" distR="${margin.right}"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${drawingId}" name="${escapeXml(metadata.name)}"${metadata.description ? ` descr="${escapeXml(metadata.description)}"` : ''}/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>${graphic}</wp:inline></w:drawing>`;
}

function renderAnchoredDrawing(
  object: Element,
  position: Element,
  cx: number,
  cy: number,
  drawingId: number,
  graphic: string,
  ctx: RenderContext,
  metadata: { name: string; description?: string } = { name: `Picture ${drawingId}` },
): string {
  const margin = objectMarginEmu(object);
  const wrap = renderAnchorWrap(object, ctx);
  const textWrap = (object.getAttribute('textWrap') ?? '').toUpperCase();
  const behindDoc = textWrap === 'BEHIND_TEXT' ? '1' : '0';
  const zOrder = Math.max(0, Math.trunc(finiteNumber(object.getAttribute('zOrder')) ?? 0));
  const allowOverlap = isOne(position.getAttribute('allowOverlap')) ? '1' : '0';
  const lock = isOne(object.getAttribute('lock')) ? '1' : '0';
  return `<w:drawing><wp:anchor distT="${margin.top}" distB="${margin.bottom}" distL="${margin.left}" distR="${margin.right}" simplePos="0" relativeHeight="${Math.max(1, zOrder + 1)}" behindDoc="${behindDoc}" locked="${lock}" layoutInCell="1" allowOverlap="${allowOverlap}"><wp:simplePos x="0" y="0"/>${renderAnchorPosition(position, 'horizontal', ctx)}${renderAnchorPosition(position, 'vertical', ctx)}<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrap}<wp:docPr id="${drawingId}" name="${escapeXml(metadata.name)}"${metadata.description ? ` descr="${escapeXml(metadata.description)}"` : ''}/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>${graphic}</wp:anchor></w:drawing>`;
}

function renderAnchorPosition(
  position: Element,
  axis: 'horizontal' | 'vertical',
  ctx: RenderContext,
): string {
  const horizontal = axis === 'horizontal';
  const relation = position.getAttribute(horizontal ? 'horzRelTo' : 'vertRelTo');
  const align = position.getAttribute(horizontal ? 'horzAlign' : 'vertAlign');
  const offsetHu = finiteNumber(position.getAttribute(horizontal ? 'horzOffset' : 'vertOffset')) ?? 0;
  const tag = horizontal ? 'positionH' : 'positionV';
  const relativeFrom = horizontal
    ? mapHorizontalRelativeFrom(relation, ctx)
    : mapVerticalRelativeFrom(relation, ctx);
  const mappedAlign = horizontal ? mapHorizontalObjectAlign(align) : mapVerticalObjectAlign(align);
  if (offsetHu !== 0 || !mappedAlign) {
    return `<wp:${tag} relativeFrom="${relativeFrom}"><wp:posOffset>${Math.round(offsetHu * 127)}</wp:posOffset></wp:${tag}>`;
  }
  return `<wp:${tag} relativeFrom="${relativeFrom}"><wp:align>${mappedAlign}</wp:align></wp:${tag}>`;
}

function renderAnchorWrap(object: Element, ctx: RenderContext): string {
  const wrap = (object.getAttribute('textWrap') ?? 'SQUARE').toUpperCase();
  const flow = (object.getAttribute('textFlow') ?? 'BOTH_SIDES').toUpperCase();
  if (wrap === 'TOP_AND_BOTTOM') return '<wp:wrapTopAndBottom/>';
  if (wrap === 'BEHIND_TEXT' || wrap === 'IN_FRONT_OF_TEXT' || wrap === 'NONE') return '<wp:wrapNone/>';
  if (wrap === 'TIGHT' || wrap === 'THROUGH') {
    ctx.warnings.add(`HWPX 그림의 ${wrap} 외곽선 감싸기는 DOCX square wrapping으로 근사됩니다.`);
  }
  return `<wp:wrapSquare wrapText="${mapWrapSide(flow)}"/>`;
}

function objectMarginEmu(object: Element): { top: number; right: number; bottom: number; left: number } {
  const margin = firstDirectChild(object, 'outMargin');
  return {
    top: Math.max(0, Math.round((finiteNumber(margin?.getAttribute('top')) ?? 0) * 127)),
    right: Math.max(0, Math.round((finiteNumber(margin?.getAttribute('right')) ?? 0) * 127)),
    bottom: Math.max(0, Math.round((finiteNumber(margin?.getAttribute('bottom')) ?? 0) * 127)),
    left: Math.max(0, Math.round((finiteNumber(margin?.getAttribute('left')) ?? 0) * 127)),
  };
}

function renderEquation(equation: Element, ctx: RenderContext): string {
  const script = firstDirectChild(equation, 'script')?.textContent?.trim();
  if (!script) return '';
  const style = renderMathRunStyle(equation);
  const rendered = renderMathExpression(script, style);
  if (hasAdvancedEquationSyntax(script)) {
    ctx.warnings.add('복잡한 HWP 수식 문법 일부는 편집 가능한 OMML 수식 텍스트로 보존됩니다.');
  }
  return `<m:oMath>${rendered}</m:oMath>`;
}

function renderMathRunStyle(equation: Element): string {
  const props: string[] = [];
  const fontSize = finiteNumber(equation.getAttribute('baseUnit'));
  if (fontSize && fontSize > 0) props.push(`<w:sz w:val="${Math.max(1, Math.round(fontSize / 50))}"/>`);
  const color = normalizeHex(equation.getAttribute('textColor'));
  if (color) props.push(`<w:color w:val="${color}"/>`);
  return props.length ? `<w:rPr>${props.join('')}</w:rPr>` : '';
}

function renderMathExpression(source: string, style: string): string {
  const value = stripOuterMathBraces(source.trim());
  const over = findTopLevelKeyword(value, 'over');
  if (over >= 0) {
    const left = value.slice(0, over).trim();
    const right = value.slice(over + 4).trim();
    return `<m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num>${renderMathExpression(left, style)}</m:num><m:den>${renderMathExpression(right, style)}</m:den></m:f>`;
  }
  return renderMathSequence(value, style);
}

function renderMathSequence(source: string, style: string): string {
  const chunks: string[] = [];
  let index = 0;
  while (index < source.length) {
    if (/\s/.test(source[index])) {
      index += 1;
      continue;
    }

    const word = readMathWord(source, index);
    if (word) {
      const keyword = word.value.toLowerCase();
      if (keyword === 'sqrt') {
        const atom = readMathAtom(source, word.end);
        if (atom) {
          chunks.push(`<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${renderMathExpression(atom.value, style)}</m:e></m:rad>`);
          index = atom.end;
          continue;
        }
      }
      if (keyword === 'sum' || keyword === 'prod' || keyword === 'int') {
        const nary = renderMathNary(source, word.end, keyword, style);
        chunks.push(nary.xml);
        index = nary.end;
        continue;
      }
    }

    const atom = readMathAtom(source, index);
    if (!atom) {
      chunks.push(renderMathText(source[index], style));
      index += 1;
      continue;
    }
    index = atom.end;
    let base = renderMathAtomValue(atom.value, style);
    let sub: string | undefined;
    let sup: string | undefined;
    for (;;) {
      index = skipMathSpaces(source, index);
      const marker = source[index];
      if (marker !== '^' && marker !== '_') break;
      const script = readMathAtom(source, index + 1);
      if (!script) break;
      if (marker === '^') sup = renderMathExpression(script.value, style);
      else sub = renderMathExpression(script.value, style);
      index = script.end;
    }
    if (sub && sup) base = `<m:sSubSup><m:e>${base}</m:e><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`;
    else if (sub) base = `<m:sSub><m:e>${base}</m:e><m:sub>${sub}</m:sub></m:sSub>`;
    else if (sup) base = `<m:sSup><m:e>${base}</m:e><m:sup>${sup}</m:sup></m:sSup>`;
    chunks.push(base);
  }
  return chunks.join('');
}

function renderMathNary(
  source: string,
  start: number,
  keyword: 'sum' | 'prod' | 'int',
  style: string,
): { xml: string; end: number } {
  let index = skipMathSpaces(source, start);
  let sub = '';
  let sup = '';
  const from = readMathWord(source, index);
  if (from?.value.toLowerCase() === 'from') {
    const atom = readMathAtom(source, from.end);
    if (atom) {
      sub = renderMathExpression(atom.value, style);
      index = skipMathSpaces(source, atom.end);
    }
  }
  const to = readMathWord(source, index);
  if (to?.value.toLowerCase() === 'to') {
    const atom = readMathAtom(source, to.end);
    if (atom) {
      sup = renderMathExpression(atom.value, style);
      index = skipMathSpaces(source, atom.end);
    }
  }
  const body = readMathAtom(source, index);
  const bodyXml = body ? renderMathExpression(body.value, style) : renderMathText('', style);
  if (body) index = body.end;
  const symbol = keyword === 'sum' ? '∑' : keyword === 'prod' ? '∏' : '∫';
  const pr = `<m:naryPr><m:chr m:val="${symbol}"/>${sub ? '' : '<m:subHide m:val="1"/>'}${sup ? '' : '<m:supHide m:val="1"/>'}</m:naryPr>`;
  return {
    xml: `<m:nary>${pr}<m:sub>${sub}</m:sub><m:sup>${sup}</m:sup><m:e>${bodyXml}</m:e></m:nary>`,
    end: index,
  };
}

function renderMathAtomValue(value: string, style: string): string {
  const stripped = stripOuterMathBraces(value);
  if (stripped !== value || /\s/.test(stripped)) return renderMathExpression(stripped, style);
  const mapped = MATH_SYMBOLS[stripped.toLowerCase()] ?? stripped;
  return renderMathText(mapped, style);
}

function renderMathText(value: string, style: string): string {
  return `<m:r>${style}<m:t xml:space="preserve">${escapeXml(value)}</m:t></m:r>`;
}

const MATH_SYMBOLS: Record<string, string> = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', theta: 'θ', lambda: 'λ',
  mu: 'μ', pi: 'π', rho: 'ρ', sigma: 'σ', tau: 'τ', phi: 'φ', psi: 'ψ', omega: 'ω',
  infinity: '∞', inf: '∞', times: '×', cdot: '·', pm: '±', leq: '≤', geq: '≥', neq: '≠',
};

function readMathWord(source: string, start: number): { value: string; end: number } | undefined {
  const index = skipMathSpaces(source, start);
  const match = source.slice(index).match(/^[A-Za-z]+/);
  return match ? { value: match[0], end: index + match[0].length } : undefined;
}

function readMathAtom(source: string, start: number): { value: string; end: number } | undefined {
  let index = skipMathSpaces(source, start);
  if (index >= source.length) return undefined;
  if (source[index] === '{') {
    const end = findMatchingMathBrace(source, index);
    if (end < 0) return { value: source.slice(index + 1), end: source.length };
    return { value: source.slice(index + 1, end), end: end + 1 };
  }
  if ('()+-=*/,[]<>'.includes(source[index])) return { value: source[index], end: index + 1 };
  const begin = index;
  while (index < source.length && !/\s/.test(source[index]) && !'{}^_()+-=*/,[]<>'.includes(source[index])) index += 1;
  if (index === begin) return { value: source[index], end: index + 1 };
  return { value: source.slice(begin, index), end: index };
}

function skipMathSpaces(source: string, start: number): number {
  let index = start;
  while (index < source.length && /\s/.test(source[index])) index += 1;
  return index;
}

function findMatchingMathBrace(source: string, start: number): number {
  let depth = 0;
  for (let index = start; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    else if (source[index] === '}') {
      depth -= 1;
      if (depth === 0) return index;
    }
  }
  return -1;
}

function stripOuterMathBraces(value: string): string {
  let result = value.trim();
  while (result.startsWith('{') && findMatchingMathBrace(result, 0) === result.length - 1) {
    result = result.slice(1, -1).trim();
  }
  return result;
}

function findTopLevelKeyword(source: string, keyword: string): number {
  let depth = 0;
  for (let index = 0; index <= source.length - keyword.length; index += 1) {
    if (source[index] === '{') depth += 1;
    else if (source[index] === '}') depth = Math.max(0, depth - 1);
    if (depth !== 0) continue;
    if (source.slice(index, index + keyword.length).toLowerCase() !== keyword) continue;
    const before = index === 0 ? ' ' : source[index - 1];
    const after = index + keyword.length >= source.length ? ' ' : source[index + keyword.length];
    if (/\s|\}/.test(before) && /\s|\{/.test(after)) return index;
  }
  return -1;
}

function hasAdvancedEquationSyntax(script: string): boolean {
  return /\b(matrix|cases|pile|atop|left|right|phantom|accent|vec|hat|bar|under|overline|color|font|size)\b/i.test(script);
}

function renderTable(table: Element, ctx: RenderContext): string {
  const rows = directChildren(table, 'tr');
  const rowCount = Math.max(rows.length, Math.trunc(finiteNumber(table.getAttribute('rowCnt')) ?? rows.length));
  const colCount = Math.max(1, Math.trunc(finiteNumber(table.getAttribute('colCnt')) ?? inferColumnCount(rows)));
  const gridWidths = inferGridWidths(rows, colCount);
  const tableSize = firstDirectChild(table, 'sz');
  const declaredWidth = finiteNumber(tableSize?.getAttribute('width'));
  const width = declaredWidth && declaredWidth > 0 ? hwpUnitToTwip(declaredWidth) : gridWidths.reduce((sum, item) => sum + item, 0);
  const border = ctx.catalog.borders.get(table.getAttribute('borderFillIDRef') ?? '');
  const position = firstDirectChild(table, 'pos');
  const floatingProperties = position && position.getAttribute('treatAsChar') !== '1'
    ? renderFloatingTableProperties(table, position, ctx)
    : '';

  const origins = new Map<string, Element>();
  const covering = new Map<string, { origin: Element; row: number; col: number; colSpan: number; rowSpan: number }>();
  rows.forEach((row, fallbackRow) => {
    for (const cell of directChildren(row, 'tc')) {
      const address = firstDirectChild(cell, 'cellAddr');
      const span = firstDirectChild(cell, 'cellSpan');
      const rowAddr = Math.max(0, Math.trunc(finiteNumber(address?.getAttribute('rowAddr')) ?? fallbackRow));
      const colAddr = Math.max(0, Math.trunc(finiteNumber(address?.getAttribute('colAddr')) ?? 0));
      const colSpan = Math.max(1, Math.trunc(finiteNumber(span?.getAttribute('colSpan')) ?? 1));
      const rowSpan = Math.max(1, Math.trunc(finiteNumber(span?.getAttribute('rowSpan')) ?? 1));
      origins.set(`${rowAddr}:${colAddr}`, cell);
      for (let rr = rowAddr; rr < Math.min(rowCount, rowAddr + rowSpan); rr += 1) {
        for (let cc = colAddr; cc < Math.min(colCount, colAddr + colSpan); cc += 1) {
          covering.set(`${rr}:${cc}`, { origin: cell, row: rowAddr, col: colAddr, colSpan, rowSpan });
        }
      }
    }
  });

  const rowXml: string[] = [];
  for (let row = 0; row < rowCount; row += 1) {
    const cells: string[] = [];
    let col = 0;
    while (col < colCount) {
      const origin = origins.get(`${row}:${col}`);
      if (origin) {
        const cover = covering.get(`${row}:${col}`)!;
        cells.push(renderTableCell(origin, ctx, gridWidths, cover.col, cover.colSpan, cover.rowSpan > 1 ? 'restart' : undefined));
        col += cover.colSpan;
        continue;
      }
      const cover = covering.get(`${row}:${col}`);
      if (cover && cover.row < row && cover.col === col) {
        cells.push(renderTableCell(undefined, ctx, gridWidths, cover.col, cover.colSpan, 'continue'));
        col += cover.colSpan;
        continue;
      }
      if (cover) {
        col += 1;
        continue;
      }
      cells.push(renderTableCell(undefined, ctx, gridWidths, col, 1));
      col += 1;
    }
    rowXml.push(`<w:tr>${cells.join('')}</w:tr>`);
  }

  return `<w:tbl><w:tblPr><w:tblW w:w="${Math.max(0, width)}" w:type="dxa"/><w:tblLayout w:type="fixed"/>${floatingProperties}${border?.tableBorders ?? ''}</w:tblPr><w:tblGrid>${gridWidths.map((item) => `<w:gridCol w:w="${item}"/>`).join('')}</w:tblGrid>${rowXml.join('')}</w:tbl>`;
}

function renderFloatingTableProperties(table: Element, position: Element, ctx: RenderContext): string {
  const margin = firstDirectChild(table, 'outMargin');
  const attrs: string[] = [
    `w:horzAnchor="${mapTableHorizontalAnchor(position.getAttribute('horzRelTo'), ctx)}"`,
    `w:vertAnchor="${mapTableVerticalAnchor(position.getAttribute('vertRelTo'), ctx)}"`,
    `w:leftFromText="${hwpUnitToTwip(finiteNumber(margin?.getAttribute('left')) ?? 0)}"`,
    `w:rightFromText="${hwpUnitToTwip(finiteNumber(margin?.getAttribute('right')) ?? 0)}"`,
    `w:topFromText="${hwpUnitToTwip(finiteNumber(margin?.getAttribute('top')) ?? 0)}"`,
    `w:bottomFromText="${hwpUnitToTwip(finiteNumber(margin?.getAttribute('bottom')) ?? 0)}"`,
  ];
  const x = finiteNumber(position.getAttribute('horzOffset')) ?? 0;
  const y = finiteNumber(position.getAttribute('vertOffset')) ?? 0;
  const xSpec = mapHorizontalObjectAlign(position.getAttribute('horzAlign'));
  const ySpec = mapVerticalObjectAlign(position.getAttribute('vertAlign'));
  if (x !== 0 || !xSpec) attrs.push(`w:tblpX="${hwpUnitToTwip(x)}"`);
  else attrs.push(`w:tblpXSpec="${xSpec}"`);
  if (y !== 0 || !ySpec) attrs.push(`w:tblpY="${hwpUnitToTwip(y)}"`);
  else attrs.push(`w:tblpYSpec="${ySpec}"`);
  const overlap = isOne(position.getAttribute('allowOverlap')) ? 'overlap' : 'never';
  return `<w:tblpPr ${attrs.join(' ')}/><w:tblOverlap w:val="${overlap}"/>`;
}

function renderTableCell(
  cell: Element | undefined,
  ctx: RenderContext,
  gridWidths: number[],
  col: number,
  colSpan: number,
  vMerge?: 'restart' | 'continue',
): string {
  const size = cell ? firstDirectChild(cell, 'cellSz') : undefined;
  const width = finiteNumber(size?.getAttribute('width'));
  const fallbackWidth = gridWidths.slice(col, col + colSpan).reduce((sum, item) => sum + item, 0);
  const tcPr: string[] = [`<w:tcW w:w="${width && width > 0 ? hwpUnitToTwip(width) : fallbackWidth}" w:type="dxa"/>`];
  if (colSpan > 1) tcPr.push(`<w:gridSpan w:val="${colSpan}"/>`);
  if (vMerge) tcPr.push(vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>');

  if (cell) {
    const margin = firstDirectChild(cell, 'cellMargin');
    if (margin) {
      tcPr.push(`<w:tcMar>${(['top', 'left', 'bottom', 'right'] as const).map((side) => `<w:${side} w:w="${hwpUnitToTwip(finiteNumber(margin.getAttribute(side)) ?? 0)}" w:type="dxa"/>`).join('')}</w:tcMar>`);
    }
    const subList = firstDirectChild(cell, 'subList');
    const vertical = subList?.getAttribute('vertAlign');
    if (vertical) tcPr.push(`<w:vAlign w:val="${vertical === 'CENTER' ? 'center' : vertical === 'BOTTOM' ? 'bottom' : 'top'}"/>`);
    const border = ctx.catalog.borders.get(cell.getAttribute('borderFillIDRef') ?? '');
    if (border?.cellBorders) tcPr.push(border.cellBorders);
    if (border?.shading) tcPr.push(border.shading);
  }

  const paragraphs = cell
    ? directChildren(firstDirectChild(cell, 'subList') ?? cell, 'p').map((paragraph) => renderParagraph(paragraph, ctx)).join('')
    : '';
  return `<w:tc><w:tcPr>${tcPr.join('')}</w:tcPr>${paragraphs || '<w:p/>'}</w:tc>`;
}

function renderHeaderFooterParagraphs(element: Element, ctx: RenderContext): string {
  const subList = firstDirectChild(element, 'subList');
  if (!subList) return '<w:p/>';
  const paragraphs = directChildren(subList, 'p').map((paragraph) => {
    if (directRunChildren(paragraph, 'pic').length) {
      ctx.warnings.add('머리말/꼬리말 안의 그림은 현재 DOCX 직접 저장에서 제외됩니다.');
    }
    return renderParagraph(paragraph, { ...ctx, images: new Map() });
  });
  return paragraphs.join('');
}

function renderNotesPart(kind: NotePart['kind'], notes: NotePart[], ctx: RenderContext): string {
  const rootTag = kind === 'footnote' ? 'footnotes' : 'endnotes';
  const separator = `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>`;
  const continuation = `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`;
  const actual = notes.map((note) => renderNote(note, ctx)).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${rootTag} xmlns:w="${WORD_NS}" xmlns:r="${REL_NS}" xmlns:m="${MATH_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="${WPS_NS}">${separator}${continuation}${actual}</w:${rootTag}>`;
}

function renderNote(note: NotePart, ctx: RenderContext): string {
  const subList = firstDirectChild(note.element, 'subList');
  const paragraphs = subList ? directChildren(subList, 'p') : [];
  const style = note.kind === 'footnote' ? 'FootnoteReference' : 'EndnoteReference';
  const refTag = note.kind === 'footnote' ? 'footnoteRef' : 'endnoteRef';
  const marker = `<w:r><w:rPr><w:rStyle w:val="${style}"/></w:rPr><w:${refTag}/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>`;
  const body = paragraphs.length
    ? paragraphs.map((paragraph, index) => renderParagraph(paragraph, ctx, false, index === 0 ? marker : '')).join('')
    : `<w:p>${marker}</w:p>`;
  return `<w:${note.kind} w:id="${note.id}">${body}</w:${note.kind}>`;
}

async function collectImages(zip: JSZip): Promise<Map<string, ImageResource>> {
  const result = new Map<string, ImageResource>();
  const content = await readOptionalText(zip, ['Contents/content.hpf', 'contents/content.hpf']);
  if (!content) return result;
  const document = parseXml(content, 'HWPX content.hpf');
  let index = 1;
  for (const item of descendants(document.documentElement, 'item')) {
    const id = item.getAttribute('id');
    const href = item.getAttribute('href');
    if (!id || !href || !/^BinData\//i.test(href)) continue;
    const mediaType = item.getAttribute('media-type') ?? mediaTypeFromPath(href);
    const extension = href.split('.').pop()?.toLowerCase() || extensionFromMediaType(mediaType);
    const targetName = `image${index}.${sanitizeExtension(extension)}`;
    result.set(id, {
      id,
      sourcePath: href,
      targetName,
      mediaType,
      relationshipId: `rIdImage${index}`,
    });
    index += 1;
  }
  return result;
}

function renderNumbering(numberingIds: Map<string, number>): string {
  const abstract: string[] = [];
  const instances: string[] = [];
  for (const [key, numId] of numberingIds) {
    const kind = key.startsWith('bullet:') ? 'bullet' : 'number';
    const levels: string[] = [];
    for (let level = 0; level < 9; level += 1) {
      const left = 720 * (level + 1);
      const hanging = 360;
      if (kind === 'bullet') {
        const bullets = ['•', '◦', '▪', '▫', '◆', '◇', '●', '○', '■'];
        levels.push(`<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${bullets[level]}"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="${left}"/></w:tabs><w:ind w:left="${left}" w:hanging="${hanging}"/></w:pPr></w:lvl>`);
      } else {
        const pattern = Array.from({ length: level + 1 }, (_, i) => `%${i + 1}`).join('.') + '.';
        levels.push(`<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="${pattern}"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="${left}"/></w:tabs><w:ind w:left="${left}" w:hanging="${hanging}"/></w:pPr></w:lvl>`);
      }
    }
    abstract.push(`<w:abstractNum w:abstractNumId="${numId}"><w:multiLevelType w:val="multilevel"/>${levels.join('')}</w:abstractNum>`);
    instances.push(`<w:num w:numId="${numId}"><w:abstractNumId w:val="${numId}"/></w:num>`);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${WORD_NS}">${abstract.join('')}${instances.join('')}</w:numbering>`;
}

function renderStyles(catalog: StyleCatalog): string {
  const defaultFont = catalog.fonts.hangul.values().next().value
    ?? catalog.fonts.latin.values().next().value
    ?? 'Arial';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${WORD_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${escapeXml(defaultFont)}" w:hAnsi="${escapeXml(defaultFont)}" w:eastAsia="${escapeXml(defaultFont)}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:semiHidden/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style><w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:semiHidden/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>`;
}

function renderSettings(ctx: PackageContext): string {
  const hasEven = ctx.headerFooterParts.some((part) => part.type === 'even');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${WORD_NS}">${hasEven ? '<w:evenAndOddHeaders/>' : ''}<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;
}

function renderContentTypes(ctx: PackageContext): string {
  const defaults = new Map<string, string>([
    ['rels', 'application/vnd.openxmlformats-package.relationships+xml'],
    ['xml', 'application/xml'],
  ]);
  for (const image of ctx.images.values()) defaults.set(image.targetName.split('.').pop() ?? 'bin', image.mediaType);
  const overrides = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
  ];
  if (ctx.numberingIds.size) overrides.push('<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>');
  if (ctx.footnotes.length) overrides.push('<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>');
  if (ctx.endnotes.length) overrides.push('<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>');
  for (const part of ctx.headerFooterParts) {
    const type = part.kind === 'header'
      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
      : 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';
    overrides.push(`<Override PartName="/word/${part.targetName}" ContentType="${type}"/>`);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CONTENT_TYPES_NS}">${[...defaults].map(([extension, type]) => `<Default Extension="${escapeXml(extension)}" ContentType="${escapeXml(type)}"/>`).join('')}${overrides.join('')}</Types>`;
}

function renderCoreProperties(): string {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>OmniDocs</dc:creator><cp:lastModifiedBy>OmniDocs</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
}

function renderAppProperties(): string {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>OmniDocs</Application></Properties>';
}

function relationshipXml(id: string, type: string, target: string): string {
  return `<Relationship Id="${escapeXml(id)}" Type="${escapeXml(type)}" Target="${escapeXml(target)}"/>`;
}

function ensureNumberingId(ctx: RenderContext, key: string): number {
  const existing = ctx.numberingIds.get(key);
  if (existing) return existing;
  ctx.warnings.add('HWPX의 사용자 정의 문단 번호/글머리표 모양은 DOCX 다단계 목록 규칙으로 정규화됩니다.');
  const next = ctx.numberingIds.size + 1;
  ctx.numberingIds.set(key, next);
  return next;
}

function detectUnsupportedHwpxFeatures(root: Element, warnings: Set<string>): void {
  const names = new Set(Array.from(root.getElementsByTagName('*'), (element) => element.localName.toLowerCase()));
  const hasAny = (...values: string[]) => values.some((value) => names.has(value.toLowerCase()));
  if (hasAny('polygon', 'curve', 'arc', 'container', 'textart', 'ole', 'chart', 'video')) {
    warnings.add('HWPX 자유곡선/복합 도형/차트/OLE 일부는 현재 DOCX 직접 저장에서 제외되거나 단순화될 수 있습니다.');
  }
  if (hasAny('fieldbegin', 'fieldend', 'fieldBegin', 'fieldEnd')) {
    warnings.add('HWPX 동적 필드는 현재 DOCX 직접 저장에서 Word 필드 코드로 완전히 변환되지 않습니다.');
  }
}

function mapHorizontalRelativeFrom(value: string | null, ctx: RenderContext): string {
  switch ((value ?? 'PAPER').toUpperCase()) {
    case 'COLUMN': return 'column';
    case 'PARA':
      ctx.warnings.add('HWPX의 문단 기준 가로 떠있는 개체 위치는 DOCX character 기준으로 매핑됩니다.');
      return 'character';
    case 'MARGIN': return 'margin';
    case 'PAGE':
    case 'PAPER': return 'page';
    default:
      ctx.warnings.add(`알 수 없는 HWPX 가로 위치 기준(${value ?? ''})을 DOCX page 기준으로 정규화했습니다.`);
      return 'page';
  }
}

function mapVerticalRelativeFrom(value: string | null, ctx: RenderContext): string {
  switch ((value ?? 'PAPER').toUpperCase()) {
    case 'PARA': return 'paragraph';
    case 'LINE': return 'line';
    case 'MARGIN': return 'margin';
    case 'PAGE':
    case 'PAPER': return 'page';
    default:
      ctx.warnings.add(`알 수 없는 HWPX 세로 위치 기준(${value ?? ''})을 DOCX page 기준으로 정규화했습니다.`);
      return 'page';
  }
}

function mapHorizontalObjectAlign(value: string | null): string | undefined {
  switch ((value ?? '').toUpperCase()) {
    case 'LEFT': return 'left';
    case 'CENTER': return 'center';
    case 'RIGHT': return 'right';
    case 'INSIDE': return 'inside';
    case 'OUTSIDE': return 'outside';
    default: return undefined;
  }
}

function mapVerticalObjectAlign(value: string | null): string | undefined {
  switch ((value ?? '').toUpperCase()) {
    case 'TOP': return 'top';
    case 'CENTER': return 'center';
    case 'BOTTOM': return 'bottom';
    case 'INSIDE': return 'inside';
    case 'OUTSIDE': return 'outside';
    default: return undefined;
  }
}

function mapWrapSide(value: string): string {
  switch (value) {
    case 'LEFT_ONLY': return 'left';
    case 'RIGHT_ONLY': return 'right';
    case 'LARGEST_ONLY': return 'largest';
    default: return 'bothSides';
  }
}

function mapTableHorizontalAnchor(value: string | null, ctx: RenderContext): string {
  switch ((value ?? 'PAPER').toUpperCase()) {
    case 'MARGIN': return 'margin';
    case 'COLUMN':
    case 'PARA': return 'text';
    case 'PAGE':
    case 'PAPER': return 'page';
    default:
      ctx.warnings.add(`알 수 없는 HWPX 표 가로 위치 기준(${value ?? ''})을 DOCX page 기준으로 정규화했습니다.`);
      return 'page';
  }
}

function mapTableVerticalAnchor(value: string | null, ctx: RenderContext): string {
  switch ((value ?? 'PAPER').toUpperCase()) {
    case 'MARGIN': return 'margin';
    case 'PARA':
    case 'LINE': return 'text';
    case 'PAGE':
    case 'PAPER': return 'page';
    default:
      ctx.warnings.add(`알 수 없는 HWPX 표 세로 위치 기준(${value ?? ''})을 DOCX page 기준으로 정규화했습니다.`);
      return 'page';
  }
}

function inferColumnCount(rows: Element[]): number {
  let max = 1;
  for (const row of rows) {
    for (const cell of directChildren(row, 'tc')) {
      const address = firstDirectChild(cell, 'cellAddr');
      const span = firstDirectChild(cell, 'cellSpan');
      const col = finiteNumber(address?.getAttribute('colAddr')) ?? 0;
      const colSpan = finiteNumber(span?.getAttribute('colSpan')) ?? 1;
      max = Math.max(max, col + colSpan);
    }
  }
  return Math.trunc(max);
}

function inferGridWidths(rows: Element[], colCount: number): number[] {
  const widths = Array.from({ length: colCount }, () => 0);
  for (const row of rows) {
    for (const cell of directChildren(row, 'tc')) {
      const address = firstDirectChild(cell, 'cellAddr');
      const span = firstDirectChild(cell, 'cellSpan');
      const size = firstDirectChild(cell, 'cellSz');
      const col = Math.max(0, Math.trunc(finiteNumber(address?.getAttribute('colAddr')) ?? 0));
      const colSpan = Math.max(1, Math.trunc(finiteNumber(span?.getAttribute('colSpan')) ?? 1));
      const cellWidth = finiteNumber(size?.getAttribute('width'));
      if (!cellWidth || cellWidth <= 0) continue;
      const perColumn = Math.max(1, Math.round(hwpUnitToTwip(cellWidth) / colSpan));
      for (let index = col; index < Math.min(colCount, col + colSpan); index += 1) widths[index] = Math.max(widths[index], perColumn);
    }
  }
  const fallback = 1440;
  return widths.map((value) => value || fallback);
}

function directRuns(paragraph: Element): Element[] {
  return directChildren(paragraph, 'run');
}

function directRunChildren(paragraph: Element, localName: string): Element[] {
  return directRuns(paragraph).flatMap((run) => directChildren(run, localName));
}

function directChildren(parent: Element, localName: string): Element[] {
  return Array.from(parent.children).filter((child) => child.localName === localName);
}

function firstDirectChild(parent: Element, localName: string): Element | undefined {
  return directChildren(parent, localName)[0];
}

function descendants(parent: Element, localName: string): Element[] {
  return Array.from(parent.getElementsByTagName('*')).filter((element) => element.localName === localName);
}

function firstDescendant(parent: Element, localName: string): Element | undefined {
  return descendants(parent, localName)[0];
}

function parseXml(xml: string, label: string): XMLDocument {
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  if (document.getElementsByTagName('parsererror').length) throw new Error(`${label} XML을 해석하지 못했습니다.`);
  return document;
}

async function readRequiredText(zip: JSZip, paths: string[]): Promise<string> {
  const value = await readOptionalText(zip, paths);
  if (value === undefined) throw new Error(`HWPX 필수 항목을 찾지 못했습니다: ${paths[0]}`);
  return value;
}

async function readOptionalText(zip: JSZip, paths: string[]): Promise<string | undefined> {
  for (const path of paths) {
    const file = zip.file(path);
    if (file) return file.async('text');
  }
  return undefined;
}

function sectionIndex(path: string): number {
  return Number(path.match(/section(\d+)\.xml/i)?.[1] ?? 0);
}

function mapParagraphAlignment(value: string): string {
  switch (value.toUpperCase()) {
    case 'LEFT': return 'left';
    case 'RIGHT': return 'right';
    case 'CENTER': return 'center';
    case 'DISTRIBUTE': return 'distribute';
    default: return 'both';
  }
}

function mapUnderline(shape: string | null): string {
  const value = (shape ?? '').toUpperCase();
  if (value.includes('DOUBLE')) return 'double';
  if (value.includes('DASH')) return 'dash';
  if (value.includes('DOT')) return 'dotted';
  if (value.includes('WAVE')) return 'wave';
  return 'single';
}

function mapBorderType(value: string): string {
  if (value === 'NONE') return 'nil';
  if (value.includes('DOUBLE')) return 'double';
  if (value.includes('DASH_DOT_DOT')) return 'dashDotDot';
  if (value.includes('DASH_DOT')) return 'dashDotStroked';
  if (value.includes('DASH')) return 'dashed';
  if (value.includes('DOT')) return 'dotted';
  return 'single';
}

function mapHeaderFooterType(value: string | null): HeaderFooterPart['type'] {
  switch ((value ?? 'BOTH').toUpperCase()) {
    case 'EVEN': return 'even';
    case 'FIRST': return 'first';
    default: return 'default';
  }
}

function fontName(map: Map<string, string>, id: string | null): string | undefined {
  return id === null ? undefined : map.get(id);
}

function hwpUnitValue(element: Element | undefined): number | undefined {
  return finiteNumber(element?.getAttribute('value'));
}

function hwpUnitToTwip(value: number): number {
  return Math.round(value / 5);
}

function borderWidthEighthPoints(value: string | null): number {
  if (!value) return 4;
  const match = value.match(/(-?\d+(?:\.\d+)?)\s*(mm|pt)?/i);
  if (!match) return 4;
  const number = Number(match[1]);
  if (!Number.isFinite(number) || number <= 0) return 4;
  const points = (match[2] ?? 'mm').toLowerCase() === 'pt' ? number : number * 72 / 25.4;
  return Math.max(2, Math.min(96, Math.round(points * 8)));
}

function normalizeHex(value: string | null | undefined): string | undefined {
  if (!value || value === 'none' || value === 'auto') return undefined;
  const normalized = value.replace(/^#/, '');
  return /^[0-9a-f]{6}$/i.test(normalized) ? normalized.toUpperCase() : undefined;
}

function finiteNumber(value: string | null | undefined): number | undefined {
  if (value === null || value === undefined || value === '') return undefined;
  const number = Number(value);
  return Number.isFinite(number) ? number : undefined;
}

function isOne(value: string | null): boolean {
  return value === '1' || value?.toLowerCase() === 'true';
}

function meaningfulText(value: string): boolean {
  return Boolean(value.replace('[이미지]', '').trim());
}

function mediaTypeFromPath(path: string): string {
  const extension = path.split('.').pop()?.toLowerCase();
  switch (extension) {
    case 'png': return 'image/png';
    case 'jpg':
    case 'jpeg': return 'image/jpeg';
    case 'gif': return 'image/gif';
    case 'bmp': return 'image/bmp';
    case 'tif':
    case 'tiff': return 'image/tiff';
    case 'svg': return 'image/svg+xml';
    case 'emf': return 'image/x-emf';
    case 'wmf': return 'image/x-wmf';
    default: return 'application/octet-stream';
  }
}

function extensionFromMediaType(type: string): string {
  switch (type) {
    case 'image/png': return 'png';
    case 'image/jpeg': return 'jpg';
    case 'image/gif': return 'gif';
    case 'image/bmp': return 'bmp';
    case 'image/tiff': return 'tiff';
    case 'image/svg+xml': return 'svg';
    default: return 'bin';
  }
}

function sanitizeExtension(value: string): string {
  return /^[a-z0-9]{1,8}$/i.test(value) ? value.toLowerCase() : 'bin';
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;');
}
