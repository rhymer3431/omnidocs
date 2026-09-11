import JSZip from 'jszip';
import type { RhwpPageDef } from './rhwpRuntime';

export interface DocxLayoutProfile {
  pageDef?: Partial<RhwpPageDef>;
  sectionCount: number;
  headerText?: string;
  footerText?: string;
  sections: DocxSectionLayout[];
  warnings: string[];
}

export type DocxHeaderFooterType = 'default' | 'first' | 'even';

export interface DocxSectionLayout {
  pageDef?: Partial<RhwpPageDef>;
  breakType?: string;
  titlePage: boolean;
  headers: Partial<Record<DocxHeaderFooterType, string>>;
  footers: Partial<Record<DocxHeaderFooterType, string>>;
}

const TWIP_TO_HWPUNIT = 5;

export async function extractDocxLayout(bytes: Uint8Array): Promise<DocxLayoutProfile> {
  const zip = await JSZip.loadAsync(bytes);
  const documentXml = await zip.file('word/document.xml')?.async('text');
  if (!documentXml) {
    return { sectionCount: 0, sections: [], warnings: ['DOCX document.xml을 찾지 못해 페이지 설정을 가져오지 못했습니다.'] };
  }

  const sectionXmls = [...documentXml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((match) => match[0]);
  const sectionXml = sectionXmls[0];
  const warnings: string[] = [];
  const relationshipsXml = await zip.file('word/_rels/document.xml.rels')?.async('text');
  const relationships = relationshipsXml ? parseRelationships(relationshipsXml) : new Map<string, string>();
  const sections: DocxSectionLayout[] = [];

  for (const currentXml of sectionXmls.length ? sectionXmls : ['']) {
    sections.push({
      pageDef: currentXml ? parsePageDef(currentXml) : undefined,
      breakType: currentXml ? sectionBreakType(currentXml) : undefined,
      titlePage: currentXml ? /<w:titlePg\b/.test(currentXml) : false,
      headers: currentXml ? await readHeaderFooterReferences(zip, currentXml, 'headerReference', relationships) : {},
      footers: currentXml ? await readHeaderFooterReferences(zip, currentXml, 'footerReference', relationships) : {},
    });
  }

  // Word sections inherit a missing header/footer reference from the preceding
  // section. Materialise that inheritance in the OMDX compatibility metadata.
  for (let index = 1; index < sections.length; index += 1) {
    sections[index].headers = { ...sections[index - 1].headers, ...sections[index].headers };
    sections[index].footers = { ...sections[index - 1].footers, ...sections[index].footers };
  }

  const firstSection = sections[0];
  const pageDef = firstSection?.pageDef;
  const headerText = firstSection?.headers.default ?? firstSection?.headers.first ?? firstSection?.headers.even;
  const footerText = firstSection?.footers.default ?? firstSection?.footers.first ?? firstSection?.footers.even;

  if (sections.length > 1) {
    warnings.push(`DOCX의 ${sections.length}개 구역 용지 설정을 OMDX 구역 메타데이터로 보존합니다.`);
    const hasPerSectionHeaderFooter = sections.slice(1).some((section) =>
      JSON.stringify(section.headers) !== JSON.stringify(firstSection?.headers ?? {})
      || JSON.stringify(section.footers) !== JSON.stringify(firstSection?.footers ?? {}));
    if (hasPerSectionHeaderFooter) {
      warnings.push('구역마다 다른 DOCX 머리말/꼬리말은 현재 첫 구역 기준으로 표시되며 원본 정보는 OMDX 메타데이터에 보존됩니다.');
    }
  }
  if (sections.some((section) => section.titlePage && (section.headers.first || section.footers.first))) {
    warnings.push('DOCX 첫 페이지 전용 머리말/꼬리말은 현재 일반 머리말/꼬리말로 근사될 수 있습니다.');
  }

  if (documentXml.includes('<w:instrText')) {
    warnings.push('DOCX 필드가 포함되어 있어 동적 필드는 OMDX에서 단순화될 수 있습니다.');
  }

  return {
    pageDef,
    sectionCount: sections.length || 1,
    headerText,
    footerText,
    sections,
    warnings,
  };
}

function parsePageDef(sectionXml: string): Partial<RhwpPageDef> | undefined {
  const pageSize = sectionXml.match(/<w:pgSz\b([^>]*)\/?\s*>/);
  const pageMargin = sectionXml.match(/<w:pgMar\b([^>]*)\/?\s*>/);
  if (!pageSize && !pageMargin) return undefined;

  const result: Partial<RhwpPageDef> = {};
  if (pageSize) {
    const width = numericAttr(pageSize[1], 'w');
    const height = numericAttr(pageSize[1], 'h');
    const orient = stringAttr(pageSize[1], 'orient');
    if (width) result.width = width * TWIP_TO_HWPUNIT;
    if (height) result.height = height * TWIP_TO_HWPUNIT;
    if (orient) result.landscape = orient === 'landscape';
  }
  if (pageMargin) {
    assignTwip(result, 'marginTop', numericAttr(pageMargin[1], 'top'));
    assignTwip(result, 'marginRight', numericAttr(pageMargin[1], 'right'));
    assignTwip(result, 'marginBottom', numericAttr(pageMargin[1], 'bottom'));
    assignTwip(result, 'marginLeft', numericAttr(pageMargin[1], 'left'));
    assignTwip(result, 'marginHeader', numericAttr(pageMargin[1], 'header'));
    assignTwip(result, 'marginFooter', numericAttr(pageMargin[1], 'footer'));
    assignTwip(result, 'marginGutter', numericAttr(pageMargin[1], 'gutter'));
  }
  return result;
}

function assignTwip(target: Partial<RhwpPageDef>, key: keyof RhwpPageDef, value?: number) {
  if (value !== undefined) (target as Record<string, number | boolean>)[key] = value * TWIP_TO_HWPUNIT;
}

function numericAttr(attributes: string, name: string): number | undefined {
  const value = stringAttr(attributes, name);
  if (value === undefined) return undefined;
  const number = Number(value);
  return Number.isFinite(number) ? number : undefined;
}

function stringAttr(attributes: string, name: string): string | undefined {
  return attributes.match(new RegExp(`(?:w:)?${name}="([^"]+)"`))?.[1];
}

function parseRelationships(xml: string): Map<string, string> {
  const result = new Map<string, string>();
  for (const match of xml.matchAll(/<Relationship\b([^>]*)\/?\s*>/g)) {
    const id = match[1].match(/\bId="([^"]+)"/)?.[1];
    const target = match[1].match(/\bTarget="([^"]+)"/)?.[1];
    if (id && target) result.set(id, target);
  }
  return result;
}

function sectionBreakType(sectionXml: string): string | undefined {
  const element = sectionXml.match(/<w:type\b([^>]*)\/?>/)?.[1];
  return element ? stringAttr(element, 'val') : undefined;
}

async function readHeaderFooterReferences(
  zip: JSZip,
  sectionXml: string,
  elementName: 'headerReference' | 'footerReference',
  relationships: Map<string, string>,
): Promise<Partial<Record<DocxHeaderFooterType, string>>> {
  const result: Partial<Record<DocxHeaderFooterType, string>> = {};
  const regex = new RegExp(`<w:${elementName}\\b([^>]*)/?>`, 'g');
  for (const match of sectionXml.matchAll(regex)) {
    const attributes = match[1];
    const relationId = attributes.match(/r:id="([^"]+)"/)?.[1];
    const type = stringAttr(attributes, 'type');
    if (!relationId || (type !== 'default' && type !== 'first' && type !== 'even')) continue;
    const target = relationships.get(relationId);
    const text = target ? await readPartText(zip, target) : undefined;
    if (text) result[type] = text;
  }
  return result;
}

async function readPartText(zip: JSZip, target: string): Promise<string | undefined> {
  const normalized = target.startsWith('/') ? target.slice(1) : `word/${target.replace(/^\.\//, '')}`;
  const xml = await zip.file(normalized)?.async('text');
  if (!xml) return undefined;

  const paragraphs = [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map((match) => match[0]);
  const sources = paragraphs.length ? paragraphs : [xml];
  const text = sources.map((source) => {
    const tokens = [...source.matchAll(/<w:(t)\b[^>]*>([\s\S]*?)<\/w:\1>|<w:(tab|br)\b[^>]*\/>/g)];
    return tokens.map((match) => {
      const kind = match[1] ?? match[3];
      if (kind === 'tab') return '\t';
      if (kind === 'br') return '\n';
      return decodeXml(match[2] ?? '');
    }).join('');
  }).join('\n').trim();
  return text || undefined;
}

function decodeXml(value: string): string {
  return value
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
    .replaceAll('&amp;', '&');
}
