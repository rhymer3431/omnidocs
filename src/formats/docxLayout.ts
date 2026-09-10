import JSZip from 'jszip';
import type { RhwpPageDef } from './rhwpRuntime';

export interface DocxLayoutProfile {
  pageDef?: Partial<RhwpPageDef>;
  sectionCount: number;
  headerText?: string;
  footerText?: string;
  warnings: string[];
}

const TWIP_TO_HWPUNIT = 5;

export async function extractDocxLayout(bytes: Uint8Array): Promise<DocxLayoutProfile> {
  const zip = await JSZip.loadAsync(bytes);
  const documentXml = await zip.file('word/document.xml')?.async('text');
  if (!documentXml) {
    return { sectionCount: 0, warnings: ['DOCX document.xml을 찾지 못해 페이지 설정을 가져오지 못했습니다.'] };
  }

  const sections = [...documentXml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((match) => match[0]);
  const sectionXml = sections[0];
  const warnings: string[] = [];
  if (sections.length > 1) {
    warnings.push(`DOCX에 ${sections.length}개 구역이 있습니다. 현재 OMDX 변환은 첫 구역의 페이지 설정을 기준으로 정규화합니다.`);
  }

  const pageDef = sectionXml ? parsePageDef(sectionXml) : undefined;
  const relationshipsXml = await zip.file('word/_rels/document.xml.rels')?.async('text');
  const relationships = relationshipsXml ? parseRelationships(relationshipsXml) : new Map<string, string>();

  const headerTarget = sectionXml ? referenceTarget(sectionXml, 'headerReference', relationships) : undefined;
  const footerTarget = sectionXml ? referenceTarget(sectionXml, 'footerReference', relationships) : undefined;
  const headerText = headerTarget ? await readPartText(zip, headerTarget) : undefined;
  const footerText = footerTarget ? await readPartText(zip, footerTarget) : undefined;

  if (documentXml.includes('<w:instrText')) {
    warnings.push('DOCX 필드가 포함되어 있어 동적 필드는 OMDX에서 단순화될 수 있습니다.');
  }

  return {
    pageDef,
    sectionCount: sections.length || 1,
    headerText,
    footerText,
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

function referenceTarget(sectionXml: string, elementName: string, relationships: Map<string, string>): string | undefined {
  const element = sectionXml.match(new RegExp(`<w:${elementName}\\b([^>]*)/?>`))?.[1];
  if (!element) return undefined;
  const relationId = element.match(/r:id="([^"]+)"/)?.[1];
  return relationId ? relationships.get(relationId) : undefined;
}

async function readPartText(zip: JSZip, target: string): Promise<string | undefined> {
  const normalized = target.startsWith('/') ? target.slice(1) : `word/${target.replace(/^\.\//, '')}`;
  const xml = await zip.file(normalized)?.async('text');
  if (!xml) return undefined;

  const tokens = [...xml.matchAll(/<w:(t|tab|br)\b[^>]*>([\s\S]*?)<\/w:\1>|<w:(tab|br)\b[^>]*\/>/g)];
  const text = tokens.map((match) => {
    const kind = match[1] ?? match[3];
    if (kind === 'tab') return '\t';
    if (kind === 'br') return '\n';
    return decodeXml(match[2] ?? '');
  }).join('').trim();
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
