import JSZip from 'jszip';
import type { RhwpPageDef } from './rhwpRuntime';

export interface HwpxSectionPatchResult {
  bytes: Uint8Array;
  patchedSections: number;
  warnings: string[];
}

export async function patchHwpxSectionPageDefs(
  hwpxBytes: Uint8Array,
  pageDefs: Array<Partial<RhwpPageDef> | undefined>,
): Promise<HwpxSectionPatchResult> {
  if (pageDefs.length <= 1) {
    return { bytes: Uint8Array.from(hwpxBytes), patchedSections: 0, warnings: [] };
  }

  const zip = await JSZip.loadAsync(hwpxBytes);
  const sectionPath = zip.file('Contents/section0.xml')
    ? 'Contents/section0.xml'
    : zip.file('contents/section0.xml')
      ? 'contents/section0.xml'
      : undefined;
  if (!sectionPath) {
    return {
      bytes: Uint8Array.from(hwpxBytes),
      patchedSections: 0,
      warnings: ['HWPX section0.xml을 찾지 못해 DOCX 다중 구역 용지 설정을 적용하지 못했습니다.'],
    };
  }

  const xml = await zip.file(sectionPath)!.async('text');
  const matches = [...xml.matchAll(/<hp:secPr\b[\s\S]*?<\/hp:secPr>/g)];
  const patchCount = Math.min(matches.length, pageDefs.length);
  let index = 0;
  const patchedXml = xml.replace(/<hp:secPr\b[\s\S]*?<\/hp:secPr>/g, (sectionXml) => {
    const current = index < patchCount ? pageDefs[index] : undefined;
    index += 1;
    return current ? patchSectionPageDef(sectionXml, current) : sectionXml;
  });

  zip.file(sectionPath, patchedXml);
  const warnings = matches.length < pageDefs.length
    ? [`DOCX 구역 ${pageDefs.length}개 중 ${matches.length}개만 canonical HWPX 구역으로 대응되었습니다.`]
    : [];
  return {
    bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } }),
    patchedSections: patchCount,
    warnings,
  };
}

function patchSectionPageDef(sectionXml: string, pageDef: Partial<RhwpPageDef>): string {
  let result = replaceTagAttributes(sectionXml, 'hp:pagePr', {
    width: integerValue(pageDef.width),
    height: integerValue(pageDef.height),
    landscape: pageDef.landscape === undefined ? undefined : pageDef.landscape ? 'NARROWLY' : 'WIDELY',
  });
  result = replaceTagAttributes(result, 'hp:margin', {
    header: integerValue(pageDef.marginHeader),
    footer: integerValue(pageDef.marginFooter),
    gutter: integerValue(pageDef.marginGutter),
    left: integerValue(pageDef.marginLeft),
    right: integerValue(pageDef.marginRight),
    top: integerValue(pageDef.marginTop),
    bottom: integerValue(pageDef.marginBottom),
  });
  return result;
}

function replaceTagAttributes(
  xml: string,
  tagName: string,
  values: Record<string, string | undefined>,
): string {
  const escaped = tagName.replace(':', '\\:');
  const regex = new RegExp(`(<${escaped}\\b)([^>]*)(/?>)`);
  return xml.replace(regex, (_whole, prefix: string, rawAttributes: string, suffix: string) => {
    let attributes = rawAttributes;
    for (const [name, value] of Object.entries(values)) {
      if (value === undefined) continue;
      const existing = new RegExp(`\\s${name}="[^"]*"`);
      if (existing.test(attributes)) attributes = attributes.replace(existing, ` ${name}="${value}"`);
      else attributes += ` ${name}="${value}"`;
    }
    return `${prefix}${attributes}${suffix}`;
  });
}

function integerValue(value: number | undefined): string | undefined {
  return value === undefined || !Number.isFinite(value) ? undefined : String(Math.round(value));
}
