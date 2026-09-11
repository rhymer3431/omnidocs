import JSZip from 'jszip';
import type { DocumentFeatureEntry, DocumentFeatureInventory, FeatureHandling } from './canonical';

interface FeatureRule {
  id: string;
  label: string;
  handling: FeatureHandling;
  pattern: RegExp;
  scope?: 'document' | 'all-word-xml';
  note?: string;
}

const RULES: FeatureRule[] = [
  { id: 'paragraph', label: '문단', handling: 'native', pattern: /<w:p\b/g, scope: 'document' },
  { id: 'table', label: '표', handling: 'native', pattern: /<w:tbl\b/g, scope: 'document' },
  { id: 'section', label: '구역', handling: 'normalized', pattern: /<w:sectPr\b/g, scope: 'document' },
  { id: 'drawing', label: 'DrawingML 그림/도형', handling: 'approximated', pattern: /<w:drawing\b/g, scope: 'all-word-xml', note: '그림 크기는 보존하지만 inline/floating 배치는 일부 근사됩니다.' },
  { id: 'legacy-picture', label: 'VML/legacy 그림', handling: 'approximated', pattern: /<w:pict\b/g, scope: 'all-word-xml' },
  { id: 'hyperlink', label: '하이퍼링크', handling: 'normalized', pattern: /<w:hyperlink\b/g, scope: 'all-word-xml' },
  { id: 'field', label: 'Word 필드', handling: 'source-only', pattern: /<w:instrText\b|<w:fldSimple\b/g, scope: 'all-word-xml' },
  { id: 'content-control', label: '콘텐츠 컨트롤', handling: 'source-only', pattern: /<w:sdt\b/g, scope: 'all-word-xml' },
  { id: 'track-change', label: '변경 내용 추적', handling: 'source-only', pattern: /<w:(?:ins|del|moveFrom|moveTo)\b/g, scope: 'all-word-xml' },
  { id: 'comment-reference', label: '주석', handling: 'source-only', pattern: /<w:commentReference\b/g, scope: 'all-word-xml' },
  { id: 'footnote-reference', label: '각주', handling: 'approximated', pattern: /<w:footnoteReference\b/g, scope: 'document' },
  { id: 'endnote-reference', label: '미주', handling: 'approximated', pattern: /<w:endnoteReference\b/g, scope: 'document' },
  { id: 'equation', label: '수식', handling: 'approximated', pattern: /<m:oMath(?:Para)?\b/g, scope: 'all-word-xml' },
  { id: 'textbox', label: '텍스트 상자', handling: 'approximated', pattern: /<w:txbxContent\b/g, scope: 'all-word-xml' },
  { id: 'bookmark', label: '책갈피', handling: 'source-only', pattern: /<w:bookmarkStart\b/g, scope: 'document' },
  { id: 'page-break', label: '명시적 쪽 나누기', handling: 'normalized', pattern: /<w:br\b[^>]*w:type=["']page["']/g, scope: 'document' },
];

export async function analyzeDocxFeatures(bytes: Uint8Array): Promise<DocumentFeatureInventory> {
  const zip = await JSZip.loadAsync(bytes);
  const wordXml: string[] = [];
  let documentXml = '';

  const reads: Promise<void>[] = [];
  zip.forEach((path, entry) => {
    if (entry.dir || !/^word\/.*\.xml$/i.test(path)) return;
    reads.push(entry.async('text').then((xml) => {
      wordXml.push(xml);
      if (path.toLowerCase() === 'word/document.xml') documentXml = xml;
    }));
  });
  await Promise.all(reads);

  const allXml = wordXml.join('\n');
  const features: DocumentFeatureEntry[] = RULES.map((rule) => ({
    id: rule.id,
    label: rule.label,
    handling: rule.handling,
    count: countMatches(rule.scope === 'document' ? documentXml : allXml, rule.pattern),
    note: rule.note,
  }));

  const headerCount = countPackageParts(zip, /^word\/header\d+\.xml$/i);
  const footerCount = countPackageParts(zip, /^word\/footer\d+\.xml$/i);
  if (headerCount > 0) features.push({
    id: 'header-part', label: '머리말', count: headerCount, handling: 'normalized',
    note: '기본/홀짝 머리말은 정규화하며 일부 구역별 차이는 원본 메타데이터에 보존됩니다.',
  });
  if (footerCount > 0) features.push({
    id: 'footer-part', label: '꼬리말', count: footerCount, handling: 'normalized',
    note: '기본/홀짝 꼬리말은 정규화하며 일부 구역별 차이는 원본 메타데이터에 보존됩니다.',
  });

  return {
    sourceFormat: 'docx',
    adapter: 'omnidocs.docx-ooxml-v1',
    features,
  };
}

function countMatches(source: string, pattern: RegExp): number {
  if (!source) return 0;
  const flags = pattern.flags.includes('g') ? pattern.flags : `${pattern.flags}g`;
  return [...source.matchAll(new RegExp(pattern.source, flags))].length;
}

function countPackageParts(zip: JSZip, pattern: RegExp): number {
  let count = 0;
  zip.forEach((path, entry) => {
    if (!entry.dir && pattern.test(path)) count += 1;
    pattern.lastIndex = 0;
  });
  return count;
}
