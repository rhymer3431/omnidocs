import mammoth from 'mammoth/mammoth.browser';
import { asBlob } from 'html-docx-js-typescript';
import { createRhwpFromHtml, openRhwp } from './rhwpRuntime';
import { extractDocxLayout, type DocxLayoutProfile } from './docxLayout';

export interface DocxImportResult {
  hwpxBytes: Uint8Array;
  warnings: string[];
  layout: DocxLayoutProfile;
}

export async function importDocx(bytes: Uint8Array): Promise<DocxImportResult> {
  const layout = await extractDocxLayout(bytes);
  const arrayBuffer = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
  const converted = await convertDocxToLayoutHtml(arrayBuffer);
  const document = await createRhwpFromHtml(converted.html, {
    pageDef: layout.pageDef,
    headerText: layout.headerText,
    footerText: layout.footerText,
  });
  try {
    return {
      hwpxBytes: document.exportHwpx(),
      warnings: [
        ...converted.warnings,
        ...layout.warnings,
      ],
      layout,
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

  return {
    html: inlineGeneratedStyles(result.value, classStyles),
    warnings: result.messages.map((message: { message: string }) => message.message),
  };
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

export async function exportDocx(hwpxBytes: Uint8Array): Promise<Uint8Array> {
  const document = await openRhwp(hwpxBytes);
  try {
    const pages: string[] = [];
    for (let page = 0; page < document.pageCount(); page += 1) {
      pages.push(`<section class="omnidocs-page">${document.renderPageHtml(page)}</section>`);
    }
    const html = `<!doctype html><html><head><meta charset="utf-8"></head><body>${pages.join('')}</body></html>`;
    const binary = await asBlob(html);
    if (binary instanceof Uint8Array) return Uint8Array.from(binary);
    return new Uint8Array(await binary.arrayBuffer());
  } finally {
    document.free();
  }
}
