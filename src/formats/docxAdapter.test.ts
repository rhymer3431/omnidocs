import { JSDOM } from 'jsdom';
import {
  AlignmentType,
  BorderStyle,
  Document,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { convertDocxToLayoutHtml } from './docxAdapter';
import { decodeOmdxFormat, OMDX_CHAR_FORMAT_ATTR, OMDX_PARA_FORMAT_ATTR } from './docxOoxml';

const originalDOMParser = globalThis.DOMParser;
const originalNode = globalThis.Node;

beforeAll(() => {
  const dom = new JSDOM();
  Object.assign(globalThis, {
    DOMParser: dom.window.DOMParser,
    Node: dom.window.Node,
  });
});

afterAll(() => {
  Object.assign(globalThis, { DOMParser: originalDOMParser, Node: originalNode });
});

const ONE_PIXEL_PNG = Uint8Array.from([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
  0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196, 137,
  0, 0, 0, 13, 73, 68, 65, 84, 120, 218, 99, 252, 207, 192, 80,
  15, 0, 5, 254, 2, 254, 227, 21, 199, 205, 0, 0, 0, 0, 73, 69,
  78, 68, 174, 66, 96, 130,
]);

describe('DOCX semantic style bridge', () => {
  it('keeps Word paragraph geometry and character formatting as rHWP-native metadata', async () => {
    const fixture = new Document({
      sections: [{
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            indent: { left: 720, firstLine: 360 },
            spacing: { before: 240, after: 120, line: 360 },
            keepNext: true,
            pageBreakBefore: true,
            children: [new TextRun({
              text: '서식 보존',
              font: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'Malgun Gothic' },
              size: 28,
              bold: true,
              italics: true,
              underline: {},
              color: 'FF0000',
              superScript: true,
            })],
          }),
        ],
      }],
    });
    const buffer = await Packer.toBuffer(fixture);
    const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer;
    const converted = await convertDocxToLayoutHtml(arrayBuffer);

    expect(converted.html).toContain('font-weight:bold');
    expect(converted.html).toContain('font-size:14pt');
    expect(converted.html).toContain('서식 보존');

    const parsed = new DOMParser().parseFromString(`<body>${converted.html}</body>`, 'text/html');
    const paragraph = parsed.body.querySelector('p');
    const formattedRun = parsed.body.querySelector(`[${OMDX_CHAR_FORMAT_ATTR}]`);
    const paraFormat = decodeOmdxFormat<Record<string, unknown>>(
      paragraph?.getAttribute(OMDX_PARA_FORMAT_ATTR) ?? null,
    );
    const charFormat = decodeOmdxFormat<Record<string, unknown>>(
      formattedRun?.getAttribute(OMDX_CHAR_FORMAT_ATTR) ?? null,
    );

    expect(paragraph?.style.textAlign).toBe('center');
    expect(paraFormat).toMatchObject({
      alignment: 'center',
      marginLeft: 3600,
      indent: 1800,
      spacingBefore: 1200,
      spacingAfter: 600,
      lineSpacing: 150,
      lineSpacingType: 'Percent',
      keepWithNext: true,
      pageBreakBefore: true,
    });
    expect(charFormat).toMatchObject({
      fontName: 'Malgun Gothic',
      fontSize: 1400,
      textColor: '#FF0000',
      bold: true,
      italic: true,
      underline: true,
      superscript: true,
    });
  });

  it('restores Word table geometry and promotes inline images so rHWP does not drop them', async () => {
    const fixture = new Document({
      sections: [{
        children: [
          new Paragraph({
            children: [
              new TextRun('앞'),
              new ImageRun({
                data: ONE_PIXEL_PNG,
                transformation: { width: 120, height: 60 },
                type: 'png',
              }),
              new TextRun('뒤'),
            ],
          }),
          new Table({
            width: { size: 8000, type: WidthType.DXA },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 8, color: 'FF0000' },
              bottom: { style: BorderStyle.SINGLE, size: 8, color: 'FF0000' },
              left: { style: BorderStyle.SINGLE, size: 8, color: 'FF0000' },
              right: { style: BorderStyle.SINGLE, size: 8, color: 'FF0000' },
              insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: '999999' },
              insideVertical: { style: BorderStyle.SINGLE, size: 4, color: '999999' },
            },
            rows: [new TableRow({
              children: [
                new TableCell({
                  width: { size: 3000, type: WidthType.DXA },
                  margins: { top: 100, right: 200, bottom: 100, left: 200 },
                  children: [new Paragraph('A')],
                }),
                new TableCell({
                  width: { size: 5000, type: WidthType.DXA },
                  children: [new Paragraph('B')],
                }),
              ],
            })],
          }),
        ],
      }],
    });

    const buffer = await Packer.toBuffer(fixture);
    const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer;
    const converted = await convertDocxToLayoutHtml(arrayBuffer);
    const parsed = new DOMParser().parseFromString(`<body>${converted.html}</body>`, 'text/html');
    const image = parsed.body.querySelector('img');
    const table = parsed.body.querySelector('table');
    const firstCell = parsed.body.querySelector('td');

    expect(image).not.toBeNull();
    expect(image?.parentElement).toBe(parsed.body);
    expect(image?.getAttribute('width')).toBe('120');
    expect(image?.getAttribute('height')).toBe('60');
    expect(parsed.body.textContent).toContain('앞');
    expect(parsed.body.textContent).toContain('뒤');

    expect(table?.style.width).toBe('400pt');
    expect(table?.style.borderCollapse).toBe('collapse');
    expect(firstCell?.style.width).toBe('150pt');
    expect(firstCell?.style.paddingLeft).toBe('10pt');
    expect(firstCell?.style.paddingTop).toBe('5pt');
    expect(firstCell?.style.borderTop).toContain('1pt');
  });
});
