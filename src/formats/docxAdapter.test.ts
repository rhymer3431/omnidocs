import { AlignmentType, Document, Packer, Paragraph, TextRun } from 'docx';
import { describe, expect, it } from 'vitest';
import { convertDocxToLayoutHtml } from './docxAdapter';

describe('DOCX semantic style bridge', () => {
  it('keeps explicit paragraph alignment, run font/size, bold and underline in HTML', async () => {
    const fixture = new Document({
      sections: [{
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: '서식 보존', font: 'Arial', size: 28, bold: true, underline: {} })],
          }),
        ],
      }],
    });
    const buffer = await Packer.toBuffer(fixture);
    const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer;
    const converted = await convertDocxToLayoutHtml(arrayBuffer);

    expect(converted.html).toContain('text-align:center');
    expect(converted.html).toContain('font-family:Arial');
    expect(converted.html).toContain('font-size:14pt');
    expect(converted.html).toContain('<strong>');
    expect(converted.html).toContain('<u>');
    expect(converted.html).toContain('서식 보존');
  });
});
