import init, { HwpDocument } from '@rhwp/core';
import mammoth from 'mammoth';
import {
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
} from 'docx';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { JSDOM } from 'jsdom';

globalThis.measureTextWidth ??= (_font, text) => text.length * 8;
const wasmUrl = new URL('../public/vendor/rhwp/rhwp_bg.wasm', import.meta.url);
const wasm = await readFile(wasmUrl);
await init({ module_or_path: wasm });

const fixture = new Document({
  sections: [{
    children: [
      new Paragraph({ text: 'OmniDocs DOCX compatibility', heading: HeadingLevel.HEADING_1 }),
      new Paragraph({
        children: [
          new TextRun({ text: '굵은 글씨', bold: true }),
          new TextRun('와 일반 문단입니다.'),
        ],
      }),
      new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph('항목')] }),
              new TableCell({ children: [new Paragraph('값')] }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph('테스트')] }),
              new TableCell({ children: [new Paragraph('42')] }),
            ],
          }),
        ],
      }),
    ],
  }],
});

const sourceBytes = new Uint8Array(await Packer.toBuffer(fixture));
const converted = await mammoth.convertToHtml({ buffer: Buffer.from(sourceBytes) });
assert.match(converted.value, /OmniDocs DOCX compatibility/);
assert.match(converted.value, /<strong>굵은 글씨<\/strong>/);
assert.match(converted.value, /<table>/);
assert.match(converted.value, /테스트/);

const doc = HwpDocument.createEmpty();
let poisoned = false;
try {
  const dom = new JSDOM(`<body>${converted.value}</body>`);
  const blocks = Array.from(dom.window.document.body.children, (node) => node.outerHTML);
  let paraIdx = 0;
  let charOffset = 0;
  for (const block of blocks) {
    const paste = JSON.parse(doc.pasteHtml(0, paraIdx, charOffset, block));
    assert.notEqual(paste.ok, false, paste.error ?? 'pasteHtml failed');
    paraIdx = paste.paraIdx ?? paraIdx;
    charOffset = paste.charOffset ?? charOffset;
  }

  const hwpx = doc.exportHwpx();
  assert.ok(hwpx.byteLength > 1000, 'HWPX export is unexpectedly small');
  assert.ok(doc.pageCount() >= 1, 'Imported document has no pages');

  console.log(`DOCX import smoke OK: source=${sourceBytes.byteLength} bytes, html=${converted.value.length} chars, hwpx=${hwpx.byteLength} bytes`);
} catch (error) {
  // A WASM panic can leave the instance borrowed. Avoid masking the actual failure with free().
  poisoned = true;
  throw error;
} finally {
  if (!poisoned) doc.free();
}
