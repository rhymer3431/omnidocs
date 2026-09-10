import { readFile } from 'node:fs/promises';
import init, { HwpDocument } from '@rhwp/core';

globalThis.measureTextWidth = (_font, text) => text.length * 8;

const wasmUrl = new URL('../public/vendor/rhwp/rhwp_bg.wasm', import.meta.url);
const wasm = await readFile(wasmUrl);
await init({ module_or_path: wasm });

const document = HwpDocument.createEmpty();
try {
  const pasted = JSON.parse(document.pasteHtml(0, 0, 0, '<p>OmniDocs 포맷 브리지 테스트</p><table><tr><td>A</td><td>B</td></tr></table>'));
  if (pasted.ok === false) throw new Error(pasted.error ?? 'pasteHtml failed');

  const hwpx = document.exportHwpx();
  const hwp = document.exportHwp();
  if (hwpx.byteLength < 100 || hwp.byteLength < 100) {
    throw new Error(`unexpected export size: hwpx=${hwpx.byteLength}, hwp=${hwp.byteLength}`);
  }

  console.log(`rHWP smoke OK: HWPX=${hwpx.byteLength} bytes, HWP=${hwp.byteLength} bytes`);
} finally {
  document.free();
}
