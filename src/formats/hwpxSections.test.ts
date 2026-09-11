import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { patchHwpxSectionPageDefs } from './hwpxSections';

describe('patchHwpxSectionPageDefs', () => {
  it('patches logical HWPX secPr blocks with per-section Word page definitions', async () => {
    const zip = new JSZip();
    zip.file('Contents/section0.xml', `
      <hs:sec xmlns:hs="hs" xmlns:hp="hp">
        <hp:p><hp:run><hp:secPr><hp:pagePr landscape="WIDELY" width="60000" height="80000"><hp:margin left="5000" right="5000" top="5000" bottom="5000" header="3000" footer="3000" gutter="0"/></hp:pagePr></hp:secPr></hp:run></hp:p>
        <hp:p><hp:run><hp:secPr><hp:pagePr landscape="WIDELY" width="60000" height="80000"><hp:margin left="5000" right="5000" top="5000" bottom="5000" header="3000" footer="3000" gutter="0"/></hp:pagePr></hp:secPr></hp:run></hp:p>
      </hs:sec>`, { compression: 'DEFLATE' });
    const source = await zip.generateAsync({ type: 'uint8array' });

    const patched = await patchHwpxSectionPageDefs(source, [
      { width: 61200, height: 79200, landscape: false, marginLeft: 7200, marginRight: 7200 },
      { width: 79200, height: 61200, landscape: true, marginLeft: 3600, marginRight: 3600, marginTop: 3600, marginBottom: 3600 },
    ]);
    const resultZip = await JSZip.loadAsync(patched.bytes);
    const xml = await resultZip.file('Contents/section0.xml')!.async('text');

    expect(patched.patchedSections).toBe(2);
    expect(xml).toContain('landscape="WIDELY" width="61200" height="79200"');
    expect(xml).toContain('landscape="NARROWLY" width="79200" height="61200"');
    expect(xml).toContain('left="3600" right="3600" top="3600" bottom="3600"');
  });
});
