import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { analyzeDocxFeatures } from './docxFeatures';

describe('analyzeDocxFeatures', () => {
  it('classifies native, approximated, and source-only DOCX features', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', `
      <w:document xmlns:w="w" xmlns:m="m"><w:body>
        <w:p><w:r><w:t>A</w:t></w:r></w:p>
        <w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
        <w:p><w:r><w:drawing/></w:r></w:p>
        <w:p><w:r><w:instrText>PAGE</w:instrText></w:r></w:p>
        <w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt>
        <w:p><w:ins><w:r><w:t>B</w:t></w:r></w:ins></w:p>
        <w:p><m:oMath/></w:p>
        <w:sectPr/>
      </w:body></w:document>`);
    zip.file('word/header1.xml', '<w:hdr xmlns:w="w"><w:p/></w:hdr>');
    zip.file('word/footer1.xml', '<w:ftr xmlns:w="w"><w:p/></w:ftr>');

    const inventory = await analyzeDocxFeatures(await zip.generateAsync({ type: 'uint8array' }));
    const feature = (id: string) => inventory.features.find((item) => item.id === id);

    expect(feature('table')).toMatchObject({ count: 1, handling: 'native' });
    expect(feature('drawing')).toMatchObject({ count: 1, handling: 'approximated' });
    expect(feature('field')).toMatchObject({ count: 1, handling: 'source-only' });
    expect(feature('content-control')).toMatchObject({ count: 1, handling: 'source-only' });
    expect(feature('track-change')).toMatchObject({ count: 1, handling: 'source-only' });
    expect(feature('equation')).toMatchObject({ count: 1, handling: 'approximated' });
    expect(feature('header-part')).toMatchObject({ count: 1, handling: 'normalized' });
    expect(feature('footer-part')).toMatchObject({ count: 1, handling: 'normalized' });
  });
});
