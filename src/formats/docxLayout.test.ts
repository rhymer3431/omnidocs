import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { extractDocxLayout } from './docxLayout';

describe('extractDocxLayout', () => {
  it('maps Word twips to HWP units and preserves plain header/footer text', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', `
      <w:document xmlns:w="w" xmlns:r="r"><w:body><w:p><w:r><w:t>본문</w:t></w:r></w:p>
      <w:sectPr>
        <w:headerReference w:type="default" r:id="rId1"/>
        <w:footerReference w:type="default" r:id="rId2"/>
        <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
        <w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>
      </w:sectPr></w:body></w:document>`);
    zip.file('word/_rels/document.xml.rels', `
      <Relationships>
        <Relationship Id="rId1" Target="header1.xml"/>
        <Relationship Id="rId2" Target="footer1.xml"/>
      </Relationships>`);
    zip.file('word/header1.xml', '<w:hdr xmlns:w="w"><w:p><w:r><w:t>회사 머리말</w:t></w:r></w:p></w:hdr>');
    zip.file('word/footer1.xml', '<w:ftr xmlns:w="w"><w:p><w:r><w:t>문서 꼬리말</w:t></w:r></w:p></w:ftr>');

    const profile = await extractDocxLayout(await zip.generateAsync({ type: 'uint8array' }));
    expect(profile.pageDef).toMatchObject({
      width: 61200,
      height: 79200,
      marginTop: 7200,
      marginRight: 5400,
      marginBottom: 7200,
      marginLeft: 5400,
      marginHeader: 3600,
      marginFooter: 3600,
      marginGutter: 0,
      landscape: false,
    });
    expect(profile.headerText).toBe('회사 머리말');
    expect(profile.footerText).toBe('문서 꼬리말');
  });

  it('preserves per-section page setup and default/even header-footer inheritance', async () => {
    const zip = new JSZip();
    zip.file('word/document.xml', `
      <w:document xmlns:w="w" xmlns:r="r"><w:body>
        <w:p><w:pPr><w:sectPr>
          <w:headerReference w:type="default" r:id="rId1"/>
          <w:headerReference w:type="even" r:id="rId2"/>
          <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
          <w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>
        </w:sectPr></w:pPr></w:p>
        <w:p><w:r><w:t>둘째 구역</w:t></w:r></w:p>
        <w:sectPr>
          <w:footerReference w:type="default" r:id="rId3"/>
          <w:type w:val="nextPage"/>
          <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
          <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/>
        </w:sectPr>
      </w:body></w:document>`);
    zip.file('word/_rels/document.xml.rels', `
      <Relationships>
        <Relationship Id="rId1" Target="header1.xml"/>
        <Relationship Id="rId2" Target="header2.xml"/>
        <Relationship Id="rId3" Target="footer1.xml"/>
      </Relationships>`);
    zip.file('word/header1.xml', '<w:hdr xmlns:w="w"><w:p><w:r><w:t>기본 머리말</w:t></w:r></w:p></w:hdr>');
    zip.file('word/header2.xml', '<w:hdr xmlns:w="w"><w:p><w:r><w:t>짝수 머리말</w:t></w:r></w:p></w:hdr>');
    zip.file('word/footer1.xml', '<w:ftr xmlns:w="w"><w:p><w:r><w:t>둘째 꼬리말</w:t></w:r></w:p></w:ftr>');

    const profile = await extractDocxLayout(await zip.generateAsync({ type: 'uint8array' }));

    expect(profile.sectionCount).toBe(2);
    expect(profile.sections[0].headers).toEqual({ default: '기본 머리말', even: '짝수 머리말' });
    expect(profile.sections[1].headers).toEqual({ default: '기본 머리말', even: '짝수 머리말' });
    expect(profile.sections[1].footers.default).toBe('둘째 꼬리말');
    expect(profile.sections[1].breakType).toBe('nextPage');
    expect(profile.sections[1].pageDef).toMatchObject({
      width: 79200,
      height: 61200,
      landscape: true,
      marginLeft: 3600,
      marginRight: 3600,
    });
  });
});
