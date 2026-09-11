import JSZip from 'jszip';
import { JSDOM } from 'jsdom';
import mammoth from 'mammoth/mammoth.browser';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { exportHwpxToDocx } from './hwpxDocxWriter';

const originalDOMParser = globalThis.DOMParser;

const ONE_PIXEL_PNG = Uint8Array.from([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
  0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196, 137,
  0, 0, 0, 13, 73, 68, 65, 84, 120, 218, 99, 252, 207, 192, 80,
  15, 0, 5, 254, 2, 254, 227, 21, 199, 205, 0, 0, 0, 0, 73, 69,
  78, 68, 174, 66, 96, 130,
]);

beforeAll(() => {
  const dom = new JSDOM();
  Object.assign(globalThis, { DOMParser: dom.window.DOMParser });
});

afterAll(() => {
  Object.assign(globalThis, { DOMParser: originalDOMParser });
});

describe('direct HWPX -> DOCX writer', () => {
  it('writes semantic OOXML for text, formatting, table, image, header/footer and page geometry', async () => {
    const hwpx = new JSZip();
    hwpx.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
      <hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
        <hh:refList>
          <hh:fontfaces><hh:fontface lang="HANGUL"><hh:font id="0" face="Malgun Gothic"/></hh:fontface><hh:fontface lang="LATIN"><hh:font id="0" face="Arial"/></hh:fontface></hh:fontfaces>
          <hh:borderFills><hh:borderFill id="1"><hh:leftBorder type="SOLID" width="0.4 mm" color="#112233"/><hh:rightBorder type="SOLID" width="0.4 mm" color="#112233"/><hh:topBorder type="SOLID" width="0.4 mm" color="#112233"/><hh:bottomBorder type="SOLID" width="0.4 mm" color="#112233"/></hh:borderFill></hh:borderFills>
          <hh:charProperties><hh:charPr id="1" height="1400" textColor="#336699"><hh:fontRef hangul="0" latin="0"/><hh:bold/><hh:italic/><hh:underline type="BOTTOM" shape="SOLID" color="#336699"/></hh:charPr></hh:charProperties>
          <hh:paraProperties><hh:paraPr id="1"><hh:align horizontal="CENTER"/><hh:heading type="NUMBER" idRef="1" level="0"/><hh:breakSetting widowOrphan="1" keepWithNext="1" keepLines="0" pageBreakBefore="0"/><hp:switch><hp:default><hh:margin><hc:intent value="900"/><hc:left value="1800"/><hc:right value="0"/><hc:prev value="600"/><hc:next value="300"/></hh:margin><hh:lineSpacing type="PERCENT" value="150"/></hp:default></hp:switch></hh:paraPr></hh:paraProperties>
        </hh:refList>
      </hh:head>`);
    hwpx.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?><opf:package xmlns:opf="http://www.idpf.org/2007/opf"><opf:manifest><opf:item id="image1" href="BinData/image1.png" media-type="image/png"/></opf:manifest></opf:package>`);
    hwpx.file('BinData/image1.png', Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]));
    hwpx.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
      <hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
        <hp:p paraPrIDRef="1"><hp:run charPrIDRef="1">
          <hp:secPr><hp:pagePr landscape="WIDELY" width="59528" height="84188"><hp:margin left="8504" right="8504" top="5669" bottom="4252" header="4252" footer="4252" gutter="0"/></hp:pagePr></hp:secPr>
          <hp:ctrl><hp:colPr colCount="1" sameGap="0"/></hp:ctrl>
          <hp:t>직접 저장</hp:t>
          <hp:ctrl><hp:header applyPageType="BOTH"><hp:subList><hp:p paraPrIDRef="1"><hp:run charPrIDRef="1"><hp:t>머리말</hp:t></hp:run></hp:p></hp:subList></hp:header></hp:ctrl>
          <hp:ctrl><hp:footer applyPageType="BOTH"><hp:subList><hp:p paraPrIDRef="1"><hp:run charPrIDRef="1"><hp:t>꼬리말</hp:t></hp:run></hp:p></hp:subList></hp:footer></hp:ctrl>
        </hp:run></hp:p>
        <hp:p paraPrIDRef="1"><hp:run charPrIDRef="1"><hp:t>[이미지]</hp:t><hp:pic zOrder="3" textWrap="SQUARE" textFlow="RIGHT_ONLY"><hp:curSz width="9000" height="4500"/><hp:pos treatAsChar="0" horzRelTo="PAPER" vertRelTo="PAPER" horzAlign="LEFT" vertAlign="TOP" horzOffset="1270" vertOffset="2540" allowOverlap="1"/><hp:outMargin left="100" right="200" top="300" bottom="400"/><hc:img binaryItemIDRef="image1"/></hp:pic></hp:run></hp:p>
        <hp:p paraPrIDRef="1"><hp:run charPrIDRef="1"><hp:equation textColor="#112233" baseUnit="1200"><hp:script>{a over b} + sqrt {x} + x^2 + sum from {i=1} to n {i}</hp:script><hp:pos treatAsChar="1"/></hp:equation></hp:run></hp:p>
        <hp:p paraPrIDRef="1"><hp:run charPrIDRef="1"><hp:rect zOrder="4" textWrap="SQUARE" textFlow="BOTH_SIDES" ratio="0"><hp:curSz width="12000" height="6000"/><hp:lineShape color="#112233" width="50" style="DASH" headStyle="NORMAL" tailStyle="NORMAL"/><hc:fillBrush><hc:winBrush faceColor="#DDEEFF"/></hc:fillBrush><hp:pos treatAsChar="0" horzRelTo="PAPER" vertRelTo="PAPER" horzAlign="LEFT" vertAlign="TOP" horzOffset="1800" vertOffset="2400" allowOverlap="0"/><hp:shapeComment>테스트 사각형</hp:shapeComment></hp:rect></hp:run></hp:p>
        <hp:p paraPrIDRef="1"><hp:run charPrIDRef="1"><hp:t>주석 본문</hp:t><hp:ctrl><hp:footNote number="1"><hp:subList><hp:p paraPrIDRef="1"><hp:run charPrIDRef="1"><hp:t>각주 내용</hp:t></hp:run></hp:p></hp:subList></hp:footNote></hp:ctrl><hp:ctrl><hp:endNote number="1"><hp:subList><hp:p paraPrIDRef="1"><hp:run charPrIDRef="1"><hp:t>미주 내용</hp:t></hp:run></hp:p></hp:subList></hp:endNote></hp:ctrl></hp:run></hp:p>
        <hp:p paraPrIDRef="1"><hp:run charPrIDRef="1"><hp:tbl rowCnt="2" colCnt="2" borderFillIDRef="1"><hp:sz width="0" height="0"/><hp:pos treatAsChar="0" horzRelTo="PAPER" vertRelTo="PAPER" horzAlign="CENTER" vertAlign="BOTTOM" horzOffset="0" vertOffset="0" allowOverlap="0"/><hp:outMargin left="200" right="200" top="100" bottom="100"/>
          <hp:tr><hp:tc borderFillIDRef="1"><hp:subList vertAlign="CENTER"><hp:p paraPrIDRef="1"><hp:run charPrIDRef="1"><hp:t>A</hp:t></hp:run></hp:p></hp:subList><hp:cellAddr rowAddr="0" colAddr="0"/><hp:cellSpan rowSpan="2" colSpan="1"/><hp:cellSz width="10000" height="2000"/><hp:cellMargin left="100" right="100" top="100" bottom="100"/></hp:tc><hp:tc borderFillIDRef="1"><hp:subList><hp:p paraPrIDRef="1"><hp:run charPrIDRef="1"><hp:t>B</hp:t></hp:run></hp:p></hp:subList><hp:cellAddr rowAddr="0" colAddr="1"/><hp:cellSpan rowSpan="1" colSpan="1"/><hp:cellSz width="12000" height="1000"/></hp:tc></hp:tr>
          <hp:tr><hp:tc borderFillIDRef="1"><hp:subList><hp:p paraPrIDRef="1"><hp:run charPrIDRef="1"><hp:t>D</hp:t></hp:run></hp:p></hp:subList><hp:cellAddr rowAddr="1" colAddr="1"/><hp:cellSpan rowSpan="1" colSpan="1"/><hp:cellSz width="12000" height="1000"/></hp:tc></hp:tr>
        </hp:tbl></hp:run></hp:p>
      </hs:sec>`);

    const result = await exportHwpxToDocx(await hwpx.generateAsync({ type: 'uint8array' }));
    const docx = await JSZip.loadAsync(result.bytes);
    const documentXml = await docx.file('word/document.xml')!.async('text');
    const rels = await docx.file('word/_rels/document.xml.rels')!.async('text');
    const numbering = await docx.file('word/numbering.xml')!.async('text');
    const header = await docx.file('word/header1.xml')!.async('text');
    const footer = await docx.file('word/footer1.xml')!.async('text');
    const footnotes = await docx.file('word/footnotes.xml')!.async('text');
    const endnotes = await docx.file('word/endnotes.xml')!.async('text');

    expect(documentXml).toContain('직접 저장');
    expect(documentXml).toContain('<w:b/>');
    expect(documentXml).toContain('<w:i/>');
    expect(documentXml).toContain('<w:sz w:val="28"/>');
    expect(documentXml).toContain('<w:color w:val="336699"/>');
    expect(documentXml).toContain('<w:jc w:val="center"/>');
    expect(documentXml).toContain('<w:pgSz w:w="11906" w:h="16838"/>');
    expect(documentXml).toContain('<w:tbl>');
    expect(documentXml).toContain('<w:vMerge w:val="restart"/>');
    expect(documentXml).toContain('<w:vMerge/>');
    expect(documentXml).toContain('<w:drawing>');
    expect(documentXml).toContain('<wp:anchor');
    expect(documentXml).toContain('<wp:positionH relativeFrom="page"><wp:posOffset>161290</wp:posOffset></wp:positionH>');
    expect(documentXml).toContain('<wp:positionV relativeFrom="page"><wp:posOffset>322580</wp:posOffset></wp:positionV>');
    expect(documentXml).toContain('<wp:wrapSquare wrapText="right"/>');
    expect(documentXml).toContain('r:embed="rIdImage1"');
    expect(documentXml).toContain('<w:tblpPr w:horzAnchor="page" w:vertAnchor="page"');
    expect(documentXml).toContain('w:tblpXSpec="center"');
    expect(documentXml).toContain('w:tblpYSpec="bottom"');
    expect(documentXml).toContain('<w:tblOverlap w:val="never"/>');
    expect(documentXml).toContain('<m:oMath>');
    expect(documentXml).toContain('<m:f>');
    expect(documentXml).toContain('<m:rad>');
    expect(documentXml).toContain('<m:sSup>');
    expect(documentXml).toContain('<m:nary>');
    expect(documentXml).toContain('<m:chr m:val="∑"/>');
    expect(documentXml).toContain('<wps:wsp>');
    expect(documentXml).toContain('<a:prstGeom prst="rect">');
    expect(documentXml).toContain('<a:srgbClr val="DDEEFF"/>');
    expect(documentXml).toContain('descr="테스트 사각형"');
    expect(documentXml).toContain('<w:footnoteReference w:id="1"/>');
    expect(documentXml).toContain('<w:endnoteReference w:id="1"/>');
    expect(numbering).toContain('<w:numFmt w:val="decimal"/>');
    expect(header).toContain('머리말');
    expect(footer).toContain('꼬리말');
    expect(rels).toContain('Target="media/image1.png"');
    expect(rels).toContain('Target="header1.xml"');
    expect(rels).toContain('Target="footer1.xml"');
    expect(rels).toContain('Target="footnotes.xml"');
    expect(rels).toContain('Target="endnotes.xml"');
    expect(footnotes).toContain('각주 내용');
    expect(footnotes).toContain('<w:footnoteRef/>');
    expect(endnotes).toContain('미주 내용');
    expect(endnotes).toContain('<w:endnoteRef/>');
    expect(docx.file('word/media/image1.png')).not.toBeNull();
    expect(result.warnings).toContain('HWPX의 사용자 정의 문단 번호/글머리표 모양은 DOCX 다단계 목록 규칙으로 정규화됩니다.');

    const arrayBuffer = result.bytes.buffer.slice(
      result.bytes.byteOffset,
      result.bytes.byteOffset + result.bytes.byteLength,
    ) as ArrayBuffer;
    const reopened = await mammoth.convertToHtml({ arrayBuffer });
    expect(reopened.value).toContain('직접 저장');
    expect(reopened.value).toContain('<strong><em>직접 저장</em></strong>');
    expect(reopened.value).toContain('<table>');
    expect(reopened.value).toContain('A');
    expect(reopened.value).toContain('D');
  });

  it('accepts the canonical HWPX emitted by the real rHWP engine', async () => {
    const [{ readFile }, rhwp] = await Promise.all([
      import('node:fs/promises'),
      import('@rhwp/core'),
    ]);
    const wasm = await readFile(new URL('../../public/vendor/rhwp/rhwp_bg.wasm', import.meta.url));
    await rhwp.default({ module_or_path: wasm });
    Object.assign(globalThis, { measureTextWidth: (_font: string, text: string) => text.length * 8 });

    const document = rhwp.HwpDocument.createEmpty();
    let poisoned = false;
    try {
      document.createBlankDocument();
      const pasted = JSON.parse(document.pasteHtml(
        0,
        0,
        0,
        '<p><b>실제 canonical HWPX</b></p><table><tr><td>항목</td><td>값</td></tr><tr><td>테스트</td><td>42</td></tr></table>',
      )) as { ok?: boolean; error?: string };
      expect(pasted.ok).not.toBe(false);

      const equation = JSON.parse(document.insertEquation(
        0,
        0,
        0,
        '{a over b} + sqrt {x} + x^2 + sum from {i=1} to n {i}',
        1200,
        0x112233,
      )) as { ok?: boolean; error?: string };
      expect(equation.ok).not.toBe(false);

      const picture = JSON.parse(document.insertPicture(
        0,
        0,
        0,
        '',
        ONE_PIXEL_PNG,
        9000,
        4500,
        1,
        1,
        'png',
        'floating test',
        1200,
        2400,
      )) as { ok?: boolean; error?: string };
      expect(picture.ok).not.toBe(false);

      const shape = JSON.parse(document.createShapeControl(JSON.stringify({
        sectionIdx: 0,
        paraIdx: 0,
        charOffset: 0,
        width: 12000,
        height: 6000,
        horzOffset: 1800,
        vertOffset: 2400,
        treatAsChar: false,
        textWrap: 'Square',
      }))) as { ok?: boolean; error?: string };
      expect(shape.ok).not.toBe(false);

      const footnote = JSON.parse(document.insertFootnote(0, 0, 0)) as {
        ok?: boolean;
        error?: string;
        controlIdx: number;
      };
      expect(footnote.ok).not.toBe(false);
      document.insertTextInFootnote(0, 0, footnote.controlIdx, 0, 0, '실제 각주');
      const endnote = JSON.parse(document.insertEndnote(0, 0, 0)) as { ok?: boolean; error?: string };
      expect(endnote.ok).not.toBe(false);

      const exported = await exportHwpxToDocx(document.exportHwpx());
      const packageZip = await JSZip.loadAsync(exported.bytes);
      const documentXml = await packageZip.file('word/document.xml')!.async('text');
      const footnotesXml = await packageZip.file('word/footnotes.xml')!.async('text');
      const endnotesXml = await packageZip.file('word/endnotes.xml')!.async('text');
      const arrayBuffer = exported.bytes.buffer.slice(
        exported.bytes.byteOffset,
        exported.bytes.byteOffset + exported.bytes.byteLength,
      ) as ArrayBuffer;
      const reopened = await mammoth.convertToHtml({ arrayBuffer });

      expect(reopened.value).toContain('실제 canonical HWPX');
      expect(reopened.value).toContain('<strong>');
      expect(reopened.value).toContain('<table>');
      expect(reopened.value).toContain('테스트');
      expect(reopened.value).toContain('42');
      expect(documentXml).toContain('<wp:anchor');
      expect(documentXml).toContain('<wp:posOffset>152400</wp:posOffset>');
      expect(documentXml).toContain('<wp:posOffset>304800</wp:posOffset>');
      expect(documentXml).toContain('<m:f>');
      expect(documentXml).toContain('<m:rad>');
      expect(documentXml).toContain('<m:sSup>');
      expect(documentXml).toContain('<m:nary>');
      expect(documentXml).toContain('<wps:wsp>');
      expect(documentXml).toContain('<w:footnoteReference');
      expect(documentXml).toContain('<w:endnoteReference');
      expect(footnotesXml).toContain('실제 각주');
      expect(endnotesXml).toContain('<w:endnoteRef/>');
    } catch (error) {
      poisoned = error instanceof WebAssembly.RuntimeError;
      throw error;
    } finally {
      if (!poisoned) document.free();
    }
  });
});
