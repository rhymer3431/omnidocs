import { JSDOM } from 'jsdom';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { pasteTopLevelHtmlBlocks, splitTopLevelHtmlBlocks } from './rhwpRuntime';
import { encodeOmdxFormat, OMDX_CHAR_FORMAT_ATTR, OMDX_PARA_FORMAT_ATTR, OMDX_SECTION_BREAK_ATTR } from './docxOoxml';

const originalDOMParser = globalThis.DOMParser;
const originalNode = globalThis.Node;

afterEach(() => {
  Object.assign(globalThis, { DOMParser: originalDOMParser, Node: originalNode });
});

describe('splitTopLevelHtmlBlocks', () => {
  it('keeps nested table markup together while separating top-level blocks', () => {
    const dom = new JSDOM();
    Object.assign(globalThis, {
      DOMParser: dom.window.DOMParser,
      Node: dom.window.Node,
    });

    const blocks = splitTopLevelHtmlBlocks(
      '<h1>제목</h1><p>본문</p><table><tr><td><p>셀</p></td></tr></table>',
    );

    expect(blocks).toHaveLength(3);
    expect(blocks[0]).toBe('<h1>제목</h1>');
    expect(blocks[1]).toBe('<p>본문</p>');
    expect(blocks[2]).toContain('<table>');
    expect(blocks[2]).toContain('셀');
  });

  it('creates an explicit paragraph boundary between sequential pasteHtml blocks', () => {
    const dom = new JSDOM();
    Object.assign(globalThis, {
      DOMParser: dom.window.DOMParser,
      Node: dom.window.Node,
    });

    const pasteHtml = vi.fn()
      .mockReturnValueOnce(JSON.stringify({ ok: true, paraIdx: 0, charOffset: 1 }))
      .mockReturnValueOnce(JSON.stringify({ ok: true, paraIdx: 1, charOffset: 1 }))
      .mockReturnValueOnce(JSON.stringify({ ok: true, paraIdx: 2, charOffset: 0 }));
    const insertParagraph = vi.fn()
      .mockReturnValueOnce(JSON.stringify({ ok: true, paraIdx: 1, newParagraphCount: 2 }))
      .mockReturnValueOnce(JSON.stringify({ ok: true, paraIdx: 2, newParagraphCount: 3 }));

    pasteTopLevelHtmlBlocks(
      { pasteHtml, insertParagraph } as never,
      '<p>A</p><p>B</p><table><tr><td><p>C</p></td></tr></table>',
    );

    expect(pasteHtml).toHaveBeenNthCalledWith(1, 0, 0, 0, '<p>A</p>');
    expect(insertParagraph).toHaveBeenNthCalledWith(1, 0, 1);
    expect(pasteHtml).toHaveBeenNthCalledWith(2, 0, 1, 0, '<p>B</p>');
    expect(insertParagraph).toHaveBeenNthCalledWith(2, 0, 2);
    expect(pasteHtml).toHaveBeenNthCalledWith(3, 0, 2, 0, '<table><tbody><tr><td><p>C</p></td></tr></tbody></table>');
  });

  it('applies DOCX-only paragraph and character metadata through native rHWP APIs', () => {
    const dom = new JSDOM();
    Object.assign(globalThis, {
      DOMParser: dom.window.DOMParser,
      Node: dom.window.Node,
    });

    const paragraphFormat = encodeOmdxFormat({ marginLeft: 3600, indent: 1800, spacingAfter: 600 });
    const characterFormat = encodeOmdxFormat({ fontName: 'Malgun Gothic', fontSize: 1400, textColor: '#FF0000', superscript: true });
    const pasteHtml = vi.fn().mockReturnValue(JSON.stringify({ ok: true, paraIdx: 0, charOffset: 2 }));
    const applyParaFormat = vi.fn().mockReturnValue(JSON.stringify({ ok: true }));
    const applyCharFormat = vi.fn().mockReturnValue(JSON.stringify({ ok: true }));
    const findOrCreateFontId = vi.fn().mockReturnValue(7);

    pasteTopLevelHtmlBlocks({
      pasteHtml,
      insertParagraph: vi.fn(),
      applyParaFormat,
      applyCharFormat,
      findOrCreateFontId,
      breakAtCursor: vi.fn(),
    } as never, `<p ${OMDX_PARA_FORMAT_ATTR}="${paragraphFormat}"><span ${OMDX_CHAR_FORMAT_ATTR}="${characterFormat}">AB</span></p>`);

    expect(applyParaFormat).toHaveBeenCalledTimes(1);
    expect(JSON.parse(applyParaFormat.mock.calls[0][2])).toMatchObject({ marginLeft: 3600, indent: 1800, spacingAfter: 600 });
    expect(findOrCreateFontId).toHaveBeenCalledWith('Malgun Gothic');
    expect(applyCharFormat).toHaveBeenCalledWith(0, 0, 0, 2, expect.any(String));
    expect(JSON.parse(applyCharFormat.mock.calls[0][4])).toMatchObject({
      fontId: 7,
      fontSize: 1400,
      textColor: '#FF0000',
      superscript: true,
    });
  });

  it('turns an OOXML section boundary into an rHWP logical section break', () => {
    const dom = new JSDOM();
    Object.assign(globalThis, {
      DOMParser: dom.window.DOMParser,
      Node: dom.window.Node,
    });

    const pasteHtml = vi.fn()
      .mockReturnValueOnce(JSON.stringify({ ok: true, paraIdx: 0, charOffset: 1 }))
      .mockReturnValueOnce(JSON.stringify({ ok: true, paraIdx: 1, charOffset: 17 }));
    const insertParagraph = vi.fn().mockReturnValue(JSON.stringify({ ok: true, paraIdx: 1 }));
    const breakAtCursor = vi.fn().mockReturnValue(JSON.stringify({ ok: true, para: 1, pos: 16 }));

    pasteTopLevelHtmlBlocks({
      pasteHtml,
      insertParagraph,
      applyParaFormat: vi.fn(),
      applyCharFormat: vi.fn(),
      findOrCreateFontId: vi.fn(),
      breakAtCursor,
    } as never, `<p ${OMDX_SECTION_BREAK_ATTR}="1">A</p><p>B</p>`);

    expect(insertParagraph).toHaveBeenCalledWith(0, 1);
    expect(breakAtCursor).toHaveBeenCalledWith(0, 1, 0, 'section');
    expect(pasteHtml).toHaveBeenNthCalledWith(2, 0, 1, 16, '<p>B</p>');
  });
});
