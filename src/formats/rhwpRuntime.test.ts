import { JSDOM } from 'jsdom';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { pasteTopLevelHtmlBlocks, splitTopLevelHtmlBlocks } from './rhwpRuntime';

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
});
