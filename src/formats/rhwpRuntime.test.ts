import { JSDOM } from 'jsdom';
import { afterEach, describe, expect, it } from 'vitest';
import { splitTopLevelHtmlBlocks } from './rhwpRuntime';

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
});
