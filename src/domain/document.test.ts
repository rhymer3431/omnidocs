import { describe, expect, it } from 'vitest';
import { formatFromFileName, replaceExtension } from './document';

describe('document helpers', () => {
  it('recognizes supported formats case-insensitively', () => {
    expect(formatFromFileName('report.HWP')).toBe('hwp');
    expect(formatFromFileName('report.hwpx')).toBe('hwpx');
    expect(formatFromFileName('report.DOCX')).toBe('docx');
    expect(formatFromFileName('report.OMDX')).toBe('omdx');
  });

  it('rejects unsupported formats', () => {
    expect(() => formatFromFileName('report.pdf')).toThrow(/OMDX/);
  });

  it('replaces the extension without changing the base name', () => {
    expect(replaceExtension('회의록.final.hwp', 'docx')).toBe('회의록.final.docx');
    expect(replaceExtension('회의록.final.docx', 'omdx')).toBe('회의록.final.omdx');
  });
});
