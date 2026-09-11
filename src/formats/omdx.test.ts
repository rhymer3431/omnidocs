import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { buildOmdx, canUseOriginalSource, OMDX_MIME, readOmdx } from './omdx';

describe('OMDX package', () => {
  it('stores a canonical HWPX and the untouched source document', async () => {
    const canonical = new Uint8Array([1, 2, 3, 4, 5, 6]);
    const source = new Uint8Array([9, 8, 7, 6]);
    const built = await buildOmdx({
      canonicalHwpx: canonical,
      pageCount: 2,
      source: { format: 'docx', fileName: '원본.docx', bytes: source },
      importWarnings: ['compatibility note'],
      layout: { paper: 'A4' },
      features: {
        sourceFormat: 'docx',
        adapter: 'omnidocs.docx-ooxml-v1',
        features: [{ id: 'field', label: 'Word 필드', count: 1, handling: 'source-only' }],
      },
    });

    const opened = await readOmdx(built.bytes);
    expect(opened.canonicalHwpx).toEqual(canonical);
    expect(opened.source?.bytes).toEqual(source);
    expect(opened.source?.fileName).toBe('원본.docx');
    expect(opened.manifest.canonical.pageCount).toBe(2);
    expect(opened.manifest.source?.canonicalSha256AtImport).toBe(opened.manifest.canonical.sha256);
    expect(opened.manifest.compatibility.importWarnings).toEqual(['compatibility note']);
    expect(opened.manifest.compatibility.features?.features[0]).toMatchObject({ id: 'field', handling: 'source-only' });

    const zip = await JSZip.loadAsync(built.bytes);
    expect(await zip.file('mimetype')?.async('text')).toBe(OMDX_MIME);
    expect(zip.file('content/document.hwpx')).not.toBeNull();
    expect(zip.file('source/original.docx')).not.toBeNull();
    expect(await canUseOriginalSource(opened, 'docx', canonical)).toBe(true);
    expect(await canUseOriginalSource(opened, 'docx', new Uint8Array([1, 2, 3]))).toBe(false);
    expect(await canUseOriginalSource(opened, 'hwp', canonical)).toBe(false);
  });

  it('rejects a package whose canonical payload was changed', async () => {
    const built = await buildOmdx({ canonicalHwpx: new Uint8Array([1, 2, 3]), pageCount: 1 });
    const zip = await JSZip.loadAsync(built.bytes);
    zip.file('content/document.hwpx', new Uint8Array([3, 2, 1]));
    const tampered = await zip.generateAsync({ type: 'uint8array' });

    await expect(readOmdx(tampered)).rejects.toThrow(/체크섬/);
  });
});
