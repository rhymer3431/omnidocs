import JSZip from 'jszip';
import type { SourceFormat } from '../domain/document';
import type { DocumentFeatureInventory } from './canonical';

export const OMDX_MIME = 'application/vnd.omnidocs.document+zip';
export const OMDX_SCHEMA = 'org.omnidocs.omdx';
export const OMDX_VERSION = 1;

export interface OmdxSourceFile {
  format: SourceFormat;
  fileName: string;
  bytes: Uint8Array;
}

export interface OmdxManifest {
  schema: typeof OMDX_SCHEMA;
  version: typeof OMDX_VERSION;
  createdAt: string;
  modifiedAt: string;
  canonical: {
    path: 'content/document.hwpx';
    format: 'hwpx';
    mediaType: string;
    sha256: string;
    engine: 'rhwp';
    engineVersion: string;
    pageCount: number;
  };
  source?: {
    path: string;
    format: SourceFormat;
    fileName: string;
    sha256: string;
    canonicalSha256AtImport: string;
  };
  compatibility: {
    importWarnings: string[];
    contentLoss?: unknown;
    layout?: unknown;
    features?: DocumentFeatureInventory;
  };
}

export interface OmdxDocument {
  manifest: OmdxManifest;
  canonicalHwpx: Uint8Array;
  source?: OmdxSourceFile;
}

export interface BuildOmdxOptions {
  canonicalHwpx: Uint8Array;
  pageCount: number;
  source?: OmdxSourceFile;
  importWarnings?: string[];
  contentLoss?: unknown;
  layout?: unknown;
  features?: DocumentFeatureInventory;
  previousManifest?: OmdxManifest;
}

export async function buildOmdx(options: BuildOmdxOptions): Promise<{ bytes: Uint8Array; document: OmdxDocument }> {
  const now = new Date().toISOString();
  const sourcePath = options.source ? `source/original.${options.source.format}` : undefined;
  const canonicalSha256 = await sha256Hex(options.canonicalHwpx);
  const manifest: OmdxManifest = {
    schema: OMDX_SCHEMA,
    version: OMDX_VERSION,
    createdAt: options.previousManifest?.createdAt ?? now,
    modifiedAt: now,
    canonical: {
      path: 'content/document.hwpx',
      format: 'hwpx',
      mediaType: 'application/vnd.hancom.hwpx+zip',
      sha256: canonicalSha256,
      engine: 'rhwp',
      engineVersion: '0.8.6',
      pageCount: options.pageCount,
    },
    source: options.source && sourcePath ? {
      path: sourcePath,
      format: options.source.format,
      fileName: options.source.fileName,
      sha256: await sha256Hex(options.source.bytes),
      canonicalSha256AtImport: options.previousManifest?.source?.canonicalSha256AtImport ?? canonicalSha256,
    } : undefined,
    compatibility: {
      importWarnings: options.importWarnings ?? options.previousManifest?.compatibility.importWarnings ?? [],
      contentLoss: options.contentLoss ?? options.previousManifest?.compatibility.contentLoss,
      layout: options.layout ?? options.previousManifest?.compatibility.layout,
      features: options.features ?? options.previousManifest?.compatibility.features,
    },
  };

  const zip = new JSZip();
  zip.file('mimetype', OMDX_MIME, { compression: 'STORE' });
  zip.file('manifest.json', JSON.stringify(manifest, null, 2));
  zip.file(manifest.canonical.path, options.canonicalHwpx);
  if (options.source && manifest.source) zip.file(manifest.source.path, options.source.bytes);

  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  return {
    bytes,
    document: {
      manifest,
      canonicalHwpx: Uint8Array.from(options.canonicalHwpx),
      source: options.source ? { ...options.source, bytes: Uint8Array.from(options.source.bytes) } : undefined,
    },
  };
}

export async function readOmdx(bytes: Uint8Array): Promise<OmdxDocument> {
  const zip = await JSZip.loadAsync(bytes);
  const mime = await zip.file('mimetype')?.async('text');
  if (mime?.trim() !== OMDX_MIME) throw new Error('유효한 OMDX 문서가 아닙니다: mimetype이 일치하지 않습니다.');

  const manifestText = await zip.file('manifest.json')?.async('text');
  if (!manifestText) throw new Error('유효한 OMDX 문서가 아닙니다: manifest.json이 없습니다.');
  const manifest = JSON.parse(manifestText) as OmdxManifest;
  if (manifest.schema !== OMDX_SCHEMA || manifest.version !== OMDX_VERSION) {
    throw new Error(`지원하지 않는 OMDX 버전입니다: ${String(manifest.version)}`);
  }

  const canonical = await zip.file(manifest.canonical.path)?.async('uint8array');
  if (!canonical) throw new Error('OMDX canonical HWPX가 없습니다.');
  if (await sha256Hex(canonical) !== manifest.canonical.sha256) {
    throw new Error('OMDX canonical HWPX 체크섬이 일치하지 않습니다.');
  }

  let source: OmdxSourceFile | undefined;
  if (manifest.source) {
    const sourceBytes = await zip.file(manifest.source.path)?.async('uint8array');
    if (sourceBytes) {
      if (await sha256Hex(sourceBytes) !== manifest.source.sha256) {
        throw new Error('OMDX 원본 파일 체크섬이 일치하지 않습니다.');
      }
      source = {
        format: manifest.source.format,
        fileName: manifest.source.fileName,
        bytes: sourceBytes,
      };
    }
  }

  return { manifest, canonicalHwpx: canonical, source };
}

export async function canUseOriginalSource(
  document: OmdxDocument | undefined,
  targetFormat: SourceFormat,
  currentCanonicalHwpx: Uint8Array,
): Promise<boolean> {
  if (!document?.source || !document.manifest.source) return false;
  if (document.source.format !== targetFormat) return false;
  return await sha256Hex(currentCanonicalHwpx) === document.manifest.source.canonicalSha256AtImport;
}

export async function sha256Hex(bytes: Uint8Array): Promise<string> {
  if (!globalThis.crypto?.subtle) throw new Error('이 환경에서는 OMDX 무결성 검사를 위한 SHA-256을 사용할 수 없습니다.');
  const stable = Uint8Array.from(bytes);
  const digest = await globalThis.crypto.subtle.digest('SHA-256', stable.buffer);
  return Array.from(new Uint8Array(digest), (value) => value.toString(16).padStart(2, '0')).join('');
}
