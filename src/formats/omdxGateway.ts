import type { OmniFile, SourceFormat } from '../domain/document';
import { buildOmdx, readOmdx, type OmdxDocument, type OmdxSourceFile } from './omdx';
import { importSourceToCanonical } from './sourceAdapters';

export interface OmdxImportResult {
  document: OmdxDocument;
  packageBytes: Uint8Array;
  warnings: string[];
}

export async function importAsOmdx(file: OmniFile): Promise<OmdxImportResult> {
  if (file.format === 'omdx') {
    const document = await readOmdx(file.bytes);
    return {
      document,
      packageBytes: Uint8Array.from(file.bytes),
      warnings: [...document.manifest.compatibility.importWarnings],
    };
  }

  const source: OmdxSourceFile = {
    format: file.format as SourceFormat,
    fileName: file.name,
    bytes: Uint8Array.from(file.bytes),
  };

  const imported = await importSourceToCanonical(file.format as SourceFormat, file.bytes);
  const built = await buildOmdx({
    canonicalHwpx: imported.canonicalHwpx,
    pageCount: imported.pageCount,
    source,
    importWarnings: imported.warnings,
    contentLoss: imported.contentLoss,
    layout: imported.layout,
    features: imported.features,
  });
  return { document: built.document, packageBytes: built.bytes, warnings: imported.warnings };
}
