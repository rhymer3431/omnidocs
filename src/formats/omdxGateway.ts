import type { OmniFile, SourceFormat } from '../domain/document';
import { importDocx } from './docxAdapter';
import { buildOmdx, readOmdx, type OmdxDocument, type OmdxSourceFile } from './omdx';
import { openRhwp } from './rhwpRuntime';

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

  if (file.format === 'docx') {
    const imported = await importDocx(file.bytes);
    const pageCount = await pageCountOf(imported.hwpxBytes);
    const built = await buildOmdx({
      canonicalHwpx: imported.hwpxBytes,
      pageCount,
      source,
      importWarnings: imported.warnings,
      layout: imported.layout,
    });
    return { document: built.document, packageBytes: built.bytes, warnings: imported.warnings };
  }

  const document = await openRhwp(file.bytes);
  try {
    const pageCount = document.pageCount();
    const exported = document.exportHwpxWithReport();
    try {
      const contentLoss = parseJson(exported.contentLoss());
      const canonicalHwpx = exported.takeBytes();
      const warnings = contentLossWarnings(contentLoss);
      const layout = collectRhwpLayout(document);
      const built = await buildOmdx({
        canonicalHwpx,
        pageCount,
        source,
        importWarnings: warnings,
        contentLoss,
        layout,
      });
      return { document: built.document, packageBytes: built.bytes, warnings };
    } finally {
      exported.free();
    }
  } finally {
    document.free();
  }
}

async function pageCountOf(bytes: Uint8Array): Promise<number> {
  const document = await openRhwp(bytes);
  try {
    return document.pageCount();
  } finally {
    document.free();
  }
}

function collectRhwpLayout(document: Awaited<ReturnType<typeof openRhwp>>): unknown {
  const sections = [];
  for (let index = 0; index < document.getSectionCount(); index += 1) {
    sections.push({
      pageDef: parseJson(document.getPageDef(index)),
      sectionDef: parseJson(document.getSectionDef(index)),
    });
  }
  return { sections };
}

function contentLossWarnings(report: unknown): string[] {
  if (!report || typeof report !== 'object') return [];
  const record = report as { count?: number; losses?: Array<{ kind?: string; message?: string }> };
  if (!record.count) return [];
  const details = (record.losses ?? []).slice(0, 5).map((loss) => loss.message ?? loss.kind ?? 'unknown loss');
  return [`OMDX 정규화 중 ${record.count}개의 호환성 손실이 보고되었습니다.`, ...details];
}

function parseJson(value: string): unknown {
  try {
    return JSON.parse(value);
  } catch {
    return value;
  }
}
