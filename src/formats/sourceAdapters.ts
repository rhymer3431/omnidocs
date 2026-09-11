import type { SourceFormat } from '../domain/document';
import type { CanonicalDocumentPayload, DocumentFeatureInventory, SourceFormatAdapter } from './canonical';
import { sourceOnlyFeatureWarnings } from './canonical';
import { importDocx } from './docxAdapter';
import { openRhwp } from './rhwpRuntime';

const docxAdapter: SourceFormatAdapter = {
  format: 'docx',
  id: 'omnidocs.docx-ooxml-v1',
  async import(bytes) {
    const imported = await importDocx(bytes);
    const pageCount = await pageCountOf(imported.hwpxBytes);
    return {
      canonicalHwpx: imported.hwpxBytes,
      pageCount,
      warnings: dedupeWarnings([...imported.warnings, ...sourceOnlyFeatureWarnings(imported.features)]),
      layout: imported.layout,
      features: imported.features,
    };
  },
};

function rhwpAdapter(format: 'hwp' | 'hwpx'): SourceFormatAdapter {
  return {
    format,
    id: `omnidocs.${format}-rhwp-v1`,
    async import(bytes) {
      const document = await openRhwp(bytes);
      try {
        const pageCount = document.pageCount();
        const exported = document.exportHwpxWithReport();
        try {
          const contentLoss = parseJson(exported.contentLoss());
          const canonicalHwpx = exported.takeBytes();
          const warnings = contentLossWarnings(contentLoss);
          const layout = collectRhwpLayout(document);
          return {
            canonicalHwpx,
            pageCount,
            warnings,
            contentLoss,
            layout,
            features: rhwpFeatureInventory(format, document.getSectionCount()),
          };
        } finally {
          exported.free();
        }
      } finally {
        document.free();
      }
    },
  };
}

const adapters: Record<SourceFormat, SourceFormatAdapter> = {
  hwp: rhwpAdapter('hwp'),
  hwpx: rhwpAdapter('hwpx'),
  docx: docxAdapter,
};

export function getSourceFormatAdapter(format: SourceFormat): SourceFormatAdapter {
  return adapters[format];
}

export async function importSourceToCanonical(format: SourceFormat, bytes: Uint8Array): Promise<CanonicalDocumentPayload> {
  return getSourceFormatAdapter(format).import(bytes);
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

function rhwpFeatureInventory(format: 'hwp' | 'hwpx', sectionCount: number): DocumentFeatureInventory {
  return {
    sourceFormat: format,
    adapter: `omnidocs.${format}-rhwp-v1`,
    features: [
      { id: 'section', label: '구역', count: sectionCount, handling: 'native' },
    ],
  };
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

function dedupeWarnings(warnings: string[]): string[] {
  return [...new Set(warnings.filter(Boolean))];
}
