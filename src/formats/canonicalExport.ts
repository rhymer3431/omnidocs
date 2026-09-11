import type { SourceFormat } from '../domain/document';
import { exportHwpxToDocx } from './hwpxDocxWriter';
import { openRhwp } from './rhwpRuntime';

export interface CanonicalExportPayload {
  bytes: Uint8Array;
  warnings: string[];
}

type CanonicalExporter = (canonicalHwpx: Uint8Array) => Promise<CanonicalExportPayload>;

const exporters: Record<SourceFormat, CanonicalExporter> = {
  hwpx: async (canonicalHwpx) => ({
    bytes: Uint8Array.from(canonicalHwpx),
    warnings: [],
  }),
  docx: exportHwpxToDocx,
  hwp: exportHwp,
};

export async function exportCanonical(format: SourceFormat, canonicalHwpx: Uint8Array): Promise<CanonicalExportPayload> {
  return exporters[format](canonicalHwpx);
}

async function exportHwp(canonicalHwpx: Uint8Array): Promise<CanonicalExportPayload> {
  const document = await openRhwp(canonicalHwpx);
  try {
    const exported = document.exportHwpWithReport();
    try {
      const bytes = exported.takeBytes();
      const contentLoss = safeJson(exported.contentLoss());
      const lossCount = getLossCount(contentLoss);
      return {
        bytes,
        warnings: lossCount > 0
          ? [`OMDX → HWP 저장 과정에서 ${lossCount}개의 호환성 손실이 보고되었습니다.`]
          : [],
      };
    } finally {
      exported.free();
    }
  } finally {
    document.free();
  }
}

function getLossCount(report: unknown): number {
  if (!report || typeof report !== 'object') return 0;
  const count = (report as { count?: unknown }).count;
  return typeof count === 'number' ? count : 0;
}

function safeJson(value: string): unknown {
  try {
    return JSON.parse(value);
  } catch {
    return value;
  }
}
