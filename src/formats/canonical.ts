import type { SourceFormat } from '../domain/document';

/**
 * OMDX keeps HWPX as the concrete layout/editing authority, but format
 * adapters communicate through this envelope instead of leaking parser-specific
 * return types into the editor. This is intentionally a thin IR: paragraph and
 * table layout remain owned by rHWP rather than being reimplemented in JS.
 */
export interface CanonicalDocumentPayload {
  canonicalHwpx: Uint8Array;
  pageCount: number;
  warnings: string[];
  contentLoss?: unknown;
  layout?: unknown;
  features?: DocumentFeatureInventory;
}

export interface SourceFormatAdapter {
  readonly format: SourceFormat;
  readonly id: string;
  import(bytes: Uint8Array): Promise<CanonicalDocumentPayload>;
}

export type FeatureHandling = 'native' | 'normalized' | 'approximated' | 'source-only';

export interface DocumentFeatureEntry {
  id: string;
  label: string;
  count: number;
  handling: FeatureHandling;
  note?: string;
}

export interface DocumentFeatureInventory {
  sourceFormat: SourceFormat;
  adapter: string;
  features: DocumentFeatureEntry[];
}

export function presentFeatures(inventory?: DocumentFeatureInventory): DocumentFeatureEntry[] {
  return inventory?.features.filter((feature) => feature.count > 0) ?? [];
}

export function sourceOnlyFeatureWarnings(inventory?: DocumentFeatureInventory): string[] {
  if (!inventory) return [];
  return inventory.features
    .filter((feature) => feature.count > 0 && feature.handling === 'source-only')
    .map((feature) => `${feature.label} ${feature.count}개는 원본 파일에 보존되지만 현재 canonical 편집 모델에서는 단순화될 수 있습니다.`);
}
