export type SourceFormat = 'hwp' | 'hwpx' | 'docx';
export type OmniFormat = SourceFormat | 'omdx';

export interface OmniFile {
  name: string;
  format: OmniFormat;
  bytes: Uint8Array;
}

export interface LoadResult {
  pageCount: number;
  warnings: string[];
  internalFormat?: 'omdx';
}

export interface ExportResult {
  bytes: Uint8Array;
  mimeType: string;
  fileName: string;
  warnings: string[];
}

export function formatFromFileName(fileName: string): OmniFormat {
  const extension = fileName.split('.').pop()?.toLowerCase();
  if (extension === 'hwp' || extension === 'hwpx' || extension === 'docx' || extension === 'omdx') {
    return extension;
  }
  throw new Error('OmniDocs는 현재 HWP, HWPX, DOCX, OMDX 파일을 지원합니다.');
}

export function replaceExtension(fileName: string, format: OmniFormat): string {
  const base = fileName.replace(/\.[^.]+$/, '') || 'document';
  return `${base}.${format}`;
}
