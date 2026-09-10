import { createStudio } from '@rhwp/editor';
import type { ExportResult, LoadResult, OmniFile, OmniFormat } from '../domain/document';
import { replaceExtension } from '../domain/document';
import { exportDocx } from '../formats/docxAdapter';
import { buildOmdx, canUseOriginalSource, OMDX_MIME, type OmdxDocument } from '../formats/omdx';
import { importAsOmdx } from '../formats/omdxGateway';
import { openRhwp } from '../formats/rhwpRuntime';
import type { EditorCommand, OmniEditorEngine } from './contracts';
import { resolveStudioUrl } from './studioUrl';

const MIME: Record<OmniFormat, string> = {
  hwp: 'application/x-hwp',
  hwpx: 'application/zip',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  omdx: OMDX_MIME,
};

export class RhwpStudioEngine implements OmniEditorEngine {
  private studio: Awaited<ReturnType<typeof createStudio>> | null = null;
  private dirty = false;
  private unsubscribe: (() => void) | null = null;
  private warnings: string[] = [];
  private workingDocument: OmdxDocument | null = null;
  private workingPackageBytes: Uint8Array | null = null;

  async mount(container: HTMLElement): Promise<void> {
    if (this.studio) return;
    this.studio = await createStudio(container, {
      studioUrl: resolveStudioUrl(import.meta.env.VITE_RHWP_STUDIO_URL),
      plugins: ['hwpctrl'],
      chrome: { statusbar: true },
      width: '100%',
      height: '100%',
    });
    this.unsubscribe = this.studio.onDocumentChanged(() => {
      this.dirty = true;
    });
  }

  async load(file: OmniFile): Promise<LoadResult> {
    const studio = this.requireStudio();
    const imported = await importAsOmdx(file);
    this.workingDocument = imported.document;
    this.workingPackageBytes = imported.packageBytes;
    this.warnings = imported.warnings;

    const result = await studio.loadFile(imported.document.canonicalHwpx, replaceExtension(file.name, 'hwpx'), {
      skipUnsavedGuard: true,
      suppressDialogs: true,
    });
    this.dirty = false;
    return { pageCount: result.pageCount, warnings: [...this.warnings], internalFormat: 'omdx' };
  }

  async export(format: OmniFormat, sourceName: string): Promise<ExportResult> {
    const studio = this.requireStudio();
    const warnings = [...this.warnings];
    const currentHwpx = await studio.hwpctrl.exportBytes('hwpx');
    const original = this.workingDocument?.source;

    if (
      format !== 'omdx'
      && original
      && format === original.format
      && await canUseOriginalSource(this.workingDocument ?? undefined, format, currentHwpx)
    ) {
      this.dirty = false;
      return {
        bytes: Uint8Array.from(original.bytes),
        mimeType: MIME[format],
        fileName: replaceExtension(sourceName, format),
        warnings,
      };
    }

    if (format === 'omdx') {
      const built = await buildOmdx({
        canonicalHwpx: currentHwpx,
        pageCount: await studio.pageCount(),
        source: original,
        previousManifest: this.workingDocument?.manifest,
      });
      this.workingDocument = built.document;
      this.workingPackageBytes = built.bytes;
      this.dirty = false;
      return {
        bytes: built.bytes,
        mimeType: MIME.omdx,
        fileName: replaceExtension(sourceName, 'omdx'),
        warnings,
      };
    }

    if (format === 'docx') {
      const bytes = await exportDocx(currentHwpx);
      this.dirty = false;
      warnings.push('편집된 문서의 DOCX 저장은 현재 HTML 브리지를 사용하므로 Word 전용 필드/도형 일부는 단순화될 수 있습니다. 원본 DOCX를 수정하지 않은 경우에는 OMDX가 원본 바이트를 그대로 반환합니다.');
      return {
        bytes,
        mimeType: MIME.docx,
        fileName: replaceExtension(sourceName, 'docx'),
        warnings,
      };
    }

    if (format === 'hwpx') {
      this.dirty = false;
      return {
        bytes: currentHwpx,
        mimeType: MIME.hwpx,
        fileName: replaceExtension(sourceName, 'hwpx'),
        warnings,
      };
    }

    const { bytes, contentLoss } = await exportHwpWithReport(currentHwpx);
    const lossCount = getLossCount(contentLoss);
    if (lossCount > 0) warnings.push(`OMDX → HWP 저장 과정에서 ${lossCount}개의 호환성 손실이 보고되었습니다.`);
    this.dirty = false;
    return {
      bytes,
      mimeType: MIME.hwp,
      fileName: replaceExtension(sourceName, 'hwp'),
      warnings,
    };
  }

  async execute(commandId: string, params?: unknown): Promise<unknown> {
    return this.requireStudio().commands.execute(
      commandId,
      params as Record<string, unknown> | undefined,
      { allowDialog: true },
    );
  }

  async listCommands(): Promise<EditorCommand[]> {
    return this.requireStudio().commands.list();
  }

  isDirty(): boolean {
    return this.dirty;
  }

  destroy(): void {
    this.unsubscribe?.();
    this.unsubscribe = null;
    this.studio?.destroy();
    this.studio = null;
    this.workingDocument = null;
    this.workingPackageBytes = null;
  }

  private requireStudio(): Awaited<ReturnType<typeof createStudio>> {
    if (!this.studio) throw new Error('편집기가 아직 초기화되지 않았습니다.');
    return this.studio;
  }
}

async function exportHwpWithReport(hwpxBytes: Uint8Array): Promise<{ bytes: Uint8Array; contentLoss: unknown }> {
  const document = await openRhwp(hwpxBytes);
  try {
    const exported = document.exportHwpWithReport();
    try {
      return {
        bytes: exported.takeBytes(),
        contentLoss: safeJson(exported.contentLoss()),
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
