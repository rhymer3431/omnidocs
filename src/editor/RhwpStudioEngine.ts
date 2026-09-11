import { createStudio } from '@rhwp/editor';
import type { ExportResult, LoadResult, OmniFile, OmniFormat } from '../domain/document';
import { replaceExtension } from '../domain/document';
import { exportCanonical } from '../formats/canonicalExport';
import { buildOmdx, canUseOriginalSource, OMDX_MIME, type OmdxDocument } from '../formats/omdx';
import { importAsOmdx } from '../formats/omdxGateway';
import { createBlankHwpx } from '../formats/rhwpRuntime';
import type { EditorCommand, OmniEditorEngine } from './contracts';
import { resolveStudioUrl, unregisterLegacyRhwpServiceWorker } from './studioUrl';

const MIME: Record<OmniFormat, string> = {
  hwp: 'application/x-hwp',
  hwpx: 'application/zip',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  omdx: OMDX_MIME,
};

export class RhwpStudioEngine implements OmniEditorEngine {
  private studio: Awaited<ReturnType<typeof createStudio>> | null = null;
  private mountPromise: Promise<void> | null = null;
  private mountGeneration = 0;
  private dirty = false;
  private dirtyListeners = new Set<(dirty: boolean) => void>();
  private unsubscribe: (() => void) | null = null;
  private warnings: string[] = [];
  private workingDocument: OmdxDocument | null = null;
  private workingPackageBytes: Uint8Array | null = null;

  async mount(container: HTMLElement): Promise<void> {
    if (this.studio) return;
    if (this.mountPromise) return this.mountPromise;

    const generation = ++this.mountGeneration;
    const promise = (async () => {
      await unregisterLegacyRhwpServiceWorker();

      // This container is exclusively owned by the editor. Clearing it also
      // removes an orphaned iframe left by a previous interrupted dev mount.
      container.replaceChildren();

      const studio = await createStudio(container, {
        studioUrl: resolveStudioUrl(import.meta.env.VITE_RHWP_STUDIO_URL),
        plugins: ['hwpctrl'],
        // OmniDocs owns the visible editor chrome. Keep Studio as the document
        // canvas/command engine so the UI is not duplicated inside the iframe.
        chrome: { menu: false, toolbar: false, statusbar: false },
        width: '100%',
        height: '100%',
      });

      // React StrictMode deliberately performs mount -> cleanup -> mount in
      // development. createStudio() is asynchronous, so the first mount may
      // finish after cleanup. Never publish that stale Studio instance.
      if (generation !== this.mountGeneration) {
        studio.destroy();
        return;
      }

      this.studio = studio;
      applyOmniDocsStudioTheme(studio.element);
      this.unsubscribe = studio.onDocumentChanged(() => {
        this.setDirty(true);
      });
    })();

    this.mountPromise = promise;
    try {
      await promise;
    } finally {
      if (this.mountPromise === promise) this.mountPromise = null;
    }
  }

  async newDocument(fileName = 'Untitled.omdx'): Promise<LoadResult> {
    const studio = this.requireStudio();
    const canonicalHwpx = await createBlankHwpx();
    const result = await studio.loadFile(canonicalHwpx, replaceExtension(fileName, 'hwpx'), {
      skipUnsavedGuard: true,
      suppressDialogs: true,
    });
    const built = await buildOmdx({
      canonicalHwpx,
      pageCount: result.pageCount,
    });
    this.workingDocument = built.document;
    this.workingPackageBytes = built.bytes;
    this.warnings = [];
    this.setDirty(false);
    return { pageCount: result.pageCount, warnings: [], internalFormat: 'omdx' };
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
    this.setDirty(false);
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
      this.setDirty(false);
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
      this.setDirty(false);
      return {
        bytes: built.bytes,
        mimeType: MIME.omdx,
        fileName: replaceExtension(sourceName, 'omdx'),
        warnings,
      };
    }

    const exported = await exportCanonical(format, currentHwpx);
    warnings.push(...exported.warnings);
    this.setDirty(false);
    return {
      bytes: exported.bytes,
      mimeType: MIME[format],
      fileName: replaceExtension(sourceName, format),
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

  onDirtyChange(listener: (dirty: boolean) => void): () => void {
    this.dirtyListeners.add(listener);
    listener(this.dirty);
    return () => this.dirtyListeners.delete(listener);
  }

  destroy(): void {
    // Invalidate an in-flight asynchronous mount. It cannot be cancelled, so
    // mount() will destroy the resulting stale iframe when it eventually
    // resolves. Clearing mountPromise lets a StrictMode remount start at once.
    this.mountGeneration += 1;
    this.mountPromise = null;
    this.unsubscribe?.();
    this.unsubscribe = null;
    this.studio?.destroy();
    this.studio = null;
    this.workingDocument = null;
    this.workingPackageBytes = null;
    this.setDirty(false);
  }

  private setDirty(next: boolean): void {
    if (this.dirty === next) return;
    this.dirty = next;
    for (const listener of this.dirtyListeners) listener(next);
  }

  private requireStudio(): Awaited<ReturnType<typeof createStudio>> {
    if (!this.studio) throw new Error('편집기가 아직 초기화되지 않았습니다.');
    return this.studio;
  }
}

function applyOmniDocsStudioTheme(iframe: HTMLIFrameElement): void {
  let document: Document | null = null;
  try {
    document = iframe.contentDocument;
  } catch {
    // A user supplied cross-origin Studio override cannot be themed from the
    // host page. The default self-hosted Studio is same-origin.
    return;
  }
  if (!document) return;
  document.documentElement.dataset.themeSkin = 'flat';
  const existing = document.getElementById('omnidocs-studio-theme');
  if (existing) return;

  const style = document.createElement('style');
  style.id = 'omnidocs-studio-theme';
  style.textContent = `
    :root {
      --ui-bg: #f7f7f5;
      --ui-bg-light: #fbfbfa;
      --ui-surface: #ffffff;
      --ui-surface-raised: #ffffff;
      --ui-surface-muted: #f3f3f1;
      --ui-border: #e4e4df;
      --ui-border-light: #e8e8e4;
      --ui-border-subtle: #eeeeeb;
      --ui-border-strong: #d4d4cf;
      --ui-text: #292928;
      --ui-text-secondary: #575754;
      --ui-text-muted: #858581;
      --ui-text-hint: #9a9a96;
      --ui-text-disabled: #b6b6b1;
      --ui-hover: #f3f3f1;
      --ui-hover-strong: #efefec;
      --ui-active: #e8e8f7;
      --ui-selected: #eeeeff;
      --ui-menu-open: #414094;
      --ui-menu-open-border: #414094;
      --accent-primary: #414094;
      --accent-strong: #35347e;
      --accent-light: #6563b5;
      --accent-hover: #c8c7e9;
      --accent-subtle: #7775bd;
      --focus-ring: #6563b5;
      --doc-workspace: #f7f7f5;
      --doc-shadow: #2929281c;
      --ruler-bg: #f7f7f5;
      --ruler-body: #ffffff;
      --ruler-tick: #9b9b96;
      --ruler-text: #767672;
      --ruler-marker: #6563b5;
      --shadow-dropdown: 0 10px 28px rgba(30, 30, 20, .10);
      --shadow-dialog: 0 18px 50px rgba(31, 31, 28, .14);
      --shadow-light: 0 4px 14px rgba(31, 31, 28, .08);
      --radius-sm: 5px;
      --radius-md: 7px;
      --radius-lg: 11px;
      --font-family-ui: Inter, Pretendard, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }
    body { background: #f7f7f5; }
    #scroll-container { background: #f7f7f5 !important; }
    .modal-overlay { background: rgba(41, 41, 40, .14) !important; backdrop-filter: blur(2px); }
    .dialog-wrap,
    .compare-dialog,
    .compare-inspector-window,
    .cp-panel,
    .shape-picker,
    .font-picker-menu,
    .context-menu,
    .tb-split-menu,
    .form-combo-dropdown {
      border: 1px solid #e7e7e3 !important;
      border-radius: 11px !important;
      background: #fff !important;
      box-shadow: 0 18px 50px rgba(31, 31, 28, .14) !important;
      overflow: hidden;
    }
    .dialog-title,
    .compare-dialog-title,
    .compare-inspector-head,
    .shape-picker-title {
      min-height: 45px;
      display: flex;
      align-items: center;
      border-bottom: 1px solid #eeeeeb !important;
      background: #fff !important;
      color: #343432 !important;
      font-size: 14px !important;
      font-weight: 650 !important;
      letter-spacing: -.015em;
      padding: 0 16px !important;
    }
    .dialog-body { padding: 16px !important; }
    .dialog-footer {
      gap: 6px !important;
      border-top: 1px solid #eeeeeb !important;
      background: #fbfbfa !important;
      padding: 10px 14px !important;
    }
    .dialog-btn {
      min-height: 32px !important;
      border: 1px solid #e4e4df !important;
      border-radius: 7px !important;
      background: #fff !important;
      color: #4f4f4b !important;
      padding: 0 12px !important;
      font-weight: 550 !important;
      box-shadow: none !important;
    }
    .dialog-btn:hover:not(:disabled) { background: #f3f3f1 !important; }
    .dialog-btn-primary {
      border-color: #414094 !important;
      background: #414094 !important;
      color: #fff !important;
    }
    .dialog-btn-primary:hover:not(:disabled) { background: #35347e !important; }
    .dialog-close {
      border-radius: 6px !important;
      color: #8b8b86 !important;
    }
    .dialog-close:hover { background: #f3f3f1 !important; color: #4f4f4b !important; }
    .dialog-input,
    .dialog-text-input,
    .dialog-select,
    .find-dialog-input,
    .compare-strategy-select,
    .history-label-input,
    .dialog-wrap input, .dialog-wrap select, .dialog-wrap textarea,
    .compare-dialog input, .compare-dialog select, .compare-dialog textarea,
    .compare-inspector-window input, .compare-inspector-window select, .compare-inspector-window textarea,
    .cp-panel input {
      border: 1px solid #dfdfda !important;
      border-radius: 7px !important;
      background: #fff !important;
      color: #343432 !important;
      box-shadow: none !important;
      outline: none;
    }
    .dialog-input:focus,
    .dialog-text-input:focus,
    .dialog-select:focus,
    .dialog-wrap input:focus, .dialog-wrap select:focus, .dialog-wrap textarea:focus,
    .compare-dialog input:focus, .compare-dialog select:focus, .compare-dialog textarea:focus,
    .compare-inspector-window input:focus, .compare-inspector-window select:focus, .compare-inspector-window textarea:focus,
    .cp-panel input:focus {
      border-color: #7775bd !important;
      box-shadow: 0 0 0 2px #eeeeff !important;
    }
    .dialog-section {
      border-color: #eeeeeb !important;
      border-radius: 8px !important;
      background: #fff !important;
    }
    .dialog-section-title,
    .dialog-label { color: #696965 !important; font-weight: 600 !important; }
    .dialog-tabs { border-bottom: 1px solid #eeeeeb !important; gap: 3px; }
    .dialog-tab {
      border: 0 !important;
      border-radius: 6px 6px 0 0 !important;
      background: transparent !important;
      color: #858581 !important;
      padding: 7px 10px !important;
    }
    .dialog-tab.active,
    .dialog-tab[aria-selected="true"] { color: #414094 !important; background: #eeeeff !important; }
    .shape-picker-grid { gap: 4px !important; padding: 8px !important; }
    .shape-picker-btn { border-radius: 7px !important; padding: 7px 5px !important; }
    .shape-picker-btn:hover { border-color: transparent !important; background: #f3f3f1 !important; }
    .context-menu { padding: 5px !important; }
    .context-menu > *, .font-picker-option, .tb-split-item, .form-combo-dropdown > * { border-radius: 6px !important; }
    .font-picker-option:hover, .tb-split-item:hover { background: #f3f3f1 !important; }
    .cp-input-wrap { background: #fff !important; border-color: #eeeeeb !important; }
    .cp-item { border-radius: 7px !important; }
    .cp-item:hover, .cp-item--selected { background: #eeeeff !important; }
    .cp-item-shortcut { border-radius: 5px !important; background: #f5f5f3 !important; }
  `;
  document.head.appendChild(style);
}
