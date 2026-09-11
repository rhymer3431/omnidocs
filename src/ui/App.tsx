import { useEffect, useMemo, useRef, useState } from 'react';
import { formatFromFileName, type LoadResult, type OmniFormat, type OmniFile } from '../domain/document';
import { RhwpStudioEngine } from '../editor/RhwpStudioEngine';

type View = 'home' | 'editor';
type ToolbarTab = 'basic' | 'advanced';
type PendingAction = { kind: 'new' } | { kind: 'file'; file: File };
type RecentDocument = { name: string; format: OmniFormat; file: File };

function download(bytes: Uint8Array, mimeType: string, fileName: string) {
  const stableBytes = Uint8Array.from(bytes);
  const blob = new Blob([stableBytes.buffer], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

function ToolButton({
  label,
  children,
  disabled,
  onClick,
}: {
  label: string;
  children: React.ReactNode;
  disabled?: boolean;
  onClick?: () => void;
}) {
  return (
    <button className="tool-button" title={label} aria-label={label} disabled={disabled} onClick={onClick}>
      {children}
    </button>
  );
}

function Divider() {
  return <span className="toolbar-divider" aria-hidden="true" />;
}

function FileIcon() {
  return <span className="file-icon" aria-hidden="true">▤</span>;
}

function formatLabel(format: OmniFormat) {
  return format.toUpperCase();
}

export function App() {
  const mountRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const engine = useMemo(() => new RhwpStudioEngine(), []);

  const [view, setView] = useState<View>('home');
  const [ready, setReady] = useState(false);
  const [busy, setBusy] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [fileName, setFileName] = useState('Untitled.omdx');
  const [format, setFormat] = useState<OmniFormat>('omdx');
  const [pageCount, setPageCount] = useState(0);
  const [message, setMessage] = useState('Ready');
  const [warnings, setWarnings] = useState<string[]>([]);
  const [pendingAction, setPendingAction] = useState<PendingAction | null>(null);
  const [recentDocuments, setRecentDocuments] = useState<RecentDocument[]>([]);
  const [openModal, setOpenModal] = useState(false);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [exportOpen, setExportOpen] = useState(false);
  const [insertOpen, setInsertOpen] = useState(false);
  const [toolbarTab, setToolbarTab] = useState<ToolbarTab>('basic');

  useEffect(() => engine.onDirtyChange(setDirty), [engine]);

  useEffect(() => {
    if (view !== 'editor' || !mountRef.current) return;
    let cancelled = false;
    const action = pendingAction;

    const start = async () => {
      try {
        await engine.mount(mountRef.current!);
        if (cancelled) return;
        setReady(true);

        if (!action) return;
        setBusy(true);
        let result: LoadResult;
        if (action.kind === 'new') {
          result = await engine.newDocument('Untitled.omdx');
          if (cancelled) return;
          setFileName('Untitled.omdx');
          setFormat('omdx');
        } else {
          const nextFormat = formatFromFileName(action.file.name);
          const omniFile: OmniFile = {
            name: action.file.name,
            format: nextFormat,
            bytes: new Uint8Array(await action.file.arrayBuffer()),
          };
          result = await engine.load(omniFile);
          if (cancelled) return;
          setFileName(action.file.name);
          setFormat(nextFormat);
          setRecentDocuments((current) => [
            { name: action.file.name, format: nextFormat, file: action.file },
            ...current.filter((item) => item.name !== action.file.name),
          ].slice(0, 8));
        }

        setPageCount(result.pageCount);
        setWarnings(result.warnings);
        setMessage(`${result.pageCount} page${result.pageCount === 1 ? '' : 's'} · OMDX`);
        setPendingAction(null);
      } catch (error) {
        if (!cancelled) setMessage(error instanceof Error ? error.message : String(error));
      } finally {
        if (!cancelled) setBusy(false);
      }
    };

    void start();
    return () => {
      cancelled = true;
      setReady(false);
      engine.destroy();
    };
  }, [engine, view]);

  function beginNewDocument() {
    setPendingAction({ kind: 'new' });
    setWarnings([]);
    setToolbarTab('basic');
    setView('editor');
  }

  function beginOpenFile(file: File) {
    setSelectedFile(null);
    setOpenModal(false);
    setPendingAction({ kind: 'file', file });
    setWarnings([]);
    setToolbarTab('basic');
    setView('editor');
  }

  async function openFromEditor(file: File) {
    setBusy(true);
    setWarnings([]);
    try {
      const nextFormat = formatFromFileName(file.name);
      const omniFile: OmniFile = {
        name: file.name,
        format: nextFormat,
        bytes: new Uint8Array(await file.arrayBuffer()),
      };
      const result = await engine.load(omniFile);
      setFileName(file.name);
      setFormat(nextFormat);
      setPageCount(result.pageCount);
      setWarnings(result.warnings);
      setToolbarTab('basic');
      setMessage(`${result.pageCount} page${result.pageCount === 1 ? '' : 's'} · OMDX`);
      setRecentDocuments((current) => [
        { name: file.name, format: nextFormat, file },
        ...current.filter((item) => item.name !== file.name),
      ].slice(0, 8));
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
      if (inputRef.current) inputRef.current.value = '';
    }
  }

  async function saveAs(target: OmniFormat) {
    if (!ready || pageCount === 0) return;
    setBusy(true);
    setExportOpen(false);
    try {
      const result = await engine.export(target, fileName);
      download(result.bytes, result.mimeType, result.fileName);
      setFormat(target);
      setFileName(result.fileName);
      setWarnings(result.warnings);
      setMessage(`${result.fileName} saved`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  async function runCommand(commandId: string, params?: Record<string, unknown>) {
    if (!ready || pageCount === 0) return;
    try {
      await engine.execute(commandId, params);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    }
  }

  function triggerEditorFilePicker() {
    if (!inputRef.current) return;
    inputRef.current.dataset.mode = 'editor';
    inputRef.current.click();
  }

  const editorDisabled = !ready || busy || pageCount === 0;

  return (
    <>
      <input
        ref={inputRef}
        className="file-input"
        type="file"
        accept=".hwp,.hwpx,.docx,.omdx"
        onChange={(event) => {
          const file = event.currentTarget.files?.[0];
          if (!file) return;
          const mode = event.currentTarget.dataset.mode;
          delete event.currentTarget.dataset.mode;
          if (mode === 'editor') {
            void openFromEditor(file);
          } else {
            setSelectedFile(file);
            event.currentTarget.value = '';
          }
        }}
      />

      {view === 'home' ? (
        <main className="home-shell">
          <header className="home-header">
            <button className="brand-button" type="button">
              <span className="brand-mark">O</span>
              <span>OmniDocs</span>
            </button>
            <div className="home-actions">
              <button className="primary-button" onClick={beginNewDocument}>New document</button>
              <button className="ghost-button" onClick={() => setOpenModal(true)}>Open file</button>
              <button className="icon-button" aria-label="Settings" title="Settings">⚙</button>
            </div>
          </header>

          <section className="documents-view">
            <p className="eyebrow">Workspace</p>
            <h1>Documents</h1>
            <p className="documents-subtitle">Your recent work, kept in one quiet place.</p>

            {recentDocuments.length ? (
              <div className="document-list">
                {recentDocuments.map((document) => (
                  <button
                    className="document-row"
                    key={`${document.name}-${document.file.lastModified}`}
                    onClick={() => beginOpenFile(document.file)}
                  >
                    <FileIcon />
                    <span className="document-copy">
                      <strong>{document.name}</strong>
                      <small>{formatLabel(document.format)} · Opened this session</small>
                    </span>
                    <span className="row-more" aria-hidden="true">···</span>
                  </button>
                ))}
              </div>
            ) : (
              <div className="empty-state">
                <FileIcon />
                <strong>No documents yet</strong>
                <p>Open an existing file or create a clean OMDX document.</p>
                <div>
                  <button className="primary-button" onClick={beginNewDocument}>Create document</button>
                  <button className="ghost-button" onClick={() => setOpenModal(true)}>Open existing file</button>
                </div>
              </div>
            )}
          </section>
        </main>
      ) : (
        <main className="editor-shell">
          <header className="editor-topbar">
            <div className="editor-title-group">
              <button className="icon-button back-button" aria-label="Back to documents" onClick={() => setView('home')}>←</button>
              <button className="icon-button compact-open" aria-label="Open document" title="Open document" onClick={triggerEditorFilePicker}>▤</button>
              <input
                className="document-title"
                value={fileName.replace(/\.[^.]+$/, '')}
                aria-label="Document title"
                onChange={(event) => {
                  const extension = fileName.includes('.') ? fileName.slice(fileName.lastIndexOf('.')) : `.${format}`;
                  setFileName(`${event.target.value}${extension}`);
                }}
              />
            </div>

            <div className="editor-actions">
              <span className={`save-state ${dirty ? 'unsaved' : ''}`}>
                {busy ? 'Saving…' : dirty ? 'Unsaved' : 'Saved'}
              </span>
              <button className="ghost-button share-button" disabled>Share</button>
              <div className="menu-anchor">
                <button className="export-button" disabled={editorDisabled} onClick={() => setExportOpen((open) => !open)}>Export</button>
                {exportOpen && (
                  <div className="popup-menu export-menu">
                    {(['omdx', 'docx', 'hwp', 'hwpx'] as OmniFormat[]).map((target) => (
                      <button key={target} onClick={() => void saveAs(target)}>
                        <span>Export as {target.toUpperCase()}</span>
                        {target === format && <span className="menu-check">✓</span>}
                      </button>
                    ))}
                  </div>
                )}
              </div>
              <button className="icon-button" aria-label="More options">···</button>
            </div>
          </header>

          <section className="editor-ribbon" aria-label="Document tools">
            <nav className="toolbar-tabs" aria-label="Tool categories">
              <button className={toolbarTab === 'basic' ? 'active' : ''} onClick={() => setToolbarTab('basic')}>기본</button>
              <button className={toolbarTab === 'advanced' ? 'active' : ''} onClick={() => setToolbarTab('advanced')}>고급</button>
              <span className="tab-help">고급 도구도 DOCX · HWP · HWPX에서 사용할 수 있으며, 변환 시 일부 고유 표현은 단순화될 수 있습니다.</span>
            </nav>

            {toolbarTab === 'basic' && (
              <nav className="editor-toolbar" aria-label="Basic editor toolbar">
                <ToolButton label="Undo" disabled={editorDisabled} onClick={() => void runCommand('edit:undo')}>↶</ToolButton>
                <ToolButton label="Redo" disabled={editorDisabled} onClick={() => void runCommand('edit:redo')}>↷</ToolButton>
                <Divider />
                <button className="toolbar-select" disabled={editorDisabled} onClick={() => void runCommand('format:style-dialog')}>Normal <span>⌄</span></button>
                <button className="toolbar-select font-select" disabled={editorDisabled} onClick={() => void runCommand('format:char-shape')}>Font <span>⌄</span></button>
                <Divider />
                <ToolButton label="Bold" disabled={editorDisabled} onClick={() => void runCommand('format:bold')}><b>B</b></ToolButton>
                <ToolButton label="Italic" disabled={editorDisabled} onClick={() => void runCommand('format:italic')}><i>I</i></ToolButton>
                <ToolButton label="Underline" disabled={editorDisabled} onClick={() => void runCommand('format:underline')}><u>U</u></ToolButton>
                <ToolButton label="Strikethrough" disabled={editorDisabled} onClick={() => void runCommand('format:strikethrough')}><s>S</s></ToolButton>
                <Divider />
                <ToolButton label="Align left" disabled={editorDisabled} onClick={() => void runCommand('format:align-left')}><span className="align-icon left">≡</span></ToolButton>
                <ToolButton label="Align center" disabled={editorDisabled} onClick={() => void runCommand('format:align-center')}><span className="align-icon">≡</span></ToolButton>
                <ToolButton label="Align right" disabled={editorDisabled} onClick={() => void runCommand('format:align-right')}><span className="align-icon right">≡</span></ToolButton>
                <ToolButton label="Justify" disabled={editorDisabled} onClick={() => void runCommand('format:align-justify')}>☰</ToolButton>
                <Divider />
                <ToolButton label="Bullet list" disabled={editorDisabled} onClick={() => void runCommand('format:toggle-bullet')}>•☰</ToolButton>
                <ToolButton label="Numbered list" disabled={editorDisabled} onClick={() => void runCommand('format:toggle-numbering')}>1☰</ToolButton>
                <Divider />
                <div className="menu-anchor">
                  <button className="insert-button" disabled={editorDisabled} onClick={() => setInsertOpen((open) => !open)}>＋ Insert</button>
                  {insertOpen && (
                    <div className="popup-menu insert-menu">
                      <button onClick={() => { setInsertOpen(false); void runCommand('insert:image'); }}><span>Image</span><small>DOCX / HWP</small></button>
                      <button onClick={() => { setInsertOpen(false); void runCommand('table:create'); }}><span>Table</span><small>DOCX / HWP</small></button>
                      <button onClick={() => { setInsertOpen(false); void runCommand('insert:textbox'); }}><span>Text box</span><small>Basic floating text</small></button>
                    </div>
                  )}
                </div>
              </nav>
            )}

            {toolbarTab === 'advanced' && (
              <nav className="editor-toolbar advanced-toolbar" aria-label="Advanced document tools">
                <div className="tool-group">
                  <span className="tool-group-label">서식</span>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:char-shape')}>글자 모양</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:para-shape')}>문단 모양</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:style-dialog')}>스타일</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:superscript')}>위 첨자</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:subscript')}>아래 첨자</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:emboss')}>양각</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:engrave')}>음각</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:outline')}>외곽선</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:align-split')}>나눔 정렬</button>
                </div>
                <Divider />
                <div className="tool-group">
                  <span className="tool-group-label">문단</span>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:line-spacing-decrease')}>줄 간격 −</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:line-spacing-increase')}>줄 간격 +</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:level-increase')}>수준 증가</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:level-decrease')}>수준 감소</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:para-num-shape')}>문단 번호</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:bullet-shape')}>글머리표</button>
                </div>
                <Divider />
                <div className="tool-group">
                  <span className="tool-group-label">쪽</span>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('page:setup')}>편집 용지</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('page:break')}>쪽 나누기</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('page:page-border')}>테두리/배경</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('page:section-settings')}>구역</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('page:col-settings')}>단</button>
                </div>
                <Divider />
                <div className="tool-group">
                  <span className="tool-group-label">삽입</span>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:footnote')}>각주</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:endnote')}>미주</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:endnote-shape')}>미주 모양</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:equation')}>수식</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:symbols')}>문자표</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:bookmark')}>책갈피</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:shape')}>도형</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:caption-toggle')}>캡션</button>
                </div>
                <Divider />
                <div className="tool-group">
                  <span className="tool-group-label">머리말</span>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('page:header-create')}>머리말</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('page:footer-create')}>꼬리말</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('page:new-page-num')}>새 쪽 번호</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('page:insert-field-pagenum')}>쪽 번호 필드</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('page:insert-field-totalpage')}>전체 쪽수</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('page:insert-field-filename')}>파일명</button>
                </div>
                <Divider />
                <div className="tool-group">
                  <span className="tool-group-label">표</span>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('table:cell-props')}>셀 속성</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('table:cell-merge')}>셀 합치기</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('table:cell-split')}>셀 나누기</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('table:insert-row-below')}>행 추가</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('table:insert-col-right')}>열 추가</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('table:delete-row')}>행 삭제</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('table:delete-col')}>열 삭제</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('table:border-each')}>셀 테두리</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('table:cell-width-equal')}>너비 같게</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('table:cell-height-equal')}>높이 같게</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('table:formula')}>계산식</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('table:block-formula')}>블록 계산식</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('table:thousand-sep')}>천 단위</button>
                </div>
                <Divider />
                <div className="tool-group">
                  <span className="tool-group-label">개체</span>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('format:object-properties')}>개체 속성</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:picture-props')}>그림 속성</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:arrange-front')}>맨 앞으로</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:arrange-back')}>맨 뒤로</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:rotate-cw')}>90° 회전</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:flip-horz')}>좌우 대칭</button>
                  <button className="text-tool" disabled={editorDisabled} onClick={() => void runCommand('insert:flip-vert')}>상하 대칭</button>
                </div>
              </nav>
            )}
          </section>

          <section className="document-canvas">
            {!ready && <div className="editor-loading">Preparing document editor…</div>}
            <div className="editor-frame" ref={mountRef} />
            {warnings.length > 0 && (
              <aside className="compatibility-note">
                <strong>Compatibility</strong>
                <span>{warnings.slice(0, 2).join(' · ')}</span>
              </aside>
            )}
          </section>

          <footer className="editor-footer">
            <span>{message}</span>
            <span>{pageCount ? `${pageCount} page${pageCount === 1 ? '' : 's'}` : formatLabel(format)}</span>
          </footer>
        </main>
      )}

      {openModal && (
        <div className="modal-backdrop" role="presentation" onMouseDown={() => setOpenModal(false)}>
          <section className="open-modal" role="dialog" aria-modal="true" aria-labelledby="open-document-title" onMouseDown={(event) => event.stopPropagation()}>
            <div className="modal-heading">
              <div>
                <h2 id="open-document-title">Open document</h2>
                <p>Choose a compatible file from your computer.</p>
              </div>
              <button className="icon-button" aria-label="Close" onClick={() => setOpenModal(false)}>×</button>
            </div>

            <button
              className={`file-drop-choice ${selectedFile ? 'selected' : ''}`}
              onClick={() => inputRef.current?.click()}
            >
              <FileIcon />
              <span>
                <strong>{selectedFile?.name ?? 'Choose a document'}</strong>
                <small>{selectedFile ? `${Math.max(1, Math.round(selectedFile.size / 1024))} KB` : 'Browse local files'}</small>
              </span>
              <span>{selectedFile ? '✓' : 'Browse'}</span>
            </button>

            <p className="supported-formats">Supports DOCX, HWP, HWPX and OMDX</p>
            <div className="modal-actions">
              <button className="ghost-button" onClick={() => { setSelectedFile(null); setOpenModal(false); }}>Cancel</button>
              <button className="primary-button" disabled={!selectedFile} onClick={() => selectedFile && beginOpenFile(selectedFile)}>Open</button>
            </div>
          </section>
        </div>
      )}
    </>
  );
}
