import { useEffect, useMemo, useRef, useState } from 'react';
import { formatFromFileName, type OmniFormat, type OmniFile } from '../domain/document';
import { RhwpStudioEngine } from '../editor/RhwpStudioEngine';

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

export function App() {
  const mountRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const engine = useMemo(() => new RhwpStudioEngine(), []);
  const [ready, setReady] = useState(false);
  const [busy, setBusy] = useState(false);
  const [fileName, setFileName] = useState('새 문서.omdx');
  const [format, setFormat] = useState<OmniFormat>('omdx');
  const [pageCount, setPageCount] = useState(0);
  const [message, setMessage] = useState('편집기를 준비하고 있습니다.');
  const [warnings, setWarnings] = useState<string[]>([]);

  useEffect(() => {
    let cancelled = false;
    if (!mountRef.current) return;
    engine
      .mount(mountRef.current)
      .then(() => {
        if (!cancelled) {
          setReady(true);
          setMessage('HWP, HWPX, DOCX 또는 OMDX 파일을 열면 내부 OMDX 문서로 정규화됩니다.');
        }
      })
      .catch((error: unknown) => {
        if (!cancelled) setMessage(error instanceof Error ? error.message : String(error));
      });
    return () => {
      cancelled = true;
      engine.destroy();
    };
  }, [engine]);

  async function openFile(file: File) {
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
      setMessage(`${file.name} · ${result.pageCount}페이지 · 내부 OMDX`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
      if (inputRef.current) inputRef.current.value = '';
    }
  }

  async function saveAs(target: OmniFormat) {
    setBusy(true);
    try {
      const result = await engine.export(target, fileName);
      download(result.bytes, result.mimeType, result.fileName);
      setFormat(target);
      setFileName(result.fileName);
      setWarnings(result.warnings);
      setMessage(`${result.fileName} 저장 완료`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand">
          <div className="brand-mark">O</div>
          <div>
            <strong>OmniDocs</strong>
            <span>OMDX · HWP · HWPX · DOCX</span>
          </div>
        </div>
        <div className="document-info">
          <strong>{fileName}</strong>
          <span>{pageCount ? `${pageCount} pages` : format.toUpperCase()}</span>
        </div>
        <div className="actions">
          <input
            ref={inputRef}
            className="file-input"
            type="file"
            accept=".hwp,.hwpx,.docx,.omdx"
            onChange={(event) => {
              const file = event.currentTarget.files?.[0];
              if (file) void openFile(file);
            }}
          />
          <button disabled={!ready || busy} onClick={() => inputRef.current?.click()}>
            열기
          </button>
          <div className="save-group">
            <button disabled={!ready || busy} onClick={() => void saveAs(format)}>
              저장
            </button>
            <select
              aria-label="저장 형식"
              value={format}
              disabled={!ready || busy}
              onChange={(event) => void saveAs(event.target.value as OmniFormat)}
            >
              <option value="hwp">HWP로 저장</option>
              <option value="hwpx">HWPX로 저장</option>
              <option value="docx">DOCX로 저장</option>
              <option value="omdx">OMDX로 저장</option>
            </select>
          </div>
        </div>
      </header>

      <section className="status-strip" aria-live="polite">
        <span className={busy ? 'status-dot busy' : 'status-dot'} />
        <span>{busy ? '처리 중…' : message}</span>
      </section>

      {warnings.length > 0 && (
        <aside className="warning-panel">
          <strong>호환성 알림</strong>
          <span>{warnings.slice(0, 3).join(' · ')}</span>
        </aside>
      )}

      <section className="editor-frame" ref={mountRef} />
    </main>
  );
}
