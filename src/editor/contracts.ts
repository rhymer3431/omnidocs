import type { ExportResult, LoadResult, OmniFile, OmniFormat } from '../domain/document';

export interface EditorCommand {
  id: string;
  label: string;
  enabled: boolean;
  opensDialog?: boolean;
}

export interface OmniEditorEngine {
  mount(container: HTMLElement): Promise<void>;
  load(file: OmniFile): Promise<LoadResult>;
  export(format: OmniFormat, sourceName: string): Promise<ExportResult>;
  execute(commandId: string, params?: unknown): Promise<unknown>;
  listCommands(): Promise<EditorCommand[]>;
  isDirty(): boolean;
  destroy(): void;
}
