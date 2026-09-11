// @vitest-environment jsdom
import { beforeEach, describe, expect, it, vi } from 'vitest';

const { createStudio } = vi.hoisted(() => ({
  createStudio: vi.fn(),
}));

vi.mock('@rhwp/editor', () => ({ createStudio }));
vi.mock('./studioUrl', () => ({
  resolveStudioUrl: () => 'http://localhost/rhwp/index.html',
  unregisterLegacyRhwpServiceWorker: vi.fn().mockResolvedValue(undefined),
}));

import { RhwpStudioEngine } from './RhwpStudioEngine';

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => {
    resolve = done;
  });
  return { promise, resolve };
}

function studio(commandId: string) {
  const iframe = document.createElement('iframe');
  document.body.appendChild(iframe);
  return {
    element: iframe,
    destroy: vi.fn(),
    onDocumentChanged: vi.fn(() => vi.fn()),
    commands: {
      list: vi.fn().mockResolvedValue([{ id: commandId, label: commandId, enabled: true }]),
      execute: vi.fn(),
    },
  };
}

describe('RhwpStudioEngine mount lifecycle', () => {
  beforeEach(() => {
    createStudio.mockReset();
  });

  it('discards a stale asynchronous StrictMode mount instead of keeping two editors', async () => {
    const first = deferred<ReturnType<typeof studio>>();
    const second = deferred<ReturnType<typeof studio>>();
    const firstStudio = studio('first');
    const secondStudio = studio('second');
    createStudio
      .mockReturnValueOnce(first.promise)
      .mockReturnValueOnce(second.promise);

    const engine = new RhwpStudioEngine();
    const container = document.createElement('div');

    const firstMount = engine.mount(container);
    await vi.waitFor(() => expect(createStudio).toHaveBeenCalledTimes(1));

    // React StrictMode cleanup between its development-only double mounts.
    engine.destroy();
    const secondMount = engine.mount(container);
    await vi.waitFor(() => expect(createStudio).toHaveBeenCalledTimes(2));

    second.resolve(secondStudio);
    await secondMount;
    first.resolve(firstStudio);
    await firstMount;

    expect(firstStudio.destroy).toHaveBeenCalledTimes(1);
    expect(secondStudio.destroy).not.toHaveBeenCalled();
    expect(secondStudio.element.contentDocument?.getElementById('omnidocs-studio-theme')).not.toBeNull();
    await expect(engine.listCommands()).resolves.toEqual([
      { id: 'second', label: 'second', enabled: true },
    ]);
  });

  it('does not restrict export formats after advanced rHWP commands', async () => {
    const mountedStudio = studio('format:emboss');
    mountedStudio.commands.execute.mockResolvedValue({ ok: true });
    createStudio.mockResolvedValue(mountedStudio);

    const engine = new RhwpStudioEngine();
    const container = document.createElement('div');
    await engine.mount(container);
    await engine.execute('format:emboss');

    expect(mountedStudio.commands.execute).toHaveBeenCalledWith(
      'format:emboss', undefined, { allowDialog: true },
    );
  });
});
