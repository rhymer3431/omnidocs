import { describe, expect, it } from 'vitest';
import { resolveStudioUrl, unregisterLegacyRhwpServiceWorker } from './studioUrl';

describe('resolveStudioUrl', () => {
  it('uses the self-hosted studio by default', () => {
    expect(resolveStudioUrl(undefined, 'http://localhost:5173/app/')).toBe(
      'http://localhost:5173/rhwp/index.html',
    );
  });

  it('keeps an explicit studio override', () => {
    expect(resolveStudioUrl('https://example.com/editor/', 'http://localhost:5173/')).toBe(
      'https://example.com/editor/',
    );
  });
});

describe('unregisterLegacyRhwpServiceWorker', () => {
  it('is a no-op in environments without serviceWorker support', async () => {
    await expect(unregisterLegacyRhwpServiceWorker()).resolves.toBeUndefined();
  });
});
