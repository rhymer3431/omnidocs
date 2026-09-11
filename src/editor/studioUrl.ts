const DEFAULT_LOCAL_STUDIO_PATH = '/rhwp/index.html';
const LOCAL_STUDIO_SCOPE_PATH = '/rhwp/';

export function resolveStudioUrl(
  configuredUrl: string | undefined,
  baseUrl: string = document.baseURI,
): string {
  if (configuredUrl?.trim()) {
    return new URL(configuredUrl, baseUrl).href;
  }

  const base = new URL(baseUrl);
  return new URL(DEFAULT_LOCAL_STUDIO_PATH, base.origin).href;
}

/**
 * Older vendored rHWP Studio builds registered the upstream PWA service
 * worker under /rhwp/. That worker can keep serving a stale Studio entry
 * bundle even after OmniDocs has been updated, so remove only registrations
 * scoped to the embedded Studio before creating its iframe.
 */
export async function unregisterLegacyRhwpServiceWorker(): Promise<void> {
  if (typeof navigator === 'undefined' || !('serviceWorker' in navigator)) return;

  try {
    const registrations = await navigator.serviceWorker.getRegistrations();
    const scopeRoot = new URL(LOCAL_STUDIO_SCOPE_PATH, window.location.origin).href;
    await Promise.all(
      registrations
        .filter((registration) => registration.scope.startsWith(scopeRoot))
        .map((registration) => registration.unregister()),
    );
  } catch (error) {
    // Service-worker cleanup is defensive. Do not prevent the editor from
    // starting if the browser blocks registration enumeration.
    console.warn('[OmniDocs] rHWP service worker cleanup failed:', error);
  }
}
