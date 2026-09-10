const DEFAULT_LOCAL_STUDIO_PATH = '/rhwp/index.html';

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
