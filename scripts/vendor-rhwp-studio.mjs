import { access, mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, extname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const VERSION = '0.8.6';
const ORIGIN = 'https://edwardkim.github.io';
const ROOT_PATH = '/rhwp/';
const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '..');
const outputRoot = resolve(root, 'public/rhwp');
const queue = [new URL(ROOT_PATH, ORIGIN)];
const seen = new Set();
let downloaded = 0;

const allowedExtensions = new Set([
  '', '.html', '.js', '.css', '.json', '.webmanifest', '.ico', '.png', '.svg', '.wasm',
  '.woff', '.woff2', '.ttf', '.otf', '.map',
]);

function localPath(url) {
  let pathname = decodeURIComponent(url.pathname);
  if (!pathname.startsWith(ROOT_PATH)) {
    throw new Error(`Refusing to vendor path outside ${ROOT_PATH}: ${pathname}`);
  }
  if (pathname.endsWith('/')) pathname += 'index.html';
  return join(outputRoot, pathname.slice(ROOT_PATH.length));
}

function enqueue(raw, from) {
  if (!raw || raw.startsWith('data:') || raw.startsWith('blob:') || raw.startsWith('#')) return;
  let url;
  try {
    if (raw.startsWith('fonts/') || raw.startsWith('./fonts/') || raw.startsWith('../fonts/')) {
      url = new URL(`${ROOT_PATH}fonts/${raw.split('fonts/')[1]}`, ORIGIN);
    } else if (raw.startsWith('images/') || raw.startsWith('./images/') || raw.startsWith('../images/')) {
      url = new URL(`${ROOT_PATH}images/${raw.split('images/')[1]}`, ORIGIN);
    } else {
      url = new URL(raw, from);
    }
  } catch {
    return;
  }
  if (url.origin !== ORIGIN || !url.pathname.startsWith(ROOT_PATH)) return;
  const extension = extname(url.pathname).toLowerCase();
  if (!allowedExtensions.has(extension)) return;
  url.hash = '';
  queue.push(url);
}

function discover(text, from) {
  const patterns = [
    /(?:src|href)=["']([^"']+)["']/g,
    /url\(["']?([^)'"\s]+)["']?\)/g,
    /["'](\/rhwp\/[^"'\s)]+)["']/g,
    /["']((?:\.\.\/|\.\/)?fonts\/[^"'\s)]+)["']/g,
    /["']((?:\.\.\/|\.\/)?images\/[^"'\s)]+)["']/g,
    // Vite lazy chunks, including `import(`./studio-plugin-*.js`)`.
    /\bimport\s*\(\s*["'`]([^"'`]+)["'`]\s*\)/g,
    // Vite PWA precache manifest entries. These also include lazy chunks that
    // are not referenced by ordinary src/href attributes in index.html.
    /\burl\s*:\s*["'`]([^"'`]+)["'`]/g,
    /["'`]((?:\.\.\/|\.\/)?assets\/[^"'`\s)]+)["'`]/g,
  ];
  for (const pattern of patterns) {
    for (const match of text.matchAll(pattern)) enqueue(match[1], from);
  }

  // Vite can emit asset URLs inside JS expressions rather than quoted literals.
  for (const match of text.matchAll(/\/rhwp\/[A-Za-z0-9_./?=&%:+~-]+/g)) {
    enqueue(match[0], from);
  }
}

while (queue.length) {
  const url = queue.shift();
  if (seen.has(url.href)) continue;
  seen.add(url.href);

  const response = await fetch(url, { redirect: 'follow' });
  if (!response.ok) {
    console.warn(`Skipping ${url.href}: ${response.status}`);
    continue;
  }

  const target = localPath(url);
  await mkdir(dirname(target), { recursive: true });
  const contentType = response.headers.get('content-type') ?? '';
  const extension = extname(url.pathname).toLowerCase();
  const isText = /text|javascript|json|xml|manifest/.test(contentType)
    || ['.html', '.js', '.css', '.json', '.webmanifest', ''].includes(extension);

  if (isText) {
    const text = await response.text();
    await writeFile(target, text);
    discover(text, url);
  } else {
    await writeFile(target, new Uint8Array(await response.arrayBuffer()));
  }
  downloaded += 1;
  console.log(`Vendored ${url.pathname} -> ${relative(root, target)}`);
}

const builtJs = [...seen].find((value) => /\/assets\/index-[^/]+\.js$/.test(new URL(value).pathname));
if (builtJs) {
  const text = await readFile(localPath(new URL(builtJs)), 'utf8');
  if (!text.includes(VERSION)) {
    throw new Error(`Published rHWP Studio does not appear to match pinned version ${VERSION}.`);
  }

  const requiredLazyChunks = [
    ...text.matchAll(/(?:import\s*\(\s*)?["'`]\.\/(studio-plugin-[^"'`]+\.js)["'`]/g),
  ].map((match) => match[1]);
  if (!requiredLazyChunks.length) {
    throw new Error('Published rHWP Studio bundle does not expose the expected hwpctrl plugin chunk.');
  }
  for (const fileName of requiredLazyChunks) {
    await access(join(outputRoot, 'assets', fileName));
  }
}

// OmniDocs embeds Studio in an iframe and owns its lifecycle. Registering the
// upstream PWA service worker inside that iframe can keep an older Studio
// bundle cached after an OmniDocs update, which in turn produces misleading
// plugin errors. Keep the vendored runtime network-local but disable PWA
// registration for the embedded copy.
const indexPath = join(outputRoot, 'index.html');
let indexHtml = await readFile(indexPath, 'utf8');
indexHtml = indexHtml
  .replace(/<script\b[^>]*id=["']vite-plugin-pwa:register-sw["'][^>]*><\/script>/gi, '')
  .replace(/<link\b[^>]*rel=["']manifest["'][^>]*>/gi, '');
await writeFile(indexPath, indexHtml);

await writeFile(
  join(outputRoot, '.omnidocs-vendor.json'),
  JSON.stringify({
    source: `${ORIGIN}${ROOT_PATH}`,
    rhwpVersion: VERSION,
    vendoredAt: new Date().toISOString(),
    files: downloaded,
  }, null, 2),
);

console.log(`rHWP Studio ${VERSION} self-host bundle ready: ${downloaded} resources.`);
