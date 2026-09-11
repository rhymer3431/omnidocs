import { access, readFile } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '..');
const studioRoot = join(root, 'public', 'rhwp');
const indexHtml = await readFile(join(studioRoot, 'index.html'), 'utf8');

if (/vite-plugin-pwa:register-sw|registerSW\.js/.test(indexHtml)) {
  throw new Error('Embedded rHWP Studio must not register the upstream service worker.');
}

const entryMatch = indexHtml.match(/src=["']\/rhwp\/assets\/(index-[^"']+\.js)["']/);
if (!entryMatch) throw new Error('Could not find the rHWP Studio entry bundle.');

const entryPath = join(studioRoot, 'assets', entryMatch[1]);
const entry = await readFile(entryPath, 'utf8');
if (!entry.includes('hwpctrl')) throw new Error('Vendored Studio does not advertise hwpctrl support.');

const lazyChunks = new Set(
  [...entry.matchAll(/["'`]\.\/(studio-plugin-[^"'`]+\.js)["'`]/g)].map((match) => match[1]),
);
if (!lazyChunks.size) throw new Error('hwpctrl lazy plugin chunk reference was not found.');

for (const chunk of lazyChunks) {
  const path = join(studioRoot, 'assets', chunk);
  await access(path);
  const text = await readFile(path, 'utf8');
  if (!text.includes('hwpctrl')) {
    throw new Error(`${basename(path)} does not look like the hwpctrl plugin bundle.`);
  }
}

console.log(`rHWP Studio embed OK: hwpctrl chunks=${[...lazyChunks].join(', ')}`);
