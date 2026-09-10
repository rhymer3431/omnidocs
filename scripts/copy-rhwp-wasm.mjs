import { copyFile, mkdir } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '..');
const source = resolve(root, 'node_modules/@rhwp/core/rhwp_bg.wasm');
const targetDir = resolve(root, 'public/vendor/rhwp');
const target = resolve(targetDir, 'rhwp_bg.wasm');

await mkdir(targetDir, { recursive: true });
await copyFile(source, target);
console.log(`Copied ${source} -> ${target}`);
