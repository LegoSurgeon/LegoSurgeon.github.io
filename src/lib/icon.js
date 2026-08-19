import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';

// createRequire walks node_modules up the tree from wherever this file is compiled,
// so it resolves correctly whether running from src/ or dist/.
const _require = createRequire(import.meta.url);
const iconsDir = path.join(path.dirname(_require.resolve('lucide-static/package.json')), 'icons');

const aliases = {
  'check-circle-2': 'circle-check-big',
  'alert-circle': 'circle-alert',
  'alert-triangle': 'triangle-alert',
  'loader-2': 'loader-circle',
};

const cache = new Map();

export function getIconSvg(name) {
  const resolved = aliases[name] ?? name;
  if (cache.has(resolved)) return cache.get(resolved);
  const file = path.join(iconsDir, `${resolved}.svg`);
  const raw = fs.readFileSync(file, 'utf-8').replace(/<!--[\s\S]*?-->/g, '').trim();
  cache.set(resolved, raw);
  return raw;
}
