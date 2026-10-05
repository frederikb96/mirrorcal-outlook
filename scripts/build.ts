#!/usr/bin/env bun
/** Bundles the extension into dist/: service worker, options page, manifest and icons. */
import { cpSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { basename, join, resolve } from 'node:path';
import pkg from '../package.json' with { type: 'json' };
import { buildManifest } from '../src/manifest.ts';

const root = resolve(import.meta.dir, '..');
const dist = join(root, 'dist');
if (basename(dist) !== 'dist') throw new Error(`Refusing to clean ${dist}`);

rmSync(dist, { recursive: true, force: true });
mkdirSync(dist, { recursive: true });

const result = await Bun.build({
  entrypoints: [
    join(root, 'src/extension/service-worker.ts'),
    join(root, 'src/extension/options.ts'),
  ],
  outdir: dist,
  target: 'browser',
  format: 'esm',
  naming: '[name].js',
});
if (!result.success) {
  for (const log of result.logs) console.error(log);
  process.exit(1);
}

cpSync(join(root, 'src/extension/options.html'), join(dist, 'options.html'));
cpSync(join(root, 'src/extension/options.css'), join(dist, 'options.css'));
cpSync(join(root, 'icons'), join(dist, 'icons'), { recursive: true });
writeFileSync(
  join(dist, 'manifest.json'),
  `${JSON.stringify(buildManifest(pkg.version), null, 2)}\n`,
);
console.log(`Built ${pkg.version} into ${dist}`);
