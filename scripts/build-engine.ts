#!/usr/bin/env bun
/** Bundles the chrome-free sync engine into dist-engine/engine.mjs for plain Node or Bun. */
import { mkdirSync, rmSync } from 'node:fs';
import { basename, join, resolve } from 'node:path';

const root = resolve(import.meta.dir, '..');
const out = join(root, 'dist-engine');
if (basename(out) !== 'dist-engine') throw new Error(`Refusing to clean ${out}`);

rmSync(out, { recursive: true, force: true });
mkdirSync(out, { recursive: true });

const result = await Bun.build({
  entrypoints: [join(root, 'src/core/engine.ts')],
  outdir: out,
  target: 'node',
  format: 'esm',
  naming: 'engine.mjs',
});
if (!result.success) {
  for (const log of result.logs) console.error(log);
  process.exit(1);
}
console.log(`Built ${join(out, 'engine.mjs')}`);
