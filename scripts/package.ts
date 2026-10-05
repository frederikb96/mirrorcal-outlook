#!/usr/bin/env bun
/** Zips dist/ into mirrorcal-outlook.zip with manifest.json at the archive root. */
import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';
import { deflateRawSync } from 'node:zlib';

const root = resolve(import.meta.dir, '..');
const dist = join(root, 'dist');
const output = join(root, 'mirrorcal-outlook.zip');

function files(dir: string): string[] {
  return readdirSync(dir)
    .sort()
    .flatMap((name) => {
      const path = join(dir, name);
      return statSync(path).isDirectory() ? files(path) : [path];
    });
}

const names = files(dist).map((path) => relative(dist, path).split('\\').join('/'));
if (!names.includes('manifest.json'))
  throw new Error('dist/manifest.json is missing; run build first');

// Fixed time stamps keep the archive identical for identical input.
const DOS_TIME = 0;
const DOS_DATE = (0 << 9) | (1 << 5) | 1;
const encoder = new TextEncoder();

const chunks: Uint8Array[] = [];
const central: Uint8Array[] = [];
let offset = 0;

function header(size: number, fill: (view: DataView) => void): Uint8Array {
  const bytes = new Uint8Array(size);
  fill(new DataView(bytes.buffer));
  return bytes;
}

for (const name of names) {
  const data = readFileSync(join(dist, name));
  const packed = deflateRawSync(data);
  const crc = Bun.hash.crc32(data);
  const nameBytes = encoder.encode(name);
  const local = header(30, (v) => {
    v.setUint32(0, 0x04034b50, true);
    v.setUint16(4, 20, true);
    v.setUint16(6, 0x0800, true);
    v.setUint16(8, 8, true);
    v.setUint16(10, DOS_TIME, true);
    v.setUint16(12, DOS_DATE, true);
    v.setUint32(14, crc, true);
    v.setUint32(18, packed.length, true);
    v.setUint32(22, data.length, true);
    v.setUint16(26, nameBytes.length, true);
  });
  central.push(
    header(46, (v) => {
      v.setUint32(0, 0x02014b50, true);
      v.setUint16(4, 0x031e, true);
      v.setUint16(6, 20, true);
      v.setUint16(8, 0x0800, true);
      v.setUint16(10, 8, true);
      v.setUint16(12, DOS_TIME, true);
      v.setUint16(14, DOS_DATE, true);
      v.setUint32(16, crc, true);
      v.setUint32(20, packed.length, true);
      v.setUint32(24, data.length, true);
      v.setUint16(28, nameBytes.length, true);
      v.setUint32(38, (0o100644 << 16) >>> 0, true);
      v.setUint32(42, offset, true);
    }),
    nameBytes,
  );
  chunks.push(local, nameBytes, packed);
  offset += local.length + nameBytes.length + packed.length;
}

const centralSize = central.reduce((sum, c) => sum + c.length, 0);
const end = header(22, (v) => {
  v.setUint32(0, 0x06054b50, true);
  v.setUint16(8, names.length, true);
  v.setUint16(10, names.length, true);
  v.setUint32(12, centralSize, true);
  v.setUint32(16, offset, true);
});

writeFileSync(output, Buffer.concat([...chunks, ...central, end]));
console.log(`Wrote ${output} (${names.length} files)`);
