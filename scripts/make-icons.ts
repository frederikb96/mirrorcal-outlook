#!/usr/bin/env bun
/** Regenerates icons/icon-{16,48,128}.png: a calendar page on a rounded blue square. */
import { mkdirSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { deflateSync } from 'node:zlib';

type Rgba = [number, number, number, number];

const BLUE: Rgba = [43, 108, 176, 255];
const WHITE: Rgba = [255, 255, 255, 255];
const ORANGE: Rgba = [237, 137, 54, 255];
const GREY: Rgba = [160, 174, 192, 255];

function inRoundRect(
  x: number,
  y: number,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  r: number,
) {
  if (x < x0 || x > x1 || y < y0 || y > y1) return false;
  const cx = Math.min(Math.max(x, x0 + r), x1 - r);
  const cy = Math.min(Math.max(y, y0 + r), y1 - r);
  return (x - cx) ** 2 + (y - cy) ** 2 <= r * r;
}

function sample(x: number, y: number): Rgba | null {
  if (!inRoundRect(x, y, 0, 0, 1, 1, 0.22)) return null;
  if (!inRoundRect(x, y, 0.18, 0.2, 0.82, 0.82, 0.07)) return BLUE;
  if (y < 0.4) return ORANGE;
  const col = Math.floor((x - 0.24) / 0.2);
  const row = Math.floor((y - 0.46) / 0.15);
  const inCellX = (x - 0.24) % 0.2 < 0.13 && col >= 0 && col < 3;
  const inCellY = (y - 0.46) % 0.15 < 0.09 && row >= 0 && row < 2;
  return inCellX && inCellY ? GREY : WHITE;
}

function render(size: number): Uint8Array {
  const grid = 4;
  const raw = new Uint8Array((size * 4 + 1) * size);
  for (let py = 0; py < size; py++) {
    raw[py * (size * 4 + 1)] = 0;
    for (let px = 0; px < size; px++) {
      const acc: Rgba = [0, 0, 0, 0];
      for (let sy = 0; sy < grid; sy++) {
        for (let sx = 0; sx < grid; sx++) {
          const c = sample((px + (sx + 0.5) / grid) / size, (py + (sy + 0.5) / grid) / size);
          if (c) {
            acc[0] += c[0];
            acc[1] += c[1];
            acc[2] += c[2];
            acc[3] += 255;
          }
        }
      }
      const covered = acc[3] / 255;
      const o = py * (size * 4 + 1) + 1 + px * 4;
      raw[o] = covered ? Math.round(acc[0] / covered) : 0;
      raw[o + 1] = covered ? Math.round(acc[1] / covered) : 0;
      raw[o + 2] = covered ? Math.round(acc[2] / covered) : 0;
      raw[o + 3] = Math.round(acc[3] / (grid * grid));
    }
  }
  return raw;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const out = Buffer.alloc(body.length + 8);
  out.writeUInt32BE(data.length, 0);
  body.copy(out, 4);
  out.writeUInt32BE(Bun.hash.crc32(body), body.length + 4);
  return out;
}

function png(size: number): Uint8Array {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr.set([8, 6, 0, 0, 0], 8);
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(render(size))),
    chunk('IEND', new Uint8Array()),
  ]);
}

const dir = join(resolve(import.meta.dir, '..'), 'icons');
mkdirSync(dir, { recursive: true });
for (const size of [16, 48, 128]) writeFileSync(join(dir, `icon-${size}.png`), png(size));
