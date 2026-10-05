import { describe, expect, test } from 'bun:test';
import {
  buildEventIcs,
  contentHash,
  escapeText,
  foldLine,
  readRemoteInfo,
  unfoldLines,
} from '../src/core/ics.ts';
import type { MirrorEvent } from '../src/core/types.ts';

const encoder = new TextEncoder();
const now = new Date('2026-10-05T12:00:00Z');

function event(overrides: Partial<MirrorEvent> = {}): MirrorEvent {
  return {
    summary: 'Planning',
    allDay: false,
    start: '2026-10-06T08:00:00',
    end: '2026-10-06T09:00:00',
    showAs: 'Busy',
    attendees: [],
    ...overrides,
  };
}

const meta = { uid: 'mirrorcal-abc', owner: 'outlook', now };

describe('folding', () => {
  test('no physical line exceeds 75 octets and unfolding restores the line', () => {
    const line = `DESCRIPTION:${'word '.repeat(60)}`;
    const folded = foldLine(line);
    const physical = folded.split('\r\n');
    expect(physical.length).toBeGreaterThan(1);
    for (const p of physical) expect(encoder.encode(p).length).toBeLessThanOrEqual(75);
    for (const p of physical.slice(1)) expect(p.startsWith(' ')).toBe(true);
    expect(unfoldLines(folded)).toEqual([line]);
  });

  test('never splits a multi-byte character', () => {
    const line = `SUMMARY:${'日本語のタイトル🙂'.repeat(12)}`;
    const folded = foldLine(line);
    for (const p of folded.split('\r\n')) expect(encoder.encode(p).length).toBeLessThanOrEqual(75);
    expect(unfoldLines(folded)).toEqual([line]);
  });

  test('a line of exactly 75 octets stays on one line', () => {
    const line = `X:${'a'.repeat(73)}`;
    expect(foldLine(line)).toBe(line);
    expect(foldLine(`${line}a`).split('\r\n')).toHaveLength(2);
  });
});

describe('escaping', () => {
  test('escapes backslash, semicolon, comma and newlines', () => {
    const input = ['a,b;c', 'd', 'e'].join('\n') + String.raw` back\slash`;
    expect(escapeText(input)).toBe(String.raw`a\,b\;c\nd\ne back\\slash`);
  });

  test('a subject cannot inject further properties', async () => {
    const { ics } = await buildEventIcs(
      event({ summary: 'x\r\nBEGIN:VALARM\r\nACTION:AUDIO' }),
      meta,
    );
    expect(unfoldLines(ics)).not.toContain('BEGIN:VALARM');
    expect(unfoldLines(ics).filter((l) => l.startsWith('SUMMARY:'))).toEqual([
      String.raw`SUMMARY:x\nBEGIN:VALARM\nACTION:AUDIO`,
    ]);
  });
});

describe('generated ICS', () => {
  test('uses CRLF everywhere and ends with a line break', async () => {
    const { ics } = await buildEventIcs(event({ description: 'one\ntwo' }), meta);
    expect(ics.endsWith('\r\n')).toBe(true);
    expect(ics.replace(/\r\n/g, '')).not.toMatch(/[\r\n]/);
  });

  test('timed events are UTC, all-day events are dates', async () => {
    const timed = await buildEventIcs(event(), meta);
    expect(timed.ics).toContain('DTSTART:20261006T080000Z\r\n');
    expect(timed.ics).toContain('DTEND:20261006T090000Z\r\n');
    const allDay = await buildEventIcs(
      event({ allDay: true, start: '2026-10-12', end: '2026-10-15' }),
      meta,
    );
    expect(allDay.ics).toContain('DTSTART;VALUE=DATE:20261012\r\n');
    expect(allDay.ics).toContain('DTEND;VALUE=DATE:20261015\r\n');
  });

  test('carries owner and hash, writes no alarm', async () => {
    const { ics, hash } = await buildEventIcs(event(), meta);
    const lines = unfoldLines(ics);
    expect(lines).toContain('X-MIRRORCAL-OWNER:outlook');
    expect(lines).toContain(`X-MIRRORCAL-HASH:${hash}`);
    expect(ics).not.toContain('VALARM');
  });

  test('the stored hash equals the hash recomputed from the text', async () => {
    const { ics, hash } = await buildEventIcs(event({ description: 'x'.repeat(300) }), meta);
    expect(await contentHash(ics)).toBe(hash);
  });
});

describe('content hash', () => {
  test('ignores the time stamp, line folding and property order', async () => {
    const a = await buildEventIcs(event({ description: 'x'.repeat(200) }), meta);
    const b = await buildEventIcs(event({ description: 'x'.repeat(200) }), {
      ...meta,
      now: new Date('2027-01-01T00:00:00Z'),
    });
    expect(a.ics).not.toBe(b.ics);
    expect(a.hash).toBe(b.hash);
    const lines = unfoldLines(a.ics);
    const reordered = [...lines.slice(0, 5), ...lines.slice(5, -3).reverse(), ...lines.slice(-3)];
    expect(await contentHash(reordered.join('\n'))).toBe(a.hash);
  });

  test('changes with any written field', async () => {
    const base = (await buildEventIcs(event(), meta)).hash;
    const variants: Partial<MirrorEvent>[] = [
      { summary: 'Other' },
      { start: '2026-10-06T08:30:00' },
      { end: '2026-10-06T09:30:00' },
      { location: 'Room' },
      { description: 'd' },
      { showAs: 'Free' },
    ];
    for (const v of variants) {
      expect((await buildEventIcs(event(v), meta)).hash).not.toBe(base);
    }
    expect((await buildEventIcs(event(), { ...meta, owner: 'other' })).hash).not.toBe(base);
    expect((await buildEventIcs(event(), { ...meta, uid: 'mirrorcal-xyz' })).hash).not.toBe(base);
  });
});

describe('reading server-side ICS', () => {
  test('extracts identity and a label across folded lines', async () => {
    const summary = `Long ${'title '.repeat(30)}end, with; punctuation`;
    const { ics } = await buildEventIcs(event({ summary }), meta);
    expect(readRemoteInfo(ics)).toEqual({
      uid: 'mirrorcal-abc',
      owner: 'outlook',
      summary,
      start: '20261006T080000Z',
    });
  });

  test('objects without our marker have no owner', () => {
    const foreign =
      'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:abc\r\nSUMMARY:Dentist\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n';
    expect(readRemoteInfo(foreign).owner).toBeUndefined();
    expect(readRemoteInfo('not a calendar')).toEqual({});
  });

  test('a marker outside the event does not count', () => {
    const tricky =
      'BEGIN:VCALENDAR\r\nX-MIRRORCAL-OWNER:outlook\r\nBEGIN:VEVENT\r\nUID:abc\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n';
    expect(readRemoteInfo(tricky).owner).toBeUndefined();
  });
});
