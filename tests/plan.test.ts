import { describe, expect, test } from 'bun:test';
import { unfoldLines } from '../src/core/ics.ts';
import { buildDesired } from '../src/core/mapping.ts';
import { planDigest, planSync } from '../src/core/plan.ts';
import type { RemoteObject, SourceEvent } from '../src/core/types.ts';
import { COLLECTION, NOW, outlookId, testSettings } from './helpers/fixtures.ts';

function source(n: number, overrides: Partial<SourceEvent> = {}): SourceEvent {
  return {
    id: outlookId(n),
    subject: `Meeting ${n}`,
    start: '2026-10-06T08:00:00',
    end: '2026-10-06T09:00:00',
    allDay: false,
    showAs: 'Busy',
    ...overrides,
  };
}

const settings = testSettings();
const desiredFor = async (events: SourceEvent[], s = settings) =>
  (await buildDesired(events, s, COLLECTION, NOW)).desired;

/** What the server would hold after the desired objects were written. */
const stored = (desired: Awaited<ReturnType<typeof desiredFor>>): RemoteObject[] =>
  desired.map((d, i) => ({ href: d.href, etag: `"e${i}"`, ics: d.ics }));

const foreign = (name: string, extra = ''): RemoteObject => ({
  href: `${COLLECTION}${name}.ics`,
  etag: '"f"',
  ics: `BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:${name}\r\nSUMMARY:Dentist\r\n${extra}END:VEVENT\r\nEND:VCALENDAR\r\n`,
});

describe('planning', () => {
  test('an empty destination gets everything created', async () => {
    const desired = await desiredFor([source(1), source(2)]);
    const plan = await planSync(desired, [], 'outlook');
    expect(plan.creations).toHaveLength(2);
    expect([plan.updates.length, plan.deletions.length, plan.unchanged]).toEqual([0, 0, 0]);
  });

  test('is idempotent: a second plan over unchanged input is empty', async () => {
    const desired = await desiredFor([source(1), source(2), source(3)]);
    const plan = await planSync(desired, stored(desired), 'outlook');
    expect([plan.creations.length, plan.updates.length, plan.deletions.length]).toEqual([0, 0, 0]);
    expect(plan.unchanged).toBe(3);
    const later = await buildDesired(
      [source(1), source(2), source(3)],
      settings,
      COLLECTION,
      new Date('2027-01-01T00:00:00Z'),
    );
    const again = await planSync(later.desired, stored(desired), 'outlook');
    expect(again.unchanged).toBe(3);
  });

  test('a changed title is exactly one update, at the existing address and etag', async () => {
    const before = await desiredFor([source(1), source(2)]);
    const after = await desiredFor([source(1), source(2, { subject: 'Renamed' })]);
    const plan = await planSync(after, stored(before), 'outlook');
    expect(plan.updates).toHaveLength(1);
    expect(plan.updates[0]?.summary).toBe('Renamed');
    expect(plan.updates[0]?.href).toBe(before[1]?.href);
    expect(plan.updates[0]?.etag).toBe('"e1"');
    expect(plan.creations.length + plan.deletions.length).toBe(0);
    expect(plan.unchanged).toBe(1);
  });

  test('a moved event is an update, not a delete and create', async () => {
    const before = await desiredFor([source(1)]);
    const after = await desiredFor([
      source(1, { start: '2026-10-07T08:00:00', end: '2026-10-07T09:00:00' }),
    ]);
    const plan = await planSync(after, stored(before), 'outlook');
    expect([plan.creations.length, plan.updates.length, plan.deletions.length]).toEqual([0, 1, 0]);
  });

  test('a removed source event is exactly one deletion', async () => {
    const before = await desiredFor([source(1), source(2)]);
    const after = await desiredFor([source(1)]);
    const plan = await planSync(after, stored(before), 'outlook');
    expect(plan.deletions).toHaveLength(1);
    expect(plan.deletions[0]?.summary).toBe('Meeting 2');
    expect(plan.deletions[0]?.etag).toBe('"e1"');
    expect(plan.creations.length + plan.updates.length).toBe(0);
  });

  test('a hand-edited destination object is repaired even though its stored hash is intact', async () => {
    const desired = await desiredFor([source(1)]);
    const remote = stored(desired);
    remote[0] = {
      ...remote[0],
      ics: (remote[0]?.ics ?? '').replace('SUMMARY:Meeting 1', 'SUMMARY:Edited by hand'),
    } as RemoteObject;
    const plan = await planSync(desired, remote, 'outlook');
    expect(plan.updates).toHaveLength(1);
  });

  test('a time edited on the server is repaired', async () => {
    const desired = await desiredFor([source(1)]);
    const remote = stored(desired);
    remote[0] = {
      ...remote[0],
      ics: (remote[0]?.ics ?? '').replace('T080000Z', 'T100000Z'),
    } as RemoteObject;
    expect((await planSync(desired, remote, 'outlook')).updates).toHaveLength(1);
  });

  test('a server that refolds or reorders lines does not cause updates', async () => {
    const desired = await desiredFor([source(1)]);
    const lines = unfoldLines(desired[0]?.ics ?? '');
    const shuffled = [
      lines[0],
      lines[1],
      lines[4],
      ...lines.slice(5, -2).reverse(),
      ...lines.slice(-2),
    ].join('\n');
    const plan = await planSync(
      desired,
      [{ href: desired[0]?.href ?? '', etag: '"e"', ics: shuffled }],
      'outlook',
    );
    expect(plan.unchanged).toBe(1);
  });

  test('a hand-deleted object is created again', async () => {
    const desired = await desiredFor([source(1), source(2)]);
    const plan = await planSync(desired, stored(desired).slice(1), 'outlook');
    expect(plan.creations.map((c) => c.uid)).toEqual([desired[0]?.uid as string]);
  });
});

describe('ownership', () => {
  test('objects without the marker are never updated or deleted, only reported', async () => {
    const desired = await desiredFor([source(1)]);
    const remote = [...stored(desired), foreign('dentist'), foreign('holiday')];
    const plan = await planSync(desired, remote, 'outlook');
    expect(plan.deletions).toHaveLength(0);
    expect(plan.updates).toHaveLength(0);
    expect(plan.foreign.map((f) => f.summary)).toEqual(['Dentist', 'Dentist']);
    expect(plan.ownCount).toBe(1);
    expect(plan.ownHrefs).toEqual([desired[0]?.href as string]);
  });

  test('objects of another owner are foreign', async () => {
    const other = foreign('theirs', 'X-MIRRORCAL-OWNER:someone-else\r\n');
    const plan = await planSync([], [other], 'outlook');
    expect(plan.deletions).toHaveLength(0);
    expect(plan.foreign).toHaveLength(1);
  });

  test('objects that cannot be read as ICS are foreign', async () => {
    const plan = await planSync(
      [],
      [{ href: `${COLLECTION}junk.ics`, etag: '"j"', ics: 'garbage' }],
      'outlook',
    );
    expect(plan.deletions).toHaveLength(0);
    expect(plan.foreign).toHaveLength(1);
  });

  test('a foreign object on a wanted identity blocks creation instead of being overwritten', async () => {
    const desired = await desiredFor([source(1)]);
    const squatter: RemoteObject = foreign('squatter');
    squatter.ics = squatter.ics.replace('UID:squatter', `UID:${desired[0]?.uid}`);
    const byAddress: RemoteObject = { ...foreign('x'), href: desired[0]?.href as string };
    for (const occupant of [squatter, byAddress]) {
      const plan = await planSync(desired, [occupant], 'outlook');
      expect(plan.creations).toHaveLength(0);
      expect(plan.conflicts).toHaveLength(1);
      expect(plan.deletions).toHaveLength(0);
    }
  });

  test('own objects that are no longer wanted are deleted, duplicates collapse to one', async () => {
    const desired = await desiredFor([source(1)]);
    const copy: RemoteObject = {
      href: `${COLLECTION}copy.ics`,
      etag: '"c"',
      ics: desired[0]?.ics ?? '',
    };
    const plan = await planSync(desired, [...stored(desired), copy], 'outlook');
    expect(plan.deletions.map((d) => d.href)).toEqual([`${COLLECTION}copy.ics`]);
    expect(plan.unchanged).toBe(1);
  });
});

describe('digest', () => {
  test('identifies the plan', async () => {
    const a = await desiredFor([source(1), source(2)]);
    const b = await desiredFor([source(1)]);
    const base = await planDigest(await planSync(b, stored(a), 'outlook'));
    expect(base).toBe(await planDigest(await planSync(b, stored(a), 'outlook')));
    expect(base).not.toBe(await planDigest(await planSync(a, stored(a), 'outlook')));
    expect(base).not.toBe(await planDigest(await planSync([], stored(a), 'outlook')));
  });
});
