import { beforeEach, describe, expect, test } from 'bun:test';
import { OUTLOOK_HOSTS } from '../src/core/constants.ts';
import { type RunOptions, resetMirror, runSync } from '../src/core/engine.ts';
import { unfoldLines } from '../src/core/ics.ts';
import type { Settings } from '../src/core/types.ts';
import { FakeCalDav, routeFetch } from './helpers/fake-caldav.ts';
import { FakeOutlook } from './helpers/fake-outlook.ts';
import {
  COLLECTION,
  NOW,
  outlookId,
  rawEvent,
  sampleWeek,
  TOKEN,
  testSettings,
} from './helpers/fixtures.ts';

let outlook: FakeOutlook;
let dav: FakeCalDav;

beforeEach(() => {
  outlook = new FakeOutlook(TOKEN, sampleWeek());
  dav = new FakeCalDav();
});

function run(overrides: Partial<RunOptions> & { settings?: Settings } = {}) {
  return runSync({
    fetch: routeFetch({ 'outlook.cloud.microsoft': outlook.fetch, 'dav.example.test': dav.fetch }),
    token: TOKEN,
    settings: testSettings(),
    now: NOW,
    mode: 'apply',
    ...overrides,
  });
}

const summaries = () =>
  [...dav.objects.values()]
    .map((o) =>
      unfoldLines(o.ics)
        .find((l) => l.startsWith('SUMMARY:'))
        ?.slice(8),
    )
    .sort();

const filters = ['Sport', 'Mittagessen', 'Abwesend'];

describe('mirroring', () => {
  test('first sync writes every non-filtered event once, conditionally', async () => {
    const result = await run({ settings: testSettings({ titleFilters: filters }) });
    expect(result.outcome).toBe('applied');
    expect(result.counts).toMatchObject({
      source: 7,
      filtered: 3,
      mirrored: 4,
      created: 4,
      errors: 0,
    });
    expect(summaries()).toEqual([
      'Café — Überprüfung 日本',
      'Conference day',
      'Offsite',
      'Planning',
    ]);
    const puts = dav.writes;
    expect(puts).toHaveLength(4);
    for (const p of puts) {
      expect(p.method).toBe('PUT');
      expect(p.headers['if-none-match']).toBe('*');
    }
  });

  test('a second sync with nothing changed sends no writes', async () => {
    const settings = testSettings({ titleFilters: filters });
    await run({ settings });
    dav.clearRequests();
    const again = await run({ settings, now: new Date('2026-10-05T18:30:00Z') });
    expect(again.outcome).toBe('applied');
    expect(again.counts).toMatchObject({ unchanged: 4, create: 0, update: 0, delete: 0 });
    expect(dav.writes).toHaveLength(0);
  });

  test('changing one source event is exactly one PUT, removing one exactly one DELETE', async () => {
    const settings = testSettings();
    await run({ settings });
    dav.clearRequests();
    outlook.events[0] = rawEvent({
      id: outlookId(1),
      subject: 'Planning v2',
      start: '2026-10-06T08:00:00',
      end: '2026-10-06T09:00:00',
    });
    await run({ settings });
    expect(dav.writes.map((w) => w.method)).toEqual(['PUT']);
    expect(dav.writes[0]?.headers['if-match']).toMatch(/^"etag-/);
    expect(summaries()).toContain('Planning v2');

    dav.clearRequests();
    outlook.events.splice(1, 1);
    const result = await run({ settings });
    expect(dav.writes.map((w) => w.method)).toEqual(['DELETE']);
    expect(result.counts.deleted).toBe(1);
    expect(dav.objects.size).toBe(6);
  });

  test('a cancelled instance disappears from the destination', async () => {
    await run();
    outlook.events[0] = rawEvent({
      id: outlookId(1),
      start: '2026-10-06T08:00:00',
      end: '2026-10-06T09:00:00',
      cancelled: true,
    });
    const result = await run();
    expect(result.counts.deleted).toBe(1);
    expect(dav.objects.size).toBe(6);
  });

  test('stores exactly the events inside the span', async () => {
    outlook.events.push(
      rawEvent({
        id: 'old',
        subject: 'Too old',
        start: '2026-08-20T08:00:00',
        end: '2026-08-20T09:00:00',
      }),
      rawEvent({
        id: 'far',
        subject: 'Too far',
        start: '2027-05-01T08:00:00',
        end: '2027-05-01T09:00:00',
      }),
      rawEvent({
        id: 'edge',
        subject: 'Last day',
        start: '2027-04-02T08:00:00',
        end: '2027-04-02T09:00:00',
      }),
    );
    const result = await run();
    expect(result.window).toEqual({
      start: '2026-09-05T00:00:00.000Z',
      end: '2027-04-04T00:00:00.000Z',
    });
    expect(summaries()).toContain('Last day');
    expect(summaries()).not.toContain('Too old');
    expect(summaries()).not.toContain('Too far');
  });

  test('events that age out of the window are removed from the destination', async () => {
    const settings = testSettings({ daysBack: 5 });
    outlook.events.push(
      rawEvent({
        id: 'later',
        subject: 'Later',
        start: '2026-11-20T08:00:00',
        end: '2026-11-20T09:00:00',
      }),
    );
    await run({ settings });
    expect(dav.objects.size).toBe(8);
    const result = await run({ settings, now: new Date('2026-10-30T12:00:00Z') });
    expect(result.outcome).toBe('applied');
    expect(result.counts.deleted).toBeGreaterThan(0);
    expect(summaries()).not.toContain('Planning');
    expect(summaries()).toContain('Later');
  });

  test('a span change that shrinks the mirror moves the destination with it', async () => {
    await run();
    const settings = testSettings({ daysAhead: 7 });
    await run({ settings });
    expect(summaries()).toEqual([
      'Café — Überprüfung 日本',
      'Conference day',
      'Mittagessen',
      'Offsite',
      'Planning',
      'Sport',
    ]);
  });

  test('a large calendar syncs in one run and is stable afterwards', async () => {
    outlook.events = Array.from({ length: 300 }, (_, i) =>
      rawEvent({ id: outlookId(i), start: '2026-10-06T08:00:00', end: '2026-10-06T09:00:00' }),
    );
    outlook.pageSize = 100;
    expect((await run()).counts.created).toBe(300);
    dav.clearRequests();
    expect((await run()).counts.unchanged).toBe(300);
    expect(dav.writes).toHaveLength(0);
  });
});

describe('filters and field options', () => {
  test('adding a filter removes the mirrored copies, removing it restores them', async () => {
    await run();
    expect(summaries()).toContain('Sport');
    const withFilter = await run({ settings: testSettings({ titleFilters: ['sport'] }) });
    expect(withFilter.counts).toMatchObject({ deleted: 1, created: 0 });
    expect(summaries()).not.toContain('Sport');
    const without = await run();
    expect(without.counts).toMatchObject({ created: 1, deleted: 0 });
    expect(summaries()).toContain('Sport');
  });

  test('availability exclusion works like a filter', async () => {
    const result = await run({ settings: testSettings({ excludeShowAs: ['Free'] }) });
    expect(result.counts.filtered).toBe(2);
    expect(summaries()).not.toContain('Sport');
  });

  test('toggling a field re-syncs the existing events', async () => {
    await run();
    const text = () => [...dav.objects.values()].map((o) => o.ics).join('');
    expect(text()).not.toContain('Room 4');

    const copy = testSettings();
    copy.policies.location = { mode: 'copy' };
    copy.policies.description = { mode: 'copy' };
    const on = await run({ settings: copy });
    expect(on.counts.updated).toBe(1);
    expect(text()).toContain('LOCATION:Room 4');
    expect(text()).toContain(String.raw`DESCRIPTION:Agenda\, items\; more`);

    const off = await run();
    expect(off.counts.updated).toBe(1);
    expect(text()).not.toContain('Room 4');
    expect(text()).not.toContain('Agenda');
  });

  test('hiding the title replaces it everywhere and a later change updates all', async () => {
    const hidden = testSettings();
    hidden.policies.title = { mode: 'replace', text: 'Busy' };
    await run({ settings: hidden });
    expect(new Set(summaries())).toEqual(new Set(['Busy']));
    const shown = await run();
    expect(shown.counts.updated).toBe(7);
  });

  test('attendees are only fetched and written when asked for', async () => {
    await run();
    expect(outlook.requests.every((r) => !r.url.includes('Attendees'))).toBe(true);
    expect([...dav.objects.values()].some((o) => o.ics.includes('sam@example.test'))).toBe(false);
    const copy = testSettings();
    copy.policies.attendees = 'copy';
    await run({ settings: copy });
    expect(outlook.requests.at(-1)?.url).toContain('Attendees');
    expect([...dav.objects.values()].some((o) => o.ics.includes('sam@example.test'))).toBe(true);
  });
});

describe('the destination is treated as owned, but only by marker', () => {
  test('hand-deleted objects come back', async () => {
    await run();
    const [name] = [...dav.objects.keys()];
    dav.remove(name as string);
    const result = await run();
    expect(result.counts.created).toBe(1);
    expect(dav.objects.size).toBe(7);
  });

  test('hand-edited objects are restored', async () => {
    await run();
    const [name] = [...dav.objects.keys()];
    dav.edit(name as string, (ics) => ics.replace(/SUMMARY:[^\r\n]*/, 'SUMMARY:Scribbled'));
    dav.clearRequests();
    const result = await run();
    expect(result.counts.updated).toBe(1);
    expect(dav.writes).toHaveLength(1);
    expect(summaries()).not.toContain('Scribbled');
  });

  test('objects without our marker are never touched and are reported', async () => {
    const mine =
      'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:dentist\r\nSUMMARY:Dentist\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n';
    const theirs = mine
      .replace('UID:dentist', 'UID:other')
      .replace('SUMMARY:Dentist', 'SUMMARY:Theirs\r\nX-MIRRORCAL-OWNER:someone-else');
    dav.seed('dentist.ics', mine);
    dav.seed('other.ics', theirs);
    const before = [...dav.objects.entries()].map(([k, v]) => [k, v.ics, v.etag]);
    const result = await run({ settings: testSettings({ titleFilters: ['*'] }) });
    expect(result.counts.foreign).toBe(2);
    expect(dav.objects.get('dentist.ics')?.ics).toBe(before[0]?.[1] as string);
    expect(dav.objects.get('other.ics')?.ics).toBe(before[1]?.[1] as string);
    expect(dav.writes).toHaveLength(0);
    const second = await run();
    expect(second.counts.foreign).toBe(2);
    for (const w of dav.writes) {
      expect(w.url).not.toContain('dentist');
      expect(w.url).not.toContain('other');
    }
    expect(dav.objects.get('dentist.ics')?.ics).toBe(before[0]?.[1] as string);
  });

  test('a foreign object on a wanted address is never overwritten', async () => {
    const preview = await run({ mode: 'preview', includeIcs: true });
    const target = preview.plan?.creations[0];
    const squat =
      'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:squatter\r\nSUMMARY:Mine\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n';
    dav.seed(`${target?.uid}.ics`, squat);
    const result = await run({ confirmMassChange: true });
    expect(result.counts.conflicts).toBe(1);
    expect(dav.objects.get(`${target?.uid}.ics`)?.ics).toBe(squat);
    expect(dav.objects.size).toBe(7);
  });

  test('a different owner id does not see, update or delete the first mirror', async () => {
    await run();
    const snapshot = new Map([...dav.objects].map(([k, v]) => [k, v.ics]));
    outlook.events = [];
    const result = await run({ settings: testSettings({ ownerId: 'second' }) });
    expect(result.counts).toMatchObject({ foreign: 7, own: 0, deleted: 0 });
    for (const [k, ics] of snapshot) expect(dav.objects.get(k)?.ics).toBe(ics);
  });
});

describe('safety guards', () => {
  test('an empty read never empties the destination', async () => {
    await run();
    outlook.events = [];
    dav.clearRequests();
    const result = await run();
    expect(result.outcome).toBe('needs-confirmation');
    expect(result.guard?.code).toBe('empty-source');
    expect(dav.writes).toHaveLength(0);
    expect(dav.objects.size).toBe(7);
  });

  test('an empty read can be confirmed explicitly', async () => {
    await run();
    outlook.events = [];
    const result = await run({ confirmMassChange: true });
    expect(result.outcome).toBe('applied');
    expect(dav.objects.size).toBe(0);
  });

  test('an empty first sync into an empty calendar is fine', async () => {
    outlook.events = [];
    expect((await run()).outcome).toBe('applied');
  });

  describe('mass deletion', () => {
    beforeEach(() => {
      outlook.events = Array.from({ length: 30 }, (_, i) =>
        rawEvent({
          id: outlookId(i),
          subject: `Item ${i}`,
          start: '2026-10-06T08:00:00',
          end: '2026-10-06T09:00:00',
        }),
      );
    });

    test('a sweeping filter is previewed, refused, and applied only when confirmed', async () => {
      await run();
      expect(dav.objects.size).toBe(30);
      const sweeping = testSettings({ titleFilters: ['*'] });

      const preview = await run({ settings: sweeping, mode: 'preview' });
      expect(preview.outcome).toBe('preview');
      expect(preview.guard?.code).toBe('mass-deletion');
      expect(preview.plan?.deletions).toHaveLength(30);
      expect(dav.objects.size).toBe(30);

      dav.clearRequests();
      const refused = await run({ settings: sweeping });
      expect(refused.outcome).toBe('needs-confirmation');
      expect(dav.writes).toHaveLength(0);

      const confirmed = await run({
        settings: sweeping,
        confirmMassChange: true,
        confirmedPlanDigest: refused.digest,
      });
      expect(confirmed.outcome).toBe('applied');
      expect(confirmed.counts.deleted).toBe(30);
      expect(dav.objects.size).toBe(0);
    });

    test('a confirmation only holds for the plan it was given for', async () => {
      await run();
      const sweeping = testSettings({ titleFilters: ['*'] });
      const shown = await run({ settings: sweeping, mode: 'preview' });
      dav.remove([...dav.objects.keys()][0] as string);
      const result = await run({
        settings: sweeping,
        confirmMassChange: true,
        confirmedPlanDigest: shown.digest,
      });
      expect(result.outcome).toBe('needs-confirmation');
      expect(result.message).toContain('changed');
      expect(dav.objects.size).toBe(29);
    });

    test('a few deletions need no confirmation', async () => {
      await run();
      outlook.events.splice(0, 5);
      const result = await run();
      expect(result.outcome).toBe('applied');
      expect(result.counts.deleted).toBe(5);
    });
  });
});

describe('failed reads never delete', () => {
  beforeEach(async () => {
    await run();
    dav.clearRequests();
  });

  const faults = [
    { kind: 'status', status: 500 },
    { kind: 'status', status: 503, onPage: 1 },
    { kind: 'status', status: 404 },
    { kind: 'network' },
    { kind: 'network', onPage: 2 },
    { kind: 'body', body: '<html>sign in</html>' },
    { kind: 'body', body: '{"value": null}' },
    { kind: 'next-link', link: 'https://evil.example.test/api/v2.0/me/calendarview', onPage: 0 },
  ] as const;

  for (const fault of faults) {
    test(`${JSON.stringify(fault)}`, async () => {
      outlook.fault = fault;
      const result = await run();
      expect(result.outcome).toBe('read-failed');
      expect(result.error?.source).toBe('outlook');
      expect(dav.requests).toHaveLength(0);
      expect(dav.objects.size).toBe(7);
    });
  }

  test('rejected or missing tokens say so and touch nothing', async () => {
    const expired = await run({ token: 'Bearer eyJhbGciOiJub25lIn0.eyJhdWQiOiJ4In0.c2ln' });
    expect(expired.outcome).toBe('auth-failed');
    const none = await run({ token: '' });
    expect(none.outcome).toBe('auth-failed');
    expect(none.error?.code).toBe('no-token');
    expect(dav.requests).toHaveLength(0);
    expect(dav.objects.size).toBe(7);
  });
});

describe('failed destination', () => {
  test('rejected credentials write nothing', async () => {
    const result = await run({ settings: testSettings({ password: 'wrong' }) });
    expect(result.outcome).toBe('destination-failed');
    expect(result.error).toMatchObject({ source: 'caldav', code: 'auth' });
    expect(dav.writes).toHaveLength(0);
  });

  test('a listing that fails writes nothing', async () => {
    dav.failWhen = (method) => (method === 'REPORT' ? 500 : undefined);
    const result = await run();
    expect(result.outcome).toBe('destination-failed');
    expect(dav.writes).toHaveLength(0);
  });

  test('one failing object does not stop the others and is retried next run', async () => {
    let failed = false;
    dav.failWhen = (method) => {
      if (method !== 'PUT' || failed) return undefined;
      failed = true;
      return 500;
    };
    const first = await run();
    expect(first.outcome).toBe('applied');
    expect(first.counts).toMatchObject({ created: 6, errors: 1 });
    expect(first.errors[0]).toMatchObject({ op: 'create' });
    const second = await run();
    expect(second.counts).toMatchObject({ created: 1, errors: 0 });
    expect(dav.objects.size).toBe(7);
  });

  test('credentials rejected mid-run stop everything, deletions included', async () => {
    await run();
    outlook.events.splice(0, 2);
    outlook.events.push(
      rawEvent({ id: 'new', start: '2026-10-20T08:00:00', end: '2026-10-20T09:00:00' }),
    );
    dav.failWhen = (method) => (method === 'PUT' ? 401 : undefined);
    dav.clearRequests();
    const result = await run();
    expect(result.errors.length).toBeGreaterThan(0);
    expect(result.message).toContain('401');
    expect(dav.requests.some((r) => r.method === 'DELETE')).toBe(false);
    expect(dav.objects.size).toBe(7);
  });
});

describe('modes and switches', () => {
  test('preview plans without writing anything', async () => {
    const result = await run({ mode: 'preview' });
    expect(result.outcome).toBe('preview');
    expect(result.counts).toMatchObject({ create: 7, created: 0 });
    expect(dav.writes).toHaveLength(0);
    expect(result.plan?.creations.map((c) => c.summary)).toContain('Planning');
    expect(result.plan?.creations[0]?.ics).toBeUndefined();
    const withIcs = await run({ mode: 'preview', includeIcs: true });
    expect(withIcs.plan?.creations[0]?.ics).toContain('BEGIN:VCALENDAR');
  });

  test('the switch stops applying before any request is made', async () => {
    const result = await run({ settings: testSettings({ enabled: false }) });
    expect(result.outcome).toBe('disabled');
    expect(outlook.requests).toHaveLength(0);
    expect(dav.requests).toHaveLength(0);
    const preview = await run({ settings: testSettings({ enabled: false }), mode: 'preview' });
    expect(preview.outcome).toBe('preview');
    expect(dav.writes).toHaveLength(0);
  });

  test('invalid settings are reported before any request', async () => {
    const result = await run({ settings: testSettings({ daysBack: -3, password: '' }) });
    expect(result.outcome).toBe('invalid-settings');
    expect(outlook.requests).toHaveLength(0);
  });
});

describe('what is ever sent', () => {
  test('Outlook only ever receives GETs for the calendar view on its own hosts', async () => {
    await run();
    outlook.events.pop();
    await run();
    await run({ mode: 'preview', settings: testSettings({ titleFilters: ['*'] }) });
    expect(outlook.requests.length).toBeGreaterThan(3);
    for (const r of outlook.requests) {
      const url = new URL(r.url);
      expect(r.method).toBe('GET');
      expect((OUTLOOK_HOSTS as readonly string[]).includes(url.hostname)).toBe(true);
      expect(url.pathname).toBe('/api/v2.0/me/calendarview');
    }
  });

  test('the destination only ever receives requests for its own collection', async () => {
    await run();
    outlook.events.pop();
    await run();
    await resetMirror({ fetch: dav.fetch, settings: testSettings(), mode: 'count' });
    for (const r of dav.requests) {
      expect(['PROPFIND', 'REPORT', 'PUT', 'DELETE']).toContain(r.method);
      expect(r.url.startsWith(COLLECTION)).toBe(true);
    }
  });
});

describe('quiet runs', () => {
  const methods = () => dav.requests.map((r) => r.method);

  test('an unchanged calendar costs one small request and no read of the objects', async () => {
    const first = await run();
    expect(first.state).toBeUndefined();
    const quiet = await run();
    expect(quiet.state).toBeDefined();
    dav.clearRequests();
    const skipped = await run({ lastState: quiet.state });
    expect(skipped).toMatchObject({ outcome: 'applied', skipped: true });
    expect(skipped.counts).toMatchObject({ own: 7, unchanged: 7, mirrored: 7 });
    expect(methods()).toEqual(['PROPFIND']);
  });

  test('an edit on the server moves the change tag and is repaired', async () => {
    await run();
    const { state } = await run();
    const [name] = [...dav.objects.keys()];
    dav.edit(name as string, (ics) => ics.replace(/SUMMARY:[^\r\n]*/, 'SUMMARY:Scribbled'));
    dav.clearRequests();
    const result = await run({ lastState: state });
    expect(result.skipped).toBeUndefined();
    expect(result.counts.updated).toBe(1);
    expect(methods()).toContain('REPORT');
  });

  test('a deletion on the server is noticed the same way', async () => {
    await run();
    const { state } = await run();
    dav.remove([...dav.objects.keys()][0] as string);
    const result = await run({ lastState: state });
    expect(result.counts.created).toBe(1);
  });

  test('a changed source or changed settings read the calendar again', async () => {
    await run();
    const { state } = await run();
    outlook.events[0] = rawEvent({
      id: outlookId(1),
      subject: 'Moved',
      start: '2026-10-06T10:00:00',
      end: '2026-10-06T11:00:00',
    });
    expect((await run({ lastState: state })).counts.updated).toBe(1);
    const again = await run();
    const hidden = testSettings();
    hidden.policies.title = { mode: 'replace', text: 'Busy' };
    const result = await run({ settings: hidden, lastState: again.state });
    expect(result.skipped).toBeUndefined();
    expect(result.counts.updated).toBe(7);
  });

  test('a run that wrote something leaves no state, so the next one verifies', async () => {
    const first = await run();
    expect(first.counts.created).toBe(7);
    expect(first.state).toBeUndefined();
  });

  test('previews always read everything', async () => {
    await run();
    const { state } = await run();
    dav.clearRequests();
    const preview = await run({ mode: 'preview', lastState: state });
    expect(preview.skipped).toBeUndefined();
    expect(methods()).toEqual(['REPORT']);
  });

  test('a server without a change tag is read in full every time', async () => {
    await run();
    const { state } = await run();
    dav.failWhen = (method) => (method === 'PROPFIND' ? 501 : undefined);
    dav.clearRequests();
    const result = await run({ lastState: state });
    expect(result.skipped).toBeUndefined();
    expect(result.state).toBeUndefined();
    expect(methods()).toContain('REPORT');
  });

  test('state from another calendar or owner is not trusted', async () => {
    await run();
    const { state } = await run();
    const other = await run({ settings: testSettings({ ownerId: 'second' }), lastState: state });
    expect(other.skipped).toBeUndefined();
  });
});

describe('reset', () => {
  test('counting changes nothing', async () => {
    await run();
    dav.seed(
      'dentist.ics',
      'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:d\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n',
    );
    dav.clearRequests();
    const result = await resetMirror({ fetch: dav.fetch, settings: testSettings(), mode: 'count' });
    expect(result).toMatchObject({ outcome: 'counted', own: 7, foreign: 1, deleted: 0 });
    expect(dav.writes).toHaveLength(0);
  });

  test('deletes every owned object and nothing else', async () => {
    await run();
    const foreignIcs =
      'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:d\r\nSUMMARY:Dentist\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n';
    dav.seed('dentist.ics', foreignIcs);
    const result = await resetMirror({
      fetch: dav.fetch,
      settings: testSettings(),
      mode: 'apply',
      confirmedCount: 7,
    });
    expect(result).toMatchObject({ outcome: 'reset', deleted: 7 });
    expect([...dav.objects.keys()]).toEqual(['dentist.ics']);
    expect(dav.objects.get('dentist.ics')?.ics).toBe(foreignIcs);
  });

  test('refuses when the count differs from the one that was confirmed', async () => {
    await run();
    for (const confirmedCount of [6, 8, undefined]) {
      const result = await resetMirror({
        fetch: dav.fetch,
        settings: testSettings(),
        mode: 'apply',
        confirmedCount,
      });
      expect(result.outcome).toBe('refused');
    }
    expect(dav.objects.size).toBe(7);
  });

  test('a failed listing deletes nothing', async () => {
    await run();
    dav.failWhen = () => 500;
    const result = await resetMirror({
      fetch: dav.fetch,
      settings: testSettings(),
      mode: 'apply',
      confirmedCount: 7,
    });
    expect(result.outcome).toBe('destination-failed');
    expect(dav.objects.size).toBe(7);
  });
});
