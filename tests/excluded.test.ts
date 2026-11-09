import { beforeEach, describe, expect, test } from 'bun:test';
import { checkDestination, type RunOptions, resetMirror, runSync } from '../src/core/engine.ts';
import { unfoldLines } from '../src/core/ics.ts';
import type { Settings } from '../src/core/types.ts';
import { FakeCalDav, routeCalendars, routeFetch } from './helpers/fake-caldav.ts';
import { FakeOutlook } from './helpers/fake-outlook.ts';
import {
  EXCLUDED_COLLECTION,
  NOW,
  outlookId,
  rawEvent,
  sampleWeek,
  TOKEN,
  testSettings,
} from './helpers/fixtures.ts';

let outlook: FakeOutlook;
let main: FakeCalDav;
let excluded: FakeCalDav;

beforeEach(() => {
  outlook = new FakeOutlook(TOKEN, sampleWeek());
  main = new FakeCalDav();
  excluded = new FakeCalDav(EXCLUDED_COLLECTION);
  excluded.collection = { kind: 'calendar', name: 'Other' };
});

const filters = ['Sport', 'Mittagessen', 'Abwesend'];

const settings = (overrides: Partial<Settings> = {}) =>
  testSettings({ titleFilters: filters, excludedCaldavUrl: EXCLUDED_COLLECTION, ...overrides });

const davFetch = () => routeCalendars(main, excluded);

function run(overrides: Partial<RunOptions> = {}) {
  return runSync({
    fetch: routeFetch({ 'outlook.cloud.microsoft': outlook.fetch, 'dav.example.test': davFetch() }),
    token: TOKEN,
    settings: settings(),
    now: NOW,
    mode: 'apply',
    ...overrides,
  });
}

const summaries = (dav: FakeCalDav) =>
  [...dav.objects.values()]
    .map((o) =>
      unfoldLines(o.ics)
        .find((l) => l.startsWith('SUMMARY:'))
        ?.slice(8),
    )
    .sort();

describe('the calendar for excluded events', () => {
  test('receives exactly the filtered events, the main calendar the rest', async () => {
    const result = await run();
    expect(result.outcome).toBe('applied');
    expect(result.counts).toMatchObject({ source: 7, filtered: 3, mirrored: 4, created: 4 });
    expect(result.excluded?.counts).toMatchObject({ mirrored: 3, created: 3, errors: 0 });
    expect(summaries(main)).toEqual([
      'Café — Überprüfung 日本',
      'Conference day',
      'Offsite',
      'Planning',
    ]);
    expect(summaries(excluded)).toEqual(['Abwesend', 'Mittagessen', 'Sport']);
  });

  test('availability exclusions go there too', async () => {
    await run({ settings: settings({ titleFilters: [], excludeShowAs: ['Free'] }) });
    expect(summaries(excluded)).toEqual(['Offsite', 'Sport']);
    expect(main.objects.size).toBe(5);
  });

  test('removing a filter moves the event back to the main calendar', async () => {
    await run();
    const moved = await run({ settings: settings({ titleFilters: ['Sport', 'Abwesend'] }) });
    expect(moved.counts.created).toBe(1);
    expect(moved.excluded?.counts.deleted).toBe(1);
    expect(summaries(main)).toContain('Mittagessen');
    expect(summaries(excluded)).toEqual(['Abwesend', 'Sport']);
  });

  test('private events set to not mirror appear in neither calendar', async () => {
    outlook.events.push(
      rawEvent({
        id: outlookId(9),
        subject: 'Sport',
        start: '2026-10-08T07:00:00',
        end: '2026-10-08T08:00:00',
        sensitivity: 'Private',
      }),
    );
    const s = settings();
    s.policies.privateEvents = { mode: 'skip' };
    const result = await run({ settings: s });
    expect(result.counts.filtered).toBe(4);
    expect(result.excluded?.counts.mirrored).toBe(3);
    expect(main.objects.size + excluded.objects.size).toBe(7);
  });

  test('without a second calendar the filtered events are not mirrored and nothing else is asked', async () => {
    const result = await run({ settings: settings({ excludedCaldavUrl: '' }) });
    expect(result.excluded).toBeUndefined();
    expect(main.objects.size).toBe(4);
    expect(excluded.requests).toHaveLength(0);
  });

  test('every write lies inside the calendar it was planned for', async () => {
    await run();
    for (const r of main.writes) expect(main.owns(r.url)).toBe(true);
    for (const r of excluded.writes) expect(r.url.startsWith(EXCLUDED_COLLECTION)).toBe(true);
    expect(excluded.writes).toHaveLength(3);
  });

  test('objects of others in the second calendar are left alone', async () => {
    excluded.seed(
      'hand-made.ics',
      'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:x\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n',
    );
    const held = await run();
    expect(held.outcome).toBe('needs-confirmation');
    expect(held.guard?.code).toBe('foreign-destination');
    expect(held.guard?.message).toStartWith('Calendar for excluded events: ');
    expect(main.writes).toHaveLength(0);
    expect(excluded.writes).toHaveLength(0);
    const confirmed = await run({
      confirmMassChange: true,
      confirmedPlanDigest: held.digest as string,
    });
    expect(confirmed.outcome).toBe('applied');
    expect(excluded.objects.has('hand-made.ics')).toBe(true);
    expect(excluded.objects.size).toBe(4);
  });
});

describe('guards and confirmation', () => {
  test('a mass deletion in either calendar holds every write, and the preview shows both plans', async () => {
    const many = Array.from({ length: 30 }, (_, i) =>
      rawEvent({
        id: outlookId(100 + i),
        subject: 'Sport',
        start: `2026-10-${String(6 + (i % 20)).padStart(2, '0')}T06:00:00`,
        end: `2026-10-${String(6 + (i % 20)).padStart(2, '0')}T07:00:00`,
      }),
    );
    outlook.events.push(...many);
    await run();
    expect(excluded.objects.size).toBe(33);
    outlook.events.push(
      rawEvent({
        id: outlookId(200),
        subject: 'New',
        start: '2026-10-20T08:00:00',
        end: '2026-10-20T09:00:00',
      }),
    );
    main.clearRequests();
    const unfiltered = settings({ titleFilters: ['Mittagessen', 'Abwesend'] });
    const held = await run({ settings: unfiltered });
    expect(held.outcome).toBe('needs-confirmation');
    expect(held.guard?.code).toBe('mass-deletion');
    expect(main.writes).toHaveLength(0);
    expect(held.plan?.creations.length).toBe(32);
    expect(held.excluded?.plan?.deletions.length).toBe(31);

    const preview = await run({ settings: unfiltered, mode: 'preview' });
    expect(preview.digest).toBe(held.digest as string);
    const applied = await run({
      settings: unfiltered,
      confirmMassChange: true,
      confirmedPlanDigest: preview.digest as string,
    });
    expect(applied.outcome).toBe('applied');
    expect(excluded.objects.size).toBe(2);
  });

  test('a failing read of the second calendar stops the run before any write', async () => {
    excluded.failWhen = (method) => (method === 'REPORT' ? 500 : undefined);
    const result = await run();
    expect(result.outcome).toBe('destination-failed');
    expect(result.message).toStartWith('Calendar for excluded events: ');
    expect(main.writes).toHaveLength(0);
  });
});

describe('quiet runs', () => {
  test('two unchanged calendars cost one change-tag request each', async () => {
    await run();
    const quiet = await run();
    expect(quiet.state).toBeDefined();
    expect(quiet.excluded?.state).toBeDefined();
    main.clearRequests();
    excluded.clearRequests();
    const skipped = await run({
      lastState: quiet.state as never,
      lastExcludedState: quiet.excluded?.state as never,
    });
    expect(skipped).toMatchObject({ outcome: 'applied', skipped: true });
    expect(skipped.excluded?.counts).toMatchObject({ own: 3, unchanged: 3 });
    expect(main.requests.map((r) => r.method)).toEqual(['PROPFIND']);
    expect(excluded.requests.map((r) => r.method)).toEqual(['PROPFIND']);
  });

  test('a change in one calendar reads only that one', async () => {
    await run();
    const quiet = await run();
    excluded.remove([...excluded.objects.keys()][0] as string);
    main.clearRequests();
    const result = await run({
      lastState: quiet.state as never,
      lastExcludedState: quiet.excluded?.state as never,
    });
    expect(result.excluded?.counts.created).toBe(1);
    expect(main.requests.map((r) => r.method)).toEqual(['PROPFIND']);
    expect(result.state).toEqual(quiet.state as never);
  });
});

describe('reset and check', () => {
  test('a reset counts and deletes the owned events of both calendars, and nothing else', async () => {
    await run();
    excluded.seed(
      'keep.ics',
      'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:keep\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n',
    );
    const reset = (mode: 'count' | 'apply', confirmedCount?: number) =>
      resetMirror({
        fetch: davFetch(),
        settings: settings(),
        mode,
        ...(confirmedCount === undefined ? {} : { confirmedCount }),
      });
    const counted = await reset('count');
    expect(counted).toMatchObject({ outcome: 'counted', own: 7, foreign: 1 });
    expect(main.writes).toHaveLength(4);
    expect((await reset('apply', 4)).outcome).toBe('refused');
    const done = await reset('apply', 7);
    expect(done).toMatchObject({ outcome: 'reset', deleted: 7 });
    expect(main.objects.size).toBe(0);
    expect([...excluded.objects.keys()]).toEqual(['keep.ics']);
  });

  test('the check names both calendars and refuses a second one that holds no events', async () => {
    const ok = await checkDestination({ fetch: davFetch(), settings: settings() });
    expect(ok).toMatchObject({ outcome: 'ok', name: 'Work' });
    expect(ok.message).toBe(
      "Connected to calendar 'Work', and to calendar 'Other' for excluded events",
    );
    excluded.collection = { kind: 'tasks', name: 'Tasks' };
    const refused = await checkDestination({ fetch: davFetch(), settings: settings() });
    expect(refused.outcome).toBe('not-a-calendar');
    expect(refused.message).toBe(
      'Calendar for excluded events: This calendar does not hold events',
    );
  });
});
