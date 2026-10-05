import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import type { Settings } from '../src/core/types.ts';
import type { RunRecord } from '../src/extension/messages.ts';
import { runResetFlight, runSyncFlight } from '../src/extension/runner.ts';
import { loadStatus, loadToken, saveSettings, saveToken } from '../src/extension/storage.ts';
import { FakeCalDav, routeFetch } from './helpers/fake-caldav.ts';
import { installFakeChrome } from './helpers/fake-chrome.ts';
import { FakeOutlook } from './helpers/fake-outlook.ts';
import { outlookId, rawEvent, sampleWeek, TOKEN, testSettings } from './helpers/fixtures.ts';

const realFetch = globalThis.fetch;
let fake: ReturnType<typeof installFakeChrome>;
let outlook: FakeOutlook;
let dav: FakeCalDav;

const FAR_FUTURE = 4_000_000_000;

beforeEach(async () => {
  fake = installFakeChrome();
  outlook = new FakeOutlook(TOKEN, sampleWeek());
  dav = new FakeCalDav();
  globalThis.fetch = routeFetch({
    'outlook.cloud.microsoft': outlook.fetch,
    'dav.example.test': dav.fetch,
  }) as typeof fetch;
});

afterEach(() => {
  globalThis.fetch = realFetch;
});

async function configure(settings: Settings = testSettings()) {
  await saveSettings(settings);
  await saveToken({ token: TOKEN, exp: FAR_FUTURE });
}

async function sync(mode: 'apply' | 'preview' = 'apply', confirm?: { digest: string }) {
  const admitted = await runSyncFlight({
    trigger: 'manual',
    mode,
    ...(confirm ? { confirm } : {}),
  });
  if (admitted.busy) throw new Error('busy');
  return admitted.value;
}

describe('before anything can run', () => {
  test('an unconfigured extension says so and makes no request', async () => {
    const { record } = await sync();
    expect(record.outcome).toBe('not-configured');
    expect(outlook.requests).toHaveLength(0);
    expect((await loadStatus()).lastRun?.outcome).toBe('not-configured');
  });

  test('without an Outlook session it says to open Outlook and deletes nothing', async () => {
    await saveSettings(testSettings());
    dav.seed('keep.ics', 'BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n');
    const { record } = await sync();
    expect(record.outcome).toBe('no-token');
    expect(record.message).toContain('open Outlook');
    expect(outlook.requests).toHaveLength(0);
    expect(dav.requests).toHaveLength(0);
  });

  test('an expired session counts as none', async () => {
    await saveSettings(testSettings());
    await saveToken({ token: TOKEN, exp: 1_000 });
    expect((await sync()).record.outcome).toBe('no-token');
  });

  test('a missing server permission is reported before any request', async () => {
    await configure();
    fake.state.permissionGranted = false;
    expect((await sync()).record.outcome).toBe('no-permission');
    expect(outlook.requests).toHaveLength(0);
  });
});

describe('a run', () => {
  test('mirrors, records counts and the time of the last success', async () => {
    await configure();
    const { record } = await sync();
    expect(record.outcome).toBe('applied');
    expect(record.counts?.created).toBe(7);
    const status = await loadStatus();
    expect(status.lastSuccessAt).toBe(record.finishedAt);
    expect(status.log).toHaveLength(1);
    expect(dav.objects.size).toBe(7);
  });

  test('no event titles are kept in the stored status', async () => {
    await configure();
    await sync();
    const stored = JSON.stringify(await loadStatus());
    expect(stored).not.toContain('Planning');
    expect(stored).not.toContain('Sport');
  });

  test('identical quiet runs are one log entry with a repeat count', async () => {
    await configure();
    await sync();
    await sync();
    await sync();
    await sync();
    const { log } = await loadStatus();
    expect(log).toHaveLength(3);
    expect(log[0]?.repeats).toBe(2);
    expect(log[2]?.counts?.created).toBe(7);
  });

  test('previews and a switched-off mirror leave no trace in the log', async () => {
    await configure();
    await sync('preview');
    expect((await loadStatus()).log).toHaveLength(0);
    await configure(testSettings({ enabled: false }));
    dav.clearRequests();
    expect((await sync()).record.outcome).toBe('disabled');
    expect((await loadStatus()).log).toHaveLength(0);
    expect(dav.requests).toHaveLength(0);
  });

  test('an Outlook session the server rejects is forgotten', async () => {
    await configure();
    outlook.expectedToken = 'another-token';
    const { record } = await sync();
    expect(record.outcome).toBe('auth-failed');
    expect(await loadToken()).toBeNull();
    expect((await sync()).record.outcome).toBe('no-token');
  });

  test('the token is only ever kept in session storage', async () => {
    await configure();
    await sync();
    expect(JSON.stringify([...fake.local.values()])).not.toContain(TOKEN);
    expect(JSON.stringify([...fake.session.values()])).toContain(TOKEN);
  });
});

describe('quiet runs', () => {
  test('after one quiet run the calendar is only asked for its change tag', async () => {
    await configure();
    await sync();
    await sync();
    dav.clearRequests();
    const { record } = await sync();
    expect(record.outcome).toBe('applied');
    expect(record.message).toContain('Nothing changed');
    expect(dav.requests.map((r) => r.method)).toEqual(['PROPFIND']);
  });

  test('a changed destination drops back to a full read', async () => {
    await configure();
    await sync();
    await sync();
    dav.remove([...dav.objects.keys()][0] as string);
    dav.clearRequests();
    const { record } = await sync();
    expect(record.counts?.created).toBe(1);
    expect(dav.requests.map((r) => r.method)).toContain('REPORT');
  });
});

describe('holding back and confirming', () => {
  async function wipeSetup() {
    await configure();
    await sync();
    outlook.events = [];
    return (await sync()).record;
  }

  test('an empty Outlook answer is held back, recorded, and nothing is deleted', async () => {
    const held = await wipeSetup();
    expect(held.outcome).toBe('needs-confirmation');
    expect(held.guard?.code).toBe('empty-source');
    expect(held.digest).toBeString();
    expect(dav.objects.size).toBe(7);
    expect((await loadStatus()).lastRun?.outcome).toBe('needs-confirmation');
  });

  test('a preview offers the plan and a confirmation applies exactly that plan', async () => {
    await wipeSetup();
    const preview = await sync('preview');
    expect(preview.plan?.deletions).toHaveLength(7);
    expect(preview.record.guard).toBeDefined();
    const digest = preview.record.digest as string;

    const wrong = await sync('apply', { digest: 'not-the-plan' });
    expect(wrong.record.outcome).toBe('needs-confirmation');
    expect(dav.objects.size).toBe(7);

    const done = await sync('apply', { digest });
    expect(done.record.outcome).toBe('applied');
    expect(done.record.counts?.deleted).toBe(7);
    expect(dav.objects.size).toBe(0);
  });
});

describe('overlapping requests', () => {
  test('interactive requests report busy, scheduled ones coalesce into one follow-up', async () => {
    await configure();
    outlook.events = Array.from({ length: 30 }, (_, i) =>
      rawEvent({ id: outlookId(i), start: '2026-10-06T08:00:00', end: '2026-10-06T09:00:00' }),
    );
    const first = runSyncFlight({ trigger: 'alarm', mode: 'apply' });
    const preview = await runSyncFlight({ trigger: 'manual', mode: 'preview' });
    expect(preview.busy).toBe(true);
    const reset = await runResetFlight('count');
    expect(reset.busy).toBe(true);
    const second = runSyncFlight({ trigger: 'alarm', mode: 'apply' });
    const third = runSyncFlight({ trigger: 'manual', mode: 'apply' });
    const results = await Promise.all([first, second, third]);
    expect(results.every((r) => !r.busy)).toBe(true);
    const created = (r: (typeof results)[number]) => (r.busy ? -1 : r.value.record.counts?.created);
    expect(created(results[0] as (typeof results)[number])).toBe(30);
    expect(created(results[1] as (typeof results)[number])).toBe(0);
    expect(created(results[2] as (typeof results)[number])).toBe(0);
    expect(dav.writes).toHaveLength(30);
  });
});

describe('reset', () => {
  test('counts first, deletes only on the confirmed count, and records it', async () => {
    await configure();
    await sync();
    const counted = await runResetFlight('count');
    expect(counted.busy ? undefined : counted.value.own).toBe(7);
    const done = await runResetFlight('apply', 7);
    expect(done.busy ? undefined : done.value.outcome).toBe('reset');
    expect(dav.objects.size).toBe(0);
    const { log } = await loadStatus();
    expect(log[0]).toMatchObject({ kind: 'reset', outcome: 'reset' } satisfies Partial<RunRecord>);
  });
});
