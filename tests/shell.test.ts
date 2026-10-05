import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import type { Settings } from '../src/core/types.ts';
import { captureToken, resetCaptureMemory } from '../src/extension/capture.ts';
import { keepAliveWhile, runResetFlight, runSyncFlight } from '../src/extension/runner.ts';
import { buildStatus } from '../src/extension/status.ts';
import {
  clearAccount,
  loadAccount,
  loadBackoff,
  loadPasswordRejected,
  loadStatus,
  loadToken,
  saveSettings,
  saveToken,
} from '../src/extension/storage.ts';
import { FakeCalDav, routeFetch } from './helpers/fake-caldav.ts';
import { installFakeChrome } from './helpers/fake-chrome.ts';
import { FakeOutlook } from './helpers/fake-outlook.ts';
import { makeJwt, outlookId, rawEvent, sampleWeek, testSettings } from './helpers/fixtures.ts';

const realFetch = globalThis.fetch;
const FAR_FUTURE = 4_000_000_000;
const USER_A = { tenant: 't1', oid: 'user-a', name: 'a@example.test', label: 'a@example.test' };
const TOKEN = makeJwt({ exp: FAR_FUTURE, tid: 't1', oid: 'user-a', upn: 'a@example.test' });

let fake: ReturnType<typeof installFakeChrome>;
let outlook: FakeOutlook;
let dav: FakeCalDav;

beforeEach(() => {
  fake = installFakeChrome();
  resetCaptureMemory();
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
  await saveToken({
    token: TOKEN,
    exp: FAR_FUTURE,
    account: USER_A,
  });
}

async function sync(
  trigger: 'alarm' | 'manual' = 'alarm',
  mode: 'apply' | 'preview' = 'apply',
  confirm?: { digest: string },
) {
  const admitted = await runSyncFlight({ trigger, mode, ...(confirm ? { confirm } : {}) });
  if (admitted.busy) throw new Error('busy');
  return admitted.value;
}

const requestCount = () => outlook.requests.length + dav.requests.length;

describe('a rejected calendar password', () => {
  test('is reported once, stops scheduled runs and shows on the toolbar', async () => {
    await configure();
    dav.password = 'rotated';
    const first = await sync();
    expect(first.record.outcome).toBe('credentials-rejected');
    expect(first.record.message).toContain('enter it again');
    expect(dav.requests).toHaveLength(1);
    expect(await loadPasswordRejected()).toBe(true);
    expect(fake.state.badge).toBe('!');

    outlook.requests.length = 0;
    dav.clearRequests();
    for (const trigger of ['alarm', 'startup', 'token'] as const) {
      const run = await runSyncFlight({ trigger, mode: 'apply' });
      expect(run.busy ? '' : run.value.record.outcome).toBe('skipped');
    }
    expect(requestCount()).toBe(0);
    expect((await buildStatus()).passwordRejected).toBe(true);
  });

  test('saving the settings again lifts the stop', async () => {
    await configure();
    dav.password = 'rotated';
    await sync();
    dav.password = 'fresh';
    await saveSettings(testSettings({ password: 'fresh' }));
    expect(await loadPasswordRejected()).toBe(false);
    expect((await sync()).record.outcome).toBe('applied');
    expect(fake.state.badge).toBe('');
  });

  test('an explicit sync tries once more, and a working one clears the stop', async () => {
    await configure();
    dav.password = 'rotated';
    await sync();
    dav.clearRequests();
    expect((await sync('manual')).record.outcome).toBe('credentials-rejected');
    expect(dav.requests).toHaveLength(1);
    dav.password = 'test-app-password';
    expect((await sync('manual')).record.outcome).toBe('applied');
    expect(await loadPasswordRejected()).toBe(false);
  });
});

describe('the toolbar badge', () => {
  test('shows a held-back plan and clears once it is applied', async () => {
    await configure();
    await sync();
    expect(fake.state.badge).toBe('');
    outlook.events = [];
    const held = await sync();
    expect(held.record.outcome).toBe('needs-confirmation');
    expect(fake.state.badge).toBe('!');
    expect(fake.state.title).toContain('held back');
    const applied = await sync('manual', 'apply', { digest: held.record.digest as string });
    expect(applied.record.outcome).toBe('applied');
    expect(fake.state.badge).toBe('');
  });

  test('shows a missing permission at once', async () => {
    await configure();
    fake.state.permissionGranted = false;
    await sync();
    expect(fake.state.badge).toBe('!');
  });

  test('shows a lost Outlook session only after it persisted, and clears on success', async () => {
    await configure();
    await saveToken({ token: TOKEN, exp: 1_000 });
    await sync();
    await sync();
    expect(fake.state.badge).toBe('');
    await sync();
    expect(fake.state.badge).toBe('!');
    await saveToken({ token: TOKEN, exp: FAR_FUTURE });
    expect((await sync()).record.outcome).toBe('applied');
    expect(fake.state.badge).toBe('');
  });

  test('is cleared when syncing is switched off', async () => {
    await configure();
    fake.state.permissionGranted = false;
    await sync();
    fake.state.permissionGranted = true;
    await configure(testSettings({ enabled: false }));
    await sync();
    expect(fake.state.badge).toBe('');
  });

  test('a previewed plan does not touch it', async () => {
    await configure();
    fake.state.permissionGranted = false;
    await sync('manual', 'preview');
    expect(fake.state.badge).toBe('');
  });
});

describe('write failures', () => {
  test('are explained with their reason and labelled as partial', async () => {
    await configure();
    dav.failWhen = (method) => (method === 'PUT' ? 403 : undefined);
    const { record } = await sync();
    expect(record.outcome).toBe('applied');
    expect(record.counts?.errors).toBe(7);
    expect(record.message).toBe('Writing failed with HTTP 403');
  });

  test('an identical failure is not repeated blindly: scheduled runs back off', async () => {
    await configure();
    dav.failWhen = (method) => (method === 'PUT' ? 403 : undefined);
    for (let i = 0; i < 3; i++) await sync();
    expect((await loadBackoff())?.retryAt).toBeNumber();
    expect((await buildStatus()).retryAt).toBeString();
    dav.clearRequests();
    outlook.requests.length = 0;
    const skipped = await sync();
    expect(skipped.record.outcome).toBe('skipped');
    expect(requestCount()).toBe(0);
    const manual = await sync('manual');
    expect(manual.record.outcome).toBe('applied');
    expect(dav.writes.length).toBeGreaterThan(0);
  });

  test('saving the settings or a different result ends the back-off', async () => {
    await configure();
    dav.failWhen = (method) => (method === 'PUT' ? 403 : undefined);
    for (let i = 0; i < 3; i++) await sync();
    await saveSettings(testSettings());
    expect(await loadBackoff()).toBeUndefined();
    dav.failWhen = undefined;
    expect((await sync()).record.outcome).toBe('applied');
    expect(await loadBackoff()).toBeUndefined();
  });

  test('a repeating hold-back backs off too, until the plan changes', async () => {
    await configure();
    await sync();
    outlook.events = [];
    for (let i = 0; i < 3; i++) expect((await sync()).record.outcome).toBe('needs-confirmation');
    expect((await sync()).record.outcome).toBe('skipped');
  });
});

describe('the mirrored account', () => {
  test('is bound by the first successful sync and shown in the status', async () => {
    await configure();
    expect((await buildStatus()).account).toEqual({ label: 'a@example.test', pinned: false });
    expect(await loadAccount()).toBeUndefined();
    await sync();
    expect(await loadAccount()).toEqual(USER_A);
    expect((await buildStatus()).account).toEqual({ label: 'a@example.test', pinned: true });
  });

  test('a held-back first run does not bind it', async () => {
    await configure();
    dav.seed(
      'x.ics',
      'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:x\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n',
    );
    expect((await sync()).record.outcome).toBe('needs-confirmation');
    expect(await loadAccount()).toBeUndefined();
  });

  test('a session of anyone else is neither captured nor used afterwards', async () => {
    await configure();
    await sync();
    const other = makeJwt({ exp: FAR_FUTURE, tid: 't1', oid: 'user-b', upn: 'b@example.test' });
    const captured = await captureToken({
      initiator: 'https://outlook.cloud.microsoft',
      requestHeaders: [{ name: 'Authorization', value: `Bearer ${other}` }],
    });
    expect(captured.stored).toBe(false);
    expect((await loadToken())?.token).toBe(TOKEN);

    // Even a stored token of someone else is refused before any request goes out.
    await saveToken({
      token: other,
      exp: FAR_FUTURE,
      account: { tenant: 't1', oid: 'user-b', label: 'b' },
    });
    outlook.requests.length = 0;
    outlook.expectedToken = other;
    dav.clearRequests();
    const run = await sync('manual');
    expect(run.record.outcome).toBe('auth-failed');
    expect(requestCount()).toBe(0);
  });

  test('a reset or an explicit forget releases it', async () => {
    await configure();
    await sync();
    const done = await runResetFlight('apply', 7);
    expect(done.busy ? undefined : done.value.outcome).toBe('reset');
    expect(await loadAccount()).toBeUndefined();
    expect(await loadToken()).toBeNull();
    await configure();
    await sync();
    expect(await loadAccount()).toBeDefined();
    await clearAccount();
    expect(await loadAccount()).toBeUndefined();
  });
});

describe('keeping the worker alive', () => {
  test('an extension API is called repeatedly while a task runs, and not afterwards', async () => {
    await keepAliveWhile(() => Bun.sleep(70), 10);
    const during = fake.state.platformCalls;
    expect(during).toBeGreaterThanOrEqual(3);
    await Bun.sleep(40);
    expect(fake.state.platformCalls).toBe(during);
  });
});

describe('another owner id in the calendar', () => {
  const foreignOwner = (uid: string) =>
    `BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:${uid}\r\nX-MIRRORCAL-OWNER:laptop\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n`;

  test('a confirmation is remembered, so later creations are not held back again', async () => {
    await configure();
    dav.seed('l.ics', foreignOwner('l'));
    const held = await sync();
    expect(held.record.guard?.code).toBe('other-owner');
    const done = await sync('manual', 'apply', { digest: held.record.digest as string });
    expect(done.record.outcome).toBe('applied');
    outlook.events.push(
      rawEvent({ id: outlookId(9), start: '2026-10-10T08:00:00', end: '2026-10-10T09:00:00' }),
    );
    expect((await sync()).record.outcome).toBe('applied');
  });

  test('its events can be removed on their own, without releasing the account', async () => {
    await configure();
    await sync();
    dav.seed('l.ics', foreignOwner('l'));
    const counted = await runResetFlight('count', undefined, 'laptop');
    expect(counted.busy ? undefined : counted.value.own).toBe(1);
    const done = await runResetFlight('apply', 1, 'laptop');
    expect(done.busy ? undefined : done.value.deleted).toBe(1);
    expect(dav.objects.size).toBe(7);
    expect(await loadAccount()).toEqual(USER_A);
  });
});

describe('waiting for a session while bound', () => {
  test('says how many sign-ins of another account were ignored', async () => {
    await configure();
    await sync();
    for (const oid of ['b', 'c']) {
      const jwt = makeJwt({ exp: FAR_FUTURE, tid: 't1', oid, upn: `${oid}@example.test` });
      await captureToken({
        initiator: 'https://outlook.cloud.microsoft',
        requestHeaders: [{ name: 'Authorization', value: `Bearer ${jwt}` }],
      });
    }
    await saveToken({ token: TOKEN, exp: 1_000 });
    const { record } = await sync();
    expect(record.outcome).toBe('no-token');
    expect(record.message).toContain('2 sign-ins of another account ignored');
    expect((await buildStatus()).ignoredTokens).toBe(2);
  });
});

describe('failures that read differently each time', () => {
  test('still count as the same failure', async () => {
    await configure();
    let flip = false;
    globalThis.fetch = routeFetch({
      'outlook.cloud.microsoft': outlook.fetch,
      'dav.example.test': async (_url, init) => {
        if (init?.method !== 'REPORT') throw new TypeError('offline');
        flip = !flip;
        if (flip) throw new TypeError('offline');
        const timeout = new Error('slow');
        timeout.name = 'TimeoutError';
        throw timeout;
      },
    }) as typeof fetch;
    for (let i = 0; i < 3; i++) {
      expect((await sync()).record.outcome).toBe('destination-failed');
    }
    expect((await loadStatus()).log).toHaveLength(1);
    expect((await loadStatus()).log[0]?.repeats).toBe(3);
    expect(fake.state.badge).toBe('!');
    expect((await sync()).record.outcome).toBe('skipped');
  });
});
