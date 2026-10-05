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
  loadToken,
  saveSettings,
  saveToken,
} from '../src/extension/storage.ts';
import { FakeCalDav, routeFetch } from './helpers/fake-caldav.ts';
import { installFakeChrome } from './helpers/fake-chrome.ts';
import { FakeOutlook } from './helpers/fake-outlook.ts';
import { makeJwt, sampleWeek, testSettings } from './helpers/fixtures.ts';

const realFetch = globalThis.fetch;
const FAR_FUTURE = 4_000_000_000;
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
    account: { id: 't1/user-a', label: 'a@example.test' },
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
    expect(await loadAccount()).toEqual({ id: 't1/user-a', label: 'a@example.test' });
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
    await saveToken({ token: other, exp: FAR_FUTURE, account: { id: 't1/user-b', label: 'b' } });
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
