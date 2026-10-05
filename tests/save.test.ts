import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { saveSettingsFlow } from '../src/extension/save.ts';
import {
  loadBackoff,
  loadPasswordRejected,
  loadSettings,
  saveBackoff,
  savePasswordRejected,
  saveSettings,
  saveToken,
} from '../src/extension/storage.ts';
import { FakeCalDav, routeFetch } from './helpers/fake-caldav.ts';
import { installFakeChrome } from './helpers/fake-chrome.ts';
import { FakeOutlook } from './helpers/fake-outlook.ts';
import { sampleWeek, TOKEN, testSettings } from './helpers/fixtures.ts';

const realFetch = globalThis.fetch;
let fake: ReturnType<typeof installFakeChrome>;
let outlook: FakeOutlook;
let dav: FakeCalDav;

beforeEach(async () => {
  fake = installFakeChrome();
  outlook = new FakeOutlook(TOKEN, sampleWeek());
  dav = new FakeCalDav();
  globalThis.fetch = routeFetch({
    'outlook.cloud.microsoft': outlook.fetch,
    'dav.example.test': dav.fetch,
    'other.example.test': dav.fetch,
  }) as typeof fetch;
  await saveToken({ token: TOKEN, exp: 4_000_000_000 });
});

afterEach(() => {
  globalThis.fetch = realFetch;
});

describe('saving settings', () => {
  test('names the calendar, stores the settings and shows the first sync at once', async () => {
    const response = await saveSettingsFlow(testSettings());
    expect(response.saved).toBe(true);
    expect(response.connection).toBe("Connected to calendar 'Work'");
    const sync = response.sync;
    expect(sync && !sync.busy && sync.record?.outcome).toBe('applied');
    expect(dav.objects.size).toBe(7);
    expect((await loadSettings()).stored).toBe(true);
  });

  test('a held-back plan comes back with the save', async () => {
    dav.seed(
      'x.ics',
      'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:x\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n',
    );
    const response = await saveSettingsFlow(testSettings());
    const sync = response.sync;
    expect(sync && !sync.busy && sync.record?.outcome).toBe('needs-confirmation');
    expect(sync && !sync.busy && sync.plan?.creations).toHaveLength(7);
    expect(fake.state.badge).toBe('!');
  });

  test('refuses a URL that is not a calendar and stores nothing', async () => {
    dav.collection = { kind: 'home' };
    const response = await saveSettingsFlow(testSettings());
    expect(response.saved).toBe(false);
    expect(response.errors?.[0]).toContain('not a calendar collection');
    expect((await loadSettings()).stored).toBe(false);
    expect(dav.writes).toHaveLength(0);
  });

  test('refuses a password the server rejects', async () => {
    dav.password = 'rotated';
    const response = await saveSettingsFlow(testSettings());
    expect(response.saved).toBe(false);
    expect(response.errors?.[0]).toContain('rejected the credentials');
  });

  test('saves when the server cannot be reached, and says so', async () => {
    globalThis.fetch = (async () => {
      throw new TypeError('offline');
    }) as unknown as typeof fetch;
    const response = await saveSettingsFlow(testSettings());
    expect(response.saved).toBe(true);
    expect(response.connection).toContain('Could not reach');
  });

  test('refuses an Outlook host as the calendar server', async () => {
    const response = await saveSettingsFlow(
      testSettings({ caldavUrl: 'https://outlook.office.com/x/' }),
    );
    expect(response.saved).toBe(false);
    expect(outlook.requests).toHaveLength(0);
  });

  test('clears a rejected-password stop and a failure back-off', async () => {
    await saveSettings(testSettings());
    await savePasswordRejected(true);
    await saveBackoff({ signature: 's', count: 5, retryAt: 4_000_000_000_000 });
    await saveSettingsFlow(testSettings());
    expect(await loadPasswordRejected()).toBe(false);
    expect(await loadBackoff()).toBeUndefined();
  });
});

describe('the stored password and another server', () => {
  test('an empty password field keeps the password for the same server', async () => {
    await saveSettings(testSettings());
    const response = await saveSettingsFlow(testSettings({ password: '', username: 'alice' }));
    expect(response.saved).toBe(true);
    expect((await loadSettings()).settings.password).toBe('test-app-password');
  });

  test('is never sent to another origin: the save is refused without a request', async () => {
    await saveSettings(testSettings());
    dav.clearRequests();
    const response = await saveSettingsFlow(
      testSettings({
        password: '',
        caldavUrl: 'https://other.example.test/remote.php/dav/calendars/alice/work/',
      }),
    );
    expect(response.saved).toBe(false);
    expect(response.errors?.[0]).toContain('new calendar server');
    expect(dav.requests).toHaveLength(0);
    expect((await loadSettings()).settings.caldavUrl).toContain('dav.example.test');
  });

  test('moving to another server with its own password drops the old origin permission', async () => {
    await saveSettings(testSettings());
    const response = await saveSettingsFlow(
      testSettings({
        password: 'test-app-password',
        caldavUrl: 'https://other.example.test/remote.php/dav/calendars/alice/work/',
      }),
    );
    expect(response.saved).toBe(true);
    expect(fake.state.removedOrigins).toEqual(['https://dav.example.test/*']);
  });

  test('staying on the same server releases nothing', async () => {
    await saveSettings(testSettings());
    await saveSettingsFlow(testSettings({ daysBack: 10 }));
    expect(fake.state.removedOrigins).toEqual([]);
  });
});
