/**
 * End-to-end run of the built extension in Chromium against a local fake Outlook and a local fake
 * CalDAV server. The browser resolves the Outlook and CalDAV host names to this machine only
 * (every other host fails to resolve), so no request can leave it.
 *
 * Run with `bun run test:e2e` after `bun run build`. Needs a Chromium that still honours
 * --load-extension (Playwright's bundled one); set CHROMIUM_PATH to override its location.
 */
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { Page } from 'playwright-core';
import type { StatusResponse } from '../../src/extension/messages.ts';
import {
  browserAvailable,
  DAV_HOST,
  type Harness,
  OUTLOOK_HOST,
  root,
  startHarness,
  waitFor,
} from '../helpers/browser-harness.ts';
import { extensionIdFromKey } from '../helpers/extension-id.ts';
import { FakeCalDav, routeCalendars } from '../helpers/fake-caldav.ts';
import { FakeOutlook } from '../helpers/fake-outlook.ts';
import {
  EXCLUDED_COLLECTION,
  makeJwt,
  outlookId,
  rawEvent,
  sampleWeek,
  testSettings,
} from '../helpers/fixtures.ts';

// A token that stays valid however long the suite runs; the fake server expects exactly this one.
const TOKEN = makeJwt({ exp: Math.floor(Date.now() / 1000) + 3600 });

const outlook = new FakeOutlook(TOKEN, sampleWeek());
const dav = new FakeCalDav();
const excludedDav = new FakeCalDav(EXCLUDED_COLLECTION);
excludedDav.collection = { kind: 'calendar', name: 'Other' };
const davFetch = routeCalendars(dav, excludedDav);
// The browser runs on the real clock and the fixture week is fixed: the window reaches back to it.
const settings = (overrides: Parameters<typeof testSettings>[0] = {}) =>
  testSettings({ daysBack: 730, ...overrides });

let harness: Harness;
let context: Harness['context'];
let extensionId: string;
let pageRequests: string[];

const optionsPage = () => harness.optionsPage();

async function message<T>(page: Page, request: unknown): Promise<T> {
  return page.evaluate((r) => chrome.runtime.sendMessage(r), request) as Promise<T>;
}

beforeAll(async () => {
  if (!browserAvailable) return;
  harness = await startHarness({ token: TOKEN, outlook, davFetch });
  ({ context, extensionId, pageRequests } = harness);
}, 60_000);

afterAll(async () => {
  await harness?.close();
});

describe.skipIf(!browserAvailable)('extension in Chromium', () => {
  test('observes the page token, then mirrors on its own and shows it in the settings page', async () => {
    const options = await optionsPage();
    await options.evaluate(
      (settings) => chrome.storage.local.set({ settings }),
      settings({ titleFilters: ['Sport', 'Mittagessen', 'Abwesend'] }),
    );

    const outlookTab = await context.newPage();
    await outlookTab.goto(`https://${OUTLOOK_HOST}/owa/`);

    await waitFor('the first sync to finish', () => dav.objects.size === 4 && dav.objects.size);
    const status = await waitFor('the status', async () => {
      const s = await message<StatusResponse>(options, { type: 'status' });
      return s.lastRun?.outcome === 'applied' ? s : undefined;
    });
    expect(status.token.usable).toBe(true);
    expect(status.lastRun?.trigger).toBe('token');
    expect(status.lastRun?.counts?.created).toBe(4);
    expect(outlook.requests.every((r) => r.method === 'GET')).toBe(true);
    expect(outlook.requests.some((r) => r.url.includes('/api/v2.0/me/calendarview'))).toBe(true);

    // The token is memory-only: not in storage.local.
    const local = JSON.stringify(await options.evaluate(() => chrome.storage.local.get(null)));
    expect(local).not.toContain(TOKEN);

    await options.reload();
    await options.locator('#status:has-text("Synced")').waitFor({ timeout: 10_000 });
  });

  test('a second sync writes nothing, a changed event writes exactly one object', async () => {
    const options = await optionsPage();
    dav.clearRequests();
    const again = await message<{ record: { outcome: string; counts: { unchanged: number } } }>(
      options,
      {
        type: 'sync',
        mode: 'apply',
      },
    );
    expect(again.record.outcome).toBe('applied');
    expect(again.record.counts.unchanged).toBe(4);
    expect(dav.writes).toHaveLength(0);

    outlook.events[0] = rawEvent({
      id: outlookId(1),
      subject: 'Planning, renamed',
      start: '2026-10-06T08:00:00',
      end: '2026-10-06T09:00:00',
    });
    dav.clearRequests();
    await message(options, { type: 'sync', mode: 'apply' });
    expect(dav.writes.map((w) => w.method)).toEqual(['PUT']);
  });

  test('the alarm is registered at the configured interval', async () => {
    const options = await optionsPage();
    const alarm = await waitFor('the alarm', () =>
      options.evaluate(() => chrome.alarms.get('mirrorcal-sync')),
    );
    expect(alarm.periodInMinutes).toBe(15);
  });

  test('an empty Outlook answer is held back and shown for confirmation', async () => {
    const options = await optionsPage();
    outlook.events = [];
    dav.clearRequests();
    await options.getByRole('button', { name: 'Sync now' }).click();
    await options.locator('#confirm-box').waitFor({ state: 'visible', timeout: 10_000 });
    expect(dav.writes).toHaveLength(0);
    expect(dav.objects.size).toBe(4);
  });

  test('confirming in the settings page applies the held-back plan', async () => {
    const options = await optionsPage();
    await options.getByRole('button', { name: 'Review changes', exact: true }).click();
    await options.locator('#confirm-box').waitFor({ state: 'visible', timeout: 10_000 });
    expect(dav.objects.size).toBe(4);
    await options.getByRole('button', { name: 'Confirm and apply' }).click();
    await waitFor('the deletions', () => dav.objects.size === 0);
    expect(dav.writes.every((w) => w.method === 'DELETE')).toBe(true);
  });

  test('saving checks the calendar, names it and shows the first result at once', async () => {
    const options = await optionsPage();
    await options.getByRole('button', { name: 'Save', exact: true }).click();
    await options
      .locator('#save-message:has-text("Connected to calendar")')
      .waitFor({ timeout: 10_000 });
    await options.locator('#save-message:has-text("Synced")').waitFor({ timeout: 10_000 });
    expect(dav.requests.some((r) => r.method === 'PROPFIND')).toBe(true);
  });

  test('a calendar for excluded events set in the page receives the filtered events', async () => {
    outlook.events = sampleWeek();
    const options = await optionsPage();
    await options.locator('#excludedCaldavUrl').fill(EXCLUDED_COLLECTION);
    await options
      .locator('#excluded-where:has-text("mirrored into the calendar for excluded events")')
      .waitFor({ timeout: 5_000 });
    await options.getByRole('button', { name: 'Save', exact: true }).click();
    await options
      .locator('#save-message:has-text("for excluded events")')
      .waitFor({ timeout: 10_000 });
    await waitFor('both calendars', () => dav.objects.size === 4 && excludedDav.objects.size === 3);
    expect(excludedDav.writes.every((w) => w.url.startsWith(EXCLUDED_COLLECTION))).toBe(true);
    await options
      .locator('#status:has-text("Calendar for excluded events")')
      .waitFor({ timeout: 10_000 });
    const screenshot = process.env.E2E_SCREENSHOT;
    if (screenshot) await options.screenshot({ path: screenshot, fullPage: true });
  });

  test('the extension id is the one the manifest key fixes', async () => {
    const manifest = JSON.parse(readFileSync(join(root, 'dist', 'manifest.json'), 'utf8'));
    expect(extensionId).toBe(await extensionIdFromKey(manifest.key));
  });

  test('only the two local hosts were ever contacted', () => {
    const hosts = new Set(pageRequests.map((u) => new URL(u).hostname));
    for (const host of hosts) {
      expect([OUTLOOK_HOST, DAV_HOST, extensionId]).toContain(host);
    }
  });
});
