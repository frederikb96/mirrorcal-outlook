/**
 * End-to-end run of the built extension in Chromium against a local fake Outlook and a local fake
 * CalDAV server. The browser resolves the Outlook and CalDAV host names to this machine only
 * (every other host fails to resolve), so no request can leave it.
 *
 * Run with `bun run test:e2e` after `bun run build`. Needs a Chromium that still honours
 * --load-extension (Playwright's bundled one); set CHROMIUM_PATH to override its location.
 */
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import {
  cpSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { type BrowserContext, chromium, type Page, type Worker } from 'playwright-core';
import type { StatusResponse } from '../../src/extension/messages.ts';
import { extensionIdFromKey } from '../helpers/extension-id.ts';
import { FakeCalDav } from '../helpers/fake-caldav.ts';
import { FakeOutlook } from '../helpers/fake-outlook.ts';
import {
  COLLECTION,
  makeJwt,
  outlookId,
  rawEvent,
  sampleWeek,
  testSettings,
} from '../helpers/fixtures.ts';

const root = resolve(import.meta.dir, '..', '..');
const executable = process.env.CHROMIUM_PATH ?? chromium.executablePath();
const OUTLOOK_HOST = 'outlook.cloud.microsoft';
const DAV_HOST = new URL(COLLECTION).hostname;

// A token that stays valid however long the suite runs; the fake server expects exactly this one.
const TOKEN = makeJwt({ exp: Math.floor(Date.now() / 1000) + 3600 });

const outlook = new FakeOutlook(TOKEN, sampleWeek());
const dav = new FakeCalDav();
const pageRequests: string[] = [];

let tmp: string;
let server: ReturnType<typeof Bun.serve>;
let context: BrowserContext;
let worker: Worker;
let extensionId: string;

const PAGE = `<!doctype html><title>Outlook</title><script>
  fetch('/api/beta/me/calendars', { headers: { Authorization: 'Bearer ${TOKEN}' } });
</script>`;

function selfSignedCert(dir: string) {
  const key = join(dir, 'key.pem');
  const cert = join(dir, 'cert.pem');
  const run = Bun.spawnSync([
    'openssl',
    'req',
    '-x509',
    '-newkey',
    'rsa:2048',
    '-nodes',
    '-keyout',
    key,
    '-out',
    cert,
    '-days',
    '1',
    '-subj',
    '/CN=mirrorcal-e2e',
    '-addext',
    `subjectAltName=DNS:${OUTLOOK_HOST},DNS:${DAV_HOST}`,
  ]);
  if (run.exitCode !== 0) throw new Error(`openssl failed: ${run.stderr.toString()}`);
  return { key: readFileSync(key, 'utf8'), cert: readFileSync(cert, 'utf8') };
}

async function waitFor<T>(
  what: string,
  probe: () => Promise<T | undefined | false> | T | undefined | false,
  ms = 20_000,
): Promise<T> {
  const deadline = Date.now() + ms;
  while (Date.now() < deadline) {
    const value = await probe();
    if (value) return value;
    await Bun.sleep(150);
  }
  throw new Error(`Timed out waiting for ${what}`);
}

async function optionsPage(): Promise<Page> {
  const page = await context.newPage();
  await page.goto(`chrome-extension://${extensionId}/options.html`);
  return page;
}

async function message<T>(page: Page, request: unknown): Promise<T> {
  return page.evaluate((r) => chrome.runtime.sendMessage(r), request) as Promise<T>;
}

const available = existsSync(executable);

beforeAll(async () => {
  if (!available) return;
  tmp = mkdtempSync(join(tmpdir(), 'mirrorcal-e2e-'));
  const build = Bun.spawnSync(['bun', 'run', 'scripts/build.ts'], { cwd: root });
  if (build.exitCode !== 0) throw new Error(`build failed: ${build.stderr.toString()}`);

  // The test copy may reach the fake CalDAV host without the runtime permission prompt.
  const extDir = join(tmp, 'extension');
  cpSync(join(root, 'dist'), extDir, { recursive: true });
  const manifestPath = join(extDir, 'manifest.json');
  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));
  manifest.host_permissions.push(`https://${DAV_HOST}/*`);
  writeFileSync(manifestPath, JSON.stringify(manifest));

  const tls = selfSignedCert(tmp);
  server = Bun.serve({
    port: 0,
    tls,
    async fetch(req) {
      const host = (req.headers.get('host') ?? '').split(':')[0];
      const url = new URL(req.url);
      const external = `https://${host}${url.pathname}${url.search}`;
      const headers = Object.fromEntries(req.headers);
      if (host === OUTLOOK_HOST) {
        if (url.pathname === '/owa/')
          return new Response(PAGE, { headers: { 'Content-Type': 'text/html' } });
        if (url.pathname.startsWith('/api/beta/')) return new Response('{}');
        return outlook.fetch(external, { method: req.method, headers });
      }
      if (host === DAV_HOST) {
        const withBody = ['PUT', 'REPORT', 'PROPFIND'].includes(req.method);
        const body = withBody ? await req.text() : undefined;
        return dav.fetch(external, {
          method: req.method,
          headers,
          ...(body === undefined ? {} : { body }),
        });
      }
      return new Response('unknown host', { status: 404 });
    },
  });

  const userData = join(tmp, 'profile');
  mkdirSync(userData);
  context = await chromium.launchPersistentContext(userData, {
    executablePath: executable,
    headless: true,
    ignoreHTTPSErrors: true,
    args: [
      `--disable-extensions-except=${extDir}`,
      `--load-extension=${extDir}`,
      `--host-resolver-rules=MAP ${OUTLOOK_HOST} 127.0.0.1:${server.port}, MAP ${DAV_HOST} 127.0.0.1:${server.port}, MAP * ~NOTFOUND, EXCLUDE 127.0.0.1`,
      '--ignore-certificate-errors',
      '--no-sandbox',
    ],
  });
  context.on('request', (r) => pageRequests.push(r.url()));
  worker = context.serviceWorkers()[0] ?? (await context.waitForEvent('serviceworker'));
  extensionId = new URL(worker.url()).host;
}, 60_000);

afterAll(async () => {
  await context?.close();
  server?.stop(true);
  if (tmp?.includes('mirrorcal-e2e-')) rmSync(tmp, { recursive: true, force: true });
});

describe.skipIf(!available)('extension in Chromium', () => {
  test('observes the page token, then mirrors on its own and shows it in the settings page', async () => {
    const options = await optionsPage();
    await options.evaluate(
      (settings) => chrome.storage.local.set({ settings }),
      testSettings({ titleFilters: ['Sport', 'Mittagessen', 'Abwesend'] }),
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
