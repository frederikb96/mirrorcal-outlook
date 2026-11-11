/**
 * Chromium with the built extension loaded, against a local fake Outlook and fake CalDAV server.
 * The browser resolves the Outlook and CalDAV host names to this machine only (every other host
 * fails to resolve), so no request can leave it. Needs a Chromium that still honours
 * --load-extension (Playwright's bundled one); set CHROMIUM_PATH to override its location.
 */
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
import { COLLECTION } from './fixtures.ts';

export const root = resolve(import.meta.dir, '..', '..');
export const OUTLOOK_HOST = 'outlook.cloud.microsoft';
export const DAV_HOST = new URL(COLLECTION).hostname;
export const chromiumPath = process.env.CHROMIUM_PATH ?? chromium.executablePath();
export const browserAvailable = existsSync(chromiumPath);

/** What the harness needs from the fake Outlook. */
interface OutlookFake {
  fetch(url: string, init: { method: string; headers: Record<string, string> }): Promise<Response>;
}

export interface Harness {
  context: BrowserContext;
  worker: Worker;
  extensionId: string;
  /** URLs of every request the browser made. */
  pageRequests: string[];
  optionsPage(): Promise<Page>;
  close(): Promise<void>;
}

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

/**
 * Builds the extension and starts the browser. `token` is the bearer token the fake Outlook page
 * sends; `davFetch` answers every request to the CalDAV host.
 */
export async function startHarness(options: {
  token: string;
  outlook: OutlookFake;
  davFetch: (url: string, init: RequestInit) => Promise<Response>;
  /** Passed to Chromium as the viewport of every page; the window default when absent. */
  viewport?: { width: number; height: number };
}): Promise<Harness> {
  const tmp = mkdtempSync(join(tmpdir(), 'mirrorcal-e2e-'));
  const build = Bun.spawnSync(['bun', 'run', 'scripts/build.ts'], { cwd: root });
  if (build.exitCode !== 0) throw new Error(`build failed: ${build.stderr.toString()}`);

  // The test copy may reach the fake CalDAV host without the runtime permission prompt.
  const extDir = join(tmp, 'extension');
  cpSync(join(root, 'dist'), extDir, { recursive: true });
  const manifestPath = join(extDir, 'manifest.json');
  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));
  manifest.host_permissions.push(`https://${DAV_HOST}/*`);
  writeFileSync(manifestPath, JSON.stringify(manifest));

  const page = `<!doctype html><title>Outlook</title><script>
  fetch('/api/beta/me/calendars', { headers: { Authorization: 'Bearer ${options.token}' } });
</script>`;
  const server = Bun.serve({
    port: 0,
    tls: selfSignedCert(tmp),
    async fetch(req) {
      const host = (req.headers.get('host') ?? '').split(':')[0];
      const url = new URL(req.url);
      const external = `https://${host}${url.pathname}${url.search}`;
      const headers = Object.fromEntries(req.headers);
      if (host === OUTLOOK_HOST) {
        if (url.pathname === '/owa/')
          return new Response(page, { headers: { 'Content-Type': 'text/html' } });
        if (url.pathname.startsWith('/api/beta/')) return new Response('{}');
        return options.outlook.fetch(external, { method: req.method, headers });
      }
      if (host === DAV_HOST) {
        const withBody = ['PUT', 'REPORT', 'PROPFIND'].includes(req.method);
        const body = withBody ? await req.text() : undefined;
        return options.davFetch(external, {
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
  const context = await chromium.launchPersistentContext(userData, {
    executablePath: chromiumPath,
    headless: true,
    ignoreHTTPSErrors: true,
    ...(options.viewport ? { viewport: options.viewport } : {}),
    args: [
      `--disable-extensions-except=${extDir}`,
      `--load-extension=${extDir}`,
      `--host-resolver-rules=MAP ${OUTLOOK_HOST} 127.0.0.1:${server.port}, MAP ${DAV_HOST} 127.0.0.1:${server.port}, MAP * ~NOTFOUND, EXCLUDE 127.0.0.1`,
      '--ignore-certificate-errors',
      '--no-sandbox',
    ],
  });
  const pageRequests: string[] = [];
  context.on('request', (r) => pageRequests.push(r.url()));
  const worker = context.serviceWorkers()[0] ?? (await context.waitForEvent('serviceworker'));
  const extensionId = new URL(worker.url()).host;

  return {
    context,
    worker,
    extensionId,
    pageRequests,
    async optionsPage() {
      const options = await context.newPage();
      await options.goto(`chrome-extension://${extensionId}/options.html`);
      return options;
    },
    async close() {
      await context.close();
      server.stop(true);
      if (tmp.includes('mirrorcal-e2e-')) rmSync(tmp, { recursive: true, force: true });
    },
  };
}

export async function waitFor<T>(
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
