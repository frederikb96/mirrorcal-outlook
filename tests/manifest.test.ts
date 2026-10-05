import { describe, expect, test } from 'bun:test';
import pkg from '../package.json' with { type: 'json' };
import { OUTLOOK_HOSTS } from '../src/core/constants.ts';
import { buildManifest } from '../src/manifest.ts';

describe('manifest', () => {
  const manifest = buildManifest(pkg.version);

  test('is Manifest V3 with the version taken from package.json', () => {
    expect(manifest.manifest_version).toBe(3);
    expect(manifest.version).toBe(pkg.version);
  });

  test('asks only for what the sync needs', () => {
    expect([...manifest.permissions].sort()).toEqual(['alarms', 'storage', 'webRequest']);
    expect(manifest.host_permissions).toEqual(OUTLOOK_HOSTS.map((h) => `https://${h}/*`));
    expect(manifest.optional_host_permissions).toEqual(['https://*/*']);
  });

  test('declares a module service worker and an options page', () => {
    expect(manifest.background).toEqual({ service_worker: 'service-worker.js', type: 'module' });
    expect(manifest.options_ui.page).toBe('options.html');
  });
});
