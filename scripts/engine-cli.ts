#!/usr/bin/env bun
/**
 * Runs the sync engine outside the browser, with the real network.
 *
 *   OUTLOOK_TOKEN=<jwt or "Bearer <jwt>"> bun scripts/engine-cli.ts --settings settings.json [--mode preview|apply]
 *
 * Options
 *   --settings <file>   JSON with any Settings fields; missing ones take the defaults
 *   --mode <mode>       preview (default) plans only; apply writes
 *   --confirm           let a plan through that the safety guards would hold back
 *   --digest <hash>     with --confirm: only for the plan carrying this digest
 *   --ics               include the generated ICS in the plan entries
 *   --now <iso>         run as if it were this moment
 *   --reset-count       count the mirrored objects in the destination
 *   --reset <count>     delete every mirrored object; <count> must match the current count
 *
 * Environment: OUTLOOK_TOKEN (required except for reset), MIRRORCAL_CALDAV_URL,
 * MIRRORCAL_USER and MIRRORCAL_PASSWORD override the matching settings fields.
 * The same engine is importable from dist-engine/engine.mjs (`bun run build:engine`).
 */
import { readFileSync } from 'node:fs';
import { resetMirror, runSync } from '../src/core/engine.ts';
import { mergeWithDefaults } from '../src/core/settings.ts';

function option(name: string): string | undefined {
  const index = process.argv.indexOf(name);
  return index < 0 ? undefined : process.argv[index + 1];
}
const flag = (name: string) => process.argv.includes(name);

const file = option('--settings');
const stored: Record<string, unknown> = file ? JSON.parse(readFileSync(file, 'utf8')) : {};
const env = process.env;
const settings = mergeWithDefaults({
  ...stored,
  ...(env.MIRRORCAL_CALDAV_URL ? { caldavUrl: env.MIRRORCAL_CALDAV_URL } : {}),
  ...(env.MIRRORCAL_USER ? { username: env.MIRRORCAL_USER } : {}),
  ...(env.MIRRORCAL_PASSWORD ? { password: env.MIRRORCAL_PASSWORD } : {}),
});
const doFetch = (input: string, init?: RequestInit) => fetch(input, init);

if (flag('--reset-count') || option('--reset') !== undefined) {
  const confirmed = option('--reset');
  const result = await resetMirror({
    fetch: doFetch,
    settings,
    mode: confirmed === undefined ? 'count' : 'apply',
    ...(confirmed === undefined ? {} : { confirmedCount: Number(confirmed) }),
  });
  console.log(JSON.stringify(result, null, 2));
} else {
  const mode = option('--mode') ?? 'preview';
  if (mode !== 'preview' && mode !== 'apply') throw new Error('--mode must be preview or apply');
  const digest = option('--digest');
  const result = await runSync({
    fetch: doFetch,
    token: env.OUTLOOK_TOKEN ?? '',
    settings,
    now: new Date(option('--now') ?? Date.now()),
    mode,
    includeIcs: flag('--ics'),
    ...(flag('--confirm') ? { confirmMassChange: true } : {}),
    ...(digest ? { confirmedPlanDigest: digest } : {}),
  });
  console.log(JSON.stringify(result, null, 2));
}
