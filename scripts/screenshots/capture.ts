#!/usr/bin/env bun
/**
 * Regenerates the pictures in docs/images/ of the options page: the built extension runs in
 * Playwright's Chromium against local fakes of Outlook and the calendar server, filled with
 * invented events dated relative to now. Run `bun run screenshots`.
 */
import { join } from 'node:path';
import {
  browserAvailable,
  OUTLOOK_HOST,
  root,
  startHarness,
  waitFor,
} from '../../tests/helpers/browser-harness.ts';
import { FakeCalDav, routeCalendars } from '../../tests/helpers/fake-caldav.ts';
import { FakeOutlook } from '../../tests/helpers/fake-outlook.ts';
import {
  EXCLUDED_COLLECTION,
  makeJwt,
  outlookId,
  rawEvent,
  testSettings,
} from '../../tests/helpers/fixtures.ts';

if (!browserAvailable)
  throw new Error('Playwright Chromium is missing: bunx playwright install chromium');

const out = join(root, 'docs', 'images');
const TOKEN = makeJwt({
  exp: Math.floor(Date.now() / 1000) + 3600,
  tid: '3f2a9c7e-5d41-4b8a-9e60-2c1d7a4b8f10',
  upn: 'alex.rivera@example.com',
});

function at(dayOffset: number, hour: number, minutes = 0): string {
  const d = new Date();
  d.setUTCDate(d.getUTCDate() + dayOffset);
  d.setUTCHours(hour, minutes, 0, 0);
  return d.toISOString().slice(0, 19);
}

const events = [
  rawEvent({
    id: outlookId(1),
    subject: 'Sprint planning',
    start: at(1, 8),
    end: at(1, 9),
    location: 'Room 4',
  }),
  rawEvent({ id: outlookId(2), subject: 'Design review', start: at(1, 13), end: at(1, 14) }),
  rawEvent({ id: outlookId(3), subject: 'Lunch', start: at(1, 11, 30), end: at(1, 12, 30) }),
  rawEvent({
    id: outlookId(4),
    subject: 'Focus time',
    start: at(2, 8),
    end: at(2, 11),
    showAs: 'Free',
  }),
  rawEvent({ id: outlookId(5), subject: 'Customer call', start: at(2, 14), end: at(2, 15) }),
  rawEvent({
    id: outlookId(6),
    subject: 'Team offsite',
    start: at(4, 0),
    end: at(6, 0),
    allDay: true,
  }),
  rawEvent({
    id: outlookId(7),
    subject: 'Doctor appointment',
    start: at(3, 9),
    end: at(3, 10),
    sensitivity: 'Private',
  }),
];

const outlook = new FakeOutlook(TOKEN, events);
const dav = new FakeCalDav();
dav.collection = { kind: 'calendar', name: 'Work mirror' };
const excludedDav = new FakeCalDav(EXCLUDED_COLLECTION);
excludedDav.collection = { kind: 'calendar', name: 'Filtered events' };

const harness = await startHarness({
  token: TOKEN,
  outlook,
  davFetch: routeCalendars(dav, excludedDav),
  viewport: { width: 800, height: 900 },
});

try {
  const options = await harness.optionsPage();
  await options.evaluate(
    (settings) => chrome.storage.local.set({ settings }),
    testSettings({
      excludedCaldavUrl: EXCLUDED_COLLECTION,
      titleFilters: ['Lunch', 'Focus time'],
    }),
  );
  const outlookTab = await harness.context.newPage();
  await outlookTab.goto(`https://${OUTLOOK_HOST}/owa/`);
  await waitFor('both calendars', () => dav.objects.size === 5 && excludedDav.objects.size === 2);
  await outlookTab.close();

  await options.reload();
  await options.locator('#status:has-text("Synced")').waitFor({ timeout: 10_000 });
  // A hairline keeps the white page visible against GitHub's white background.
  await options.addStyleTag({ content: 'html { box-shadow: inset 0 0 0 1px #d4d4d4; }' });

  await options.screenshot({ path: join(out, 'options-status.png') });

  await options.evaluate(() => {
    const legend = [...document.querySelectorAll('legend')].find(
      (l) => l.textContent === 'What is mirrored',
    );
    legend?.closest('fieldset')?.scrollIntoView({ block: 'start' });
    window.scrollBy(0, -16);
  });
  await options.screenshot({ path: join(out, 'options-settings.png') });
  console.log(`Wrote ${out}/options-status.png and options-settings.png`);
} finally {
  await harness.close();
}
