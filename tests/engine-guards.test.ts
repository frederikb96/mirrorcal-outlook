import { beforeEach, describe, expect, test } from 'bun:test';
import { checkDestination, type RunOptions, resetMirror, runSync } from '../src/core/engine.ts';
import { unfoldLines } from '../src/core/ics.ts';
import type { Settings } from '../src/core/types.ts';
import { FakeCalDav, routeFetch } from './helpers/fake-caldav.ts';
import { FakeOutlook } from './helpers/fake-outlook.ts';
import {
  makeJwt,
  outlookId,
  rawEvent,
  sampleWeek,
  TOKEN,
  testSettings,
} from './helpers/fixtures.ts';

let outlook: FakeOutlook;
let dav: FakeCalDav;

beforeEach(() => {
  outlook = new FakeOutlook(TOKEN, sampleWeek());
  dav = new FakeCalDav();
});

function run(overrides: Partial<RunOptions> & { settings?: Settings } = {}) {
  return runSync({
    fetch: routeFetch({ 'outlook.cloud.microsoft': outlook.fetch, 'dav.example.test': dav.fetch }),
    token: TOKEN,
    settings: testSettings(),
    now: new Date('2026-10-05T12:00:00Z'),
    mode: 'apply',
    ...overrides,
  });
}

const FOREIGN =
  'BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:dentist\r\nSUMMARY:Dentist\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n';

describe('creations are guarded too', () => {
  test('changing the owner id holds the plan back instead of duplicating every event', async () => {
    await run();
    expect(dav.objects.size).toBe(7);
    dav.clearRequests();
    const result = await run({ settings: testSettings({ ownerId: 'laptop' }) });
    expect(result.outcome).toBe('needs-confirmation');
    expect(result.guard?.code).toBe('other-owner');
    expect(result.guard?.message).toContain('outlook');
    expect(dav.writes).toHaveLength(0);
    expect(dav.objects.size).toBe(7);
  });

  test('the held-back owner change can still be confirmed on purpose', async () => {
    await run();
    const held = await run({ settings: testSettings({ ownerId: 'laptop' }) });
    const result = await run({
      settings: testSettings({ ownerId: 'laptop' }),
      confirmMassChange: true,
      confirmedPlanDigest: held.digest as string,
    });
    expect(result.outcome).toBe('applied');
    expect(dav.objects.size).toBe(14);
  });

  test('a first sync into a calendar that only holds foreign events asks first', async () => {
    dav.seed('dentist.ics', FOREIGN);
    const result = await run();
    expect(result.outcome).toBe('needs-confirmation');
    expect(result.guard?.code).toBe('foreign-destination');
    expect(result.plan?.creations).toHaveLength(7);
    expect(dav.writes).toHaveLength(0);
  });

  test('once the mirror owns events, foreign objects beside them are fine', async () => {
    dav.seed('dentist.ics', FOREIGN);
    const held = await run();
    const first = await run({
      confirmMassChange: true,
      confirmedPlanDigest: held.digest as string,
    });
    expect(first.outcome).toBe('applied');
    outlook.events.push(
      rawEvent({ id: outlookId(9), start: '2026-10-10T08:00:00', end: '2026-10-10T09:00:00' }),
    );
    const next = await run();
    expect(next.outcome).toBe('applied');
    expect(next.counts.created).toBe(1);
    expect(dav.objects.has('dentist.ics')).toBe(true);
  });

  test('an empty destination and a calendar of its own events are not held back', async () => {
    expect((await run()).outcome).toBe('applied');
    expect((await run()).outcome).toBe('applied');
  });
});

describe('a rejected calendar password', () => {
  test('stops the run at the first request when applying', async () => {
    dav.password = 'rotated';
    const result = await run();
    expect(result.outcome).toBe('destination-failed');
    expect(result.error).toMatchObject({ source: 'caldav', code: 'auth' });
    expect(dav.requests).toHaveLength(1);
  });

  test('stops at the first request when previewing', async () => {
    dav.password = 'rotated';
    const result = await run({ mode: 'preview' });
    expect(result.error).toMatchObject({ source: 'caldav', code: 'auth' });
    expect(dav.requests).toHaveLength(1);
  });
});

describe('write failures', () => {
  test('carry the reason, never a title, and leave the destination alone', async () => {
    dav.failWhen = (method) => (method === 'PUT' ? 403 : undefined);
    const result = await run();
    expect(result.outcome).toBe('applied');
    expect(result.counts.errors).toBe(7);
    expect(result.message).toBe('Writing failed with HTTP 403');
    expect(result.state).toBeUndefined();
  });
});

describe('account binding', () => {
  const BOUND = {
    tenant: 'tenant-1',
    oid: 'user-a',
    name: 'user-a@example.test',
    label: 'user-a@example.test',
  };
  const tokenFor = (oid: string, tid = 'tenant-1') =>
    makeJwt({ oid, tid, upn: `${oid}@example.test` });

  test('the account of the token is reported once Outlook was read', async () => {
    const token = tokenFor('user-a');
    outlook.expectedToken = token;
    const result = await run({ token });
    expect(result.account).toEqual({
      tenant: 'tenant-1',
      oid: 'user-a',
      name: 'user-a@example.test',
      label: 'user-a@example.test',
    });
  });

  test('a token of anyone else is refused before anything is read', async () => {
    for (const token of [tokenFor('user-b'), tokenFor('user-a', 'tenant-2')]) {
      outlook.expectedToken = token;
      outlook.requests.length = 0;
      const result = await run({ token, expectedAccount: BOUND });
      expect(result.outcome).toBe('auth-failed');
      expect(result.error?.code).toBe('wrong-account');
      expect(outlook.requests).toHaveLength(0);
      expect(dav.requests).toHaveLength(0);
    }
  });

  test('a token that carries nothing comparable is let through', async () => {
    const result = await run({ token: TOKEN, expectedAccount: BOUND });
    expect(result.outcome).toBe('applied');
  });

  test('the bound account passes', async () => {
    const token = tokenFor('user-a');
    outlook.expectedToken = token;
    const result = await run({ token, expectedAccount: BOUND });
    expect(result.outcome).toBe('applied');
  });
});

describe('private events', () => {
  const privateWeek = () => [
    rawEvent({
      id: outlookId(1),
      subject: 'Therapy',
      start: '2026-10-06T08:00:00',
      end: '2026-10-06T09:00:00',
      sensitivity: 'Private',
      location: 'Clinic',
      body: 'Details',
    }),
    rawEvent({
      id: outlookId(2),
      subject: 'Board',
      start: '2026-10-07T08:00:00',
      end: '2026-10-07T09:00:00',
      sensitivity: 'Confidential',
    }),
    rawEvent({
      id: outlookId(3),
      subject: 'Standup',
      start: '2026-10-08T08:00:00',
      end: '2026-10-08T09:00:00',
      sensitivity: 'Normal',
    }),
  ];

  const written = () => [...dav.objects.values()].map((o) => unfoldLines(o.ics).join('\n')).sort();

  test('the sensitivity is requested from Outlook', async () => {
    await run();
    const select = new URL(outlook.requests[0]?.url as string).searchParams.get('$select');
    expect(select?.split(',')).toContain('Sensitivity');
  });

  test('by default they are flagged and carry a neutral title and nothing else', async () => {
    outlook.events = privateWeek();
    const settings = testSettings();
    settings.policies.description = { mode: 'copy' };
    settings.policies.location = { mode: 'copy' };
    await run({ settings });
    const objects = written();
    expect(objects).toHaveLength(3);
    const therapy = objects.find((o) => o.includes('CLASS:PRIVATE')) as string;
    expect(therapy).toContain('SUMMARY:Private');
    expect(therapy).not.toContain('Therapy');
    expect(therapy).not.toContain('Clinic');
    expect(therapy).not.toContain('Details');
    const board = objects.find((o) => o.includes('CLASS:CONFIDENTIAL')) as string;
    expect(board).toContain('SUMMARY:Private');
    const standup = objects.find((o) => o.includes('Standup')) as string;
    expect(standup).not.toContain('CLASS:');
  });

  test('copy keeps the real title but still flags the event', async () => {
    outlook.events = privateWeek();
    const settings = testSettings();
    settings.policies.privateEvents = { mode: 'copy' };
    await run({ settings });
    const therapy = written().find((o) => o.includes('Therapy')) as string;
    expect(therapy).toContain('CLASS:PRIVATE');
  });

  test('replace uses the configured text', async () => {
    outlook.events = privateWeek();
    const settings = testSettings();
    settings.policies.privateEvents = { mode: 'replace', text: 'Out' };
    await run({ settings });
    expect(written().filter((o) => o.includes('SUMMARY:Out'))).toHaveLength(2);
  });

  test('skip leaves them out of the destination', async () => {
    outlook.events = privateWeek();
    const settings = testSettings();
    settings.policies.privateEvents = { mode: 'skip' };
    const result = await run({ settings });
    expect(result.counts).toMatchObject({ filtered: 2, created: 1 });
    expect(written()).toHaveLength(1);
    expect(written()[0]).toContain('Standup');
  });

  test('a label this code does not know is treated as private', async () => {
    outlook.events = [
      rawEvent({
        id: outlookId(1),
        subject: 'Mystery',
        start: '2026-10-06T08:00:00',
        end: '2026-10-06T09:00:00',
        sensitivity: 'TopSecret',
      }),
    ];
    await run();
    expect(written()[0]).toContain('CLASS:PRIVATE');
    expect(written()[0]).not.toContain('Mystery');
  });
});

describe('checking the destination', () => {
  const check = () => checkDestination({ fetch: dav.fetch, settings: testSettings() });

  test('names the calendar it found', async () => {
    const result = await check();
    expect(result).toMatchObject({ outcome: 'ok', name: 'Work' });
    expect(result.message).toBe("Connected to calendar 'Work'");
    expect(dav.requests.map((r) => r.method)).toEqual(['PROPFIND']);
    expect(dav.requests[0]?.headers.depth).toBe('0');
  });

  test('refuses a URL that is not a calendar collection', async () => {
    dav.collection = { kind: 'home' };
    expect((await check()).outcome).toBe('not-a-calendar');
  });

  test('refuses a calendar that holds no events', async () => {
    dav.collection = { kind: 'tasks', name: 'Todo' };
    const result = await check();
    expect(result.outcome).toBe('not-a-calendar');
    expect(result.message).toContain('does not hold events');
  });

  test('reports a rejected password and an unreachable server', async () => {
    dav.password = 'rotated';
    expect((await check()).outcome).toBe('auth-failed');
    const down = await checkDestination({
      fetch: async () => {
        throw new TypeError('offline');
      },
      settings: testSettings(),
    });
    expect(down.outcome).toBe('unreachable');
  });

  test('refuses settings that cannot work without asking the server', async () => {
    const result = await checkDestination({
      fetch: dav.fetch,
      settings: testSettings({ caldavUrl: 'https://outlook.office.com/x/' }),
    });
    expect(result.outcome).toBe('invalid-settings');
    expect(dav.requests).toHaveLength(0);
  });
});

describe('another owner id', () => {
  const marked = (uid: string, owner: string | undefined) =>
    `BEGIN:VCALENDAR\r\nBEGIN:VEVENT\r\nUID:${uid}\r\nSUMMARY:${uid}\r\n${
      owner ? `X-MIRRORCAL-OWNER:${owner}\r\n` : ''
    }END:VEVENT\r\nEND:VCALENDAR\r\n`;
  const reset = (over: Partial<Parameters<typeof resetMirror>[0]> = {}) =>
    resetMirror({ fetch: dav.fetch, settings: testSettings(), mode: 'count', ...over });

  beforeEach(() => {
    dav.seed('l1.ics', marked('l1', 'laptop'));
    dav.seed('l2.ics', marked('l2', 'laptop'));
    dav.seed('p1.ics', marked('p1', 'phone'));
    dav.seed('plain.ics', marked('plain', undefined));
  });

  test('is listed with counts when counting', async () => {
    const result = await reset();
    expect(result.otherOwners).toEqual([
      { id: 'laptop', count: 2 },
      { id: 'phone', count: 1 },
    ]);
    expect(result.own).toBe(0);
  });

  test('its events are counted first and then deleted, and nothing else', async () => {
    const counted = await reset({ ownerId: 'laptop' });
    expect(counted.own).toBe(2);
    expect(dav.writes).toHaveLength(0);
    const done = await reset({ mode: 'apply', ownerId: 'laptop', confirmedCount: 2 });
    expect(done).toMatchObject({ outcome: 'reset', deleted: 2 });
    expect([...dav.objects.keys()].sort()).toEqual(['p1.ics', 'plain.ics']);
  });

  test('an object without any owner marker is never deleted, whatever id is named', async () => {
    for (const ownerId of ['laptop', 'phone', 'plain', 'x']) {
      const counted = await reset({ ownerId });
      await reset({ mode: 'apply', ownerId, confirmedCount: counted.own });
    }
    expect([...dav.objects.keys()]).toEqual(['plain.ics']);
  });

  test('a changed count refuses the deletion, and an invalid id is refused', async () => {
    expect((await reset({ mode: 'apply', ownerId: 'laptop', confirmedCount: 1 })).outcome).toBe(
      'refused',
    );
    expect((await reset({ ownerId: 'bad id!' })).outcome).toBe('invalid-settings');
    expect(dav.objects.size).toBe(4);
  });

  test('confirming it once is remembered for that set of owners', async () => {
    dav.remove('l2.ics');
    dav.remove('p1.ics');
    dav.remove('plain.ics');
    const held = await run();
    expect(held.guard?.code).toBe('other-owner');
    const confirmed = await run({
      confirmMassChange: true,
      confirmedPlanDigest: held.digest as string,
    });
    expect(confirmed.acknowledgedOwners).toEqual(['laptop']);
    outlook.events.push(
      rawEvent({ id: outlookId(9), start: '2026-10-10T08:00:00', end: '2026-10-10T09:00:00' }),
    );
    expect((await run()).outcome).toBe('needs-confirmation');
    const next = await run({ acknowledgedOwners: ['laptop'] });
    expect(next.outcome).toBe('applied');
    expect(next.counts.created).toBe(1);
  });

  test('a new owner among acknowledged ones holds back again', async () => {
    dav.remove('plain.ics');
    const result = await run({ acknowledgedOwners: ['laptop'] });
    expect(result.guard?.code).toBe('other-owner');
    expect(result.guard?.message).toContain('phone');
  });
});

describe('failure kinds', () => {
  test('write failures carry a code that does not depend on the wording', async () => {
    dav.failWhen = (method) => (method === 'PUT' ? 403 : undefined);
    const result = await run();
    expect(result.errors.every((e) => e.code === 'http:403')).toBe(true);
  });

  test('a destination failure carries its status', async () => {
    dav.failWhen = () => 503;
    const result = await run();
    expect(result.error).toMatchObject({ source: 'caldav', code: 'http', status: 503 });
  });
});
