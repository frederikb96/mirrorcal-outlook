import { describe, expect, test } from 'bun:test';
import { MAX_PAGES, OUTLOOK_HOSTS } from '../src/core/constants.ts';
import {
  buildCalendarViewRequest,
  calendarViewUrl,
  fieldNeeds,
  OutlookReadError,
  readCalendarView,
} from '../src/core/outlook.ts';
import { FakeOutlook } from './helpers/fake-outlook.ts';
import { outlookId, rawEvent, sampleWeek, TOKEN, testSettings } from './helpers/fixtures.ts';

const start = new Date('2026-09-05T00:00:00Z');
const end = new Date('2027-04-05T00:00:00Z');
const NEEDS = { location: false, body: false, bodyPresence: false, attendees: false };
const GOOD = calendarViewUrl('outlook.cloud.microsoft', start, end, NEEDS);

function query(outlook: FakeOutlook, needs = NEEDS, token = TOKEN) {
  return readCalendarView({
    fetch: outlook.fetch,
    token,
    host: 'outlook.cloud.microsoft',
    start,
    end,
    needs,
  });
}

async function readError(outlook: FakeOutlook): Promise<OutlookReadError> {
  const error = await query(outlook).then(
    () => undefined,
    (e: unknown) => e,
  );
  expect(error).toBeInstanceOf(OutlookReadError);
  return error as OutlookReadError;
}

describe('request builder', () => {
  test('produces a credential-free GET with the token and UTC preference', () => {
    const { url, init } = buildCalendarViewRequest(GOOD, TOKEN);
    expect(init.method).toBe('GET');
    expect(init.credentials).toBe('omit');
    expect(init.redirect).toBe('error');
    expect(init.body).toBeUndefined();
    const headers = init.headers as Record<string, string>;
    expect(headers.Authorization).toBe(`Bearer ${TOKEN}`);
    expect(headers.Prefer).toContain('outlook.timezone="UTC"');
    expect(new URL(url).pathname).toBe('/api/v2.0/me/calendarview');
  });

  test('accepts a token that already carries the Bearer prefix', () => {
    const { init } = buildCalendarViewRequest(GOOD, `Bearer ${TOKEN}`);
    expect((init.headers as Record<string, string>).Authorization).toBe(`Bearer ${TOKEN}`);
  });

  test('refuses every other host, scheme, path and decoration', () => {
    const bad = [
      'https://evil.example.test/api/v2.0/me/calendarview',
      'https://outlook.office.com.evil.example.test/api/v2.0/me/calendarview',
      'http://outlook.office.com/api/v2.0/me/calendarview',
      'https://outlook.office.com:8443/api/v2.0/me/calendarview',
      'https://user:pw@outlook.office.com/api/v2.0/me/calendarview',
      'https://outlook.office.com/api/v2.0/me/events',
      'https://outlook.office.com/api/v2.0/me/sendmail',
      'https://outlook.office.com/api/v2.0/me/calendarview/../events',
      'https://outlook.office.com/api/v2.0/me/calendars/abc/events',
      'https://outlook.office.com/api/beta/me/calendarview',
      'https://outlook.office.com/owa/service.svc?action=GetCalendarView',
      'not a url',
    ];
    for (const url of bad) {
      expect(() => buildCalendarViewRequest(url, TOKEN)).toThrow(OutlookReadError);
    }
  });

  test('accepts each allow-listed host and the user-scoped calendar view path', () => {
    for (const host of OUTLOOK_HOSTS) {
      expect(() =>
        buildCalendarViewRequest(`https://${host}/api/v2.0/Me/CalendarView?$skip=3`, TOKEN),
      ).not.toThrow();
    }
    expect(() =>
      buildCalendarViewRequest(
        "https://outlook.office.com/api/v2.0/Users('abc-123')/CalendarView?$skip=3",
        TOKEN,
      ),
    ).not.toThrow();
  });

  test('refuses a token that is not a JWT', () => {
    expect(() => buildCalendarViewRequest(GOOD, 'abc\r\nX-Evil: 1')).toThrow(OutlookReadError);
    expect(() => buildCalendarViewRequest(GOOD, '')).toThrow(OutlookReadError);
  });
});

describe('reading the calendar view', () => {
  test('follows continuation links and returns every event', async () => {
    const outlook = new FakeOutlook(TOKEN, sampleWeek());
    const read = await query(outlook);
    expect(read.events).toHaveLength(7);
    expect(read.pages).toBe(3);
    expect(read.rawCount).toBe(7);
    expect(read.events[0]?.id).toBe(outlookId(1));
  });

  test('sends nothing but GETs to calendar-view URLs on allow-listed hosts', async () => {
    const outlook = new FakeOutlook(TOKEN, sampleWeek());
    await query(outlook);
    expect(outlook.requests.length).toBeGreaterThan(1);
    for (const r of outlook.requests) {
      const url = new URL(r.url);
      expect(r.method).toBe('GET');
      expect(url.protocol).toBe('https:');
      expect((OUTLOOK_HOSTS as readonly string[]).includes(url.hostname)).toBe(true);
      expect(url.pathname.toLowerCase()).toBe('/api/v2.0/me/calendarview');
    }
  });

  test('cancelled instances are skipped but counted as read', async () => {
    const outlook = new FakeOutlook(TOKEN, [
      rawEvent({ id: 'a', start: '2026-10-06T08:00:00', end: '2026-10-06T09:00:00' }),
      rawEvent({
        id: 'b',
        start: '2026-10-06T10:00:00',
        end: '2026-10-06T11:00:00',
        cancelled: true,
      }),
    ]);
    const read = await query(outlook);
    expect(read.events.map((e) => e.id)).toEqual(['a']);
    expect(read.rawCount).toBe(2);
  });

  test('normalises times and availability', async () => {
    const outlook = new FakeOutlook(TOKEN, [
      rawEvent({
        id: 'a',
        subject: null,
        start: '2026-10-06T08:00:00',
        end: '2026-10-06T09:00:00',
        showAs: 'Strange',
      }),
    ]);
    const [event] = (await query(outlook)).events;
    expect(event).toMatchObject({
      subject: '',
      start: '2026-10-06T08:00:00',
      end: '2026-10-06T09:00:00',
      showAs: 'Unknown',
    });
  });

  test('requests optional fields only when a policy needs them', async () => {
    const select = (needs: typeof NEEDS) =>
      new URL(calendarViewUrl('outlook.office.com', start, end, needs)).searchParams.get('$select');
    expect(select(NEEDS)).not.toMatch(/Body|Location|Attendees|Organizer/);
    expect(select({ ...NEEDS, body: true })).toContain('Body');
    expect(select({ ...NEEDS, bodyPresence: true })).toMatch(/BodyPreview/);
    expect(select({ ...NEEDS, bodyPresence: true })).not.toMatch(/Body,|Body$/);
    expect(select({ ...NEEDS, location: true })).toContain('Location');
    expect(select({ ...NEEDS, attendees: true })).toContain('Attendees');
    expect(fieldNeeds(testSettings())).toEqual(NEEDS);
    const s = testSettings();
    s.policies.attendees = 'copy';
    expect(fieldNeeds(s).attendees).toBe(true);
    s.policies.description = { mode: 'replace', text: 'x' };
    s.policies.location = { mode: 'replace', text: 'y' };
    expect(fieldNeeds(s)).toMatchObject({ body: false, bodyPresence: true, location: true });
  });

  test('parses people and text when requested', async () => {
    const outlook = new FakeOutlook(TOKEN, [
      rawEvent({
        id: 'a',
        start: '2026-10-06T08:00:00',
        end: '2026-10-06T09:00:00',
        location: 'Room 4',
        body: 'Agenda',
        organizer: { name: 'Pat', address: 'pat@example.test' },
        attendees: [
          { name: 'Sam', address: 'sam@example.test', type: 'Optional', response: 'Declined' },
        ],
      }),
    ]);
    const [event] = (
      await query(outlook, { location: true, body: true, bodyPresence: false, attendees: true })
    ).events;
    expect(event?.location).toBe('Room 4');
    expect(event?.body).toBe('Agenda');
    expect(event?.organizer?.address).toBe('pat@example.test');
    expect(event?.attendees).toEqual([
      { name: 'Sam', address: 'sam@example.test', type: 'Optional', response: 'Declined' },
    ]);
  });
});

describe('failed or partial reads', () => {
  const events = () => sampleWeek();

  test('HTTP errors abort', async () => {
    for (const status of [403, 404, 410, 429, 500, 503]) {
      const outlook = new FakeOutlook(TOKEN, events());
      outlook.fault = { kind: 'status', status };
      const error = await readError(outlook);
      expect(error.code).toBe('http');
      expect(error.message).toContain(String(status));
    }
  });

  test('401 is reported as an authentication problem', async () => {
    const error = await readError(new FakeOutlook('some-other-token', events()));
    expect(error.code).toBe('auth');
  });

  test('a failing later page discards everything read so far', async () => {
    const outlook = new FakeOutlook(TOKEN, events());
    outlook.fault = { kind: 'status', status: 500, onPage: 1 };
    const error = await readError(outlook);
    expect(error.code).toBe('http');
    expect(outlook.requests).toHaveLength(2);
  });

  test('network failure and timeouts abort', async () => {
    const outlook = new FakeOutlook(TOKEN, events());
    outlook.fault = { kind: 'network' };
    expect((await readError(outlook)).code).toBe('network');
  });

  test('bodies that are not an event list abort', async () => {
    for (const body of ['<html>login</html>', '{}', '{"value": 5}', 'null', '[]', '{"error":{}}']) {
      const outlook = new FakeOutlook(TOKEN, events());
      outlook.fault = { kind: 'body', body };
      expect((await readError(outlook)).code).toBe('malformed');
    }
  });

  test('an unusable item aborts the whole read', async () => {
    const broken: Record<string, unknown>[] = [
      {
        ...rawEvent({ id: 'x', start: '2026-10-06T08:00:00', end: '2026-10-06T09:00:00' }),
        Id: '',
      },
      {
        ...rawEvent({ id: 'x', start: '2026-10-06T08:00:00', end: '2026-10-06T09:00:00' }),
        Start: undefined,
      },
      {
        ...rawEvent({ id: 'x', start: '2026-10-06T08:00:00', end: '2026-10-06T09:00:00' }),
        IsAllDay: 'no',
      },
      rawEvent({ id: 'x', start: '2026-10-06T09:00:00', end: '2026-10-06T08:00:00' }),
      {
        ...rawEvent({ id: 'x', start: '2026-10-06T08:00:00', end: '2026-10-06T09:00:00' }),
        Start: { DateTime: '2026-10-06T08:00:00.0000000', TimeZone: 'Europe/Berlin' },
      },
      {
        ...rawEvent({ id: 'x', start: '2026-10-06T08:00:00', end: '2026-10-06T09:00:00' }),
        Start: { DateTime: 'soon', TimeZone: 'UTC' },
      },
    ];
    for (const item of broken) {
      const outlook = new FakeOutlook(TOKEN, [
        rawEvent({ id: 'fine', start: '2026-10-06T06:00:00', end: '2026-10-06T07:00:00' }),
        item,
      ]);
      expect((await readError(outlook)).code).toBe('malformed');
    }
  });

  test('a continuation link leaving the allow-list aborts before it is followed', async () => {
    const links = [
      'https://evil.example.test/api/v2.0/me/calendarview?$skip=3',
      'https://outlook.office.com/api/v2.0/me/events?$skip=3',
      'http://outlook.office.com/api/v2.0/me/calendarview?$skip=3',
    ];
    for (const link of links) {
      const outlook = new FakeOutlook(TOKEN, events());
      outlook.fault = { kind: 'next-link', link };
      const error = await readError(outlook);
      expect(error.code).toBe('request');
      expect(outlook.requests).toHaveLength(1);
    }
  });

  test('looping or endless paging aborts', async () => {
    const loop = new FakeOutlook(TOKEN, events());
    loop.fault = { kind: 'next-link', link: GOOD, onPage: 0 };
    expect((await readError(loop)).code).toBe('paging');
    expect(loop.requests).toHaveLength(1);

    const many = Array.from({ length: MAX_PAGES + 5 }, (_, i) =>
      rawEvent({ id: `e${i}`, start: '2026-10-06T08:00:00', end: '2026-10-06T09:00:00' }),
    );
    const endless = new FakeOutlook(TOKEN, many);
    endless.pageSize = 1;
    expect((await readError(endless)).code).toBe('paging');
  });

  test('error messages never contain the token', async () => {
    const faults = [
      { kind: 'status', status: 500 },
      { kind: 'network' },
      { kind: 'body', body: 'x' },
    ] as const;
    for (const fault of faults) {
      const outlook = new FakeOutlook(TOKEN, events());
      outlook.fault = fault;
      const error = await readError(outlook);
      expect(error.message).not.toContain(TOKEN);
      expect(error.message).not.toContain(TOKEN.split('.')[1] as string);
    }
  });
});
