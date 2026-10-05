import { DEFAULT_SETTINGS } from '../../src/core/settings.ts';
import type { Settings } from '../../src/core/types.ts';

export const NOW = new Date('2026-10-05T12:00:00Z');
export const COLLECTION = 'https://dav.example.test/remote.php/dav/calendars/alice/work/';

export function testSettings(overrides: Partial<Settings> = {}): Settings {
  return {
    ...structuredClone(DEFAULT_SETTINGS),
    caldavUrl: COLLECTION,
    username: 'alice',
    password: 'test-app-password',
    ...overrides,
  };
}

function base64url(value: string): string {
  return btoa(value).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

export function makeJwt(claims: Record<string, unknown> = {}): string {
  const header = base64url(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const payload = base64url(
    JSON.stringify({ aud: 'https://outlook.office.com', exp: 1_900_000_000, ...claims }),
  );
  return `${header}.${payload}.${base64url('signature')}`;
}

export const TOKEN = makeJwt();

export interface RawEventInit {
  id: string;
  subject?: string | null;
  /** `YYYY-MM-DDTHH:MM:SS` (UTC) */
  start: string;
  end: string;
  allDay?: boolean;
  showAs?: string;
  cancelled?: boolean;
  sensitivity?: string;
  location?: string;
  body?: string;
  organizer?: { name: string; address: string };
  attendees?: { name: string; address: string; type?: string; response?: string }[];
}

/** An event in the shape Outlook's calendar view returns it with a UTC time zone preference. */
export function rawEvent(init: RawEventInit): Record<string, unknown> {
  const event: Record<string, unknown> = {
    Id: init.id,
    Subject: init.subject === undefined ? `Meeting ${init.id}` : init.subject,
    Start: { DateTime: `${init.start}.0000000`, TimeZone: 'UTC' },
    End: { DateTime: `${init.end}.0000000`, TimeZone: 'UTC' },
    IsAllDay: init.allDay ?? false,
    IsCancelled: init.cancelled ?? false,
    ShowAs: init.showAs ?? 'Busy',
  };
  if (init.sensitivity !== undefined) event.Sensitivity = init.sensitivity;
  if (init.location !== undefined) event.Location = { DisplayName: init.location };
  if (init.body !== undefined) event.Body = { ContentType: 'Text', Content: init.body };
  if (init.organizer) {
    event.Organizer = {
      EmailAddress: { Name: init.organizer.name, Address: init.organizer.address },
    };
  }
  if (init.attendees) {
    event.Attendees = init.attendees.map((a) => ({
      Type: a.type ?? 'Required',
      Status: { Response: a.response ?? 'None' },
      EmailAddress: { Name: a.name, Address: a.address },
    }));
  }
  return event;
}

/** A long, unique id with the characters real Outlook ids contain. */
export function outlookId(n: number): string {
  return `AAMkAGI2${n}+/${'x'.repeat(60)}${String(n).padStart(4, '0')}=`;
}

/** A week of mixed events: timed, all-day, multi-day, tentative, free, with unicode. */
export function sampleWeek(): Record<string, unknown>[] {
  return [
    rawEvent({
      id: outlookId(1),
      subject: 'Planning',
      start: '2026-10-06T08:00:00',
      end: '2026-10-06T09:00:00',
      location: 'Room 4',
      body: 'Agenda, items; more',
      organizer: { name: 'Pat Doe', address: 'pat@example.test' },
      attendees: [{ name: 'Sam Roe', address: 'sam@example.test', response: 'Accepted' }],
    }),
    rawEvent({
      id: outlookId(2),
      subject: 'Conference day',
      start: '2026-10-08T00:00:00',
      end: '2026-10-09T00:00:00',
      allDay: true,
    }),
    rawEvent({
      id: outlookId(3),
      subject: 'Offsite',
      start: '2026-10-12T00:00:00',
      end: '2026-10-15T00:00:00',
      allDay: true,
      showAs: 'Free',
    }),
    rawEvent({
      id: outlookId(4),
      subject: 'Sport',
      start: '2026-10-07T11:30:00',
      end: '2026-10-07T12:30:00',
      showAs: 'Free',
    }),
    rawEvent({
      id: outlookId(5),
      subject: 'Mittagessen',
      start: '2026-10-07T12:00:00',
      end: '2026-10-07T13:00:00',
    }),
    rawEvent({
      id: outlookId(6),
      subject: 'Café — Überprüfung 日本',
      start: '2026-10-09T14:00:00',
      end: '2026-10-09T15:00:00',
      showAs: 'Tentative',
    }),
    rawEvent({
      id: outlookId(7),
      subject: 'Abwesend',
      start: '2026-10-16T00:00:00',
      end: '2026-10-17T00:00:00',
      allDay: true,
      showAs: 'Oof',
    }),
  ];
}
