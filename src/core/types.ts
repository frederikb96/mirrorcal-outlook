import type { OutlookHost } from './constants.ts';

export type FetchLike = (input: string, init?: RequestInit) => Promise<Response>;

export type ShowAs = 'Free' | 'Tentative' | 'Busy' | 'Oof' | 'WorkingElsewhere' | 'Unknown';

export const SHOW_AS_VALUES: readonly ShowAs[] = [
  'Free',
  'Tentative',
  'Busy',
  'Oof',
  'WorkingElsewhere',
  'Unknown',
];

export interface Person {
  name: string;
  address: string;
}

export interface SourceAttendee extends Person {
  /** Required | Optional | Resource */
  type: string;
  /** Outlook response: Accepted, Declined, TentativelyAccepted, NotResponded, Organizer, None */
  response: string;
}

/** One Outlook calendar instance, normalised. Optional fields are only read when a policy needs them. */
export interface SourceEvent {
  id: string;
  subject: string;
  /** `YYYY-MM-DDTHH:MM:SS`, UTC. For all-day events the date part is the calendar date. */
  start: string;
  /** Same format; exclusive end for all-day events. */
  end: string;
  allDay: boolean;
  showAs: ShowAs;
  location?: string;
  body?: string;
  organizer?: Person;
  attendees?: SourceAttendee[];
}

export type TextPolicy = { mode: 'copy' } | { mode: 'drop' } | { mode: 'replace'; text: string };

export type AttendeePolicy = 'copy' | 'drop';

export interface Settings {
  enabled: boolean;
  caldavUrl: string;
  username: string;
  password: string;
  ownerId: string;
  outlookHost: OutlookHost;
  daysBack: number;
  daysAhead: number;
  intervalMinutes: number;
  policies: {
    title: TextPolicy;
    description: TextPolicy;
    location: TextPolicy;
    attendees: AttendeePolicy;
  };
  /** Title patterns; matched case-insensitively against the trimmed title, `*` is a wildcard. */
  titleFilters: string[];
  /** Events with one of these availabilities are not mirrored. */
  excludeShowAs: ShowAs[];
}

/** An event as it will be written, after filters and policies. */
export interface MirrorEvent {
  summary: string;
  description?: string;
  location?: string;
  allDay: boolean;
  /** `YYYY-MM-DD` for all-day events, `YYYY-MM-DDTHH:MM:SS` (UTC) otherwise. */
  start: string;
  end: string;
  showAs: ShowAs;
  organizer?: Person;
  attendees: SourceAttendee[];
}

export interface DesiredObject {
  uid: string;
  href: string;
  ics: string;
  hash: string;
  summary: string;
  start: string;
  allDay: boolean;
}

export interface RemoteObject {
  /** Absolute URL of the object on the server. */
  href: string;
  etag: string;
  ics: string;
}
