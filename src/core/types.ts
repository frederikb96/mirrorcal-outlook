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

/** Outlook's sensitivity label of an event. */
export type Sensitivity = 'Normal' | 'Personal' | 'Private' | 'Confidential';

export const SENSITIVITY_VALUES: readonly Sensitivity[] = [
  'Normal',
  'Personal',
  'Private',
  'Confidential',
];

/** CLASS value an event is written with; absent for ordinary events. */
export type Classification = 'PRIVATE' | 'CONFIDENTIAL';

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
  sensitivity: Sensitivity;
  location?: string;
  /** The body, or only its preview when the description policy replaces it. */
  body?: string;
  organizer?: Person;
  attendees?: SourceAttendee[];
}

export type TextPolicy = { mode: 'copy' } | { mode: 'drop' } | { mode: 'replace'; text: string };

/** What happens to events Outlook marks as private, personal or confidential. */
export type PrivatePolicy = { mode: 'copy' } | { mode: 'replace'; text: string } | { mode: 'skip' };

export type AttendeePolicy = 'copy' | 'drop';

export interface Settings {
  enabled: boolean;
  caldavUrl: string;
  /**
   * Optional second calendar on the same server. Events the filters and availability exclusions
   * leave out of the main calendar are mirrored here instead; empty means they are not mirrored.
   */
  excludedCaldavUrl: string;
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
    privateEvents: PrivatePolicy;
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
  classification?: Classification;
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
