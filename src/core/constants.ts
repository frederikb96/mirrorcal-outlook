/** Hosts the Outlook web app is served from; the only hosts an Outlook request may address. */
export const OUTLOOK_HOSTS = [
  'outlook.cloud.microsoft',
  'outlook.office.com',
  'outlook.office365.com',
] as const;

export type OutlookHost = (typeof OUTLOOK_HOSTS)[number];

/** Path of the one Outlook endpoint that is ever called. */
export const CALENDAR_VIEW_PATH = '/api/v2.0/me/calendarview';

/**
 * Paths accepted for a calendar-view continuation link. Outlook may spell the continuation
 * as `/Me/CalendarView` or `/Users('<id>')/CalendarView`; both are the same read-only resource.
 */
export const CALENDAR_VIEW_PATH_PATTERN = /^\/api\/v2\.0\/(?:me|users\([^)/]*\))\/calendarview$/;

/** The access token audience that Outlook web's own API calls carry. */
export const OUTLOOK_TOKEN_AUDIENCE = 'https://outlook.office.com';

/** Request paths (per Outlook host) on which the page's own bearer token is observed. */
export const TOKEN_CAPTURE_PATHS = ['/api/*'] as const;

export const PAGE_SIZE = 200;
export const MAX_PAGES = 100;
export const REQUEST_TIMEOUT_MS = 25_000;
export const WRITE_CONCURRENCY = 4;

/** A plan is refused unless confirmed when it deletes more than this many own objects... */
export const MASS_DELETE_MIN = 20;
/** ...and more than this share of all own objects. */
export const MASS_DELETE_SHARE = 0.5;

export const FALLBACK_TITLE = 'Busy';
export const UID_PREFIX = 'mirrorcal-';
export const UID_HEX_LENGTH = 32;
export const PROP_OWNER = 'X-MIRRORCAL-OWNER';
export const PROP_HASH = 'X-MIRRORCAL-HASH';
/** Part of every content hash; bump when the generated ICS changes shape. */
export const FORMAT_VERSION = 1;
export const PRODID = '-//mirrorcal-outlook//EN';
