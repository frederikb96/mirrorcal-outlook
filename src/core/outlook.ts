/**
 * Outlook reader. The calendar-view read in this module is the only thing that ever talks to
 * Outlook: there is no generic request helper, the single request constructor can only produce a
 * GET to the calendar-view path on an allow-listed host, and nothing here accepts a method.
 */
import {
  CALENDAR_VIEW_PATH,
  CALENDAR_VIEW_PATH_PATTERN,
  MAX_PAGES,
  OUTLOOK_HOSTS,
  type OutlookHost,
  PAGE_SIZE,
  REQUEST_TIMEOUT_MS,
} from './constants.ts';
import { normalizeToken } from './token.ts';
import {
  type FetchLike,
  type Person,
  type Settings,
  SHOW_AS_VALUES,
  type ShowAs,
  type SourceAttendee,
  type SourceEvent,
} from './types.ts';

export type OutlookErrorCode = 'request' | 'auth' | 'http' | 'network' | 'malformed' | 'paging';

export class OutlookReadError extends Error {
  readonly code: OutlookErrorCode;
  constructor(code: OutlookErrorCode, message: string) {
    super(message);
    this.name = 'OutlookReadError';
    this.code = code;
  }
}

/** Which optional fields the configured policies need; everything else is not requested. */
export interface FieldNeeds {
  location: boolean;
  body: boolean;
  attendees: boolean;
}

export function fieldNeeds(settings: Settings): FieldNeeds {
  const { policies } = settings;
  return {
    location: policies.location.mode === 'copy',
    body: policies.description.mode === 'copy',
    attendees: policies.attendees === 'copy',
  };
}

const BASE_FIELDS = ['Id', 'Subject', 'Start', 'End', 'IsAllDay', 'IsCancelled', 'ShowAs'];

function selectedFields(needs: FieldNeeds): string[] {
  const fields = [...BASE_FIELDS];
  if (needs.location) fields.push('Location');
  if (needs.body) fields.push('Body');
  if (needs.attendees) fields.push('Attendees', 'Organizer');
  return fields;
}

export function calendarViewUrl(host: OutlookHost, start: Date, end: Date, needs: FieldNeeds) {
  const query = [
    `startDateTime=${encodeURIComponent(start.toISOString())}`,
    `endDateTime=${encodeURIComponent(end.toISOString())}`,
    `$top=${PAGE_SIZE}`,
    `$select=${selectedFields(needs).join(',')}`,
  ].join('&');
  return `https://${host}${CALENDAR_VIEW_PATH}?${query}`;
}

/** Throws unless the URL is an https GET target on an Outlook host at a calendar-view path. */
function assertCalendarViewUrl(raw: string): URL {
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    throw new OutlookReadError('request', 'Outlook request URL is not valid');
  }
  const hostAllowed = (OUTLOOK_HOSTS as readonly string[]).includes(url.hostname);
  if (
    url.protocol !== 'https:' ||
    !hostAllowed ||
    url.port !== '' ||
    url.username !== '' ||
    url.password !== '' ||
    !CALENDAR_VIEW_PATH_PATTERN.test(url.pathname.toLowerCase())
  ) {
    throw new OutlookReadError(
      'request',
      `Refusing Outlook request to ${url.protocol}//${url.host}${url.pathname}`,
    );
  }
  return url;
}

/** The single way to build an Outlook request. Always a credential-free GET. */
export function buildCalendarViewRequest(
  rawUrl: string,
  token: string,
): { url: string; init: RequestInit } {
  const url = assertCalendarViewUrl(rawUrl);
  let bearer: string;
  try {
    bearer = normalizeToken(token);
  } catch (error) {
    throw new OutlookReadError('request', (error as Error).message);
  }
  return {
    url: url.toString(),
    init: {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${bearer}`,
        Accept: 'application/json',
        Prefer: 'outlook.timezone="UTC", outlook.body-content-type="text"',
      },
      credentials: 'omit',
      redirect: 'error',
      cache: 'no-store',
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    },
  };
}

type Json = Record<string, unknown>;

function isRecord(value: unknown): value is Json {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

const DATE_TIME = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.\d+)?$/;

function malformed(what: string): never {
  throw new OutlookReadError('malformed', `Outlook returned an unexpected response: ${what}`);
}

function parseDateTime(value: unknown, label: string): string {
  if (!isRecord(value)) return malformed(`${label} is missing`);
  const { DateTime, TimeZone } = value;
  if (typeof DateTime !== 'string') return malformed(`${label} has no date`);
  if (typeof TimeZone !== 'string' || TimeZone.toUpperCase() !== 'UTC') {
    return malformed(`${label} is not in UTC`);
  }
  const m = DATE_TIME.exec(DateTime);
  if (!m) return malformed(`${label} is not a date`);
  return `${m[1]}-${m[2]}-${m[3]}T${m[4]}:${m[5]}:${m[6]}`;
}

function parsePerson(value: unknown): Person | undefined {
  if (!isRecord(value)) return undefined;
  const email = isRecord(value.EmailAddress) ? value.EmailAddress : undefined;
  if (!email || typeof email.Address !== 'string' || email.Address === '') return undefined;
  return { name: typeof email.Name === 'string' ? email.Name : '', address: email.Address };
}

function parseAttendees(value: unknown): SourceAttendee[] {
  if (!Array.isArray(value)) return [];
  const attendees: SourceAttendee[] = [];
  for (const item of value) {
    const person = parsePerson(item);
    if (!person || !isRecord(item)) continue;
    const status = isRecord(item.Status) ? item.Status : {};
    attendees.push({
      ...person,
      type: typeof item.Type === 'string' ? item.Type : 'Required',
      response: typeof status.Response === 'string' ? status.Response : 'None',
    });
  }
  return attendees;
}

/** Returns undefined for cancelled instances; throws on anything that is not a usable event. */
function parseEvent(item: unknown, index: number): SourceEvent | undefined {
  if (!isRecord(item)) return malformed(`item ${index} is not an object`);
  if (typeof item.Id !== 'string' || item.Id === '') return malformed(`item ${index} has no id`);
  if (typeof item.IsCancelled !== 'boolean') return malformed(`item ${index} has no cancel flag`);
  if (item.IsCancelled) return undefined;
  if (typeof item.IsAllDay !== 'boolean') return malformed(`item ${index} has no all-day flag`);
  const start = parseDateTime(item.Start, `item ${index} start`);
  const end = parseDateTime(item.End, `item ${index} end`);
  if (end < start) return malformed(`item ${index} ends before it starts`);
  const showAs = SHOW_AS_VALUES.find((v) => v === item.ShowAs) ?? ('Unknown' as ShowAs);
  const event: SourceEvent = {
    id: item.Id,
    subject: typeof item.Subject === 'string' ? item.Subject : '',
    start,
    end,
    allDay: item.IsAllDay,
    showAs,
  };
  if (isRecord(item.Location) && typeof item.Location.DisplayName === 'string') {
    event.location = item.Location.DisplayName;
  }
  if (isRecord(item.Body) && typeof item.Body.Content === 'string') event.body = item.Body.Content;
  if ('Attendees' in item) event.attendees = parseAttendees(item.Attendees);
  const organizer = parsePerson(item.Organizer);
  if (organizer) event.organizer = organizer;
  return event;
}

async function fetchPage(fetchImpl: FetchLike, url: string, token: string): Promise<Json> {
  const request = buildCalendarViewRequest(url, token);
  let response: Response;
  try {
    response = await fetchImpl(request.url, request.init);
  } catch (error) {
    const timedOut = error instanceof Error && error.name === 'TimeoutError';
    throw new OutlookReadError(
      'network',
      timedOut ? 'Outlook did not answer in time' : 'Could not reach Outlook',
    );
  }
  if (response.status === 401) {
    throw new OutlookReadError('auth', 'Outlook rejected the token (401)');
  }
  if (response.status !== 200) {
    throw new OutlookReadError('http', `Outlook answered HTTP ${response.status}`);
  }
  let body: unknown;
  try {
    body = await response.json();
  } catch {
    return malformed('the body is not JSON');
  }
  if (!isRecord(body) || !Array.isArray(body.value)) return malformed('no event list');
  return body;
}

export interface CalendarViewRead {
  /** Non-cancelled instances. */
  events: SourceEvent[];
  /** Every item the server returned, cancelled ones included. */
  rawCount: number;
  pages: number;
}

export interface CalendarViewQuery {
  fetch: FetchLike;
  token: string;
  host: OutlookHost;
  start: Date;
  end: Date;
  needs: FieldNeeds;
}

/**
 * Reads every instance in the window, following continuation links. Any failure, unexpected
 * shape or off-allow-list link aborts the whole read: a partial result is never returned.
 */
export async function readCalendarView(query: CalendarViewQuery): Promise<CalendarViewRead> {
  const events: SourceEvent[] = [];
  const seen = new Set<string>();
  let rawCount = 0;
  let pages = 0;
  let next: string | undefined = calendarViewUrl(query.host, query.start, query.end, query.needs);
  while (next !== undefined) {
    if (seen.has(next)) throw new OutlookReadError('paging', 'Outlook paging looped');
    if (pages >= MAX_PAGES) throw new OutlookReadError('paging', 'Outlook returned too many pages');
    seen.add(next);
    const page = await fetchPage(query.fetch, next, query.token);
    pages += 1;
    const items = page.value as unknown[];
    for (const item of items) {
      const event = parseEvent(item, rawCount);
      rawCount += 1;
      if (event) events.push(event);
    }
    const link = page['@odata.nextLink'];
    if (link === undefined || link === null) {
      next = undefined;
    } else if (typeof link === 'string') {
      next = link;
    } else {
      return malformed('the continuation link is not text');
    }
  }
  return { events, rawCount, pages };
}
