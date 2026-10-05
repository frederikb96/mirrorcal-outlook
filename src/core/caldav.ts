/**
 * Minimal CalDAV client for one calendar collection: list every event object, create, update and
 * delete by URL. Every URL it touches is checked to lie inside the configured collection.
 */
import { REQUEST_TIMEOUT_MS } from './constants.ts';
import { normalizeCaldavUrl } from './settings.ts';
import type { FetchLike, RemoteObject } from './types.ts';

export interface CalDavTarget {
  url: string;
  username: string;
  password: string;
}

export type CalDavErrorCode = 'auth' | 'http' | 'network' | 'malformed' | 'scope' | 'conflict';

export class CalDavError extends Error {
  readonly code: CalDavErrorCode;
  readonly status?: number;
  constructor(code: CalDavErrorCode, message: string, status?: number) {
    super(message);
    this.name = 'CalDavError';
    this.code = code;
    if (status !== undefined) this.status = status;
  }
}

function authorization(target: CalDavTarget): string {
  const bytes = new TextEncoder().encode(`${target.username}:${target.password}`);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return `Basic ${btoa(binary)}`;
}

function decodedPath(url: URL): string {
  try {
    return decodeURIComponent(url.pathname);
  } catch {
    return url.pathname;
  }
}

/** Resolves a server-reported href or URL and checks it lies inside the collection. */
export function resolveInCollection(target: CalDavTarget, href: string): string {
  const collection = new URL(normalizeCaldavUrl(target.url));
  let resolved: URL;
  try {
    resolved = new URL(href, collection);
  } catch {
    throw new CalDavError('scope', 'The server reported an invalid object address');
  }
  const prefix = decodedPath(collection);
  if (resolved.origin !== collection.origin || !decodedPath(resolved).startsWith(prefix)) {
    throw new CalDavError('scope', 'The server reported an object outside the configured calendar');
  }
  return resolved.toString();
}

export function objectUrlForUid(target: CalDavTarget, uid: string): string {
  return resolveInCollection(target, `${encodeURIComponent(uid)}.ics`);
}

export interface MultistatusEntry {
  href: string;
  etag?: string;
  calendarData?: string;
}

const XML_ENTITIES: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

function decodeXml(text: string): string {
  const plain = (chunk: string) =>
    chunk.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, (_, ref: string) => {
      if (ref.startsWith('#x')) return String.fromCodePoint(Number.parseInt(ref.slice(2), 16));
      if (ref.startsWith('#')) return String.fromCodePoint(Number.parseInt(ref.slice(1), 10));
      return XML_ENTITIES[ref] as string;
    });
  // CDATA sections are literal; everything between them carries entities.
  return text
    .split(/(<!\[CDATA\[[\s\S]*?\]\]>)/)
    .map((part) => (part.startsWith('<![CDATA[') ? part.slice(9, -3) : plain(part)))
    .join('');
}

function elementText(block: string, localName: string): string | undefined {
  const re = new RegExp(
    `<(?:[\\w.-]+:)?${localName}(?:\\s[^>]*)?>([\\s\\S]*?)</(?:[\\w.-]+:)?${localName}>`,
  );
  const match = re.exec(block);
  return match?.[1] === undefined ? undefined : decodeXml(match[1]);
}

/** Parses a WebDAV multistatus without a DOM; prefixes are ignored, only 200 propstats count. */
export function parseMultistatus(xml: string): MultistatusEntry[] {
  if (!/<(?:[\w.-]+:)?multistatus[\s>]/.test(xml)) {
    throw new CalDavError('malformed', 'The server answer is not a multistatus document');
  }
  const entries: MultistatusEntry[] = [];
  const responses = xml.matchAll(
    /<(?:[\w.-]+:)?response(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w.-]+:)?response>/g,
  );
  for (const response of responses) {
    const block = response[1] as string;
    const href = elementText(block, 'href')?.trim();
    if (!href) throw new CalDavError('malformed', 'A multistatus response has no href');
    const entry: MultistatusEntry = { href };
    const propstats = block.matchAll(
      /<(?:[\w.-]+:)?propstat(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w.-]+:)?propstat>/g,
    );
    for (const propstat of propstats) {
      const body = propstat[1] as string;
      const status = /HTTP\/\d(?:\.\d)?\s+(\d{3})/.exec(elementText(body, 'status') ?? '');
      if (status?.[1] !== '200') continue;
      const etag = elementText(body, 'getetag')?.trim();
      const data = elementText(body, 'calendar-data');
      if (etag) entry.etag = etag;
      if (data !== undefined) entry.calendarData = data;
    }
    entries.push(entry);
  }
  return entries;
}

async function send(
  fetchImpl: FetchLike,
  target: CalDavTarget,
  url: string,
  init: RequestInit & { headers?: Record<string, string> },
): Promise<Response> {
  const scoped = resolveInCollection(target, url);
  try {
    return await fetchImpl(scoped, {
      ...init,
      headers: { ...init.headers, Authorization: authorization(target) },
      credentials: 'omit',
      redirect: 'error',
      cache: 'no-store',
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    });
  } catch (error) {
    const timedOut = error instanceof Error && error.name === 'TimeoutError';
    throw new CalDavError(
      'network',
      timedOut
        ? 'The calendar server did not answer in time'
        : 'Could not reach the calendar server',
    );
  }
}

function failure(response: Response, doing: string): CalDavError {
  if (response.status === 401) {
    return new CalDavError('auth', 'The calendar server rejected the credentials (401)', 401);
  }
  return new CalDavError('http', `${doing} failed with HTTP ${response.status}`, response.status);
}

const LIST_BODY = `<?xml version="1.0" encoding="utf-8"?>
<c:calendar-query xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">
  <d:prop><d:getetag/><c:calendar-data/></d:prop>
  <c:filter><c:comp-filter name="VCALENDAR"><c:comp-filter name="VEVENT"/></c:comp-filter></c:filter>
</c:calendar-query>`;

const CTAG_BODY = `<?xml version="1.0" encoding="utf-8"?>
<d:propfind xmlns:d="DAV:" xmlns:cs="http://calendarserver.org/ns/"><d:prop><cs:getctag/></d:prop></d:propfind>`;

/**
 * The collection's change tag: one small request that moves whenever anything in the calendar
 * changes. Undefined when the server does not offer it or the request fails; callers then read
 * the whole calendar.
 */
export async function getCtag(
  fetchImpl: FetchLike,
  target: CalDavTarget,
): Promise<string | undefined> {
  try {
    const collection = resolveInCollection(target, normalizeCaldavUrl(target.url));
    const response = await send(fetchImpl, target, collection, {
      method: 'PROPFIND',
      headers: { Depth: '0', 'Content-Type': 'application/xml; charset=utf-8' },
      body: CTAG_BODY,
    });
    if (response.status !== 207) return undefined;
    const ctag = elementText(await response.text(), 'getctag')?.trim();
    return ctag || undefined;
  } catch {
    return undefined;
  }
}

export interface ObjectListing {
  objects: RemoteObject[];
  /** Entries the server listed without readable calendar data. */
  unreadable: number;
}

/** Reads every event object in the collection, whatever its date. */
export async function listObjects(
  fetchImpl: FetchLike,
  target: CalDavTarget,
): Promise<ObjectListing> {
  const collection = resolveInCollection(target, normalizeCaldavUrl(target.url));
  const response = await send(fetchImpl, target, collection, {
    method: 'REPORT',
    headers: { Depth: '1', 'Content-Type': 'application/xml; charset=utf-8' },
    body: LIST_BODY,
  });
  if (response.status !== 207) throw failure(response, 'Listing the calendar');
  const entries = parseMultistatus(await response.text());
  const objects: RemoteObject[] = [];
  let unreadable = 0;
  for (const entry of entries) {
    const href = resolveInCollection(target, entry.href);
    if (href === collection) continue;
    if (entry.calendarData === undefined) {
      unreadable += 1;
      continue;
    }
    objects.push({ href, etag: entry.etag ?? '', ics: entry.calendarData });
  }
  objects.sort((a, b) => a.href.localeCompare(b.href));
  return { objects, unreadable };
}

export type WriteCondition = { create: true } | { replace: string };

/** Creates (`create`) or replaces (`replace: etag`) one object; never overwrites unconditionally. */
export async function putObject(
  fetchImpl: FetchLike,
  target: CalDavTarget,
  url: string,
  ics: string,
  condition: WriteCondition,
): Promise<void> {
  const headers: Record<string, string> = { 'Content-Type': 'text/calendar; charset=utf-8' };
  if ('create' in condition) headers['If-None-Match'] = '*';
  else if (condition.replace !== '') headers['If-Match'] = condition.replace;
  const response = await send(fetchImpl, target, url, { method: 'PUT', headers, body: ics });
  if (response.status === 412) {
    throw new CalDavError('conflict', 'The object changed on the server', 412);
  }
  if (response.status !== 201 && response.status !== 204) throw failure(response, 'Writing');
}

export async function deleteObject(
  fetchImpl: FetchLike,
  target: CalDavTarget,
  url: string,
  etag: string,
): Promise<void> {
  const headers: Record<string, string> = etag === '' ? {} : { 'If-Match': etag };
  const response = await send(fetchImpl, target, url, { method: 'DELETE', headers });
  if (response.status === 412) {
    throw new CalDavError('conflict', 'The object changed on the server', 412);
  }
  // A 404 means it is already gone, which is the state a delete asks for.
  if (response.status !== 204 && response.status !== 200 && response.status !== 404) {
    throw failure(response, 'Deleting');
  }
}
