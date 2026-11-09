import type { FetchLike } from '../../src/core/types.ts';
import { COLLECTION } from './fixtures.ts';

export interface DavRequest {
  method: string;
  url: string;
  headers: Record<string, string>;
  body?: string;
}

interface Stored {
  ics: string;
  etag: string;
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/\r/g, '&#13;');
}

/**
 * A CalDAV calendar that keeps the exact ICS text it was given and answers like Nextcloud:
 * conditional writes, 207 listings with escaped calendar data, path-only hrefs, unique UIDs.
 */
export class FakeCalDav {
  readonly objects = new Map<string, Stored>();
  readonly requests: DavRequest[] = [];
  readonly username = 'alice';
  /** The password the server accepts; change it to simulate a rotated app password. */
  password = 'test-app-password';
  /** What the collection URL is, for the inspection request. */
  collection: { kind: 'calendar' | 'home' | 'tasks'; name?: string } = {
    kind: 'calendar',
    name: 'Work',
  };
  private counter = 0;
  private ctag = 0;
  /** Fail the next request(s) matching this predicate with the given status. */
  failWhen?: (method: string, url: string) => number | undefined;

  private root: string;

  constructor(collection = COLLECTION) {
    this.root = new URL(collection).pathname;
  }

  /** Whether a request URL addresses this calendar. */
  owns(url: string): boolean {
    return new URL(url).pathname.startsWith(this.root);
  }

  private nextEtag(): string {
    this.counter += 1;
    return `"etag-${this.counter}"`;
  }

  get writes(): DavRequest[] {
    return this.requests.filter((r) => r.method === 'PUT' || r.method === 'DELETE');
  }

  clearRequests() {
    this.requests.length = 0;
  }

  /** Puts an object in place without going through HTTP (a foreign or hand-made object). */
  seed(name: string, ics: string) {
    this.ctag += 1;
    this.objects.set(name, { ics, etag: this.nextEtag() });
  }

  /** Changes a stored object the way another client would. */
  edit(name: string, change: (ics: string) => string) {
    this.ctag += 1;
    const object = this.objects.get(name);
    if (!object) throw new Error(`no object ${name}`);
    this.objects.set(name, { ics: change(object.ics), etag: this.nextEtag() });
  }

  remove(name: string) {
    this.ctag += 1;
    this.objects.delete(name);
  }

  private uidOf(ics: string): string | undefined {
    return /^UID:(.*)$/m.exec(ics.replace(/\r\n[ \t]/g, ''))?.[1]?.trim();
  }

  fetch: FetchLike = async (input, init) => {
    const url = new URL(input);
    const method = init?.method ?? 'GET';
    const headers: Record<string, string> = {};
    for (const [k, v] of Object.entries((init?.headers ?? {}) as Record<string, string>)) {
      headers[k.toLowerCase()] = v;
    }
    this.requests.push({
      method,
      url: input,
      headers,
      ...(typeof init?.body === 'string' ? { body: init.body } : {}),
    });
    const expected = `Basic ${btoa(`${this.username}:${this.password}`)}`;
    if (headers.authorization !== expected) return new Response('', { status: 401 });
    const forced = this.failWhen?.(method, input);
    if (forced !== undefined) return new Response('', { status: forced });

    if (method === 'PROPFIND' && String(init?.body).includes('resourcetype')) {
      const { kind, name } = this.collection;
      const type = kind === 'home' ? '<d:collection/>' : '<d:collection/><cal:calendar/>';
      const components =
        kind === 'calendar'
          ? '<cal:comp name="VEVENT"/>'
          : kind === 'tasks'
            ? '<cal:comp name="VTODO"/>'
            : '';
      const set =
        kind === 'home'
          ? ''
          : `<cal:supported-calendar-component-set>${components}</cal:supported-calendar-component-set>`;
      const xml = `<?xml version="1.0"?>
<d:multistatus xmlns:d="DAV:" xmlns:cal="urn:ietf:params:xml:ns:caldav"><d:response><d:href>${this.root}</d:href><d:propstat><d:prop><d:resourcetype>${type}</d:resourcetype>${name ? `<d:displayname>${name}</d:displayname>` : ''}${set}</d:prop><d:status>HTTP/1.1 200 OK</d:status></d:propstat></d:response></d:multistatus>`;
      return new Response(xml, { status: 207 });
    }

    if (method === 'PROPFIND') {
      const xml = `<?xml version="1.0"?>
<d:multistatus xmlns:d="DAV:" xmlns:cs="http://calendarserver.org/ns/"><d:response><d:href>${this.root}</d:href><d:propstat><d:prop><cs:getctag>http://sabre.io/ns/sync/${this.ctag}</cs:getctag></d:prop><d:status>HTTP/1.1 200 OK</d:status></d:propstat></d:response></d:multistatus>`;
      return new Response(xml, { status: 207 });
    }

    if (method === 'REPORT') {
      const parts = [...this.objects.entries()].map(
        ([name, o]) => `<d:response>
<d:href>${this.root}${name}</d:href>
<d:propstat>
<d:prop>
<d:getetag>${escapeXml(o.etag)}</d:getetag>
<cal:calendar-data>${escapeXml(o.ics)}</cal:calendar-data>
</d:prop>
<d:status>HTTP/1.1 200 OK</d:status>
</d:propstat>
</d:response>`,
      );
      const xml = `<?xml version="1.0"?>
<d:multistatus xmlns:d="DAV:" xmlns:s="http://sabredav.org/ns" xmlns:cal="urn:ietf:params:xml:ns:caldav" xmlns:cs="http://calendarserver.org/ns/" xmlns:oc="http://owncloud.org/ns" xmlns:nc="http://nextcloud.org/ns">
${parts.join('\n')}
</d:multistatus>`;
      return new Response(xml, {
        status: 207,
        headers: { 'Content-Type': 'application/xml; charset=utf-8' },
      });
    }

    const name = url.pathname.slice(this.root.length);
    const existing = this.objects.get(name);
    if (method === 'PUT') {
      const ics = String(init?.body ?? '');
      if (headers['if-none-match'] === '*' && existing) return new Response('', { status: 412 });
      if (headers['if-match'] !== undefined && existing?.etag !== headers['if-match']) {
        return new Response('', { status: 412 });
      }
      const uid = this.uidOf(ics);
      for (const [other, o] of this.objects) {
        if (other !== name && uid && this.uidOf(o.ics) === uid)
          return new Response('', { status: 409 });
      }
      const etag = this.nextEtag();
      this.ctag += 1;
      this.objects.set(name, { ics, etag });
      return new Response('', { status: existing ? 204 : 201, headers: { ETag: etag } });
    }
    if (method === 'DELETE') {
      if (!existing) return new Response('', { status: 404 });
      if (headers['if-match'] !== undefined && existing.etag !== headers['if-match']) {
        return new Response('', { status: 412 });
      }
      this.ctag += 1;
      this.objects.delete(name);
      return new Response('', { status: 204 });
    }
    return new Response('', { status: 405 });
  };
}

/** Routes by host so one fetch can serve both fakes. */
export function routeFetch(routes: Record<string, FetchLike>): FetchLike {
  return async (input, init) => {
    const host = new URL(input).hostname;
    const handler = routes[host];
    if (!handler) throw new Error(`unexpected request to ${host}`);
    return handler(input, init);
  };
}

/** Serves several calendars on one host, each request going to the calendar whose path it names. */
export function routeCalendars(...calendars: FakeCalDav[]): FetchLike {
  return async (input, init) => {
    const calendar = calendars.find((c) => c.owns(input));
    if (!calendar) throw new Error(`no calendar at ${input}`);
    return calendar.fetch(input, init);
  };
}
