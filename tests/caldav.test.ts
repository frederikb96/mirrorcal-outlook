import { describe, expect, test } from 'bun:test';
import {
  CalDavError,
  type CalDavTarget,
  deleteObject,
  listObjects,
  objectUrlForUid,
  parseMultistatus,
  putObject,
  resolveInCollection,
} from '../src/core/caldav.ts';
import type { FetchLike } from '../src/core/types.ts';
import { FakeCalDav } from './helpers/fake-caldav.ts';
import { COLLECTION } from './helpers/fixtures.ts';

const target: CalDavTarget = { url: COLLECTION, username: 'alice', password: 'test-app-password' };

/** Shaped like a Nextcloud calendar-query answer: path hrefs, quoted etags, CR as a character reference. */
const NEXTCLOUD_REPORT = `<?xml version="1.0"?>
<d:multistatus xmlns:d="DAV:" xmlns:s="http://sabredav.org/ns" xmlns:cal="urn:ietf:params:xml:ns:caldav" xmlns:cs="http://calendarserver.org/ns/" xmlns:oc="http://owncloud.org/ns" xmlns:nc="http://nextcloud.org/ns">
 <d:response>
  <d:href>/remote.php/dav/calendars/alice/work/mirrorcal-0123.ics</d:href>
  <d:propstat>
   <d:prop>
    <d:getetag>&quot;5f3a9c1b2d&quot;</d:getetag>
    <cal:calendar-data>BEGIN:VCALENDAR&#13;
VERSION:2.0&#13;
BEGIN:VEVENT&#13;
UID:mirrorcal-0123&#13;
SUMMARY:Fish &amp; Chips &lt;lunch&gt;&#13;
X-MIRRORCAL-OWNER:outlook&#13;
END:VEVENT&#13;
END:VCALENDAR&#13;
</cal:calendar-data>
   </d:prop>
   <d:status>HTTP/1.1 200 OK</d:status>
  </d:propstat>
 </d:response>
 <d:response>
  <d:href>/remote.php/dav/calendars/alice/work/Dentist%20%28Tue%29.ics</d:href>
  <d:propstat>
   <d:prop>
    <d:getetag>&quot;aa11&quot;</d:getetag>
    <cal:calendar-data><![CDATA[BEGIN:VCALENDAR
BEGIN:VEVENT
UID:dentist
SUMMARY:Raw & unescaped
END:VEVENT
END:VCALENDAR
]]></cal:calendar-data>
   </d:prop>
   <d:status>HTTP/1.1 200 OK</d:status>
  </d:propstat>
 </d:response>
 <d:response>
  <d:href>/remote.php/dav/calendars/alice/work/no-data.ics</d:href>
  <d:propstat>
   <d:prop><d:getetag>&quot;bb22&quot;</d:getetag></d:prop>
   <d:status>HTTP/1.1 200 OK</d:status>
  </d:propstat>
  <d:propstat>
   <d:prop><cal:calendar-data/></d:prop>
   <d:status>HTTP/1.1 404 Not Found</d:status>
  </d:propstat>
 </d:response>
</d:multistatus>`;

describe('multistatus parsing', () => {
  test('reads hrefs, etags and calendar data of a Nextcloud answer', () => {
    const entries = parseMultistatus(NEXTCLOUD_REPORT);
    expect(entries.map((e) => e.href)).toEqual([
      '/remote.php/dav/calendars/alice/work/mirrorcal-0123.ics',
      '/remote.php/dav/calendars/alice/work/Dentist%20%28Tue%29.ics',
      '/remote.php/dav/calendars/alice/work/no-data.ics',
    ]);
    expect(entries[0]?.etag).toBe('"5f3a9c1b2d"');
    expect(entries[0]?.calendarData).toBe(
      'BEGIN:VCALENDAR\r\nVERSION:2.0\r\nBEGIN:VEVENT\r\nUID:mirrorcal-0123\r\nSUMMARY:Fish & Chips <lunch>\r\nX-MIRRORCAL-OWNER:outlook\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n',
    );
    expect(entries[1]?.calendarData).toContain('SUMMARY:Raw & unescaped');
    expect(entries[2]?.calendarData).toBeUndefined();
    expect(entries[2]?.etag).toBe('"bb22"');
  });

  test('does not depend on the namespace prefix', () => {
    const xml = `<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav"><D:response><D:href>/a.ics</D:href><D:propstat><D:prop><D:getetag>"e"</D:getetag><C:calendar-data>X</C:calendar-data></D:prop><D:status>HTTP/1.1 200 OK</D:status></D:propstat></D:response></D:multistatus>`;
    expect(parseMultistatus(xml)).toEqual([{ href: '/a.ics', etag: '"e"', calendarData: 'X' }]);
  });

  test('only properties reported with status 200 are used', () => {
    const xml = `<d:multistatus xmlns:d="DAV:" xmlns:cal="urn:ietf:params:xml:ns:caldav"><d:response><d:href>/a.ics</d:href><d:propstat><d:prop><d:getetag>"ok"</d:getetag></d:prop><d:status>HTTP/1.1 200 OK</d:status></d:propstat><d:propstat><d:prop><cal:calendar-data>left over</cal:calendar-data></d:prop><d:status>HTTP/1.1 404 Not Found</d:status></d:propstat></d:response></d:multistatus>`;
    expect(parseMultistatus(xml)).toEqual([{ href: '/a.ics', etag: '"ok"' }]);
  });

  test('an empty listing is valid, anything else that is not a multistatus is not', () => {
    expect(parseMultistatus('<d:multistatus xmlns:d="DAV:"></d:multistatus>')).toEqual([]);
    expect(() => parseMultistatus('<html>Sign in</html>')).toThrow(CalDavError);
    expect(() => parseMultistatus('')).toThrow(CalDavError);
  });

  test('entities are decoded once', () => {
    const xml =
      '<d:multistatus xmlns:d="DAV:"><d:response><d:href>/a</d:href><d:propstat><d:prop><cal:calendar-data xmlns:cal="x">&amp;lt;&#x41;&#66;</cal:calendar-data></d:prop><d:status>HTTP/1.1 200 OK</d:status></d:propstat></d:response></d:multistatus>';
    expect(parseMultistatus(xml)[0]?.calendarData).toBe('&lt;AB');
  });
});

describe('listing', () => {
  test('resolves hrefs against the collection and keeps the ICS text intact', async () => {
    const fetchImpl: FetchLike = async () => new Response(NEXTCLOUD_REPORT, { status: 207 });
    const { objects, unreadable } = await listObjects(fetchImpl, target);
    expect(unreadable).toBe(1);
    expect(objects.map((o) => o.href)).toEqual([
      `${COLLECTION}Dentist%20%28Tue%29.ics`,
      `${COLLECTION}mirrorcal-0123.ics`,
    ]);
    expect(objects[1]?.etag).toBe('"5f3a9c1b2d"');
    expect(objects[1]?.ics).toContain('SUMMARY:Fish & Chips <lunch>\r\n');
  });

  test('asks for every event with Depth 1, without credentials, with basic auth', async () => {
    const dav = new FakeCalDav();
    await listObjects(dav.fetch, target);
    const [request] = dav.requests;
    expect(request?.method).toBe('REPORT');
    expect(request?.url).toBe(COLLECTION);
    expect(request?.headers.depth).toBe('1');
    expect(request?.headers.authorization).toBe(`Basic ${btoa('alice:test-app-password')}`);
  });

  test('sends non-ASCII passwords as UTF-8', async () => {
    let seen = '';
    const fetchImpl: FetchLike = async (_url, init) => {
      seen = ((init?.headers ?? {}) as Record<string, string>).Authorization as string;
      return new Response('<d:multistatus xmlns:d="DAV:"/>', { status: 207 });
    };
    await listObjects(fetchImpl, { ...target, password: 'pässwörd' });
    const bytes = Uint8Array.from(atob(seen.replace('Basic ', '')), (c) => c.charCodeAt(0));
    expect(new TextDecoder().decode(bytes)).toBe('alice:pässwörd');
  });

  test('failures abort with a typed error', async () => {
    const answer =
      (status: number, body = ''): FetchLike =>
      async () =>
        new Response(body, { status });
    const code = async (f: FetchLike) =>
      (await listObjects(f, target).catch((e: unknown) => e)) as CalDavError;
    expect((await code(answer(401))).code).toBe('auth');
    expect((await code(answer(500))).code).toBe('http');
    expect((await code(answer(200, '<d:multistatus xmlns:d="DAV:"/>'))).code).toBe('http');
    expect((await code(answer(207, 'login page'))).code).toBe('malformed');
    expect(
      (
        await code(async () => {
          throw new TypeError('offline');
        })
      ).code,
    ).toBe('network');
  });

  test('an object outside the configured calendar aborts the listing', async () => {
    const xml = (href: string) =>
      `<d:multistatus xmlns:d="DAV:"><d:response><d:href>${href}</d:href><d:propstat><d:prop><cal:calendar-data xmlns:cal="x">X</cal:calendar-data></d:prop><d:status>HTTP/1.1 200 OK</d:status></d:propstat></d:response></d:multistatus>`;
    for (const href of [
      '/remote.php/dav/calendars/alice/personal/x.ics',
      '/remote.php/dav/calendars/alice/work-2/x.ics',
      '/remote.php/dav/calendars/alice/work/../personal/x.ics',
      'https://elsewhere.example.test/remote.php/dav/calendars/alice/work/x.ics',
    ]) {
      const fetchImpl: FetchLike = async () => new Response(xml(href), { status: 207 });
      const error = (await listObjects(fetchImpl, target).catch((e: unknown) => e)) as CalDavError;
      expect(error).toBeInstanceOf(CalDavError);
      expect(error.code).toBe('scope');
    }
  });
});

describe('writing', () => {
  test('create is conditional on the object not existing', async () => {
    const dav = new FakeCalDav();
    const url = objectUrlForUid(target, 'mirrorcal-1');
    await putObject(dav.fetch, target, url, 'BODY', { create: true });
    expect(dav.requests[0]?.headers['if-none-match']).toBe('*');
    expect(dav.requests[0]?.headers['content-type']).toContain('text/calendar');
    const again = await putObject(dav.fetch, target, url, 'BODY', { create: true }).catch(
      (e: unknown) => e,
    );
    expect((again as CalDavError).code).toBe('conflict');
  });

  test('replace and delete are conditional on the etag', async () => {
    const dav = new FakeCalDav();
    dav.seed('x.ics', 'OLD');
    const etag = dav.objects.get('x.ics')?.etag as string;
    const url = `${COLLECTION}x.ics`;
    const stale = await putObject(dav.fetch, target, url, 'NEW', { replace: '"stale"' }).catch(
      (e: unknown) => e,
    );
    expect((stale as CalDavError).code).toBe('conflict');
    await putObject(dav.fetch, target, url, 'NEW', { replace: etag });
    expect(dav.objects.get('x.ics')?.ics).toBe('NEW');
    const staleDelete = await deleteObject(dav.fetch, target, url, etag).catch((e: unknown) => e);
    expect((staleDelete as CalDavError).code).toBe('conflict');
    await deleteObject(dav.fetch, target, url, dav.objects.get('x.ics')?.etag as string);
    expect(dav.objects.size).toBe(0);
  });

  test('deleting something already gone is fine', async () => {
    const dav = new FakeCalDav();
    await deleteObject(dav.fetch, target, `${COLLECTION}gone.ics`, '');
  });

  test('refuses to write outside the collection', async () => {
    const dav = new FakeCalDav();
    for (const url of [
      'https://dav.example.test/remote.php/dav/calendars/alice/personal/x.ics',
      'https://elsewhere.example.test/remote.php/dav/calendars/alice/work/x.ics',
    ]) {
      await expect(putObject(dav.fetch, target, url, 'X', { create: true })).rejects.toThrow(
        CalDavError,
      );
      await expect(deleteObject(dav.fetch, target, url, '')).rejects.toThrow(CalDavError);
    }
    expect(dav.requests).toHaveLength(0);
  });

  test('uses no cookies and does not follow redirects', async () => {
    let init: RequestInit | undefined;
    const fetchImpl: FetchLike = async (_url, i) => {
      init = i;
      return new Response('', { status: 201 });
    };
    await putObject(fetchImpl, target, `${COLLECTION}a.ics`, 'X', { create: true });
    expect(init?.credentials).toBe('omit');
    expect(init?.redirect).toBe('error');
  });

  test('rejected credentials are reported as such', async () => {
    const fetchImpl: FetchLike = async () => new Response('', { status: 401 });
    const error = await putObject(fetchImpl, target, `${COLLECTION}a.ics`, 'X', {
      create: true,
    }).catch((e: unknown) => e);
    expect((error as CalDavError).code).toBe('auth');
  });
});

describe('addresses', () => {
  test('relative and absolute forms resolve inside the collection', () => {
    expect(resolveInCollection(target, 'a.ics')).toBe(`${COLLECTION}a.ics`);
    expect(resolveInCollection(target, '/remote.php/dav/calendars/alice/work/a%20b.ics')).toBe(
      `${COLLECTION}a%20b.ics`,
    );
    expect(() => resolveInCollection(target, '/remote.php/dav/calendars/alice/work')).toThrow(
      CalDavError,
    );
  });
});
