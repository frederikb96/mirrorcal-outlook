import type { FetchLike } from '../../src/core/types.ts';

export interface OutlookRequest {
  method: string;
  url: string;
  headers: Record<string, string>;
}

export type OutlookFault =
  | { kind: 'status'; status: number; onPage?: number }
  | { kind: 'body'; body: string; onPage?: number }
  | { kind: 'network'; onPage?: number }
  | { kind: 'next-link'; link: string; onPage?: number };

/** Serves a calendar view over an event list: window overlap, paging, optional faults. */
export class FakeOutlook {
  events: Record<string, unknown>[] = [];
  expectedToken: string;
  pageSize = 3;
  fault?: OutlookFault;
  requests: OutlookRequest[] = [];

  constructor(token: string, events: Record<string, unknown>[] = []) {
    this.expectedToken = token;
    this.events = events;
  }

  fetch: FetchLike = async (input, init) => {
    const url = new URL(input);
    const headers: Record<string, string> = {};
    for (const [k, v] of Object.entries((init?.headers ?? {}) as Record<string, string>)) {
      headers[k.toLowerCase()] = v;
    }
    this.requests.push({ method: init?.method ?? 'GET', url: input, headers });
    const skip = Number(url.searchParams.get('$skip') ?? 0);
    const page = skip / this.pageSize;
    const fault = this.fault;
    if (fault && (fault.onPage === undefined || fault.onPage === page)) {
      if (fault.kind === 'network') throw new TypeError('network down');
      if (fault.kind === 'status') return new Response('{}', { status: fault.status });
      if (fault.kind === 'body') return new Response(fault.body, { status: 200 });
    }
    if (headers.authorization !== `Bearer ${this.expectedToken}`) {
      return new Response('{}', { status: 401 });
    }
    const start = url.searchParams.get('startDateTime') ?? '';
    const end = url.searchParams.get('endDateTime') ?? '';
    const inWindow = this.events.filter((e) => {
      const s = (e.Start as { DateTime?: string } | undefined)?.DateTime;
      const en = (e.End as { DateTime?: string } | undefined)?.DateTime;
      // Items without usable times are served as they are, so tests can feed broken ones.
      if (!s || !en || !/^\d{4}-/.test(s)) return true;
      return (
        `${s.slice(0, 19)}Z` < `${end.slice(0, 19)}Z` &&
        `${en.slice(0, 19)}Z` > `${start.slice(0, 19)}Z`
      );
    });
    const items = inWindow.slice(skip, skip + this.pageSize);
    const body: Record<string, unknown> = { value: items };
    if (skip + this.pageSize < inWindow.length) {
      const next = new URL(input);
      next.searchParams.set('$skip', String(skip + this.pageSize));
      body['@odata.nextLink'] =
        fault?.kind === 'next-link' && (fault.onPage === undefined || fault.onPage === page)
          ? fault.link
          : next.toString();
    }
    return new Response(JSON.stringify(body), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  };
}
