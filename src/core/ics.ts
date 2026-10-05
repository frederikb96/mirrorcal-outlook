import { FORMAT_VERSION, PRODID, PROP_HASH, PROP_OWNER } from './constants.ts';
import { sha256Hex } from './crypto.ts';
import type { MirrorEvent, ShowAs, SourceAttendee } from './types.ts';

const CRLF = '\r\n';
const MAX_LINE_OCTETS = 75;

function isControl(char: string): boolean {
  const code = char.codePointAt(0) as number;
  return code === 0x7f || (code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d);
}

/** Drops control characters other than tab and line breaks, which have meaning in ICS text. */
function stripControls(value: string): string {
  return Array.from(value)
    .filter((char) => !isControl(char))
    .join('');
}

export function escapeText(value: string): string {
  return stripControls(value)
    .replace(/\\/g, '\\\\')
    .replace(/;/g, String.raw`\;`)
    .replace(/,/g, '\\,')
    .replace(/\r\n|\r|\n/g, '\\n');
}

export function unescapeText(value: string): string {
  return value.replace(/\\([\\,;nN])/g, (_, c: string) => (c === 'n' || c === 'N' ? '\n' : c));
}

function octets(codePoint: number): number {
  if (codePoint < 0x80) return 1;
  if (codePoint < 0x800) return 2;
  if (codePoint < 0x10000) return 3;
  return 4;
}

/** Folds a content line so that no physical line exceeds 75 octets; never splits a character. */
export function foldLine(line: string): string {
  const physical: string[] = [];
  let current = '';
  let used = 0;
  for (const char of line) {
    const size = octets(char.codePointAt(0) as number);
    if (used + size > MAX_LINE_OCTETS) {
      physical.push(current);
      current = ' ';
      used = 1;
    }
    current += char;
    used += size;
  }
  physical.push(current);
  return physical.join(CRLF);
}

/** Joins folded physical lines back into content lines; tolerates bare LF line ends. */
export function unfoldLines(ics: string): string[] {
  return ics
    .replace(/\r?\n[ \t]/g, '')
    .split(/\r?\n/)
    .filter((line) => line !== '');
}

function utcStamp(iso: string): string {
  return `${iso.replace(/[-:]/g, '').slice(0, 15)}Z`;
}

function dateStamp(date: string): string {
  return date.replace(/-/g, '');
}

const BUSY_STATUS: Record<ShowAs, string> = {
  Free: 'FREE',
  Tentative: 'TENTATIVE',
  Busy: 'BUSY',
  Oof: 'OOF',
  WorkingElsewhere: 'WORKINGELSEWHERE',
  Unknown: 'BUSY',
};

const PARTSTAT: Record<string, string> = {
  Accepted: 'ACCEPTED',
  Organizer: 'ACCEPTED',
  Declined: 'DECLINED',
  TentativelyAccepted: 'TENTATIVE',
};

const ROLE: Record<string, string> = {
  Required: 'REQ-PARTICIPANT',
  Optional: 'OPT-PARTICIPANT',
  Resource: 'NON-PARTICIPANT',
};

function quoted(value: string): string {
  return `"${stripControls(value).replace(/["\r\n]/g, '')}"`;
}

function mailto(address: string): string {
  const safe = Array.from(address).filter((char) => !/[\s<>"\\]/.test(char) && !isControl(char));
  return `mailto:${safe.join('')}`;
}

/** Scheduling stays with the client so the server never sends invitations for mirrored people. */
function personParams(name: string): string {
  return `${name ? `;CN=${quoted(name)}` : ''};SCHEDULE-AGENT=CLIENT`;
}

function attendeeLine(attendee: SourceAttendee): string {
  const role = ROLE[attendee.type] ?? 'REQ-PARTICIPANT';
  const partstat = PARTSTAT[attendee.response] ?? 'NEEDS-ACTION';
  const cutype = attendee.type === 'Resource' ? ';CUTYPE=RESOURCE' : '';
  return `ATTENDEE${cutype};ROLE=${role};PARTSTAT=${partstat}${personParams(attendee.name)}:${mailto(attendee.address)}`;
}

export interface IcsMeta {
  uid: string;
  owner: string;
  /** Written as DTSTAMP; not part of the content hash. */
  now: Date;
}

function eventLines(event: MirrorEvent, meta: IcsMeta): string[] {
  const lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    `PRODID:${PRODID}`,
    'CALSCALE:GREGORIAN',
    'BEGIN:VEVENT',
    `UID:${meta.uid}`,
    `DTSTAMP:${utcStamp(meta.now.toISOString())}`,
  ];
  if (event.allDay) {
    lines.push(`DTSTART;VALUE=DATE:${dateStamp(event.start)}`);
    lines.push(`DTEND;VALUE=DATE:${dateStamp(event.end)}`);
  } else {
    lines.push(`DTSTART:${utcStamp(event.start)}`);
    lines.push(`DTEND:${utcStamp(event.end)}`);
  }
  lines.push(`SUMMARY:${escapeText(event.summary)}`);
  if (event.description) lines.push(`DESCRIPTION:${escapeText(event.description)}`);
  if (event.location) lines.push(`LOCATION:${escapeText(event.location)}`);
  lines.push(`TRANSP:${event.showAs === 'Free' ? 'TRANSPARENT' : 'OPAQUE'}`);
  if (event.showAs === 'Tentative') lines.push('STATUS:TENTATIVE');
  lines.push(`X-MICROSOFT-CDO-BUSYSTATUS:${BUSY_STATUS[event.showAs]}`);
  if (event.organizer) {
    lines.push(`ORGANIZER${personParams(event.organizer.name)}:${mailto(event.organizer.address)}`);
  }
  for (const attendee of event.attendees) lines.push(attendeeLine(attendee));
  lines.push(`${PROP_OWNER}:${meta.owner}`);
  return lines;
}

function render(lines: string[]): string {
  return lines.map((line) => `${foldLine(line)}${CRLF}`).join('');
}

/** Builds the ICS for one mirrored event, stamped with owner and content hash. */
export async function buildEventIcs(
  event: MirrorEvent,
  meta: IcsMeta,
): Promise<{ ics: string; hash: string }> {
  const lines = eventLines(event, meta);
  const hash = await contentHash(render([...lines, 'END:VEVENT', 'END:VCALENDAR']));
  const ics = render([...lines, `${PROP_HASH}:${hash}`, 'END:VEVENT', 'END:VCALENDAR']);
  return { ics, hash };
}

interface ContentLine {
  name: string;
  params: string[];
  value: string;
}

/** Splits at the first character of `separator` that is outside double quotes. */
function splitOutsideQuotes(text: string, separator: string, limit = Number.POSITIVE_INFINITY) {
  const parts: string[] = [];
  let quotes = false;
  let start = 0;
  for (let i = 0; i < text.length && parts.length < limit - 1; i++) {
    const char = text[i];
    if (char === '"') quotes = !quotes;
    else if (char === separator && !quotes) {
      parts.push(text.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(text.slice(start));
  return parts;
}

function parseLine(line: string): ContentLine | null {
  const [head, value] = splitOutsideQuotes(line, ':', 2);
  if (value === undefined || head === undefined) return null;
  const [name, ...params] = splitOutsideQuotes(head, ';');
  if (!name) return null;
  return { name: name.toUpperCase(), params, value };
}

/** Properties a server or client may legitimately rewrite without the content having changed. */
const VOLATILE = new Set(['DTSTAMP', 'CREATED', 'LAST-MODIFIED', 'SEQUENCE', PROP_HASH]);

function canonicalLines(ics: string): string[] {
  const lines: string[] = [];
  const stack: string[] = [];
  for (const raw of unfoldLines(ics)) {
    const line = parseLine(raw);
    if (!line) continue;
    if (line.name === 'BEGIN') stack.push(line.value.toUpperCase());
    const inEvent = stack.includes('VEVENT');
    if (inEvent && !VOLATILE.has(line.name)) {
      const params = line.params
        .map((p) => {
          const eq = p.indexOf('=');
          return eq < 0 ? p.toUpperCase() : `${p.slice(0, eq).toUpperCase()}${p.slice(eq)}`;
        })
        .sort();
      lines.push(`${[line.name, ...params].join(';')}:${line.value}`);
    }
    if (line.name === 'END') stack.pop();
  }
  return lines.sort();
}

/**
 * Hash of what an event says, independent of line folding, property order and server-maintained
 * stamps. The same function reads freshly built ICS and ICS fetched from the server, so a hand
 * edit on the server changes the hash exactly like a source edit does.
 */
export async function contentHash(ics: string): Promise<string> {
  return sha256Hex([`v${FORMAT_VERSION}`, ...canonicalLines(ics)].join('\n'));
}

export interface RemoteInfo {
  uid?: string;
  owner?: string;
  summary?: string;
  start?: string;
}

/** Reads identity and a display label from the first event of a server-side ICS body. */
export function readRemoteInfo(ics: string): RemoteInfo {
  const info: RemoteInfo = {};
  const stack: string[] = [];
  for (const raw of unfoldLines(ics)) {
    const line = parseLine(raw);
    if (!line) continue;
    if (line.name === 'BEGIN') stack.push(line.value.toUpperCase());
    else if (line.name === 'END') stack.pop();
    else if (stack.at(-1) === 'VEVENT') {
      if (line.name === 'UID') info.uid ??= line.value.trim();
      else if (line.name === PROP_OWNER) info.owner ??= line.value.trim();
      else if (line.name === 'SUMMARY') info.summary ??= unescapeText(line.value);
      else if (line.name === 'DTSTART') info.start ??= line.value.trim();
    }
  }
  return info;
}
