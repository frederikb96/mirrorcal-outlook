import { FALLBACK_TITLE, UID_HEX_LENGTH, UID_PREFIX } from './constants.ts';
import { sha256Hex } from './crypto.ts';
import { buildEventIcs } from './ics.ts';
import type { DesiredObject, MirrorEvent, Settings, SourceEvent, TextPolicy } from './types.ts';

/** `*` is the only wildcard; everything else matches literally, ignoring case. */
export function compileTitleFilter(pattern: string): RegExp | null {
  const trimmed = pattern.trim();
  if (trimmed === '') return null;
  const source = trimmed
    .split('*')
    .map((part) => part.replace(/[.+?^${}()|[\]\\]/g, '\\$&'))
    .join('.*');
  return new RegExp(`^${source}$`, 'isu');
}

export function compileTitleFilters(patterns: string[]): RegExp[] {
  return patterns.flatMap((p) => compileTitleFilter(p) ?? []);
}

export function isExcluded(event: SourceEvent, settings: Settings, filters: RegExp[]): boolean {
  if (settings.excludeShowAs.includes(event.showAs)) return true;
  const title = event.subject.trim();
  return filters.some((filter) => filter.test(title));
}

function cleanText(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  const text = value.replace(/\r\n?/g, '\n').trim();
  return text === '' ? undefined : text;
}

function applyPolicy(value: string | undefined, policy: TextPolicy): string | undefined {
  if (policy.mode === 'drop') return undefined;
  if (policy.mode === 'replace') return cleanText(policy.text);
  return cleanText(value);
}

function addDay(date: string): string {
  const [y, m, d] = date.split('-').map(Number) as [number, number, number];
  return new Date(Date.UTC(y, m - 1, d + 1)).toISOString().slice(0, 10);
}

export function mapEvent(event: SourceEvent, settings: Settings): MirrorEvent {
  const { policies } = settings;
  const summary = applyPolicy(event.subject, policies.title)?.replace(/\s*\n\s*/g, ' ');
  let start = event.start;
  let end = event.end;
  if (event.allDay) {
    start = start.slice(0, 10);
    end = end.slice(0, 10);
    if (end <= start) end = addDay(start);
  }
  const mapped: MirrorEvent = {
    summary: summary ?? FALLBACK_TITLE,
    allDay: event.allDay,
    start,
    end,
    showAs: event.showAs,
    attendees: [],
  };
  const description = applyPolicy(event.body, policies.description);
  if (description) mapped.description = description;
  const location = applyPolicy(event.location, policies.location)?.replace(/\s*\n\s*/g, ' ');
  if (location) mapped.location = location;
  if (policies.attendees === 'copy') {
    mapped.attendees = event.attendees ?? [];
    if (event.organizer) mapped.organizer = event.organizer;
  }
  return mapped;
}

/** Stable identity of a mirrored instance: derived from the owner and the Outlook instance id. */
export async function mirrorUid(ownerId: string, sourceId: string): Promise<string> {
  const hash = await sha256Hex(`${ownerId}|${sourceId}`);
  return `${UID_PREFIX}${hash.slice(0, UID_HEX_LENGTH)}`;
}

export interface DesiredSet {
  desired: DesiredObject[];
  /** Events dropped by the title filters or availability exclusions. */
  filtered: number;
  /** Instances sharing an id with an earlier one; only the first is mirrored. */
  duplicates: number;
}

/** Filters, applies policies and renders every source event to the ICS that should exist. */
export async function buildDesired(
  events: SourceEvent[],
  settings: Settings,
  collectionUrl: string,
  now: Date,
): Promise<DesiredSet> {
  const filters = compileTitleFilters(settings.titleFilters);
  const byUid = new Map<string, DesiredObject>();
  let filtered = 0;
  let duplicates = 0;
  for (const event of events) {
    if (isExcluded(event, settings, filters)) {
      filtered += 1;
      continue;
    }
    const uid = await mirrorUid(settings.ownerId, event.id);
    if (byUid.has(uid)) {
      duplicates += 1;
      continue;
    }
    const mirror = mapEvent(event, settings);
    const { ics, hash } = await buildEventIcs(mirror, { uid, owner: settings.ownerId, now });
    byUid.set(uid, {
      uid,
      href: new URL(`${encodeURIComponent(uid)}.ics`, collectionUrl).toString(),
      ics,
      hash,
      summary: mirror.summary,
      start: mirror.start,
      allDay: mirror.allDay,
    });
  }
  const desired = [...byUid.values()].sort((a, b) => a.uid.localeCompare(b.uid));
  return { desired, filtered, duplicates };
}
