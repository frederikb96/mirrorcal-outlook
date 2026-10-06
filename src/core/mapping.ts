import { FALLBACK_TITLE, UID_HEX_LENGTH, UID_PREFIX } from './constants.ts';
import { sha256Hex } from './crypto.ts';
import { buildEventIcs } from './ics.ts';
import type {
  Classification,
  DesiredObject,
  MirrorEvent,
  Sensitivity,
  Settings,
  SourceEvent,
  TextPolicy,
} from './types.ts';

export type TitleMatcher = (title: string) => boolean;

/** Case-folds one code point at a time, so context-sensitive forms (final sigma) fold like the rest. */
const foldCodePoint = (char: string) => (char === 'ς' ? 'σ' : char.toLowerCase());
const codePoints = (text: string) => Array.from(text, foldCodePoint);

/** Glob match, `*` being the only wildcard. Backtracks to the last `*` only, so the work is bounded by title length times pattern length. */
function globMatches(pattern: string[], text: string[]): boolean {
  let p = 0;
  let t = 0;
  let star = -1;
  let resume = 0;
  while (t < text.length) {
    if (pattern[p] === '*') {
      star = p;
      resume = t;
      p += 1;
    } else if (p < pattern.length && pattern[p] === text[t]) {
      p += 1;
      t += 1;
    } else if (star >= 0) {
      p = star + 1;
      resume += 1;
      t = resume;
    } else {
      return false;
    }
  }
  while (pattern[p] === '*') p += 1;
  return p === pattern.length;
}

/** `*` is the only wildcard; everything else matches literally, ignoring case. */
export function compileTitleFilter(pattern: string): TitleMatcher | null {
  const trimmed = pattern.trim();
  if (trimmed === '') return null;
  const glob = codePoints(trimmed);
  return (title) => globMatches(glob, codePoints(title));
}

export function compileTitleFilters(patterns: string[]): TitleMatcher[] {
  return patterns.flatMap((p) => compileTitleFilter(p) ?? []);
}

/** Outlook's own label, mapped to what CLASS says; ordinary events have none. */
export function classificationOf(sensitivity: Sensitivity): Classification | undefined {
  if (sensitivity === 'Normal') return undefined;
  return sensitivity === 'Confidential' ? 'CONFIDENTIAL' : 'PRIVATE';
}

export function isExcluded(
  event: SourceEvent,
  settings: Settings,
  filters: TitleMatcher[],
): boolean {
  if (settings.excludeShowAs.includes(event.showAs)) return true;
  if (classificationOf(event.sensitivity) && settings.policies.privateEvents.mode === 'skip') {
    return true;
  }
  const title = event.subject.trim();
  return filters.some((matches) => matches(title));
}

function cleanText(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  const text = value.replace(/\r\n?/g, '\n').trim();
  return text === '' ? undefined : text;
}

/** Replacement text stands in for a value the source has; an empty field stays empty. */
function applyPolicy(value: string | undefined, policy: TextPolicy): string | undefined {
  if (policy.mode === 'drop') return undefined;
  const text = cleanText(value);
  if (policy.mode === 'replace') return text === undefined ? undefined : cleanText(policy.text);
  return text;
}

function addDay(date: string): string {
  const [y, m, d] = date.split('-').map(Number) as [number, number, number];
  return new Date(Date.UTC(y, m - 1, d + 1)).toISOString().slice(0, 10);
}

export function mapEvent(event: SourceEvent, settings: Settings): MirrorEvent {
  const { policies } = settings;
  const classification = classificationOf(event.sensitivity);
  const privatePolicy = policies.privateEvents;
  // A private event shown under a neutral title must not give itself away through other fields.
  const concealed = classification !== undefined && privatePolicy.mode !== 'copy';
  const titlePolicy: TextPolicy =
    concealed && privatePolicy.mode === 'replace' ? privatePolicy : policies.title;
  const summary = applyPolicy(event.subject, titlePolicy)?.replace(/\s*\n\s*/g, ' ');
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
  if (classification) mapped.classification = classification;
  if (concealed) return mapped;
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
