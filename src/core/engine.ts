/**
 * Sync engine. Runs anywhere `fetch`, `crypto.subtle` and `AbortSignal.timeout` exist (browser
 * service worker, Node, Bun); it touches no `chrome.*` API.
 *
 * Entry points:
 * - `runSync({ fetch, token, settings, now, mode })` reads the Outlook window and the whole
 *   destination calendar, plans the difference and, in `apply` mode, writes it.
 * - `resetMirror({ fetch, settings, mode })` counts or deletes every object carrying the owner
 *   marker.
 *
 * Both return plain JSON-serialisable results and never throw for expected failures; the
 * `outcome` field says what happened. Bundled for Node as `dist-engine/engine.mjs`
 * (`bun run build:engine`).
 */
import { applyPlan, type OpError } from './apply.ts';
import {
  CalDavError,
  type CalDavTarget,
  getCtag,
  inspectCollection,
  listObjects,
} from './caldav.ts';
import { evaluateGuards, type GuardVerdict } from './guards.ts';
import { buildDesired } from './mapping.ts';
import { fieldNeeds, OutlookReadError, readCalendarView } from './outlook.ts';
import {
  classifyRemote,
  desiredDigest,
  otherOwnersOf,
  type Plan,
  planDigest,
  planSync,
} from './plan.ts';
import { normalizeCaldavUrl, validateDestination, validateSettings } from './settings.ts';
import { normalizeToken, type TokenAccount, tokenAccount } from './token.ts';
import type { FetchLike, Settings } from './types.ts';

export type Outcome =
  | 'applied'
  | 'preview'
  | 'needs-confirmation'
  | 'disabled'
  | 'invalid-settings'
  | 'auth-failed'
  | 'read-failed'
  | 'destination-failed'
  | 'failed';

export interface RunOptions {
  fetch: FetchLike;
  /** Outlook bearer token, bare or as a full `Bearer …` header value. */
  token: string;
  /**
   * Id of the account this mirror is bound to (`SyncResult.account.id` of an earlier run). A token
   * issued to anyone else is refused before anything is read.
   */
  expectedAccount?: string;
  settings: Settings;
  now: Date;
  /** `preview` plans without writing anything. */
  mode: 'preview' | 'apply';
  /** Lets a plan through that the safety guards would otherwise refuse. */
  confirmMassChange?: boolean;
  /** When set, the confirmation only holds for the plan with this digest. */
  confirmedPlanDigest?: string;
  /** Include the generated ICS text in the plan entries. */
  includeIcs?: boolean;
  /**
   * State returned by the last quiet run. In apply mode, when the calendar's change tag and the
   * desired state still match it, the full destination read and the plan are skipped.
   */
  lastState?: SyncState;
}

/** What a run that wrote nothing and failed nowhere leaves behind for the next one. */
export interface SyncState {
  ctag: string;
  digest: string;
  own: number;
  foreign: number;
}

export interface PlanEntry {
  uid: string;
  summary: string;
  start: string;
  allDay?: boolean;
  href?: string;
  ics?: string;
}

export interface PlanView {
  creations: PlanEntry[];
  updates: PlanEntry[];
  deletions: PlanEntry[];
  conflicts: PlanEntry[];
  foreign: { href: string; uid?: string; summary?: string }[];
}

export interface SyncCounts {
  /** Items Outlook returned for the window, cancelled and filtered ones included. */
  source: number;
  /** Events that should exist in the destination. */
  mirrored: number;
  filtered: number;
  duplicates: number;
  /** Objects in the destination carrying the owner marker. */
  own: number;
  foreign: number;
  unreadable: number;
  create: number;
  update: number;
  delete: number;
  unchanged: number;
  conflicts: number;
  /** What was actually written in `apply` mode. */
  created: number;
  updated: number;
  deleted: number;
  errors: number;
}

export interface SyncResult {
  outcome: Outcome;
  message?: string;
  error?: { source: 'outlook' | 'caldav' | 'settings' | 'engine'; code: string; message: string };
  window?: { start: string; end: string };
  counts: SyncCounts;
  plan?: PlanView;
  /** Present when the safety guards refused the plan. */
  guard?: Exclude<GuardVerdict, { ok: true }>;
  /** Fingerprint of the plan; pass it back as `confirmedPlanDigest` together with `confirmMassChange`. */
  digest?: string;
  /** Whom the Outlook token was issued to, when it says; set once Outlook has been read. */
  account?: TokenAccount;
  /** True when the destination was found unchanged and not read. */
  skipped?: boolean;
  /** Present after a quiet, error-free apply run: pass it back as `lastState`. */
  state?: SyncState;
  errors: OpError[];
}

function emptyCounts(): SyncCounts {
  return {
    source: 0,
    mirrored: 0,
    filtered: 0,
    duplicates: 0,
    own: 0,
    foreign: 0,
    unreadable: 0,
    create: 0,
    update: 0,
    delete: 0,
    unchanged: 0,
    conflicts: 0,
    created: 0,
    updated: 0,
    deleted: 0,
    errors: 0,
  };
}

const DAY_MS = 86_400_000;

/** Whole UTC days around `now`, so the window only moves once a day. */
export function syncWindow(now: Date, daysBack: number, daysAhead: number) {
  const today = Math.floor(now.getTime() / DAY_MS) * DAY_MS;
  return {
    start: new Date(today - daysBack * DAY_MS),
    end: new Date(today + (daysAhead + 1) * DAY_MS),
  };
}

export function destinationTarget(settings: Settings): CalDavTarget {
  return {
    url: normalizeCaldavUrl(settings.caldavUrl),
    username: settings.username,
    password: settings.password,
  };
}

function view(plan: Plan, includeIcs: boolean): PlanView {
  const entry = (o: {
    uid: string;
    summary: string;
    start: string;
    allDay?: boolean;
    href: string;
    ics?: string;
  }): PlanEntry => {
    const e: PlanEntry = { uid: o.uid, summary: o.summary, start: o.start, href: o.href };
    if (o.allDay !== undefined) e.allDay = o.allDay;
    if (includeIcs && o.ics !== undefined) e.ics = o.ics;
    return e;
  };
  return {
    creations: plan.creations.map(entry),
    updates: plan.updates.map(entry),
    deletions: plan.deletions.map(entry),
    conflicts: plan.conflicts.map(entry),
    foreign: plan.foreign,
  };
}

function fail(
  outcome: Outcome,
  source: NonNullable<SyncResult['error']>['source'],
  code: string,
  message: string,
  counts = emptyCounts(),
): SyncResult {
  return { outcome, message, error: { source, code, message }, counts, errors: [] };
}

/** Compares a fresh Outlook read with a fresh destination read and, in apply mode, writes the difference. */
export async function runSync(options: RunOptions): Promise<SyncResult> {
  const { settings } = options;
  const problems = validateSettings(settings);
  if (problems.length > 0) {
    return fail('invalid-settings', 'settings', 'invalid', problems.join('; '));
  }
  if (options.mode === 'apply' && !settings.enabled) {
    return {
      outcome: 'disabled',
      message: 'Syncing is switched off',
      counts: emptyCounts(),
      errors: [],
    };
  }
  let token: string;
  try {
    token = normalizeToken(options.token);
  } catch {
    return fail('auth-failed', 'outlook', 'no-token', 'No usable Outlook token');
  }
  const account = tokenAccount(token);
  if (options.expectedAccount !== undefined && account?.id !== options.expectedAccount) {
    return fail(
      'auth-failed',
      'outlook',
      'wrong-account',
      'The Outlook session belongs to a different account than the mirrored one',
    );
  }

  const counts = emptyCounts();
  const target = destinationTarget(settings);
  const { start, end } = syncWindow(options.now, settings.daysBack, settings.daysAhead);
  const window = { start: start.toISOString(), end: end.toISOString() };

  try {
    // Both reads complete before anything is planned, let alone written.
    let read: Awaited<ReturnType<typeof readCalendarView>>;
    try {
      read = await readCalendarView({
        fetch: options.fetch,
        token,
        host: settings.outlookHost,
        start,
        end,
        needs: fieldNeeds(settings),
      });
    } catch (error) {
      if (error instanceof OutlookReadError) {
        const outcome = error.code === 'auth' ? 'auth-failed' : 'read-failed';
        return { ...fail(outcome, 'outlook', error.code, error.message), window };
      }
      throw error;
    }
    counts.source = read.rawCount;
    const accountField = account ? { account } : {};

    const { desired, filtered, duplicates } = await buildDesired(
      read.events,
      settings,
      target.url,
      options.now,
    );
    counts.mirrored = desired.length;
    counts.filtered = filtered;
    counts.duplicates = duplicates;

    const digestOfDesired = await desiredDigest(desired, target.url);
    // One small request tells whether anything in the calendar moved since the last quiet run.
    let ctag: string | undefined;
    if (options.mode === 'apply') {
      try {
        ctag = await getCtag(options.fetch, target);
      } catch (error) {
        if (error instanceof CalDavError) {
          return {
            ...fail('destination-failed', 'caldav', error.code, error.message, counts),
            window,
            ...accountField,
          };
        }
        throw error;
      }
    }
    const last = options.lastState;
    if (ctag && last && last.ctag === ctag && last.digest === digestOfDesired) {
      counts.own = last.own;
      counts.foreign = last.foreign;
      counts.unchanged = desired.length;
      return {
        outcome: 'applied',
        message: 'Nothing changed since the last sync',
        window,
        counts,
        skipped: true,
        state: last,
        ...accountField,
        errors: [],
      };
    }

    let listing: Awaited<ReturnType<typeof listObjects>>;
    try {
      listing = await listObjects(options.fetch, target);
    } catch (error) {
      if (error instanceof CalDavError) {
        return {
          ...fail('destination-failed', 'caldav', error.code, error.message, counts),
          window,
          ...accountField,
        };
      }
      throw error;
    }
    counts.unreadable = listing.unreadable;

    const plan = await planSync(desired, listing.objects, settings.ownerId);
    counts.own = plan.ownCount;
    counts.foreign = plan.foreign.length;
    counts.create = plan.creations.length;
    counts.update = plan.updates.length;
    counts.delete = plan.deletions.length;
    counts.unchanged = plan.unchanged;
    counts.conflicts = plan.conflicts.length;

    const digest = await planDigest(plan);
    const verdict = evaluateGuards({
      sourceCount: read.rawCount,
      deletions: plan.deletions.length,
      ownCount: plan.ownCount,
      creations: plan.creations.length,
      foreignCount: plan.foreign.length,
      otherOwners: otherOwnersOf(plan.foreign),
    });
    const result: SyncResult = {
      outcome: 'preview',
      window,
      ...accountField,
      counts,
      plan: view(plan, options.includeIcs === true),
      digest,
      errors: [],
    };
    if (!verdict.ok) result.guard = verdict;
    if (options.mode === 'preview') return result;

    const confirmed =
      options.confirmMassChange === true &&
      (options.confirmedPlanDigest === undefined || options.confirmedPlanDigest === digest);
    if (!verdict.ok && !confirmed) {
      const stale = options.confirmMassChange === true && !confirmed;
      result.outcome = 'needs-confirmation';
      result.message = stale
        ? `${verdict.message}; the plan changed since it was confirmed`
        : verdict.message;
      return result;
    }

    const ownedUrls = new Set(plan.ownHrefs);
    const applied = await applyPlan(plan, { fetch: options.fetch, target, ownedUrls });
    counts.created = applied.created;
    counts.updated = applied.updated;
    counts.deleted = applied.deleted;
    counts.errors = applied.errors.length;
    result.outcome = 'applied';
    result.errors = applied.errors;
    const quiet = applied.created + applied.updated + applied.deleted === 0;
    if (ctag && quiet && applied.errors.length === 0 && !applied.stopped) {
      result.state = {
        ctag,
        digest: digestOfDesired,
        own: plan.ownCount,
        foreign: plan.foreign.length,
      };
    }
    const first = applied.errors[0];
    if (applied.stopped) {
      result.message = applied.stopped;
      result.error = { source: 'caldav', code: 'auth', message: applied.stopped };
    } else if (first) {
      result.message = first.message;
    }
    return result;
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unexpected error';
    return { ...fail('failed', 'engine', 'unexpected', message, counts), window };
  }
}

export interface CheckResult {
  outcome: 'ok' | 'not-a-calendar' | 'auth-failed' | 'unreachable' | 'failed' | 'invalid-settings';
  message: string;
  /** The calendar's display name, when the server gives one. */
  name?: string;
}

/** Asks the server what the configured URL is, without writing anything. */
export async function checkDestination(options: {
  fetch: FetchLike;
  settings: Settings;
}): Promise<CheckResult> {
  const problems = validateDestination(options.settings);
  if (problems.length > 0) {
    return { outcome: 'invalid-settings', message: problems.join('; ') };
  }
  try {
    const info = await inspectCollection(options.fetch, destinationTarget(options.settings));
    if (!info.isCalendar) {
      return { outcome: 'not-a-calendar', message: 'This URL is not a calendar collection' };
    }
    if (!info.holdsEvents) {
      return { outcome: 'not-a-calendar', message: 'This calendar does not hold events' };
    }
    return {
      outcome: 'ok',
      message: info.name ? `Connected to calendar '${info.name}'` : 'Connected to the calendar',
      ...(info.name ? { name: info.name } : {}),
    };
  } catch (error) {
    if (error instanceof CalDavError) {
      const outcome =
        error.code === 'auth' ? 'auth-failed' : error.code === 'network' ? 'unreachable' : 'failed';
      return { outcome, message: error.message };
    }
    return { outcome: 'failed', message: error instanceof Error ? error.message : 'Failed' };
  }
}

export interface ResetOptions {
  fetch: FetchLike;
  settings: Settings;
  /** `count` only reports; `apply` deletes. */
  mode: 'count' | 'apply';
  /** Required for `apply`: the count the user was shown. A different current count refuses the reset. */
  confirmedCount?: number;
}

export interface ResetResult {
  outcome: 'counted' | 'reset' | 'refused' | 'invalid-settings' | 'destination-failed' | 'failed';
  message?: string;
  own: number;
  foreign: number;
  deleted: number;
  errors: OpError[];
}

/** Deletes every object carrying the owner marker, and nothing else. */
export async function resetMirror(options: ResetOptions): Promise<ResetResult> {
  const result: ResetResult = { outcome: 'counted', own: 0, foreign: 0, deleted: 0, errors: [] };
  const problems = validateDestination(options.settings);
  if (problems.length > 0)
    return { ...result, outcome: 'invalid-settings', message: problems.join('; ') };
  const target = destinationTarget(options.settings);
  try {
    const listing = await listObjects(options.fetch, target);
    const { own, foreign } = await classifyRemote(listing.objects, options.settings.ownerId);
    result.own = own.length;
    result.foreign = foreign.length;
    if (options.mode === 'count') return result;
    if (options.confirmedCount !== own.length) {
      return {
        ...result,
        outcome: 'refused',
        message: `The calendar now holds ${own.length} mirrored events, not ${options.confirmedCount ?? 'the confirmed number'}`,
      };
    }
    const plan: Plan = {
      creations: [],
      updates: [],
      conflicts: [],
      unchanged: 0,
      ownCount: own.length,
      ownHrefs: own.map((o) => o.href),
      foreign,
      deletions: own.map((o) => ({
        kind: 'delete',
        uid: o.uid,
        href: o.href,
        etag: o.etag,
        summary: o.summary,
        start: o.start,
      })),
    };
    const ownedUrls = new Set(plan.ownHrefs);
    const applied = await applyPlan(plan, { fetch: options.fetch, target, ownedUrls });
    return { ...result, outcome: 'reset', deleted: applied.deleted, errors: applied.errors };
  } catch (error) {
    if (error instanceof CalDavError) {
      return { ...result, outcome: 'destination-failed', message: error.message };
    }
    const message = error instanceof Error ? error.message : 'Unexpected error';
    return { ...result, outcome: 'failed', message };
  }
}
