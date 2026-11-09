/**
 * Sync engine. Runs anywhere `fetch`, `crypto.subtle` and `AbortSignal.timeout` exist (browser
 * service worker, Node, Bun); it touches no `chrome.*` API.
 *
 * Entry points:
 * - `runSync({ fetch, token, settings, now, mode })` reads the Outlook window and every
 *   destination calendar in full (the main one and, when configured, the one for excluded
 *   events), plans the difference and, in `apply` mode, writes it.
 * - `resetMirror({ fetch, settings, mode })` counts or deletes every object carrying the owner
 *   marker, in every destination calendar.
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
import { sha256Hex } from './crypto.ts';
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
import {
  normalizeCaldavUrl,
  OWNER_ID_PATTERN,
  validateDestination,
  validateSettings,
} from './settings.ts';
import { matchAccount, normalizeToken, type TokenAccount, tokenAccount } from './token.ts';
import type { DesiredObject, FetchLike, Settings } from './types.ts';

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
   * The account this mirror is bound to (`SyncResult.account` of an earlier run). A token that is
   * provably issued to anyone else is refused before anything is read; one that carries nothing
   * comparable is let through.
   */
  expectedAccount?: TokenAccount;
  /**
   * Owner ids of other mirrors whose events the user already confirmed to share the calendar
   * with; they no longer hold back creations.
   */
  acknowledgedOwners?: string[];
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
  /** The same, for the calendar for excluded events. */
  lastExcludedState?: SyncState;
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
  error?: {
    source: 'outlook' | 'caldav' | 'settings' | 'engine';
    code: string;
    message: string;
    /** HTTP status, for failures that came with one. */
    status?: number;
  };
  window?: { start: string; end: string };
  counts: SyncCounts;
  plan?: PlanView;
  /** Present when the safety guards refused the plan. */
  guard?: Exclude<GuardVerdict, { ok: true }>;
  /** Fingerprint of the plan; pass it back as `confirmedPlanDigest` together with `confirmMassChange`. */
  digest?: string;
  /** Whom the Outlook token was issued to, when it says; set once Outlook has been read. */
  account?: TokenAccount;
  /** Present after a confirmed plan that shares the calendar with other mirrors' events. */
  acknowledgedOwners?: string[];
  /** True when the destination was found unchanged and not read. */
  skipped?: boolean;
  /** Present after a quiet, error-free apply run: pass it back as `lastState`. */
  state?: SyncState;
  /** The calendar for excluded events, when one is configured and the run got to it. */
  excluded?: ExcludedResult;
  errors: OpError[];
}

/** The second calendar's part of a run; `counts` covers only its own reads and writes. */
export interface ExcludedResult {
  counts: SyncCounts;
  plan?: PlanView;
  skipped?: boolean;
  /** Pass it back as `lastExcludedState`. */
  state?: SyncState;
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

/** The calendar for excluded events on the same server, with the same credentials. */
export function excludedTarget(settings: Settings): CalDavTarget | undefined {
  if (settings.excludedCaldavUrl === '') return undefined;
  return { ...destinationTarget(settings), url: normalizeCaldavUrl(settings.excludedCaldavUrl) };
}

/** Every calendar the settings write to, main first. */
function targetsOf(settings: Settings): { role: CalendarRole; target: CalDavTarget }[] {
  const second = excludedTarget(settings);
  return [
    { role: 'main', target: destinationTarget(settings) },
    ...(second ? [{ role: 'excluded' as const, target: second }] : []),
  ];
}

export type CalendarRole = 'main' | 'excluded';

/** Prefixes what is said about the second calendar, so a message names the calendar it is about. */
function about(role: CalendarRole, message: string): string {
  return role === 'main' ? message : `Calendar for excluded events: ${message}`;
}

const EMPTY_PLAN: Plan = {
  creations: [],
  updates: [],
  deletions: [],
  unchanged: 0,
  conflicts: [],
  ownCount: 0,
  ownHrefs: [],
  foreign: [],
};

/** One calendar's share of a run, from its read to its writes. */
interface CalendarPass {
  role: CalendarRole;
  target: CalDavTarget;
  desired: DesiredObject[];
  last: SyncState | undefined;
  digestOfDesired: string;
  counts: SyncCounts;
  ctag?: string;
  /** Found unchanged since the last quiet run and not read. */
  skipped: boolean;
  plan?: Plan;
  verdict: GuardVerdict;
  state?: SyncState;
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
  status?: number,
): SyncResult {
  const error = { source, code, message, ...(status === undefined ? {} : { status }) };
  return { outcome, message, error, counts, errors: [] };
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
  if (options.expectedAccount && matchAccount(options.expectedAccount, account) === 'mismatch') {
    return fail(
      'auth-failed',
      'outlook',
      'wrong-account',
      'The Outlook session belongs to a different account than the mirrored one',
    );
  }

  const { start, end } = syncWindow(options.now, settings.daysBack, settings.daysAhead);
  const window = { start: start.toISOString(), end: end.toISOString() };
  const targets = targetsOf(settings);
  const main = targets[0] as { role: CalendarRole; target: CalDavTarget };
  const second = targets[1];

  try {
    // Every read completes before anything is planned, let alone written.
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
    const accountField = account ? { account } : {};

    const built = await buildDesired(
      read.events,
      settings,
      main.target.url,
      options.now,
      second?.target.url,
    );
    const passes: CalendarPass[] = [];
    for (const { role, target } of targets) {
      const desired = role === 'main' ? built.desired : built.excluded;
      passes.push({
        role,
        target,
        desired,
        last: role === 'main' ? options.lastState : options.lastExcludedState,
        digestOfDesired: await desiredDigest(desired, target.url),
        counts: { ...emptyCounts(), source: read.rawCount, mirrored: desired.length },
        skipped: false,
        verdict: { ok: true },
      });
    }
    const mainPass = passes[0] as CalendarPass;
    mainPass.counts.filtered = built.filtered;
    mainPass.counts.duplicates = built.duplicates;

    const acknowledged = new Set(options.acknowledgedOwners);
    for (const pass of passes) {
      const failed = await readCalendar(pass, options, acknowledged);
      if (failed) {
        return {
          ...fail(
            'destination-failed',
            'caldav',
            failed.code,
            about(pass.role, failed.message),
            mainPass.counts,
            failed.status,
          ),
          window,
          ...accountField,
        };
      }
    }

    const result: SyncResult = {
      outcome: 'applied',
      window,
      ...accountField,
      counts: mainPass.counts,
      errors: [],
    };
    const excludedPass = passes[1];
    const attachExcluded = () => {
      if (!excludedPass) return;
      const part: ExcludedResult = { counts: excludedPass.counts };
      if (excludedPass.plan) part.plan = view(excludedPass.plan, options.includeIcs === true);
      if (excludedPass.skipped) part.skipped = true;
      if (excludedPass.state) part.state = excludedPass.state;
      result.excluded = part;
    };
    const keepState = () => {
      if (mainPass.state) result.state = mainPass.state;
      attachExcluded();
    };

    if (passes.every((pass) => pass.skipped)) {
      result.message = 'Nothing changed since the last sync';
      result.skipped = true;
      keepState();
      return result;
    }

    const digests = await Promise.all(passes.map((pass) => planDigest(pass.plan ?? EMPTY_PLAN)));
    const digest =
      digests.length === 1
        ? (digests[0] as string)
        : await sha256Hex(passes.map((pass, i) => `${pass.role} ${digests[i]}`).join('\n'));
    const held = passes.find((pass) => !pass.verdict.ok);
    result.outcome = 'preview';
    result.plan = view(mainPass.plan ?? EMPTY_PLAN, options.includeIcs === true);
    result.digest = digest;
    if (held && !held.verdict.ok) {
      result.guard = { ...held.verdict, message: about(held.role, held.verdict.message) };
    }
    attachExcluded();
    if (options.mode === 'preview') return result;

    const confirmed =
      options.confirmMassChange === true &&
      (options.confirmedPlanDigest === undefined || options.confirmedPlanDigest === digest);
    if (result.guard && !confirmed) {
      const stale = options.confirmMassChange === true && !confirmed;
      result.outcome = 'needs-confirmation';
      result.message = stale
        ? `${result.guard.message}; the plan changed since it was confirmed`
        : result.guard.message;
      return result;
    }

    const acknowledgedNow = passes.flatMap((pass) =>
      !pass.verdict.ok && pass.verdict.code === 'other-owner' && pass.plan
        ? otherOwnersOf(pass.plan.foreign)
        : [],
    );
    if (acknowledgedNow.length > 0) result.acknowledgedOwners = [...new Set(acknowledgedNow)];

    let stopped: string | undefined;
    for (const pass of passes) {
      if (!pass.plan) continue;
      if (stopped) break;
      const applied = await applyPlan(pass.plan, {
        fetch: options.fetch,
        target: pass.target,
        ownedUrls: new Set(pass.plan.ownHrefs),
      });
      pass.counts.created = applied.created;
      pass.counts.updated = applied.updated;
      pass.counts.deleted = applied.deleted;
      pass.counts.errors = applied.errors.length;
      result.errors.push(
        ...applied.errors.map((e) => ({ ...e, message: about(pass.role, e.message) })),
      );
      stopped = applied.stopped;
      const quiet = applied.created + applied.updated + applied.deleted === 0;
      if (pass.ctag && quiet && applied.errors.length === 0 && !applied.stopped) {
        pass.state = {
          ctag: pass.ctag,
          digest: pass.digestOfDesired,
          own: pass.plan.ownCount,
          foreign: pass.plan.foreign.length,
        };
      }
    }
    result.outcome = 'applied';
    keepState();
    const first = result.errors[0];
    if (stopped) {
      result.message = stopped;
      result.error = { source: 'caldav', code: 'auth', message: stopped, status: 401 };
    } else if (first) {
      result.message = first.message;
    }
    return result;
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unexpected error';
    return { ...fail('failed', 'engine', 'unexpected', message), window };
  }
}

/**
 * Reads one calendar and plans it: in apply mode its change tag first, and when that and the
 * desired state match the last quiet run, nothing more. Returns the failure that stops the run.
 */
async function readCalendar(
  pass: CalendarPass,
  options: RunOptions,
  acknowledged: Set<string>,
): Promise<CalDavError | undefined> {
  try {
    if (options.mode === 'apply') pass.ctag = await getCtag(options.fetch, pass.target);
    const last = pass.last;
    if (pass.ctag && last && last.ctag === pass.ctag && last.digest === pass.digestOfDesired) {
      pass.skipped = true;
      pass.state = last;
      pass.counts.own = last.own;
      pass.counts.foreign = last.foreign;
      pass.counts.unchanged = pass.desired.length;
      return undefined;
    }
    const listing = await listObjects(options.fetch, pass.target);
    const plan = await planSync(pass.desired, listing.objects, options.settings.ownerId);
    const { counts } = pass;
    counts.unreadable = listing.unreadable;
    counts.own = plan.ownCount;
    counts.foreign = plan.foreign.length;
    counts.create = plan.creations.length;
    counts.update = plan.updates.length;
    counts.delete = plan.deletions.length;
    counts.unchanged = plan.unchanged;
    counts.conflicts = plan.conflicts.length;
    pass.plan = plan;
    pass.verdict = evaluateGuards({
      sourceCount: counts.source,
      deletions: plan.deletions.length,
      ownCount: plan.ownCount,
      creations: plan.creations.length,
      foreignCount: plan.foreign.length,
      otherOwners: otherOwnersOf(plan.foreign).filter((id) => !acknowledged.has(id)),
    });
    return undefined;
  } catch (error) {
    if (error instanceof CalDavError) return error;
    throw error;
  }
}

export interface CheckResult {
  outcome: 'ok' | 'not-a-calendar' | 'auth-failed' | 'unreachable' | 'failed' | 'invalid-settings';
  message: string;
  /** The calendar's display name, when the server gives one. */
  name?: string;
}

/** Asks the server what the configured URLs are, without writing anything. */
export async function checkDestination(options: {
  fetch: FetchLike;
  settings: Settings;
}): Promise<CheckResult> {
  const problems = validateDestination(options.settings);
  if (problems.length > 0) {
    return { outcome: 'invalid-settings', message: problems.join('; ') };
  }
  const names: (string | undefined)[] = [];
  for (const { role, target } of targetsOf(options.settings)) {
    const checked = await checkCalendar(options.fetch, target);
    if ('outcome' in checked) return { ...checked, message: about(role, checked.message) };
    names.push(checked.name);
  }
  const [name, excludedName] = names;
  const calendar = (n: string | undefined) => (n ? `calendar '${n}'` : 'the calendar');
  const message =
    names.length === 1
      ? `Connected to ${calendar(name)}`
      : `Connected to ${calendar(name)}, and to ${calendar(excludedName)} for excluded events`;
  return { outcome: 'ok', message, ...(name ? { name } : {}) };
}

async function checkCalendar(
  fetchImpl: FetchLike,
  target: CalDavTarget,
): Promise<{ name?: string } | Omit<CheckResult, 'name'>> {
  try {
    const info = await inspectCollection(fetchImpl, target);
    if (!info.isCalendar) {
      return { outcome: 'not-a-calendar', message: 'This URL is not a calendar collection' };
    }
    if (!info.holdsEvents) {
      return { outcome: 'not-a-calendar', message: 'This calendar does not hold events' };
    }
    return info.name ? { name: info.name } : {};
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
  /**
   * Whose events to count and delete: the configured owner id by default, or the id of another
   * mirror. Only objects carrying exactly that owner marker are ever listed for deletion.
   */
  ownerId?: string;
}

export interface ResetResult {
  outcome: 'counted' | 'reset' | 'refused' | 'invalid-settings' | 'destination-failed' | 'failed';
  message?: string;
  /** Objects carrying the owner marker that was asked about. */
  own: number;
  foreign: number;
  /** Other mirrors' owner ids found among the foreign objects, with their object counts. */
  otherOwners: { id: string; count: number }[];
  deleted: number;
  errors: OpError[];
}

/**
 * Deletes every object carrying the chosen owner marker, and nothing else, in every configured
 * calendar. All of them are read before anything is deleted; counts are totals.
 */
export async function resetMirror(options: ResetOptions): Promise<ResetResult> {
  const result: ResetResult = {
    outcome: 'counted',
    own: 0,
    foreign: 0,
    otherOwners: [],
    deleted: 0,
    errors: [],
  };
  const problems = validateDestination(options.settings);
  const ownerId = options.ownerId ?? options.settings.ownerId;
  if (!OWNER_ID_PATTERN.test(ownerId)) problems.push('Owner id is not valid');
  if (problems.length > 0)
    return { ...result, outcome: 'invalid-settings', message: problems.join('; ') };
  const plans: { target: CalDavTarget; plan: Plan }[] = [];
  let role: CalendarRole = 'main';
  try {
    for (const calendar of targetsOf(options.settings)) {
      role = calendar.role;
      const listing = await listObjects(options.fetch, calendar.target);
      const { own, foreign } = await classifyRemote(listing.objects, ownerId);
      result.own += own.length;
      result.foreign += foreign.length;
      for (const object of foreign) {
        if (!object.owner) continue;
        const entry = result.otherOwners.find((o) => o.id === object.owner);
        if (entry) entry.count += 1;
        else result.otherOwners.push({ id: object.owner, count: 1 });
      }
      plans.push({
        target: calendar.target,
        plan: {
          ...EMPTY_PLAN,
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
        },
      });
    }
    if (options.mode === 'count') return result;
    if (options.confirmedCount !== result.own) {
      return {
        ...result,
        outcome: 'refused',
        message: `The calendars now hold ${result.own} events of this owner id, not ${options.confirmedCount ?? 'the confirmed number'}`,
      };
    }
    for (const { target, plan } of plans) {
      const applied = await applyPlan(plan, {
        fetch: options.fetch,
        target,
        ownedUrls: new Set(plan.ownHrefs),
      });
      result.deleted += applied.deleted;
      result.errors.push(...applied.errors);
      if (applied.stopped) break;
    }
    return { ...result, outcome: 'reset' };
  } catch (error) {
    if (error instanceof CalDavError) {
      return { ...result, outcome: 'destination-failed', message: about(role, error.message) };
    }
    const message = error instanceof Error ? error.message : 'Unexpected error';
    return { ...result, outcome: 'failed', message };
  }
}
