import type { SyncState } from '../core/engine.ts';
import { mergeWithDefaults } from '../core/settings.ts';
import type { CapturedToken, TokenAccount } from '../core/token.ts';
import type { Settings } from '../core/types.ts';
import type { RunRecord } from './messages.ts';

const SETTINGS_KEY = 'settings';
const STATUS_KEY = 'status';
const TOKEN_KEY = 'outlookToken';
const STATE_KEY = 'syncState';
const ACCOUNT_KEY = 'account';
const REJECTED_PASSWORD_KEY = 'credentialsRejected';
const BACKOFF_KEY = 'backoff';
const REJECTED_TOKEN_KEY = 'rejectedToken';
const LOG_LIMIT = 50;

/** Identical failures in a row, and when scheduled runs may try again. */
export interface Backoff {
  signature: string;
  count: number;
  /** Epoch milliseconds; absent while the failure has not repeated often enough to back off. */
  retryAt?: number;
}

interface StoredStatus {
  lastRun?: RunRecord;
  lastSuccessAt?: string;
  log: RunRecord[];
}

/** Settings and credentials persist in storage.local; see the README for what that means. */
export async function loadSettings(): Promise<{ settings: Settings; stored: boolean }> {
  const items = await chrome.storage.local.get(SETTINGS_KEY);
  return { settings: mergeWithDefaults(items[SETTINGS_KEY]), stored: SETTINGS_KEY in items };
}

/** Saving settings is the user's answer to a rejected password or a failure loop: both reset. */
export async function saveSettings(settings: Settings): Promise<void> {
  await chrome.storage.local.set({ [SETTINGS_KEY]: settings });
  await chrome.storage.local.remove([REJECTED_PASSWORD_KEY, BACKOFF_KEY]);
}

/** The account the mirror is bound to, set by the first successful sync. */
export async function loadAccount(): Promise<TokenAccount | undefined> {
  const items = await chrome.storage.local.get(ACCOUNT_KEY);
  return items[ACCOUNT_KEY] as TokenAccount | undefined;
}

export async function saveAccount(account: TokenAccount): Promise<void> {
  await chrome.storage.local.set({ [ACCOUNT_KEY]: account });
}

export async function clearAccount(): Promise<void> {
  await chrome.storage.local.remove(ACCOUNT_KEY);
}

/** Set when the calendar server rejected the password; scheduled runs stop until settings are saved. */
export async function loadPasswordRejected(): Promise<boolean> {
  const items = await chrome.storage.local.get(REJECTED_PASSWORD_KEY);
  return REJECTED_PASSWORD_KEY in items;
}

export async function savePasswordRejected(rejected: boolean): Promise<void> {
  if (rejected) await chrome.storage.local.set({ [REJECTED_PASSWORD_KEY]: true });
  else await chrome.storage.local.remove(REJECTED_PASSWORD_KEY);
}

export async function loadBackoff(): Promise<Backoff | undefined> {
  const items = await chrome.storage.local.get(BACKOFF_KEY);
  return items[BACKOFF_KEY] as Backoff | undefined;
}

export async function saveBackoff(backoff: Backoff | undefined): Promise<void> {
  if (backoff) await chrome.storage.local.set({ [BACKOFF_KEY]: backoff });
  else await chrome.storage.local.remove(BACKOFF_KEY);
}

/** The observed Outlook token lives in memory-backed session storage and never in storage.local. */
export async function loadToken(): Promise<CapturedToken | null> {
  const items = await chrome.storage.session.get(TOKEN_KEY);
  return (items[TOKEN_KEY] as CapturedToken | undefined) ?? null;
}

export async function saveToken(token: CapturedToken): Promise<void> {
  await chrome.storage.session.set({ [TOKEN_KEY]: token });
}

export async function clearToken(): Promise<void> {
  await chrome.storage.session.remove(TOKEN_KEY);
}

/** Fingerprint of the last token Outlook rejected, so the page re-sending it is not taken up again. */
export async function loadRejectedToken(): Promise<string | undefined> {
  const items = await chrome.storage.session.get(REJECTED_TOKEN_KEY);
  return items[REJECTED_TOKEN_KEY] as string | undefined;
}

export async function saveRejectedToken(fingerprint: string): Promise<void> {
  await chrome.storage.session.set({ [REJECTED_TOKEN_KEY]: fingerprint });
}

export async function loadStatus(): Promise<StoredStatus> {
  const items = await chrome.storage.local.get(STATUS_KEY);
  const stored = items[STATUS_KEY] as StoredStatus | undefined;
  return { log: [], ...stored };
}

/** Two runs that said and did the same thing are one log entry with a repeat count. */
function sameOutcome(a: RunRecord, b: RunRecord): boolean {
  return (
    a.kind === b.kind &&
    a.outcome === b.outcome &&
    a.message === b.message &&
    a.digest === b.digest &&
    (a.counts?.errors ?? 0) === (b.counts?.errors ?? 0)
  );
}

/** Appends a run to the log and returns the log's newest entry, repeat count included. */
export async function recordRun(record: RunRecord): Promise<RunRecord> {
  const status = await loadStatus();
  const [latest, ...older] = status.log;
  const quiet = (r: RunRecord) =>
    !r.counts || r.counts.created + r.counts.updated + r.counts.deleted === 0;
  if (latest && quiet(latest) && quiet(record) && sameOutcome(latest, record)) {
    status.log = [
      { ...latest, finishedAt: record.finishedAt, repeats: latest.repeats + 1 },
      ...older,
    ];
  } else {
    status.log = [record, ...status.log].slice(0, LOG_LIMIT);
  }
  status.lastRun = status.log[0] as RunRecord;
  if (record.outcome === 'applied' && !record.counts?.errors) {
    status.lastSuccessAt = record.finishedAt;
  }
  await chrome.storage.local.set({ [STATUS_KEY]: status });
  return status.lastRun;
}

/** Change tag and digest from the last quiet run; lets the next run skip reading the calendar. */
export async function loadSyncState(): Promise<SyncState | undefined> {
  const items = await chrome.storage.local.get(STATE_KEY);
  return items[STATE_KEY] as SyncState | undefined;
}

export async function saveSyncState(state: SyncState | undefined): Promise<void> {
  if (state) await chrome.storage.local.set({ [STATE_KEY]: state });
  else await chrome.storage.local.remove(STATE_KEY);
}
