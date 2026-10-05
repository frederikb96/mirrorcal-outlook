import type { SyncState } from '../core/engine.ts';
import { mergeWithDefaults } from '../core/settings.ts';
import type { CapturedToken } from '../core/token.ts';
import type { Settings } from '../core/types.ts';
import type { RunRecord } from './messages.ts';

const SETTINGS_KEY = 'settings';
const STATUS_KEY = 'status';
const TOKEN_KEY = 'outlookToken';
const STATE_KEY = 'syncState';
const LOG_LIMIT = 50;

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

export async function saveSettings(settings: Settings): Promise<void> {
  await chrome.storage.local.set({ [SETTINGS_KEY]: settings });
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

export async function recordRun(record: RunRecord): Promise<void> {
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
