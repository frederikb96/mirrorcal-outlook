import type { Outcome, PlanView, ResetResult, SyncCounts } from '../core/engine.ts';
import type { Settings } from '../core/types.ts';

export type Trigger = 'alarm' | 'startup' | 'token' | 'manual';

/**
 * Outcomes the shell can produce besides the engine's own. `skipped` is a scheduled run that was
 * not attempted (rejected password, failure back-off); it is never logged.
 */
export type RunOutcome =
  | Outcome
  | 'no-token'
  | 'not-configured'
  | 'no-permission'
  | 'credentials-rejected'
  | 'skipped'
  | 'reset';

/** What is kept of a run: counts and messages only, never event titles. */
export interface RunRecord {
  kind: 'sync' | 'reset';
  trigger: Trigger;
  startedAt: string;
  finishedAt: string;
  outcome: RunOutcome;
  message?: string;
  counts?: SyncCounts;
  guard?: { code: string; message: string };
  digest?: string;
  /** How many identical consecutive runs this entry stands for. */
  repeats: number;
}

export type Request =
  | { type: 'status' }
  | { type: 'sync'; mode: 'apply' | 'preview'; confirm?: { digest: string } }
  | { type: 'reset'; mode: 'count' | 'apply'; confirmedCount?: number }
  | { type: 'save'; settings: Settings }
  | { type: 'forget-account' };

export interface StatusResponse {
  configured: boolean;
  enabled: boolean;
  token: { present: boolean; usable: boolean; expiresAt?: string };
  /** The account being mirrored; `pinned` once a sync has succeeded for it. */
  account?: { label: string; pinned: boolean };
  /** The calendar server rejected the password and nothing is retried until settings are saved. */
  passwordRejected: boolean;
  /** ISO time before which scheduled runs do not retry a repeating failure. */
  retryAt?: string;
  lastRun?: RunRecord;
  lastSuccessAt?: string;
  log: RunRecord[];
}

export interface SyncResponse {
  busy: boolean;
  record?: RunRecord;
  plan?: PlanView;
}

export interface ResetResponse {
  busy: boolean;
  result?: ResetResult;
}

export interface SaveResponse {
  saved: boolean;
  /** Why nothing was saved. */
  errors?: string[];
  /** What the server said about the destination, shown next to the save confirmation. */
  connection?: string;
  /** The sync the save started. */
  sync?: SyncResponse;
}

export type Response = StatusResponse | SyncResponse | ResetResponse | SaveResponse;
