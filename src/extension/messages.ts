import type { Outcome, PlanView, ResetResult, SyncCounts } from '../core/engine.ts';

export type Trigger = 'alarm' | 'startup' | 'token' | 'manual';

/** Outcomes the shell can produce before the engine runs, next to the engine's own. */
export type RunOutcome = Outcome | 'no-token' | 'not-configured' | 'no-permission' | 'reset';

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
  | { type: 'reset'; mode: 'count' | 'apply'; confirmedCount?: number };

export interface StatusResponse {
  configured: boolean;
  enabled: boolean;
  token: { present: boolean; usable: boolean; expiresAt?: string };
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

export type Response = StatusResponse | SyncResponse | ResetResponse;
