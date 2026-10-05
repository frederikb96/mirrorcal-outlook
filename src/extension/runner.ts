import { KEEP_ALIVE_MS } from '../core/constants.ts';
import { sha256Hex } from '../core/crypto.ts';
import {
  type PlanView,
  type ResetResult,
  resetMirror,
  runSync,
  type SyncResult,
} from '../core/engine.ts';
import { originPattern, validateDestination } from '../core/settings.ts';
import { isTokenUsable } from '../core/token.ts';
import { showAttention } from './attention.ts';
import { nextBackoff } from './backoff.ts';
import { releaseAccount } from './capture.ts';
import type { RunOutcome, RunRecord, Trigger } from './messages.ts';
import { type Admission, createSingleFlight } from './single-flight.ts';
import { ignoredTokensNote } from './status.ts';
import {
  addAcknowledgedOwners,
  clearToken,
  loadAccount,
  loadAcknowledgedOwners,
  loadBackoff,
  loadIgnoredTokens,
  loadPasswordRejected,
  loadSettings,
  loadSyncState,
  loadToken,
  recordRun,
  saveAccount,
  saveBackoff,
  savePasswordRejected,
  saveRejectedToken,
  saveSyncState,
} from './storage.ts';

type SyncRun = { record: RunRecord; plan?: PlanView };

/** One flight for everything that touches the destination, so a reset never overlaps a sync. */
const flight = createSingleFlight<SyncRun | ResetResult>();

export const doFetch = (input: string, init?: RequestInit) => fetch(input, init);

const PASSWORD_REJECTED_MESSAGE =
  'The calendar server rejected the password: enter it again in the settings. Nothing is retried until the settings are saved';

/**
 * Calls an extension API now and then while `task` runs: a worker is only kept alive by events
 * and extension API calls, and a long first sync makes neither.
 */
export async function keepAliveWhile<T>(
  task: () => Promise<T>,
  intervalMs = KEEP_ALIVE_MS,
): Promise<T> {
  const timer = setInterval(
    () => void chrome.runtime.getPlatformInfo().catch(() => {}),
    intervalMs,
  );
  try {
    return await task();
  } finally {
    clearInterval(timer);
  }
}

const passwordRejected = (result: SyncResult) =>
  result.error?.source === 'caldav' && result.error.code === 'auth';

function recordOf(
  trigger: Trigger,
  startedAt: Date,
  outcome: RunOutcome,
  rest: Partial<RunRecord> = {},
): RunRecord {
  return {
    kind: 'sync',
    trigger,
    startedAt: startedAt.toISOString(),
    finishedAt: new Date().toISOString(),
    outcome,
    repeats: 1,
    ...rest,
  };
}

function recordFromResult(trigger: Trigger, startedAt: Date, result: SyncResult): RunRecord {
  const rejected = passwordRejected(result);
  const record = recordOf(trigger, startedAt, rejected ? 'credentials-rejected' : result.outcome, {
    counts: result.counts,
  });
  const message = rejected ? PASSWORD_REJECTED_MESSAGE : (result.message ?? result.error?.message);
  if (message) record.message = message;
  const failure = result.error
    ? `${result.error.code}${result.error.status === undefined ? '' : `:${result.error.status}`}`
    : result.errors[0]?.code;
  if (failure) record.failure = failure;
  if (result.guard) record.guard = { code: result.guard.code, message: result.guard.message };
  if (result.digest && (result.outcome === 'needs-confirmation' || result.outcome === 'preview')) {
    record.digest = result.digest;
  }
  return record;
}

export interface SyncRequest {
  trigger: Trigger;
  mode: 'apply' | 'preview';
  confirm?: { digest: string };
}

async function syncOnce(request: SyncRequest): Promise<{ record: RunRecord; plan?: PlanView }> {
  const startedAt = new Date();
  const preview = request.mode === 'preview';
  const finish = async (record: RunRecord, plan?: PlanView) => {
    // Previews and runs that were never attempted are not worth remembering.
    if (!preview && record.outcome !== 'skipped') {
      const latest = record.outcome === 'disabled' ? record : await recordRun(record);
      await showAttention(latest);
    }
    return plan ? { record, plan } : { record };
  };

  const { settings, stored } = await loadSettings();
  if (!stored || validateDestination(settings).length > 0) {
    return finish(
      recordOf(request.trigger, startedAt, 'not-configured', {
        message: 'Open the settings and enter the calendar URL, user name and password',
      }),
    );
  }
  if (request.trigger !== 'manual' && !preview) {
    const backoff = await loadBackoff();
    const waiting = (backoff?.retryAt ?? 0) > Date.now();
    if ((await loadPasswordRejected()) || waiting) {
      return finish(recordOf(request.trigger, startedAt, 'skipped'));
    }
  }
  const token = await loadToken();
  if (!isTokenUsable(token, Date.now())) {
    const ignored = await loadIgnoredTokens();
    return finish(
      recordOf(request.trigger, startedAt, 'no-token', {
        message: `No valid Outlook session seen: open Outlook in this browser${
          ignored > 0 ? `. ${ignoredTokensNote(ignored)}` : ''
        }`,
      }),
    );
  }
  const origin = originPattern(settings.caldavUrl) as string;
  if (!(await chrome.permissions.contains({ origins: [origin] }))) {
    return finish(
      recordOf(request.trigger, startedAt, 'no-permission', {
        message: 'The extension has no access to the calendar server: save the settings again',
      }),
    );
  }

  const lastState = preview ? undefined : await loadSyncState();
  const pinned = await loadAccount();
  const result = await runSync({
    fetch: doFetch,
    token: (token as { token: string }).token,
    ...(pinned ? { expectedAccount: pinned } : {}),
    acknowledgedOwners: await loadAcknowledgedOwners(),
    settings,
    now: new Date(),
    mode: request.mode,
    ...(lastState ? { lastState } : {}),
    ...(request.confirm
      ? { confirmMassChange: true, confirmedPlanDigest: request.confirm.digest }
      : {}),
  });
  if (result.outcome === 'auth-failed') {
    await saveRejectedToken(await sha256Hex((token as { token: string }).token));
    await clearToken();
  }
  const record = recordFromResult(request.trigger, startedAt, result);
  if (!preview) {
    if (result.outcome === 'applied') await saveSyncState(result.state);
    if (record.outcome === 'credentials-rejected') await savePasswordRejected(true);
    else if (result.outcome === 'applied') await savePasswordRejected(false);
    if (result.outcome === 'applied' && result.account && !pinned)
      await saveAccount(result.account);
    if (result.outcome === 'applied' && result.acknowledgedOwners) {
      await addAcknowledgedOwners(result.acknowledgedOwners);
    }
    await saveBackoff(
      nextBackoff(await loadBackoff(), record, Date.now(), settings.intervalMinutes),
    );
  }
  if (preview && result.guard)
    record.guard = { code: result.guard.code, message: result.guard.message };
  return finish(
    record,
    preview || result.outcome === 'needs-confirmation' ? result.plan : undefined,
  );
}

/** Scheduled runs coalesce into one follow-up; interactive ones report busy instead. */
export async function runSyncFlight(request: SyncRequest): Promise<Admission<SyncRun>> {
  const guarded = async (): Promise<SyncRun> => {
    try {
      return await keepAliveWhile(() => syncOnce(request));
    } catch (error) {
      // Nobody watches a scheduled run, so even an unexpected failure ends up in the status.
      const record = recordOf(request.trigger, new Date(), 'failed', {
        message: error instanceof Error ? error.message : 'Unexpected error',
      });
      if (request.mode === 'apply') await showAttention(await recordRun(record));
      return { record };
    }
  };
  const admitted = await flight.run(guarded, {
    coalesce: request.mode === 'apply' && !request.confirm,
  });
  return admitted as Admission<SyncRun>;
}

export async function runResetFlight(
  mode: 'count' | 'apply',
  confirmedCount?: number,
  ownerId?: string,
): Promise<Admission<ResetResult>> {
  const admitted = await flight.run(() =>
    keepAliveWhile(async () => {
      const { settings } = await loadSettings();
      const startedAt = new Date();
      const result = await resetMirror({
        fetch: doFetch,
        settings,
        mode,
        ...(confirmedCount === undefined ? {} : { confirmedCount }),
        ...(ownerId === undefined ? {} : { ownerId }),
      });
      if (mode === 'apply' && result.outcome === 'reset') {
        const other = ownerId !== undefined && ownerId !== settings.ownerId;
        if (!other) await releaseAccount();
        await showAttention(
          await recordRun({
            kind: 'reset',
            trigger: 'manual',
            startedAt: startedAt.toISOString(),
            finishedAt: new Date().toISOString(),
            outcome: 'reset',
            message: other
              ? `Deleted ${result.deleted} events of owner id ${ownerId}`
              : `Deleted ${result.deleted} mirrored events`,
            repeats: 1,
          }),
        );
      }
      return result;
    }),
  );
  return admitted as Admission<ResetResult>;
}
