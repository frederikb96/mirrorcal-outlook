import {
  type PlanView,
  type ResetResult,
  resetMirror,
  runSync,
  type SyncResult,
} from '../core/engine.ts';
import { validateDestination } from '../core/settings.ts';
import { isTokenUsable } from '../core/token.ts';
import type { RunOutcome, RunRecord, Trigger } from './messages.ts';
import { type Admission, createSingleFlight } from './single-flight.ts';
import {
  clearToken,
  loadSettings,
  loadSyncState,
  loadToken,
  recordRun,
  saveSyncState,
} from './storage.ts';

type SyncRun = { record: RunRecord; plan?: PlanView };

/** One flight for everything that touches the destination, so a reset never overlaps a sync. */
const flight = createSingleFlight<SyncRun | ResetResult>();

const doFetch = (input: string, init?: RequestInit) => fetch(input, init);

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
  const record = recordOf(trigger, startedAt, result.outcome, { counts: result.counts });
  const message = result.message ?? result.error?.message;
  if (message) record.message = message;
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
    // Previews and idle switch-offs are not runs worth remembering.
    if (!preview && record.outcome !== 'disabled') await recordRun(record);
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
  const token = await loadToken();
  if (!isTokenUsable(token, Date.now())) {
    return finish(
      recordOf(request.trigger, startedAt, 'no-token', {
        message: 'No valid Outlook session seen: open Outlook in this browser',
      }),
    );
  }
  const origin = `${new URL(settings.caldavUrl).origin}/*`;
  if (!(await chrome.permissions.contains({ origins: [origin] }))) {
    return finish(
      recordOf(request.trigger, startedAt, 'no-permission', {
        message: 'The extension has no access to the calendar server: save the settings again',
      }),
    );
  }

  const lastState = preview ? undefined : await loadSyncState();
  const result = await runSync({
    fetch: doFetch,
    token: (token as { token: string }).token,
    settings,
    now: new Date(),
    mode: request.mode,
    ...(lastState ? { lastState } : {}),
    ...(request.confirm
      ? { confirmMassChange: true, confirmedPlanDigest: request.confirm.digest }
      : {}),
  });
  if (result.outcome === 'auth-failed') await clearToken();
  if (!preview && result.outcome === 'applied') await saveSyncState(result.state);
  const record = recordFromResult(request.trigger, startedAt, result);
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
      return await syncOnce(request);
    } catch (error) {
      // Nobody watches a scheduled run, so even an unexpected failure ends up in the status.
      const record = recordOf(request.trigger, new Date(), 'failed', {
        message: error instanceof Error ? error.message : 'Unexpected error',
      });
      if (request.mode === 'apply') await recordRun(record);
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
): Promise<Admission<ResetResult>> {
  const admitted = await flight.run(async () => {
    const { settings } = await loadSettings();
    const startedAt = new Date();
    const result = await resetMirror({
      fetch: doFetch,
      settings,
      mode,
      ...(confirmedCount === undefined ? {} : { confirmedCount }),
    });
    if (mode === 'apply' && result.outcome === 'reset') {
      await recordRun({
        kind: 'reset',
        trigger: 'manual',
        startedAt: startedAt.toISOString(),
        finishedAt: new Date().toISOString(),
        outcome: 'reset',
        message: `Deleted ${result.deleted} mirrored events`,
        repeats: 1,
      });
    }
    return result;
  });
  return admitted as Admission<ResetResult>;
}
