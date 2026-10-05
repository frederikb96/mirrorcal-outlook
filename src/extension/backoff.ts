import { BACKOFF_AFTER, BACKOFF_MINUTES } from '../core/constants.ts';
import type { RunRecord } from './messages.ts';
import type { Backoff } from './storage.ts';

/** Outcomes that retrying every interval cannot fix on its own. */
function isRepeatableFailure(record: RunRecord): boolean {
  return (
    record.outcome === 'destination-failed' ||
    record.outcome === 'failed' ||
    record.outcome === 'needs-confirmation' ||
    (record.outcome === 'applied' && (record.counts?.errors ?? 0) > 0)
  );
}

const signatureOf = (record: RunRecord) =>
  [record.outcome, record.counts?.errors ?? 0, record.message ?? '', record.digest ?? ''].join('|');

/**
 * Tracks identical failures in a row. After a few, scheduled runs wait an hour (or the sync
 * interval when that is longer) before trying again; any other result starts over.
 */
export function nextBackoff(
  previous: Backoff | undefined,
  record: RunRecord,
  nowMs: number,
  intervalMinutes: number,
): Backoff | undefined {
  if (!isRepeatableFailure(record)) return undefined;
  const signature = signatureOf(record);
  const count = previous?.signature === signature ? previous.count + 1 : 1;
  if (count < BACKOFF_AFTER) return { signature, count };
  const minutes = Math.max(BACKOFF_MINUTES, intervalMinutes);
  return { signature, count, retryAt: nowMs + minutes * 60_000 };
}
