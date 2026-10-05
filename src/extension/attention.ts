import { ATTENTION_AFTER_REPEATS } from '../core/constants.ts';
import type { RunRecord } from './messages.ts';

/** Outcomes that only need the user once they keep repeating. */
const PERSISTENT = new Set([
  'no-token',
  'auth-failed',
  'read-failed',
  'destination-failed',
  'failed',
]);

/**
 * What the toolbar icon should say about the newest run, or undefined when nothing needs the
 * user. Computed here once; the badge, its tooltip and nothing else read this answer.
 */
export function attentionOf(record: RunRecord): string | undefined {
  const { outcome, repeats } = record;
  if (outcome === 'invalid-settings') return 'The settings are invalid';
  if (outcome === 'needs-confirmation')
    return 'A change is held back and waits for your confirmation';
  if (outcome === 'credentials-rejected') return 'The calendar server rejected the password';
  if (outcome === 'no-permission') return 'The extension has no access to the calendar server';
  const failing =
    PERSISTENT.has(outcome) || (outcome === 'applied' && (record.counts?.errors ?? 0) > 0);
  if (failing && repeats >= ATTENTION_AFTER_REPEATS) {
    return record.message ?? 'Syncing keeps failing';
  }
  return undefined;
}

export async function showAttention(record: RunRecord): Promise<void> {
  const reason = attentionOf(record);
  await chrome.action.setBadgeText({ text: reason ? '!' : '' });
  if (reason) await chrome.action.setBadgeBackgroundColor({ color: '#b83232' });
  await chrome.action.setTitle({
    title: reason ? `MirrorCal Outlook: ${reason}` : 'MirrorCal Outlook settings',
  });
}
