import { OUTLOOK_HOSTS } from '../core/constants.ts';
import { checkDestination } from '../core/engine.ts';
import { originPattern, prepareSave, validateSettings } from '../core/settings.ts';
import type { Settings } from '../core/types.ts';
import type { SaveResponse } from './messages.ts';
import { doFetch, runSyncFlight } from './runner.ts';
import { loadSettings, saveSettings } from './storage.ts';

/** Drops the access to a calendar server the settings no longer point at. */
async function releaseOldOrigin(before: Settings, after: Settings): Promise<void> {
  const old = originPattern(before.caldavUrl);
  if (!old || old === originPattern(after.caldavUrl)) return;
  if (OUTLOOK_HOSTS.some((host) => old === `https://${host}/*`)) return;
  try {
    await chrome.permissions.remove({ origins: [old] });
  } catch {
    // Not granted any more, or granted by the manifest: nothing to release.
  }
}

/**
 * Validates the draft, asks the calendar server what the URL is, stores the settings and starts a
 * sync so the result shows at once. A server that answers "no" (wrong password, not a calendar)
 * stops the save; one that cannot be reached does not, so settings can be entered offline.
 */
export async function saveSettingsFlow(draft: Settings): Promise<SaveResponse> {
  const { settings: existing } = await loadSettings();
  const prepared = prepareSave(existing, draft);
  if ('errors' in prepared) return { saved: false, errors: prepared.errors };
  const settings = prepared.settings;
  const errors = validateSettings(settings);
  if (errors.length > 0) return { saved: false, errors };

  const check = await checkDestination({ fetch: doFetch, settings });
  if (check.outcome !== 'ok' && check.outcome !== 'unreachable') {
    return { saved: false, errors: [check.message] };
  }
  await saveSettings(settings);
  await releaseOldOrigin(existing, settings);
  const admitted = await runSyncFlight({ trigger: 'manual', mode: 'apply' });
  return {
    saved: true,
    connection:
      check.outcome === 'ok'
        ? check.message
        : `Could not reach the calendar server to check it: ${check.message}`,
    sync: admitted.busy ? { busy: true } : { busy: false, ...admitted.value },
  };
}
