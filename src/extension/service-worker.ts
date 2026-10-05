import { OUTLOOK_HOSTS, TOKEN_CAPTURE_PATHS } from '../core/constants.ts';
import { captureToken, releaseAccount } from './capture.ts';
import type {
  Request,
  ResetResponse,
  SaveResponse,
  StatusResponse,
  SyncResponse,
} from './messages.ts';
import { runResetFlight, runSyncFlight } from './runner.ts';
import { saveSettingsFlow } from './save.ts';
import { buildStatus } from './status.ts';
import { loadSettings } from './storage.ts';

const ALARM = 'mirrorcal-sync';

/** Keeps exactly one repeating alarm, at the configured interval. */
async function ensureAlarm(): Promise<void> {
  const { settings } = await loadSettings();
  const existing = await chrome.alarms.get(ALARM);
  if (existing?.periodInMinutes === settings.intervalMinutes) return;
  await chrome.alarms.create(ALARM, {
    delayInMinutes: 1,
    periodInMinutes: settings.intervalMinutes,
  });
}

async function handle(
  request: Request,
): Promise<StatusResponse | SyncResponse | ResetResponse | SaveResponse | Record<string, never>> {
  switch (request.type) {
    case 'status':
      return buildStatus();
    case 'sync': {
      const admitted = await runSyncFlight({
        trigger: 'manual',
        mode: request.mode,
        ...(request.confirm ? { confirm: request.confirm } : {}),
      });
      return admitted.busy ? { busy: true } : { busy: false, ...admitted.value };
    }
    case 'reset': {
      const admitted = await runResetFlight(request.mode, request.confirmedCount, request.ownerId);
      return admitted.busy ? { busy: true } : { busy: false, result: admitted.value };
    }
    case 'save':
      return saveSettingsFlow(request.settings);
    case 'forget-account':
      await releaseAccount();
      return {};
  }
}

// Only requests the Outlook page itself sent are observed; the token is never logged.
chrome.webRequest.onSendHeaders.addListener(
  (details) => {
    void captureToken(details).then(({ stored, hadUsable }) => {
      if (stored && !hadUsable) void runSyncFlight({ trigger: 'token', mode: 'apply' });
    });
  },
  {
    urls: OUTLOOK_HOSTS.flatMap((host) =>
      TOKEN_CAPTURE_PATHS.map((path) => `https://${host}${path}`),
    ),
    types: ['xmlhttprequest'],
  },
  ['requestHeaders', 'extraHeaders'],
);

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === ALARM) void runSyncFlight({ trigger: 'alarm', mode: 'apply' });
});

// Every worker start re-checks the alarm, so one lost across a browser restart comes back.
void ensureAlarm();

chrome.runtime.onInstalled.addListener(() => void ensureAlarm());

chrome.runtime.onStartup.addListener(() => {
  void runSyncFlight({ trigger: 'startup', mode: 'apply' });
});

chrome.storage.onChanged.addListener((changes, area) => {
  if (area === 'local' && 'settings' in changes) void ensureAlarm();
});

chrome.action.onClicked.addListener(() => void chrome.runtime.openOptionsPage());

chrome.runtime.onMessage.addListener((message: Request, sender, sendResponse) => {
  if (sender.id !== chrome.runtime.id) return false;
  handle(message).then(sendResponse, (error: unknown) =>
    sendResponse({ busy: false, error: error instanceof Error ? error.message : 'failed' }),
  );
  return true;
});
