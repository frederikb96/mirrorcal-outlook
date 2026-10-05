import { OUTLOOK_HOSTS, TOKEN_CAPTURE_PATHS } from '../core/constants.ts';
import { isTokenUsable, parseAuthorizationHeader } from '../core/token.ts';
import type { Request, ResetResponse, StatusResponse, SyncResponse } from './messages.ts';
import { runResetFlight, runSyncFlight } from './runner.ts';
import { loadSettings, loadStatus, loadToken, saveToken } from './storage.ts';

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

async function status(): Promise<StatusResponse> {
  const { settings, stored } = await loadSettings();
  const token = await loadToken();
  const saved = await loadStatus();
  const response: StatusResponse = {
    configured: stored,
    enabled: settings.enabled,
    token: {
      present: token !== null,
      usable: isTokenUsable(token, Date.now()),
      ...(token ? { expiresAt: new Date(token.exp * 1000).toISOString() } : {}),
    },
    log: saved.log,
  };
  if (saved.lastRun) response.lastRun = saved.lastRun;
  if (saved.lastSuccessAt) response.lastSuccessAt = saved.lastSuccessAt;
  return response;
}

async function handle(request: Request): Promise<StatusResponse | SyncResponse | ResetResponse> {
  switch (request.type) {
    case 'status':
      return status();
    case 'sync': {
      const admitted = await runSyncFlight({
        trigger: 'manual',
        mode: request.mode,
        ...(request.confirm ? { confirm: request.confirm } : {}),
      });
      return admitted.busy ? { busy: true } : { busy: false, ...admitted.value };
    }
    case 'reset': {
      const admitted = await runResetFlight(request.mode, request.confirmedCount);
      return admitted.busy ? { busy: true } : { busy: false, result: admitted.value };
    }
  }
}

// Only the page-level GET of the bearer header is observed; the token is never logged.
let lastHeader = '';

async function remember(header: string): Promise<void> {
  const parsed = parseAuthorizationHeader(header);
  if (!parsed) return;
  const current = await loadToken();
  if (current && current.exp > parsed.exp && isTokenUsable(current, Date.now())) return;
  await saveToken(parsed);
  if (!isTokenUsable(current, Date.now())) void runSyncFlight({ trigger: 'token', mode: 'apply' });
}

chrome.webRequest.onSendHeaders.addListener(
  (details) => {
    if (details.initiator?.startsWith(`chrome-extension://${chrome.runtime.id}`)) return;
    const header = details.requestHeaders?.find((h) => h.name.toLowerCase() === 'authorization');
    if (!header?.value || header.value === lastHeader) return;
    lastHeader = header.value;
    void remember(header.value);
  },
  {
    urls: OUTLOOK_HOSTS.flatMap((host) =>
      TOKEN_CAPTURE_PATHS.map((path) => `https://${host}${path}`),
    ),
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
