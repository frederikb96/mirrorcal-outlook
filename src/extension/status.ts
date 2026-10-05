import { isTokenUsable } from '../core/token.ts';
import type { StatusResponse } from './messages.ts';
import {
  loadAccount,
  loadBackoff,
  loadIgnoredTokens,
  loadPasswordRejected,
  loadSettings,
  loadStatus,
  loadToken,
} from './storage.ts';

/** Said wherever the mirror is waiting for a session, so a bound account never fails silently. */
export const ignoredTokensNote = (count: number) =>
  `${count} sign-in${count === 1 ? '' : 's'} of another account ignored: use Forget account if you switched`;

/** Everything the settings page shows about the mirror, read straight from storage. */
export async function buildStatus(nowMs = Date.now()): Promise<StatusResponse> {
  const { settings, stored } = await loadSettings();
  const [token, saved, pinned, passwordRejected, backoff, ignoredTokens] = await Promise.all([
    loadToken(),
    loadStatus(),
    loadAccount(),
    loadPasswordRejected(),
    loadBackoff(),
    loadIgnoredTokens(),
  ]);
  const response: StatusResponse = {
    configured: stored,
    enabled: settings.enabled,
    token: {
      present: token !== null,
      usable: isTokenUsable(token, nowMs),
      ...(token ? { expiresAt: new Date(token.exp * 1000).toISOString() } : {}),
    },
    passwordRejected,
    ignoredTokens,
    log: saved.log,
  };
  const account = pinned ?? token?.account;
  if (account) response.account = { label: account.label, pinned: pinned !== undefined };
  if (backoff?.retryAt !== undefined && backoff.retryAt > nowMs) {
    response.retryAt = new Date(backoff.retryAt).toISOString();
  }
  if (saved.lastRun) response.lastRun = saved.lastRun;
  if (saved.lastSuccessAt) response.lastSuccessAt = saved.lastSuccessAt;
  return response;
}
