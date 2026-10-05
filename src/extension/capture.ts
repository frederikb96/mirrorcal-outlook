import { TOKEN_REPLACE_MARGIN_SECONDS } from '../core/constants.ts';
import { sha256Hex } from '../core/crypto.ts';
import { isOutlookInitiator, isTokenUsable, parseAuthorizationHeader } from '../core/token.ts';
import { loadAccount, loadRejectedToken, loadToken, saveToken } from './storage.ts';

export interface ObservedRequest {
  initiator?: string | undefined;
  requestHeaders?: { name: string; value?: string | undefined }[] | undefined;
}

export interface CaptureResult {
  /** The token was stored. */
  stored: boolean;
  /** A usable token was already stored before this one arrived. */
  hadUsable: boolean;
}

const NOTHING: CaptureResult = { stored: false, hadUsable: false };

// The last header value seen; the page repeats the same one on most requests.
let lastHeader = '';

/** Forgets the last header seen (a fresh worker start does the same). */
export function resetCaptureMemory(): void {
  lastHeader = '';
}

/**
 * Takes up the bearer token of a request the Outlook page sent. Only the page's own requests
 * count, a pinned mirror only follows its own account, a token Outlook already rejected is not
 * taken again, and a usable token is only replaced by one that clearly outlives it.
 */
export async function captureToken(
  request: ObservedRequest,
  nowMs = Date.now(),
): Promise<CaptureResult> {
  if (!isOutlookInitiator(request.initiator)) return NOTHING;
  const header = request.requestHeaders?.find((h) => h.name.toLowerCase() === 'authorization');
  if (!header?.value || header.value === lastHeader) return NOTHING;
  lastHeader = header.value;
  const parsed = parseAuthorizationHeader(header.value);
  if (!parsed) return NOTHING;

  const pinned = await loadAccount();
  if (pinned && parsed.account?.id !== pinned.id) return NOTHING;
  if ((await loadRejectedToken()) === (await sha256Hex(parsed.token))) return NOTHING;

  const current = await loadToken();
  const hadUsable = isTokenUsable(current, nowMs);
  if (current && hadUsable) {
    const currentFits = !pinned || current.account?.id === pinned.id;
    // Without a pinned account the first one seen is kept rather than whichever expires last.
    if (currentFits && current.account?.id !== parsed.account?.id) return NOTHING;
    if (currentFits && parsed.exp <= current.exp + TOKEN_REPLACE_MARGIN_SECONDS) return NOTHING;
  }
  await saveToken(parsed);
  return { stored: true, hadUsable };
}
