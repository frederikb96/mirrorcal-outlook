import { TOKEN_REPLACE_MARGIN_SECONDS } from '../core/constants.ts';
import { sha256Hex } from '../core/crypto.ts';
import {
  isOutlookInitiator,
  isTokenUsable,
  matchAccount,
  parseAuthorizationHeader,
} from '../core/token.ts';
import {
  clearAccount,
  clearIgnoredTokens,
  clearToken,
  countIgnoredToken,
  loadAccount,
  loadRejectedToken,
  loadToken,
  saveToken,
} from './storage.ts';

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
const MAX_REMEMBERED = 20;

// Header values that need no second look: the one that was taken (or is the stored token) ...
let lastHeader = '';
// ... and the ones refused because they belong to another account or were rejected.
const refused = new Set<string>();

/** Forgets every header seen (a fresh worker start does the same). */
export function resetCaptureMemory(): void {
  lastHeader = '';
  refused.clear();
}

function refuse(header: string) {
  if (refused.size >= MAX_REMEMBERED) refused.clear();
  refused.add(header);
}

/**
 * Releases the account binding: the stored session and everything remembered about other
 * sessions go too, so the next token the Outlook page sends decides which account is mirrored.
 */
export async function releaseAccount(): Promise<void> {
  await clearAccount();
  await clearToken();
  await clearIgnoredTokens();
  resetCaptureMemory();
}

/**
 * Takes up the bearer token of a request the Outlook page sent. Only the page's own requests
 * count, a bound mirror only follows its own account (other sign-ins are counted, so the status
 * can say so), a token Outlook already rejected is not taken again, and a usable token is only
 * replaced by one that clearly outlives it.
 */
export async function captureToken(
  request: ObservedRequest,
  nowMs = Date.now(),
): Promise<CaptureResult> {
  if (!isOutlookInitiator(request.initiator)) return NOTHING;
  const header = request.requestHeaders?.find((h) => h.name.toLowerCase() === 'authorization');
  const value = header?.value;
  if (!value || value === lastHeader || refused.has(value)) return NOTHING;
  const parsed = parseAuthorizationHeader(value);
  if (!parsed) return NOTHING;

  const pinned = await loadAccount();
  if (pinned && matchAccount(pinned, parsed.account) === 'mismatch') {
    refuse(value);
    await countIgnoredToken();
    return NOTHING;
  }
  if ((await loadRejectedToken()) === (await sha256Hex(parsed.token))) {
    refuse(value);
    return NOTHING;
  }

  const current = await loadToken();
  const hadUsable = isTokenUsable(current, nowMs);
  if (current?.token === parsed.token) {
    lastHeader = value;
    return NOTHING;
  }
  if (current && hadUsable) {
    const currentFits = !pinned || matchAccount(pinned, current.account) !== 'mismatch';
    // Without a bound account the first one seen is kept rather than whichever expires last.
    if (
      currentFits &&
      current.account &&
      parsed.account &&
      matchAccount(current.account, parsed.account) === 'mismatch'
    ) {
      refuse(value);
      await countIgnoredToken();
      return NOTHING;
    }
    if (currentFits && parsed.exp <= current.exp + TOKEN_REPLACE_MARGIN_SECONDS) return NOTHING;
  }
  await saveToken(parsed);
  lastHeader = value;
  return { stored: true, hadUsable };
}
