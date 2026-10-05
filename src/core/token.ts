import { OUTLOOK_HOSTS, OUTLOOK_TOKEN_AUDIENCE } from './constants.ts';

/**
 * Who a token was issued to. The tenant is always present, next to the user's object id and/or
 * name claim; tokens of one user do not always carry the same mix of claims.
 */
export interface TokenAccount {
  tenant: string;
  oid?: string;
  name?: string;
  /** What is shown to the user. */
  label: string;
}

/** `unknown` when the two carry no claim that can be compared. */
export type AccountMatch = 'match' | 'mismatch' | 'unknown';

/**
 * Compares a token's account with the one the mirror is bound to. Within one tenant a matching
 * object id or a matching name is enough; a candidate that cannot be compared is not a mismatch.
 */
export function matchAccount(
  bound: TokenAccount,
  candidate: TokenAccount | undefined,
): AccountMatch {
  if (!candidate) return 'unknown';
  if (candidate.tenant !== bound.tenant) return 'mismatch';
  let differs = false;
  for (const claim of ['oid', 'name'] as const) {
    const a = bound[claim]?.toLowerCase();
    const b = candidate[claim]?.toLowerCase();
    if (a === undefined || b === undefined) continue;
    if (a === b) return 'match';
    differs = true;
  }
  return differs ? 'mismatch' : 'unknown';
}

export interface CapturedToken {
  /** The bare JWT, without the `Bearer ` prefix. */
  token: string;
  /** Expiry, seconds since the epoch. */
  exp: number;
  /** Absent when the token carries no tenant and user claims. */
  account?: TokenAccount;
}

const JWT_PART = '[A-Za-z0-9_-]+';
const JWT_PATTERN = new RegExp(`^${JWT_PART}\\.${JWT_PART}\\.${JWT_PART}$`);

function decodePayload(jwt: string): Record<string, unknown> | null {
  const payload = jwt.split('.')[1];
  if (!payload) return null;
  try {
    const padded = payload
      .replace(/-/g, '+')
      .replace(/_/g, '/')
      .padEnd(Math.ceil(payload.length / 4) * 4, '=');
    const bytes = Uint8Array.from(atob(padded), (c) => c.charCodeAt(0));
    const parsed: unknown = JSON.parse(new TextDecoder().decode(bytes));
    return typeof parsed === 'object' && parsed !== null
      ? (parsed as Record<string, unknown>)
      : null;
  } catch {
    return null;
  }
}

function claim(payload: Record<string, unknown>, name: string): string | undefined {
  const value = payload[name];
  return typeof value === 'string' && value !== '' ? value : undefined;
}

function accountOf(payload: Record<string, unknown>): TokenAccount | undefined {
  const tenant = claim(payload, 'tid');
  const name =
    claim(payload, 'upn') ??
    claim(payload, 'unique_name') ??
    claim(payload, 'preferred_username') ??
    claim(payload, 'email');
  const oid = claim(payload, 'oid');
  if (!tenant || (!oid && !name)) return undefined;
  return {
    tenant,
    ...(oid ? { oid } : {}),
    ...(name ? { name } : {}),
    label: (name ?? oid) as string,
  };
}

/** The tenant and user a token was issued to, or undefined when it does not say. */
export function tokenAccount(jwt: string): TokenAccount | undefined {
  const payload = decodePayload(jwt);
  return payload ? accountOf(payload) : undefined;
}

/** True for a request that the Outlook web page itself sent, and for nothing else. */
export function isOutlookInitiator(initiator: string | undefined): boolean {
  return OUTLOOK_HOSTS.some((host) => initiator === `https://${host}`);
}

/**
 * Reads the value of an `Authorization` request header. Returns the token only when it is a bearer
 * JWT issued for the Outlook API; tokens for any other service the page talks to are ignored.
 */
export function parseAuthorizationHeader(header: string): CapturedToken | null {
  const match = /^Bearer\s+(\S+)$/i.exec(header.trim());
  const token = match?.[1];
  if (!token || !JWT_PATTERN.test(token)) return null;
  const payload = decodePayload(token);
  if (!payload) return null;
  const { aud, exp } = payload;
  if (aud !== OUTLOOK_TOKEN_AUDIENCE) return null;
  if (typeof exp !== 'number' || !Number.isFinite(exp)) return null;
  const account = accountOf(payload);
  return account ? { token, exp, account } : { token, exp };
}

/** Accepts a bare token or a full `Bearer …` header value; rejects anything that is not a JWT. */
export function normalizeToken(raw: string): string {
  const bare = raw.trim().replace(/^Bearer\s+/i, '');
  if (!JWT_PATTERN.test(bare)) throw new Error('The Outlook token is not a JWT');
  return bare;
}

const EXPIRY_SKEW_SECONDS = 60;

export function isTokenUsable(token: CapturedToken | null | undefined, nowMs: number) {
  return !!token && token.exp * 1000 - EXPIRY_SKEW_SECONDS * 1000 > nowMs;
}
