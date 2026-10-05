import { OUTLOOK_TOKEN_AUDIENCE } from './constants.ts';

export interface CapturedToken {
  /** The bare JWT, without the `Bearer ` prefix. */
  token: string;
  /** Expiry, seconds since the epoch. */
  exp: number;
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
  return { token, exp };
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
