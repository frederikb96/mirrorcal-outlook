import { describe, expect, test } from 'bun:test';
import { isTokenUsable, normalizeToken, parseAuthorizationHeader } from '../src/core/token.ts';
import { makeJwt } from './helpers/fixtures.ts';

describe('authorization header capture', () => {
  test('accepts a bearer JWT for the Outlook audience and reads its expiry', () => {
    const jwt = makeJwt({ exp: 1_900_000_123 });
    expect(parseAuthorizationHeader(`Bearer ${jwt}`)).toEqual({ token: jwt, exp: 1_900_000_123 });
    expect(parseAuthorizationHeader(`bearer ${jwt}`)?.token).toBe(jwt);
  });

  test('ignores tokens for other audiences', () => {
    expect(
      parseAuthorizationHeader(`Bearer ${makeJwt({ aud: 'https://graph.microsoft.com' })}`),
    ).toBeNull();
    expect(parseAuthorizationHeader(`Bearer ${makeJwt({ aud: undefined })}`)).toBeNull();
  });

  test('ignores anything that is not a bearer JWT', () => {
    expect(parseAuthorizationHeader('Basic YWxpY2U6cHc=')).toBeNull();
    expect(parseAuthorizationHeader('Bearer opaque-token')).toBeNull();
    expect(parseAuthorizationHeader('Bearer a.b.c')).toBeNull();
    expect(parseAuthorizationHeader('')).toBeNull();
    expect(parseAuthorizationHeader(`Bearer ${makeJwt({ exp: 'soon' })}`)).toBeNull();
  });
});

describe('token use', () => {
  test('normalizeToken strips the prefix and rejects non-JWTs', () => {
    const jwt = makeJwt();
    expect(normalizeToken(`Bearer ${jwt}`)).toBe(jwt);
    expect(normalizeToken(`  ${jwt}  `)).toBe(jwt);
    expect(() => normalizeToken('nope')).toThrow();
    expect(() => normalizeToken(`${jwt}\nX: y`)).toThrow();
  });

  test('is unusable shortly before it expires', () => {
    const now = 1_000_000 * 1000;
    expect(isTokenUsable({ token: 't', exp: 1_000_000 + 3600 }, now)).toBe(true);
    expect(isTokenUsable({ token: 't', exp: 1_000_000 + 30 }, now)).toBe(false);
    expect(isTokenUsable({ token: 't', exp: 1_000_000 - 1 }, now)).toBe(false);
    expect(isTokenUsable(null, now)).toBe(false);
  });
});
