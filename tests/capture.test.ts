import { beforeEach, describe, expect, test } from 'bun:test';
import { sha256Hex } from '../src/core/crypto.ts';
import {
  captureToken,
  type ObservedRequest,
  releaseAccount,
  resetCaptureMemory,
} from '../src/extension/capture.ts';
import {
  loadAccount,
  loadIgnoredTokens,
  loadToken,
  saveAccount,
  saveRejectedToken,
  saveToken,
} from '../src/extension/storage.ts';
import { installFakeChrome } from './helpers/fake-chrome.ts';
import { makeJwt } from './helpers/fixtures.ts';

const PAGE = 'https://outlook.cloud.microsoft';
const NOW = Date.UTC(2026, 9, 5, 12);
const nowSeconds = Math.floor(NOW / 1000);

const tokenFor = (oid: string, exp = nowSeconds + 3600) =>
  makeJwt({ oid, tid: 'tenant-1', upn: `${oid}@example.test`, exp });

const request = (jwt: string, initiator: string | null = PAGE): ObservedRequest => ({
  ...(initiator === null ? {} : { initiator }),
  requestHeaders: [{ name: 'Authorization', value: `Bearer ${jwt}` }],
});

let sessionWrites: number;

beforeEach(() => {
  const fake = installFakeChrome();
  resetCaptureMemory();
  sessionWrites = 0;
  const set = chrome.storage.session.set.bind(chrome.storage.session);
  chrome.storage.session.set = async (items) => {
    sessionWrites += 1;
    await set(items);
  };
  void fake;
});

describe('where a token may come from', () => {
  test('the Outlook page is the only source', async () => {
    const jwt = tokenFor('a');
    for (const initiator of [null, 'https://evil.test', 'chrome-extension://test-extension-id']) {
      expect((await captureToken(request(jwt, initiator), NOW)).stored).toBe(false);
      resetCaptureMemory();
    }
    expect(await loadToken()).toBeNull();
    expect((await captureToken(request(jwt), NOW)).stored).toBe(true);
    expect((await loadToken())?.token).toBe(jwt);
  });

  test('a token Outlook already rejected is not taken up again', async () => {
    const jwt = tokenFor('a');
    await saveRejectedToken(await sha256Hex(jwt));
    expect((await captureToken(request(jwt), NOW)).stored).toBe(false);
    expect((await captureToken(request(tokenFor('a', nowSeconds + 7200)), NOW)).stored).toBe(true);
  });
});

describe('which account is followed', () => {
  test('a pinned mirror ignores every other account', async () => {
    await saveAccount({
      tenant: 'tenant-1',
      oid: 'a',
      name: 'a@example.test',
      label: 'a@example.test',
    });
    expect((await captureToken(request(tokenFor('b')), NOW)).stored).toBe(false);
    expect((await captureToken(request(tokenFor('a')), NOW)).stored).toBe(true);
  });

  test('a pinned mirror drops a stored token of another account for its own', async () => {
    await saveToken({
      token: tokenFor('b', nowSeconds + 9000),
      exp: nowSeconds + 9000,
      account: { tenant: 'tenant-1', oid: 'b', label: 'b' },
    });
    await saveAccount({
      tenant: 'tenant-1',
      oid: 'a',
      name: 'a@example.test',
      label: 'a@example.test',
    });
    const own = tokenFor('a');
    expect((await captureToken(request(own), NOW)).stored).toBe(true);
    expect((await loadToken())?.token).toBe(own);
  });

  test('before the first sync the first account seen is kept', async () => {
    const first = tokenFor('a');
    await captureToken(request(first), NOW);
    expect((await captureToken(request(tokenFor('b', nowSeconds + 9000)), NOW)).stored).toBe(false);
    expect((await loadToken())?.token).toBe(first);
  });
});

describe('replacing a usable token', () => {
  test('one that does not clearly outlive it causes no storage write', async () => {
    await captureToken(request(tokenFor('a', nowSeconds + 3600)), NOW);
    sessionWrites = 0;
    const result = await captureToken(request(tokenFor('a', nowSeconds + 3700)), NOW);
    expect(result.stored).toBe(false);
    expect(sessionWrites).toBe(0);
  });

  test('one that outlives it by more than ten minutes replaces it', async () => {
    await captureToken(request(tokenFor('a', nowSeconds + 3600)), NOW);
    const newer = tokenFor('a', nowSeconds + 3600 + 700);
    const result = await captureToken(request(newer), NOW);
    expect(result).toEqual({ stored: true, hadUsable: true });
    expect((await loadToken())?.token).toBe(newer);
  });

  test('an expired token is replaced by any valid one, reporting that none was usable', async () => {
    await saveToken({ token: tokenFor('a', nowSeconds - 10), exp: nowSeconds - 10 });
    const result = await captureToken(request(tokenFor('a', nowSeconds + 60 * 5)), NOW);
    expect(result).toEqual({ stored: true, hadUsable: false });
  });

  test('the same header twice is looked at once', async () => {
    const jwt = tokenFor('a');
    await captureToken(request(jwt), NOW);
    sessionWrites = 0;
    expect((await captureToken(request(jwt), NOW)).stored).toBe(false);
    expect(sessionWrites).toBe(0);
  });
});

describe('a bound mirror that must not lock its user out', () => {
  const pin = () =>
    saveAccount({ tenant: 'tenant-1', oid: 'a', name: 'a@example.test', label: 'a@example.test' });

  test('a token of the same user with another mix of claims is accepted', async () => {
    await pin();
    const nameOnly = makeJwt({ tid: 'tenant-1', upn: 'a@example.test', exp: nowSeconds + 3600 });
    expect((await captureToken(request(nameOnly), NOW)).stored).toBe(true);
  });

  test('a changed object id with the same name is still the same user', async () => {
    await pin();
    const moved = makeJwt({
      tid: 'tenant-1',
      oid: 'new',
      upn: 'a@example.test',
      exp: nowSeconds + 3600,
    });
    expect((await captureToken(request(moved), NOW)).stored).toBe(true);
  });

  test('a token that carries nothing comparable is accepted', async () => {
    await pin();
    expect((await captureToken(request(makeJwt({ exp: nowSeconds + 3600 })), NOW)).stored).toBe(
      true,
    );
  });

  test('another tenant is ignored and counted once per distinct sign-in', async () => {
    await pin();
    const other = makeJwt({ tid: 'tenant-2', oid: 'a', exp: nowSeconds + 3600 });
    await captureToken(request(other), NOW);
    await captureToken(request(other), NOW);
    expect(await loadIgnoredTokens()).toBe(1);
    await captureToken(
      request(makeJwt({ tid: 'tenant-3', oid: 'a', exp: nowSeconds + 3600 })),
      NOW,
    );
    expect(await loadIgnoredTokens()).toBe(2);
  });
});

describe('releasing the account', () => {
  test('drops the binding and the stored session, so the next token decides', async () => {
    await saveAccount({ tenant: 'tenant-1', oid: 'a', name: 'a@example.test', label: 'a' });
    const old = tokenFor('a', nowSeconds + 3600);
    await captureToken(request(old), NOW);
    const switched = tokenFor('b', nowSeconds + 3600);
    await captureToken(request(switched), NOW);
    expect(await loadIgnoredTokens()).toBe(1);

    await releaseAccount();
    expect(await loadAccount()).toBeUndefined();
    expect(await loadToken()).toBeNull();
    expect(await loadIgnoredTokens()).toBe(0);
    // The sign-in that was ignored while bound is looked at again.
    expect((await captureToken(request(switched), NOW)).stored).toBe(true);
    expect((await loadToken())?.token).toBe(switched);
  });
});

describe('a token that is not taken is not remembered as seen', () => {
  test('a not-newer token is taken when the stored one has run out', async () => {
    const a = tokenFor('a', nowSeconds + 120);
    const b = tokenFor('a', nowSeconds + 200);
    await captureToken(request(a), NOW);
    expect((await captureToken(request(b), NOW)).stored).toBe(false);
    const later = NOW + 100_000;
    expect((await captureToken(request(b), later)).stored).toBe(true);
  });
});
