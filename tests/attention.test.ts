import { describe, expect, test } from 'bun:test';
import { attentionOf } from '../src/extension/attention.ts';
import { nextBackoff } from '../src/extension/backoff.ts';
import type { RunOutcome, RunRecord } from '../src/extension/messages.ts';

function record(outcome: RunOutcome, rest: Partial<RunRecord> = {}): RunRecord {
  return {
    kind: 'sync',
    trigger: 'alarm',
    startedAt: '2026-10-05T12:00:00.000Z',
    finishedAt: '2026-10-05T12:00:01.000Z',
    outcome,
    repeats: 1,
    ...rest,
  };
}

describe('what needs the user', () => {
  test('a held-back plan, a rejected password and a missing permission show at once', () => {
    for (const outcome of [
      'needs-confirmation',
      'credentials-rejected',
      'no-permission',
    ] as const) {
      expect(attentionOf(record(outcome))).toBeString();
    }
    expect(attentionOf(record('invalid-settings'))).toBeString();
  });

  test('a missing or rejected Outlook session only shows once it persists', () => {
    for (const outcome of [
      'no-token',
      'auth-failed',
      'read-failed',
      'destination-failed',
    ] as const) {
      expect(attentionOf(record(outcome, { repeats: 2 }))).toBeUndefined();
      expect(attentionOf(record(outcome, { repeats: 3 }))).toBeString();
    }
  });

  test('failed writes show once they repeat', () => {
    const counts = { errors: 4 } as RunRecord['counts'];
    expect(attentionOf(record('applied', { counts, repeats: 1 }))).toBeUndefined();
    expect(attentionOf(record('applied', { counts, repeats: 3, message: 'HTTP 403' }))).toBe(
      'HTTP 403',
    );
  });

  test('a clean run, a reset and an unconfigured extension show nothing', () => {
    expect(attentionOf(record('applied', { repeats: 9 }))).toBeUndefined();
    expect(attentionOf(record('reset'))).toBeUndefined();
    expect(attentionOf(record('not-configured', { repeats: 9 }))).toBeUndefined();
  });
});

describe('back-off', () => {
  const failing = record('applied', {
    counts: { errors: 3 } as RunRecord['counts'],
    message: 'Writing failed with HTTP 403',
  });
  const NOW = 1_000_000;

  test('starts at the third identical failure and lasts an hour', () => {
    const first = nextBackoff(undefined, failing, NOW, 15);
    expect(first).toEqual({ signature: expect.any(String), count: 1 });
    const second = nextBackoff(first, failing, NOW, 15);
    expect(second?.retryAt).toBeUndefined();
    const third = nextBackoff(second, failing, NOW, 15);
    expect(third?.retryAt).toBe(NOW + 60 * 60_000);
  });

  test('never retries sooner than the sync interval', () => {
    let state = nextBackoff(undefined, failing, NOW, 1440);
    state = nextBackoff(state, failing, NOW, 1440);
    expect(nextBackoff(state, failing, NOW, 1440)?.retryAt).toBe(NOW + 1440 * 60_000);
  });

  test('a different failure or a success starts over', () => {
    const a = nextBackoff(nextBackoff(undefined, failing, NOW, 15), failing, NOW, 15);
    const other = record('applied', { counts: { errors: 3 } as RunRecord['counts'], message: 'x' });
    expect(nextBackoff(a, other, NOW, 15)?.count).toBe(1);
    expect(nextBackoff(a, record('applied'), NOW, 15)).toBeUndefined();
  });

  test('problems the user must fix are not backed off here', () => {
    expect(nextBackoff(undefined, record('no-token'), NOW, 15)).toBeUndefined();
    expect(nextBackoff(undefined, record('read-failed'), NOW, 15)).toBeUndefined();
  });
});
