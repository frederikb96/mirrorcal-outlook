import { describe, expect, test } from 'bun:test';
import { evaluateGuards } from '../src/core/guards.ts';

const verdict = (sourceCount: number, deletions: number, ownCount: number) =>
  evaluateGuards({ sourceCount, deletions, ownCount });

describe('empty source', () => {
  test('refused when mirrored events exist', () => {
    const v = verdict(0, 5, 5);
    expect(v).toMatchObject({ ok: false, code: 'empty-source' });
  });

  test('fine on the very first sync', () => {
    expect(verdict(0, 0, 0)).toEqual({ ok: true });
  });

  test('refused even for a handful of events', () => {
    expect(verdict(0, 1, 1).ok).toBe(false);
  });
});

describe('mass deletion', () => {
  test('refused above 20 deletions and half of what is owned', () => {
    expect(verdict(50, 21, 30)).toMatchObject({ ok: false, code: 'mass-deletion' });
    expect(verdict(50, 51, 100)).toMatchObject({ ok: false, code: 'mass-deletion' });
  });

  test('allowed at or below either threshold', () => {
    expect(verdict(50, 20, 21).ok).toBe(true);
    expect(verdict(50, 50, 100).ok).toBe(true);
    expect(verdict(50, 21, 100).ok).toBe(true);
    expect(verdict(50, 3, 3).ok).toBe(true);
  });

  test('messages state the numbers', () => {
    const v = verdict(50, 30, 40);
    expect(v.ok ? '' : v.message).toContain('30 of 40');
  });
});
