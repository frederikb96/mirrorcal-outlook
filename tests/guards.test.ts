import { describe, expect, test } from 'bun:test';
import { evaluateGuards } from '../src/core/guards.ts';

const verdict = (sourceCount: number, deletions: number, ownCount: number) =>
  evaluateGuards({
    sourceCount,
    deletions,
    ownCount,
    creations: 0,
    foreignCount: 0,
    otherOwners: [],
  });

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

describe('creations', () => {
  const creating = (over: Partial<Parameters<typeof evaluateGuards>[0]>) =>
    evaluateGuards({
      sourceCount: 5,
      deletions: 0,
      ownCount: 0,
      creations: 5,
      foreignCount: 0,
      otherOwners: [],
      ...over,
    });

  test('another mirror owner holds the plan back and names it', () => {
    const v = creating({ ownCount: 5, foreignCount: 5, otherOwners: ['laptop'] });
    expect(v).toMatchObject({ ok: false, code: 'other-owner' });
    expect(v.ok ? '' : v.message).toContain('laptop');
  });

  test('a destination with only foreign events holds the plan back', () => {
    expect(creating({ foreignCount: 30 })).toMatchObject({
      ok: false,
      code: 'foreign-destination',
    });
  });

  test('foreign events beside our own, or nothing to create, are fine', () => {
    expect(creating({ ownCount: 3, foreignCount: 30 })).toEqual({ ok: true });
    expect(creating({ creations: 0, foreignCount: 30 })).toEqual({ ok: true });
    expect(creating({ creations: 0, otherOwners: ['laptop'] })).toEqual({ ok: true });
  });
});
