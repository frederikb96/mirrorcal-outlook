import { describe, expect, test } from 'bun:test';
import { applyPlan } from '../src/core/apply.ts';
import type { CalDavTarget } from '../src/core/caldav.ts';
import type { Plan } from '../src/core/plan.ts';
import { FakeCalDav } from './helpers/fake-caldav.ts';
import { COLLECTION } from './helpers/fixtures.ts';

const target: CalDavTarget = { url: COLLECTION, username: 'alice', password: 'test-app-password' };

function emptyPlan(overrides: Partial<Plan> = {}): Plan {
  return {
    creations: [],
    updates: [],
    deletions: [],
    unchanged: 0,
    conflicts: [],
    ownCount: 0,
    ownHrefs: [],
    foreign: [],
    ...overrides,
  };
}

const deletion = (name: string) => ({
  kind: 'delete' as const,
  uid: name,
  href: `${COLLECTION}${name}.ics`,
  etag: '"e"',
  summary: name,
  start: '',
});

describe('applying a plan', () => {
  test('refuses to delete or replace an object that is not known to carry the owner marker', async () => {
    const dav = new FakeCalDav();
    dav.seed('theirs.ics', 'THEIRS');
    const plan = emptyPlan({ deletions: [deletion('theirs')] });
    await expect(
      applyPlan(plan, { fetch: dav.fetch, target, ownedUrls: new Set() }),
    ).rejects.toThrow('owner marker');
    const update = {
      kind: 'update' as const,
      uid: 'theirs',
      href: `${COLLECTION}theirs.ics`,
      etag: '"e"',
      ics: 'X',
      hash: 'h',
      summary: 's',
      start: '',
      allDay: false,
    };
    await expect(
      applyPlan(emptyPlan({ updates: [update] }), {
        fetch: dav.fetch,
        target,
        ownedUrls: new Set([`${COLLECTION}other.ics`]),
      }),
    ).rejects.toThrow('owner marker');
    expect(dav.requests).toHaveLength(0);
    expect(dav.objects.get('theirs.ics')?.ics).toBe('THEIRS');
  });

  test('deletes owned objects and counts them', async () => {
    const dav = new FakeCalDav();
    dav.seed('mine.ics', 'MINE');
    const href = `${COLLECTION}mine.ics`;
    const etag = dav.objects.get('mine.ics')?.etag ?? '';
    const plan = emptyPlan({ deletions: [{ ...deletion('mine'), etag }] });
    const result = await applyPlan(plan, { fetch: dav.fetch, target, ownedUrls: new Set([href]) });
    expect(result).toMatchObject({ deleted: 1, errors: [] });
    expect(dav.objects.size).toBe(0);
  });
});
