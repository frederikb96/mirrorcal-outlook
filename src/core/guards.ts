import { MASS_DELETE_MIN, MASS_DELETE_SHARE } from './constants.ts';

export type GuardVerdict =
  | { ok: true }
  | { ok: false; code: 'empty-source' | 'mass-deletion'; message: string };

export interface GuardInput {
  /** Items Outlook returned for the window, before any filtering. */
  sourceCount: number;
  deletions: number;
  /** Objects in the destination carrying this owner's marker. */
  ownCount: number;
}

/**
 * Decides whether a plan may be applied without an explicit confirmation. A plan that would wipe
 * the mirror is far more likely to come from a bad read, a wrong window or a sweeping new filter
 * than from the calendar really emptying.
 */
export function evaluateGuards(input: GuardInput): GuardVerdict {
  const { sourceCount, deletions, ownCount } = input;
  if (sourceCount === 0 && ownCount > 0) {
    return {
      ok: false,
      code: 'empty-source',
      message: `Outlook returned no events but the calendar holds ${ownCount} mirrored events`,
    };
  }
  if (deletions > MASS_DELETE_MIN && deletions > ownCount * MASS_DELETE_SHARE) {
    return {
      ok: false,
      code: 'mass-deletion',
      message: `The plan deletes ${deletions} of ${ownCount} mirrored events`,
    };
  }
  return { ok: true };
}
