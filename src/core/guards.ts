import { MASS_DELETE_MIN, MASS_DELETE_SHARE } from './constants.ts';

export type GuardCode = 'empty-source' | 'mass-deletion' | 'other-owner' | 'foreign-destination';

export type GuardVerdict = { ok: true } | { ok: false; code: GuardCode; message: string };

export interface GuardInput {
  /** Items Outlook returned for the window, before any filtering. */
  sourceCount: number;
  deletions: number;
  /** Objects in the destination carrying this owner's marker. */
  ownCount: number;
  creations: number;
  /** Objects in the destination that do not carry this owner's marker. */
  foreignCount: number;
  /** Distinct owner ids found on other mirrors' objects in the destination. */
  otherOwners: string[];
}

/**
 * Decides whether a plan may be applied without an explicit confirmation. A plan that would wipe
 * the mirror is far more likely to come from a bad read, a wrong window or a sweeping new filter
 * than from the calendar really emptying. Creating events is held back too when the destination
 * already holds another mirror's events (they would be duplicated) or only events that are not
 * this mirror's (most likely the wrong calendar).
 */
export function evaluateGuards(input: GuardInput): GuardVerdict {
  const { sourceCount, deletions, ownCount, creations, foreignCount, otherOwners } = input;
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
  if (creations > 0 && otherOwners.length > 0) {
    const ids = otherOwners.slice(0, 3).join(', ');
    return {
      ok: false,
      code: 'other-owner',
      message: `The calendar holds events of another mirror owner id (${ids}); creating events would duplicate them. Remove them with the settings page's "another owner" reset, or confirm to create the events next to them`,
    };
  }
  if (creations > 0 && ownCount === 0 && foreignCount > 0) {
    return {
      ok: false,
      code: 'foreign-destination',
      message: `The calendar already holds ${foreignCount} events that are not mirrored by this extension; is this the right calendar?`,
    };
  }
  return { ok: true };
}
