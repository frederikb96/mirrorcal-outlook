import {
  CalDavError,
  type CalDavTarget,
  deleteObject,
  putObject,
  resolveInCollection,
} from './caldav.ts';
import { WRITE_CONCURRENCY } from './constants.ts';
import type { Plan } from './plan.ts';
import type { FetchLike } from './types.ts';

export interface OpError {
  op: 'create' | 'update' | 'delete';
  uid: string;
  message: string;
}

export interface ApplyResult {
  created: number;
  updated: number;
  deleted: number;
  errors: OpError[];
  /** Set when the run stopped early (credentials rejected); deletions are then skipped. */
  stopped?: string;
}

export interface ApplyContext {
  fetch: FetchLike;
  target: CalDavTarget;
  /** URLs of the objects that carry this owner's marker; the only ones that may be changed. */
  ownedUrls: ReadonlySet<string>;
}

async function pool<T>(
  items: T[],
  size: number,
  work: (item: T) => Promise<void>,
  halted: () => boolean,
) {
  let next = 0;
  const worker = async () => {
    while (next < items.length && !halted()) {
      const item = items[next++] as T;
      await work(item);
    }
  };
  await Promise.all(Array.from({ length: Math.min(size, items.length) }, worker));
}

/**
 * Writes the plan: creations and updates first, deletions last. A failing object is recorded and
 * the rest carry on; rejected credentials stop everything, deletions included.
 */
export async function applyPlan(plan: Plan, ctx: ApplyContext): Promise<ApplyResult> {
  for (const op of [...plan.updates, ...plan.deletions]) {
    if (!ctx.ownedUrls.has(op.href)) {
      throw new Error('Refusing to modify an object that does not carry the owner marker');
    }
  }
  const result: ApplyResult = { created: 0, updated: 0, deleted: 0, errors: [] };
  const halted = () => result.stopped !== undefined;
  const guarded = async (op: OpError['op'], uid: string, run: () => Promise<void>) => {
    try {
      await run();
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Unexpected error';
      result.errors.push({ op, uid, message });
      if (error instanceof CalDavError && error.code === 'auth') result.stopped = message;
    }
  };

  await pool(
    plan.creations,
    WRITE_CONCURRENCY,
    (op) =>
      guarded('create', op.uid, async () => {
        await putObject(ctx.fetch, ctx.target, resolveInCollection(ctx.target, op.href), op.ics, {
          create: true,
        });
        result.created += 1;
      }),
    halted,
  );
  await pool(
    plan.updates,
    WRITE_CONCURRENCY,
    (op) =>
      guarded('update', op.uid, async () => {
        await putObject(ctx.fetch, ctx.target, op.href, op.ics, { replace: op.etag });
        result.updated += 1;
      }),
    halted,
  );
  await pool(
    plan.deletions,
    WRITE_CONCURRENCY,
    (op) =>
      guarded('delete', op.uid, async () => {
        await deleteObject(ctx.fetch, ctx.target, op.href, op.etag);
        result.deleted += 1;
      }),
    halted,
  );
  return result;
}
