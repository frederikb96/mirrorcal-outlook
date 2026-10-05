import { sha256Hex } from './crypto.ts';
import { contentHash, readRemoteInfo } from './ics.ts';
import type { DesiredObject, RemoteObject } from './types.ts';

export interface OwnedObject {
  href: string;
  etag: string;
  uid: string;
  /** Hash of the content as it stands on the server, not of anything remembered. */
  hash: string;
  summary: string;
  start: string;
}

export interface ForeignObject {
  href: string;
  uid?: string;
  summary?: string;
  /** Owner id of another mirror, when the object carries a marker with a different value. */
  owner?: string;
}

export interface CreateOp {
  kind: 'create';
  uid: string;
  href: string;
  ics: string;
  hash: string;
  summary: string;
  start: string;
  allDay: boolean;
}

export interface UpdateOp extends Omit<CreateOp, 'kind'> {
  kind: 'update';
  etag: string;
}

export interface DeleteOp {
  kind: 'delete';
  uid: string;
  href: string;
  etag: string;
  summary: string;
  start: string;
}

export interface Plan {
  creations: CreateOp[];
  updates: UpdateOp[];
  deletions: DeleteOp[];
  unchanged: number;
  /** Desired objects that cannot be created because a foreign object occupies the identity. */
  conflicts: CreateOp[];
  /** Objects carrying this owner's marker. */
  ownCount: number;
  /** Addresses of those objects: the only ones a write may target. */
  ownHrefs: string[];
  /** Objects without it: never modified, only reported. */
  foreign: ForeignObject[];
}

/** Splits the destination into objects this owner may touch and everything else. */
export async function classifyRemote(
  remote: RemoteObject[],
  ownerId: string,
): Promise<{ own: OwnedObject[]; foreign: ForeignObject[] }> {
  const own: OwnedObject[] = [];
  const foreign: ForeignObject[] = [];
  for (const object of remote) {
    const info = readRemoteInfo(object.ics);
    if (info.owner === ownerId && info.uid) {
      own.push({
        href: object.href,
        etag: object.etag,
        uid: info.uid,
        hash: await contentHash(object.ics),
        summary: info.summary ?? '',
        start: info.start ?? '',
      });
    } else {
      const entry: ForeignObject = { href: object.href };
      if (info.uid) entry.uid = info.uid;
      if (info.summary) entry.summary = info.summary;
      if (info.owner && info.owner !== ownerId) entry.owner = info.owner;
      foreign.push(entry);
    }
  }
  return { own, foreign };
}

/**
 * Pure comparison of what should exist with what does: no I/O, no remembered state. Only objects
 * that carry the owner marker can appear as an update or a deletion.
 */
export async function planSync(
  desired: DesiredObject[],
  remote: RemoteObject[],
  ownerId: string,
): Promise<Plan> {
  const { own, foreign } = await classifyRemote(remote, ownerId);
  const ownByUid = new Map<string, OwnedObject>();
  const surplus: OwnedObject[] = [];
  // Of several objects with one identity, the one at the canonical address survives.
  const canonical = (o: OwnedObject) =>
    o.href.endsWith(`/${encodeURIComponent(o.uid)}.ics`) ? 0 : 1;
  const ranked = [...own].sort(
    (a, b) => canonical(a) - canonical(b) || a.href.localeCompare(b.href),
  );
  for (const object of ranked) {
    if (ownByUid.has(object.uid)) surplus.push(object);
    else ownByUid.set(object.uid, object);
  }
  const occupied = new Set([
    ...foreign.map((f) => f.href),
    ...foreign.flatMap((f) => (f.uid ? [f.uid] : [])),
  ]);

  const plan: Plan = {
    creations: [],
    updates: [],
    deletions: [],
    unchanged: 0,
    conflicts: [],
    ownCount: own.length,
    ownHrefs: own.map((o) => o.href),
    foreign,
  };
  const wanted = new Set<string>();
  for (const item of desired) {
    wanted.add(item.uid);
    const fields = {
      uid: item.uid,
      href: item.href,
      ics: item.ics,
      hash: item.hash,
      summary: item.summary,
      start: item.start,
      allDay: item.allDay,
    };
    const existing = ownByUid.get(item.uid);
    if (!existing) {
      if (occupied.has(item.uid) || occupied.has(item.href)) {
        plan.conflicts.push({ kind: 'create', ...fields });
      } else {
        plan.creations.push({ kind: 'create', ...fields });
      }
    } else if (existing.hash === item.hash) {
      plan.unchanged += 1;
    } else {
      plan.updates.push({ kind: 'update', ...fields, href: existing.href, etag: existing.etag });
    }
  }
  const toDelete = [...ownByUid.values()].filter((o) => !wanted.has(o.uid)).concat(surplus);
  plan.deletions = toDelete
    .sort((a, b) => a.href.localeCompare(b.href))
    .map((o) => ({
      kind: 'delete',
      uid: o.uid,
      href: o.href,
      etag: o.etag,
      summary: o.summary,
      start: o.start,
    }));
  return plan;
}

/** Distinct owner ids of other mirrors' objects, in order of appearance. */
export function otherOwnersOf(foreign: ForeignObject[]): string[] {
  return [...new Set(foreign.flatMap((f) => (f.owner ? [f.owner] : [])))];
}

/** Fingerprint of what a plan would do; a confirmation is only valid for the plan it was given for. */
export async function planDigest(plan: Plan): Promise<string> {
  const lines = [
    ...plan.creations.map((o) => `create ${o.uid} ${o.hash}`),
    ...plan.updates.map((o) => `update ${o.uid} ${o.hash}`),
    ...plan.deletions.map((o) => `delete ${o.href}`),
  ].sort();
  return sha256Hex(lines.join('\n'));
}

/** Fingerprint of everything a sync would want to exist, for deciding that nothing needs doing. */
export async function desiredDigest(
  desired: DesiredObject[],
  collectionUrl: string,
): Promise<string> {
  const lines = desired.map((d) => `${d.uid} ${d.hash}`).sort();
  return sha256Hex([collectionUrl, ...lines].join('\n'));
}
