import { OUTLOOK_HOSTS } from '../core/constants.ts';
import type { PlanEntry, PlanView } from '../core/engine.ts';
import {
  originPattern,
  parseWholeNumber,
  prepareSave,
  validateSettings,
} from '../core/settings.ts';
import {
  type PrivatePolicy,
  type Settings,
  SHOW_AS_VALUES,
  type ShowAs,
  type TextPolicy,
} from '../core/types.ts';
import type {
  Request,
  ResetResponse,
  RunRecord,
  SaveResponse,
  StatusResponse,
  SyncResponse,
} from './messages.ts';
import { buildStatus, ignoredTokensNote } from './status.ts';
import { loadSettings } from './storage.ts';

const OUTCOME_LABEL: Record<string, string> = {
  applied: 'Synced',
  preview: 'Previewed',
  'needs-confirmation': 'Waiting for your confirmation',
  disabled: 'Switched off',
  'invalid-settings': 'Invalid settings',
  'auth-failed': 'Outlook session rejected: open Outlook',
  'read-failed': 'Could not read Outlook',
  'destination-failed': 'Could not use the calendar server',
  failed: 'Failed',
  'no-token': 'Outlook is not open',
  'not-configured': 'Not configured',
  'no-permission': 'No access to the calendar server',
  'credentials-rejected': 'Calendar server rejected the password',
  reset: 'Mirror reset',
};

const MODE_LABEL: Record<string, string> = {
  copy: 'Copy',
  drop: 'Drop',
  replace: 'Replace with',
  skip: 'Do not mirror',
};

const SHOW_AS_LABEL: Record<ShowAs, string> = {
  Free: 'Free',
  Tentative: 'Tentative',
  Busy: 'Busy',
  Oof: 'Out of office',
  WorkingElsewhere: 'Working elsewhere',
  Unknown: 'Unknown',
};

/** A run that wrote everything it meant to; one with failed writes is not a clean sync. */
const hasErrors = (record: RunRecord) => (record.counts?.errors ?? 0) > 0;

function outcomeLabel(record: RunRecord): string {
  if (record.outcome === 'applied' && hasErrors(record)) return 'Synced with errors';
  return OUTCOME_LABEL[record.outcome] ?? record.outcome;
}

const isClean = (record: RunRecord | undefined, outcome: string) =>
  record?.outcome === outcome && !hasErrors(record);

const PREVIEW_LIMIT = 200;

function el<T extends HTMLElement>(id: string): T {
  const found = document.getElementById(id);
  if (!found) throw new Error(`Missing element #${id}`);
  return found as T;
}

function send<T>(request: Request): Promise<T> {
  return chrome.runtime.sendMessage(request) as Promise<T>;
}

function setMessage(text: string, kind: 'info' | 'error' = 'info', target = 'action-message') {
  const node = el(target);
  node.textContent = text;
  node.classList.toggle('error', kind === 'error');
}

// ---- settings form -------------------------------------------------------------------------

type PolicyField = 'title' | 'description' | 'location';
const POLICY_FIELDS: PolicyField[] = ['title', 'description', 'location'];
const MODES: Record<string, string[]> = {
  attendees: ['copy', 'drop'],
  privateEvents: ['copy', 'replace', 'skip'],
};

function buildForm() {
  for (const host of OUTLOOK_HOSTS) {
    el<HTMLSelectElement>('outlookHost').append(new Option(host, host));
  }
  for (const row of document.querySelectorAll<HTMLElement>('.policy')) {
    const field = row.dataset.field as string;
    const modes = MODES[field] ?? ['copy', 'drop', 'replace'];
    const select = document.createElement('select');
    select.id = `policy-${field}`;
    for (const mode of modes) select.append(new Option(MODE_LABEL[mode], mode));
    row.append(select);
    if (field !== 'attendees') {
      const text = document.createElement('input');
      text.type = 'text';
      text.id = `policy-${field}-text`;
      text.placeholder = 'Replacement text';
      row.append(text);
      select.addEventListener('change', () => {
        text.hidden = select.value !== 'replace';
      });
      text.hidden = select.value !== 'replace';
    }
  }
  const box = el('showAsBox');
  for (const value of SHOW_AS_VALUES) {
    const label = document.createElement('label');
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.value = value;
    input.name = 'excludeShowAs';
    label.append(input, ` ${SHOW_AS_LABEL[value]}`);
    box.append(label);
  }
}

function fillForm(settings: Settings) {
  el<HTMLInputElement>('enabled').checked = settings.enabled;
  el<HTMLInputElement>('caldavUrl').value = settings.caldavUrl;
  el<HTMLInputElement>('username').value = settings.username;
  el<HTMLInputElement>('password').value = '';
  el<HTMLInputElement>('password').placeholder = settings.password
    ? 'Saved; leave empty to keep'
    : '';
  el<HTMLInputElement>('ownerId').value = settings.ownerId;
  el<HTMLInputElement>('daysBack').value = String(settings.daysBack);
  el<HTMLInputElement>('daysAhead').value = String(settings.daysAhead);
  el<HTMLInputElement>('intervalMinutes').value = String(settings.intervalMinutes);
  el<HTMLSelectElement>('outlookHost').value = settings.outlookHost;
  for (const field of POLICY_FIELDS) {
    const policy = settings.policies[field];
    const select = el<HTMLSelectElement>(`policy-${field}`);
    const text = el<HTMLInputElement>(`policy-${field}-text`);
    select.value = policy.mode;
    text.value = policy.mode === 'replace' ? policy.text : '';
    text.hidden = policy.mode !== 'replace';
  }
  el<HTMLSelectElement>('policy-attendees').value = settings.policies.attendees;
  const hidden = settings.policies.privateEvents;
  el<HTMLSelectElement>('policy-privateEvents').value = hidden.mode;
  el<HTMLInputElement>('policy-privateEvents-text').value =
    hidden.mode === 'replace' ? hidden.text : '';
  el<HTMLInputElement>('policy-privateEvents-text').hidden = hidden.mode !== 'replace';
  el<HTMLTextAreaElement>('titleFilters').value = settings.titleFilters.join('\n');
  for (const box of document.querySelectorAll<HTMLInputElement>('input[name="excludeShowAs"]')) {
    box.checked = settings.excludeShowAs.includes(box.value as ShowAs);
  }
}

function readPolicy(field: PolicyField): TextPolicy {
  const mode = el<HTMLSelectElement>(`policy-${field}`).value;
  if (mode === 'replace') return { mode, text: el<HTMLInputElement>(`policy-${field}-text`).value };
  return { mode: mode as 'copy' | 'drop' };
}

function readPrivatePolicy(): PrivatePolicy {
  const mode = el<HTMLSelectElement>('policy-privateEvents').value;
  if (mode === 'replace') {
    return { mode, text: el<HTMLInputElement>('policy-privateEvents-text').value };
  }
  return { mode: mode as 'copy' | 'skip' };
}

/** The form as typed; an empty password field stays empty (the stored one is applied on save). */
function readForm(): Settings {
  const value = (id: string) => el<HTMLInputElement>(id).value.trim();
  return {
    enabled: el<HTMLInputElement>('enabled').checked,
    caldavUrl: value('caldavUrl'),
    username: value('username'),
    password: el<HTMLInputElement>('password').value,
    ownerId: value('ownerId'),
    outlookHost: el<HTMLSelectElement>('outlookHost').value as Settings['outlookHost'],
    daysBack: parseWholeNumber(value('daysBack')),
    daysAhead: parseWholeNumber(value('daysAhead')),
    intervalMinutes: parseWholeNumber(value('intervalMinutes')),
    policies: {
      title: readPolicy('title'),
      description: readPolicy('description'),
      location: readPolicy('location'),
      attendees: el<HTMLSelectElement>('policy-attendees').value as 'copy' | 'drop',
      privateEvents: readPrivatePolicy(),
    },
    titleFilters: el<HTMLTextAreaElement>('titleFilters')
      .value.split('\n')
      .map((line) => line.trim())
      .filter((line) => line !== ''),
    excludeShowAs: [
      ...document.querySelectorAll<HTMLInputElement>('input[name="excludeShowAs"]:checked'),
    ].map((box) => box.value as ShowAs),
  };
}

let current: Settings;

async function save() {
  const draft = readForm();
  const prepared = prepareSave(current, draft);
  const errors = 'errors' in prepared ? prepared.errors : validateSettings(prepared.settings);
  if (errors.length > 0) {
    setMessage(errors.join('. '), 'error', 'save-message');
    return;
  }
  const origins = [originPattern(draft.caldavUrl) as string];
  const alreadyHeld = await chrome.permissions.contains({ origins });
  // Requested first: the browser only grants it inside the click's user gesture.
  const granted = await chrome.permissions.request({ origins });
  if (!granted) {
    setMessage(
      'Access to the calendar server was not granted, nothing saved',
      'error',
      'save-message',
    );
    return;
  }
  setMessage('Checking the calendar and saving…', 'info', 'save-message');
  const response = await send<SaveResponse>({ type: 'save', settings: draft });
  if (!response.saved) {
    // A refused save does not keep the access it asked for.
    if (!alreadyHeld) await chrome.permissions.remove({ origins });
    setMessage(
      `Nothing saved: ${(response.errors ?? ['unknown error']).join('. ')}`,
      'error',
      'save-message',
    );
    return;
  }
  current = (await loadSettings()).settings;
  fillForm(current);
  const sync = response.sync;
  const outcome = sync && !sync.busy ? describe(sync) : 'A sync was already running';
  setMessage(`Saved. ${response.connection ?? ''} ${outcome}`.trim(), 'info', 'save-message');
  if (sync) showSyncResult(sync);
}

async function onSave(event: Event) {
  event.preventDefault();
  await busyWhile(async () => {
    try {
      await save();
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'Failed', 'error', 'save-message');
    }
  });
}

// ---- status ---------------------------------------------------------------------------------

const when = (iso: string) => new Date(iso).toLocaleString();

function countsLine(record: RunRecord): string {
  const c = record.counts;
  if (!c) return '';
  if (record.outcome === 'applied') {
    return `${c.created} created, ${c.updated} updated, ${c.deleted} deleted, ${c.unchanged} unchanged${c.errors ? `, ${c.errors} failed` : ''}`;
  }
  return `${c.create} to create, ${c.update} to update, ${c.delete} to delete, ${c.unchanged} unchanged`;
}

function addRow(list: HTMLElement, label: string, value: string) {
  const dt = document.createElement('dt');
  dt.textContent = label;
  const dd = document.createElement('dd');
  dd.textContent = value;
  list.append(dt, dd);
}

let pendingDigest: string | undefined;

function renderStatus(status: StatusResponse) {
  const list = el('status');
  list.replaceChildren();
  addRow(
    list,
    'Syncing',
    status.configured ? (status.enabled ? 'on' : 'switched off') : 'not configured',
  );
  addRow(
    list,
    'Mirroring account',
    status.account
      ? `${status.account.label}${status.account.pinned ? '' : ' (bound after the first successful sync)'}`
      : 'not seen yet: open Outlook in this browser',
  );
  if (status.ignoredTokens > 0) addRow(list, 'Ignored', ignoredTokensNote(status.ignoredTokens));
  addRow(
    list,
    'Outlook session',
    status.token.usable
      ? `seen, valid until ${when(status.token.expiresAt as string)}`
      : 'none seen or expired: open Outlook in this browser',
  );
  const run = status.lastRun;
  if (run) {
    addRow(list, 'Last run', `${when(run.finishedAt)}: ${outcomeLabel(run)}`);
    const counts = countsLine(run);
    if (counts) addRow(list, 'Changes', counts);
    if (run.counts) {
      addRow(
        list,
        'Calendar',
        `${run.counts.own} mirrored, ${run.counts.foreign} other objects left alone, ${run.counts.source} events read from Outlook`,
      );
    }
    if (run.message) addRow(list, 'Message', run.message);
  } else {
    addRow(list, 'Last run', 'never');
  }
  if (status.lastSuccessAt) addRow(list, 'Last success', when(status.lastSuccessAt));
  if (status.retryAt) {
    addRow(list, 'Next attempt', `${when(status.retryAt)} (the same failure keeps repeating)`);
  }

  const banner = el('banner');
  const waiting = run?.outcome === 'needs-confirmation';
  banner.hidden = !waiting && !status.passwordRejected;
  banner.className = 'banner warn';
  banner.replaceChildren();
  if (status.passwordRejected) {
    const text = document.createElement('p');
    text.textContent =
      'The calendar server rejected the password. Nothing is retried until you save the settings again.';
    banner.append(text);
  }
  if (waiting) {
    const text = document.createElement('p');
    text.textContent = `The last run was held back: ${run.guard?.message ?? run.message ?? ''}. Nothing was changed.`;
    const review = document.createElement('button');
    review.type = 'button';
    review.textContent = 'Review changes';
    review.addEventListener('click', () => void onPreview());
    banner.append(text, review);
  }
  renderHistory(status.log);
}

function renderHistory(log: RunRecord[]) {
  const list = el('history-list');
  list.replaceChildren();
  for (const record of log) {
    const parts = [
      when(record.finishedAt),
      outcomeLabel(record) + (record.repeats > 1 ? ` (${record.repeats} times)` : ''),
      countsLine(record),
      record.message,
    ];
    const item = document.createElement('li');
    item.textContent = parts.filter(Boolean).join(': ');
    list.append(item);
  }
  el('history').hidden = log.length === 0;
}

async function refreshStatus() {
  renderStatus(await buildStatus());
}

// ---- plan preview ---------------------------------------------------------------------------

function formatStart(raw: string): string {
  const compact = /^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})\d{2}Z)?$/.exec(raw);
  if (compact) {
    const [, y, m, d, hh, mm] = compact;
    return hh ? `${y}-${m}-${d} ${hh}:${mm} UTC` : `${y}-${m}-${d}`;
  }
  const iso = /^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})/.exec(raw);
  return iso ? `${iso[1]} ${iso[2]} UTC` : raw;
}

function group(title: string, entries: PlanEntry[]): HTMLElement | undefined {
  if (entries.length === 0) return undefined;
  const details = document.createElement('details');
  const summary = document.createElement('summary');
  summary.textContent = `${title} (${entries.length})`;
  const list = document.createElement('ul');
  for (const entry of entries.slice(0, PREVIEW_LIMIT)) {
    const item = document.createElement('li');
    item.textContent = `${formatStart(entry.start)}  ${entry.summary}`;
    list.append(item);
  }
  if (entries.length > PREVIEW_LIMIT) {
    const more = document.createElement('li');
    more.textContent = `… and ${entries.length - PREVIEW_LIMIT} more`;
    list.append(more);
  }
  details.append(summary, list);
  return details;
}

function renderPlan(plan: PlanView, open: 'deletions' | 'creations' | undefined) {
  const box = el('plan');
  box.replaceChildren();
  const groups = [
    group('Create', plan.creations),
    group('Update', plan.updates),
    group('Delete', plan.deletions),
    group('Blocked by an object that is not owned', plan.conflicts),
  ];
  for (const g of groups) if (g) box.append(g);
  const sections = box.querySelectorAll('details');
  const focus = open === 'deletions' ? 'Delete' : open === 'creations' ? 'Create' : undefined;
  for (const d of sections) if (focus && d.textContent?.startsWith(focus)) d.open = true;
  if (sections.length === 0) box.textContent = 'Nothing to change.';
  box.hidden = false;
}

function showConfirm(record: RunRecord | undefined) {
  const box = el('confirm-box');
  pendingDigest = record?.guard && record.digest ? record.digest : undefined;
  box.hidden = pendingDigest === undefined;
  if (record?.guard) {
    el('confirm-text').textContent = `${record.guard.message}. Apply this plan anyway?`;
  }
}

// ---- actions --------------------------------------------------------------------------------

const BUTTONS = ['sync', 'preview', 'reset', 'reset-other', 'confirm', 'save', 'forget'];

async function busyWhile(work: () => Promise<void>) {
  for (const id of BUTTONS) el<HTMLButtonElement>(id).disabled = true;
  try {
    await work();
  } catch (error) {
    setMessage(error instanceof Error ? error.message : 'Failed', 'error');
  } finally {
    for (const id of BUTTONS) el<HTMLButtonElement>(id).disabled = false;
    await refreshStatus();
  }
}

function describe(response: SyncResponse): string {
  if (response.busy) return 'A sync is already running. Try again in a moment.';
  const record = response.record as RunRecord;
  const parts = [outcomeLabel(record), countsLine(record), record.message];
  return parts.filter(Boolean).join(': ');
}

/** The group of the plan the held-back reason is about. */
function planFocus(record: RunRecord | undefined): 'deletions' | 'creations' | undefined {
  const code = record?.guard?.code;
  if (code === 'other-owner' || code === 'foreign-destination') return 'creations';
  return code ? 'deletions' : undefined;
}

/** Shows what a sync that was started from the page left behind: plan and confirmation. */
function showSyncResult(response: SyncResponse) {
  if (response.plan) renderPlan(response.plan, planFocus(response.record));
  if (response.record?.outcome === 'needs-confirmation') {
    showConfirm(response.record);
    el('status-panel').scrollIntoView();
  }
}

async function onSync() {
  await busyWhile(async () => {
    el('plan').hidden = true;
    showConfirm(undefined);
    setMessage('Syncing…');
    const response = await send<SyncResponse>({ type: 'sync', mode: 'apply' });
    setMessage(describe(response), isClean(response.record, 'applied') ? 'info' : 'error');
    showSyncResult(response);
  });
}

async function onPreview() {
  await busyWhile(async () => {
    showConfirm(undefined);
    setMessage('Reading Outlook and the calendar…');
    const response = await send<SyncResponse>({ type: 'sync', mode: 'preview' });
    setMessage(describe(response), isClean(response.record, 'preview') ? 'info' : 'error');
    if (response.plan) renderPlan(response.plan, planFocus(response.record));
    showConfirm(response.record);
  });
}

async function onConfirm() {
  const digest = pendingDigest;
  if (!digest) return;
  await busyWhile(async () => {
    setMessage('Applying…');
    const response = await send<SyncResponse>({ type: 'sync', mode: 'apply', confirm: { digest } });
    setMessage(describe(response), isClean(response.record, 'applied') ? 'info' : 'error');
    // A busy answer changed nothing: the confirmation stays available.
    if (response.busy) return;
    showConfirm(response.record?.outcome === 'needs-confirmation' ? response.record : undefined);
    if (response.record?.outcome === 'applied') el('plan').hidden = true;
  });
}

async function onForgetAccount() {
  await busyWhile(async () => {
    await send({ type: 'forget-account' });
    setMessage('Released. The next Outlook session the page shows is mirrored.');
  });
}

/** Counts, asks, then deletes the events of one owner id (the configured one when omitted). */
async function resetOwner(
  ownerId: string | undefined,
  question: (own: number, foreign: number) => string,
) {
  const owner = ownerId === undefined ? {} : { ownerId };
  const counted = await send<ResetResponse>({ type: 'reset', mode: 'count', ...owner });
  const result = counted.result;
  if (counted.busy || !result || result.outcome !== 'counted') {
    setMessage(result?.message ?? 'A sync is already running. Try again in a moment.', 'error');
    return;
  }
  if (result.own === 0 || !window.confirm(question(result.own, result.foreign))) {
    setMessage(result.own === 0 ? 'Nothing to delete.' : 'Cancelled.');
    return;
  }
  const done = await send<ResetResponse>({
    type: 'reset',
    mode: 'apply',
    confirmedCount: result.own,
    ...owner,
  });
  const r = done.result;
  setMessage(
    r
      ? `${r.outcome}: deleted ${r.deleted}${r.errors.length ? `, ${r.errors.length} failed` : ''}${r.message ? ` (${r.message})` : ''}`
      : 'Busy',
    r?.outcome === 'reset' ? 'info' : 'error',
  );
}

async function onReset() {
  await busyWhile(() =>
    resetOwner(
      undefined,
      (own, foreign) =>
        `Delete all ${own} mirrored events from the calendar? ${foreign} other objects are left alone. While syncing is on, the next run recreates the mirror (and asks for confirmation if other objects remain): switch syncing off first to stop mirroring.`,
    ),
  );
}

async function onResetOther() {
  await busyWhile(async () => {
    const overview = await send<ResetResponse>({ type: 'reset', mode: 'count' });
    if (overview.busy || !overview.result) {
      setMessage('A sync is already running. Try again in a moment.', 'error');
      return;
    }
    const owners = overview.result.otherOwners;
    if (owners.length === 0) {
      setMessage('No events of another owner id found.');
      return;
    }
    const listing = owners.map((o) => `${o.id} (${o.count})`).join(', ');
    const chosen =
      owners.length === 1
        ? owners[0]?.id
        : window.prompt(`Remove the events of which owner id? ${listing}`, owners[0]?.id)?.trim();
    if (!chosen || !owners.some((o) => o.id === chosen)) {
      setMessage('Cancelled.');
      return;
    }
    await resetOwner(
      chosen,
      (own) =>
        `Delete the ${own} events of owner id ${chosen}? Objects without any mirror owner id, and those of other owner ids, are left alone.`,
    );
  });
}

let refreshing = false;

/** Redraws the status whenever the stored state changes; several changes at once draw once. */
function watchStorage() {
  chrome.storage.onChanged.addListener(() => {
    if (refreshing) return;
    refreshing = true;
    queueMicrotask(() => {
      refreshing = false;
      void refreshStatus();
    });
  });
}

async function init() {
  buildForm();
  current = (await loadSettings()).settings;
  fillForm(current);
  el('settings').addEventListener('submit', (e) => void onSave(e));
  el('sync').addEventListener('click', () => void onSync());
  el('preview').addEventListener('click', () => void onPreview());
  el('confirm').addEventListener('click', () => void onConfirm());
  el('reset').addEventListener('click', () => void onReset());
  el('forget').addEventListener('click', () => void onForgetAccount());
  el('reset-other').addEventListener('click', () => void onResetOther());
  await refreshStatus();
  watchStorage();
}

void init();
