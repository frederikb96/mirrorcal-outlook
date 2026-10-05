import { OUTLOOK_HOSTS } from './constants.ts';
import { type Settings, SHOW_AS_VALUES } from './types.ts';

export const LIMITS = {
  daysBack: { min: 0, max: 730 },
  daysAhead: { min: 0, max: 730 },
  intervalMinutes: { min: 5, max: 1440 },
} as const;

/** The one place defaults live. The engine never fills gaps; callers merge with these first. */
export const DEFAULT_SETTINGS: Settings = {
  enabled: true,
  caldavUrl: '',
  username: '',
  password: '',
  ownerId: 'outlook',
  outlookHost: 'outlook.cloud.microsoft',
  daysBack: 30,
  daysAhead: 180,
  intervalMinutes: 15,
  policies: {
    title: { mode: 'copy' },
    description: { mode: 'drop' },
    location: { mode: 'drop' },
    attendees: 'drop',
    privateEvents: { mode: 'replace', text: 'Private' },
  },
  titleFilters: [],
  excludeShowAs: [],
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/** Overlays stored (possibly partial or foreign-shaped) settings on the defaults. */
export function mergeWithDefaults(stored: unknown): Settings {
  const base = structuredClone(DEFAULT_SETTINGS);
  if (!isRecord(stored)) return base;
  const { policies, ...rest } = stored;
  const merged = { ...base, ...rest } as Settings;
  merged.policies = { ...base.policies, ...(isRecord(policies) ? policies : {}) } as never;
  return merged;
}

export function normalizeCaldavUrl(url: string): string {
  const trimmed = url.trim();
  return trimmed.endsWith('/') ? trimmed : `${trimmed}/`;
}

/** The origin of a URL, or undefined when it is not one. */
function originOf(url: string): string | undefined {
  try {
    return new URL(url).origin;
  } catch {
    return undefined;
  }
}

/** Whole-number text from a form field; empty or blank text is not a number. */
export function parseWholeNumber(text: string): number {
  return text.trim() === '' ? Number.NaN : Number(text);
}

function isLoopback(hostname: string): boolean {
  return hostname === 'localhost' || hostname === '127.0.0.1' || hostname === '[::1]';
}

function checkInteger(
  errors: string[],
  label: string,
  value: unknown,
  range: { min: number; max: number },
) {
  if (!Number.isInteger(value) || (value as number) < range.min || (value as number) > range.max) {
    errors.push(`${label} must be a whole number from ${range.min} to ${range.max}`);
  }
}

function checkTextPolicy(errors: string[], label: string, policy: unknown) {
  if (!isRecord(policy)) {
    errors.push(`${label} policy is missing`);
    return;
  }
  const { mode } = policy;
  if (mode === 'replace') {
    if (typeof policy.text !== 'string') errors.push(`${label} replacement text is missing`);
  } else if (mode !== 'copy' && mode !== 'drop') {
    errors.push(`${label} policy must be copy, drop or replace`);
  }
}

function checkPrivatePolicy(errors: string[], policy: unknown) {
  if (!isRecord(policy)) {
    errors.push('Private events policy is missing');
  } else if (policy.mode === 'replace') {
    if (typeof policy.text !== 'string') errors.push('Private events replacement text is missing');
  } else if (policy.mode !== 'copy' && policy.mode !== 'skip') {
    errors.push('Private events policy must be copy, replace or skip');
  }
}

/** Problems that make the destination unusable (also what a mirror reset needs). */
export const OWNER_ID_PATTERN = /^[A-Za-z0-9._-]{1,64}$/;

export function validateDestination(settings: Settings): string[] {
  const errors: string[] = [];
  if (typeof settings.caldavUrl !== 'string' || settings.caldavUrl.trim() === '') {
    errors.push('Calendar URL is required');
  } else {
    try {
      const url = new URL(settings.caldavUrl);
      if (url.username || url.password) errors.push('Calendar URL must not contain credentials');
      if (url.protocol !== 'https:' && !(url.protocol === 'http:' && isLoopback(url.hostname))) {
        errors.push('Calendar URL must use https');
      }
      // `host.` and `host%2e` are the same host as `host`.
      if (
        (OUTLOOK_HOSTS as readonly string[]).includes(
          url.hostname.replace(/\.+$/, '').toLowerCase(),
        )
      ) {
        errors.push('Calendar URL must not point at Outlook');
      }
    } catch {
      errors.push('Calendar URL is not a valid URL');
    }
  }
  if (typeof settings.username !== 'string' || settings.username === '') {
    errors.push('User name is required');
  }
  if (typeof settings.password !== 'string' || settings.password === '') {
    errors.push('Password is required');
  }
  if (typeof settings.ownerId !== 'string' || !OWNER_ID_PATTERN.test(settings.ownerId)) {
    errors.push('Owner id must be 1-64 letters, digits, dot, dash or underscore');
  }
  return errors;
}

export function validateSettings(settings: Settings): string[] {
  const errors = validateDestination(settings);
  if (typeof settings.enabled !== 'boolean') errors.push('enabled must be true or false');
  if (!(OUTLOOK_HOSTS as readonly string[]).includes(settings.outlookHost)) {
    errors.push(`Outlook host must be one of ${OUTLOOK_HOSTS.join(', ')}`);
  }
  checkInteger(errors, 'Days back', settings.daysBack, LIMITS.daysBack);
  checkInteger(errors, 'Days ahead', settings.daysAhead, LIMITS.daysAhead);
  checkInteger(errors, 'Interval', settings.intervalMinutes, LIMITS.intervalMinutes);
  const policies = settings.policies;
  if (!isRecord(policies)) {
    errors.push('Policies are missing');
  } else {
    checkTextPolicy(errors, 'Title', policies.title);
    checkTextPolicy(errors, 'Description', policies.description);
    checkTextPolicy(errors, 'Location', policies.location);
    if (policies.attendees !== 'copy' && policies.attendees !== 'drop') {
      errors.push('Attendees policy must be copy or drop');
    }
    checkPrivatePolicy(errors, policies.privateEvents);
  }
  if (
    !Array.isArray(settings.titleFilters) ||
    settings.titleFilters.some((f) => typeof f !== 'string')
  ) {
    errors.push('Title filters must be a list of text patterns');
  }
  if (
    !Array.isArray(settings.excludeShowAs) ||
    settings.excludeShowAs.some((s) => !SHOW_AS_VALUES.includes(s))
  ) {
    errors.push('Excluded availabilities must be a list of known values');
  }
  return errors;
}

/**
 * Turns a form draft into the settings to store. An empty password field keeps the stored one,
 * but only for the server it was entered for: a draft that points at another origin must bring
 * its own password.
 */
export function prepareSave(
  existing: Settings,
  draft: Settings,
): { settings: Settings } | { errors: string[] } {
  if (draft.password !== '') return { settings: draft };
  const sameServer =
    existing.password !== '' &&
    originOf(draft.caldavUrl) !== undefined &&
    originOf(draft.caldavUrl) === originOf(existing.caldavUrl);
  if (sameServer) return { settings: { ...draft, password: existing.password } };
  return {
    errors: [
      existing.password === ''
        ? 'Password is required'
        : 'Enter the password for the new calendar server',
    ],
  };
}

/** The origin permission a destination URL needs, as a match pattern. */
export function originPattern(url: string): string | undefined {
  const origin = originOf(url);
  return origin === undefined ? undefined : `${origin}/*`;
}
