import { describe, expect, test } from 'bun:test';
import {
  DEFAULT_SETTINGS,
  mergeWithDefaults,
  normalizeCaldavUrl,
  parseWholeNumber,
  prepareSave,
  validateDestination,
  validateSettings,
} from '../src/core/settings.ts';
import { testSettings } from './helpers/fixtures.ts';

describe('defaults', () => {
  test('match the documented behaviour', () => {
    expect(DEFAULT_SETTINGS.daysBack).toBe(30);
    expect(DEFAULT_SETTINGS.daysAhead).toBe(180);
    expect(DEFAULT_SETTINGS.intervalMinutes).toBe(15);
    expect(DEFAULT_SETTINGS.policies).toEqual({
      title: { mode: 'copy' },
      description: { mode: 'drop' },
      location: { mode: 'drop' },
      attendees: 'drop',
      privateEvents: { mode: 'replace', text: 'Private' },
    });
  });

  test('merging fills gaps without sharing state with the defaults', () => {
    const merged = mergeWithDefaults({ daysBack: 7, policies: { title: { mode: 'drop' } } });
    expect(merged.daysBack).toBe(7);
    expect(merged.daysAhead).toBe(180);
    expect(merged.policies.title).toEqual({ mode: 'drop' });
    expect(merged.policies.description).toEqual({ mode: 'drop' });
    merged.titleFilters.push('x');
    expect(DEFAULT_SETTINGS.titleFilters).toEqual([]);
    expect(mergeWithDefaults(undefined)).toEqual(DEFAULT_SETTINGS);
    expect(mergeWithDefaults('junk')).toEqual(DEFAULT_SETTINGS);
  });
});

describe('validation', () => {
  test('a complete configuration is valid', () => {
    expect(validateSettings(testSettings())).toEqual([]);
  });

  test('the destination must be configured', () => {
    expect(validateSettings(mergeWithDefaults({})).length).toBeGreaterThanOrEqual(3);
    expect(validateDestination(testSettings({ password: '' }))).toEqual(['Password is required']);
  });

  test('plain http is only allowed on loopback', () => {
    expect(
      validateDestination(testSettings({ caldavUrl: 'http://dav.example.test/x/' })),
    ).not.toEqual([]);
    expect(validateDestination(testSettings({ caldavUrl: 'http://localhost:8080/x/' }))).toEqual(
      [],
    );
    expect(
      validateDestination(testSettings({ caldavUrl: 'https://u:p@dav.example.test/x/' })),
    ).not.toEqual([]);
    expect(validateDestination(testSettings({ caldavUrl: 'dav.example.test' }))).not.toEqual([]);
  });

  test('numbers and enums are range checked', () => {
    expect(validateSettings(testSettings({ daysBack: -1 }))).not.toEqual([]);
    expect(validateSettings(testSettings({ daysAhead: 1.5 }))).not.toEqual([]);
    expect(validateSettings(testSettings({ intervalMinutes: 1 }))).not.toEqual([]);
    expect(validateSettings(testSettings({ ownerId: 'bad id' }))).not.toEqual([]);
    expect(
      validateSettings(testSettings({ outlookHost: 'evil.example.test' as never })),
    ).not.toEqual([]);
    expect(validateSettings(testSettings({ excludeShowAs: ['Nope' as never] }))).not.toEqual([]);
    const s = testSettings();
    s.policies.title = { mode: 'replace' } as never;
    expect(validateSettings(s)).not.toEqual([]);
  });

  test('urls get a trailing slash', () => {
    expect(normalizeCaldavUrl(' https://h/x ')).toBe('https://h/x/');
    expect(normalizeCaldavUrl('https://h/x/')).toBe('https://h/x/');
  });
});

describe('destination safety', () => {
  test('the Outlook hosts are not a calendar server', () => {
    for (const host of ['outlook.office.com', 'outlook.office365.com', 'outlook.cloud.microsoft']) {
      expect(validateDestination(testSettings({ caldavUrl: `https://${host}/x/` }))).toContain(
        'Calendar URL must not point at Outlook',
      );
    }
  });
});

describe('saving a draft', () => {
  const existing = testSettings();

  test('an empty password keeps the stored one for the same server', () => {
    const result = prepareSave(existing, testSettings({ password: '', username: 'bob' }));
    expect('settings' in result && result.settings.password).toBe(existing.password);
  });

  test('an empty password is refused for another origin', () => {
    const result = prepareSave(
      existing,
      testSettings({ password: '', caldavUrl: 'https://other.example.test/dav/' }),
    );
    expect('errors' in result && result.errors[0]).toContain('new calendar server');
  });

  test('another origin with its own password is accepted', () => {
    const draft = testSettings({ password: 'new-one', caldavUrl: 'https://other.example.test/d/' });
    expect(prepareSave(existing, draft)).toEqual({ settings: draft });
  });

  test('a password is required when none is stored', () => {
    const result = prepareSave(testSettings({ password: '' }), testSettings({ password: '' }));
    expect('errors' in result && result.errors).toEqual(['Password is required']);
  });
});

describe('form numbers', () => {
  test('empty or blank text is not a number', () => {
    expect(Number.isNaN(parseWholeNumber(''))).toBe(true);
    expect(Number.isNaN(parseWholeNumber('  '))).toBe(true);
    expect(parseWholeNumber('0')).toBe(0);
    expect(parseWholeNumber(' 30 ')).toBe(30);
    expect(validateSettings(testSettings({ daysBack: parseWholeNumber('') }))).toHaveLength(1);
  });
});

describe('private events policy', () => {
  test('accepts the three modes and refuses others', () => {
    for (const privateEvents of [
      { mode: 'copy' },
      { mode: 'skip' },
      { mode: 'replace', text: 'Private' },
    ] as const) {
      const s = testSettings();
      s.policies.privateEvents = privateEvents;
      expect(validateSettings(s)).toEqual([]);
    }
    const bad = testSettings();
    bad.policies.privateEvents = { mode: 'drop' } as never;
    expect(validateSettings(bad)).toHaveLength(1);
  });
});
