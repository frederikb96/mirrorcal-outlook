import { describe, expect, test } from 'bun:test';
import { problems, sections } from '../scripts/release-notes.ts';

const CHANGELOG = `# Changelog

## [Unreleased]

### Added

- Something pending.

## [1.2.0] - 2026-02-03

### Fixed

- A fix.

## [1.1.0]

## [1.0.0] - 2026-01-01

### Added

- First release.
`;

describe('changelog sections', () => {
  test('splits at version headings and keeps the body', () => {
    const found = sections(CHANGELOG);
    expect(found.map((s) => s.version)).toEqual(['Unreleased', '1.2.0', '1.1.0', '1.0.0']);
    expect(found[1]?.date).toBe('2026-02-03');
    expect(found[1]?.body).toBe('### Fixed\n\n- A fix.');
  });
});

describe('release guard', () => {
  test('passes for a dated, non-empty section and a matching tag', () => {
    expect(problems(CHANGELOG, '1.2.0', 'v1.2.0')).toEqual([]);
    expect(problems(CHANGELOG, '1.2.0')).toEqual([]);
  });

  test('fails when the tag names another version than package.json', () => {
    expect(problems(CHANGELOG, '1.2.0', 'v1.3.0')).toHaveLength(1);
  });

  test('fails for a missing, undated or empty section', () => {
    expect(problems(CHANGELOG, '2.0.0')[0]).toContain('no section');
    expect(problems(CHANGELOG, '1.1.0')).toHaveLength(2);
  });
});
