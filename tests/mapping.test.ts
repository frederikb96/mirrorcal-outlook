import { describe, expect, test } from 'bun:test';
import { unfoldLines } from '../src/core/ics.ts';
import {
  buildDesired,
  compileTitleFilter,
  exclusionOf,
  mapEvent,
  mirrorUid,
  type TitleMatcher,
} from '../src/core/mapping.ts';
import type { SourceEvent } from '../src/core/types.ts';
import { COLLECTION, NOW, outlookId, testSettings } from './helpers/fixtures.ts';

function source(overrides: Partial<SourceEvent> = {}): SourceEvent {
  return {
    id: outlookId(1),
    subject: 'Planning',
    start: '2026-10-06T08:00:00',
    end: '2026-10-06T09:00:00',
    allDay: false,
    showAs: 'Busy',
    sensitivity: 'Normal',
    location: 'Room 4',
    body: 'Agenda',
    organizer: { name: 'Pat Doe', address: 'pat@example.test' },
    attendees: [
      { name: 'Sam Roe', address: 'sam@example.test', type: 'Required', response: 'Accepted' },
    ],
    ...overrides,
  };
}

describe('time mapping', () => {
  test('timed events keep their UTC instants', () => {
    const m = mapEvent(source(), testSettings());
    expect([m.allDay, m.start, m.end]).toEqual([
      false,
      '2026-10-06T08:00:00',
      '2026-10-06T09:00:00',
    ]);
  });

  test('all-day events reduce to dates, multi-day end stays exclusive', () => {
    const m = mapEvent(
      source({ allDay: true, start: '2026-10-12T00:00:00', end: '2026-10-15T00:00:00' }),
      testSettings(),
    );
    expect([m.allDay, m.start, m.end]).toEqual([true, '2026-10-12', '2026-10-15']);
  });

  test('a zero-length all-day event spans one day', () => {
    const m = mapEvent(
      source({ allDay: true, start: '2026-12-31T00:00:00', end: '2026-12-31T00:00:00' }),
      testSettings(),
    );
    expect(m.end).toBe('2027-01-01');
  });
});

describe('field policies', () => {
  test('defaults copy the title and drop the rest', () => {
    const m = mapEvent(source(), testSettings());
    expect(m.summary).toBe('Planning');
    expect(m.description).toBeUndefined();
    expect(m.location).toBeUndefined();
    expect(m.attendees).toEqual([]);
    expect(m.organizer).toBeUndefined();
  });

  test('copy includes the field', () => {
    const s = testSettings();
    s.policies.description = { mode: 'copy' };
    s.policies.location = { mode: 'copy' };
    s.policies.attendees = 'copy';
    const m = mapEvent(source(), s);
    expect(m.description).toBe('Agenda');
    expect(m.location).toBe('Room 4');
    expect(m.attendees).toHaveLength(1);
    expect(m.organizer?.address).toBe('pat@example.test');
  });

  test('replace substitutes fixed text where the source has a value', () => {
    const s = testSettings();
    s.policies.title = { mode: 'replace', text: 'Meeting' };
    s.policies.location = { mode: 'replace', text: 'Office' };
    s.policies.description = { mode: 'replace', text: 'Details in Outlook' };
    const m = mapEvent(source({ location: 'Room 4', body: 'Agenda' }), s);
    expect(m.summary).toBe('Meeting');
    expect(m.location).toBe('Office');
    expect(m.description).toBe('Details in Outlook');
  });

  test('replace leaves a field empty when the source has none', () => {
    const s = testSettings();
    s.policies.location = { mode: 'replace', text: 'Office' };
    s.policies.description = { mode: 'replace', text: 'Details in Outlook' };
    for (const src of [
      source({ location: undefined, body: undefined }),
      source({ location: '  ', body: '\r\n' }),
    ]) {
      const m = mapEvent(src, s);
      expect(m.location).toBeUndefined();
      expect(m.description).toBeUndefined();
    }
  });

  test('the summary is never empty', () => {
    const s = testSettings();
    expect(mapEvent(source({ subject: '' }), s).summary).toBe('Busy');
    expect(mapEvent(source({ subject: '   ' }), s).summary).toBe('Busy');
    s.policies.title = { mode: 'drop' };
    expect(mapEvent(source(), s).summary).toBe('Busy');
    s.policies.title = { mode: 'replace', text: '' };
    expect(mapEvent(source(), s).summary).toBe('Busy');
  });

  test('line breaks in title and location collapse', () => {
    const s = testSettings();
    s.policies.location = { mode: 'copy' };
    const m = mapEvent(source({ subject: 'a\nb', location: 'x\r\ny' }), s);
    expect(m.summary).toBe('a b');
    expect(m.location).toBe('x y');
  });

  test('copied attendees are written with client-side scheduling only', async () => {
    const s = testSettings();
    s.policies.attendees = 'copy';
    const { desired } = await buildDesired([source()], s, COLLECTION, NOW);
    const lines = unfoldLines(desired[0]?.ics ?? '');
    const people = lines.filter((l) => l.startsWith('ATTENDEE') || l.startsWith('ORGANIZER'));
    expect(people).toHaveLength(2);
    for (const l of people) expect(l).toContain('SCHEDULE-AGENT=CLIENT');
    expect(people.join('\n')).toContain('PARTSTAT=ACCEPTED');
  });

  test('dropped attendees leave no trace in the ICS', async () => {
    const { desired } = await buildDesired([source()], testSettings(), COLLECTION, NOW);
    const ics = desired[0]?.ics ?? '';
    expect(ics).not.toContain('ATTENDEE');
    expect(ics).not.toContain('ORGANIZER');
    expect(ics).not.toContain('sam@example.test');
    expect(ics).not.toContain('Agenda');
    expect(ics).not.toContain('Room 4');
  });
});

describe('availability', () => {
  test('is always kept', async () => {
    const check = async (showAs: SourceEvent['showAs']) => {
      const { desired } = await buildDesired([source({ showAs })], testSettings(), COLLECTION, NOW);
      return unfoldLines(desired[0]?.ics ?? '');
    };
    const free = await check('Free');
    expect(free).toContain('TRANSP:TRANSPARENT');
    expect(free).toContain('X-MICROSOFT-CDO-BUSYSTATUS:FREE');
    const tentative = await check('Tentative');
    expect(tentative).toContain('STATUS:TENTATIVE');
    expect(tentative).toContain('TRANSP:OPAQUE');
    expect(await check('Oof')).toContain('X-MICROSOFT-CDO-BUSYSTATUS:OOF');
    expect(await check('Busy')).toContain('TRANSP:OPAQUE');
  });
});

describe('title filters', () => {
  const matches = (pattern: string, title: string) =>
    compileTitleFilter(pattern)?.(title.trim()) ?? false;

  test('exact by default, ignoring case and surrounding space', () => {
    expect(matches('Sport', 'Sport')).toBe(true);
    expect(matches('sport', 'SPORT')).toBe(true);
    expect(matches(' Sport ', '  Sport\t')).toBe(true);
    expect(matches('Sport', 'Sport (Büro)')).toBe(false);
    expect(matches('Sport', 'Teamsport')).toBe(false);
  });

  test('star is a wildcard', () => {
    expect(matches('Sport*', 'Sport (Büro)')).toBe(true);
    expect(matches('*essen', 'Mittagessen')).toBe(true);
    expect(matches('a*c', 'abbbc')).toBe(true);
    expect(matches('a*c', 'abbbd')).toBe(false);
    expect(matches('*', 'anything at all')).toBe(true);
  });

  test('everything else is literal', () => {
    expect(matches('a.c', 'abc')).toBe(false);
    expect(matches('a.c', 'a.c')).toBe(true);
    expect(matches('Sport (Büro)', 'Sport (Büro)')).toBe(true);
    expect(matches('[x]+', '[x]+')).toBe(true);
    expect(matches('Überprüfung', 'ÜBERPRÜFUNG')).toBe(true);
  });

  test('blank patterns match nothing', () => {
    expect(compileTitleFilter('')).toBeNull();
    expect(compileTitleFilter('   ')).toBeNull();
  });

  test('filters use the original title, not the replacement', () => {
    const s = testSettings({ titleFilters: ['Planning'] });
    s.policies.title = { mode: 'replace', text: 'Meeting' };
    expect(exclusionOf(source(), s, [compileTitleFilter('Planning') as TitleMatcher])).toBe(
      'filtered',
    );
  });

  test('availability exclusion', () => {
    const s = testSettings({ excludeShowAs: ['Free'] });
    expect(exclusionOf(source({ showAs: 'Free' }), s, [])).toBe('filtered');
    expect(exclusionOf(source({ showAs: 'Busy' }), s, [])).toBeUndefined();
  });

  test('excluded events are not rendered and are counted', async () => {
    const s = testSettings({ titleFilters: ['Sport', 'Mittag*'] });
    const events = [
      source({ id: 'a', subject: 'Sport' }),
      source({ id: 'b', subject: 'mittagessen im Büro' }),
      source({ id: 'c', subject: 'Planning' }),
    ];
    const { desired, filtered } = await buildDesired(events, s, COLLECTION, NOW);
    expect(filtered).toBe(2);
    expect(desired.map((d) => d.summary)).toEqual(['Planning']);
  });
});

describe('identity', () => {
  test('is stable, owner specific and safe for any Outlook id', async () => {
    const id = outlookId(7);
    const a = await mirrorUid('outlook', id);
    expect(a).toBe(await mirrorUid('outlook', id));
    expect(a).toMatch(/^mirrorcal-[0-9a-f]{32}$/);
    expect(a).not.toBe(await mirrorUid('other', id));
    expect(a).not.toBe(await mirrorUid('outlook', outlookId(8)));
  });

  test('href is the uid under the collection', async () => {
    const { desired } = await buildDesired([source()], testSettings(), COLLECTION, NOW);
    expect(desired[0]?.href).toBe(`${COLLECTION}${desired[0]?.uid}.ics`);
  });

  test('instances sharing an id are mirrored once', async () => {
    const { desired, duplicates } = await buildDesired(
      [source(), source({ subject: 'Again' })],
      testSettings(),
      COLLECTION,
      NOW,
    );
    expect(desired).toHaveLength(1);
    expect(duplicates).toBe(1);
  });
});

describe('title filter cost', () => {
  test('a many-star pattern against a long title answers in milliseconds', () => {
    const matcher = compileTitleFilter('*a*a*a*a*a*a*a*a*b') as TitleMatcher;
    const started = performance.now();
    expect(matcher('a'.repeat(60))).toBe(false);
    expect(matcher(`${'a'.repeat(60)}b`)).toBe(true);
    expect(performance.now() - started).toBeLessThan(100);
  });

  test('star runs and edges behave', () => {
    const matches = (p: string, t: string) => compileTitleFilter(p)?.(t) ?? false;
    expect(matches('**a**', 'xax')).toBe(true);
    expect(matches('a*', 'a')).toBe(true);
    expect(matches('*a', 'ba')).toBe(true);
    expect(matches('*ab*cd*', 'xxabyycdzz')).toBe(true);
    expect(matches('*ab*cd*', 'xxcdyyabzz')).toBe(false);
    expect(matches('a*b*c', 'abc')).toBe(true);
    expect(matches('a*b*c', 'acb')).toBe(false);
  });
});

describe('private events', () => {
  const secret = (sensitivity: SourceEvent['sensitivity']) =>
    source({ subject: 'Therapy', sensitivity });

  test('are flagged, retitled and stripped of other fields by default', () => {
    const s = testSettings();
    s.policies.description = { mode: 'copy' };
    s.policies.location = { mode: 'copy' };
    s.policies.attendees = 'copy';
    const mapped = mapEvent(secret('Private'), s);
    expect(mapped).toMatchObject({ summary: 'Private', classification: 'PRIVATE' });
    expect(mapped.description).toBeUndefined();
    expect(mapped.location).toBeUndefined();
    expect(mapped.attendees).toEqual([]);
    expect(mapped.organizer).toBeUndefined();
  });

  test('personal is flagged private and confidential confidential', () => {
    const s = testSettings();
    expect(mapEvent(secret('Personal'), s).classification).toBe('PRIVATE');
    expect(mapEvent(secret('Confidential'), s).classification).toBe('CONFIDENTIAL');
    expect(mapEvent(secret('Normal'), s).classification).toBeUndefined();
  });

  test('copy keeps every field and the flag', () => {
    const s = testSettings();
    s.policies.privateEvents = { mode: 'copy' };
    s.policies.location = { mode: 'copy' };
    expect(mapEvent(secret('Private'), s)).toMatchObject({
      summary: 'Therapy',
      location: 'Room 4',
      classification: 'PRIVATE',
    });
  });

  test('skip excludes them and only them', () => {
    const s = testSettings();
    s.policies.privateEvents = { mode: 'skip' };
    expect(exclusionOf(secret('Private'), s, [])).toBe('private');
    expect(exclusionOf(secret('Normal'), s, [])).toBeUndefined();
  });
});

describe('case folding', () => {
  const matches = (p: string, t: string) => compileTitleFilter(p)?.(t) ?? false;

  test('Greek final sigma folds like any other letter', () => {
    expect(matches('ΟΔΟΣ', 'οδος')).toBe(true);
    expect(matches('οδοσ', 'ΟΔΟΣ')).toBe(true);
    expect(matches('ς', 'Σ')).toBe(true);
  });
});
