import { describe, expect, test } from 'bun:test';
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';

const root = resolve(import.meta.dir, '..');
const read = (path: string) => readFileSync(resolve(root, path), 'utf8');
const docs = readdirSync(resolve(root, 'docs'))
  .filter((name) => name.endsWith('.md'))
  .map((name) => `docs/${name}`);

/** Relative targets of markdown links and html src/href attributes. */
function localTargets(text: string): string[] {
  const targets = [
    ...[...text.matchAll(/\]\(([^)\s]+)\)/g)].map((m) => m[1] as string),
    ...[...text.matchAll(/(?:src|href)="([^"]+)"/g)].map((m) => m[1] as string),
  ];
  return targets.filter((t) => !/^([a-z]+:|#)/i.test(t)).map((t) => t.split('#')[0] as string);
}

describe('documentation', () => {
  test('every docs file is linked from the README', () => {
    const readme = read('README.md');
    for (const doc of docs) expect(readme).toContain(`](${doc})`);
  });

  test('every relative link and image resolves', () => {
    for (const file of ['README.md', ...docs]) {
      for (const target of localTargets(read(file))) {
        const path = resolve(root, dirname(file), target);
        expect(existsSync(path), `${file} links ${target}`).toBe(true);
      }
    }
  });

  test('the settings reference names every button and field group of the options page', () => {
    const page = read('src/extension/options.html');
    const reference = read('docs/settings.md');
    const names = [
      ...[...page.matchAll(/<button[^>]*>([^<]+)<\/button>/g)].map((m) => m[1]),
      ...[...page.matchAll(/<legend>([^<]+)<\/legend>/g)].map((m) => m[1]),
    ] as string[];
    expect(names.length).toBeGreaterThan(5);
    for (const name of names) {
      // "Calendars" and "Time span and schedule" are legends; the reference names their fields.
      if (['Calendars', 'Time span and schedule', 'Save'].includes(name)) continue;
      expect(reference, `docs/settings.md mentions ${name}`).toContain(name);
    }
  });
});
