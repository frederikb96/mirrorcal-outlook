#!/usr/bin/env bun
/**
 * Release helpers over CHANGELOG.md and package.json.
 *
 *   bun run scripts/release-notes.ts check [tag]   package.json has a dated, non-empty changelog
 *                                                   section; with a tag, the tag names that version
 *   bun run scripts/release-notes.ts notes <tag>   prints that version's section, the release body
 */
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dir, '..');
const HEADING = /^## \[(?<version>[^\]]+)\](?: - (?<date>\d{4}-\d{2}-\d{2}))?\s*$/;

export interface Section {
  version: string;
  date: string | undefined;
  body: string;
}

export function sections(changelog: string): Section[] {
  const found: Section[] = [];
  let current: { version: string; date: string | undefined; lines: string[] } | undefined;
  const close = () => {
    if (current)
      found.push({
        version: current.version,
        date: current.date,
        body: current.lines.join('\n').trim(),
      });
  };
  for (const line of changelog.split('\n')) {
    const match = HEADING.exec(line);
    if (match?.groups) {
      close();
      current = { version: match.groups.version as string, date: match.groups.date, lines: [] };
    } else if (line.startsWith('## ')) {
      close();
      current = undefined;
    } else current?.lines.push(line);
  }
  close();
  return found;
}

/** Problems that would make a release with this version wrong; empty when the guard passes. */
export function problems(changelog: string, packageVersion: string, tag?: string): string[] {
  const found: string[] = [];
  if (tag !== undefined && tag.replace(/^v/, '') !== packageVersion)
    found.push(`tag ${tag} does not match package.json version ${packageVersion}`);
  const section = sections(changelog).find((s) => s.version === packageVersion);
  if (!section) found.push(`CHANGELOG.md has no section for ${packageVersion}`);
  else {
    if (!section.date) found.push(`the CHANGELOG.md section for ${packageVersion} has no date`);
    if (!section.body) found.push(`the CHANGELOG.md section for ${packageVersion} is empty`);
  }
  return found;
}

function main(argv: string[]): number {
  const [command, arg] = argv;
  const changelog = readFileSync(resolve(root, 'CHANGELOG.md'), 'utf8');
  const { version } = JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8'));
  if (command === 'check') {
    const list = problems(changelog, version, arg);
    for (const p of list) console.error(`error: ${p}`);
    if (list.length === 0) console.log(`release metadata is consistent for ${version}`);
    return list.length === 0 ? 0 : 1;
  }
  if (command === 'notes' && arg) {
    const wanted = arg.replace(/^v/, '');
    const section = sections(changelog).find((s) => s.version === wanted);
    if (!section?.body) {
      console.error(`error: CHANGELOG.md has no section for ${wanted}`);
      return 1;
    }
    console.log(section.body);
    return 0;
  }
  console.error('usage: release-notes.ts check [tag] | notes <tag>');
  return 2;
}

if (import.meta.main) process.exit(main(process.argv.slice(2)));
