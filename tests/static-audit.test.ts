import { describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync } from 'node:fs';
import { join, relative } from 'node:path';

const SRC = join(import.meta.dir, '..', 'src');

function sources(dir: string): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) return sources(path);
    return path.endsWith('.ts') ? [path] : [];
  });
}

const files = sources(SRC).map((path) => ({
  name: relative(SRC, path),
  text: readFileSync(path, 'utf8'),
}));

const methodsIn = (text: string) =>
  [...text.matchAll(/method:\s*'([A-Z]+)'/g)].map((match) => match[1] as string);

describe('what the code can send', () => {
  test('only the Outlook reader and the CalDAV client name an HTTP method', () => {
    const naming = files.filter((f) => methodsIn(f.text).length > 0).map((f) => f.name);
    expect(naming.sort()).toEqual(['core/caldav.ts', 'core/outlook.ts']);
  });

  test('the Outlook reader names exactly one method, GET', () => {
    const outlook = files.find((f) => f.name === 'core/outlook.ts');
    expect(methodsIn(outlook?.text ?? '')).toEqual(['GET']);
  });

  test('the CalDAV client only lists, inspects, reads the change tag, writes and deletes', () => {
    const caldav = files.find((f) => f.name === 'core/caldav.ts');
    expect([...new Set(methodsIn(caldav?.text ?? ''))].sort()).toEqual([
      'DELETE',
      'PROPFIND',
      'PUT',
      'REPORT',
    ]);
  });

  test('every request method is a literal, so none can be chosen at run time', () => {
    for (const f of files) {
      const code = f.text.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
      const all = [...code.matchAll(/\bmethod\s*:/g)].length;
      const literal = [...code.matchAll(/\bmethod:\s*'[A-Z]+'/g)].length;
      expect({ file: f.name, dynamic: all - literal }).toEqual({ file: f.name, dynamic: 0 });
    }
  });

  test('every PROPFIND asks for the collection itself, never its contents', () => {
    const caldav = files.find((f) => f.name === 'core/caldav.ts')?.text ?? '';
    const requests = [...caldav.matchAll(/method: 'PROPFIND',\s*headers: \{ Depth: '(\d)'/g)];
    expect(requests.length).toBe([...caldav.matchAll(/method: 'PROPFIND'/g)].length);
    expect(requests.every((m) => m[1] === '0')).toBe(true);
  });

  test('the extension shell reaches the calendar server only through the engine', () => {
    for (const f of files.filter((file) => file.name.startsWith('extension/'))) {
      expect({ file: f.name, direct: /core\/caldav\.ts/.test(f.text) }).toEqual({
        file: f.name,
        direct: false,
      });
    }
  });

  test('every flight against the destination runs under the worker keep-alive', () => {
    const runner = files.find((f) => f.name === 'extension/runner.ts')?.text ?? '';
    expect([...runner.matchAll(/keepAliveWhile\(/g)].length).toBe(2);
  });

  test('no other module performs requests', () => {
    const callers = files
      .filter((f) => /\bfetch(?:Impl)?\(|XMLHttpRequest|sendBeacon|WebSocket/.test(f.text))
      .map((f) => f.name);
    // runner and the CLI-style adapters only hand `fetch` to the engine.
    expect(callers.sort()).toEqual(['core/caldav.ts', 'core/outlook.ts', 'extension/runner.ts']);
  });

  test('no Outlook action other than the calendar view appears anywhere', () => {
    const forbidden =
      /service\.svc|sendmail|me\/(?:events|messages|calendars)|CreateItem|UpdateItem|DeleteItem|\/(?:accept|decline|respond)\b/i;
    for (const f of files) {
      const code = f.text.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
      expect({ file: f.name, hit: forbidden.exec(code)?.[0] }).toEqual({
        file: f.name,
        hit: undefined,
      });
    }
  });

  test('page scripting and cookie access are not used', () => {
    for (const f of files) {
      expect(f.text).not.toMatch(/chrome\.(scripting|cookies|debugger|tabs)\b/);
    }
  });
});
