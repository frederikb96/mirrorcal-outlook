# Development

The extension is TypeScript, built with Bun, with no framework. Logic lives in `src/core` and has no browser APIs; `src/extension` holds the thin adapters, the service worker and the options page.

## Commands

```
bun install
bun run typecheck
bun run lint
bun run test
bun run build
bun run package
bun run test:e2e
bun run screenshots
```

- `bun run test` runs the unit tests with injected fetch and a fake storage.
- `bun run test:e2e` loads the built extension in Chromium against local fakes of Outlook and the calendar server (no real network). It needs the Chromium that Playwright ships: `bunx playwright-core install chromium`. CI runs it on every pull request.
- `bun run package` writes `mirrorcal-outlook.zip` in the repository root, with the manifest at the top level of the archive.

## Screenshots

The pictures in `docs/images/` are made by `bun run screenshots`: the browser test's harness with invented events dated relative to now. A change to the options page re-runs it in the same pull request. `bun run scripts/make-icons.ts` regenerates the toolbar icons and the README logo.

## Releases

The version lives in `package.json`; the manifest takes it from there. A release is a pull request that bumps it and moves the `CHANGELOG.md` entries from Unreleased into a dated section. After the merge, tag the merge commit `vX.Y.Z` and push the tag. The release workflow checks the tag against `package.json` and the changelog, runs the checks, and publishes a GitHub Release whose body is that changelog section, with the zip attached. `bun run scripts/release-notes.ts check` runs the same guard without a tag, and the pull-request checks run it.
