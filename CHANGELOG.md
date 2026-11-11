# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/) and the project uses
[Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- `bun run screenshots` regenerates the options page pictures used in the README.
- The browser test runs in the pull-request checks, and so does a check that `package.json` and the changelog agree on the version.

### Changed

- The README is a front page with links into `docs/` (how it works, settings, safety, development).
- Release notes are the version's changelog section instead of a generated list of pull requests; the release workflow refuses a tag that does not match `package.json` or has no changelog section.
- Workflow actions are pinned to commit SHAs.

## [0.2.0] - 2026-11-09

### Added

- Optional calendar for excluded events: events left out by the title filters or availability exclusions are mirrored into a second calendar on the same server instead of being dropped. Both calendars are read and planned before anything is written, a held-back plan in either holds the whole run, and Reset mirror covers both.

### Changed

- The status shows how many events a calendar holds after the run's own writes, and the save message reports the changes of both calendars.

### Fixed

- The unit tests pin the clock and the browser test widens its window, so the fixture week no longer drops out of the sync window as time passes.

## [0.1.1] - 2026-10-06

### Fixed

- A replaced description or location is only written for events that have one; events without a description or location no longer receive the replacement text.

## [0.1.0] - 2026-10-05

### Added

- One-way mirror of an Outlook web calendar into a CalDAV calendar, read with the token the Outlook page itself uses. Outlook is only ever read.
- Scheduled, on-demand and preview syncs; a quiet run only asks the calendar for its change tag.
- Field policies for title, description, location and attendees; title filters with `*` wildcards; availability exclusions; configurable time window and interval.
- Private, personal and confidential Outlook events are written with `CLASS`; their title is replaced ("Private" by default), copied, or the event is not mirrored.
- Only events carrying the mirror's owner id are modified or deleted. Large deletions, events of another owner id and a first sync into a calendar that only holds other events are held back until confirmed on the options page.
- The mirror binds to the Outlook account of its first successful sync and ignores tokens of any other account (counted in the status); Forget account releases the binding.
- Saving the settings checks that the URL is a calendar that holds events, shows its name and starts a sync. The stored password is only sent to the server it was entered for.
- Toolbar badge for a held-back plan, a rejected password, a missing permission, invalid settings and failures that keep repeating. After a rejected password scheduled runs stop; an unchanged repeating failure is retried hourly.
- Reset mirror and removal of another owner id's events, each after showing the count.
- A fixed public key in the manifest keeps the extension id independent of the install folder.
- Options page with status, recent runs and plan preview.
