# MirrorCal Outlook

[![CI](https://github.com/frederikb96/mirrorcal-outlook/actions/workflows/ci.yaml/badge.svg)](https://github.com/frederikb96/mirrorcal-outlook/actions/workflows/ci.yaml)
[![Release](https://img.shields.io/github/v/release/frederikb96/mirrorcal-outlook)](https://github.com/frederikb96/mirrorcal-outlook/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A Manifest V3 browser extension for Chromium and Edge. It reads your calendar from Outlook on the web and mirrors it, one way, into a CalDAV calendar such as one on Nextcloud. That puts work meetings into the calendar app you already use, without any Outlook account setup on the device.

## How it works

- While you are signed in to Outlook on the web, the extension notes the access token the Outlook page itself uses and keeps it in memory only.
- On a schedule, at browser start and on demand, it reads the calendar view for a time window from Outlook.
- It reads every object in the destination calendar, compares both sides and sends only the creations, updates and deletions that are needed. A second run over unchanged data writes nothing.
- Outlook is only ever read. The extension has no code path that sends anything else to Outlook.

The destination calendar belongs to the mirror: edits made there are overwritten by the next sync. Use a calendar that nothing else writes to.

## Install

Build and load it as an unpacked extension.

- Build with Bun: `bun install && bun run build`. The extension ends up in `dist/`.
- Or download `mirrorcal-outlook.zip` from the [releases](https://github.com/frederikb96/mirrorcal-outlook/releases) page and unzip it.
- Open `chrome://extensions` (Edge: `edge://extensions`) and switch on developer mode.
- Choose **Load unpacked** and select the folder that contains `manifest.json`.
- Open Outlook on the web and sign in. The extension can only sync while it has seen a valid token from that page.

## Settings

Open the extension's options page.

- **CalDAV calendar URL, user and app password.** The URL is the collection of the destination calendar, for example `https://host/remote.php/dav/calendars/<user>/<calendar>/`. Saving asks the browser for permission to reach that server.
- **Time window.** Days back and days ahead, relative to now.
- **Sync interval.** How often the scheduled sync runs.
- **Field policies.** Title, description and location can each be copied, dropped or replaced with fixed text. Attendees can be copied or dropped. By default the title is copied and everything else is dropped.
- **Title filters.** Events whose title matches an entry are not mirrored. Matching ignores case and surrounding whitespace, is exact by default, and `*` is a wildcard. Adding a filter removes the matching copies on the next sync, removing it brings them back.
- **Owner id.** A constant label written into every mirrored event. Only events that carry it are ever modified or deleted by the extension, so other events in the same calendar stay untouched and are listed in the status.

The page also shows the current status (token present and its expiry, last run, counts, last error) and has these buttons:

- **Sync now** runs a sync immediately.
- **Preview changes** computes the plan without writing anything.
- **Reset mirror** deletes all events that carry your owner id, after showing how many that are.

## Safety

- Outlook is accessed read-only. The only request the extension can build towards Outlook is the calendar view read, and a test checks that.
- A sync stops before any write if the Outlook read fails, returns something unexpected or is incomplete. Nothing is deleted on a partial read.
- A plan that would remove everything, or a large share of the mirrored events, is refused until you confirm that run on the options page. Changing a filter can legitimately cause this, so the page shows the preview and asks once.
- Without a valid token the status asks you to open Outlook, and nothing is deleted.
- The CalDAV app password is stored unencrypted in the browser's extension storage on your machine. Use an app password that is limited to what the mirror needs. The Outlook token is held in session storage only and is gone when the browser restarts.

## Limitations

- Sync only runs while the browser is running and Outlook on the web has been open and signed in long enough for the extension to see a token.
- The extension depends on Outlook web's internal calendar endpoint and on the token that the page sends. Microsoft can change either without notice.
- One direction only. Changes made in the destination calendar are never read back.

## Development

```
bun install
bun run typecheck
bun run lint
bun run test
bun run build
bun run package
```

`bun run package` writes `mirrorcal-outlook.zip` in the repository root, with the manifest at the top level of the archive. Pushing a tag that starts with `v` runs the same checks and attaches the zip to a GitHub release.
