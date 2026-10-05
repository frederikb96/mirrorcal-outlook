# MirrorCal Outlook

[![CI](https://github.com/frederikb96/mirrorcal-outlook/actions/workflows/ci.yaml/badge.svg)](https://github.com/frederikb96/mirrorcal-outlook/actions/workflows/ci.yaml)
[![Release](https://img.shields.io/github/v/release/frederikb96/mirrorcal-outlook)](https://github.com/frederikb96/mirrorcal-outlook/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A Manifest V3 browser extension for Chromium and Edge. It reads your calendar from Outlook on the web and mirrors it, one way, into a CalDAV calendar such as one on Nextcloud. That puts work meetings into the calendar app you already use, without any Outlook account setup on the device.

## How it works

- While you are signed in to Outlook on the web, the extension notes the access token the Outlook page itself uses and keeps it in memory only. Only requests sent by the Outlook page count, and once a sync has succeeded the mirror stays bound to that account: tokens of any other account or tenant are ignored.
- On a schedule, at browser start and on demand, it reads the calendar view for a time window from Outlook.
- It reads every object in the destination calendar, compares both sides and sends only the creations, updates and deletions that are needed. A second run over unchanged data writes nothing and, while the calendar's change tag is unchanged, only asks for that tag instead of reading the calendar.
- Outlook is only ever read. The extension has no code path that sends anything else to Outlook.

The destination calendar belongs to the mirror: edits made there are overwritten by the next sync. Use a calendar that nothing else writes to.

## Install

Build and load it as an unpacked extension.

- Build with Bun: `bun install && bun run build`. The extension ends up in `dist/`.
- Or download `mirrorcal-outlook.zip` from the [releases](https://github.com/frederikb96/mirrorcal-outlook/releases) page and unzip it.
- Open `chrome://extensions` (Edge: `edge://extensions`) and switch on developer mode.
- Choose **Load unpacked** and select the folder that contains `manifest.json`.
- Open Outlook on the web and sign in. The extension can only sync while it has seen a valid token from that page.

The manifest carries a fixed public key, so the extension id is the same whichever folder it is loaded from.

### Updating

Replace the files in the folder the extension was loaded from with the new build and press **Reload** on its card in the extensions page. Settings, the saved password and the account binding stay. Loading the new files from a different folder adds a second extension next to the old one; remove the old one first, as two copies with the same owner id would both mirror into the same calendar.

## Settings

Open the extension's options page.

- **Enabled.** The master switch. While it is off the extension reads nothing and writes nothing.
- **CalDAV calendar URL, user and app password.** The URL is the collection of the destination calendar, for example `https://host/remote.php/dav/calendars/<user>/<calendar>/`. Saving asks the browser for permission to reach that server, asks the server what the URL is (it must be a calendar that holds events; the page shows its name) and then starts a sync, so the result is visible at once. An empty password field keeps the stored password, but only for the server it was entered for: pointing the URL at another server requires entering the password again, and the access to the old server is released.
- **Time window.** Days back and days ahead, relative to now.
- **Sync interval.** How often the scheduled sync runs.
- **Outlook host.** Which of the Outlook web hosts the calendar is read from (`outlook.cloud.microsoft`, `outlook.office.com`, `outlook.office365.com`). Pick the one the Outlook tab uses.
- **Field policies.** Title, description and location can each be copied, dropped or replaced with fixed text. Attendees can be copied or dropped. By default the title is copied and everything else is dropped. Copying descriptions reads every event body on each sync.
- **Private events.** Events that Outlook marks private, personal or confidential are written with a matching `CLASS`, so calendar apps and sharing treat them as private. By default they get the title "Private" and no description, location or attendees. They can instead be copied as they are, replaced with another text, or not mirrored at all.
- **Title filters.** Events whose title matches an entry are not mirrored. Matching ignores case and surrounding whitespace, is exact by default, and `*` is a wildcard. Adding a filter removes the matching copies on the next sync, removing it brings them back.
- **Availability.** Events with a checked availability (free, tentative, busy, out of office, working elsewhere, unknown) are not mirrored.
- **Owner id.** A constant label written into every mirrored event. Only events that carry it are ever modified or deleted by the extension. Events without it, and events of another owner id, are never touched; the status counts them. Keep the id identical, together with filters and policies, on every browser that mirrors the same calendar. Changing the id leaves the events of the old id behind and holds the next sync back until it is confirmed.

The page also shows the current status: syncing state, the mirrored account, Outlook session and its expiry, the last run with its counts and, when something failed, the reason and how many writes failed, plus a short history. It has these buttons:

- **Sync now** runs a sync immediately.
- **Preview changes** computes the plan without writing anything.
- **Reset mirror** deletes all events that carry your owner id, after showing how many that are. While syncing is on, the next run recreates them; switch syncing off to stop mirroring.
- **Forget account** releases the account binding; the next successful sync binds the account of the open Outlook session. A reset releases it as well.

The toolbar icon shows a badge when something needs you: a held-back plan, a rejected password, a missing permission, an invalid setting, or a problem that keeps repeating (no Outlook session, failing reads or writes). The tooltip says which.

## Safety

- Outlook is accessed read-only. The only request the extension can build towards Outlook is the calendar view read, and a test checks that. An Outlook host is refused as the calendar server.
- A sync stops before any write if the Outlook read fails, returns something unexpected or is incomplete. Nothing is deleted on a partial read.
- A plan that would remove everything, or a large share of the mirrored events, is refused until you confirm that run on the options page. Changing a filter can legitimately cause this, so the page shows the preview and asks once. Creating events is held back the same way when the calendar already holds events of another owner id, or only events that are not this mirror's (usually the wrong calendar).
- Without a valid token the status asks you to open Outlook, and nothing is deleted.
- After the calendar server rejects the password, no further requests are made until the settings are saved again. A failure that repeats unchanged is retried at most hourly.
- The CalDAV app password is stored unencrypted in the browser's extension storage on your machine. A Nextcloud app password can be revoked but gives access to the whole account, not just one calendar. It is only sent to the server it was entered for. The Outlook token is held in session storage only and is gone when the browser restarts.

## Limitations

- Sync only runs while the browser is running and Outlook on the web has been open and signed in long enough for the extension to see a token.
- The extension depends on Outlook web's internal calendar endpoint and on the token that the page sends. Microsoft can change either without notice.
- One direction only. Changes made in the destination calendar are never read back.
- The mirror follows one Outlook account. A second account open in the same browser is ignored once the first has synced.

## Development

```
bun install
bun run typecheck
bun run lint
bun run test
bun run build
bun run package
bun run test:e2e
```

`bun run test:e2e` loads the built extension in Chromium against local fakes of Outlook and the calendar server (no real network; it needs the Chromium that Playwright ships). `bun run package` writes `mirrorcal-outlook.zip` in the repository root, with the manifest at the top level of the archive. Pushing a tag that starts with `v` runs the same checks and attaches the zip to a GitHub release.
