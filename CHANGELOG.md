# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/) and the project uses
[Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- One-way mirror of an Outlook web calendar into a CalDAV calendar.
- The mirror binds to the Outlook account of the first successful sync and ignores tokens of any other account; the status names it. Only requests from the Outlook page supply tokens.
- Saving the settings checks that the URL is a calendar that holds events, shows its name and starts a sync.
- Toolbar badge for a held-back plan, a rejected password, a missing permission, invalid settings and failures that keep repeating.
- Private, personal and confidential Outlook events are written with `CLASS`. A new setting decides their title: copy, replace (default "Private") or do not mirror.
- A fixed public key in the manifest keeps the extension id independent of the install folder. The README explains how to update.
- Recent runs are listed on the options page.

### Changed

- Creating events is held back for confirmation when the destination holds events of another owner id or only foreign events.
- After the calendar server rejects the password, scheduled runs stop until the settings are saved again.
- Per-object write failures show their reason and count, the run is labelled "Synced with errors", and an identical failure or held-back plan is retried hourly after three repeats.
- An empty password field keeps the stored password only for the same server origin; the old server's access is released when the URL moves.
- Title filters match in linear time.
- An Outlook host is refused as calendar URL; empty number fields are invalid instead of 0; a refused permission request is reported.
- Token capture is limited to fetch/XHR requests and does not write storage for a token that is not clearly newer; a rejected token is not taken up again.
- The options page reads storage and listens for changes instead of polling the worker.
- The worker is kept alive with a periodic extension API call while a run is active.
- A confirmation that meets a running sync stays available.
