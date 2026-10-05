# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/) and the project uses
[Semantic Versioning](https://semver.org/).

## [Unreleased]

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
