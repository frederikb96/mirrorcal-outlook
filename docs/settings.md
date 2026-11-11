# Settings

The extension's options page has a status panel and the settings form. Defaults are in `src/core/settings.ts`.

## Settings

- **Syncing enabled.** The master switch. While it is off the extension reads nothing and writes nothing.
- **Calendar URL, User name and App password.** The URL is the collection of the destination calendar, for example `https://host/remote.php/dav/calendars/<user>/<calendar>/`. Saving asks the browser for permission to reach that server, asks the server what the URL is (it must be a calendar that holds events; the page shows its name) and then starts a sync, so the result is visible at once. An empty password field keeps the stored password, but only for the server it was entered for: pointing the URL at another server requires entering the password again, and the access to the old server is released.
- **Calendar URL for excluded events.** Optional. The events that the title filters and availability exclusions leave out are mirrored into this second calendar instead, with the same field policies and owner id. It has to be another calendar on the same server, reached with the same user and password; saving checks it like the first. Empty means excluded events are not mirrored at all.
- **Owner id.** A constant label written into every mirrored event. Only events that carry it are ever modified or deleted by the extension. Keep the id identical, together with filters and policies, on every browser that mirrors the same calendar. Changing the id leaves the events of the old id behind and holds the next sync back until it is confirmed; confirming once lets later syncs create events next to them, and **Remove another owner's events** deletes them.
- **Days back, Days ahead and Sync every (minutes).** The time window relative to now, and how often the scheduled sync runs.
- **Outlook host.** Which of the Outlook web hosts the calendar is read from. Pick the one the Outlook tab uses.
- **What is mirrored.** Title, description and location can each be copied, dropped or replaced with fixed text. The replacement only stands in where the event has that field; an event without a description or location stays without one. Attendees can be copied or dropped. Copying descriptions reads every event body on each sync.
- **Private events.** Events that Outlook marks private, personal or confidential are written with a matching `CLASS`, so calendar apps and sharing treat them as private. They can get a replacement title, be copied as they are, or not be mirrored at all (in neither calendar).
- **Excluded events.** Events whose title matches a line of the filter are left out of the calendar: mirrored into the calendar for excluded events if one is set, otherwise not at all. Matching ignores case and surrounding whitespace, is exact by default, and `*` is a wildcard. Events with a checked **Availability** (free, tentative, busy, out of office, working elsewhere, unknown) are left out the same way. Adding a filter moves the matching events out on the next sync, removing it brings them back.

## Status and buttons

The status panel shows the syncing state, the mirrored account, the Outlook session and its expiry, the last run with its counts and, when something failed, the reason and how many writes failed, plus a short history under **Recent runs**.

- **Sync now** runs a sync immediately.
- **Preview changes** computes the plan without writing anything.
- **Reset mirror** deletes all events that carry your owner id, in both calendars, after showing how many that are. While syncing is on, the next run recreates them (and asks for confirmation if other objects remain); switch syncing off to stop mirroring.
- **Remove another owner's events** lists the other owner ids found in the calendars and deletes the events of the one you pick, after showing how many that are. Objects without any mirror owner id are never deleted.
- **Review changes** and **Confirm and apply** appear when a plan is held back (see [safety](safety.md)). The first opens the plan, the second applies that run once.
- **Forget account** releases the account binding and drops the stored Outlook session; the next session the Outlook page shows is mirrored and bound after its first successful sync. A reset releases it as well.

## Toolbar badge

The toolbar icon shows a badge when something needs you: a held-back plan, a rejected password, a missing permission, an invalid setting, or a problem that keeps repeating (no Outlook session, failing reads or writes). The tooltip says which.
