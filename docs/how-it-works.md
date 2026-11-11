# How it works

One sync run, from the token to the writes.

## The Outlook session

While you are signed in to Outlook on the web, the extension notes the access token the Outlook page itself uses for its own API calls and keeps it in memory only. Only requests sent by the Outlook page count. The token is checked for its audience and expiry, and a token that outlives the current one by a margin replaces it.

Once a sync has succeeded, the mirror is bound to that account. Tokens of any other account or tenant are ignored; the status counts them, so a switched account is not silent. **Forget account** releases the binding.

## A run

- A run starts on a schedule (an alarm), at browser start, when a new token arrives, on **Sync now** and on **Preview changes**.
- It reads the calendar view for the configured time window from Outlook, page by page. A read that fails, returns something unexpected or is incomplete ends the run before any write.
- It reads every object of the destination calendar, and of the calendar for excluded events when one is set.
- It maps each Outlook event to a calendar object according to the field policies, private-event handling, title filters and availability exclusions, and compares both sides.
- It sends only the creations, updates and deletions that are needed. A second run over unchanged data writes nothing; while the calendar's change tag is unchanged it asks only for that tag instead of reading the calendar.

The service worker the extension runs in can be stopped by the browser at any time, so state lives in extension storage, only one run is in flight at a time, and every worker start re-checks the schedule.

## Ownership

Every mirrored object carries the owner id and a content hash. Only objects with this owner id are ever updated or deleted. Objects without one, and objects of another owner id, are left alone and counted in the status. The destination calendars belong to the mirror: edits made there are overwritten by the next sync, so use calendars that nothing else writes to.

## Held-back plans

A plan can be refused until you confirm that run on the options page. The page shows the preview and asks once. The cases are listed in [safety.md](safety.md).

## Failures

After the calendar server rejects the password, scheduled runs make no requests until the settings are saved again; a manual sync still tries once. A failing run that repeats unchanged is retried at most hourly by the schedule, and the toolbar icon shows a badge for it.
