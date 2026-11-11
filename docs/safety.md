# Safety

What the extension guarantees, and what it stores.

## Outlook is read-only

The only request the extension can build towards Outlook is the calendar-view read, and a test checks that. An Outlook host is refused as the calendar server.

## The destination is guarded

- A sync stops before any write if the Outlook read fails, returns something unexpected or is incomplete. Nothing is deleted on a partial read.
- Every calendar is read before anything is written, and a run that fails to read either writes nothing.
- A plan that would remove everything, or a large share of the mirrored events, of either calendar is refused until you confirm that run on the options page; until then neither calendar is changed. Changing a filter can legitimately cause this, so the page shows the preview and asks once.
- Creating events is held back the same way when the calendar already holds events of another owner id, or only events that are not this mirror's (usually the wrong calendar).
- Without a valid token the status asks you to open Outlook, and nothing is deleted.
- Only objects carrying this mirror's owner id are ever updated or deleted. Foreign objects are counted, never touched.

The thresholds are constants in `src/core/constants.ts`.

## What is stored

- The CalDAV app password is stored unencrypted in the browser's extension storage on your machine. A Nextcloud app password can be revoked but gives access to the whole account, not just one calendar. It is only sent to the server it was entered for.
- The Outlook token is held in session storage only and is gone when the browser restarts.
- Settings, the account binding and the retry state live in the extension's local storage.
