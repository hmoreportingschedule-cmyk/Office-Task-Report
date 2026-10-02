# Office Task Report V.67 — Notes Save Fix

## Fix
- Fixed `Unknown action: saveNote` by ensuring the Google Apps Script route explicitly supports `saveNote`.
- Added backward-compatible aliases `createNote` and `addNote`.
- Notes sheet is automatically created if missing, with the required columns.
- Existing Notes data is preserved.

## Deploy
Replace only the root `Code.gs` in the existing Google Apps Script project, then create a new deployment version and update the Web App deployment.

Do not use `Apps Script/Code.gs`; the root `Code.gs` is the canonical backend file.
