# Office Task Report V.73

## Changes
- Manual Attendance: Attendance Date, IN Time and OUT Time shown in one 3-column row.
- Break section title changed to `Break Time (Only Namaz & Lunch)` and the `Namaz · Lunch Time` helper text removed.
- Break 1: For Namaz — Namaz Name, Start Time, End Time aligned on one row.
- Break 2: For Namaz — Namaz Name, Start Time, End Time aligned on one row.
- Break 3: For Lunch Time — Start Time and End Time aligned on one row.
- Notes save compatibility improved: frontend tries `saveNote`, then legacy aliases when an older deployed backend responds with `Unknown action`; backend also supports `saveNotes`, `noteSave`, and `save_note` aliases.

## Deployment
- Replace root `Code.gs` in the existing Google Apps Script project and deploy a new version.
- Replace `src/main.jsx` in GitHub with this `main.jsx`.
- `Apps Script/Code.gs` is not the canonical backend file.
