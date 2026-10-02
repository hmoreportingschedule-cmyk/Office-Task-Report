# Office Task Report V.29

## Changed files
1. `frontend/src/main.jsx`
2. `Code.gs`

## Replacement
- Replace `frontend/src/main.jsx` in the GitHub/Vercel project.
- Replace the Apps Script project's existing `Code.gs` with this `Code.gs`.
- Deploy Apps Script as a **new Web App version** (Execute as: Me; Who has access: Anyone).
- Use the new `/exec` URL in `frontend/src/main.jsx` if Google gives a different URL.

## What V.29 fixes
- Uses the latest Apps Script URL supplied by the user.
- Login request timeout increased to 60 seconds to accommodate first-time Apps Script/Drive initialization.
- `doGet?action=health` now actually checks/initializes the backend and confirms the master Users sheet. Normal `doGet` remains lightweight.
- Backend version is V.29.
- Existing manually finalized Users A:S schema is untouched.
- Password remains column E (`password`).

## Test
After deploying, open:
`<new exec URL>?action=health`

It should return JSON with `ok:true` and `ready:true`. If it returns an error, that error identifies the backend issue directly.
