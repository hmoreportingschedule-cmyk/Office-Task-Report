# Office Task Report V.31

## Changed files only

1. `src/main.jsx` — replace the existing file.
2. `api/office-task.js` — add this new file at the repository root under `api/`.

## Important

Your GitHub project structure is already:

- `src/main.jsx`
- `src/styles.css`
- `index.html`
- `package.json`
- `api/office-task.js`

Do NOT create a second root-level `main.jsx`.

The Vercel proxy forwards login requests to the current Apps Script Web App URL.

Apps Script URL:
`https://script.google.com/macros/s/AKfycbyfL_lBPCkUlHzvv0iWvbUyNmhjSE7dVh6yqHp0L4E9JAs8S6e9jDx3v2dKku3ClK6M/exec`

Master Google Drive path used by Apps Script:
`My Drive/Dashboard Working/office-task-report`

Expected master spreadsheet name:
`office-task-report`

Expected Users tab schema remains A:S; no Users sheet migration is performed.

Login:
- username: `admin`
- password: `Admin@123`
