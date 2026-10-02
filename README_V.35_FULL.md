# Office Task Report — Full V.35

This ZIP contains the complete project, not only changed files.

## Deployment

1. Install Node.js.
2. Run `npm install`.
3. Run `npm run build`.
4. Run `npx wrangler login`.
5. Run `npm run deploy`.

The Cloudflare Worker serves the Vite `dist` site and proxies `/api/office-task` to the Google Apps Script Web App. The Apps Script URL is already configured in `worker.js`.

## Google Apps Script

Current endpoint:
`https://script.google.com/macros/s/AKfycbyfL_lBPCkUlHzvv0iWvbUyNmhjSE7dVh6yqHp0L4E9JAs8S6e9jDx3v2dKku3ClK6M/exec`

Existing Google Sheet path:
`My Drive/Dashboard Working/office-task-report`

Do not change the Users sheet manually if it already has the finalized A:S schema.

## Important

The browser uses the same-origin `/api/office-task` route. The Worker manually follows Apps Script redirects while preserving the POST body, avoiding the direct browser-to-Apps-Script CORS/redirect issue.
