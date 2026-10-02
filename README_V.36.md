# Office Task Report — FULL V.36

This package is the complete project, including the frontend and the Cloudflare Worker proxy.

## Deployment

1. Extract this entire project.
2. Run:
   `npm install`
3. Build:
   `npm run build`
4. Authenticate Cloudflare:
   `npx wrangler login`
5. Deploy:
   `npx wrangler deploy`

The Worker uses the existing Apps Script Web App endpoint:

https://script.google.com/macros/s/AKfycbyfL_lBPCkUlHzvv0iWvbUyNmhjSE7dVh6yqHp0L4E9JAs8S6e9jDx3v2dKku3ClK6M/exec

The login path is:

Browser -> Cloudflare Worker `/api/office-task` -> Apps Script -> Users sheet.

The Worker manually follows Apps Script redirects while preserving the POST body and returns JSON to the frontend.

## Important

Do not rebuild or manually modify the Users sheet for this deployment. Keep the existing master Google Sheet:

My Drive / Dashboard Working / office-task-report

and the existing Users A:S structure.

Test Worker health after deployment:

`https://YOUR-WORKER-DOMAIN/api/office-task/health`

Expected JSON contains `"ok":true` and `"worker":"V.36"`.

Then test login with the credentials already configured in the Users sheet.
