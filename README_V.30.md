# Office Task Report V.30

## Why V.30
The Apps Script `/exec` URL is healthy in the browser, but direct browser POST login remains stuck because Google Apps Script Web Apps can redirect the POST response to `googleusercontent.com`, where browser CORS handling can prevent the frontend from receiving the response.

V.30 fixes this by routing the frontend request through a same-origin Vercel serverless function. The Vercel function calls Apps Script server-to-server and returns the JSON response to the dashboard.

## Changed files
1. `frontend/src/main.jsx` — replace the existing file.
2. `api/office-task.js` — add this new file at the project root `api/office-task.js`.

## Apps Script
No Code.gs change is required for this V.30 fix. Keep the deployed V.29 Apps Script URL active.

## Vercel
Redeploy after adding both files. No Vercel environment variable is required because the Apps Script URL is already configured in the proxy.

Optional environment variable: `OFFICE_TASK_APPS_SCRIPT_URL` can override the hardcoded endpoint.

## Login
Username: `admin`
Password: `Admin@123`
