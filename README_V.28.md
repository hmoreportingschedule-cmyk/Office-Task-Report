# Office Task Report — V.28

## Changed file
- `frontend/src/main.jsx`

## What changed
1. Apps Script Web App endpoint updated to the latest URL supplied by the user.
2. Sign-in request timeout increased from 20 seconds to 30 seconds so slower Apps Script responses are not cut off too early.
3. No Google Users sheet/schema/backend data structure was changed.
4. Existing V.27 login logic remains in place, including the manual A:S Users schema and plain-text `password` column matching requested by the user.

## Replacement
Replace only:
`frontend/src/main.jsx`

No Code.gs replacement is required for this V.28 update.

## Important deployment check
The new Apps Script URL must be an active Web App deployment. Use:
- Execute as: Me
- Who has access: Anyone (for an external Vercel frontend)

After replacing `main.jsx`, rebuild/redeploy the Vercel project.
