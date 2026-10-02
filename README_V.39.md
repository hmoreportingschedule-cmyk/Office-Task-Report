# Office Task Report V.39

## Updates
- Employee Profile module with photo, office city/address, office time and weekoff.
- Employee can change own password from Profile.
- Employee ID table has Edit/Delete actions only; no Settings button.
- Manual attendance date is explicitly shown; employee can add only Today or Previous 1 Day, never future dates.
- Weekoff, Leave and approved Weekoff Adjustment remain automatic; no manual attendance selection is required for those dates.
- Attendance IN/OUT generates fast notifications for active HOD/Admin/Master Admin users.
- Approval generates notifications for the requesting employee and active approvers, including who approved/rejected.
- Time Adjustment request immediately shows IN/OUT + time; approved adjustment updates the employee roster time.
- Day Adjustment renamed to Weekoff Adjustment.
- Meeting request provides Online / Physically options.
- Meeting Journey provides From Time / To Time and validates duration.
- Notifications poll every 10 seconds for fast dashboard updates.

## Deployment
1. Deploy `Code.gs` to the existing Apps Script Web App as a new version.
2. Deploy the complete project to the existing Cloudflare Worker.
3. Keep the existing master sheet: `My Drive / Dashboard Working / office-task-report`.
4. Do not rebuild the Users sheet or overwrite existing passwords.

## Health
After Worker deployment:
`/api/office-task/health`
should report `worker: V.39`.
