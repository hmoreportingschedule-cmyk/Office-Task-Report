# Office Task Report V.69

## Updates
- Check Task: employees no longer use Complete directly. After Start, employee uses **Save Minutes** to record time spent on the task. Minutes accumulate across the task period. A task assigned for a date-to-date range is not auto-completed before its end date; on/after the end date, a saved time entry can close it and set progress to 100%.
- Task actual minutes remain cumulative and are stored server-side.
- Removed the **CSV** action from My Requests.
- My Profile: employee can upload/change their own profile photo and update only Mobile Number and Email Id. Other profile fields remain read-only to the employee.
- Profile photo continues to use `My Drive/Dashboard Working/office-task-report/Employees Photo` and the backend stores the Drive file reference/path.
- Existing admin/HOD/master-admin employee profile/photo management remains available.
- Existing Users sheet schema is preserved; email is stored in EmployeeSettings rather than changing the finalized Users column layout.

## Changed files
- `Code.gs`
- `src/main.jsx`

## Deployment
- Replace the root `Code.gs` in Google Apps Script and deploy a new version.
- Replace `src/main.jsx` in the GitHub project.
- Do not use a duplicate `Apps Script/Code.gs`.
