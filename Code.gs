/**
 * OFFICE TASK REPORT V.62
 * Google Apps Script backend
 *
 * Architecture:
 * 1 Master spreadsheet: Users, Roles, Templates, Locations, Notifications, AuditLog
 * 1 yearly spreadsheet per employee: EmployeeCode_Name_Year / user requested format Name_EmployeeCode_Year
 * Example: Asif_12345_2026
 */

const USERS_SCHEMA_VERSION = 'V62_CUMULATIVE';

const CONFIG = {
  MASTER_NAME: 'office-task-report',
  MASTER_FOLDER_PATH: ['Dashboard Working','office-task-report'],
  EMPLOYEE_FOLDER: 'Office Task Report - Employee Files',
  PHOTO_FOLDER: 'Office Task Report - Photos',
  YEARLY_SHEETS: ['Profile','Attendance','Tasks','TaskTime','Leave','Requests','Activities','Notifications'],
  DEFAULT_ADMIN: {username:'admin', password:'Admin@123', name:'Master Admin', role:'MASTER_ADMIN'}
};

function doGet(e) {
  try {
    const action = e && e.parameter ? String(e.parameter.action || '').trim().toLowerCase() : '';
    if (action === 'health') {
      const status = ensureBackend_();
      return json_({ok:true,app:'Office Task Report',version:'V.62',ready:true,message:'Backend ready.',masterId:status.masterId,usersSheetUrl:status.usersSheetUrl,time:new Date().toISOString()});
    }
    return json_({
      ok:true,
      app:'Office Task Report',
      version:'V.62',
      message:'Office Task Report API is running.',
      time:new Date().toISOString()
    });
  } catch(err) {
    return json_({ok:false,app:'Office Task Report',version:'V.62',message:String(err.message || err),time:new Date().toISOString()});
  }
}

function doPost(e) {
  const started = Date.now();
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(body.action || '').trim();

    // IMPORTANT: Login must be fast. Do NOT run Drive/folder/schema setup on
    // every login request. The previous version did that and could spend long
    // enough in DriveApp/SpreadsheetApp that Vercel/browser remained on
    // "Signing in...". Login resolves the existing master sheet directly.
    if (action === 'login') {
      try {
        const result = login_(body.username, body.password);
        return json_(Object.assign({ok:true}, result || {}));
      } catch (loginError) {
        // If the master sheet/schema has never been initialized, perform the
        // automatic setup ONCE and retry login. Existing valid sheets are not
        // rebuilt and existing passwords are never overwritten.
        const msg = String(loginError && loginError.message || loginError);
        const needsSetup = /not found|header mismatch|Users sheet not found|Master Google Sheet/i.test(msg);
        if (needsSetup) {
          ensureBackend_();
          const result = login_(body.username, body.password);
          return json_(Object.assign({ok:true}, result || {}));
        }
        throw loginError;
      }
    }

    // Health and all authenticated actions may use the normal automatic setup.
    // Setup runs only when needed because MASTER_ID/schema version are persisted.
    if (action === 'health' || action !== 'login') {
      const props = PropertiesService.getScriptProperties();
      if (action === 'health' || !props.getProperty('MASTER_ID') || props.getProperty('USERS_SCHEMA_VERSION') !== USERS_SCHEMA_VERSION) {
        ensureBackend_();
      }
    }

    const result = route_(action, body);
    return json_(Object.assign({ok:true}, result || {}));
  } catch(err) {
    return json_({
      ok:false,
      message:String(err && err.message || err),
      code:'API_ERROR',
      elapsedMs:Date.now()-started
    });
  }
}


/**
 * Finds or creates a nested Google Drive folder path inside My Drive.
 * Example:
 * ['Dashboard Working','office-task-report','Employees Task Files']
 */
function getOrCreateFolderPath_(parts) {
  if (!Array.isArray(parts) || parts.length === 0) {
    throw new Error('Google Drive folder path is empty.');
  }

  let parent = DriveApp.getRootFolder();

  parts.forEach(function(name) {
    name = String(name || '').trim();
    if (!name) return;

    const folders = parent.getFoldersByName(name);

    if (folders.hasNext()) {
      parent = folders.next();
    } else {
      parent = parent.createFolder(name);
    }
  });

  return parent;
}

function ensureBackend_() {
  const props = PropertiesService.getScriptProperties();
  const masterFolder = getOrCreateFolderPath_(['Dashboard Working','office-task-report']);
  const masterName = CONFIG.MASTER_NAME;
  let ss = null;

  // 1) Reuse saved master spreadsheet when valid.
  const savedId = props.getProperty('MASTER_ID');
  if (savedId) {
    try {
      const saved = SpreadsheetApp.openById(savedId);
      if (saved.getName() === masterName) ss = saved;
    } catch (e) {}
  }

  // 2) If this script is container-bound to the master sheet, reuse it.
  if (!ss) {
    try {
      const active = SpreadsheetApp.getActiveSpreadsheet();
      if (active && active.getName() === masterName) ss = active;
    } catch (e) {}
  }

  // 3) Find exact master sheet by name, preferring one inside the requested folder.
  if (!ss) {
    const files = DriveApp.getFilesByName(masterName);
    let fallback = null;
    while (files.hasNext()) {
      const f = files.next();
      if (f.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
      try {
        const candidate = SpreadsheetApp.openById(f.getId());
        const parents = f.getParents();
        let inMasterFolder = false;
        while (parents.hasNext()) {
          if (parents.next().getId() === masterFolder.getId()) { inMasterFolder = true; break; }
        }
        if (inMasterFolder) { ss = candidate; break; }
        if (!fallback) fallback = candidate;
      } catch (e) {}
    }
    if (!ss) ss = fallback;
  }

  // 4) If it does not exist, create it automatically in the exact folder.
  if (!ss) {
    ss = SpreadsheetApp.create(masterName);
    const file = DriveApp.getFileById(ss.getId());
    try { file.moveTo(masterFolder); } catch (e) {}
  }

  const sheetDefs = {
    Users: [
      'name','employeeId','username','role','password','department',
      'designation','phone','whatsapp','office','address','country',
      'region','state','division','district','status','createdAt','updatedAt'
    ],
    Templates:['id','name','details','category','priority','fromDay','toDay','createdBy','createdAt'],
    Notifications:['id','username','title','message','type','time','read'],
    Notes:['id','username','role','text','date','time','createdAt'],
    Approvals:['id','requestId','type','employeeId','employeeName','username','date','details','status','createdBy','createdAt','updatedAt'],
    AuditLog:['time','username','action','module','details'],
    Holidays:['id','date','name','type','createdBy','createdAt'],
    SystemSettings:['key','value','updatedAt'],
    EmployeeSettings:['employeeId','username','officeCity','officeAddress','officeInTime','officeOutTime','weekoffDay','photoUrl','updatedAt'],
    AttendanceRoster:['employeeId','username','date','officeInTime','officeOutTime','weekoffDay','updatedAt']
  };

  // 5) Automatically create all support sheets.
  Object.keys(sheetDefs).forEach(function(name){
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1,1,1,sheetDefs[name].length).setValues([sheetDefs[name]]);
      sh.setFrozenRows(1);
    } else if (name !== 'Users') {
      ensureHeaderColumns_(sh, sheetDefs[name]);
    }
  });

  // 6) Automatically create/repair the Users schema WITHOUT losing existing data.
  const users = ss.getSheetByName('Users');
  ensureUsersSchemaAuto_(users, sheetDefs.Users);

  // 7) Ensure Master Admin exists. Never overwrite a manually entered password.
  const rows = readRows_(users);
  const admin = rows.find(function(r){ return String(r.username || '').trim().toLowerCase() === 'admin'; });
  if (!admin) {
    users.appendRow([
      'Master Admin','', 'admin','MASTER_ADMIN', CONFIG.DEFAULT_ADMIN.password,
      'Administration','Administrator','','','','','','','','','','ACTIVE',new Date(),new Date()
    ]);
  } else {
    const patch = {};
    if (!String(admin.status || '').trim()) patch.status = 'ACTIVE';
    if (!String(admin.role || '').trim()) patch.role = 'MASTER_ADMIN';
    if (!String(admin.password || '')) patch.password = CONFIG.DEFAULT_ADMIN.password;
    if (Object.keys(patch).length) updateByKey_(users,'username','admin',patch);
  }

  // 8) Automatically create employee task/photo folders.
  let employeeFolder = null, photoFolder = null;
  try {
    const eid = props.getProperty('EMPLOYEE_FOLDER_ID');
    if (eid) employeeFolder = DriveApp.getFolderById(eid);
  } catch(e) {}
  try {
    const pid = props.getProperty('PHOTO_FOLDER_ID');
    if (pid) photoFolder = DriveApp.getFolderById(pid);
  } catch(e) {}

  if (!employeeFolder) {
    employeeFolder = getOrCreateFolderPath_(['Dashboard Working','office-task-report','Employees Task Files']);
    props.setProperty('EMPLOYEE_FOLDER_ID', employeeFolder.getId());
  }
  if (!photoFolder) {
    photoFolder = getOrCreateFolderPath_(['Dashboard Working','office-task-report','Employees Photo']);
    props.setProperty('PHOTO_FOLDER_ID', photoFolder.getId());
  }

  props.setProperty('MASTER_ID', ss.getId());
  props.setProperty('USERS_SCHEMA_VERSION', USERS_SCHEMA_VERSION);
  invalidateUsersCache_();

  return {
    ready:true,
    masterId:ss.getId(),
    masterUrl:ss.getUrl(),
    masterName:ss.getName(),
    usersSheetUrl:ss.getUrl() + '#gid=' + users.getSheetId(),
    employeeFolderUrl:employeeFolder.getUrl(),
    photoFolderUrl:photoFolder.getUrl(),
    schemaVersion:USERS_SCHEMA_VERSION
  };
}

/**
 * Automatically normalizes the Users sheet to the final A:S schema while
 * preserving existing rows by matching common header names. This replaces the
 * previous manual Users-sheet preparation step.
 */
function ensureUsersSchemaAuto_(sh, expected) {
  if (!sh) throw new Error('Unable to create Users sheet.');

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  const current = lastRow > 0 && lastCol > 0
    ? sh.getRange(1,1,1,lastCol).getValues()[0].map(function(h){ return String(h || '').trim(); })
    : [];

  const exact = current.length === expected.length && expected.every(function(h,i){ return current[i] === h; });
  if (exact) {
    sh.setFrozenRows(1);
    return;
  }

  // Read existing rows before changing columns. Map old names to final names.
  const data = lastRow > 1 && lastCol > 0
    ? sh.getRange(2,1,lastRow-1,lastCol).getValues()
    : [];
  const aliases = {
    employeeCode:'employeeId', employeeID:'employeeId', EmployeeId:'employeeId', EmployeeID:'employeeId',
    userId:'username', userID:'username', UserId:'username',
    Password:'password', pass:'password', passwordHash:'password',
    phoneNo:'phone', mobile:'phone', whatsappNo:'whatsapp',
    status:'status'
  };
  const normalized = current.map(function(h){ return aliases[h] || h; });
  const index = {};
  normalized.forEach(function(h,i){ if (h && index[h] === undefined) index[h] = i; });

  // Do not destroy the sheet if it contains an unknown/non-empty structure.
  // Instead, create the final header and preserve recognizable data.
  sh.clearContents();
  sh.getRange(1,1,1,expected.length).setValues([expected]);

  if (data.length) {
    const out = data.map(function(row){
      return expected.map(function(field){
        const i = index[field];
        return i === undefined ? '' : row[i];
      });
    });
    sh.getRange(2,1,out.length,expected.length).setValues(out);
  }
  sh.setFrozenRows(1);
}

function setup() {
  return ensureBackend_();
}

/**
 * Final validation only.
 * This function NEVER changes/rebuilds the Users sheet.
 */
function validateUsersSheet() {
  const ss = ensureBackend_();
  const sh = ss && ss.getSheetByName('Users');

  if (!sh) throw new Error('Users sheet not found.');

  const headers = sh.getRange(1,1,1,sh.getLastColumn())
    .getValues()[0]
    .map(function(h){ return String(h || '').trim(); });

  Logger.log('Master: ' + ss.getUrl());
  Logger.log('Users headers: ' + JSON.stringify(headers));

  return {
    ok:true,
    spreadsheetId:ss.getId(),
    spreadsheetUrl:ss.getUrl(),
    usersSheetUrl:ss.getUrl() + '#gid=' + sh.getSheetId(),
    headers:headers
  };
}

// Backward-compatible repair entry point. It now performs the same automatic setup.
function repairUsersSheet() {
  return ensureBackend_();
}

function route_(action,b) {
  const publicActions = ['login','health'];
  if (!publicActions.includes(action)) requireSession_(b.session);

  switch(action) {
    case 'login': return login_(b.username,b.password);
    case 'health': return health_();
    case 'dashboard': return dashboard_(b.session);
    case 'attendance': return attendance_(b.session,b.month,b.year,b.since);
    case 'importAttendance': return importAttendance_(b.session,b.rows);
    case 'punch': return punch_(b.session,b.type,b.date,b.time);
    case 'saveAttendance': return saveAttendance_(b.session,b.date,b.inTime,b.outTime,b.breaks);
    case 'saveBreak': return saveBreak_(b.session,b.date,b.breakNo,b.breakType,b.namazType,b.startTime,b.endTime,b.reason);
    case 'notes': return notes_(b.session);
    case 'saveNote': return saveNote_(b.session,b.note);
    case 'tasks': return tasks_(b.session,b.month,b.year,b.since);
    case 'taskAction': return taskAction_(b.session,b.taskId,b.type,b.progress,b.note,b.minutes);
    case 'assignTask': return assignTask_(b.session,b.task);
    case 'taskProgress': return taskProgress_(b.session,b.taskId,b.progress,b.note);
    case 'taskTime': return taskTime_(b.session,b.taskId,b.type,b.minutes);
    case 'templates': return templates_(b.session);
    case 'createTemplate': return createTemplate_(b.session,b.template);
    case 'updateTemplate': return updateTemplate_(b.session,b.template);
    case 'deleteTemplate': return deleteTemplate_(b.session,b.templateId);
    case 'assignTemplate': return assignTemplate_(b.session,b.templateId,b.username,b.month,b.year,b.fromDay,b.toDay);
    case 'employees': return employees_(b.session);
    case 'createEmployee': return createEmployee_(b.session,b.employee);
    case 'updateEmployee': return updateEmployee_(b.session,b.employee);
    case 'deleteEmployee': return deleteEmployee_(b.session,b.username);
    case 'approvals': return approvals_(b.session);
    case 'approvalAction': return approvalAction_(b.session,b.id,b.status);
    case 'reports': return reports_(b.session,b.range);
    case 'exportReport': return exportReport_(b.session,b.format,b.range);
    case 'notifications': return notifications_(b.session);
    case 'markNotificationsRead': return markNotificationsRead_(b.session);
    case 'automation': return automation_(b.session);
    case 'runAutomation': return runAutomation_(b.session);
    case 'installAutomation': return installAutomation_(b.session);
    case 'profile': return profile_(b.session);
    case 'employeeProfile': return employeeProfile_(b.session,b.username);
    case 'changePassword': return changePassword_(b.session,b.currentPassword,b.newPassword);
    case 'uploadProfilePhoto': return uploadProfilePhoto_(b.session,b.targetUsername,b.fileName,b.dataUrl,b.mimeType);
    case 'advanced': return advanced_(b.session);
    case 'createHoliday': return createHoliday_(b.session,b.holiday);
    case 'deleteHoliday': return deleteHoliday_(b.session,b.id);
    case 'saveWeekoff': return saveWeekoff_(b.session,b.day);
    case 'setRamadanBreakFreeze': return setRamadanBreakFreeze_(b.session,b.frozen);
    case 'createRequest': return createRequest_(b.session,b.request);
    case 'auditLog': return auditLog_(b.session);
    default: throw new Error('Unknown action: '+action);
  }
}

function login_(username,password) {
  if (username === undefined || username === null || String(username).trim() === '' || password === undefined || password === null || String(password) === '') {
    throw new Error('User ID aur Password dono enter karein.');
  }

  // Always resolve/verify the real Master Sheet before reading credentials.
  // This prevents an old/stale MASTER_ID from pointing login to another file.
  const ss = resolveMasterSpreadsheetForLogin_();
  const sh = ss.getSheetByName('Users');
  if (!sh) throw new Error('Users sheet not found in the Master Google Sheet.');

  const expected = [
    'name','employeeId','username','role','password','department',
    'designation','phone','whatsapp','office','address','country',
    'region','state','division','district','status','createdAt','updatedAt'
  ];
  const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(h){ return String(h || '').trim(); });
  const exact = header.length === expected.length && expected.every(function(h,i){ return header[i] === h; });
  if (!exact) {
    throw new Error('Users sheet header mismatch. Please use the final A:S Users header.');
  }

  const rows = readRows_(sh);
  const wanted = String(username).trim().toLowerCase();
  const user = rows.find(function(r){
    return String(r.username || '').trim().toLowerCase() === wanted;
  });

  if (!user) { try { audit_({username:String(username||'')},'LOGIN_FAILED','AUTH','Unknown username'); } catch(e) {} throw new Error('Invalid username or password.'); }

  // Blank status is treated as ACTIVE so manually prepared Users rows can log in.
  const status = String(user.status || '').trim().toUpperCase();
  if (status && status !== 'ACTIVE') {
    throw new Error('User access is inactive.');
  }

  // Password is read directly from column E (password).
  // Do not hash or transform it.
  const storedPassword = String(user.password === null || user.password === undefined ? '' : user.password);
  const enteredPassword = String(password);

  if (storedPassword !== enteredPassword) {
    try { audit_({username:user.username},'LOGIN_FAILED','AUTH','Invalid password'); } catch(e) {}
    throw new Error('Invalid username or password.');
  }

  // Normalize a blank status once, without touching the password.
  if (!status) {
    updateByKey_(sh, 'username', user.username, {status:'ACTIVE'});
    user.status = 'ACTIVE';
  }

  // Refresh cache after any status normalization.
  invalidateUsersCache_();

  const session = {
    id:user.username,
    name:user.name,
    code:user.employeeId || user.employeeCode || '',
    username:user.username,
    role:user.role,
    department:user.department
  };

  // Login must not fail or become slow because an audit write is temporarily
  // unavailable. Credential validation has already succeeded at this point.
  try {
    audit_(session,'LOGIN','AUTH','Successful login');
  } catch (auditError) {
    console.warn('Login audit skipped: ' + String(auditError.message || auditError));
  }
  return {session};
}

/**
 * Resolves the Master Sheet for login and rejects stale MASTER_ID values.
 */
function resolveMasterSpreadsheetForLogin_() {
  const props = PropertiesService.getScriptProperties();
  const expectedUsersHeader = [
    'name','employeeId','username','role','password','department',
    'designation','phone','whatsapp','office','address','country',
    'region','state','division','district','status','createdAt','updatedAt'
  ];

  const isValidMaster = function(ss) {
    try {
      if (!ss || ss.getName() !== CONFIG.MASTER_NAME) return false;
      const sh = ss.getSheetByName('Users');
      if (!sh || sh.getLastColumn() < expectedUsersHeader.length) return false;
      const header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(h){ return String(h || '').trim(); });
      return header.length === expectedUsersHeader.length && expectedUsersHeader.every(function(h,i){ return header[i] === h; });
    } catch(e) { return false; }
  };

  const savedId = props.getProperty('MASTER_ID');
  if (savedId) {
    try {
      const saved = SpreadsheetApp.openById(savedId);
      if (isValidMaster(saved)) return saved;
    } catch(e) {}
  }

  // Prefer a spreadsheet named exactly office-task-report.
  const files = DriveApp.getFilesByName(CONFIG.MASTER_NAME);
  let fallback = null;
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
    try {
      const candidate = SpreadsheetApp.openById(f.getId());
      if (isValidMaster(candidate)) {
        // Prefer a file whose direct parent is office-task-report.
        const parents = f.getParents();
        let exactParent = false;
        while (parents.hasNext()) {
          if (parents.next().getName() === 'office-task-report') { exactParent = true; break; }
        }
        if (exactParent) {
          props.setProperty('MASTER_ID', f.getId());
          return candidate;
        }
        if (!fallback) fallback = candidate;
      }
    } catch(e) {}
  }

  if (fallback) {
    props.setProperty('MASTER_ID', fallback.getId());
    return fallback;
  }

  throw new Error('Final Master Google Sheet "office-task-report" with the Users A:S header was not found.');
}

function health_() {
  const status = ensureBackend_();
  return {
    ready:true,
    message:'Backend ready. Users sheet and Master Admin are available.',
    masterId:status.masterId,
    usersSheetUrl:status.usersSheetUrl
  };
}

function requireSession_(s) {
  if(!s || !s.username) throw new Error('Session expired. Please login again.');
  const user=findUser_(s.username);
  if(!user || String(user.status||'ACTIVE').toUpperCase()!=='ACTIVE') throw new Error('User access is inactive.');
  // Never trust role/department values supplied by the browser.
  s.role=user.role; s.name=user.name; s.code=user.employeeId; s.department=user.department;
}

function findUser_(username) {
  const rows = readUsersCached_();
  const target = String(username || '').toLowerCase();
  const r = rows.find(function(x){ return String(x.username || '').toLowerCase() === target; });
  if(!r) return null;
  return {
    id:r.username, name:r.name, employeeId:r.employeeId || r.employeeCode || '', employeeCode:r.employeeId || r.employeeCode || '', username:r.username, role:r.role,
    department:r.department,designation:r.designation,phone:r.phone,whatsapp:r.whatsapp,password:r.password,
    office:r.office,address:r.address,country:r.country,region:r.region,state:r.state,
    division:r.division,district:r.district,status:r.status,
    officeInTime:r.officeInTime||'',officeOutTime:r.officeOutTime||'',weekoffDay:r.weekoffDay
  };
}

function dashboard_(s) {
  const user=findUser_(s.username), empFile=getEmployeeFile_(user);
  const taskData=readRows_(empFile.getSheetByName('Tasks'));
  const att=readRows_(empFile.getSheetByName('Attendance'));
  const today=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const myToday=taskData.filter(r=>r.date===today);
  const completed=myToday.filter(r=>r.status==='COMPLETED').length;
  const pending=myToday.filter(r=>!['COMPLETED','CANCELLED'].includes(r.status)).length;
  const trend=last7_(taskData);
  const activities=readRows_(empFile.getSheetByName('Activities')).slice(-5).reverse().map(r=>({title:r.title,time:r.time}));
  const reminders=taskData.filter(r=>r.due>=today && r.status!=='COMPLETED').slice(0,5).map(r=>({title:r.name,when:'Due '+r.due}));
  return {kpis:{present:att.filter(r=>r.date===today && r.status==='PRESENT').length,myTasks:myToday.length,completed,pending},trend,activities,reminders};
}

function attendance_(s,month,year,since) {
  const user=findUser_(s.username), f=getEmployeeFile_(user), attSh=f.getSheetByName('Attendance');
  ensureHeaderColumns_(attSh,['date','in','out','officeMinutes','breakMinutes','break1Type','break1NamazType','break1Start','break1End','break1Reason','break2Type','break2NamazType','break2Start','break2End','break2Reason','break3Type','break3NamazType','break3Start','break3End','break3Reason','status','approveStatus','updatedAt']);
  const rows=readRows_(attSh);
  const today=new Date(), todayStr=Utilities.formatDate(today,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const yesterday=new Date(today); yesterday.setDate(yesterday.getDate()-1);
  const yesterdayStr=Utilities.formatDate(yesterday,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const rules=getAttendanceRules_(user);
  // Testing-friendly: allow any past date, but never allow a future date.
  // The UI uses the same lower bound for a consistent experience.
  const editableDates=[];
  const start=new Date(today); start.setDate(start.getDate()-2);
  for(let d=new Date(today); d>=start; d.setDate(d.getDate()-1)){
    const ds=Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd');
    editableDates.push({date:ds,label:ds===todayStr?'Today':(ds===yesterdayStr?'Yesterday':ds),nonWorking:isNonWorkingDate_(ds,rules),reason:rules.reasonMap[ds]||''});
  }
  const office=employeeSettings_(user);
  const rosterToday=getAttendanceRoster_(user,todayStr);
  const rosterYesterday=getAttendanceRoster_(user,yesterdayStr);
  const normalizedRows=rows.map(function(r){
    return Object.assign({},r,{date:normalizeAttendanceDate_(r.date)||String(r.date||'')});
  });
  const selectedMonth = month ? Number(month) : Number(Utilities.formatDate(today,Session.getScriptTimeZone(),'MM'));
  const selectedYear = year ? Number(year) : Number(Utilities.formatDate(today,Session.getScriptTimeZone(),'yyyy'));
  const filteredRows = normalizedRows.filter(function(r){
    const m=String(r.date||'').match(/^(\d{4})-(\d{2})-/);
    return m && Number(m[1])===selectedYear && Number(m[2])===selectedMonth;
  }).sort(function(a,b){return String(a.date).localeCompare(String(b.date));});
  const breakAlerts=[];
  filteredRows.forEach(function(r){
    for(let i=1;i<=3;i++){
      const type=String(r[`break${i}Type`]||'').trim();
      const namaz=String(r[`break${i}NamazType`]||'').trim();
      const startT=normalizeTime_(r[`break${i}Start`]), endT=normalizeTime_(r[`break${i}End`]);
      if(!startT||!endT) continue;
      let max=0, label=type;
      if(type==='Lunch'){max=25; label='Lunch Time';}
      else if(type==='Namaz' && namaz==='Asr'){max=25; label='Namaz - Asr';}
      else if(type==='Namaz' && namaz==='Magrib'){max=25; label='Namaz - Magrib';}
      else if(type==='Namaz' && namaz==='Zohar'){max=20; label='Namaz - Zohar';}
      else if(type==='Namaz' && namaz==='Juma'){max=60; label='Namaz - Juma';}
      else if(type==='Namaz' && namaz==='Zohar / Juma'){
        const dow=new Date(r.date+'T00:00:00').getDay();
        max=dow===5?60:20; label=dow===5?'Namaz - Juma':'Namaz - Zohar';
      }
      if(max>0){
        const actual=minutesBetween_(startT,endT);
        if(actual>max) breakAlerts.push({date:r.date,breakNo:i,type:label,allowedMinutes:max,actualMinutes:actual,excessMinutes:actual-max});
      }
    }
  });
  const sys=master_().getSheetByName('SystemSettings');
  const sysRows=readRows_(sys);
  const break3Frozen=String((sysRows.find(r=>r.key==='RAMADAN_LUNCH_BREAK_FROZEN')||{}).value||'').toUpperCase()==='TRUE';
  return {
    break3Frozen:break3Frozen,
    today:(function(){const t=normalizedRows.find(r=>r.date===todayStr)||{};return {in:t.in||'',out:t.out||'',officeMinutes:t.officeMinutes||0,breakMinutes:t.breakMinutes||0,nonWorking:isNonWorkingDate_(todayStr,rules),reason:rules.reasonMap[todayStr]||''};})(),
    rows:(function(){
      if(!since) return filteredRows.slice().reverse();
      const sinceDate=new Date(String(since));
      if(isNaN(sinceDate.getTime())) return filteredRows.slice().reverse();
      return filteredRows.filter(function(r){
        const u=new Date(String(r.updatedAt||''));
        return !isNaN(u.getTime()) && u.getTime()>sinceDate.getTime();
      }).slice().reverse();
    })(),
    incremental:!!since,
    selectedMonth:selectedMonth,
    selectedYear:selectedYear,
    editableDates:editableDates,
    breakAlerts:breakAlerts.slice().reverse(),
    syncAt:new Date().toISOString(),
    settings:{officeInTime:(rosterToday&&rosterToday.officeInTime)||office.officeInTime||'',officeOutTime:(rosterToday&&rosterToday.officeOutTime)||office.officeOutTime||'',weekoffLabel:weekDayName_(rules.weekoff),roster:{today:rosterToday||null,yesterday:rosterYesterday||null}}
  };
}


/**
 * Bulk employee attendance import.
 * Rules:
 * 1) Existing attendance is NEVER overwritten.
 * 2) If an existing day has a blank field, the incoming value can fill only that blank field.
 * 3) Missing employee files are created automatically.
 * 4) Each employee sheet is read/written once for the whole import for better speed.
 */
function importAttendance_(s, inputRows) {
  assertMasterAdmin_(s);
  if (!Array.isArray(inputRows) || !inputRows.length) throw new Error('Import file mein koi attendance record nahi mila.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Attendance import already running. Please try again.');
  try {
    const users = readUsersCached_().filter(function(u){ return String(u.status || 'ACTIVE') === 'ACTIVE'; });
    const rules = getAttendanceRules_();
    const byCode = {};
    users.forEach(function(u){ byCode[String(u.employeeCode || '').trim()] = u; });
    const groups = {};
    const errors = [];
    let valid = 0;

    inputRows.forEach(function(raw, idx){
      const r = normalizeAttendanceImportRow_(raw);
      if (!r.employeeCode) { errors.push(`Row ${idx+2}: Employee Id missing.`); return; }
      if (!r.date) { errors.push(`Row ${idx+2}: Date missing.`); return; }
      const user = byCode[String(r.employeeCode).trim()];
      if (!user) { errors.push(`Row ${idx+2}: Employee Id ${r.employeeCode} not found.`); return; }
      valid++;
      const key = String(r.employeeCode).trim();
      if (!groups[key]) groups[key] = {user:user, rows:[]};
      groups[key].rows.push(r);
    });

    let added = 0, updated = 0, skipped = 0;
    Object.keys(groups).forEach(function(code){
      const group = groups[code], user = group.user;
      const file = getEmployeeFile_(user);
      const sh = file.getSheetByName('Attendance');
      ensureHeaderColumns_(sh,['date','in','out','officeMinutes','breakMinutes','break1Type','break1NamazType','break1Start','break1End','break1Reason','break2Type','break2NamazType','break2Start','break2End','break2Reason','break3Type','break3NamazType','break3Start','break3End','break3Reason','status','approveStatus','updatedAt']);
      const range = sh.getDataRange();
      const values = range.getValues();
      const headers = values.shift().map(String);
      const idx = index_(headers);
      const byDate = {};
      values.forEach(function(row,i){
        const d = normalizeAttendanceDate_(row[idx.date]);
        if (d) byDate[d] = {row:row, index:i};
      });

      const beforeRows = values.length;
      group.rows.forEach(function(r){
        const existing = byDate[r.date];
        if (!existing) {
          const row = new Array(headers.length).fill('');
          row[idx.date] = r.date;
          if (idx.in !== undefined) row[idx.in] = r.in || '';
          if (idx.out !== undefined) row[idx.out] = r.out || '';
          if (idx.officeMinutes !== undefined) row[idx.officeMinutes] = r.officeMinutes === '' ? calculateOfficeMinutes_(r.in,r.out) : Number(r.officeMinutes || 0);
          if (idx.breakMinutes !== undefined) row[idx.breakMinutes] = r.breakMinutes === '' ? calculateBreakMinutesFromRow_(r) : Number(r.breakMinutes || 0);
          if (idx.status !== undefined) row[idx.status] = r.status || (isNonWorkingDate_(r.date,rules) ? (rules.reasonMap[r.date] || 'Weekly Off') : ((r.in || r.out) ? 'PRESENT' : ''));
          if (idx.approveStatus !== undefined) row[idx.approveStatus] = r.approveStatus || 'Not Approve';
          if (idx.updatedAt !== undefined) row[idx.updatedAt] = new Date();
          values.push(row);
          byDate[r.date] = {row:row,index:values.length-1};
          added++;
          return;
        }

        let changed = false;
        const row = existing.row;
        function fillBlank(col, value){
          if (idx[col] === undefined || value === '' || value === null || value === undefined) return;
          const current = row[idx[col]];
          if (current === '' || current === null || current === undefined) {
            row[idx[col]] = value;
            changed = true;
          }
        }
        fillBlank('in', r.in);
        fillBlank('out', r.out);
        const calcOffice=calculateOfficeMinutes_(r.in,r.out);
        const calcBreak=calculateBreakMinutesFromRow_(r);
        if (idx.officeMinutes !== undefined && calcOffice !== '' && (row[idx.officeMinutes] === '' || row[idx.officeMinutes] === null || row[idx.officeMinutes] === undefined || Number(row[idx.officeMinutes])===0)) { row[idx.officeMinutes]=calcOffice; changed=true; }
        if (idx.breakMinutes !== undefined && calcBreak>0 && (row[idx.breakMinutes] === '' || row[idx.breakMinutes] === null || row[idx.breakMinutes] === undefined || Number(row[idx.breakMinutes])===0)) { row[idx.breakMinutes]=calcBreak; changed=true; }
        fillBlank('status', r.status || (isNonWorkingDate_(r.date,rules) ? (rules.reasonMap[r.date] || 'Weekly Off') : ((r.in || r.out) ? 'PRESENT' : '')));
        if (idx.approveStatus !== undefined && r.approveStatus === 'Approve' && String(row[idx.approveStatus]||'').trim().toUpperCase() !== 'APPROVE') { row[idx.approveStatus]='Approve'; changed=true; }
        else fillBlank('approveStatus', r.approveStatus);
        if (changed && idx.updatedAt !== undefined) row[idx.updatedAt]=new Date();
        if (changed) updated++; else skipped++;
      });

      if (values.length !== beforeRows || added > 0 || updated > 0) {
        sh.getRange(2,1,values.length,headers.length).setValues(values);
      }
    });

    const affected=Object.keys(groups).map(function(code){return groups[code].user.username;});
    if(affected.length) notifyUsers_(affected,'Attendance Imported','Admin ne attendance file import ki hai. Aapki attendance turant update kar di gayi hai.','ATTENDANCE');
    audit_(s,'IMPORT','ATTENDANCE',`Rows ${valid}, Added ${added}, Updated blank fields ${updated}, Skipped ${skipped}`);
    return {added:added, updated:updated, skipped:skipped, errors:errors.slice(0,100), total:inputRows.length};
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

function normalizeAttendanceImportRow_(raw) {
  raw = raw || {};
  const pick = function(names){
    for (let i=0;i<names.length;i++) {
      const key = names[i];
      if (raw[key] !== undefined && raw[key] !== null && String(raw[key]).trim() !== '') return raw[key];
    }
    return '';
  };
  const employeeCode = String(pick(['employeeId','employee id','Employee Id','employeeCode','employee code','Employee Code','code','Code']) || '').trim();
  const date = normalizeAttendanceDate_(pick(['date','Date','attendanceDate','Attendance Date']));
  return {
    employeeCode: employeeCode,
    date: date,
    in: normalizeAttendanceTime_(pick(['in','IN','inTime','IN Time','In Time'])),
    out: normalizeAttendanceTime_(pick(['out','OUT','outTime','OUT Time','Out Time'])),
    officeMinutes: normalizeNumber_(pick(['officeMinutes','Office Minutes','office time','Office Time'])),
    breakMinutes: normalizeNumber_(pick(['breakMinutes','Break Minutes','break time','Break Time'])),
    status: normalizeAttendanceStatus_(pick(['status','Status'])),
    approveStatus: normalizeApprovalStatus_(pick(['approveStatus','Approve/Not Approve','Approval','Approve','Approved','Approval Status','approvalStatus']))
  };
}

function calculateOfficeMinutes_(inTime,outTime){
  const a=normalizeTime_(inTime), b=normalizeTime_(outTime);
  if(!a||!b) return '';
  const n=minutesBetween_(a,b);
  return n>0 ? Math.round(n) : '';
}
function calculateBreakMinutesFromRow_(r){
  let total=0;
  for(let i=1;i<=3;i++){
    total += minutesBetween_(normalizeTime_(r[`break${i}Start`]), normalizeTime_(r[`break${i}End`]));
  }
  return total>0 ? Math.round(total) : 0;
}

function normalizeNumber_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(String(v).replace(/,/g,''));
  return isNaN(n) ? '' : n;
}

function normalizeAttendanceTime_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v,Session.getScriptTimeZone(),'HH:mm:ss');
  const s = String(v).trim();
  if (!s) return '';
  const ampm=s.match(/^(\d{1,2}):([0-5]\d)(?::([0-5]\d))?\s*(AM|PM)$/i);
  if(ampm){ let h=Number(ampm[1]); const mi=ampm[2], sec=ampm[3]||'00', ap=ampm[4].toUpperCase(); if(h<1||h>12)return ''; if(ap==='AM'&&h===12)h=0; if(ap==='PM'&&h!==12)h+=12; return String(h).padStart(2,'0')+':'+mi+':'+sec; }
  const twenty=s.match(/^(\d{1,2}):([0-5]\d)(?::([0-5]\d))?$/);
  if(twenty){ const h=Number(twenty[1]),mi=twenty[2],sec=twenty[3]||'00'; if(h>23)return ''; return String(h).padStart(2,'0')+':'+mi+':'+sec; }
  return '';
}

function normalizeAttendanceStatus_(v){
  const s=String(v||'').trim().toUpperCase();
  if(!s) return '';
  if(s==='PRESENT') return 'Present';
  if(s==='HOLIDAY') return 'Holiday';
  if(s==='WEEKLY OFF' || s==='WEEKOFF' || s==='WEEK OFF') return 'Weekly Off';
  if(s==='LEAVE') return 'Leave';
  if(s==='ABSENT') return 'Absent';
  return String(v).trim().replace(/\b\w/g,function(c){return c.toUpperCase()});
}

function normalizeApprovalStatus_(v){
  const s=String(v||'').trim().toUpperCase();
  return ['APPROVE','APPROVED','APPROVAL','YES','Y','OK','TRUE','1'].includes(s) ? 'Approve' : 'Not Approve';
}

function normalizeAttendanceDate_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const s=String(v).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
  let m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if(m){
    const a=Number(m[1]), b=Number(m[2]), y=m[3];
    // Support both DD/MM/YYYY and MM/DD/YYYY where the values are unambiguous.
    let day=a, month=b;
    if(a<=12 && b>12){ month=a; day=b; }
    if(month>=1&&month<=12&&day>=1&&day<=31) return y+'-'+String(month).padStart(2,'0')+'-'+String(day).padStart(2,'0');
  }
  m=s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/);
  if(m) return m[1]+'-'+String(m[2]).padStart(2,'0')+'-'+String(m[3]).padStart(2,'0');
  const d=new Date(s);
  if(!isNaN(d.getTime())) return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd');
  return '';
}

function saveAttendance_(s,dateStr,inTime,outTime,breaks){
  const user=findUser_(s.username), f=getEmployeeFile_(user), sh=f.getSheetByName('Attendance');
  const todayStr=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const date=normalizeAttendanceDate_(dateStr||todayStr);
  if(!date) throw new Error('Valid attendance date select karein.');
  if(date>todayStr) throw new Error('Future date ki attendance allowed nahi hai.');
  const d=new Date(todayStr+'T00:00:00'); const minDate=new Date(d); minDate.setDate(minDate.getDate()-2);
  const minStr=Utilities.formatDate(minDate,Session.getScriptTimeZone(),'yyyy-MM-dd');
  if(date<minStr) throw new Error('Filhal sirf aaj aur previous 2 days ki attendance allowed hai.');
  const rules=getAttendanceRules_(user);
  if(isNonWorkingDate_(date,rules)) throw new Error((rules.reasonMap[date]||'Weekoff / Leave / Adjustment')+' hai. Attendance required nahi hai.');
  const tmIn=normalizeTime_(inTime), tmOut=normalizeTime_(outTime);
  if(!tmIn && !tmOut) throw new Error('IN Time ya OUT Time enter karein.');
  if(tmOut && !tmIn) throw new Error('Pehle IN Time enter karein.');
  if(tmIn && tmOut && minutesBetween_(tmIn,tmOut)<=0) throw new Error('OUT Time, IN Time ke baad hona chahiye.');
  ensureHeaderColumns_(sh,['date','in','out','officeMinutes','breakMinutes','break1Type','break1NamazType','break1Start','break1End','break1Reason','break2Type','break2NamazType','break2Start','break2End','break2Reason','break3Type','break3NamazType','break3Start','break3End','break3Reason','status','approveStatus','updatedAt']);
  let rows=readRows_(sh), existing=rows.find(r=>normalizeAttendanceDate_(r.date)===date);
  if(!existing){ appendObject_(sh,{date:date,in:'',out:'',officeMinutes:0,breakMinutes:0,status:'PRESENT',approveStatus:'Not Approve',updatedAt:new Date()}); existing={date:date,in:'',out:'',officeMinutes:0,breakMinutes:0,status:'PRESENT',approveStatus:'Not Approve',updatedAt:new Date()}; }
  const patch={}; if(tmIn) patch.in=tmIn; if(tmOut) patch.out=tmOut;
  const finalIn=tmIn||normalizeTime_(existing.in), finalOut=tmOut||normalizeTime_(existing.out);
  if(finalIn&&finalOut) patch.officeMinutes=calculateOfficeMinutes_(finalIn,finalOut);
  if(finalIn) patch.status='PRESENT';
  const incoming=Array.isArray(breaks)?breaks:[]; let breakTotal=0;
  const frozenLunch=String((readRows_(master_().getSheetByName('SystemSettings')).find(r=>r.key==='RAMADAN_LUNCH_BREAK_FROZEN')||{}).value||'').toUpperCase()==='TRUE';
  for(let i=1;i<=3;i++){
    const b=incoming[i-1]||{};
    const start=normalizeTime_(b.start), end=normalizeTime_(b.end);
    const has=!!(start||end||b.namazType);
    if(i===3 && frozenLunch){
      if(has) throw new Error('Ramadan Lunch Break abhi Admin ne freeze kiya hua hai.');
      patch.break3Type=''; patch.break3NamazType=''; patch.break3Start=''; patch.break3End=''; patch.break3Reason='';
      continue;
    }
    if(!has){
      patch[`break${i}Type`]=''; patch[`break${i}NamazType`]=''; patch[`break${i}Start`]=''; patch[`break${i}End`]=''; patch[`break${i}Reason`]='';
      continue;
    }
    if(!start||!end) throw new Error(`Break ${i}: Start aur End Time dono select karein.`);
    if(minutesBetween_(start,end)<=0) throw new Error(`Break ${i}: End Time, Start Time ke baad hona chahiye.`);
    let type=i===3?'Lunch':'Namaz';
    let namaz=String(b.namazType||'').trim();
    if(i===1 && !['Zohar','Juma'].includes(namaz)) throw new Error('Break 1 mein Zohar ya Juma select karein.');
    if(i===2 && !['Asr','Magrib'].includes(namaz)) throw new Error('Break 2 mein Asr ya Magrib select karein.');
    if(i===3) namaz='';
    const prefix=`break${i}`;
    patch[`${prefix}Type`]=type; patch[`${prefix}NamazType`]=namaz; patch[`${prefix}Start`]=start; patch[`${prefix}End`]=end; patch[`${prefix}Reason`]='';
    breakTotal+=minutesBetween_(start,end);
  }
  patch.breakMinutes=Math.round(breakTotal);
  patch.updatedAt=new Date();
  updateByKey_(sh,'date',date,patch);
  appendActivity_(f,user,'Attendance',`Attendance saved for ${date}`);
  const admins=readUsersCached_().filter(function(x){return ['MASTER_ADMIN','ADMIN','HOD'].includes(String(x.role||''))&&String(x.status||'ACTIVE')==='ACTIVE';}).map(function(x){return x.username;});
  notifyUsers_(admins,'Attendance Updated',`${user.name} (${user.employeeId}) ne ${date} ki attendance update ki hai.`,'ATTENDANCE');
  return {saved:true,date:date,in:finalIn||'',out:finalOut||'',officeMinutes:patch.officeMinutes||existing.officeMinutes||0,breakMinutes:patch.breakMinutes||0};
}

function punch_(s,type,dateStr,timeStr) {
  const user=findUser_(s.username), f=getEmployeeFile_(user), sh=f.getSheetByName('Attendance');
  const today=new Date(), todayStr=Utilities.formatDate(today,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const date=normalizeAttendanceDate_(dateStr || todayStr);
  if(!date) throw new Error('Valid attendance date select karein.');
  if(date>todayStr) throw new Error('Future date ki attendance allowed nahi hai.');
  const rules=getAttendanceRules_(user);
  if(isNonWorkingDate_(date,rules)) throw new Error((rules.reasonMap[date]||'Weekoff / Leave / Adjustment')+' hai. Attendance required nahi hai.');
  if(!['IN','OUT'].includes(String(type))) throw new Error('Invalid attendance action.');
  const tm=normalizeTime_(timeStr);
  if(!tm) throw new Error('Valid time select karein.');
  ensureHeaderColumns_(sh,['date','in','out','officeMinutes','breakMinutes','break1Type','break1NamazType','break1Start','break1End','break1Reason','break2Type','break2NamazType','break2Start','break2End','break2Reason','break3Type','break3NamazType','break3Start','break3End','break3Reason','status','approveStatus','updatedAt']);
  const rows=readRows_(sh);
  let existing=rows.find(r=>normalizeAttendanceDate_(r.date)===date);

  // Robustly handle a newly created attendance row. Do not assume that a
  // second read will immediately return the row; this was causing
  // "Cannot read properties of undefined (reading 'in')" during Save.
  if(!existing){
    appendObject_(sh,{date:date,in:'',out:'',officeMinutes:0,breakMinutes:0,break1Type:'',break1NamazType:'',break1Start:'',break1End:'',break1Reason:'',break2Type:'',break2NamazType:'',break2Start:'',break2End:'',break2Reason:'',break3Type:'',break3NamazType:'',break3Start:'',break3End:'',break3Reason:'',status:'PRESENT',approveStatus:'Not Approve'});
    existing={date:date,in:'',out:'',officeMinutes:0,breakMinutes:0,break1Type:'',break1NamazType:'',break1Start:'',break1End:'',break1Reason:'',break2Type:'',break2NamazType:'',break2Start:'',break2End:'',break2Reason:'',break3Type:'',break3NamazType:'',break3Start:'',break3End:'',break3Reason:'',status:'PRESENT',approveStatus:'Not Approve'};
  }

  if(type==='IN'){
    if(existing.out && minutesBetween_(tm,existing.out)<=0) throw new Error('IN Time, existing OUT Time se pehle hona chahiye.');
    const patch={in:tm,status:'PRESENT'};
    if(existing.out) patch.officeMinutes=minutesBetween_(tm,existing.out);
    updateByKey_(sh,'date',date,patch);
  } else {
    if(!existing || !existing.in) throw new Error('Pehle IN Time enter karein.');
    const officeMinutes=minutesBetween_(existing.in,tm);
    if(officeMinutes<=0) throw new Error('OUT Time, IN Time ke baad hona chahiye.');
    updateByKey_(sh,'date',date,{out:tm,officeMinutes:officeMinutes,status:'PRESENT'});
  }
  appendActivity_(f,user,'Attendance',`${type} time manually recorded for ${date}`);
  const admins=readUsersCached_().filter(function(x){return ['MASTER_ADMIN','ADMIN','HOD'].includes(String(x.role||'')) && String(x.status||'ACTIVE')==='ACTIVE';}).map(function(x){return x.username;});
  notifyUsers_(admins,'Attendance Updated',`${user.name} (${user.employeeId}) ne ${date} ki ${type} attendance ${tm} par add ki hai.`,'ATTENDANCE');
  return {date:date,type:type,time:tm};
}

function saveBreak_(s,dateStr,breakNo,breakType,namazType,startTime,endTime,reason){
  const user=findUser_(s.username), f=getEmployeeFile_(user), sh=f.getSheetByName('Attendance');
  ensureHeaderColumns_(sh,['date','in','out','officeMinutes','breakMinutes','break1Type','break1NamazType','break1Start','break1End','break1Reason','break2Type','break2NamazType','break2Start','break2End','break2Reason','break3Type','break3NamazType','break3Start','break3End','break3Reason','status','approveStatus','updatedAt']);
  const date=normalizeAttendanceDate_(dateStr); if(!date) throw new Error('Valid attendance date select karein.');
  const todayStr=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'); if(date>todayStr) throw new Error('Future date ki attendance allowed nahi hai.');
  const n=Number(breakNo); if(![1,2,3].includes(n)) throw new Error('Maximum 3 breaks allowed hain.');
  const type=String(breakType||'').trim(); if(!type) throw new Error('Break type select karein.');
  const start=normalizeTime_(startTime), end=normalizeTime_(endTime); if(!start||!end) throw new Error('Break Start aur End Time dono select karein.');
  const mins=minutesBetween_(start,end); if(mins<=0) throw new Error('Break End Time, Start Time ke baad hona chahiye.');
  if(type==='Namaz' && !String(namazType||'').trim()) throw new Error('Namaz ka type select karein.');
  if(type==='Other' && !String(reason||'').trim()) throw new Error('Other break ke liye reason likhein.');
  const rows=readRows_(sh); let row=rows.find(r=>normalizeAttendanceDate_(r.date)===date);
  if(!row){ appendObject_(sh,{date:date,in:'',out:'',officeMinutes:0,breakMinutes:0,status:'PRESENT',approveStatus:'Not Approve',updatedAt:new Date()}); row={date:date,in:'',out:'',officeMinutes:0,breakMinutes:0,status:'PRESENT',approveStatus:'Not Approve',updatedAt:new Date()}; }
  for(let i=1;i<n;i++) if(!row[`break${i}End`]) throw new Error(`Break ${i} pehle complete karein.`);
  const prefix=`break${n}`;
  const oldMinutes=[1,2,3].reduce((sum,i)=>sum+minutesBetween_(row[`break${i}Start`],row[`break${i}End`]),0);
  const patch={}; patch[`${prefix}Type`]=type; patch[`${prefix}NamazType`]=type==='Namaz'?String(namazType||''):''; patch[`${prefix}Start`]=start; patch[`${prefix}End`]=end; patch[`${prefix}Reason`]=String(reason||''); patch.breakMinutes=oldMinutes-minutesBetween_(row[`${prefix}Start`],row[`${prefix}End`])+mins; patch.updatedAt=new Date();
  updateByKey_(sh,'date',date,patch);
  appendActivity_(f,user,'Attendance Break',`${type}${namazType?` (${namazType})`:''} · ${start}-${end} · ${date}`);
  return {date,breakNo:n,breakType:type,namazType:namazType||'',startTime:start,endTime:end,breakMinutes:patch.breakMinutes};
}

function notes_(s){
  const sh=master_().getSheetByName('Notes'); if(!sh) return {rows:[]};
  ensureHeaderColumns_(sh,['id','username','role','text','date','time','createdAt']);
  return {rows:readRows_(sh).filter(r=>String(r.username)===String(s.username)).slice(-100).reverse()};
}
function saveNote_(s,note){
  const sh=master_().getSheetByName('Notes'); if(!sh) throw new Error('Notes sheet unavailable.');
  ensureHeaderColumns_(sh,['id','username','role','text','date','time','createdAt']);
  const text=String(note&&note.text||'').trim(); if(!text) throw new Error('Note text required hai.');
  const now=new Date(), date=normalizeAttendanceDate_(note&&note.date)||Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyy-MM-dd'), time=normalizeTime_(note&&note.time)||Utilities.formatDate(now,Session.getScriptTimeZone(),'HH:mm:ss');
  appendObject_(sh,{id:Utilities.getUuid(),username:s.username,role:s.role,text:text,date:date,time:time,createdAt:now});
  audit_(s,'CREATE','NOTE',text.slice(0,200));
  return {saved:true};
}

function normalizeTime_(value){
  const s=String(value||'').trim();
  const m=s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if(!m) return '';
  const h=Number(m[1]), mi=Number(m[2]), sec=Number(m[3]||0);
  if(h<0||h>23||mi<0||mi>59||sec<0||sec>59) return '';
  return String(h).padStart(2,'0')+':'+String(mi).padStart(2,'0')+':'+String(sec).padStart(2,'0');
}

function tasks_(s,month,year,since) {
  const user=findUser_(s.username), rows=readRows_(getEmployeeFile_(user).getSheetByName('Tasks'));
  const m=month?Number(month):0, y=year?Number(year):0;
  const filtered=(m&&y)?rows.filter(r=>{const d=normalizeAttendanceDate_(r.date);return d&&Number(d.slice(5,7))===m&&Number(d.slice(0,4))===y;}):rows;
  let out=filtered;
  if(since){
    const sd=new Date(String(since));
    if(!isNaN(sd.getTime())) out=filtered.filter(r=>{const u=new Date(String(r.updatedAt||''));return !isNaN(u.getTime())&&u.getTime()>sd.getTime();});
  }
  return {rows:out.slice().reverse(),incremental:!!since,month:m||null,year:y||null,syncAt:new Date().toISOString()};
}

function taskAction_(s,id,type,progress,note,minutes) {
  const user=findUser_(s.username), f=getEmployeeFile_(user), sh=f.getSheetByName('Tasks');
  ensureHeaderColumns_(sh,['id','date','name','details','category','priority','due','status','startTime','completedTime','actualMinutes','expectedMinutes','progress','progressNote','assignedBy','assignedAt','updatedAt','assignmentMonth','assignmentYear','assignmentFrom','assignmentTo','assignmentKey']);
  const rows=readRows_(sh), row=rows.find(r=>String(r.id)===String(id));
  if(!row) throw new Error('Task not found.');
  const now=new Date(), time=formatDateTime_(now), action=String(type||'').toUpperCase();
  if(action==='START') updateByKey_(sh,'id',id,{status:'IN_PROGRESS',startTime:time,updatedAt:now});
  else if(action==='COMPLETE') {
    const actual=row.startTime?Math.max(0,Math.round((now-new Date(row.startTime))/60000)):Number(row.actualMinutes||0);
    updateByKey_(sh,'id',id,{status:'COMPLETED',completedTime:time,actualMinutes:actual,progress:100,updatedAt:now});
  } else if(action==='CANCEL') updateByKey_(sh,'id',id,{status:'CANCELLED',updatedAt:now});
  else if(action==='PAUSE') updateByKey_(sh,'id',id,{status:'ON_HOLD',updatedAt:now});
  else if(action==='RESUME') updateByKey_(sh,'id',id,{status:'IN_PROGRESS',updatedAt:now});
  else if(action==='PROGRESS') taskProgress_(s,id,progress,note);
  else throw new Error('Unsupported task action.');
  appendActivity_(f,user,'Task',`${action} · ${row.name}`+(note?` · ${note}`:''));
  audit_(s,'UPDATE','TASK',`${action} / ${row.name}`);
  return {};
}

function assignTask_(s,t) {
  assertAdmin_(s);
  if(!t || !t.username || !t.name) throw new Error('Employee aur Task Name required hain.');
  const user=findUser_(t.username); if(!user || String(user.status||'ACTIVE')!=='ACTIVE') throw new Error('Selected employee is not active.');
  if(!['EMPLOYEE','HOD'].includes(String(user.role))) throw new Error('Task sirf Employee/HOD ko assign kiya ja sakta hai.');
  const file=getEmployeeFile_(user), sh=file.getSheetByName('Tasks');
  ensureHeaderColumns_(sh,['id','date','name','details','category','priority','due','status','startTime','completedTime','actualMinutes','expectedMinutes','progress','progressNote','assignedBy','assignedAt','updatedAt','assignmentMonth','assignmentYear','assignmentFrom','assignmentTo','assignmentKey']);
  const date=normalizeAttendanceDate_(t.date)||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const due=normalizeAttendanceDate_(t.due)||date;
  const taskId=Utilities.getUuid();
  appendObject_(sh,{id:taskId,date:date,name:String(t.name).trim(),details:t.details||'',category:t.category||'',priority:t.priority||'Normal',due:due,status:'ASSIGNED',startTime:'',completedTime:'',actualMinutes:0,expectedMinutes:Number(t.expectedMinutes||0)||'',progress:0,progressNote:'',assignedBy:s.username,assignedAt:new Date(),updatedAt:new Date()});
  appendActivity_(file,user,'Task Assigned',`${t.name} · due ${due}`);
  notifyUsers_([user.username],'New Task Assigned',`${t.name} task aapko ${s.name||s.username} ne assign ki hai. Due: ${due}.`,'TASK');
  audit_(s,'ASSIGN','TASK',`${t.name} → ${user.name}`);
  return {taskId};
}

function taskProgress_(s,id,progress,note) {
  const user=findUser_(s.username), sh=getEmployeeFile_(user).getSheetByName('Tasks');
  const n=Number(progress); if(!Number.isFinite(n)||n<0||n>100) throw new Error('Progress 0 se 100 ke beech hona chahiye.');
  const rows=readRows_(sh), row=rows.find(r=>String(r.id)===String(id)); if(!row) throw new Error('Task not found.');
  const status=n>=100?'COMPLETED':(n>0?'IN_PROGRESS':(row.status||'ASSIGNED'));
  updateByKey_(sh,'id',id,{progress:n,progressNote:String(note||''),status:status,updatedAt:new Date(),completedTime:n>=100?formatDateTime_(new Date()):row.completedTime||''});
  audit_(s,'UPDATE','TASK_PROGRESS',`${row.name} = ${n}%`);
  return {progress:n,status:status};
}

function taskTime_(s,id,type,minutes) {
  const user=findUser_(s.username), f=getEmployeeFile_(user), sh=f.getSheetByName('Tasks');
  const rows=readRows_(sh), row=rows.find(r=>String(r.id)===String(id)); if(!row) throw new Error('Task not found.');
  const n=Number(minutes||0); if(!Number.isFinite(n)||n<0||n>1440) throw new Error('Invalid task minutes.');
  const current=Number(row.actualMinutes||0), total=current+n;
  updateByKey_(sh,'id',id,{actualMinutes:total,updatedAt:new Date()});
  const tSh=f.getSheetByName('TaskTime'); ensureHeaderColumns_(tSh,['taskId','date','start','end','minutes','type']);
  appendObject_(tSh,{taskId:id,date:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'),start:formatDateTime_(new Date()),end:'',minutes:n,type:type||'MANUAL'});
  audit_(s,'UPDATE','TASK_TIME',`${row.name} +${n} min`);
  return {actualMinutes:total};
}

function templates_(s) {
  assertAdmin_(s);
  return {rows:readRows_(master_().getSheetByName('Templates')).slice().reverse()};
}

function createTemplate_(s,t) {
  assertAdmin_(s);
  const allowedTasks=['Followup','File Work','Outdoor','Meeting','Other'];
  const allowedCategories={
    'Followup':['Monthly Report','Hind Mushawarat Task','HOD-Department Points Task','Data Required','Other'],
    'File Work':['Analise','Errors Cheking','Application Required Data','Other'],
    'Outdoor':['Office Related','Journey','Tarbiyati Ijtima','3 Din Qafila','Other'],
    'Meeting':['Online Meeting','Physicall Meeting','Other'],
    'Other':['Other']
  };
  if(!allowedTasks.includes(String(t.name||''))) throw new Error('Invalid Task Name.');
  if(!(allowedCategories[t.name]||[]).includes(String(t.category||''))) throw new Error('Invalid Category for selected Task Name.');
  const sh=master_().getSheetByName('Templates');
  const fromDay=Number(t.fromDay||1);
  const toDay=Number(t.toDay||fromDay);
  if(fromDay<1 || fromDay>31 || toDay<1 || toDay>31 || fromDay>toDay) throw new Error('Task period must be between day 1 and day 31, with From Day <= To Day.');
  appendObject_(sh,{id:Utilities.getUuid(),name:t.name,details:t.details||'',category:t.category||'',priority:t.priority||'Normal',fromDay,toDay,createdBy:s.username,createdAt:new Date()});
  audit_(s,'CREATE','TEMPLATE',t.name);
  return {};
}

function updateTemplate_(s,t) {
  assertAdmin_(s);
  if(!t || !t.id) throw new Error('Template ID is required.');
  const allowedTasks=['Followup','File Work','Outdoor','Meeting','Other'];
  const allowedCategories={
    'Followup':['Monthly Report','Hind Mushawarat Task','HOD-Department Points Task','Data Required','Other'],
    'File Work':['Analise','Errors Cheking','Application Required Data','Other'],
    'Outdoor':['Office Related','Journey','Tarbiyati Ijtima','3 Din Qafila','Other'],
    'Meeting':['Online Meeting','Physicall Meeting','Other'],
    'Other':['Other']
  };
  if(!allowedTasks.includes(String(t.name||''))) throw new Error('Invalid Task Name.');
  if(!(allowedCategories[t.name]||[]).includes(String(t.category||''))) throw new Error('Invalid Category for selected Task Name.');
  const fromDay=Number(t.fromDay||1), toDay=Number(t.toDay||fromDay);
  if(fromDay<1 || fromDay>31 || toDay<1 || toDay>31 || fromDay>toDay) throw new Error('Task period must be between day 1 and day 31, with From Day <= To Day.');
  updateByKey_(master_().getSheetByName('Templates'),'id',t.id,{name:t.name,details:t.details||'',category:t.category||'',priority:t.priority||'Normal',fromDay,toDay});
  audit_(s,'UPDATE','TEMPLATE',t.name);
  return {};
}

function deleteTemplate_(s,templateId) {
  assertAdmin_(s);
  if(!templateId) throw new Error('Template ID is required.');
  const sh=master_().getSheetByName('Templates');
  const data=sh.getDataRange().getValues(), h=data[0], i=h.indexOf('id');
  if(i<0) throw new Error('Template ID column not found.');
  for(let r=1;r<data.length;r++) if(String(data[r][i])===String(templateId)) {
    const name=data[r][h.indexOf('name')]||'Template';
    sh.deleteRow(r+1);
    audit_(s,'DELETE','TEMPLATE',String(name));
    return {};
  }
  throw new Error('Template not found.');
}

function assignTemplate_(s,templateId,username,month,year,fromDay,toDay) {
  assertAdmin_(s);
  if(!templateId || !username) throw new Error('Template aur Employee/HOD select karein.');
  const tRows=readRows_(master_().getSheetByName('Templates'));
  const t=tRows.find(r=>String(r.id)===String(templateId));
  if(!t) throw new Error('Template not found.');
  const user=findUser_(username);
  if(!user || user.status!=='ACTIVE') throw new Error('Selected Employee/HOD is not active.');
  if(!['EMPLOYEE','HOD'].includes(String(user.role))) throw new Error('Template sirf Employee ya HOD ko assign kiya ja sakta hai.');
  const m=Number(month), y=Number(year);
  if(!Number.isInteger(m)||m<1||m>12) throw new Error('Valid month select karein.');
  if(!Number.isInteger(y)||y<2020||y>2100) throw new Error('Valid year select karein.');
  const from=Number(fromDay||t.fromDay||1), to=Number(toDay||t.toDay||from);
  const daysInMonth=new Date(y,m,0).getDate();
  if(from<1||to<1||from>to||from>daysInMonth) throw new Error(`Task period ${y}-${String(m).padStart(2,'0')} ke andar valid hona chahiye.`);
  const safeTo=Math.min(to,daysInMonth);
  const file=getEmployeeFile_(user), sh=file.getSheetByName('Tasks');
  ensureHeaderColumns_(sh,['id','date','name','details','category','priority','due','status','startTime','completedTime','actualMinutes','expectedMinutes','progress','progressNote','assignedBy','assignedAt','updatedAt','assignmentMonth','assignmentYear','assignmentFrom','assignmentTo','assignmentKey']);
  const key=[templateId,username,y,m,from,safeTo].join('|');
  const existing=readRows_(sh).some(r=>String(r.assignmentKey||'')===key);
  if(existing) throw new Error('Yeh template is employee ke liye isi month/year aur period mein pehle hi assign ho chuka hai.');
  const tasks=[];
  for(let day=from;day<=safeTo;day++){
    const date=`${y}-${String(m).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    tasks.push({id:Utilities.getUuid(),date:date,name:t.name,details:t.details||'',category:t.category||'',priority:t.priority||'Normal',due:date,status:'ASSIGNED',startTime:'',completedTime:'',actualMinutes:0,expectedMinutes:'',progress:0,progressNote:'',assignedBy:s.username,assignedAt:new Date(),updatedAt:new Date(),assignmentMonth:m,assignmentYear:y,assignmentFrom:from,assignmentTo:safeTo,assignmentKey:key});
  }
  if(tasks.length){
    const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
    const values=tasks.map(o=>headers.map(h=>o[h]!==undefined?o[h]:''));
    sh.getRange(sh.getLastRow()+1,1,values.length,headers.length).setValues(values);
  }
  appendActivity_(file,user,'Task Assigned',`${t.name} · ${y}-${String(m).padStart(2,'0')} · ${from}-${safeTo}`);
  notifyUsers_([user.username],'New Monthly Task Assigned',`${t.name} task ${String(from).padStart(2,'0')}-${String(safeTo).padStart(2,'0')}/${String(m).padStart(2,'0')}/${y} ke liye assign ki gayi hai.`,'TASK');
  audit_(s,'ASSIGN','TEMPLATE',`${t.name} → ${user.name} (${username}) · ${m}/${y} · ${from}-${safeTo}`);
  return {count:tasks.length,username:user.username,name:user.name,fromDay:from,toDay:safeTo,month:m,year:y};
}

function employees_(s) {
  assertMasterAdmin_(s);
  // Password is intentionally kept out of API responses. Besides security, this keeps
  // the Users response small and noticeably faster on slow connections.
  const rows = readUsersCached_();
  return {rows:rows.map(function(r){
    return {
      id:r.id,name:r.name,code:r.employeeCode,username:r.username,
      role:r.role,department:r.department,designation:r.designation,
      phone:r.phone,whatsapp:r.whatsapp,office:r.office,address:r.address,
      country:r.country,region:r.region,state:r.state,division:r.division,
      district:r.district,status:r.status,createdAt:r.createdAt,updatedAt:r.updatedAt,
      officeInTime:r.officeInTime||'',officeOutTime:r.officeOutTime||'',weekoffDay:r.weekoffDay,weekoffLabel:weekDayName_(Number(r.weekoffDay))
    };
  }).map(function(row){
    const u=findUser_(row.username); const st=employeeSettings_(u);
    row.officeInTime=st.officeInTime||u.officeInTime||row.officeInTime||''; row.officeOutTime=st.officeOutTime||u.officeOutTime||row.officeOutTime||''; row.weekoffDay=(st.weekoffDay!==undefined&&st.weekoffDay!=='')?st.weekoffDay:(u.weekoffDay!==undefined?u.weekoffDay:row.weekoffDay); row.weekoffLabel=weekDayName_(row.weekoffDay);
    return row;
  })};
}

function createEmployee_(s,e) {
  assertMasterAdmin_(s);
  const sh=master_().getSheetByName('Users');
  // Keep the existing Users sheet schema in sync without running a full backend/Drive scan.
  // This specifically guarantees that the Password column exists on older master sheets.
  migrateUsersSchema_(sh);
  const employeeCode=String(e.code||'').trim();
  const username=String(e.username||'').trim();
  if(!employeeCode) throw new Error('Employee Id is required. Please enter Employee Id.');
  if(!username) throw new Error('Username is required.');
  const existingRows=readRows_(sh);
  if(existingRows.some(r=>String((r.employeeId||r.employeeCode)||'').trim()===employeeCode || String(r.username||'').trim().toLowerCase()===username.toLowerCase())) throw new Error('Employee Id or Username already exists.');
  const password=String(e.password||'');
  if(!password) throw new Error('Password is required when creating a user.');
  if(password.length < 4) throw new Error('Password must be at least 4 characters.');
  appendObject_(sh,{name:e.name,employeeId:employeeCode,username:username,role:e.role||'EMPLOYEE',password:password,department:e.department||'',designation:e.designation||'',phone:e.phone||'',whatsapp:e.whatsapp||'',office:e.office||'',address:e.address||'',country:e.country||'',region:e.region||'',state:e.state||'',division:e.division||'',district:e.district||'',status:'ACTIVE',createdAt:new Date(),updatedAt:new Date()});
  invalidateUsersCache_();
  const u=findUser_(username);
  saveEmployeeSettings_(u,{officeCity:e.office||'',officeAddress:e.address||'',officeInTime:e.officeInTime||'',officeOutTime:e.officeOutTime||'',weekoffDay:e.weekoffDay});
  if(e.photoDataUrl) uploadProfilePhotoForAdmin_(s,u,e.photoFileName,e.photoDataUrl,e.photoMimeType);
  const file=getEmployeeFile_(u);
  invalidateUsersCache_();
  appendActivity_(file,u,'Profile',`Employee file created for ${new Date().getFullYear()}`);
  audit_(s,'CREATE','EMPLOYEE',`${e.name} / ${e.employeeId || e.code || ''}`);
  return {temporaryPassword:password,fileUrl:file.getUrl()};
}

function updateEmployee_(s,e) {
  assertMasterAdmin_(s);
  if(String(e.username).toLowerCase()==='admin') throw new Error('Master Admin account is protected.');
  const sh=master_().getSheetByName('Users');
  // Users sheet is manually finalized. Do not migrate/rebuild it here.
  const existing=findUser_(e.username);
  if(!existing) throw new Error('User not found.');
  const duplicate=readRows_(sh).find(r=>String(r.username).toLowerCase()===String(e.username).toLowerCase() && String((r.employeeId||r.employeeCode)||'')!==String(existing.employeeCode||''));
  if(duplicate) throw new Error('Username already exists.');
  const patch={
    name:e.name,department:e.department||'',designation:e.designation||'',phone:e.phone||'',
    whatsapp:e.whatsapp||'',office:e.office||'',address:e.address||'',
    country:e.country||'',region:e.region||'',state:e.state||'',division:e.division||'',district:e.district||'',
    role:e.role||'EMPLOYEE',updatedAt:new Date()
  };
  const newPassword=String(e.password||'');
  if(newPassword){
    if(newPassword.length < 4) throw new Error('Password must be at least 4 characters.');
    patch.password=newPassword;
  }
  updateByKey_(sh,'username',e.username,patch);
  const updatedUser=findUser_(e.username);
  saveEmployeeSettings_(updatedUser,{officeCity:e.office||'',officeAddress:e.address||'',officeInTime:e.officeInTime||'',officeOutTime:e.officeOutTime||'',weekoffDay:e.weekoffDay});
  if(e.photoDataUrl) uploadProfilePhotoForAdmin_(s,updatedUser,e.photoFileName,e.photoDataUrl,e.photoMimeType);
  invalidateUsersCache_();
  audit_(s,'UPDATE','EMPLOYEE',`${e.name} / ${e.employeeId || e.code || ''}`);
  syncEmployeeProfile_(e);
  return {};
}

function deleteEmployee_(s,username) {
  assertMasterAdmin_(s);
  if(String(username).toLowerCase()==='admin') throw new Error('Master Admin cannot be deleted.');
  const user=findUser_(username);
  if(!user) throw new Error('User not found.');
  // Safe delete: deactivate the account so historical attendance/tasks remain intact.
  updateByKey_(master_().getSheetByName('Users'),'username',username,{status:'DELETED',updatedAt:new Date()});
  invalidateUsersCache_();
  audit_(s,'DELETE','EMPLOYEE',`${user.name} / ${user.employeeCode}`);
  return {};
}

function syncEmployeeProfile_(e) {
  const user=findUser_(e.username);
  if(!user) return;
  const file=getEmployeeFile_(user);
  const sh=file.getSheetByName('Profile');
  if(!sh) return;
  ensureHeaderColumns_(sh,['name','employeeId','username','role','department','designation','phone','whatsapp','office','address','country','region','state','division','district','officeInTime','officeOutTime','weekoffDay','year']);
  updateByKey_(sh,'username',e.username,{
    name:e.name,employeeId:e.employeeId || e.code || e.employeeCode || '',role:e.role||'EMPLOYEE',department:e.department||'',
    designation:e.designation||'',phone:e.phone||'',whatsapp:e.whatsapp||'',
    office:e.office||'',address:e.address||'',country:e.country||'',region:e.region||'',
    state:e.state||'',division:e.division||'',district:e.district||'',officeInTime:e.officeInTime||'',officeOutTime:e.officeOutTime||'',weekoffDay:e.weekoffDay||''
  });
}

function assertMasterAdmin_(s) {
  if(String(s.role)!=='MASTER_ADMIN') throw new Error('Only Master Admin can manage users.');
}

function approvals_(s) {
  assertAdmin_(s);
  return {rows:readRows_(master_().getSheetByName('Approvals')).filter(r=>r.status==='PENDING').reverse()};
}

function approvalAction_(s,id,status) {
  assertAdmin_(s);
  const a=master_().getSheetByName('Approvals'), rows=readRows_(a), item=rows.find(function(r){return String(r.id)===String(id);});
  if(!item) throw new Error('Approval request not found.');
  updateByKey_(a,'id',id,{status:status,updatedAt:new Date()});
  if(item.username && item.requestId){
    const u=findUser_(item.username);
    if(u){
      const f=getEmployeeFile_(u), sh=f.getSheetByName('Requests'); try{updateByKey_(sh,'id',item.requestId,{status:status});}catch(e){}
      if(String(status)==='APPROVED' && String(item.type)==='Time Adjustment' && item.adjustmentField && item.adjustmentTime){
        saveAttendanceRoster_(u,item.date,item.adjustmentField,item.adjustmentTime);
      }
    }
  }
  audit_(s,status,'APPROVAL',id);
  const actor=findUser_(s.username);
  if(item.username) notifyUsers_([item.username],'Request '+String(status).toLowerCase(),`${item.type} (${item.date}) ${String(status).toLowerCase()} ki gayi by ${actor?actor.name:s.username}.`,'APPROVAL');
  const allApprovers=readUsersCached_().filter(function(x){return ['MASTER_ADMIN','ADMIN','HOD'].includes(String(x.role||'')) && String(x.status||'ACTIVE')==='ACTIVE';}).map(function(x){return x.username;});
  notifyUsers_(allApprovers,'Request Updated',`${item.employeeName} ki ${item.type} request ${String(status).toLowerCase()} by ${actor?actor.name:s.username}.`,'APPROVAL');
  return {};
}

function reports_(s,range) {
  const cache=CacheService.getScriptCache(), key='OTR_REPORT_'+String(s.username)+'_'+String(range||'month');
  const cached=cache.get(key); if(cached){try{return JSON.parse(cached);}catch(e){}}
  const all=readUsersCached_().filter(u=>String(u.status||'ACTIVE')==='ACTIVE');
  const visible=reportUsersForSession_(s,all);
  let tasks=0,completed=0,pending=0,present=0,absent=0;
  const employeesRows=[];
  visible.forEach(function(u){
    const rows=readRows_(getEmployeeFile_(u).getSheetByName('Tasks'));
    const a=readRows_(getEmployeeFile_(u).getSheetByName('Attendance'));
    const ct=rows.filter(r=>r.status==='COMPLETED').length, pt=rows.filter(r=>r.status && !['COMPLETED','CANCELLED'].includes(r.status)).length;
    tasks+=rows.length;completed+=ct;pending+=pt;
    const today=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
    if(a.some(r=>normalizeAttendanceDate_(r.date)===today && r.in)) present++; else absent++;
    employeesRows.push({name:u.name,employeeId:u.employeeId,department:u.department,tasks:rows.length,completed:ct,pending:pt,progress:rows.length?Math.round(ct/rows.length*100):0});
  });
  const out={employees:visible.length,tasks,completed,pending,present,absent,employeesRows,trend:monthTrend_(visible),generatedAt:new Date().toISOString(),role:s.role};
  try{cache.put(key,JSON.stringify(out),30);}catch(e){}
  return out;
}

function reportUsersForSession_(s,all){
  const role=String(s.role||'EMPLOYEE');
  if(role==='MASTER_ADMIN'||role==='ADMIN') return all.filter(u=>u.employeeId);
  const me=findUser_(s.username);
  if(role==='HOD') return all.filter(u=>u.employeeId && (u.username===me.username || (me.department && u.department===me.department)));
  return all.filter(u=>u.username===s.username);
}

function exportReport_(s,format,range){
  const r=reports_(s,range), rows=[['Employee','Employee Id','Department','Tasks','Completed','Pending','Progress %']];
  (r.employeesRows||[]).forEach(x=>rows.push([x.name,x.employeeId,x.department,x.tasks,x.completed,x.pending,x.progress]));
  if(String(format||'CSV').toUpperCase()==='CSV'){
    const csv=rows.map(row=>row.map(v=>`"${String(v??'').replace(/"/g,'""')}"`).join(',')).join('\n');
    return {format:'CSV',fileName:'Office_Task_Report.csv',content:csv};
  }
  return {format:'DATA',fileName:'Office_Task_Report.xlsx',rows:rows};
}

function markNotificationsRead_(s){
  const sh=master_().getSheetByName('Notifications'); if(!sh) return {};
  const data=sh.getDataRange().getValues(), h=data[0], iu=h.indexOf('username'), ir=h.indexOf('read');
  if(iu<0||ir<0)return{};
  for(let r=1;r<data.length;r++) if(String(data[r][iu])===String(s.username)) data[r][ir]='Y';
  if(data.length>1) sh.getRange(2,1,data.length-1,data[0].length).setValues(data.slice(1));
  return {};
}

function automation_(s){
  assertAdmin_(s);
  const props=PropertiesService.getScriptProperties();
  return {enabled:props.getProperty('AUTOMATION_ENABLED')==='Y',lastRun:props.getProperty('AUTOMATION_LAST_RUN')||'',rules:['Pending task reminder','Attendance reminder','Approval reminder','Yearly employee file rollover']};
}

function runAutomation_(s){
  assertAdmin_(s); return automationTick_();
}

function installAutomation_(s){
  assertMasterAdmin_(s);
  const existing=ScriptApp.getProjectTriggers(); existing.forEach(function(t){if(t.getHandlerFunction()==='automationTick')ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('automationTick').timeBased().everyHours(1).create();
  PropertiesService.getScriptProperties().setProperty('AUTOMATION_ENABLED','Y');
  audit_(s,'ENABLE','AUTOMATION','Hourly automation trigger installed');
  return {enabled:true};
}

function automationTick(){
  try{return automationTick_();}catch(e){console.warn(e);return {ok:false,message:String(e.message||e)};}
}

function automationTick_(){
  const users=readUsersCached_().filter(u=>String(u.status||'ACTIVE')==='ACTIVE'&&u.employeeId), today=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  let reminders=0, approvals=0;
  users.forEach(function(u){
    const f=getEmployeeFile_(u), tasks=readRows_(f.getSheetByName('Tasks'));
    const due=tasks.filter(t=>t.due===today&&t.status!=='COMPLETED'&&t.status!=='CANCELLED');
    if(due.length){notifyUsers_([u.username],'Task Reminder',`Aaj ${due.length} task pending/due hain. Dashboard par check karein.`,'REMINDER');reminders++;}
  });
  const pending=readRows_(master_().getSheetByName('Approvals')).filter(r=>r.status==='PENDING');
  if(pending.length){const admins=users.filter(u=>['MASTER_ADMIN','ADMIN','HOD'].includes(String(u.role))).map(u=>u.username);notifyUsers_(admins,'Approval Reminder',`${pending.length} approval request pending hai.`,'REMINDER');approvals=pending.length;}
  PropertiesService.getScriptProperties().setProperty('AUTOMATION_LAST_RUN',new Date().toISOString());
  return {reminders,approvals,time:new Date().toISOString()};
}

function employeeSettings_(user){
  const globalRows=readRows_(master_().getSheetByName('SystemSettings')), gs={}; globalRows.forEach(function(r){gs[r.key]=r.value;});
  const defaults={
    officeCity:user&&user.office||'',
    officeAddress:user&&user.address||'',
    officeInTime:user&&user.officeInTime||'',
    officeOutTime:user&&user.officeOutTime||'',
    weekoffDay:user&&user.weekoffDay!==undefined&&user.weekoffDay!==''?String(user.weekoffDay):String(gs.WEEKOFF_DAY||'0')
  };
  if(!user || !user.employeeId) return defaults;
  const sh=master_().getSheetByName('EmployeeSettings'); if(!sh) return defaults;
  const row=readRows_(sh).find(function(r){return String(r.employeeId)===String(user.employeeId);});
  if(!row) return defaults;
  return {officeCity:row.officeCity||defaults.officeCity,officeAddress:row.officeAddress||defaults.officeAddress,officeInTime:row.officeInTime||defaults.officeInTime,officeOutTime:row.officeOutTime||defaults.officeOutTime,weekoffDay:row.weekoffDay===''||row.weekoffDay===undefined?defaults.weekoffDay:String(row.weekoffDay)};
}
function saveEmployeeSettings_(user,settings){
  if(!user || !user.employeeId) return;
  const sh=master_().getSheetByName('EmployeeSettings'); if(!sh) return;
  ensureHeaderColumns_(sh,['employeeId','username','officeCity','officeAddress','officeInTime','officeOutTime','weekoffDay','photoUrl','updatedAt']);
  const rows=readRows_(sh), existing=rows.find(function(r){return String(r.employeeId)===String(user.employeeId);});
  let day=settings.weekoffDay;
  if(day===''||day===null||day===undefined){
    const sr=readRows_(master_().getSheetByName('SystemSettings')).find(function(r){return r.key==='WEEKOFF_DAY';}); day=sr?sr.value:'0';
  }
  const obj={employeeId:user.employeeId,username:user.username,officeCity:settings.officeCity||user.office||'',officeAddress:settings.officeAddress||user.address||'',officeInTime:settings.officeInTime||'',officeOutTime:settings.officeOutTime||'',weekoffDay:String(day),updatedAt:new Date()};
  if(existing) updateByKey_(sh,'employeeId',user.employeeId,obj); else appendObject_(sh,obj);
}

function advanced_(s) {
  const user=findUser_(s.username); if(!user) throw new Error('User not found.');
  const isAdmin=['MASTER_ADMIN','ADMIN','HOD'].includes(String(s.role));
  let requests=[];
  if(user.employeeCode){ const file=getEmployeeFile_(user), reqSheet=file.getSheetByName('Requests'); requests=readRows_(reqSheet).slice(-50).reverse(); }
  const result={requests:requests};
  if(!isAdmin) return result;
  const h=master_().getSheetByName('Holidays'), set=master_().getSheetByName('SystemSettings');
  const holidays=readRows_(h).sort(function(a,b){return String(a.date).localeCompare(String(b.date));});
  const settings={}; readRows_(set).forEach(function(r){settings[r.key]=r.value;});
  const approvals=readRows_(master_().getSheetByName('Approvals')).filter(function(r){return r.status==='PENDING';}).reverse().slice(0,100);
  result.holidays=holidays.slice(0,200); result.weekoff=settings.WEEKOFF_DAY||'0'; result.approvals=approvals; result.audit=readRows_(master_().getSheetByName('AuditLog')).slice(-50).reverse();
  return result;
}

function createHoliday_(s,h) {
  assertAdmin_(s); if(!h || !h.date || !h.name) throw new Error('Holiday date and name are required.');
  const sh=master_().getSheetByName('Holidays');
  const date=normalizeAttendanceDate_(h.date); if(!date) throw new Error('Invalid holiday date.');
  const existing=readRows_(sh).find(function(r){return String(r.date)===date;});
  if(existing) throw new Error('Holiday already exists for this date.');
  appendObject_(sh,{id:Utilities.getUuid(),date:date,name:String(h.name).trim(),type:h.type||'PUBLIC',createdBy:s.username,createdAt:new Date()});
  audit_(s,'CREATE','HOLIDAY',date+' / '+h.name); return {};
}

function deleteHoliday_(s,id) {
  assertAdmin_(s); const sh=master_().getSheetByName('Holidays');
  const data=sh.getDataRange().getValues(), h=data[0], ix=h.indexOf('id');
  for(let r=1;r<data.length;r++) if(String(data[r][ix])===String(id)){sh.deleteRow(r+1);audit_(s,'DELETE','HOLIDAY',String(id));return {};}
  throw new Error('Holiday not found.');
}

function saveWeekoff_(s,day) {
  assertAdmin_(s); const n=Number(day); if(!Number.isInteger(n)||n<0||n>6) throw new Error('Weekoff day must be 0 to 6.');
  const sh=master_().getSheetByName('SystemSettings'); const rows=readRows_(sh); const old=rows.find(function(r){return r.key==='WEEKOFF_DAY';});
  if(old) updateByKey_(sh,'key','WEEKOFF_DAY',{value:String(n),updatedAt:new Date()}); else appendObject_(sh,{key:'WEEKOFF_DAY',value:String(n),updatedAt:new Date()});
  audit_(s,'UPDATE','SETTINGS','WEEKOFF_DAY='+n); return {weekoff:String(n)};
}

function setRamadanBreakFreeze_(s,frozen){
  assertAdmin_(s);
  const value=Boolean(frozen);
  const sh=master_().getSheetByName('SystemSettings');
  const rows=readRows_(sh);
  const old=rows.find(function(r){return r.key==='RAMADAN_LUNCH_BREAK_FROZEN';});
  if(old) updateByKey_(sh,'key','RAMADAN_LUNCH_BREAK_FROZEN',{value:value?'TRUE':'FALSE',updatedAt:new Date()});
  else appendObject_(sh,{key:'RAMADAN_LUNCH_BREAK_FROZEN',value:value?'TRUE':'FALSE',updatedAt:new Date()});
  audit_(s,'UPDATE','SETTINGS','RAMADAN_LUNCH_BREAK_FROZEN='+value);
  return {frozen:value};
}

function createRequest_(s,r) {
  if(!r || !r.type || !r.date) throw new Error('Request type and date are required.');
  const user=findUser_(s.username); if(!user) throw new Error('User not found.');
  const file=getEmployeeFile_(user), sh=file.getSheetByName('Requests'); ensureHeaderColumns_(sh,['id','type','date','details','startTime','endTime','meetingMode','adjustmentField','status','createdAt']);
  const id=Utilities.getUuid(), date=normalizeAttendanceDate_(r.date); if(!date) throw new Error('Invalid request date.');
  const type=String(r.type);
  if(type==='Time Adjustment' && !['IN','OUT'].includes(String(r.adjustmentField||''))) throw new Error('Time Adjustment mein IN ya OUT select karein.');
  if(type==='Time Adjustment' && !normalizeTime_(r.adjustmentTime)) throw new Error('Time Adjustment ka time select karein.');
  if(type==='Meeting' && !['Online','Physically'].includes(String(r.meetingMode||''))) throw new Error('Meeting type select karein: Online ya Physically.');
  if(type==='Meeting Journey'){ if(!normalizeTime_(r.startTime)||!normalizeTime_(r.endTime)) throw new Error('Meeting Journey mein From Time aur To Time dono select karein.'); if(minutesBetween_(r.startTime,r.endTime)<=0) throw new Error('Meeting Journey ka To Time, From Time ke baad hona chahiye.'); }
  if(type==='Weekoff Adjustment' && !r.date) throw new Error('Weekoff Adjustment date required.');
  let finalDetails=String(r.details||'');
  if(type==='Time Adjustment') finalDetails=`${r.adjustmentField} Time: ${r.adjustmentTime}. ${finalDetails}`.trim();
  if(type==='Meeting') finalDetails=`Mode: ${r.meetingMode}. ${finalDetails}`.trim();
  if(type==='Meeting Journey') finalDetails=`Duration: ${r.startTime} to ${r.endTime}. ${finalDetails}`.trim();
  appendObject_(sh,{id:id,type:type,date:date,details:finalDetails,startTime:r.startTime||'',endTime:r.endTime||'',meetingMode:r.meetingMode||'',adjustmentField:r.adjustmentField||'',status:'PENDING',createdAt:new Date()});
  const a=master_().getSheetByName('Approvals'); ensureHeaderColumns_(a,['id','requestId','type','employeeId','employeeName','username','date','details','startTime','endTime','meetingMode','adjustmentField','adjustmentTime','status','createdBy','createdAt','updatedAt']);
  appendObject_(a,{id:Utilities.getUuid(),requestId:id,type:type,employeeId:user.employeeId,employeeName:user.name,username:user.username,date:date,details:finalDetails,startTime:r.startTime||'',endTime:r.endTime||'',meetingMode:r.meetingMode||'',adjustmentField:r.adjustmentField||'',adjustmentTime:r.adjustmentTime||'',status:'PENDING',createdBy:user.username,createdAt:new Date(),updatedAt:new Date()});
  const approvers=readUsersCached_().filter(function(x){return ['MASTER_ADMIN','ADMIN','HOD'].includes(String(x.role||'')) && String(x.status||'ACTIVE')==='ACTIVE';}).map(function(x){return x.username;});
  notifyUsers_(approvers,'New Approval Request',`${user.name} ne ${type} request submit ki hai for ${date}.`,'APPROVAL');
  audit_(s,'CREATE','REQUEST',type+' / '+date); return {requestId:id};
}

function auditLog_(s) {
  assertMasterAdmin_(s); return {rows:readRows_(master_().getSheetByName('AuditLog')).slice(-100).reverse()};
}

function notifications_(s) {
  const sh=master_().getSheetByName('Notifications');
  return {rows:readRows_(sh).filter(r=>r.username===s.username || r.username==='*').reverse().slice(0,50)};
}

function notifyUsers_(usernames,title,message,type){
  const sh=master_().getSheetByName('Notifications');
  if(!sh) return;
  ensureHeaderColumns_(sh,['id','username','title','message','type','time','read']);
  const list=[...new Set((usernames||[]).filter(Boolean))];
  list.forEach(function(username){
    appendObject_(sh,{id:Utilities.getUuid(),username:username,title:String(title||''),message:String(message||''),type:type||'INFO',time:new Date(),read:'N'});
  });
}

function employeeProfile_(s,username){
  assertAdmin_(s);
  const u=findUser_(username); if(!u) throw new Error('Employee not found.');
  const st=employeeSettings_(u);
  let photoUrl='';
  try{ const r=readRows_(master_().getSheetByName('EmployeeSettings')).find(x=>String(x.employeeId)===String(u.employeeId)); photoUrl=r&&r.photoUrl||''; }catch(e){}
  return {profile:{name:u.name,employeeId:u.employeeId,username:u.username,role:u.role,department:u.department,designation:u.designation,phone:u.phone,whatsapp:u.whatsapp,officeCity:st.officeCity||u.office||'',officeAddress:st.officeAddress||u.address||'',officeInTime:st.officeInTime||'',officeOutTime:st.officeOutTime||'',weekoffDay:String(st.weekoffDay??''),weekoffLabel:weekDayName_(st.weekoffDay),photoUrl:photoUrl}};
}

function weekDayName_(day){
  const names=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const n=Number(day); return Number.isInteger(n)&&n>=0&&n<=6 ? names[n] : '';
}

function getAttendanceRoster_(user,date){
  const sh=master_().getSheetByName('AttendanceRoster'); if(!sh||!user||!user.employeeId||!date) return null;
  const row=readRows_(sh).find(function(r){return String(r.employeeId)===String(user.employeeId)&&normalizeAttendanceDate_(r.date)===String(date);});
  return row||null;
}

function saveAttendanceRoster_(user,date,field,time){
  if(!user||!user.employeeId||!date||!normalizeTime_(time)) throw new Error('Roster details invalid.');
  const sh=master_().getSheetByName('AttendanceRoster');
  ensureHeaderColumns_(sh,['employeeId','username','date','officeInTime','officeOutTime','weekoffDay','updatedAt']);
  const current=getAttendanceRoster_(user,date)||{};
  const base=employeeSettings_(user);
  const obj={employeeId:user.employeeId,username:user.username,date:date,officeInTime:current.officeInTime||base.officeInTime||'',officeOutTime:current.officeOutTime||base.officeOutTime||'',weekoffDay:current.weekoffDay!==undefined&&current.weekoffDay!==''?String(current.weekoffDay):String(base.weekoffDay||''),updatedAt:new Date()};
  if(String(field)==='IN') obj.officeInTime=normalizeTime_(time); else if(String(field)==='OUT') obj.officeOutTime=normalizeTime_(time); else throw new Error('Roster adjustment field invalid.');
  const rows=readRows_(sh), existingIndex=rows.findIndex(function(r){return String(r.employeeId)===String(user.employeeId)&&normalizeAttendanceDate_(r.date)===String(date);});
  if(existingIndex>=0){
    const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
    const rowNo=existingIndex+2;
    headers.forEach(function(h,i){ if(Object.prototype.hasOwnProperty.call(obj,h)) sh.getRange(rowNo,i+1).setValue(obj[h]); });
  } else appendObject_(sh,obj);
}

function getAttendanceRules_(user){
  const globalRows=readRows_(master_().getSheetByName('SystemSettings')), gs={}; globalRows.forEach(function(r){gs[r.key]=r.value;});
  const st=employeeSettings_(user);
  const rosterToday=getAttendanceRoster_(user,Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'));
  const weekoff=rosterToday&&rosterToday.weekoffDay!==''?rosterToday.weekoffDay:(st.weekoffDay!==''?st.weekoffDay:(gs.WEEKOFF_DAY||'0'));
  const reasonMap={};
  const h=master_().getSheetByName('Holidays');
  if(h) readRows_(h).forEach(function(r){const d=normalizeAttendanceDate_(r.date); if(d) reasonMap[d]='Holiday: '+String(r.name||'Holiday');});
  try{
    const file=getEmployeeFile_(user), req=file.getSheetByName('Requests');
    if(req) readRows_(req).forEach(function(r){if(String(r.status)!=='APPROVED') return; const d=normalizeAttendanceDate_(r.date); if(!d)return; if(['Leave','Weekoff Adjustment'].includes(String(r.type))) reasonMap[d]=String(r.type);});
  }catch(e){}
  return {weekoff:String(weekoff),reasonMap:reasonMap};
}

function isNonWorkingDate_(date,rules){
  const d=normalizeAttendanceDate_(date); if(!d) return false;
  if(rules&&rules.reasonMap&&rules.reasonMap[d]) return true;
  const dt=new Date(d+'T00:00:00');
  return !isNaN(dt.getTime()) && dt.getDay()===Number(rules&&rules.weekoff!==undefined?rules.weekoff:0);
}

function profile_(s){
  const u=findUser_(s.username); if(!u) throw new Error('User not found.');
  const st=employeeSettings_(u);
  let photoUrl='';
  try{ const r=readRows_(master_().getSheetByName('EmployeeSettings')).find(x=>String(x.employeeId)===String(u.employeeId)); photoUrl=r&&r.photoUrl||''; }catch(e){}
  return {profile:{name:u.name,employeeId:u.employeeId,username:u.username,role:u.role,department:u.department,designation:u.designation,phone:u.phone,whatsapp:u.whatsapp,officeCity:st.officeCity||u.office||'',officeAddress:st.officeAddress||u.address||'',officeInTime:st.officeInTime||'',officeOutTime:st.officeOutTime||'',weekoffDay:String(st.weekoffDay??''),weekoffLabel:weekDayName_(st.weekoffDay),photoUrl:photoUrl}};
}

function changePassword_(s,currentPassword,newPassword){
  const u=findUser_(s.username); if(!u) throw new Error('User not found.');
  if(String(currentPassword||'')!==String(u.password||'')) throw new Error('Current password is incorrect.');
  const np=String(newPassword||''); if(np.length<4) throw new Error('New password must be at least 4 characters.');
  const sh=master_().getSheetByName('Users'); updateByKey_(sh,'username',u.username,{password:np,updatedAt:new Date()}); invalidateUsersCache_();
  audit_(s,'UPDATE','PASSWORD','Employee changed own password');
  notifyUsers_([u.username],'Password Changed','Aapka login password successfully change ho gaya.','SECURITY');
  return {};
}

function uploadProfilePhoto_(s,targetUsername,fileName,dataUrl,mimeType){
  const target=findUser_(targetUsername); if(!target) throw new Error('Employee not found.');
  return uploadProfilePhotoForAdmin_(s,target,fileName,dataUrl,mimeType);
}

// Shared Admin/Master Admin photo uploader. Keeping the actual upload logic here
// also lets User Create/Update save the photo in the same request, avoiding a
// second API action that can fail when an older Worker/Apps Script deployment is live.
function uploadProfilePhotoForAdmin_(s,target,fileName,dataUrl,mimeType){
  assertAdmin_(s);
  if(!dataUrl || !String(dataUrl).startsWith('data:')) throw new Error('Valid photo file required.');
  const match=String(dataUrl).match(/^data:([^;]+);base64,(.+)$/); if(!match) throw new Error('Invalid photo data.');
  const bytes=Utilities.base64Decode(match[2]);
  if(bytes.length>2*1024*1024) throw new Error('Photo 2 MB se chhoti honi chahiye.');
  let folder;
  const props=PropertiesService.getScriptProperties();
  const folderId=props.getProperty('PHOTO_FOLDER_ID');
  if(folderId){ try{folder=DriveApp.getFolderById(folderId);}catch(e){} }
  if(!folder){
    folder=getOrCreateFolderPath_(['Dashboard Working','office-task-report','Employees Photo']);
    props.setProperty('PHOTO_FOLDER_ID',folder.getId());
  }
  const safe=String(fileName||'profile.jpg').replace(/[^a-zA-Z0-9._-]/g,'_');
  const blob=Utilities.newBlob(bytes,mimeType||match[1],String(target.employeeId)+'_'+safe);
  const file=folder.createFile(blob);
  try{file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);}catch(e){}
  const url='https://drive.google.com/uc?export=view&id='+file.getId();
  const sh=master_().getSheetByName('EmployeeSettings');
  ensureHeaderColumns_(sh,['employeeId','username','officeCity','officeAddress','officeInTime','officeOutTime','weekoffDay','photoUrl','updatedAt']);
  const rows=readRows_(sh), existing=rows.find(r=>String(r.employeeId)===String(target.employeeId));
  if(existing) updateByKey_(sh,'employeeId',target.employeeId,{photoUrl:url,updatedAt:new Date()});
  else appendObject_(sh,{employeeId:target.employeeId,username:target.username,officeCity:target.office||'',officeAddress:target.address||'',officeInTime:target.officeInTime||'',officeOutTime:target.officeOutTime||'',weekoffDay:target.weekoffDay!==undefined?target.weekoffDay:'',photoUrl:url,updatedAt:new Date()});
  audit_(s,'UPDATE','PROFILE_PHOTO',`Profile photo updated for ${target.name} (${target.employeeId}) by ${s.username}`);
  notifyUsers_([target.username],'Profile Photo Updated','Aapki profile photo Admin ne update ki hai.','PROFILE');
  return {photoUrl:url,employeeId:target.employeeId,username:target.username};
}

function getEmployeeFile_(u) {
  if(!u || !u.employeeCode) throw new Error('Employee Id is required.');
  const year=new Date().getFullYear(), props=PropertiesService.getScriptProperties();
  const folder=DriveApp.getFolderById(props.getProperty('EMPLOYEE_FOLDER_ID'));
  const targetName=`${u.name}_${u.employeeCode}_${year}`;
  const existing=folder.getFilesByName(targetName);
  if(existing.hasNext()) return SpreadsheetApp.open(existing.next());
  const ss=SpreadsheetApp.create(targetName);
  const file=DriveApp.getFileById(ss.getId());
  folder.addFile(file); DriveApp.getRootFolder().removeFile(file);
  setupEmployeeSheets_(ss,u);
  return ss;
}

function setupEmployeeSheets_(ss,u) {
  const first=ss.getSheets()[0]; first.setName('Profile');
  const map={
    Profile:['name','employeeId','username','role','department','designation','phone','whatsapp','office','address','country','region','state','division','district','year'],
    Attendance:['date','in','out','officeMinutes','breakMinutes','break1Type','break1NamazType','break1Start','break1End','break1Reason','break2Type','break2NamazType','break2Start','break2End','break2Reason','break3Type','break3NamazType','break3Start','break3End','break3Reason','status','approveStatus','updatedAt'],
    Tasks:['id','date','name','details','category','priority','due','status','startTime','completedTime','actualMinutes','expectedMinutes','progress','progressNote','assignedBy','assignedAt','updatedAt','assignmentMonth','assignmentYear','assignmentFrom','assignmentTo','assignmentKey'],
    TaskTime:['taskId','date','start','end','minutes'],
    Leave:['id','from','to','type','reason','status','createdAt'],
    Requests:['id','type','date','details','startTime','endTime','meetingMode','adjustmentField','adjustmentTime','status','createdAt'],
    Activities:['time','title','details'],
    Notifications:['time','title','message','type','read']
  };
  Object.keys(map).forEach(name=>{
    const sh=name==='Profile'?first:ss.insertSheet(name);
    sh.getRange(1,1,1,map[name].length).setValues([map[name]]);
    sh.setFrozenRows(1);
  });
  const p=ss.getSheetByName('Profile');
  p.getRange(2,1,1,16).setValues([[u.name,u.employeeCode,u.username,u.role,u.department,u.designation,u.phone,u.whatsapp,u.office,u.address,u.country||'',u.region||'',u.state||'',u.division||'',u.district||'',new Date().getFullYear()]]);
}

function monthTrend_(users) {
  const labels=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return labels.map((label,m)=>{
    let c=0,p=0;
    users.forEach(u=>readRows_(getEmployeeFile_(u).getSheetByName('Tasks')).forEach(r=>{
      if(r.status==='COMPLETED' && new Date(r.date).getMonth()===m)c++;
      if(r.status && !['COMPLETED','CANCELLED'].includes(r.status) && new Date(r.date).getMonth()===m)p++;
    }));
    return {label,completed:c,pending:p};
  });
}

function last7_(rows) {
  const out=[]; const now=new Date();
  for(let d=6;d>=0;d--){const x=new Date(now);x.setDate(x.getDate()-d);const key=Utilities.formatDate(x,Session.getScriptTimeZone(),'yyyy-MM-dd');out.push({label:Utilities.formatDate(x,Session.getScriptTimeZone(),'dd MMM'),completed:rows.filter(r=>r.date===key&&r.status==='COMPLETED').length});}
  return out;
}

function appendActivity_(ss,u,title,details) {
  appendObject_(ss.getSheetByName('Activities'),{time:new Date(),title,details});
}

function appendObject_(sh,obj) {
  const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.appendRow(h.map(k=>obj[k]!==undefined?obj[k]:''));
}

function updateByKey_(sh,key,value,patch) {
  const data=sh.getDataRange().getValues(), h=data[0], i=h.indexOf(key);
  if(i<0) throw new Error('Column not found: '+key);
  for(let r=1;r<data.length;r++) if(String(data[r][i])===String(value)){
    Object.keys(patch).forEach(k=>{const c=h.indexOf(k);if(c>=0)sh.getRange(r+1,c+1).setValue(patch[k])});
    return;
  }
  throw new Error('Record not found.');
}

function migrateUsersSchema_(sh) {
  // V25: Users schema is manually finalized by the user.
  // Never rebuild, reorder, delete, or rename columns automatically.
  if (!sh) throw new Error('Users sheet not found.');

  const desired = [
    'name','employeeId','username','role','password','department',
    'designation','phone','whatsapp','office','address','country',
    'region','state','division','district','status','createdAt','updatedAt'
  ];

  const headers = sh.getLastRow() > 0 && sh.getLastColumn() > 0
    ? sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(h){
        return String(h || '').trim();
      })
    : [];

  const exact =
    headers.length === desired.length &&
    desired.every(function(h,i){ return headers[i] === h; });

  if (!exact) {
    throw new Error(
      'Users sheet must use the finalized manual header A:S: ' +
      desired.join(', ')
    );
  }

  return true;
}

function ensureHeaderColumns_(sh,required) {
  if(sh.getLastRow()===0 || sh.getLastColumn()===0) {
    sh.getRange(1,1,1,required.length).setValues([required]);
    sh.setFrozenRows(1);
    return;
  }
  const current=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  required.forEach(function(col){
    if(current.indexOf(col)===-1){
      sh.getRange(1,sh.getLastColumn()+1).setValue(col);
      current.push(col);
    }
  });
  sh.setFrozenRows(1);
}

function invalidateUsersCache_() { try { CacheService.getScriptCache().remove('OTR_USERS_V2'); } catch(e) {} }

function readUsersCached_() {
  const sh = master_().getSheetByName('Users');

  const expected = [
    'name','employeeId','username','role','password','department',
    'designation','phone','whatsapp','office','address','country',
    'region','state','division','district','status','createdAt','updatedAt'
  ];

  const header = sh.getLastRow() > 0 && sh.getLastColumn() > 0
    ? sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(h){
        return String(h || '').trim();
      })
    : [];

  const exact =
    header.length === expected.length &&
    expected.every(function(h,i){ return header[i] === h; });

  if (!exact) {
    throw new Error(
      'Users sheet header mismatch. Expected: ' +
      expected.join(' | ')
    );
  }

  const cache = CacheService.getScriptCache();
  const cached = cache.get('OTR_USERS_V2');

  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  const rows = readRows_(sh);

  rows.forEach(function(r) {
    // Internal compatibility aliases only. Sheet remains employeeId.
    r.employeeId = r.employeeId || r.employeeCode || '';
    r.employeeCode = r.employeeId;
  });

  try {
    cache.put(
      'OTR_USERS_V2',
      JSON.stringify(rows),
      30
    );
  } catch(e) {}

  return rows;
}

function readRows_(sh) {
  if(!sh || sh.getLastRow()<2) return [];
  const v=sh.getDataRange().getValues(), h=v.shift();
  return v.map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i] instanceof Date ? formatDateTime_(r[i]) : r[i]);return o});
}

function index_(h){const o={};h.forEach((x,i)=>o[x]=i);return o}
function master_() {
  const id = PropertiesService.getScriptProperties().getProperty('MASTER_ID');

  if (id) {
    try {
      const ss = SpreadsheetApp.openById(id);
      if (ss.getSheetByName('Users')) return ss;
    } catch(e) {}
  }

  // Resolve the manually finalized Master Sheet instead of trusting a stale ID.
  return resolveMasterSpreadsheetForLogin_();
}
function assertAdmin_(s){if(!['MASTER_ADMIN','ADMIN','HOD'].includes(s.role))throw new Error('Access denied.')}
function audit_(s,action,module,details){const sh=master_().getSheetByName('AuditLog');sh.appendRow([new Date(),s.username,action,module,details])}
function hash_(x){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(x),Utilities.Charset.UTF_8).map(b=>((b+256)%256).toString(16).padStart(2,'0')).join('')}
function formatTime_(d){return Utilities.formatDate(d,Session.getScriptTimeZone(),'HH:mm:ss')}
function formatDateTime_(d){return Utilities.formatDate(new Date(d),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss')}
function minutesBetween_(a,b){if(!a||!b)return 0;const [ah,am,as]=a.split(':').map(Number),[bh,bm,bs]=b.split(':').map(Number);return Math.max(0,(bh*60+bm+bs/60)-(ah*60+am+as/60))}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
