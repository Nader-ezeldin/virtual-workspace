var SHEET_USERS = 'Users';
var SHEET_TASKS = 'Tasks';
var SHEET_BUDGETS = 'Budgets';
var SHEET_CHAT = 'Chat';
var SHEET_FILE_REQUESTS = 'FileRequests';
var SHEET_EVENTS = 'CalendarEvents';
var SHEET_OBJECTIVES = 'WeeklyObjectives';
var SHEET_PLANNER = 'PlannerItems';
var SHEET_OFFICES = 'Offices';
var SHEET_ASSETS = 'Assets';

// ترويسات موحدة تضمن التخزين في العمود والمكان الصحيح
var HEADERS = {
  Users:            ['Email', 'Name', 'Role', 'Status', 'AddedAt', 'LastSeen', 'SubOffice', 'PIN', 'SessionToken'],
  Tasks:            ['Id', 'Title', 'Description', 'Status', 'Priority', 'AssignedTo', 'DueDate', 'CreatedBy', 'CreatedAt', 'UpdatedAt'],
  Budgets:          ['Id', 'Category', 'Description', 'Amount', 'Type', 'Date', 'CreatedBy', 'CreatedAt'],
  Chat:             ['Id', 'Email', 'Name', 'Message', 'CreatedAt'],
  FileRequests:     ['Id', 'Serial', 'Title', 'Description', 'RequestedFrom', 'RequestedBy', 'Status', 'FileUrl', 'FileName', 'UploadedBy', 'CreatedAt', 'UpdatedAt'],
  CalendarEvents:   ['Id', 'Title', 'Date', 'Time', 'Type', 'GoogleEventId', 'CreatedBy', 'CreatedAt'],
  WeeklyObjectives: ['Id', 'Objective', 'CreatedBy', 'CreatedAt'],
  PlannerItems:     ['Id', 'Day', 'Text', 'CreatedBy', 'CreatedAt'],
  Offices:          ['SubOfficeName', 'Location', 'WifiSSID', 'WifiPass', 'Phone', 'Status'],
  Assets:           ['Id', 'Serial', 'Name', 'Center', 'AssignedTo', 'AssignedBy', 'Date']
};

var ROLE_ADMIN = 'Admin';
var ROLE_STAFF = 'Staff'; 
var STATUS_ACTIVE = 'Active';
var STATUS_PENDING = 'Pending';

function doGet(e) {
  ensureInitialized_();
  var tmpl = HtmlService.createTemplateFromFile('index');
  return tmpl.evaluate()
    .setTitle('Workspace Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupApp() {
  ensureInitialized_();
  return 'Setup complete. All sheets and schemas initialized with dynamic col mapping.';
}

function requireAuth_(token) {
  if (!token) throw new Error("Unauthorized: Missing session token.");
  var users = sheetToObjects_(SHEET_USERS);
  var user = users.filter(function(u) { return String(u.SessionToken) === String(token); })[0];
  if (!user) throw new Error("Unauthorized: Session is expired or invalid.");
  if (user.Status !== STATUS_ACTIVE) throw new Error("Unauthorized: Your account is pending admin approval.");
  return user;
}

function requireAdmin_(user) {
  if (user.Role !== ROLE_ADMIN) throw new Error('This action requires Admin permissions.');
}

function getColumnIndex_(sheetName, headerName) {
  var sheet = getSheet_(sheetName);
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return -1;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  headers = headers.map(function(h) { return String(h).trim().toLowerCase(); });
  return headers.indexOf(headerName.toLowerCase()) + 1;
}

function setSheetValue_(sheetName, row, headerName, value) {
  var sheet = getSheet_(sheetName);
  var colIdx = getColumnIndex_(sheetName, headerName);
  if (colIdx > 0) {
    sheet.getRange(row, colIdx).setValue(value);
  } else {
    var newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol).setValue(headerName)
         .setFontWeight('bold')
         .setBackground('#0f172a')
         .setFontColor('#ffffff');
    sheet.getRange(row, newCol).setValue(value);
    SpreadsheetApp.flush();
  }
}

function api_login(email, pin) {
  email = String(email || '').trim().toLowerCase();
  ensureInitialized_();
  var users = sheetToObjects_(SHEET_USERS);
  
  if (users.length === 0) {
    var sheet = getSheet_(SHEET_USERS);
    var newRowIdx = sheet.getLastRow() + 1;
    sheet.appendRow([email]);
    
    setSheetValue_(SHEET_USERS, newRowIdx, 'Email', email);
    setSheetValue_(SHEET_USERS, newRowIdx, 'Name', email.split('@')[0]);
    setSheetValue_(SHEET_USERS, newRowIdx, 'Role', ROLE_ADMIN);
    setSheetValue_(SHEET_USERS, newRowIdx, 'Status', STATUS_ACTIVE);
    setSheetValue_(SHEET_USERS, newRowIdx, 'AddedAt', nowIso_());
    setSheetValue_(SHEET_USERS, newRowIdx, 'LastSeen', nowIso_());
    setSheetValue_(SHEET_USERS, newRowIdx, 'SubOffice', 'Main Office');
    setSheetValue_(SHEET_USERS, newRowIdx, 'PIN', String(pin));
    setSheetValue_(SHEET_USERS, newRowIdx, 'SessionToken', '');
    
    SpreadsheetApp.flush();
    users = sheetToObjects_(SHEET_USERS);
  }
  
  var user = users.filter(function(u) { return String(u.Email).toLowerCase() === email; })[0];
  if (!user) throw new Error("Invalid credentials.");
  
  var sheetPin = String(user.PIN || '').trim();
  if (sheetPin.indexOf('.') !== -1) {
    sheetPin = sheetPin.split('.')[0];
  }
  
  if (sheetPin === '') {
    setSheetValue_(SHEET_USERS, user._row, 'PIN', String(pin));
    SpreadsheetApp.flush();
    sheetPin = String(pin).trim();
  }
  
  if (sheetPin !== String(pin).trim()) throw new Error("Invalid credentials.");
  if (user.Status !== STATUS_ACTIVE) throw new Error("Your account is pending admin approval.");
  
  var token = newId_();
  setSheetValue_(SHEET_USERS, user._row, 'SessionToken', token);
  setSheetValue_(SHEET_USERS, user._row, 'LastSeen', nowIso_());
  SpreadsheetApp.flush();
  
  return {
    token: token,
    user: { email: user.Email, name: user.Name, role: user.Role, subOffice: user.SubOffice }
  };
}

function api_logout(token) {
  var users = sheetToObjects_(SHEET_USERS);
  var user = users.filter(function(u) { return String(u.SessionToken) === String(token); })[0];
  if (user) {
    setSheetValue_(SHEET_USERS, user._row, 'SessionToken', "");
    SpreadsheetApp.flush();
  }
  return true;
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function standardizeHeaders_(sheetName) {
  var sheet = getSheet_(sheetName);
  if (!sheet) return;
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  
  var range = sheet.getRange(1, 1, 1, lastCol);
  var headers = range.getValues()[0].map(function(h) { return String(h).trim(); });
  var canonical = HEADERS[sheetName];
  
  var changed = false;
  var updatedHeaders = headers.map(function(h) {
    for (var i = 0; i < canonical.length; i++) {
      if (canonical[i].toLowerCase() === h.toLowerCase()) {
        if (canonical[i] !== h) {
          changed = true;
          return canonical[i];
        }
        return h;
      }
    }
    return h;
  });
  
  if (changed) {
    range.setValues([updatedHeaders]);
    SpreadsheetApp.flush();
  }
}

function ensureInitialized_() {
  var ss = ss_();
  Object.keys(HEADERS).forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(HEADERS[name]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    } else {
      standardizeHeaders_(name);
      
      var lastCol = sheet.getLastColumn();
      var existingHeaders = [];
      if (lastCol > 0) {
        existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      }
      
      existingHeaders = existingHeaders.map(function(h) { return String(h).trim().toLowerCase(); });
      
      var missingHeaders = [];
      HEADERS[name].forEach(function (h) {
        if (existingHeaders.indexOf(h.toLowerCase()) === -1) {
          missingHeaders.push(h);
        }
      });
      
      if (missingHeaders.length > 0) {
        var startCol = existingHeaders.length + 1;
        sheet.getRange(1, startCol, 1, missingHeaders.length)
             .setValues([missingHeaders])
             .setFontWeight('bold')
             .setBackground('#0f172a')
             .setFontColor('#ffffff');
      }
    }
  });
  var def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 4) {
    ss.deleteSheet(def);
  }
}

function getSheet_(name) { return ss_().getSheetByName(name); }

function sheetToObjects_(name) {
  var sheet = getSheet_(name);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function(h) { return String(h).trim(); });
  var rows = values.slice(1);
  return rows
    .map(function (row, i) {
      var obj = { _row: i + 2 };
      headers.forEach(function (h, idx) { obj[h] = row[idx]; });
      return obj;
    })
    .filter(function (obj) {
      return headers.some(function (h) { return obj[h] !== '' && obj[h] !== null && obj[h] !== undefined; });
    });
}

function backfillMissingIds_(name) {
  var idColIdx = getColumnIndex_(name, 'Id') - 1;
  if (idColIdx < 0) return;
  var sheet = getSheet_(name);
  var last = sheet.getLastRow();
  if (last < 2) return;
  var range = sheet.getRange(2, 1, last - 1, sheet.getLastColumn());
  var values = range.getValues();
  var changed = false;
  values.forEach(function (row) {
    var hasContent = row.some(function (v) { return v !== '' && v !== null && v !== undefined; });
    if (hasContent && (row[idColIdx] === '' || row[idColIdx] === null || row[idColIdx] === undefined)) {
      row[idColIdx] = newId_();
      changed = true;
    }
  });
  if (changed) range.setValues(values);
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function newId_() { return Utilities.getUuid(); }
function nowIso_() { return new Date().toISOString(); }

function sanitizeForClient_(value) {
  if (value === undefined || value === null) return null;
  if (value instanceof Date) return isNaN(value.getTime()) ? '' : value.toISOString();
  if (Array.isArray(value)) return value.map(sanitizeForClient_);
  if (typeof value === 'object') {
    var out = {};
    Object.keys(value).forEach(function (k) {
      if (k === '_row') return;
      out[k] = sanitizeForClient_(value[k]);
    });
    return out;
  }
  return value;
}

function api_getUsers(token) {
  var me = requireAuth_(token);
  requireAdmin_(me);
  return sanitizeForClient_(sheetToObjects_(SHEET_USERS).map(function (u) {
    return { email: u.Email, name: u.Name, role: u.Role, status: u.Status, addedAt: u.AddedAt, lastSeen: u.LastSeen, subOffice: u.SubOffice };
  }));
}

function api_addUser(token, email, name, role, subOffice, pin) {
  var me = requireAuth_(token);
  requireAdmin_(me);
  email = String(email).trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Enter a valid email address.');
  
  var users = sheetToObjects_(SHEET_USERS);
  var match = users.filter(function(u) { return String(u.Email).toLowerCase() === email; })[0];
  if (match) throw new Error('That user already exists.');
  
  var userPin = String(pin || '1234');
  return withLock_(function () {
    var sheet = getSheet_(SHEET_USERS);
    var newRowIdx = sheet.getLastRow() + 1;
    sheet.appendRow([email]);
    
    setSheetValue_(SHEET_USERS, newRowIdx, 'Email', email);
    setSheetValue_(SHEET_USERS, newRowIdx, 'Name', name || email.split('@')[0]);
    setSheetValue_(SHEET_USERS, newRowIdx, 'Role', role || ROLE_STAFF);
    setSheetValue_(SHEET_USERS, newRowIdx, 'Status', STATUS_ACTIVE);
    setSheetValue_(SHEET_USERS, newRowIdx, 'AddedAt', nowIso_());
    setSheetValue_(SHEET_USERS, newRowIdx, 'LastSeen', '');
    setSheetValue_(SHEET_USERS, newRowIdx, 'SubOffice', subOffice || '');
    setSheetValue_(SHEET_USERS, newRowIdx, 'PIN', userPin);
    setSheetValue_(SHEET_USERS, newRowIdx, 'SessionToken', '');
    
    SpreadsheetApp.flush();
    return api_getUsers(token);
  });
}

function api_updateUserStatus(token, email, status) {
  var me = requireAuth_(token);
  requireAdmin_(me);
  setUserField_(email, 'Status', status === STATUS_ACTIVE ? STATUS_ACTIVE : STATUS_PENDING);
  return api_getUsers(token);
}

function api_updateUserRole(token, email, role) {
  var me = requireAuth_(token);
  requireAdmin_(me);
  if (String(email).toLowerCase() === me.Email.toLowerCase() && role !== ROLE_ADMIN) {
    throw new Error("You can't remove your own Admin role.");
  }
  setUserField_(email, 'Role', role === ROLE_ADMIN ? ROLE_ADMIN : ROLE_STAFF);
  return api_getUsers(token);
}

function api_removeUser(token, email) {
  var me = requireAuth_(token);
  requireAdmin_(me);
  if (String(email).toLowerCase() === me.Email.toLowerCase()) throw new Error("You can't remove yourself.");
  withLock_(function () {
    var sheet = getSheet_(SHEET_USERS);
    var users = sheetToObjects_(SHEET_USERS);
    for (var i = users.length - 1; i >= 0; i--) {
      if (String(users[i].Email).toLowerCase() === String(email).toLowerCase()) {
        sheet.deleteRow(users[i]._row);
      }
    }
  });
  return api_getUsers(token);
}

function setUserField_(email, field, value) {
  withLock_(function () {
    var users = sheetToObjects_(SHEET_USERS);
    for (var i = 0; i < users.length; i++) {
      if (String(users[i].Email).toLowerCase() === String(email).toLowerCase()) {
        setSheetValue_(SHEET_USERS, users[i]._row, field, value);
        return;
      }
    }
    throw new Error('User not found.');
  });
}

function api_getOfficeDetails(token, userSubOffice) {
  var me = requireAuth_(token);
  if (me.Role !== ROLE_STAFF) {
    throw new Error('Access Denied: Only Office Staff are authorized to view office records.');
  }
  
  var offices = sheetToObjects_(SHEET_OFFICES);
  var myOffice = null;
  
  for (var i = 0; i < offices.length; i++) {
    if (offices[i].SubOfficeName === userSubOffice) {
      myOffice = offices[i];
      break;
    }
  }
  
  if (!myOffice) return null;
  
  return {
    subOfficeName: myOffice.SubOfficeName,
    location: myOffice.Location,
    wifiSSID: myOffice.WifiSSID,
    wifiPass: myOffice.WifiPass,
    phone: myOffice.Phone,
    status: myOffice.Status
  };
}

function api_getInitialData(token) {
  var me = requireAuth_(token);
  var officeDetails = null;
  try {
    if (me.Role === ROLE_STAFF && me.SubOffice) {
      officeDetails = api_getOfficeDetails(token, me.SubOffice);
    }
  } catch (err) {}

  return sanitizeForClient_({
    me: { email: me.Email, name: me.Name, role: me.Role, subOffice: me.SubOffice },
    tasks: api_getTasks(token),
    budgets: api_getBudgets(token),
    chat: api_getChatMessages(token),
    fileRequests: api_getFileRequests(token),
    users: me.Role === ROLE_ADMIN ? api_getUsers(token) : [],
    metrics: computeMetrics_(token),
    calendarEvents: api_getEvents(token),
    weeklyObjectives: api_getWeeklyObjectives(token),
    plannerDays: api_getPlannerDays(token),
    officeDetails: officeDetails,
    assets: api_getAssets(token)
  });
}

function computeMetrics_(token) {
  var tasks = sheetToObjects_(SHEET_TASKS);
  var budgets = sheetToObjects_(SHEET_BUDGETS);
  var now = new Date();

  var byStatus = { 'To Do': 0, 'In Progress': 0, 'Done': 0 };
  var overdue = 0;
  tasks.forEach(function (t) {
    var status = t.Status || 'To Do';
    byStatus[status] = (byStatus[status] || 0) + 1;
    if (status !== 'Done' && t.DueDate) {
      var dDate = new Date(t.DueDate);
      if (!isNaN(dDate.getTime()) && dDate < now) overdue++;
    }
  });

  var income = 0, expense = 0;
  var byCategory = {};
  budgets.forEach(function (b) {
    var amt = Number(b.Amount) || 0;
    var type = b.Type || 'Expense';
    if (type === 'Income') {
      income += amt;
    } else {
      expense += amt;
      var cat = b.Category || 'Uncategorized';
      byCategory[cat] = (byCategory[cat] || 0) + amt;
    }
  });

  return {
    totalTasks: tasks.length,
    tasksByStatus: byStatus,
    overdueTasks: overdue,
    income: income,
    expense: expense,
    balance: income - expense,
    expenseByCategory: byCategory
  };
}

function api_getMetrics(token) {
  requireAuth_(token);
  return sanitizeForClient_(computeMetrics_(token));
}

function api_getTasks(token) {
  requireAuth_(token);
  backfillMissingIds_(SHEET_TASKS);
  return sanitizeForClient_(sheetToObjects_(SHEET_TASKS).map(function (t) {
    return { id: t.Id, title: t.Title, description: t.Description, status: t.Status, priority: t.Priority,
      assignedTo: t.AssignedTo, dueDate: t.DueDate, createdBy: t.CreatedBy, createdAt: t.CreatedAt, updatedAt: t.UpdatedAt };
  }));
}

function api_addTask(token, task) {
  var me = requireAuth_(token);
  if (!task || !task.title) throw new Error('A task needs a title.');
  return withLock_(function () {
    var id = newId_();
    var ts = nowIso_();
    getSheet_(SHEET_TASKS).appendRow([
      id, task.title, task.description || '', task.status || 'To Do', task.priority || 'Medium',
      task.assignedTo || me.Email, task.dueDate || '', me.Email, ts, ts
    ]);
    SpreadsheetApp.flush();
    return api_getTasks(token);
  });
}

function api_updateTask(token, id, updates) {
  requireAuth_(token);
  return withLock_(function () {
    var rows = sheetToObjects_(SHEET_TASKS);
    var row = rows.filter(function (r) { return r.Id === id; })[0];
    if (!row) throw new Error('Task not found.');
    var fieldMap = { title: 'Title', description: 'Description', status: 'Status', priority: 'Priority', assignedTo: 'AssignedTo', dueDate: 'DueDate' };
    Object.keys(updates || {}).forEach(function (k) {
      if (fieldMap[k]) {
        setSheetValue_(SHEET_TASKS, row._row, fieldMap[k], updates[k]);
      }
    });
    setSheetValue_(SHEET_TASKS, row._row, 'UpdatedAt', nowIso_());
    SpreadsheetApp.flush();
    return api_getTasks(token);
  });
}

function api_deleteTask(token, id) {
  requireAuth_(token);
  return withLock_(function () {
    var sheet = getSheet_(SHEET_TASKS);
    var rows = sheetToObjects_(SHEET_TASKS);
    var row = rows.filter(function (r) { return r.Id === id; })[0];
    if (row) sheet.deleteRow(row._row);
    SpreadsheetApp.flush();
    return api_getTasks(token);
  });
}

function api_getBudgets(token) {
  requireAuth_(token);
  backfillMissingIds_(SHEET_BUDGETS);
  return sanitizeForClient_(sheetToObjects_(SHEET_BUDGETS).map(function (b) {
    return { id: b.Id, category: b.Category, description: b.Description, amount: Number(b.Amount) || 0,
      type: b.Type, date: b.Date, createdBy: b.CreatedBy, createdAt: b.CreatedAt };
  }));
}

function api_addBudget(token, entry) {
  var me = requireAuth_(token);
  if (!entry || !entry.category || entry.amount == null) throw new Error('A budget entry needs a category and amount.');
  return withLock_(function () {
    var id = newId_();
    getSheet_(SHEET_BUDGETS).appendRow([
      id, entry.category, entry.description || '', Number(entry.amount) || 0,
      entry.type === 'Income' ? 'Income' : 'Expense', entry.date || nowIso_().slice(0, 10), me.Email, nowIso_()
    ]);
    SpreadsheetApp.flush();
    return api_getBudgets(token);
  });
}

function api_updateBudget(token, id, updates) {
  requireAuth_(token);
  return withLock_(function () {
    var rows = sheetToObjects_(SHEET_BUDGETS);
    var row = rows.filter(function (r) { return r.Id === id; })[0];
    if (!row) throw new Error('Budget entry not found.');
    var fieldMap = { category: 'Category', description: 'Description', amount: 'Amount', type: 'Type', date: 'Date' };
    Object.keys(updates || {}).forEach(function (k) {
      if (fieldMap[k]) {
        var val = updates[k];
        if (k === 'amount') val = Number(val);
        setSheetValue_(SHEET_BUDGETS, row._row, fieldMap[k], val);
      }
    });
    SpreadsheetApp.flush();
    return api_getBudgets(token);
  });
}

function api_deleteBudget(token, id) {
  requireAuth_(token);
  return withLock_(function () {
    var sheet = getSheet_(SHEET_BUDGETS);
    var rows = sheetToObjects_(SHEET_BUDGETS);
    var row = rows.filter(function (r) { return r.Id === id; })[0];
    if (row) sheet.deleteRow(row._row);
    SpreadsheetApp.flush();
    return api_getBudgets(token);
  });
}

function api_getChatMessages(token) {
  requireAuth_(token);
  var msgs = sheetToObjects_(SHEET_CHAT).map(function (c) {
    return { id: c.Id, email: c.Email, name: c.Name, message: c.Message, createdAt: c.CreatedAt };
  });
  return sanitizeForClient_(msgs.slice(Math.max(0, msgs.length - 200)));
}

function api_postChatMessage(token, message) {
  var me = requireAuth_(token);
  message = String(message || '').trim();
  if (!message) throw new Error('Message is empty.');
  if (message.length > 2000) message = message.slice(0, 2000);
  withLock_(function () {
    getSheet_(SHEET_CHAT).appendRow([newId_(), me.Email, me.Name, message, nowIso_()]);
    SpreadsheetApp.flush();
  });
  return api_getChatMessages(token);
}

function api_getFileRequests(token) {
  var me = requireAuth_(token);
  backfillMissingIds_(SHEET_FILE_REQUESTS);
  return sanitizeForClient_(sheetToObjects_(SHEET_FILE_REQUESTS).map(function (r) {
    return {
      id: r.Id, 
      serial: r.Serial || 'REQ-000', 
      title: r.Title, 
      description: r.Description, 
      requestedFrom: String(r.RequestedFrom || '').trim().toLowerCase(), // تأكيد حالة الأحرف الصغيرة
      requestedBy: String(r.RequestedBy || '').trim().toLowerCase(), 
      status: r.Status, 
      fileUrl: r.FileUrl, 
      fileName: r.FileName,
      uploadedBy: String(r.UploadedBy || '').trim().toLowerCase(),
      createdAt: r.CreatedAt, 
      updatedAt: r.UpdatedAt
    };
  }));
}

function api_addFileRequest(token, request) {
  var me = requireAuth_(token);
  if (!request || !request.title || !request.requestedFrom) {
    throw new Error('Title and Requested From fields are required.');
  }
  return withLock_(function () {
    var id = newId_();
    var ts = nowIso_();
    
    var startSerial = 1001;
    var existing = sheetToObjects_(SHEET_FILE_REQUESTS);
    existing.forEach(function (r) {
      var num = parseInt(String(r.Serial).replace(/^\D+/g, ''), 10);
      if (!isNaN(num) && num >= startSerial) {
        startSerial = num + 1;
      }
    });
    var serial = 'REQ-' + startSerial;

    // متوافق تماماً مع ترويسة FileRequests ذات الـ 12 عموداً
    getSheet_(SHEET_FILE_REQUESTS).appendRow([
      id, 
      serial, 
      request.title, 
      request.description || '', 
      request.requestedFrom.trim().toLowerCase(),
      me.Email.trim().toLowerCase(), 
      'Pending', 
      '', 
      '', 
      '', 
      ts, 
      ts
    ]);
    SpreadsheetApp.flush();
    return api_getFileRequests(token);
  });
}

function api_shareFileDirectly(token, title, description, fileUrl, fileName) {
  var me = requireAuth_(token);
  if (!fileUrl) throw new Error('A file URL link is required to share a file.');
  return withLock_(function () {
    var id = newId_();
    var ts = nowIso_();
    
    var startSerial = 1001;
    var existing = sheetToObjects_(SHEET_FILE_REQUESTS);
    existing.forEach(function (r) {
      var num = parseInt(String(r.Serial).replace(/^\D+/g, ''), 10);
      if (!isNaN(num) && num >= startSerial) {
        startSerial = num + 1;
      }
    });
    var serial = 'SHR-' + startSerial;

    // مشاركة فورية - متوافق مع هيكل الـ 12 عموداً
    getSheet_(SHEET_FILE_REQUESTS).appendRow([
      id, 
      serial, 
      title, 
      description || '', 
      'everyone',
      me.Email.trim().toLowerCase(), 
      'Fulfilled', 
      fileUrl, 
      fileName || 'Shared File', 
      me.Email.trim().toLowerCase(), 
      ts, 
      ts
    ]);
    SpreadsheetApp.flush();
    return api_getFileRequests(token);
  });
}

function api_fulfillFileRequest(token, id, fileUrl, fileName) {
  var me = requireAuth_(token);
  if (!fileUrl) throw new Error('A file URL/link is required to fulfill this request.');
  return withLock_(function () {
    var rows = sheetToObjects_(SHEET_FILE_REQUESTS);
    var row = rows.filter(function (r) { return r.Id === id; })[0];
    if (!row) throw new Error('Request not found.');
    
    setSheetValue_(SHEET_FILE_REQUESTS, row._row, 'FileUrl', fileUrl);
    setSheetValue_(SHEET_FILE_REQUESTS, row._row, 'FileName', fileName || 'Uploaded File');
    setSheetValue_(SHEET_FILE_REQUESTS, row._row, 'UploadedBy', me.Email.trim().toLowerCase());
    setSheetValue_(SHEET_FILE_REQUESTS, row._row, 'Status', 'Fulfilled');
    setSheetValue_(SHEET_FILE_REQUESTS, row._row, 'UpdatedAt', nowIso_());
    SpreadsheetApp.flush();
    return api_getFileRequests(token);
  });
}

function api_getMyDriveFiles(token) {
  requireAuth_(token);
  var files = [];
  try {
    var iterator = DriveApp.getFiles();
    var count = 0;
    while (iterator.hasNext() && count < 25) {
      var file = iterator.next();
      if (!file.isTrashed()) { 
        files.push({
          id: file.getId(),
          name: file.getName(),
          url: file.getUrl(),
          mimeType: file.getMimeType()
        });
        count++;
      }
    }
  } catch (err) {
    throw new Error('Unable to access Google Drive files: ' + err.message);
  }
  return sanitizeForClient_(files);
}

function api_getEvents(token) {
  var me = requireAuth_(token);
  backfillMissingIds_(SHEET_EVENTS);
  
  return withLock_(function() {
    var sheet = getSheet_(SHEET_EVENTS);
    var existingEvents = sheetToObjects_(SHEET_EVENTS);
    var googleEventIds = existingEvents.map(function(e) { return e.GoogleEventId; }).filter(Boolean);
    
    try {
      var calendar = CalendarApp.getDefaultCalendar();
      if (calendar) {
        var now = new Date();
        var startDate = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000); 
        var endDate = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
        var gEvents = calendar.getEvents(startDate, endDate);
        
        gEvents.forEach(function(ge) {
          var geId = ge.getId();
          if (googleEventIds.indexOf(geId) === -1) {
            var id = newId_();
            var title = ge.getTitle() || "Unnamed Event";
            var start = ge.getStartTime();
            var dateStr = Utilities.formatDate(start, Session.getScriptTimeZone() || 'GMT', 'yyyy-MM-dd');
            var timeStr = Utilities.formatDate(start, Session.getScriptTimeZone() || 'GMT', 'HH:mm');
            
            sheet.appendRow([
              id, title, dateStr, timeStr, 'Meeting', geId, me.Email, nowIso_()
            ]);
          }
        });
        SpreadsheetApp.flush();
        existingEvents = sheetToObjects_(SHEET_EVENTS);
      }
    } catch(err) {
      console.warn('Google Calendar bi-sync failed: ' + err.message);
    }
    
    return sanitizeForClient_(existingEvents.map(function (e) {
      return { id: e.Id, title: e.Title, date: e.Date, time: e.Time, type: e.Type, googleEventId: e.GoogleEventId, createdBy: e.CreatedBy, createdAt: e.CreatedAt };
    }));
  });
}

function api_addEvent(token, ev) {
  var me = requireAuth_(token);
  if (!ev || !ev.title || !ev.date) throw new Error('Title and Date are required.');
  return withLock_(function () {
    var id = newId_();
    var googleEventId = '';
    
    try {
      var calendar = CalendarApp.getDefaultCalendar();
      if (calendar) {
        var startDateTime = new Date(ev.date + 'T' + (ev.time || '12:00') + ':00');
        var endDateTime = new Date(startDateTime.getTime() + 60 * 60 * 1000);
        var gEvent = calendar.createEvent(ev.title, startDateTime, endDateTime, {
          description: 'Synchronized dynamically via Workspace Dashboard.'
        });
        if (gEvent) {
          googleEventId = gEvent.getId();
        }
      }
    } catch (calendarErr) {
      console.warn('Google Calendar creation sync failed: ' + calendarErr.message);
    }

    getSheet_(SHEET_EVENTS).appendRow([
      id, ev.title, ev.date, ev.time || '12:00', ev.type || 'Meeting', googleEventId, me.Email, nowIso_()
    ]);
    SpreadsheetApp.flush();
    return api_getEvents(token);
  });
}

function api_deleteEvent(token, id) {
  requireAuth_(token);
  return withLock_(function () {
    var sheet = getSheet_(SHEET_EVENTS);
    var rows = sheetToObjects_(SHEET_EVENTS);
    var row = rows.filter(function (r) { return r.Id === id; })[0];
    if (row) {
      if (row.GoogleEventId) {
        try {
          var calendar = CalendarApp.getDefaultCalendar();
          var gEvent = calendar.getEventById(row.GoogleEventId);
          if (gEvent) {
            gEvent.deleteEvent();
          }
        } catch (calendarErr) {
          console.warn('Could not locate or delete event on Google Calendar: ' + calendarErr.message);
        }
      }
      sheet.deleteRow(row._row);
    }
    SpreadsheetApp.flush();
    return api_getEvents(token);
  });
}

function api_getWeeklyObjectives(token) {
  requireAuth_(token);
  backfillMissingIds_(SHEET_OBJECTIVES);
  return sanitizeForClient_(sheetToObjects_(SHEET_OBJECTIVES).map(function (o) {
    return o.Objective;
  }));
}

function api_addWeeklyObjective(token, obj) {
  var me = requireAuth_(token);
  if (!obj || !obj.trim()) throw new Error('Objective title is required.');
  return withLock_(function () {
    var id = newId_();
    getSheet_(SHEET_OBJECTIVES).appendRow([id, obj.trim(), me.Email, nowIso_()]);
    SpreadsheetApp.flush();
    return api_getWeeklyObjectives(token);
  });
}

function api_removeWeeklyObjective(token, idx) {
  requireAuth_(token);
  return withLock_(function () {
    var sheet = getSheet_(SHEET_OBJECTIVES);
    var rows = sheetToObjects_(SHEET_OBJECTIVES);
    if (rows[idx]) sheet.deleteRow(rows[idx]._row);
    SpreadsheetApp.flush();
    return api_getWeeklyObjectives(token);
  });
}

function api_getPlannerDays(token) {
  requireAuth_(token);
  backfillMissingIds_(SHEET_PLANNER);
  var items = sheetToObjects_(SHEET_PLANNER);
  var days = { Mon: [], Tue: [], Wed: [], Thu: [], Fri: [], Sat: [], Sun: [] };
  items.forEach(function (item) {
    if (days[item.Day]) {
      days[item.Day].push({ id: item.Id, text: item.Text });
    }
  });
  return sanitizeForClient_(days);
}

function api_addPlannerItem(token, day, text) {
  var me = requireAuth_(token);
  if (!day || !text || !text.trim()) throw new Error('Day and plan metrics are required.');
  return withLock_(function () {
    var id = newId_();
    getSheet_(SHEET_PLANNER).appendRow([id, day, text.trim(), me.Email, nowIso_()]);
    SpreadsheetApp.flush();
    return api_getPlannerDays(token);
  });
}

function api_removePlannerItem(token, day, id) {
  requireAuth_(token);
  return withLock_(function () {
    var sheet = getSheet_(SHEET_PLANNER);
    var rows = sheetToObjects_(SHEET_PLANNER);
    var row = rows.filter(function (r) { return r.Id === id; })[0];
    if (row) sheet.deleteRow(row._row);
    SpreadsheetApp.flush();
    return api_getPlannerDays(token);
  });
}

function api_getAssets(token) {
  requireAuth_(token);
  backfillMissingIds_(SHEET_ASSETS);
  return sanitizeForClient_(sheetToObjects_(SHEET_ASSETS).map(function (a) {
    return { 
      id: a.Id, 
      serial: a.Serial || 'AST-000', 
      name: a.Name, 
      center: a.Center, 
      assignedTo: a.AssignedTo, 
      assignedBy: a.AssignedBy, 
      date: a.Date 
    };
  }));
}

function api_addAsset(token, asset) {
  var me = requireAuth_(token);
  if (!asset || !asset.name) throw new Error('An asset needs a name.');
  return withLock_(function () {
    var id = newId_();
    var startSerial = startSerialNum_(SHEET_ASSETS);
    var serial = 'AST-' + startSerial;
    var dateVal = asset.date || nowIso_().slice(0, 10);
    var assignedBy = asset.assignedBy || me.Email;

    getSheet_(SHEET_ASSETS).appendRow([
      id, serial, asset.name, asset.center || 'Main HQ', asset.assignedTo || '', assignedBy, dateVal
    ]);
    SpreadsheetApp.flush();
    return api_getAssets(token);
  });
}

function api_updateAsset(token, id, updates) {
  requireAuth_(token);
  return withLock_(function () {
    var rows = sheetToObjects_(SHEET_ASSETS);
    var row = rows.filter(function (r) { return r.Id === id; })[0];
    if (!row) throw new Error('Asset not found.');
    var fieldMap = {
      name: 'Name',
      center: 'Center',
      assignedTo: 'AssignedTo',
      assignedBy: 'AssignedBy',
      date: 'Date'
    };
    Object.keys(updates || {}).forEach(function (k) {
      if (fieldMap[k]) {
        setSheetValue_(SHEET_ASSETS, row._row, fieldMap[k], updates[k]);
      }
    });
    SpreadsheetApp.flush();
    return api_getAssets(token);
  });
}

function startSerialNum_(sheetName) {
  var startSerial = 1001;
  var existing = sheetToObjects_(sheetName);
  existing.forEach(function (x) {
    var num = parseInt(String(x.Serial).replace(/^\D+/g, ''), 10);
    if (!isNaN(num) && num >= startSerial) {
      startSerial = num + 1;
    }
  });
  return startSerial;
}

function api_importAssets(token, assetsList) {
  var me = requireAuth_(token);
  requireAdmin_(me);
  if (!Array.isArray(assetsList)) throw new Error('Invalid assets dataset.');
  
  return withLock_(function() {
    var sheet = getSheet_(SHEET_ASSETS);
    var startSerial = startSerialNum_(SHEET_ASSETS);
    var existing = sheetToObjects_(SHEET_ASSETS);
    
    assetsList.forEach(function(item) {
      var id = item.id || newId_();
      var serial = item.serial || ('AST-' + startSerial++);
      var dateVal = item.date || nowIso_().slice(0, 10);
      var assignedBy = item.assignedBy || me.Email;

      var existingRow = existing.filter(function(r) { return String(r.Id) === String(id); })[0];
      if (existingRow) {
        setSheetValue_(SHEET_ASSETS, existingRow._row, 'Serial', serial);
        setSheetValue_(SHEET_ASSETS, existingRow._row, 'Name', item.name);
        setSheetValue_(SHEET_ASSETS, existingRow._row, 'Center', item.center || 'Main HQ');
        setSheetValue_(SHEET_ASSETS, existingRow._row, 'AssignedTo', item.assignedTo || '');
        setSheetValue_(SHEET_ASSETS, existingRow._row, 'AssignedBy', assignedBy);
        setSheetValue_(SHEET_ASSETS, existingRow._row, 'Date', dateVal);
      } else {
        sheet.appendRow([
          id, serial, item.name, item.center || 'Main HQ', item.assignedTo || '', assignedBy, dateVal
        ]);
      }
    });
    SpreadsheetApp.flush();
    return api_getAssets(token);
  });
}

function api_deleteAsset(token, id) {
  requireAuth_(token);
  return withLock_(function () {
    var sheet = getSheet_(SHEET_ASSETS);
    var rows = sheetToObjects_(SHEET_ASSETS);
    var row = rows.filter(function (r) { return r.Id === id; })[0];
    if (row) sheet.deleteRow(row._row);
    SpreadsheetApp.flush();
    return api_getAssets(token);
  });
}

function api_resetDatabase(token, scope, confirmText) {
  var me = requireAuth_(token);
  requireAdmin_(me);
  if (confirmText !== 'RESET') throw new Error('Type RESET to confirm.');
  withLock_(function () {
    var targets = scope === 'all' ? [SHEET_TASKS, SHEET_BUDGETS, SHEET_CHAT, SHEET_FILE_REQUESTS, SHEET_EVENTS, SHEET_OBJECTIVES, SHEET_PLANNER, SHEET_ASSETS] : [scope];
    targets.forEach(function (name) {
      if (!HEADERS[name] || name === SHEET_USERS) return;
      var sheet = getSheet_(name);
      var last = sheet.getLastRow();
      if (last > 1) sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).clearContent();
    });
  });
  return api_getInitialData(token);
}

function test_diagnostic_drive_api() {
  try {
    var iterator = DriveApp.getFiles();
    var count = 0;
    while (iterator.hasNext() && count < 5) {
      var file = iterator.next();
      Logger.log('Success! Found File: ' + file.getName() + ' (ID: ' + file.getId() + ')');
      count++;
    }
    if (count === 0) {
      Logger.log('The DriveApp call worked perfectly, but your Google Drive contains absolutely 0 files.');
    }
  } catch (err) {
    Logger.log('Diagnostic Error: ' + err.message);
  }
}