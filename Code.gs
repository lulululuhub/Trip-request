// UICUSA Business Trip Suite - Code.gs
// Auto-detects Google account for admin vs employee view

// -- CONFIGURATION --
var CONFIG = {
  SPREADSHEET_ID:        '1j762rgWVPNGemVp7FYEQIXG4NLyLkGoPtGNSyaxagZs',
  CALENDAR_ID:           'c_9afb084439b154efd07562acf496b2f05d3601a9e459cf6aaa09453470ab591d@group.calendar.google.com',
  GENERAL_MANAGER_EMAIL: 'lchao@uicusa.com',  // Testing: change to edwin.young@uicusa.com for production
  GENERAL_MANAGER_NAME:  'Edwin Young',
  COMPANY_NAME:          'UICUSA',
  WEBAPP_URL:            'https://script.google.com/a/macros/uicusa.com/s/AKfycbznLSIzhNB6seFHli61D9spiP4MrlPDmrjg2gFOlc26VE7RyPn8NYQJw_7MboFrc7nXOQ/exec',
  RECEIPT_FOLDER_NAME:   'UICUSA_Expense_Receipts',
  ADMIN_EMAILS:          ['lchao@uicusa.com', 'cleo@uicusa.com'],
};

var SHEETS = {
  REQUESTS:  'TripRequests',
  REPORTS:   'TripReports',
  EXPENSES:  'ExpenseReports',
  SALES:     'SalesFollowups',
  EMPLOYEES: 'Employees',
  LOG:       'ApprovalLog',
};

// -- SESSION & ROLE DETECTION --

function getCurrentUser() {
  var email = Session.getActiveUser().getEmail();
  var isAdmin = CONFIG.ADMIN_EMAILS.indexOf(email) > -1;
  var empInfo = getEmployeeByEmail_(email);
  return {
    email:     email,
    isAdmin:   isAdmin,
    name:      empInfo ? empInfo.name : email,
    dept:      empInfo ? empInfo.dept : '',
    role:      empInfo ? empInfo.role : 'Employee',
    title:     empInfo ? empInfo.title : '',
    isManager: empInfo ? empInfo.role === 'Manager' : false,
  };
}

// -- WEB APP ENTRY POINTS --

function doGet(e) {
  var page   = (e && e.parameter && e.parameter.page)   ? e.parameter.page   : 'form';
  var token  = (e && e.parameter && e.parameter.token)  ? e.parameter.token  : '';
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  if (page === 'approve' && token) { return handleApprovalPage_(token, action); }
  if (page === 'status'  && token) { return handleStatusPage_(token); }
  return buildMainApp_();
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var result;
    if      (data.action === 'submit_request')        { result = submitTripRequest_(data); }
    else if (data.action === 'submit_trip_report')    { result = submitTripReport_(data); }
    else if (data.action === 'submit_expense')        { result = submitExpense_(data); }
    else if (data.action === 'submit_sales_followup') { result = submitSalesFollowup_(data); }
    else { result = { success: false, error: 'Unknown action: ' + data.action }; }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// -- SERVER FUNCTIONS CALLED FROM CLIENT --

function getSessionInfo() {
  return getCurrentUser();
}

function getMyHistory() {
  var user = getCurrentUser();
  var ss   = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  function readMine(sheetName, emailCol) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) { return []; }
    var rows = sh.getDataRange().getValues();
    var h    = rows[0];
    var col  = h.indexOf(emailCol);
    if (col < 0) { return []; }
    return rows.slice(1)
      .filter(function(r){ return String(r[col]).trim().toLowerCase() === user.email.toLowerCase(); })
      .map(function(r){ return rowToObject_(r, h); });
  }

  return {
    requests: readMine(SHEETS.REQUESTS, 'EmployeeEmail'),
    reports:  readMine(SHEETS.REPORTS,  'EmployeeEmail'),
    expenses: readMine(SHEETS.EXPENSES, 'EmployeeEmail'),
    sales:    readMine(SHEETS.SALES,    'EmployeeEmail'),
  };
}

function getReportData(fromDate, toDate, deptFilter, statusFilter) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  function readSheet(name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) { return []; }
    var rows = sh.getDataRange().getValues();
    var h    = rows[0];
    return rows.slice(1).map(function(r){ return rowToObject_(r, h); });
  }
  var requests = readSheet(SHEETS.REQUESTS);
  var reports  = readSheet(SHEETS.REPORTS);
  var expenses = readSheet(SHEETS.EXPENSES);
  var sales    = readSheet(SHEETS.SALES);

  var from = fromDate ? new Date(fromDate) : null;
  var to   = toDate   ? new Date(toDate)   : null;
  if (to) { to.setHours(23,59,59); }

  function applyFilters(arr, deptField) {
    return arr.filter(function(r) {
      if (deptFilter && String(r[deptField]||'').trim() !== deptFilter) { return false; }
      if (statusFilter && String(r['Status']||'').trim() !== statusFilter) { return false; }
      if (from || to) {
        var d = r['SubmittedAt'] ? new Date(r['SubmittedAt']) : null;
        if (d) {
          if (from && d < from) { return false; }
          if (to   && d > to)   { return false; }
        }
      }
      return true;
    });
  }

  return {
    requests: applyFilters(requests, 'Department'),
    reports:  reports,
    expenses: applyFilters(expenses, 'Department'),
    sales:    sales,
  };
}

function getReportsPageHtml() {
  return HtmlService.createHtmlOutputFromFile('ReportsPage').getContent();
}

// -- RECEIPT UPLOAD --

function getReceiptUploadUrl(expenseId) {
  var folder = getOrCreateReceiptFolder_();
  var subFolder;
  try {
    var iter = folder.getFoldersByName(expenseId);
    subFolder = iter.hasNext() ? iter.next() : folder.createFolder(expenseId);
  } catch(e) {
    subFolder = folder.createFolder(expenseId);
  }
  return { folderId: subFolder.getId(), folderUrl: subFolder.getUrl() };
}

function saveReceiptFile(expenseId, fileName, base64Data, mimeType) {
  try {
    var folder = getOrCreateReceiptFolder_();
    var subFolder;
    var iter = folder.getFoldersByName(expenseId);
    subFolder = iter.hasNext() ? iter.next() : folder.createFolder(expenseId);
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    var file = subFolder.createFile(blob);
    // Save receipt URL back to expense sheet
    updateExpenseReceiptUrl_(expenseId, file.getUrl(), subFolder.getUrl());
    return { success: true, fileUrl: file.getUrl(), folderUrl: subFolder.getUrl() };
  } catch (err) {
    Logger.log('Receipt save error: ' + err.message);
    return { success: false, error: err.message };
  }
}

function getOrCreateReceiptFolder_() {
  var folders = DriveApp.getFoldersByName(CONFIG.RECEIPT_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.RECEIPT_FOLDER_NAME);
}

function updateExpenseReceiptUrl_(expenseId, fileUrl, folderUrl) {
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.EXPENSES);
  var rows  = sheet.getDataRange().getValues();
  var h     = rows[0];
  var idCol = h.indexOf('ExpenseID');
  var urlCol = h.indexOf('ReceiptFolderUrl');
  if (idCol < 0 || urlCol < 0) { return; }
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][idCol] === expenseId) {
      sheet.getRange(i+1, urlCol+1).setValue(folderUrl);
      return;
    }
  }
}

// -- TRIP REQUEST --

function submitTripRequest_(data) {
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.REQUESTS);
  var requestId = 'TRIP-' + new Date().getFullYear() + '-' + Utilities.getUuid().substring(0,6).toUpperCase();
  var token     = Utilities.getUuid();

  var approver = getFirstApprover_(data.employeeEmail, data.department);
  if (!approver) { return { success: false, error: 'No approver found for department: ' + data.department }; }

  sheet.appendRow([
    requestId, token, new Date().toISOString(), 'PENDING_DEPT',
    data.employeeName, data.employeeEmail, data.department, approver.email,
    data.destination, data.clientName, data.departureDate, data.returnDate,
    data.purpose, data.justification,
    data.estAirfare||0, data.estHotel||0, data.estNights||0,
    data.estPerDiem||75, data.estDays||0, data.estOther||0, data.estTotal||0,
    '','','','','','','',''
  ]);

  sendApproverEmail_(requestId, token, data, approver, 'trip');
  sendEmployeeConfirmation_(requestId, token, data.employeeEmail, data.employeeName,
    'Trip Request', data.destination, data.departureDate, data.returnDate);
  return { success: true, requestId: requestId };
}

// -- TRIP REPORT --

function submitTripReport_(data) {
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.REPORTS);
  var reportId = 'RPT-' + new Date().getFullYear() + '-' + Utilities.getUuid().substring(0,6).toUpperCase();
  var token    = Utilities.getUuid();

  var approver = getFirstApprover_(data.employeeEmail, data.department||'');
  if (!approver) { return { success: false, error: 'No approver found for: ' + data.employeeEmail }; }

  sheet.appendRow([
    reportId, token, new Date().toISOString(), 'PENDING_DEPT',
    data.employeeName, data.employeeEmail, data.department||'', approver.email,
    data.requestId||'', data.destination, data.clientName, data.departureDate, data.returnDate,
    data.contacts||'', data.objectives||'', data.outcomes||'', data.challenges||'',
    data.intel||'', data.actions||'', data.timeline||'', data.recommendations||'',
    '','','','','','',''
  ]);

  sendApproverEmail_(reportId, token, {
    employeeName:  data.employeeName,
    employeeEmail: data.employeeEmail,
    department:    data.department||'',
    destination:   data.destination,
    clientName:    data.clientName,
    departureDate: data.departureDate,
    returnDate:    data.returnDate,
    purpose:       'Trip report review',
    justification: (data.outcomes||'').substring(0,300),
    estTotal:      0,
  }, approver, 'report');

  sendEmployeeConfirmation_(reportId, token, data.employeeEmail, data.employeeName,
    'Trip Report', data.destination, data.departureDate, data.returnDate);

  return { success: true, reportId: reportId };
}

// -- EXPENSE REPORT --

function submitExpense_(data) {
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.EXPENSES);
  var expenseId = 'EXP-' + new Date().getFullYear() + '-' + Utilities.getUuid().substring(0,6).toUpperCase();
  var token     = Utilities.getUuid();

  var approver = getFirstApprover_(data.employeeEmail, data.department);
  if (!approver) { approver = { name: 'Manager', email: CONFIG.GENERAL_MANAGER_EMAIL }; }

  // Create receipt folder proactively
  var folder = getOrCreateReceiptFolder_();
  var sub    = folder.createFolder(expenseId);

  sheet.appendRow([
    expenseId, token, new Date().toISOString(), 'PENDING_DEPT',
    data.employeeName, data.employeeEmail, data.department||'',
    data.requestId||'', data.reference||'', data.submissionDate,
    JSON.stringify(data.items||[]), data.totalAmount||0,
    data.notes||'', approver.email,
    sub.getUrl(),   // ReceiptFolderUrl
    '','','','','','',''
  ]);

  sendApproverEmail_(expenseId, token, {
    employeeName:  data.employeeName,
    employeeEmail: data.employeeEmail,
    department:    data.department||'',
    destination:   data.reference||'',
    clientName:    data.requestId||'N/A',
    departureDate: data.submissionDate,
    returnDate:    data.submissionDate,
    purpose:       'Expense reimbursement',
    justification: data.notes||'',
    estTotal:      data.totalAmount||0,
    items:         data.items||[],
  }, approver, 'expense');

  sendEmployeeConfirmation_(expenseId, token, data.employeeEmail, data.employeeName,
    'Expense Report', data.reference||'', data.submissionDate, data.submissionDate);

  return {
    success:    true,
    expenseId:  expenseId,
    receiptUrl: sub.getUrl(),
    folderId:   sub.getId()
  };
}

// -- SALES FOLLOW-UP --

function submitSalesFollowup_(data) {
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.SALES);
  var followupId = 'SFU-' + new Date().getFullYear() + '-' + Utilities.getUuid().substring(0,6).toUpperCase();
  sheet.appendRow([
    followupId, new Date().toISOString(),
    data.employeeName, data.employeeEmail,
    data.companyName, data.contactName, data.contactEmail, data.contactPhone||'',
    data.meetingDate, data.stage, data.dealValue||0, data.closeDate||'',
    data.products||'', data.notes, data.rating||0, data.ratingLabel||'',
    data.competitors||'', data.decisionMaker||'', data.checklist||'', data.extra||''
  ]);
  return { success: true, followupId: followupId };
}

// -- APPROVAL --

function processApproval(token, decision, comment, approverEmail) {
  var result = processApprovalForSheet_(SHEETS.REQUESTS, token, decision, comment, approverEmail, 'trip');
  if (result) { return result; }
  result = processApprovalForSheet_(SHEETS.EXPENSES, token, decision, comment, approverEmail, 'expense');
  if (result) { return result; }
  result = processApprovalForSheet_(SHEETS.REPORTS, token, decision, comment, approverEmail, 'report');
  if (result) { return result; }
  return { success: false, error: 'Request not found or already processed.' };
}

function processApprovalForSheet_(sheetName, token, decision, comment, approverEmail, type) {
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  var rows  = sheet.getDataRange().getValues();
  var h     = rows[0];
  var now   = new Date().toISOString();
  var col   = h.indexOf('Token');
  if (col < 0) { return null; }

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][col] !== token) { continue; }

    var statusIdx = h.indexOf('Status');
    var status    = rows[i][statusIdx];
    var rowData   = rowToObject_(rows[i], h);
    var id        = rowData['RequestID'] || rowData['ExpenseID'] || '';

    if (decision === 'approve') {
      if (status === 'PENDING_DEPT') {
        // Forward to GM
        sheet.getRange(i+1, statusIdx+1).setValue('PENDING_GM');
        sheet.getRange(i+1, h.indexOf('DeptApprovedAt')+1).setValue(now);
        sheet.getRange(i+1, h.indexOf('DeptApprovedBy')+1).setValue(approverEmail);
        sendGMEmail_(rowData, token, comment, type);
        notifyEmployeeDeptApproved_(rowData, type);
        logAction_(id, approverEmail, 'DEPT_APPROVED', comment);
        return { success: true, message: 'Approved and forwarded to General Manager.' };
      }
      if (status === 'PENDING_GM') {
        sheet.getRange(i+1, statusIdx+1).setValue('APPROVED');
        sheet.getRange(i+1, h.indexOf('GMApprovedAt')+1).setValue(now);
        sheet.getRange(i+1, h.indexOf('GMApprovedBy')+1).setValue(approverEmail);
        if (type === 'trip') {
          var evtId = addToCalendar_(rowData);
          var calIdx = h.indexOf('CalendarEventId');
          if (calIdx >= 0) { sheet.getRange(i+1, calIdx+1).setValue(evtId); }
        }
        sendFinalApprovalEmail_(rowData, type);
        logAction_(id, approverEmail, 'GM_APPROVED', comment);
        var msg = type === 'trip'
          ? 'Trip fully approved and added to the attendance calendar.'
          : 'Expense report approved. Accounting team notified.';
        return { success: true, message: msg };
      }
    }

    if (decision === 'reject') {
      sheet.getRange(i+1, statusIdx+1).setValue('REJECTED');
      sheet.getRange(i+1, h.indexOf('RejectedAt')+1).setValue(now);
      sheet.getRange(i+1, h.indexOf('RejectedBy')+1).setValue(approverEmail);
      sheet.getRange(i+1, h.indexOf('RejectionReason')+1).setValue(comment);
      sendRejectionEmail_(rowData, comment, approverEmail, type);
      logAction_(id, approverEmail, 'REJECTED', comment);
      return { success: true, message: 'Rejected. Employee has been notified.' };
    }
  }
  return null;
}

// -- APPROVAL ROUTING --
// Managers go directly to GM (skip dept approval)
// Employees go to their dept manager first

function getFirstApprover_(employeeEmail, department) {
  var emp = getEmployeeByEmail_(employeeEmail);
  if (emp && emp.role === 'Manager') {
    // Manager: first approver is GM directly
    return { name: CONFIG.GENERAL_MANAGER_NAME, email: CONFIG.GENERAL_MANAGER_EMAIL, isGM: true };
  }
  // Employee: first approver is dept manager
  var mgr = getDeptManager_(department) || getDeptManagerByEmail_(employeeEmail);
  return mgr || null;
}

function submitTripRequest_forManager_(data, approver) {
  // When submitter is a manager, status starts as PENDING_GM not PENDING_DEPT
  return 'PENDING_GM';
}

// -- CALENDAR --

function addToCalendar_(data) {
  try {
    var cal   = CONFIG.CALENDAR_ID ? CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) : CalendarApp.getDefaultCalendar();
    var start = new Date(data['DepartureDate']);
    var end   = new Date(data['ReturnDate']);
    end.setDate(end.getDate() + 1);
    var title = '[Business Trip] ' + data['EmployeeName'] + ' - ' + data['Destination'];
    var desc  = 'Employee: '   + data['EmployeeName'] +
                '\nDept: '     + (data['Department']||'') +
                '\nDest: '     + data['Destination'] +
                '\nClient: '   + (data['ClientName']||'') +
                '\nPurpose: '  + (data['Purpose']||'') +
                '\nReq ID: '   + (data['RequestID']||'') +
                '\nApproved: ' + (data['GMApprovedBy']||'') +
                '\nBudget: $'  + Number(data['EstTotal']||0).toFixed(2);
    var event = cal.createAllDayEvent(title, start, end, {
      description: desc, guests: data['EmployeeEmail'], sendInvites: true
    });
    return event.getId();
  } catch (err) {
    Logger.log('Calendar error: ' + err.message);
    return 'ERROR: ' + err.message;
  }
}

// -- EMAILS --

function sendApproverEmail_(id, token, data, approver, type) {
  var base  = CONFIG.WEBAPP_URL + '?page=approve&token=' + token;
  var label = type === 'expense' ? 'Expense Report' : 'Trip Request';
  var isGM  = approver.isGM || false;
  var rows  = [
    [label + ' ID', id],
    ['Employee',    data.employeeName + ' (' + data.employeeEmail + ')'],
    ['Department',  data.department||''],
  ];
  if (type === 'expense') {
    rows = rows.concat([
      ['Reference',     data.destination||data.reference||''],
      ['Total amount',  '$' + Number(data.estTotal||0).toFixed(2)],
      ['Notes',         (data.justification||data.notes||'None').substring(0,200)],
    ]);
    if (data.items && data.items.length) {
      var itemSummary = data.items.slice(0,5).map(function(it){
        return it.description + ' (' + it.category + '): $' + Number(it.amount||0).toFixed(2);
      }).join(' | ');
      rows.push(['Items', itemSummary]);
    }
  } else if (type === 'report') {
    rows = rows.concat([
      ['Destination',   data.destination],
      ['Client',        data.clientName],
      ['Dates',         data.departureDate + ' to ' + data.returnDate],
      ['Key outcomes',  (data.justification||'').substring(0,300)],
    ]);
  } else {
    rows = rows.concat([
      ['Destination',   data.destination],
      ['Client',        data.clientName],
      ['Dates',         data.departureDate + ' to ' + data.returnDate],
      ['Purpose',       data.purpose],
      ['Justification', (data.justification||'').substring(0,300)],
      ['Est. budget',   '$' + Number(data.estTotal||0).toFixed(2)],
    ]);
  }
  if (isGM) { rows.push(['Note', 'This request is from a department manager — forwarded directly to GM for approval.']); }

  var body = buildEmailHtml_({
    heading:    label + ' — ' + (isGM ? 'GM approval needed' : 'your approval needed'),
    badge:      isGM ? 'Pending GM approval' : 'Pending 1st approval',
    badgeColor: isGM ? '#1a3a6b' : '#e65100',
    rows:       rows,
    approveUrl: base + '&action=approve',
    rejectUrl:  base + '&action=reject',
    viewUrl:    base,
    footerNote: isGM
      ? 'This request is from a manager and requires your final approval.'
      : 'You are the 1st approver. Approving will forward to the General Manager for final approval.',
  });
  GmailApp.sendEmail(approver.email,
    '[' + CONFIG.COMPANY_NAME + '] ' + label + ' approval needed — ' + data.employeeName,
    '', { htmlBody: body });
}

function sendGMEmail_(data, token, note, type) {
  var base  = CONFIG.WEBAPP_URL + '?page=approve&token=' + token;
  var label = type === 'expense' ? 'Expense Report' : 'Trip Request';
  var id    = data['RequestID'] || data['ExpenseID'] || '';
  var rows  = [
    [label + ' ID', id],
    ['Employee',    data['EmployeeName'] + ' (' + data['EmployeeEmail'] + ')'],
    ['Department',  data['Department']||''],
    ['Reference',   data['Destination']||data['Reference']||''],
    ['Total',       '$' + Number(data['EstTotal']||data['TotalAmount']||0).toFixed(2)],
    ['1st approver note', note||'None'],
  ];
  var body = buildEmailHtml_({
    heading: label + ' — final approval needed',
    badge: '1st approved — Pending GM', badgeColor: '#1a3a6b',
    rows: rows,
    approveUrl: base + '&action=approve',
    rejectUrl:  base + '&action=reject',
    viewUrl:    base,
    footerNote: 'Department manager has approved. Your final approval is needed.'
  });
  GmailApp.sendEmail(CONFIG.GENERAL_MANAGER_EMAIL,
    '[' + CONFIG.COMPANY_NAME + '] Final approval needed — ' + (data['EmployeeName']||''),
    '', { htmlBody: body });
}

function sendEmployeeConfirmation_(id, token, email, name, type, ref, d1, d2) {
  var statusUrl = CONFIG.WEBAPP_URL + '?page=status&token=' + token;
  var body = buildEmailHtml_({
    heading: type + ' submitted',
    badge: 'Submitted', badgeColor: '#2e7d32',
    rows: [
      ['ID',    id],
      ['Type',  type],
      ['Ref',   ref||''],
      ['Date',  d1 + (d2 && d2!==d1 ? ' to '+d2 : '')],
    ],
    statusUrl:  statusUrl,
    footerNote: 'Your approver has been notified. Track your request status using the button above. Keep your ID for reference.'
  });
  GmailApp.sendEmail(email,
    '[' + CONFIG.COMPANY_NAME + '] ' + type + ' submitted — ' + id, '', { htmlBody: body });
}

function notifyEmployeeDeptApproved_(data, type) {
  var label = type === 'expense' ? 'Expense Report' : 'Trip Request';
  var id    = data['RequestID'] || data['ExpenseID'] || '';
  var body  = buildEmailHtml_({
    heading: label + ' — 1st approval done',
    badge: '1st approved — awaiting GM', badgeColor: '#1a3a6b',
    rows: [
      ['ID',   id],
      ['Ref',  data['Destination']||data['Reference']||''],
    ],
    footerNote: 'Your department manager approved. Awaiting General Manager final approval.'
  });
  GmailApp.sendEmail(data['EmployeeEmail'],
    '[' + CONFIG.COMPANY_NAME + '] 1st approval done — ' + id, '', { htmlBody: body });
}

function sendFinalApprovalEmail_(data, type) {
  var label   = type === 'expense' ? 'Expense Report' : 'Trip Request';
  var id      = data['RequestID'] || data['ExpenseID'] || '';
  var subject = '[' + CONFIG.COMPANY_NAME + '] APPROVED — ' + label + ' ' + id;
  var footerNote = type === 'trip'
    ? 'Your trip has been added to the UICUSA attendance calendar. A calendar invite will arrive shortly. Safe travels!'
    : 'Your expense report is approved. Accounting will process your reimbursement shortly.';
  var rows = [
    ['ID',         id],
    ['Employee',   data['EmployeeName']||''],
    ['Reference',  data['Destination']||data['Reference']||''],
    ['Amount',     '$' + Number(data['EstTotal']||data['TotalAmount']||0).toFixed(2)],
    ['Approved by', data['GMApprovedBy']||''],
  ];
  var body = buildEmailHtml_({
    heading: label + ' fully approved!',
    badge: 'Fully approved', badgeColor: '#2e7d32',
    rows: rows, footerNote: footerNote
  });
  GmailApp.sendEmail(data['EmployeeEmail'],    subject, '', { htmlBody: body });
  if (data['DeptManagerEmail']) {
    GmailApp.sendEmail(data['DeptManagerEmail'], subject, '', { htmlBody: body });
  }
  if (type === 'expense') {
    var accMgr = getDeptManager_('Accounting');
    if (accMgr) {
      GmailApp.sendEmail(accMgr.email,
        '[' + CONFIG.COMPANY_NAME + '] Expense approved for processing — ' + id, '', { htmlBody: body });
    }
  }
}

function sendRejectionEmail_(data, reason, rejectedBy, type) {
  var label = type === 'expense' ? 'Expense Report' : 'Trip Request';
  var id    = data['RequestID'] || data['ExpenseID'] || '';
  var body  = buildEmailHtml_({
    heading: label + ' — not approved',
    badge: 'Not approved', badgeColor: '#c62828',
    rows: [
      ['ID',          id],
      ['Reference',   data['Destination']||data['Reference']||''],
      ['Rejected by', rejectedBy],
      ['Reason',      reason||'No reason provided'],
    ],
    footerNote: 'Please speak with your manager if you have questions.'
  });
  GmailApp.sendEmail(data['EmployeeEmail'],
    '[' + CONFIG.COMPANY_NAME + '] Not approved — ' + label + ' ' + id, '', { htmlBody: body });
}

// -- EMAIL HTML BUILDER --

function buildEmailHtml_(opts) {
  var rows = opts.rows.map(function(r) {
    return '<tr>'
      + '<td style="padding:7px 12px;color:#666;font-size:13px;white-space:nowrap;vertical-align:top;border-bottom:1px solid #f0f0f0">' + r[0] + '</td>'
      + '<td style="padding:7px 12px;font-size:13px;color:#111;vertical-align:top;border-bottom:1px solid #f0f0f0">' + r[1] + '</td>'
      + '</tr>';
  }).join('');
  var btns = '';
  if (opts.approveUrl) { btns += '<a href="' + opts.approveUrl + '" style="display:inline-block;padding:11px 28px;background:#1a3a6b;color:#fff;text-decoration:none;border-radius:6px;font-size:14px;font-weight:600;margin-right:8px">Approve</a>'; }
  if (opts.rejectUrl)  { btns += '<a href="' + opts.rejectUrl  + '" style="display:inline-block;padding:11px 28px;background:#fff;color:#c62828;text-decoration:none;border-radius:6px;font-size:14px;font-weight:600;border:1px solid #c62828">Reject</a>'; }
  if (opts.statusUrl)  { btns += '<a href="' + opts.statusUrl  + '" style="display:inline-block;padding:11px 28px;background:#f5f5f5;color:#333;text-decoration:none;border-radius:6px;font-size:14px">Track my request</a>'; }
  if (opts.viewUrl && !opts.statusUrl) { btns += '<a href="' + opts.viewUrl + '" style="display:inline-block;padding:11px 28px;background:#f5f5f5;color:#333;text-decoration:none;border-radius:6px;font-size:14px;margin-left:8px">View details</a>'; }
  return '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif">'
    + '<div style="max-width:580px;margin:32px auto;background:#fff;border-radius:8px;border:1px solid #e0e0e0">'
    + '<div style="background:#1a3a6b;padding:18px 24px">'
    + '<span style="background:#fff;color:#1a3a6b;font-size:11px;font-weight:700;padding:4px 8px;border-radius:4px;margin-right:10px">UICUSA</span>'
    + '<span style="color:#fff;font-size:15px;font-weight:500">' + opts.heading + '</span></div>'
    + '<div style="padding:20px 24px 8px">'
    + '<span style="display:inline-block;padding:4px 12px;background:' + opts.badgeColor + '22;color:' + opts.badgeColor + ';font-size:12px;font-weight:700;border-radius:20px;border:1px solid ' + opts.badgeColor + '88;margin-bottom:16px">' + opts.badge + '</span>'
    + '<table style="width:100%;border-collapse:collapse;border:1px solid #e8e8e8">' + rows + '</table></div>'
    + (btns ? '<div style="padding:20px 24px">' + btns + '</div>' : '')
    + '<div style="padding:14px 24px;border-top:1px solid #eee;font-size:12px;color:#888">' + (opts.footerNote||'') + '</div>'
    + '</div></body></html>';
}

// -- PAGE BUILDERS --

function buildMainApp_() {
  var html = HtmlService.createHtmlOutputFromFile('EmployeeForm');
  html.setTitle('UICUSA - Business Trip Suite');
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

function handleApprovalPage_(token, action) {
  var row = findRowByToken_(SHEETS.REQUESTS, token) || findRowByToken_(SHEETS.EXPENSES, token) || findRowByToken_(SHEETS.REPORTS, token);
  if (!row) {
    return HtmlService.createHtmlOutput('<div style="font-family:Arial;padding:40px;text-align:center"><h2>Invalid or expired link</h2></div>');
  }
  return buildApprovalForm_(row.obj, token, action, row.type);
}

function buildApprovalForm_(data, token, preAction, type) {
  var t = HtmlService.createTemplateFromFile('ApprovalForm');
  t.data = data; t.token = token; t.preAction = preAction; t.formType = type||'trip';
  var html = t.evaluate();
  html.setTitle('UICUSA - Approval');
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

function handleStatusPage_(token) {
  var row = findRowByToken_(SHEETS.REQUESTS, token) || findRowByToken_(SHEETS.EXPENSES, token) || findRowByToken_(SHEETS.REPORTS, token);
  if (!row) {
    return HtmlService.createHtmlOutput('<div style="font-family:Arial;padding:40px;text-align:center"><h2>Request not found</h2></div>');
  }
  var t = HtmlService.createTemplateFromFile('StatusPage');
  t.data = row.obj;
  var html = t.evaluate();
  html.setTitle('UICUSA - Status');
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

// -- HELPERS --

function getEmployeeByEmail_(email) {
  if (!email) { return null; }
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.EMPLOYEES);
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().trim().toLowerCase() === email.trim().toLowerCase()) {
      return { name: rows[i][0], email: rows[i][1], dept: rows[i][2], role: rows[i][3], title: rows[i][4]||'' };
    }
  }
  return null;
}

function getDeptManager_(department) {
  if (!department) { return null; }
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEETS.EMPLOYEES);
  var rows  = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][2].toString().trim().toLowerCase() === department.trim().toLowerCase() &&
        rows[i][3].toString().trim().toLowerCase() === 'manager') {
      return { name: rows[i][0], email: rows[i][1] };
    }
  }
  return null;
}

function getDeptManagerByEmail_(email) {
  if (!email) { return null; }
  var emp = getEmployeeByEmail_(email);
  return emp ? getDeptManager_(emp.dept) : null;
}

function findRowByToken_(sheetName, token) {
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) { return null; }
  var rows = sheet.getDataRange().getValues();
  var h    = rows[0];
  var col  = h.indexOf('Token');
  if (col < 0) { return null; }
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][col] === token) {
      return { obj: rowToObject_(rows[i], h), type: sheetName === SHEETS.EXPENSES ? 'expense' : 'trip' };
    }
  }
  return null;
}

function rowToObject_(row, headers) {
  var obj = {};
  headers.forEach(function(h, i){ obj[h] = row[i] !== undefined ? row[i] : ''; });
  return obj;
}

function logAction_(id, actor, action, note) {
  var ss  = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var log = ss.getSheetByName(SHEETS.LOG);
  log.appendRow([new Date().toISOString(), id||'', actor, action, note||'']);
}

// -- ONE-TIME SETUP --

function setupSpreadsheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Trip Requests
  var req = ss.getSheetByName(SHEETS.REQUESTS) || ss.insertSheet(SHEETS.REQUESTS);
  req.clearContents();
  req.appendRow(['RequestID','Token','SubmittedAt','Status','EmployeeName','EmployeeEmail','Department','DeptManagerEmail','Destination','ClientName','DepartureDate','ReturnDate','Purpose','Justification','EstAirfare','EstHotel','EstNights','EstPerDiem','EstDays','EstOther','EstTotal','DeptApprovedAt','DeptApprovedBy','GMApprovedAt','GMApprovedBy','RejectedAt','RejectedBy','RejectionReason','CalendarEventId']);
  req.getRange(1,1,1,29).setFontWeight('bold').setBackground('#1a3a6b').setFontColor('#fff');
  req.setFrozenRows(1);

  // Trip Reports
  var rpt = ss.getSheetByName(SHEETS.REPORTS) || ss.insertSheet(SHEETS.REPORTS);
  rpt.clearContents();
  rpt.appendRow(['ReportID','Token','SubmittedAt','Status','EmployeeName','EmployeeEmail','Department','DeptManagerEmail','RequestID','Destination','ClientName','DepartureDate','ReturnDate','Contacts','Objectives','Outcomes','Challenges','Intel','Actions','Timeline','Recommendations','DeptApprovedAt','DeptApprovedBy','GMApprovedAt','GMApprovedBy','RejectedAt','RejectedBy','RejectionReason']);
  rpt.getRange(1,1,1,28).setFontWeight('bold').setBackground('#1a3a6b').setFontColor('#fff');
  rpt.setFrozenRows(1);

  // Expense Reports
  var exp = ss.getSheetByName(SHEETS.EXPENSES) || ss.insertSheet(SHEETS.EXPENSES);
  exp.clearContents();
  exp.appendRow(['ExpenseID','Token','SubmittedAt','Status','EmployeeName','EmployeeEmail','Department','RequestID','Reference','SubmissionDate','ItemsJSON','TotalAmount','Notes','DeptManagerEmail','ReceiptFolderUrl','DeptApprovedAt','DeptApprovedBy','GMApprovedAt','GMApprovedBy','RejectedAt','RejectedBy','RejectionReason']);
  exp.getRange(1,1,1,22).setFontWeight('bold').setBackground('#1a3a6b').setFontColor('#fff');
  exp.setFrozenRows(1);

  // Sales
  var sfu = ss.getSheetByName(SHEETS.SALES) || ss.insertSheet(SHEETS.SALES);
  sfu.clearContents();
  sfu.appendRow(['FollowupID','SubmittedAt','EmployeeName','EmployeeEmail','CompanyName','ContactName','ContactEmail','ContactPhone','MeetingDate','Stage','DealValue','CloseDate','Products','Notes','Rating','RatingLabel','Competitors','DecisionMaker','Checklist','Extra']);
  sfu.getRange(1,1,1,20).setFontWeight('bold').setBackground('#1a3a6b').setFontColor('#fff');
  sfu.setFrozenRows(1);

  // Employees
  var emp = ss.getSheetByName(SHEETS.EMPLOYEES) || ss.insertSheet(SHEETS.EMPLOYEES);
  emp.clearContents();
  emp.appendRow(['Name','Email','Department','Role','Title']);
  emp.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1a3a6b').setFontColor('#fff');
  var staff = [
    ['Edwin Young',    'Edwin@uicusa.com',    'General Manager',      'Manager',  'General Manager'],
    ['Lulu Chao',      'lchao@uicusa.com',           'HR',                   'Manager',  'HR Manager'],
    ['Cleo Chiang',    'cleo@uicusa.com',             'Accounting',           'Manager',  'Accounting Manager'],
    ['Robin Tang',     'robin@uicusa.com',            'Engineering',          'Manager',  'Engineering Director'],
    ['Linda Liu',      'linda@uicusa.com',            'Operations',           'Manager',  'Operations Manager'],
    ['Robert Wang',    'rwang@uicusa.com',            'Sales',                'Manager',  'Executive Sales Manager'],
    ['Carlos Sedano',  'carlos@uicusa.com',           'Latin American Sales', 'Manager',  'Latin American Territory Manager'],
    ['Kevin Chang',    'kevin@uicusa.com',            'Engineering',          'Employee', 'Software Engineer'],
    ['Jerry Wang',     'jwang@uicusa.com',            'Engineering',          'Employee', 'Project Management Analyst'],
    ['Steve Wu',       'steve@uicusa.com',            'Engineering',          'Employee', 'Software Engineer'],
    ['Roger Navai',    'roger@uicusa.com',            'Engineering',          'Employee', 'Technical Support'],
    ['Fausto Sandoval','fausto.sandoval@uicusa.com',  'Latin American Sales', 'Employee', 'Latin American Sales Manager'],
    ['Fabian Ramirez', 'fabian.ramirez@uicusa.com',  'Latin American Sales', 'Employee', 'Latin American Technical Support'],
  ];
  staff.forEach(function(r){ emp.appendRow(r); });
  emp.getRange(2,4,7,1).setBackground('#e8edf4');
  emp.setFrozenRows(1);
  emp.autoResizeColumns(1,5);

  // Log
  var log = ss.getSheetByName(SHEETS.LOG) || ss.insertSheet(SHEETS.LOG);
  log.clearContents();
  log.appendRow(['Timestamp','ID','Actor','Action','Note']);
  log.getRange(1,1,1,5).setFontWeight('bold').setBackground('#1a3a6b').setFontColor('#fff');
  log.setFrozenRows(1);

  Logger.log('Setup complete.');
  SpreadsheetApp.getUi().alert('Setup complete!\n\nAll sheets created.\nIMPORTANT: Update real email addresses in the Employees sheet before going live.\n\nReceipt folder "UICUSA_Expense_Receipts" will be created automatically in Google Drive on first expense submission.');
}
