// UICUSA Reports & Export Functions
// ADD this entire file as a new Script file in Apps Script named "ReportsFunctions"

function getReportsPageHtml() {
  return HtmlService.createHtmlOutputFromFile('ReportsPage').getContent();
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
      if (deptFilter  && String(r[deptField]||'').trim() !== deptFilter)  { return false; }
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

function exportToExcel(fromDate, toDate, deptFilter, statusFilter) {
  try {
    var data  = getReportData(fromDate, toDate, deptFilter, statusFilter);
    var title = 'UICUSA_Export_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    var ss    = SpreadsheetApp.create(title);
    var ssId  = ss.getId();

    buildTripRequestsSheet_(ss.getSheets()[0], data.requests);
    buildExpenseSheet_(ss.insertSheet('Expense Reports'), data.expenses);
    buildTripReportsSheet_(ss.insertSheet('Trip Reports'), data.reports);
    buildSalesSheet_(ss.insertSheet('Sales Follow-ups'), data.sales);
    buildSummarySheet_(ss.insertSheet('Summary'), data);

    DriveApp.getFileById(ssId).setSharing(
      DriveApp.Access.DOMAIN_WITH_LINK,
      DriveApp.Permission.VIEW
    );
    return 'https://docs.google.com/spreadsheets/d/' + ssId + '/edit';
  } catch(err) {
    Logger.log('Export error: ' + err.message);
    return null;
  }
}

function buildTripRequestsSheet_(sh, rows) {
  sh.setName('Trip Requests');
  var headers = ['Request ID','Employee','Email','Department','Destination','Client',
                 'Departure','Return','Purpose','Airfare','Hotel/Night','Nights',
                 'Per Diem/Day','Days','Other','Est. Total','Status',
                 'Dept Approved By','Dept Approved At','GM Approved By','GM Approved At',
                 'Rejected By','Rejection Reason','Submitted At'];
  sh.appendRow(headers);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1a3a6b').setFontColor('#fff').setFontFamily('Arial').setFontSize(10);
  rows.forEach(function(r){
    sh.appendRow([
      r['RequestID'], r['EmployeeName'], r['EmployeeEmail'], r['Department'],
      r['Destination'], r['ClientName'], r['DepartureDate'], r['ReturnDate'], r['Purpose'],
      Number(r['EstAirfare']||0), Number(r['EstHotel']||0), Number(r['EstNights']||0),
      Number(r['EstPerDiem']||0), Number(r['EstDays']||0), Number(r['EstOther']||0), Number(r['EstTotal']||0),
      r['Status'], r['DeptApprovedBy'], r['DeptApprovedAt'],
      r['GMApprovedBy'], r['GMApprovedAt'], r['RejectedBy'], r['RejectionReason'], r['SubmittedAt']
    ]);
  });
  if (rows.length > 0) {
    sh.getRange(2,10,rows.length,7).setNumberFormat('$#,##0.00');
    for (var i=0; i<rows.length; i++) {
      var cell = sh.getRange(i+2,17);
      var s = rows[i]['Status'];
      if (s==='APPROVED')      { cell.setBackground('#e8f5e9').setFontColor('#2e7d32').setFontWeight('bold'); }
      else if (s==='REJECTED') { cell.setBackground('#ffebee').setFontColor('#c62828').setFontWeight('bold'); }
      else                     { cell.setBackground('#fff3e0').setFontColor('#e65100').setFontWeight('bold'); }
    }
  }
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}

function buildExpenseSheet_(sh, rows) {
  // Header matching UICUSA expense format
  sh.getRange('A1').setValue('UICUSA — Expense Report Export');
  sh.getRange('A1:N1').merge().setFontWeight('bold').setFontSize(13).setBackground('#1a3a6b').setFontColor('#fff').setHorizontalAlignment('center').setFontFamily('Arial');
  sh.getRange('A2').setValue('Generated: ' + new Date().toLocaleString());
  sh.getRange('A2').setFontColor('#888').setFontSize(10);

  var headers = ['Expense ID','Employee','Email','Department','Date Submitted',
                 'Trip Ref / Request ID','Lodging','Airfare','Ground Transportation',
                 'Reimbursement $/mile','Meals & Tips','Conferences & Seminars',
                 'Entertainment','Office Supply','Mobile / Telecom','Miscellaneous',
                 'Total Amount','Status','Approved By','Receipt Folder','Notes'];
  sh.getRange(4,1,1,headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1a3a6b').setFontColor('#fff')
    .setFontFamily('Arial').setFontSize(10).setWrap(true);
  sh.setRowHeight(4,40);

  var startRow = 5;
  rows.forEach(function(r, idx) {
    var cats = { lodging:0, airfare:0, ground:0, mileage:0, meals:0, conf:0, entertain:0, office:0, mobile:0, misc:0 };
    try {
      var items = JSON.parse(r['ItemsJSON']||'[]');
      items.forEach(function(it){
        var c = String(it.category||'').toLowerCase();
        var a = parseFloat(it.amount||0);
        if      (c.indexOf('lodging')>-1||c.indexOf('hotel')>-1)           { cats.lodging   += a; }
        else if (c.indexOf('airfare')>-1)                                   { cats.airfare   += a; }
        else if (c.indexOf('ground')>-1||c.indexOf('transport')>-1)        { cats.ground    += a; }
        else if (c.indexOf('mile')>-1||c.indexOf('reimburse')>-1)          { cats.mileage   += a; }
        else if (c.indexOf('meal')>-1||c.indexOf('tip')>-1)                { cats.meals     += a; }
        else if (c.indexOf('conf')>-1||c.indexOf('seminar')>-1)            { cats.conf      += a; }
        else if (c.indexOf('entertain')>-1)                                 { cats.entertain += a; }
        else if (c.indexOf('office')>-1||c.indexOf('supply')>-1)           { cats.office    += a; }
        else if (c.indexOf('mobile')>-1||c.indexOf('telecom')>-1)          { cats.mobile    += a; }
        else                                                                 { cats.misc      += a; }
      });
    } catch(e){}
    var row = startRow + idx;
    sh.getRange(row,1,1,21).setValues([[
      r['ExpenseID'], r['EmployeeName'], r['EmployeeEmail'], r['Department']||'',
      r['SubmissionDate'], r['Reference']||r['RequestID']||'',
      cats.lodging, cats.airfare, cats.ground, cats.mileage,
      cats.meals, cats.conf, cats.entertain, cats.office, cats.mobile, cats.misc,
      Number(r['TotalAmount']||0), r['Status'],
      r['GMApprovedBy']||r['DeptApprovedBy']||'',
      r['ReceiptFolderUrl']||'', r['Notes']||''
    ]]);
    sh.getRange(row,7,1,11).setNumberFormat('$#,##0.00');
    sh.getRange(row,17,1,1).setNumberFormat('$#,##0.00').setFontWeight('bold');
    var sc = sh.getRange(row,18);
    if (r['Status']==='APPROVED')      { sc.setBackground('#e8f5e9').setFontColor('#2e7d32').setFontWeight('bold'); }
    else if (r['Status']==='REJECTED') { sc.setBackground('#ffebee').setFontColor('#c62828').setFontWeight('bold'); }
    else                               { sc.setBackground('#fff3e0').setFontColor('#e65100').setFontWeight('bold'); }
    if (idx%2===1) { sh.getRange(row,1,1,16).setBackground('#f9f9f9'); }
  });

  // Subtotals
  if (rows.length > 0) {
    var totRow = startRow + rows.length;
    sh.getRange(totRow,1).setValue('SUBTOTALS').setFontWeight('bold').setBackground('#e8edf4');
    for (var c=7; c<=17; c++) {
      var col = colLetter_(c);
      sh.getRange(totRow,c).setFormula('=SUM('+col+startRow+':'+col+(totRow-1)+')')
        .setNumberFormat('$#,##0.00').setFontWeight('bold').setBackground('#e8edf4');
    }
    // Accounting summary block
    var accRow = totRow + 2;
    sh.getRange(accRow,1).setValue('FOR ACCOUNTING').setFontWeight('bold').setBackground('#1a3a6b').setFontColor('#fff');
    var accItems = [
      ['Lodging',          'G'],['Airfare',        'H'],['Ground Transport','I'],
      ['Mileage',          'J'],['Meals & Tips',   'K'],['Conferences',     'L'],
      ['Entertainment',    'M'],['Office Supplies', 'N'],['Mobile/Telecom',  'O'],
      ['Miscellaneous',    'P'],['GRAND TOTAL',     'Q']
    ];
    accItems.forEach(function(ai, i){
      sh.getRange(accRow+1+i,1).setValue(ai[0]).setFontWeight(ai[0]==='GRAND TOTAL'?'bold':'normal');
      sh.getRange(accRow+1+i,2)
        .setFormula('=SUM('+ai[1]+startRow+':'+ai[1]+(totRow-1)+')')
        .setNumberFormat('$#,##0.00').setFontWeight(ai[0]==='GRAND TOTAL'?'bold':'normal');
    });
  }
  sh.setFrozenRows(4); sh.setFrozenColumns(2);
  sh.autoResizeColumns(1, headers.length);
}

function buildTripReportsSheet_(sh, rows) {
  var headers = ['Report ID','Submitted','Employee','Email','Trip Request ID','Destination',
                 'Client','Departure','Return','Contacts Met','Objectives','Outcomes',
                 'Challenges','Market Intelligence','Action Items','Timeline','Recommendations'];
  sh.appendRow(headers);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1a3a6b').setFontColor('#fff').setFontFamily('Arial').setFontSize(10);
  rows.forEach(function(r){
    sh.appendRow([
      r['ReportID'], r['SubmittedAt'], r['EmployeeName'], r['EmployeeEmail'],
      r['RequestID']||'', r['Destination'], r['ClientName'],
      r['DepartureDate'], r['ReturnDate'], r['Contacts']||'',
      r['Objectives']||'', r['Outcomes']||'', r['Challenges']||'',
      r['Intel']||'', r['Actions']||'', r['Timeline']||'', r['Recommendations']||''
    ]);
  });
  sh.setFrozenRows(1);
  [11,12,13,14,15].forEach(function(c){ sh.setColumnWidth(c,200); });
  sh.getRange(1,1,1+rows.length,17).setWrap(true).setVerticalAlignment('top');
  sh.autoResizeColumns(1,10);
}

function buildSalesSheet_(sh, rows) {
  var headers = ['Follow-up ID','Submitted','Employee','Email','Company','Contact',
                 'Contact Email','Contact Phone','Meeting Date','Stage','Deal Value',
                 'Close Date','Products','Notes','Rating','Rating Label',
                 'Competitors','Decision Maker','Checklist','Extra Notes'];
  sh.appendRow(headers);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1a3a6b').setFontColor('#fff').setFontFamily('Arial').setFontSize(10);
  rows.forEach(function(r){
    sh.appendRow([
      r['FollowupID'], r['SubmittedAt'], r['EmployeeName'], r['EmployeeEmail'],
      r['CompanyName'], r['ContactName'], r['ContactEmail'], r['ContactPhone']||'',
      r['MeetingDate'], r['Stage'], Number(r['DealValue']||0), r['CloseDate']||'',
      r['Products']||'', r['Notes']||'', Number(r['Rating']||0), r['RatingLabel']||'',
      r['Competitors']||'', r['DecisionMaker']||'', r['Checklist']||'', r['Extra']||''
    ]);
  });
  if (rows.length > 0) { sh.getRange(2,11,rows.length,1).setNumberFormat('$#,##0.00'); }
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}

function buildSummarySheet_(sh, data) {
  sh.getRange('A1').setValue('UICUSA Business Trip Suite — Summary').setFontWeight('bold').setFontSize(14).setFontColor('#1a3a6b').setFontFamily('Arial');
  sh.getRange('A2').setValue('Generated: ' + new Date().toLocaleString()).setFontColor('#888').setFontSize(10);

  var reqs = data.requests||[], exps = data.expenses||[], rpts = data.reports||[], sfu = data.sales||[];
  var kpis = [
    ['Total trip requests',       reqs.length],
    ['Approved trips',            reqs.filter(function(r){ return r['Status']==='APPROVED'; }).length],
    ['Pending approvals (trips)', reqs.filter(function(r){ return r['Status']==='PENDING_DEPT'||r['Status']==='PENDING_GM'; }).length],
    ['Rejected trips',            reqs.filter(function(r){ return r['Status']==='REJECTED'; }).length],
    ['Total expense reports',     exps.length],
    ['Approved expense total',    '$'+exps.filter(function(r){ return r['Status']==='APPROVED'; }).reduce(function(s,r){ return s+parseFloat(r['TotalAmount']||0);},0).toFixed(2)],
    ['Pending expenses',          exps.filter(function(r){ return r['Status']==='PENDING_DEPT'||r['Status']==='PENDING_GM'; }).length],
    ['Trip reports submitted',    rpts.length],
    ['Sales follow-ups logged',   sfu.length],
    ['Total deal pipeline ($)',   '$'+sfu.reduce(function(s,r){ return s+parseFloat(r['DealValue']||0);},0).toFixed(2)],
  ];

  sh.getRange(4,1).setValue('KEY METRICS').setFontWeight('bold').setFontColor('#1a3a6b').setFontFamily('Arial');
  kpis.forEach(function(k,i){
    sh.getRange(5+i,1).setValue(k[0]).setFontFamily('Arial').setFontSize(11);
    sh.getRange(5+i,2).setValue(k[1]).setFontWeight('bold').setFontFamily('Arial').setFontSize(11);
    if(i%2===0){ sh.getRange(5+i,1,1,2).setBackground('#f5f7fa'); }
  });

  // Dept breakdown
  var dRow = 5 + kpis.length + 2;
  sh.getRange(dRow,1).setValue('TRIPS BY DEPARTMENT').setFontWeight('bold').setFontColor('#1a3a6b');
  sh.getRange(dRow+1,1,1,3).setValues([['Department','Trips','Approved']]).setFontWeight('bold').setBackground('#1a3a6b').setFontColor('#fff');
  var deptMap = {};
  reqs.forEach(function(r){
    var d = r['Department']||'Unknown';
    if (!deptMap[d]) { deptMap[d]={total:0,approved:0}; }
    deptMap[d].total++;
    if (r['Status']==='APPROVED') { deptMap[d].approved++; }
  });
  Object.keys(deptMap).forEach(function(d,i){
    sh.getRange(dRow+2+i,1).setValue(d);
    sh.getRange(dRow+2+i,2).setValue(deptMap[d].total);
    sh.getRange(dRow+2+i,3).setValue(deptMap[d].approved);
    if(i%2===0){ sh.getRange(dRow+2+i,1,1,3).setBackground('#f5f7fa'); }
  });

  sh.setColumnWidth(1,240); sh.setColumnWidth(2,140); sh.setFrozenRows(1);
}

function colLetter_(col) {
  var s='';
  while(col>0){ var r=(col-1)%26; s=String.fromCharCode(65+r)+s; col=Math.floor((col-1)/26); }
  return s;
}
