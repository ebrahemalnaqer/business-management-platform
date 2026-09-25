// ============================================================
// TECHNO CLUB — Enterprise Co-working SaaS | Code.gs V3
// Google Apps Script Backend Engine — Production Ready
// ============================================================

var SHEET_NAMES = {
  CLIENTS: "Clients",
  ATTENDANCE: "Attendance",
  SERVICES: "Services",
  INVOICES: "Invoices",
  INVOICE_ITEMS: "InvoiceItems",
  EXPENSES: "Expenses",
  MANUAL_INFLOWS: "ManualInflows",
  SHIFTS: "Shifts",
  FEEDBACK: "Feedback"
};

var HEADERS = {
  Clients: ["ClientID","CustomerName","PhoneNumber","LeadSource","AccountType","HourlyBalance","CreatedAt"],
  Attendance: ["SessionID","ClientID","CustomerName","ResourceID","ResourceName","CheckInTime","CheckOutTime","DurationMinutes","TotalCharge","Status"],
  Services: ["ServiceID","ServiceName","Category","CostPrice","Price","CapacityOrStock","MinAlert","CreatedAt"],
  Invoices: ["InvoiceID","Date","CustomerName","ClientID","Subtotal","Discount","Tax","Total","PaymentMethod","B2B_TabStatus","HandledBy"],
  InvoiceItems: ["InvoiceID","ServiceID","ServiceName","Quantity","UnitPrice","LineTotal"],
  Expenses: ["ExpenseID","Date","Category","Description","Amount","LoggedBy","CreatedAt"],
  ManualInflows: ["InflowID","Date","Category","Description","Amount","LoggedBy","CreatedAt"],
  Shifts: ["ShiftID","Date","StaffName","CashExpected","CashReported","Variance","Status"],
  Feedback: ["FeedbackID","Timestamp","ClientName","Score","Message","Status"]
};

// ============================================================
// ENTRY POINT
// ============================================================

function doGet(e) {
  initializeAllSheets();
  var template = HtmlService.createTemplateFromFile("index");
  var html = template.evaluate()
    .setTitle("Techno Club — Co-working SaaS")
    .addMetaTag("viewport","width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// SHEET INITIALIZATION ENGINE
// ============================================================

function initializeAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var sheetName in HEADERS) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    var headers = HEADERS[sheetName];
    var firstRow = sheet.getRange(1,1,1,headers.length).getValues()[0];
    var isEmpty = firstRow.every(function(cell){ return cell === ""; });
    if (isEmpty) {
      sheet.getRange(1,1,1,headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1,1,1,headers.length)
        .setBackground("#12b981")
        .setFontColor("#f8fafc")
        .setFontWeight("bold");
    }
  }
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function generateID(prefix) {
  return prefix + "-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
}

// ============================================================
// CLIENTS MODULE — Phone as Primary UID + Deduplication Guard
// ============================================================

function getClients() {
  var sheet = getSheet(SHEET_NAMES.CLIENTS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h,i){ obj[h] = row[i]; });
    return obj;
  });
}

function addClient(params) {
  try {
    var sheet = getSheet(SHEET_NAMES.CLIENTS);
    var data = sheet.getDataRange().getValues();
    var cleanPhone = String(params.phoneNumber || params.clientId || "").trim();
    if (!cleanPhone) return { success:false, error:"Phone number is required as the primary unique ID." };
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === cleanPhone || String(data[i][2]).trim() === cleanPhone) {
        return { success:false, error:"This phone number is already registered in the CRM." };
      }
    }
    var now = new Date().toISOString();
    sheet.appendRow([
      cleanPhone,
      params.customerName,
      cleanPhone,
      params.leadSource || "Other Sources/أخرى",
      params.accountType || "Drop-in",
      parseFloat(params.hourlyBalance) || 0,
      now
    ]);
    return { success:true, clientId:cleanPhone };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateClient(params) {
  try {
    var sheet = getSheet(SHEET_NAMES.CLIENTS);
    var data = sheet.getDataRange().getValues();
    var targetID = String(params.clientId).trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetID) {
        sheet.getRange(i+1,2).setValue(params.customerName);
        sheet.getRange(i+1,4).setValue(params.leadSource);
        sheet.getRange(i+1,5).setValue(params.accountType);
        sheet.getRange(i+1,6).setValue(parseFloat(params.hourlyBalance)||0);
        return { success:true };
      }
    }
    return { success:false, error:"Client not found" };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteClient(clientId) {
  try {
    var sheet = getSheet(SHEET_NAMES.CLIENTS);
    var data = sheet.getDataRange().getValues();
    var targetID = String(clientId).trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetID) {
        sheet.deleteRow(i+1);
        return { success:true };
      }
    }
    return { success:false, error:"Client not found" };
  } catch(e) { return { success:false, error:e.message }; }
}

function topUpClientHours(params) {
  try {
    var sheet = getSheet(SHEET_NAMES.CLIENTS);
    var data = sheet.getDataRange().getValues();
    var targetID = String(params.clientId).trim();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetID) {
        var cur = parseFloat(data[i][5])||0;
        var add = parseFloat(params.hours)||0;
        sheet.getRange(i+1,6).setValue(cur+add);
        return { success:true, newBalance:cur+add };
      }
    }
    return { success:false, error:"Client not found" };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// ATTENDANCE MODULE
// ============================================================

function getActiveSessions() {
  var sheet = getSheet(SHEET_NAMES.ATTENDANCE);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).filter(function(row){ return row[9]==="Active"; }).map(function(row){
    var obj={}; headers.forEach(function(h,i){ obj[h]=row[i]; }); return obj;
  });
}

function getAllSessions() {
  var sheet = getSheet(SHEET_NAMES.ATTENDANCE);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row){
    var obj={}; headers.forEach(function(h,i){ obj[h]=row[i]; }); return obj;
  });
}

function executeCheckIn(params) {
  try {
    var sheet = getSheet(SHEET_NAMES.ATTENDANCE);
    var id = generateID("SES");
    var now = new Date();
    sheet.appendRow([id, String(params.clientId).trim(), params.customerName, params.resourceId, params.resourceName, now.toISOString(), "", 0, 0, "Active"]);
    return { success:true, sessionId:id, checkInTime:now.toISOString() };
  } catch(e) { return { success:false, error:e.message }; }
}

function executeCheckOut(params) {
  try {
    var attSheet = getSheet(SHEET_NAMES.ATTENDANCE);
    var data = attSheet.getDataRange().getValues();
    var sessionRow = -1, sessionData = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === params.sessionId && data[i][9] === "Active") {
        sessionRow = i+1; sessionData = data[i]; break;
      }
    }
    if (sessionRow === -1) return { success:false, error:"Active session not found" };

    var checkInTime = new Date(sessionData[5]);
    var checkOutTime = new Date();
    var durationMinutes = Math.round((checkOutTime - checkInTime) / 60000);
    var resourceId = sessionData[3];
    var resourceName = sessionData[4];
    var totalCharge = 0;
    var isShared = resourceName.toLowerCase().indexOf("shared") !== -1 || resourceName.toLowerCase().indexOf("workspace") !== -1;

    if (isShared && durationMinutes >= 240) {
      totalCharge = 200;
    } else {
      var svcSheet = getSheet(SHEET_NAMES.SERVICES);
      var svcData = svcSheet.getDataRange().getValues();
      var hourlyRate = 30;
      for (var j = 1; j < svcData.length; j++) {
        if (svcData[j][0] === resourceId) { hourlyRate = parseFloat(svcData[j][4])||30; break; }
      }
      totalCharge = Math.round((durationMinutes/60) * hourlyRate * 100) / 100;
      if (isShared && totalCharge > 200) totalCharge = 200;
    }

    var clientSheet = getSheet(SHEET_NAMES.CLIENTS);
    var clientData = clientSheet.getDataRange().getValues();
    var targetClientID = String(sessionData[1]).trim();
    for (var k = 1; k < clientData.length; k++) {
      if (String(clientData[k][0]).trim() === targetClientID && clientData[k][4] === "Hourly Package") {
        var bal = parseFloat(clientData[k][5])||0;
        clientSheet.getRange(k+1,6).setValue(Math.max(0, bal - durationMinutes/60));
        break;
      }
    }

    attSheet.getRange(sessionRow,7).setValue(checkOutTime.toISOString());
    attSheet.getRange(sessionRow,8).setValue(durationMinutes);
    attSheet.getRange(sessionRow,9).setValue(totalCharge);
    attSheet.getRange(sessionRow,10).setValue("Completed");

    return {
      success:true, sessionId:params.sessionId, clientId:targetClientID,
      customerName:sessionData[2], resourceId:resourceId, resourceName:resourceName,
      durationMinutes:durationMinutes, totalCharge:totalCharge
    };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// SERVICES MODULE
// ============================================================

function getServices() {
  var sheet = getSheet(SHEET_NAMES.SERVICES);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row){
    var obj={}; headers.forEach(function(h,i){ obj[h]=row[i]; }); return obj;
  });
}

function addService(params) {
  try {
    var sheet = getSheet(SHEET_NAMES.SERVICES);
    var id = generateID("SVC");
    sheet.appendRow([id, params.serviceName, params.category, parseFloat(params.costPrice)||0, parseFloat(params.price)||0, parseFloat(params.capacityOrStock)||0, parseFloat(params.minAlert)||0, new Date().toISOString()]);
    return { success:true, serviceId:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateService(params) {
  try {
    var sheet = getSheet(SHEET_NAMES.SERVICES);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === params.serviceId) {
        sheet.getRange(i+1,2).setValue(params.serviceName);
        sheet.getRange(i+1,3).setValue(params.category);
        sheet.getRange(i+1,4).setValue(parseFloat(params.costPrice)||0);
        sheet.getRange(i+1,5).setValue(parseFloat(params.price)||0);
        sheet.getRange(i+1,6).setValue(parseFloat(params.capacityOrStock)||0);
        sheet.getRange(i+1,7).setValue(parseFloat(params.minAlert)||0);
        return { success:true };
      }
    }
    return { success:false, error:"Service not found" };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteService(serviceId) {
  try {
    var sheet = getSheet(SHEET_NAMES.SERVICES);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === serviceId) { sheet.deleteRow(i+1); return { success:true }; }
    }
    return { success:false, error:"Not found" };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// INVOICES MODULE
// ============================================================

function getInvoices() {
  var sheet = getSheet(SHEET_NAMES.INVOICES);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row){
    var obj={}; headers.forEach(function(h,i){ obj[h]=row[i]; }); return obj;
  });
}

function getInvoiceItems(invoiceId) {
  var sheet = getSheet(SHEET_NAMES.INVOICE_ITEMS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).filter(function(row){ return row[0]===invoiceId; }).map(function(row){
    var obj={}; headers.forEach(function(h,i){ obj[h]=row[i]; }); return obj;
  });
}

function createInvoice(params) {
  try {
    var invSheet = getSheet(SHEET_NAMES.INVOICES);
    var itemSheet = getSheet(SHEET_NAMES.INVOICE_ITEMS);
    var id = generateID("INV");
    var now = new Date().toISOString();
    var items = params.items || [];
    var subtotal = items.reduce(function(s,i){ return s+(parseFloat(i.lineTotal)||0); }, 0);
    var discount = parseFloat(params.discount)||0;
    var tax = parseFloat(params.tax)||0;
    var total = subtotal - discount + tax;

    invSheet.appendRow([id, now, params.customerName, params.clientId||"", subtotal, discount, tax, total, params.paymentMethod, params.b2bTabStatus||"N/A", params.handledBy||"Staff"]);

    items.forEach(function(item){
      itemSheet.appendRow([id, item.serviceId, item.serviceName, parseFloat(item.quantity)||1, parseFloat(item.unitPrice)||0, parseFloat(item.lineTotal)||0]);
    });

    var svcSheet = getSheet(SHEET_NAMES.SERVICES);
    var svcData = svcSheet.getDataRange().getValues();
    items.forEach(function(item){
      if (item.category==="Product"||item.category==="Cafe") {
        for (var i = 1; i < svcData.length; i++) {
          if (svcData[i][0]===item.serviceId) {
            svcSheet.getRange(i+1,6).setValue(Math.max(0,(parseFloat(svcData[i][5])||0)-(parseFloat(item.quantity)||1)));
            break;
          }
        }
      }
    });

    return { success:true, invoiceId:id, total:total };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// EXPENSES MODULE
// ============================================================

function getExpenses() {
  var sheet = getSheet(SHEET_NAMES.EXPENSES);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row){
    var obj={}; headers.forEach(function(h,i){ obj[h]=row[i]; }); return obj;
  });
}

function addExpense(params) {
  try {
    var sheet = getSheet(SHEET_NAMES.EXPENSES);
    var id = generateID("EXP");
    var now = new Date().toISOString();
    sheet.appendRow([id, params.date||now.split("T")[0], params.category, params.description, parseFloat(params.amount)||0, params.loggedBy||"Admin", now]);
    return { success:true, expenseId:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteExpense(expenseId) {
  try {
    var sheet = getSheet(SHEET_NAMES.EXPENSES);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]===expenseId) { sheet.deleteRow(i+1); return { success:true }; }
    }
    return { success:false, error:"Not found" };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// MANUAL INFLOWS MODULE
// ============================================================

function getManualInflows() {
  var sheet = getSheet(SHEET_NAMES.MANUAL_INFLOWS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row){
    var obj={}; headers.forEach(function(h,i){ obj[h]=row[i]; }); return obj;
  });
}

function addManualInflow(params) {
  try {
    var sheet = getSheet(SHEET_NAMES.MANUAL_INFLOWS);
    var id = generateID("INF");
    var now = new Date().toISOString();
    sheet.appendRow([id, params.date||now.split("T")[0], params.category, params.description, parseFloat(params.amount)||0, params.loggedBy||"Admin", now]);
    return { success:true, inflowId:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteManualInflow(inflowId) {
  try {
    var sheet = getSheet(SHEET_NAMES.MANUAL_INFLOWS);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]===inflowId) { sheet.deleteRow(i+1); return { success:true }; }
    }
    return { success:false, error:"Not found" };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// SHIFTS MODULE — Server-side anti-fraud expected revenue calc
// ============================================================

function getShifts() {
  var sheet = getSheet(SHEET_NAMES.SHIFTS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row){
    var obj={}; headers.forEach(function(h,i){ obj[h]=row[i]; }); return obj;
  });
}

function calculateExpectedRevenueForDate(targetDate) {
  // FIX BUG 6: Normalize date string robustly
  if (!targetDate) return 0;
  var dateStr = String(targetDate).trim().split("T")[0];

  var invoices = getInvoices();
  var expenses = getExpenses();
  var inflows = getManualInflows();

  var totalPOS = 0, totalInflows = 0, totalExpenses = 0;

  invoices.forEach(function(inv) {
    var d = String(inv.Date||"").split("T")[0].trim();
    if (d === dateStr) totalPOS += parseFloat(inv.Total)||0;
  });
  inflows.forEach(function(inf) {
    var d = String(inf.Date||"").split("T")[0].trim();
    if (d === dateStr) totalInflows += parseFloat(inf.Amount)||0;
  });
  expenses.forEach(function(exp) {
    var d = String(exp.Date||exp.CreatedAt||"").split("T")[0].trim();
    if (d === dateStr) totalExpenses += parseFloat(exp.Amount)||0;
  });

  var result = totalPOS + totalInflows - totalExpenses;
  // Always return a plain Number — critical so frontend parses cleanly
  return Math.round(result * 100) / 100;
}

function addShift(params) {
  try {
    var sheet = getSheet(SHEET_NAMES.SHIFTS);
    var id = generateID("SHF");
    var dateStr = String(params.date||new Date().toISOString()).split("T")[0].trim();
    // Server re-computes expected to block any client-side manipulation
    var cashExpected = calculateExpectedRevenueForDate(dateStr);
    var cashReported = parseFloat(params.cashReported)||0;
    var variance = Math.round((cashReported - cashExpected) * 100) / 100;
    var status = variance===0?"Balanced":(variance>0?"Surplus":"Deficit");
    sheet.appendRow([id, dateStr, params.staffName, cashExpected, cashReported, variance, status]);
    return { success:true, shiftId:id, cashExpected:cashExpected, variance:variance, status:status };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// FEEDBACK MODULE
// ============================================================

function getFeedback() {
  var sheet = getSheet(SHEET_NAMES.FEEDBACK);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row){
    var obj={}; headers.forEach(function(h,i){ obj[h]=row[i]; }); return obj;
  });
}

function addFeedback(params) {
  try {
    var sheet = getSheet(SHEET_NAMES.FEEDBACK);
    var id = generateID("FBK");
    var now = new Date().toISOString();
    sheet.appendRow([id, now, params.clientName, parseInt(params.score)||5, params.message, "New"]);
    return { success:true, feedbackId:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateFeedbackStatus(params) {
  try {
    var sheet = getSheet(SHEET_NAMES.FEEDBACK);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]===params.feedbackId) { sheet.getRange(i+1,6).setValue(params.status); return { success:true }; }
    }
    return { success:false, error:"Not found" };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// FINANCIAL AGGREGATION ENGINE
// ============================================================

function getFinancialSummary() {
  try {
    var invoices = getInvoices();
    var expenses = getExpenses();
    var inflows = getManualInflows();

    var posRevenue = invoices.reduce(function(s,i){ return s+(parseFloat(i.Total)||0); },0);
    var manualInflowTotal = inflows.reduce(function(s,i){ return s+(parseFloat(i.Amount)||0); },0);
    var totalExpenses = expenses.reduce(function(s,e){ return s+(parseFloat(e.Amount)||0); },0);
    var grossIncome = posRevenue + manualInflowTotal;
    var netYield = grossIncome - totalExpenses;

    var now = new Date();
    var cm = now.getMonth(), cy = now.getFullYear();

    function sameMonth(d){ var x=new Date(d); return x.getMonth()===cm&&x.getFullYear()===cy; }
    var monthlyPOS = invoices.filter(function(i){ return sameMonth(i.Date); }).reduce(function(s,i){ return s+(parseFloat(i.Total)||0); },0);
    var monthlyExpenses = expenses.filter(function(e){ return sameMonth(e.CreatedAt||e.Date); }).reduce(function(s,e){ return s+(parseFloat(e.Amount)||0); },0);
    var monthlyInflows = inflows.filter(function(i){ return sameMonth(i.CreatedAt||i.Date); }).reduce(function(s,i){ return s+(parseFloat(i.Amount)||0); },0);

    var dailyRevenue = {};
    for (var d = 6; d >= 0; d--) {
      var day = new Date(); day.setDate(day.getDate()-d);
      dailyRevenue[day.toISOString().split("T")[0]] = 0;
    }
    invoices.forEach(function(inv){
      var k = String(inv.Date||"").split("T")[0];
      if (dailyRevenue.hasOwnProperty(k)) dailyRevenue[k] += parseFloat(inv.Total)||0;
    });

    var expenseByCategory = {};
    expenses.forEach(function(e){
      var c = e.Category||"Other";
      expenseByCategory[c] = (expenseByCategory[c]||0) + (parseFloat(e.Amount)||0);
    });

    return {
      success:true,
      posRevenue:Math.round(posRevenue*100)/100,
      manualInflowTotal:Math.round(manualInflowTotal*100)/100,
      grossIncome:Math.round(grossIncome*100)/100,
      totalExpenses:Math.round(totalExpenses*100)/100,
      netYield:Math.round(netYield*100)/100,
      monthlyPOS:Math.round(monthlyPOS*100)/100,
      monthlyExpenses:Math.round(monthlyExpenses*100)/100,
      monthlyInflows:Math.round(monthlyInflows*100)/100,
      dailyRevenue:dailyRevenue,
      expenseByCategory:expenseByCategory,
      invoiceCount:invoices.length,
      activeSessionCount:getActiveSessions().length
    };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// DASHBOARD KPI ENGINE
// ============================================================

function getDashboardKPIs() {
  try {
    var clients = getClients();
    var sessions = getAllSessions();
    var activeSessions = sessions.filter(function(s){ return s.Status==="Active"; });
    var financial = getFinancialSummary();
    var feedback = getFeedback();
    var services = getServices();

    var avgScore = 0;
    if (feedback.length > 0) {
      avgScore = Math.round(feedback.reduce(function(s,f){ return s+(parseInt(f.Score)||0); },0) / feedback.length * 10) / 10;
    }
    var lowStockAlerts = services.filter(function(s){ return parseFloat(s.CapacityOrStock)<=parseFloat(s.MinAlert)&&parseFloat(s.MinAlert)>0; });

    return {
      success:true,
      totalClients:clients.length,
      activeSessions:activeSessions.length,
      grossIncome:financial.grossIncome,
      netYield:financial.netYield,
      totalInvoices:getInvoices().length,
      avgFeedbackScore:avgScore,
      lowStockAlerts:lowStockAlerts.length,
      dailyRevenue:financial.dailyRevenue,
      expenseByCategory:financial.expenseByCategory,
      monthlyPOS:financial.monthlyPOS,
      monthlyExpenses:financial.monthlyExpenses,
      monthlyInflows:financial.monthlyInflows
    };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// DEMO DATA GENERATOR
// ============================================================

function createDemoData() {
  try {
    initializeAllSheets();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    [SHEET_NAMES.CLIENTS,SHEET_NAMES.ATTENDANCE,SHEET_NAMES.SERVICES,
     SHEET_NAMES.INVOICES,SHEET_NAMES.INVOICE_ITEMS,SHEET_NAMES.EXPENSES,
     SHEET_NAMES.MANUAL_INFLOWS,SHEET_NAMES.SHIFTS,SHEET_NAMES.FEEDBACK].forEach(function(name){
      var sheet = ss.getSheetByName(name);
      var lr = sheet.getLastRow();
      if (lr > 1) sheet.deleteRows(2, lr-1);
    });

    var svcSheet = getSheet(SHEET_NAMES.SERVICES);
    [["SVC-001","Shared Workspace","Workspace",0,30,20,2],
     ["SVC-002","Private Office","Workspace",0,60,5,1],
     ["SVC-003","Conference Room","Workspace",0,120,2,0],
     ["SVC-004","High-Speed Internet","Service",0,20,100,10],
     ["SVC-005","Printing B&W","Service",1,5,500,50],
     ["SVC-006","Printing Color","Service",2,10,200,20],
     ["SVC-007","Coffee","Cafe",5,15,50,5],
     ["SVC-008","Water Bottle","Cafe",3,8,100,10],
     ["SVC-009","AI Course - Beginner","Training",0,500,30,2],
     ["SVC-010","AI Course - Advanced","Training",0,900,20,2]
    ].forEach(function(r){ svcSheet.appendRow(r.concat([new Date().toISOString()])); });

    var cltSheet = getSheet(SHEET_NAMES.CLIENTS);
    var now = new Date().toISOString();
    [["01001234567","Ahmed Hassan","01001234567","Facebook Page/فيسبوك","Monthly Member",0,now],
     ["01112345678","Sara Mohamed","01112345678","Friends/أصحاب","Hourly Package",20,now],
     ["01223456789","Omar Khaled","01223456789","Mansoura University Campaign","Drop-in",0,now],
     ["01334567890","Nour Ibrahim","01334567890","Sponsored Ads/إعلان ممول","Monthly Member",0,now],
     ["01445678901","Kareem Ali","01445678901","Corporate Outflow/وفد شركة","Corporate",0,now]
    ].forEach(function(r){ cltSheet.appendRow(r); });

    var attSheet = getSheet(SHEET_NAMES.ATTENDANCE);
    var t1 = new Date(Date.now()-3*3600000).toISOString();
    var t2 = new Date(Date.now()-1*3600000).toISOString();
    attSheet.appendRow(["SES-D1","01001234567","Ahmed Hassan","SVC-001","Shared Workspace",new Date(Date.now()-5*3600000).toISOString(),new Date(Date.now()-2*3600000).toISOString(),180,90,"Completed"]);
    attSheet.appendRow(["SES-D2","01112345678","Sara Mohamed","SVC-002","Private Office",t1,"",0,0,"Active"]);

    var today = new Date().toISOString().split("T")[0];
    var invSheet = getSheet(SHEET_NAMES.INVOICES);
    var itemSheet = getSheet(SHEET_NAMES.INVOICE_ITEMS);
    invSheet.appendRow(["INV-D1",today+"T10:00:00.000Z","Ahmed Hassan","01001234567",500,0,0,500,"Cash","N/A","Staff"]);
    itemSheet.appendRow(["INV-D1","SVC-009","AI Course - Beginner",1,500,500]);
    invSheet.appendRow(["INV-D2",today+"T12:00:00.000Z","Sara Mohamed","01112345678",38,0,0,38,"Card","N/A","Staff"]);
    itemSheet.appendRow(["INV-D2","SVC-007","Coffee",2,15,30]);
    itemSheet.appendRow(["INV-D2","SVC-005","Printing B&W",2,5,10]);

    var expSheet = getSheet(SHEET_NAMES.EXPENSES);
    expSheet.appendRow(["EXP-D1",today,"Internet","Monthly Fiber — WE Telecom",450,"Admin",new Date().toISOString()]);
    expSheet.appendRow(["EXP-D2",today,"Utilities","Electricity Bill",620,"Admin",new Date().toISOString()]);

    var infSheet = getSheet(SHEET_NAMES.MANUAL_INFLOWS);
    infSheet.appendRow(["INF-D1",today,"AI Course Enrollment","AI Beginner Batch #3 — 5 Students",2500,"Admin",new Date().toISOString()]);
    infSheet.appendRow(["INF-D2",today,"Corporate Package","Mansoura Corp Monthly Office",3000,"Admin",new Date().toISOString()]);

    var fbkSheet = getSheet(SHEET_NAMES.FEEDBACK);
    fbkSheet.appendRow(["FBK-D1",new Date().toISOString(),"Ahmed Hassan",5,"Amazing space, perfect internet!","Reviewed"]);
    fbkSheet.appendRow(["FBK-D2",new Date().toISOString(),"Sara Mohamed",4,"Great vibes, needs more power outlets.","New"]);

    return { success:true, message:"Demo data loaded: 5 clients, 2 sessions, 2 invoices, 2 expenses, 2 inflows, 2 feedback records." };
  } catch(e) { return { success:false, error:e.message }; }
}

// ============================================================
// UNIVERSAL DISPATCHER
// ============================================================

function dispatch(action, params) {
  params = params || {};
  switch(action) {
    case "getClients":                    return getClients();
    case "addClient":                     return addClient(params);
    case "updateClient":                  return updateClient(params);
    case "deleteClient":                  return deleteClient(params.clientId);
    case "topUpClientHours":              return topUpClientHours(params);
    case "getActiveSessions":             return getActiveSessions();
    case "getAllSessions":                 return getAllSessions();
    case "executeCheckIn":                return executeCheckIn(params);
    case "executeCheckOut":               return executeCheckOut(params);
    case "getServices":                   return getServices();
    case "addService":                    return addService(params);
    case "updateService":                 return updateService(params);
    case "deleteService":                 return deleteService(params.serviceId);
    case "getInvoices":                   return getInvoices();
    case "getInvoiceItems":               return getInvoiceItems(params.invoiceId);
    case "createInvoice":                 return createInvoice(params);
    case "getExpenses":                   return getExpenses();
    case "addExpense":                    return addExpense(params);
    case "deleteExpense":                 return deleteExpense(params.expenseId);
    case "getManualInflows":              return getManualInflows();
    case "addManualInflow":               return addManualInflow(params);
    case "deleteManualInflow":            return deleteManualInflow(params.inflowId);
    case "getShifts":                     return getShifts();
    case "addShift":                      return addShift(params);
    case "calculateExpectedRevenueForDate": return calculateExpectedRevenueForDate(params.date);
    case "getFeedback":                   return getFeedback();
    case "addFeedback":                   return addFeedback(params);
    case "updateFeedbackStatus":          return updateFeedbackStatus(params);
    case "getFinancialSummary":           return getFinancialSummary();
    case "getDashboardKPIs":              return getDashboardKPIs();
    case "createDemoData":                return createDemoData();
    default:
      var v4Result = (typeof dispatchV4 === "function") ? dispatchV4(action, params) : null;
      if (v4Result !== null) return v4Result;
      return { success:false, error:"Unknown action: " + action };
  }
}

// ============================================================
// TECHNO CLUB V4 — ADD-ONLY MODULES
// Room Bookings + Google Calendar Sync + Daily Reports
// This section is appended to V3 without removing old features.
// ============================================================

// Extend old sheet names safely
SHEET_NAMES.ROOM_BOOKINGS = "RoomBookings";
SHEET_NAMES.ROOM_DAILY_STATS = "RoomDailyStats";
SHEET_NAMES.FOLLOW_UPS = "FollowUps";
SHEET_NAMES.SYSTEM_SETTINGS = "SystemSettings";

// Extend old headers safely
HEADERS.RoomBookings = [
  "BookingID","CustomerName","Phone","Room","Date","StartTime","EndTime","Hours",
  "HourlyRate","Subtotal","DiscountType","DiscountValue","DiscountReason","DiscountAmount",
  "FinalAmount","PaymentMethod","Status","CalendarEventID","CreatedBy","CreatedAt","UpdatedAt","Notes"
];

HEADERS.RoomDailyStats = [
  "Date","Room","BookingsCount","BookedHours","Revenue","Discounts","OccupancyRate","UpdatedAt"
];

HEADERS.FollowUps = [
  "ClientID","CustomerName","Phone","LastVisit","DaysAbsent","Status","Notes","UpdatedAt"
];

HEADERS.SystemSettings = [
  "Key","Value","Description","UpdatedAt"
];

var TECHNO_WORKING_HOURS = {
  normal: { startHour: 9, endHour: 24, workingHours: 15 },
  friday: { startHour: 13, endHour: 24, workingHours: 11 }
};

var ROOM_CALENDARS = {
  "ALPHA": {
    calendarId:"YOUR_A_ROOM_CALENDAR_ID@group.calendar.google.com",
    capacity: 22,
    hourlyRate: 140
  },
  "BETA": {
    calendarId: "YOUR_B_ROOM_CALENDAR_ID@group.calendar.google.com",
    capacity: 16,
    hourlyRate: 110
  },
  "GAMMA": {
    calendarId: "YOUR_C_ROOM_CALENDAR_ID@group.calendar.google.com",
    capacity: 10,
    hourlyRate: 90
  },
  "DELTA": {
    calendarId: "YOUR_D_ROOM_CALENDAR_ID@group.calendar.google.com",
    capacity: 5,
    hourlyRate: 60
  }
};

function initializeTechnoClubV4() {
  initializeAllSheets();
  ensureRoomServicesExist();
  return { success:true, message:"Techno Club V4 sheets and room services are ready." };
}

function getRoomConfigs() {
  var rooms = [];
  Object.keys(ROOM_CALENDARS).forEach(function(room){
    rooms.push({
      room: room,
      calendarId: ROOM_CALENDARS[room].calendarId,
      capacity: ROOM_CALENDARS[room].capacity,
      hourlyRate: ROOM_CALENDARS[room].hourlyRate
    });
  });
  return rooms;
}

function ensureRoomServicesExist() {
  var sheet = getSheet(SHEET_NAMES.SERVICES);
  if (!sheet) return { success:false, error:"Services sheet not found" };
  var data = sheet.getDataRange().getValues();
  var existingNames = {};
  for (var i = 1; i < data.length; i++) {
    existingNames[String(data[i][1] || "").toUpperCase()] = true;
  }
  Object.keys(ROOM_CALENDARS).forEach(function(room){
    var serviceName = room + " Room";
    if (!existingNames[serviceName.toUpperCase()]) {
      sheet.appendRow([
        "ROOM-" + room,
        serviceName,
        "Workspace",
        0,
        ROOM_CALENDARS[room].hourlyRate,
        ROOM_CALENDARS[room].capacity,
        0,
        new Date().toISOString()
      ]);
    }
  });
  return { success:true };
}

function normalizeRoomName(room) {
  return String(room || "").trim().toUpperCase();
}

function parseDateSafe(value) {
  if (value instanceof Date) return value;
  return new Date(value);
}

function formatDateKey(dateValue) {
  var d = parseDateSafe(dateValue);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function roundMoney(value) {
  return Math.round((parseFloat(value) || 0) * 100) / 100;
}

function calculateRoomBookingPrice(room, startTime, endTime, discountType, discountValue) {
  room = normalizeRoomName(room);
  if (!ROOM_CALENDARS[room]) return { success:false, error:"Invalid room name" };

  var start = parseDateSafe(startTime);
  var end = parseDateSafe(endTime);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return { success:false, error:"Invalid date/time" };
  if (end <= start) return { success:false, error:"End time must be after start time" };

  var hours = roundMoney((end - start) / 1000 / 60 / 60);
  var hourlyRate = ROOM_CALENDARS[room].hourlyRate;
  var subtotal = roundMoney(hours * hourlyRate);

  discountType = String(discountType || "none").toLowerCase();
  discountValue = parseFloat(discountValue) || 0;
  var discountAmount = 0;

  if (discountType === "percentage" || discountType === "percent") {
    if (discountValue < 0) discountValue = 0;
    if (discountValue > 100) discountValue = 100;
    discountAmount = subtotal * discountValue / 100;
  } else if (discountType === "fixed" || discountType === "amount") {
    discountAmount = discountValue;
  }

  discountAmount = roundMoney(Math.max(0, Math.min(discountAmount, subtotal)));
  var finalAmount = roundMoney(subtotal - discountAmount);

  return {
    success:true,
    room:room,
    hours:hours,
    hourlyRate:hourlyRate,
    subtotal:subtotal,
    discountType:discountType,
    discountValue:discountValue,
    discountAmount:discountAmount,
    finalAmount:finalAmount
  };
}

function getCalendarEventIdRaw(event) {
  try { return event.getId(); } catch(e) { return ""; }
}

function checkRoomAvailability(room, startTime, endTime, excludeCalendarEventId) {
  room = normalizeRoomName(room);
  if (!ROOM_CALENDARS[room]) return { success:false, error:"Invalid room name" };

  var calendar = CalendarApp.getCalendarById(ROOM_CALENDARS[room].calendarId);
  if (!calendar) return { success:false, error:"Calendar not found for room " + room };

  var start = parseDateSafe(startTime);
  var end = parseDateSafe(endTime);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return { success:false, error:"Invalid date/time" };
  if (end <= start) return { success:false, error:"End time must be after start time" };

  var events = calendar.getEvents(start, end);
  var conflicts = [];
  events.forEach(function(ev){
    var evId = getCalendarEventIdRaw(ev);
    if (excludeCalendarEventId && evId === excludeCalendarEventId) return;
    conflicts.push({
      id: evId,
      title: ev.getTitle(),
      start: ev.getStartTime().toISOString(),
      end: ev.getEndTime().toISOString()
    });
  });

  return {
    success:true,
    available: conflicts.length === 0,
    conflictCount: conflicts.length,
    conflicts: conflicts,
    message: conflicts.length === 0 ? "Room is available." : "Room is already booked in this time."
  };
}

function buildRoomBookingDescription(params, price) {
  return [
    "Techno Club Room Booking",
    "Customer: " + (params.customerName || ""),
    "Phone: " + (params.phone || ""),
    "Room: " + price.room,
    "Hours: " + price.hours,
    "Hourly Rate: " + price.hourlyRate + " EGP",
    "Subtotal: " + price.subtotal + " EGP",
    "Discount: " + price.discountAmount + " EGP",
    "Discount Reason: " + (params.discountReason || ""),
    "Final Amount: " + price.finalAmount + " EGP",
    "Payment: " + (params.paymentMethod || "Cash"),
    "Notes: " + (params.notes || "")
  ].join("\n");
}

function getRoomBookings() {
  initializeAllSheets();
  var sheet = getSheet(SHEET_NAMES.ROOM_BOOKINGS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row){
    var obj = {};
    headers.forEach(function(h,i){ obj[h] = row[i]; });
    return obj;
  });
}

function getRoomBookingsByDate(dateStr) {
  dateStr = String(dateStr || formatDateKey(new Date())).split("T")[0];
  return getRoomBookings().filter(function(b){
    return String(b.Date || "").split("T")[0] === dateStr && String(b.Status || "") !== "Cancelled";
  });
}

function createRoomBooking(params) {
  try {
    initializeAllSheets();
    params = params || {};
    var room = normalizeRoomName(params.room);
    if (!params.customerName) return { success:false, error:"Customer name is required" };
    if (!params.phone) return { success:false, error:"Phone is required" };
    if (!room || !ROOM_CALENDARS[room]) return { success:false, error:"Valid room is required" };

    var availability = checkRoomAvailability(room, params.startTime, params.endTime);
    if (!availability.success) return availability;
    if (!availability.available) return availability;

    var price = calculateRoomBookingPrice(room, params.startTime, params.endTime, params.discountType || "none", params.discountValue || 0);
    if (!price.success) return price;

    var calendar = CalendarApp.getCalendarById(ROOM_CALENDARS[room].calendarId);
    var title = room + " Booking - " + params.customerName;
    var event = calendar.createEvent(
      title,
      parseDateSafe(params.startTime),
      parseDateSafe(params.endTime),
      { description: buildRoomBookingDescription(params, price) }
    );

    var bookingId = generateID("BOOK");
    var now = new Date().toISOString();
    var sheet = getSheet(SHEET_NAMES.ROOM_BOOKINGS);
    sheet.appendRow([
      bookingId,
      params.customerName,
      params.phone,
      room,
      formatDateKey(params.startTime),
      parseDateSafe(params.startTime).toISOString(),
      parseDateSafe(params.endTime).toISOString(),
      price.hours,
      price.hourlyRate,
      price.subtotal,
      price.discountType,
      price.discountValue,
      params.discountReason || "",
      price.discountAmount,
      price.finalAmount,
      params.paymentMethod || "Cash",
      params.status || "Confirmed",
      event.getId(),
      params.createdBy || "Staff",
      now,
      now,
      params.notes || ""
    ]);

    upsertClientFromBooking(params.customerName, params.phone, params.leadSource || "Booking");
    updateRoomDailyStats(formatDateKey(params.startTime));

    return {
      success:true,
      bookingId:bookingId,
      calendarEventId:event.getId(),
      finalAmount:price.finalAmount,
      discountAmount:price.discountAmount
    };
  } catch(e) {
    return { success:false, error:e.message };
  }
}

function findRoomBookingRow(bookingId) {
  var sheet = getSheet(SHEET_NAMES.ROOM_BOOKINGS);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(bookingId)) {
      return { sheet:sheet, rowIndex:i+1, row:data[i], headers:data[0] };
    }
  }
  return null;
}

function updateRoomBooking(params) {
  try {
    initializeAllSheets();
    params = params || {};
    if (!params.bookingId) return { success:false, error:"Booking ID is required" };
    var found = findRoomBookingRow(params.bookingId);
    if (!found) return { success:false, error:"Booking not found" };

    var old = {};
    found.headers.forEach(function(h,i){ old[h] = found.row[i]; });

    var room = normalizeRoomName(params.room || old.Room);
    var startTime = params.startTime || old.StartTime;
    var endTime = params.endTime || old.EndTime;
    var existingEventId = old.CalendarEventID || "";

    var availability = checkRoomAvailability(room, startTime, endTime, existingEventId);
    if (!availability.success) return availability;
    if (!availability.available) return availability;

    var price = calculateRoomBookingPrice(room, startTime, endTime, params.discountType || old.DiscountType || "none", params.discountValue || old.DiscountValue || 0);
    if (!price.success) return price;

    var calendar = CalendarApp.getCalendarById(ROOM_CALENDARS[room].calendarId);
    var event = existingEventId ? calendar.getEventById(existingEventId) : null;
    if (!event) {
      event = calendar.createEvent(
        room + " Booking - " + (params.customerName || old.CustomerName),
        parseDateSafe(startTime),
        parseDateSafe(endTime),
        { description: buildRoomBookingDescription(params, price) }
      );
    } else {
      event.setTitle(room + " Booking - " + (params.customerName || old.CustomerName));
      event.setTime(parseDateSafe(startTime), parseDateSafe(endTime));
      event.setDescription(buildRoomBookingDescription(params, price));
    }

    var now = new Date().toISOString();
    found.sheet.getRange(found.rowIndex, 1, 1, HEADERS.RoomBookings.length).setValues([[
      params.bookingId,
      params.customerName || old.CustomerName,
      params.phone || old.Phone,
      room,
      formatDateKey(startTime),
      parseDateSafe(startTime).toISOString(),
      parseDateSafe(endTime).toISOString(),
      price.hours,
      price.hourlyRate,
      price.subtotal,
      price.discountType,
      price.discountValue,
      params.discountReason || old.DiscountReason || "",
      price.discountAmount,
      price.finalAmount,
      params.paymentMethod || old.PaymentMethod || "Cash",
      params.status || old.Status || "Confirmed",
      event.getId(),
      old.CreatedBy || params.createdBy || "Staff",
      old.CreatedAt || now,
      now,
      params.notes || old.Notes || ""
    ]]);

    updateRoomDailyStats(formatDateKey(startTime));
    if (old.Date && old.Date !== formatDateKey(startTime)) updateRoomDailyStats(old.Date);

    return { success:true, bookingId:params.bookingId, finalAmount:price.finalAmount };
  } catch(e) {
    return { success:false, error:e.message };
  }
}

function cancelRoomBooking(params) {
  try {
    params = params || {};
    if (!params.bookingId) return { success:false, error:"Booking ID is required" };
    var found = findRoomBookingRow(params.bookingId);
    if (!found) return { success:false, error:"Booking not found" };
    var old = {};
    found.headers.forEach(function(h,i){ old[h] = found.row[i]; });

    if (old.CalendarEventID && ROOM_CALENDARS[normalizeRoomName(old.Room)]) {
      var calendar = CalendarApp.getCalendarById(ROOM_CALENDARS[normalizeRoomName(old.Room)].calendarId);
      var event = calendar.getEventById(old.CalendarEventID);
      if (event) event.deleteEvent();
    }

    found.sheet.getRange(found.rowIndex, 17).setValue("Cancelled");
    found.sheet.getRange(found.rowIndex, 21).setValue(new Date().toISOString());
    updateRoomDailyStats(old.Date);
    return { success:true };
  } catch(e) {
    return { success:false, error:e.message };
  }
}

function getWorkingHoursForDate(dateStr) {
  var d = parseDateSafe(dateStr + "T12:00:00");
  var day = d.getDay(); // Friday = 5
  return day === 5 ? TECHNO_WORKING_HOURS.friday.workingHours : TECHNO_WORKING_HOURS.normal.workingHours;
}

function getDailyRoomReport(dateStr) {
  initializeAllSheets();
  dateStr = String(dateStr || formatDateKey(new Date())).split("T")[0];
  var bookings = getRoomBookingsByDate(dateStr);
  var rooms = {};
  Object.keys(ROOM_CALENDARS).forEach(function(room){
    rooms[room] = {
      room:room,
      capacity:ROOM_CALENDARS[room].capacity,
      hourlyRate:ROOM_CALENDARS[room].hourlyRate,
      bookingsCount:0,
      bookedHours:0,
      revenue:0,
      discounts:0,
      occupancyRate:0
    };
  });

  bookings.forEach(function(b){
    var room = normalizeRoomName(b.Room);
    if (!rooms[room]) return;
    rooms[room].bookingsCount += 1;
    rooms[room].bookedHours += parseFloat(b.Hours) || 0;
    rooms[room].revenue += parseFloat(b.FinalAmount) || 0;
    rooms[room].discounts += parseFloat(b.DiscountAmount) || 0;
  });

  var workingHours = getWorkingHoursForDate(dateStr);
  var totalRoomRevenue = 0, totalRoomHours = 0, totalDiscounts = 0, totalBookings = 0;
  Object.keys(rooms).forEach(function(room){
    rooms[room].bookedHours = roundMoney(rooms[room].bookedHours);
    rooms[room].revenue = roundMoney(rooms[room].revenue);
    rooms[room].discounts = roundMoney(rooms[room].discounts);
    rooms[room].occupancyRate = roundMoney((rooms[room].bookedHours / workingHours) * 100);
    totalRoomRevenue += rooms[room].revenue;
    totalRoomHours += rooms[room].bookedHours;
    totalDiscounts += rooms[room].discounts;
    totalBookings += rooms[room].bookingsCount;
  });

  var financial = getFinancialSummaryForDateV4(dateStr);

  return {
    success:true,
    date:dateStr,
    workingHours:workingHours,
    rooms:Object.keys(rooms).map(function(k){ return rooms[k]; }),
    summary:{
      totalBookings:totalBookings,
      totalRoomHours:roundMoney(totalRoomHours),
      totalRoomRevenue:roundMoney(totalRoomRevenue),
      totalDiscounts:roundMoney(totalDiscounts),
      posRevenue:financial.posRevenue,
      manualInflows:financial.manualInflows,
      expenses:financial.expenses,
      grossIncome:financial.grossIncome,
      netProfit:financial.netProfit
    }
  };
}

function updateRoomDailyStats(dateStr) {
  dateStr = String(dateStr || formatDateKey(new Date())).split("T")[0];
  var report = getDailyRoomReport(dateStr);
  if (!report.success) return report;
  var sheet = getSheet(SHEET_NAMES.ROOM_DAILY_STATS);
  var data = sheet.getDataRange().getValues();
  var now = new Date().toISOString();

  report.rooms.forEach(function(r){
    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === dateStr && String(data[i][1]) === r.room) {
        foundRow = i + 1;
        break;
      }
    }
    var row = [dateStr, r.room, r.bookingsCount, r.bookedHours, r.revenue, r.discounts, r.occupancyRate, now];
    if (foundRow > -1) sheet.getRange(foundRow, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);
  });

  return { success:true, date:dateStr };
}

function getFinancialSummaryForDateV4(dateStr) {
  dateStr = String(dateStr || formatDateKey(new Date())).split("T")[0];
  var invoices = getInvoices();
  var expenses = getExpenses();
  var inflows = getManualInflows();
  var bookings = getRoomBookingsByDate(dateStr);

  var posRevenue = 0, totalExpenses = 0, manualInflows = 0, roomRevenue = 0, roomDiscounts = 0;

  invoices.forEach(function(inv){
    var d = String(inv.Date || "").split("T")[0];
    if (d === dateStr) posRevenue += parseFloat(inv.Total) || 0;
  });
  inflows.forEach(function(inf){
    var d = String(inf.Date || "").split("T")[0];
    if (d === dateStr) manualInflows += parseFloat(inf.Amount) || 0;
  });
  expenses.forEach(function(exp){
    var d = String(exp.Date || exp.CreatedAt || "").split("T")[0];
    if (d === dateStr) totalExpenses += parseFloat(exp.Amount) || 0;
  });
  bookings.forEach(function(b){
    roomRevenue += parseFloat(b.FinalAmount) || 0;
    roomDiscounts += parseFloat(b.DiscountAmount) || 0;
  });

  var grossIncome = posRevenue + manualInflows + roomRevenue;
  var netProfit = grossIncome - totalExpenses;
  return {
    date:dateStr,
    posRevenue:roundMoney(posRevenue),
    manualInflows:roundMoney(manualInflows),
    roomRevenue:roundMoney(roomRevenue),
    roomDiscounts:roundMoney(roomDiscounts),
    expenses:roundMoney(totalExpenses),
    grossIncome:roundMoney(grossIncome),
    netProfit:roundMoney(netProfit)
  };
}

// Override V3 function to include Room Bookings in expected cash calculation.
function calculateExpectedRevenueForDate(targetDate) {
  if (!targetDate) return 0;
  var dateStr = String(targetDate).trim().split("T")[0];
  var f = getFinancialSummaryForDateV4(dateStr);
  return roundMoney(f.grossIncome - f.expenses);
}

function getOwnerDashboardData(dateStr) {
  dateStr = String(dateStr || formatDateKey(new Date())).split("T")[0];
  var daily = getDailyRoomReport(dateStr);
  var monthStart = dateStr.substring(0, 7) + "-01";
  var month = getRoomRevenueByRange(monthStart, dateStr);
  return {
    success:true,
    today:daily,
    month:month,
    followUps:getFollowUps()
  };
}

function getRoomRevenueByRange(startDate, endDate) {
  startDate = String(startDate).split("T")[0];
  endDate = String(endDate).split("T")[0];
  var bookings = getRoomBookings().filter(function(b){
    var d = String(b.Date || "").split("T")[0];
    return d >= startDate && d <= endDate && String(b.Status || "") !== "Cancelled";
  });

  var rooms = {};
  Object.keys(ROOM_CALENDARS).forEach(function(room){
    rooms[room] = { room:room, bookingsCount:0, bookedHours:0, revenue:0, discounts:0 };
  });

  bookings.forEach(function(b){
    var room = normalizeRoomName(b.Room);
    if (!rooms[room]) return;
    rooms[room].bookingsCount += 1;
    rooms[room].bookedHours += parseFloat(b.Hours) || 0;
    rooms[room].revenue += parseFloat(b.FinalAmount) || 0;
    rooms[room].discounts += parseFloat(b.DiscountAmount) || 0;
  });

  return Object.keys(rooms).map(function(room){
    rooms[room].bookedHours = roundMoney(rooms[room].bookedHours);
    rooms[room].revenue = roundMoney(rooms[room].revenue);
    rooms[room].discounts = roundMoney(rooms[room].discounts);
    return rooms[room];
  });
}

function upsertClientFromBooking(customerName, phone, leadSource) {
  if (!phone) return;
  var clients = getClients();
  var cleanPhone = String(phone).trim();
  var exists = clients.some(function(c){ return String(c.ClientID).trim() === cleanPhone || String(c.PhoneNumber).trim() === cleanPhone; });
  if (!exists) {
    addClient({ customerName:customerName, phoneNumber:cleanPhone, leadSource:leadSource || "Booking", accountType:"Drop-in", hourlyBalance:0 });
  }
}

function getFollowUps() {
  initializeAllSheets();
  generateFollowUps();
  var sheet = getSheet(SHEET_NAMES.FOLLOW_UPS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  return data.slice(1).map(function(row){
    var obj = {};
    headers.forEach(function(h,i){ obj[h] = row[i]; });
    return obj;
  });
}

function generateFollowUps() {
  initializeAllSheets();
  var clients = getClients();
  var bookings = getRoomBookings();
  var sessions = getAllSessions();
  var today = new Date();
  var followRows = [];
  var now = new Date().toISOString();

  clients.forEach(function(c){
    var clientId = String(c.ClientID || c.PhoneNumber || "").trim();
    if (!clientId) return;
    var lastVisit = null;

    bookings.forEach(function(b){
      if (String(b.Phone || "").trim() === clientId && String(b.Status || "") !== "Cancelled") {
        var d = parseDateSafe(b.Date + "T12:00:00");
        if (!lastVisit || d > lastVisit) lastVisit = d;
      }
    });
    sessions.forEach(function(s){
      if (String(s.ClientID || "").trim() === clientId) {
        var d = parseDateSafe(s.CheckInTime);
        if (!lastVisit || d > lastVisit) lastVisit = d;
      }
    });

    if (!lastVisit && c.CreatedAt) lastVisit = parseDateSafe(c.CreatedAt);
    if (!lastVisit || isNaN(lastVisit.getTime())) return;

    var daysAbsent = Math.floor((today - lastVisit) / 86400000);
    var status = "Active";
    if (daysAbsent >= 60) status = "Lost Customer";
    else if (daysAbsent >= 30) status = "Needs Follow-up";
    else if (daysAbsent >= 14) status = "Warm Follow-up";

    if (daysAbsent >= 14) {
      followRows.push([
        clientId,
        c.CustomerName || "",
        c.PhoneNumber || clientId,
        formatDateKey(lastVisit),
        daysAbsent,
        status,
        "Auto-generated by V4 Follow-up Engine",
        now
      ]);
    }
  });

  var sheet = getSheet(SHEET_NAMES.FOLLOW_UPS);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).clearContent();
  if (followRows.length) sheet.getRange(2, 1, followRows.length, followRows[0].length).setValues(followRows);
  return { success:true, count:followRows.length };
}

function normalizePhoneForWhatsapp(phone) {
  var p = String(phone || "").replace(/[^0-9]/g, "");
  if (p.indexOf("20") === 0) return p;
  if (p.indexOf("0") === 0) return "20" + p.substring(1);
  if (p.length === 10) return "20" + p;
  return p;
}

function getWhatsappBookingMessage(params) {
  params = params || {};
  var room = normalizeRoomName(params.room);
  var start = parseDateSafe(params.startTime);
  var end = parseDateSafe(params.endTime);
  var dateText = Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var startText = Utilities.formatDate(start, Session.getScriptTimeZone(), "hh:mm a");
  var endText = Utilities.formatDate(end, Session.getScriptTimeZone(), "hh:mm a");
  return "تم تأكيد حجزك في Techno Club ✅\n\n" +
    "الاسم: " + (params.customerName || "") + "\n" +
    "القاعة: " + room + "\n" +
    "التاريخ: " + dateText + "\n" +
    "المعاد: من " + startText + " إلى " + endText + "\n" +
    "العنوان: المنصورة - حي الجامعة - تشكا - شارع استون يارد - فوق هايبر أبو ليلة - الدور الثالث\n" +
    "Google Maps: https://maps.app.goo.gl/u4yWHj5ECFoDzubZ8?g_st=ic\n\n" +
    "مستنينك تنورنا 🌿";
}

function getWhatsappReminderMessage(params) {
  params = params || {};
  var room = normalizeRoomName(params.room);
  var start = parseDateSafe(params.startTime);
  var dateText = Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var startText = Utilities.formatDate(start, Session.getScriptTimeZone(), "hh:mm a");
  return "تذكير بحجزك في Techno Club ⏰\n\n" +
    "القاعة: " + room + "\n" +
    "التاريخ: " + dateText + "\n" +
    "المعاد: " + startText + "\n\n" +
    "لو محتاج أي تعديل ابعتلنا على نفس الرقم.";
}

function getWhatsappBookingLink(params) {
  params = params || {};
  var msgType = params.type || "confirmation";
  var message = msgType === "reminder" ? getWhatsappReminderMessage(params) : getWhatsappBookingMessage(params);
  var phone = normalizePhoneForWhatsapp(params.phone);
  return { success:true, phone:phone, message:message, url:"https://wa.me/" + phone + "?text=" + encodeURIComponent(message) };
}

function getBookingById(bookingId) {
  var found = findRoomBookingRow(bookingId);
  if (!found) return null;
  var obj = {};
  found.headers.forEach(function(h,i){ obj[h] = found.row[i]; });
  return obj;
}

function getWhatsappLinkForBooking(params) {
  params = params || {};
  var booking = getBookingById(params.bookingId);
  if (!booking) return { success:false, error:"Booking not found" };
  return getWhatsappBookingLink({
    customerName: booking.CustomerName,
    phone: booking.Phone,
    room: booking.Room,
    startTime: booking.StartTime,
    endTime: booking.EndTime,
    type: params.type || "confirmation"
  });
}

function dispatchV4(action, params) {
  params = params || {};
  switch(action) {
    case "initializeTechnoClubV4": return initializeTechnoClubV4();
    case "getRoomConfigs": return getRoomConfigs();
    case "ensureRoomServicesExist": return ensureRoomServicesExist();
    case "calculateRoomBookingPrice": return calculateRoomBookingPrice(params.room, params.startTime, params.endTime, params.discountType, params.discountValue);
    case "checkRoomAvailability": return checkRoomAvailability(params.room, params.startTime, params.endTime, params.excludeCalendarEventId);
    case "createRoomBooking": return createRoomBooking(params);
    case "updateRoomBooking": return updateRoomBooking(params);
    case "cancelRoomBooking": return cancelRoomBooking(params);
    case "getRoomBookings": return getRoomBookings();
    case "getRoomBookingsByDate": return getRoomBookingsByDate(params.date);
    case "getDailyRoomReport": return getDailyRoomReport(params.date);
    case "updateRoomDailyStats": return updateRoomDailyStats(params.date);
    case "getOwnerDashboardData": return getOwnerDashboardData(params.date);
    case "getRoomRevenueByRange": return getRoomRevenueByRange(params.startDate, params.endDate);
    case "generateFollowUps": return generateFollowUps();
    case "getFollowUps": return getFollowUps();
    case "getWhatsappBookingLink": return getWhatsappBookingLink(params);
    case "getWhatsappLinkForBooking": return getWhatsappLinkForBooking(params);
    default: return null;
  }
}


function testRoomCalendars() {
  var ids = {
    ALPHA: "YOUR_A_ROOM_CALENDAR_ID@group.calendar.google.com",
    BETA: "YOUR_B_ROOM_CALENDAR_ID@group.calendar.google.com",
    GAMMA: "YOUR_C_ROOM_CALENDAR_ID@group.calendar.google.com",
    DELTA: "YOUR_D_ROOM_CALENDAR_ID@group.calendar.google.com"
  };

  Object.keys(ids).forEach(function(room) {
    var cal = CalendarApp.getCalendarById(ids[room]);
    Logger.log(room + " = " + (cal ? "FOUND: " + cal.getName() : "NOT FOUND"));
  });
}
function testGetRoomBookings() {
  var rows = getRoomBookings();
  Logger.log(JSON.stringify(rows, null, 2));
  return rows;
}




// ============================================================
// TECHNO CLUB — STRONG BOOKING CONFLICT ENGINE
// Works with Web App RoomBookings + old Techno Booking sheet
// Paste this block at the END of Code.gs
// ============================================================

var TECHNO_BOOKING_CALENDARS = {
  "ALPHA": "YOUR_A_ROOM_CALENDAR_ID@group.calendar.google.com",
  "BETA": "YOUR_B_ROOM_CALENDAR_ID@group.calendar.google.com",
  "GAMMA": "YOUR_C_ROOM_CALENDAR_ID@group.calendar.google.com",
  "DELTA": "YOUR_D_ROOM_CALENDAR_ID@group.calendar.google.com"
};

var TECHNO_ROOM_PRICES = {
  "ALPHA": 140,
  "BETA": 110,
  "GAMMA": 90,
  "DELTA": 60
};

var TECHNO_ROOM_CAPACITY = {
  "ALPHA": 22,
  "BETA": 16,
  "GAMMA": 10,
  "DELTA": 5
};

function technoNormalizeRoom(room) {
  var r = String(room || "").trim().toUpperCase();

  if (r === "ALPHA ROOM" || r === "TECHNOCLUB – ALPHA ROOM" || r === "TECHNOCLUB - ALPHA ROOM") return "ALPHA";
  if (r === "BETA ROOM" || r === "TECHNOCLUB – BETA ROOM" || r === "TECHNOCLUB - BETA ROOM") return "BETA";
  if (r === "GAMMA ROOM" || r === "TECHNOCLUB – GAMMA ROOM" || r === "TECHNOCLUB - GAMMA ROOM") return "GAMMA";
  if (r === "DELTA ROOM" || r === "TECHNOCLUB – DELTA ROOM" || r === "TECHNOCLUB - DELTA ROOM") return "DELTA";

  return r;
}

function technoParseDate(value) {
  if (value instanceof Date) return value;
  return new Date(value);
}

function technoFormatDateKey(value) {
  var d = technoParseDate(value);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function technoRound(value) {
  return Math.round((parseFloat(value) || 0) * 100) / 100;
}

function technoBuildDateTime(dateValue, timeValue) {
  var d = new Date(dateValue);

  if (timeValue instanceof Date) {
    d.setHours(timeValue.getHours(), timeValue.getMinutes(), 0, 0);
    return d;
  }

  var timeText = String(timeValue || "").trim();

  if (timeText.indexOf("T") !== -1) {
    var full = new Date(timeText);
    d.setHours(full.getHours(), full.getMinutes(), 0, 0);
    return d;
  }

  var parts = timeText.split(":");
  var h = parseInt(parts[0], 10) || 0;
  var m = parseInt(parts[1], 10) || 0;

  d.setHours(h, m, 0, 0);
  return d;
}

function technoGetCalendar(room) {
  room = technoNormalizeRoom(room);
  var calendarId = TECHNO_BOOKING_CALENDARS[room];

  if (!calendarId) return null;

  return CalendarApp.getCalendarById(calendarId);
}

function technoCheckSheetConflict(room, start, end, excludeBookingId) {
  var conflicts = [];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RoomBookings");

  if (!sheet || sheet.getLastRow() <= 1) {
    return conflicts;
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idxBookingID = headers.indexOf("BookingID");
  var idxCustomer = headers.indexOf("CustomerName");
  var idxRoom = headers.indexOf("Room");
  var idxStart = headers.indexOf("StartTime");
  var idxEnd = headers.indexOf("EndTime");
  var idxStatus = headers.indexOf("Status");

  if (idxRoom === -1 || idxStart === -1 || idxEnd === -1) {
    return conflicts;
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    var bookingId = String(row[idxBookingID] || "");
    var rowRoom = technoNormalizeRoom(row[idxRoom]);
    var status = String(row[idxStatus] || "");

    if (excludeBookingId && bookingId === excludeBookingId) continue;
    if (rowRoom !== room) continue;
    if (status === "Cancelled" || status === "ملغي") continue;

    var rowStart = technoParseDate(row[idxStart]);
    var rowEnd = technoParseDate(row[idxEnd]);

    if (isNaN(rowStart.getTime()) || isNaN(rowEnd.getTime())) continue;

    // True overlap condition
    var overlaps = start < rowEnd && end > rowStart;

    if (overlaps) {
      conflicts.push({
        source: "RoomBookings",
        bookingId: bookingId,
        customerName: row[idxCustomer],
        start: rowStart.toISOString(),
        end: rowEnd.toISOString()
      });
    }
  }

  return conflicts;
}

function checkRoomAvailability(room, startTime, endTime, excludeCalendarEventId, excludeBookingId) {
  room = technoNormalizeRoom(room);

  if (!TECHNO_BOOKING_CALENDARS[room]) {
    return {
      success: false,
      available: false,
      error: "Invalid room name: " + room
    };
  }

  var start = technoParseDate(startTime);
  var end = technoParseDate(endTime);

  if (isNaN(start.getTime()) || isNaN(end.getTime())) {
    return {
      success: false,
      available: false,
      error: "Invalid date/time"
    };
  }

  if (end <= start) {
    return {
      success: false,
      available: false,
      error: "End time must be after start time"
    };
  }

  var calendar = technoGetCalendar(room);

  if (!calendar) {
    return {
      success: false,
      available: false,
      error: "Calendar not found for room " + room
    };
  }

  var conflicts = [];

  // 1) Check Google Calendar
  var events = calendar.getEvents(start, end);

  events.forEach(function(event) {
    var eventId = "";

    try {
      eventId = event.getId();
    } catch (e) {}

    if (excludeCalendarEventId && eventId === excludeCalendarEventId) return;

    conflicts.push({
      source: "Google Calendar",
      id: eventId,
      title: event.getTitle(),
      start: event.getStartTime().toISOString(),
      end: event.getEndTime().toISOString()
    });
  });

  // 2) Check RoomBookings Sheet
  var sheetConflicts = technoCheckSheetConflict(room, start, end, excludeBookingId);

  sheetConflicts.forEach(function(c) {
    conflicts.push(c);
  });

  if (conflicts.length > 0) {
    return {
      success: true,
      available: false,
      conflictCount: conflicts.length,
      conflicts: conflicts,
      message: "محجوز بالفعل في نفس الميعاد"
    };
  }

  return {
    success: true,
    available: true,
    conflictCount: 0,
    conflicts: [],
    message: "Room is available"
  };
}

function technoCalculateBookingPrice(room, startTime, endTime, discountType, discountValue) {
  room = technoNormalizeRoom(room);

  var start = technoParseDate(startTime);
  var end = technoParseDate(endTime);

  if (end <= start) {
    return {
      success: false,
      error: "End time must be after start time"
    };
  }

  var hours = technoRound((end - start) / 1000 / 60 / 60);
  var hourlyRate = TECHNO_ROOM_PRICES[room] || 0;
  var subtotal = technoRound(hours * hourlyRate);

  discountType = String(discountType || "none").toLowerCase();
  discountValue = parseFloat(discountValue) || 0;

  var discountAmount = 0;

  if (discountType === "percentage" || discountType === "percent") {
    if (discountValue < 0) discountValue = 0;
    if (discountValue > 100) discountValue = 100;
    discountAmount = subtotal * discountValue / 100;
  } else if (discountType === "fixed" || discountType === "amount") {
    discountAmount = discountValue;
  }

  discountAmount = technoRound(Math.max(0, Math.min(discountAmount, subtotal)));

  return {
    success: true,
    room: room,
    hours: hours,
    hourlyRate: hourlyRate,
    subtotal: subtotal,
    discountType: discountType,
    discountValue: discountValue,
    discountAmount: discountAmount,
    finalAmount: technoRound(subtotal - discountAmount)
  };
}

function technoBookingDescription(params, price, bookingId) {
  return [
    "Techno Club Room Booking",
    "BookingID: " + bookingId,
    "Customer: " + (params.customerName || ""),
    "Phone: " + (params.phone || ""),
    "Room: " + price.room,
    "Hours: " + price.hours,
    "Hourly Rate: " + price.hourlyRate + " EGP",
    "Subtotal: " + price.subtotal + " EGP",
    "Discount: " + price.discountAmount + " EGP",
    "Discount Reason: " + (params.discountReason || ""),
    "Final Amount: " + price.finalAmount + " EGP",
    "Payment: " + (params.paymentMethod || "Cash"),
    "Created By: " + (params.createdBy || "Staff"),
    "Notes: " + (params.notes || "")
  ].join("\n");
}

function createRoomBooking(params) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    if (typeof initializeAllSheets === "function") {
      initializeAllSheets();
    }

    params = params || {};

    var room = technoNormalizeRoom(params.room);

    if (!params.customerName) {
      return { success: false, error: "Customer name is required" };
    }

    if (!params.phone) {
      return { success: false, error: "Phone is required" };
    }

    if (!TECHNO_BOOKING_CALENDARS[room]) {
      return { success: false, error: "Valid room is required" };
    }

    var start = technoParseDate(params.startTime);
    var end = technoParseDate(params.endTime);

    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
      return { success: false, error: "Invalid date/time" };
    }

    if (end <= start) {
      return { success: false, error: "End time must be after start time" };
    }

    // Strong conflict check before any calendar event or sheet write
    var availability = checkRoomAvailability(room, start, end);

    if (!availability.success) {
      return availability;
    }

    if (!availability.available) {
      return {
        success: false,
        error: "محجوز بالفعل في نفس الميعاد",
        details: availability
      };
    }

    var price = technoCalculateBookingPrice(
      room,
      start,
      end,
      params.discountType || "none",
      params.discountValue || 0
    );

    if (!price.success) return price;

    var calendar = technoGetCalendar(room);

    if (!calendar) {
      return { success: false, error: "Calendar not found for room " + room };
    }

    var bookingId = generateID("BOOK");
    var now = new Date().toISOString();

    var title = "حجز: " + params.customerName + " - " + room + " - " + bookingId;

    var event = calendar.createEvent(
      title,
      start,
      end,
      {
        description: technoBookingDescription(params, price, bookingId)
      }
    );

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("RoomBookings");

    if (!sheet) {
      sheet = ss.insertSheet("RoomBookings");
      sheet.appendRow([
        "BookingID","CustomerName","Phone","Room","Date","StartTime","EndTime","Hours",
        "HourlyRate","Subtotal","DiscountType","DiscountValue","DiscountReason","DiscountAmount",
        "FinalAmount","PaymentMethod","Status","CalendarEventID","CreatedBy","CreatedAt","UpdatedAt","Notes"
      ]);
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      bookingId,
      params.customerName,
      params.phone,
      room,
      technoFormatDateKey(start),
      start.toISOString(),
      end.toISOString(),
      price.hours,
      price.hourlyRate,
      price.subtotal,
      price.discountType,
      price.discountValue,
      params.discountReason || "",
      price.discountAmount,
      price.finalAmount,
      params.paymentMethod || "Cash",
      params.status || "Confirmed",
      event.getId(),
      params.createdBy || "Staff",
      now,
      now,
      params.notes || ""
    ]);

    if (typeof upsertClientFromBooking === "function") {
      upsertClientFromBooking(
        params.customerName,
        params.phone,
        params.leadSource || "Booking"
      );
    }

    if (typeof updateRoomDailyStats === "function") {
      updateRoomDailyStats(technoFormatDateKey(start));
    }

    return {
      success: true,
      bookingId: bookingId,
      calendarEventId: event.getId(),
      finalAmount: price.finalAmount,
      discountAmount: price.discountAmount,
      message: "Booking created successfully"
    };

  } catch (e) {
    return {
      success: false,
      error: e.message
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (err) {}
  }
}

// ============================================================
// OLD SHEET SYSTEM SUPPORT — Techno Booking onEdit
// This keeps your old manual sheet booking system working too.
// ============================================================

function setupTechnoBookingSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Techno Booking");

  if (!sheet) {
    sheet = ss.insertSheet("Techno Booking");
  }

  var headers = [
    "BookingID",
    "ClientName",
    "Phone",
    "Room",
    "Date",
    "StartTime",
    "EndTime",
    "Price",
    "Status",
    "Delete?",
    "Notes"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return {
    success: true,
    message: "Techno Booking sheet is ready"
  };
}

function onEdit(e) {
  try {
    if (!e || !e.source || !e.range) return;

    var sheet = e.source.getActiveSheet();
    var range = e.range;

    if (sheet.getName() !== "Techno Booking" || range.getRow() === 1) {
      return;
    }

    var row = range.getRow();

    var dataRow = sheet.getRange(row, 1, 1, 11).getValues()[0];

    var bookingId = dataRow[0];
    var clientName = dataRow[1];
    var phone = dataRow[2];
    var room = technoNormalizeRoom(dataRow[3]);
    var date = dataRow[4];
    var startTime = dataRow[5];
    var endTime = dataRow[6];
    var statusCell = sheet.getRange(row, 9);
    var action = dataRow[9];
    var notes = dataRow[10];

    // Delete action column J
    if (range.getColumn() === 10 && action === true) {
      technoDeleteLegacyBooking(sheet, row, bookingId, clientName, room, date, statusCell);
      return;
    }

    // Only trigger when EndTime column G is edited
    if (range.getColumn() !== 7) return;

    if (!clientName || !room || !date || !startTime || !endTime) {
      statusCell.setValue("⚠️ بيانات ناقصة");
      return;
    }

    var startDateTime = technoBuildDateTime(date, startTime);
    var endDateTime = technoBuildDateTime(date, endTime);

    if (endDateTime <= startDateTime) {
      statusCell.setValue("⚠️ وقت النهاية غير صحيح");
      return;
    }

    if (!bookingId) {
      bookingId = generateID("BOOK");
      sheet.getRange(row, 1).setValue(bookingId);
    }

    var availability = checkRoomAvailability(
      room,
      startDateTime,
      endDateTime,
      null,
      bookingId
    );

    if (!availability.success) {
      statusCell.setValue("⚠️ " + availability.error);
      return;
    }

    if (!availability.available) {
      statusCell.setValue("⚠️ محجوز بالفعل");
      return;
    }

    var calendar = technoGetCalendar(room);

    if (!calendar) {
      statusCell.setValue("⚠️ خطأ في معرف الغرفة");
      return;
    }

    var price = technoCalculateBookingPrice(room, startDateTime, endDateTime, "none", 0);

    var title = "حجز: " + clientName + " - " + room + " - " + bookingId;

    var event = calendar.createEvent(
      title,
      startDateTime,
      endDateTime,
      {
        description: "Techno Booking Sheet\nBookingID: " + bookingId + "\nClient: " + clientName + "\nPhone: " + phone + "\nNotes: " + (notes || "")
      }
    );

    sheet.getRange(row, 8).setValue(price.finalAmount);
    statusCell.setValue("✅ تم التأكيد");

    technoAppendLegacyToRoomBookings({
      bookingId: bookingId,
      clientName: clientName,
      phone: phone,
      room: room,
      start: startDateTime,
      end: endDateTime,
      price: price,
      eventId: event.getId(),
      notes: notes
    });

  } catch (err) {
    try {
      e.source.getActiveSheet().getRange(e.range.getRow(), 9).setValue("⚠️ تفقد الصلاحية");
    } catch (x) {}
  }
}

function technoAppendLegacyToRoomBookings(obj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("RoomBookings");

  if (!sheet) {
    sheet = ss.insertSheet("RoomBookings");
    sheet.appendRow([
      "BookingID","CustomerName","Phone","Room","Date","StartTime","EndTime","Hours",
      "HourlyRate","Subtotal","DiscountType","DiscountValue","DiscountReason","DiscountAmount",
      "FinalAmount","PaymentMethod","Status","CalendarEventID","CreatedBy","CreatedAt","UpdatedAt","Notes"
    ]);
    sheet.setFrozenRows(1);
  }

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(obj.bookingId)) {
      return;
    }
  }

  var now = new Date().toISOString();

  sheet.appendRow([
    obj.bookingId,
    obj.clientName,
    obj.phone,
    obj.room,
    technoFormatDateKey(obj.start),
    obj.start.toISOString(),
    obj.end.toISOString(),
    obj.price.hours,
    obj.price.hourlyRate,
    obj.price.subtotal,
    "none",
    0,
    "",
    0,
    obj.price.finalAmount,
    "Cash",
    "Confirmed",
    obj.eventId,
    "Techno Booking Sheet",
    now,
    now,
    obj.notes || ""
  ]);
}

function technoDeleteLegacyBooking(sheet, row, bookingId, clientName, room, date, statusCell) {
  var calendar = technoGetCalendar(room);

  if (calendar) {
    var events = calendar.getEventsForDay(new Date(date));

    var eventToDelete = events.find(function(event) {
      var title = event.getTitle();

      if (bookingId && title.indexOf(bookingId) !== -1) return true;
      if (clientName && title.indexOf(clientName) !== -1) return true;

      return false;
    });

    if (eventToDelete) {
      eventToDelete.deleteEvent();
      statusCell.setValue("❌ تم الحذف من التقويم");
    } else {
      statusCell.setValue("⚠️ لم يعثر عليه بالتقويم");
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rb = ss.getSheetByName("RoomBookings");

  if (rb && bookingId) {
    var data = rb.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(bookingId)) {
        rb.getRange(i + 1, 17).setValue("Cancelled");
        rb.getRange(i + 1, 21).setValue(new Date().toISOString());
        break;
      }
    }
  }

  Utilities.sleep(500);
  sheet.deleteRow(row);
}
