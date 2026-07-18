/**
 * Business Finance — Apps Script JSON API backend.
 * Deployed as a Web App (Anyone can access) so the GitHub Pages dashboard
 * can fetch/write data cross-origin.
 *
 * Sheets used (already present in this spreadsheet):
 *  - Daily_Logs:      A=Date  B=Revenue  C=Daily Expenses  D=Just_Date (=INT(A))  E=Profit (=B-C)
 *  - Fixed_Expenses:  A=Expense Name  B=Amount
 *  - Monthly_Summary: A=Month (YYYY-MM)  B=Total Revenue  C=Total Daily Expenses
 *                      D=Fixed Expenses  E=Net Profit (Before Fixed)  F=Actual Profit (After Fixed)
 *    (Monthly_Summary is formula-driven off Daily_Logs/Fixed_Expenses — this script
 *    only reads it, never writes it.)
 */

var TZ_FALLBACK = 'Asia/Riyadh';

function getTimeZone_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || TZ_FALLBACK;
}

function fmtDate_(date) {
  return Utilities.formatDate(date, getTimeZone_(), 'yyyy-MM-dd');
}

/** instanceof Date can fail on values coming back from Range#getValues(); duck-type instead. */
function isDateValue_(v) {
  return !!v && typeof v.getFullYear === 'function' && !isNaN(v.getTime());
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    return jsonOut_(getDashboardPayload_());
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var date = String(body.date || '').trim();
    var revenue = Number(body.revenue) || 0;
    var expenses = Number(body.expenses) || 0;

    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      throw new Error('date must be in YYYY-MM-DD format');
    }

    upsertDailyLog_(date, revenue, expenses);

    var result = { ok: true, date: date, revenue: revenue, expenses: expenses, profit: revenue - expenses };
    result.data = getDashboardPayload_();
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/** Removes any existing Daily_Logs rows for `date`, then appends one fresh row. */
function upsertDailyLog_(date, revenue, expenses) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily_Logs');
  var tz = getTimeZone_();
  var lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    var dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    // Delete bottom-up so row indices don't shift while removing matches.
    for (var i = dates.length - 1; i >= 0; i--) {
      var cell = dates[i][0];
      if (isDateValue_(cell) && Utilities.formatDate(cell, tz, 'yyyy-MM-dd') === date) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  var newRow = sheet.getLastRow() + 1;
  var dateObj = new Date(date + 'T12:00:00');
  sheet.getRange(newRow, 1, 1, 3).setValues([[dateObj, revenue, expenses]]);
  sheet.getRange(newRow, 4).setFormula('=INT(A' + newRow + ')');
  sheet.getRange(newRow, 5).setFormula('=B' + newRow + '-C' + newRow);
}

/** Reads Daily_Logs and collapses duplicate-date rows (sum) into one entry per day. */
function getDailyLogs_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily_Logs');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var tz = getTimeZone_();
  var rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var byDate = {};
  var order = [];

  rows.forEach(function (row) {
    var d = row[0];
    if (!isDateValue_(d)) return;
    var key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    var revenue = Number(row[1]) || 0;
    var expenses = Number(row[2]) || 0;
    if (!byDate[key]) {
      byDate[key] = { date: key, revenue: 0, expenses: 0 };
      order.push(key);
    }
    byDate[key].revenue += revenue;
    byDate[key].expenses += expenses;
  });

  order.sort();
  return order.map(function (key) {
    var e = byDate[key];
    return { date: e.date, revenue: e.revenue, expenses: e.expenses, profit: e.revenue - e.expenses };
  });
}

function getFixedExpenses_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Fixed_Expenses');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { items: [], total: 0 };

  var rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var items = [];
  var total = 0;
  rows.forEach(function (row) {
    var name = row[0];
    var amount = Number(row[1]) || 0;
    if (!name) return;
    items.push({ name: name, amount: amount });
    total += amount;
  });
  return { items: items, total: total };
}

function getMonthlySummary_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Monthly_Summary');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var tz = getTimeZone_();
  var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var out = [];
  rows.forEach(function (row) {
    var m = row[0];
    if (!m) return;
    var month = isDateValue_(m) ? Utilities.formatDate(m, tz, 'yyyy-MM') : String(m);
    out.push({
      month: month,
      revenue: Number(row[1]) || 0,
      dailyExpenses: Number(row[2]) || 0,
      fixedExpenses: Number(row[3]) || 0,
      netBeforeFixed: Number(row[4]) || 0,
      actualProfit: Number(row[5]) || 0
    });
  });
  return out;
}

function getDashboardPayload_() {
  var tz = getTimeZone_();
  var fixed = getFixedExpenses_();
  return {
    ok: true,
    today: fmtDate_(new Date()),
    timeZone: tz,
    daily: getDailyLogs_(),
    monthly: getMonthlySummary_(),
    fixedExpenses: fixed.items,
    fixedExpensesTotal: fixed.total,
    shares: { operator: 0.10, partner1: 0.30, partner2: 0.30, partner3: 0.30 }
  };
}
