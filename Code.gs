// ============================================================
// CD Money — Google Apps Script Backend (full replacement)
// Deploy: Extensions > Apps Script > Deploy > Web App
//   Execute as: Me
//   Access: Anyone
//
// Required Script Properties:
//   GEMINI_API_KEY = your Gemini Developer API key
//   SPREADSHEET_ID = your Google Sheet ID
//
// Optional Script Properties:
//   GEMINI_MODEL = gemini-2.5-flash
//   APP_SECRET   = shared secret; if set, every request must include token
//                  either in JSON body or query string.
// ============================================================

var DEFAULT_GEMINI_MODEL = 'gemini-2.5-flash';
var MUTATING_ACTIONS = {
  upsertExpense: true,
  deleteExpense: true,
  saveRule: true,
  saveMonthSettings: true,
  saveRecurring: true,
  deleteRecurring: true,
  runDueRecurring: true,
  runRecurringNow: true,
  postRecurringExpense: true,
  saveFlexBudget: true,
  deleteFlexBudget: true,
  saveReceiptItems: true
};

function doGet(e) {
  try {
    var payload = {};
    assertAuthorized_(payload, e);
    var action = (e && e.parameter && e.parameter.action) || 'getState';
    var result = dispatch_(action, payload, e);
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_(errorResponse_(err));
  }
}

function doPost(e) {
  try {
    var payload = parseBody_(e);

    // APP_SECRET check must happen before any action routing.
    assertAuthorized_(payload, e);

    var action = payload.action || ((e && e.parameter && e.parameter.action) || '');
    delete payload.action;

    if (!action) {
      if (looksLikeExpensePayload_(payload)) {
        action = 'upsertExpense';
      } else {
        throw userError_('Missing action in request body.', false);
      }
    }

    var result = dispatch_(action, payload, e);
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_(errorResponse_(err));
  }
}

function dispatch_(action, payload, e) {
  if (MUTATING_ACTIONS[action]) {
    return withScriptLock_(function () {
      return dispatchUnlocked_(action, payload, e);
    });
  }
  return dispatchUnlocked_(action, payload, e);
}

function dispatchUnlocked_(action, p, e) {
  switch (action) {
    case 'getState':             return actionGetState_();
    case 'upsertExpense':        return actionUpsertExpense_(p);
    case 'deleteExpense':        return actionDeleteExpense_(p);
    case 'saveRule':             return actionSaveRule_(p);
    case 'saveMonthSettings':    return actionSaveMonthSettings_(p);
    case 'saveRecurring':        return actionSaveRecurring_(p);
    case 'deleteRecurring':      return actionDeleteRecurring_(p);
    case 'runDueRecurring':
    case 'runRecurringNow':      return actionRunDueRecurring_(p);
    case 'postRecurringExpense': return actionPostRecurringExpense_(p);
    case 'saveFlexBudget':       return actionSaveFlexBudget_(p);
    case 'deleteFlexBudget':     return actionDeleteFlexBudget_(p);
    case 'saveReceiptItems':     return actionSaveReceiptItems_(p);
    case 'parseReceiptImage':    return actionParseReceiptImage_(p);
    case 'getGeminiInsights':    return actionGetGeminiInsights_(p);
    case 'askGeminiFollowup':    return actionAskGeminiFollowup_(p);
    default:
      throw userError_('Unknown action: ' + action, false);
  }
}

function parseBody_(e) {
  var raw = (e && e.postData && typeof e.postData.contents === 'string') ? e.postData.contents : '';
  raw = String(raw || '').replace(/^\uFEFF/, '').trim();
  if (!raw) return {};
  try {
    return JSON.parse(raw);
  } catch (err) {
    throw userError_('Malformed JSON in request body.', false);
  }
}

function looksLikeExpensePayload_(payload) {
  return !!(
    payload && (
      Object.prototype.hasOwnProperty.call(payload, 'amount') ||
      Object.prototype.hasOwnProperty.call(payload, 'date') ||
      Object.prototype.hasOwnProperty.call(payload, 'category') ||
      Object.prototype.hasOwnProperty.call(payload, 'note') ||
      Object.prototype.hasOwnProperty.call(payload, 'source') ||
      Object.prototype.hasOwnProperty.call(payload, 'id')
    )
  );
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse_(err) {
  return {
    success: false,
    error: (err && err.message) ? err.message : String(err),
    retryable: !!(err && err.retryable)
  };
}

function userError_(message, retryable) {
  var err = new Error(message);
  err.retryable = !!retryable;
  return err;
}

function withScriptLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getConfig_() {
  var props = PropertiesService.getScriptProperties();
  return {
    geminiApiKey: props.getProperty('GEMINI_API_KEY') || '',
    spreadsheetId: props.getProperty('SPREADSHEET_ID') || '',
    geminiModel: props.getProperty('GEMINI_MODEL') || DEFAULT_GEMINI_MODEL,
    appSecret: props.getProperty('APP_SECRET') || ''
  };
}

function assertAuthorized_(payload, e) {
  var config = getConfig_();
  if (!config.appSecret) return;

  var tokenFromBody = payload && payload.token ? String(payload.token) : '';
  var tokenFromQuery = (e && e.parameter && e.parameter.token) ? String(e.parameter.token) : '';
  var token = tokenFromBody || tokenFromQuery;

  if (payload && Object.prototype.hasOwnProperty.call(payload, 'token')) {
    delete payload.token;
  }

  if (!token || token !== config.appSecret) {
    throw userError_('Unauthorized request.', false);
  }
}

function getSpreadsheet_() {
  var config = getConfig_();
  if (!config.spreadsheetId) {
    throw userError_('SPREADSHEET_ID is not set in Script Properties.', false);
  }
  return SpreadsheetApp.openById(config.spreadsheetId);
}

function getOrCreateSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }

  if (headers && headers.length) {
    var lastRow = sh.getLastRow();
    if (lastRow === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      var currentHeaders = sh.getRange(1, 1, 1, Math.max(headers.length, sh.getLastColumn())).getValues()[0]
        .slice(0, headers.length)
        .map(function (v) { return String(v || '').trim(); });
      var expected = headers.map(function (v) { return String(v || '').trim(); });
      if (JSON.stringify(currentHeaders) !== JSON.stringify(expected)) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
    }
  }

  return sh;
}

function sheetToObjects_(sh) {
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var data = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0].map(function (h) { return String(h || '').trim(); });

  return data.slice(1).map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) {
      obj[h] = row[i] === '' ? null : row[i];
    });
    return obj;
  });
}

function ensureCoreSheets_(ss) {
  return {
    expenses: getOrCreateSheet_(ss, 'Expenses', ['id', 'date', 'amount', 'category', 'note', 'source']),
    budgets: getOrCreateSheet_(ss, 'Budgets', ['month', 'income', 'flex']),
    rules: getOrCreateSheet_(ss, 'Rules', ['category', 'mode']),
    flexBudgets: getOrCreateSheet_(ss, 'FlexBudgets', ['month', 'category', 'budget']),
    recurring: getOrCreateSheet_(ss, 'Recurring', ['id', 'title', 'amount', 'category', 'dayOfMonth', 'status', 'lastAddedMonth'])
  };
}

function monthKeyFromDateLike_(value) {
  var str = String(value || '').trim();
  if (!str) return '';
  if (/^\d{4}-\d{2}$/.test(str)) return str;
  var d = new Date(str);
  if (isNaN(d.getTime())) return '';
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function addMonthsToKey_(key, delta) {
  var parts = String(key || '').split('-');
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10);
  if (!year || !month) return '';
  var d = new Date(year, month - 1 + delta, 1);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function isFixedCategoryByRules_(category, rules) {
  return String((rules && rules[category]) || 'flexible').toLowerCase() === 'fixed';
}

function collectKnownMonths_(expenses, budgets, flexBudgets) {
  var monthMap = {};
  monthMap[currentMonthKey_()] = true;

  (expenses || []).forEach(function (item) {
    var key = monthKeyFromDateLike_(item && item.date);
    if (key) monthMap[key] = true;
  });

  Object.keys(budgets || {}).forEach(function (key) {
    if (/^\d{4}-\d{2}$/.test(String(key))) monthMap[String(key)] = true;
  });

  Object.keys(flexBudgets || {}).forEach(function (key) {
    if (/^\d{4}-\d{2}$/.test(String(key))) monthMap[String(key)] = true;
  });

  return Object.keys(monthMap).sort();
}

function buildMonthlySpendTotals_(expenses, rules) {
  var out = {};
  (expenses || []).forEach(function (item) {
    var month = monthKeyFromDateLike_(item && item.date);
    if (!month) return;
    if (!out[month]) {
      out[month] = { allIn: 0, flexible: 0 };
    }
    var amount = parseAmount_(item && item.amount);
    out[month].allIn += amount;
    if (!isFixedCategoryByRules_(item && item.category, rules)) {
      out[month].flexible += amount;
    }
  });
  return out;
}

function buildRollingWindowAverages_(monthlyTotals, baseMonth, windowSize) {
  var months = [];
  var allInTotal = 0;
  var flexibleTotal = 0;

  for (var i = windowSize; i >= 1; i--) {
    var month = addMonthsToKey_(baseMonth, -i);
    months.push(month);
    var totals = monthlyTotals[month] || { allIn: 0, flexible: 0 };
    allInTotal += parseAmount_(totals.allIn);
    flexibleTotal += parseAmount_(totals.flexible);
  }

  return {
    months: months,
    allIn: allInTotal / windowSize,
    flexible: flexibleTotal / windowSize
  };
}

function buildRollingAveragesByBaseMonth_(expenses, rules, budgets, flexBudgets) {
  var monthlyTotals = buildMonthlySpendTotals_(expenses, rules);
  var knownMonths = collectKnownMonths_(expenses, budgets, flexBudgets);
  var out = {};

  knownMonths.forEach(function (baseMonth) {
    var current = monthlyTotals[baseMonth] || { allIn: 0, flexible: 0 };
    var avg3 = buildRollingWindowAverages_(monthlyTotals, baseMonth, 3);
    var avg6 = buildRollingWindowAverages_(monthlyTotals, baseMonth, 6);
    var avg12 = buildRollingWindowAverages_(monthlyTotals, baseMonth, 12);

    out[baseMonth] = {
      current: {
        allIn: parseAmount_(current.allIn),
        flexible: parseAmount_(current.flexible)
      },
      averages: {
        '3': avg3,
        '6': avg6,
        '12': avg12
      }
    };
  });

  return out;
}

function resolveMonthlyIncome_(budgets, month) {
  var budgetRow = (budgets && budgets[month]) || {};
  var income = parseAmount_(budgetRow.income);
  return income > 0 ? income : 5000;
}

function buildMacroIncomeByMonth_(expenses, budgets, flexBudgets) {
  var knownMonths = collectKnownMonths_(expenses, budgets, flexBudgets);
  var out = {};
  knownMonths.forEach(function (month) {
    out[month] = resolveMonthlyIncome_(budgets, month);
  });
  if (!out[currentMonthKey_()]) {
    out[currentMonthKey_()] = resolveMonthlyIncome_(budgets, currentMonthKey_());
  }
  return out;
}

function actionGetState_() {
  var ss = getSpreadsheet_();
  var sheets = ensureCoreSheets_(ss);

  var expenses = sheetToObjects_(sheets.expenses).map(function (r) {
    return {
      id: String(r.id || ''),
      date: String(r.date || ''),
      amount: parseAmount_(r.amount),
      category: String(r.category || ''),
      note: String(r.note || ''),
      source: String(r.source || 'manual')
    };
  });

  var budgets = {};
  sheetToObjects_(sheets.budgets).forEach(function (r) {
    var month = String(r.month || '').trim();
    if (!month) return;
    budgets[month] = {
      income: parseAmount_(r.income),
      flex: parseAmount_(r.flex)
    };
  });

  var rules = {};
  sheetToObjects_(sheets.rules).forEach(function (r) {
    var category = String(r.category || '').trim();
    if (!category) return;
    rules[category] = String(r.mode || 'flexible');
  });

  var flexBudgets = {};
  sheetToObjects_(sheets.flexBudgets).forEach(function (r) {
    var month = String(r.month || '').trim();
    var category = String(r.category || '').trim();
    if (!month || !category) return;
    if (!flexBudgets[month]) flexBudgets[month] = {};
    flexBudgets[month][category] = parseAmount_(r.budget);
  });

  var recurring = sheetToObjects_(sheets.recurring).map(function (r) {
    return {
      id: String(r.id || ''),
      title: String(r.title || ''),
      amount: parseAmount_(r.amount),
      category: String(r.category || ''),
      dayOfMonth: Math.max(1, parseInt(r.dayOfMonth, 10) || 1),
      status: String(r.status || 'Active'),
      lastAddedMonth: String(r.lastAddedMonth || '')
    };
  });

  var rollingAverages = buildRollingAveragesByBaseMonth_(expenses, rules, budgets, flexBudgets);
  var macroIncomeByMonth = buildMacroIncomeByMonth_(expenses, budgets, flexBudgets);

  return {
    success: true,
    state: {
      expenses: expenses,
      budgets: budgets,
      flexBudgets: flexBudgets,
      rules: rules,
      recurring: recurring,
      rollingAverages: rollingAverages,
      macroIncomeByMonth: macroIncomeByMonth
    }
  };
}

function actionUpsertExpense_(p) {
  var ss = getSpreadsheet_();
  var sh = getOrCreateSheet_(ss, 'Expenses', ['id', 'date', 'amount', 'category', 'note', 'source']);

  var id = String(p.id || generateId_('exp'));
  var rowData = [
    id,
    String(p.date || new Date().toISOString()),
    toNumber_(p.amount),
    String(p.category || ''),
    String(p.note || ''),
    String(p.source || 'manual')
  ];

  var lastRow = sh.getLastRow();
  if (lastRow > 1) {
    var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === id) {
        sh.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
        return { success: true, id: id, updated: true };
      }
    }
  }

  sh.appendRow(rowData);
  return { success: true, id: id, updated: false };
}

function actionDeleteExpense_(p) {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName('Expenses');
  if (!sh || sh.getLastRow() < 2) return { success: true };

  var targetId = String(p.id || '');
  if (!targetId) return { success: true };

  var values = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]) === targetId) {
      sh.deleteRow(i + 2);
      break;
    }
  }
  return { success: true };
}

function actionSaveRule_(p) {
  var ss = getSpreadsheet_();
  var sh = getOrCreateSheet_(ss, 'Rules', ['category', 'mode']);
  var category = String(p.category || '').trim();
  var mode = String(p.mode || 'flexible').trim();

  if (!category) {
    throw userError_('Category is required for saveRule.', false);
  }

  var lastRow = sh.getLastRow();
  if (lastRow > 1) {
    var values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]) === category) {
        sh.getRange(i + 2, 2).setValue(mode);
        return { success: true, updated: true };
      }
    }
  }

  sh.appendRow([category, mode]);
  return { success: true, updated: false };
}

function actionSaveMonthSettings_(p) {
  var ss = getSpreadsheet_();
  var sh = getOrCreateSheet_(ss, 'Budgets', ['month', 'income', 'flex']);
  var month = String(p.month || '').trim();
  var income = toNumber_(p.income);
  var flex = toNumber_(p.flex);

  if (!month) {
    throw userError_('Month is required for saveMonthSettings.', false);
  }

  var rowData = [month, income, flex];
  var lastRow = sh.getLastRow();
  if (lastRow > 1) {
    var values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === month) {
        sh.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
        return { success: true, updated: true };
      }
    }
  }

  sh.appendRow(rowData);
  return { success: true, updated: false };
}

function actionSaveRecurring_(p) {
  var ss = getSpreadsheet_();
  var sh = getOrCreateSheet_(ss, 'Recurring', ['id', 'title', 'amount', 'category', 'dayOfMonth', 'status', 'lastAddedMonth']);

  var id = String(p.id || generateId_('rec'));
  var rowData = [
    id,
    String(p.title || ''),
    toNumber_(p.amount),
    String(p.category || ''),
    Math.max(1, parseInt(p.dayOfMonth, 10) || 1),
    String(p.status || 'Active'),
    String(p.lastAddedMonth || '')
  ];

  var lastRow = sh.getLastRow();
  if (lastRow > 1) {
    var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === id) {
        sh.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
        return { success: true, id: id, updated: true };
      }
    }
  }

  sh.appendRow(rowData);
  return { success: true, id: id, updated: false };
}

function actionDeleteRecurring_(p) {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName('Recurring');
  if (!sh || sh.getLastRow() < 2) return { success: true };

  var targetId = String(p.id || '');
  if (!targetId) return { success: true };

  var values = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]) === targetId) {
      sh.deleteRow(i + 2);
      break;
    }
  }
  return { success: true };
}

function actionRunDueRecurring_(p) {
  var targetMonth = String(p.month || currentMonthKey_());
  var ss = getSpreadsheet_();
  var recSh = getOrCreateSheet_(ss, 'Recurring', ['id', 'title', 'amount', 'category', 'dayOfMonth', 'status', 'lastAddedMonth']);
  var expSh = getOrCreateSheet_(ss, 'Expenses', ['id', 'date', 'amount', 'category', 'note', 'source']);

  var recurring = sheetToObjects_(recSh);
  var addedRows = [];
  var updatedLastAdded = [];

  recurring.forEach(function (rec, index) {
    if (String(rec.status || '') !== 'Active') return;
    if (String(rec.lastAddedMonth || '') === targetMonth) return;

    var parts = targetMonth.split('-');
    var year = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10);
    if (!year || !month) return;

    var maxDay = new Date(year, month, 0).getDate();
    var day = Math.min(Math.max(1, parseInt(rec.dayOfMonth, 10) || 1), maxDay);
    var isoDate = new Date(year, month - 1, day).toISOString();

    addedRows.push([
      generateId_('rec'),
      isoDate,
      toNumber_(rec.amount),
      String(rec.category || ''),
      String(rec.title || ''),
      'recurring'
    ]);
    updatedLastAdded.push({ row: index + 2, value: targetMonth });
  });

  if (addedRows.length) {
    expSh.getRange(expSh.getLastRow() + 1, 1, addedRows.length, addedRows[0].length).setValues(addedRows);
    updatedLastAdded.forEach(function (entry) {
      recSh.getRange(entry.row, 7).setValue(entry.value);
    });
  }

  return { success: true, added: addedRows.length };
}


// ============================================================
// V2.0 — Midnight recurring automation
// ============================================================

function cronRunRecurring() {
  return actionRunDueRecurring_({ month: currentMonthKey_() });
}

function installMidnightTrigger() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'cronRunRecurring') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('cronRunRecurring')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .inTimezone('Europe/London')
    .create();

  return 'Midnight trigger installed. It will run daily between 00:00 and 01:00 Europe/London.';
}

function actionPostRecurringExpense_(p) {
  var recurringId = String(p.recurringId || p.id || '').trim();
  var targetMonth = String(p.lastAddedMonth || p.month || currentMonthKey_()).trim();
  if (!recurringId) {
    throw userError_('Recurring ID is required for postRecurringExpense.', false);
  }

  var ss = getSpreadsheet_();
  var recSh = getOrCreateSheet_(ss, 'Recurring', ['id', 'title', 'amount', 'category', 'dayOfMonth', 'status', 'lastAddedMonth']);
  var expSh = getOrCreateSheet_(ss, 'Expenses', ['id', 'date', 'amount', 'category', 'note', 'source']);

  var rows = recSh.getDataRange().getValues();
  if (rows.length < 2) {
    return { success: true, added: 0 };
  }

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== recurringId) continue;

    var title = String(rows[i][1] || '');
    var amount = toNumber_(rows[i][2]);
    var category = String(rows[i][3] || '');
    var dayOfMonth = Math.max(1, parseInt(rows[i][4], 10) || 1);
    var status = String(rows[i][5] || 'Active');
    var lastAddedMonth = String(rows[i][6] || '');

    if (status !== 'Active') {
      return { success: true, added: 0, skipped: 'inactive' };
    }
    if (lastAddedMonth === targetMonth) {
      return { success: true, added: 0, skipped: 'already_posted' };
    }

    var parts = targetMonth.split('-');
    var year = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10);
    if (!year || !month) {
      throw userError_('Invalid target month for postRecurringExpense.', false);
    }

    var maxDay = new Date(year, month, 0).getDate();
    var safeDay = Math.min(Math.max(1, dayOfMonth), maxDay);
    var isoDate = new Date(year, month - 1, safeDay).toISOString();

    expSh.appendRow([
      generateId_('rec'),
      isoDate,
      amount,
      category,
      title,
      'recurring'
    ]);

    recSh.getRange(i + 1, 7).setValue(targetMonth);
    return { success: true, added: 1 };
  }

  throw userError_('Recurring item not found for postRecurringExpense.', false);
}

function actionSaveFlexBudget_(p) {
  var ss = getSpreadsheet_();
  var sh = getOrCreateSheet_(ss, 'FlexBudgets', ['month', 'category', 'budget']);
  var month = String(p.month || '').trim();
  var category = String(p.category || '').trim();
  var budget = toNumber_(p.budget);

  if (!month || !category) {
    throw userError_('Month and category are required for saveFlexBudget.', false);
  }

  var rowData = [month, category, budget];
  var lastRow = sh.getLastRow();
  if (lastRow > 1) {
    var values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]) === month && String(values[i][1]) === category) {
        sh.getRange(i + 2, 1, 1, rowData.length).setValues([rowData]);
        return { success: true, updated: true };
      }
    }
  }

  sh.appendRow(rowData);
  return { success: true, updated: false };
}

function actionDeleteFlexBudget_(p) {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName('FlexBudgets');
  if (!sh || sh.getLastRow() < 2) return { success: true };

  var month = String(p.month || '').trim();
  var category = String(p.category || '').trim();
  if (!month || !category) return { success: true };

  var values = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]) === month && String(values[i][1]) === category) {
      sh.deleteRow(i + 2);
      break;
    }
  }
  return { success: true };
}

function actionSaveReceiptItems_(p) {
  var items = Array.isArray(p.items) ? p.items : [];
  var source = String(p.source || 'receipt_scan');
  var ss = getSpreadsheet_();
  var sh = getOrCreateSheet_(ss, 'Expenses', ['id', 'date', 'amount', 'category', 'note', 'source']);
  var nowIso = new Date().toISOString();

  var rows = [];
  items.forEach(function (item) {
    var amount = toNumber_(item.amount);
    if (amount <= 0) return;
    rows.push([
      generateId_('scan'),
      String(item.date || nowIso),
      amount,
      String(item.category || item.suggestedCategory || 'Food'),
      String(item.name || 'Receipt item'),
      source
    ]);
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  return { success: true, saved: rows.length };
}

function actionParseReceiptImage_(p) {
  var base64 = String(p.imageBase64 || '');
  var mimeType = String(p.mimeType || 'image/jpeg');

  if (!base64 || base64.length < 100) {
    throw userError_('No image data received.', false);
  }

  if (base64.length > 7000000) {
    throw userError_('Receipt image is too large. Please crop/compress the image and try again.', false);
  }

  var prompt = [
    'Extract the receipt into structured JSON.',
    'Include merchant, receiptDate, and item lines only.',
    'Do not include subtotal, total, VAT, tax, service charge, card payment lines, or change.',
    'If the receipt is unreadable, return merchant=null, receiptDate=null, items=[].',
    'Use one of these categories when possible: Food, Travel, Household, Personal, Entertainment, Gifts, Holiday, Bills, Transport, Other.'
  ].join('\n');

  var schema = {
    type: 'object',
    properties: {
      merchant: { type: ['string', 'null'] },
      receiptDate: { type: ['string', 'null'] },
      items: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            name: { type: 'string' },
            amount: { type: 'number' },
            suggestedCategory: { type: 'string' }
          },
          required: ['name', 'amount', 'suggestedCategory'],
          additionalProperties: false
        }
      }
    },
    required: ['merchant', 'receiptDate', 'items'],
    additionalProperties: false
  };

  var requestBody = {
    contents: [{
      parts: [
        {
          inline_data: {
            mime_type: mimeType,
            data: base64
          }
        },
        { text: prompt }
      ]
    }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseJsonSchema: schema,
      temperature: 0.1,
      maxOutputTokens: 1024
    }
  };

  var parsed = callGeminiJson_(requestBody);
  var parsedText = stripMarkdownJson_(JSON.stringify(parsed));

  try {
    parsed = JSON.parse(parsedText);
  } catch (err) {
    throw userError_('Gemini returned malformed JSON content.', true);
  }

  var items = Array.isArray(parsed.items) ? parsed.items : [];
  items = items
    .map(function (item) {
      return {
        name: String(item.name || '').trim(),
        amount: toNumber_(item.amount),
        suggestedCategory: String(item.suggestedCategory || 'Other').trim() || 'Other'
      };
    })
    .filter(function (item) {
      return item.name && item.amount > 0;
    });

  return {
    success: true,
    merchant: parsed.merchant || null,
    receiptDate: parsed.receiptDate || null,
    items: items,
    debug: {
      model: getConfig_().geminiModel,
      itemCount: items.length
    }
  };
}

function actionGetGeminiInsights_(p) {
  var summary = p.summary || {};
  var topCategories = Array.isArray(p.topCategories) ? p.topCategories : [];
  var topItems = Array.isArray(p.topItems) ? p.topItems : [];
  var repeatedItems = Array.isArray(p.repeatedItems) ? p.repeatedItems : [];
  var categoryBudgets = p.categoryBudgets || {};
  var safeDailySpend = toNumber_(p.safeDailySpend);

  var contextLines = [
    'Month: ' + String(p.month || 'unknown'),
    'Total spend: £' + toNumber_(summary.totalSpend).toFixed(2),
    'Fixed spend: £' + toNumber_(summary.fixedSpend).toFixed(2),
    'Flexible spend: £' + toNumber_(summary.flexibleSpend).toFixed(2),
    'Flexible budget remaining: £' + toNumber_(summary.flexibleLeft).toFixed(2),
    'Income: £' + toNumber_(summary.income).toFixed(2),
    'Safe daily spend: £' + safeDailySpend.toFixed(2),
    'Top categories: ' + topCategories.map(function (x) {
      return String(x.name || '?') + ' £' + toNumber_(x.total).toFixed(2);
    }).join(', '),
    'Top items: ' + topItems.map(function (x) {
      return String(x.name || x.note || '?') + ' £' + toNumber_(x.total).toFixed(2) + ' x' + (parseInt(x.count, 10) || 0);
    }).join(', '),
    'Repeated items: ' + repeatedItems.map(function (x) {
      return String(x.name || x.note || '?') + ' x' + (parseInt(x.count, 10) || 0);
    }).join(', '),
    'Category budgets: ' + Object.keys(categoryBudgets).map(function (key) {
      var b = categoryBudgets[key] || {};
      return key + ' budget £' + toNumber_(b.budget).toFixed(2) +
        ' spent £' + toNumber_(b.spent).toFixed(2) +
        ' left £' + toNumber_(b.left).toFixed(2);
    }).join('; ')
  ].join('\n');

  var prompt = [
    'You are a concise personal finance coach.',
    'Analyze the data and return structured JSON.',
    'Be specific and practical. Avoid generic advice.',
    '',
    contextLines
  ].join('\n');

  var schema = {
    type: 'object',
    properties: {
      headline: { type: 'string' },
      top_drains: { type: 'array', items: { type: 'string' } },
      repeated_spend: { type: 'array', items: { type: 'string' } },
      budget_pressure: { type: 'array', items: { type: 'string' } },
      actions: { type: 'array', items: { type: 'string' } }
    },
    required: ['headline', 'top_drains', 'repeated_spend', 'budget_pressure', 'actions'],
    additionalProperties: false
  };

  var requestBody = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseJsonSchema: schema,
      temperature: 0.3,
      maxOutputTokens: 900
    }
  };

  var insights = callGeminiJson_(requestBody);
  var insightsText = stripMarkdownJson_(JSON.stringify(insights));

  try {
    insights = JSON.parse(insightsText);
  } catch (err) {
    throw userError_('Gemini returned malformed JSON content.', true);
  }

  return {
    success: true,
    insights: {
      headline: String(insights.headline || 'AI analysis unavailable.'),
      top_drains: Array.isArray(insights.top_drains) ? insights.top_drains : [],
      repeated_spend: Array.isArray(insights.repeated_spend) ? insights.repeated_spend : [],
      budget_pressure: Array.isArray(insights.budget_pressure) ? insights.budget_pressure : [],
      actions: Array.isArray(insights.actions) ? insights.actions : []
    }
  };
}

function actionAskGeminiFollowup_(p) {
  var question = String(p.question || '').trim();
  if (!question) {
    throw userError_('Question is required for askGeminiFollowup.', false);
  }

  var summary = p.summary || {};
  var topCategories = Array.isArray(p.topCategories) ? p.topCategories : [];
  var topItems = Array.isArray(p.topItems) ? p.topItems : [];
  var repeatedItems = Array.isArray(p.repeatedItems) ? p.repeatedItems : [];
  var categoryBudgets = p.categoryBudgets || {};
  var safeDailySpend = toNumber_(p.safeDailySpend);

  var contextLines = [
    'Month: ' + String(p.month || 'unknown'),
    'Total spend: £' + toNumber_(summary.totalSpend).toFixed(2),
    'Fixed spend: £' + toNumber_(summary.fixedSpend).toFixed(2),
    'Flexible spend: £' + toNumber_(summary.flexibleSpend).toFixed(2),
    'Flexible budget remaining: £' + toNumber_(summary.flexibleLeft).toFixed(2),
    'Income: £' + toNumber_(summary.income).toFixed(2),
    'Safe daily spend: £' + safeDailySpend.toFixed(2),
    'Top categories: ' + topCategories.map(function (x) {
      return String(x.name || '?') + ' £' + toNumber_(x.total).toFixed(2);
    }).join(', '),
    'Top items: ' + topItems.map(function (x) {
      return String(x.name || x.note || '?') + ' £' + toNumber_(x.total).toFixed(2) + ' x' + (parseInt(x.count, 10) || 0);
    }).join(', '),
    'Repeated items: ' + repeatedItems.map(function (x) {
      return String(x.name || x.note || '?') + ' x' + (parseInt(x.count, 10) || 0);
    }).join(', '),
    'Category budgets: ' + Object.keys(categoryBudgets).map(function (key) {
      var b = categoryBudgets[key] || {};
      return key + ' budget £' + toNumber_(b.budget).toFixed(2) +
        ' spent £' + toNumber_(b.spent).toFixed(2) +
        ' left £' + toNumber_(b.left).toFixed(2);
    }).join('; ')
  ].join('\n');

  var prompt = [
    'You are a strict, analytical personal finance advisor.',
    'Do not answer lazily and do not simply repeat balances back to the user.',
    'Use the provided numbers to reason forward, not just describe the current state.',
    'You must think like a budgeting strategist and give predictive, actionable advice.',
    'When relevant, estimate what the user can save over the next 7 days, 14 days, and by month-end.',
    'When relevant, use safeDailySpend, current flexible spend, repeated purchases, category budgets, and remaining category headroom.',
    'When relevant, identify likely wasteful spending patterns, daily burn rate risk, and whether the user is on track or drifting off track.',
    'If the user asks what they can afford, how much they can save, or how much they can spend, give a practical answer with a strict recommendation.',
    'If the question is time-based, project the answer using the remaining budget and current pace rather than only stating the current remaining balance.',
    'Keep the answer concise but sharp, specific, and useful.',
    'Return ONLY valid JSON in this exact structure: {"answer":"..."}',
    '',
    'Budget context:',
    contextLines,
    '',
    'User question: ' + question
  ].join('\n');

  var schema = {
    type: 'object',
    properties: {
      answer: { type: 'string' }
    },
    required: ['answer'],
    additionalProperties: false
  };

  var requestBody = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseJsonSchema: schema,
      temperature: 0.35,
      maxOutputTokens: 700
    }
  };

  var result = callGeminiJson_(requestBody);
  return {
    success: true,
    answer: String(result.answer || '').trim() || 'No answer returned.'
  };
}

function callGeminiJson_(requestBody) {
  var config = getConfig_();
  if (!config.geminiApiKey) {
    throw userError_('GEMINI_API_KEY is not set in Script Properties.', false);
  }

  requestBody = requestBody || {};
  requestBody.generationConfig = requestBody.generationConfig || {};
  requestBody.generationConfig.responseMimeType = 'application/json';

  var endpoint = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(config.geminiModel) +
    ':generateContent?key=' +
    encodeURIComponent(config.geminiApiKey);

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  var lastError = null;
  for (var attempt = 1; attempt <= 3; attempt++) {
    var response = UrlFetchApp.fetch(endpoint, options);
    var status = response.getResponseCode();
    var text = response.getContentText();

    if (status >= 200 && status < 300) {
      var envelope;
      try {
        envelope = JSON.parse(text);
      } catch (err) {
        throw userError_('Gemini returned invalid JSON envelope.', true);
      }

      var candidate = envelope && envelope.candidates && envelope.candidates[0];
      var part = candidate && candidate.content && candidate.content.parts && candidate.content.parts[0];
      var outputText = part && part.text;
      if (!outputText) {
        throw userError_('Gemini returned no content.', true);
      }

      var cleanedText = stripMarkdownJson_(outputText);
      try {
        return JSON.parse(cleanedText);
      } catch (err2) {
        throw userError_('Gemini returned malformed JSON content.', true);
      }
    }

    lastError = { status: status, text: text };
    if ((status === 429 || status >= 500) && attempt < 3) {
      Utilities.sleep(attempt * 800);
      continue;
    }

    var retryable = (status === 429 || status >= 500);
    throw userError_('Gemini API error ' + status + ': ' + String(text || '').slice(0, 300), retryable);
  }

  throw userError_('Gemini call failed: ' + JSON.stringify(lastError), true);
}

function stripMarkdownJson_(text) {
  var str = String(text || '').trim();

  // Brute force: find the first { and the last }
  var startIndex = str.indexOf('{');
  var endIndex = str.lastIndexOf('}');

  if (startIndex !== -1 && endIndex !== -1 && endIndex >= startIndex) {
    // Extract ONLY the JSON object, completely ignoring conversational text
    return str.substring(startIndex, endIndex + 1);
  }

  // Fallback if no brackets are found
  return str.replace(/```json|```/gi, '').trim();
}

function parseAmount_(value) {
  var num = parseFloat(String(value == null ? '' : value).replace(/,/g, '').trim());
  return isNaN(num) ? 0 : num;
}

function toNumber_(value) {
  var num = Number(value);
  return isNaN(num) ? 0 : num;
}

function generateId_(prefix) {
  return String(prefix || 'id') + '-' + new Date().getTime() + '-' + Math.random().toString(16).slice(2, 10);
}

function currentMonthKey_() {
  var d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}