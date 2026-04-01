// ─── Configuration ───────────────────────────────────────────────
const SPREADSHEET_ID = '1xHqwlTWMthb7U_L1ymMcxyCgpvPzFr3n4zrapqO5o9Y';

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: ' + name);
  return sheet;
}

// ─── Helpers ─────────────────────────────────────────────────────
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).filter(row => row[0] !== '').map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      // Convert Date objects to YYYY-MM-DD string
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[h] = val;
    });
    return obj;
  });
}

function uid() {
  return 'id' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function simpleHash(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + ch;
    hash |= 0;
  }
  return 'h' + Math.abs(hash).toString(16);
}

function jsonResp(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function ok(data)  { return jsonResp({ ok: true, data: data }); }
function fail(msg) { return jsonResp({ ok: false, error: msg }); }

// ─── CORS-friendly GET + POST handlers ──────────────────────────
function doGet(e)  { return route(e.parameter); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (_) {}
  if (e.parameter) Object.keys(e.parameter).forEach(k => p[k] = p[k] || e.parameter[k]);
  return route(p);
}

// ─── Router ─────────────────────────────────────────────────────
function route(p) {
  try {
    // Clean action parameter (strip trailing ? or whitespace)
    if (p.action) p.action = p.action.replace(/[?\s]/g, '').trim();
    switch (p.action) {
      case 'sync':                return doSync(p);
      case 'addTransactions':     return doAddTransactions(p);
      case 'updateTransaction':   return doUpdateTransaction(p);
      case 'bulkUpdateCategory':  return doBulkUpdateCategory(p);
      case 'deleteTransactions':  return doDeleteTransactions(p);
      case 'saveBudgets':         return doSaveBudgets(p);
      case 'saveSettings':        return doSaveSettings(p);
      case 'saveCustomCategory':  return doSaveCustomCategory(p);
      case 'deleteCustomCategory':return doDeleteCustomCategory(p);
      default:                    return fail('Unknown action: ' + p.action);
    }
  } catch (err) {
    return fail(err.toString());
  }
}

// ─── Sync: return everything in one call ────────────────────────
function doSync() {
  const transactions = sheetToObjects(getSheet('Transactions'));
  const budgets = sheetToObjects(getSheet('Budgets'));
  const settingsRows = sheetToObjects(getSheet('Settings'));
  const customCategories = sheetToObjects(getSheet('CustomCategories'));

  // Convert settings rows to an object
  const settings = {};
  settingsRows.forEach(row => {
    try {
      settings[row.key] = JSON.parse(row.value);
    } catch (_) {
      settings[row.key] = row.value;
    }
  });

  return ok({
    transactions: transactions,
    budgets: budgets,
    settings: settings,
    customCategories: customCategories
  });
}

// ─── Add Transactions (with dedup) ──────────────────────────────
function doAddTransactions(p) {
  const incoming = p.transactions || [];
  if (!incoming.length) return ok({ added: 0, skipped: 0 });

  const sheet = getSheet('Transactions');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const hashCol = headers.indexOf('hash');

  // Collect existing hashes
  const existingHashes = new Set();
  if (hashCol >= 0) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][hashCol]) existingHashes.add(String(data[i][hashCol]));
    }
  }

  let added = 0;
  let skipped = 0;
  const newRows = [];

  incoming.forEach(tx => {
    const hash = tx.hash || simpleHash(tx.date + '|' + tx.description + '|' + tx.amount + '|' + tx.account);
    if (existingHashes.has(hash)) {
      skipped++;
      return;
    }
    existingHashes.add(hash); // prevent dupes within same batch
    const id = tx.id || uid();
    newRows.push([
      id,
      tx.date || '',
      tx.type || '',
      tx.description || '',
      tx.amount || 0,
      tx.balance || '',
      tx.category || 'Other',
      tx.account || '',
      hash
    ]);
    added++;
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }

  return ok({ added: added, skipped: skipped });
}

// ─── Update single transaction ──────────────────────────────────
function doUpdateTransaction(p) {
  const sheet = getSheet('Transactions');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  const catCol = headers.indexOf('category');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === p.id) {
      if (p.category !== undefined) sheet.getRange(i + 1, catCol + 1).setValue(p.category);
      return ok({ updated: true });
    }
  }
  return fail('Transaction not found: ' + p.id);
}

// ─── Bulk update category by merchant name hash ─────────────────
function doBulkUpdateCategory(p) {
  const sheet = getSheet('Transactions');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const catCol = headers.indexOf('category');
  const descCol = headers.indexOf('description');

  // p.merchantName is the simplified merchant name
  // p.category is the new category
  // p.ids is an array of transaction IDs to update
  const idSet = new Set(p.ids || []);
  let count = 0;

  if (idSet.size > 0) {
    const idCol = headers.indexOf('id');
    for (let i = 1; i < data.length; i++) {
      if (idSet.has(data[i][idCol])) {
        sheet.getRange(i + 1, catCol + 1).setValue(p.category);
        count++;
      }
    }
  }

  return ok({ updated: count });
}

// ─── Delete transactions ────────────────────────────────────────
function doDeleteTransactions(p) {
  const ids = p.ids || [];
  if (!ids.length) return ok({ deleted: 0 });

  const sheet = getSheet('Transactions');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  const idSet = new Set(ids);

  // Delete from bottom to top to preserve row indices
  let deleted = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (idSet.has(data[i][idCol])) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }

  return ok({ deleted: deleted });
}

// ─── Save Budgets (replace all) ─────────────────────────────────
function doSaveBudgets(p) {
  const sheet = getSheet('Budgets');
  const budgets = p.budgets || {};

  // Clear existing data (keep headers)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 3).clearContent();

  const entries = Object.entries(budgets).filter(([, v]) => v > 0);
  if (entries.length > 0) {
    const rows = entries.map(([cat, amt]) => [uid(), cat, amt]);
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  return ok({ saved: entries.length });
}

// ─── Save Settings ──────────────────────────────────────────────
function doSaveSettings(p) {
  const sheet = getSheet('Settings');
  const settings = p.settings || {};
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol = headers.indexOf('key');
  const valCol = headers.indexOf('value');

  // Build a map of existing rows
  const rowMap = {};
  for (let i = 1; i < data.length; i++) {
    rowMap[data[i][keyCol]] = i + 1; // 1-indexed sheet row
  }

  for (const [key, value] of Object.entries(settings)) {
    const valStr = typeof value === 'string' ? value : JSON.stringify(value);
    if (rowMap[key]) {
      sheet.getRange(rowMap[key], valCol + 1).setValue(valStr);
    } else {
      sheet.appendRow([key, valStr]);
    }
  }

  return ok({ saved: Object.keys(settings).length });
}

// ─── Save Custom Category ───────────────────────────────────────
function doSaveCustomCategory(p) {
  const sheet = getSheet('CustomCategories');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('name');

  // Check if exists (update)
  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === p.name) {
      if (p.color !== undefined) sheet.getRange(i + 1, headers.indexOf('color') + 1).setValue(p.color);
      if (p.keywords !== undefined) sheet.getRange(i + 1, headers.indexOf('keywords') + 1).setValue(p.keywords);
      return ok({ updated: true });
    }
  }

  // Insert new
  sheet.appendRow([uid(), p.name, p.color || '#e84393', p.keywords || '']);
  return ok({ created: true });
}

// ─── Delete Custom Category ─────────────────────────────────────
function doDeleteCustomCategory(p) {
  const sheet = getSheet('CustomCategories');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('name');

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][nameCol] === p.name) {
      sheet.deleteRow(i + 1);
      return ok({ deleted: true });
    }
  }
  return fail('Category not found: ' + p.name);
}
