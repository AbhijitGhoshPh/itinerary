// ================================================================
// SIMLIPAL TRIP — Shared Expense Tracker
// Google Apps Script — paste this into your script editor
//
// Web App URL (already deployed):
//   https://script.google.com/macros/s/AKfycby5e_fZyXSC6i_p20_IMbjnEMtRCjQ2MUVJ_xGK7MwQF3SieV-qnL2E7kAt2UTW-AosDw/exec
//
// Trip Code:  SIMLIPAL26
//
// ⚠️  IMPORTANT — IF THE PAGE SHOWS "Network error" OR "Sign-in page":
//   The web app is redirecting to a Google login page. This happens when
//   the deployment was not set to "Anyone" access correctly.
//   Fix it in 3 clicks:
//     1. Open the Sheet → Extensions → Apps Script
//     2. Click Deploy → Manage deployments
//     3. Click the ✏️ pencil (edit) on your deployment
//     4. Under "Who has access" — make sure it says "Anyone"
//        (NOT "Anyone with Google account" and NOT "Only myself")
//     5. Click Deploy → copy the URL again (it may change)
//     6. Paste the new URL into the Expenses tab → Change URL field
//
// ================================================================

const SHEET_NAME = 'Expenses';
const TRIP_CODE  = 'SIMLIPAL26';   // shared with the group via WhatsApp

// ------------------------------------------------------------------
// doGet — the page calls this to load all expenses
// Called as:  GET  <url>?code=SIMLIPAL26
// ------------------------------------------------------------------
function doGet(e) {
  const params = e.parameter || {};

  if (params.code !== TRIP_CODE) {
    return jsonOut({ error: 'Invalid trip code' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(SHEET_NAME);
  if (!sheet) {
    return jsonOut({
      error: 'Sheet tab named "' + SHEET_NAME + '" not found. ' +
             'Rename your first sheet tab to exactly: Expenses'
    });
  }

  const allValues = sheet.getDataRange().getValues();

  if (allValues.length < 2) {
    return jsonOut({ rows: [] });   // only the header row exists, no data yet
  }

  const headers = allValues[0].map(h => String(h).trim());
  const rows = allValues.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });

  return jsonOut({ rows: rows });
}

// ------------------------------------------------------------------
// doPost — the page calls this to add or delete one expense
//
// Add body:
//   { action:"add", code:"SIMLIPAL26",
//     id, date, paidBy, category, amount, note, sharedWith, addedBy }
//
// Delete body:
//   { action:"delete", code:"SIMLIPAL26", id }
// ------------------------------------------------------------------
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ error: 'Bad JSON: ' + err.toString() });
  }

  if (!body || body.code !== TRIP_CODE) {
    return jsonOut({ error: 'Invalid trip code' });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(SHEET_NAME);
  if (!sheet) {
    return jsonOut({ error: 'Sheet "' + SHEET_NAME + '" not found' });
  }

  // ---- ADD a new row ----
  if (body.action === 'add') {
    sheet.appendRow([
      String(body.id       || Date.now()),
      String(body.date     || ''),
      String(body.paidBy   || ''),
      String(body.category || ''),
      Number(body.amount   || 0),
      String(body.note     || ''),
      String(body.sharedWith || 'all'),
      String(body.addedBy  || ''),
      new Date().toISOString()          // server timestamp
    ]);
    return jsonOut({ ok: true });
  }

  // ---- DELETE a row by id ----
  if (body.action === 'delete') {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(body.id)) {
        sheet.deleteRow(i + 1);         // getDataRange is 0-indexed; sheet rows are 1-indexed
        return jsonOut({ ok: true });
      }
    }
    return jsonOut({ error: 'Row not found for id: ' + body.id });
  }

  return jsonOut({ error: 'Unknown action: ' + (body.action || 'undefined') });
}

// ------------------------------------------------------------------
// Helper — JSON response
// Note: Apps Script Web Apps handle CORS automatically when deployed
// with "Execute as: Me" and "Who has access: Anyone".
// ------------------------------------------------------------------
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
