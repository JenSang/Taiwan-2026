/**
 * Google Apps Script web app that lets the Taiwan Trip Planner read and write
 * expenses in this spreadsheet.
 *
 * Expected layout of each day tab (named "Day 1 - 14 Oct", "Day 2 - 15 Oct", ...):
 *   Row 1: title
 *   Row 2: headers  Name | Item | Amount | Payment Methods
 *   Row 3+: one expense per row (hidden column E holds a row id used by the web page)
 *   Next:  "ruam (Total)" row, day sum in column C (read by the Summary tab)
 * Rows are inserted/deleted above the total row and its SUM formula is rewritten.
 *
 * Deploy: Deploy -> New deployment -> Web app -> Execute as: Me, Who has access: Anyone.
 */

const FIRST_DATA_ROW = 3;
const ID_COL = 5; // column E

function doGet(e) {
  return json_(handle_({ action: 'list', day: Number(e.parameter.day) }));
}

function doPost(e) {
  return json_(handle_(JSON.parse(e.postData.contents)));
}

function handle_(req) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // three people may save at the same moment
  try {
    const sheet = daySheet_(req.day);
    if (req.action === 'sync') {
      const remove = new Set(req.remove || []);
      let rows = readRows_(sheet).filter(r => !remove.has(r.id));
      const have = new Set(rows.map(r => r.id));
      (req.add || []).forEach(a => {
        if (!a.id || have.has(a.id)) return; // already saved (retry after a network error)
        rows.push({ id: a.id, name: a.name, item: a.item, amount: Number(a.amount) || 0, payment: a.payment });
      });
      writeRows_(sheet, rows);
    }
    return { ok: true, rows: readRows_(sheet) };
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  } finally {
    lock.releaseLock();
  }
}

function daySheet_(day) {
  const sheet = SpreadsheetApp.getActive().getSheets()
    .find(s => s.getName().indexOf('Day ' + day + ' ') === 0);
  if (!sheet) throw new Error('No sheet found for day ' + day);
  return sheet;
}

// Matches the "(sparkle) ruam (Total)" label. Thai written as \u escapes so the
// check survives copy/paste through tools that mangle non-ASCII text.
function isTotalLabel_(value) {
  const s = String(value);
  return s.indexOf('\u0E23\u0E27\u0E21') !== -1 || /total/i.test(s);
}

// Number of expense rows between the header and the total row.
function dataRowCount_(sheet) {
  const lastRow = sheet.getMaxRows();
  if (lastRow < FIRST_DATA_ROW) return 0;
  const col = sheet.getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, 1).getDisplayValues();
  for (let i = 0; i < col.length; i++) {
    if (isTotalLabel_(col[i][0])) {
      sheet.getRange(FIRST_DATA_ROW + i, ID_COL).clearContent(); // never keep a row id on the total row
      return i;
    }
  }
  throw new Error('Total row not found in ' + sheet.getName());
}

function readRows_(sheet) {
  const count = dataRowCount_(sheet);
  if (count <= 0) return [];
  const range = sheet.getRange(FIRST_DATA_ROW, 1, count, ID_COL);
  const values = range.getValues();
  let idsAdded = false;
  const rows = [];
  values.forEach(r => {
    if (r[0] === '' && r[1] === '' && r[2] === '') return;
    if (!r[4]) { r[4] = Utilities.getUuid(); idsAdded = true; } // row typed directly into the sheet
    rows.push({ id: String(r[4]), name: String(r[0]), item: String(r[1]), amount: Number(r[2]) || 0, payment: String(r[3]) });
  });
  if (idsAdded) range.setValues(values);
  return rows;
}

function writeRows_(sheet, rows) {
  // Data rows sit between the header and the total row. Grow or shrink that
  // block so the total row always sits right under the last expense.
  const have = dataRowCount_(sheet);
  const need = rows.length;
  const totalRow = FIRST_DATA_ROW + have;
  if (need > have) sheet.insertRowsBefore(totalRow, need - have);
  else if (need < have) sheet.deleteRows(FIRST_DATA_ROW + need, have - need);
  if (!sheet.isColumnHiddenByUser(ID_COL)) sheet.hideColumns(ID_COL);

  if (need) {
    sheet.getRange(FIRST_DATA_ROW, 1, need, ID_COL)
      .setValues(rows.map(r => [r.name, r.item, r.amount, r.payment, r.id]));
  }

  // The Summary tab reads each day's total from this cell.
  const total = sheet.getRange(FIRST_DATA_ROW + need, 3);
  if (need) total.setFormula('=SUM(C' + FIRST_DATA_ROW + ':C' + (FIRST_DATA_ROW + need - 1) + ')');
  else total.setValue(0);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
