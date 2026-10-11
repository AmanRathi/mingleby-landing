/**
 * Mingleby signups → Google Sheet.
 *
 * Step 1 ("signup")    appends a row with name/mobile right away.
 * Step 2 ("interests") finds that row by lead_id and fills in interests/city.
 *                      If the row is missing, it appends a full row instead,
 *                      so name and mobile are never lost.
 *
 * Setup: see apps-script/README.md
 */

const SHEET_NAME = 'Signups';
const HEADERS = [
  'lead_id', 'created_at', 'updated_at', 'name', 'partner_name', 'mobile', 'source',
  'interests', 'city', 'step2_status',
  'utm_source', 'utm_medium', 'utm_campaign', 'utm_term', 'utm_content',
  'page', 'user_agent',
];
// Fields step 2 is allowed to overwrite on an existing row.
const STEP2_FIELDS = ['interests', 'city', 'step2_status'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (!data.lead_id) throw new Error('missing lead_id');
    const sheet = getSheet_();
    const now = new Date();

    if (data.action === 'signup') {
      if (!/^[6-9]\d{9}$/.test(String(data.mobile || ''))) throw new Error('invalid mobile');
      appendRow_(sheet, { ...data, created_at: now, updated_at: now, step2_status: 'pending' });
    } else if (data.action === 'interests') {
      const values = { ...data, step2_status: data.status || 'completed', updated_at: now };
      const row = findRow_(sheet, data.lead_id);
      if (row) {
        STEP2_FIELDS.concat('updated_at').forEach(function (key) {
          if (values[key] !== undefined) {
            sheet.getRange(row, HEADERS.indexOf(key) + 1).setValue(clean_(values[key]));
          }
        });
      } else {
        appendRow_(sheet, { ...values, created_at: now });
      }
    } else {
      throw new Error('unknown action');
    }
    return json_({ ok: true, lead_id: data.lead_id });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

// Visit the /exec URL in a browser to check the deployment is live.
function doGet() {
  return json_({ ok: true, service: 'mingleby-signups' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const header = sheet.getRange(1, 1, 1, HEADERS.length);
  if (header.getValues()[0].join() !== HEADERS.join()) {
    header.setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // Store mobiles as plain text so Sheets doesn't treat them as numbers.
    if (sheet.getMaxRows() > 1) sheet.getRange(2, HEADERS.indexOf('mobile') + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
  }
  return sheet;
}

function appendRow_(sheet, obj) {
  sheet.appendRow(HEADERS.map(function (h) { return obj[h] === undefined ? '' : clean_(obj[h]); }));
}

function findRow_(sheet, leadId) {
  const cell = sheet.getRange('A:A').createTextFinder(String(leadId)).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : null;
}

// Strings starting with = + - @ would be run as formulas; prefix them with '.
function clean_(v) {
  if (v instanceof Date) return v;
  const s = String(v).slice(0, 1000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
