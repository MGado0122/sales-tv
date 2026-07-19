/**
 * MEILS Sales TV — Backend (Google Apps Script Web App)
 * ============================================================
 * This REPLACES the entire current script in your Apps Script project.
 *
 * WHAT'S NEW vs. the current version:
 *   1. doGet() now caps Entries to the last 30 rows (was unbounded).
 *   2. doGet(?tab=History) serves a new "History" tab for yesterday-comparison.
 *   3. logDailyHistory() — run once/day via a time trigger — snapshots that
 *      day's totals per team into the History tab, INCLUDING real hourly
 *      breakdowns (Hourly_09..Hourly_18) now that Column I has real times.
 *   4. onEdit() — a simple trigger that auto-stamps the time of day into
 *      Column I whenever Column A (Date) is filled in on a row.
 *
 * DEPLOYMENT STEPS (you've already done 1–4 for a prior version of this file;
 * only the History columns changed, so you need one extra step this time):
 *   1. Open the Google Sheet → Extensions → Apps Script.
 *   2. Select ALL existing code in the editor and replace it with this file.
 *   3. Deploy → Manage Deployments → click the pencil (Edit) on the existing
 *      deployment → Version: "New version" → Deploy.
 *   4. ⚠️ DELETE THE EXISTING "History" TAB before re-running. Its header
 *      row was written with the OLD 8-column layout; this version writes 18
 *      columns (8 + 10 hourly), so the old header would no longer line up.
 *      Right-click the History tab → Delete. It's safe — it only ever had
 *      today's partial data in it, nothing depends on it yet.
 *   5. Run logDailyHistory once manually (▶ Run button, function dropdown
 *      set to logDailyHistory) to recreate the tab with the new headers.
 *   6. The existing daily trigger (Time-driven, ~11pm) still applies — no
 *      need to recreate it, it just calls this same function name.
 *      onEdit() needs no trigger setup at all — simple triggers run
 *      automatically off whatever's currently saved in the editor.
 * ============================================================
 */

const ENTRIES_SHEET = 'Entries';
const HISTORY_SHEET = 'History';
const HOUR_START = 9;  // 9 AM — must match OPENING_HOUR in index.html
const HOUR_END = 18;   // 6 PM — must match CLOSING_HOUR in index.html
const HOURLY_HEADERS = (function () {
  const h = [];
  for (let i = HOUR_START; i <= HOUR_END; i++) h.push('Hourly_' + String(i).padStart(2, '0'));
  return h;
})();
const HISTORY_HEADERS = ['Date','Team','Total_Entries','Confirmed','Mega','RFQ','Opportunity','Cancelled'].concat(HOURLY_HEADERS);
const TEAM_NAMES = ['Whale','Lion','Tiger','Ali','Meram','Sayeda','Vivian'];
const OVERSEAS_PEOPLE = ['gendia','esraa khaled']; // rolls up to Tiger (team 3), matches frontend logic

function doGet(e) {
  const tab = (e && e.parameter && e.parameter.tab) || 'Entries';
  const sheetName = tab === 'History' ? HISTORY_SHEET : ENTRIES_SHEET;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift();
  let rows = values
    .filter(function (r) { return r.join('').trim() !== ''; })
    .map(function (r) {
      var obj = {};
      headers.forEach(function (h, i) { obj[String(h).trim()] = r[i]; });
      return obj;
    });
  if (sheetName === ENTRIES_SHEET) rows = rows.slice(-30); // cap — keeps the TV payload small forever
  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Run daily (time trigger, ~11:55–11:59 PM) to snapshot today's totals —
 * including real per-hour breakdowns — into the History tab. Creates the
 * tab + headers automatically on first run. Status/team matching mirrors
 * the frontend's normStatus()/teamKey() exactly so numbers never disagree
 * with what the TV showed that day.
 */
function logDailyHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entries = ss.getSheetByName(ENTRIES_SHEET);
  let history = ss.getSheetByName(HISTORY_SHEET);
  if (!history) {
    history = ss.insertSheet(HISTORY_SHEET);
    history.appendRow(HISTORY_HEADERS);
  }

  const values = entries.getDataRange().getDisplayValues();
  const headers = values.shift();
  const rows = values.filter(function (r) { return r.join('').trim() !== ''; })
    .map(function (r) { var o = {}; headers.forEach(function (h,i){ o[String(h).trim()] = r[i]; }); return o; });

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const todayRows = rows.filter(function (r) { return normalizeDate_(r['Date']) === todayStr; });

  for (let t = 1; t <= 7; t++) {
    const teamRows = todayRows.filter(function (r) { return effTeamNum_(r) === t; });
    const count = function (status) {
      return teamRows.filter(function (r) { return normStatus_(r['Status']) === status; }).length;
    };
    // Hourly breakdown counts confirmed+mega only, matching the dashboard's
    // heatmap definition (renderHeatmap is called with winToday = confirmed+mega).
    const winRows = teamRows.filter(function (r) {
      const s = normStatus_(r['Status']);
      return s === 'confirmed' || s === 'mega';
    });
    const hourlyCounts = [];
    for (let h = HOUR_START; h <= HOUR_END; h++) {
      hourlyCounts.push(winRows.filter(function (r) { return parseTimeToHour_(r['Time']) === h; }).length);
    }
    history.appendRow([
      todayStr, TEAM_NAMES[t-1], teamRows.length,
      count('confirmed'), count('mega'), count('rfq'), count('opportunity'), count('cancelled')
    ].concat(hourlyCounts));
  }
}

function normalizeDate_(raw) {
  raw = String(raw || '').trim();
  let m = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return raw.slice(0, 10);
  m = raw.match(/^(\d{1,2})[\/](\d{1,2})[\/](\d{4})/);
  if (m) return m[3] + '-' + ('0'+m[1]).slice(-2) + '-' + ('0'+m[2]).slice(-2);
  return '';
}

function normStatus_(s) {
  s = String(s || '').trim().toLowerCase();
  if (s.indexOf('cancel') > -1) return 'cancelled';
  if (s.indexOf('mega') > -1) return 'mega';
  if (s.indexOf('opportun') > -1 || s === 'opp') return 'opportunity';
  if (s.indexOf('rfq') > -1 || s.indexOf('quote') > -1) return 'rfq';
  if (s.indexOf('confirm') > -1 || s.indexOf('shipment') > -1) return 'confirmed';
  return s;
}

function effTeamNum_(row) {
  const person = String(row['Salesperson'] || '').trim().toLowerCase();
  if (OVERSEAS_PEOPLE.indexOf(person) !== -1) return 3; // Tiger
  let team = String(row['Team'] || '').trim();
  const nameMap = { whale:1, lion:2, tiger:3, ali:4, meram:5, sayeda:6, vivian:7 };
  if (/[a-z]/i.test(team)) return nameMap[team.toLowerCase()] || 1;
  return parseInt(team, 10) || 1;
}

// Parses the "hh:mm am/pm" display format written by onEdit (Column I) into
// a 24-hour integer hour, or null if blank/unparseable (e.g. legacy rows).
function parseTimeToHour_(timeStr) {
  const m = String(timeStr || '').trim().match(/^(\d{1,2}):(\d{2})\s*([AaPp][Mm])$/);
  if (!m) return null;
  let hh = parseInt(m[1], 10) % 12;
  if (/[Pp]/.test(m[3])) hh += 12;
  return hh;
}

/**
 * Simple trigger — Google runs this automatically on every edit to the
 * spreadsheet, no manual trigger setup needed. Stamps the current time into
 * Column I whenever Column A (Date) is filled in on a data row, giving every
 * entry a real timestamp for the first time. Column I was unused before this.
 *
 * Reads column A's actual value per-row (rather than trusting the edit
 * event's e.value) so this works whether a date is typed into a single cell
 * or a whole row is pasted/filled at once — e.value only reports reliably
 * for single-cell edits, which silently broke the first version of this.
 * Also won't overwrite an existing timestamp if the date is corrected later.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== ENTRIES_SHEET) return;

  const rowStart = e.range.getRow();
  const rowEnd = e.range.getLastRow();
  const colStart = e.range.getColumn();
  const colEnd = e.range.getLastColumn();
  const timestampCol = 9; // Column I (Time)

  // Only act if the edited range includes column A, and isn't the header row
  if (colStart <= 1 && colEnd >= 1 && rowEnd > 1) {
    for (let r = Math.max(2, rowStart); r <= rowEnd; r++) {
      const dateValue = sheet.getRange(r, 1).getValue();
      const timeCell = sheet.getRange(r, timestampCol);
      if (dateValue !== '') {
        if (timeCell.getValue() === '') {
          timeCell.setValue(new Date()).setNumberFormat('hh:mm am/pm');
        }
      } else {
        timeCell.clearContent();
      }
    }
  }
}
