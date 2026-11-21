/*
  TOILET PASS CONNECTOR
  Writes Toilet Pass trips into ONE school Google Sheet.

  ONE-TIME SETUP (done once, by one person):
   1. Make the Google Sheet for the logs and copy its link.
   2. Go to script.google.com and click New project.
   3. Delete everything in the editor and paste in this whole file.
   4. Paste your Sheet link between the quotes on the SHEET_URL line below. Click Save.
   5. Click Deploy, then New deployment. Click the gear next to "Select type" and pick Web app.
        Execute as:      Me
        Who has access:  Anyone
      (If "Anyone" isn't in the list, your school blocks it. Ask IT to allow it.)
   6. Click Deploy, then Authorize access, and choose your account. If Google says
      "Google hasn't verified this app", click Advanced, then Go to Toilet Pass.
      That warning appears because this is your own script, not a published app.
   7. Copy the Web app URL (it ends in /exec) and put it in toilet-pass.html
      (the WEB_APP_URL line near the top of the script), or send it to whoever maintains the page.

  This script can only ever write to the one Sheet named below.
  The KEY must match the KEY in toilet-pass.html.
*/
const SHEET_URL = 'PASTE YOUR GOOGLE SHEET LINK HERE';
const KEY = 'e8e5efaa70e253484909a21e';
const HEADERS = ['Date', 'Student', 'Partner', 'Time Out', 'Time Back', 'Minutes Away', 'Trip ID'];
const VERSION = 2;   // Toilet Pass checks this before asking for today's trips

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.key !== KEY) return reply({ ok: false, error: 'The connector key does not match. Ask whoever set up Toilet Pass.' });
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    if (d.action === 'ping') return reply({ ok: true, name: ss.getName(), version: VERSION });
    if (d.action === 'today') return reply({ ok: true, trips: todaysTrips(ss, d.className, d.date) });
    if (['out', 'back', 'cancel'].indexOf(d.action) < 0 || !d.id) return reply({ ok: false, error: 'Unknown request.' });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = getTab(ss, d.className);
      const row = findRow(sh, d.id);
      if (d.action === 'cancel') {
        if (row) sh.deleteRow(row);
      } else if (!row) {
        sh.appendRow([d.date, d.student, d.partner || '', d.timeOut, d.timeBack || '', d.minutes === '' ? '' : Number(d.minutes), d.id]);
        formatRow(sh, sh.getLastRow());
      } else if (d.action === 'back') {
        sh.getRange(row, 5, 1, 2).setValues([[d.timeBack, Number(d.minutes)]]);
        formatRow(sh, row);
      }
    } finally {
      lock.releaseLock();
    }
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err && err.message || err) });
  }
}

// Today's trips for one class tab. Only the most recent rows are read, since today's are at the bottom.
function todaysTrips(ss, className, date) {
  const sh = ss.getSheetByName(tabName(className));
  if (!sh || sh.getLastRow() < 2) return [];
  const last = sh.getLastRow(), first = Math.max(2, last - 299);
  const range = sh.getRange(first, 1, last - first + 1, 7);
  const values = range.getValues(), shown = range.getDisplayValues();
  const tz = ss.getSpreadsheetTimeZone();
  const trips = [];
  values.forEach((row, i) => {
    const day = row[0] instanceof Date ? Utilities.formatDate(row[0], tz, 'yyyy-MM-dd') : String(row[0]);
    if (day !== date || !row[6]) return;
    trips.push({ id: String(row[6]), student: String(row[1]), partner: String(row[2]),
      timeOut: shown[i][3], timeBack: shown[i][4], minutes: row[5] === '' ? '' : Number(row[5]) });
  });
  return trips;
}

// Opening the Web app URL in a browser shows this, which is a quick way to check it's working.
function doGet() {
  let name = '';
  try { name = SpreadsheetApp.openByUrl(SHEET_URL).getName(); } catch (err) { name = 'ERROR: ' + err.message; }
  return reply({ ok: true, message: 'Toilet Pass connector is running.', sheet: name, version: VERSION });
}

function tabName(name) {
  return String(name || 'Class').replace(/[\[\]\*\?\/\\:]/g, ' ').trim().slice(0, 90) || 'Class';
}

function getTab(ss, name) {
  name = tabName(name);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 6, 130);
    sh.setColumnWidth(2, 170);
    sh.setColumnWidth(3, 170);
    sh.hideColumns(7);
    const blank = ss.getSheetByName('Sheet1');
    if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  }
  return sh;
}

function findRow(sh, id) {
  if (!id || sh.getLastRow() < 2) return 0;
  const hit = sh.getRange(2, 7, sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function formatRow(sh, row) {
  sh.getRange(row, 1).setNumberFormat('ddd d mmm yyyy');
  sh.getRange(row, 4, 1, 2).setNumberFormat('h:mm am/pm');
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
