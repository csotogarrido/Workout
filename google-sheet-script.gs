/**
 * Home Strength Sessions -> Google Sheet
 * Paste this whole file into Extensions > Apps Script of a blank Google Sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Every "Save session" in the app adds one row to "Sessions" and one row per set to "Sets".
 */

const SESSIONS_HEADER = ['Date', 'Session', 'Name', 'Result', 'Sets as planned', 'Pain after', 'Feel', 'Stability',
  'Check: signs', 'Check: felt', 'Check: pain areas', 'Check: pain level', 'Check: cueca', 'Check: breathing', 'Comments', 'Entry id'];
const SETS_HEADER = ['Date', 'Session', 'Exercise', 'Set', 'Planned', 'Done', 'Unit', 'As planned', 'Entry id'];

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const entry = JSON.parse(e.postData.contents);
    const sessions = sheet_('Sessions', SESSIONS_HEADER);
    const sets = sheet_('Sets', SETS_HEADER);

    // Skip duplicates if the app sends the same session twice.
    const ids = sessions.getLastRow() > 1
      ? sessions.getRange(2, 16, sessions.getLastRow() - 1, 1).getValues().flat()
      : [];
    if (ids.indexOf(entry.id) !== -1) return json_({ ok: true, duplicate: true });

    const date = new Date(entry.date);
    const c = entry.check || {};
    const setList = entry.sets || [];
    const planned = setList.filter(s => s.full).length;
    const steps = String(entry.steps || '0/0').split('/');
    const standard = setList.length ? planned === setList.length : steps[0] === steps[1];

    sessions.appendRow([
      date, entry.session, entry.name, standard ? 'Standard' : 'Adapted',
      setList.length ? planned + '/' + setList.length : (entry.steps || ''),
      entry.pain, entry.feel || '', entry.stab || '',
      c.flag || '', c.felt || '', c.noPain ? 'None' : (c.areas || []).join(', '), c.pain || '',
      c.dance || '', c.breath || '', entry.notes || '', entry.id
    ]);
    setList.forEach(s => sets.appendRow([
      date, entry.session, s.ex, s.set, s.plan, s.done === null ? 'not done' : s.done, s.unit, s.full ? 'Yes' : 'No', entry.id
    ]));
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, message: 'Home Strength Sessions sheet is connected.' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
