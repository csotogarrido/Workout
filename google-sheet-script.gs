/**
 * Home Strength Sessions -> Google Sheet
 * Paste this whole file into Extensions > Apps Script of a blank Google Sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Every "Save session" in the app adds one row to "Sessions" and one row per set to "Sets".
 */

// Pick your own private code here. The app asks for the same code once per device
// before it can read your history back. Without it, nobody can read the log.
const SYNC_CODE = 'CHANGE-ME';

const SESSIONS_HEADER = ['Date', 'Session', 'Name', 'Result', 'Sets as planned', 'Pain after', 'Feel', 'Stability',
  'Check: signs', 'Check: felt', 'Check: pain areas', 'Check: pain level', 'Check: cueca', 'Check: breathing', 'Comments', 'Entry id', 'Edited'];
const SETS_HEADER = ['Date', 'Session', 'Exercise', 'Set', 'Planned', 'Done', 'Unit', 'As planned', 'Entry id'];

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  // Add any header cells a newer version of this script introduced.
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
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
    const found = ids.indexOf(entry.id);
    if (found !== -1) {
      // Same session sent again: ignore it, unless it was edited in the app, then replace the old rows.
      if (!entry.updated) return json_({ ok: true, duplicate: true });
      sessions.deleteRow(found + 2);
      const setIds = sets.getLastRow() > 1 ? sets.getRange(2, 9, sets.getLastRow() - 1, 1).getValues().flat() : [];
      for (let i = setIds.length - 1; i >= 0; i--) if (setIds[i] === entry.id) sets.deleteRow(i + 2);
    }

    const date = new Date(entry.date);
    const c = entry.check || {};
    const setList = entry.sets || [];
    const planned = setList.filter(s => s.full).length;
    const steps = String(entry.steps || '0/0').split('/');
    const standard = setList.length ? planned === setList.length : steps[0] === steps[1];

    sessions.appendRow([
      date, entry.session, entry.name, standard ? 'Standard' : 'Adapted',
      "'" + (setList.length ? planned + '/' + setList.length : (entry.steps || '')),  // leading ' keeps Sheets from turning 6/6 into a date
      entry.pain, entry.feel || '', entry.stab || '',
      c.flag || '', c.felt || '', c.noPain ? 'None' : (c.areas || []).join(', '), c.pain || '',
      c.dance || '', c.breath || '', entry.notes || '', entry.id, entry.updated ? new Date(entry.updated) : ''
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

function doGet(e) {
  if (!e || !e.parameter || e.parameter.action !== 'list') {
    return json_({ ok: true, message: 'Home Strength Sessions sheet is connected.' });
  }
  if (SYNC_CODE === 'CHANGE-ME' || e.parameter.key !== SYNC_CODE) {
    return json_({ ok: false, error: 'wrong code' });
  }
  // Rebuild every saved session so the app shows the same history on every device.
  const sessions = sheet_('Sessions', SESSIONS_HEADER);
  const sets = sheet_('Sets', SETS_HEADER);
  const sRows = sessions.getLastRow() > 1 ? sessions.getRange(2, 1, sessions.getLastRow() - 1, 17).getValues() : [];
  const sShown = sessions.getLastRow() > 1 ? sessions.getRange(2, 1, sessions.getLastRow() - 1, 17).getDisplayValues() : [];
  const tRows = sets.getLastRow() > 1 ? sets.getRange(2, 1, sets.getLastRow() - 1, 9).getValues() : [];
  const setsById = {};
  tRows.forEach(r => {
    const id = r[8]; if (!id) return;
    (setsById[id] = setsById[id] || []).push({
      ex: r[2], set: Number(r[3]), plan: Number(r[4]),
      done: r[5] === 'not done' || r[5] === '' ? null : Number(r[5]), unit: r[6], full: r[7] === 'Yes'
    });
  });
  const iso = v => (v instanceof Date ? v : new Date(v)).toISOString();
  const entries = sRows.map((r, i) => [r, sShown[i]]).filter(p => p[0][15]).map(([r, shown]) => {
    const id = String(r[15]);
    const areas = r[10] && r[10] !== 'None' ? String(r[10]).split(', ') : [];
    const entry = {
      id: id, date: iso(r[0]), session: String(r[1]), name: String(r[2]),
      pain: r[5] === '' ? null : Number(r[5]), feel: r[6] || null, stab: r[7] || null, notes: r[14] || '',
      check: r[8] ? { flag: r[8], felt: r[9] || null, areas: areas, noPain: r[10] === 'None', pain: r[11] || null, dance: r[12] || null, breath: r[13] || null } : { areas: [] },
      updated: r[16] ? iso(r[16]) : null, synced: true
    };
    if (setsById[id]) entry.sets = setsById[id]; else entry.steps = shown[4];
    return entry;
  });
  return json_({ ok: true, entries: entries });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
