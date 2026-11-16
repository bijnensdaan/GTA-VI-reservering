/**
 * Google Apps Script backend voor de PlayStation-reserveringsapp.
 * Koppel dit script aan de Google Sheet en voer setup() eenmalig uit.
 */
const BOOKINGS_TAB = 'Reserveringen';
const SETTINGS_TAB = 'Instellingen';
const BOOKING_HEADERS = ['ID', 'Datum', 'Starttijd', 'Duur (min)', 'Aanvragers', 'Winnaar', 'Status', 'Gestart op', 'Afgelopen op'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Koppel dit Apps Script aan jullie Google Sheet.');
  ss.setSpreadsheetTimeZone('Europe/Brussels');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  let bookings = ss.getSheetByName(BOOKINGS_TAB);
  if (!bookings) bookings = ss.insertSheet(BOOKINGS_TAB);
  if (bookings.getLastRow() === 0) bookings.getRange(1, 1, 1, BOOKING_HEADERS.length).setValues([BOOKING_HEADERS]);
  bookings.setFrozenRows(1);
  bookings.getRange(1, 1, 1, BOOKING_HEADERS.length).setFontWeight('bold');
  bookings.getRange('B:B').setNumberFormat('yyyy-mm-dd');
  bookings.getRange('C:C').setNumberFormat('hh:mm');
  dedupeSlots_(bookings);
  let settings = ss.getSheetByName(SETTINGS_TAB);
  if (!settings) settings = ss.insertSheet(SETTINGS_TAB);
  if (settings.getLastRow() === 0) settings.getRange('A1:C2').setValues([['Broer 1', 'Broer 2', 'Broer 3'], ['', '', '']]);
  settings.setFrozenRows(1);
  settings.getRange('A1:C1').setFontWeight('bold');
  return 'Klaar. De tabbladen zijn ingesteld.';
}

function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback;
  const result = readState_();
  const json = JSON.stringify(result);
  if (callback && /^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const action = String((e && e.parameter && e.parameter.action) || '');
  let payload = {};
  try { payload = JSON.parse((e.postData && e.postData.contents) || '{}'); } catch (err) { return output_({ ok: false, error: 'Ongeldige JSON.' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = spreadsheet_();
    const sheet = ss.getSheetByName(BOOKINGS_TAB);
    if (action === 'configure') return configure_(ss, payload);
    if (action === 'reserve') return reserve_(sheet, payload);
    if (action === 'delete') return deleteBooking_(sheet, payload.id);
    if (action === 'start') return changeStatus_(sheet, payload.id, 'started');
    if (action === 'finish') return changeStatus_(sheet, payload.id, 'done');
    if (action === 'clear') {
      if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
      return output_({ ok: true });
    }
    return output_({ ok: false, error: 'Onbekende actie.' });
  } finally { lock.releaseLock(); }
}

function reserve_(sheet, request) {
  const player = String(request.player || '').trim();
  const date = String(request.date || '');
  const time = String(request.time || '');
  const duration = Number(request.duration);
  if (!player || !/^\d{4}-\d{2}-\d{2}$/.test(date) || !/^\d{2}:\d{2}$/.test(time) || ![60, 90, 120, 180].includes(duration)) return output_({ ok: false, error: 'Ongeldige reservering.' });
  const timezone = spreadsheet_().getSpreadsheetTimeZone();
  const startsAt = Utilities.parseDate(date + ' ' + time, timezone, 'yyyy-MM-dd HH:mm');
  if (startsAt.getTime() < Date.now() + 2 * 60 * 60 * 1000) return output_({ ok: false, error: 'Reserveer minstens twee uur van tevoren.' });
  const names = readNames_(spreadsheet_());
  if (!names.includes(player)) return output_({ ok: false, error: 'Naam staat niet in de instellingen.' });
  dedupeSlots_(sheet);
  const all = bookingRows_(sheet);
  let found = null;
  for (const entry of all) if (entry.date === date && entry.time === time) { found = entry; break; }
  if (found && (found.status === 'started' || found.status === 'done')) return output_({ ok: false, error: 'Dit slot kan niet meer worden geloot.' });
  const slotDuration = found ? found.duration : duration;
  const players = found ? found.players.slice() : [];
  if (players.includes(player)) return output_({ ok: false, error: 'Je staat al op de lijst.' });
  const start = toMinutes_(time);
  const collision = all.some(entry => entry.date === date && entry.winner && entry.status !== 'done' && entry.id !== (found && found.id) && start < toMinutes_(entry.time) + entry.duration && toMinutes_(entry.time) < start + slotDuration);
  if (collision) return output_({ ok: false, error: 'Dit tijdslot overlapt met een andere reservering.' });
  players.push(player);
  const winner = players[Math.floor(Math.random() * players.length)];
  if (found) {
    sheet.getRange(found.row, 4, 1, 3).setValues([[slotDuration, JSON.stringify(players), winner]]);
  } else {
    sheet.appendRow([Utilities.getUuid(), date, time, slotDuration, JSON.stringify(players), winner, 'booked', '', '']);
  }
  SpreadsheetApp.flush();
  return output_({ ok: true, winner: winner, players: players });
}

function configure_(ss, payload) {
  const names = Array.isArray(payload.names) ? payload.names.map(name => String(name).trim()) : [];
  if (names.length !== 3 || names.some(name => !name) || new Set(names.map(name => name.toLowerCase())).size !== 3) return output_({ ok: false, error: 'Vul drie verschillende namen in.' });
  ss.getSheetByName(SETTINGS_TAB).getRange(2, 1, 1, 3).setValues([names]);
  return output_({ ok: true });
}

function changeStatus_(sheet, id, status) {
  const row = bookingRows_(sheet).find(entry => entry.id === String(id || ''));
  if (!row) return output_({ ok: false, error: 'Reservering niet gevonden.' });
  if (status === 'started' && row.status !== 'booked') return output_({ ok: false, error: 'Deze reservering is niet meer beschikbaar.' });
  if (status === 'done' && row.status !== 'started') return output_({ ok: false, error: 'Deze sessie is niet gestart.' });
  sheet.getRange(row.row, 7).setValue(status);
  sheet.getRange(row.row, status === 'started' ? 8 : 9).setValue(new Date());
  return output_({ ok: true });
}

function deleteBooking_(sheet, id) {
  const row = bookingRows_(sheet).find(entry => entry.id === String(id || ''));
  if (!row) return output_({ ok: false, error: 'Reservering niet gevonden.' });
  if (row.status !== 'booked') return output_({ ok: false, error: 'Alleen een nog niet gestarte reservering kan worden verwijderd.' });
  sheet.deleteRow(row.row);
  SpreadsheetApp.flush();
  return output_({ ok: true });
}

function readState_() {
  const ss = spreadsheet_();
  const sheet = ss.getSheetByName(BOOKINGS_TAB);
  return { names: readNames_(ss), bookings: bookingRows_(sheet, ss.getSpreadsheetTimeZone()).map(({ row, ...item }) => item) };
}
function bookingRows_(sheet, timezone) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const tz = timezone || spreadsheet_().getSpreadsheetTimeZone();
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, BOOKING_HEADERS.length).getValues().map((r, index) => ({
    row: index + 2, id: String(r[0]), date: normalizeDate_(r[1], tz), time: normalizeTime_(r[2], tz), duration: Number(r[3]),
    players: safePlayers_(r[4]), winner: String(r[5]), status: String(r[6]), started: String(r[6]) === 'started', startedAt: r[7] || '', endedAt: r[8] || ''
  }));
}
function dedupeSlots_(sheet) {
  const rows = bookingRows_(sheet);
  const groupsBySlot = new Map();
  rows.forEach(row => {
    const key = row.date + '|' + row.time;
    if (!groupsBySlot.has(key)) groupsBySlot.set(key, []);
    groupsBySlot.get(key).push(row);
  });
  const plans = [];
  const deleteRows = [];
  groupsBySlot.forEach(group => {
    if (group.length < 2) return;
    group.sort((a, b) => a.row - b.row);
    const players = [...new Set(group.flatMap(row => row.players).filter(Boolean))];
    if (!players.length) return;
    const keep = group[0];
    const statuses = group.map(row => row.status);
    const status = statuses.includes('started') ? 'started' : statuses.includes('done') ? 'done' : 'booked';
    const oldWinners = new Set(group.map(row => row.winner));
    const winner = oldWinners.size === 1 && players.includes(keep.winner)
      ? keep.winner
      : players[Math.floor(Math.random() * players.length)];
    const startedAt = group.find(row => row.status === 'started' && row.startedAt)?.startedAt || keep.startedAt || '';
    const endedAt = status === 'done' ? (group.find(row => row.endedAt)?.endedAt || keep.endedAt || '') : '';
    plans.push({ date: keep.date, time: keep.time, duration: keep.duration, players, winner, status, startedAt, endedAt });
    group.slice(1).forEach(row => deleteRows.push(row.row));
  });
  deleteRows.sort((a, b) => b - a).forEach(row => sheet.deleteRow(row));
  if (plans.length) {
    const remaining = bookingRows_(sheet);
    plans.forEach(plan => {
      const keep = remaining.find(row => row.date === plan.date && row.time === plan.time);
      if (keep) sheet.getRange(keep.row, 4, 1, 6).setValues([[plan.duration, JSON.stringify(plan.players), plan.winner, plan.status, plan.startedAt, plan.endedAt]]);
    });
    SpreadsheetApp.flush();
  }
}
function normalizeDate_(value, timezone) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
  return String(value || '').trim();
}
function normalizeTime_(value, timezone) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, timezone, 'HH:mm');
  if (typeof value === 'number' && isFinite(value)) {
    const total = ((Math.round(value * 1440) % 1440) + 1440) % 1440;
    return String(Math.floor(total / 60)).padStart(2, '0') + ':' + String(total % 60).padStart(2, '0');
  }
  const text = String(value || '').trim();
  const match = text.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  return match ? String(Number(match[1])).padStart(2, '0') + ':' + match[2] : text;
}
function safePlayers_(value) { try { const result = JSON.parse(String(value || '[]')); return Array.isArray(result) ? result : []; } catch (err) { return []; } }
function readNames_(ss) { const values = ss.getSheetByName(SETTINGS_TAB).getRange('A2:C2').getDisplayValues()[0]; return values.map(value => String(value).trim()); }
function spreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Voer eerst setup() uit vanuit de gekoppelde Sheet.');
  return SpreadsheetApp.openById(id);
}
function toMinutes_(time) { const parts = String(time).split(':').map(Number); return parts[0] * 60 + parts[1]; }
function output_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
