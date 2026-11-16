const KEY = 'ps-reserveringen-v1';
const $ = (selector) => document.querySelector(selector);
const nameForm = $('#names-form');
const bookingForm = $('#booking-form');
const dialog = $('#result-dialog');
const state = loadState();
let toastTimer;
const API_URL = String(window.PS_API_URL || '').trim();
const isShared = Boolean(API_URL);

function loadState() {
  try { return JSON.parse(localStorage.getItem(KEY)) || { names: ['', '', ''], bookings: [] }; }
  catch { return { names: ['', '', ''], bookings: [] }; }
}
function save() { localStorage.setItem(KEY, JSON.stringify(state)); }
function dateString(date) { return [date.getFullYear(), String(date.getMonth()+1).padStart(2,'0'), String(date.getDate()).padStart(2,'0')].join('-'); }
function seedInputs() {
  nameForm.querySelectorAll('input').forEach((input, i) => input.value = state.names[i] || '');
  const today = dateString(new Date());
  $('#date').value = today;
  $('#filter-date').value = today;
  const timeSelect = $('#time');
  timeSelect.innerHTML = '';
  for (let hour = 8; hour <= 23; hour++) {
    const option = document.createElement('option');
    option.value = `${String(hour).padStart(2,'0')}:00`;
    option.textContent = option.value;
    timeSelect.append(option);
  }
  updatePlayers();
  render();
  if (isShared) {
    $('#storage-tag').textContent = 'Verbinden…';
    $('#storage-note').textContent = 'Verbinding maken met de gedeelde Google Sheet…';
    refreshShared().catch(() => { setConnectionError(); showToast('Google Sheet niet bereikbaar. Probeer Ververs.'); });
  }
}
function setConnectionReady() {
  $('#storage-tag').textContent = 'Google Sheet verbonden';
  $('#storage-tag').classList.add('connected');
  $('#storage-tag').classList.remove('disconnected');
  $('#storage-note').textContent = 'Reserveringen en lotingen worden gedeeld via Google Sheet.';
}
function setConnectionError() {
  $('#storage-tag').textContent = 'Sheet niet bereikbaar';
  $('#storage-tag').classList.remove('connected');
  $('#storage-tag').classList.add('disconnected');
  $('#storage-note').textContent = 'De gedeelde Sheet reageert niet. Controleer je verbinding of probeer Ververs.';
}
function jsonp(url) {
  return new Promise((resolve, reject) => {
    const callback = `__ps_cb_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const script = document.createElement('script');
    const timeout = setTimeout(() => finish(new Error('timeout')), 10000);
    function finish(error, value) {
      clearTimeout(timeout); delete window[callback]; script.remove();
      error ? reject(error) : resolve(value);
    }
    window[callback] = value => finish(null, value);
    script.onerror = () => finish(new Error('network'));
    script.src = `${API_URL}${API_URL.includes('?') ? '&' : '?'}action=get&callback=${callback}`;
    document.body.append(script);
  });
}
async function refreshShared() {
  if (!isShared) return;
  const remote = await jsonp(API_URL);
  if (!remote || !Array.isArray(remote.bookings) || !Array.isArray(remote.names)) throw new Error('Invalid response');
  state.names = remote.names;
  state.bookings = remote.bookings;
  setConnectionReady();
  save(); updatePlayers(); render();
  if (!nameForm.contains(document.activeElement)) nameForm.querySelectorAll('input').forEach((input, i) => input.value = state.names[i] || '');
}
async function postShared(action, data = {}) {
  const response = await fetch(`${API_URL}${API_URL.includes('?') ? '&' : '?'}action=${encodeURIComponent(action)}`, {
    method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(data)
  });
  // Apps Script's cross-origin POST response is opaque; read back the authoritative sheet state.
  await new Promise(resolve => setTimeout(resolve, 1200));
  await refreshShared();
  return response;
}
function updatePlayers() {
  const select = $('#player');
  const previous = select.value;
  select.innerHTML = '';
  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = state.names.every(Boolean) ? 'Kies je naam' : 'Vul eerst de namen in';
  select.append(placeholder);
  state.names.filter(Boolean).forEach(name => {
    const option = document.createElement('option'); option.value = name; option.textContent = name; select.append(option);
  });
  if (state.names.includes(previous)) select.value = previous;
}
function showToast(message) {
  const toast = $('#toast'); toast.textContent = message; toast.classList.add('show');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => toast.classList.remove('show'), 2800);
}
function slotBookings(date, time) { return state.bookings.find(item => item.date === date && item.time === time); }
function handleBooking(event) {
  event.preventDefault();
  if (isShared) { handleSharedBooking(); return; }
  if (state.names.filter(Boolean).length !== 3) return showToast('Vul eerst de drie namen in.');
  const player = $('#player').value, date = $('#date').value, time = $('#time').value, duration = Number($('#duration').value);
  if (!player || !date) return showToast('Kies je naam en een datum.');
  if (!isAtLeastTwoHoursAhead(date, time)) return showToast('Je moet minstens 2 uur van tevoren reserveren.');
  const existing = slotBookings(date, time);
  if (existing?.started) return showToast('Dit slot is al begonnen en kan niet meer worden geloot.');
  if (existing?.players.includes(player)) return showToast(`${player} staat al op de lijst voor dit slot.`);
  const overlaps = state.bookings.find(item => item.date === date && item.time !== time && item.winner &&
    item.status !== 'done' && minutes(time) < minutes(item.time) + item.duration && minutes(item.time) < minutes(time) + duration);
  if (overlaps) return showToast(`Dit slot overlapt met ${overlaps.winner} om ${overlaps.time}.`);

  if (!existing) {
    state.bookings.push({ id: crypto.randomUUID(), date, time, duration, players: [player], winner: player, status: 'booked', started: false });
    save(); $('#filter-date').value = date; render(); showToast(`${player} heeft het slot gereserveerd.`); bookingForm.reset(); $('#date').value = date; return;
  }
  existing.players.push(player);
  const previousWinner = existing.winner;
  existing.winner = existing.players[Math.floor(Math.random() * existing.players.length)];
  save(); $('#filter-date').value = date; render();
  const isWinner = existing.winner === player;
  showResult(isWinner, existing.winner, existing.players.length, player, previousWinner);
  bookingForm.reset(); $('#date').value = date;
}
async function handleSharedBooking() {
  if (state.names.filter(Boolean).length !== 3) return showToast('Vul eerst de drie namen in.');
  const player = $('#player').value, date = $('#date').value, time = $('#time').value, duration = Number($('#duration').value);
  if (!player || !date) return showToast('Kies je naam en een datum.');
  if (!isAtLeastTwoHoursAhead(date, time)) return showToast('Je moet minstens 2 uur van tevoren reserveren.');
  const previous = slotBookings(date, time), oldWinner = previous?.winner || '';
  try {
    await postShared('reserve', { player, date, time, duration });
    const slot = slotBookings(date, time);
    $('#filter-date').value = date; render();
    if (!slot || !slot.players.includes(player)) return showToast('Reservering geweigerd: het slot kan niet meer worden geboekt.');
    if (slot.players.length === 1) showToast(`${player} heeft het slot gereserveerd.`);
    else showResult(slot.winner === player, slot.winner, slot.players.length, player, oldWinner);
    bookingForm.reset(); $('#date').value = date;
  } catch { showToast('Opslaan mislukt. Controleer de Apps Script-URL en verbinding.'); }
}
function minutes(time) { const [hours, mins] = time.split(':').map(Number); return hours * 60 + mins; }
function isAtLeastTwoHoursAhead(date, time) {
  const start = new Date(`${date}T${time}:00`);
  return Number.isFinite(start.getTime()) && start.getTime() >= Date.now() + 2 * 60 * 60 * 1000;
}
function showResult(won, winner, players, requester, previousWinner) {
  $('#result-icon').textContent = won ? '✦' : '↗';
  $('#result-title').textContent = won ? `${requester}, jij hebt gewonnen!` : `${winner} krijgt dit tijdslot`;
  $('#result-copy').textContent = won
    ? `De loting is opnieuw gedaan met ${players} spelers. Jij hebt ${Math.round(100 / players)}% kans gehad en mag dit slot reserveren.`
    : `De loting is opnieuw gedaan met ${players} spelers. ${winner} mag dit slot reserveren. Je kans was ${Math.round(100 / players)}%.`;
  dialog.showModal();
}
function render() {
  const date = $('#filter-date').value;
  const list = $('#reservation-list'); list.innerHTML = '';
  const rows = state.bookings.filter(item => !date || item.date === date).sort((a,b) => a.time.localeCompare(b.time));
  $('#empty-state').classList.toggle('hidden', rows.length > 0);
  rows.forEach(item => {
    const card = document.createElement('article'); card.className = 'reservation';
    const time = document.createElement('div'); time.className = 'time-block'; time.textContent = item.time;
    const info = document.createElement('div');
    const title = document.createElement('h3'); title.textContent = item.winner;
    const detail = document.createElement('p'); detail.textContent = `${item.duration / 60} uur${item.players.length > 1 ? ` · loting met ${item.players.join(', ')}` : ''}`;
    info.append(title, detail);
    const controls = document.createElement('div');
    const status = document.createElement('span'); status.className = 'status';
    status.textContent = item.started ? 'Bezig' : item.status === 'done' ? 'Afgelopen' : 'Gereserveerd';
    if (item.started) status.classList.add('started');
    controls.append(status);
    if (!item.started && item.status !== 'done') {
      const start = document.createElement('button'); start.className = 'action'; start.textContent = 'Ik begin nu';
      start.addEventListener('click', () => confirmStart(item.id)); controls.append(start);
    }
    if (item.started) {
      const finish = document.createElement('button'); finish.className = 'action'; finish.textContent = 'Sessie klaar';
      finish.addEventListener('click', () => finishSession(item.id)); controls.append(finish);
    }
    card.append(time, info, controls); list.append(card);
  });
}
function confirmStart(id) {
  const item = state.bookings.find(booking => booking.id === id);
  if (!item || !confirm(`Bevestig: ${item.winner} begint nu met spelen?`)) return;
  if (isShared) { postShared('start', { id }).then(() => showToast(`Veel plezier, ${item.winner}!`)).catch(() => showToast('Startbevestiging mislukt.')); return; }
  item.started = true; item.status = 'started'; item.startedAt = new Date().toISOString(); save(); render();
  showToast(`Veel plezier, ${item.winner}!`);
}
function finishSession(id) {
  const item = state.bookings.find(booking => booking.id === id);
  if (!item || !confirm(`Markeer de sessie van ${item.winner} als afgelopen?`)) return;
  if (isShared) { postShared('finish', { id }).then(() => showToast('Sessie afgerond.')).catch(() => showToast('Kon sessie niet afronden.')); return; }
  item.status = 'done'; item.endedAt = new Date().toISOString(); save(); render(); showToast('Sessie afgerond.');
}
nameForm.addEventListener('submit', async event => {
  event.preventDefault();
  const names = [...nameForm.querySelectorAll('input')].map(input => input.value.trim());
  if (new Set(names.map(name => name.toLowerCase())).size !== 3) return showToast('Gebruik drie verschillende namen.');
  if (state.bookings.length && state.names.some((name, i) => name !== names[i]) && !confirm('Namen wijzigen? Bestaande reserveringen blijven staan met hun huidige naam.')) return;
  if (isShared) {
    try { await postShared('configure', { names }); showToast('Namen opgeslagen in de gedeelde Sheet.'); }
    catch { showToast('Namen konden niet worden opgeslagen. Controleer je verbinding.'); }
  } else { state.names = names; save(); updatePlayers(); showToast('Namen opgeslagen.'); }
});
bookingForm.addEventListener('submit', handleBooking);
$('#filter-date').addEventListener('change', render);
$('#refresh-button').addEventListener('click', () => isShared ? refreshShared().then(() => showToast('Overzicht bijgewerkt.')).catch(() => { setConnectionError(); showToast('Verversen mislukt.'); }) : render());
$('#close-dialog').addEventListener('click', () => dialog.close());
$('#result-ok').addEventListener('click', () => dialog.close());
$('#clear-button').addEventListener('click', async () => {
  if (!state.bookings.length) return showToast('Er zijn geen reserveringen om te wissen.');
  if (!confirm('Alle reserveringen wissen?')) return;
  if (isShared) {
    try { await postShared('clear'); showToast('Alle reserveringen gewist.'); }
    catch { showToast('Wissen mislukt.'); }
    return;
  }
  state.bookings = []; save(); render(); showToast('Alle reserveringen gewist.');
});
seedInputs();
if (isShared) setInterval(() => refreshShared().catch(() => {}), 20000);
