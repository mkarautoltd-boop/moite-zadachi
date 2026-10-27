/* =====================================================================
   МОИТЕ ЗАДАЧИ — сървър (Google Apps Script)
   Пазиш данните в тази Google таблица, синхронизира календар
   „Моите задачи“ (известия на телефона) и праща сутрешно резюме.
   ===================================================================== */

// ▼▼▼ СМЕНИ ПИН КОДА (с него влизаш в приложението) ▼▼▼
const PIN = '1234';
// ▲▲▲

const CAL_NAME = 'Моите задачи';
const APP_URL  = 'https://mkarautoltd-boop.github.io/moite-zadachi/';
const SEND_BRIEF_HOUR = 7;            // ежедневна проверка на наемите в 7:00
const SEND_BRIEF = false;             // true = сутрешно резюме по имейл пон–пет

const COLS  = ['items','ledger','adv','meets','plan','rents','wexp','settle','rentPaid','cfg'];
const SHEET = { items:'Записи', ledger:'Сервизи и доставчици', adv:'Касата ми дължи', meets:'Срещи', plan:'План',
  rents:'Наеми', wexp:'Автомивка разходи', settle:'Автомивка изравняване', rentPaid:'Платени наеми', cfg:'Настройки' };
const EV_SHEET = 'Събития';

/* ---------- уеб API ---------- */
function doGet() { return out({ ok: true, app: 'Моите задачи' }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, err: 'bad request' }); }
  if (String(req.pin) !== String(PIN)) return out({ ok: false, err: 'pin' });
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (req.ops && req.ops.length) applyOps(req.ops);
    return out({ ok: true, data: readAll(), t: Date.now() });
  } catch (err) {
    return out({ ok: false, err: String(err && err.message || err) });
  } finally { lock.releaseLock(); }
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- таблица ---------- */
function sheetOf(c) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(SHEET[c]);
  if (!s) { s = ss.insertSheet(SHEET[c]); s.getRange(1, 1, 1, 4).setValues([['id', 'u', 'del', 'данни']]).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function readRows(c) {
  const s = sheetOf(c), n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, 4).getValues().map(r => ({ id: String(r[0]), u: Number(r[1]) || 0, del: r[2] === true || r[2] === 'TRUE', json: String(r[3] || '') }));
}
function parse(j) { try { return JSON.parse(j) || {}; } catch (e) { return {}; } }
function readCol(c) { return readRows(c).filter(r => !r.del && r.id).map(r => Object.assign(parse(r.json), { id: r.id })); }
function readAll() { const d = {}; COLS.forEach(c => d[c] = readCol(c)); return d; }

function applyOps(ops) {
  const byCol = {};
  ops.forEach(o => { if (COLS.indexOf(o.c) >= 0 && o.id) (byCol[o.c] = byCol[o.c] || []).push(o); });
  const changed = [];
  Object.keys(byCol).forEach(c => {
    const s = sheetOf(c), rows = readRows(c), idx = {};
    rows.forEach((r, i) => idx[r.id] = i);
    byCol[c].forEach(o => {
      const i = idx[o.id];
      if (i !== undefined && rows[i].u > o.u) return;               // на сървъра има по-нова версия
      const row = { id: o.id, u: o.u, del: !!o.del, json: o.del ? '' : o.json };
      if (i === undefined) { idx[o.id] = rows.length; rows.push(row); } else rows[i] = row;
      changed.push({ c: c, id: o.id, del: !!o.del, obj: o.del ? null : Object.assign(parse(o.json), { id: o.id }) });
    });
    // чистим стари изтрити редове (над 60 дни)
    const old = Date.now() - 60 * 864e5;
    const keep = rows.filter(r => !(r.del && r.u < old));
    s.getRange(2, 1, Math.max(s.getLastRow() - 1, 1), 4).clearContent();
    if (keep.length) s.getRange(2, 1, keep.length, 4).setValues(keep.map(r => [r.id, r.u, r.del, r.json]));
  });
  try { syncCalendar(changed); } catch (err) { console.error('Календар: ' + err); }
}

/* ---------- логика (същата като в приложението) ---------- */
const DEPTS = { obsh: 'Община', babh: 'Агенция храни', kzp: 'КЗП', pol: 'Полиция', sad: 'Съд', nap: 'НАП' };
const EXPIRY = ['Удължено работно време', 'Тротоарно право', 'Табела'];
const isFiled = i => i.stage === 'filed' && !i.done;
const waitNP  = i => i.kind === 'АУАН' && !i.np && !i.done;
const isFine  = i => !i.done && ((i.kind === 'АУАН' && i.np) || i.kind === 'Споразумение');
const fineAmt = i => i.kind === 'АУАН' ? i.fine : i.amount;
const effDate = i => isFiled(i) ? (i.pickup || i.date) : (i.kind === 'АУАН' && i.np ? i.payDue : i.date);
const defDays = i => EXPIRY.indexOf(i.kind) >= 0 ? 30 : (['Предписание', 'АУАН', 'Споразумение'].indexOf(i.kind) >= 0 ? 7 : 3);
const eur = n => (Math.round((+n || 0) * 100) / 100).toFixed(2).replace('.', ',') + ' €';
const bg  = s => s ? s.slice(8, 10) + '.' + s.slice(5, 7) + '.' + s.slice(2, 4) : '';
function todayIso() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function dt(iso, hm) { const p = iso.split('-').map(Number), t = (hm || '09:00').split(':').map(Number); return new Date(p[0], p[1] - 1, p[2], t[0] || 0, t[1] || 0); }
function diffDays(iso) { const a = dt(todayIso(), '00:00'), b = dt(iso, '00:00'); return Math.round((b - a) / 864e5); }
function ym(iso) { return iso.slice(0, 7); }

/* ---------- календар ---------- */
function cal() {
  const c = CalendarApp.getCalendarsByName(CAL_NAME);
  if (c.length) return c[0];
  const n = CalendarApp.createCalendar(CAL_NAME, { color: CalendarApp.Color.RED, summary: 'Срокове, проверки и срещи от приложението „Моите задачи“' });
  return n;
}
function evMap() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(EV_SHEET);
  if (!s) { s = ss.insertSheet(EV_SHEET); s.getRange(1, 1, 1, 2).setValues([['ключ', 'събитие']]).setFontWeight('bold'); s.hideSheet(); }
  const n = s.getLastRow(), m = {};
  if (n > 1) s.getRange(2, 1, n - 1, 2).getValues().forEach(r => { if (r[0]) m[r[0]] = r[1]; });
  return { s: s, m: m };
}
function saveEvMap(em) {
  const rows = Object.keys(em.m).filter(k => em.m[k]).map(k => [k, em.m[k]]);
  em.s.getRange(2, 1, Math.max(em.s.getLastRow() - 1, 1), 2).clearContent();
  if (rows.length) em.s.getRange(2, 1, rows.length, 2).setValues(rows);
}
const clampR = m => Math.max(0, Math.min(40320, Math.round(m)));   // Google позволява до 4 седмици

function itemSpec(i) {
  if (!i || i.done || waitNP(i)) return null;
  const d = effDate(i); if (!d) return null;
  const dep = DEPTS[i.dept] || '';
  let title, time = '09:00', rem;
  if (isFiled(i)) { title = 'Община: Вземи разрешителното (' + i.kind + ')'; rem = [1440, 60]; }
  else if (isFine(i)) { title = dep + ': Плати ' + (i.kind === 'АУАН' ? 'глоба' : 'споразумение') + ' ' + eur(fineAmt(i)) + ' — ' + (i.place || i.title); rem = [i.remind || 7 * 1440, 60]; }
  else if (EXPIRY.indexOf(i.kind) >= 0) { title = dep + ': Изтича ' + i.kind + ' — ' + (i.place || i.title); rem = [i.remind || 30 * 1440, 1440]; }
  else { title = dep + ': ' + i.kind + ' — ' + i.title; time = i.time || '09:00'; rem = [i.remind || defDays(i) * 1440, 60]; }
  const lines = [i.kind, i.title, i.store && 'Магазин: ' + i.store, i.place && (i.dept === 'pol' ? 'Районно: ' : 'Обект: ') + i.place, i.muni,
    i.firm && 'Фирма: ' + i.firm, i.dept === 'pol' ? (i.poa ? 'Пълномощно: готово' : '⚠ НЯМА ПЪЛНОМОЩНО') : '', i.note && 'Бележка: ' + i.note, '', APP_URL];
  return { title: title + (i.dept === 'pol' && !i.poa ? ' ⚠ пълномощно' : ''), start: dt(d, time), min: 30, rem: rem, desc: lines.filter(x => x !== undefined && x !== null && x !== false).join('\n') };
}
function meetSpec(m) {
  if (!m || m.done || !m.date) return null;
  return { title: (m.title || 'Среща'), start: dt(m.date, m.time || '09:00'), min: 60, rem: m.remind ? [m.remind] : [],
    location: m.place || '', desc: [m.place, m.note, '', APP_URL].filter(x => x !== undefined).join('\n') };
}
function putEvent(c, em, key, spec) {
  const old = em.m[key];
  let ev = null;
  if (old) { try { ev = c.getEventById(old); } catch (e) { ev = null; } }
  if (!spec) { if (ev) try { ev.deleteEvent(); } catch (e) {} delete em.m[key]; return; }
  const end = new Date(spec.start.getTime() + spec.min * 60000);
  if (ev) { ev.setTitle(spec.title); ev.setTime(spec.start, end); ev.setDescription(spec.desc || ''); if (spec.location !== undefined) ev.setLocation(spec.location); }
  else ev = c.createEvent(spec.title, spec.start, end, { description: spec.desc || '', location: spec.location || '' });
  ev.removeAllReminders();
  (spec.rem || []).map(clampR).filter((v, k, a) => a.indexOf(v) === k).slice(0, 5).forEach(r => ev.addPopupReminder(r));
  em.m[key] = ev.getId();
}
function syncCalendar(changed) {
  if (!changed.length) return;
  const c = cal(), em = evMap();
  let rents = false;
  changed.forEach(x => {
    if (x.c === 'items') putEvent(c, em, 'i:' + x.id, x.del ? null : itemSpec(x.obj));
    else if (x.c === 'meets') putEvent(c, em, 'm:' + x.id, x.del ? null : meetSpec(x.obj));
    else if (x.c === 'rents' || x.c === 'rentPaid') rents = true;
  });
  if (rents) syncRents(c, em);
  saveEvMap(em);
}
function syncRents(c, em) {
  const rents = readCol('rents'), paid = {};
  readCol('rentPaid').forEach(p => paid[p.id] = p.date);
  const t = todayIso(), months = [ym(t)];
  const nx = dt(t, '12:00'); nx.setMonth(nx.getMonth() + 1, 1); months.push(Utilities.formatDate(nx, Session.getScriptTimeZone(), 'yyyy-MM'));
  const want = {};
  rents.forEach(r => months.forEach(m => {
    const p = m.split('-').map(Number), dim = new Date(p[0], p[1], 0).getDate();
    const due = m + '-' + ('0' + Math.min(+r.day || 1, dim)).slice(-2), key = 'r:' + r.id + '|' + m;
    if (m !== ym(t) && diffDays(due) > 10) return;                    // следващия месец — само ако наближава
    want[key] = paid[r.id + '|' + m] ? null : { title: 'Наем: ' + r.obj + ' — ' + eur(r.amount), start: dt(due, '09:00'), min: 30, rem: [5 * 1440, 60], desc: (r.who || '') + '\n\n' + APP_URL };
  }));
  Object.keys(em.m).filter(k => k.indexOf('r:') === 0 && !(k in want)).forEach(k => want[k] = null);
  Object.keys(want).forEach(k => putEvent(c, em, k, want[k]));
}
function syncAll() {
  const c = cal(), em = evMap();
  readCol('items').forEach(i => putEvent(c, em, 'i:' + i.id, itemSpec(i)));
  readCol('meets').forEach(m => putEvent(c, em, 'm:' + m.id, meetSpec(m)));
  const ids = {}; readCol('items').forEach(i => ids['i:' + i.id] = 1); readCol('meets').forEach(m => ids['m:' + m.id] = 1);
  Object.keys(em.m).filter(k => /^[im]:/.test(k) && !ids[k]).forEach(k => putEvent(c, em, k, null));
  syncRents(c, em);
  saveEvMap(em);
}

/* ---------- сутрешно резюме ---------- */
function briefText() {
  const t = todayIso(), items = readCol('items'), meets = readCol('meets'), plan = readCol('plan'), rents = readCol('rents');
  const paid = {}; readCol('rentPaid').forEach(p => paid[p.id] = p.date);
  const up = [];
  items.filter(i => !i.done && !waitNP(i)).forEach(i => { const e = effDate(i), d = diffDays(e), lim = isFiled(i) ? 3 : (i.remind ? Math.ceil(i.remind / 1440) : defDays(i)); if (d <= lim) up.push({ d: d, s: line(i, e) }); });
  meets.filter(m => !m.done).forEach(m => { const d = diffDays(m.date); if (d <= 3) up.push({ d: d, s: '• ' + (m.time ? m.time + ' ' : '') + m.title + (m.place ? ' — ' + m.place : '') + ' (' + bg(m.date) + ')' }); });
  rents.forEach(r => { if (paid[r.id + '|' + ym(t)]) return; const p = t.split('-').map(Number), dim = new Date(p[0], p[1], 0).getDate(), due = ym(t) + '-' + ('0' + Math.min(+r.day || 1, dim)).slice(-2), d = diffDays(due); if (d <= 5) up.push({ d: d, s: '• Наем: ' + r.obj + ' — ' + eur(r.amount) + ' (до ' + bg(due) + ')' }); });
  up.sort((a, b) => a.d - b.d);
  const D = new Date(), dn = ['неделя', 'понеделник', 'вторник', 'сряда', 'четвъртък', 'петък', 'събота'][D.getDay()];
  let s = 'Добро утро! ' + dn + ', ' + bg(t) + '\n\n';
  [['ПРОСРОЧЕНО', u => u.d < 0], ['ДНЕС И УТРЕ', u => u.d >= 0 && u.d <= 1], ['ПРЕДСТОИ', u => u.d > 1]].forEach(g => { const a = up.filter(g[1]); if (a.length) s += g[0] + '\n' + a.map(u => u.s).join('\n') + '\n\n'; });
  const pt = plan.filter(p => !p.done && p.date <= t);
  if (pt.length) s += 'ПЛАН ЗА ДНЕС\n' + pt.map(p => '☐ ' + p.text + (p.date < t ? ' (от ' + bg(p.date).slice(0, 5) + ')' : '')).join('\n') + '\n\n';
  const owed = readCol('adv').filter(a => !a.back).reduce((x, a) => x + (+a.amount || 0), 0);
  const bal = {}; readCol('ledger').forEach(l => bal[l.party] = (bal[l.party] || 0) + (l.type === 'bill' ? +l.amount : -l.amount));
  s += 'ПАРИ\n• Касата ми дължи: ' + eur(owed) + '\n' + Object.keys(bal).filter(k => bal[k] > 0.005).map(k => '• Дължа на ' + k + ': ' + eur(bal[k])).join('\n');
  return s + '\n\n' + APP_URL;
}
function line(i, e) {
  if (isFiled(i)) return '• Вземи разрешително (' + i.kind + ') — ' + (i.place || i.title) + ' (' + bg(e) + ')';
  if (isFine(i)) return '• ' + DEPTS[i.dept] + ': плати ' + (i.kind === 'АУАН' ? 'глоба' : 'споразумение') + ' ' + eur(fineAmt(i)) + ' — ' + (i.place || i.title) + ' (до ' + bg(e) + ')';
  return '• ' + DEPTS[i.dept] + ': ' + i.title + (i.place ? ' — ' + i.place : '') + (i.time ? ' ' + i.time : '') + ' (' + bg(e) + ')' + (i.dept === 'pol' && !i.poa ? ' ⚠ пълномощно!' : '');
}

/* ---------- всеки ден ---------- */
function daily() {
  try { const c = cal(), em = evMap(); syncRents(c, em); saveEvMap(em); } catch (e) { console.error(e); }
  const wd = new Date().getDay();
  if (SEND_BRIEF && wd >= 1 && wd <= 5) MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Моите задачи — ' + bg(todayIso()), briefText());
}
function testBrief() { MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Моите задачи — проба', briefText()); }

/* ---------- ПЪРВОНАЧАЛНА НАСТРОЙКА: пусни веднъж ---------- */
function setup() {
  COLS.forEach(sheetOf);
  const def = SpreadsheetApp.getActive().getSheetByName('Sheet1') || SpreadsheetApp.getActive().getSheetByName('Лист1');
  if (def && SpreadsheetApp.getActive().getSheets().length > 1) SpreadsheetApp.getActive().deleteSheet(def);
  cal(); evMap();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'daily').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('daily').timeBased().everyDays(1).atHour(SEND_BRIEF_HOUR).nearMinute(0).create();
  syncAll();
  console.log('Готово. Календар „' + CAL_NAME + '“ е създаден, резюмето ще идва в ' + SEND_BRIEF_HOUR + ':00 пон–пет.');
}
