/* PREVIEW ONLY. When config.js has no SCRIPT_URL, this runs the real apps-script/Code.gs inside the browser
   against pretend Google services, with made-up example data. It does nothing on the live page. */
(function () {
  if ((window.FAMCAL_CONFIG || {}).SCRIPT_URL) return;
  const TZ = 'Europe/London';

  // ---------- time ----------
  function parts(d, tz) {
    const f = new Intl.DateTimeFormat('en-GB', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' });
    const o = {}; f.formatToParts(d).forEach(p => { o[p.type] = p.value; }); return o;
  }
  function offset(ms, tz) { const p = parts(new Date(ms), tz); return Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour, +p.minute, +p.second) - ms; }
  window.Utilities = {
    formatDate(d, tz, fmt) {
      const p = parts(d, tz);
      return fmt.replace(/'([^']*)'|yyyy|yy|MM|dd|HH|mm|ss/g, (m, lit) => lit !== undefined ? lit : ({ yyyy: p.year, yy: p.year.slice(2), MM: p.month, dd: p.day, HH: p.hour, mm: p.minute, ss: p.second })[m]);
    },
    parseDate(s, tz) {
      const m = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}))?$/.exec(s);
      const guess = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0));
      let t = guess - offset(guess, tz); t = guess - offset(t, tz);
      return new Date(t);
    },
    getUuid() { return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => { const r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 3 | 8)).toString(16); }); }
  };

  // ---------- Sheets ----------
  function Sheet(name) { this.name = name; this.rows = []; }
  Sheet.prototype = {
    appendRow(r) { this.rows.push(r.map(String)); },
    setFrozenRows() {},
    getLastRow() { return this.rows.length; },
    getDataRange() { const s = this; return { getDisplayValues: () => s.rows.map(r => r.slice()), getValues: () => s.rows.map(r => r.slice()) }; },
    getRange(r, c, nr, nc) {
      const s = this;
      return {
        setNumberFormat() { return this; },
        setValues(vals) { vals.forEach((row, i) => { while (s.rows.length < r + i) s.rows.push([]); const tgt = s.rows[r + i - 1]; row.forEach((v, j) => { tgt[c - 1 + j] = String(v); }); }); return this; },
        clearContent() { s.rows.splice(r - 1, nr); return this; }
      };
    }
  };
  const sheets = { Sheet1: new Sheet('Sheet1') };
  const book = {
    getSheetByName: n => sheets[n] || null,
    insertSheet: n => (sheets[n] = new Sheet(n)),
    deleteSheet: sh => { delete sheets[sh.name]; }
  };
  window.SpreadsheetApp = { getActiveSpreadsheet: () => book };

  // ---------- Calendar ----------
  let evSeq = 1;
  function Ev(cal, title, start, end, allDay, opts) {
    this.id = 'ev' + (evSeq++); this.cal = cal; this.title = title; this.start = start; this.end = end; this.allDay = allDay; this.tags = {}; this.loc = (opts || {}).location || '';
  }
  Ev.prototype = {
    getId() { return this.id; }, getTitle() { return this.title; }, getStartTime() { return this.start; }, getEndTime() { return this.end; },
    isAllDayEvent() { return this.allDay; }, getAllDayStartDate() { return this.start; }, getAllDayEndDate() { return this.end; },
    getMyStatus() { return 'YES'; }, getLocation() { return this.loc; },
    setTag(k, v) { this.tags[k] = v; }, getTag(k) { return this.tags[k] || null; }, deleteTag(k) { delete this.tags[k]; }, setColor() {},
    setTime(s, e) { this.start = s; this.end = e; }, setAllDayDate(d) { this.start = d; this.end = new Date(d.getTime() + 86400000); this.allDay = true; }, setLocation(l) { this.loc = l; },
    deleteEvent() { const i = this.cal.events.indexOf(this); if (i >= 0) this.cal.events.splice(i, 1); }
  };
  function Cal(id, name) { this.id = id; this.name = name; this.events = []; }
  Cal.prototype = {
    getName() { return this.name; },
    getEvents(a, b) { return this.events.filter(e => e.start < b && e.end > a); },
    createEvent(t, s, e, o) { const ev = new Ev(this, t, s, e, false, o); this.events.push(ev); return ev; },
    createAllDayEvent(t, d, o) { const ev = new Ev(this, t, d, new Date(d.getTime() + 86400000), true, o); this.events.push(ev); return ev; },
    getEventById(id) { return this.events.filter(e => e.id === id)[0] || null; }
  };
  const cals = { primary: new Cal('lottewarren95@gmail.com', 'Charlotte Warren'), 'cwa@hereslaw.co.uk': new Cal('cwa@hereslaw.co.uk', 'cwa@hereslaw.co.uk'), 'office@hammer-and-pen.com': new Cal('office@hammer-and-pen.com', 'Hammer & Pen office') };
  window.CalendarApp = {
    GuestStatus: { NO: 'NO' }, EventColor: { GRAY: '8' },
    getDefaultCalendar: () => cals.primary,
    getCalendarById: id => cals[id] || null
  };

  // ---------- the rest ----------
  const store = {};
  window.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {} }) };
  window.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  window.FAMCAL_SENT = [];
  window.MailApp = { sendEmail: o => { window.FAMCAL_SENT.push(o); console.log('[preview email] to ' + o.to + ': ' + o.subject); }, getRemainingDailyQuota: () => 100 };
  window.Logger = { log: m => console.log(m) };
  const chain = { timeBased() { return chain; }, everyDays() { return chain; }, atHour() { return chain; }, inTimezone() { return chain; }, create() { return {}; } };
  window.ScriptApp = { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => chain };
  window.ContentService = { MimeType: { JSON: 'json', ICAL: 'ical' }, createTextOutput: t => ({ text: t, setMimeType() { return this; } }) };

  window.FAMCAL_MOCK = function (fn, args) {
    return new Promise((resolve, reject) => {
      setTimeout(() => {
        try {
          const out = JSON.parse(doPost({ postData: { contents: JSON.stringify({ fn: fn, args: args }) } }).text);
          out.ok ? resolve(out.data) : reject(new Error(out.error));
        } catch (e) { reject(e); }
      }, 150);
    });
  };

  // ---------- example data, set up once Code.gs has loaded ----------
  window.FAMCAL_SEED = function () {
    setUpEverything();
    setSetting_('admin_key', 'demo-charlotte-key-00000000');
    setSetting_('site_url', 'https://charlotte644.github.io/family-calendar/');
    const tokens = { mumdad: 'demo-mumdad-token-0000000', tomsmum: 'demo-tomsmum-token-000000', tom: 'demo-tom-token-0000000000' };
    readAll_('People').forEach(p => { p.token = tokens[p.id]; p.email = p.id + '@example.com'; saveObj_('People', p); });
    const t = today_(), mon = dateAdd_(t, 8 - weekdayOf_(t));   // next Monday
    const at = (d, a, b) => [dt_(d, a), dt_(d, b)];
    const add = (cal, d, a, b, title) => { const x = at(d, a, b); cals[cal].createEvent(title, x[0], x[1], {}); };
    for (let w = 0; w < 4; w++) {
      const m = dateAdd_(mon, w * 7);
      [0, 1, 2, 3].forEach(i => add('cwa@hereslaw.co.uk', dateAdd_(m, i), '09:00', '15:00', 'Heres Law: drafting'));
      add('office@hammer-and-pen.com', dateAdd_(m, 2), '10:00', '11:30', 'Discovery call');
      add('office@hammer-and-pen.com', dateAdd_(m, 4), '15:00', '16:00', 'Weekly catch-up');
      add('primary', dateAdd_(m, 2), '18:00', '20:00', 'Heres Law seminar');   // ignored: contains "seminar"
      add('primary', dateAdd_(m, 5), '10:00', '12:00', 'Swimming with Sam');
    }
    add('primary', dateAdd_(mon, 3), '17:00', '19:00', 'Dinner at JRC Buffet');
    cals.primary.createAllDayEvent('Josh\'s stag', dt_(dateAdd_(mon, 4), '00:00'), {});
    const plan = (by, d, s, e, title, place, people) => {
      const p = { id: nextId_('P'), created: '', created_by: by, kind: 'plan', date: d, start: s, end: e, title: title, place: place, notes: '', people: people, status: 'on', event_id: '', updated: '' };
      p.event_id = makeEvent_(p); appendObj_('Plans', p);
    };
    plan('mumdad', dateAdd_(mon, 6), '13:00', '15:30', 'Sunday lunch', 'Nanay & Tatay\'s house', 'mumdad');
    plan('charlotte', dateAdd_(mon, 12), '', '', 'Day out at the zoo', 'Bristol Zoo Project', 'tomsmum');
    haveCare('demo-tomsmum-token-000000', { kind: 'nixon', date: dateAdd_(mon, 3), start: '09:00', end: '17:00', notes: 'I\'ll walk him at lunchtime' });
    if (window.FAMCAL_IDEAS) adminImportIdeas('demo-charlotte-key-00000000', window.FAMCAL_IDEAS);
    adminSaveIdea('demo-charlotte-key-00000000', { title: 'Forest walk and hot chocolate', place: 'Wentwood Forest', price: 'Free (parking £3)', dog: 'yes', notes: 'Bring the sledge if it snows' });
    adminSaveCard('demo-charlotte-key-00000000', 'sam', [{ section: 'Daily routine & naps', text: 'Up about 6:30. Nap 12:30–2pm with his bunny.' }, { section: 'Food & drink', text: 'Loves pasta and bananas. Milk before bed.' }]);
    adminSaveBirthdays('demo-charlotte-key-00000000', [{ name: 'Tatay', date: '1960-' + dateAdd_(mon, 2).slice(5), kind: 'birthday', gift: 'Golf balls' }]);
    window.FAMCAL_SENT.length = 0;
  };
  window.FAMCAL_DEMO_PEOPLE = [['Charlotte', 'demo-charlotte-key-00000000'], ['Nanay & Tatay', 'demo-mumdad-token-0000000'], ['Nana', 'demo-tomsmum-token-000000'], ['Tom', 'demo-tom-token-0000000000']];

  if (!window.FAMCAL_BUNDLED) document.write('<script src="../apps-script/Code.gs"><\/script><script>FAMCAL_SEED();<\/script>');
  document.addEventListener('DOMContentLoaded', () => {
    const bar = document.getElementById('demo');
    bar.className = 'demo';
    bar.innerHTML = '<span>Preview with made-up data. View as:</span>' + window.FAMCAL_DEMO_PEOPLE.map(p => '<button type="button" data-demo="' + p[1] + '">' + p[0] + '</button>').join('');
    bar.addEventListener('click', e => { const b = e.target.closest('[data-demo]'); if (b) { try { localStorage.setItem('famcal_k', b.dataset.demo); } catch (x) { } window.FAMCAL_RELOAD(b.dataset.demo); } });
  });
})();
