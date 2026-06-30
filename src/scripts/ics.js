// lilICS: turn event details into add-to-calendar links (Google, Outlook,
// Microsoft 365, Yahoo), a downloadable .ics, and a paste-ready HTML snippet.
// Fully client-side; times are read in the user's local timezone.

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];

/* ---------- theme (OS-aware, matches the family) ---------- */
const MOON_SVG = '<svg viewBox="0 0 24 24" width="17" height="17" aria-hidden="true"><path fill="currentColor" d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8Z"/></svg>';
const SUN_SVG = '<svg viewBox="0 0 24 24" width="17" height="17" aria-hidden="true"><g fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2.5v2M12 19.5v2M4.6 4.6l1.4 1.4M18 18l1.4 1.4M2.5 12h2M19.5 12h2M4.6 19.4l1.4-1.4M18 6l1.4-1.4"/></g></svg>';

function setThemeIcon(btn, theme) {
  if (theme === 'dark') { btn.innerHTML = SUN_SVG; btn.setAttribute('aria-label', 'Switch to light mode'); }
  else { btn.innerHTML = MOON_SVG; btn.setAttribute('aria-label', 'Switch to dark mode'); }
}
function initTheme() {
  const btn = $('#ui-theme-btn');
  const current = () => (document.documentElement.dataset.theme === 'dark' ? 'dark' : 'light');
  setThemeIcon(btn, current());
  btn.addEventListener('click', () => {
    const next = current() === 'dark' ? 'light' : 'dark';
    document.documentElement.dataset.theme = next;
    try { localStorage.setItem('lilics-theme', next); } catch (e) { /* storage may be unavailable; safe to ignore */ }
    setThemeIcon(btn, next);
  });
}

/* ---------- date plumbing ---------- */
const pad = (n) => String(n).padStart(2, '0');

// 20260618T170000Z from a local Date
function utcStamp(d) {
  return d.getUTCFullYear() + pad(d.getUTCMonth() + 1) + pad(d.getUTCDate()) +
    'T' + pad(d.getUTCHours()) + pad(d.getUTCMinutes()) + '00Z';
}
// 20260618 (all-day, date only)
function dayStamp(d) {
  return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate());
}

function readEvent() {
  const title = $('#f-title').value.trim();
  const location = $('#f-location').value.trim();
  const desc = $('#f-desc').value.trim();
  const allDay = $('#f-allday').checked;
  const startRaw = $('#f-start').value;
  const endRaw = $('#f-end').value;
  if (!startRaw) return null;
  const start = new Date(startRaw);
  let end = endRaw ? new Date(endRaw) : new Date(start.getTime() + 3600000);
  if (isNaN(start) || isNaN(end)) return null;
  if (end <= start) end = new Date(start.getTime() + 3600000);
  return { title: title || 'Untitled event', location, desc, allDay, start, end };
}

function dates(ev) {
  if (ev.allDay) {
    // all-day ranges use date-only stamps with an exclusive end
    const endNext = new Date(ev.end.getFullYear(), ev.end.getMonth(), ev.end.getDate() + 1);
    return { s: dayStamp(ev.start), e: dayStamp(endNext), iso: false };
  }
  return { s: utcStamp(ev.start), e: utcStamp(ev.end), iso: true };
}

/* ---------- link builders ---------- */
function googleUrl(ev) {
  const d = dates(ev);
  const p = new URLSearchParams({ action: 'TEMPLATE', text: ev.title, dates: `${d.s}/${d.e}` });
  if (ev.desc) p.set('details', ev.desc);
  if (ev.location) p.set('location', ev.location);
  return 'https://calendar.google.com/calendar/render?' + p.toString();
}
function outlookUrl(ev, host) {
  const p = new URLSearchParams({ path: '/calendar/action/compose', rru: 'addevent', subject: ev.title });
  if (ev.allDay) {
    p.set('allday', 'true');
    p.set('startdt', ev.start.getFullYear() + '-' + pad(ev.start.getMonth() + 1) + '-' + pad(ev.start.getDate()));
    p.set('enddt', ev.end.getFullYear() + '-' + pad(ev.end.getMonth() + 1) + '-' + pad(ev.end.getDate()));
  } else {
    p.set('startdt', ev.start.toISOString());
    p.set('enddt', ev.end.toISOString());
  }
  if (ev.desc) p.set('body', ev.desc);
  if (ev.location) p.set('location', ev.location);
  return `https://${host}/calendar/0/action/compose?` + p.toString();
}
function yahooUrl(ev) {
  const d = dates(ev);
  const p = new URLSearchParams({ v: '60', title: ev.title, st: d.s, et: d.e });
  if (ev.allDay) p.set('dur', 'allday');
  if (ev.desc) p.set('desc', ev.desc);
  if (ev.location) p.set('in_loc', ev.location);
  return 'https://calendar.yahoo.com/?' + p.toString();
}

/* ---------- ics ---------- */
const icsEsc = (s) => s.replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');

function buildIcs(ev) {
  const d = dates(ev);
  const uid = 'lilics-' + Math.abs([...(ev.title + d.s)].reduce((a, c) => (a * 31 + c.charCodeAt(0)) | 0, 7)) + '@lilagents.com';
  const lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//lilAgents//lilICS//EN',
    'BEGIN:VEVENT',
    `UID:${uid}`,
    `DTSTAMP:${utcStamp(new Date())}`,
    ev.allDay ? `DTSTART;VALUE=DATE:${d.s}` : `DTSTART:${d.s}`,
    ev.allDay ? `DTEND;VALUE=DATE:${d.e}` : `DTEND:${d.e}`,
    `SUMMARY:${icsEsc(ev.title)}`,
  ];
  if (ev.location) lines.push(`LOCATION:${icsEsc(ev.location)}`);
  if (ev.desc) lines.push(`DESCRIPTION:${icsEsc(ev.desc)}`);
  lines.push('END:VEVENT', 'END:VCALENDAR');
  return lines.join('\r\n') + '\r\n';
}

/* ---------- share links (the event travels inside the URL) ---------- */
const b64urlEnc = (obj) => btoa(unescape(encodeURIComponent(JSON.stringify(obj)))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const b64urlDec = (s) => JSON.parse(decodeURIComponent(escape(atob(s.replace(/-/g, '+').replace(/_/g, '/')))));

const localDateStr = (d) => d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
const localInputStr = (d) => localDateStr(d) + 'T' + pad(d.getHours()) + ':' + pad(d.getMinutes());

function sharePayload(ev) {
  // Timed events travel as absolute epoch times so they land correctly in
  // any timezone; all-day events travel as plain dates so they never shift.
  return ev.allDay
    ? { t: ev.title, l: ev.location, d: ev.desc, a: 1, s: localDateStr(ev.start), e: localDateStr(ev.end) }
    : { t: ev.title, l: ev.location, d: ev.desc, a: 0, s: ev.start.getTime(), e: ev.end.getTime() };
}
function shareUrl(ev) {
  return location.origin + location.pathname + '#e=' + b64urlEnc(sharePayload(ev));
}

function applyShared(p) {
  $('#f-title').value = p.t || '';
  $('#f-location').value = p.l || '';
  $('#f-desc').value = p.d || '';
  $('#f-allday').checked = !!p.a;
  if (p.a) {
    $('#f-start').value = p.s + 'T09:00';
    $('#f-end').value = p.e + 'T17:00';
  } else {
    $('#f-start').value = localInputStr(new Date(p.s));
    $('#f-end').value = localInputStr(new Date(p.e));
  }
}

/* ---------- render ---------- */
function snippet(ev) {
  return `<a href="${googleUrl(ev)}" target="_blank">Add to Google Calendar</a>
<a href="${outlookUrl(ev, 'outlook.live.com')}" target="_blank">Add to Outlook</a>
<a href="${yahooUrl(ev)}" target="_blank">Add to Yahoo Calendar</a>
<a href="${shareUrl(ev)}&dl=1">Add to Apple Calendar (.ics)</a>`;
}

function render() {
  const ev = readEvent();
  const notes = $('#notes');
  if (!ev) {
    $('#code').textContent = 'Set a start time on the left (or check All-day event for a date-only event) and the links build themselves.';
    $('#share-url').value = '';
    $('#share-dl').value = '';
    notes.innerHTML = '';
    ['google', 'outlook', 'office', 'yahoo'].forEach((k) => { $('#link-' + k).removeAttribute('href'); });
    return;
  }
  $('#link-google').href = googleUrl(ev);
  $('#link-outlook').href = outlookUrl(ev, 'outlook.live.com');
  $('#link-office').href = outlookUrl(ev, 'outlook.office.com');
  $('#link-yahoo').href = yahooUrl(ev);
  const share = shareUrl(ev);
  $('#share-url').value = share;
  $('#share-dl').value = share + '&dl=1';
  $('#code').textContent = snippet(ev);
  notes.innerHTML = `<div class="note note--ok">${ev.allDay ? 'All-day event' : 'Timed event, converted from your local timezone'}: ${ev.title}. Links, share URL, and .ics are live.</div>`;
}

function flash(btn, label) {
  const prev = btn.textContent;
  btn.textContent = label;
  btn.classList.add('btn--done');
  setTimeout(() => { btn.textContent = prev; btn.classList.remove('btn--done'); }, 1100);
}
function fallbackCopy(text, done) {
  const ta = document.createElement('textarea');
  ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
  document.body.appendChild(ta); ta.select();
  try { document.execCommand('copy'); } catch (e) { /* storage may be unavailable; safe to ignore */ }
  document.body.removeChild(ta); done();
}

function downloadIcs() {
  const ev = readEvent();
  if (!ev) return false;
  const blob = new Blob([buildIcs(ev)], { type: 'text/calendar' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = (ev.title.replace(/[^a-z0-9 _-]/gi, '') || 'event') + '.ics';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
  return true;
}

function copyHandler(btnSel, getText) {
  $(btnSel).addEventListener('click', (e) => {
    const btn = e.currentTarget;
    const text = getText();
    if (!text) return;
    const done = () => flash(btn, 'Copied');
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(done).catch(() => fallbackCopy(text, done));
    else fallbackCopy(text, done);
  });
}

// All-day events only need a date. A datetime-local input with no time entered
// is an empty value, which is why picking just a date looked broken, so swap
// the start/end inputs to plain date pickers while All-day is on.
function syncAllDay() {
  const allDay = $('#f-allday').checked;
  ['#f-start', '#f-end'].forEach((sel) => {
    const el = $(sel);
    const v = el.value;
    el.type = allDay ? 'date' : 'datetime-local';
    if (v) el.value = allDay ? v.slice(0, 10) : (v.length === 10 ? v + 'T09:00' : v);
  });
}

function initIcs() {
  initTheme();
  syncAllDay();
  ['#f-title', '#f-location', '#f-desc', '#f-start', '#f-end'].forEach((sel) =>
    $(sel).addEventListener('input', render));
  $('#f-allday').addEventListener('change', () => { syncAllDay(); render(); });

  $('#dl-ics').addEventListener('click', (e) => {
    if (downloadIcs()) flash(e.currentTarget, 'Saved');
  });

  copyHandler('#copy-btn', () => $('#code').textContent);
  copyHandler('#copy-share', () => $('#share-url').value);
  copyHandler('#copy-dl', () => $('#share-dl').value);

  // Opening a shared event link prefills everything and arms the buttons;
  // the &dl=1 variant downloads the .ics immediately.
  const hashEvent = (location.hash.match(/[#&]e=([^&]+)/) || [])[1];
  if (hashEvent) {
    try {
      applyShared(b64urlDec(hashEvent));
      syncAllDay();
      render();
      $('#notes').innerHTML = `<div class="note note--ok">You opened a shared event. The buttons above are ready to click.</div>` + $('#notes').innerHTML;
      if (/[#&]dl=1/.test(location.hash)) downloadIcs();
      return;
    } catch (e) { /* malformed hash; fall through to a blank form */ }
  }

  render();
}

export { initIcs };
