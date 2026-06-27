/**
 * date-format.js — shared display-date formatter.
 *
 * The user picks one of DD/MM/YYYY, DD-MM-YYYY, YYYY-MM-DD in
 * Settings → Display → Date format. This module reads the choice from
 * company_settings (lazily cached) and exposes a single fmtDate(d, fmt?)
 * helper that any PDF / XLSX / UI-feeding code path can call.
 *
 * NOT used for:
 *   • Tally XML (Tally requires its own YYYYMMDD — see tally-xml.js)
 *   • DBF exports (FoxPro requires its own DD/MM/YYYY — see dbf-exports.js)
 *   • DB storage (always ISO YYYY-MM-DD — see normalizeDate in server.js)
 */

let _cache = null;

function getDateFormat() {
  if (_cache != null) return _cache;
  try {
    const { getSetting } = require('./company-config');
    const { getDb }      = require('./db');
    _cache = String(getSetting(getDb(), 'date_format') || 'dd/mm/yyyy');
  } catch (_) {
    _cache = 'dd/mm/yyyy';
  }
  return _cache;
}

function invalidateDateFormatCache() { _cache = null; }

// Display: ISO yyyy-mm-dd (or any date-ish input) → user-chosen format.
// Delegates to the single canonical formatter (report-formatters) so the
// XLSX exports follow the exact same date format as the rest of the app,
// including the month-name variants. Pass `fmt` to override the setting.
function fmtDate(d, fmt) {
  if (!d && d !== 0) return '';
  try {
    return require('./report-formatters').formatDateForDisplay(d, fmt || getDateFormat());
  } catch (_) {
    return String(d);
  }
}

// Today's date in local time as YYYY-MM-DD — use this instead of
// `new Date()` when you need "today", because Date.toISOString() is
// UTC and rolls back a day for users east of UTC during early-morning
// hours (e.g. IST at 5 AM is still the previous day in UTC).
function todayLocalISO() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

// ── IST (Asia/Kolkata) time helpers ──────────────────────────────
// All human-facing TIME display in the app is shown in IST regardless of
// the server's own system timezone, so a report generated on a UTC host
// still reads the wall-clock the (Indian) users expect.

// Break a Date down into IST wall-clock parts. Uses Intl so the offset
// (incl. the :30) is handled correctly without hard-coding +5:30.
function istParts(d) {
  d = d || new Date();
  const f = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kolkata',
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
    hour12: false, weekday: 'short',
  });
  const p = {};
  for (const part of f.formatToParts(d)) p[part.type] = part.value;
  const WD = { Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6 };
  return {
    y: p.year, m: p.month, d: p.day,
    hh: p.hour === '24' ? '00' : p.hour, mm: p.minute, ss: p.second,
    weekday: WD[p.weekday],
  };
}

// "HH:MM" 24-hour wall-clock in IST — for matching scheduled times.
function istHHMM(d) {
  const p = istParts(d);
  return `${p.hh}:${p.mm}`;
}

// Full date+time in IST, date part following the user's chosen format,
// suffixed " IST". For report footers / "Generated:" stamps.
function nowIST(d) {
  const p = istParts(d);
  return `${fmtDate(`${p.y}-${p.m}-${p.d}`)} ${p.hh}:${p.mm} IST`;
}

module.exports = { fmtDate, getDateFormat, invalidateDateFormatCache, todayLocalISO, istParts, istHHMM, nowIST };
