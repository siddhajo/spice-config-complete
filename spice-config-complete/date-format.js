/**
 * date-format.js â shared display-date formatter.
 *
 * The user picks one of DD/MM/YYYY, DD-MM-YYYY, YYYY-MM-DD in
 * Settings â Display â Date format. This module reads the choice from
 * company_settings (lazily cached) and exposes a single fmtDate(d, fmt?)
 * helper that any PDF / XLSX / UI-feeding code path can call.
 *
 * NOT used for:
 *   â¢ Tally XML (Tally requires its own YYYYMMDD â see tally-xml.js)
 *   â¢ DBF exports (FoxPro requires its own DD/MM/YYYY â see dbf-exports.js)
 *   â¢ DB storage (always ISO YYYY-MM-DD â see normalizeDate in server.js)
 */

let _cache = null;

function getDateFormat() {
  if (_cache != null) return _cache;
  try {
    const { getSetting } = require('./company-config');
    const { getDb }      = require('./db');
    _cache = String(getSetting(getDb(), 'date_format') || 'DD/MM/YYYY').toUpperCase();
  } catch (_) {
    _cache = 'DD/MM/YYYY';
  }
  return _cache;
}

function invalidateDateFormatCache() { _cache = null; }

// Display: ISO yyyy-mm-dd (or any date-ish input) â user-chosen format.
// Pass `fmt` explicitly when you already have the format string handy
// (e.g. inside a loop building a PDF), otherwise the cached default is
// used. Falls through to the original string when the input can't be
// parsed.
function fmtDate(d, fmt) {
  if (!d && d !== 0) return '';
  let iso = '';
  if (d instanceof Date && !isNaN(d)) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    iso = `${y}-${m}-${day}`;
  } else {
    const s = String(d).trim();
    // ISO yyyy-mm-dd (possibly with time)
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) iso = `${m[1]}-${m[2]}-${m[3]}`;
    else {
      // dd/mm/yyyy or dd-mm-yyyy
      m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
      if (m) iso = `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
    }
  }
  if (!iso) return String(d);
  const [y, mo, day] = iso.split('-');
  const f = String(fmt || getDateFormat() || 'DD/MM/YYYY').toUpperCase();
  if (f === 'YYYY-MM-DD') return `${y}-${mo}-${day}`;
  if (f === 'DD-MM-YYYY') return `${day}-${mo}-${y}`;
  return `${day}/${mo}/${y}`;   // DD/MM/YYYY (default)
}

// Today's date in local time as YYYY-MM-DD â use this instead of
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

module.exports = { fmtDate, getDateFormat, invalidateDateFormatCache, todayLocalISO };
