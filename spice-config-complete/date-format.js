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

module.exports = { fmtDate, getDateFormat, invalidateDateFormatCache, todayLocalISO };
