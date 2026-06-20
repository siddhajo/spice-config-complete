/**
 * mask-fields.js — Sensitive-field masking for client-facing documents.
 *
 * Mirrors the front-end maskField() in public/index.html EXACTLY so the
 * masked value shown on screen matches the value printed on receipts,
 * invoices and reports. Policy lives in three company_settings keys
 * (display category): mask_acct / mask_ifsc / mask_phone.
 *
 * Modes: none | last4 | last6 | first4 | first6 | first2last2 | full
 *
 * IMPORTANT: functional outputs — the bank payment (NEFT/RTGS) file, DBF,
 * Tally XML and WhatsApp send targets — must NEVER be masked (they need the
 * real values). Only apply these helpers to human-facing documents.
 */

// Mask a single value per the given mode. Returns the value unchanged for
// empty input or mode 'none'/unknown.
function maskField(value, mode) {
  const s = value == null ? '' : String(value);
  if (!s || !mode || mode === 'none') return s;
  if (mode === 'full') return '*'.repeat(s.length);
  if (mode === 'first2last2') {
    if (s.length <= 4) return s; // nothing left to hide between first/last 2
    return s.slice(0, 2) + '*'.repeat(s.length - 4) + s.slice(-2);
  }
  const show = (mode === 'last4' || mode === 'first4') ? 4
             : (mode === 'last6' || mode === 'first6') ? 6 : 0;
  if (!show || s.length <= show) return s;
  const stars = '*'.repeat(s.length - show);
  return (mode === 'last4' || mode === 'last6')
    ? stars + s.slice(-show)
    : s.slice(0, show) + stars;
}

// Build per-field maskers from a flat company-settings object (cfg). Returns
// { maskAcct, maskIfsc, maskPhone } closures so callers don't repeat the key
// lookups. Defaults match the company-config seed (acct → last4, others none).
function makeMaskers(cfg) {
  cfg = cfg || {};
  const acct  = cfg.mask_acct  || 'none';
  const ifsc  = cfg.mask_ifsc  || 'none';
  const phone = cfg.mask_phone || 'none';
  return {
    maskAcct:  (v) => maskField(v, acct),
    maskIfsc:  (v) => maskField(v, ifsc),
    maskPhone: (v) => maskField(v, phone),
  };
}

module.exports = { maskField, makeMaskers };
