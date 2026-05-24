/**
 * _company-identity-fallback.js
 *
 * Tiny shared helper that returns a working `getCompanyIdentity`
 * function regardless of whether the main `report-formatters.js`
 * exports one. Used as a safety net by every consumer (invoice-pdf,
 * server, exports, auction-reports, tally-xml) so a stale deploy of
 * report-formatters.js can't crash PDF/Excel generation with
 * "getCompanyIdentity is not a function".
 *
 * The fallback resolver mirrors the priority chain of the real one:
 *   name      : company_name → tally_company_name → short_name
 *   shortName : short_name → logo (uppercased) → first word of name
 *   address1  : tn_address1 → address1 → address
 *   address2  : tn_address2 → address2 → tn_branch → branch
 *   gstin     : gstin → tn_gstin → business_gstin
 *   pan       : pan → tn_pan → business_pan → derived (gstin[2..12])
 *   state     : tn_state → business_state → state (uppercased)
 *   stateCode : tally_state_code → derived (gstin[0..2]) → ''
 */

function inlineResolver(cfg) {
  cfg = cfg || {};
  const pick = (...keys) => {
    for (const k of keys) {
      const v = cfg[k];
      if (v != null && String(v).trim() !== '') return String(v).trim();
    }
    return '';
  };
  const name      = pick('trade_name', 'company_name', 'tally_company_name', 'short_name');
  const logoCode  = pick('logo');
  const gstin     = pick('gstin', 'tn_gstin', 'business_gstin');
  return {
    name,
    shortName: pick('short_name') || logoCode.toUpperCase() || (name.split(/\s+/)[0] || ''),
    logoCode,
    address1: pick('tn_address1', 'address1', 'address'),
    address2: pick('tn_address2', 'address2', 'tn_branch', 'branch'),
    gstin,
    pan: pick('pan', 'tn_pan', 'business_pan')
      || (gstin && gstin.length >= 12 ? gstin.slice(2, 12) : ''),
    state: pick('tn_state', 'business_state', 'state').toUpperCase(),
    stateCode: pick('tally_state_code') || (gstin && gstin.length >= 2 ? gstin.slice(0, 2) : ''),
    cin: pick('cin'),
    idLine: (() => {
      const isPart = String(cfg.is_partnership || '').toLowerCase() === 'true';
      return isPart
        ? { label: 'Partnership', value: pick('partnership_name'), isPartnership: true }
        : { label: 'CIN',         value: pick('cin'),               isPartnership: false };
    })(),
  };
}

// Returns the real resolver if exported, else the fallback.
function resolve() {
  try {
    const _rf = require('./report-formatters');
    if (typeof _rf.getCompanyIdentity === 'function') return _rf.getCompanyIdentity;
  } catch (_) { /* fall through */ }
  return inlineResolver;
}

module.exports = { resolve, inlineResolver };
