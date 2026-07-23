/**
 * tally-xml.js — Tally-importable XML generators
 *
 * Ports the VBA macros (ConvertSales, ConvertPurchase, ConvertDebit) from
 * the IDEAL_V5_6 / ASPPL_V5_6 .xlsm files into pure-JS functions that build
 * the same <ENVELOPE>...</ENVELOPE> Tally XML payloads.
 *
 * Four export types:
 *   generSalesXML       — registered dealer sales (generXML)
 *   generRDPurchaseXML  — registered dealer purchases (generRD)
 *   generURDPurchaseXML — agriculturist / unregistered purchases (generURD)
 *   generDebitNoteXML   — discount debit notes against suppliers (generDN)
 *
 * Each function receives {rows, cfg, opts} where:
 *   rows = pre-grouped invoice/voucher records pulled from the SQLite DB
 *   cfg  = company_settings flat object (getSettingsFlat output)
 *   opts = { season, separator, voucherStart } overrides per call
 *
 * The XML is text-only; we return a string ready for download. We don't
 * touch ExcelJS or PDF here — that's a separate path.
 */

// ── Indian state code → name (matches FindState in VBA) ──────────
const STATES = {
  '01': 'Jammu & Kashmir', '02': 'Himachal Pradesh', '03': 'Punjab',
  '04': 'Chandigarh', '05': 'Uttarakhand', '06': 'Haryana',
  '07': 'Delhi', '08': 'Rajasthan', '09': 'Uttar Pradesh',
  '10': 'Bihar', '11': 'Sikkim', '12': 'Arunachal Pradesh',
  '13': 'Nagaland', '14': 'Manipur', '15': 'Mizoram',
  '16': 'Tripura', '17': 'Meghalaya', '18': 'Assam',
  '19': 'West Bengal', '20': 'Jharkhand', '21': 'Odisha',
  '22': 'Chhattisgarh', '23': 'Madhya Pradesh', '24': 'Gujarat',
  '25': 'Daman & Diu', '26': 'Dadra & Nagar Haveli', '27': 'Maharashtra',
  '28': 'Andhra Pradesh', '29': 'Karnataka', '30': 'Goa',
  '31': 'Lakshadweep', '32': 'Kerala', '33': 'Tamil Nadu',
  '34': 'Puducherry', '35': 'Andaman & Nicobar Islands',
  '36': 'Telangana', '37': 'Andhra Pradesh (New)',
  '97': 'Other Territory', '99': 'Other Country',
};

const findState = (gstin) => {
  if (!gstin) return '';
  const code = String(gstin).trim().slice(0, 2);
  return STATES[code] || '';
};

// ── XML escaping ─────────────────────────────────────────────────
const xe = (v) => {
  if (v === null || v === undefined) return '';
  return String(v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
};

const r2 = (n) => Math.round(Number(n || 0) * 100) / 100;
const r0 = (n) => Math.round(Number(n || 0));

// yyyymmdd from any date-ish string ("2026-04-28", "28/04/2026", or Date)
const toTallyDate = (d) => {
  if (!d) return '';
  if (d instanceof Date) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}${m}${day}`;
  }
  const s = String(d).trim();
  // yyyy-mm-dd or yyyy-mm-ddT...
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}${m[2]}${m[3]}`;
  // dd/mm/yyyy
  m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return `${m[3]}${m[2]}${m[1]}`;
  // yyyymmdd (already correct)
  if (/^\d{8}$/.test(s)) return s;
  return s.replace(/\D/g, '').slice(0, 8);
};

// ── Tally XML constants (mirror VBA constants in ConvertSales.bas) ──
const TAGS = {
  STARTENV:  '<ENVELOPE>',
  HEADER:    '<HEADER>\n<TALLYREQUEST>Import Data</TALLYREQUEST>\n</HEADER>',
  SIMPDATA:  '<IMPORTDATA>',
  EREQDESC:  '</REQUESTDESC>',
  SREQDATA:  '<REQUESTDATA>',
  STARTDATA: '<TALLYMESSAGE xmlns:UDF="TallyUDF">',
  ENDDATA:   '</TALLYMESSAGE>',
  EREQDATA:  '</REQUESTDATA>',
  EIMPDATA:  '</IMPORTDATA>',
  ENDBODY:   '</BODY>',
  ENDVOUCHER:'</VOUCHER>',
  DEEMNO:    '<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>\n<ISLASTDEEMEDPOSITIVE>No</ISLASTDEEMEDPOSITIVE>',
  DEEMYES:   '<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>\n<ISLASTDEEMEDPOSITIVE>Yes</ISLASTDEEMEDPOSITIVE>',
};

// ── GST rate detail blocks ─────────────────────────────────────
const rateBlock = (head, rate, valuation = 'Based on Value') => {
  let body = `<GSTRATEDUTYHEAD>${head}</GSTRATEDUTYHEAD>\n<GSTRATEVALUATIONTYPE>${valuation}</GSTRATEVALUATIONTYPE>`;
  if (rate !== null && rate !== undefined) body += `\n<GSTRATE>${rate}</GSTRATE>`;
  return `<RATEDETAILS.LIST>\n${body}\n</RATEDETAILS.LIST>`;
};

const rateDetails = (gstRate /* e.g. 5 = full IGST */) => {
  const half = (gstRate / 2).toFixed(2);
  const full = String(gstRate);
  return {
    cgst: rateBlock('CGST', half),
    sgst: rateBlock('SGST/UTGST', half),
    igst: rateBlock('IGST', full),
    // Cess + State Cess: emit head + valuation but no rate (matches target schemas)
    cess:  rateBlock('Cess',       null),
    scess: rateBlock('State Cess', null),
  };
};

// ── Envelope helpers ──────────────────────────────────────────
const startEnvelope = (companyName, reportName = 'Vouchers') => {
  const sreqdesc = `<REQUESTDESC>\n<REPORTNAME>${reportName}</REPORTNAME>`;
  const stat = `<STATICVARIABLES>\n<SVCURRENTCOMPANY>${xe(companyName)}</SVCURRENTCOMPANY>\n</STATICVARIABLES>`;
  const startBody = `<BODY>\n${TAGS.SIMPDATA}\n${sreqdesc}\n${stat}\n${TAGS.EREQDESC}\n${TAGS.SREQDATA}\n${TAGS.STARTDATA}`;
  return `${TAGS.STARTENV}\n${TAGS.HEADER}\n${startBody}`;
};

const endEnvelope = () => {
  return `${TAGS.ENDDATA}\n${TAGS.EREQDATA}\n${TAGS.EIMPDATA}\n${TAGS.ENDBODY}\n</ENVELOPE>`;
};

// ── Cfg accessors with sensible fallbacks ──────────────────────
const cfgGet = (cfg, key, def = '') => {
  if (!cfg) return def;
  const v = cfg[key];
  if (v === undefined || v === null || v === '') return def;
  return v;
};

const cfgBool = (cfg, key, def = false) => {
  const v = cfgGet(cfg, key, null);
  if (v === null) return def;
  if (typeof v === 'boolean') return v;
  const s = String(v).toLowerCase();
  return s === 'true' || s === '1' || s === 'yes' || s === 'on';
};

const cfgNum = (cfg, key, def = 0) => {
  const v = Number(cfgGet(cfg, key, def));
  return isFinite(v) ? v : def;
};

// SQL fragments for filtering lot rows by whether the seller has a GSTIN.
// The `cr` column may carry either format:
//   • Legacy UI:  "GSTIN.<15-char-gstin>"  → starts with "GSTIN" (uppercase)
//   • Bulk import: bare 15-char GSTIN     → starts with 2 digits
// Both need to be classified the same way; use these fragments everywhere
// instead of bare LIKE clauses to keep the rules consistent.
//
//   HAS_GSTIN_SQL — match registered-dealer rows (RD vouchers, RD ledgers)
//   NO_GSTIN_SQL  — match agriculturist rows  (URD vouchers, URD ledgers)
const HAS_GSTIN_SQL = `(UPPER(COALESCE(cr,'')) LIKE 'GSTIN%' OR cr GLOB '[0-9][0-9]*')`;
const NO_GSTIN_SQL  = `(cr IS NULL OR cr = '' OR (UPPER(cr) NOT LIKE 'GSTIN%' AND cr NOT GLOB '[0-9][0-9]*'))`;

// =====================================================================
// Helpers shared by ISP/ASP sales generators
// =====================================================================

// Two-line "BASICBUYERADDRESS" / "ADDRESS.LIST" — exactly what the
// reference XMLs do: line 1 = the joined address line, line 2 = place.
const _addrLines = (address, place) => {
  const a = String(address || '').trim();
  const p = String(place || '').trim();
  return [a, p].filter(Boolean);
};

// Pull the PAN out of a 15-char GSTIN (chars 3..12 inclusive, 1-indexed
// chars 3..12 = 0-indexed 2..11). Reference uses this as
// <CONSIGNEEPINNUMBER>HZSPM0006E</CONSIGNEEPINNUMBER>.
const _panFromGstin = (gstin) => {
  const s = String(gstin || '').replace(/\s+/g, '').toUpperCase();
  if (s.length < 15) return '';
  return s.slice(2, 12);
};

// Sale type letter ("L"/"I"/"E") → human label used in dispatch terms /
// fallbacks. Currently unused but handy.
const _saleLabel = (s) => {
  const u = String(s || '').toUpperCase();
  if (u === 'L') return 'Local';
  if (u === 'I') return 'Inter-state';
  if (u === 'E') return 'Export';
  return u;
};

// =====================================================================
// 1A. ISP SALES VOUCHERS (sales to outside customer)
// =====================================================================
//
// Matches reference SALE_TRANS-03.xml — full format with consignee,
// e-way bill, dispatch-from (sister company), BASICORDERREF (= the
// matching ASP voucher), order terms list.
//
// Input shape (rows from buildSalesIspRows):
//   rows = [{
//     ano, date, sale ('L'|'I'|'E'), invo, aspInvo (optional, used in
//     BASICORDERREF if present), partyName, address, place, pin,
//     partyGstin, lots: [{lot, bag, qty, rate, amount}, ...],
//     amounttot, gunnyAmt, gunnyBags (optional, count of gunny bags
//     across all lots), cgst, sgst, igst, tcsamt, total, totalRounded,
//     vehicleNo (optional), shippedBy (optional), distance (optional),
//     finalDestination (optional, defaults to place),
//   }, ...]
//
function generSalesIspXML(rows, cfg, opts = {}) {
  const company       = opts.companyName || cfgGet(cfg, 'tally_company_name', cfgGet(cfg, 'short_name', 'Ideal Spices Private Limited'));
  const season        = opts.season || cfgGet(cfg, 'tally_season', cfgGet(cfg, 'season_code', '2026-27'));
  const separator     = opts.separator || cfgGet(cfg, 'tally_separator', '/');
  const invPrefix     = cfgGet(cfg, 'tally_inv_prefix', 'ISP/');
  const ainvPrefix    = cfgGet(cfg, 'tally_ainv_prefix', 'ASP/');
  const detailed      = cfgBool(cfg, 'tally_detailed', true);
  const dispatchEnabled = cfgBool(cfg, 'tally_dispatch_from', true);
  const tcs           = cfgBool(cfg, 'tally_tcs_enabled', false);
  const shipToOverride = cfgBool(cfg, 'tally_ship_to', false);
  // Emit the Round Off ledger on every voucher (default ON), matching the
  // reference: every sale carries the line even when the delta is 0, so
  // whole-rupee totals aren't left without it.
  const tlyrnd        = cfgBool(cfg, 'tally_round_enabled', true);
  const intra         = cfgGet(cfg, 'tally_state_code', '33');
  const homeState     = cfgGet(cfg, 'tally_home_state', 'Tamil Nadu');

  // Ledgers
  const SalesInter   = cfgGet(cfg, 'tally_sales_inter',  'Cardamom Inter-State Sales');
  const SalesIntra   = cfgGet(cfg, 'tally_sales_intra',  'Cardamom Local Sales');
  const SalesExport  = cfgGet(cfg, 'tally_sales_export', 'Cardamom Sales - Export');
  const GunnyInter   = cfgGet(cfg, 'tally_gunny_inter',  'Gunny Interstate Sales');
  const GunnyIntra   = cfgGet(cfg, 'tally_gunny_intra',  'Gunny Local Sales');
  const GunnyExport  = cfgGet(cfg, 'tally_gunny_export', 'Gunny Sales - Export');
  const Tax_CGST     = cfgGet(cfg, 'tally_cgst', 'OUTPUT CGST 2.5%');
  const Tax_SGST     = cfgGet(cfg, 'tally_sgst', 'OUTPUT SGST 2.5%');
  const Tax_IGST     = cfgGet(cfg, 'tally_igst', 'OUTPUT IGST 5%');
  const Tax_TCS      = cfgGet(cfg, 'tally_tcs',  'TCS on Sale of Goods');
  const Round_LDR    = cfgGet(cfg, 'tally_round', 'Round On/Off');
  const Item_Card    = cfgGet(cfg, 'tally_item_cardamom', 'Cardamom');
  const Item_Gunny   = cfgGet(cfg, 'tally_item_gunny',    'Gunny');
  const HSN_Card     = cfgGet(cfg, 'tally_hsn_cardamom',  '09083120');
  const HSN_Gunny    = cfgGet(cfg, 'tally_hsn_gunny',     '63051040');
  const GunnyRate    = cfgNum(cfg, 'tally_gunny_rate',     165);
  // Service ledgers (intra-state ISP only — reference uses fully-spelled
  // names with the rate baked in; user can override in Settings)
  const LDR_Transport  = cfgGet(cfg, 'tally_transport', 'Transport Rs.2.50/per Kg');
  const LDR_Insurance  = cfgGet(cfg, 'tally_insurance', 'Insurance Rs.0.75/per Thousand');
  const SAC_Transport  = cfgGet(cfg, 'tally_hsn_transport', '996791');
  const SAC_Insurance  = cfgGet(cfg, 'tally_hsn_insurance', '997136');

  // Dispatch-from (sister company by default — ASP)
  const d_company    = cfgGet(cfg, 'tally_dispatch_company', cfgGet(cfg, 's_company', cfgGet(cfg, 's_short_name', '')));
  const d_add        = cfgGet(cfg, 'tally_dispatch_address', cfgGet(cfg, 's_address1', ''));
  const d_add2       = cfgGet(cfg, 's_address2', '');
  const d_place      = cfgGet(cfg, 'tally_dispatch_place',   cfgGet(cfg, 'kl_branch', 'NEDUMKANDAM'));
  const d_pin        = cfgGet(cfg, 'tally_dispatch_pin',     '685553');
  const d_state      = cfgGet(cfg, 'tally_dispatch_state',   'Kerala');
  const d_state_code = cfgGet(cfg, 'tally_state_code_amazing', '32');
  const d_gstin      = cfgGet(cfg, 'tally_dispatch_gstin',   cfgGet(cfg, 's_gstin', ''));

  // E-way bill consignor type — reference uses "Amazing"
  const consignorType = cfgGet(cfg, 'tally_consignor_type', 'Amazing');

  let xml = '\n' + startEnvelope(company, 'Vouchers');

  for (const row of rows) {
    const dateval     = toTallyDate(row.date);
    const partyName   = xe(row.partyName);
    const buyerAddrLines = _addrLines(row.address, row.place);
    // E-way bill CONSIGNEEADDRESS prefers the buyer's ship-to (consignee)
    // address when configured; otherwise falls back to the bill-to address.
    const consigneeAddr = String(row.consigneeAddress || '').trim() || (buyerAddrLines[0] || '');
    // BASICBUYERADDRESS: when tally_ship_to is enabled, report the buyer's
    // ship-to (consignee) address instead of bill-to. Falls back to bill-to
    // when the buyer has no consignee address configured. The party's own
    // ADDRESS.LIST (ledger/bill-to mailing address) is left unchanged.
    const shipToLines = _addrLines(row.consigneeAddress, row.consigneePlace);
    const basicBuyerAddrLines = (shipToOverride && shipToLines.length)
      ? shipToLines
      : buyerAddrLines;
    const partyGstin  = xe(row.partyGstin);
    const partyState  = xe(findState(row.partyGstin));
    const partyPin    = xe(row.pin || '');
    const partyPlace  = xe(row.place || '');
    const consigneePAN = _panFromGstin(row.partyGstin);
    const isIntra     = String(row.partyGstin || '').slice(0, 2) === String(intra);
    const sale        = String(row.sale || 'L').toUpperCase();
    const isExport    = sale === 'E';
    const invoNo      = String(row.invo || '').trim();
    const taxNm       = `${sale}${separator}${invoNo}`;
    const voucherNum  = `${invPrefix}${taxNm}/${season}`;
    const aspVoucherRef = row.aspInvo
      ? `${ainvPrefix}I${separator}${String(row.aspInvo).trim()}/${season}`
      : '';

    const rates       = rateDetails(cfgNum(cfg, 'tally_gst_rate', 5));

    // Cardamom ledger + nature
    const cardLedger = isExport ? SalesExport : (isIntra ? SalesIntra : SalesInter);
    const cardNature = isExport
      ? 'Exports - Taxable'
      : (isIntra ? 'Local Sales - Taxable' : 'Interstate Sales - Taxable');
    const gunnyLedger = isExport ? GunnyExport : (isIntra ? GunnyIntra : GunnyInter);

    // Final destination defaults to buyer's place
    const finalDest = xe(row.finalDestination || row.place || '');
    const shippedBy = xe(row.shippedBy || '');
    const vehicleNo = xe(row.vehicleNo || '');
    const distance  = xe(row.distance || '');
    const transportMode = isExport ? '4 - Ship' : '1 - Road';
    const vehicleType   = 'R - Regular';

    const startVoucher = `<VOUCHER VCHTYPE="Sales" ACTION="Create" OBJVIEW="Invoice Voucher View">`;

    xml += `\n${startVoucher}
<PARTYNAME>${partyName}</PARTYNAME>
<ADDRESS.LIST TYPE="String">
${buyerAddrLines.map(l => `<ADDRESS>${xe(l)}</ADDRESS>`).join('\n')}
</ADDRESS.LIST>
<PARTYGSTIN>${partyGstin}</PARTYGSTIN>
<PARTYLEDGERNAME>${partyName}</PARTYLEDGERNAME>
<PARTYMAILINGNAME>${partyName}</PARTYMAILINGNAME>
<PARTYPINCODE>${partyPin}</PARTYPINCODE>
<BASICBUYERNAME>${partyName}</BASICBUYERNAME>
<BASICBUYERADDRESS.LIST TYPE="String">
${basicBuyerAddrLines.map(l => `<BASICBUYERADDRESS>${xe(l)}</BASICBUYERADDRESS>`).join('\n')}
</BASICBUYERADDRESS.LIST>`;

    if (dispatchEnabled) {
      // Dispatch-from address — 5 lines as in reference, blanks for unused
      const dispatchLines = [
        d_add,
        d_add2 || `${d_place}-${d_pin}`,
        '', '', '',
      ];
      xml += `
<DISPATCHFROMADDRESS.LIST TYPE="String">
${dispatchLines.map(l => `<DISPATCHFROMADDRESS>${xe(l)}</DISPATCHFROMADDRESS>`).join('\n')}
</DISPATCHFROMADDRESS.LIST>`;
    }

    xml += `
<DATE>${dateval}</DATE>
<REFERENCEDATE></REFERENCEDATE>
<IRNACKDATE>${dateval}</IRNACKDATE>
<VCHSTATUSDATE>${dateval}</VCHSTATUSDATE>
<GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>
<STATENAME>${partyState}</STATENAME>
<COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
<PLACEOFSUPPLY>${partyState}</PLACEOFSUPPLY>
<VOUCHERNUMBER>${xe(voucherNum)}</VOUCHERNUMBER>
<REFERENCE></REFERENCE>
<BILLTOPLACE>${partyPlace}</BILLTOPLACE>`;

    if (dispatchEnabled) {
      xml += `
<DISPATCHFROMNAME>${xe(d_company)}</DISPATCHFROMNAME>
<DISPATCHFROMSTATENAME>${xe(d_state)}</DISPATCHFROMSTATENAME>
<DISPATCHFROMPINCODE>${xe(d_pin)}</DISPATCHFROMPINCODE>
<DISPATCHFROMPLACE>${xe(d_place)}</DISPATCHFROMPLACE>`;
    }

    xml += `
<SHIPTOPLACE>${partyPlace}</SHIPTOPLACE>
<CONSIGNEEGSTIN>${partyGstin}</CONSIGNEEGSTIN>
<CONSIGNEEMAILINGNAME>${partyName}</CONSIGNEEMAILINGNAME>
<CONSIGNEEPINCODE>${partyPin}</CONSIGNEEPINCODE>
<CONSIGNEESTATENAME>${partyState}</CONSIGNEESTATENAME>
<CONSIGNEEPINNUMBER>${xe(consigneePAN)}</CONSIGNEEPINNUMBER>
<CONSIGNEECOUNTRYNAME>India</CONSIGNEECOUNTRYNAME>`;

    if (dispatchEnabled) {
      xml += `
<BASICORDERTERMS.LIST TYPE="String">
<BASICORDERTERMS>Dispatch From:</BASICORDERTERMS>
<BASICORDERTERMS>GSTIN.${xe(d_gstin)}</BASICORDERTERMS>
<BASICORDERTERMS>${xe(d_company)}</BASICORDERTERMS>
<BASICORDERTERMS>${xe(d_add)}</BASICORDERTERMS>
<BASICORDERTERMS>${xe(d_place)}-${xe(d_pin)}</BASICORDERTERMS>
<BASICORDERTERMS>${xe(d_state)} Code:${xe(d_state_code)}</BASICORDERTERMS>
</BASICORDERTERMS.LIST>`;
    }

    xml += `
<BASICBASEPARTYNAME>${partyName}</BASICBASEPARTYNAME>
<NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
<VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
<BASICDATETIMEOFINVOICE>${dateval}</BASICDATETIMEOFINVOICE>
<BASICDATETIMEOFREMOVAL>${dateval}</BASICDATETIMEOFREMOVAL>
<BASICFINALDESTINATION>${finalDest}</BASICFINALDESTINATION>
<BASICSHIPPEDBY>${shippedBy}</BASICSHIPPEDBY>
<BASICORDERREF>${xe(aspVoucherRef)}</BASICORDERREF>
<BASICSHIPVESSELNO>${vehicleNo}</BASICSHIPVESSELNO>
<VCHENTRYMODE>Item Invoice</VCHENTRYMODE>
<DIFFACTUALQTY>Yes</DIFFACTUALQTY>
<ISSECURITYONWHENENTERED>Yes</ISSECURITYONWHENENTERED>
<EFFECTIVEDATE>${dateval}</EFFECTIVEDATE>
<ISGSTOVERRIDDEN>Yes</ISGSTOVERRIDDEN>
<ISELIGIBLEFORITC>Yes</ISELIGIBLEFORITC>
<VCHGSTSTATUSISINCLUDED>Yes</VCHGSTSTATUSISINCLUDED>
<VCHGSTSTATUSISAPPLICABLE>Yes</VCHGSTSTATUSISAPPLICABLE>
<VCHSTATUSISREACCEPHSNSIXONEDONE>Yes</VCHSTATUSISREACCEPHSNSIXONEDONE>
<ISINVOICE>Yes</ISINVOICE>
<ISVATDUTYPAID>Yes</ISVATDUTYPAID>`;

    // E-way bill block — only when dispatch is enabled (matches reference)
    if (dispatchEnabled) {
      const consignorAddrFlat = `${d_company} ${d_add} GSTIN:${d_gstin}`.replace(/\s+/g, ' ').trim();
      xml += `
<EWAYBILLDETAILS.LIST>
<CONSIGNORADDRESSTYPE>${xe(consignorType)}</CONSIGNORADDRESSTYPE>
<CONSIGNORADDRESS.LIST TYPE="String">
<CONSIGNORADDRESS>${xe(consignorAddrFlat)}</CONSIGNORADDRESS>
</CONSIGNORADDRESS.LIST>
<CONSIGNEEADDRESS.LIST TYPE="String">
<CONSIGNEEADDRESS>${xe(consigneeAddr)}</CONSIGNEEADDRESS>
</CONSIGNEEADDRESS.LIST>
<DOCUMENTTYPE>Tax Invoice</DOCUMENTTYPE>
<SUBTYPE>Supply</SUBTYPE>
<CONSIGNEEPINCODE>${partyPin}</CONSIGNEEPINCODE>
<CONSIGNORPLACE>${xe(d_place)}</CONSIGNORPLACE>
<CONSIGNORPINCODE>${xe(d_pin)}</CONSIGNORPINCODE>
<SHIPPEDFROMSTATE>${xe(d_state)}</SHIPPEDFROMSTATE>
<CONSIGNEEPLACE>${partyPlace}</CONSIGNEEPLACE>
<SHIPPEDTOSTATE>${partyState}</SHIPPEDTOSTATE>
<ISCANCELLED>No</ISCANCELLED>
<TRANSPORTDETAILS.LIST>
<TRANSPORTMODE>${transportMode}</TRANSPORTMODE>
<VEHICLENUMBER>${vehicleNo}</VEHICLENUMBER>
<OLDVEHICLETYPE>${vehicleType}</OLDVEHICLETYPE>
<VEHICLETYPE>${vehicleType}</VEHICLETYPE>
<DISTANCE>${distance}</DISTANCE>
</TRANSPORTDETAILS.LIST>
</EWAYBILLDETAILS.LIST>`;
    }

    // Party (debtor) ledger — negative amount (party owes us)
    const totalRound = r0(row.totalRounded || row.total);
    const totalAmt   = r2(row.total);
    const rnd        = r2(totalRound - totalAmt);

    xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${partyName}</LEDGERNAME>
<ISPARTYLEDGER>Yes</ISPARTYLEDGER>
${TAGS.DEEMYES}
<AMOUNT>${-totalRound}</AMOUNT>
<BILLALLOCATIONS.LIST>
<NAME>${xe(voucherNum)}</NAME>
<BILLTYPE>New Ref</BILLTYPE>
<AMOUNT>${-totalRound}</AMOUNT>
</BILLALLOCATIONS.LIST>
</LEDGERENTRIES.LIST>`;

    // Transport + Insurance — only in intra-state ISP vouchers (reference
    // omits both for inter-state and export, matching calculations.js
    // hideTI flag). Emitted before the tax ledgers; each carries its own
    // RATEDETAILS.LIST (CGST/SGST/Cess) since they're separately taxable.
    const wantTI = isIntra && !isExport;
    const transportAmt = r2(row.transportAmt || 0);
    const insuranceAmt = r2(row.insuranceAmt || 0);
    if (wantTI && transportAmt > 0) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(LDR_Transport)}</LEDGERNAME>
<GSTONINEVRDLIGIBLEITC>Applicable</GSTONINEVRDLIGIBLEITC>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
<GSTSOURCETYPE>Ledger</GSTSOURCETYPE>
<GSTLEDGERSOURCE>${xe(LDR_Transport)}</GSTLEDGERSOURCE>
<GSTOVRDNTYPEOFSUPPLY>Services</GSTOVRDNTYPEOFSUPPLY>
<GSTRATEINFERAPPLICABILITY>As per Masters/Company</GSTRATEINFERAPPLICABILITY>
<GSTHSNNAME>${xe(SAC_Transport)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>${xe(LDR_Transport)}</GSTHSNDESCRIPTION>
<GSTHSNINFERAPPLICABILITY>As per Masters/Company</GSTHSNINFERAPPLICABILITY>
${TAGS.DEEMNO}
<AMOUNT>${transportAmt}</AMOUNT>
<VATEXPAMOUNT>${transportAmt}</VATEXPAMOUNT>
${rates.cgst}
${rates.sgst}
${rates.cess}
</LEDGERENTRIES.LIST>`;
    }
    if (wantTI && insuranceAmt > 0) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(LDR_Insurance)}</LEDGERNAME>
<GSTONINEVRDLIGIBLEITC>Applicable</GSTONINEVRDLIGIBLEITC>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
<GSTSOURCETYPE>Ledger</GSTSOURCETYPE>
<GSTLEDGERSOURCE>${xe(LDR_Insurance)}</GSTLEDGERSOURCE>
<GSTOVRDNTYPEOFSUPPLY>Services</GSTOVRDNTYPEOFSUPPLY>
<GSTRATEINFERAPPLICABILITY>As per Masters/Company</GSTRATEINFERAPPLICABILITY>
<GSTHSNNAME>${xe(SAC_Insurance)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>${xe(LDR_Insurance)}</GSTHSNDESCRIPTION>
<GSTHSNINFERAPPLICABILITY>As per Masters/Company</GSTHSNINFERAPPLICABILITY>
${TAGS.DEEMNO}
<AMOUNT>${insuranceAmt}</AMOUNT>
<VATEXPAMOUNT>${insuranceAmt}</VATEXPAMOUNT>
${rates.cgst}
${rates.sgst}
${rates.cess}
</LEDGERENTRIES.LIST>`;
    }

    // Tax ledgers
    if (isExport) {
      // Export — no tax ledger (rate 0)
    } else if (isIntra) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_CGST)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(row.cgst || 0)}</AMOUNT>
<VATEXPAMOUNT>${r2(row.cgst || 0)}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_SGST)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(row.sgst || 0)}</AMOUNT>
<VATEXPAMOUNT>${r2(row.sgst || 0)}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    } else {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_IGST)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(row.igst || 0)}</AMOUNT>
<VATEXPAMOUNT>${r2(row.igst || 0)}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    }

    // TCS
    if (tcs && row.tcsamt && row.tcsamt > 0) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_TCS)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(row.tcsamt)}</AMOUNT>
<VATEXPAMOUNT>${r2(row.tcsamt)}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    }

    // Round-off — always emit when enabled (matches reference; every
    // voucher carries a Round Off line, even with a 0 amount). The sign
    // reflects the adjustment to reach the rounded total; ISDEEMEDPOSITIVE
    // is No and the AMOUNT carries the sign. The previous
    // `Math.abs(rnd) > 0.001` gate dropped the line on whole-rupee totals.
    if (tlyrnd) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Round_LDR)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(rnd)}</AMOUNT>
<VATEXPAMOUNT>${r2(rnd)}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    }

    // Inventory — one entry per lot (Cardamom) when detailed; else one aggregate
    xml += '\n';
    if (detailed && Array.isArray(row.lots)) {
      for (const lot of row.lots) {
        xml += `
<ALLINVENTORYENTRIES.LIST>
<STOCKITEMNAME>${xe(Item_Card)}</STOCKITEMNAME>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
<HSNSOURCETYPE>Stock Item</HSNSOURCETYPE>
<HSNITEMSOURCE>${xe(Item_Card)}</HSNITEMSOURCE>
<GSTOVRDNSTOREDNATURE>${cardNature}</GSTOVRDNSTOREDNATURE>
<GSTOVRDNTYPEOFSUPPLY>Goods</GSTOVRDNTYPEOFSUPPLY>
<GSTHSNNAME>${xe(HSN_Card)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>${xe(Item_Card)}</GSTHSNDESCRIPTION>
<BASICPACKAGEMARKS>${xe(lot.lot || '')}</BASICPACKAGEMARKS>
<BASICNUMPACKAGES>${r0(lot.bag)}</BASICNUMPACKAGES>
${TAGS.DEEMNO}
<RATE>${r2(lot.rate)}/Kgs.</RATE>
<AMOUNT>${r2(lot.amount)}</AMOUNT>
<ACTUALQTY>${r2(lot.qty)}Kgs.</ACTUALQTY>
<BILLEDQTY>${r2(lot.qty)}Kgs.</BILLEDQTY>
<BATCHALLOCATIONS.LIST>
<GODOWNNAME>Main Location</GODOWNNAME>
<BATCHNAME>Primary Batch</BATCHNAME>
<DESTINATIONGODOWNNAME>Main Location</DESTINATIONGODOWNNAME>
<AMOUNT>${r2(lot.amount)}</AMOUNT>
<ACTUALQTY>${r2(lot.qty)}Kgs.</ACTUALQTY>
<BILLEDQTY>${r2(lot.qty)}Kgs.</BILLEDQTY>
</BATCHALLOCATIONS.LIST>
<ACCOUNTINGALLOCATIONS.LIST>
<LEDGERNAME>${xe(cardLedger)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(lot.amount)}</AMOUNT>
</ACCOUNTINGALLOCATIONS.LIST>
${isIntra && !isExport ? `${rates.cgst}\n${rates.sgst}` : (isExport ? '' : rates.igst)}
${rates.cess}
</ALLINVENTORYENTRIES.LIST>`;
      }
    }

    // Gunny inventory — one aggregate entry across all lots if any
    const totalGunnyBags = r0(row.gunnyBags || (Array.isArray(row.lots) ? row.lots.reduce((s, l) => s + Number(l.bag || 0), 0) : 0));
    const totalGunnyAmt  = r2(row.gunnyAmt || (totalGunnyBags * GunnyRate));
    if (totalGunnyAmt > 0) {
      xml += `
<ALLINVENTORYENTRIES.LIST>
<STOCKITEMNAME>${xe(Item_Gunny)}</STOCKITEMNAME>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
<GSTOVRDNSTOREDNATURE>${cardNature}</GSTOVRDNSTOREDNATURE>
<GSTOVRDNTYPEOFSUPPLY>Goods</GSTOVRDNTYPEOFSUPPLY>
<GSTHSNNAME>${xe(HSN_Gunny)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>${xe(Item_Gunny)}</GSTHSNDESCRIPTION>
${TAGS.DEEMNO}
<RATE>${r0(GunnyRate)}/Nos.</RATE>
<AMOUNT>${totalGunnyAmt}</AMOUNT>
<ACTUALQTY>${totalGunnyBags}Nos.</ACTUALQTY>
<BATCHALLOCATIONS.LIST>
<GODOWNNAME>Main Location</GODOWNNAME>
<BATCHNAME>Primary Batch</BATCHNAME>
<DESTINATIONGODOWNNAME>Main Location</DESTINATIONGODOWNNAME>
<AMOUNT>${totalGunnyAmt}</AMOUNT>
<ACTUALQTY>${totalGunnyBags}Nos.</ACTUALQTY>
</BATCHALLOCATIONS.LIST>
<ACCOUNTINGALLOCATIONS.LIST>
<LEDGERNAME>${xe(gunnyLedger)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${totalGunnyAmt}</AMOUNT>
</ACCOUNTINGALLOCATIONS.LIST>
${isIntra && !isExport ? `${rates.cgst}\n${rates.sgst}` : (isExport ? '' : rates.igst)}
${rates.cess}
</ALLINVENTORYENTRIES.LIST>`;
    }

    xml += `\n${TAGS.ENDVOUCHER}`;
  }

  xml += '\n' + endEnvelope();
  const BOM = '\uFEFF';
  return BOM + xml.replace(/\r?\n/g, '\r\n');
}

// =====================================================================
// 1B. ASP SALES VOUCHERS (sister-company internal transfer to ISP)
// =====================================================================
//
// Matches reference ASP_SALE_TRANS-03.xml — leaner format. The buyer
// is ALWAYS ISP (the sister/main company), so address/GSTIN come from
// the home-company config. Lot rates/amounts come from the ASP-side
// planter calculation (`lots.asp_prate` / `lots.asp_puramt`).
//
// Input shape (rows from buildSalesAspRows):
//   rows = [{
//     ano, date, sale ('L'|'I'|'E'), invo (= ASP invoice number),
//     ispPartyName, ispAddress (joined string), ispPlace, ispPin,
//     ispGstin, lots: [{lot, bag, qty, rate, amount}, ...],
//     amounttot, gunnyAmt, gunnyBags, igst, cgst, sgst,
//     total, totalRounded,
//   }, ...]
//
function generSalesAspXML(rows, cfg, opts = {}) {
  const company       = opts.companyName || cfgGet(cfg, 'tally_asp_company_name', 'Amazing Spice Park Private Limited');
  const season        = opts.season || cfgGet(cfg, 'tally_season', cfgGet(cfg, 'season_code', '2026-27'));
  const separator     = opts.separator || cfgGet(cfg, 'tally_separator', '/');
  const ainvPrefix    = cfgGet(cfg, 'tally_ainv_prefix', 'ASP/');
  const detailed      = cfgBool(cfg, 'tally_detailed', true);
  // Always emit the Round Off ledger (default ON) — see generSalesIspXML.
  const tlyrnd        = cfgBool(cfg, 'tally_round_enabled', true);

  // ASP's home-state code (for intra/inter detection w.r.t. the buyer)
  const intra         = cfgGet(cfg, 'tally_state_code_amazing', '32');

  // Ledgers — ASP's books use the same names by convention; if user
  // wants ASP-specific ledger names, add tally_asp_* keys later.
  const SalesInter   = cfgGet(cfg, 'tally_sales_inter',  'Cardamom Inter-State Sales');
  const SalesIntra   = cfgGet(cfg, 'tally_sales_intra',  'Cardamom Local Sales');
  const SalesExport  = cfgGet(cfg, 'tally_sales_export', 'Cardamom Sales - Export');
  const GunnyInter   = cfgGet(cfg, 'tally_gunny_inter',  'Gunny Interstate Sales');
  const GunnyIntra   = cfgGet(cfg, 'tally_gunny_intra',  'Gunny Local Sales');
  const GunnyExport  = cfgGet(cfg, 'tally_gunny_export', 'Gunny Sales - Export');
  const Tax_CGST     = cfgGet(cfg, 'tally_cgst', 'OUTPUT CGST 2.5%');
  const Tax_SGST     = cfgGet(cfg, 'tally_sgst', 'OUTPUT SGST 2.5%');
  const Tax_IGST     = cfgGet(cfg, 'tally_igst', 'OUTPUT IGST 5%');
  const Round_LDR    = cfgGet(cfg, 'tally_round', 'Round On/Off');
  const Item_Card    = cfgGet(cfg, 'tally_item_cardamom', 'Cardamom');
  const Item_Gunny   = cfgGet(cfg, 'tally_item_gunny',    'Gunny');
  const HSN_Card     = cfgGet(cfg, 'tally_hsn_cardamom',  '09083120');
  const HSN_Gunny    = cfgGet(cfg, 'tally_hsn_gunny',     '63051040');
  const GunnyRate    = cfgNum(cfg, 'tally_gunny_rate',     165);

  // ASP's own dispatch-from address (own premises in Kerala). Reference
  // shows full address with SBL line. We reuse the home company's KL
  // address fields for this.
  const own_addr1    = cfgGet(cfg, 's_dispatch_address1', '650,Ward 6, Ellikkanam, Nedumkandam');
  const own_addr2    = cfgGet(cfg, 's_dispatch_address2', 'Idukki, Kerala, 685553');
  const own_sbl      = cfgGet(cfg, 's_sbl', '');
  const own_place    = cfgGet(cfg, 'tally_dispatch_place', 'NEDUMKANDAM');
  const own_pin      = cfgGet(cfg, 'tally_dispatch_pin', '685553');
  const own_state    = cfgGet(cfg, 'tally_dispatch_state', 'Kerala');
  const own_company  = cfgGet(cfg, 's_company', cfgGet(cfg, 's_short_name', 'AMAZING SPICE PARK PRIVATE LIMITED'));

  let xml = '\n' + startEnvelope(company, 'Vouchers');

  for (const row of rows) {
    const dateval     = toTallyDate(row.date);
    const ispName     = xe(row.ispPartyName);
    const ispGstin    = xe(row.ispGstin);
    const ispState    = xe(findState(row.ispGstin));
    const ispPin      = xe(row.ispPin || '');
    const ispPlace    = xe(row.ispPlace || '');
    const ispAddrLines = _addrLines(row.ispAddress, row.ispPlace);
    const isIntra     = String(row.ispGstin || '').slice(0, 2) === String(intra);
    // ASP→ISP transfer is always inter-state (Kerala→TN by default), so
    // the voucher number's sale letter is always 'I'. The upstream
    // invoice's `sale` (which reflects the ISP→outside-customer leg —
    // could be L/I/E) is preserved on the row but not used in the
    // voucher number. Reference uses ASP/I-61, ASP/I-62, ... regardless.
    const sale        = 'I';
    const isExport    = false;
    const invoNo      = String(row.invo || '').trim();
    const taxNm       = `${sale}${separator}${invoNo}`;
    const voucherNum  = `${ainvPrefix}${taxNm}/${season}`;

    const rates       = rateDetails(cfgNum(cfg, 'tally_gst_rate', 5));

    const cardLedger = isIntra ? SalesIntra : SalesInter;
    const cardNature = isIntra ? 'Local Sales - Taxable' : 'Interstate Sales - Taxable';
    const gunnyLedger = isIntra ? GunnyIntra : GunnyInter;

    const startVoucher = `<VOUCHER VCHTYPE="Sales" ACTION="Create" OBJVIEW="Invoice Voucher View">`;

    xml += `\n${startVoucher}
<PARTYNAME>${ispName}</PARTYNAME>
<ADDRESS.LIST TYPE="String">
${ispAddrLines.map(l => `<ADDRESS>${xe(l)}</ADDRESS>`).join('\n')}
</ADDRESS.LIST>
<PARTYGSTIN>${ispGstin}</PARTYGSTIN>
<PARTYLEDGERNAME>${ispName}</PARTYLEDGERNAME>
<PARTYMAILINGNAME>${ispName}</PARTYMAILINGNAME>
<PARTYPINCODE>${ispPin}</PARTYPINCODE>
<DISPATCHFROMADDRESS.LIST TYPE="String">
<DISPATCHFROMADDRESS>${xe(own_addr1)}</DISPATCHFROMADDRESS>
<DISPATCHFROMADDRESS>${xe(own_addr2)}</DISPATCHFROMADDRESS>
<DISPATCHFROMADDRESS>${xe(own_sbl ? 'SBL:' + own_sbl : '')}</DISPATCHFROMADDRESS>
<DISPATCHFROMADDRESS> </DISPATCHFROMADDRESS>
<DISPATCHFROMADDRESS> </DISPATCHFROMADDRESS>
</DISPATCHFROMADDRESS.LIST>
<DATE>${dateval}</DATE>
<REFERENCEDATE></REFERENCEDATE>
<VCHSTATUSDATE>${dateval}</VCHSTATUSDATE>
<GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>
<STATENAME>${ispState}</STATENAME>
<COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
<PLACEOFSUPPLY>${ispState}</PLACEOFSUPPLY>
<VOUCHERNUMBER>${xe(voucherNum)}</VOUCHERNUMBER>
<REFERENCE></REFERENCE>
<BILLTOPLACE>${ispPlace}</BILLTOPLACE>
<DISPATCHFROMNAME>${xe(own_company)}</DISPATCHFROMNAME>
<DISPATCHFROMSTATENAME>${xe(own_state)}</DISPATCHFROMSTATENAME>
<DISPATCHFROMPINCODE>${xe(own_pin)}</DISPATCHFROMPINCODE>
<DISPATCHFROMPLACE>${xe(own_place)}</DISPATCHFROMPLACE>
<BASICBASEPARTYNAME>${ispName}</BASICBASEPARTYNAME>
<NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
<VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
<BASICDATETIMEOFINVOICE>${dateval}</BASICDATETIMEOFINVOICE>
<BASICDATETIMEOFREMOVAL>${dateval}</BASICDATETIMEOFREMOVAL>
<BASICFINALDESTINATION>${xe(own_place)}</BASICFINALDESTINATION>
<BASICSHIPVESSELNO></BASICSHIPVESSELNO>
<VCHENTRYMODE>Item Invoice</VCHENTRYMODE>
<DIFFACTUALQTY>Yes</DIFFACTUALQTY>
<ISSECURITYONWHENENTERED>Yes</ISSECURITYONWHENENTERED>
<EFFECTIVEDATE>${dateval}</EFFECTIVEDATE>
<ISGSTOVERRIDDEN>Yes</ISGSTOVERRIDDEN>
<ISELIGIBLEFORITC>Yes</ISELIGIBLEFORITC>
<VCHGSTSTATUSISINCLUDED>Yes</VCHGSTSTATUSISINCLUDED>
<VCHGSTSTATUSISAPPLICABLE>Yes</VCHGSTSTATUSISAPPLICABLE>
<VCHSTATUSISREACCEPHSNSIXONEDONE>Yes</VCHSTATUSISREACCEPHSNSIXONEDONE>
<ISINVOICE>Yes</ISINVOICE>
<ISVATDUTYPAID>Yes</ISVATDUTYPAID>`;

    // Party (debtor = ISP) ledger
    const totalRound = r0(row.totalRounded || row.total);
    const totalAmt   = r2(row.total);
    const rnd        = r2(totalRound - totalAmt);

    xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${ispName}</LEDGERNAME>
<ISPARTYLEDGER>Yes</ISPARTYLEDGER>
${TAGS.DEEMYES}
<AMOUNT>${-totalRound}</AMOUNT>
<BILLALLOCATIONS.LIST>
<NAME>${xe(voucherNum)}</NAME>
<BILLTYPE>New Ref</BILLTYPE>
<AMOUNT>${-totalRound}</AMOUNT>
</BILLALLOCATIONS.LIST>
</LEDGERENTRIES.LIST>`;

    // Tax ledgers
    if (isIntra) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_CGST)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(row.cgst || 0)}</AMOUNT>
<VATEXPAMOUNT>${r2(row.cgst || 0)}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_SGST)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(row.sgst || 0)}</AMOUNT>
<VATEXPAMOUNT>${r2(row.sgst || 0)}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    } else {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_IGST)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(row.igst || 0)}</AMOUNT>
<VATEXPAMOUNT>${r2(row.igst || 0)}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    }

    // Round-off — always emit when enabled (even with a 0 amount), so
    // whole-rupee totals still carry the Round Off line.
    if (tlyrnd) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Round_LDR)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(rnd)}</AMOUNT>
<VATEXPAMOUNT>${r2(rnd)}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    }

    // Inventory — one entry per lot
    xml += '\n';
    if (detailed && Array.isArray(row.lots)) {
      for (const lot of row.lots) {
        xml += `
<ALLINVENTORYENTRIES.LIST>
<STOCKITEMNAME>${xe(Item_Card)}</STOCKITEMNAME>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
<HSNSOURCETYPE>Stock Item</HSNSOURCETYPE>
<HSNITEMSOURCE>${xe(Item_Card)}</HSNITEMSOURCE>
<GSTOVRDNSTOREDNATURE>${cardNature}</GSTOVRDNSTOREDNATURE>
<GSTOVRDNTYPEOFSUPPLY>Goods</GSTOVRDNTYPEOFSUPPLY>
<GSTHSNNAME>${xe(HSN_Card)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>${xe(Item_Card)}</GSTHSNDESCRIPTION>
<BASICPACKAGEMARKS>${xe(lot.lot || '')}</BASICPACKAGEMARKS>
<BASICNUMPACKAGES>${r0(lot.bag)}</BASICNUMPACKAGES>
${TAGS.DEEMNO}
<RATE>${r2(lot.rate)}/Kgs.</RATE>
<AMOUNT>${r2(lot.amount)}</AMOUNT>
<ACTUALQTY>${r2(lot.qty)}Kgs.</ACTUALQTY>
<BILLEDQTY>${r2(lot.qty)}Kgs.</BILLEDQTY>
<BATCHALLOCATIONS.LIST>
<GODOWNNAME>Main Location</GODOWNNAME>
<BATCHNAME>Primary Batch</BATCHNAME>
<DESTINATIONGODOWNNAME>Main Location</DESTINATIONGODOWNNAME>
<AMOUNT>${r2(lot.amount)}</AMOUNT>
<ACTUALQTY>${r2(lot.qty)}Kgs.</ACTUALQTY>
<BILLEDQTY>${r2(lot.qty)}Kgs.</BILLEDQTY>
</BATCHALLOCATIONS.LIST>
<ACCOUNTINGALLOCATIONS.LIST>
<LEDGERNAME>${xe(cardLedger)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(lot.amount)}</AMOUNT>
</ACCOUNTINGALLOCATIONS.LIST>
${isIntra ? `${rates.cgst}\n${rates.sgst}` : rates.igst}
${rates.cess}
</ALLINVENTORYENTRIES.LIST>`;
      }
    }

    // Gunny aggregate
    const totalGunnyBags = r0(row.gunnyBags || (Array.isArray(row.lots) ? row.lots.reduce((s, l) => s + Number(l.bag || 0), 0) : 0));
    const totalGunnyAmt  = r2(row.gunnyAmt || (totalGunnyBags * GunnyRate));
    if (totalGunnyAmt > 0) {
      xml += `
<ALLINVENTORYENTRIES.LIST>
<STOCKITEMNAME>${xe(Item_Gunny)}</STOCKITEMNAME>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
<GSTOVRDNSTOREDNATURE>${cardNature}</GSTOVRDNSTOREDNATURE>
<GSTOVRDNTYPEOFSUPPLY>Goods</GSTOVRDNTYPEOFSUPPLY>
<GSTHSNNAME>${xe(HSN_Gunny)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>${xe(Item_Gunny)}</GSTHSNDESCRIPTION>
${TAGS.DEEMNO}
<RATE>${r0(GunnyRate)}/Nos.</RATE>
<AMOUNT>${totalGunnyAmt}</AMOUNT>
<ACTUALQTY>${totalGunnyBags}Nos.</ACTUALQTY>
<BATCHALLOCATIONS.LIST>
<GODOWNNAME>Main Location</GODOWNNAME>
<BATCHNAME>Primary Batch</BATCHNAME>
<DESTINATIONGODOWNNAME>Main Location</DESTINATIONGODOWNNAME>
<AMOUNT>${totalGunnyAmt}</AMOUNT>
<ACTUALQTY>${totalGunnyBags}Nos.</ACTUALQTY>
</BATCHALLOCATIONS.LIST>
<ACCOUNTINGALLOCATIONS.LIST>
<LEDGERNAME>${xe(gunnyLedger)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${totalGunnyAmt}</AMOUNT>
</ACCOUNTINGALLOCATIONS.LIST>
${isIntra ? `${rates.cgst}\n${rates.sgst}` : rates.igst}
${rates.cess}
</ALLINVENTORYENTRIES.LIST>`;
    }

    xml += `\n${TAGS.ENDVOUCHER}`;
  }

  xml += '\n' + endEnvelope();
  const BOM = '\uFEFF';
  return BOM + xml.replace(/\r?\n/g, '\r\n');
}

// =====================================================================
// ISP PURCHASE (mirror of ASP→ISP transfer, ISP-side books)
// =====================================================================
//
// For every ASP sale voucher (ASP→ISP internal transfer), the ISP company
// needs a matching purchase voucher in its books. Same line items, same
// amounts, but viewed from the buyer side:
//   • PARTYNAME = ASP (the supplier)
//   • Imports into the ISP Tally company
//   • VOUCHERNUMBER = same as the ASP sale (e.g. ASP/I-61/26-27) — this is
//     the cross-reference key. NOT a fresh ISP-prefixed number.
//   • Inventory AMOUNTs are negative (purchase debits stock)
//   • Tax ledger AMOUNTs are negative (input GST claimed)
//   • Round-off sign is flipped vs the sale (preRound − rounded)
//   • Single accounting ledger: "Trade Purchase From Dealer" (no
//     -Local/-Inter_State variant — always inter-state Kerala→TN)
//   • Sales-only fields (DISPATCHFROMNAME, BILLTOPLACE, ORDERREF, etc.)
//     are absent
//
// Input shape: same as buildSalesAspRows produces. We re-use that builder
// directly — no new builder needed.
//
function generIspPurchaseXML(rows, cfg, opts = {}) {
  // Imports into ISP company (the buyer's books)
  const company       = opts.companyName || cfgGet(cfg, 'tally_company_name', cfgGet(cfg, 'short_name', 'Ideal Spices Private Limited'));
  const tlyrnd        = cfgBool(cfg, 'flag_tally_round', true);

  // Ledger names — Trade Purchase From Dealer is the single accounting
  // ledger used in the reference for ALL inventory entries (cardamom +
  // gunny). Configurable via tally_isp_purchase_ledger.
  const Purchase_LDR = cfgGet(cfg, 'tally_isp_purchase_ledger', 'Trade Purchase From Dealer');
  const Tax_IGST_IN  = cfgGet(cfg, 'tally_igst_input',          'INPUT IGST 5%');
  const Round_LDR    = cfgGet(cfg, 'tally_round',               'Round On/Off');
  const Item_Card    = cfgGet(cfg, 'tally_item_cardamom',       'Cardamom');
  const Item_Gunny   = cfgGet(cfg, 'tally_item_gunny',          'Gunny');
  const HSN_Card     = cfgGet(cfg, 'tally_hsn_cardamom',        '09083120');
  const HSN_Gunny    = cfgGet(cfg, 'tally_hsn_gunny',           '63051040');

  // Sister/ASP party identity — fetched from sister-company cfg.
  const aspName    = cfgGet(cfg, 's_company',
                       cfgGet(cfg, 's_short_name', 'AMAZING SPICE PARK PRIVATE LIMITED'));
  const aspAddr1   = cfgGet(cfg, 's_address1', cfgGet(cfg, 's_dispatch_address1', ''));
  const aspAddr2   = cfgGet(cfg, 's_address2', cfgGet(cfg, 's_dispatch_address2', ''));
  const aspPlace   = cfgGet(cfg, 's_place',    cfgGet(cfg, 'tally_dispatch_place', 'NEDUMKANDAM'));
  const aspPin     = cfgGet(cfg, 's_pin',      cfgGet(cfg, 'tally_dispatch_pin', '685553'));
  const aspState   = cfgGet(cfg, 's_state',    cfgGet(cfg, 'tally_dispatch_state', 'Kerala'));
  const aspGstin   = cfgGet(cfg, 's_gstin',    '');

  // Address lines for the ASP party. Reference shows two lines:
  //   line 1 = full street address
  //   line 2 = town/place
  const aspAddrLines = _addrLines(aspAddr1, aspPlace || aspAddr2);

  // Place of supply for the receiving company (ISP = TN by default). Sale
  // letter on the matching ASP sale voucher is always 'I' (inter-state),
  // so place of supply = TN.
  const ispPlaceOfSupply = cfgGet(cfg, 'tally_home_state', 'Tamil Nadu');

  // CONSIGNEEPINNUMBER in the reference is the PAN-portion of the
  // CONSIGNEEGSTIN (chars 3-12). Reference shows "ABDCA2636B" extracted
  // from "32ABDCA2636B1ZE".
  const aspPan = aspGstin ? String(aspGstin).slice(2, 12) : '';

  let xml = '\n' + startEnvelope(company, 'Vouchers');

  for (const row of rows) {
    const dateval     = toTallyDate(row.date);
    // Same voucher number as the matching ASP sale — this is the
    // cross-reference. The builder already produced the per-row data
    // including invo and sale='I' implicit; we re-derive the number here
    // exactly like generSalesAspXML does, so the two stay in lockstep.
    const ainvPrefix  = cfgGet(cfg, 'tally_ainv_prefix', 'ASP/');
    const separator   = opts.separator || cfgGet(cfg, 'tally_separator', '/');
    const season      = opts.season    || cfgGet(cfg, 'tally_season', cfgGet(cfg, 'season_code', '2026-27'));
    const sale        = 'I';  // always inter-state for ASP→ISP
    const invoNo      = String(row.invo || '').trim();
    const taxNm       = `${sale}${separator}${invoNo}`;
    const voucherNum  = `${ainvPrefix}${taxNm}/${season}`;

    // Pre-round total (matches the ASP sale's pre-round total). The
    // builder computed `total` as r2(taxableTotal + cgst + sgst + igst).
    const preRound    = r2(row.total || 0);
    const rounded     = r0(preRound);
    // Purchase round-off has opposite sign vs sale (preRound − rounded)
    const roundOff    = r2(preRound - rounded);

    // For the party AMOUNT, Tally expects the rounded total (positive,
    // since we owe ASP). The reference shows e.g. <AMOUNT>985352</AMOUNT>.
    const partyAmt    = tlyrnd ? rounded : preRound;

    // IGST always (Kerala→TN inter-state). The reference shows IGST
    // ledger AMOUNT as negative (input tax). cgst/sgst should be 0 here
    // for the standard ASP→ISP flow.
    const igstAmt     = r2(row.igst || 0);
    const cgstAmt     = r2(row.cgst || 0);
    const sgstAmt     = r2(row.sgst || 0);

    const startVoucher = `<VOUCHER VCHTYPE="Purchase" ACTION="Create" OBJVIEW="Invoice Voucher View">`;

    xml += `\n${startVoucher}
<PARTYNAME>${xe(aspName)}</PARTYNAME>
<ADDRESS.LIST TYPE="String">
${aspAddrLines.map(l => `<ADDRESS>${xe(l)}</ADDRESS>`).join('\n')}
</ADDRESS.LIST>
<PARTYGSTIN>${xe(aspGstin)}</PARTYGSTIN>
<PARTYLEDGERNAME>${xe(aspName)}</PARTYLEDGERNAME>
<PARTYMAILINGNAME>${xe(aspName)}</PARTYMAILINGNAME>
<PARTYPINCODE>${xe(aspPin)}</PARTYPINCODE>
<BASICBUYERNAME>${xe(aspName)}</BASICBUYERNAME>
<BASICBUYERADDRESS.LIST TYPE="String">
${aspAddrLines.map(l => `<BASICBUYERADDRESS>${xe(l)}</BASICBUYERADDRESS>`).join('\n')}
</BASICBUYERADDRESS.LIST>
<DATE>${dateval}</DATE>
<REFERENCEDATE></REFERENCEDATE>
<IRNACKDATE>${dateval}</IRNACKDATE>
<VCHSTATUSDATE>${dateval}</VCHSTATUSDATE>
<GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>
<STATENAME>${xe(aspState)}</STATENAME>
<COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
<PLACEOFSUPPLY>${xe(ispPlaceOfSupply)}</PLACEOFSUPPLY>
<VOUCHERNUMBER>${xe(voucherNum)}</VOUCHERNUMBER>
<REFERENCE></REFERENCE>
<CONSIGNEEGSTIN>${xe(aspGstin)}</CONSIGNEEGSTIN>
<CONSIGNEEMAILINGNAME>${xe(aspName)}</CONSIGNEEMAILINGNAME>
<CONSIGNEEPINCODE>${xe(aspPin)}</CONSIGNEEPINCODE>
<CONSIGNEESTATENAME>${xe(aspState)}</CONSIGNEESTATENAME>
<CONSIGNEEPINNUMBER>${xe(aspPan)}</CONSIGNEEPINNUMBER>
<CONSIGNEECOUNTRYNAME>India</CONSIGNEECOUNTRYNAME>
<BASICBASEPARTYNAME>${xe(aspName)}</BASICBASEPARTYNAME>
<NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
<VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>
<BASICDATETIMEOFINVOICE>${dateval}</BASICDATETIMEOFINVOICE>
<BASICDATETIMEOFREMOVAL>${dateval}</BASICDATETIMEOFREMOVAL>
<VCHENTRYMODE>Item Invoice</VCHENTRYMODE>
<DIFFACTUALQTY>Yes</DIFFACTUALQTY>
<EFFECTIVEDATE>${dateval}</EFFECTIVEDATE>
<ISELIGIBLEFORITC>Yes</ISELIGIBLEFORITC>
<VCHSTATUSISREACCEPHSNSIXONEDONE>Yes</VCHSTATUSISREACCEPHSNSIXONEDONE>
<VCHGSTSTATUSISAPPLICABLE>Yes</VCHGSTSTATUSISAPPLICABLE>
<VCHGSTSTATUSISOVERRDN>Yes</VCHGSTSTATUSISOVERRDN>
<ISINVOICE>Yes</ISINVOICE>
<ISVATDUTYPAID>Yes</ISVATDUTYPAID>
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(aspName)}</LEDGERNAME>
<ISPARTYLEDGER>Yes</ISPARTYLEDGER>
${TAGS.DEEMNO}
<AMOUNT>${partyAmt}</AMOUNT>
<BILLALLOCATIONS.LIST>
<NAME>${xe(voucherNum)}</NAME>
<BILLTYPE>New Ref</BILLTYPE>
<AMOUNT>${partyAmt}</AMOUNT>
</BILLALLOCATIONS.LIST>
</LEDGERENTRIES.LIST>`;

    // Tax ledger — IGST only for inter-state (the standard ASP→ISP path).
    // CGST/SGST emitted only if the row was somehow computed as intra
    // (rare; would happen only if ASP and ISP shared a state).
    if (igstAmt) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_IGST_IN)}</LEDGERNAME>
${TAGS.DEEMYES}
<AMOUNT>${-igstAmt}</AMOUNT>
<VATEXPAMOUNT>${-igstAmt}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    }
    if (cgstAmt) {
      const Tax_CGST_IN = cfgGet(cfg, 'tally_cgst_input', 'INPUT CGST 2.5%');
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_CGST_IN)}</LEDGERNAME>
${TAGS.DEEMYES}
<AMOUNT>${-cgstAmt}</AMOUNT>
<VATEXPAMOUNT>${-cgstAmt}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    }
    if (sgstAmt) {
      const Tax_SGST_IN = cfgGet(cfg, 'tally_sgst_input', 'INPUT SGST 2.5%');
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_SGST_IN)}</LEDGERNAME>
${TAGS.DEEMYES}
<AMOUNT>${-sgstAmt}</AMOUNT>
<VATEXPAMOUNT>${-sgstAmt}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    }

    // Round-off — always emit when rounding is enabled (even with a 0
    // amount), matching the sales/RD/URD generators. Sign is preRound −
    // rounded (opposite of the sale voucher). DEEMEDPOSITIVE=Yes for
    // purchase round-off (matches reference). Gating on tlyrnd (rather
    // than roundOff !== 0) keeps the voucher balanced: when rounding is
    // off, partyAmt = preRound and no round-off line is needed.
    if (tlyrnd) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Round_LDR)}</LEDGERNAME>
${TAGS.DEEMYES}
<AMOUNT>${roundOff}</AMOUNT>
<VATEXPAMOUNT>${roundOff}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    }

    // Inventory entries — one per cardamom lot + one for gunny
    // (aggregate, since gunny is sold by bag count not lot). All AMOUNTs
    // negative (purchase debits stock).
    xml += '\n';
    for (const lot of (row.lots || [])) {
      const amt = r2(lot.amount || 0);
      if (!amt) continue;
      xml += `
<ALLINVENTORYENTRIES.LIST>
<STOCKITEMNAME>${xe(Item_Card)}</STOCKITEMNAME>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
<HSNSOURCETYPE>Stock Item</HSNSOURCETYPE>
<HSNITEMSOURCE>${xe(Item_Card)}</HSNITEMSOURCE>
<GSTOVRDNSTOREDNATURE>Interstate Purchase - Taxable</GSTOVRDNSTOREDNATURE>
<GSTOVRDNTYPEOFSUPPLY>Goods</GSTOVRDNTYPEOFSUPPLY>
<GSTHSNNAME>${xe(HSN_Card)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>${xe(Item_Card)}</GSTHSNDESCRIPTION>
<BASICPACKAGEMARKS>${xe(lot.lot || '')}</BASICPACKAGEMARKS>
<BASICNUMPACKAGES>${r0(lot.bag)}</BASICNUMPACKAGES>
${TAGS.DEEMYES}
<RATE>${r0(lot.rate)}/Kgs.</RATE>
<AMOUNT>${-amt}</AMOUNT>
<ACTUALQTY>${r2(lot.qty)}Kgs.</ACTUALQTY>
<BILLEDQTY>${r2(lot.qty)}Kgs.</BILLEDQTY>
<BATCHALLOCATIONS.LIST>
<GODOWNNAME>Main Location</GODOWNNAME>
<BATCHNAME>Primary Batch</BATCHNAME>
<DESTINATIONGODOWNNAME>Main Location</DESTINATIONGODOWNNAME>
<AMOUNT>${-amt}</AMOUNT>
<ACTUALQTY>${r2(lot.qty)}Kgs.</ACTUALQTY>
<BILLEDQTY>${r2(lot.qty)}Kgs.</BILLEDQTY>
</BATCHALLOCATIONS.LIST>
<ACCOUNTINGALLOCATIONS.LIST>
<LEDGERNAME>${xe(Purchase_LDR)}</LEDGERNAME>
${TAGS.DEEMYES}
<AMOUNT>${-amt}</AMOUNT>
</ACCOUNTINGALLOCATIONS.LIST>
<RATEDETAILS.LIST>
<GSTRATEDUTYHEAD>IGST</GSTRATEDUTYHEAD>
<GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>
<GSTRATE>5</GSTRATE>
</RATEDETAILS.LIST>
<RATEDETAILS.LIST>
<GSTRATEDUTYHEAD>Cess</GSTRATEDUTYHEAD>
<GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>
</RATEDETAILS.LIST>
</ALLINVENTORYENTRIES.LIST>`;
    }

    // Gunny inventory — aggregate, with a few structural differences
    // from cardamom (matches reference exactly):
    //   • No HSNSOURCETYPE / HSNITEMSOURCE
    //   • No BASICPACKAGEMARKS / BASICNUMPACKAGES
    //   • No BILLEDQTY (only ACTUALQTY)
    //   • AccAlloc has DEEMEDPOSITIVE=No (vs Yes for cardamom)
    if (row.gunnyAmt && row.gunnyBags) {
      const gAmt = r2(row.gunnyAmt);
      const gBags = r0(row.gunnyBags);
      const gunnyRate = cfgNum(cfg, 'tally_gunny_rate', 165);
      xml += `
<ALLINVENTORYENTRIES.LIST>
<STOCKITEMNAME>${xe(Item_Gunny)}</STOCKITEMNAME>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
<GSTOVRDNSTOREDNATURE>Interstate Purchase - Taxable</GSTOVRDNSTOREDNATURE>
<GSTOVRDNTYPEOFSUPPLY>Goods</GSTOVRDNTYPEOFSUPPLY>
<GSTHSNNAME>${xe(HSN_Gunny)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>${xe(Item_Gunny)}</GSTHSNDESCRIPTION>
${TAGS.DEEMYES}
<RATE>${r0(gunnyRate)}/Nos.</RATE>
<AMOUNT>${-gAmt}</AMOUNT>
<ACTUALQTY>${gBags}Nos.</ACTUALQTY>
<BATCHALLOCATIONS.LIST>
<GODOWNNAME>Main Location</GODOWNNAME>
<BATCHNAME>Primary Batch</BATCHNAME>
<DESTINATIONGODOWNNAME>Main Location</DESTINATIONGODOWNNAME>
<AMOUNT>${-gAmt}</AMOUNT>
<ACTUALQTY>${gBags}Nos.</ACTUALQTY>
</BATCHALLOCATIONS.LIST>
<ACCOUNTINGALLOCATIONS.LIST>
<LEDGERNAME>${xe(Purchase_LDR)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${-gAmt}</AMOUNT>
</ACCOUNTINGALLOCATIONS.LIST>
<RATEDETAILS.LIST>
<GSTRATEDUTYHEAD>IGST</GSTRATEDUTYHEAD>
<GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>
<GSTRATE>5</GSTRATE>
</RATEDETAILS.LIST>
<RATEDETAILS.LIST>
<GSTRATEDUTYHEAD>Cess</GSTRATEDUTYHEAD>
<GSTRATEVALUATIONTYPE>Based on Value</GSTRATEVALUATIONTYPE>
</RATEDETAILS.LIST>
</ALLINVENTORYENTRIES.LIST>`;
    }

    xml += `\n${TAGS.ENDVOUCHER}`;
  }

  xml += '\n' + endEnvelope();
  const BOM = '\uFEFF';
  return BOM + xml.replace(/\r?\n/g, '\r\n');
}

// =====================================================================
// 1. SALES (registered dealer sales — VBA generXML)  [LEGACY]
// =====================================================================
//
// Kept for back-compat; new callers should use generSalesIspXML or
// generSalesAspXML. This original generator is unchanged and still
// referenced by the deprecated `sales` Tally export key in server.js.
//
// Input shape (what the route layer should pass us — already grouped by
// invoice from the invoices table):
//
//   rows = [{
//     ano, date, sale, invo, partyName, address, place, pin,
//     partyGstin, lots: [{lot, bag, qty, rate, amount}, ...],
//     amounttot, gunnyAmt, cgst, sgst, igst, tcsamt, total, totalRounded, rnd,
//   }, ...]
//
function generSalesXML(rows, cfg, opts = {}) {
  const company       = opts.companyName || cfgGet(cfg, 'tally_company_name', cfgGet(cfg, 'short_name', 'Ideal Spices Private Limited'));
  const season        = opts.season || cfgGet(cfg, 'tally_season', cfgGet(cfg, 'season_code', '2026-27'));
  const separator     = opts.separator || cfgGet(cfg, 'tally_separator', '/');
  const invPrefix     = cfgGet(cfg, 'tally_inv_prefix', 'ISP/');
  const ainvPrefix    = cfgGet(cfg, 'tally_ainv_prefix', 'ASP/');
  const amazing       = cfgBool(cfg, 'tally_amazing_mode', false);
  const detailed      = cfgBool(cfg, 'tally_detailed', true);
  const dispatchEnabled = cfgBool(cfg, 'tally_dispatch_from', true);
  const tcs           = cfgBool(cfg, 'tally_tcs_enabled', false);
  // Always emit the Round Off ledger (default ON) — see generSalesIspXML.
  const tlyrnd        = cfgBool(cfg, 'tally_round_enabled', true);
  const intra         = cfgGet(cfg, 'tally_state_code', '33'); // ISP=33, ASP=32 (set in cfg)

  // Ledgers
  const SalesInter   = cfgGet(cfg, 'tally_sales_inter',  'Cardamom Sales 5%');
  const SalesIntra   = cfgGet(cfg, 'tally_sales_intra',  'Cardamom Sales 5% - Local');
  const SalesExport  = cfgGet(cfg, 'tally_sales_export', 'Cardamom Sales - Export');
  const GunnyInter   = cfgGet(cfg, 'tally_gunny_inter',  'Gunny Sales 5%');
  const GunnyIntra   = cfgGet(cfg, 'tally_gunny_intra',  'Gunny Sales 5% - Local');
  const Tax_CGST     = cfgGet(cfg, 'tally_cgst', 'OUTPUT CGST 2.5%');
  const Tax_SGST     = cfgGet(cfg, 'tally_sgst', 'OUTPUT SGST 2.5%');
  const Tax_IGST     = cfgGet(cfg, 'tally_igst', 'OUTPUT IGST 5%');
  const Tax_TCS      = cfgGet(cfg, 'tally_tcs',  'TCS on Sale of Goods');
  const Round_LDR    = cfgGet(cfg, 'tally_round', 'Round Off');
  const Item_Card    = cfgGet(cfg, 'tally_item_cardamom', 'Cardamom');
  const Item_Gunny   = cfgGet(cfg, 'tally_item_gunny',    'Gunny Bag');
  const HSN_Card     = cfgGet(cfg, 'tally_hsn_cardamom',  '09083110');
  const HSN_Gunny    = cfgGet(cfg, 'tally_hsn_gunny',     '63053200');

  // Dispatch-from address (sister-company despatch, ASP source)
  const d_company    = cfgGet(cfg, 'tally_dispatch_company', cfgGet(cfg, 's_short_name', ''));
  const d_add        = cfgGet(cfg, 'tally_dispatch_address', cfgGet(cfg, 's_address1', ''));
  const d_place      = cfgGet(cfg, 'tally_dispatch_place',   cfgGet(cfg, 's_place', ''));
  const d_pin        = cfgGet(cfg, 'tally_dispatch_pin',     cfgGet(cfg, 's_pin', ''));
  const d_state      = cfgGet(cfg, 'tally_dispatch_state',   cfgGet(cfg, 's_state', 'Kerala'));
  const d_gstin      = cfgGet(cfg, 'tally_dispatch_gstin',   cfgGet(cfg, 's_gstin', ''));

  let xml = '\n' + startEnvelope(company, 'Vouchers');

  for (const row of rows) {
    const dateval     = toTallyDate(row.date);
    const partyName   = xe(row.partyName);
    const address     = xe(row.address);
    const place       = xe(row.place);
    const pin         = xe(row.pin);
    const partyGstin  = xe(row.partyGstin);
    const state       = xe(findState(partyGstin));
    const isIntra     = String(partyGstin).slice(0, 2) === String(intra);
    const isExport    = (row.sale || '').toUpperCase() === 'E';
    const sale        = row.sale || 'L';
    const invoNo      = String(row.invo || '').trim();
    const taxNm       = `${sale}${separator}${invoNo}`;
    const voucherNum  = `${amazing ? ainvPrefix : invPrefix}${taxNm}/${season}`;
    const rates       = rateDetails(cfgNum(cfg, 'gst_goods', 5));

    const startVoucher = `<VOUCHER VCHTYPE="Sales" ACTION="Create" OBJVIEW="Invoice Voucher View">`;

    // Inventory entries — one per lot if detailed, else aggregate
    let invEntries = '';
    if (detailed && Array.isArray(row.lots)) {
      for (const lot of row.lots) {
        const ledger = amazing
          ? SalesInter
          : (isExport ? SalesExport : (isIntra ? SalesIntra : SalesInter));
        const stockNature = amazing
          ? 'Interstate Sales - Taxable'
          : (isIntra ? 'Local Sales - Taxable' : 'Interstate Sales - Taxable');
        invEntries += `\n<ALLINVENTORYENTRIES.LIST>
<STOCKITEMNAME>${xe(Item_Card)}</STOCKITEMNAME>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
<HSNSOURCETYPE>Stock Item</HSNSOURCETYPE>
<HSNITEMSOURCE>${xe(Item_Card)}</HSNITEMSOURCE>
<GSTOVRDNSTOREDNATURE>${stockNature}</GSTOVRDNSTOREDNATURE>
<GSTOVRDNTYPEOFSUPPLY>Goods</GSTOVRDNTYPEOFSUPPLY>
<GSTHSNNAME>${xe(HSN_Card)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>${xe(Item_Card)}</GSTHSNDESCRIPTION>
<BASICPACKAGEMARKS>${xe(lot.lot || '')}</BASICPACKAGEMARKS>
<BASICNUMPACKAGES>${r2(lot.bag)}</BASICNUMPACKAGES>
${TAGS.DEEMNO}
<RATE>${r2(lot.rate)}/Kgs.</RATE>
<AMOUNT>${r2(lot.amount)}</AMOUNT>
<ACTUALQTY>${r2(lot.qty)}Kgs.</ACTUALQTY>
<BILLEDQTY>${r2(lot.qty)}Kgs.</BILLEDQTY>
<BATCHALLOCATIONS.LIST>
<GODOWNNAME>Main Location</GODOWNNAME>
<BATCHNAME>Primary Batch</BATCHNAME>
<DESTINATIONGODOWNNAME>Main Location</DESTINATIONGODOWNNAME>
<AMOUNT>${r2(lot.amount)}</AMOUNT>
<ACTUALQTY>${r2(lot.qty)}Kgs.</ACTUALQTY>
<BILLEDQTY>${r2(lot.qty)}Kgs.</BILLEDQTY>
</BATCHALLOCATIONS.LIST>
<ACCOUNTINGALLOCATIONS.LIST>
<LEDGERNAME>${xe(ledger)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(lot.amount)}</AMOUNT>
</ACCOUNTINGALLOCATIONS.LIST>
${amazing ? rates.igst : (isIntra ? `${rates.cgst}\n${rates.sgst}` : rates.igst)}
${rates.cess}
</ALLINVENTORYENTRIES.LIST>`;
      }
    }

    const totalAmt    = r2(row.total);
    const totalRound  = r0(row.totalRounded || row.total);
    const rnd         = r2(totalRound - totalAmt);
    const gunny       = r2(row.gunnyAmt || 0);
    const amounttot   = r2(row.amounttot);

    xml += `\n${startVoucher}
<PARTYNAME>${partyName}</PARTYNAME>
<ADDRESS.LIST TYPE="String">
<ADDRESS>${address}</ADDRESS>
<ADDRESS>${place}</ADDRESS>
</ADDRESS.LIST>
<PARTYGSTIN>${partyGstin}</PARTYGSTIN>
<PARTYLEDGERNAME>${partyName}</PARTYLEDGERNAME>
<PARTYMAILINGNAME>${partyName}</PARTYMAILINGNAME>
<PARTYPINCODE>${pin}</PARTYPINCODE>`;

    if (dispatchEnabled) {
      xml += `
<DISPATCHFROMADDRESS.LIST TYPE="String">
<DISPATCHFROMADDRESS>${xe(d_add)}</DISPATCHFROMADDRESS>
<DISPATCHFROMADDRESS>${xe(d_place)}</DISPATCHFROMADDRESS>
</DISPATCHFROMADDRESS.LIST>
<DISPATCHFROMNAME>${xe(d_company)}</DISPATCHFROMNAME>
<DISPATCHFROMSTATENAME>${xe(d_state)}</DISPATCHFROMSTATENAME>
<DISPATCHFROMPINCODE>${xe(d_pin)}</DISPATCHFROMPINCODE>
<DISPATCHFROMPLACE>${xe(d_place)}</DISPATCHFROMPLACE>`;
    }

    xml += `
<DATE>${dateval}</DATE>
<REFERENCEDATE></REFERENCEDATE>
<VCHSTATUSDATE>${dateval}</VCHSTATUSDATE>
<GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>
<STATENAME>${state}</STATENAME>
<COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
<PLACEOFSUPPLY>${state}</PLACEOFSUPPLY>
<VOUCHERNUMBER>${xe(voucherNum)}</VOUCHERNUMBER>
<REFERENCE></REFERENCE>
<CONSIGNEEGSTIN>${partyGstin}</CONSIGNEEGSTIN>
<CONSIGNEEMAILINGNAME>${partyName}</CONSIGNEEMAILINGNAME>
<CONSIGNEEPINCODE>${pin}</CONSIGNEEPINCODE>
<CONSIGNEESTATENAME>${state}</CONSIGNEESTATENAME>
<CONSIGNEECOUNTRYNAME>India</CONSIGNEECOUNTRYNAME>
<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
<VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>
<VCHENTRYMODE>Item Invoice</VCHENTRYMODE>
<DIFFACTUALQTY>Yes</DIFFACTUALQTY>
<EFFECTIVEDATE>${dateval}</EFFECTIVEDATE>
<ISINVOICE>Yes</ISINVOICE>
<NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>

<LEDGERENTRIES.LIST>
<LEDGERNAME>${partyName}</LEDGERNAME>
<ISPARTYLEDGER>Yes</ISPARTYLEDGER>
${TAGS.DEEMYES}
<AMOUNT>${-totalRound}</AMOUNT>
<BILLALLOCATIONS.LIST>
<NAME>${xe(voucherNum)}</NAME>
<BILLTYPE>New Ref</BILLTYPE>
<AMOUNT>${-totalRound}</AMOUNT>
</BILLALLOCATIONS.LIST>
</LEDGERENTRIES.LIST>`;

    // Tax ledgers
    if (amazing || !isIntra) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_IGST)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(row.igst || 0)}</AMOUNT>
${rates.igst}
${rates.cess}
</LEDGERENTRIES.LIST>`;
    } else {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_CGST)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(row.cgst || 0)}</AMOUNT>
${rates.cgst}
${rates.cess}
</LEDGERENTRIES.LIST>
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_SGST)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(row.sgst || 0)}</AMOUNT>
${rates.sgst}
${rates.cess}
</LEDGERENTRIES.LIST>`;
    }

    // TCS
    if (tcs && row.tcsamt && row.tcsamt > 0) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_TCS)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(row.tcsamt)}</AMOUNT>
</LEDGERENTRIES.LIST>`;
    }

    // Round off — always emit when enabled (even with a 0 amount).
    if (tlyrnd) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Round_LDR)}</LEDGERNAME>
${rnd < 0 ? TAGS.DEEMYES : TAGS.DEEMNO}
<AMOUNT>${r2(-rnd)}</AMOUNT>
</LEDGERENTRIES.LIST>`;
    }

    xml += invEntries;
    xml += `\n${TAGS.ENDVOUCHER}`;
  }

  xml += '\n' + endEnvelope();
  const BOM = '\uFEFF';
  return BOM + xml.replace(/\r?\n/g, '\r\n');
}

// =====================================================================
// 2. RD PURCHASE (registered dealer purchases — VBA generRD)
// =====================================================================
//
// rows = [{
//   ano, date, name, address, place, pin, gstin (full "GSTIN.xxxx" or bare),
//   pan, lots: [{lot, bag, qty, rate, amount, bilamt}, ...],
//   amounttot, qtytot, bilamttot, cgst, sgst, igst, tdsamt,
//   total, totalRounded, voucherNum,
// }, ...]
//
function generRDPurchaseXML(rows, cfg, opts = {}) {
  const company    = opts.companyName || cfgGet(cfg, 'tally_asp_company_name', cfgGet(cfg, 'tally_company_name', 'Amazing Spice Park Private Limited'));
  const season     = opts.season || cfgGet(cfg, 'tally_season', cfgGet(cfg, 'season_code', '2026-27'));
  // Short season ("26-27") for the dedicated TDS bill-allocation ref —
  // separate from the full season above ("2026-27") used in goods/GST
  // refs so it lines up with the abbreviation already used in invoice
  // numbering elsewhere in the app.
  const seasonShort= opts.seasonShort || cfgGet(cfg, 'season_short', '26-27');
  const detailed   = cfgBool(cfg, 'tally_detailed_purchase', false);
  const tlyrnd     = cfgBool(cfg, 'tally_round_enabled', true);
  const opt        = cfgBool(cfg, 'tally_optional', false);
  // RD vouchers go to the ASP company. The intra/inter test must therefore
  // compare the seller's GSTIN state against ASP's home state code (32 by
  // default for Kerala), not the ISP one. Falls back to ISP if ASP not set.
  const intra      = cfgGet(cfg, 'tally_state_code_amazing', '') || cfgGet(cfg, 'tally_state_code', '32');
  const ainvPrefix = cfgGet(cfg, 'tally_ainv_prefix', 'ASP/');
  const sStateName = cfgGet(cfg, 'tally_home_state_amazing', cfgGet(cfg, 'tally_home_state', 'Kerala'));

  const Purchase_LDR    = cfgGet(cfg, 'tally_purchase_dealer', 'Trade Purchase From Dealer');
  const Tax_CGST_IN     = cfgGet(cfg, 'tally_cgst_input', 'INPUT CGST 2.5%');
  const Tax_SGST_IN     = cfgGet(cfg, 'tally_sgst_input', 'INPUT SGST 2.5%');
  const Tax_IGST_IN     = cfgGet(cfg, 'tally_igst_input', 'INPUT IGST 5%');
  const TDS_LDR         = cfgGet(cfg, 'tally_tds_ledger', 'TDS on Purchase of Goods');
  const Round_LDR       = cfgGet(cfg, 'tally_round', 'Round On/Off');
  const Item_Card       = cfgGet(cfg, 'tally_item_cardamom', 'Cardamom');
  const HSN_Card        = cfgGet(cfg, 'tally_hsn_cardamom',  '09083110');
  const TDS_Rate        = cfgNum(cfg, 'tds_purchase_rate', 0.1);  // % rate for 194Q

  let xml = '\n' + startEnvelope(company, 'Vouchers');

  for (const row of rows) {
    const dateval    = toTallyDate(row.date);
    const ano        = xe(row.ano);
    const taxNm      = xe(row.voucherNum || row.invo || row.id || '');
    const name       = xe(row.name);
    const address    = xe(row.address);
    const place      = xe(row.place);
    const pin        = xe(row.pin);
    const fullGstin  = String(row.gstin || '');
    const partyGstin = fullGstin.toUpperCase().startsWith('GST') ? fullGstin.slice(6, 21) : fullGstin;
    const state      = xe(findState(partyGstin));
    const isIntra    = String(partyGstin).slice(0, 2) === String(intra);
    const rates      = rateDetails(cfgNum(cfg, 'gst_goods', 5));
    // Voucher number / reference = <tradeno>/<purchase-inv-no>/<season>.
    // The purchase-inv-no comes from `purchases.invo` (the dealer's
    // own invoice number entered when the purchase was recorded). Per-lot
    // bill allocations stay in their own <ano>/<lot>/<season> format.
    const voucherRef = `${row.ano}/${taxNm}/${season}`;

    const startVoucher = `<VOUCHER VCHTYPE="Purchase" ACTION="Create" OBJVIEW="Invoice Voucher View">`;
    const total       = r2(row.total);
    // Prefer the builder-supplied rounded total (matches purchases.total
    // exactly so Tally's audit trail reconciles). Falls back to a fresh
    // round when the row doesn't carry it.
    const totalRound  = tlyrnd ? r0(row.totalRounded != null ? row.totalRounded : total) : total;
    const rnd         = tlyrnd ? r2(totalRound - total) : 0;
    const cgst        = r2(row.cgst || 0);
    const sgst        = r2(row.sgst || 0);
    const igst        = r2(row.igst || 0);
    // TDS amount is always passed through if computed by the builder; the
    // TDS LEDGERENTRIES block is always emitted (per-conversation decision)
    // so even a zero TDS amount shows up as a placeholder ledger entry.
    const tdsamt      = r2(row.tdsamt || 0);
    const bilamttot   = r2(row.bilamttot || total);
    const amounttot   = r2(row.amounttot || 0);
    const qtytot      = r2(row.qtytot || 0);
    const rt          = r2(row.rate || (qtytot > 0 ? amounttot / qtytot : 0));

    // Bill allocations per lot. When rounding is on, distribute so the
    // per-lot New Refs sum EXACTLY to round0(Σ bilamt) — the last lot
    // absorbs the remainder. Rounding each lot independently (the old
    // r0 per line) let the goods refs drift from the party amount and
    // broke the bill-wise balance Tally checks on import.
    let billAlloc1 = '';
    let billAllocSum = 0;
    if (detailed && Array.isArray(row.lots) && row.lots.length) {
      const lotsArr = row.lots;
      if (tlyrnd) {
        const target = r0(lotsArr.reduce((s, l) => s + Number(l.bilamt || 0), 0));
        let running = 0;
        for (let i = 0; i < lotsArr.length; i++) {
          const lot = lotsArr[i];
          const amt = (i === lotsArr.length - 1) ? (target - running) : r0(lot.bilamt || 0);
          running += amt;
          billAllocSum += amt;
          billAlloc1 += `
<BILLALLOCATIONS.LIST>
<NAME>${xe(`${row.ano}/${lot.lot}/${season}`)}</NAME>
<BILLTYPE>New Ref</BILLTYPE>
<AMOUNT>${amt}</AMOUNT>
</BILLALLOCATIONS.LIST>`;
        }
      } else {
        for (const lot of lotsArr) {
          const amt = r2(lot.bilamt || 0);
          billAllocSum += amt;
          billAlloc1 += `
<BILLALLOCATIONS.LIST>
<NAME>${xe(`${row.ano}/${lot.lot}/${season}`)}</NAME>
<BILLTYPE>New Ref</BILLTYPE>
<AMOUNT>${amt}</AMOUNT>
</BILLALLOCATIONS.LIST>`;
        }
      }
    }

    // Inventory blocks (per lot when detailed)
    let invEntries = '';
    if (detailed && Array.isArray(row.lots)) {
      for (const lot of row.lots) {
        const ledger = isIntra ? `${Purchase_LDR}-Local` : `${Purchase_LDR}-Inter_State`;
        const nature = isIntra ? 'Local Purchase - Taxable' : 'Interstate Purchase - Taxable';
        invEntries += `\n<ALLINVENTORYENTRIES.LIST>
<STOCKITEMNAME>${xe(Item_Card)}</STOCKITEMNAME>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
<GSTSOURCETYPE>Ledger</GSTSOURCETYPE>
<HSNLEDGERSOURCE>${xe(ledger)}</HSNLEDGERSOURCE>
<GSTOVRDNSTOREDNATURE>${nature}</GSTOVRDNSTOREDNATURE>
<GSTOVRDNTYPEOFSUPPLY>Goods</GSTOVRDNTYPEOFSUPPLY>
<GSTHSNNAME>${xe(HSN_Card)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>Cardamom</GSTHSNDESCRIPTION>
<BASICPACKAGEMARKS>${xe(lot.lot || '')}</BASICPACKAGEMARKS>
<BASICNUMPACKAGES>${r0(lot.bag)} Bags</BASICNUMPACKAGES>
${TAGS.DEEMYES}
<RATE>${r2(lot.rate)}/Kgs.</RATE>
<AMOUNT>${-r2(lot.amount)}</AMOUNT>
<ACTUALQTY>${r2(lot.qty)}Kgs.</ACTUALQTY>
<BILLEDQTY>${r2(lot.qty)}Kgs.</BILLEDQTY>
<BATCHALLOCATIONS.LIST>
<GODOWNNAME>Main Location</GODOWNNAME>
<BATCHNAME>${xe(`${row.ano}/${lot.lot}`)}</BATCHNAME>
<DESTINATIONGODOWNNAME>Main Location</DESTINATIONGODOWNNAME>
<AMOUNT>${-r2(lot.amount)}</AMOUNT>
<ACTUALQTY>${r2(lot.qty)}Kgs.</ACTUALQTY>
<BILLEDQTY>${r2(lot.qty)}Kgs.</BILLEDQTY>
</BATCHALLOCATIONS.LIST>
<ACCOUNTINGALLOCATIONS.LIST>
<LEDGERNAME>${xe(ledger)}</LEDGERNAME>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
${TAGS.DEEMYES}
<AMOUNT>${-r2(lot.amount)}</AMOUNT>
</ACCOUNTINGALLOCATIONS.LIST>
${isIntra ? rates.igst : `${rates.cgst}\n${rates.sgst}`}
${rates.cess}
</ALLINVENTORYENTRIES.LIST>`;
      }
    } else {
      // Aggregate single inventory entry
      const ledger = isIntra ? `${Purchase_LDR}-Local` : `${Purchase_LDR}-Inter_State`;
      const nature = isIntra ? 'Local Purchase - Taxable' : 'Interstate Purchase - Taxable';
      invEntries += `\n<ALLINVENTORYENTRIES.LIST>
<STOCKITEMNAME>${xe(Item_Card)}</STOCKITEMNAME>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
<GSTSOURCETYPE>Ledger</GSTSOURCETYPE>
<HSNLEDGERSOURCE>${xe(ledger)}</HSNLEDGERSOURCE>
<GSTOVRDNSTOREDNATURE>${nature}</GSTOVRDNSTOREDNATURE>
<GSTOVRDNTYPEOFSUPPLY>Goods</GSTOVRDNTYPEOFSUPPLY>
<GSTHSNNAME>${xe(HSN_Card)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>Cardamom</GSTHSNDESCRIPTION>
<BASICPACKAGEMARKS></BASICPACKAGEMARKS>
<BASICNUMPACKAGES></BASICNUMPACKAGES>
${TAGS.DEEMYES}
<RATE>${rt}/Kgs.</RATE>
<AMOUNT>${-amounttot}</AMOUNT>
<ACTUALQTY>${qtytot}Kgs.</ACTUALQTY>
<BILLEDQTY>${qtytot}Kgs.</BILLEDQTY>
<ACCOUNTINGALLOCATIONS.LIST>
<LEDGERNAME>${xe(ledger)}</LEDGERNAME>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
${TAGS.DEEMYES}
<AMOUNT>${-amounttot}</AMOUNT>
</ACCOUNTINGALLOCATIONS.LIST>
${isIntra ? rates.igst : `${rates.cgst}\n${rates.sgst}`}
${rates.cess}
</ALLINVENTORYENTRIES.LIST>`;
    }

    // Bill allocations split into "goods" + "GST" + "TDS" line items so
    // Tally can age each separately in receivables/payables. Three cases:
    //   • detailed   → one allocation per lot for goods (billAlloc1)
    //                  + one GST allocation = sum(taxes)
    //                  + one TDS allocation = −tdsamt (when TDS > 0)
    //   • aggregate  → single goods allocation = bilamttot
    //                  + one GST allocation = sum(taxes)
    //                  + one TDS allocation = −tdsamt (when TDS > 0)
    // The total of all allocations = the party ledger AMOUNT. Until now
    // TDS was absorbed into the GST line; we now break it out as its own
    // New Ref ({ano}/TDS/{seasonShort}) so Tally shows the gross GST and
    // the TDS withholding separately. The TDS bill carries a negative
    // amount so the sum still equals the party AMOUNT.
    const gstSum     = cgst + sgst + igst;
    // Goods bill-ref total (the per-lot New Refs, or the single aggregate).
    const goodsAllocSum = (detailed && Array.isArray(row.lots) && row.lots.length)
      ? billAllocSum
      : (tlyrnd ? r0(bilamttot) : r2(bilamttot));
    // GST bill-ref carries the COMBINED rounding so goods + GST sum to the
    // rounded grand total (totalRound), matching the PDF's single round0.
    // Rounding goods and GST separately drifted by ±1 rupee and broke the
    // bill-wise balance on import.
    const gstAllocAmt = tlyrnd ? (totalRound - goodsAllocSum) : r2(gstSum);
    const tdsAllocAmt = tlyrnd ? -r0(tdsamt) : -r2(tdsamt);
    const gstAlloc   = `
<BILLALLOCATIONS.LIST>
<NAME>${xe(`${row.ano}/GST/${season}`)}</NAME>
<BILLTYPE>New Ref</BILLTYPE>
<AMOUNT>${r2(gstAllocAmt)}</AMOUNT>
</BILLALLOCATIONS.LIST>`;
    const tdsAlloc   = tdsamt > 0 ? `
<BILLALLOCATIONS.LIST>
<NAME>${xe(`${row.ano}/TDS/${seasonShort}`)}</NAME>
<BILLTYPE>New Ref</BILLTYPE>
<AMOUNT>${tdsAllocAmt}</AMOUNT>
</BILLALLOCATIONS.LIST>` : '';
    // Party AMOUNT must equal the sum of its bill allocations. In a 194Q
    // purchase voucher the party is owed (grandTotal − TDS): crediting the
    // full grandTotal double-credited by exactly the TDS, so debit≠credit
    // and Tally only reconciled after the operator re-saved each voucher.
    const partyAmt = tlyrnd
      ? (goodsAllocSum + gstAllocAmt + tdsAllocAmt)
      : r2(goodsAllocSum + gstAllocAmt + tdsAllocAmt);

    xml += `\n${startVoucher}
<ADDRESS.LIST TYPE="String">
<ADDRESS>${address}</ADDRESS>
<ADDRESS>${place}</ADDRESS>
</ADDRESS.LIST>
<DATE>${dateval}</DATE>
<REFERENCEDATE></REFERENCEDATE>
<VCHSTATUSDATE>${dateval}</VCHSTATUSDATE>
<GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>
<STATENAME>${state}</STATENAME>
<COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
<PARTYGSTIN>${xe(partyGstin)}</PARTYGSTIN>
<PLACEOFSUPPLY>${xe(sStateName)}</PLACEOFSUPPLY>
<PARTYNAME>${name}</PARTYNAME>
<PARTYLEDGERNAME>${name}</PARTYLEDGERNAME>
<VOUCHERNUMBER>${xe(voucherRef)}</VOUCHERNUMBER>
<REFERENCE>${xe(voucherRef)}</REFERENCE>
<PARTYMAILINGNAME>${name}</PARTYMAILINGNAME>
<PARTYPINCODE>${pin}</PARTYPINCODE>
<BASICBASEPARTYNAME>${name}</BASICBASEPARTYNAME>
<NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
<VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>
<BASICDATETIMEOFINVOICE></BASICDATETIMEOFINVOICE>
<BASICDATETIMEOFREMOVAL></BASICDATETIMEOFREMOVAL>
<VCHENTRYMODE>Item Invoice</VCHENTRYMODE>
<DIFFACTUALQTY>Yes</DIFFACTUALQTY>
<EFFECTIVEDATE>${dateval}</EFFECTIVEDATE>
<ISELIGIBLEFORITC>Yes</ISELIGIBLEFORITC>
<ISINVOICE>Yes</ISINVOICE>
<ISOPTIONAL>${opt ? 'Yes' : 'No'}</ISOPTIONAL>
<ISVATDUTYPAID>Yes</ISVATDUTYPAID>

<LEDGERENTRIES.LIST>
<LEDGERNAME>${name}</LEDGERNAME>
${TAGS.DEEMNO}
<ISPARTYLEDGER>Yes</ISPARTYLEDGER>
<AMOUNT>${partyAmt}</AMOUNT>${detailed ? billAlloc1 : `
<BILLALLOCATIONS.LIST>
<NAME>${xe(`${row.ano}/${taxNm}/${season}`)}</NAME>
<BILLTYPE>New Ref</BILLTYPE>
<AMOUNT>${goodsAllocSum}</AMOUNT>
</BILLALLOCATIONS.LIST>`}${gstAlloc}${tdsAlloc}
</LEDGERENTRIES.LIST>`;

    // Tax ledgers
    if (isIntra) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_CGST_IN)}</LEDGERNAME>
${TAGS.DEEMYES}
<AMOUNT>${-cgst}</AMOUNT>
<VATEXPAMOUNT>${-cgst}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_SGST_IN)}</LEDGERNAME>
${TAGS.DEEMYES}
<AMOUNT>${-sgst}</AMOUNT>
<VATEXPAMOUNT>${-sgst}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    } else {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_IGST_IN)}</LEDGERNAME>
${TAGS.DEEMYES}
<AMOUNT>${-igst}</AMOUNT>
<VATEXPAMOUNT>${-igst}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    }

    // TDS — emit only when the "TDS on Purchase of Goods" flag is ON
    // (cfg.flag_tds_purchase). The block is the standard 194Q structure
    // with 4 subcategories (Income Tax + 3 placeholder cesses); Tally
    // requires all 4 even if only the first carries the actual TDS amount.
    // When the flag is OFF, we skip emission entirely so the LEDGERENTRIES
    // count drops by 1 — matching the reference XMLs that show some
    // vouchers with TDS and some without.
    if (cfgBool(cfg, 'flag_tds_purchase', false)) {
      const expensesLdr = isIntra ? `${Purchase_LDR}-Local` : `${Purchase_LDR}-Inter_State`;
      const assessable  = amounttot;
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(TDS_LDR)}</LEDGERNAME>
${TAGS.DEEMYES}
<AMOUNT>${tlyrnd ? r0(tdsamt) : tdsamt}</AMOUNT>
<VATEXPAMOUNT>${tlyrnd ? r0(tdsamt) : tdsamt}</VATEXPAMOUNT>
<TAXOBJECTALLOCATIONS.LIST>
<CATEGORY>${xe(TDS_LDR)}</CATEGORY>
<TAXTYPE>TDS</TAXTYPE>
<PARTYLEDGER>${name}</PARTYLEDGER>
<EXPENSES>${xe(expensesLdr)}</EXPENSES>
<REFTYPE>New Ref</REFTYPE>
<EXEMPTED>Yes</EXEMPTED>
<SUBCATEGORYALLOCATION.LIST>
<SUBCATEGORY>Income Tax</SUBCATEGORY>
<DUTYLEDGER>${xe(TDS_LDR)}</DUTYLEDGER>
<TAXRATE>${TDS_Rate}</TAXRATE>
<ASSESSABLEAMOUNT>${tlyrnd ? r0(assessable) : assessable}</ASSESSABLEAMOUNT>
<TAX>${tlyrnd ? r0(tdsamt) : tdsamt}</TAX>
</SUBCATEGORYALLOCATION.LIST>
<SUBCATEGORYALLOCATION.LIST>
<SUBCATEGORY>Surcharge</SUBCATEGORY>
</SUBCATEGORYALLOCATION.LIST>
<SUBCATEGORYALLOCATION.LIST>
<SUBCATEGORY>Education Cess</SUBCATEGORY>
</SUBCATEGORYALLOCATION.LIST>
<SUBCATEGORYALLOCATION.LIST>
<SUBCATEGORY>Secondary Education Cess</SUBCATEGORY>
</SUBCATEGORYALLOCATION.LIST>
</TAXOBJECTALLOCATIONS.LIST>
</LEDGERENTRIES.LIST>`;
    }

    // Round Off — always emit (matches target which has it on every
    // voucher). Computed as the exact balancing figure: credits
    // (party + TDS) minus debits (goods + GST). This guarantees the
    // voucher's debit and credit totals are equal on import — it works
    // out to the PDF's round-off when the allocations are consistent.
    const balancingRnd = r2((partyAmt + tdsamt) - amounttot - cgst - sgst - igst);
    xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Round_LDR)}</LEDGERNAME>
${TAGS.DEEMYES}
<AMOUNT>${-balancingRnd}</AMOUNT>
<VATEXPAMOUNT>${-balancingRnd}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;

    xml += invEntries;
    xml += `\n${TAGS.ENDVOUCHER}`;
  }

  xml += '\n' + endEnvelope();
  const BOM = '\uFEFF';
  return BOM + xml.replace(/\r?\n/g, '\r\n');
}

// =====================================================================
// 3. URD PURCHASE (agriculturist / unregistered — VBA generURD)
// =====================================================================
//
// rows = [{
//   ano, date, name, address, place, pin, lots: [...], amounttot,
//   qtytot, bilamttot, total, voucherNum
// }, ...]
//
function generURDPurchaseXML(rows, cfg, opts = {}) {
  const company   = opts.companyName || cfgGet(cfg, 'tally_company_name', cfgGet(cfg, 'short_name', 'Ideal Spices Private Limited'));
  const season    = opts.season || cfgGet(cfg, 'tally_season', cfgGet(cfg, 'season_code', '2026-27'));
  const detailed  = cfgBool(cfg, 'tally_detailed_purchase', false);
  const tlyrnd    = cfgBool(cfg, 'tally_round_enabled', true);
  const opt       = cfgBool(cfg, 'tally_optional', false);
  const amazing   = cfgBool(cfg, 'tally_amazing_mode', false);
  const invPrefix = cfgGet(cfg, 'tally_inv_prefix', 'ISP/');
  const ainvPrefix= cfgGet(cfg, 'tally_ainv_prefix', 'ASP/');
  const sStateName= cfgGet(cfg, 'tally_urd_state', 'Kerala');

  const Auction_LDR    = cfgGet(cfg, 'tally_purchase_auction', 'Auction Purchase Account');
  const Round_LDR      = cfgGet(cfg, 'tally_round', 'Round Off');
  const Item_Card      = cfgGet(cfg, 'tally_item_cardamom', 'Cardamom');
  const HSN_Card       = cfgGet(cfg, 'tally_hsn_cardamom',  '09083110');

  const rates = rateDetails(cfgNum(cfg, 'gst_goods', 5));

  let xml = '\n' + startEnvelope(company, 'Vouchers');

  for (const row of rows) {
    const dateval    = toTallyDate(row.date);
    const ano        = xe(row.ano);
    const taxNm      = xe(row.voucherNum || row.invo || row.id || '');
    const name       = xe(row.name);
    const address    = xe(row.address);
    const place      = xe(row.place);
    const pin        = xe(row.pin);
    const total      = r2(row.total);
    const totalRound = tlyrnd ? r0(total) : total;
    const rnd        = tlyrnd ? r2(totalRound - total) : 0;
    const amounttot  = r2(row.amounttot);
    const qtytot     = r2(row.qtytot);
    const rt         = r2(qtytot > 0 ? amounttot / qtytot : 0);
    // URD purchase vouchers always go to the ASP Tally company, so the
    // voucher number always uses the ASP prefix regardless of amazing mode.
    const voucherRef = `${ainvPrefix}P-${taxNm}/${season}`;

    const startVoucher = `<VOUCHER VCHTYPE="Purchase" ACTION="Create" OBJVIEW="Invoice Voucher View">`;

    // Bill allocations (per lot)
    let billAlloc = '';
    if (Array.isArray(row.lots)) {
      for (const lot of row.lots) {
        billAlloc += `
<BILLALLOCATIONS.LIST>
<NAME>${xe(`${row.ano}/${lot.lot}/${season}`)}</NAME>
<BILLTYPE>New Ref</BILLTYPE>
<AMOUNT>${tlyrnd ? r0(lot.bilamt || lot.amount) : r2(lot.bilamt || lot.amount)}</AMOUNT>
</BILLALLOCATIONS.LIST>`;
      }
    }

    // Inventory: detailed-per-lot or aggregated
    let invEntries = '';
    if (detailed && Array.isArray(row.lots)) {
      for (const lot of row.lots) {
        invEntries += `\n<ALLINVENTORYENTRIES.LIST>
<STOCKITEMNAME>${xe(Item_Card)}</STOCKITEMNAME>
<GSTOVRDNTAXABILITY>Nil Rated</GSTOVRDNTAXABILITY>
<GSTSOURCETYPE>Ledger</GSTSOURCETYPE>
<GSTLEDGERSOURCE>${xe(Auction_LDR)}</GSTLEDGERSOURCE>
<HSNSOURCETYPE>Stock Item</HSNSOURCETYPE>
<HSNITEMSOURCE>${xe(Item_Card)}</HSNITEMSOURCE>
<GSTOVRDNTYPEOFSUPPLY>Goods</GSTOVRDNTYPEOFSUPPLY>
<GSTHSNNAME>${xe(HSN_Card)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>${xe(Item_Card)}</GSTHSNDESCRIPTION>
<BASICPACKAGEMARKS>${xe(lot.lot || '')}</BASICPACKAGEMARKS>
<BASICNUMPACKAGES>${r0(lot.bag)} Bags</BASICNUMPACKAGES>
${TAGS.DEEMYES}
<RATE>${r2(lot.rate)}/Kgs.</RATE>
<AMOUNT>${-r2(lot.amount)}</AMOUNT>
<ACTUALQTY>${r2(lot.qty)}Kgs.</ACTUALQTY>
<BILLEDQTY>${r2(lot.qty)}Kgs.</BILLEDQTY>
<BATCHALLOCATIONS.LIST>
<GODOWNNAME>Main Location</GODOWNNAME>
<BATCHNAME>${xe(`${row.ano}/${lot.lot}`)}</BATCHNAME>
<DESTINATIONGODOWNNAME>Main Location</DESTINATIONGODOWNNAME>
<AMOUNT>${-r2(lot.amount)}</AMOUNT>
<ACTUALQTY>${r2(lot.qty)}Kgs.</ACTUALQTY>
<BILLEDQTY>${r2(lot.qty)}Kgs.</BILLEDQTY>
</BATCHALLOCATIONS.LIST>
<ACCOUNTINGALLOCATIONS.LIST>
<LEDGERNAME>${xe(Auction_LDR)}</LEDGERNAME>
<GSTOVRDNTAXABILITY>Nil Rated</GSTOVRDNTAXABILITY>
${TAGS.DEEMYES}
<AMOUNT>${-r2(lot.amount)}</AMOUNT>
</ACCOUNTINGALLOCATIONS.LIST>
${rates.cgst}
${rates.sgst}
${rates.igst}
${rates.cess}
${rates.scess}
</ALLINVENTORYENTRIES.LIST>`;
      }
    } else {
      invEntries = `\n<ALLINVENTORYENTRIES.LIST>
<STOCKITEMNAME>${xe(Item_Card)}</STOCKITEMNAME>
<GSTOVRDNTAXABILITY>Nil Rated</GSTOVRDNTAXABILITY>
<GSTSOURCETYPE>Ledger</GSTSOURCETYPE>
<GSTLEDGERSOURCE>${xe(Auction_LDR)}</GSTLEDGERSOURCE>
<HSNSOURCETYPE>Stock Item</HSNSOURCETYPE>
<HSNITEMSOURCE>${xe(Item_Card)}</HSNITEMSOURCE>
<GSTOVRDNTYPEOFSUPPLY>Goods</GSTOVRDNTYPEOFSUPPLY>
<GSTHSNNAME>${xe(HSN_Card)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>${xe(Item_Card)}</GSTHSNDESCRIPTION>
<BASICPACKAGEMARKS></BASICPACKAGEMARKS>
<BASICNUMPACKAGES></BASICNUMPACKAGES>
${TAGS.DEEMYES}
<RATE>${rt}/Kgs.</RATE>
<AMOUNT>${-amounttot}</AMOUNT>
<ACTUALQTY>${qtytot}Kgs.</ACTUALQTY>
<BILLEDQTY>${qtytot}Kgs.</BILLEDQTY>
<ACCOUNTINGALLOCATIONS.LIST>
<LEDGERNAME>${xe(Auction_LDR)}</LEDGERNAME>
<GSTOVRDNTAXABILITY>Nil Rated</GSTOVRDNTAXABILITY>
${TAGS.DEEMYES}
<AMOUNT>${-amounttot}</AMOUNT>
</ACCOUNTINGALLOCATIONS.LIST>
${rates.cgst}
${rates.sgst}
${rates.igst}
${rates.cess}
${rates.scess}
</ALLINVENTORYENTRIES.LIST>`;
    }

    xml += `\n${startVoucher}
<ADDRESS.LIST TYPE="String">
<ADDRESS>${address}</ADDRESS>
<ADDRESS>${place}</ADDRESS>
</ADDRESS.LIST>
<DATE>${dateval}</DATE>
<REFERENCEDATE></REFERENCEDATE>
<VCHSTATUSDATE>${dateval}</VCHSTATUSDATE>
<GSTREGISTRATIONTYPE>Unregistered/Consumer</GSTREGISTRATIONTYPE>
<VATDEALERTYPE>Unregistered/Consumer</VATDEALERTYPE>
<STATENAME>${xe(sStateName)}</STATENAME>
<COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
<PLACEOFSUPPLY>${xe(sStateName)}</PLACEOFSUPPLY>
<PARTYNAME>${name}</PARTYNAME>
<REFERENCE></REFERENCE>
<PARTYLEDGERNAME>${name}</PARTYLEDGERNAME>
<VOUCHERNUMBER>${xe(voucherRef)}</VOUCHERNUMBER>
<PARTYMAILINGNAME>${name}</PARTYMAILINGNAME>
<PARTYPINCODE>${pin}</PARTYPINCODE>
<BASICBASEPARTYNAME>${name}</BASICBASEPARTYNAME>
<NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
<FBTPAYMENTTYPE>Default</FBTPAYMENTTYPE>
<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
<VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>
<BASICDATETIMEOFINVOICE></BASICDATETIMEOFINVOICE>
<BASICDATETIMEOFREMOVAL></BASICDATETIMEOFREMOVAL>
<VCHENTRYMODE>Item Invoice</VCHENTRYMODE>
<DIFFACTUALQTY>Yes</DIFFACTUALQTY>
<EFFECTIVEDATE>${dateval}</EFFECTIVEDATE>
<ISELIGIBLEFORITC>Yes</ISELIGIBLEFORITC>
<VCHGSTSTATUSISUNCERTAIN>No</VCHGSTSTATUSISUNCERTAIN>
<VCHGSTSTATUSISAPPLICABLE>Applicable</VCHGSTSTATUSISAPPLICABLE>
<ISINVOICE>Yes</ISINVOICE>
<ISOPTIONAL>${opt ? 'Yes' : 'No'}</ISOPTIONAL>
<ISVATDUTYPAID>Yes</ISVATDUTYPAID>

<LEDGERENTRIES.LIST>
<LEDGERNAME>${name}</LEDGERNAME>
${TAGS.DEEMNO}
<ISPARTYLEDGER>Yes</ISPARTYLEDGER>
<AMOUNT>${tlyrnd ? r0(total) : total}</AMOUNT>${billAlloc}
</LEDGERENTRIES.LIST>`;

    // Always emit the Round Off line when enabled (even with a 0 amount),
    // matching the sales generators.
    if (tlyrnd) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Round_LDR)}</LEDGERNAME>
${TAGS.DEEMYES}
<AMOUNT>${-r2(rnd)}</AMOUNT>
<VATEXPAMOUNT>${-r2(rnd)}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    }

    xml += invEntries;
    xml += `\n${TAGS.ENDVOUCHER}`;
  }

  xml += '\n' + endEnvelope();

  // Match the byte sequence Tally has been importing successfully from
  // the macros: UTF-8 BOM + CRLF line endings.
  const BOM = '\uFEFF';
  return BOM + xml.replace(/\r?\n/g, '\r\n');
}

// =====================================================================
// 4. DEBIT NOTE (discount received from supplier — VBA generDN)
// =====================================================================
//
// rows = [{
//   ano, date, name (with -PURCHASE suffix or we add it),
//   address, place, pin, gstin, partyGstin,
//   refundtot, cgsttot, sgsttot, igsttot, total, voucherNum
// }, ...]
//
function generDebitNoteXML(rows, cfg, opts = {}) {
  const company    = opts.companyName || cfgGet(cfg, 'tally_company_name', cfgGet(cfg, 'short_name', 'Ideal Spices Private Limited'));
  const season     = opts.season || cfgGet(cfg, 'tally_season', cfgGet(cfg, 'season_code', '2026-27'));
  const tlyrnd     = cfgBool(cfg, 'tally_round_enabled', true);
  const exempt     = cfgBool(cfg, 'tally_dn_exempt', false);
  const opt        = cfgBool(cfg, 'tally_optional', false);
  const intra      = cfgGet(cfg, 'tally_state_code', '33');
  const sStateName = cfgGet(cfg, 'tally_home_state', 'Tamil Nadu');

  const Discount_LDR  = cfgGet(cfg, 'tally_dn_discount', 'Discount Received');
  const Tax_CGST      = cfgGet(cfg, 'tally_dn_cgst', 'OUTPUT CGST 9%');
  const Tax_SGST      = cfgGet(cfg, 'tally_dn_sgst', 'OUTPUT SGST 9%');
  const Tax_IGST      = cfgGet(cfg, 'tally_dn_igst', 'OUTPUT IGST 18%');
  const Round_LDR     = cfgGet(cfg, 'tally_round', 'Round Off');
  const HSN_Service   = cfgGet(cfg, 'tally_hsn_service', '996111');
  const dnGstRate     = cfgNum(cfg, 'tally_dn_gst_rate', 18);

  const rates = rateDetails(dnGstRate);

  let xml = '\n' + startEnvelope(company, 'Vouchers');

  for (const row of rows) {
    const dateval     = toTallyDate(row.date);
    const ano         = xe(row.ano);
    const taxNm       = xe(row.voucherNum || row.note_no || row.id || '');
    const name        = xe(row.name);
    const address     = xe(row.address);
    const place       = xe(row.place);
    const pin         = xe(row.pin);
    const fullGstin   = String(row.gstin || '');
    const partyGstin  = row.partyGstin || (fullGstin.toUpperCase().startsWith('GST') ? fullGstin.slice(6, 21) : fullGstin);
    const state       = xe(findState(partyGstin));
    // Registered dealers carry a PARTYGSTIN, so its 2-digit state code is
    // authoritative. URD / planter debit notes have NO GSTIN — for those we
    // fall back to the GST split stored on the row (CGST/SGST present ⇒
    // intra, else inter). The fallback only fires when partyGstin is empty,
    // so registered-dealer behaviour is unchanged.
    const isIntra     = partyGstin
      ? (String(partyGstin).slice(0, 2) === String(intra))
      : (Number(row.cgsttot || row.cgst || 0) > 0 || Number(row.sgsttot || row.sgst || 0) > 0);
    const refundtot   = r2(row.refundtot || row.amount || 0);
    const cgsttot     = r2(row.cgsttot || row.cgst || 0);
    const sgsttot     = r2(row.sgsttot || row.sgst || 0);
    const igsttot     = r2(row.igsttot || row.igst || 0);
    const total       = r2(row.total || (refundtot + cgsttot + sgsttot + igsttot));
    const totalRound  = tlyrnd ? r0(total) : total;
    const rnd         = tlyrnd ? r2(totalRound - total) : 0;
    // Planter (URD) debit notes post their commission/handling discount to
    // their own ledger (tally_dnp_discount); registered-dealer DNs use the
    // regular DN discount ledger. Falls back to the dealer ledger when the
    // planter ledger isn't separately configured.
    const discountLdr = row.planter
      ? cfgGet(cfg, 'tally_dnp_discount', Discount_LDR)
      : Discount_LDR;

    const startVoucher = `<VOUCHER VCHTYPE="Debit Note" ACTION="Create" OBJVIEW="Invoice Voucher View">`;

    xml += `\n${startVoucher}
<ADDRESS.LIST TYPE="String">
<ADDRESS>${address}</ADDRESS>
<ADDRESS>${place}</ADDRESS>
</ADDRESS.LIST>
<DATE>${dateval}</DATE>
<REFERENCEDATE>${dateval}</REFERENCEDATE>
<VCHSTATUSDATE>${dateval}</VCHSTATUSDATE>
<GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>
<STATENAME>${state}</STATENAME>
<COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
<PARTYGSTIN>${xe(partyGstin)}</PARTYGSTIN>
<PLACEOFSUPPLY>${xe(sStateName)}</PLACEOFSUPPLY>
<PARTYNAME>${name}</PARTYNAME>
<PARTYLEDGERNAME>${name}</PARTYLEDGERNAME>
<VOUCHERNUMBER>${xe(`DN/${taxNm}/${season}`)}</VOUCHERNUMBER>
<REFERENCE>${xe(`DN/${taxNm}/${season}`)}</REFERENCE>
<PARTYMAILINGNAME>${name}</PARTYMAILINGNAME>
<PARTYPINCODE>${pin}</PARTYPINCODE>
<NUMBERINGSTYLE>Manual</NUMBERINGSTYLE>
<PERSISTEDVIEW>Invoice Voucher View</PERSISTEDVIEW>
<VOUCHERTYPENAME>Debit Note</VOUCHERTYPENAME>
<VCHENTRYMODE>Item Invoice</VCHENTRYMODE>
<EFFECTIVEDATE>${dateval}</EFFECTIVEDATE>
<ISINVOICE>Yes</ISINVOICE>
<ISOPTIONAL>${opt ? 'Yes' : 'No'}</ISOPTIONAL>

<LEDGERENTRIES.LIST>
<LEDGERNAME>${name}</LEDGERNAME>
${TAGS.DEEMNO}
<ISPARTYLEDGER>Yes</ISPARTYLEDGER>
<AMOUNT>${-totalRound}</AMOUNT>
<BILLALLOCATIONS.LIST>
<NAME>${xe(`DN/${taxNm}/${season}`)}</NAME>
<BILLTYPE>New Ref</BILLTYPE>
<AMOUNT>${-totalRound}</AMOUNT>
</BILLALLOCATIONS.LIST>
</LEDGERENTRIES.LIST>

<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(discountLdr)}</LEDGERNAME>
<GSTOVRDNTAXABILITY>Taxable</GSTOVRDNTAXABILITY>
<HSNSOURCETYPE>Ledger</HSNSOURCETYPE>
<HSNLEDGERSOURCE>${xe(discountLdr)}</HSNLEDGERSOURCE>
<GSTOVRDNTYPEOFSUPPLY>Services</GSTOVRDNTYPEOFSUPPLY>
<GSTHSNNAME>${xe(HSN_Service)}</GSTHSNNAME>
<GSTHSNDESCRIPTION>${xe(discountLdr)}</GSTHSNDESCRIPTION>
${TAGS.DEEMNO}
<AMOUNT>${refundtot}</AMOUNT>
<VATEXPAMOUNT>${refundtot}</VATEXPAMOUNT>
${rates.cgst}
${rates.sgst}
${rates.igst}
${rates.cess}
${rates.scess}
</LEDGERENTRIES.LIST>`;

    if (!exempt) {
      if (isIntra) {
        xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_CGST)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${cgsttot}</AMOUNT>
<VATEXPAMOUNT>${cgsttot}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_SGST)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${sgsttot}</AMOUNT>
<VATEXPAMOUNT>${sgsttot}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
      } else {
        xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Tax_IGST)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${igsttot}</AMOUNT>
<VATEXPAMOUNT>${igsttot}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
      }
    }

    // Always emit the Round Off line when enabled (even with a 0 amount),
    // matching the sales generators.
    if (tlyrnd) {
      xml += `
<LEDGERENTRIES.LIST>
<LEDGERNAME>${xe(Round_LDR)}</LEDGERNAME>
${TAGS.DEEMNO}
<AMOUNT>${r2(rnd)}</AMOUNT>
<VATEXPAMOUNT>${r2(rnd)}</VATEXPAMOUNT>
</LEDGERENTRIES.LIST>`;
    }

    xml += `\n${TAGS.ENDVOUCHER}`;
  }

  xml += '\n' + endEnvelope();
  const BOM = '\uFEFF';
  return BOM + xml.replace(/\r?\n/g, '\r\n');
}

// =====================================================================
// Data builders — convert DB rows into the {rows} shape each XML fn wants
// =====================================================================

// ── ISP / ASP invoice classification ──────────────────────────
// `invoices.state` is stamped with the *business context* state at the
// time of invoice generation (TAMIL NADU=ISP, KERALA=ASP). For legacy
// rows that pre-date the state stamping, we fall back to a heuristic:
// if the invoice's `invo` matches `lots.asp_invo` for any lot of that
// buyer/auction → ASP, else ISP.
const ISP_STATE_SQL = `(UPPER(COALESCE(i.state,'')) = 'TAMIL NADU' OR UPPER(COALESCE(i.state,'')) = '')`;
const ASP_STATE_SQL = `UPPER(COALESCE(i.state,'')) = 'KERALA'`;

/**
 * Pull ISP sales (state = TAMIL NADU) for an auction. The `invoices`
 * table stores ONE summary row per buyer per auction (the writer in
 * server.js inserts the buyer aggregate; per-lot detail lives in the
 * `lots` table). So we read each invoice row directly, then fetch the
 * matching lots separately.
 *
 * Output shape (per row): one voucher per (buyer, sale-type, ISP invo).
 * `aspInvo` is the matching ASP invoice number (for BASICORDERREF).
 */
function buildSalesIspRows(db, auctionId, cfg) {
  const stmt = db.prepare(`
    SELECT i.*, b.add1, b.add2, b.pla AS buyer_pla, b.pin AS buyer_pin, b.cpin AS buyer_cpin,
           b.cadd1 AS buyer_cadd1, b.cadd2 AS buyer_cadd2, b.cpla AS buyer_cpla
    FROM invoices i
    LEFT JOIN buyers b ON b.buyer = i.buyer
    WHERE i.auction_id = ? AND ${ISP_STATE_SQL}
    ORDER BY CAST(i.invo AS INTEGER), i.sale, i.id
  `);
  const raw = stmt.all(auctionId);

  // Per-lot details for ISP voucher inventory entries. ISP voucher uses
  // the SALES rate (lots.price / lots.amount), not the planter rate.
  //
  // We deliberately do NOT filter on `lots.invo = invoice.invo`. The
  // `lots.invo` column gets overwritten by whichever side (ISP or ASP)
  // ran most recently, so it can hold either invoice number — meaning
  // a strict equality filter silently drops all lots when the ASP step
  // was last to write. Since each buyer has at most ONE invoice per
  // auction per state, scoping by (auction_id, buyer) and `amount > 0`
  // is enough to pick the right lots.
  const lotsStmt = db.prepare(`
    SELECT lot_no AS lot, bags AS bag, qty, price AS rate, amount, asp_invo
    FROM lots
    WHERE auction_id = ? AND buyer = ? AND amount > 0
    ORDER BY CAST(lot_no AS INTEGER), lot_no
  `);

  // Fallback lookup for the cross-referenced ASP invoice number when the
  // lots weren't stamped with `asp_invo` (older data, or the ASP step
  // didn't write it back). The matching ASP invoice shares the same
  // auction + buyer under the KERALA state, so we resolve it directly —
  // this is what populates <BASICORDERREF> when lots.asp_invo is blank.
  const aspInvoStmt = db.prepare(`
    SELECT invo FROM invoices
    WHERE auction_id = ? AND buyer = ? AND UPPER(COALESCE(state,'')) = 'KERALA'
    ORDER BY CAST(invo AS INTEGER), id
    LIMIT 1
  `);

  // E-way bill DISTANCE resolution. Priority order:
  //   1. invoices.distance_km — manual per-invoice override
  //   2. route_distances[(dispatch, buyer)] — saved route value
  //   3. blank (Tally accepts empty <DISTANCE>)
  // The dispatch PIN comes from cfg with the same fallback chain the
  // header generator uses, so the route lookup keys match what the user
  // sees on the screen. Route table keys are normalised — we always
  // store the lexicographically smaller PIN as `from_pin`, so we have
  // to do the same when looking up.
  const dispatchPin = String(
    cfgGet(cfg, 'tally_dispatch_pin', '') ||
    cfgGet(cfg, 's_pin', '') ||
    '685553'
  ).trim();
  const routeStmt = db.prepare(
    'SELECT km FROM route_distances WHERE from_pin = ? AND to_pin = ?'
  );

  const out = [];
  for (const r of raw) {
    const lotRows = lotsStmt.all(auctionId, r.buyer);
    // Pick the first non-empty asp_invo from the matching lots — used by
    // generSalesIspXML to build BASICORDERREF (the matching ASP voucher).
    // When no lot carries asp_invo (it wasn't stamped), fall back to the
    // matching ASP invoice (same auction + buyer, KERALA) so the
    // cross-reference still appears on the ISP voucher.
    const aspInvoRow = lotRows.find(l => l.asp_invo && String(l.asp_invo).trim());
    let aspInvo = aspInvoRow ? String(aspInvoRow.asp_invo).trim() : '';
    if (!aspInvo) {
      try {
        const aspRow = aspInvoStmt.get(auctionId, r.buyer);
        if (aspRow && aspRow.invo != null && String(aspRow.invo).trim()) {
          aspInvo = String(aspRow.invo).trim();
        }
      } catch (_) { /* invoices table shape guard */ }
    }

    // Total bags/qty for the gunny inventory line. Use lot sums when
    // available; fall back to invoice aggregate.
    const gunnyBags = lotRows.length
      ? lotRows.reduce((s, l) => s + Number(l.bag || 0), 0)
      : Number(r.bag || 0);

    // PRE-round grand total, taken from the ledger components this
    // voucher actually posts (cardamom + gunny + transport + insurance +
    // CGST + SGST + IGST + TCS). Deriving it here — instead of from the
    // stored tot/rund — keeps the Round Off correct even for legacy /
    // imported invoices that saved `tot` already-rounded with `rund = 0`
    // (which made the round-off come out as 0). It also guarantees the
    // voucher balances: party = round(compSum), round-off =
    // round(compSum) − compSum. For freshly-generated invoices compSum
    // equals tot − rund, so the value is unchanged.
    // Prefer the stored round-off when present (reconciles exactly with
    // the invoice PDF); fall back to the component sum for legacy /
    // imported invoices saved with rund = 0.
    const compSum = r2(
      (r.amount || 0) + (r.gunny || 0) + (r.pava_hc || 0) + (r.ins || 0) +
      (r.cgst || 0) + (r.sgst || 0) + (r.igst || 0) + (r.tcs || 0)
    );
    const total = Number(r.rund)
      ? r2((r.tot || 0) - (r.rund || 0))
      : (compSum > 0 ? compSum : r2((r.tot || 0) - (r.rund || 0)));
    const totalRounded = r0(total);

    // Resolve <DISTANCE> per the priority above.
    let distance = '';
    if (r.distance_km != null && r.distance_km !== '') {
      distance = String(r.distance_km);
    } else {
      // Consignee PIN drives the distance: ship-to (cpin) preferred, bill-to
      // (pin) fallback — same rule as the UI + route-save endpoints so the
      // <DISTANCE> matches what the operator saw and saved on screen.
      const consigneePin = String(r.buyer_cpin || '').trim() || String(r.buyer_pin || '').trim();
      if (consigneePin && dispatchPin) {
        const [k1, k2] = dispatchPin < consigneePin
          ? [dispatchPin, consigneePin]
          : [consigneePin, dispatchPin];
        try {
          const hit = routeStmt.get(k1, k2);
          if (hit && hit.km != null) distance = String(hit.km);
        } catch (e) { /* table may not exist on very old DBs */ }
      }
    }

    out.push({
      ano: r.ano,
      date: r.date,
      sale: r.sale,
      invo: r.invo,
      aspInvo,
      buyer: r.buyer,
      partyName: r.buyer1 || r.buyer || '',
      address: [r.add1, r.add2].filter(Boolean).join(', '),
      place: r.place || r.buyer_pla || '',
      pin: r.buyer_pin || '',
      // Ship-to (consignee) address, when the buyer has one configured. Used
      // for the e-way bill CONSIGNEEADDRESS and, when tally_ship_to is on, the
      // voucher BASICBUYERADDRESS; blank falls back to bill-to.
      consigneeAddress: [r.buyer_cadd1, r.buyer_cadd2].filter(Boolean).join(', '),
      consigneePlace: r.buyer_cpla || '',
      // <BASICSHIPPEDBY>: the "Dispatched Through" value the operator entered
      // in the print modal, persisted on the invoice at PDF-download time.
      // Falls back to the config default (same chain the ISP PDF uses) when
      // the invoice has no saved value.
      shippedBy: String(r.disp_through || '').trim()
        || cfgGet(cfg, 'dispatched_through_isp', '')
        || cfgGet(cfg, 'dispatched_through', ''),
      partyGstin: r.gstin || '',
      lots: lotRows.map(l => ({
        lot: l.lot,
        bag: Number(l.bag || 0),
        qty: Number(l.qty || 0),
        rate: Number(l.rate || 0),
        amount: Number(l.amount || 0),
      })),
      distance,
      // Vehicle number for the e-way bill (<VEHICLENUMBER> / <BASICSHIPVESSELNO>).
      // Stored on invoices.lorry_no via the bulk "Set Lorry No" action.
      vehicleNo: r.lorry_no || '',
      // Aggregates straight from the (single) invoice row — already
      // pre-summed by the invoice writer in server.js.
      amounttot: r2(r.amount || 0),
      gunnyBags,
      gunnyAmt: r2(r.gunny || 0),
      transportAmt: r2(r.pava_hc || 0),
      insuranceAmt: r2(r.ins || 0),
      cgst: r2(r.cgst || 0),
      sgst: r2(r.sgst || 0),
      igst: r2(r.igst || 0),
      tcsamt: r2(r.tcs || 0),
      total,
      totalRounded,
    });
  }
  return out;
}

/**
 * Pull ASP sales (state = KERALA) for an auction. Each ASP voucher is
 * an internal ASP→ISP transfer: ASP sells its lot inventory to the
 * sister company (ISP) at the asp-side planter rate. The buyer is
 * always ISP itself, so we read the ISP party fields from cfg
 * (tn_address1, tn_address2, tn_gstin, tn_branch).
 *
 * Lot rates/amounts come from `lots.asp_prate` / `lots.asp_puramt`
 * (the ASP-side planter calc). Bag counts come from the lots table.
 *
 * GST is recomputed on (cardamom + gunny) only — ASP vouchers never
 * carry transport/insurance (matches calculations.js isASP branch).
 */
function buildSalesAspRows(db, auctionId, cfg) {
  const stmt = db.prepare(`
    SELECT i.*
    FROM invoices i
    WHERE i.auction_id = ? AND ${ASP_STATE_SQL}
    ORDER BY CAST(i.invo AS INTEGER), i.sale, i.id
  `);
  const raw = stmt.all(auctionId);

  // ASP-side per-lot data. Prefer the dual-storage asp_* columns; fall
  // back to legacy (puramt/prate/pqty) for rows that pre-date the
  // dual-storage migration.
  //
  // Like the ISP builder, we don't filter on `lots.asp_invo = invoice.invo`
  // — each buyer has at most ONE ASP invoice per auction, so scoping by
  // (auction_id, buyer) plus a positive asp_puramt is sufficient. This
  // avoids surprises when asp_invo was cleared/changed by a regen.
  const lotsStmt = db.prepare(`
    SELECT lot_no AS lot, bags AS bag,
           qty AS asp_qty,
           CASE WHEN asp_puramt > 0 THEN asp_prate  ELSE prate  END AS asp_rate,
           CASE WHEN asp_puramt > 0 THEN asp_puramt ELSE puramt END AS asp_amount
    FROM lots
    WHERE auction_id = ? AND buyer = ?
      AND (asp_puramt > 0 OR puramt > 0)
    ORDER BY CAST(lot_no AS INTEGER), lot_no
  `);

  // ISP party identity — the ASP voucher's customer is always ISP.
  const ispPartyName = cfgGet(cfg, 'tally_company_name', cfgGet(cfg, 'short_name', 'IDEAL SPICES PRIVATE LIMITED'));
  const ispAddr1     = cfgGet(cfg, 'tn_address1', '');
  const ispAddr2     = cfgGet(cfg, 'tn_address2', '');
  const ispBranch    = cfgGet(cfg, 'tn_branch', cfgGet(cfg, 'br1', ''));
  const ispGstin     = cfgGet(cfg, 'tn_gstin', '');
  const pinMatch     = String(ispAddr2).match(/-(\d{6})/);
  const ispPin       = cfgGet(cfg, 'tn_pin', pinMatch ? pinMatch[1] : '');

  const gstRate = cfgNum(cfg, 'tally_gst_rate', 5);
  const gunnyRate = cfgNum(cfg, 'tally_gunny_rate', 165);
  const aspIntraCode = String(cfgGet(cfg, 'tally_state_code_amazing', '32'));

  const out = [];
  for (const r of raw) {
    const lotRows = lotsStmt.all(auctionId, r.buyer);
    const lots = lotRows.map(l => ({
      lot: l.lot,
      bag: Number(l.bag || 0),
      qty: Number(l.asp_qty || 0),
      rate: Number(l.asp_rate || 0),
      amount: Number(l.asp_amount || 0),
    }));
    const cardAmt = lots.reduce((s, l) => s + l.amount, 0);
    const gunnyBags = lots.length
      ? lots.reduce((s, l) => s + l.bag, 0)
      : Number(r.bag || 0);
    const gunnyAmt = r2(gunnyBags * gunnyRate);
    const taxableTotal = r2(r2(cardAmt) + gunnyAmt);

    // ASP→ISP is intra-state when ASP and ISP share the GSTIN state
    // code (rare; default config has ASP=32 Kerala, ISP=33 TN, so it's
    // always inter-state). Recompute fresh because the invoice row's
    // cgst/sgst/igst was computed for the ISP-side rates.
    const isIntra = String(ispGstin || '').slice(0, 2) === aspIntraCode;
    let cgst = 0, sgst = 0, igst = 0;
    if (isIntra) {
      const half = r2(taxableTotal * (gstRate / 2) / 100);
      cgst = half; sgst = half;
    } else {
      igst = r2(taxableTotal * gstRate / 100);
    }
    const total = r2(taxableTotal + cgst + sgst + igst);

    out.push({
      ano: r.ano,
      date: r.date,
      sale: r.sale,
      invo: r.invo,
      buyer: r.buyer,
      buyerName: r.buyer1 || r.buyer || '',  // downstream ISP-side buyer; surfaced for filter UIs
      ispPartyName,
      ispAddress: ispAddr1,
      ispPlace: ispBranch,
      ispPin,
      ispGstin,
      lots,
      amounttot: r2(cardAmt),
      gunnyBags,
      gunnyAmt,
      cgst, sgst, igst,
      total,
      totalRounded: r0(total),
    });
  }
  return out;
}

/**
 * Pull invoices for an auction, group by invoice number, attach lots.
 * Used by Sales export.
 */
function buildSalesRows(db, auctionId, cfg) {
  const stmt = db.prepare(`
    SELECT i.*, b.add1, b.add2, b.pla AS buyer_pla, b.pin AS buyer_pin
    FROM invoices i
    LEFT JOIN buyers b ON b.buyer = i.buyer
    WHERE i.auction_id = ?
    ORDER BY CAST(i.invo AS INTEGER), i.sale, i.id
  `);
  const raw = stmt.all(auctionId);

  // Group by party — one voucher per buyer.
  // Macro behaviour: a single voucher captures all lots for a buyer with a
  // shared invoice header. We mirror that by keying on buyer + sale + gstin
  // so two distinct invoice numbers for the same buyer (rare but possible
  // when a buyer is split across two trades) still merge into one voucher,
  // and the invoice number used is the lowest one for that buyer. If the
  // user wants strict per-invoice-number vouchers they can give each lot a
  // different buyer code.
  const grouped = {};
  for (const r of raw) {
    const partyKey = `${r.buyer}|${r.gstin}|${r.sale || 'L'}`;
    if (!grouped[partyKey]) {
      grouped[partyKey] = {
        ano: r.ano,
        date: r.date,
        sale: r.sale,
        invo: r.invo,                      // first/lowest invoice number wins
        buyer: r.buyer,
        partyName: r.buyer1 || r.buyer || '',
        address: [r.add1, r.add2].filter(Boolean).join(', '),
        place: r.place || r.buyer_pla || '',
        pin: r.buyer_pin || '',
        partyGstin: r.gstin || '',
        lots: [],
        amounttot: 0,
        gunnyAmt: 0,
        cgst: 0, sgst: 0, igst: 0, tcsamt: 0,
        total: 0,
      };
    }
    const g = grouped[partyKey];
    g.lots.push({
      lot: r.lot,
      bag: r.bag,
      qty: r.qty,
      rate: r.price,
      amount: r.amount,
    });
    g.amounttot += Number(r.amount || 0);
    g.gunnyAmt  += Number(r.gunny || 0);
    g.cgst      += Number(r.cgst || 0);
    g.sgst      += Number(r.sgst || 0);
    g.igst      += Number(r.igst || 0);
    g.tcsamt    += Number(r.tcs || 0);
    g.total     += Number(r.tot || 0);
  }

  // round
  const out = Object.values(grouped);
  for (const g of out) {
    g.amounttot = r2(g.amounttot);
    g.gunnyAmt  = r2(g.gunnyAmt);
    g.cgst = r2(g.cgst); g.sgst = r2(g.sgst); g.igst = r2(g.igst);
    g.tcsamt = r2(g.tcsamt);
    g.total = r2(g.total);
    g.totalRounded = r0(g.total);
  }
  return out;
}

/**
 * Pull purchases (registered dealers) for an auction.
 * RD = gstin starts with "GSTIN." marker (matches the macro convention).
 */
function buildRDPurchaseRows(db, auctionId, cfg) {
  // Pull from purchases table (one row per voucher already aggregated).
  // Accept both "GSTIN.<gstin>" (legacy UI) and bare 15-char GSTIN (Excel
  // import) — both forms identify a registered dealer.
  const stmt = db.prepare(`
    SELECT p.*
    FROM purchases p
    WHERE p.auction_id = ?
      AND (UPPER(COALESCE(p.gstin,'')) LIKE 'GSTIN%' OR p.gstin GLOB '[0-9][0-9]*')
    ORDER BY CAST(p.invo AS INTEGER), p.id
  `);
  const rawAll = stmt.all(auctionId);

  // Collapse duplicate purchase vouchers for the same dealer. Repeated
  // invoice generation leaves multiple `purchases` rows per name; since
  // lots are attached BY NAME below, each duplicate would re-count the
  // same lots (and emit a duplicate voucher). Keep the latest (highest
  // id) so every registered dealer yields exactly ONE RD voucher.
  const byNameRD = new Map();
  for (const p of rawAll) {
    const key = String(p.name || '').trim().toUpperCase();
    const prev = byNameRD.get(key);
    if (!prev || (Number(p.id) || 0) > (Number(prev.id) || 0)) byNameRD.set(key, p);
  }
  const raw = Array.from(byNameRD.values()).sort((a, b) => {
    const ai = parseInt(a.invo, 10), bi = parseInt(b.invo, 10);
    if (Number.isFinite(ai) && Number.isFinite(bi) && ai !== bi) return ai - bi;
    return (Number(a.id) || 0) - (Number(b.id) || 0);
  });

  // Per-lot price (P_Rate / PurAmt / P_Qty), GST and totals are taken from
  // the SAME builder the purchase-invoice screen uses — buildPurchaseInvoice
  // with ispView (the ISP planter figures the printed purchase invoice
  // shows). This keeps the RD Tally voucher's prices and amounts identical
  // to the purchase invoice regardless of the active business state (the
  // legacy path read lots.prate/puramt, the active view, which showed ASP
  // numbers when the active company was ASP). The `purchases` row supplies
  // only the voucher metadata (number, date, party, GSTIN).
  const { buildPurchaseInvoice } = require('./calculations');

  return raw.map((p) => {
    const inv = buildPurchaseInvoice(db, auctionId, p.name, cfg, { ispView: true, useStoredTds: true });
    const li  = inv ? inv.lineItems : [];
    const s   = inv ? inv.summary : null;
    const lots = li.map(x => ({
      lot: x.lot, bag: x.bags, qty: x.pqty, rate: x.prate,
      amount: x.puramt, bilamt: x.puramt,
    }));
    const amounttot = s ? r2(s.totalPuramt) : 0;
    const qtytot    = s ? r2(s.totalQty)    : 0;
    const cgst = s ? r2(s.totalCgst) : 0;
    const sgst = s ? r2(s.totalSgst) : 0;
    const igst = s ? r2(s.totalIgst) : 0;
    // PRE-round grand total + rounded total straight from the purchase
    // invoice, so the Round Off (totalRounded − total) matches it too.
    const preRound     = s ? r2(s.totalPuramt + s.totalCgst + s.totalSgst + s.totalIgst) : 0;
    const totalRounded = s ? r0(s.grandTotal) : 0;
    return {
      ano: p.ano,
      date: p.date,
      name: p.name,
      address: p.add_line,
      place: p.place,
      pin: '',
      gstin: p.gstin,
      pan: '',
      lots,
      qtytot,
      amounttot,
      bilamttot: amounttot,
      cgst, sgst, igst,
      // TDS from the purchase invoice (recomputed on the ISP grand total),
      // falling back to the stored purchase row's value.
      tdsamt: (s && s.tdsAmount != null) ? s.tdsAmount : p.tds,
      total: preRound,
      totalRounded,
      voucherNum: p.invo || String(p.id),
    };
  });
}

/**
 * Pull bills of supply (URD/agriculturist) for an auction.
 */
function buildURDPurchaseRows(db, auctionId, cfg) {
  const stmt = db.prepare(`
    SELECT * FROM bills WHERE ano IN (
      SELECT ano FROM auctions WHERE id = ?
    )
    ORDER BY CAST(bil AS INTEGER), id
  `);
  const rawAll = stmt.all(auctionId);

  // Collapse duplicate bills for the same agriculturist. Like the RD
  // side, lots are attached BY NAME below, so a re-generated bill (a
  // second `bills` row with the same name) would re-count the seller's
  // lots. Keep the latest (highest id) → one URD voucher per seller.
  const byNameURD = new Map();
  for (const b of rawAll) {
    const key = String(b.name || '').trim().toUpperCase();
    const prev = byNameURD.get(key);
    if (!prev || (Number(b.id) || 0) > (Number(prev.id) || 0)) byNameURD.set(key, b);
  }
  const raw = Array.from(byNameURD.values()).sort((a, b) => {
    const ai = parseInt(a.bil, 10), bi = parseInt(b.bil, 10);
    if (Number.isFinite(ai) && Number.isFinite(bi) && ai !== bi) return ai - bi;
    return (Number(a.id) || 0) - (Number(b.id) || 0);
  });

  // Read ISP planter values directly from the dedicated columns. These are
  // populated by calculateLot() on every save (regardless of which
  // business_state mode dad is currently in), so the URD voucher always
  // gets the right view without recomputing or flipping settings.
  // Falls back to the legacy active-view columns for any rows that
  // pre-date the dual-storage migration (isp_puramt = 0).
  // The exclusion uses NO_GSTIN_SQL so we accept both prefixed ("GSTIN.<gstin>")
  // and bare 15-char GSTINs as registered dealers (and exclude both from URD).
  const lotsStmt = db.prepare(`
    SELECT lot_no AS lot, bags AS bag,
           CASE WHEN isp_puramt > 0 THEN isp_pqty   ELSE pqty   END AS qty,
           CASE WHEN isp_puramt > 0 THEN isp_prate  ELSE prate  END AS rate,
           CASE WHEN isp_puramt > 0 THEN isp_puramt ELSE puramt END AS amount,
           bilamt
    FROM lots
    WHERE auction_id = ? AND name = ?
      AND (isp_puramt > 0 OR puramt > 0)
      AND ${NO_GSTIN_SQL}
    ORDER BY lot_no
  `);

  return raw.map((b) => {
    const lots = lotsStmt.all(auctionId, b.name).map(l => ({
      lot: l.lot, bag: l.bag,
      qty: r2(l.qty), rate: r2(l.rate), amount: r2(l.amount),
      bilamt: r2(l.bilamt || l.amount),
    }));
    const qtytot = lots.reduce((s, l) => s + Number(l.qty || 0), 0);
    const amounttot = lots.reduce((s, l) => s + Number(l.amount || 0), 0);
    const bilamttot = lots.reduce((s, l) => s + Number(l.bilamt || 0), 0);
    return {
      ano: b.ano,
      date: b.date,
      name: b.name,
      address: b.add_line,
      place: b.pla,
      pin: '',
      lots,
      qtytot: r2(qtytot),
      amounttot: r2(amounttot),
      bilamttot: r2(bilamttot),
      // Voucher total = sum(isp_puramt). Refunds and commissions get
      // emitted as separate ledger entries inside the voucher, so they
      // are NOT subtracted from the voucher total. Matches the macro.
      total: r2(amounttot),
      voucherNum: String(b.bil),
    };
  });
}

/**
 * Pull debit notes for an auction.
 */
function buildDebitNoteRows(db, auctionId, cfg) {
  // Match debit_notes by `ano` (auction number) — the same column the
  // auction itself uses. Legacy debit notes have `date` set to creation
  // timestamp, not auction date, so date-based lookups miss them.
  const a = db.prepare('SELECT ano FROM auctions WHERE id = ?').get(auctionId);
  if (!a) return [];
  const stmt = db.prepare(`
    SELECT * FROM debit_notes WHERE ano = ? ORDER BY CAST(note_no AS INTEGER), id
  `);
  const raw = stmt.all(a.ano);
  return raw.map((d) => {
    // PRE-round total from components (discount + GST) so the Round Off
    // is correct even when the stored debit_notes.total was pre-rounded
    // (which would make the round-off come out as 0). The generator
    // rounds this and posts the delta. Falls back to the stored total.
    const dnPre = r2((Number(d.amount) || 0) + (Number(d.cgst) || 0) + (Number(d.sgst) || 0) + (Number(d.igst) || 0));
    return {
      ano: d.ano,
      date: d.date,
      name: d.name,
      address: '',
      place: '',
      pin: '',
      gstin: '',
      refundtot: d.amount,
      cgsttot: d.cgst, sgsttot: d.sgst, igsttot: d.igst,
      total: dnPre > 0 ? dnPre : r2(d.total || 0),
      voucherNum: d.note_no || String(d.id),
    };
  });
}

// URD/planter party-ledger name convention: `<name>-[<PAN>]` when a PAN is
// known, else `<name>-PURCHASE`. Matches the URD party ledgers emitted for
// agriculturist purchases so the planter debit-note voucher references an
// existing ledger in Tally.
function _urdPurchaseLedgerName(n, pan) {
  const s = String(n || '').trim();
  if (!s) return s;
  const p = String(pan || '').trim();
  return p ? `${s}-[${p}]` : `${s}-PURCHASE`;
}

// PLANTER debit notes — mirrors buildDebitNoteRows but reads the
// `debit_notes_planter` table. Parties are agriculturists / unregistered
// sellers (URD), so the PARTYLEDGERNAME uses the URD ledger-name convention
// and `gstin` is left blank (planters have none) — generDebitNoteXML then
// classifies intra/inter from the stored CGST/SGST/IGST amounts and posts
// the discount to the planter discount ledger. Address / place / PAN come
// from the seller's own bill of supply for the trade.
function buildDebitNotePlanterRows(db, auctionId, cfg) {
  const a = db.prepare('SELECT ano FROM auctions WHERE id = ?').get(auctionId);
  if (!a) return [];
  const raw = db.prepare(`
    SELECT * FROM debit_notes_planter WHERE ano = ? ORDER BY CAST(note_no AS INTEGER), id
  `).all(a.ano);

  // Pull the matching bill-of-supply row (address / place / PAN) for each
  // distinct planter in this DN batch. Bills store single-digit trade
  // numbers space-padded → match on TRIM(ano).
  const planterNames = [...new Set(raw.map(d => String(d.name || '').trim()).filter(Boolean))];
  const planters = {};
  for (const name of planterNames) {
    const b = db.prepare(
      `SELECT name, add_line, pla, st_code, pan
         FROM bills
        WHERE TRIM(ano) = ? AND UPPER(name) = UPPER(?)
        ORDER BY id DESC LIMIT 1`
    ).get(String(a.ano).trim(), name);
    if (b) planters[name] = b;
  }

  return raw.map((d) => {
    const planter = planters[String(d.name || '').trim()] || {};
    const dnPre = r2((Number(d.amount) || 0) + (Number(d.cgst) || 0) + (Number(d.sgst) || 0) + (Number(d.igst) || 0));
    return {
      ano: d.ano,
      date: d.date,
      name: _urdPurchaseLedgerName(d.name, planter.pan),
      address: planter.add_line || '',
      place:   planter.pla || '',
      pin:     '',
      // URD/planter — no GSTIN. generDebitNoteXML reads the GST amounts to
      // pick intra vs inter when the GSTIN is blank.
      gstin: '',
      refundtot: d.amount,
      cgsttot: d.cgst, sgsttot: d.sgst, igsttot: d.igst,
      total: dnPre > 0 ? dnPre : r2(d.total || 0),
      voucherNum: d.note_no || String(d.id),
      planter: true,
    };
  });
}

// =====================================================================
// 5. LEDGER MASTERS (party + tax + sales + purchase ledgers)
// =====================================================================
//
// Mirrors the macro's generLED / generLEDGERD / generLEDGERP, plus emits
// the standard tax / sales / purchase ledgers configured in Settings →
// To Tally so a fresh Tally company can be primed with all the ledgers
// the voucher imports will reference. Without these ledgers in Tally,
// every voucher import will fail with "ledger not found".
//
// rows = [{ kind: 'party'|'tax'|'sales'|'purchase'|'group',
//           name, parent (group), gstin, pan, address, place, pin,
//           state, applicableFrom (yyyymmdd) }]
//
function generLedgerXML(rows, cfg, opts = {}) {
  const company = opts.companyName || cfgGet(cfg, 'tally_company_name', cfgGet(cfg, 'short_name', 'Ideal Spices Private Limited'));
  const today = toTallyDate(new Date());

  // Use "All Masters" reportname for ledger imports.
  // Target XMLs use CRLF line endings + UTF-8 BOM at start of file. We mirror
  // that here so the byte sequence matches what Tally has been importing
  // successfully from the macros. We build with \n then convert at the end.
  let xml = '\n' + startEnvelope(company, 'All Masters');

  for (const r of rows) {
    const name = xe(r.name || '');
    if (!name) continue;
    const parent = xe(r.parent || 'Sundry Debtors');
    const dateval = r.applicableFrom || today;
    // Always derive state from GSTIN if available (title-cased) so Tally
    // accepts the value. Fall back to the row's `state` only if no GSTIN.
    const stateRaw = (r.gstin ? findState(r.gstin) : r.state) || '';
    const state = xe(stateRaw);
    const gstin = xe(r.gstin || '');
    const pan = xe(r.pan || (r.gstin ? String(r.gstin).slice(2, 12) : ''));
    const address = xe(r.address || '');
    const place = xe(r.place || '');
    const pin = xe(r.pin || '');
    const isParty = r.kind === 'party';
    const hasGst = isParty && gstin;

    // Sub-kind hint from the builder so we know whether this party is sales,
    // RD, or URD. Falls back to a heuristic based on parent group name when
    // not provided (keeps backward compat with the all-in-one builder).
    let partyKind = r.partyKind || '';
    if (!partyKind && isParty) {
      const p = String(r.parent || '').toLowerCase();
      if (p.includes('agriculturist'))         partyKind = 'urd';
      else if (p.includes('dealer-purchase'))  partyKind = 'sales'; // sales-side
      else if (p.includes('dealer'))           partyKind = 'rd';    // RD purchase-side
    }
    const isRdParty  = isParty && partyKind === 'rd';
    const isUrdParty = isParty && partyKind === 'urd';

    // ── Build the LEDGER body in the exact order the target XMLs use ──
    // Order: CURRENCYNAME, PRIORSTATENAME, INCOMETAXNUMBER, [VATDEALERTYPE],
    //        PARENT, COUNTRYOFRESIDENCE, LEDGERCOUNTRYISDCODE,
    //        ISBILLWISEON, ASORIGINAL, ISCHEQUEPRINTINGENABLED,
    //        LANGUAGENAME.LIST, LEDGSTREGDETAILS.LIST, LEDMAILINGDETAILS.LIST.
    // The opening tag has just NAME="" — no top-level <NAME> child, no
    // RESERVEDNAME attribute (matches every sample we've seen).
    let block = `\n<LEDGER NAME="${name}">
<CURRENCYNAME>₹</CURRENCYNAME>`;

    if (isParty) {
      // For sales/RD parties we have state + PAN; for URD we still emit the
      // tags but leave them empty (the target URD XML uses empty tags here,
      // which is what Tally expects for unregistered parties).
      if (isUrdParty) {
        block += `
<PRIORSTATENAME></PRIORSTATENAME>
<INCOMETAXNUMBER></INCOMETAXNUMBER>`;
      } else {
        block += `
<PRIORSTATENAME>${state}</PRIORSTATENAME>
<INCOMETAXNUMBER>${pan}</INCOMETAXNUMBER>`;
      }
      // VATDEALERTYPE: only RD and URD purchase parties carry this (matches
      // the schema in RD_LEDGER_MASTER and URD_LEDGER_MASTER; SALE doesn't).
      if (isRdParty) {
        block += `\n<VATDEALERTYPE>Regular</VATDEALERTYPE>`;
      } else if (isUrdParty) {
        block += `\n<VATDEALERTYPE>Unregistered/Consumer</VATDEALERTYPE>`;
      }
    }

    block += `
<PARENT>${parent}</PARENT>
<COUNTRYOFRESIDENCE>India</COUNTRYOFRESIDENCE>
<LEDGERCOUNTRYISDCODE>+91</LEDGERCOUNTRYISDCODE>`;

    if (isParty) {
      block += `
<ISBILLWISEON>Yes</ISBILLWISEON>
<ASORIGINAL>Yes</ASORIGINAL>
<ISCHEQUEPRINTINGENABLED>Yes</ISCHEQUEPRINTINGENABLED>`;
    }

    block += `
<LANGUAGENAME.LIST>
<NAME.LIST>
<NAME>${name}</NAME>
</NAME.LIST>
</LANGUAGENAME.LIST>`;

    // GST registration block.
    //   • Sales / RD party with GSTIN → full block with GSTIN
    //   • URD party (no GSTIN) → block without GSTIN, GSTREGISTRATIONTYPE = Unregistered/Consumer
    //   • Master ledgers (sales/purchase/tax) → no block
    if (isParty) {
      if (hasGst) {
        block += `
<LEDGSTREGDETAILS.LIST>
<APPLICABLEFROM>${dateval}</APPLICABLEFROM>
<GSTREGISTRATIONTYPE>Regular</GSTREGISTRATIONTYPE>
<PLACEOFSUPPLY>${state}</PLACEOFSUPPLY>
<GSTIN>${gstin}</GSTIN>
</LEDGSTREGDETAILS.LIST>`;
      } else if (isUrdParty) {
        // URD parties still get the block, just without GSTIN
        const placeOfSupply = state || cfgGet(cfg, 'tally_urd_state', 'Kerala');
        block += `
<LEDGSTREGDETAILS.LIST>
<APPLICABLEFROM>${dateval}</APPLICABLEFROM>
<GSTREGISTRATIONTYPE>Unregistered/Consumer</GSTREGISTRATIONTYPE>
<PLACEOFSUPPLY>${xe(placeOfSupply)}</PLACEOFSUPPLY>
</LEDGSTREGDETAILS.LIST>`;
      }
    }

    if (isParty && (address || place || pin)) {
      // For URD where state is unknown, fall back to the configured URD state
      const mailingState = state || (isUrdParty ? xe(cfgGet(cfg, 'tally_urd_state', 'Kerala')) : '');
      block += `
<LEDMAILINGDETAILS.LIST>
<ADDRESS.LIST>
<ADDRESS>${address}</ADDRESS>
<ADDRESS>${place}</ADDRESS>
</ADDRESS.LIST>
<APPLICABLEFROM>${dateval}</APPLICABLEFROM>
<PINCODE>${pin}</PINCODE>
<MAILINGNAME>${name}</MAILINGNAME>
<STATE>${mailingState}</STATE>
<COUNTRY>India</COUNTRY>
</LEDMAILINGDETAILS.LIST>`;
    }

    // Tax ledgers: tag with the right rate-of-tax info (only used by the
    // all-in-one ledger export; the 3 party-only exports never include
    // these).
    if (r.kind === 'tax') {
      const rateOfTax = Number(r.rateOfTax || 0);
      const dutyHead = r.dutyHead || 'GST';
      block += `
<TAXTYPE>GST</TAXTYPE>
<GSTDUTYHEAD>${xe(dutyHead)}</GSTDUTYHEAD>
<RATEOFTAXCALCULATION>${rateOfTax}</RATEOFTAXCALCULATION>`;
    }

    // Sales/Purchase master ledgers: rounding + taxability + HSN
    if (r.kind === 'sales' || r.kind === 'purchase') {
      block += `
<ROUNDINGMETHOD>Normal Rounding</ROUNDINGMETHOD>
<ROUNDINGLIMIT>1</ROUNDINGLIMIT>
<GSTOVRDNTAXABILITY>${xe(r.taxability || 'Taxable')}</GSTOVRDNTAXABILITY>`;
      if (r.hsn) {
        block += `
<HSNCODE>${xe(r.hsn)}</HSNCODE>`;
      }
    }

    block += '\n</LEDGER>';
    xml += block;
  }

  xml += '\n' + endEnvelope();

  // Convert to CRLF line endings + prepend UTF-8 BOM so the byte sequence
  // matches what Tally has been importing successfully from the macros.
  const BOM = '\uFEFF';
  return BOM + xml.replace(/\r?\n/g, '\r\n');
}

// ── Helpers shared by the 3 ledger builders ─────────────────
// Each builder returns rows in the same { kind, name, parent, gstin, ... }
// shape that generLedgerXML already consumes — only the source-of-data
// and the parent group differ.

function _buyerRow(b, todayDate, intra, interDealer, localDealer) {
  const isIntra = String(b.gstin || '').slice(0, 2) === String(intra);
  return {
    kind: 'party',
    partyKind: 'sales',
    name: b.buyer1 || b.buyer || '',
    parent: isIntra ? localDealer : interDealer,
    gstin: b.gstin || '',
    pan: b.pan || '',
    address: [b.add1, b.add2].filter(Boolean).join(', '),
    place: b.pla || '',
    pin: b.pin || '',
    state: b.state || '',
    applicableFrom: todayDate,
  };
}

function _rdTraderRow(t, todayDate, intra, interDealPur, localDealPur) {
  // `cr` carries the GSTIN with a "GSTIN." prefix for registered dealers
  const fullGstin = String(t.cr || '');
  const partyGstin = fullGstin.toUpperCase().startsWith('GST') ? fullGstin.slice(6, 21) : fullGstin;
  const isIntra = String(partyGstin).slice(0, 2) === String(intra);
  return {
    kind: 'party',
    partyKind: 'rd',
    name: t.name || '',
    parent: isIntra ? localDealPur : interDealPur,
    gstin: partyGstin,
    pan: t.pan || '',
    address: t.padd || '',
    place: t.ppla || '',
    pin: t.ppin || '',
    state: t.pstate || '',
    applicableFrom: todayDate,
  };
}

function _urdTraderRow(t, todayDate, auctionLDR) {
  return {
    kind: 'party',
    partyKind: 'urd',
    name: t.name || '',
    parent: auctionLDR,           // Agriculturists go under the auction-purchase parent
    gstin: '',
    pan: t.pan || '',
    address: t.padd || '',
    place: t.ppla || '',
    pin: t.ppin || '',
    state: t.pstate || '',
    applicableFrom: todayDate,
  };
}

/**
 * Build sales-party ledger rows (mirrors generLED).
 * One row per distinct buyer that appears in this auction's invoices.
 * Filter: optional `partyName` to limit output to a single buyer.
 */
function buildSalesPartyLedgerRows(db, auctionId, cfg, opts = {}) {
  const todayDate = toTallyDate(new Date());
  const intra = cfgGet(cfg, 'tally_state_code', '33');
  const interDealer = cfgGet(cfg, 'tally_dealer_sale_inter', 'Interstate Dealer-Purchase');
  const localDealer = cfgGet(cfg, 'tally_dealer_sale_intra', 'Local Dealer-Purchase');

  let sql = `
    SELECT DISTINCT b.*
    FROM invoices i
    JOIN buyers b ON b.buyer = i.buyer
    WHERE i.auction_id = ?
  `;
  const params = [auctionId];
  if (opts.partyName) {
    sql += ` AND (b.buyer1 = ? OR b.buyer = ?)`;
    params.push(opts.partyName, opts.partyName);
  }
  sql += ` ORDER BY b.buyer1`;

  const buyers = db.prepare(sql).all(...params);
  return buyers.map(b => _buyerRow(b, todayDate, intra, interDealer, localDealer));
}

/**
 * Build RD-purchase party ledger rows (mirrors generLEDGERD).
 * One row per distinct trader with a GSTIN (either "GSTIN.<gstin>" or
 * bare 15-char) in the auction's lots.
 */
function buildRDPartyLedgerRows(db, auctionId, cfg, opts = {}) {
  const todayDate = toTallyDate(new Date());
  // RD party ledgers belong in the ASP Tally company. The intra/local
  // determination must therefore use the ASP company's home GSTIN state
  // code (defaults to 32 = Kerala), not the ISP code (33 = Tamil Nadu).
  // Falls back to the ISP code if the ASP one isn't configured, so existing
  // setups without a separate ASP code keep working.
  const intra = cfgGet(cfg, 'tally_state_code_amazing', '') || cfgGet(cfg, 'tally_state_code', '33');
  const interDealPur = cfgGet(cfg, 'tally_purchase_dealer_inter', 'Interstate Dealer');
  const localDealPur = cfgGet(cfg, 'tally_purchase_dealer_intra', 'Local Dealer');

  let sql = `
    SELECT DISTINCT name, padd, ppla, ppin, pstate, cr, pan
    FROM lots
    WHERE auction_id = ? AND ${HAS_GSTIN_SQL}
  `;
  const params = [auctionId];
  if (opts.partyName) {
    sql += ` AND name = ?`;
    params.push(opts.partyName);
  }
  sql += ` ORDER BY name`;

  const traders = db.prepare(sql).all(...params);
  return traders.map(t => _rdTraderRow(t, todayDate, intra, interDealPur, localDealPur));
}

/**
 * Build URD-purchase party ledger rows (mirrors generLEDGERP).
 * One row per distinct agriculturist (no GSTIN) in the auction's lots.
 */
function buildURDPartyLedgerRows(db, auctionId, cfg, opts = {}) {
  const todayDate = toTallyDate(new Date());
  const auctionLDR = cfgGet(cfg, 'tally_purchase_auction', 'Purchase From Agriculturist');

  let sql = `
    SELECT DISTINCT name, padd, ppla, ppin, pstate, pan
    FROM lots
    WHERE auction_id = ? AND ${NO_GSTIN_SQL}
      AND name != ''
  `;
  const params = [auctionId];
  if (opts.partyName) {
    sql += ` AND name = ?`;
    params.push(opts.partyName);
  }
  sql += ` ORDER BY name`;

  const traders = db.prepare(sql).all(...params);
  return traders.map(t => _urdTraderRow(t, todayDate, auctionLDR));
}

/**
 * List every party in an auction with the kind it would be exported as.
 * Powers the "single-party" picker UI on the To Tally page so dad can
 * pick exactly one and emit just that ledger.
 *
 * Returns: [{ kind: 'sales'|'rd_purchase'|'urd_purchase', name, gstin }]
 */
function listAuctionParties(db, auctionId) {
  const out = [];

  // Sales parties — buyers
  const buyers = db.prepare(`
    SELECT DISTINCT b.buyer1 AS name, b.gstin, b.pla AS place
    FROM invoices i
    JOIN buyers b ON b.buyer = i.buyer
    WHERE i.auction_id = ?
    ORDER BY b.buyer1
  `).all(auctionId);
  for (const b of buyers) {
    out.push({ kind: 'sales', name: b.name || '', gstin: b.gstin || '', place: b.place || '' });
  }

  // RD purchase parties — traders with GSTIN in lots (either format)
  const rd = db.prepare(`
    SELECT DISTINCT name, ppla AS place, cr
    FROM lots
    WHERE auction_id = ? AND ${HAS_GSTIN_SQL} AND name != ''
    ORDER BY name
  `).all(auctionId);
  for (const t of rd) {
    // Strip "GSTIN." prefix if present, otherwise the value already IS the
    // bare 15-char GSTIN. Handles both formats produced by UI vs Excel import.
    const raw = String(t.cr || '').trim();
    const gstin = raw.toUpperCase().startsWith('GSTIN.') ? raw.slice(6) : raw;
    out.push({ kind: 'rd_purchase', name: t.name || '', gstin, place: t.place || '' });
  }

  // URD purchase parties — agriculturists (no GSTIN, either form)
  const urd = db.prepare(`
    SELECT DISTINCT name, ppla AS place
    FROM lots
    WHERE auction_id = ? AND ${NO_GSTIN_SQL} AND name != ''
    ORDER BY name
  `).all(auctionId);
  for (const t of urd) {
    out.push({ kind: 'urd_purchase', name: t.name || '', gstin: '', place: t.place || '' });
  }

  return out;
}

/**
 * Backwards-compatible "all ledgers in one shot" builder.
 * Combines the 3 party builders + the master ledgers (sales/purchase/tax/
 * service) seeded from cfg. Useful for a one-click "prime a fresh Tally
 * company" workflow.
 */
function buildLedgerRows(db, auctionId, cfg) {
  const rows = [
    ...buildSalesPartyLedgerRows(db, auctionId, cfg),
    ...buildRDPartyLedgerRows(db, auctionId, cfg),
    ...buildURDPartyLedgerRows(db, auctionId, cfg),
  ];
  const todayDate = toTallyDate(new Date());

  // ── Master ledgers from cfg (sales / purchase / tax / service) ─
  const gstRate = cfgNum(cfg, 'tally_gst_rate', 5);
  const dnRate  = cfgNum(cfg, 'tally_dn_gst_rate', 18);
  const hsnCard = cfgGet(cfg, 'tally_hsn_cardamom', '09083120');
  const hsnService = cfgGet(cfg, 'tally_hsn_service', '996111');

  for (const [k, parent, taxability, hsn] of [
    ['tally_sales_inter',   'Sales Accounts',  'Taxable',   hsnCard],
    ['tally_sales_intra',   'Sales Accounts',  'Taxable',   hsnCard],
    ['tally_sales_export',  'Sales Accounts',  'Exempt',    hsnCard],
    ['tally_gunny_inter',   'Sales Accounts',  'Taxable',   cfgGet(cfg, 'tally_hsn_gunny', '63051040')],
    ['tally_gunny_intra',   'Sales Accounts',  'Taxable',   cfgGet(cfg, 'tally_hsn_gunny', '63051040')],
    ['tally_gunny_export',  'Sales Accounts',  'Exempt',    cfgGet(cfg, 'tally_hsn_gunny', '63051040')],
  ]) {
    const name = cfgGet(cfg, k, '');
    if (name) rows.push({ kind: 'sales', name, parent, taxability, hsn, applicableFrom: todayDate });
  }

  const purBase = cfgGet(cfg, 'tally_purchase_dealer', 'Trade Purchase From Dealer');
  if (purBase) {
    rows.push({ kind: 'purchase', name: `${purBase}-Local`,       parent: 'Purchase Accounts', taxability: 'Taxable', hsn: hsnCard, applicableFrom: todayDate });
    rows.push({ kind: 'purchase', name: `${purBase}-Inter_State`, parent: 'Purchase Accounts', taxability: 'Taxable', hsn: hsnCard, applicableFrom: todayDate });
  }
  const auctionLDRname = cfgGet(cfg, 'tally_purchase_auction', '');
  if (auctionLDRname) rows.push({ kind: 'purchase', name: auctionLDRname, parent: 'Purchase Accounts', taxability: 'Nil Rated', hsn: hsnCard, applicableFrom: todayDate });

  const tax = (key, dutyHead, rate) => {
    const name = cfgGet(cfg, key, '');
    if (name) rows.push({ kind: 'tax', name, parent: 'Duties & Taxes', dutyHead, rateOfTax: rate, applicableFrom: todayDate });
  };
  tax('tally_cgst',        'CGST', gstRate / 2);
  tax('tally_sgst',        'SGST/UTGST', gstRate / 2);
  tax('tally_igst',        'IGST', gstRate);
  tax('tally_cgst_input',  'CGST', gstRate / 2);
  tax('tally_sgst_input',  'SGST/UTGST', gstRate / 2);
  tax('tally_igst_input',  'IGST', gstRate);
  tax('tally_dn_cgst',     'CGST', dnRate / 2);
  tax('tally_dn_sgst',     'SGST/UTGST', dnRate / 2);
  tax('tally_dn_igst',     'IGST', dnRate);
  tax('tally_tcs',         'TCS',  cfgNum(cfg, 'tally_tcs_rate', 0.1));
  tax('tally_tds_ledger',  'TDS',  cfgNum(cfg, 'tally_tcs_rate', 0.1));

  const services = [
    ['tally_dn_discount',          'Indirect Incomes',   hsnService],
    ['tally_commission',           'Indirect Expenses',  hsnService],
    ['tally_cash_handling',        'Indirect Expenses',  hsnService],
    ['tally_cash_handling_planter','Indirect Expenses',  hsnService],
    ['tally_chc_planter',          'Indirect Expenses',  hsnService],
    ['tally_sample_planter',       'Indirect Expenses',  hsnService],
    ['tally_sample_dealer',        'Indirect Expenses',  hsnService],
    ['tally_transport',            'Indirect Expenses',  cfgGet(cfg, 'tally_hsn_transport', '996791')],
    ['tally_insurance',            'Indirect Expenses',  cfgGet(cfg, 'tally_hsn_insurance', '997136')],
    ['tally_round',                'Indirect Expenses',  ''],
    ['tally_tds_paid_sales',       'Duties & Taxes',     ''],
  ];
  for (const [k, parent, hsn] of services) {
    const name = cfgGet(cfg, k, '');
    if (name) rows.push({ kind: 'sales', name, parent, taxability: 'Taxable', hsn, applicableFrom: todayDate });
  }

  return rows;
}

module.exports = {
  generSalesXML,
  generSalesIspXML,
  generSalesAspXML,
  generIspPurchaseXML,
  generRDPurchaseXML,
  generURDPurchaseXML,
  generDebitNoteXML,
  generLedgerXML,
  buildSalesRows,
  buildSalesIspRows,
  buildSalesAspRows,
  buildRDPurchaseRows,
  buildURDPurchaseRows,
  buildDebitNoteRows,
  buildDebitNotePlanterRows,
  buildLedgerRows,
  buildSalesPartyLedgerRows,
  buildRDPartyLedgerRows,
  buildURDPartyLedgerRows,
  listAuctionParties,
  // helpers (exported for tests)
  toTallyDate,
  findState,
};
