/**
 * company-config.js — Replaces TOOL.DBF + DEPOTS.PRG + COMPANY.PRG
 * All company configuration stored as key-value pairs in SQLite.
 */

const DEFAULTS = [
  // ── COMPANY (Primary - ISP) ────────────────────────────────
  { key: 'logo',            value: 'ISP',           category: 'company',   label: 'Logo Code',                type: 'text' },
  { key: 'trade_name',      value: 'IDEAL SPICES',  category: 'company',   label: 'Trade Name',               type: 'text' },
  { key: 'legal_name',      value: ' PRIVATE LIMITED', category: 'company', label: 'Legal Name Suffix',        type: 'text' },
  { key: 'short_name',      value: 'IDEAL SPICES PRIVATE LIMITED', category: 'company', label: 'Short Name', type: 'text' },
  { key: 'pan',             value: 'AAICI5415L',    category: 'company',   label: 'PAN',                      type: 'text' },
  { key: 'cin',             value: 'U47211TN2025PTC186657', category: 'company', label: 'CIN',                type: 'text' },
  // Partnership Firm toggle: when ON, the invoice letterhead's "CIN" line
  // is replaced by "Partnership: <partnership_name>" (a partnership firm
  // has no CIN). In Settings → Company the two fields are mutually
  // exclusive — turning this ON greys out CIN and vice-versa.
  { key: 'is_partnership',  value: 'false',          category: 'company',   label: 'Partnership Firm',        type: 'boolean' },
  { key: 'partnership_name', value: '',              category: 'company',   label: 'Partnership Name / No.',  type: 'text' },
  { key: 'fssai',           value: '',               category: 'company',   label: 'FSSAI No.',               type: 'text' },
  { key: 'sbl',             value: '',               category: 'company',   label: 'SBL No.',                 type: 'text' },

  // ── ADDRESS (Kerala) ───────────────────────────────────────
  { key: 'kl_address1',     value: 'FLAT No.42,V.O.C.1ST STREET,MELACHOKKANATHAPURAM', category: 'address_kl', label: 'Address Line 1', type: 'text' },
  { key: 'kl_address2',     value: 'BODINAYAKANUR, THENI-625582 TAMIL NADU CODE:33 Mobile:8610943865', category: 'address_kl', label: 'Address Line 2', type: 'text' },
  { key: 'kl_phone',        value: '8610943865',    category: 'address_kl', label: 'Phone',                   type: 'text' },
  { key: 'kl_email',        value: 'idealspicesbodi@gmail.com', category: 'address_kl', label: 'Email',       type: 'text' },
  { key: 'kl_gstin',        value: '32AAICI5415L1ZX', category: 'address_kl', label: 'GSTIN',                 type: 'text' },
  { key: 'kl_branch',       value: 'BODINAYAKANUR', category: 'address_kl', label: 'Office Branch',           type: 'text' },

  // ── ADDRESS (Tamil Nadu) ───────────────────────────────────
  { key: 'tn_address1',     value: 'DOOR No.42,V.O.C.1ST STREET,MELACHOKKANATHAPURAM', category: 'address_tn', label: 'Address Line 1', type: 'text' },
  { key: 'tn_address2',     value: 'BODINAYAKANUR, THENI-625582 TAMIL NADU CODE:33 Mobile:8610943865', category: 'address_tn', label: 'Address Line 2', type: 'text' },
  { key: 'tn_dispatch',     value: 'AMAZING SPICE PARK PVT LTD WARD No.6 ELLIKKANAM DOOR No.650 NEDUMKANDAM IDUKKI KERALA CODE:32', category: 'address_tn', label: 'Dispatch Address', type: 'text' },
  { key: 'tn_phone',        value: '8610943865',    category: 'address_tn', label: 'Phone',                   type: 'text' },
  { key: 'tn_email',        value: 'idealspicesbodi@gmail.com', category: 'address_tn', label: 'Email',       type: 'text' },
  { key: 'tn_gstin',        value: '33AAICI5415L1ZH', category: 'address_tn', label: 'GSTIN',                 type: 'text' },
  { key: 'tn_branch',       value: 'BODINAYAKANUR', category: 'address_tn', label: 'Office Branch',           type: 'text' },

  // ── SISTER COMPANY (ASP) ───────────────────────────────────
  { key: 's_logo',          value: 'ASP',            category: 'sister',    label: 'Logo Code',                type: 'text' },
  { key: 's_company',       value: 'AMAZING SPICE PARK PRIVATE LIMITED', category: 'sister', label: 'Company Name', type: 'text' },
  { key: 's_short_name',    value: 'AMAZING SPICE PARK PVT LTD', category: 'sister', label: 'Short Name',     type: 'text' },
  { key: 's_address1',      value: 'WARD No.6, ELLIKKANAM, DOOR No.650, NEDUMKANDAM', category: 'sister', label: 'Address Line 1', type: 'text' },
  { key: 's_address2',      value: 'UDUMBANCHOLA, IDUKKI-685553 KERALA CODE:32 MOBILE:9843338633', category: 'sister', label: 'Address Line 2', type: 'text' },
  { key: 's_phone',         value: '9843338633',     category: 'sister',    label: 'Phone',                    type: 'text' },
  { key: 's_email',         value: 'amazingspicepark@gmail.com', category: 'sister', label: 'Email',           type: 'text' },
  { key: 's_gstin',         value: '32ABDCA2636B1ZE', category: 'sister',   label: 'GSTIN',                    type: 'text' },
  { key: 's_cin',           value: 'U46305KL2025PTC095544', category: 'sister', label: 'CIN',                  type: 'text' },
  // Partnership Firm toggle for the sister (ASP) company — mirrors the
  // company-level pair so an ASP-active install can show "Partnership:
  // <name>" on its letterhead instead of a CIN.
  { key: 's_is_partnership',  value: 'false',        category: 'sister',    label: 'Partnership Firm',        type: 'boolean' },
  { key: 's_partnership_name', value: '',            category: 'sister',    label: 'Partnership Name / No.',  type: 'text' },
  { key: 's_pan',           value: 'ABDCA2636B',    category: 'sister',    label: 'PAN',                      type: 'text' },
  { key: 's_fssai',         value: '',               category: 'sister',    label: 'FSSAI No.',                type: 'text' },
  { key: 's_sbl',           value: 'CS/55884/950/2026-27', category: 'sister', label: 'SBL No.',               type: 'text' },

  // ── BRANCHES ───────────────────────────────────────────────
  { key: 'br1',             value: 'NEDUMKANDAM',    category: 'branches',  label: 'Branch 1',                type: 'text' },
  { key: 'br2',             value: 'UDUBANCHOLA',    category: 'branches',  label: 'Branch 2',                type: 'text' },
  { key: 'br3',             value: 'MARUKKUMTOTTI',  category: 'branches',  label: 'Branch 3',                type: 'text' },
  { key: 'br4',             value: 'ANAVILASAM',     category: 'branches',  label: 'Branch 4',                type: 'text' },
  { key: 'br5',             value: 'VANDANMEDU',     category: 'branches',  label: 'Branch 5',                type: 'text' },
  { key: 'br6',             value: '',               category: 'branches',  label: 'Branch 6',                type: 'text' },
  { key: 'br7',             value: '',               category: 'branches',  label: 'Branch 7',                type: 'text' },
  { key: 'br8',             value: '',               category: 'branches',  label: 'Branch 8',                type: 'text' },
  { key: 'br9',             value: '',               category: 'branches',  label: 'Branch 9',                type: 'text' },
  { key: 'br1_tel',         value: '9786069799',     category: 'branches',  label: 'Branch 1 Mobile',         type: 'text' },
  { key: 'br2_tel',         value: '',               category: 'branches',  label: 'Branch 2 Mobile',         type: 'text' },
  { key: 'br3_tel',         value: '9080248574',     category: 'branches',  label: 'Branch 3 Mobile',         type: 'text' },
  { key: 'br4_tel',         value: '',               category: 'branches',  label: 'Branch 4 Mobile',         type: 'text' },
  { key: 'br5_tel',         value: '',               category: 'branches',  label: 'Branch 5 Mobile',         type: 'text' },
  { key: 'br6_tel',         value: '',               category: 'branches',  label: 'Branch 6 Mobile',         type: 'text' },
  { key: 'br7_tel',         value: '',               category: 'branches',  label: 'Branch 7 Mobile',         type: 'text' },
  { key: 'br8_tel',         value: '',               category: 'branches',  label: 'Branch 8 Mobile',         type: 'text' },

  // ── RATES ──────────────────────────────────────────────────
  { key: 'commission',      value: '1',              category: 'rates',     label: 'Commission %',             type: 'number' },
  { key: 'hpc',             value: '10',             category: 'rates',     label: 'Handling %',               type: 'number' },
  { key: 'deduction1',      value: '1.25',           category: 'rates',     label: 'Deduction (Pooler)',       type: 'number' },
  { key: 'deduction2',      value: '1.25',           category: 'rates',     label: 'Deduction (Dealer)',       type: 'number' },
  // Used by flag_discount_in_prate — applies ONLY to Grade 1 lots.
  { key: 'deduction1_inclusive', value: '1.25',      category: 'rates',     label: 'Deduction (Pooler) — discount-inclusive (Grade 1 only)', type: 'number' },
  { key: 'asp_profit_pooler', value: '0.75',         category: 'rates',     label: 'ASP Profit Ratio (Pooler)', type: 'number' },
  { key: 'asp_profit_dealer', value: '0.75',         category: 'rates',     label: 'ASP Profit Ratio (Dealer)', type: 'number' },
  { key: 'isp_profit_pooler', value: '0.5',          category: 'rates',     label: 'ISP Profit Ratio (Pooler)', type: 'number' },
  { key: 'isp_profit_dealer', value: '0.5',          category: 'rates',     label: 'ISP Profit Ratio (Dealer)', type: 'number' },
  { key: 'refund',          value: '1.9',            category: 'rates',     label: 'Sample Refund (Kgs)',      type: 'number' },
  { key: 'sb_refund',       value: '2.85',           category: 'rates',     label: 'SB Sample Refund (Kgs)',   type: 'number' },
  { key: 'gst_goods',       value: '5',              category: 'rates',     label: 'GST Goods Rate %',         type: 'number' },
  { key: 'gst_service',     value: '18',             category: 'rates',     label: 'GST Service Rate %',       type: 'number' },
  { key: 'tcs_tds',         value: '0.1',            category: 'rates',     label: 'TCS / TDS Rate %',         type: 'number' },
  { key: 'tds_purchase_rate', value: '0.1',          category: 'rates',     label: 'TDS on Purchase Rate % (Section 194Q)',  type: 'number' },
  { key: 'tds_threshold',   value: '5000000',        category: 'rates',     label: 'TDS / TCS Annual Threshold (₹) — default ₹50 lakh per Section 194Q/206C(1H)',  type: 'number' },
  { key: 'gunny_rate',      value: '165',            category: 'rates',     label: 'Gunny Rate (₹)',           type: 'number' },
  // When flag_inter_transport / flag_inter_insurance is OFF the inter-state rate is forced to 0.
  { key: 'flag_inter_transport', value: 'true',      category: 'rates',     label: 'Inter-State Transport (use inter-state transport rate)', type: 'boolean' },
  { key: 'transport',       value: '2.5',            category: 'rates',     label: 'Transport (₹/kg)',         type: 'number' },
  { key: 'flag_inter_insurance', value: 'true',      category: 'rates',     label: 'Inter-State Insurance (use inter-state insurance rate)', type: 'boolean' },
  { key: 'insurance',       value: '0.75',           category: 'rates',     label: 'Insurance (₹/kg)',         type: 'number' },
  { key: 'flag_local_transport', value: 'true',      category: 'rates',     label: 'Local Transport (use local transport rate)', type: 'boolean' },
  { key: 'local_transport', value: '2.5',            category: 'rates',     label: 'Local Transport (₹/kg)',   type: 'number' },
  { key: 'flag_local_insurance', value: 'true',      category: 'rates',     label: 'Local Insurance (use local insurance rate)', type: 'boolean' },
  { key: 'local_insurance', value: '0.75',           category: 'rates',     label: 'Local Insurance (₹/kg)',   type: 'number' },
  { key: 'discount_pct',    value: '0',              category: 'rates',     label: 'Discount %',               type: 'number' },
  { key: 'discount_days',   value: '0',              category: 'rates',     label: 'No. of Days for Discount', type: 'number' },
  { key: 'dealer_days',     value: '0',              category: 'rates',     label: 'No. of Days for Dealer',   type: 'number' },
  { key: 'addl_charge_name',  value: '',             category: 'rates',     label: 'Additional Charge — Name', type: 'text' },
  { key: 'addl_charge_value', value: '0',            category: 'rates',     label: 'Additional Charge — % of cardamom amount (0 to disable)', type: 'number' },

  // ── HSN / SAC CODES ────────────────────────────────────────
  { key: 'hsn_cardamom',    value: '09083120',       category: 'hsn',       label: 'Cardamom HSN',             type: 'text' },
  { key: 'hsn_gunny',       value: '63051040',       category: 'hsn',       label: 'Gunny HSN',                type: 'text' },
  { key: 'sac_transport',   value: '996791',         category: 'hsn',       label: 'Transport SAC',            type: 'text' },
  { key: 'sac_insurance',   value: '997136',         category: 'hsn',       label: 'Insurance SAC',            type: 'text' },
  { key: 'sac_service',     value: '996111',         category: 'hsn',       label: 'Service SAC',              type: 'text' },

  // ── BANK DETAILS ───────────────────────────────────────────
  { key: 'bank_kl_name',    value: 'FEDERAL BANK - PUTTADY', category: 'bank', label: 'Kerala Bank Name',      type: 'text' },
  { key: 'bank_kl_acct',    value: '10735500094452', category: 'bank',      label: 'Kerala Account No.',       type: 'text' },
  { key: 'bank_kl_ifsc',    value: 'FDRL0001073',   category: 'bank',      label: 'Kerala IFSC Code',         type: 'text' },
  { key: 'bank_tn_name',    value: 'CITY UNION BANK-BODINAYAKANUR', category: 'bank', label: 'TN Bank Name',   type: 'text' },
  { key: 'bank_tn_acct',    value: '510909010383556', category: 'bank',     label: 'TN Account No.',           type: 'text' },
  { key: 'bank_tn_ifsc',    value: 'CIUB0000346',   category: 'bank',      label: 'TN IFSC Code',             type: 'text' },

  // ── SEASON ─────────────────────────────────────────────────
  { key: 'season',          value: '2026 - 27',      category: 'season',    label: 'Season Name',              type: 'text' },
  { key: 'season_short',    value: '26-27',          category: 'season',    label: 'Season Short',             type: 'text' },
  { key: 'season_start',    value: '2026-04-01',     category: 'season',    label: 'FY Start Date',            type: 'date' },
  { key: 'season_end',      value: '2027-03-31',     category: 'season',    label: 'FY End Date',              type: 'date' },

  // ── INVOICE SETTINGS ───────────────────────────────────────
  { key: 'inv_prefix',      value: 'ISP',            category: 'invoice',   label: 'Invoice Prefix',           type: 'text' },
  { key: 'inv_prefix_sister', value: 'ASP',          category: 'invoice',   label: 'Sister Invoice Prefix (Other Ref.)', type: 'text' },
  { key: 'separator',       value: '-',              category: 'invoice',   label: 'Separator Symbol',         type: 'text' },
  { key: 'hsn_cardamom',    value: '09083120',       category: 'invoice',   label: 'HSN/SAC — Cardamom',       type: 'text' },
  { key: 'hsn_gunny',       value: '63051040',       category: 'invoice',   label: 'HSN/SAC — Gunny',          type: 'text' },
  { key: 'dispatched_through_isp', value: '',         category: 'invoice',   label: 'Dispatched Through (ISP)', type: 'text' },
  { key: 'dispatched_through_asp', value: '',         category: 'invoice',   label: 'Dispatched Through (ASP)', type: 'text' },
  { key: 'dispatch_destination', value: 'NEDUMKANDAM', category: 'invoice', label: 'Dispatch Destination',     type: 'text' },
  { key: 'duplicate_text',  value: 'DUMMY INVOICE',  category: 'invoice',   label: 'Dummy Invoice Text',       type: 'text' },
  { key: 'commission_bill', value: 'COMMISSION BILL', category: 'invoice',  label: 'Commission Bill Name',     type: 'text' },
  { key: 'memorandum_text', value: 'MEMORANDAM OF CARDAMOM SOLD THROUGH', category: 'invoice', label: 'Memorandum Text', type: 'text' },
  { key: 'signature_text',  value: 'Signature of the Authorised Buyer', category: 'invoice', label: 'Signature Label', type: 'text' },

  // ── FEATURE FLAGS ──────────────────────────────────────────
  // Sidebar visibility toggles for optional document types. When OFF,
  // the sidebar entry, related dashboard tiles, and Backup → Danger Zone
  // Delete buttons all hide via the body[data-feat-*] / .feat-* CSS
  // pairing. Server endpoints stay live (data already in those tables
  // remains accessible if the user re-enables) — only the UI surface
  // is gated.
  { key: 'flag_bills',         value: 'true',        category: 'flags',     label: 'Bills of Supply Module',       type: 'boolean' },
  { key: 'flag_debit_notes',   value: 'true',        category: 'flags',     label: 'Debit Notes Module',           type: 'boolean' },
  { key: 'flag_pooling',    value: 'false',          category: 'flags',     label: 'Pooling (Single State)',    type: 'boolean' },
  { key: 'flag_sister',     value: 'true',           category: 'flags',     label: 'Sister Concern Active',    type: 'boolean' },
  { key: 'flag_tnpa',       value: 'true',           category: 'flags',     label: 'ASP Ship To Address',      type: 'boolean' },
  { key: 'flag_sample',     value: 'false',          category: 'flags',     label: 'Discount in Invoice',      type: 'boolean' },
  { key: 'flag_dispatch',   value: 'true',           category: 'flags',     label: 'Show Dispatch Address',    type: 'boolean' },
  { key: 'flag_ship',       value: 'true',           category: 'flags',     label: 'Show Ship To Address',     type: 'boolean' },
  { key: 'flag_hsn',        value: 'true',           category: 'flags',     label: 'Show HSN Codes',           type: 'boolean' },
  { key: 'flag_bank',       value: 'true',           category: 'flags',     label: 'Bank Details in Invoice',  type: 'boolean' },
  { key: 'flag_tds_purchase', value: 'true',         category: 'flags',     label: 'TDS on Purchase Invoice',  type: 'boolean' },
  { key: 'flag_tds_sales',  value: 'false',          category: 'flags',     label: 'TDS on Sales Invoice',     type: 'boolean' },
  { key: 'flag_rtds_inv',   value: 'true',           category: 'flags',     label: 'TDS in ASP Purchase',      type: 'boolean' },
  { key: 'flag_wgst',       value: 'false',          category: 'flags',     label: 'TDS on Full Invoice Amount', type: 'boolean' },
  { key: 'flag_disc_gst',   value: 'false',          category: 'flags',     label: 'Discount includes GST',    type: 'boolean' },
  // Rolls the per-lot Discount into P_Rate using deduction1_inclusive — ONLY for Grade 1 lots.
  { key: 'flag_discount_in_prate', value: 'false',   category: 'flags',     label: 'Roll Discount into P_Rate (Grade 1 only)', type: 'boolean' },
  { key: 'flag_debit_note', value: 'false',          category: 'flags',     label: 'Debit Note for Discount',  type: 'boolean' },
  { key: 'flag_invoice_stripe', value: 'true',       category: 'flags',     label: 'Alternate Row Stripe in Invoice', type: 'boolean' },
  { key: 'flag_dummy',      value: 'true',           category: 'flags',     label: 'Allow Dummy Invoices',     type: 'boolean' },
  { key: 'flag_round',      value: 'true',           category: 'flags',     label: 'Round Invoice Amounts',    type: 'boolean' },
  { key: 'flag_eway',       value: 'false',          category: 'flags',     label: 'ASP eWay Bill / Transport', type: 'boolean' },
  { key: 'flag_export',     value: 'false',          category: 'flags',     label: 'Export Invoices',          type: 'boolean' },
  // e-Auction-only fields. Both gated by business_mode === 'e-Auction'
  // on top of the flag, so flipping the mode away from e-Auction
  // automatically hides them regardless of the flag value. Default OFF
  // so a fresh install matches the pre-feature behaviour.
  { key: 'flag_crop_receipt',  value: 'false',       category: 'flags',     label: 'Crop Receipt (e-Auction)',   type: 'boolean' },
  { key: 'flag_reserved_price',value: 'false',       category: 'flags',     label: 'Reserved Price (e-Auction)', type: 'boolean' },
  // Price List Mapping — sister tool of the Lots → Price Import button.
  // When ON, a "Price List Mapping" sidebar entry (under Lots) and a
  // quick-access "🗺 Price List Mapping" button on the Lots toolbar
  // appear; the tab fills the CODE column of a Price List (Before) sheet
  // from the Buyers master. Gated by business_mode === 'e-Auction' on
  // top of the flag — flipping the mode away from e-Auction hides both
  // surfaces regardless of the flag value. Default ON because most
  // e-Auction operators use the mapping flow.
  { key: 'flag_price_list_mapping', value: 'true',   category: 'flags',     label: 'Price List Mapping (e-Auction)', type: 'boolean' },
  // Operational feature flags (default ON to preserve existing behaviour;
  // flag_price_check defaults OFF — matches the previous unset = gate-off state).
  { key: 'flag_whatsapp',       value: 'true',  category: 'flags',     label: 'WhatsApp Share Buttons',                 type: 'boolean' },
  { key: 'flag_set_buyer',      value: 'true',  category: 'flags',     label: 'Lots \u2192 Set Buyer (bulk action)',       type: 'boolean' },
  { key: 'flag_print_purchase', value: 'true',  category: 'flags',     label: 'Print Selected Purchase (ASP / Kerala)', type: 'boolean' },
  { key: 'flag_price_check',    value: 'false', category: 'flags',     label: 'Price Check + transaction gate',         type: 'boolean' },

  // ── BUSINESS MODE ──────────────────────────────────────────
  // Single-mode e-Auction build. Default flipped from 'e-Trade' to 'e-Auction'
  // so fresh installs match the readonly UI input on the Settings → Business
  // Mode panel and the Spice Board sidebar entry shows up immediately.
  { key: 'business_mode',   value: 'e-Auction',      category: 'mode',      label: 'Business Mode',            type: 'select' },
  { key: 'business_state',  value: 'TAMIL NADU',     category: 'mode',      label: 'Business State',           type: 'select' },

  // ── INTEGRATIONS ───────────────────────────────────────────
  { key: 'gst_api_key',     value: '',               category: 'integrations', label: 'GST Lookup API Key (gstincheck.co.in)', type: 'text' },

  // ── TALLY EXPORT ──────────────────────────────────────────
  // Settings here mirror the macro's Configration form (UserForm1) field-for-field.
  // Identity & defaults
  { key: 'tally_company_name',     value: 'Ideal Spices Private Limited',      category: 'tally', label: 'ISP Tally Company Name (used for Sales — must match Tally company exactly)', type: 'text' },
  { key: 'tally_asp_company_name', value: 'Amazing Spice Park Private Limited', category: 'tally', label: 'ASP Tally Company Name (used for RD / URD Purchase / Debit Note — must match Tally company exactly)', type: 'text' },
  { key: 'tally_season',          value: '2026-27',        category: 'tally', label: 'Season Suffix',                  type: 'text' },
  { key: 'tally_separator',       value: '/',              category: 'tally', label: 'Voucher Separator',              type: 'text' },
  { key: 'tally_inv_prefix',      value: 'ISP/',           category: 'tally', label: 'ISP Voucher Prefix',             type: 'text' },
  { key: 'tally_ainv_prefix',     value: 'ASP/',           category: 'tally', label: 'ASP Voucher Prefix',             type: 'text' },
  { key: 'tally_state_code',      value: '33',             category: 'tally', label: 'Home GSTIN State Code (intra)',  type: 'text' },
  { key: 'tally_state_code_amazing', value: '32',          category: 'tally', label: 'ASP Home GSTIN State Code',      type: 'text' },
  { key: 'tally_home_state',      value: 'Tamil Nadu',     category: 'tally', label: 'Home Place of Supply',           type: 'text' },
  { key: 'tally_urd_state',       value: 'Kerala',         category: 'tally', label: 'URD Purchase State (agriculturist)', type: 'text' },

  // Mode toggles (mirror the macro checkboxes)
  { key: 'tally_amazing_mode',    value: 'false',          category: 'tally', label: 'Amazing (generate as ASP / sister company)',type: 'boolean' },
  { key: 'tally_detailed',        value: 'true',           category: 'tally', label: 'Detailed Inv (one inventory entry per lot)',type: 'boolean' },
  { key: 'tally_dispatch_from',   value: 'true',           category: 'tally', label: 'Dispatch (include Dispatch-From block)',    type: 'boolean' },
  { key: 'tally_round_enabled',   value: 'true',           category: 'tally', label: 'Round (Round On/Off ledger)',               type: 'boolean' },
  { key: 'tally_tcs_enabled',     value: 'false',          category: 'tally', label: 'TCS (apply on Sales when applicable)',      type: 'boolean' },
  { key: 'tally_tds_enabled',     value: 'false',          category: 'tally', label: 'TDS (apply 194Q on RD Purchases)',          type: 'boolean' },
  { key: 'tally_optional',        value: 'false',          category: 'tally', label: 'Optional (mark vouchers as Optional)',      type: 'boolean' },
  { key: 'tally_dn_exempt',       value: 'false',          category: 'tally', label: 'Exempted (Debit Note: skip GST tax ledgers)', type: 'boolean' },
  { key: 'tally_local_transport', value: 'true',           category: 'tally', label: 'Local Transport (use local transport rate)', type: 'boolean' },
  { key: 'tally_local_insurance', value: 'true',           category: 'tally', label: 'Local Insurance (use local insurance rate)', type: 'boolean' },
  { key: 'tally_ship_to',         value: 'false',          category: 'tally', label: 'Ship To (override consignee with separate Ship-To party)', type: 'boolean' },

  // Sales Account Ledgers (Cardamom)
  { key: 'tally_sales_inter',     value: 'Cardamom Sales 5%',          category: 'tally', label: 'Cardamom Inter-State Sales',  type: 'text' },
  { key: 'tally_sales_intra',     value: 'Cardamom Sales 5% - Local',  category: 'tally', label: 'Cardamom Local Sales',        type: 'text' },
  { key: 'tally_sales_export',    value: 'Cardamom Sales - Export',    category: 'tally', label: 'Cardamom Export Sales (Deemed)', type: 'text' },

  // Sales Account Ledgers (Gunny)
  { key: 'tally_gunny_inter',     value: 'Gunny Sales 5%',             category: 'tally', label: 'Gunny Interstate Sales',      type: 'text' },
  { key: 'tally_gunny_intra',     value: 'Gunny Sales 5% - Local',     category: 'tally', label: 'Gunny Local Sales',           type: 'text' },
  { key: 'tally_gunny_export',    value: 'Gunny Sales - Export',       category: 'tally', label: 'Gunny Export Sales',          type: 'text' },

  // Dealer-Side Sales (when ISP sells to a dealer)
  { key: 'tally_dealer_sale_inter', value: 'Interstate Dealer-Purchase', category: 'tally', label: 'Interstate Dealer-Purch (sales-side)', type: 'text' },
  { key: 'tally_dealer_sale_intra', value: 'Local Dealer-Purchase',      category: 'tally', label: 'Local Dealer-Purcha (sales-side)',      type: 'text' },

  // RD Purchase ledgers (when ISP buys from a dealer)
  { key: 'tally_purchase_dealer',     value: 'Trade Purchase From Dealer',category: 'tally', label: 'Trade Purchase From Dealer (base; gets -Local / -Inter_State suffix)', type: 'text' },
  { key: 'tally_purchase_dealer_inter', value: 'Interstate Dealer',      category: 'tally', label: 'Interstate Dealer (purchase-side)',     type: 'text' },
  { key: 'tally_purchase_dealer_intra', value: 'Local Dealer',           category: 'tally', label: 'Local Dealer (purchase-side)',          type: 'text' },

  // Agriculturist & TDS-on-sales
  { key: 'tally_purchase_auction',value: 'Purchase From Agriculturist', category: 'tally', label: 'Purchase From Agriculturist (URD ledger)', type: 'text' },
  { key: 'tally_tds_paid_sales',  value: 'TDS Paid on Sales',           category: 'tally', label: 'TDS Paid on Sales',           type: 'text' },

  // Tax Ledger Names — Sales 5% (output) and Purchase (input)
  { key: 'tally_cgst',            value: 'OUTPUT CGST 2.5%',           category: 'tally', label: 'CGST 2.5% (output)',          type: 'text' },
  { key: 'tally_sgst',            value: 'OUTPUT SGST 2.5%',           category: 'tally', label: 'SGST 2.5% (output)',          type: 'text' },
  { key: 'tally_igst',            value: 'OUTPUT IGST 5%',             category: 'tally', label: 'IGST 5% (output)',            type: 'text' },
  { key: 'tally_cgst_input',      value: 'INPUT CGST 2.5%',            category: 'tally', label: 'INPUT CGST 2.5%',             type: 'text' },
  { key: 'tally_sgst_input',      value: 'INPUT SGST 2.5%',            category: 'tally', label: 'INPUT SGST 2.5%',             type: 'text' },
  { key: 'tally_igst_input',      value: 'INPUT IGST 5%',              category: 'tally', label: 'INPUT IGST 5%',               type: 'text' },
  { key: 'tally_tcs',             value: 'TCS on Sale of Goods',       category: 'tally', label: 'TCS on Sale of Goods',        type: 'text' },
  { key: 'tally_tds_ledger',      value: 'TDS on Purchase of Goods',   category: 'tally', label: 'TDS on Purchase of Goods', type: 'text' },

  // Tax Ledger Names — Debit Note 18%
  { key: 'tally_dn_discount',     value: 'Discount on Purchase',       category: 'tally', label: 'Discount on Purch (Debit Note ledger)', type: 'text' },
  { key: 'tally_dn_cgst',         value: 'OUTPUT CGST 9%',             category: 'tally', label: 'CGST 9% (Debit Note)',        type: 'text' },
  { key: 'tally_dn_sgst',         value: 'OUTPUT SGST 9%',             category: 'tally', label: 'SGST 9% (Debit Note)',        type: 'text' },
  { key: 'tally_dn_igst',         value: 'OUTPUT IGST 18%',            category: 'tally', label: 'IGST 18% (Debit Note)',       type: 'text' },
  { key: 'tally_dn_gst_rate',     value: '18',                         category: 'tally', label: 'Debit Note GST Rate %',       type: 'number' },

  // Other operational ledgers
  { key: 'tally_commission',      value: 'Commission-Planter',         category: 'tally', label: 'Commission-Planter',          type: 'text' },
  { key: 'tally_cash_handling',   value: 'Cash Handling Charges',      category: 'tally', label: 'Cash Handling Charges',       type: 'text' },
  { key: 'tally_cash_handling_planter', value: 'Cash Handling Charges-Planter', category: 'tally', label: 'Cash Handling Charges-Planter', type: 'text' },
  { key: 'tally_chc_planter',     value: 'CHC From Planter',           category: 'tally', label: 'CHC From Planter',            type: 'text' },
  { key: 'tally_sample_planter',  value: 'Sample Refund to Planter',   category: 'tally', label: 'Sample Refund to Planter',    type: 'text' },
  { key: 'tally_sample_dealer',   value: 'Sample Refund to Dealer',    category: 'tally', label: 'Sample Refund to Dealer',     type: 'text' },
  { key: 'tally_sample_stock',    value: 'false',                      category: 'tally', label: 'Stock (track sample refund as inventory)', type: 'boolean' },
  { key: 'tally_round',           value: 'Round On/Off',               category: 'tally', label: 'Round On/Off Ledger',         type: 'text' },
  { key: 'tally_transport',       value: 'Transport Charges',          category: 'tally', label: 'Transport Charges Ledger',    type: 'text' },
  { key: 'tally_insurance',       value: 'Insurance Charges',          category: 'tally', label: 'Insurance Charges Ledger',    type: 'text' },

  // Tax / commercial rates (the right-hand "Tax Rate" / "Item Rates" block)
  { key: 'tally_gst_rate',        value: '5',                          category: 'tally', label: 'GST Goods Rate %',            type: 'number' },
  { key: 'tally_service_rate',    value: '18',                         category: 'tally', label: 'Service Rate % (DN/Discount)', type: 'number' },
  { key: 'tally_tcs_rate',        value: '0.1',                        category: 'tally', label: 'TCS / TDS Rate %',            type: 'number' },
  { key: 'tally_export_rate',     value: '0',                          category: 'tally', label: 'Export GST Rate %',           type: 'number' },
  { key: 'tally_sample_kgs',      value: '1.900',                      category: 'tally', label: 'Sample Refund (Kgs)',         type: 'number' },
  { key: 'tally_unit_rate',       value: '0.1',                        category: 'tally', label: 'Sample Unit Rate (per Kg)',   type: 'number' },
  { key: 'tally_gunny_rate',      value: '165',                        category: 'tally', label: 'Gunny Rate (₹ per bag)',      type: 'number' },
  { key: 'tally_transport_rate',  value: '2.50',                       category: 'tally', label: 'Transport Rate (₹/Kg, inter-state)', type: 'number' },
  { key: 'tally_local_trans_rate',value: '2.50',                       category: 'tally', label: 'Local Transport Rate (₹/Kg)', type: 'number' },
  { key: 'tally_insurance_rate',  value: '0.75',                       category: 'tally', label: 'Insurance Rate (₹/₹1000)',    type: 'number' },
  { key: 'tally_local_ins_rate',  value: '0.75',                       category: 'tally', label: 'Local Insurance Rate (₹/₹1000)', type: 'number' },

  // Stock Item Names + HSN
  { key: 'tally_item_cardamom',   value: 'Cardamom',                   category: 'tally', label: 'Stock Item — Cardamom',       type: 'text' },
  { key: 'tally_item_gunny',      value: 'Gunny Bag',                  category: 'tally', label: 'Stock Item — Gunny',          type: 'text' },
  { key: 'tally_hsn_cardamom',    value: '09083120',                   category: 'tally', label: 'HSN — Cardamom',              type: 'text' },
  { key: 'tally_hsn_gunny',       value: '63051040',                   category: 'tally', label: 'HSN — Gunny',                 type: 'text' },
  { key: 'tally_hsn_service',     value: '996111',                     category: 'tally', label: 'SAC — Service / Discount',    type: 'text' },
  { key: 'tally_hsn_transport',   value: '996791',                     category: 'tally', label: 'SAC — Transport',             type: 'text' },
  { key: 'tally_hsn_insurance',   value: '997136',                     category: 'tally', label: 'SAC — Insurance',             type: 'text' },

  // Dispatch-from address (optional override; defaults to Sister Company config)
  { key: 'tally_dispatch_company',value: '',                           category: 'tally', label: 'Dispatch-From Company (blank = use sister)',     type: 'text' },
  { key: 'tally_dispatch_address',value: '',                           category: 'tally', label: 'Dispatch-From Address (blank = use sister)',     type: 'text' },
  { key: 'tally_dispatch_place',  value: '',                           category: 'tally', label: 'Dispatch-From Place (blank = use sister)',       type: 'text' },
  { key: 'tally_dispatch_pin',    value: '',                           category: 'tally', label: 'Dispatch-From PIN (blank = use sister)',         type: 'text' },
  { key: 'tally_dispatch_state',  value: '',                           category: 'tally', label: 'Dispatch-From State (blank = use sister)',       type: 'text' },
  { key: 'tally_dispatch_gstin',  value: '',                           category: 'tally', label: 'Dispatch-From GSTIN (blank = use sister)',       type: 'text' },
  // ── E-way bill DISTANCE estimation ────────────────────────────
  // Auto-fills <DISTANCE> on ISP sales vouchers using haversine ×
  // multiplier between dispatch PIN and consignee PIN. The multiplier
  // converts straight-line km to road km — bump it for hilly terrain
  // (Western Ghats), lower it for plains. Per-invoice manual override
  // is supported via the invoices.distance_km column.
  //
  // CAVEAT: haversine × multiplier is a rough estimate. For Western
  // Ghats routes (Kerala↔Tamil Nadu cardamom belt) it can under-shoot
  // real road distance by 30–50%. The auto-compute is OFF by default
  // — turn it on only if you've tuned the multiplier for your routes
  // or you're OK with the estimate. The recommended workflow is to
  // populate invoices.distance_km manually (or via an external tool)
  // and let the generator use those values verbatim.
  { key: 'distance_auto_enabled',    value: 'false',                   category: 'tally', label: 'Auto-fill <DISTANCE> from PIN coordinates (rough estimate — manual override always wins)', type: 'check' },
  { key: 'distance_road_multiplier', value: '1.5',                     category: 'tally', label: 'Road-distance multiplier (haversine × this = road km)', type: 'number' },

  // ── LOT ENTRY DEFAULTS ─────────────────────────────────────
  // Pre-populate the Lot Entry form so field staff don't re-type the
  // same numbers every lot. Sample weight is the cardamom sample
  // taken from each lot for grading, typically a constant per season.
  // Edit timeout controls how long non-admin users can edit their own
  // saved lots (0 = unlimited).
  { key: 'sample_weight',     value: '0.000',  category: 'lot_entry', label: 'Default Sample Weight (kg)',            type: 'number' },
  { key: 'show_moisture',     value: 'false',  category: 'lot_entry', label: 'Show Moisture Column',                  type: 'boolean' },
  { key: 'default_litre',     value: '',       category: 'lot_entry', label: 'Default Litre Weight',                  type: 'text' },
  // Gunny tare default + the unified "extra lot fields" toggle. Gunny
  // Weight is the per-bag tare: on the entry form the operator types
  // Weight-w/-Gunny and Net Wt auto-derives as WwG − (gunny_weight ×
  // bags). show_extra_lot_fields is a single switch that reveals the
  // Weight-w/-Gunny + Gunny Wt pair AND the Crop Receipt + Reserved
  // Price inputs together (OR'ed with the per-field e-Auction flags so
  // an install already using those flags keeps working).
  { key: 'gunny_weight',          value: '0.000', category: 'lot_entry', label: 'Default Gunny Weight (kg)',                          type: 'number' },
  { key: 'show_extra_lot_fields', value: 'false', category: 'lot_entry', label: 'Show Extra Lot Fields (Crop Receipt, Reserved Price)', type: 'boolean' },
  { key: 'default_crop_type', value: '',       category: 'lot_entry', label: 'Default Crop Type',                     type: 'text' },
  { key: 'edit_enabled',      value: 'true',   category: 'lot_entry', label: 'Allow Lot Edits (non-admin)',           type: 'boolean' },
  { key: 'edit_timeout_sec',  value: '0',      category: 'lot_entry', label: 'Edit Timeout (sec; 0 = no limit)',      type: 'number' },
  // Default lot receipt format. The Lot Entry print modal lets the
  // user override this per-print, but this setting decides which
  // option is pre-selected. "compact" = thermal-printer slip;
  // "detailed" = A4-style with seller bank details.
  { key: 'lot_receipt_format', value: 'detailed', category: 'lot_entry', label: 'Lot Receipt Format (compact|detailed)', type: 'text' },

  // ── SPICE BOARD REPORTS ───────────────────────────────────
  // Statutory cardamom-auction reports submitted to the Spices Board.
  // formd_places drives the "Place of Auction" dropdown on the Spice
  // Board → FORM-D panel — one place per line. When the operator picks
  // a value there, it overrides the configured branch on the printed
  // Form-D for that one report run.
  { key: 'formd_places', value: '', category: 'spice_board',
    label: 'FORM-D Place of Auction (one per line)',
    type: 'textarea' },
];

const CATEGORIES = {
  mode:       { order: 0, title: 'Business Mode',        icon: '⚙' },
  company:    { order: 1, title: 'Company Details',       icon: '🏢' },
  address_kl: { order: 2, title: 'Address (Kerala)',      icon: '📍' },
  address_tn: { order: 3, title: 'Address (Tamil Nadu)',  icon: '📍' },
  sister:     { order: 4, title: 'Sister Company (ASP)',  icon: '🤝' },
  branches:   { order: 5, title: 'Branches & Contacts',  icon: '🏪' },
  rates:      { order: 6, title: 'Rates & Charges',       icon: '💰' },
  hsn:        { order: 7, title: 'HSN / SAC Codes',       icon: '🏷' },
  bank:       { order: 8, title: 'Bank Details',          icon: '🏦' },
  season:     { order: 9, title: 'Season / Financial Year', icon: '📅' },
  invoice:    { order: 10, title: 'Invoice Settings',     icon: '📄' },
  flags:      { order: 11, title: 'Feature Flags',        icon: '🔧' },
  lot_entry:  { order: 11.5, title: 'Lot Entry Defaults',  icon: '📝', description: 'Defaults used by the Lot Entry tab — sample weight, gunny tare, default crop, moisture visibility, extra-field (crop receipt / reserved price) visibility, edit window, and receipt format.' },
  integrations: { order: 12, title: 'Integrations',       icon: '🔌', description: 'Optional third-party services. The GST API key enables auto-fetching trade name and address when you enter a GSTIN. Get a free key at gstincheck.co.in — sign up, copy the key from your dashboard, paste here.' },
  tally:      { order: 13, title: 'To Tally',             icon: '📤', description: 'Configure all settings for the Tally XML export — laid out exactly like the original Configration form. Ledger names here MUST match what exists in your Tally company; if a ledger is missing or misspelled, Tally will reject the import.' },
  spice_board:{ order: 14, title: 'Spice Board Reports',   icon: '🌶', description: 'Statutory cardamom-auction reports. Place values entered below populate the Place of Auction dropdown on the FORM-D report.' },
};

function initCompanySettings(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS company_settings (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL DEFAULT '',
      category TEXT NOT NULL DEFAULT 'company',
      label TEXT NOT NULL DEFAULT '',
      field_type TEXT NOT NULL DEFAULT 'text'
    );
  `);

  const insert = db.prepare(
    'INSERT OR IGNORE INTO company_settings (key, value, category, label, field_type) VALUES (?, ?, ?, ?, ?)'
  );
  const seed = db.transaction(() => {
    for (const d of DEFAULTS) insert.run(d.key, d.value, d.category, d.label, d.type);
  });
  seed();

  // Migration: asp_profit was split into asp_profit_pooler and asp_profit_dealer.
  // If an existing DB still has the old row, copy its value to both new keys
  // (preserving the user's configured rate), then remove the legacy row.
  const legacy = db.prepare('SELECT value FROM company_settings WHERE key = ?').get('asp_profit');
  if (legacy && legacy.value != null && legacy.value !== '') {
    const upd = db.prepare('UPDATE company_settings SET value = ? WHERE key = ?');
    upd.run(legacy.value, 'asp_profit_pooler');
    upd.run(legacy.value, 'asp_profit_dealer');
    db.prepare('DELETE FROM company_settings WHERE key = ?').run('asp_profit');
    console.log('Migrated asp_profit → asp_profit_pooler/asp_profit_dealer (value=%s)', legacy.value);
  }

  // Migration: isp_profit was split into isp_profit_pooler and isp_profit_dealer.
  // These now drive P_Rate calculations for Kerala + e-Trade (ASP invoices).
  // Copy the legacy value into both new keys, then remove the legacy row.
  const legacyIsp = db.prepare('SELECT value FROM company_settings WHERE key = ?').get('isp_profit');
  if (legacyIsp && legacyIsp.value != null && legacyIsp.value !== '') {
    const upd = db.prepare('UPDATE company_settings SET value = ? WHERE key = ?');
    upd.run(legacyIsp.value, 'isp_profit_pooler');
    upd.run(legacyIsp.value, 'isp_profit_dealer');
    db.prepare('DELETE FROM company_settings WHERE key = ?').run('isp_profit');
    console.log('Migrated isp_profit → isp_profit_pooler/isp_profit_dealer (value=%s)', legacyIsp.value);
  } else {
    // No legacy value but the row may still exist from a prior install. Drop it.
    db.prepare('DELETE FROM company_settings WHERE key = ?').run('isp_profit');
  }

  // Migration: dispatched_through was split into _isp and _asp variants.
  // Copy the legacy value into both new keys (user can customize per company
  // afterward), then drop the legacy row.
  const legacyDT = db.prepare('SELECT value FROM company_settings WHERE key = ?').get('dispatched_through');
  if (legacyDT && legacyDT.value != null && legacyDT.value !== '') {
    const upd = db.prepare('UPDATE company_settings SET value = ? WHERE key = ?');
    upd.run(legacyDT.value, 'dispatched_through_isp');
    upd.run(legacyDT.value, 'dispatched_through_asp');
    db.prepare('DELETE FROM company_settings WHERE key = ?').run('dispatched_through');
    console.log('Migrated dispatched_through → dispatched_through_isp/_asp (value=%s)', legacyDT.value);
  } else {
    db.prepare('DELETE FROM company_settings WHERE key = ?').run('dispatched_through');
  }

  // Migration: business_mode was historically seeded as 'e-Trade' on older
  // installs, but this build is e-Auction only (the UI input is readonly
  // and forces 'e-Auction' on save). Silently rewrite any legacy value
  // so the Spice Board sidebar gate works without forcing every operator
  // to open Settings → Save once.
  try {
    const cur = db.prepare('SELECT value FROM company_settings WHERE key = ?').get('business_mode');
    if (cur && cur.value && String(cur.value).trim() !== 'e-Auction') {
      db.prepare('UPDATE company_settings SET value = ? WHERE key = ?').run('e-Auction', 'business_mode');
      console.log('Migrated business_mode "%s" → "e-Auction" (single-mode build)', cur.value);
    }
  } catch (_) { /* non-fatal */ }

  console.log('Company settings ready (%d defaults)', DEFAULTS.length);
}

function getSetting(db, key) {
  const r = db.prepare('SELECT value FROM company_settings WHERE key = ?').get(key);
  return r ? r.value : null;
  // ── Presets (ISP / ASP) for the "Company" category ───────────────────
  // Two named snapshots of the 8 fields in category='company' (logo,
  // trade_name, legal_name, short_name, pan, cin, fssai, sbl). The user
  // flips between them via the Logo Code dropdown; the active preset's
  // values overlay onto company_settings so invoice PDFs and exports
  // continue reading from the familiar flat key-value store.
  //
  // Schema: composite PK of (preset_code, field_key). One row per
  // (preset, field). active_preset_code is tracked in a meta row.
  db.exec(`
    CREATE TABLE IF NOT EXISTS company_presets (
      preset_code TEXT NOT NULL,
      field_key   TEXT NOT NULL,
      field_value TEXT NOT NULL DEFAULT '',
      PRIMARY KEY (preset_code, field_key)
    );
    CREATE TABLE IF NOT EXISTS company_preset_meta (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL DEFAULT ''
    );
  `);

  // Fields that belong to a preset: all keys in the 'company' category.
  // Derived dynamically so adding a new company-category field to DEFAULTS
  // automatically becomes part of both presets.
  const companyFieldRows = db.prepare(
    "SELECT key, value FROM company_settings WHERE category = 'company'"
  ).all();
  const companyFieldKeys = companyFieldRows.map(r => r.key);

  // Hardcoded ASP defaults — used when seeding the ASP preset for the
  // first time. ISP preset is seeded from the CURRENT company_settings
  // values (i.e., whatever's live = already dad's ISP data).
  const ASP_DEFAULTS = {
    logo:        'ASP',
    trade_name:  'AMAZING SPICE PARK',
    legal_name:  ' PRIVATE LIMITED',
    short_name:  'AMAZING SPICE PARK PRIVATE LIMITED',
    pan:         'ABDCA2636B',
    cin:         'U46305KL2025PTC095544',
    fssai:       '',
    sbl:         '',
  };

  const insertPreset = db.prepare(
    'INSERT OR IGNORE INTO company_presets (preset_code, field_key, field_value) VALUES (?, ?, ?)'
  );
  const seedPresets = db.transaction(() => {
    // ISP preset: seeded from current live values so first-time upgrade
    // preserves whatever the user had configured before presets existed.
    for (const r of companyFieldRows) {
      insertPreset.run('ISP', r.key, r.value);
    }
    // ASP preset: seeded from hardcoded defaults; user edits afterward.
    for (const k of companyFieldKeys) {
      const v = (ASP_DEFAULTS[k] !== undefined) ? ASP_DEFAULTS[k] : '';
      insertPreset.run('ASP', k, v);
    }
  });
  seedPresets();

  // Meta: active_preset_code defaults to 'ISP' on first install so
  // nothing visibly changes until the user flips it.
  db.prepare(
    "INSERT OR IGNORE INTO company_preset_meta (key, value) VALUES ('active_preset_code', 'ISP')"
  ).run();

  // Heal-step (one-off): if the active preset is currently ASP AND the
  // previous preset-overlay architecture had mirrored ASP values into
  // company_settings, those company_settings values are now incorrect
  // (they hold ASP identity where PDFs expect ISP identity). Restore
  // the stable ISP values from the ISP preset into company_settings so
  // invoices/purchase PDFs produce correct output again.
  //
  // Safe to run every startup: it only rewrites company_settings rows
  // that exist in the ISP preset (the 8 company-category fields), using
  // the ISP preset's values. If you WANT company_settings to hold ASP
  // (e.g. you deliberately renamed the primary company), re-save the
  // ISP preset with your desired values afterward.
  try {
    const ispRows = db.prepare(
      "SELECT field_key, field_value FROM company_presets WHERE preset_code = 'ISP'"
    ).all();
    // Detect corruption: if company_settings.short_name doesn't match
    // ISP preset's short_name, we almost certainly have contamination
    // from the old overlay path. Heal unconditionally for the 8 fields.
    const cur = db.prepare("SELECT value FROM company_settings WHERE key = 'short_name'").get();
    const ispShort = (ispRows.find(r => r.field_key === 'short_name') || {}).field_value;
    if (cur && ispShort && cur.value !== ispShort) {
      const upd = db.prepare('UPDATE company_settings SET value = ? WHERE key = ?');
      for (const r of ispRows) upd.run(r.field_value, r.field_key);
      console.log('[presets] Healed company_settings from ISP preset (previous ASP overlay cleaned up)');
    }
  } catch (e) {
    console.warn('[presets] Heal step skipped:', e.message);
  }
}

function getSettingBool(db, key) {
  const v = getSetting(db, key);
  return v === 'true' || v === '1';
}

function getSettingNum(db, key) {
  return parseFloat(getSetting(db, key)) || 0;
}

function getAllSettings(db) {
  const rows = db.prepare('SELECT key, value, category, label, field_type FROM company_settings ORDER BY rowid').all();
  const grouped = {};
  for (const r of rows) {
    if (!grouped[r.category]) grouped[r.category] = [];
    grouped[r.category].push(r);
  }
  return grouped;
}

function updateSettings(db, settings) {
  const upd = db.prepare('UPDATE company_settings SET value = ? WHERE key = ?');
  const batch = db.transaction((items) => {
    let n = 0;
    for (const [k, v] of Object.entries(items)) { upd.run(String(v), k); n++; }
    return n;
  });
  return batch(settings);
}

function getSettingsFlat(db) {
  const rows = db.prepare('SELECT key, value, field_type FROM company_settings').all();
  const flat = {};
  for (const r of rows) {
    if (r.field_type === 'boolean') flat[r.key] = r.value === 'true';
    else if (r.field_type === 'number') flat[r.key] = parseFloat(r.value) || 0;
    else flat[r.key] = r.value;
  }
  // NOTE: preset values are NOT overlaid here any more. Previously we
  // pushed the active preset's (ISP or ASP) company-category values into
  // this flat object so downstream consumers (invoice PDFs, exports) would
  // transparently see the "current identity". That broke ISP invoices
  // when ASP preset was active — ISP invoices read cfg.short_name and got
  // AMAZING SPICE PARK instead of IDEAL SPICES.
  //
  // Instead: company_settings holds the ISP identity (stable). s_* fields
  // hold ASP identity (stable). effectiveCompany() picks between them
  // based on business_state + business_mode. Presets are edit-time only
  // and stored in company_presets; the UI shows the active preset.
  return flat;
}

// ── Preset helpers ─────────────────────────────────────────────────────
// These are the authoritative accessors for preset data. The UI calls
// these via new API endpoints.

// Defensive: make sure the preset tables and seed rows exist before any
// preset operation runs. Covers the case where the server was running
// when new code was deployed and initCompanySettings() never executed
// the new migration block. Cheap (CREATE TABLE IF NOT EXISTS + SELECT).
function ensurePresetsInitialized(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS company_presets (
      preset_code TEXT NOT NULL,
      field_key   TEXT NOT NULL,
      field_value TEXT NOT NULL DEFAULT '',
      PRIMARY KEY (preset_code, field_key)
    );
    CREATE TABLE IF NOT EXISTS company_preset_meta (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL DEFAULT ''
    );
  `);
  const meta = db.prepare("SELECT value FROM company_preset_meta WHERE key = 'active_preset_code'").get();
  if (!meta) {
    db.prepare("INSERT OR IGNORE INTO company_preset_meta (key, value) VALUES ('active_preset_code', 'ISP')").run();
  }
  // Seed presets if the table is empty (e.g., after this defensive
  // create-table just created it for the first time).
  const count = db.prepare('SELECT COUNT(*) as c FROM company_presets').get();
  if (!count || count.c === 0) {
    const companyFieldRows = db.prepare(
      "SELECT key, value FROM company_settings WHERE category = 'company'"
    ).all();
    const ASP_DEFAULTS = {
      logo:       'ASP',
      trade_name: 'AMAZING SPICE PARK',
      legal_name: ' PRIVATE LIMITED',
      short_name: 'AMAZING SPICE PARK PRIVATE LIMITED',
      pan:        'ABDCA2636B',
      cin:        'U46305KL2025PTC095544',
      fssai:      '',
      sbl:        '',
    };
    const ins = db.prepare('INSERT OR IGNORE INTO company_presets (preset_code, field_key, field_value) VALUES (?, ?, ?)');
    for (const r of companyFieldRows) ins.run('ISP', r.key, r.value);
    for (const r of companyFieldRows) {
      const v = (ASP_DEFAULTS[r.key] !== undefined) ? ASP_DEFAULTS[r.key] : '';
      ins.run('ASP', r.key, v);
    }
  }
}

function getActivePresetCode(db) {
  ensurePresetsInitialized(db);
  const r = db.prepare(
    "SELECT value FROM company_preset_meta WHERE key = 'active_preset_code'"
  ).get();
  return (r && r.value) || 'ISP';
}

function setActivePresetCode(db, code) {
  ensurePresetsInitialized(db);
  if (code !== 'ISP' && code !== 'ASP') throw new Error('Invalid preset code: ' + code);
  // UPSERT the meta row in case the UPDATE affects 0 rows (meta row missing)
  const existing = db.prepare("SELECT value FROM company_preset_meta WHERE key = 'active_preset_code'").get();
  if (existing) {
    db.prepare("UPDATE company_preset_meta SET value = ? WHERE key = 'active_preset_code'").run(code);
  } else {
    db.prepare("INSERT INTO company_preset_meta (key, value) VALUES ('active_preset_code', ?)").run(code);
  }
  // NOTE: Previously we ALSO mirrored the preset's values into
  // company_settings here. Removed — that corrupted ISP invoice output
  // when ASP preset was active (PDFs read short_name/pan/cin and got
  // ASP values even for TN/ISP invoices). Identity fields are now stored
  // stably: ISP in company_settings, ASP in s_* keys.
}

function getPreset(db, code) {
  ensurePresetsInitialized(db);
  const rows = db.prepare(
    'SELECT field_key, field_value FROM company_presets WHERE preset_code = ?'
  ).all(code);
  const obj = {};
  for (const r of rows) obj[r.field_key] = r.field_value;
  return obj;
}

function getAllPresets(db) {
  ensurePresetsInitialized(db);
  return {
    ISP:    getPreset(db, 'ISP'),
    ASP:    getPreset(db, 'ASP'),
    active: getActivePresetCode(db),
  };
}

function savePreset(db, code, values) {
  ensurePresetsInitialized(db);
  if (code !== 'ISP' && code !== 'ASP') throw new Error('Invalid preset code: ' + code);
  const ins = db.prepare(
    'INSERT OR REPLACE INTO company_presets (preset_code, field_key, field_value) VALUES (?, ?, ?)'
  );
  for (const [k, v] of Object.entries(values || {})) {
    ins.run(code, k, String(v ?? ''));
  }
  // If this is the ISP preset, also mirror into company_settings since
  // that's where downstream code (effectiveCompany, invoice PDFs) reads
  // ISP identity from. ASP preset edits never touch company_settings —
  // ASP lives in the s_* keys, which are edited separately via the
  // Sister Company tab.
  if (code === 'ISP') {
    const syncUpd = db.prepare('UPDATE company_settings SET value = ? WHERE key = ?');
    for (const [k, v] of Object.entries(values || {})) {
      syncUpd.run(String(v ?? ''), k);
    }
  }
}

function getGSTRates(db) {
  const g = getSettingNum(db, 'gst_goods');
  return { cgst: g / 2, sgst: g / 2, igst: g, service: getSettingNum(db, 'gst_service'), tcs: getSettingNum(db, 'tcs_tds') };
}

module.exports = { DEFAULTS, CATEGORIES, initCompanySettings, getSetting, getSettingBool, getSettingNum, getAllSettings, updateSettings, getSettingsFlat, getGSTRates, getActivePresetCode, setActivePresetCode, getPreset, getAllPresets, savePreset };
