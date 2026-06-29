/**
 * calculations.js — Core business logic
 * Replaces: GENERATE.PRG, parts of GSTKBILT/GSTKBILP/GSTBILP/PAYCHECK
 */

const { getSettingsFlat, getGSTRates } = require('./company-config');

// Display date formatter for journal/register exports — honours the user's
// Settings → Display → Date format choice via the shared module.
const { fmtDate: _ddmmyyyy } = require('./date-format');

// `round0` = round to integer (whole rupee — used by the Round on/off line).
// Sign-aware (Excel's "round half away from zero").
const round0 = (n) => {
  const x = Number(n);
  if (!isFinite(x)) return 0;
  if (x === 0) return 0;
  return (x < 0 ? -1 : 1) * Math.round(Math.abs(x));
};

// `round2` = round to 2 decimals (paise). Sign-aware (Excel's "round half
// away from zero"). Used by the payment-TDS netting so on-screen, statement
// and bank-file figures all foot to the same paise.
const round2 = (n) => {
  const x = Number(n);
  if (!isFinite(x)) return 0;
  if (x === 0) return 0;
  return (x < 0 ? -1 : 1) * Math.round(Math.abs(x) * 100) / 100;
};

/**
 * Compact a GROUP_CONCAT'd lot-number list into a clean, de-duplicated,
 * sorted comma string (e.g. "12,13,14"). Lots that look numeric sort
 * ascending; mixed/text lot numbers fall back to lexical order. Returns
 * '' when there are no lots so callers can omit the suffix entirely.
 */
function formatLotList(raw) {
  if (!raw) return '';
  const uniq = [...new Set(
    String(raw).split(',').map(s => s.trim()).filter(Boolean)
  )];
  const allNumeric = uniq.every(x => /^\d+$/.test(x));
  uniq.sort(allNumeric ? (a, b) => Number(a) - Number(b) : undefined);
  return uniq.join(',');
}

/**
 * Extract the 2-digit state code from a seller's `cr` field.
 *
 * The `cr` column historically stored "GSTIN.<15-char-gstin>" (where the
 * "GSTIN." prefix was added by the UI for sellers). Some import paths
 * (Excel seller import, GST portal lookup, edge-case manual entry) end up
 * with a bare GSTIN without the prefix. Both forms should be supported,
 * because the state code only depends on the first 2 chars of the GSTIN
 * itself, not on the prefix.
 *
 * Logic:
 *   "GSTIN.32AAACG1234F1Z2" → strip "GSTIN." prefix → first 2 chars → "32"
 *   "32AAACG1234F1Z2"       → already bare         → first 2 chars → "32"
 *   ""  / null / undefined  → ""                                    (no GSTIN)
 *   "CR.001" or other       → ""                                    (not a GSTIN)
 *
 * Only a value that looks like a valid 15-char GSTIN starting with 2 digits
 * yields a state code; anything else returns "" (which means no intra-match,
 * so IGST applies — safe default for non-GSTIN sellers).
 */
function gstinStateCode(cr) {
  if (!cr) return '';
  let s = String(cr).trim().toUpperCase();
  if (s.startsWith('GSTIN.')) s = s.substring(6);
  else if (s.startsWith('GSTIN')) s = s.substring(5);  // e.g. "GSTIN<no-dot>33AAA..."
  // GSTIN format: 2 digits + 5 letters + 4 digits + 1 letter + 1 digit + Z + 1 alphanumeric
  // We only need the first 2 chars, but verify they're digits.
  if (s.length < 2) return '';
  const head = s.substring(0, 2);
  if (!/^\d{2}$/.test(head)) return '';
  return head;
}

/**
 * Calculate purchase amounts for a lot (after trade)
 * This is what GENERATE.PRG does — fills pqty, prate, puramt, com, gst, etc.
 */
function calculateLot(lot, cfg) {
  const result = { ...lot };
  const gstGoods = cfg.gst_goods || 5;
  const cgstRate = gstGoods / 2;
  const sgstRate = gstGoods / 2;
  const igstRate = gstGoods;
  
  // Purchase qty = qty + Sample Refund (from Rates & Charges settings, global)
  // Falls back to per-lot `refud` column for back-compat with older data.
  const sampleRefundKg = (cfg.refund != null && cfg.refund !== '') ? Number(cfg.refund) : (lot.refud || 0);
  result.pqty = (lot.qty || 0) + (Number(sampleRefundKg) || 0);

  // flag_discount_in_prate roll-in applies ONLY to Grade 1 lots: when ON,
  // deduction1_inclusive is used for the ISP pooler rate (instead of
  // deduction1) and the per-lot Discount is forced to 0 below because it's
  // baked into the rate. Grade 2 / ungraded lots keep the normal deduction +
  // discount behaviour regardless of the flag.
  const lotGradeStr = String(lot.grade || '').trim();
  const discInPrateFlag = cfg.flag_discount_in_prate === true
    || String(cfg.flag_discount_in_prate || '').toLowerCase() === 'true';
  const applyRollIn = discInPrateFlag && lotGradeStr === '1';

  // ── Compute planter-side numbers for BOTH ISP and ASP views ────────
  // Refer to lot fields directly; doesn't depend on cfg.business_state.
  // Returns { pqty, prate, puramt } — the legacy active-view trio.
  // Used both to populate the active-view columns (legacy compat) and
  // the new isp_*/asp_* columns (Option B dual-storage).
  function planterCalc(stateName) {
    const isKerala = stateName === 'KERALA';
    const pqty = (lot.qty || 0) + (Number(sampleRefundKg) || 0);
    const gradeStr = String(lot.grade || '').trim();
    let prate, puramt;
    if (cfg.business_mode === 'e-Auction') {
      // e-Auction: rate = lot.price; the purchase TAXABLE VALUE = Amount +
      // Refund (Amount = qty × price, Refund = price × SB Sample Refund).
      // Commission, handling and their GST are billed SEPARATELY on the
      // Commission Bill and are NOT netted into the goods value here — so the
      // purchase invoice / Tally voucher tax base is the gross goods value.
      // These numbers are state-agnostic by design; for dual-view storage both
      // ISP and ASP get the same values. If rules ever diverge per state in
      // e-Auction mode, fork here.
      const sbRefund  = Number(cfg.sb_refund) || 0;
      const refundAmt = Math.round((lot.price || 0) * sbRefund * 100) / 100;
      prate = lot.price || 0;
      puramt = Math.round(((lot.amount || 0) + refundAmt) * 100) / 100;
    } else {
      // e-Trade: deduction-based, with different deduction sources per state.
      //   ISP (TN)  → cfg.deduction1 / cfg.deduction2
      //   ASP (KL)  → cfg.isp_profit_pooler / cfg.isp_profit_dealer
      //              (margin ISP earns on the ASP→ISP internal transfer)
      let deduction;
      if (isKerala) {
        deduction = (gradeStr === '2')
          ? Number(cfg.isp_profit_dealer || 0)
          : Number(cfg.isp_profit_pooler || 0);
      } else {
        if (gradeStr === '2') {
          deduction = Number(cfg.deduction2 || 0);
        } else if (applyRollIn) {
          // Grade 1 + roll-in: use the discount-inclusive pooler deduction.
          deduction = Number(cfg.deduction1_inclusive || 0);
        } else {
          deduction = Number(cfg.deduction1 || 0);
        }
      }
      const rawRate = (lot.price || 0) * (1 - deduction / 100);
      prate = Math.round(rawRate);
      // PurAmt formula:
      //   ASP (Kerala) → Qty × P_Rate    (sample refund EXCLUDED — inter-company transfer)
      //   ISP (TN)     → P_Qty × P_Rate  (sample refund INCLUDED — direct purchase)
      const puramtQty = isKerala ? (lot.qty || 0) : pqty;
      puramt = Math.round(puramtQty * prate * 100) / 100;
    }
    return { pqty, prate, puramt };
  }

  // Compute both views and stash them in the result for callers to persist.
  // The dual-storage columns let the URD voucher (and future ISP-only or
  // ASP-only reports) read the right values without flipping business_state.
  const ispView = planterCalc('TAMIL NADU');
  const aspView = planterCalc('KERALA');
  result.isp_pqty = ispView.pqty;
  result.isp_prate = ispView.prate;
  result.isp_puramt = ispView.puramt;
  result.asp_pqty = aspView.pqty;
  result.asp_prate = aspView.prate;
  result.asp_puramt = aspView.puramt;

  // Active-view legacy fields: pqty/prate/puramt mirror whichever set
  // matches the current cfg.business_state, so existing readers (Lots tab,
  // reports, e-Trade calculations downstream) keep working unchanged.
  if (cfg.business_mode === 'e-Auction') {
    // e-Auction: same values for both views currently
    const sbRefund  = Number(cfg.sb_refund) || 0;
    const commPct   = Number(cfg.commission) || 0;
    const handlingPct = Number(cfg.hpc) || 0;
    const refundAmt = Math.round((lot.price || 0) * sbRefund * 100) / 100;
    const commAmt   = Math.round((((lot.amount || 0) + refundAmt) * commPct / 100) * 100) / 100;
    const handling  = Math.round(commAmt * handlingPct / 100 * 100) / 100;
    result.refund = refundAmt;
    result.com    = commAmt;
    result.sertax = handling;
    result.prate  = ispView.prate;       // same as aspView.prate in e-Auction
    result.puramt = ispView.puramt;
  } else {
    // e-Trade: pick whichever view matches active business_state
    const isKerala = (String(cfg.business_state || '').toUpperCase() === 'KERALA');
    const active = isKerala ? aspView : ispView;
    result.prate = active.prate;
    result.puramt = active.puramt;
    result.com = 0;
    result.sertax = 0;
  }

  // Intra/inter-state detection (shared between both modes):
  //   Seller's GSTIN state code (first 2 digits) vs company's state code.
  //   Uses gstinStateCode() to handle both "GSTIN.<gstin>" and bare "<gstin>".
  const sellerGstState = gstinStateCode(lot.cr);
  const companyGstState = cfg.business_state === 'KERALA' ? '32' : '33';
  const isIntra = (sellerGstState === companyGstState);

  // GST rate: both modes tax a SERVICE component (commission/handling for e-Auction,
  // trade-credit discount for e-Trade), so both use gst_service rate.
  const gstServiceRate = Number(cfg.gst_service) || 18;
  const halfRate = gstServiceRate / 2;

  // Default CGST/SGST/IGST to 0; mode-specific branches fill them in below.
  result.cgst = 0;
  result.sgst = 0;
  result.igst = 0;

  if (cfg.business_mode === 'e-Auction') {
    // e-Auction: GST on (Commission + Handling) using Service Rate
    const gstBase = result.com + result.sertax;
    if (isIntra) {
      result.cgst = Math.round(gstBase * halfRate / 100 * 100) / 100;
      result.sgst = Math.round(gstBase * halfRate / 100 * 100) / 100;
    } else {
      result.igst = Math.round(gstBase * gstServiceRate / 100 * 100) / 100;
    }

    // Advance (sum of deductions) — kept for downstream compatibility (PDFs, exports)
    result.advance = result.com + result.sertax + result.cgst + result.sgst + result.igst;

    // Payable = (Amount + Refund) − (Commission + Handling + CGST + SGST + IGST)
    result.balance = Math.round(((lot.amount || 0) + result.refund - result.advance) * 100) / 100;

  } else {
    // e-Trade: compute Discount first, then GST on the Discount (if flag is ON).
    //
    //   Discount = round(PurAmt / 1000 * days * discount%)  — nearest rupee, half-up
    //   Payable  = PurAmt − Discount − CGST − SGST − IGST
    //
    // result.refund is reused as the Discount amount for e-Trade mode.
    // Days source depends on seller type:
    //   GSTIN present (registered dealer) → cfg.dealer_days
    //   no GSTIN (CR / agriculturist)     → cfg.discount_days
    // SKIPPED entirely when the roll-in is in effect for THIS lot (flag ON
    // AND Grade 1) — the discount is already baked into P_Rate via
    // deduction1_inclusive, so a separate refund would double-count it. GST
    // on Discount is skipped for the same reason.
    const sellerHasGstin = sellerGstState !== '';
    if (applyRollIn) {
      result.refund = 0;
    } else {
      const days    = sellerHasGstin
        ? (Number(cfg.dealer_days) || 0)
        : (Number(cfg.discount_days) || 0);
      const discPct = Number(cfg.discount_pct)  || 0;
      result.refund = Math.round((result.puramt / 1000) * days * discPct);

      // Only apply GST on the Discount when flag_disc_gst is ON.
      // Uses Service Rate because the discount is treated as a credit/finance service.
      if (cfg.flag_disc_gst && result.refund > 0) {
        if (isIntra) {
          result.cgst = Math.round(result.refund * halfRate / 100 * 100) / 100;
          result.sgst = Math.round(result.refund * halfRate / 100 * 100) / 100;
        } else {
          result.igst = Math.round(result.refund * gstServiceRate / 100 * 100) / 100;
        }
      }
    }

    // e-Trade has no commission/handling, so advance = GST only (informational).
    result.advance = result.cgst + result.sgst + result.igst;

    // Payable = PurAmt − Discount − GST-on-Discount
    const totalDeductions = result.refund + result.cgst + result.sgst + result.igst;
    result.balance = Math.round((result.puramt - totalDeductions) * 100) / 100;
  }

  // Bill amount (for agriculturist bills) — always equals PurAmt
  result.bilamt = result.puramt;

  return result;
}

/**
 * Calculate TDS under Section 194Q (TDS on Purchase of Goods).
 *
 * Per Section 194Q, a buyer whose turnover exceeds ₹10 cr in the prior FY
 * must deduct TDS at 0.1% of the purchase amount in EXCESS of ₹50 lakh
 * paid to a single seller in the current FY. Once the threshold is crossed,
 * TDS applies to every subsequent rupee bought from that seller for the
 * rest of the year.
 *
 * Inputs (must all be on the SAME basis — either all incl-GST or all
 *   excl-GST; the caller is responsible for keeping units consistent):
 *   purchaseAmount  — current trade's purchase amount (this voucher)
 *   priorPurchases  — sum of all prior purchases from the same seller in
 *                      the current FY (excluding this trade)
 *   cfg.tds_threshold     — usually 5000000 (₹50 L); configurable
 *   cfg.tds_purchase_rate — usually 0.1 (%); configurable, fall-back to
 *                            tcs_tds for back-compat with older configs
 *
 * Returns: TDS amount in rupees (rounded up to nearest paisa).
 */
function calculateTDS(purchaseAmount, priorPurchases, cfg) {
  const threshold = Number(cfg.tds_threshold) || 5000000;
  // Prefer the dedicated TDS-on-purchase rate; fall back to the legacy
  // shared tcs_tds setting if the new key isn't configured yet.
  const tdsRate = Number(cfg.tds_purchase_rate) || Number(cfg.tcs_tds) || 0.1;

  if (priorPurchases > threshold) {
    // Already crossed threshold this FY — TDS on the full new purchase
    return Math.ceil(purchaseAmount * tdsRate / 100);
  } else if ((priorPurchases + purchaseAmount) > threshold) {
    // This trade crosses the threshold — TDS only on the portion above
    const excess = priorPurchases + purchaseAmount - threshold;
    return Math.ceil(excess * tdsRate / 100);
  }
  return 0;
}

/**
 * Calculate TCS under Section 206C(1H) — TCS on Sale of Goods.
 * Threshold logic mirrors TDS-on-purchase: TCS applies to amounts in
 * EXCESS of ₹50 lakh per buyer per FY, then to every subsequent rupee.
 */
function calculateTCS(invoiceAmount, priorSales, cfg) {
  const threshold = Number(cfg.tds_threshold) || 5000000;
  const tcsRate = Number(cfg.tcs_tds) || 0.1;

  if (priorSales > threshold) {
    return Math.ceil(invoiceAmount * tcsRate / 100);
  } else if ((priorSales + invoiceAmount) > threshold) {
    const excess = priorSales + invoiceAmount - threshold;
    return Math.ceil(excess * tcsRate / 100);
  }
  return 0;
}

/**
 * Build sales invoice data for a buyer
 * Aggregates lots by buyer for a given auction
 * Sale type filter is optional — if lots don't have sale set yet, filter by buyer only
 */
function buildSalesInvoice(db, auctionId, buyerCode, saleType, cfg, opts = {}) {
  // ASP context = Kerala + e-Trade (sister-company ASP billing). Every ASP
  // sales invoice is an inter-state transfer regardless of the buyer's GST
  // state. We don't mutate `saleType` before the lots query (that would
  // wrongly exclude lots already assigned a different type); instead the
  // effective sale type is forced to 'I' AFTER the query — see
  // effectiveSaleType below. It drives IGST (not CGST/SGST), hides
  // Transport/Insurance, and — because the generate endpoints persist
  // invoice.saleType — stamps the stored invoice's `sale` column as 'I' to
  // match the PDF (which already prints 'I' for ASP).
  const isASP = (String(cfg.business_mode || '').toLowerCase() === 'e-trade')
             && (String(cfg.business_state || '').toUpperCase() === 'KERALA');
  // In e-Auction, inter-state invoices DO bill Transport/Insurance (unlike
  // ISP e-Trade interstate, where the buyer covers freight). Used below to
  // skip the hideTI suppression for auction inter-state sales.
  const isEAuction = (String(cfg.business_mode || '').toLowerCase() === 'e-auction');

  // Get all lots for this buyer in this auction that have amounts.
  // Don't filter by sale — we're ASSIGNING the sale type now.
  // Skip code='WD' lots: those are withdrawn (no actual buyer
  // transaction), so they must not appear as line items, contribute
  // to totals, or get an invo stamped on them.
  // ASP sales invoices select lots by auction_id + buyer ONLY — the sale
  // filter is dropped. Why: ASP generation deliberately never stamps
  // lots.sale (it's left free for the later ISP step — see the generate
  // route), and ASP invoices imported from old data carry a sale value that
  // doesn't line up with the lots. Filtering on sale in those cases excludes
  // every lot, so buildSalesInvoice returns null and the PDF falls back to a
  // single consolidated line with no lot numbers. `opts.aspInvoice` is set by
  // the reprint routes from the stored invoice's KERALA state; `isASP` covers
  // live generation while in ASP (Kerala + e-Trade) context.
  const aspLotPick = isASP || opts.aspInvoice === true;
  let lotSql = `SELECT * FROM lots WHERE auction_id = ? AND buyer = ? AND amount > 0
                AND UPPER(COALESCE(code, '')) != 'WD'`;
  const lotParams = [auctionId, buyerCode];
  if (!aspLotPick) {
    // ISP invoices keep the sale filter so a buyer's lots split across
    // different sale types don't bleed across invoices.
    lotSql += ` AND (sale IS NULL OR sale = '' OR sale = ?)`;
    lotParams.push(saleType);
  }
  lotSql += ` ORDER BY lot_no`;
  const lots = db.all(lotSql, lotParams);
  
  if (!lots.length) return null;

  const gstGoods = cfg.gst_goods || 5;
  const companyState = cfg.business_state === 'KERALA' ? '32' : '33';
  
  // Get buyer details
  const buyer = db.get('SELECT * FROM buyers WHERE buyer = ?', [buyerCode]);
  const buyerState = buyer ? buyer.gstin.substring(0, 2) : companyState;
  // Effective sale type: ASP forces inter-state ('I'); ISP keeps the caller's
  // type. Used for the GST split, transport/insurance suppression, and the
  // returned/stored sale type.
  const effectiveSaleType = isASP ? 'I' : saleType;
  // Inter-state vs intra-state drives the GST split (IGST vs CGST+SGST).
  // The explicit sale type is the legal source of truth — the operator
  // picks L/I/E per invoice and the sales-list IGST/CGST columns already
  // follow it — so honour it directly here so the PDF matches the list.
  // 'I'/'E' → inter-state (IGST); 'L' → intra-state (CGST+SGST). Fall back
  // to comparing the buyer's GSTIN state with the company's home state only
  // when no sale type was supplied. ASP transfers are always inter-state.
  const _est = String(effectiveSaleType || '').toUpperCase();
  const isInterState = isASP ? true
    : (_est === 'I' || _est === 'E') ? true
    : (_est === 'L') ? false
    : (buyerState !== companyState);

  let totalQty = 0, totalBags = 0, totalAmount = 0;
  const lineItems = [];

  // For ASP invoices (Kerala + e-Trade), invoice values are based on the
  // intra-company transfer price (PurAmt / P_Rate), not the external auction
  // price. Compute both sets of numbers per lot so the PDF can show the right
  // ones AND the totals/GST align with what's printed. `isASP` is computed at
  // the top of the function (it also forces saleType to 'I').
  // Purchase view = the ISPL-side print of an ASP sale (Sales-invoice
  // screen → "Print Purchase Invoice"). It bills the goods using ASP's
  // planter numbers — Qty (raw lot qty, sample refund excluded), P_Rate
  // (asp_prate) and PurAmt (asp_puramt). Because ASP's PurAmt is
  // qty × asp_prate, the printed Qty stays lot.qty so the line's
  // Qty × Rate matches PurAmt.
  const purchaseView = opts.purchaseView === true;
  for (const lot of lots) {
    totalBags += lot.bags;
    // Run calculateLot to derive prate/puramt (uses isp_profit_pooler/dealer
    // for ASP, deduction1/deduction2 for ISP). Doing this here keeps the
    // calculation logic in one place. It also exposes the asp_* view used
    // by the purchase print.
    const calc = calculateLot(lot, cfg);
    const prate  = purchaseView ? calc.asp_prate  : calc.prate;
    const puramt = purchaseView ? calc.asp_puramt : calc.puramt;

    // Totals depend on which view is billing:
    //   ASP / purchase → total qty = lot.qty (no sample refund), total = Σ puramt
    //   ISP (sales)    → total qty = lot.qty, total = Σ amount   (unchanged)
    totalQty += lot.qty;
    totalAmount += (isASP || purchaseView) ? puramt : lot.amount;

    lineItems.push({
      lot: lot.lot_no, grade: lot.grade, bags: lot.bags, qty: lot.qty,
      price: lot.price, amount: lot.amount,
      // Extra fields used by ASP / purchase invoice rendering. Qty column
      // falls back to lot.qty (no pqty set), keeping Qty × P_Rate == PurAmt.
      prate: prate, puramt: puramt,
    });
  }

  // Gunny cost (HSN: jute bags). Always billed for every sale type and
  // business mode, INCLUDING e-Auction LOCAL sales — those carry a Gunny
  // line even though their Transport/Insurance is still suppressed (see
  // hideTI below). Gunny feeds the taxable value, GST and the PDF row.
  const gunnyCost = totalBags * (cfg.gunny_rate || 165);

  // Transport & Insurance rates depend on sale type:
  //   L (Local)        → local_transport / local_insurance
  //   I (Inter-state)  → transport / insurance
  //   E (Export)       → use inter-state rates (same interstate logistics)
  // All rates are in ₹/kg (transport) or paise-per-thousand-units (insurance).
  // IMPORTANT: Use `??` not `||` so that an explicit 0 from settings is
  // respected (e.g. user wants no transport charge). `||` would treat 0 as
  // falsy and fall back to the default, silently adding unwanted charges.
  const pickRate = (...vals) => {
    for (const v of vals) {
      if (v === undefined || v === null || v === '') continue;
      const n = typeof v === 'number' ? v : parseFloat(v);
      if (!Number.isNaN(n)) return n;
    }
    return 0;
  };
  const isLocal = (effectiveSaleType === 'L');
  // e-Auction bills Transport/Insurance at the LOCAL rates regardless of
  // sale type — interstate auction sales should carry the same charges as
  // local ones (the dedicated interstate rates aren't maintained in auction
  // mode). So treat e-Auction like local for rate selection.
  const useLocalRates = isLocal || isEAuction;
  // ASP invoices (Kerala + e-Trade, already computed in `isASP` above) do NOT
  // bill Transport/Insurance as separate line-items — only Cardamom + Gunny.
  // Force both to zero so subtotal, GST, and grand total agree with the PDF.
  // ISP inter-state invoices ('I') don't bill transport/insurance separately
  // (the buyer covers freight). Match the rendering rule that hides these
  // rows from the PDF — see invoice-pdf.js hideTransportInsurance.
  // Transport/Insurance suppression:
  //   - ISP e-Trade INTER-state ('I'): buyer covers freight (original rule).
  //   - e-Auction LOCAL ('L'): local auction invoices bill Cardamom + Gunny
  //     but no Transport/Insurance. e-Auction interstate KEEPS them.
  // ASP suppression is handled separately via `isASP` in the rate calc below.
  const hideTI = (!isASP && !isEAuction && _est === 'I')
              || (isEAuction && _est === 'L');
  // Per-invoice Transport & Insurance switch (opts.includeTI). Defaults ON
  // so existing/generated invoices are unchanged. When the operator turns
  // it OFF for an invoice, force both charges to 0 — same effect as hideTI,
  // but driven by the stored per-invoice flag so it persists across reprints
  // and flows into the Tally export (which reads the stored zeros).
  const includeTI = opts.includeTI !== false;
  // Per-component enable flags from Rates & Charges. When the matching flag
  // is OFF the rate is forced to 0, so the component drops out of the taxable
  // value, GST, and PDF entirely. Blank/legacy values default to ON.
  const flagOn = (k, defaultOn) => {
    const v = cfg[k];
    if (v === undefined || v === null || v === '') return defaultOn;
    return v === true || String(v).toLowerCase() === 'true';
  };
  const useLocalTransport = flagOn('flag_local_transport', true);
  const useLocalInsurance = flagOn('flag_local_insurance', true);
  const useInterTransport = flagOn('flag_inter_transport', true);
  const useInterInsurance = flagOn('flag_inter_insurance', true);
  const transportRate = isASP || hideTI || !includeTI ? 0 : (useLocalRates
    ? (useLocalTransport ? pickRate(cfg.local_transport, cfg.transport, 2.5) : 0)
    : (useInterTransport ? pickRate(cfg.transport, 2.5) : 0));
  const insuranceRate = isASP || hideTI || !includeTI ? 0 : (useLocalRates
    ? (useLocalInsurance ? pickRate(cfg.local_insurance, cfg.insurance, 0.75) : 0)
    : (useInterInsurance ? pickRate(cfg.insurance, 0.75) : 0));

  const transportCost = Math.round(totalQty * transportRate * 100) / 100;

  // Insurance formula (per spec):
  //   insurance = ((cardamom_amount + gunny_cost) + GST on cardamom+gunny) / 1000 × insurance_rate
  const subtotalGoods = totalAmount + gunnyCost;
  const gstOnGoods = subtotalGoods * gstGoods / 100;
  const insuranceCost = Math.round((subtotalGoods + gstOnGoods) / 1000 * insuranceRate * 100) / 100;

  // Taxable value = cardamom + gunny + transport + insurance
  const taxableValue = subtotalGoods + transportCost + insuranceCost;

  // All four components get the SAME gstGoods rate (per user confirmation).
  let cgst = 0, sgst = 0, igst = 0;
  if (isInterState) {
    igst = Math.round(taxableValue * gstGoods / 100 * 100) / 100;
  } else {
    cgst = Math.round(taxableValue * (gstGoods / 2) / 100 * 100) / 100;
    sgst = Math.round(taxableValue * (gstGoods / 2) / 100 * 100) / 100;
  }

  const totalBeforeRound = taxableValue + cgst + sgst + igst;
  const subtotalRounded = Math.round(totalBeforeRound);
  const roundDiff = subtotalRounded - totalBeforeRound;

  // Additional Charge — sum(cardamom) × cfg.addl_charge_value % .
  // The configured value is a PERCENTAGE (e.g. 2 means 2%). Sits BELOW the
  // Round line and adds straight onto the grand total — does not feed into
  // GST or round-off math. When the percentage is 0 the charge is fully
  // skipped (no row, no effect on grand total).
  const addlChargePct = Number(cfg.addl_charge_value) || 0;
  const addlCharge = addlChargePct > 0
    ? Math.round(totalAmount * addlChargePct / 100 * 100) / 100
    : 0;
  const addlChargeName = addlCharge > 0 ? String(cfg.addl_charge_name || '').trim() : '';
  const grandTotal = addlCharge > 0
    ? Math.round((subtotalRounded + addlCharge) * 100) / 100
    : subtotalRounded;

  return {
    buyer: buyer || {},
    saleType: effectiveSaleType,
    lineItems,
    summary: {
      totalQty, totalBags, totalAmount,
      gunnyCost, transportCost, insuranceCost,
      taxableValue, cgst, sgst, igst,
      roundDiff, grandTotal,
      addlCharge, addlChargeName,
      isInterState
    }
  };
}

/**
 * Build purchase invoice data for a seller
 * Aggregates lots by seller for a given auction (registered dealers only)
 */
function buildPurchaseInvoice(db, auctionId, sellerName, cfg, opts = {}) {
  // ispView: print the ISP planter figures (isp_pqty/isp_prate/isp_puramt)
  // on the purchase invoice regardless of the active business state. The
  // legacy pqty/prate/puramt fields mirror whichever state is active, so
  // in Kerala (ASP) mode they hold ASP figures — but the downloaded
  // registered-dealer purchase invoice should always reflect ISP. Set by
  // the purchase-PDF download routes only; the stored-generate / TDS path
  // leaves it off and keeps its existing behaviour.
  const ispView = !!opts.ispView;
  // A lot qualifies for a Purchase Invoice if it has a GSTIN-bearing seller —
  // i.e. cr is either "GSTIN.<15-char>" (legacy UI format) or a bare 15-char
  // GSTIN starting with 2 digits (Excel import format). We accept both.
  const lots = db.all(
    `SELECT * FROM lots
     WHERE auction_id = ? AND name = ? AND amount > 0
       AND (UPPER(cr) LIKE 'GSTIN%' OR cr GLOB '[0-9][0-9]*')
     ORDER BY lot_no`,
    [auctionId, sellerName]
  );
  
  if (!lots.length) return null;

  const gstGoods = cfg.gst_goods || 5;
  const companyState = cfg.business_state === 'KERALA' ? '32' : '33';

  let totalQty = 0, totalPuramt = 0, totalBags = 0;
  const lineItems = [];

  for (const lot of lots) {
    const sellerState = gstinStateCode(lot.cr);
    const isInter = sellerState !== companyState;
    // In ISP-view, source the ISP planter trio; otherwise the legacy
    // active-view fields. Fall back to the legacy values when the ISP
    // columns are blank (e-Auction, or lots predating the dual-view
    // backfill — where isp_* == legacy anyway). In non-ISP mode these
    // resolve to exactly the previous values, so that path is unchanged.
    const prate  = ispView ? (lot.isp_prate  || lot.prate  || 0)           : (lot.prate  || 0);
    const puramt = ispView ? (lot.isp_puramt || lot.puramt || 0)           : (lot.puramt || 0);
    const pqty   = ispView ? (lot.isp_pqty   || lot.pqty   || lot.qty || 0) : (lot.pqty   || lot.qty || 0);

    const rcgst = isInter ? 0 : Math.round(puramt * (gstGoods / 2) / 100 * 100) / 100;
    const rsgst = isInter ? 0 : Math.round(puramt * (gstGoods / 2) / 100 * 100) / 100;
    const rigst = isInter ? Math.round(puramt * gstGoods / 100 * 100) / 100 : 0;

    totalQty += pqty;
    totalPuramt += puramt;
    totalBags += lot.bags || 0;

    lineItems.push({
      lot: lot.lot_no, bags: lot.bags, grade: lot.grade,
      qty: lot.qty, pqty: pqty,
      price: lot.price, prate: prate,
      amount: lot.amount, puramt,
      com: lot.com, sertax: lot.sertax,
      cgst: rcgst, sgst: rsgst, igst: rigst
    });
  }

  const firstLot = lots[0];
  const sellerState = gstinStateCode(firstLot.cr);
  const isInter = sellerState !== companyState;

  // GST is charged once on the whole taxable value (single round) — the
  // GST-standard method and exactly what Tally recomputes when it imports
  // the voucher. Summing the per-lot rounded GST instead drifts by a
  // paisa or two on multi-lot invoices, which made Tally flag a "tax
  // amount mismatch" on import (cleared only after a refresh/recompute).
  const half = gstGoods / 2;
  const totalCgst = isInter ? 0 : Math.round(totalPuramt * half / 100 * 100) / 100;
  const totalSgst = isInter ? 0 : Math.round(totalPuramt * half / 100 * 100) / 100;
  const totalIgst = isInter ? Math.round(totalPuramt * gstGoods / 100 * 100) / 100 : 0;
  // Reconcile the per-lot line GST so it sums EXACTLY to the aggregate
  // (push the rounding remainder onto the last line). The purchase-invoice
  // PDF totals the per-line GST column, so this keeps the printed lines,
  // the invoice total, and Tally all in agreement.
  if (lineItems.length) {
    const last = lineItems[lineItems.length - 1];
    const sumc = lineItems.reduce((s, li) => s + (li.cgst || 0), 0);
    const sums = lineItems.reduce((s, li) => s + (li.sgst || 0), 0);
    const sumi = lineItems.reduce((s, li) => s + (li.igst || 0), 0);
    last.cgst = Math.round((last.cgst + (totalCgst - sumc)) * 100) / 100;
    last.sgst = Math.round((last.sgst + (totalSgst - sums)) * 100) / 100;
    last.igst = Math.round((last.igst + (totalIgst - sumi)) * 100) / 100;
  }

  const totalBeforeRound = totalPuramt + totalCgst + totalSgst + totalIgst;
  const roundDiff = Math.round(totalBeforeRound) - totalBeforeRound;
  const grandTotal = Math.round(totalBeforeRound);

  // ── TDS calculation (Section 194Q) ──
  //
  // 1) GSTIN format compatibility: the purchases table may have rows with
  //    gstin in either form ("GSTIN.32AAA..." or bare "32AAA..."). We
  //    derive both candidates from the current lot's cr and match either.
  //
  // 2) Amount basis must match: this trade's amount and the running prior
  //    total must be on the SAME basis (both with-GST or both excl-GST),
  //    otherwise the threshold check is inconsistent. The `purchases.total`
  //    column = puramt + GST = grand total (with GST). So:
  //      • flag_wgst=true  → prior=SUM(total), current=grandTotal       ✓
  //      • flag_wgst=false → prior=SUM(amount), current=totalPuramt    ✓
  //    (`purchases.amount` is stored as the pre-GST puramt subtotal.)
  const cr = String(firstLot.cr || '').trim();
  const gstinPrefixed = cr.toUpperCase().startsWith('GSTIN.') ? cr : ('GSTIN.' + cr);
  const gstinBare     = cr.toUpperCase().startsWith('GSTIN.') ? cr.substring(6) : cr;
  const priorAmountCol = cfg.flag_wgst ? 'total' : 'amount';
  const priorPurchases = db.get(
    `SELECT COALESCE(SUM(${priorAmountCol}),0) as total
       FROM purchases
      WHERE (gstin = ? OR gstin = ?) AND date >= ?`,
    [gstinPrefixed, gstinBare, cfg.season_start || '2026-04-01']
  );
  const computedTds = cfg.flag_tds_purchase
    ? calculateTDS(cfg.flag_wgst ? grandTotal : totalPuramt, priorPurchases ? priorPurchases.total : 0, cfg)
    : 0;
  // When reprinting / exporting an already-recorded purchase
  // (opts.useStoredTds), honour the value stored on the purchase row
  // rather than re-deriving 194Q. This keeps imported purchases — saved
  // with tds = 0 — from showing a freshly-computed TDS on the invoice /
  // Tally voucher that they were never recorded with. Fresh generation
  // leaves the flag off, so it still computes (and the generate route
  // then stores) the 194Q amount as before.
  let tdsAmount = computedTds;
  if (opts.useStoredTds) {
    const storedPur = db.get(
      `SELECT tds FROM purchases WHERE auction_id = ? AND LOWER(TRIM(name)) = LOWER(TRIM(?)) ORDER BY id DESC LIMIT 1`,
      [auctionId, sellerName]
    );
    if (storedPur) tdsAmount = Number(storedPur.tds) || 0;
  }
  const invoiceAmount = grandTotal - tdsAmount;

  return {
    seller: { name: firstLot.name, address: firstLot.padd, place: firstLot.ppla, 
              cr: firstLot.cr, pan: firstLot.pan, state: firstLot.pstate },
    lineItems,
    summary: {
      totalQty, totalBags, totalPuramt, totalCgst, totalSgst, totalIgst,
      roundDiff, grandTotal, tdsAmount, invoiceAmount, isInter
    }
  };
}

/**
 * ISP-view discount for a single lot — the e-Trade discount calculateLot
 * would compute, but on the ISP planter PurAmt (isp_puramt) so Payments
 * shows the ISP discount even when the lot was calculated in the ASP
 * company. Mirrors the calculateLot formula exactly (and the SQL in
 * getPaymentSummary). e-Auction keeps the stored `advance`.
 * Falls back to the active `puramt` when isp_puramt isn't populated.
 */
function ispLotDiscount(lot, cfg) {
  const mode = (cfg && cfg.business_mode || 'e-Trade').toLowerCase();
  // Non-e-Trade: mirror getPaymentSummary's stored discount column exactly
  // (advance for 'auction', refund otherwise) so the modal/statement stay
  // in step with the screen. ISP P_Qty/P_Rate don't diverge in these modes.
  if (mode !== 'e-trade') {
    const col = (mode === 'auction') ? 'advance' : 'refund';
    return Number(lot[col]) || 0;
  }
  const grade = String(lot.grade || '').trim();
  const rollIn = (cfg.flag_discount_in_prate === true
      || String(cfg.flag_discount_in_prate || '').toLowerCase() === 'true') && grade === '1';
  if (rollIn) return 0;
  const hasGstin = gstinStateCode(lot.cr) !== '';
  const days = hasGstin ? (Number(cfg.dealer_days) || 0) : (Number(cfg.discount_days) || 0);
  const discPct = Number(cfg.discount_pct) || 0;
  const puramt = (Number(lot.isp_puramt) > 0) ? Number(lot.isp_puramt) : (Number(lot.puramt) || 0);
  return Math.round((puramt / 1000) * days * discPct);
}

/**
 * Generate payment summary for sellers (PAYCHECK.PRG equivalent)
 */
/**
 * Build a per-seller TDS context for a trade, sourced from the seller's
 * stamped Section-194Q purchase-invoice TDS (purchases.tds). Returns a
 * helper that, given a seller name and the puramt of a lot subset, returns
 * the proportionate TDS to withhold:
 *   - paying the whole seller  → the seller's full purchase TDS
 *   - paying a state/lot subset → TDS spread ∝ puramt of that subset
 * TDS is 0 for unregistered/agriculturist sellers (no purchase invoice) and
 * for dealers below the ₹50-lakh threshold (purchases.tds stamped as 0).
 */
function paymentTdsContext(db, auctionId) {
  const tdsByName = {};     // seller name (UPPER/trim) → stamped purchase TDS
  const puramtByName = {};  // seller name → full payable-lot puramt (denominator)
  try {
    for (const r of db.all(
      `SELECT name, COALESCE(SUM(tds),0) AS tds
         FROM purchases WHERE auction_id = ? GROUP BY name`, [auctionId])) {
      // Accumulate (+=) rather than assign: the same seller can appear under
      // more than one casing (a known name-drift data state), which GROUP BY
      // keeps as separate rows — collapsing them by upper-cased key must SUM,
      // not overwrite, or one casing would zero out the other's TDS.
      const key = String(r.name || '').trim().toUpperCase();
      tdsByName[key] = (tdsByName[key] || 0) + (Number(r.tds) || 0);
    }
  } catch (_) { /* purchases table may be absent on partial migrations */ }
  try {
    for (const r of db.all(
      `SELECT name, COALESCE(SUM(puramt),0) AS puramt
         FROM lots WHERE auction_id = ? AND amount > 0 GROUP BY name`, [auctionId])) {
      const key = String(r.name || '').trim().toUpperCase();
      puramtByName[key] = (puramtByName[key] || 0) + (Number(r.puramt) || 0);
    }
  } catch (_) { /* lots always present, but stay defensive */ }
  return {
    tdsByName,
    puramtByName,
    share(name, puramt) {
      const key = String(name || '').trim().toUpperCase();
      const tds = tdsByName[key] || 0;
      if (!(tds > 0)) return 0;
      const full = puramtByName[key] || 0;
      if (!(full > 0)) return 0;
      const frac = Math.min(1, Math.max(0, (Number(puramt) || 0) / full));
      return round2(tds * frac);
    },
  };
}

/**
 * Distribute a set of fractional payables to whole rupees so the lines stay
 * whole integers AND foot exactly to round0(Σ). Largest fractional parts get
 * the +1 leftover. Used when invoice rounding (flag_round) is on so per-lot
 * Payable rows in the modal/statement add up to the rounded seller total.
 */
function distributeRoundedPayable(values) {
  const nums = (values || []).map(v => Number(v) || 0);
  if (!nums.length) return [];
  const target = round0(nums.reduce((a, b) => a + b, 0));
  const out = nums.map(v => Math.floor(v));
  let leftover = target - out.reduce((a, b) => a + b, 0);
  if (leftover > 0) {
    const order = nums
      .map((v, i) => ({ i, frac: v - Math.floor(v) }))
      .sort((a, b) => b.frac - a.frac);
    for (let k = 0; k < order.length && leftover > 0; k++) { out[order[k].i] += 1; leftover--; }
  }
  return out;
}

function getPaymentSummary(db, auctionId, state, cfg) {
  // The "discount" column is the sum of two parts per seller per auction:
  //   1. Per-lot computed discount (lots.refund in e-Trade, lots.advance in
  //      auction mode) — based on discount_pct × days × puramt
  //   2. Per-seller debit notes for this auction — manual adjustments
  //      (e.g., quality complaints, settlement deductions). Joined by
  //      seller name + auction ano so we sum all debit_notes that apply.
  // Total payable already accounts for these via balance recalc, but the
  // displayed "Discount" column needs the COMBINED figure so the user
  // sees both the policy discount and any manual adjustments.
  const mode = (cfg && cfg.business_mode || 'e-Trade').toLowerCase();
  const discountCol = (mode === 'auction') ? 'advance' : 'refund';
  // Payments are purchase-side: show the ISP planter trio (P_Qty / P_Rate /
  // PurAmt) consistently — total_qty = Σ isp_pqty and total_amount = Σ
  // isp_puramt — so the screen, the lots modal and the printed statement
  // all agree. Falls back to the active-view columns for e-Auction / older
  // lots where isp_* wasn't backfilled (isp_puramt = 0).
  //
  // Discount: e-Trade shows the ISP discount, derived from isp_puramt with
  // the same formula calculateLot uses (round(puramt/1000 × days × pct),
  // days = dealer_days for GSTIN sellers else discount_days, 0 on the
  // Grade-1 roll-in). This matches ispLotDiscount() — kept in lockstep so
  // the screen, modal (via /api/lots), and statement agree. e-Auction
  // keeps the stored advance/refund.
  const useIspDisc = (mode === 'e-trade');
  const rollFlag   = (cfg && (cfg.flag_discount_in_prate === true || String(cfg.flag_discount_in_prate || '').toLowerCase() === 'true')) ? 1 : 0;
  const dealerDays = Number(cfg && cfg.dealer_days)   || 0;
  const discDays   = Number(cfg && cfg.discount_days) || 0;
  const discPct    = Number(cfg && cfg.discount_pct)  || 0;
  const ispDiscSql = `SUM(CASE WHEN ? = 1 AND TRIM(COALESCE(l.grade,'')) = '1' THEN 0
      ELSE ROUND( (CASE WHEN l.isp_puramt > 0 THEN l.isp_puramt ELSE l.puramt END) / 1000.0
           * (CASE WHEN (UPPER(COALESCE(l.cr,'')) LIKE 'GSTIN%' OR l.cr GLOB '[0-9][0-9]*') THEN ? ELSE ? END)
           * ? ) END)`;
  const discSql = useIspDisc ? ispDiscSql : `SUM(l.${discountCol})`;
  let query = `SELECT l.name, l.cr,
    SUM(CASE WHEN l.isp_puramt > 0 THEN l.isp_pqty   ELSE l.pqty   END) as total_qty,
    SUM(CASE WHEN l.isp_puramt > 0 THEN l.isp_puramt ELSE l.puramt END) as total_amount,
    SUM(CASE WHEN l.isp_puramt > 0 THEN l.isp_pqty   ELSE l.pqty   END) as total_pqty,
    SUM(CASE WHEN l.isp_puramt > 0 THEN l.isp_prate  ELSE l.prate  END) as avg_prate,
    SUM(CASE WHEN l.isp_puramt > 0 THEN l.isp_puramt ELSE l.puramt END) as total_puramt,
    ${discSql} as lot_discount,
    SUM(l.balance) as total_payable,
    COUNT(*) as lot_count
    FROM lots l WHERE l.auction_id = ? AND l.amount > 0`;
  // discSql params (if any) appear in the SELECT, so they precede auctionId.
  const params = useIspDisc ? [rollFlag, dealerDays, discDays, discPct, auctionId] : [auctionId];
  if (state) { query += ' AND l.state = ?'; params.push(state); }
  query += ' GROUP BY l.name, l.cr ORDER BY l.state, l.name';
  const sellers = db.all(query, params);

  // Fetch this auction's identifier (ano) so we can match debit_notes.
  // Debit notes are keyed by ano + seller name (no FK to auctions.id),
  // mirroring the legacy FoxPro flow.
  const auction = db.get('SELECT ano FROM auctions WHERE id = ?', [auctionId]);
  const ano = auction ? auction.ano : null;
  // Build a name → debit_note total map for fast lookup
  const debitMap = {};
  if (ano) {
    const debits = db.all(
      'SELECT name, SUM(amount) as total FROM debit_notes WHERE ano = ? GROUP BY name',
      [ano]
    );
    for (const d of debits) debitMap[d.name] = Number(d.total) || 0;
  }
  // Payment TDS mirrors the stamped purchase-invoice TDS exactly (see
  // paymentTdsContext): the seller's full TDS when paying the whole seller,
  // spread ∝ puramt when a state filter narrows the lot set. 0 for
  // unregistered/agriculturist sellers and dealers below the threshold.
  const tdsCtx = paymentTdsContext(db, auctionId);
  // Merge: total_discount = lot-policy discount + any manual debit notes
  return sellers.map(s => {
    const lotDisc = Number(s.lot_discount) || 0;
    const manualDisc = Number(debitMap[s.name]) || 0;
    // Pre-TDS net = balance − manual debit notes (balance already nets the
    // lot-policy discount and GST). This is the "Total" column shown before
    // TDS in the Payments views.
    const totalBeforeTds = (Number(s.total_payable) || 0) - manualDisc;
    const tds = tdsCtx.share(s.name, s.total_puramt);
    return {
      ...s,
      total_discount: lotDisc + manualDisc,
      // "Total" column — pre-TDS net (PurAmt − Discount − GST).
      total_total: totalBeforeTds,
      // TDS = the seller's stamped Section 194Q purchase-invoice TDS for this
      // trade (0 until that invoice is generated, or below threshold).
      total_tds: tds,
      // "Payable" column — Total − TDS. This is what the seller is actually
      // paid (TDS withheld), so the bank file nets it too (see
      // getBankPaymentData).
      total_payable: totalBeforeTds - tds,
    };
  });
}

/**
 * Generate bank payment data (BANKPAY.PRG — RTGS/NEFT format).
 * Used by both the "after discount" Bank Payment export (default) and
 * the "Bank Payment (Before)" export when `opts.before === true`.
 */
function getBankPaymentData(db, auctionId, cfg, opts) {
  opts = opts || {};
  const useBefore = !!opts.before;
  // Bank Payment lists every seller in the trade with a non-zero
  // payable (or non-zero pre-discount amount in 'before' mode) — both
  // registered dealers AND unregistered (URD/agriculturist) farmers.
  // The earlier WHERE clause filtered to URD-only by excluding rows
  // whose `cr` looked like a GSTIN. That came from the legacy FoxPro
  // BANKPAY.PRG which only handled farmers — but the e-Trade flow pays
  // every seller via RTGS/NEFT, so all sellers must be included.
  // Result was: registered dealers had IFSC + acctnum on file, but the
  // SQL excluded them and returned empty rows, so the export was blank.
  //
  // Bank details come from `traders` (single-bank legacy) or
  // `trader_banks` (multi-bank). The LEFT JOIN to traders pulls
  // address/IFSC; we then COALESCE with trader_banks default for
  // sellers who maintain multiple bank accounts.
  const payments = db.all(
    // GROUP BY l.name (only) — same fix as getPaymentSummary. Splitting
    // by `cr` produced duplicate bank-payment rows whenever a seller's
    // lots held inconsistent GSTIN values, leading to NEFT files with
    // the dealer listed twice for partial amounts.
    // JOIN trader by lots.trader_id (FK), not by name. Joining by name
    // multiplied each lot row by the number of traders sharing that
    // name (multi-branch sellers / accidental dupes), then SUM(puramt)
    // etc. summed those duplicates → inflated payable. GROUP BY name
    // alone wasn't enough; the fan-out happened BEFORE the aggregate.
    `SELECT MAX(l.state) AS state, l.name, MAX(l.cr) AS cr,
      SUM(l.puramt) as puramt, SUM(l.refund) as advance, SUM(l.balance) as payable,
      GROUP_CONCAT(l.lot_no) as lot_nos,
      MAX(t.id) AS trader_id,
      MAX(t.ifsc) AS t_ifsc, MAX(t.acctnum) AS t_acctnum, MAX(t.holder_name) AS t_holder,
      MAX(t.padd) AS padd, MAX(t.ppla) AS ppla, MAX(t.pin) AS pin,
      -- Per-lot bank routing: distinct non-null bank_ids across this
      -- seller's payable lots, plus counts so we can tell whether they
      -- ALL share one account (single → use it) or differ (mixed → keep
      -- the default account and flag multipleBanks for the UI).
      GROUP_CONCAT(DISTINCT l.bank_id) AS bank_ids,
      COUNT(*) AS lot_count,
      COUNT(l.bank_id) AS bank_lot_count
    FROM lots l
    LEFT JOIN traders t ON t.id = l.trader_id
    WHERE l.auction_id = ? AND l.amount > 0
      AND (l.paid IS NULL OR l.paid = '')
    GROUP BY l.name
    ORDER BY MAX(l.state), l.name`,
    [auctionId]
  );

  // Per-seller bank-details fallback chain:
  //   1. trader_banks default (is_default=1) — picks the explicitly
  //      flagged primary account when the seller has multiple banks
  //   2. trader_banks first row — when no default flagged
  //   3. traders.ifsc/acctnum — legacy single-bank
  // Pre-fetch all default banks once (cheaper than per-seller query).
  const bankByTraderId = {};
  // Also index every bank row by its own id so per-lot bank_id routing can
  // resolve the exact account a seller's lots were tagged with.
  const bankById = {};
  try {
    const banks = db.all(`
      SELECT trader_id, ifsc, acctnum, holder_name, bank_name, is_default, id
        FROM trader_banks
       ORDER BY trader_id, is_default DESC, id ASC
    `);
    for (const b of banks) {
      // First row per trader_id wins (already sorted by is_default DESC).
      if (bankByTraderId[b.trader_id] == null) bankByTraderId[b.trader_id] = b;
      bankById[b.id] = b;
    }
  } catch (_) { /* trader_banks may not exist on partial migrations */ }

  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]);
  const roundAmounts = cfg.flag_round;
  // Same stamped-purchase-TDS source the Payments tab uses, so the NEFT
  // amount matches the screen exactly (Payable = Total − TDS).
  const tdsCtx = paymentTdsContext(db, auctionId);

  return payments.map(p => {
    // 'before' uses puramt — pre-discount, useful when paying suppliers
    // before the deduction policy is applied. 'after' (default) uses
    // payable = puramt − discount − GST, then nets the seller's purchase
    // TDS (Section 194Q) so the credited amount matches the Payments tab.
    const tds = useBefore ? 0 : tdsCtx.share(p.name, p.puramt);
    const rawAmount = (useBefore ? (p.puramt || 0) : (p.payable || 0)) - tds;
    const amount = roundAmounts ? round0(rawAmount) : rawAmount;
    const tb = p.trader_id != null ? bankByTraderId[p.trader_id] : null;
    // Per-lot bank routing. Distinct non-null bank_ids tagged on this
    // seller's payable lots:
    const distinctBankIds = String(p.bank_ids || '')
      .split(',').map(s => s.trim()).filter(s => s !== '' && s !== 'null')
      .map(Number).filter(Number.isFinite);
    const hasUntagged = Number(p.lot_count || 0) > Number(p.bank_lot_count || 0);
    // Use the lot-tagged account ONLY when every payable lot points at the
    // same single account (no untagged lots). Otherwise keep the seller's
    // default account and flag `multipleBanks` so the UI can warn the user
    // to export each account's lots separately via the lot picker.
    const lotBank = (distinctBankIds.length === 1 && !hasUntagged)
      ? bankById[distinctBankIds[0]] : null;
    const multipleBanks = distinctBankIds.length > 1
      || (distinctBankIds.length >= 1 && hasUntagged);
    const ifsc      = (lotBank && lotBank.ifsc)        || (tb && tb.ifsc)        || p.t_ifsc    || '';
    const acctnum   = (lotBank && lotBank.acctnum)     || (tb && tb.acctnum)     || p.t_acctnum || '';
    const holderNm  = (lotBank && lotBank.holder_name) || (tb && tb.holder_name) || p.t_holder  || p.name;
    // Lots this seller's payment covers — surfaced in REMARKS so the
    // beneficiary can reconcile the credit against the specific lots.
    const lots = formatLotList(p.lot_nos);
    return {
      // Seller name preserved on the row so callers can filter by the
      // same key the Payments tab UI uses (the ticked checkbox value).
      // beneficiaryName below can diverge from this (it tracks the bank
      // account holder, which may be a different person/entity) so it's
      // not safe to filter against beneficiaryName.
      name: p.name,
      transactionType: rawAmount >= 200000 ? 'RTGS' : 'NEFT',
      ifsc,
      accountNo: acctnum,
      beneficiaryName: holderNm,
      address1: p.padd || '',
      address2: p.ppla || '',
      pin: p.pin || '',
      amount,
      lots,
      remarks: `${auction ? auction.ano : ''} ${p.name} PAYMENT ${rawAmount.toFixed(2)} Credited${lots ? ` for lot${lots.includes(',') ? 's' : ''} ${lots}` : ''}`,
      holderName: holderNm,
      // True when this seller's lots point at more than one bank account
      // (or a mix of tagged + untagged). The row still pays a single
      // account (the default); the Payments UI shows a badge prompting the
      // user to export each account's lots separately via the lot picker.
      multipleBanks,
    };
  });
}

/**
 * TDS return data (TDSRETU.PRG equivalent)
 */
function getTDSReturnData(db, fromDate, toDate, orderBy) {
  const order = orderBy === 'party' ? 'name' : 'date, invo';
  // PAN extraction. The gstin column holds either:
  //   "GSTIN.32AAHCE4551A1Z8" (21 chars, with prefix — most common)
  //   "32AAHCE4551A1Z8"       (15 chars, bare GSTIN)
  // Strip the optional "GSTIN." prefix first, then take chars 3-12 of
  // the bare GSTIN to get the 10-char PAN ("AAHCE4551A").
  return db.all(
    `SELECT invo as invoice, date, name,
      SUBSTR(
        CASE WHEN UPPER(SUBSTR(COALESCE(gstin,''), 1, 6)) = 'GSTIN.'
             THEN SUBSTR(gstin, 7)
             ELSE COALESCE(gstin,'') END,
        3, 10
      ) as pan,
      amount as assess_value, tds
    FROM purchases
    WHERE date BETWEEN ? AND ? AND tds > 0
    ORDER BY ${order}`,
    [fromDate, toDate]
  );
}

/**
 * Build Agriculturist Bill of Supply (GSTKBILP/GSTBILP equivalent)
 * For sellers WITHOUT GSTIN — agricultural produce from farmers.
 * No GST charged (exempt/reverse charge).
 * 
 * Returns: { seller, lineItems, summary } if successful
 *          { error, detail } object if no data (to help debug)
 */
function buildAgriBill(db, auctionId, sellerName, cfg) {
  const trimmedName = String(sellerName || '').trim();
  if (!trimmedName) return { error: 'Seller name is empty' };

  // First check: any lots at all for this seller (case-insensitive)?
  const allLots = db.all(
    `SELECT * FROM lots WHERE auction_id = ? AND UPPER(TRIM(name)) = UPPER(?) ORDER BY lot_no`,
    [auctionId, trimmedName]
  );
  
  if (!allLots.length) {
    return { error: `No lots found for seller "${trimmedName}" in this auction. Check the exact spelling.` };
  }

  // Check if any have GSTIN — those aren't eligible for Bills of Supply
  const withGstin = allLots.filter(l => l.cr && l.cr.toUpperCase().startsWith('GSTIN'));
  const withoutGstin = allLots.filter(l => !l.cr || !l.cr.toUpperCase().startsWith('GSTIN'));
  
  if (withGstin.length && !withoutGstin.length) {
    return { error: `Seller "${trimmedName}" has GSTIN (${withGstin[0].cr}). Use Generate Purchase Invoice instead — Bills of Supply are only for agriculturists without GSTIN.` };
  }

  // Filter to agri-eligible lots with amount > 0
  const lots = withoutGstin.filter(l => (l.amount || 0) > 0);
  
  if (!lots.length) {
    if (withoutGstin.length) {
      return { error: `Seller "${trimmedName}" has ${withoutGstin.length} lot(s) but none have amount > 0. Set prices on the lots first (or click Calculate All).` };
    }
    return { error: `No eligible lots for "${trimmedName}"` };
  }

  let totalQty = 0, totalPuramt = 0;
  const lineItems = [];

  // A Bill of Supply documents the ISP company's purchase FROM the
  // agriculturist, so in e-Trade it must always use the ISP-view planter
  // numbers (isp_pqty / isp_prate / isp_puramt) — NOT the active-state view,
  // which would show ASP internal-transfer figures when business_state is
  // KERALA. Fall back to the legacy active columns when the isp_* dual-storage
  // columns weren't backfilled (isp_puramt = 0). In e-Auction the isp_/asp_
  // values are identical to the active ones, so this is a no-op there.
  const isETrade = String(cfg.business_mode || '').toLowerCase() === 'e-trade';

  for (const lot of lots) {
    const useIsp = isETrade && Number(lot.isp_puramt) > 0;
    const pqty   = useIsp ? lot.isp_pqty   : lot.pqty;
    const prate  = useIsp ? lot.isp_prate  : lot.prate;
    const puramt = useIsp ? lot.isp_puramt : lot.puramt;
    totalQty += pqty || lot.qty;
    totalPuramt += puramt || 0;
    lineItems.push({
      lot: lot.lot_no, qty: lot.qty, pqty: pqty,
      price: lot.price, prate: prate,
      amount: lot.amount, puramt: puramt,
      com: lot.com, sertax: lot.sertax
    });
  }

  const firstLot = lots[0];
  const roundDiff = cfg.flag_round ? Math.round(totalPuramt) - totalPuramt : 0;
  const netAmount = Math.round(totalPuramt);

  return {
    seller: {
      name: firstLot.name,
      address: firstLot.padd,
      place: firstLot.ppla,
      pin: firstLot.ppin,
      state: firstLot.pstate,
      st_code: firstLot.pst_code,
      cr: firstLot.cr,
      pan: firstLot.pan,
      aadhar: firstLot.aadhar,
      tel: firstLot.tel,
    },
    lineItems,
    summary: {
      totalQty, totalPuramt, 
      roundDiff, netAmount,
      cgst: 0, sgst: 0, igst: 0,
      tax: 0
    }
  };
}

/**
 * List agri-eligible sellers for an auction
 * (sellers without GSTIN who have lots with amount > 0)
 */
function listAgriSellers(db, auctionId) {
  // An "agri seller" is one without a GSTIN. Reject both prefixed
  // ("GSTIN.<gstin>") and bare ("<gstin>") forms — anything else (empty,
  // CR codes, plain text) qualifies.
  return db.all(
    `SELECT name, COUNT(*) as lot_count, SUM(qty) as total_qty, SUM(amount) as total_amount
     FROM lots 
     WHERE auction_id = ? 
       AND (cr IS NULL OR cr = ''
            OR (UPPER(cr) NOT LIKE 'GSTIN%' AND cr NOT GLOB '[0-9][0-9]*'))
       AND amount > 0
     GROUP BY name
     ORDER BY name`,
    [auctionId]
  );
}

/**
 * Sales Journal (JOUR.PRG)
 * Date-wise sales invoice register
 */
function getSalesJournal(db, fromDate, toDate, saleType) {
  let query = `SELECT date, sale, invo, buyer, buyer1, gstin, place,
      bag, qty, amount as cardamom, gunny, pava_hc as transport, ins as insurance,
      cgst, sgst, igst, tcs, rund, tot as total
    FROM invoices WHERE date BETWEEN ? AND ?`;
  const params = [fromDate, toDate];
  if (saleType) { query += ' AND sale = ?'; params.push(saleType); }
  query += ' ORDER BY date, sale, invo';
  return db.all(query, params);
}

/**
 * Purchase Journal (PUJOUR.PRG / PPUJOUR.PRG)
 * Date-wise purchase invoice register
 * type: 'dealer' (registered) or 'agri' (agriculturist bills)
 */
function getPurchaseJournal(db, fromDate, toDate, type) {
  if (type === 'agri') {
    return db.all(
      `SELECT date, bil as bill_no, name, add_line as address, pla as place, pstate as state,
        crr as cr, pan, qty, cost, igst, net
      FROM bills WHERE date BETWEEN ? AND ? ORDER BY date, bil`,
      [fromDate, toDate]
    );
  }
  // Dealer purchases
  return db.all(
    `SELECT date, invo as invoice_no, name, add_line as address, place, state,
      gstin, qty, amount, cgst, sgst, igst, rund, total, tds
    FROM purchases WHERE date BETWEEN ? AND ? ORDER BY date, invo`,
    [fromDate, toDate]
  );
}

/**
 * Debit Note calculation
 * For discounts or adjustments against invoices
 */
function buildDebitNote(db, invoiceNo, saleType, discount, cfg) {
  const inv = db.get('SELECT * FROM invoices WHERE invo = ? AND sale = ?', [String(invoiceNo), saleType]);
  if (!inv) return null;

  const gstGoods = cfg.gst_goods || 5;
  const isInter = inv.igst > 0;

  const amount = Math.round(discount * 100) / 100;
  let cgst = 0, sgst = 0, igst = 0;
  
  if (cfg.flag_disc_gst) {
    // Discount amount includes GST — extract it
    const factor = 100 / (100 + gstGoods);
    const taxable = amount * factor;
    if (isInter) igst = Math.round((amount - taxable) * 100) / 100;
    else { 
      const tax = (amount - taxable) / 2;
      cgst = Math.round(tax * 100) / 100;
      sgst = Math.round(tax * 100) / 100;
    }
  } else {
    // Discount is pre-tax — add GST on top
    if (isInter) igst = Math.round(amount * gstGoods / 100 * 100) / 100;
    else {
      cgst = Math.round(amount * (gstGoods / 2) / 100 * 100) / 100;
      sgst = Math.round(amount * (gstGoods / 2) / 100 * 100) / 100;
    }
  }
  
  const total = amount + cgst + sgst + igst;
  return { invoice: inv, amount, cgst, sgst, igst, total };
}

/**
 * Purchase Register (lot-wise)
 * One row PER LOT — the seller-side purchase detail. Unlike the Purchase
 * Journal (one row per dealer invoice / agri bill), this is the raw lot
 * ledger: STATE, TNO, DATE, LOT, BRANCH, NAME, PLACE, GSTIN, BAG, QTY,
 * PRICE, AMOUNT, PQTY, PRATE, PURAMT, DISCOUNT, GST5, PAYABLE.
 *
 * DISCOUNT = refund, GST5 = stored GST-on-discount (`advance`), PAYABLE =
 * balance (GST already netted) — see [[payment-field-semantics]]. In
 * auction mode `advance` is the discount, so GST5 → 0.
 *
 * Withdrawn lots (code = 'WD') ARE included even though withdrawal zeroes
 * their price/amount, so the register accounts for every lot in the trade —
 * they appear with their real BAG/QTY but zero money columns (the only
 * zero-AMOUNT rows here, since unsold lots with no code stay excluded).
 *
 * Scope: a specific trade (opts.auctionId) OR a date range across trades
 * (opts.from/opts.to over the auction date). Trade wins when both given.
 */
function getPurchaseRegister(db, opts = {}) {
  const mode = String(opts.mode || 'e-Trade').toLowerCase();
  const discountCol = (mode === 'auction') ? 'advance' : 'refund';
  const gstCol = (mode === 'auction') ? '0' : 'advance';
  let q = `SELECT l.state AS state, a.ano AS tno, a.date AS date, l.lot_no AS lot,
      l.branch AS branch, l.name AS name, l.ppla AS place, l.cr AS gstin,
      l.bags AS bag, l.qty AS qty, l.price AS price, l.amount AS amount,
      l.pqty AS pqty, l.prate AS prate, l.puramt AS puramt,
      l.${discountCol} AS discount, l.${gstCol} AS gst5, l.balance AS payable
    FROM lots l JOIN auctions a ON a.id = l.auction_id
    WHERE (l.amount > 0 OR UPPER(TRIM(COALESCE(l.code,''))) = 'WD')`;
  const params = [];
  if (opts.auctionId) { q += ' AND l.auction_id = ?'; params.push(opts.auctionId); }
  else if (opts.from && opts.to) { q += ' AND a.date BETWEEN ? AND ?'; params.push(opts.from, opts.to); }
  q += ' ORDER BY l.state, a.ano, CAST(l.lot_no AS INTEGER), l.lot_no';
  const rows = db.all(q, params);
  return rows.map(r => ({ ...r, date: _ddmmyyyy(r.date) }));
}

/**
 * Sales Register (invoice-wise)
 * One row PER INVOICE: STATE, TNO, DATE, SALE, INVO, TRADERNAME, BIDDER,
 * BAG, QTY, AMOUNT, LORRY, GUNNY, IGST, CGST, SGST, INS, INVAMT.
 * LORRY = freight charge (pava_hc); INVAMT = invoice grand total (tot).
 *
 * Scope: a specific trade (matched by auction_id OR ano for legacy rows)
 * OR a date range across trades. Optional saleType filter.
 */
function getSalesRegister(db, opts = {}) {
  let q = `SELECT i.state AS state, i.ano AS tno, i.date AS date, i.sale AS sale,
      i.invo AS invo, i.buyer1 AS tradername, i.buyer AS bidder,
      i.bag AS bag, i.qty AS qty, i.amount AS amount,
      i.pava_hc AS lorry, i.gunny AS gunny, i.igst AS igst, i.cgst AS cgst,
      i.sgst AS sgst, i.ins AS ins, i.tot AS invamt
    FROM invoices i`;
  const params = [];
  const where = [];
  if (opts.auctionId) {
    const a = db.get('SELECT id, ano FROM auctions WHERE id = ?', [opts.auctionId]);
    if (a) { where.push('(i.auction_id = ? OR i.ano = ?)'); params.push(a.id, a.ano); }
    else { where.push('i.auction_id = ?'); params.push(opts.auctionId); }
  } else if (opts.from && opts.to) {
    where.push('i.date BETWEEN ? AND ?'); params.push(opts.from, opts.to);
  }
  if (opts.saleType) { where.push('i.sale = ?'); params.push(opts.saleType); }
  if (where.length) q += ' WHERE ' + where.join(' AND ');
  q += ' ORDER BY i.state, i.ano, i.date, i.sale, i.invo';
  const rows = db.all(q, params);
  return rows.map(r => ({ ...r, date: _ddmmyyyy(r.date) }));
}

// ═══════════════════════════════════════════════════════════════
// PER-PARTY "INDIVIDUAL" REGISTERS (cross-auction, date-range)
// ───────────────────────────────────────────────────────────────
// Three party-statement registers that span MULTIPLE trades within a
// date range (unlike the lot/invoice registers above which are per-trade
// OR a flat date-range list). Each returns { kind, parties: [...] } where
// every party carries its own rows + summary totals so the export layer
// can render one section/page per party (with an optional single-party
// filter). Rows are returned pre-sorted by party so callers can group in
// one pass.
//   • Pooler   — the seller's own lots (lots table). Sold = amount>0.
//   • Seller   — purchase invoices raised TO the pooler (purchases table),
//                summarised per trade. INVO = count of invoices in the trade.
//   • Merchant — sales invoices raised to the buyer (invoices table), one
//                row per invoice. RECEIPT has no data source yet (no
//                receipts table) so it renders 0 / blank; closing balance
//                therefore equals the invoice total.
function _groupRegister(rows, summaryFn) {
  const parties = [];
  let cur = null;
  for (const r of rows) {
    const name = r.party || '';
    if (!cur || cur.name !== name) {
      cur = { name, gstin: '', rows: [] };
      parties.push(cur);
    }
    if (!cur.gstin && r.gstin) cur.gstin = String(r.gstin).trim();
    // The party + gstin live on the group, not on each row.
    const { party, gstin, ...rest } = r;
    cur.rows.push(rest);
  }
  for (const p of parties) p.summary = summaryFn(p.rows);
  return parties;
}
const _num = (v) => Number(v) || 0;
const _sum = (rows, k) => rows.reduce((s, r) => s + _num(r[k]), 0);

// Pooler Register — one row per lot the pooler put up, across all trades
// in range. TNo | Date | Lot | Qty | Rate | Value | PQty | PRate | PurAmt.
// Withdrawn lots (code 'WD') ARE included now (so the register reconciles to
// every lot the pooler handled); the per-party summary breaks the totals into
// Sold (code != WD, value > 0) vs Withdrawn (code = WD; amount is always 0).
function getPoolerRegister(db, opts = {}) {
  let q = `SELECT a.ano AS tno, a.date AS date, l.lot_no AS lot, l.name AS party,
      l.cr AS gstin, l.qty AS qty, l.price AS rate, l.amount AS value,
      l.pqty AS pqty, l.prate AS prate, l.puramt AS puramt,
      UPPER(TRIM(COALESCE(l.code,''))) AS code
    FROM lots l JOIN auctions a ON a.id = l.auction_id
    WHERE 1=1`;
  const params = [];
  if (opts.from && opts.to) { q += ' AND a.date BETWEEN ? AND ?'; params.push(opts.from, opts.to); }
  if (opts.party) { q += ' AND UPPER(TRIM(l.name)) = UPPER(?)'; params.push(String(opts.party).trim()); }
  q += ' ORDER BY l.name, a.date, a.ano, CAST(l.lot_no AS INTEGER), l.lot_no';
  const rows = db.all(q, params).map(r => ({ ...r, date: _ddmmyyyy(r.date) }));
  const isWd = (r) => String(r.code || '').trim().toUpperCase() === 'WD';
  const parties = _groupRegister(rows, (rs) => {
    const qty = _sum(rs, 'qty');
    const value = _sum(rs, 'value');
    const pqty = _sum(rs, 'pqty');
    const puramt = _sum(rs, 'puramt');
    // Sold = not withdrawn and actually fetched a price. Withdrawn = code 'WD'
    // (amount is 0, so withdrawn VALUE is always 0 — only qty is meaningful).
    const soldQty = rs.reduce((s, r) => s + (!isWd(r) && _num(r.value) > 0 ? _num(r.qty) : 0), 0);
    const soldValue = rs.reduce((s, r) => s + (!isWd(r) ? _num(r.value) : 0), 0);
    const withdrawnQty = rs.reduce((s, r) => s + (isWd(r) ? _num(r.qty) : 0), 0);
    const withdrawnValue = 0;
    let notSoldQty = qty - soldQty - withdrawnQty;
    if (Math.abs(notSoldQty) < 1e-6) notSoldQty = 0;   // kill float residue (-0.000)
    return {
      qty, value, pqty, puramt,
      soldQty, soldValue, withdrawnQty, withdrawnValue,
      notSoldQty,
    };
  });
  return { kind: 'pooler', parties };
}

// Seller Register ("SELLERS INDIVIDUAL") — purchase invoices to the pooler,
// summarised per trade. DATE | ANO | INVO(count) | QTY | INVOICE.
function getSellerRegister(db, opts = {}) {
  let q = `SELECT p.name AS party, MAX(p.gstin) AS gstin, p.ano AS ano, p.date AS date,
      COUNT(*) AS invo, SUM(p.qty) AS qty,
      SUM(CASE WHEN COALESCE(p.total,0) > 0 THEN p.total ELSE p.amount END) AS invoice
    FROM purchases p WHERE 1=1`;
  const params = [];
  if (opts.from && opts.to) { q += ' AND p.date BETWEEN ? AND ?'; params.push(opts.from, opts.to); }
  if (opts.party) { q += ' AND UPPER(TRIM(p.name)) = UPPER(?)'; params.push(String(opts.party).trim()); }
  q += ' GROUP BY p.name, p.ano, p.date ORDER BY p.name, p.date, p.ano';
  const rows = db.all(q, params).map(r => ({ ...r, date: _ddmmyyyy(r.date) }));
  const parties = _groupRegister(rows, (rs) => {
    const invoice = _sum(rs, 'invoice');
    return { qty: _sum(rs, 'qty'), invoice, closing: invoice };
  });
  return { kind: 'seller', parties };
}

// Merchant Register ("MERCHANTS INDIVIDUAL") — sales invoices to the buyer,
// one row per invoice. DATE | TNo | INVO | RECP | QTY | INVOICE | RECEIPT.
function getMerchantRegister(db, opts = {}) {
  let q = `SELECT i.buyer1 AS party, i.gstin AS gstin, i.ano AS tno, i.date AS date,
      i.invo AS invo, '' AS recp, i.qty AS qty, i.tot AS invoice, 0 AS receipt
    FROM invoices i WHERE 1=1`;
  const params = [];
  if (opts.from && opts.to) { q += ' AND i.date BETWEEN ? AND ?'; params.push(opts.from, opts.to); }
  if (opts.party) { q += ' AND UPPER(TRIM(i.buyer1)) = UPPER(?)'; params.push(String(opts.party).trim()); }
  q += " ORDER BY i.buyer1, i.date, i.ano, CAST(NULLIF(i.invo,'') AS INTEGER), i.invo";
  const rows = db.all(q, params).map(r => ({ ...r, date: _ddmmyyyy(r.date) }));
  const parties = _groupRegister(rows, (rs) => {
    const invoice = _sum(rs, 'invoice');
    const receipt = _sum(rs, 'receipt');
    return { qty: _sum(rs, 'qty'), invoice, receipt, closing: invoice - receipt };
  });
  return { kind: 'merchant', parties };
}

// Distinct party names for the picker dropdown, scoped to the same source
// table + date range as the matching register.
function listRegisterParties(db, opts = {}) {
  const kind = String(opts.kind || '').toLowerCase();
  const params = [];
  let q;
  if (kind === 'merchant') {
    q = `SELECT DISTINCT i.buyer1 AS name FROM invoices i WHERE COALESCE(i.buyer1,'') != ''`;
    if (opts.from && opts.to) { q += ' AND i.date BETWEEN ? AND ?'; params.push(opts.from, opts.to); }
    q += ' ORDER BY i.buyer1';
  } else if (kind === 'seller') {
    q = `SELECT DISTINCT p.name AS name FROM purchases p WHERE COALESCE(p.name,'') != ''`;
    if (opts.from && opts.to) { q += ' AND p.date BETWEEN ? AND ?'; params.push(opts.from, opts.to); }
    q += ' ORDER BY p.name';
  } else {
    q = `SELECT DISTINCT l.name AS name FROM lots l JOIN auctions a ON a.id = l.auction_id
         WHERE COALESCE(l.name,'') != '' AND UPPER(TRIM(COALESCE(l.code,''))) != 'WD'`;
    if (opts.from && opts.to) { q += ' AND a.date BETWEEN ? AND ?'; params.push(opts.from, opts.to); }
    q += ' ORDER BY l.name';
  }
  return db.all(q, params).map(r => r.name);
}

module.exports = {
  calculateLot,
  calculateTDS,
  calculateTCS,
  buildSalesInvoice,
  buildPurchaseInvoice,
  buildAgriBill,
  buildDebitNote,
  listAgriSellers,
  getPaymentSummary,
  paymentTdsContext,
  distributeRoundedPayable,
  ispLotDiscount,
  getBankPaymentData,
  formatLotList,
  getTDSReturnData,
  getSalesJournal,
  getPurchaseJournal,
  getPurchaseRegister,
  getSalesRegister,
  getPoolerRegister,
  getSellerRegister,
  getMerchantRegister,
  listRegisterParties,
  gstinStateCode,
};
