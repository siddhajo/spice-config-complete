/**
 * exports.js — All export formats
 * Replaces: EXP.PRG (11 types), TALY.PRG, KOTALLY.PRG, BANKPAY export
 */

const ExcelJS = require('exceljs');
const { collectionXlsx: newCollectionXlsx, tradeReportXlsx } = require('./auction-reports');
const { REPORTS: SPICE_BOARD_REPORTS } = require('./spice-board-reports');
const {
  getCompanyHeader, writeXlsxCompanyHeader, xlsxNumFmtForHeader,
} = require('./report-formatters');
const { fmtDate: fmtUserDate } = require('./date-format');
const { getSettingsFlat } = require('./company-config');

// Build an XLSX buffer with a unified brand band on top and Indian-format
// numeric columns. `opts.title` is the report title shown in the middle of
// the band; `opts.metaLines` is an array of right-aligned meta strings
// (e.g. ["Trade #3", "15/04/2026", "ASP"]).
async function createExcelBuffer(sheetName, columns, rows, opts) {
  opts = opts || {};
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName);

  // Apply column widths up front (the brand band uses these widths too).
  ws.columns = columns.map(c => ({ key: c.key, width: c.width || 15 }));

  // Brand band: logo + name + address (left), title (middle), meta (right).
  // `opts.noBrandBand` produces a bare sheet — no logo, no company/title/meta
  // rows — with the column headers on row 1. Used where the caller wants a
  // plain data grid (e.g. e-Trade Tamil Nadu Price List (Before)).
  let startRow;
  if (opts.noBrandBand) {
    startRow = 1;
  } else {
    const header = opts.companyHeader || getCompanyHeader(opts.db);
    startRow = writeXlsxCompanyHeader(wb, ws, header, {
      colCount: columns.length,
      title: opts.title || sheetName,
      metaLines: opts.metaLines || [],
    });
  }

  // Column-header row (right after the brand band, with the spacer row).
  const headerRow = ws.getRow(startRow);
  columns.forEach((c, i) => {
    headerRow.getCell(i + 1).value = c.header;
  });
  headerRow.font = { bold: true, size: 10 };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
  headerRow.eachCell((cell) => {
    cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    cell.alignment = { horizontal: 'center' };
  });

  // Apply Indian-format numFmt to each numeric column. We do this on the
  // worksheet column object so every data row picks it up automatically.
  columns.forEach((c, i) => {
    // `c.text` forces the column to stay textual (e.g. zero-padded lot
    // numbers like "007") — apply Excel's text format and skip numerics.
    if (c.text) {
      ws.getColumn(i + 1).numFmt = '@';
      return;
    }
    const fmt = xlsxNumFmtForHeader(c.header);
    if (fmt) {
      const colObj = ws.getColumn(i + 1);
      colObj.numFmt = fmt;
      colObj.alignment = { horizontal: 'right' };
    }
  });

  // Data rows. addRow uses keys from ws.columns to map object → cells.
  // Rows flagged `_isSubtotal: true` get distinct styling (bold + light
  // yellow fill + thin top border) so callers can interleave per-group
  // subtotal rows directly in `rows` and have them styled automatically.
  const emitDataRow = (rowObj) => {
    const dataRow = ws.addRow({});
    columns.forEach((c, i) => {
      let v = rowObj[c.key];
      // Coerce string-numbers to numbers so Excel applies the numFmt — but
      // never for `text` columns, which must keep their literal string (e.g.
      // a zero-padded "007" lot number would otherwise collapse to 7).
      if (!c.text && typeof v === 'string' && v !== '' && !isNaN(Number(v))) {
        const n = Number(v);
        if (!Number.isNaN(n) && xlsxNumFmtForHeader(c.header)) v = n;
      }
      dataRow.getCell(i + 1).value = v == null ? '' : v;
    });
    if (rowObj && rowObj._isSubtotal) {
      dataRow.font = { bold: true };
      dataRow.eachCell((cell, ci) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF8E1' } };
        cell.border = { top: { style: 'thin' } };
        const fmt = xlsxNumFmtForHeader(columns[ci - 1] && columns[ci - 1].header);
        if (fmt) cell.numFmt = fmt;
      });
    }
  };

  // ── Section-grouped mode (optional) ──
  // When `opts.sections` is provided, we ignore `rows` and emit each
  // section as: section header (merged, light-yellow) → its rows. Reused
  // by the per-party "Individual" registers (one section per party).
  if (Array.isArray(opts.sections) && opts.sections.length) {
    const colLetter = (n) => {
      let s = '';
      while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
      return s;
    };
    opts.sections.forEach((sec, sIdx) => {
      const titleRow = ws.addRow([sec.title || '']);
      ws.mergeCells(`A${titleRow.number}:${colLetter(columns.length)}${titleRow.number}`);
      titleRow.font = { bold: true, size: 10 };
      titleRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
      titleRow.alignment = { horizontal: 'left', vertical: 'middle' };
      (sec.rows || []).forEach(emitDataRow);
      if (opts.spacerBetween && sIdx < opts.sections.length - 1) ws.addRow([]);
    });
  } else {
    // Flat mode — original behaviour.
    rows.forEach(emitDataRow);
  }

  // ── Grand total footer (optional) ──
  // Bold 11pt, yellow `#FFF3CD` fill, double bottom border. Pass values
  // keyed by column key; only the listed columns get numbers, the rest are
  // blank. `label` is placed in the first non-numeric column (defaults to
  // 'GRAND TOTAL').
  if (opts.grandTotal) {
    const gt = opts.grandTotal;
    const fmts = columns.map(c => xlsxNumFmtForHeader(c.header));
    const cells = columns.map(c => (gt.values && gt.values[c.key] != null) ? gt.values[c.key] : '');
    if (gt.label) {
      const labelIdx = fmts.findIndex(f => !f);
      const idx = labelIdx >= 0 ? labelIdx : 0;
      if (cells[idx] === '') cells[idx] = gt.label;
    }
    const gRow = ws.addRow(cells);
    gRow.font = { bold: true, size: 11 };
    gRow.height = 22;
    const fill = gt.fillArgb || 'FFFFF3CD';
    gRow.eachCell((cell, ci) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
      cell.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
      if (fmts[ci - 1]) {
        cell.numFmt = fmts[ci - 1];
        cell.alignment = { horizontal: 'right' };
      }
    });
  }

  return wb.xlsx.writeBuffer();
}

// Build the common XLSX header meta lines for a given auction. Returns
// an array like ["e-TRADE No: 3", "Date: 15/04/2026"]. The crop type
// (ISP/ASP) is omitted — the active preset is already shown via the logo
// and company name in the brand block.
function auctionMeta(db, auctionId) {
  if (!auctionId) return [];
  try {
    const a = db.get(
      'SELECT ano, date, crop_type, mode FROM auctions WHERE id = ?', [auctionId]
    );
    if (!a) return [];
    const dt = String(a.date || '').slice(0, 10).split('-').reverse().join('/');
    const meta = [];
    // Mode-aware: 'e-TRADE No:' or 'e-AUCTION No:' from auction.mode.
    // Empty/legacy mode falls back to e-AUCTION (matches the historical
    // wording exports printed before the mode tag existed).
    const eLbl = (a.mode === 'e-Trade') ? 'e-TRADE' : 'e-AUCTION';
    if (a.ano) meta.push(`${eLbl} No: ${a.ano}`);
    if (dt) meta.push(`Date: ${dt}`);
    return meta;
  } catch (_) { return []; }
}

// ── Export Type 1: Lot Slip (before trade) ───────────────────
async function exportLotSlip(db, auctionId, state) {
  const rows = db.all(
    `SELECT state, lot_no as lot, name, grade, bags as bag, qty, litre
     FROM lots WHERE auction_id = ? ${state ? 'AND state = ?' : ''}
     ORDER BY lot_no`, state ? [auctionId, state] : [auctionId]
  );
  const cols = [
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'LOT', key: 'lot', width: 8 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'GRADE', key: 'grade', width: 8 },
    { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'LITRE', key: 'litre', width: 10 },
  ];
  return createExcelBuffer('LotSlip', cols, rows, {
    db, title: 'Lot Slip', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export Type 2: Lot Slip After Trade (with price/buyer) ───
async function exportLotSlipAfter(db, auctionId, state) {
  const rows = db.all(
    `SELECT state, lot_no as lot, name, bags as bag, qty, price, amount, code
     FROM lots WHERE auction_id = ? ${state ? 'AND state = ?' : ''}
     ORDER BY lot_no`, state ? [auctionId, state] : [auctionId]
  );
  const cols = [
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'LOT', key: 'lot', width: 8 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'PRICE', key: 'price', width: 10 },
    { header: 'AMOUNT', key: 'amount', width: 14 },
    { header: 'CODE', key: 'code', width: 8 },
  ];
  return createExcelBuffer('LotSlipAfter', cols, rows, {
    db, title: 'Lot Slip (After Trade)', metaLines: auctionMeta(db, auctionId),
  });
}

// ── "Before trade" snapshot exports ───────────────────────────
// Surfaced from the Price Import modal: importing prices overwrites the
// lots table with the auctioneer's data, so these reports are the
// operator's record of the pre-trade state. Ported from the eTrade
// build. State → BR is folded inline (KERALA→KL, TAMIL NADU→TN, else
// first two letters) so each function stays self-contained.

// Lot ↔ Buyer crosswalk: which buyer bought each lot, with branch code.
async function exportLotBuyer(db, auctionId) {
  const rows = db.all(
    `SELECT lot_no AS lot, COALESCE(buyer,'') AS buyer,
            CASE UPPER(COALESCE(state,''))
              WHEN 'KERALA' THEN 'KL'
              WHEN 'TAMIL NADU' THEN 'TN'
              ELSE UPPER(SUBSTR(COALESCE(state,''), 1, 2))
            END AS br,
            bags AS bag, qty
     FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]
  );
  const cols = [
    { header: 'LOT',   key: 'lot',   width: 8  },
    { header: 'BUYER', key: 'buyer', width: 24 },
    { header: 'BR',    key: 'br',    width: 6  },
    { header: 'BAG',   key: 'bag',   width: 6  },
    { header: 'QTY',   key: 'qty',   width: 12 },
  ];
  return createExcelBuffer('LotBuyer', cols, rows, {
    db, title: 'Lot Buyer', metaLines: auctionMeta(db, auctionId),
  });
}

// Lot ↔ Seller name crosswalk, with a blank CONTROL column for hand notes.
async function exportLotName(db, auctionId) {
  const rows = db.all(
    `SELECT lot_no AS lot, COALESCE(name,'') AS name,
            CASE UPPER(COALESCE(state,''))
              WHEN 'KERALA' THEN 'KL'
              WHEN 'TAMIL NADU' THEN 'TN'
              ELSE UPPER(SUBSTR(COALESCE(state,''), 1, 2))
            END AS br,
            bags AS bag, qty,
            CASE WHEN COALESCE(price,0) = 0 THEN '' ELSE price END AS price,
            '' AS control
     FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]
  );
  const cols = [
    { header: 'LOT',     key: 'lot',     width: 8  },
    { header: 'NAME',    key: 'name',    width: 30 },
    { header: 'BR',      key: 'br',      width: 6  },
    { header: 'BAG',     key: 'bag',     width: 6  },
    { header: 'QTY',     key: 'qty',     width: 12 },
    { header: 'PRICE',   key: 'price',   width: 10 },
    { header: 'CONTROL', key: 'control', width: 12 },
  ];
  return createExcelBuffer('LotName', cols, rows, {
    db, title: 'Lot Name', metaLines: auctionMeta(db, auctionId),
  });
}

// Price List (Before): same shape as Price List minus the trade-result
// columns — typically printed empty so buyers can hand-fill PRICE / CODE
// during the auction. PRICE is blanked when 0 (lot not yet priced) so the
// column reads empty rather than "0.00".
async function exportPriceListBefore(db, auctionId) {
  const a = db.get('SELECT ano, date FROM auctions WHERE id = ?', [auctionId]) || {};
  const tradeNo = a.ano || '';
  const tradeDate = String(a.date || '').slice(0, 10).split('-').reverse().join('/');

  // e-Trade / Tamil Nadu wants a bare grid: no brand band (logo + title +
  // meta rows) and lot numbers zero-padded to 3 digits (e.g. "007").
  const cfg = getSettingsFlat(db);
  const bare = cfg.business_mode === 'e-Trade'
    && String(cfg.business_state || '').toUpperCase().includes('TAMIL NADU');

  const rawRows = db.all(
    `SELECT lot_no as lot, bags as bag, qty,
            CASE WHEN COALESCE(price,0) = 0 THEN '' ELSE price END AS price,
            COALESCE(code,'') AS code
     FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]
  );
  const padLot = (v) => {
    const s = String(v == null ? '' : v).trim();
    return /^\d+$/.test(s) ? s.padStart(3, '0') : s;
  };
  const rows = rawRows.map(r => ({
    trade_no: tradeNo, date: tradeDate, ...r,
    lot: bare ? padLot(r.lot) : r.lot,
  }));
  const cols = [
    { header: 'TNO',   key: 'trade_no', width: 10 },
    { header: 'DATE',  key: 'date',     width: 12 },
    { header: 'LOT',   key: 'lot',      width: 10, text: bare },
    { header: 'BAG',   key: 'bag',      width: 8  },
    { header: 'QTY',   key: 'qty',      width: 14 },
    { header: 'PRICE', key: 'price',    width: 10 },
    { header: 'CODE',  key: 'code',     width: 10 },
  ];
  return createExcelBuffer('PriceListBefore', cols, rows, {
    db, title: 'Price List (Before)', metaLines: auctionMeta(db, auctionId),
    noBrandBand: bare,
  });
}

// ── Export Type 3: Price List ─────────────────────────────────
async function exportPriceList(db, auctionId) {
  const rows = db.all(
    `SELECT lot_no as lot, bags as bag, qty, price, code, buyer as bidder
     FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]
  );
  const cols = [
    { header: 'LOT', key: 'lot', width: 8 },
    { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'PRICE', key: 'price', width: 10 },
    { header: 'CODE', key: 'code', width: 8 },
    { header: 'BIDDER', key: 'bidder', width: 20 },
  ];
  return createExcelBuffer('PriceList', cols, rows, {
    db, title: 'Price List', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export Type 4: Bank Payment (RTGS/NEFT — bank-import format) ─
// Bare 13-column sheet (no brand band, no totals) matching the bank's
// upload template:
//   TRANSACT_A | MESSAGETYP | DEBITACCOU | PAYMENTAMO | TRANSACT_B |
//   VALUEDATE  | BENEFICI_A | BENEFIARYN | BENEFIARYB | BENEFIARYE |
//   BENEFICI_B | REMARKS    | CLIENTCODE
// Header on row 1, data from row 2. Bank software auto-ingests this
// shape — adding a brand band would break the import.
async function exportBankPayment(db, auctionId, cfg, _state, opts) {
  const { getBankPaymentData, formatLotList } = require('./calculations');
  let payments = getBankPaymentData(db, auctionId, cfg);
  // Optional seller-name filter — when the user clicks "Export Bank
  // Payment (Selected)" in the Payments tab, only the ticked sellers'
  // rows should appear in the bank upload file. Match against `p.name`
  // (the seller name) because that's the value the UI checkbox holds;
  // beneficiaryName tracks the bank account holder which can be a
  // different entity. When opts.names is absent or empty the full
  // payment set is exported, preserving the original behaviour.
  if (opts && Array.isArray(opts.names) && opts.names.length) {
    const wanted = new Set(opts.names.map(n => String(n || '').trim().toUpperCase()));
    payments = payments.filter(p =>
      wanted.has(String(p.name || '').trim().toUpperCase())
    );
  }
  // Optional per-seller lot-picks AND already-exported exclusions.
  //
  //   opts.lots[seller]        — operator-picked subset; only these
  //                              lots' balances count for the bank row.
  //   opts.excludeLots[seller] — lots that have already gone out in a
  //                              previous export and must be skipped
  //                              automatically so re-exporting the
  //                              seller doesn't double-pay them.
  //
  // For each seller in either map we re-query SUM(balance) with the
  // appropriate WHERE clauses and override the payments[] row in
  // place. RTGS/NEFT is re-picked from the new amount (₹2L threshold).
  // REMARKS is rebuilt later at row.map time using p.amount, so just
  // updating p.amount + p.transactionType here is enough.
  const lotPicks = (opts && opts.lots && typeof opts.lots === 'object') ? opts.lots : null;
  const excludeLots = (opts && opts.excludeLots && typeof opts.excludeLots === 'object') ? opts.excludeLots : null;
  if (lotPicks || excludeLots) {
    // Index every trader bank by id so a picked-lot subset that all points
    // at one account can route this export row to THAT account (this is how
    // "select Account-1's lots → Export Selected" credits Account 1, then a
    // second export of Account-2's lots credits Account 2).
    const bankById = {};
    try {
      for (const b of db.all('SELECT id, ifsc, acctnum, holder_name FROM trader_banks')) bankById[b.id] = b;
    } catch (_) { /* table may not exist on partial migrations */ }
    const sellersToRecompute = new Set();
    if (lotPicks)     for (const k of Object.keys(lotPicks))     sellersToRecompute.add(k);
    if (excludeLots)  for (const k of Object.keys(excludeLots))  sellersToRecompute.add(k);
    for (const sellerName of sellersToRecompute) {
      const picksArr   = lotPicks    && Array.isArray(lotPicks[sellerName])    ? lotPicks[sellerName]    : null;
      const excludeArr = excludeLots && Array.isArray(excludeLots[sellerName]) ? excludeLots[sellerName] : null;
      if ((!picksArr || !picksArr.length) && (!excludeArr || !excludeArr.length)) continue;
      const wantedUpper = String(sellerName || '').trim().toUpperCase();
      const idx = payments.findIndex(p =>
        String(p.name || '').trim().toUpperCase() === wantedUpper
      );
      if (idx < 0) continue;   // seller not in the current payments set (e.g. fully paid)
      const params = [auctionId, wantedUpper];
      let extraWhere = '';
      if (picksArr && picksArr.length) {
        extraWhere += ` AND l.lot_no IN (${picksArr.map(() => '?').join(',')})`;
        for (const lot of picksArr) params.push(String(lot));
      }
      if (excludeArr && excludeArr.length) {
        extraWhere += ` AND l.lot_no NOT IN (${excludeArr.map(() => '?').join(',')})`;
        for (const lot of excludeArr) params.push(String(lot));
      }
      const sub = db.get(
        `SELECT COALESCE(SUM(l.balance),0) AS payable,
                COALESCE(SUM(l.puramt), 0) AS puramt,
                GROUP_CONCAT(l.lot_no) AS lot_nos,
                GROUP_CONCAT(DISTINCT l.bank_id) AS bank_ids,
                COUNT(*) AS lot_count,
                COUNT(l.bank_id) AS bank_lot_count
           FROM lots l
          WHERE l.auction_id = ? AND l.amount > 0
            AND (l.paid IS NULL OR l.paid = '')
            AND UPPER(TRIM(l.name)) = ?${extraWhere}`,
        params
      ) || { payable: 0, puramt: 0, lot_nos: '' };
      const rawAmount = Number(sub.payable) || 0;
      const roundedAmount = cfg.flag_round ? Math.round(rawAmount) : rawAmount;
      const isRTGS = roundedAmount >= 200000;
      // If every picked lot points at the same single bank account, route
      // this row to that account (overrides the seller-default account that
      // getBankPaymentData put on the base row). Mixed/untagged → leave the
      // default account as-is.
      const subBankIds = String(sub.bank_ids || '')
        .split(',').map(s => s.trim()).filter(s => s !== '' && s !== 'null')
        .map(Number).filter(Number.isFinite);
      const subUntagged = Number(sub.lot_count || 0) > Number(sub.bank_lot_count || 0);
      const subBank = (subBankIds.length === 1 && !subUntagged) ? bankById[subBankIds[0]] : null;
      payments[idx] = {
        ...payments[idx],
        amount: roundedAmount,
        transactionType: isRTGS ? 'RTGS' : 'NEFT',
        // Re-derive the covered-lots list from the same picked/excluded
        // subset so REMARKS lists exactly the lots this row pays for.
        lots: formatLotList(sub.lot_nos),
        ...(subBank ? {
          ifsc: subBank.ifsc || payments[idx].ifsc,
          accountNo: subBank.acctnum || payments[idx].accountNo,
          beneficiaryName: subBank.holder_name || payments[idx].beneficiaryName,
        } : {}),
      };
    }
    // Drop any zero-amount rows produced by the recompute — banks reject
    // zero-value RTGS rows, and a seller whose remaining lots all net to
    // zero (everything already exported, or only zero-balance lots
    // picked) shouldn't appear at all.
    payments = payments.filter(p => Number(p.amount) > 0);
  }

  // Sender-side context (state-aware): debit account, IFSC for BT/LBT
  // detection, and the email used in BENEFIARYE.
  const isKL = String(cfg.business_state || cfg.state || '').toUpperCase().includes('KERALA');
  const senderAcct  = (isKL ? cfg.bank_kl_acct  : cfg.bank_tn_acct)  || cfg.bank_tn_acct  || cfg.bank_kl_acct  || '';
  const senderIfsc  = (isKL ? cfg.bank_kl_ifsc  : cfg.bank_tn_ifsc)  || cfg.bank_tn_ifsc  || cfg.bank_kl_ifsc  || '';
  const senderEmail = (isKL ? cfg.kl_email      : cfg.tn_email)      || cfg.tn_email      || cfg.kl_email      || '';
  const senderBankPrefix = String(senderIfsc).slice(0, 4).toUpperCase();
  // Short tag inserted into REMARKS (e.g. "VSTL" → "5 ANN MARIA SPICES VSTL PAYMENT 5945275.00 Credited").
  // Falls back to the leading word of trade_name when short_name isn't set.
  const shortTag = String(cfg.short_name || (cfg.trade_name || '').split(/\s+/)[0] || '').toUpperCase();

  // Auction context: ano (REMARKS prefix) + value date (DD/MM/YYYY).
  const a = db.get('SELECT ano, date FROM auctions WHERE id = ?', [auctionId]) || {};
  const ano = a.ano || '';
  const valueDate = fmtUserDate(String(a.date || '').slice(0, 10));

  const rows = payments.map(p => {
    const amount = Number(p.amount) || 0;
    const beneIfsc = String(p.ifsc || '').toUpperCase();
    const benePrefix = beneIfsc.slice(0, 4);
    // BT  = book transfer (same bank as sender)
    // LBT = local bank transfer (different bank, RTGS/NEFT routed)
    const transactA = (senderBankPrefix && benePrefix === senderBankPrefix) ? 'BT' : 'LBT';
    return {
      TRANSACT_A:  transactA,
      MESSAGETYP:  p.transactionType || 'RTGS',
      DEBITACCOU:  senderAcct,
      PAYMENTAMO:  amount,
      TRANSACT_B:  'INR',
      VALUEDATE:   valueDate,
      BENEFICI_A:  p.accountNo || '',
      BENEFIARYN:  String(p.beneficiaryName || '').toUpperCase(),
      BENEFIARYB:  beneIfsc,
      BENEFIARYE:  senderEmail,
      BENEFICI_B:  '',
      REMARKS:     `${ano} ${String(p.beneficiaryName || '').toUpperCase()}${shortTag ? ' ' + shortTag : ''} PAYMENT ${amount.toFixed(2)} Credited${p.lots ? ` for lot${p.lots.includes(',') ? 's' : ''} ${p.lots}` : ''}`,
      CLIENTCODE:  '',
    };
  });

  // Build the sheet directly (bypass createExcelBuffer's brand-band).
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('BANK_PAYMENT');
  const cols = [
    { key: 'TRANSACT_A',  header: 'TRANSACT_A',  width: 12 },
    { key: 'MESSAGETYP',  header: 'MESSAGETYP',  width: 12 },
    { key: 'DEBITACCOU',  header: 'DEBITACCOU',  width: 18 },
    { key: 'PAYMENTAMO',  header: 'PAYMENTAMO',  width: 14, numFmt: '#,##0.00' },
    { key: 'TRANSACT_B',  header: 'TRANSACT_B',  width: 10 },
    { key: 'VALUEDATE',   header: 'VALUEDATE',   width: 12 },
    { key: 'BENEFICI_A',  header: 'BENEFICI_A',  width: 22 },
    { key: 'BENEFIARYN',  header: 'BENEFIARYN',  width: 32 },
    { key: 'BENEFIARYB',  header: 'BENEFIARYB',  width: 16 },
    { key: 'BENEFIARYE',  header: 'BENEFIARYE',  width: 28 },
    { key: 'BENEFICI_B',  header: 'BENEFICI_B',  width: 12 },
    { key: 'REMARKS',     header: 'REMARKS',     width: 60 },
    { key: 'CLIENTCODE',  header: 'CLIENTCODE',  width: 14 },
  ];
  ws.columns = cols.map(c => ({ key: c.key, width: c.width }));
  cols.forEach((c, i) => {
    if (c.numFmt) ws.getColumn(i + 1).numFmt = c.numFmt;
  });
  // Header row 1 — plain bold, no fill (so bank importers don't choke).
  const head = ws.getRow(1);
  cols.forEach((c, i) => { head.getCell(i + 1).value = c.header; });
  head.font = { bold: true };
  // Data rows from row 2.
  rows.forEach(r => ws.addRow(r));
  return wb.xlsx.writeBuffer();
}

// ── Export Type 4b: Bank Payment (Before discount) ───────────
// Same data shape as bank_payment except `amount` is the pre-discount
// puramt (raw purchase amount before refund/GST). Per the e-Trade spec
// the Amount + SendertoRcvrInfo columns are omitted from this variant.
async function exportBankPaymentBefore(db, auctionId, cfg) {
  const { getBankPaymentData } = require('./calculations');
  const payments = getBankPaymentData(db, auctionId, cfg, { before: true });
  const cols = [
    { header: 'TransactionType', key: 'transactionType', width: 16 },
    { header: 'BeneIFSCode',     key: 'ifsc',            width: 14 },
    { header: 'BeneAcctNo',      key: 'accountNo',       width: 20 },
    { header: 'BeneName',        key: 'beneficiaryName', width: 30 },
    { header: 'BeneAddLine1',    key: 'address1',        width: 30 },
    { header: 'BeneAddLine2',    key: 'address2',        width: 20 },
    { header: 'BeneAddLine3',    key: 'pin',             width: 10 },
  ];
  return createExcelBuffer('BankPaymentBefore', cols, payments, {
    db, title: 'Bank Payment (Before)', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export Type 5: Pooler-wise Register ───────────────────────
async function exportPoolerRegister(db, auctionId) {
  const rows = db.all(
    `SELECT state, lot_no as lot, name as poolername, branch as br, qty, price, amount, pqty, prate, puramt
     FROM lots WHERE auction_id = ? AND amount > 0
     ORDER BY name`, [auctionId]
  );
  const cols = [
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'NAME', key: 'poolername', width: 30 },
    { header: 'BRANCH', key: 'br', width: 15 },
    { header: 'LOT', key: 'lot', width: 8 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'PRICE', key: 'price', width: 10 },
    { header: 'AMOUNT', key: 'amount', width: 14 },
    { header: 'PQTY', key: 'pqty', width: 12 },
    { header: 'PRATE', key: 'prate', width: 10 },
    { header: 'PURAMT', key: 'puramt', width: 14 },
  ];
  return createExcelBuffer('PoolerRegister', cols, rows, {
    db, title: 'Pooler Register', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export Type 6: Full File ─────────────────────────────────
async function exportFullFile(db, auctionId) {
  const rows = db.all(`SELECT * FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]);
  const cols = [
    { header: 'STATE', key: 'state' }, { header: 'LOT', key: 'lot_no', width: 8 },
    { header: 'CROP', key: 'crop' }, { header: 'GRADE', key: 'grade' },
    { header: 'CRPT', key: 'crpt' }, { header: 'BRANCH', key: 'branch', width: 15 },
    { header: 'NAME', key: 'name', width: 30 }, { header: 'CR', key: 'cr', width: 25 },
    { header: 'PAN', key: 'pan' }, { header: 'TEL', key: 'tel' },
    { header: 'BAG', key: 'bags', width: 6 }, { header: 'QTY', key: 'qty', width: 12 },
    { header: 'PRICE', key: 'price', width: 10 }, { header: 'AMOUNT', key: 'amount', width: 14 },
    { header: 'CODE', key: 'code' }, { header: 'BUYER', key: 'buyer', width: 15 },
    { header: 'BUYER1', key: 'buyer1', width: 20 }, { header: 'SALE', key: 'sale' },
    { header: 'INVO', key: 'invo' }, { header: 'PQTY', key: 'pqty', width: 12 },
    { header: 'PRATE', key: 'prate', width: 10 }, { header: 'PURAMT', key: 'puramt', width: 14 },
    { header: 'COM', key: 'com' }, { header: 'CGST', key: 'cgst' },
    { header: 'SGST', key: 'sgst' }, { header: 'IGST', key: 'igst' },
    { header: 'ADVANCE', key: 'advance', width: 14 }, { header: 'BALANCE', key: 'balance', width: 14 },
  ];
  return createExcelBuffer('FullFile', cols, rows, {
    db, title: 'Full File', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export Type 7: Collection (invoice register) ─────────────
// Mirrors COLLECTION.pdf: one row per sales invoice issued, grouped by buyer
// state. Columns: SALE+INVO | TRADE NAME (firm) | NAME (buyer) | QTY | VALUE.
async function exportCollection(db, auctionId) {
  return newCollectionXlsx(db, auctionId);
}

// ── Export Type 8: Dealer List ────────────────────────────────
async function exportDealerList(db, auctionId) {
  // Registered-dealer roster: a pre-trade export, so it must NOT depend on
  // `amount` (lots carry no price/amount until prices are imported — see
  // exportPriceListBefore). Filtering on amount>0 made the pre-trade Dealer
  // List come back empty. Qualify on GSTIN presence + a real (qty>0) lot.
  const rows = db.all(
    `SELECT state, name, SUBSTR(cr, 7, 15) as gstin,
      COUNT(lot_no) as lots, SUM(bags) as bags, SUM(qty) as qty
     FROM lots WHERE auction_id = ? AND cr LIKE '%GST%' AND COALESCE(qty,0) > 0
     GROUP BY state, name, cr ORDER BY state, name`, [auctionId]
  );
  const cols = [
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'GSTIN', key: 'gstin', width: 18 },
    { header: 'LOTS', key: 'lots', width: 6 },
    { header: 'BAGS', key: 'bags', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
  ];
  return createExcelBuffer('DealerList', cols, rows, {
    db, title: 'Dealer List', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export Type: Planter List (Grade 1) ──────────────────────
// Pre-trade roster of planters (Grade 1 = pooler lots, no GSTIN). Run
// BEFORE prices are imported, so it must NOT filter on `amount`/`price`
// (those stay unset until price import — same reasoning as the Dealer List
// and Price List (Before)). Qualifies purely on grade = '1'. The CR column
// shows the control/registration number with the "CR." prefix stripped.
async function exportPlanterList(db, auctionId) {
  const rows = db.all(
    `SELECT state, name,
        CASE WHEN UPPER(COALESCE(cr,'')) LIKE 'CR.%' THEN TRIM(SUBSTR(cr, 4))
             ELSE COALESCE(cr,'') END AS cr,
        COUNT(lot_no) as lots, SUM(bags) as bags, SUM(qty) as qty
     FROM lots WHERE auction_id = ? AND TRIM(COALESCE(grade,'')) = '1'
     GROUP BY state, name, cr ORDER BY state, name`, [auctionId]
  );
  const cols = [
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'NAME',  key: 'name',  width: 30 },
    { header: 'CR',    key: 'cr',    width: 22 },
    { header: 'LOTS',  key: 'lots',  width: 6 },
    { header: 'BAGS',  key: 'bags',  width: 6 },
    { header: 'QTY',   key: 'qty',   width: 12 },
  ];
  return createExcelBuffer('PlanterList', cols, rows, {
    db, title: 'Planter List (Grade 1)', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export Type 9: Sales & Taxes ─────────────────────────────
async function exportSalesTaxes(db, auctionId) {
  const rows = db.all(
    `SELECT state, sale, invo, buyer1 as tradername, bags as bag, qty, 
      amount as cardamom_cost, gunny as gunny_cost,
      cgst, sgst, igst, tcs, pava_hc as transport, ins as insurance, tot as total
     FROM invoices WHERE ano = (SELECT ano FROM auctions WHERE id = ?)
     ORDER BY sale, invo`, [auctionId]
  );
  const cols = [
    { header: 'STATE', key: 'state' }, { header: 'SALE', key: 'sale' },
    { header: 'INVO', key: 'invo' }, { header: 'TRADERNAME', key: 'tradername', width: 25 },
    { header: 'BAG', key: 'bag', width: 6 }, { header: 'QTY', key: 'qty', width: 12 },
    { header: 'CARDAMOM', key: 'cardamom_cost', width: 14 },
    { header: 'GUNNY', key: 'gunny_cost', width: 10 },
    { header: 'CGST', key: 'cgst', width: 12 }, { header: 'SGST', key: 'sgst', width: 12 },
    { header: 'IGST', key: 'igst', width: 12 }, { header: 'TCS', key: 'tcs', width: 10 },
    { header: 'TRANSPORT', key: 'transport', width: 10 },
    { header: 'INSURANCE', key: 'insurance', width: 10 },
    { header: 'TOTAL', key: 'total', width: 14 },
  ];
  return createExcelBuffer('SalesTaxes', cols, rows, {
    db, title: 'Sales & Taxes', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export: Payment Summary ──────────────────────────────────
async function exportPaymentSummary(db, auctionId, cfg, _state, opts) {
  // Match getPaymentSummary semantics: discount includes BOTH the per-lot
  // policy discount AND any manual debit_notes for this auction's sellers.
  // We compute it per-row by adding debit_notes (joined by ano + name).
  const mode = (cfg && cfg.business_mode || 'e-Trade').toLowerCase();
  const discountCol = (mode === 'auction') ? 'advance' : 'refund';
  const auction = db.get('SELECT ano FROM auctions WHERE id = ?', [auctionId]);
  const ano = auction ? auction.ano : null;
  // Build name → manual debit total map (debit_notes can have multiple
  // rows per seller per auction; we sum)
  const debitMap = {};
  if (ano) {
    const debits = db.all(
      'SELECT name, SUM(amount) as total FROM debit_notes WHERE ano = ? GROUP BY name',
      [ano]
    );
    for (const d of debits) debitMap[d.name] = Number(d.total) || 0;
  }
  // Optional seller-name filter — "Export Payment XLSX (Selected)"
  // limits the rows to the ticked sellers. We push the filter into the
  // SQL with a `name IN (…)` clause so we don't waste the SELECT on rows
  // we'll throw away. Names are matched case-insensitively to be
  // resilient to slight casing drift between the UI and the DB.
  const filterNames = (opts && Array.isArray(opts.names) && opts.names.length)
    ? opts.names.map(n => String(n || '').trim()).filter(Boolean)
    : null;
  let whereExtra = '';
  const params = [auctionId];
  if (filterNames && filterNames.length) {
    const placeholders = filterNames.map(() => '?').join(',');
    whereExtra = ` AND UPPER(TRIM(name)) IN (${placeholders})`;
    for (const n of filterNames) params.push(n.toUpperCase());
  }
  let rows = db.all(
    `SELECT name as poolername, lot_no as lot, bags as bag, qty, price, amount,
      pqty, prate, puramt, ${discountCol} as lot_discount, balance as payable
     FROM lots WHERE auction_id = ? AND amount > 0${whereExtra}
     ORDER BY state, name`, params
  );
  // Optional per-seller lot-picks AND already-exported exclusions.
  // Same shape as exportBankPayment:
  //   opts.lots[seller]        — keep ONLY these lot rows for the seller
  //   opts.excludeLots[seller] — drop these lots (already shipped before)
  // Both filters compose: if a seller has both, the row must satisfy
  // BOTH conditions to survive. Match name case-insensitively to be
  // tolerant of slight casing drift between localStorage and the DB.
  const lotPicks    = (opts && opts.lots         && typeof opts.lots         === 'object') ? opts.lots         : null;
  const excludeLots = (opts && opts.excludeLots  && typeof opts.excludeLots  === 'object') ? opts.excludeLots  : null;
  if (lotPicks || excludeLots) {
    const picksUpper   = {};
    const excludeUpper = {};
    if (lotPicks) for (const k of Object.keys(lotPicks)) {
      const arr = Array.isArray(lotPicks[k]) ? lotPicks[k] : [];
      if (arr.length) picksUpper[k.trim().toUpperCase()] = new Set(arr.map(x => String(x)));
    }
    if (excludeLots) for (const k of Object.keys(excludeLots)) {
      const arr = Array.isArray(excludeLots[k]) ? excludeLots[k] : [];
      if (arr.length) excludeUpper[k.trim().toUpperCase()] = new Set(arr.map(x => String(x)));
    }
    rows = rows.filter(r => {
      const key = String(r.poolername || '').trim().toUpperCase();
      const lotKey = String(r.lot);
      const picks = picksUpper[key];
      if (picks && !picks.has(lotKey)) return false;
      const excl = excludeUpper[key];
      if (excl && excl.has(lotKey)) return false;
      return true;
    });
  }
  // Spread debit_notes amount across the seller's lots proportionally so
  // every row totals to the same SUM as the payments view. Simpler approach:
  // attribute the FULL manual debit on the FIRST row for each seller; later
  // rows show only the lot policy discount. Avoids per-row arithmetic but
  // still preserves the seller-level total.
  // TDS: the seller's stamped Section-194Q purchase TDS, spread ∝ each lot's
  // puramt so a lot-picked subset nets the proportionate share. "Total" is
  // the pre-TDS payable; "Payable" = Total − TDS (what the seller is paid).
  const { paymentTdsContext } = require('./calculations');
  const tdsCtx = paymentTdsContext(db, auctionId);
  const seenSellers = new Set();
  const enrichedFlat = rows.map(r => {
    const lotDisc = Number(r.lot_discount) || 0;
    const manualDisc = (!seenSellers.has(r.poolername))
      ? (Number(debitMap[r.poolername]) || 0)
      : 0;
    seenSellers.add(r.poolername);
    const total = (Number(r.payable) || 0) - manualDisc;   // pre-TDS payable
    const tds = tdsCtx.share(r.poolername, r.puramt);
    return {
      ...r,
      discount: lotDisc + manualDisc,
      total,
      tds,
      payable: total - tds,
    };
  });

  // Interleave per-pooler subtotal rows after each name group — mirrors
  // the PDF's groupByKey:'poolername' subtotalKeys behaviour. Rows are
  // already sorted by (state, name) so a single linear pass groups them.
  const SUB_KEYS = ['bag', 'qty', 'amount', 'pqty', 'puramt', 'discount', 'total', 'tds', 'payable'];
  const enriched = [];
  let curName = null;
  let acc = null;
  const flushSub = () => {
    if (!acc || curName == null) return;
    const sub = { _isSubtotal: true, poolername: `${curName} TOTAL` };
    SUB_KEYS.forEach(k => { sub[k] = acc[k] || 0; });
    enriched.push(sub);
  };
  for (const r of enrichedFlat) {
    const k = r.poolername || '';
    if (k !== curName) {
      flushSub();
      curName = k;
      acc = Object.fromEntries(SUB_KEYS.map(x => [x, 0]));
    }
    SUB_KEYS.forEach(x => { acc[x] += Number(r[x]) || 0; });
    enriched.push(r);
  }
  flushSub();
  const cols = [
    { header: 'POOLERNAME', key: 'poolername', width: 30 },
    { header: 'LOT', key: 'lot', width: 8 }, { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 }, { header: 'PRICE', key: 'price', width: 10 },
    { header: 'AMOUNT', key: 'amount', width: 14 }, { header: 'PQTY', key: 'pqty', width: 12 },
    { header: 'PRATE', key: 'prate', width: 10 }, { header: 'PURAMT', key: 'puramt', width: 14 },
    { header: 'DISCOUNT', key: 'discount', width: 14 },
    { header: 'TOTAL', key: 'total', width: 14 },
    { header: 'TDS', key: 'tds', width: 12 },
    { header: 'PAYABLE', key: 'payable', width: 14 },
  ];
  // Footer totals — sum every numeric column. The earlier export had no
  // totals row, so users had to compute payable/discount sums manually
  // in Excel before reconciling with bank transfers. PRICE/PRATE are
  // omitted from the sum (averaging rates makes no business sense; a
  // sum would mislead readers). Sum over enrichedFlat (data rows only)
  // so interleaved subtotals don't double-count.
  const sum = (key) => enrichedFlat.reduce((s, r) => s + (Number(r[key]) || 0), 0);
  const grandTotal = {
    label: 'GRAND TOTAL',
    values: {
      bag:     sum('bag'),
      qty:     sum('qty'),
      amount:  sum('amount'),
      pqty:    sum('pqty'),
      puramt:  sum('puramt'),
      discount:sum('discount'),
      total:   sum('total'),
      tds:     sum('tds'),
      payable: sum('payable'),
    },
  };
  return createExcelBuffer('Payment', cols, enriched, {
    db, title: 'Payment Summary', metaLines: auctionMeta(db, auctionId),
    grandTotal,
  });
}

// ── Export: Payment Summary (Party-wise) ─────────────────────
// One aggregated row per seller (party) — NO per-lot breakdown. Mirrors the
// Payments screen rollup: Qty, Purchase Amount, Discount (policy + manual
// debit notes), Total (pre-TDS), TDS, and Payable (= Total − TDS). Built
// straight from getPaymentSummary so the figures match the on-screen view
// and the per-lot Payment Summary export to the rupee.
async function exportPaymentPartyWise(db, auctionId, cfg, state, _opts) {
  const { getPaymentSummary } = require('./calculations');
  const sellers = getPaymentSummary(db, auctionId, state || '', cfg);
  const cols = [
    { header: 'SELLER',   key: 'name',           width: 32 },
    { header: 'LOTS',     key: 'lot_count',      width: 8  },
    { header: 'QTY',      key: 'total_qty',      width: 12 },
    { header: 'PUR.AMT',  key: 'total_puramt',   width: 14 },
    { header: 'DISCOUNT', key: 'total_discount', width: 14 },
    { header: 'TOTAL',    key: 'total_total',    width: 14 },
    { header: 'TDS',      key: 'total_tds',      width: 12 },
    { header: 'PAYABLE',  key: 'total_payable',  width: 14 },
  ];
  const sum = (k) => sellers.reduce((s, r) => s + (Number(r[k]) || 0), 0);
  const grandTotal = {
    label: 'GRAND TOTAL',
    values: {
      lot_count:      sum('lot_count'),
      total_qty:      sum('total_qty'),
      total_puramt:   sum('total_puramt'),
      total_discount: sum('total_discount'),
      total_total:    sum('total_total'),
      total_tds:      sum('total_tds'),
      total_payable:  sum('total_payable'),
    },
  };
  return createExcelBuffer('PaymentPartyWise', cols, sellers, {
    db, title: 'Payment Summary (Party-wise)', metaLines: auctionMeta(db, auctionId),
    grandTotal,
  });
}

// ── Export: Lot Payment Summary ──────────────────────────────
// Fully-populated post-auction payment summary, ordered by branch then
// seller name so the natural printed layout (branch header followed by
// that branch's lots) emerges from the row sequence.
async function exportLotPayment(db, auctionId) {
  const rows = db.all(
    `SELECT COALESCE(branch,'') AS branch,
            lot_no AS lot, qty, price AS rate, amount AS cost,
            pqty, prate, puramt AS purchamt,
            COALESCE(name,'') AS seller_name
     FROM lots WHERE auction_id = ? ORDER BY branch, name, lot_no`,
    [auctionId]
  );
  const cols = [
    { header: 'BRANCH',      key: 'branch',      width: 14 },
    { header: 'LOT',         key: 'lot',         width: 6  },
    { header: 'QTY',         key: 'qty',         width: 10 },
    { header: 'RATE',        key: 'rate',        width: 10 },
    { header: 'COST',        key: 'cost',        width: 14 },
    { header: 'PQTY',        key: 'pqty',        width: 10 },
    { header: 'PRATE',       key: 'prate',       width: 10 },
    { header: 'PURCHAMT',    key: 'purchamt',    width: 14 },
    { header: 'SELLER NAME', key: 'seller_name', width: 26 },
  ];
  return createExcelBuffer('LotPayment', cols, rows, {
    db, title: 'Lot Payment Summary', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export: TDS Return ───────────────────────────────────────
async function exportTDSReturn(db, fromDate, toDate) {
  const { getTDSReturnData } = require('./calculations');
  const rows = getTDSReturnData(db, fromDate, toDate, 'invoice');
  const cols = [
    { header: 'INVOICE', key: 'invoice', width: 10 },
    { header: 'DATE', key: 'date', width: 12 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'PAN', key: 'pan', width: 12 },
    { header: 'ASSESS_VALUE', key: 'assess_value', width: 14 },
    { header: 'TDS', key: 'tds', width: 12 },
  ];
  return createExcelBuffer('TDSReturn', cols, rows, {
    db, title: 'TDS Return', metaLines: [`From: ${fromDate}`, `To: ${toDate}`],
  });
}

// ── Export: Tally format (TALY.PRG — purchase data for accounting)
async function exportTallyPurchase(db, auctionId, cfg) {
  const mode = (cfg && cfg.business_mode || 'e-Trade').toLowerCase();
  const discountCol = (mode === 'auction') ? 'advance' : 'refund';
  const rows = db.all(
    `SELECT name, padd as add, ppla as place, cr as gstin, tel,
      lot_no as lot, bags as bag, pqty as qty, prate as price, puramt as amount,
      cgst, sgst, igst, ${discountCol} as discount, puramt as bilamt
     FROM lots WHERE auction_id = ? AND amount > 0
      AND cr NOT LIKE 'GSTIN.%'
     ORDER BY name`, [auctionId]
  );
  const cols = [
    { header: 'NAME', key: 'name', width: 30 }, { header: 'ADD', key: 'add', width: 30 },
    { header: 'PLACE', key: 'place', width: 15 }, { header: 'GSTIN', key: 'gstin', width: 20 },
    { header: 'TEL', key: 'tel', width: 14 }, { header: 'LOT', key: 'lot', width: 8 },
    { header: 'BAG', key: 'bag', width: 6 }, { header: 'QTY', key: 'qty', width: 12 },
    { header: 'PRICE', key: 'price', width: 10 }, { header: 'AMOUNT', key: 'amount', width: 14 },
    { header: 'CGST', key: 'cgst', width: 12 }, { header: 'SGST', key: 'sgst', width: 12 },
    { header: 'IGST', key: 'igst', width: 12 }, { header: 'DISCOUNT', key: 'discount', width: 14 },
    { header: 'BILAMT', key: 'bilamt', width: 14 },
  ];
  return createExcelBuffer('TallyPurchase', cols, rows, {
    db, title: 'Tally Purchase', metaLines: auctionMeta(db, auctionId),
  });
}

// ── Export: Sales Journal (JOUR.PRG) ────────────────────────
async function exportSalesJournal(db, fromDate, toDate, saleType) {
  const { getSalesJournal } = require('./calculations');
  const rows = getSalesJournal(db, fromDate, toDate, saleType);
  const cols = [
    { header: 'DATE', key: 'date', width: 12 },
    { header: 'SALE', key: 'sale', width: 6 },
    { header: 'INV#', key: 'invo', width: 8 },
    { header: 'BUYER', key: 'buyer', width: 8 },
    { header: 'TRADE NAME', key: 'buyer1', width: 30 },
    { header: 'GSTIN', key: 'gstin', width: 20 },
    { header: 'PLACE', key: 'place', width: 15 },
    { header: 'BAGS', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'CARDAMOM', key: 'cardamom', width: 14 },
    { header: 'GUNNY', key: 'gunny', width: 10 },
    { header: 'TRANSPORT', key: 'transport', width: 10 },
    { header: 'INSURANCE', key: 'insurance', width: 10 },
    { header: 'CGST', key: 'cgst', width: 10 },
    { header: 'SGST', key: 'sgst', width: 10 },
    { header: 'IGST', key: 'igst', width: 10 },
    { header: 'TCS', key: 'tcs', width: 10 },
    { header: 'ROUND', key: 'rund', width: 8 },
    { header: 'TOTAL', key: 'total', width: 14 },
  ];
  return createExcelBuffer('SalesJournal', cols, rows, {
    db, title: 'Sales Journal',
    metaLines: [`From: ${fromDate}`, `To: ${toDate}`, saleType ? `Type: ${saleType}` : ''].filter(Boolean),
  });
}

// ── Export: Purchase Journal (PUJOUR.PRG / PPUJOUR.PRG) ────
async function exportPurchaseJournal(db, fromDate, toDate, type) {
  const { getPurchaseJournal } = require('./calculations');
  const rows = getPurchaseJournal(db, fromDate, toDate, type);
  const cols = type === 'agri' ? [
    { header: 'DATE', key: 'date', width: 12 },
    { header: 'BILL#', key: 'bill_no', width: 8 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'ADDRESS', key: 'address', width: 30 },
    { header: 'PLACE', key: 'place', width: 15 },
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'CR', key: 'cr', width: 15 },
    { header: 'PAN', key: 'pan', width: 12 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'COST', key: 'cost', width: 14 },
    { header: 'IGST', key: 'igst', width: 10 },
    { header: 'NET', key: 'net', width: 14 },
  ] : [
    { header: 'DATE', key: 'date', width: 12 },
    { header: 'INV#', key: 'invoice_no', width: 8 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'ADDRESS', key: 'address', width: 30 },
    { header: 'PLACE', key: 'place', width: 15 },
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'GSTIN', key: 'gstin', width: 20 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'AMOUNT', key: 'amount', width: 14 },
    { header: 'CGST', key: 'cgst', width: 10 },
    { header: 'SGST', key: 'sgst', width: 10 },
    { header: 'IGST', key: 'igst', width: 10 },
    { header: 'ROUND', key: 'rund', width: 8 },
    { header: 'TOTAL', key: 'total', width: 14 },
    { header: 'TDS', key: 'tds', width: 10 },
  ];
  const name = type === 'agri' ? 'AgriBillJournal' : 'PurchaseJournal';
  return createExcelBuffer(name, cols, rows, {
    db,
    title: type === 'agri' ? 'Agri Bill Journal' : 'Purchase Journal',
    metaLines: [`From: ${fromDate}`, `To: ${toDate}`],
  });
}

// ── Export: Praman CSV (Lot Slip in Praman auction platform format) ──
// Produces a CSV (NOT xlsx) matching the column layout required by Praman's
// lot-upload interface. Returns a Buffer of CSV text.
//
// Special rule (item #9): Grade 1 lots → Lot Company = 'ASP' on the CSV
// output only (doesn't change stored data). All other grades → 'ISPL'.
// Rationale: Grade 1 (pooler) lots are routed to ASP for tax/accounting
// reasons, but they still appear as ISPL lots in the local DB.
async function exportPramanCSV(db, auctionId, cfg, state) {
  const rows = db.all(
    `SELECT lot_no, branch, grade, name, cr, qty, litre, bags, tel
     FROM lots WHERE auction_id = ? ${state ? 'AND state = ?' : ''}
     ORDER BY CAST(lot_no AS INTEGER), lot_no`,
    state ? [auctionId, state] : [auctionId]
  );

  const header = [
    'Lot Number', 'Lot Company', 'Collection Centre', 'Planter/Dealer',
    'Planter Name', 'CRNO/SBL No', 'Quantity(Kg)', 'Litre Weight(Gms)',
    'Bags', 'Grade Type', 'Grade', 'Reserved Price', 'Auction Start Price(Rs)',
    'Immature Seeds(%)', 'Moisture Content(%)', 'Planter Mobile Number',
    'Youtube Video Link'
  ];

  // Escape a CSV field: wrap in quotes if it contains comma/quote/newline,
  // and double-up any embedded quotes. Undefined/null → empty.
  const csvEscape = (v) => {
    if (v === null || v === undefined) return '';
    const s = String(v);
    if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  };

  // Per business rule: every Praman row reports ASP as the planter,
  // regardless of which trader actually supplied the lot. This is for the
  // Praman platform's expected upload format — internal records still keep
  // the actual trader name on `lots.name`.
  const aspName    = (cfg && cfg.s_company) || 'AMAZING SPICE PARK PRIVATE LIMITED';
  const aspGstin   = (cfg && cfg.s_gstin)   || '';
  // Planter Mobile is the Kerala address phone (kl_phone), since ASP
  // is registered at the Kerala address. Falls back to s_mobile if
  // kl_phone is blank.
  const aspMobile  = (cfg && (cfg.kl_phone || cfg.s_mobile)) || '';
  const planterDealer = 2; // 2 = Dealer (always, since ASP is the legal seller)

  const lines = [header.join(',')];
  for (const r of rows) {
    // Per business rule: every Praman row reports ISPL as the lot company
    // regardless of grade. Earlier rule (Grade 1 → ASP) is no longer applied
    // since the upload flow now treats all e-Trade lots as ISPL-fronted.
    const lotCompany = 'ISPL';

    lines.push([
      r.lot_no || '',
      lotCompany,
      r.branch || '',
      planterDealer,
      aspName,
      aspGstin,
      r.qty || '',
      r.litre || '',
      r.bags || '',
      '', // Grade Type (not captured — blank as per sample)
      '', // Grade (Praman's own grade codes, not ours — blank)
      '', // Reserved Price (blank)
      '', // Auction Start Price (blank)
      '', // Immature Seeds (blank)
      '', // Moisture Content (blank)
      aspMobile,
      '', // Youtube link (blank)
    ].map(csvEscape).join(','));
  }

  // CSV text → Buffer. Prefix with BOM so Excel on Windows opens with
  // UTF-8 correctly (otherwise accented characters break).
  return Buffer.from('\uFEFF' + lines.join('\r\n'), 'utf8');
}

// ── Export Type 12: Trade Report (BUYERS LIST FOR VERIFICATION) ──
async function exportTradeReport(db, auctionId) {
  return tradeReportXlsx(db, auctionId);
}

// ── Export: e-Auction (Spices Board) CSV ──────────────────────
// The e-Auction equivalent of the Praman (e-Trade) CSV upload. Reuses the
// Spice Board tab's eauction_csv builder, but exports ALL lots in the
// auction (allLots:true) — like the Praman CSV it replaces, this is a
// pre-auction lot-catalogue upload, so it must work before prices land.
// (The Spice Board tab's own copy keeps its sold-only filter unchanged.)
async function exportEauctionCSV(db, auctionId) {
  return SPICE_BOARD_REPORTS.eauction_csv.csv(db, { auctionId, allLots: true });
}

// ── Export router ────────────────────────────────────────────
const EXPORT_TYPES = {
  lot_slip:       { fn: exportLotSlip,       name: 'LotSlip' },
  lot_slip_after: { fn: exportLotSlipAfter,  name: 'LotSlipAfter' },
  // "Before trade" snapshots surfaced from the Price Import modal (not
  // listed in the Export Center, which is driven by the frontend
  // EXP_LABELS map).
  lot_buyer:         { fn: exportLotBuyer,        name: 'LotBuyer' },
  lot_name:          { fn: exportLotName,         name: 'LotName' },
  price_list_before: { fn: exportPriceListBefore, name: 'PriceListBefore' },
  lot_payment:       { fn: exportLotPayment,      name: 'LotPayment' },
  praman_csv:     { fn: exportPramanCSV,     name: 'eTrade_Praman', ext: 'csv', mime: 'text/csv', needsCfg: true },
  eauction_csv:   { fn: exportEauctionCSV,   name: 'EAuctionCSV',   ext: 'csv', mime: 'text/csv' },
  price_list:     { fn: exportPriceList,     name: 'PriceList' },
  bank_payment_before:{ fn: exportBankPaymentBefore, name: 'BankPaymentBefore', needsCfg: true },
  bank_payment:   { fn: exportBankPayment,   name: 'BankPayment', needsCfg: true },
  pooler_register:{ fn: exportPoolerRegister,name: 'PoolerRegister' },
  full_file:      { fn: exportFullFile,      name: 'FullFile' },
  collection:     { fn: exportCollection,    name: 'Collection' },
  trade_report:   { fn: exportTradeReport,   name: 'TradeReport' },
  dealer_list:    { fn: exportDealerList,    name: 'DealerList' },
  planter_list:   { fn: exportPlanterList,   name: 'PlanterList' },
  sales_taxes:    { fn: exportSalesTaxes,    name: 'SalesTaxes' },
  payment:        { fn: exportPaymentSummary,name: 'Payment',        needsCfg: true },
  payment_partywise: { fn: exportPaymentPartyWise, name: 'PaymentPartyWise', needsCfg: true },
  tally_purchase: { fn: exportTallyPurchase, name: 'TallyPurchase',  needsCfg: true },
};

// Header meta lines for the Registers — trade (when scoped to one) or a
// date range (when spanning trades), plus an optional sale-type note.
function registerMeta(db, opts) {
  const lines = [];
  if (opts && opts.auctionId) lines.push(...auctionMeta(db, opts.auctionId));
  else if (opts && opts.from && opts.to) lines.push(`Period: ${opts.from} to ${opts.to}`);
  else lines.push('All trades');
  if (opts && opts.saleType) lines.push(`Sale: ${opts.saleType}`);
  return lines.filter(Boolean);
}

// ── Export: Purchase Register (lot-wise) ───────────────────
async function exportPurchaseRegister(db, opts = {}) {
  const { getPurchaseRegister } = require('./calculations');
  const rows = getPurchaseRegister(db, opts);
  const cols = [
    { header: 'STATE',  key: 'state',  width: 14 },
    { header: 'TNO',    key: 'tno',    width: 6  },
    { header: 'DATE',   key: 'date',   width: 12 },
    { header: 'LOT',    key: 'lot',    width: 8  },
    { header: 'BRANCH', key: 'branch', width: 10 },
    { header: 'NAME',   key: 'name',   width: 28 },
    { header: 'PLACE',  key: 'place',  width: 14 },
    { header: 'GSTIN',  key: 'gstin',  width: 18 },
    { header: 'BAG',    key: 'bag',    width: 6  },
    { header: 'QTY',    key: 'qty',    width: 11, numFmt: '#,##0.000' },
    { header: 'PRICE',  key: 'price',  width: 10, numFmt: '#,##0.00' },
    { header: 'AMOUNT', key: 'amount', width: 14, numFmt: '#,##0.00' },
    { header: 'PQTY',   key: 'pqty',   width: 11, numFmt: '#,##0.000' },
    { header: 'PRATE',  key: 'prate',  width: 10, numFmt: '#,##0.00' },
    { header: 'PURAMT', key: 'puramt', width: 14, numFmt: '#,##0.00' },
    { header: 'DISCOUNT', key: 'discount', width: 12, numFmt: '#,##0.00' },
    { header: 'GST5',   key: 'gst5',   width: 11, numFmt: '#,##0.00' },
    { header: 'PAYABLE', key: 'payable', width: 14, numFmt: '#,##0.00' },
  ];
  const sum = (k) => rows.reduce((s, r) => s + (Number(r[k]) || 0), 0);
  const grandTotal = { label: 'TOTAL', values: {
    bag: sum('bag'), qty: sum('qty'), amount: sum('amount'), pqty: sum('pqty'),
    puramt: sum('puramt'), discount: sum('discount'), gst5: sum('gst5'), payable: sum('payable'),
  }};
  return createExcelBuffer('PurchaseRegister', cols, rows, {
    db, title: 'Purchase Register', metaLines: registerMeta(db, opts), grandTotal,
  });
}

// ── Export: Sales Register (invoice-wise) ──────────────────
async function exportSalesRegister(db, opts = {}) {
  const { getSalesRegister } = require('./calculations');
  const rows = getSalesRegister(db, opts);
  const cols = [
    { header: 'STATE',  key: 'state',  width: 14 },
    { header: 'TNO',    key: 'tno',    width: 6  },
    { header: 'DATE',   key: 'date',   width: 12 },
    { header: 'SALE',   key: 'sale',   width: 6  },
    { header: 'INVO',   key: 'invo',   width: 8  },
    { header: 'TRADERNAME', key: 'tradername', width: 30 },
    { header: 'BIDDER', key: 'bidder', width: 10 },
    { header: 'BAG',    key: 'bag',    width: 6  },
    { header: 'QTY',    key: 'qty',    width: 11, numFmt: '#,##0.000' },
    { header: 'AMOUNT', key: 'amount', width: 14, numFmt: '#,##0.00' },
    { header: 'LORRY',  key: 'lorry',  width: 10, numFmt: '#,##0.00' },
    { header: 'GUNNY',  key: 'gunny',  width: 10, numFmt: '#,##0.00' },
    { header: 'IGST',   key: 'igst',   width: 10, numFmt: '#,##0.00' },
    { header: 'CGST',   key: 'cgst',   width: 10, numFmt: '#,##0.00' },
    { header: 'SGST',   key: 'sgst',   width: 10, numFmt: '#,##0.00' },
    { header: 'INS',    key: 'ins',    width: 10, numFmt: '#,##0.00' },
    { header: 'INVAMT', key: 'invamt', width: 14, numFmt: '#,##0.00' },
  ];
  const sum = (k) => rows.reduce((s, r) => s + (Number(r[k]) || 0), 0);
  const grandTotal = { label: 'TOTAL', values: {
    bag: sum('bag'), qty: sum('qty'), amount: sum('amount'), lorry: sum('lorry'),
    gunny: sum('gunny'), igst: sum('igst'), cgst: sum('cgst'), sgst: sum('sgst'),
    ins: sum('ins'), invamt: sum('invamt'),
  }};
  return createExcelBuffer('SalesRegister', cols, rows, {
    db, title: 'Sales Register', metaLines: registerMeta(db, opts), grandTotal,
  });
}

// ── Export: Per-party "Individual" Registers (cross-auction) ───────
// Pooler / Seller / Merchant statements, one section per party. Shares the
// createExcelBuffer section-grouped mode: each party becomes a banded
// section (name + GSTIN) followed by its rows, a bold TOTAL subtotal, and a
// summary line (Sold/Not Sold for poolers, Closing Balance for the others).
// `labelKey` is the first column the TOTAL/summary labels land in.
const INDIVIDUAL_REG_DEFS = {
  pooler: {
    sheet: 'PoolerRegister', title: 'Pooler Register', labelKey: 'tno',
    cols: [
      { header: 'TNO',    key: 'tno',    width: 8  },
      { header: 'DATE',   key: 'date',   width: 12 },
      { header: 'LOT',    key: 'lot',    width: 8  },
      { header: 'QTY',    key: 'qty',    width: 12, numFmt: '#,##0.000' },
      { header: 'RATE',   key: 'rate',   width: 11, numFmt: '#,##0.00'  },
      { header: 'VALUE',  key: 'value',  width: 16, numFmt: '#,##0.00'  },
      { header: 'PQTY',   key: 'pqty',   width: 12, numFmt: '#,##0.000' },
      { header: 'PRATE',  key: 'prate',  width: 11, numFmt: '#,##0.00'  },
      { header: 'PURAMT', key: 'puramt', width: 16, numFmt: '#,##0.00'  },
    ],
    summaryRows: (p) => ([
      { _isSubtotal: true, tno: 'Total',     qty: p.summary.qty,          value: p.summary.value, pqty: p.summary.pqty, puramt: p.summary.puramt },
      { _isSubtotal: true, tno: 'Sold',      qty: p.summary.soldQty,      value: p.summary.soldValue },
      { _isSubtotal: true, tno: 'Withdrawn', qty: p.summary.withdrawnQty, value: p.summary.withdrawnValue },
      { _isSubtotal: true, tno: 'Not Sold',  qty: p.summary.notSoldQty },
    ]),
    grandKeys: ['qty', 'value', 'pqty', 'puramt'],
  },
  seller: {
    sheet: 'SellerRegister', title: 'Sellers Individual', labelKey: 'date',
    cols: [
      { header: 'DATE',    key: 'date',    width: 12 },
      { header: 'ANO',     key: 'ano',     width: 8  },
      { header: 'INVO',    key: 'invo',    width: 8,  numFmt: '#,##0' },
      { header: 'QTY',     key: 'qty',     width: 12, numFmt: '#,##0.000' },
      { header: 'INVOICE', key: 'invoice', width: 16, numFmt: '#,##0.00' },
    ],
    summaryRows: (p) => ([
      { _isSubtotal: true, date: 'Total',           qty: p.summary.qty, invoice: p.summary.invoice },
      { _isSubtotal: true, date: 'Closing Balance', invoice: p.summary.closing },
    ]),
    grandKeys: ['qty', 'invoice'],
  },
  merchant: {
    sheet: 'MerchantRegister', title: 'Merchants Individual', labelKey: 'date',
    cols: [
      { header: 'DATE',    key: 'date',    width: 12 },
      { header: 'TNO',     key: 'tno',     width: 8  },
      { header: 'INVO',    key: 'invo',    width: 8  },
      { header: 'RECP',    key: 'recp',    width: 8  },
      { header: 'QTY',     key: 'qty',     width: 12, numFmt: '#,##0.000' },
      { header: 'INVOICE', key: 'invoice', width: 16, numFmt: '#,##0.00' },
      { header: 'RECEIPT', key: 'receipt', width: 16, numFmt: '#,##0.00' },
    ],
    summaryRows: (p) => ([
      { _isSubtotal: true, date: 'Total',           qty: p.summary.qty, invoice: p.summary.invoice, receipt: p.summary.receipt },
      { _isSubtotal: true, date: 'Closing Balance', invoice: p.summary.closing },
    ]),
    grandKeys: ['qty', 'invoice', 'receipt'],
  },
};

function individualRegisterData(db, kind, opts) {
  const { getPoolerRegister, getSellerRegister, getMerchantRegister } = require('./calculations');
  if (kind === 'seller')   return getSellerRegister(db, opts);
  if (kind === 'merchant') return getMerchantRegister(db, opts);
  return getPoolerRegister(db, opts);
}

async function exportIndividualRegister(db, kind, opts = {}) {
  const def = INDIVIDUAL_REG_DEFS[kind];
  if (!def) throw new Error(`Unknown individual register kind: ${kind}`);
  const data = individualRegisterData(db, kind, opts);
  const sections = data.parties.map(p => ({
    title: p.name + (p.gstin ? `      GSTIN: ${p.gstin}` : ''),
    rows: [...p.rows, ...def.summaryRows(p)],
  }));
  // Grand total across every party in the file.
  const gv = {};
  def.grandKeys.forEach(k => {
    gv[k] = data.parties.reduce((s, p) => s + (Number(p.summary[k]) || 0), 0);
  });
  gv[def.labelKey] = 'GRAND TOTAL';
  return createExcelBuffer(def.sheet, def.cols, [], {
    db, title: def.title, metaLines: registerMeta(db, opts),
    sections, spacerBetween: true,
    grandTotal: { values: gv },
  });
}

module.exports = {
  EXPORT_TYPES,
  exportLotSlip, exportLotSlipAfter, exportLotBuyer, exportLotName, exportLotPayment, exportPriceListBefore,
  exportPramanCSV, exportPriceList, exportBankPayment, exportBankPaymentBefore,
  exportPoolerRegister, exportFullFile, exportCollection, exportTradeReport, exportDealerList, exportPlanterList,
  exportSalesTaxes, exportPaymentSummary, exportPaymentPartyWise, exportTDSReturn, exportTallyPurchase,
  exportSalesJournal, exportPurchaseJournal,
  exportPurchaseRegister, exportSalesRegister, exportIndividualRegister,
};
