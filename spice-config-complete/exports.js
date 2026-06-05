/**
 * exports.js — All export formats
 * Replaces: EXP.PRG (11 types), TALY.PRG, KOTALLY.PRG, BANKPAY export
 */

const ExcelJS = require('exceljs');
const { collectionXlsx: newCollectionXlsx, tradeReportXlsx } = require('./auction-reports');
const {
  getCompanyHeader, writeXlsxCompanyHeader, xlsxNumFmtForHeader,
} = require('./report-formatters');

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
  const header = opts.companyHeader || getCompanyHeader(opts.db);
  const startRow = writeXlsxCompanyHeader(wb, ws, header, {
    colCount: columns.length,
    title: opts.title || sheetName,
    metaLines: opts.metaLines || [],
  });

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
    const fmt = xlsxNumFmtForHeader(c.header);
    if (fmt) {
      const colObj = ws.getColumn(i + 1);
      colObj.numFmt = fmt;
      colObj.alignment = { horizontal: 'right' };
    }
  });

  // Data rows. addRow uses keys from ws.columns to map object → cells.
  rows.forEach((rowObj) => {
    const dataRow = ws.addRow({});
    columns.forEach((c, i) => {
      let v = rowObj[c.key];
      // Coerce string-numbers to numbers so Excel applies the numFmt.
      if (typeof v === 'string' && v !== '' && !isNaN(Number(v))) {
        const n = Number(v);
        if (!Number.isNaN(n) && xlsxNumFmtForHeader(c.header)) v = n;
      }
      dataRow.getCell(i + 1).value = v == null ? '' : v;
    });
  });

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

// ── Pre-Trade Export: blank-priced sheet for the auction floor ──
// The complement to the "Price Import" modal. Emits exactly the columns
// that import reads back (STATE, TNO, DATE, LOT, BAG, QTY, PRICE, CODE)
// so the sheet round-trips: export before the auction, fill in PRICE +
// CODE on the floor, then re-import via "Price Import". STATE/TNO/DATE/
// LOT/BAG/QTY come pre-filled; PRICE and CODE are blank when the lot is
// still unpriced (price 0 / no code) so the floor sees empty cells to
// fill. NAME is included for readability and is ignored on import.
async function exportPreTrade(db, auctionId) {
  const auc = db.get('SELECT ano, date FROM auctions WHERE id = ?', [auctionId]) || {};
  const rows = db.all(
    `SELECT state, lot_no as lot, name, bags as bag, qty, price, code
       FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]
  ).map(r => ({
    state: r.state || '',
    tno:   auc.ano  || '',
    date:  auc.date || '',
    lot:   r.lot,
    name:  r.name || '',
    bag:   r.bag,
    qty:   r.qty,
    // Blank (not 0) when unpriced so the floor gets an empty cell to fill;
    // a re-export of an already-priced lot still carries its value through.
    price: r.price || '',
    code:  r.code  || '',
  }));
  const cols = [
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'TNO',   key: 'tno',   width: 8 },
    { header: 'DATE',  key: 'date',  width: 12 },
    { header: 'LOT',   key: 'lot',   width: 8 },
    { header: 'NAME',  key: 'name',  width: 30 },
    { header: 'BAG',   key: 'bag',   width: 6 },
    { header: 'QTY',   key: 'qty',   width: 12 },
    { header: 'PRICE', key: 'price', width: 10 },
    { header: 'CODE',  key: 'code',  width: 8 },
  ];
  return createExcelBuffer('PreTrade', cols, rows, {
    db, title: 'Pre-Trade Price Sheet', metaLines: auctionMeta(db, auctionId),
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

// ── Export Type 4: Bank Payment (RTGS/NEFT format) ───────────
async function exportBankPayment(db, auctionId, cfg) {
  const { getBankPaymentData } = require('./calculations');
  const payments = getBankPaymentData(db, auctionId, cfg);
  const cols = [
    { header: 'TransactionType', key: 'transactionType', width: 16 },
    { header: 'BeneIFSCode', key: 'ifsc', width: 14 },
    { header: 'BeneAcctNo', key: 'accountNo', width: 20 },
    { header: 'BeneName', key: 'beneficiaryName', width: 30 },
    { header: 'BeneAddLine1', key: 'address1', width: 30 },
    { header: 'BeneAddLine2', key: 'address2', width: 20 },
    { header: 'BeneAddLine3', key: 'pin', width: 10 },
    { header: 'Amount', key: 'amount', width: 14 },
    { header: 'SendertoRcvrInfo', key: 'remarks', width: 50 },
  ];
  return createExcelBuffer('BankPayment', cols, payments, {
    db, title: 'Bank Payment (RTGS/NEFT)', metaLines: auctionMeta(db, auctionId),
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
  const rows = db.all(
    `SELECT state, name, SUBSTR(cr, 7, 15) as gstin, 
      COUNT(lot_no) as lots, SUM(bags) as bags, SUM(qty) as qty
     FROM lots WHERE auction_id = ? AND cr LIKE '%GST%' AND amount > 0
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
async function exportPaymentSummary(db, auctionId, cfg) {
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
  const rows = db.all(
    `SELECT name as poolername, lot_no as lot, bags as bag, qty, price, amount,
      pqty, prate, puramt, ${discountCol} as lot_discount, balance as payable
     FROM lots WHERE auction_id = ? AND amount > 0
     ORDER BY state, name`, [auctionId]
  );
  // Spread debit_notes amount across the seller's lots proportionally so
  // every row totals to the same SUM as the payments view. Simpler approach:
  // attribute the FULL manual debit on the FIRST row for each seller; later
  // rows show only the lot policy discount. Avoids per-row arithmetic but
  // still preserves the seller-level total.
  const seenSellers = new Set();
  const enriched = rows.map(r => {
    const lotDisc = Number(r.lot_discount) || 0;
    const manualDisc = (!seenSellers.has(r.poolername))
      ? (Number(debitMap[r.poolername]) || 0)
      : 0;
    seenSellers.add(r.poolername);
    return {
      ...r,
      discount: lotDisc + manualDisc,
      payable: (Number(r.payable) || 0) - manualDisc,
    };
  });
  const cols = [
    { header: 'POOLERNAME', key: 'poolername', width: 30 },
    { header: 'LOT', key: 'lot', width: 8 }, { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 }, { header: 'PRICE', key: 'price', width: 10 },
    { header: 'AMOUNT', key: 'amount', width: 14 }, { header: 'PQTY', key: 'pqty', width: 12 },
    { header: 'PRATE', key: 'prate', width: 10 }, { header: 'PURAMT', key: 'puramt', width: 14 },
    { header: 'DISCOUNT', key: 'discount', width: 14 },
    { header: 'PAYABLE', key: 'payable', width: 14 },
  ];
  return createExcelBuffer('Payment', cols, enriched, {
    db, title: 'Payment Summary', metaLines: auctionMeta(db, auctionId),
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

// ── Export router ────────────────────────────────────────────
const EXPORT_TYPES = {
  lot_slip:       { fn: exportLotSlip,       name: 'LotSlip' },
  lot_slip_after: { fn: exportLotSlipAfter,  name: 'LotSlipAfter' },
  // Served only from the Price Import modal (not listed in the Export
  // Center, which is driven by the frontend EXP_LABELS map).
  pre_trade:      { fn: exportPreTrade,      name: 'PreTradePriceSheet' },
  praman_csv:     { fn: exportPramanCSV,     name: 'eTrade_Praman', ext: 'csv', mime: 'text/csv', needsCfg: true },
  price_list:     { fn: exportPriceList,     name: 'PriceList' },
  bank_payment:   { fn: exportBankPayment,   name: 'BankPayment', needsCfg: true },
  pooler_register:{ fn: exportPoolerRegister,name: 'PoolerRegister' },
  full_file:      { fn: exportFullFile,      name: 'FullFile' },
  collection:     { fn: exportCollection,    name: 'Collection' },
  trade_report:   { fn: exportTradeReport,   name: 'TradeReport' },
  dealer_list:    { fn: exportDealerList,    name: 'DealerList' },
  sales_taxes:    { fn: exportSalesTaxes,    name: 'SalesTaxes' },
  payment:        { fn: exportPaymentSummary,name: 'Payment',        needsCfg: true },
  tally_purchase: { fn: exportTallyPurchase, name: 'TallyPurchase',  needsCfg: true },
};

module.exports = {
  EXPORT_TYPES,
  exportLotSlip, exportLotSlipAfter, exportPreTrade, exportPramanCSV, exportPriceList, exportBankPayment,
  exportPoolerRegister, exportFullFile, exportCollection, exportTradeReport, exportDealerList,
  exportSalesTaxes, exportPaymentSummary, exportTDSReturn, exportTallyPurchase,
  exportSalesJournal, exportPurchaseJournal,
};
