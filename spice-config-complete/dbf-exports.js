/**
 * dbf-exports.js — DBF (FoxPro) format exports
 * 
 * Exports spice-config data in FoxPro-compatible DBF format so the legacy
 * application can continue to consume records during the transition period.
 * 
 * Field types:
 *   C = Character (text)
 *   N = Numeric
 *   D = Date
 *   L = Logical (boolean)
 * 
 * Key DBF rules learned from the previous chat:
 *   - LotNo, grade, pst_code, ppin, litre → store as TEXT (preserves leading zeros)
 *   - Date fields: real DBF `D` (date) type — stored as YYYYMMDD, so FoxPro
 *     and other consumers read them as dates rather than plain text
 *   - Qty: 3 decimal places
 *   - Amount: 2 decimal places
 */

const { DBFFile } = require('dbffile');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const TMP_DIR = path.join(__dirname, 'data', 'tmp-dbf');
if (!fs.existsSync(TMP_DIR)) fs.mkdirSync(TMP_DIR, { recursive: true });

// Parse a stored date into a JS Date (UTC) so it can be written to a real
// DBF `D` (date) field. dbffile serialises `D` values via Date#toISOString,
// so we build the Date in UTC to avoid any timezone day-shift. Accepts the
// ISO `YYYY-MM-DD` we store as well as legacy `DD/MM/YYYY` strings. Returns
// null for blank/unparseable input, which dbffile writes as an empty date.
function toDate(d) {
  if (!d) return null;
  const s = String(d).trim();
  if (!s) return null;
  // ISO format YYYY-MM-DD
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  // Legacy DD/MM/YYYY
  m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
  return null;
}

// Safely trim string values to fit DBF field size
function fit(val, maxLen) {
  if (val === null || val === undefined) return '';
  return String(val).substring(0, maxLen);
}

// Round number to N decimals
function num(val, dec = 2) {
  const n = parseFloat(val);
  if (isNaN(n)) return 0;
  return parseFloat(n.toFixed(dec));
}

// Normalise the filter argument so every exporter accepts either a rich
// filter object ({ auctionId, ano, from, to }) or — for backward
// compatibility — a bare auctionId (string/number). Lets the same module
// be called auction-wise OR date-wise.
function normFilters(f) {
  if (f == null) return {};
  if (typeof f === 'object') return f;
  return { auctionId: f };
}

// Build a WHERE clause for the ano-keyed transactional tables (invoices,
// purchases, bills, debit_notes). Auction-wise wins when an `ano` is
// supplied; otherwise a date range narrows by the `date` column. With
// neither, the whole table is exported.
function anoDateWhere(filters) {
  filters = normFilters(filters);
  if (filters.ano != null && String(filters.ano).trim() !== '') {
    return { sql: 'ano = ?', params: [filters.ano] };
  }
  if (filters.from && filters.to) {
    return { sql: 'date BETWEEN ? AND ?', params: [filters.from, filters.to] };
  }
  return { sql: '1=1', params: [] };
}

/**
 * Build an XLSX buffer for a flat list of master-data rows. Simple
 * branded-free sheet: a bold header row plus one row per record. Used for
 * the Sellers / Buyers "export as .xlsx" option alongside the .dbf form.
 */
async function writeXlsxBuffer(sheetName, columns, records) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName);
  ws.columns = columns.map(c => ({ header: c.header, key: c.key, width: c.width || 18 }));
  const headerRow = ws.getRow(1);
  headerRow.font = { bold: true };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
  headerRow.eachCell((cell) => {
    cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
  });
  records.forEach(rec => ws.addRow(rec));
  return wb.xlsx.writeBuffer();
}

/**
 * Emit the SAME field/record shape that backs a .dbf as a plain .xlsx, so
 * every DBF export is also downloadable as a spreadsheet. Headers are the
 * DBF field names; numeric fields keep their native number type (the records
 * are already keyed by field name, with numbers for N fields).
 */
async function writeXlsxFromFields(sheetName, fields, records) {
  const columns = fields.map(f => ({
    header: f.name,
    key: f.name,
    width: Math.min(40, Math.max(10, (f.size || 12) + 2)),
  }));
  return writeXlsxBuffer(sheetName, columns, records);
}

// Route a built field/record set to the requested output format ('dbf' | 'xlsx').
function emitExport(format, sheetName, fields, records) {
  return format === 'xlsx'
    ? writeXlsxFromFields(sheetName, fields, records)
    : writeDbfBuffer(fields, records);
}

/**
 * Create a DBF file and write records to it
 * Returns a Buffer containing the DBF file contents
 */
async function writeDbfBuffer(fields, records) {
  const tmpFile = path.join(TMP_DIR, `export-${Date.now()}-${Math.random().toString(36).slice(2)}.dbf`);
  try {
    const dbf = await DBFFile.create(tmpFile, fields);
    if (records.length) await dbf.appendRecords(records);
    const buffer = fs.readFileSync(tmpFile);
    return buffer;
  } finally {
    // Cleanup temp file
    try { fs.unlinkSync(tmpFile); } catch(e) {}
  }
}

// ── LOTS (CPA1.DBF structure) ─────────────────────────────────
// Accepts either a bare auctionId (legacy) or a filter object. Auction-wise
// filters on lots.auction_id; date-wise filters on the parent auction's date.
async function exportLotsDbf(db, filters, format = 'dbf') {
  filters = normFilters(filters);
  let where = '1=1';
  const params = [];
  if (filters.auctionId != null && String(filters.auctionId).trim() !== '') {
    where = 'l.auction_id = ?';
    params.push(filters.auctionId);
  } else if (filters.from && filters.to) {
    where = 'a.date BETWEEN ? AND ?';
    params.push(filters.from, filters.to);
  }
  const rows = db.all(`
    SELECT l.*, a.ano as trade_no, a.date as trade_date
    FROM lots l JOIN auctions a ON a.id = l.auction_id
    WHERE ${where} ORDER BY a.date, l.lot_no
  `, params);

  const fields = [
    { name: 'ANO',      type: 'C', size: 10 },
    { name: 'DATE',     type: 'D', size: 8 },
    { name: 'LOT',      type: 'C', size: 10 },
    { name: 'CROP',     type: 'C', size: 10 },
    { name: 'GRADE',    type: 'C', size: 10 },
    { name: 'CRPT',     type: 'C', size: 10 },
    { name: 'BR',       type: 'C', size: 30 },
    { name: 'STATE',    type: 'C', size: 20 },
    { name: 'NAME',     type: 'C', size: 50 },
    { name: 'PADD',     type: 'C', size: 80 },
    { name: 'PPLA',     type: 'C', size: 30 },
    { name: 'PPIN',     type: 'C', size: 10 },
    { name: 'PSTATE',   type: 'C', size: 20 },
    { name: 'PST_CODE', type: 'C', size: 10 },
    { name: 'CR',       type: 'C', size: 40 },
    { name: 'PAN',      type: 'C', size: 14 },
    { name: 'TEL',      type: 'C', size: 20 },
    { name: 'AADHAR',   type: 'C', size: 20 },
    { name: 'BAG',      type: 'N', size: 6, decimalPlaces: 0 },
    { name: 'LITRE',    type: 'C', size: 10 },
    { name: 'QTY',      type: 'N', size: 12, decimalPlaces: 3 },
    { name: 'PRICE',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'AMOUNT',   type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'CODE',     type: 'C', size: 10 },
    { name: 'BUYER',    type: 'C', size: 10 },
    { name: 'BUYER1',   type: 'C', size: 50 },
    { name: 'SALE',     type: 'C', size: 2 },
    { name: 'INVO',     type: 'C', size: 10 },
    { name: 'PQTY',     type: 'N', size: 12, decimalPlaces: 3 },
    { name: 'PRATE',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'PURAMT',   type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'COM',      type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'CGST',     type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'SGST',     type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'IGST',     type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'ADVANCE',  type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'BALANCE',  type: 'N', size: 14, decimalPlaces: 2 },
  ];

  const records = rows.map(r => ({
    ANO: fit(r.trade_no, 10),
    DATE: toDate(r.trade_date),
    LOT: fit(r.lot_no, 10),
    CROP: '',
    GRADE: fit(r.grade, 10),
    CRPT: fit(r.crpt, 10),
    BR: fit(r.branch, 30),
    STATE: fit(r.state, 20),
    NAME: fit(r.name, 50),
    PADD: fit(r.padd, 80),
    PPLA: fit(r.ppla, 30),
    PPIN: fit(r.ppin, 10),
    PSTATE: fit(r.pstate, 20),
    PST_CODE: fit(r.pst_code, 10),
    CR: fit(r.cr, 40),
    PAN: fit(r.pan, 14),
    TEL: fit(r.tel, 20),
    AADHAR: fit(r.aadhar, 20),
    BAG: parseInt(r.bags) || 0,
    LITRE: fit(r.litre, 10),
    QTY: num(r.qty, 3),
    PRICE: num(r.price, 2),
    AMOUNT: num(r.amount, 2),
    CODE: fit(r.code, 10),
    BUYER: fit(r.buyer, 10),
    BUYER1: fit(r.buyer1, 50),
    SALE: fit(r.sale, 2),
    INVO: fit(r.invo, 10),
    PQTY: num(r.pqty, 3),
    PRATE: num(r.prate, 2),
    PURAMT: num(r.puramt, 2),
    COM: num(r.com, 2),
    CGST: num(r.cgst, 2),
    SGST: num(r.sgst, 2),
    IGST: num(r.igst, 2),
    ADVANCE: num(r.advance, 2),
    BALANCE: num(r.balance, 2),
  }));

  return emitExport(format, 'Lots', fields, records);
}

// ── SALES INVOICES (INV.DBF structure) ────────────────────────
async function exportInvoicesDbf(db, filters = {}, format = 'dbf') {
  const w = anoDateWhere(filters);
  const rows = db.all(`SELECT * FROM invoices WHERE ${w.sql} ORDER BY date, sale, invo`, w.params);

  const fields = [
    { name: 'ANO',     type: 'C', size: 10 },
    { name: 'DATE',    type: 'D', size: 8 },
    { name: 'STATE',   type: 'C', size: 20 },
    { name: 'SALE',    type: 'C', size: 2 },
    { name: 'INVO',    type: 'C', size: 10 },
    { name: 'BUYER',   type: 'C', size: 10 },
    { name: 'BUYER1',  type: 'C', size: 50 },
    { name: 'GSTIN',   type: 'C', size: 20 },
    { name: 'PLACE',   type: 'C', size: 30 },
    { name: 'BAG',     type: 'N', size: 6, decimalPlaces: 0 },
    { name: 'QTY',     type: 'N', size: 12, decimalPlaces: 3 },
    { name: 'AMOUNT',  type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'GUNNY',   type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'PAVA_HC', type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'INS',     type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'CGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'SGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'IGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'TCS',     type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'RUND',    type: 'N', size: 8, decimalPlaces: 2 },
    { name: 'TOT',     type: 'N', size: 14, decimalPlaces: 2 },
  ];

  const records = rows.map(r => ({
    ANO: fit(r.ano, 10),
    DATE: toDate(r.date),
    STATE: fit(r.state, 20),
    SALE: fit(r.sale, 2),
    INVO: fit(r.invo, 10),
    BUYER: fit(r.buyer, 10),
    BUYER1: fit(r.buyer1, 50),
    GSTIN: fit(r.gstin, 20),
    PLACE: fit(r.place, 30),
    BAG: parseInt(r.bag) || 0,
    QTY: num(r.qty, 3),
    AMOUNT: num(r.amount, 2),
    GUNNY: num(r.gunny, 2),
    PAVA_HC: num(r.pava_hc, 2),
    INS: num(r.ins, 2),
    CGST: num(r.cgst, 2),
    SGST: num(r.sgst, 2),
    IGST: num(r.igst, 2),
    TCS: num(r.tcs, 2),
    RUND: num(r.rund, 2),
    TOT: num(r.tot, 2),
  }));

  return emitExport(format, 'Sales Invoices', fields, records);
}

// ── PURCHASES (PURCHASE.DBF structure) ────────────────────────
async function exportPurchasesDbf(db, filters = {}, format = 'dbf') {
  const w = anoDateWhere(filters);
  const rows = db.all(`SELECT * FROM purchases WHERE ${w.sql} ORDER BY date, invo`, w.params);

  const fields = [
    { name: 'ANO',     type: 'C', size: 10 },
    { name: 'DATE',    type: 'D', size: 8 },
    { name: 'STATE',   type: 'C', size: 20 },
    { name: 'BR',      type: 'C', size: 30 },
    { name: 'NAME',    type: 'C', size: 50 },
    { name: 'ADD_LINE',type: 'C', size: 80 },
    { name: 'PLACE',   type: 'C', size: 30 },
    { name: 'GSTIN',   type: 'C', size: 40 },
    { name: 'INVO',    type: 'C', size: 10 },
    { name: 'QTY',     type: 'N', size: 12, decimalPlaces: 3 },
    { name: 'AMOUNT',  type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'CGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'SGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'IGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'RUND',    type: 'N', size: 8, decimalPlaces: 2 },
    { name: 'TOTAL',   type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'TDS',     type: 'N', size: 12, decimalPlaces: 2 },
  ];

  const records = rows.map(r => ({
    ANO: fit(r.ano, 10),
    DATE: toDate(r.date),
    STATE: fit(r.state, 20),
    BR: fit(r.br, 30),
    NAME: fit(r.name, 50),
    ADD_LINE: fit(r.add_line, 80),
    PLACE: fit(r.place, 30),
    GSTIN: fit(r.gstin, 40),
    INVO: fit(r.invo, 10),
    QTY: num(r.qty, 3),
    AMOUNT: num(r.amount, 2),
    CGST: num(r.cgst, 2),
    SGST: num(r.sgst, 2),
    IGST: num(r.igst, 2),
    RUND: num(r.rund, 2),
    TOTAL: num(r.total, 2),
    TDS: num(r.tds, 2),
  }));

  return emitExport(format, 'Purchases', fields, records);
}

// ── BILLS of SUPPLY (BILL.DBF structure) ──────────────────────
async function exportBillsDbf(db, filters = {}, format = 'dbf') {
  const w = anoDateWhere(filters);
  const rows = db.all(`SELECT * FROM bills WHERE ${w.sql} ORDER BY date, bil`, w.params);

  const fields = [
    { name: 'ANO',     type: 'C', size: 10 },
    { name: 'DATE',    type: 'D', size: 8 },
    { name: 'STATE',   type: 'C', size: 20 },
    { name: 'BR',      type: 'C', size: 30 },
    { name: 'CRPT',    type: 'C', size: 10 },
    { name: 'BIL',     type: 'N', size: 8, decimalPlaces: 0 },
    { name: 'NAME',    type: 'C', size: 50 },
    { name: 'ADD_LINE',type: 'C', size: 80 },
    { name: 'PLA',     type: 'C', size: 30 },
    { name: 'PSTATE',  type: 'C', size: 20 },
    { name: 'ST_CODE', type: 'C', size: 10 },
    { name: 'CRR',     type: 'C', size: 20 },
    { name: 'PAN',     type: 'C', size: 14 },
    { name: 'QTY',     type: 'N', size: 12, decimalPlaces: 3 },
    { name: 'COST',    type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'IGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'NET',     type: 'N', size: 14, decimalPlaces: 2 },
  ];

  const records = rows.map(r => ({
    ANO: fit(r.ano, 10),
    DATE: toDate(r.date),
    STATE: fit(r.state, 20),
    BR: fit(r.br, 30),
    CRPT: fit(r.crpt, 10),
    BIL: parseInt(r.bil) || 0,
    NAME: fit(r.name, 50),
    ADD_LINE: fit(r.add_line, 80),
    PLA: fit(r.pla, 30),
    PSTATE: fit(r.pstate, 20),
    ST_CODE: fit(r.st_code, 10),
    CRR: fit(r.crr, 20),
    PAN: fit(r.pan, 14),
    QTY: num(r.qty, 3),
    COST: num(r.cost, 2),
    IGST: num(r.igst, 2),
    NET: num(r.net, 2),
  }));

  return emitExport(format, 'Bills of Supply', fields, records);
}

// ── TRADERS / SELLERS (NAM.DBF structure) ─────────────────────
async function exportTradersDbf(db) {
  const rows = db.all('SELECT * FROM traders ORDER BY name');

  const fields = [
    { name: 'NAME',       type: 'C', size: 50 },
    { name: 'CR',         type: 'C', size: 40 },
    { name: 'PAN',        type: 'C', size: 14 },
    { name: 'TEL',        type: 'C', size: 20 },
    { name: 'AADHAR',     type: 'C', size: 20 },
    { name: 'PADD',       type: 'C', size: 80 },
    { name: 'PPLA',       type: 'C', size: 30 },
    { name: 'PIN',        type: 'C', size: 10 },
    { name: 'PSTATE',     type: 'C', size: 20 },
    { name: 'PST_CODE',   type: 'C', size: 10 },
    { name: 'IFSC',       type: 'C', size: 15 },
    { name: 'ACCTNUM',    type: 'C', size: 25 },
    { name: 'HOLDER_NM',  type: 'C', size: 50 },
  ];

  const records = rows.map(r => ({
    NAME: fit(r.name, 50),
    CR: fit(r.cr, 40),
    PAN: fit(r.pan, 14),
    TEL: fit(r.tel, 20),
    AADHAR: fit(r.aadhar, 20),
    PADD: fit(r.padd, 80),
    PPLA: fit(r.ppla, 30),
    PIN: fit(r.pin, 10),
    PSTATE: fit(r.pstate, 20),
    PST_CODE: fit(r.pst_code, 10),
    IFSC: fit(r.ifsc, 15),
    ACCTNUM: fit(r.acctnum, 25),
    HOLDER_NM: fit(r.holder_name, 50),
  }));

  return writeDbfBuffer(fields, records);
}

// ── BUYERS / DEALERS (SBL.DBF structure) ──────────────────────
async function exportBuyersDbf(db) {
  const rows = db.all('SELECT * FROM buyers ORDER BY buyer');

  const fields = [
    { name: 'BUYER',    type: 'C', size: 10 },
    { name: 'BUYER1',   type: 'C', size: 50 },
    { name: 'ADD1',     type: 'C', size: 80 },
    { name: 'ADD2',     type: 'C', size: 80 },
    { name: 'PLA',      type: 'C', size: 30 },
    { name: 'PIN',      type: 'C', size: 10 },
    { name: 'STATE',    type: 'C', size: 20 },
    { name: 'ST_CODE',  type: 'C', size: 10 },
    { name: 'GSTIN',    type: 'C', size: 20 },
    { name: 'PAN',      type: 'C', size: 14 },
    { name: 'TEL',      type: 'C', size: 20 },
    { name: 'TI',       type: 'C', size: 20 },
    { name: 'SALE',     type: 'C', size: 2 },
  ];

  const records = rows.map(r => ({
    BUYER: fit(r.buyer, 10),
    BUYER1: fit(r.buyer1, 50),
    ADD1: fit(r.add1, 80),
    ADD2: fit(r.add2, 80),
    PLA: fit(r.pla, 30),
    PIN: fit(r.pin, 10),
    STATE: fit(r.state, 20),
    ST_CODE: fit(r.st_code, 10),
    GSTIN: fit(r.gstin, 20),
    PAN: fit(r.pan, 14),
    TEL: fit(r.tel, 20),
    TI: fit(r.ti, 20),
    SALE: fit(r.sale, 2) || 'L',
  }));

  return writeDbfBuffer(fields, records);
}

// ── SELLERS / BUYERS as XLSX ──────────────────────────────────
// Same master-data rows as the .dbf forms above, emitted as a plain
// spreadsheet so the data is usable outside the legacy FoxPro app.
async function exportTradersXlsx(db) {
  const rows = db.all('SELECT * FROM traders ORDER BY name');
  const columns = [
    { header: 'Name',           key: 'name',        width: 28 },
    { header: 'CR / GSTIN',     key: 'cr',          width: 24 },
    { header: 'PAN',            key: 'pan',         width: 14 },
    { header: 'Tel',            key: 'tel',         width: 16 },
    { header: 'Aadhar',         key: 'aadhar',      width: 16 },
    { header: 'Address',        key: 'padd',        width: 36 },
    { header: 'Place',          key: 'ppla',        width: 20 },
    { header: 'PIN',            key: 'pin',         width: 10 },
    { header: 'State',          key: 'pstate',      width: 16 },
    { header: 'State Code',     key: 'pst_code',    width: 10 },
    { header: 'IFSC',           key: 'ifsc',        width: 14 },
    { header: 'Account Number', key: 'acctnum',     width: 22 },
    { header: 'Account Holder', key: 'holder_name', width: 26 },
  ];
  const records = rows.map(r => ({
    name: r.name || '', cr: r.cr || '', pan: r.pan || '', tel: r.tel || '',
    aadhar: r.aadhar || '', padd: r.padd || '', ppla: r.ppla || '', pin: r.pin || '',
    pstate: r.pstate || '', pst_code: r.pst_code || '', ifsc: r.ifsc || '',
    acctnum: r.acctnum || '', holder_name: r.holder_name || '',
  }));
  return writeXlsxBuffer('Sellers', columns, records);
}

async function exportBuyersXlsx(db) {
  const rows = db.all('SELECT * FROM buyers ORDER BY buyer');
  const columns = [
    { header: 'Code',       key: 'buyer',    width: 12 },
    { header: 'Name',       key: 'buyer1',   width: 30 },
    { header: 'Address 1',  key: 'add1',     width: 32 },
    { header: 'Address 2',  key: 'add2',     width: 32 },
    { header: 'Place',      key: 'pla',      width: 20 },
    { header: 'PIN',        key: 'pin',      width: 10 },
    { header: 'State',      key: 'state',    width: 16 },
    { header: 'State Code', key: 'st_code',  width: 10 },
    { header: 'GSTIN',      key: 'gstin',    width: 20 },
    { header: 'PAN',        key: 'pan',      width: 14 },
    { header: 'Tel',        key: 'tel',      width: 16 },
    { header: 'TI',         key: 'ti',       width: 16 },
    { header: 'Sale',       key: 'sale',     width: 8 },
  ];
  const records = rows.map(r => ({
    buyer: r.buyer || '', buyer1: r.buyer1 || '', add1: r.add1 || '', add2: r.add2 || '',
    pla: r.pla || '', pin: r.pin || '', state: r.state || '', st_code: r.st_code || '',
    gstin: r.gstin || '', pan: r.pan || '', tel: r.tel || '', ti: r.ti || '',
    sale: r.sale || 'L',
  }));
  return writeXlsxBuffer('Buyers', columns, records);
}

// ── DEBIT NOTES ───────────────────────────────────────────────
async function exportDebitNotesDbf(db, filters = {}, format = 'dbf') {
  const w = anoDateWhere(filters);
  const rows = db.all(`SELECT * FROM debit_notes WHERE ${w.sql} ORDER BY date, note_no`, w.params);

  const fields = [
    { name: 'ANO',     type: 'C', size: 10 },
    { name: 'DATE',    type: 'D', size: 8 },
    { name: 'STATE',   type: 'C', size: 20 },
    { name: 'NAME',    type: 'C', size: 50 },
    { name: 'NOTE_NO', type: 'C', size: 10 },
    { name: 'AMOUNT',  type: 'N', size: 14, decimalPlaces: 2 },
    { name: 'CGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'SGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'IGST',    type: 'N', size: 12, decimalPlaces: 2 },
    { name: 'TOTAL',   type: 'N', size: 14, decimalPlaces: 2 },
  ];

  const records = rows.map(r => ({
    ANO: fit(r.ano, 10),
    DATE: toDate(r.date),
    STATE: fit(r.state, 20),
    NAME: fit(r.name, 50),
    NOTE_NO: fit(r.note_no, 10),
    AMOUNT: num(r.amount, 2),
    CGST: num(r.cgst, 2),
    SGST: num(r.sgst, 2),
    IGST: num(r.igst, 2),
    TOTAL: num(r.total, 2),
  }));

  return emitExport(format, 'Debit Notes', fields, records);
}

// ── Registry for easy routing ─────────────────────────────────
// Capability flags:
//   auctionFilter  → can be narrowed to one trade/auction
//   dateFilter     → can be narrowed to a date range
//   master         → master data (no transactional filter)
//   xlsx           → also exportable as .xlsx. Transactional modules reuse
//                    their .dbf field/record shape (fn called with format
//                    'xlsx'); master modules use a hand-tuned xlsxFn instead.
// Transactional modules carry BOTH auctionFilter and dateFilter so the
// user can export either trade/auction-wise or date-wise.
const DBF_EXPORTS = {
  lots:         { fn: exportLotsDbf,        name: 'CPA1',     auctionFilter: true, dateFilter: true, xlsx: true, label: 'Lots (CPA1.DBF)' },
  invoices:     { fn: exportInvoicesDbf,    name: 'INV',      auctionFilter: true, dateFilter: true, xlsx: true, label: 'Sales Invoices (INV.DBF)' },
  purchases:    { fn: exportPurchasesDbf,   name: 'PURCHASE', auctionFilter: true, dateFilter: true, xlsx: true, label: 'Purchases (PURCHASE.DBF)' },
  bills:        { fn: exportBillsDbf,       name: 'BILL',     auctionFilter: true, dateFilter: true, xlsx: true, label: 'Bills of Supply (BILL.DBF)' },
  debit_notes:  { fn: exportDebitNotesDbf,  name: 'DEBIT',    auctionFilter: true, dateFilter: true, xlsx: true, label: 'Debit Notes' },
  traders:      { fn: exportTradersDbf,     name: 'NAM',      master: true, xlsx: true, xlsxFn: exportTradersXlsx, label: 'Sellers (NAM.DBF)' },
  buyers:       { fn: exportBuyersDbf,      name: 'SBL',      master: true, xlsx: true, xlsxFn: exportBuyersXlsx, label: 'Buyers (SBL.DBF)' },
};

module.exports = {
  DBF_EXPORTS,
  exportLotsDbf,
  exportInvoicesDbf,
  exportPurchasesDbf,
  exportBillsDbf,
  exportTradersDbf,
  exportBuyersDbf,
  exportTradersXlsx,
  exportBuyersXlsx,
  exportDebitNotesDbf,
};
