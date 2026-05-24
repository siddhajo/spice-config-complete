/**
 * price-check.js â Validate SERVER PRICE column in an uploaded price
 * sheet against the live `lots` table.
 *
 * Replaces the legacy PriceCheck_VSTL.xlsm workbook. The macro flow
 * was:
 *
 *   1. `mapping`     â imports the auctioneer's manual record into
 *                      cols A-H (KERALA/TNO/DATE/LOT/BAG/QTY/PRICE/CODE)
 *   2. `priceCheck`  â imports the server's price export into col I
 *                      (SERVER PRICE)
 *   3. `checkRate`   â diffs col G (manual PRICE) vs col I (SERVER
 *                      PRICE), paints non-zero diffs red.
 *
 * In this app the "server" *is* the live SQLite `lots` table, so
 * step 2 â the second Excel import â is the wrong abstraction. What
 * the user actually wants is to validate that the SERVER PRICE in
 * their file (whether they pasted it in by hand, imported from an
 * older export, or got it from somewhere else) still matches what's
 * in the DB *now*. A mismatch can mean either:
 *   - the file was generated before a price correction was made
 *     (file is stale relative to DB)
 *   - the file was hand-edited and is wrong
 * Either way, the operator wants to see those rows.
 *
 * Public API:
 *   - locateColumns(ws)               â column map (header row + indexes)
 *   - processFile(filePath, db, opts) â { wb, ws, cols, perRow, summary }
 *   - annotateWorkbook(wb, ws, cols, perRow)  (mutates ws for download)
 *
 * `opts.auctionId` (optional but recommended) forces every row to
 * match against that auction. Without it, rows are resolved per-row
 * via TNO + DATE â which works but is more permissive (a typo in TNO
 * silently drops a row into `no_auction`).
 */

const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
// SheetJS â used to bridge legacy .xls (BIFF) files. ExcelJS only
// reads modern .xlsx (OOXML); when the operator uploads a .xls (e.g.
// the legacy PriceCheck_VSTL.xls saved by the macro workbook), SheetJS
// reads it and re-emits an .xlsx buffer that ExcelJS can then process
// normally. The downloaded annotated file is always .xlsx.
const XLSX = require('xlsx');

// ââ Normalisation helpers âââââââââââââââââââââââââââââââââââââ

// Tolerant header normaliser â collapses spaces/underscores/dashes
// so "SERVER PRICE" / "SERVER_PRICE" / "server-price" all hash to the
// same key.
const normHeader = (s) =>
  String(s == null ? '' : s).trim().toUpperCase().replace(/[\s_\-]+/g, ' ');

// Strip leading zeros from a lot number so "001" and "1" match.
// We keep the original string as the display lot but use this for
// keyed lookups against the DB (which may store either form).
const normLot = (v) => {
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  if (/^0\d+$/.test(s)) return s.replace(/^0+/, '') || '0';
  return s;
};

// Excel serial date â ISO yyyy-mm-dd. Excel epoch is 1899-12-30
// (account for the 1900 leap-year bug). Accepts Date / number / string.
const toIsoDate = (v) => {
  if (!v) return '';
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  if (typeof v === 'number' && v > 25569) {
    const ms = (v - 25569) * 86400 * 1000;
    return toIsoDate(new Date(ms));
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    let y = m[3].length === 2 ? '20' + m[3] : m[3];
    return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  return s;
};

// Price tolerance â prices are integers in practice but Excel can
// round-trip floats through binary, so half a rupee absorbs noise
// without masking real discrepancies.
const PRICE_TOL = 0.5;

// ââ Column locator ââââââââââââââââââââââââââââââââââââââââââââ

// Scan the first 15 rows for a header row that contains LOT plus a
// server-price column. Any column not found stays 0.
//
// SERVER PRICE is the validation target. PRICE (col G) is optional
// context â preserved if present, ignored if absent. The matcher
// requires LOT + (SERVER PRICE or PRICE) to consider a header valid;
// if only PRICE is present, it's used as the validation target
// (graceful fallback for files that don't have a SERVER PRICE column
// at all).
function locateColumns(ws) {
  const maxRow = ws.rowCount || 0;
  let headerRow = 0;
  let map = {};
  let lastCol = 0;

  for (let r = 1; r <= Math.min(maxRow, 15); r++) {
    const row = ws.getRow(r);
    const cells = {};
    let localLast = 0;
    row.eachCell({ includeEmpty: false }, (cell, col) => {
      const k = normHeader(cell.value);
      if (k) cells[k] = col;
      if (col > localLast) localLast = col;
    });
    const hasLot = cells['LOT'] || cells['LOT NO'] || cells['LOTNO'];
    const hasAnyPrice =
      cells['SERVER PRICE'] || cells['SYSTEM PRICE'] || cells['DB PRICE'] ||
      cells['PRICE'] || cells['RATE'];
    if (hasLot && hasAnyPrice) {
      headerRow = r;
      map = cells;
      lastCol = localLast;
      break;
    }
  }

  const lotCol         = map['LOT']         || map['LOT NO'] || map['LOTNO']  || 0;
  const manualPriceCol = map['PRICE']       || map['RATE']                    || 0;
  const serverPriceCol = map['SERVER PRICE']|| map['SYSTEM PRICE'] || map['DB PRICE'] || 0;
  const codeCol        = map['CODE']        || map['BUYER CODE']              || 0;
  const anoCol         = map['TNO']         || map['ANO']    || map['AUCTION NO'] || map['TRADE'] || map['TRADE NO'] || 0;
  const dateCol        = map['DATE']                                          || 0;
  const bagCol         = map['BAG']         || map['BAGS']                    || 0;
  const qtyCol         = map['QTY']         || map['QUANTITY']  || map['WEIGHT'] || map['WT'] || 0;
  const stateCol       = map['KERALA']      || map['STATE']                   || 0;
  const diffCol        = map['DIFF']                                          || 0;

  // The validation target: prefer SERVER PRICE, fall back to PRICE
  // (so files that pre-date the SERVER PRICE column still work).
  const validateCol = serverPriceCol || manualPriceCol;
  const hasServerPriceCol = !!serverPriceCol;

  return {
    headerRow, lotCol, manualPriceCol, serverPriceCol, validateCol,
    hasServerPriceCol,
    codeCol, anoCol, dateCol, bagCol, qtyCol, stateCol, diffCol,
    lastDataCol: lastCol,
  };
}

// ââ Cell readers ââââââââââââââââââââââââââââââââââââââââââââââ

function readCell(row, col) {
  if (!col) return null;
  const v = row.getCell(col).value;
  if (v == null) return null;
  if (typeof v === 'object') {
    if (v.result !== undefined) return v.result;
    if (v.text !== undefined)   return v.text;
    if (v.richText)             return v.richText.map(t => t.text).join('');
  }
  return v;
}
function readStr(row, col) {
  const v = readCell(row, col);
  return v == null ? '' : String(v).trim();
}
function readNum(row, col) {
  const v = readCell(row, col);
  if (v == null || v === '') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

// ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
// processFile â read XLSX, compare every row's SERVER PRICE against
// the lots table, build perRow + summary. Does NOT mutate the
// workbook.
// ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

async function processFile(filePath, db, opts) {
  opts = opts || {};
  const wb = new ExcelJS.Workbook();
  // Detect legacy .xls and bridge through SheetJS. ExcelJS's
  // `.xlsx.readFile` doesn't understand the older binary BIFF
  // format â it opens the file without throwing but returns a
  // workbook with zero worksheets, which surfaced as the cryptic
  // "No worksheet found in uploaded file" error. SheetJS handles
  // both formats transparently, so we read the file there and
  // re-serialize as an .xlsx buffer that ExcelJS can consume.
  const ext = String(path.extname(filePath || '') || '').toLowerCase();
  if (ext === '.xls') {
    let xlsxBuf;
    try {
      const inputBuf = fs.readFileSync(filePath);
      const sjsWb = XLSX.read(inputBuf, { type: 'buffer', cellDates: true });
      xlsxBuf = XLSX.write(sjsWb, { bookType: 'xlsx', type: 'buffer' });
    } catch (e) {
      throw new Error('Could not read .xls file (' + (e.message || 'unknown error') + '). Try opening it in Excel and saving as .xlsx, then re-upload.');
    }
    await wb.xlsx.load(xlsxBuf);
  } else {
    await wb.xlsx.readFile(filePath);
  }
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('No worksheet found in uploaded file. If you exported as .xls, save the file as .xlsx and try again.');

  const cols = locateColumns(ws);
  if (!cols.headerRow) {
    throw new Error('Could not find a header row with LOT and SERVER PRICE (or PRICE) columns. Check that the file uses the standard Price Check layout (KERALA, TNO, DATE, LOT, BAG, QTY, PRICE, CODE, SERVER PRICE).');
  }
  if (!cols.validateCol) {
    throw new Error('No SERVER PRICE or PRICE column found in the header â nothing to validate.');
  }

  // Resolve auctions once. Two modes:
  //   1. opts.auctionId set â every row uses it. The UI pre-picks
  //      an auction, the file then doesn't need TNO/DATE columns,
  //      and we can also enumerate db-only lots (missing_file).
  //   2. opts.auctionId unset â resolve per row from TNO + DATE.
  //      missing_file is not reported (we don't know which auctions
  //      to enumerate).
  const forcedAuctionId = opts.auctionId ? Number(opts.auctionId) : null;
  let forcedAuction = null;
  let lotIdx = null;

  if (forcedAuctionId) {
    forcedAuction = db.get('SELECT * FROM auctions WHERE id = ?', [forcedAuctionId]);
    if (!forcedAuction) throw new Error('Selected auction not found.');
    const lots = db.all(
      'SELECT id, lot_no, price, code FROM lots WHERE auction_id = ?',
      [forcedAuctionId]
    );
    lotIdx = new Map();
    for (const l of lots) lotIdx.set(normLot(l.lot_no), l);
  }

  const auctionCache = new Map();
  const auctionLotIdx = new Map();
  const resolveAuctionByRow = (anoStr, dateStr) => {
    const key = `${anoStr}|${dateStr}`;
    if (auctionCache.has(key)) return auctionCache.get(key);
    let auc = null;
    if (anoStr && dateStr) {
      auc = db.get('SELECT * FROM auctions WHERE ano = ? AND date = ?', [anoStr, dateStr]);
    }
    if (!auc && anoStr) {
      auc = db.get('SELECT * FROM auctions WHERE ano = ? ORDER BY date DESC LIMIT 1', [anoStr]);
    }
    auctionCache.set(key, auc);
    return auc;
  };
  const getLotIdx = (auctionId) => {
    if (auctionLotIdx.has(auctionId)) return auctionLotIdx.get(auctionId);
    const lots = db.all(
      'SELECT id, lot_no, price, code FROM lots WHERE auction_id = ?',
      [auctionId]
    );
    const idx = new Map();
    for (const l of lots) idx.set(normLot(l.lot_no), l);
    auctionLotIdx.set(auctionId, idx);
    return idx;
  };

  // Helper: classify the buyer code on each matched row.
  // 'match'      â file CODE matches DB CODE (case-insensitive)
  // 'diff'       â both sides have CODE but they disagree
  // 'file_blank' â file CODE blank, DB has one
  // 'db_blank'   â DB CODE blank, file has one (correct-able)
  // 'both_blank' â neither side has CODE (no action needed)
  function classifyCode(fileCode, dbCode) {
    const f = String(fileCode || '').trim();
    const d = String(dbCode   || '').trim();
    if (!f && !d) return 'both_blank';
    if (!f)        return 'file_blank';
    if (!d)        return 'db_blank';
    return f.toUpperCase() === d.toUpperCase() ? 'match' : 'diff';
  }

  const perRow = [];
  const maxRow = ws.rowCount || 0;
  const seenLotsByAuction = new Map();

  for (let r = cols.headerRow + 1; r <= maxRow; r++) {
    const row = ws.getRow(r);

    const lotRaw   = readStr(row, cols.lotCol);
    // In fallback mode (no separate SERVER PRICE column), the file's
    // PRICE col IS the validated col â they're literally the same cell.
    // Surfacing "manual_price" in that case would just duplicate
    // file_server_price in every row of the response. Null it instead
    // so the UI can hide the redundant column.
    const manualPrice = cols.hasServerPriceCol
      ? readNum(row, cols.manualPriceCol)
      : null;
    const filePrice   = readNum(row, cols.validateCol);
    const anoStr      = readStr(row, cols.anoCol);
    const dateStr     = toIsoDate(readCell(row, cols.dateCol));
    const codeStr     = readStr(row, cols.codeCol);
    const bagFile     = readNum(row, cols.bagCol);
    const qtyFile     = readNum(row, cols.qtyCol);
    const stateStr    = readStr(row, cols.stateCol);

    // Blank/footer row â skip silently (XLSM has a TOTAL row at the
    // bottom of col C with no lot/price).
    if (!lotRaw && manualPrice == null && filePrice == null && !codeStr) continue;
    if (!lotRaw) {
      perRow.push({
        row: r, lot: '', state: stateStr, ano: anoStr, date: dateStr,
        bag: bagFile, qty: qtyFile, code: codeStr,
        manual_price: manualPrice, file_server_price: filePrice,
        db_price: null, diff: null,
        file_code: codeStr, db_code: '', code_status: 'both_blank', lot_id: null,
        status: 'blank', issues: ['no_lot'],
      });
      continue;
    }

    let auction, idx;
    if (forcedAuctionId) {
      auction = forcedAuction;
      idx = lotIdx;
    } else {
      auction = resolveAuctionByRow(anoStr, dateStr);
      idx = auction ? getLotIdx(auction.id) : null;
    }

    if (!auction) {
      perRow.push({
        row: r, lot: lotRaw, state: stateStr, ano: anoStr, date: dateStr,
        bag: bagFile, qty: qtyFile, code: codeStr,
        manual_price: manualPrice, file_server_price: filePrice,
        db_price: null, diff: null,
        file_code: codeStr, db_code: '', code_status: classifyCode(codeStr, ''),
        lot_id: null,
        status: 'no_auction', issues: ['no_auction'],
      });
      continue;
    }

    const lot = idx.get(normLot(lotRaw));
    if (!seenLotsByAuction.has(auction.id)) seenLotsByAuction.set(auction.id, new Set());
    seenLotsByAuction.get(auction.id).add(normLot(lotRaw));

    if (!lot) {
      perRow.push({
        row: r, lot: lotRaw, state: stateStr, ano: anoStr, date: dateStr,
        auction_id: auction.id, auction_ano: auction.ano,
        bag: bagFile, qty: qtyFile, code: codeStr,
        manual_price: manualPrice, file_server_price: filePrice,
        db_price: null, diff: null,
        file_code: codeStr, db_code: '', code_status: classifyCode(codeStr, ''),
        lot_id: null,
        status: 'missing_server', issues: ['missing_server'],
      });
      continue;
    }

    // The validation. File's SERVER PRICE (or PRICE, in fallback
    // mode) vs lots.price. An empty file SERVER PRICE is treated as
    // its own status, not as 0 â "blank" and "0" are semantically
    // different (withdrawn lots legitimately have price 0).
    const dbPrice = Number(lot.price) || 0;
    let status, issues;
    if (filePrice == null) {
      status = 'server_empty';
      issues = ['server_empty'];
    } else if (Math.abs(filePrice - dbPrice) > PRICE_TOL) {
      status = 'server_diff';
      issues = ['server_diff'];
    } else {
      status = 'match';
      issues = [];
    }
    const diff = (filePrice == null) ? null : Math.round((filePrice - dbPrice) * 100) / 100;

    perRow.push({
      row: r, lot: lotRaw, state: stateStr, ano: anoStr, date: dateStr,
      auction_id: auction.id, auction_ano: auction.ano,
      bag: bagFile, qty: qtyFile, code: codeStr,
      manual_price: manualPrice, file_server_price: filePrice,
      db_price: dbPrice, diff,
      file_code: codeStr,
      db_code: String(lot.code || '').trim(),
      code_status: classifyCode(codeStr, lot.code),
      lot_id: lot.id,
      status, issues,
    });
  }

  // Missing-in-file pass â only meaningful when an auction is forced.
  if (forcedAuctionId && lotIdx) {
    const seen = seenLotsByAuction.get(forcedAuctionId) || new Set();
    for (const [key, lot] of lotIdx) {
      if (seen.has(key)) continue;
      perRow.push({
        row: null, lot: lot.lot_no,
        state: forcedAuction.state || '',
        ano: forcedAuction.ano, date: forcedAuction.date,
        auction_id: forcedAuction.id, auction_ano: forcedAuction.ano,
        bag: null, qty: null, code: '',
        manual_price: null, file_server_price: null,
        db_price: Number(lot.price) || 0, diff: null,
        file_code: '', db_code: String(lot.code || '').trim(),
        code_status: classifyCode('', lot.code), lot_id: lot.id,
        status: 'missing_file', issues: ['missing_file'],
      });
    }
  }

  let matched = 0, mismatched = 0, missingServer = 0, missingFile = 0,
      serverEmpty = 0, blank = 0, noAuction = 0;
  let codeMatched = 0, codeMismatched = 0, codeFileBlank = 0,
      codeDbBlank = 0, codeBothBlank = 0;
  // Withdrawn lots â sellers/operators write "WD" into the buyer CODE
  // column on the auction sheet when a lot is pulled before sale. Counting
  // these separately means a clean price check can co-exist with a few
  // legitimately-blank prices (withdrawn lots have no buyer / no price).
  let withdrawnFile = 0, withdrawnDb = 0;
  let totalAbsDiff = 0, totalSignedDiff = 0;
  const isWD = v => String(v || '').trim().toUpperCase() === 'WD';
  for (const e of perRow) {
    switch (e.status) {
      case 'match':          matched++;        break;
      case 'server_diff':    mismatched++;     break;
      case 'server_empty':   serverEmpty++;    break;
      case 'missing_server': missingServer++;  break;
      case 'missing_file':   missingFile++;    break;
      case 'blank':          blank++;          break;
      case 'no_auction':     noAuction++;      break;
    }
    switch (e.code_status) {
      case 'match':      codeMatched++;     break;
      case 'diff':       codeMismatched++;  break;
      case 'file_blank': codeFileBlank++;   break;
      case 'db_blank':   codeDbBlank++;     break;
      case 'both_blank': codeBothBlank++;   break;
    }
    if (isWD(e.file_code)) withdrawnFile++;
    if (isWD(e.db_code))   withdrawnDb++;
    if (e.diff != null) {
      totalAbsDiff += Math.abs(e.diff);
      totalSignedDiff += e.diff;
    }
  }

  // Gate ready = no code mismatches + no DB-blank rows (those are the
  // two buckets the Apply button can fix). File-blank rows are NOT
  // counted â the operator was already warned via the inline note that
  // they need to be fixed manually via Lots â Set Buyer Code.
  const codeFixesPending = codeMismatched + codeDbBlank;
  const gateReady = !!forcedAuction && codeFixesPending === 0;
  const summary = {
    total: perRow.length,
    matched, mismatched, missingServer, missingFile, serverEmpty,
    blank, noAuction,
    codeMatched, codeMismatched, codeFileBlank, codeDbBlank, codeBothBlank,
    codeFixesPending, gateReady,
    withdrawnFile, withdrawnDb,
    totalAbsDiff:    Math.round(totalAbsDiff * 100) / 100,
    totalSignedDiff: Math.round(totalSignedDiff * 100) / 100,
    hasServerPriceCol: cols.hasServerPriceCol,
    hasCodeCol: !!cols.codeCol,
    validatedColumn: cols.hasServerPriceCol ? 'SERVER PRICE' : 'PRICE',
    forcedAuction: forcedAuction
      ? { id: forcedAuction.id, ano: forcedAuction.ano, date: forcedAuction.date, state: forcedAuction.state }
      : null,
  };

  return { wb, ws, cols, perRow, summary };
}

// ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
// annotateWorkbook â round-trip the uploaded XLSX with a DB PRICE
// column added (or repopulated) and a DIFF column showing where the
// file's SERVER PRICE disagrees with the DB. Mirrors the VBA
// checkRate macro's visual style: red on mismatches, full-row red
// on missing-in-server, amber row for missing-in-file.
// ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

const RED   = 'FFFCA5A5';
const AMBER = 'FFFED7AA';
const GREEN = 'FFBBF7D0';
const HEAD  = 'FF14532D';

function annotateWorkbook(wb, ws, cols, perRow) {
  // We add (or reuse) two columns at the right edge of the data:
  //   - DB PRICE  â the value we're validating against
  //   - DIFF      â file SERVER PRICE minus DB price
  // The existing SERVER PRICE column in the file is never overwritten.
  let dbCol, diffCol;
  if (cols.diffCol) {
    diffCol = cols.diffCol;
    dbCol = diffCol - 1;   // legacy layout â DB price sits left of DIFF
  } else {
    dbCol = (cols.lastDataCol || cols.validateCol) + 1;
    diffCol = dbCol + 1;
    const dbHead   = ws.getRow(cols.headerRow).getCell(dbCol);
    const diffHead = ws.getRow(cols.headerRow).getCell(diffCol);
    dbHead.value   = 'DB PRICE';
    diffHead.value = 'DIFF';
    // ExcelJS shares style objects across cells â assigning to
    // cell.fill / .font / .border directly mutates other cells.
    // Use the .style spread to force a new style object.
    [dbHead, diffHead].forEach((cell) => {
      cell.style = {
        ...cell.style,
        font:      { bold: true, size: 10, color: { argb: 'FFFFFFFF' } },
        fill:      { type: 'pattern', pattern: 'solid', fgColor: { argb: HEAD } },
        alignment: { horizontal: 'center', vertical: 'middle' },
        border:    { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } },
      };
    });
    ws.getColumn(dbCol).width = 12;
    ws.getColumn(diffCol).width = 10;
  }

  const byRow = new Map();
  let nextSyntheticRow = (ws.rowCount || cols.headerRow) + 1;
  for (const e of perRow) {
    if (e.row == null) {
      // Missing-in-file rows â append below the data.
      const r = nextSyntheticRow++;
      e.row = r;
      const newRow = ws.getRow(r);
      if (cols.stateCol)  newRow.getCell(cols.stateCol).value  = e.state || '';
      if (cols.anoCol)    newRow.getCell(cols.anoCol).value    = e.ano || '';
      if (cols.dateCol)   newRow.getCell(cols.dateCol).value   = e.date || '';
      if (cols.lotCol)    newRow.getCell(cols.lotCol).value    = e.lot || '';
      if (cols.bagCol)    newRow.getCell(cols.bagCol).value    = e.bag;
      if (cols.qtyCol)    newRow.getCell(cols.qtyCol).value    = e.qty;
    }
    byRow.set(e.row, e);
  }

  const paint = (rowNum, col, argb) => {
    if (!col) return;
    const cell = ws.getRow(rowNum).getCell(col);
    cell.style = { ...cell.style, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb } } };
  };

  for (const e of byRow.values()) {
    const row = ws.getRow(e.row);

    const dbCell = row.getCell(dbCol);
    dbCell.value = e.db_price == null ? '' : e.db_price;
    dbCell.numFmt = '0';
    dbCell.alignment = { horizontal: 'right', vertical: 'middle' };

    const diffCell = row.getCell(diffCol);
    diffCell.value = e.diff == null ? '' : e.diff;
    diffCell.numFmt = '0';
    diffCell.alignment = { horizontal: 'right', vertical: 'middle' };

    const status = e.status;
    const paintRow = (argb) => {
      const lastCol = Math.max(diffCol, cols.lastDataCol || diffCol);
      for (let c = 1; c <= lastCol; c++) paint(e.row, c, argb);
    };

    if (status === 'match') {
      paint(e.row, diffCol, GREEN);
    } else if (status === 'server_diff') {
      // Paint the file's SERVER PRICE cell (the value being
      // challenged) and the DIFF cell.
      paint(e.row, cols.validateCol, RED);
      paint(e.row, diffCol, RED);
    } else if (status === 'server_empty') {
      // File's SERVER PRICE was blank â amber on the price + DIFF
      // cells so the gap is obvious without shouting.
      paint(e.row, cols.validateCol, AMBER);
      paint(e.row, diffCol, AMBER);
    } else if (status === 'missing_server') {
      paintRow(RED);
    } else if (status === 'missing_file') {
      paintRow(AMBER);
    } else {
      paintRow(AMBER);
    }
  }

  // Total row two below the last data row, with SUM(DIFF). Net zero
  // is the success signal; non-zero tells the operator the file has
  // drifted from the DB by that amount in total.
  const lastDataRow = Math.max(...Array.from(byRow.keys()), cols.headerRow);
  const totalRow = ws.getRow(lastDataRow + 2);
  if (cols.codeCol) {
    const lbl = totalRow.getCell(cols.codeCol);
    lbl.value = 'TOTAL';
    lbl.font = { bold: true };
    lbl.alignment = { horizontal: 'right', vertical: 'middle' };
  }
  const dfTot = totalRow.getCell(diffCol);
  dfTot.value = { formula: `SUM(${colLetter(diffCol)}${cols.headerRow + 1}:${colLetter(diffCol)}${lastDataRow})` };
  dfTot.numFmt = '0';
  dfTot.font = { bold: true };
}

function colLetter(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

module.exports = {
  locateColumns,
  processFile,
  annotateWorkbook,
  _normLot: normLot,
  _toIsoDate: toIsoDate,
};
