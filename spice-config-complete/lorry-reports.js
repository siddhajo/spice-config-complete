/**
 * lorry-reports.js — Lorry side-menu reports (XLSX + PDF)
 *
 * Three reports based on dad's FoxPro outputs:
 *   1. Lot Slip (Code)   — per-lot listing with bidder code, carbon-copy
 *      layout in PDF (two columns side-by-side per page)
 *   2. Truck List        — buyer-code summary: lot count, bag, qty, amount
 *   3. Buyer Lot Lorry   — lots grouped by Inter-State / Intra-State
 *      buyer, with full GSTIN headers and per-buyer subtotals
 *
 * State classification rule:
 *   intra-state when UPPER(buyer.state) === UPPER(auction.state)
 *   else inter-state. Buyers with no state info are bucketed as intra-state
 *   (matches the FoxPro fallback — local until proven otherwise).
 */

const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const {
  fmtMoney, fmtQty, fmtPrice,
  getCompanyHeader, drawCompanyHeader,
  writeXlsxCompanyHeader,
} = require('./report-formatters');

// ── Helpers ──────────────────────────────────────────────────
function fmtDateDMY(iso) {
  if (!iso) return '';
  const s = String(iso);
  if (s.includes('-') && s.length >= 10) return s.slice(0, 10).split('-').reverse().join('/');
  return s;
}
// fmtMoney / fmtQty / fmtPrice come from report-formatters.js (Indian comma
// grouping; 2 decimals for rupees, 3 for kilos).

// Manually truncate `text` so doc.widthOfString(out) <= maxWidth, appending an
// ellipsis when truncated. PDFKit 0.15's `lineBreak: false` + `ellipsis: true`
// is unreliable for long single tokens — we ellipsize ourselves so multi-word
// names like "EMPEROR SPICES PRIVATE LIMITED" never wrap into the next row.
// Caller must already have set the desired font/size on `doc`.
function fitText(doc, text, maxWidth) {
  const s = String(text == null ? '' : text);
  if (!s) return '';
  if (doc.widthOfString(s) <= maxWidth) return s;
  const ell = '…';
  const ellW = doc.widthOfString(ell);
  if (ellW >= maxWidth) return '';
  let lo = 0, hi = s.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (doc.widthOfString(s.slice(0, mid)) + ellW <= maxWidth) lo = mid;
    else hi = mid - 1;
  }
  return s.slice(0, lo).trimEnd() + ell;
}

// Wrap `text` into one or more lines fitting maxWidth (word-aware, with
// character-level fallback for over-long tokens). Returns array of lines.
function wrapText(doc, text, maxWidth) {
  const s = String(text == null ? '' : text).trim();
  if (!s) return [''];
  if (doc.widthOfString(s) <= maxWidth) return [s];
  const words = s.split(/\s+/);
  const lines = [];
  let cur = '';
  for (const w of words) {
    const probe = cur ? cur + ' ' + w : w;
    if (doc.widthOfString(probe) <= maxWidth) { cur = probe; continue; }
    if (cur) { lines.push(cur); cur = ''; }
    if (doc.widthOfString(w) > maxWidth) {
      let chunk = '';
      for (const ch of w) {
        if (doc.widthOfString(chunk + ch) <= maxWidth) chunk += ch;
        else { if (chunk) lines.push(chunk); chunk = ch; }
      }
      cur = chunk;
    } else {
      cur = w;
    }
  }
  if (cur) lines.push(cur);
  return lines.length ? lines : [''];
}

// Pull auction header (ano + date + state + mode) for subtitle.
// `mode` lets report headers adapt their "e-TRADE No:" vs
// "e-AUCTION No:" label per-trade. Empty/legacy mode falls back to
// e-AUCTION wording (matches historical reports).
function getAuctionHeader(db, auctionId) {
  const a = db.get('SELECT id, ano, date, crop_type, state, mode FROM auctions WHERE id = ?', [auctionId]);
  if (!a) throw new Error('Auction not found');
  return a;
}
// Mode-aware label helper for the report headers below.
function modeHdrLabel(auction) {
  return (auction && auction.mode === 'e-Trade') ? 'e-TRADE' : 'e-AUCTION';
}

// Pull lots joined with buyer master to get state + GSTIN + firm name.
// We try lots.buyer first (full code, set during price import) and fall back
// to lots.code (short alias) so reports still work for older imports where
// only the short code was captured.
function getLotsWithBuyer(db, auctionId) {
  return db.all(`
    SELECT
      l.lot_no                       AS lot,
      l.bags                         AS bag,
      l.qty                          AS qty,
      l.price                        AS price,
      l.amount                       AS amount,
      l.code                         AS lot_code,
      l.buyer                        AS lot_buyer,
      l.buyer1                       AS lot_buyer1,
      COALESCE(b.code,    l.code,    '') AS code,
      COALESCE(b.buyer,   l.buyer,   '') AS buyer,
      COALESCE(b.buyer1,  l.buyer1,  '') AS buyer1,
      COALESCE(b.sbl,     '')        AS sbl,
      COALESCE(b.gstin,   '')        AS gstin,
      COALESCE(b.state,   '')        AS state,
      COALESCE(b.st_code, '')        AS st_code
    FROM lots l
    LEFT JOIN buyers b
      ON UPPER(TRIM(b.code))  = UPPER(TRIM(l.code))
      OR UPPER(TRIM(b.buyer)) = UPPER(TRIM(l.buyer))
    WHERE l.auction_id = ?
      AND l.amount > 0
    ORDER BY CAST(l.lot_no AS INTEGER), l.lot_no
  `, [auctionId]);
}

// ════════════════════════════════════════════════════════════
// REPORT 1 — LOT SLIP (CODE)
// ════════════════════════════════════════════════════════════
function getLotSlipRows(db, auctionId) {
  return db.all(`
    SELECT
      l.lot_no AS lot,
      l.bags   AS bags,
      l.qty    AS kilos,
      l.price  AS price,
      COALESCE(NULLIF(l.code,''), '') AS bidder
    FROM lots l
    WHERE l.auction_id = ?
      AND l.amount > 0
    ORDER BY CAST(l.lot_no AS INTEGER), l.lot_no
  `, [auctionId]);
}

async function lotSlipCodeXlsx(db, auctionId) {
  const auction = getAuctionHeader(db, auctionId);
  const rows    = getLotSlipRows(db, auctionId);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('LotSlipCode');
  ws.columns = [
    { width: 8 },  // Lot
    { width: 8 },  // Bags
    { width: 12 }, // Kilos
    { width: 12 }, // Price
    { width: 14 }, // Bidder
    { width: 8 },  // Lot (repeated, for matching the PDF carbon-copy layout)
  ];

  // Three-column brand band.
  const startRow = writeXlsxCompanyHeader(wb, ws, getCompanyHeader(db), {
    colCount: 6,
    title: 'LOT SLIP CODE',
    metaLines: [
      `${modeHdrLabel(auction)} No: ${auction.ano}`,
      `Date: ${fmtDateDMY(auction.date)}`,
    ],
  });

  const head = ws.getRow(startRow);
  ['Lot', 'Bags', 'Kilos', 'Price', 'Bidder', 'Lot'].forEach((h, i) => {
    head.getCell(i + 1).value = h;
  });
  head.font = { bold: true };
  head.eachCell((c, ci) => {
    if (ci > 6) return;
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
    c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    c.alignment = { horizontal: 'center' };
  });

  let totBag = 0, totKilo = 0;
  rows.forEach(r => {
    const row = ws.addRow([
      r.lot,
      r.bags,
      Number(r.kilos),
      Number(r.price),
      r.bidder,
      r.lot,           // repeat the lot number on the right (matches PDF)
    ]);
    row.getCell(2).alignment = { horizontal: 'right' };
    row.getCell(3).alignment = { horizontal: 'right' };
    row.getCell(3).numFmt = '#,##0.000';
    row.getCell(4).alignment = { horizontal: 'right' };
    row.getCell(4).numFmt = '#,##0.00';
    row.getCell(5).alignment = { horizontal: 'center' };
    row.getCell(6).alignment = { horizontal: 'center' };
    totBag  += Number(r.bags)  || 0;
    totKilo += Number(r.kilos) || 0;
  });

  // Total row — Bags and Kilos totals, no Price/Bidder/Lot totals
  // (matches the PDF which only sums BAG and QTY).
  const tot = ws.addRow(['Total', totBag, totKilo, '', '', '']);
  tot.font = { bold: true };
  tot.getCell(3).numFmt = '#,##0.000';
  tot.eachCell(c => {
    c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
  });

  return wb.xlsx.writeBuffer();
}

// PDF — carbon-copy layout (two identical halves side-by-side per page)
async function lotSlipCodePdf(db, auctionId) {
  const auction = getAuctionHeader(db, auctionId);
  const rows    = getLotSlipRows(db, auctionId);

  const doc = new PDFDocument({ size: 'A4', layout: 'portrait', margin: 18 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));

  const pageW = doc.page.width;
  const pageH = doc.page.height;
  const m = 18;
  const gutter = 40;
  const halfW = (pageW - m * 2 - gutter) / 2;

  // 6 cols per half: Lot | Bags | Kilos | Price | Bidder | Lot
  const colW = [
    Math.floor(halfW * 0.13),
    Math.floor(halfW * 0.13),
    Math.floor(halfW * 0.20),
    Math.floor(halfW * 0.18),
    Math.floor(halfW * 0.22),
  ];
  // Last col absorbs rounding
  colW.push(halfW - colW.reduce((s, w) => s + w, 0));

  const ROW_H  = 14;
  const HEAD_H = 16;
  const TOP_H  = 50;   // reserved for company header lines per half
  const BODY_TOP = m + TOP_H;
  const BODY_MAX_Y = pageH - m - 8;
  const ROWS_PER_HALF = Math.floor((BODY_MAX_Y - BODY_TOP - HEAD_H) / ROW_H);

  let pageNum = 1;
  let totalPages = 1;
  // Pre-compute total pages so the per-half "Page X" footer is accurate.
  // Each printed page contains ROWS_PER_HALF rows (carbon-copy: same rows
  // appear in both halves so users get an original + a duplicate slip).
  totalPages = Math.max(1, Math.ceil(rows.length / ROWS_PER_HALF));

  // Resolve company branding once. The carbon-copy halves are narrow, so the
  // brand band uses a small (22pt) logo and skips the multi-line address —
  // only the company name fits.
  const companyHeader = getCompanyHeader(db);

  function drawHalfHeader(xOrigin, page) {
    // Compact company brand band (no title/meta — too narrow for three cols).
    const afterY = drawCompanyHeader(doc, companyHeader, {
      x: xOrigin, y: m, width: halfW,
      logoH: 22, logoW: 22, showAddress: false,
    });
    // Page / e-TRADE / Date metadata stacks beneath the brand band.
    doc.font('Helvetica').fontSize(8).fillColor('#000')
       .text(`Page: ${page}`, xOrigin, afterY, { width: halfW, align: 'right' });
    doc.font('Helvetica-Bold').fontSize(9)
       .text(`${modeHdrLabel(auction)} No: ${auction.ano}`, xOrigin, afterY + 12, { width: halfW / 2, align: 'left' });
    doc.text(`Date: ${fmtDateDMY(auction.date)}`, xOrigin + halfW / 2, afterY + 12, { width: halfW / 2, align: 'right' });

    // Column header row
    const hy = BODY_TOP;
    doc.rect(xOrigin, hy, halfW, HEAD_H).fillAndStroke('#E8E4DD', '#444');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(8);
    const headers = ['Lot', 'Bags', 'Kilos', 'Price', 'Bidder', 'Lot'];
    let cx = xOrigin;
    headers.forEach((h, i) => {
      doc.text(h, cx + 2, hy + 4, { width: colW[i] - 4, align: 'center', lineBreak: false });
      cx += colW[i];
    });
  }

  function drawHalfRows(xOrigin, sliceRows, isLastPage) {
    let ry = BODY_TOP + HEAD_H;
    const tblTop = BODY_TOP;
    // Pass 1 — zebra fills first, so the per-row separators below stay
    // visible (a shaded row's fill would otherwise paint over the bottom
    // rule of the unshaded row above it, hiding every other line).
    let fy = ry;
    sliceRows.forEach((r, i) => {
      if (i % 2 === 1) doc.rect(xOrigin, fy, halfW, ROW_H).fill('#F7F5F2');
      fy += ROW_H;
    });

    // Pass 2 — cell text + horizontal separators, drawn on top of the fills.
    sliceRows.forEach((r, i) => {
      doc.fillColor('#000').font('Helvetica').fontSize(8);
      const cells = [
        String(r.lot),
        String(r.bags || 0),
        fmtQty(r.kilos),
        fmtPrice(r.price),
        String(r.bidder || ''),
        String(r.lot),
      ];
      let cx = xOrigin;
      cells.forEach((v, ci) => {
        const isNum = ci === 1 || ci === 2 || ci === 3;
        const fitted = fitText(doc, v, colW[ci] - 4);
        doc.text(fitted, cx + 2, ry + 3, {
          width: colW[ci] - 4,
          align: isNum ? 'right' : 'center',
          lineBreak: false,
        });
        cx += colW[ci];
      });
      // Row separator
      doc.moveTo(xOrigin, ry + ROW_H).lineTo(xOrigin + halfW, ry + ROW_H)
         .lineWidth(0.25).strokeColor('#999').stroke();
      ry += ROW_H;
    });

    // Vertical column separators — one line at every interior column boundary,
    // running from the top of the header down through every body row.
    let vx = xOrigin;
    for (let ci = 0; ci < colW.length - 1; ci++) {
      vx += colW[ci];
      doc.moveTo(vx, tblTop).lineTo(vx, ry)
         .lineWidth(0.25).strokeColor('#888').stroke();
    }

    // Outer border around the whole half-table area
    const tblH = ry - tblTop;
    doc.rect(xOrigin, tblTop, halfW, tblH).lineWidth(0.5).strokeColor('#444').stroke();

    // Grand total on the last page
    if (isLastPage) {
      const totBag  = rows.reduce((s, r) => s + (Number(r.bags)  || 0), 0);
      const totKilo = rows.reduce((s, r) => s + (Number(r.kilos) || 0), 0);
      doc.rect(xOrigin, ry + 2, halfW, ROW_H + 2).fillAndStroke('#FFF3CD', '#E0B020');
      doc.fillColor('#000').font('Helvetica-Bold').fontSize(8.5);
      doc.text(`Total      ${totBag}      ${fmtQty(totKilo)}`, xOrigin, ry + 6, {
        width: halfW, align: 'center', lineBreak: false,
      });
    }
  }

  // Slice rows into chunks of ROWS_PER_HALF and emit one printed page per chunk.
  for (let i = 0; i < totalPages; i++) {
    if (i > 0) doc.addPage();
    const slice = rows.slice(i * ROWS_PER_HALF, (i + 1) * ROWS_PER_HALF);
    const isLast = (i === totalPages - 1);
    // Left half
    drawHalfHeader(m, i + 1);
    drawHalfRows(m, slice, isLast);
    // Right half (carbon copy)
    drawHalfHeader(m + halfW + gutter, i + 1);
    drawHalfRows(m + halfW + gutter, slice, isLast);

    // Vertical dashed cut-line down the middle so the page can be torn into
    // two physical copies.
    const cutX = m + halfW + gutter / 2;
    doc.save();
    doc.dash(3, { space: 3 });
    doc.moveTo(cutX, m).lineTo(cutX, pageH - m)
       .lineWidth(0.5).strokeColor('#888').stroke();
    doc.undash();
    doc.restore();
    doc.font('Helvetica').fontSize(10).fillColor('#888')
       .text('✂', cutX - 4, m - 12, { lineBreak: false });
  }

  return new Promise(resolve => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

// ════════════════════════════════════════════════════════════
// REPORT 2 — TRUCK LIST (buyer-code summary)
// ════════════════════════════════════════════════════════════
function getTruckListRows(db, auctionId) {
  // Group by short CODE — that is what's printed on the truck/lorry tag
  // and what dad's truck-list FoxPro report keys on.
  return db.all(`
    SELECT
      COALESCE(NULLIF(l.code,''), 'UNKNOWN') AS code,
      COUNT(*)        AS lot_count,
      SUM(l.bags)     AS bag,
      SUM(l.qty)      AS qty,
      SUM(l.amount)   AS amount
    FROM lots l
    WHERE l.auction_id = ?
      AND l.amount > 0
    GROUP BY COALESCE(NULLIF(l.code,''), 'UNKNOWN')
    ORDER BY CASE WHEN COALESCE(NULLIF(l.code,''), 'UNKNOWN')='UNKNOWN' THEN 1 ELSE 0 END,
             code
  `, [auctionId]);
}

async function truckListXlsx(db, auctionId) {
  const auction = getAuctionHeader(db, auctionId);
  const rows = getTruckListRows(db, auctionId);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('TruckList');
  ws.columns = [
    { width: 8 },  // SL.NO
    { width: 8 },  // LOT
    { width: 8 },  // BAG
    { width: 12 }, // CODE
    { width: 14 }, // QTY
    { width: 18 }, // AMOUNT
  ];

  // Three-column brand band.
  const startRow = writeXlsxCompanyHeader(wb, ws, getCompanyHeader(db), {
    colCount: 6,
    title: 'TRUCK LIST',
    metaLines: [
      `${modeHdrLabel(auction)} No: ${auction.ano}`,
      `Date: ${fmtDateDMY(auction.date)}`,
    ],
  });

  const head = ws.getRow(startRow);
  ['SL.NO', 'LOT', 'BAG', 'CODE', 'QTY', 'AMOUNT'].forEach((h, i) => {
    head.getCell(i + 1).value = h;
  });
  head.font = { bold: true };
  head.eachCell((c, ci) => {
    if (ci > 6) return;
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
    c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    c.alignment = { horizontal: 'center' };
  });

  let tLot = 0, tBag = 0, tQty = 0, tAmt = 0;
  rows.forEach((r, idx) => {
    const row = ws.addRow([
      idx + 1,           // SL.NO (sequential)
      r.lot_count,       // LOT (count of lots in this truck/code)
      r.bag,
      r.code,
      Number(r.qty),
      Number(r.amount),
    ]);
    row.getCell(1).alignment = { horizontal: 'center' };
    row.getCell(2).alignment = { horizontal: 'right' };
    row.getCell(3).alignment = { horizontal: 'right' };
    row.getCell(4).alignment = { horizontal: 'center' };
    row.getCell(5).alignment = { horizontal: 'right' };
    row.getCell(5).numFmt = '#,##0.000';
    row.getCell(6).alignment = { horizontal: 'right' };
    row.getCell(6).numFmt = '#,##,##0.00';

    tLot += Number(r.lot_count) || 0;
    tBag += Number(r.bag) || 0;
    tQty += Number(r.qty) || 0;
    tAmt += Number(r.amount) || 0;
  });

  const tot = ws.addRow(['', tLot, tBag, 'TOTAL', tQty, tAmt]);
  tot.font = { bold: true };
  tot.getCell(5).numFmt = '#,##0.000';
  tot.getCell(6).numFmt = '#,##,##0.00';
  tot.getCell(4).alignment = { horizontal: 'center' };
  tot.eachCell(c => {
    c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
  });

  return wb.xlsx.writeBuffer();
}

async function truckListPdf(db, auctionId) {
  const auction = getAuctionHeader(db, auctionId);
  const rows = getTruckListRows(db, auctionId);

  const doc = new PDFDocument({ size: 'A4', layout: 'portrait', margin: 36 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));

  const m = 36;
  const usableW = doc.page.width - m * 2;
  const colW = [
    Math.floor(usableW * 0.08),  // SL.NO
    Math.floor(usableW * 0.10),  // LOT
    Math.floor(usableW * 0.10),  // BAG
    Math.floor(usableW * 0.18),  // CODE
    Math.floor(usableW * 0.20),  // QTY
    0,                           // AMOUNT
  ];
  colW[5] = usableW - colW[0] - colW[1] - colW[2] - colW[3] - colW[4];
  const colX = [m];
  for (let i = 0; i < colW.length - 1; i++) colX.push(colX[i] + colW[i]);

  const ROW_H = 18;
  const HEAD_H = 20;
  let y;
  // Track each printed page's table top so vertical column separators run from
  // the very top of the column-header strip down through every body row.
  let pageStartY;
  const pageStarts = [];

  // Resolve company branding once for use across pages
  const companyHeader = getCompanyHeader(db);

  function drawTopHeader() {
    const afterY = drawCompanyHeader(doc, companyHeader, {
      x: m, y: m, width: usableW,
      title: 'TRUCK LIST',
      metaLines: [
        `${modeHdrLabel(auction)} No: ${auction.ano}`,
        `Date: ${fmtDateDMY(auction.date)}`,
      ],
    });
    y = afterY;
  }
  function drawColHeader() {
    pageStartY = y;
    pageStarts.push({ page: doc.bufferedPageRange().count - 1, top: y });
    doc.rect(m, y, usableW, HEAD_H).fillAndStroke('#E8E4DD', '#444');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(10);
    const heads = ['SL.NO', 'LOT', 'BAG', 'CODE', 'QTY', 'AMOUNT'];
    const aligns = ['center', 'center', 'center', 'center', 'right', 'right'];
    heads.forEach((h, i) => {
      doc.text(h, colX[i] + 4, y + 5, { width: colW[i] - 8, align: aligns[i], lineBreak: false });
    });
    y += HEAD_H;
  }
  function drawRow(r, idx) {
    // Inset the fill's top edge by the rule width so it doesn't paint over
    // the previous row's separator (drawn at this row's top y).
    if (idx % 2 === 1) doc.rect(m, y + 0.5, usableW, ROW_H - 0.5).fill('#F7F5F2');
    doc.fillColor('#000').font('Helvetica').fontSize(10);
    doc.text(String(idx + 1),     colX[0] + 4, y + 5, { width: colW[0] - 8, align: 'center', lineBreak: false });
    doc.text(String(r.lot_count), colX[1] + 4, y + 5, { width: colW[1] - 8, align: 'center', lineBreak: false });
    doc.text(String(r.bag),       colX[2] + 4, y + 5, { width: colW[2] - 8, align: 'center', lineBreak: false });
    doc.font('Helvetica-Bold');
    doc.text(String(r.code || ''),colX[3] + 4, y + 5, { width: colW[3] - 8, align: 'center', lineBreak: false });
    doc.font('Helvetica');
    doc.text(fmtQty(r.qty),       colX[4] + 4, y + 5, { width: colW[4] - 8, align: 'right',  lineBreak: false });
    doc.text(fmtMoney(r.amount),  colX[5] + 4, y + 5, { width: colW[5] - 8, align: 'right',  lineBreak: false });
    doc.moveTo(m, y + ROW_H).lineTo(m + usableW, y + ROW_H).lineWidth(0.25).strokeColor('#BBB').stroke();
    y += ROW_H;
  }

  // Draw vertical column separators on the *current* page from `top` down to `bottom`.
  function drawVerticalsOnCurrentPage(top, bottom) {
    for (let ci = 0; ci < colW.length - 1; ci++) {
      const vx = colX[ci] + colW[ci];
      doc.moveTo(vx, top).lineTo(vx, bottom)
         .lineWidth(0.25).strokeColor('#888').stroke();
    }
  }

  drawTopHeader();
  drawColHeader();

  let tLot = 0, tBag = 0, tQty = 0, tAmt = 0;
  rows.forEach((r, i) => {
    if (y + ROW_H > doc.page.height - m - 30) {
      // Close out vertical lines on the page we're leaving — table ends at y here.
      const cur = pageStarts[pageStarts.length - 1];
      drawVerticalsOnCurrentPage(cur.top, y);
      doc.rect(m, cur.top, usableW, y - cur.top).lineWidth(0.5).strokeColor('#444').stroke();
      doc.addPage();
      drawTopHeader();
      drawColHeader();
    }
    drawRow(r, i);
    tLot += Number(r.lot_count) || 0;
    tBag += Number(r.bag) || 0;
    tQty += Number(r.qty) || 0;
    tAmt += Number(r.amount) || 0;
  });

  // Total row
  if (y + ROW_H + 4 > doc.page.height - m) {
    const cur = pageStarts[pageStarts.length - 1];
    drawVerticalsOnCurrentPage(cur.top, y);
    doc.rect(m, cur.top, usableW, y - cur.top).lineWidth(0.5).strokeColor('#444').stroke();
    doc.addPage(); drawTopHeader(); drawColHeader();
  }
  doc.rect(m, y, usableW, ROW_H + 4).fillAndStroke('#FFF3CD', '#E0B020');
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(10.5);
  // SL.NO column blank in total row
  doc.text(String(tLot),    colX[1] + 4, y + 6, { width: colW[1] - 8, align: 'center', lineBreak: false });
  doc.text(String(tBag),    colX[2] + 4, y + 6, { width: colW[2] - 8, align: 'center', lineBreak: false });
  doc.text('TOTAL',         colX[3] + 4, y + 6, { width: colW[3] - 8, align: 'center', lineBreak: false });
  doc.text(fmtQty(tQty),    colX[4] + 4, y + 6, { width: colW[4] - 8, align: 'right',  lineBreak: false });
  doc.text(fmtMoney(tAmt),  colX[5] + 4, y + 6, { width: colW[5] - 8, align: 'right',  lineBreak: false });
  y += ROW_H + 4;

  // Vertical column separators on the final page + outer border
  const last = pageStarts[pageStarts.length - 1];
  drawVerticalsOnCurrentPage(last.top, y);
  doc.rect(m, last.top, usableW, y - last.top).lineWidth(0.5).strokeColor('#444').stroke();

  return new Promise(resolve => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

// ════════════════════════════════════════════════════════════
// REPORT 3 — BUYER LOT LORRY
// ════════════════════════════════════════════════════════════
//
// Returns rows pre-grouped: { interState: [...buyerGroup], intraState: [...] }
// Each buyer group:
//   { code, buyer1, sbl, gstin, state, lots: [{lot,bag,qty,rate,amount}],
//     totalLotCount, totalBag, totalQty, totalAmount }
function getBuyerLotLorryData(db, auctionId) {
  const auction = getAuctionHeader(db, auctionId);
  const rows    = getLotsWithBuyer(db, auctionId);
  const auctionState = String(auction.state || '').trim().toUpperCase();

  // Group by buyer code (short) — falls back to buyer (full) if code is empty
  const groups = new Map();
  for (const r of rows) {
    const groupKey = (r.code || r.buyer || 'UNKNOWN').toUpperCase();
    if (!groups.has(groupKey)) {
      groups.set(groupKey, {
        code:    r.code || '',
        buyer:   r.buyer  || '',
        buyer1:  r.buyer1 || '',
        sbl:     r.sbl    || '',
        gstin:   r.gstin  || '',
        state:   r.state  || '',
        lots: [],
      });
    }
    groups.get(groupKey).lots.push({
      lot:    r.lot,
      bag:    Number(r.bag) || 0,
      qty:    Number(r.qty) || 0,
      rate:   Number(r.price) || 0,
      amount: Number(r.amount) || 0,
    });
  }

  // Compute subtotals + classify each group as inter/intra
  const interState = [];
  const intraState = [];
  for (const g of groups.values()) {
    g.totalLotCount = g.lots.length;
    g.totalBag      = g.lots.reduce((s, x) => s + x.bag, 0);
    g.totalQty      = g.lots.reduce((s, x) => s + x.qty, 0);
    g.totalAmount   = g.lots.reduce((s, x) => s + x.amount, 0);

    const buyerState = String(g.state || '').trim().toUpperCase();
    // If buyer state is unknown, default to intra (matches FoxPro fallback).
    const isIntra = !buyerState || buyerState === auctionState;
    if (isIntra) intraState.push(g);
    else         interState.push(g);
  }

  // Sort by full buyer name (buyer1) then code, alphabetically — matches the
  // sample BUYER_LOT_LORRY.pdf which is alpha by full buyer name.
  const sortFn = (a, b) => {
    const an = (a.buyer1 || a.sbl || a.code || '').toUpperCase();
    const bn = (b.buyer1 || b.sbl || b.code || '').toUpperCase();
    if (an !== bn) return an < bn ? -1 : 1;
    // Same name (e.g. SIVA TEACHER appears for both SIVA and SIVA1) — sort by code
    return (a.code || '').localeCompare(b.code || '');
  };
  interState.sort(sortFn);
  intraState.sort(sortFn);

  return { auction, interState, intraState };
}

async function buyerLotLorryXlsx(db, auctionId) {
  const { auction, interState, intraState } = getBuyerLotLorryData(db, auctionId);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('BuyerLotLorry');

  ws.columns = [
    { width: 8 },   // SL.NO
    { width: 10 },  // LOT
    { width: 8 },   // BAG
    { width: 14 },  // QTY
    { width: 12 },  // RATE
    { width: 18 },  // AMOUNT
  ];

  // Three-column brand band (6 columns).
  writeXlsxCompanyHeader(wb, ws, getCompanyHeader(db), {
    colCount: 6,
    title: 'BUYER LOT LORRY',
    metaLines: [
      `${modeHdrLabel(auction)} No: ${auction.ano}`,
      `Date: ${fmtDateDMY(auction.date)}`,
      auction.state || '',
    ].filter(Boolean),
  });

  // Grand counters across the whole report
  let gLot = 0, gBag = 0, gQty = 0, gAmt = 0;
  // Counters for the intra-only mid-subtotal (matches the sample's intra-state subtotal)
  let intraLot = 0, intraBag = 0, intraQty = 0, intraAmt = 0;
  // Buyer counter — increments across both inter and intra so each buyer gets
  // a unique sequence number prefix matching the PDF ("1. FLORA SPICES", etc.)
  let buyerSeq = 0;

  function emitSection(title, groups) {
    if (!groups.length) return;
    const sec = ws.addRow([title]);
    ws.mergeCells(`A${sec.number}:F${sec.number}`);
    sec.font = { bold: true, size: 11 };
    sec.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD4EDDA' } };
    sec.alignment = { horizontal: 'center' };

    groups.forEach(g => {
      buyerSeq += 1;
      // Buyer header row 1: numbered name + code (e.g. "1. FLORA SPICES   [FSS]")
      const h1 = ws.addRow([`${buyerSeq}. ${g.buyer1 || g.sbl || g.buyer}    [${g.code}]`]);
      ws.mergeCells(`A${h1.number}:F${h1.number}`);
      h1.font = { bold: true, size: 10 };
      h1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };

      // Buyer header row 2: proprietor name + GSTIN
      if (g.sbl || g.gstin) {
        const h2 = ws.addRow([`${g.sbl || ''}    ${g.gstin || ''}`]);
        ws.mergeCells(`A${h2.number}:F${h2.number}`);
        h2.font = { italic: true, size: 9, color: { argb: 'FF555555' } };
      }

      // Column header — 6 columns matching PDF
      const ch = ws.addRow(['SL.NO', 'LOT', 'BAG', 'QTY', 'RATE', 'AMOUNT']);
      ch.font = { bold: true, size: 9 };
      ch.eachCell(c => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
        c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
        c.alignment = { horizontal: 'center' };
      });

      // Lot rows (sorted by lot number)
      g.lots.sort((a, b) => {
        const na = parseInt(a.lot, 10), nb = parseInt(b.lot, 10);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return String(a.lot).localeCompare(String(b.lot));
      });
      // Per-buyer lot serial resets to 1 at the start of each buyer block —
      // matches the PDF's SL.NO column behaviour.
      g.lots.forEach((lt, idx) => {
        const r = ws.addRow([idx + 1, lt.lot, lt.bag, lt.qty, lt.rate, lt.amount]);
        r.getCell(1).alignment = { horizontal: 'center' };
        r.getCell(2).alignment = { horizontal: 'center' };
        r.getCell(3).alignment = { horizontal: 'right' };
        r.getCell(4).alignment = { horizontal: 'right' };
        r.getCell(4).numFmt = '#,##0.000';
        r.getCell(5).alignment = { horizontal: 'right' };
        r.getCell(5).numFmt = '#,##0.00';
        r.getCell(6).alignment = { horizontal: 'right' };
        r.getCell(6).numFmt = '#,##,##0.00';
      });

      // Buyer subtotal — blank SL.NO, sum BAG/QTY/AMOUNT, blank RATE
      const sub = ws.addRow(['', g.totalLotCount, g.totalBag, g.totalQty, '', g.totalAmount]);
      sub.font = { bold: true };
      sub.getCell(4).numFmt = '#,##0.000';
      sub.getCell(6).numFmt = '#,##,##0.00';
      sub.eachCell(c => {
        c.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F5F2' } };
      });
      ws.addRow([]); // spacer

      gLot += g.totalLotCount;
      gBag += g.totalBag;
      gQty += g.totalQty;
      gAmt += g.totalAmount;
    });
  }

  emitSection('INTER-STATE SALES', interState);
  // Intra-state group — capture the intra-only subtotal between the two
  // sections (matches the FoxPro layout which prints it before the grand total)
  const beforeIntra = { gLot, gBag, gQty, gAmt };
  emitSection('INTRA-STATE SALES', intraState);
  intraLot = gLot - beforeIntra.gLot;
  intraBag = gBag - beforeIntra.gBag;
  intraQty = gQty - beforeIntra.gQty;
  intraAmt = gAmt - beforeIntra.gAmt;

  // Intra-state subtotal (matches FoxPro: only printed when both sections exist)
  if (interState.length && intraState.length) {
    const intraTot = ws.addRow(['', intraLot, intraBag, intraQty, '', intraAmt]);
    intraTot.font = { bold: true };
    intraTot.getCell(4).numFmt = '#,##0.000';
    intraTot.getCell(6).numFmt = '#,##,##0.00';
    intraTot.eachCell((c, ci) => {
      c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F4EA' } };
      if (ci === 5) c.value = 'INTRA-STATE TOTAL';
    });
  }

  // Grand total
  const grand = ws.addRow(['', gLot, gBag, gQty, '', gAmt]);
  grand.font = { bold: true, size: 11 };
  grand.getCell(4).numFmt = '#,##0.000';
  grand.getCell(6).numFmt = '#,##,##0.00';
  grand.eachCell((c, ci) => {
    c.border = { top: { style: 'double' }, bottom: { style: 'double' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
    if (ci === 5) c.value = 'GRAND TOTAL';
  });

  return wb.xlsx.writeBuffer();
}

async function buyerLotLorryPdf(db, auctionId) {
  const { auction, interState, intraState } = getBuyerLotLorryData(db, auctionId);

  const doc = new PDFDocument({ size: 'A4', layout: 'portrait', margin: 30 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));

  const m = 30;
  const pageW = doc.page.width;
  const pageH = doc.page.height;
  const usableW = pageW - m * 2;

  // Cols: SL.NO | LOT | BAG | QTY | RATE | AMOUNT
  // Serial restarts within each buyer's block.
  const colW = [
    Math.floor(usableW * 0.08),  // SL.NO
    Math.floor(usableW * 0.10),  // LOT
    Math.floor(usableW * 0.10),  // BAG
    Math.floor(usableW * 0.18),  // QTY
    Math.floor(usableW * 0.18),  // RATE
    0,                           // AMOUNT
  ];
  colW[5] = usableW - colW[0] - colW[1] - colW[2] - colW[3] - colW[4];
  const colX = [m];
  for (let i = 0; i < colW.length - 1; i++) colX.push(colX[i] + colW[i]);

  const ROW_H   = 14;
  const HEAD_H  = 16;
  const BUYER_H = 30;  // 2 lines of buyer info
  const SECT_H  = 20;
  let y;

  function ensureRoom(needed) {
    if (y + needed > pageH - m - 14) {
      // If we're mid-block (between drawColHeader and drawSubtotal), close out
      // the vertical lines and outer border on the page we're leaving so the
      // partial block on this page looks complete.
      if (blockTop !== null) {
        for (let ci = 0; ci < colW.length - 1; ci++) {
          const vx = colX[ci] + colW[ci];
          doc.moveTo(vx, blockTop).lineTo(vx, y)
             .lineWidth(0.3).strokeColor('#777').stroke();
        }
        doc.rect(m, blockTop, usableW, y - blockTop)
           .lineWidth(0.5).strokeColor('#444').stroke();
      }
      doc.addPage();
      drawTopHeader();
      // Continuation: the next thing drawn is more lot rows, so the block
      // effectively restarts at the new y (just below the page top header).
      if (blockTop !== null) blockTop = y;
    }
  }

  // Resolve company branding once for use across pages
  const companyHeader = getCompanyHeader(db);

  function drawTopHeader() {
    const afterY = drawCompanyHeader(doc, companyHeader, {
      x: m, y: m, width: usableW,
      title: 'BUYER LOT LORRY',
      metaLines: [
        `${modeHdrLabel(auction)} No: ${auction.ano}`,
        `Date: ${fmtDateDMY(auction.date)}`,
      ],
    });
    y = afterY;
  }

  function drawSection(label) {
    ensureRoom(SECT_H + HEAD_H + BUYER_H + ROW_H + 8);
    doc.rect(m, y, usableW, SECT_H).fillAndStroke('#D4EDDA', '#5A8F62');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(11)
       .text(label, m, y + 4, { width: usableW, align: 'center', lineBreak: false });
    y += SECT_H + 4;
  }

  function drawBuyerHeader(g, sn) {
    // Two lines of buyer info: name+code, then firm-name+GSTIN.
    // Long names wrap to multiple lines instead of being truncated. The
    // buyer serial number (one per unique buyer) is prefixed onto the
    // firm name like "1. EMPEROR SPICES PRIVATE LIMITED".
    const nameW = usableW * 0.7;
    const codeW = usableW * 0.3;
    const sblW  = usableW * 0.6;
    const gstW  = usableW * 0.4;

    const firmRaw = g.buyer1 || g.sbl || g.buyer || '';
    const firmTxt = sn != null ? `${sn}. ${firmRaw}` : firmRaw;

    doc.fillColor('#7A4400').font('Helvetica-Bold').fontSize(10);
    const nameLines = wrapText(doc, firmTxt, nameW - 4);
    nameLines.forEach((ln, i) => {
      doc.text(ln, m, y + i * 12, { width: nameW, align: 'left', lineBreak: false });
    });

    doc.font('Helvetica-Bold').fontSize(11).fillColor('#B85C00');
    doc.text(`[${g.code}]`, m + nameW, y, { width: codeW, align: 'right', lineBreak: false });
    y += Math.max(13, nameLines.length * 12 + 1);

    if (g.sbl || g.gstin) {
      doc.font('Helvetica').fontSize(8).fillColor('#555');
      const sblLines = wrapText(doc, g.sbl || '', sblW - 4);
      sblLines.forEach((ln, i) => {
        doc.text(ln, m, y + i * 10, { width: sblW, align: 'left', lineBreak: false });
      });

      doc.font('Helvetica-Oblique').fontSize(8).fillColor('#555');
      doc.text(g.gstin || '', m + sblW, y, { width: gstW, align: 'right', lineBreak: false });
      y += Math.max(11, sblLines.length * 10 + 1);
    }
    y += 2;
  }

  // Track the top-Y of the current per-buyer table block so we can draw
  // vertical column separators after the buyer's subtotal closes the block.
  let blockTop = null;

  function drawColHeader() {
    blockTop = y;  // remember where this block's table begins
    doc.rect(m, y, usableW, HEAD_H).fillAndStroke('#E8E4DD', '#444');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(8.5);
    const heads = ['SL.NO', 'LOT', 'BAG', 'QTY', 'RATE', 'AMOUNT'];
    const aligns = ['center', 'center', 'center', 'right', 'right', 'right'];
    heads.forEach((h, i) => {
      doc.text(h, colX[i] + 4, y + 4, { width: colW[i] - 8, align: aligns[i], lineBreak: false });
    });
    y += HEAD_H;
  }

  function drawLotRow(lt, idx) {
    // Inset the fill's top edge by the rule width so it doesn't paint over
    // the previous row's separator (drawn at this row's top y).
    if (idx % 2 === 1) doc.rect(m, y + 0.5, usableW, ROW_H - 0.5).fill('#F7F5F2');
    doc.fillColor('#000').font('Helvetica').fontSize(8.5);
    const cells = [
      String(idx + 1),         // SL.NO — restarts at 1 within each buyer
                                // block (idx is the lot index inside the
                                // current buyer's lots loop).
      String(lt.lot),
      String(lt.bag),
      fmtQty(lt.qty),
      fmtPrice(lt.rate),
      fmtMoney(lt.amount),
    ];
    const aligns = ['center', 'center', 'center', 'right', 'right', 'right'];
    cells.forEach((v, ci) => {
      doc.text(v, colX[ci] + 4, y + 3, { width: colW[ci] - 8, align: aligns[ci], lineBreak: false });
    });
    doc.moveTo(m, y + ROW_H).lineTo(m + usableW, y + ROW_H).lineWidth(0.25).strokeColor('#CCC').stroke();
    y += ROW_H;
  }

  function drawSubtotal(g) {
    doc.rect(m, y, usableW, ROW_H + 2).fillAndStroke('#FFF3CD', '#E0B020');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(8.5);
    const cells = [
      '',                          // SL.NO blank — the buyer serial number
                                    // appears in the buyer header above.
      String(g.totalLotCount),
      String(g.totalBag),
      fmtQty(g.totalQty),
      '',
      fmtMoney(g.totalAmount),
    ];
    const aligns = ['center', 'center', 'center', 'right', 'right', 'right'];
    cells.forEach((v, ci) => {
      doc.text(v, colX[ci] + 4, y + 5, { width: colW[ci] - 8, align: aligns[ci], lineBreak: false });
    });
    const blockBottom = y + ROW_H + 2;
    // Vertical column separators from the top of this buyer's column-header
    // strip down through the subtotal — only inside this per-buyer block.
    if (blockTop !== null) {
      for (let ci = 0; ci < colW.length - 1; ci++) {
        const vx = colX[ci] + colW[ci];
        doc.moveTo(vx, blockTop).lineTo(vx, blockBottom)
           .lineWidth(0.3).strokeColor('#777').stroke();
      }
      // Outer border around the buyer block (header strip + lot rows + subtotal)
      doc.rect(m, blockTop, usableW, blockBottom - blockTop)
         .lineWidth(0.5).strokeColor('#444').stroke();
    }
    blockTop = null;
    y += ROW_H + 6;
  }

  drawTopHeader();

  let gLot = 0, gBag = 0, gQty = 0, gAmt = 0;
  // Counter for the per-unique-buyer serial number. Increments once per
  // buyer block (across both INTER-STATE and INTRA-STATE sections so the
  // serials run continuously through the report).
  let buyerSerial = 0;

  function emitSection(label, groups) {
    if (!groups.length) return;
    drawSection(label);
    groups.forEach(g => {
      // Reserve enough room for header + at least one row + subtotal
      ensureRoom(BUYER_H + HEAD_H + ROW_H * 2 + 8);
      buyerSerial += 1;
      drawBuyerHeader(g, buyerSerial);
      drawColHeader();
      g.lots.sort((a, b) => {
        const na = parseInt(a.lot, 10), nb = parseInt(b.lot, 10);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return String(a.lot).localeCompare(String(b.lot));
      });
      g.lots.forEach((lt, i) => {
        ensureRoom(ROW_H + ROW_H + 4);
        drawLotRow(lt, i);
      });
      ensureRoom(ROW_H + 6);
      drawSubtotal(g);
      gLot += g.totalLotCount; gBag += g.totalBag; gQty += g.totalQty; gAmt += g.totalAmount;
    });
  }

  emitSection('INTER-STATE SALES', interState);
  const beforeIntra = { gLot, gBag, gQty, gAmt };
  emitSection('INTRA-STATE SALES', intraState);

  // Intra-state subtotal (only when both sections present)
  if (interState.length && intraState.length) {
    const intraLot = gLot - beforeIntra.gLot;
    const intraBag = gBag - beforeIntra.gBag;
    const intraQty = gQty - beforeIntra.gQty;
    const intraAmt = gAmt - beforeIntra.gAmt;

    ensureRoom(ROW_H + 8);
    const intraTop = y;
    doc.rect(m, y, usableW, ROW_H + 4).fillAndStroke('#E6F4EA', '#5A8F62');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
    const cells = [String(intraLot), String(intraBag), fmtQty(intraQty), 'INTRA-STATE TOTAL', fmtMoney(intraAmt)];
    cells.forEach((v, ci) => {
      const align = (ci === 0 || ci === 1) ? 'center' : (ci === 3 ? 'center' : 'right');
      doc.text(v, colX[ci] + 4, y + 6, { width: colW[ci] - 8, align, lineBreak: false });
    });
    // Vertical separators inside the intra-state subtotal strip
    for (let ci = 0; ci < colW.length - 1; ci++) {
      const vx = colX[ci] + colW[ci];
      doc.moveTo(vx, intraTop).lineTo(vx, intraTop + ROW_H + 4)
         .lineWidth(0.3).strokeColor('#5A8F62').stroke();
    }
    y += ROW_H + 8;
  }

  // Grand total
  ensureRoom(ROW_H + 10);
  const grandTop = y;
  doc.rect(m, y, usableW, ROW_H + 6).fillAndStroke('#FFF3CD', '#9A6700');
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(10);
  // 6 cols: SL.NO | LOT | BAG | QTY | RATE | AMOUNT
  // Put count totals in LOT/BAG, "GRAND TOTAL" label in RATE column, amount in AMOUNT.
  const gcells = ['', String(gLot), String(gBag), fmtQty(gQty), 'GRAND TOTAL', fmtMoney(gAmt)];
  const galigns = ['center', 'center', 'center', 'right', 'center', 'right'];
  gcells.forEach((v, ci) => {
    doc.text(v, colX[ci] + 4, y + 6, { width: colW[ci] - 8, align: galigns[ci], lineBreak: false });
  });
  // Vertical separators inside the grand total strip
  for (let ci = 0; ci < colW.length - 1; ci++) {
    const vx = colX[ci] + colW[ci];
    doc.moveTo(vx, grandTop).lineTo(vx, grandTop + ROW_H + 6)
       .lineWidth(0.3).strokeColor('#9A6700').stroke();
  }

  return new Promise(resolve => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

// ════════════════════════════════════════════════════════════
// Dispatcher — used by the /api/lorry-reports/:type/:auctionId route
// ════════════════════════════════════════════════════════════
const REPORTS = {
  lot_slip_code:     { name: 'LotSlipCode',     xlsx: lotSlipCodeXlsx,    pdf: lotSlipCodePdf    },
  truck_list:        { name: 'TruckList',       xlsx: truckListXlsx,      pdf: truckListPdf      },
  buyer_lot_lorry:   { name: 'BuyerLotLorry',   xlsx: buyerLotLorryXlsx,  pdf: buyerLotLorryPdf  },
};

module.exports = { REPORTS };
