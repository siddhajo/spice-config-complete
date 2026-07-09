/**
 * auction-reports.js — Specialized auction reports that don't fit the generic
 * column-based table renderer in exports-pdf.js. Each one mirrors a specific
 * FoxPro report layout the customer is replacing.
 *
 *   1. Lot Slip (PDF)         — pre-trade carbon-copy slip with LOT|BAG|QTY|PRICE
 *   2. Collection (XLSX+PDF)  — invoice register: INVO|TRADE NAME|NAME|QTY|VALUE
 *   3. Trade Report (XLSX+PDF)— BUYERS LIST FOR VERIFICATION grouped by state
 *
 * The XLSX path goes through a custom worksheet builder (not createExcelBuffer)
 * so we can emit headers/subtotals/group-rows that the generic builder can't.
 */

const ExcelJS = require('exceljs');
const PDFDocument = require('pdfkit');
const {
  fmtMoney, fmtQty, fmtPrice,
  getCompanyHeader, drawCompanyHeader,
} = require('./report-formatters');
const { getSettingsFlat } = require('./company-config');

// ── Number formatting ────────────────────────────────────────
// fmtMoney / fmtQty / fmtPrice come from report-formatters.js (Indian comma
// grouping; 2 decimals for rupees, 3 for kilos).
function fmtDateDMY(iso) {
  if (!iso) return '';
  const s = String(iso);
  if (s.includes('-') && s.length >= 10) return s.slice(0, 10).split('-').reverse().join('/');
  return s;
}

// ── Mode-aware label helpers ────────────────────────────────
// PDF report headers say "e-TRADE No: 3" or "e-AUCTION No: 3" based on
// the auction's mode. Footer totals adapt the same way:
//   e-Trade   → "NOT eTRADED"
//   e-Auction → "NOT AUCTIONED"
// Pass the auction row (which now carries `mode`) and the helpers do
// the right thing. NULL/empty mode rows fall back to e-AUCTION so
// legacy data prints with the historical wording.
function modeHeaderLabel(auction) {
  return (auction && auction.mode === 'e-Trade') ? 'e-TRADE' : 'e-AUCTION';
}
function modeNotXyzLabel(auction) {
  return (auction && auction.mode === 'e-Trade') ? 'NOT eTRADED' : 'NOT AUCTIONED';
}

// Manually truncate `text` so doc.widthOfString(out) <= maxWidth, appending an
// ellipsis when truncated. PDFKit 0.15's `lineBreak: false` + `ellipsis: true`
// is unreliable for long single tokens — we ellipsize ourselves so multi-word
// names like "EMPEROR SPICES PRIVATE LIMITED" never wrap into the next row.
// Caller must already have set the desired font/size on `doc` before invoking.
function fitText(doc, text, maxWidth) {
  const s = String(text == null ? '' : text);
  if (!s) return '';
  if (doc.widthOfString(s) <= maxWidth) return s;
  const ell = '…';
  const ellW = doc.widthOfString(ell);
  if (ellW >= maxWidth) return '';
  // Binary search for longest prefix that fits with the ellipsis appended.
  let lo = 0, hi = s.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (doc.widthOfString(s.slice(0, mid)) + ellW <= maxWidth) lo = mid;
    else hi = mid - 1;
  }
  return s.slice(0, lo).trimEnd() + ell;
}

// Wrap `text` into one or more lines, each fitting within `maxWidth`. Tries
// to break on word boundaries; if a single token is wider than maxWidth, it
// breaks at character level so nothing ever overflows. Caller must have
// already set the desired font/size on `doc` before invoking. Returns an
// array of strings (one per line), at least one element.
function wrapText(doc, text, maxWidth) {
  const s = String(text == null ? '' : text).trim();
  if (!s) return [''];
  if (doc.widthOfString(s) <= maxWidth) return [s];

  const words = s.split(/\s+/);
  const lines = [];
  let cur = '';
  for (const w of words) {
    const probe = cur ? cur + ' ' + w : w;
    if (doc.widthOfString(probe) <= maxWidth) {
      cur = probe;
      continue;
    }
    // probe doesn't fit. Push current line (if any) and start fresh with w.
    if (cur) { lines.push(cur); cur = ''; }
    // If the word itself is too wide, break it at character level.
    if (doc.widthOfString(w) > maxWidth) {
      let chunk = '';
      for (const ch of w) {
        if (doc.widthOfString(chunk + ch) <= maxWidth) {
          chunk += ch;
        } else {
          if (chunk) lines.push(chunk);
          chunk = ch;
        }
      }
      cur = chunk;
    } else {
      cur = w;
    }
  }
  if (cur) lines.push(cur);
  return lines.length ? lines : [''];
}

function getAuctionHeader(db, auctionId) {
  // Include `mode` so PDF/XLSX headers can adapt their "e-TRADE No:" /
  // "NOT eTRADED" labels per-trade. Mode is empty for pre-migration
  // auctions; the mode helpers fall back to e-AUCTION wording in that
  // case so legacy data prints with the historical labels.
  const a = db.get('SELECT id, ano, date, crop_type, state, mode FROM auctions WHERE id = ?', [auctionId]);
  if (!a) throw new Error('Auction not found');
  return a;
}

// ════════════════════════════════════════════════════════════
// REPORT 1 — LOT SLIP (PDF carbon-copy, pre-trade)
// ════════════════════════════════════════════════════════════
//
// Mirrors LOT_SLIP.pdf. Two identical halves per page (carbon copy), 4 cols
// per half: LOT | BAG | QTY | PRICE — PRICE is intentionally blank because
// this slip is printed BEFORE the auction so the auctioneer can fill in
// hammer prices by hand. Total bags + qty at the end.

function getLotSlipPreRows(db, auctionId, state) {
  const where = state ? 'AND state = ?' : '';
  const params = state ? [auctionId, state] : [auctionId];
  return db.all(
    `SELECT lot_no AS lot, bags AS bag, qty FROM lots
     WHERE auction_id = ? ${where}
     ORDER BY CAST(lot_no AS INTEGER), lot_no`, params
  );
}

async function lotSlipPdf(db, auctionId, _cfg, extra) {
  const auction = getAuctionHeader(db, auctionId);
  const rows = getLotSlipPreRows(db, auctionId, extra && extra.state);

  const doc = new PDFDocument({ size: 'A4', layout: 'portrait', margin: 18 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));

  const m = 18;
  // Wide gutter between the two halves so the page can be torn or cut down
  // the middle to give two physical copies (one for the auction office, one
  // for the seller). A dashed cut-line is drawn in the gutter as a guide.
  const gutter = 40;
  const pageW = doc.page.width;
  const pageH = doc.page.height;
  const halfW = (pageW - m * 2 - gutter) / 2;

  // 4 cols: LOT | BAG | QTY | PRICE
  const colW = [
    Math.floor(halfW * 0.20),
    Math.floor(halfW * 0.18),
    Math.floor(halfW * 0.30),
    0,
  ];
  colW[3] = halfW - colW[0] - colW[1] - colW[2];

  const ROW_H = 16;
  const HEAD_H = 18;
  const TOP_H = 50;
  const BODY_TOP = m + TOP_H;
  const BODY_MAX_Y = pageH - m - 8;
  const ROWS_PER_HALF = Math.floor((BODY_MAX_Y - BODY_TOP - HEAD_H) / ROW_H);
  const totalPages = Math.max(1, Math.ceil(rows.length / ROWS_PER_HALF));

  // Resolve company branding once. The lot-slip carbon-copy halves are
  // narrow, so the brand band uses a small (22pt) logo and skips the
  // multi-line address — only the company name fits.
  const companyHeader = getCompanyHeader(db);

  function drawHalfHeader(xOrigin, page) {
    // Compact company brand band (no title/meta — too narrow for three cols).
    const afterY = drawCompanyHeader(doc, companyHeader, {
      x: xOrigin, y: m, width: halfW,
      logoH: 22, logoW: 22, showAddress: false,
    });
    // Page / e-TRADE / Date metadata stacks beneath the brand band.
    doc.font('Helvetica').fontSize(8)
       .text(`Page: ${page}`, xOrigin, afterY, { width: halfW, align: 'right' });
    doc.font('Helvetica-Bold').fontSize(9)
       .text(`${modeHeaderLabel(auction)} No:${auction.ano}`, xOrigin, afterY + 12, { width: halfW / 2, align: 'left' });
    doc.text(`Date:${fmtDateDMY(auction.date)}`, xOrigin + halfW / 2, afterY + 12, { width: halfW / 2, align: 'right' });

    const hy = BODY_TOP;
    doc.rect(xOrigin, hy, halfW, HEAD_H).fillAndStroke('#E8E4DD', '#444');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
    const heads = ['LOT', 'BAG', 'QTY', 'PRICE'];
    let cx = xOrigin;
    heads.forEach((h, i) => {
      doc.text(h, cx + 2, hy + 5, { width: colW[i] - 4, align: 'center', lineBreak: false });
      cx += colW[i];
    });
  }

  function drawHalfRows(xOrigin, sliceRows, isLastPage) {
    let ry = BODY_TOP + HEAD_H;
    const tblTop = BODY_TOP;

    // Pass 1 — zebra fills first. Drawing every fill before any rule keeps
    // the row separators visible: otherwise a shaded row's fill paints over
    // the bottom rule of the (unshaded) row above it, so every other line
    // disappears.
    let fy = ry;
    sliceRows.forEach((r, i) => {
      if (i % 2 === 1) doc.rect(xOrigin, fy, halfW, ROW_H).fill('#F7F5F2');
      fy += ROW_H;
    });

    // Pass 2 — cell text + horizontal separators, drawn on top of the fills.
    sliceRows.forEach((r, i) => {
      doc.fillColor('#000').font('Helvetica').fontSize(9);
      const cells = [
        String(r.lot),
        String(r.bag || 0),
        fmtQty(r.qty),
        '', // PRICE intentionally blank
      ];
      let cx = xOrigin;
      cells.forEach((v, ci) => {
        const align = ci === 0 ? 'center' : 'right';
        doc.text(v, cx + 2, ry + 4, {
          width: colW[ci] - 4,
          align,
          lineBreak: false, ellipsis: true,
        });
        cx += colW[ci];
      });
      // Horizontal row separator
      doc.moveTo(xOrigin, ry + ROW_H).lineTo(xOrigin + halfW, ry + ROW_H)
         .lineWidth(0.25).strokeColor('#999').stroke();
      ry += ROW_H;
    });

    // Vertical column separators
    let vx = xOrigin;
    for (let ci = 0; ci < colW.length - 1; ci++) {
      vx += colW[ci];
      doc.moveTo(vx, tblTop).lineTo(vx, ry)
         .lineWidth(0.25).strokeColor('#888').stroke();
    }
    // Outer table border
    doc.rect(xOrigin, tblTop, halfW, ry - tblTop)
       .lineWidth(0.5).strokeColor('#444').stroke();

    // Grand total on the last page
    if (isLastPage) {
      const totBag = rows.reduce((s, r) => s + (Number(r.bag) || 0), 0);
      const totQty = rows.reduce((s, r) => s + (Number(r.qty) || 0), 0);
      const ty = ry + 2;
      doc.rect(xOrigin, ty, halfW, ROW_H + 2).fillAndStroke('#FFF3CD', '#E0B020');
      doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
      // Distribute totals: 'Total' label in LOT col, then BAG, QTY (PRICE blank)
      const cells = ['Total', String(totBag), fmtQty(totQty), ''];
      let cx = xOrigin;
      cells.forEach((v, ci) => {
        const align = ci === 0 ? 'center' : 'right';
        doc.text(v, cx + 2, ty + 5, { width: colW[ci] - 4, align, lineBreak: false });
        cx += colW[ci];
      });
      // Verticals through the total row too
      let vx2 = xOrigin;
      for (let ci = 0; ci < colW.length - 1; ci++) {
        vx2 += colW[ci];
        doc.moveTo(vx2, ty).lineTo(vx2, ty + ROW_H + 2)
           .lineWidth(0.25).strokeColor('#E0B020').stroke();
      }
    }
  }

  for (let i = 0; i < totalPages; i++) {
    if (i > 0) doc.addPage();
    const slice = rows.slice(i * ROWS_PER_HALF, (i + 1) * ROWS_PER_HALF);
    const isLast = (i === totalPages - 1);
    drawHalfHeader(m, i + 1);
    drawHalfRows(m, slice, isLast);
    drawHalfHeader(m + halfW + gutter, i + 1);
    drawHalfRows(m + halfW + gutter, slice, isLast);

    // Vertical dashed cut-line down the middle of the gutter — a tearing
    // guide so the page can be split into two copies. A small "✂" hint at
    // the top makes the intent obvious at a glance.
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
// GENERALIZED CARBON-COPY SLIP (PDF)
// ════════════════════════════════════════════════════════════
//
// The same two-identical-halves-per-page tear-off layout as the Lot Slip,
// but driven by a column spec so other per-lot slips (Lot Buyer, Lot Name)
// can share it. Long text columns (fit:true) are ellipsized with fitText so
// a long buyer/seller name can never overlap the adjacent column — and
// every cell is clipped to its width (lineBreak:false) so nothing bleeds
// across a column divider.
//   opts.title      — heading shown under the brand band
//   opts.columns    — [{ header, key, weight, align, fit, kind, blank }]
//                     weight = relative share of the half-page width;
//                     kind 'qty'|'price' formats numbers; blank renders ''.
//   opts.totalKeys  — numeric keys summed into a Total row on the last page
//   opts.rows       — data rows
async function carbonSlipPdf(db, auctionId, _cfg, extra, opts) {
  const auction = getAuctionHeader(db, auctionId);
  const cols = opts.columns || [];
  const rows = opts.rows || [];
  const title = opts.title || '';
  const totalKeys = opts.totalKeys || [];

  const fmtCell = (kind, raw) => {
    if (kind === 'qty')   return fmtQty(raw);
    // Blank an unpriced cell rather than rendering "0.00" — the SQL already
    // returns '' for price=0 (pre-trade lots), but fmtPrice('') would coerce
    // it back to 0. Keep it empty so the PRICE column reads blank pre-trade.
    if (kind === 'price') return (raw === '' || raw == null || Number(raw) === 0) ? '' : fmtPrice(raw);
    return raw == null ? '' : String(raw);
  };

  const doc = new PDFDocument({ size: 'A4', layout: 'portrait', margin: 18 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));

  const m = 18;
  const gutter = 40;
  const pageW = doc.page.width;
  const pageH = doc.page.height;
  const halfW = (pageW - m * 2 - gutter) / 2;

  // Absolute column widths from weights; the last column absorbs rounding so
  // the columns always exactly fill halfW (no gap, no overhang).
  const wsum = cols.reduce((s, c) => s + (c.weight || 1), 0);
  let acc = 0;
  const colW = cols.map((c, i) => {
    if (i === cols.length - 1) return Math.max(8, halfW - acc);
    const w = Math.max(8, Math.floor(halfW * (c.weight || 1) / wsum));
    acc += w;
    return w;
  });

  const ROW_H = 16, HEAD_H = 18, TOP_H = 50;
  const BODY_TOP = m + TOP_H;
  const BODY_MAX_Y = pageH - m - 8;
  const ROWS_PER_HALF = Math.max(1, Math.floor((BODY_MAX_Y - BODY_TOP - HEAD_H) / ROW_H));
  const totalPages = Math.max(1, Math.ceil(rows.length / ROWS_PER_HALF));
  const companyHeader = getCompanyHeader(db);

  // Draw `text` on ONE line, shrinking the font (down to 5.5pt) until it
  // fits `w`. Used for headers and numeric cells so values like "2,463.00"
  // and headers like "CONTROL" never wrap to a second line (which would
  // overlap the row below) in the narrow half-page columns.
  function drawFit(text, x, y, w, align, baseSize, bold) {
    const s = String(text == null ? '' : text);
    doc.font(bold ? 'Helvetica-Bold' : 'Helvetica');
    let size = baseSize;
    doc.fontSize(size);
    while (size > 5.5 && doc.widthOfString(s) > w) { size -= 0.5; doc.fontSize(size); }
    doc.text(s, x, y, { width: w, align, lineBreak: false });
  }

  function drawHalfHeader(xOrigin, page) {
    const afterY = drawCompanyHeader(doc, companyHeader, {
      x: xOrigin, y: m, width: halfW, logoH: 22, logoW: 22, showAddress: false,
    });
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(8)
       .text(title, xOrigin, afterY, { width: halfW / 2, align: 'left', lineBreak: false });
    doc.font('Helvetica').fontSize(8)
       .text(`Page: ${page}`, xOrigin + halfW / 2, afterY, { width: halfW / 2, align: 'right' });
    doc.font('Helvetica-Bold').fontSize(9)
       .text(`${modeHeaderLabel(auction)} No:${auction.ano}`, xOrigin, afterY + 12, { width: halfW / 2, align: 'left' });
    doc.text(`Date:${fmtDateDMY(auction.date)}`, xOrigin + halfW / 2, afterY + 12, { width: halfW / 2, align: 'right' });

    const hy = BODY_TOP;
    doc.rect(xOrigin, hy, halfW, HEAD_H).fillAndStroke('#E8E4DD', '#444');
    doc.fillColor('#000');
    let cx = xOrigin;
    cols.forEach((c, i) => {
      drawFit(c.header, cx + 2, hy + 5, colW[i] - 4, 'center', 8.5, true);
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
      doc.fillColor('#000');
      let cx = xOrigin;
      cols.forEach((c, ci) => {
        const align = c.align || 'left';
        const cellW = colW[ci] - 4;
        if (c.blank) { cx += colW[ci]; return; }
        if (c.fit) {
          // Free-text (buyer / seller name): keep the base size and
          // ellipsize so a long name never spills into the next column.
          doc.font('Helvetica').fontSize(8.5);
          const v = fitText(doc, fmtCell(c.kind, r[c.key]), cellW);
          doc.text(v, cx + 2, ry + 4, { width: cellW, align, lineBreak: false });
        } else {
          // Numeric / short cells: shrink to fit so values never wrap.
          drawFit(fmtCell(c.kind, r[c.key]), cx + 2, ry + 4, cellW, align, 8.5, false);
        }
        cx += colW[ci];
      });
      doc.moveTo(xOrigin, ry + ROW_H).lineTo(xOrigin + halfW, ry + ROW_H)
         .lineWidth(0.25).strokeColor('#999').stroke();
      ry += ROW_H;
    });

    let vx = xOrigin;
    for (let ci = 0; ci < colW.length - 1; ci++) {
      vx += colW[ci];
      doc.moveTo(vx, tblTop).lineTo(vx, ry).lineWidth(0.25).strokeColor('#888').stroke();
    }
    doc.rect(xOrigin, tblTop, halfW, ry - tblTop).lineWidth(0.5).strokeColor('#444').stroke();

    if (isLastPage && totalKeys.length) {
      const ty = ry + 2;
      doc.rect(xOrigin, ty, halfW, ROW_H + 2).fillAndStroke('#FFF3CD', '#E0B020');
      doc.fillColor('#000');
      let cx = xOrigin;
      cols.forEach((c, ci) => {
        let v = '';
        if (ci === 0) v = 'Total';
        else if (totalKeys.includes(c.key)) {
          const sum = rows.reduce((s, r) => s + (Number(r[c.key]) || 0), 0);
          v = fmtCell(c.kind, sum);
        }
        const align = ci === 0 ? 'center' : (c.align || 'right');
        drawFit(v, cx + 2, ty + 5, colW[ci] - 4, align, 8.5, true);
        cx += colW[ci];
      });
      let vx2 = xOrigin;
      for (let ci = 0; ci < colW.length - 1; ci++) {
        vx2 += colW[ci];
        doc.moveTo(vx2, ty).lineTo(vx2, ty + ROW_H + 2).lineWidth(0.25).strokeColor('#E0B020').stroke();
      }
    }
  }

  for (let i = 0; i < totalPages; i++) {
    if (i > 0) doc.addPage();
    const slice = rows.slice(i * ROWS_PER_HALF, (i + 1) * ROWS_PER_HALF);
    const isLast = (i === totalPages - 1);
    drawHalfHeader(m, i + 1);
    drawHalfRows(m, slice, isLast);
    drawHalfHeader(m + halfW + gutter, i + 1);
    drawHalfRows(m + halfW + gutter, slice, isLast);

    const cutX = m + halfW + gutter / 2;
    doc.save();
    doc.dash(3, { space: 3 });
    doc.moveTo(cutX, m).lineTo(cutX, pageH - m).lineWidth(0.5).strokeColor('#888').stroke();
    doc.undash();
    doc.restore();
    doc.font('Helvetica').fontSize(10).fillColor('#888').text('✂', cutX - 4, m - 12, { lineBreak: false });
  }

  return new Promise(resolve => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

// ════════════════════════════════════════════════════════════
// REPORT 2 — COLLECTION (XLSX + PDF, invoice register)
// ════════════════════════════════════════════════════════════
//
// Mirrors COLLECTION.pdf. One row per sales invoice issued:
//   SALE+INVO | TRADE NAME (firm) | NAME (buyer) | QUANTITY | VALUE
// Grouped by buyer state — TAMIL NADU header etc. Total at the end.
//
// We pull from the `invoices` table (one row per invoice, sale='L'/'I').

function getCollectionRows(db, auctionId) {
  // Each invoice header row aggregates across its lots, but in the FoxPro
  // collection report we show one row per buyer/firm. The invoices table
  // already stores one row per invoice with totals, so we group by the
  // (sale, invo, buyer1, place, code) tuple.
  //
  // Column mapping (matches user's data convention):
  //   TRADE NAME ← i.buyer1   (firm/trade name on the invoice itself —
  //                              "EMPEROR SPICES PRIVATE LIMITED",
  //                              "MAR TRADERS", "VARDHAN TRADING COMPANY")
  //   NAME       ← b.buyer    (the full buyer name from the buyers master).
  //                              Falls back to b.sbl, then i.buyer if the
  //                              master record is missing.
  // The buyers master often carries variant/duplicate records that share a
  // trade name (buyer1): e.g. "KARTHIKEYAN SPICES" belongs to MARREESWARAN KKS,
  // PANDI RAVI *and* MAREESWARAN; "ANKIT SPICES" to ASHOK GUPTA AG/AGT/AGT1.
  // A plain LEFT JOIN on (buyer OR buyer1) fans out — one output row per
  // matching master record — so a single invoice repeats several times with
  // identical qty/value. To guarantee ONE row per invoice we resolve the buyer
  // name/state via correlated scalar subqueries (each picks a single master
  // row) instead of joining, and group by the invoice identity so SUM stays
  // correct.
  //
  // CRUCIAL: the buyer NAME (i.buyer) is the unambiguous key; the trade name
  // (i.buyer1) is shared by several buyers, so a buyer1 match can resolve to the
  // WRONG person (e.g. invoice for PANDI RAVI showing MARREESWARAN KKS). We
  // therefore match on the exact (buyer + buyer1) pair first, then on buyer
  // name alone, and only fall back to a trade-name (buyer1) match when the
  // invoice carries no usable buyer name. NOTE: this SQLite build won't resolve
  // an outer-table reference inside a subquery's ORDER BY, so the subqueries
  // correlate only in WHERE and order by bb.id.
  return db.all(`
    SELECT
      i.sale                                                AS sale,
      i.invo                                                AS invo,
      COALESCE(i.buyer1, '')                                AS trade_name,
      COALESCE(
        -- 1) exact buyer + trade-name pair (the true identity)
        (SELECT NULLIF(bb.buyer,'') FROM buyers bb
           WHERE UPPER(TRIM(bb.buyer))  = UPPER(TRIM(i.buyer))
             AND UPPER(TRIM(bb.buyer1)) = UPPER(TRIM(i.buyer1))
             AND NULLIF(bb.buyer,'') IS NOT NULL
           ORDER BY bb.id LIMIT 1),
        -- 2) buyer name alone (unambiguous vs. shared trade names)
        (SELECT NULLIF(bb.buyer,'') FROM buyers bb
           WHERE UPPER(TRIM(bb.buyer))  = UPPER(TRIM(i.buyer))
             AND NULLIF(bb.buyer,'') IS NOT NULL
           ORDER BY bb.id LIMIT 1),
        -- 3) the invoice's own buyer name
        NULLIF(i.buyer,''),
        -- 4) last resort: sbl of any buyer under this trade name
        (SELECT NULLIF(bb.sbl,'') FROM buyers bb
           WHERE UPPER(TRIM(bb.buyer1)) = UPPER(TRIM(i.buyer1))
             AND NULLIF(bb.sbl,'') IS NOT NULL
           ORDER BY bb.id LIMIT 1),
        '')                                                 AS buyer_name,
      COALESCE(
        -- buyer code (buyers.code, e.g. "AG", "AKS") resolved with the same
        -- priority as buyer_name: exact (buyer + trade) pair, then buyer name,
        -- then trade name alone.
        (SELECT NULLIF(bb.code,'') FROM buyers bb
           WHERE UPPER(TRIM(bb.buyer))  = UPPER(TRIM(i.buyer))
             AND UPPER(TRIM(bb.buyer1)) = UPPER(TRIM(i.buyer1))
             AND NULLIF(bb.code,'') IS NOT NULL
           ORDER BY bb.id LIMIT 1),
        (SELECT NULLIF(bb.code,'') FROM buyers bb
           WHERE UPPER(TRIM(bb.buyer))  = UPPER(TRIM(i.buyer))
             AND NULLIF(bb.code,'') IS NOT NULL
           ORDER BY bb.id LIMIT 1),
        (SELECT NULLIF(bb.code,'') FROM buyers bb
           WHERE UPPER(TRIM(bb.buyer1)) = UPPER(TRIM(i.buyer1))
             AND NULLIF(bb.code,'') IS NOT NULL
           ORDER BY bb.id LIMIT 1),
        '')                                                 AS buyer_code,
      SUM(i.qty)                                            AS qty,
      SUM(i.tot)                                            AS value,
      COALESCE(
        -- state of the exact (buyer + trade) match, then buyer-name match,
        -- then a trade-name match — same priority as buyer_name above
        (SELECT bb.state FROM buyers bb
           WHERE UPPER(TRIM(bb.buyer))  = UPPER(TRIM(i.buyer))
             AND UPPER(TRIM(bb.buyer1)) = UPPER(TRIM(i.buyer1))
           ORDER BY bb.id LIMIT 1),
        (SELECT bb.state FROM buyers bb
           WHERE UPPER(TRIM(bb.buyer))  = UPPER(TRIM(i.buyer))
           ORDER BY bb.id LIMIT 1),
        (SELECT bb.state FROM buyers bb
           WHERE UPPER(TRIM(bb.buyer1)) = UPPER(TRIM(i.buyer1))
           ORDER BY bb.id LIMIT 1),
        '')                                                 AS buyer_state
    FROM invoices i
    WHERE i.auction_id = ?
      -- ISP only (IDEAL SPICES PRIVATE LIMITED). The invoices table holds
      -- both the ISP (Tamil Nadu) and its paired ASP (Kerala) row for each
      -- trade; without this filter every buyer appears twice. Matches the
      -- ISP_STATE_SQL classification used by the Tally export.
      AND (UPPER(COALESCE(i.state,'')) = 'TAMIL NADU' OR UPPER(COALESCE(i.state,'')) = '')
    GROUP BY i.sale, i.invo, i.buyer1, i.buyer
    ORDER BY i.sale, CAST(i.invo AS INTEGER), i.invo
  `, [auctionId]);
}

function classifyByState(rows, auctionState) {
  // Group rows by buyer state. Auction's home state goes last (after any
  // out-of-state buyers — matches the FoxPro layout where TAMIL NADU comes
  // first only because it's the only state in this trade).
  const groups = new Map();
  const auctionSt = String(auctionState || '').trim().toUpperCase();
  for (const r of rows) {
    const st = (r.buyer_state || '').trim().toUpperCase() || auctionSt;
    if (!groups.has(st)) groups.set(st, []);
    groups.get(st).push(r);
  }
  return [...groups.entries()].sort(([a], [b]) => {
    // Auction's home state first
    if (a === auctionSt) return -1;
    if (b === auctionSt) return 1;
    return a.localeCompare(b);
  });
}

async function collectionXlsx(db, auctionId) {
  const auction = getAuctionHeader(db, auctionId);
  const rows = getCollectionRows(db, auctionId);
  const groups = classifyByState(rows, auction.state);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Collection');
  ws.columns = [
    { width: 12 }, { width: 30 }, { width: 10 }, { width: 24 }, { width: 14 }, { width: 18 },
  ];

  ws.mergeCells('A1:F1');
  ws.getCell('A1').value = 'IDEAL SPICES PRIVATE LIMITED';
  ws.getCell('A1').font = { bold: true, size: 14 };
  ws.getCell('A1').alignment = { horizontal: 'center' };
  ws.mergeCells('A2:F2');
  ws.getCell('A2').value = `${modeHeaderLabel(auction)} No: ${auction.ano}    Date: ${fmtDateDMY(auction.date)}`;
  ws.getCell('A2').font = { bold: true, size: 11 };
  ws.getCell('A2').alignment = { horizontal: 'center' };
  ws.addRow([]);

  const head = ws.addRow(['INVO', 'TRADE NAME', 'CODE', 'NAME', 'QUANTITY', 'VALUE']);
  head.font = { bold: true };
  head.eachCell(c => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
    c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    c.alignment = { horizontal: 'center' };
  });

  let gQty = 0, gValue = 0;
  groups.forEach(([state, items]) => {
    const sec = ws.addRow([state || 'OTHER']);
    ws.mergeCells(`A${sec.number}:F${sec.number}`);
    sec.font = { bold: true, size: 10 };
    sec.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
    sec.alignment = { horizontal: 'left' };

    items.forEach(it => {
      const r = ws.addRow([
        `${it.sale || ''} ${it.invo || ''}`.trim(),
        it.trade_name || '',
        it.buyer_code || '',
        it.buyer_name || '',
        Number(it.qty) || 0,
        Number(it.value) || 0,
      ]);
      r.getCell(5).numFmt = '#,##0.000';
      r.getCell(5).alignment = { horizontal: 'right' };
      r.getCell(6).numFmt = '#,##,##0.00';
      r.getCell(6).alignment = { horizontal: 'right' };
      gQty += Number(it.qty) || 0;
      gValue += Number(it.value) || 0;
    });
  });

  const tot = ws.addRow(['', '', '', 'Total', gQty, gValue]);
  tot.font = { bold: true };
  tot.getCell(5).numFmt = '#,##0.000';
  tot.getCell(6).numFmt = '#,##,##0.00';
  tot.eachCell((c) => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
    c.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
  });

  return wb.xlsx.writeBuffer();
}

async function collectionPdf(db, auctionId) {
  const auction = getAuctionHeader(db, auctionId);
  const rows = getCollectionRows(db, auctionId);
  const groups = classifyByState(rows, auction.state);

  const doc = new PDFDocument({ size: 'A4', layout: 'portrait', margin: 24 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));

  const m = 24;
  const pageW = doc.page.width;
  const pageH = doc.page.height;
  const usableW = pageW - m * 2;

  // Cols: INVO | TRADE NAME | CODE | NAME | QTY | VALUE
  // Portrait A4 → ~547pt usable. Long firm names like "PERFECT CARDAMOM AND
  // SPICES MARKETING COMPANY PVT LTD" wrap to multiple lines via wrapText
  // (drawDataRow grows the row height to fit the tallest cell).
  const colW = [
    Math.floor(usableW * 0.10),  // INVO
    Math.floor(usableW * 0.27),  // TRADE NAME
    Math.floor(usableW * 0.08),  // CODE
    Math.floor(usableW * 0.25),  // NAME
    Math.floor(usableW * 0.13),  // QUANTITY
    0,                           // VALUE absorbs rounding (~17%)
  ];
  colW[5] = usableW - colW[0] - colW[1] - colW[2] - colW[3] - colW[4];
  const colX = [m];
  for (let i = 0; i < colW.length - 1; i++) colX.push(colX[i] + colW[i]);
  // Per-column text alignment, shared by header + data + total rows.
  const COL_ALIGN = ['center', 'left', 'left', 'left', 'right', 'right'];

  const ROW_H = 16;
  const HEAD_H = 18;
  const SECT_H = 16;
  let y;
  let pageStartY;

  // Segment tracking — verticals only inside data segments, never through
  // full-width section header strips or the Total row strip.
  const dataSegments = [];   // per-page segments
  let curSegStart = null;

  function startSegment() { if (curSegStart === null) curSegStart = y; }
  function closeSegment() {
    if (curSegStart === null) return;
    if (y > curSegStart + 0.5) {
      const segs = dataSegments[dataSegments.length - 1];
      segs.push({ top: curSegStart, bottom: y });
    }
    curSegStart = null;
  }

  // Resolve company branding once (logo + name + address) for use across pages
  const companyHeader = getCompanyHeader(db);

  function drawTopHeader() {
    // Three-column brand band: company on left, "Collection" centered,
    // trade meta right-aligned. The page number changes per page.
    const afterY = drawCompanyHeader(doc, companyHeader, {
      x: m, y: m, width: usableW,
      title: 'COLLECTION',
      metaLines: [
        `${modeHeaderLabel(auction)} No: ${auction.ano}`,
        `Date: ${fmtDateDMY(auction.date)}`,
        `Page: ${doc.bufferedPageRange().count}`,
      ],
    });
    y = afterY;
    pageStartY = y;
    dataSegments.push([]);
  }

  function drawColHeader() {
    const headTop = y;
    doc.rect(m, y, usableW, HEAD_H).fillAndStroke('#E8E4DD', '#444');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(9.5);
    const heads = ['INVO', 'TRADE NAME', 'CODE', 'NAME', 'QUANTITY', 'VALUE'];
    heads.forEach((h, i) => {
      doc.text(fitText(doc, h, colW[i] - 8), colX[i] + 4, y + 5, {
        width: colW[i] - 8, align: COL_ALIGN[i], lineBreak: false,
      });
    });
    y += HEAD_H;
    curSegStart = headTop;  // header strip + following data rows = one segment
  }

  function drawSectionRow(label) {
    if (y + SECT_H > pageH - m - 12) {
      finishPage(); doc.addPage(); drawTopHeader(); drawColHeader();
    }
    closeSegment();  // section header is NOT in a data segment
    doc.rect(m, y, usableW, SECT_H).fillAndStroke('#FFF3CD', '#9A6700');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(10);
    doc.text(label, m + 6, y + 4, { width: usableW - 12, align: 'left', lineBreak: false });
    y += SECT_H;
  }

  function drawDataRow(it, idx) {
    doc.font('Helvetica').fontSize(9);
    const invoDisplay = `${it.sale || ''} ${it.invo || ''}`.trim();
    const cells = [invoDisplay, it.trade_name || '', it.buyer_code || '',
                   it.buyer_name || '', fmtQty(it.qty), fmtMoney(it.value)];
    const aligns = COL_ALIGN;

    const LINE_H = 12;
    const PAD_TOP = 4, PAD_BOT = 4;
    const wrapped = cells.map((v, ci) => wrapText(doc, v, colW[ci] - 8));
    const maxLines = Math.max(1, ...wrapped.map(ls => ls.length));
    const rowH = maxLines * LINE_H + PAD_TOP + PAD_BOT;

    if (y + rowH > pageH - m - 12) {
      finishPage(); doc.addPage(); drawTopHeader(); drawColHeader();
    }
    startSegment();
    // Inset the fill's top edge by the rule width so it doesn't paint over
    // the previous row's separator (drawn at this row's top y).
    if (idx % 2 === 1) doc.rect(m, y + 0.5, usableW, rowH - 0.5).fill('#F7F5F2');
    doc.fillColor('#000').font('Helvetica').fontSize(9);

    cells.forEach((v, ci) => {
      const lines = wrapped[ci];
      lines.forEach((line, li) => {
        doc.text(line, colX[ci] + 4, y + PAD_TOP + li * LINE_H, {
          width: colW[ci] - 8, align: aligns[ci], lineBreak: false,
        });
      });
    });
    doc.moveTo(m, y + rowH).lineTo(m + usableW, y + rowH)
       .lineWidth(0.25).strokeColor('#CCC').stroke();
    y += rowH;
  }

  // Draw verticals (only inside data segments) + outer border for current page.
  function finishPage() {
    if (pageStartY === null || pageStartY === undefined) return;
    closeSegment();
    const top = pageStartY;
    const bottom = y;
    const segs = dataSegments[dataSegments.length - 1] || [];
    for (let ci = 0; ci < colW.length - 1; ci++) {
      const vx = colX[ci] + colW[ci];
      segs.forEach(s => {
        doc.moveTo(vx, s.top).lineTo(vx, s.bottom).lineWidth(0.3).strokeColor('#888').stroke();
      });
    }
    doc.rect(m, top, usableW, bottom - top).lineWidth(0.5).strokeColor('#444').stroke();
    pageStartY = null;
  }

  drawTopHeader();
  drawColHeader();

  let gQty = 0, gValue = 0;
  let rowIdx = 0;
  groups.forEach(([state, items]) => {
    drawSectionRow(state || 'OTHER');
    items.forEach(it => {
      drawDataRow(it, rowIdx++);
      gQty += Number(it.qty) || 0;
      gValue += Number(it.value) || 0;
    });
  });

  // Total row — full-width strip, close segment first
  if (y + ROW_H + 4 > pageH - m - 12) {
    finishPage(); doc.addPage(); drawTopHeader(); drawColHeader();
  }
  closeSegment();
  doc.rect(m, y, usableW, ROW_H + 4).fillAndStroke('#FFF3CD', '#E0B020');
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(10);
  // The total occupies cols: empty | empty | empty | "Total" | qty | value
  doc.text('Total', colX[3] + 4, y + 6, { width: colW[3] - 8, align: 'right', lineBreak: false });
  doc.text(fmtQty(gQty),   colX[4] + 4, y + 6, { width: colW[4] - 8, align: 'right', lineBreak: false });
  doc.text(fmtMoney(gValue), colX[5] + 4, y + 6, { width: colW[5] - 8, align: 'right', lineBreak: false });
  y += ROW_H + 4;

  finishPage();

  return new Promise(resolve => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

// ════════════════════════════════════════════════════════════
// REPORT 3 — TRADE REPORT (XLSX + PDF, BUYERS LIST FOR VERIFICATION)
// ════════════════════════════════════════════════════════════
//
// Mirrors TRADE_REPORT.pdf. Per buyer-code in the auction:
//   SALE | BIDDER (short name) | TRADE NAME (firm) | BAG | QUANTITY | AMOUNT | INV.AMOUNT | CODE
// Grouped by state, with subtotals:
//   - INTER-STATE SALES (within state)
//   - INTRA STATE SALES (within state)
//   - <state> STATE TOTAL
//   - GRAND TOTAL
// Plus a footer block with TOTAL ARRIVALS / WITHDRAWN / SOLD / NOT eTRADED
// counts and MAX/MIN/AVERAGE prices and COST OF CARDAMOM.

function getTradeReportData(db, auctionId) {
  const auction = getAuctionHeader(db, auctionId);

  // One row per buyer code, summed across all lots.
  // INV.AMOUNT comes from the invoices table (sum of `tot` per buyer).
  // We join on buyers master to get state and full buyer info.
  //
  // Column mapping (matches user's data convention):
  //   TRADE NAME ← b.buyer1     (firm/trade name as user stores it —
  //                                "EMPEROR SPICES PRIVATE LIMITED",
  //                                "MAR TRADERS", "VARDHAN TRADING COMPANY")
  //   BIDDER     ← b.sbl/b.code  (proprietor name from sbl if distinct,
  //                                falls back to short code)
  const rows = db.all(`
    SELECT
      l.code                                                  AS code,
      COALESCE(NULLIF(b.sbl,    b.buyer1),
               NULLIF(b.code,   ''),
               l.buyer1,
               '')                                            AS bidder,
      COALESCE(b.buyer, l.buyer, '')                          AS buyer_name,
      COALESCE(b.buyer1, l.buyer1, '')                        AS trade_name,
      COALESCE(b.state,  '')                                  AS state,
      COALESCE(b.gstin,  '')                                  AS gstin,
      COALESCE(l.sale,   '')                                  AS lot_sale,
      SUM(l.bags)                                             AS bag,
      SUM(l.qty)                                              AS qty,
      SUM(l.amount)                                           AS amount
    FROM lots l
    LEFT JOIN buyers b
      ON UPPER(TRIM(b.code))  = UPPER(TRIM(l.code))
      OR UPPER(TRIM(b.buyer)) = UPPER(TRIM(l.buyer))
    WHERE l.auction_id = ?
      AND l.amount > 0
    GROUP BY l.code, b.buyer, l.buyer, b.buyer1, b.sbl, b.code, b.state, b.gstin, l.sale
    ORDER BY UPPER(COALESCE(b.buyer1, l.buyer1, l.code)), l.code
  `, [auctionId]);

  // INV.AMOUNT per code from invoices table
  const invRows = db.all(`
    SELECT
      COALESCE(b.code, '')       AS code,
      SUM(i.tot)                 AS inv_amount
    FROM invoices i
    LEFT JOIN buyers b
      ON UPPER(TRIM(b.buyer))  = UPPER(TRIM(i.buyer))
      OR UPPER(TRIM(b.buyer1)) = UPPER(TRIM(i.buyer1))
    WHERE i.auction_id = ?
    GROUP BY b.code
  `, [auctionId]);
  const invByCode = {};
  invRows.forEach(r => { if (r.code) invByCode[r.code] = Number(r.inv_amount) || 0; });

  // Stamp each buyer-row with inv_amount and a uniform sale code (I/L)
  const auctionState = String(auction.state || '').trim().toUpperCase();
  const cfg = getSettingsFlat(db);
  // e-Trade: the BIDDER column shows the buyer name (e.g. "NAGENDRAN")
  // rather than the proprietor short-name used for e-Auction.
  const isETrade = auction && auction.mode === 'e-Trade';
  // Home state for the inter/intra split. e-Trade compares the buyer against
  // the configured business_state (the selling company's own state); e-Auction
  // keeps comparing against the auction's stored state.
  const homeState = isETrade
    ? String(cfg.business_state || auction.state || '').trim().toUpperCase()
    : auctionState;
  // e-Trade INV.AMOUNT is COMPUTED from each buyer's lot aggregates so it's
  // available even before sales invoices are generated (e-Auction keeps reading
  // the stored invoice grand-totals via invByCode). Transport & Insurance are
  // billed on LOCAL sales only ('L'); inter-state buyers cover their own
  // freight, so both are 0 for 'I'. Rounded to the nearest rupee:
  //   cardamomGunny = Σ amount + (Σ bags × Gunny rate)
  //   transport     = Σ qty × Transport (₹/kg)           [Local only, else 0]
  //   insurance     = ((cardamomGunny + cardamomGunny × Goods GST %) / 1000)
  //                     × Insurance (₹/₹1000)             [Local only, else 0]
  //   gst           = (cardamomGunny + transport + insurance) × Goods GST %
  //   INV.AMOUNT    = round( cardamomGunny + transport + insurance + gst )
  const _gunnyRate     = Number(cfg.gunny_rate) || 165;
  const _transportRate = Number(cfg.transport)  || 2.5;
  const _insuranceRate = Number(cfg.insurance)  || 0.75;
  const _gstGoods      = Number(cfg.gst_goods)  || 5;
  const _tradeInvAmount = (r) => {
    const bags = Number(r.bag) || 0;
    const qty = Number(r.qty) || 0;
    const cardamomGunny = (Number(r.amount) || 0) + bags * _gunnyRate;
    const isLocal = r.sale === 'L';
    const transport = isLocal ? qty * _transportRate : 0;
    // Insurance base = goods value grossed up by Goods GST (per ₹1000).
    const goodsWithGst = cardamomGunny + cardamomGunny * _gstGoods / 100;
    const insurance = isLocal ? (goodsWithGst / 1000) * _insuranceRate : 0;
    const gst = (cardamomGunny + transport + insurance) * _gstGoods / 100;
    // Rounded to the nearest rupee.
    return Math.round(cardamomGunny + transport + insurance + gst);
  };
  rows.forEach(r => {
    if (isETrade) r.bidder = r.buyer_name || '';
    // Sale type first (INV.AMOUNT depends on it: transport/insurance are
    // Local-only). Inter-state if buyer state ≠ home state; empty buyer
    // state defaults to intra-state (matches the FoxPro fallback).
    const buyerSt = String(r.state || '').trim().toUpperCase();
    r.sale = (buyerSt && buyerSt !== homeState) ? 'I' : 'L';
    r.inv_amount = isETrade ? _tradeInvAmount(r) : (invByCode[r.code] || 0);
  });

  // Group by buyer-state (column on the report). Within each state, split into
  // inter-state vs intra-state sub-buckets for the section subtotals.
  const stateGroups = new Map();
  for (const r of rows) {
    const st = (r.state || '').trim().toUpperCase() || auctionState;
    if (!stateGroups.has(st)) stateGroups.set(st, { inter: [], intra: [] });
    if (r.sale === 'I') stateGroups.get(st).inter.push(r);
    else                stateGroups.get(st).intra.push(r);
  }
  // Sort: auction's home state first
  const sortedStates = [...stateGroups.entries()].sort(([a], [b]) => {
    if (a === auctionState) return -1;
    if (b === auctionState) return 1;
    return a.localeCompare(b);
  });

  // Statistics for the footer
  const allLots = db.all(
    `SELECT bags, qty, price, amount, code FROM lots WHERE auction_id = ?`, [auctionId]
  );
  // e-Trade: Withdrawn lots (code = 'WD') are reported in the WITHDRAWN row,
  // separate from the NOT eTRADED (unsold, non-withdrawn) bucket where they
  // used to land. e-Auction keeps its prior behaviour (no WD split).
  const isWd = (l) => isETrade && String(l.code || '').trim().toUpperCase() === 'WD';
  const withdrawn = allLots.filter(isWd);
  const sold = allLots.filter(l => !isWd(l) && Number(l.amount) > 0);
  const notSold = allLots.filter(l => !isWd(l) && !(Number(l.amount) > 0));
  const sumLots = (xs, k) => xs.reduce((s, r) => s + (Number(r[k]) || 0), 0);
  const stats = {
    arrivals_qty:  sumLots(allLots, 'qty'),
    arrivals_bags: sumLots(allLots, 'bags'),
    arrivals_lots: allLots.length,
    withdrawn_qty:  sumLots(withdrawn, 'qty'),
    withdrawn_bags: sumLots(withdrawn, 'bags'),
    withdrawn_lots: withdrawn.length,
    sold_qty:      sumLots(sold, 'qty'),
    sold_bags:     sumLots(sold, 'bags'),
    sold_lots:     sold.length,
    not_qty:       sumLots(notSold, 'qty'),
    not_bags:      sumLots(notSold, 'bags'),
    not_lots:      notSold.length,
    cost:          sumLots(sold, 'amount'),
    max_price:     sold.length ? Math.max(...sold.map(l => Number(l.price) || 0)) : 0,
    min_price:     sold.length ? Math.min(...sold.map(l => Number(l.price) || 0)) : 0,
    avg_price:     0,
  };
  if (stats.sold_qty > 0) stats.avg_price = stats.cost / stats.sold_qty;

  return { auction, sortedStates, stats };
}

async function tradeReportXlsx(db, auctionId) {
  const { auction, sortedStates, stats } = getTradeReportData(db, auctionId);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('TradeReport');
  ws.columns = [
    { width: 6 },   // SALE
    { width: 22 },  // BIDDER
    { width: 26 },  // TRADE NAME
    { width: 6 },   // BAG
    { width: 12 },  // QUANTITY
    { width: 16 },  // AMOUNT
    { width: 16 },  // INV.AMOUNT
    { width: 8 },   // CODE
  ];

  ws.mergeCells('A1:H1');
  ws.getCell('A1').value = 'IDEAL SPICES PRIVATE LIMITED';
  ws.getCell('A1').font = { bold: true, size: 14 };
  ws.getCell('A1').alignment = { horizontal: 'center' };

  ws.mergeCells('A2:H2');
  ws.getCell('A2').value = 'BUYERS LIST FOR VERIFICATION';
  ws.getCell('A2').font = { bold: true, size: 12 };
  ws.getCell('A2').alignment = { horizontal: 'center' };

  ws.mergeCells('A3:H3');
  ws.getCell('A3').value = `${modeHeaderLabel(auction)} No: ${auction.ano}    DATE: ${fmtDateDMY(auction.date)}`;
  ws.getCell('A3').font = { bold: true, size: 11 };
  ws.getCell('A3').alignment = { horizontal: 'center' };

  ws.addRow([]);

  const head = ws.addRow(['SALE', 'BIDDER', 'TRADE NAME', 'BAG', 'QUANTITY', 'AMOUNT', 'INV.AMOUNT', 'CODE']);
  head.font = { bold: true };
  head.eachCell(c => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
    c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    c.alignment = { horizontal: 'center' };
  });

  function emitRow(r) {
    const row = ws.addRow([
      r.sale,
      r.bidder || '',
      r.trade_name || '',
      Number(r.bag) || 0,
      Number(r.qty) || 0,
      Number(r.amount) || 0,
      Number(r.inv_amount) || 0,
      r.code || '',
    ]);
    row.getCell(4).alignment = { horizontal: 'right' };
    row.getCell(5).alignment = { horizontal: 'right' };
    row.getCell(5).numFmt = '#,##0.000';
    row.getCell(6).alignment = { horizontal: 'right' };
    row.getCell(6).numFmt = '#,##,##0.00';
    row.getCell(7).alignment = { horizontal: 'right' };
    row.getCell(7).numFmt = '#,##,##0.00';
  }

  function emitSubtotal(label, items) {
    const sumKey = k => items.reduce((s, r) => s + (Number(r[k]) || 0), 0);
    const tot = ws.addRow([
      '', label, '',
      sumKey('bag'), sumKey('qty'), sumKey('amount'), '', '',
    ]);
    ws.mergeCells(`B${tot.number}:C${tot.number}`);
    tot.font = { bold: true };
    tot.getCell(5).numFmt = '#,##0.000';
    tot.getCell(6).numFmt = '#,##,##0.00';
    tot.eachCell((c) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F4EA' } };
      c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    });
    return { bag: sumKey('bag'), qty: sumKey('qty'), amount: sumKey('amount') };
  }

  let gBag = 0, gQty = 0, gAmt = 0;
  sortedStates.forEach(([state, group]) => {
    const stateRow = ws.addRow([state]);
    ws.mergeCells(`A${stateRow.number}:H${stateRow.number}`);
    stateRow.font = { bold: true, size: 11 };
    stateRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
    stateRow.alignment = { horizontal: 'left' };

    let stBag = 0, stQty = 0, stAmt = 0;
    if (group.inter.length) {
      group.inter.forEach(emitRow);
      const t = emitSubtotal('INTER-STATE SALES', group.inter);
      stBag += t.bag; stQty += t.qty; stAmt += t.amount;
    }
    if (group.intra.length) {
      group.intra.forEach(emitRow);
      const t = emitSubtotal('INTRA STATE SALES', group.intra);
      stBag += t.bag; stQty += t.qty; stAmt += t.amount;
    }
    // State total
    const stRow = ws.addRow(['', `${state} STATE TOTAL`, '', stBag, stQty, stAmt, '', '']);
    ws.mergeCells(`B${stRow.number}:C${stRow.number}`);
    stRow.font = { bold: true };
    stRow.getCell(5).numFmt = '#,##0.000';
    stRow.getCell(6).numFmt = '#,##,##0.00';
    stRow.eachCell((c) => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
      c.border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
    });
    gBag += stBag; gQty += stQty; gAmt += stAmt;
  });

  // Grand total
  const grand = ws.addRow(['', 'GRAND TOTAL', '', gBag, gQty, gAmt, '', '']);
  ws.mergeCells(`B${grand.number}:C${grand.number}`);
  grand.font = { bold: true, size: 11 };
  grand.getCell(5).numFmt = '#,##0.000';
  grand.getCell(6).numFmt = '#,##,##0.00';
  grand.eachCell((c) => {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' } };
    c.border = { top: { style: 'double' }, bottom: { style: 'double' } };
  });

  // Footer stats block
  ws.addRow([]);
  const addFootRow = (lhs, rhs) => {
    const r = ws.addRow([lhs, '', '', '', '', rhs, '', '']);
    ws.mergeCells(`A${r.number}:E${r.number}`);
    ws.mergeCells(`F${r.number}:H${r.number}`);
    r.font = { size: 10 };
    return r;
  };
  addFootRow(
    `TOTAL ARRIVALS  Kgs. ${fmtQty(stats.arrivals_qty)}  Bags. ${stats.arrivals_bags}  Lot. ${stats.arrivals_lots}`,
    `MAXIMUM Rs. ${fmtMoney(stats.max_price)}`,
  );
  addFootRow(
    `WITH DRAWN     Kgs. ${fmtQty(stats.withdrawn_qty)}  Bags. ${stats.withdrawn_bags}  Lot. ${stats.withdrawn_lots}`,
    `MINIMUM Rs. ${fmtMoney(stats.min_price)}`,
  );
  addFootRow(
    `SOLD           Kgs. ${fmtQty(stats.sold_qty)}  Bags. ${stats.sold_bags}  Lot. ${stats.sold_lots}`,
    `AVERAGE Rs. ${fmtMoney(stats.avg_price)}`,
  );
  addFootRow(
    `${modeNotXyzLabel(auction)}    Kgs. ${fmtQty(stats.not_qty)}  Bags. ${stats.not_bags}  Lot. ${stats.not_lots}`,
    '',
  );
  addFootRow(
    `COST OF CARDAMOM   Rs. ${fmtMoney(stats.cost)}`,
    '',
  );

  return wb.xlsx.writeBuffer();
}

async function tradeReportPdf(db, auctionId) {
  const { auction, sortedStates, stats } = getTradeReportData(db, auctionId);

  const doc = new PDFDocument({ size: 'A4', layout: 'portrait', margin: 18 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));

  const m = 18;
  const pageW = doc.page.width;
  const pageH = doc.page.height;
  const usableW = pageW - m * 2;

  // Cols: SL.NO | SALE | BIDDER | TRADE NAME | BAG | QTY | AMOUNT | INV.AMOUNT | CODE
  // Portrait A4 → ~559pt usable. Long firm/person names wrap to multiple
  // lines via wrapText (drawDataRow grows the row height to fit). Numeric
  // columns are sized to fit Indian-format money values without truncation.
  const colW = [
    Math.floor(usableW * 0.06),  // SL.NO
    Math.floor(usableW * 0.06),  // SALE
    Math.floor(usableW * 0.16),  // BIDDER
    Math.floor(usableW * 0.18),  // TRADE NAME
    Math.floor(usableW * 0.05),  // BAG
    Math.floor(usableW * 0.11),  // QTY
    Math.floor(usableW * 0.16),  // AMOUNT
    Math.floor(usableW * 0.16),  // INV.AMOUNT
    0,                           // CODE absorbs rounding
  ];
  colW[8] = usableW - colW.slice(0, 8).reduce((s, w) => s + w, 0);
  const colX = [m];
  for (let i = 0; i < colW.length - 1; i++) colX.push(colX[i] + colW[i]);

  const ROW_H = 14;
  const HEAD_H = 18;
  const SECT_H = 16;
  let y;
  let pageStartY;
  let pageNum = 0;

  // Segment tracking — verticals only inside column-header + data-row regions,
  // never through full-width section headers, subtotal strips, state-total
  // strips, or the grand-total strip.
  const dataSegments = [];   // per-page list of {top, bottom} segments
  let curSegStart = null;

  function startSegment() { if (curSegStart === null) curSegStart = y; }
  function closeSegment() {
    if (curSegStart === null) return;
    if (y > curSegStart + 0.5) {
      const segs = dataSegments[dataSegments.length - 1];
      segs.push({ top: curSegStart, bottom: y });
    }
    curSegStart = null;
  }

  // Resolve company branding once (logo + name + address) for use across pages
  const companyHeader = getCompanyHeader(db);

  function drawTopHeader() {
    pageNum += 1;
    // Three-column brand band: company on left, "BUYERS LIST FOR VERIFICATION"
    // centered, trade meta right-aligned. Page number updates per page.
    const afterY = drawCompanyHeader(doc, companyHeader, {
      x: m, y: m, width: usableW,
      title: 'BUYERS LIST FOR VERIFICATION',
      metaLines: [
        `${modeHeaderLabel(auction)} No: ${auction.ano}`,
        `Date: ${fmtDateDMY(auction.date)}`,
        `Page: ${pageNum}`,
      ],
    });
    y = afterY;
    pageStartY = y;
    dataSegments.push([]);
  }

  function drawColHeader() {
    const headTop = y;
    doc.rect(m, y, usableW, HEAD_H).fillAndStroke('#E8E4DD', '#444');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(8.5);
    const heads = ['S.NO', 'SALE', 'BIDDER', 'TRADE NAME', 'BAG', 'QUANTITY', 'AMOUNT', 'INV.AMOUNT', 'CODE'];
    const aligns = ['center', 'center', 'left', 'left', 'center', 'right', 'right', 'right', 'center'];
    heads.forEach((h, i) => {
      doc.text(fitText(doc, h, colW[i] - 8), colX[i] + 4, y + 5, {
        width: colW[i] - 8, align: aligns[i], lineBreak: false,
      });
    });
    y += HEAD_H;
    curSegStart = headTop;
  }

  // Draws verticals + outer table border on the current page.
  function finishPage() {
    if (pageStartY === null || pageStartY === undefined) return;
    closeSegment();
    const top = pageStartY, bottom = y;
    // Draw verticals only inside data segments (header strip + lot rows)
    const segs = dataSegments[dataSegments.length - 1] || [];
    for (let ci = 0; ci < colW.length - 1; ci++) {
      const vx = colX[ci] + colW[ci];
      segs.forEach(s => {
        doc.moveTo(vx, s.top).lineTo(vx, s.bottom).lineWidth(0.3).strokeColor('#888').stroke();
      });
    }
    // Outer table border still spans the whole page table area
    doc.rect(m, top, usableW, bottom - top).lineWidth(0.5).strokeColor('#444').stroke();
    pageStartY = null;
  }

  function ensureRoom(needed) {
    if (y + needed > pageH - m - 12) {
      finishPage(); doc.addPage(); drawTopHeader(); drawColHeader();
    }
  }

  function drawStateRow(label) {
    ensureRoom(SECT_H + ROW_H);
    closeSegment();  // section header is NOT in a data segment
    doc.rect(m, y, usableW, SECT_H).fillAndStroke('#FFF3CD', '#9A6700');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(10);
    doc.text(label, m + 6, y + 4, { width: usableW - 12, align: 'left', lineBreak: false });
    y += SECT_H;
  }

  function drawDataRow(r, idx) {
    doc.font('Helvetica').fontSize(8.5);
    const cells = [
      String(idx + 1),         // SL.NO
      r.sale,
      r.bidder || '',
      r.trade_name || '',
      String(r.bag || 0),
      fmtQty(r.qty),
      fmtMoney(r.amount),
      fmtMoney(r.inv_amount),
      r.code || '',
    ];
    const aligns = ['center', 'center', 'left', 'left', 'center', 'right', 'right', 'right', 'center'];

    // Wrap each cell text into lines that fit its column width. Use the
    // tallest cell to set this row's height so nothing overflows.
    const LINE_H = 11;
    const PAD_TOP = 3, PAD_BOT = 3;
    const wrapped = cells.map((v, ci) => wrapText(doc, v, colW[ci] - 8));
    const maxLines = Math.max(1, ...wrapped.map(ls => ls.length));
    const rowH = maxLines * LINE_H + PAD_TOP + PAD_BOT;

    ensureRoom(rowH);
    startSegment();
    // Inset the fill's top edge by the rule width so it doesn't paint over
    // the previous row's separator (drawn at this row's top y).
    if (idx % 2 === 1) doc.rect(m, y + 0.5, usableW, rowH - 0.5).fill('#F7F5F2');
    doc.fillColor('#000').font('Helvetica').fontSize(8.5);

    cells.forEach((v, ci) => {
      const lines = wrapped[ci];
      lines.forEach((line, li) => {
        doc.text(line, colX[ci] + 4, y + PAD_TOP + li * LINE_H, {
          width: colW[ci] - 8, align: aligns[ci], lineBreak: false,
        });
      });
    });
    doc.moveTo(m, y + rowH).lineTo(m + usableW, y + rowH)
       .lineWidth(0.2).strokeColor('#CCC').stroke();
    y += rowH;
  }

  function drawSubtotal(label, items, color) {
    ensureRoom(ROW_H + 2);
    closeSegment();  // subtotal strip is NOT in a data segment
    doc.rect(m, y, usableW, ROW_H + 2).fillAndStroke(color || '#E6F4EA', '#5A8F62');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
    const sum = k => items.reduce((s, r) => s + (Number(r[k]) || 0), 0);
    // Label spans SL.NO + SALE + BIDDER + TRADE NAME (cols 0..3)
    const labelW = colW[0] + colW[1] + colW[2] + colW[3] - 8;
    doc.text(fitText(doc, label, labelW), colX[0] + 4, y + 4, {
      width: labelW, align: 'left', lineBreak: false,
    });
    doc.text(fitText(doc, String(sum('bag')),      colW[4] - 8), colX[4] + 4, y + 4, { width: colW[4] - 8, align: 'center', lineBreak: false });
    doc.text(fitText(doc, fmtQty(sum('qty')),      colW[5] - 8), colX[5] + 4, y + 4, { width: colW[5] - 8, align: 'right',  lineBreak: false });
    doc.text(fitText(doc, fmtMoney(sum('amount')), colW[6] - 8), colX[6] + 4, y + 4, { width: colW[6] - 8, align: 'right',  lineBreak: false });
    y += ROW_H + 2;
    return { bag: sum('bag'), qty: sum('qty'), amount: sum('amount') };
  }

  drawTopHeader();
  drawColHeader();

  let gBag = 0, gQty = 0, gAmt = 0;
  sortedStates.forEach(([state, group]) => {
    drawStateRow(state);
    // Serial number (and zebra striping) restart at the beginning of each
    // state so the rows count Tamil Nadu 1..N independently of Kerala 1..M.
    let rowIdx = 0;
    let stBag = 0, stQty = 0, stAmt = 0;
    if (group.inter.length) {
      group.inter.forEach(r => drawDataRow(r, rowIdx++));
      const t = drawSubtotal('INTER-STATE SALES', group.inter, '#E6F4EA');
      stBag += t.bag; stQty += t.qty; stAmt += t.amount;
    }
    if (group.intra.length) {
      group.intra.forEach(r => drawDataRow(r, rowIdx++));
      const t = drawSubtotal('INTRA STATE SALES', group.intra, '#E6F4EA');
      stBag += t.bag; stQty += t.qty; stAmt += t.amount;
    }
    // State total (use a slightly different colour) — also full-width strip,
    // so close the data segment first.
    ensureRoom(ROW_H + 2);
    closeSegment();
    doc.rect(m, y, usableW, ROW_H + 2).fillAndStroke('#FFF3CD', '#9A6700');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
    const stLabelW = colW[0] + colW[1] + colW[2] + colW[3] - 8;
    doc.text(fitText(doc, `${state} STATE TOTAL`, stLabelW), colX[0] + 4, y + 4, {
      width: stLabelW, align: 'left', lineBreak: false,
    });
    doc.text(fitText(doc, String(stBag),       colW[4] - 8), colX[4] + 4, y + 4, { width: colW[4] - 8, align: 'center', lineBreak: false });
    doc.text(fitText(doc, fmtQty(stQty),       colW[5] - 8), colX[5] + 4, y + 4, { width: colW[5] - 8, align: 'right',  lineBreak: false });
    doc.text(fitText(doc, fmtMoney(stAmt),     colW[6] - 8), colX[6] + 4, y + 4, { width: colW[6] - 8, align: 'right',  lineBreak: false });
    y += ROW_H + 2;
    gBag += stBag; gQty += stQty; gAmt += stAmt;
  });

  // Grand total — full-width strip, so close segment first
  ensureRoom(ROW_H + 4);
  closeSegment();
  doc.rect(m, y, usableW, ROW_H + 4).fillAndStroke('#FFF3CD', '#7A4400');
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(10);
  const gtLabelW = colW[0] + colW[1] + colW[2] + colW[3] - 8;
  doc.text(fitText(doc, 'GRAND TOTAL', gtLabelW), colX[0] + 4, y + 5, {
    width: gtLabelW, align: 'left', lineBreak: false,
  });
  doc.text(fitText(doc, String(gBag),         colW[4] - 8), colX[4] + 4, y + 5, { width: colW[4] - 8, align: 'center', lineBreak: false });
  doc.text(fitText(doc, fmtQty(gQty),         colW[5] - 8), colX[5] + 4, y + 5, { width: colW[5] - 8, align: 'right',  lineBreak: false });
  doc.text(fitText(doc, fmtMoney(gAmt),       colW[6] - 8), colX[6] + 4, y + 5, { width: colW[6] - 8, align: 'right',  lineBreak: false });
  y += ROW_H + 4;

  finishPage();

  // ── Footer: stats summary as two clean tables ──────────────────
  // Left  table — INVENTORY: STATUS | KGS | BAGS | LOTS  (4 rows)
  // Right table — PRICE STATS: METRIC | RS              (3 rows)
  // Bottom — COST OF CARDAMOM full-width, prominent.
  const FOOT_GAP        = 16;
  const FOOT_HEAD_H     = 18;
  const FOOT_ROW_H      = 16;
  const FOOT_COST_H     = 20;
  const FOOT_BLOCK_H    = FOOT_HEAD_H + FOOT_ROW_H * 4 + FOOT_COST_H + 8;

  if (y + FOOT_GAP + FOOT_BLOCK_H > pageH - m) { doc.addPage(); y = m; }
  y += FOOT_GAP;

  // Left & right tables share the page width with a small gutter
  const gutter = 12;
  const leftW  = Math.floor((usableW - gutter) * 0.62);
  const rightW = usableW - gutter - leftW;
  const leftX  = m;
  const rightX = m + leftW + gutter;

  // Left INVENTORY table: STATUS (45%) | KGS (22%) | BAGS (16%) | LOTS (17%)
  const lcW = [
    Math.floor(leftW * 0.45),
    Math.floor(leftW * 0.22),
    Math.floor(leftW * 0.16),
    0,
  ];
  lcW[3] = leftW - lcW[0] - lcW[1] - lcW[2];
  const lcX = [leftX];
  for (let i = 0; i < lcW.length - 1; i++) lcX.push(lcX[i] + lcW[i]);

  // Header strip
  doc.rect(leftX, y, leftW, FOOT_HEAD_H).fillAndStroke('#E8E4DD', '#444');
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
  ['STATUS', 'KGS', 'BAGS', 'LOTS'].forEach((h, i) => {
    const align = i === 0 ? 'left' : (i === 1 ? 'right' : 'center');
    doc.text(h, lcX[i] + 6, y + 5, { width: lcW[i] - 12, align, lineBreak: false });
  });
  let ly = y + FOOT_HEAD_H;

  function leftRow(label, kgs, bags, lots, idx) {
    if (idx % 2 === 1) doc.rect(leftX, ly, leftW, FOOT_ROW_H).fill('#F7F5F2');
    doc.fillColor('#000').font('Helvetica').fontSize(9);
    doc.text(label, lcX[0] + 6, ly + 4, { width: lcW[0] - 12, align: 'left', lineBreak: false });
    doc.text(kgs,   lcX[1] + 6, ly + 4, { width: lcW[1] - 12, align: 'right', lineBreak: false });
    doc.text(bags,  lcX[2] + 6, ly + 4, { width: lcW[2] - 12, align: 'center', lineBreak: false });
    doc.text(lots,  lcX[3] + 6, ly + 4, { width: lcW[3] - 12, align: 'center', lineBreak: false });
    doc.moveTo(leftX, ly + FOOT_ROW_H).lineTo(leftX + leftW, ly + FOOT_ROW_H)
       .lineWidth(0.25).strokeColor('#CCC').stroke();
    ly += FOOT_ROW_H;
  }
  leftRow('TOTAL ARRIVALS',  fmtQty(stats.arrivals_qty),  String(stats.arrivals_bags), String(stats.arrivals_lots), 0);
  leftRow('WITHDRAWN',       fmtQty(stats.withdrawn_qty), String(stats.withdrawn_bags), String(stats.withdrawn_lots), 1);
  leftRow('SOLD',            fmtQty(stats.sold_qty),      String(stats.sold_bags),     String(stats.sold_lots), 2);
  leftRow(modeNotXyzLabel(auction), fmtQty(stats.not_qty), String(stats.not_bags), String(stats.not_lots), 3);

  // Vertical column separators inside the left table
  for (let ci = 0; ci < lcW.length - 1; ci++) {
    const vx = lcX[ci] + lcW[ci];
    doc.moveTo(vx, y).lineTo(vx, ly).lineWidth(0.3).strokeColor('#888').stroke();
  }
  doc.rect(leftX, y, leftW, ly - y).lineWidth(0.5).strokeColor('#444').stroke();

  // Right PRICE STATS table: METRIC (50%) | RS (50%)
  const rcW = [Math.floor(rightW * 0.5), 0];
  rcW[1] = rightW - rcW[0];
  const rcX = [rightX, rightX + rcW[0]];

  doc.rect(rightX, y, rightW, FOOT_HEAD_H).fillAndStroke('#E8E4DD', '#444');
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
  doc.text('PRICE',  rcX[0] + 6, y + 5, { width: rcW[0] - 12, align: 'left',  lineBreak: false });
  doc.text('Rs.',    rcX[1] + 6, y + 5, { width: rcW[1] - 12, align: 'right', lineBreak: false });
  let ry2 = y + FOOT_HEAD_H;

  function priceRow(label, val, idx) {
    if (idx % 2 === 1) doc.rect(rightX, ry2, rightW, FOOT_ROW_H).fill('#F7F5F2');
    doc.fillColor('#000').font('Helvetica').fontSize(9);
    doc.text(label, rcX[0] + 6, ry2 + 4, { width: rcW[0] - 12, align: 'left',  lineBreak: false });
    doc.text(val,   rcX[1] + 6, ry2 + 4, { width: rcW[1] - 12, align: 'right', lineBreak: false });
    doc.moveTo(rightX, ry2 + FOOT_ROW_H).lineTo(rightX + rightW, ry2 + FOOT_ROW_H)
       .lineWidth(0.25).strokeColor('#CCC').stroke();
    ry2 += FOOT_ROW_H;
  }
  priceRow('MAXIMUM', fmtMoney(stats.max_price), 0);
  priceRow('MINIMUM', fmtMoney(stats.min_price), 1);
  priceRow('AVERAGE', fmtMoney(stats.avg_price), 2);
  // Pad with empty row so right table aligns with left's 4 rows
  if (ry2 < ly) {
    doc.rect(rightX, ry2, rightW, ly - ry2).fill('#F7F5F2');
    ry2 = ly;
  }

  // Verticals + outer for right table
  for (let ci = 0; ci < rcW.length - 1; ci++) {
    const vx = rcX[ci] + rcW[ci];
    doc.moveTo(vx, y).lineTo(vx, ry2).lineWidth(0.3).strokeColor('#888').stroke();
  }
  doc.rect(rightX, y, rightW, ry2 - y).lineWidth(0.5).strokeColor('#444').stroke();

  y = Math.max(ly, ry2) + 8;

  // Cost of cardamom — full-width prominent strip
  doc.rect(m, y, usableW, FOOT_COST_H).fillAndStroke('#FFF3CD', '#9A6700');
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(11);
  doc.text('COST OF CARDAMOM', m + 8, y + 5, { width: usableW * 0.55, align: 'left', lineBreak: false });
  doc.text(`Rs. ${fmtMoney(stats.cost)}`, m + usableW * 0.55, y + 5, {
    width: usableW * 0.45 - 8, align: 'right', lineBreak: false,
  });
  y += FOOT_COST_H;

  return new Promise(resolve => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

module.exports = {
  // Lot slip — PDF only (XLSX stays in exports.js)
  lotSlipPdf,
  // Generalized carbon-copy slip (Lot Buyer / Lot Name share this layout)
  carbonSlipPdf,
  // Collection — both formats
  collectionXlsx,
  collectionPdf,
  // Trade report — both formats
  tradeReportXlsx,
  tradeReportPdf,
};
