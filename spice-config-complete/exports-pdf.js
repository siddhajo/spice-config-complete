/**
 * exports-pdf.js — PDF versions of all XLSX exports
 * Column structures match exports.js exactly, rendered as landscape A4 tables
 *
 * Three exports use specialized renderers (in auction-reports.js) instead of
 * the generic table renderer because their layouts don't fit a flat grid:
 *   - lot_slip      → carbon-copy slip with empty PRICE column
 *   - collection    → invoice register grouped by buyer state
 *   - trade_report  → BUYERS LIST FOR VERIFICATION with state subtotals
 *
 * Full File is too wide (27 columns) to render usably in PDF — that one is
 * XLSX-only. The dispatcher returns null for full_file PDFs; callers should
 * not request that combination, but if they do, the server returns an error.
 */

const PDFDocument = require('pdfkit');
const auctionReports = require('./auction-reports');
const {
  fmtMoney, fmtQty, fmtPrice,
  getCompanyHeader, drawCompanyHeader,
} = require('./report-formatters');

// Manually truncate `text` to fit `maxWidth` using doc.widthOfString. PDFKit
// 0.15's `lineBreak: false` + `ellipsis: true` is unreliable for long single
// tokens — we ellipsize ourselves so multi-word names don't wrap into next row.
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

// Wrap `text` into one or more lines fitting maxWidth. Breaks on word
// boundaries; falls back to character-level break for tokens wider than the
// column. Returns at least one line. Caller must set font/size on doc first.
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

// Preprocess rows for the table renderer.
// Options:
//   serialKey:    column key under which the per-row serial number should be
//                 stored. If set, every data row gets a sequential number.
//   groupByKey:   if set, rows are grouped by this field (preserving the
//                 caller's existing order *within* each group — caller is
//                 responsible for sorting beforehand if a particular group
//                 order is desired). Each group gets:
//                   - serials restart from 1 within the group
//                   - a subtotal row (sumKeys, marked _isSubtotal) appended
//                     after the last row of the group
//   subtotalKeys: list of numeric keys to sum into the subtotal row
//   subtotalLabelKey: key under which to put the subtotal label
//                     (e.g. "ABUDUL BASITH SUBTOTAL")
function preprocessRows(rows, opts) {
  const { serialKey, groupByKey, subtotalKeys = [], subtotalLabelKey } = opts || {};
  if (!groupByKey) {
    // Simple serial numbering, no grouping.
    if (!serialKey) return rows.slice();
    return rows.map((r, i) => ({ ...r, [serialKey]: String(i + 1) }));
  }
  // Group while preserving caller's order. Caller should sort by groupByKey
  // first if they want all rows of one name to appear together.
  // Serial numbering: each unique group gets ONE serial number, placed on
  // the subtotal row. Individual lot rows have a blank serial column.
  const out = [];
  let curKey = null;
  let curGroup = [];
  let groupNo = 0;
  function flushSubtotal() {
    if (!curGroup.length) return;
    const sub = { _isSubtotal: true };
    subtotalKeys.forEach(k => {
      sub[k] = curGroup.reduce((s, r) => s + (Number(r[k]) || 0), 0);
    });
    if (subtotalLabelKey) {
      sub[subtotalLabelKey] = `${curKey || ''} TOTAL`;
    }
    if (serialKey) sub[serialKey] = String(groupNo);
    out.push(sub);
  }
  rows.forEach((r) => {
    const k = r[groupByKey] || '';
    if (k !== curKey) {
      if (curKey !== null) flushSubtotal();
      curKey = k;
      curGroup = [];
      groupNo += 1;
    }
    curGroup.push(r);
    const stamped = { ...r };
    if (serialKey) stamped[serialKey] = '';   // blank on per-row entries
    out.push(stamped);
  });
  if (curKey !== null) flushSubtotal();
  return out;
}

// ── Generic table-to-PDF renderer ───────────────────────────
function renderTablePdf({ title, subtitle, columns, rows, totals, layout, companyHeader }) {
  // layout: 'portrait' (default) or 'landscape'. All exports default to
  // portrait per user preference. Override per-type via PDF_LAYOUT below
  // if a specific report ever needs landscape (e.g. very wide column sets).
  const pageLayout = layout === 'landscape' ? 'landscape' : 'portrait';
  const doc = new PDFDocument({ size: 'A4', layout: pageLayout, margin: 24 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));

  const pageW = doc.page.width;
  const pageH = doc.page.height;
  const m = 24;
  const usableW = pageW - m * 2;

  // Column widths proportional to exports.js width values.
  // Scale so the sum exactly matches usableW; enforce a small min and rescale.
  const totalWeight = columns.reduce((s, c) => s + (c.width || 12), 0);
  const MIN_COL = 22;
  let colWidths = columns.map(c => (c.width || 12) / totalWeight * usableW);
  // Bump narrow columns to MIN_COL, shrink wider cols proportionally to compensate
  const deficit = colWidths.reduce((s, w) => s + Math.max(0, MIN_COL - w), 0);
  if (deficit > 0) {
    const donatePool = colWidths.reduce((s, w) => s + Math.max(0, w - MIN_COL), 0);
    if (donatePool > 0) {
      colWidths = colWidths.map(w => {
        if (w < MIN_COL) return MIN_COL;
        const share = (w - MIN_COL) / donatePool;
        return w - deficit * share;
      });
    }
  }
  colWidths = colWidths.map(w => Math.max(MIN_COL, Math.floor(w)));
  // Final correction so widths sum exactly to usableW
  const diff = usableW - colWidths.reduce((s, w) => s + w, 0);
  colWidths[colWidths.length - 1] = Math.max(MIN_COL, colWidths[colWidths.length - 1] + diff);

  const colX = [m];
  for (let i = 0; i < colWidths.length - 1; i++) colX.push(colX[i] + colWidths[i]);

  const ROW_H = 13;
  const HEAD_H = 16;
  let y;

  function isNumericCol(col) {
    const h = (col.header || '').toUpperCase();
    return /^(QTY|BAG|BAGS|PRICE|RATE|AMOUNT|PQTY|PRATE|PURAMT|CGST|SGST|IGST|TCS|TOTAL|DISCOUNT|PAYABLE|ADVANCE|BALANCE|LITRE|LOTS|TDS|ASSESS_VALUE|COST|NET|GUNNY|TRANSPORT|INSURANCE|CARDAMOM|CARDAMOM_COST|GUNNY_COST|ROUND|BILAMT|COM)$/.test(h);
  }

  function fmtCell(val, col) {
    if (val === null || val === undefined || val === '') return '';
    if (typeof val === 'number') {
      const h = (col.header || '').toUpperCase();
      // Kilos / pure quantities: 3 decimals, Indian commas (1,100.000)
      if (h === 'QTY' || h === 'PQTY' || h === 'LITRE' || h === 'KILOS' || h === 'QUANTITY') {
        return fmtQty(val);
      }
      // Bag counts and other non-decimal integers stay plain
      if (Number.isInteger(val) && (h === 'BAG' || h === 'BAGS' || h === 'LOTS' || h === 'LOT')) {
        return String(val);
      }
      // Everything else numeric is treated as rupees: 2 decimals, Indian commas.
      return fmtMoney(val);
    }
    return String(val);
  }

  // Track the top-Y of the table on each page so we can draw vertical column
  // separators (only inside data-row regions) and an outer border (around the
  // whole table including the totals strip) once the body section closes.
  // Without verticals, columns with right-aligned numbers next to left-aligned
  // text columns look jammed (e.g. PRICE 2163 right-edge sitting next to CODE
  // RSH left-edge). Without borders the table edges look ragged.
  let pageTableTop = null;

  // Draw verticals through the data-row region only (header + rows), without
  // closing the outer border. Caller must still draw the outer border later.
  function drawDataVerticals() {
    if (pageTableTop === null) return;
    const top = pageTableTop, bottom = y;
    for (let ci = 0; ci < colWidths.length - 1; ci++) {
      const vx = colX[ci] + colWidths[ci];
      doc.moveTo(vx, top).lineTo(vx, bottom).lineWidth(0.3).strokeColor('#888').stroke();
    }
  }

  // Draw verticals + outer border for the whole table on this page. Used at
  // page breaks and at the very end of the report (after the totals strip).
  function closePageBorders(extraBottomY) {
    if (pageTableTop === null) return;
    const top = pageTableTop;
    const bottom = (extraBottomY !== undefined) ? extraBottomY : y;
    drawDataVerticals();
    doc.rect(m, top, usableW, bottom - top).lineWidth(0.5).strokeColor('#444').stroke();
    pageTableTop = null;
  }

  // Track where the data section ends on the current page so verticals stop
  // there but the outer border can extend to include the totals strip.
  let dataBottomY = null;

  function drawHeader(firstPage) {
    if (firstPage) {
      // Three-column brand band: company on left, report title centered,
      // subtitle pieces (trade no, date, etc.) right-aligned. The subtitle
      // string for these reports is e.g. "e-TRADE No: 3 — Date: 15/04/2026",
      // so split it on " — " into separate meta lines.
      const metaLines = [];
      if (subtitle) {
        for (const part of String(subtitle).split(' — ')) {
          if (part.trim()) metaLines.push(part.trim());
        }
      }

      const afterY = drawCompanyHeader(doc, companyHeader || {}, {
        x: m, y: m, width: usableW,
        title: title,
        metaLines: metaLines,
      });
      y = afterY;
    } else {
      doc.font('Helvetica-Bold').fontSize(10).fillColor('#000')
         .text(`${title} (continued)`, m, m, { width: usableW, align: 'left' });
      y = m + 18;
    }

    // Compute header height by wrapping each header label
    const HEAD_LINE_H = 10;
    const HEAD_PAD = 4;
    doc.font('Helvetica-Bold').fontSize(8);
    const headerWrapped = columns.map((c, i) => wrapText(doc, c.header, colWidths[i] - 6));
    const headerLines = Math.max(1, ...headerWrapped.map(ls => ls.length));
    const headH = headerLines * HEAD_LINE_H + HEAD_PAD * 2;

    pageTableTop = y;  // remember where this page's column-strip starts
    doc.rect(m, y, usableW, headH).fillAndStroke('#E8E4DD', '#999');
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(8);
    columns.forEach((c, i) => {
      const lines = headerWrapped[i];
      lines.forEach((line, li) => {
        doc.text(line, colX[i] + 3, y + HEAD_PAD + li * HEAD_LINE_H, {
          width: colWidths[i] - 6,
          align: isNumericCol(c) ? 'right' : 'left',
          lineBreak: false,
        });
      });
    });
    y += headH;
  }

  // For numeric cells: if the rendered string is wider than the column,
  // shrink the font from BASE down to a minimum until it fits. Returns the
  // size to use (the caller restores after). Non-numeric cells still wrap.
  function fitNumericFontSize(text, maxWidth, baseSize, isBold) {
    doc.font(isBold ? 'Helvetica-Bold' : 'Helvetica');
    let size = baseSize;
    doc.fontSize(size);
    while (size > 5 && doc.widthOfString(text) > maxWidth) {
      size -= 0.5;
      doc.fontSize(size);
    }
    return size;
  }

  function drawRow(row, i, rowH, wrapped) {
    if (row._isSubtotal) {
      // Subtotal row — full-width yellow strip styled like the grand total.
      doc.rect(m, y, usableW, rowH).fillAndStroke('#FFF3CD', '#E0B020');
      const BASE = 7.5;
      const LINE_H = 10;
      const PAD_TOP = 3;
      columns.forEach((c, ci) => {
        const lines = wrapped[ci];
        const cellW = colWidths[ci] - 6;
        if (isNumericCol(c) && lines.length === 1) {
          // Numeric: auto-shrink to fit on one line.
          const size = fitNumericFontSize(lines[0], cellW, BASE, true);
          doc.fillColor('#000').font('Helvetica-Bold').fontSize(size);
          doc.text(lines[0], colX[ci] + 3, y + PAD_TOP, {
            width: cellW, align: 'right', lineBreak: false,
          });
        } else {
          doc.fillColor('#000').font('Helvetica-Bold').fontSize(BASE);
          lines.forEach((line, li) => {
            doc.text(line, colX[ci] + 3, y + PAD_TOP + li * LINE_H, {
              width: cellW, align: isNumericCol(c) ? 'right' : 'left', lineBreak: false,
            });
          });
        }
      });
      // Bottom separator so subtotal rows are ruled off like every other row.
      doc.moveTo(m, y + rowH).lineTo(m + usableW, y + rowH).lineWidth(0.5).strokeColor('#999').stroke();
      y += rowH;
      return;
    }
    if (i % 2 === 1) doc.rect(m, y, usableW, rowH).fill('#F7F5F2');
    const BASE = 7.5;
    const LINE_H = 10;
    const PAD_TOP = 3;
    columns.forEach((c, ci) => {
      const lines = wrapped[ci];
      const cellW = colWidths[ci] - 6;
      if (isNumericCol(c) && lines.length === 1) {
        // Numeric: auto-shrink to fit on one line — never wrap a number.
        const size = fitNumericFontSize(lines[0], cellW, BASE, false);
        doc.fillColor('#000').font('Helvetica').fontSize(size);
        doc.text(lines[0], colX[ci] + 3, y + PAD_TOP, {
          width: cellW, align: 'right', lineBreak: false,
        });
      } else {
        doc.fillColor('#000').font('Helvetica').fontSize(BASE);
        lines.forEach((line, li) => {
          doc.text(line, colX[ci] + 3, y + PAD_TOP + li * LINE_H, {
            width: cellW, align: isNumericCol(c) ? 'right' : 'left', lineBreak: false,
          });
        });
      }
    });
    // Horizontal rule under every data row so the table reads as fully ruled.
    doc.moveTo(m, y + rowH).lineTo(m + usableW, y + rowH).lineWidth(0.4).strokeColor('#999').stroke();
    y += rowH;
  }

  // Pre-measure a row's required height by wrapping each cell.
  // Numeric cells are NOT wrapped — they're laid out single-line and the font
  // shrinks if the value overflows, since wrapping a number across lines
  // (e.g. "10,71,225." / "00") looks broken. Non-numeric cells word-wrap.
  function measureRow(row) {
    doc.font('Helvetica').fontSize(7.5);
    const LINE_H = 10;
    const PAD_TOP = 3, PAD_BOT = 3;
    const MIN_ROW = 14;
    const wrapped = columns.map((c, ci) => {
      const cellW = colWidths[ci] - 6;
      const text = fmtCell(row[c.key], c);
      if (isNumericCol(c)) {
        // Single-line; auto-shrink handled at draw time.
        return [String(text)];
      }
      return wrapText(doc, text, cellW);
    });
    const maxLines = Math.max(1, ...wrapped.map(ls => ls.length));
    const rowH = Math.max(MIN_ROW, maxLines * LINE_H + PAD_TOP + PAD_BOT);
    return { rowH, wrapped };
  }

  drawHeader(true);

  rows.forEach((row, i) => {
    const { rowH, wrapped } = measureRow(row);
    if (y + rowH > pageH - m - (totals ? 28 : 12)) {
      closePageBorders();
      doc.addPage();
      drawHeader(false);
    }
    drawRow(row, i, rowH, wrapped);
  });

  if (totals) {
    if (y + 28 > pageH - m) { closePageBorders(); doc.addPage(); drawHeader(false); }
    // Draw verticals through the data-row region only — they must stop before
    // the totals strip so column dividers don't cut through it.
    drawDataVerticals();
    y += 2;
    doc.rect(m, y, usableW, ROW_H + 2).fillAndStroke('#FFF3CD', '#E0B020');
    columns.forEach((c, ci) => {
      const val = totals[c.key];
      if (val === undefined || val === null || val === '') return;
      const cellW = colWidths[ci] - 6;
      const text = fmtCell(val, c);
      if (isNumericCol(c)) {
        // Auto-shrink numeric totals so they never get truncated with an
        // ellipsis — losing digits in a total is much worse than a slightly
        // smaller font.
        const size = fitNumericFontSize(text, cellW, 8, true);
        doc.fillColor('#000').font('Helvetica-Bold').fontSize(size);
        doc.text(text, colX[ci] + 3, y + 4, {
          width: cellW, align: 'right', lineBreak: false,
        });
      } else {
        doc.fillColor('#000').font('Helvetica-Bold').fontSize(8);
        const fitted = fitText(doc, text, cellW);
        doc.text(fitted, colX[ci] + 3, y + 4, {
          width: cellW, align: 'left', lineBreak: false,
        });
      }
    });
    y += ROW_H + 2;
    // Outer border now encloses data + totals; the verticals were drawn
    // already so closePageBorders should not draw them again. Inline the
    // outer border draw and reset pageTableTop.
    if (pageTableTop !== null) {
      doc.rect(m, pageTableTop, usableW, y - pageTableTop).lineWidth(0.5).strokeColor('#444').stroke();
      pageTableTop = null;
    }
  } else {
    // No totals — close verticals + outer border in one go on the final page
    closePageBorders();
  }

  doc.fillColor('#888').font('Helvetica').fontSize(7)
     .text(`Rows: ${rows.length}`, m, pageH - m - 10, { width: usableW, align: 'left' });
  doc.text('Generated by Spice Config', m, pageH - m - 10, { width: usableW, align: 'right' });

  return new Promise(resolve => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

function sumKeys(rows, keys) {
  const out = {};
  keys.forEach(k => { out[k] = rows.reduce((s, r) => s + (Number(r[k]) || 0), 0); });
  return out;
}

// ── Column defs — must match exports.js columns exactly ─────
const COLS = {
  lot_slip: [
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'LOT', key: 'lot', width: 8 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'GRADE', key: 'grade', width: 8 },
    { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'LITRE', key: 'litre', width: 10 },
  ],
  lot_slip_after: [
    { header: 'STATE', key: 'state', width: 12 },
    { header: 'LOT', key: 'lot', width: 8 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'PRICE', key: 'price', width: 10 },
    { header: 'AMOUNT', key: 'amount', width: 14 },
    { header: 'CODE', key: 'code', width: 8 },
  ],
  lot_buyer: [
    { header: 'LOT',   key: 'lot',   width: 8  },
    { header: 'BUYER', key: 'buyer', width: 24 },
    { header: 'BR',    key: 'br',    width: 6  },
    { header: 'BAG',   key: 'bag',   width: 6  },
    { header: 'QTY',   key: 'qty',   width: 12 },
  ],
  lot_name: [
    { header: 'LOT',     key: 'lot',     width: 8  },
    { header: 'NAME',    key: 'name',    width: 30 },
    { header: 'BR',      key: 'br',      width: 6  },
    { header: 'BAG',     key: 'bag',     width: 6  },
    { header: 'QTY',     key: 'qty',     width: 12 },
    { header: 'PRICE',   key: 'price',   width: 10 },
    { header: 'CONTROL', key: 'control', width: 12 },
  ],
  price_list: [
    { header: 'LOT', key: 'lot', width: 8 },
    { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'PRICE', key: 'price', width: 10 },
    { header: 'CODE', key: 'code', width: 8 },
    { header: 'BIDDER', key: 'bidder', width: 20 },
  ],
  price_list_before: [
    { header: 'TNO',   key: 'trade_no', width: 10 },
    { header: 'DATE',  key: 'date',     width: 12 },
    { header: 'LOT',   key: 'lot',      width: 10 },
    { header: 'BAG',   key: 'bag',      width: 8  },
    { header: 'QTY',   key: 'qty',      width: 14 },
    { header: 'PRICE', key: 'price',    width: 10 },
    { header: 'CODE',  key: 'code',     width: 10 },
  ],
  bank_payment: [
    // PDF-only display columns — restructured for portrait so all data fits
    // without wrapping single-token values like "HDFC0001234" or 6-digit
    // PINs. The XLSX export in exports.js still has the full 9-column
    // bank-required schema (TransactionType, BeneIFSCode, BeneAcctNo,
    // BeneName, BeneAddLine1, BeneAddLine2, BeneAddLine3, Amount,
    // SendertoRcvrInfo) — only PDF combines address+city+pin for display.
    { header: 'SL.NO',    key: '_sn',              width: 6  },
    { header: 'TYPE',     key: 'transactionType',  width: 10 },
    { header: 'IFSC',     key: 'ifsc',             width: 18 },
    { header: 'A/C NO',   key: 'accountNo',        width: 22 },
    { header: 'NAME',     key: 'beneficiaryName',  width: 24 },
    { header: 'ADDRESS',  key: 'address_combined', width: 32 },
    { header: 'AMOUNT',   key: 'amount',           width: 16 },
    { header: 'REMARKS',  key: 'remarks',          width: 26 },
  ],
  pooler_register: [
    // STATE column dropped per user request; serial number shown per-name
    // (resets within each name group), with a subtotal row for each name.
    { header: 'SL.NO',  key: '_sn',         width: 6  },
    { header: 'NAME',   key: 'poolername',  width: 28 },
    { header: 'BRANCH', key: 'br',          width: 12 },
    { header: 'LOT',    key: 'lot',         width: 7  },
    { header: 'QTY',    key: 'qty',         width: 12 },
    { header: 'PRICE',  key: 'price',       width: 9  },
    { header: 'AMOUNT', key: 'amount',      width: 16 },
    { header: 'PQTY',   key: 'pqty',        width: 12 },
    { header: 'PRATE',  key: 'prate',       width: 9  },
    { header: 'PURAMT', key: 'puramt',      width: 16 },
  ],
  full_file: [
    { header: 'STATE', key: 'state', width: 10 }, { header: 'LOT', key: 'lot_no', width: 8 },
    { header: 'CROP', key: 'crop', width: 8 }, { header: 'GRADE', key: 'grade', width: 8 },
    { header: 'CRPT', key: 'crpt', width: 8 }, { header: 'BRANCH', key: 'branch', width: 12 },
    { header: 'NAME', key: 'name', width: 24 }, { header: 'CR', key: 'cr', width: 18 },
    { header: 'PAN', key: 'pan', width: 12 }, { header: 'TEL', key: 'tel', width: 12 },
    { header: 'BAG', key: 'bags', width: 6 }, { header: 'QTY', key: 'qty', width: 10 },
    { header: 'PRICE', key: 'price', width: 10 }, { header: 'AMOUNT', key: 'amount', width: 12 },
    { header: 'CODE', key: 'code', width: 8 }, { header: 'BUYER', key: 'buyer', width: 12 },
    { header: 'BUYER1', key: 'buyer1', width: 16 }, { header: 'SALE', key: 'sale', width: 6 },
    { header: 'INVO', key: 'invo', width: 8 }, { header: 'PQTY', key: 'pqty', width: 10 },
    { header: 'PRATE', key: 'prate', width: 10 }, { header: 'PURAMT', key: 'puramt', width: 12 },
    { header: 'COM', key: 'com', width: 8 }, { header: 'CGST', key: 'cgst', width: 8 },
    { header: 'SGST', key: 'sgst', width: 8 }, { header: 'IGST', key: 'igst', width: 8 },
    { header: 'ADVANCE', key: 'advance', width: 10 }, { header: 'BALANCE', key: 'balance', width: 10 },
  ],
  collection: [
    { header: 'BRANCH', key: 'branch', width: 15 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'CR', key: 'cr', width: 25 },
    { header: 'BAG', key: 'bag', width: 6 },
    { header: 'QTY', key: 'qty', width: 12 },
    { header: 'LITRE', key: 'litre', width: 10 },
    { header: 'GRADE', key: 'grade', width: 8 },
  ],
  dealer_list: [
    // STATE column dropped per user request; simple sequential serial number.
    { header: 'SL.NO', key: '_sn',   width: 6  },
    { header: 'NAME',  key: 'name',  width: 30 },
    { header: 'GSTIN', key: 'gstin', width: 18 },
    { header: 'LOTS',  key: 'lots',  width: 6  },
    { header: 'BAGS',  key: 'bags',  width: 6  },
    { header: 'QTY',   key: 'qty',   width: 12 },
  ],
  sales_taxes: [
    { header: 'STATE', key: 'state', width: 10 }, { header: 'SALE', key: 'sale', width: 6 },
    { header: 'INVO', key: 'invo', width: 8 }, { header: 'TRADERNAME', key: 'tradername', width: 22 },
    { header: 'BAG', key: 'bag', width: 6 }, { header: 'QTY', key: 'qty', width: 10 },
    { header: 'CARDAMOM', key: 'cardamom_cost', width: 12 },
    { header: 'GUNNY', key: 'gunny_cost', width: 10 },
    { header: 'CGST', key: 'cgst', width: 10 }, { header: 'SGST', key: 'sgst', width: 10 },
    { header: 'IGST', key: 'igst', width: 10 }, { header: 'TCS', key: 'tcs', width: 8 },
    { header: 'TRANSPORT', key: 'transport', width: 10 },
    { header: 'INSURANCE', key: 'insurance', width: 10 },
    { header: 'TOTAL', key: 'total', width: 12 },
  ],
  payment: [
    // Serial number resets per pooler; each pooler gets a subtotal row.
    // Money columns sized for Indian-format 7-digit values like
    // "1,73,31,966.50" without wrapping.
    { header: 'SL.NO',      key: '_sn',         width: 5  },
    { header: 'POOLERNAME', key: 'poolername',  width: 22 },
    { header: 'LOT',        key: 'lot',         width: 7  },
    { header: 'BAG',        key: 'bag',         width: 5  },
    { header: 'QTY',        key: 'qty',         width: 10 },
    { header: 'PRICE',      key: 'price',       width: 9  },
    { header: 'AMOUNT',     key: 'amount',      width: 14 },
    { header: 'PQTY',       key: 'pqty',        width: 10 },
    { header: 'PRATE',      key: 'prate',       width: 9  },
    { header: 'PURAMT',     key: 'puramt',      width: 14 },
    { header: 'DISCOUNT',   key: 'discount',    width: 10 },
    { header: 'PAYABLE',    key: 'payable',     width: 14 },
  ],
  tally_purchase: [
    { header: 'NAME', key: 'name', width: 24 }, { header: 'ADD', key: 'add', width: 24 },
    { header: 'PLACE', key: 'place', width: 12 }, { header: 'GSTIN', key: 'gstin', width: 16 },
    { header: 'TEL', key: 'tel', width: 12 }, { header: 'LOT', key: 'lot', width: 8 },
    { header: 'BAG', key: 'bag', width: 6 }, { header: 'QTY', key: 'qty', width: 10 },
    { header: 'PRICE', key: 'price', width: 10 }, { header: 'AMOUNT', key: 'amount', width: 12 },
    { header: 'CGST', key: 'cgst', width: 10 }, { header: 'SGST', key: 'sgst', width: 10 },
    { header: 'IGST', key: 'igst', width: 10 }, { header: 'DISCOUNT', key: 'discount', width: 10 },
    { header: 'BILAMT', key: 'bilamt', width: 12 },
  ],
  tds_return: [
    { header: 'INVOICE', key: 'invoice', width: 10 },
    { header: 'DATE', key: 'date', width: 12 },
    { header: 'NAME', key: 'name', width: 30 },
    { header: 'PAN', key: 'pan', width: 12 },
    { header: 'ASSESS_VALUE', key: 'assess_value', width: 14 },
    { header: 'TDS', key: 'tds', width: 12 },
  ],
};

// Column specs for the carbon-copy (two-up) PDF variants of Lot Buyer /
// Lot Name. `weight` is the column's share of the half-page width; the wide
// BUYER / NAME column carries fit:true so long names ellipsize cleanly
// instead of overlapping. kind 'qty'/'price' formats numbers; the blank
// CONTROL column stays empty for hand-written notes on the printed slip.
const CARBON_COLS = {
  lot_buyer: [
    { header: 'LOT',   key: 'lot',   weight: 0.16, align: 'center' },
    { header: 'BUYER', key: 'buyer', weight: 0.40, align: 'left', fit: true },
    { header: 'BR',    key: 'br',    weight: 0.12, align: 'center' },
    { header: 'BAG',   key: 'bag',   weight: 0.13, align: 'right' },
    { header: 'QTY',   key: 'qty',   weight: 0.19, align: 'right', kind: 'qty' },
  ],
  lot_name: [
    { header: 'LOT',     key: 'lot',     weight: 0.10, align: 'center' },
    { header: 'NAME',    key: 'name',    weight: 0.27, align: 'left', fit: true },
    { header: 'BR',      key: 'br',      weight: 0.08, align: 'center' },
    { header: 'BAG',     key: 'bag',     weight: 0.09, align: 'right' },
    { header: 'QTY',     key: 'qty',     weight: 0.16, align: 'right', kind: 'qty' },
    { header: 'PRICE',   key: 'price',   weight: 0.16, align: 'right', kind: 'price' },
    { header: 'CONTROL', key: 'control', weight: 0.14, align: 'center', blank: true },
  ],
};

const TOTAL_KEYS = {
  lot_slip:        ['bag', 'qty'],
  lot_slip_after:  ['bag', 'qty', 'amount'],
  lot_buyer:       ['bag', 'qty'],
  lot_name:        ['bag', 'qty'],
  price_list:      ['bag', 'qty'],
  price_list_before: ['bag', 'qty'],
  bank_payment:    ['amount'],
  pooler_register: ['qty', 'amount', 'pqty', 'puramt'],
  full_file:       ['bags', 'qty', 'amount', 'pqty', 'puramt', 'cgst', 'sgst', 'igst', 'advance', 'balance'],
  collection:      ['bag', 'qty'],
  dealer_list:     ['lots', 'bags', 'qty'],
  sales_taxes:     ['bag', 'qty', 'cardamom_cost', 'gunny_cost', 'cgst', 'sgst', 'igst', 'tcs', 'transport', 'insurance', 'total'],
  payment:         ['bag', 'qty', 'amount', 'pqty', 'puramt', 'discount', 'payable'],
  tally_purchase:  ['bag', 'qty', 'amount', 'cgst', 'sgst', 'igst', 'discount', 'bilamt'],
  tds_return:      ['assess_value', 'tds'],
};

const TITLES = {
  lot_slip:        'Lot Slip',
  lot_slip_after:  'Lot Slip (After Trade)',
  lot_buyer:       'Lot Buyer',
  lot_name:        'Lot Name',
  price_list:      'Price List',
  price_list_before: 'Price List (Before)',
  bank_payment:    'Bank Payment (RTGS/NEFT)',
  pooler_register: 'Pooler Register',
  full_file:       'Full File',
  collection:      'Collection / Lorry',
  dealer_list:     'Dealer List',
  sales_taxes:     'Sales & Taxes',
  payment:         'Payment Summary',
  tally_purchase:  'Tally Purchase',
  tds_return:      'TDS Return',
};

// Per-type page orientation override. Portrait is the default (set in
// renderTablePdf). Add a type here only if a specific report needs landscape
// because its column count is too high to fit comfortably in portrait.
const PDF_LAYOUT = {
  // Full file is a 28-column raw export (STATE, LOT, CROP, GRADE, CRPT,
  // BRANCH, NAME, CR, PAN, TEL, BAG, QTY, PRICE, AMOUNT, CODE, BUYER,
  // BUYER1, SALE, INVO, PQTY, PRATE, PURAMT, COM, CGST, SGST, IGST,
  // ADVANCE, BALANCE). Portrait is impossible — stays landscape.
  full_file: 'landscape',
};

// Per-type row preprocessing: add a serial-number column, optionally group
// rows by a name field with a subtotal row inserted after each group. Keys
// here line up with the COLS definitions (e.g. `_sn` for the SL.NO column,
// `poolername` for group-by). See `preprocessRows` for semantics.
const ROW_PREPROCESS = {
  // Bank payment — flat sequential serial.
  bank_payment: {
    serialKey: '_sn',
  },
  // Pooler register — serial restarts per pooler name; subtotal of qty,
  // amount, pqty, puramt at the end of each pooler's rows.
  pooler_register: {
    serialKey: '_sn',
    groupByKey: 'poolername',
    subtotalKeys: ['qty', 'amount', 'pqty', 'puramt'],
    subtotalLabelKey: 'poolername',
  },
  // Dealer list — flat sequential serial (no grouping).
  dealer_list: {
    serialKey: '_sn',
  },
  // Payment summary — serial restarts per pooler name; subtotal of bag, qty,
  // amount, pqty, puramt, discount, payable at the end of each pooler's rows.
  payment: {
    serialKey: '_sn',
    groupByKey: 'poolername',
    subtotalKeys: ['bag', 'qty', 'amount', 'pqty', 'puramt', 'discount', 'payable'],
    subtotalLabelKey: 'poolername',
  },
};

async function getRowsForType(db, type, auctionId, cfg, extra) {
  switch (type) {
    case 'lot_slip':
      return db.all(
        `SELECT state, lot_no as lot, name, grade, bags as bag, qty, litre
         FROM lots WHERE auction_id = ? ${extra.state ? 'AND state = ?' : ''}
         ORDER BY lot_no`, extra.state ? [auctionId, extra.state] : [auctionId]);

    case 'lot_slip_after':
      return db.all(
        `SELECT state, lot_no as lot, name, bags as bag, qty, price, amount, code
         FROM lots WHERE auction_id = ? ${extra.state ? 'AND state = ?' : ''}
         ORDER BY lot_no`, extra.state ? [auctionId, extra.state] : [auctionId]);

    case 'lot_buyer':
      return db.all(
        `SELECT lot_no as lot, COALESCE(buyer,'') as buyer,
                CASE UPPER(COALESCE(state,''))
                  WHEN 'KERALA' THEN 'KL'
                  WHEN 'TAMIL NADU' THEN 'TN'
                  ELSE UPPER(SUBSTR(COALESCE(state,''), 1, 2))
                END AS br,
                bags as bag, qty
         FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]);

    case 'lot_name':
      return db.all(
        `SELECT lot_no as lot, COALESCE(name,'') as name,
                CASE UPPER(COALESCE(state,''))
                  WHEN 'KERALA' THEN 'KL'
                  WHEN 'TAMIL NADU' THEN 'TN'
                  ELSE UPPER(SUBSTR(COALESCE(state,''), 1, 2))
                END AS br,
                bags as bag, qty, price, '' as control
         FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]);

    case 'price_list':
      return db.all(
        `SELECT lot_no as lot, bags as bag, qty, price, code, buyer as bidder
         FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]);

    case 'price_list_before': {
      const a = db.get('SELECT ano, date FROM auctions WHERE id = ?', [auctionId]) || {};
      const tradeNo = a.ano || '';
      const tradeDate = String(a.date || '').slice(0, 10).split('-').reverse().join('/');
      // PRICE blanked when 0 so the column reads empty instead of "0.00",
      // matching CODE's blank-when-unset behaviour.
      return db.all(
        `SELECT lot_no as lot, bags as bag, qty,
                CASE WHEN COALESCE(price,0) = 0 THEN '' ELSE price END AS price,
                COALESCE(code,'') AS code
         FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]
      ).map(r => ({ trade_no: tradeNo, date: tradeDate, ...r }));
    }

    case 'bank_payment': {
      const { getBankPaymentData } = require('./calculations');
      const rows = getBankPaymentData(db, auctionId, cfg);
      // For PDF display, combine address1 + address2 (city) + pin into a
      // single ADDRESS column. The underlying data still has the original
      // separate fields (used by XLSX export for the bank-required format).
      return rows.map(r => ({
        ...r,
        address_combined: [r.address1, r.address2, r.pin]
          .filter(s => s != null && String(s).trim() !== '').join(', '),
      }));
    }

    case 'pooler_register':
      return db.all(
        `SELECT state, lot_no as lot, name as poolername, branch as br, qty, price, amount, pqty, prate, puramt
         FROM lots WHERE auction_id = ? AND amount > 0 ORDER BY name`, [auctionId]);

    case 'full_file':
      return db.all(`SELECT * FROM lots WHERE auction_id = ? ORDER BY lot_no`, [auctionId]);

    case 'collection':
      return db.all(
        `SELECT branch, name, cr, bags as bag, qty, litre, grade
         FROM lots WHERE auction_id = ? ORDER BY branch, name`, [auctionId]);

    case 'dealer_list':
      return db.all(
        `SELECT state, name, SUBSTR(cr, 7, 15) as gstin,
          COUNT(lot_no) as lots, SUM(bags) as bags, SUM(qty) as qty
         FROM lots WHERE auction_id = ? AND cr LIKE '%GST%' AND amount > 0
         GROUP BY state, name, cr ORDER BY name`, [auctionId]);

    case 'sales_taxes':
      return db.all(
        `SELECT state, sale, invo, buyer1 as tradername, bags as bag, qty,
          amount as cardamom_cost, gunny as gunny_cost,
          cgst, sgst, igst, tcs, pava_hc as transport, ins as insurance, tot as total
         FROM invoices WHERE ano = (SELECT ano FROM auctions WHERE id = ?)
         ORDER BY sale, invo`, [auctionId]);

    case 'payment': {
      // Mode-aware discount column — see exports.js exportPaymentSummary.
      const mode = (cfg && cfg.business_mode || 'e-Trade').toLowerCase();
      const discountCol = (mode === 'auction') ? 'advance' : 'refund';
      return db.all(
        `SELECT name as poolername, lot_no as lot, bags as bag, qty, price, amount,
          pqty, prate, puramt, ${discountCol} as discount, balance as payable
         FROM lots WHERE auction_id = ? AND amount > 0
         ORDER BY state, name`, [auctionId]);
    }

    case 'tally_purchase': {
      const mode = (cfg && cfg.business_mode || 'e-Trade').toLowerCase();
      const discountCol = (mode === 'auction') ? 'advance' : 'refund';
      return db.all(
        `SELECT name, padd as add, ppla as place, cr as gstin, tel,
          lot_no as lot, bags as bag, pqty as qty, prate as price, puramt as amount,
          cgst, sgst, igst, ${discountCol} as discount, puramt as bilamt
         FROM lots WHERE auction_id = ? AND amount > 0
          AND cr NOT LIKE 'GSTIN.%'
         ORDER BY name`, [auctionId]);
    }

    case 'tds_return': {
      const { getTDSReturnData } = require('./calculations');
      return getTDSReturnData(db, extra.from, extra.to, 'invoice');
    }

    default:
      throw new Error(`Unknown export type: ${type}`);
  }
}

async function exportPdf(db, type, auctionId, cfg, extra = {}) {
  // Specialized renderers — these don't use the generic table layout.
  if (type === 'lot_slip') {
    return auctionReports.lotSlipPdf(db, auctionId, cfg, extra);
  }
  if (type === 'collection') {
    return auctionReports.collectionPdf(db, auctionId);
  }
  if (type === 'trade_report') {
    return auctionReports.tradeReportPdf(db, auctionId);
  }
  if (type === 'full_file') {
    throw new Error('Full File is XLSX-only — PDF version is not supported (too many columns to fit on a page).');
  }
  // Lot Buyer / Lot Name → carbon-copy tear-off slip (two identical halves
  // per page), matching the Lot Slip. The wide BUYER / NAME column is
  // ellipsized (fit:true) so long names never overlap the next column.
  if (type === 'lot_buyer' || type === 'lot_name') {
    const rows = await getRowsForType(db, type, auctionId, cfg, extra);
    return auctionReports.carbonSlipPdf(db, auctionId, cfg, extra, {
      title: TITLES[type] || type,
      columns: CARBON_COLS[type],
      totalKeys: TOTAL_KEYS[type] || [],
      rows,
    });
  }

  const columns = COLS[type];
  if (!columns) throw new Error(`No PDF column def for type: ${type}`);

  let rows = await getRowsForType(db, type, auctionId, cfg, extra);

  // Per-type row preprocessing: serial numbers + group-by-name subtotals.
  const ppCfg = ROW_PREPROCESS[type];
  if (ppCfg) {
    rows = preprocessRows(rows, ppCfg);
  }

  const totalKeys = TOTAL_KEYS[type] || [];
  const totals = totalKeys.length && rows.length ? (() => {
    const t = sumKeys(rows.filter(r => !r._isSubtotal), totalKeys);
    // Place "TOTAL" in the first non-serial column so the label is visible.
    // If column[0] is the SL.NO column, the label goes in column[1] instead.
    const labelCol = (columns[0] && columns[0].key === '_sn') ? columns[1] : columns[0];
    if (labelCol) t[labelCol.key] = 'TOTAL';
    return t;
  })() : null;

  let subtitle = '';
  if (type === 'tds_return') {
    subtitle = `Period: ${extra.from || ''} to ${extra.to || ''}`;
  } else if (auctionId) {
    const auction = db.get('SELECT ano, date, crop_type, mode FROM auctions WHERE id = ?', [auctionId]);
    if (auction) {
      const d = auction.date ? auction.date.split('-').reverse().join('/') : '';
      // Two clean meta lines, joined by " — " so renderTablePdf can split
      // them back into separate right-side rows. The crop type (ISP/ASP) is
      // omitted — the active preset is already obvious from the logo and
      // company name in the brand block.
      // Mode-aware header: shows "e-TRADE No:" or "e-AUCTION No:" based
      // on the auction's stamped mode. Empty/legacy mode falls back to
      // e-AUCTION wording.
      const modeLbl = (auction.mode === 'e-Trade') ? 'e-TRADE' : 'e-AUCTION';
      subtitle = `${modeLbl} No: ${auction.ano} — Date: ${d}`;
      if (extra.state) subtitle += ` — State: ${extra.state}`;
    }
  }

  return renderTablePdf({
    title: TITLES[type] || type,
    subtitle,
    columns,
    rows,
    totals,
    layout: PDF_LAYOUT[type],
    companyHeader: getCompanyHeader(db),
  });
}

module.exports = { exportPdf, TITLES, COLS };
