/**
 * reports.js — Reports Hub data aggregation + summary PDF.
 *
 * Three pieces consumed by the Reports tab in the Admin Console:
 *   getTradeSummary(db, auctionId, branchFilter)
 *     → totals, branch-level aggregates, breakdown tables, hourly pace,
 *       per-user (when settings.show_username is on).
 *   getBranchComparison(db)
 *     → lifetime per-branch rollup + per-trade per-branch rows, for the
 *       chart-first comparison view.
 *   generateTradeSummaryPDF(db, auctionId, branchFilter)
 *     → single multi-section A4 PDF mirroring the on-screen card so
 *       dad can hand a stapled copy to the auctioneer.
 *
 * All three respect the active business mode (e-Trade vs e-Auction) by
 * filtering on the parent auction's mode column when present. Lots
 * without an auction get included regardless — the mode filter shouldn't
 * orphan rows in installations that pre-date the mode column.
 */

const PDFDocument = require('pdfkit');
const {
  fmtMoney, fmtQty, fmtPrice,
  getCompanyHeader, drawCompanyHeader,
} = require('./report-formatters');

// ── Helpers ──────────────────────────────────────────────────────────
function _num(v) { return Number(v) || 0; }
function _esc(s) { return String(s == null ? '' : s); }

// Wrapper around prepared-row aggregation that tolerates empty arrays
// and undefined values — half the columns in `lots` are NULL for legacy
// rows so we guard everywhere.
function _sumKey(rows, key) {
  let s = 0;
  for (const r of rows || []) s += _num(r[key]);
  return s;
}

// True when settings.show_username is on. Cached in module scope but
// re-read every request — settings change rarely enough that we don't
// need a TTL, but the per-call DB read is cheap and avoids stale state
// after the operator toggles it.
function _isUsernameShown(db) {
  try {
    const r = db.get(`SELECT value FROM company_settings WHERE key = 'show_username' AND business_mode = ?`,
      [require('./company-config').getActiveMode(db)]);
    return r && String(r.value || '').toLowerCase() === 'true';
  } catch (_) {
    return false;
  }
}

// Returns the noun for an "auction event" based on company_settings.
// e-Trade mode → "Trade(s)"; e-Auction mode (default) → "Auction(s)".
// Mirrors the client-side termAuction() helper so the PDF speaks the
// same language as the UI that triggered it.
function _termAuction(db, plural) {
  let mode = 'e-Auction';
  try {
    const r = db.get(`SELECT value FROM company_settings WHERE key = 'business_mode' AND business_mode = '*'`);
    if (r && r.value) mode = String(r.value);
  } catch (_) {}
  const isTrade = (mode === 'e-Trade');
  if (isTrade)  return plural ? 'Trades'   : 'Trade';
  return            plural ? 'Auctions' : 'Auction';
}

// ─────────────────────────────────────────────────────────────────────
// Trade Summary — per-auction multi-dimensional rollup
// ─────────────────────────────────────────────────────────────────────
function getTradeSummary(db, auctionId, branchFilter) {
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]);
  if (!auction) {
    const err = new Error('Auction not found');
    err.status = 404;
    throw err;
  }

  // Branch filter is APPLIED to the aggregates row (sold/withdrawn/
  // min/max/avg) but NOT to the breakdown tables (branch-wise, grade,
  // sellers) — those always show every branch so the user can pivot.
  // Frontend matches this contract.
  const _branch = String(branchFilter || '').trim();

  // ── Lots query: pull everything we need in one pass to keep the
  // request cheap. We do all the per-dimension aggregations in JS
  // because (a) the row count is bounded to one auction (typically
  // <2000) and (b) it keeps the SQL readable.
  const lots = db.all(
    `SELECT lot_no, branch, grade, name, bags, qty, price, amount, code,
            created_at, user_id
     FROM lots
     WHERE auction_id = ?`,
    [auctionId]
  );

  // ── Totals (the top strip) ─────────────────────────────────────
  const sellers = new Set();
  const branches = new Set();
  let totalBags = 0, totalQty = 0;
  for (const r of lots) {
    if (r.name)   sellers.add(String(r.name).trim().toUpperCase());
    if (r.branch) branches.add(String(r.branch).trim().toUpperCase());
    totalBags += _num(r.bags);
    totalQty  += _num(r.qty);
  }
  const totals = {
    lot_count:    lots.length,
    total_bags:   totalBags,
    total_qty:    totalQty,
    seller_count: sellers.size,
    branch_count: branches.size,
  };

  // ── Branch-filtered aggregates ─────────────────────────────────
  // "Sold" = lots with a code AND code != 'WD'; "Withdrawn" = code='WD'.
  // Min/max/avg are computed over SOLD lots only (withdrawn have no
  // transacted price and would skew the bounds).
  const filtered = _branch
    ? lots.filter(r => String(r.branch || '').trim().toUpperCase() === _branch.toUpperCase())
    : lots;
  const soldRows = [];
  const wdRows = [];
  for (const r of filtered) {
    const code = String(r.code || '').trim().toUpperCase();
    if (!code) continue;
    if (code === 'WD') wdRows.push(r);
    else               soldRows.push(r);
  }
  let pMin = Infinity, pMax = -Infinity, pAmt = 0, pQty = 0;
  for (const r of soldRows) {
    const price = _num(r.price);
    const amt   = _num(r.amount);
    const q     = _num(r.qty);
    if (price > 0 && amt > 0) {
      if (price < pMin) pMin = price;
      if (price > pMax) pMax = price;
    }
    pAmt += amt;
    pQty += q;
  }
  const branchAggregates = {
    branch: _branch || '',
    sold:      { n: soldRows.length, qty: _sumKey(soldRows, 'qty') },
    withdrawn: { n: wdRows.length,   qty: _sumKey(wdRows,   'qty') },
    min: pMin === Infinity   ? 0 : pMin,
    max: pMax === -Infinity  ? 0 : pMax,
    avg: pQty > 0 ? (pAmt / pQty) : 0,
  };

  // ── Per-branch breakdown ───────────────────────────────────────
  // Each branch gets sold/withdrawn breakdowns + min/max/avg over its
  // SOLD lots. Branch == NULL or '' is bucketed as "(unspecified)".
  const byBranch = new Map();
  for (const r of lots) {
    const key = (r.branch || '(unspecified)').toUpperCase();
    if (!byBranch.has(key)) byBranch.set(key, {
      branch: r.branch || '(unspecified)',
      lots: [], sellers: new Set(),
    });
    const g = byBranch.get(key);
    g.lots.push(r);
    if (r.name) g.sellers.add(String(r.name).trim().toUpperCase());
  }
  const branchWise = [];
  for (const g of byBranch.values()) {
    const sold = [], wd = [];
    let bMin = Infinity, bMax = -Infinity, bAmt = 0, bQty = 0;
    let bBags = 0;
    for (const r of g.lots) {
      bBags += _num(r.bags);
      const code = String(r.code || '').trim().toUpperCase();
      if (!code) continue;
      if (code === 'WD') { wd.push(r); continue; }
      sold.push(r);
      const price = _num(r.price);
      const amt   = _num(r.amount);
      if (price > 0 && amt > 0) {
        if (price < bMin) bMin = price;
        if (price > bMax) bMax = price;
      }
      bAmt += amt;
      bQty += _num(r.qty);
    }
    branchWise.push({
      branch:         g.branch,
      lot_count:      g.lots.length,
      total_bags:     bBags,
      total_qty:      _sumKey(g.lots, 'qty'),
      sold_qty:       _sumKey(sold, 'qty'),
      withdrawn_qty:  _sumKey(wd,   'qty'),
      min_price:      bMin === Infinity   ? 0 : bMin,
      max_price:      bMax === -Infinity  ? 0 : bMax,
      avg_price:      bQty > 0 ? bAmt / bQty : 0,
      seller_count:   g.sellers.size,
    });
  }
  branchWise.sort((a, b) => b.total_qty - a.total_qty);

  // ── Per-grade breakdown ────────────────────────────────────────
  const byGrade = new Map();
  for (const r of lots) {
    const key = (r.grade || '(none)').toUpperCase();
    if (!byGrade.has(key)) byGrade.set(key, {
      grade: r.grade || '(none)',
      lots: [], sellers: new Set(),
    });
    const g = byGrade.get(key);
    g.lots.push(r);
    if (r.name) g.sellers.add(String(r.name).trim().toUpperCase());
  }
  const gradeWise = [];
  for (const g of byGrade.values()) {
    gradeWise.push({
      grade:        g.grade,
      lot_count:    g.lots.length,
      total_bags:   _sumKey(g.lots, 'bags'),
      total_qty:    _sumKey(g.lots, 'qty'),
      seller_count: g.sellers.size,
    });
  }
  gradeWise.sort((a, b) => b.total_qty - a.total_qty);

  // ── Top sellers ────────────────────────────────────────────────
  // One row per seller across the whole auction. The frontend caps the
  // display to "Top N" via a dropdown (5 / 10 / 25 / 50 / 100 / All),
  // so we return everything and let the UI slice.
  const bySeller = new Map();
  for (const r of lots) {
    const key = (r.name || '(unknown)').toUpperCase();
    if (!bySeller.has(key)) bySeller.set(key, {
      seller_name: r.name || '(unknown)',
      branch: r.branch || '',
      lots: [],
    });
    bySeller.get(key).lots.push(r);
  }
  const sellerWise = [];
  for (const s of bySeller.values()) {
    sellerWise.push({
      seller_name: s.seller_name,
      branch:      s.branch,
      lot_count:   s.lots.length,
      total_bags:  _sumKey(s.lots, 'bags'),
      total_qty:   _sumKey(s.lots, 'qty'),
    });
  }
  sellerWise.sort((a, b) => b.total_qty - a.total_qty);

  // ── Hourly pace ────────────────────────────────────────────────
  // Buckets by HH of created_at. Useful for spotting "slow morning,
  // busy afternoon" patterns. Lots with no created_at (rare — set by
  // SQLite default) get dropped.
  const byHour = new Map();
  for (const r of lots) {
    if (!r.created_at) continue;
    const m = String(r.created_at).match(/T?(\d{2}):/);
    if (!m) continue;
    const h = m[1];
    if (!byHour.has(h)) byHour.set(h, { hour: h, lots: [] });
    byHour.get(h).lots.push(r);
  }
  const hourly = [];
  for (const g of byHour.values()) {
    hourly.push({
      hour:      g.hour,
      lot_count: g.lots.length,
      total_qty: _sumKey(g.lots, 'qty'),
    });
  }
  hourly.sort((a, b) => a.hour.localeCompare(b.hour));

  // ── Per-user (operator) breakdown ──────────────────────────────
  // Only computed when settings.show_username is on — hidden by default
  // because most installs don't tag lots with a user_id and the column
  // would render empty.
  const showUsername = _isUsernameShown(db);
  let userWise = [];
  if (showUsername) {
    const byUser = new Map();
    for (const r of lots) {
      const u = r.user_id || '';
      if (!u) continue;
      if (!byUser.has(u)) byUser.set(u, { user_id: u, lots: [] });
      byUser.get(u).lots.push(r);
    }
    for (const g of byUser.values()) {
      userWise.push({
        user_id:    g.user_id,
        lot_count:  g.lots.length,
        total_bags: _sumKey(g.lots, 'bags'),
        total_qty:  _sumKey(g.lots, 'qty'),
      });
    }
    userWise.sort((a, b) => b.total_qty - a.total_qty);
  }

  return {
    auction: {
      id:        auction.id,
      ano:       auction.ano,
      date:      auction.date,
      crop_type: auction.crop_type,
      state:     auction.state,
    },
    totals,
    branchAggregates,
    branchWise,
    gradeWise,
    sellerWise,
    hourly,
    userWise,
    showUsername,
  };
}

// ─────────────────────────────────────────────────────────────────────
// Branch Comparison — lifetime rollup + per-trade per-branch rows
// ─────────────────────────────────────────────────────────────────────
function getBranchComparison(db) {
  // `data` powers the per-trade stacked chart (recent 50 trades, every
  // branch as a stack segment). `overall` powers the lifetime totals
  // table + bar chart.
  // The frontend already filters to the recent 50 trades client-side
  // for the chart; we return everything so the data table can show
  // history.
  // Mode-aware via the auctions.mode column.
  let modeFilter = '';
  try {
    const m = db.get(`SELECT value FROM company_settings WHERE key = 'business_mode' AND business_mode = '*'`);
    const mode = m && m.value ? String(m.value).trim() : '';
    if (mode) modeFilter = ` AND (a.mode = '${mode.replace(/'/g, "''")}' OR a.mode IS NULL OR a.mode = '')`;
  } catch (_) { /* no mode filter when settings missing */ }

  // Per-trade per-branch rows. Empty-branch lots get bucketed as
  // "(unspecified)" so totals reconcile.
  const data = db.all(
    `SELECT a.id   AS auction_id,
            a.ano  AS ano,
            a.date AS date,
            a.crop_type,
            COALESCE(NULLIF(TRIM(l.branch), ''), '(unspecified)') AS branch,
            COUNT(l.id)                  AS lot_count,
            COALESCE(SUM(l.bags), 0)     AS total_bags,
            COALESCE(SUM(l.qty),  0)     AS total_qty
     FROM auctions a
     LEFT JOIN lots l ON l.auction_id = a.id
     WHERE l.id IS NOT NULL ${modeFilter}
     GROUP BY a.id, a.ano, a.date, a.crop_type, branch
     ORDER BY a.date DESC, a.id DESC, branch`
  );

  // Lifetime per-branch rollup. trade_count counts distinct auctions
  // that branch participated in.
  const overall = db.all(
    `SELECT COALESCE(NULLIF(TRIM(l.branch), ''), '(unspecified)') AS branch,
            COUNT(DISTINCT l.auction_id) AS trade_count,
            COUNT(l.id)                  AS lot_count,
            COALESCE(SUM(l.bags), 0)     AS total_bags,
            COALESCE(SUM(l.qty),  0)     AS total_qty,
            COUNT(DISTINCT UPPER(TRIM(l.name))) AS seller_count
     FROM lots l
     JOIN auctions a ON a.id = l.auction_id
     WHERE 1=1 ${modeFilter}
     GROUP BY branch
     ORDER BY total_qty DESC`
  );

  return { overall, data };
}

// ─────────────────────────────────────────────────────────────────────
// Trade Summary PDF
// ─────────────────────────────────────────────────────────────────────
// One multi-section A4 portrait page (or two if the seller list is
// long) mirroring the on-screen Trade Summary card. Uses the same
// PDFKit + drawCompanyHeader pattern as every other PDF in the app.
async function generateTradeSummaryPDF(db, auctionId, branchFilter) {
  const summary = getTradeSummary(db, auctionId, branchFilter);
  const a = summary.auction;
  const companyHeader = getCompanyHeader(db);

  const doc = new PDFDocument({ size: 'A4', layout: 'portrait', margin: 28 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));

  const m  = 28;
  const pageW = doc.page.width;
  const pageH = doc.page.height;
  const usableW = pageW - m * 2;
  let y;

  // ── Brand band ──────────────────────────────────────────────────
  const fmtDate = (iso) => {
    const s = String(iso || '').slice(0, 10);
    const mm = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    return mm ? `${mm[3]}/${mm[2]}/${mm[1]}` : s;
  };
  // Noun ("Trade" vs "Auction") follows business_mode so the PDF
  // matches the on-screen card the operator just clicked through.
  const aucWord = _termAuction(db, false);
  const metaLines = [
    `${aucWord} #${a.ano}`,
    `Date: ${fmtDate(a.date)}`,
  ];
  if (a.crop_type) metaLines.push(`Crop: ${a.crop_type}`);
  if (branchFilter) metaLines.push(`Branch: ${branchFilter}`);

  y = drawCompanyHeader(doc, companyHeader || {}, {
    x: m, y: m, width: usableW,
    title: `${aucWord} Summary`,
    metaLines,
  });

  // ── Totals strip ────────────────────────────────────────────────
  const t = summary.totals;
  const tiles = [
    { label: 'Lots',       value: String(t.lot_count) },
    { label: 'Bags',       value: String(t.total_bags) },
    { label: 'Qty (kg)',   value: fmtQty(t.total_qty) },
    { label: 'Sellers',    value: String(t.seller_count) },
    { label: 'Branches',   value: String(t.branch_count) },
  ];
  const tileW = (usableW - (tiles.length - 1) * 6) / tiles.length;
  const tileH = 38;
  for (let i = 0; i < tiles.length; i++) {
    const x = m + i * (tileW + 6);
    doc.roundedRect(x, y, tileW, tileH, 4).fillAndStroke('#F0FDF4', '#86EFAC');
    doc.fillColor('#166534').font('Helvetica-Bold').fontSize(13)
       .text(tiles[i].value, x + 4, y + 5, { width: tileW - 8, align: 'center', lineBreak: false });
    doc.fillColor('#4B5563').font('Helvetica').fontSize(8)
       .text(tiles[i].label, x + 4, y + 22, { width: tileW - 8, align: 'center', lineBreak: false });
  }
  y += tileH + 10;

  // ── Branch-filtered aggregates ──────────────────────────────────
  const ba = summary.branchAggregates;
  const aggCells = [
    { label: 'Sold (kg)',      value: fmtQty(ba.sold.qty) },
    { label: 'Withdrawn (kg)', value: fmtQty(ba.withdrawn.qty) },
    { label: 'Min Price',      value: fmtMoney(ba.min) },
    { label: 'Max Price',      value: fmtMoney(ba.max) },
    { label: 'Avg Price',      value: fmtMoney(ba.avg) },
  ];
  doc.fillColor('#475569').font('Helvetica-Bold').fontSize(9)
     .text(ba.branch ? `Branch: ${ba.branch}` : 'All Branches', m, y, { width: usableW, lineBreak: false });
  y += 12;
  const aggW = (usableW - (aggCells.length - 1) * 6) / aggCells.length;
  for (let i = 0; i < aggCells.length; i++) {
    const x = m + i * (aggW + 6);
    doc.roundedRect(x, y, aggW, tileH, 4).fillAndStroke('#FAFAFA', '#D1D5DB');
    doc.fillColor('#0F172A').font('Helvetica-Bold').fontSize(11)
       .text(aggCells[i].value, x + 4, y + 6, { width: aggW - 8, align: 'center', lineBreak: false });
    doc.fillColor('#6B7280').font('Helvetica').fontSize(8)
       .text(aggCells[i].label, x + 4, y + 23, { width: aggW - 8, align: 'center', lineBreak: false });
  }
  y += tileH + 14;

  // ── Section header helper ───────────────────────────────────────
  function drawSectionHead(text) {
    doc.rect(m, y, usableW, 18).fillAndStroke('#E0F2FE', '#0369A1');
    doc.fillColor('#0C4A6E').font('Helvetica-Bold').fontSize(10)
       .text(text, m + 8, y + 4, { width: usableW - 16, lineBreak: false });
    y += 20;
  }

  // ── Table helper ────────────────────────────────────────────────
  // Each table is a header row + body rows. Caller passes column
  // widths (sums to usableW). Handles page break when out of room.
  function drawTable(headers, widths, rows, opts) {
    const rowH = (opts && opts.rowH) || 14;
    const headH = 16;
    // Header
    doc.rect(m, y, usableW, headH).fillAndStroke('#F3F4F6', '#9CA3AF');
    doc.fillColor('#111827').font('Helvetica-Bold').fontSize(8.5);
    let cx = m;
    for (let i = 0; i < headers.length; i++) {
      doc.text(headers[i], cx + 4, y + 4, {
        width: widths[i] - 8,
        align: opts && opts.aligns ? opts.aligns[i] : 'left',
        lineBreak: false,
      });
      cx += widths[i];
    }
    y += headH;

    // Rows
    doc.font('Helvetica').fontSize(8.5);
    for (let r = 0; r < rows.length; r++) {
      if (y + rowH > pageH - m - 30) {
        doc.addPage();
        y = m;
      }
      // Inset the fill's top edge by the rule width so it doesn't paint over
      // the previous row's separator (drawn at this row's top y).
      if (r % 2 === 1) doc.rect(m, y + 0.5, usableW, rowH - 0.5).fill('#F9FAFB');
      doc.fillColor('#111827');
      let dx = m;
      for (let i = 0; i < rows[r].length; i++) {
        doc.text(String(rows[r][i] == null ? '' : rows[r][i]),
          dx + 4, y + 3, {
            width: widths[i] - 8,
            align: opts && opts.aligns ? opts.aligns[i] : 'left',
            lineBreak: false,
          });
        dx += widths[i];
      }
      // Horizontal rule under every row so the striped table reads as fully ruled.
      doc.moveTo(m, y + rowH).lineTo(m + usableW, y + rowH)
         .lineWidth(0.4).strokeColor('#D1D5DB').stroke();
      y += rowH;
    }
    y += 6;
  }

  // ── Per-Branch table ────────────────────────────────────────────
  drawSectionHead('Per-Branch');
  const bwWidths = [110, 50, 70, 80, 60, 60, 60, usableW - 110 - 50 - 70 - 80 - 60 - 60 - 60];
  const bwRows = summary.branchWise.map(b => [
    b.branch,
    String(b.lot_count),
    fmtQty(b.sold_qty),
    fmtQty(b.withdrawn_qty),
    fmtMoney(b.min_price),
    fmtMoney(b.max_price),
    fmtMoney(b.avg_price),
    String(b.seller_count || 0),
  ]);
  drawTable(
    ['Branch', 'Lots', 'Sold (kg)', 'Withdrawn (kg)', 'Min', 'Max', 'Avg', 'Sellers'],
    bwWidths,
    bwRows.length ? bwRows : [['(no lots yet)', '', '', '', '', '', '', '']],
    { aligns: ['left', 'right', 'right', 'right', 'right', 'right', 'right', 'right'] }
  );

  // ── By Grade table ──────────────────────────────────────────────
  if (summary.gradeWise.length) {
    drawSectionHead('By Grade');
    const gwWidths = [120, 60, 70, 90, usableW - 120 - 60 - 70 - 90];
    const gwRows = summary.gradeWise.map(g => [
      g.grade,
      String(g.lot_count),
      String(g.total_bags || 0),
      fmtQty(g.total_qty),
      String(g.seller_count || 0),
    ]);
    drawTable(
      ['Grade', 'Lots', 'Bags', 'Qty (kg)', 'Sellers'],
      gwWidths, gwRows,
      { aligns: ['left', 'right', 'right', 'right', 'right'] }
    );
  }

  // ── Top Sellers (cap at 50 in PDF — beyond that the table runs
  // ── off the page and the data is better consumed in XLSX). The
  // ── on-screen view lets the user pick higher limits.
  if (summary.sellerWise.length) {
    const cap = Math.min(50, summary.sellerWise.length);
    drawSectionHead(`Top Sellers (showing ${cap} of ${summary.sellerWise.length})`);
    const swWidths = [30, 200, 80, 60, 60, usableW - 30 - 200 - 80 - 60 - 60];
    const swRows = summary.sellerWise.slice(0, cap).map((s, i) => [
      String(i + 1),
      s.seller_name,
      s.branch || '',
      String(s.lot_count),
      String(s.total_bags || 0),
      fmtQty(s.total_qty),
    ]);
    drawTable(
      ['#', 'Seller', 'Branch', 'Lots', 'Bags', 'Qty (kg)'],
      swWidths, swRows,
      { aligns: ['right', 'left', 'left', 'right', 'right', 'right'] }
    );
  }

  // ── Hourly pace (one row per hour) ──────────────────────────────
  if (summary.hourly.length) {
    drawSectionHead('Hourly Pace');
    const hWidths = [80, 90, usableW - 80 - 90];
    const hRows = summary.hourly.map(h => [
      `${h.hour}:00`,
      String(h.lot_count),
      fmtQty(h.total_qty),
    ]);
    drawTable(
      ['Hour', 'Lots', 'Qty (kg)'],
      hWidths, hRows,
      { aligns: ['left', 'right', 'right'] }
    );
  }

  // ── Per-user (only when settings.show_username is on) ───────────
  if (summary.showUsername && summary.userWise.length) {
    drawSectionHead('By Operator');
    const uWidths = [180, 80, 80, usableW - 180 - 80 - 80];
    const uRows = summary.userWise.map(u => [
      u.user_id || '',
      String(u.lot_count),
      String(u.total_bags || 0),
      fmtQty(u.total_qty),
    ]);
    drawTable(
      ['Operator', 'Lots', 'Bags', 'Qty (kg)'],
      uWidths, uRows,
      { aligns: ['left', 'right', 'right', 'right'] }
    );
  }

  doc.end();
  return new Promise((resolve, reject) => {
    doc.on('end',   () => resolve(Buffer.concat(buffers)));
    doc.on('error', reject);
  });
}

module.exports = {
  getTradeSummary,
  getBranchComparison,
  generateTradeSummaryPDF,
};
