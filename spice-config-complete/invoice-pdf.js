/**
 * invoice-pdf.js — GST Invoice PDF generation
 * Replaces: GSTKBILT.PRG, GSTKBILP.PRG, GSTIN.PRG printer output
 */

const PDFDocument = require('pdfkit');
const { amountToWords } = require('./amount-words');
// Sensitive-field masking (bank account / IFSC) per the company's display
// policy — applied to client-facing invoices so a configured mask is honoured.
const { makeMaskers } = require('./mask-fields');

// Shared flag-reader: cfg flags can be `true|false` (booleans), `'true'|'false'`
// (strings, from JSON storage), or `undefined`/empty (treat as defaultOn).
function readFlag(val, defaultOn) {
  if (val === undefined || val === null || val === '') return defaultOn;
  if (typeof val === 'boolean') return val;
  return String(val).toLowerCase() === 'true';
}

// ── Aligned label/colon/value row ────────────────────────────────
// Draws "Label : value" with the COLON pinned to a fixed x, so every row
// in a block lines up vertically (labels are left-aligned, colons column-
// aligned, values start at the same x). Proportional fonts make padded-
// string alignment ("GSTIN    : x") drift, so we draw the three pieces
// separately. Pass `labelColW` = the widest label width in the block
// (measure with doc.widthOfString under the active font). `blockX` is the
// row's left edge and `blockW` its full width so the value wraps in place.
// Caller controls the y-cursor and font.
function drawKVRow(doc, label, value, blockX, y, labelColW, blockW, opts = {}) {
  const colonGap = opts.colonGap != null ? opts.colonGap : 3; // label → colon
  const valueGap = opts.valueGap != null ? opts.valueGap : 4; // colon → value
  const colonX = blockX + labelColW + colonGap;
  const valueX = colonX + doc.widthOfString(':') + valueGap;
  doc.text(String(label == null ? '' : label), blockX, y, { width: labelColW });
  doc.text(':', colonX, y);
  const valueW = Math.max(10, (blockX + blockW) - valueX);
  doc.text(String(value == null ? '' : value), valueX, y, { width: valueW });
  return valueW;
}
// Widest label width (current font) across a list of label strings.
function widestLabel(doc, labels) {
  let w = 0;
  for (const l of labels) { const x = doc.widthOfString(String(l)); if (x > w) w = x; }
  return w;
}

// ── Invoice number formatter ──────────────────────────────────────
// Format: {inv_prefix}/{saleType}-{invoiceNo}/{season_short}
// Examples:
//   Sales:    "ISP/L-9/26-27"
//   Purchase: "ISP/9/26-27"   (no saleType segment)
// Separators are hardcoded: "/" outer, "-" between saleType & invoiceNo.
function formatInvoiceNo(cfg, saleType, invoiceNo) {
  const prefix = cfg.inv_prefix || '';
  const season = cfg.season_short || '';
  // Middle segment: "L-9" if saleType present, else just "9"
  const middle = saleType ? `${saleType}-${invoiceNo}` : String(invoiceNo);
  const parts  = [prefix, middle, season].filter(p => p !== '' && p != null);
  return parts.join('/');
}

// ── Indian number formatter (lakhs style) ──────────────────────
// Produces strings like "4,25,356.80" instead of "425,356.80"
function formatINR(n, decimals = 2) {
  const num = Number(n || 0);
  const sign = num < 0 ? '-' : '';
  const abs  = Math.abs(num);
  const parts = abs.toFixed(decimals).split('.');
  let intPart = parts[0];
  const dec = parts[1] || '';
  // Indian grouping: last 3, then pairs
  let formatted;
  if (intPart.length <= 3) {
    formatted = intPart;
  } else {
    const last3 = intPart.slice(-3);
    const rest  = intPart.slice(0, -3);
    formatted = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + last3;
  }
  return sign + formatted + (dec ? '.' + dec : '');
}

// ── Effective company details based on business_state ────────────
// ── Effective company details based on state ────────────────────
// KL (e-Trade or e-Auction) → use sister.* (ASP)
// TN (e-Trade or e-Auction) → use company.* / address_tn.* (ISP)
//
// Kerala resolves to the Sister Company (ASP) in BOTH modes — so the
// Sister Company section must be configured for Kerala installs.
function effectiveCompany(cfg) {
  const state = (cfg.business_state || '').toUpperCase();
  const useASP = (state === 'KERALA');

  // Partnership Firm toggle → the letterhead identity line shows
  // "Partnership: <name>" instead of "CIN: <cin>". cinLabel/cinValue
  // carry the resolved label + value so the renderer doesn't repeat
  // this branch; `cin` stays the raw CIN for any other consumer.
  const _bool = (v) => v === true || String(v || '').toLowerCase() === 'true';
  if (useASP) {
    const isPartner = _bool(cfg.s_is_partnership);
    return {
      logo:    cfg.s_logo     || 'ASP',
      name:    cfg.s_company  || cfg.s_short_name || '',
      short:   cfg.s_short_name || cfg.s_company  || '',
      pan:     cfg.s_pan      || '',
      cin:     cfg.s_cin      || '',
      isPartnership: isPartner,
      partnershipName: cfg.s_partnership_name || '',
      cinLabel: isPartner ? 'Partnership' : 'CIN',
      cinValue: isPartner ? (cfg.s_partnership_name || '') : (cfg.s_cin || ''),
      fssai:   cfg.s_fssai    || '',
      sbl:     cfg.s_sbl      || '',
      msme:    cfg.s_msme     || '',
      address1: cfg.s_address1 || '',
      address2: cfg.s_address2 || '',
      place:   cfg.s_place    || '',
      pin:     cfg.s_pin      || '',
      stateName: cfg.s_state  || 'KERALA',
      stateCode: cfg.s_st_code || '32',
      phone:   cfg.s_phone    || '',
      email:   cfg.s_email    || '',
      gstin:   cfg.s_gstin    || '',
    };
  }
  // ISP: company.* + address matching state (TN uses address_tn, KL uses address_kl)
  const isStateKL = (state === 'KERALA');
  const ispPartner = _bool(cfg.is_partnership);
  return {
    logo:    cfg.logo        || 'ISP',
    name:    cfg.short_name  || cfg.trade_name || '',
    short:   cfg.short_name  || cfg.trade_name || '',
    pan:     cfg.pan         || '',
    cin:     cfg.cin         || '',
    isPartnership: ispPartner,
    partnershipName: cfg.partnership_name || '',
    cinLabel: ispPartner ? 'Partnership' : 'CIN',
    cinValue: ispPartner ? (cfg.partnership_name || '') : (cfg.cin || ''),
    fssai:   cfg.fssai       || '',
    sbl:     cfg.sbl         || '',
    msme:    cfg.msme_no     || '',
    address1: isStateKL ? (cfg.kl_address1 || cfg.tn_address1 || '') : (cfg.tn_address1 || ''),
    address2: isStateKL ? (cfg.kl_address2 || cfg.tn_address2 || '') : (cfg.tn_address2 || ''),
    place:   isStateKL ? (cfg.kl_place || cfg.tn_place || '') : (cfg.tn_place || ''),
    pin:     isStateKL ? (cfg.kl_pin || cfg.tn_pin || '') : (cfg.tn_pin || ''),
    stateName: isStateKL ? 'KERALA' : 'TAMIL NADU',
    stateCode: isStateKL ? '32' : '33',
    phone:   isStateKL ? (cfg.kl_phone || cfg.tn_phone || '') : (cfg.tn_phone || ''),
    email:   isStateKL ? (cfg.kl_email || cfg.tn_email || '') : (cfg.tn_email || ''),
    gstin:   isStateKL ? (cfg.kl_gstin || cfg.tn_gstin || '') : (cfg.tn_gstin || ''),
  };
}

// Purchase Invoice — layout matches ALL_PURCHASES-1-2.pdf reference.
// Structure: TAX INVOICE title → seller (supplier) block → 2x4 grid of
// transport + invoice details → BILLED/SHIPPED TO → HSN → 9-col table →
// right-side totals block → TDS row → Invoice Amount → words → signatory.
function generatePurchaseInvoicePDF(invoiceData, cfg, invoiceNo, externalDoc) {
  const isBatch = !!externalDoc;
  let doc, buffers;
  if (isBatch) {
    doc = externalDoc;
    if (doc._purchaseCount && doc._purchaseCount > 0) doc.addPage({ size: 'A4', margin: 20 });
    doc._purchaseCount = (doc._purchaseCount || 0) + 1;
  } else {
    doc = new PDFDocument({ size: 'A4', margin: 20 });
    buffers = [];
    doc.on('data', b => buffers.push(b));
  }

  const PAGE_W = doc.page.width;
  const MX = 20;
  const x0 = MX, x1 = PAGE_W - MX, W = x1 - x0;
  let y = MX;

  // Normalize stroke style globally so every line — whether drawn via
  // explicit moveTo/lineTo/stroke OR via rect.fill().stroke() — renders
  // with the same width and color. Prevents the "some lines look darker
  // than others" inconsistency caused by pdfkit defaulting to thicker
  // strokes after a fill operation.
  doc.lineWidth(0.5).strokeColor('#000');

  const fmtQty = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  const fmtRup = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const sumKey = (rows, primary, fallback) => rows.reduce((s, r) => s + Number(r[primary] || (fallback ? r[fallback] : 0) || 0), 0);

  // Extract data
  const seller = invoiceData.seller || {};
  const lineItems = invoiceData.lineItems || [];
  const s = invoiceData.summary || {};

  // ── ORIGINAL/DUPLICATE/TRIPLICATE + TAX INVOICE title above outer border ──
  doc.fontSize(7.5).font('Helvetica').fillColor('#000');
  doc.text('ORIGINAL/DUPLICATE/TRIPLICATE', x0, y, { width: W, align: 'right' });
  y += 9;
  doc.font('Helvetica-Bold').fontSize(10);
  doc.text('TAX INVOICE', x0, y, { width: W, align: 'center' });
  y += 12;

  const boxTopY = y;

  // ── Seller's company header block (top, centered) ──
  // On a Purchase Invoice, the TOP block is the SELLER (supplier) — the
  // company we're buying from. Their full address, state, GSTIN printed.
  y += 4;
  doc.font('Helvetica-Bold').fontSize(11);
  doc.text((seller.name || '').toUpperCase(), x0, y, { width: W, align: 'center' });
  y += 13;
  doc.font('Helvetica').fontSize(8);
  const sellerAddrBits = [seller.address, seller.place, seller.pin].filter(Boolean).join(', ');
  if (sellerAddrBits) { doc.text(sellerAddrBits, x0, y, { width: W, align: 'center' }); y += 10; }
  if (seller.gstin || seller.cr) { doc.text('GSTIN:' + (seller.gstin || seller.cr || ''), x0, y, { width: W, align: 'center' }); y += 10; }
  y += 2;

  // Divider between seller block and details grid
  doc.moveTo(x0, y).lineTo(x1, y).stroke();

  // ── 4-field grid: 2 columns × 4 rows for supplier details + invoice details ──
  //  Left column:  TRANSPORT / VEHICLE NO / STATION / e-TRADE No
  //  Right column: INVOICE NO / DATE / PLACE OF SUPPLY / REVERSE CHARGE
  const gridLeftW = Math.floor(W / 2);
  const gridRightW = W - gridLeftW;
  const gridSplitX = x0 + gridLeftW;
  const gridRowH = 13;
  const gridRows = 4;
  const gridH = gridRowH * gridRows;
  doc.moveTo(gridSplitX, y).lineTo(gridSplitX, y + gridH).stroke();
  doc.moveTo(x0, y + gridH).lineTo(x1, y + gridH).stroke();
  doc.font('Helvetica').fontSize(8);
  // Mode-aware: 'e-TRADE No' / 'e-AUCTION No' driven by cfg.business_mode.
  // Defaults to e-TRADE for legacy installs where the setting is empty.
  const eTradeLabel = (cfg.business_mode === 'e-Auction') ? 'e-AUCTION No' : 'e-TRADE No';
  const leftPairs = [
    ['TRANSPORT', invoiceData.transport || 'BY ROAD'],
    ['VEHICLE NO', invoiceData.vehicleNo || ''],
    ['STATION', invoiceData.station || (seller.place || '').toUpperCase()],
    [eTradeLabel, invoiceData.eTradeNo || ''],
  ];
  const rightPairs = [
    ['INVOICE NO', ''], // value blank per reference
    ['DATE', invoiceData.invoiceDate || new Date().toLocaleDateString('en-GB')],
    ['PLACE OF SUPPLY', (cfg.s_place || '').toUpperCase() + (cfg.s_state ? '  [' + (cfg.s_state || '').toUpperCase() + ']' : '')],
    ['REVERSE CHARGE', ''],
  ];
  // Align the colons within each column (labels left, colons in one column).
  const gridLeftLabelW  = widestLabel(doc, leftPairs.map(p => p[0]));
  const gridRightLabelW = widestLabel(doc, rightPairs.map(p => p[0]));
  for (let i = 0; i < gridRows; i++) {
    const rowY = y + i * gridRowH + 2;
    const [lL, lV] = leftPairs[i];
    const [rL, rV] = rightPairs[i];
    drawKVRow(doc, lL, lV || '', x0 + 6, rowY, gridLeftLabelW, gridLeftW - 12);
    drawKVRow(doc, rL, rV || '', gridSplitX + 6, rowY, gridRightLabelW, gridRightW - 12);
  }
  y += gridH;

  // ── BILLED TO / SHIPPED TO row ──
  // On a purchase invoice from seller's perspective, BILLED TO is US (the
  // purchaser — typically ASP when buying). Both columns show the same
  // ASP block unless shipping to a different address.
  const headH = 14;
  // Explicit horizontal divider ABOVE the header band
  doc.moveTo(x0, y).lineTo(x1, y).stroke();
  doc.rect(x0, y, gridLeftW, headH).fill('#e8e8e8').stroke();
  doc.rect(gridSplitX, y, gridRightW, headH).fill('#e8e8e8').stroke();
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
  doc.text('BILLED TO', x0, y + 3, { width: gridLeftW, align: 'center' });
  doc.text('SHIPPED TO', gridSplitX, y + 3, { width: gridRightW, align: 'center' });
  // Explicit vertical divider between BILLED TO and SHIPPED TO header
  // (drawn AFTER fill so it's not overwritten by the grey)
  doc.moveTo(gridSplitX, y).lineTo(gridSplitX, y + headH).stroke();
  // Explicit horizontal divider BELOW the header band
  doc.moveTo(x0, y + headH).lineTo(x1, y + headH).stroke();
  y += headH;

  // Body
  const buyerLines = [];
  const buyer = invoiceData.buyer || {
    name: cfg.s_company || 'AMAZING SPICE PARK PRIVATE LIMITED',
    address: cfg.s_address1 || '',
    place: cfg.s_place || '',
    pin: cfg.s_pin || '',
    state: cfg.s_state || '',
    st_code: cfg.s_st_code || '',
    gstin: cfg.s_gstin || '',
    pan: cfg.s_pan || cfg.pan || '',
  };
  // Each entry is either a plain line ({ t:'text' }) or an aligned
  // label/colon/value row ({ t:'kv' }) so GSTIN/PAN/STATE colons line up.
  // e-Auction: BILLED/SHIPPED TO uses the company short name instead of the
  // full legal name, per spec.
  const isEAuction = (cfg.business_mode || '').toLowerCase() === 'e-auction';
  const buyerDisplayName = isEAuction
    ? (cfg.s_short_name || cfg.short_name || buyer.name)
    : buyer.name;
  buyerLines.push({ t: 'text', s: (buyerDisplayName || '').toUpperCase() });
  const buyerAddr = [buyer.address, buyer.place ? 'DOOR No.650, ' + buyer.place : ''].filter(Boolean).join(', ').toUpperCase();
  if (buyerAddr) buyerLines.push({ t: 'text', s: buyerAddr });
  if (buyer.gstin) buyerLines.push({ t: 'kv', l: 'GSTIN', v: buyer.gstin });
  if (buyer.pan)   buyerLines.push({ t: 'kv', l: 'PAN',   v: buyer.pan });
  if (buyer.state) buyerLines.push({ t: 'kv', l: 'STATE', v: (buyer.state || '').toUpperCase() + '     CODE: ' + (buyer.st_code || '') });

  const bodyLineH = 10;
  const buyerBodyH = Math.max(6, buyerLines.length + 1) * bodyLineH;
  doc.moveTo(x0, y).lineTo(x0, y + buyerBodyH).stroke();
  doc.moveTo(gridSplitX, y).lineTo(gridSplitX, y + buyerBodyH).stroke();
  doc.moveTo(x1, y).lineTo(x1, y + buyerBodyH).stroke();
  doc.moveTo(x0, y + buyerBodyH).lineTo(x1, y + buyerBodyH).stroke();
  doc.font('Helvetica').fontSize(8);
  const buyerLabelW = widestLabel(doc, buyerLines.filter(it => it.t === 'kv').map(it => it.l));
  let ly = y + 3;
  for (const item of buyerLines) {
    if (item.t === 'kv') {
      drawKVRow(doc, item.l, item.v, x0 + 6, ly, buyerLabelW, gridLeftW - 12);
      drawKVRow(doc, item.l, item.v, gridSplitX + 6, ly, buyerLabelW, gridRightW - 12);
    } else {
      doc.text(item.s, x0 + 6, ly, { width: gridLeftW - 12 });
      doc.text(item.s, gridSplitX + 6, ly, { width: gridRightW - 12 });
    }
    ly += bodyLineH;
  }
  y += buyerBodyH;

  // ── Description of Goods / HSN CODE row ──
  // Honor flag_hsn — when OFF, hide HSN reference and span description
  // across full row width.
  const showHsn = readFlag(cfg.flag_hsn, true);
  const descH = 14;
  if (showHsn) {
    doc.rect(x0, y, gridLeftW, descH).fill('#e8e8e8').stroke();
    doc.rect(gridSplitX, y, gridRightW, descH).fill('#e8e8e8').stroke();
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
    doc.text('Description of Goods:' + (cfg.desc_goods || 'CARDAMOM'), x0, y + 3, { width: gridLeftW, align: 'center' });
    doc.text('HSN CODE:' + (cfg.hsn_cardamom || '09083120'), gridSplitX, y + 3, { width: gridRightW, align: 'center' });
  } else {
    doc.rect(x0, y, W, descH).fill('#e8e8e8').stroke();
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
    doc.text('Description of Goods:' + (cfg.desc_goods || 'CARDAMOM'), x0, y + 3, { width: W, align: 'center' });
  }
  // Horizontal dividers ABOVE and BELOW the description row — drawn
  // AFTER the grey fill so they stay visible (otherwise the fill paints
  // right over them).
  doc.moveTo(x0, y).lineTo(x1, y).stroke();
  doc.moveTo(x0, y + descH).lineTo(x1, y + descH).stroke();
  y += descH;

  // ── Line-item table (same 9-column layout as bill of supply) ──
  const colSpec = [
    { key: 'lot',   label: 'LOT\nNO',                    w: 34 },
    { key: 'qty',   label: 'QUANTITY\nKGS',              w: 66 },
    { key: 'gap1',  label: '',                           w: 6 },
    { key: 'tqty',  label: 'TOTAL\nQTY/KGS',             w: 64 },
    { key: 'price', label: 'PRICE\nRs.  P.',             w: 52 },
    { key: 'value', label: 'VALUE\nRs.     P.',          w: 72 },
    { key: 'gap2',  label: '',                           w: 6 },
    { key: 'tax',   label: 'TAXABLE VALUE\n( Rs.    P.)', w: 80 },
    { key: 'cgst',  label: 'C G S T\n( 2.5%)',           w: 56 },
    { key: 'sgst',  label: 'S G S T\n( 2.5%)',           w: 56 },
    { key: 'igst',  label: 'I G S T\n( 5.0%)',           w: 64 },
  ];
  const totCol = colSpec.reduce((s, c) => s + c.w, 0);
  const scale = W / totCol;
  for (const c of colSpec) c.w = c.w * scale;
  let cx = x0;
  for (const c of colSpec) { c.x = cx; cx += c.w; }

  const hdrH = 22;
  doc.rect(x0, y, W, hdrH).fill('#e8e8e8').stroke();
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(7.5);
  for (const c of colSpec) {
    if (!c.label) continue;
    doc.text(c.label, c.x + 1, y + 2, { width: c.w - 2, align: 'center' });
  }
  // Explicit horizontal divider beneath the LOT NO column header
  doc.moveTo(x0, y + hdrH).lineTo(x1, y + hdrH).stroke();
  y += hdrH;

  // Body rows with alternate-row stripes (matches sales invoice format).
  // flag_invoice_stripe defaults ON; toggle in Settings → Flags.
  const STRIPE_ON = (() => {
    const v = cfg.flag_invoice_stripe;
    if (v === undefined || v === null || v === '') return true;
    if (typeof v === 'boolean') return v;
    return String(v).toLowerCase() === 'true';
  })();
  const STRIPE_COLOR = '#ECECEC';
  function stripeRow(ry, rh, rowIndex) {
    if (!STRIPE_ON) return;
    if (rowIndex % 2 !== 1) return;
    doc.save();
    doc.rect(x0, ry, W, rh).fillColor(STRIPE_COLOR).fill();
    doc.restore();
    doc.fillColor('#000');
  }

  const lineH = 12;
  let bodyStartY = y;
  doc.font('Helvetica').fontSize(9).fillColor('#000');
  // Adaptive row count + multi-page support — same approach as agri bill.
  const PAGE_H = doc.page.height;
  const PAGE_BOTTOM_MARGIN = 20;
  const BOTTOM_RESERVE = 230;
  const targetTableBottom = PAGE_H - PAGE_BOTTOM_MARGIN - BOTTOM_RESERVE;
  let maxRowsThisPage = Math.max(1, Math.floor((targetTableBottom - bodyStartY) / lineH) - 1);
  const MIN_ROWS = 10;

  const tableHeaderTop = bodyStartY - hdrH;
  function paintHeaderAt(y0) {
    doc.rect(x0, y0, W, hdrH).fill('#e8e8e8').stroke();
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(7.5);
    for (const c of colSpec) {
      if (!c.label) continue;
      doc.text(c.label, c.x + 1, y0 + 2, { width: c.w - 2, align: 'center' });
    }
  }
  const pageSegments = [{ start: tableHeaderTop, end: 0 }];
  let drawnRows = 0;

  for (let i = 0; i < lineItems.length; i++) {
    if (drawnRows >= maxRowsThisPage) {
      const segEnd = y;
      pageSegments[pageSegments.length - 1].end = segEnd;
      doc.moveTo(x0, segEnd).lineTo(x1, segEnd).stroke();
      for (const c of colSpec) {
        if (c.key.startsWith('gap')) continue;
        doc.moveTo(c.x, pageSegments[pageSegments.length - 1].start).lineTo(c.x, segEnd).stroke();
      }
      doc.moveTo(x1, pageSegments[pageSegments.length - 1].start).lineTo(x1, segEnd).stroke();
      doc.font('Helvetica-Oblique').fontSize(8).fillColor('#666');
      doc.text('— Continued on next page —', x0, segEnd + 4, { width: W, align: 'center' });
      doc.fillColor('#000');

      doc.addPage({ size: 'A4', margin: 20 });
      y = 20;
      paintHeaderAt(y);
      y += hdrH;
      bodyStartY = y;
      pageSegments.push({ start: y - hdrH, end: 0 });
      maxRowsThisPage = Math.max(1, Math.floor((targetTableBottom - bodyStartY) / lineH) - 1);
      drawnRows = 0;
      doc.font('Helvetica').fontSize(9).fillColor('#000');
    }

    const li = lineItems[i];
    stripeRow(y, lineH, i);
    doc.text(String(li.lot || '').padStart(3, '0'), colSpec[0].x, y + 2, { width: colSpec[0].w, align: 'center' });
    doc.text(fmtQty(li.qty), colSpec[1].x, y + 2, { width: colSpec[1].w - 3, align: 'right' });
    doc.text(fmtQty(li.pqty || li.qty), colSpec[3].x, y + 2, { width: colSpec[3].w - 3, align: 'right' });
    doc.text(fmtRup(li.prate || li.price), colSpec[4].x, y + 2, { width: colSpec[4].w - 3, align: 'right' });
    const value = (li.prate || li.price) * (li.pqty || li.qty);
    doc.text(fmtRup(value), colSpec[5].x, y + 2, { width: colSpec[5].w - 3, align: 'right' });
    doc.text(fmtRup(li.puramt || li.amount), colSpec[7].x, y + 2, { width: colSpec[7].w - 3, align: 'right' });
    if (li.cgst) doc.text(fmtRup(li.cgst), colSpec[8].x, y + 2, { width: colSpec[8].w - 3, align: 'right' });
    if (li.sgst) doc.text(fmtRup(li.sgst), colSpec[9].x, y + 2, { width: colSpec[9].w - 3, align: 'right' });
    if (li.igst) doc.text(fmtRup(li.igst), colSpec[10].x, y + 2, { width: colSpec[10].w - 3, align: 'right' });
    y += lineH;
    drawnRows++;
  }

  // Pad empty rows on the LAST page so totals land at bottom
  const remainingHeight = targetTableBottom - y;
  const minRowsTarget = Math.max(MIN_ROWS - lineItems.length - 1, 0);
  const padRows = Math.max(minRowsTarget, Math.floor(remainingHeight / lineH) - 1);
  for (let i = 0; i < padRows; i++) {
    stripeRow(y, lineH, lineItems.length + i);
    y += lineH;
  }

  // TOTAL row
  const totalRowY = y;
  doc.font('Helvetica-Bold').fontSize(9);
  doc.text('TOTAL', colSpec[0].x, totalRowY + 2, { width: colSpec[0].w, align: 'center' });
  doc.text(fmtQty(sumKey(lineItems, 'qty')), colSpec[1].x, totalRowY + 2, { width: colSpec[1].w - 3, align: 'right' });
  doc.text(fmtQty(sumKey(lineItems, 'pqty', 'qty')), colSpec[3].x, totalRowY + 2, { width: colSpec[3].w - 3, align: 'right' });
  const valueSum = lineItems.reduce((s, li) => s + (li.prate || li.price) * (li.pqty || li.qty), 0);
  doc.text(fmtRup(valueSum), colSpec[5].x, totalRowY + 2, { width: colSpec[5].w - 3, align: 'right' });
  const taxableSum = sumKey(lineItems, 'puramt', 'amount');
  doc.text(fmtRup(taxableSum), colSpec[7].x, totalRowY + 2, { width: colSpec[7].w - 3, align: 'right' });
  const totalCgst = sumKey(lineItems, 'cgst') || s.cgst || 0;
  const totalSgst = sumKey(lineItems, 'sgst') || s.sgst || 0;
  const totalIgst = sumKey(lineItems, 'igst') || s.igst || 0;
  if (totalCgst) doc.text(fmtRup(totalCgst), colSpec[8].x, totalRowY + 2, { width: colSpec[8].w - 3, align: 'right' });
  if (totalSgst) doc.text(fmtRup(totalSgst), colSpec[9].x, totalRowY + 2, { width: colSpec[9].w - 3, align: 'right' });
  if (totalIgst) doc.text(fmtRup(totalIgst), colSpec[10].x, totalRowY + 2, { width: colSpec[10].w - 3, align: 'right' });
  y += lineH;

  // Table borders. For multi-page tables, draw verticals per page segment.
  const tableEndY = y;
  pageSegments[pageSegments.length - 1].end = tableEndY;
  doc.moveTo(x0, tableEndY).lineTo(x1, tableEndY).stroke();
  doc.moveTo(x0, totalRowY).lineTo(x1, totalRowY).stroke();
  for (const seg of pageSegments) {
    for (const c of colSpec) {
      if (c.key.startsWith('gap')) continue;
      doc.moveTo(c.x, seg.start).lineTo(c.x, seg.end).stroke();
    }
    doc.moveTo(x1, seg.start).lineTo(x1, seg.end).stroke();
  }

  // ── Right-side totals block ──
  const sumX = x0 + W * 0.48;
  const sumBlockW = W - (sumX - x0);
  const sumLabelW = sumBlockW * 0.58;
  const sumValW = sumBlockW - sumLabelW;
  const sumRowsData = [
    ['Total Taxable Value', taxableSum],
    ['Total Integrated Tax', totalIgst],
    ['Total Central Tax', totalCgst],
    ['Total State Tax', totalSgst],
    ['Round UP/DOWN', s.roundDiff || 0],
  ];
  // Additional Charge (Settings → Rates & Charges) — adds onto Total Value
  // below the round line. Hidden when zero.
  if (s.addlCharge && s.addlCharge !== 0) {
    sumRowsData.push([(s.addlChargeName || cfg.addl_charge_name || 'Additional Charge').toString(), s.addlCharge]);
  }
  // Capture the y-start so we can later draw a vertical separator
  // between the label column and the value column (spans from the first
  // row down through the Total Value row).
  const sumStartY = y;
  doc.font('Helvetica').fontSize(8.5);
  for (const [lbl, v] of sumRowsData) {
    doc.text(lbl, sumX + 4, y + 2, { width: sumLabelW });
    doc.text(fmtRup(v), sumX + sumLabelW, y + 2, { width: sumValW - 4, align: 'right' });
    y += 12;
  }
  doc.font('Helvetica-Bold').fontSize(9);
  doc.text('Total Value', sumX + 4, y + 2, { width: sumLabelW });
  const totalValue = s.grandTotal || (taxableSum + totalCgst + totalSgst + totalIgst + (s.roundDiff || 0));
  doc.text(fmtRup(totalValue), sumX + sumLabelW, y + 2, { width: sumValW - 4, align: 'right' });
  y += 14;
  // Vertical separator between the label column and the value column,
  // spanning the full totals block height.
  doc.moveTo(sumX + sumLabelW, sumStartY).lineTo(sumX + sumLabelW, y).stroke();

  // DC No. / Date: on the left, aligned with Total Value row
  doc.font('Helvetica').fontSize(8);
  doc.text('DC No.', x0 + 6, y - 14);
  doc.text('Date:', x0 + W * 0.22, y - 14);

  // Separator before TDS / invoice amount
  doc.moveTo(x0, y).lineTo(x1, y).stroke(); y += 4;

  // ── TDS + Invoice Amount rows ──
  // Purchase invoices for sellers with GSTIN carry 0.1% TDS U/S 194Q
  const tdsAmount = s.tdsAmount || 0;
  const invoiceAmount = s.invoiceAmount || (totalValue - tdsAmount);
  const tdsRowH = 14;
  const amtX = x1 - 110;  // start of amount column (right-aligned inside 100pt)
  const amtW = 104;
  doc.font('Helvetica').fontSize(8.5);
  doc.text('TDS on Purchase of Goods [ U/S 194Q ]', x0 + 6, y + 2, { width: amtX - x0 - 12 });
  doc.text('-' + fmtRup(Math.abs(tdsAmount)), amtX, y + 2, { width: amtW, align: 'right' });
  y += tdsRowH;
  doc.font('Helvetica-Bold').fontSize(9);
  doc.text('Invoice Amount', x0 + 6, y + 2, { width: amtX - x0 - 12 });
  doc.text(fmtRup(invoiceAmount), amtX, y + 2, { width: amtW, align: 'right' });
  y += tdsRowH;

  // Separator before amount in words
  doc.moveTo(x0, y).lineTo(x1, y).stroke(); y += 6;

  // ── Amount in words ──
  doc.fontSize(9).font('Helvetica');
  const forWords = Math.round(invoiceAmount);
  doc.text(amountToWords(forWords) + ' Only', x0 + 6, y, { width: W - 12 });
  y += 20;

  // ── For {SELLER} at right ──
  doc.fontSize(9).font('Helvetica-Bold');
  doc.text('For ' + (seller.name || '').toUpperCase(), x0, y, { width: W - 6, align: 'right' });
  y += 30;

  // Authorized signatory
  doc.font('Helvetica').fontSize(8);
  doc.text('Authorised Signatory', x0, y, { width: W - 6, align: 'right' });
  y += 14;

  // Outer border. For single-page invoices, wrap from the title strip to
  // bottom. For multi-page, only wrap the last page's bottom block.
  const boxBottomY = y;
  if (pageSegments.length === 1) {
    doc.rect(x0, boxTopY, W, boxBottomY - boxTopY).lineWidth(1).stroke();
  } else {
    doc.rect(x0, pageSegments[pageSegments.length - 1].start, W, boxBottomY - pageSegments[pageSegments.length - 1].start).lineWidth(1).stroke();
  }

  if (isBatch) return Promise.resolve(null);
  return new Promise((resolve) => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}


/**
 * Generate a crop receipt PDF (CROASP.PRG equivalent)
 */
function generateCropReceiptPDF(lot, cfg) {
  const co = effectiveCompany(cfg);
  const doc = new PDFDocument({ size: [595, 420], margin: 25 }); // half-page
  const buffers = [];
  doc.on('data', b => buffers.push(b));
  
  const w = 545; const x = 25; let y = 25;

  doc.rect(x, y, w, 370).stroke();
  y += 10;
  doc.fontSize(16).font('Helvetica-Bold').text('RECEIPT', x, y, { align: 'center', width: w });
  y += 20;
  doc.fontSize(8).font('Helvetica').text(`Sl.No: ${lot.crop || ''}`, x + w - 100, y - 10, { width: 90, align: 'right' });
  
  const companyName = co.name;
  doc.fontSize(11).font('Helvetica-Bold').text(companyName, x, y, { align: 'center', width: w });
  y += 14;
  doc.fontSize(7).font('Helvetica');
  doc.text(co.address1, x, y, { align: 'center', width: w }); y += 10;
  doc.text(`GST No. ${co.gstin}`, x, y, { align: 'center', width: w }); y += 16;

  // Details grid
  const details = [
    ['Trade No', lot.ano || ''],
    ['Lot No', lot.lot_no || ''],
    ['Date', new Date().toLocaleDateString('en-GB')],
    ['No. of Bags', String(lot.bags || '')],
    ['Nett Weight', String(lot.qty || '')],
    ['Depot', lot.branch || ''],
  ];

  doc.fontSize(8);
  let col = 0;
  for (const [label, val] of details) {
    const cx = x + 10 + (col % 3) * 180;
    const cy = y + Math.floor(col / 3) * 16;
    doc.font('Helvetica').text(`${label}: `, cx, cy, { continued: true });
    doc.font('Helvetica-Bold').text(val);
    col++;
  }
  y += 40;

  // Declaration text
  doc.font('Helvetica').fontSize(7);
  doc.text(`We acknowledge the receipt of Cardamom as per the description above, from`, x + 10, y, { width: w - 20 });
  y += 11;
  doc.text(`M/s. ${lot.name || ''}`, x + 10, y, { width: w - 20 }); y += 11;
  doc.text(`GSTIN/CR No. ${lot.cr || ''}`, x + 10, y, { width: w - 20 }); y += 18;

  // Signatures
  y += 30;
  doc.text('Pooler Signature', x + 10, y);
  doc.text('[ Contact Number ]', x + w/2 - 50, y, { width: 100, align: 'center' });
  doc.text('Depot in Charge', x + w - 120, y);

  return new Promise((resolve) => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

// ── Lot-Entry Receipt PDF ────────────────────────────────────
// A4-portrait receipt for one seller's lots — the printable companion to
// the lot-entry WhatsApp text summary. Company header + Trade/date + a
// lots table (Lot / Bags / Qty / Rate / Amount) + totals. Mirrors the
// content of the on-screen lot-entry receipt so the seller gets the same
// document over WhatsApp.
function generateLotReceiptPDF(lots, cfg, opts) {
  opts = opts || {};
  const co = effectiveCompany(cfg);
  const traderName = opts.traderName || (lots[0] && lots[0].name) || '';
  const ano = opts.ano != null ? opts.ano : (lots[0] && lots[0].ano) || '';
  const dateStr = opts.date
    ? String(opts.date).slice(0, 10).split('-').reverse().join('/')
    : new Date().toLocaleDateString('en-GB');

  const doc = new PDFDocument({ size: 'A4', margin: 36 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));

  const x = 36, w = 523; let y = 36;

  // Header
  doc.fillColor('#000').fontSize(15).font('Helvetica-Bold')
     .text((co.name || 'COMPANY').toUpperCase(), x, y, { align: 'center', width: w });
  y += 19;
  doc.fontSize(8).font('Helvetica');
  const addrLine = [co.address1, co.address2].filter(Boolean).join(', ');
  if (addrLine) { doc.text(addrLine, x, y, { align: 'center', width: w }); y += 10; }
  const placeLine = [co.place, co.pin].filter(Boolean).join(' - ');
  if (placeLine) { doc.text(placeLine, x, y, { align: 'center', width: w }); y += 10; }
  if (co.gstin) { doc.text('GSTIN: ' + co.gstin, x, y, { align: 'center', width: w }); y += 12; }

  doc.moveTo(x, y).lineTo(x + w, y).stroke(); y += 9;
  doc.fontSize(12).font('Helvetica-Bold').text('LOT RECEIPT', x, y, { align: 'center', width: w });
  y += 18;

  // Meta line: Trade No (left) + Date (right)
  doc.fontSize(9).font('Helvetica-Bold').text('Trade No: ', x, y, { continued: true })
     .font('Helvetica').text(String(ano || '-'));
  doc.font('Helvetica-Bold').text('Date: ', x + w - 150, y, { continued: true, width: 150 })
     .font('Helvetica').text(dateStr);
  y += 14;
  doc.font('Helvetica-Bold').text('To: ', x, y, { continued: true }).font('Helvetica').text(traderName || '-');
  y += 16;

  // Table columns
  const cols = [
    { key: 'lot',  label: 'Lot',     cx: x,        cw: 70,        align: 'left'  },
    { key: 'bags', label: 'Bags',    cx: x + 70,   cw: 70,        align: 'right' },
    { key: 'qty',  label: 'Qty (kg)',cx: x + 140,  cw: 120,       align: 'right' },
    { key: 'rate', label: 'Rate',    cx: x + 260,  cw: 120,       align: 'right' },
    { key: 'amt',  label: 'Amount',  cx: x + 380,  cw: w - 380,   align: 'right' },
  ];
  const fmtN = (n, dp = 2) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: dp, maximumFractionDigits: dp });

  const drawHeader = () => {
    doc.rect(x, y, w, 16).fill('#f0f0f0'); doc.fillColor('#000');
    doc.fontSize(8).font('Helvetica-Bold');
    for (const c of cols) doc.text(c.label, c.cx + 3, y + 4, { width: c.cw - 6, align: c.align });
    y += 16;
  };
  drawHeader();

  let totBags = 0, totQty = 0, totAmt = 0;
  doc.font('Helvetica').fontSize(8);
  for (const l of lots) {
    if (y > 760) { doc.addPage(); y = 40; drawHeader(); doc.font('Helvetica').fontSize(8); }
    const bags = Number(l.bags) || 0, qty = Number(l.qty) || 0, price = Number(l.price) || 0;
    const amt = Number(l.amount) || (qty * price);
    totBags += bags; totQty += qty; totAmt += amt;
    doc.rect(x, y, w, 14).stroke();
    doc.fillColor('#000');
    doc.text(String(l.lot_no || ''), cols[0].cx + 3, y + 3, { width: cols[0].cw - 6, align: cols[0].align });
    doc.text(String(bags),           cols[1].cx + 3, y + 3, { width: cols[1].cw - 6, align: cols[1].align });
    doc.text(fmtN(qty, 3),           cols[2].cx + 3, y + 3, { width: cols[2].cw - 6, align: cols[2].align });
    doc.text(fmtN(price),            cols[3].cx + 3, y + 3, { width: cols[3].cw - 6, align: cols[3].align });
    doc.text(fmtN(amt),              cols[4].cx + 3, y + 3, { width: cols[4].cw - 6, align: cols[4].align });
    y += 14;
  }

  // Totals row
  doc.rect(x, y, w, 16).fill('#f0f0f0'); doc.fillColor('#000');
  doc.font('Helvetica-Bold').fontSize(8);
  doc.text('Total',          cols[0].cx + 3, y + 4, { width: cols[0].cw - 6, align: 'left'  });
  doc.text(String(totBags),  cols[1].cx + 3, y + 4, { width: cols[1].cw - 6, align: 'right' });
  doc.text(fmtN(totQty, 3),  cols[2].cx + 3, y + 4, { width: cols[2].cw - 6, align: 'right' });
  doc.text(fmtN(totAmt),     cols[4].cx + 3, y + 4, { width: cols[4].cw - 6, align: 'right' });
  y += 28;

  doc.font('Helvetica').fontSize(7).fillColor('#555')
     .text('This is a computer-generated lot receipt.', x, y, { align: 'center', width: w });

  return new Promise((resolve) => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

module.exports = { generatePurchaseInvoicePDF, generateCropReceiptPDF, generateAgriBillPDF, generateSalesInvoicePDF, generateSalesInvoicesBatchPDF, generatePurchaseInvoicesBatchPDF, generateAgriBillsBatchPDF, generateLotReceiptPDF, generateCommissionBoSPDF, generateCommissionBoSBatchPDF };

/**
 * Sales Invoice PDF (Tax Invoice)
 * Grid-based layout matching the legal GST format:
 *   - Supplier (ISPL) top-left with logo
 *   - Invoice metadata grid top-right (Invoice No, Date, Other References, etc.)
 *   - Consignee (Ship to) + Buyer (Bill to) stacked on left
 *   - Dispatch From (sister company ASP) on right
 *   - Line items table with HSN/SAC, Shipped/Billed qty, Rate, Amount
 *   - HSN summary + tax breakup
 *   - Bank details + signature block
 * GST logic:
 *   - Sale 'L' (Local)       → CGST + SGST
 *   - Sale 'I' (Inter-state) → IGST
 *   - Sale 'E' (Export)      → Zero-rated
 */
// When called without `externalDoc`, builds a standalone PDF for one invoice.
// When called with `externalDoc` (a PDFKit doc), appends a new page and draws
// this invoice into it — used by the batch-print endpoint to merge many
// invoices into a single file. Returns the Buffer only in standalone mode.
function generateSalesInvoicePDF(invoiceData, cfg, saleType, invoiceNo, invoiceDate, externalDoc, variant) {
  const isBatch = !!externalDoc;
  // Variant = 'purchase' renders the mirror-image view of an ASP sales invoice:
  //   - Top-left company block shows ISPL (the receiving company) instead of ASP
  //   - "Buyer (Bill to)" label becomes "Seller (Bill from)" and displays ASP
  //   - Bank details show TN bank (ISPL's)
  //   - Title shows "Purchase Invoice" instead of "Tax Invoice"
  // All line-item math stays identical (same P_Rate, PurAmt, totals, HSN).
  // Only valid for ASP invoices; caller must ensure that.
  const isPurchaseView = (variant === 'purchase');
  // For the purchase view, we FLIP the top-of-page issuer to ISPL by pretending
  // the effective company is the primary (not sister). Easiest way: compute
  // effectiveCompany with a forced-TN cfg clone. The rest of the function still
  // uses the original cfg so ASP business rules (P_Rate formula, etc.) stay in
  // effect — only the display swaps.
  const co = isPurchaseView
    ? effectiveCompany({ ...cfg, business_state: 'TAMIL NADU' })
    : effectiveCompany(cfg);
  let doc, buffers;
  if (isBatch) {
    doc = externalDoc;
    // For invoices after the first, start on a fresh page; first call re-uses the empty initial page
    if (doc._invoiceCount && doc._invoiceCount > 0) {
      doc.addPage({ size: 'A4', margin: 20 });
    }
    doc._invoiceCount = (doc._invoiceCount || 0) + 1;
  } else {
    doc = new PDFDocument({ size: 'A4', margin: 20 });
    buffers = [];
    doc.on('data', b => buffers.push(b));
  }

  const { buyer, lineItems, summary } = invoiceData;

  // ── Invoice provenance flags ────────────────────────────────
  // isASP = this invoice is issued by ASP (sister company) instead of ISPL.
  // Triggered when business mode = e-Trade AND state = KERALA.
  // Drives: invoice-prefix swap, Transport/Insurance hidden, bank details
  //         picked from KL bank settings, "Dispatch From" hidden (ASP
  //         invoices don't reference a separate dispatch), logo swap.
  const isASP = (cfg.business_mode || '').toLowerCase() === 'e-trade'
             && (cfg.business_state || '').toUpperCase() === 'KERALA';
  // e-Auction mode: no ASP↔ISP mirror, so the "Other References" cross-ref
  // is meaningless and is blanked below.
  const isEAuction = (cfg.business_mode || '').toLowerCase() === 'e-auction';
  // ISP inter-state invoices should not bill Transport/Insurance separately
  // (item 5: those costs are absorbed in the buyer's freight, not invoiced).
  // Combine with the existing ASP rule: ASP also hides them.
  // e-Auction interstate is an exception — it DOES bill Transport/Insurance,
  // so don't suppress them there (mirrors calculations.js hideTI).
  const hideTransportInsurance = isASP || (!isEAuction && String(saleType || '').toUpperCase() === 'I');
  // Read a boolean-ish flag cfg value, respecting undefined→default semantics.
  // Used for user-facing toggle flags like flag_ship, flag_dispatch.
  function readFlagSafe(val, defaultOn) {
    if (val === undefined || val === null || val === '') return defaultOn;
    if (typeof val === 'boolean') return val;
    return String(val).toLowerCase() === 'true';
  }
  // Ship-To visibility:
  //   ASP invoices → ALWAYS hidden (Consignee block doesn't apply to the
  //     ASP→ISP internal transfer, per user spec)
  //   ISP invoices → ALWAYS shown (per business rule: TN invoices need
  //     buyer ship-to and dispatch addresses on the PDF). The cfg flags
  //     flag_ship / flag_dispatch are kept for backward compat but
  //     overridden in the TN/ISP path.
  const showShipTo   = isASP ? false : true;
  const showDispatch = isASP ? readFlagSafe(cfg.flag_dispatch, true) : true;

  // ── Page geometry ───────────────────────────────────────────
  const pageW = doc.page.width;
  const pageH = doc.page.height;
  const margin = 20;
  const x0 = margin;
  const x1 = pageW - margin;
  const W  = x1 - x0;
  let y = margin;

  // Normalize stroke style — uniform 0.5pt black for all lines including
  // grey-filled rect outlines.
  doc.lineWidth(0.5).strokeColor('#000');

  // ── Cell drawing helpers ────────────────────────────────────
  // Draws a bordered box and optionally text inside. Does NOT advance y.
  function box(bx, by, bw, bh) {
    doc.lineWidth(0.5).rect(bx, by, bw, bh).stroke();
  }
  // Draw text inside a cell with optional padding and label/value style.
  // opts: {font, size, align, label, labelFont, labelSize, color}
  function cellText(bx, by, bw, bh, text, opts = {}) {
    const pad = opts.pad != null ? opts.pad : 3;
    const size = opts.size || 8;
    const font = opts.font || 'Helvetica';
    doc.font(font).fontSize(size);
    if (opts.color) doc.fillColor(opts.color);
    const align = opts.align || 'left';
    doc.text(text || '', bx + pad, by + pad, { width: bw - pad * 2, height: bh - pad * 2, align, lineBreak: opts.lineBreak !== false });
    if (opts.color) doc.fillColor('#000');
  }
  // Draw a labeled cell: small label on top, value below in bold
  function labeledCell(bx, by, bw, bh, label, value, opts = {}) {
    box(bx, by, bw, bh);
    doc.font('Helvetica').fontSize(7).fillColor('#000');
    doc.text(label, bx + 3, by + 2, { width: bw - 6 });
    if (value) {
      doc.font(opts.valueFont || 'Helvetica-Bold').fontSize(opts.valueSize || 9);
      doc.text(value, bx + 3, by + 11, { width: bw - 6 });
    }
  }

  // Draws ONLY the vertical dividers + outer left/right borders for a row
  // in the line-items table. Use this instead of `box()` for row rendering
  // so horizontal lines don't appear between rows. Horizontal boundaries
  // around the whole table are drawn once at the top (under header) and once
  // at the bottom (under the total row).
  function rowVerticals(ry, rh, skipBilledSplit) {
    // Outer left/right borders
    doc.lineWidth(0.5);
    doc.moveTo(x0, ry).lineTo(x0, ry + rh).stroke();
    doc.moveTo(x0 + W, ry).lineTo(x0 + W, ry + rh).stroke();
    // Inner column dividers
    for (const k of cols) {
      const cx = colX(k);
      if (cx <= x0) continue;
      if (skipBilledSplit && k === 'billed') continue;
      doc.moveTo(cx, ry).lineTo(cx, ry + rh).stroke();
    }
  }

  // Alternate-row stripes for line-items + HSN summary tables.
  // Controlled by the `flag_invoice_stripe` setting (default on).
  // IMPORTANT: getSettingsFlat coerces boolean settings to real JS booleans,
  // so cfg.flag_invoice_stripe may be `true` or `false` — NOT strings.
  // Using `|| 'true'` would convert `false` to the default, hiding the toggle.
  // Treat undefined/null as the default (on); any other value must be interpreted literally.
  function readFlag(val, defaultOn) {
    if (val === undefined || val === null || val === '') return defaultOn;
    if (typeof val === 'boolean') return val;
    return String(val).toLowerCase() === 'true';
  }
  const STRIPE_ON = readFlag(cfg.flag_invoice_stripe, true);
  const STRIPE_COLOR = '#ECECEC';
  function stripeFill(ry, rh, rowIndex) {
    if (!STRIPE_ON) return;
    if (rowIndex % 2 !== 1) return; // only odd rows (alternate)
    doc.save();
    doc.rect(x0, ry, W, rh).fillColor(STRIPE_COLOR).fill();
    doc.restore();
    doc.fillColor('#000'); // reset for text
  }

  // ── Title ───────────────────────────────────────────────────
  doc.font('Helvetica-Bold').fontSize(10).text(isPurchaseView ? 'Purchase Invoice' : 'Tax Invoice', x0, y, { width: W, align: 'center' });
  y += 14;

  // ── TOP HEADER BLOCK ────────────────────────────────────────
  // Left half: Logo + ISPL details
  // Right half: 2-col sub-grid (Invoice No, e-Way Bill No, Dated / Delivery Note, Mode/Terms / Ref No., Other Refs)
  const topY = y;
  const leftW = W * 0.5;
  const rightW = W - leftW;
  const leftX = x0;
  const rightX = x0 + leftW;

  // Right-side grid cell sizes — 2 rows, sized so combined height matches
  // the left company-details block (logo + name + address + GSTIN + State + CIN + optional FSSAI + SBL).
  const topHeaderH = 100;
  const rRow = topHeaderH / 2;       // 50pt each row
  const rCell = rightW / 2;          // each cell's width for 2-col rows

  // ── LEFT BLOCK: Logo + company details ──────────────────────
  box(leftX, topY, leftW, topHeaderH);
  // Logo file pick. Kerala installs operate as the Sister Company (ASP) in
  // BOTH e-Trade and e-Auction modes (matches effectiveCompany), so the logo
  // follows business_state, not mode. For ASP purchase view (issuer is ISPL,
  // not ASP) → ISPL logo.
  const useASPLogo = !isPurchaseView
                  && (cfg.business_state || '').toUpperCase() === 'KERALA';
  const logoFile = useASPLogo ? 'logo-asp.png' : 'logo-ispl.png';
  const logoPath = require('path').join(__dirname, 'public', logoFile);
  const fs = require('fs');
  let logoDrawn = false;
  if (fs.existsSync(logoPath)) {
    try {
      doc.image(logoPath, leftX + 4, topY + 4, { fit: [60, 60] });
      logoDrawn = true;
    } catch (_) { /* fall through to text */ }
  }
  const textX = leftX + (logoDrawn ? 70 : 8);
  const textW = leftW - (logoDrawn ? 78 : 16);
  let ty = topY + 4;
  doc.font('Helvetica-Bold').fontSize(10).text(co.name || '', textX, ty, { width: textW });
  ty += 12;
  doc.font('Helvetica').fontSize(8);
  const addrLine = [co.address1, co.address2, co.place, co.stateName, co.pin].filter(Boolean).join(', ');
  doc.text(addrLine, textX, ty, { width: textW });
  ty += doc.heightOfString(addrLine, { width: textW });
  // Identity line: "Partnership: <name>" when the partnership toggle is
  // on for the active company, else "CIN: <cin>". Falls back to the raw
  // CIN if cinLabel/cinValue aren't resolved (older cfg shapes).
  const _idLabel = co.cinLabel || 'CIN';
  const _idValue = (co.cinValue !== undefined ? co.cinValue : co.cin) || '';
  // Build the label/value rows present, then draw them with a single
  // colon column so every ":" lines up. (MSME / SBL / FSSAI fit inside the
  // fixed 100pt header box alongside name + address + GSTIN + State + CIN.)
  const coRows = [];
  if (co.gstin)     coRows.push(['GSTIN/UIN', co.gstin]);
  if (co.stateName) coRows.push(['State Name', `${co.stateName}, Code : ${co.stateCode}`]);
  if (_idValue)     coRows.push([_idLabel, _idValue]);
  if (co.fssai)     coRows.push(['FSSAI No.', co.fssai]);
  if (co.sbl)       coRows.push(['SBL No.', co.sbl]);
  if (co.msme)      coRows.push(['MSME No.', co.msme]);
  if (coRows.length) {
    const coLabelW = widestLabel(doc, coRows.map(r => r[0]));
    for (const [lab, val] of coRows) {
      drawKVRow(doc, lab, val, textX, ty, coLabelW, textW);
      ty += 10;
    }
  }

  // ── RIGHT BLOCK: 2-row metadata grid ────────────────────────
  // Row 1: Invoice No | e-Way Bill No | Dated
  // Row 2: Reference No. & Date | Other References

  const r1W = rightW / 3;
  let ry = topY;

  // Row 1
  // Invoice No: primary prefix flips to ASP when this is an ASP-issued invoice.
  // Other Refs is always the OTHER company (ISP vs ASP are mirrors).
  // The prefix is driven SOLELY by the "Invoice Prefix" / "Sister Invoice
  // Prefix" settings. When the user clears one, NO prefix is emitted — the
  // number renders prefix-less (e.g. "L-9/26-27"). There is deliberately NO
  // fallback to the Logo Code or company short name.
  const primaryPrefix = isASP ? (cfg.inv_prefix_sister || '') : (cfg.inv_prefix || '');
  const otherPrefix   = isASP ? (cfg.inv_prefix || '')         : (cfg.inv_prefix_sister || '');
  const primaryCfg    = { ...cfg, inv_prefix: primaryPrefix };
  // ASP invoices always use the "I" segment irrespective of local/interstate
  // sale type (format: ASP/I-{invno}/{season_short}).
  const displaySaleType = isASP ? 'I' : saleType;
  // Row-1 height: expand to fill both rows' worth of space for ASP (since
  // Row 2 — Reference No. + Other References — is hidden for ASP).
  const row1H = isASP ? (rRow * 2) : rRow;
  labeledCell(rightX,             ry, r1W, row1H, 'Invoice No.', formatInvoiceNo(primaryCfg, displaySaleType, invoiceNo));
  labeledCell(rightX + r1W,       ry, r1W, row1H, 'e-Way Bill No.', '');
  labeledCell(rightX + r1W * 2,   ry, rightW - r1W * 2, row1H, 'Dated', (() => {
    const d = invoiceDate ? new Date(invoiceDate) : new Date();
    const day = String(d.getDate()).padStart(2, '0');
    const mon = d.toLocaleDateString('en-US', { month: 'short' });
    const yr  = String(d.getFullYear()).slice(-2);
    return `${day}-${mon}-${yr}`;
  })());
  ry += row1H;

  // Row 2 — ONLY for ISP invoices (ASP omits Reference No. / Other References)
  if (!isASP) {
    // Other References shows the cross-referenced ASP invoice number when
    // available. Format is fixed: ASP/I-{aspInvo}/{season} — the middle
    // segment is hardcoded to 'I' per business convention (ASP→ISP transfer
    // is always treated as inter-state for the cross-reference label,
    // regardless of the actual ISP sale type to the external buyer).
    const aspInvo = invoiceData && invoiceData.aspInvo;
    const otherRefCfg = { ...cfg, inv_prefix: otherPrefix };
    // e-Auction has no mirrored ASP invoice to cross-reference → leave blank.
    const otherRefs = isEAuction ? '' : formatInvoiceNo(otherRefCfg, 'I', aspInvo || invoiceNo);
    labeledCell(rightX,         ry, rCell, rRow, 'Reference No. & Date.', '');
    labeledCell(rightX + rCell, ry, rCell, rRow, 'Other References', otherRefs);
  }

  y = topY + topHeaderH;

  // ── MIDDLE BLOCK: Consignee + Buyer + Dispatch ──────────────
  // Left column: Consignee (Ship to) stacked above Buyer (Bill to)
  // Right column: Dispatched through | Destination (top row)
  //               Dispatch From (ASP) — fills remaining height

  const midH = 150; // total middle-block height
  const midY = y;

  // Left column structure:
  //   - Both blocks shown: 2 stacked cells (Consignee on top, Buyer below)
  //   - flag_ship OFF: only Buyer block, full midH height
  const leftCellH = showShipTo ? midH / 2 : midH;
  const consigneeY = midY;
  const buyerY = showShipTo ? (midY + (midH / 2)) : midY;

  // Consignee (Ship to) — uses consignee fields if present, else falls back to buyer
  const hasConsignee = !!(buyer.cbuyer1 || buyer.cadd1 || buyer.cpla || buyer.cgstin);
  const ship = hasConsignee ? {
    name: buyer.cbuyer1 || buyer.buyer1 || buyer.buyer || '',
    addr: buyer.cadd1 || '',
    pla:  buyer.cpla  || '',
    pin:  buyer.cpin  || '',
    state: buyer.cstate || '',
    stCode: buyer.cst_code || '',
    gstin: buyer.cgstin || '',
  } : {
    name: buyer.buyer1 || buyer.buyer || '',
    addr: [buyer.add1, buyer.add2].filter(Boolean).join(','),
    pla: buyer.pla || '',
    pin: buyer.pin || '',
    state: buyer.state || '',
    stCode: buyer.st_code || '',
    gstin: buyer.gstin || '',
  };

  // Shared helper for left-column writing, advancing the anchor by actual height
  const lW = leftW - 6;
  const writeLeft = (txt, anchor) => {
    if (!txt) return;
    doc.text(txt, leftX + 3, anchor.v, { width: lW });
    anchor.v += doc.heightOfString(txt, { width: lW }) + 1;
  };
  // Aligned label/colon/value variant — GSTIN/UIN and State Name share one
  // colon column so they line up. labelColW is the widest of the two labels
  // (both blocks use the same pair). Measured lazily under the active font.
  const kvLabelW = () => widestLabel(doc, ['GSTIN/UIN', 'State Name']);
  const writeLeftKV = (label, value, anchor) => {
    const vw = drawKVRow(doc, label, value, leftX + 3, anchor.v, kvLabelW(), lW);
    anchor.v += doc.heightOfString(String(value == null ? '' : value), { width: vw }) + 1;
  };

  // Draw Consignee cell (skip when flag_ship is OFF — item 1)
  if (showShipTo) {
    box(leftX, consigneeY, leftW, (midH / 2));
    let cy = consigneeY + 3;
    doc.font('Helvetica').fontSize(7).text('Consignee (Ship to)', leftX + 3, cy); cy += 9;
    doc.font('Helvetica-Bold').fontSize(9).text(ship.name, leftX + 3, cy, { width: leftW - 6 });
    // Advance by actual rendered height (prevents overlap for long company names)
    cy += doc.heightOfString(ship.name, { width: leftW - 6 }) + 2;
    doc.font('Helvetica').fontSize(8);
    const cAnchor = { v: cy };
    writeLeft(ship.addr, cAnchor);
    writeLeft(ship.pla,  cAnchor);
    if (ship.gstin) writeLeftKV('GSTIN/UIN', ship.gstin, cAnchor);
    if (ship.state) writeLeftKV('State Name', `${ship.state}, Code : ${ship.stCode || ''}`, cAnchor);
  }

  // Draw Buyer (Bill to) / Seller (Bill from) cell
  box(leftX, buyerY, leftW, leftCellH);
  let by = buyerY + 3;
  // Label flips for purchase view: "Seller (Bill from)" since ISPL is the buyer
  // here, and ASP is the seller we're buying from.
  const blockLabel = isPurchaseView ? 'Seller (Bill from)' : 'Buyer (Bill to)';
  doc.font('Helvetica').fontSize(7).text(blockLabel, leftX + 3, by); by += 9;

  // Entity displayed in this block:
  //   Purchase view (ASP sales mirrored)  → ASP sister company (seller)
  //   ASP sales invoice (normal)          → ISP (the buyer in ASP→ISP flow)
  //   ISP sales invoice                   → external customer
  let billTo;
  if (isPurchaseView) {
    billTo = {
      name:  cfg.s_company || 'AMAZING SPICE PARK PRIVATE LIMITED',
      add1:  cfg.s_address1 || '',
      add2:  cfg.s_address2 || '',
      pla:   cfg.s_place || '',
      gstin: cfg.s_gstin || '',
      state: cfg.s_state || 'KERALA',
      stCode: cfg.s_st_code || '32',
    };
  } else if (isASP) {
    billTo = {
      name:  cfg.short_name || cfg.trade_name || 'IDEAL SPICES PRIVATE LIMITED',
      add1:  cfg.tn_address1 || '',
      add2:  cfg.tn_address2 || '',
      pla:   cfg.tn_place || '',
      gstin: cfg.tn_gstin || '',
      state: cfg.tn_state || 'TAMIL NADU',
      stCode: cfg.tn_st_code || '33',
    };
  } else {
    billTo = {
      name:  buyer.buyer1 || buyer.buyer || '',
      add1:  buyer.add1 || '',
      add2:  buyer.add2 || '',
      pla:   buyer.pla || '',
      gstin: buyer.gstin || '',
      state: buyer.state || '',
      stCode: buyer.st_code || '',
    };
  }

  doc.font('Helvetica-Bold').fontSize(9).text(billTo.name, leftX + 3, by, { width: leftW - 6 });
  // Advance by ACTUAL rendered height — long company names wrap to 2+ lines.
  // Using a fixed advance caused subsequent address lines to overlap the name.
  by += doc.heightOfString(billTo.name, { width: leftW - 6 }) + 2;
  doc.font('Helvetica').fontSize(8);
  const bAddr = [billTo.add1, billTo.add2].filter(Boolean).join(',');
  const bAnchor = { v: by };
  writeLeft(bAddr, bAnchor);
  writeLeft(billTo.pla, bAnchor);
  if (billTo.gstin) writeLeftKV('GSTIN/UIN', billTo.gstin, bAnchor);
  if (billTo.state) writeLeftKV('State Name', `${billTo.state}, Code : ${billTo.stCode || ''}`, bAnchor);

  // Right column: 2 rows now (Dispatched through | Destination, Dispatch From)
  const rSmall = 28;
  let ry2 = midY;

  // Dispatched through: prefer a per-call override (passed in invoiceData
  // from the print modal) over the stored default. Falls back to ISP/ASP
  // variants in cfg, then the global default.
  const dispThrough = (invoiceData && invoiceData.dispatchedThrough)
    || (isASP
      ? (cfg.dispatched_through_asp || cfg.dispatched_through || '')
      : (cfg.dispatched_through_isp || cfg.dispatched_through || ''));
  labeledCell(rightX,         ry2, rCell, rSmall, 'Dispatched through', dispThrough);
  // Destination for ISP invoices uses the buyer's CONSIGNEE place (cpla)
  // since ISP delivers TO the consignee/ship-to address. Falls back to
  // buyer's main place (pla) if no separate consignee is set, then to
  // configured dispatch_destination as a last resort.
  // ASP invoices use the configured dispatch_destination directly because
  // ASP→ISPL is internal and the destination is always ISPL's location.
  const destination = isASP
    ? (cfg.dispatch_destination || '')
    : (buyer.cpla || buyer.pla || cfg.dispatch_destination || '');
  labeledCell(rightX + rCell, ry2, rCell, rSmall, 'Destination', destination);
  ry2 += rSmall;

  // Dispatch From block (sister company) — fills remaining middle height.
  // Hidden when flag_dispatch is OFF (item 2) — draw an empty bordered cell so
  // the right-column visual frame still matches the left column's height.
  const dispatchFromH = midH - rSmall;
  if (showDispatch) {
    box(rightX, ry2, rightW, dispatchFromH);
    const dispatchY = ry2;
    doc.font('Helvetica-Bold').fontSize(8).text('Dispatch From:', rightX + 3, dispatchY + 4, { width: rightW - 6 });
    doc.font('Helvetica-Bold').fontSize(9).text(cfg.s_company || 'AMAZING SPICE PARK PRIVATE LIMITED', rightX + 3, dispatchY + 14, { width: rightW - 6 });
    doc.font('Helvetica').fontSize(8);
    // Advance dy by actual rendered height so wrapped text doesn't overlap next line.
    let dy = dispatchY + 26;
    const dispW = rightW - 6;
    const writeLine = (txt) => {
      if (!txt) return;
      doc.text(txt, rightX + 3, dy, { width: dispW });
      dy += doc.heightOfString(txt, { width: dispW }) + 1;
    };
    writeLine(cfg.s_address1);
    writeLine(cfg.s_address2);
    if (cfg.s_state) writeLine(`${cfg.s_state} Code:${cfg.s_st_code || '32'}`);
    if (cfg.s_gstin) writeLine(`GSTIN.${cfg.s_gstin}`);
  } else {
    // Empty bordered cell to keep the frame consistent with the left column
    box(rightX, ry2, rightW, dispatchFromH);
  }

  y = midY + midH;

  // ── LINE ITEMS TABLE ────────────────────────────────────────
  // Columns: Sl | Token | Bags | Description | HSN/SAC | Shipped | Billed | Rate | per | Amount
  // Honor flag_hsn: when OFF, drop the HSN/SAC column entirely. The
  // Description column fills the freed space (it's the remainder column).
  const showHsn = readFlag(cfg.flag_hsn, true);

  // Amount is now fixed (~95pt) — enough for values up to "1,00,00,000.00" (1 crore).
  // Extra width goes to Description instead of Amount.
  const colW = {
    sl:    22,
    token: 40,
    bags:  34,
    desc:  0,    // fills remainder
    hsn:   showHsn ? 52 : 0,
    shipped: 60,
    billed:  60,
    rate:  48,
    per:   20,
    amount: 95,
  };
  const fixedSum = Object.values(colW).reduce((a, b) => a + b, 0);
  colW.desc = W - fixedSum;

  const cols = showHsn
    ? ['sl', 'token', 'bags', 'desc', 'hsn', 'shipped', 'billed', 'rate', 'per', 'amount']
    : ['sl', 'token', 'bags', 'desc',        'shipped', 'billed', 'rate', 'per', 'amount'];
  function colX(key) {
    let cx = x0;
    for (const k of cols) { if (k === key) return cx; cx += colW[k]; }
    return cx;
  }

  const hdrH = 24;
  const rowH = 14;
  // Reserve space at bottom of every page for a "continued" notice + margin
  const bottomReserve = 30;
  const pageBottom = pageH - margin - bottomReserve;

  // Draw the line-items table header at the current y.
  // Extracted into a function so it can be re-drawn at the top of each new page.
  function drawTableHeader() {
    box(x0, y, W, hdrH);
    for (const k of cols) {
      const cx = colX(k);
      if (cx <= x0) continue;
      if (k === 'billed') {
        doc.moveTo(cx, y + 12).lineTo(cx, y + hdrH).stroke();
      } else {
        doc.moveTo(cx, y).lineTo(cx, y + hdrH).stroke();
      }
    }
    const qtyX = colX('shipped');
    const qtyW = colW.shipped + colW.billed;
    doc.moveTo(qtyX, y + 12).lineTo(qtyX + qtyW, y + 12).stroke();
    doc.font('Helvetica').fontSize(7);
    doc.text('Quantity', qtyX, y + 3, { width: qtyW, align: 'center' });
    doc.text('Shipped', qtyX, y + 14, { width: colW.shipped, align: 'center' });
    doc.text('Billed', qtyX + colW.shipped, y + 14, { width: colW.billed, align: 'center' });
    const hdr = {
      sl:    ['SI', 'No.'],
      token: ['Lot', 'No'],
      bags:  ['No. of', 'Bags'],
      desc:  ['Description of Goods', ''],
      hsn:   ['HSN/SAC', ''],
      rate:  ['Rate', ''],
      per:   ['per', ''],
      amount: ['Amount', ''],
    };
    // Only iterate columns that are actually rendered (skips hsn when off)
    for (const k of cols) {
      if (!hdr[k]) continue;  // skip shipped/billed (handled separately above)
      const [l1, l2] = hdr[k];
      const cx = colX(k);
      doc.text(l1, cx + 2, y + 3, { width: colW[k] - 4, align: 'center' });
      if (l2) doc.text(l2, cx + 2, y + 13, { width: colW[k] - 4, align: 'center' });
    }
    y += hdrH;
  }

  // If the next row wouldn't fit on the current page, close the current table,
  // emit a new page, and redraw the table header at the top.
  function ensureRoomFor(neededH) {
    if (y + neededH <= pageBottom) return;
    // Close the table with a horizontal line so the last row has a bottom border
    doc.lineWidth(0.5).moveTo(x0, y).lineTo(x0 + W, y).stroke();
    // Bottom notice: "Continued..." right-aligned
    doc.font('Helvetica-Oblique').fontSize(7)
       .text('Continued on next page...', x0, y + 4, { width: W, align: 'right' });
    doc.font('Helvetica');
    // New page
    doc.addPage({ size: 'A4', margin: margin });
    y = margin;
    // Small top note that this is a continuation
    doc.font('Helvetica-Oblique').fontSize(7)
       .text(`Tax Invoice — ${formatInvoiceNo(primaryCfg, saleType, invoiceNo)} (continued)`,
             x0, y, { width: W, align: 'center' });
    doc.font('Helvetica');
    y += 12;
    // Redraw the column header
    drawTableHeader();
  }

  // Draw initial table header
  drawTableHeader();

  // Line item rows
  doc.font('Helvetica').fontSize(8);
  const hsnCardamom = cfg.hsn_cardamom || '09083120';
  const hsnGunny    = cfg.hsn_gunny    || '63051040';

  let sl = 1;
  for (const li of lineItems) {
    ensureRoomFor(rowH);
    stripeFill(y, rowH, sl - 1);
    rowVerticals(y, rowH);
    doc.text(String(sl), colX('sl') + 2, y + 3, { width: colW.sl - 4, align: 'center' });
    doc.text(String(li.lot || ''), colX('token') + 2, y + 3, { width: colW.token - 4, align: 'center' });
    doc.text(String(li.bags || ''), colX('bags') + 2, y + 3, { width: colW.bags - 4, align: 'center' });
    doc.font('Helvetica-Bold').text('Cardamom', colX('desc') + 4, y + 3, { width: colW.desc - 8 });
    doc.font('Helvetica');
    if (showHsn) doc.text(hsnCardamom, colX('hsn') + 2, y + 3, { width: colW.hsn - 4, align: 'center' });
    doc.text(`${li.qty.toFixed(3)} Kgs.`, colX('shipped') + 2, y + 3, { width: colW.shipped - 4, align: 'right' , lineBreak: false});
    doc.font('Helvetica-Bold').text(`${li.qty.toFixed(3)} Kgs.`, colX('billed') + 2, y + 3, { width: colW.billed - 4, align: 'right' , lineBreak: false});
    // ASP invoices: bill at P_Rate and show PurAmt (ASP→ISP internal transfer price)
    // ISP invoices: bill at Price and show Amount (external customer price)
    const lineRate = isASP ? (li.prate != null ? li.prate : li.price) : li.price;
    const lineAmount = isASP ? (li.puramt != null ? li.puramt : li.amount) : li.amount;
    doc.font('Helvetica-Bold').text(formatINR(lineRate), colX('rate') + 2, y + 3, { width: colW.rate - 4, align: 'right' });
    doc.font('Helvetica').text('Kgs.', colX('per') + 2, y + 3, { width: colW.per - 4, align: 'center' });
    doc.font('Helvetica-Bold').text(formatINR(lineAmount), colX('amount') + 2, y + 3, { width: colW.amount - 4, align: 'right' });
    doc.font('Helvetica');
    y += rowH;
    sl++;
  }

  // ── Adaptive empty-row padding ──
  // For small invoices, pad blank striped rows after the data so the
  // summary/total/signature block lands near the bottom of A4. Only
  // applied on a single-page invoice (won't pad if data already pushed
  // us past where padding would help). The bottom reserve below covers
  // every optional summary row + amount-words + HSN summary + signature.
  const gunnyH_pre     = (summary.totalBags > 0 && summary.gunnyCost > 0) ? rowH : 0;
  const transportH_pre = (summary.transportCost > 0 && !hideTransportInsurance) ? rowH : 0;
  const insuranceH_pre = (summary.insuranceCost > 0 && !hideTransportInsurance) ? rowH : 0;
  const gstRowCount_pre = summary.isInterState ? 1 : 2;
  const addlH_pre      = (summary.addlCharge && summary.addlCharge !== 0) ? rowH : 0;
  // Required height for the summary + bottom blocks (matches the existing
  // ensureRoomFor estimate below)
  const requiredBottomH =
      gunnyH_pre + transportH_pre + insuranceH_pre
    + rowH + (rowH * gstRowCount_pre) + rowH + addlH_pre + (rowH + 2)
    + 24 + 90 + 16 + 90 + 10;
  // Only pad if we're still on the first page and have room for it.
  // pageBottom is the y at which content should stop on this page.
  const targetSummaryStartY = pageBottom - requiredBottomH;
  if (y < targetSummaryStartY) {
    const fillH = targetSummaryStartY - y;
    const fillRows = Math.floor(fillH / rowH);
    for (let i = 0; i < fillRows; i++) {
      stripeFill(y, rowH, sl - 1);
      rowVerticals(y, rowH);
      // Empty cells — vertical separators are drawn by rowVerticals above,
      // and the stripeFill provides the visual band for alternating rows.
      y += rowH;
      sl++;
    }
  }

  // Before drawing summary rows: estimate the total remaining height and
  // push to next page if it won't fit. Keeps the summary block intact.
  //   Gunny row:    rowH (if applicable)
  //   Transport:    rowH (if applicable)
  //   Insurance:    rowH (if applicable)
  //   Subtotal:     rowH
  //   GST rows:     rowH × (isInterState ? 1 : 2)
  //   Round-off:    rowH
  //   Total:        rowH + 2
  //   Amount words: 24
  //   HSN summary:  ~20 header + 14×up-to-4 rows + 14 total = ~90
  //   Tax words:    16
  //   Bank/Sig:     90
  const gunnyH     = (summary.totalBags > 0 && summary.gunnyCost > 0) ? rowH : 0;
  const transportH = (summary.transportCost > 0 && !hideTransportInsurance) ? rowH : 0;
  const insuranceH = (summary.insuranceCost > 0 && !hideTransportInsurance) ? rowH : 0;
  const gstRowCount = summary.isInterState ? 1 : 2;
  const addlH = (summary.addlCharge && summary.addlCharge !== 0) ? rowH : 0;
  const summaryBlockH = gunnyH + transportH + insuranceH
                      + rowH + (rowH * gstRowCount) + rowH + addlH + (rowH + 2)
                      + 24 + 90 + 16 + 90 + 10; // +10 buffer
  ensureRoomFor(summaryBlockH);

  // Gunny row
  if (summary.totalBags > 0 && summary.gunnyCost > 0) {
    stripeFill(y, rowH, sl - 1);
    rowVerticals(y, rowH);
    // sl.no shown only on lot rows (not on Gunny/Transport/Insurance footer rows)
    doc.font('Helvetica-Bold').text('Gunny', colX('desc') + 4, y + 3, { width: colW.desc - 8 });
    doc.font('Helvetica');
    if (showHsn) doc.text(hsnGunny, colX('hsn') + 2, y + 3, { width: colW.hsn - 4, align: 'center' });
    doc.text(`${summary.totalBags} Nos.`, colX('shipped') + 2, y + 3, { width: colW.shipped - 4, align: 'right' });
    const gunnyRate = (cfg.gunny_rate || 165).toFixed(2);
    doc.font('Helvetica-Bold').text(gunnyRate, colX('rate') + 2, y + 3, { width: colW.rate - 4, align: 'right' });
    doc.font('Helvetica').text('Nos.', colX('per') + 2, y + 3, { width: colW.per - 4, align: 'center' });
    doc.font('Helvetica-Bold').text(formatINR(summary.gunnyCost), colX('amount') + 2, y + 3, { width: colW.amount - 4, align: 'right' });
    doc.font('Helvetica');
    y += rowH;
    sl++;
  }

  // Transport row (SAC: transport service)
  // Rate depends on sale type: L → local_transport, else → transport
  // Use pickRate (not `||`) so that 0 is respected as an explicit user value.
  const pickRate = (...vals) => {
    for (const v of vals) {
      if (v === undefined || v === null || v === '') continue;
      const n = typeof v === 'number' ? v : parseFloat(v);
      if (!Number.isNaN(n)) return n;
    }
    return 0;
  };
  const isLocalSale = (saleType === 'L');
  const transportRate = isLocalSale
    ? pickRate(cfg.local_transport, cfg.transport, 2.5)
    : pickRate(cfg.transport, 2.5);
  // ASP invoices never bill Transport/Insurance separately (item 5).
  // Wrap both rows + their HSN summary entries in an isASP guard.
  const sacTransport = cfg.sac_transport || '996791';
  if (summary.transportCost > 0 && !hideTransportInsurance) {
    stripeFill(y, rowH, sl - 1);
    rowVerticals(y, rowH);
    // sl.no shown only on lot rows (not on Gunny/Transport/Insurance footer rows)
    doc.font('Helvetica-Bold').text('Transport', colX('desc') + 4, y + 3, { width: colW.desc - 8 });
    doc.font('Helvetica');
    if (showHsn) doc.text(sacTransport, colX('hsn') + 2, y + 3, { width: colW.hsn - 4, align: 'center' });
    doc.text(`${summary.totalQty.toFixed(3)} Kgs.`, colX('shipped') + 2, y + 3, { width: colW.shipped - 4, align: 'right' , lineBreak: false});
    doc.font('Helvetica-Bold').text(transportRate.toFixed(2), colX('rate') + 2, y + 3, { width: colW.rate - 4, align: 'right' });
    doc.font('Helvetica').text('Kgs.', colX('per') + 2, y + 3, { width: colW.per - 4, align: 'center' });
    doc.font('Helvetica-Bold').text(formatINR(summary.transportCost), colX('amount') + 2, y + 3, { width: colW.amount - 4, align: 'right' });
    doc.font('Helvetica');
    y += rowH;
    sl++;
  }

  // Insurance row (SAC: insurance service)
  // Amount = ((cardamom + gunny) + GST on them) / 1000 × insurance_rate
  // Rate depends on sale type: L → local_insurance, else → insurance
  const insuranceRate = isLocalSale
    ? pickRate(cfg.local_insurance, cfg.insurance, 0.75)
    : pickRate(cfg.insurance, 0.75);
  const sacInsurance = cfg.sac_insurance || '997136';
  if (summary.insuranceCost > 0 && !hideTransportInsurance) {
    stripeFill(y, rowH, sl - 1);
    rowVerticals(y, rowH);
    // sl.no shown only on lot rows (not on Gunny/Transport/Insurance footer rows)
    doc.font('Helvetica-Bold').text('Insurance', colX('desc') + 4, y + 3, { width: colW.desc - 8 });
    doc.font('Helvetica');
    if (showHsn) doc.text(sacInsurance, colX('hsn') + 2, y + 3, { width: colW.hsn - 4, align: 'center' });
    doc.font('Helvetica-Bold').text(insuranceRate.toFixed(2), colX('rate') + 2, y + 3, { width: colW.rate - 4, align: 'right' });
    doc.font('Helvetica-Bold').text(formatINR(summary.insuranceCost), colX('amount') + 2, y + 3, { width: colW.amount - 4, align: 'right' });
    doc.font('Helvetica');
    y += rowH;
    sl++;
  }

  // Horizontal line under the last line-item's Amount cell — visually groups
  // the line-items (Cardamom/Gunny/Transport/Insurance) and separates them
  // from the subtotal that follows. Drawn ONLY in the Amount column so it
  // doesn't disrupt the description/HSN/Qty cells to the left.
  doc.lineWidth(0.5).moveTo(colX('amount'), y).lineTo(colX('amount') + colW.amount, y).stroke();

  // Subtotal row = taxable value (cardamom + gunny + transport + insurance)
  const subtotal = summary.taxableValue;
  rowVerticals(y, rowH);
  doc.font('Helvetica-Bold').text(formatINR(subtotal), colX('amount') + 2, y + 3, { width: colW.amount - 4, align: 'right' });
  y += rowH;

  // GST rows — separate row per tax component
  const gstGoods = cfg.gst_goods || 5;
  const gstRate = gstGoods / 2;

  function drawTaxRow(label, amount) {
    rowVerticals(y, rowH);
    doc.font('Helvetica-BoldOblique').text(label, colX('desc') + 4, y + 3, { width: colW.desc + colW.hsn + colW.shipped + colW.billed - 8 });
    doc.font('Helvetica-Bold').text(formatINR(amount), colX('amount') + 2, y + 3, { width: colW.amount - 4, align: 'right' });
    doc.font('Helvetica');
    y += rowH;
  }

  // GST row label prefix: "OUTPUT" for sales (tax collected by seller),
  // "INPUT" for the purchase view (tax paid by buyer — eligible as input tax credit)
  const gstDir = isPurchaseView ? 'INPUT' : 'OUTPUT';
  if (summary.isInterState) {
    drawTaxRow(`${gstDir} IGST ${gstGoods}%`, summary.igst);
  } else {
    drawTaxRow(`${gstDir} CGST ${gstRate}%`, summary.cgst);
    drawTaxRow(`${gstDir} SGST ${gstRate}%`, summary.sgst);
  }

  // Round on/off row
  rowVerticals(y, rowH);
  doc.font('Helvetica-BoldOblique').text('Round On/off', colX('desc') + 4, y + 3, { width: colW.desc + colW.hsn - 8 });
  doc.font('Helvetica-Bold').text(formatINR(summary.roundDiff), colX('amount') + 2, y + 3, { width: colW.amount - 4, align: 'right' });
  doc.font('Helvetica');
  y += rowH;

  // Additional Charge row — name + value from Settings → Rates & Charges.
  // Computed = sum(cardamom) × cfg.addl_charge_value %. Sits below Round and
  // adds onto the grand total. Hidden when zero.
  if (summary.addlCharge && summary.addlCharge !== 0) {
    rowVerticals(y, rowH);
    const addlLabel = (summary.addlChargeName || cfg.addl_charge_name || 'Additional Charge').toString();
    doc.font('Helvetica-Bold').text(addlLabel, colX('desc') + 4, y + 3, { width: colW.desc + colW.hsn - 8 });
    doc.font('Helvetica-Bold').text(formatINR(summary.addlCharge), colX('amount') + 2, y + 3, { width: colW.amount - 4, align: 'right' });
    doc.font('Helvetica');
    y += rowH;
  }

  // Total row (bold) — shipped|billed divider skipped so total qty spans both
  // Draw a horizontal line above Total to visually separate it from the summary rows.
  doc.lineWidth(0.5).moveTo(x0, y).lineTo(x0 + W, y).stroke();
  const totalRowH = rowH + 2;
  rowVerticals(y, totalRowH, /*skipBilledSplit*/ true);
  doc.font('Helvetica-Bold').fontSize(8);
  doc.text(String(summary.totalBags), colX('bags') + 2, y + 4, { width: colW.bags - 4, align: 'center' });
  doc.text('Total', colX('desc') + 4, y + 4, { width: colW.desc - 8 });
  doc.text(`${summary.totalQty.toFixed(3)} Kgs.`, colX('shipped') + 2, y + 4, { width: colW.shipped + colW.billed - 4, align: 'right' , lineBreak: false});
  // Grand total: right-align with extra padding so it doesn't touch the border
  // (No ₹ symbol — PDFKit's Helvetica doesn't support U+20B9 and renders garbage)
  doc.fontSize(9).text(formatINR(summary.grandTotal), colX('amount') + 2, y + 4, { width: colW.amount - 6, align: 'right' });
  doc.fontSize(8);
  y += totalRowH;
  // Close the line-items table with a single horizontal line at the bottom
  doc.lineWidth(0.5).moveTo(x0, y).lineTo(x0 + W, y).stroke();

  // ── AMOUNT IN WORDS ─────────────────────────────────────────
  const amtWordsH = 24;
  box(x0, y, W, amtWordsH);
  doc.font('Helvetica').fontSize(7).text('Amount Chargeable (in words)', x0 + 3, y + 2);
  doc.font('Helvetica').fontSize(7).text('E. & O.E', x0, y + 2, { width: W - 4, align: 'right' });
  doc.font('Helvetica-Bold').fontSize(9).text(`INR ${amountToWords(summary.grandTotal)} Only`, x0 + 3, y + 11, { width: W - 6 });
  y += amtWordsH;

  // ── HSN SUMMARY TABLE ───────────────────────────────────────
  // When flag_hsn is OFF, the HSN/SAC column is hidden but the rest of
  // the summary (Taxable Value + GST breakdown) stays — that's the only
  // place where Taxable Value is totaled, so it must not vanish.
  // Columns: HSN/SAC | Taxable Value | IGST Rate | IGST Amount | Total Tax Amount
  //   OR: HSN/SAC | Taxable Value | CGST Rate | CGST Amt | SGST Rate | SGST Amt | Total Tax Amount
  const isInter = summary.isInterState;
  const hsnRows = [];
  // Helper: build one HSN summary row for a taxable amount at gstGoods rate.
  // `desc` is the human-readable label (Cardamom / Gunny / Transport /
  // Insurance) — used as the row label when flag_hsn is OFF (since we
  // don't show the HSN code, we want to know WHAT each row is).
  function hsnRow(hsn, desc, taxable) {
    return {
      hsn, desc, taxable,
      rate: gstGoods,
      cgst: isInter ? 0 : +(taxable * gstGoods / 2 / 100).toFixed(2),
      sgst: isInter ? 0 : +(taxable * gstGoods / 2 / 100).toFixed(2),
      igst: isInter ? +(taxable * gstGoods / 100).toFixed(2) : 0,
    };
  }
  hsnRows.push(hsnRow(hsnCardamom, 'Cardamom', summary.totalAmount));
  if (summary.gunnyCost > 0)     hsnRows.push(hsnRow(hsnGunny, 'Gunny', summary.gunnyCost));
  // ASP invoices: Transport/Insurance are not billed, so skip them from HSN summary too
  if (summary.transportCost > 0 && !hideTransportInsurance) hsnRows.push(hsnRow(sacTransport, 'Transport', summary.transportCost));
  if (summary.insuranceCost > 0 && !hideTransportInsurance) hsnRows.push(hsnRow(sacInsurance, 'Insurance', summary.insuranceCost));

  const hsnHdrH = 20;
  const hsnRowH = 12;
  // The first column shows EITHER HSN/SAC (when flag_hsn ON) or
  // "Description of Goods" with the per-row description (when OFF).
  // We keep the same column WIDTH either way so the rest of the layout
  // is identical between flag states — only the label and cell content
  // change.
  const firstColW = isInter ? 160 : 130;
  const hsnCols = isInter
    ? { hsn: firstColW, taxable: 90, rateLbl: 'IGST', rate: 60, amt: 90, total: 0 }
    : { hsn: firstColW, taxable: 80, cgstRate: 40, cgstAmt: 60, sgstRate: 40, sgstAmt: 60, total: 0 };
  const hsnFixed = Object.entries(hsnCols).filter(([k]) => typeof hsnCols[k] === 'number').reduce((a, [, v]) => a + v, 0);
  hsnCols.total = W - hsnFixed;

  // Header
  box(x0, y, W, hsnHdrH);
  // First column label: HSN/SAC (default) or "Description of Goods" when
  // flag_hsn is OFF — column width is identical either way.
  const firstColLabel = showHsn ? 'HSN/SAC' : 'Description of Goods';
  if (isInter) {
    let cx = x0;
    doc.font('Helvetica').fontSize(7);
    doc.text(firstColLabel, cx + 2, y + 6, { width: hsnCols.hsn - 4, align: 'center' });
    doc.moveTo(cx + hsnCols.hsn, y).lineTo(cx + hsnCols.hsn, y + hsnHdrH).stroke();
    cx += hsnCols.hsn;
    doc.text('Taxable', cx + 2, y + 2, { width: hsnCols.taxable - 4, align: 'center' });
    doc.text('Value', cx + 2, y + 11, { width: hsnCols.taxable - 4, align: 'center' });
    doc.moveTo(cx + hsnCols.taxable, y).lineTo(cx + hsnCols.taxable, y + hsnHdrH).stroke();
    cx += hsnCols.taxable;
    const igstW = hsnCols.rate + hsnCols.amt;
    doc.text('IGST', cx + 2, y + 2, { width: igstW - 4, align: 'center' });
    doc.moveTo(cx, y + 10).lineTo(cx + igstW, y + 10).stroke();
    doc.text('Rate', cx + 2, y + 12, { width: hsnCols.rate - 4, align: 'center' });
    doc.text('Amount', cx + hsnCols.rate + 2, y + 12, { width: hsnCols.amt - 4, align: 'center' });
    doc.moveTo(cx + hsnCols.rate, y + 10).lineTo(cx + hsnCols.rate, y + hsnHdrH).stroke();
    doc.moveTo(cx + igstW, y).lineTo(cx + igstW, y + hsnHdrH).stroke();
    cx += igstW;
    doc.text('Total', cx + 2, y + 2, { width: hsnCols.total - 4, align: 'center' });
    doc.text('Tax Amount', cx + 2, y + 11, { width: hsnCols.total - 4, align: 'center' });
  } else {
    let cx = x0;
    doc.font('Helvetica').fontSize(7);
    doc.text(firstColLabel, cx + 2, y + 6, { width: hsnCols.hsn - 4, align: 'center' });
    doc.moveTo(cx + hsnCols.hsn, y).lineTo(cx + hsnCols.hsn, y + hsnHdrH).stroke();
    cx += hsnCols.hsn;
    doc.text('Taxable Value', cx + 2, y + 6, { width: hsnCols.taxable - 4, align: 'center' });
    doc.moveTo(cx + hsnCols.taxable, y).lineTo(cx + hsnCols.taxable, y + hsnHdrH).stroke();
    cx += hsnCols.taxable;
    const cgstW = hsnCols.cgstRate + hsnCols.cgstAmt;
    doc.text('CGST', cx + 2, y + 2, { width: cgstW - 4, align: 'center' });
    doc.moveTo(cx, y + 10).lineTo(cx + cgstW, y + 10).stroke();
    doc.text('Rate', cx + 2, y + 12, { width: hsnCols.cgstRate - 4, align: 'center' });
    doc.text('Amount', cx + hsnCols.cgstRate + 2, y + 12, { width: hsnCols.cgstAmt - 4, align: 'center' });
    doc.moveTo(cx + hsnCols.cgstRate, y + 10).lineTo(cx + hsnCols.cgstRate, y + hsnHdrH).stroke();
    doc.moveTo(cx + cgstW, y).lineTo(cx + cgstW, y + hsnHdrH).stroke();
    cx += cgstW;
    const sgstW = hsnCols.sgstRate + hsnCols.sgstAmt;
    doc.text('SGST', cx + 2, y + 2, { width: sgstW - 4, align: 'center' });
    doc.moveTo(cx, y + 10).lineTo(cx + sgstW, y + 10).stroke();
    doc.text('Rate', cx + 2, y + 12, { width: hsnCols.sgstRate - 4, align: 'center' });
    doc.text('Amount', cx + hsnCols.sgstRate + 2, y + 12, { width: hsnCols.sgstAmt - 4, align: 'center' });
    doc.moveTo(cx + hsnCols.sgstRate, y + 10).lineTo(cx + hsnCols.sgstRate, y + hsnHdrH).stroke();
    doc.moveTo(cx + sgstW, y).lineTo(cx + sgstW, y + hsnHdrH).stroke();
    cx += sgstW;
    doc.text('Total Tax Amount', cx + 2, y + 6, { width: hsnCols.total - 4, align: 'center' });
  }
  y += hsnHdrH;

  // HSN rows — vertical-only dividers + alternate-row stripe
  // Builds column x-positions once per call to avoid repeating arithmetic.
  function hsnColBoundaries() {
    // Returns [x0, afterFirstCol, afterTaxable, ..., x0+W]
    // First col is HSN/SAC or Description of Goods — same width either way.
    const xs = [x0];
    let cx = x0 + hsnCols.hsn;  xs.push(cx);
    cx += hsnCols.taxable;      xs.push(cx);
    if (isInter) {
      cx += hsnCols.rate;        xs.push(cx);
      cx += hsnCols.amt;         xs.push(cx);
    } else {
      cx += hsnCols.cgstRate;    xs.push(cx);
      cx += hsnCols.cgstAmt;     xs.push(cx);
      cx += hsnCols.sgstRate;    xs.push(cx);
      cx += hsnCols.sgstAmt;     xs.push(cx);
    }
    xs.push(x0 + W); // right edge (Total Tax Amount)
    return xs;
  }
  const hsnXs = hsnColBoundaries();

  function hsnRowVerticals(ry, rh) {
    doc.lineWidth(0.5);
    for (const hx of hsnXs) {
      doc.moveTo(hx, ry).lineTo(hx, ry + rh).stroke();
    }
  }

  function drawHsnRow(row, isTotal, rowIndex) {
    // Stripe first, then verticals, then text
    stripeFill(y, hsnRowH, rowIndex);
    hsnRowVerticals(y, hsnRowH);
    let cx = x0;
    doc.font(isTotal ? 'Helvetica-Bold' : 'Helvetica').fontSize(8);
    // First column: HSN/SAC code when flag on, item description when off.
    // For Total row both show "Total" (set in caller).
    const firstColVal = showHsn ? row.hsn : (row.desc || row.hsn);
    doc.text(firstColVal, cx + 2, y + 2, { width: hsnCols.hsn - 4 });
    cx += hsnCols.hsn;
    doc.text(formatINR(row.taxable), cx + 2, y + 2, { width: hsnCols.taxable - 4, align: 'right' });
    cx += hsnCols.taxable;
    if (isInter) {
      doc.text(`${row.rate}%`, cx + 2, y + 2, { width: hsnCols.rate - 4, align: 'center' });
      doc.text(formatINR(row.igst), cx + hsnCols.rate + 2, y + 2, { width: hsnCols.amt - 4, align: 'right' });
      cx += hsnCols.rate + hsnCols.amt;
      doc.text(formatINR(row.igst), cx + 2, y + 2, { width: hsnCols.total - 4, align: 'right' });
    } else {
      doc.text(`${row.rate / 2}%`, cx + 2, y + 2, { width: hsnCols.cgstRate - 4, align: 'center' });
      doc.text(formatINR(row.cgst), cx + hsnCols.cgstRate + 2, y + 2, { width: hsnCols.cgstAmt - 4, align: 'right' });
      cx += hsnCols.cgstRate + hsnCols.cgstAmt;
      doc.text(`${row.rate / 2}%`, cx + 2, y + 2, { width: hsnCols.sgstRate - 4, align: 'center' });
      doc.text(formatINR(row.sgst), cx + hsnCols.sgstRate + 2, y + 2, { width: hsnCols.sgstAmt - 4, align: 'right' });
      cx += hsnCols.sgstRate + hsnCols.sgstAmt;
      doc.text(formatINR(row.cgst + row.sgst), cx + 2, y + 2, { width: hsnCols.total - 4, align: 'right' });
    }
    doc.font('Helvetica');
    y += hsnRowH;
  }

  for (let i = 0; i < hsnRows.length; i++) drawHsnRow(hsnRows[i], false, i);
  // Horizontal line above Total row — visual separator between data rows and total
  doc.lineWidth(0.5).moveTo(x0, y).lineTo(x0 + W, y).stroke();
  // Total row — first column reads "Total" regardless of flag_hsn state
  const totRow = {
    hsn: 'Total',
    desc: 'Total',
    taxable: hsnRows.reduce((a, r) => a + r.taxable, 0),
    rate: gstGoods,
    cgst: hsnRows.reduce((a, r) => a + r.cgst, 0),
    sgst: hsnRows.reduce((a, r) => a + r.sgst, 0),
    igst: hsnRows.reduce((a, r) => a + r.igst, 0),
  };
  drawHsnRow(totRow, true, hsnRows.length);
  // Close the HSN summary with a horizontal line so the last row has a bottom border
  doc.lineWidth(0.5).moveTo(x0, y).lineTo(x0 + W, y).stroke();

  // ── TAX IN WORDS ────────────────────────────────────────────
  const taxAmount = summary.cgst + summary.sgst + summary.igst;
  const taxWordsH = 16;
  box(x0, y, W, taxWordsH);
  doc.font('Helvetica').fontSize(7).text('Tax Amount (in words) :', x0 + 3, y + 4);
  doc.font('Helvetica-Bold').fontSize(9).text(`INR ${amountToWords(taxAmount)} Only`, x0 + 110, y + 3, { width: W - 120 });
  y += taxWordsH;

  // ── BANK + SIGNATURE BLOCK ──────────────────────────────────
  const footerH = 90;
  box(x0, y, W, footerH);
  doc.moveTo(x0 + leftW, y).lineTo(x0 + leftW, y + footerH).stroke();

  // Left: Company PAN + Declaration
  let fy = y + 6;
  doc.font('Helvetica').fontSize(8);
  doc.text(`Company's PAN     : `, x0 + 4, fy, { continued: true }).font('Helvetica-Bold').text(co.pan || '');
  fy += 16;
  doc.font('Helvetica').fontSize(8).text('Declaration', x0 + 4, fy); fy += 11;
  doc.fontSize(7).text('We declare that this invoice shows the actual price of the goods', x0 + 4, fy, { width: leftW - 8 }); fy += 10;
  doc.text('described and that all particulars are true and correct.', x0 + 4, fy, { width: leftW - 8 });

  // Right: Bank details (top), "for COMPANY" (middle), Authorised Signatory (bottom-right)
  let bky = y + 6;
  const bkX = rightX + 4;
  const bkInnerW = rightW - 8;
  // Honor flag_bank — when OFF, skip the bank details block entirely
  // ("for COMPANY NAME" + Authorised Signatory remain since they're
  // independent of bank info).
  const showBank = readFlag(cfg.flag_bank, true);
  if (showBank) {
    doc.font('Helvetica').fontSize(8).text("Company's Bank Details", bkX, bky); bky += 11;
    // Bank picks:
    //   Purchase view         → KL bank (ASP is still the selling company, so its bank receives payment)
    //   ASP sales invoice     → KL bank
    //   ISP invoices          → TN bank
    const useKLBank = isASP; // same for both sales AND purchase view when isASP
    const bankName = useKLBank ? (cfg.bank_kl_name || '') : (cfg.bank_tn_name || '');
    // Account no. + IFSC honour the display masking policy (mask_acct /
    // mask_ifsc). Defaults: acct → last4, ifsc → none, so by default only the
    // account is partially masked and the IFSC prints in full.
    const { maskAcct, maskIfsc } = makeMaskers(cfg);
    const bankAcct = maskAcct(useKLBank ? (cfg.bank_kl_acct || '') : (cfg.bank_tn_acct || ''));
    const bankIfsc = maskIfsc(useKLBank ? (cfg.bank_kl_ifsc || '') : (cfg.bank_tn_ifsc || ''));
    // Align values at a fixed x so all three rows start at the same column
    const labelW = 90;
    const valX = bkX + labelW;
    const valW = bkInnerW - labelW;
    doc.font('Helvetica').fontSize(8).text('Bank Name', bkX, bky, { width: labelW });
    doc.font('Helvetica').text(':', bkX + labelW - 8, bky);
    doc.font('Helvetica-Bold').text(bankName, valX, bky, { width: valW });
    bky += 10;
    doc.font('Helvetica').text('A/c No.', bkX, bky, { width: labelW });
    doc.font('Helvetica').text(':', bkX + labelW - 8, bky);
    doc.font('Helvetica-Bold').text(bankAcct, valX, bky, { width: valW });
    bky += 10;
    doc.font('Helvetica').text('Branch & IFS Code', bkX, bky, { width: labelW });
    doc.font('Helvetica').text(':', bkX + labelW - 8, bky);
    const branchLbl = (bankName.split('-')[1] || bankName).trim();
    doc.font('Helvetica-Bold').text(`${branchLbl} & ${bankIfsc}`, valX, bky, { width: valW });
    bky += 14;
    // Horizontal line above "for COMPANY NAME" — separates bank details from signatory section.
    // Drawn only in the right cell (bank details column), not across the Declaration block.
    doc.lineWidth(0.5).moveTo(rightX, bky - 2).lineTo(x0 + W, bky - 2).stroke();
    bky += 2;
  } // end if (showBank)
  // "for COMPANY NAME" right-aligned
  // In purchase view, the issuer header is ISPL but the signatory block
  // represents the SELLING company (ASP) — whose bank received the payment
  // and whose authorised signatory certifies the sale.
  const forCompanyName = isPurchaseView
    ? (cfg.s_company || 'AMAZING SPICE PARK PRIVATE LIMITED')
    : (co.name || '');
  doc.font('Helvetica-Bold').fontSize(8).text(`for ${forCompanyName}`, bkX, bky, { width: bkInnerW, align: 'right' });
  // Authorised Signatory at bottom-right of footer
  doc.font('Helvetica').fontSize(8).text('Authorised Signatory', bkX, y + footerH - 12, { width: bkInnerW, align: 'right' });

  y += footerH;

  // Batch mode: caller owns the doc lifecycle; just return (no buffer).
  if (isBatch) return null;

  return new Promise((resolve) => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

/**
 * Agriculturist Bill of Supply PDF (GSTKBILP.PRG / GSTBILP.PRG equivalent)
 * For non-GSTIN sellers — no GST charged
 */
// Agri Bill (Bill of Supply for agriculturist sellers) — layout matches
// PLANTER-TF-03-1-2.pdf reference. Key differences from a tax invoice:
//   - Title: "Bill for Purchase from Agriculturist" with Section 2(7) subtitle
//   - Issuer block = ASP (sister company) — agri purchases are made via ASP
//   - Buyer/Consignee columns: seller is "DETAILS OF SELLER [BILLED FOR]"
//     and the consignee column is typically empty for self-generated bills
//   - GST columns are PRESENT but left BLANK (agriculturist is unregistered
//     under GST Sec 2(7) — no tax liability on the purchase)
//   - Bottom-right totals block: Taxable / Integrated / Central / State /
//     Round / Total Value
//   - DC No. / Date: line at bottom-left, matches reference positioning
function generateAgriBillPDF(billData, cfg, billNo, externalDoc) {
  const isBatch = !!externalDoc;
  let doc, buffers;
  if (isBatch) {
    doc = externalDoc;
    if (doc._billCount && doc._billCount > 0) doc.addPage({ size: 'A4', margin: 20 });
    doc._billCount = (doc._billCount || 0) + 1;
  } else {
    doc = new PDFDocument({ size: 'A4', margin: 20 });
    buffers = [];
    doc.on('data', b => buffers.push(b));
  }

  const PAGE_W = doc.page.width;
  const MX = 20;
  const x0 = MX;
  const x1 = PAGE_W - MX;
  const W = x1 - x0;
  let y = MX;

  // Normalize stroke style — same fix as purchase invoice. Ensures all
  // strokes (including grey-fill rect outlines) render at uniform 0.5pt
  // black lines.
  doc.lineWidth(0.5).strokeColor('#000');

  // Helper: Indian-style number formatter reused across the layout
  const fmtQty = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  const fmtRup = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const sumKey = (rows, primary, fallback) => rows.reduce((s, r) => s + Number(r[primary] || (fallback ? r[fallback] : 0) || 0), 0);

  // ── ORIGINAL/DUPLICATE/TRIPLICATE above outer border ──
  doc.fontSize(7.5).font('Helvetica').fillColor('#000');
  doc.text('ORIGINAL/DUPLICATE/TRIPLICATE', x0, y, { width: W, align: 'right' });
  y += 9;

  const boxTopY = y;

  // ── ASP header block (company identity) ──
  y += 4;
  // Capture y BEFORE writing the company text, so we can position the logo
  // alongside the text block (vertically centered against it).
  const headerStartY = y;
  // Logo: Bill of Supply is always issued by ASP, so always use the ASP
  // logo. Optional — falls through silently if the file isn't present.
  const _logoPath = require('path').join(__dirname, 'public', 'logo-asp.png');
  const _fs = require('fs');
  let logoOffsetX = 0;
  if (_fs.existsSync(_logoPath)) {
    try {
      doc.image(_logoPath, x0 + 6, headerStartY, { fit: [60, 60] });
      logoOffsetX = 0; // logo lives in left gutter — text still centered
    } catch (_) { /* fall through silently */ }
  }
  doc.font('Helvetica-Bold').fontSize(11);
  doc.text(cfg.s_company || 'AMAZING SPICE PARK PRIVATE LIMITED', x0, y, { width: W, align: 'center' });
  y += 13;

  doc.font('Helvetica').fontSize(8);
  const addr1 = (cfg.s_address1 || '').toUpperCase();
  const addr2Bits = [
    (cfg.s_place || '').toUpperCase(),
    (cfg.s_pin ? '-' + cfg.s_pin : ''),
    (cfg.s_state ? (cfg.s_state.toUpperCase() + ' CODE:' + (cfg.s_st_code || '')) : '')
  ].filter(Boolean).join(' ').trim();
  const mobile = cfg.s_mobile || cfg.mobile || '';
  const addr2Full = addr2Bits + (mobile ? ' MOBILE:' + mobile : '');
  doc.text(addr1, x0, y, { width: W, align: 'center' }); y += 10;
  doc.text(addr2Full, x0, y, { width: W, align: 'center' }); y += 10;
  doc.text(`CIN:${cfg.s_cin || ''}  PAN:${cfg.s_pan || cfg.pan || ''}`, x0, y, { width: W, align: 'center' }); y += 10;
  doc.text(`GSTIN:${cfg.s_gstin || ''}  SBL:${cfg.s_sbl || cfg.sbl || ''}`, x0, y, { width: W, align: 'center' }); y += 10;
  if (cfg.s_email || cfg.email) {
    doc.text('e-Mail ID:' + (cfg.s_email || cfg.email), x0, y, { width: W, align: 'center' });
    y += 10;
  }
  y += 4;

  // ── Title + subtitle ──
  doc.font('Helvetica-Bold').fontSize(12);
  doc.text('Bill for Purchase from Agriculturist', x0, y, { width: W, align: 'center' });
  y += 14;
  doc.font('Helvetica-Oblique').fontSize(8);
  doc.text('[Self-Generated Bill for Unregistered Purchases from Agriculturist Sec 2(7) of the CGST Act,2017]', x0, y, { width: W, align: 'center' });
  y += 14;

  // ── Invoice No / e-TRADE No / Date strip ──
  const infoY = y;
  const infoH = 16;
  doc.moveTo(x0, infoY).lineTo(x1, infoY).stroke();
  doc.moveTo(x0, infoY + infoH).lineTo(x1, infoY + infoH).stroke();
  const invDate = (billData && billData.billDate) || new Date().toLocaleDateString('en-GB');
  const eTradeNo = (billData && billData.eTradeNo) || cfg.e_trade_no || '';
  // Mode-aware: 'e-TRADE No:' or 'e-AUCTION No:' from cfg.business_mode.
  const eTradeLabel2 = (cfg.business_mode === 'e-Auction') ? 'e-AUCTION No' : 'e-TRADE No';
  doc.font('Helvetica').fontSize(8.5);
  doc.text(`Invoice No: ${billNo || ''}`, x0 + 6, infoY + 4);
  doc.text(`${eTradeLabel2}: ${eTradeNo}`, x0, infoY + 4, { width: W, align: 'center' });
  doc.text(`Date: ${invDate}`, x0, infoY + 4, { width: W - 6, align: 'right' });
  y = infoY + infoH;

  // ── Two-column seller / consignee header block ──
  const sellerW = Math.floor(W / 2);
  const conW = W - sellerW;
  const splitX = x0 + sellerW;
  const headH = 14;
  // Horizontal divider ABOVE the header band
  doc.moveTo(x0, y).lineTo(x1, y).stroke();
  doc.rect(x0, y, sellerW, headH).fill('#e8e8e8').stroke();
  doc.rect(splitX, y, conW, headH).fill('#e8e8e8').stroke();
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
  doc.text('DETAILS OF SELLER [BILLED FOR]', x0, y + 3, { width: sellerW, align: 'center' });
  doc.text('DETAILS OF CONSIGNEE [SHIPPED TO]', splitX, y + 3, { width: conW, align: 'center' });
  // Explicit vertical divider between SELLER and CONSIGNEE columns
  doc.moveTo(splitX, y).lineTo(splitX, y + headH).stroke();
  // Horizontal divider BELOW the header band
  doc.moveTo(x0, y + headH).lineTo(x1, y + headH).stroke();
  y += headH;

  // Seller body (up to ~6 rows visible)
  const bodyLineH = 10;
  const sellerLines = [];
  const seller = billData.seller || {};
  if (seller.name) sellerLines.push('M/s.' + seller.name);
  if (seller.address) sellerLines.push(seller.address);
  if (seller.place) sellerLines.push((seller.place || '').toUpperCase() + (seller.pan ? '   PAN:' + seller.pan : ''));
  if (seller.state) sellerLines.push('STATE:' + (seller.state || '').toUpperCase() + '   CODE:' + (seller.st_code || ''));
  sellerLines.push('CR.' + (seller.crno || ''));

  const bodyH = Math.max(6, sellerLines.length + 2) * bodyLineH;
  doc.moveTo(x0, y).lineTo(x0, y + bodyH).stroke();
  doc.moveTo(splitX, y).lineTo(splitX, y + bodyH).stroke();
  doc.moveTo(x1, y).lineTo(x1, y + bodyH).stroke();
  doc.moveTo(x0, y + bodyH).lineTo(x1, y + bodyH).stroke();
  doc.font('Helvetica').fontSize(8);
  let ly = y + 3;
  for (const line of sellerLines) { doc.text(line, x0 + 6, ly, { width: sellerW - 12 }); ly += bodyLineH; }
  y += bodyH;

  // ── Description of Goods / HSN row ──
  // When flag_hsn is OFF, hide the HSN CODE column and span Description
  // across the full row width (matches user expectation that disabling the
  // flag removes any reference to HSN from the PDF).
  const showHsn = readFlag(cfg.flag_hsn, true);
  const descH = 14;
  if (showHsn) {
    doc.rect(x0, y, sellerW, descH).fill('#e8e8e8').stroke();
    doc.rect(splitX, y, conW, descH).fill('#e8e8e8').stroke();
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
    doc.text('Description of Goods:' + (cfg.desc_goods || 'CARDAMOM'), x0, y + 3, { width: sellerW, align: 'center' });
    doc.text('HSN CODE:' + (cfg.hsn_cardamom || '09083120'), splitX, y + 3, { width: conW, align: 'center' });
  } else {
    doc.rect(x0, y, W, descH).fill('#e8e8e8').stroke();
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
    doc.text('Description of Goods:' + (cfg.desc_goods || 'CARDAMOM'), x0, y + 3, { width: W, align: 'center' });
  }
  // Horizontal dividers ABOVE and BELOW (drawn AFTER fill so they survive)
  doc.moveTo(x0, y).lineTo(x1, y).stroke();
  doc.moveTo(x0, y + descH).lineTo(x1, y + descH).stroke();
  y += descH;

  // ── Line-item table header ──
  const colSpec = [
    { key: 'lot',   label: 'LOT\nNO',                    w: 34 },
    { key: 'qty',   label: 'QUANTITY\nKGS',              w: 66 },
    { key: 'gap1',  label: '',                           w: 6 },
    { key: 'tqty',  label: 'TOTAL\nQTY/KGS',             w: 64 },
    { key: 'price', label: 'PRICE\nRs.  P.',             w: 52 },
    { key: 'value', label: 'VALUE\nRs.     P.',          w: 72 },
    { key: 'gap2',  label: '',                           w: 6 },
    { key: 'tax',   label: 'TAXABLE VALUE\n( Rs.    P.)', w: 80 },
    { key: 'cgst',  label: 'C G S T\n( 2.5%)',           w: 56 },
    { key: 'sgst',  label: 'S G S T\n( 2.5%)',           w: 56 },
    { key: 'igst',  label: 'I G S T\n( 5.0%)',           w: 64 },
  ];
  const totCol = colSpec.reduce((s, c) => s + c.w, 0);
  const scale = W / totCol;
  for (const c of colSpec) c.w = c.w * scale;
  let cx = x0;
  for (const c of colSpec) { c.x = cx; cx += c.w; }

  const hdrH = 22;
  doc.rect(x0, y, W, hdrH).fill('#e8e8e8').stroke();
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(7.5);
  for (const c of colSpec) {
    if (!c.label) continue;
    doc.text(c.label, c.x + 1, y + 2, { width: c.w - 2, align: 'center' });
  }
  // Explicit horizontal divider beneath the LOT NO column header
  doc.moveTo(x0, y + hdrH).lineTo(x1, y + hdrH).stroke();
  y += hdrH;

  // ── Body rows ──
  // Alternate-row stripes (light grey on odd rows, like sales invoice).
  // Controlled by flag_invoice_stripe (default ON). Fill is drawn BEFORE
  // text so text sits on top of the colored row band.
  const STRIPE_ON = (() => {
    const v = cfg.flag_invoice_stripe;
    if (v === undefined || v === null || v === '') return true;
    if (typeof v === 'boolean') return v;
    return String(v).toLowerCase() === 'true';
  })();
  const STRIPE_COLOR = '#ECECEC';
  function stripeRow(ry, rh, rowIndex) {
    if (!STRIPE_ON) return;
    if (rowIndex % 2 !== 1) return;
    doc.save();
    doc.rect(x0, ry, W, rh).fillColor(STRIPE_COLOR).fill();
    doc.restore();
    doc.fillColor('#000');
  }

  const lineH = 12;
  let bodyStartY = y;
  doc.font('Helvetica').fontSize(9).fillColor('#000');
  // Adaptive row count: fill the table down to a fixed Y position near the
  // bottom of the A4 page, so the totals/signature block always lands at
  // the page bottom regardless of how many line items there are.
  //
  // For invoices whose data exceeds what fits on one page, we paginate:
  // draw rows up to the page's max, then addPage() and redraw the table
  // header on the new page. The totals/signature block always lives at
  // the bottom of the LAST page.
  const PAGE_H = doc.page.height;          // A4 height = 842pt
  const PAGE_BOTTOM_MARGIN = 20;
  // Reserved bottom-block height (totals + amount-in-words + signatures)
  const BOTTOM_RESERVE = 200;
  const targetTableBottom = PAGE_H - PAGE_BOTTOM_MARGIN - BOTTOM_RESERVE;
  // How many data rows fit on the FIRST page (header consumed bodyStartY)
  let maxRowsThisPage = Math.max(1, Math.floor((targetTableBottom - bodyStartY) / lineH) - 1);
  const MIN_ROWS = 10;
  const rows = billData.lineItems || [];

  // Helper to repaint the column header on a continuation page
  const tableHeaderTop = bodyStartY - hdrH;
  function paintHeaderAt(y0) {
    doc.rect(x0, y0, W, hdrH).fill('#e8e8e8').stroke();
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(7.5);
    for (const c of colSpec) {
      if (!c.label) continue;
      doc.text(c.label, c.x + 1, y0 + 2, { width: c.w - 2, align: 'center' });
    }
  }

  // Track per-page state so we can draw vertical column dividers correctly
  // for each page segment of the table.
  const pageSegments = [{ start: tableHeaderTop, end: 0 }]; // end filled in below
  let drawnRows = 0;

  for (let i = 0; i < rows.length; i++) {
    // If this row would push us past the target table bottom, start a new page
    if (drawnRows >= maxRowsThisPage) {
      // Close out the current page's table — vertical dividers + bottom border
      const segEnd = y;
      pageSegments[pageSegments.length - 1].end = segEnd;
      doc.moveTo(x0, segEnd).lineTo(x1, segEnd).stroke();
      for (const c of colSpec) {
        if (c.key.startsWith('gap')) continue;
        doc.moveTo(c.x, pageSegments[pageSegments.length - 1].start).lineTo(c.x, segEnd).stroke();
      }
      doc.moveTo(x1, pageSegments[pageSegments.length - 1].start).lineTo(x1, segEnd).stroke();
      // Optional "Continued..." footer note
      doc.font('Helvetica-Oblique').fontSize(8).fillColor('#666');
      doc.text('— Continued on next page —', x0, segEnd + 4, { width: W, align: 'center' });
      doc.fillColor('#000');

      // New page
      doc.addPage({ size: 'A4', margin: 20 });
      y = 20;
      // Redraw column header at top of new page
      paintHeaderAt(y);
      y += hdrH;
      bodyStartY = y;
      pageSegments.push({ start: y - hdrH, end: 0 });
      // Reset per-page counters; full page available now (no header block above)
      maxRowsThisPage = Math.max(1, Math.floor((targetTableBottom - bodyStartY) / lineH) - 1);
      drawnRows = 0;
      doc.font('Helvetica').fontSize(9).fillColor('#000');
    }

    const li = rows[i];
    stripeRow(y, lineH, i);
    doc.text(String(li.lot || '').padStart(3, '0'), colSpec[0].x, y + 2, { width: colSpec[0].w, align: 'center' });
    doc.text(fmtQty(li.qty), colSpec[1].x, y + 2, { width: colSpec[1].w - 3, align: 'right' });
    doc.text(fmtQty(li.pqty || li.qty), colSpec[3].x, y + 2, { width: colSpec[3].w - 3, align: 'right' });
    doc.text(fmtRup(li.prate || li.price), colSpec[4].x, y + 2, { width: colSpec[4].w - 3, align: 'right' });
    doc.text(fmtRup((li.prate || li.price) * (li.pqty || li.qty)), colSpec[5].x, y + 2, { width: colSpec[5].w - 3, align: 'right' });
    doc.text(fmtRup(li.puramt || li.amount), colSpec[7].x, y + 2, { width: colSpec[7].w - 3, align: 'right' });
    y += lineH;
    drawnRows++;
  }

  // After all data rows: pad empty rows on the LAST page so totals land
  // at the bottom of the page (only when there's room — large invoices
  // that already filled the page won't get extra padding).
  const remainingHeight = targetTableBottom - y;
  const minRowsTarget = Math.max(MIN_ROWS - rows.length - 1, 0);
  const padRows = Math.max(minRowsTarget, Math.floor(remainingHeight / lineH) - 1);
  for (let i = 0; i < padRows; i++) {
    stripeRow(y, lineH, rows.length + i);
    y += lineH;
  }
  pageSegments[pageSegments.length - 1].end = y; // updated again after TOTAL below

  // ── TOTAL row ──
  const totalRowY = y;
  doc.font('Helvetica-Bold').fontSize(9);
  doc.text('TOTAL', colSpec[0].x, totalRowY + 2, { width: colSpec[0].w, align: 'center' });
  doc.text(fmtQty(sumKey(rows, 'qty')), colSpec[1].x, totalRowY + 2, { width: colSpec[1].w - 3, align: 'right' });
  doc.text(fmtQty(sumKey(rows, 'pqty', 'qty')), colSpec[3].x, totalRowY + 2, { width: colSpec[3].w - 3, align: 'right' });
  const valueSum = rows.reduce((s, li) => s + (li.prate || li.price) * (li.pqty || li.qty), 0);
  doc.text(fmtRup(valueSum), colSpec[5].x, totalRowY + 2, { width: colSpec[5].w - 3, align: 'right' });
  doc.text(fmtRup(sumKey(rows, 'puramt', 'amount')), colSpec[7].x, totalRowY + 2, { width: colSpec[7].w - 3, align: 'right' });
  y += lineH;

  // Table borders. For multi-page tables, draw verticals per page segment
  // (each page has its own header→tableEnd top/bottom range). The current
  // (last) page's segment ends at tableEndY below.
  const tableEndY = y;
  pageSegments[pageSegments.length - 1].end = tableEndY;
  doc.moveTo(x0, tableEndY).lineTo(x1, tableEndY).stroke();
  doc.moveTo(x0, totalRowY).lineTo(x1, totalRowY).stroke();
  for (const seg of pageSegments) {
    for (const c of colSpec) {
      if (c.key.startsWith('gap')) continue;
      doc.moveTo(c.x, seg.start).lineTo(c.x, seg.end).stroke();
    }
    doc.moveTo(x1, seg.start).lineTo(x1, seg.end).stroke();
  }

  // ── Totals summary bottom-right ──
  const sumX = x0 + W * 0.48;
  const sumBlockW = W - (sumX - x0);
  const sumLabelW = sumBlockW * 0.58;
  const sumValW = sumBlockW - sumLabelW;
  const s = billData.summary || {};
  const sumRows = [
    ['Total Taxable Value', s.totalPuramt || valueSum],
    ['Total Integrated Tax', s.igst || 0],
    ['Total Central Tax', s.cgst || 0],
    ['Total State Tax', s.sgst || 0],
    ['Round UP/DOWN', s.roundDiff || 0],
  ];
  doc.font('Helvetica').fontSize(8.5);
  const sumStartY = y;
  for (const [lbl, v] of sumRows) {
    doc.text(lbl, sumX + 4, y + 2, { width: sumLabelW });
    doc.text(fmtRup(v), sumX + sumLabelW, y + 2, { width: sumValW - 4, align: 'right' });
    y += 12;
  }
  doc.font('Helvetica-Bold').fontSize(9);
  doc.text('Total Value', sumX + 4, y + 2, { width: sumLabelW });
  doc.text(fmtRup(s.netAmount || s.totalPuramt || valueSum), sumX + sumLabelW, y + 2, { width: sumValW - 4, align: 'right' });
  y += 14;
  // Vertical separator between label column and value column
  doc.moveTo(sumX + sumLabelW, sumStartY).lineTo(sumX + sumLabelW, y).stroke();

  // DC No. / Date: aligned with "Total Value" row on the left
  doc.font('Helvetica').fontSize(8);
  doc.text('DC No.', x0 + 6, y - 14);
  doc.text('Date:', x0 + W * 0.22, y - 14);

  // Separator before amount-in-words
  doc.moveTo(x0, y).lineTo(x1, y).stroke(); y += 6;

  // ── Amount in words ──
  const netForWords = Math.round(s.netAmount || s.totalPuramt || valueSum);
  doc.fontSize(9).font('Helvetica');
  doc.text(amountToWords(netForWords) + ' Only', x0 + 6, y, { width: W - 12 });
  y += 16;

  // ── "for COMPANY" right-aligned ──
  doc.fontSize(9).font('Helvetica-Bold');
  doc.text('for ' + (cfg.s_company || 'AMAZING SPICE PARK PVT LTD'), x0, y, { width: W - 6, align: 'right' });
  y += 30;

  // ── Signatures ──
  doc.font('Helvetica').fontSize(8);
  doc.text('Signature of Seller', x0 + 6, y);
  doc.text('Authorized Signatory', x0, y, { width: W - 6, align: 'right' });
  y += 14;

  // Outer border. For single-page invoices, wrap from the title strip to
  // bottom. For multi-page, the page-spanning rectangle would render
  // incorrectly so skip it — table verticals + horizontal borders below
  // each section already provide visual structure.
  const boxBottomY = y;
  if (pageSegments.length === 1) {
    doc.rect(x0, boxTopY, W, boxBottomY - boxTopY).lineWidth(1).stroke();
  } else {
    // Last page only: simple bottom-portion outer rectangle
    doc.rect(x0, pageSegments[pageSegments.length - 1].start, W, boxBottomY - pageSegments[pageSegments.length - 1].start).lineWidth(1).stroke();
  }

  if (isBatch) return Promise.resolve(null);
  return new Promise((resolve) => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

/**
 * Generate ONE merged PDF containing multiple sales invoices (one per page
 * group, with continuation pages for long ones). Used by the bulk-print
 * endpoint so the user gets a single file they can print in one go.
 *
 * invoices: array of { invoiceData, saleType, invoiceNo, invoiceDate }
 * cfg: shared company-settings config (same for all)
 */
function generateSalesInvoicesBatchPDF(invoices, cfg, variant) {
  if (!Array.isArray(invoices) || invoices.length === 0) {
    return Promise.reject(new Error('No invoices to print'));
  }
  const doc = new PDFDocument({ size: 'A4', margin: 20 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));
  doc._invoiceCount = 0; // tracked so subsequent invoices addPage before drawing

  // `variant` (e.g. 'purchase') is a uniform display-flip applied to every
  // invoice in the batch — used by the bulk purchase-view endpoint to
  // render the ISPL-side mirror of every selected ASP sales invoice.
  for (const inv of invoices) {
    generateSalesInvoicePDF(
      inv.invoiceData,
      cfg,
      inv.saleType,
      inv.invoiceNo,
      inv.invoiceDate,
      doc, // externalDoc — triggers batch mode
      variant
    );
  }

  return new Promise((resolve) => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

// ── Batch: merge N purchase invoices into one PDF ──────────────────────
// Same pattern as sales batch. Each item: { invoiceData, invoiceNo }.
function generatePurchaseInvoicesBatchPDF(invoices, cfg) {
  if (!Array.isArray(invoices) || invoices.length === 0) {
    return Promise.reject(new Error('No purchase invoices to print'));
  }
  const doc = new PDFDocument({ size: 'A4', margin: 30 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));
  doc._purchaseCount = 0;

  for (const inv of invoices) {
    generatePurchaseInvoicePDF(inv.invoiceData, cfg, inv.invoiceNo, doc);
  }

  return new Promise((resolve) => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

// ── Batch: merge N agri bills into one PDF ─────────────────────────────
// Each item: { billData, billNo }.
function generateAgriBillsBatchPDF(bills, cfg) {
  if (!Array.isArray(bills) || bills.length === 0) {
    return Promise.reject(new Error('No bills to print'));
  }
  const doc = new PDFDocument({ size: 'A4', margin: 30 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));
  doc._billCount = 0;

  for (const bill of bills) {
    generateAgriBillPDF(bill.billData, cfg, bill.billNo, doc);
  }

  return new Promise((resolve) => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

// ── Commission Bill (Bills of Supply, Commission-Bill mode) ──────────
// IMCPC "MEMORANDUM OF CARDAMOM SOLD" format — one A4 page per LOT.
// Seller (billed-for) on the left, purchaser (sold-to) on the right,
// a cardamom line + optional sample-refund sub-row, a [-] COMMISSION row
// (HSN 996111) with CGST/SGST (intra) or IGST (inter), a TAX AMOUNT band,
// a NETT AMOUNT strip and amount-in-words. Self-contained — depends only
// on effectiveCompany() + amountToWords(). billData shape:
//   { seller:{name,address,place,state,st_code,cr,pan},
//     purchaser:{name,invo,address,place,pin,state,st_code,sbl,gstin,pan},
//     auction:{ano,date}, crpt, lineItems:[{lot,qty,bags,rate,cardamomCost,
//     refundQty,refundRate,refundAmount}], commission, gstRate, cgst, sgst,
//     igst, interState, nett, hsnCardamom, hsnCommission }
function generateCommissionBoSPDF(billData, cfg, billNo, externalDoc) {
  const isBatch = !!externalDoc;
  let doc, buffers;
  if (isBatch) {
    doc = externalDoc;
    if (doc._commBoSCount && doc._commBoSCount > 0) doc.addPage({ size: 'A4', margin: 20 });
    doc._commBoSCount = (doc._commBoSCount || 0) + 1;
  } else {
    doc = new PDFDocument({ size: 'A4', margin: 20 });
    buffers = [];
    doc.on('data', b => buffers.push(b));
  }

  const PAGE_W = doc.page.width;
  const MX = 20;
  const x0 = MX;
  const x1 = PAGE_W - MX;
  const W = x1 - x0;
  let y = MX;

  doc.lineWidth(0.5).strokeColor('#000');

  const fmtQty = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  const fmtRup = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // Resolve company identity using the same state-aware helper the rest
  // of this file uses (effectiveCompany picks KL/TN block based on
  // cfg.business_state) — guarantees the address actually renders.
  const co = effectiveCompany(cfg);
  const hsnCardamom = billData.hsnCardamom || cfg.hsn_cardamom || '09083120';
  const hsnCommission = billData.hsnCommission || '996111';

  // Draw one centered header line, shrinking the font (down to minSize)
  // until it fits within W so a long company name / address / CIN never
  // wraps to a second line and overprints the row below it.
  const drawCenteredFit = (text, yPos, font, baseSize, minSize) => {
    const s = String(text == null ? '' : text);
    let size = baseSize;
    doc.font(font).fontSize(size);
    while (size > minSize && doc.widthOfString(s) > W - 6) { size -= 0.25; doc.fontSize(size); }
    doc.text(s, x0, yPos, { width: W, align: 'center', lineBreak: false });
    doc.fontSize(baseSize);
  };

  // ── "ORIGINAL/DUPLICATE/TRIPLICATE" tag (top-right, above outer border) ──
  doc.font('Helvetica').fontSize(7.5).fillColor('#000');
  doc.text('ORIGINAL/DUPLICATE/TRIPLICATE', x0, y, { width: W - 4, align: 'right' });
  y += 9;

  const boxTopY = y;

  // ── Company header block ──
  y += 4;
  drawCenteredFit((co.name || '').toUpperCase(), y, 'Helvetica-Bold', 11, 8);
  y += 13;

  // REGD OFFICE line: assemble from whichever address fields the company
  // has configured. Some users embed the literal text "REGD OFFICE:" in
  // their address1 field, and others bake the state name + code into
  // address2 — we detect and skip those so the prefix and " STATE CODE:"
  // don't get printed twice.
  doc.font('Helvetica').fontSize(8);
  const addrJoined = [co.address1 || '', co.address2 || '']
    .filter(p => p && String(p).trim()).join(' ');
  const stateCodeAlreadyInAddr = co.stateName
    && new RegExp(co.stateName + '\\s*CODE\\s*:', 'i').test(addrJoined);
  const stateTail = (co.stateName && !stateCodeAlreadyInAddr)
    ? (co.stateName + ' CODE:' + (co.stateCode || '')) : '';
  const phoneTail = co.phone ? 'Ph:' + co.phone : '';
  const regdBody = [addrJoined, stateTail, phoneTail].filter(Boolean).join(' ');
  if (regdBody) {
    // Only add the "REGD OFFICE:" prefix when the user didn't already
    // bake one into their address line.
    const alreadyHasPrefix = /^\s*REGD\s+OFFICE\s*:/i.test(addrJoined);
    const regdLine = alreadyHasPrefix ? regdBody : ('REGD OFFICE:' + regdBody);
    drawCenteredFit(regdLine, y, 'Helvetica', 8, 6);
    y += 10;
  }
  const idLine = co.cinLabel
    ? { label: co.cinLabel, value: co.cinValue || '' }
    : { label: 'CIN', value: cfg.cin || '' };
  drawCenteredFit(`${idLine.label}:${idLine.value}   PAN:${co.pan || cfg.pan || ''}`, y, 'Helvetica', 8, 6);
  y += 10;
  drawCenteredFit(`GSTIN:${co.gstin || ''}   SBL:${co.sbl || cfg.sbl || ''}`, y, 'Helvetica', 8, 6);
  y += 10;
  if (co.email) {
    drawCenteredFit('e-Mail ID:' + co.email, y, 'Helvetica', 8, 6);
    y += 10;
  }
  y += 4;

  // ── Crop Receipt No (left) | Bill No (right) ──
  doc.font('Helvetica').fontSize(9);
  doc.text(`Cropt Receipt No.:${billData.crpt || ''}`, x0 + 6, y);
  doc.text(`Bill No: ${billNo || ''}`, x0, y, { width: W - 6, align: 'right' });
  y += 14;

  // ── Title + subtitle ──
  doc.font('Helvetica-Bold').fontSize(12);
  doc.text('COMMISSION BILL', x0, y, { width: W, align: 'center' });
  y += 14;
  doc.font('Helvetica').fontSize(8);
  doc.text('[ MEMORANDAM OF CARDAMOM SOLD THROUGH ' + (co.short || co.name || 'COMPANY') + ' ]',
    x0, y, { width: W, align: 'center' });
  y += 14;

  // ── eAuction No + Date strip ──
  const stripH = 16;
  doc.moveTo(x0, y).lineTo(x1, y).stroke();
  doc.moveTo(x0, y + stripH).lineTo(x1, y + stripH).stroke();
  doc.font('Helvetica').fontSize(8.5);
  const auc = billData.auction || {};
  doc.text(`eAuction No: ${auc.ano || ''}`, x0 + 6, y + 4);
  doc.text(`Date :${auc.date || ''}`, x0, y + 4, { width: W - 6, align: 'right' });
  y += stripH;

  // ── Two-column seller / purchaser header band ──
  const headColW = Math.floor(W / 2);
  const splitX = x0 + headColW;
  const bandH = 14;
  doc.rect(x0, y, headColW, bandH).fill('#e8e8e8').stroke();
  doc.rect(splitX, y, W - headColW, bandH).fill('#e8e8e8').stroke();
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(9);
  doc.text('DETAILS OF SELLER [BILLED FOR]', x0, y + 3, { width: headColW, align: 'center' });
  doc.text('DETAILS OF PURCHASER [SOLD TO]', splitX, y + 3, { width: W - headColW, align: 'center' });
  doc.moveTo(splitX, y).lineTo(splitX, y + bandH).stroke();
  doc.moveTo(x0, y + bandH).lineTo(x1, y + bandH).stroke();
  y += bandH;

  // ── Two-column seller / purchaser body ──
  const partyLineH = 10;
  const seller = billData.seller || {};
  const pur = billData.purchaser || {};
  const sellerLines = [];
  if (seller.name) sellerLines.push('Sri/M/s.' + seller.name);
  if (seller.address) sellerLines.push(seller.address);
  if (seller.place) sellerLines.push(String(seller.place || '').toUpperCase());
  if (seller.state) sellerLines.push(String(seller.state || '').toUpperCase() + '   CODE:' + (seller.st_code || ''));
  sellerLines.push('CR.' + (seller.cr || '') + (seller.pan ? '   PAN:' + seller.pan : ''));

  const purLines = [];
  if (pur.name) purLines.push('M/s.' + pur.name + (pur.invo ? '   INV:' + pur.invo : ''));
  if (pur.address) purLines.push(pur.address);
  if (pur.place) purLines.push(String(pur.place || '').toUpperCase() + (pur.pin ? '   PIN:' + pur.pin : ''));
  if (pur.state) purLines.push(String(pur.state || '').toUpperCase() + '   CODE:' + (pur.st_code || '') + (pur.sbl ? '   SBL:' + pur.sbl : ''));
  if (pur.gstin) purLines.push('GSTIN:' + pur.gstin + (pur.pan ? '   PAN:' + pur.pan : ''));

  // Height: tall enough for the longest side (purchaser block has up to
  // 5 long lines: name, two address lines, state+SBL, GSTIN+PAN), plus
  // a one-line margin at the bottom.
  const partyMaxLines = Math.max(sellerLines.length, purLines.length);
  const partyH = (Math.max(6, partyMaxLines) + 1) * partyLineH;
  doc.moveTo(x0, y).lineTo(x0, y + partyH).stroke();
  doc.moveTo(splitX, y).lineTo(splitX, y + partyH).stroke();
  doc.moveTo(x1, y).lineTo(x1, y + partyH).stroke();
  doc.moveTo(x0, y + partyH).lineTo(x1, y + partyH).stroke();
  // lineBreak:false is critical — without it, a long address line wraps
  // back to x=0 and overprints the next y-positioned line. To avoid losing
  // the END of long names / the GSTIN+PAN line to an ellipsis, pick the
  // largest font (8pt → 6pt floor) at which EVERY line on that side fits
  // its column, then draw uniformly.
  const fitPartySize = (lines, colW) => {
    let size = 8;
    while (size > 6) {
      doc.font('Helvetica').fontSize(size);
      if (lines.every(l => doc.widthOfString(String(l)) <= colW - 12)) break;
      size -= 0.25;
    }
    return size;
  };
  const sellerSize = fitPartySize(sellerLines, headColW);
  const purSize = fitPartySize(purLines, W - headColW);
  let sy = y + 3;
  doc.font('Helvetica').fontSize(sellerSize);
  for (const line of sellerLines) {
    doc.text(line, x0 + 6, sy, { width: headColW - 12, lineBreak: false, ellipsis: true });
    sy += partyLineH;
  }
  let py = y + 3;
  doc.font('Helvetica').fontSize(purSize);
  for (const line of purLines) {
    doc.text(line, splitX + 6, py, { width: W - headColW - 12, lineBreak: false, ellipsis: true });
    py += partyLineH;
  }
  y += partyH;

  // ── Line-item table column geometry ──
  // 13 leaf columns. CGST/SGST/IGST each split RATE | AMOUNT.
  const cols = [
    { key: 'lot',   w: 26 },   // 0
    { key: 'desc',  w: 74 },   // 1
    { key: 'hsn',   w: 44 },   // 2
    { key: 'qty',   w: 50 },   // 3
    { key: 'bags',  w: 32 },   // 4
    { key: 'rate',  w: 44 },   // 5
    { key: 'cost',  w: 79 },   // 6  CARDAMOM COST / TAXABLE VALUE
    { key: 'cgstR', w: 28 },   // 7
    { key: 'cgstA', w: 50 },   // 8
    { key: 'sgstR', w: 28 },   // 9
    { key: 'sgstA', w: 50 },   // 10
    { key: 'igstR', w: 28 },   // 11
    { key: 'igstA', w: 50 },   // 12
  ];
  const colsTot = cols.reduce((s, c) => s + c.w, 0);
  const colScale = W / colsTot;
  for (const c of cols) c.w = c.w * colScale;
  let cx = x0;
  for (const c of cols) { c.x = cx; cx += c.w; }
  const colEndX = c => c.x + c.w;

  // Right-aligned numeric drawer that shrinks the font from `baseSize` down to
  // a 6pt floor so a wide ₹ amount fits on one line instead of being clipped.
  const drawNum = (text, x, yPos, w, baseSize) => {
    const s = String(text == null ? '' : text);
    let size = baseSize;
    doc.font('Helvetica').fontSize(size);
    while (size > 6 && doc.widthOfString(s) > w) { size -= 0.25; doc.fontSize(size); }
    doc.text(s, x, yPos, { width: w, align: 'right', lineBreak: false });
    doc.fontSize(baseSize);
  };

  // ── Table header (two-band): top band has merged "CGST"/"SGST"/"IGST"
  // labels; bottom band has per-leaf sub-labels. ──
  const hdrTopH = 11;
  const hdrBotH = 14;
  const hdrTotalH = hdrTopH + hdrBotH;
  doc.rect(x0, y, W, hdrTotalH).fill('#e8e8e8').stroke();
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(7.5);

  const topLeafLabels = [
    'LOT', 'DESCRIPTION', 'HSN', 'QTY', 'UNIT', 'RATE', 'CARDAMOM',
  ];
  const botLeafLabelsLeft = [
    'NO', 'OF GOODS', 'SAC', 'Kg. gr.', 'BAGS', 'Rs. P.', 'COST',
  ];
  for (let i = 0; i < 7; i++) {
    doc.text(topLeafLabels[i], cols[i].x + 1, y + 2, { width: cols[i].w - 2, align: 'center' });
    doc.text(botLeafLabelsLeft[i], cols[i].x + 1, y + 2 + hdrTopH, { width: cols[i].w - 2, align: 'center' });
  }
  // Merged group labels
  const grpDefs = [
    { label: 'CGST', from: 7,  to: 8  },
    { label: 'SGST', from: 9,  to: 10 },
    { label: 'IGST', from: 11, to: 12 },
  ];
  for (const g of grpDefs) {
    const gx = cols[g.from].x;
    const gw = colEndX(cols[g.to]) - gx;
    doc.text(g.label, gx, y + 2, { width: gw, align: 'center' });
    doc.text('RATE', cols[g.from].x, y + 2 + hdrTopH, { width: cols[g.from].w, align: 'center' });
    doc.text('AMOUNT', cols[g.to].x, y + 2 + hdrTopH, { width: cols[g.to].w, align: 'center' });
  }

  // Header dividers
  doc.moveTo(cols[7].x, y + hdrTopH).lineTo(x1, y + hdrTopH).stroke();
  for (const c of cols) doc.moveTo(c.x, y).lineTo(c.x, y + hdrTotalH).stroke();
  doc.moveTo(x1, y).lineTo(x1, y + hdrTotalH).stroke();
  doc.moveTo(x0, y + hdrTotalH).lineTo(x1, y + hdrTotalH).stroke();
  y += hdrTotalH;

  // ── Body rows ──
  const items = billData.lineItems || [];
  const rowH = 14;
  const bodyStartY = y;
  doc.font('Helvetica').fontSize(7.5).fillColor('#000');

  let totalCost = 0, totalRefund = 0;
  for (const li of items) {
    const cost = Number(li.cardamomCost || 0);
    const refAmt = Number(li.refundAmount || 0);
    totalCost += cost;
    totalRefund += refAmt;

    const cellOpts = (col, align = 'center') => ({
      width: col.w - (align === 'right' ? 3 : 0),
      align, lineBreak: false, ellipsis: true,
    });
    doc.save(); doc.moveTo(x0, y).lineTo(x1, y).lineWidth(0.25).strokeColor('#CCC').stroke(); doc.restore();
    // Pad numeric lot numbers; leave blank/non-numeric (consolidated line) as-is.
    const _lotStrC = /^\d+$/.test(String(li.lot || '')) ? String(li.lot).padStart(3, '0') : String(li.lot || '');
    doc.text(_lotStrC, cols[0].x, y + 3, cellOpts(cols[0]));
    doc.text('CARDAMOM',                           cols[1].x, y + 3, cellOpts(cols[1]));
    doc.text(hsnCardamom,                          cols[2].x, y + 3, cellOpts(cols[2]));
    drawNum(fmtQty(li.qty),  cols[3].x, y + 3, cols[3].w - 3, 7.5);
    doc.text(String(li.bags || ''),                cols[4].x, y + 3, cellOpts(cols[4]));
    drawNum(fmtRup(li.rate), cols[5].x, y + 3, cols[5].w - 3, 7.5);
    drawNum(fmtRup(cost),    cols[6].x, y + 3, cols[6].w - 3, 7.5);
    y += rowH;

    // [+] SAMPLE REFUND row (only if refund > 0)
    if (refAmt > 0 || Number(li.refundQty || 0) > 0) {
      doc.save(); doc.moveTo(x0, y).lineTo(x1, y).lineWidth(0.25).strokeColor('#CCC').stroke(); doc.restore();
      doc.text('[+]',                          cols[0].x, y + 3, cellOpts(cols[0]));
      doc.text('SAMPLE REFUND',                cols[1].x, y + 3, cellOpts(cols[1]));
      drawNum(fmtQty(li.refundQty),           cols[3].x, y + 3, cols[3].w - 3, 7.5);
      drawNum(fmtRup(li.refundRate||li.rate), cols[5].x, y + 3, cols[5].w - 3, 7.5);
      drawNum(fmtRup(refAmt),                 cols[6].x, y + 3, cols[6].w - 3, 7.5);
      y += rowH;
    }
  }

  // Resolve GST amounts. If not supplied, derive from gstRate (intra-state
  // → CGST + SGST each = commission × gstRate%; inter-state → IGST = commission × 2*gstRate%).
  const commission = Number(billData.commission || 0);
  const gstRate = Number(billData.gstRate != null ? billData.gstRate : 9.0);
  const interState = !!billData.interState;
  let cgstAmt = Number(billData.cgst != null ? billData.cgst : 0);
  let sgstAmt = Number(billData.sgst != null ? billData.sgst : 0);
  let igstAmt = Number(billData.igst != null ? billData.igst : 0);
  if (!cgstAmt && !sgstAmt && !igstAmt && commission > 0 && gstRate > 0) {
    if (interState) {
      igstAmt = Math.round(commission * (2 * gstRate)) / 100;
    } else {
      cgstAmt = Math.round(commission * gstRate) / 100;
      sgstAmt = cgstAmt;
    }
  }

  // [-] COMMISSION row — separator above so it visually parts from lot rows
  doc.moveTo(x0, y).lineTo(x1, y).stroke();
  const cellOptsCom = (col, align = 'center') => ({
    width: col.w - (align === 'right' ? 3 : 0),
    align, lineBreak: false, ellipsis: true,
  });
  doc.font('Helvetica').fontSize(7.5);
  doc.text('[-]',                 cols[0].x, y + 3, cellOptsCom(cols[0]));
  doc.text('COMMISSION',          cols[1].x, y + 3, cellOptsCom(cols[1]));
  doc.text(hsnCommission,         cols[2].x, y + 3, cellOptsCom(cols[2]));
  drawNum(fmtRup(commission),     cols[6].x, y + 3, cols[6].w - 3, 7.5);
  if (cgstAmt) {
    doc.text(gstRate.toFixed(1) + '%', cols[7].x, y + 3, cellOptsCom(cols[7]));
    drawNum(fmtRup(cgstAmt),           cols[8].x, y + 3, cols[8].w - 3, 7.5);
  }
  if (sgstAmt) {
    doc.text(gstRate.toFixed(1) + '%', cols[9].x,  y + 3, cellOptsCom(cols[9]));
    drawNum(fmtRup(sgstAmt),           cols[10].x, y + 3, cols[10].w - 3, 7.5);
  }
  if (igstAmt) {
    doc.text((2 * gstRate).toFixed(1) + '%', cols[11].x, y + 3, cellOptsCom(cols[11]));
    drawNum(fmtRup(igstAmt),                 cols[12].x, y + 3, cols[12].w - 3, 7.5);
  }
  y += rowH;

  const tableEndY = y;
  // Verticals for the body
  for (const c of cols) doc.moveTo(c.x, bodyStartY).lineTo(c.x, tableEndY).stroke();
  doc.moveTo(x1, bodyStartY).lineTo(x1, tableEndY).stroke();
  doc.moveTo(x0, tableEndY).lineTo(x1, tableEndY).stroke();

  // ── Summary band: TAXABLE VALUE | CGST | SGST | IGST headers ──
  const sumHdrH = 14;
  const leftEndX = colEndX(cols[5]);
  doc.rect(x0, y, leftEndX - x0, sumHdrH).stroke();
  doc.rect(cols[6].x, y, cols[6].w, sumHdrH).fill('#e8e8e8').stroke();
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(8);
  doc.text('TAXABLE VALUE', cols[6].x, y + 3, { width: cols[6].w, align: 'center' });
  for (const g of grpDefs) {
    const gx = cols[g.from].x;
    const gw = colEndX(cols[g.to]) - gx;
    doc.rect(gx, y, gw, sumHdrH).fill('#e8e8e8').stroke();
    doc.fillColor('#000').text(g.label, gx, y + 3, { width: gw, align: 'center' });
  }
  y += sumHdrH;

  // TAX AMOUNT row: "TAX AMOUNT" spans cols 0..5; then per-group AMOUNTs.
  const taxRowH = 16;
  doc.rect(x0, y, leftEndX - x0, taxRowH).stroke();
  doc.font('Helvetica-Bold').fontSize(9);
  doc.text('TAX AMOUNT', x0, y + 4, { width: leftEndX - x0, align: 'center' });
  doc.font('Helvetica').fontSize(8.5);
  doc.rect(cols[6].x, y, cols[6].w, taxRowH).stroke();
  drawNum(fmtRup(commission), cols[6].x, y + 4, cols[6].w - 3, 8.5);
  for (const g of grpDefs) {
    doc.rect(cols[g.from].x, y, cols[g.from].w, taxRowH).stroke();
    doc.rect(cols[g.to].x,   y, cols[g.to].w,   taxRowH).stroke();
  }
  if (cgstAmt) drawNum(fmtRup(cgstAmt), cols[8].x,  y + 4, cols[8].w - 3,  8.5);
  if (sgstAmt) drawNum(fmtRup(sgstAmt), cols[10].x, y + 4, cols[10].w - 3, 8.5);
  if (igstAmt) drawNum(fmtRup(igstAmt), cols[12].x, y + 4, cols[12].w - 3, 8.5);
  y += taxRowH;

  // ── NETT AMOUNT strip ──
  const nett = Number(
    billData.nett != null
      ? billData.nett
      : (totalCost + totalRefund - commission - cgstAmt - sgstAmt - igstAmt)
  );
  const nettH = 18;
  const nettValueX = cols[10].x;
  doc.rect(x0, y, nettValueX - x0, nettH).stroke();
  doc.rect(nettValueX, y, x1 - nettValueX, nettH).stroke();
  doc.font('Helvetica-Bold').fontSize(10);
  doc.text('NETT AMOUNT', x0, y + 4, {
    width: nettValueX - x0 - 6, align: 'right', lineBreak: false,
  });
  {
    let _ns = 10;
    doc.font('Helvetica-Bold').fontSize(_ns);
    const _nw = x1 - nettValueX - 3;
    while (_ns > 7 && doc.widthOfString(fmtRup(nett)) > _nw) { _ns -= 0.25; doc.fontSize(_ns); }
    doc.text(fmtRup(nett), nettValueX, y + 4, { width: _nw, align: 'right', lineBreak: false });
  }
  y += nettH;

  // ── Amount in words ──
  // amountToWords already prefixes "Rupees " — do not add it again.
  y += 10;
  doc.font('Helvetica').fontSize(9);
  doc.text(amountToWords(Math.round(nett)) + ' Only', x0 + 6, y, {
    width: W - 12, lineBreak: false, ellipsis: true,
  });
  y += 24;

  // ── "for COMPANY" + signatures ──
  doc.font('Helvetica-Bold').fontSize(9);
  doc.text('for ' + (co.short || co.name || 'COMPANY'), x0, y, { width: W - 6, align: 'right' });
  y += 30;
  doc.font('Helvetica').fontSize(8);
  doc.text('Signature of Seller', x0 + 6, y);
  doc.text('Authorized Signatory', x0, y, { width: W - 6, align: 'right' });
  y += 14;

  // ── Outer rectangle border ──
  doc.rect(x0, boxTopY, W, y - boxTopY).lineWidth(1).stroke();

  if (isBatch) return Promise.resolve(null);
  return new Promise((resolve) => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}

/**
 * Merge N Commission Bills into a single PDF (one per page).
 * Each item: { billData, billNo }.
 */
function generateCommissionBoSBatchPDF(bills, cfg) {
  if (!Array.isArray(bills) || bills.length === 0) {
    return Promise.reject(new Error('No bills to print'));
  }
  const doc = new PDFDocument({ size: 'A4', margin: 20 });
  const buffers = [];
  doc.on('data', b => buffers.push(b));
  doc._commBoSCount = 0;

  for (const bill of bills) {
    generateCommissionBoSPDF(bill.billData, cfg, bill.billNo, doc);
  }

  return new Promise((resolve) => {
    doc.on('end', () => resolve(Buffer.concat(buffers)));
    doc.end();
  });
}
