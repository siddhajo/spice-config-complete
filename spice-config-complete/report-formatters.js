/**
 * report-formatters.js
 *
 * Shared utilities for export PDFs / XLSX:
 *   - Indian-format number formatters (rupees, kilos, prices)
 *   - Company header (logo + name + address) drawn at the top of every PDF
 *
 * The number formatters apply Indian-style comma grouping (last 3 digits, then
 * pairs going left): 1,98,000.00 — not 198,000.00 — for amounts in rupees,
 * and 1,100.000 for kilos.
 */

const fs = require('fs');
const path = require('path');
const { AsyncLocalStorage } = require('async_hooks');

// Per-request override for which company identity (ISP vs ASP) exports render
// under. When a request runs inside runWithCompanyOverride('ASP'|'ISP', fn),
// getCompanyHeader() below honours that instead of the DB's active preset —
// this powers the Export Center's "Use ISP / ASP company details" chooser
// (Tamil Nadu + e-Trade) without threading a param through every builder.
const _companyOverrideStore = new AsyncLocalStorage();
function runWithCompanyOverride(code, fn) {
  const norm = String(code || '').trim().toUpperCase();
  if (norm !== 'ISP' && norm !== 'ASP') return fn();
  return _companyOverrideStore.run(norm, fn);
}

// ── Indian-format number formatters ─────────────────────────

// Group an integer string (no sign) using Indian comma convention:
// last 3 digits keep together, every preceding pair gets a comma.
//   "1198000" -> "11,98,000"
//   "100"     -> "100"
//   "1000"    -> "1,000"
function groupIndian(digits) {
  if (digits.length <= 3) return digits;
  const last3 = digits.slice(-3);
  const rest  = digits.slice(0, -3);
  const restGrouped = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',');
  return `${restGrouped},${last3}`;
}

// Format a number with `decimals` decimal places and Indian comma grouping
// on the integer part.
function fmtIndian(n, decimals) {
  const num = Number(n) || 0;
  const fixed = num.toFixed(decimals);
  const [intPart, decPart] = fixed.split('.');
  const sign = intPart.startsWith('-') ? '-' : '';
  const digits = sign ? intPart.slice(1) : intPart;
  const grouped = groupIndian(digits);
  return decPart != null ? `${sign}${grouped}.${decPart}` : `${sign}${grouped}`;
}

// Rupees: 2 decimals, Indian commas.  Examples:
//   198000     -> "1,98,000.00"
//   1100.5     -> "1,100.50"
//   89.1       -> "89.10"
function fmtMoney(n) { return fmtIndian(n, 2); }

// Kilos: 3 decimals, Indian commas.  Examples:
//   1100       -> "1,100.000"
//   90.5       -> "90.500"
//   123456.789 -> "1,23,456.789"
function fmtQty(n) { return fmtIndian(n, 3); }

// Per-kg price (also rupees): same as money — 2 decimals, Indian commas.
function fmtPrice(n) { return fmtIndian(n, 2); }

// ── Date formatting (settings-driven) ───────────────────────
// Accepts ISO `yyyy-mm-dd[Thh:mm:ss]`, dd/mm/yyyy, dd-mm-yyyy, a JS Date, or an
// Excel serial number, and renders per the operator's `date_format` setting.
// Mirrors the front-end's fmtD() so a date round-trips identically from list
// view to PDF/XLSX. Unknown inputs return the input as-is.
const _MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function _parseDateParts(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date && !isNaN(v)) {
    const pad = x => String(x).padStart(2, '0');
    return { y: String(v.getFullYear()), m: pad(v.getMonth() + 1), d: pad(v.getDate()) };
  }
  const s = String(v).trim();
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return { y: iso[1], m: iso[2], d: iso[3] };
  const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmy) return { y: dmy[3], m: dmy[2].padStart(2, '0'), d: dmy[1].padStart(2, '0') };
  if (/^\d+(\.\d+)?$/.test(s)) {
    const n = Number(s);
    if (n > 0 && n < 80000) {
      const ms = (n - 25569) * 86400 * 1000;
      const date = new Date(ms);
      if (!isNaN(date)) {
        const pad = x => String(x).padStart(2, '0');
        return {
          y: String(date.getUTCFullYear()),
          m: pad(date.getUTCMonth() + 1),
          d: pad(date.getUTCDate()),
        };
      }
    }
  }
  return null;
}
function formatDateForDisplay(input, fmt) {
  const parts = _parseDateParts(input);
  if (!parts) return input == null ? '' : String(input);
  const { y, m, d } = parts;
  const mmm = _MONTH_ABBR[Number(m) - 1] || m;
  const want = String(fmt || 'dd/mm/yyyy').toLowerCase();
  switch (want) {
    case 'mm/dd/yyyy':   return `${m}/${d}/${y}`;
    case 'yyyy-mm-dd':   return `${y}-${m}-${d}`;
    case 'dd-mmm-yyyy':  return `${d}-${mmm}-${y}`;
    case 'dd mmm yyyy':  return `${d} ${mmm} ${y}`;
    case 'dd/mm/yyyy':
    default:             return `${d}/${m}/${y}`;
  }
}

// ── Company header (logo + name + address) ──────────────────

// Resolve the active company branding from company_settings. Falls back to
// hard-coded ISP defaults if the settings table doesn't exist or returns
// nothing — useful for export tools that run before the DB is initialized.
function getCompanyHeader(db) {
  // Defensive defaults so the header still renders if anything is missing
  let name = 'IDEAL SPICES PRIVATE LIMITED';
  let logoFile = 'logo-ispl.png';
  let address1 = '';
  let branch = '';

  if (db && typeof db.get === 'function') {
    try {
      // Determine which preset (ISP vs ASP) is active. Default to ISP.
      // A per-request override (Export Center chooser) wins over the DB's
      // stored active preset when present.
      let activeCode = 'ISP';
      const _ov = _companyOverrideStore.getStore();
      if (_ov === 'ISP' || _ov === 'ASP') {
        activeCode = _ov;
      } else {
        try {
          const r = db.get(
            "SELECT value FROM company_preset_meta WHERE key = 'active_preset_code'"
          );
          if (r && r.value) activeCode = r.value;
        } catch (_) {}
      }

      // Read short_name + logo for whichever preset is active.
      // The non-active preset's fields are stored under "s_" prefix (sister).
      // When active is ISP, use logo + short_name (primary fields).
      // When active is ASP, use s_logo + s_short_name (sister fields).
      const isASP = activeCode === 'ASP';
      const nameKey = isASP ? 's_short_name' : 'short_name';
      const logoKey = isASP ? 's_logo'        : 'logo';

      // Company-identity settings are per-mode now — scope to the active mode.
      const _bm = require('./company-config').getActiveMode(db);

      const nameRow = db.get('SELECT value FROM company_settings WHERE key = ? AND business_mode = ?', [nameKey, _bm]);
      if (nameRow && nameRow.value) name = nameRow.value;

      const logoRow = db.get('SELECT value FROM company_settings WHERE key = ? AND business_mode = ?', [logoKey, _bm]);
      const logoCode = logoRow && logoRow.value ? String(logoRow.value).toLowerCase() : (isASP ? 'asp' : 'isp');
      // logo-ispl.png is the actual filename in /public for the ISP preset
      logoFile = logoCode === 'asp' ? 'logo-asp.png' : 'logo-ispl.png';

      // Brand block shows three lines: company name, address line 1, and
      // office branch. For ISP we use the Tamil Nadu fields (tn_*); for
      // ASP we use the sister/Kerala fields (s_address1 + kl_branch).
      const addrPrefix = isASP ? 's_' : 'tn_';
      const branchKey = isASP ? 'kl_branch' : 'tn_branch';
      const a1Row = db.get('SELECT value FROM company_settings WHERE key = ? AND business_mode = ?', [addrPrefix + 'address1', _bm]);
      if (a1Row && a1Row.value) address1 = a1Row.value;
      const bRow = db.get('SELECT value FROM company_settings WHERE key = ? AND business_mode = ?', [branchKey, _bm]);
      if (bRow && bRow.value) branch = bRow.value;
    } catch (_) {
      // Ignore — fall through to defaults
    }
  }

  // Resolve the logo to an absolute path on disk; null if missing.
  const logoPath = path.join(__dirname, 'public', logoFile);
  const logoOnDisk = fs.existsSync(logoPath) ? logoPath : null;

  // The header object exposes `address1` and `address2` for backward
  // compatibility with PDF/XLSX renderers — `address2` now carries the
  // office branch instead of the second address line.
  return { name, logoPath: logoOnDisk, address1, address2: branch, branch };
}

// Truncate `text` so doc.widthOfString(out) <= maxWidth, appending an ellipsis
// when it doesn't fit. PDFKit 0.15 doesn't reliably honor `lineBreak: false`
// when text exceeds the column — pre-fitting avoids visual overlap.
function fitText(doc, text, maxWidth) {
  if (text == null) return '';
  const s = String(text);
  if (doc.widthOfString(s) <= maxWidth) return s;
  const ELLIP = '…';
  let lo = 0, hi = s.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    const candidate = s.slice(0, mid) + ELLIP;
    if (doc.widthOfString(candidate) <= maxWidth) lo = mid;
    else hi = mid - 1;
  }
  return lo > 0 ? s.slice(0, lo) + ELLIP : ELLIP;
}

// Draw a unified three-column header band on the current page:
//
//   ┌─────────────────────────────────────────────────────────────┐
//   │ ┌──────┐ COMPANY NAME              REPORT       Trade #3    │
//   │ │ LOGO │ Address line 1            TITLE        15/04/2026  │
//   │ └──────┘ Address line 2                                     │
//   └─────────────────────────────────────────────────────────────┘
//      (left block)            (middle, big)    (right, meta)
//
// The logo is sized to vertically span the company-text block so it acts as
// a proper brand mark instead of a floating ornament. Returns the y-coord
// just below the band — caller continues drawing from there.
//
// opts:
//   x, y, width — band position and width
//   title       — report title (e.g. "Bank Payment") shown in the middle
//   metaLines   — array of strings shown right-aligned (e.g. ["e-TRADE No: 3", "Date: 15/04/2026"])
//   showAddress — whether to show the address (default true; false on narrow halves)
//   logoH       — explicit logo height. Defaults to fit the company-text block.
function drawCompanyHeader(doc, header, opts) {
  const { x, y, width } = opts;
  const showAddress = opts.showAddress !== false;
  const title = opts.title || '';
  const metaLines = Array.isArray(opts.metaLines) ? opts.metaLines : [];
  const startY = y;
  const NAME_FONT = 10;
  const ADDR_FONT = 7;
  const TITLE_FONT = 16;
  const META_FONT = 9;
  const NAME_LINE_H = 13;
  const ADDR_LINE_H = 9;
  const GAP_LOGO_TEXT = 6;

  // ── Compute company-text block height so the logo can match it ──
  // 1 line of name + (0/1/2) lines of address.
  const addressLineCount = showAddress
    ? (header.address1 ? 1 : 0) + (header.address2 ? 1 : 0)
    : 0;
  const textBlockH = NAME_LINE_H + addressLineCount * ADDR_LINE_H;

  // Logo sized to match the text block height, but never smaller than a
  // legible brand-mark minimum — a 1-line-name + 2-line-address block is only
  // ~31pt, which rendered the logo as a tiny, barely-visible ornament. Floor it
  // so the company logo actually reads as a logo (the band height, left-column
  // reservation and title/meta centering below all key off logoH/logoW, so a
  // larger logo cascades correctly and never overlaps the text). Callers can
  // still override with opts.logoH for special cases (carbon-copy halves).
  const MIN_LOGO = 46;
  const logoH = opts.logoH || Math.max(textBlockH, MIN_LOGO);
  const logoW = opts.logoW || logoH;  // square by default; aspect-fit handles rectangular logos

  // Band height = the taller of logo vs text. Vertically center the company
  // text block against it so the name/address sit level with the logo instead
  // of top-aligning while a taller logo hangs below.
  const bandH = Math.max(logoH, textBlockH);
  const textTop = startY + (bandH - textBlockH) / 2;

  // ── Three-column geometry (adaptive, symmetric) ──
  // For the title to actually appear centered on the page, the left and
  // right reservations must be equal — so middle is symmetric around the
  // page center. When there's no title we don't need symmetry.
  const hasTitle = !!title;
  const hasMeta = metaLines.length > 0;
  let leftW, midW, rightW;
  if (hasTitle) {
    // Reserve equal-width side columns so the middle title column is
    // page-centered. Side width is sized to fit the company name (measured
    // at NAME_FONT) plus the logo + gap, with a percentage cap so the title
    // always has reasonable room.
    doc.font('Helvetica-Bold').fontSize(NAME_FONT);
    const nameW = doc.widthOfString(header.name || '');
    const desiredLeftContent = logoW + GAP_LOGO_TEXT + nameW + 8;
    // Clamp side width between 28% (so title keeps reasonable room) and 42%.
    // The upper bound was raised from 38% so a prominent logo + a longer
    // company name still fit on one line without ellipsizing; the title column
    // auto-shrinks to fit whatever middle width remains. Short brand names
    // clamp to the 28% floor, so their title width is unchanged.
    const sideW = Math.max(width * 0.28, Math.min(width * 0.42, desiredLeftContent));
    leftW = sideW;
    rightW = sideW;
    midW = width - leftW - rightW;
  } else if (hasMeta) {
    // Meta-only: small right column, left takes the rest.
    rightW = Math.min(width * 0.22, 110);
    leftW = width - rightW;
    midW = 0;
  } else {
    // No title and no meta — left block uses the full width.
    leftW = width;
    rightW = 0;
    midW = 0;
  }
  const midX = x + leftW;
  const rightX = x + width - rightW;

  // ── LEFT block: logo + company text ──
  let logoDrawn = false;
  if (header.logoPath) {
    try {
      doc.image(header.logoPath, x, startY, { fit: [logoW, logoH] });
      logoDrawn = true;
    } catch (_) { /* fall through to text-only */ }
  }
  const textX = logoDrawn ? x + logoW + GAP_LOGO_TEXT : x;
  const textW = leftW - (logoDrawn ? logoW + GAP_LOGO_TEXT : 0);

  // Company name (top of left text block). Use fitText so a long name doesn't
  // overrun and overlap into the middle column.
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(NAME_FONT)
     .text(fitText(doc, header.name || '', textW - 2), textX, textTop, {
       width: textW, align: 'left', lineBreak: false,
     });

  // Address lines beneath the name, in the same left column. Use fitText to
  // ellipsize lines that exceed the column width — PDFKit's `lineBreak: false`
  // is unreliable with long single-token strings.
  if (showAddress) {
    let aY = textTop + NAME_LINE_H;
    doc.font('Helvetica').fontSize(ADDR_FONT).fillColor('#444');
    if (header.address1) {
      doc.text(fitText(doc, header.address1, textW - 2), textX, aY, {
        width: textW, align: 'left', lineBreak: false,
      });
      aY += ADDR_LINE_H;
    }
    if (header.address2) {
      doc.text(fitText(doc, header.address2, textW - 2), textX, aY, {
        width: textW, align: 'left', lineBreak: false,
      });
    }
  }

  // ── MIDDLE: report title, vertically centered against the brand band ──
  if (title && midW > 30) {
    // Auto-shrink the title font size until it fits in one line within the
    // middle column. PDFKit 0.15 does not reliably honor `lineBreak: false`
    // when the text exceeds the column width — measuring + scaling avoids
    // ugly wrap.
    let titleSize = TITLE_FONT;
    doc.font('Helvetica-Bold');
    while (titleSize > 9 && (doc.fontSize(titleSize), doc.widthOfString(title)) > midW - 6) {
      titleSize -= 0.5;
    }
    const titleY = startY + (bandH - titleSize) / 2;
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(titleSize)
       .text(title, midX, titleY, {
         width: midW, align: 'center', lineBreak: false,
       });
  }

  // ── RIGHT: meta lines (Trade #, date, page, etc.) ──
  // Vertically center the meta block against the brand band.
  if (metaLines.length) {
    const totalMetaH = metaLines.length * (META_FONT + 4);
    let mY = startY + (bandH - totalMetaH) / 2;
    doc.fillColor('#000').font('Helvetica-Bold').fontSize(META_FONT);
    metaLines.forEach((line) => {
      doc.text(fitText(doc, line, rightW - 2), rightX, mY, {
        width: rightW, align: 'right', lineBreak: false,
      });
      mY += META_FONT + 4;
    });
  }

  // Small gap before the report body begins, measured from the band bottom.
  doc.fillColor('#000');
  return startY + bandH + 6;
}

// ── XLSX helpers ────────────────────────────────────────────

// Number-format hint for ExcelJS based on a column header label.
// Matches the PDF formatters: rupees → 2 decimals + Indian commas,
// kilos → 3 decimals + Indian commas, integer counts stay plain.
function xlsxNumFmtForHeader(header) {
  const h = String(header || '').toUpperCase();
  if (h === 'QTY' || h === 'PQTY' || h === 'LITRE' || h === 'KILOS' || h === 'QUANTITY') {
    return '#,##0.000';   // Excel will format Indian-style with the right locale; pattern is the standard 3-decimal lakh format
  }
  if (h === 'BAG' || h === 'BAGS' || h === 'LOTS' || h === 'LOT' || h === 'SL.NO' || h === 'S.NO') {
    return '#,##0';        // integer
  }
  // Treat all other numerics as rupees (price/amount/total/etc.)
  if (
    h === 'PRICE'   || h === 'AMOUNT' || h === 'PURAMT' ||
    h === 'PRATE'   || h === 'PAYABLE'|| h === 'DISCOUNT' ||
    h === 'BALANCE' || h === 'ADVANCE'|| h === 'VALUE' ||
    h === 'INV.AMOUNT' || h === 'TOTAL' || h === 'COMMISSION' ||
    h === 'CGST'    || h === 'SGST'   || h === 'IGST' ||
    h === 'REFUND'
  ) {
    return '#,##,##0.00';  // Indian-style 2-decimal with lakh grouping
  }
  return null;             // not numeric — no format
}

// Draw a unified three-column brand band at the top of an XLSX worksheet.
// Mirrors the PDF drawCompanyHeader layout: logo + name + address (left),
// report title (middle), meta lines (right). Returns the row number of the
// first row AFTER the brand band — caller continues writing from there.
//
// `ws` is an ExcelJS worksheet. `wb` is the workbook (needed to register
// the logo image). `colCount` is how many columns the band should span
// (typically the worksheet's data column count).
//
// Layout: 4 rows tall.
//   Row 1: [logo cell merged 1×3] [company name] [report title centered] [meta line 1]
//   Row 2: [logo cell continued]  [address 1]                            [meta line 2]
//   Row 3: [logo cell continued]  [address 2]                            [meta line 3]
//   Row 4: blank spacer
//
// The brand-band rows are merged into 3 horizontal bands so name/title/meta
// each get their own column groups regardless of the underlying data column
// widths. We split the available columns: ~30% left, ~40% middle, ~30% right.
function writeXlsxCompanyHeader(wb, ws, header, opts) {
  const colCount = Math.max(opts.colCount || 1, 1);
  const title = opts.title || '';
  const metaLines = Array.isArray(opts.metaLines) ? opts.metaLines : [];

  function colLetter(n) {
    let s = '';
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }
  function range(c1, r1, c2, r2) {
    return `${colLetter(c1)}${r1}:${colLetter(c2)}${r2}`;
  }

  // Build a list of cumulative character-widths across the data columns so
  // we can pick split points that approximate the PDF's 30/40/30 split.
  const colWidths = [];
  for (let i = 1; i <= colCount; i++) {
    const w = (ws.getColumn(i) && ws.getColumn(i).width) || 12;
    colWidths.push(w);
  }
  const totalW = colWidths.reduce((a, b) => a + b, 0);

  // Find the column index whose cumulative width crosses a given fraction of
  // totalW. Returns a 1-based column index.
  function colAtFraction(frac) {
    const target = totalW * frac;
    let acc = 0;
    for (let i = 0; i < colCount; i++) {
      acc += colWidths[i];
      if (acc >= target) return i + 1;
    }
    return colCount;
  }

  // ── Layout zones ──
  // Logo: column 1 only (we widen col 1 below if needed so the logo has room).
  // Left text: cols 2 .. ~40% of total width (room for company name).
  // Middle title: cols (left end + 1) .. ~75% of total width (lots of room for the title).
  // Right meta: cols (mid end + 1) .. last column (~25% for meta lines).
  const logoCol = 1;
  const leftStart = 2;
  // Ensure first column is wide enough for a 60×60-px logo (approx 9 char-units)
  if ((ws.getColumn(1).width || 0) < 9) ws.getColumn(1).width = 9;

  let leftEnd, midEnd;
  if (colCount <= 3) {
    // Too few columns to do a real 3-way split. Stack vertically: brand on
    // row 1-3 spanning all columns, title on row 1 right side if room.
    leftEnd = colCount;
    midEnd = colCount;
  } else {
    leftEnd = Math.max(leftStart, colAtFraction(0.40));
    midEnd = Math.max(leftEnd + 1, colAtFraction(0.75));
    if (midEnd >= colCount) midEnd = colCount - 1;

    function leftZoneWidth() {
      let w = 0;
      for (let i = leftStart; i <= leftEnd; i++) w += colWidths[i - 1];
      return w;
    }
    function midZoneWidth() {
      let w = 0;
      for (let i = leftEnd + 1; i <= midEnd; i++) w += colWidths[i - 1];
      return w;
    }

    // Make sure middle zone has room for the title (~18 units).
    let guard = 0;
    while (midZoneWidth() < 18 && guard++ < colCount) {
      if (leftEnd > leftStart && midZoneWidth() < 18 && leftZoneWidth() > 20) {
        leftEnd--;
        continue;
      }
      if (midEnd < colCount - 1) {
        midEnd++;
        continue;
      }
      break;
    }

    // Left zone needs enough width to fit the company name on a single line.
    // Address lines may wrap, but row heights below are tall enough to show
    // those wrapped lines without clipping.
    const nameLen = (header.name || '').length;          // chars at 11pt bold
    const LOGO_INDENT = 9;
    // 11pt bold needs ~1.1 char-unit per character + padding for indent and margin.
    const nameNeed = LOGO_INDENT + Math.ceil(nameLen * 1.1) + 2;
    // leftZoneWidth() doesn't include the logo column (col 1), so subtract
    // the logo column width from the budget. Cap at ~40 char-units so we
    // don't push other columns off the printable page.
    const NEEDED_LEFT = Math.max(0, Math.min(40, nameNeed - colWidths[0]));
    const shortfall = NEEDED_LEFT - leftZoneWidth();
    if (shortfall > 0) {
      const widenCol = ws.getColumn(leftEnd);
      widenCol.width = (widenCol.width || 12) + shortfall;
      colWidths[leftEnd - 1] += shortfall;  // keep our local copy in sync
    }

    // If meta lines are present, ensure the right zone is wide enough to fit
    // the longest meta line (~16 units for "Date: 15/04/2026" at 10pt bold).
    if (metaLines.length) {
      let rightZoneW = 0;
      for (let i = midEnd + 1; i <= colCount; i++) rightZoneW += colWidths[i - 1];
      const longestMeta = metaLines.reduce((m, s) => Math.max(m, String(s).length), 0);
      const NEEDED_RIGHT = Math.max(16, longestMeta + 2);
      const rShortfall = NEEDED_RIGHT - rightZoneW;
      if (rShortfall > 0) {
        const widenRightCol = ws.getColumn(colCount);
        widenRightCol.width = (widenRightCol.width || 12) + rShortfall;
        colWidths[colCount - 1] += rShortfall;
      }
    }

    // Note: we don't try to perfectly balance the visual width of the left
    // and right sides — that would push columns off the printable page when
    // the data has many columns. Instead we rely on align:center within the
    // merged title cell, which puts the title at the middle of the middle
    // data columns — close enough to page-center for the brand band to look
    // balanced.
  }
  const midStart = leftEnd + 1;
  const rightStart = midEnd + 1;
  const lastCol = colCount;

  // Brand band: 3 rows of content + 1 spacer = 4 rows total.
  // Tall rows (~78pt total) so the merged left cell can fit the company name
  // plus 2 address lines, with each address wrapping to up to 2 lines.
  ws.getRow(1).height = 26;
  ws.getRow(2).height = 26;
  ws.getRow(3).height = 26;
  ws.getRow(4).height = 6;

  // ── Brand block: ONE merged cell spanning cols 1..leftEnd × rows 1-3 ──
  // The logo image is anchored in the leftmost portion of the merged cell;
  // the company name + addresses (richText with line breaks) sit alongside
  // it, indented past the logo. Logo + text live as one visual unit.
  const brandLeftCol = logoCol;     // col 1
  const brandRightCol = leftEnd;    // last column of the left zone
  ws.mergeCells(range(brandLeftCol, 1, brandRightCol, 3));

  if (header.logoPath) {
    try {
      // Detect actual image format from the file's magic bytes — the logo
      // files are sometimes JPEGs masquerading as .png, and Excel's strict
      // parser rejects mismatched extensions (LibreOffice is more lenient).
      const fs = require('fs');
      const path = require('path');
      const buf = fs.readFileSync(header.logoPath);
      let extension = 'png';
      if (buf.length >= 3 && buf[0] === 0xFF && buf[1] === 0xD8 && buf[2] === 0xFF) {
        extension = 'jpeg';
      } else if (buf.length >= 8 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4E && buf[3] === 0x47) {
        extension = 'png';
      } else if (buf.length >= 6 && buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46) {
        extension = 'gif';
      } else {
        // Fallback: use the file extension if magic bytes don't match.
        const ext = path.extname(header.logoPath).toLowerCase();
        if (ext === '.jpg' || ext === '.jpeg') extension = 'jpeg';
      }
      const imgId = wb.addImage({ buffer: buf, extension });
      // Anchor the logo within the leftmost portion of the merged cell:
      // 0-based column 0 → 1, rows 0 → 3. The image overlays the merged cell
      // but the richText below renders to the right of it (indent compensates).
      ws.addImage(imgId, {
        tl: { col: 0.05, row: 0.05 },
        br: { col: 0.95, row: 2.95 },
        editAs: 'oneCell',
      });
    } catch (_) { /* swallow — render without logo if image lookup fails */ }
  }

  // Place name + address as multi-line richText inside the same merged cell.
  // Indent so the text doesn't overlap the logo (~9 chars of indent ≈ 1 col).
  const cell = ws.getCell(colLetter(brandLeftCol) + '1');
  const runs = [];
  runs.push({
    text: header.name || '',
    font: { bold: true, size: 11, color: { argb: 'FF000000' } },
  });
  if (header.address1) {
    runs.push({
      text: '\n' + header.address1,
      font: { size: 8, color: { argb: 'FF555555' } },
    });
  }
  if (header.address2) {
    runs.push({
      text: '\n' + header.address2,
      font: { size: 8, color: { argb: 'FF555555' } },
    });
  }
  cell.value = { richText: runs };
  // wrapText must be enabled for richText \n line breaks to render in
  // Excel/Calc. Indent the text 9 char-units so it sits to the RIGHT of
  // the logo (which occupies col 1 ≈ 9 char-units wide).
  cell.alignment = {
    horizontal: 'left',
    vertical: 'middle',
    wrapText: true,
    indent: 9,
  };

  // ── Middle: report title spanning rows 1-3 of the middle column band ──
  if (title && midStart <= midEnd) {
    ws.mergeCells(range(midStart, 1, midEnd, 3));
    const titleCell = ws.getCell(colLetter(midStart) + '1');
    titleCell.value = title;
    // Pick a title font size that fits the available column width.
    // ExcelJS char-units roughly correspond to "1 unit ≈ 1 character at the
    // sheet's default font". 16pt bold characters are ~1.5x wider, so to fit
    // a title of N characters we need roughly N * 1.5 units. Scale font size
    // down if the available zone is too narrow.
    let midZoneW = 0;
    for (let i = midStart; i <= midEnd; i++) midZoneW += colWidths[i - 1];
    const titleLen = String(title).length;
    const charsPerUnit = midZoneW / Math.max(1, titleLen);
    let titleSize = 16;
    if (charsPerUnit < 1.6) titleSize = 14;
    if (charsPerUnit < 1.4) titleSize = 12;
    if (charsPerUnit < 1.2) titleSize = 11;
    titleCell.font = { bold: true, size: titleSize };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: false };
  }

  // ── Right: meta lines, one per row, right-aligned ──
  if (metaLines.length && rightStart <= lastCol) {
    metaLines.slice(0, 3).forEach((line, i) => {
      ws.mergeCells(range(rightStart, i + 1, lastCol, i + 1));
      const cell = ws.getCell(colLetter(rightStart) + (i + 1));
      cell.value = line;
      cell.font = { bold: true, size: 10 };
      cell.alignment = { horizontal: 'right', vertical: 'middle' };
    });
  }

  return 5;  // first row available for callers (row 4 is the spacer)
}

module.exports = {
  fmtMoney, fmtQty, fmtPrice, fmtIndian,
  formatDateForDisplay,
  getCompanyHeader, drawCompanyHeader,
  xlsxNumFmtForHeader, writeXlsxCompanyHeader,
  runWithCompanyOverride,
};
