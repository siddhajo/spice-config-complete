/**
 * mobile-bridge.js — PWA compatibility layer for the merged app.
 *
 * Mounts the mobile lot-entry UI (PWA's app.html) at `/mobile` and
 * registers the alias/shim endpoints under `/api/auth/*`, `/api/config`,
 * `/api/status`, `/api/logo`, `/api/lots` (query-string filter form),
 * etc. so the mobile UI can talk to spice-config's existing API without
 * the HTML knowing about the rename.
 *
 * Why a bridge instead of patching app.html:
 *   - app.html is a known-good piece of code (1600+ lines). Patching ~30
 *     fetch() callsites inline would be a maintenance burden every time
 *     we re-import the PWA's UI.
 *   - All deltas live in this one file — easy to audit, easy to remove
 *     later if we ever rewrite the mobile UI on spice-config's native API.
 *
 * Pass 1 scope (this file): everything app.html needs to log in, browse
 *   auctions, search/create sellers + banks, and CRUD lots.
 * Pass 2 scope (next iteration): receipt PDF, batch/seller print routes
 *   — currently stubbed with 501 so missing-feature buttons surface a
 *   clear "not yet" message instead of generic network errors.
 */

const path = require('path');
const express = require('express');
const crypto = require('crypto');
const fs = require('fs');
const PDFDocument = require('pdfkit');

// ──────────────────────────────────────────────────────────────────────
// RECEIPT-PRINT HELPERS — ported from PWA server.js's renderer code.
// ──────────────────────────────────────────────────────────────────────
// All field-name remappings live here so the renderer can stay agnostic:
//   spice-config column → PWA renderer expects
//     lots.gross_wt     → lot.gross_weight
//     lots.sample_wt    → lot.sample_weight
//     lots.name         → lot.trader_name (denormalised seller name)
//     lots.ppin         → lot.pin
// The query selects these aliased fields directly so renderSellerReceipt
// (which is a verbatim port) doesn't need to know about the rename.

// Spice-config logo location. The mobile app defaults to the sister
// company (Amazing Spice Park / ASPL) branding, so receipts use the ASP
// logo. resolveLogoPath checks SPICE_DATA_DIR/logos first so cloud-
// persisted uploads aren't shadowed by the bundled default of the same name.
const { resolveLogoPath: _rlpMb } = require('./logo-paths');
function getLogoPath() {
  return _rlpMb('logo-asp.png');
}

// Sensitive-field masking — shared with the rest of the app (screen +
// invoices + reports). Uses the real mask_acct / mask_ifsc / mask_phone
// policy (modes: none/last4/last6/first4/first6/first2last2/full).
const { makeMaskers } = require('./mask-fields');

// Pull the receipt-relevant settings from spice-config's company_settings.
// Field names match what the PWA renderer reads off cfg.
function getReceiptConfig(db) {
  const { modeForKey } = require('./company-config');
  const get = (key, fb = '') => {
    const r = db.get('SELECT value FROM company_settings WHERE key = ? AND business_mode = ?', [key, modeForKey(db, key)]);
    return r ? r.value : fb;
  };
  const getBool = (key, fb = false) => {
    const v = get(key, '');
    if (!v) return fb;
    return v === 'true' || v === '1';
  };
  return {
    appTitle:     get('trade_name', 'Spice Auction'),
    showUser:     getBool('show_username', false),
    // Real masking policy keys (the old code read a non-existent 'acct_mask'
    // with the wrong mode vocabulary, so masking never applied). Defaults
    // match the company-config seed: acct → last4, ifsc/phone → none.
    acctMask:     get('mask_acct',  'last4'),
    ifscMask:     get('mask_ifsc',  'none'),
    phoneMask:    get('mask_phone', 'none'),
    showMoisture: getBool('show_moisture', false),
    sampleWeight: parseFloat(get('sample_weight', '0')) || 0,
    labels:       {},  // spice-config doesn't customize labels; defaults fine
  };
}

// ── HEADER (full size: ~340pt wide) ──────────────────────────────
function addReceiptHeader(doc, appTitle, branch, dateFmt, tradeNo) {
  const w = 300, m = 20;
  const logoPath = getLogoPath();
  if (logoPath) {
    try {
      doc.image(logoPath, (340 - 45) / 2, doc.y, { width: 45, height: 45 });
      doc.y += 50;
    } catch (e) {}
  }
  doc.font('Helvetica-Bold').fontSize(14).text(appTitle, m, doc.y, { width: w, align: 'center' });
  doc.fontSize(10).text((branch || '') + ' BRANCH', m, doc.y, { width: w, align: 'center' });
  doc.moveDown(0.4);
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.5).stroke(); doc.moveDown(0.4);

  doc.font('Helvetica').fontSize(10);
  const y0 = doc.y;
  doc.text('Date: ' + dateFmt, m, y0, { width: w / 2 });
  doc.text('Trade #' + tradeNo, m + w / 2, y0, { width: w / 2, align: 'right' });
  doc.y = y0 + 16;
  doc.moveDown(0.2);
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).dash(3, { space: 3 }).lineWidth(0.5).stroke().undash();
  doc.moveDown(0.4);
}

// ── HEADER (compact: ~180pt wide, thermal-printer friendly) ──────
function addReceiptHeaderCompact(doc, appTitle, branch, dateFmt, tradeNo) {
  const w = 160, m = 10;
  const logoPath = getLogoPath();
  if (logoPath) {
    try {
      doc.image(logoPath, (180 - 28) / 2, doc.y, { width: 28, height: 28 });
      doc.y += 30;
    } catch (e) {}
  }
  doc.font('Helvetica-Bold').fontSize(10).text(appTitle, m, doc.y, { width: w, align: 'center' });
  doc.fontSize(7.5).text((branch || '') + ' BRANCH', m, doc.y, { width: w, align: 'center' });
  doc.moveDown(0.2);
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.4).stroke(); doc.moveDown(0.2);
  doc.font('Helvetica').fontSize(7);
  const y0 = doc.y;
  doc.text('Date: ' + dateFmt, m, y0, { width: w / 2 });
  doc.text('Trade #' + tradeNo, m + w / 2, y0, { width: w / 2, align: 'right' });
  doc.y = y0 + 10;
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).dash(2, { space: 2 }).lineWidth(0.4).stroke().undash();
  doc.moveDown(0.2);
}

// ── RENDERER (full) ──────────────────────────────────────────────
function renderSellerReceipt(doc, sellerLots, cfg) {
  const w = 300, m = 20;
  const lot = sellerLots[0];
  const dateFmt = lot.date ? String(lot.date).split('-').reverse().join('/') : '';
  const L = cfg.labels || {};
  const lb = (k, d) => L[k] || d;
  const headerBranch = cfg.branch || lot.branch;

  addReceiptHeader(doc, cfg.appTitle, headerBranch, dateFmt, lot.ano);

  const lw = 70;
  const { maskAcct, maskIfsc } = makeMaskers({ mask_acct: cfg.acctMask, mask_ifsc: cfg.ifscMask });
  const maskedAcct = lot.acctnum ? maskAcct(lot.acctnum) : '';
  const maskedIfsc = lot.ifsc ? maskIfsc(lot.ifsc) : '';
  const sellerFields = [
    [lb('seller', 'Seller'), lot.trader_name],
    [lb('place',  'Place'),  [lot.ppla, lot.pin].filter(Boolean).join(', ')],
    [lb('gstin',  'GSTIN'),  lot.cr],
    [lb('acct_no','A/C No'), maskedAcct || '--NIL--'],
    [lb('ifsc',   'IFSC'),   maskedIfsc || '--NIL--'],
  ];
  doc.fontSize(9);
  sellerFields.forEach(([label, value]) => {
    if (!value) return;
    const y = doc.y;
    doc.font('Helvetica-Bold').text(label, m, y, { width: lw });
    doc.font('Helvetica').text(String(value), m + lw, y, { width: w - lw });
    if (doc.y < y + 13) doc.y = y + 13;
  });

  doc.moveDown(0.3);
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.5).stroke(); doc.moveDown(0.3);

  const cols = [50, 46, 64, 50, 60];
  const hdrs = [lb('lot_no','Lot#'), lb('bags','Bags'), lb('net_wt','Net'), lb('sample_wt','Smp'), lb('gross_wt','Gross')];
  if (cfg.showMoisture) { cols.push(38); hdrs.push(lb('moisture','Mst%')); }

  const hdrY = doc.y;
  doc.font('Helvetica-Bold').fontSize(7.5);
  let cx = m;
  hdrs.forEach((h, i) => { doc.text(h, cx, hdrY, { width: cols[i], align: 'center' }); cx += cols[i]; });
  doc.y = hdrY + 11;
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.3).stroke(); doc.moveDown(0.2);

  doc.font('Helvetica').fontSize(8);
  let totalQty = 0, totalGross = 0, totalBags = 0, totalSample = 0;
  sellerLots.forEach(l => {
    const ry = doc.y;
    cx = m;
    const sw = Number(l.sample_weight) || cfg.sampleWeight || 0;
    const rowData = [
      l.lot_no,
      l.bags,
      Number(l.qty).toFixed(3),
      sw ? sw.toFixed(3) : '',
      l.gross_weight != null ? Number(l.gross_weight).toFixed(3) : '',
    ];
    if (cfg.showMoisture) rowData.push(l.moisture ? Number(l.moisture).toFixed(1) : '');
    rowData.forEach((v, i) => { doc.text(String(v), cx, ry, { width: cols[i], align: 'center' }); cx += cols[i]; });
    doc.y = ry + 13;
    totalQty    += Number(l.qty) || 0;
    totalGross  += Number(l.gross_weight) || 0;
    totalBags   += Number(l.bags) || 0;
    totalSample += sw;
  });

  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.5).stroke(); doc.moveDown(0.3);
  doc.font('Helvetica-Bold').fontSize(8);
  let totLine = sellerLots.length + ' lot(s) | ' + totalBags + ' ' + lb('bags','bags') +
                ' | ' + lb('net_wt','Net') + ': ' + totalQty.toFixed(3);
  if (totalSample) totLine += ' | ' + lb('sample_wt','Smp') + ': ' + totalSample.toFixed(3);
  if (totalGross)  totLine += ' | ' + lb('gross_wt','Grs') + ': ' + totalGross.toFixed(3);
  doc.text(totLine, m, doc.y, { width: w, align: 'center' });

  doc.moveDown(0.4);
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.5).stroke(); doc.moveDown(0.2);
  if (cfg.showUser) {
    doc.font('Helvetica').fontSize(8).fillColor('#888')
       .text('Entered by: ' + (lot.user_id || ''), m, doc.y, { width: w });
    doc.moveDown(0.2);
  }
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(10)
     .text('** THANK YOU **', m, doc.y, { width: w, align: 'center' });
}

// ── RENDERER (compact, thermal-printer / ~2.5"×3.5") ─────────────
function renderSellerReceiptCompact(doc, sellerLots, cfg) {
  const w = 160, m = 10;
  const lot = sellerLots[0];
  const dateFmt = lot.date ? String(lot.date).split('-').reverse().join('/') : '';
  const L = cfg.labels || {};
  const lb = (k, d) => L[k] || d;
  const headerBranch = cfg.branch || lot.branch;

  addReceiptHeaderCompact(doc, cfg.appTitle, headerBranch, dateFmt, lot.ano);

  const lw = 32;
  const { maskAcct, maskIfsc } = makeMaskers({ mask_acct: cfg.acctMask, mask_ifsc: cfg.ifscMask });
  const maskedAcct = lot.acctnum ? maskAcct(lot.acctnum) : '';
  const maskedIfsc = lot.ifsc ? maskIfsc(lot.ifsc) : '';
  const sellerFields = [
    [lb('seller','Seller'), lot.trader_name],
    [lb('place', 'Place'),  [lot.ppla, lot.pin].filter(Boolean).join(', ')],
    [lb('acct_no','A/C'),   maskedAcct || '--NIL--'],
    [lb('ifsc',  'IFSC'),   maskedIfsc || '--NIL--'],
  ];
  doc.fontSize(7);
  sellerFields.forEach(([label, value]) => {
    if (!value) return;
    const y = doc.y;
    doc.font('Helvetica-Bold').text(label, m, y, { width: lw });
    doc.font('Helvetica').text(String(value), m + lw, y, { width: w - lw });
    if (doc.y < y + 10) doc.y = y + 10;
  });

  doc.moveDown(0.2);
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.4).stroke(); doc.moveDown(0.2);

  const cols = [28, 28, 50, 54];
  const hdrs = [lb('lot_no','Lot#'), lb('bags','Bags'), lb('net_wt','Net'), lb('gross_wt','Gross')];

  const hdrY = doc.y;
  doc.font('Helvetica-Bold').fontSize(6.5);
  let cx = m;
  hdrs.forEach((h, i) => { doc.text(h, cx, hdrY, { width: cols[i], align: 'center' }); cx += cols[i]; });
  doc.y = hdrY + 9;
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.3).stroke(); doc.moveDown(0.15);

  doc.font('Helvetica').fontSize(7);
  let totalQty = 0, totalGross = 0, totalBags = 0;
  sellerLots.forEach(l => {
    const ry = doc.y;
    cx = m;
    const rowData = [
      l.lot_no, l.bags,
      Number(l.qty).toFixed(3),
      l.gross_weight != null ? Number(l.gross_weight).toFixed(3) : '',
    ];
    rowData.forEach((v, i) => { doc.text(String(v), cx, ry, { width: cols[i], align: 'center' }); cx += cols[i]; });
    doc.y = ry + 11;
    totalQty   += Number(l.qty) || 0;
    totalGross += Number(l.gross_weight) || 0;
    totalBags  += Number(l.bags) || 0;
  });

  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.4).stroke(); doc.moveDown(0.15);

  const sumCols = [40, 40, 40, 40];
  const sumHdrs = ['Lots', lb('bags','Bags'), lb('net_wt','Net'), lb('gross_wt','Gross')];
  const sumVals = [String(sellerLots.length), String(totalBags), totalQty.toFixed(3),
                   totalGross ? totalGross.toFixed(3) : '-'];
  const sHdrY = doc.y;
  doc.font('Helvetica-Bold').fontSize(6.5);
  let sx = m;
  sumHdrs.forEach((h, i) => { doc.text(h, sx, sHdrY, { width: sumCols[i], align: 'center' }); sx += sumCols[i]; });
  doc.y = sHdrY + 9;
  const sValY = doc.y;
  doc.font('Helvetica-Bold').fontSize(8.5);
  sx = m;
  sumVals.forEach((v, i) => { doc.text(v, sx, sValY, { width: sumCols[i], align: 'center' }); sx += sumCols[i]; });
  doc.y = sValY + 12;

  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.4).stroke(); doc.moveDown(0.2);
  if (cfg.showUser) {
    doc.font('Helvetica').fontSize(6).fillColor('#888')
       .text('Entered by: ' + (lot.user_id || ''), m, doc.y, { width: w });
    doc.moveDown(0.15);
  }
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(9)
     .text('** THANK YOU **', m, doc.y, { width: w, align: 'center' });
}

function pickReceiptRenderer(fmt) {
  return fmt === 'compact'
    ? { render: renderSellerReceiptCompact, pageSize: [180, 252], compact: true }
    : { render: renderSellerReceipt,        pageSize: [340, 550], compact: false };
}

// ── LOT SELECT — single helper used by every print endpoint ─────
// Spice-config has denormalised seller fields on the lots row
// (lots.name, lots.cr, lots.ppla, lots.ppin, lots.tel) so we don't
// strictly need the traders join, BUT joining gives us the freshest
// values when the seller's master record has changed since the lot was
// booked. We also pick the bank from trader_banks (per-lot bank_id pin
// > seller's default), falling back to lots.acctnum/ifsc for legacy data.
// Output columns are aliased to the names PWA's renderer expects.
const LOT_SELECT_SQL = `
  SELECT
    l.id, l.lot_no, l.branch, l.bags, l.litre, l.qty,
    l.gross_wt  AS gross_weight,
    l.sample_wt AS sample_weight,
    l.moisture, l.user_id, l.trader_id,
    COALESCE(t.name, l.name, 'Unknown') AS trader_name,
    COALESCE(t.cr,   l.cr,   '') AS cr,
    COALESCE(t.ppla, l.ppla, '') AS ppla,
    COALESCE(t.pin,  l.ppin, '') AS pin,
    COALESCE(
      (SELECT tb.acctnum FROM trader_banks tb WHERE tb.id = l.bank_id),
      (SELECT tb.acctnum FROM trader_banks tb WHERE tb.trader_id = t.id ORDER BY tb.is_default DESC, tb.id LIMIT 1),
      t.acctnum, l.cr, ''
    ) AS acctnum,
    COALESCE(
      (SELECT tb.ifsc FROM trader_banks tb WHERE tb.id = l.bank_id),
      (SELECT tb.ifsc FROM trader_banks tb WHERE tb.trader_id = t.id ORDER BY tb.is_default DESC, tb.id LIMIT 1),
      t.ifsc, ''
    ) AS ifsc,
    a.ano, a.date, a.crop_type
  FROM lots l
  JOIN auctions a ON a.id = l.auction_id
  LEFT JOIN traders t ON t.id = l.trader_id
`;

function mountMobile(app, deps) {
  const { getDb, requireAuth, verifyPassword, hashPassword, isLegacyHash, ROLE_PERMISSIONS } = deps;
  // auditLog is injected from server.js so mobile-side writes (e.g. a
  // field user clearing their own lots) land in the same activity feed
  // the admin console reads. No-op shim if an older host didn't pass it.
  const auditLog = deps.auditLog || function () {};

  // ── 0. LAZY SELF-HEAL SCHEMA (defence in depth) ───────────────────
  // db.js is now the canonical source for every column/table the bridge
  // needs (login_history, traders.whatsapp/email, lots.bank_id,
  // users.must_change_password). These ALTERs are kept as a safety net
  // so the bridge still functions if someone ever loads it against an
  // older db.js (forked install, partial upgrade) — every statement is
  // idempotent and silently no-ops once the columns already exist.
  // Runs on the first /api/* request rather than at mount time because
  // initDb() finishes after mountMobile is called.
  let _healed = false;
  function ensureBridgeSchema() {
    if (_healed) return;
    try {
      const db = getDb();
      db.exec(`CREATE TABLE IF NOT EXISTS login_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL,
        username TEXT NOT NULL,
        ip TEXT DEFAULT '',
        user_agent TEXT DEFAULT '',
        created_at TEXT DEFAULT (datetime('now','localtime'))
      )`);
      try { db.exec("ALTER TABLE traders ADD COLUMN whatsapp TEXT DEFAULT ''"); } catch (_) {}
      try { db.exec("ALTER TABLE traders ADD COLUMN email TEXT DEFAULT ''"); } catch (_) {}
      try { db.exec("ALTER TABLE lots ADD COLUMN bank_id INTEGER"); } catch (_) {}
      try { db.exec("ALTER TABLE users ADD COLUMN must_change_password INTEGER DEFAULT 0"); } catch (_) {}
      _healed = true;
    } catch (e) {
      console.warn('[mobile-bridge] self-heal deferred:', e.message);
    }
  }
  // Express middleware: heal once before the first /api/* request.
  app.use('/api', (req, _res, next) => { ensureBridgeSchema(); next(); });

  // ── 0b. requireAuthFlex — accepts Authorization header OR ?token= ─
  // Print URLs are opened via window.open() (the only way mobile browsers
  // give us a real print dialog with a renderable PDF preview). window.open
  // can NOT set the Authorization header, so the mobile UI appends the
  // token as a query string. Spice-config's native requireAuth only reads
  // the header, so query-string tokens 401 there. This helper closes that
  // gap for the print routes ONLY — every other route stays on the strict
  // header-only requireAuth.
  function requireAuthFlex(req, res, next) {
    const hdr = (req.headers.authorization || '').replace('Bearer ', '');
    const tok = hdr || String(req.query.token || '');
    if (!tok) return res.status(401).json({ error: 'No token' });
    const db = getDb();
    const session = db.get('SELECT * FROM sessions WHERE token = ?', [tok]);
    if (!session) return res.status(403).json({ error: 'Session expired — please sign in again' });
    const user = db.get('SELECT * FROM users WHERE id = ?', [session.user_id]);
    if (!user) return res.status(403).json({ error: 'Unauthorized' });
    db.run(`UPDATE sessions SET last_used_at = datetime('now','localtime') WHERE token = ?`, [tok]);
    req.user = user;
    req.session = session;
    next();
  }

  // ── 1. STATIC MOUNT ──────────────────────────────────────────────
  // Serves /mobile, /mobile/app.html, /mobile/manifest.json, /mobile/icon.svg.
  // Phones will install the PWA from /mobile/ (manifest scope = /mobile/).
  const mobileDir = path.join(__dirname, 'public-mobile');
  // The explicit route MUST be registered BEFORE express.static. Otherwise
  // static() auto-redirects `/mobile` → `/mobile/` with a 301 (its built-in
  // directory-handling) before our app.get gets a chance to run.
  app.get('/mobile', (_req, res) => res.sendFile(path.join(mobileDir, 'app.html')));
  app.use('/mobile', express.static(mobileDir, { maxAge: 0 }));

  // ── 2. AUTH ALIASES ─────────────────────────────────────────────
  // PWA uses /api/auth/* paths. spice-config uses /api/login etc.
  // We wrap login/me so the response shape matches what app.html expects
  // ({user: {...}, token} rather than {token, role, username}).

  app.post('/api/auth/login', async (req, res) => {
    const { username, password } = req.body || {};
    if (!username || !password) {
      return res.status(400).json({ error: 'Username and password required' });
    }
    const db = getDb();
    const user = db.get('SELECT * FROM users WHERE username = ?', [username]);
    // Bcrypt verify — matches the desktop admin's /api/login flow.
    // verifyPassword tolerates both legacy SHA-256 rows AND bcrypt rows
    // so a user who logged in via the desktop and got their hash
    // upgraded to bcrypt earlier can still log in here.
    const ok = user ? await verifyPassword(password, user.password_hash) : false;
    if (!user || !ok) {
      return res.status(401).json({ error: 'Invalid username or password' });
    }
    // Opportunistic rehash to bcrypt if the stored hash is legacy SHA-256.
    // Same logic as desktop /api/login — first successful mobile login
    // also upgrades the row.
    if (isLegacyHash(user.password_hash)) {
      try {
        const upgraded = await hashPassword(password);
        db.run('UPDATE users SET password_hash = ? WHERE id = ?', [upgraded, user.id]);
      } catch (_) { /* non-fatal */ }
    }
    const token = crypto.randomBytes(32).toString('hex');
    // Multi-device sessions — DON'T delete existing sessions. Field staff
    // can keep the desktop admin UI logged in while using the phone too.
    db.run(
      'INSERT INTO sessions (token, user_id, device_label) VALUES (?, ?, ?)',
      [token, user.id, (req.headers['user-agent'] || '').slice(0, 80)]
    );
    db.run(
      'INSERT INTO login_history (user_id, username, ip, user_agent) VALUES (?, ?, ?, ?)',
      [
        user.id,
        user.username,
        req.headers['x-forwarded-for'] || req.connection.remoteAddress || '',
        /Mobile|Android|iPhone/i.test(req.headers['user-agent'] || '') ? 'Mobile' : 'Desktop',
      ]
    );
    res.json({
      user: {
        id: user.id,
        username: user.username,
        role: user.role,
        branch: user.branch || '',
      },
      token,
    });
  });

  app.post('/api/auth/logout', (req, res) => {
    const t = (req.headers.authorization || '').replace('Bearer ', '');
    if (t) getDb().run('DELETE FROM sessions WHERE token = ?', [t]);
    res.json({ success: true });
  });

  app.get('/api/auth/me', requireAuth, (req, res) => {
    res.json({
      user: {
        id: req.user.id,
        username: req.user.username,
        role: req.user.role,
        branch: req.user.branch || '',
      },
    });
  });

  app.post('/api/auth/change-password', requireAuth, async (req, res) => {
    const { current_password, new_password } = req.body || {};
    if (!current_password || !new_password) {
      return res.status(400).json({ error: 'Both current and new password required' });
    }
    if (new_password.length < 4) {
      return res.status(400).json({ error: 'New password must be at least 4 characters' });
    }
    const db = getDb();
    const user = db.get('SELECT * FROM users WHERE id = ?', [req.user.id]);
    // Same bcrypt verify as desktop /api/me/password — accepts legacy
    // SHA-256 rows so the first password change after upgrade still works.
    const ok = user ? await verifyPassword(current_password, user.password_hash) : false;
    if (!user || !ok) {
      return res.status(401).json({ error: 'Current password is incorrect' });
    }
    const newHash = await hashPassword(new_password);
    // Clear must_change_password alongside the hash update — without
    // this the user stays gated behind the forced-change wall even
    // after a successful change. The desktop endpoint (/api/me/password
    // in server.js) does the same in one UPDATE.
    db.run(
      'UPDATE users SET password_hash = ?, must_change_password = 0 WHERE id = ?',
      [newHash, user.id]
    );
    // Kill all OTHER sessions of this user
    db.run('DELETE FROM sessions WHERE user_id = ? AND token != ?', [user.id, req.session.token]);
    res.json({ success: true });
  });

  // ── 3. CONFIG SHIM (branches / crop types / title / settings) ──
  // The PWA's app.html hits GET /api/config (no query string) on every
  // login + session start, and expects a SINGLE FLAT OBJECT with every
  // setting it cares about. Build that shape from spice-config's
  // company_settings table.
  //
  // PWA-expected fields:
  //   branches[]      → br1..br9 (skip blanks)
  //   cropTypes[]     → default_crop_type + sensible fallbacks
  //   title           → trade_name
  //   sampleWeight    → sample_weight (lot_entry category)
  //   showMoisture    → show_moisture
  //   defaultLitre    → default_litre
  //   editEnabled     → edit_enabled (boolean)
  //   editTimeout     → edit_timeout_sec
  //   labels{}        → reserved; safe to leave empty (PWA has defaults)
  //   pageLimit, showUsername, tradeTileTitle, acctMask → defaults
  app.get('/api/config', (req, res) => {
    const db = getDb();
    const type = String(req.query.type || '').toLowerCase();

    // Helper: read a single setting (scoped to the active mode; globals use '*')
    const { modeForKey } = require('./company-config');
    const get = (key, fallback = '') => {
      const r = db.get(`SELECT value FROM company_settings WHERE key = ? AND business_mode = ?`, [key, modeForKey(db, key)]);
      return r ? r.value : fallback;
    };
    const getNum  = (key, fallback = 0) => { const v = parseFloat(get(key, '')); return isNaN(v) ? fallback : v; };
    const getBool = (key, fallback = false) => {
      const v = get(key, '');
      if (v === '' || v == null) return fallback;
      return v === 'true' || v === '1';
    };

    // Branch list — keys br1..br9, blank values dropped (active mode only)
    const brRows = db.all(
      `SELECT key, value FROM company_settings
       WHERE category = 'branches' AND key LIKE 'br_' AND business_mode = ?
       ORDER BY key`,
      [require('./company-config').getActiveMode(db)]
    );
    const branches = brRows
      .filter(r => r.value && String(r.value).trim())
      .map((r, i) => ({
        id: i + 1, type: 'branch',
        value: String(r.value).trim().toUpperCase(),
        sort_order: i,
      }));

    // Crop types — synthesise from default_crop_type + fallbacks
    const defCrop = String(get('default_crop_type', '')).trim().toUpperCase();
    const cropSet = new Set();
    if (defCrop) cropSet.add(defCrop);
    ['ASP', 'VST'].forEach(c => cropSet.add(c));
    const cropTypes = Array.from(cropSet).map((v, i) => ({
      id: i + 1, type: 'crop_type', value: v, sort_order: i,
    }));

    // Single-type shorthand: PWA's original /api/config?type=branch
    // returned { items: [...] }. Preserve that for any caller that uses it.
    if (type === 'branch')    return res.json({ items: branches });
    if (type === 'crop_type') return res.json({ items: cropTypes });
    if (type === 'title')     return res.json({ items: [{ id: 1, type: 'title', value: get('s_short_name', 'AMAZING SPICE PARK PVT LTD'), sort_order: 0 }] });

    // e-Auction-only fields. Both inputs are AND-gated by the flag
    // AND business_mode === 'e-Auction' so flipping the mode away
    // hides the inputs even if an admin leaves the flag on.
    const isEAuc = String(get('business_mode', 'e-Auction')).toLowerCase() === 'e-auction';
    const showCropReceipt   = isEAuc && getBool('flag_crop_receipt',   false);
    const showReservedPrice = isEAuc && getBool('flag_reserved_price', false);

    // Full config object — what app.html expects from a no-arg GET.
    res.json({
      branches,
      cropTypes,
      // Mobile app defaults to the sister company (ASPL) identity. Prefer
      // the ASP short name from settings; fall back to the brand string.
      title:           get('s_short_name', 'AMAZING SPICE PARK PVT LTD'),
      editTimeout:     parseInt(get('edit_timeout_sec', '0'), 10) || 0,
      editEnabled:     getBool('edit_enabled', true),
      sampleWeight:    getNum('sample_weight', 0),
      showMoisture:    getBool('show_moisture', false),
      defaultLitre:    get('default_litre', ''),
      showCropReceipt,
      showReservedPrice,
      // Gunny tare default + unified extra-fields toggle. The mobile UI
      // uses gunnyWeight to derive Net Wt (= Weight-w/-Gunny − gunny ×
      // bags) and showExtraLotFields to reveal the Weight-w/-Gunny +
      // Gunny Wt pair AND the Crop Receipt / Reserved Price inputs.
      gunnyWeight:        getNum('gunny_weight', 0),
      showExtraLotFields: getBool('show_extra_lot_fields', false),
      // PWA defaults — surfaced here for completeness; not currently
      // backed by spice-config settings, so static-ish values are fine.
      pageLimit:       20,
      showUsername:    false,
      tradeTileTitle:  'Active Trade',
      acctMask:        'none',
      labels:          {},
    });
  });

  // ── 4. AUCTIONS ENVELOPE ────────────────────────────────────────
  // PWA's app.html does `const trades = d.auctions || [];` so it expects
  // an envelope object, not the flat array spice-config returns natively.
  // Wrap the native array in {auctions: [...]} for the mobile client.
  //
  // Mode filter (mirrors the desktop's /api/auctions): rows whose
  // `mode` column matches the current company_settings.business_mode
  // setting, OR rows whose mode is NULL/empty (legacy data pre-mode
  // column, kept visible during soft cutover). Without this, switching
  // a phone-toting field user's install from e-Auction to e-Trade
  // doesn't hide the old e-Auction trades from the picker.
  //
  // Errors are caught and surfaced as JSON so the mobile client sees
  // an actual diagnostic instead of a generic "Failed to load".
  app.get('/api/mobile/auctions', requireAuth, (_req, res) => {
    try {
      const db = getDb();
      const modeRow = db.get("SELECT value FROM company_settings WHERE key = 'business_mode' AND business_mode = '*'");
      const mode = modeRow && modeRow.value ? String(modeRow.value) : '';
      const whereSql = mode
        ? `WHERE (auctions.mode = ? OR auctions.mode IS NULL OR auctions.mode = '')`
        : '';
      const params = mode ? [mode] : [];
      const rows = db.all(
        `SELECT *, (SELECT COUNT(*) FROM lots WHERE auction_id=auctions.id) AS lot_count
         FROM auctions ${whereSql}
         ORDER BY date DESC, ano DESC LIMIT 100`,
        params
      );
      res.json({ auctions: rows, businessMode: mode });
    } catch (e) {
      console.error('[/api/mobile/auctions] failed:', e && (e.stack || e.message || e));
      res.status(500).json({ error: e && (e.message || String(e)) || 'Failed to load auctions' });
    }
  });

  // ── 3b. ACTIVE TRADE (shared with the admin app) ────────────────
  // The admin app's topbar trade picker pushes its current selection to
  // company_settings.active_auction_id (POST /api/active-auction, admin
  // side). The mobile app reads it here and defaults to that trade
  // instead of showing its own trade selector. Returns null when nothing
  // is selected yet, so the mobile app falls back to manual selection.
  app.get('/api/active-auction', requireAuth, (_req, res) => {
    try {
      // Read from the dedicated app_state table (written by the admin
      // side's POST /api/active-auction). The table may not exist yet if
      // the admin has never picked a trade — treat that as "none".
      const db = getDb();
      const row = db.get("SELECT value FROM app_state WHERE key = 'active_auction_id'");
      const aid = row && row.value ? parseInt(row.value, 10) : null;
      res.json({ auctionId: Number.isFinite(aid) && aid > 0 ? aid : null });
    } catch (e) {
      res.json({ auctionId: null });
    }
  });

  // ── 3c. FAVOURITE TRADE (mobile default) ────────────────────────
  // The mobile lot-entry app defaults to the admin's "favourite" trade,
  // set via the star on the desktop Trades screen (POST /api/favourite-trade).
  // This is stored under app_state.favourite_auction_id — SEPARATE from
  // active_auction_id — so changing the desktop global trade selector does
  // NOT change the trade the mobile app shows. Returns null when no
  // favourite is set, so the mobile app falls back to its own picker.
  app.get('/api/mobile/favourite-trade', requireAuth, (_req, res) => {
    try {
      const db = getDb();
      const row = db.get("SELECT value FROM app_state WHERE key = 'favourite_auction_id'");
      const aid = row && row.value ? parseInt(row.value, 10) : null;
      res.json({ auctionId: Number.isFinite(aid) && aid > 0 ? aid : null });
    } catch (e) {
      res.json({ auctionId: null });
    }
  });

  // ── 4. STATUS ALIAS ─────────────────────────────────────────────
  // PWA's app.html pings /api/status on boot to detect "logged out vs
  // server unreachable". spice-config has /api/health; alias it.
  app.get('/api/status', (_req, res) => res.json({ ok: true, ts: Date.now() }));

  // ── 5. LOGO ALIAS ───────────────────────────────────────────────
  // PWA loads the brand logo via `<img src="/api/logo">`, which means
  // the response MUST be a raw image, not JSON. The mobile app defaults
  // to the sister company (ASPL) branding, so redirect to /logo-asp.png
  // — served statically from public/ (and overridable via an admin logo
  // upload for 'asp'). The previous redirect to /api/branding returned
  // JSON, which the <img> tag couldn't decode and showed as a broken
  // image placeholder.
  app.get('/api/logo', (req, res) => {
    res.redirect(302, '/logo-asp.png');
  });

  // ── 6. LOTS — query-string filter form ──────────────────────────
  // PWA: GET /api/lots?auction_id=N&branch=X (returns {lots, stats}).
  // Spice-config: GET /api/lots/:auctionId  (returns flat array).
  // We add a new endpoint at the PWA path that reshapes to PWA's expected
  // {lots, stats} envelope. Spice-config's own UI keeps using the path
  // form unchanged.
  app.get('/api/lots', requireAuth, (req, res) => {
    const { auction_id, branch, user_id, seller, page, limit } = req.query;
    if (!auction_id) return res.status(400).json({ error: 'auction_id is required' });
    const db = getDb();
    let where = 'l.auction_id = ?';
    const params = [parseInt(auction_id, 10)];
    if (branch)  { where += ' AND l.branch = ?'; params.push(branch); }
    if (user_id) { where += ' AND l.user_id = ?'; params.push(user_id); }
    if (seller)  {
      // Match against fresh master data first, falling back to the
      // denormalised name on the lot row for legacy lots without a
      // trader_id back-reference.
      where += ' AND (COALESCE(t.name, l.name, "") LIKE ? COLLATE NOCASE)';
      params.push(`%${seller}%`);
    }

    const stats = db.get(
      `SELECT COUNT(*) AS lot_count,
              COALESCE(SUM(l.qty), 0)  AS total_qty,
              COALESCE(SUM(l.bags), 0) AS total_bags
       FROM lots l
       LEFT JOIN traders t ON t.id = l.trader_id
       WHERE ${where}`,
      params
    ) || { lot_count: 0, total_qty: 0, total_bags: 0 };

    // Pagination (opt-in)
    const pageNum = Math.max(1, parseInt(page, 10) || 1);
    const hasLimit = limit !== undefined && limit !== '' && limit !== null;
    const pageSize = hasLimit ? Math.min(100, Math.max(1, parseInt(limit, 10))) : 0;

    let q = `SELECT l.*,
                    COALESCE(t.name, l.name, 'Unknown Trader') AS trader_name,
                    COALESCE(t.cr,   l.cr,   '') AS cr,
                    COALESCE(t.pan,  l.pan,  '') AS pan,
                    COALESCE(t.ppla, l.ppla, '') AS ppla,
                    COALESCE(t.pin,  l.ppin, '') AS pin,
                    COALESCE(t.tel,  l.tel,  '') AS tel,
                    (SELECT tb.acctnum FROM trader_banks tb
                       WHERE tb.id = l.bank_id) AS lot_bank_acctnum,
                    (SELECT tb.ifsc FROM trader_banks tb
                       WHERE tb.id = l.bank_id) AS lot_bank_ifsc,
                    (SELECT tb.acctnum FROM trader_banks tb
                       WHERE tb.trader_id = l.trader_id
                       ORDER BY tb.is_default DESC, tb.id ASC LIMIT 1) AS def_acctnum,
                    (SELECT tb.ifsc FROM trader_banks tb
                       WHERE tb.trader_id = l.trader_id
                       ORDER BY tb.is_default DESC, tb.id ASC LIMIT 1) AS def_ifsc
             FROM lots l
             LEFT JOIN traders t ON t.id = l.trader_id
             WHERE ${where}
             ORDER BY CAST(l.lot_no AS INTEGER) ASC, l.lot_no ASC`;
    const qParams = [...params];
    if (pageSize > 0) {
      q += ' LIMIT ? OFFSET ?';
      qParams.push(pageSize, (pageNum - 1) * pageSize);
    }
    const lots = db.all(q, qParams).map(r => ({
      ...r,
      // Normalize the bank columns the PWA expects:
      acctnum: r.lot_bank_acctnum || r.def_acctnum || '',
      ifsc:    r.lot_bank_ifsc    || r.def_ifsc    || '',
      // PWA reads gross_weight/sample_weight; spice-config stores gross_wt/sample_wt
      gross_weight:  r.gross_wt  || null,
      sample_weight: r.sample_wt || 0,
    }));

    const users = db.all(
      'SELECT DISTINCT user_id FROM lots WHERE auction_id = ? ORDER BY user_id',
      [parseInt(auction_id, 10)]
    );
    const branches = db.all(
      'SELECT DISTINCT branch FROM lots WHERE auction_id = ? ORDER BY branch',
      [parseInt(auction_id, 10)]
    );
    const totalPages = pageSize > 0 ? Math.ceil(stats.lot_count / pageSize) : 1;
    res.json({
      lots,
      stats: {
        lotCount:  stats.lot_count,
        totalQty:  Math.round(stats.total_qty * 1000) / 1000,
        totalBags: stats.total_bags,
      },
      pagination: { page: pageNum, totalPages, total: stats.lot_count, pageSize },
      filters: {
        users: users.map(u => u.user_id).filter(Boolean),
        branches: branches.map(b => b.branch).filter(Boolean),
      },
    });
  });

  // ── 7. LOT DETAIL (for edit modal) ──────────────────────────────
  // PWA pre-fills the edit form's gross/sample weights from this endpoint.
  app.get('/api/lots/:id/detail', requireAuth, (req, res) => {
    const db = getDb();
    const id = parseInt(req.params.id, 10);
    const row = db.get('SELECT * FROM lots WHERE id = ?', [id]);
    if (!row) return res.status(404).json({ error: 'Lot not found' });
    res.json({
      lot: {
        ...row,
        // Field-name mapping back to PWA's expected names
        gross_weight:  row.gross_wt  || null,
        sample_weight: row.sample_wt || 0,
      },
    });
  });

  // ── 8. TRADER QUICK-CREATE (PWA POST /api/traders) ─────────────
  // Single unified seller-create path used by BOTH apps. Mobile PWA hits
  // /api/traders directly; desktop's "Add Seller" buttons also reach here
  // (the bridge mounts before the native /api/traders POST, so the bridge
  // wins the route).
  //
  // Strong uniqueness:
  //   - GSTIN (cr) — if present, must be unique across all traders
  //   - PAN        — if present, must be unique
  //   - Phone (tel) + Name — same combo treated as a duplicate
  // These checks run BEFORE insert so neither app can race two creates
  // for the same person.
  app.post('/api/traders', requireAuth, (req, res) => {
    const t = req.body || {};
    if (!t.name || !String(t.name).trim()) {
      return res.status(400).json({ error: 'Seller name is required' });
    }
    const emailClean = (t.email || '').toString().trim();
    if (emailClean && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailClean)) {
      return res.status(400).json({ error: 'Invalid email format' });
    }
    const db = getDb();
    const nameTrim = String(t.name).trim().toUpperCase();
    const crTrim   = String(t.cr  || '').trim();
    const panTrim  = String(t.pan || '').trim().toUpperCase();
    const telTrim  = String(t.tel || '').trim();

    // Strict uniqueness — GSTIN (cr) is the strongest identifier
    if (crTrim) {
      const dup = db.get('SELECT * FROM traders WHERE cr = ? COLLATE NOCASE LIMIT 1', [crTrim]);
      if (dup) {
        dup.banks = db.all(
          'SELECT * FROM trader_banks WHERE trader_id = ? ORDER BY is_default DESC, id', [dup.id]
        );
        return res.json({ trader: dup, deduped: true, reason: 'GSTIN match' });
      }
    }
    // PAN is the next strongest
    if (panTrim) {
      const dup = db.get('SELECT * FROM traders WHERE pan = ? COLLATE NOCASE LIMIT 1', [panTrim]);
      if (dup) {
        dup.banks = db.all(
          'SELECT * FROM trader_banks WHERE trader_id = ? ORDER BY is_default DESC, id', [dup.id]
        );
        return res.json({ trader: dup, deduped: true, reason: 'PAN match' });
      }
    }
    // Soft dedup — same name + same phone is treated as a single person
    if (telTrim) {
      const dup = db.get('SELECT * FROM traders WHERE name = ? AND tel = ? LIMIT 1',
        [nameTrim, telTrim]);
      if (dup) {
        dup.banks = db.all(
          'SELECT * FROM trader_banks WHERE trader_id = ? ORDER BY is_default DESC, id', [dup.id]
        );
        return res.json({ trader: dup, deduped: true, reason: 'name+phone match' });
      }
    }

    const info = db.run(
      `INSERT INTO traders (name,cr,pan,tel,aadhar,padd,ppla,pin,pstate,pst_code,ifsc,acctnum,holder_name,whatsapp,email)
       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
      [
        nameTrim,
        crTrim,
        panTrim,
        telTrim,
        (t.aadhar || '').toString().trim(),
        (t.padd || '').toString().trim(),
        (t.ppla || '').toString().trim().toUpperCase(),
        (t.pin || '').toString().trim(),
        (t.pstate || 'TAMIL NADU').toString().trim().toUpperCase(),
        (t.pst_code || '33').toString().trim(),
        '', '', '',
        (t.whatsapp || '').toString().trim(),
        emailClean,
      ]
    );
    const newId = info.lastInsertRowid;
    // Desktop UI sends `banks` as part of the trader payload. Persist
    // them now so the response carries the populated array. PWA omits
    // `banks` here and writes them via /api/traders/:id/banks instead,
    // so the no-op branch in the helper covers that path.
    syncTraderBanksFromArray(db, newId, t.banks);
    const created = db.get('SELECT * FROM traders WHERE id = ?', [newId]);
    if (created) {
      created.banks = db.all(
        'SELECT * FROM trader_banks WHERE trader_id = ? ORDER BY is_default DESC, id', [newId]
      );
    }
    res.status(201).json({ trader: created });
  });

  // ── 8b. TRADER UPDATE (PWA PUT /api/traders/:id) ───────────────
  // Mobile PWA hits this to update whatsapp/email/contact fields. Desktop
  // also hits the same path. Single write path = single source of truth.
  //
  // The bridge handles whatsapp + email (which spice-config's native PUT
  // doesn't know about) and delegates everything else to a full UPDATE
  // covering all editable fields. Uniqueness re-checked on cr/pan changes
  // so an edit can't introduce a duplicate either.
  app.put('/api/traders/:id', requireAuth, (req, res) => {
    const db = getDb();
    const id = parseInt(req.params.id, 10);
    const trader = db.get('SELECT * FROM traders WHERE id = ?', [id]);
    if (!trader) return res.status(404).json({ error: 'Seller not found' });
    const t = req.body || {};
    const emailClean = t.email != null ? String(t.email).trim() : null;
    if (emailClean && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailClean)) {
      return res.status(400).json({ error: 'Invalid email format' });
    }
    // Uniqueness re-check on cr / pan changes
    if (t.cr != null && String(t.cr).trim() && String(t.cr).trim() !== trader.cr) {
      const dup = db.get('SELECT id FROM traders WHERE cr = ? COLLATE NOCASE AND id != ?',
        [String(t.cr).trim(), id]);
      if (dup) return res.status(409).json({ error: 'Another seller already has this GSTIN' });
    }
    if (t.pan != null && String(t.pan).trim() && String(t.pan).trim().toUpperCase() !== trader.pan) {
      const dup = db.get('SELECT id FROM traders WHERE pan = ? COLLATE NOCASE AND id != ?',
        [String(t.pan).trim().toUpperCase(), id]);
      if (dup) return res.status(409).json({ error: 'Another seller already has this PAN' });
    }
    // Partial update — only write fields that were sent. Mobile sends a
    // subset (just acctnum/ifsc/whatsapp/email on edit-from-banks);
    // desktop sends the full record.
    const sets = []; const vals = [];
    const setField = (col, val, transform = (v) => v) => {
      if (val !== undefined) { sets.push(col + ' = ?'); vals.push(transform(val)); }
    };
    setField('name',        t.name,        (v) => String(v).trim().toUpperCase());
    setField('cr',          t.cr,          (v) => String(v).trim());
    setField('pan',         t.pan,         (v) => String(v).trim().toUpperCase());
    setField('tel',         t.tel,         (v) => String(v).trim());
    setField('aadhar',      t.aadhar,      (v) => String(v).trim());
    setField('padd',        t.padd,        (v) => String(v).trim());
    setField('ppla',        t.ppla,        (v) => String(v).trim().toUpperCase());
    setField('pin',         t.pin,         (v) => String(v).trim());
    setField('pstate',      t.pstate,      (v) => String(v).trim().toUpperCase());
    setField('pst_code',    t.pst_code,    (v) => String(v).trim());
    setField('ifsc',        t.ifsc,        (v) => String(v).trim().toUpperCase());
    setField('acctnum',     t.acctnum,     (v) => String(v).trim());
    setField('holder_name', t.holder_name, (v) => String(v).trim());
    setField('whatsapp',    t.whatsapp,    (v) => String(v).trim());
    if (emailClean !== null) { sets.push('email = ?'); vals.push(emailClean); }
    // No flat-field changes is fine — we may still have a `banks` array
    // to sync below. Only short-circuit when neither flat fields nor
    // banks were sent.
    if (sets.length > 0) {
      vals.push(id);
      db.run(`UPDATE traders SET ${sets.join(', ')} WHERE id = ?`, vals);
    } else if (!Array.isArray(t.banks)) {
      return res.json({ success: true, noop: true, trader });
    }
    // Desktop sends the whole banks array on every edit; persist it
    // here so the bridge handler (which wins route-matching over
    // server.js) keeps trader_banks in sync. Mobile PWA omits `banks`
    // and edits rows individually via /api/traders/:id/banks, so the
    // no-op branch in the helper covers that path.
    syncTraderBanksFromArray(db, id, t.banks);
    const updated = db.get('SELECT * FROM traders WHERE id = ?', [id]);
    updated.banks = db.all(
      'SELECT * FROM trader_banks WHERE trader_id = ? ORDER BY is_default DESC, id', [id]
    );
    res.json({ success: true, trader: updated });
  });

  // ── 8c. TRADER GET BY ID — ensures fresh fetch ──────────────────
  // Mobile uses this after edits to refresh the displayed trader. Always
  // reads from the DB (no cache); both apps see the same data.
  app.get('/api/traders/:id', requireAuth, (req, res) => {
    const db = getDb();
    const id = parseInt(req.params.id, 10);
    const trader = db.get('SELECT * FROM traders WHERE id = ?', [id]);
    if (!trader) return res.status(404).json({ error: 'Seller not found' });
    trader.banks = db.all(
      'SELECT * FROM trader_banks WHERE trader_id = ? ORDER BY is_default DESC, id', [id]
    );
    // Explicit no-cache headers so phones with aggressive PWA caching
    // never serve stale seller data.
    res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
    res.set('Pragma', 'no-cache');
    res.json(trader);
  });

  // ── 9. TRADER LAST-LOT + BANKS (PWA helper) ────────────────────
  app.get('/api/traders/:id/last-lot', requireAuth, (req, res) => {
    const db = getDb();
    const traderId = parseInt(req.params.id, 10);
    const lot = db.get(
      `SELECT grade, litre, bags, branch
         FROM lots WHERE trader_id = ?
         ORDER BY created_at DESC LIMIT 1`,
      [traderId]
    );
    let banks = db.all(
      `SELECT id, trader_id, bank_name, acctnum, ifsc, holder_name, is_default
         FROM trader_banks WHERE trader_id = ?
         ORDER BY is_default DESC, id ASC`,
      [traderId]
    );
    // Auto-migrate: if no rows in trader_banks but the trader row has an
    // account number, copy it across so the picker has something to show.
    if (!banks.length) {
      const t = db.get('SELECT acctnum, ifsc, holder_name FROM traders WHERE id = ?', [traderId]);
      if (t && t.acctnum && String(t.acctnum).trim()) {
        db.run(
          `INSERT INTO trader_banks (trader_id, bank_name, acctnum, ifsc, holder_name, is_default)
           VALUES (?, '', ?, ?, ?, 1)`,
          [traderId, String(t.acctnum).trim(), t.ifsc || '', t.holder_name || '']
        );
        banks = db.all(
          `SELECT id, trader_id, bank_name, acctnum, ifsc, holder_name, is_default
             FROM trader_banks WHERE trader_id = ?
             ORDER BY is_default DESC, id ASC`,
          [traderId]
        );
      }
    }
    res.json({ lastLot: lot || null, banks });
  });

  // ── 10. TRADER BANK CRUD ───────────────────────────────────────
  // PWA exposes per-trader bank management. Spice-config does this through
  // a different route shape; replicate the PWA contract here.
  // Replace the trader's bank-account rows from a `banks` array on the
  // parent payload. Mirrors server.js's syncTraderBanks: clear existing
  // rows, insert what was sent, then copy the FIRST bank back into the
  // legacy traders.ifsc/acctnum/holder_name columns so callers that
  // haven't been migrated to read trader_banks still see a valid
  // primary account.
  //
  // The desktop UI sends the whole array in one shot via POST/PUT
  // /api/traders ({...trader, banks: [...]}); the mobile PWA sends
  // individual rows via /api/traders/:id/banks below. This helper is
  // the single source of truth — keep both paths consistent.
  //
  // Skips entirely when `banks` is missing or not an array, so the PWA
  // payload (which doesn't include `banks`) is unchanged.
  function syncTraderBanksFromArray(db, traderId, banks) {
    if (!Array.isArray(banks)) return;
    const arr = banks.filter(b => b && (b.acctnum || b.ifsc));
    db.run('DELETE FROM trader_banks WHERE trader_id = ?', [traderId]);
    for (const b of arr) {
      db.run(
        `INSERT INTO trader_banks (trader_id, bank_name, branch, acctnum, ifsc, holder_name)
         VALUES (?, ?, ?, ?, ?, ?)`,
        [
          traderId,
          String(b.bank_name || '').trim(),
          String(b.branch || '').trim(),
          String(b.acctnum || '').trim(),
          String(b.ifsc || '').trim().toUpperCase(),
          String(b.holder_name || '').trim(),
        ]
      );
    }
    const first = arr[0] || {};
    db.run(
      'UPDATE traders SET ifsc = ?, acctnum = ?, holder_name = ? WHERE id = ?',
      [
        String(first.ifsc || '').trim().toUpperCase(),
        String(first.acctnum || '').trim(),
        String(first.holder_name || '').trim(),
        traderId,
      ]
    );
  }

  app.post('/api/traders/:id/banks', requireAuth, (req, res) => {
    const db = getDb();
    const traderId = parseInt(req.params.id, 10);
    const { acctnum, ifsc, label, holder_name, is_default } = req.body || {};
    if (!acctnum || !String(acctnum).trim()) {
      return res.status(400).json({ error: 'Account number is required' });
    }
    const trader = db.get('SELECT id FROM traders WHERE id = ?', [traderId]);
    if (!trader) return res.status(404).json({ error: 'Trader not found' });
    if (is_default) {
      db.run('UPDATE trader_banks SET is_default = 0 WHERE trader_id = ?', [traderId]);
    }
    // Spice-config stores the user-visible bank label in `bank_name`. PWA
    // calls it `label`. Map across.
    const info = db.run(
      `INSERT INTO trader_banks (trader_id, bank_name, acctnum, ifsc, holder_name, is_default)
       VALUES (?, ?, ?, ?, ?, ?)`,
      [
        traderId,
        String(label || '').trim(),
        String(acctnum).trim(),
        String(ifsc || '').trim().toUpperCase(),
        String(holder_name || '').trim(),
        is_default ? 1 : 0,
      ]
    );
    // Sync default to traders row (legacy callers read traders.acctnum/ifsc)
    if (is_default) {
      db.run(
        'UPDATE traders SET acctnum = ?, ifsc = ?, holder_name = ? WHERE id = ?',
        [String(acctnum).trim(), String(ifsc || '').trim().toUpperCase(), String(holder_name || '').trim(), traderId]
      );
    }
    res.json({ id: info.lastInsertRowid });
  });

  app.put('/api/traders/:tid/banks/:bid', requireAuth, (req, res) => {
    const db = getDb();
    const tid = parseInt(req.params.tid, 10);
    const bid = parseInt(req.params.bid, 10);
    const { acctnum, ifsc, label, holder_name } = req.body || {};
    const bank = db.get(
      'SELECT * FROM trader_banks WHERE id = ? AND trader_id = ?', [bid, tid]
    );
    if (!bank) return res.status(404).json({ error: 'Bank not found' });
    db.run(
      `UPDATE trader_banks
       SET acctnum = COALESCE(?, acctnum),
           ifsc = COALESCE(?, ifsc),
           bank_name = COALESCE(?, bank_name),
           holder_name = COALESCE(?, holder_name)
       WHERE id = ?`,
      [
        acctnum != null ? String(acctnum).trim() : null,
        ifsc != null ? String(ifsc).trim().toUpperCase() : null,
        label != null ? String(label).trim() : null,
        holder_name != null ? String(holder_name).trim() : null,
        bid,
      ]
    );
    res.json({ success: true });
  });

  app.delete('/api/traders/:tid/banks/:bid', requireAuth, (req, res) => {
    const db = getDb();
    const tid = parseInt(req.params.tid, 10);
    const bid = parseInt(req.params.bid, 10);
    const bank = db.get(
      'SELECT * FROM trader_banks WHERE id = ? AND trader_id = ?', [bid, tid]
    );
    if (!bank) return res.status(404).json({ error: 'Bank not found' });
    db.run('DELETE FROM trader_banks WHERE id = ?', [bid]);
    // If we deleted the default, promote the next-oldest to default
    if (bank.is_default) {
      const next = db.get(
        'SELECT id FROM trader_banks WHERE trader_id = ? ORDER BY id LIMIT 1', [tid]
      );
      if (next) db.run('UPDATE trader_banks SET is_default = 1 WHERE id = ?', [next.id]);
    }
    res.json({ success: true });
  });

  app.post('/api/traders/:tid/banks/:bid/default', requireAuth, (req, res) => {
    const db = getDb();
    const tid = parseInt(req.params.tid, 10);
    const bid = parseInt(req.params.bid, 10);
    const bank = db.get(
      'SELECT * FROM trader_banks WHERE id = ? AND trader_id = ?', [bid, tid]
    );
    if (!bank) return res.status(404).json({ error: 'Bank not found' });
    db.run('UPDATE trader_banks SET is_default = 0 WHERE trader_id = ?', [tid]);
    db.run('UPDATE trader_banks SET is_default = 1 WHERE id = ?', [bid]);
    // Sync traders row to new default
    db.run(
      `UPDATE traders SET acctnum = ?, ifsc = ?, holder_name = ? WHERE id = ?`,
      [bank.acctnum || '', bank.ifsc || '', bank.holder_name || '', tid]
    );
    res.json({ success: true });
  });

  // ── 11. LOTS — CLEAR MINE ───────────────────────────────────────
  // PWA admin button: delete all of MY lots in the current auction.
  app.post('/api/lots/clear-mine', requireAuth, (req, res) => {
    const db = getDb();
    const { auction_id } = req.body || {};
    if (!auction_id) return res.status(400).json({ error: 'auction_id required' });
    const result = db.run(
      'DELETE FROM lots WHERE auction_id = ? AND user_id = ?',
      [parseInt(auction_id, 10), req.user.username]
    );
    auditLog(req, 'bulk-delete', 'lot', null, {
      auction_id: parseInt(auction_id, 10), count: result.changes || 0, scope: 'clear-mine',
    });
    res.json({ success: true, deleted: result.changes });
  });

  // ── 11b. SELLER HISTORY (existing bookings panel) ───────────────
  // The lot-entry screen shows a "📋 EXISTING BOOKINGS" box once a
  // seller is selected, listing every lot already booked for that
  // seller in the current trade across ALL branches (so a field user
  // in NEDUMKANDAM sees that the same seller already has lots in
  // PAMPUPARA). PWA had /api/reports/seller-history/:traderId — port
  // it here over spice-config's schema.
  //
  // Response shape (matches PWA so app.html parses unchanged):
  //   { lots: [{ id, lot_no, branch, grade, bags, qty, created_at }],
  //     summary: { total_qty, total_bags, lot_count } }
  app.get('/api/reports/seller-history/:traderId', requireAuth, (req, res) => {
    const db = getDb();
    const traderId  = parseInt(req.params.traderId, 10);
    const auctionId = parseInt(req.query.auction_id, 10);
    if (!traderId || !auctionId) {
      // Return an empty-but-valid shape rather than 400 — the mobile UI
      // gracefully shows "None in this trade" for empty arrays and the
      // worst we'd accomplish with 400 is the same "Failed to load"
      // message that prompted this fix.
      return res.json({ lots: [], summary: { total_qty: 0, total_bags: 0, lot_count: 0 } });
    }
    const lots = db.all(
      `SELECT l.id, l.lot_no, l.branch, l.grade,
              l.bags, l.qty, l.created_at, l.user_id
         FROM lots l
        WHERE l.trader_id = ? AND l.auction_id = ?
        ORDER BY CAST(l.lot_no AS INTEGER), l.lot_no`,
      [traderId, auctionId]
    );
    const summary = lots.reduce((acc, l) => {
      acc.total_qty  += Number(l.qty)  || 0;
      acc.total_bags += Number(l.bags) || 0;
      acc.lot_count  += 1;
      return acc;
    }, { total_qty: 0, total_bags: 0, lot_count: 0 });
    // No-cache: this panel must always be fresh — bookings can be added
    // by another field user a moment ago.
    res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
    res.set('Pragma', 'no-cache');
    res.json({ lots, summary });
  });

  // ── 12. PRINT / RECEIPT ENDPOINTS (Pass 2) ──────────────────────
  // Six routes — receipt for one lot, batch by lot-id list, all lots for
  // a single seller, all lots for all sellers in an auction. Both GET and
  // POST variants for the batch/seller calls so the mobile UI can use
  // window.open() (GET) for the print dialog while desktop callers can
  // POST a JSON array of ids if they prefer.

  // (1) Single-lot receipt — printed right after a save from the mobile UI.
  app.get('/api/lots/:id/receipt', requireAuthFlex, (req, res) => {
    const db = getDb();
    const lot = db.get(LOT_SELECT_SQL + ' WHERE l.id = ?', [parseInt(req.params.id, 10)]);
    if (!lot) return res.status(404).json({ error: 'Lot not found' });

    // If a branch was passed, enforce it — prevents printing a receipt
    // whose header would lie about which branch the lot lives in.
    const branch = req.query && req.query.branch;
    if (branch && lot.branch !== branch) {
      return res.status(404).json({ error: `Lot ${lot.lot_no} is not in ${branch}` });
    }

    const cfg = getReceiptConfig(db);
    if (branch) cfg.branch = branch;
    const r = pickReceiptRenderer(req.query.format);

    const doc = new PDFDocument({ size: r.pageSize, margin: r.compact ? 10 : 20 });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="Lot_${lot.lot_no}_Receipt.pdf"`);
    doc.pipe(res);
    r.render(doc, [lot], cfg);
    doc.end();
  });

  // Shared helper — groups arbitrary lot rows by seller, then renders
  // one receipt page per seller. Used by print-batch and print-all-sellers.
  function streamGroupedReceipts(lots, req, res, cfg, filename) {
    const r = pickReceiptRenderer(req.query.format || (req.body && req.body.format));
    const groups = {};
    for (const l of lots) {
      const key = l.trader_id || ('u_' + (l.trader_name || 'unknown'));
      (groups[key] || (groups[key] = [])).push(l);
    }
    const doc = new PDFDocument({ size: r.pageSize, margin: r.compact ? 10 : 20 });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="${filename}"`);
    doc.pipe(res);
    Object.values(groups).forEach((group, idx) => {
      if (idx > 0) doc.addPage();
      r.render(doc, group, cfg);
    });
    doc.end();
  }

  // (2) Batch by explicit lot-id list. Mobile passes ?ids=1,2,3 via GET
  // (window.open); desktop posts { ids: [...] } as JSON.
  function handlePrintBatch(ids, req, res) {
    const db = getDb();
    if (!ids || !ids.length) return res.status(400).json({ error: 'No lot IDs provided' });
    const branch = (req.query && req.query.branch) || (req.body && req.body.branch) || '';
    let lots = ids
      .map(id => db.get(LOT_SELECT_SQL + ' WHERE l.id = ?', [parseInt(id, 10)]))
      .filter(Boolean);
    if (branch) lots = lots.filter(l => l.branch === branch);
    if (!lots.length) {
      return res.status(404).json({ error: branch ? `No lots in ${branch}` : 'No lots found' });
    }
    const cfg = getReceiptConfig(db);
    if (branch) cfg.branch = branch;
    streamGroupedReceipts(lots, req, res, cfg, `Lots_Receipt_${lots.length}.pdf`);
  }
  app.post('/api/lots/print-batch', requireAuthFlex, (req, res) =>
    handlePrintBatch(req.body && req.body.ids, req, res));
  app.get('/api/lots/print-batch', requireAuthFlex, (req, res) => {
    const ids = String(req.query.ids || '').split(',').map(Number).filter(n => n > 0);
    handlePrintBatch(ids, req, res);
  });

  // (3) All lots for one seller in one auction — "📋 All by Seller"
  function handlePrintSeller(traderId, auctionId, req, res) {
    const db = getDb();
    if (!traderId || !auctionId) {
      return res.status(400).json({ error: 'trader_id and auction_id required' });
    }
    const branch = (req.query && req.query.branch) || (req.body && req.body.branch) || '';
    const params = [parseInt(auctionId, 10), parseInt(traderId, 10)];
    let where = 'l.auction_id = ? AND l.trader_id = ?';
    if (branch) { where += ' AND l.branch = ?'; params.push(branch); }
    const lots = db.all(LOT_SELECT_SQL + ' WHERE ' + where + ' ORDER BY CAST(l.lot_no AS INTEGER), l.lot_no', params);
    if (!lots.length) {
      return res.status(404).json({
        error: branch ? `No lots for this seller in ${branch}` : 'No lots found',
      });
    }
    const cfg = getReceiptConfig(db);
    if (branch) cfg.branch = branch;
    const fmt = (req.query && req.query.format) || (req.body && req.body.format);
    const r = pickReceiptRenderer(fmt);
    // Auto-grow page for long seller histories
    const pageSize = r.compact
      ? [180, Math.min(160 + lots.length * 12 + 60, 700)]
      : [340, Math.min(200 + lots.length * 18 + 80, 800)];
    const doc = new PDFDocument({ size: pageSize, margin: r.compact ? 10 : 20 });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition',
      `inline; filename="Seller_${(lots[0].trader_name || 'Receipt').replace(/[^A-Za-z0-9]+/g,'_')}.pdf"`);
    doc.pipe(res);
    r.render(doc, lots, cfg);
    doc.end();
  }
  app.post('/api/lots/print-seller', requireAuthFlex, (req, res) =>
    handlePrintSeller(req.body && req.body.trader_id, req.body && req.body.auction_id, req, res));
  app.get('/api/lots/print-seller', requireAuthFlex, (req, res) =>
    handlePrintSeller(req.query.trader_id, req.query.auction_id, req, res));

  // (4) Every seller in an auction (optionally branch-scoped) — admin's
  // end-of-day bulk print.
  app.get('/api/lots/print-all-sellers/:auctionId', requireAuthFlex, (req, res) => {
    const db = getDb();
    const auctionId = parseInt(req.params.auctionId, 10);
    const branch = req.query.branch || '';
    const params = [auctionId];
    let where = 'l.auction_id = ?';
    if (branch) { where += ' AND l.branch = ?'; params.push(branch); }
    // Order by seller name first so each PDF page covers one seller in lot order.
    const lots = db.all(
      LOT_SELECT_SQL + ' WHERE ' + where +
      ' ORDER BY COALESCE(t.name, l.name), CAST(l.lot_no AS INTEGER), l.lot_no',
      params
    );
    if (!lots.length) return res.status(404).json({ error: 'No lots found' });
    const cfg = getReceiptConfig(db);
    if (branch) cfg.branch = branch;
    streamGroupedReceipts(lots, req, res, cfg, 'All_Sellers_Receipt.pdf');
  });

  console.log('[mobile-bridge] mounted /mobile + PWA-compat routes (Pass 1 + Pass 2 receipts)');
}

module.exports = { mountMobile };
