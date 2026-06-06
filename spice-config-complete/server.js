const express = require('express');
const crypto = require('crypto');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { initDb, getDb, flushDb, replaceDbFromBuffer, DB_PATH } = require('./db');
const { initCompanySettings, CATEGORIES, getAllSettings, updateSettings, getSettingsFlat, getGSTRates, getAllPresets, setActivePresetCode, savePreset, getActivePresetCode, getPreset } = require('./company-config');
const { calculateLot, buildSalesInvoice, buildPurchaseInvoice, buildAgriBill, buildDebitNote, listAgriSellers, getPaymentSummary, getBankPaymentData, getTDSReturnData, getSalesJournal, getPurchaseJournal } = require('./calculations');
const { generatePurchaseInvoicePDF, generateCropReceiptPDF, generateAgriBillPDF, generateSalesInvoicePDF, generateSalesInvoicesBatchPDF, generatePurchaseInvoicesBatchPDF, generateAgriBillsBatchPDF } = require('./invoice-pdf');
const { generateDebitNoteBatchPDF } = require('./debit-note-print');
const { EXPORT_TYPES } = require('./exports');
const { exportPdf: exportAnyPdf } = require('./exports-pdf');
const { DBF_EXPORTS } = require('./dbf-exports');
const { REPORTS: LORRY_REPORTS } = require('./lorry-reports');
const { REPORTS: SPICE_BOARD_REPORTS, getReportFilters: getSpiceBoardFilters } = require('./spice-board-reports');
const { getTradeSummary, getBranchComparison, generateTradeSummaryPDF } = require('./reports');
// Per-install time-bombed licensing — see license.js for the model.
// Token signing/verification + the license_state row helpers live there;
// db.js owns the schema migration.
const license = require('./license');
const {
  generSalesXML, generSalesIspXML, generSalesAspXML, generIspPurchaseXML,
  generRDPurchaseXML, generURDPurchaseXML, generDebitNoteXML, generLedgerXML,
  buildSalesRows, buildSalesIspRows, buildSalesAspRows,
  buildRDPurchaseRows, buildURDPurchaseRows, buildDebitNoteRows, buildLedgerRows,
  buildSalesPartyLedgerRows, buildRDPartyLedgerRows, buildURDPartyLedgerRows,
  listAuctionParties,
} = require('./tally-xml');

const app = express();
app.use(express.json({ limit: '50mb' }));

// Disable caching of HTML files so users always get the latest UI without
// needing a hard-reload. This is critical for ngrok-tunnelled deployments
// where intermediate proxies may cache aggressively. JavaScript/CSS/
// images can still be cached normally (handled by the static middleware
// after this).
app.use((req, res, next) => {
  if (req.path === '/' || req.path.endsWith('.html')) {
    res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
    res.set('Pragma', 'no-cache');
    res.set('Expires', '0');
  }
  next();
});

app.use(express.static(path.join(__dirname, 'public')));

// Prevent browser/proxy caching of API responses so Refresh buttons actually
// fetch fresh data (without this, fetch() may return stale cached JSON)
app.use('/api', (req, res, next) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  next();
});

// Health check — used by the Electron wrapper to wait until the server
// is ready to accept requests before loading the window URL. Returns a
// minimal 200 with no auth required.
app.get('/api/health', (req, res) => {
  res.json({ ok: true, ts: Date.now() });
});

// Public brand probe — used by the pre-auth login page so the welcome
// card can show the actual company short name + logo image instead of
// a hardcoded literal. Returns only the active preset's short name,
// the static logo URL, and a fallback text mark. No other settings
// leak (auth-free endpoint).
app.get('/api/brand', (req, res) => {
  try {
    const db = getDb();
    const active = (typeof getActivePresetCode === 'function')
      ? getActivePresetCode(db)
      : 'ISP';
    const preset = (typeof getPreset === 'function') ? getPreset(db, active) : {};
    const flat = getSettingsFlat(db);
    // Short-name preference: preset's own short_name → flat fallback →
    // sister-company fallback for ASP → final literal default.
    const pickName = () => {
      if (preset.short_name)  return preset.short_name;
      if (preset.trade_name)  return preset.trade_name;
      if (active === 'ASP') {
        return flat.s_short_name || 'Amazing Spice Park';
      }
      return flat.short_name || flat.trade_name || 'Spice Config';
    };
    const name = String(pickName()).trim();
    // Logo: ONLY the active preset's file. We deliberately do not
    // fall back to the other preset's logo — showing the ISPL bee on
    // an ASP-active session (or vice versa) is misbranding. When the
    // active preset's file is missing the client falls through to the
    // text-mark tile / .dot background.
    const which = active === 'ASP' ? 'asp' : 'ispl';
    const target = LOGO_FILES && LOGO_FILES[which];
    const hasLogo = !!(target && fs.existsSync(target));
    const logoUrl = hasLogo ? ('/logo-' + which + '.png?v=' + Math.floor(fs.statSync(target).mtimeMs)) : null;
    // Two-letter tile (used only when logoUrl is null).
    const mark = active === 'ASP' ? 'AS' : 'IS';
    res.json({ active, name, code: active, mark, logoUrl, hasLogo });
  } catch (e) {
    res.json({ active: 'ISP', name: '', code: 'ISP', mark: 'IS', logoUrl: null, hasLogo: false });
  }
});

// File upload setup
// Honor SPICE_DATA_DIR so uploads also land in userData when packaged.
const uploadDir = path.join(process.env.SPICE_DATA_DIR || path.join(__dirname, 'data'), 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
const upload = multer({ dest: uploadDir, limits: { fileSize: 20 * 1024 * 1024 } });

const hash = pw => crypto.createHash('sha256').update(pw).digest('hex');

// ── Password helpers ─────────────────────────────────────────
// Bcrypt is preferred for new/changed passwords; SHA-256 rows are
// honored for legacy users so existing installs keep working. On a
// successful login against a legacy row, the caller opportunistically
// re-hashes to bcrypt (see /api/auth/login in mobile-bridge.js).
// We use bcryptjs (pure JS) so Electron-packaged builds don't need an
// extra native-module compile/asar-unpack step.
const bcrypt = require('bcryptjs');
function isLegacyHash(h) {
  // Bcrypt hashes start with $2a$ / $2b$ / $2y$. Anything else (here:
  // 64-char SHA-256 hex) is treated as legacy and eligible for upgrade.
  return !h || !/^\$2[aby]\$/.test(h);
}
async function verifyPassword(plain, stored) {
  if (!stored) return false;
  if (isLegacyHash(stored)) return hash(plain) === stored;
  return bcrypt.compare(plain, stored);
}
async function hashPassword(plain) {
  return bcrypt.hash(plain, 10);
}

// ── Date helpers ──
// Convert any date-ish input (Date object, Excel serial number, dd/mm/yyyy,
// yyyy-mm-dd, etc.) to canonical ISO yyyy-mm-dd for storage.
function normalizeDate(v) {
  if (v === null || v === undefined || v === '') return '';
  // Date object
  if (v instanceof Date && !isNaN(v)) {
    return v.toISOString().slice(0, 10);
  }
  // Number = Excel serial (days since 1900-01-01, with the famous 1900 leap-year bug)
  if (typeof v === 'number' && v > 0 && v < 80000) {
    // Excel epoch: 1899-12-30 (accounts for the 1900 leap-year bug)
    const ms = (v - 25569) * 86400 * 1000;
    const d = new Date(ms);
    if (!isNaN(d)) return d.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  // ISO yyyy-mm-dd (or yyyy-mm-dd HH:MM:SS)
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  // dd/mm/yyyy or dd-mm-yyyy
  const m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m1) return `${m1[3]}-${m1[2].padStart(2, '0')}-${m1[1].padStart(2, '0')}`;
  // Pure numeric string Excel serial
  if (/^\d+(\.\d+)?$/.test(s)) {
    const n = Number(s);
    if (n > 0 && n < 80000) {
      const ms = (n - 25569) * 86400 * 1000;
      const d = new Date(ms);
      if (!isNaN(d)) return d.toISOString().slice(0, 10);
    }
  }
  // Last resort: try Date parsing
  const d = new Date(s);
  if (!isNaN(d)) return d.toISOString().slice(0, 10);
  return s;
}

// Display: yyyy-mm-dd → dd/mm/yyyy (handles Excel serials defensively too)
function fmtDate(d) {
  if (!d && d !== 0) return '';
  const iso = normalizeDate(d);
  if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) {
    const [y, m, day] = iso.split('-');
    return `${day}/${m}/${y}`;
  }
  return String(d);
}

function withFmtDate(rows, field = 'date') {
  if (!Array.isArray(rows)) return rows;
  return rows.map(r => ({ ...r, date_fmt: fmtDate(r[field]) }));
}

// Auth middleware: verify a valid session, attach req.user/req.session.
// DOES NOT check role — use this for endpoints that any logged-in user
// (admin OR regular user) should be able to hit (GET list endpoints mostly).
function requireAuth(req, res, next) {
  const token = (req.headers.authorization || '').replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'No token' });
  const db = getDb();
  const session = db.get('SELECT * FROM sessions WHERE token = ?', [token]);
  // 401 (not 403) so the client's auto-logout flow triggers silently
  // instead of surfacing a "Session expired" toast to the user.
  if (!session) return res.status(401).json({ error: 'Session expired — please sign in again' });
  const user = db.get('SELECT * FROM users WHERE id = ?', [session.user_id]);
  if (!user) return res.status(401).json({ error: 'Unauthorized' });
  // Touch last_used_at for cleanup / activity display
  db.run(`UPDATE sessions SET last_used_at = datetime('now','localtime') WHERE token = ?`, [token]);
  req.user = user;
  req.session = session;
  next();
}

// Admin-only middleware: gates mutations, settings, deletes, user management.
// Runs requireAuth first, then verifies role.
function requireAdmin(req, res, next) {
  requireAuth(req, res, (err) => {
    if (err) return next(err);
    if (!req.user || req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Admin access required for this action' });
    }
    next();
  });
}

// ══════════════════════════════════════════════════════════════
// ROLE-BASED PERMISSIONS
// ══════════════════════════════════════════════════════════════
// Four pre-defined role tiers, each granting a fixed set of capabilities.
// Capabilities are referenced by name everywhere in the code that needs
// permission gating, so adding/removing a capability touches one place.
//
// Hierarchy (least → most privileged):
//   viewer    — read-only
//   operator  — daily auction-floor work (lots, invoices, buyers/traders)
//   manager   — branch oversight (auctions, settings, revert)
//   admin     — full control (delete, user management, business state)
//
// Capability names are short snake_case strings. New capabilities are
// added by including the name in the appropriate role(s) below.
const ROLE_PERMISSIONS = {
  viewer: new Set([
    'view',           // read any list / detail
    'export',         // download XLSX / PDF / CSV exports
    'self_password',  // change own password
    'lot_entry_view'  // also see the Lot Entry tab + its data (read-only)
  ]),
  // Field-staff role for the auction-hall lot entry workflow. Sees only
  // the Lot Entry tab in the sidebar (everything else hidden via the
  // role-scoped CSS in index.html). The narrow scope — create trades,
  // search sellers, create/edit own lots — matches what the auction-floor
  // workflow actually needs.
  lot_entry: new Set([
    'self_password',
    'view',            // read shared trade/lot data so multi-user sessions
                       // can all see the same in-progress entries
    'lot_entry_view',  // Lot Entry tab + its endpoints
    'lot_write',       // create/edit own lots
    'auction_write'    // create new trades on-the-fly during an auction day
  ]),
  operator: new Set([
    'view', 'export', 'self_password',
    'lot_entry_view',// operators can also use the Lot Entry tab
    'lot_write',     // create/edit lots, calculate, validate, price-import
    'invoice_write', // generate sales/purchase/bills + edit
    'trader_write',  // create/edit/delete-bank traders
    'buyer_write'    // create/edit buyers (per user decision: tax fields editable)
  ]),
  manager: new Set([
    'view', 'export', 'self_password', 'lot_entry_view',
    'lot_write', 'invoice_write', 'trader_write', 'buyer_write',
    'auction_write',  // create/edit auctions (trades)
    'invoice_revert', // revert sales/purchase/bills (undo invoice)
    'settings_write', // edit company settings (rates, addresses, flags)
    'state_toggle'    // toggle business state TN ↔ KL
  ]),
  admin: new Set([
    'view', 'export', 'self_password', 'lot_entry_view',
    'lot_write', 'invoice_write', 'trader_write', 'buyer_write',
    'auction_write', 'invoice_revert', 'settings_write', 'state_toggle',
    'delete',       // delete any individual record
    'delete_all',   // bulk Delete All (sales, purchases, lots, etc.)
    'user_manage'   // create/delete users, reset passwords, revoke sessions
  ])
};

// Best-effort capability lookup. Unknown roles get treated as 'viewer'
// (safest default — fails closed instead of open).
function userHas(role, capability) {
  const perms = ROLE_PERMISSIONS[role] || ROLE_PERMISSIONS.viewer;
  return perms.has(capability);
}

// Middleware factory: returns an Express middleware that requires the
// authenticated user to have a specific capability.
//
// Usage:
//   app.post('/api/lots', requirePermission('lot_write'), handler)
//   app.delete('/api/invoices/:id', requirePermission('delete'), handler)
//
// Falls through to next() on success; sends 403 with a clear message
// indicating both the user's current role AND the capability required
// (helps the client show a useful error rather than a generic "denied").
function requirePermission(capability) {
  return (req, res, next) => {
    requireAuth(req, res, (err) => {
      if (err) return next(err);
      if (!req.user) return res.status(401).json({ error: 'Authentication required' });
      if (!userHas(req.user.role, capability)) {
        return res.status(403).json({
          error: `Your role (${req.user.role}) does not allow this action`,
          required: capability,
          role: req.user.role
        });
      }
      next();
    });
  };
}

// Same idea but accepts ANY of a list of capabilities. Used for endpoints
// that serve multiple roles — e.g., the trader search endpoint, which
// general operators reach through 'view' and lot-entry users reach
// through their own lot_entry_view capability.
function requireAnyPermission(...capabilities) {
  return (req, res, next) => {
    requireAuth(req, res, (err) => {
      if (err) return next(err);
      if (!req.user) return res.status(401).json({ error: 'Authentication required' });
      const hasAny = capabilities.some(c => userHas(req.user.role, c));
      if (!hasAny) {
        return res.status(403).json({
          error: `Your role (${req.user.role}) does not allow this action`,
          required: capabilities.join(' or '),
          role: req.user.role
        });
      }
      next();
    });
  };
}

// Convenience aliases — readable names for the most common gate points.
// Encapsulates the permission name so callers don't repeat string literals.
const requireView          = requirePermission('view');
const requireViewOrLotEntry = requireAnyPermission('view', 'lot_entry_view');
const requireLotWrite      = requirePermission('lot_write');
const requireInvoiceWrite  = requirePermission('invoice_write');
const requireInvoiceRevert = requirePermission('invoice_revert');
const requireTraderWrite   = requirePermission('trader_write');
const requireBuyerWrite    = requirePermission('buyer_write');
const requireAuctionWrite  = requirePermission('auction_write');
const requireSettingsWrite = requirePermission('settings_write');
const requireStateToggle   = requirePermission('state_toggle');
const requireDelete        = requirePermission('delete');
const requireDeleteAll     = requirePermission('delete_all');
const requireUserManage    = requirePermission('user_manage');
const requireExport        = requirePermission('export');

// ══════════════════════════════════════════════════════════════
// AUTH
// ══════════════════════════════════════════════════════════════
app.post('/api/login', async (req, res) => {
  // Hard licence gate before the credential check — when the install
  // is expired we don't even pretend to authenticate. The frontend
  // login script catches the 451 and redirects to /renew.html with
  // the install ID pre-filled.
  try {
    const lstatus = license.getStatus(getDb());
    if (lstatus.expired) {
      return res.status(451).json({
        error: 'license_expired',
        message: `Your access window ended on ${lstatus.expires_at}. Send the install ID below to your provider to receive a renewal token.`,
        install_id: lstatus.install_id,
        expires_at: lstatus.expires_at,
      });
    }
  } catch (_) { /* license check failed → let login through; fail-open is friendlier than a brick */ }
  const { username, password, device_label } = req.body;
  if (!username || !password) return res.status(400).json({ error: 'Missing credentials' });
  const db = getDb();
  const user = db.get('SELECT * FROM users WHERE username = ?', [username]);
  const ok = user ? await verifyPassword(password, user.password_hash) : false;
  if (!user || !ok) return res.status(401).json({ error: 'Invalid credentials' });
  // Opportunistic upgrade: legacy SHA-256 rows get re-hashed to bcrypt on
  // first successful desktop login so mobile + desktop stay in sync.
  if (isLegacyHash(user.password_hash)) {
    try {
      const upgraded = await hashPassword(password);
      db.run('UPDATE users SET password_hash = ? WHERE id = ?', [upgraded, user.id]);
    } catch (_) { /* non-fatal */ }
  }
  const token = crypto.randomBytes(32).toString('hex');
  // Create a new session row WITHOUT deleting any existing sessions —
  // this lets the same user stay logged in on multiple devices simultaneously.
  db.run('INSERT INTO sessions (token, user_id, device_label) VALUES (?, ?, ?)', [token, user.id, device_label || '']);
  // Clean up very old sessions (> 30 days) so the table doesn't grow forever
  db.run(`DELETE FROM sessions WHERE last_used_at < datetime('now','-30 days')`);
  // Return the user's capabilities array so the client can hide buttons
  // they're not allowed to use. Server still validates every request.
  const permissions = Array.from(ROLE_PERMISSIONS[user.role] || ROLE_PERMISSIONS.viewer);
  res.json({ token, role: user.role, username: user.username, permissions });
});
app.post('/api/logout', (req, res) => {
  const t = (req.headers.authorization||'').replace('Bearer ','');
  if (t) getDb().run('DELETE FROM sessions WHERE token = ?', [t]);
  res.json({ success: true });
});
app.get('/api/me', requireView, (req, res) => {
  const permissions = Array.from(ROLE_PERMISSIONS[req.user.role] || ROLE_PERMISSIONS.viewer);
  res.json({ username: req.user.username, role: req.user.role, permissions });
});

// ══════════════════════════════════════════════════════════════
// LICENSING — install ID, expiry probe, token apply
// ══════════════════════════════════════════════════════════════
//
// Both endpoints are deliberately auth-free: the renewal page lives
// outside the normal authenticated flow (the operator is locked out
// of /api/login while expired, so they need a way to apply a new
// token from a logged-out state). status is also used by the topbar
// countdown pill on the main UI.
app.get('/api/license/status', (req, res) => {
  res.set('Cache-Control', 'no-store');
  try {
    const status = license.getStatus(getDb());
    res.json(status);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.post('/api/license/apply', (req, res) => {
  const token = (req.body && (req.body.token || req.body.license_token) || '').trim();
  if (!token) return res.status(400).json({ error: 'token required' });
  const result = license.applyToken(getDb(), token);
  if (!result.ok) return res.status(400).json(result);
  res.json(result);
});

// ─────────────────────────────────────────────────────────────
// Admin endpoint — back-channel for ops without shell access.
//
// Authenticated by the same LICENSE_SECRET used to sign tokens
// (only the developer knows it). When LICENSE_SECRET is unset on
// the server, this endpoint refuses ALL requests so the fallback
// development secret in license.js can never double as a remote
// admin backdoor on a misconfigured production deploy.
//
// Usage from anywhere (no shell required):
//
//   curl -X POST https://<app>/api/license/admin/set-expiry \
//     -H "content-type: application/json" \
//     -H "X-License-Secret: $LICENSE_SECRET" \
//     -d '{"expires_at":"2020-01-01"}'      # force expired (test)
//
//   curl -X POST https://<app>/api/license/admin/set-expiry \
//     -H "content-type: application/json" \
//     -H "X-License-Secret: $LICENSE_SECRET" \
//     -d '{"expires_at":"2026-07-02"}'      # restore
//
// Sets the row's active_token to NULL so the audit trail doesn't
// falsely attribute the new expiry to a previously-applied token.
// For token-driven history use POST /api/license/apply.
// ─────────────────────────────────────────────────────────────
function _requireLicenseAdmin(req) {
  const envSecret = String(process.env.LICENSE_SECRET || '').trim();
  if (!envSecret) {
    return { status: 403, error: 'admin disabled: LICENSE_SECRET env var is not set on this server' };
  }
  const provided = String(req.headers['x-license-secret'] || '').trim();
  if (!provided) {
    return { status: 403, error: 'X-License-Secret header required' };
  }
  // Constant-time compare on equal-length buffers. Bail before the
  // compare when lengths differ so we don't leak the secret length
  // (and so timingSafeEqual doesn't throw on a length mismatch).
  const a = Buffer.from(envSecret, 'utf8');
  const b = Buffer.from(provided, 'utf8');
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) {
    return { status: 403, error: 'invalid X-License-Secret' };
  }
  return null;
}

app.post('/api/license/admin/set-expiry', (req, res) => {
  res.set('Cache-Control', 'no-store');
  const denied = _requireLicenseAdmin(req);
  if (denied) return res.status(denied.status).json({ error: denied.error });

  const body = req.body || {};
  const expires_at = String(body.expires_at || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(expires_at)) {
    return res.status(400).json({ error: 'expires_at must be YYYY-MM-DD (e.g. 2026-07-31)' });
  }
  // Sanity bound — anything outside 2020-2099 is almost certainly a typo.
  // Tokens already issued past 2099 would also be unusable.
  const yr = Number(expires_at.slice(0, 4));
  if (yr < 2020 || yr > 2099) {
    return res.status(400).json({ error: 'expires_at year out of range (2020-2099)' });
  }

  try {
    const db = getDb();
    // Make sure the row exists — on a fresh install this endpoint may
    // be hit before any normal traffic has triggered the bootstrap.
    license.ensureState(db);
    db.run(
      'UPDATE license_state SET expires_at = ?, active_token = NULL WHERE id = 1',
      [expires_at]
    );
    res.json({ ok: true, status: license.getStatus(db) });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// USER MANAGEMENT (admin-only)
// ══════════════════════════════════════════════════════════════
app.get('/api/users', requireUserManage, (req, res) => {
  const db = getDb();
  const users = db.all(`
    SELECT u.id, u.username, u.role, u.created_at,
      (SELECT COUNT(*) FROM sessions s WHERE s.user_id = u.id) as active_sessions,
      (SELECT MAX(last_used_at) FROM sessions s WHERE s.user_id = u.id) as last_active
    FROM users u ORDER BY u.id ASC
  `);
  res.json(users);
});

app.post('/api/users', requireUserManage, async (req, res) => {
  const { username, password, role } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: 'username and password are required' });
  if (password.length < 4) return res.status(400).json({ error: 'Password must be at least 4 characters' });
  if (!/^[a-zA-Z0-9_.-]{3,30}$/.test(username)) return res.status(400).json({ error: 'Username: 3–30 chars, letters/digits/._- only' });
  // Validate role against the known set. Default to 'operator' (the most
  // common and least privileged write-capable role) if missing or invalid.
  // Legacy 'user' role is mapped to 'viewer' for backward compat.
  const VALID_ROLES = ['viewer', 'lot_entry', 'operator', 'manager', 'admin'];
  let finalRole = (role || '').toLowerCase();
  if (finalRole === 'user') finalRole = 'viewer';
  if (!VALID_ROLES.includes(finalRole)) finalRole = 'operator';
  const db = getDb();
  const existing = db.get('SELECT id FROM users WHERE username = ?', [username]);
  if (existing) return res.status(400).json({ error: 'Username already exists' });
  const newHash = await hashPassword(password);
  db.run(
    'INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)',
    [username, newHash, finalRole]
  );
  const created = db.get('SELECT id, username, role FROM users WHERE username = ?', [username]);
  res.json({ success: true, id: created ? created.id : null, username, role: finalRole });
});

// Update an existing user's role (for promoting/demoting users without
// recreating them). Admin-only — same gate as creating users.
app.put('/api/users/:id/role', requireUserManage, (req, res) => {
  const { role } = req.body || {};
  const VALID_ROLES = ['viewer', 'lot_entry', 'operator', 'manager', 'admin'];
  let finalRole = String(role || '').toLowerCase();
  if (finalRole === 'user') finalRole = 'viewer';
  if (!VALID_ROLES.includes(finalRole)) {
    return res.status(400).json({ error: `Role must be one of: ${VALID_ROLES.join(', ')}` });
  }
  const db = getDb();
  const target = db.get('SELECT id, username, role FROM users WHERE id = ?', [req.params.id]);
  if (!target) return res.status(404).json({ error: 'User not found' });
  // Safety: don't let admins demote the last remaining admin (would lock
  // everyone out of user management).
  if (target.role === 'admin' && finalRole !== 'admin') {
    const adminCount = db.get(`SELECT COUNT(*) as n FROM users WHERE role = 'admin'`).n;
    if (adminCount <= 1) {
      return res.status(400).json({ error: 'Cannot demote the last admin — promote someone else first' });
    }
  }
  db.run('UPDATE users SET role = ? WHERE id = ?', [finalRole, target.id]);
  res.json({ success: true, username: target.username, role: finalRole });
});

app.put('/api/users/:id/password', requireUserManage, async (req, res) => {
  const { password } = req.body || {};
  if (!password || password.length < 4) return res.status(400).json({ error: 'Password must be at least 4 characters' });
  const db = getDb();
  const user = db.get('SELECT id, username FROM users WHERE id = ?', [req.params.id]);
  if (!user) return res.status(404).json({ error: 'User not found' });
  const newHash = await hashPassword(password);
  db.run('UPDATE users SET password_hash = ? WHERE id = ?', [newHash, user.id]);
  // Invalidate all sessions of that user (force re-login after password change)
  db.run('DELETE FROM sessions WHERE user_id = ?', [user.id]);
  res.json({ success: true, username: user.username });
});

app.delete('/api/users/:id', requireUserManage, (req, res) => {
  const db = getDb();
  const target = db.get('SELECT id, username, role FROM users WHERE id = ?', [req.params.id]);
  if (!target) return res.status(404).json({ error: 'User not found' });
  // Safety: don't let admin delete themselves
  if (target.id === req.user.id) return res.status(400).json({ error: 'You cannot delete your own account while signed in' });
  // Safety: never allow deleting the last remaining user
  const total = db.get('SELECT COUNT(*) as c FROM users').c;
  if (total <= 1) return res.status(400).json({ error: 'Cannot delete the last remaining user' });
  db.run('DELETE FROM sessions WHERE user_id = ?', [target.id]);
  db.run('DELETE FROM users WHERE id = ?', [target.id]);
  res.json({ success: true, username: target.username });
});

// Change own password — shortcut that doesn't require user id
app.put('/api/me/password', requirePermission('self_password'), async (req, res) => {
  const { current_password, new_password } = req.body || {};
  if (!current_password || !new_password) return res.status(400).json({ error: 'Both current and new password required' });
  if (new_password.length < 4) return res.status(400).json({ error: 'New password must be at least 4 characters' });
  const db = getDb();
  const user = db.get('SELECT * FROM users WHERE id = ?', [req.user.id]);
  const ok = user ? await verifyPassword(current_password, user.password_hash) : false;
  if (!user || !ok) return res.status(401).json({ error: 'Current password is incorrect' });
  const newHash = await hashPassword(new_password);
  db.run('UPDATE users SET password_hash = ? WHERE id = ?', [newHash, user.id]);
  // Kill all OTHER sessions (keep current one)
  db.run('DELETE FROM sessions WHERE user_id = ? AND token != ?', [user.id, req.session.token]);
  res.json({ success: true });
});

// See my active sessions (all devices signed in as me)
app.get('/api/me/sessions', requireView, (req, res) => {
  const db = getDb();
  const sessions = db.all(
    `SELECT token, device_label, created_at, last_used_at,
            CASE WHEN token = ? THEN 1 ELSE 0 END as is_current
     FROM sessions WHERE user_id = ? ORDER BY last_used_at DESC`,
    [req.session.token, req.user.id]
  );
  // Mask tokens — only show last 8 chars
  res.json(sessions.map(s => ({ ...s, token: '…' + (s.token || '').slice(-8) })));
});

// Revoke (log out) another session I own
app.delete('/api/me/sessions/:tokenSuffix', requireView, (req, res) => {
  const suffix = req.params.tokenSuffix;
  const db = getDb();
  // Find session by matching suffix, for THIS user only
  const sessions = db.all('SELECT token FROM sessions WHERE user_id = ?', [req.user.id]);
  const match = sessions.find(s => (s.token || '').endsWith(suffix));
  if (!match) return res.status(404).json({ error: 'Session not found' });
  if (match.token === req.session.token) return res.status(400).json({ error: 'Use Logout to end your current session' });
  db.run('DELETE FROM sessions WHERE token = ?', [match.token]);
  res.json({ success: true });
});

// ══════════════════════════════════════════════════════════════
// MOBILE PWA BRIDGE
// ══════════════════════════════════════════════════════════════
// Mount BEFORE company-settings + the rest so the bridge wins route
// matching for /api/traders, /api/traders/:id, etc. — those are the
// unified seller-write endpoints that both apps share. The desktop's
// own /api/traders POST further down still loads but is shadowed.
// Mounting here also serves /mobile (the PWA) and the /api/auth/*,
// /api/config, /api/lots (query form), /api/logo, and receipt-print
// routes the PWA's app.html depends on.
const { mountMobile } = require('./mobile-bridge');
mountMobile(app, { getDb, requireAuth, verifyPassword, hashPassword, isLegacyHash, ROLE_PERMISSIONS });

// ══════════════════════════════════════════════════════════════
// COMPANY SETTINGS
// ══════════════════════════════════════════════════════════════
app.get('/api/company-settings', requireView, (req, res) => {
  res.json({ categories: CATEGORIES, settings: getAllSettings(getDb()) });
});
app.put('/api/company-settings', requireSettingsWrite, (req, res) => {
  const count = updateSettings(getDb(), req.body.settings || {});
  res.json({ success: true, updated: count });
});
app.get('/api/company-settings/flat', requireView, (req, res) => res.json(getSettingsFlat(getDb())));

// ── Company identity presets (ISP / ASP) ─────────────────────────────
// Two named snapshots of the 8 fields in category='company'. The active
// preset's values overlay onto company_settings so invoice PDFs and
// exports continue reading from the flat settings object. Dad edits each
// preset independently; the Logo Code dropdown in Settings flips which
// one is active.

// Returns: {ISP: {logo:..., trade_name:..., ...}, ASP: {...}, active: 'ISP'}
app.get('/api/company-presets', requireView, (req, res) => {
  try { res.json(getAllPresets(getDb())); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

// Switch which preset is active. Overlay values are pushed into
// company_settings immediately so downstream reads reflect the switch
// without a server restart.
// Body: {code: 'ASP'}
// NOTE: This route MUST be declared before PUT /:code because Express
// routes match in declaration order. If /:code comes first, 'active'
// gets captured as the code parameter and this route is unreachable.
app.put('/api/company-presets/active', requireStateToggle, (req, res) => {
  try {
    const code = req.body.code;
    if (code !== 'ISP' && code !== 'ASP') {
      return res.status(400).json({ error: 'Invalid preset code — must be ISP or ASP' });
    }
    setActivePresetCode(getDb(), code);
    res.json({ success: true, active: code });
  } catch (e) {
    console.error('[presets/active] Failed to switch preset:', e);
    res.status(500).json({ error: e.message });
  }
});

// Save values to a specific preset (does not change the active preset).
// Body: {values: {logo: 'ASP', trade_name: 'AMAZING SPICE PARK', ...}}
app.put('/api/company-presets/:code', requireSettingsWrite, (req, res) => {
  try {
    const code = req.params.code;
    if (code !== 'ISP' && code !== 'ASP') {
      return res.status(400).json({ error: 'Invalid preset code — must be ISP or ASP' });
    }
    savePreset(getDb(), code, req.body.values || {});
    res.json({ success: true });
  } catch (e) {
    console.error('[presets/save] Failed to save preset:', e);
    res.status(500).json({ error: e.message });
  }
});

// ── Logo upload/delete ────────────────────────────────────────
// Supports two logos: 'ispl' (main company) and 'asp' (sister company).
// Files are saved to public/ so the PDF generator can read them at runtime.
// Also served to the browser for preview via GET /logo-<which>.png.
const LOGO_FILES = {
  ispl: path.join(__dirname, 'public', 'logo-ispl.png'),
  asp:  path.join(__dirname, 'public', 'logo-asp.png'),
};
app.post('/api/company-settings/logo/:which', requireSettingsWrite, upload.single('file'), (req, res) => {
  const which = req.params.which;
  if (!LOGO_FILES[which]) return res.status(400).json({ error: 'Invalid logo type (use ispl or asp)' });
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  // Only allow image types
  const ext = (req.file.originalname.split('.').pop() || '').toLowerCase();
  if (!['png', 'jpg', 'jpeg'].includes(ext)) {
    fs.unlinkSync(req.file.path);
    return res.status(400).json({ error: 'Only PNG or JPG images allowed' });
  }
  // Always save as .png at the fixed path (PDFKit handles both PNG and JPEG from PNG extension? No — rename to real ext)
  // Simpler: keep .png in the PDF code always pointing to PNG. For JPEG uploads, save as .jpg alongside.
  const target = LOGO_FILES[which];
  fs.copyFileSync(req.file.path, target);
  fs.unlinkSync(req.file.path);
  res.json({ success: true, path: `/logo-${which}.png`, size: fs.statSync(target).size });
});
app.delete('/api/company-settings/logo/:which', requireSettingsWrite, (req, res) => {
  const which = req.params.which;
  if (!LOGO_FILES[which]) return res.status(400).json({ error: 'Invalid logo type' });
  const target = LOGO_FILES[which];
  if (fs.existsSync(target)) fs.unlinkSync(target);
  res.json({ success: true });
});
// Quick probe so the UI knows whether a logo is uploaded
app.get('/api/company-settings/logo/:which', requireView, (req, res) => {
  const which = req.params.which;
  if (!LOGO_FILES[which]) return res.status(400).json({ error: 'Invalid logo type' });
  const target = LOGO_FILES[which];
  if (!fs.existsSync(target)) return res.json({ exists: false });
  const stat = fs.statSync(target);
  res.json({ exists: true, size: stat.size, mtime: stat.mtime });
});

app.get('/api/company-settings/export', requireExport, (req, res) => {
  res.setHeader('Content-Disposition', 'attachment; filename="company-settings.json"');
  res.json(getSettingsFlat(getDb()));
});
app.post('/api/company-settings/import', requireSettingsWrite, (req, res) => {
  const count = updateSettings(getDb(), req.body.settings || {});
  res.json({ success: true, imported: count });
});

// ══════════════════════════════════════════════════════════════
// BACKUP & RESTORE  (admin only)
// ══════════════════════════════════════════════════════════════
// Used by both the Backup tab UI and the per-table Delete All
// snapshots (which capture a recovery point before any wipe). The
// snapshots live under data/backups/ so the operator can restore via
// the same Restore endpoint.

const BACKUP_DIR = path.join(process.env.SPICE_DATA_DIR || path.join(__dirname, 'data'), 'backups');
if (!fs.existsSync(BACKUP_DIR)) fs.mkdirSync(BACKUP_DIR, { recursive: true });

// Build a timestamped filename like
// `spice-etrade-backup-2026-05-24T10-15-30.db`. Strips colons that some
// filesystems don't tolerate, and the millisecond + Z suffix because the
// operator never needs sub-second precision in a backup name.
function backupFilename(prefix) {
  const iso = new Date().toISOString();
  const safe = iso.replace(/:/g, '-').replace(/\.\d+Z$/, '');
  return `${prefix}-${safe}.db`;
}

// Snapshot the current DB to disk under data/backups/. Used by
// Delete All so a misclick is recoverable. Returns the absolute path
// of the snapshot file or null on failure.
function snapshotDbToFile(prefix) {
  try {
    flushDb();
    if (!fs.existsSync(DB_PATH)) return null;
    const name = backupFilename(prefix);
    const out = path.join(BACKUP_DIR, name);
    fs.copyFileSync(DB_PATH, out);
    return out;
  } catch (e) {
    console.error('[backup] snapshot failed:', e.message);
    return null;
  }
}

// GET /api/system/backup — streams the live DB file as an attachment.
// Flushes the in-memory state first so the download captures the latest
// writes (sql.js debounces saves by 200ms).
app.get('/api/system/backup', requireUserManage, (req, res) => {
  try {
    flushDb();
    if (!fs.existsSync(DB_PATH)) return res.status(500).json({ error: 'Database file not found on disk' });
    const filename = backupFilename('spice-etrade-backup');
    res.setHeader('Content-Type', 'application/octet-stream');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    fs.createReadStream(DB_PATH).pipe(res);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// POST /api/system/restore — replaces the live DB with an uploaded file.
// Snapshots the current DB FIRST (so the restore itself is undoable),
// validates that the upload is a real SQLite file (4-byte header check),
// then swaps it in. Subsequent reads come from the new state.
app.post('/api/system/restore', requireUserManage, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  let tmpPath = req.file.path;
  let buf;
  try {
    buf = fs.readFileSync(tmpPath);
  } catch (e) {
    return res.status(500).json({ error: 'Failed to read upload: ' + e.message });
  } finally {
    // Best-effort cleanup of the multer temp file
    try { fs.unlinkSync(tmpPath); } catch (_) {}
  }
  // SQLite file header is the ASCII string "SQLite format 3\0" (16 bytes).
  // Reject anything that doesn't start with that — protects against
  // accidental upload of an XLSX, JSON, or unrelated file.
  const SQLITE_MAGIC = 'SQLite format 3\0';
  if (buf.length < 16 || buf.toString('utf8', 0, 16) !== SQLITE_MAGIC) {
    return res.status(400).json({ error: 'Not a SQLite database file (wrong header).' });
  }
  // Snapshot the current DB before we clobber it
  const backupPath = snapshotDbToFile('before-restore');
  try {
    replaceDbFromBuffer(buf);
    res.json({
      success: true,
      restoredBytes: buf.length,
      snapshotBeforeRestore: backupPath,
    });
  } catch (e) {
    res.status(500).json({ error: 'Restore failed: ' + e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// BULK DELETE ROUTES (DELETE ALL records from a given table)
// ══════════════════════════════════════════════════════════════
// Each delete-all action follows the same flow:
//   1. Preflight count (target + cascade tables)
//   2. Snapshot the DB to data/backups/ before-delete-{resource}-{ts}.db
//   3. Run the deletes inside a transaction
//   4. Insert a row into delete_log for the audit trail
//   5. Return { deleted, cascade_counts, backupPath }

// Maps a public resource name (URL slug) → its primary table + any
// cascade tables that must be wiped first (children of FKs). Used by
// both the preflight endpoint and the actual delete handler.
const DELETE_ALL_MAP = {
  traders:       { table: 'traders',     cascade: ['trader_banks'] },
  buyers:        { table: 'buyers',      cascade: [] },
  invoices:      { table: 'invoices',    cascade: [] },
  purchases:     { table: 'purchases',   cascade: [] },
  bills:         { table: 'bills',       cascade: [] },
  'debit-notes': { table: 'debit_notes', cascade: [] },
  // Auctions cascade through every transactional table — lots,
  // invoices, purchases, bills, debit_notes — because the FKs would
  // otherwise leave orphan rows. Order matters: wipe children first.
  auctions:      { table: 'auctions',    cascade: ['lots', 'invoices', 'purchases', 'bills', 'debit_notes'] },
};

// Returns { target: { table, count }, cascade: { table: count, ... } } for
// the given resource. Used by the UI to show concrete numbers in the
// "Are you sure?" prompt.
function preflightCounts(resource) {
  const spec = DELETE_ALL_MAP[resource];
  if (!spec) return null;
  const db = getDb();
  const out = {};
  out[spec.table] = (db.get(`SELECT COUNT(*) as c FROM ${spec.table}`) || { c: 0 }).c || 0;
  for (const t of spec.cascade) {
    try { out[t] = (db.get(`SELECT COUNT(*) as c FROM ${t}`) || { c: 0 }).c || 0; }
    catch (_) { out[t] = 0; }
  }
  return out;
}

app.get('/api/admin/delete-all/preflight', requireDeleteAll, (req, res) => {
  const resource = String(req.query.resource || '');
  const counts = preflightCounts(resource);
  if (!counts) return res.status(400).json({ error: 'Unknown resource: ' + resource });
  res.json({ resource, counts });
});

// Records the wipe in delete_log. Safe-by-default: if the insert fails
// the wipe still succeeds — we'd rather lose the audit entry than leave
// the operator's data in an inconsistent state.
function recordDeleteLog(req, resource, deletedCount, cascadeCounts, backupPath) {
  try {
    const db = getDb();
    const ip = (req.headers['x-forwarded-for'] || req.socket.remoteAddress || '').toString().split(',')[0].trim();
    db.run(
      `INSERT INTO delete_log (user_id, username, resource, deleted_count, cascade_counts, backup_path, ip)
       VALUES (?, ?, ?, ?, ?, ?, ?)`,
      [
        req.user ? req.user.id : null,
        req.user ? req.user.username : '',
        resource,
        deletedCount || 0,
        JSON.stringify(cascadeCounts || {}),
        backupPath || '',
        ip,
      ]
    );
  } catch (e) {
    console.error('[delete-log] insert failed:', e.message);
  }
}

// GET /api/admin/delete-log — last 200 wipes, newest first.
app.get('/api/admin/delete-log', requireDeleteAll, (req, res) => {
  try {
    const db = getDb();
    const rows = db.all(
      `SELECT id, user_id, username, resource, deleted_count,
              cascade_counts, backup_path, ip, created_at
       FROM delete_log ORDER BY id DESC LIMIT 200`
    );
    // Parse cascade_counts JSON on the way out so the client doesn't
    // have to. Tolerant of corrupt/missing values.
    const out = rows.map(r => {
      let cascade = {};
      try { cascade = r.cascade_counts ? JSON.parse(r.cascade_counts) : {}; }
      catch (_) { cascade = {}; }
      return { ...r, cascade_counts: cascade };
    });
    res.json(out);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Performs the actual delete + cascade + snapshot + log. Wrapped in a
// single try/catch so partial failures (e.g. the snapshot fails) still
// return a useful error rather than half-completing.
function performDeleteAll(req, res, resource) {
  const spec = DELETE_ALL_MAP[resource];
  if (!spec) return res.status(400).json({ error: 'Unknown resource: ' + resource });
  try {
    const db = getDb();
    const counts = preflightCounts(resource) || {};
    // Snapshot first — if this fails we don't proceed; recovery would
    // be impossible without it.
    const backupPath = snapshotDbToFile(`before-delete-${resource}`);
    if (!backupPath) {
      return res.status(500).json({ error: 'Snapshot failed — refusing to delete. Check server logs.' });
    }
    // Cascade tables first (children), then target (parent). sqlite_sequence
    // wipes are best-effort: if the table isn't in sqlite_sequence (no rows
    // ever inserted), the DELETE is a no-op. Wrapped in try/catch so a
    // missing sqlite_sequence (some empty DBs) doesn't fail the operation.
    for (const child of spec.cascade) {
      db.run(`DELETE FROM ${child}`);
      try { db.exec(`DELETE FROM sqlite_sequence WHERE name = '${child}'`); } catch (_) {}
    }
    db.run(`DELETE FROM ${spec.table}`);
    try { db.exec(`DELETE FROM sqlite_sequence WHERE name = '${spec.table}'`); } catch (_) {}
    // Flush so the snapshot we just made + the new empty state are both
    // on disk before we report success to the client.
    flushDb();
    recordDeleteLog(req, resource, counts[spec.table] || 0, counts, backupPath);
    res.json({
      success: true,
      deleted: counts[spec.table] || 0,
      cascade_counts: counts,
      backupPath,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}

// Public delete-all routes — all gated by `delete_all` capability.
app.delete('/api/traders/delete-all',     requireDeleteAll, (req, res) => performDeleteAll(req, res, 'traders'));
app.delete('/api/buyers/delete-all',      requireDeleteAll, (req, res) => performDeleteAll(req, res, 'buyers'));
app.delete('/api/invoices/delete-all',    requireDeleteAll, (req, res) => performDeleteAll(req, res, 'invoices'));
app.delete('/api/purchases/delete-all',   requireDeleteAll, (req, res) => performDeleteAll(req, res, 'purchases'));
app.delete('/api/bills/delete-all',       requireDeleteAll, (req, res) => performDeleteAll(req, res, 'bills'));
app.delete('/api/debit-notes/delete-all', requireDeleteAll, (req, res) => performDeleteAll(req, res, 'debit-notes'));
app.delete('/api/auctions/delete-all',    requireDeleteAll, (req, res) => performDeleteAll(req, res, 'auctions'));

// ══════════════════════════════════════════════════════════════
// GST LOOKUP — fetch trade name/address/state from GSTIN
// Uses gstincheck.co.in if an API key is configured in settings
// (company-config: gst_api_key). Falls back to structural validation.
// ══════════════════════════════════════════════════════════════
const STATE_CODES = {
  '01':'JAMMU AND KASHMIR','02':'HIMACHAL PRADESH','03':'PUNJAB','04':'CHANDIGARH','05':'UTTARAKHAND',
  '06':'HARYANA','07':'DELHI','08':'RAJASTHAN','09':'UTTAR PRADESH','10':'BIHAR','11':'SIKKIM',
  '12':'ARUNACHAL PRADESH','13':'NAGALAND','14':'MANIPUR','15':'MIZORAM','16':'TRIPURA','17':'MEGHALAYA',
  '18':'ASSAM','19':'WEST BENGAL','20':'JHARKHAND','21':'ODISHA','22':'CHATTISGARH','23':'MADHYA PRADESH',
  '24':'GUJARAT','25':'DAMAN AND DIU','26':'DADRA AND NAGAR HAVELI','27':'MAHARASHTRA','28':'ANDHRA PRADESH',
  '29':'KARNATAKA','30':'GOA','31':'LAKSHADWEEP','32':'KERALA','33':'TAMIL NADU','34':'PUDUCHERRY',
  '35':'ANDAMAN AND NICOBAR ISLANDS','36':'TELANGANA','37':'ANDHRA PRADESH','38':'LADAKH'
};
const GSTIN_RE = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/;

// Lazily ensure the single-row GST API credit-state table exists.
function _ensureGstApiState(db) {
  db.exec(`CREATE TABLE IF NOT EXISTS gst_api_state (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    credits_remaining INTEGER,
    credits_total INTEGER,
    plan_expires_at TEXT,
    last_checked_at TEXT,
    last_envelope TEXT
  )`);
}

// Probe a gstincheck response envelope for a remaining-credit count and
// persist it. Returns a small {credits_remaining,...} summary (or nulls).
// The provider's field name varies, so we try a list of likely keys; if
// none match we still store the raw body so the operator can paste it to
// a dev who can add the alias.
function _gstCaptureCredits(db, body) {
  _ensureGstApiState(db);
  const out = { credits_remaining: null, credits_total: null, plan_expires_at: null };
  try {
    const probe = (obj, keys) => {
      if (!obj || typeof obj !== 'object') return null;
      for (const k of Object.keys(obj)) {
        const lk = k.toLowerCase().replace(/[^a-z]/g, '');
        if (keys.includes(lk)) {
          const n = Number(obj[k]);
          if (Number.isFinite(n)) return n;
        }
      }
      return null;
    };
    const REM = ['creditsremaining', 'remainingcredits', 'creditleft', 'creditsleft', 'availablesearch', 'availablesearches', 'searchleft', 'searchesleft', 'balance', 'credit', 'credits'];
    const TOT = ['creditstotal', 'totalcredits', 'totalsearch', 'totalsearches', 'plancredits'];
    out.credits_remaining = probe(body, REM);
    if (out.credits_remaining == null && body) out.credits_remaining = probe(body.data, REM);
    out.credits_total = probe(body, TOT);
    if (out.credits_total == null && body) out.credits_total = probe(body.data, TOT);
    const exp = body && (body.expiry || body.plan_expiry || body.expires_at || body.validtill);
    if (exp) out.plan_expires_at = String(exp);
  } catch (_) {}
  const nowIso = (db.get(`SELECT datetime('now','localtime') AS t`) || {}).t || null;
  try {
    db.run(
      `INSERT INTO gst_api_state (id, credits_remaining, credits_total, plan_expires_at, last_checked_at, last_envelope)
       VALUES (1, ?, ?, ?, ?, ?)
       ON CONFLICT(id) DO UPDATE SET
         credits_remaining = COALESCE(excluded.credits_remaining, gst_api_state.credits_remaining),
         credits_total     = COALESCE(excluded.credits_total, gst_api_state.credits_total),
         plan_expires_at   = COALESCE(excluded.plan_expires_at, gst_api_state.plan_expires_at),
         last_checked_at   = excluded.last_checked_at,
         last_envelope     = excluded.last_envelope`,
      [out.credits_remaining, out.credits_total, out.plan_expires_at, nowIso, JSON.stringify(body || null)]
    );
  } catch (_) {}
  out.last_checked_at = nowIso;
  return out;
}

// GST lookup API status — credits remaining, plan expiry, last-checked
// timestamp + raw envelope. Drives the Settings → Integrations card and
// the topbar credit pill. Read-only; never triggers a live lookup.
app.get('/api/gst-lookup/status', requireView, (req, res) => {
  const db = getDb();
  _ensureGstApiState(db);
  const cfg = getSettingsFlat(db);
  const hasKey = !!(cfg.gst_api_key && String(cfg.gst_api_key).trim());
  const st = db.get('SELECT * FROM gst_api_state WHERE id = 1') || {};
  const warn_below = parseInt(cfg.gst_warn_below, 10) || 50;
  const critical_below = parseInt(cfg.gst_critical_below, 10) || 10;
  const left = st.credits_remaining == null ? null : Number(st.credits_remaining);
  let level;
  if (!hasKey || left == null) level = 'unknown';
  else if (left <= 0) level = 'exhausted';
  else if (left <= critical_below) level = 'critical';
  else if (left <= warn_below) level = 'warning';
  else level = 'ok';
  let envelope = null;
  try { envelope = st.last_envelope ? JSON.parse(st.last_envelope) : null; } catch (_) {}
  res.json({
    has_api_key: hasKey,
    level,
    credits_remaining: left,
    credits_total: st.credits_total != null ? Number(st.credits_total) : null,
    plan_expires_at: st.plan_expires_at || null,
    last_checked_at: st.last_checked_at || null,
    last_envelope: envelope,
    warn_below, critical_below,
    recharge_url: 'https://gstincheck.co.in/',
  });
});

// ══════════════════════════════════════════════════════════════
// WHATSAPP BUSINESS (Cloud API) — config storage + status + send.
// Secrets live in the single-row whatsapp_config table and are NEVER
// returned to the browser (status hands back only booleans + non-secret
// identifiers). Environment variables override DB values so a managed
// deployment can inject credentials without touching the UI.
// Sends call Meta's Graph API when fully configured; otherwise they
// return 501 so the frontend falls back to the wa.me / Web-Share flow.
// ══════════════════════════════════════════════════════════════
const WA_GRAPH = 'https://graph.facebook.com/v21.0';
function _ensureWaConfig(db) {
  db.exec(`CREATE TABLE IF NOT EXISTS whatsapp_config (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    token TEXT, phone_id TEXT, waba_id TEXT, app_secret TEXT, verify_token TEXT,
    display_number TEXT, tpl_document TEXT, tpl_document_lang TEXT,
    tpl_text TEXT, tpl_text_lang TEXT
  )`);
}
// Merge DB row with env overrides. `source` is 'env' when any secret
// comes from the environment, else 'db'.
function _waConfig(db) {
  _ensureWaConfig(db);
  const row = db.get('SELECT * FROM whatsapp_config WHERE id = 1') || {};
  const E = process.env;
  let source = 'db';
  const pick = (envKey, dbVal) => {
    if (E[envKey]) { source = 'env'; return E[envKey]; }
    return dbVal || '';
  };
  return {
    token:           pick('WHATSAPP_TOKEN', row.token),
    phoneId:         pick('WHATSAPP_PHONE_ID', row.phone_id),
    wabaId:          pick('WHATSAPP_WABA_ID', row.waba_id),
    appSecret:       pick('WHATSAPP_APP_SECRET', row.app_secret),
    verifyToken:     pick('WHATSAPP_VERIFY_TOKEN', row.verify_token),
    displayNumber:   row.display_number || '',
    tplDocument:     row.tpl_document || '',
    tplDocumentLang: row.tpl_document_lang || 'en',
    tplText:         row.tpl_text || '',
    tplTextLang:     row.tpl_text_lang || 'en',
    source,
  };
}
function _waNormPhone(p) {
  const d = String(p || '').replace(/\D/g, '');
  if (!d) return '';
  return d.length === 10 ? '91' + d : d;
}

app.get('/api/whatsapp/status', requireView, async (req, res) => {
  const db = getDb();
  const c = _waConfig(db);
  const configured = !!(c.token && c.phoneId);
  const out = {
    configured,
    hasToken: !!c.token,
    hasAppSecret: !!c.appSecret,
    hasVerifyToken: !!c.verifyToken,
    phoneId: c.phoneId,
    wabaId: c.wabaId,
    displayNumber: c.displayNumber,
    tplDocument: c.tplDocument,
    tplDocumentLang: c.tplDocumentLang,
    tplText: c.tplText,
    tplTextLang: c.tplTextLang,
    webhookReady: !!(c.verifyToken && c.appSecret),
    source: c.source,
    live: false,
    liveError: null,
    displayPhone: '',
    qualityRating: '',
  };
  if (configured) {
    try {
      const r = await fetch(
        `${WA_GRAPH}/${encodeURIComponent(c.phoneId)}?fields=display_phone_number,verified_name,quality_rating`,
        { headers: { Authorization: 'Bearer ' + c.token }, signal: AbortSignal.timeout(5000) }
      );
      const d = await r.json().catch(() => ({}));
      if (!r.ok) throw new Error((d.error && d.error.message) || ('HTTP ' + r.status));
      out.live = true;
      out.displayPhone = d.display_phone_number || c.displayNumber || '';
      out.qualityRating = d.quality_rating || '';
    } catch (e) {
      out.liveError = (e && e.message) || 'check failed';
    }
  }
  res.json(out);
});

app.put('/api/whatsapp/config', requireSettingsWrite, (req, res) => {
  const db = getDb();
  _ensureWaConfig(db);
  const b = req.body || {};
  const cur = db.get('SELECT * FROM whatsapp_config WHERE id = 1') || {};
  // Blank secret inputs mean "keep existing" — the UI sends '' to leave a
  // configured token untouched. Non-secret fields always overwrite.
  const keepIfBlank = (incoming, existing) =>
    (incoming === undefined || incoming === null || incoming === '') ? (existing || '') : String(incoming);
  const setAlways = (incoming, existing) =>
    (incoming === undefined || incoming === null) ? (existing || '') : String(incoming);
  const next = {
    token:        keepIfBlank(b.token, cur.token),
    phone_id:     setAlways(b.phoneId, cur.phone_id),
    waba_id:      setAlways(b.wabaId, cur.waba_id),
    app_secret:   keepIfBlank(b.appSecret, cur.app_secret),
    verify_token: keepIfBlank(b.verifyToken, cur.verify_token),
    display_number:    setAlways(b.displayNumber, cur.display_number),
    tpl_document:      setAlways(b.tplDocument, cur.tpl_document),
    tpl_document_lang: setAlways(b.tplDocumentLang, cur.tpl_document_lang) || 'en',
    tpl_text:          setAlways(b.tplText, cur.tpl_text),
    tpl_text_lang:     setAlways(b.tplTextLang, cur.tpl_text_lang) || 'en',
  };
  db.run(
    `INSERT INTO whatsapp_config
       (id, token, phone_id, waba_id, app_secret, verify_token, display_number,
        tpl_document, tpl_document_lang, tpl_text, tpl_text_lang)
     VALUES (1,?,?,?,?,?,?,?,?,?,?)
     ON CONFLICT(id) DO UPDATE SET
       token=excluded.token, phone_id=excluded.phone_id, waba_id=excluded.waba_id,
       app_secret=excluded.app_secret, verify_token=excluded.verify_token,
       display_number=excluded.display_number, tpl_document=excluded.tpl_document,
       tpl_document_lang=excluded.tpl_document_lang, tpl_text=excluded.tpl_text,
       tpl_text_lang=excluded.tpl_text_lang`,
    [next.token, next.phone_id, next.waba_id, next.app_secret, next.verify_token,
     next.display_number, next.tpl_document, next.tpl_document_lang, next.tpl_text, next.tpl_text_lang]
  );
  res.json({ ok: true });
});

// Low-level Graph template send. Returns { ok, error }.
async function _waSendTemplate(c, { phone, templateName, langCode, bodyParams, headerDocument }) {
  const components = [];
  if (headerDocument) {
    components.push({ type: 'header', parameters: [{ type: 'document', document: headerDocument }] });
  }
  if (bodyParams && bodyParams.length) {
    components.push({ type: 'body', parameters: bodyParams.map(t => ({ type: 'text', text: String(t == null ? '' : t) })) });
  }
  const payload = {
    messaging_product: 'whatsapp',
    to: _waNormPhone(phone),
    type: 'template',
    template: { name: templateName, language: { code: langCode || 'en' }, components },
  };
  const r = await fetch(`${WA_GRAPH}/${encodeURIComponent(c.phoneId)}/messages`, {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + c.token, 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
    signal: AbortSignal.timeout(15000),
  });
  const d = await r.json().catch(() => ({}));
  if (!r.ok) return { ok: false, error: (d.error && d.error.message) || ('HTTP ' + r.status) };
  return { ok: true, id: (d.messages && d.messages[0] && d.messages[0].id) || null };
}

app.post('/api/whatsapp/test', requireSettingsWrite, async (req, res) => {
  const db = getDb();
  const c = _waConfig(db);
  if (!c.token || !c.phoneId) return res.status(501).json({ error: 'WhatsApp not configured' });
  if (!c.tplText) return res.status(400).json({ error: 'Set a Text template name first', fallback: true });
  const phone = _waNormPhone((req.body || {}).phone);
  if (!phone) return res.status(400).json({ error: 'Phone required' });
  const company = (getSettingsFlat(db).trade_name || 'Test').toUpperCase();
  const r = await _waSendTemplate(c, {
    phone, templateName: c.tplText, langCode: c.tplTextLang,
    bodyParams: ['there', 'this is a test message from your admin console.', company],
  });
  if (!r.ok) return res.status(502).json({ error: r.error });
  res.json({ ok: true, id: r.id });
});

app.post('/api/whatsapp/send-template-text', requireView, async (req, res) => {
  const db = getDb();
  const c = _waConfig(db);
  if (!c.token || !c.phoneId) return res.status(501).json({ error: 'WhatsApp not configured' });
  if (!c.tplText) return res.status(400).json({ error: 'No text template configured', fallback: true });
  const b = req.body || {};
  const phone = _waNormPhone(b.phone);
  if (!phone) return res.status(400).json({ error: 'Phone required' });
  const params = Array.isArray(b.params) ? b.params.map(x => String(x == null ? '' : x).replace(/\s+/g, ' ').trim()) : [];
  const r = await _waSendTemplate(c, { phone, templateName: c.tplText, langCode: c.tplTextLang, bodyParams: params });
  if (!r.ok) return res.status(502).json({ error: r.error });
  res.json({ ok: true, id: r.id });
});

app.post('/api/whatsapp/send-template-document', requireView, upload.single('file'), async (req, res) => {
  const db = getDb();
  const c = _waConfig(db);
  const cleanup = () => { try { if (req.file) fs.unlinkSync(req.file.path); } catch (_) {} };
  if (!c.token || !c.phoneId) { cleanup(); return res.status(501).json({ error: 'WhatsApp not configured' }); }
  if (!c.tplDocument) { cleanup(); return res.status(400).json({ error: 'No document template configured', fallback: true }); }
  const phone = _waNormPhone((req.body || {}).phone);
  if (!phone) { cleanup(); return res.status(400).json({ error: 'Phone required' }); }
  if (!req.file) { cleanup(); return res.status(400).json({ error: 'File required' }); }
  let params = [];
  try { params = JSON.parse((req.body || {}).params || '[]'); } catch (_) {}
  const filename = (req.body || {}).filename || req.file.originalname || 'document.pdf';
  try {
    // Upload the PDF to Meta's media store first, then reference it in
    // the template header.
    const buf = fs.readFileSync(req.file.path);
    const fd = new FormData();
    fd.append('messaging_product', 'whatsapp');
    fd.append('type', 'application/pdf');
    fd.append('file', new Blob([buf], { type: 'application/pdf' }), filename);
    const up = await fetch(`${WA_GRAPH}/${encodeURIComponent(c.phoneId)}/media`, {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + c.token },
      body: fd,
      signal: AbortSignal.timeout(20000),
    });
    const upd = await up.json().catch(() => ({}));
    if (!up.ok || !upd.id) throw new Error((upd.error && upd.error.message) || 'media upload failed');
    const r = await _waSendTemplate(c, {
      phone, templateName: c.tplDocument, langCode: c.tplDocumentLang,
      bodyParams: Array.isArray(params) ? params.map(x => String(x == null ? '' : x).replace(/\s+/g, ' ').trim()) : [],
      headerDocument: { id: upd.id, filename },
    });
    if (!r.ok) return res.status(502).json({ error: r.error });
    res.json({ ok: true, id: r.id });
  } catch (e) {
    res.status(502).json({ error: (e && e.message) || 'send failed' });
  } finally {
    cleanup();
  }
});

// Webhook verification handshake (Meta GET) + delivery receipts (POST).
// No auth: Meta calls these directly. Verification compares the token
// against the configured verify_token.
app.get('/api/whatsapp/webhook', (req, res) => {
  const db = getDb();
  const c = _waConfig(db);
  const mode = req.query['hub.mode'];
  const token = req.query['hub.verify_token'];
  const challenge = req.query['hub.challenge'];
  if (mode === 'subscribe' && c.verifyToken && token === c.verifyToken) {
    return res.status(200).send(String(challenge || ''));
  }
  res.sendStatus(403);
});
app.post('/api/whatsapp/webhook', (req, res) => {
  // Delivery receipts are acknowledged but not persisted in this build.
  res.sendStatus(200);
});

app.get('/api/gst-lookup/:gstin', requireView, async (req, res) => {
  const gstin = String(req.params.gstin || '').toUpperCase().trim();
  if (!GSTIN_RE.test(gstin)) {
    return res.status(400).json({ valid: false, error: 'Invalid GSTIN format' });
  }
  const stCode = gstin.substring(0, 2);
  const pan    = gstin.substring(2, 12);
  const state  = STATE_CODES[stCode] || '';

  const cfg = getSettingsFlat(getDb());
  const apiKey = cfg.gst_api_key || '';

  // No API key → return structural details only
  if (!apiKey) {
    return res.json({
      valid: true, gstin, pan, st_code: stCode, state,
      source: 'structural',
      note: 'Set "gst_api_key" in settings to auto-fetch trade name/address.'
    });
  }

  // With API key → attempt live lookup
  try {
    const url = `https://sheet.gstincheck.co.in/check/${apiKey}/${gstin}`;
    const r = await fetch(url);
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const body = await r.json();
    // Persist whatever credit envelope the provider returned so the
    // Settings → Integrations card + topbar pill can show how many
    // searches remain. gstincheck doesn't document a single field name,
    // so probe common aliases and stash the raw body either way.
    const api = _gstCaptureCredits(db, body);
    if (body && body.flag && body.data) {
      const d = body.data;
      const addr = (d.pradr && d.pradr.addr) || {};
      return res.json({
        valid: true, gstin, pan, st_code: stCode,
        name:     d.lgnm || d.tradeNam || '',
        tradeName:d.tradeNam || d.lgnm || '',
        address:  [addr.bno, addr.bnm, addr.st, addr.loc].filter(Boolean).join(', '),
        place:    addr.dst || addr.loc || '',
        pin:      addr.pncd || '',
        state:    addr.stcd || state,
        status:   d.sts || '',
        regDate:  d.rgdt || '',
        source:   'live',
        api,
      });
    }
    return res.json({
      valid: true, gstin, pan, st_code: stCode, state,
      source: 'structural',
      note: body && body.message ? body.message : 'GST portal returned no data'
    });
  } catch (e) {
    return res.json({
      valid: true, gstin, pan, st_code: stCode, state,
      source: 'structural',
      note: 'GST lookup failed: ' + e.message
    });
  }
});

// ══════════════════════════════════════════════════════════════
// TRADERS (NAM.DBF — sellers/poolers)
// ══════════════════════════════════════════════════════════════
app.get('/api/traders', requireView, (req, res) => {
  const { search, limit, page, pageSize } = req.query;
  const db = getDb();
  // Helper: attach the `banks` array to each trader row (from trader_banks
  // table). Kept as a post-query hydration step so we don't bloat the main
  // query with joins or GROUP_CONCAT — easier to read, and the N+1 is fine
  // for the small trader counts this app handles (~few hundred max).
  const hydrateBanks = (rows) => {
    if (!rows.length) return rows;
    const ids = rows.map(r => r.id);
    const placeholders = ids.map(() => '?').join(',');
    const banks = db.all(
      `SELECT trader_id, bank_name, branch, acctnum, ifsc, holder_name
       FROM trader_banks WHERE trader_id IN (${placeholders})
       ORDER BY trader_id, id`, ids
    );
    const byTrader = new Map();
    for (const b of banks) {
      if (!byTrader.has(b.trader_id)) byTrader.set(b.trader_id, []);
      byTrader.get(b.trader_id).push(b);
    }
    for (const r of rows) r.banks = byTrader.get(r.id) || [];
    return rows;
  };
  // Paginated mode: caller passed page= and/or pageSize=. Returns
  // { rows, total, page, pageSize } so the UI can render a pager.
  // Active search filter is honored across the full table — the LIKE
  // clause runs against every row, not just the current page-window.
  // Legacy callers that only pass search= (or nothing) get the old
  // flat-array shape so existing code keeps working unchanged.
  const isPaginated = (page != null || pageSize != null);
  if (isPaginated) {
    const p  = Math.max(1, parseInt(page, 10) || 1);
    const ps = Math.max(1, Math.min(500, parseInt(pageSize, 10) || 50));
    const offset = (p - 1) * ps;
    let whereSql = '';
    let params = [];
    if (search) {
      const q = `%${search}%`;
      whereSql = 'WHERE name LIKE ? OR tel LIKE ? OR cr LIKE ? OR pan LIKE ? OR ppla LIKE ? OR aadhar LIKE ?';
      params = [q, q, q, q, q, q];
    }
    const total = (db.get(`SELECT COUNT(*) as c FROM traders ${whereSql}`, params) || { c: 0 }).c || 0;
    const rows = db.all(
      `SELECT * FROM traders ${whereSql} ORDER BY name LIMIT ? OFFSET ?`,
      [...params, ps, offset]
    );
    return res.json({ rows: hydrateBanks(rows), total, page: p, pageSize: ps });
  }
  if (search) {
    const q = `%${search}%`;
    const rows = db.all(
      `SELECT * FROM traders
       WHERE name LIKE ? OR tel LIKE ? OR cr LIKE ? OR pan LIKE ? OR ppla LIKE ? OR aadhar LIKE ?
       ORDER BY name LIMIT ?`,
      [q, q, q, q, q, q, parseInt(limit)||50]
    );
    return res.json(hydrateBanks(rows));
  }
  res.json(hydrateBanks(db.all('SELECT * FROM traders ORDER BY name LIMIT 500')));
});
// Phone lookup by exact (case-insensitive) seller name — powers the
// Payments tab WhatsApp share. Registered before /api/traders/:id (it
// has a distinct 2-segment path so there's no route conflict).
app.get('/api/traders/by-name/:name', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const nm = String(req.params.name || '').trim();
  if (!nm) return res.status(400).json({ error: 'name required' });
  const row = db.get('SELECT id, name, tel FROM traders WHERE LOWER(name) = LOWER(?) LIMIT 1', [nm]);
  if (!row) return res.status(404).json({ error: 'Not found' });
  res.json(row);
});

app.get('/api/traders/:id', requireView, (req, res) => {
  const db = getDb();
  const row = db.get('SELECT * FROM traders WHERE id = ?', [req.params.id]);
  if (!row) return res.status(404).json({ error: 'Not found' });
  // Attach banks array so the edit modal sees all bank accounts.
  // id + is_default are needed by the Payments lot-picker to resolve
  // each lot's routing account and label the default (★).
  row.banks = db.all(
    'SELECT id, trader_id, bank_name, branch, acctnum, ifsc, holder_name, is_default FROM trader_banks WHERE trader_id = ? ORDER BY is_default DESC, id',
    [row.id]
  );
  res.json(row);
});
// Sync a trader's banks array into the trader_banks table.
// Strategy: clear existing rows for this trader, reinsert. Simple and
// correct; the number of banks per trader is tiny (typically 1-3) so
// the delete+reinsert cost is negligible.
// Also mirrors the FIRST bank back into the parent traders.ifsc/acctnum/
// holder_name columns so older code paths that haven't been migrated to
// read trader_banks yet still see a valid primary account.
function syncTraderBanks(db, traderId, banks) {
  const arr = Array.isArray(banks) ? banks.filter(b => b && (b.acctnum || b.ifsc)) : [];
  db.run('DELETE FROM trader_banks WHERE trader_id = ?', [traderId]);
  for (const b of arr) {
    db.run(
      'INSERT INTO trader_banks (trader_id, bank_name, branch, acctnum, ifsc, holder_name) VALUES (?,?,?,?,?,?)',
      [traderId, b.bank_name||'', b.branch||'', String(b.acctnum||''), String(b.ifsc||''), b.holder_name||'']
    );
  }
  // Mirror first bank into traders row for legacy compatibility
  const first = arr[0] || {};
  db.run(
    'UPDATE traders SET ifsc=?, acctnum=?, holder_name=? WHERE id=?',
    [first.ifsc||'', first.acctnum||'', first.holder_name||'', traderId]
  );
}

// Duplicate-seller check: a seller is identified by PAN (the canonical
// taxpayer ID). If a different row already has the same PAN, the create
// is rejected with 409 + `{duplicate: true, existing}` so the client
// can show the operator the EXISTING row's identity instead of letting
// them silently create a second copy. Empty PAN is intentionally NOT
// gated — agriculturist sellers often have no PAN on file, and a
// blanket "no two rows with empty PAN" rule would block them.
function findDuplicateSeller(db, pan, excludeId) {
  const cleanPan = String(pan || '').trim().toUpperCase();
  if (!cleanPan) return null;
  // UPPER(TRIM(pan)) on the stored side too, so " ABC123 " and "abc123"
  // are treated as the same PAN regardless of how either was entered.
  const sql = excludeId
    ? 'SELECT id, name, pan, cr, tel FROM traders WHERE UPPER(TRIM(pan)) = ? AND id != ? LIMIT 1'
    : 'SELECT id, name, pan, cr, tel FROM traders WHERE UPPER(TRIM(pan)) = ? LIMIT 1';
  return db.get(sql, excludeId ? [cleanPan, excludeId] : [cleanPan]);
}

app.post('/api/traders', requireTraderWrite, (req, res) => {
  const t = req.body;
  const db = getDb();
  // Duplicate-PAN guard. Hard block (409) so the operator goes back
  // and edits the existing row rather than maintaining two.
  const dup = findDuplicateSeller(db, t.pan);
  if (dup) {
    return res.status(409).json({
      error: `A seller with PAN "${dup.pan}" already exists: ${dup.name || '(unnamed)'}`,
      duplicate: true,
      field: 'pan',
      existing: dup,
    });
  }
  const info = db.run(`INSERT INTO traders (name,cr,pan,tel,aadhar,padd,ppla,pin,pstate,pst_code,ifsc,acctnum,holder_name)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [t.name,t.cr||'',t.pan||'',t.tel||'',t.aadhar||'',t.padd||'',t.ppla||'',t.pin||'',t.pstate||'',t.pst_code||'',t.ifsc||'',t.acctnum||'',t.holder_name||'']);
  // If the client sent a banks array (new multi-bank UI), persist them.
  // Otherwise honor the legacy single-bank fields already inserted above.
  if (Array.isArray(t.banks)) {
    syncTraderBanks(db, info.lastInsertRowid, t.banks);
  }
  res.json({ success: true, id: info.lastInsertRowid });
});
app.put('/api/traders/:id', requireTraderWrite, (req, res) => {
  const t = req.body;
  const db = getDb();
  // Same duplicate-PAN check applied to updates, excluding the row
  // being edited. Prevents renaming a PAN to collide with a sibling.
  const dup = findDuplicateSeller(db, t.pan, parseInt(req.params.id, 10));
  if (dup) {
    return res.status(409).json({
      error: `Another seller with PAN "${dup.pan}" already exists: ${dup.name || '(unnamed)'}`,
      duplicate: true,
      field: 'pan',
      existing: dup,
    });
  }
  db.run(`UPDATE traders SET name=?,cr=?,pan=?,tel=?,aadhar=?,padd=?,ppla=?,pin=?,pstate=?,pst_code=?,ifsc=?,acctnum=?,holder_name=? WHERE id=?`,
    [t.name,t.cr||'',t.pan||'',t.tel||'',t.aadhar||'',t.padd||'',t.ppla||'',t.pin||'',t.pstate||'',t.pst_code||'',t.ifsc||'',t.acctnum||'',t.holder_name||'',req.params.id]);
  if (Array.isArray(t.banks)) {
    syncTraderBanks(db, parseInt(req.params.id), t.banks);
  }
  res.json({ success: true });
});
app.delete('/api/traders/:id', requireDelete, (req, res) => {
  const db = getDb();
  // Clear child rows first (trader_banks FK) before deleting the parent
  db.run('DELETE FROM trader_banks WHERE trader_id = ?', [req.params.id]);
  db.run('DELETE FROM traders WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// ── Quick-create trader (Lot Entry hall workflow) ───────────────
// Used by the Lot Entry tab's "Add New Seller" modal. Minimal-fields
// form for the auction-hall flow where the field user just needs to
// register a new seller fast — full edits can come later from the
// Sellers tab. Permissioned for trader_write OR lot_write so both
// office operators AND lot_entry-role users can hit it.
app.post('/api/traders/quick', requireAnyPermission('trader_write', 'lot_write'), (req, res) => {
  const t = req.body || {};
  if (!t.name || !String(t.name).trim()) {
    return res.status(400).json({ error: 'Name is required' });
  }
  const db = getDb();
  // Hard duplicate-PAN block — same logic as POST /api/traders so an
  // auction-hall user cannot accidentally re-create a seller that
  // already exists with the same PAN under a different name spelling.
  const panDup = findDuplicateSeller(db, t.pan);
  if (panDup) return res.status(409).json({
    duplicate: true, field: 'pan', existing: panDup,
    error: `A seller with PAN "${panDup.pan}" already exists: ${panDup.name || '(unnamed)'}`,
  });
  // De-dupe: if a seller with the same name AND (CR or phone) already
  // exists, return that one instead of creating a duplicate. Helps when
  // multiple field users create the same seller around the same time.
  let existing = null;
  if (t.cr && String(t.cr).trim()) {
    existing = db.get('SELECT * FROM traders WHERE name = ? AND cr = ? LIMIT 1',
      [String(t.name).trim(), String(t.cr).trim()]);
  }
  if (!existing && t.tel && String(t.tel).trim()) {
    existing = db.get('SELECT * FROM traders WHERE name = ? AND tel = ? LIMIT 1',
      [String(t.name).trim(), String(t.tel).trim()]);
  }
  if (existing) {
    return res.json({ success: true, id: existing.id, deduped: true, trader: existing });
  }
  const info = db.run(`INSERT INTO traders (name,cr,pan,tel,aadhar,padd,ppla,pin,pstate,pst_code,ifsc,acctnum,holder_name)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [
      String(t.name).trim().toUpperCase(),
      (t.cr || '').toString().trim(),
      (t.pan || '').toString().trim().toUpperCase(),
      (t.tel || '').toString().trim(),
      (t.aadhar || '').toString().trim(),
      (t.padd || '').toString().trim(),
      (t.ppla || '').toString().trim().toUpperCase(),
      (t.pin || '').toString().trim(),
      (t.pstate || 'TAMIL NADU').toString().trim().toUpperCase(),
      (t.pst_code || '33').toString().trim(),
      '', '', ''
    ]);
  const created = db.get('SELECT * FROM traders WHERE id = ?', [info.lastInsertRowid]);
  if (created) created.banks = [];
  res.json({ success: true, id: info.lastInsertRowid, trader: created });
});

// Set a bank as the trader's default. Used by the Lot Entry bank-picker
// so picking a bank on a lot save updates the trader's default for next
// time. Also syncs the legacy traders.acctnum/ifsc/holder_name fields
// since several existing exports read directly from the traders row.
app.put('/api/traders/:id/bank-default/:bankId', requireAnyPermission('trader_write', 'lot_write'), (req, res) => {
  const traderId = parseInt(req.params.id, 10);
  const bankId   = parseInt(req.params.bankId, 10);
  if (!Number.isFinite(traderId) || !Number.isFinite(bankId)) {
    return res.status(400).json({ error: 'Invalid trader or bank id' });
  }
  const db = getDb();
  const bank = db.get('SELECT * FROM trader_banks WHERE id = ? AND trader_id = ?', [bankId, traderId]);
  if (!bank) return res.status(404).json({ error: 'Bank not found for this trader' });
  // Clear is_default on every bank for this trader, then set it on the chosen one.
  db.run('UPDATE trader_banks SET is_default = 0 WHERE trader_id = ?', [traderId]);
  db.run('UPDATE trader_banks SET is_default = 1 WHERE id = ?', [bankId]);
  // Sync the legacy single-bank fields on traders so existing DBF/XLSX
  // exports read the chosen bank.
  db.run('UPDATE traders SET acctnum = ?, ifsc = ?, holder_name = ? WHERE id = ?',
    [bank.acctnum || '', bank.ifsc || '', bank.holder_name || '', traderId]);
  res.json({ success: true });
});

// ── Import Sellers from XLS/XLSX ──────────────────────────────
app.post('/api/traders/import', requireTraderWrite, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const workbook = XLSX.readFile(req.file.path);
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (!rows.length) throw new Error('File is empty');

    const db = getDb();
    const mode = req.body.mode || 'append';
    if (mode === 'replace') {
      // Wipe child rows (trader_banks FK) before parents — avoids FK error
      db.run('DELETE FROM trader_banks');
      db.run('DELETE FROM traders');
    }

    // Build flexible header map — normalize all keys to uppercase
    const mapCol = (row, ...names) => {
      for (const n of names) { if (row[n] !== undefined) return String(row[n]).trim(); }
      // Also try uppercase/lowercase variants
      const keys = Object.keys(row);
      for (const n of names) {
        const found = keys.find(k => k.toUpperCase() === n.toUpperCase());
        if (found && row[found] !== undefined) return String(row[found]).trim();
      }
      return '';
    };

    let imported = 0, skipped = 0;
    for (const row of rows) {
      const name = mapCol(row, 'NAME', 'SELLER', 'POOLER', 'TRADER');
      if (!name) { skipped++; continue; }

      const cr = mapCol(row, 'CR', 'GSTIN', 'CR_NO', 'CRNO');
      if (mode === 'append') {
        const existing = db.get('SELECT id FROM traders WHERE name = ? AND cr = ?', [name, cr]);
        if (existing) { skipped++; continue; }
      }

      db.run(`INSERT INTO traders (name,cr,pan,tel,aadhar,padd,ppla,pin,pstate,pst_code,ifsc,acctnum,holder_name) 
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [name, cr,
         mapCol(row, 'PAN', 'PAN_NO'),
         mapCol(row, 'TEL', 'PHONE', 'MOBILE', 'CONTACT'),
         mapCol(row, 'AADHAR', 'AADHAAR', 'AADHAR_NO'),
         mapCol(row, 'PADD', 'ADDRESS', 'ADD', 'ADD1', 'ADDRESS1'),
         mapCol(row, 'PPLA', 'PLACE', 'PLA', 'CITY'),
         mapCol(row, 'PIN', 'PPIN', 'PINCODE', 'ZIP'),
         mapCol(row, 'PSTATE', 'STATE'),
         mapCol(row, 'PST_CODE', 'ST_CODE', 'STATE_CODE', 'STATECODE'),
         mapCol(row, 'IFSC', 'IFS_CODE', 'IFSCODE', 'IFS'),
         mapCol(row, 'ACCTNUM', 'ACCOUNT', 'ACCNO', 'ACC_NO', 'ACCOUNT_NO', 'ACCOUNTNO'),
         mapCol(row, 'HOLDER_NAME', 'HOLDER', 'ACCOUNT_HOLDER')]);
      imported++;
    }

    fs.unlink(req.file.path, () => {});
    res.json({ success: true, imported, skipped, total: rows.length });
  } catch (e) {
    if (req.file) fs.unlink(req.file.path, () => {});
    res.status(400).json({ error: e.message });
  }
});

// ── Download Seller template XLSX ────────────────────────────
app.get('/api/traders/template', requireExport, async (req, res) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sellers');
  ws.columns = ['NAME','CR','PAN','TEL','AADHAR','PADD','PPLA','PIN','PSTATE','PST_CODE','IFSC','ACCTNUM','HOLDER_NAME']
    .map(h => ({ header: h, key: h.toLowerCase(), width: 18 }));
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
  // Add one sample row
  ws.addRow({ name: 'SAMPLE SELLER', cr: 'CR.12345', pan: 'ABCDE1234F', tel: '9876543210',
    aadhar: '', padd: '123 MAIN STREET', ppla: 'BODINAYAKANUR', pin: '625582',
    pstate: 'TAMIL NADU', pst_code: '33', ifsc: 'FDRL0001073', acctnum: '1234567890', holder_name: 'SAMPLE SELLER' });
  const buf = await wb.xlsx.writeBuffer();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="sellers-template.xlsx"');
  res.send(Buffer.from(buf));
});

// ══════════════════════════════════════════════════════════════
// BUYERS (SBL.DBF — dealers/traders)
// ══════════════════════════════════════════════════════════════
app.get('/api/buyers', requireView, (req, res) => {
  const { search, page, pageSize } = req.query;
  const db = getDb();
  // Paginated mode: caller passed page= and/or pageSize=. Returns
  // { rows, total, page, pageSize }. Search filter applies across the
  // full table (LIKE runs over every row, not just the page-window).
  // Legacy callers (no page/pageSize) keep the flat-array response so
  // existing code paths don't break.
  const isPaginated = (page != null || pageSize != null);
  if (isPaginated) {
    const p  = Math.max(1, parseInt(page, 10) || 1);
    const ps = Math.max(1, Math.min(500, parseInt(pageSize, 10) || 50));
    const offset = (p - 1) * ps;
    let whereSql = '';
    let params = [];
    if (search) {
      const q = `%${search}%`;
      whereSql = 'WHERE buyer LIKE ? OR buyer1 LIKE ? OR tel LIKE ? OR gstin LIKE ? OR pan LIKE ? OR pla LIKE ? OR ti LIKE ? OR code LIKE ?';
      params = [q, q, q, q, q, q, q, q];
    }
    const total = (db.get(`SELECT COUNT(*) as c FROM buyers ${whereSql}`, params) || { c: 0 }).c || 0;
    const rows = db.all(
      `SELECT * FROM buyers ${whereSql} ORDER BY buyer1 LIMIT ? OFFSET ?`,
      [...params, ps, offset]
    );
    return res.json({ rows, total, page: p, pageSize: ps });
  }
  if (search) {
    const q = `%${search}%`;
    return res.json(db.all(
      `SELECT * FROM buyers
       WHERE buyer LIKE ? OR buyer1 LIKE ? OR tel LIKE ? OR gstin LIKE ? OR pan LIKE ? OR pla LIKE ? OR ti LIKE ? OR code LIKE ?
       ORDER BY buyer1 LIMIT 50`,
      [q, q, q, q, q, q, q, q]
    ));
  }
  // `?all=1` lifts the default 500-row cap — used by the lot-edit Code
  // dropdown so buyer codes past the first 500 (alphabetical by buyer1)
  // still appear in the type-ahead. Backward-compatible: no `all` → cap.
  const allBuyers = req.query.all != null && !['', '0', 'false'].includes(String(req.query.all).toLowerCase());
  res.json(db.all('SELECT * FROM buyers ORDER BY buyer1' + (allBuyers ? '' : ' LIMIT 500')));
});
// Duplicate-buyer check. Buyers carry two identifiers:
//   buyer   — primary buyer code (required, e.g. "B042")
//   code    — optional short alias / mnemonic (e.g. "RSH")
// Either is a collision-worthy key, so we check both. Returns the
// EXISTING row that matches, with a `field` flag so the client can
// show "buyer code already taken" vs "short alias already taken".
// Both lookups are case-insensitive; empty values aren't gated.
function findDuplicateBuyer(db, buyer, code, excludeId) {
  const cleanBuyer = String(buyer || '').trim().toUpperCase();
  const cleanCode  = String(code  || '').trim().toUpperCase();
  if (cleanBuyer) {
    const sql = excludeId
      ? 'SELECT id, buyer, buyer1, code FROM buyers WHERE UPPER(TRIM(buyer)) = ? AND id != ? LIMIT 1'
      : 'SELECT id, buyer, buyer1, code FROM buyers WHERE UPPER(TRIM(buyer)) = ? LIMIT 1';
    const hit = db.get(sql, excludeId ? [cleanBuyer, excludeId] : [cleanBuyer]);
    if (hit) return { ...hit, field: 'buyer' };
  }
  if (cleanCode) {
    const sql = excludeId
      ? "SELECT id, buyer, buyer1, code FROM buyers WHERE UPPER(TRIM(code)) = ? AND TRIM(code) != '' AND id != ? LIMIT 1"
      : "SELECT id, buyer, buyer1, code FROM buyers WHERE UPPER(TRIM(code)) = ? AND TRIM(code) != '' LIMIT 1";
    const hit = db.get(sql, excludeId ? [cleanCode, excludeId] : [cleanCode]);
    if (hit) return { ...hit, field: 'code' };
  }
  return null;
}

app.post('/api/buyers', requireBuyerWrite, (req, res) => {
  const b = req.body;
  const db = getDb();
  // Duplicate-by-code guard. Hard block so two buyers can't share the
  // same primary code or short alias (both are used as lookup keys
  // elsewhere — duplicates would make alloc / invoice flows ambiguous).
  const dup = findDuplicateBuyer(db, b.buyer, b.code);
  if (dup) {
    return res.status(409).json({
      error: `A buyer with ${dup.field === 'buyer' ? `code "${dup.buyer}"` : `short alias "${dup.code}"`} already exists${dup.buyer1 ? `: ${dup.buyer1}` : ''}`,
      duplicate: true,
      field: dup.field,
      existing: dup,
    });
  }
  db.run(`INSERT INTO buyers (
      buyer, buyer1, code, sbl, add1, add2, pla, pin, state, st_code,
      gstin, pan, tel, ti, sale, email, tdsq,
      cbuyer1, cadd1, cadd2, cpla, cpin, cstate, cst_code, cgstin
    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [b.buyer, b.buyer1||'', b.code||'', b.sbl||'', b.add1||'', b.add2||'', b.pla||'', b.pin||'', b.state||'', b.st_code||'',
     b.gstin||'', b.pan||'', b.tel||'', b.ti||'', b.sale||'L', b.email||'', b.tdsq||'',
     b.cbuyer1||'', b.cadd1||'', b.cadd2||'', b.cpla||'', b.cpin||'', b.cstate||'', b.cst_code||'', b.cgstin||'']);
  res.json({ success: true });
});
app.put('/api/buyers/:id', requireBuyerWrite, (req, res) => {
  const b = req.body;
  const db = getDb();
  // Same duplicate guard on updates, excluding the row being edited.
  const dup = findDuplicateBuyer(db, b.buyer, b.code, parseInt(req.params.id, 10));
  if (dup) {
    return res.status(409).json({
      error: `Another buyer with ${dup.field === 'buyer' ? `code "${dup.buyer}"` : `short alias "${dup.code}"`} already exists${dup.buyer1 ? `: ${dup.buyer1}` : ''}`,
      duplicate: true,
      field: dup.field,
      existing: dup,
    });
  }
  db.run(`UPDATE buyers SET
      buyer=?, buyer1=?, code=?, sbl=?, add1=?, add2=?, pla=?, pin=?, state=?, st_code=?,
      gstin=?, pan=?, tel=?, ti=?, sale=?, email=?, tdsq=?,
      cbuyer1=?, cadd1=?, cadd2=?, cpla=?, cpin=?, cstate=?, cst_code=?, cgstin=?
    WHERE id=?`,
    [b.buyer, b.buyer1||'', b.code||'', b.sbl||'', b.add1||'', b.add2||'', b.pla||'', b.pin||'', b.state||'', b.st_code||'',
     b.gstin||'', b.pan||'', b.tel||'', b.ti||'', b.sale||'L', b.email||'', b.tdsq||'',
     b.cbuyer1||'', b.cadd1||'', b.cadd2||'', b.cpla||'', b.cpin||'', b.cstate||'', b.cst_code||'', b.cgstin||'',
     req.params.id]);
  res.json({ success: true });
});
app.delete('/api/buyers/:id', requireDelete, (req, res) => {
  getDb().run('DELETE FROM buyers WHERE id = ?', [req.params.id]); res.json({ success: true });
});

// ── Import Buyers from XLS/XLSX ───────────────────────────────
app.post('/api/buyers/import', requireBuyerWrite, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const workbook = XLSX.readFile(req.file.path);
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (!rows.length) throw new Error('File is empty');

    const db = getDb();
    const mode = req.body.mode || 'append';
    if (mode === 'replace') db.run('DELETE FROM buyers');

    const mapCol = (row, ...names) => {
      for (const n of names) { if (row[n] !== undefined) return String(row[n]).trim(); }
      const keys = Object.keys(row);
      for (const n of names) {
        const found = keys.find(k => k.toUpperCase() === n.toUpperCase());
        if (found && row[found] !== undefined) return String(row[found]).trim();
      }
      return '';
    };

    let imported = 0, skipped = 0;
    for (const row of rows) {
      // BUYER = full buyer code (primary key in lot.buyer → matches invoice lookup)
      // CODE  = short alias printed on tags (e.g. RSH, TE, SL) — used by post-auction price files
      // The two may be the same value in some files, or different. Treat them as distinct columns.
      const buyer = mapCol(row, 'BUYER', 'BUYER_CODE', 'BUYERCODE');
      const code  = mapCol(row, 'CODE', 'SHORT_CODE', 'ALIAS');
      if (!buyer && !code) { skipped++; continue; }
      // If BUYER column missing, fall back to CODE, then trade name
      const buyerVal = buyer || code || mapCol(row, 'BUYER1', 'TRADE_NAME', 'TRADENAME', 'NAME');
      if (!buyerVal) { skipped++; continue; }

      if (mode === 'append') {
        const existing = db.get('SELECT id FROM buyers WHERE buyer = ?', [buyerVal]);
        if (existing) { skipped++; continue; }
      }

      db.run(`INSERT INTO buyers (
        buyer, buyer1, code, sbl, add1, add2, pla, pin, state, st_code,
        gstin, pan, tel, ti, sale, email, tdsq,
        cbuyer1, cadd1, cadd2, cpla, cpin, cstate, cst_code, cgstin
      ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [buyerVal,
         mapCol(row, 'BUYER1', 'TRADE_NAME', 'TRADENAME', 'NAME'),
         code,
         mapCol(row, 'SBL', 'SBLNO'),
         mapCol(row, 'ADD1', 'ADDRESS1', 'ADDRESS'),
         mapCol(row, 'ADD2', 'ADDRESS2'),
         mapCol(row, 'PLA', 'PLACE', 'CITY'),
         mapCol(row, 'PIN', 'PINCODE', 'ZIP'),
         mapCol(row, 'STATE'),
         mapCol(row, 'ST_CODE', 'STATE_CODE', 'STATECODE'),
         mapCol(row, 'GSTIN', 'GST', 'GSTNO', 'GST_NO'),
         mapCol(row, 'PAN', 'PAN_NO'),
         mapCol(row, 'TEL', 'PHONE', 'MOBILE', 'CONTACT'),
         mapCol(row, 'TI'),
         mapCol(row, 'SALE', 'SALE_TYPE') || 'L',
         mapCol(row, 'EMAIL', 'E_MAIL', 'MAIL'),
         mapCol(row, 'TDSQ', 'TDS_Q', 'TDS'),
         // Consignee (ship-to) details
         mapCol(row, 'CBUYER1', 'CONSIGNEE', 'CONSIGNEE_NAME'),
         mapCol(row, 'CADD1', 'CONS_ADD1', 'CONSIGNEE_ADDRESS1'),
         mapCol(row, 'CADD2', 'CONS_ADD2', 'CONSIGNEE_ADDRESS2'),
         mapCol(row, 'CPLA', 'CONS_PLA', 'CONSIGNEE_PLACE'),
         mapCol(row, 'CPIN', 'CONS_PIN', 'CONSIGNEE_PIN'),
         mapCol(row, 'CSTATE', 'CONS_STATE'),
         mapCol(row, 'CST_CODE', 'CONS_ST_CODE'),
         mapCol(row, 'CGSTIN', 'CONS_GSTIN')]);
      imported++;
    }

    fs.unlink(req.file.path, () => {});
    res.json({ success: true, imported, skipped, total: rows.length });
  } catch (e) {
    if (req.file) fs.unlink(req.file.path, () => {});
    res.status(400).json({ error: e.message });
  }
});

// ── Download Buyer template XLSX ─────────────────────────────
app.get('/api/buyers/template', requireExport, async (req, res) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Buyers');
  ws.columns = ['BUYER','BUYER1','ADD1','ADD2','PLA','PIN','STATE','ST_CODE','GSTIN','PAN','TEL','TI','SALE']
    .map(h => ({ header: h, key: h.toLowerCase(), width: 18 }));
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
  ws.addRow({ buyer: 'ABC', buyer1: 'ABC TRADERS', add1: '10 MARKET ROAD', add2: '', pla: 'KUMILY',
    pin: '685509', state: 'KERALA', st_code: '32', gstin: '32AABCT1234L1ZP', pan: 'AABCT1234L', tel: '9876543210', ti: '', sale: 'L' });
  const buf = await wb.xlsx.writeBuffer();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="buyers-template.xlsx"');
  res.send(Buffer.from(buf));
});

// ══════════════════════════════════════════════════════════════
// BUSINESS MODE FILTERING
// ══════════════════════════════════════════════════════════════
// Auctions are tagged with `mode` at creation time (e-Trade or e-Auction).
// Lists filter by the current company-settings mode so users only see
// trades — and the downstream invoices/purchases/bills/debit notes/
// payments that belong to those trades — that match the mode they're
// currently working in.
//
// Two helpers below:
//   currentBusinessMode(db) — reads the active mode setting.
//   modeWhere(prefix)       — returns a SQL fragment + params suitable
//                             for AND-ing into a query that joins or
//                             filters by an auctions row. NULL/empty
//                             mode columns always match (legacy data
//                             stays visible during the soft cutover).
function currentBusinessMode(db) {
  try {
    const row = db.get("SELECT value FROM company_settings WHERE key = 'business_mode'");
    return row && row.value ? String(row.value) : '';
  } catch (_) { return ''; }
}
// User-facing noun for the active business mode.
// e-Trade → "Trade" / "Trades"; e-Auction → "Auction" / "Auctions".
// Accepts either a cfg object (preferred — caller already loaded
// settings) or a db handle. Used in filenames, PDF headers, XLSX sheet
// names, and any other human-readable string emitted by the server.
function termAuction(cfgOrDb, plural, lower) {
  let mode = '';
  if (cfgOrDb && typeof cfgOrDb.business_mode === 'string') {
    mode = cfgOrDb.business_mode;
  } else if (cfgOrDb && typeof cfgOrDb.get === 'function') {
    mode = currentBusinessMode(cfgOrDb);
  }
  if (!mode) mode = 'e-Auction';
  const isTrade = (mode === 'e-Trade');
  const word = isTrade ? (plural ? 'Trades' : 'Trade') : (plural ? 'Auctions' : 'Auction');
  return lower ? word.toLowerCase() : word;
}
// Returns { sql, params } that filter rows to the current business mode.
// `prefix` is the SQL prefix that resolves to the auctions row's `mode`
// column — usually 'auctions.mode' (when the query already JOINs auctions)
// or 'a.mode' (when the JOIN uses alias `a`). For tables that don't store
// auction_id, use a subquery: prefix='(SELECT mode FROM auctions WHERE id=invoices.auction_id)'.
function modeWhereClause(db, prefix) {
  const mode = currentBusinessMode(db);
  if (!mode) return { sql: '', params: [] };       // no filter when mode unset
  return {
    sql: ` AND (${prefix} = ? OR ${prefix} IS NULL OR ${prefix} = '')`,
    params: [mode],
  };
}

// ══════════════════════════════════════════════════════════════
// AUCTIONS
// ══════════════════════════════════════════════════════════════
app.get('/api/auctions', requireView, (req, res) => {
  const db = getDb();
  const mw = modeWhereClause(db, 'auctions.mode');
  // Build the WHERE clause manually (no `WHERE 1=1` shortcut because the
  // mode filter is the only optional one and we want clean SQL).
  const whereSql = mw.sql ? 'WHERE 1=1' + mw.sql : '';
  const rows = db.all(
    `SELECT *, (SELECT COUNT(*) FROM lots WHERE auction_id=auctions.id) as lot_count
     FROM auctions ${whereSql}
     ORDER BY date DESC, ano DESC LIMIT 100`,
    mw.params
  );
  res.json(withFmtDate(rows));
});
app.post('/api/auctions', requireAuctionWrite, (req, res) => {
  const { ano, date, crop_type, state } = req.body;
  const db = getDb();
  const d = normalizeDate(date);
  // Mode is ALWAYS stamped from the current setting at creation time —
  // never honored from the request body. Locking the mode means an
  // operator can't accidentally create a trade in the wrong mode by
  // sending a stale value from a stale browser tab.
  const mode = currentBusinessMode(db);
  db.run('INSERT INTO auctions (ano,date,crop_type,state,mode) VALUES (?,?,?,?,?)',
    [ano, d, crop_type||'ASP', state||'TAMIL NADU', mode]);
  const created = db.get('SELECT id FROM auctions WHERE ano = ? AND date = ? ORDER BY id DESC LIMIT 1', [ano, d]);
  res.json({ success: true, id: created ? created.id : null, mode });
});
app.put('/api/auctions/:id', requireAuctionWrite, (req, res) => {
  // Edits NEVER change mode — it's locked at creation. We just refuse to
  // accept a mode field in the body, and don't include it in the UPDATE.
  // The auction's mode follows its trade for its entire lifetime.
  const { ano, date, crop_type, state } = req.body;
  getDb().run('UPDATE auctions SET ano=?, date=?, crop_type=?, state=? WHERE id=?',
    [ano, normalizeDate(date), crop_type||'ASP', state||'TAMIL NADU', req.params.id]);
  res.json({ success: true });
});
app.delete('/api/auctions/:id', requireDelete, (req, res) => {
  const db = getDb();
  // Cascade through every child table (lots, allocations) before parent.
  db.run('DELETE FROM lots WHERE auction_id = ?', [req.params.id]);
  db.run('DELETE FROM lot_allocations WHERE auction_id = ?', [req.params.id]);
  db.run('DELETE FROM auctions WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// ══════════════════════════════════════════════════════════════
// LOT ALLOCATIONS
// ══════════════════════════════════════════════════════════════
// Per-auction, per-branch lot-number ranges. The Auctions tab opens
// a modal (Edit + Reassign tabs) that drives these endpoints. Edit
// rewrites the whole allocation set; Reassign moves an unused range
// from one branch to another by re-bucketing the affected rows.
//
// Lot number format: anything matching ^[A-Za-z]*\d+$ (optional alpha
// prefix + digits). We enumerate by parsing the digit tail. Lot
// numbers with mixed prefix WITHIN a range are rejected — start_lot
// and end_lot must share the same prefix.

// Parse "001" → {prefix:'', num:1, padLen:3}, "A12" → {prefix:'A', num:12, padLen:2}.
// Returns null on invalid input so the caller can validate.
function parseLotNo(s) {
  const m = String(s || '').trim().match(/^([A-Za-z]*)(\d+)$/);
  if (!m) return null;
  return { prefix: m[1].toUpperCase(), num: parseInt(m[2], 10), padLen: m[2].length };
}
// Build a normalized lot number from a prefix + integer + pad length.
function buildLotNo(prefix, num, padLen) {
  return prefix + String(num).padStart(padLen, '0');
}
// Enumerate every lot number in [start, end] inclusive. Throws if
// start/end don't parse, have different prefixes, or end < start.
function enumerateRange(startLot, endLot) {
  const s = parseLotNo(startLot);
  const e = parseLotNo(endLot);
  if (!s) throw new Error(`Invalid start lot "${startLot}"`);
  if (!e) throw new Error(`Invalid end lot "${endLot}"`);
  if (s.prefix !== e.prefix) throw new Error(`Start and end must share the same prefix (${startLot} vs ${endLot})`);
  if (e.num < s.num) throw new Error(`End lot (${endLot}) must be >= start lot (${startLot})`);
  // Pad to the longer of the two so "1" and "100" produce "001..100"
  const pad = Math.max(s.padLen, e.padLen);
  const out = [];
  for (let n = s.num; n <= e.num; n++) out.push(buildLotNo(s.prefix, n, pad));
  return out;
}

// True iff lotNo is in [startLot..endLot] with the same prefix. Used by
// the Lot Entry validate-lot endpoint to gate against out-of-allocation
// inputs.
function isLotInRange(lotNo, startLot, endLot) {
  const lot = parseLotNo(lotNo);
  const s = parseLotNo(startLot);
  const e = parseLotNo(endLot);
  if (!lot || !s || !e) return false;
  if (lot.prefix !== s.prefix || s.prefix !== e.prefix) return false;
  return lot.num >= s.num && lot.num <= e.num;
}

// GET /api/auctions/:id/allocations — list. Empty array when none.
app.get('/api/auctions/:id/allocations', requireView, (req, res) => {
  const db = getDb();
  const rows = db.all(
    `SELECT id, branch, start_lot, end_lot FROM lot_allocations
     WHERE auction_id = ? ORDER BY branch, id`,
    [parseInt(req.params.id, 10)]
  );
  res.json({ allocations: rows });
});

// POST /api/auctions/:id/allocations — bulk-replace. Body: { allocations: [{branch, start_lot, end_lot}, ...] }.
// Validation:
//   • Every range must parse cleanly (start/end share prefix, end >= start).
//   • Within a branch, ranges must NOT overlap.
//   • If an existing range is being dropped AND saved lots fall inside
//     it, the request is rejected. The operator must delete the lots
//     first (so they aren't silently orphaned from any branch).
app.post('/api/auctions/:id/allocations', requireAuctionWrite, (req, res) => {
  const auctionId = parseInt(req.params.id, 10);
  const next = Array.isArray(req.body.allocations) ? req.body.allocations : [];
  const db = getDb();
  try {
    // Normalize + parse every range up-front so the validation errors
    // include the bad input verbatim.
    const ranges = next.map((r, i) => {
      const branch = String(r.branch || '').trim().toUpperCase();
      const startLot = String(r.start_lot || '').trim();
      const endLot   = String(r.end_lot || '').trim();
      if (!branch) throw new Error(`Row ${i + 1}: branch is required`);
      if (!startLot || !endLot) throw new Error(`Row ${i + 1}: start and end lots required`);
      const lots = enumerateRange(startLot, endLot);   // throws on bad data
      return { branch, startLot, endLot, lots };
    });
    // Overlap check within each branch
    const byBranch = new Map();
    for (const r of ranges) {
      if (!byBranch.has(r.branch)) byBranch.set(r.branch, new Set());
      const seen = byBranch.get(r.branch);
      for (const lot of r.lots) {
        if (seen.has(lot)) throw new Error(`Overlap in ${r.branch}: lot ${lot} appears in two ranges`);
        seen.add(lot);
      }
    }
    // Build set of lot_no values across the NEW allocations so we can
    // tell whether each existing saved-lot still has a home.
    const nextLotsByBranch = new Map();
    for (const r of ranges) {
      if (!nextLotsByBranch.has(r.branch)) nextLotsByBranch.set(r.branch, new Set());
      for (const l of r.lots) nextLotsByBranch.get(r.branch).add(l);
    }
    // Saved lots for this auction — any that fall outside the new
    // allocation set are "orphaned" and we refuse to save.
    const savedLots = db.all(
      `SELECT lot_no, branch FROM lots WHERE auction_id = ?`,
      [auctionId]
    );
    const orphaned = [];
    for (const sl of savedLots) {
      const br = String(sl.branch || '').trim().toUpperCase();
      const lot = String(sl.lot_no || '').trim();
      const set = nextLotsByBranch.get(br);
      if (!set || !set.has(lot)) orphaned.push(`${br || '(no branch)'} #${lot}`);
    }
    if (orphaned.length) {
      const sample = orphaned.slice(0, 6).join(', ') + (orphaned.length > 6 ? `, …+${orphaned.length - 6} more` : '');
      throw new Error(`Cannot save: ${orphaned.length} saved lot(s) would fall outside the new allocations — delete those lots first or extend the ranges. Examples: ${sample}`);
    }
    // Replace the allocation set atomically.
    db.run('DELETE FROM lot_allocations WHERE auction_id = ?', [auctionId]);
    for (const r of ranges) {
      db.run(
        'INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?,?,?,?)',
        [auctionId, r.branch, r.startLot, r.endLot]
      );
    }
    res.json({ success: true, saved: ranges.length });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

// GET /api/auctions/:id/allocation-stats — the rich payload the modal
// uses to render chips. For each branch + range, returns every lot
// number in the range with `used: true|false` and (when used) the
// seller name. Server-side enumeration avoids the client needing to
// know how lot numbers are constructed.
app.get('/api/auctions/:id/allocation-stats', requireView, (req, res) => {
  const auctionId = parseInt(req.params.id, 10);
  const db = getDb();
  // Pull every saved lot for this auction, keyed by branch+lot_no →
  // seller. We do one query instead of N (one per range) because the
  // total lot count per auction is small.
  const saved = db.all(
    `SELECT lot_no, branch, name FROM lots WHERE auction_id = ?`,
    [auctionId]
  );
  const savedMap = new Map();
  for (const r of saved) {
    const k = String(r.branch || '').trim().toUpperCase() + '::' + String(r.lot_no || '').trim();
    savedMap.set(k, r.name || '');
  }
  const rows = db.all(
    `SELECT id, branch, start_lot, end_lot FROM lot_allocations
     WHERE auction_id = ? ORDER BY branch, id`,
    [auctionId]
  );
  // Group ranges by branch. Each branch's `ranges` array carries the
  // chip-level detail the UI renders.
  const byBranch = new Map();
  for (const row of rows) {
    const br = String(row.branch || '').trim().toUpperCase();
    if (!byBranch.has(br)) byBranch.set(br, { branch: br, total: 0, used: 0, ranges: [] });
    const entry = byBranch.get(br);
    let lots = [];
    try {
      lots = enumerateRange(row.start_lot, row.end_lot).map(lotNo => {
        const key = br + '::' + lotNo;
        const seller = savedMap.get(key);
        // `state` is what the reassign tile UI keys colors off of:
        //   'free'      → allocated but no lot saved (assignable)
        //   'booked'    → a lot already exists, can't reassign
        const used = savedMap.has(key);
        return {
          lot: lotNo,
          used,
          seller: used ? seller : '',
          state: used ? 'booked' : 'allocated',
          booked: used,
        };
      });
    } catch (e) {
      // A corrupted range (bad start/end) gets logged but doesn't
      // crash the stats endpoint — the UI still shows the branch.
      console.error(`[allocation-stats] range parse failed for auction ${auctionId} branch ${br}:`, e.message);
      lots = [];
    }
    const rangeTotal = lots.length;
    const rangeUsed  = lots.filter(l => l.used).length;
    entry.total += rangeTotal;
    entry.used  += rangeUsed;
    entry.ranges.push({
      start: row.start_lot,
      end:   row.end_lot,
      total: rangeTotal,
      used:  rangeUsed,
      lots,
    });
  }
  res.json({ stats: Array.from(byBranch.values()) });
});

// POST /api/auctions/:id/reassign-lots — move the unused [start..end]
// range from one branch to another. Body: { from_branch, to_branch,
// start_lot, end_lot }. Refuses to move a range that has saved lots
// (use the Edit panel + manual lot deletion for that). The source
// branch's covering range is split around the moved range; the dest
// branch gains a new range (or extends an existing one).
app.post('/api/auctions/:id/reassign-lots', requireAuctionWrite, (req, res) => {
  const auctionId = parseInt(req.params.id, 10);
  const fromBranch = String(req.body.from_branch || '').trim().toUpperCase();
  const toBranch   = String(req.body.to_branch || '').trim().toUpperCase();
  const startLot   = String(req.body.start_lot || '').trim();
  const endLot     = String(req.body.end_lot || '').trim();
  if (!fromBranch || !toBranch) return res.status(400).json({ error: 'from_branch and to_branch required' });
  if (fromBranch === toBranch)  return res.status(400).json({ error: 'FROM and TO branches must differ' });
  if (!startLot || !endLot)     return res.status(400).json({ error: 'start_lot and end_lot required' });
  const db = getDb();
  try {
    const moving = enumerateRange(startLot, endLot);     // throws on bad input
    const movingSet = new Set(moving);
    // Refuse to move a range that's booked.
    const booked = db.all(
      `SELECT lot_no FROM lots WHERE auction_id = ? AND UPPER(branch) = ?`,
      [auctionId, fromBranch]
    ).filter(r => movingSet.has(String(r.lot_no || '').trim()));
    if (booked.length) {
      return res.status(400).json({
        error: `Cannot reassign — ${booked.length} lot(s) in this range already have sellers. Delete those lots first.`,
        bookedLots: booked.map(r => r.lot_no),
      });
    }
    // Load FROM branch's covering allocation. We need to find the
    // single range that contains ALL of [start..end]. If no single
    // range covers it, we refuse — multi-range reassign would split
    // the responsibility and isn't worth the complexity.
    const fromRanges = db.all(
      `SELECT id, start_lot, end_lot FROM lot_allocations
       WHERE auction_id = ? AND UPPER(branch) = ?`,
      [auctionId, fromBranch]
    );
    let host = null;
    let hostLots = null;
    for (const r of fromRanges) {
      try {
        const lots = enumerateRange(r.start_lot, r.end_lot);
        if (moving.every(l => lots.includes(l))) { host = r; hostLots = lots; break; }
      } catch (_) {}
    }
    if (!host) {
      return res.status(400).json({ error: `No allocation on ${fromBranch} covers the range ${startLot}-${endLot}` });
    }
    // Split the host range around the moved chunk:
    //   leftover = hostLots − moving
    //   then collapse consecutive lots into sub-ranges.
    const movedSet = new Set(moving);
    const keep = hostLots.filter(l => !movedSet.has(l));
    // Group `keep` into contiguous chunks (using the same enumeration
    // logic so we preserve prefix + padding).
    const chunks = [];
    if (keep.length) {
      let chunk = [keep[0]];
      for (let i = 1; i < keep.length; i++) {
        const prev = parseLotNo(chunk[chunk.length - 1]);
        const curr = parseLotNo(keep[i]);
        if (prev && curr && prev.prefix === curr.prefix && (curr.num === prev.num + 1)) {
          chunk.push(keep[i]);
        } else {
          chunks.push(chunk);
          chunk = [keep[i]];
        }
      }
      chunks.push(chunk);
    }
    // Apply: delete the host range, insert the leftover chunks, insert
    // the moved range under the dest branch.
    db.run('DELETE FROM lot_allocations WHERE id = ?', [host.id]);
    for (const ch of chunks) {
      db.run(
        'INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?,?,?,?)',
        [auctionId, fromBranch, ch[0], ch[ch.length - 1]]
      );
    }
    db.run(
      'INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?,?,?,?)',
      [auctionId, toBranch, startLot, endLot]
    );
    res.json({
      success: true,
      message: `Moved ${moving.length} lot(s) from ${fromBranch} to ${toBranch}`,
      moved: moving.length,
      leftoverChunks: chunks.length,
    });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

// POST /api/auctions/:id/allocations/auto-fill — synthesize allocation
// rows from existing saved lots. Useful when a trade was imported via
// Excel without ever having allocations configured: every imported lot
// has a branch tag, but no `lot_allocations` row covers it. This
// endpoint scans `lots`, groups by (branch, lot prefix), and inserts
// a wide catch-all range per group sized to span every existing lot.
// Existing allocations are NEVER overwritten — only gaps are filled.
app.post('/api/auctions/:id/allocations/auto-fill', requireAuctionWrite, (req, res) => {
  const auctionId = parseInt(req.params.id, 10);
  const db = getDb();
  try {
    // Gather existing allocations into a quick "is this lot covered?" map.
    const existing = db.all(
      `SELECT branch, start_lot, end_lot FROM lot_allocations WHERE auction_id = ?`,
      [auctionId]
    );
    const coveredByBranch = new Map();
    for (const r of existing) {
      const br = String(r.branch || '').trim().toUpperCase();
      if (!coveredByBranch.has(br)) coveredByBranch.set(br, new Set());
      try {
        for (const lot of enumerateRange(r.start_lot, r.end_lot)) {
          coveredByBranch.get(br).add(lot);
        }
      } catch (_) { /* skip malformed existing range */ }
    }
    // Group existing lots by branch+prefix, tracking the numeric tail range.
    const lots = db.all(
      `SELECT lot_no, branch FROM lots WHERE auction_id = ?`,
      [auctionId]
    );
    const groups = new Map();   // key = `${branch}::${prefix}` → { branch, prefix, minN, maxN, padLen }
    for (const l of lots) {
      const br = String(l.branch || '').trim().toUpperCase();
      const parsed = parseLotNo(l.lot_no);
      if (!br || !parsed) continue;
      const key = `${br}::${parsed.prefix}`;
      if (!groups.has(key)) {
        groups.set(key, { branch: br, prefix: parsed.prefix, minN: parsed.num, maxN: parsed.num, padLen: parsed.padLen });
      } else {
        const g = groups.get(key);
        if (parsed.num < g.minN) g.minN = parsed.num;
        if (parsed.num > g.maxN) g.maxN = parsed.num;
        if (parsed.padLen > g.padLen) g.padLen = parsed.padLen;
      }
    }
    // For each group, check whether every saved lot is already covered.
    // If even one isn't, insert a catch-all range that spans min..max.
    let created = 0;
    for (const g of groups.values()) {
      const set = coveredByBranch.get(g.branch) || new Set();
      let allCovered = true;
      for (let n = g.minN; n <= g.maxN; n++) {
        if (!set.has(buildLotNo(g.prefix, n, g.padLen))) { allCovered = false; break; }
      }
      if (allCovered) continue;
      db.run(
        'INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?,?,?,?)',
        [auctionId, g.branch, buildLotNo(g.prefix, g.minN, g.padLen), buildLotNo(g.prefix, g.maxN, g.padLen)]
      );
      created++;
    }
    res.json({ success: true, created });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Validate a single lot number against (a) duplicates and (b) the
// branch's allocation. Returns { valid: bool, error?: string }. Used
// by Lot Entry's lot-no input as a server-side sanity check before save.
app.get('/api/auctions/:id/validate-lot', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const lotNo = String(req.query.lot_no || '').trim();
  const branch = String(req.query.branch || '').trim();
  if (!lotNo) return res.json({ valid: false, error: 'Enter lot number' });

  const dup = db.get('SELECT id FROM lots WHERE auction_id = ? AND lot_no = ?', [auctionId, lotNo]);
  if (dup) return res.json({ valid: false, error: 'Lot #' + lotNo + ' already exists' });

  const allocs = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ?', [auctionId, branch]);
  if (allocs.length > 0) {
    const inRange = allocs.some(a => isLotInRange(lotNo, a.start_lot, a.end_lot));
    if (!inRange) {
      const ranges = allocs.map(a => a.start_lot + '-' + a.end_lot).join(', ');
      return res.json({ valid: false, error: 'Outside allocation (' + ranges + ')' });
    }
  }
  res.json({ valid: true });
});

// Next available lot number for a branch — used by Lot Entry to
// auto-suggest after every save and after a seller is picked. Walks the
// branch's allocation ranges in order, returning the first lot_no not
// already saved to the auction.
app.get('/api/auctions/:id/next-lot/:branch', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const branch = req.params.branch;
  const allocations = db.all(
    'SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ? ORDER BY start_lot',
    [auctionId, branch]
  );
  if (!allocations.length) return res.json({ next_lot: null, error: 'No allocation for this branch' });

  const usedLots = db.all('SELECT lot_no FROM lots WHERE auction_id = ?', [auctionId]).map(l => l.lot_no);
  const usedSet = new Set(usedLots);

  for (const a of allocations) {
    const s = parseLotNo(a.start_lot);
    const e = parseLotNo(a.end_lot);
    if (!s || !e) continue;
    for (let n = s.num; n <= e.num; n++) {
      const lotNo = buildLotNo(s.prefix, n, s.padLen);
      if (!usedSet.has(lotNo)) return res.json({ next_lot: lotNo });
    }
  }
  res.json({ next_lot: null, error: 'All lots in this branch are used' });
});

// ── Import Auction + Lots from XLS/XLSX (replaces APPA.PRG) ──
app.post('/api/auctions/import', requireAuctionWrite, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const workbook = XLSX.readFile(req.file.path);
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (!rows.length) throw new Error('File is empty');

    const db = getDb();
    const mode = req.body.mode || 'full'; // 'full' = new lots, 'price' = update price/buyer only

    const mapCol = (row, ...names) => {
      for (const n of names) { if (row[n] !== undefined) return String(row[n]).trim(); }
      const keys = Object.keys(row);
      for (const n of names) {
        const found = keys.find(k => k.toUpperCase() === n.toUpperCase());
        if (found && row[found] !== undefined) return String(row[found]).trim();
      }
      return '';
    };
    const mapNum = (row, ...names) => parseFloat(mapCol(row, ...names)) || 0;

    // If user specified ano/date in the form → that OVERRIDES every row (single-auction import)
    // Otherwise → resolve auction per row from its own ANO/DATE columns (multi-auction import)
    const overrideAno = req.body.ano;
    const overrideDate = normalizeDate(req.body.date);
    const cropType = req.body.crop_type || mapCol(rows[0], 'CRPT', 'CROP_TYPE', 'CROPTYPE') || 'ASP';
    const state = req.body.state || mapCol(rows[0], 'STATE') || 'TAMIL NADU';

    // Cache of resolved auctions so we don't query the DB for every row
    const auctionCache = new Map(); // key = "ano|date" → {id, ano, date}
    // Stamp the current business mode on auctions created by import, matching
    // the manual-create path (POST /api/auctions). Without this, imported
    // auctions get mode='' and leak into both e-Trade and e-Auction views.
    const importMode = currentBusinessMode(db);
    const resolveAuction = (ano, dateStr) => {
      const key = `${ano}|${dateStr}`;
      if (auctionCache.has(key)) return auctionCache.get(key);
      let auc = db.get('SELECT * FROM auctions WHERE ano = ? AND date = ?', [ano, dateStr]);
      if (!auc) {
        db.run('INSERT INTO auctions (ano, date, crop_type, state, mode) VALUES (?,?,?,?,?)',
          [ano, dateStr || new Date().toISOString().slice(0, 10), cropType, state, importMode]);
        auc = db.get('SELECT * FROM auctions WHERE ano = ? AND date = ? ORDER BY id DESC LIMIT 1', [ano, dateStr]);
      }
      auctionCache.set(key, auc);
      return auc;
    };

    // Pre-validate: if no form override AND no ANO column anywhere, bail early with a clear message
    if (!overrideAno) {
      const firstAno = rows.length ? mapCol(rows[0], 'ANO', 'TNO', 'TRADE', 'TRADE_NO', 'TRADENO') : '';
      if (!firstAno) throw new Error('No ANO column found in file. Add ANO/TRADE/TRADE_NO column, or specify Trade No in the form to override.');
    }

    let imported = 0, updated = 0, skipped = 0;
    const skipReasons = []; // [{row, lot, reason}]
    const auctionStats = new Map(); // key = "ano|date" → count

    // Helper: check if row is completely empty (all values blank/undefined)
    const isBlankRow = (row) => {
      const vals = Object.values(row);
      return !vals.length || vals.every(v => v === '' || v === null || v === undefined);
    };

    if (mode === 'price') {
      // Price update mode — only update price, amount, code, buyer fields on existing lots
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowNum = i + 2; // +2 because: 1-based + header row
        if (isBlankRow(row)) continue; // truly empty rows — don't count as skipped
        
        // Resolve this row's auction (form override OR read from row's ANO/DATE columns)
        const rowAno = overrideAno || mapCol(row, 'ANO', 'TNO', 'TRADE', 'TRADE_NO', 'TRADENO');
        // Read raw (un-stringified) DATE cell so Excel Date objects / serial numbers normalize correctly
        const rawDate = row.DATE !== undefined ? row.DATE
                      : row.date !== undefined ? row.date
                      : row.TRADE_DATE !== undefined ? row.TRADE_DATE
                      : '';
        const rowDate = overrideDate || normalizeDate(rawDate);
        if (!rowAno) { skipped++; skipReasons.push({row: rowNum, lot: '', reason: 'Missing ANO/TRADE_NO for this row'}); continue; }
        const auc = resolveAuction(rowAno, rowDate);
        const auctionId = auc.id;
        
        const lotNo = mapCol(row, 'LOT', 'LOT_NO', 'LOTNO');
        if (!lotNo) { skipped++; skipReasons.push({row: rowNum, lot: '', reason: 'Missing LOT / LOT_NO column value'}); continue; }
        
        // (price mode continues below in original code)
        const existing = db.get('SELECT id FROM lots WHERE auction_id = ? AND lot_no = ?', [auctionId, lotNo]);
        if (!existing) { skipped++; skipReasons.push({row: rowNum, lot: lotNo, reason: `Lot ${lotNo} does not exist in Trade ${rowAno} (price-update requires existing lot)`}); continue; }

        try {
          // Parse each field from the row using generous synonyms so different XLSX layouts work
          const price = mapNum(row, 'PRICE');
          const qty   = mapNum(row, 'QTY', 'QUANTITY', 'WEIGHT', 'WT');
          const bag   = mapNum(row, 'BAG', 'BAGS', 'NO_OF_BAGS');
          // If file didn't provide AMOUNT, compute qty × price (common in post-auction price sheets)
          let amount  = mapNum(row, 'AMOUNT', 'AMT', 'VALUE', 'TOTAL');
          if (!amount && qty && price) amount = qty * price;

          // Build UPDATE dynamically — only touch fields the file provided, so a sparse "price-only"
          // file doesn't wipe pre-existing bag/qty/buyer values
          const sets = []; const vals = [];
          if (row.PRICE !== undefined || row.price !== undefined) { sets.push('price=?');  vals.push(price); }
          if (amount)                                              { sets.push('amount=?'); vals.push(amount); }
          if (row.QTY !== undefined || row.qty !== undefined)      { sets.push('qty=?');    vals.push(qty); }
          if (row.BAG !== undefined || row.bag !== undefined ||
              row.BAGS !== undefined || row.bags !== undefined)    { sets.push('bags=?');   vals.push(bag); }
          const codeVal  = mapCol(row, 'CODE', 'BUYER_CODE');
          if (codeVal)                                             { sets.push('code=?');   vals.push(codeVal); }

          // Auto-resolve short CODE (e.g. RSH, TE, SL) to the full buyer record.
          // Priority: explicit BUYER/BIDDER column in file → matching buyers.code → matching buyers.ti → matching buyers.buyer
          let resolvedBuyer  = mapCol(row, 'BUYER', 'BIDDER', 'BUYER_NAME');
          let resolvedBuyer1 = mapCol(row, 'BUYER1', 'TRADE_NAME', 'TRADENAME');
          let resolvedSale   = mapCol(row, 'SALE', 'SALE_TYPE');

          if (codeVal && (!resolvedBuyer || !resolvedBuyer1)) {
            // Look the code up in the buyers master (case-insensitive match on code, ti, or buyer)
            const match = db.get(
              `SELECT buyer, buyer1, sale FROM buyers
               WHERE UPPER(TRIM(code))  = UPPER(TRIM(?))
                  OR UPPER(TRIM(ti))    = UPPER(TRIM(?))
                  OR UPPER(TRIM(buyer)) = UPPER(TRIM(?))
               LIMIT 1`,
              [codeVal, codeVal, codeVal]
            );
            if (match) {
              if (!resolvedBuyer)  resolvedBuyer  = match.buyer  || '';
              if (!resolvedBuyer1) resolvedBuyer1 = match.buyer1 || '';
              if (!resolvedSale)   resolvedSale   = match.sale   || '';
            } else {
              // Not found — record a warning but DON'T fail the row (we still update price/qty/bag)
              skipReasons.push({
                row: rowNum, lot: lotNo,
                reason: `Warning: CODE "${codeVal}" not found in Buyers master — price updated but buyer NOT assigned. Add this code to Buyers to enable invoicing.`
              });
            }
          }

          if (resolvedBuyer)  { sets.push('buyer=?');  vals.push(resolvedBuyer); }
          if (resolvedBuyer1) { sets.push('buyer1=?'); vals.push(resolvedBuyer1); }
          if (resolvedSale)   { sets.push('sale=?');   vals.push(resolvedSale); }

          if (!sets.length) { skipped++; skipReasons.push({row: rowNum, lot: lotNo, reason: 'Row has no updatable fields (price/qty/bag/code/buyer/sale)'}); continue; }

          vals.push(existing.id);
          db.run(`UPDATE lots SET ${sets.join(', ')} WHERE id=?`, vals);
          updated++;
          const key = `${rowAno}|${rowDate}`;
          auctionStats.set(key, (auctionStats.get(key) || 0) + 1);
        } catch (e) {
          skipped++;
          skipReasons.push({row: rowNum, lot: lotNo, reason: `DB error: ${e.message}`});
        }
      }
    } else {
      // Full import — insert new lots (skip if lot_no already exists for this auction)
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowNum = i + 2;
        if (isBlankRow(row)) continue;
        
        // Resolve this row's auction (form override OR read from row's ANO/DATE columns)
        const rowAno = overrideAno || mapCol(row, 'ANO', 'TNO', 'TRADE', 'TRADE_NO', 'TRADENO');
        // Read raw DATE cell (may be Date object or Excel serial number) and normalize
        const rawDate = row.DATE !== undefined ? row.DATE
                      : row.date !== undefined ? row.date
                      : row.TRADE_DATE !== undefined ? row.TRADE_DATE
                      : '';
        const rowDate = overrideDate || normalizeDate(rawDate);
        if (!rowAno) { skipped++; skipReasons.push({row: rowNum, lot: '', reason: 'Missing ANO/TRADE_NO for this row'}); continue; }
        const auc = resolveAuction(rowAno, rowDate);
        const auctionId = auc.id;
        
        const lotNo = mapCol(row, 'LOT', 'LOT_NO', 'LOTNO');
        if (!lotNo) { skipped++; skipReasons.push({row: rowNum, lot: '', reason: 'Missing LOT / LOT_NO column value'}); continue; }

        const existing = db.get('SELECT id FROM lots WHERE auction_id = ? AND lot_no = ?', [auctionId, lotNo]);
        if (existing) { skipped++; skipReasons.push({row: rowNum, lot: lotNo, reason: `Duplicate — lot ${lotNo} already exists in Trade ${rowAno}`}); continue; }

        // Try to find trader by name for linking
        const sellerName = mapCol(row, 'NAME', 'SELLER', 'POOLER', 'TRADER');
        let traderId = null;
        if (sellerName) {
          const trader = db.get('SELECT id FROM traders WHERE name = ?', [sellerName]);
          if (trader) traderId = trader.id;
        }

        try {
          db.run(`INSERT INTO lots (auction_id, lot_no, crop, grade, crpt, branch, state, trader_id,
            name, padd, ppla, ppin, pstate, pst_code, cr, pan, tel, aadhar,
            bags, litre, qty, price, amount, code, buyer, buyer1, sale, invo,
            pqty, prate, puramt, com, sertax, cgst, sgst, igst, advance, balance, bilamt, user_id)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
            [auctionId, lotNo,
             mapCol(row, 'CROP'),
             mapCol(row, 'GRADE'),
             mapCol(row, 'CRPT', 'CROP_TYPE') || cropType,
             mapCol(row, 'BR', 'BRANCH', 'DEPOT'),
             mapCol(row, 'STATE') || state,
             traderId,
             sellerName,
             mapCol(row, 'PADD', 'ADDRESS', 'ADD', 'ADD1'),
             mapCol(row, 'PPLA', 'PLACE', 'PLA'),
             mapCol(row, 'PPIN', 'PIN', 'PINCODE'),
             mapCol(row, 'PSTATE'),
             mapCol(row, 'PST_CODE', 'ST_CODE', 'STATE_CODE'),
             mapCol(row, 'CR', 'GSTIN', 'CR_NO'),
             mapCol(row, 'PAN', 'PAN_NO'),
             mapCol(row, 'TEL', 'PHONE', 'MOBILE'),
             mapCol(row, 'AADHAR', 'AADHAAR'),
             mapNum(row, 'BAG', 'BAGS'),
             mapCol(row, 'LITRE', 'LITRE_WT'),
             mapNum(row, 'QTY', 'QUANTITY', 'NET_QTY'),
             mapNum(row, 'PRICE', 'RATE'),
             mapNum(row, 'AMOUNT'),
             mapCol(row, 'CODE', 'BUYER_CODE'),
             mapCol(row, 'BUYER', 'BIDDER'),
             mapCol(row, 'BUYER1', 'TRADE_NAME', 'TRADENAME'),
             mapCol(row, 'SALE', 'SALE_TYPE'),
             mapCol(row, 'INVO', 'INVOICE'),
             mapNum(row, 'PQTY', 'PUR_QTY'),
             mapNum(row, 'PRATE', 'PUR_RATE'),
             mapNum(row, 'PURAMT', 'PUR_AMT', 'PURCHASE_AMT'),
             mapNum(row, 'COM', 'COMMISSION'),
             mapNum(row, 'SERTAX', 'HPC'),
             mapNum(row, 'CGST'),
             mapNum(row, 'SGST'),
             mapNum(row, 'IGST'),
             mapNum(row, 'ADVANCE', 'DISCOUNT'),
             mapNum(row, 'BALANCE', 'PAYABLE'),
             mapNum(row, 'BILAMT', 'BILL_AMT'),
             mapCol(row, 'USER_ID', 'USER') || 'import']);
          imported++;
          const key = `${rowAno}|${rowDate}`;
          auctionStats.set(key, (auctionStats.get(key) || 0) + 1);
        } catch (e) {
          skipped++;
          skipReasons.push({row: rowNum, lot: lotNo, reason: `DB error: ${e.message}`});
        }
      }
    }

    // Build auction breakdown for the response
    const auctionBreakdown = [];
    for (const [key, count] of auctionStats) {
      const [ano, date] = key.split('|');
      const auc = auctionCache.get(key);
      auctionBreakdown.push({ id: auc?.id, ano, date, count });
    }
    auctionBreakdown.sort((a,b) => String(a.ano).localeCompare(String(b.ano), undefined, {numeric:true}));

    fs.unlink(req.file.path, () => {});
    res.json({ 
      success: true, 
      imported, updated, skipped, total: rows.length,
      auctionCount: auctionBreakdown.length,
      auctionBreakdown,
      skipReasons 
    });
  } catch (e) {
    if (req.file) fs.unlink(req.file.path, () => {});
    res.status(400).json({ error: e.message });
  }
});

// ── Download Trade/Lots template XLSX ──────────────────────
// Filename adapts to the active business mode so the user sees a
// "trade-lots-template.xlsx" download in e-Trade and
// "auction-lots-template.xlsx" in e-Auction.
app.get('/api/auctions/template', requireExport, async (req, res) => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Lots');
  ws.columns = ['ANO','DATE','LOT','CROP','GRADE','CRPT','BR','STATE','NAME','PADD','PPLA','PPIN','PSTATE','PST_CODE',
    'CR','PAN','TEL','AADHAR','BAG','LITRE','QTY','PRICE','AMOUNT','CODE','BUYER','BUYER1','SALE','INVO',
    'PQTY','PRATE','PURAMT','CGST','SGST','IGST','ADVANCE','BALANCE']
    .map(h => ({ header: h, key: h.toLowerCase(), width: h.length < 5 ? 8 : 14 }));
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };
  ws.addRow({ ano: '1', date: '2026-04-15', lot: '001', crop: '', grade: '1', crpt: 'ASP', br: 'VANDANMEDU',
    state: 'TAMIL NADU', name: 'SAMPLE SELLER', padd: '123 MAIN ST', ppla: 'KUMILY', ppin: '685509',
    pstate: 'KERALA', pst_code: '32', cr: 'CR.001', pan: 'ABCDE1234F', tel: '9876543210', aadhar: '',
    bag: 5, litre: '380', qty: 100.567, price: 0, amount: 0, code: '', buyer: '', buyer1: '', sale: '', invo: '',
    pqty: 0, prate: 0, puramt: 0, cgst: 0, sgst: 0, igst: 0, advance: 0, balance: 0 });
  const buf = await wb.xlsx.writeBuffer();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  const tName = termAuction(getDb(), false, true);   // 'trade' or 'auction'
  res.setHeader('Content-Disposition', `attachment; filename="${tName}-lots-template.xlsx"`);
  res.send(Buffer.from(buf));
});

// ══════════════════════════════════════════════════════════════
// LOTS (CPA1.DBF — main data)
// ══════════════════════════════════════════════════════════════
app.get('/api/lots/:auctionId', requireView, (req, res) => {
  const { branch, name, buyer, search } = req.query;
  // Correlated subquery (not LEFT JOIN) to avoid any risk of row duplication
  // if the same buyer code exists multiple times in the buyers table.
  let q = `SELECT lots.*,
             (SELECT b.code FROM buyers b WHERE b.buyer = lots.buyer LIMIT 1) AS buyer_code
           FROM lots
           WHERE lots.auction_id = ?`;
  const p = [req.params.auctionId];
  if (branch) { q += ' AND lots.branch = ?'; p.push(branch); }
  if (name)   { q += ' AND lots.name LIKE ?'; p.push(`%${name}%`); }
  if (buyer)  { q += ' AND lots.buyer = ?'; p.push(buyer); }
  // Free-text search across lot_no / seller name / buyer code / buyer
  // short alias / invoice no / branch — every column a user would
  // type when looking for a specific lot. LIKE is case-insensitive in
  // SQLite for ASCII; matches anywhere in the cell.
  if (search) {
    const q2 = `%${search}%`;
    q += ` AND (lots.lot_no LIKE ? OR lots.name LIKE ? OR lots.buyer LIKE ?
                 OR (SELECT b.code FROM buyers b WHERE b.buyer = lots.buyer LIMIT 1) LIKE ?
                 OR lots.invo LIKE ? OR lots.branch LIKE ?)`;
    p.push(q2, q2, q2, q2, q2, q2);
  }
  // Paginated mode: caller passed paginated=1 (with optional limit/offset).
  // Used by Lot Entry's recent-entries panel where the lot list can be
  // hundreds of rows. Sort newest-first so the just-saved lot appears at
  // the top of page 1. Returns { rows, total } so the client can paint a
  // pager. Legacy callers (no paginated flag) still get the flat array
  // sorted by lot_no ASC.
  const db = getDb();
  if (req.query.paginated === '1' || req.query.paginated === 'true') {
    const limit  = Math.max(1, Math.min(500, parseInt(req.query.limit, 10)  || 25));
    const offset = Math.max(0,            parseInt(req.query.offset, 10) || 0);
    // COUNT query re-uses the same WHERE clause shape but selects COUNT
    // instead of lots.*. Build it from the same filter helper variables
    // so any future filter additions land in both queries together.
    const countQ = q.replace(
      /^SELECT[\s\S]*?FROM lots\s+WHERE/,
      'SELECT COUNT(*) AS c FROM lots WHERE'
    );
    const total = (db.get(countQ, p) || { c: 0 }).c || 0;
    q += ' ORDER BY lots.id DESC LIMIT ? OFFSET ?';
    const rows = db.all(q, [...p, limit, offset]);
    return res.json({ rows, total, limit, offset });
  }
  q += ' ORDER BY lots.lot_no';
  res.json(db.all(q, p));
});

app.post('/api/lots', requireLotWrite, (req, res) => {
  const l = req.body;
  // reserved_price is persisted unconditionally — server doesn't gate
  // on flag_reserved_price because flipping the flag later mustn't
  // wipe values that the operator already entered. UI hides the input
  // when the flag is off so 0 is what flows in by default.
  const reservedPrice = Number(l.reserved_price);
  const _db = getDb();
  _db.run(`INSERT INTO lots (auction_id,lot_no,crop,grade,crpt,reserved_price,branch,state,trader_id,name,padd,ppla,ppin,pstate,pst_code,cr,pan,tel,aadhar,bags,litre,qty,gross_wt,sample_wt,moisture,user_id)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [l.auction_id,l.lot_no,l.crop||'',l.grade||'',l.crpt||'',Number.isFinite(reservedPrice)?reservedPrice:0,l.branch||'',l.state||'TAMIL NADU',l.trader_id||null,l.name||'',l.padd||'',l.ppla||'',l.ppin||'',l.pstate||'',l.pst_code||'',l.cr||'',l.pan||'',l.tel||'',l.aadhar||'',l.bags||0,l.litre||'',l.qty||0,l.gross_wt||0,l.sample_wt||0,l.moisture||'',l.user_id||'']);
  // New lot in this trade → reconciliation is stale.
  pcClearGate(_db, l.auction_id);
  res.json({ success: true });
});

// Returns true if `req.user` has the admin role. Admin bypasses the
// lock gate — they can edit/delete locked lots and run unlock.
function isAdminUser(req) {
  return req.user && req.user.role === 'admin';
}
// Returns the lot row's lock state (or null if the row doesn't exist).
// Used by the PUT/DELETE gates below.
function getLotLock(db, lotId) {
  const r = db.get('SELECT id, locked_at, locked_by FROM lots WHERE id = ?', [parseInt(lotId, 10)]);
  return r || null;
}

app.put('/api/lots/:id', requireLotWrite, (req, res) => {
  const db = getDb();
  // Lock gate: a locked lot is editable only by admins. Server returns
  // 423 Locked so the client can distinguish from a 403 permission
  // refusal. The client surfaces the message in a toast.
  const lock = getLotLock(db, req.params.id);
  if (lock && lock.locked_at && !isAdminUser(req)) {
    return res.status(423).json({
      error: `This lot is locked${lock.locked_by ? ' by ' + lock.locked_by : ''} (${lock.locked_at}). Ask an admin to unlock it before editing.`,
      locked: true,
    });
  }
  const l = req.body; const sets = []; const vals = [];
  // Withdrawn lots (code='WD') carry no sale: force the price (and amount)
  // to 0 so every derived figure recomputes to 0. Done server-side so all
  // edit paths (modal, inline, future bulk) behave identically. The code is
  // normalised to canonical 'WD' to match the `code === 'WD'` checks used
  // across reporting and invoice eligibility.
  const codeIsWD = String(l.code || '').trim().toUpperCase() === 'WD';
  if (codeIsWD) { l.code = 'WD'; l.price = 0; l.amount = 0; }
  for (const [k,v] of Object.entries(l)) {
    // `locked_at` / `locked_by` are managed by the lock/unlock
    // endpoints — never let the generic update slot rewrite them.
    if (k !== 'id' && k !== 'auction_id' && k !== 'created_at' && k !== 'locked_at' && k !== 'locked_by') {
      sets.push(`${k}=?`); vals.push(v);
    }
  }
  vals.push(req.params.id);
  // Capture auction_id BEFORE the write so a price/code edit drops
  // the price-check gate to 'stale' (re-surfacing the lots-tab banner).
  const _pcRow = db.get('SELECT auction_id FROM lots WHERE id = ?', [req.params.id]);
  db.run(`UPDATE lots SET ${sets.join(',')} WHERE id=?`, vals);
  // After a withdrawal, recompute the derived columns from the now-zeroed
  // price so prate/puramt/GST/balance all read 0 without a Calculate All.
  if (codeIsWD) {
    const fresh = db.get('SELECT * FROM lots WHERE id = ?', [req.params.id]);
    if (fresh) {
      const cfg = getSettingsFlat(db);
      const c = calculateLot(fresh, cfg);
      db.run(
        `UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
        [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,req.params.id]
      );
    }
  }
  if (_pcRow && _pcRow.auction_id) pcClearGate(db, _pcRow.auction_id);
  res.json({ success: true });
});

app.delete('/api/lots/:id', requireDelete, (req, res) => {
  const db = getDb();
  // Same lock gate as PUT — server is the authoritative gate; the
  // client's row-tinting is purely cosmetic.
  const lock = getLotLock(db, req.params.id);
  if (lock && lock.locked_at && !isAdminUser(req)) {
    return res.status(423).json({
      error: `This lot is locked${lock.locked_by ? ' by ' + lock.locked_by : ''} (${lock.locked_at}). Ask an admin to unlock it before deleting.`,
      locked: true,
    });
  }
  const _pcRow = db.get('SELECT auction_id FROM lots WHERE id = ?', [req.params.id]);
  db.run('DELETE FROM lots WHERE id = ?', [req.params.id]);
  if (_pcRow && _pcRow.auction_id) pcClearGate(db, _pcRow.auction_id);
  res.json({ success: true });
});

// ══════════════════════════════════════════════════════════════
// BULK LOT ACTIONS
// ══════════════════════════════════════════════════════════════
// Each endpoint accepts { ids: [42, 43, ...] } and operates on the
// matching lots. Lock-aware: non-admins can't touch locked rows; the
// server returns the count of rows actually changed AND a list of
// skipped (locked) lot_no's so the client can show what was protected.

// Filter a list of lot ids to the subset that the caller is allowed to
// mutate (admin → all; non-admin → only unlocked). Returns:
//   { allowedIds: [...], skipped: [{id, lot_no}, ...] }
// `skipped` contains the locked lots the caller wasn't allowed to touch.
function partitionLotsByLock(db, ids, req) {
  if (!ids || !ids.length) return { allowedIds: [], skipped: [] };
  const placeholders = ids.map(() => '?').join(',');
  const rows = db.all(
    `SELECT id, lot_no, locked_at FROM lots WHERE id IN (${placeholders})`,
    ids
  );
  if (isAdminUser(req)) return { allowedIds: rows.map(r => r.id), skipped: [] };
  const allowedIds = [], skipped = [];
  for (const r of rows) {
    if (r.locked_at) skipped.push({ id: r.id, lot_no: r.lot_no });
    else allowedIds.push(r.id);
  }
  return { allowedIds, skipped };
}

// POST /api/lots/bulk-buyer — set buyer code on multiple lots. The
// new buyer's `buyer1` (trade name) and `sale` type are auto-resolved
// from the buyers master, so the operator only has to type the code.
// Body: { ids:[...], buyer:'B042' }
app.post('/api/lots/bulk-buyer', requireLotWrite, (req, res) => {
  const ids = Array.isArray(req.body.ids) ? req.body.ids.map(n => parseInt(n, 10)).filter(Boolean) : [];
  const buyerCode = String(req.body.buyer || '').trim();
  if (!ids.length) return res.status(400).json({ error: 'ids array required' });
  if (!buyerCode) return res.status(400).json({ error: 'buyer code required' });
  const db = getDb();
  // Resolve buyer details once so all updates share one read.
  // Case-insensitive match because operators type codes in any case.
  const buyer = db.get(
    'SELECT buyer, buyer1, sale FROM buyers WHERE UPPER(buyer) = ? LIMIT 1',
    [buyerCode.toUpperCase()]
  );
  if (!buyer) return res.status(404).json({ error: `No buyer registered with code "${buyerCode}". Add them in the Buyers tab first.` });
  const part = partitionLotsByLock(db, ids, req);
  if (!part.allowedIds.length) {
    return res.status(423).json({
      error: 'Every selected lot is locked. Ask an admin to unlock first.',
      locked: true,
      skipped: part.skipped,
    });
  }
  // Stamp buyer, buyer1, and sale on each lot in one batched UPDATE.
  // We also clear `invo` so the lot is treated as un-invoiced again
  // (it'll need a fresh invoice after a buyer reassignment).
  const placeholders = part.allowedIds.map(() => '?').join(',');
  db.run(
    `UPDATE lots SET buyer=?, buyer1=?, sale=?, invo='' WHERE id IN (${placeholders})`,
    [buyer.buyer, buyer.buyer1 || '', buyer.sale || 'L', ...part.allowedIds]
  );
  // Stale price-check for every affected trade.
  const _aids = db.all(`SELECT DISTINCT auction_id FROM lots WHERE id IN (${placeholders})`, part.allowedIds);
  for (const r of _aids) pcClearGate(db, r.auction_id);
  res.json({
    success: true,
    updated: part.allowedIds.length,
    buyer: buyer.buyer,
    buyer1: buyer.buyer1 || '',
    sale: buyer.sale || 'L',
    skipped: part.skipped,
  });
});

// POST /api/lots/bulk-grade — set grade on multiple lots, then run
// calculateLot on each so prate/puramt/discount reflect the new grade
// without forcing the operator to click Calculate All. Body:
// { ids:[...], grade:'1' }
app.post('/api/lots/bulk-grade', requireLotWrite, (req, res) => {
  const ids = Array.isArray(req.body.ids) ? req.body.ids.map(n => parseInt(n, 10)).filter(Boolean) : [];
  const grade = String(req.body.grade || '').trim();
  if (!ids.length) return res.status(400).json({ error: 'ids array required' });
  if (!grade)     return res.status(400).json({ error: 'grade required' });
  const db = getDb();
  const cfg = getSettingsFlat(db);
  const part = partitionLotsByLock(db, ids, req);
  if (!part.allowedIds.length) {
    return res.status(423).json({
      error: 'Every selected lot is locked. Ask an admin to unlock first.',
      locked: true,
      skipped: part.skipped,
    });
  }
  // Stamp grade, then recalc (only the rows that already had a price
  // entered; un-priced lots have nothing to recompute).
  const placeholders = part.allowedIds.map(() => '?').join(',');
  db.run(`UPDATE lots SET grade=? WHERE id IN (${placeholders})`, [grade, ...part.allowedIds]);
  const toCalc = db.all(`SELECT * FROM lots WHERE id IN (${placeholders}) AND amount > 0`, part.allowedIds);
  let recalced = 0;
  for (const lot of toCalc) {
    const c = calculateLot(lot, cfg);
    db.run(
      `UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]
    );
    recalced++;
  }
  const _aids = db.all(`SELECT DISTINCT auction_id FROM lots WHERE id IN (${placeholders})`, part.allowedIds);
  for (const r of _aids) pcClearGate(db, r.auction_id);
  res.json({
    success: true,
    updated: part.allowedIds.length,
    recalced,
    grade,
    skipped: part.skipped,
  });
});

// POST /api/lots/bulk-delete — delete multiple lots. Lock-aware.
app.post('/api/lots/bulk-delete', requireDelete, (req, res) => {
  const ids = Array.isArray(req.body.ids) ? req.body.ids.map(n => parseInt(n, 10)).filter(Boolean) : [];
  if (!ids.length) return res.status(400).json({ error: 'ids array required' });
  const db = getDb();
  const part = partitionLotsByLock(db, ids, req);
  if (!part.allowedIds.length) {
    return res.status(423).json({
      error: 'Every selected lot is locked. Ask an admin to unlock first.',
      locked: true,
      skipped: part.skipped,
    });
  }
  const placeholders = part.allowedIds.map(() => '?').join(',');
  const _aids = db.all(`SELECT DISTINCT auction_id FROM lots WHERE id IN (${placeholders})`, part.allowedIds);
  db.run(`DELETE FROM lots WHERE id IN (${placeholders})`, part.allowedIds);
  for (const r of _aids) pcClearGate(db, r.auction_id);
  res.json({
    success: true,
    deleted: part.allowedIds.length,
    skipped: part.skipped,
  });
});

// POST /api/lots/lock — lock multiple lots. Open to anyone with
// lot_write (the role allowed to create them). Idempotent: already-
// locked lots stay locked with their original timestamp.
app.post('/api/lots/lock', requireLotWrite, (req, res) => {
  const ids = Array.isArray(req.body.ids) ? req.body.ids.map(n => parseInt(n, 10)).filter(Boolean) : [];
  if (!ids.length) return res.status(400).json({ error: 'ids array required' });
  const db = getDb();
  const username = (req.user && req.user.username) || '';
  // Only stamp `locked_at` on currently-unlocked rows so re-locking
  // preserves the original lock timestamp.
  const placeholders = ids.map(() => '?').join(',');
  const info = db.run(
    `UPDATE lots SET locked_at = datetime('now','localtime'), locked_by = ?
     WHERE id IN (${placeholders}) AND (locked_at IS NULL OR locked_at = '')`,
    [username, ...ids]
  );
  res.json({
    success: true,
    locked: info.changes || 0,
    requested: ids.length,
  });
});

// POST /api/lots/unlock — admin only. Clears the lock so non-admins
// can edit/delete the lot again. We DO NOT log an audit row here; the
// admin doing the unlock is implicit (only they can hit this route).
app.post('/api/lots/unlock', requireAdmin, (req, res) => {
  const ids = Array.isArray(req.body.ids) ? req.body.ids.map(n => parseInt(n, 10)).filter(Boolean) : [];
  if (!ids.length) return res.status(400).json({ error: 'ids array required' });
  const db = getDb();
  const placeholders = ids.map(() => '?').join(',');
  const info = db.run(
    `UPDATE lots SET locked_at = NULL, locked_by = NULL
     WHERE id IN (${placeholders}) AND locked_at IS NOT NULL`,
    ids
  );
  res.json({
    success: true,
    unlocked: info.changes || 0,
    requested: ids.length,
  });
});

// ══════════════════════════════════════════════════════════════
// PRICE LIST (BEFORE) — code mapping tool
// ══════════════════════════════════════════════════════════════
// Operator workflow:
//   1. Export an empty Price List (Before) sheet (Exports tab)
//   2. Print → buyers write their TRADE NAME (and prices) by hand
//   3. Type the trade names back into the file
//   4. Upload here → server resolves CODE from the buyers master
//   5. Preview the matches; download the updated file
//   6. Feed the downloaded file into Lots → Price Import
//
// Two endpoints share the same parsing/matching code:
//   POST /api/price-list/map-preview  → JSON summary only
//   POST /api/price-list/map-download → updated XLSX (Buffer)
//
// `ExcelJS` is used end-to-end so the brand header / total row /
// column widths from the original export survive the round-trip.
//
// UI surface (sidebar entry + Lots toolbar button + the Price List
// Mapping tab) is gated client-side by flag_price_list_mapping AND
// business_mode === 'e-Auction'. These endpoints stay live regardless
// so existing data/round-trips keep working if the flag is toggled.
function _plLocateColumns(ws) {
  // Find the header row containing both "TRADE NAME" and "CODE".
  // Match is case-insensitive + whitespace-tolerant so renamed headers
  // (e.g. "Trade Name", "trade_name") still resolve.
  const normalize = s => String(s == null ? '' : s).trim().toUpperCase().replace(/[\s_\-]+/g, ' ');
  let headerRow = 0, tradeCol = 0, codeCol = 0, anoCol = 0, dateCol = 0, lotCol = 0;
  const maxRow = ws.rowCount || 0;
  for (let r = 1; r <= maxRow; r++) {
    const row = ws.getRow(r);
    const cells = {};
    row.eachCell({ includeEmpty: false }, (cell, col) => {
      cells[normalize(cell.value)] = col;
    });
    if (cells['TRADE NAME'] && cells['CODE']) {
      headerRow = r;
      tradeCol = cells['TRADE NAME'];
      codeCol = cells['CODE'];
      anoCol = cells['AUCTION NO'] || cells['ANO'] || cells['TNO'] || 0;
      dateCol = cells['DATE'] || 0;
      lotCol = cells['LOT'] || cells['LOT NO'] || cells['LOTNO'] || 0;
      break;
    }
  }
  return { headerRow, tradeCol, codeCol, anoCol, dateCol, lotCol };
}
function _plBuildTradeIndex(db) {
  // Pre-index every buyer by their trade name so a 1000-row file is one
  // DB query, not 1000. We index BOTH buyer1 and buyer because operators
  // sometimes write the full buyer-code string in the TRADE NAME column.
  const buyers = db.all('SELECT id, buyer, buyer1, code, ti, sale, gstin FROM buyers');
  const idx = new Map();
  const push = (key, row) => {
    if (!key) return;
    const k = key.trim().toUpperCase();
    if (!k) return;
    if (!idx.has(k)) idx.set(k, []);
    // Avoid duplicate entries when buyer === buyer1.
    const arr = idx.get(k);
    if (!arr.some(b => b.id === row.id)) arr.push(row);
  };
  for (const b of buyers) {
    push(b.buyer1, b);
    push(b.buyer, b);
  }
  return idx;
}
async function _plProcessFile(filePath) {
  // Returns { wb, ws, cols, perRow: [{row, tradeName, status, pickedCode, candidates}], summary }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('No worksheet found');
  const cols = _plLocateColumns(ws);
  if (!cols.headerRow) throw new Error('Could not find a row with both "TRADE NAME" and "CODE" columns — is this a Price List (Before) file?');

  const idx = _plBuildTradeIndex(getDb());
  const perRow = [];
  let matched = 0, ambiguous = 0, unmatched = 0, blank = 0;
  const maxRow = ws.rowCount || 0;
  for (let r = cols.headerRow + 1; r <= maxRow; r++) {
    const row = ws.getRow(r);
    const tradeCell = row.getCell(cols.tradeCol);
    // Skip a TOTAL footer row — the export marks it by leaving TRADE NAME
    // blank and putting the literal "TOTAL" in the first text column. We
    // only fill rows that look like data: must have some non-empty
    // identifier in the row (lot/ano/trade name).
    const tradeRaw = tradeCell.value;
    const tradeName = String(tradeRaw == null ? '' : tradeRaw).trim();
    const lotVal = cols.lotCol ? String(row.getCell(cols.lotCol).value || '').trim() : '';
    if (!tradeName && !lotVal) continue;
    const entry = {
      row: r,
      lot: lotVal,
      ano: cols.anoCol ? String(row.getCell(cols.anoCol).value || '').trim() : '',
      date: cols.dateCol ? String(row.getCell(cols.dateCol).value || '').trim() : '',
      tradeName,
      candidates: [],
      pickedCode: '',
      status: 'blank',
    };
    if (!tradeName) {
      entry.status = 'blank';
      blank++;
      perRow.push(entry);
      continue;
    }
    const key = tradeName.toUpperCase();
    const cands = idx.get(key) || [];
    entry.candidates = cands.map(b => ({
      id: b.id, code: b.code, buyer: b.buyer, buyer1: b.buyer1, sale: b.sale, gstin: b.gstin,
    }));
    if (cands.length === 0) {
      entry.status = 'unmatched';
      unmatched++;
    } else if (cands.length === 1) {
      entry.status = 'matched';
      entry.pickedCode = cands[0].code || '';
      matched++;
    } else {
      // Ambiguous — multiple buyers share this trade name. Pick the
      // first by code-sort order. Operator resolves per-lot later
      // using the multi-code picker in the Lot edit modal.
      entry.status = 'ambiguous';
      // Prefer a candidate with a non-blank code; fall back to first.
      const withCode = cands.find(c => c.code && String(c.code).trim());
      entry.pickedCode = (withCode || cands[0]).code || '';
      ambiguous++;
    }
    perRow.push(entry);
  }
  const summary = {
    total: perRow.length,
    matched, ambiguous, unmatched, blank,
    uniqueTradeNames: Array.from(new Set(perRow.filter(p => p.tradeName).map(p => p.tradeName))).length,
  };
  return { wb, ws, cols, perRow, summary };
}
app.post('/api/price-list/map-preview', requireView, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const { perRow, summary } = await _plProcessFile(req.file.path);
    res.json({ ...summary, rows: perRow });
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    if (req.file) fs.unlink(req.file.path, () => {});
  }
});
app.post('/api/price-list/map-download', requireView, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const { wb, ws, cols, perRow } = await _plProcessFile(req.file.path);
    // Write the resolved code back into each row's CODE cell. We
    // explicitly set the cell value so the existing column-level numFmt
    // (which Excel uses to right-pad short codes like "RSH") still
    // applies — modifying `.value` keeps the format intact.
    for (const entry of perRow) {
      if (!entry.pickedCode) continue;
      ws.getRow(entry.row).getCell(cols.codeCol).value = entry.pickedCode;
    }
    const buf = await wb.xlsx.writeBuffer();
    const baseName = (req.file.originalname || 'price-list-before.xlsx')
      .replace(/\.xlsx?$/i, '')
      .replace(/[^A-Za-z0-9._-]+/g, '-');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${baseName}-mapped.xlsx"`);
    res.send(Buffer.from(buf));
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    if (req.file) fs.unlink(req.file.path, () => {});
  }
});

// ══════════════════════════════════════════════════════════════
// PRICE CHECK — reconcile uploaded price sheet vs. lots table
// ══════════════════════════════════════════════════════════════
// Replaces the legacy PriceCheck_VSTL.xlsm macro workbook. The
// operator uploads a price sheet (typically the auction-floor record
// or a third-party export) and the server reports:
//   • status per row (match / price-diff / missing / withdrawn / etc.)
//   • code reconciliation (file CODE vs DB buyer code)
//   • totals: matched / mismatched / total |diff| / total signed diff
// Used by the Price Check tab; the operator can apply fixes per-row
// via /api/price-check/apply-fix.

// Reads the uploaded XLSX into an array of plain objects. Tolerant of
// blank rows and header-case differences. The "header row" is the
// first row that contains a "LOT" column (case-insensitive); rows
// above it are treated as metadata and skipped.
function _pcReadRows(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const sn = wb.SheetNames[0];
  if (!sn) throw new Error('Workbook has no sheets');
  const ws = wb.Sheets[sn];
  // Walk the sheet as 2D array first to find the header row, then
  // re-emit as objects keyed by uppercase header. XLSX.utils' raw
  // array form is the most flexible way to handle inconsistent
  // header positions.
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: false });
  if (!grid.length) return { rows: [], headers: [] };
  let headerRow = -1;
  for (let i = 0; i < Math.min(grid.length, 30); i++) {
    const cells = (grid[i] || []).map(c => String(c || '').trim().toUpperCase());
    if (cells.includes('LOT')) { headerRow = i; break; }
  }
  if (headerRow < 0) throw new Error('Could not find a "LOT" column — first 30 rows scanned.');
  const headers = grid[headerRow].map(c => String(c || '').trim().toUpperCase());
  const rows = [];
  for (let i = headerRow + 1; i < grid.length; i++) {
    const cells = grid[i] || [];
    // Skip rows that have NO LOT value (totals/separator lines etc.)
    const lotIdx = headers.indexOf('LOT');
    const lotVal = String((cells[lotIdx] || '')).trim();
    if (!lotVal) continue;
    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      if (headers[c]) obj[headers[c]] = cells[c] == null ? '' : cells[c];
    }
    obj.__row = i + 1;   // 1-based original spreadsheet row number
    rows.push(obj);
  }
  return { rows, headers };
}

// Resolve the file-row's price column. Prefer "SERVER PRICE" (the
// canonical column on the legacy macro's output); fall back to
// "PRICE" so operators uploading raw price files without the macro
// post-process still get a reconciliation.
function _pcExtractPrice(row, hasServerPriceCol) {
  if (hasServerPriceCol) {
    const v = row['SERVER PRICE'];
    return (v === '' || v == null) ? null : Number(v);
  }
  const v = row['PRICE'];
  return (v === '' || v == null) ? null : Number(v);
}

// Build the per-row reconciliation. Shared by /verify and /download
// so the rendered table and the annotated XLSX agree exactly.
function _pcBuildReport(filePath, opts) {
  const { rows, headers } = _pcReadRows(filePath);
  const hasServerPriceCol = headers.includes('SERVER PRICE');
  const hasCodeCol        = headers.includes('CODE');
  const validatedColumn   = hasServerPriceCol ? 'SERVER PRICE' : 'PRICE';

  const db = getDb();
  // Forced-auction mode: caller passed a specific auction_id, so we
  // ignore the file's TNO/ANO columns and reconcile against that one
  // trade. Without it, we resolve per-row via (TNO, DATE).
  let forcedAuction = null;
  if (opts.auctionId) {
    forcedAuction = db.get(
      'SELECT id, ano, date FROM auctions WHERE id = ?',
      [parseInt(opts.auctionId, 10)]
    );
    if (!forcedAuction) throw new Error('Selected auction not found');
  }

  // Helper: resolve an auction by ANO (and optionally DATE) so file
  // rows without an explicit auction_id can be paired with the right
  // trade. Memoised within one report run.
  const auctionByAno = new Map();
  function findAuctionForRow(row) {
    if (forcedAuction) return forcedAuction;
    const ano = String(row['TNO'] || row['ANO'] || '').trim();
    if (!ano) return null;
    const key = ano.toUpperCase();
    if (auctionByAno.has(key)) return auctionByAno.get(key);
    // ANO is a text column in the DB (because lot numbers can be
    // alphanumeric); match case-insensitively.
    const hit = db.get('SELECT id, ano, date FROM auctions WHERE UPPER(ano) = ? ORDER BY date DESC LIMIT 1', [key]);
    auctionByAno.set(key, hit || null);
    return hit || null;
  }

  // Pre-load every lot for the forced auction (when set) so the
  // per-row lookup is a map read instead of an N-queries fan-out.
  let dbLotsByLotNo = null;   // Map<lot_no_upper, lot_row>
  if (forcedAuction) {
    const lots = db.all(
      `SELECT lots.*, (SELECT b.code FROM buyers b WHERE UPPER(b.buyer) = UPPER(lots.buyer) LIMIT 1) AS buyer_code
       FROM lots WHERE auction_id = ?`,
      [forcedAuction.id]
    );
    dbLotsByLotNo = new Map();
    for (const l of lots) {
      const k = String(l.lot_no || '').trim().toUpperCase();
      if (k) dbLotsByLotNo.set(k, l);
    }
  }

  // Walk each file row, classify, and emit a result row.
  const out = [];
  const seenLotIds = new Set();      // lots referenced by the file (for missing_file detection)
  for (const r of rows) {
    const result = {
      row: r.__row,
      lot: String(r['LOT'] || '').trim(),
      auction_ano: forcedAuction ? forcedAuction.ano : String(r['TNO'] || r['ANO'] || '').trim(),
      file_server_price: _pcExtractPrice(r, hasServerPriceCol),
      manual_price: hasServerPriceCol && r['PRICE'] !== '' && r['PRICE'] != null ? Number(r['PRICE']) : null,
      file_code: String(r['CODE'] || '').trim(),
      db_price: null,
      db_code: '',
      diff: null,
      lot_id: null,
      status: 'unmatched',
      code_status: 'both_blank',
      issues: [],
    };

    // Resolve auction first
    const auc = findAuctionForRow(r);
    if (!auc) { result.status = 'no_auction'; result.issues.push('no_auction'); out.push(result); continue; }

    // Resolve the DB lot. If forcedAuction is set, use the pre-loaded
    // map; otherwise hit the DB once per row.
    let dbLot;
    if (forcedAuction && dbLotsByLotNo) {
      dbLot = dbLotsByLotNo.get(result.lot.toUpperCase()) || null;
    } else {
      dbLot = db.get(
        `SELECT lots.*, (SELECT b.code FROM buyers b WHERE UPPER(b.buyer) = UPPER(lots.buyer) LIMIT 1) AS buyer_code
         FROM lots WHERE auction_id = ? AND UPPER(lot_no) = ?`,
        [auc.id, result.lot.toUpperCase()]
      );
    }
    if (!dbLot) { result.status = 'missing_server'; result.issues.push('missing_server'); out.push(result); continue; }

    result.lot_id = dbLot.id;
    result.db_price = Number(dbLot.price) || 0;
    result.db_code  = String(dbLot.buyer_code || dbLot.buyer || '').trim();
    seenLotIds.add(dbLot.id);

    // ── Price comparison ────────────────────────────────────
    const fileP = result.file_server_price;
    if (fileP == null) {
      result.status = 'server_empty';
      result.issues.push('server_empty');
    } else if (Math.abs(fileP - result.db_price) < 0.005) {
      result.status = 'match';
    } else {
      result.status = 'server_diff';
      result.issues.push('server_diff');
      result.diff = +(fileP - result.db_price).toFixed(2);
    }

    // ── Code comparison ─────────────────────────────────────
    // Case-insensitive. 'WD' is the withdrawn marker — we treat
    // it like any other code for the diff check.
    const fc = result.file_code.toUpperCase();
    const dc = result.db_code.toUpperCase();
    if (!fc && !dc)        result.code_status = 'both_blank';
    else if (!fc &&  dc)   result.code_status = 'file_blank';
    else if ( fc && !dc)   result.code_status = 'db_blank';
    else if ( fc === dc)   result.code_status = 'match';
    else                   result.code_status = 'diff';

    out.push(result);
  }

  // Add "missing_file" rows: lots in DB that the file didn't mention.
  // Skipped when no forced auction (we can't enumerate DB lots without
  // an auction_id; the file would need to span multiple auctions).
  if (forcedAuction && dbLotsByLotNo) {
    for (const [lotNoUpper, dbLot] of dbLotsByLotNo) {
      if (seenLotIds.has(dbLot.id)) continue;
      out.push({
        row: null,
        lot: dbLot.lot_no || '',
        auction_ano: forcedAuction.ano,
        file_server_price: null,
        manual_price: null,
        file_code: '',
        db_price: Number(dbLot.price) || 0,
        db_code: String(dbLot.buyer_code || dbLot.buyer || '').trim(),
        diff: null,
        lot_id: dbLot.id,
        status: 'missing_file',
        code_status: dbLot.buyer ? 'file_blank' : 'both_blank',
        issues: ['missing_file'],
      });
    }
  }

  // Aggregate counters.
  const counts = {
    total: out.length,
    matched: 0, mismatched: 0, missingServer: 0, missingFile: 0, noAuction: 0, serverEmpty: 0,
    codeMatched: 0, codeMismatched: 0, codeFileBlank: 0, codeDbBlank: 0,
    withdrawnFile: 0, withdrawnDb: 0,
  };
  let totalAbsDiff = 0, totalSignedDiff = 0;
  for (const r of out) {
    if (r.status === 'match')          counts.matched++;
    if (r.status === 'server_diff')    counts.mismatched++;
    if (r.status === 'missing_server') counts.missingServer++;
    if (r.status === 'missing_file')   counts.missingFile++;
    if (r.status === 'no_auction')     counts.noAuction++;
    if (r.status === 'server_empty')   counts.serverEmpty++;
    if (r.code_status === 'match')      counts.codeMatched++;
    if (r.code_status === 'diff')       counts.codeMismatched++;
    if (r.code_status === 'file_blank') counts.codeFileBlank++;
    if (r.code_status === 'db_blank')   counts.codeDbBlank++;
    if (String(r.file_code).toUpperCase() === 'WD') counts.withdrawnFile++;
    if (String(r.db_code).toUpperCase()   === 'WD') counts.withdrawnDb++;
    if (r.diff != null) {
      totalAbsDiff    += Math.abs(r.diff);
      totalSignedDiff += r.diff;
    }
  }

  return {
    ...counts,
    totalAbsDiff: +totalAbsDiff.toFixed(2),
    totalSignedDiff: +totalSignedDiff.toFixed(2),
    hasServerPriceCol,
    hasCodeCol,
    validatedColumn,
    forcedAuction,
    rows: out,
  };
}

// ── Price-check gate (per-auction tri-state) ──────────────────
// Each auction carries two timestamps:
//   • price_check_first_passed_at — stamped on the FIRST successful
//     verify, never cleared. "This auction has been reconciled at
//     least once."
//   • price_checked_at — stamped on every successful verify AND
//     cleared by any endpoint that mutates lot price/code. "The
//     current reconciliation is still in sync with the lot data."
//
// Derived state:
//   'off'   — feature flag disabled (treat as clean, gate is a no-op)
//   'never' — first-passed-at not set → show banner, hard "pending"
//   'stale' — first-passed-at set but price_checked_at empty → show
//             banner softly ("lots changed since last Price Check")
//   'clean' — both set → no banner
function pcFlagOn(db) {
  try {
    const cfg = getSettingsFlat(db || getDb());
    return String(cfg.flag_price_check || '').toLowerCase() === 'true';
  } catch (_) { return false; }
}
function pcStampGate(db, auctionId) {
  if (!auctionId || !pcFlagOn(db)) return;
  db.run(
    `UPDATE auctions
        SET price_checked_at = datetime('now','localtime'),
            price_check_first_passed_at = COALESCE(NULLIF(price_check_first_passed_at, ''), datetime('now','localtime'))
      WHERE id = ?`,
    [auctionId]
  );
}
function pcClearGate(db, auctionId) {
  if (!auctionId || !pcFlagOn(db)) return;
  // First-pass stamp is permanent; clearing only drops state to 'stale'.
  db.run(`UPDATE auctions SET price_checked_at = '' WHERE id = ?`, [auctionId]);
}
function pcGateState(db, auctionId) {
  if (!auctionId)    return 'never';
  if (!pcFlagOn(db)) return 'off';
  const row = db.get(
    'SELECT price_checked_at, price_check_first_passed_at FROM auctions WHERE id = ?',
    [auctionId]
  );
  if (!row) return 'never';
  if (!row.price_check_first_passed_at) return 'never';
  return row.price_checked_at ? 'clean' : 'stale';
}
// Compute "gate-ready" from a verify summary. We use the same lenient
// rule as the sibling: a verify counts as fully reconciled when no
// fixable code rows remain (CODE Δ + DB BLANK = 0). File-blank codes
// are operator-fixable manually and aren't gated on.
function _pcReportIsGateReady(report) {
  if (!report) return false;
  const codeFixesPending = (report.codeMismatched || 0) + (report.codeDbBlank || 0);
  return codeFixesPending === 0;
}

// POST /api/price-check/verify — returns JSON reconciliation.
app.post('/api/price-check/verify', requireView, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const auctionId = req.body.auction_id ? Number(req.body.auction_id) : null;
    const report = _pcBuildReport(req.file.path, { auctionId });
    // Stamp the gate when the verify ran against a specific auction
    // AND all fixable code rows are resolved. Drives the lots-tab
    // "Lots changed since last Price Check" banner clear-out.
    if (auctionId && _pcReportIsGateReady(report)) {
      pcStampGate(getDb(), auctionId);
    }
    res.json(report);
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    // Best-effort cleanup of the multer temp file.
    try { fs.unlinkSync(req.file.path); } catch (_) {}
  }
});

// Lightweight status probe — drives the lots-tab banner state.
// Tri-state response stays small so the client can poll cheaply.
app.get('/api/auctions/:id/price-check-status', requireView, (req, res) => {
  const id = parseInt(req.params.id, 10);
  if (!id) return res.status(400).json({ error: 'invalid auction id' });
  const db = getDb();
  const row = db.get(
    'SELECT id, ano, date, price_checked_at, price_check_first_passed_at FROM auctions WHERE id = ?', [id]
  );
  if (!row) return res.status(404).json({ error: 'auction not found' });
  const state = pcGateState(db, id);
  res.json({
    auctionId: id,
    ano: row.ano, date: row.date,
    state,                                     // 'off' | 'never' | 'stale' | 'clean'
    checked: state === 'clean' || state === 'off',
    everPassed: !!row.price_check_first_passed_at,
    checkedAt: row.price_checked_at || null,
    firstPassedAt: row.price_check_first_passed_at || null,
  });
});

// POST /api/price-check/download — same data, returns an XLSX with
// extra columns: STATUS, DB PRICE, DIFF, DB CODE, CODE STATUS. The
// original file's structure is preserved as much as possible by
// reading rows + writing a new workbook in the same column order.
app.post('/api/price-check/download', requireView, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const report = _pcBuildReport(req.file.path, { auctionId: req.body.auction_id });
    // Build a new workbook from the report. One row per result row,
    // ordered the same way verify returns them.
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Price Check');
    ws.columns = [
      { header: 'Row',         key: 'row',   width: 6 },
      { header: 'Lot',         key: 'lot',   width: 10 },
      { header: 'Auction No',  key: 'ano',   width: 12 },
      { header: 'File Price',  key: 'fp',    width: 14 },
      { header: 'DB Price',    key: 'dp',    width: 14 },
      { header: 'Diff',        key: 'diff',  width: 12 },
      { header: 'File Code',   key: 'fc',    width: 12 },
      { header: 'DB Code',     key: 'dc',    width: 12 },
      { header: 'Status',      key: 'st',    width: 18 },
      { header: 'Code Status', key: 'cs',    width: 14 },
    ];
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E4DD' } };

    // Per-row tint so a quick scan tells the operator what's wrong.
    const tintByStatus = {
      'match':         null,
      'server_diff':   'FFFEE2E2',
      'server_empty':  'FFFEF3C7',
      'missing_server':'FFFEE2E2',
      'missing_file':  'FFFEF3C7',
      'no_auction':    'FFFEF3C7',
    };
    for (const r of report.rows) {
      const row = ws.addRow({
        row: r.row || '+',
        lot: r.lot,
        ano: r.auction_ano || '',
        fp:  r.file_server_price == null ? '' : r.file_server_price,
        dp:  r.db_price == null ? '' : r.db_price,
        diff: r.diff == null ? '' : r.diff,
        fc:  r.file_code || '',
        dc:  r.db_code || '',
        st:  r.status,
        cs:  r.code_status,
      });
      const fill = tintByStatus[r.status];
      if (fill) {
        for (let c = 1; c <= 10; c++) {
          row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
        }
      }
    }
    const buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="price-check-result.xlsx"');
    res.send(Buffer.from(buf));
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    try { fs.unlinkSync(req.file.path); } catch (_) {}
  }
});

// POST /api/price-check/apply-fix — apply the file's price and/or
// code to a single lot. Body: { lot_id, price?, code? }. At least
// ONE of price/code must be provided. When `code` is supplied we
// look it up in the buyers master so buyer/buyer1/sale stay
// consistent. Lock-aware via the same gate used by other lot writes.
app.post('/api/price-check/apply-fix', requireLotWrite, (req, res) => {
  const lotId = parseInt(req.body.lot_id, 10);
  const price = req.body.price === undefined || req.body.price === null || req.body.price === '' ? null : Number(req.body.price);
  const code  = req.body.code  === undefined || req.body.code  === null ? null : String(req.body.code).trim();
  if (!Number.isFinite(lotId) || lotId <= 0) return res.status(400).json({ error: 'lot_id required' });
  if (price == null && (code == null || code === '')) {
    return res.status(400).json({ error: 'Provide price, code, or both' });
  }
  const db = getDb();
  // Lock gate. Admins bypass.
  const lock = getLotLock(db, lotId);
  if (lock && lock.locked_at && !isAdminUser(req)) {
    return res.status(423).json({
      error: `This lot is locked${lock.locked_by ? ' by ' + lock.locked_by : ''}. Ask an admin to unlock.`,
      locked: true,
    });
  }
  const lot = db.get('SELECT id, qty, buyer FROM lots WHERE id = ?', [lotId]);
  if (!lot) return res.status(404).json({ error: 'Lot not found' });

  // Resolve buyer (if a CODE was provided). We match against either
  // `buyers.buyer` (the primary key) OR `buyers.code` (the short
  // alias) since price files use either form.
  let buyerRow = null;
  if (code) {
    const upper = code.toUpperCase();
    buyerRow = db.get(
      `SELECT buyer, buyer1, sale FROM buyers
       WHERE UPPER(buyer) = ? OR UPPER(code) = ? LIMIT 1`,
      [upper, upper]
    );
    // Note: code could be 'WD' (withdrawn) — no buyer match, but
    // we still write 'WD' into lots.code so the lot is flagged.
    // For other unknown codes, we ALSO let the write proceed (the
    // operator may be using a code that's not in the master yet);
    // it just won't populate buyer1/sale.
  }

  // Build the UPDATE. Always write whatever was supplied; don't touch
  // unrelated columns.
  const sets = [];
  const vals = [];
  if (price != null) {
    sets.push('price = ?'); vals.push(price);
    // Re-derive amount = qty × price so totals stay coherent without
    // a separate Calculate All click. If qty is 0 we leave amount as-is.
    if ((Number(lot.qty) || 0) > 0) {
      sets.push('amount = ?'); vals.push(+(Number(lot.qty) * price).toFixed(2));
    }
  }
  if (code != null) {
    sets.push('code = ?'); vals.push(code);
    if (buyerRow) {
      sets.push('buyer = ?');  vals.push(buyerRow.buyer || '');
      sets.push('buyer1 = ?'); vals.push(buyerRow.buyer1 || '');
      sets.push('sale = ?');   vals.push(buyerRow.sale || 'L');
    }
  }
  vals.push(lotId);
  db.run(`UPDATE lots SET ${sets.join(', ')} WHERE id = ?`, vals);
  res.json({
    success: true,
    applied: { price, code, buyer: buyerRow ? buyerRow.buyer : null },
  });
});

// ── Calculate all lots for an auction (GENERATE.PRG) ─────────
app.post('/api/lots/calculate/:auctionId', requireLotWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const lots = db.all('SELECT * FROM lots WHERE auction_id = ? AND amount > 0', [req.params.auctionId]);
  let count = 0;
  for (const lot of lots) {
    // Re-derive amount = qty × price first so calculateLot (which reads
    // lot.amount as its base) doesn't compound stale data when price has
    // been edited without a matching amount write.
    const qty = Number(lot.qty) || 0, price = Number(lot.price) || 0;
    const newAmount = (qty > 0 && price > 0) ? +(qty * price).toFixed(2) : (Number(lot.amount) || 0);
    lot.amount = newAmount;
    const calc = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET amount=?,pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [newAmount,calc.pqty,calc.prate,calc.puramt,calc.com,calc.sertax,calc.cgst,calc.sgst,calc.igst,calc.advance,calc.balance,calc.bilamt,calc.refund||0,calc.refud||0,calc.isp_pqty||0,calc.isp_prate||0,calc.isp_puramt||0,calc.asp_pqty||0,calc.asp_prate||0,calc.asp_puramt||0,lot.id]);
    count++;
  }
  res.json({ success: true, calculated: count });
});

// Recalculate every lot in every auction with the CURRENT business
// settings. Used by the client when business_state changes — calculations
// like CGST/SGST/IGST and prate are state-sensitive (intra vs inter), so
// the saved values become stale on a state flip and must be refreshed.
//
// CRITICAL: scoped to lots whose parent auction matches the current mode.
// Without this, flipping to e-Auction and clicking "Calculate All" would
// overwrite every e-Trade lot's `prate`/`puramt`/`com`/`sertax` with
// e-Auction formulas (e.g. com=0 in e-Trade becomes a real commission
// number, prate flips from deduction-based to direct copy-of-price).
// NULL/empty mode rows pass the filter so legacy data still recalculates
// during the soft-cutover window.
//
// Only touches lots with `amount > 0` (skips empty/auction-floor entries).
app.post('/api/lots/calculate-all', requireLotWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const mwLots = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=lots.auction_id)');
  const lots = db.all(`SELECT * FROM lots WHERE amount > 0 ${mwLots.sql}`, mwLots.params);
  let count = 0;
  for (const lot of lots) {
    // Re-derive amount = qty × price first so calculateLot (which reads
    // lot.amount as its base) doesn't compound stale data when price has
    // been edited without a matching amount write.
    const qty = Number(lot.qty) || 0, price = Number(lot.price) || 0;
    const newAmount = (qty > 0 && price > 0) ? +(qty * price).toFixed(2) : (Number(lot.amount) || 0);
    lot.amount = newAmount;
    const calc = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET amount=?,pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [newAmount,calc.pqty,calc.prate,calc.puramt,calc.com,calc.sertax,calc.cgst,calc.sgst,calc.igst,calc.advance,calc.balance,calc.bilamt,calc.refund||0,calc.refud||0,calc.isp_pqty||0,calc.isp_prate||0,calc.isp_puramt||0,calc.asp_pqty||0,calc.asp_prate||0,calc.asp_puramt||0,lot.id]);
    count++;
  }
  res.json({ success: true, calculated: count });
});

// ── Data validation (PRICHECK.PRG) ───────────────────────────
app.get('/api/lots/validate/:auctionId', requireView, (req, res) => {
  const rows = getDb().all(
    `SELECT * FROM lots WHERE auction_id = ? AND (price = 0 OR amount = 0 OR buyer = '' OR code = '' OR ROUND(qty*price,2) <> ROUND(amount,2))`,
    [req.params.auctionId]);
  res.json(rows);
});

// ══════════════════════════════════════════════════════════════
// GENERATION LOCK — once EVERY eligible party in a trade has its
// invoice/purchase/bill/debit-note, the matching "Generate" actions
// are blocked. While at least one party is still un-documented, the
// operator can keep generating (single OR bulk) freely.
//
// Admin can grant ONE regeneration at a time by inserting a row
// into generation_overrides; the row is consumed (deleted) the first
// time the corresponding generate endpoint runs.
//
// Doc types: 'invoices' | 'purchases' | 'bills' | 'debit_notes'.
// Payments aren't generated (they're a derived view of lots) so
// they're outside this gate.
// ══════════════════════════════════════════════════════════════
const _GEN_TABLE = {
  invoices:    'invoices',
  purchases:   'purchases',
  bills:       'bills',
  debit_notes: 'debit_notes',
};
function _hasGeneratedDocs(db, docType, auctionId) {
  const table = _GEN_TABLE[docType];
  if (!table || !auctionId) return false;
  if (docType === 'debit_notes') {
    // debit_notes is keyed by trade `ano`, not auction_id.
    const auc = db.get('SELECT ano FROM auctions WHERE id = ?', [auctionId]);
    if (!auc || !auc.ano) return false;
    return !!db.get(`SELECT 1 FROM debit_notes WHERE ano = ? LIMIT 1`, [String(auc.ano)]);
  }
  return !!db.get(`SELECT 1 FROM ${table} WHERE auction_id = ? LIMIT 1`, [auctionId]);
}
// True iff at least one eligible party in the trade is still missing
// its doc. Mirrors each generate-all endpoint's "what's left to do"
// query, so the gate engages exactly when those endpoints would
// return "nothing to do".
function _hasRemainingParties(db, docType, auctionId) {
  if (!auctionId) return false;
  if (docType === 'invoices') {
    const cfg = getSettingsFlat(db);
    const isASPState = String(cfg.business_state || '').toUpperCase() === 'KERALA';
    const uninvoicedExpr = isASPState
      ? `(l.invo IS NULL OR l.invo = '')`
      : `(l.invo IS NULL OR l.invo = '' OR (l.asp_invo IS NOT NULL AND l.asp_invo != '' AND l.invo = l.asp_invo))`;
    // WD = withdrawn lot, no buyer transaction → never gets an invoice.
    // The price-check Apply-fix flow writes code='WD' but leaves the
    // buyer column intact (since 'WD' doesn't match a buyers row), so
    // we have to exclude WD here explicitly or the gate never engages
    // for a trade that has any withdrawn lots.
    return !!db.get(
      `SELECT 1 FROM lots l
       WHERE l.auction_id = ? AND l.buyer IS NOT NULL AND l.buyer != ''
         AND l.amount > 0
         AND UPPER(COALESCE(l.code, '')) != 'WD'
         AND ${uninvoicedExpr}
       LIMIT 1`,
      [auctionId]
    );
  }
  if (docType === 'purchases') {
    // Mirrors /api/purchases/generate-all (UPPER(cr) LIKE 'GSTIN%').
    return !!db.get(
      `SELECT 1 FROM lots l
       WHERE l.auction_id = ? AND l.amount > 0
         AND l.name IS NOT NULL AND l.name != ''
         AND UPPER(l.cr) LIKE 'GSTIN%'
         AND NOT EXISTS (
           SELECT 1 FROM purchases p WHERE p.auction_id = l.auction_id AND p.name = l.name
         )
       LIMIT 1`,
      [auctionId]
    );
  }
  if (docType === 'bills') {
    // Mirrors listAgriSellers in calculations.js — no GSTIN in either
    // prefixed or bare form qualifies as an agri seller.
    return !!db.get(
      `SELECT 1 FROM lots l
       WHERE l.auction_id = ? AND l.amount > 0
         AND l.name IS NOT NULL AND l.name != ''
         AND (l.cr IS NULL OR l.cr = ''
              OR (UPPER(l.cr) NOT LIKE 'GSTIN%' AND l.cr NOT GLOB '[0-9][0-9]*'))
         AND NOT EXISTS (
           SELECT 1 FROM bills b WHERE b.auction_id = l.auction_id AND b.name = l.name
         )
       LIMIT 1`,
      [auctionId]
    );
  }
  if (docType === 'debit_notes') {
    // debit_notes is scoped by trade `ano`. One DN per purchase row
    // (matched by dealer name), so a remaining party = any purchase
    // in this trade without a matching DN.
    const auc = db.get('SELECT ano FROM auctions WHERE id = ?', [auctionId]);
    if (!auc || !auc.ano) return false;
    return !!db.get(
      `SELECT 1 FROM purchases p
       WHERE p.ano = ?
         AND NOT EXISTS (
           SELECT 1 FROM debit_notes d WHERE d.ano = p.ano AND d.name = p.name
         )
       LIMIT 1`,
      [String(auc.ano)]
    );
  }
  return false;
}
// "Fully generated" = at least one doc exists AND no remaining party.
// Empty trades report false (nothing to lock against).
function _isFullyGenerated(db, docType, auctionId) {
  if (!_hasGeneratedDocs(db, docType, auctionId)) return false;
  return !_hasRemainingParties(db, docType, auctionId);
}
function _getGenerationOverride(db, docType, auctionId) {
  if (!auctionId) return null;
  return db.get(
    'SELECT auction_id, doc_type, granted_at, granted_by FROM generation_overrides WHERE auction_id = ? AND doc_type = ?',
    [auctionId, docType]
  );
}
function _consumeGenerationOverride(db, docType, auctionId) {
  if (!auctionId) return;
  db.run('DELETE FROM generation_overrides WHERE auction_id = ? AND doc_type = ?', [auctionId, docType]);
}
// Pre-flight gate check used at the top of every generate endpoint.
// Returns { allowed: true } when (a) at least one party still needs
// a doc, or (b) an admin override is present (consumed here, one-shot).
// Returns { allowed: false, error } with a 412 payload to send.
function _checkGenerationGate(db, docType, auctionId) {
  if (!auctionId) return { allowed: true };
  if (!_isFullyGenerated(db, docType, auctionId)) return { allowed: true };
  const override = _getGenerationOverride(db, docType, auctionId);
  if (override) {
    _consumeGenerationOverride(db, docType, auctionId);
    return { allowed: true, usedOverride: true };
  }
  const labels = {
    invoices: 'Invoices', purchases: 'Purchases',
    bills: 'Bills of Supply', debit_notes: 'Debit Notes',
  };
  const label = labels[docType] || docType;
  return {
    allowed: false,
    error: {
      error: 'Generation locked',
      detail: `${label} have already been generated for every party in this trade. An admin must click 🔓 Allow regeneration before another generate can run.`,
      auctionId, docType, gate: 'generation',
    },
  };
}

// Status endpoint — drives the client's button-disable state.
// `has` is true only when EVERY eligible party has its doc (= the
// gate is engaged). Response shape: { auctionId, invoices:{has,override},
// purchases:{...}, bills:{...}, debit_notes:{...} }.
app.get('/api/auctions/:id/generation-status', requireView, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  if (!auctionId) return res.status(400).json({ error: 'Invalid auction id' });
  const out = {};
  for (const docType of Object.keys(_GEN_TABLE)) {
    out[docType] = {
      has: _isFullyGenerated(db, docType, auctionId),
      override: !!_getGenerationOverride(db, docType, auctionId),
    };
  }
  res.json({ auctionId, ...out });
});

// Admin: grant one regeneration for (auction, doc_type). Idempotent —
// re-granting before consumption is a no-op (PRIMARY KEY collision
// is resolved by INSERT OR REPLACE).
app.post('/api/auctions/:id/generation-override', requireAdmin, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const docType = String(req.body.docType || '').trim();
  if (!auctionId) return res.status(400).json({ error: 'Invalid auction id' });
  if (!_GEN_TABLE[docType]) return res.status(400).json({ error: `Invalid docType: ${docType}` });
  const grantedBy = (req.user && (req.user.username || req.user.name)) || 'admin';
  db.run(
    `INSERT OR REPLACE INTO generation_overrides
       (auction_id, doc_type, granted_at, granted_by)
       VALUES (?, ?, datetime('now','localtime'), ?)`,
    [auctionId, docType, grantedBy]
  );
  res.json({ ok: true, auctionId, docType, grantedBy });
});

// ══════════════════════════════════════════════════════════════
// INVOICES — Sales (GSTIN.PRG / KGSTIN.PRG)
// ══════════════════════════════════════════════════════════════
app.get('/api/invoices', requireView, (req, res) => {
  const { ano, auction_id, from, to, search, saleType } = req.query;
  const db = getDb();
  const cfg = getSettingsFlat(db);
  // Filter list by active business context: when state=KERALA show only
  // ASP-stamped invoices, when state=TAMIL NADU show only ISP-stamped.
  // This avoids the "two rows per buyer" confusion in users who run the
  // ASP→ISP flow on the same auction.
  const businessState = String(cfg.business_state || 'TAMIL NADU').toUpperCase();
  let q = 'SELECT * FROM invoices WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  // Sale-type filter (L/I/E) — set by the toolbar dropdown. Empty means
  // "All sale types" so we skip the predicate entirely.
  if (saleType) {
    q += ' AND UPPER(sale) = ?';
    p.push(String(saleType).toUpperCase());
  }
  // Free-text search across invoice no, buyer name (both display + short),
  // GSTIN, place, and lorry no. Active filter bypasses the auction_id
  // narrowing above (so cross-trade search works) but still honours the
  // business-state / mode filters below.
  if (search) {
    const s = `%${String(search).trim()}%`;
    q += ` AND (invo LIKE ? OR buyer LIKE ? OR buyer1 LIKE ? OR gstin LIKE ? OR place LIKE ? OR lorry_no LIKE ?)`;
    p.push(s, s, s, s, s, s);
  }
  // Match the invoice's stamped state to the current business context.
  // We allow both spellings (TAMIL NADU / Tamil Nadu / TN) just in case.
  if (businessState === 'KERALA') {
    q += " AND UPPER(state) = 'KERALA'";
  } else {
    q += " AND UPPER(state) IN ('TAMIL NADU', 'TAMILNADU', 'TN')";
  }
  // Mode filter via parent auction. NULL/empty mode rows pass (legacy data).
  const mw = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=invoices.auction_id)');
  q += mw.sql; p.push(...mw.params);
  q += ' ORDER BY date DESC, invo DESC LIMIT 500';
  const rows = db.all(q, p);
  // Hydrate asp_invo: for each invoice, find the ASP invoice number
  // recorded on its lots. Multiple distinct asp_invos are concatenated
  // (rare — usually one ASP invoice maps 1:1 to one ISP invoice for the
  // same buyer/auction). Empty for ASP invoices themselves.
  const aspStmt = db.prepare(
    `SELECT DISTINCT asp_invo FROM lots
     WHERE auction_id = ? AND buyer = ? AND invo = ?
       AND asp_invo IS NOT NULL AND asp_invo != ''`
  );
  for (const r of rows) {
    // For ASP invoices (state contains "Kerala"), the asp_invo column
    // would just be a copy of `invo` — show blank instead of duplicating.
    const isASPRow = String(r.state || '').toLowerCase().includes('kerala');
    if (isASPRow) { r.asp_invo = ''; continue; }
    const aspRows = aspStmt.all(r.auction_id, r.buyer, r.invo);
    r.asp_invo = aspRows.map(x => x.asp_invo).filter(Boolean).join(', ');
  }
  res.json(rows);
});

// Remove any prior sales-invoice rows this generation would duplicate, so a
// (re)generation REPLACES instead of APPENDS. Without this, granting "Allow
// regeneration" and generating again leaves the old rows in place — the
// root cause of the "two rows per buyer" duplicates. Context-aware:
//   • Kerala / ASP   → ASP bills exactly one inter-state invoice per buyer
//                      per auction, so drop every prior KERALA row for the
//                      buyer (any sale type — also clears stale pre-change
//                      CGST/SGST rows).
//   • Tamil Nadu/ISP → preserve a buyer's distinct L / I / E invoices; drop
//                      only the matching (buyer, sale type) row.
function clearPriorSalesInvoice(db, auctionId, buyer, invoiceState, saleType) {
  const isKerala = String(invoiceState || '').toUpperCase() === 'KERALA';
  if (isKerala) {
    db.run("DELETE FROM invoices WHERE auction_id=? AND buyer=? AND UPPER(state)='KERALA'",
      [auctionId, buyer]);
  } else {
    db.run("DELETE FROM invoices WHERE auction_id=? AND buyer=? AND UPPER(COALESCE(sale,''))=? AND UPPER(state) IN ('TAMIL NADU','TAMILNADU','TN')",
      [auctionId, buyer, String(saleType || '').toUpperCase()]);
  }
}

app.post('/api/invoices/generate/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const auctionIdForGate = parseInt(req.params.auctionId, 10);
  const _gen = _checkGenerationGate(db, 'invoices', auctionIdForGate);
  if (!_gen.allowed) return res.status(412).json(_gen.error);
  const { saleType, buyerCode, invoiceNo } = req.body;
  
  if (!saleType || !buyerCode || !invoiceNo) {
    return res.status(400).json({ error: 'saleType, buyerCode, and invoiceNo are required' });
  }
  
  // Auto-calculate lots if puramt is missing (user might not have clicked Calculate)
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  
  const invoice = buildSalesInvoice(db, req.params.auctionId, buyerCode, saleType, cfg);
  if (!invoice) return res.status(404).json({ error: `No lots found for buyer "${buyerCode}" in this auction. Make sure lots have this buyer code assigned.` });
  
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const s = invoice.summary;
  // Store the BUSINESS context state (TAMIL NADU=ISP, KERALA=ASP), not
  // the auction's physical state. This lets us distinguish ASP invoices
  // from ISP invoices in the same auction, which matters for the sales
  // list cross-reference (ASP Inv# column).
  const invoiceState = cfg.business_state || auction.state || '';
  // Idempotent (re)generation — clear any prior matching row(s) first.
  clearPriorSalesInvoice(db, req.params.auctionId, buyerCode, invoiceState, invoice.saleType);
  db.run(`INSERT INTO invoices (auction_id,ano,date,state,sale,invo,buyer,buyer1,gstin,place,bag,qty,amount,gunny,pava_hc,ins,cgst,sgst,igst,tcs,rund,tot)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [req.params.auctionId,auction.ano,auction.date,invoiceState,invoice.saleType,String(invoiceNo),buyerCode,invoice.buyer.buyer1||'',
     invoice.buyer.gstin||'',invoice.buyer.pla||'',s.totalBags,s.totalQty,s.totalAmount,s.gunnyCost,s.transportCost,s.insuranceCost,
     s.cgst,s.sgst,s.igst,0,s.roundDiff,s.grandTotal]);
  
  // Update lots with sale type and invoice number.
  // Workflow trace:
  //   - In Kerala (ASP) context: set `invo` AND `asp_invo` to the new ASP
  //     invoice number. DON'T update `lots.sale` — sale type is determined
  //     by the ISP→external transaction (could be Local/Inter-state/Export
  //     depending on buyer's GST state). ASP→ISP is a fixed intra-Kerala
  //     transfer, so any `sale` value would constrain the later ISP step.
  //     EXCEPT: if `invo` already holds a non-ASP value (i.e., an ISP
  //     invoice was already generated), preserve `invo` and only refresh
  //     `asp_invo`. This prevents accidentally destroying the ISP invoice
  //     number when re-running ASP after the full ASP→ISP cycle.
  //   - In Tamil Nadu (ISP) context: update `sale` and `invo` — `asp_invo`
  //     retains the prior ASP number from the earlier ASP-state generation.
  const isASPState = String(cfg.business_state || '').toUpperCase() === 'KERALA';
  for (const li of invoice.lineItems) {
    if (isASPState) {
      const existing = db.get(
        'SELECT invo, asp_invo FROM lots WHERE auction_id=? AND lot_no=? AND buyer=? LIMIT 1',
        [req.params.auctionId, li.lot, buyerCode]
      );
      const hasIspInvo = existing && existing.invo && existing.invo !== existing.asp_invo;
      if (hasIspInvo) {
        // Preserve invo (ISP); refresh only asp_invo. Sale stays as set by ISP step.
        db.run('UPDATE lots SET asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
          [String(invoiceNo), req.params.auctionId, li.lot, buyerCode]);
      } else {
        // First-time ASP: don't touch `sale` so ISP step has a clean slate
        db.run('UPDATE lots SET invo=?, asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
          [String(invoiceNo), String(invoiceNo), req.params.auctionId, li.lot, buyerCode]);
      }
    } else {
      db.run('UPDATE lots SET sale=?, invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
        [saleType, String(invoiceNo), req.params.auctionId, li.lot, buyerCode]);
    }
  }
  res.json({ success: true, invoice: invoice.summary });
});

// List eligible buyers for an auction (distinct buyers with lots in amount > 0)
app.get('/api/invoices/eligible-buyers/:auctionId', requireView, (req, res) => {
  const { saleType } = req.query;
  const db = getDb();
  const cfg = getSettingsFlat(db);
  const params = [req.params.auctionId];

  // Match buyers by sale type via their default (b.sale) when a type is specified.
  // A buyer is eligible when any of their lots in this auction isn't yet invoiced
  // for the current state context (so user can always see/regenerate; server
  // endpoint has stricter filter).
  let saleClause = '';
  if (saleType) {
    saleClause = ` AND (COALESCE(NULLIF(l.sale,''), b.sale, 'L') = ?)`;
    params.push(saleType);
  }

  // State-aware eligibility:
  //   - In Kerala (ASP) context: a lot is eligible if no `invo` set yet,
  //     OR if `invo == asp_invo` (i.e., it was previously invoiced in
  //     ASP — user is regenerating).
  //   - In Tamil Nadu (ISP) context: a lot is eligible if `invo` is empty
  //     OR if `invo == asp_invo` (lot only has its ASP invoice, still
  //     needs ISP invoicing). This is the key case that was broken.
  // In both states, lots with `invo != asp_invo AND invo != ''` are
  // considered "fully invoiced" for the current state and excluded.
  const isASPState = String(cfg.business_state || '').toUpperCase() === 'KERALA';
  // Both states share the same eligibility expression — what differs is
  // the meaning. The expression: lot is eligible if no `invo` OR `invo
  // matches asp_invo` (meaning the only existing invoice on this lot is
  // an ASP one, which doesn't count toward "ISP-invoiced" status).
  const eligibleExpr = isASPState
    ? `(l.invo IS NULL OR l.invo = '')`
    : `(l.invo IS NULL OR l.invo = '' OR (l.asp_invo IS NOT NULL AND l.asp_invo != '' AND l.invo = l.asp_invo))`;

  res.json(db.all(
    `SELECT l.buyer, COALESCE(b.buyer1, MAX(l.buyer1), l.buyer) as buyer1,
        b.code as code,
        COUNT(*) as lot_count, SUM(l.qty) as total_qty, SUM(l.amount) as total_amount,
        b.gstin, b.sale as default_sale
     FROM lots l
     LEFT JOIN buyers b ON b.buyer = l.buyer
     WHERE l.auction_id = ?
       AND l.buyer IS NOT NULL AND l.buyer != ''
       AND UPPER(COALESCE(l.code, '')) != 'WD'
       ${saleClause}
     GROUP BY l.buyer
     HAVING COUNT(CASE WHEN ${eligibleExpr} THEN 1 END) > 0
     ORDER BY l.buyer`,
    params
  ));
});

// ── Diagnostic: show EVERYTHING about buyers in an auction ──
// Helps troubleshoot why eligible-buyers returns an unexpected count.
app.get('/api/invoices/eligibility-debug/:auctionId', requireView, (req, res) => {
  const db = getDb();
  const aid = req.params.auctionId;
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [aid]);
  if (!auction) return res.status(404).json({ error: 'Auction not found' });

  // Every distinct value in lots.buyer (including blanks), with counts
  const allBuyerGroups = db.all(
    `SELECT
       COALESCE(NULLIF(TRIM(l.buyer),''), '<BLANK>') as buyer_raw,
       COUNT(*) as total_lots,
       COUNT(CASE WHEN l.invo IS NULL OR l.invo = '' THEN 1 END) as uninvoiced_lots,
       COUNT(CASE WHEN l.amount > 0 THEN 1 END) as priced_lots,
       SUM(l.amount) as total_amount,
       MAX(l.buyer1) as lot_buyer1,
       (SELECT buyer1 FROM buyers WHERE buyer = l.buyer LIMIT 1) as master_buyer1,
       (SELECT sale    FROM buyers WHERE buyer = l.buyer LIMIT 1) as master_sale,
       (SELECT gstin   FROM buyers WHERE buyer = l.buyer LIMIT 1) as master_gstin,
       (SELECT id      FROM buyers WHERE buyer = l.buyer LIMIT 1) as master_id
     FROM lots l
     WHERE l.auction_id = ?
     GROUP BY TRIM(l.buyer)
     ORDER BY total_lots DESC`,
    [aid]
  );

  const total_lots      = db.get('SELECT COUNT(*) as c FROM lots WHERE auction_id = ?', [aid]).c;
  const lots_no_buyer   = db.get(`SELECT COUNT(*) as c FROM lots WHERE auction_id = ? AND (buyer IS NULL OR TRIM(buyer) = '')`, [aid]).c;
  const lots_invoiced   = db.get(`SELECT COUNT(*) as c FROM lots WHERE auction_id = ? AND invo IS NOT NULL AND invo != ''`, [aid]).c;
  const distinct_buyers_in_lots = db.get(
    `SELECT COUNT(*) as c FROM (
       SELECT DISTINCT TRIM(buyer) as b FROM lots
       WHERE auction_id = ? AND buyer IS NOT NULL AND TRIM(buyer) != ''
     )`, [aid]).c;

  res.json({
    auction: { id: auction.id, ano: auction.ano, date: auction.date, crop_type: auction.crop_type },
    totals: {
      total_lots,
      lots_with_blank_buyer: lots_no_buyer,
      lots_already_invoiced: lots_invoiced,
      distinct_buyer_codes_in_lots: distinct_buyers_in_lots,
      buyers_table_total: db.get('SELECT COUNT(*) as c FROM buyers').c,
    },
    breakdown: allBuyerGroups.map(r => ({
      buyer_code: r.buyer_raw,
      master_match: r.master_id ? 'yes' : 'NO — not in buyers table',
      master_buyer1: r.master_buyer1 || null,
      lot_buyer1:    r.lot_buyer1 || null,
      master_sale:   r.master_sale || null,
      total_lots:      r.total_lots,
      uninvoiced_lots: r.uninvoiced_lots,
      priced_lots:     r.priced_lots,
      total_amount:    r.total_amount,
      eligible: r.buyer_raw !== '<BLANK>' && r.uninvoiced_lots > 0 ? 'yes' : 'NO',
      eligibility_reason: r.buyer_raw === '<BLANK>' ? 'buyer code is blank'
        : r.uninvoiced_lots === 0 ? 'all lots already invoiced'
        : 'eligible'
    }))
  });
});

// Batch: generate sales invoice for ALL buyers in an auction
app.post('/api/invoices/generate-all/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const auctionIdForGate = parseInt(req.params.auctionId, 10);
  const _gen = _checkGenerationGate(db, 'invoices', auctionIdForGate);
  if (!_gen.allowed) return res.status(412).json(_gen.error);
  const { startInvoiceNo, saleType } = req.body;
  
  let nextNo = parseInt(startInvoiceNo);
  if (!nextNo || nextNo < 1) return res.status(400).json({ error: 'startInvoiceNo must be a positive integer' });
  
  // Auto-calculate uncalculated lots
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  
  // Get distinct buyers. When saleType filter is set, only buyers whose
  // default sale matches (or whose lots already have that sale assigned) are included.
  // The "un-invoiced" check is state-aware:
  //   - In Tamil Nadu (ISP): a lot is un-invoiced if `invo` is empty OR
  //     if the only existing invoice is the ASP one (invo == asp_invo).
  //   - In Kerala (ASP): un-invoiced means `invo` is empty.
  const isASPState = String(cfg.business_state || '').toUpperCase() === 'KERALA';
  const uninvoicedExpr = isASPState
    ? `(l.invo IS NULL OR l.invo = '')`
    : `(l.invo IS NULL OR l.invo = '' OR (l.asp_invo IS NOT NULL AND l.asp_invo != '' AND l.invo = l.asp_invo))`;
  const params = [req.params.auctionId];
  let saleClause = '';
  if (saleType) {
    saleClause = ` AND (COALESCE(NULLIF(l.sale,''), b.sale, 'L') = ?)`;
    params.push(saleType);
  }
  const buyers = db.all(
    `SELECT DISTINCT l.buyer, b.sale as default_sale
     FROM lots l LEFT JOIN buyers b ON b.buyer = l.buyer
     WHERE l.auction_id = ? AND l.buyer IS NOT NULL AND l.buyer != '' AND l.amount > 0
       AND ${uninvoicedExpr}
       ${saleClause}`,
    params
  );
  
  if (!buyers.length) return res.status(404).json({ error: saleType ? `No un-invoiced buyers for sale type ${saleType}` : 'No un-invoiced buyers with lots in this auction' });
  
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const results = [];
  const errors = [];
  
  for (const row of buyers) {
    const useSaleType = saleType || row.default_sale || 'L';
    try {
      const invoice = buildSalesInvoice(db, req.params.auctionId, row.buyer, useSaleType, cfg);
      if (!invoice) { errors.push({ buyer: row.buyer, error: 'No matching lots' }); continue; }
      const s = invoice.summary;
      const invoNo = String(nextNo);
      // Store BUSINESS context state — see single-invoice handler for rationale
      const invoiceState = cfg.business_state || auction.state || '';
      // Idempotent (re)generation — clear any prior matching row(s) first.
      clearPriorSalesInvoice(db, req.params.auctionId, row.buyer, invoiceState, invoice.saleType);
      db.run(`INSERT INTO invoices (auction_id,ano,date,state,sale,invo,buyer,buyer1,gstin,place,bag,qty,amount,gunny,pava_hc,ins,cgst,sgst,igst,tcs,rund,tot)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [req.params.auctionId,auction.ano,auction.date,invoiceState,invoice.saleType,invoNo,row.buyer,invoice.buyer.buyer1||'',
         invoice.buyer.gstin||'',invoice.buyer.pla||'',s.totalBags,s.totalQty,s.totalAmount,s.gunnyCost,s.transportCost,s.insuranceCost,
         s.cgst,s.sgst,s.igst,0,s.roundDiff,s.grandTotal]);
      // ASP-aware lot update: see single-invoice handler above for rationale.
      const isASPStateBulk = String(cfg.business_state || '').toUpperCase() === 'KERALA';
      for (const li of invoice.lineItems) {
        if (isASPStateBulk) {
          const existing = db.get(
            'SELECT invo, asp_invo FROM lots WHERE auction_id=? AND lot_no=? AND buyer=? LIMIT 1',
            [req.params.auctionId, li.lot, row.buyer]
          );
          const hasIspInvo = existing && existing.invo && existing.invo !== existing.asp_invo;
          if (hasIspInvo) {
            db.run('UPDATE lots SET asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
              [invoNo, req.params.auctionId, li.lot, row.buyer]);
          } else {
            // Don't set `sale` in ASP context — ISP step decides
            db.run('UPDATE lots SET invo=?, asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
              [invoNo, invoNo, req.params.auctionId, li.lot, row.buyer]);
          }
        } else {
          db.run('UPDATE lots SET sale=?, invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
            [useSaleType, invoNo, req.params.auctionId, li.lot, row.buyer]);
        }
      }
      results.push({ buyer: row.buyer, invoiceNo: invoNo, sale: invoice.saleType, grandTotal: s.grandTotal });
      nextNo++;
    } catch (e) { errors.push({ buyer: row.buyer, error: e.message }); }
  }
  
  res.json({ success: true, generated: results.length, results, errors });
});

// ── LORRY / VEHICLE NUMBER — bulk-set on selected invoices ──
// Stored on invoices.lorry_no (already in db.js). The Tally sales voucher
// generator reads this column and emits it as the e-way bill VehicleNo.
// Body: { ids: [1,2,3], lorry_no: 'TN66H1234' }  (lorry_no '' = clear)
//
// CRITICAL: declared BEFORE `app.put('/api/invoices/:id')` because Express
// matches in declaration order — a generic `:id` route declared first
// would capture 'lorry-no' as the id and 404 the request.
app.put('/api/invoices/lorry-no', requireInvoiceWrite, (req, res) => {
  const { ids, lorry_no } = req.body || {};
  if (!Array.isArray(ids) || !ids.length) {
    return res.status(400).json({ error: 'ids array required' });
  }
  const cleanIds = ids.map(Number).filter(Number.isFinite);
  if (!cleanIds.length) return res.status(400).json({ error: 'No valid invoice IDs' });
  // Normalise: trim, uppercase, strip spaces. Empty → NULL (clear). We
  // deliberately DON'T validate format — Indian plates have many regional
  // variants and rejecting valid ones is worse than accepting a typo.
  let v = null;
  if (lorry_no != null && String(lorry_no).trim() !== '') {
    v = String(lorry_no).trim().toUpperCase().replace(/\s+/g, '');
    if (v.length > 20) return res.status(400).json({ error: 'Lorry no too long (max 20 chars)' });
  }
  try {
    const db = getDb();
    const placeholders = cleanIds.map(() => '?').join(',');
    const r = db.run(
      `UPDATE invoices SET lorry_no = ? WHERE id IN (${placeholders})`,
      [v, ...cleanIds]
    );
    res.json({ ok: true, updated: r.changes, lorry_no: v });
  } catch (e) {
    console.error('[lorry-no] bulk update failed:', e);
    res.status(500).json({ error: e.message });
  }
});

// ── BULK DELETE — selected invoices ──
// Body: { ids: [1,2,3] }. Frees the lots tied to each deleted invoice so
// they can be re-invoiced. Single-round-trip replacement for the
// per-row Delete All flow which nuked every invoice in the trade.
app.post('/api/invoices/bulk-delete', requireDelete, (req, res) => {
  const { ids } = req.body || {};
  if (!Array.isArray(ids) || !ids.length) {
    return res.status(400).json({ error: 'ids array required' });
  }
  const cleanIds = ids.map(Number).filter(Number.isFinite);
  if (!cleanIds.length) return res.status(400).json({ error: 'No valid invoice IDs' });
  const db = getDb();
  let deleted = 0;
  let lotsFreed = 0;
  for (const id of cleanIds) {
    const inv = db.get('SELECT * FROM invoices WHERE id=?', [id]);
    if (!inv) continue;
    // Clear sale/invo from related lots so they become eligible again.
    if (inv.auction_id) {
      const before = db.get(
        'SELECT COUNT(*) as c FROM lots WHERE auction_id=? AND sale=? AND invo=? AND buyer=?',
        [inv.auction_id, inv.sale, inv.invo, inv.buyer]
      ).c;
      db.run(
        `UPDATE lots SET sale='', invo='' WHERE auction_id=? AND sale=? AND invo=? AND buyer=?`,
        [inv.auction_id, inv.sale, inv.invo, inv.buyer]
      );
      lotsFreed += before;
    }
    db.run('DELETE FROM invoices WHERE id=?', [id]);
    deleted++;
  }
  res.json({ ok: true, deleted, lotsFreed });
});

// Update invoice fields (edit)
app.put('/api/invoices/:id', requireInvoiceWrite, (req, res) => {
  const i = req.body;
  const fields = ['ano','date','state','sale','invo','buyer','buyer1','gstin','place',
    'bag','qty','amount','gunny','pava_hc','ins','cgst','sgst','igst','tcs','rund','tot'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (i[f] !== undefined) { sets.push(`${f}=?`); vals.push(i[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  getDb().run(`UPDATE invoices SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

// Delete invoice
app.delete('/api/invoices/:id', requireDelete, (req, res) => {
  const db = getDb();
  const inv = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
  if (!inv) return res.status(404).json({ error: 'Invoice not found' });
  // Clear sale/invo from the related lots so they're eligible again
  let lotsFreed = 0;
  if (inv.auction_id) {
    const before = db.get('SELECT COUNT(*) as c FROM lots WHERE auction_id=? AND sale=? AND invo=? AND buyer=?',
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]).c;
    db.run(`UPDATE lots SET sale='', invo='' WHERE auction_id=? AND sale=? AND invo=? AND buyer=?`,
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]);
    lotsFreed = before;
  }
  db.run('DELETE FROM invoices WHERE id=?', [req.params.id]);
  res.json({ success: true, invoiceId: Number(req.params.id), lotsFreed });
});

// Explicit revert route (same effect as DELETE but returns richer info)
app.post('/api/invoices/:id/revert', requireInvoiceRevert, (req, res) => {
  const db = getDb();
  const inv = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
  if (!inv) return res.status(404).json({ error: 'Invoice not found' });
  let lotsFreed = 0;
  if (inv.auction_id) {
    const affected = db.all(
      'SELECT lot_no FROM lots WHERE auction_id=? AND sale=? AND invo=? AND buyer=?',
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]
    );
    lotsFreed = affected.length;
    db.run(`UPDATE lots SET sale='', invo='' WHERE auction_id=? AND sale=? AND invo=? AND buyer=?`,
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]);
    db.run('DELETE FROM invoices WHERE id=?', [req.params.id]);
    return res.json({
      success: true,
      invoice: { sale: inv.sale, invo: inv.invo, buyer: inv.buyer, buyer1: inv.buyer1 },
      lotsFreed,
      lots: affected.map(r => r.lot_no),
    });
  }
  db.run('DELETE FROM invoices WHERE id=?', [req.params.id]);
  res.json({ success: true, lotsFreed: 0 });
});

// Bulk revert: revert ALL invoices in an auction
app.post('/api/invoices/revert-all/:auctionId', requireInvoiceRevert, (req, res) => {
  const db = getDb();
  const aid = req.params.auctionId;
  const invoices = db.all('SELECT * FROM invoices WHERE auction_id = ?', [aid]);
  let lotsFreed = 0;
  for (const inv of invoices) {
    const n = db.get('SELECT COUNT(*) as c FROM lots WHERE auction_id=? AND sale=? AND invo=? AND buyer=?',
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]).c;
    lotsFreed += n;
    db.run(`UPDATE lots SET sale='', invo='' WHERE auction_id=? AND sale=? AND invo=? AND buyer=?`,
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]);
  }
  db.run('DELETE FROM invoices WHERE auction_id = ?', [aid]);
  // Safety net: clear any orphan invo values from lots in this auction
  const orphan = db.get(
    `SELECT COUNT(*) as c FROM lots WHERE auction_id = ? AND invo IS NOT NULL AND invo != ''`, [aid]
  ).c;
  if (orphan) {
    db.run(`UPDATE lots SET sale='', invo='' WHERE auction_id = ?`, [aid]);
    lotsFreed += orphan;
  }
  res.json({ success: true, invoicesReverted: invoices.length, lotsFreed });
});

// Sales Invoice PDF
app.get('/api/invoices/pdf/:id', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const stored = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
    if (!stored) return res.status(404).json({ error: 'Invoice not found' });

    // Try to rebuild fresh from lots (gives line-item detail), fall back to stored summary
    let invoice = stored.auction_id
      ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg, { aspInvoice: String(stored.state || '').toUpperCase() === 'KERALA' })
      : null;

    // Defensive: even when lots exist, if buyer lookup missed, enrich from stored invoice fields
    // so BILL TO / SHIPPED TO isn't blank.
    const enrichBuyer = (buyer) => {
      if (!buyer) buyer = {};
      // If buyer has no recognizable display field, try several fallbacks
      if (!buyer.buyer1 && !buyer.buyer) {
        const looked = db.get('SELECT * FROM buyers WHERE buyer=? OR buyer1=? LIMIT 1',
          [stored.buyer, stored.buyer1 || stored.tradername || '']);
        if (looked) buyer = looked;
      }
      // Last-resort fill from stored invoice row
      if (!buyer.buyer1 && stored.buyer1)   buyer.buyer1   = stored.buyer1;
      if (!buyer.buyer1 && stored.tradername) buyer.buyer1 = stored.tradername;
      if (!buyer.buyer  && stored.buyer)    buyer.buyer    = stored.buyer;
      if (!buyer.gstin  && stored.gstin)    buyer.gstin    = stored.gstin;
      if (!buyer.pla    && stored.place)    buyer.pla      = stored.place;
      if (!buyer.state  && stored.state)    buyer.state    = stored.state;
      if (!buyer.add1   && stored.add_line) buyer.add1     = stored.add_line;
      return buyer;
    };

    if (invoice) {
      invoice.buyer = enrichBuyer(invoice.buyer);
    } else {
      // Build a minimal invoice object from stored fields (lots may have been deleted)
      const buyer = enrichBuyer(db.get('SELECT * FROM buyers WHERE buyer=? LIMIT 1', [stored.buyer]));
      invoice = {
        buyer,
        lineItems: [{ lot: '—', grade: '', bags: stored.bag || 0, qty: stored.qty || 0, price: 0, amount: stored.amount || 0 }],
        summary: {
          totalBags: stored.bag || 0,
          totalQty: stored.qty || 0,
          totalAmount: stored.amount || 0,
          gunnyCost: stored.gunny || 0,
          transportCost: stored.pava_hc || 0,
          insuranceCost: stored.ins || 0,
          cgst: stored.cgst || 0,
          sgst: stored.sgst || 0,
          igst: stored.igst || 0,
          tcs: stored.tcs || 0,
          roundDiff: stored.rund || 0,
          grandTotal: stored.tot || 0,
          isInterState: stored.sale === 'I',
        }
      };
    }

    // Optional dispatched-through override from print modal (URL-encoded)
    const dispatchedThrough = req.query.dispatchedThrough || '';
    if (dispatchedThrough) invoice.dispatchedThrough = dispatchedThrough;

    // Look up the ASP invoice number from lots so the ISP PDF can show
    // the cross-reference under "Other References" as ASP/I-{asp}/{season}.
    // When this invoice IS an ASP one (state=KERALA), aspInvo stays empty.
    if (String(stored.state || '').toUpperCase() !== 'KERALA') {
      const aspRow = db.get(
        `SELECT asp_invo FROM lots
         WHERE auction_id = ? AND buyer = ? AND invo = ?
           AND asp_invo IS NOT NULL AND asp_invo != ''
         LIMIT 1`,
        [stored.auction_id, stored.buyer, stored.invo]
      );
      if (aspRow && aspRow.asp_invo) invoice.aspInvo = aspRow.asp_invo;
    }

    const pdf = await generateSalesInvoicePDF(invoice, cfg, stored.sale, stored.invo, stored.date);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="Invoice_${stored.sale}_${stored.invo}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Sales invoice PDF error:', e);
    res.status(500).json({ error: 'PDF generation failed: ' + e.message });
  }
});

// ASP Purchase-view PDF — renders the mirror-image of an ASP sales invoice
// with ISPL at the top (issuer), ASP as Seller (Bill from), TN bank details.
// Only valid when current business_mode is e-Trade and business_state is KERALA;
// otherwise returns 400. Uses the same `generateSalesInvoicePDF` code path with
// variant='purchase' so the math (P_Rate, PurAmt, totals, HSN) stays identical
// to the source sales invoice.
app.get('/api/invoices/purchase-pdf/:id', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    // Guard: purchase view is meaningful only for ASP invoices
    const isASPContext = (String(cfg.business_mode || '').toLowerCase() === 'e-trade')
                      && (String(cfg.business_state || '').toUpperCase() === 'KERALA');
    if (!isASPContext) {
      return res.status(400).json({
        error: 'Purchase view is only available for ASP invoices. Switch business state to KERALA + e-Trade mode.'
      });
    }
    const stored = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
    if (!stored) return res.status(404).json({ error: 'Invoice not found' });

    // Same enrichment pattern as the sales-invoice endpoint
    let invoice = stored.auction_id
      ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg, { aspInvoice: String(stored.state || '').toUpperCase() === 'KERALA' })
      : null;

    const enrichBuyer = (buyer) => {
      if (!buyer) buyer = {};
      if (!buyer.buyer1 && !buyer.buyer) {
        const looked = db.get('SELECT * FROM buyers WHERE buyer=? OR buyer1=? LIMIT 1',
          [stored.buyer, stored.buyer1 || stored.tradername || '']);
        if (looked) buyer = looked;
      }
      if (!buyer.buyer1 && stored.buyer1)   buyer.buyer1   = stored.buyer1;
      if (!buyer.buyer1 && stored.tradername) buyer.buyer1 = stored.tradername;
      if (!buyer.buyer  && stored.buyer)    buyer.buyer    = stored.buyer;
      if (!buyer.gstin  && stored.gstin)    buyer.gstin    = stored.gstin;
      if (!buyer.pla    && stored.place)    buyer.pla      = stored.place;
      if (!buyer.state  && stored.state)    buyer.state    = stored.state;
      if (!buyer.add1   && stored.add_line) buyer.add1     = stored.add_line;
      return buyer;
    };

    if (invoice) {
      invoice.buyer = enrichBuyer(invoice.buyer);
    } else {
      const buyer = enrichBuyer(db.get('SELECT * FROM buyers WHERE buyer=? LIMIT 1', [stored.buyer]));
      invoice = {
        buyer,
        lineItems: [{ lot: '—', grade: '', bags: stored.bag || 0, qty: stored.qty || 0, price: 0, amount: stored.amount || 0 }],
        summary: {
          totalBags: stored.bag || 0,
          totalQty: stored.qty || 0,
          totalAmount: stored.amount || 0,
          gunnyCost: stored.gunny || 0,
          transportCost: stored.pava_hc || 0,
          insuranceCost: stored.ins || 0,
          cgst: stored.cgst || 0,
          sgst: stored.sgst || 0,
          igst: stored.igst || 0,
          tcs: stored.tcs || 0,
          roundDiff: stored.rund || 0,
          grandTotal: stored.tot || 0,
          isInterState: stored.sale === 'I',
        }
      };
    }

    // variant='purchase' flips the display: ISPL at top, ASP as seller, TN bank
    const pdf = await generateSalesInvoicePDF(invoice, cfg, stored.sale, stored.invo, stored.date, undefined, 'purchase');
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="PurchaseView_${stored.sale}_${stored.invo}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Purchase-view PDF error:', e);
    res.status(500).json({ error: 'PDF generation failed: ' + e.message });
  }
});

// Bulk Sales Invoice PDF — merges N invoices into a single PDF
// Body: { ids: [1, 2, 3, ...] }
// Returns: one PDF with each invoice on fresh page(s), in the order given.
app.post('/api/invoices/pdf-bulk', requireView, async (req, res) => {
  try {
    const ids = Array.isArray(req.body?.ids) ? req.body.ids.map(Number).filter(Boolean) : [];
    if (!ids.length) return res.status(400).json({ error: 'No invoice IDs provided' });
    // Optional dispatched-through override applied to every invoice in the batch.
    const dispatchedThrough = (req.body?.dispatchedThrough || '').toString();

    const db = getDb();
    const cfg = getSettingsFlat(db);

    // Enrich-buyer helper (same logic as single-invoice endpoint) — ensures
    // Bill-To / Ship-To have values even when rebuilding from stored data.
    const enrichBuyer = (buyer, stored) => {
      if (!buyer) buyer = {};
      if (!buyer.buyer1 && !buyer.buyer) {
        const looked = db.get('SELECT * FROM buyers WHERE buyer=? OR buyer1=? LIMIT 1',
          [stored.buyer, stored.buyer1 || stored.tradername || '']);
        if (looked) buyer = looked;
      }
      if (!buyer.buyer1 && stored.buyer1)   buyer.buyer1   = stored.buyer1;
      if (!buyer.buyer1 && stored.tradername) buyer.buyer1 = stored.tradername;
      if (!buyer.buyer  && stored.buyer)    buyer.buyer    = stored.buyer;
      if (!buyer.gstin  && stored.gstin)    buyer.gstin    = stored.gstin;
      if (!buyer.pla    && stored.place)    buyer.pla      = stored.place;
      if (!buyer.state  && stored.state)    buyer.state    = stored.state;
      if (!buyer.add1   && stored.add_line) buyer.add1     = stored.add_line;
      return buyer;
    };

    // Build each invoice's data
    const payloads = [];
    for (const id of ids) {
      const stored = db.get('SELECT * FROM invoices WHERE id=?', [id]);
      if (!stored) continue; // silently skip missing IDs
      let invoice = stored.auction_id
        ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg, { aspInvoice: String(stored.state || '').toUpperCase() === 'KERALA' })
        : null;
      if (invoice) {
        invoice.buyer = enrichBuyer(invoice.buyer, stored);
      } else {
        const buyer = enrichBuyer(db.get('SELECT * FROM buyers WHERE buyer=? LIMIT 1', [stored.buyer]), stored);
        invoice = {
          buyer,
          lineItems: [{ lot: '—', grade: '', bags: stored.bag || 0, qty: stored.qty || 0, price: 0, amount: stored.amount || 0 }],
          summary: {
            totalBags: stored.bag || 0, totalQty: stored.qty || 0,
            totalAmount: stored.amount || 0, gunnyCost: stored.gunny || 0,
            transportCost: stored.pava_hc || 0, insuranceCost: stored.ins || 0,
            taxableValue: (stored.amount || 0) + (stored.gunny || 0) + (stored.pava_hc || 0) + (stored.ins || 0),
            cgst: stored.cgst || 0, sgst: stored.sgst || 0, igst: stored.igst || 0,
            roundDiff: stored.rund || 0, grandTotal: stored.tot || 0,
            isInterState: stored.sale === 'I',
          }
        };
      }
      if (dispatchedThrough) invoice.dispatchedThrough = dispatchedThrough;
      // ASP cross-reference for ISP invoices — see single endpoint for rationale
      if (String(stored.state || '').toUpperCase() !== 'KERALA') {
        const aspRow = db.get(
          `SELECT asp_invo FROM lots
           WHERE auction_id = ? AND buyer = ? AND invo = ?
             AND asp_invo IS NOT NULL AND asp_invo != ''
           LIMIT 1`,
          [stored.auction_id, stored.buyer, stored.invo]
        );
        if (aspRow && aspRow.asp_invo) invoice.aspInvo = aspRow.asp_invo;
      }
      payloads.push({
        invoiceData: invoice,
        saleType: stored.sale,
        invoiceNo: stored.invo,
        invoiceDate: stored.date,
      });
    }

    if (!payloads.length) return res.status(404).json({ error: 'No invoices resolved from the provided IDs' });

    const pdf = await generateSalesInvoicesBatchPDF(payloads, cfg);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="Invoices_Batch_${payloads.length}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Bulk sales invoice PDF error:', e);
    res.status(500).json({ error: 'Batch PDF generation failed: ' + e.message });
  }
});

// Bulk Purchase-View PDF — like /pdf-bulk but renders each invoice with
// variant='purchase' so the buyer (ISPL) appears as the issuing company
// and the active ASP company appears as the seller. ASP context only.
// Body: { ids: [1, 2, 3, ...] }
app.post('/api/invoices/purchase-pdf-bulk', requireView, async (req, res) => {
  try {
    const ids = Array.isArray(req.body?.ids) ? req.body.ids.map(Number).filter(Boolean) : [];
    if (!ids.length) return res.status(400).json({ error: 'No invoice IDs provided' });

    const db = getDb();
    const cfg = getSettingsFlat(db);

    const isASPContext = (String(cfg.business_mode || '').toLowerCase() === 'e-trade')
                      && (String(cfg.business_state || '').toUpperCase() === 'KERALA');
    if (!isASPContext) {
      return res.status(400).json({
        error: 'Purchase view is only available for ASP invoices. Switch business state to KERALA + e-Trade mode.'
      });
    }

    // Same enrichBuyer pattern as /pdf-bulk
    const enrichBuyer = (buyer, stored) => {
      if (!buyer) buyer = {};
      if (!buyer.buyer1 && !buyer.buyer) {
        const looked = db.get('SELECT * FROM buyers WHERE buyer=? OR buyer1=? LIMIT 1',
          [stored.buyer, stored.buyer1 || stored.tradername || '']);
        if (looked) buyer = looked;
      }
      if (!buyer.buyer1 && stored.buyer1)   buyer.buyer1   = stored.buyer1;
      if (!buyer.buyer1 && stored.tradername) buyer.buyer1 = stored.tradername;
      if (!buyer.buyer  && stored.buyer)    buyer.buyer    = stored.buyer;
      if (!buyer.gstin  && stored.gstin)    buyer.gstin    = stored.gstin;
      if (!buyer.pla    && stored.place)    buyer.pla      = stored.place;
      if (!buyer.state  && stored.state)    buyer.state    = stored.state;
      if (!buyer.add1   && stored.add_line) buyer.add1     = stored.add_line;
      return buyer;
    };

    const payloads = [];
    for (const id of ids) {
      const stored = db.get('SELECT * FROM invoices WHERE id=?', [id]);
      if (!stored) continue;
      let invoice = stored.auction_id
        ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg, { aspInvoice: String(stored.state || '').toUpperCase() === 'KERALA' })
        : null;
      if (invoice) {
        invoice.buyer = enrichBuyer(invoice.buyer, stored);
      } else {
        const buyer = enrichBuyer(db.get('SELECT * FROM buyers WHERE buyer=? LIMIT 1', [stored.buyer]), stored);
        invoice = {
          buyer,
          lineItems: [{ lot: '—', grade: '', bags: stored.bag || 0, qty: stored.qty || 0, price: 0, amount: stored.amount || 0 }],
          summary: {
            totalBags: stored.bag || 0, totalQty: stored.qty || 0,
            totalAmount: stored.amount || 0, gunnyCost: stored.gunny || 0,
            transportCost: stored.pava_hc || 0, insuranceCost: stored.ins || 0,
            taxableValue: (stored.amount || 0) + (stored.gunny || 0) + (stored.pava_hc || 0) + (stored.ins || 0),
            cgst: stored.cgst || 0, sgst: stored.sgst || 0, igst: stored.igst || 0,
            roundDiff: stored.rund || 0, grandTotal: stored.tot || 0,
            isInterState: stored.sale === 'I',
          }
        };
      }
      payloads.push({
        invoiceData: invoice,
        saleType: stored.sale,
        invoiceNo: stored.invo,
        invoiceDate: stored.date,
      });
    }
    if (!payloads.length) return res.status(404).json({ error: 'No invoices resolved from the provided IDs' });

    // 'purchase' variant: ISPL at top, ASP as seller, TN bank — applied to every page
    const pdf = await generateSalesInvoicesBatchPDF(payloads, cfg, 'purchase');
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="PurchaseView_Batch_${payloads.length}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Bulk purchase-view PDF error:', e);
    res.status(500).json({ error: 'Batch PDF generation failed: ' + e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// PURCHASES (GSTKBILT.PRG — registered dealer invoices)
// ══════════════════════════════════════════════════════════════
app.get('/api/purchases', requireView, (req, res) => {
  const { auction_id, ano, from, to, sale, search } = req.query;
  const db = getDb();
  let q = 'SELECT * FROM purchases WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  // Free-text search across invoice no, seller name, GSTIN. Bypasses
  // the auction filter so cross-trade lookups work — the URL builder on
  // the client drops auction_id when search is non-empty.
  const searchTerm = String(search || '').trim();
  if (searchTerm) {
    const wild = `%${searchTerm}%`;
    q += ` AND (
            COALESCE(invo,'')  LIKE ?
            OR COALESCE(name,'')  LIKE ?
            OR COALESCE(gstin,'') LIKE ?
            OR COALESCE(place,'') LIKE ?
          )`;
    p.push(wild, wild, wild, wild);
  }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  // Sale-type filter (L / I / E). Purchases don't carry a sale column,
  // but the GST split on each row is deterministic:
  //   • L (Local / intra-state)  → CGST + SGST > 0, IGST = 0
  //   • I (Inter-state)          → IGST > 0, dealer has NO sale='E' lot
  //                                in the same auction
  //   • E (Export)               → IGST > 0, dealer HAS ≥1 sale='E' lot
  //                                in the same auction
  // Inferring directly from a sale column on purchases would fail —
  // that column doesn't exist. The GST-column approach is data-driven
  // and always correct as long as lots.sale was populated upstream.
  const saleNorm = String(sale || '').trim().toUpperCase();
  if (saleNorm === 'L') {
    q += ' AND COALESCE(igst,0) = 0 AND (COALESCE(cgst,0) > 0 OR COALESCE(sgst,0) > 0)';
  } else if (saleNorm === 'I') {
    q += ' AND COALESCE(igst,0) > 0';
    q += ` AND NOT EXISTS (
            SELECT 1 FROM lots l
             WHERE l.auction_id = purchases.auction_id
               AND UPPER(TRIM(COALESCE(l.name,''))) = UPPER(TRIM(COALESCE(purchases.name,'')))
               AND UPPER(TRIM(COALESCE(l.sale,''))) = 'E'
          )`;
  } else if (saleNorm === 'E') {
    q += ' AND COALESCE(igst,0) > 0';
    q += ` AND EXISTS (
            SELECT 1 FROM lots l
             WHERE l.auction_id = purchases.auction_id
               AND UPPER(TRIM(COALESCE(l.name,''))) = UPPER(TRIM(COALESCE(purchases.name,'')))
               AND UPPER(TRIM(COALESCE(l.sale,''))) = 'E'
          )`;
  }
  // Company split — match the purchase's stamped state to the active
  // business context, mirroring the Sales list. ASP rows (state contains
  // KERALA) show only when business state = KERALA; everything else
  // (including blank/legacy state on generated rows) shows in Tamil
  // Nadu. Import stamps state from the BR "ASP" prefix (rowDefaults);
  // generation stamps the auction state. The lenient NOT-LIKE on the
  // ISP side keeps blank-state generated purchases from disappearing.
  const _purBiz = String(getSettingsFlat(db).business_state || 'TAMIL NADU').toUpperCase();
  if (_purBiz === 'KERALA') {
    q += " AND UPPER(COALESCE(state,'')) LIKE '%KERALA%'";
  } else {
    q += " AND UPPER(COALESCE(state,'')) NOT LIKE '%KERALA%'";
  }
  const mw = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=purchases.auction_id)');
  q += mw.sql; p.push(...mw.params);
  q += ' ORDER BY date DESC LIMIT 500';
  res.json(db.all(q, p));
});

app.post('/api/purchases/generate/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const auctionIdForGate = parseInt(req.params.auctionId, 10);
  const _gen = _checkGenerationGate(db, 'purchases', auctionIdForGate);
  if (!_gen.allowed) return res.status(412).json(_gen.error);
  const { sellerName, invoiceNo } = req.body;
  const invoice = buildPurchaseInvoice(db, req.params.auctionId, sellerName, cfg);
  if (!invoice) return res.status(404).json({ error: 'No data for this seller' });
  
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const s = invoice.summary;
  db.run(`INSERT INTO purchases (auction_id,ano,date,state,br,name,add_line,place,gstin,invo,qty,amount,cgst,sgst,igst,rund,total,tds)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [req.params.auctionId,auction.ano,auction.date,auction.state||'','',invoice.seller.name,invoice.seller.address||'',
     invoice.seller.place||'',invoice.seller.cr||'',String(invoiceNo),s.totalQty,s.totalPuramt,
     s.totalCgst,s.totalSgst,s.totalIgst,s.roundDiff,s.grandTotal,s.tdsAmount]);
  res.json({ success: true, invoice: s });
});

// List eligible sellers for purchase invoices (with GSTIN, amount > 0)
app.get('/api/purchases/eligible-sellers/:auctionId', requireView, (req, res) => {
  res.json(getDb().all(
    `SELECT name, COUNT(*) as lot_count, SUM(qty) as total_qty, SUM(amount) as total_amount, MAX(cr) as cr
     FROM lots
     WHERE auction_id = ? AND name IS NOT NULL AND name != ''
       AND UPPER(cr) LIKE 'GSTIN%' AND amount > 0
     GROUP BY name
     ORDER BY name`,
    [req.params.auctionId]
  ));
});

// Batch: generate purchase invoice for ALL registered dealers in an auction
app.post('/api/purchases/generate-all/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const auctionIdForGate = parseInt(req.params.auctionId, 10);
  const _gen = _checkGenerationGate(db, 'purchases', auctionIdForGate);
  if (!_gen.allowed) return res.status(412).json(_gen.error);
  const { startInvoiceNo } = req.body;
  
  let nextNo = parseInt(startInvoiceNo);
  if (!nextNo || nextNo < 1) return res.status(400).json({ error: 'startInvoiceNo must be a positive integer' });
  
  // Auto-calculate
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  
  const sellers = db.all(
    `SELECT DISTINCT name FROM lots
     WHERE auction_id = ? AND UPPER(cr) LIKE 'GSTIN%' AND amount > 0 AND name IS NOT NULL AND name != ''`,
    [req.params.auctionId]
  );
  
  if (!sellers.length) return res.status(404).json({ error: 'No registered dealers (with GSTIN) in this auction' });
  
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const results = [];
  const errors = [];
  
  for (const row of sellers) {
    try {
      const invoice = buildPurchaseInvoice(db, req.params.auctionId, row.name, cfg);
      if (!invoice) { errors.push({ seller: row.name, error: 'Build failed' }); continue; }
      const s = invoice.summary;
      const invoNo = String(nextNo);
      db.run(`INSERT INTO purchases (auction_id,ano,date,state,br,name,add_line,place,gstin,invo,qty,amount,cgst,sgst,igst,rund,total,tds)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [req.params.auctionId,auction.ano,auction.date,auction.state||'','',invoice.seller.name,invoice.seller.address||'',
         invoice.seller.place||'',invoice.seller.cr||'',invoNo,s.totalQty,s.totalPuramt,
         s.totalCgst,s.totalSgst,s.totalIgst,s.roundDiff,s.grandTotal,s.tdsAmount]);
      results.push({ seller: row.name, invoiceNo: invoNo, grandTotal: s.grandTotal });
      nextNo++;
    } catch (e) { errors.push({ seller: row.name, error: e.message }); }
  }
  
  res.json({ success: true, generated: results.length, results, errors });
});

// Update purchase (edit)
app.put('/api/purchases/:id', requireInvoiceWrite, (req, res) => {
  const p = req.body;
  const fields = ['ano','date','state','br','name','add_line','place','gstin','invo',
    'qty','amount','cgst','sgst','igst','rund','total','tds'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (p[f] !== undefined) { sets.push(`${f}=?`); vals.push(p[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  getDb().run(`UPDATE purchases SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

// Delete purchase
app.delete('/api/purchases/:id', requireDelete, (req, res) => {
  getDb().run('DELETE FROM purchases WHERE id=?', [req.params.id]);
  res.json({ success: true });
});

// ── BULK DELETE — selected purchases ──
// Body: { ids: [1,2,3] }. Single round-trip replacement for the old
// per-trade Delete All flow. Returns the count actually deleted.
app.post('/api/purchases/bulk-delete', requireDelete, (req, res) => {
  const { ids } = req.body || {};
  if (!Array.isArray(ids) || !ids.length) {
    return res.status(400).json({ error: 'ids array required' });
  }
  const cleanIds = ids.map(Number).filter(Number.isFinite);
  if (!cleanIds.length) return res.status(400).json({ error: 'No valid purchase IDs' });
  const db = getDb();
  const placeholders = cleanIds.map(() => '?').join(',');
  const r = db.run(`DELETE FROM purchases WHERE id IN (${placeholders})`, cleanIds);
  res.json({ ok: true, deleted: r.changes });
});

// ── Purchase Invoice PDF ─────────────────────────────────────
app.get('/api/purchases/pdf/:auctionId/:sellerName', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const sellerName = decodeURIComponent(req.params.sellerName);
    const auctionId = req.params.auctionId;
    const invoiceNo = req.query.invoiceNo || '001';
    
    // Try to build fresh invoice from lots
    let invoice = buildPurchaseInvoice(db, auctionId, sellerName, cfg);
    
    // Fallback: if lots data missing, rebuild from stored purchase record
    if (!invoice) {
      // Try with auction_id first, then fall back to name+invo match (for older records)
      let stored = db.get(
        `SELECT * FROM purchases WHERE auction_id = ? AND name = ? AND invo = ? LIMIT 1`,
        [auctionId, sellerName, String(invoiceNo)]
      );
      if (!stored) {
        stored = db.get(
          `SELECT * FROM purchases WHERE name = ? AND invo = ? LIMIT 1`,
          [sellerName, String(invoiceNo)]
        );
      }
      if (!stored) {
        return res.status(404).json({ 
          error: `No purchase data found for seller "${sellerName}" with invoice ${invoiceNo}. Lots may have been deleted.` 
        });
      }
      // Reconstruct minimal invoice object from stored data
      invoice = {
        seller: { name: stored.name, address: stored.add_line, place: stored.place, cr: stored.gstin, pan: '', state: stored.state },
        lineItems: [{ lot: '—', qty: stored.qty, pqty: stored.qty, price: 0, prate: 0, amount: stored.amount, puramt: stored.amount, com: 0, sertax: 0, cgst: stored.cgst, sgst: stored.sgst, igst: stored.igst }],
        summary: {
          totalQty: stored.qty, totalPuramt: stored.amount,
          totalCgst: stored.cgst, totalSgst: stored.sgst, totalIgst: stored.igst,
          roundDiff: stored.rund, grandTotal: stored.total,
          tdsAmount: stored.tds, invoiceAmount: stored.total - stored.tds,
          isInter: stored.igst > 0
        }
      };
    }
    
    const pdf = await generatePurchaseInvoicePDF(
      enrichPurchaseForPDF(invoice, cfg, db, auctionId), cfg, invoiceNo
    );
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="PurchaseInvoice_${sellerName.replace(/[^\w]/g, '_')}_${invoiceNo}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('PDF generation error:', e);
    res.status(500).json({ error: 'PDF generation failed: ' + e.message });
  }
});

// Attach buyer block + e-TRADE No + invoice date to a purchase invoice
// object so the new renderer can show the full BILLED/SHIPPED TO + top
// grid correctly. The buyer is whichever identity is currently active in
// the app (ASP when Kerala+e-Trade, else ISP from company_settings).
function enrichPurchaseForPDF(invoice, cfg, db, auctionId) {
  if (!invoice) return invoice;
  // Stamp auction date + e-TRADE no
  if (!invoice.invoiceDate && auctionId) {
    const auction = db.get('SELECT date FROM auctions WHERE id = ?', [auctionId]);
    if (auction && auction.date) {
      const d = new Date(auction.date);
      if (!isNaN(d)) invoice.invoiceDate = d.toLocaleDateString('en-GB');
    }
  }
  if (!invoice.invoiceDate) invoice.invoiceDate = new Date().toLocaleDateString('en-GB');
  if (!invoice.eTradeNo) invoice.eTradeNo = String(auctionId || '');

  // Buyer block — ISPL or ASP depending on active context
  const isASP = cfg.business_mode === 'e-Trade' && String(cfg.business_state || '').toUpperCase() === 'KERALA';
  if (!invoice.buyer) {
    invoice.buyer = isASP ? {
      name: cfg.s_company || 'AMAZING SPICE PARK PRIVATE LIMITED',
      address: cfg.s_address1 || '',
      place: cfg.s_place || '',
      pin: cfg.s_pin || '',
      state: cfg.s_state || 'Kerala',
      st_code: cfg.s_st_code || '32',
      gstin: cfg.s_gstin || '',
      pan: cfg.s_pan || cfg.pan || '',
    } : {
      name: cfg.short_name || cfg.trade_name || 'IDEAL SPICES PRIVATE LIMITED',
      address: cfg.tn_address1 || '',
      place: cfg.tn_place || '',
      pin: cfg.tn_pin || '',
      state: cfg.tn_state || 'Tamil Nadu',
      st_code: cfg.tn_st_code || '33',
      gstin: cfg.tn_gstin || '',
      pan: cfg.pan || '',
    };
  }
  return invoice;
}

// Bulk Purchase Invoice PDF — merges N purchases into a single PDF
// Body: { ids: [1, 2, 3, ...] } — database row IDs from `purchases` table.
// Returns: one PDF with each purchase on its own page(s), in the order given.
// Same rebuild-from-lots OR fallback-to-stored pattern as the single route.
app.post('/api/purchases/pdf-bulk', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const ids = Array.isArray(req.body?.ids) ? req.body.ids.filter(Boolean) : [];
    if (!ids.length) return res.status(400).json({ error: 'No purchase IDs provided' });

    const placeholders = ids.map(() => '?').join(',');
    const rows = db.all(`SELECT * FROM purchases WHERE id IN (${placeholders})`, ids);
    if (!rows.length) return res.status(404).json({ error: 'No matching purchases found' });

    // Preserve the order the user ticked them in by looking up each ID in
    // the returned set (the IN query doesn't preserve order).
    const byId = new Map(rows.map(r => [r.id, r]));
    const ordered = ids.map(id => byId.get(Number(id))).filter(Boolean);

    const payloads = [];
    for (const stored of ordered) {
      // Try fresh rebuild from lots first (richer line-item detail)
      let invoice = stored.auction_id
        ? buildPurchaseInvoice(db, stored.auction_id, stored.name, cfg)
        : null;
      if (!invoice) {
        // Fallback: stored summary only (one line item)
        invoice = {
          seller: {
            name: stored.name, address: stored.add_line, place: stored.place,
            cr: stored.gstin, pan: '', state: stored.state
          },
          lineItems: [{
            lot: '—', qty: stored.qty, pqty: stored.qty, price: 0, prate: 0,
            amount: stored.amount, puramt: stored.amount,
            com: 0, sertax: 0, cgst: stored.cgst, sgst: stored.sgst, igst: stored.igst
          }],
          summary: {
            totalQty: stored.qty, totalPuramt: stored.amount,
            totalCgst: stored.cgst, totalSgst: stored.sgst, totalIgst: stored.igst,
            roundDiff: stored.rund, grandTotal: stored.total,
            tdsAmount: stored.tds, invoiceAmount: stored.total - stored.tds,
            isInter: stored.igst > 0
          }
        };
      }
      payloads.push({ invoiceData: enrichPurchaseForPDF(invoice, cfg, db, stored.auction_id), invoiceNo: stored.invo });
    }

    const pdf = await generatePurchaseInvoicesBatchPDF(payloads, cfg);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="PurchaseInvoices_Batch_${payloads.length}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Bulk purchase PDF error:', e);
    res.status(500).json({ error: 'Bulk PDF generation failed: ' + e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// BILLS — Agriculturist Bills of Supply (GSTKBILP/GSTBILP)
// ══════════════════════════════════════════════════════════════
app.get('/api/bills', requireView, (req, res) => {
  const { auction_id, ano, from, to, search, branch } = req.query;
  const db = getDb();
  let q = 'SELECT * FROM bills WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  // Free-text search across bill no, seller name and place. Bypasses
  // the auction filter when active so cross-trade lookups work — the
  // client drops auction_id when search is non-empty.
  const searchTerm = String(search || '').trim();
  if (searchTerm) {
    const wild = `%${searchTerm}%`;
    q += ` AND (
            COALESCE(CAST(bil AS TEXT),'') LIKE ?
            OR COALESCE(name,'')              LIKE ?
            OR COALESCE(pla,'')               LIKE ?
          )`;
    p.push(wild, wild, wild);
  }
  // Branch filter — populated from Settings → Branches (br1..br9). Matches
  // against bills.br first (when populated by the generator); falls back
  // to the underlying lot's branch via an EXISTS subquery so older bills
  // that didn't get the column stamped still surface under the right tab.
  const branchTerm = String(branch || '').trim();
  if (branchTerm) {
    q += ` AND (
            UPPER(COALESCE(bills.br,'')) = UPPER(?)
            OR EXISTS (
              SELECT 1 FROM lots l
               WHERE l.auction_id = bills.auction_id
                 AND UPPER(COALESCE(l.name,'')) = UPPER(COALESCE(bills.name,''))
                 AND UPPER(COALESCE(l.branch,'')) = UPPER(?)
            )
          )`;
    p.push(branchTerm, branchTerm);
  }
  // Two-company split — same convention as the Sales/Purchases lists:
  // `state` carries the issuing company (KERALA = ASP, else ISP). Import
  // Old Data tags Bills of Supply from the BR "ASP" prefix; generation
  // stamps it from the auction state. Filter by the active business
  // context so ASP bills show only in ASP and ISP only in ISP. Blank/
  // legacy state falls into the ISP bucket so nothing disappears.
  const _billBiz = String(getSettingsFlat(db).business_state || 'TAMIL NADU').toUpperCase();
  if (_billBiz === 'KERALA') {
    q += " AND UPPER(COALESCE(state,'')) LIKE '%KERALA%'";
  } else {
    q += " AND UPPER(COALESCE(state,'')) NOT LIKE '%KERALA%'";
  }
  const mw = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=bills.auction_id)');
  q += mw.sql; p.push(...mw.params);
  q += ' ORDER BY date DESC, bil DESC LIMIT 500';
  res.json(withFmtDate(db.all(q, p)));
});

// Generate agri bill for a seller
app.post('/api/bills/generate/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const auctionIdForGate = parseInt(req.params.auctionId, 10);
  const _gen = _checkGenerationGate(db, 'bills', auctionIdForGate);
  if (!_gen.allowed) return res.status(412).json(_gen.error);
  const { sellerName, billNo } = req.body;
  
  if (!sellerName || !billNo) {
    return res.status(400).json({ error: 'sellerName and billNo are required' });
  }
  
  // Auto-calculate if needed
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  
  const bill = buildAgriBill(db, req.params.auctionId, sellerName, cfg);
  if (!bill || bill.error) {
    return res.status(404).json({ error: bill?.error || 'No eligible lots found' });
  }
  
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const s = bill.summary;
  db.run(`INSERT INTO bills (auction_id,ano,date,state,br,crpt,bil,name,add_line,pla,pstate,st_code,crr,pan,qty,cost,igst,net)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [req.params.auctionId,auction.ano,auction.date,auction.state||'','',auction.crop_type||'ASP',
     parseInt(billNo),bill.seller.name,bill.seller.address||'',bill.seller.place||'',
     bill.seller.state||'',bill.seller.st_code||'',bill.seller.cr||'',bill.seller.pan||'',
     s.totalQty,s.totalPuramt,0,s.netAmount]);
  
  res.json({ success: true, bill: s });
});

// List eligible agri sellers for an auction (no GSTIN + amount > 0)
app.get('/api/bills/eligible-sellers/:auctionId', requireView, (req, res) => {
  res.json(listAgriSellers(getDb(), req.params.auctionId));
});

// Batch: generate bill of supply for ALL agriculturists (no GSTIN) in an auction
app.post('/api/bills/generate-all/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const auctionIdForGate = parseInt(req.params.auctionId, 10);
  const _gen = _checkGenerationGate(db, 'bills', auctionIdForGate);
  if (!_gen.allowed) return res.status(412).json(_gen.error);
  const { startBillNo } = req.body;
  
  let nextNo = parseInt(startBillNo);
  if (!nextNo || nextNo < 1) return res.status(400).json({ error: 'startBillNo must be a positive integer' });
  
  // Auto-calculate
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  
  const sellers = listAgriSellers(db, req.params.auctionId);
  if (!sellers.length) return res.status(404).json({ error: 'No agriculturist sellers (without GSTIN) with lots in this auction' });
  
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const results = [];
  const errors = [];
  
  for (const row of sellers) {
    try {
      const bill = buildAgriBill(db, req.params.auctionId, row.name, cfg);
      if (!bill || bill.error) { errors.push({ seller: row.name, error: bill?.error || 'Build failed' }); continue; }
      const s = bill.summary;
      db.run(`INSERT INTO bills (auction_id,ano,date,state,br,crpt,bil,name,add_line,pla,pstate,st_code,crr,pan,qty,cost,igst,net)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [req.params.auctionId,auction.ano,auction.date,auction.state||'','',auction.crop_type||'ASP',
         nextNo,bill.seller.name,bill.seller.address||'',bill.seller.place||'',
         bill.seller.state||'',bill.seller.st_code||'',bill.seller.cr||'',bill.seller.pan||'',
         s.totalQty,s.totalPuramt,0,s.netAmount]);
      results.push({ seller: row.name, billNo: nextNo, netAmount: s.netAmount });
      nextNo++;
    } catch (e) { errors.push({ seller: row.name, error: e.message }); }
  }
  
  res.json({ success: true, generated: results.length, results, errors });
});

// Agri bill PDF
app.get('/api/bills/pdf/:auctionId/:sellerName', requireView, async (req, res) => {
  try {
    const db = getDb(); const cfg = getSettingsFlat(db);
    const sellerName = decodeURIComponent(req.params.sellerName);
    const billNo = req.query.billNo || '001';
    
    let bill = buildAgriBill(db, req.params.auctionId, sellerName, cfg);
    if (!bill || bill.error) {
      // Fallback to stored record
      const stored = db.get('SELECT * FROM bills WHERE name = ? AND bil = ? LIMIT 1', [sellerName, parseInt(billNo)]);
      if (!stored) return res.status(404).json({ error: bill?.error || `No bill data found for "${sellerName}"` });
      bill = {
        seller: { name: stored.name, address: stored.add_line, place: stored.pla, state: stored.pstate, st_code: stored.st_code, cr: stored.crr, crno: stored.crr, pan: stored.pan },
        lineItems: [{ lot: '—', qty: stored.qty, pqty: stored.qty, prate: 0, amount: stored.cost, puramt: stored.cost }],
        summary: { totalQty: stored.qty, totalPuramt: stored.cost, roundDiff: 0, netAmount: stored.net, cgst: 0, sgst: 0, igst: 0, tax: 0 }
      };
    }
    // Enrich seller.crno so the new renderer can display "CR.<n>" in the
    // details block when CR/GSTIN-style id is available on the trader row
    if (bill.seller && !bill.seller.crno) bill.seller.crno = bill.seller.cr || '';
    // Stamp the bill date + e-TRADE number so the new layout can render them
    // in the top strip (Invoice No / e-TRADE No / Date).
    const auction = db.get('SELECT date FROM auctions WHERE id = ?', [req.params.auctionId]);
    if (auction && auction.date) {
      const d = new Date(auction.date);
      if (!isNaN(d)) bill.billDate = d.toLocaleDateString('en-GB');
    }
    if (!bill.billDate) bill.billDate = new Date().toLocaleDateString('en-GB');
    bill.eTradeNo = req.query.eTradeNo || req.params.auctionId;

    const pdf = await generateAgriBillPDF(bill, cfg, billNo);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="BillOfSupply_${sellerName.replace(/[^\w]/g,'_')}_${billNo}.pdf"`);
    res.send(pdf);
  } catch(e) {
    console.error('Agri Bill PDF error:', e);
    res.status(500).json({ error: 'PDF generation failed: ' + e.message });
  }
});

// Bulk Agri Bill PDF — merges N bills into a single PDF.
// Body: { ids: [1, 2, 3, ...] } — DB row IDs from the `bills` table.
app.post('/api/bills/pdf-bulk', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const ids = Array.isArray(req.body?.ids) ? req.body.ids.filter(Boolean) : [];
    if (!ids.length) return res.status(400).json({ error: 'No bill IDs provided' });

    const placeholders = ids.map(() => '?').join(',');
    const rows = db.all(`SELECT * FROM bills WHERE id IN (${placeholders})`, ids);
    if (!rows.length) return res.status(404).json({ error: 'No matching bills found' });

    const byId = new Map(rows.map(r => [r.id, r]));
    const ordered = ids.map(id => byId.get(Number(id))).filter(Boolean);

    const payloads = [];
    for (const stored of ordered) {
      let bill = stored.auction_id
        ? buildAgriBill(db, stored.auction_id, stored.name, cfg)
        : null;
      if (!bill || bill.error) {
        bill = {
          seller: {
            name: stored.name, address: stored.add_line, place: stored.pla,
            state: stored.pstate, st_code: stored.st_code, cr: stored.crr, crno: stored.crr, pan: stored.pan
          },
          lineItems: [{
            lot: '—', qty: stored.qty, pqty: stored.qty, prate: 0,
            amount: stored.cost, puramt: stored.cost
          }],
          summary: {
            totalQty: stored.qty, totalPuramt: stored.cost, roundDiff: 0,
            netAmount: stored.net, cgst: 0, sgst: 0, igst: 0, tax: 0
          }
        };
      }
      // Enrich for new renderer layout (Invoice No / e-TRADE No / Date strip)
      if (bill.seller && !bill.seller.crno) bill.seller.crno = bill.seller.cr || '';
      if (stored.auction_id) {
        const auction = db.get('SELECT date FROM auctions WHERE id = ?', [stored.auction_id]);
        if (auction && auction.date) {
          const d = new Date(auction.date);
          if (!isNaN(d)) bill.billDate = d.toLocaleDateString('en-GB');
        }
      }
      if (!bill.billDate) bill.billDate = new Date().toLocaleDateString('en-GB');
      bill.eTradeNo = stored.auction_id || '';
      payloads.push({ billData: bill, billNo: stored.bil });
    }

    const pdf = await generateAgriBillsBatchPDF(payloads, cfg);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="BillsOfSupply_Batch_${payloads.length}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Bulk bill PDF error:', e);
    res.status(500).json({ error: 'Bulk PDF generation failed: ' + e.message });
  }
});

app.delete('/api/bills/:id', requireDelete, (req, res) => {
  getDb().run('DELETE FROM bills WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// ── BULK DELETE — selected bills ──
// Body: { ids: [1,2,3] }. Single round-trip replacement for the old
// trade-wide Delete All flow.
app.post('/api/bills/bulk-delete', requireDelete, (req, res) => {
  const { ids } = req.body || {};
  if (!Array.isArray(ids) || !ids.length) {
    return res.status(400).json({ error: 'ids array required' });
  }
  const cleanIds = ids.map(Number).filter(Number.isFinite);
  if (!cleanIds.length) return res.status(400).json({ error: 'No valid bill IDs' });
  const db = getDb();
  const placeholders = cleanIds.map(() => '?').join(',');
  const r = db.run(`DELETE FROM bills WHERE id IN (${placeholders})`, cleanIds);
  res.json({ ok: true, deleted: r.changes });
});

// Update bill (edit)
app.put('/api/bills/:id', requireInvoiceWrite, (req, res) => {
  const b = req.body;
  const fields = ['ano','date','state','br','crpt','bil','name','add_line','pla','pstate','st_code','crr','pan','qty','cost','igst','net'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (b[f] !== undefined) { sets.push(`${f}=?`); vals.push(b[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  getDb().run(`UPDATE bills SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

// ══════════════════════════════════════════════════════════════
// DEBIT NOTES (for discounts/adjustments)
// ══════════════════════════════════════════════════════════════
app.get('/api/debit-notes', requireView, (req, res) => {
  const { auction_id, ano, from, to, search } = req.query;
  const db = getDb();
  let q = 'SELECT * FROM debit_notes WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  // Free-text search across note no, seller name, and trade no (ano).
  // Cross-trade — the client drops the auction filter when search is
  // non-empty so dad can find a stray note from another trade by
  // typing the dealer name.
  const searchTerm = String(search || '').trim();
  if (searchTerm) {
    const wild = `%${searchTerm}%`;
    q += ` AND (
            COALESCE(note_no,'') LIKE ?
            OR COALESCE(name,'')    LIKE ?
            OR COALESCE(ano,'')     LIKE ?
          )`;
    p.push(wild, wild, wild);
  }
  const mw = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=debit_notes.auction_id)');
  q += mw.sql; p.push(...mw.params);
  q += ' ORDER BY date DESC, note_no DESC LIMIT 500';
  res.json(withFmtDate(db.all(q, p)));
});

// ── Debit note math helpers ──────────────────────────────────────
// Local to the DN endpoints since these were imported from the
// reference build's standalone DN routes. Both are stateless one-liners.
function _dnRound2(n) { return Math.round(Number(n || 0) * 100) / 100; }
function _dnAddDays(iso, days) {
  const d = new Date(String(iso || '').slice(0, 10) + 'T00:00:00');
  if (isNaN(d.getTime())) return new Date().toISOString().slice(0, 10);
  d.setDate(d.getDate() + (Number(days) || 0));
  return d.toISOString().slice(0, 10);
}
// Strip a "GSTIN."/"GSTIN" prefix, uppercase, trim. Used for the
// state-code extraction below.
function _dnNormGstin(s) {
  let v = String(s == null ? '' : s).trim().toUpperCase();
  if (v.startsWith('GSTIN.')) v = v.slice(6);
  else if (v.startsWith('GSTIN')) v = v.slice(5);
  return v.trim();
}

// ── Generate Debit Note (single, trade-scoped) ───────────────────
// Inputs: purchno (purchase invoice number) + ano (trade) + optional
// startNoteNo. Discount is auto-derived from settings (discount_pct
// and discount_days). GST split classified by the DEALER's GSTIN
// state code vs the company's state code. Numbering is TRADE-WISE —
// each trade has its own 1..N sequence.
app.post('/api/debit-notes/generate', requireInvoiceWrite, (req, res) => {
  const db = getDb();
  const cfg = getSettingsFlat(db);
  const purchno = String(req.body.purchno || req.body.invoiceNo || '').trim();
  const ano     = String(req.body.ano || '').trim();

  if (!purchno) return res.status(400).json({ error: 'purchno (purchase invoice number) is required' });
  if (!ano)     return res.status(400).json({ error: 'ano (trade number) is required' });

  // Generation lock: once every purchase in this trade has a debit
  // note, the generate is blocked until an admin grants a one-shot
  // override. While any purchase is still un-DN'd it stays open.
  {
    const _dnAuc = db.get('SELECT id FROM auctions WHERE ano = ? ORDER BY date DESC LIMIT 1', [ano]);
    if (_dnAuc) {
      const _gen = _checkGenerationGate(db, 'debit_notes', _dnAuc.id);
      if (!_gen.allowed) return res.status(412).json(_gen.error);
    }
  }

  // Look up the purchase. Multiple rows can share an invo across
  // trades — pick the one matching the requested trade.
  const candidates = db.all(
    `SELECT * FROM purchases WHERE invo = ? ORDER BY date DESC, id DESC`,
    [purchno]
  );
  if (!candidates.length) {
    // Distinguish "is a sales invoice" from "doesn't exist".
    const isSalesInv = db.get(`SELECT id FROM invoices WHERE invo = ? LIMIT 1`, [purchno]);
    if (isSalesInv) {
      return res.status(400).json({
        error: `${purchno} is a SALES invoice. Debit notes can only be generated against PURCHASE invoices.`
      });
    }
    return res.status(404).json({ error: `Purchase invoice ${purchno} not found` });
  }

  const purchase = candidates.find(p => String(p.ano) === ano);
  if (!purchase) {
    const otherAnos = [...new Set(candidates.map(p => String(p.ano)))].join(', ');
    return res.status(400).json({
      error: `Purchase invoice ${purchno} does not belong to trade #${ano}. It belongs to trade #${otherAnos}.`
    });
  }

  // Idempotency: skip if a DN for (ano, dealer) already exists.
  const dealerName = purchase.name || '';
  const dupe = db.get(
    `SELECT id, note_no FROM debit_notes WHERE ano = ? AND name = ? LIMIT 1`,
    [ano, dealerName]
  );
  if (dupe) {
    return res.status(409).json({
      error: `Debit note #${dupe.note_no} already exists for ${dealerName} in trade #${ano}`,
      existingId: dupe.id,
      existingNoteNo: dupe.note_no,
    });
  }

  // Discount: explicit override > computed from settings.
  const baseAmt = Number(purchase.amount || 0);
  if (baseAmt <= 0) {
    return res.status(400).json({ error: 'Purchase amount is zero — cannot compute discount' });
  }
  let discountAmt = req.body.discount != null ? parseFloat(req.body.discount) : NaN;
  if (!Number.isFinite(discountAmt) || discountAmt <= 0) {
    const discountPct  = Number(cfg.discount_pct)  || 0;
    const discountDays = Number(cfg.discount_days) || 0;
    if (discountPct <= 0) {
      return res.status(400).json({ error: 'Discount % not configured in settings' });
    }
    discountAmt = discountDays > 0
      ? Math.round((baseAmt / 1000) * discountDays * discountPct)
      : Math.round(baseAmt * discountPct / 100);
  }
  if (discountAmt <= 0) {
    return res.status(400).json({ error: 'Computed discount is zero — check settings or invoice amount' });
  }

  // GST split — classify by the DEALER's GSTIN state code vs the
  // company's state code. Not by purchase.igst (which can be stale).
  const dealerG = _dnNormGstin(purchase.gstin);
  const dealerStateCode = /^\d{2}/.test(dealerG) ? dealerG.slice(0, 2) : '';
  const companyStateCode = String(cfg.tally_state_code
      || (String(cfg.business_state || '').toUpperCase() === 'KERALA' ? '32' : '33'));
  const isInter = !!dealerStateCode && dealerStateCode !== companyStateCode;

  const dnGstRate = Number(cfg.gst_service) || 18;
  // Only emit GST when (a) flag_disc_gst is on AND (b) the source
  // purchase carried GST (registered dealer). URD/agri purchases
  // produce exempt DNs regardless.
  const flagDiscGst = String(cfg.flag_disc_gst || '').toLowerCase() === 'true' || cfg.flag_disc_gst === true;
  const dealerCarriedGst = Number(purchase.cgst) || Number(purchase.sgst) || Number(purchase.igst);
  let cgst = 0, sgst = 0, igst = 0;
  if (flagDiscGst && dealerCarriedGst) {
    if (isInter) {
      igst = _dnRound2(discountAmt * dnGstRate / 100);
    } else {
      const half = _dnRound2(discountAmt * (dnGstRate / 2) / 100);
      cgst = half; sgst = half;
    }
  }
  const total = _dnRound2(discountAmt + cgst + sgst + igst);

  // DN date = trade.date + 1.
  const trade = db.get('SELECT date FROM auctions WHERE ano = ? LIMIT 1', [ano]);
  const dnDate = trade && trade.date
    ? _dnAddDays(trade.date, 1)
    : new Date().toISOString().slice(0, 10);

  // Note number: client-supplied `startNoteNo` (preferred) or fall
  // back to MAX(note_no)+1 within this trade. Validated as a positive
  // integer; uniqueness scoped to the selected trade only.
  const rawStart = req.body.startNoteNo != null ? req.body.startNoteNo : req.body.noteNo;
  let noteNo;
  if (rawStart != null && String(rawStart).trim() !== '') {
    const n = parseInt(String(rawStart).trim(), 10);
    if (!Number.isFinite(n) || n < 1) {
      return res.status(400).json({ error: 'Starting Number must be a positive integer' });
    }
    noteNo = String(n);
    const taken = db.get(
      `SELECT id FROM debit_notes WHERE ano = ? AND CAST(note_no AS INTEGER) = ? LIMIT 1`,
      [ano, n]
    );
    if (taken) {
      const row = db.get('SELECT MAX(CAST(note_no AS INTEGER)) AS mx FROM debit_notes WHERE ano = ?', [ano]);
      const mx = parseInt(row && row.mx, 10);
      const safe = Number.isFinite(mx) && mx > 0 ? mx + 1 : 1;
      return res.status(409).json({
        error: `Debit note #${n} is already used in trade #${ano}. Choose a different number.`,
        suggested: safe,
      });
    }
  } else {
    const row = db.get('SELECT MAX(CAST(note_no AS INTEGER)) AS mx FROM debit_notes WHERE ano = ?', [ano]);
    const mx = parseInt(row && row.mx, 10);
    noteNo = String(Number.isFinite(mx) && mx > 0 ? mx + 1 : 1);
  }

  db.run(
    `INSERT INTO debit_notes (ano,date,state,name,note_no,amount,cgst,sgst,igst,total)
     VALUES (?,?,?,?,?,?,?,?,?,?)`,
    [ano, dnDate, purchase.state || '', dealerName,
     noteNo, discountAmt, cgst, sgst, igst, total]
  );

  res.json({
    success: true, created: 1, note_no: noteNo,
    purchno, ano, dealer: dealerName,
    amount: discountAmt, cgst, sgst, igst, total,
  });
});

// ── Generate All Debit Notes (trade-scoped bulk) ─────────────────
// Body: { ano, startNoteNo }. Generates one DN per eligible purchase
// in the selected trade (i.e. purchases without an existing DN for
// the same (ano, dealer) pair, amount > 0). DNs get sequential note
// numbers starting at startNoteNo. Returns { created, skipped,
// generated[], skippedDetails[] }.
app.post('/api/debit-notes/generate-bulk', requireInvoiceWrite, (req, res) => {
  const db = getDb();
  const cfg = getSettingsFlat(db);

  // Resolve trade number. Two input shapes: { ano } (preferred) or
  // { purchno } (legacy — derive ano from the single purchase).
  let ano = String(req.body.ano || '').trim();
  if (!ano) {
    const purchno = String(req.body.purchno || '').trim();
    if (!purchno) return res.status(400).json({ error: 'Trade number (ano) is required' });
    const p = db.get(
      `SELECT ano FROM purchases WHERE invo = ? ORDER BY date DESC, id DESC LIMIT 1`,
      [purchno]
    );
    if (!p) return res.status(404).json({ error: `Purchase invoice ${purchno} not found` });
    ano = String(p.ano || '').trim();
    if (!ano) return res.status(400).json({ error: 'Purchase row has no trade number' });
  }

  // Generation lock — admin must grant an override before re-running
  // generate-all once every purchase in this trade has a debit note.
  // While any purchase is still un-DN'd, generate-all stays open and
  // its dedupe layer skips the already-done rows.
  {
    const _dnAuc = db.get('SELECT id FROM auctions WHERE ano = ? ORDER BY date DESC LIMIT 1', [ano]);
    if (_dnAuc) {
      const _gen = _checkGenerationGate(db, 'debit_notes', _dnAuc.id);
      if (!_gen.allowed) return res.status(412).json(_gen.error);
    }
  }

  // Pull every purchase row for this trade.
  const purchases = db.all(`SELECT * FROM purchases WHERE ano = ? ORDER BY id`, [ano]);
  if (!purchases.length) {
    return res.json({
      success: true, created: 0, skipped: 0, generated: [], skippedDetails: [],
      note: `No purchase invoices in trade #${ano}`,
    });
  }

  // Existing DN keys for this trade (single query — cheap).
  const existingKeys = new Set(
    db.all(`SELECT name FROM debit_notes WHERE ano = ?`, [ano]).map(r => r.name || '')
  );

  // Resolve DN date once per trade.
  const trade = db.get('SELECT date FROM auctions WHERE ano = ? LIMIT 1', [ano]);
  const dnDate = trade && trade.date
    ? _dnAddDays(trade.date, 1)
    : new Date().toISOString().slice(0, 10);

  // Discount math constants — read once, applied per-purchase.
  const discountPct  = Number(cfg.discount_pct)  || 0;
  const discountDays = Number(cfg.discount_days) || 0;
  const dnGstRate    = Number(cfg.gst_service) || 18;
  const flagDiscGst  = String(cfg.flag_disc_gst || '').toLowerCase() === 'true' || cfg.flag_disc_gst === true;
  if (discountPct <= 0) {
    return res.status(400).json({ error: 'Discount % not configured in settings' });
  }
  const companyStateCode = String(cfg.tally_state_code
      || (String(cfg.business_state || '').toUpperCase() === 'KERALA' ? '32' : '33'));

  // Eligible count for the up-front range check.
  const eligibleCount = purchases.filter(
    p => !existingKeys.has(p.name || '') && Number(p.amount || 0) > 0
  ).length;

  // Resolve next note number. User-supplied `startNoteNo` anchors the
  // sequence; omitted → MAX+1 within this trade.
  let nextNoteNo;
  const rawStart = req.body.startNoteNo != null ? req.body.startNoteNo : req.body.startInvoiceNo;
  if (rawStart != null && String(rawStart).trim() !== '') {
    const n = parseInt(String(rawStart).trim(), 10);
    if (!Number.isFinite(n) || n < 1) {
      return res.status(400).json({ error: 'Starting Number must be a positive integer' });
    }
    nextNoteNo = n;
    if (eligibleCount > 0) {
      // Range claim — scoped to THIS TRADE only.
      const upper = nextNoteNo + eligibleCount - 1;
      const collisions = db.all(
        `SELECT CAST(note_no AS INTEGER) AS n
           FROM debit_notes
          WHERE ano = ? AND CAST(note_no AS INTEGER) BETWEEN ? AND ?
          ORDER BY n`,
        [ano, nextNoteNo, upper]
      );
      if (collisions.length) {
        const row = db.get('SELECT MAX(CAST(note_no AS INTEGER)) AS mx FROM debit_notes WHERE ano = ?', [ano]);
        const mx = parseInt(row && row.mx, 10);
        const safe = Number.isFinite(mx) && mx > 0 ? mx + 1 : 1;
        return res.status(409).json({
          error: `Starting Number ${nextNoteNo} would overlap existing debit note(s) in trade #${ano} `
               + `(${collisions.slice(0, 5).map(c => '#' + c.n).join(', ')}`
               + `${collisions.length > 5 ? `, +${collisions.length - 5} more` : ''}). Try ${safe} or higher.`,
          collisions: collisions.map(c => c.n),
          suggested: safe,
        });
      }
    }
  } else {
    const row = db.get('SELECT MAX(CAST(note_no AS INTEGER)) AS mx FROM debit_notes WHERE ano = ?', [ano]);
    const mx = parseInt(row && row.mx, 10);
    nextNoteNo = Number.isFinite(mx) && mx > 0 ? mx + 1 : 1;
  }

  const generated = [];
  const skipped   = [];

  for (const p of purchases) {
    const dealerName = p.name || '';
    if (existingKeys.has(dealerName)) {
      skipped.push({ invo: p.invo, ano, buyer: dealerName, reason: 'duplicate (DN already exists for this dealer in this trade)' });
      continue;
    }
    const baseAmt = Number(p.amount || 0);
    if (baseAmt <= 0) {
      skipped.push({ invo: p.invo, ano, buyer: dealerName, reason: 'zero amount' });
      continue;
    }
    const discountAmt = discountDays > 0
      ? Math.round((baseAmt / 1000) * discountDays * discountPct)
      : Math.round(baseAmt * discountPct / 100);
    if (discountAmt <= 0) {
      skipped.push({ invo: p.invo, ano, buyer: dealerName, reason: 'computed discount is zero' });
      continue;
    }

    // Intra/inter classification by dealer's GSTIN state code.
    const dealerG = _dnNormGstin(p.gstin);
    const dealerStateCode = /^\d{2}/.test(dealerG) ? dealerG.slice(0, 2) : '';
    const isInter = !!dealerStateCode && dealerStateCode !== companyStateCode;
    const dealerCarriedGst = Number(p.cgst) || Number(p.sgst) || Number(p.igst);
    let cgst = 0, sgst = 0, igst = 0;
    if (flagDiscGst && dealerCarriedGst) {
      if (isInter) {
        igst = _dnRound2(discountAmt * dnGstRate / 100);
      } else {
        const half = _dnRound2(discountAmt * (dnGstRate / 2) / 100);
        cgst = half; sgst = half;
      }
    }
    const total = _dnRound2(discountAmt + cgst + sgst + igst);

    db.run(
      `INSERT INTO debit_notes (ano,date,state,name,note_no,amount,cgst,sgst,igst,total)
       VALUES (?,?,?,?,?,?,?,?,?,?)`,
      [ano, dnDate, p.state || '', dealerName,
       String(nextNoteNo), discountAmt, cgst, sgst, igst, total]
    );
    generated.push({ note_no: nextNoteNo, purchno: p.invo, dealer: dealerName, total });
    existingKeys.add(dealerName);
    nextNoteNo++;
  }

  res.json({
    success: true,
    created: generated.length,
    skipped: skipped.length,
    generated,
    skippedDetails: skipped,
    note: generated.length === 0 && skipped.length === 0
      ? `No eligible purchases in trade #${ano}`
      : undefined,
  });
});

// List purchases in a trade that don't yet have a DN. Drives the
// preview panel in the Generate All modal so the user sees exactly
// what will be created before clicking Generate.
app.get('/api/debit-notes/eligible-purchases/:auctionId', requireView, (req, res) => {
  const db = getDb();
  const auction = db.get('SELECT ano FROM auctions WHERE id = ?', [req.params.auctionId]);
  if (!auction) return res.status(404).json({ error: 'Auction not found' });
  const ano = auction.ano;
  const rows = db.all(
    `SELECT p.id, p.invo, p.name, p.amount, p.cgst, p.sgst, p.igst, p.total, p.date, p.state
       FROM purchases p
      WHERE p.ano = ? AND p.amount > 0
        AND NOT EXISTS (
          SELECT 1 FROM debit_notes dn WHERE dn.ano = p.ano AND dn.name = p.name
        )
      ORDER BY p.id`,
    [ano]
  );
  res.json(rows);
});

// Next-available DN number for a trade. Trade-wise numbering — each
// trade has its own 1..N sequence. Used by both Generate modals to
// pre-fill the Starting Number field.
app.get('/api/debit-notes/next-note-no', requireView, (req, res) => {
  const db = getDb();
  const ano = String(req.query.ano || '').trim();
  if (!ano) return res.status(400).json({ error: 'ano (trade number) is required' });
  const row = db.get(
    'SELECT MAX(CAST(note_no AS INTEGER)) AS mx FROM debit_notes WHERE ano = ?', [ano]
  );
  const mx = parseInt(row && row.mx, 10);
  const next = Number.isFinite(mx) && mx > 0 ? mx + 1 : 1;
  res.json({ next, ano });
});

app.delete('/api/debit-notes/:id', requireDelete, (req, res) => {
  getDb().run('DELETE FROM debit_notes WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// ── BULK DELETE — selected debit notes ──
// Body: { ids: [1,2,3] }. Single round-trip replacement for the trade-
// wide Delete All. Returns the number of rows actually deleted.
app.post('/api/debit-notes/bulk-delete', requireDelete, (req, res) => {
  const { ids } = req.body || {};
  if (!Array.isArray(ids) || !ids.length) {
    return res.status(400).json({ error: 'ids array required' });
  }
  const cleanIds = ids.map(Number).filter(Number.isFinite);
  if (!cleanIds.length) return res.status(400).json({ error: 'No valid debit note IDs' });
  const db = getDb();
  const placeholders = cleanIds.map(() => '?').join(',');
  const r = db.run(`DELETE FROM debit_notes WHERE id IN (${placeholders})`, cleanIds);
  res.json({ ok: true, deleted: r.changes });
});

// Update debit note (edit)
app.put('/api/debit-notes/:id', requireInvoiceWrite, (req, res) => {
  const n = req.body;
  const fields = ['ano','date','state','name','note_no','amount','cgst','sgst','igst','total'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (n[f] !== undefined) { sets.push(`${f}=?`); vals.push(n[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  getDb().run(`UPDATE debit_notes SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

// ══════════════════════════════════════════════════════════════
// JOURNALS (JOUR.PRG, PUJOUR.PRG, PPUJOUR.PRG)
// ══════════════════════════════════════════════════════════════
app.get('/api/journals/sales', requireView, (req, res) => {
  const { from, to, saleType } = req.query;
  if (!from || !to) return res.status(400).json({ error: 'from and to required' });
  res.json(getSalesJournal(getDb(), from, to, saleType));
});

app.get('/api/journals/purchase', requireView, (req, res) => {
  const { from, to, type } = req.query;
  if (!from || !to) return res.status(400).json({ error: 'from and to required' });
  res.json(getPurchaseJournal(getDb(), from, to, type || 'dealer'));
});

// Journal exports (XLSX only)
app.get('/api/exports/sales-journal', requireExport, async (req, res) => {
  const { from, to, saleType } = req.query;
  if (!from || !to) return res.status(400).json({ error: 'from/to required' });
  const { exportSalesJournal } = require('./exports');
  const buffer = await exportSalesJournal(getDb(), from, to, saleType);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="SalesJournal.xlsx"');
  res.send(Buffer.from(buffer));
});

app.get('/api/exports/purchase-journal', requireExport, async (req, res) => {
  const { from, to, type } = req.query;
  if (!from || !to) return res.status(400).json({ error: 'from/to required' });
  const baseName = type === 'agri' ? 'AgriBillJournal' : 'PurchaseJournal';
  const { exportPurchaseJournal } = require('./exports');
  const buffer = await exportPurchaseJournal(getDb(), from, to, type || 'dealer');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${baseName}.xlsx"`);
  res.send(Buffer.from(buffer));
});

// ══════════════════════════════════════════════════════════════
// INVOICE PREVIEW (PREINVO.PRG) — dry-run, no save
// ══════════════════════════════════════════════════════════════
app.post('/api/invoices/preview/:auctionId', requireView, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { saleType, buyerCode, type } = req.body;
  
  // Auto-calculate any uncalculated lots first (read-only would be better but we need the data)
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  
  let invoice;
  if (type === 'purchase') {
    invoice = buildPurchaseInvoice(db, req.params.auctionId, buyerCode, cfg); // buyerCode = sellerName for purchase
  } else if (type === 'agri') {
    invoice = buildAgriBill(db, req.params.auctionId, buyerCode, cfg);
    if (invoice && invoice.error) return res.status(404).json({ error: invoice.error });
  } else {
    invoice = buildSalesInvoice(db, req.params.auctionId, buyerCode, saleType, cfg);
  }
  
  if (!invoice) return res.status(404).json({ error: 'No data found' });
  res.json({ preview: true, invoice });
});

// ══════════════════════════════════════════════════════════════
// PAYMENTS (PAYCHECK.PRG)
// ══════════════════════════════════════════════════════════════
app.get('/api/payments/:auctionId', requireView, (req, res) => {
  const db = getDb();
  const cfg = getSettingsFlat(db);
  const summary = getPaymentSummary(db, req.params.auctionId, req.query.state, cfg);
  res.json(summary);
});

// ── Bank payment data (BANKPAY.PRG) ──────────────────────────
app.get('/api/payments/bank/:auctionId', requireView, (req, res) => {
  const cfg = getSettingsFlat(getDb());
  const data = getBankPaymentData(getDb(), req.params.auctionId, cfg);
  res.json(data);
});

// ── Payment statement PDF ────────────────────────────────────
// Per-seller payment statement. `lotIds` (optional) narrows the
// statement to a caller-chosen subset of the seller's lots — powers
// the Payments tab "partial payment" flow (operator opens the lots
// modal, ticks the lots being settled now, prints just those).
function _renderPaymentStatement(doc, db, auctionId, sellerName, cfg, lotIds) {
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]) || { ano:'', date:'' };
  // Match seller by trimmed/case-insensitive name so legacy rows whose
  // `name` was stored with trailing whitespace or mixed case still pair
  // up. `prate` (per-kg purchase rate) is aliased as `rate` for display.
  const lotIdFilter = Array.isArray(lotIds)
    ? lotIds.map(n => parseInt(n, 10)).filter(Number.isFinite)
    : [];
  let lotSql = `SELECT id, lot_no, qty, prate AS rate, amount, puramt, refund, balance, cgst, sgst, igst
       FROM lots
      WHERE auction_id = ?
        AND TRIM(LOWER(COALESCE(name,''))) = TRIM(LOWER(?))
        AND amount > 0`;
  const lotParams = [auctionId, sellerName];
  if (lotIdFilter.length) {
    const placeholders = lotIdFilter.map(() => '?').join(',');
    lotSql += ` AND id IN (${placeholders})`;
    lotParams.push(...lotIdFilter);
  }
  lotSql += ' ORDER BY CAST(lot_no AS INTEGER), lot_no';
  const lots = db.all(lotSql, lotParams) || [];
  const trader = db.get('SELECT * FROM traders WHERE LOWER(name) = LOWER(?) LIMIT 1', [sellerName]);
  const fmtAmt = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmtQty = (n) => Number(n || 0).toLocaleString('en-IN', { minimumFractionDigits: 3, maximumFractionDigits: 3 });
  const company = (cfg.trade_name || cfg.short_name || cfg.tally_company_name || cfg.legal_name || 'Company').toString();

  const PAGE_L = 30, PAGE_R = 565, PAGE_W = PAGE_R - PAGE_L;
  let y = 40;
  doc.font('Helvetica-Bold').fontSize(16).text(company.toUpperCase(), PAGE_L, y, { width: PAGE_W, align: 'center' });
  y = doc.y + 4;
  doc.font('Helvetica-Bold').fontSize(13).text('PAYMENT STATEMENT', PAGE_L, y, { width: PAGE_W, align: 'center' });
  y = doc.y + 10;
  doc.moveTo(PAGE_L, y).lineTo(PAGE_R, y).lineWidth(1).stroke();
  y += 10;

  doc.font('Helvetica').fontSize(10);
  doc.text(`Seller: ${sellerName}`, PAGE_L, y); doc.text(`Auction: ${auction.ano}`, PAGE_L + 280, y);
  y += 14;
  doc.text(`Phone: ${trader && trader.tel ? trader.tel : '-'}`, PAGE_L, y);
  doc.text(`Date: ${fmtDate(auction.date)}`, PAGE_L + 280, y);
  y += 18;

  const cols = [
    { k: 'lot_no', label: 'Lot#',     x: PAGE_L,        w: 60,  align: 'left' },
    { k: 'qty',    label: 'Qty',      x: PAGE_L + 60,   w: 70,  align: 'right', fmt: fmtQty },
    { k: 'rate',   label: 'Rate',     x: PAGE_L + 130,  w: 60,  align: 'right', fmt: fmtAmt },
    { k: 'amount', label: 'Amount',   x: PAGE_L + 190,  w: 80,  align: 'right', fmt: fmtAmt },
    { k: 'refund', label: 'Discount', x: PAGE_L + 270,  w: 75,  align: 'right', fmt: fmtAmt },
    { k: 'tax',    label: 'GST',      x: PAGE_L + 345,  w: 70,  align: 'right', fmt: fmtAmt },
    { k: 'balance',label: 'Payable',  x: PAGE_L + 415,  w: 120, align: 'right', fmt: fmtAmt },
  ];
  doc.font('Helvetica-Bold').fontSize(9);
  doc.rect(PAGE_L, y, PAGE_W, 18).fillAndStroke('#f3f4f6', '#999').fillColor('#000');
  for (const c of cols) doc.text(c.label, c.x + 2, y + 5, { width: c.w - 4, align: c.align });
  y += 18;
  doc.font('Helvetica').fontSize(9).fillColor('#000');
  let tQty=0,tAmt=0,tDisc=0,tTax=0,tPay=0;
  for (const l of lots) {
    const tax = (Number(l.cgst)||0)+(Number(l.sgst)||0)+(Number(l.igst)||0);
    const row = { ...l, tax };
    tQty+=Number(l.qty)||0; tAmt+=Number(l.amount)||0; tDisc+=Number(l.refund)||0; tTax+=tax; tPay+=Number(l.balance)||0;
    if (y > 770) { doc.addPage(); y = 40; }
    for (const c of cols) {
      const v = c.fmt ? c.fmt(row[c.k]) : String(row[c.k] ?? '');
      doc.text(v, c.x + 2, y + 4, { width: c.w - 4, align: c.align });
    }
    y += 14;
    doc.moveTo(PAGE_L, y).lineTo(PAGE_R, y).strokeColor('#e5e7eb').lineWidth(0.5).stroke().strokeColor('#000');
  }
  doc.font('Helvetica-Bold').fontSize(10);
  doc.rect(PAGE_L, y, PAGE_W, 20).fillAndStroke('#f3f4f6', '#666').fillColor('#000');
  doc.text('TOTAL', PAGE_L + 2, y + 6);
  doc.text(fmtQty(tQty), PAGE_L + 62, y + 6, { width: 66, align: 'right' });
  doc.text(fmtAmt(tAmt), PAGE_L + 192, y + 6, { width: 76, align: 'right' });
  doc.text(fmtAmt(tDisc),PAGE_L + 272, y + 6, { width: 71, align: 'right' });
  doc.text(fmtAmt(tTax), PAGE_L + 347, y + 6, { width: 66, align: 'right' });
  doc.text(fmtAmt(tPay), PAGE_L + 417, y + 6, { width: 116,align: 'right' });
  y += 30;
  doc.font('Helvetica').fontSize(9).text(`Generated: ${new Date().toLocaleString('en-IN')}`, PAGE_L, y, { width: PAGE_W, align: 'right' });
  return tPay;
}

// Full per-seller statement (every payable lot).
app.get('/api/payments/pdf/:auctionId/:sellerName', requireView, (req, res) => {
  let doc, piped = false;
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const auctionId = req.params.auctionId;
    const sellerName = decodeURIComponent(req.params.sellerName);
    const auction = db.get('SELECT id FROM auctions WHERE id = ?', [auctionId]);
    if (!auction) return res.status(404).json({ error: 'Auction not found' });
    const PDFDocument = require('pdfkit');
    doc = new PDFDocument({ size: 'A4', margin: 30 });
    res.setHeader('Content-Type', 'application/pdf');
    const safeName = String(sellerName || '').replace(/[^\w]/g, '_').slice(0, 80) || 'seller';
    res.setHeader('Content-Disposition', `inline; filename="Payment_${safeName}_${auctionId}.pdf"`);
    doc.pipe(res); piped = true;
    res.on('close', () => { try { doc.destroy(); } catch(_){} });
    _renderPaymentStatement(doc, db, auctionId, sellerName, cfg);
    doc.end();
  } catch (e) {
    if (piped && doc) { try { doc.end(); } catch(_){} }
    else if (!res.headersSent) res.status(500).json({ error: e.message || 'PDF failed' });
  }
});

// Per-seller, lot-filtered statement. Body { auction_id, seller_name, lot_ids:[...] }
// Powers the Payments tab "partial payment" flow ("Print Payment for Selected").
app.post('/api/payments/pdf-lots', requireView, (req, res) => {
  let doc, piped = false;
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const auctionId = Number(req.body.auction_id);
    const sellerName = String(req.body.seller_name || '').trim();
    const lotIds = Array.isArray(req.body.lot_ids) ? req.body.lot_ids : [];
    if (!auctionId || !sellerName || !lotIds.length) {
      return res.status(400).json({ error: 'auction_id, seller_name and lot_ids[] are required' });
    }
    const auction = db.get('SELECT id FROM auctions WHERE id = ?', [auctionId]);
    if (!auction) return res.status(404).json({ error: 'Auction not found' });
    const PDFDocument = require('pdfkit');
    doc = new PDFDocument({ size: 'A4', margin: 30 });
    res.setHeader('Content-Type', 'application/pdf');
    const safeName = sellerName.replace(/[^\w]/g, '_').slice(0, 80) || 'seller';
    res.setHeader('Content-Disposition', `inline; filename="Payment_${safeName}_${auctionId}_partial.pdf"`);
    doc.pipe(res); piped = true;
    res.on('close', () => { try { doc.destroy(); } catch(_){} });
    _renderPaymentStatement(doc, db, auctionId, sellerName, cfg, lotIds);
    doc.end();
  } catch (e) {
    if (piped && doc) { try { doc.end(); } catch(_){} }
    else if (!res.headersSent) res.status(500).json({ error: e.message || 'PDF failed' });
  }
});

// Bulk: Body { auction_id, names: [...] } → one merged PDF, page-break
// per seller. Powers the Payments tab "Print Selected" button.
app.post('/api/payments/pdf-bulk', requireView, (req, res) => {
  let doc, piped = false;
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const auctionId = Number(req.body.auction_id);
    const names = Array.isArray(req.body.names) ? req.body.names : [];
    if (!auctionId || !names.length) return res.status(400).json({ error: 'auction_id and names[] required' });
    const auction = db.get('SELECT id FROM auctions WHERE id = ?', [auctionId]);
    if (!auction) return res.status(404).json({ error: 'Auction not found' });
    const PDFDocument = require('pdfkit');
    doc = new PDFDocument({ size: 'A4', margin: 30 });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="Payments_Batch_${names.length}.pdf"`);
    doc.pipe(res); piped = true;
    res.on('close', () => { try { doc.destroy(); } catch(_){} });
    names.forEach((nm, i) => {
      if (i > 0) doc.addPage();
      try { _renderPaymentStatement(doc, db, auctionId, nm, cfg); }
      catch (e) { try { doc.font('Helvetica').fontSize(10).text(`Error rendering ${nm}: ${e.message}`); } catch(_){} }
    });
    doc.end();
  } catch (e) {
    if (piped && doc) { try { doc.end(); } catch(_){} }
    else if (!res.headersSent) res.status(500).json({ error: e.message || 'PDF failed' });
  }
});

// ══════════════════════════════════════════════════════════════
// TDS RETURNS (TDSRETU.PRG)
// ══════════════════════════════════════════════════════════════
app.get('/api/tds-return', requireView, (req, res) => {
  const { from, to, orderBy } = req.query;
  if (!from || !to) return res.status(400).json({ error: 'from and to dates required' });
  const data = getTDSReturnData(getDb(), from, to, orderBy || 'invoice');
  res.json(data);
});

// ══════════════════════════════════════════════════════════════
// EXPORTS (EXP.PRG — all 11 types + TDS + Tally)
// ══════════════════════════════════════════════════════════════
app.get('/api/exports/:type/:auctionId', requireExport, async (req, res) => {
  const { type, auctionId } = req.params;
  const format = (req.query.format || 'xlsx').toLowerCase();

  if (format === 'pdf') {
    try {
      const db = getDb();
      const cfg = getSettingsFlat(db);
      const buffer = await exportAnyPdf(db, type, auctionId, cfg, { state: req.query.state });
      const niceName = (EXPORT_TYPES[type] && EXPORT_TYPES[type].name) || type;
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${niceName}_${auctionId}.pdf"`);
      return res.send(buffer);
    } catch (e) {
      return res.status(500).json({ error: e.message });
    }
  }

  const exportDef = EXPORT_TYPES[type];
  if (!exportDef) return res.status(400).json({ error: 'Unknown export type', available: Object.keys(EXPORT_TYPES) });

  try {
    const db = getDb();
    let buffer;
    // Optional seller-name filter — drives the "Export Selected" buttons
    // in the Payments tab. The Bank Payment + Payment XLSX exporters
    // read this through their 5th arg; every other exporter ignores
    // unknown opts, so this is a no-op for them. Accept BOTH shapes:
    //   ?names=A&names=B&names=C   → req.query.names = ['A','B','C']
    //   ?names=A,B,C               → req.query.names = 'A,B,C'  → split
    let rawNames = req.query.names;
    if (typeof rawNames === 'string') rawNames = rawNames.split(',');
    if (!Array.isArray(rawNames)) rawNames = [];
    const names = rawNames.map(s => String(s || '').trim()).filter(Boolean);
    const opts = names.length ? { names } : undefined;
    if (exportDef.needsCfg) {
      const cfg = getSettingsFlat(db);
      // Pass state too so exports that need both (e.g. Praman) can filter
      // by state without losing cfg context. Backward-compatible: existing
      // needsCfg exports that ignore the 4th/5th args are unaffected.
      buffer = await exportDef.fn(db, auctionId, cfg, req.query.state, opts);
    } else {
      buffer = await exportDef.fn(db, auctionId, req.query.state, opts);
    }
    // Per-export-type content-type/extension override (defaults to xlsx).
    // CSV exports like Praman use ext:'csv', mime:'text/csv'.
    const ext  = exportDef.ext  || 'xlsx';
    const mime = exportDef.mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    // Tag the download filename with "_selected" when a names filter is
    // active so a partial-export file is obviously a subset on disk and
    // doesn't get confused with the full export later.
    const suffix = opts ? '_selected' : '';
    res.setHeader('Content-Type', mime);
    res.setHeader('Content-Disposition', `attachment; filename="${exportDef.name}${suffix}_${auctionId}.${ext}"`);
    res.send(Buffer.from(buffer));
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// POST sibling of the GET export route — same response shape, but the
// "selected" filters live in the JSON body so:
//   (a) URL length isn't a constraint when many sellers are ticked,
//   (b) the per-seller lot-pick map ({ "Alice": ["12","15"] }) can be
//       sent as a normal nested object instead of being marshalled
//       through query-string bracket notation.
//
// Used by the Payments tab's "Export Bank Payment (Selected)" and
// "Export Payment XLSX (Selected)" buttons. The GET route stays the
// canonical path for the no-filter case (all sellers, all lots).
app.post('/api/exports/:type/:auctionId', requireExport, async (req, res) => {
  const { type, auctionId } = req.params;
  const format = String((req.body && req.body.format) || 'xlsx').toLowerCase();

  if (format === 'pdf') {
    // PDF generation on this route is intentionally not implemented —
    // the Payments-tab selected exports are XLSX-only. Adding PDF later
    // is straightforward: thread `opts` through exportAnyPdf the same
    // way the GET route does for ?state=...
    return res.status(400).json({ error: 'PDF format is not supported on the selected-export route. Use the GET endpoint for a full PDF.' });
  }

  const exportDef = EXPORT_TYPES[type];
  if (!exportDef) return res.status(400).json({ error: 'Unknown export type', available: Object.keys(EXPORT_TYPES) });

  try {
    const db = getDb();
    // Normalise the body. names → array of trimmed non-empty strings.
    // lots → object mapping seller-name → array of lot_no strings.
    // excludeLots → same shape; lots that already shipped in earlier
    //               exports and must not be included again. Lets the
    //               client re-export a seller's remaining lots without
    //               accidentally double-paying the ones already paid.
    const body = req.body || {};
    const names = Array.isArray(body.names)
      ? body.names.map(s => String(s || '').trim()).filter(Boolean)
      : [];
    const lots        = (body.lots        && typeof body.lots        === 'object' && !Array.isArray(body.lots))        ? body.lots        : null;
    const excludeLots = (body.excludeLots && typeof body.excludeLots === 'object' && !Array.isArray(body.excludeLots)) ? body.excludeLots : null;
    const opts = (names.length || lots || excludeLots) ? {} : undefined;
    if (opts) {
      if (names.length) opts.names = names;
      const cleanMap = (src) => {
        const cleaned = {};
        for (const k of Object.keys(src)) {
          const arr = Array.isArray(src[k]) ? src[k].map(v => String(v || '').trim()).filter(Boolean) : [];
          if (arr.length) cleaned[k] = arr;
        }
        return Object.keys(cleaned).length ? cleaned : null;
      };
      if (lots) {
        const c = cleanMap(lots);
        if (c) opts.lots = c;
      }
      if (excludeLots) {
        const c = cleanMap(excludeLots);
        if (c) opts.excludeLots = c;
      }
    }
    let buffer;
    if (exportDef.needsCfg) {
      const cfg = getSettingsFlat(db);
      buffer = await exportDef.fn(db, auctionId, cfg, body.state || null, opts);
    } else {
      buffer = await exportDef.fn(db, auctionId, body.state || null, opts);
    }
    const ext  = exportDef.ext  || 'xlsx';
    const mime = exportDef.mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const suffix = opts ? '_selected' : '';
    res.setHeader('Content-Type', mime);
    res.setHeader('Content-Disposition', `attachment; filename="${exportDef.name}${suffix}_${auctionId}.${ext}"`);
    res.send(Buffer.from(buffer));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// TDS export (supports ?format=pdf)
app.get('/api/exports/tds-return', requireExport, async (req, res) => {
  const { from, to } = req.query;
  const format = (req.query.format || 'xlsx').toLowerCase();
  if (!from || !to) return res.status(400).json({ error: 'from/to required' });
  if (format === 'pdf') {
    const buffer = await exportAnyPdf(getDb(), 'tds_return', null, null, { from, to });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="TDSReturn.pdf"');
    return res.send(buffer);
  }
  const { exportTDSReturn } = require('./exports');
  const buffer = await exportTDSReturn(getDb(), from, to);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="TDSReturn.xlsx"');
  res.send(Buffer.from(buffer));
});

// ══════════════════════════════════════════════════════════════
// LORRY REPORTS (Lot Slip Code / Truck List / Buyer Lot Lorry)
// ══════════════════════════════════════════════════════════════
app.get('/api/lorry-reports/:type/:auctionId', requireExport, async (req, res) => {
  const { type, auctionId } = req.params;
  const format = (req.query.format || 'xlsx').toLowerCase();
  const def = LORRY_REPORTS[type];
  if (!def) return res.status(400).json({ error: 'Unknown lorry report', available: Object.keys(LORRY_REPORTS) });
  try {
    const db = getDb();
    if (format === 'pdf') {
      const buf = await def.pdf(db, auctionId);
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${def.name}_${auctionId}.pdf"`);
      return res.send(buf);
    }
    const buf = await def.xlsx(db, auctionId);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${def.name}_${auctionId}.xlsx"`);
    return res.send(Buffer.from(buf));
  } catch (e) {
    console.error('lorry-reports error:', e);
    return res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// SPICE BOARD REPORTS — statutory cardamom-auction reports
// ══════════════════════════════════════════════════════════════
// All four report types share one dispatcher:
//   buyers_statement / form_d / form_c → json + xlsx + pdf
//   eauction_csv                       → csv only (Spices Board portal)
//
// The data layer is keyed off `auctionId` with optional branch / seller /
// buyer / place overrides — picked up by getReportContext() in
// spice-board-reports.js.

function _sbBuildOpts(req) {
  return {
    auctionId: req.query.auctionId || req.params.auctionId,
    branch:    req.query.branch    || null,
    sellerId:  req.query.sellerId  || null,
    buyerCode: req.query.buyerCode || null,
    dateFrom:  req.query.dateFrom  || null,
    dateTo:    req.query.dateTo    || null,
    place:     req.query.place     || null,
  };
}

// Filter dropdowns for the Spice Board tab — branches, sellers, buyers
// seen in the picked auction.
app.get('/api/spice-board-reports/filters', requireAuth, (req, res) => {
  try {
    const aid = req.query.auctionId;
    if (!aid) return res.status(400).json({ error: 'auctionId is required' });
    const db = getDb();
    res.json(getSpiceBoardFilters(db, aid));
  } catch (e) {
    console.error('spice-board-reports filters error:', e);
    res.status(500).json({ error: e.message });
  }
});

// JSON preview — returns the full report shape rendered by the front-end.
// CSV-only reports (eauction_csv) have no json() so we 400 there.
app.get('/api/spice-board-reports/:type/data', requireAuth, (req, res) => {
  const { type } = req.params;
  const def = SPICE_BOARD_REPORTS[type];
  if (!def) return res.status(400).json({ error: 'Unknown report type', available: Object.keys(SPICE_BOARD_REPORTS) });
  if (!def.json) return res.status(400).json({ error: 'No JSON preview for ' + type });
  try {
    const db = getDb();
    const out = def.json(db, _sbBuildOpts(req));
    res.json(out);
  } catch (e) {
    console.error('spice-board-reports data error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Export — xlsx | pdf | csv depending on the report. `inline=1` keeps the
// browser in-tab so the print dialog can fire (used by sbPrint() on the
// front end).
app.get('/api/spice-board-reports/:type/export', requireExport, async (req, res) => {
  const { type } = req.params;
  const def = SPICE_BOARD_REPORTS[type];
  if (!def) return res.status(400).json({ error: 'Unknown report type', available: Object.keys(SPICE_BOARD_REPORTS) });
  const format = String(req.query.format || (def.csv ? 'csv' : 'xlsx')).toLowerCase();
  const inline = String(req.query.inline || '') === '1';
  const opts = _sbBuildOpts(req);
  if (!opts.auctionId) return res.status(400).json({ error: 'auctionId is required' });
  try {
    const db = getDb();
    const name = def.name || type;
    const dispo = (mime, ext) => {
      res.setHeader('Content-Type', mime);
      res.setHeader('Content-Disposition',
        `${inline ? 'inline' : 'attachment'}; filename="${name}_${opts.auctionId}.${ext}"`);
    };
    if (format === 'csv') {
      if (!def.csv) return res.status(400).json({ error: 'CSV not supported for ' + type });
      const buf = await def.csv(db, opts);
      dispo('text/csv; charset=utf-8', 'csv');
      return res.send(buf);
    }
    if (format === 'pdf') {
      if (!def.pdf) return res.status(400).json({ error: 'PDF not supported for ' + type });
      const buf = await def.pdf(db, opts);
      dispo('application/pdf', 'pdf');
      return res.send(buf);
    }
    // default: xlsx
    if (!def.xlsx) return res.status(400).json({ error: 'XLSX not supported for ' + type });
    const buf = await def.xlsx(db, opts);
    dispo('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'xlsx');
    return res.send(Buffer.from(buf));
  } catch (e) {
    console.error('spice-board-reports export error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// DBF EXPORTS (FoxPro-compatible format)
// ══════════════════════════════════════════════════════════════

// List all available DBF export types with labels
app.get('/api/dbf-exports/list', requireExport, (req, res) => {
  const list = {};
  for (const [key, def] of Object.entries(DBF_EXPORTS)) {
    list[key] = {
      label: def.label,
      name: def.name,
      needsAuction: !!def.needsAuction,
      needsDateRange: !!def.needsDateRange,
    };
  }
  res.json(list);
});

// Generic DBF export endpoint
app.get('/api/dbf-exports/:type', requireExport, async (req, res) => {
  const { type } = req.params;
  const def = DBF_EXPORTS[type];
  if (!def) return res.status(400).json({ error: 'Unknown DBF export type', available: Object.keys(DBF_EXPORTS) });

  try {
    const db = getDb();
    let buffer;

    if (def.needsAuction) {
      const { auctionId } = req.query;
      if (!auctionId) return res.status(400).json({ error: 'auctionId query parameter required' });
      buffer = await def.fn(db, auctionId);
    } else if (def.needsDateRange) {
      const { from, to, ano } = req.query;
      const filters = {};
      if (ano) filters.ano = ano;
      if (from && to) { filters.from = from; filters.to = to; }
      buffer = await def.fn(db, filters);
    } else {
      buffer = await def.fn(db);
    }

    // Build filename: LOTS_1.dbf, INV_2026-04-01_to_2026-04-30.dbf, NAM.dbf
    let filename = def.name;
    if (def.needsAuction && req.query.auctionId) filename += `_${req.query.auctionId}`;
    if (def.needsDateRange && req.query.from) filename += `_${req.query.from}_to_${req.query.to}`;
    filename += '.dbf';

    res.setHeader('Content-Type', 'application/x-dbase');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch(e) {
    console.error('DBF export error:', e);
    res.status(500).json({ error: 'DBF export failed: ' + e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// TO TALLY — XML exports for Tally accounting software
// ══════════════════════════════════════════════════════════════

// Definitions of available Tally exports — keep in sync with frontend.
//
// `company` resolves which Tally company name goes in the XML's
// <SVCURRENTCOMPANY> tag. For now this routing only differs across the
// 3 party-ledger types:
//   • Sales Party Ledgers → ISP (sales-side parties = ISP customers)
//   • RD / URD Party Ledgers → ASP (purchase-side parties = ASP suppliers)
// All other exports (vouchers, the all-in-one ledger master) currently
// import into the ISP company. We can split vouchers later if dad asks.
const TALLY_EXPORTS = {
  ledger_sales:        { label: 'Sales Party Ledgers',                              name: 'SalesPartyLedgers',  builder: buildSalesPartyLedgerRows, generator: generLedgerXML, isLedger: true, company: 'isp' },
  ledger_rd_purchase:  { label: 'RD Purchase Party Ledgers',                        name: 'RDPartyLedgers',     builder: buildRDPartyLedgerRows,    generator: generLedgerXML, isLedger: true, company: 'asp' },
  ledger_urd_purchase: { label: 'URD Purchase Party Ledgers (Agriculturist)',       name: 'URDPartyLedgers',    builder: buildURDPartyLedgerRows,   generator: generLedgerXML, isLedger: true, company: 'asp' },
  ledger:              { label: 'All Ledger Masters (parties + tax + sales + purchase)', name: 'AllLedgers',  builder: buildLedgerRows,           generator: generLedgerXML, isLedger: true, company: 'isp' },
  // ── Sales Vouchers — split into two purpose-built exports ────
  // sales_isp = ISP→outside-customer sales (full e-way bill, dispatch
  //   from sister, BASICORDERREF to matching ASP voucher).
  // sales_asp = ASP→ISP internal transfers (lean format, no e-way
  //   bill, customer is always ISP, lot rates from asp_prate/asp_puramt).
  // The legacy `sales` key is kept as an alias for sales_isp so any old
  // bookmarks / API callers don't break; new UI buttons use the split keys.
  sales_isp:           { label: 'Sales Vouchers — ISP',                             name: 'SalesISP',           builder: buildSalesIspRows,         generator: generSalesIspXML,     company: 'isp' },
  sales_asp:           { label: 'Sales Vouchers — ASP',                             name: 'SalesASP',           builder: buildSalesAspRows,         generator: generSalesAspXML,     company: 'asp' },
  sales:               { label: 'Sales Vouchers (legacy alias for ISP)',            name: 'Sales',              builder: buildSalesIspRows,         generator: generSalesIspXML,     company: 'isp' },
  // ISP Purchase = the buyer-side mirror of an ASP→ISP transfer. Each
  // sales_asp row produces one isp_purchase voucher into ISP's books with
  // the same VOUCHERNUMBER (e.g. ASP/I-61/26-27) for cross-reference. We
  // re-use buildSalesAspRows directly since the row shape is identical.
  isp_purchase:        { label: 'ISP Purchase Vouchers (mirror of ASP→ISP)',        name: 'ISPPurchase',        builder: buildSalesAspRows,         generator: generIspPurchaseXML,  company: 'isp' },
  rd_purchase:         { label: 'RD Purchase Vouchers',                             name: 'RDPurchase',         builder: buildRDPurchaseRows,       generator: generRDPurchaseXML,   company: 'asp' },
  urd_purchase:        { label: 'URD Purchase Vouchers (Agriculturist)',            name: 'URDPurchase',        builder: buildURDPurchaseRows,      generator: generURDPurchaseXML,  company: 'asp' },
  debit_note:          { label: 'Debit Notes (Discount)',                           name: 'DebitNote',          builder: buildDebitNoteRows,        generator: generDebitNoteXML,    company: 'isp' },
};

// Resolve the Tally company name for a given export type.
// 'isp' → tally_company_name; 'asp' → tally_asp_company_name (falls
// back to ISP if the ASP name is blank, but logs a warning so misconfig
// is visible — silently falling back has caused confusion when a user
// sees ISP in <SVCURRENTCOMPANY> but expected ASP).
function resolveTallyCompanyName(cfg, target) {
  const isp = (cfg.tally_company_name || '').trim();
  const asp = (cfg.tally_asp_company_name || '').trim();
  if (target === 'asp') {
    if (!asp) {
      console.warn('[tally] tally_asp_company_name is empty — falling back to ISP company name. Set it via Settings → To Tally → "ASP Tally Company Name".');
    }
    return asp || isp;
  }
  return isp;
}

// Map a single-party `kind` (sales|rd_purchase|urd_purchase) to its
// dedicated builder + which Tally company its ledger belongs to.
const PARTY_LEDGER_BUILDERS = {
  sales:        { builder: buildSalesPartyLedgerRows, company: 'isp' },
  rd_purchase:  { builder: buildRDPartyLedgerRows,    company: 'asp' },
  urd_purchase: { builder: buildURDPartyLedgerRows,   company: 'asp' },
};

// List endpoint — used by the To Tally tab to render export buttons
app.get('/api/tally/list', requireExport, (req, res) => {
  const list = {};
  for (const [key, def] of Object.entries(TALLY_EXPORTS)) {
    list[key] = { label: def.label, name: def.name };
  }
  res.json(list);
});

// Preview endpoint — returns row counts so the user knows how many vouchers
// will be in the XML before downloading
app.get('/api/tally/preview/:type/:auctionId', requireExport, (req, res) => {
  const { type, auctionId } = req.params;
  const def = TALLY_EXPORTS[type];
  if (!def) return res.status(400).json({ error: 'Unknown Tally export', available: Object.keys(TALLY_EXPORTS) });
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const rows = def.builder(db, auctionId, cfg);
    const targetCompany = resolveTallyCompanyName(cfg, def.company);
    if (def.isLedger) {
      // Ledger rows have a different shape — count by kind
      const byKind = rows.reduce((acc, r) => {
        const k = r.kind || 'other';
        acc[k] = (acc[k] || 0) + 1;
        return acc;
      }, {});
      return res.json({
        type, auctionId,
        ledgerCount: rows.length,
        byKind,
        targetCompany,
        sample: rows.slice(0, 6).map(r => ({ kind: r.kind, name: r.name, parent: r.parent, gstin: r.gstin || '' })),
      });
    }
    const totalLots = rows.reduce((s, r) => s + (Array.isArray(r.lots) ? r.lots.length : 0), 0);
    res.json({
      type, auctionId,
      voucherCount: rows.length,
      lotCount: totalLots,
      targetCompany,
      sample: rows.slice(0, 3).map(r => ({
        ano: r.ano, date: r.date, name: r.partyName || r.name,
        voucher: r.voucherNum || r.invo,
        amount: r.total,
      })),
    });
  } catch (e) {
    console.error('tally preview error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// E-WAY BILL DISTANCE — route table + per-invoice override
// ══════════════════════════════════════════════════════════════
// Two storage paths feed the <DISTANCE> field on ISP sales vouchers:
//
//   1. invoices.distance_km — per-invoice override. Wins when set.
//   2. route_distances[(from_pin, to_pin)] — saved route value, applied
//      to every invoice between the same two PINs.
//
// Workflow: user clicks NIC, looks up dispatch→consignee distance on
// the portal, types it in, clicks Save. We write to route_distances —
// every other invoice between the same two PINs (this auction and all
// future ones) auto-resolves.

// Resolve the configured dispatch PIN with the same fallback chain the
// voucher generator uses. Used for normalising route lookups.
function getDispatchPin(db) {
  const cfg = require('./company-config').getSettingsFlat(db);
  return String(
    cfg.tally_dispatch_pin || cfg.s_pin || '685553'
  ).trim();
}

// Normalise a (from, to) pair so A↔B share a single route_distances row.
// Always returns the lexicographically smaller PIN first.
function normalizeRouteKey(fromPin, toPin) {
  const a = String(fromPin || '').trim();
  const b = String(toPin || '').trim();
  return a < b ? [a, b] : [b, a];
}

// List ISP invoices for an auction with their resolved distance + source
// tag. The UI uses this to render the table — `km` is the value to
// display, `source` tells the user where it came from ('manual' = per-
// invoice override, 'route' = looked up by PIN pair, 'none' = blank).
app.get('/api/invoices/distances/:auctionId', requireView, (req, res) => {
  try {
    const db = getDb();
    const dispatchPin = getDispatchPin(db);
    const rows = db.all(
      `SELECT i.id, i.ano, i.invo, i.buyer, i.buyer1, i.gstin, i.state,
              b.pin AS buyer_pin, b.pla AS buyer_pla,
              i.distance_km
       FROM invoices i
       LEFT JOIN buyers b ON b.buyer = i.buyer
       WHERE i.auction_id = ? AND UPPER(COALESCE(i.state,'')) = 'TAMIL NADU'
       ORDER BY CAST(i.invo AS INTEGER), i.id`,
      [req.params.auctionId]
    );

    // Pre-fetch all route distances for this dispatch PIN — one query
    // instead of N. The set is small (a few dozen routes max) so it
    // fits comfortably in memory.
    const routes = {};
    try {
      const allRoutes = db.all(
        `SELECT from_pin, to_pin, km FROM route_distances
         WHERE from_pin = ? OR to_pin = ?`,
        [dispatchPin, dispatchPin]
      );
      for (const r of allRoutes) {
        // The "other PIN" — whichever side isn't the dispatch PIN
        const other = r.from_pin === dispatchPin ? r.to_pin : r.from_pin;
        routes[other] = r.km;
      }
    } catch (e) { /* table may not exist on very old DBs */ }

    // Annotate each row with resolved distance + source
    const enriched = rows.map(r => {
      let km = null, source = 'none';
      if (r.distance_km != null) {
        km = r.distance_km;
        source = 'manual';
      } else if (r.buyer_pin && routes[String(r.buyer_pin).trim()] != null) {
        km = routes[String(r.buyer_pin).trim()];
        source = 'route';
      }
      return { ...r, resolved_km: km, distance_source: source };
    });

    res.json({
      count: enriched.length,
      dispatchPin,
      invoices: enriched,
    });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Save a route distance. Body: { from_pin, to_pin, km }. The pair gets
// normalised (smaller PIN first) before write, so subsequent lookups
// find it regardless of direction. Empty/null `km` deletes the row.
//
// Returns: how many ISP invoices now resolve via this route (so the
// UI can show "applied to N invoices" feedback).
app.put('/api/route-distances', requireExport, (req, res) => {
  const { from_pin, to_pin, km } = req.body || {};
  if (!from_pin || !to_pin) {
    return res.status(400).json({ error: 'from_pin and to_pin required' });
  }
  if (!/^\d{6}$/.test(String(from_pin).trim()) || !/^\d{6}$/.test(String(to_pin).trim())) {
    return res.status(400).json({ error: 'PINs must be 6-digit strings' });
  }
  const [k1, k2] = normalizeRouteKey(from_pin, to_pin);

  // Empty km = delete
  if (km === '' || km == null) {
    try {
      const r = getDb().run(
        'DELETE FROM route_distances WHERE from_pin = ? AND to_pin = ?',
        [k1, k2]
      );
      return res.json({ ok: true, deleted: r.changes > 0, from_pin: k1, to_pin: k2 });
    } catch (e) { return res.status(500).json({ error: e.message }); }
  }

  const v = Math.round(Number(km));
  if (!isFinite(v) || v < 0 || v > 5000) {
    return res.status(400).json({ error: 'km must be between 0 and 5000' });
  }

  try {
    const db = getDb();
    db.run(
      `INSERT INTO route_distances (from_pin, to_pin, km, updated_at)
       VALUES (?, ?, ?, datetime('now','localtime'))
       ON CONFLICT(from_pin, to_pin) DO UPDATE
         SET km = excluded.km, updated_at = excluded.updated_at`,
      [k1, k2, v]
    );

    // Saving a route is the user's signal that "this distance applies to
    // every invoice between these PINs." Clear any legacy per-invoice
    // overrides on matching invoices so the route value actually wins —
    // otherwise leftover invoices.distance_km values from earlier saves
    // would shadow the route forever. (Per-invoice overrides have higher
    // priority by design; if we ever add a UI to set a true per-invoice
    // override, this clearing step would need to be opt-in.)
    const dispatchPin = getDispatchPin(db);
    const otherPin = k1 === dispatchPin ? k2 : (k2 === dispatchPin ? k1 : null);
    let clearedOverrides = 0;
    if (otherPin) {
      const r = db.run(
        `UPDATE invoices SET distance_km = NULL
         WHERE id IN (
           SELECT i.id FROM invoices i
           LEFT JOIN buyers b ON b.buyer = i.buyer
           WHERE UPPER(COALESCE(i.state,'')) = 'TAMIL NADU'
             AND b.pin = ?
             AND i.distance_km IS NOT NULL
         )`,
        [otherPin]
      );
      clearedOverrides = r.changes || 0;
    }

    // How many invoices now resolve via this route? Now that we cleared
    // the legacy overrides, every invoice with the matching buyer PIN
    // counts (not just the ones that were already NULL).
    let appliedCount = 0;
    if (otherPin) {
      const r = db.get(
        `SELECT COUNT(*) AS n FROM invoices i
         LEFT JOIN buyers b ON b.buyer = i.buyer
         WHERE UPPER(COALESCE(i.state,'')) = 'TAMIL NADU'
           AND b.pin = ?`,
        [otherPin]
      );
      appliedCount = r ? r.n : 0;
    }

    res.json({ ok: true, from_pin: k1, to_pin: k2, km: v, appliedCount, clearedOverrides });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Bulk-clear all per-invoice distance overrides. Used to wipe legacy
// invoices.distance_km values from before the route-table refactor —
// after running this, every ISP invoice resolves via the route table
// (or stays blank if no route exists). Confirmed via the UI before
// hitting this; no body needed.
//
// Path is under /api/distance-overrides (not /api/invoices/distance-
// overrides) to avoid Express matching it against the earlier-defined
// app.delete('/api/invoices/:id') route, which treats 'distance-
// overrides' as an :id and returns 'Invoice not found'.
app.delete('/api/distance-overrides', requireExport, (req, res) => {
  try {
    const r = getDb().run(
      `UPDATE invoices SET distance_km = NULL
       WHERE distance_km IS NOT NULL
         AND UPPER(COALESCE(state,'')) = 'TAMIL NADU'`
    );
    res.json({ ok: true, cleared: r.changes });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Per-invoice override (kept in the API even though no UI button wires
// to it directly — useful for one-off exceptions where a single invoice
// needs a different distance than the route. Set null to clear.)
app.put('/api/invoices/:id/distance', requireExport, (req, res) => {
  const id = Number(req.params.id);
  const { distance_km } = req.body || {};
  let v = null;
  if (distance_km !== '' && distance_km != null) {
    v = Math.round(Number(distance_km));
    if (!isFinite(v) || v < 0 || v > 5000) {
      return res.status(400).json({ error: 'distance_km must be between 0 and 5000 km, or null to clear' });
    }
  }
  try {
    const r = getDb().run('UPDATE invoices SET distance_km = ? WHERE id = ?', [v, id]);
    if (r.changes === 0) return res.status(404).json({ error: 'Invoice not found' });
    res.json({ ok: true, id, distance_km: v });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Party listing for an auction — used by the single-party picker UI.
// Returns every distinct party (buyer/RD/URD) with the kind it would be
// exported under so the frontend can group and filter.
app.get('/api/tally/parties/:auctionId', requireExport, (req, res) => {
  const { auctionId } = req.params;
  try {
    const db = getDb();
    const parties = listAuctionParties(db, auctionId);
    const byKind = parties.reduce((acc, p) => {
      acc[p.kind] = (acc[p.kind] || 0) + 1;
      return acc;
    }, {});
    res.json({ auctionId, total: parties.length, byKind, parties });
  } catch (e) {
    console.error('tally parties error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Single-party ledger XML — emits exactly one ledger for the named party.
// kind: 'sales'|'rd_purchase'|'urd_purchase'  (matches party's source)
// Sales parties go into the ISP Tally company; RD/URD parties go into ASP.
app.get('/api/tally/party-ledger/:kind/:auctionId', requireExport, (req, res) => {
  const { kind, auctionId } = req.params;
  const partyName = req.query.name;
  if (!partyName) return res.status(400).json({ error: 'Missing ?name=<party name>' });
  const partyDef = PARTY_LEDGER_BUILDERS[kind];
  if (!partyDef) return res.status(400).json({ error: 'Unknown party kind', available: Object.keys(PARTY_LEDGER_BUILDERS) });
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const rows = partyDef.builder(db, auctionId, cfg, { partyName });
    if (rows.length === 0) {
      return res.status(404).json({ error: `Party "${partyName}" not found in ${kind} for auction ${auctionId}` });
    }
    const xml = generLedgerXML(rows, cfg, { companyName: resolveTallyCompanyName(cfg, partyDef.company) });
    const safeName = String(partyName).replace(/[^a-zA-Z0-9]/g, '_').slice(0, 40);
    const filename = `Tally_PartyLedger_${kind}_${safeName}_${auctionId}.xml`;
    res.setHeader('Content-Type', 'application/xml; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(xml);
  } catch (e) {
    console.error('tally party-ledger error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Single-party voucher XML — emits exactly one voucher for the named
// party, for one of the voucher types: sales_isp, sales_asp, rd_purchase,
// urd_purchase, debit_note. Useful when a single buyer/dealer needs a
// voucher import in isolation (e.g. you missed one and don't want to
// re-import the whole auction).
//
// We reuse the existing TALLY_EXPORTS builders (which produce rows for
// the entire auction) and filter the rows by party name in-memory. The
// "party name" field varies per voucher type:
//   sales_isp     → row.partyName     (= invoices.buyer1, the buyer)
//   sales_asp     → row.buyerName     (= the downstream ISP-side buyer)
//   isp_purchase  → row.buyerName     (= same source as sales_asp; the
//                                       voucher's "party" is always ASP,
//                                       so we filter by the downstream
//                                       buyer to let users pick a single
//                                       transfer voucher)
//   rd_purchase   → row.name          (= purchases.name, the dealer)
//   urd_purchase  → row.name          (= bills.name, the agriculturist)
//   debit_note    → row.partyName     (= the discount-paying supplier)
const VOUCHER_PARTY_KEY = {
  sales_isp:    (r) => r.partyName || '',
  sales_asp:    (r) => r.buyerName || r.buyer || '',
  isp_purchase: (r) => r.buyerName || r.buyer || '',
  rd_purchase:  (r) => r.name || '',
  urd_purchase: (r) => r.name || '',
  debit_note:   (r) => r.partyName || r.name || '',
  // Legacy alias still works
  sales:        (r) => r.partyName || '',
};

app.get('/api/tally/party-voucher/:type/:auctionId', requireExport, (req, res) => {
  const { type, auctionId } = req.params;
  const partyName = req.query.name;
  if (!partyName) return res.status(400).json({ error: 'Missing ?name=<party name>' });
  const def = TALLY_EXPORTS[type];
  const keyFn = VOUCHER_PARTY_KEY[type];
  if (!def || !keyFn || def.isLedger) {
    return res.status(400).json({
      error: 'Unknown or unsupported voucher type for single-party export',
      supported: Object.keys(VOUCHER_PARTY_KEY),
    });
  }
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const allRows = def.builder(db, auctionId, cfg);
    // Case-insensitive exact match on the party name; fall back to
    // contains-match if no exact hit (handles minor whitespace/case
    // differences between the picker label and the underlying data).
    const target = String(partyName).trim().toUpperCase();
    let rows = allRows.filter(r => String(keyFn(r) || '').trim().toUpperCase() === target);
    if (rows.length === 0) {
      rows = allRows.filter(r => String(keyFn(r) || '').toUpperCase().includes(target));
    }
    if (rows.length === 0) {
      return res.status(404).json({
        error: `No ${def.label} found for "${partyName}" in auction ${auctionId}`,
        availableParties: [...new Set(allRows.map(keyFn).filter(Boolean))].slice(0, 20),
      });
    }
    const xml = def.generator(rows, cfg, { companyName: resolveTallyCompanyName(cfg, def.company) });
    const safeName = String(partyName).replace(/[^a-zA-Z0-9]/g, '_').slice(0, 40);
    const filename = `${def.name}_${safeName}_${auctionId}.xml`;
    res.setHeader('Content-Type', 'application/xml; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(xml);
  } catch (e) {
    console.error('tally party-voucher error:', e);
    res.status(500).json({ error: e.message });
  }
});

// XML download endpoint — the main thing
app.get('/api/tally/export/:type/:auctionId', requireExport, (req, res) => {
  const { type, auctionId } = req.params;
  const def = TALLY_EXPORTS[type];
  if (!def) return res.status(400).json({ error: 'Unknown Tally export', available: Object.keys(TALLY_EXPORTS) });
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const rows = def.builder(db, auctionId, cfg);
    if (rows.length === 0) {
      const what = def.isLedger ? def.label.toLowerCase() : `${def.label.toLowerCase()}`;
      return res.status(404).json({ error: `No ${what} found for auction ${auctionId}` });
    }
    const xml = def.generator(rows, cfg, { companyName: resolveTallyCompanyName(cfg, def.company) });
    const filename = `${def.name}_${auctionId}.xml`;
    res.setHeader('Content-Type', 'application/xml; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(xml);
  } catch (e) {
    console.error('tally export error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// DEBIT-NOTE PRINT TEMPLATE (per-buyer credit/debit note PDF)
// ══════════════════════════════════════════════════════════════
// Produces the buyer-grouped discount-note PDF in the layout the user
// supplied as the reference template. One page (or several, with C/D +
// B/F continuation lines) per buyer in the auction.
//
// Source data:
//   • debit_notes WHERE date = auction.date — gives buyer name + total
//     discount per buyer
//   • lots WHERE buyer1 = name AND auction_id = N — gives the lot rows
//     to render in the table
//   • buyers WHERE name = … — gives buyer GSTIN + address for the
//     header
//
// Per-lot discount split: the debit_notes row stores the summary total
// only, so we proportionally allocate the total across the buyer's lots
// (lot_disc = round(lot.amount × total / total_value)). Drift goes to
// the largest-amount lot so the column sum exactly equals the stored
// note total.
app.get('/api/tally/debit-note-print/:auctionId', requireExport, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const auctionId = req.params.auctionId;
    const auc = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]);
    if (!auc) return res.status(404).json({ error: 'Auction not found' });

    // Pull all debit notes for this auction. We match on `ano` (the
    // auction number) — the same column the auction itself uses — rather
    // than `date`, because legacy debit notes have `date` set to the
    // creation timestamp instead of the auction date. Going forward both
    // routes (ano + date) would work; matching by ano survives the
    // legacy data without needing a migration.
    const notes = db.all(
      'SELECT * FROM debit_notes WHERE ano = ? ORDER BY name',
      [auc.ano]
    );
    if (!notes.length) return res.status(404).json({ error: `No debit notes found for auction ${auctionId}` });

    // For each note, fetch the buyer's lots in this auction. Match on
    // buyer1 (proper buyer name) first; the short `buyer` code is also
    // tried in case `debit_notes.name` was inserted with the code form.
    // Filter amount > 0 to drop refunded/zero lots.
    const lotsStmt = db.prepare(`
      SELECT lot_no, qty, price, amount
      FROM lots
      WHERE auction_id = ? AND amount > 0
        AND (buyer1 = ? OR buyer = ?)
      ORDER BY CAST(lot_no AS INTEGER), lot_no
    `);
    // Buyers table: `buyer1` = long company name, `buyer` = short code.
    // No `name` column exists. Look up by either, since debit_notes.name
    // could hold either form historically.
    const buyerStmt = db.prepare(
      'SELECT buyer1, add1, add2, pla, gstin FROM buyers WHERE buyer1 = ? OR buyer = ? LIMIT 1'
    );

    const buyers = notes.map(n => {
      const lots = lotsStmt.all(auctionId, n.name, n.name);
      const b = buyerStmt.get(n.name, n.name) || {};
      // Compose a single-line address from add1 / add2 / pla, dropping empties.
      const addrParts = [b.add1, b.add2, b.pla].map(s => (s || '').trim()).filter(Boolean);
      return {
        name: n.name,
        address: addrParts.join(', '),
        gstin: b.gstin || '',
        noteNo: n.note_no || '',
        date: n.date,
        totalDiscount: Number(n.amount || 0),
        lots,
      };
    }).filter(b => b.lots.length > 0);  // skip buyers with no lots — no rows to render

    if (!buyers.length) return res.status(404).json({ error: 'No matching lots found for any debit note in this auction' });

    // Build the our-company cfg used in every page header. Keys here
    // line up with what generateDebitNoteBatchPDF expects.
    const ourCfg = {
      tally_company_name: cfg.tally_company_name || cfg.short_name || cfg.company || '',
      tally_dispatch_place: cfg.tally_dispatch_place || cfg.place || '',
      place: cfg.place || '',
      gstin: cfg.gstin || '',
      state: cfg.state || cfg.business_state || '',
      tally_season: cfg.tally_season || cfg.season_code || '',
      tally_hsn_cardamom: cfg.tally_hsn_cardamom || '09083120',
    };

    const pdf = await generateDebitNoteBatchPDF(buyers, ourCfg);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="DebitNotes_${auctionId}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('debit-note-print error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ── Crop Receipt PDF ─────────────────────────────────────────
app.get('/api/receipt/:lotId', requireView, async (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const lot = db.get('SELECT l.*, a.ano FROM lots l JOIN auctions a ON a.id=l.auction_id WHERE l.id=?', [req.params.lotId]);
  if (!lot) return res.status(404).json({ error: 'Lot not found' });
  const pdf = await generateCropReceiptPDF(lot, cfg);
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `inline; filename="Receipt_${lot.lot_no}.pdf"`);
  res.send(pdf);
});

// ══════════════════════════════════════════════════════════════
// SUMMARY STATS
// ══════════════════════════════════════════════════════════════
app.get('/api/stats', requireView, (req, res) => {
  const db = getDb();
  // Mode filter — every auction-scoped count and aggregate gets gated.
  // Traders + buyers stay mode-agnostic (master data isn't auction-scoped).
  // Subquery prefix is rebuilt per table to keep the SQL readable.
  const mwA = modeWhereClause(db, 'auctions.mode');
  const mwLots  = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=lots.auction_id)');
  const mwInv   = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=invoices.auction_id)');
  const mwPur   = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=purchases.auction_id)');
  const mwBill  = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=bills.auction_id)');
  const mwDN    = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=debit_notes.auction_id)');
  // Each `_w(mwX)` turns a modeWhereClause into a complete WHERE so we
  // can prepend it without dragging in `WHERE 1=1`. Empty when no mode.
  const _w = (mw) => mw.sql ? 'WHERE 1=1' + mw.sql : '';

  // Counts — auction-scoped tables get filtered, master data doesn't.
  const counts = {
    traders:    (db.get('SELECT COUNT(*) as c FROM traders') || {}).c || 0,
    buyers:     (db.get('SELECT COUNT(*) as c FROM buyers') || {}).c || 0,
    auctions:   (db.get(`SELECT COUNT(*) as c FROM auctions ${_w(mwA)}`,    mwA.params)   || {}).c || 0,
    lots:       (db.get(`SELECT COUNT(*) as c FROM lots ${_w(mwLots)}`,     mwLots.params)|| {}).c || 0,
    invoices:   (db.get(`SELECT COUNT(*) as c FROM invoices ${_w(mwInv)}`,  mwInv.params) || {}).c || 0,
    purchases:  (db.get(`SELECT COUNT(*) as c FROM purchases ${_w(mwPur)}`, mwPur.params) || {}).c || 0,
    bills:      (db.get(`SELECT COUNT(*) as c FROM bills ${_w(mwBill)}`,    mwBill.params)|| {}).c || 0,
    debit_notes:(db.get(`SELECT COUNT(*) as c FROM debit_notes ${_w(mwDN)}`,mwDN.params)  || {}).c || 0,
  };

  // All auctions (for the dashboard picker)
  const allAuctions = db.all(
    `SELECT id, ano, date, crop_type, mode FROM auctions ${_w(mwA)} ORDER BY id DESC LIMIT 50`,
    mwA.params
  );

  // Lot classification fragments shared by cumulative / per-trade /
  // branch aggregates (sold = code present & != WD; wd = code == WD).
  const _SOLD = `(code IS NOT NULL AND TRIM(code)<>'' AND UPPER(TRIM(code))<>'WD')`;
  const _WD   = `(UPPER(TRIM(COALESCE(code,'')))='WD')`;

  // ── Cumulative totals across ALL trades (lifetime) ──
  // Aggregates over every lot in every (mode-matching) auction.
  const cumRow = db.get(
    `SELECT COALESCE(SUM(qty),0) as qty,
            COALESCE(SUM(amount),0) as amount,
            COUNT(*) as lots,
            COALESCE(SUM(CASE WHEN ${_SOLD} THEN qty ELSE 0 END),0) as sold_qty,
            COALESCE(SUM(CASE WHEN ${_WD}   THEN qty ELSE 0 END),0) as wd_qty,
            COALESCE(SUM(CASE WHEN ${_SOLD} THEN amount ELSE 0 END),0) as sold_value,
            MIN(CASE WHEN ${_SOLD} AND price>0 AND amount>0 THEN price END) as min_price,
            MAX(CASE WHEN ${_SOLD} AND price>0 AND amount>0 THEN price END) as max_price
     FROM lots ${_w(mwLots)}`,
    mwLots.params
  ) || {};
  const cumulative = {
    qty:    cumRow.qty    || 0,
    amount: cumRow.amount || 0,
    lots:   cumRow.lots   || 0,
    auctions: counts.auctions,
    sold_qty: cumRow.sold_qty || 0,
    wd_qty:   cumRow.wd_qty   || 0,
    min_price: cumRow.min_price,
    max_price: cumRow.max_price,
    avg_price: (cumRow.sold_qty || 0) > 0 ? (cumRow.sold_value / cumRow.sold_qty) : 0,
  };

  // ── Branch-wise totals (dashboard branch tiles) ──
  // Scoped to the drilled-in auction when one is picked, else all lots
  // (mode-filtered). Groups by branch; blanks roll into "(unspecified)".
  let branchTotals;
  {
    const _raw = req.query.auction_id;
    const _isAll = (_raw === 'all' || _raw === '' || _raw === undefined);
    const drillId = _isAll ? null : (parseInt(_raw, 10) || null);
    if (drillId) {
      branchTotals = db.all(
        `SELECT COALESCE(NULLIF(TRIM(branch),''),'(unspecified)') as branch,
                COUNT(*) as lots, COALESCE(SUM(qty),0) as qty, COALESCE(SUM(amount),0) as amount
         FROM lots WHERE auction_id = ?
         GROUP BY COALESCE(NULLIF(TRIM(branch),''),'(unspecified)')
         ORDER BY amount DESC`,
        [drillId]
      );
    } else {
      branchTotals = db.all(
        `SELECT COALESCE(NULLIF(TRIM(branch),''),'(unspecified)') as branch,
                COUNT(*) as lots, COALESCE(SUM(qty),0) as qty, COALESCE(SUM(amount),0) as amount
         FROM lots ${_w(mwLots)}
         GROUP BY COALESCE(NULLIF(TRIM(branch),''),'(unspecified)')
         ORDER BY amount DESC`,
        mwLots.params
      );
    }
  }

  // ── Per-trade breakdown (one row per auction, newest first) ──
  // One query with a LEFT JOIN so auctions with zero lots still appear.
  // Mode-filtered via the parent auction. Alias `a` carries the mode.
  const mwAa = modeWhereClause(db, 'a.mode');
  const perTradeBreakdown = db.all(
    `SELECT a.id, a.ano, a.date, a.crop_type, a.mode,
            COUNT(l.id) as lots,
            COALESCE(SUM(l.qty),0) as qty,
            COALESCE(SUM(l.amount),0) as amount,
            COALESCE(SUM(CASE WHEN l.amount > 0 THEN 1 ELSE 0 END),0) as priced,
            COALESCE(SUM(CASE WHEN l.invo IS NOT NULL AND l.invo != '' THEN 1 ELSE 0 END),0) as invoiced,
            COALESCE(SUM(CASE WHEN ${_SOLD.replace(/code/g,'l.code')} THEN l.qty ELSE 0 END),0) as sold_qty,
            COALESCE(SUM(CASE WHEN ${_WD.replace(/code/g,'l.code')}   THEN l.qty ELSE 0 END),0) as wd_qty,
            COALESCE(SUM(CASE WHEN ${_SOLD.replace(/code/g,'l.code')} THEN l.amount ELSE 0 END),0) as sold_value,
            MIN(CASE WHEN ${_SOLD.replace(/code/g,'l.code')} AND l.price>0 AND l.amount>0 THEN l.price END) as min_price,
            MAX(CASE WHEN ${_SOLD.replace(/code/g,'l.code')} AND l.price>0 AND l.amount>0 THEN l.price END) as max_price
     FROM auctions a
     LEFT JOIN lots l ON l.auction_id = a.id
     ${mwAa.sql ? 'WHERE 1=1' + mwAa.sql : ''}
     GROUP BY a.id, a.ano, a.date, a.crop_type, a.mode
     ORDER BY a.date DESC, a.id DESC
     LIMIT 50`,
    mwAa.params
  );
  // Weighted avg sold price per trade (Σ sold value / Σ sold qty).
  for (const t of perTradeBreakdown) {
    t.avg_price = (t.sold_qty || 0) > 0 ? (t.sold_value / t.sold_qty) : 0;
  }

  // Pick: ?auction_id=N if provided
  //   - "all" (or no param) => dashboard shows cumulative view, no individual auction highlighted
  //   - specific id         => dashboard drills into that one auction
  let currentAuction = null;
  const rawAuctionId = req.query.auction_id;
  const isAllMode = (rawAuctionId === 'all' || rawAuctionId === '' || rawAuctionId === undefined);
  if (!isAllMode) {
    const requestedId = parseInt(rawAuctionId);
    if (requestedId) {
      currentAuction = db.get('SELECT * FROM auctions WHERE id = ?', [requestedId]);
    }
  }

  let auctionStats = null;
  if (currentAuction) {
    const totalLots  = (db.get('SELECT COUNT(*) as c FROM lots WHERE auction_id = ?', [currentAuction.id]) || {}).c || 0;
    const priced     = (db.get('SELECT COUNT(*) as c FROM lots WHERE auction_id = ? AND amount > 0', [currentAuction.id]) || {}).c || 0;
    const invoiced   = (db.get(`SELECT COUNT(*) as c FROM lots WHERE auction_id = ? AND invo IS NOT NULL AND invo != ''`, [currentAuction.id]) || {}).c || 0;
    const totalQty   = (db.get('SELECT COALESCE(SUM(qty),0) as s FROM lots WHERE auction_id = ?', [currentAuction.id]) || {}).s || 0;
    const totalAmt   = (db.get('SELECT COALESCE(SUM(amount),0) as s FROM lots WHERE auction_id = ?', [currentAuction.id]) || {}).s || 0;
    auctionStats = { ...currentAuction, totalLots, priced, invoiced, totalQty, totalAmt };
  }

  // Top sellers (this week — by total amount in auctions dated within last 7 days)
  // Mode-filtered via the joined auction's mode column.
  const topSellers = db.all(
    `SELECT l.name as name, COUNT(*) as lots, COALESCE(SUM(l.qty),0) as qty, COALESCE(SUM(l.amount),0) as amount
     FROM lots l JOIN auctions a ON a.id = l.auction_id
     WHERE a.date >= date('now','-7 days') AND l.name IS NOT NULL AND l.name != ''
       ${mwAa.sql}
     GROUP BY l.name
     ORDER BY amount DESC
     LIMIT 5`,
    mwAa.params
  );

  // Recent invoices (last 5) — mode-filtered via the parent auction.
  // No alias on FROM: mwInv's subquery references `invoices.auction_id`,
  // which SQLite refuses to resolve once the table is aliased.
  const recentInvoices = db.all(
    `SELECT id, sale, invo, buyer, buyer1, tot, date, place
     FROM invoices
     ${mwInv.sql ? 'WHERE 1=1' + mwInv.sql : ''}
     ORDER BY id DESC LIMIT 5`,
    mwInv.params
  );

  // Today's trade totals (active auction lots)
  const todayQty = auctionStats ? auctionStats.totalQty : 0;
  const todayAmt = auctionStats ? auctionStats.totalAmt : 0;

  // Revenue this month (sum of invoice totals in current month).
  // Mode-filtered via the parent auction.
  const monthTot = (db.get(
    `SELECT COALESCE(SUM(tot),0) as s FROM invoices
     WHERE date >= date('now','start of month') ${mwInv.sql}`,
    mwInv.params
  ) || {}).s || 0;
  // Revenue last month (for comparison)
  const lastMonthTot = (db.get(
    `SELECT COALESCE(SUM(tot),0) as s FROM invoices
     WHERE date >= date('now','start of month','-1 month')
       AND date <  date('now','start of month')
       ${mwInv.sql}`,
    mwInv.params
  ) || {}).s || 0;

  // Pending invoices:
  //   - Drilled into an auction: un-invoiced priced lots in that auction
  //     (mode-implicit — the auction picker already filtered to mode)
  //   - Cumulative mode: un-invoiced priced lots across ALL auctions
  //     matching the current mode.
  let pendingInvoices = 0;
  if (currentAuction) {
    pendingInvoices = (db.get(
      `SELECT COUNT(DISTINCT buyer) as c FROM lots
       WHERE auction_id = ? AND amount > 0 AND buyer IS NOT NULL AND buyer != ''
         AND (invo IS NULL OR invo = '')`, [currentAuction.id]
    ) || {}).c || 0;
  } else {
    pendingInvoices = (db.get(
      `SELECT COUNT(DISTINCT buyer || '|' || auction_id) as c FROM lots
       WHERE amount > 0 AND buyer IS NOT NULL AND buyer != ''
         AND (invo IS NULL OR invo = '')
         ${mwLots.sql}`,
      mwLots.params
    ) || {}).c || 0;
  }

  res.json({
    counts,
    cumulative,
    perTradeBreakdown,
    branchTotals,
    currentAuction: auctionStats,
    allAuctions,
    topSellers,
    recentInvoices,
    kpi: {
      todayQty, todayAmt,
      activeLots: auctionStats ? auctionStats.totalLots : 0,
      pendingInvoices,
      monthRevenue: monthTot,
      lastMonthRevenue: lastMonthTot,
    }
  });
});

// ══════════════════════════════════════════════════════════════
// REVENUE TREND — daily invoiced revenue for the Insights chart.
// ?days=7|14|30 (clamped 1..90). Returns a continuous series (zero
// days filled) so the line chart has an unbroken x-axis. Mode-filtered
// via the parent auction, matching /api/stats.
// ══════════════════════════════════════════════════════════════
app.get('/api/stats/revenue-trend', requireView, (req, res) => {
  const db = getDb();
  let days = parseInt(req.query.days, 10) || 7;
  days = Math.max(1, Math.min(90, days));
  const mwInv = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=invoices.auction_id)');
  const rows = db.all(
    `SELECT date(date) AS d, COALESCE(SUM(tot),0) AS total, COUNT(*) AS count
     FROM invoices
     WHERE date(date) >= date('now','localtime', ?) ${mwInv.sql}
     GROUP BY date(date)`,
    ['-' + (days - 1) + ' days', ...mwInv.params]
  );
  const map = {};
  for (const r of rows) map[r.d] = r;
  // Fill every day so the chart x-axis is continuous. Today comes from
  // SQLite localtime so it matches the GROUP BY date() bucketing above.
  const todayStr = (db.get(`SELECT date('now','localtime') AS d`) || {}).d;
  const base = new Date(todayStr + 'T00:00:00');
  const pad = n => String(n).padStart(2, '0');
  const series = [];
  for (let i = days - 1; i >= 0; i--) {
    const dt = new Date(base);
    dt.setDate(base.getDate() - i);
    const d = `${dt.getFullYear()}-${pad(dt.getMonth() + 1)}-${pad(dt.getDate())}`;
    const hit = map[d];
    series.push({ date: d, total: hit ? Number(hit.total) : 0, count: hit ? Number(hit.count) : 0 });
  }
  res.json({ days, series });
});

// ══════════════════════════════════════════════════════════════
// INSIGHTS — per-trade × per-branch analytics. Drives the Dashboard
// "headline metrics" tiles AND the Insights tab. Scope is either a
// single trade (?auction_id=N) or a date range (?from=&to=, defaulting
// to the current calendar month). Mode-filtered like /api/stats.
//
// Lot classification (matches the Lots-tab dashboard logic):
//   sold      = code present and != 'WD'
//   withdrawn = code == 'WD'
//   unsold    = no code
// Price min/max/avg are over SOLD lots only (price>0, amount>0); avg is
// the weighted Σamount/Σqty. payable_to_sellers = Σ lots.balance.
// ══════════════════════════════════════════════════════════════
app.get('/api/insights', requireView, (req, res) => {
  const db = getDb();
  const mwAa = modeWhereClause(db, 'a.mode');

  // ── Resolve scope → the auctions in play + the date range label. ──
  let auctions, range;
  const rawAid = req.query.auction_id;
  if (rawAid && String(rawAid).trim() !== '' && String(rawAid) !== 'all') {
    const aid = parseInt(rawAid, 10);
    auctions = db.all(`SELECT id, ano, date, crop_type, state FROM auctions a WHERE a.id = ?`, [aid]);
    const a0 = auctions[0];
    range = { from: (a0 && a0.date) || '', to: (a0 && a0.date) || '' };
  } else {
    let from = String(req.query.from || '').trim();
    let to   = String(req.query.to   || '').trim();
    if (!from || !to) {
      const m = db.get(`SELECT date('now','localtime','start of month') AS f, date('now','localtime') AS t`) || {};
      from = from || m.f; to = to || m.t;
    }
    range = { from, to };
    auctions = db.all(
      `SELECT id, ano, date, crop_type, state FROM auctions a
       WHERE date(a.date) BETWEEN date(?) AND date(?) ${mwAa.sql}
       ORDER BY date(a.date) ASC, a.id ASC`,
      [from, to, ...mwAa.params]
    );
  }

  const aids = auctions.map(a => a.id);
  const blankTotals = {
    trades: auctions.length, lots: 0, bags: 0, qty: 0, value: 0,
    sold: 0, sold_bags: 0, sold_qty: 0, sold_value: 0,
    withdrawn: 0, wd_bags: 0, wd_qty: 0, wd_value: 0,
    min_price: null, max_price: null, avg_price: 0,
    payable_to_sellers: 0, outstanding_by_buyers: 0,
  };
  if (!aids.length) {
    return res.json({
      range, totals: blankTotals, perTrade: [], perBranch: [],
      branchStacked: { labels: [], datasets: [] },
      outstandingByBuyer: [], buyerActivity: [],
    });
  }
  const ph = aids.map(() => '?').join(',');

  // Reusable code-classification SQL fragments.
  const SOLD = `(code IS NOT NULL AND TRIM(code)<>'' AND UPPER(TRIM(code))<>'WD')`;
  const WD   = `(UPPER(TRIM(COALESCE(code,'')))='WD')`;

  // ── Per (auction, branch) lot aggregation in one pass. ──
  const grid = db.all(
    `SELECT auction_id,
            COALESCE(NULLIF(TRIM(branch),''),'(unspecified)') AS branch,
            COUNT(*) AS lots,
            COALESCE(SUM(bags),0) AS bags,
            COALESCE(SUM(qty),0)  AS qty,
            COALESCE(SUM(amount),0) AS value,
            SUM(CASE WHEN ${SOLD} THEN 1 ELSE 0 END) AS sold,
            SUM(CASE WHEN ${WD}   THEN 1 ELSE 0 END) AS withdrawn,
            SUM(CASE WHEN COALESCE(TRIM(code),'')='' THEN 1 ELSE 0 END) AS unsold,
            COALESCE(SUM(CASE WHEN ${SOLD} THEN bags   ELSE 0 END),0) AS sold_bags,
            COALESCE(SUM(CASE WHEN ${SOLD} THEN qty    ELSE 0 END),0) AS sold_qty,
            COALESCE(SUM(CASE WHEN ${SOLD} THEN amount ELSE 0 END),0) AS sold_value,
            COALESCE(SUM(CASE WHEN ${WD}   THEN bags   ELSE 0 END),0) AS wd_bags,
            COALESCE(SUM(CASE WHEN ${WD}   THEN qty    ELSE 0 END),0) AS wd_qty,
            COALESCE(SUM(CASE WHEN ${WD}   THEN amount ELSE 0 END),0) AS wd_value,
            MIN(CASE WHEN ${SOLD} AND price>0 AND amount>0 THEN price END) AS min_price,
            MAX(CASE WHEN ${SOLD} AND price>0 AND amount>0 THEN price END) AS max_price,
            COALESCE(SUM(balance),0) AS payable
     FROM lots
     WHERE auction_id IN (${ph})
     GROUP BY auction_id, COALESCE(NULLIF(TRIM(branch),''),'(unspecified)')`,
    aids
  );

  // ── Roll the grid up into perTrade, perBranch, totals. ──
  const num = v => Number(v) || 0;
  const byTrade  = new Map();   // auction_id → trade agg
  const byBranch = new Map();   // branch     → branch agg
  const totals = { ...blankTotals };
  let priceMin = null, priceMax = null;

  const newTradeAgg = (a) => ({
    id: a.id, ano: a.ano, date: a.date, state: a.state || '', crop_type: a.crop_type || '',
    lots: 0, sold: 0, withdrawn: 0, qty: 0, value: 0,
    sold_qty: 0, sold_value: 0,
    _minP: null, _maxP: null,
    branches: [],   // [{branch, value, qty, lots}]
  });
  for (const a of auctions) byTrade.set(a.id, newTradeAgg(a));

  const newBranchAgg = (name) => ({
    branch: name, lots: 0, sold: 0, withdrawn: 0, unsold: 0,
    bags: 0, qty: 0, value: 0, sold_qty: 0, sold_value: 0,
    _minP: null, _maxP: null, payable_to_sellers: 0,
  });

  for (const g of grid) {
    const t = byTrade.get(g.auction_id);
    if (t) {
      t.lots += num(g.lots); t.sold += num(g.sold); t.withdrawn += num(g.withdrawn);
      t.qty += num(g.qty); t.value += num(g.value);
      t.sold_qty += num(g.sold_qty); t.sold_value += num(g.sold_value);
      if (g.min_price != null) t._minP = (t._minP == null) ? num(g.min_price) : Math.min(t._minP, num(g.min_price));
      if (g.max_price != null) t._maxP = (t._maxP == null) ? num(g.max_price) : Math.max(t._maxP, num(g.max_price));
      t.branches.push({ branch: g.branch, value: num(g.value), qty: num(g.qty), lots: num(g.lots) });
    }
    let b = byBranch.get(g.branch);
    if (!b) { b = newBranchAgg(g.branch); byBranch.set(g.branch, b); }
    b.lots += num(g.lots); b.sold += num(g.sold); b.withdrawn += num(g.withdrawn); b.unsold += num(g.unsold);
    b.bags += num(g.bags); b.qty += num(g.qty); b.value += num(g.value);
    b.sold_qty += num(g.sold_qty); b.sold_value += num(g.sold_value);
    b.payable_to_sellers += num(g.payable);
    if (g.min_price != null) b._minP = (b._minP == null) ? num(g.min_price) : Math.min(b._minP, num(g.min_price));
    if (g.max_price != null) b._maxP = (b._maxP == null) ? num(g.max_price) : Math.max(b._maxP, num(g.max_price));

    totals.lots += num(g.lots); totals.bags += num(g.bags); totals.qty += num(g.qty); totals.value += num(g.value);
    totals.sold += num(g.sold); totals.sold_bags += num(g.sold_bags); totals.sold_qty += num(g.sold_qty); totals.sold_value += num(g.sold_value);
    totals.withdrawn += num(g.withdrawn); totals.wd_bags += num(g.wd_bags); totals.wd_qty += num(g.wd_qty); totals.wd_value += num(g.wd_value);
    totals.payable_to_sellers += num(g.payable);
    if (g.min_price != null) priceMin = (priceMin == null) ? num(g.min_price) : Math.min(priceMin, num(g.min_price));
    if (g.max_price != null) priceMax = (priceMax == null) ? num(g.max_price) : Math.max(priceMax, num(g.max_price));
  }
  totals.min_price = priceMin;
  totals.max_price = priceMax;
  totals.avg_price = totals.sold_qty > 0 ? (totals.sold_value / totals.sold_qty) : 0;

  // ── Invoices in scope: outstanding-by-buyer + grand outstanding. ──
  const invByBuyer = db.all(
    `SELECT COALESCE(NULLIF(TRIM(buyer1),''), buyer) AS buyer_name,
            buyer AS buyer_code,
            COUNT(*) AS invoices,
            COALESCE(SUM(tot),0) AS value
     FROM invoices
     WHERE auction_id IN (${ph})
     GROUP BY COALESCE(NULLIF(TRIM(buyer1),''), buyer), buyer
     ORDER BY value DESC
     LIMIT 100`,
    aids
  ).map(r => ({ buyer_name: r.buyer_name || '(unknown)', buyer_code: r.buyer_code || '', invoices: num(r.invoices), value: num(r.value) }));
  totals.outstanding_by_buyers = invByBuyer.reduce((s, r) => s + r.value, 0);

  // ── Buyer activity leaderboard (by purchased lot value). ──
  const buyerActivity = db.all(
    `SELECT COALESCE(NULLIF(TRIM(buyer1),''), buyer) AS buyer_name,
            COUNT(*) AS lots,
            COALESCE(SUM(qty),0) AS qty,
            COALESCE(SUM(amount),0) AS value
     FROM lots
     WHERE auction_id IN (${ph}) AND ${SOLD} AND buyer IS NOT NULL AND TRIM(buyer)<>''
     GROUP BY COALESCE(NULLIF(TRIM(buyer1),''), buyer)
     ORDER BY value DESC
     LIMIT 15`,
    aids
  ).map(r => ({ buyer_name: r.buyer_name || '(unknown)', lots: num(r.lots), qty: num(r.qty), value: num(r.value) }));

  // ── Finalize perTrade + perBranch (resolve weighted avg, prune helpers). ──
  const perTrade = auctions.map(a => {
    const t = byTrade.get(a.id);
    const avg = t.sold_qty > 0 ? (t.sold_value / t.sold_qty) : 0;
    // Trade's branch list, biggest contribution first (frontend shows top 4).
    const branches = t.branches.slice().sort((x, y) => y.value - x.value);
    return {
      id: t.id, ano: t.ano, date: t.date, state: t.state, crop_type: t.crop_type,
      lots: t.lots, sold: t.sold, withdrawn: t.withdrawn, qty: t.qty, value: t.value,
      min_price: t._minP, max_price: t._maxP, avg_price: avg, branches,
    };
  });

  const perBranch = Array.from(byBranch.values())
    .map(b => ({
      branch: b.branch, lots: b.lots, sold: b.sold, withdrawn: b.withdrawn, unsold: b.unsold,
      qty: b.qty, value: b.value, min_price: b._minP, max_price: b._maxP,
      avg_price: b.sold_qty > 0 ? (b.sold_value / b.sold_qty) : 0,
      payable_to_sellers: b.payable_to_sellers,
    }))
    .sort((x, y) => y.value - x.value);

  // ── Stacked bar: one stack per trade (newest-window last 20), one
  //    dataset per branch with that branch's value in each trade. ──
  const stackTrades = perTrade.slice(-20);
  const branchNames = perBranch.map(b => b.branch);
  const branchStacked = {
    labels: stackTrades.map(t => '#' + (t.ano != null ? t.ano : '')),
    datasets: branchNames.map(bn => ({
      label: bn,
      data: stackTrades.map(t => {
        const hit = (byTrade.get(t.id).branches || []).find(x => x.branch === bn);
        return hit ? hit.value : 0;
      }),
    })),
  };

  res.json({
    range, totals, perTrade, perBranch, branchStacked,
    outstandingByBuyer: invByBuyer.slice(0, 50),
    buyerActivity,
  });
});

// ══════════════════════════════════════════════════════════════
// IMPORT OLD DATA — admin-only, generic XLSX → DB importer
// ══════════════════════════════════════════════════════════════
// Replaces the ad-hoc per-resource importers for legacy migration.
// Five endpoints share one config-driven flow:
//   POST /api/import-old-data/preview      — header detection + sample rows
//   POST /api/import-old-data/verify       — dry-resolve vs. live DB (no writes)
//   POST /api/import-old-data/run          — actual INSERT (or dry-run validation)
//   GET  /api/import-old-data/history      — recent imports for the History panel
//   POST /api/import-old-data/undo/:id     — rollback a single live import
//
// Each MODULE describes a target table: which columns to write, which
// columns are the natural key (used for duplicate detection), and
// alias lists for auto-mapping fuzzy spreadsheet headers. The user
// can override the auto-map in the UI before clicking Run Import.
//
// Modules covered:
//   auctions       — Auctions / Trades (one entity, two labels via business_mode)
//   sales_invoice  — INV.DBF style sales invoices
//   purchase       — Purchase invoices (registered dealer)
//   bills          — Agriculturist Bills of Supply
//   debit_notes    — Discount debit notes
//   sellers        — NAM.DBF sellers/poolers
//   buyers         — SBL.DBF buyers/dealers
const IMPORT_MODULES = {
  // NEW: Auctions / Trades. In this app the same `auctions` table
  // surfaces under two labels depending on business_mode (Auctions in
  // e-Auction, Trades in e-Trade), so one module covers both. The
  // `mode` column is auto-stamped from current settings when blank so
  // imported rows don't leak across the two views.
  auctions: {
    label: 'Auctions / Trades',
    table: 'auctions',
    // Natural key: ano + date. ano alone can repeat across years/states.
    keyCols: ['ano', 'date'],
    fields: ['ano', 'date', 'crop_type', 'state', 'mode'],
    aliases: {
      ano: ['ano', 'auction_no', 'auctionno', 'auction', 'trade_no', 'tradeno', 'trade', 'tno', 'no'],
      date: ['date', 'auction_date', 'trade_date', 'tdate'],
      crop_type: ['crop_type', 'crpt', 'crop', 'type'],
      state: ['state', 'auction_state'],
      mode: ['mode', 'business_mode'],
    },
    // Server-computed defaults applied when the field is unmapped OR
    // its source value is blank. Keeps imported rows from leaking
    // across e-Auction/e-Trade views when the source file omits mode.
    defaults: (db) => ({
      mode: currentBusinessMode(db),
    }),
  },
  sales_invoice: {
    label: 'Sales Invoices',
    table: 'invoices',
    // Natural key = trade no (ano) + sale type + invoice no + business
    // state. Rationale:
    //   • ano   — the same invoice number legitimately recurs across
    //             different trades, so it MUST scope the key (the old
    //             ['invo','sale'] key wrongly collided cross-trade).
    //   • sale  — L/I/E share number ranges; keep them distinct.
    //   • invo  — the invoice number itself.
    //   • state — separates the TN (ISP) and KL (ASP) books. It is NOT
    //             read from the source file: it is ALWAYS derived from
    //             the PLACE — prefix "ASP" → KERALA, everything else →
    //             TAMIL NADU (see rowDefaults below). This is a firm
    //             business rule, independent of the current business
    //             state, so `state` is intentionally absent from the
    //             aliases (it can't be mapped to a source column).
    keyCols: ['ano', 'sale', 'invo', 'state'],
    // `state` is computed, never read from the source (even if the sheet
    // has a STATE column) — see rowDefaults.
    derivedFields: ['state'],
    // auction_id is derived from `ano` at import time so the imported
    // rows show up under the matching trade in the Sales tab (the list
    // filters by auction_id, not ano).
    autoFillAuctionId: true,
    fields: ['auction_id','ano','date','state','sale','invo','buyer','buyer1','gstin','place',
             'bag','qty','amount','gunny','pava_hc','ins','cgst','sgst','igst','tcs','rund','tot'],
    aliases: {
      ano: ['ano','auction_no','trade'],
      date: ['date','invoice_date','inv_date'],
      // NOTE: `state` is deliberately NOT mapped from the source — it is
      // always derived from PLACE in rowDefaults (ASP → KERALA, else
      // TAMIL NADU). Adding a state alias here would let a source column
      // override that rule, which we don't want.
      sale: ['sale','sale_type','type'],
      invo: ['invo','invoice','invoice_no','invno'],
      buyer: ['buyer','buyer_code','code'],
      buyer1: ['buyer1','buyer_name','name'],
      gstin: ['gstin','gst','gst_no'],
      place: ['place','city','pla'],
      bag: ['bag','bags','no_of_bags'],
      qty: ['qty','kilos','weight','kgs'],
      amount: ['amount','cardamom','value'],
      tot: ['tot','total','grand_total','invoice_amount'],
    },
    // Flat default for `state` — TAMIL NADU. rowDefaults below always
    // returns an explicit state, so this is just a safety net (and it
    // tells the import UI that `state` is a server-filled field, so it's
    // shown as auto-filled rather than "missing required").
    defaults: () => ({ state: 'TAMIL NADU' }),
    // Authoritative per-row state derivation from PLACE:
    //   • PLACE prefixed "ASP" (Amazing Spice Park, Kerala) → KERALA
    //   • everything else                                   → TAMIL NADU
    // `state` is never mapped from the source (no alias), so the field is
    // always blank-from-source and this value always wins. Independent of
    // the current business state.
    rowDefaults: (row, mapping) => {
      const placeSrc = mapping.place;
      const place = placeSrc ? String(row[placeSrc] || '').trim().toUpperCase() : '';
      return { state: place.startsWith('ASP') ? 'KERALA' : 'TAMIL NADU' };
    },
  },
  purchase: {
    label: 'Purchase Invoices',
    table: 'purchases',
    keyCols: ['invo'],
    // `state` is computed from BR, never read from the source — see rowDefaults.
    derivedFields: ['state'],
    autoFillAuctionId: true,
    fields: ['auction_id','ano','date','state','br','name','add_line','place','gstin','invo',
             'qty','amount','cgst','sgst','igst','rund','total','tds'],
    aliases: {
      invo:    ['invo','invoice','invoice_no'],
      name:    ['name','seller','dealer'],
      gstin:   ['gstin','gst','gst_no','cr','registration'],
      place:   ['place','city','pla'],
      add_line:['add_line','address','add','add1','address1'],
      br:      ['br','branch'],
      qty:     ['qty','kilos','weight','kgs'],
      amount:  ['amount','cardamom','value'],
      total:   ['total','grand_total','invoice_amount'],
      rund:    ['rund','round','round_off'],
      tds:     ['tds','tds_amount'],
      // NOTE: no `state` alias — state is always derived from BR (see
      // rowDefaults), never read from a source column.
    },
    // Flat default safety net + signals the UI that `state` is server-filled.
    defaults: () => ({ state: 'TAMIL NADU' }),
    // Authoritative per-row state derivation from the BR (branch) column:
    //   • BR prefixed "ASP" (e.g. ASPNEDUMKANDAM) → KERALA
    //   • everything else                          → TAMIL NADU
    // Mirrors the sales-invoice PLACE rule, but purchases key off branch.
    rowDefaults: (row, mapping) => {
      const brSrc = mapping.br;
      const br = brSrc ? String(row[brSrc] || '').trim().toUpperCase() : '';
      return { state: br.startsWith('ASP') ? 'KERALA' : 'TAMIL NADU' };
    },
  },
  bills: {
    label: 'Bills of Supply',
    table: 'bills',
    keyCols: ['bil'],
    // `state` is computed from BR, never read from the source — see rowDefaults.
    derivedFields: ['state'],
    autoFillAuctionId: true,
    fields: ['auction_id','ano','date','state','br','crpt','bil','name','add_line','pla',
             'pstate','st_code','crr','pan','qty','cost','igst','net'],
    aliases: {
      bil: ['bil','bill','bill_no'],
      name: ['name','seller','planter'],
      br: ['br','branch'],
      qty: ['qty','kilos','weight','kgs'],
      cost: ['cost','amount','cardamom'],
      net: ['net','nett','net_amount'],
      // NOTE: no `state` alias — state is always derived from BR (see
      // rowDefaults), never read from a source column.
    },
    // Flat default safety net + signals the UI that `state` is server-filled.
    defaults: () => ({ state: 'TAMIL NADU' }),
    // Authoritative per-row state derivation from the BR (branch) column:
    //   • BR prefixed "ASP" (e.g. ASPNEDUMKANDAM) → KERALA
    //   • everything else                          → TAMIL NADU
    // Mirrors the Purchases BR rule — Bills of Supply key off branch too.
    rowDefaults: (row, mapping) => {
      const brSrc = mapping.br;
      const br = brSrc ? String(row[brSrc] || '').trim().toUpperCase() : '';
      return { state: br.startsWith('ASP') ? 'KERALA' : 'TAMIL NADU' };
    },
  },
  debit_notes: {
    label: 'Debit Notes',
    table: 'debit_notes',
    keyCols: ['note_no','ano'],
    fields: ['ano','date','state','name','note_no','amount','cgst','sgst','igst','total'],
    aliases: {
      note_no: ['note_no','note','dn_no'],
      name: ['name','dealer','buyer'],
    },
  },
  sellers: {
    label: 'Sellers',
    table: 'traders',
    keyCols: ['name','cr'],
    fields: ['name','cr','pan','tel','aadhar','padd','ppla','pin','pstate','pst_code',
             'ifsc','acctnum','holder_name'],
    aliases: {
      name: ['name','seller','planter','trader'],
      cr: ['cr','gstin'],
      padd: ['padd','address','add','add1','address1'],
      ppla: ['ppla','place','pla','city'],
      pin: ['pin','pincode','zip'],
    },
  },
  buyers: {
    label: 'Buyers',
    table: 'buyers',
    keyCols: ['buyer','code'],
    fields: ['buyer','buyer1','code','sbl','add1','add2','pla','pin','state',
             'st_code','gstin','pan','tel','ti','sale','email'],
    aliases: {
      buyer: ['buyer','buyer_code','code'],
      buyer1: ['buyer1','buyer_name','name'],
      pla: ['pla','place','city'],
      pin: ['pin','pincode','zip'],
    },
  },
};

// Header → field mapping. Normalises both sides (lowercase, collapse
// separators) before matching aliases. Returns { fieldName: srcHeader }
// for fields the importer auto-detected.
function _importMapHeaders(headers, moduleDef) {
  const norm = s => String(s || '').trim().toLowerCase().replace(/[\s\-/]+/g, '_');
  const out = {};
  for (const field of moduleDef.fields) {
    const aliases = (moduleDef.aliases && moduleDef.aliases[field]) || [field];
    for (const h of headers) {
      if (aliases.includes(norm(h))) { out[field] = h; break; }
    }
  }
  // Derived fields are ALWAYS computed server-side (e.g. `state` from the
  // PLACE/BR prefix), never read from the source — even if the sheet has a
  // matching column. Strip them so the auto-mapper's `[field]` fallback
  // can't silently bind a source column that would override the rule.
  if (Array.isArray(moduleDef.derivedFields)) {
    for (const f of moduleDef.derivedFields) delete out[f];
  }
  return out;
}

// POST /preview — read headers, detect mapping, return first 50 rows.
// No DB writes. Idempotent — used repeatedly as the user adjusts the
// mapping in the UI.
app.post('/api/import-old-data/preview', requireAdmin, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const moduleKey = req.body.module;
  const def = IMPORT_MODULES[moduleKey];
  if (!def) {
    if (req.file) fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: 'Unknown module', available: Object.keys(IMPORT_MODULES) });
  }
  try {
    const wb = XLSX.readFile(req.file.path);
    const ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    const headers = rows.length ? Object.keys(rows[0]) : [];
    const mapping = _importMapHeaders(headers, def);
    res.json({
      module: moduleKey,
      label: def.label,
      total: rows.length,
      headers,
      // Full DB field list so the UI can render a mapping editor for
      // every column, not just the auto-detected ones.
      fields:  def.fields,
      keyCols: def.keyCols,
      autoFillAuctionId: !!def.autoFillAuctionId,
      // Fields the server fills from a default when the source omits them
      // (e.g. `state` ← current business state). The UI renders these as
      // auto-filled rather than "missing", and excludes them from the
      // "rows missing required values" tally.
      autoDefaults: (() => { try { return def.defaults ? (def.defaults(getDb()) || {}) : {}; } catch (_) { return {}; } })(),
      detectedMapping: mapping,
      // Flag key columns the user still needs to map — but EXCLUDE any
      // that a module default fills (e.g. `state` ← current business
      // state), since those resolve without a source column.
      missingFields: (() => {
        let dk = [];
        try { dk = def.defaults ? Object.keys(def.defaults(getDb()) || {}) : []; } catch (_) {}
        return def.fields.filter(f => !mapping[f] && def.keyCols.includes(f) && !dk.includes(f));
      })(),
      preview: rows.slice(0, 50),
    });
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    if (req.file) fs.unlink(req.file.path, () => {});
  }
});

// POST /verify — dry-resolve the entire file against the live DB
// using the user's mapping. Returns per-row status (new / duplicate /
// invalid), reasons for invalids, the field-by-field diff vs. any
// existing row, and accurate summary counts. NO DB WRITES.
//
// Differs from /preview (header detection only) and /run (actually
// writes): same validation /run does, against the same mapping the
// user chose, but with zero side effects. Lets the operator catch
// wrong mappings, missing trades, or accidental overwrites before
// hitting production data.
app.post('/api/import-old-data/verify', requireAdmin, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const moduleKey = req.body.module;
  const def = IMPORT_MODULES[moduleKey];
  if (!def) {
    fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: 'Unknown module' });
  }
  let userMapping = {};
  if (req.body.mapping) {
    try { userMapping = JSON.parse(req.body.mapping) || {}; } catch (_) {}
  }
  const PER_BUCKET_LIMIT = 50;
  const db = getDb();
  try {
    const wb = XLSX.readFile(req.file.path);
    const ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    const total = rows.length;
    if (!total) {
      fs.unlink(req.file.path, () => {});
      return res.json({
        module: moduleKey, label: def.label, total: 0,
        fields: def.fields, keyCols: def.keyCols,
        autoFillAuctionId: !!def.autoFillAuctionId,
        sampleLimit: PER_BUCKET_LIMIT,
        counts: { new: 0, duplicate: 0, duplicateChanged: 0, invalidAno: 0, invalidRequired: 0 },
        samples: { new: [], invalid: [], dupChanges: [], dupIdentical: [] },
      });
    }

    const headers = Object.keys(rows[0]);
    // Merge user overrides on top of auto-detected mapping. An
    // explicit '' from the user means SKIP this field — pop it from
    // the merged map so /run and /verify agree on what gets written.
    const autoDetected = _importMapHeaders(headers, def);
    const mapping = Object.assign({}, autoDetected, userMapping);
    for (const k of Object.keys(userMapping || {})) {
      const v = userMapping[k];
      if (v === '' || v === null) delete mapping[k];
    }
    // Derived fields are always computed server-side — drop any mapping
    // (auto or manual) so it can never override the rule.
    if (Array.isArray(def.derivedFields)) for (const f of def.derivedFields) delete mapping[f];
    const fieldSources = def.fields.map(f => [f, mapping[f] || null]);
    const auctionIdSlot = def.fields.indexOf('auction_id');

    // Cached ano→auction_id resolver. One DB call per distinct ano.
    const auctionIdCache = new Map();
    const resolveAuctionId = (ano) => {
      const key = String(ano || '').trim();
      if (!key) return null;
      if (auctionIdCache.has(key)) return auctionIdCache.get(key);
      const row = db.get('SELECT id FROM auctions WHERE ano = ? LIMIT 1', [key]);
      const id  = row ? row.id : null;
      auctionIdCache.set(key, id);
      return id;
    };

    let cntNew = 0, cntDup = 0, cntDupChanged = 0, cntInvAno = 0, cntInvReq = 0;
    // Four independent sample buckets so the UI shows concrete rows
    // from each non-empty status even when one bucket dominates the
    // first 100 rows of the file. Counts are always over the whole file.
    const sampleNew = [];
    const sampleInvalid = [];
    const sampleDupChanges = [];
    const sampleDupIdentical = [];

    // Snapshot flat module defaults once (a DB read); per-row rowDefaults
    // (no DB) are recomputed inside the loop and override the flat ones.
    const _flatDefaults = (typeof def.defaults === 'function')
      ? (() => { try { return def.defaults(db) || {}; } catch (_) { return {}; } })()
      : {};

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const values = {};
      for (const [f, src] of fieldSources) {
        values[f] = src ? r[src] : '';
      }
      // Fill blank fields from module defaults. Per-row rowDefaults (e.g.
      // state ← KERALA when PLACE is ASP-prefixed) override the flat
      // defaults; a non-blank source value always wins over both.
      const _eff = Object.assign({}, _flatDefaults,
        (typeof def.rowDefaults === 'function' ? (() => { try { return def.rowDefaults(r, mapping, db) || {}; } catch (_) { return {}; } })() : {}));
      for (const k of Object.keys(_eff)) {
        const cur = values[k];
        if (cur == null || String(cur).trim() === '') values[k] = _eff[k];
      }

      const reasons = [];

      let anoResolutionFailed = false;
      if (def.autoFillAuctionId && auctionIdSlot >= 0) {
        const anoSrc = mapping.ano;
        const anoVal = anoSrc ? r[anoSrc] : '';
        const aid = resolveAuctionId(anoVal);
        if (aid == null) {
          anoResolutionFailed = true;
          reasons.push('No trade found for ano="' + String(anoVal || '').trim() + '" — create the auction first or fix the mapping.');
        } else {
          values.auction_id = aid;
        }
      }

      // Required-field check uses the same keyCols /run uses for dup
      // detection. A row missing any of them can't be inserted (and
      // can't be checked for duplicates).
      // Use the post-default `values` (defaults already applied above) so
      // a key column satisfied by a module default — e.g. `state` from the
      // current business state — counts as present rather than missing.
      const missingKeys = [];
      for (const k of def.keyCols) {
        const v = values[k];
        if (v == null || String(v).trim() === '') missingKeys.push(k);
      }
      const requiredMissing = missingKeys.length > 0;
      if (requiredMissing) {
        reasons.push('Missing required value(s): ' + missingKeys.join(', '));
      }

      // Duplicate detection — only meaningful when every keyCol resolves
      // to a non-blank value. Mirrors /run's gate so verify counts are
      // accurate. Keyed on the post-default values for the same reason.
      let existing = null;
      let diff = null;
      if (!requiredMissing) {
        const keyVals = def.keyCols.map(k => values[k]);
        const whereSql = def.keyCols.map(k => `${k} = ?`).join(' AND ');
        existing = db.get(`SELECT * FROM ${def.table} WHERE ${whereSql} LIMIT 1`, keyVals);
        if (existing) {
          diff = {};
          for (const f of def.fields) {
            const newVal = values[f];
            const oldVal = existing[f];
            const a = oldVal == null ? '' : String(oldVal);
            const b = newVal == null ? '' : String(newVal);
            if (a !== b) {
              diff[f] = {
                old: oldVal == null ? '' : oldVal,
                new: newVal == null ? '' : newVal,
              };
            }
          }
          if (Object.keys(diff).length === 0) diff = null;
        }
      }

      let status;
      if (requiredMissing) {
        status = 'invalid';
        cntInvReq++;
      } else if (anoResolutionFailed) {
        status = 'invalid';
        cntInvAno++;
      } else if (existing) {
        status = 'duplicate';
        cntDup++;
        if (diff) cntDupChanged++;
      } else {
        status = 'new';
        cntNew++;
      }

      const entry = {
        row: i + 2,
        status,
        reasons,
        values,
        existing: existing || null,
        diff,
      };
      if (status === 'new' && sampleNew.length < PER_BUCKET_LIMIT) {
        sampleNew.push(entry);
      } else if (status === 'invalid' && sampleInvalid.length < PER_BUCKET_LIMIT) {
        sampleInvalid.push(entry);
      } else if (status === 'duplicate' && diff && sampleDupChanges.length < PER_BUCKET_LIMIT) {
        sampleDupChanges.push(entry);
      } else if (status === 'duplicate' && !diff && sampleDupIdentical.length < PER_BUCKET_LIMIT) {
        sampleDupIdentical.push(entry);
      }
    }

    // Diagnostic: how many rows live in the target table right now?
    // Helps the operator confirm a "Delete All" actually emptied the
    // target before re-importing.
    let targetRowCount = 0;
    try {
      const r = db.get(`SELECT COUNT(*) as c FROM ${def.table}`);
      targetRowCount = r ? Number(r.c || 0) : 0;
    } catch (_) { /* table missing — leave at 0 */ }

    fs.unlink(req.file.path, () => {});
    res.json({
      module: moduleKey,
      label: def.label,
      total,
      fields: def.fields,
      keyCols: def.keyCols,
      autoFillAuctionId: !!def.autoFillAuctionId,
      autoDefaults: (() => { try { return def.defaults ? (def.defaults(getDb()) || {}) : {}; } catch (_) { return {}; } })(),
      sampleLimit: PER_BUCKET_LIMIT,
      targetTable: def.table,
      targetRowCount,
      counts: {
        new: cntNew,
        duplicate: cntDup,
        duplicateChanged: cntDupChanged,
        invalidAno: cntInvAno,
        invalidRequired: cntInvReq,
      },
      samples: {
        new: sampleNew,
        invalid: sampleInvalid,
        dupChanges: sampleDupChanges,
        dupIdentical: sampleDupIdentical,
      },
    });
  } catch (e) {
    if (req.file) fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: e.message });
  }
});

// POST /run — actually insert (or dry-run validate). Records every
// run in import_log so the History panel can show + offer Undo.
// On success, returns counts + the new import_log id so the client
// can reference it for follow-up actions.
app.post('/api/import-old-data/run', requireAdmin, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const moduleKey = req.body.module;
  const def = IMPORT_MODULES[moduleKey];
  if (!def) {
    fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: 'Unknown module' });
  }
  const dryRun = String(req.body.dryRun || '').toLowerCase() === 'true';
  let userMapping = {};
  if (req.body.mapping) {
    try { userMapping = JSON.parse(req.body.mapping) || {}; } catch (_) {}
  }

  const db = getDb();
  let imported = 0, skipped = 0, failed = 0;
  const errors = [];
  // Per-row detail of duplicate-skipped rows so the UI can show WHICH
  // rows were skipped (not just the count). Capped to bound the payload.
  const skippedDetails = [];
  let total = 0;
  // Captured for /undo so we can roll back a specific import's inserts.
  // Lives outside the try block so the audit-log INSERT below can see
  // it (an earlier bug had this scoped inside the try and successful
  // imports never made it into history).
  const insertedIds = [];

  try {
    const wb = XLSX.readFile(req.file.path);
    const ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    total = rows.length;
    if (!total) throw new Error('File is empty');

    const headers  = Object.keys(rows[0]);
    // Two distinct user signals merged on top of auto-detected mapping:
    //   • field absent       → keep auto-detected (no opinion)
    //   • field explicit ''  → SKIP this field (user picked "— skip —")
    // Without honoring explicit '', picking "— skip —" on an auto-
    // detected field has no effect.
    const autoDetected = _importMapHeaders(headers, def);
    const mapping = Object.assign({}, autoDetected, userMapping);
    for (const k of Object.keys(userMapping || {})) {
      const v = userMapping[k];
      if (v === '' || v === null) delete mapping[k];
    }
    // Derived fields are always computed server-side — drop any mapping
    // (auto or manual) so it can never override the rule.
    if (Array.isArray(def.derivedFields)) for (const f of def.derivedFields) delete mapping[f];

    const fieldSources = def.fields.map(f => [f, mapping[f] || null]);
    const valuePlaceholders = def.fields.map(() => '?').join(',');
    const insertSql = `INSERT INTO ${def.table} (${def.fields.join(',')}) VALUES (${valuePlaceholders})`;

    // ano→auction_id cache for autoFillAuctionId modules — single DB
    // call per distinct ano across the whole file.
    const auctionIdCache = new Map();
    const resolveAuctionId = (ano) => {
      const key = String(ano || '').trim();
      if (!key) return null;
      if (auctionIdCache.has(key)) return auctionIdCache.get(key);
      const row = db.get('SELECT id FROM auctions WHERE ano = ? LIMIT 1', [key]);
      const id  = row ? row.id : null;
      auctionIdCache.set(key, id);
      return id;
    };
    const auctionIdSlot = def.fields.indexOf('auction_id');

    // Snapshot module defaults once per request so each row can pick
    // them up without re-calling the function.
    const moduleDefaults = (typeof def.defaults === 'function')
      ? (() => { try { return def.defaults(db) || {}; } catch (_) { return {}; } })()
      : {};

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      try {
        // Effective per-row defaults: flat module defaults overlaid with
        // per-row rowDefaults (e.g. state ← KERALA when PLACE is ASP-
        // prefixed). Used for both dedup key resolution and the positional
        // value fill so the stored row and the dup query agree.
        const rowEff = Object.assign({}, moduleDefaults,
          (typeof def.rowDefaults === 'function' ? (() => { try { return def.rowDefaults(r, mapping, db) || {}; } catch (_) { return {}; } })() : {}));
        // Duplicate detection — resolve each key column to its effective
        // value: the mapped source cell, falling back to the per-row
        // default when the source is blank. Dedup only fires when EVERY
        // key column resolves to a non-blank value, keyed identically to
        // /verify so counts agree.
        const keyVals = def.keyCols.map(k => {
          let v = mapping[k] ? r[mapping[k]] : '';
          if ((v == null || String(v).trim() === '') && (k in rowEff)) v = rowEff[k];
          return v;
        });
        if (keyVals.every(v => v != null && String(v).trim() !== '')) {
          const whereSql = def.keyCols.map(k => `${k} = ?`).join(' AND ');
          const dup = db.get(`SELECT id FROM ${def.table} WHERE ${whereSql} LIMIT 1`, keyVals);
          if (dup) {
            skipped++;
            // Record which file row was skipped + the matching key values
            // + the existing DB row id so the UI can list them (capped).
            if (skippedDetails.length < 500) {
              const keyObj = {};
              def.keyCols.forEach((k, idx) => { keyObj[k] = keyVals[idx]; });
              skippedDetails.push({
                row: i + 2,
                keys: keyObj,
                existingId: (dup && dup.id != null) ? dup.id : null,
                reason: 'Duplicate — already in ' + def.table + ' ('
                      + def.keyCols.map((k, idx) => k + '=' + keyVals[idx]).join(', ') + ')',
              });
            }
            continue;
          }
        }

        // Build positional values from the source mapping. `date` is
        // normalised to ISO yyyy-mm-dd so downstream code (Tally XML,
        // date BETWEEN filters) sees a canonical value regardless of
        // spreadsheet format (Excel serial, dd-mm-yyyy, etc.).
        const values = fieldSources.map(([fname, src]) => {
          const v = src ? r[src] : '';
          if (fname === 'date') return normalizeDate(v);
          return v;
        });

        // Apply effective per-row defaults to any positional slot whose
        // value came back blank. e.g. state ← KERALA for ASP-prefixed
        // places, or the flat business-state / auctions.mode default.
        for (let s = 0; s < def.fields.length; s++) {
          const fname = def.fields[s];
          if (!(fname in rowEff)) continue;
          const cur = values[s];
          if (cur == null || String(cur).trim() === '') {
            values[s] = rowEff[fname];
          }
        }

        if (def.autoFillAuctionId && auctionIdSlot >= 0) {
          const anoSrc = mapping.ano;
          const anoVal = anoSrc ? r[anoSrc] : '';
          const aid    = resolveAuctionId(anoVal);
          if (aid == null) {
            failed++;
            if (errors.length < 50) errors.push({
              row: i + 2,
              error: `No trade found for ano="${String(anoVal || '').trim()}" — create the auction first or fix the column.`
            });
            continue;
          }
          values[auctionIdSlot] = aid;
        }
        if (!dryRun) {
          const info = db.run(insertSql, values);
          if (info && info.lastInsertRowid != null) insertedIds.push(Number(info.lastInsertRowid));
        }
        imported++;
      } catch (e) {
        failed++;
        if (errors.length < 50) errors.push({ row: i + 2, error: e.message });
      }
    }
  } catch (e) {
    fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: e.message });
  }

  // Back-fill auction_id on any pre-existing rows the user imported
  // before this fix landed. Idempotent — touches only NULL slots with
  // a matching auctions.ano. Runs even on dryRun so the fix applies
  // without a second pass.
  if (def.autoFillAuctionId) {
    try {
      getDb().run(
        `UPDATE ${def.table}
            SET auction_id = (SELECT id FROM auctions WHERE auctions.ano = ${def.table}.ano)
          WHERE auction_id IS NULL
            AND ano IS NOT NULL AND ano != ''
            AND EXISTS (SELECT 1 FROM auctions WHERE auctions.ano = ${def.table}.ano)`
      );
    } catch (_) { /* non-fatal */ }
  }

  // Repair any malformed `date` values from earlier imports (before
  // per-row normalize was added). Idempotent and cheap.
  if (!dryRun) {
    try { repairBadDates(db); } catch (_) { /* non-fatal */ }
  }

  // Log this run regardless of outcome. Dry-runs get an empty
  // inserted_ids so the Undo button stays disabled for that entry.
  let importLogId = null;
  try {
    const info = db.run(`INSERT INTO import_log
      (module, filename, dry_run, total, imported, skipped, failed, errors, inserted_ids, user_id, username)
      VALUES (?,?,?,?,?,?,?,?,?,?,?)`,
      [moduleKey, req.file.originalname || '', dryRun ? 1 : 0,
       total, imported, skipped, failed, JSON.stringify(errors).slice(0, 4000),
       dryRun ? '' : JSON.stringify(insertedIds),
       (req.user && req.user.id) || null, (req.user && req.user.username) || '']);
    if (info && info.lastInsertRowid != null) importLogId = Number(info.lastInsertRowid);
  } catch (e) {
    console.error('[import-old-data] Failed to write import_log entry:', e && e.message ? e.message : e);
  }

  fs.unlink(req.file.path, () => {});
  res.json({
    success: true, module: moduleKey, dryRun,
    total, imported, skipped, failed,
    // Per-row detail of duplicate-skipped rows (capped at 500) so the UI
    // can show exactly which rows were skipped and what they collided with.
    skippedDetails,
    errors, importLogId,
  });
});

// GET /history — recent imports for the History panel. Cache-busted
// so the panel always shows fresh state after an import lands.
app.get('/api/import-old-data/history', requireAdmin, (req, res) => {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
  const rows = getDb().all(
    `SELECT id, module, filename, dry_run, total, imported, skipped, failed,
            errors, inserted_ids, undone_at, username, created_at
       FROM import_log ORDER BY id DESC LIMIT 200`
  );
  res.json(rows.map(r => {
    const ids = r.inserted_ids ? _safeImportJSON(r.inserted_ids) : [];
    return {
      id: r.id, module: r.module, filename: r.filename,
      dry_run: !!r.dry_run, total: r.total, imported: r.imported,
      skipped: r.skipped, failed: r.failed,
      errors: r.errors ? _safeImportJSON(r.errors) : [],
      undone_at: r.undone_at || '',
      // Undo is available iff this wasn't a dry-run, actually inserted
      // rows, and hasn't already been rolled back.
      undoable: !r.dry_run && Array.isArray(ids) && ids.length > 0 && !r.undone_at,
      inserted_count: Array.isArray(ids) ? ids.length : 0,
      username: r.username, created_at: r.created_at,
    };
  }));
});

// POST /undo/:id — roll back a single live import. DELETEs every row
// captured in import_log.inserted_ids after snapshotting the DB so
// the rollback itself is reversible. Marks the entry as undone so a
// second click is a no-op.
app.post('/api/import-old-data/undo/:id', requireAdmin, (req, res) => {
  const db = getDb();
  const logId = Number(req.params.id);
  if (!Number.isFinite(logId)) return res.status(400).json({ error: 'Invalid import id' });
  const logRow = db.get('SELECT * FROM import_log WHERE id = ?', [logId]);
  if (!logRow) return res.status(404).json({ error: 'Import not found' });
  if (logRow.undone_at) return res.status(400).json({ error: 'This import has already been undone at ' + logRow.undone_at });
  if (logRow.dry_run)   return res.status(400).json({ error: 'Dry-run imports did not insert any rows — nothing to undo' });

  const def = IMPORT_MODULES[logRow.module];
  if (!def) return res.status(400).json({ error: 'Unknown module on this import — cannot resolve target table' });

  let ids = [];
  try { ids = JSON.parse(logRow.inserted_ids || '[]'); } catch (_) {}
  if (!Array.isArray(ids) || !ids.length) {
    return res.status(400).json({
      error: 'This import did not record its inserted row IDs (was it run before per-import Undo was added?). To clear it, use the Backup tab → Delete All for ' + def.table + '.',
    });
  }

  // Snapshot first — undo of an undo is a Restore from the file we drop
  // at this path. The existing snapshotDbToFile helper handles the
  // backups/ directory + filename stamp.
  const backupPath = snapshotDbToFile('undo-import-' + logRow.module + '-' + logId);
  if (!backupPath) {
    return res.status(500).json({ error: 'Backup snapshot failed; refusing to undo. Check disk space and try again.' });
  }

  // Bulk-delete by id, chunked. SQLite caps parameter count per
  // statement (default 999), so chunk for large imports.
  let deleted = 0;
  const CHUNK = 500;
  try {
    for (let i = 0; i < ids.length; i += CHUNK) {
      const slice = ids.slice(i, i + CHUNK).filter(n => Number.isFinite(Number(n)));
      if (!slice.length) continue;
      const placeholders = slice.map(() => '?').join(',');
      const info = db.run(`DELETE FROM ${def.table} WHERE id IN (${placeholders})`, slice);
      if (info && typeof info.changes === 'number') deleted += info.changes;
    }
    db.run('UPDATE import_log SET undone_at = datetime("now","localtime") WHERE id = ?', [logId]);
  } catch (e) {
    return res.status(500).json({ error: 'Undo failed mid-way; partial deletions may have occurred. Backup at: ' + backupPath + ' — ' + (e.message || e) });
  }

  res.json({
    success: true,
    importLogId: logId,
    module: logRow.module,
    table: def.table,
    requested: ids.length,
    deleted,
    backupPath,
  });
});

// Helper — JSON.parse that returns [] on malformed input. Used by
// /history to defensively unpack the errors / inserted_ids columns.
function _safeImportJSON(s) {
  try { return JSON.parse(s); } catch (_) { return []; }
}

// ══════════════════════════════════════════════════════════════
// REPORTS HUB
// ══════════════════════════════════════════════════════════════
// Three endpoints that back the Reports tab in the Admin Console:
//   • Per-trade summary (in-tab data table + KPIs)
//   • Lifetime branch comparison (chart-first card)
//   • Trade summary PDF (clickable "Open PDF" button on the card)
//
// The PDF endpoint accepts ?token= because window.open() can't set
// an Authorization header — the token-query path is wired explicitly
// here rather than poking requireAuth, so the relaxation is confined
// to a single route.

app.get('/api/reports/trade-summary/:id', requireView, (req, res) => {
  const id = parseInt(req.params.id, 10);
  if (!id) return res.status(400).json({ error: 'Invalid auction id' });
  try {
    const data = getTradeSummary(getDb(), id, req.query.branch || '');
    res.json(data);
  } catch (e) {
    if (e && e.status === 404) return res.status(404).json({ error: e.message });
    console.error('trade-summary error:', e);
    res.status(500).json({ error: e.message || 'Failed to build summary' });
  }
});

app.get('/api/reports/branch-comparison', requireView, (req, res) => {
  try {
    const data = getBranchComparison(getDb());
    res.json(data);
  } catch (e) {
    console.error('branch-comparison error:', e);
    res.status(500).json({ error: e.message || 'Failed to build comparison' });
  }
});

// Summary PDF — window.open() can't set Authorization, so we accept
// the token via the querystring. Validates the token by hand because
// the framework's requireView only reads the header.
app.get('/api/reports/summary-pdf/:id', (req, res, next) => {
  const headerTok = (req.headers.authorization || '').replace('Bearer ', '');
  const queryTok  = String(req.query.token || '');
  const token = headerTok || queryTok;
  if (!token) return res.status(401).json({ error: 'No token' });
  const db = getDb();
  const session = db.get('SELECT * FROM sessions WHERE token = ?', [token]);
  if (!session) return res.status(401).json({ error: 'Session expired — please sign in again' });
  const user = db.get('SELECT * FROM users WHERE id = ?', [session.user_id]);
  if (!user) return res.status(401).json({ error: 'Unauthorized' });
  if (!userHas(user.role, 'view')) {
    return res.status(403).json({ error: 'Your role does not allow viewing reports', role: user.role });
  }
  // Touch last_used_at — same housekeeping the normal requireAuth does.
  db.run(`UPDATE sessions SET last_used_at = datetime('now','localtime') WHERE token = ?`, [token]);
  req.user = user;
  req.session = session;
  next();
}, async (req, res) => {
  const id = parseInt(req.params.id, 10);
  if (!id) return res.status(400).json({ error: 'Invalid auction id' });
  try {
    const db = getDb();
    const pdf = await generateTradeSummaryPDF(db, id, req.query.branch || '');
    // Filename noun follows business_mode (matches the PDF title +
    // on-screen card) — read the setting inline rather than exporting
    // reports.js's private helper.
    let nounMode = 'e-Auction';
    try {
      const r = db.get(`SELECT value FROM company_settings WHERE key = 'business_mode'`);
      if (r && r.value) nounMode = String(r.value);
    } catch (_) {}
    const noun = (nounMode === 'e-Trade') ? 'Trade' : 'Auction';
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="${noun}Summary_${id}.pdf"`);
    res.send(pdf);
  } catch (e) {
    if (e && e.status === 404) return res.status(404).json({ error: e.message });
    console.error('summary-pdf error:', e);
    res.status(500).json({ error: e.message || 'PDF generation failed' });
  }
});

// ══════════════════════════════════════════════════════════════
// START
// ══════════════════════════════════════════════════════════════
function repairBadDates(db) {
  // Fix rows where date is an Excel serial number stored as string, or Date-object-toString garbage
  const tables = ['auctions', 'bills', 'debit_notes', 'invoices', 'purchases', 'lots'];
  let totalFixed = 0;
  for (const tbl of tables) {
    try {
      // Only tables that have a `date` column
      const hasDate = db.all(`PRAGMA table_info(${tbl})`).some(c => c.name === 'date');
      if (!hasDate) continue;
      const rows = db.all(`SELECT rowid, date FROM ${tbl} WHERE date IS NOT NULL AND date != ''`);
      let fixed = 0;
      for (const r of rows) {
        const current = String(r.date);
        // Skip if already ISO yyyy-mm-dd
        if (/^\d{4}-\d{2}-\d{2}$/.test(current)) continue;
        const iso = normalizeDate(r.date);
        if (iso && /^\d{4}-\d{2}-\d{2}$/.test(iso) && iso !== current) {
          db.run(`UPDATE ${tbl} SET date = ? WHERE rowid = ?`, [iso, r.rowid]);
          fixed++;
        }
      }
      if (fixed > 0) console.log(`  Date repair: ${tbl} — fixed ${fixed} row(s)`);
      totalFixed += fixed;
    } catch (_) { /* table may not exist yet */ }
  }
  if (totalFixed > 0) console.log(`  Date repair: ${totalFixed} total row(s) normalized to yyyy-mm-dd`);
}

const PORT = process.env.PORT || 3001;
(async () => {
  const db = await initDb();
  initCompanySettings(db);
  repairBadDates(db);
  // Bootstrap the per-install license row on first boot. This generates
  // the install_id, starts the trial window, and logs the current
  // status so the operator can spot expiry-soon at deploy time.
  try {
    const lstatus = license.getStatus(db);
    const tag = lstatus.expired
      ? `EXPIRED (was ${lstatus.expires_at})`
      : `${lstatus.days_remaining} day${lstatus.days_remaining === 1 ? '' : 's'} remaining (expires ${lstatus.expires_at})`;
    console.log(`  License: install ${lstatus.install_id} — ${tag}`);
  } catch (e) {
    console.warn('  License bootstrap failed:', e.message);
  }
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n  Spice Config running at http://localhost:${PORT}\n`);
  });
})();
