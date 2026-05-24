const express = require('express');
const crypto = require('crypto');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { initDb, getDb, flushDb, replaceDbFromBuffer, DB_PATH } = require('./db');
const { initCompanySettings, CATEGORIES, getAllSettings, updateSettings, getSettingsFlat, getGSTRates, getAllPresets, setActivePresetCode, savePreset, getActivePresetCode } = require('./company-config');
const { calculateLot, buildSalesInvoice, buildPurchaseInvoice, buildAgriBill, buildDebitNote, listAgriSellers, getPaymentSummary, getBankPaymentData, getTDSReturnData, getSalesJournal, getPurchaseJournal } = require('./calculations');
const { generatePurchaseInvoicePDF, generateCropReceiptPDF, generateAgriBillPDF, generateSalesInvoicePDF, generateSalesInvoicesBatchPDF, generatePurchaseInvoicesBatchPDF, generateAgriBillsBatchPDF } = require('./invoice-pdf');
const { generateDebitNoteBatchPDF } = require('./debit-note-print');
const { EXPORT_TYPES } = require('./exports');
const { exportPdf: exportAnyPdf } = require('./exports-pdf');
const { DBF_EXPORTS } = require('./dbf-exports');
const { REPORTS: LORRY_REPORTS } = require('./lorry-reports');
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

// File upload setup
// Honor SPICE_DATA_DIR so uploads also land in userData when packaged.
const uploadDir = path.join(process.env.SPICE_DATA_DIR || path.join(__dirname, 'data'), 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
const upload = multer({ dest: uploadDir, limits: { fileSize: 20 * 1024 * 1024 } });

const hash = pw => crypto.createHash('sha256').update(pw).digest('hex');

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
  if (!session) return res.status(403).json({ error: 'Session expired — please sign in again' });
  const user = db.get('SELECT * FROM users WHERE id = ?', [session.user_id]);
  if (!user) return res.status(403).json({ error: 'Unauthorized' });
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
    'view',         // read any list / detail
    'export',       // download XLSX / PDF / CSV exports
    'self_password' // change own password
  ]),
  operator: new Set([
    'view', 'export', 'self_password',
    'lot_write',     // create/edit lots, calculate, validate, price-import
    'invoice_write', // generate sales/purchase/bills + edit
    'trader_write',  // create/edit/delete-bank traders
    'buyer_write'    // create/edit buyers (per user decision: tax fields editable)
  ]),
  manager: new Set([
    'view', 'export', 'self_password',
    'lot_write', 'invoice_write', 'trader_write', 'buyer_write',
    'auction_write',  // create/edit auctions (trades)
    'invoice_revert', // revert sales/purchase/bills (undo invoice)
    'settings_write', // edit company settings (rates, addresses, flags)
    'state_toggle'    // toggle business state TN ↔ KL
  ]),
  admin: new Set([
    'view', 'export', 'self_password',
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

// Convenience aliases — readable names for the most common gate points.
// Encapsulates the permission name so callers don't repeat string literals.
const requireView          = requirePermission('view');
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
app.post('/api/login', (req, res) => {
  const { username, password, device_label } = req.body;
  if (!username || !password) return res.status(400).json({ error: 'Missing credentials' });
  const db = getDb();
  const user = db.get('SELECT * FROM users WHERE username = ?', [username]);
  if (!user || user.password_hash !== hash(password)) return res.status(401).json({ error: 'Invalid credentials' });
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

app.post('/api/users', requireUserManage, (req, res) => {
  const { username, password, role } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: 'username and password are required' });
  if (password.length < 4) return res.status(400).json({ error: 'Password must be at least 4 characters' });
  if (!/^[a-zA-Z0-9_.-]{3,30}$/.test(username)) return res.status(400).json({ error: 'Username: 3–30 chars, letters/digits/._- only' });
  // Validate role against the known set. Default to 'operator' (the most
  // common and least privileged write-capable role) if missing or invalid.
  // Legacy 'user' role is mapped to 'viewer' for backward compat.
  const VALID_ROLES = ['viewer', 'operator', 'manager', 'admin'];
  let finalRole = (role || '').toLowerCase();
  if (finalRole === 'user') finalRole = 'viewer';
  if (!VALID_ROLES.includes(finalRole)) finalRole = 'operator';
  const db = getDb();
  const existing = db.get('SELECT id FROM users WHERE username = ?', [username]);
  if (existing) return res.status(400).json({ error: 'Username already exists' });
  db.run(
    'INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)',
    [username, hash(password), finalRole]
  );
  const created = db.get('SELECT id, username, role FROM users WHERE username = ?', [username]);
  res.json({ success: true, id: created ? created.id : null, username, role: finalRole });
});

// Update an existing user's role (for promoting/demoting users without
// recreating them). Admin-only — same gate as creating users.
app.put('/api/users/:id/role', requireUserManage, (req, res) => {
  const { role } = req.body || {};
  const VALID_ROLES = ['viewer', 'operator', 'manager', 'admin'];
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

app.put('/api/users/:id/password', requireUserManage, (req, res) => {
  const { password } = req.body || {};
  if (!password || password.length < 4) return res.status(400).json({ error: 'Password must be at least 4 characters' });
  const db = getDb();
  const user = db.get('SELECT id, username FROM users WHERE id = ?', [req.params.id]);
  if (!user) return res.status(404).json({ error: 'User not found' });
  db.run('UPDATE users SET password_hash = ? WHERE id = ?', [hash(password), user.id]);
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
app.put('/api/me/password', requirePermission('self_password'), (req, res) => {
  const { current_password, new_password } = req.body || {};
  if (!current_password || !new_password) return res.status(400).json({ error: 'Both current and new password required' });
  if (new_password.length < 4) return res.status(400).json({ error: 'New password must be at least 4 characters' });
  const db = getDb();
  const user = db.get('SELECT * FROM users WHERE id = ?', [req.user.id]);
  if (!user || user.password_hash !== hash(current_password)) return res.status(401).json({ error: 'Current password is incorrect' });
  db.run('UPDATE users SET password_hash = ? WHERE id = ?', [hash(new_password), user.id]);
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
        source:   'live'
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
      `SELECT trader_id, bank_name, acctnum, ifsc, holder_name
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
app.get('/api/traders/:id', requireView, (req, res) => {
  const db = getDb();
  const row = db.get('SELECT * FROM traders WHERE id = ?', [req.params.id]);
  if (!row) return res.status(404).json({ error: 'Not found' });
  // Attach banks array so the edit modal sees all bank accounts.
  row.banks = db.all(
    'SELECT trader_id, bank_name, acctnum, ifsc, holder_name FROM trader_banks WHERE trader_id = ? ORDER BY id',
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
      'INSERT INTO trader_banks (trader_id, bank_name, acctnum, ifsc, holder_name) VALUES (?,?,?,?,?)',
      [traderId, b.bank_name||'', String(b.acctnum||''), String(b.ifsc||''), b.holder_name||'']
    );
  }
  // Mirror first bank into traders row for legacy compatibility
  const first = arr[0] || {};
  db.run(
    'UPDATE traders SET ifsc=?, acctnum=?, holder_name=? WHERE id=?',
    [first.ifsc||'', first.acctnum||'', first.holder_name||'', traderId]
  );
}

app.post('/api/traders', requireTraderWrite, (req, res) => {
  const t = req.body;
  const db = getDb();
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
  res.json(db.all('SELECT * FROM buyers ORDER BY buyer1 LIMIT 500'));
});
app.post('/api/buyers', requireBuyerWrite, (req, res) => {
  const b = req.body;
  getDb().run(`INSERT INTO buyers (
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
  getDb().run(`UPDATE buyers SET
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
  db.run('DELETE FROM lots WHERE auction_id = ?', [req.params.id]);
  db.run('DELETE FROM auctions WHERE id = ?', [req.params.id]);
  res.json({ success: true });
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
  const { branch, name, buyer } = req.query;
  // Correlated subquery (not LEFT JOIN) to avoid any risk of row duplication
  // if the same buyer code exists multiple times in the buyers table.
  let q = `SELECT lots.*,
             (SELECT b.code FROM buyers b WHERE b.buyer = lots.buyer LIMIT 1) AS buyer_code
           FROM lots
           WHERE lots.auction_id = ?`;
  const p = [req.params.auctionId];
  if (branch) { q += ' AND lots.branch = ?'; p.push(branch); }
  if (name) { q += ' AND lots.name LIKE ?'; p.push(`%${name}%`); }
  if (buyer) { q += ' AND lots.buyer = ?'; p.push(buyer); }
  q += ' ORDER BY lots.lot_no';
  res.json(getDb().all(q, p));
});

app.post('/api/lots', requireLotWrite, (req, res) => {
  const l = req.body;
  getDb().run(`INSERT INTO lots (auction_id,lot_no,crop,grade,crpt,branch,state,trader_id,name,padd,ppla,ppin,pstate,pst_code,cr,pan,tel,aadhar,bags,litre,qty,gross_wt,sample_wt,moisture,user_id)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [l.auction_id,l.lot_no,l.crop||'',l.grade||'',l.crpt||'',l.branch||'',l.state||'TAMIL NADU',l.trader_id||null,l.name||'',l.padd||'',l.ppla||'',l.ppin||'',l.pstate||'',l.pst_code||'',l.cr||'',l.pan||'',l.tel||'',l.aadhar||'',l.bags||0,l.litre||'',l.qty||0,l.gross_wt||0,l.sample_wt||0,l.moisture||'',l.user_id||'']);
  res.json({ success: true });
});

app.put('/api/lots/:id', requireLotWrite, (req, res) => {
  const l = req.body; const sets = []; const vals = [];
  for (const [k,v] of Object.entries(l)) {
    if (k !== 'id' && k !== 'auction_id' && k !== 'created_at') { sets.push(`${k}=?`); vals.push(v); }
  }
  vals.push(req.params.id);
  getDb().run(`UPDATE lots SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

app.delete('/api/lots/:id', requireDelete, (req, res) => {
  getDb().run('DELETE FROM lots WHERE id = ?', [req.params.id]); res.json({ success: true });
});

// ── Calculate all lots for an auction (GENERATE.PRG) ─────────
app.post('/api/lots/calculate/:auctionId', requireLotWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const lots = db.all('SELECT * FROM lots WHERE auction_id = ? AND amount > 0', [req.params.auctionId]);
  let count = 0;
  for (const lot of lots) {
    const calc = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [calc.pqty,calc.prate,calc.puramt,calc.com,calc.sertax,calc.cgst,calc.sgst,calc.igst,calc.advance,calc.balance,calc.bilamt,calc.refund||0,calc.refud||0,calc.isp_pqty||0,calc.isp_prate||0,calc.isp_puramt||0,calc.asp_pqty||0,calc.asp_prate||0,calc.asp_puramt||0,lot.id]);
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
    const calc = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [calc.pqty,calc.prate,calc.puramt,calc.com,calc.sertax,calc.cgst,calc.sgst,calc.igst,calc.advance,calc.balance,calc.bilamt,calc.refund||0,calc.refud||0,calc.isp_pqty||0,calc.isp_prate||0,calc.isp_puramt||0,calc.asp_pqty||0,calc.asp_prate||0,calc.asp_puramt||0,lot.id]);
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
// INVOICES — Sales (GSTIN.PRG / KGSTIN.PRG)
// ══════════════════════════════════════════════════════════════
app.get('/api/invoices', requireView, (req, res) => {
  const { ano, auction_id, from, to } = req.query;
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

app.post('/api/invoices/generate/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
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
  db.run(`INSERT INTO invoices (auction_id,ano,date,state,sale,invo,buyer,buyer1,gstin,place,bag,qty,amount,gunny,pava_hc,ins,cgst,sgst,igst,tcs,rund,tot)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [req.params.auctionId,auction.ano,auction.date,invoiceState,saleType,String(invoiceNo),buyerCode,invoice.buyer.buyer1||'',
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
      db.run(`INSERT INTO invoices (auction_id,ano,date,state,sale,invo,buyer,buyer1,gstin,place,bag,qty,amount,gunny,pava_hc,ins,cgst,sgst,igst,tcs,rund,tot)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [req.params.auctionId,auction.ano,auction.date,invoiceState,useSaleType,invoNo,row.buyer,invoice.buyer.buyer1||'',
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
      results.push({ buyer: row.buyer, invoiceNo: invoNo, sale: useSaleType, grandTotal: s.grandTotal });
      nextNo++;
    } catch (e) { errors.push({ buyer: row.buyer, error: e.message }); }
  }
  
  res.json({ success: true, generated: results.length, results, errors });
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
      ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg)
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
      ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg)
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
        ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg)
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
        ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg)
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
  const { auction_id, ano, from, to } = req.query;
  const db = getDb();
  let q = 'SELECT * FROM purchases WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  const mw = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=purchases.auction_id)');
  q += mw.sql; p.push(...mw.params);
  q += ' ORDER BY date DESC LIMIT 500';
  res.json(db.all(q, p));
});

app.post('/api/purchases/generate/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
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
  const { auction_id, ano, from, to } = req.query;
  const db = getDb();
  let q = 'SELECT * FROM bills WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  const mw = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=bills.auction_id)');
  q += mw.sql; p.push(...mw.params);
  q += ' ORDER BY date DESC, bil DESC LIMIT 500';
  res.json(withFmtDate(db.all(q, p)));
});

// Generate agri bill for a seller
app.post('/api/bills/generate/:auctionId', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
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
  const { auction_id, ano, from, to } = req.query;
  const db = getDb();
  let q = 'SELECT * FROM debit_notes WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  const mw = modeWhereClause(db, '(SELECT mode FROM auctions WHERE id=debit_notes.auction_id)');
  q += mw.sql; p.push(...mw.params);
  q += ' ORDER BY date DESC, note_no DESC LIMIT 500';
  res.json(withFmtDate(db.all(q, p)));
});

app.post('/api/debit-notes/generate', requireInvoiceWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { invoiceNo, saleType, discount, noteNo } = req.body;
  
  if (!invoiceNo || !saleType || !discount || !noteNo) {
    return res.status(400).json({ error: 'invoiceNo, saleType, discount, and noteNo required' });
  }
  
  const note = buildDebitNote(db, invoiceNo, saleType, parseFloat(discount), cfg);
  if (!note) return res.status(404).json({ error: `Invoice ${invoiceNo} (${saleType}) not found` });

  // Use the invoice's date (= auction date) as the debit note's date.
  // This is what every downstream consumer assumes — the PDF generator
  // and XML builder both join debit_notes back to auctions via this
  // column. Falls back to today only if the invoice somehow has no
  // date stored (rare; legacy data).
  const noteDate = note.invoice.date || new Date().toISOString().slice(0, 10);

  db.run(`INSERT INTO debit_notes (ano,date,state,name,note_no,amount,cgst,sgst,igst,total)
    VALUES (?,?,?,?,?,?,?,?,?,?)`,
    [note.invoice.ano, noteDate, note.invoice.state||'',
     note.invoice.buyer1||note.invoice.buyer, String(noteNo),
     note.amount, note.cgst, note.sgst, note.igst, note.total]);
  
  res.json({ success: true, note });
});

app.delete('/api/debit-notes/:id', requireDelete, (req, res) => {
  getDb().run('DELETE FROM debit_notes WHERE id = ?', [req.params.id]);
  res.json({ success: true });
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
    if (exportDef.needsCfg) {
      const cfg = getSettingsFlat(db);
      // Pass state too so exports that need both (e.g. Praman) can filter
      // by state without losing cfg context. Backward-compatible: existing
      // needsCfg exports that ignore the 4th arg are unaffected.
      buffer = await exportDef.fn(db, auctionId, cfg, req.query.state);
    } else {
      buffer = await exportDef.fn(db, auctionId, req.query.state);
    }
    // Per-export-type content-type/extension override (defaults to xlsx).
    // CSV exports like Praman use ext:'csv', mime:'text/csv'.
    const ext  = exportDef.ext  || 'xlsx';
    const mime = exportDef.mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    res.setHeader('Content-Type', mime);
    res.setHeader('Content-Disposition', `attachment; filename="${exportDef.name}_${auctionId}.${ext}"`);
    res.send(Buffer.from(buffer));
  } catch(e) {
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

  // ── Cumulative totals across ALL trades (lifetime) ──
  // Aggregates over every lot in every (mode-matching) auction.
  const cumRow = db.get(
    `SELECT COALESCE(SUM(qty),0) as qty,
            COALESCE(SUM(amount),0) as amount,
            COUNT(*) as lots
     FROM lots ${_w(mwLots)}`,
    mwLots.params
  ) || {};
  const cumulative = {
    qty:    cumRow.qty    || 0,
    amount: cumRow.amount || 0,
    lots:   cumRow.lots   || 0,
    auctions: counts.auctions,
  };

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
            COALESCE(SUM(CASE WHEN l.invo IS NOT NULL AND l.invo != '' THEN 1 ELSE 0 END),0) as invoiced
     FROM auctions a
     LEFT JOIN lots l ON l.auction_id = a.id
     ${mwAa.sql ? 'WHERE 1=1' + mwAa.sql : ''}
     GROUP BY a.id, a.ano, a.date, a.crop_type, a.mode
     ORDER BY a.date DESC, a.id DESC
     LIMIT 50`,
    mwAa.params
  );

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
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n  Spice Config running at http://localhost:${PORT}\n`);
  });
})();
