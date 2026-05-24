// Load env vars from .env (if present) BEFORE any other require runs.
// Anything downstream — db.js, mobile-bridge, the tenant-preset admin
// gate, the Anthropic API client — can then read process.env.<KEY>
// without caring whether the value came from .env, the shell, or
// Railway's variables panel. The {silent} option means a missing
// .env file is fine (production on Railway uses dashboard vars, not
// a checked-in .env).
try { require('dotenv').config(); } catch (_) { /* dotenv not installed — fine */ }

const express = require('express');
const crypto = require('crypto');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const bcrypt = require('bcryptjs');
const rateLimit = require('express-rate-limit');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const { initDb, getDb, DB_PATH, replaceFromBuffer } = require('./db');
const { initCompanySettings, CATEGORIES, getSetting, getAllSettings, updateSettings, getSettingsFlat, getGSTRates } = require('./company-config');
const { calculateLot, buildSalesInvoice, buildPurchaseInvoice, buildAgriBill, listAgriSellers, getPaymentSummary, getBankPaymentData, getTDSReturnData, getSalesJournal, getPurchaseJournal, round2, round0 } = require('./calculations');
const { generatePurchaseInvoicePDF, generateCropReceiptPDF, generateAgriBillPDF, generateSalesInvoicePDF, generateSalesInvoicesBatchPDF, generatePurchaseInvoicesBatchPDF, generateAgriBillsBatchPDF } = require('./invoice-pdf');
const { EXPORT_TYPES, createExcelBuffer } = require('./exports');
const { getCompanyHeader, writeXlsxCompanyHeader } = require('./report-formatters');
const { exportPdf: exportAnyPdf } = require('./exports-pdf');
const { DBF_EXPORTS } = require('./dbf-exports');
const { REPORTS: LORRY_REPORTS } = require('./lorry-reports');
// Defensive resolution — see _company-identity-fallback.js. Uses the
// real getCompanyIdentity from report-formatters.js when available,
// falls through to an inline fallback otherwise. Fixes
// "getCompanyIdentity is not a function" on partial deploys.
const getCompanyIdentity = require('./_company-identity-fallback').resolve();
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

// When the user hasn't uploaded a custom logo, /logo-ispl.png falls
// through to the bundled default (logo_kj.png). The upload widget still
// writes to logo-ispl.png, and GET /api/company-settings/logo/ispl still
// correctly returns exists:false until the user uploads, so the "Upload
// your logo" UI state stays accurate.
app.get('/logo-ispl.png', (req, res, next) => {
  const userLogo = path.join(__dirname, 'public', 'logo-ispl.png');
  if (fs.existsSync(userLogo)) return next();
  res.sendFile(path.join(__dirname, 'public', 'logo_kj.png'));
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

// ── Password hashing ────────────────────────────────────────────────
// Bcrypt cost factor. 12 ≈ 250ms per hash on modern hardware — slow enough
// to make brute-force expensive but fast enough that login latency is fine.
const BCRYPT_ROUNDS = 12;

// Pre-computed dummy bcrypt hash for unknown-user logins. Running a real
// bcrypt comparison against this when the username isn't found prevents
// timing-based user enumeration (otherwise unknown-user paths return in
// microseconds while real-user paths take ~250ms).
const DUMMY_BCRYPT_HASH = bcrypt.hashSync('__never_a_real_password__', BCRYPT_ROUNDS);

// Hash a plaintext password with bcrypt. Async — every callsite is in a
// route handler that can be made async.
async function hashPassword(plain) {
  return bcrypt.hash(String(plain), BCRYPT_ROUNDS);
}

// Detect legacy SHA-256 password hashes (64 lowercase hex chars). Bcrypt
// hashes start with "$2a$" / "$2b$" / "$2y$" so the two are unambiguous.
const LEGACY_SHA256_RE = /^[a-f0-9]{64}$/i;
function isLegacyHash(stored) { return LEGACY_SHA256_RE.test(stored || ''); }

// Verify a plaintext password against a stored hash, supporting both new
// bcrypt rows and legacy SHA-256 rows from before the migration. Returns
// true/false; never throws on bad input.
async function verifyPassword(plain, stored) {
  if (!stored || typeof stored !== 'string') return false;
  if (isLegacyHash(stored)) {
    const sha = crypto.createHash('sha256').update(String(plain)).digest('hex');
    return sha === stored;
  }
  try { return await bcrypt.compare(String(plain), stored); }
  catch { return false; }
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

function addDays(isoDate, days) {
  const iso = normalizeDate(isoDate);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return '';
  const [y, m, d] = iso.split('-').map(Number);
  const ms = Date.UTC(y, m - 1, d) + (Number(days) || 0) * 86400 * 1000;
  const out = new Date(ms);
  const yy = out.getUTCFullYear();
  const mm = String(out.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(out.getUTCDate()).padStart(2, '0');
  return `${yy}-${mm}-${dd}`;
}

const { fmtDate, todayLocalISO, invalidateDateFormatCache } = require('./date-format');

function withFmtDate(rows, field = 'date') {
  if (!Array.isArray(rows)) return rows;
  return rows.map(r => ({ ...r, date_fmt: fmtDate(r[field]) }));
}

const FORCED_CHANGE_ALLOWED = new Set([
  '/api/me', '/api/me/password',
  '/api/auth/me', '/api/auth/change-password', '/api/auth/logout',
]);

function requireAuth(req, res, next) {
  const token = (req.headers.authorization || '').replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'No token' });
  const db = getDb();
  const session = db.get(
    `SELECT * FROM sessions
     WHERE token = ?
       AND (expires_at IS NULL OR expires_at >= datetime('now','localtime'))`,
    [token]
  );
  if (!session) {
    db.run('DELETE FROM sessions WHERE token = ?', [token]);
    return res.status(403).json({ error: 'Session expired — please sign in again' });
  }
  const user = db.get('SELECT * FROM users WHERE id = ?', [session.user_id]);
  if (!user) return res.status(403).json({ error: 'Unauthorized' });
  if (user.must_change_password && !FORCED_CHANGE_ALLOWED.has(req.path)) {
    return res.status(403).json({
      error: 'Password change required before continuing',
      must_change_password: true
    });
  }
  db.run(`UPDATE sessions SET last_used_at = datetime('now','localtime') WHERE token = ?`, [token]);
  req.user = user;
  req.session = session;
  next();
}

function requireAdmin(req, res, next) {
  requireAuth(req, res, (err) => {
    if (err) return next(err);
    if (!req.user || req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Admin access required for this action' });
    }
    next();
  });
}

const ROLE_PERMISSIONS = {
  viewer: new Set([
    'view', 'export', 'self_password', 'lot_entry_view'
  ]),
  lot_entry: new Set([
    'self_password', 'view', 'lot_entry_view', 'lot_write', 'auction_write'
  ]),
  operator: new Set([
    'view', 'export', 'self_password', 'lot_entry_view',
    'lot_write', 'invoice_write', 'trader_write', 'buyer_write'
  ]),
  manager: new Set([
    'view', 'export', 'self_password', 'lot_entry_view',
    'lot_write', 'invoice_write', 'trader_write', 'buyer_write',
    'auction_write', 'invoice_revert', 'settings_write', 'state_toggle'
  ]),
  admin: new Set([
    'view', 'export', 'self_password', 'lot_entry_view',
    'lot_write', 'invoice_write', 'trader_write', 'buyer_write',
    'auction_write', 'invoice_revert', 'settings_write', 'state_toggle',
    'delete', 'delete_all', 'user_manage'
  ])
};

function userHas(role, capability) {
  const perms = ROLE_PERMISSIONS[role] || ROLE_PERMISSIONS.viewer;
  return perms.has(capability);
}

const { mountMobile } = require('./mobile-bridge');
mountMobile(app, {
  getDb,
  requireAuth,
  verifyPassword,
  hashPassword,
  isLegacyHash,
  ROLE_PERMISSIONS,
});

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

// ──────────────────────────────────────────────────────────────
// AUTH
// ──────────────────────────────────────────────────────────────

// Public branding (no auth) — login screen / topbar pulls company
// name + logo from settings. Returns only the safe-to-expose subset:
// trade name, short name, branch, GSTIN, and a logo URL when present.
app.get('/api/branding', (req, res) => {
  try {
    const cfg = getSettingsFlat(getDb());
    const isKL = String(cfg.business_state || '').toUpperCase().includes('KERALA');
    const branch = (isKL ? cfg.kl_branch : cfg.tn_branch) || cfg.tn_branch || cfg.kl_branch || '';
    const gstin  = (isKL ? cfg.kl_gstin  : cfg.tn_gstin)  || cfg.tn_gstin  || cfg.kl_gstin  || '';
    let presetConfig = null;
    if (cfg.preset_config) {
      try { presetConfig = JSON.parse(cfg.preset_config); }
      catch (_) { presetConfig = null; }
    }
    res.json({
      tradeName: cfg.trade_name || cfg.short_name || '',
      shortName: cfg.short_name || cfg.trade_name || '',
      branch,
      gstin,
      logoUrl: '/logo-ispl.png',
      theme: cfg.theme || '',
      themeCustomColor: cfg.theme_custom_color || '',
      preset: cfg.tenant_preset || '',
      presetConfig,
    });
  } catch (e) {
    res.json({ tradeName: '', shortName: '', branch: '', gstin: '', logoUrl: null, theme: '', themeCustomColor: '', preset: '', presetConfig: null });
  }
});

app.put('/api/branding', requireSettingsWrite, (req, res) => {
  try {
    const db = getDb();
    const body = req.body || {};
    const allowed = {};
    if (typeof body.theme === 'string') {
      const THEMES = ['emerald','coral','violet','sunshine','electric','ocean','tech','minimal','trust','rose','indigo','teal','slate','custom'];
      if (THEMES.includes(body.theme)) allowed.theme = body.theme;
    }
    if (typeof body.themeCustomColor === 'string' && /^#[0-9a-fA-F]{6}$/.test(body.themeCustomColor)) {
      allowed.theme_custom_color = body.themeCustomColor;
    }
    if (!Object.keys(allowed).length) {
      return res.status(400).json({ error: 'No valid branding fields supplied' });
    }
    const stmt = db.prepare(
      `INSERT INTO company_settings (key, value, category, label, field_type)
       VALUES (?, ?, 'branding', ?, 'text')
       ON CONFLICT(key) DO UPDATE SET value = excluded.value`
    );
    const labels = { theme: 'Theme', theme_custom_color: 'Custom primary color' };
    for (const [k, v] of Object.entries(allowed)) {
      stmt.run(k, String(v), labels[k] || k);
    }
    res.json({ success: true, updated: Object.keys(allowed).length });
  } catch (e) {
    res.status(500).json({ error: 'Branding save failed: ' + (e.message || e) });
  }
});

// Tenant preset (white-label switcher) — developer-only admin UI.
const TENANT_PRESETS = {
  cardamom: { label: 'Cardamom (default — premium spacious)', theme: 'emerald', customColor: '', density: 'roomy', font: 'jakarta', hideAppearance: true },
  bluehill: { label: 'Bluehill (indigo + dense + inter)', theme: 'indigo', customColor: '', density: 'compact', font: 'inter', hideAppearance: true },
  'western-ghats': { label: 'Western Ghats (teal + spacious + outfit)', theme: 'teal', customColor: '', density: 'spacious', font: 'outfit', hideAppearance: true },
  slate: { label: 'Slate (corporate grey + dense + system)', theme: 'slate', customColor: '', density: 'compact', font: 'system', hideAppearance: true },
  marigold: { label: 'Marigold (sunshine + roomy + jakarta)', theme: 'sunshine', customColor: '', density: 'roomy', font: 'jakarta', hideAppearance: true },
  ocean: { label: 'Ocean (cool blue + roomy + inter)', theme: 'ocean', customColor: '', density: 'roomy', font: 'inter', hideAppearance: true },
};
const TENANT_FONTS = ['jakarta', 'inter', 'outfit', 'system'];
const TENANT_DENSITIES = ['compact', 'roomy', 'spacious'];

function checkAdminKey(req) {
  const expected = process.env.ADMIN_BRANDING_KEY || 'change-me';
  const supplied = String(req.query.key || req.headers['x-admin-key'] || '');
  return supplied === expected && supplied.length > 0;
}

app.get('/admin/branding', (req, res) => {
  if (!checkAdminKey(req)) {
    return res.status(403).type('html').send(
      '<html><body style="font-family:system-ui;padding:40px;text-align:center;color:#666"><h2>403 — Access denied</h2><p>This page requires a valid <code>?key=</code> in the URL.</p></body></html>'
    );
  }
  const db = getDb();
  const cfg = getSettingsFlat(db);
  const currentPreset = cfg.tenant_preset || '';
  let currentConfig = null;
  try { currentConfig = JSON.parse(cfg.preset_config || 'null'); } catch (_) {}
  const presetOptions = Object.entries(TENANT_PRESETS).map(([slug, p]) => {
    const sel = slug === currentPreset ? 'selected' : '';
    return `<option value="${slug}" ${sel}>${slug} — ${p.label}</option>`;
  }).join('');
  const customSel = currentPreset === 'custom' ? 'selected' : '';
  const noneSel = !currentPreset ? 'selected' : '';
  const c = (currentPreset === 'custom' && currentConfig) ? currentConfig : {};
  const cTheme    = c.theme || 'emerald';
  const cColor    = c.customColor || '';
  const cDensity  = c.density || 'roomy';
  const cFont     = c.font || 'jakarta';
  const cHide     = c.hideAppearance !== false;
  const themeOpts = ['emerald','coral','violet','sunshine','electric','ocean','tech','minimal','trust','rose','indigo','teal','slate','custom']
    .map(t => `<option value="${t}" ${t === cTheme ? 'selected' : ''}>${t}</option>`).join('');
  const densityOpts = TENANT_DENSITIES.map(d => `<option value="${d}" ${d === cDensity ? 'selected' : ''}>${d}</option>`).join('');
  const fontOpts = TENANT_FONTS.map(f => `<option value="${f}" ${f === cFont ? 'selected' : ''}>${f}</option>`).join('');
  const keyEsc = String(req.query.key).replace(/[<>'"&]/g, '');
  res.type('html').send(`<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Tenant Branding — Admin</title>
<style>body{font-family:system-ui;max-width:720px;margin:40px auto;padding:20px;color:#1f2937;background:#f9fafb}h1{font-size:22px;margin:0 0 4px}h2{font-size:16px;margin:24px 0 10px;color:#374151}.card{background:#fff;border:1px solid #e5e7eb;border-radius:8px;padding:20px;margin-bottom:16px}label{display:block;font-size:12px;font-weight:600;color:#374151;margin:12px 0 4px;text-transform:uppercase;letter-spacing:.3px}select,input{width:100%;padding:8px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:14px;box-sizing:border-box}button{background:#166534;color:#fff;border:0;padding:10px 18px;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer}button.secondary{background:#6b7280}</style></head>
<body><h1>Tenant branding</h1><p>Hidden admin panel.</p>
<div class="card"><h2>Current</h2><p><strong>Active:</strong> ${currentPreset || '<em>none</em>'}</p></div>
<div class="card"><h2>Preset</h2><select id="p"><option value="" ${noneSel}>— none —</option>${presetOptions}<option value="custom" ${customSel}>custom</option></select></div>
<div class="card"><h2>Custom</h2><label>Theme</label><select id="t">${themeOpts}</select><label>Color</label><input type="color" id="cc" value="${cColor || '#166534'}"><label>Density</label><select id="d">${densityOpts}</select><label>Font</label><select id="f">${fontOpts}</select><label><input type="checkbox" id="h" ${cHide ? 'checked' : ''}> Hide Appearance</label></div>
<button onclick="apply()">Apply</button> <button class="secondary" onclick="clr()">Clear</button>
<script>const K=${JSON.stringify(keyEsc)};async function apply(){const p=document.getElementById('p').value;const b={preset:p};if(p==='custom')b.config={theme:document.getElementById('t').value,customColor:document.getElementById('cc').value,density:document.getElementById('d').value,font:document.getElementById('f').value,hideAppearance:document.getElementById('h').checked};const r=await fetch('/api/admin/preset?key='+K,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)});if(r.ok){alert('Applied');location.reload()}else{alert('Failed')}}async function clr(){if(!confirm('Clear?'))return;const r=await fetch('/api/admin/preset?key='+K,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({preset:''})});if(r.ok){location.reload()}}</script></body></html>`);
});

app.post('/api/admin/preset', (req, res) => {
  if (!checkAdminKey(req)) return res.status(403).json({ error: 'Invalid admin key' });
  const db = getDb();
  const { preset, config } = req.body || {};
  const slug = String(preset || '').trim();
  let finalConfig = null;
  if (slug === '') finalConfig = null;
  else if (slug === 'custom') {
    if (!config || typeof config !== 'object') return res.status(400).json({ error: 'custom requires config' });
    finalConfig = {
      theme: String(config.theme || 'emerald'),
      customColor: /^#[0-9a-fA-F]{6}$/.test(config.customColor || '') ? config.customColor : '',
      density: TENANT_DENSITIES.includes(config.density) ? config.density : 'roomy',
      font: TENANT_FONTS.includes(config.font) ? config.font : 'jakarta',
      hideAppearance: !!config.hideAppearance,
    };
  } else if (TENANT_PRESETS[slug]) finalConfig = TENANT_PRESETS[slug];
  else return res.status(400).json({ error: `Unknown preset: ${slug}` });
  const stmt = db.prepare(
    `INSERT INTO company_settings (key, value, category, label, field_type)
     VALUES (?, ?, 'branding', ?, 'text')
     ON CONFLICT(key) DO UPDATE SET value = excluded.value`
  );
  stmt.run('tenant_preset', slug, 'Tenant preset');
  stmt.run('preset_config', finalConfig ? JSON.stringify(finalConfig) : '', 'Tenant preset config (JSON)');
  res.json({ success: true, preset: slug, config: finalConfig });
});

const loginLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 10,
  standardHeaders: true,
  legacyHeaders: false,
  skipSuccessfulRequests: true,
  message: { error: 'Too many login attempts. Please wait a few minutes and try again.' }
});

const SESSION_TTL_DAYS = 30;

app.post('/api/login', loginLimiter, async (req, res) => {
  const { username, password, device_label } = req.body || {};
  if (!username || !password) return res.status(401).json({ error: 'Invalid credentials' });
  const db = getDb();
  const user = db.get('SELECT * FROM users WHERE username = ?', [username]);
  const ok = await verifyPassword(password, user ? user.password_hash : DUMMY_BCRYPT_HASH);
  if (!user || !ok) return res.status(401).json({ error: 'Invalid credentials' });
  if (isLegacyHash(user.password_hash)) {
    try {
      const upgraded = await hashPassword(password);
      db.run('UPDATE users SET password_hash = ? WHERE id = ?', [upgraded, user.id]);
    } catch (_) {}
  }
  const token = crypto.randomBytes(32).toString('hex');
  db.run(
    `INSERT INTO sessions (token, user_id, device_label, expires_at)
     VALUES (?, ?, ?, datetime('now','localtime','+${SESSION_TTL_DAYS} days'))`,
    [token, user.id, device_label || '']
  );
  db.run(
    `DELETE FROM sessions
     WHERE (expires_at IS NOT NULL AND expires_at < datetime('now','localtime'))
        OR last_used_at < datetime('now','-30 days')`
  );
  const permissions = Array.from(ROLE_PERMISSIONS[user.role] || ROLE_PERMISSIONS.viewer);
  res.json({
    token, role: user.role, username: user.username, permissions,
    must_change_password: !!user.must_change_password
  });
});

app.post('/api/logout', (req, res) => {
  const t = (req.headers.authorization||'').replace('Bearer ','');
  if (t) getDb().run('DELETE FROM sessions WHERE token = ?', [t]);
  res.json({ success: true });
});

app.get('/api/me', requireAnyPermission('view', 'lot_entry_view', 'self_password'), (req, res) => {
  const permissions = Array.from(ROLE_PERMISSIONS[req.user.role] || ROLE_PERMISSIONS.viewer);
  res.json({
    username: req.user.username,
    role: req.user.role,
    permissions,
    must_change_password: !!req.user.must_change_password
  });
});

// ──────────────────────────────────────────────────────────────
// USER MANAGEMENT (admin-only)
// ──────────────────────────────────────────────────────────────
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
  const VALID_ROLES = ['viewer', 'lot_entry', 'operator', 'manager', 'admin'];
  let finalRole = (role || '').toLowerCase();
  if (finalRole === 'user') finalRole = 'viewer';
  if (!VALID_ROLES.includes(finalRole)) finalRole = 'operator';
  const db = getDb();
  const existing = db.get('SELECT id FROM users WHERE username = ?', [username]);
  if (existing) return res.status(400).json({ error: 'Username already exists' });
  const pwHash = await hashPassword(password);
  db.run(
    'INSERT INTO users (username, password_hash, role, must_change_password) VALUES (?, ?, ?, 1)',
    [username, pwHash, finalRole]
  );
  const created = db.get('SELECT id, username, role FROM users WHERE username = ?', [username]);
  res.json({ success: true, id: created ? created.id : null, username, role: finalRole });
});

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
  const pwHash = await hashPassword(password);
  db.run(
    'UPDATE users SET password_hash = ?, must_change_password = 1 WHERE id = ?',
    [pwHash, user.id]
  );
  db.run('DELETE FROM sessions WHERE user_id = ?', [user.id]);
  res.json({ success: true, username: user.username });
});

app.delete('/api/users/:id', requireUserManage, (req, res) => {
  const db = getDb();
  const target = db.get('SELECT id, username, role FROM users WHERE id = ?', [req.params.id]);
  if (!target) return res.status(404).json({ error: 'User not found' });
  if (target.id === req.user.id) return res.status(400).json({ error: 'You cannot delete your own account while signed in' });
  const total = db.get('SELECT COUNT(*) as c FROM users').c;
  if (total <= 1) return res.status(400).json({ error: 'Cannot delete the last remaining user' });
  db.run('DELETE FROM sessions WHERE user_id = ?', [target.id]);
  db.run('DELETE FROM users WHERE id = ?', [target.id]);
  res.json({ success: true, username: target.username });
});

app.put('/api/me/password', requirePermission('self_password'), async (req, res) => {
  const { current_password, new_password } = req.body || {};
  if (!current_password || !new_password) return res.status(400).json({ error: 'Both current and new password required' });
  if (new_password.length < 4) return res.status(400).json({ error: 'New password must be at least 4 characters' });
  const db = getDb();
  const user = db.get('SELECT * FROM users WHERE id = ?', [req.user.id]);
  if (!user || !(await verifyPassword(current_password, user.password_hash))) {
    return res.status(401).json({ error: 'Current password is incorrect' });
  }
  if (await verifyPassword(new_password, user.password_hash)) {
    return res.status(400).json({ error: 'New password must be different from current password' });
  }
  const newHash = await hashPassword(new_password);
  db.run(
    'UPDATE users SET password_hash = ?, must_change_password = 0 WHERE id = ?',
    [newHash, user.id]
  );
  db.run('DELETE FROM sessions WHERE user_id = ? AND token != ?', [user.id, req.session.token]);
  res.json({ success: true });
});

app.get('/api/me/sessions', requireView, (req, res) => {
  const db = getDb();
  const sessions = db.all(
    `SELECT token, device_label, created_at, last_used_at,
            CASE WHEN token = ? THEN 1 ELSE 0 END as is_current
     FROM sessions WHERE user_id = ? ORDER BY last_used_at DESC`,
    [req.session.token, req.user.id]
  );
  res.json(sessions.map(s => ({ ...s, token: '…' + (s.token || '').slice(-8) })));
});

app.delete('/api/me/sessions/:tokenSuffix', requireView, (req, res) => {
  const suffix = req.params.tokenSuffix;
  const db = getDb();
  const sessions = db.all('SELECT token FROM sessions WHERE user_id = ?', [req.user.id]);
  const match = sessions.find(s => (s.token || '').endsWith(suffix));
  if (!match) return res.status(404).json({ error: 'Session not found' });
  if (match.token === req.session.token) return res.status(400).json({ error: 'Use Logout to end your current session' });
  db.run('DELETE FROM sessions WHERE token = ?', [match.token]);
  res.json({ success: true });
});

// ──────────────────────────────────────────────────────────────
// COMPANY SETTINGS
// ──────────────────────────────────────────────────────────────
app.get('/api/company-settings', requireViewOrLotEntry, (req, res) => {
  res.json({ categories: CATEGORIES, settings: getAllSettings(getDb()) });
});
app.put('/api/company-settings', requireSettingsWrite, (req, res) => {
  const count = updateSettings(getDb(), req.body.settings || {});
  invalidateDateFormatCache();
  res.json({ success: true, updated: count });
});
app.get('/api/company-settings/flat', requireViewOrLotEntry, (req, res) => res.json(getSettingsFlat(getDb())));

function _activeLabel() {
  try {
    const id = getCompanyIdentity(getSettingsFlat(getDb()));
    return id.shortName || id.logoCode || '';
  } catch (_) { return ''; }
}
app.get('/api/company-presets', requireView, (_req, res) => {
  const a = _activeLabel();
  res.json({ [a || 'default']: {}, active: a });
});
app.put('/api/company-presets/active', requireStateToggle, (_req, res) => {
  res.json({ success: true, active: _activeLabel() });
});
app.put('/api/company-presets/:code', requireSettingsWrite, (_req, res) => {
  res.json({ success: true });
});

const LOGO_FILES = {
  ispl: path.join(__dirname, 'public', 'logo-ispl.png'),
  asp:  path.join(__dirname, 'public', 'logo-asp.png'),
};
app.post('/api/company-settings/logo/:which', requireSettingsWrite, upload.single('file'), (req, res) => {
  const which = req.params.which;
  if (!LOGO_FILES[which]) return res.status(400).json({ error: 'Invalid logo type (use ispl or asp)' });
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const ext = (req.file.originalname.split('.').pop() || '').toLowerCase();
  if (!['png', 'jpg', 'jpeg'].includes(ext)) {
    fs.unlinkSync(req.file.path);
    return res.status(400).json({ error: 'Only PNG or JPG images allowed' });
  }
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

// ──────────────────────────────────────────────────────────────
// BULK DELETE ROUTES
// ──────────────────────────────────────────────────────────────
const DELETE_ALL_RESOURCES = {
  traders:      { table: 'traders',     cascade: ['trader_banks'],                                          scope: 'global' },
  buyers:       { table: 'buyers',      cascade: [],                                                        scope: 'global' },
  invoices:     { table: 'invoices',    cascade: [],                                                        scope: 'global' },
  purchases:    { table: 'purchases',   cascade: [],                                                        scope: 'global' },
  bills:        { table: 'bills',       cascade: [],                                                        scope: 'global' },
  auctions:     { table: 'auctions',    cascade: ['lots','invoices','purchases','bills','debit_notes','lot_allocations'], scope: 'global' },
  'debit-notes': { table: 'debit_notes', cascade: [],                                                       scope: 'trade' },
};

function _snapshotBackupBeforeDelete(resource) {
  const backupDir = path.join(process.env.SPICE_DATA_DIR || path.join(__dirname, 'data'), 'backups');
  try { fs.mkdirSync(backupDir, { recursive: true }); } catch (_) {}
  const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const filename = `before-delete-${resource}-${stamp}.db`;
  const target = path.join(backupDir, filename);
  try { require('./db').flushSave(); } catch (_) {}
  fs.copyFileSync(DB_PATH, target);
  return target;
}

function _logDelete(db, { resource, deletedCount, cascadeCounts, backupPath, req }) {
  try {
    db.run(
      `INSERT INTO delete_log (resource, deleted_count, cascade_counts, backup_path, user_id, username, ip)
       VALUES (?,?,?,?,?,?,?)`,
      [
        resource,
        deletedCount,
        JSON.stringify(cascadeCounts || {}),
        backupPath || '',
        (req.user && req.user.id) || null,
        (req.user && req.user.username) || '',
        String(req.ip || req.headers['x-forwarded-for'] || '').slice(0, 64),
      ],
    );
  } catch (_) {}
}

function _countDeleteAllImpact(db, resource, ano) {
  const def = DELETE_ALL_RESOURCES[resource];
  if (!def) return null;
  const counts = {};
  if (def.scope === 'trade') {
    counts[def.table] = db.get(`SELECT COUNT(*) as c FROM ${def.table} WHERE ano = ?`, [ano || '']).c;
  } else {
    counts[def.table] = db.get(`SELECT COUNT(*) as c FROM ${def.table}`).c;
    for (const t of def.cascade) {
      try { counts[t] = db.get(`SELECT COUNT(*) as c FROM ${t}`).c; }
      catch (_) { counts[t] = 0; }
    }
  }
  return counts;
}

app.get('/api/admin/delete-all/preflight', requireDeleteAll, (req, res) => {
  const resource = String(req.query.resource || '').trim();
  const def = DELETE_ALL_RESOURCES[resource];
  if (!def) return res.status(400).json({ error: 'Unknown resource', available: Object.keys(DELETE_ALL_RESOURCES) });
  if (def.scope === 'trade' && !String(req.query.ano || '').trim()) {
    return res.status(400).json({ error: 'ano query param required for trade-scoped delete' });
  }
  const counts = _countDeleteAllImpact(getDb(), resource, req.query.ano);
  res.json({ resource, scope: def.scope, counts });
});

app.get('/api/admin/delete-log', requireDeleteAll, (req, res) => {
  const rows = getDb().all(
    `SELECT id, resource, deleted_count, cascade_counts, backup_path, username, ip, created_at
       FROM delete_log ORDER BY id DESC LIMIT 200`
  );
  res.json(rows.map(r => ({
    ...r,
    cascade_counts: (() => { try { return JSON.parse(r.cascade_counts || '{}'); } catch (_) { return {}; } })(),
  })));
});

function makeDeleteAll(resource) {
  const def = DELETE_ALL_RESOURCES[resource];
  return (req, res) => {
    try {
      const db = getDb();
      let backupPath = '';
      try { backupPath = _snapshotBackupBeforeDelete(resource); }
      catch (e) {
        return res.status(500).json({ error: 'Backup snapshot failed; refusing to delete: ' + (e.message || e) });
      }
      const counts = _countDeleteAllImpact(db, resource);
      const before = counts[def.table] || 0;
      for (const t of def.cascade) {
        try { db.run(`DELETE FROM ${t}`); } catch (_) {}
        try { db.exec(`DELETE FROM sqlite_sequence WHERE name = '${t}'`); } catch (_) {}
      }
      db.run(`DELETE FROM ${def.table}`);
      try { db.exec(`DELETE FROM sqlite_sequence WHERE name = '${def.table}'`); } catch (_) {}
      _logDelete(db, { resource, deletedCount: before, cascadeCounts: counts, backupPath, req });
      res.json({ success: true, deleted: before, cascadeCounts: counts, backupPath });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  };
}
app.delete('/api/traders/delete-all',     requireDeleteAll, makeDeleteAll('traders'));
app.delete('/api/buyers/delete-all',      requireDeleteAll, makeDeleteAll('buyers'));
app.delete('/api/invoices/delete-all',    requireDeleteAll, makeDeleteAll('invoices'));
app.delete('/api/purchases/delete-all',   requireDeleteAll, makeDeleteAll('purchases'));
app.delete('/api/bills/delete-all',       requireDeleteAll, makeDeleteAll('bills'));
app.delete('/api/auctions/delete-all',    requireDeleteAll, makeDeleteAll('auctions'));
app.delete('/api/debit-notes/delete-all', requireDeleteAll, (req, res) => {
  try {
    const db = getDb();
    const ano = String(req.query.ano || '').trim();
    const scope = ano ? ('debit-notes-' + ano) : 'debit-notes';
    let backupPath = '';
    try { backupPath = _snapshotBackupBeforeDelete(scope); }
    catch (e) {
      return res.status(500).json({ error: 'Backup snapshot failed; refusing to delete: ' + (e.message || e) });
    }
    let before;
    if (ano) {
      before = db.get('SELECT COUNT(*) as c FROM debit_notes WHERE ano = ?', [ano]).c;
      db.run('DELETE FROM debit_notes WHERE ano = ?', [ano]);
    } else {
      before = db.get('SELECT COUNT(*) as c FROM debit_notes').c;
      db.run('DELETE FROM debit_notes');
      try { db.exec(`DELETE FROM sqlite_sequence WHERE name = 'debit_notes'`); } catch (_) {}
    }
    _logDelete(db, {
      resource: 'debit-notes',
      deletedCount: before,
      cascadeCounts: ano ? { ano, debit_notes: before } : { debit_notes: before },
      backupPath,
      req,
    });
    res.json({ success: true, deleted: before, ano: ano || null, backupPath });
  } catch (e) {
    res.status(500).json({ error: 'Delete All failed: ' + (e.message || e) });
  }
});

// ──────────────────────────────────────────────────────────────
// GST LOOKUP
// ──────────────────────────────────────────────────────────────
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
  if (!apiKey) {
    return res.json({
      valid: true, gstin, pan, st_code: stCode, state,
      source: 'structural',
      note: 'Set "gst_api_key" in settings to auto-fetch trade name/address.'
    });
  }
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

// ──────────────────────────────────────────────────────────────
// TRADERS
// ──────────────────────────────────────────────────────────────
app.get('/api/traders', requireViewOrLotEntry, (req, res) => {
  const { search, limit } = req.query;
  const db = getDb();
  const hydrateBanks = (rows) => {
    if (!rows.length) return rows;
    const ids = rows.map(r => r.id);
    const placeholders = ids.map(() => '?').join(',');
    const banks = db.all(
      `SELECT id, trader_id, bank_name, acctnum, ifsc, holder_name, is_default
       FROM trader_banks WHERE trader_id IN (${placeholders})
       ORDER BY trader_id, is_default DESC, id`, ids
    );
    const byTrader = new Map();
    for (const b of banks) {
      if (!byTrader.has(b.trader_id)) byTrader.set(b.trader_id, []);
      byTrader.get(b.trader_id).push(b);
    }
    for (const r of rows) r.banks = byTrader.get(r.id) || [];
    return rows;
  };
  const page     = Math.max(1, parseInt(req.query.page, 10) || 1);
  const pageSize = Math.min(500, Math.max(1, parseInt(req.query.pageSize || req.query.limit, 10) || 50));
  const offset   = (page - 1) * pageSize;
  const wantPaged = req.query.page != null || req.query.pageSize != null;
  let where = '';
  let params = [];
  if (search) {
    const q = `%${search}%`;
    where = 'WHERE name LIKE ? OR tel LIKE ? OR cr LIKE ? OR pan LIKE ? OR ppla LIKE ? OR aadhar LIKE ?';
    params = [q, q, q, q, q, q];
  }
  if (wantPaged) {
    const total = db.get(`SELECT COUNT(*) as c FROM traders ${where}`, params).c;
    const rows = db.all(
      `SELECT * FROM traders ${where} ORDER BY name LIMIT ? OFFSET ?`,
      [...params, pageSize, offset]
    );
    return res.json({ rows: hydrateBanks(rows), total, page, pageSize });
  }
  if (search) {
    const rows = db.all(
      `SELECT * FROM traders ${where} ORDER BY name LIMIT ?`,
      [...params, pageSize]
    );
    const normalize = (s) => String(s || '').trim().toUpperCase();
    const stripGstinPrefix = (s) => {
      let v = normalize(s);
      if (v.startsWith('GSTIN.')) v = v.slice(6);
      else if (v.startsWith('GSTIN')) v = v.slice(5);
      return v;
    };
    const seen = new Map();
    for (const r of rows) {
      const key = normalize(r.name) + '|' + stripGstinPrefix(r.cr);
      const prev = seen.get(key);
      if (!prev || (r.id || 0) > (prev.id || 0)) seen.set(key, r);
    }
    return res.json(hydrateBanks([...seen.values()]));
  }
  res.json(hydrateBanks(db.all('SELECT * FROM traders ORDER BY name LIMIT 500')));
});

app.get('/api/traders/by-name/:name', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const nm = String(req.params.name || '').trim();
  if (!nm) return res.status(400).json({ error: 'name required' });
  const row = db.get('SELECT id, name, tel FROM traders WHERE LOWER(name) = LOWER(?) LIMIT 1', [nm]);
  if (!row) return res.status(404).json({ error: 'Not found' });
  res.json(row);
});

// ── WhatsApp Cloud API (Meta) ──
function _waConfigured() {
  return Boolean(process.env.WHATSAPP_TOKEN && process.env.WHATSAPP_PHONE_ID);
}
async function _waPost(path, body) {
  const url = `https://graph.facebook.com/v18.0/${process.env.WHATSAPP_PHONE_ID}${path}`;
  const r = await fetch(url, {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + process.env.WHATSAPP_TOKEN, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(j.error && j.error.message ? j.error.message : `Cloud API error ${r.status}`);
  return j;
}
function _waNormalizePhone(tel) {
  const d = String(tel || '').replace(/\D/g, '');
  if (!d) return '';
  return d.length === 10 ? '91' + d : d;
}

app.post('/api/whatsapp/send-text', requireView, async (req, res) => {
  if (!_waConfigured()) return res.status(501).json({ error: 'WhatsApp Cloud API not configured', fallback: true });
  try {
    const phone = _waNormalizePhone(req.body.phone);
    const message = String(req.body.message || '');
    if (!phone || !message) return res.status(400).json({ error: 'phone and message required' });
    const out = await _waPost('/messages', {
      messaging_product: 'whatsapp', to: phone, type: 'text', text: { body: message },
    });
    res.json({ ok: true, id: out.messages && out.messages[0] && out.messages[0].id });
  } catch (e) { res.status(502).json({ error: e.message }); }
});

app.post('/api/whatsapp/send-document', requireView, async (req, res) => {
  if (!_waConfigured()) return res.status(501).json({ error: 'WhatsApp Cloud API not configured', fallback: true });
  try {
    const phone = _waNormalizePhone(req.body.phone);
    const caption = String(req.body.caption || '');
    const docUrl = String(req.body.doc_url || '');
    if (!phone || !docUrl) return res.status(400).json({ error: 'phone and doc_url required' });
    const out = await _waPost('/messages', {
      messaging_product: 'whatsapp', to: phone, type: 'document',
      document: { link: docUrl, caption, filename: req.body.filename || 'document.pdf' },
    });
    res.json({ ok: true, id: out.messages && out.messages[0] && out.messages[0].id });
  } catch (e) { res.status(502).json({ error: e.message }); }
});

app.post('/api/whatsapp/send-pdf', requireView, upload.single('file'), async (req, res) => {
  if (!_waConfigured()) {
    if (req.file) fs.unlink(req.file.path, () => {});
    return res.status(501).json({ error: 'WhatsApp Cloud API not configured', fallback: true });
  }
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const phone = _waNormalizePhone(req.body.phone);
    const caption = String(req.body.caption || '');
    const filename = String(req.body.filename || req.file.originalname || 'document.pdf');
    if (!phone) return res.status(400).json({ error: 'phone required' });
    const fileBuffer = fs.readFileSync(req.file.path);
    const fd = new FormData();
    fd.append('messaging_product', 'whatsapp');
    fd.append('type', 'application/pdf');
    fd.append('file', new Blob([fileBuffer], { type: 'application/pdf' }), filename);
    const mediaUrl = `https://graph.facebook.com/v18.0/${process.env.WHATSAPP_PHONE_ID}/media`;
    const mediaRes = await fetch(mediaUrl, {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + process.env.WHATSAPP_TOKEN },
      body: fd,
    });
    const mediaJson = await mediaRes.json().catch(() => ({}));
    if (!mediaRes.ok || !mediaJson.id) {
      const msg = (mediaJson.error && mediaJson.error.message) || `Media upload failed ${mediaRes.status}`;
      throw new Error(msg);
    }
    const out = await _waPost('/messages', {
      messaging_product: 'whatsapp', to: phone, type: 'document',
      document: { id: mediaJson.id, caption, filename },
    });
    res.json({
      ok: true,
      id: out.messages && out.messages[0] && out.messages[0].id,
      mediaId: mediaJson.id,
    });
  } catch (e) {
    res.status(502).json({ error: e.message });
  } finally {
    if (req.file) fs.unlink(req.file.path, () => {});
  }
});

app.get('/api/traders/:id', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const row = db.get('SELECT * FROM traders WHERE id = ?', [req.params.id]);
  if (!row) return res.status(404).json({ error: 'Not found' });
  row.banks = db.all(
    'SELECT id, trader_id, bank_name, acctnum, ifsc, holder_name, is_default FROM trader_banks WHERE trader_id = ? ORDER BY is_default DESC, id',
    [row.id]
  );
  res.json(row);
});

function syncTraderBanks(db, traderId, banks) {
  const arr = Array.isArray(banks) ? banks.filter(b => b && (b.acctnum || b.ifsc)) : [];
  db.run('DELETE FROM trader_banks WHERE trader_id = ?', [traderId]);
  for (const b of arr) {
    db.run(
      'INSERT INTO trader_banks (trader_id, bank_name, acctnum, ifsc, holder_name) VALUES (?,?,?,?,?)',
      [traderId, b.bank_name||'', String(b.acctnum||''), String(b.ifsc||''), b.holder_name||'']
    );
  }
  const first = arr[0] || {};
  db.run(
    'UPDATE traders SET ifsc=?, acctnum=?, holder_name=? WHERE id=?',
    [first.ifsc||'', first.acctnum||'', first.holder_name||'', traderId]
  );
}

// Duplicate-PAN guard for sellers. Case-insensitive trim match.
function _findTraderDuplicateByPan(db, pan, excludeId) {
  const norm = String(pan || '').trim().toUpperCase();
  if (!norm) return null;
  const params = [norm];
  let sql = 'SELECT id, name, cr, pan, tel FROM traders WHERE UPPER(TRIM(pan)) = ?';
  if (excludeId) { sql += ' AND id != ?'; params.push(excludeId); }
  sql += ' LIMIT 1';
  return db.get(sql, params);
}

app.post('/api/traders', requireTraderWrite, (req, res) => {
  const t = req.body;
  const db = getDb();
  const dup = _findTraderDuplicateByPan(db, t.pan);
  if (dup) return res.status(409).json({
    duplicate: true, field: 'pan', existing: dup,
    error: `A seller with PAN "${dup.pan}" already exists: ${dup.name || '(unnamed)'}`,
  });
  const info = db.run(`INSERT INTO traders (name,cr,pan,tel,aadhar,padd,ppla,pin,pstate,pst_code,ifsc,acctnum,holder_name)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [t.name,t.cr||'',t.pan||'',t.tel||'',t.aadhar||'',t.padd||'',t.ppla||'',t.pin||'',t.pstate||'',t.pst_code||'',t.ifsc||'',t.acctnum||'',t.holder_name||'']);
  if (Array.isArray(t.banks)) {
    syncTraderBanks(db, info.lastInsertRowid, t.banks);
  }
  res.json({ success: true, id: info.lastInsertRowid });
});

app.put('/api/traders/:id', requireTraderWrite, (req, res) => {
  const t = req.body;
  const db = getDb();
  const tid = parseInt(req.params.id, 10);
  const dup = _findTraderDuplicateByPan(db, t.pan, tid);
  if (dup) return res.status(409).json({
    duplicate: true, field: 'pan', existing: dup,
    error: `Another seller with PAN "${dup.pan}" already exists: ${dup.name || '(unnamed)'}`,
  });
  db.run(`UPDATE traders SET name=?,cr=?,pan=?,tel=?,aadhar=?,padd=?,ppla=?,pin=?,pstate=?,pst_code=?,ifsc=?,acctnum=?,holder_name=? WHERE id=?`,
    [t.name,t.cr||'',t.pan||'',t.tel||'',t.aadhar||'',t.padd||'',t.ppla||'',t.pin||'',t.pstate||'',t.pst_code||'',t.ifsc||'',t.acctnum||'',t.holder_name||'',tid]);
  if (Array.isArray(t.banks)) {
    syncTraderBanks(db, tid, t.banks);
  }
  res.json({ success: true });
});

app.delete('/api/traders/:id', requireDelete, (req, res) => {
  const db = getDb();
  db.run('DELETE FROM trader_banks WHERE trader_id = ?', [req.params.id]);
  db.run('DELETE FROM traders WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

// Lot-entry quick-add seller (Feature #6 + #1)
app.post('/api/traders/quick', requireAnyPermission('trader_write', 'lot_write'), (req, res) => {
  const t = req.body || {};
  if (!t.name || !String(t.name).trim()) {
    return res.status(400).json({ error: 'Name is required' });
  }
  const db = getDb();
  const panDup = _findTraderDuplicateByPan(db, t.pan);
  if (panDup) return res.status(409).json({
    duplicate: true, field: 'pan', existing: panDup,
    error: `A seller with PAN "${panDup.pan}" already exists: ${panDup.name || '(unnamed)'}`,
  });
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

// Set a bank as the trader's default
app.put('/api/traders/:id/bank-default/:bankId', requireAnyPermission('trader_write', 'lot_write'), (req, res) => {
  const traderId = parseInt(req.params.id, 10);
  const bankId   = parseInt(req.params.bankId, 10);
  if (!Number.isFinite(traderId) || !Number.isFinite(bankId)) {
    return res.status(400).json({ error: 'Invalid trader or bank id' });
  }
  const db = getDb();
  const bank = db.get('SELECT * FROM trader_banks WHERE id = ? AND trader_id = ?', [bankId, traderId]);
  if (!bank) return res.status(404).json({ error: 'Bank not found for this trader' });
  db.run('UPDATE trader_banks SET is_default = 0 WHERE trader_id = ?', [traderId]);
  db.run('UPDATE trader_banks SET is_default = 1 WHERE id = ?', [bankId]);
  db.run('UPDATE traders SET acctnum = ?, ifsc = ?, holder_name = ? WHERE id = ?',
    [bank.acctnum || '', bank.ifsc || '', bank.holder_name || '', traderId]);
  res.json({ success: true });
});

// Import Sellers from XLS/XLSX
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
      db.run('DELETE FROM trader_banks');
      db.run('DELETE FROM traders');
    }
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

app.get('/api/traders/template', requireExport, async (req, res) => {
  const db = getDb();
  const cols = [
    { header: 'NAME',         key: 'name',         width: 30 },
    { header: 'CR',           key: 'cr',           width: 28 },
    { header: 'PAN',          key: 'pan',          width: 14 },
    { header: 'TEL',          key: 'tel',          width: 16 },
    { header: 'AADHAR',       key: 'aadhar',       width: 16 },
    { header: 'PADD',         key: 'padd',         width: 50 },
    { header: 'PPLA',         key: 'ppla',         width: 20 },
    { header: 'PIN',          key: 'pin',          width: 10 },
    { header: 'PSTATE',       key: 'pstate',       width: 14 },
    { header: 'PST_CODE',     key: 'pst_code',     width: 10, align: 'left' },
    { header: 'IFSC',         key: 'ifsc',         width: 18 },
    { header: 'ACCTNUM',      key: 'acctnum',      width: 24, align: 'left' },
    { header: 'HOLDER_NAME',  key: 'holder_name',  width: 22 },
  ];
  const bizState = (getSetting(db, 'business_state') || 'TAMIL NADU').toUpperCase();
  const stCode = bizState === 'KERALA' ? '32' : '33';
  const sample = [{
    name: 'SAMPLE SELLER', cr: 'CR.12345', pan: 'ABCDE1234F', tel: '9876543210',
    aadhar: '', padd: '123 MAIN STREET', ppla: '', pin: '',
    pstate: bizState, pst_code: stCode, ifsc: '', acctnum: '', holder_name: 'SAMPLE SELLER',
  }];
  const buf = await createExcelBuffer('Sellers', cols, sample, {
    db, title: 'SELLERS TEMPLATE',
    metaLines: [`Date: ${fmtDate(todayLocalISO())}`],
  });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="sellers-template.xlsx"');
  res.send(Buffer.from(buf));
});

// ──────────────────────────────────────────────────────────────
// BUYERS
// ──────────────────────────────────────────────────────────────
app.get('/api/buyers', requireView, (req, res) => {
  const { search } = req.query;
  const db = getDb();
  const wantAll   = String(req.query.all || '') === '1';
  const wantPaged = req.query.page != null || req.query.pageSize != null;
  let where = '';
  let params = [];
  if (search) {
    const q = `%${search}%`;
    where = `WHERE buyer LIKE ? OR buyer1 LIKE ? OR tel LIKE ? OR gstin LIKE ? OR pan LIKE ? OR pla LIKE ? OR ti LIKE ? OR code LIKE ?`;
    params = [q, q, q, q, q, q, q, q];
  }
  if (wantAll) {
    return res.json(db.all(`SELECT * FROM buyers ${where} ORDER BY buyer1`, params));
  }
  if (wantPaged) {
    const page     = Math.max(1, parseInt(req.query.page, 10) || 1);
    const pageSize = Math.min(500, Math.max(1, parseInt(req.query.pageSize, 10) || 50));
    const offset   = (page - 1) * pageSize;
    const total = db.get(`SELECT COUNT(*) as c FROM buyers ${where}`, params).c;
    const rows = db.all(
      `SELECT * FROM buyers ${where} ORDER BY buyer1 LIMIT ? OFFSET ?`,
      [...params, pageSize, offset]
    );
    return res.json({ rows, total, page, pageSize });
  }
  if (search) {
    return res.json(db.all(
      `SELECT * FROM buyers ${where} ORDER BY buyer1 LIMIT 50`,
      params
    ));
  }
  res.json(db.all('SELECT * FROM buyers ORDER BY buyer1 LIMIT 500'));
});

// Duplicate-key guard for buyers (Feature #6).
function _findBuyerDuplicate(db, buyer, code, excludeId) {
  const nb = String(buyer || '').trim().toUpperCase();
  const nc = String(code  || '').trim().toUpperCase();
  const buildSql = (col, val) => {
    const params = [val];
    let sql = `SELECT id, buyer, buyer1, code FROM buyers WHERE UPPER(TRIM(${col})) = ?`;
    if (excludeId) { sql += ' AND id != ?'; params.push(excludeId); }
    sql += ' LIMIT 1';
    return { sql, params };
  };
  if (nb) {
    const q = buildSql('buyer', nb);
    const hit = db.get(q.sql, q.params);
    if (hit) return { row: hit, field: 'buyer' };
  }
  if (nc) {
    const q = buildSql('code', nc);
    const hit = db.get(q.sql, q.params);
    if (hit) return { row: hit, field: 'code' };
  }
  return null;
}

app.post('/api/buyers', requireBuyerWrite, (req, res) => {
  const b = req.body;
  const db = getDb();
  const dup = _findBuyerDuplicate(db, b.buyer, b.code);
  if (dup) return res.status(409).json({
    duplicate: true, field: dup.field, existing: dup.row,
    error: `A buyer with ${dup.field === 'buyer' ? `code "${dup.row.buyer}"` : `short alias "${dup.row.code}"`} already exists${dup.row.buyer1 ? `: ${dup.row.buyer1}` : ''}`,
  });
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
  const bid = parseInt(req.params.id, 10);
  const dup = _findBuyerDuplicate(db, b.buyer, b.code, bid);
  if (dup) return res.status(409).json({
    duplicate: true, field: dup.field, existing: dup.row,
    error: `Another buyer with ${dup.field === 'buyer' ? `code "${dup.row.buyer}"` : `short alias "${dup.row.code}"`} already exists${dup.row.buyer1 ? `: ${dup.row.buyer1}` : ''}`,
  });
  db.run(`UPDATE buyers SET
      buyer=?, buyer1=?, code=?, sbl=?, add1=?, add2=?, pla=?, pin=?, state=?, st_code=?,
      gstin=?, pan=?, tel=?, ti=?, sale=?, email=?, tdsq=?,
      cbuyer1=?, cadd1=?, cadd2=?, cpla=?, cpin=?, cstate=?, cst_code=?, cgstin=?
    WHERE id=?`,
    [b.buyer, b.buyer1||'', b.code||'', b.sbl||'', b.add1||'', b.add2||'', b.pla||'', b.pin||'', b.state||'', b.st_code||'',
     b.gstin||'', b.pan||'', b.tel||'', b.ti||'', b.sale||'L', b.email||'', b.tdsq||'',
     b.cbuyer1||'', b.cadd1||'', b.cadd2||'', b.cpla||'', b.cpin||'', b.cstate||'', b.cst_code||'', b.cgstin||'',
     bid]);
  res.json({ success: true });
});

app.delete('/api/buyers/:id', requireDelete, (req, res) => {
  getDb().run('DELETE FROM buyers WHERE id = ?', [req.params.id]); res.json({ success: true });
});

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
      const buyer = mapCol(row, 'BUYER', 'BUYER_CODE', 'BUYERCODE');
      const code  = mapCol(row, 'CODE', 'SHORT_CODE', 'ALIAS');
      if (!buyer && !code) { skipped++; continue; }
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

app.get('/api/buyers/template', requireExport, async (req, res) => {
  const db = getDb();
  const cols = [
    { header: 'BUYER',    key: 'buyer',    width: 14 },
    { header: 'BUYER1',   key: 'buyer1',   width: 30 },
    { header: 'ADD1',     key: 'add1',     width: 30 },
    { header: 'ADD2',     key: 'add2',     width: 30 },
    { header: 'PLA',      key: 'pla',      width: 18 },
    { header: 'PIN',      key: 'pin',      width: 10, align: 'left' },
    { header: 'STATE',    key: 'state',    width: 16 },
    { header: 'ST_CODE',  key: 'st_code',  width: 10, align: 'left' },
    { header: 'GSTIN',    key: 'gstin',    width: 18 },
    { header: 'PAN',      key: 'pan',      width: 14 },
    { header: 'TEL',      key: 'tel',      width: 14 },
    { header: 'TI',       key: 'ti',       width: 10 },
    { header: 'SALE',     key: 'sale',     width: 8  },
  ];
  const bizState = (getSetting(db, 'business_state') || 'TAMIL NADU').toUpperCase();
  const stCode = bizState === 'KERALA' ? '32' : '33';
  const sample = [{
    buyer: 'ABC', buyer1: 'ABC TRADERS', add1: '10 MARKET ROAD', add2: '', pla: '',
    pin: '', state: bizState, st_code: stCode, gstin: '', pan: '', tel: '', ti: '', sale: 'L',
  }];
  const buf = await createExcelBuffer('Buyers', cols, sample, {
    db, title: 'BUYERS TEMPLATE',
    metaLines: [`Date: ${fmtDate(todayLocalISO())}`],
  });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="buyers-template.xlsx"');
  res.send(Buffer.from(buf));
});

// Multi-code buyer picker for Lot edit modal
app.get('/api/buyers/by-tradename', requireView, (req, res) => {
  const name = String(req.query.name || '').trim();
  if (!name) return res.json([]);
  const rows = getDb().all(
    `SELECT id, buyer, buyer1, code, ti, sale, gstin, pla, tel
       FROM buyers
      WHERE UPPER(TRIM(buyer1)) = UPPER(TRIM(?))
         OR UPPER(TRIM(buyer))  = UPPER(TRIM(?))
      ORDER BY code, buyer`,
    [name, name]
  );
  res.json(rows);
});

// ──────────────────────────────────────────────────────────────
// PRICE LIST (BEFORE) — code mapping tool
// ──────────────────────────────────────────────────────────────
function _plLocateColumns(ws) {
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
  const buyers = db.all('SELECT id, buyer, buyer1, code, ti, sale, gstin FROM buyers');
  const idx = new Map();
  const push = (key, row) => {
    if (!key) return;
    const k = key.trim().toUpperCase();
    if (!k) return;
    if (!idx.has(k)) idx.set(k, []);
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
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error('Workbook has no worksheets.');
  const cols = _plLocateColumns(ws);
  if (!cols.headerRow || !cols.tradeCol || !cols.codeCol) {
    throw new Error('Could not locate TRADE NAME and CODE columns. The sheet must have both headers.');
  }
  const idx = _plBuildTradeIndex(getDb());
  const perRow = [];
  let matched = 0, unmatched = 0, ambiguous = 0, blank = 0;
  const maxRow = ws.rowCount || 0;
  for (let r = cols.headerRow + 1; r <= maxRow; r++) {
    const row = ws.getRow(r);
    const tradeRaw = row.getCell(cols.tradeCol).value;
    const tradeName = String(tradeRaw == null ? '' : tradeRaw).trim();
    const entry = {
      row: r,
      tradeName,
      currentCode: String(row.getCell(cols.codeCol).value || '').trim(),
      ano:   cols.anoCol  ? String(row.getCell(cols.anoCol).value  || '').trim() : '',
      date:  cols.dateCol ? String(row.getCell(cols.dateCol).value || '').trim() : '',
      lot:   cols.lotCol  ? String(row.getCell(cols.lotCol).value  || '').trim() : '',
      status: 'blank',
      pickedCode: '',
      candidates: [],
    };
    if (!tradeName) { blank++; perRow.push(entry); continue; }
    const key = tradeName.toUpperCase();
    const cands = idx.get(key) || [];
    entry.candidates = cands.map(b => ({
      id: b.id, code: b.code, buyer: b.buyer, buyer1: b.buyer1, sale: b.sale, gstin: b.gstin,
    }));
    if (cands.length === 0) { entry.status = 'unmatched'; unmatched++; }
    else if (cands.length === 1) {
      entry.status = 'matched'; entry.pickedCode = cands[0].code || ''; matched++;
    } else {
      entry.status = 'ambiguous';
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
  } catch (e) { res.status(400).json({ error: e.message }); }
  finally { if (req.file) fs.unlink(req.file.path, () => {}); }
});
app.post('/api/price-list/map-download', requireView, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const { wb, ws, cols, perRow } = await _plProcessFile(req.file.path);
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
  } catch (e) { res.status(400).json({ error: e.message }); }
  finally { if (req.file) fs.unlink(req.file.path, () => {}); }
});

// ──────────────────────────────────────────────────────────────
// PRICE CHECK (Feature #10)
// ──────────────────────────────────────────────────────────────
const priceCheck = require('./price-check');

function pcFlagOn(db) {
  try {
    const cfg = getSettingsFlat(db || getDb());
    return String(cfg.flag_price_check || '').toLowerCase() === 'true';
  } catch (_) { return false; }
}
function pcStampGate(db, auctionId) {
  if (!auctionId) return;
  if (!pcFlagOn(db)) return;
  db.run(
    `UPDATE auctions
        SET price_checked_at = datetime('now','localtime'),
            price_check_first_passed_at = COALESCE(NULLIF(price_check_first_passed_at, ''), datetime('now','localtime'))
      WHERE id = ?`,
    [auctionId]
  );
}
function pcClearGate(db, auctionId) {
  if (!auctionId) return;
  if (!pcFlagOn(db)) return;
  db.run(`UPDATE auctions SET price_checked_at = '' WHERE id = ?`, [auctionId]);
}
function pcGateState(db, auctionId) {
  if (!auctionId)   return 'never';
  if (!pcFlagOn(db)) return 'off';
  const row = db.get(
    'SELECT price_checked_at, price_check_first_passed_at FROM auctions WHERE id = ?',
    [auctionId]
  );
  if (!row) return 'never';
  if (!row.price_check_first_passed_at) return 'never';
  return row.price_checked_at ? 'clean' : 'stale';
}
function pcGateReady(db, auctionId) {
  const s = pcGateState(db, auctionId);
  return s === 'clean' || s === 'off';
}
function requirePriceChecked(getAuctionId) {
  return (req, res, next) => {
    if (!pcFlagOn()) return next();
    const aid = getAuctionId(req);
    if (!aid) return next();
    if (pcGateState(getDb(), aid) !== 'never') return next();
    return res.status(412).json({
      error: 'Price check required',
      detail: 'Run Reports → Price Check against the auction (and apply any code fixes) before this action.',
      auctionId: aid,
      gate: 'price_check',
    });
  };
}

app.post('/api/price-check/verify', requireView, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const auctionId = req.body.auction_id ? Number(req.body.auction_id) : null;
    const { perRow, summary } = await priceCheck.processFile(
      req.file.path, getDb(),
      { auctionId }
    );
    if (auctionId && summary.gateReady) {
      pcStampGate(getDb(), auctionId);
    }
    res.json({ ...summary, rows: perRow });
  } catch (e) { res.status(400).json({ error: e.message }); }
  finally { if (req.file) fs.unlink(req.file.path, () => {}); }
});

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
    state,
    checked: state === 'clean' || state === 'off',
    everPassed: !!row.price_check_first_passed_at,
    checkedAt: row.price_checked_at || null,
    firstPassedAt: row.price_check_first_passed_at || null,
  });
});

app.post('/api/price-check/download', requireView, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const { wb, ws, cols, perRow } = await priceCheck.processFile(
      req.file.path, getDb(),
      { auctionId: req.body.auction_id || null }
    );
    priceCheck.annotateWorkbook(wb, ws, cols, perRow);
    const buf = await wb.xlsx.writeBuffer();
    const aid = req.body.auction_id;
    let baseName;
    if (aid) {
      const auc = getDb().get('SELECT ano FROM auctions WHERE id = ?', [aid]);
      baseName = `Price${auc && auc.ano ? auc.ano : aid}-checked`;
    } else {
      baseName = (req.file.originalname || 'price-check.xlsx')
        .replace(/\.xlsx?$/i, '')
        .replace(/[^A-Za-z0-9._-]+/g, '-') + '-checked';
    }
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${baseName}.xlsx"`);
    res.send(Buffer.from(buf));
  } catch (e) { res.status(400).json({ error: e.message }); }
  finally { if (req.file) fs.unlink(req.file.path, () => {}); }
});

// ──────────────────────────────────────────────────────────────
// AUCTIONS
// ──────────────────────────────────────────────────────────────
app.get('/api/auctions', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const search   = String(req.query.search || '').trim();
  const wantPaged = req.query.page != null || req.query.pageSize != null;
  let where = '';
  let params = [];
  if (search) {
    where = 'WHERE ano LIKE ?';
    params = [`%${search}%`];
  }
  const sel = `SELECT *, (SELECT COUNT(*) FROM lots WHERE auction_id=auctions.id) as lot_count
               FROM auctions ${where}
               ORDER BY date DESC, ano DESC`;
  if (wantPaged) {
    const page     = Math.max(1, parseInt(req.query.page, 10) || 1);
    const pageSize = Math.min(500, Math.max(1, parseInt(req.query.pageSize, 10) || 50));
    const offset   = (page - 1) * pageSize;
    const total = db.get(`SELECT COUNT(*) as c FROM auctions ${where}`, params).c;
    const rows = db.all(sel + ' LIMIT ? OFFSET ?', [...params, pageSize, offset]);
    return res.json({ rows: withFmtDate(rows), total, page, pageSize });
  }
  const rows = db.all(sel + ' LIMIT 100', params);
  res.json(withFmtDate(rows));
});

app.post('/api/auctions', requireAuctionWrite, (req, res) => {
  const { ano, date, crop_type, state } = req.body;
  const db = getDb();
  const d = normalizeDate(date);
  const defaultCrop  = getSetting(db, 'default_crop_type') || 'VST';
  const defaultState = getSetting(db, 'business_state')    || 'TAMIL NADU';
  db.run('INSERT INTO auctions (ano,date,crop_type,state) VALUES (?,?,?,?)', [ano, d, crop_type||defaultCrop, state||defaultState]);
  const created = db.get('SELECT id FROM auctions WHERE ano = ? AND date = ? ORDER BY id DESC LIMIT 1', [ano, d]);
  res.json({ success: true, id: created ? created.id : null });
});

app.put('/api/auctions/:id', requireAuctionWrite, (req, res) => {
  const { ano, date, crop_type, state } = req.body;
  const db = getDb();
  const defaultCrop  = getSetting(db, 'default_crop_type') || 'VST';
  const defaultState = getSetting(db, 'business_state')    || 'TAMIL NADU';
  db.run('UPDATE auctions SET ano=?, date=?, crop_type=?, state=? WHERE id=?',
    [ano, normalizeDate(date), crop_type||defaultCrop, state||defaultState, req.params.id]);
  res.json({ success: true });
});

app.delete('/api/auctions/:id', requireDelete, (req, res) => {
  const db = getDb();
  const auction = db.get('SELECT id, ano FROM auctions WHERE id = ?', [req.params.id]);
  if (!auction) return res.status(404).json({ error: 'Trade not found' });
  const ano = String(auction.ano || '').trim();
  const removed = { lots: 0, lot_allocations: 0, invoices: 0, purchases: 0, bills: 0, debit_notes: 0 };
  const cnt = (sql, params) => {
    try { const r = db.get(sql, params); return Number(r && (r.c ?? r.count ?? 0)) || 0; } catch (_) { return 0; }
  };
  removed.lots            = cnt('SELECT COUNT(*) AS c FROM lots            WHERE auction_id = ?', [auction.id]);
  removed.lot_allocations = cnt('SELECT COUNT(*) AS c FROM lot_allocations WHERE auction_id = ?', [auction.id]);
  removed.invoices        = cnt('SELECT COUNT(*) AS c FROM invoices        WHERE auction_id = ? OR ano = ?', [auction.id, ano]);
  removed.purchases       = cnt('SELECT COUNT(*) AS c FROM purchases       WHERE auction_id = ? OR ano = ?', [auction.id, ano]);
  removed.bills           = cnt('SELECT COUNT(*) AS c FROM bills           WHERE auction_id = ? OR ano = ?', [auction.id, ano]);
  removed.debit_notes     = cnt('SELECT COUNT(*) AS c FROM debit_notes     WHERE ano = ?',                  [ano]);
  const safeRun = (sql, params) => { try { db.run(sql, params); } catch (e) { console.warn('[delete trade]', e.message); } };
  safeRun('DELETE FROM lot_allocations WHERE auction_id = ?', [auction.id]);
  safeRun('DELETE FROM lots            WHERE auction_id = ?', [auction.id]);
  safeRun('DELETE FROM invoices        WHERE auction_id = ? OR ano = ?', [auction.id, ano]);
  safeRun('DELETE FROM purchases       WHERE auction_id = ? OR ano = ?', [auction.id, ano]);
  safeRun('DELETE FROM bills           WHERE auction_id = ? OR ano = ?', [auction.id, ano]);
  safeRun('DELETE FROM debit_notes     WHERE ano = ?',                  [ano]);
  safeRun('DELETE FROM auctions        WHERE id  = ?', [auction.id]);
  res.json({ success: true, deleted: removed });
});

// ──────────────────────────────────────────────────────────────
// LOT ALLOCATIONS
// ──────────────────────────────────────────────────────────────
function parseLotNo(lot) {
  const match = String(lot).match(/^([A-Za-z]*)(\d+)$/);
  if (!match) return null;
  return { prefix: match[1].toUpperCase(), num: parseInt(match[2], 10), padLen: match[2].length };
}
function buildLotNo(prefix, num, padLen) {
  return prefix + String(num).padStart(padLen, '0');
}
function isLotInRange(lotNo, startLot, endLot) {
  const lot = parseLotNo(lotNo);
  const s = parseLotNo(startLot);
  const e = parseLotNo(endLot);
  if (!lot || !s || !e) return false;
  if (lot.prefix !== s.prefix || s.prefix !== e.prefix) return false;
  return lot.num >= s.num && lot.num <= e.num;
}
function rangeSize(startLot, endLot) {
  const s = parseLotNo(startLot);
  const e = parseLotNo(endLot);
  if (!s || !e) return 0;
  return e.num - s.num + 1;
}

app.get('/api/auctions/:id/allocations', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const allocations = db.all(
    'SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot',
    [auctionId]
  );
  res.json({ allocations });
});

app.post('/api/auctions/:id/allocations', requireAuctionWrite, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const { allocations } = req.body;
  if (!allocations || !Array.isArray(allocations) || !allocations.length) {
    return res.status(400).json({ error: 'At least one allocation is required' });
  }
  for (const a of allocations) {
    if (!a.branch || !a.start_lot || !a.end_lot) {
      return res.status(400).json({ error: 'Branch, start_lot, end_lot required for each allocation' });
    }
    const s = parseLotNo(a.start_lot);
    const e = parseLotNo(a.end_lot);
    if (!s || !e) return res.status(400).json({ error: `Invalid lot format: ${a.start_lot} or ${a.end_lot}. Use format like 001, A001` });
    if (s.prefix !== e.prefix) return res.status(400).json({ error: `Prefix mismatch: ${a.start_lot} vs ${a.end_lot}` });
    if (s.num > e.num) return res.status(400).json({ error: `Start (${a.start_lot}) must be <= End (${a.end_lot})` });
  }
  for (let i = 0; i < allocations.length; i++) {
    for (let j = i + 1; j < allocations.length; j++) {
      const a = allocations[i], b = allocations[j];
      const ap = parseLotNo(a.start_lot), ae = parseLotNo(a.end_lot);
      const bp = parseLotNo(b.start_lot), be = parseLotNo(b.end_lot);
      if (ap.prefix === bp.prefix && ap.num <= be.num && bp.num <= ae.num) {
        return res.status(400).json({ error: `Ranges overlap: ${a.branch} (${a.start_lot}-${a.end_lot}) and ${b.branch} (${b.start_lot}-${b.end_lot})` });
      }
    }
  }
  const usedLots = db.all('SELECT lot_no, branch FROM lots WHERE auction_id = ?', [auctionId]);
  const orphans = [];
  for (const ul of usedLots) {
    const covered = allocations.some(a => isLotInRange(ul.lot_no, a.start_lot, a.end_lot));
    if (!covered) orphans.push(ul.lot_no);
  }
  if (orphans.length > 0) {
    return res.status(400).json({
      error: `Cannot save — ${orphans.length} entered lot${orphans.length === 1 ? '' : 's'} would be orphaned: ${orphans.slice(0, 5).join(', ')}${orphans.length > 5 ? '…' : ''}`
    });
  }
  db.run('DELETE FROM lot_allocations WHERE auction_id = ?', [auctionId]);
  for (const a of allocations) {
    db.run(
      'INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
      [auctionId, a.branch, String(a.start_lot).trim(), String(a.end_lot).trim()]
    );
  }
  const saved = db.all(
    'SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot',
    [auctionId]
  );
  res.json({ allocations: saved });
});

app.get('/api/auctions/:id/allocation-stats', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const allocations = db.all(
    'SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot',
    [auctionId]
  );
  const lots = db.all(
    'SELECT lot_no, branch, name, amount FROM lots WHERE auction_id = ?',
    [auctionId]
  );
  const lotInfo = {};
  for (const l of lots) {
    lotInfo[l.lot_no] = { branch: l.branch || '', seller: l.name || '', booked: Number(l.amount) > 0 };
  }
  const stats = {};
  for (const a of allocations) {
    if (!stats[a.branch]) stats[a.branch] = { branch: a.branch, total: 0, used: 0, ranges: [] };
    const total = rangeSize(a.start_lot, a.end_lot);
    const usedInRange = lots.filter(l => isLotInRange(l.lot_no, a.start_lot, a.end_lot));
    stats[a.branch].total += total;
    stats[a.branch].used += usedInRange.length;
    const s = parseLotNo(a.start_lot);
    const e = parseLotNo(a.end_lot);
    const lotGrid = [];
    if (s && e) {
      for (let n = s.num; n <= e.num; n++) {
        const lotNo = buildLotNo(s.prefix, n, s.padLen);
        const info = lotInfo[lotNo];
        let state = 'free';
        if (info && info.booked) state = 'booked';
        else if (info)           state = 'allocated';
        lotGrid.push({
          lot: lotNo,
          used: !!info,
          booked: !!(info && info.booked),
          seller: info ? info.seller : '',
          branch: a.branch,
          state,
        });
      }
    }
    stats[a.branch].ranges.push({
      start: a.start_lot, end: a.end_lot, total,
      used: usedInRange.length, lots: lotGrid
    });
  }
  res.json({ stats: Object.values(stats), allocations });
});

app.post('/api/auctions/:id/reassign-lots', requireAuctionWrite, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id, 10);
  const { from_branch, to_branch, start_lot, end_lot } = req.body;
  if (!from_branch || !to_branch || !start_lot || !end_lot) {
    return res.status(400).json({ error: 'All fields required: from_branch, to_branch, start_lot, end_lot' });
  }
  if (from_branch === to_branch) return res.status(400).json({ error: 'FROM and TO branch must be different' });
  const s = parseLotNo(start_lot);
  const e = parseLotNo(end_lot);
  if (!s || !e) return res.status(400).json({ error: 'Invalid lot number format' });
  if (s.prefix !== e.prefix) return res.status(400).json({ error: 'Start and end must have same prefix' });
  if (s.num > e.num) return res.status(400).json({ error: 'Start must be <= End' });
  const fromAllocs = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ?', [auctionId, from_branch]);
  for (let n = s.num; n <= e.num; n++) {
    const lotNo = buildLotNo(s.prefix, n, s.padLen);
    const inRange = fromAllocs.some(a => isLotInRange(lotNo, a.start_lot, a.end_lot));
    if (!inRange) return res.status(400).json({ error: `Lot ${lotNo} is not allocated to ${from_branch}` });
  }
  const usedLots = db.all('SELECT lot_no FROM lots WHERE auction_id = ?', [auctionId]).map(l => l.lot_no);
  const usedSet = new Set(usedLots);
  const usedInRange = [];
  for (let n = s.num; n <= e.num; n++) {
    const lotNo = buildLotNo(s.prefix, n, s.padLen);
    if (usedSet.has(lotNo)) usedInRange.push(lotNo);
  }
  if (usedInRange.length > 0) {
    return res.status(400).json({
      error: `Cannot reassign — ${usedInRange.length} lot(s) already used: ${usedInRange.slice(0, 5).join(', ')}${usedInRange.length > 5 ? '...' : ''}`
    });
  }
  const fromAllocsAll = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ?', [auctionId, from_branch]);
  db.run('DELETE FROM lot_allocations WHERE auction_id = ? AND branch = ?', [auctionId, from_branch]);
  for (const alloc of fromAllocsAll) {
    const as = parseLotNo(alloc.start_lot);
    const ae = parseLotNo(alloc.end_lot);
    if (!as || !ae) continue;
    const overlapStart = Math.max(as.num, s.num);
    const overlapEnd = Math.min(ae.num, e.num);
    if (overlapStart > overlapEnd) {
      db.run('INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
        [auctionId, from_branch, alloc.start_lot, alloc.end_lot]);
    } else {
      if (as.num < overlapStart) {
        db.run('INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
          [auctionId, from_branch, buildLotNo(as.prefix, as.num, as.padLen), buildLotNo(as.prefix, overlapStart - 1, as.padLen)]);
      }
      if (ae.num > overlapEnd) {
        db.run('INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
          [auctionId, from_branch, buildLotNo(ae.prefix, overlapEnd + 1, ae.padLen), buildLotNo(ae.prefix, ae.num, ae.padLen)]);
      }
    }
  }
  db.run('INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
    [auctionId, to_branch, String(start_lot).trim(), String(end_lot).trim()]);
  const allocs = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot', [auctionId]);
  res.json({ success: true, allocations: allocs, message: `Lots ${start_lot}-${end_lot} reassigned from ${from_branch} to ${to_branch}` });
});

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

// Auction + lots import (large helper — preserved from previous build)
app.post('/api/auctions/import', requireAuctionWrite, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const workbook = XLSX.readFile(req.file.path);
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (!rows.length) throw new Error('File is empty');
    const db = getDb();
    const mode = req.body.mode || 'full';
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
    const overrideAno = req.body.ano;
    const overrideDate = normalizeDate(req.body.date);
    const cropType = req.body.crop_type || mapCol(rows[0], 'CRPT', 'CROP_TYPE', 'CROPTYPE') || (getSetting(db, 'default_crop_type') || 'VST');
    const state = req.body.state || mapCol(rows[0], 'STATE') || 'TAMIL NADU';
    const auctionCache = new Map();
    const resolveAuction = (ano, dateStr) => {
      const key = `${ano}|${dateStr}`;
      if (auctionCache.has(key)) return auctionCache.get(key);
      let auc = db.get('SELECT * FROM auctions WHERE ano = ? AND date = ?', [ano, dateStr]);
      if (!auc) {
        db.run('INSERT INTO auctions (ano, date, crop_type, state) VALUES (?,?,?,?)',
          [ano, dateStr || new Date().toISOString().slice(0, 10), cropType, state]);
        auc = db.get('SELECT * FROM auctions WHERE ano = ? AND date = ? ORDER BY id DESC LIMIT 1', [ano, dateStr]);
      }
      auctionCache.set(key, auc);
      return auc;
    };
    if (!overrideAno) {
      const firstAno = rows.length ? mapCol(rows[0], 'ANO', 'TNO', 'TRADE', 'TRADE_NO', 'TRADENO') : '';
      if (!firstAno) throw new Error('No ANO column found in file. Add ANO/TRADE/TRADE_NO column, or specify Trade No in the form to override.');
    }
    let imported = 0, updated = 0, skipped = 0;
    const skipReasons = [];
    const auctionStats = new Map();
    const isBlankRow = (row) => {
      const vals = Object.values(row);
      return !vals.length || vals.every(v => v === '' || v === null || v === undefined);
    };
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowNum = i + 2;
      if (isBlankRow(row)) continue;
      const rowAno = overrideAno || mapCol(row, 'ANO', 'TNO', 'TRADE', 'TRADE_NO', 'TRADENO');
      const rawDate = row.DATE !== undefined ? row.DATE : row.date !== undefined ? row.date : row.TRADE_DATE !== undefined ? row.TRADE_DATE : '';
      const rowDate = overrideDate || normalizeDate(rawDate);
      if (!rowAno) { skipped++; skipReasons.push({row: rowNum, lot: '', reason: 'Missing ANO/TRADE_NO for this row'}); continue; }
      const auc = resolveAuction(rowAno, rowDate);
      const auctionId = auc.id;
      const lotNo = mapCol(row, 'LOT', 'LOT_NO', 'LOTNO');
      if (!lotNo) { skipped++; skipReasons.push({row: rowNum, lot: '', reason: 'Missing LOT / LOT_NO column value'}); continue; }
      if (mode === 'price') {
        const existing = db.get('SELECT id FROM lots WHERE auction_id = ? AND lot_no = ?', [auctionId, lotNo]);
        if (!existing) { skipped++; skipReasons.push({row: rowNum, lot: lotNo, reason: `Lot ${lotNo} does not exist in Trade ${rowAno} (price-update requires existing lot)`}); continue; }
        try {
          const price = mapNum(row, 'PRICE');
          const qty   = mapNum(row, 'QTY', 'QUANTITY', 'WEIGHT', 'WT');
          const bag   = mapNum(row, 'BAG', 'BAGS', 'NO_OF_BAGS');
          let amount  = mapNum(row, 'AMOUNT', 'AMT', 'VALUE', 'TOTAL');
          if (!amount && qty && price) amount = qty * price;
          const sets = []; const vals = [];
          if (row.PRICE !== undefined || row.price !== undefined) { sets.push('price=?');  vals.push(price); }
          if (amount)                                              { sets.push('amount=?'); vals.push(amount); }
          if (row.QTY !== undefined || row.qty !== undefined)      { sets.push('qty=?');    vals.push(qty); }
          if (row.BAG !== undefined || row.bag !== undefined || row.BAGS !== undefined || row.bags !== undefined) { sets.push('bags=?'); vals.push(bag); }
          const codeVal  = mapCol(row, 'CODE', 'BUYER_CODE');
          if (codeVal) { sets.push('code=?'); vals.push(codeVal); }
          let resolvedBuyer  = mapCol(row, 'BUYER', 'BIDDER', 'BUYER_NAME');
          let resolvedBuyer1 = mapCol(row, 'BUYER1', 'TRADE_NAME', 'TRADENAME');
          let resolvedSale   = mapCol(row, 'SALE', 'SALE_TYPE');
          if (codeVal && (!resolvedBuyer || !resolvedBuyer1)) {
            const match = db.get(
              `SELECT buyer, buyer1, sale FROM buyers
               WHERE UPPER(TRIM(code)) = UPPER(TRIM(?)) OR UPPER(TRIM(ti)) = UPPER(TRIM(?)) OR UPPER(TRIM(buyer)) = UPPER(TRIM(?))
               LIMIT 1`,
              [codeVal, codeVal, codeVal]
            );
            if (match) {
              if (!resolvedBuyer)  resolvedBuyer  = match.buyer  || '';
              if (!resolvedBuyer1) resolvedBuyer1 = match.buyer1 || '';
              if (!resolvedSale)   resolvedSale   = match.sale   || '';
            } else {
              skipReasons.push({ row: rowNum, lot: lotNo, reason: `Warning: CODE "${codeVal}" not found in Buyers master — price updated but buyer NOT assigned.` });
            }
          }
          if (resolvedBuyer)  { sets.push('buyer=?');  vals.push(resolvedBuyer); }
          if (resolvedBuyer1) { sets.push('buyer1=?'); vals.push(resolvedBuyer1); }
          if (resolvedSale)   { sets.push('sale=?');   vals.push(resolvedSale); }
          if (!sets.length) { skipped++; skipReasons.push({row: rowNum, lot: lotNo, reason: 'Row has no updatable fields'}); continue; }
          vals.push(existing.id);
          db.run(`UPDATE lots SET ${sets.join(', ')} WHERE id=?`, vals);
          updated++;
          const key = `${rowAno}|${rowDate}`;
          auctionStats.set(key, (auctionStats.get(key) || 0) + 1);
        } catch (e) { skipped++; skipReasons.push({row: rowNum, lot: lotNo, reason: `DB error: ${e.message}`}); }
      } else {
        const existing = db.get('SELECT id FROM lots WHERE auction_id = ? AND lot_no = ?', [auctionId, lotNo]);
        if (existing) { skipped++; skipReasons.push({row: rowNum, lot: lotNo, reason: `Duplicate — lot ${lotNo} already exists in Trade ${rowAno}`}); continue; }
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
             mapCol(row, 'CROP'), mapCol(row, 'GRADE'),
             mapCol(row, 'CRPT', 'CROP_TYPE') || cropType,
             mapCol(row, 'BR', 'BRANCH', 'DEPOT'),
             mapCol(row, 'STATE') || state, traderId, sellerName,
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
             mapNum(row, 'CGST'), mapNum(row, 'SGST'), mapNum(row, 'IGST'),
             mapNum(row, 'ADVANCE', 'DISCOUNT'),
             mapNum(row, 'BALANCE', 'PAYABLE'),
             mapNum(row, 'BILAMT', 'BILL_AMT'),
             mapCol(row, 'USER_ID', 'USER') || 'import']);
          imported++;
          const key = `${rowAno}|${rowDate}`;
          auctionStats.set(key, (auctionStats.get(key) || 0) + 1);
        } catch (e) { skipped++; skipReasons.push({row: rowNum, lot: lotNo, reason: `DB error: ${e.message}`}); }
      }
    }
    const auctionBreakdown = [];
    for (const [key, count] of auctionStats) {
      const [ano, date] = key.split('|');
      const auc = auctionCache.get(key);
      auctionBreakdown.push({ id: auc?.id, ano, date, count });
    }
    auctionBreakdown.sort((a,b) => String(a.ano).localeCompare(String(b.ano), undefined, {numeric:true}));
    fs.unlink(req.file.path, () => {});
    res.json({ success: true, imported, updated, skipped, total: rows.length, auctionCount: auctionBreakdown.length, auctionBreakdown, skipReasons });
  } catch (e) {
    if (req.file) fs.unlink(req.file.path, () => {});
    res.status(400).json({ error: e.message });
  }
});

app.get('/api/auctions/template', requireExport, async (req, res) => {
  const db = getDb();
  const headers = ['ANO','DATE','LOT','CROP','GRADE','CRPT','BR','STATE','NAME','PADD','PPLA','PPIN','PSTATE','PST_CODE',
    'CR','PAN','TEL','AADHAR','BAG','LITRE','QTY','PRICE','AMOUNT','CODE','BUYER','BUYER1','SALE','INVO',
    'PQTY','PRATE','PURAMT','CGST','SGST','IGST','ADVANCE','BALANCE'];
  const cols = headers.map(h => ({ header: h, key: h.toLowerCase(), width: h.length < 5 ? 9 : 14 }));
  const defaultCrop  = getSetting(db, 'default_crop_type') || '';
  const bizState     = (getSetting(db, 'business_state') || 'TAMIL NADU').toUpperCase();
  const stCode       = bizState === 'KERALA' ? '32' : '33';
  const sample = [{
    ano: '1', date: '2026-04-15', lot: '001', crop: '', grade: '1',
    crpt: defaultCrop, br: '', state: bizState,
    name: 'SAMPLE SELLER', padd: '123 MAIN ST', ppla: '', ppin: '',
    pstate: bizState, pst_code: stCode, cr: 'CR.001', pan: 'ABCDE1234F', tel: '9876543210', aadhar: '',
    bag: 5, litre: '380', qty: 100.567, price: 0, amount: 0, code: '', buyer: '', buyer1: '', sale: '', invo: '',
    pqty: 0, prate: 0, puramt: 0, cgst: 0, sgst: 0, igst: 0, advance: 0, balance: 0,
  }];
  const buf = await createExcelBuffer('Lots', cols, sample, {
    db, title: 'AUCTION / LOTS TEMPLATE',
    metaLines: [`Date: ${fmtDate(todayLocalISO())}`],
  });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="auction-lots-template.xlsx"');
  res.send(Buffer.from(buf));
});

// ──────────────────────────────────────────────────────────────
// LOTS
// ──────────────────────────────────────────────────────────────
app.get('/api/lots/:auctionId', requireViewOrLotEntry, (req, res) => {
  const { branch, name, buyer, limit, offset, paginated, summary, search } = req.query;
  const db = getDb();
  let q = `SELECT lots.*,
             (SELECT b.code FROM buyers b WHERE b.buyer = lots.buyer LIMIT 1) AS buyer_code
           FROM lots
           WHERE lots.auction_id = ?`;
  const p = [req.params.auctionId];
  if (branch) { q += ' AND lots.branch = ?'; p.push(branch); }
  if (name)   { q += ' AND lots.name LIKE ?'; p.push(`%${name}%`); }
  if (buyer)  { q += ' AND lots.buyer = ?'; p.push(buyer); }
  const searchTerm = String(search || '').trim();
  if (searchTerm) {
    const wild = `%${searchTerm}%`;
    q += ` AND (
            COALESCE(lots.lot_no,'') LIKE ?
            OR COALESCE(lots.name,'')   LIKE ?
            OR COALESCE(lots.buyer,'')  LIKE ?
            OR COALESCE(lots.buyer1,'') LIKE ?
            OR COALESCE(lots.code,'')   LIKE ?
            OR COALESCE(lots.invo,'')   LIKE ?
            OR COALESCE(lots.branch,'') LIKE ?
            OR EXISTS (
              SELECT 1 FROM buyers b
               WHERE b.buyer = lots.buyer
                 AND COALESCE(b.code,'') LIKE ?
            )
          )`;
    p.push(wild, wild, wild, wild, wild, wild, wild, wild);
  }
  if (summary === '1') {
    let aggSql = `SELECT COUNT(*) AS n, COALESCE(SUM(CAST(bags AS INTEGER)), 0) AS bags, COALESCE(SUM(qty), 0) AS qty FROM lots WHERE lots.auction_id = ?`
      + (branch ? ' AND lots.branch = ?' : '')
      + (name   ? ' AND lots.name LIKE ?' : '')
      + (buyer  ? ' AND lots.buyer = ?' : '');
    if (searchTerm) {
      aggSql += ` AND (
            COALESCE(lots.lot_no,'') LIKE ?
            OR COALESCE(lots.name,'')   LIKE ?
            OR COALESCE(lots.buyer,'')  LIKE ?
            OR COALESCE(lots.buyer1,'') LIKE ?
            OR COALESCE(lots.code,'')   LIKE ?
            OR COALESCE(lots.invo,'')   LIKE ?
            OR COALESCE(lots.branch,'') LIKE ?
            OR EXISTS (
              SELECT 1 FROM buyers b
               WHERE b.buyer = lots.buyer
                 AND COALESCE(b.code,'') LIKE ?
            )
          )`;
    }
    const row = db.get(aggSql, p) || { n:0, bags:0, qty:0 };
    return res.json({ n: row.n, bags: row.bags, qty: row.qty });
  }
  if (paginated === '1') {
    const lim = Math.min(Math.max(parseInt(limit, 10) || 25, 1), 200);
    const off = Math.max(parseInt(offset, 10) || 0, 0);
    let cq = q.replace(/^SELECT[\s\S]+?FROM lots/, 'SELECT COUNT(*) AS n FROM lots');
    const total = (db.get(cq, p) || {}).n || 0;
    q += ' ORDER BY CAST(lots.lot_no AS INTEGER) DESC, lots.lot_no DESC LIMIT ? OFFSET ?';
    const rows = db.all(q, [...p, lim, off]);
    return res.json({ rows, total, limit: lim, offset: off });
  }
  q += ' ORDER BY lots.lot_no';
  res.json(db.all(q, p));
});

app.post('/api/lots', requireLotWrite, (req, res) => {
  const l = req.body;
  const db = getDb();
  const auctionId = parseInt(l.auction_id, 10);
  const lotNoStr  = String(l.lot_no || '').trim();
  const branch    = String(l.branch || '').trim();
  if (!auctionId || !lotNoStr) {
    return res.status(400).json({ error: 'auction_id and lot_no are required' });
  }
  const existing = db.get('SELECT id FROM lots WHERE auction_id = ? AND lot_no = ?', [auctionId, lotNoStr]);
  if (existing) {
    return res.status(409).json({ error: `Lot #${lotNoStr} already exists in this auction` });
  }
  const traderId = l.trader_id != null ? parseInt(l.trader_id, 10) : null;
  let trader = null;
  if (traderId) {
    trader = db.get('SELECT * FROM traders WHERE id = ?', [traderId]);
    if (!trader) {
      return res.status(400).json({ error: `Selected seller (trader id ${traderId}) no longer exists. Re-pick from the search.` });
    }
  }
  if (trader) {
    l.name   = trader.name   || l.name   || '';
    l.cr     = trader.cr     || '';
    l.pan    = trader.pan    || '';
    l.tel    = trader.tel    || '';
    l.aadhar = trader.aadhar || '';
    l.padd   = trader.padd   || '';
    l.ppla   = trader.ppla   || '';
    l.ppin   = trader.pin    || '';
    l.pstate = trader.pstate || l.pstate || '';
    l.pst_code = trader.pst_code || l.pst_code || '';
  }
  if (branch) {
    const allocs = db.all(
      'SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ?',
      [auctionId, branch]
    );
    if (allocs.length > 0) {
      const inRange = allocs.some(a => isLotInRange(lotNoStr, a.start_lot, a.end_lot));
      if (!inRange) {
        const ranges = allocs.map(a => a.start_lot + '-' + a.end_lot).join(', ');
        return res.status(400).json({ error: `Lot #${lotNoStr} is outside ${branch} allocation (${ranges})` });
      }
    }
  }
  db.run(`INSERT INTO lots (auction_id,lot_no,crop,grade,crpt,branch,state,trader_id,name,padd,ppla,ppin,pstate,pst_code,cr,pan,tel,aadhar,bags,litre,qty,gross_wt,sample_wt,moisture,user_id)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [auctionId,lotNoStr,l.crop||'',l.grade||'',l.crpt||'',branch,l.state||'TAMIL NADU',l.trader_id||null,l.name||'',l.padd||'',l.ppla||'',l.ppin||'',l.pstate||'',l.pst_code||'',l.cr||'',l.pan||'',l.tel||'',l.aadhar||'',l.bags||0,l.litre||'',l.qty||0,l.gross_wt||0,l.sample_wt||0,l.moisture||'',l.user_id||'']);
  pcClearGate(db, auctionId);
  res.json({ success: true });
});

const LOT_UPDATE_COLUMNS = new Set([
  'lot_no','crop','grade','crpt','branch','state','trader_id',
  'name','padd','ppla','ppin','pstate','pst_code','cr','pan','tel','aadhar',
  'bags','litre','qty','gross_wt','sample_wt','moisture',
  'price','amount','code','buyer','buyer1','sale','invo',
  'pqty','prate','puramt','com','sertax','cgst','sgst','igst',
  'dcgst','dsgst','digst','refud','refund','advance','balance','bilamt','paid',
  'user_id','asp_invo',
  'isp_pqty','isp_prate','isp_puramt','asp_pqty','asp_prate','asp_puramt'
]);

app.put('/api/lots/:id', requireLotWrite, (req, res) => {
  const l = req.body; const sets = []; const vals = [];
  const db = getDb();
  const lotId = parseInt(req.params.id, 10);
  const current = db.get('SELECT auction_id, lot_no, branch, locked_at FROM lots WHERE id = ?', [lotId]);
  if (current && current.locked_at && !isAdmin(req)) {
    return res.status(423).json({ error: 'This lot is locked — only an admin can edit it.' });
  }
  if (current) {
    const newLotNo = (l.lot_no != null) ? String(l.lot_no).trim() : current.lot_no;
    const newBranch = (l.branch != null) ? String(l.branch).trim() : current.branch;
    if (newLotNo !== current.lot_no) {
      const dup = db.get(
        'SELECT id FROM lots WHERE auction_id = ? AND lot_no = ? AND id != ?',
        [current.auction_id, newLotNo, lotId]
      );
      if (dup) return res.status(409).json({ error: `Lot #${newLotNo} already exists in this auction` });
    }
    if (newBranch && (newLotNo !== current.lot_no || newBranch !== current.branch)) {
      const allocs = db.all(
        'SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ?',
        [current.auction_id, newBranch]
      );
      if (allocs.length > 0) {
        const inRange = allocs.some(a => isLotInRange(newLotNo, a.start_lot, a.end_lot));
        if (!inRange) {
          const ranges = allocs.map(a => a.start_lot + '-' + a.end_lot).join(', ');
          return res.status(400).json({ error: `Lot #${newLotNo} is outside ${newBranch} allocation (${ranges})` });
        }
      }
    }
  }
  for (const [k,v] of Object.entries(l)) {
    if (!LOT_UPDATE_COLUMNS.has(k)) continue;
    sets.push(`${k}=?`); vals.push(v);
  }
  if (sets.length === 0) return res.json({ success: true });
  vals.push(lotId);
  db.run(`UPDATE lots SET ${sets.join(',')} WHERE id=?`, vals);
  if (current && current.auction_id) pcClearGate(db, current.auction_id);
  res.json({ success: true });
});

app.delete('/api/lots/:id', requireDelete, (req, res) => {
  const db = getDb();
  const cur = db.get('SELECT auction_id, locked_at FROM lots WHERE id = ?', [req.params.id]);
  if (cur && cur.locked_at && !isAdmin(req)) {
    return res.status(423).json({ error: 'This lot is locked — only an admin can delete it.' });
  }
  db.run('DELETE FROM lots WHERE id = ?', [req.params.id]);
  if (cur && cur.auction_id) pcClearGate(db, cur.auction_id);
  res.json({ success: true });
});

// ──────────────────────────────────────────────────────────────
// LOT LOCK / UNLOCK (Feature #7)
// ──────────────────────────────────────────────────────────────
function isAdmin(req) {
  return !!(req && req.user && req.user.role === 'admin');
}

function filterLockedLotIds(db, ids) {
  const nums = (ids || []).map(x => Number(x)).filter(Number.isFinite);
  if (!nums.length) return { allowed: [], skipped: [] };
  const placeholders = nums.map(() => '?').join(',');
  const locked = db.all(
    `SELECT id FROM lots WHERE id IN (${placeholders}) AND locked_at IS NOT NULL`,
    nums
  );
  const lockedSet = new Set(locked.map(r => Number(r.id)));
  const allowed = nums.filter(id => !lockedSet.has(id));
  const skipped = nums.filter(id =>  lockedSet.has(id));
  return { allowed, skipped };
}

function lotsLockedForInvoice(db, invoiceId) {
  const inv = db.get('SELECT auction_id, buyer FROM invoices WHERE id = ?', [invoiceId]);
  if (!inv) return false;
  const hit = db.get(
    `SELECT 1 FROM lots WHERE auction_id = ? AND buyer = ? AND locked_at IS NOT NULL LIMIT 1`,
    [inv.auction_id, inv.buyer || '']
  );
  return !!hit;
}
function lotsLockedForPurchase(db, purchaseId) {
  const pur = db.get('SELECT auction_id, name FROM purchases WHERE id = ?', [purchaseId]);
  if (!pur) return false;
  const hit = db.get(
    `SELECT 1 FROM lots WHERE auction_id = ? AND name = ? AND locked_at IS NOT NULL LIMIT 1`,
    [pur.auction_id, pur.name || '']
  );
  return !!hit;
}
function lotsLockedForDebitNote(db, noteId) {
  const dn = db.get('SELECT auction_id, name FROM debit_notes WHERE id = ?', [noteId]);
  if (!dn) return false;
  const hit = db.get(
    `SELECT 1 FROM lots WHERE auction_id = ? AND name = ? AND locked_at IS NOT NULL LIMIT 1`,
    [dn.auction_id, dn.name || '']
  );
  return !!hit;
}

app.post('/api/lots/lock', requireLotWrite, (req, res) => {
  const ids = Array.isArray(req.body && req.body.ids) ? req.body.ids : [];
  const numericIds = ids.map(x => Number(x)).filter(Number.isFinite);
  if (!numericIds.length) return res.status(400).json({ error: 'ids[] is required' });
  const db = getDb();
  const username = (req.user && req.user.username) || '';
  const placeholders = numericIds.map(() => '?').join(',');
  const info = db.run(
    `UPDATE lots SET locked_at = datetime('now','localtime'), locked_by = ?
      WHERE id IN (${placeholders}) AND locked_at IS NULL`,
    [username, ...numericIds]
  );
  res.json({ success: true, locked: (info && info.changes) || 0, requested: numericIds.length });
});

app.post('/api/lots/unlock', requireAdmin, (req, res) => {
  const ids = Array.isArray(req.body && req.body.ids) ? req.body.ids : [];
  const numericIds = ids.map(x => Number(x)).filter(Number.isFinite);
  if (!numericIds.length) return res.status(400).json({ error: 'ids[] is required' });
  const db = getDb();
  const placeholders = numericIds.map(() => '?').join(',');
  const info = db.run(
    `UPDATE lots SET locked_at = NULL, locked_by = NULL
      WHERE id IN (${placeholders}) AND locked_at IS NOT NULL`,
    numericIds
  );
  res.json({ success: true, unlocked: (info && info.changes) || 0, requested: numericIds.length });
});

// Calculate all lots for an auction
app.post('/api/lots/calculate/:auctionId',
  requireLotWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const lots = db.all('SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND locked_at IS NULL', [req.params.auctionId]);
  let count = 0;
  for (const lot of lots) {
    const calc = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [calc.pqty,calc.prate,calc.puramt,calc.com,calc.sertax,calc.cgst,calc.sgst,calc.igst,calc.advance,calc.balance,calc.bilamt,calc.refund||0,calc.refud||0,calc.isp_pqty||0,calc.isp_prate||0,calc.isp_puramt||0,calc.asp_pqty||0,calc.asp_prate||0,calc.asp_puramt||0,lot.id]);
    count++;
  }
  res.json({ success: true, calculated: count });
});

app.post('/api/lots/calculate-all', requireLotWrite, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const lots = db.all('SELECT * FROM lots WHERE amount > 0 AND locked_at IS NULL');
  let count = 0;
  for (const lot of lots) {
    const calc = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [calc.pqty,calc.prate,calc.puramt,calc.com,calc.sertax,calc.cgst,calc.sgst,calc.igst,calc.advance,calc.balance,calc.bilamt,calc.refund||0,calc.refud||0,calc.isp_pqty||0,calc.isp_prate||0,calc.isp_puramt||0,calc.asp_pqty||0,calc.asp_prate||0,calc.asp_puramt||0,lot.id]);
    count++;
  }
  res.json({ success: true, calculated: count });
});

app.get('/api/lots/validate/:auctionId', requireViewOrLotEntry, (req, res) => {
  const rows = getDb().all(
    `SELECT * FROM lots WHERE auction_id = ? AND (price = 0 OR amount = 0 OR buyer = '' OR code = '' OR ROUND(qty*price,2) <> ROUND(amount,2))`,
    [req.params.auctionId]);
  res.json(rows);
});

// ── Feature #12: Bulk grade update ──
app.post('/api/lots/bulk-grade', requireLotWrite, (req, res) => {
  try {
    const { ids, grade } = req.body || {};
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'ids[] is required' });
    }
    const validGrades = new Set(['', '1', '2', '3']);
    const g = String(grade == null ? '' : grade).trim();
    if (!validGrades.has(g)) {
      return res.status(400).json({ error: `grade must be one of: ${[...validGrades].map(v => v || '(blank)').join(', ')}` });
    }
    const numericIds = ids.map(x => Number(x)).filter(Number.isFinite);
    if (!numericIds.length) {
      return res.status(400).json({ error: 'ids[] contains no valid numeric ids' });
    }
    const db = getDb();
    const { allowed, skipped } = filterLockedLotIds(db, numericIds);
    if (!allowed.length) {
      return res.json({ success: true, updated: 0, grade: g, skipped_locked: skipped.length });
    }
    const placeholders = allowed.map(() => '?').join(',');
    db.run(`UPDATE lots SET grade = ? WHERE id IN (${placeholders})`, [g, ...allowed]);
    res.json({ success: true, updated: allowed.length, grade: g, skipped_locked: skipped.length });
  } catch (e) {
    res.status(500).json({ error: 'Bulk grade update failed: ' + (e.message || e) });
  }
});

// ── Feature #11: Bulk set buyer (price+code combined helper) ──
app.post('/api/lots/bulk-set-buyer', requireLotWrite, (req, res) => {
  const ids = Array.isArray(req.body.ids) ? req.body.ids.map(Number).filter(Number.isFinite) : [];
  if (!ids.length) return res.status(400).json({ error: 'No lot ids provided' });
  const code   = String(req.body.code   || '').trim();
  const buyer  = String(req.body.buyer  || '').trim();
  const buyer1 = String(req.body.buyer1 || '').trim();
  const sale   = String(req.body.sale   || '').trim();
  const hasPrice = req.body.price !== undefined && req.body.price !== null && req.body.price !== '';
  const priceNum = hasPrice ? Number(req.body.price) : null;
  if (hasPrice && !Number.isFinite(priceNum)) {
    return res.status(400).json({ error: 'price must be a number' });
  }
  if (!code && !hasPrice) {
    return res.status(400).json({ error: 'At least one of code or price is required' });
  }
  const db = getDb();
  const { allowed: mutableIds, skipped: lockedIds } = filterLockedLotIds(db, ids);
  const sets = [];
  const vals = [];
  if (code)   { sets.push('code = ?');   vals.push(code);   }
  if (buyer)  { sets.push('buyer = ?');  vals.push(buyer);  }
  if (buyer1) { sets.push('buyer1 = ?'); vals.push(buyer1); }
  if (req.body.sale !== undefined) { sets.push('sale = ?'); vals.push(sale); }
  if (hasPrice) { sets.push('price = ?'); vals.push(priceNum); }
  const CHUNK = 500;
  let updated = 0;
  const touchedAuctions = new Set();
  for (let i = 0; i < mutableIds.length; i += CHUNK) {
    const slice = mutableIds.slice(i, i + CHUNK);
    const placeholders = slice.map(() => '?').join(',');
    const affectedRows = db.all(
      `SELECT DISTINCT auction_id FROM lots WHERE id IN (${placeholders})`,
      slice
    );
    affectedRows.forEach(r => { if (r.auction_id) touchedAuctions.add(r.auction_id); });
    const info = db.run(
      `UPDATE lots SET ${sets.join(', ')} WHERE id IN (${placeholders})`,
      [...vals, ...slice]
    );
    if (info && typeof info.changes === 'number') updated += info.changes;
  }
  for (const aid of touchedAuctions) pcClearGate(db, aid);
  res.json({ success: true, updated, requested: ids.length, skipped_locked: lockedIds.length });
});

// Bulk buyer-by-code helper (looks up buyer master to fill buyer/buyer1/sale)
app.post('/api/lots/bulk-buyer', requireLotWrite, (req, res) => {
  try {
    const { ids, buyer } = req.body || {};
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'ids[] is required' });
    }
    const buyerCode = String(buyer == null ? '' : buyer).trim();
    if (!buyerCode) {
      return res.status(400).json({ error: 'buyer code is required' });
    }
    const numericIds = ids.map(x => Number(x)).filter(Number.isFinite);
    if (!numericIds.length) {
      return res.status(400).json({ error: 'ids[] contains no valid numeric ids' });
    }
    const db = getDb();
    const b = db.get(
      `SELECT buyer, buyer1, code, sale FROM buyers
        WHERE UPPER(TRIM(buyer)) = UPPER(TRIM(?)) LIMIT 1`,
      [buyerCode]
    );
    if (!b) {
      return res.status(404).json({ error: `No buyer found with code "${buyerCode}". Register the buyer first in the Buyers tab.` });
    }
    const buyerSale = String(b.sale || '').trim().toUpperCase();
    const saleVal = ['L', 'I', 'E'].includes(buyerSale) ? buyerSale : 'L';
    const { allowed: mutableIds, skipped: lockedIds } = filterLockedLotIds(db, numericIds);
    if (!mutableIds.length) {
      return res.json({
        success: true, updated: 0,
        buyer: b.buyer, buyer1: b.buyer1 || '', code: b.code || '', sale: saleVal,
        skipped_locked: lockedIds.length,
      });
    }
    const placeholders = mutableIds.map(() => '?').join(',');
    const touchedAuctions = new Set();
    const affectedRows = db.all(
      `SELECT DISTINCT auction_id FROM lots WHERE id IN (${placeholders})`,
      mutableIds
    );
    affectedRows.forEach(r => { if (r.auction_id) touchedAuctions.add(r.auction_id); });
    db.run(
      `UPDATE lots SET buyer = ?, buyer1 = ?, code = ?, sale = ? WHERE id IN (${placeholders})`,
      [b.buyer, b.buyer1 || '', b.code || '', saleVal, ...mutableIds]
    );
    for (const aid of touchedAuctions) pcClearGate(db, aid);
    res.json({
      success: true,
      updated: mutableIds.length,
      buyer: b.buyer,
      buyer1: b.buyer1 || '',
      code: b.code || '',
      sale: saleVal,
      skipped_locked: lockedIds.length,
    });
  } catch (e) {
    res.status(500).json({ error: 'Bulk buyer update failed: ' + (e.message || e) });
  }
});

// ── Feature #2: Bulk seller reassign (Change Seller) ──
app.post('/api/lots/bulk-seller', requireLotWrite, (req, res) => {
  try {
    const { ids, trader_id } = req.body || {};
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'ids[] is required' });
    }
    const tid = parseInt(trader_id, 10);
    if (!Number.isFinite(tid)) {
      return res.status(400).json({ error: 'trader_id must be a numeric id' });
    }
    const numericIds = ids.map(x => Number(x)).filter(Number.isFinite);
    if (!numericIds.length) {
      return res.status(400).json({ error: 'ids[] contains no valid numeric ids' });
    }
    const db = getDb();
    const t = db.get('SELECT * FROM traders WHERE id = ?', [tid]);
    if (!t) {
      return res.status(404).json({ error: `No seller found with id ${tid}. Refresh the seller list and try again.` });
    }
    const { allowed: mutableIds, skipped: lockedIds } = filterLockedLotIds(db, numericIds);
    if (!mutableIds.length) {
      return res.json({
        success: true, updated: 0, trader_id: tid, name: t.name || '',
        skipped_locked: lockedIds.length,
      });
    }
    const placeholders = mutableIds.map(() => '?').join(',');
    db.run(
      `UPDATE lots SET
         trader_id = ?, name = ?, cr = ?, pan = ?, tel = ?, aadhar = ?,
         padd = ?, ppla = ?, ppin = ?, pstate = ?, pst_code = ?
       WHERE id IN (${placeholders})`,
      [tid, t.name || '', t.cr || '', t.pan || '', t.tel || '', t.aadhar || '',
       t.padd || '', t.ppla || '', t.pin || '', t.pstate || '', t.pst_code || '',
       ...mutableIds]
    );
    res.json({
      success: true,
      updated: mutableIds.length,
      trader_id: tid,
      name: t.name || '',
      skipped_locked: lockedIds.length,
    });
  } catch (e) {
    res.status(500).json({ error: 'Bulk seller update failed: ' + (e.message || e) });
  }
});

// ──────────────────────────────────────────────────────────────
// INVOICES — Sales
// ──────────────────────────────────────────────────────────────
app.get('/api/invoices', requireView, (req, res) => {
  const { ano, auction_id, from, to, sale, search } = req.query;
  const db = getDb();
  const cfg = getSettingsFlat(db);
  const businessState = String(cfg.business_state || 'TAMIL NADU').toUpperCase();
  let q = 'SELECT * FROM invoices WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  const searchTerm = String(search || '').trim();
  if (searchTerm) {
    const wild = `%${searchTerm}%`;
    q += ` AND (COALESCE(invo,'') LIKE ? OR COALESCE(buyer,'') LIKE ? OR COALESCE(buyer1,'') LIKE ? OR COALESCE(gstin,'') LIKE ?)`;
    p.push(wild, wild, wild, wild);
  }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  if (sale && ['L','I','E'].includes(String(sale).toUpperCase())) {
    q += ' AND UPPER(sale) = ?';
    p.push(String(sale).toUpperCase());
  }
  q += ' ORDER BY date DESC, invo DESC LIMIT 500';
  const rows = db.all(q, p);
  const aspStmt = db.prepare(
    `SELECT DISTINCT asp_invo FROM lots
     WHERE auction_id = ? AND buyer = ? AND invo = ?
       AND asp_invo IS NOT NULL AND asp_invo != ''`
  );
  for (const r of rows) {
    const isASPRow = String(r.state || '').toLowerCase().includes('kerala');
    if (isASPRow) { r.asp_invo = ''; }
    else {
      const aspRows = aspStmt.all(r.auction_id, r.buyer, r.invo);
      r.asp_invo = aspRows.map(x => x.asp_invo).filter(Boolean).join(', ');
    }
  }
  try {
    const dispatchPin = String(cfg.tally_dispatch_pin || cfg.s_pin || cfg.kl_pin || cfg.tn_pin || '').trim();
    const routeStmt = db.prepare(
      `SELECT km FROM route_distances
       WHERE (from_pin = ? AND to_pin = ?) OR (from_pin = ? AND to_pin = ?)
       LIMIT 1`
    );
    const buyerPinStmt = db.prepare('SELECT pin, cpin FROM buyers WHERE buyer = ? LIMIT 1');
    for (const r of rows) {
      if (r.distance_km != null && r.distance_km !== '') {
        r.resolved_distance_km = Number(r.distance_km);
        continue;
      }
      const b = buyerPinStmt.get(r.buyer);
      const shipPin = b && b.cpin ? String(b.cpin).trim() : '';
      const billPin = b && b.pin  ? String(b.pin ).trim() : '';
      const buyerPin = shipPin || billPin;
      if (!buyerPin || !dispatchPin) { r.resolved_distance_km = null; continue; }
      const hit = routeStmt.get(dispatchPin, buyerPin, buyerPin, dispatchPin);
      r.resolved_distance_km = hit && hit.km != null ? Number(hit.km) : null;
    }
  } catch (e) {
    console.warn('[invoices] resolved distance hydration failed:', e.message);
  }
  res.json(rows);
});

app.post('/api/invoices/generate/:auctionId',
  requireInvoiceWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { saleType, buyerCode, invoiceNo } = req.body;
  if (!saleType || !buyerCode || !invoiceNo) {
    return res.status(400).json({ error: 'saleType, buyerCode, and invoiceNo are required' });
  }
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0) AND locked_at IS NULL`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  const invoice = buildSalesInvoice(db, req.params.auctionId, buyerCode, saleType, cfg);
  if (!invoice) return res.status(404).json({ error: `No lots found for buyer "${buyerCode}" in this auction. Make sure lots have this buyer code assigned.` });
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const dupBuyer = db.get(
    'SELECT id, invo FROM invoices WHERE auction_id = ? AND sale = ? AND buyer = ? LIMIT 1',
    [req.params.auctionId, saleType, buyerCode]
  );
  if (dupBuyer) {
    return res.status(409).json({
      error: `Invoice already exists for buyer "${buyerCode}" (${saleType}) in this trade — invoice #${dupBuyer.invo}.`,
      existingId: dupBuyer.id, existingInvo: dupBuyer.invo,
    });
  }
  const dupNo = db.get(
    'SELECT id, buyer FROM invoices WHERE auction_id = ? AND sale = ? AND invo = ? LIMIT 1',
    [req.params.auctionId, saleType, String(invoiceNo)]
  );
  if (dupNo) {
    return res.status(409).json({
      error: `Invoice number ${invoiceNo} (${saleType}) is already used in this trade by buyer "${dupNo.buyer}".`,
      existingId: dupNo.id, existingBuyer: dupNo.buyer,
    });
  }
  const s = invoice.summary;
  const invoiceState = cfg.business_state || auction.state || '';
  db.run(`INSERT INTO invoices (auction_id,ano,date,state,sale,invo,buyer,buyer1,gstin,place,bag,qty,amount,gunny,pava_hc,ins,cgst,sgst,igst,tcs,rund,tot,addl_chg,addl_name)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [req.params.auctionId,auction.ano,auction.date,invoiceState,saleType,String(invoiceNo),buyerCode,invoice.buyer.buyer1||'',
     invoice.buyer.gstin||'',invoice.buyer.pla||'',s.totalBags,s.totalQty,s.totalAmount,s.gunnyCost,s.transportCost,s.insuranceCost,
     s.cgst,s.sgst,s.igst,0,s.roundDiff,s.grandTotal,s.addlCharge||0,s.addlChargeName||'']);
  const isASPState = String(cfg.business_state || '').toUpperCase() === 'KERALA';
  for (const li of invoice.lineItems) {
    if (isASPState) {
      const existing = db.get(
        'SELECT invo, asp_invo FROM lots WHERE auction_id=? AND lot_no=? AND buyer=? LIMIT 1',
        [req.params.auctionId, li.lot, buyerCode]
      );
      const hasIspInvo = existing && existing.invo && existing.invo !== existing.asp_invo;
      if (hasIspInvo) {
        db.run('UPDATE lots SET asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
          [String(invoiceNo), req.params.auctionId, li.lot, buyerCode]);
      } else {
        db.run('UPDATE lots SET invo=?, asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=?',
          [String(invoiceNo), String(invoiceNo), req.params.auctionId, li.lot, buyerCode]);
      }
    } else {
      db.run('UPDATE lots SET sale=?, invo=? WHERE auction_id=? AND lot_no=? AND buyer=? AND locked_at IS NULL',
        [saleType, String(invoiceNo), req.params.auctionId, li.lot, buyerCode]);
    }
  }
  res.json({ success: true, invoice: invoice.summary });
});

app.get('/api/invoices/eligible-buyers/:auctionId', requireView, (req, res) => {
  const { saleType } = req.query;
  const db = getDb();
  const cfg = getSettingsFlat(db);
  const params = [req.params.auctionId];
  let saleClause = '';
  if (saleType) {
    saleClause = ` AND (COALESCE(NULLIF(l.sale,''), b.sale, 'L') = ?)`;
    params.push(saleType);
  }
  const isASPState = String(cfg.business_state || '').toUpperCase() === 'KERALA';
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
       AND l.locked_at IS NULL
       ${saleClause}
     GROUP BY l.buyer
     HAVING COUNT(CASE WHEN ${eligibleExpr} THEN 1 END) > 0
     ORDER BY l.buyer`,
    params
  ));
});

app.get('/api/invoices/eligibility-debug/:auctionId', requireView, (req, res) => {
  const db = getDb();
  const aid = req.params.auctionId;
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [aid]);
  if (!auction) return res.status(404).json({ error: 'Auction not found' });
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

app.post('/api/invoices/generate-all/:auctionId',
  requireInvoiceWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { startInvoiceNo, saleType } = req.body;
  let nextNo = parseInt(startInvoiceNo);
  if (!nextNo || nextNo < 1) return res.status(400).json({ error: 'startInvoiceNo must be a positive integer' });
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0) AND locked_at IS NULL`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
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
       AND l.locked_at IS NULL AND ${uninvoicedExpr} ${saleClause}`,
    params
  );
  if (!buyers.length) return res.status(404).json({ error: saleType ? `No un-invoiced buyers for sale type ${saleType}` : 'No un-invoiced buyers with lots in this auction' });
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const results = [];
  const errors = [];
  for (const row of buyers) {
    const useSaleType = saleType || row.default_sale || 'L';
    try {
      const dupBuyer = db.get(
        'SELECT id, invo FROM invoices WHERE auction_id = ? AND sale = ? AND buyer = ? LIMIT 1',
        [req.params.auctionId, useSaleType, row.buyer]
      );
      if (dupBuyer) {
        errors.push({ buyer: row.buyer, error: `Already invoiced as #${dupBuyer.invo}` });
        continue;
      }
      const invoice = buildSalesInvoice(db, req.params.auctionId, row.buyer, useSaleType, cfg);
      if (!invoice) { errors.push({ buyer: row.buyer, error: 'No matching lots' }); continue; }
      const s = invoice.summary;
      const invoNo = String(nextNo);
      const invoiceState = cfg.business_state || auction.state || '';
      const dupNo = db.get(
        'SELECT id, buyer FROM invoices WHERE auction_id = ? AND sale = ? AND invo = ? LIMIT 1',
        [req.params.auctionId, useSaleType, invoNo]
      );
      if (dupNo) {
        errors.push({ buyer: row.buyer, error: `Invoice #${invoNo} already used by ${dupNo.buyer}` });
        continue;
      }
      db.run(`INSERT INTO invoices (auction_id,ano,date,state,sale,invo,buyer,buyer1,gstin,place,bag,qty,amount,gunny,pava_hc,ins,cgst,sgst,igst,tcs,rund,tot,addl_chg,addl_name)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [req.params.auctionId,auction.ano,auction.date,invoiceState,useSaleType,invoNo,row.buyer,invoice.buyer.buyer1||'',
         invoice.buyer.gstin||'',invoice.buyer.pla||'',s.totalBags,s.totalQty,s.totalAmount,s.gunnyCost,s.transportCost,s.insuranceCost,
         s.cgst,s.sgst,s.igst,0,s.roundDiff,s.grandTotal,s.addlCharge||0,s.addlChargeName||'']);
      const isASPStateBulk = String(cfg.business_state || '').toUpperCase() === 'KERALA';
      for (const li of invoice.lineItems) {
        if (isASPStateBulk) {
          const existing = db.get(
            'SELECT invo, asp_invo FROM lots WHERE auction_id=? AND lot_no=? AND buyer=? LIMIT 1',
            [req.params.auctionId, li.lot, row.buyer]
          );
          const hasIspInvo = existing && existing.invo && existing.invo !== existing.asp_invo;
          if (hasIspInvo) {
            db.run('UPDATE lots SET asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=? AND locked_at IS NULL',
              [invoNo, req.params.auctionId, li.lot, row.buyer]);
          } else {
            db.run('UPDATE lots SET invo=?, asp_invo=? WHERE auction_id=? AND lot_no=? AND buyer=? AND locked_at IS NULL',
              [invoNo, invoNo, req.params.auctionId, li.lot, row.buyer]);
          }
        } else {
          db.run('UPDATE lots SET sale=?, invo=? WHERE auction_id=? AND lot_no=? AND buyer=? AND locked_at IS NULL',
            [useSaleType, invoNo, req.params.auctionId, li.lot, row.buyer]);
        }
      }
      results.push({ buyer: row.buyer, invoiceNo: invoNo, sale: useSaleType, grandTotal: s.grandTotal });
      nextNo++;
    } catch (e) { errors.push({ buyer: row.buyer, error: e.message }); }
  }
  res.json({ success: true, generated: results.length, results, errors });
});

// Lorry no bulk-set. Must precede the generic /:id PUT below.
app.put('/api/invoices/lorry-no', requireInvoiceWrite, (req, res) => {
  const { ids, lorry_no } = req.body || {};
  if (!Array.isArray(ids) || !ids.length) return res.status(400).json({ error: 'ids array required' });
  const cleanIds = ids.map(Number).filter(Number.isFinite);
  if (!cleanIds.length) return res.status(400).json({ error: 'No valid invoice IDs' });
  let v = null;
  if (lorry_no != null && String(lorry_no).trim() !== '') {
    v = String(lorry_no).trim().toUpperCase().replace(/\s+/g, '');
    if (v.length > 20) return res.status(400).json({ error: 'Lorry no too long (max 20 chars)' });
  }
  try {
    const db = getDb();
    let mutableIds = cleanIds;
    let skippedLocked = 0;
    if (!isAdmin(req)) {
      mutableIds = [];
      for (const id of cleanIds) {
        if (lotsLockedForInvoice(db, id)) { skippedLocked++; continue; }
        mutableIds.push(id);
      }
    }
    if (!mutableIds.length) {
      return res.json({ ok: true, updated: 0, lorry_no: v, skipped_locked: skippedLocked });
    }
    const placeholders = mutableIds.map(() => '?').join(',');
    const r = db.run(`UPDATE invoices SET lorry_no = ? WHERE id IN (${placeholders})`, [v, ...mutableIds]);
    res.json({ ok: true, updated: r.changes, lorry_no: v, skipped_locked: skippedLocked });
  } catch (e) {
    console.error('[lorry-no] bulk update failed:', e);
    res.status(500).json({ error: e.message });
  }
});

app.put('/api/invoices/:id', requireInvoiceWrite, (req, res) => {
  const i = req.body;
  const db = getDb();
  if (!isAdmin(req) && lotsLockedForInvoice(db, req.params.id)) {
    return res.status(423).json({ error: 'This invoice is locked because at least one of its lots is locked — only an admin can edit it.' });
  }
  const fields = ['ano','date','state','sale','invo','buyer','buyer1','gstin','place',
    'bag','qty','amount','gunny','pava_hc','ins','cgst','sgst','igst','tcs','rund','tot'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (i[f] !== undefined) { sets.push(`${f}=?`); vals.push(i[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  db.run(`UPDATE invoices SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

app.delete('/api/invoices/:id', requireDelete, (req, res) => {
  const db = getDb();
  const inv = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
  if (!inv) return res.status(404).json({ error: 'Invoice not found' });
  if (!isAdmin(req) && lotsLockedForInvoice(db, req.params.id)) {
    return res.status(423).json({ error: 'This invoice is locked because at least one of its lots is locked — only an admin can delete it.' });
  }
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

app.post('/api/invoices/:id/revert', requireInvoiceRevert, (req, res) => {
  const db = getDb();
  const inv = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
  if (!inv) return res.status(404).json({ error: 'Invoice not found' });
  if (!isAdmin(req) && lotsLockedForInvoice(db, req.params.id)) {
    return res.status(423).json({ error: 'This invoice is locked because at least one of its lots is locked — only an admin can revert it.' });
  }
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

app.post('/api/invoices/revert-all/:auctionId', requireInvoiceRevert, (req, res) => {
  const db = getDb();
  const aid = req.params.auctionId;
  const invoices = db.all('SELECT * FROM invoices WHERE auction_id = ?', [aid]);
  const admin = isAdmin(req);
  let lotsFreed = 0;
  let skippedLocked = 0;
  const revertedIds = [];
  for (const inv of invoices) {
    if (!admin && lotsLockedForInvoice(db, inv.id)) {
      skippedLocked++;
      continue;
    }
    const n = db.get('SELECT COUNT(*) as c FROM lots WHERE auction_id=? AND sale=? AND invo=? AND buyer=? AND locked_at IS NULL',
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]).c;
    lotsFreed += n;
    db.run(`UPDATE lots SET sale='', invo='' WHERE auction_id=? AND sale=? AND invo=? AND buyer=? AND locked_at IS NULL`,
      [inv.auction_id, inv.sale, inv.invo, inv.buyer]);
    db.run('DELETE FROM invoices WHERE id=?', [inv.id]);
    revertedIds.push(inv.id);
  }
  const orphan = db.get(
    `SELECT COUNT(*) as c FROM lots WHERE auction_id = ? AND invo IS NOT NULL AND invo != '' AND locked_at IS NULL`, [aid]
  ).c;
  if (orphan) {
    db.run(`UPDATE lots SET sale='', invo='' WHERE auction_id = ? AND locked_at IS NULL`, [aid]);
    lotsFreed += orphan;
  }
  res.json({ success: true, invoicesReverted: revertedIds.length, lotsFreed, skipped_locked: skippedLocked });
});

// Sales Invoice PDF
app.get('/api/invoices/pdf/:id', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const stored = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
    if (!stored) return res.status(404).json({ error: 'Invoice not found' });
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
          totalBags: stored.bag || 0, totalQty: stored.qty || 0, totalAmount: stored.amount || 0,
          gunnyCost: stored.gunny || 0, transportCost: stored.pava_hc || 0, insuranceCost: stored.ins || 0,
          cgst: stored.cgst || 0, sgst: stored.sgst || 0, igst: stored.igst || 0, tcs: stored.tcs || 0,
          roundDiff: stored.rund || 0,
          subtotalRounded: (stored.tot || 0) - (stored.addl_chg || 0),
          addlCharge: stored.addl_chg || 0, addlChargeName: stored.addl_name || '',
          grandTotal: stored.tot || 0, isInterState: stored.sale === 'I',
        }
      };
    }
    const dispatchedThrough = req.query.dispatchedThrough || '';
    if (dispatchedThrough) invoice.dispatchedThrough = dispatchedThrough;
    if (String(stored.state || '').toUpperCase() !== 'KERALA') {
      const aspRow = db.get(
        `SELECT asp_invo FROM lots WHERE auction_id = ? AND buyer = ? AND invo = ?
           AND asp_invo IS NOT NULL AND asp_invo != '' LIMIT 1`,
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

// Bulk Sales Invoice PDF
app.post('/api/invoices/pdf-bulk', requireView, async (req, res) => {
  try {
    const ids = Array.isArray(req.body?.ids) ? req.body.ids.map(Number).filter(Boolean) : [];
    if (!ids.length) return res.status(400).json({ error: 'No invoice IDs provided' });
    const dispatchedThrough = (req.body?.dispatchedThrough || '').toString();
    const db = getDb();
    const cfg = getSettingsFlat(db);
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
            roundDiff: stored.rund || 0,
            subtotalRounded: (stored.tot || 0) - (stored.addl_chg || 0),
            addlCharge: stored.addl_chg || 0, addlChargeName: stored.addl_name || '',
            grandTotal: stored.tot || 0, isInterState: stored.sale === 'I',
          }
        };
      }
      if (dispatchedThrough) invoice.dispatchedThrough = dispatchedThrough;
      if (String(stored.state || '').toUpperCase() !== 'KERALA') {
        const aspRow = db.get(
          `SELECT asp_invo FROM lots WHERE auction_id = ? AND buyer = ? AND invo = ?
             AND asp_invo IS NOT NULL AND asp_invo != '' LIMIT 1`,
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

// ASP Purchase-view PDF (stub — full version routes to generateSalesInvoicePDF with variant='purchase')
app.get('/api/invoices/purchase-pdf/:id', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const isASPContext = (String(cfg.business_mode || '').toLowerCase() === 'e-trade')
                      && (String(cfg.business_state || '').toUpperCase() === 'KERALA');
    if (!isASPContext) {
      return res.status(400).json({ error: 'Purchase view is only available when business state is Kerala (e-Trade mode).' });
    }
    const stored = db.get('SELECT * FROM invoices WHERE id=?', [req.params.id]);
    if (!stored) return res.status(404).json({ error: 'Invoice not found' });
    let invoice = stored.auction_id
      ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg) : null;
    if (!invoice) {
      const buyer = db.get('SELECT * FROM buyers WHERE buyer=? LIMIT 1', [stored.buyer]) || {};
      invoice = {
        buyer,
        lineItems: [{ lot: '—', grade: '', bags: stored.bag || 0, qty: stored.qty || 0, price: 0, amount: stored.amount || 0 }],
        summary: {
          totalBags: stored.bag || 0, totalQty: stored.qty || 0,
          totalAmount: stored.amount || 0, gunnyCost: stored.gunny || 0,
          transportCost: stored.pava_hc || 0, insuranceCost: stored.ins || 0,
          cgst: stored.cgst || 0, sgst: stored.sgst || 0, igst: stored.igst || 0,
          roundDiff: stored.rund || 0, grandTotal: stored.tot || 0,
          isInterState: stored.sale === 'I',
        }
      };
    }
    const pdf = await generateSalesInvoicePDF(invoice, cfg, stored.sale, stored.invo, stored.date, undefined, 'purchase');
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="PurchaseView_${stored.sale}_${stored.invo}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Purchase-view PDF error:', e);
    res.status(500).json({ error: 'PDF generation failed: ' + e.message });
  }
});

// Bulk purchase-view PDF
app.post('/api/invoices/purchase-pdf-bulk', requireView, async (req, res) => {
  try {
    const ids = Array.isArray(req.body?.ids) ? req.body.ids.map(Number).filter(Boolean) : [];
    if (!ids.length) return res.status(400).json({ error: 'No invoice IDs provided' });
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const isASPContext = (String(cfg.business_mode || '').toLowerCase() === 'e-trade')
                      && (String(cfg.business_state || '').toUpperCase() === 'KERALA');
    if (!isASPContext) {
      return res.status(400).json({ error: 'Purchase view is only available when business state is Kerala (e-Trade mode).' });
    }
    const payloads = [];
    for (const id of ids) {
      const stored = db.get('SELECT * FROM invoices WHERE id=?', [id]);
      if (!stored) continue;
      let invoice = stored.auction_id
        ? buildSalesInvoice(db, stored.auction_id, stored.buyer, stored.sale, cfg) : null;
      if (!invoice) {
        const buyer = db.get('SELECT * FROM buyers WHERE buyer=? LIMIT 1', [stored.buyer]) || {};
        invoice = {
          buyer,
          lineItems: [{ lot: '—', grade: '', bags: stored.bag || 0, qty: stored.qty || 0, price: 0, amount: stored.amount || 0 }],
          summary: {
            totalBags: stored.bag || 0, totalQty: stored.qty || 0,
            totalAmount: stored.amount || 0, gunnyCost: stored.gunny || 0,
            transportCost: stored.pava_hc || 0, insuranceCost: stored.ins || 0,
            cgst: stored.cgst || 0, sgst: stored.sgst || 0, igst: stored.igst || 0,
            roundDiff: stored.rund || 0, grandTotal: stored.tot || 0,
            isInterState: stored.sale === 'I',
          }
        };
      }
      payloads.push({ invoiceData: invoice, saleType: stored.sale, invoiceNo: stored.invo, invoiceDate: stored.date });
    }
    if (!payloads.length) return res.status(404).json({ error: 'No invoices resolved from the provided IDs' });
    const pdf = await generateSalesInvoicesBatchPDF(payloads, cfg, 'purchase');
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="PurchaseView_Batch_${payloads.length}.pdf"`);
    res.send(pdf);
  } catch (e) {
    console.error('Bulk purchase-view PDF error:', e);
    res.status(500).json({ error: 'Batch PDF generation failed: ' + e.message });
  }
});

// ──────────────────────────────────────────────────────────────
// PURCHASES (RD)
// ──────────────────────────────────────────────────────────────
app.get('/api/purchases', requireView, (req, res) => {
  const { auction_id, ano, from, to, sale, search } = req.query;
  let q = 'SELECT * FROM purchases WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  const searchTerm = String(search || '').trim();
  if (searchTerm) {
    const wild = `%${searchTerm}%`;
    q += ` AND (COALESCE(invo,'') LIKE ? OR COALESCE(name,'') LIKE ? OR COALESCE(gstin,'') LIKE ?)`;
    p.push(wild, wild, wild);
  }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  const saleNorm = String(sale || '').trim().toUpperCase();
  if (saleNorm === 'L') {
    q += ' AND COALESCE(igst,0) = 0 AND (COALESCE(cgst,0) > 0 OR COALESCE(sgst,0) > 0)';
  } else if (saleNorm === 'I') {
    q += ' AND COALESCE(igst,0) > 0';
    q += ` AND NOT EXISTS (SELECT 1 FROM lots l WHERE l.auction_id = purchases.auction_id
            AND UPPER(TRIM(COALESCE(l.name,''))) = UPPER(TRIM(COALESCE(purchases.name,'')))
            AND UPPER(TRIM(COALESCE(l.sale,''))) = 'E')`;
  } else if (saleNorm === 'E') {
    q += ' AND COALESCE(igst,0) > 0';
    q += ` AND EXISTS (SELECT 1 FROM lots l WHERE l.auction_id = purchases.auction_id
            AND UPPER(TRIM(COALESCE(l.name,''))) = UPPER(TRIM(COALESCE(purchases.name,'')))
            AND UPPER(TRIM(COALESCE(l.sale,''))) = 'E')`;
  }
  q += ' ORDER BY date DESC LIMIT 500';
  res.json(getDb().all(q, p));
});

function _normGstin(s) {
  let v = String(s == null ? '' : s).trim().toUpperCase();
  if (v.startsWith('GSTIN.')) v = v.slice(6);
  else if (v.startsWith('GSTIN')) v = v.slice(5);
  return v.trim();
}
function _resolveBuyerGstin(cfg) {
  const stateUpper = String(cfg && cfg.business_state || '').toUpperCase();
  const isKerala = stateUpper === 'KERALA';
  const candidates = isKerala
    ? [cfg.kl_gstin, cfg.tn_gstin, cfg.gstin, cfg.business_gstin]
    : [cfg.tn_gstin, cfg.kl_gstin, cfg.gstin, cfg.business_gstin];
  for (const c of candidates) {
    const norm = _normGstin(c);
    if (norm) return norm;
  }
  try {
    const id = getCompanyIdentity(cfg);
    return _normGstin(id && id.gstin);
  } catch (_) { return ''; }
}
function _normName(s) {
  return String(s == null ? '' : s).trim().toUpperCase().replace(/\s+/g, ' ');
}

app.post('/api/purchases/generate/:auctionId',
  requireInvoiceWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { sellerName, invoiceNo } = req.body;
  const invoice = buildPurchaseInvoice(db, req.params.auctionId, sellerName, cfg);
  if (!invoice) return res.status(404).json({ error: 'No data for this seller' });
  const sellerGstin = _normGstin(invoice.seller && (invoice.seller.cr || invoice.seller.gstin));
  const buyerGstin  = _resolveBuyerGstin(cfg);
  if (sellerGstin && buyerGstin && sellerGstin === buyerGstin) {
    return res.status(400).json({
      error: `Seller GSTIN (${sellerGstin}) matches the buyer/company GSTIN — same legal entity, no purchase invoice can be raised. Treat as an internal stock transfer.`,
      sellerGstin, buyerGstin,
    });
  }
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const dupSeller = db.get(
    'SELECT id, invo FROM purchases WHERE auction_id = ? AND name = ? LIMIT 1',
    [req.params.auctionId, invoice.seller.name]
  );
  if (dupSeller) {
    return res.status(409).json({
      error: `Purchase already exists for dealer "${invoice.seller.name}" in this trade — invoice #${dupSeller.invo}.`,
      existingId: dupSeller.id, existingInvo: dupSeller.invo,
    });
  }
  const dupNo = db.get(
    'SELECT id, name FROM purchases WHERE auction_id = ? AND invo = ? LIMIT 1',
    [req.params.auctionId, String(invoiceNo)]
  );
  if (dupNo) {
    return res.status(409).json({
      error: `Purchase invoice number ${invoiceNo} is already used in this trade by dealer "${dupNo.name}".`,
      existingId: dupNo.id, existingDealer: dupNo.name,
    });
  }
  const s = invoice.summary;
  db.run(`INSERT INTO purchases (auction_id,ano,date,state,br,name,add_line,place,gstin,invo,qty,amount,cgst,sgst,igst,rund,total,tds)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [req.params.auctionId,auction.ano,auction.date,auction.state||'','',invoice.seller.name,invoice.seller.address||'',
     invoice.seller.place||'',invoice.seller.cr||'',String(invoiceNo),s.totalQty,s.totalPuramt,
     s.totalCgst,s.totalSgst,s.totalIgst,s.roundDiff,s.grandTotal,s.tdsAmount]);
  res.json({ success: true, invoice: s });
});

app.get('/api/purchases/eligible-sellers/:auctionId', requireView, (req, res) => {
  const cfgEs = getSettingsFlat(getDb());
  const buyerGstinEs = _resolveBuyerGstin(cfgEs);
  const allSellers = getDb().all(
    `SELECT name, COUNT(*) as lot_count, SUM(qty) as total_qty, SUM(amount) as total_amount, MAX(cr) as cr
     FROM lots
     WHERE auction_id = ? AND name IS NOT NULL AND name != ''
       AND amount > 0
       AND (UPPER(cr) LIKE 'GSTIN%' OR (cr GLOB '[0-9][0-9]*' AND LENGTH(cr) >= 15))
     GROUP BY name ORDER BY name`,
    [req.params.auctionId]
  );
  const eligible = buyerGstinEs
    ? allSellers.filter(r => _normGstin(r.cr) !== buyerGstinEs)
    : allSellers;
  res.json(eligible);
});

app.post('/api/purchases/generate-all/:auctionId',
  requireInvoiceWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { startInvoiceNo } = req.body;
  let nextNo = parseInt(startInvoiceNo);
  if (!nextNo || nextNo < 1) return res.status(400).json({ error: 'startInvoiceNo must be a positive integer' });
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  const sellers = db.all(
    `SELECT DISTINCT name FROM lots
     WHERE auction_id = ? AND amount > 0 AND name IS NOT NULL AND name != ''
       AND (UPPER(cr) LIKE 'GSTIN%' OR (cr GLOB '[0-9][0-9]*' AND LENGTH(cr) >= 15))`,
    [req.params.auctionId]
  );
  if (!sellers.length) return res.status(404).json({ error: 'No registered dealers (with GSTIN) in this auction' });
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [req.params.auctionId]);
  const results = [];
  const errors = [];
  const buyerGstinAll = _resolveBuyerGstin(cfg);
  for (const row of sellers) {
    try {
      const dupSeller = db.get(
        'SELECT id, invo FROM purchases WHERE auction_id = ? AND name = ? LIMIT 1',
        [req.params.auctionId, row.name]
      );
      if (dupSeller) {
        errors.push({ seller: row.name, error: `Already invoiced as #${dupSeller.invo}` });
        continue;
      }
      const invoice = buildPurchaseInvoice(db, req.params.auctionId, row.name, cfg);
      if (!invoice) { errors.push({ seller: row.name, error: 'Build failed' }); continue; }
      const sellerGstinAll = _normGstin(invoice.seller && (invoice.seller.cr || invoice.seller.gstin));
      if (sellerGstinAll && buyerGstinAll && sellerGstinAll === buyerGstinAll) {
        errors.push({ seller: row.name, error: `Same-entity (GSTIN ${sellerGstinAll}) — no purchase invoice raised` });
        continue;
      }
      const s = invoice.summary;
      const invoNo = String(nextNo);
      const dupNo = db.get(
        'SELECT id, name FROM purchases WHERE auction_id = ? AND invo = ? LIMIT 1',
        [req.params.auctionId, invoNo]
      );
      if (dupNo) {
        errors.push({ seller: row.name, error: `Invoice #${invoNo} already used by ${dupNo.name}` });
        continue;
      }
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

app.put('/api/purchases/:id', requireInvoiceWrite, (req, res) => {
  const p = req.body;
  const db = getDb();
  if (!isAdmin(req) && lotsLockedForPurchase(db, req.params.id)) {
    return res.status(423).json({ error: 'This purchase is locked because at least one of its lots is locked — only an admin can edit it.' });
  }
  const fields = ['ano','date','state','br','name','add_line','place','gstin','invo',
    'qty','amount','cgst','sgst','igst','rund','total','tds'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (p[f] !== undefined) { sets.push(`${f}=?`); vals.push(p[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  db.run(`UPDATE purchases SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

app.delete('/api/purchases/:id', requireDelete, (req, res) => {
  const db = getDb();
  if (!isAdmin(req) && lotsLockedForPurchase(db, req.params.id)) {
    return res.status(423).json({ error: 'This purchase is locked because at least one of its lots is locked — only an admin can delete it.' });
  }
  db.run('DELETE FROM purchases WHERE id=?', [req.params.id]);
  res.json({ success: true });
});

function enrichPurchaseForPDF(invoice, cfg, db, auctionId) {
  if (!invoice) return invoice;
  if (!invoice.invoiceDate && auctionId) {
    const auction = db.get('SELECT date FROM auctions WHERE id = ?', [auctionId]);
    if (auction && auction.date) {
      const d = new Date(auction.date);
      if (!isNaN(d)) invoice.invoiceDate = fmtDate(auction.date);
    }
  }
  if (!invoice.invoiceDate) invoice.invoiceDate = fmtDate(todayLocalISO());
  if (!invoice.eTradeNo) invoice.eTradeNo = String(auctionId || '');
  if (!invoice.buyer) {
    const _ident = getCompanyIdentity(cfg);
    const isKerala = String(cfg.business_state || '').toUpperCase() === 'KERALA';
    if (isKerala) {
      invoice.buyer = {
        name:    _ident.name || cfg.short_name || cfg.trade_name || '',
        address: cfg.kl_address1 || _ident.address1 || '',
        place:   cfg.kl_place || '', pin: cfg.kl_pin || '',
        state:   cfg.kl_state || _ident.state || 'Kerala', st_code: '32',
        gstin:   cfg.kl_gstin || _ident.gstin || '',
        pan:     _ident.pan || cfg.pan || '',
      };
    } else {
      invoice.buyer = {
        name:    _ident.name || cfg.short_name || cfg.trade_name || '',
        address: cfg.tn_address1 || _ident.address1 || '',
        place:   cfg.tn_place || '', pin: cfg.tn_pin || '',
        state:   cfg.tn_state || _ident.state || 'Tamil Nadu',
        st_code: cfg.tn_st_code || _ident.stateCode || '33',
        gstin:   cfg.tn_gstin || _ident.gstin || '',
        pan:     _ident.pan || cfg.pan || '',
      };
    }
  }
  return invoice;
}

app.get('/api/purchases/pdf/:auctionId/:sellerName', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const sellerName = decodeURIComponent(req.params.sellerName);
    const auctionId = req.params.auctionId;
    const invoiceNo = req.query.invoiceNo || '001';
    let invoice = buildPurchaseInvoice(db, auctionId, sellerName, cfg);
    if (!invoice) {
      let stored = db.get(
        `SELECT * FROM purchases WHERE auction_id = ? AND name = ? AND invo = ? LIMIT 1`,
        [auctionId, sellerName, String(invoiceNo)]
      );
      if (!stored) {
        stored = db.get(`SELECT * FROM purchases WHERE name = ? AND invo = ? LIMIT 1`, [sellerName, String(invoiceNo)]);
      }
      if (!stored) {
        return res.status(404).json({ error: `No purchase data found for seller "${sellerName}" with invoice ${invoiceNo}. Lots may have been deleted.` });
      }
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

app.post('/api/purchases/pdf-bulk', requireView, async (req, res) => {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const ids = Array.isArray(req.body?.ids) ? req.body.ids.filter(Boolean) : [];
    if (!ids.length) return res.status(400).json({ error: 'No purchase IDs provided' });
    const placeholders = ids.map(() => '?').join(',');
    const rows = db.all(`SELECT * FROM purchases WHERE id IN (${placeholders})`, ids);
    if (!rows.length) return res.status(404).json({ error: 'No matching purchases found' });
    const byId = new Map(rows.map(r => [r.id, r]));
    const ordered = ids.map(id => byId.get(Number(id))).filter(Boolean);
    const payloads = [];
    for (const stored of ordered) {
      let invoice = stored.auction_id ? buildPurchaseInvoice(db, stored.auction_id, stored.name, cfg) : null;
      if (!invoice) {
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

// ──────────────────────────────────────────────────────────────
// BILLS (Agriculturist Bills of Supply)
// ──────────────────────────────────────────────────────────────
app.get('/api/bills', requireView, (req, res) => {
  const { auction_id, ano, from, to, branch, search } = req.query;
  let q = 'SELECT * FROM bills WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  const searchTerm = String(search || '').trim();
  if (searchTerm) {
    const wild = `%${searchTerm}%`;
    q += ` AND (COALESCE(bil,'') LIKE ? OR COALESCE(name,'') LIKE ?)`;
    p.push(wild, wild);
  }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  const branchFilter = String(branch || '').trim();
  if (branchFilter) {
    q += ` AND (
            UPPER(TRIM(COALESCE(br,''))) = UPPER(TRIM(?))
            OR EXISTS (
              SELECT 1 FROM lots l
               WHERE l.auction_id = COALESCE(bills.auction_id, (SELECT a.id FROM auctions a WHERE a.ano = bills.ano LIMIT 1))
                 AND UPPER(TRIM(COALESCE(l.name,''))) = UPPER(TRIM(COALESCE(bills.name,'')))
                 AND UPPER(TRIM(COALESCE(l.branch,''))) = UPPER(TRIM(?))
            )
          )`;
    p.push(branchFilter, branchFilter);
  }
  q += ' ORDER BY date DESC, bil DESC LIMIT 500';
  res.json(withFmtDate(getDb().all(q, p)));
});

app.post('/api/bills/generate/:auctionId',
  requireInvoiceWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { sellerName, billNo } = req.body;
  if (!sellerName || !billNo) return res.status(400).json({ error: 'sellerName and billNo are required' });
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
  const dupSeller = db.get(
    'SELECT id, bil FROM bills WHERE auction_id = ? AND name = ? LIMIT 1',
    [req.params.auctionId, bill.seller.name]
  );
  if (dupSeller) {
    return res.status(409).json({
      error: `Bill of supply already exists for seller "${bill.seller.name}" in this trade — bill #${dupSeller.bil}.`,
      existingId: dupSeller.id, existingBil: dupSeller.bil,
    });
  }
  const dupNo = db.get(
    'SELECT id, name FROM bills WHERE auction_id = ? AND bil = ? LIMIT 1',
    [req.params.auctionId, parseInt(billNo)]
  );
  if (dupNo) {
    return res.status(409).json({
      error: `Bill number ${billNo} is already used in this trade by seller "${dupNo.name}".`,
      existingId: dupNo.id, existingSeller: dupNo.name,
    });
  }
  const s = bill.summary;
  db.run(`INSERT INTO bills (auction_id,ano,date,state,br,crpt,bil,name,add_line,pla,pstate,st_code,crr,pan,qty,cost,igst,net)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [req.params.auctionId,auction.ano,auction.date,auction.state||'','',auction.crop_type||(getSetting(getDb(),'default_crop_type')||'VST'),
     parseInt(billNo),bill.seller.name,bill.seller.address||'',bill.seller.place||'',
     bill.seller.state||'',bill.seller.st_code||'',bill.seller.cr||'',bill.seller.pan||'',
     s.totalQty,s.totalPuramt,0,s.netAmount]);
  res.json({ success: true, bill: s });
});

app.get('/api/bills/eligible-sellers/:auctionId', requireView, (req, res) => {
  res.json(listAgriSellers(getDb(), req.params.auctionId));
});

app.post('/api/bills/generate-all/:auctionId',
  requireInvoiceWrite,
  requirePriceChecked(req => parseInt(req.params.auctionId, 10)),
  (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { startBillNo } = req.body;
  let nextNo = parseInt(startBillNo);
  if (!nextNo || nextNo < 1) return res.status(400).json({ error: 'startBillNo must be a positive integer' });
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
      const dupSeller = db.get(
        'SELECT id, bil FROM bills WHERE auction_id = ? AND name = ? LIMIT 1',
        [req.params.auctionId, row.name]
      );
      if (dupSeller) {
        errors.push({ seller: row.name, error: `Already billed as #${dupSeller.bil}` });
        continue;
      }
      const bill = buildAgriBill(db, req.params.auctionId, row.name, cfg);
      if (!bill || bill.error) { errors.push({ seller: row.name, error: bill?.error || 'Build failed' }); continue; }
      const s = bill.summary;
      const dupNo = db.get(
        'SELECT id, name FROM bills WHERE auction_id = ? AND bil = ? LIMIT 1',
        [req.params.auctionId, nextNo]
      );
      if (dupNo) {
        errors.push({ seller: row.name, error: `Bill #${nextNo} already used by ${dupNo.name}` });
        continue;
      }
      db.run(`INSERT INTO bills (auction_id,ano,date,state,br,crpt,bil,name,add_line,pla,pstate,st_code,crr,pan,qty,cost,igst,net)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
        [req.params.auctionId,auction.ano,auction.date,auction.state||'','',auction.crop_type||(getSetting(getDb(),'default_crop_type')||'VST'),
         nextNo,bill.seller.name,bill.seller.address||'',bill.seller.place||'',
         bill.seller.state||'',bill.seller.st_code||'',bill.seller.cr||'',bill.seller.pan||'',
         s.totalQty,s.totalPuramt,0,s.netAmount]);
      results.push({ seller: row.name, billNo: nextNo, netAmount: s.netAmount });
      nextNo++;
    } catch (e) { errors.push({ seller: row.name, error: e.message }); }
  }
  res.json({ success: true, generated: results.length, results, errors });
});

app.get('/api/bills/pdf/:auctionId/:sellerName', requireView, async (req, res) => {
  try {
    const db = getDb(); const cfg = getSettingsFlat(db);
    const sellerName = decodeURIComponent(req.params.sellerName);
    const billNo = req.query.billNo || '001';
    let bill = buildAgriBill(db, req.params.auctionId, sellerName, cfg);
    if (!bill || bill.error) {
      const stored = db.get('SELECT * FROM bills WHERE name = ? AND bil = ? LIMIT 1', [sellerName, parseInt(billNo)]);
      if (!stored) return res.status(404).json({ error: bill?.error || `No bill data found for "${sellerName}"` });
      bill = {
        seller: { name: stored.name, address: stored.add_line, place: stored.pla, state: stored.pstate, st_code: stored.st_code, cr: stored.crr, crno: stored.crr, pan: stored.pan },
        lineItems: [{ lot: '—', qty: stored.qty, pqty: stored.qty, prate: 0, amount: stored.cost, puramt: stored.cost }],
        summary: { totalQty: stored.qty, totalPuramt: stored.cost, roundDiff: 0, netAmount: stored.net, cgst: 0, sgst: 0, igst: 0, tax: 0 }
      };
    }
    if (bill.seller && !bill.seller.crno) bill.seller.crno = bill.seller.cr || '';
    const auction = db.get('SELECT date FROM auctions WHERE id = ?', [req.params.auctionId]);
    if (auction && auction.date) {
      const d = new Date(auction.date);
      if (!isNaN(d)) bill.billDate = fmtDate(auction.date);
    }
    if (!bill.billDate) bill.billDate = fmtDate(todayLocalISO());
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
      let bill = stored.auction_id ? buildAgriBill(db, stored.auction_id, stored.name, cfg) : null;
      if (!bill || bill.error) {
        bill = {
          seller: { name: stored.name, address: stored.add_line, place: stored.pla, state: stored.pstate, st_code: stored.st_code, cr: stored.crr, crno: stored.crr, pan: stored.pan },
          lineItems: [{ lot: '—', qty: stored.qty, pqty: stored.qty, prate: 0, amount: stored.cost, puramt: stored.cost }],
          summary: { totalQty: stored.qty, totalPuramt: stored.cost, roundDiff: 0, netAmount: stored.net, cgst: 0, sgst: 0, igst: 0, tax: 0 }
        };
      }
      if (bill.seller && !bill.seller.crno) bill.seller.crno = bill.seller.cr || '';
      if (stored.auction_id) {
        const auction = db.get('SELECT date FROM auctions WHERE id = ?', [stored.auction_id]);
        if (auction && auction.date) {
          const d = new Date(auction.date);
          if (!isNaN(d)) bill.billDate = fmtDate(auction.date);
        }
      }
      if (!bill.billDate) bill.billDate = fmtDate(todayLocalISO());
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

// ──────────────────────────────────────────────────────────────
// DEBIT NOTES
// ──────────────────────────────────────────────────────────────
app.get('/api/debit-notes', requireView, (req, res) => {
  const { auction_id, ano, from, to, search } = req.query;
  let q = 'SELECT * FROM debit_notes WHERE 1=1'; const p = [];
  if (auction_id) { q += ' AND auction_id = ?'; p.push(parseInt(auction_id)); }
  if (ano) { q += ' AND ano = ?'; p.push(ano); }
  const searchTerm = String(search || '').trim();
  if (searchTerm) {
    const wild = `%${searchTerm}%`;
    q += ` AND (COALESCE(note_no,'') LIKE ? OR COALESCE(name,'') LIKE ?)`;
    p.push(wild, wild);
  }
  if (from && to) { q += ' AND date BETWEEN ? AND ?'; p.push(from, to); }
  q += ' ORDER BY date DESC, note_no DESC LIMIT 500';
  res.json(withFmtDate(getDb().all(q, p)));
});

app.post('/api/debit-notes/generate', requireInvoiceWrite, (req, res) => {
  const db = getDb();
  const cfg = getSettingsFlat(db);
  const purchno = String(req.body.purchno || req.body.invoiceNo || '').trim();
  const ano     = String(req.body.ano || '').trim();
  if (!purchno) return res.status(400).json({ error: 'purchno (purchase invoice number) is required' });
  if (!ano)     return res.status(400).json({ error: 'ano (trade number) is required' });
  if (pcFlagOn(db)) {
    const gateAuction = db.get('SELECT id, price_check_first_passed_at FROM auctions WHERE ano = ? ORDER BY date DESC LIMIT 1', [ano]);
    if (gateAuction && !gateAuction.price_check_first_passed_at) {
      return res.status(412).json({
        error: 'Price check required',
        detail: 'Run Reports → Price Check against the auction (and apply any code fixes) before generating debit notes.',
        auctionId: gateAuction.id, gate: 'price_check',
      });
    }
  }
  const candidates = db.all(`SELECT * FROM purchases WHERE invo = ? ORDER BY date DESC, id DESC`, [purchno]);
  if (!candidates.length) {
    const isSalesInv = db.get(`SELECT id FROM invoices WHERE invo = ? LIMIT 1`, [purchno]);
    if (isSalesInv) {
      return res.status(400).json({ error: `${purchno} is a SALES invoice. Debit notes can only be generated against PURCHASE invoices.` });
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
  const dealerName = purchase.name || '';
  const dupe = db.get(`SELECT id, note_no FROM debit_notes WHERE ano = ? AND name = ? LIMIT 1`, [ano, dealerName]);
  if (dupe) {
    return res.status(409).json({
      error: `Debit note #${dupe.note_no} already exists for ${dealerName} in trade #${ano}`,
      existingId: dupe.id, existingNoteNo: dupe.note_no,
    });
  }
  const baseAmt = Number(purchase.amount || 0);
  if (baseAmt <= 0) return res.status(400).json({ error: 'Purchase amount is zero — cannot compute discount' });
  let discountAmt = req.body.discount != null ? parseFloat(req.body.discount) : NaN;
  if (!Number.isFinite(discountAmt) || discountAmt <= 0) {
    const discountPct  = Number(cfg.discount_pct)  || 0;
    let _g = String(purchase.gstin || '').trim().toUpperCase();
    if (_g.startsWith('GSTIN.')) _g = _g.slice(6);
    else if (_g.startsWith('GSTIN')) _g = _g.slice(5);
    const sellerHasGstin = /^\d{2}/.test(_g);
    const discountDays = sellerHasGstin ? (Number(cfg.dealer_days) || 0) : (Number(cfg.discount_days) || 0);
    if (discountPct <= 0) return res.status(400).json({ error: 'Discount % not configured in settings' });
    discountAmt = discountDays > 0
      ? Math.round((baseAmt / 1000) * discountDays * discountPct)
      : Math.round(baseAmt * discountPct / 100);
  }
  if (discountAmt <= 0) return res.status(400).json({ error: 'Computed discount is zero — check settings or invoice amount' });
  let dealerStateCode = '';
  {
    let g = String(purchase.gstin || '').trim().toUpperCase();
    if (g.startsWith('GSTIN.')) g = g.slice(6);
    else if (g.startsWith('GSTIN')) g = g.slice(5);
    if (/^\d{2}/.test(g)) dealerStateCode = g.slice(0, 2);
  }
  const companyStateCode = String(cfg.tally_state_code
      || (String(cfg.business_state || '').toUpperCase() === 'KERALA' ? '32' : '33'));
  const isInter = !!dealerStateCode && dealerStateCode !== companyStateCode;
  const dnGstRate = Number(cfg.discount_gst) || Number(cfg.gst_service) || 18;
  const flagDiscGst = String(cfg.flag_disc_gst || '').toLowerCase() === 'true' || cfg.flag_disc_gst === true;
  const dealerCarriedGst = Number(purchase.cgst) || Number(purchase.sgst) || Number(purchase.igst);
  let cgst = 0, sgst = 0, igst = 0;
  if (flagDiscGst && dealerCarriedGst) {
    if (isInter) {
      igst = round2(discountAmt * dnGstRate / 100);
    } else {
      const half = round2(discountAmt * (dnGstRate / 2) / 100);
      cgst = half; sgst = half;
    }
  }
  const total = round2(discountAmt + cgst + sgst + igst);
  const trade = db.get('SELECT date FROM auctions WHERE ano = ? LIMIT 1', [ano]);
  const dnDate = trade && trade.date ? addDays(trade.date, 1) : new Date().toISOString().slice(0, 10);
  const rawStart = req.body.startNoteNo != null ? req.body.startNoteNo : req.body.noteNo;
  let noteNo;
  if (rawStart != null && String(rawStart).trim() !== '') {
    const n = parseInt(String(rawStart).trim(), 10);
    if (!Number.isFinite(n) || n < 1) return res.status(400).json({ error: 'Starting Number must be a positive integer' });
    noteNo = String(n);
    const taken = db.get(
      `SELECT id FROM debit_notes WHERE ano = ? AND CAST(note_no AS INTEGER) = ? LIMIT 1`,
      [ano, n]
    );
    if (taken) {
      return res.status(409).json({
        error: `Debit note #${n} is already used in trade #${ano}. Choose a different number.`,
        suggested: (() => {
          const row = db.get('SELECT MAX(CAST(note_no AS INTEGER)) AS mx FROM debit_notes WHERE ano = ?', [ano]);
          const mx = parseInt(row && row.mx, 10);
          return Number.isFinite(mx) && mx > 0 ? mx + 1 : 1;
        })(),
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
    [ano, dnDate, purchase.state || '', dealerName, noteNo, discountAmt, cgst, sgst, igst, total]
  );
  res.json({
    success: true, created: 1, note_no: noteNo, purchno, ano,
    dealer: dealerName, amount: discountAmt, cgst, sgst, igst, total,
  });
});

app.post('/api/debit-notes/generate-bulk', requireInvoiceWrite, (req, res) => {
  const db = getDb();
  const cfg = getSettingsFlat(db);
  let ano = String(req.body.ano || '').trim();
  if (!ano) {
    const purchno = String(req.body.purchno || '').trim();
    if (!purchno) return res.status(400).json({ error: 'Trade number (ano) is required' });
    const p = db.get(`SELECT ano FROM purchases WHERE invo = ? ORDER BY date DESC, id DESC LIMIT 1`, [purchno]);
    if (!p) return res.status(404).json({ error: `Purchase invoice ${purchno} not found` });
    ano = String(p.ano || '').trim();
    if (!ano) return res.status(400).json({ error: 'Purchase row has no trade number' });
  }
  if (pcFlagOn(db)) {
    const ga = db.get('SELECT id, price_check_first_passed_at FROM auctions WHERE ano = ? ORDER BY date DESC LIMIT 1', [ano]);
    if (ga && !ga.price_check_first_passed_at) {
      return res.status(412).json({
        error: 'Price check required',
        detail: 'Run Reports → Price Check against the auction (and apply any code fixes) before generating debit notes.',
        auctionId: ga.id, gate: 'price_check',
      });
    }
  }
  const purchases = db.all(`SELECT * FROM purchases WHERE ano = ? ORDER BY id`, [ano]);
  if (!purchases.length) {
    return res.json({
      success: true, created: 0, skipped: 0, generated: [], skippedDetails: [],
      note: `No purchase invoices in trade #${ano}`,
    });
  }
  const existingKeys = new Set(
    db.all(`SELECT name FROM debit_notes WHERE ano = ?`, [ano]).map(r => r.name || '')
  );
  const trade = db.get('SELECT date FROM auctions WHERE ano = ? LIMIT 1', [ano]);
  const dnDate = trade && trade.date ? addDays(trade.date, 1) : new Date().toISOString().slice(0, 10);
  const discountPct = Number(cfg.discount_pct) || 0;
  const dealerDays  = Number(cfg.dealer_days)  || 0;
  const crDays      = Number(cfg.discount_days) || 0;
  const dnGstRate   = Number(cfg.discount_gst) || Number(cfg.gst_service) || 18;
  const flagDiscGst = String(cfg.flag_disc_gst || '').toLowerCase() === 'true' || cfg.flag_disc_gst === true;
  if (discountPct <= 0) return res.status(400).json({ error: 'Discount % not configured in settings' });
  const eligibleCount = purchases.filter(
    p => !existingKeys.has(p.name || '') && Number(p.amount || 0) > 0
  ).length;
  let nextNoteNo;
  const rawStart = req.body.startNoteNo != null ? req.body.startNoteNo : req.body.startInvoiceNo;
  if (rawStart != null && String(rawStart).trim() !== '') {
    const n = parseInt(String(rawStart).trim(), 10);
    if (!Number.isFinite(n) || n < 1) return res.status(400).json({ error: 'Starting Number must be a positive integer' });
    nextNoteNo = n;
    if (eligibleCount > 0) {
      const upper = nextNoteNo + eligibleCount - 1;
      const collisions = db.all(
        `SELECT CAST(note_no AS INTEGER) AS n FROM debit_notes
          WHERE ano = ? AND CAST(note_no AS INTEGER) BETWEEN ? AND ? ORDER BY n`,
        [ano, nextNoteNo, upper]
      );
      if (collisions.length) {
        const safe = (() => {
          const row = db.get('SELECT MAX(CAST(note_no AS INTEGER)) AS mx FROM debit_notes WHERE ano = ?', [ano]);
          const mx = parseInt(row && row.mx, 10);
          return Number.isFinite(mx) && mx > 0 ? mx + 1 : 1;
        })();
        return res.status(409).json({
          error: `Starting Number ${nextNoteNo} would overlap existing debit note(s) in trade #${ano} ` +
                 `(${collisions.slice(0, 5).map(c => '#' + c.n).join(', ')}` +
                 `${collisions.length > 5 ? `, +${collisions.length - 5} more` : ''}). Try ${safe} or higher.`,
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
    let _dg = String(p.gstin || '').trim().toUpperCase();
    if (_dg.startsWith('GSTIN.')) _dg = _dg.slice(6);
    else if (_dg.startsWith('GSTIN')) _dg = _dg.slice(5);
    const sellerHasGstin = /^\d{2}/.test(_dg);
    const days = sellerHasGstin ? dealerDays : crDays;
    const discountAmt = days > 0
      ? Math.round((baseAmt / 1000) * days * discountPct)
      : Math.round(baseAmt * discountPct / 100);
    if (discountAmt <= 0) {
      skipped.push({ invo: p.invo, ano, buyer: dealerName, reason: 'computed discount is zero' });
      continue;
    }
    let dealerStateCode = '';
    {
      let g = String(p.gstin || '').trim().toUpperCase();
      if (g.startsWith('GSTIN.')) g = g.slice(6);
      else if (g.startsWith('GSTIN')) g = g.slice(5);
      if (/^\d{2}/.test(g)) dealerStateCode = g.slice(0, 2);
    }
    const companyStateCode = String(cfg.tally_state_code
        || (String(cfg.business_state || '').toUpperCase() === 'KERALA' ? '32' : '33'));
    const isInter = !!dealerStateCode && dealerStateCode !== companyStateCode;
    const dealerCarriedGst = Number(p.cgst) || Number(p.sgst) || Number(p.igst);
    let cgst = 0, sgst = 0, igst = 0;
    if (flagDiscGst && dealerCarriedGst) {
      if (isInter) igst = round2(discountAmt * dnGstRate / 100);
      else {
        const half = round2(discountAmt * (dnGstRate / 2) / 100);
        cgst = half; sgst = half;
      }
    }
    const total = round2(discountAmt + cgst + sgst + igst);
    db.run(
      `INSERT INTO debit_notes (ano,date,state,name,note_no,amount,cgst,sgst,igst,total)
       VALUES (?,?,?,?,?,?,?,?,?,?)`,
      [ano, dnDate, p.state || '', dealerName, String(nextNoteNo), discountAmt, cgst, sgst, igst, total]
    );
    generated.push({ note_no: nextNoteNo, purchno: p.invo, dealer: dealerName, total });
    existingKeys.add(dealerName);
    nextNoteNo++;
  }
  res.json({
    success: true,
    created: generated.length, skipped: skipped.length,
    generated, skippedDetails: skipped,
    note: generated.length === 0 && skipped.length === 0
      ? `No eligible purchases in trade #${ano}` : undefined,
  });
});

app.get('/api/debit-notes/eligible-purchases/:auctionId', requireView, (req, res) => {
  const db = getDb();
  const auction = db.get('SELECT ano FROM auctions WHERE id = ?', [req.params.auctionId]);
  if (!auction) return res.status(404).json({ error: 'Auction not found' });
  const ano = auction.ano;
  const rows = db.all(
    `SELECT p.id, p.invo, p.name, p.amount, p.cgst, p.sgst, p.igst, p.total, p.date, p.state
       FROM purchases p
      WHERE p.ano = ? AND p.amount > 0
        AND NOT EXISTS (SELECT 1 FROM debit_notes dn WHERE dn.ano = p.ano AND dn.name = p.name)
      ORDER BY p.id`,
    [ano]
  );
  res.json(rows);
});

app.get('/api/debit-notes/next-note-no', requireView, (req, res) => {
  const db = getDb();
  const ano = String(req.query.ano || '').trim();
  if (!ano) return res.status(400).json({ error: 'ano (trade number) is required for trade-wise numbering' });
  const row = db.get('SELECT MAX(CAST(note_no AS INTEGER)) AS mx FROM debit_notes WHERE ano = ?', [ano]);
  const mx = parseInt(row && row.mx, 10);
  const next = Number.isFinite(mx) && mx > 0 ? mx + 1 : 1;
  res.json({ next, ano });
});

app.post('/api/debit-notes/generate-all', requireInvoiceWrite, (req, res) => {
  res.status(410).json({
    error: 'Cross-trade DN generation is no longer supported. Use POST /api/debit-notes/generate-bulk with { ano } to generate DNs for a specific trade.',
  });
});

app.delete('/api/debit-notes/:id', requireDelete, (req, res) => {
  const db = getDb();
  if (!isAdmin(req) && lotsLockedForDebitNote(db, req.params.id)) {
    return res.status(423).json({ error: 'This debit note is locked because at least one of its lots is locked — only an admin can delete it.' });
  }
  db.run('DELETE FROM debit_notes WHERE id = ?', [req.params.id]);
  res.json({ success: true });
});

app.put('/api/debit-notes/:id', requireInvoiceWrite, (req, res) => {
  const n = req.body;
  const db = getDb();
  if (!isAdmin(req) && lotsLockedForDebitNote(db, req.params.id)) {
    return res.status(423).json({ error: 'This debit note is locked because at least one of its lots is locked — only an admin can edit it.' });
  }
  const fields = ['ano','date','state','name','note_no','amount','cgst','sgst','igst','total'];
  const sets = []; const vals = [];
  for (const f of fields) {
    if (n[f] !== undefined) { sets.push(`${f}=?`); vals.push(n[f]); }
  }
  if (!sets.length) return res.status(400).json({ error: 'No fields to update' });
  vals.push(req.params.id);
  db.run(`UPDATE debit_notes SET ${sets.join(',')} WHERE id=?`, vals);
  res.json({ success: true });
});

// Debit Note PDF — stub that emits a basic PDF. Original was a 400-line
// hand-rolled PDFKit layout; this minimal version still renders a valid
// PDF so the print button works while a full restore is pending.
app.get('/api/debit-notes/:id/pdf', requireView, (req, res) => {
  try {
    const PDFDocument = require('pdfkit');
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const dn = db.get('SELECT * FROM debit_notes WHERE id = ?', [req.params.id]);
    if (!dn) return res.status(404).json({ error: 'Debit note not found' });
    const doc = new PDFDocument({ size: 'A4', margin: 40 });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="DebitNote_${dn.note_no || dn.id}.pdf"`);
    doc.pipe(res);
    const company = (cfg.trade_name || cfg.short_name || 'Company').toString().toUpperCase();
    doc.font('Helvetica-Bold').fontSize(16).text(company, { align: 'center' });
    doc.moveDown(0.3);
    doc.font('Helvetica-Bold').fontSize(13).text('DEBIT NOTE', { align: 'center' });
    doc.moveDown(0.8);
    doc.font('Helvetica').fontSize(11);
    doc.text(`Note No: ${dn.note_no || '-'}`);
    doc.text(`Date: ${fmtDate(dn.date)}`);
    doc.text(`Trade: ${dn.ano}`);
    doc.text(`Dealer: ${dn.name || ''}`);
    doc.moveDown(0.5);
    doc.text(`Amount: ₹${Number(dn.amount || 0).toLocaleString('en-IN', { minimumFractionDigits: 2 })}`);
    if (dn.cgst || dn.sgst) {
      doc.text(`CGST: ₹${Number(dn.cgst || 0).toFixed(2)}`);
      doc.text(`SGST: ₹${Number(dn.sgst || 0).toFixed(2)}`);
    }
    if (dn.igst) doc.text(`IGST: ₹${Number(dn.igst || 0).toFixed(2)}`);
    doc.moveDown(0.5);
    doc.font('Helvetica-Bold').text(`GRAND TOTAL: ₹${Number(dn.total || 0).toLocaleString('en-IN', { minimumFractionDigits: 2 })}`);
    doc.end();
  } catch (e) {
    console.error('[dn-pdf] failed:', e);
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
});

// ──────────────────────────────────────────────────────────────
// JOURNALS (trade-based)
// ──────────────────────────────────────────────────────────────
app.get('/api/journals/sales', requireView, (req, res) => {
  const { auctionId, saleType } = req.query;
  if (!auctionId) return res.status(400).json({ error: 'auctionId required' });
  res.json(getSalesJournal(getDb(), auctionId, saleType));
});

app.get('/api/journals/purchase', requireView, (req, res) => {
  const { auctionId, type } = req.query;
  if (!auctionId) return res.status(400).json({ error: 'auctionId required' });
  res.json(getPurchaseJournal(getDb(), auctionId, type || 'dealer'));
});

app.get('/api/exports/sales-journal', requireExport, async (req, res) => {
  const { auctionId, saleType } = req.query;
  if (!auctionId) return res.status(400).json({ error: 'auctionId required' });
  const { exportSalesJournal } = require('./exports');
  const buffer = await exportSalesJournal(getDb(), auctionId, saleType);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="SalesJournal.xlsx"');
  res.send(Buffer.from(buffer));
});

app.get('/api/exports/purchase-journal', requireExport, async (req, res) => {
  const { auctionId, type } = req.query;
  if (!auctionId) return res.status(400).json({ error: 'auctionId required' });
  const baseName = type === 'agri' ? 'AgriBillJournal' : 'PurchaseJournal';
  const { exportPurchaseJournal } = require('./exports');
  const buffer = await exportPurchaseJournal(getDb(), auctionId, type || 'dealer');
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${baseName}.xlsx"`);
  res.send(Buffer.from(buffer));
});

// Preview (dry-run) endpoint shared by sales/purchase/agri
app.post('/api/invoices/preview/:auctionId', requireView, (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const { saleType, buyerCode, type } = req.body;
  const uncalc = db.all(`SELECT * FROM lots WHERE auction_id = ? AND amount > 0 AND (puramt IS NULL OR puramt = 0)`, [req.params.auctionId]);
  for (const lot of uncalc) {
    const c = calculateLot(lot, cfg);
    db.run(`UPDATE lots SET pqty=?,prate=?,puramt=?,com=?,sertax=?,cgst=?,sgst=?,igst=?,advance=?,balance=?,bilamt=?,refund=?,refud=?,isp_pqty=?,isp_prate=?,isp_puramt=?,asp_pqty=?,asp_prate=?,asp_puramt=? WHERE id=?`,
      [c.pqty,c.prate,c.puramt,c.com,c.sertax,c.cgst,c.sgst,c.igst,c.advance,c.balance,c.bilamt,c.refund||0,c.refud||0,c.isp_pqty||0,c.isp_prate||0,c.isp_puramt||0,c.asp_pqty||0,c.asp_prate||0,c.asp_puramt||0,lot.id]);
  }
  let invoice;
  if (type === 'purchase') {
    invoice = buildPurchaseInvoice(db, req.params.auctionId, buyerCode, cfg);
  } else if (type === 'agri') {
    invoice = buildAgriBill(db, req.params.auctionId, buyerCode, cfg);
    if (invoice && invoice.error) return res.status(404).json({ error: invoice.error });
  } else {
    invoice = buildSalesInvoice(db, req.params.auctionId, buyerCode, saleType, cfg);
  }
  if (!invoice) return res.status(404).json({ error: 'No data found' });
  res.json({ preview: true, invoice });
});

// ──────────────────────────────────────────────────────────────
// PAYMENTS
// ──────────────────────────────────────────────────────────────
app.get('/api/payments/:auctionId', requireView, (req, res) => {
  const db = getDb();
  const cfg = getSettingsFlat(db);
  const summary = getPaymentSummary(db, req.params.auctionId, req.query.state, cfg);
  res.json(summary);
});

app.post('/api/payments/:auctionId/delete-sellers', requireDelete, (req, res) => {
  try {
    const db = getDb();
    const auctionId = req.params.auctionId;
    const names = Array.isArray(req.body.sellerNames) ? req.body.sellerNames : [];
    if (!names.length) return res.status(400).json({ error: 'sellerNames array is required' });
    const auction = db.get('SELECT ano FROM auctions WHERE id = ?', [auctionId]);
    if (!auction) return res.status(404).json({ error: 'Auction not found' });
    const CHUNK = 500;
    let lotsDeleted = 0, dnsDeleted = 0;
    for (let i = 0; i < names.length; i += CHUNK) {
      const batch = names.slice(i, i + CHUNK).map(s => String(s).trim()).filter(Boolean);
      if (!batch.length) continue;
      const placeholders = batch.map(() => '?').join(',');
      const upperBatch = batch.map(n => n.toUpperCase());
      const lotsBefore = db.get(
        `SELECT COUNT(*) AS c FROM lots WHERE auction_id = ? AND UPPER(COALESCE(name,'')) IN (${placeholders})`,
        [auctionId, ...upperBatch]
      ).c;
      db.run(
        `DELETE FROM lots WHERE auction_id = ? AND UPPER(COALESCE(name,'')) IN (${placeholders})`,
        [auctionId, ...upperBatch]
      );
      lotsDeleted += lotsBefore;
      const dnsBefore = db.get(
        `SELECT COUNT(*) AS c FROM debit_notes WHERE ano = ? AND UPPER(COALESCE(name,'')) IN (${placeholders})`,
        [auction.ano, ...upperBatch]
      ).c;
      db.run(
        `DELETE FROM debit_notes WHERE ano = ? AND UPPER(COALESCE(name,'')) IN (${placeholders})`,
        [auction.ano, ...upperBatch]
      );
      dnsDeleted += dnsBefore;
    }
    res.json({ success: true, sellers: names.length, lotsDeleted, dnsDeleted });
  } catch (e) {
    res.status(500).json({ error: 'Delete failed: ' + (e.message || e) });
  }
});

app.get('/api/payments/bank/:auctionId', requireView, (req, res) => {
  const cfg = getSettingsFlat(getDb());
  const data = getBankPaymentData(getDb(), req.params.auctionId, cfg);
  res.json(data);
});

function _renderPaymentStatement(doc, db, auctionId, sellerName, cfg) {
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]) || { ano:'', date:'' };
  const lots = db.all(
    `SELECT lot_no, qty, prate AS rate, amount, puramt, refund, balance, cgst, sgst, igst
       FROM lots
      WHERE auction_id = ? AND TRIM(LOWER(COALESCE(name,''))) = TRIM(LOWER(?)) AND amount > 0
      ORDER BY CAST(lot_no AS INTEGER), lot_no`,
    [auctionId, sellerName]
  ) || [];
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
  doc.text(`Seller: ${sellerName}`, PAGE_L, y);
  doc.text(`Auction: ${auction.ano}`, PAGE_L + 280, y);
  y += 14;
  doc.text(`Phone: ${trader && trader.tel ? trader.tel : '-'}`, PAGE_L, y);
  doc.text(`Date: ${fmtDate(auction.date)}`, PAGE_L + 280, y);
  y += 18;
  let tQty=0,tAmt=0,tDisc=0,tTax=0,tPay=0;
  for (const l of lots) {
    const tax = (Number(l.cgst)||0)+(Number(l.sgst)||0)+(Number(l.igst)||0);
    tQty+=Number(l.qty)||0; tAmt+=Number(l.amount)||0; tDisc+=Number(l.refund)||0; tTax+=tax; tPay+=Number(l.balance)||0;
    if (y > 770) { doc.addPage(); y = 40; }
    doc.fontSize(9).text(`Lot ${l.lot_no}  ${fmtQty(l.qty)}kg  ₹${fmtAmt(l.balance)}`, PAGE_L, y);
    y += 12;
  }
  doc.moveDown(0.5);
  doc.font('Helvetica-Bold').fontSize(11).text(`TOTAL PAYABLE: ₹${fmtAmt(tPay)}`, PAGE_L, doc.y);
  return tPay;
}

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

// ──────────────────────────────────────────────────────────────
// TDS RETURNS
// ──────────────────────────────────────────────────────────────
app.get('/api/tds-return', requireView, (req, res) => {
  const { from, to, orderBy } = req.query;
  if (!from || !to) return res.status(400).json({ error: 'from and to dates required' });
  const data = getTDSReturnData(getDb(), from, to, orderBy || 'invoice');
  res.json(data);
});

// ──────────────────────────────────────────────────────────────
// EXPORTS
// ──────────────────────────────────────────────────────────────
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
      buffer = await exportDef.fn(db, auctionId, cfg, req.query.state);
    } else {
      buffer = await exportDef.fn(db, auctionId, req.query.state);
    }
    const ext  = exportDef.ext  || 'xlsx';
    const mime = exportDef.mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    res.setHeader('Content-Type', mime);
    res.setHeader('Content-Disposition', `attachment; filename="${exportDef.name}_${auctionId}.${ext}"`);
    res.send(Buffer.from(buffer));
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

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

// ──────────────────────────────────────────────────────────────
// LORRY REPORTS
// ──────────────────────────────────────────────────────────────
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

// ──────────────────────────────────────────────────────────────
// DBF EXPORTS
// ──────────────────────────────────────────────────────────────
app.get('/api/dbf-exports/list', requireExport, (req, res) => {
  const list = {};
  for (const [key, def] of Object.entries(DBF_EXPORTS)) {
    list[key] = {
      label: def.label, name: def.name,
      needsAuction: !!def.needsAuction,
      needsDateRange: !!def.needsDateRange,
    };
  }
  res.json(list);
});

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

// ──────────────────────────────────────────────────────────────
// TALLY XML EXPORTS
// ──────────────────────────────────────────────────────────────
const TALLY_EXPORTS = {
  ledger_sales:        { label: 'Sales Party Ledgers',                              name: 'SalesPartyLedgers',  builder: buildSalesPartyLedgerRows, generator: generLedgerXML, isLedger: true, company: 'isp' },
  ledger_rd_purchase:  { label: 'RD Purchase Party Ledgers',                        name: 'RDPartyLedgers',     builder: buildRDPartyLedgerRows,    generator: generLedgerXML, isLedger: true, company: 'isp' },
  ledger_urd_purchase: { label: 'URD Purchase Party Ledgers (Agriculturist)',       name: 'URDPartyLedgers',    builder: buildURDPartyLedgerRows,   generator: generLedgerXML, isLedger: true, company: 'isp' },
  ledger:              { label: 'All Ledger Masters (parties + tax + sales + purchase)', name: 'AllLedgers',  builder: buildLedgerRows,           generator: generLedgerXML, isLedger: true, company: 'isp' },
  sales_isp:           { label: 'Sales Vouchers',                                   name: 'Sales',              builder: buildSalesIspRows,         generator: generSalesIspXML,     company: 'isp' },
  sales_asp:           { label: 'Sales Vouchers (Kerala / Intra-Company)',          name: 'SalesIntra',         builder: buildSalesAspRows,         generator: generSalesAspXML,     company: 'isp' },
  sales:               { label: 'Sales Vouchers (legacy alias)',                    name: 'Sales',              builder: buildSalesIspRows,         generator: generSalesIspXML,     company: 'isp' },
  isp_purchase:        { label: 'Intra-Company Purchase Vouchers',                  name: 'IntraPurchase',      builder: buildSalesAspRows,         generator: generIspPurchaseXML,  company: 'isp' },
  rd_purchase:         { label: 'RD Purchase Vouchers',                             name: 'RDPurchase',         builder: buildRDPurchaseRows,       generator: generRDPurchaseXML,   company: 'isp' },
  urd_purchase:        { label: 'URD Purchase Vouchers (Agriculturist)',            name: 'URDPurchase',        builder: buildURDPurchaseRows,      generator: generURDPurchaseXML,  company: 'isp' },
  debit_note:          { label: 'Debit Notes (Discount)',                           name: 'DebitNote',          builder: buildDebitNoteRows,        generator: generDebitNoteXML,    company: 'isp' },
};

function resolveTallyCompanyName(cfg, target) {
  const isp = (cfg.tally_company_name || '').trim();
  const asp = (cfg.tally_asp_company_name || '').trim();
  if (target === 'asp') {
    if (!asp) console.warn('[tally] tally_asp_company_name is empty — falling back to ISP company name.');
    return asp || isp;
  }
  return isp;
}

const PARTY_LEDGER_BUILDERS = {
  sales:        { builder: buildSalesPartyLedgerRows, company: 'isp' },
  rd_purchase:  { builder: buildRDPartyLedgerRows,    company: 'isp' },
  urd_purchase: { builder: buildURDPartyLedgerRows,   company: 'isp' },
};

app.get('/api/tally/list', requireExport, (req, res) => {
  const list = {};
  for (const [key, def] of Object.entries(TALLY_EXPORTS)) {
    list[key] = { label: def.label, name: def.name };
  }
  res.json(list);
});

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
      const byKind = rows.reduce((acc, r) => {
        const k = r.kind || 'other';
        acc[k] = (acc[k] || 0) + 1;
        return acc;
      }, {});
      return res.json({
        type, auctionId,
        ledgerCount: rows.length,
        byKind, targetCompany,
        sample: rows.slice(0, 6).map(r => ({ kind: r.kind, name: r.name, parent: r.parent, gstin: r.gstin || '' })),
      });
    }
    const totalLots = rows.reduce((s, r) => s + (Array.isArray(r.lots) ? r.lots.length : 0), 0);
    const distinctParties = new Set();
    for (const r of rows) {
      const n = String(r.partyName || r.name || '').trim();
      if (n) distinctParties.add(n.toUpperCase());
    }
    res.json({
      type, auctionId,
      voucherCount: rows.length,
      lotCount: totalLots,
      partyCount: distinctParties.size,
      hasLots: rows.some(r => Array.isArray(r.lots) && r.lots.length > 0),
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

// E-way bill distance management
function getDispatchPin(db) {
  const cfg = require('./company-config').getSettingsFlat(db);
  return String(cfg.tally_dispatch_pin || cfg.s_pin || cfg.kl_pin || cfg.tn_pin || '').trim();
}
function normalizeRouteKey(fromPin, toPin) {
  const a = String(fromPin || '').trim();
  const b = String(toPin || '').trim();
  return a < b ? [a, b] : [b, a];
}

app.get('/api/invoices/distances/:auctionId', requireView, (req, res) => {
  try {
    const db = getDb();
    const dispatchPin = getDispatchPin(db);
    const rows = db.all(
      `SELECT i.id, i.ano, i.invo, i.buyer, i.buyer1, i.gstin, i.state,
              b.pin AS buyer_pin, b.pla AS buyer_pla, i.distance_km
       FROM invoices i LEFT JOIN buyers b ON b.buyer = i.buyer
       WHERE i.auction_id = ? ORDER BY CAST(i.invo AS INTEGER), i.id`,
      [req.params.auctionId]
    );
    const routes = {};
    try {
      const allRoutes = db.all(
        `SELECT from_pin, to_pin, km FROM route_distances WHERE from_pin = ? OR to_pin = ?`,
        [dispatchPin, dispatchPin]
      );
      for (const r of allRoutes) {
        const other = r.from_pin === dispatchPin ? r.to_pin : r.from_pin;
        routes[other] = r.km;
      }
    } catch (e) {}
    const enriched = rows.map(r => {
      let km = null, source = 'none';
      if (r.distance_km != null) { km = r.distance_km; source = 'manual'; }
      else if (r.buyer_pin && routes[String(r.buyer_pin).trim()] != null) {
        km = routes[String(r.buyer_pin).trim()]; source = 'route';
      }
      return { ...r, resolved_km: km, distance_source: source };
    });
    res.json({ count: enriched.length, dispatchPin, invoices: enriched });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/route-distances', requireExport, (req, res) => {
  const { from_pin, to_pin, km } = req.body || {};
  if (!from_pin || !to_pin) return res.status(400).json({ error: 'from_pin and to_pin required' });
  if (!/^\d{6}$/.test(String(from_pin).trim()) || !/^\d{6}$/.test(String(to_pin).trim())) {
    return res.status(400).json({ error: 'PINs must be 6-digit strings' });
  }
  const [k1, k2] = normalizeRouteKey(from_pin, to_pin);
  if (km === '' || km == null) {
    try {
      const r = getDb().run('DELETE FROM route_distances WHERE from_pin = ? AND to_pin = ?', [k1, k2]);
      return res.json({ ok: true, deleted: r.changes > 0, from_pin: k1, to_pin: k2 });
    } catch (e) { return res.status(500).json({ error: e.message }); }
  }
  const v = Math.round(Number(km));
  if (!isFinite(v) || v < 0 || v > 5000) return res.status(400).json({ error: 'km must be between 0 and 5000' });
  try {
    const db = getDb();
    db.run(
      `INSERT INTO route_distances (from_pin, to_pin, km, updated_at)
       VALUES (?, ?, ?, datetime('now','localtime'))
       ON CONFLICT(from_pin, to_pin) DO UPDATE SET km = excluded.km, updated_at = excluded.updated_at`,
      [k1, k2, v]
    );
    const dispatchPin = getDispatchPin(db);
    const otherPin = k1 === dispatchPin ? k2 : (k2 === dispatchPin ? k1 : null);
    let clearedOverrides = 0;
    if (otherPin) {
      const r = db.run(
        `UPDATE invoices SET distance_km = NULL
         WHERE id IN (
           SELECT i.id FROM invoices i LEFT JOIN buyers b ON b.buyer = i.buyer
           WHERE COALESCE(NULLIF(TRIM(b.cpin), ''), TRIM(b.pin)) = ? AND i.distance_km IS NOT NULL
         )`,
        [otherPin]
      );
      clearedOverrides = r.changes || 0;
    }
    let appliedCount = 0;
    if (otherPin) {
      const r = db.get(
        `SELECT COUNT(*) AS n FROM invoices i LEFT JOIN buyers b ON b.buyer = i.buyer
         WHERE COALESCE(NULLIF(TRIM(b.cpin), ''), TRIM(b.pin)) = ?`,
        [otherPin]
      );
      appliedCount = r ? r.n : 0;
    }
    res.json({ ok: true, from_pin: k1, to_pin: k2, km: v, appliedCount, clearedOverrides });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/distance-overrides', requireExport, (req, res) => {
  try {
    const r = getDb().run(`UPDATE invoices SET distance_km = NULL WHERE distance_km IS NOT NULL`);
    res.json({ ok: true, cleared: r.changes });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

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

const VOUCHER_PARTY_KEY = {
  sales_isp:    (r) => r.partyName || '',
  sales_asp:    (r) => r.buyerName || r.buyer || '',
  isp_purchase: (r) => r.buyerName || r.buyer || '',
  rd_purchase:  (r) => r.name || '',
  urd_purchase: (r) => r.name || '',
  debit_note:   (r) => r.partyName || r.name || '',
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

// Crop Receipt PDF
app.get('/api/receipt/:lotId', requireView, async (req, res) => {
  const db = getDb(); const cfg = getSettingsFlat(db);
  const lot = db.get('SELECT l.*, a.ano FROM lots l JOIN auctions a ON a.id=l.auction_id WHERE l.id=?', [req.params.lotId]);
  if (!lot) return res.status(404).json({ error: 'Lot not found' });
  const pdf = await generateCropReceiptPDF(lot, cfg);
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `inline; filename="Receipt_${lot.lot_no}.pdf"`);
  res.send(pdf);
});

// ──────────────────────────────────────────────────────────────
// SUMMARY STATS
// ──────────────────────────────────────────────────────────────
app.get('/api/stats', requireView, (req, res) => {
  res.set('Cache-Control', 'no-store');
  const db = getDb();
  const counts = {
    traders:    (db.get('SELECT COUNT(*) as c FROM traders') || {}).c || 0,
    buyers:     (db.get('SELECT COUNT(*) as c FROM buyers') || {}).c || 0,
    auctions:   (db.get('SELECT COUNT(*) as c FROM auctions') || {}).c || 0,
    lots:       (db.get('SELECT COUNT(*) as c FROM lots') || {}).c || 0,
    invoices:   (db.get('SELECT COUNT(*) as c FROM invoices') || {}).c || 0,
    purchases:  (db.get('SELECT COUNT(*) as c FROM purchases') || {}).c || 0,
    bills:      (db.get('SELECT COUNT(*) as c FROM bills') || {}).c || 0,
    debit_notes:(db.get('SELECT COUNT(*) as c FROM debit_notes') || {}).c || 0,
  };
  const allAuctions = db.all(`SELECT id, ano, date, crop_type FROM auctions ORDER BY id DESC LIMIT 50`);
  const cumRow = db.get(
    `SELECT COALESCE(SUM(qty),0) as qty, COALESCE(SUM(amount),0) as amount, COUNT(*) as lots,
            COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(code,''))) NOT IN ('','WD') THEN qty ELSE 0 END),0) as sold_qty,
            COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(code,'')))  =  'WD' THEN qty ELSE 0 END),0) as wd_qty,
            COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(code,''))) NOT IN ('','WD') AND amount > 0 THEN amount ELSE 0 END),0) as sold_amount,
            (SELECT COALESCE(MIN(price),0) FROM lots WHERE price > 0 AND amount > 0 AND UPPER(TRIM(COALESCE(code,''))) NOT IN ('','WD')) as min_price,
            (SELECT COALESCE(MAX(price),0) FROM lots WHERE price > 0 AND amount > 0 AND UPPER(TRIM(COALESCE(code,''))) NOT IN ('','WD')) as max_price
     FROM lots`
  ) || {};
  const cumSoldQty = Number(cumRow.sold_qty) || 0;
  const cumSoldAmt = Number(cumRow.sold_amount) || 0;
  const cumulative = {
    qty: cumRow.qty || 0,
    amount: cumRow.amount || 0,
    lots: cumRow.lots || 0,
    auctions: counts.auctions,
    sold_qty: cumSoldQty,
    wd_qty: cumRow.wd_qty || 0,
    min_price: Number(cumRow.min_price) || 0,
    max_price: Number(cumRow.max_price) || 0,
    avg_price: cumSoldQty > 0 ? round2(cumSoldAmt / cumSoldQty) : 0,
  };
  const perTradeBreakdown = db.all(
    `SELECT a.id, a.ano, a.date, a.crop_type,
            COUNT(l.id) as lots, COALESCE(SUM(l.qty),0) as qty,
            COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(l.code,''))) NOT IN ('','WD') THEN l.qty ELSE 0 END),0) as sold_qty,
            COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(l.code,'')))  =  'WD' THEN l.qty ELSE 0 END),0) as wd_qty,
            COALESCE(SUM(l.amount),0) as amount,
            COALESCE(SUM(CASE WHEN UPPER(TRIM(COALESCE(l.code,''))) NOT IN ('','WD') AND l.amount > 0 THEN l.amount ELSE 0 END),0) as sold_amount,
            COALESCE(MIN(CASE WHEN l.price > 0 AND l.amount > 0 AND UPPER(TRIM(COALESCE(l.code,''))) NOT IN ('','WD') THEN l.price END),0) as min_price,
            COALESCE(MAX(CASE WHEN l.price > 0 AND l.amount > 0 AND UPPER(TRIM(COALESCE(l.code,''))) NOT IN ('','WD') THEN l.price END),0) as max_price,
            COALESCE(SUM(CASE WHEN l.amount > 0 THEN 1 ELSE 0 END),0) as priced,
            COALESCE(SUM(CASE WHEN l.invo IS NOT NULL AND l.invo != '' THEN 1 ELSE 0 END),0) as invoiced
     FROM auctions a LEFT JOIN lots l ON l.auction_id = a.id
     GROUP BY a.id, a.ano, a.date, a.crop_type
     ORDER BY a.date DESC, a.id DESC LIMIT 50`
  ).map(r => {
    const soldQty = Number(r.sold_qty) || 0;
    const soldAmt = Number(r.sold_amount) || 0;
    return { ...r, avg_price: soldQty > 0 ? round2(soldAmt / soldQty) : 0 };
  });

  let currentAuction = null;
  const rawAuctionId = req.query.auction_id;
  const isAllMode = (rawAuctionId === 'all' || rawAuctionId === '' || rawAuctionId === undefined);
  if (!isAllMode) {
    const requestedId = parseInt(rawAuctionId);
    if (requestedId) currentAuction = db.get('SELECT * FROM auctions WHERE id = ?', [requestedId]);
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
  const topSellers = db.all(
    `SELECT l.name as name, COUNT(*) as lots, COALESCE(SUM(l.qty),0) as qty, COALESCE(SUM(l.amount),0) as amount
     FROM lots l JOIN auctions a ON a.id = l.auction_id
     WHERE a.date >= date('now','-7 days') AND l.name IS NOT NULL AND l.name != ''
     GROUP BY l.name ORDER BY amount DESC LIMIT 5`
  );
  const recentInvoices = db.all(
    `SELECT i.id, i.sale, i.invo, i.buyer, i.buyer1, i.tot, i.date, i.place
     FROM invoices i ORDER BY i.id DESC LIMIT 5`
  );
  const todayQty = auctionStats ? auctionStats.totalQty : 0;
  const todayAmt = auctionStats ? auctionStats.totalAmt : 0;
  const monthTot = (db.get(`SELECT COALESCE(SUM(tot),0) as s FROM invoices WHERE date >= date('now','start of month')`) || {}).s || 0;
  const lastMonthTot = (db.get(
    `SELECT COALESCE(SUM(tot),0) as s FROM invoices
     WHERE date >= date('now','start of month','-1 month') AND date <  date('now','start of month')`
  ) || {}).s || 0;
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
       WHERE amount > 0 AND buyer IS NOT NULL AND buyer != '' AND (invo IS NULL OR invo = '')`
    ) || {}).c || 0;
  }
  const cfgBranches = [];
  for (let i = 1; i <= 9; i++) {
    const r = db.get('SELECT value FROM company_settings WHERE key = ?', [`br${i}`]);
    const v = r && r.value ? String(r.value).trim() : '';
    if (v) cfgBranches.push(v);
  }
  const aggSql = currentAuction
    ? `SELECT COALESCE(TRIM(branch), '') AS branch, COUNT(*) AS lots, COALESCE(SUM(qty), 0) AS qty, COALESCE(SUM(amount), 0) AS amount
       FROM lots WHERE auction_id = ? GROUP BY UPPER(COALESCE(TRIM(branch), ''))`
    : `SELECT COALESCE(TRIM(branch), '') AS branch, COUNT(*) AS lots, COALESCE(SUM(qty), 0) AS qty, COALESCE(SUM(amount), 0) AS amount
       FROM lots GROUP BY UPPER(COALESCE(TRIM(branch), ''))`;
  const rawAgg = currentAuction ? db.all(aggSql, [currentAuction.id]) : db.all(aggSql);
  const aggIdx = {};
  for (const a of rawAgg) {
    const k = String(a.branch || '').toUpperCase();
    aggIdx[k] = a;
  }
  const branchTotals = cfgBranches.map(name => {
    const hit = aggIdx[name.toUpperCase()];
    return {
      branch: name,
      lots:   Number((hit && hit.lots)   || 0),
      qty:    Number((hit && hit.qty)    || 0),
      amount: Number((hit && hit.amount) || 0),
      configured: true,
    };
  });
  res.json({
    counts, cumulative, perTradeBreakdown, branchTotals,
    currentAuction: auctionStats, allAuctions, topSellers, recentInvoices,
    kpi: {
      todayQty, todayAmt,
      activeLots: auctionStats ? auctionStats.totalLots : 0,
      pendingInvoices,
      monthRevenue: monthTot,
      lastMonthRevenue: lastMonthTot,
    }
  });
});

app.get('/api/stats/revenue-trend', requireView, (req, res) => {
  res.set('Cache-Control', 'no-store');
  const raw = parseInt(req.query.days, 10);
  const days = Math.max(1, Math.min(90, Number.isFinite(raw) ? raw : 7));
  const db = getDb();
  const rows = db.all(
    `SELECT date(date) as day, COALESCE(SUM(tot), 0) as total, COUNT(*) as count
       FROM invoices
      WHERE date IS NOT NULL AND date != ''
        AND date(date) >= date('now', '-' || ? || ' days')
        AND date(date) <= date('now')
      GROUP BY date(date) ORDER BY day ASC`,
    [days - 1]
  );
  const byDay = new Map(rows.map(r => [r.day, r]));
  const series = [];
  const today = new Date();
  for (let i = days - 1; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const iso = d.toISOString().slice(0, 10);
    const hit = byDay.get(iso);
    series.push({
      date: iso,
      total: hit ? Number(hit.total) || 0 : 0,
      count: hit ? Number(hit.count) || 0 : 0,
    });
  }
  res.json({ days, series });
});

// Trade summary report
app.get('/api/reports/trade-summary/:auctionId', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id || req.params.auctionId, 10);
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]);
  if (!auction) return res.status(404).json({ error: 'Trade not found' });
  const branchFilter = String(req.query.branch || '').trim();
  const bWhere = branchFilter ? ' AND branch = ?' : '';
  const bParams = branchFilter ? [branchFilter] : [];
  const branchWise = db.all(
    `SELECT branch, COUNT(*) AS lot_count, SUM(bags) AS total_bags, SUM(qty) AS total_qty,
            COUNT(DISTINCT trader_id) AS seller_count,
            SUM(CASE WHEN amount > 0 THEN 1 ELSE 0 END) AS sold_lots,
            SUM(CASE WHEN amount > 0 THEN qty ELSE 0 END) AS sold_qty,
            SUM(CASE WHEN COALESCE(amount,0) <= 0 THEN 1 ELSE 0 END) AS withdrawn_lots,
            SUM(CASE WHEN COALESCE(amount,0) <= 0 THEN qty ELSE 0 END) AS withdrawn_qty,
            MAX(CASE WHEN amount > 0 THEN price END) AS max_price,
            MIN(CASE WHEN amount > 0 THEN price END) AS min_price,
            CASE WHEN SUM(CASE WHEN amount > 0 THEN qty ELSE 0 END) > 0
                 THEN SUM(CASE WHEN amount > 0 THEN amount ELSE 0 END) * 1.0 / SUM(CASE WHEN amount > 0 THEN qty ELSE 0 END)
                 ELSE 0 END AS avg_price
       FROM lots WHERE auction_id = ? GROUP BY branch ORDER BY total_qty DESC`,
    [auctionId]
  );
  const sellerWise = db.all(
    `SELECT COALESCE(t.name, l.name, 'Unknown') AS seller_name, l.trader_id, l.branch,
            COUNT(*) AS lot_count, SUM(l.bags) AS total_bags, SUM(l.qty) AS total_qty
       FROM lots l LEFT JOIN traders t ON t.id = l.trader_id
      WHERE l.auction_id = ? GROUP BY COALESCE(l.trader_id, l.name) ORDER BY total_qty DESC`,
    [auctionId]
  );
  const totals = db.get(
    `SELECT COUNT(*) AS lot_count, SUM(bags) AS total_bags, SUM(qty) AS total_qty,
            COUNT(DISTINCT trader_id) AS seller_count, COUNT(DISTINCT branch) AS branch_count
       FROM lots WHERE auction_id = ?`,
    [auctionId]
  ) || { lot_count: 0, total_bags: 0, total_qty: 0, seller_count: 0, branch_count: 0 };
  res.json({ auction, totals, branchWise, sellerWise });
});

app.get('/api/reports/branch-comparison', requireViewOrLotEntry, (req, res) => {
  const db = getDb();
  const data = db.all(
    `SELECT l.branch, a.id AS auction_id, a.ano, a.date, a.crop_type,
            COUNT(*) AS lot_count, SUM(l.bags) AS total_bags, SUM(l.qty) AS total_qty
       FROM lots l JOIN auctions a ON a.id = l.auction_id
      GROUP BY l.branch, l.auction_id
      ORDER BY a.date DESC, l.branch ASC`
  );
  const overall = db.all(
    `SELECT branch, COUNT(*) AS lot_count, SUM(bags) AS total_bags, SUM(qty) AS total_qty,
            COUNT(DISTINCT auction_id) AS trade_count, COUNT(DISTINCT trader_id) AS seller_count
       FROM lots GROUP BY branch ORDER BY total_qty DESC`
  );
  res.json({ data, overall });
});

// ──────────────────────────────────────────────────────────────
// IMPORT OLD DATA (Feature #9)
// ──────────────────────────────────────────────────────────────
const IMPORT_MODULES = {
  sales_invoice: {
    label: 'Sales Invoices', table: 'invoices', keyCols: ['invo', 'sale'],
    autoFillAuctionId: true,
    fields: ['auction_id','ano','date','state','sale','invo','buyer','buyer1','gstin','place',
             'bag','qty','amount','gunny','pava_hc','ins','cgst','sgst','igst','tcs','rund','tot'],
    aliases: {
      ano: ['ano','auction_no','trade'],
      date: ['date','invoice_date','inv_date'],
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
  },
  purchase: {
    label: 'Purchase Invoices', table: 'purchases', keyCols: ['invo'],
    autoFillAuctionId: true,
    fields: ['auction_id','ano','date','state','br','name','add_line','place','gstin','invo',
             'qty','amount','cgst','sgst','igst','rund','total','tds'],
    aliases: {
      invo: ['invo','invoice','invoice_no'],
      name: ['name','seller','dealer'],
      gstin: ['gstin','gst','gst_no','cr','registration'],
      place: ['place','city','pla'],
      add_line: ['add_line','address','add','add1','address1'],
      br: ['br','branch'],
      qty: ['qty','kilos','weight','kgs'],
      amount: ['amount','cardamom','value'],
      total: ['total','grand_total','invoice_amount'],
      rund: ['rund','round','round_off'],
      tds: ['tds','tds_amount'],
    },
  },
  bills: {
    label: 'Bills of Supply', table: 'bills', keyCols: ['bil'],
    autoFillAuctionId: true,
    fields: ['auction_id','ano','date','state','br','crpt','bil','name','add_line','pla',
             'pstate','st_code','crr','pan','qty','cost','igst','net'],
    aliases: {
      bil: ['bil','bill','bill_no'],
      name: ['name','seller','planter'],
      qty: ['qty','kilos','weight','kgs'],
      cost: ['cost','amount','cardamom'],
      net: ['net','nett','net_amount'],
    },
  },
  debit_notes: {
    label: 'Debit Notes', table: 'debit_notes', keyCols: ['note_no','ano'],
    autoFillAuctionId: true,
    fields: ['auction_id','ano','date','state','name','note_no','amount','cgst','sgst','igst','total'],
    aliases: {
      note_no: ['note_no','note','dn_no'],
      name: ['name','dealer','buyer'],
    },
  },
  sellers: {
    label: 'Sellers', table: 'traders', keyCols: ['name','cr'],
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
    label: 'Buyers', table: 'buyers', keyCols: ['buyer','code'],
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

function _importMapHeaders(headers, moduleDef) {
  const norm = s => String(s || '').trim().toLowerCase().replace(/[\s\-/]+/g, '_');
  const out = {};
  for (const field of moduleDef.fields) {
    const aliases = (moduleDef.aliases && moduleDef.aliases[field]) || [field];
    for (const h of headers) {
      if (aliases.includes(norm(h))) { out[field] = h; break; }
    }
  }
  return out;
}

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
      module: moduleKey, label: def.label, total: rows.length,
      headers, fields: def.fields, keyCols: def.keyCols,
      autoFillAuctionId: !!def.autoFillAuctionId,
      detectedMapping: mapping,
      missingFields: def.fields.filter(f => !mapping[f] && def.keyCols.includes(f)),
      preview: rows.slice(0, 50),
    });
  } catch (e) {
    res.status(400).json({ error: e.message });
  } finally {
    if (req.file) fs.unlink(req.file.path, () => {});
  }
});

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
        samples: { new: [], invalid: [], dupChanges: [] },
      });
    }
    const headers = Object.keys(rows[0]);
    const autoDetected = _importMapHeaders(headers, def);
    const mapping = Object.assign({}, autoDetected, userMapping);
    for (const k of Object.keys(userMapping || {})) {
      const v = userMapping[k];
      if (v === '' || v === null) delete mapping[k];
    }
    const fieldSources = def.fields.map(f => [f, mapping[f] || null]);
    const auctionIdSlot = def.fields.indexOf('auction_id');
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
    const sampleNew = [], sampleInvalid = [], sampleDupChanges = [], sampleDupIdentical = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const values = {};
      for (const [f, src] of fieldSources) values[f] = src ? r[src] : '';
      const reasons = [];
      let anoResolutionFailed = false;
      if (def.autoFillAuctionId && auctionIdSlot >= 0) {
        const anoSrc = mapping.ano;
        const anoVal = anoSrc ? r[anoSrc] : '';
        const aid = resolveAuctionId(anoVal);
        if (aid == null) {
          anoResolutionFailed = true;
          reasons.push('No trade found for ano="' + String(anoVal || '').trim() + '" — create the auction first or fix the mapping.');
        } else { values.auction_id = aid; }
      }
      const missingKeys = [];
      for (const k of def.keyCols) {
        const src = mapping[k];
        const v = src ? r[src] : null;
        if (v == null || String(v).trim() === '') missingKeys.push(k);
      }
      const requiredMissing = missingKeys.length > 0;
      if (requiredMissing) reasons.push('Missing required value(s): ' + missingKeys.join(', '));
      let existing = null;
      let diff = null;
      if (!requiredMissing) {
        const keyVals = def.keyCols.map(k => r[mapping[k]]);
        const whereSql = def.keyCols.map(k => `${k} = ?`).join(' AND ');
        existing = db.get(`SELECT * FROM ${def.table} WHERE ${whereSql} LIMIT 1`, keyVals);
        if (existing) {
          diff = {};
          for (const f of def.fields) {
            const newVal = values[f];
            const oldVal = existing[f];
            const a = oldVal == null ? '' : String(oldVal);
            const b = newVal == null ? '' : String(newVal);
            if (a !== b) diff[f] = { old: oldVal == null ? '' : oldVal, new: newVal == null ? '' : newVal };
          }
          if (Object.keys(diff).length === 0) diff = null;
        }
      }
      let status;
      if (requiredMissing) { status = 'invalid'; cntInvReq++; }
      else if (anoResolutionFailed) { status = 'invalid'; cntInvAno++; }
      else if (existing) {
        status = 'duplicate'; cntDup++;
        if (diff) cntDupChanged++;
      } else { status = 'new'; cntNew++; }
      const entry = { row: i + 2, status, reasons, values, existing: existing || null, diff };
      if (status === 'new' && sampleNew.length < PER_BUCKET_LIMIT) sampleNew.push(entry);
      else if (status === 'invalid' && sampleInvalid.length < PER_BUCKET_LIMIT) sampleInvalid.push(entry);
      else if (status === 'duplicate' && diff && sampleDupChanges.length < PER_BUCKET_LIMIT) sampleDupChanges.push(entry);
      else if (status === 'duplicate' && !diff && sampleDupIdentical.length < PER_BUCKET_LIMIT) sampleDupIdentical.push(entry);
    }
    let targetRowCount = 0;
    try {
      const r = db.get(`SELECT COUNT(*) as c FROM ${def.table}`);
      targetRowCount = r ? Number(r.c || 0) : 0;
    } catch (_) {}
    fs.unlink(req.file.path, () => {});
    res.json({
      module: moduleKey, label: def.label, total,
      fields: def.fields, keyCols: def.keyCols,
      autoFillAuctionId: !!def.autoFillAuctionId,
      sampleLimit: PER_BUCKET_LIMIT,
      targetTable: def.table, targetRowCount,
      counts: { new: cntNew, duplicate: cntDup, duplicateChanged: cntDupChanged, invalidAno: cntInvAno, invalidRequired: cntInvReq, nameCorrected: 0, rundDerived: 0 },
      samples: { new: sampleNew, invalid: sampleInvalid, dupChanges: sampleDupChanges, dupIdentical: sampleDupIdentical },
    });
  } catch (e) {
    if (req.file) fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: e.message });
  }
});

app.post('/api/import-old-data/run', requireAdmin, upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const moduleKey = req.body.module;
  const def = IMPORT_MODULES[moduleKey];
  if (!def) {
    fs.unlink(req.file.path, () => {});
    return res.status(400).json({ error: 'Unknown module' });
  }
  const dryRun  = String(req.body.dryRun || '').toLowerCase() === 'true';
  let userMapping = {};
  if (req.body.mapping) {
    try { userMapping = JSON.parse(req.body.mapping) || {}; } catch (_) {}
  }
  const db = getDb();
  let imported = 0, skipped = 0, failed = 0;
  let nameCorrected = 0;
  let rundDerived = 0;
  const errors = [];
  let total = 0;
  const insertedIds = [];
  try {
    const wb = XLSX.readFile(req.file.path);
    const ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws) throw new Error('No worksheet found');
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    total = rows.length;
    if (!total) throw new Error('File is empty');
    const headers  = Object.keys(rows[0]);
    const autoDetected = _importMapHeaders(headers, def);
    const mapping = Object.assign({}, autoDetected, userMapping);
    for (const k of Object.keys(userMapping || {})) {
      const v = userMapping[k];
      if (v === '' || v === null) delete mapping[k];
    }
    const fieldSources = def.fields.map(f => [f, mapping[f] || null]);
    const valuePlaceholders = def.fields.map(() => '?').join(',');
    const insertSql = `INSERT INTO ${def.table} (${def.fields.join(',')}) VALUES (${valuePlaceholders})`;
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
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      try {
        const keyChecks = def.keyCols.map(k => mapping[k] ? r[mapping[k]] : null).filter(v => v != null && v !== '');
        if (keyChecks.length === def.keyCols.length) {
          const whereSql = def.keyCols.map(k => `${k} = ?`).join(' AND ');
          const dup = db.get(`SELECT 1 FROM ${def.table} WHERE ${whereSql} LIMIT 1`, keyChecks);
          if (dup) { skipped++; continue; }
        }
        const values = fieldSources.map(([fname, src]) => {
          const v = src ? r[src] : '';
          if (fname === 'date') return normalizeDate(v);
          return v;
        });
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
  // Back-fill auction_id idempotently
  if (def.autoFillAuctionId) {
    try {
      getDb().run(
        `UPDATE ${def.table}
            SET auction_id = (SELECT id FROM auctions WHERE auctions.ano = ${def.table}.ano)
          WHERE auction_id IS NULL AND ano IS NOT NULL AND ano != ''
            AND EXISTS (SELECT 1 FROM auctions WHERE auctions.ano = ${def.table}.ano)`
      );
    } catch (_) {}
  }
  if (!dryRun) {
    try { repairBadDates(db); } catch (_) {}
  }
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
    nameCorrected, rundDerived,
    errors, importLogId,
  });
});

app.get('/api/import-old-data/history', requireAdmin, (req, res) => {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
  const rows = getDb().all(
    `SELECT id, module, filename, dry_run, total, imported, skipped, failed,
            errors, inserted_ids, undone_at, username, created_at
       FROM import_log ORDER BY id DESC LIMIT 200`
  );
  res.json(rows.map(r => {
    const ids = r.inserted_ids ? safeJSON(r.inserted_ids) : [];
    return {
      id: r.id, module: r.module, filename: r.filename,
      dry_run: !!r.dry_run, total: r.total, imported: r.imported,
      skipped: r.skipped, failed: r.failed,
      errors: r.errors ? safeJSON(r.errors) : [],
      undone_at: r.undone_at || '',
      undoable: !r.dry_run && Array.isArray(ids) && ids.length > 0 && !r.undone_at,
      inserted_count: Array.isArray(ids) ? ids.length : 0,
      username: r.username, created_at: r.created_at,
    };
  }));
});

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
      error: 'This import did not record its inserted row IDs. To clear it, use the Backup tab → Delete All for ' + def.table + '.',
    });
  }
  let backupPath = '';
  try { backupPath = _snapshotBackupBeforeDelete('undo-import-' + logRow.module + '-' + logId); }
  catch (e) {
    return res.status(500).json({ error: 'Backup snapshot failed; refusing to undo: ' + (e.message || e) });
  }
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
    success: true, importLogId: logId,
    module: logRow.module, table: def.table,
    requested: ids.length, deleted, backupPath,
  });
});

function safeJSON(s){ try { return JSON.parse(s); } catch(_) { return []; } }

// ──────────────────────────────────────────────────────────────
// SYSTEM: DB backup & restore (admin-only)
// ──────────────────────────────────────────────────────────────
app.get('/api/system/backups', requireAdmin, (req, res) => {
  try {
    const bkDir = path.join(path.dirname(DB_PATH), 'backups');
    if (!fs.existsSync(bkDir)) return res.json({ backups: [] });
    const out = fs.readdirSync(bkDir)
      .filter(f => f.endsWith('.db'))
      .map(f => {
        const st = fs.statSync(path.join(bkDir, f));
        return { name: f, size: st.size, mtime: st.mtimeMs };
      })
      .sort((a, b) => b.mtime - a.mtime);
    res.json({ backups: out });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/system/backup-now', requireAdmin, (req, res) => {
  try {
    const bkDir = path.join(path.dirname(DB_PATH), 'backups');
    if (!fs.existsSync(bkDir)) fs.mkdirSync(bkDir, { recursive: true });
    require('./db').flushSave();
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
    const out = path.join(bkDir, `manual-${stamp}.db`);
    fs.copyFileSync(DB_PATH, out);
    res.json({ success: true, file: path.basename(out) });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/system/backup', requireAdmin, (req, res) => {
  try {
    require('./db').flushSave();
    if (!fs.existsSync(DB_PATH)) {
      return res.status(500).json({ error: 'Database file not found at ' + DB_PATH });
    }
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
    const filename = `spice-etrade-backup-${stamp}.db`;
    res.setHeader('Content-Type', 'application/octet-stream');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    fs.createReadStream(DB_PATH).pipe(res);
  } catch (e) {
    console.error('[backup] failed:', e);
    res.status(500).json({ error: e.message });
  }
});

const restoreUpload = multer({ dest: uploadDir, limits: { fileSize: 200 * 1024 * 1024 } });
app.post('/api/system/restore', requireAdmin, restoreUpload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded (field name: file)' });
  const tmpPath = req.file.path;
  try {
    const buf = fs.readFileSync(tmpPath);
    const r = await replaceFromBuffer(buf);
    res.json({ ok: true, restoredBytes: r.size });
  } catch (e) {
    console.error('[restore] failed:', e);
    res.status(400).json({ error: e.message });
  } finally {
    try { fs.unlinkSync(tmpPath); } catch(_) {}
  }
});

// ── /api JSON-safety middleware ──
app.use('/api', (req, res) => {
  res.status(404).json({
    error: `Not Found: ${req.method} ${req.originalUrl}`,
    code: 'route_not_found'
  });
});

app.use((err, req, res, next) => {
  console.error('[server] unhandled error:', err);
  if (res.headersSent) return next(err);
  res.status(err.status || 500).json({
    error: err.message || 'Internal server error',
    code: err.code || 'internal_error'
  });
});

function assertSchemaSanity(db) {
  const checks = [
    ['lots',         'auction_id', 'lots tracks the trade via auction_id'],
    ['lots',         'buyer',      'lots.buyer is the buyer code on the lot'],
    ['lots',         'name',       'lots.name is the seller name'],
    ['auctions',     'ano',        'trade number lives on auctions.ano (string)'],
    ['debit_notes',  'ano',        'DN keeps the trade number denormalised on the row'],
    ['purchases',    'invo',       'purchase invoice number'],
    ['purchases',    'name',       'purchases.name is the seller'],
  ];
  for (const [tbl, col, hint] of checks) {
    try {
      const cols = db.all(`PRAGMA table_info(${tbl})`).map(r => r.name);
      if (!cols.includes(col)) {
        console.warn(`[schema-check] ⚠  table ${tbl} is missing expected column "${col}" — ${hint}`);
      }
    } catch (e) {
      console.warn(`[schema-check] could not introspect ${tbl}:`, e.message);
    }
  }
}

function runBackupTickerOnce() {
  try {
    const db = getDb();
    const cfg = getSettingsFlat(db);
    const enabled = String(cfg.backup_auto_enabled || '').toLowerCase() === 'true';
    if (!enabled) return;
    const intervalHrs = Math.max(1, Number(cfg.backup_interval_hours) || 24);
    const keepN       = Math.max(1, Number(cfg.backup_keep_count) || 14);
    const dbDir = path.dirname(DB_PATH);
    const bkDir = path.join(dbDir, 'backups');
    if (!fs.existsSync(bkDir)) fs.mkdirSync(bkDir, { recursive: true });
    const files = fs.readdirSync(bkDir)
      .filter(f => f.endsWith('.db'))
      .map(f => ({ f, m: fs.statSync(path.join(bkDir, f)).mtimeMs }))
      .sort((a, b) => b.m - a.m);
    const latest = files[0] ? files[0].m : 0;
    const dueAt  = latest + intervalHrs * 3600 * 1000;
    if (Date.now() < dueAt) return;
    require('./db').flushSave();
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
    const out = path.join(bkDir, `auto-${stamp}.db`);
    fs.copyFileSync(DB_PATH, out);
    console.log('[backup] auto snapshot written:', out);
    const fresh = fs.readdirSync(bkDir)
      .filter(f => f.endsWith('.db'))
      .map(f => ({ f, m: fs.statSync(path.join(bkDir, f)).mtimeMs }))
      .sort((a, b) => b.m - a.m);
    for (const old of fresh.slice(keepN)) {
      try { fs.unlinkSync(path.join(bkDir, old.f)); console.log('[backup] pruned', old.f); }
      catch (_) {}
    }
  } catch (e) {
    console.error('[backup] ticker failed:', e.message);
  }
}

// ──────────────────────────────────────────────────────────────
// START
// ──────────────────────────────────────────────────────────────
function repairBadDates(db) {
  const tables = ['auctions', 'bills', 'debit_notes', 'invoices', 'purchases', 'lots'];
  let totalFixed = 0;
  for (const tbl of tables) {
    try {
      const hasDate = db.all(`PRAGMA table_info(${tbl})`).some(c => c.name === 'date');
      if (!hasDate) continue;
      const rows = db.all(`SELECT rowid, date FROM ${tbl} WHERE date IS NOT NULL AND date != ''`);
      let fixed = 0;
      for (const r of rows) {
        const current = String(r.date);
        if (/^\d{4}-\d{2}-\d{2}$/.test(current)) continue;
        const iso = normalizeDate(r.date);
        if (iso && /^\d{4}-\d{2}-\d{2}$/.test(iso) && iso !== current) {
          db.run(`UPDATE ${tbl} SET date = ? WHERE rowid = ?`, [iso, r.rowid]);
          fixed++;
        }
      }
      if (fixed > 0) console.log(`  Date repair: ${tbl} — fixed ${fixed} row(s)`);
      totalFixed += fixed;
    } catch (_) {}
  }
  if (totalFixed > 0) console.log(`  Date repair: ${totalFixed} total row(s) normalized to yyyy-mm-dd`);
}

const PORT = process.env.PORT || 3001;
(async () => {
  const db = await initDb();
  initCompanySettings(db);
  repairBadDates(db);
  assertSchemaSanity(db);
  setInterval(runBackupTickerOnce, 60 * 1000);
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n  Admin Console running at http://localhost:${PORT}\n`);
  });
})();
