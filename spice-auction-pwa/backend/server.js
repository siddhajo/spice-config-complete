const express = require('express');
const cors = require('cors');
const crypto = require('crypto');
const path = require('path');
const os = require('os');
const multer = require('multer');
const { initDb, getDb, closeDb } = require('./db');
const { exportXlsx, exportDbf } = require('./export');
const { importSource } = require('./import-source');

const app = express();
const PORT = process.env.PORT || 3000;

// Behind a TLS-terminating proxy (Railway / Render / Fly / nginx / Cloudflare etc.)
// the platform handles HTTPS at the edge and forwards plain HTTP to us. In that
// mode we must NOT start our own TLS listener and must NOT 301-redirect to an
// unreachable internal HTTPS port — we just trust X-Forwarded-Proto.
//
// Auto-detect common PaaS env vars; allow explicit override via TRUST_PROXY=1
// or disable via TRUST_PROXY=0.
const BEHIND_PROXY = (() => {
  if (process.env.TRUST_PROXY === '1') return true;
  if (process.env.TRUST_PROXY === '0') return false;
  // Railway: any of RAILWAY_ENVIRONMENT / RAILWAY_PROJECT_ID / RAILWAY_PUBLIC_DOMAIN.
  // Render: RENDER. Fly: FLY_APP_NAME. Heroku: DYNO. Vercel: VERCEL. Cloud Run: K_SERVICE.
  return !!(
    process.env.RAILWAY_ENVIRONMENT || process.env.RAILWAY_PROJECT_ID || process.env.RAILWAY_PUBLIC_DOMAIN ||
    process.env.RENDER || process.env.FLY_APP_NAME || process.env.DYNO ||
    process.env.VERCEL || process.env.K_SERVICE
  );
})();
if (BEHIND_PROXY) app.set('trust proxy', true);
console.log(`  Detected proxy mode: ${BEHIND_PROXY ? 'YES (HTTP only, trust X-Forwarded-Proto)' : 'NO (local dev)'}`);

const upload = multer({ dest: path.join(__dirname, 'data', 'uploads') });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── HELPERS ─────────────────────────────────────────────────

function hashPassword(password) {
  return crypto.createHash('sha256').update(password).digest('hex');
}

function generateToken() {
  return crypto.randomBytes(32).toString('hex');
}

function auditLog(userId, action, entity, entityId, details) {
  try {
    const db = getDb();
    db.run('INSERT INTO audit_log (user_id, action, entity, entity_id, details) VALUES (?, ?, ?, ?, ?)',
      [userId, action, entity, entityId || null, typeof details === 'object' ? JSON.stringify(details) : (details || null)]);
  } catch (e) { console.error('Audit log error:', e.message); }
}

// ── PII MASKING ────────────────────────────────────────────

function maskPII(val, showLast = 4) {
  if (!val || val.length <= showLast) return val || '';
  return '●'.repeat(val.length - showLast) + val.slice(-showLast);
}

function maskLotPII(lot) {
  if (!lot) return lot;
  return { ...lot,
    acctnum: maskPII(lot.acctnum),
    ifsc: maskPII(lot.ifsc, 4),
    cr: maskPII(lot.cr, 4),
    pan: maskPII(lot.pan, 4),
    tel: maskPII(lot.tel, 4),
  };
}

// ── AUTH MIDDLEWARE ──────────────────────────────────────────

const SESSION_TIMEOUT_MS = 30 * 60 * 1000; // 30 minutes

/** Require a valid user token (any role) */
function requireAuth(req, res, next) {
  const token = req.headers['authorization']?.replace('Bearer ', '') || req.query.token;
  if (!token) return res.status(401).json({ error: 'Login required' });

  const db = getDb();
  const session = db.get(`
    SELECT u.*, s.created_at as session_created FROM sessions s
    JOIN users u ON u.id = s.user_id
    WHERE s.token = ?
  `, [token]);
  if (!session) return res.status(401).json({ error: 'Invalid or expired token' });

  // Session timeout check
  if (session.session_created) {
    const created = new Date(session.session_created).getTime();
    const now = Date.now();
    if (now - created > SESSION_TIMEOUT_MS) {
      db.run('DELETE FROM sessions WHERE token = ?', [token]);
      return res.status(401).json({ error: 'Session expired. Please login again.' });
    }
  }

  req.user = session;
  next();
}

/** Require admin role */
function requireAdmin(req, res, next) {
  requireAuth(req, res, () => {
    if (req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Admin access required' });
    }
    next();
  });
}

// ── AUTH ROUTES ─────────────────────────────────────────────

/**
 * POST /api/auth/login
 * Body: { username, password }
 * Returns: { user, token }
 */
app.post('/api/auth/login', (req, res) => {
  const db = getDb();
  const { username, password, force } = req.body;

  if (!username || !password) {
    return res.status(400).json({ error: 'Username and password required' });
  }

  const user = db.get('SELECT * FROM users WHERE username = ? COLLATE NOCASE', [username]);
  if (!user) return res.status(401).json({ error: 'Invalid username or password' });

  const hash = hashPassword(password);
  if (hash !== user.password_hash) {
    return res.status(401).json({ error: 'Invalid username or password' });
  }

  // Check for existing active session
  const existingSession = db.get('SELECT * FROM sessions WHERE user_id = ?', [user.id]);
  if (existingSession && !force) {
    return res.status(409).json({
      error: 'already_logged_in',
      message: user.username + ' is already logged in on another device. Switch to this device?'
    });
  }

  // Remove all existing sessions for this user (single device only)
  db.run('DELETE FROM sessions WHERE user_id = ?', [user.id]);

  const token = generateToken();
  db.run('INSERT INTO sessions (user_id, token) VALUES (?, ?)', [user.id, token]);
  auditLog(user.username, 'login', 'user', user.id, {role: user.role, force: !!force});

  // Record login history
  const ip = req.headers['x-forwarded-for'] || req.connection.remoteAddress || '';
  const ua = req.headers['user-agent'] || '';
  const device = /Mobile|Android|iPhone/i.test(ua) ? 'Mobile' : 'Desktop';
  db.run('INSERT INTO login_history (user_id, username, ip, user_agent) VALUES (?, ?, ?, ?)',
    [user.id, user.username, ip, device]);
  // Keep only last 100 entries per user
  db.run('DELETE FROM login_history WHERE user_id = ? AND id NOT IN (SELECT id FROM login_history WHERE user_id = ? ORDER BY id DESC LIMIT 100)', [user.id, user.id]);

  res.json({
    user: { id: user.id, username: user.username, role: user.role, branch: user.branch || '' },
    token,
  });
});

/**
 * POST /api/auth/logout
 * Clears the user's token
 */
app.post('/api/auth/logout', requireAuth, (req, res) => {
  const db = getDb();
  const token = req.headers['authorization']?.replace('Bearer ', '') || req.query.token;
  db.run('DELETE FROM sessions WHERE token = ?', [token]);
  res.json({ success: true });
});

// ── PASSWORD CHANGE ──
app.post('/api/auth/change-password', requireAuth, (req, res) => {
  const db = getDb();
  const { current_password, new_password } = req.body;
  if (!current_password || !new_password) return res.status(400).json({ error: 'Both current and new password required' });
  if (new_password.length < 6) return res.status(400).json({ error: 'New password must be at least 6 characters' });

  const user = db.get('SELECT * FROM users WHERE id = ?', [req.user.id]);
  if (!user) return res.status(404).json({ error: 'User not found' });

  if (hashPassword(current_password) !== user.password_hash) {
    return res.status(401).json({ error: 'Current password is incorrect' });
  }

  db.run('UPDATE users SET password_hash = ? WHERE id = ?', [hashPassword(new_password), req.user.id]);
  auditLog(req.user.username, 'password-change', 'user', req.user.id, {});
  res.json({ success: true, message: 'Password changed successfully' });
});

// Admin reset password for any user
app.post('/api/auth/reset-password/:userId', requireAdmin, (req, res) => {
  const db = getDb();
  const userId = parseInt(req.params.userId);
  const { new_password } = req.body;
  if (!new_password || new_password.length < 6) return res.status(400).json({ error: 'Password must be at least 6 characters' });

  const user = db.get('SELECT * FROM users WHERE id = ?', [userId]);
  if (!user) return res.status(404).json({ error: 'User not found' });

  db.run('UPDATE users SET password_hash = ? WHERE id = ?', [hashPassword(new_password), userId]);
  // Invalidate their sessions
  db.run('DELETE FROM sessions WHERE user_id = ?', [userId]);
  auditLog(req.user.username, 'reset-password', 'user', userId, {target: user.username});
  res.json({ success: true, message: 'Password reset for ' + user.username });
});

// ── LOGIN HISTORY ──
app.delete('/api/login-history', requireAdmin, (req, res) => {
  const db = getDb();
  db.run('DELETE FROM login_history');
  auditLog(req.user.username, 'clear', 'login_history');
  res.json({ success: true });
});

app.get('/api/login-history', requireAdmin, (req, res) => {
  const db = getDb();
  const userId = req.query.user_id;
  let logs;
  if (userId) {
    logs = db.all('SELECT * FROM login_history WHERE user_id = ? ORDER BY created_at DESC LIMIT 50', [parseInt(userId)]);
  } else {
    logs = db.all('SELECT * FROM login_history ORDER BY created_at DESC LIMIT 100');
  }
  res.json({ logs });
});

/**
 * GET /api/auth/me
 * Returns current user info (validates token)
 */
app.get('/api/auth/me', requireAuth, (req, res) => {
  res.json({
    user: { id: req.user.id, username: req.user.username, role: req.user.role, branch: req.user.branch || '' }
  });
});

/**
 * POST /api/auth/users — Admin only: create a new user
 * Body: { username, password, role, branch }
 */
app.post('/api/auth/users', requireAdmin, (req, res) => {
  const db = getDb();
  const { username, password, role, branch } = req.body;

  if (!username || !password) {
    return res.status(400).json({ error: 'Username and password required' });
  }

  const existing = db.get('SELECT id FROM users WHERE username = ? COLLATE NOCASE', [username]);
  if (existing) return res.status(409).json({ error: 'Username already exists' });

  const hash = hashPassword(password);
  const result = db.run(
    'INSERT INTO users (username, password_hash, role, branch) VALUES (?, ?, ?, ?)',
    [username, hash, role || 'user', branch || '']
  );

  res.status(201).json({
    user: { id: result.lastInsertRowid, username, role: role || 'user', branch: branch || '' }
  });
});

/**
 * GET /api/auth/users — Admin only: list all users
 */
app.get('/api/auth/users', requireAdmin, (req, res) => {
  const db = getDb();
  const users = db.all(`SELECT u.id, u.username, u.role, u.branch, u.created_at,
    s.created_at as session_start,
    (SELECT lh.created_at FROM login_history lh WHERE lh.user_id = u.id ORDER BY lh.id DESC LIMIT 1) as last_login,
    (SELECT lh.user_agent FROM login_history lh WHERE lh.user_id = u.id ORDER BY lh.id DESC LIMIT 1) as last_device,
    CASE WHEN s.id IS NOT NULL THEN 1 ELSE 0 END as is_online
    FROM users u LEFT JOIN sessions s ON s.user_id = u.id ORDER BY is_online DESC, u.id ASC`);
  res.json({ users });
});

/**
 * DELETE /api/auth/users/:id — Admin only: delete a user
 */
app.delete('/api/auth/users/:id', requireAdmin, (req, res) => {
  const db = getDb();
  const userId = parseInt(req.params.id);

  if (userId === req.user.id) {
    return res.status(400).json({ error: 'Cannot delete yourself' });
  }

  const user = db.get('SELECT * FROM users WHERE id = ?', [userId]);
  if (!user) return res.status(404).json({ error: 'User not found' });

  db.run('DELETE FROM sessions WHERE user_id = ?', [userId]);
  db.run('DELETE FROM users WHERE id = ?', [userId]);
  res.json({ deleted: true, username: user.username });
});

/**
 * PUT /api/auth/users/:id/password — Admin only: reset a user's password
 * Body: { password }
 */
app.put('/api/auth/users/:id/password', requireAdmin, (req, res) => {
  const db = getDb();
  const { password } = req.body;
  if (!password) return res.status(400).json({ error: 'Password required' });

  const user = db.get('SELECT * FROM users WHERE id = ?', [parseInt(req.params.id)]);
  if (!user) return res.status(404).json({ error: 'User not found' });

  const hash = hashPassword(password);
  db.run('UPDATE users SET password_hash = ?, token = NULL WHERE id = ?', [hash, user.id]);
  res.json({ success: true, username: user.username });
});

// ── CONFIG (branches, crop types, title) ────────────────────

/**
 * GET /api/config?type=branch  (public for app dropdowns)
 * Returns config values by type
 */
app.get('/api/config', (req, res) => {
  const db = getDb();
  const { type } = req.query;

  if (type) {
    const items = db.all('SELECT * FROM config WHERE type = ? ORDER BY sort_order ASC, value ASC', [type]);
    return res.json({ items });
  }

  // Return all config grouped by type
  const branches = db.all("SELECT * FROM config WHERE type = 'branch' ORDER BY sort_order ASC, value ASC");
  const cropTypes = db.all("SELECT * FROM config WHERE type = 'crop_type' ORDER BY sort_order ASC, value ASC");
  const titleRow = db.get("SELECT value FROM config WHERE type = 'title' LIMIT 1");
  const timeoutRow = db.get("SELECT value FROM config WHERE type = 'edit_timeout' LIMIT 1");
  const editEnabledRow = db.get("SELECT value FROM config WHERE type = 'edit_enabled' LIMIT 1");
  const pageLimitRow = db.get("SELECT value FROM config WHERE type = 'page_limit' LIMIT 1");
  const showUsernameRow = db.get("SELECT value FROM config WHERE type = 'show_username' LIMIT 1");
  const tradeTileTitleRow = db.get("SELECT value FROM config WHERE type = 'trade_tile_title' LIMIT 1");
  const sampleWeightRow = db.get("SELECT value FROM config WHERE type = 'sample_weight' LIMIT 1");
  const showMoistureRow = db.get("SELECT value FROM config WHERE type = 'show_moisture' LIMIT 1");
  const acctMaskRow = db.get("SELECT value FROM config WHERE type = 'acct_mask' LIMIT 1");
  const defaultLitreRow = db.get("SELECT value FROM config WHERE type = 'default_litre' LIMIT 1");
  const labelsRow = db.get("SELECT value FROM config WHERE type = 'labels' LIMIT 1");
  let labels = {};
  try { if (labelsRow) labels = JSON.parse(labelsRow.value); } catch(e) {}
  res.json({ branches, cropTypes, title: titleRow ? titleRow.value : 'Spice Auction', editTimeout: timeoutRow ? parseInt(timeoutRow.value) : 0, editEnabled: editEnabledRow ? editEnabledRow.value === 'true' : true, pageLimit: pageLimitRow ? parseInt(pageLimitRow.value) : 20, showUsername: showUsernameRow ? showUsernameRow.value === 'true' : false, tradeTileTitle: tradeTileTitleRow ? tradeTileTitleRow.value : 'Active Trade', sampleWeight: sampleWeightRow ? parseFloat(sampleWeightRow.value) : 0, showMoisture: showMoistureRow ? showMoistureRow.value === 'true' : false, acctMask: acctMaskRow ? acctMaskRow.value : 'none', defaultLitre: defaultLitreRow ? defaultLitreRow.value : '', labels });
});

/**
 * POST /api/config — Admin only: add a config value
 * Body: { type, value }
 */
app.post('/api/config', requireAdmin, (req, res) => {
  const db = getDb();
  const { type, value } = req.body;

  if (!type || !value) {
    return res.status(400).json({ error: 'type and value required' });
  }
  if (!['branch', 'crop_type', 'title', 'edit_timeout', 'edit_enabled', 'page_limit', 'show_username', 'trade_tile_title', 'sample_weight', 'show_moisture', 'acct_mask', 'default_litre', 'labels'].includes(type)) {
    return res.status(400).json({ error: 'Invalid config type' });
  }

  // Single-value types — replace instead of add
  if (type === 'title' || type === 'edit_timeout' || type === 'edit_enabled' || type === 'page_limit' || type === 'show_username' || type === 'trade_tile_title' || type === 'sample_weight' || type === 'show_moisture' || type === 'acct_mask' || type === 'default_litre' || type === 'labels') {
    db.run(`DELETE FROM config WHERE type = '${type}'`);
    const result = db.run(
      'INSERT INTO config (type, value, sort_order) VALUES (?, ?, ?)',
      [type, value, 0]
    );
    return res.status(201).json({ id: result.lastInsertRowid, type, value });
  }

  const storeValue = value.toUpperCase();
  const existing = db.get('SELECT id FROM config WHERE type = ? AND value = ? COLLATE NOCASE', [type, storeValue]);
  if (existing) return res.status(409).json({ error: 'Already exists' });

  const maxOrder = db.get('SELECT COALESCE(MAX(sort_order), 0) as m FROM config WHERE type = ?', [type]);
  const result = db.run(
    'INSERT INTO config (type, value, sort_order) VALUES (?, ?, ?)',
    [type, storeValue, maxOrder.m + 1]
  );

  res.status(201).json({ id: result.lastInsertRowid, type, value: storeValue });
});

/**
 * DELETE /api/config/:id — Admin only: remove a config value
 */
app.delete('/api/config/:id', requireAdmin, (req, res) => {
  const db = getDb();
  const item = db.get('SELECT * FROM config WHERE id = ?', [parseInt(req.params.id)]);
  if (!item) return res.status(404).json({ error: 'Not found' });

  db.run('DELETE FROM config WHERE id = ?', [parseInt(req.params.id)]);
  res.json({ deleted: true, value: item.value });
});

// ── TRADERS ─────────────────────────────────────────────────

app.get('/api/traders', requireAuth, (req, res) => {
  const db = getDb();
  const query = (req.query.q || '').trim();
  const limit = parseInt(req.query.limit) || 500;
  const isAdmin = req.user.role === 'admin';

  function maskTraders(list) {
    if (isAdmin) return list;
    return list.map(t => ({ ...t, acctnum: maskPII(t.acctnum), ifsc: maskPII(t.ifsc, 4), pan: maskPII(t.pan, 4), aadhar: maskPII(t.aadhar, 4) }));
  }

  if (!query) {
    const traders = db.all('SELECT * FROM traders ORDER BY name ASC LIMIT ?', [limit]);
    return res.json({ traders: maskTraders(traders), total: traders.length });
  }

  const traders = db.all(
    `SELECT * FROM traders WHERE
      name LIKE ? COLLATE NOCASE OR
      tel LIKE ? COLLATE NOCASE OR
      cr LIKE ? COLLATE NOCASE OR
      pan LIKE ? COLLATE NOCASE OR
      whatsapp LIKE ? COLLATE NOCASE OR
      email LIKE ? COLLATE NOCASE
    ORDER BY name ASC`,
    [`%${query}%`, `%${query}%`, `%${query}%`, `%${query}%`, `%${query}%`, `%${query}%`]
  );
  res.json({ traders: maskTraders(traders), total: traders.length });
});

app.get('/api/traders/:id', requireAuth, (req, res) => {
  const db = getDb();
  const trader = db.get('SELECT * FROM traders WHERE id = ?', [parseInt(req.params.id)]);
  if (!trader) return res.status(404).json({ error: 'Trader not found' });
  res.json(trader);
});

/**
 * POST /api/traders — Add a new trader/seller
 * Body: { name, cr, pan, tel, aadhar, padd, ppla, pin, pstate, pst_code, ifsc, acctnum, whatsapp, email }
 */
app.post('/api/traders', requireAuth, (req, res) => {
  const db = getDb();
  const { name, cr, pan, tel, aadhar, padd, ppla, pin, pstate, pst_code, ifsc, acctnum, whatsapp, email } = req.body;

  if (!name || !name.trim()) {
    return res.status(400).json({ error: 'Seller name is required' });
  }

  const emailClean = (email || '').trim();
  if (emailClean && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailClean)) {
    return res.status(400).json({ error: 'Invalid email format' });
  }

  const result = db.run(
    `INSERT INTO traders (name, cr, pan, tel, aadhar, padd, ppla, pin, pstate, pst_code, ifsc, acctnum, whatsapp, email)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [name.trim().toUpperCase(), cr||'', pan||'', tel||'', aadhar||'', padd||'', ppla||'', pin||'', pstate||'', pst_code||'', ifsc||'', acctnum||'', (whatsapp||'').trim(), emailClean]
  );

  const trader = db.get('SELECT * FROM traders WHERE id = ?', [result.lastInsertRowid]);
  scheduleSyncSource();
  res.status(201).json({ trader });
});

// Get seller's last lot + all bank accounts
app.get('/api/traders/:id/last-lot', requireAuth, (req, res) => {
  const db = getDb();
  const traderId = parseInt(req.params.id);
  const lot = db.get(`
    SELECT l.grade, l.litre, l.bags, l.branch
    FROM lots l WHERE l.trader_id = ?
    ORDER BY l.created_at DESC LIMIT 1
  `, [traderId]);
  var banks = db.all('SELECT * FROM trader_banks WHERE trader_id = ? ORDER BY is_default DESC, id ASC', [traderId]);
  // Auto-migrate: if no records in trader_banks but traders table has bank details, migrate now
  if (!banks.length) {
    const trader = db.get('SELECT acctnum, ifsc FROM traders WHERE id = ?', [traderId]);
    if (trader && trader.acctnum && trader.acctnum.trim()) {
      db.run('INSERT INTO trader_banks (trader_id, acctnum, ifsc, is_default) VALUES (?, ?, ?, 1)',
        [traderId, trader.acctnum.trim(), trader.ifsc || '']);
      banks = db.all('SELECT * FROM trader_banks WHERE trader_id = ? ORDER BY is_default DESC, id ASC', [traderId]);
    }
  }
  res.json({ lastLot: lot || null, banks });
});

// Update trader basic info
app.put('/api/traders/:id', requireAuth, (req, res) => {
  const db = getDb();
  const id = parseInt(req.params.id);
  const { acctnum, ifsc, whatsapp, email } = req.body;
  const trader = db.get('SELECT * FROM traders WHERE id = ?', [id]);
  if (!trader) return res.status(404).json({ error: 'Trader not found' });
  if (acctnum !== undefined) db.run('UPDATE traders SET acctnum = ? WHERE id = ?', [acctnum, id]);
  if (ifsc !== undefined) db.run('UPDATE traders SET ifsc = ? WHERE id = ?', [ifsc, id]);
  if (whatsapp !== undefined) db.run('UPDATE traders SET whatsapp = ? WHERE id = ?', [String(whatsapp).trim(), id]);
  if (email !== undefined) {
    const emailClean = String(email).trim();
    if (emailClean && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailClean)) {
      return res.status(400).json({ error: 'Invalid email format' });
    }
    db.run('UPDATE traders SET email = ? WHERE id = ?', [emailClean, id]);
  }
  const updated = db.get('SELECT * FROM traders WHERE id = ?', [id]);
  scheduleSyncSource();
  res.json({ trader: updated });
});

// ── BANK ACCOUNTS CRUD ──

// Helper: sync traders.acctnum/ifsc with the default bank account
function syncTraderBank(traderId) {
  const db = getDb();
  const def = db.get('SELECT acctnum, ifsc FROM trader_banks WHERE trader_id = ? ORDER BY is_default DESC, id DESC LIMIT 1', [traderId]);
  if (def) {
    db.run('UPDATE traders SET acctnum = ?, ifsc = ? WHERE id = ?', [def.acctnum || '', def.ifsc || '', traderId]);
  } else {
    // No banks left — clear traders table so auto-migrate doesn't re-create
    db.run('UPDATE traders SET acctnum = ?, ifsc = ? WHERE id = ?', ['', '', traderId]);
  }
  scheduleSyncSource();
}

// ── AUTO-SYNC SOURCE FILE ──
let syncSourceTimer = null;
function scheduleSyncSource() {
  // Debounce: wait 2 seconds after last change before writing
  if (syncSourceTimer) clearTimeout(syncSourceTimer);
  syncSourceTimer = setTimeout(syncSourceFile, 2000);
}

async function syncSourceFile() {
  try {
    const db = getDb();
    const traders = db.all('SELECT * FROM traders ORDER BY id ASC');
    if (!traders.length) return;

    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('NAM');
    sheet.columns = [
      { header: 'NAME', key: 'name', width: 30 }, { header: 'CR', key: 'cr', width: 28 },
      { header: 'PAN', key: 'pan', width: 14 }, { header: 'TEL', key: 'tel', width: 16 },
      { header: 'AADHAR', key: 'aadhar', width: 16 }, { header: 'PADD', key: 'padd', width: 50 },
      { header: 'PPLA', key: 'ppla', width: 20 }, { header: 'PIN', key: 'pin', width: 10 },
      { header: 'PSTATE', key: 'pstate', width: 14 }, { header: 'PST_CODE', key: 'pst_code', width: 10 },
      { header: 'IFSC', key: 'ifsc', width: 14 }, { header: 'ACCTNUM', key: 'acctnum', width: 20 },
      { header: 'WHATSAPP', key: 'whatsapp', width: 16 }, { header: 'EMAIL', key: 'email', width: 28 },
    ];
    sheet.getRow(1).font = { bold: true };
    traders.forEach(t => sheet.addRow(t));

    const sourcePath = path.join(__dirname, 'data', 'SOURCE.xlsx');
    await workbook.xlsx.writeFile(sourcePath);
    console.log('📄 SOURCE.xlsx auto-synced (' + traders.length + ' traders)');
  } catch (e) {
    console.error('SOURCE.xlsx sync error:', e.message);
  }
}

// Add bank account (becomes default automatically)
app.post('/api/traders/:id/banks', requireAuth, (req, res) => {
  const db = getDb();
  const traderId = parseInt(req.params.id);
  const { acctnum, ifsc, label, holder_name, confirm_acctnum } = req.body;
  if (!acctnum || !acctnum.trim()) return res.status(400).json({ error: 'Account number required' });
  if (confirm_acctnum && acctnum.trim() !== confirm_acctnum.trim()) return res.status(400).json({ error: 'Account numbers do not match' });

  const acctClean = acctnum.trim();
  if (!/^\d{9,18}$/.test(acctClean)) return res.status(400).json({ error: 'Account number must be 9-18 digits' });
  const ifscClean = (ifsc || '').trim().toUpperCase();
  if (ifscClean && !/^[A-Z]{4}0[A-Z0-9]{6}$/.test(ifscClean)) return res.status(400).json({ error: 'Invalid IFSC format (e.g., ICIC0003816)' });

  const dup = db.get('SELECT id FROM trader_banks WHERE trader_id = ? AND acctnum = ?', [traderId, acctClean]);
  if (dup) return res.status(409).json({ error: 'This account already exists' });

  db.run('UPDATE trader_banks SET is_default = 0 WHERE trader_id = ?', [traderId]);
  const result = db.run('INSERT INTO trader_banks (trader_id, acctnum, ifsc, label, holder_name, is_default) VALUES (?, ?, ?, ?, ?, 1)',
    [traderId, acctClean, ifscClean, (label || '').trim(), (holder_name || '').trim()]);

  syncTraderBank(traderId);
  const bank = db.get('SELECT * FROM trader_banks WHERE id = ?', [result.lastInsertRowid]);
  res.status(201).json({ bank });
});

// Update bank account
app.put('/api/traders/:tid/banks/:bid', requireAuth, (req, res) => {
  const db = getDb();
  const bid = parseInt(req.params.bid);
  const { acctnum, ifsc, label, holder_name } = req.body;
  const bank = db.get('SELECT * FROM trader_banks WHERE id = ?', [bid]);
  if (!bank) return res.status(404).json({ error: 'Bank account not found' });

  const acctClean = (acctnum || bank.acctnum).trim();
  if (!/^\d{9,18}$/.test(acctClean)) return res.status(400).json({ error: 'Account number must be 9-18 digits' });
  const ifscClean = (ifsc || bank.ifsc || '').trim().toUpperCase();
  if (ifscClean && !/^[A-Z]{4}0[A-Z0-9]{6}$/.test(ifscClean)) return res.status(400).json({ error: 'Invalid IFSC format (e.g., ICIC0003816)' });

  db.run('UPDATE trader_banks SET acctnum = ?, ifsc = ?, label = ?, holder_name = ? WHERE id = ?',
    [acctClean, ifscClean, label !== undefined ? label : bank.label, holder_name !== undefined ? holder_name : (bank.holder_name || ''), bid]);

  syncTraderBank(bank.trader_id);
  res.json({ bank: db.get('SELECT * FROM trader_banks WHERE id = ?', [bid]) });
});

// Delete bank account
app.delete('/api/traders/:tid/banks/:bid', requireAuth, (req, res) => {
  const db = getDb();
  const bid = parseInt(req.params.bid);
  const bank = db.get('SELECT * FROM trader_banks WHERE id = ?', [bid]);
  if (!bank) return res.status(404).json({ error: 'Bank account not found' });
  db.run('DELETE FROM trader_banks WHERE id = ?', [bid]);
  if (bank.is_default) {
    const next = db.get('SELECT id FROM trader_banks WHERE trader_id = ? LIMIT 1', [bank.trader_id]);
    if (next) db.run('UPDATE trader_banks SET is_default = 1 WHERE id = ?', [next.id]);
  }
  syncTraderBank(bank.trader_id);
  res.json({ deleted: true });
});

// Set default bank account
app.post('/api/traders/:tid/banks/:bid/default', requireAuth, (req, res) => {
  const db = getDb();
  const tid = parseInt(req.params.tid);
  const bid = parseInt(req.params.bid);
  db.run('UPDATE trader_banks SET is_default = 0 WHERE trader_id = ?', [tid]);
  db.run('UPDATE trader_banks SET is_default = 1 WHERE id = ?', [bid]);
  syncTraderBank(tid);
  res.json({ success: true });
});

// ── LOT ALLOCATIONS ─────────────────────────────────────────

// Parse lot number into prefix + numeric parts
function parseLotNo(lot) {
  const match = String(lot).match(/^([A-Za-z]*)(\d+)$/);
  if (!match) return null;
  return { prefix: match[1].toUpperCase(), num: parseInt(match[2]), padLen: match[2].length };
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

// Get allocations for a trade
app.get('/api/auctions/:id/allocations', requireAuth, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id);
  const allocations = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot', [auctionId]);
  res.json({ allocations });
});

// Save allocations for a trade (bulk replace)
app.post('/api/auctions/:id/allocations', requireAdmin, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id);
  const { allocations } = req.body; // [{ branch, start_lot, end_lot }]
  if (!allocations || !Array.isArray(allocations) || !allocations.length) {
    return res.status(400).json({ error: 'At least one allocation is required' });
  }

  // Validate each allocation
  for (const a of allocations) {
    if (!a.branch || !a.start_lot || !a.end_lot) return res.status(400).json({ error: 'Branch, start_lot, end_lot required for each allocation' });
    const s = parseLotNo(a.start_lot);
    const e = parseLotNo(a.end_lot);
    if (!s || !e) return res.status(400).json({ error: `Invalid lot format: ${a.start_lot} or ${a.end_lot}. Use format like 001, A001` });
    if (s.prefix !== e.prefix) return res.status(400).json({ error: `Prefix mismatch: ${a.start_lot} vs ${a.end_lot}` });
    if (s.num > e.num) return res.status(400).json({ error: `Start (${a.start_lot}) must be <= End (${a.end_lot})` });
  }

  // Check for overlapping ranges
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

  // Coverage-based validation: every already-booked lot must still be covered
  // by an allocation for its own branch in the new allocation set. (We can't
  // match by id because the frontend doesn't round-trip the id field, so an
  // id-based "kept" check would flag every row as removed.)
  const bookedLots = db.all('SELECT lot_no, branch FROM lots WHERE auction_id = ?', [auctionId]);
  // Group orphaned (no-longer-covered) booked lots by their existing allocation range
  // so the error message points to the specific row the user effectively removed.
  const orphaned = bookedLots.filter(l => !allocations.some(a =>
    a.branch === l.branch && isLotInRange(l.lot_no, a.start_lot, a.end_lot)
  ));
  if (orphaned.length > 0) {
    const existing = db.all('SELECT * FROM lot_allocations WHERE auction_id = ?', [auctionId]);
    const offendingRange = existing.find(ex =>
      orphaned.some(l => l.branch === ex.branch && isLotInRange(l.lot_no, ex.start_lot, ex.end_lot))
    );
    const sample = orphaned.slice(0, 5).map(l => l.lot_no).join(', ');
    const detail = offendingRange
      ? `${offendingRange.branch} (${offendingRange.start_lot}-${offendingRange.end_lot})`
      : `${orphaned[0].branch}`;
    return res.status(400).json({
      error: `Cannot remove ${detail}: ${orphaned.length} lot(s) already entered (${sample}${orphaned.length > 5 ? '…' : ''})`,
      orphaned_lots: orphaned.map(l => l.lot_no)
    });
  }

  // Clear and re-insert
  db.run('DELETE FROM lot_allocations WHERE auction_id = ?', [auctionId]);
  for (const a of allocations) {
    db.run('INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
      [auctionId, a.branch, String(a.start_lot).trim(), String(a.end_lot).trim()]);
  }

  auditLog(req.user.username, 'edit', 'lot_allocations', auctionId, { count: allocations.length });
  const saved = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot', [auctionId]);
  res.json({ allocations: saved });
});

// Get allocation stats (used/total per branch)
app.get('/api/auctions/:id/allocation-stats', requireAuth, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id);
  const allocations = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot', [auctionId]);
  const lots = db.all('SELECT lot_no, branch FROM lots WHERE auction_id = ?', [auctionId]);
  const usedSet = new Set(lots.map(l => l.lot_no));

  const stats = {};
  for (const a of allocations) {
    if (!stats[a.branch]) stats[a.branch] = { branch: a.branch, total: 0, used: 0, ranges: [] };
    const total = rangeSize(a.start_lot, a.end_lot);
    const usedInRange = lots.filter(l => isLotInRange(l.lot_no, a.start_lot, a.end_lot));
    stats[a.branch].total += total;
    stats[a.branch].used += usedInRange.length;

    // Build lot grid for this range
    const s = parseLotNo(a.start_lot);
    const e = parseLotNo(a.end_lot);
    const lotGrid = [];
    if (s && e) {
      for (let n = s.num; n <= e.num; n++) {
        const lotNo = buildLotNo(s.prefix, n, s.padLen);
        lotGrid.push({ lot: lotNo, used: usedSet.has(lotNo) });
      }
    }
    stats[a.branch].ranges.push({ start: a.start_lot, end: a.end_lot, total, used: usedInRange.length, lots: lotGrid });
  }

  res.json({ stats: Object.values(stats), allocations });
});

// Reassign unused lots from one branch to another
app.post('/api/auctions/:id/reassign-lots', requireAdmin, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id);
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

  // Check the range belongs to from_branch
  const fromAllocs = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ?', [auctionId, from_branch]);
  for (let n = s.num; n <= e.num; n++) {
    const lotNo = buildLotNo(s.prefix, n, s.padLen);
    const inRange = fromAllocs.some(a => isLotInRange(lotNo, a.start_lot, a.end_lot));
    if (!inRange) return res.status(400).json({ error: `Lot ${lotNo} is not allocated to ${from_branch}` });
  }

  // Check none of the lots in the range are already booked (entered).
  // Compare by parsed prefix+num (NOT padded string) so "49" and "049" match correctly.
  const lotKey = (prefix, num) => `${(prefix || '').toUpperCase()}:${num}`;
  const allLots = db.all('SELECT id, lot_no, branch FROM lots WHERE auction_id = ?', [auctionId]);
  const usedByKey = new Map(); // key → { lot_no, branch, id }
  for (const l of allLots) {
    const p = parseLotNo(l.lot_no);
    if (p) usedByKey.set(lotKey(p.prefix, p.num), l);
  }
  const usedInRange = [];
  for (let n = s.num; n <= e.num; n++) {
    const hit = usedByKey.get(lotKey(s.prefix, n));
    if (hit) usedInRange.push(hit.lot_no);
  }
  if (usedInRange.length > 0) {
    return res.status(409).json({
      error: `Cannot reassign — ${usedInRange.length} lot(s) already booked: ${usedInRange.slice(0, 8).join(', ')}${usedInRange.length > 8 ? '…' : ''}`,
      booked_lots: usedInRange
    });
  }

  // Perform reassignment:
  // 1. Shrink/split from_branch allocations
  // 2. Add new allocation to to_branch
  
  // Remove all from_branch allocations and rebuild without the reassigned range
  const fromAllocsAll = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ?', [auctionId, from_branch]);
  db.run('DELETE FROM lot_allocations WHERE auction_id = ? AND branch = ?', [auctionId, from_branch]);

  for (const alloc of fromAllocsAll) {
    const as = parseLotNo(alloc.start_lot);
    const ae = parseLotNo(alloc.end_lot);
    if (!as || !ae) continue;

    // Check if this allocation overlaps with the reassign range
    const overlapStart = Math.max(as.num, s.num);
    const overlapEnd = Math.min(ae.num, e.num);

    if (overlapStart > overlapEnd) {
      // No overlap — keep entire allocation
      db.run('INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
        [auctionId, from_branch, alloc.start_lot, alloc.end_lot]);
    } else {
      // Has overlap — split into before and after
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

  // Add to to_branch
  db.run('INSERT INTO lot_allocations (auction_id, branch, start_lot, end_lot) VALUES (?, ?, ?, ?)',
    [auctionId, to_branch, start_lot, end_lot]);

  auditLog(req.user.username, 'edit', 'lot_reassign', auctionId, { from: from_branch, to: to_branch, range: start_lot + '-' + end_lot });

  const allocs = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? ORDER BY branch, start_lot', [auctionId]);
  res.json({ success: true, allocations: allocs, message: `Lots ${start_lot}-${end_lot} reassigned from ${from_branch} to ${to_branch}` });
});

// Validate a lot number (duplicate + allocation check)
app.get('/api/auctions/:id/validate-lot', requireAuth, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id);
  const lotNo = String(req.query.lot_no || '').trim();
  const branch = req.query.branch || '';
  if (!lotNo) return res.json({ valid: false, error: 'Enter lot number' });

  // Check duplicate
  const dup = db.get('SELECT id FROM lots WHERE auction_id = ? AND lot_no = ?', [auctionId, lotNo]);
  if (dup) return res.json({ valid: false, error: 'Lot #' + lotNo + ' already exists' });

  // Check allocation
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

// Get next available lot number for a branch
app.get('/api/auctions/:id/next-lot/:branch', requireAuth, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.id);
  const branch = req.params.branch;
  const allocations = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ? ORDER BY start_lot', [auctionId, branch]);
  if (!allocations.length) return res.json({ next_lot: null, error: 'No allocation for this branch' });

  const usedLots = db.all('SELECT lot_no FROM lots WHERE auction_id = ?', [auctionId]).map(l => l.lot_no);
  const usedSet = new Set(usedLots);

  for (const a of allocations) {
    const s = parseLotNo(a.start_lot);
    const e = parseLotNo(a.end_lot);
    if (!s || !e) continue;
    for (let n = s.num; n <= e.num; n++) {
      const lotNo = buildLotNo(s.prefix, n, s.padLen);
      if (!usedSet.has(lotNo)) {
        return res.json({ next_lot: lotNo });
      }
    }
  }

  res.json({ next_lot: null, error: 'All lots in this branch are used' });
});

// Logo upload
app.post('/api/logo', requireAdmin, upload.single('logo'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const fs = require('fs');
  const ext = path.extname(req.file.originalname).toLowerCase();
  if (!['.png', '.jpg', '.jpeg'].includes(ext)) return res.status(400).json({ error: 'Only PNG/JPG allowed' });
  const dest = path.join(__dirname, 'data', 'logo' + ext);
  // Remove old logos
  ['.png', '.jpg', '.jpeg'].forEach(e => { try { fs.unlinkSync(path.join(__dirname, 'data', 'logo' + e)); } catch(e) {} });
  fs.copyFileSync(req.file.path, dest);
  res.json({ success: true, url: '/api/logo' });
});

app.get('/api/logo', (req, res) => {
  const fs = require('fs');
  for (const ext of ['.png', '.jpg', '.jpeg']) {
    const p = path.join(__dirname, 'data', 'logo' + ext);
    if (fs.existsSync(p)) { return res.sendFile(p); }
  }
  res.status(404).json({ error: 'No logo' });
});

app.delete('/api/logo', requireAdmin, (req, res) => {
  const fs = require('fs');
  ['.png', '.jpg', '.jpeg'].forEach(ext => { try { fs.unlinkSync(path.join(__dirname, 'data', 'logo' + ext)); } catch(e) {} });
  res.json({ deleted: true });
});

// ── AUCTIONS ────────────────────────────────────────────────

app.post('/api/auctions', requireAuth, (req, res) => {
  const db = getDb();
  const { ano, date, crop_type, start_time, end_time } = req.body;

  if (!ano || !date) {
    return res.status(400).json({ error: 'ano and date are required' });
  }

  const existing = db.get('SELECT * FROM auctions WHERE ano = ? AND date = ?', [ano, date]);
  if (existing) {
    // Admin can update time window on existing auction
    if (start_time !== undefined || end_time !== undefined) {
      db.run('UPDATE auctions SET start_time = ?, end_time = ? WHERE id = ?',
        [start_time !== undefined ? start_time : existing.start_time, end_time !== undefined ? end_time : existing.end_time, existing.id]);
    }
    const auction = db.get('SELECT * FROM auctions WHERE id = ?', [existing.id]);
    const lotCount = db.get('SELECT COUNT(*) as cnt FROM lots WHERE auction_id = ?', [existing.id]).cnt;
    return res.json({ auction, lotCount, isNew: false });
  }

  const result = db.run(
    'INSERT INTO auctions (ano, date, crop_type, start_time, end_time) VALUES (?, ?, ?, ?, ?)',
    [ano, date, crop_type || 'ASP', start_time || null, end_time || null]
  );
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [result.lastInsertRowid]);
  res.status(201).json({ auction, lotCount: 0, isNew: true });
});

app.get('/api/auctions', requireAuth, (req, res) => {
  const db = getDb();
  const auctions = db.all(`
    SELECT a.*,
      (SELECT COUNT(*) FROM lots WHERE auction_id = a.id) as lot_count,
      (SELECT COALESCE(SUM(qty), 0) FROM lots WHERE auction_id = a.id) as total_qty
    FROM auctions a ORDER BY a.date DESC, a.ano DESC
  `);
  res.json({ auctions });
});

// ── LOTS ────────────────────────────────────────────────────

// DELETE auction (only if empty)
app.delete('/api/auctions/:id', requireAdmin, (req, res) => {
  const db = getDb();
  const id = parseInt(req.params.id);
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [id]);
  if (!auction) return res.status(404).json({ error: 'Auction not found' });
  // Delete lots first, then the auction
  db.run('DELETE FROM lots WHERE auction_id = ?', [id]);
  db.run('DELETE FROM auctions WHERE id = ?', [id]);
  res.json({ deleted: true });
});

app.post('/api/lots', requireAuth, (req, res) => {
  const db = getDb();
  const { auction_id, trader_id, branch, grade, bags, litre, qty, lot_no, state, bank_id, gross_weight, sample_weight, moisture } = req.body;
  const user_id = req.user.username;

  if (!auction_id || !trader_id || !branch || bags == null || litre == null || qty == null || !lot_no) {
    return res.status(400).json({
      error: 'Required: auction_id, trader_id, branch, bags, litre, qty, lot_no'
    });
  }

  // Validate no negative values
  if (bags < 0) return res.status(400).json({ error: 'Bags cannot be negative' });
  if (litre < 0) return res.status(400).json({ error: 'Litre weight cannot be negative' });
  if (qty < 0) return res.status(400).json({ error: 'Net weight cannot be negative' });
  if (gross_weight != null && gross_weight < 0) return res.status(400).json({ error: 'Gross weight cannot be negative' });
  if (moisture != null && moisture < 0) return res.status(400).json({ error: 'Moisture cannot be negative' });

  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auction_id]);
  if (!auction) return res.status(404).json({ error: 'Auction not found' });
  if (auction.status !== 'open') return res.status(400).json({ error: 'Auction is closed' });

  // Check time window
  if (auction.start_time || auction.end_time) {
    const now = new Date();
    if (auction.start_time) {
      const start = new Date(auction.date + 'T' + auction.start_time);
      if (now < start) return res.status(400).json({ error: 'Trade #' + auction.ano + ' has not started yet. Entry opens at ' + auction.start_time + '. Please wait.' });
    }
    if (auction.end_time) {
      const end = new Date(auction.date + 'T' + auction.end_time);
      if (now > end) return res.status(400).json({ error: 'Trade #' + auction.ano + ' session closed at ' + auction.end_time + '. No more entries allowed. Please contact admin.' });
    }
  }

  // Branch restriction — users can only enter lots for their assigned branch
  if (req.user.role !== 'admin' && req.user.branch && branch !== req.user.branch) {
    return res.status(403).json({ error: 'You can only enter lots for your branch (' + req.user.branch + ')' });
  }

  // Check lot allocation — lot must fall within branch's assigned range
  const allocs = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ?', [auction_id, branch]);
  if (allocs.length > 0) {
    const lotNoStr = String(lot_no).trim();
    const inRange = allocs.some(a => isLotInRange(lotNoStr, a.start_lot, a.end_lot));
    if (!inRange) {
      const ranges = allocs.map(a => a.start_lot + '-' + a.end_lot).join(', ');
      return res.status(400).json({ error: `Lot #${lot_no} is outside ${branch} allocation (${ranges})` });
    }
  }

  // Check for duplicate lot number
  const existing = db.get('SELECT id FROM lots WHERE auction_id = ? AND lot_no = ?', [auction_id, lot_no]);
  if (existing) {
    return res.status(409).json({ error: `Lot #${lot_no} already exists in this auction` });
  }

  try {
    const lotNoStr = String(lot_no).trim();
    const result = db.run(
      `INSERT INTO lots (auction_id, lot_no, trader_id, branch, grade, bags, litre, qty, user_id, state, bank_id, gross_weight, sample_weight, moisture)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [auction_id, lotNoStr, trader_id, branch, String(grade || '1'), bags, litre, qty, user_id, state || 'TAMIL NADU', bank_id || null, gross_weight || null, sample_weight || 0, moisture || null]
    );

    const lot = db.get(`
      SELECT l.*, COALESCE(t.name, 'Unknown') as trader_name, COALESCE(t.cr,'') as cr, COALESCE(t.pan,'') as pan, COALESCE((SELECT tb.acctnum FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.acctnum FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.acctnum,'') as acctnum, COALESCE((SELECT tb.ifsc FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.ifsc FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.ifsc,'') as ifsc, COALESCE(t.tel,'') as tel, COALESCE(t.ppla,'') as ppla, COALESCE(t.pin,'') as pin
      FROM lots l LEFT JOIN traders t ON t.id = l.trader_id WHERE l.id = ?
    `, [result.lastInsertRowid]);

    auditLog(user_id, 'create', 'lot', result.lastInsertRowid, {lot_no: lotNoStr, trader: lot.trader_name, qty, branch});

    const stats = db.get(
      'SELECT COUNT(*) as lot_count, COALESCE(SUM(qty), 0) as total_qty FROM lots WHERE auction_id = ?',
      [auction_id]
    );

    // Suggest next lot number from allocation
    let suggestedNext = '';
    const branchAllocs = db.all('SELECT * FROM lot_allocations WHERE auction_id = ? AND branch = ? ORDER BY start_lot', [auction_id, branch]);
    if (branchAllocs.length) {
      const usedLots = db.all('SELECT lot_no FROM lots WHERE auction_id = ?', [auction_id]).map(l => l.lot_no);
      const usedSet = new Set(usedLots);
      outer: for (const a of branchAllocs) {
        const s = parseLotNo(a.start_lot);
        const e = parseLotNo(a.end_lot);
        if (!s || !e) continue;
        for (let n = s.num; n <= e.num; n++) {
          const candidate = buildLotNo(s.prefix, n, s.padLen);
          if (!usedSet.has(candidate)) { suggestedNext = candidate; break outer; }
        }
      }
    } else {
      const numPart = lotNoStr.match(/(\d+)$/);
      if (numPart) {
        const nextNum = parseInt(numPart[1]) + 1;
        suggestedNext = lotNoStr.replace(/(\d+)$/, String(nextNum).padStart(numPart[1].length, '0'));
      }
    }

    res.status(201).json({
      lot,
      nextLotNo: suggestedNext,
      sessionStats: {
        lotCount: stats.lot_count,
        totalQty: Math.round(stats.total_qty * 1000) / 1000,
      }
    });
  } catch (err) {
    console.error('Error saving lot:', err.message);
    res.status(500).json({ error: 'Failed to save lot' });
  }
});

app.get('/api/lots', requireAuth, (req, res) => {
  const db = getDb();
  const { auction_id, branch, user_id, seller, page, limit } = req.query;
  if (!auction_id) return res.status(400).json({ error: 'auction_id is required' });

  let where = 'l.auction_id = ?';
  const params = [parseInt(auction_id)];

  if (branch) { where += ' AND l.branch = ?'; params.push(branch); }
  if (user_id) { where += ' AND l.user_id = ?'; params.push(user_id); }
  if (seller) { where += ' AND t.name LIKE ? COLLATE NOCASE'; params.push(`%${seller}%`); }

  // Stats (always from full filtered set)
  const stats = db.get(`
    SELECT COUNT(*) as lot_count, COALESCE(SUM(l.qty), 0) as total_qty,
           COALESCE(SUM(l.bags), 0) as total_bags
    FROM lots l LEFT JOIN traders t ON t.id = l.trader_id WHERE ${where}
  `, params);

  // Pagination — only apply if limit is explicitly provided
  const pageNum = Math.max(1, parseInt(page) || 1);
  const hasLimit = limit !== undefined && limit !== '' && limit !== null;
  const pageSize = hasLimit ? Math.min(100, Math.max(1, parseInt(limit))) : 0;
  let lotsQuery = `
    SELECT l.*, COALESCE(t.name, 'Unknown Trader') as trader_name,
      COALESCE(t.cr,'') as cr, COALESCE(t.pan,'') as pan,
      COALESCE(t.ppla,'') as ppla, COALESCE(t.pin,'') as pin,
      COALESCE((SELECT tb.acctnum FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.acctnum FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.acctnum,'') as acctnum, COALESCE((SELECT tb.ifsc FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.ifsc FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.ifsc,'') as ifsc, COALESCE(t.tel,'') as tel
    FROM lots l LEFT JOIN traders t ON t.id = l.trader_id
    WHERE ${where} ORDER BY CAST(l.lot_no AS INTEGER) ASC, l.lot_no ASC`;

  const lotsParams = [...params];
  if (pageSize > 0) {
    const offset = (pageNum - 1) * pageSize;
    lotsQuery += ' LIMIT ? OFFSET ?';
    lotsParams.push(pageSize, offset);
  }

  const lots = db.all(lotsQuery, lotsParams);

  // Distinct users and branches for filter dropdowns
  const users = db.all('SELECT DISTINCT user_id FROM lots WHERE auction_id = ? ORDER BY user_id', [parseInt(auction_id)]);
  const branches = db.all('SELECT DISTINCT branch FROM lots WHERE auction_id = ? ORDER BY branch', [parseInt(auction_id)]);

  const totalPages = pageSize > 0 ? Math.ceil(stats.lot_count / pageSize) : 1;

  const isAdmin = req.user.role === 'admin';
  const maskedLots = isAdmin ? lots : lots.map(l => maskLotPII(l));

  res.json({
    lots: maskedLots,
    stats: {
      lotCount: stats.lot_count,
      totalQty: Math.round(stats.total_qty * 1000) / 1000,
      totalBags: stats.total_bags,
    },
    pagination: {
      page: pageNum,
      limit: pageSize,
      totalPages,
      totalItems: stats.lot_count,
    },
    filters: {
      users: users.map(u => u.user_id),
      branches: branches.map(b => b.branch),
    }
  });
});

// Print ALL lots for a trade, grouped by seller (one page per seller)
// NOTE: Must be before /api/lots/:id routes to avoid route conflict
app.get('/api/lots/print-all-sellers/:auctionId', requireAuth, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.auctionId);
  const branch = req.query.branch || '';

  let query = `
    SELECT l.*, a.ano, a.date, a.crop_type,
      COALESCE(t.name,'Unknown') as trader_name, COALESCE(t.cr,'') as cr,
      COALESCE(t.ppla,'') as ppla, COALESCE(t.pin,'') as pin,
      COALESCE((SELECT tb.acctnum FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.acctnum FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.acctnum,'') as acctnum,
      COALESCE((SELECT tb.ifsc FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.ifsc FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.ifsc,'') as ifsc
    FROM lots l JOIN auctions a ON a.id = l.auction_id LEFT JOIN traders t ON t.id = l.trader_id
    WHERE l.auction_id = ?`;
  const params = [auctionId];
  if (branch) { query += ' AND l.branch = ?'; params.push(branch); }
  query += ' ORDER BY t.name, l.lot_no';

  const lots = db.all(query, params);
  if (!lots.length) return res.status(404).json({ error: 'No lots found' });

  const cfg = getReceiptConfig(db);

  const sellerGroups = {};
  lots.forEach(l => {
    const key = l.trader_id || 'unknown';
    if (!sellerGroups[key]) sellerGroups[key] = [];
    sellerGroups[key].push(l);
  });
  const groups = Object.values(sellerGroups);
  const r = pickReceiptRenderer(req.query.format);

  const PDFDocument = require('pdfkit');
  const doc = new PDFDocument({ size: r.pageSize, margin: r.compact ? 10 : 20 });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `inline; filename="All_Sellers_Receipt.pdf"`);
  doc.pipe(res);

  groups.forEach((group, idx) => {
    if (idx > 0) doc.addPage();
    r.render(doc, group, cfg);
  });

  doc.end();
});

/**
 * PUT /api/lots/:id — Edit an existing lot
 * Body: { lot_no, trader_id, branch, grade, bags, litre, qty }
 */
app.put('/api/lots/:id', requireAuth, (req, res) => {
  const db = getDb();
  const lotId = parseInt(req.params.id);
  const lot = db.get('SELECT * FROM lots WHERE id = ?', [lotId]);
  if (!lot) return res.status(404).json({ error: 'Lot not found' });

  const { lot_no, trader_id, branch, grade, bags, litre, qty, bank_id, gross_weight, sample_weight, moisture } = req.body;

  // Validate no negative values
  if (bags != null && bags < 0) return res.status(400).json({ error: 'Bags cannot be negative' });
  if (litre != null && litre < 0) return res.status(400).json({ error: 'Litre weight cannot be negative' });
  if (qty != null && qty < 0) return res.status(400).json({ error: 'Net weight cannot be negative' });

  // Recalculate gross weight if qty changed but gross_weight not explicitly provided
  let finalGrossWt = gross_weight;
  let finalSampleWt = sample_weight;
  if (qty != null && gross_weight == null) {
    const sw = sample_weight != null ? sample_weight : (lot.sample_weight || 0);
    finalGrossWt = qty + sw;
    finalSampleWt = sw;
  }

  // Check for duplicate lot number if it changed
  if (lot_no && String(lot_no) !== String(lot.lot_no)) {
    const dup = db.get('SELECT id FROM lots WHERE auction_id = ? AND lot_no = ? AND id != ?',
      [lot.auction_id, String(lot_no), lotId]);
    if (dup) return res.status(409).json({ error: `Lot #${lot_no} already exists` });
  }

  db.run(`UPDATE lots SET
    lot_no = ?, trader_id = ?, branch = ?, grade = ?,
    bags = ?, litre = ?, qty = ?, bank_id = ?,
    gross_weight = ?, sample_weight = ?, moisture = ?
    WHERE id = ?`, [
    String(lot_no || lot.lot_no),
    trader_id || lot.trader_id,
    branch || lot.branch,
    grade || lot.grade,
    bags != null ? bags : lot.bags,
    litre != null ? litre : lot.litre,
    qty != null ? qty : lot.qty,
    bank_id !== undefined ? (bank_id || null) : lot.bank_id,
    finalGrossWt != null ? finalGrossWt : lot.gross_weight,
    finalSampleWt != null ? finalSampleWt : lot.sample_weight,
    moisture !== undefined ? (moisture || null) : lot.moisture,
    lotId
  ]);

  const updated = db.get(`
    SELECT l.*, COALESCE(t.name, 'Unknown') as trader_name, COALESCE(t.cr,'') as cr, COALESCE(t.pan,'') as pan, COALESCE((SELECT tb.acctnum FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.acctnum FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.acctnum,'') as acctnum, COALESCE((SELECT tb.ifsc FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.ifsc FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.ifsc,'') as ifsc, COALESCE(t.tel,'') as tel, COALESCE(t.ppla,'') as ppla, COALESCE(t.pin,'') as pin
    FROM lots l LEFT JOIN traders t ON t.id = l.trader_id WHERE l.id = ?
  `, [lotId]);

  // Build diff of changed fields for audit log (before vs after)
  const diffFields = ['lot_no','trader_id','branch','grade','bags','litre','qty','bank_id','gross_weight','sample_weight','moisture'];
  const diff = {};
  diffFields.forEach(k => {
    const after = updated[k];
    const before = lot[k];
    const a = (after == null ? '' : String(after));
    const b = (before == null ? '' : String(before));
    if (a !== b) diff[k] = { from: before, to: after };
  });
  auditLog(req.user.username, 'edit', 'lot', lotId, {lot_no: updated.lot_no, trader: updated.trader_name, qty: updated.qty, changes: diff});

  res.json({ lot: updated });
});

app.delete('/api/lots/:id', requireAuth, (req, res) => {
  const db = getDb();
  const lot = db.get('SELECT * FROM lots WHERE id = ?', [parseInt(req.params.id)]);
  if (!lot) return res.status(404).json({ error: 'Lot not found' });

  auditLog(req.user.username, 'delete', 'lot', lot.id, {lot_no: lot.lot_no, qty: lot.qty, branch: lot.branch});
  db.run('DELETE FROM lots WHERE id = ?', [parseInt(req.params.id)]);
  res.json({ deleted: true, lot_no: lot.lot_no });
});

/**
 * POST /api/lots/bulk-delete — Delete multiple lots at once
 * Body: { ids: [1,2,3] } or { auction_id: 1, all: true }
 */
app.post('/api/lots/bulk-delete', requireAdmin, (req, res) => {
  const db = getDb();
  const { ids, auction_id, all } = req.body;

  if (all && auction_id) {
    const count = db.get('SELECT COUNT(*) as cnt FROM lots WHERE auction_id = ?', [auction_id]).cnt;
    auditLog(req.user.username, 'bulk-delete', 'lot', null, {auction_id, count, type: 'all'});
    db.run('DELETE FROM lots WHERE auction_id = ?', [auction_id]);
    return res.json({ deleted: count });
  }

  if (ids && Array.isArray(ids) && ids.length > 0) {
    const placeholders = ids.map(() => '?').join(',');
    const count = db.get(`SELECT COUNT(*) as cnt FROM lots WHERE id IN (${placeholders})`, ids).cnt;
    auditLog(req.user.username, 'bulk-delete', 'lot', null, {ids, count, type: 'selected'});
    db.run(`DELETE FROM lots WHERE id IN (${placeholders})`, ids);
    return res.json({ deleted: count });
  }

  res.status(400).json({ error: 'Provide ids array or auction_id with all:true' });
});

// User: clear all MY lots in a trade+branch
app.post('/api/lots/clear-mine', requireAuth, (req, res) => {
  const db = getDb();
  const { auction_id, branch } = req.body;
  if (!auction_id) return res.status(400).json({ error: 'auction_id required' });

  const username = req.user.username;
  const userBranch = branch || req.user.branch || '';

  const count = db.get(
    'SELECT COUNT(*) as cnt FROM lots WHERE auction_id = ? AND user_id = ? AND branch = ?',
    [auction_id, username, userBranch]
  ).cnt;

  if (count === 0) return res.json({ deleted: 0 });

  auditLog(username, 'bulk-delete', 'lot', null, { auction_id, branch: userBranch, count, type: 'clear-mine' });
  db.run('DELETE FROM lots WHERE auction_id = ? AND user_id = ? AND branch = ?',
    [auction_id, username, userBranch]);

  res.json({ deleted: count });
});

// ── EXPORT ──────────────────────────────────────────────────

app.get('/api/export/:auctionId/:format', requireAdmin, async (req, res) => {
  const { auctionId, format } = req.params;
  auditLog(req.user.username, 'export', 'trade', parseInt(auctionId), {format, type: 'lots'});
  try {
    if (format === 'dbf') {
      const result = await exportDbf(parseInt(auctionId));
      res.download(result.filePath, result.fileName, () => {
        result.cleanup(); // Delete temp file after download
      });
    } else {
      const result = await exportXlsx(parseInt(auctionId));
      res.setHeader('Content-Disposition', `attachment; filename="${result.fileName}"`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(result.buffer));
    }
  } catch (err) {
    console.error('Export error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ── LOT RECEIPT PDF ─────────────────────────────────────────

function getLogoPath() {
  const fs = require('fs');
  for (const ext of ['.png', '.jpg', '.jpeg']) {
    const p = path.join(__dirname, 'data', 'logo' + ext);
    if (fs.existsSync(p)) return p;
  }
  return null;
}

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
  doc.fontSize(10).text(branch + ' BRANCH', m, doc.y, { width: w, align: 'center' });
  doc.moveDown(0.4);
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.5).stroke(); doc.moveDown(0.4);

  doc.font('Helvetica').fontSize(10);
  const dateTradeY = doc.y;
  doc.text('Date: ' + dateFmt, m, dateTradeY, { width: w / 2 });
  doc.text('Trade #' + tradeNo, m + w / 2, dateTradeY, { width: w / 2, align: 'right' });
  doc.y = dateTradeY + 16;
  doc.moveDown(0.2);
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).dash(3,{space:3}).lineWidth(0.5).stroke().undash(); doc.moveDown(0.4);
}

// Get lot detail (for edit modal)
app.get('/api/lots/:id/detail', requireAuth, (req, res) => {
  const db = getDb();
  const lot = db.get('SELECT * FROM lots WHERE id = ?', [parseInt(req.params.id)]);
  if (!lot) return res.status(404).json({ error: 'Not found' });
  res.json({ lot });
});

app.get('/api/lots/:id/receipt', requireAuth, (req, res) => {
  const db = getDb();
  const lot = db.get(`
    SELECT l.*, a.ano, a.date, a.crop_type,
      COALESCE(t.name,'Unknown') as trader_name, COALESCE(t.cr,'') as cr,
      COALESCE(t.ppla,'') as ppla, COALESCE(t.pin,'') as pin,
      COALESCE(t.pstate,'') as pstate,
      COALESCE((SELECT tb.acctnum FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.acctnum FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.acctnum,'') as acctnum, COALESCE((SELECT tb.ifsc FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.ifsc FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.ifsc,'') as ifsc
    FROM lots l
    JOIN auctions a ON a.id = l.auction_id
    LEFT JOIN traders t ON t.id = l.trader_id
   
    WHERE l.id = ?
  `, [parseInt(req.params.id)]);
  if (!lot) return res.status(404).json({ error: 'Lot not found' });

  const cfg = getReceiptConfig(db);
  const r = pickReceiptRenderer(req.query.format);

  const PDFDocument = require('pdfkit');
  const doc = new PDFDocument({ size: r.pageSize, margin: r.compact ? 10 : 20 });

  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `inline; filename="Lot_${lot.lot_no}_Receipt.pdf"`);
  doc.pipe(res);

  r.render(doc, [lot], cfg);
  doc.end();
});

// ── BATCH PRINT RECEIPTS ───────────────────────────────────

// Shared helper: render a seller's lots as table receipt on a PDF page
function maskAcctForReceipt(acctnum, maskType) {
  if (!acctnum || !maskType || maskType === 'none') return acctnum;
  const a = String(acctnum);
  if (maskType === 'show_last4') {
    // XXXXXXXX5678
    if (a.length <= 4) return a;
    return '*'.repeat(a.length - 4) + a.slice(-4);
  }
  if (maskType === 'show_first4_last4') {
    // 1234XXXX5678
    if (a.length <= 8) return a;
    return a.slice(0, 4) + '*'.repeat(a.length - 8) + a.slice(-4);
  }
  if (maskType === 'show_last4_star') {
    // ****5678
    if (a.length <= 4) return a;
    return '*'.repeat(a.length - 4) + a.slice(-4);
  }
  return acctnum;
}

function renderSellerReceipt(doc, sellerLots, cfg) {
  const w = 300, m = 20;
  const lot = sellerLots[0];
  const dateFmt = lot.date ? lot.date.split('-').reverse().join('/') : '';
  const L = cfg.labels || {};
  const lb = (key, def) => L[key] || def;

  addReceiptHeader(doc, cfg.appTitle, lot.branch, dateFmt, lot.ano);

  const lw = 70;
  const maskedAcct = maskAcctForReceipt(lot.acctnum, cfg.acctMask);
  const sellerFields = [
    [lb('seller','Seller'), lot.trader_name],
    [lb('place','Place'), [lot.ppla, lot.pin].filter(Boolean).join(', ')],
    [lb('gstin','GSTIN'), lot.cr],
    [lb('acct_no','A/C No'), maskedAcct || '--NIL--'],
    [lb('ifsc','IFSC'), lot.ifsc || '--NIL--'],
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

  // Build columns dynamically (Grade + Litre removed)
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
    const rowData = [l.lot_no, l.bags, Number(l.qty).toFixed(3), sw ? sw.toFixed(3) : '', l.gross_weight != null ? Number(l.gross_weight).toFixed(3) : ''];
    if (cfg.showMoisture) rowData.push(l.moisture ? Number(l.moisture).toFixed(1) : '');
    rowData.forEach((v, i) => { doc.text(String(v), cx, ry, { width: cols[i], align: 'center' }); cx += cols[i]; });
    doc.y = ry + 13;
    totalQty += Number(l.qty) || 0;
    totalGross += Number(l.gross_weight) || 0;
    totalBags += Number(l.bags) || 0;
    totalSample += sw;
  });

  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.5).stroke(); doc.moveDown(0.3);
  doc.font('Helvetica-Bold').fontSize(8);
  let totLine = sellerLots.length + ' lot(s) | ' + totalBags + ' ' + lb('bags','bags') + ' | ' + lb('net_wt','Net') + ': ' + totalQty.toFixed(3);
  if (totalSample) totLine += ' | ' + lb('sample_wt','Smp') + ': ' + totalSample.toFixed(3);
  if (totalGross) totLine += ' | ' + lb('gross_wt','Grs') + ': ' + totalGross.toFixed(3);
  doc.text(totLine, m, doc.y, { width: w, align: 'center' });

  doc.moveDown(0.4);
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.5).stroke(); doc.moveDown(0.2);
  if (cfg.showUser) {
    doc.font('Helvetica').fontSize(8).fillColor('#888').text('Entered by: ' + (lot.user_id || ''), m, doc.y, { width: w });
    doc.moveDown(0.2);
  }
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(10).text('** THANK YOU **', m, doc.y, { width: w, align: 'center' });
}

// ── COMPACT RECEIPT (~2.5" × 3.5" / 180×252pt) ──────────────
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
  doc.fontSize(7.5).text(branch + ' BRANCH', m, doc.y, { width: w, align: 'center' });
  doc.moveDown(0.2);
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.4).stroke(); doc.moveDown(0.2);
  doc.font('Helvetica').fontSize(7);
  const y = doc.y;
  doc.text('Date: ' + dateFmt, m, y, { width: w / 2 });
  doc.text('Trade #' + tradeNo, m + w / 2, y, { width: w / 2, align: 'right' });
  doc.y = y + 10;
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).dash(2,{space:2}).lineWidth(0.4).stroke().undash();
  doc.moveDown(0.2);
}

function renderSellerReceiptCompact(doc, sellerLots, cfg) {
  const w = 160, m = 10;
  const lot = sellerLots[0];
  const dateFmt = lot.date ? lot.date.split('-').reverse().join('/') : '';
  const L = cfg.labels || {};
  const lb = (key, def) => L[key] || def;

  addReceiptHeaderCompact(doc, cfg.appTitle, lot.branch, dateFmt, lot.ano);

  const lw = 32;
  const maskedAcct = maskAcctForReceipt(lot.acctnum, cfg.acctMask);
  const sellerFields = [
    [lb('seller','Seller'), lot.trader_name],
    [lb('place','Place'), [lot.ppla, lot.pin].filter(Boolean).join(', ')],
    [lb('acct_no','A/C'), maskedAcct || '--NIL--'],
    [lb('ifsc','IFSC'), lot.ifsc || '--NIL--'],
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

  // Lot table: Lot, Bags, Net, Gross (Grade + Litre removed)
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
      l.lot_no,
      l.bags,
      Number(l.qty).toFixed(3),
      l.gross_weight != null ? Number(l.gross_weight).toFixed(3) : ''
    ];
    rowData.forEach((v, i) => { doc.text(String(v), cx, ry, { width: cols[i], align: 'center' }); cx += cols[i]; });
    doc.y = ry + 11;
    totalQty += Number(l.qty) || 0;
    totalGross += Number(l.gross_weight) || 0;
    totalBags += Number(l.bags) || 0;
  });

  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.4).stroke(); doc.moveDown(0.15);

  // 2-row summary (no sample wt): row 1 = headers, row 2 = values
  const sumCols = [40, 40, 40, 40];
  const sumHdrs = ['Lots', lb('bags','Bags'), lb('net_wt','Net'), lb('gross_wt','Gross')];
  const sumVals = [
    String(sellerLots.length),
    String(totalBags),
    totalQty.toFixed(3),
    totalGross ? totalGross.toFixed(3) : '-'
  ];
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
    doc.font('Helvetica').fontSize(6).fillColor('#888').text('Entered by: ' + (lot.user_id || ''), m, doc.y, { width: w });
    doc.moveDown(0.15);
  }
  doc.fillColor('#000').font('Helvetica-Bold').fontSize(9).text('** THANK YOU **', m, doc.y, { width: w, align: 'center' });
}

function pickReceiptRenderer(fmt) {
  return fmt === 'compact'
    ? { render: renderSellerReceiptCompact, pageSize: [180, 252], compact: true }
    : { render: renderSellerReceipt, pageSize: [340, 550], compact: false };
}

// Helper to load receipt config
function getReceiptConfig(db) {
  const titleRow = db.get("SELECT value FROM config WHERE type = 'title' LIMIT 1");
  const showUserRow = db.get("SELECT value FROM config WHERE type = 'show_username' LIMIT 1");
  const acctMaskRow = db.get("SELECT value FROM config WHERE type = 'acct_mask' LIMIT 1");
  const labelsRow = db.get("SELECT value FROM config WHERE type = 'labels' LIMIT 1");
  const showMoistureRow = db.get("SELECT value FROM config WHERE type = 'show_moisture' LIMIT 1");
  const sampleWeightRow = db.get("SELECT value FROM config WHERE type = 'sample_weight' LIMIT 1");
  let labels = {};
  try { if (labelsRow) labels = JSON.parse(labelsRow.value); } catch(e) {}
  return {
    appTitle: titleRow ? titleRow.value : 'Spice Auction',
    showUser: showUserRow && showUserRow.value === 'true',
    acctMask: acctMaskRow ? acctMaskRow.value : 'none',
    showMoisture: showMoistureRow && showMoistureRow.value === 'true',
    sampleWeight: sampleWeightRow ? parseFloat(sampleWeightRow.value) : 0,
    labels
  };
}

app.post('/api/lots/print-batch', requireAuth, (req, res) => {
  handlePrintBatch(req.body.ids, req, res);
});

// GET version for mobile window.open
app.get('/api/lots/print-batch', requireAuth, (req, res) => {
  const ids = (req.query.ids || '').split(',').map(Number).filter(n => n > 0);
  handlePrintBatch(ids, req, res);
});

function handlePrintBatch(ids, req, res) {
  const db = getDb();
  if (!ids || !ids.length) return res.status(400).json({ error: 'No lot IDs provided' });

  const cfg = getReceiptConfig(db);
  const r = pickReceiptRenderer(req.query.format || (req.body && req.body.format));

  const lots = ids.map(id => db.get(`
    SELECT l.*, a.ano, a.date, a.crop_type,
      COALESCE(t.name,'Unknown') as trader_name, COALESCE(t.cr,'') as cr,
      COALESCE(t.ppla,'') as ppla, COALESCE(t.pin,'') as pin,
      COALESCE((SELECT tb.acctnum FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.acctnum FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.acctnum,'') as acctnum, COALESCE((SELECT tb.ifsc FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.ifsc FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.ifsc,'') as ifsc
    FROM lots l
    JOIN auctions a ON a.id = l.auction_id
    LEFT JOIN traders t ON t.id = l.trader_id
    WHERE l.id = ?
  `, [parseInt(id)])).filter(Boolean);

  if (!lots.length) return res.status(404).json({ error: 'No lots found' });

  // Group lots by seller
  const sellerGroups = {};
  lots.forEach(l => {
    const key = l.trader_id || 'unknown';
    if (!sellerGroups[key]) sellerGroups[key] = [];
    sellerGroups[key].push(l);
  });
  const groups = Object.values(sellerGroups);

  const PDFDocument = require('pdfkit');
  const doc = new PDFDocument({ size: r.pageSize, margin: r.compact ? 10 : 20 });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `inline; filename="Lots_Receipt_${lots.length}.pdf"`);
  doc.pipe(res);

  groups.forEach((group, idx) => {
    if (idx > 0) doc.addPage();
    r.render(doc, group, cfg);
  });

  doc.end();
}

// Print all lots for a seller on one receipt
app.post('/api/lots/print-seller', requireAuth, (req, res) => {
  handlePrintSeller(req.body.trader_id, req.body.auction_id, req, res);
});

// GET version for mobile window.open
app.get('/api/lots/print-seller', requireAuth, (req, res) => {
  handlePrintSeller(req.query.trader_id, req.query.auction_id, req, res);
});

function handlePrintSeller(trader_id, auction_id, req, res) {
  const db = getDb();
  if (!trader_id || !auction_id) return res.status(400).json({ error: 'trader_id and auction_id required' });

  const lots = db.all(`
    SELECT l.*, a.ano, a.date, a.crop_type,
      COALESCE(t.name,'Unknown') as trader_name, COALESCE(t.cr,'') as cr,
      COALESCE(t.ppla,'') as ppla, COALESCE(t.pin,'') as pin,
      COALESCE((SELECT tb.acctnum FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.acctnum FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.acctnum,'') as acctnum,
      COALESCE((SELECT tb.ifsc FROM trader_banks tb WHERE tb.id=l.bank_id),(SELECT tb.ifsc FROM trader_banks tb WHERE tb.trader_id=t.id ORDER BY tb.is_default DESC,tb.id ASC LIMIT 1),t.ifsc,'') as ifsc
    FROM lots l JOIN auctions a ON a.id = l.auction_id LEFT JOIN traders t ON t.id = l.trader_id
    WHERE l.auction_id = ? AND l.trader_id = ? ORDER BY l.lot_no
  `, [auction_id, trader_id]);

  if (!lots.length) return res.status(404).json({ error: 'No lots found' });

  const cfg = getReceiptConfig(db);
  const fmt = (req.query && req.query.format) || (req.body && req.body.format);
  const r = pickReceiptRenderer(fmt);

  const PDFDocument = require('pdfkit');
  let pageSize;
  if (r.compact) {
    pageSize = [180, Math.min(160 + lots.length * 12 + 60, 700)];
  } else {
    pageSize = [340, Math.min(200 + lots.length * 18 + 80, 800)];
  }
  const doc = new PDFDocument({ size: pageSize, margin: r.compact ? 10 : 20 });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `inline; filename="Seller_${lots[0].trader_name}_Receipt.pdf"`);
  doc.pipe(res);

  r.render(doc, lots, cfg);
  doc.end();
}

// ── SOURCE EXPORT ───────────────────────────────────────────

app.get('/api/export-source/:format', requireAdmin, async (req, res) => {
  const db = getDb();
  const { format } = req.params;
  auditLog(req.user.username, 'export', 'source', null, {format, type: 'traders'});
  
  // traders.acctnum/ifsc is always kept in sync by syncTraderBank()
  const traders = db.all('SELECT * FROM traders ORDER BY id ASC');
  if (traders.length === 0) return res.status(404).json({ error: 'No traders found' });

  // Get all bank accounts for dropdown in export
  const allBanks = db.all('SELECT trader_id, acctnum, ifsc FROM trader_banks ORDER BY trader_id, is_default DESC, id DESC');
  const banksByTrader = {};
  allBanks.forEach(b => {
    if (!banksByTrader[b.trader_id]) banksByTrader[b.trader_id] = [];
    banksByTrader[b.trader_id].push(b);
  });

  try {
    if (format === 'xlsx') {
      const ExcelJS = require('exceljs');
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('NAM');
      sheet.columns = [
        { header: 'NAME', key: 'name', width: 30 }, { header: 'CR', key: 'cr', width: 28 },
        { header: 'PAN', key: 'pan', width: 14 }, { header: 'TEL', key: 'tel', width: 16 },
        { header: 'AADHAR', key: 'aadhar', width: 16 }, { header: 'PADD', key: 'padd', width: 50 },
        { header: 'PPLA', key: 'ppla', width: 20 }, { header: 'PIN', key: 'pin', width: 10 },
        { header: 'PSTATE', key: 'pstate', width: 14 }, { header: 'PST_CODE', key: 'pst_code', width: 10 },
        { header: 'IFSC', key: 'ifsc', width: 18 }, { header: 'ACCTNUM', key: 'acctnum', width: 24 },
        { header: 'WHATSAPP', key: 'whatsapp', width: 16 }, { header: 'EMAIL', key: 'email', width: 28 },
      ];
      sheet.getRow(1).font = { bold: true };

      traders.forEach((t, idx) => {
        const row = sheet.addRow(t);
        const banks = banksByTrader[t.id];
        if (banks && banks.length > 1) {
          // ACCTNUM dropdown (column 12)
          const acctList = banks.map(b => b.acctnum).filter(Boolean);
          if (acctList.length > 1) {
            row.getCell(12).dataValidation = {
              type: 'list',
              allowBlank: true,
              formulae: ['"' + acctList.join(',') + '"']
            };
          }
          // IFSC dropdown (column 11)
          const ifscList = banks.map(b => b.ifsc).filter(Boolean);
          if (ifscList.length > 1) {
            row.getCell(11).dataValidation = {
              type: 'list',
              allowBlank: true,
              formulae: ['"' + ifscList.join(',') + '"']
            };
          }
        }
      });
      const buffer = await workbook.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="SOURCE.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buffer));
    } else if (format === 'dbf') {
      const { DBFFile } = require('dbffile');
      const filePath = require('path').join(require('os').tmpdir(), 'source_' + Date.now() + '.dbf');
      const dbf = await DBFFile.create(filePath, [
        { name: 'NAME', type: 'C', size: 50 }, { name: 'CR', type: 'C', size: 40 },
        { name: 'PAN', type: 'C', size: 14 }, { name: 'TEL', type: 'C', size: 20 },
        { name: 'AADHAR', type: 'C', size: 20 }, { name: 'PADD', type: 'C', size: 80 },
        { name: 'PPLA', type: 'C', size: 30 }, { name: 'PIN', type: 'C', size: 10 },
        { name: 'PSTATE', type: 'C', size: 20 }, { name: 'PST_CODE', type: 'C', size: 10 },
        { name: 'IFSC', type: 'C', size: 14 }, { name: 'ACCTNUM', type: 'C', size: 20 },
        { name: 'WHATSAPP', type: 'C', size: 20 }, { name: 'EMAIL', type: 'C', size: 60 },
      ]);
      await dbf.appendRecords(traders.map(t => ({
        NAME:t.name||'',CR:t.cr||'',PAN:t.pan||'',TEL:t.tel||'',AADHAR:t.aadhar||'',
        PADD:t.padd||'',PPLA:t.ppla||'',PIN:t.pin||'',PSTATE:t.pstate||'',
        PST_CODE:t.pst_code||'',IFSC:t.ifsc||'',ACCTNUM:t.acctnum||'',
        WHATSAPP:t.whatsapp||'',EMAIL:t.email||'',
      })));
      res.download(filePath, 'SOURCE.dbf', () => { try{require('fs').unlinkSync(filePath)}catch(e){} });
    } else {
      res.status(400).json({ error: 'Format must be xlsx or dbf' });
    }
  } catch (err) { res.status(500).json({ error: err.message }); }
});


// ── UPLOAD SOURCE ───────────────────────────────────────────

app.post('/api/upload-source', requireAdmin, upload.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  try {
    const db = getDb();
    const count = await importSource(req.file.path, db);
    res.json({
      success: true,
      traders: count,
      message: `Successfully imported ${count} traders`,
    });
  } catch (err) {
    console.error('Upload error:', err.message);
    res.status(500).json({ error: 'Failed to import: ' + err.message });
  }
});

// ── REPORTS & ANALYTICS ─────────────────────────────────────

// Trade summary — branch-wise and seller-wise breakdown
app.get('/api/reports/trade-summary/:auctionId', requireAdmin, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.auctionId);
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]);
  if (!auction) return res.status(404).json({ error: 'Auction not found' });

  const branchWise = db.all(`
    SELECT branch, COUNT(*) as lot_count, SUM(bags) as total_bags, SUM(qty) as total_qty, COUNT(DISTINCT trader_id) as seller_count
    FROM lots WHERE auction_id = ? GROUP BY branch ORDER BY total_qty DESC
  `, [auctionId]);

  const sellerWise = db.all(`
    SELECT t.name as seller_name, l.branch, COUNT(*) as lot_count, SUM(l.bags) as total_bags, SUM(l.qty) as total_qty
    FROM lots l LEFT JOIN traders t ON t.id = l.trader_id
    WHERE l.auction_id = ? GROUP BY l.trader_id ORDER BY total_qty DESC
  `, [auctionId]);

  const userWise = db.all(`
    SELECT user_id, COUNT(*) as lot_count, SUM(bags) as total_bags, SUM(qty) as total_qty
    FROM lots WHERE auction_id = ? GROUP BY user_id ORDER BY lot_count DESC
  `, [auctionId]);

  const hourly = db.all(`
    SELECT substr(created_at, 12, 2) as hour, COUNT(*) as lot_count, SUM(qty) as total_qty
    FROM lots WHERE auction_id = ? GROUP BY hour ORDER BY hour ASC
  `, [auctionId]);

  const totals = db.get(`
    SELECT COUNT(*) as lot_count, SUM(bags) as total_bags, SUM(qty) as total_qty, COUNT(DISTINCT trader_id) as seller_count, COUNT(DISTINCT branch) as branch_count
    FROM lots WHERE auction_id = ?
  `, [auctionId]);

  const gradeWise = db.all(`
    SELECT grade, COUNT(*) as lot_count, SUM(bags) as total_bags, SUM(qty) as total_qty, COUNT(DISTINCT trader_id) as seller_count
    FROM lots WHERE auction_id = ? GROUP BY grade ORDER BY grade ASC
  `, [auctionId]);

  const showUsernameRow = db.get("SELECT value FROM config WHERE type = 'show_username' LIMIT 1");
  const showUsername = showUsernameRow ? showUsernameRow.value === 'true' : false;

  res.json({ auction, branchWise, sellerWise, userWise: showUsername ? userWise : [], hourly, totals, gradeWise, showUsername });
});

// Seller history across all trades
app.get('/api/reports/seller-history/:traderId', requireAuth, (req, res) => {
  const db = getDb();
  const traderId = parseInt(req.params.traderId);
  const auctionId = req.query.auction_id ? parseInt(req.query.auction_id) : null;
  const trader = db.get('SELECT * FROM traders WHERE id = ?', [traderId]);
  if (!trader) return res.status(404).json({ error: 'Trader not found' });

  let lots, summary;
  if (auctionId) {
    lots = db.all(`
      SELECT l.*, a.ano, a.date, a.crop_type
      FROM lots l JOIN auctions a ON a.id = l.auction_id
      WHERE l.trader_id = ? AND l.auction_id = ? ORDER BY l.lot_no ASC
    `, [traderId, auctionId]);
    summary = db.get(`
      SELECT COUNT(*) as total_lots, SUM(bags) as total_bags, SUM(qty) as total_qty, 1 as trade_count
      FROM lots WHERE trader_id = ? AND auction_id = ?
    `, [traderId, auctionId]);
  } else {
    lots = db.all(`
      SELECT l.*, a.ano, a.date, a.crop_type
      FROM lots l JOIN auctions a ON a.id = l.auction_id
      WHERE l.trader_id = ? ORDER BY a.date DESC, l.lot_no ASC
    `, [traderId]);
    summary = db.get(`
      SELECT COUNT(*) as total_lots, SUM(bags) as total_bags, SUM(qty) as total_qty, COUNT(DISTINCT auction_id) as trade_count
      FROM lots WHERE trader_id = ?
    `, [traderId]);
  }

  res.json({ trader, lots, summary });
});

// Branch comparison across all trades
app.get('/api/reports/branch-comparison', requireAdmin, (req, res) => {
  const db = getDb();
  const data = db.all(`
    SELECT l.branch, a.ano, a.date, COUNT(*) as lot_count, SUM(l.bags) as total_bags, SUM(l.qty) as total_qty
    FROM lots l JOIN auctions a ON a.id = l.auction_id
    GROUP BY l.branch, l.auction_id ORDER BY a.date DESC, l.branch ASC
  `);

  const overall = db.all(`
    SELECT branch, COUNT(*) as lot_count, SUM(bags) as total_bags, SUM(qty) as total_qty, COUNT(DISTINCT auction_id) as trade_count
    FROM lots GROUP BY branch ORDER BY total_qty DESC
  `);

  res.json({ data, overall });
});

// Summary PDF for a trade
app.get('/api/reports/summary-pdf/:auctionId', requireAdmin, (req, res) => {
  const db = getDb();
  const auctionId = parseInt(req.params.auctionId);
  const auction = db.get('SELECT * FROM auctions WHERE id = ?', [auctionId]);
  if (!auction) return res.status(404).json({ error: 'Auction not found' });

  const titleRow = db.get("SELECT value FROM config WHERE type = 'title' LIMIT 1");
  const appTitle = titleRow ? titleRow.value : 'Spice Auction';
  const dateFmt = auction.date ? auction.date.split('-').reverse().join('/') : '';

  const totals = db.get('SELECT COUNT(*) as lots, SUM(bags) as bags, SUM(qty) as qty, COUNT(DISTINCT trader_id) as sellers, COUNT(DISTINCT branch) as branches FROM lots WHERE auction_id = ?', [auctionId]);
  const branchWise = db.all('SELECT branch, COUNT(*) as lots, SUM(bags) as bags, SUM(qty) as qty FROM lots WHERE auction_id = ? GROUP BY branch ORDER BY qty DESC', [auctionId]);
  const sellerWise = db.all('SELECT t.name, l.branch, COUNT(*) as lots, SUM(l.bags) as bags, SUM(l.qty) as qty FROM lots l LEFT JOIN traders t ON t.id = l.trader_id WHERE l.auction_id = ? GROUP BY l.trader_id ORDER BY qty DESC', [auctionId]);

  const PDFDocument = require('pdfkit');
  const doc = new PDFDocument({ size: 'A4', margin: 40 });
  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `attachment; filename="Trade_${auction.ano}_Summary_${auction.date}.pdf"`);
  doc.pipe(res);

  const m = 40, w = 515;

  // Helper: draw a row of columns at fixed y
  function drawRow(y, cols, font, size) {
    doc.font(font || 'Helvetica').fontSize(size || 9);
    cols.forEach(c => { doc.text(String(c.val || ''), c.x, y, { width: c.w, align: c.align || 'left' }); });
    return y + (size || 9) + 5;
  }

  // Logo
  const logoPath = getLogoPath();
  if (logoPath) {
    try { doc.image(logoPath, (595 - 45) / 2, doc.y, { width: 45, height: 45 }); doc.y += 50; } catch (e) {}
  }

  // Header
  doc.font('Helvetica-Bold').fontSize(16).text(appTitle, m, doc.y, { width: w, align: 'center' });
  doc.moveDown(0.2);
  doc.font('Helvetica').fontSize(11).text('Trade #' + auction.ano + '  |  ' + dateFmt + '  |  ' + (auction.crop_type || 'ASP'), m, doc.y, { width: w, align: 'center' });
  doc.moveDown(0.5);
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(1).strokeColor('#1D9E75').stroke();
  doc.strokeColor('#000');
  doc.moveDown(0.6);

  // Summary boxes
  const sY = doc.y;
  doc.font('Helvetica-Bold').fontSize(10);
  const sItems = [
    { label: 'Lots', val: totals.lots || 0 },
    { label: 'Bags', val: totals.bags || 0 },
    { label: 'Qty (kg)', val: Number(totals.qty || 0).toFixed(3) },
    { label: 'Sellers', val: totals.sellers || 0 },
    { label: 'Branches', val: totals.branches || 0 },
  ];
  const sW = w / sItems.length;
  sItems.forEach((s, i) => {
    const sx = m + i * sW;
    doc.font('Helvetica-Bold').fontSize(14).text(String(s.val), sx, sY, { width: sW, align: 'center' });
    doc.font('Helvetica').fontSize(8).fillColor('#666').text(s.label, sx, sY + 16, { width: sW, align: 'center' });
  });
  doc.fillColor('#000');
  doc.y = sY + 36;
  doc.moveDown(0.8);

  // Branch-wise table
  doc.font('Helvetica-Bold').fontSize(11).text('Branch-wise Breakdown', m);
  doc.moveDown(0.4);
  const bX = { branch: m, lots: m + 200, bags: m + 270, qty: m + 340 };
  let y = doc.y;
  y = drawRow(y, [
    { x: bX.branch, w: 200, val: 'Branch' },
    { x: bX.lots, w: 70, val: 'Lots', align: 'right' },
    { x: bX.bags, w: 70, val: 'Bags', align: 'right' },
    { x: bX.qty, w: 100, val: 'Qty (kg)', align: 'right' },
  ], 'Helvetica-Bold', 9);
  doc.moveTo(m, y - 2).lineTo(m + 440, y - 2).lineWidth(0.5).stroke();
  branchWise.forEach(b => {
    y = drawRow(y, [
      { x: bX.branch, w: 200, val: b.branch || '' },
      { x: bX.lots, w: 70, val: b.lots, align: 'right' },
      { x: bX.bags, w: 70, val: b.bags, align: 'right' },
      { x: bX.qty, w: 100, val: Number(b.qty).toFixed(3), align: 'right' },
    ]);
  });
  doc.y = y;
  doc.moveDown(0.8);

  // Top sellers table
  doc.font('Helvetica-Bold').fontSize(11).text('Top Sellers', m);
  doc.moveDown(0.4);
  const sX = { name: m, branch: m + 170, lots: m + 280, bags: m + 330, qty: m + 390 };
  y = doc.y;
  y = drawRow(y, [
    { x: sX.name, w: 170, val: 'Seller' },
    { x: sX.branch, w: 110, val: 'Branch' },
    { x: sX.lots, w: 50, val: 'Lots', align: 'right' },
    { x: sX.bags, w: 60, val: 'Bags', align: 'right' },
    { x: sX.qty, w: 80, val: 'Qty (kg)', align: 'right' },
  ], 'Helvetica-Bold', 9);
  doc.moveTo(m, y - 2).lineTo(m + 470, y - 2).lineWidth(0.5).stroke();
  sellerWise.slice(0, 30).forEach(s => {
    if (y > 750) { doc.addPage(); y = 40; }
    y = drawRow(y, [
      { x: sX.name, w: 170, val: s.name || 'Unknown' },
      { x: sX.branch, w: 110, val: s.branch || '' },
      { x: sX.lots, w: 50, val: s.lots, align: 'right' },
      { x: sX.bags, w: 60, val: s.bags, align: 'right' },
      { x: sX.qty, w: 80, val: Number(s.qty).toFixed(3), align: 'right' },
    ]);
  });

  doc.y = y + 10;
  doc.moveTo(m, doc.y).lineTo(m + w, doc.y).lineWidth(0.5).stroke();
  doc.moveDown(0.4);
  doc.font('Helvetica').fontSize(8).fillColor('#888').text('Generated on ' + new Date().toLocaleString('en-IN'), m, doc.y, { width: w, align: 'center' });

  doc.end();
});

// ── AUDIT LOG ───────────────────────────────────────────────

app.get('/api/audit-log', requireAdmin, (req, res) => {
  const db = getDb();
  const { page, limit } = req.query;
  const pageNum = Math.max(1, parseInt(page) || 1);
  const pageSize = Math.min(100, Math.max(10, parseInt(limit) || 50));
  const offset = (pageNum - 1) * pageSize;

  const total = db.get('SELECT COUNT(*) as cnt FROM audit_log').cnt;
  const logs = db.all(
    'SELECT * FROM audit_log ORDER BY created_at DESC LIMIT ? OFFSET ?',
    [pageSize, offset]
  );

  res.json({ logs, total, page: pageNum, totalPages: Math.ceil(total / pageSize) });
});

app.delete('/api/audit-log', requireAdmin, (req, res) => {
  const db = getDb();
  db.run('DELETE FROM audit_log');
  res.json({ cleared: true });
});

// ── BACKUP & RESTORE ───────────────────────────────────────

app.get('/api/backup', requireAdmin, (req, res) => {
  auditLog(req.user.username, 'export', 'database', null, {type: 'backup'});
  const db = getDb();
  const data = db.export();
  const buffer = Buffer.from(data);
  const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  res.setHeader('Content-Disposition', `attachment; filename="auction_backup_${date}.db"`);
  res.setHeader('Content-Type', 'application/octet-stream');
  res.send(buffer);
});

app.post('/api/restore', requireAdmin, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  try {
    const fs = require('fs');
    const buffer = fs.readFileSync(req.file.path);

    // Validate it's a valid SQLite database
    const SQL = require('sql.js');
    const sqlPromise = await SQL();
    const testDb = new sqlPromise.Database(buffer);
    // Check for required tables
    const tables = testDb.exec("SELECT name FROM sqlite_master WHERE type='table'");
    const tableNames = tables[0] ? tables[0].values.map(v => v[0]) : [];
    const required = ['traders', 'auctions', 'lots', 'users'];
    const missing = required.filter(t => !tableNames.includes(t));
    testDb.close();

    if (missing.length > 0) {
      return res.status(400).json({ error: 'Invalid backup file — missing tables: ' + missing.join(', ') });
    }

    // Save the backup to the data directory
    const dbPath = path.join(__dirname, 'data', 'auction.db');
    fs.writeFileSync(dbPath, buffer);

    // Reload the database
    const { initDb, closeDb } = require('./db');
    closeDb();
    await initDb();

    auditLog('admin', 'restore', 'database', null, { size: buffer.length });
    res.json({ success: true, message: 'Database restored successfully. Please refresh the page.' });
  } catch (err) {
    res.status(500).json({ error: 'Restore failed: ' + err.message });
  }
});

// ── HEALTH CHECK ────────────────────────────────────────────
// Plain 200 with no DB hit so Railway/Render/etc. edge probes always pass.
app.get(['/healthz', '/_health'], (req, res) => res.status(200).type('text/plain').send('ok'));

// ── ADMIN ───────────────────────────────────────────────────

app.get('/admin', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ── PWA MOBILE APP ──────────────────────────────────────────

app.get('/app', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'app.html'));
});

app.get('/manifest.json', (req, res) => {
  const db = getDb();
  const titleRow = db.get("SELECT value FROM config WHERE type = 'title' LIMIT 1");
  const appName = titleRow ? titleRow.value : 'Spice Auction';
  res.json({
    name: appName,
    short_name: appName.length > 12 ? appName.substring(0, 12) : appName,
    start_url: '/app',
    display: 'standalone',
    background_color: '#f4f3ef',
    theme_color: '#1D9E75',
    icons: []
  });
});

// ── PWA MOBILE APP ──────────────────────────────────────────

app.get('/app', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'app.html'));
});

// ── STATUS (public — no auth, needed for connect screen) ───

app.get('/api/status', (req, res) => {
  const db = getDb();
  const traderCount = db.get('SELECT COUNT(*) as cnt FROM traders').cnt;
  const auctionCount = db.get('SELECT COUNT(*) as cnt FROM auctions').cnt;

  const interfaces = os.networkInterfaces();
  const addresses = [];
  for (const name of Object.keys(interfaces)) {
    for (const iface of interfaces[name]) {
      if (iface.family === 'IPv4' && !iface.internal) {
        addresses.push({ interface: name, address: iface.address });
      }
    }
  }

  res.json({
    status: 'running',
    traders: traderCount,
    auctions: auctionCount,
    serverAddresses: addresses,
    port: PORT,
    connectUrl: addresses.length > 0 ? `http://${addresses[0].address}:${PORT}` : `http://localhost:${PORT}`,
  });
});

// ── START ───────────────────────────────────────────────────

const HTTPS_PORT = process.env.HTTPS_PORT || 3443;
const CERT_DIR = path.join(__dirname, 'data', 'certs');
const CERT_PATH = path.join(CERT_DIR, 'server.crt');
const KEY_PATH = path.join(CERT_DIR, 'server.key');

/**
 * Ensure TLS certs exist. Prefers mkcert (locally-trusted, no browser warning)
 * and falls back to a self-signed openssl cert.
 */
function ensureCerts() {
  const fs = require('fs');
  if (fs.existsSync(CERT_PATH) && fs.existsSync(KEY_PATH)) return true;

  fs.mkdirSync(CERT_DIR, { recursive: true });
  const { execSync } = require('child_process');
  const os = require('os');

  // Hostnames the cert should cover. mkcert can list multiple SANs.
  const hosts = ['localhost', '127.0.0.1', '::1'];
  try {
    const interfaces = os.networkInterfaces();
    for (const name of Object.keys(interfaces)) {
      for (const iface of interfaces[name]) {
        if (iface.family === 'IPv4' && !iface.internal) hosts.push(iface.address);
      }
    }
  } catch (e) {}

  // 1) mkcert — produces a cert signed by a local CA the OS already trusts.
  try {
    execSync('mkcert -version', { stdio: 'ignore' });
    console.log('  Generating locally-trusted SSL certificate via mkcert...');
    execSync(`mkcert -install`, { stdio: 'ignore' });
    execSync(`mkcert -key-file "${KEY_PATH}" -cert-file "${CERT_PATH}" ${hosts.join(' ')}`, { stdio: 'ignore' });
    console.log('  ✅ mkcert certificate installed — browsers will trust it.');
    return true;
  } catch (e) {
    // mkcert not available or failed — fall through to openssl.
  }

  // 2) Self-signed fallback. Browser will warn "Not Secure" until trusted manually.
  console.log('  mkcert not found. Generating self-signed SSL certificate...');
  console.log('  ℹ️  For a no-warning setup install mkcert: `brew install mkcert nss`');
  try {
    execSync(`openssl req -x509 -newkey rsa:2048 -keyout "${KEY_PATH}" -out "${CERT_PATH}" -days 365 -nodes -subj "/CN=SpiceAuction/O=Auction/C=IN" 2>/dev/null`);
    console.log('  Self-signed SSL certificate generated (browser will show "Not Secure").');
    return true;
  } catch (e) {
    console.log('  Could not generate SSL cert (openssl not found). Running HTTP only.');
    return false;
  }
}

// ── AUTO-BACKUP SCHEDULER ───────────────────────────────────

const BACKUP_DIR = path.join(__dirname, 'data', 'backups');

function runAutoBackup() {
  try {
    const fs = require('fs');
    if (!fs.existsSync(BACKUP_DIR)) fs.mkdirSync(BACKUP_DIR, { recursive: true });

    const db = getDb();
    const data = db.export();
    const buffer = Buffer.from(data);
    const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const time = new Date().toTimeString().slice(0, 5).replace(':', '');
    const filename = `backup_${date}_${time}.db`;
    const filePath = path.join(BACKUP_DIR, filename);
    fs.writeFileSync(filePath, buffer);

    // Keep only last 7 backups
    const files = fs.readdirSync(BACKUP_DIR).filter(f => f.startsWith('backup_') && f.endsWith('.db')).sort();
    while (files.length > 7) {
      fs.unlinkSync(path.join(BACKUP_DIR, files.shift()));
    }

    console.log('📦 Auto-backup saved:', filename);
  } catch (e) {
    console.error('Auto-backup failed:', e.message);
  }
}

function startAutoBackup() {
  // Run first backup 1 min after start
  setTimeout(runAutoBackup, 60 * 1000);
  // Then every 6 hours
  setInterval(runAutoBackup, 6 * 60 * 60 * 1000);
  console.log('📦 Auto-backup enabled (every 6 hours, keeping last 7)');
}

async function start() {
  await initDb();
  startAutoBackup();
  const db = getDb();
  const traderCount = db.get('SELECT COUNT(*) as cnt FROM traders').cnt;
  const userCount = db.get('SELECT COUNT(*) as cnt FROM users').cnt;

  const http = require('http');
  let httpsRunning = false;

  if (BEHIND_PROXY) {
    // Hosted/PaaS deployment: platform terminates TLS at the edge and forwards
    // plain HTTP to us. Serve HTTP directly — do NOT spin up our own HTTPS
    // listener, and do NOT 301-redirect (the redirect target would be an
    // unreachable internal port).
    console.log('  ⚙  Behind TLS-terminating proxy — serving HTTP only on PORT (proxy handles HTTPS)');
    if (process.env.HSTS_ENABLE === '1') {
      app.use((req, res, next) => {
        if (req.secure) res.setHeader('Strict-Transport-Security', 'max-age=31536000; includeSubDomains');
        next();
      });
    }
    http.createServer(app).listen(PORT, '0.0.0.0');
  } else {
    // Local / self-hosted: try HTTPS first, fall back to HTTP-only with redirect.
    const hasCerts = ensureCerts();
    if (hasCerts) {
      try {
        const https = require('https');
        const fs = require('fs');
        const sslOptions = {
          key: fs.readFileSync(KEY_PATH),
          cert: fs.readFileSync(CERT_PATH),
        };
        if (process.env.HSTS_ENABLE === '1') {
          app.use((req, res, next) => {
            if (req.secure) res.setHeader('Strict-Transport-Security', 'max-age=31536000; includeSubDomains');
            next();
          });
        }
        https.createServer(sslOptions, app).listen(HTTPS_PORT, '0.0.0.0');
        httpsRunning = true;
      } catch (e) {
        console.log('  HTTPS failed to start:', e.message);
      }
    }

    if (httpsRunning) {
      // Force every plain-HTTP request to HTTPS so traffic is always encrypted.
      http.createServer((req, res) => {
        const host = (req.headers.host || '').split(':')[0];
        const target = 'https://' + host + ':' + HTTPS_PORT + req.url;
        res.writeHead(301, { Location: target });
        res.end();
      }).listen(PORT, '0.0.0.0');
    } else {
      console.log('  ⚠️  HTTPS unavailable — serving HTTP only (traffic NOT encrypted)');
      http.createServer(app).listen(PORT, '0.0.0.0');
    }
  }

  console.log('');
  console.log('='.repeat(55));
  console.log('  SPICE AUCTION SERVER');
  console.log('='.repeat(55));
  console.log(`  HTTP Port:   ${PORT}`);
  if (httpsRunning) console.log(`  HTTPS Port:  ${HTTPS_PORT}`);
  if (BEHIND_PROXY) console.log(`  Mode:        behind TLS proxy (platform handles HTTPS)`);
  console.log(`  Traders:     ${traderCount} loaded`);
  console.log(`  Users:       ${userCount} registered`);
  console.log('');

  if (BEHIND_PROXY) {
    console.log(`  App is reachable at the platform-supplied public URL.`);
    console.log(`  Local container:     http://localhost:${PORT}`);
  } else {
    const interfaces = os.networkInterfaces();
    for (const name of Object.keys(interfaces)) {
      for (const iface of interfaces[name]) {
        if (iface.family === 'IPv4' && !iface.internal) {
          if (httpsRunning) {
            console.log(`  Mobile app (HTTPS):  https://${iface.address}:${HTTPS_PORT}/app`);
            console.log(`  Admin dashboard:     https://${iface.address}:${HTTPS_PORT}/admin`);
            console.log(`  (HTTP port ${PORT} → 301 redirect to HTTPS)`);
          } else {
            console.log(`  Mobile app (HTTP):   http://${iface.address}:${PORT}/app`);
            console.log(`  Admin dashboard:     http://${iface.address}:${PORT}/admin`);
          }
          console.log('');
        }
      }
    }
    if (httpsRunning) console.log(`  Local (HTTPS):       https://localhost:${HTTPS_PORT}`);
    else console.log(`  Local:               http://localhost:${PORT}`);
  }
  console.log('');
  console.log('  Default admin:       admin / admin123');
  console.log('');
  if (httpsRunning) {
    console.log('  NOTE: Phones will show a security warning for the');
    console.log('  self-signed cert. Tap "Advanced" → "Proceed" to accept.');
  }
  console.log('');
  console.log('  Ready for lot entries!');
  console.log('='.repeat(55));
}

start().catch(err => {
  console.error('Failed to start server:', err);
  process.exit(1);
});

process.on('SIGINT', () => {
  console.log('\nShutting down...');
  closeDb();
  process.exit(0);
});
