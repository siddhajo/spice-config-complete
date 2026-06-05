/**
 * license.js — Time-bombed access control for the deployed instance.
 *
 * Model:
 *   • Each install has a unique `install_id` (uuid, generated on first
 *     boot). The dev mints license tokens bound to a specific install.
 *   • Initial state: a 30-day trial starts the moment the server first
 *     boots and writes the license_state row. After expiry the login
 *     endpoint returns 451 and the user is redirected to renew.html.
 *   • A license token is HMAC-SHA256 signed:
 *
 *         <base64url(JSON({install_id, expires_at, issued_at, note?}))>.<hex_sig>
 *
 *   • The dev signs tokens locally via `node tools/license-sign.js
 *     --install-id <id> --days 30`. The server verifies signature +
 *     install-id match before applying.
 *   • Secret: `LICENSE_SECRET` env var, with a fallback that's strictly
 *     for local dev. PRODUCTION DEPLOYS MUST set LICENSE_SECRET in the
 *     Railway dashboard — without it any laptop holding this repo can
 *     mint valid tokens.
 *
 * Design notes:
 *   • The signature scheme is intentionally simple (HMAC-SHA256, base64url).
 *     No JWT lib dependency; we own every byte of the token.
 *   • install_id is generated with crypto.randomUUID() so each install
 *     is independent — a token minted for install A can't unlock install B.
 *   • This is honest-user DRM. A determined attacker with code access
 *     can patch out the gate; that's an acceptable trade-off for solo
 *     dev → solo customer licensing.
 */

const crypto = require('crypto');

// ── Secret ────────────────────────────────────────────────────
// Read once at module load so process.env changes mid-run don't shift
// the signing/verifying behaviour. Trim because env vars on some
// hosting providers ship with stray whitespace.
const ENV_SECRET = String(process.env.LICENSE_SECRET || '').trim();
const FALLBACK_SECRET = 'spice-etrade-license-DEV-ONLY-change-via-LICENSE_SECRET';
const LICENSE_SECRET = ENV_SECRET || FALLBACK_SECRET;
if (!ENV_SECRET && process.env.NODE_ENV === 'production') {
  // Don't crash — that would lock the dev out — but make the warning
  // loud so a Railway deploy without the env var set is obvious in
  // the boot log.
  // eslint-disable-next-line no-console
  console.warn(
    '\n[license] ⚠  LICENSE_SECRET env var is not set. Falling back to ' +
    'the development default. Anyone with this codebase can mint valid ' +
    'tokens. Set LICENSE_SECRET in Railway → Variables before going live.\n'
  );
}

// ── Token codec ───────────────────────────────────────────────
function b64url(buf) {
  return Buffer.from(buf).toString('base64')
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
function b64urlDecode(str) {
  const norm = String(str).replace(/-/g, '+').replace(/_/g, '/');
  const pad = norm.length % 4 ? '='.repeat(4 - (norm.length % 4)) : '';
  return Buffer.from(norm + pad, 'base64');
}
function hmacHex(payload) {
  return crypto.createHmac('sha256', LICENSE_SECRET).update(payload).digest('hex');
}

/**
 * Sign a license payload into a token string. Used by the CLI script.
 *
 * @param {object} payload — { install_id, expires_at, issued_at, note? }
 * @returns {string} token "<b64url(payload)>.<hex_sig>"
 */
function signToken(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('payload required');
  if (!payload.install_id) throw new Error('install_id required');
  if (!payload.expires_at) throw new Error('expires_at required');
  const body = b64url(JSON.stringify(payload));
  const sig  = hmacHex(body);
  return `${body}.${sig}`;
}

/**
 * Verify a license token. Returns { ok, payload?, error? }.
 *
 * Constant-time signature compare to prevent timing leaks (low-stakes
 * here but the discipline is cheap).
 */
function verifyToken(token) {
  if (!token || typeof token !== 'string') return { ok: false, error: 'empty token' };
  const dot = token.indexOf('.');
  if (dot < 1 || dot === token.length - 1) return { ok: false, error: 'malformed token' };
  const body = token.slice(0, dot);
  const sig  = token.slice(dot + 1);
  const expected = hmacHex(body);
  // Constant-time compare — both buffers must be the same length first.
  const a = Buffer.from(sig,      'utf8');
  const b = Buffer.from(expected, 'utf8');
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) {
    return { ok: false, error: 'bad signature' };
  }
  let payload;
  try { payload = JSON.parse(b64urlDecode(body).toString('utf8')); }
  catch (_) { return { ok: false, error: 'malformed payload' }; }
  if (!payload || typeof payload !== 'object') return { ok: false, error: 'malformed payload' };
  if (!payload.install_id) return { ok: false, error: 'payload missing install_id' };
  if (!payload.expires_at) return { ok: false, error: 'payload missing expires_at' };
  return { ok: true, payload };
}

// ── Date helpers ──────────────────────────────────────────────
function todayIso() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
function addDaysIso(days) {
  const d = new Date();
  d.setDate(d.getDate() + Number(days || 0));
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
function daysRemainingFrom(expiresIso) {
  const today = todayIso();
  if (!expiresIso) return 0;
  // Compute calendar-day diff using UTC midnight to avoid TZ flips.
  const toUTC = iso => Date.UTC(
    Number(iso.slice(0, 4)),
    Number(iso.slice(5, 7)) - 1,
    Number(iso.slice(8, 10))
  );
  const diffMs = toUTC(expiresIso) - toUTC(today);
  return Math.floor(diffMs / 86400000);
}

// ── License state CRUD ────────────────────────────────────────
// The license_state table is single-row (CHECK id = 1). On first boot
// we generate a fresh install_id and start a 30-day trial. Subsequent
// boots load the row and the values stick around forever — even across
// redeploys (Railway persistent volume) — until a new token is applied
// or the volume is wiped.
const TRIAL_DAYS = Number(process.env.LICENSE_TRIAL_DAYS) || 30;

function ensureState(db) {
  const row = db.get('SELECT * FROM license_state WHERE id = 1');
  if (row) return row;
  const install_id = crypto.randomUUID();
  const first_seen_at = new Date().toISOString();
  const expires_at = addDaysIso(TRIAL_DAYS);
  db.run(
    `INSERT INTO license_state (id, install_id, first_seen_at, expires_at, active_token)
     VALUES (1, ?, ?, ?, NULL)`,
    [install_id, first_seen_at, expires_at]
  );
  return db.get('SELECT * FROM license_state WHERE id = 1');
}

function getStatus(db) {
  const row = ensureState(db);
  const days_remaining = daysRemainingFrom(row.expires_at);
  return {
    install_id: row.install_id,
    first_seen_at: row.first_seen_at,
    expires_at: row.expires_at,
    days_remaining,
    expired: days_remaining < 0,
    has_token: !!row.active_token,
    trial_days: TRIAL_DAYS,
  };
}

/**
 * Apply a signed token. Verifies signature + install-id match + that
 * the new expiry isn't BEFORE the current expiry (so a stale token
 * can't shorten the license).
 *
 * Returns { ok, status?, error? }.
 */
function applyToken(db, token) {
  const v = verifyToken(token);
  if (!v.ok) return { ok: false, error: 'Invalid token: ' + v.error };
  const current = ensureState(db);
  if (String(v.payload.install_id) !== String(current.install_id)) {
    return { ok: false, error: 'Token was issued for a different install. Send the dev your install ID: ' + current.install_id };
  }
  // Reject tokens whose expiry is in the past — applying them would
  // immediately re-lock the app and look like a no-op to the operator.
  if (daysRemainingFrom(v.payload.expires_at) < 0) {
    return { ok: false, error: 'This token has already expired (' + v.payload.expires_at + '). Ask the dev for a fresh one.' };
  }
  // Don't allow a stale-but-shorter token to roll the expiry back.
  if (v.payload.expires_at < current.expires_at) {
    return { ok: false, error: 'This token (' + v.payload.expires_at + ') would shorten the current license (' + current.expires_at + '). Already covered.' };
  }
  db.run(
    'UPDATE license_state SET expires_at = ?, active_token = ? WHERE id = 1',
    [v.payload.expires_at, token]
  );
  return { ok: true, status: getStatus(db) };
}

module.exports = {
  // Core helpers used by server.js + the CLI
  signToken,
  verifyToken,
  ensureState,
  getStatus,
  applyToken,
  // Constants exposed for the CLI to print + tests to reference
  TRIAL_DAYS,
  // For tests / introspection
  _LICENSE_SECRET_IS_FALLBACK: !ENV_SECRET,
};
