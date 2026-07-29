/**
 * trade-fair.js — outbound client for tradefair.intelloinsights.com.
 *
 * Replaces the standalone pull_trade_history.py script. The trade-fair
 * site is a Rails app behind a CSRF-protected, mobile-OTP login; the
 * thing that persists after login is the `_kcpmc_rails_session` cookie.
 * Rather than automate the OTP flow, the operator pastes that cookie
 * into Settings → Integrations (see trade_fair_config table) and this
 * module replays the two endpoints the on-screen tables use:
 *
 *   1. get_trade_fair_history  — a jQuery DataTables server-side
 *      endpoint that returns auction/session history as JSON. We
 *      replay its column map verbatim (the server's field mapping is
 *      picky) and paginate through it.
 *   2. download_price_list_report_excel — returns the per-auction
 *      price list as an .xlsx, keyed by auction_id / trade_session_id.
 *
 * Nothing here touches the DB. server.js resolves the stored config,
 * calls these functions, and feeds the downloaded .xlsx through the
 * SAME price-import pipeline a manual upload uses (runLotImport).
 *
 * Session expiry surfaces as a clear, actionable error (the history
 * endpoint answers with HTML/redirect instead of JSON once the cookie
 * dies) so the operator knows to re-copy the cookie.
 */

'use strict';

// The DataTables column map captured from the live request. The server
// maps incoming `columns[i][data]` to its own field names, so this list
// has to match what the page sends or rows come back empty. draw/start/
// length/_ are filled in per request by buildHistoryParams().
const HISTORY_COLUMNS = [
  { data: 'auction_number',          searchable: true,  orderable: true  },
  { data: 'company_name',            searchable: false, orderable: false },
  { data: 'trade_type',              searchable: true,  orderable: true  },
  { data: 'sb_auction_number',       searchable: true,  orderable: true  },
  { data: 'auction_date',            searchable: true,  orderable: true  },
  { data: 'trade_session_status',    searchable: false, orderable: false },
  { data: 'trade_session_start_time',searchable: false, orderable: false },
  { data: 'trade_session_end_time',  searchable: false, orderable: false },
  { data: 'auction_lot_count',       searchable: false, orderable: false },
  { data: 'auction_lot_qty',         searchable: false, orderable: false },
  { data: 'trade_session_sell_qty',  searchable: false, orderable: false },
];

// ── Config resolution helpers ─────────────────────────────────

const trim = (v) => (v == null ? '' : String(v).trim());
const noSlash = (u) => trim(u).replace(/\/+$/, '');

// Build the `Cookie:` header from the stored value. The operator can
// paste EITHER the raw session value OR a full `name=value; name2=...`
// header copied from DevTools — if it already contains an '=', we trust
// it verbatim; otherwise we wrap it as `<cookie_name>=<value>`.
function buildCookieHeader(cfg) {
  const raw = trim(cfg.sessionCookie);
  if (!raw) return '';
  if (raw.includes('=')) return raw;
  return `${cfg.cookieName || '_kcpmc_rails_session'}=${raw}`;
}

const UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) ' +
  'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36';

function baseHeaders(cfg, extra) {
  const cookie = buildCookieHeader(cfg);
  return Object.assign({
    'User-Agent': UA,
    'X-Requested-With': 'XMLHttpRequest',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Referer': noSlash(cfg.baseUrl) + '/spices/admin',
    'Cookie': cookie,
  }, extra || {});
}

// ── OTP login (mobile + OTP, Rails Devise + CSRF) ─────────────
// Discovered from the trade-fair site's own login JS:
//   • login page:  GET  /trading/dashboard/sign_in  (sets _kcpmc_rails_session,
//                  carries <meta name="csrf-token">)
//   • send OTP:    GET  /users/generate_otp?user[mobile_number]=<m>&source=web
//   • verify OTP:  POST /users/sign_in  user[mobile_number],user[otp],source=web
// The CSRF token + session cookie from the login page must ride along on
// both calls; the sign_in response hands back the AUTHENTICATED session
// cookie, which is what we store and reuse for history/price-list calls.

// Pull a named cookie value out of a fetch Response's Set-Cookie header(s).
function _cookieFromResponse(res, name) {
  let list = [];
  if (typeof res.headers.getSetCookie === 'function') list = res.headers.getSetCookie();
  else { const sc = res.headers.get('set-cookie'); if (sc) list = [sc]; }
  for (const c of list) {
    const m = String(c).match(new RegExp('(?:^|[;,]\\s*)' + name + '=([^;]+)'));
    if (m) return m[1];
  }
  return '';
}

// Load the login page → { csrf, session }. Both are needed to send/verify.
async function _loginPage(cfg) {
  const url = noSlash(cfg.baseUrl) + '/trading/dashboard/sign_in';
  let r;
  try {
    r = await fetch(url, { headers: { 'User-Agent': UA, 'Accept': 'text/html,application/xhtml+xml' } });
  } catch (e) {
    throw new Error('Could not reach the trade-fair login page: ' + (e.message || e));
  }
  const html = await r.text();
  const m = html.match(/<meta[^>]+name=["']csrf-token["'][^>]+content=["']([^"']+)["']/i)
         || html.match(/<meta[^>]+content=["']([^"']+)["'][^>]+name=["']csrf-token["']/i);
  const csrf = m ? m[1] : '';
  const session = _cookieFromResponse(r, cfg.cookieName || '_kcpmc_rails_session');
  if (!csrf || !session) throw new Error('Could not initialise a login session (no CSRF token / cookie from the site).');
  return { csrf, session };
}

function _normMobile(v) {
  const d = String(v == null ? '' : v).replace(/\D/g, '');
  return d.length === 11 && d[0] === '0' ? d.slice(1) : d; // drop a leading 0
}

// Step 1 — request an OTP to the mobile. Returns { session, csrf,
// otpExpiresAt } to carry into the verify step.
async function loginSendOtp(cfg, mobile) {
  const m = _normMobile(mobile);
  if (!/^\d{10}$/.test(m)) throw new Error('Enter a valid 10-digit mobile number.');
  const { csrf, session } = await _loginPage(cfg);
  const cookieName = cfg.cookieName || '_kcpmc_rails_session';
  const u = new URL(noSlash(cfg.baseUrl) + '/users/generate_otp');
  u.searchParams.set('user[mobile_number]', m);
  u.searchParams.set('source', 'web');
  const r = await fetch(u.toString(), {
    method: 'GET',
    headers: {
      'User-Agent': UA, 'X-Requested-With': 'XMLHttpRequest',
      'Accept': 'application/json, text/javascript, */*; q=0.01',
      'X-CSRF-Token': csrf, 'Cookie': `${cookieName}=${session}`,
      'Referer': noSlash(cfg.baseUrl) + '/trading/dashboard/sign_in',
    },
  });
  const j = await r.json().catch(() => ({}));
  if (j.status !== 'success') {
    throw new Error(j.error_msg || 'The trade-fair site did not send an OTP — check the mobile number is registered.');
  }
  return { session, csrf, otpExpiresAt: j.otp_expires_at || null };
}

// Step 2 — verify the OTP. Returns { sessionCookie } = the AUTHENTICATED
// cookie to store and reuse. `session`/`csrf` come from loginSendOtp.
async function loginVerifyOtp(cfg, mobile, otp, session, csrf) {
  const m = _normMobile(mobile);
  const code = String(otp == null ? '' : otp).trim();
  if (!/^\d{10}$/.test(m)) throw new Error('Enter a valid 10-digit mobile number.');
  if (!code) throw new Error('Enter the OTP.');
  const cookieName = cfg.cookieName || '_kcpmc_rails_session';
  const body = new URLSearchParams();
  body.set('user[mobile_number]', m);
  body.set('user[otp]', code);
  body.set('source', 'web');
  const r = await fetch(noSlash(cfg.baseUrl) + '/users/sign_in', {
    method: 'POST',
    headers: {
      'User-Agent': UA, 'X-Requested-With': 'XMLHttpRequest',
      'Accept': 'application/json, text/javascript, */*; q=0.01',
      'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
      'X-CSRF-Token': csrf, 'Cookie': `${cookieName}=${session}`,
      'Referer': noSlash(cfg.baseUrl) + '/trading/dashboard/sign_in',
    },
    body: body.toString(),
    redirect: 'manual',
  });
  // The authenticated session cookie is set on this response; fall back
  // to the pre-auth session if the server chose not to rotate it.
  const authCookie = _cookieFromResponse(r, cookieName) || session;
  const j = await r.json().catch(() => ({}));
  if (j.status !== 'success') {
    throw new Error(j.error_msg || 'OTP verification failed — re-check the OTP and try again.');
  }
  return { sessionCookie: authCookie };
}

function buildHistoryParams({ start = 0, length = 200, draw = 1, search = '' }) {
  const p = new URLSearchParams();
  p.set('draw', String(draw));
  HISTORY_COLUMNS.forEach((c, i) => {
    p.set(`columns[${i}][data]`, c.data);
    p.set(`columns[${i}][name]`, '');
    p.set(`columns[${i}][searchable]`, String(c.searchable));
    p.set(`columns[${i}][orderable]`, String(c.orderable));
    p.set(`columns[${i}][search][value]`, '');
    p.set(`columns[${i}][search][regex]`, 'false');
  });
  // Order by auction_number desc → newest auctions first (page 0 is the
  // most recent, which is what the "recent auctions" picker wants).
  p.set('order[0][column]', '0');
  p.set('order[0][dir]', 'desc');
  p.set('start', String(start));
  p.set('length', String(length));
  p.set('search[value]', trim(search));
  p.set('search[regex]', 'false');
  p.set('_', String(Date.now()));
  return p;
}

// DataTables can answer with a raw array OR an envelope
// {recordsTotal, data}. Normalise to { rows, total }.
function extractHistory(payload) {
  if (Array.isArray(payload)) return { rows: payload, total: null };
  if (payload && typeof payload === 'object') {
    return {
      rows: Array.isArray(payload.data) ? payload.data : [],
      total: (payload.recordsTotal != null) ? Number(payload.recordsTotal) : null,
    };
  }
  return { rows: [], total: null };
}

const SESSION_EXPIRED =
  'Trade-fair returned a non-JSON response — the session cookie is most ' +
  'likely expired or invalid. Open the trade-fair site in your browser, ' +
  'log in, copy the _kcpmc_rails_session cookie, and paste it into ' +
  'Settings → Integrations → Trade Fair.';

// ── Public API ────────────────────────────────────────────────

// Fetch a single history page. Throws a session-expired error if the
// endpoint answers with anything other than JSON (Rails redirects an
// unauthenticated XHR to the login page → text/html).
async function fetchHistoryPage(cfg, opts = {}) {
  if (!buildCookieHeader(cfg)) {
    throw new Error('No trade-fair session cookie saved. Paste the _kcpmc_rails_session cookie in Settings → Integrations → Trade Fair first.');
  }
  const url = noSlash(cfg.baseUrl) + cfg.historyPath + '?' + buildHistoryParams(opts).toString();
  let r;
  try {
    r = await fetch(url, { headers: baseHeaders(cfg), redirect: 'manual' });
  } catch (e) {
    throw new Error('Could not reach the trade-fair site: ' + (e.message || e));
  }
  // A redirect to login, or a 401/403, is exactly what an expired/
  // invalid cookie produces — surface the actionable message, not a code.
  if ((r.status >= 300 && r.status < 400) || r.status === 401 || r.status === 403) throw new Error(SESSION_EXPIRED);
  if (!r.ok) throw new Error(`Trade-fair history request failed (HTTP ${r.status}).`);
  const ct = (r.headers.get('content-type') || '').toLowerCase();
  if (!ct.includes('json')) throw new Error(SESSION_EXPIRED);
  const payload = await r.json().catch(() => { throw new Error(SESSION_EXPIRED); });
  return extractHistory(payload);
}

// Fetch up to `limit` of the most recent history rows, paginating as
// needed. The endpoint is newest-first, so we stop once we have enough
// or a short page signals the end.
async function fetchHistory(cfg, { limit = 100, pageSize = 200, search = '' } = {}) {
  const out = [];
  let start = 0, draw = 1, total = null;
  while (out.length < limit) {
    const want = Math.min(pageSize, limit - out.length);
    const { rows, total: t } = await fetchHistoryPage(cfg, { start, length: want, draw, search });
    if (t != null) total = t;
    out.push(...rows);
    if (!rows.length || rows.length < want) break;
    if (total != null && out.length >= total) break;
    start = out.length;
    draw += 1;
  }
  return { rows: out.slice(0, limit), total };
}

// Download one auction's price-list .xlsx. Returns a Buffer. Throws if
// the response is empty or HTML (expired session, or the auction has no
// price list yet).
async function downloadPriceList(cfg, idValue) {
  if (!buildCookieHeader(cfg)) {
    throw new Error('No trade-fair session cookie saved. Paste the _kcpmc_rails_session cookie in Settings → Integrations → Trade Fair first.');
  }
  if (idValue == null || idValue === '') throw new Error('Missing auction id for the price-list download.');
  const u = new URL(noSlash(cfg.baseUrl) + cfg.pricePath);
  u.searchParams.set(cfg.priceParam || 'auction_id', String(idValue));
  let r;
  try {
    r = await fetch(u.toString(), {
      headers: baseHeaders(cfg, { 'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/octet-stream, */*' }),
      redirect: 'manual',
    });
  } catch (e) {
    throw new Error('Could not reach the trade-fair site: ' + (e.message || e));
  }
  if ((r.status >= 300 && r.status < 400) || r.status === 401 || r.status === 403) throw new Error(SESSION_EXPIRED);
  if (!r.ok) throw new Error(`Price-list download failed (HTTP ${r.status}).`);
  const ct = (r.headers.get('content-type') || '').toLowerCase();
  const buf = Buffer.from(await r.arrayBuffer());
  if (!buf.length) throw new Error('The trade-fair price list came back empty for this auction.');
  // Excel files are zip-based (PK\x03\x04) or legacy OLE (\xD0\xCF). An
  // HTML/login page or JSON error means the session died or there's no
  // price list for this id.
  const isXlsx = buf[0] === 0x50 && buf[1] === 0x4b;            // 'PK'
  const isXls  = buf[0] === 0xd0 && buf[1] === 0xcf;            // OLE2
  if (!isXlsx && !isXls) {
    if (ct.includes('html') || ct.includes('json') || ct.includes('text/')) throw new Error(SESSION_EXPIRED);
    throw new Error('The trade-fair response was not a recognised Excel file.');
  }
  return { buffer: buf, contentType: ct, ext: isXls ? '.xls' : '.xlsx' };
}

module.exports = {
  buildCookieHeader,
  buildHistoryParams,
  extractHistory,
  fetchHistory,
  fetchHistoryPage,
  downloadPriceList,
  loginSendOtp,
  loginVerifyOtp,
  HISTORY_COLUMNS,
};
