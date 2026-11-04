// GSTIN lookup — two providers behind one button.
//
// Settings → Integrations → "GST Lookup Provider" chooses who answers a
// GSTIN: gstincheck.co.in (key in the URL, credits expire after a year) or
// gstinapi.in (key in a header, credits do not). Each keeps its own key,
// because a key issued by one is refused by the other.
//
// The two fail in opposite ways, and that is what this guards:
//   • gstincheck returns HTTP 200 with {"flag":false,"message":"Credit
//     Expire."} — a dead account looks like a successful request, which is
//     exactly how a month of silently-degraded lookups went unnoticed.
//   • gstinapi returns 402 with {"success":false,"error":"…"}.
// Either way the operator must be TOLD, must still get PAN/state from the
// GSTIN digits, and must never be told to "add an API key" they already set.
//
// Both providers are answered by a local stub, so this test spends no
// credits and needs no network.
const os = require('os'), path = require('path'), fs = require('fs'), http = require('http');
const { spawn } = require('child_process');

const ROOT = path.join(__dirname, '..');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'gst-prov-'));
const PORT = 47396, STUB_PORT = 47397;
const B = `http://127.0.0.1:${PORT}`;

let pass = 0, fail = 0;
function check(name, cond, detail) {
  if (cond) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + (detail ? '\n         ' + detail : '')); }
}

// ── The stub standing in for both portals ────────────────────────────────
// `mode` is flipped by the test to choose what the provider "says" next.
let mode = 'ok';
const hits = [];
const GSTIN = '33ADDFS7113C1ZK';
const stub = http.createServer((req, res) => {
  hits.push({ url: req.url, key: req.headers['x-api-key'] || null });
  const send = (code, obj) => {
    res.writeHead(code, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(obj));
  };
  // gstincheck shape: /check/<key>/<gstin>, key in the path.
  if (req.url.startsWith('/check/')) {
    if (mode === 'credit') return send(200, { flag: false, message: 'Credit Expire.', errorCode: 'CREDIT_NOT_AVAILABLE', data: {} });
    return send(200, { flag: true, message: 'GSTIN found.', data: {
      lgnm: 'Sabira Burvin Traders', tradeNam: 'Sabira Spices',
      sts: 'Active', rgdt: '01/07/2017',
      pradr: { adr: '3,Nondimagan Street, Cumbum, Theni, Tamil Nadu, 625516',
               addr: { bno: '3', st: 'Nondimagan Street', loc: 'Cumbum',
                       dst: 'Theni', stcd: 'Tamil Nadu', pncd: '625516' } },
    } });
  }
  // gstinapi shape: /v1/gstin/<gstin>, key in the x-api-key header.
  if (req.url.startsWith('/v1/gstin/')) {
    if (req.headers['x-api-key'] !== 'NEWKEY') return send(401, { success: false, error: 'Invalid API key' });
    if (mode === 'credit') return send(402, { success: false, error: 'Insufficient credits', credits_remaining: 0 });
    // The same refusal with no number anywhere — some plans answer like this.
    if (mode === 'credit_nonum') return send(402, { success: false, error: 'Insufficient credits' });
    return send(200, { success: true, gstin: GSTIN, credits_remaining: 487, data: {
      gstin: GSTIN, legal_name: 'Sabira Burvin Traders', trade_name: 'Sabira Spices',
      status: 'Active', registration_date: '2017-07-01', state_code: '33',
      address: '3,Nondimagan Street, Cumbum', city: 'Cumbum', pincode: '625516',
      address_details: { building_number: '3', street: 'Nondimagan Street',
                         locality: 'Cumbum', district: 'Theni', pincode: '625516' },
    } });
  }
  send(404, { error: 'no such stub route' });
});

const srv = spawn('node', [path.join(ROOT, 'server.js')], {
  cwd: ROOT,
  env: Object.assign({}, process.env, {
    SPICE_DATA_DIR: TMP, PORT: String(PORT), NODE_ENV: 'test',
    // Both providers answer from the local stub above — no credits, no network.
    GST_API_BASE: `http://127.0.0.1:${STUB_PORT}`,
  }),
  stdio: ['ignore', 'pipe', 'pipe'],
});
let srvLog = '';
srv.stdout.on('data', b => { srvLog += b.toString(); });
srv.stderr.on('data', b => { srvLog += b.toString(); });
function cleanup() {
  try { srv.kill('SIGKILL'); } catch (_) {}
  try { stub.close(); } catch (_) {}
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (_) {}
}

let TOKEN = '';
async function api(method, url, body) {
  const r = await fetch(B + url, {
    method,
    headers: Object.assign({ 'Content-Type': 'application/json' },
      TOKEN ? { Authorization: 'Bearer ' + TOKEN } : {}),
    body: body ? JSON.stringify(body) : undefined,
  });
  let d = null; try { d = await r.json(); } catch (_) {}
  return { status: r.status, d };
}
const setCfg = (settings) => api('PUT', '/api/company-settings', { settings });

// A server left behind by an earlier run answers on the same port, and every
// assertion below would then be made against ITS state — reporting failures
// that have nothing to do with the code under test (it happened: a piped run
// exited before cleanup and the next run "failed" on settings the stale
// process still held). Refuse to start rather than report a false result.
function portFree(port) {
  return new Promise((resolve) => {
    const probe = http.createServer();
    probe.once('error', () => resolve(false));
    probe.once('listening', () => probe.close(() => resolve(true)));
    probe.listen(port, '127.0.0.1');
  });
}

(async () => {
  for (const port of [PORT, STUB_PORT]) {
    if (!(await portFree(port))) {
      console.error(`port ${port} is already in use — a server from an earlier run is probably still alive.`);
      console.error(`kill it first:  lsof -ti:${PORT},${STUB_PORT} | xargs kill -9`);
      cleanup();
      process.exit(1);
    }
  }
  await new Promise(r => stub.listen(STUB_PORT, '127.0.0.1', r));
  for (let i = 0; i < 120; i++) {
    try { const r = await fetch(B + '/api/health'); if (r.status < 500) break; } catch (_) {}
    await new Promise(r => setTimeout(r, 250));
  }
  const login = await api('POST', '/api/login', { username: 'admin', password: 'admin123' });
  TOKEN = login.d && (login.d.token || login.d.accessToken);
  if (!TOKEN) { console.error('login failed', login.d, '\n', srvLog.slice(-2000)); cleanup(); process.exit(1); }

  console.log('[1] The provider setting ships, defaulting to the incumbent');
  const flat = (await api('GET', '/api/company-settings/flat')).d;
  check('gst_api_provider exists', flat.gst_api_provider === 'gstincheck', String(flat.gst_api_provider));
  check('each provider has its own key field',
        flat.gst_api_key === '' && flat.gst_api_key_gstinapi === '',
        JSON.stringify({ a: flat.gst_api_key, b: flat.gst_api_key_gstinapi }));

  console.log('\n[2] No key for the selected provider — structural, and it says so');
  let d = (await api('GET', '/api/gst-lookup/' + GSTIN)).d;
  check('still valid', d.valid === true);
  check('PAN comes from the GSTIN itself', d.pan === 'ADDFS7113C', d.pan);
  check('so does the state', d.state === 'TAMIL NADU' && d.st_code === '33', JSON.stringify([d.state, d.st_code]));
  check('the reason is named', d.reason === 'no_key', d.reason);
  check('and so is the provider', /gstincheck/.test(d.note || ''), d.note);

  console.log('\n[3] gstincheck: a live hit fills the party');
  await setCfg({ gst_api_provider: 'gstincheck', gst_api_key: 'OLDKEY' });
  d = (await api('GET', '/api/gst-lookup/' + GSTIN)).d;
  check('source is live', d.source === 'live', JSON.stringify(d));
  check('the key travelled in the URL path',
        hits.some(h => h.url.startsWith('/check/OLDKEY/')), JSON.stringify(hits.slice(-1)));
  check('trade name, upper-cased', d.tradeName === 'SABIRA SPICES', d.tradeName);
  check('legal name too', d.name === 'SABIRA BURVIN TRADERS', d.name);
  // The flat portal address is re-punctuated (comma spacing normalised) and
  // its district / state / PIN tail dropped, since those print on their own
  // lines in the party block.
  check('address rebuilt without the state/PIN tail',
        d.address === '3, NONDIMAGAN STREET, CUMBUM', d.address);
  check('place and PIN split out', d.place === 'THENI' && d.pin === '625516', JSON.stringify([d.place, d.pin]));
  check('the provider is reported', d.provider === 'gstincheck', d.provider);

  console.log('\n[4] gstinapi: the same GSTIN, the same answer shape');
  await setCfg({ gst_api_provider: 'gstinapi', gst_api_key_gstinapi: 'NEWKEY' });
  const before = hits.length;
  const d2 = (await api('GET', '/api/gst-lookup/' + GSTIN)).d;
  check('source is live', d2.source === 'live', JSON.stringify(d2));
  check('the key travelled in the HEADER, not the URL',
        hits[before] && hits[before].key === 'NEWKEY' && !/NEWKEY/.test(hits[before].url),
        JSON.stringify(hits[before]));
  check('same trade name', d2.tradeName === d.tradeName, `${d2.tradeName} vs ${d.tradeName}`);
  check('same legal name', d2.name === d.name, `${d2.name} vs ${d.name}`);
  check('same PIN', d2.pin === d.pin, `${d2.pin} vs ${d.pin}`);
  check('same state, derived from the GSTIN when the provider omits it',
        d2.state === 'TAMIL NADU', d2.state);
  check('the provider is reported', d2.provider === 'gstinapi', d2.provider);
  // The one field they genuinely disagree on, recorded rather than asserted
  // equal: gstincheck gives the DISTRICT, gstinapi the city.
  console.log(`         (place: gstincheck "${d.place}" vs gstinapi "${d2.place}")`);
  console.log(`         (address: "${d.address}" vs "${d2.address}")`);
  check('both return a place', !!d.place && !!d2.place, JSON.stringify([d.place, d2.place]));
  check('both return the street, and neither tacks the state onto it',
        /NONDIMAGAN STREET/.test(d.address) && /NONDIMAGAN STREET/.test(d2.address)
        && !/TAMIL NADU/.test(d.address) && !/TAMIL NADU/.test(d2.address),
        JSON.stringify([d.address, d2.address]));

  console.log('\n[5] gstinapi reports credits, so the Settings card can show them');
  let st = (await api('GET', '/api/gst-lookup/status')).d;
  check('credits are known', st.credits_remaining === 487, JSON.stringify(st.credits_remaining));
  check('the level is healthy', st.level === 'ok', st.level);
  check('the card names the provider', st.provider === 'gstinapi' && /gstinapi/.test(st.provider_label), JSON.stringify(st));
  check('and points recharge at the right site', /gstinapi\.in/.test(st.recharge_url), st.recharge_url);

  console.log('\n[6] Out of credits on gstinapi — a 402 is reported, not swallowed');
  mode = 'credit';
  d = (await api('GET', '/api/gst-lookup/' + GSTIN)).d;
  check('falls back to structural', d.source === 'structural', d.source);
  check('PAN + state still filled', d.pan === 'ADDFS7113C' && d.st_code === '33', JSON.stringify(d));
  check('the reason is the provider refusing, not a missing key',
        d.reason === 'provider_refused', d.reason);
  check('the provider\'s own words reach the operator', /insufficient credits/i.test(d.note || ''), d.note);
  check('with the HTTP status', d.http_status === 402, String(d.http_status));
  check('and a recharge link', /gstinapi\.in/.test(d.recharge_url || ''), d.recharge_url);

  console.log('\n[7] Out of credits on gstincheck — a 200 that means NO');
  await setCfg({ gst_api_provider: 'gstincheck' });
  d = (await api('GET', '/api/gst-lookup/' + GSTIN)).d;
  check('a 200 with flag:false is still a refusal', d.source === 'structural', d.source);
  check('not mistaken for a missing key', d.reason === 'provider_refused', d.reason);
  check('"Credit Expire." reaches the operator', /credit expire/i.test(d.note || ''), d.note);
  st = (await api('GET', '/api/gst-lookup/status')).d;
  check('and the card reads EXHAUSTED, not "unknown"', st.level === 'exhausted', st.level);
  check('even though the provider sent no credit number',
        st.credits_remaining == null || st.credits_remaining === 0, String(st.credits_remaining));

  console.log('\n[8] A key belongs to ONE provider — switching never reuses the wrong one');
  await setCfg({ gst_api_provider: 'gstinapi', gst_api_key_gstinapi: '' });
  d = (await api('GET', '/api/gst-lookup/' + GSTIN)).d;
  check('the gstincheck key is not offered to gstinapi', d.reason === 'no_key', d.reason);
  check('and the message names the provider that needs one',
        /gstinapi/.test(d.note || ''), d.note);
  await setCfg({ gst_api_provider: 'gstincheck' });
  d = (await api('GET', '/api/gst-lookup/' + GSTIN)).d;
  check('switching back finds the old key still in place',
        d.reason === 'provider_refused', d.reason);

  console.log('\n[9] An unknown provider id falls back rather than breaking the button');
  await setCfg({ gst_api_provider: 'nonsense' });
  d = (await api('GET', '/api/gst-lookup/' + GSTIN)).d;
  check('falls back to gstincheck', d.provider === 'gstincheck', d.provider);
  check('the lookup still answers', d.valid === true && d.pan === 'ADDFS7113C', JSON.stringify(d));

  console.log('\n[10] A malformed GSTIN never reaches a provider (no credit burnt)');
  const n0 = hits.length;
  const bad = await api('GET', '/api/gst-lookup/NOTAGSTIN');
  check('refused with 400', bad.status === 400, String(bad.status));
  check('and no provider was called', hits.length === n0, `${hits.length} vs ${n0}`);

  console.log('\n[11] Each provider keeps its OWN balance — switching never shows the other\'s');
  mode = 'ok';
  await setCfg({ gst_api_provider: 'gstinapi', gst_api_key_gstinapi: 'NEWKEY' });
  await api('GET', '/api/gst-lookup/' + GSTIN);
  st = (await api('GET', '/api/gst-lookup/status')).d;
  check('gstinapi reports its own count', st.credits_remaining === 487 && st.level === 'ok', JSON.stringify(st));
  await setCfg({ gst_api_provider: 'gstincheck' });
  st = (await api('GET', '/api/gst-lookup/status')).d;
  check('gstincheck still reads exhausted, not gstinapi\'s 487',
        st.provider === 'gstincheck' && st.credits_remaining !== 487 && st.level === 'exhausted',
        JSON.stringify(st));
  check('and recharge points at gstincheck', /gstincheck/.test(st.recharge_url || ''), st.recharge_url);
  await setCfg({ gst_api_provider: 'gstinapi' });
  st = (await api('GET', '/api/gst-lookup/status')).d;
  check('switching back restores gstinapi\'s count with no new lookup',
        st.credits_remaining === 487 && st.level === 'ok', JSON.stringify(st));
  check('and the dashboard link names the provider that sells the credits',
        /gstinapi\.in/.test(st.dashboard_url || ''), st.dashboard_url);

  console.log('\n[12] gstinapi out of credits with NO number — still EXHAUSTED, never "unknown"');
  mode = 'credit_nonum';
  d = (await api('GET', '/api/gst-lookup/' + GSTIN)).d;
  check('the lookup falls back to structural', d.source === 'structural', d.source);
  st = (await api('GET', '/api/gst-lookup/status')).d;
  check('the card reads EXHAUSTED', st.level === 'exhausted', JSON.stringify(st));
  check('the 402 is remembered as the evidence', st.last_http_status === 402, String(st.last_http_status));

  console.log(`\n${pass} passed, ${fail} failed\n`);
  cleanup();
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); console.error(srvLog.slice(-3000)); cleanup(); process.exit(1); });
