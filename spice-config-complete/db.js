/**
 * db.js — SQL.js variant for cloud deploy (Railway etc.)
 *
 * Why: better-sqlite3 needs native compilation that's been failing on
 * Railway's build infra. sql.js is pure JavaScript — no native bindings,
 * no architecture issues, no compile step.
 *
 * Trade-off: sql.js holds the entire DB in memory and writes the whole
 * file on every commit. For a single-server Railway deployment this is
 * fine since concurrent writes within one Node process are sequential.
 * For multi-machine deploys, switch back to better-sqlite3.
 *
 * Compatibility: This wrapper preserves the same API server.js,
 * calculations.js, company-config.js, exports.js, etc. already use:
 *
 *   db.run(sql, params)           // INSERT/UPDATE/DELETE (params array or spread)
 *   db.get(sql, params)           // SELECT one row
 *   db.all(sql, params)           // SELECT many rows
 *   db.exec(sql)                  // multi-statement SQL
 *   db.prepare(sql).run(...args)  // prepared INSERT/UPDATE
 *   db.prepare(sql).get(...args)  // prepared SELECT one
 *   db.prepare(sql).all(...args)  // prepared SELECT many
 *   db.transaction(fn)            // returns a wrapped function
 */

const initSqlJs = require('sql.js');
const crypto = require('crypto');
const bcrypt = require('bcryptjs');
const fs = require('fs');
const path = require('path');

// DB path: defaults to ./data/config.db (dev / standalone node).
// Electron packaging sets SPICE_DATA_DIR to %APPDATA%\SpiceConfig so the
// database survives app updates and doesn't sit inside the read-only
// installation folder.
const DB_DIR = process.env.SPICE_DATA_DIR || path.join(__dirname, 'data');
const DB_PATH = path.join(DB_DIR, 'config.db');
// Single-instance lockfile. sql.js loads the entire DB into memory and
// rewrites the whole file on every debounced commit, so two server
// processes sharing this file silently corrupt each other's writes —
// whichever flushes last wins, and previously-saved data from the loser
// vanishes. The lock makes that impossible: a second process refuses
// to boot. See acquireSingleInstanceLock() below.
const LOCK_PATH = path.join(DB_DIR, 'server.lock');

let SQL = null;        // sql.js module instance (loaded once)
let rawDb = null;       // sql.js Database instance
let wrapped = null;     // our API wrapper
let pendingSave = null; // debounced fs.writeFile timer
let lockOwned = false;  // true once we've successfully acquired LOCK_PATH
let currentActor = '';  // username stamped into modified_by by table triggers

/**
 * Set the username attributed to subsequent writes via the modified_by
 * stamping triggers. server.js's requireAuth calls this per request so
 * every INSERT/UPDATE records WHO made it without touching any call site.
 * Background / migration writes run with '' (or 'system' during boot).
 * sql.js executes synchronously, so within one request's run of writes
 * the actor can't be clobbered by another request mid-statement.
 */
function setActor(name) {
  currentActor = (name == null) ? '' : String(name);
}

/**
 * Register per-connection SQL functions. sql.js functions live on the
 * Database instance, not in the file, so this MUST be re-run after every
 * open and after replaceDbFromBuffer (restore) — otherwise the stamping
 * triggers below would fail with "no such function: current_actor".
 */
function registerDbFunctions(database) {
  try { database.create_function('current_actor', () => currentActor); } catch (_) {}
}

/**
 * Serialize the live DB to a Buffer for writing to disk.
 *
 * IMPORTANT: sql.js's export() resets the connection and DROPS every
 * user-defined function registered via create_function(). If we don't
 * re-register immediately, the very next write that fires a modified_by
 * stamping trigger fails with "no such function: current_actor". Because
 * saves are debounced, this surfaced intermittently: write → save (200ms
 * later) wipes the function → next write throws. Re-registering here keeps
 * current_actor() alive across every save.
 */
function exportDb() {
  const buf = Buffer.from(rawDb.export());
  registerDbFunctions(rawDb);
  return buf;
}

/**
 * One-time data-migration ledger. Some fixes (e.g. retagging invoice
 * state) must run exactly ONCE per database, never on every boot — an
 * unguarded re-run is exactly what silently re-flipped imported ASP
 * invoices to ISP. These helpers gate such fixes on a marker row.
 */
function dataMigrationDone(db, name) {
  db.exec(`CREATE TABLE IF NOT EXISTS schema_meta (key TEXT PRIMARY KEY, value TEXT)`);
  return !!db.get('SELECT value FROM schema_meta WHERE key = ?', [name]);
}
function markDataMigration(db, name) {
  db.run('INSERT OR REPLACE INTO schema_meta (key, value) VALUES (?, ?)',
    [name, new Date().toISOString()]);
}

/**
 * Persist the in-memory DB to disk. Debounced 200ms so a burst of writes
 * (e.g. invoice generation) only triggers one write.
 */
function scheduleSave() {
  if (pendingSave) clearTimeout(pendingSave);
  pendingSave = setTimeout(() => {
    pendingSave = null;
    if (!rawDb) return;
    try {
      const buf = exportDb();
      const tmp = DB_PATH + '.tmp';
      fs.writeFileSync(tmp, buf);
      fs.renameSync(tmp, DB_PATH);
    } catch (e) {
      console.error('[db] save failed:', e.message);
    }
  }, 200);
}

/**
 * Acquire the single-instance lock. Throws if another live server
 * already owns it. Stale locks (PID no longer running) are reclaimed.
 *
 * Race-safe: writes the lockfile with `flag: 'wx'` (O_EXCL), so two
 * simultaneous boots can't both think they won. The loser sees EEXIST,
 * reads the existing PID, checks if it's alive, and either reclaims a
 * stale lock or refuses to boot.
 *
 * On macOS/Linux, `process.kill(pid, 0)` is the standard "is this PID
 * alive?" probe — sends signal 0, which the kernel uses to validate
 * the target without delivering anything. Throws ESRCH for dead PIDs,
 * EPERM for live-but-not-ours PIDs (also counts as "alive" — someone
 * else's process is using the DB).
 */
function acquireSingleInstanceLock() {
  const dir = path.dirname(LOCK_PATH);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  const payload = JSON.stringify({
    pid: process.pid,
    bootAt: new Date().toISOString(),
    host: require('os').hostname(),
    cwd: process.cwd(),
  });
  // Try to write atomically; if it exists, inspect the holder.
  try {
    fs.writeFileSync(LOCK_PATH, payload, { flag: 'wx' });
    lockOwned = true;
    return;
  } catch (e) {
    if (e.code !== 'EEXIST') throw e;
  }
  // Lock exists — read it and decide whether to reclaim or refuse.
  let holder = null;
  try { holder = JSON.parse(fs.readFileSync(LOCK_PATH, 'utf8')); }
  catch (_) { /* corrupt lockfile — treat as stale */ }
  const holderPid = holder && Number(holder.pid);
  const stale = (() => {
    if (!holderPid) return true;             // unreadable / no pid → stale
    if (holderPid === process.pid) return true; // our own previous boot
    try {
      process.kill(holderPid, 0);
      return false;                           // pid is alive — not stale
    } catch (err) {
      return err.code === 'ESRCH';            // dead → reclaim; EPERM → assume alive
    }
  })();
  if (!stale) {
    const msg = [
      `Another server is already running and owns ${LOCK_PATH}.`,
      `  pid:     ${holderPid}`,
      `  bootAt:  ${(holder && holder.bootAt) || 'unknown'}`,
      `  host:    ${(holder && holder.host)   || 'unknown'}`,
      ``,
      `Stop that process first, OR if it's a zombie (handler swallowed SIGTERM),`,
      `force-kill it:   kill -9 ${holderPid}`,
      `then retry. If you're sure no server is running, delete the lock:`,
      `   rm ${LOCK_PATH}`,
    ].join('\n');
    const err = new Error(msg);
    err.code = 'EALREADY';
    throw err;
  }
  // Stale lock — reclaim by overwriting.
  fs.writeFileSync(LOCK_PATH, payload);
  lockOwned = true;
  console.log('[db] reclaimed stale lock from pid', holderPid);
}

/**
 * Release the single-instance lock. Idempotent; called from flushSave.
 * Only removes the lock if WE own it (avoids deleting another process's
 * lock if this one was reclaimed away from us mid-flight).
 */
function releaseSingleInstanceLock() {
  if (!lockOwned) return;
  try {
    const holder = JSON.parse(fs.readFileSync(LOCK_PATH, 'utf8'));
    if (holder && Number(holder.pid) === process.pid) {
      fs.unlinkSync(LOCK_PATH);
    }
  } catch (_) { /* lock already gone or unreadable — fine */ }
  lockOwned = false;
}

/**
 * Force-flush any pending save synchronously. Called on shutdown/close.
 * Also releases the single-instance lock so a clean exit always leaves
 * the lockfile gone — next boot doesn't have to reclaim a stale lock.
 */
function flushSave() {
  if (pendingSave) { clearTimeout(pendingSave); pendingSave = null; }
  if (rawDb) {
    try {
      const buf = exportDb();
      const tmp = DB_PATH + '.tmp';
      fs.writeFileSync(tmp, buf);
      fs.renameSync(tmp, DB_PATH);
    } catch (e) {
      console.error('[db] flush failed:', e.message);
    }
  }
  // Always release the lock, even if flush failed — holding a stale
  // lock after exit is worse than risking a partial save (which the
  // user can investigate via the .tmp file or the data/config.db.* backups).
  releaseSingleInstanceLock();
}

/**
 * Initialize the database. async/await is necessary because sql.js loads
 * its WASM module asynchronously.
 */
async function initDb() {
  if (wrapped) return wrapped;

  const dir = path.dirname(DB_PATH);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

  // Refuse to boot if another live server is already managing this DB.
  // sql.js's "load whole file into memory, rewrite on commit" model
  // means two processes silently overwrite each other's writes.
  // Throws with a clear EALREADY error that server.js's top-level
  // catch can format and exit cleanly. Stale locks from dead PIDs
  // are reclaimed automatically.
  acquireSingleInstanceLock();

  // Load sql.js wasm runtime once
  if (!SQL) SQL = await initSqlJs();

  // Open existing DB or create empty one
  if (fs.existsSync(DB_PATH)) {
    const buf = fs.readFileSync(DB_PATH);
    rawDb = new SQL.Database(buf);
  } else {
    rawDb = new SQL.Database();
  }

  // Enable foreign keys
  rawDb.run("PRAGMA foreign_keys = ON;");
  // Register current_actor() before any write so the stamping triggers
  // (created in the migrations section) can resolve it. Boot-time and
  // migration writes are attributed to 'system' until the first
  // authenticated request calls setActor().
  registerDbFunctions(rawDb);
  setActor('system');

  wrapped = makeWrapper();

  // Save on process exit (best-effort)
  const onExit = () => { flushSave(); };
  process.on('SIGINT', onExit);
  process.on('SIGTERM', onExit);
  process.on('beforeExit', onExit);

  // ── LICENSE STATE ──────────────────────────────────────────
  // Single-row table (CHECK id = 1) holding the per-install license
  // state. On first boot, ./license.js inserts a row with a fresh
  // install_id and a 30-day trial expiry. The dev's signed tokens
  // bump expires_at when applied via /api/license/apply.
  //
  // active_token stores the most recently applied token verbatim so
  // an operator can copy it back out if they need to re-apply on a
  // restored backup, and so the dev can audit who has what.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS license_state (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    install_id TEXT NOT NULL,
    first_seen_at TEXT NOT NULL,
    expires_at TEXT NOT NULL,
    active_token TEXT
  )`);

  // ── SESSIONS ───────────────────────────────────────────────
  // expires_at caps the lifetime of a leaked Authorization header even
  // if the holder keeps it warm via the sliding last_used_at sweep.
  // Pre-migration rows have NULL → grandfathered in (relies on the
  // 30-day idle sweep); new rows get a 30-day cap from creation.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS sessions (
    token TEXT PRIMARY KEY,
    user_id INTEGER NOT NULL,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    last_used_at TEXT DEFAULT (datetime('now','localtime')),
    expires_at TEXT,
    device_label TEXT DEFAULT '',
    FOREIGN KEY(user_id) REFERENCES users(id)
  )`);

  // ── USERS ──────────────────────────────────────────────────
  // must_change_password is the force-rotate gate. When non-zero, the
  // server's requireAuth blocks every endpoint except whoami / change-
  // password until the user picks a new password — closes the default-
  // creds attack window for both the seeded admin and any admin-reset
  // user. See FORCED_CHANGE_ALLOWED in server.js for the allowlist.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'user',
    token TEXT,
    must_change_password INTEGER NOT NULL DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── TRADERS (NAM.DBF — sellers/poolers) ────────────────────
  // whatsapp / email added for the mobile PWA's seller create/edit flow.
  // The desktop sellers tab also exposes them via the unified write
  // path in mobile-bridge.js (POST/PUT /api/traders).
  wrapped.exec(`CREATE TABLE IF NOT EXISTS traders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    cr TEXT DEFAULT '',
    pan TEXT DEFAULT '',
    tel TEXT DEFAULT '',
    aadhar TEXT DEFAULT '',
    padd TEXT DEFAULT '',
    ppla TEXT DEFAULT '',
    pin TEXT DEFAULT '',
    pstate TEXT DEFAULT '',
    pst_code TEXT DEFAULT '',
    ifsc TEXT DEFAULT '',
    acctnum TEXT DEFAULT '',
    holder_name TEXT DEFAULT '',
    whatsapp TEXT DEFAULT '',
    email TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── TRADER BANKS ───────────────────────────────────────────
  wrapped.exec(`CREATE TABLE IF NOT EXISTS trader_banks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    trader_id INTEGER NOT NULL,
    bank_name TEXT DEFAULT '',
    branch TEXT DEFAULT '',
    acctnum TEXT NOT NULL,
    ifsc TEXT NOT NULL,
    holder_name TEXT DEFAULT '',
    is_default INTEGER DEFAULT 0,
    FOREIGN KEY (trader_id) REFERENCES traders(id)
  )`);

  // ── BUYERS (SBL.DBF — buyers/dealers/traders) ──────────────
  wrapped.exec(`CREATE TABLE IF NOT EXISTS buyers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    buyer TEXT NOT NULL,
    buyer1 TEXT DEFAULT '',
    code TEXT DEFAULT '',
    sbl TEXT DEFAULT '',
    add1 TEXT DEFAULT '',
    add2 TEXT DEFAULT '',
    pla TEXT DEFAULT '',
    pin TEXT DEFAULT '',
    state TEXT DEFAULT '',
    st_code TEXT DEFAULT '',
    gstin TEXT DEFAULT '',
    pan TEXT DEFAULT '',
    tel TEXT DEFAULT '',
    ti TEXT DEFAULT '',
    sale TEXT DEFAULT 'L',
    email TEXT DEFAULT '',
    tdsq TEXT DEFAULT '',
    cbuyer1 TEXT DEFAULT '',
    cadd1 TEXT DEFAULT '',
    cadd2 TEXT DEFAULT '',
    cpla TEXT DEFAULT '',
    cpin TEXT DEFAULT '',
    cstate TEXT DEFAULT '',
    cst_code TEXT DEFAULT '',
    cgstin TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── AUCTIONS (trade sessions) ──────────────────────────────
  // Tri-state price-check gate:
  //   price_check_first_passed_at — stamped on the FIRST successful
  //     verify and never cleared. Tells us the operator has reconciled
  //     this auction at least once (separates 'never' from 'stale').
  //   price_checked_at — stamped on every successful verify, cleared
  //     by any endpoint that mutates a lot's price/code. Tells us the
  //     reconciliation is still current.
  // The pair drives the gate states ('off' | 'never' | 'stale' | 'clean')
  // used by the calc / invoice / purchase / bill / debit-note generators.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS auctions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ano TEXT NOT NULL,
    date TEXT NOT NULL,
    crop_type TEXT DEFAULT 'ASP',
    state TEXT DEFAULT 'TAMIL NADU',
    start_time TEXT,
    end_time TEXT,
    price_checked_at TEXT DEFAULT '',
    price_check_first_passed_at TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── GENERATION OVERRIDES ──────────────────────────────────
  // One-shot admin grants that allow a single regeneration after the
  // doc-type's first generate has already happened. Consumed (deleted)
  // by the generate endpoint the next time it runs successfully for the
  // (auction_id, doc_type) pair. Without a row, generation is allowed
  // by default; the row only exists when there's an active allowance.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS generation_overrides (
    auction_id INTEGER NOT NULL,
    doc_type   TEXT NOT NULL,
    granted_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
    granted_by TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (auction_id, doc_type)
  )`);

  // ── LOTS (CPA1.DBF — main lot data, before + after trade) ─
  wrapped.exec(`CREATE TABLE IF NOT EXISTS lots (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    auction_id INTEGER NOT NULL,
    lot_no TEXT NOT NULL,
    crop TEXT DEFAULT '',
    grade TEXT DEFAULT '',
    crpt TEXT DEFAULT '',
    branch TEXT DEFAULT '',
    state TEXT DEFAULT 'TAMIL NADU',
    trader_id INTEGER,
    name TEXT DEFAULT '',
    padd TEXT DEFAULT '',
    ppla TEXT DEFAULT '',
    ppin TEXT DEFAULT '',
    pstate TEXT DEFAULT '',
    pst_code TEXT DEFAULT '',
    cr TEXT DEFAULT '',
    pan TEXT DEFAULT '',
    tel TEXT DEFAULT '',
    aadhar TEXT DEFAULT '',
    bags INTEGER DEFAULT 0,
    litre TEXT DEFAULT '',
    qty REAL DEFAULT 0,
    gross_wt REAL DEFAULT 0,
    sample_wt REAL DEFAULT 0,
    moisture TEXT DEFAULT '',
    reserved_price REAL DEFAULT 0,
    price REAL DEFAULT 0,
    amount REAL DEFAULT 0,
    code TEXT DEFAULT '',
    buyer TEXT DEFAULT '',
    buyer1 TEXT DEFAULT '',
    sale TEXT DEFAULT '',
    invo TEXT DEFAULT '',
    pqty REAL DEFAULT 0,
    prate REAL DEFAULT 0,
    puramt REAL DEFAULT 0,
    com REAL DEFAULT 0,
    sertax REAL DEFAULT 0,
    cgst REAL DEFAULT 0,
    sgst REAL DEFAULT 0,
    igst REAL DEFAULT 0,
    dcgst REAL DEFAULT 0,
    dsgst REAL DEFAULT 0,
    digst REAL DEFAULT 0,
    refud REAL DEFAULT 0,
    refund REAL DEFAULT 0,
    advance REAL DEFAULT 0,
    balance REAL DEFAULT 0,
    bilamt REAL DEFAULT 0,
    paid TEXT DEFAULT '',
    user_id TEXT DEFAULT '',
    bank_id INTEGER,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (auction_id) REFERENCES auctions(id),
    FOREIGN KEY (trader_id) REFERENCES traders(id)
  )`);

  // ── INVOICES (INV.DBF — sales invoices) ────────────────────
  // addl_chg / addl_name = optional "Additional Charge" row that sits
  // below Round on/off — sum(cardamom) × cfg.addl_charge_value with a
  // user-defined ledger label (also used as the Tally ledger name in XML).
  // lorry_no = per-invoice vehicle number for the e-way bill
  // <VEHICLENUMBER> field. Set via the Sales tab's bulk-action button.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS invoices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    auction_id INTEGER,
    ano TEXT NOT NULL,
    date TEXT NOT NULL,
    state TEXT DEFAULT '',
    sale TEXT DEFAULT 'L',
    invo TEXT NOT NULL,
    buyer TEXT DEFAULT '',
    buyer1 TEXT DEFAULT '',
    gstin TEXT DEFAULT '',
    place TEXT DEFAULT '',
    lot TEXT DEFAULT '',
    -- asp_invo: the sister-company (ASP, Kerala) invoice number paired
    -- with this ISP (Tamil Nadu) invoice for the same trade + buyer.
    -- Populated for IMPORTED invoices by the sales-invoice import's
    -- ASP↔ISP linkage pass (the generated flow keeps the link on
    -- lots.asp_invo instead). Surfaced as the "ASP Inv#" sales column.
    asp_invo TEXT DEFAULT '',
    bag INTEGER DEFAULT 0,
    qty REAL DEFAULT 0,
    price REAL DEFAULT 0,
    amount REAL DEFAULT 0,
    gunny REAL DEFAULT 0,
    pava_hc REAL DEFAULT 0,
    ins REAL DEFAULT 0,
    cgst REAL DEFAULT 0,
    sgst REAL DEFAULT 0,
    igst REAL DEFAULT 0,
    tcs REAL DEFAULT 0,
    rund REAL DEFAULT 0,
    tot REAL DEFAULT 0,
    addl_chg REAL DEFAULT 0,
    addl_name TEXT DEFAULT '',
    lorry_no TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── PURCHASES (PURCHASE.DBF — purchase invoices for registered dealers)
  wrapped.exec(`CREATE TABLE IF NOT EXISTS purchases (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    auction_id INTEGER,
    ano TEXT NOT NULL,
    date TEXT NOT NULL,
    state TEXT DEFAULT '',
    br TEXT DEFAULT '',
    name TEXT DEFAULT '',
    add_line TEXT DEFAULT '',
    place TEXT DEFAULT '',
    gstin TEXT DEFAULT '',
    invo TEXT DEFAULT '',
    qty REAL DEFAULT 0,
    amount REAL DEFAULT 0,
    cgst REAL DEFAULT 0,
    sgst REAL DEFAULT 0,
    igst REAL DEFAULT 0,
    rund REAL DEFAULT 0,
    total REAL DEFAULT 0,
    tds REAL DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── BILLS (BILL.DBF — bills of supply for unregistered/agriculturist)
  wrapped.exec(`CREATE TABLE IF NOT EXISTS bills (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ano TEXT NOT NULL,
    date TEXT NOT NULL,
    state TEXT DEFAULT '',
    br TEXT DEFAULT '',
    crpt TEXT DEFAULT '',
    bil INTEGER DEFAULT 0,
    name TEXT DEFAULT '',
    add_line TEXT DEFAULT '',
    pla TEXT DEFAULT '',
    pstate TEXT DEFAULT '',
    st_code TEXT DEFAULT '',
    crr TEXT DEFAULT '',
    pan TEXT DEFAULT '',
    qty REAL DEFAULT 0,
    cost REAL DEFAULT 0,
    igst REAL DEFAULT 0,
    net REAL DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── DEBIT NOTES ────────────────────────────────────────────
  wrapped.exec(`CREATE TABLE IF NOT EXISTS debit_notes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ano TEXT NOT NULL,
    date TEXT NOT NULL,
    state TEXT DEFAULT '',
    name TEXT DEFAULT '',
    note_no TEXT DEFAULT '',
    amount REAL DEFAULT 0,
    cgst REAL DEFAULT 0,
    sgst REAL DEFAULT 0,
    igst REAL DEFAULT 0,
    total REAL DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── AUDIT LOG ──────────────────────────────────────────────
  wrapped.exec(`CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id TEXT NOT NULL,
    action TEXT NOT NULL,
    entity TEXT NOT NULL,
    entity_id INTEGER,
    details TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── DELETE LOG (Delete All audit trail) ────────────────────
  // Records every Delete All action so the operator can see WHO wiped
  // WHICH table WHEN, with row counts + the on-disk backup path so a
  // misclick is recoverable via Backup → Restore from File.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS delete_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    username TEXT,
    resource TEXT NOT NULL,
    deleted_count INTEGER DEFAULT 0,
    cascade_counts TEXT,
    backup_path TEXT,
    ip TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── LOGIN HISTORY (mobile + desktop sign-in audit) ────────
  // One row per successful login from either app. Captures the IP and
  // a coarse device-type tag ('Mobile' vs 'Desktop') derived from the
  // User-Agent — finer-grained device labelling stays on sessions.
  // Used by the admin console to spot anomalous access and by the
  // mobile-bridge's /api/auth/login flow.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS login_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    username TEXT NOT NULL,
    ip TEXT DEFAULT '',
    user_agent TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── IMPORT LOG (Import Old Data audit trail) ──────────────
  // One row per upload (preview or run, dry-run or live) from the
  // Import Old Data tool. inserted_ids holds the JSON-encoded list of
  // newly-inserted row PKs so the History panel's Undo button can roll
  // back a specific import; undone_at is set when undo runs so the
  // button stays disabled on a second click.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS import_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    module TEXT NOT NULL,
    filename TEXT DEFAULT '',
    dry_run INTEGER DEFAULT 0,
    total INTEGER DEFAULT 0,
    imported INTEGER DEFAULT 0,
    skipped INTEGER DEFAULT 0,
    failed INTEGER DEFAULT 0,
    errors TEXT DEFAULT '',
    inserted_ids TEXT DEFAULT '',
    undone_at TEXT DEFAULT '',
    user_id INTEGER,
    username TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── LOT ALLOCATIONS ───────────────────────────────────────
  // Per-auction, per-branch lot-number ranges. Allocations let the user
  // reserve "lots 001-050 for BODI, 051-100 for VANDANMEDU" on a given
  // trade, so each branch's auction-floor operator can enter lots in
  // parallel without lot-number collisions. Ranges are stored as text
  // because lot numbers can be alphanumeric ("A001", "001A", etc.) —
  // the application walks the numeric tail of each end to enumerate.
  //
  // Bulk-replace semantics: the POST endpoint wipes and re-inserts for
  // a given auction, so the allocation set is always consistent. The
  // server refuses to drop a range that still contains saved lots —
  // forces the operator to delete the lots first if they really mean it.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS lot_allocations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    auction_id INTEGER NOT NULL,
    branch TEXT NOT NULL,
    start_lot TEXT NOT NULL,
    end_lot TEXT NOT NULL,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (auction_id) REFERENCES auctions(id)
  )`);
  wrapped.exec(`CREATE INDEX IF NOT EXISTS idx_lot_allocations_auction ON lot_allocations(auction_id)`);

  // ── ROUTE DISTANCES ───────────────────────────────────────
  // Maps (dispatch PIN, consignee PIN) → road km, populated manually by
  // the user via the To Tally → E-way Bill Distance UI. The user looks
  // up a route once on NIC's portal, saves it here, and every future
  // invoice between the same two PINs gets the value automatically.
  //
  // Keys are normalised: we always store the lexicographically smaller
  // PIN as `from_pin` so A→B and B→A share a single row.
  wrapped.exec(`CREATE TABLE IF NOT EXISTS route_distances (
    from_pin TEXT NOT NULL,
    to_pin   TEXT NOT NULL,
    km       INTEGER NOT NULL,
    updated_at TEXT DEFAULT (datetime('now','localtime')),
    PRIMARY KEY (from_pin, to_pin)
  )`);

  // ── INDEXES ────────────────────────────────────────────────
  const indexes = [
    'CREATE INDEX IF NOT EXISTS idx_traders_name ON traders(name)',
    'CREATE INDEX IF NOT EXISTS idx_lots_auction ON lots(auction_id)',
    'CREATE INDEX IF NOT EXISTS idx_lots_lot ON lots(lot_no)',
    'CREATE INDEX IF NOT EXISTS idx_lots_name ON lots(name)',
    'CREATE INDEX IF NOT EXISTS idx_lots_buyer ON lots(buyer)',
    'CREATE INDEX IF NOT EXISTS idx_lots_sale ON lots(sale)',
    'CREATE INDEX IF NOT EXISTS idx_invoices_date ON invoices(date)',
    'CREATE INDEX IF NOT EXISTS idx_invoices_sale ON invoices(sale, invo)',
    'CREATE INDEX IF NOT EXISTS idx_purchases_date ON purchases(date)',
    'CREATE INDEX IF NOT EXISTS idx_purchases_name ON purchases(name)',
    'CREATE INDEX IF NOT EXISTS idx_bills_date ON bills(date)',
    'CREATE INDEX IF NOT EXISTS idx_bills_name ON bills(name)',
    'CREATE INDEX IF NOT EXISTS idx_buyers_buyer ON buyers(buyer)',
    'CREATE INDEX IF NOT EXISTS idx_buyers_buyer1 ON buyers(buyer1)',
    // Surfaces locked rows for the Lots tab's filter/sort without
    // scanning every lot. The auction_id index already covers cascade-
    // lock lookups (which are scoped by auction_id), so this index is
    // narrow and only pays off for "show locked rows across the DB."
    'CREATE INDEX IF NOT EXISTS idx_lots_locked ON lots(locked_at)',
    // FK child-side index. SQLite auto-indexes the parent (traders.id
    // is PK) but NOT the child column, so DELETE FROM traders triggers
    // a full scan of trader_banks per row to check for orphans. Without
    // this, bulk seller deletion is O(N·M) — quadratic.
    'CREATE INDEX IF NOT EXISTS idx_trader_banks_trader ON trader_banks(trader_id)',
  ];
  for (const idx of indexes) { try { wrapped.exec(idx); } catch (e) {} }

  // Business-data tables that carry modified_at / modified_by audit
  // columns and the stamping triggers. Deliberately EXCLUDES append-only
  // or system tables (sessions, license_state, audit_log, delete_log,
  // login_history, import_log, schema_meta) where "who last modified this
  // row" is either meaningless or already captured by the table itself.
  const AUDITED_TABLES = ['traders','trader_banks','buyers','auctions',
    'generation_overrides','lots','invoices','purchases','bills',
    'debit_notes','lot_allocations','route_distances','users'];

  // ── MIGRATIONS (for existing databases created before schema changes) ──
  const migrations = [
    'ALTER TABLE purchases ADD COLUMN auction_id INTEGER',
    'ALTER TABLE invoices ADD COLUMN auction_id INTEGER',
    'ALTER TABLE bills ADD COLUMN auction_id INTEGER',
    'ALTER TABLE debit_notes ADD COLUMN auction_id INTEGER',
    "ALTER TABLE buyers ADD COLUMN code TEXT DEFAULT ''",
    "ALTER TABLE buyers ADD COLUMN cadd2 TEXT DEFAULT ''",
    "ALTER TABLE buyers ADD COLUMN email TEXT DEFAULT ''",
    "ALTER TABLE buyers ADD COLUMN tdsq TEXT DEFAULT ''",
    "ALTER TABLE buyers ADD COLUMN sbl TEXT DEFAULT ''",
    // Discount GST columns (per-lot, when flag_disc_gst is ON)
    'ALTER TABLE lots ADD COLUMN dcgst REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN dsgst REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN digst REAL DEFAULT 0',
    // ASP invoice traceability — when a lot is first invoiced as an ASP
    // sale (state=Kerala), `lots.invo` gets the ASP invoice number AND a
    // copy is preserved here. Then when the same lot is invoiced as an
    // ISP sale (state=Tamil Nadu) later, `lots.invo` is overwritten with
    // the ISP invoice number, but `lots.asp_invo` keeps the original ASP
    // ref. This lets the sales list show both numbers side-by-side.
    "ALTER TABLE lots ADD COLUMN asp_invo TEXT DEFAULT ''",
    // ── Dual-view planter calculation columns ─────────────────
    // calculateLot() chooses ISP vs ASP rules based on cfg.business_state,
    // then writes the active view into pqty/prate/puramt. The Tally URD
    // voucher needs ISP values regardless of which mode dad is currently in,
    // so we now ALWAYS persist BOTH calculations on every save:
    //   isp_pqty/isp_prate/isp_puramt → planter side as ISP would compute
    //   asp_pqty/asp_prate/asp_puramt → planter side as ASP would compute
    // The legacy pqty/prate/puramt columns continue to mirror whichever
    // matches the current business_state, so the existing UI / reports /
    // exports keep working unchanged. Reports that need a specific view
    // (like the URD Tally voucher) read the prefixed columns directly.
    'ALTER TABLE lots ADD COLUMN isp_pqty REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN isp_prate REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN isp_puramt REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN asp_pqty REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN asp_prate REAL DEFAULT 0',
    'ALTER TABLE lots ADD COLUMN asp_puramt REAL DEFAULT 0',
    // Distance for e-way bill <DISTANCE> field on ISP sales vouchers.
    // Populated manually per-invoice from the To Tally → 🗺️ E-way Bill
    // Distance UI: user looks up the value on NIC's Pin-to-Pin Distance
    // Search page (or Google Maps), pastes it here, clicks Save. Value
    // is then emitted verbatim on the next voucher regen.
    'ALTER TABLE invoices ADD COLUMN distance_km INTEGER',
    // Drop legacy pincodes/pin_distances tables that supported the old
    // haversine auto-compute path. We replaced that with the manual-
    // override workflow (above), so these tables are now dead weight.
    // IF EXISTS makes this idempotent — fresh DBs have nothing to drop;
    // upgraded DBs shed the orphan tables on next restart.
    'DROP TABLE IF EXISTS pin_distances',
    'DROP TABLE IF EXISTS pincodes',
    // Business mode tag on auctions — stamped at trade creation from the
    // current company_settings.business_mode value (e-Trade or e-Auction)
    // and locked. Every downstream entity (lots, invoices, purchases,
    // bills, debit notes, payments) inherits via the FK to auctions, so
    // list endpoints filter by JOINing auctions on mode. Existing rows
    // stay NULL until manually re-imported; NULL rows pass any mode
    // filter so legacy data remains visible during a soft cutover.
    "ALTER TABLE auctions ADD COLUMN mode TEXT DEFAULT ''",
    // Lot locking — when a lot is finalised, lock it so non-admins
    // can no longer edit or delete it. `locked_at` doubles as the
    // boolean ("is locked?") and the timestamp; `locked_by` records
    // which user locked it for the audit log. Admin role bypasses
    // the gate. Used by the bulk Lock / Unlock buttons on the Lots
    // tab. NULL = unlocked (default).
    "ALTER TABLE lots ADD COLUMN locked_at TEXT",
    "ALTER TABLE lots ADD COLUMN locked_by TEXT",
    // Bank branch — auto-filled from the IFSC lookup (Razorpay public API)
    // in the seller edit modal, shown under the Bank Name field.
    "ALTER TABLE trader_banks ADD COLUMN branch TEXT DEFAULT ''",
    // Mobile-bridge canonical columns. The bridge's ensureBridgeSchema
    // self-heal will still attempt these idempotently as defence in
    // depth on installs that boot mobile-bridge before db.js migrations
    // settle, but they belong in the canonical schema now.
    "ALTER TABLE traders ADD COLUMN whatsapp TEXT DEFAULT ''",
    "ALTER TABLE traders ADD COLUMN email TEXT DEFAULT ''",
    "ALTER TABLE lots ADD COLUMN bank_id INTEGER",
    // Forced password change — gates seeded admin + admin-reset users
    // through the change-password screen on first login. See
    // FORCED_CHANGE_ALLOWED in server.js.
    "ALTER TABLE users ADD COLUMN must_change_password INTEGER NOT NULL DEFAULT 0",
    // Hard expiry cap on sessions. Pre-migration rows get NULL and
    // are grandfathered in (last_used_at sliding sweep still applies);
    // new rows get a 30-day cap at INSERT time in server.js so a
    // leaked Authorization header has a bounded lifetime even if
    // the holder keeps it warm.
    "ALTER TABLE sessions ADD COLUMN expires_at TEXT",
    // Price-check gate timestamps (see auctions CREATE TABLE for the
    // tri-state semantics). Backfill below stamps first_passed_at on
    // any auction that was already verified before the column existed
    // so the gate doesn't drop to 'never' for previously-verified data.
    "ALTER TABLE auctions ADD COLUMN price_checked_at TEXT DEFAULT ''",
    "ALTER TABLE auctions ADD COLUMN price_check_first_passed_at TEXT DEFAULT ''",
    // Additional charge row + per-invoice lorry number. See invoices
    // CREATE TABLE comment for what they carry.
    "ALTER TABLE invoices ADD COLUMN addl_chg REAL DEFAULT 0",
    "ALTER TABLE invoices ADD COLUMN addl_name TEXT DEFAULT ''",
    "ALTER TABLE invoices ADD COLUMN lorry_no TEXT",
    // e-Auction Reserved Price — bid floor per lot, typed by the
    // operator at entry time. Persisted regardless of the
    // flag_reserved_price toggle (so flipping the flag later doesn't
    // lose values), but the UI hides the input when the flag is off.
    // Feeds the Spices Board e-Auction CSV's reserved-price column on
    // export. Stored as REAL — value is in rupees per kg.
    'ALTER TABLE lots ADD COLUMN reserved_price REAL DEFAULT 0',
    // Sister-company (ASP/Kerala) invoice number paired with an ISP
    // (Tamil Nadu) invoice. Populated for IMPORTED sales invoices by the
    // ASP↔ISP linkage pass (matched on trade + buyer). Generated invoices
    // keep the link on lots.asp_invo. Surfaced as the "ASP Inv#" column.
    "ALTER TABLE invoices ADD COLUMN asp_invo TEXT DEFAULT ''",
    // ── modified_at / modified_by audit columns ──────────────────
    // WHEN a row last changed and WHO changed it. Stamped automatically
    // by the AFTER INSERT/UPDATE triggers created below (no call site is
    // touched). A constant '' default is required because SQLite refuses
    // ADD COLUMN with a non-constant default like CURRENT_TIMESTAMP —
    // the trigger fills the real timestamp on the very next write.
    ...AUDITED_TABLES.flatMap(t => [
      `ALTER TABLE ${t} ADD COLUMN modified_at TEXT DEFAULT ''`,
      `ALTER TABLE ${t} ADD COLUMN modified_by TEXT DEFAULT ''`,
    ]),
  ];
  for (const m of migrations) {
    try { wrapped.exec(m); console.log('Migration applied:', m); }
    catch (e) { /* column already exists — ignore */ }
  }

  // ── modified_at / modified_by stamping triggers ────────────────
  // Every write to an audited table records WHEN it changed and WHO
  // changed it, without any application call site having to pass the
  // values. modified_at is a localtime timestamp; modified_by is
  // current_actor() — the username set by setActor() for the in-flight
  // request, or 'system'/'' for background work. The trigger's own
  // UPDATE does NOT recurse: PRAGMA recursive_triggers is OFF by default
  // (verified at boot below), so an AFTER UPDATE trigger that updates its
  // own row will not re-fire itself.
  for (const t of AUDITED_TABLES) {
    try {
      wrapped.exec(`CREATE TRIGGER IF NOT EXISTS trg_${t}_modins AFTER INSERT ON ${t}
        BEGIN UPDATE ${t} SET modified_at = datetime('now','localtime'), modified_by = current_actor() WHERE rowid = NEW.rowid; END;`);
      wrapped.exec(`CREATE TRIGGER IF NOT EXISTS trg_${t}_modupd AFTER UPDATE ON ${t}
        BEGIN UPDATE ${t} SET modified_at = datetime('now','localtime'), modified_by = current_actor() WHERE rowid = NEW.rowid; END;`);
    } catch (e) { /* table may not exist on a partial schema — ignore */ }
  }

  // Price-check backfill: any auction that was already verified BEFORE
  // the price_check_first_passed_at column existed gets its first-pass
  // stamp set to the current-verify stamp. Without this, every
  // previously-verified auction would re-enter the 'never' state on
  // upgrade and force a one-off re-verify. Idempotent — only touches
  // rows where the first-pass stamp is missing.
  try {
    wrapped.run(
      `UPDATE auctions
          SET price_check_first_passed_at = price_checked_at
        WHERE price_checked_at IS NOT NULL
          AND price_checked_at != ''
          AND (price_check_first_passed_at IS NULL OR price_check_first_passed_at = '')`
    );
  } catch (_) { /* table or columns may not exist on fresh DB */ }

  // One-time data fix: legacy ASP-only lots (where invo==asp_invo) had their
  // `sale` field set during the old ASP-generation logic. The current logic
  // doesn't set it (so ISP can pick the right sale type per buyer). Clear
  // those legacy rows so they show up in ISP eligibility.
  // Idempotent: subsequent runs do nothing because the rows are already
  // cleared. Safe to run on a fresh DB (no rows match the WHERE).
  try {
    const fix = wrapped.run(
      `UPDATE lots SET sale = ''
       WHERE asp_invo IS NOT NULL AND asp_invo != ''
         AND invo = asp_invo
         AND sale IS NOT NULL AND sale != ''`
    );
    if (fix && fix.changes > 0) {
      console.log(`Migration: cleared sale on ${fix.changes} ASP-only lots so ISP eligibility works`);
    }
  } catch (e) { /* ignore — column may not exist on first run */ }

  // Data fix: lots entered from the mobile PWA were saved with only
  // trader_id — the denormalised seller fields (name/place/CR/PAN…) were
  // left blank, so those lots appeared nameless on the desktop Lot Entry
  // screen and on slips/invoices that read these columns off the lots row.
  // Backfill any blank field from the linked trader. Idempotent — only
  // touches columns that are still empty on rows that have a trader_id, so
  // it's a no-op once filled and safe on a fresh DB (no matching rows).
  // (The POST /api/lots handler now backfills on insert too, for new lots.)
  try {
    const seedFix = wrapped.run(
      `UPDATE lots
          SET name     = CASE WHEN COALESCE(name,'')     = '' THEN (SELECT t.name     FROM traders t WHERE t.id = lots.trader_id) ELSE name     END,
              padd     = CASE WHEN COALESCE(padd,'')     = '' THEN (SELECT t.padd     FROM traders t WHERE t.id = lots.trader_id) ELSE padd     END,
              ppla     = CASE WHEN COALESCE(ppla,'')     = '' THEN (SELECT t.ppla     FROM traders t WHERE t.id = lots.trader_id) ELSE ppla     END,
              ppin     = CASE WHEN COALESCE(ppin,'')     = '' THEN (SELECT t.pin      FROM traders t WHERE t.id = lots.trader_id) ELSE ppin     END,
              pstate   = CASE WHEN COALESCE(pstate,'')   = '' THEN (SELECT t.pstate   FROM traders t WHERE t.id = lots.trader_id) ELSE pstate   END,
              pst_code = CASE WHEN COALESCE(pst_code,'') = '' THEN (SELECT t.pst_code FROM traders t WHERE t.id = lots.trader_id) ELSE pst_code END,
              cr       = CASE WHEN COALESCE(cr,'')       = '' THEN (SELECT t.cr       FROM traders t WHERE t.id = lots.trader_id) ELSE cr       END,
              pan      = CASE WHEN COALESCE(pan,'')      = '' THEN (SELECT t.pan      FROM traders t WHERE t.id = lots.trader_id) ELSE pan      END,
              tel      = CASE WHEN COALESCE(tel,'')      = '' THEN (SELECT t.tel      FROM traders t WHERE t.id = lots.trader_id) ELSE tel      END,
              aadhar   = CASE WHEN COALESCE(aadhar,'')   = '' THEN (SELECT t.aadhar   FROM traders t WHERE t.id = lots.trader_id) ELSE aadhar   END
        WHERE trader_id IS NOT NULL
          AND EXISTS (SELECT 1 FROM traders t WHERE t.id = lots.trader_id)
          AND (COALESCE(name,'') = '' OR COALESCE(ppla,'') = '' OR COALESCE(cr,'') = '' OR COALESCE(pan,'') = '')`
    );
    if (seedFix && seedFix.changes > 0) {
      console.log(`Migration: backfilled seller fields on ${seedFix.changes} mobile-entered lots from their trader record`);
    }
  } catch (_) { /* tables/columns may not exist on a partial schema — ignore */ }

  // One-time data fix for the sales-invoice ISP/ASP book split.
  // `invoices.state` records which company book a row belongs to:
  // KERALA = ASP (Amazing Spice Park), TAMIL NADU = ISP (Ideal Spices).
  // This corrects two earlier problems:
  //   1. Legacy rows were stamped with the auction's PHYSICAL state, not
  //      the company-book state.
  //   2. A previous version of this fix re-ran on EVERY boot and, having
  //      no ASP signal for IMPORTED invoices (which carry no lot lineage),
  //      silently re-stamped them all to TAMIL NADU — so ASP imports
  //      drifted into ISP after a restart. That is the bug this replaces.
  //
  // It now runs EXACTLY ONCE (guarded by schema_meta) and classifies each
  // invoice by the SAME authoritative signals the importer uses, in order:
  //   • PLACE prefixed "ASP"                                  → KERALA
  //   • a lot in the same auction+buyer whose asp_invo == invo → KERALA
  //   • otherwise                                             → TAMIL NADU
  // Because it is PLACE-aware, the one-time run also RECOVERS rows the
  // buggy version already flipped (their PLACE still starts with "ASP").
  // Non-destructive: only writes when the computed book differs from the
  // stored one.
  if (!dataMigrationDone(wrapped, 'retag_invoice_state_v2')) {
    try {
      const allInvs = wrapped.all('SELECT id, auction_id, buyer, invo, place, state FROM invoices');
      let aspCount = 0, ispCount = 0, changed = 0;
      for (const inv of allInvs) {
        let isASP = String(inv.place || '').trim().toUpperCase().startsWith('ASP');
        if (!isASP) {
          // Generated ASP invoices carry no ASP-prefixed place, but their
          // lots keep the ASP invoice number in asp_invo as lineage.
          const lineage = wrapped.get(
            `SELECT 1 FROM lots
              WHERE auction_id = ? AND buyer = ? AND asp_invo = ? LIMIT 1`,
            [inv.auction_id, inv.buyer, inv.invo]
          );
          isASP = !!lineage;
        }
        const newState = isASP ? 'KERALA' : 'TAMIL NADU';
        if (String(inv.state || '').toUpperCase() !== newState) {
          wrapped.run('UPDATE invoices SET state = ? WHERE id = ?', [newState, inv.id]);
          changed++;
        }
        if (isASP) aspCount++; else ispCount++;
      }
      if (aspCount + ispCount > 0) {
        console.log(`Migration: classified ${aspCount} invoices as KERALA (ASP) and ${ispCount} as TAMIL NADU (ISP) by PLACE + lot lineage (${changed} re-stamped)`);
      }
      markDataMigration(wrapped, 'retag_invoice_state_v2');
    } catch (e) { /* table may not exist on fresh DB — ignore */ }
  }

  const row = wrapped.get('SELECT COUNT(*) as cnt FROM users');
  if (!row || row.cnt === 0) {
    // bcrypt cost 12 → ~250ms hash, only paid on first run. The seeded
    // admin is forced through the change-password screen on first sign-in
    // (must_change_password=1) so the default 'admin123' can't survive
    // into production.
    const hash = bcrypt.hashSync('admin123', 12);
    wrapped.run(
      'INSERT INTO users (username, password_hash, role, must_change_password) VALUES (?, ?, ?, 1)',
      ['admin', hash, 'admin']
    );
    console.log('Default admin created (admin / admin123) — MUST be changed on first login');
  }

  // One-shot admin password reset for environments where the DB persists
  // across deploys (e.g. Railway volume) and the admin password has drifted
  // from the default. Set RESET_ADMIN_PASSWORD on the host, restart to apply,
  // sign in, then UNSET the var — otherwise the password resets on every boot.
  // Hash is bcrypt to match the rest of the auth path; legacy SHA-256 rows
  // still verify on login and get opportunistically rehashed there.
  if (process.env.RESET_ADMIN_PASSWORD) {
    const newHash = bcrypt.hashSync(process.env.RESET_ADMIN_PASSWORD, 12);
    const existing = wrapped.get('SELECT id FROM users WHERE username = ?', ['admin']);
    if (existing) {
      wrapped.run(
        'UPDATE users SET password_hash = ?, role = ?, must_change_password = 1 WHERE username = ?',
        [newHash, 'admin', 'admin']
      );
      wrapped.run('DELETE FROM sessions WHERE user_id = ?', [existing.id]);
      console.log('[reset] admin password reset from RESET_ADMIN_PASSWORD — unset this env var after signing in');
    } else {
      wrapped.run(
        'INSERT INTO users (username, password_hash, role, must_change_password) VALUES (?, ?, ?, 1)',
        ['admin', newHash, 'admin']
      );
      console.log('[reset] admin user created from RESET_ADMIN_PASSWORD — unset this env var after signing in');
    }
    scheduleSave();
  }

  console.log('Database ready at', DB_PATH, '(sql.js, in-memory + debounced disk persist)');
  return wrapped;
}

/**
 * Normalize params so callers can pass either an array or spread arguments.
 * Accepts: fn('sql', [a, b, c])  OR  fn('sql', a, b, c)  OR  fn('sql')
 */
function normalizeParams(args) {
  if (args.length === 0) return [];
  if (args.length === 1 && Array.isArray(args[0])) return args[0];
  return args;
}

function makeWrapper() {
  // sql.js note: prepared statements need .free() to release memory.
  // We create-use-free per call to keep the API simple and match the
  // existing usage patterns (no long-lived prepared statements).

  /**
   * Run a SQL with bound params and return rows as objects.
   * Internal helper used by get/all.
   */
  function execStatement(sql, params) {
    const stmt = rawDb.prepare(sql);
    try {
      stmt.bind(params);
      const rows = [];
      while (stmt.step()) rows.push(stmt.getAsObject());
      return rows;
    } finally {
      stmt.free();
    }
  }

  /**
   * Run an INSERT/UPDATE/DELETE/etc with bound params (no result rows).
   */
  function runStatement(sql, params) {
    const stmt = rawDb.prepare(sql);
    try {
      stmt.run(params);
    } finally {
      stmt.free();
    }
    // sql.js doesn't expose lastInsertRowid/changes per-statement easily.
    // Use the connection-level helpers.
    return {
      lastInsertRowid: rawDb.exec("SELECT last_insert_rowid()")[0]?.values[0]?.[0] ?? 0,
      changes: rawDb.getRowsModified(),
    };
  }

  return {
    /**
     * Execute multi-statement SQL (no params, no return).
     */
    exec(sql) {
      rawDb.exec(sql);
      scheduleSave();
    },

    /**
     * Run an INSERT/UPDATE/DELETE. Accepts params as array or spread.
     */
    run(sql, ...rest) {
      const params = normalizeParams(rest);
      const info = runStatement(sql, params);
      scheduleSave();
      return info;
    },

    /**
     * SELECT one row. Returns row object or null.
     */
    get(sql, ...rest) {
      const params = normalizeParams(rest);
      const rows = execStatement(sql, params);
      return rows[0] || null;
    },

    /**
     * SELECT many rows. Returns array (possibly empty).
     */
    all(sql, ...rest) {
      const params = normalizeParams(rest);
      return execStatement(sql, params);
    },

    /**
     * Prepare a statement. sql.js doesn't naturally cache prepared
     * statements across reuses (and freeing too early causes errors),
     * so we re-prepare on each call. Slower than better-sqlite3 but
     * functionally equivalent.
     */
    prepare(sql) {
      return {
        run(...args) {
          const info = runStatement(sql, args);
          scheduleSave();
          return info;
        },
        get(...args) {
          const rows = execStatement(sql, args);
          return rows[0] || null;
        },
        all(...args) {
          return execStatement(sql, args);
        }
      };
    },

    /**
     * Wrap a function in a transaction. Implements via BEGIN/COMMIT/ROLLBACK.
     */
    transaction(fn) {
      return function (...args) {
        rawDb.run("BEGIN");
        try {
          const result = fn(...args);
          rawDb.run("COMMIT");
          scheduleSave();
          return result;
        } catch (e) {
          rawDb.run("ROLLBACK");
          throw e;
        }
      };
    },

    // Escape hatch — only for code that needs the raw sql.js Database.
    get raw() { return rawDb; }
  };
}

function getDb() {
  if (!wrapped) throw new Error('Call initDb() first');
  return wrapped;
}

function closeDb() {
  flushSave();
  if (rawDb) {
    rawDb.close();
    rawDb = null;
    wrapped = null;
  }
}

// Force-flush the in-memory DB to disk. Used by backup/restore endpoints
// so the file on disk matches the live state at the moment of capture.
function flushDb() {
  flushSave();
}

// Replace the live DB with the contents of `buf` (a Node Buffer holding a
// valid SQLite database). Used by the Restore endpoint. The old in-memory
// DB is closed and a fresh instance is opened from the buffer; subsequent
// scheduleSave() calls will write the new state to disk.
function replaceDbFromBuffer(buf) {
  if (!SQL) throw new Error('DB not initialized');
  if (rawDb) {
    try { rawDb.close(); } catch (_) {}
  }
  rawDb = new SQL.Database(buf);
  rawDb.run("PRAGMA foreign_keys = ON;");
  // Re-register current_actor() on the fresh connection — sql.js functions
  // are per-instance, and the restored DB's stamping triggers reference it,
  // so any write would otherwise fail with "no such function".
  registerDbFunctions(rawDb);
  // Schedule a save so the new state is persisted to disk immediately
  // (and not just held in memory until the next write).
  scheduleSave();
  flushSave();
}

module.exports = { initDb, getDb, closeDb, flushDb, replaceDbFromBuffer, setActor, DB_PATH };
