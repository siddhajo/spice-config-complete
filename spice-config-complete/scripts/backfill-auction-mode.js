// scripts/backfill-auction-mode.js — assign a business mode to legacy auctions
//
// Pre-mode-column auctions (and auctions created by import before the import
// endpoint started stamping `mode`) have mode='' or NULL, which makes them
// visible in BOTH e-Trade and e-Auction views. Run this once to lock those
// rows into a single mode.
//
// Idempotent — only touches rows where mode is empty/NULL.
//
// Usage:
//   node scripts/backfill-auction-mode.js              # defaults to e-Trade
//   node scripts/backfill-auction-mode.js --mode e-Auction
//   node scripts/backfill-auction-mode.js --dry-run    # show counts, don't write

const path = require('path');
process.chdir(path.dirname(__dirname));

const { initDb, getDb } = require('../db');

const args = process.argv.slice(2);
const DRY = args.includes('--dry-run');
const modeIdx = args.indexOf('--mode');
const TARGET_MODE = modeIdx >= 0 ? args[modeIdx + 1] : 'e-Trade';

if (!['e-Trade', 'e-Auction'].includes(TARGET_MODE)) {
  console.error(`Invalid --mode "${TARGET_MODE}". Must be "e-Trade" or "e-Auction".`);
  process.exit(1);
}

(async () => {
  await initDb();
  const db = getDb();

  const empty = db.get(
    "SELECT COUNT(*) AS n FROM auctions WHERE mode IS NULL OR mode = ''"
  );
  console.log(`Found ${empty.n} auction row(s) with empty mode.`);

  if (empty.n === 0) {
    console.log('Nothing to do.');
    process.exit(0);
  }

  if (DRY) {
    console.log(`[dry-run] Would set mode='${TARGET_MODE}' on ${empty.n} row(s).`);
    process.exit(0);
  }

  const res = db.run(
    "UPDATE auctions SET mode = ? WHERE mode IS NULL OR mode = ''",
    [TARGET_MODE]
  );
  console.log(`Updated ${res.changes} auction row(s) to mode='${TARGET_MODE}'.`);
})();
