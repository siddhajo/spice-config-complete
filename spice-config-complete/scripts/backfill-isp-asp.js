// scripts/backfill-isp-asp.js — populate isp_*/asp_* columns for existing lots
//
// Run this ONCE after deploying the dual-storage feature on a database that
// has lots pre-dating the migration. For each lot, recomputes both ISP and
// ASP planter views using current settings and writes them.
//
// Idempotent — safe to re-run. Skips rows that already have isp_puramt > 0.
//
// Usage:  node scripts/backfill-isp-asp.js [--all]
//   --all   re-backfill every row, even ones that already have isp_puramt

const path = require('path');
process.chdir(path.dirname(__dirname));   // run from project root

const { initDb, getDb } = require('../db');
const { getSettingsFlat } = require('../company-config');
const { calculateLot } = require('../calculations');

const REDO_ALL = process.argv.includes('--all');

(async () => {
  await initDb();
  const db = getDb();
  const cfg = getSettingsFlat(db);

  // Pick rows to backfill
  const filter = REDO_ALL ? '1=1' : '(isp_puramt IS NULL OR isp_puramt = 0)';
  const lots = db.prepare(`
    SELECT * FROM lots
    WHERE ${filter}
      AND qty > 0
      AND price > 0
  `).all();

  console.log(`Backfilling ${lots.length} lot row(s)...`);
  if (lots.length === 0) {
    console.log('Nothing to do. Use --all to force re-backfill.');
    process.exit(0);
  }

  const upd = db.prepare(`
    UPDATE lots SET
      isp_pqty = ?, isp_prate = ?, isp_puramt = ?,
      asp_pqty = ?, asp_prate = ?, asp_puramt = ?
    WHERE id = ?
  `);

  let count = 0;
  const tx = db.transaction(() => {
    for (const lot of lots) {
      const c = calculateLot(lot, cfg);
      upd.run(
        c.isp_pqty || 0, c.isp_prate || 0, c.isp_puramt || 0,
        c.asp_pqty || 0, c.asp_prate || 0, c.asp_puramt || 0,
        lot.id
      );
      count++;
    }
  });
  tx();

  console.log(`✓ Updated ${count} rows.`);

  // Spot-check: print one row before/after
  const sample = db.prepare(`SELECT lot_no, name, qty, price, pqty, prate, puramt, isp_pqty, isp_prate, isp_puramt, asp_pqty, asp_prate, asp_puramt FROM lots WHERE id = ?`).get(lots[0].id);
  console.log('Sample row after backfill:');
  console.log(`  lot ${sample.lot_no} (${sample.name}):`);
  console.log(`    legacy: pqty=${sample.pqty} prate=${sample.prate} puramt=${sample.puramt}`);
  console.log(`    isp:    pqty=${sample.isp_pqty} prate=${sample.isp_prate} puramt=${sample.isp_puramt}`);
  console.log(`    asp:    pqty=${sample.asp_pqty} prate=${sample.asp_prate} puramt=${sample.asp_puramt}`);
})().catch(e => { console.error('FAIL:', e); process.exit(1); });
