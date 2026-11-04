// Bank Payment: a seller whose lots credit several bank accounts must be
// paid once PER ACCOUNT, each line carrying only that account's lots.
//
// Self-contained: builds its own in-memory schema, touches no real data.
//   node tests/bank-payment-split.js
const Database = require('better-sqlite3');
const ExcelJS = require('exceljs');
const { getBankPaymentData } = require('../calculations');
const { exportBankPayment, exportBankPaymentNew } = require('../exports');

const raw = new Database(':memory:');
// Thin wrapper matching the app's db API (db.all/get/run).
const db = {
  all: (q, p) => raw.prepare(q).all(...(p || [])),
  get: (q, p) => raw.prepare(q).get(...(p || [])),
  run: (q, p) => raw.prepare(q).run(...(p || [])),
};
raw.exec(`
  CREATE TABLE auctions (id INTEGER PRIMARY KEY, ano TEXT, date TEXT);
  CREATE TABLE traders (id INTEGER PRIMARY KEY, name TEXT, cr TEXT, tel TEXT, email TEXT,
    ifsc TEXT, acctnum TEXT, holder_name TEXT, padd TEXT, ppla TEXT, pin TEXT);
  CREATE TABLE trader_banks (id INTEGER PRIMARY KEY, trader_id INTEGER, ifsc TEXT,
    acctnum TEXT, holder_name TEXT, bank_name TEXT, is_default INTEGER);
  CREATE TABLE purchases (id INTEGER PRIMARY KEY, auction_id INTEGER, trader_id INTEGER,
    name TEXT, invo TEXT, date TEXT, qty REAL, amount REAL, cgst REAL, sgst REAL,
    igst REAL, tds REAL, total REAL, gstin TEXT);
  CREATE TABLE debit_notes (id INTEGER PRIMARY KEY, ano TEXT, trader_id INTEGER,
    name TEXT, amount REAL);
  CREATE TABLE lots (id INTEGER PRIMARY KEY, auction_id INTEGER, trader_id INTEGER,
    bank_id INTEGER, name TEXT, cr TEXT, state TEXT, lot_no TEXT, grade TEXT,
    bags INTEGER, qty REAL, price REAL, amount REAL, puramt REAL, pqty REAL,
    prate REAL, isp_puramt REAL, isp_pqty REAL, isp_prate REAL, refund REAL,
    advance REAL, balance REAL, cgst REAL, sgst REAL, igst REAL, paid TEXT,
    reserved INTEGER DEFAULT 0, code TEXT, buyer TEXT, invo TEXT, branch TEXT);
  INSERT INTO auctions (id, ano, date) VALUES (1, '30', '2026-11-04');
  -- TWOBANK banks across two accounts; ONEBANK has a single default account;
  -- MIXED has one account but some lots left untagged.
  INSERT INTO traders (id, name, cr, email, padd, ppla, pin) VALUES
    (1, 'TWOBANK SPICES', 'GSTIN.33AAAAA0000A1Z5', 'two@x.com', 'Addr1', 'Cumbum', '625516'),
    (2, 'ONEBANK TRADERS', '', 'one@x.com', 'Addr2', 'Theni', '625531'),
    (3, 'MIXED ESTATES',  '', 'mix@x.com', 'Addr3', 'Bodi',  '625513'),
    (4, 'STALE TAG FARMS', '', 'stale@x.com', 'Addr4', 'Uthamapalayam', '625533');
  INSERT INTO trader_banks (id, trader_id, ifsc, acctnum, holder_name, bank_name, is_default) VALUES
    (11, 1, 'HDFC0000111', 'ACCT-A-111', 'TWOBANK SPICES A', 'HDFC', 1),
    (12, 1, 'ICIC0000222', 'ACCT-B-222', 'TWOBANK SPICES B', 'ICICI', 0),
    (13, 2, 'SBIN0000333', 'ACCT-C-333', 'ONEBANK TRADERS',  'SBI',   1),
    (14, 3, 'AXIS0000444', 'ACCT-D-444', 'MIXED ESTATES',    'AXIS',  1),
    (15, 4, 'KVBL0000555', 'ACCT-E-555', 'STALE TAG FARMS',  'KVB',   1);
`);
// Lots. balance = what's payable; puramt = pre-discount purchase amount.
const addLot = raw.prepare(`INSERT INTO lots
  (auction_id, trader_id, bank_id, name, cr, state, lot_no, grade, amount, qty,
   puramt, pqty, prate, isp_puramt, isp_pqty, isp_prate, refund, advance, balance)
  VALUES (1, ?, ?, ?, ?, 'TAMIL NADU', ?, '1', ?, 10, ?, 10, 100, ?, 10, 100, 0, 0, ?)`);
//        trader bank  name              cr                       lot    amount puramt  isp     balance
addLot.run(1, 11, 'TWOBANK SPICES', 'GSTIN.33AAAAA0000A1Z5', '001', 1000, 1000, 1000, 1000);
addLot.run(1, 11, 'TWOBANK SPICES', 'GSTIN.33AAAAA0000A1Z5', '002', 1000, 1000, 1000, 1000);
addLot.run(1, 12, 'TWOBANK SPICES', 'GSTIN.33AAAAA0000A1Z5', '003', 3000, 3000, 3000, 3000);
addLot.run(2, 13, 'ONEBANK TRADERS', '', '010', 5000, 5000, 5000, 5000);
addLot.run(3, 14, 'MIXED ESTATES', '', '020', 2000, 2000, 2000, 2000);
addLot.run(3, null, 'MIXED ESTATES', '', '021', 500, 500, 500, 500);
// Stale tag: this lot points at account 12, which belongs to TWOBANK, not to
// its own seller — the kind of leftover a seller reassignment leaves behind.
addLot.run(4, 12, 'STALE TAG FARMS', '', '030', 700, 700, 700, 700);
// TWOBANK is a registered dealer with stamped 194Q TDS on the whole trade.
raw.prepare(`INSERT INTO purchases (auction_id, trader_id, name, invo, date, amount, tds)
             VALUES (1, 1, 'TWOBANK SPICES', 'P1', '2026-11-04', 5000, 50)`).run();
// A manual debit note on the same seller — the DISCOUNT column the new-format
// sheet renders. One deduction, so it must be SPLIT across the two lines.
raw.prepare(`INSERT INTO debit_notes (ano, trader_id, name, amount)
             VALUES ('30', 1, 'TWOBANK SPICES', 100)`).run();

const cfg = { business_mode: 'e-Trade', business_state: 'TAMIL NADU', flag_round: false,
              trade_name: 'IDEAL SPICES', short_name: 'IDSP' };

let failed = 0;
const check = (label, ok, extra) => {
  console.log((ok ? '  PASS ' : '  FAIL ') + label + (ok ? '' : ' — ' + extra));
  if (!ok) failed++;
};
const r2 = n => Math.round((Number(n) || 0) * 100) / 100;

(async () => {
console.log('[1] getBankPaymentData splits a multi-bank seller into one row per account');
const rows = getBankPaymentData(db, 1, cfg);
const two = rows.filter(r => r.name === 'TWOBANK SPICES');
check('two payment rows for the two-account seller', two.length === 2, JSON.stringify(two.map(r => r.accountNo)));
const byAcct = Object.fromEntries(two.map(r => [r.accountNo, r]));
check('account A carries only its own lots', byAcct['ACCT-A-111'] && byAcct['ACCT-A-111'].lots === '001,002',
      JSON.stringify(byAcct['ACCT-A-111'] && byAcct['ACCT-A-111'].lots));
check('account B carries only its own lot', byAcct['ACCT-B-222'] && byAcct['ACCT-B-222'].lots === '003',
      JSON.stringify(byAcct['ACCT-B-222'] && byAcct['ACCT-B-222'].lots));
check('each row routes to its own IFSC',
      byAcct['ACCT-A-111'].ifsc === 'HDFC0000111' && byAcct['ACCT-B-222'].ifsc === 'ICIC0000222',
      JSON.stringify(two.map(r => r.ifsc)));
check('each row pays its own lots\' balance (2000 / 3000 pre-TDS)',
      r2(byAcct['ACCT-A-111'].payable) === 2000 && r2(byAcct['ACCT-B-222'].payable) === 3000,
      JSON.stringify(two.map(r => r.payable)));
check('the 50 TDS splits 20/30 by payable and still foots to 50',
      r2(byAcct['ACCT-A-111'].tds) === 20 && r2(byAcct['ACCT-B-222'].tds) === 30
        && r2(two.reduce((a, r) => a + r.tds, 0)) === 50,
      JSON.stringify(two.map(r => r.tds)));
check('amount is each account\'s payable net of its TDS',
      r2(byAcct['ACCT-A-111'].amount) === 1980 && r2(byAcct['ACCT-B-222'].amount) === 2970,
      JSON.stringify(two.map(r => r.amount)));
check('both rows flag the split and count the accounts',
      two.every(r => r.multipleBanks === true && r.bankCount === 2), JSON.stringify(two.map(r => r.bankCount)));
check('the default account is listed first', two[0].accountNo === 'ACCT-A-111', two[0].accountNo);
check('REMARKS name the row\'s own lots',
      /for lots 001,002$/.test(byAcct['ACCT-A-111'].remarks) && /for lot 003$/.test(byAcct['ACCT-B-222'].remarks),
      JSON.stringify(two.map(r => r.remarks)));

console.log('\n[2] A single-account seller is untouched — still one row');
const one = rows.filter(r => r.name === 'ONEBANK TRADERS');
check('exactly one row', one.length === 1, String(one.length));
check('whole payable on the default account',
      one[0] && one[0].accountNo === 'ACCT-C-333' && r2(one[0].payable) === 5000, JSON.stringify(one[0]));
check('not flagged as split', one[0] && one[0].multipleBanks === false && one[0].bankCount === 1,
      JSON.stringify([one[0].multipleBanks, one[0].bankCount]));

console.log('\n[3] Tagged + untagged lots that credit the SAME account stay one payment');
const mixed = rows.filter(r => r.name === 'MIXED ESTATES');
check('one row, not two', mixed.length === 1, JSON.stringify(mixed.map(r => r.lots)));
check('it pays both lots', mixed[0] && mixed[0].lots === '020,021' && r2(mixed[0].payable) === 2500,
      JSON.stringify([mixed[0].lots, mixed[0].payable]));
check('no spurious multi-bank flag', mixed[0] && mixed[0].multipleBanks === false, String(mixed[0].multipleBanks));

console.log('\n[3b] A lot tagged with ANOTHER party\'s account is not wired to them');
const stale = rows.filter(r => r.name === 'STALE TAG FARMS');
check('one row, paid into its own seller\'s account',
      stale.length === 1 && stale[0].accountNo === 'ACCT-E-555' && stale[0].ifsc === 'KVBL0000555',
      JSON.stringify(stale.map(r => [r.accountNo, r.ifsc])));
check('the other party gains nothing from the stale tag',
      !rows.some(r => r.name === 'TWOBANK SPICES' && String(r.lots).includes('030')),
      JSON.stringify(rows.filter(r => r.name === 'TWOBANK SPICES').map(r => r.lots)));

console.log('\n[4] The totals still foot — nothing double-paid, nothing dropped');
const wholeTrade = raw.prepare('SELECT COALESCE(SUM(balance),0) b FROM lots').get().b;
const paidOut = rows.reduce((a, r) => a + (Number(r.payable) || 0), 0);
check('Σ row payable = Σ lot balance', r2(paidOut) === r2(wholeTrade), `${paidOut} vs ${wholeTrade}`);
const allLots = rows.flatMap(r => String(r.lots).split(',')).filter(Boolean).sort();
check('every lot appears exactly once across the rows',
      allLots.join(',') === '001,002,003,010,020,021,030', allLots.join(','));

console.log('\n[5] The XLSX sheets carry a line per account');
const readSheet = async (buf) => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const ws = wb.getWorksheet('BANK_PAYMENT');
  const head = ws.getRow(1).values.slice(1);
  const out = [];
  ws.eachRow((row, n) => { if (n === 1) return;
    const o = {}; head.forEach((h, i) => { o[h] = row.getCell(i + 1).value; }); out.push(o); });
  return out;
};
const legacy = await readSheet(await exportBankPayment(db, 1, cfg, null, { names: ['id:1'] }));
check('old format: 2 rows for the ticked two-account seller', legacy.length === 2, JSON.stringify(legacy.length));
check('old format: each row holds its own account + lots',
      legacy[0].BENEFICI_A === 'ACCT-A-111' && legacy[0].BENEFICI_B === '001,002'
      && legacy[1].BENEFICI_A === 'ACCT-B-222' && legacy[1].BENEFICI_B === '003',
      JSON.stringify(legacy.map(r => [r.BENEFICI_A, r.BENEFICI_B])));
const neu = await readSheet(await exportBankPaymentNew(db, 1, cfg, null, { names: ['id:1'] }));
check('new format: 2 rows, money columns split per account',
      neu.length === 2 && r2(neu[0].PURCHASEAMT) === 2000 && r2(neu[0].TDS) === 20 && r2(neu[0].PAYMENTAMT) === 1980
      && r2(neu[1].PURCHASEAMT) === 3000 && r2(neu[1].TDS) === 30 && r2(neu[1].PAYMENTAMT) === 2970,
      JSON.stringify(neu.map(r => [r.PURCHASEAMT, r.TDS, r.PAYMENTAMT])));
check('new format: the 100 debit note splits 40/60, not 100 on each row',
      r2(neu[0].DISCOUNT) === 40 && r2(neu[1].DISCOUNT) === 60,
      JSON.stringify(neu.map(r => r.DISCOUNT)));

console.log('\n[6] A lot pick spanning both accounts still pays each account its own lots');
const picked = await readSheet(await exportBankPayment(db, 1, cfg, null,
  { names: ['id:1'], lots: { 'id:1': ['002', '003'] } }));
check('2 rows (one per account in the pick)', picked.length === 2, JSON.stringify(picked.length));
check('account A pays only lot 002, account B only lot 003',
      picked[0].BENEFICI_A === 'ACCT-A-111' && picked[0].BENEFICI_B === '002' && r2(picked[0].PAYMENTAMO) === 1000
      && picked[1].BENEFICI_A === 'ACCT-B-222' && picked[1].BENEFICI_B === '003' && r2(picked[1].PAYMENTAMO) === 3000,
      JSON.stringify(picked.map(r => [r.BENEFICI_A, r.BENEFICI_B, r.PAYMENTAMO])));

console.log('\n[7] A pick confined to ONE account produces a single line for it');
const onlyB = await readSheet(await exportBankPayment(db, 1, cfg, null,
  { names: ['id:1'], lots: { 'id:1': ['003'] } }));
check('one row, account B', onlyB.length === 1 && onlyB[0].BENEFICI_A === 'ACCT-B-222',
      JSON.stringify(onlyB.map(r => r.BENEFICI_A)));
check('it does not also emit the untouched account-A row',
      !onlyB.some(r => r.BENEFICI_A === 'ACCT-A-111'), JSON.stringify(onlyB.map(r => r.BENEFICI_A)));

console.log(failed ? `\n${failed} FAILED` : '\nall passed');
process.exit(failed ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
