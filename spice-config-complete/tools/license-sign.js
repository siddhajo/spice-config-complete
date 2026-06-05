#!/usr/bin/env node
/**
 * tools/license-sign.js
 * --------------------------------------------------------------
 * Mint a signed license token for a deployed instance.
 *
 * USAGE
 *   node tools/license-sign.js --install-id <UUID> --days 30
 *   node tools/license-sign.js --install-id <UUID> --until 2026-12-31
 *   node tools/license-sign.js --install-id <UUID> --days 30 --note "Acme Co — March renewal"
 *
 * REQUIRED ARGS
 *   --install-id <uuid>    — the customer's install ID (they read it
 *                            off the renewal page or the topbar pill).
 *
 * EXPIRY (pick ONE)
 *   --days <N>             — token expires N calendar days from today.
 *   --until <YYYY-MM-DD>   — explicit expiry date.
 *
 * OPTIONAL
 *   --note <text>          — free-form audit note baked into the
 *                            payload. Visible to anyone who decodes the
 *                            token; keep it short.
 *
 * ENV
 *   LICENSE_SECRET         — HMAC signing secret. MUST match the secret
 *                            set on the deployed server. If unset,
 *                            falls back to the in-repo development
 *                            default (loud warning).
 *
 * OUTPUT
 *   The token is printed to stdout on its own line, with a one-line
 *   confirmation to stderr (won't pollute pipes / clipboards).
 */

const path = require('path');
const license = require(path.join(__dirname, '..', 'license'));

function parseArgs(argv) {
  const out = {};
  for (let i = 2; i < argv.length; i++) {
    const k = argv[i];
    const v = argv[i + 1];
    if (k === '--install-id') { out.installId = v; i++; }
    else if (k === '--days')  { out.days  = Number(v); i++; }
    else if (k === '--until') { out.until = v; i++; }
    else if (k === '--note')  { out.note  = v; i++; }
    else if (k === '-h' || k === '--help') { out.help = true; }
  }
  return out;
}

function usage(exit) {
  const msg =
`Usage:
  node tools/license-sign.js --install-id <UUID> --days 30
  node tools/license-sign.js --install-id <UUID> --until 2026-12-31

Required:
  --install-id <uuid>   The customer's install ID (visible on /renew.html
                        and in the server boot log).

Exactly one of:
  --days <N>            Token expires N calendar days from today.
  --until <YYYY-MM-DD>  Explicit expiry date (inclusive).

Optional:
  --note <text>         Free-form note baked into the payload (audit).

Env:
  LICENSE_SECRET        HMAC secret. MUST match the deployed server.
`;
  process.stderr.write(msg);
  process.exit(exit || 0);
}

function pad(n) { return String(n).padStart(2, '0'); }
function addDaysIso(days) {
  const d = new Date();
  d.setDate(d.getDate() + Number(days));
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

const args = parseArgs(process.argv);
if (args.help) usage(0);
if (!args.installId) {
  process.stderr.write('error: --install-id is required\n\n');
  usage(2);
}
if ((args.days == null || isNaN(args.days)) && !args.until) {
  process.stderr.write('error: exactly one of --days <N> or --until <YYYY-MM-DD> is required\n\n');
  usage(2);
}
if (args.days != null && !isNaN(args.days) && args.until) {
  process.stderr.write('error: pass --days OR --until, not both\n\n');
  usage(2);
}

let expires_at;
if (args.until) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(args.until)) {
    process.stderr.write('error: --until must be YYYY-MM-DD\n');
    process.exit(2);
  }
  expires_at = args.until;
} else {
  if (args.days <= 0) {
    process.stderr.write('error: --days must be positive\n');
    process.exit(2);
  }
  expires_at = addDaysIso(args.days);
}

const payload = {
  install_id: args.installId,
  expires_at,
  issued_at: new Date().toISOString(),
};
if (args.note) payload.note = String(args.note);

const token = license.signToken(payload);

// Confirmation to STDERR (won't pollute when piping the token).
process.stderr.write(
  `\n  ✓ Token minted\n` +
  `      install_id  ${payload.install_id}\n` +
  `      expires_at  ${payload.expires_at}\n` +
  `      issued_at   ${payload.issued_at}\n` +
  (args.note ? `      note        ${payload.note}\n` : '') +
  (license._LICENSE_SECRET_IS_FALLBACK
    ? `      ⚠  Signed with the FALLBACK secret — set LICENSE_SECRET\n` +
      `         to your production secret before sending this to a customer.\n`
    : '') +
  `\n  Token (one line, copy/paste verbatim):\n\n`
);

// The token itself goes to STDOUT — clean, pipeable, single line.
process.stdout.write(token + '\n');
