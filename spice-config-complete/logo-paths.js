// logo-paths.js — resolves bundled-vs-user-uploaded logo paths.
//
// Spice-config keeps the bundled brand logo in public/ (so the static
// server can serve it directly) but cloud / persistent installs may
// override it by dropping a same-named file into SPICE_DATA_DIR/logos.
// This helper centralises the lookup so callers (mobile-bridge, PDF
// generators, etc.) all agree on precedence:
//   1. $SPICE_DATA_DIR/logos/<name>   — operator-uploaded override
//   2. <repo>/public/<name>           — bundled default
// Returns null if neither exists, so callers can fall back gracefully
// (e.g. render the receipt without a logo block instead of crashing).

const path = require('path');
const fs = require('fs');

function resolveLogoPath(name) {
  if (!name) return null;
  const dataDir = process.env.SPICE_DATA_DIR;
  if (dataDir) {
    const userPath = path.join(dataDir, 'logos', name);
    if (fs.existsSync(userPath)) return userPath;
  }
  const bundled = path.join(__dirname, 'public', name);
  if (fs.existsSync(bundled)) return bundled;
  return null;
}

module.exports = { resolveLogoPath };
