# Spice Config — Architecture & Feature Guide

> **What this app is:** A web-based admin system for **spice (cardamom) auction houses** in India.
> It manages the full back-office workflow of an auction day: recording lots, allocating them to
> buyers, generating GST-compliant invoices/bills/debit-notes, settling payments to sellers, and
> exporting everything to Tally, the Spices Board portal, e-way bills, and Excel/PDF/DBF.
>
> **Version:** 2.1.0 &nbsp;•&nbsp; **Stack:** Node.js + Express + SQLite (sql.js) &nbsp;•&nbsp; **Ships as:** web app (Railway/Docker) and Windows desktop app (Electron)

---

## Table of Contents

1. [The Big Picture](#1-the-big-picture)
2. [How It's Built (Tech Stack)](#2-how-its-built-tech-stack)
3. [System Diagram](#3-system-diagram)
4. [The Core Workflow](#4-the-core-workflow-an-auction-day)
5. [Data Model](#5-data-model)
6. [Users, Roles & Permissions](#6-users-roles--permissions)
7. [Multi-Company (ISP / ASP Presets)](#7-multi-company-isp--asp-presets)
8. [Licensing](#8-licensing)
9. [File Map (What Each File Does)](#9-file-map-what-each-file-does)
10. [Feature Inventory & Status](#10-feature-inventory--status)
11. [Deployment](#11-deployment)

---

## 1. The Big Picture

A cardamom auction works like this: **sellers** (planters / pooling agents) bring lots of spice,
**buyers** (dealers) bid on them, and the **auction house** sits in the middle — taking a commission,
handling the paperwork, collecting money from buyers, and paying sellers.

Spice Config is the software that runs the auction house's back office. One person enters lots during
the auction, and the app does the rest:

- **Calculates** every figure — purchase amount, commission, GST (CGST/SGST/IGST), TDS, net payment.
- **Generates** all statutory documents — sales invoices, purchase invoices, bills of supply, debit notes.
- **Exports** to everything downstream — Tally accounting, the Spices Board statutory forms, e-way bill
  distances, Excel reports, and legacy FoxPro DBF files.
- **Settles** — produces payment slips per seller and per bank.

A second, mobile-friendly app lets floor staff enter lots from a phone during the live auction.

### Who uses it

| User | What they do |
|------|--------------|
| **Auction-floor staff** | Enter lots live from a phone (Lot Entry app) |
| **Back-office operators** | Generate invoices, manage buyers/sellers, run exports |
| **Managers** | Configure rates/settings, revert mistakes, oversee branches |
| **Admins** | Manage users, bulk operations, backups |
| **The developer (you)** | Deploy new customers, issue licenses, set branding |

---

## 2. How It's Built (Tech Stack)

| Layer | Technology | Notes |
|-------|-----------|-------|
| **Server** | Node.js + Express | Single `server.js`, ~9,300 lines, 186 API routes |
| **Database** | SQLite via `sql.js` | Pure-JavaScript SQLite (WASM). No native compilation → runs anywhere |
| **Frontend (desktop)** | Single-page app | One compiled `public/index.html` (~1.2 MB), talks to the API |
| **Frontend (mobile)** | PWA | `public-mobile/app.html`, installable on phones, works offline |
| **PDF generation** | PDFKit | Invoices, reports, payment slips |
| **Excel** | ExcelJS / xlsx | Imports and exports |
| **Auth** | Token sessions + bcrypt | Stored in the database |
| **Packaging** | Electron + electron-builder | Optional Windows desktop installer |

**Why sql.js (not a "real" database server?)** — The whole database is one file (`config.db`). The app
loads it into memory and writes changes back to disk. This makes each customer install completely
self-contained and trivial to deploy, back up (copy one file), and move. The trade-off is that it's
single-instance: one running server owns the file at a time (enforced by a `server.lock`).

---

## 3. System Diagram

```
                          ┌─────────────────────────────────────────┐
   Auction floor          │            SPICE CONFIG SERVER           │
  ┌───────────────┐       │              (server.js)                 │
  │  Mobile PWA   │──────▶│                                          │
  │  (lot entry)  │       │   Express API  ──┐                       │
  └───────────────┘       │   186 routes      │                      │
                          │                   ▼                      │
  ┌───────────────┐       │   ┌──────────────────────────────────┐  │
  │ Desktop admin │──────▶│   │  Auth + Roles  │  Calculations    │  │
  │  (browser/    │       │   │  Company cfg   │  PDF / XLSX / XML │  │
  │   Electron)   │       │   └──────────────────────────────────┘  │
  └───────────────┘       │                   │                      │
                          │                   ▼                      │
                          │        ┌────────────────────┐           │
                          │        │   SQLite (sql.js)   │           │
                          │        │   data/config.db    │           │
                          │        └────────────────────┘           │
                          └─────────────────┬───────────────────────┘
                                            │ exports
            ┌───────────────┬───────────────┼───────────────┬───────────────┐
            ▼               ▼               ▼               ▼               ▼
       Tally ERP      Spices Board     e-Way Bill        Excel /         Legacy
       (XML import)   (statutory       (PIN distance)    PDF reports     FoxPro DBF
                       forms C/D)
```

---

## 4. The Core Workflow (an Auction Day)

This is the heart of the app. Everything flows through this sequence:

```
  1. TRADE / AUCTION          A trade session is created (date, crop, branch).
         │
         ▼
  2. LOT ENTRY                Each lot recorded: seller, grade, qty, price, buyer.
         │                    (Done live, often from the mobile app.)
         ▼
  3. ALLOCATION               Lots assigned/reassigned to buyers; lot-number
         │                    ranges allocated per branch.
         ▼
  4. CALCULATION              App computes purchase amount, commission, GST, TDS,
         │                    net payment for every lot. (calculations.js)
         ▼
  5. PRICE CHECK (optional)   Prices verified against an uploaded price list.
         │
         ▼
  6. DOCUMENT GENERATION      ┌─ Sales invoices  (to buyers)
         │                    ├─ Purchase invoices (from registered sellers)
         │                    ├─ Bills of supply  (from unregistered/farmer sellers)
         │                    └─ Debit notes      (discounts)
         ▼
  7. PAYMENTS                 Payment slips per seller / per bank.
         │
         ▼
  8. EXPORTS                  Tally XML · Spices Board forms · e-way bill ·
                              Excel · PDF · DBF.
```

**Key terms:**

- **Trade / Auction** — one auction session (the app uses both words; internally a row in `auctions`).
- **Lot** — one parcel of spice sold in the auction.
- **Trader** — a **seller** (planter or pooling agent).
- **Buyer** — a **dealer** who buys lots.
- **e-Trade vs e-Auction mode** — two business modes the operator can switch between; affects labels
  and some document formatting.

---

## 5. Data Model

Everything lives in **one SQLite file** (`data/config.db`). There is **no tenant/company ID column** —
one database = one customer (see [Onboarding](#11-deployment) and the [Onboarding Guide](./ONBOARDING.md)).

### Main tables

| Table | Holds |
|-------|-------|
| `auctions` | Trade sessions (date, crop, state/branch, price-check status) |
| `lots` | Individual lots (qty, grade, price, buyer, lock state) |
| `traders` | Sellers — name, contact, WhatsApp, email |
| `trader_banks` | Bank accounts per seller (a seller can have several) |
| `buyers` | Dealers — name, address, GSTIN, PAN, TDS quota |
| `lot_allocations` | Lot-number ranges assigned per branch per auction |
| `invoices` | Sales invoices (to buyers) |
| `purchases` | Purchase invoices (from registered-dealer sellers) |
| `bills` | Bills of supply (from unregistered/agriculturist sellers) |
| `debit_notes` | Discount notes |
| `route_distances` | Cached PIN-to-PIN distances for e-way bills |
| `company_settings` | **~194 key/value rows** — all configuration (see below) |
| `company_presets` | ISP / ASP company-identity snapshots |
| `users` | Operator accounts (bcrypt password, role) |
| `sessions` | Active login tokens (max 30-day life) |
| `license_state` | Single row — install ID, expiry, license token |
| `audit_log` / `delete_log` / `login_history` / `import_log` | Audit trails |

### How configuration is stored

The app's entire configuration is **not** in a config file — it's the `company_settings` table, a
key/value store seeded with ~194 defaults on first boot. Categories include company identity,
Kerala/Tamil-Nadu addresses, sister-company details, rates (commission, GST, TDS, handling…), HSN/SAC
codes, bank details, season dates, invoice numbering, Tally ledger mappings, and ~23 feature flags.

Edit it in the UI under **Settings**, or via `PUT /api/company-settings`.

### Audit trail

Every business row (`traders`, `buyers`, `lots`, `invoices`, …) carries `modified_at` and `modified_by`,
stamped automatically by database triggers using the logged-in username. The `audit_log` table records
create/edit/delete actions with the device type (mobile vs desktop).

---

## 6. Users, Roles & Permissions

Authentication is **token-based**: log in with username + password (bcrypt-hashed), receive a session
token, and send it as `Authorization: Bearer <token>`. Multiple devices can be logged in at once; each
gets its own token, listed and revocable under "My Sessions".

There are **5 fixed roles**, each granting a set of named capabilities. The role hierarchy is additive
(each tier includes the ones below it):

| Role | Can do | Key capabilities |
|------|--------|------------------|
| **viewer** | Read & export only | `view`, `export` |
| **lot_entry** | Auction-floor lot entry | + `lot_write`, `auction_write` (Lot Entry tab only) |
| **operator** | Daily back-office work | + `invoice_write`, `trader_write`, `buyer_write` |
| **manager** | Branch oversight | + `auction_write`, `invoice_revert`, `settings_write`, `state_toggle` |
| **admin** | Full control | + `delete`, `delete_all`, `user_manage` |

The frontend hides buttons the user's role can't use; the server enforces the same rules on every API
route (returns `403` if the capability is missing). Roles and their capabilities are defined in one
place — `server.js:257` (`ROLE_PERMISSIONS`).

**Default admin:** On first boot the app creates `admin` / `admin123` with a forced password change on
first login. This default cannot survive that first login. (Host operators can also reset it via the
`RESET_ADMIN_PASSWORD` environment variable.)

---

## 7. Multi-Company (ISP / ASP Presets)

The app supports running **two company identities** out of one install — typically a main company and
a sister concern:

- **ISP** — the primary company (e.g. Tamil Nadu base).
- **ASP** — the sister company (e.g. Kerala base).

These are stored as **presets** (snapshots of 8 identity fields: trade name, legal name, PAN, CIN,
GSTIN, etc.) in the `company_presets` table. A manager flips the active preset with the **Logo Code**
dropdown (`PUT /api/company-presets/active`, requires the `state_toggle` capability). Invoices then use
the matching company's letterhead, addresses, and GSTIN — Kerala-state invoices use ASP, Tamil-Nadu use
ISP.

> **Important:** This is **not** multi-tenancy. Both identities share the same database and data. It's a
> branding/identity toggle for businesses that legally operate as two entities. To run two *separate*
> businesses, deploy two installs — see the [Onboarding Guide](./ONBOARDING.md).

### White-label branding

A developer-only panel at `/admin/branding` (key-gated, no login) can lock a preset and apply custom
colors/fonts/density for a specific customer, then hide the appearance controls from that customer's UI.

---

## 8. Licensing

Each install is **time-limited** and must be renewed. This is honest-customer DRM, not fortress security.

- On first boot, the app generates a unique **`install_id`** (UUID) and starts a **30-day trial**.
- A license **token** is a signed blob (`HMAC-SHA256` using the `LICENSE_SECRET` env var) that encodes
  the install ID and a new expiry date.
- When the license expires, **login returns HTTP 451** and the user is sent to `/renew.html`.

### Renewal flow

```
  Customer                          Developer (you)
  ────────                          ───────────────
  Sees expiry warning pill
  (amber ≤7 days, red ≤3 days)
         │
         ▼
  Opens /renew.html,
  copies the Install ID  ──────────▶ Receives Install ID
                                            │
                                            ▼
                                     Runs:  LICENSE_SECRET=… node tools/license-sign.js \
                                              --install-id <id> --days 30
                                            │
  Pastes token, clicks Apply ◀───────  Sends token back
         │
         ▼
  Works immediately (no restart)
```

> **Setup requirement:** `LICENSE_SECRET` must be set to the **same secret value** on the customer's
> server and on your signing machine. Full details in [LICENSING.md](./LICENSING.md).

---

## 9. File Map (What Each File Does)

| File | Responsibility |
|------|----------------|
| **`server.js`** | The whole HTTP API — routes, auth, all 186 endpoints. The hub. |
| **`db.js`** | Database setup, schema, migrations, the single-instance lock, audit triggers. |
| **`company-config.js`** | The ~194 default settings + preset logic. |
| **`license.js`** / `tools/license-sign.js` | License validation / token signing. |
| **`calculations.js`** | Core auction math — purchase amount, commission, GST, TDS, net pay. |
| **`invoice-pdf.js`** | GST invoice / purchase / bill PDFs. |
| **`debit-note-print.js`** | Debit-note PDFs. |
| **`tally-xml.js`** | Tally ERP import XML (4 voucher types). |
| **`exports.js`** / `exports-pdf.js` / `dbf-exports.js` | Excel, PDF-table, and FoxPro DBF exports. |
| **`auction-reports.js`** / `lorry-reports.js` / `spice-board-reports.js` / `reports.js` | Report builders (lot slips, truck lists, statutory Spices Board forms, trade summaries). |
| **`report-formatters.js`** | Shared Indian number / date / header formatting. |
| **`distance.js`** | PIN-to-PIN road distance for e-way bills. |
| **`amount-words.js`** | Rupees-in-words (Crore/Lakh) for invoices. |
| **`date-format.js`** / `logo-paths.js` | Small shared helpers. |
| **`mobile-bridge.js`** | Mounts the mobile PWA + its API shims and receipt printing. |
| **`recover-isp.js`** / `scripts/*` | One-off recovery and backfill utilities. |

---

## 10. Feature Inventory & Status

**Legend:** ✅ Complete & in production · 🟡 Works with a known limitation · 🔲 Planned / not built

> As of the current codebase, **every shipped module is complete and production-ready** — no `TODO`,
> `FIXME`, `WIP`, or stub markers were found in feature code.

### Auction & lot management

| Feature | Status | Notes |
|---------|:------:|-------|
| Trades / auctions (create, edit, import) | ✅ | XLSX bulk import + template |
| Lot entry (create, edit, bulk ops, lock/unlock) | ✅ | Per-lot audit trail |
| Buyer allocation & auto-fill | ✅ | Lot-range allocation per branch |
| Lot calculation engine | ✅ | Matches legacy FoxPro behaviour |
| Price-list mapping & price check | ✅ | Feature-flagged (`flag_price_list_mapping`, `flag_price_check`) |
| Traders (sellers) — CRUD, banks, import | ✅ | Quick-add during lot entry |
| Buyers (dealers) — CRUD, import | ✅ | GSTIN / PAN / TDS quota |

### Documents

| Feature | Status | Notes |
|---------|:------:|-------|
| Sales invoices (PDF, single + bulk) | ✅ | State-aware (ISP/ASP), bulk capped at 60 lots |
| Purchase invoices (registered dealers) | ✅ | Single + batch PDF |
| Bills of supply (unregistered/farmers) | ✅ | Single + batch PDF |
| Debit notes | ✅ | Buyer-grouped, original/dup/triplicate |
| Amount-in-words | ✅ | Indian Crore/Lakh format |
| Invoice revert / undo | ✅ | Audit-tracked, single or whole-trade |

### Payments & tax

| Feature | Status | Notes |
|---------|:------:|-------|
| Payment summary per seller | ✅ | + bank-wise breakdown |
| Payment slip PDFs (single + bulk) | ✅ | |
| TDS return summary + XLSX | ✅ | |
| Sales / purchase journals | ✅ | Feeds Tally export |

### Exports & integrations

| Feature | Status | Notes |
|---------|:------:|-------|
| Tally ERP XML | ✅ | Sales, RD purchase, URD purchase, debit-note vouchers |
| Spices Board statutory reports | ✅ | Form C, Form D, Buyers Statement, e-Auction CSV |
| Lorry / truck reports | ✅ | Lot slip, truck list, buyer lorry (state-grouped) |
| Excel exports (11 types) | ✅ | Branded headers, Indian number format |
| PDF table exports | 🟡 | All types except `full_file` (too wide for landscape A4) |
| FoxPro DBF export | ✅ | Legacy interchange (CPA1.DBF structure) |
| e-Way bill distance (PIN-to-PIN) | ✅ | Haversine × road factor, cached, manual override |
| GST lookup (GSTIN verify) | ✅ | Status + real-time lookup endpoints |
| WhatsApp send (text + document) | ✅ | Config, test, templates, webhook |

### Platform & admin

| Feature | Status | Notes |
|---------|:------:|-------|
| Auth, roles & permissions (5 roles) | ✅ | Multi-device sessions |
| User management | ✅ | Last-admin guard |
| Company settings (~194 keys) | ✅ | Import/export, logo upload |
| ISP/ASP company presets | ✅ | Logo-code toggle |
| White-label branding (`/admin/branding`) | ✅ | Developer-gated |
| Licensing + self-service renewal | ✅ | 30-day trial, signed tokens |
| Backup & restore | ✅ | One-file DB snapshot up/download |
| Audit log + delete log + login history | ✅ | Searchable, device-aware |
| Bulk "Delete All" with snapshots | ✅ | Pre-delete backup |
| Legacy data import (old system) | ✅ | Preview → verify → run → undo |
| Reports dashboard (stats, insights, trends) | ✅ | Trade summary, branch comparison |
| Mobile PWA (lot entry + receipts) | ✅ | Offline-capable, receipt printing |
| Desktop app (Electron / Windows installer) | ✅ | Auto-update support configured |

### Known limitations / not built

| Item | Status | Notes |
|------|:------:|-------|
| `full_file` PDF export | 🟡 | Intentionally unsupported — table too wide for landscape |
| True multi-tenancy (multiple companies in one DB) | 🔲 | By design — one install per business (use ISP/ASP presets for two identities) |
| The `publish` GitHub owner/repo in `package.json` | 🟡 | Placeholders — must be filled before building auto-updating desktop releases |

---

## 11. Deployment

Each customer runs their **own install** (one server + one `config.db`). Three supported targets:

| Target | How | Files |
|--------|-----|-------|
| **Railway / cloud (Docker)** | `node server.js` in a container; mount a **persistent volume** for `data/` | `Dockerfile` (sql.js variant), `Procfile` |
| **Railway (Nixpacks)** | Build with Node 20 + python/gcc (for optional `better-sqlite3`) | `nixpacks.toml` |
| **Windows desktop** | Electron installer | `npm run build:win` |

**Key environment variables:**

| Variable | Purpose |
|----------|---------|
| `PORT` | HTTP port (default `3001`) |
| `SPICE_DATA_DIR` | Where `config.db` and backups live (default `./data`) |
| `LICENSE_SECRET` | **Required** — secret for signing/validating license tokens (must match your signer) |
| `LICENSE_TRIAL_DAYS` | Trial length on first boot (default `30`) |
| `RESET_ADMIN_PASSWORD` | One-shot admin password reset on restart |

> **For step-by-step instructions on standing up a new customer, see the
> [Customer Onboarding Guide](./ONBOARDING.md).**

---

*This document is meant to be a living overview. When you add a feature, update the
[Feature Inventory](#10-feature-inventory--status); when you add a config key or table, note it in the
[Data Model](#5-data-model).*
