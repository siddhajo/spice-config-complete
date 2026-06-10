# Customer Onboarding Guide

> **Who this is for:** the developer/operator setting up **Spice Config** for a new auction-house customer.
> **Companion doc:** [ARCHITECTURE.md](./ARCHITECTURE.md) explains how the app works internally.

---

## The one thing to understand first

**One customer = one deployment = one database file.**

Spice Config is **not** multi-tenant. There is no "Add Company" button and no customer/tenant ID in the
data. Every customer gets their **own running server** and their **own `config.db`**. Their license,
users, settings, and all auction data live in that single file, completely isolated from every other
customer.

> A customer who legally operates as **two entities** (a main company + a sister concern) is still **one
> deployment** — that case is handled by the **ISP/ASP preset toggle** inside the app, not by a second
> install. See [step 5](#step-5--optional-set-up-the-sister-company-ispasp).

So "onboarding a customer" means: **stand up a new install, license it, and configure their company.**

---

## Onboarding at a glance

```
  1. Deploy a new install          →  fresh server + persistent data volume
  2. First boot (automatic)        →  generates install ID, 30-day trial,
                                       default admin, ~194 default settings
  3. Issue a license               →  sign a token for their install ID
  4. Configure the company         →  Settings: identity, addresses, rates, flags
  5. (Optional) sister company     →  set up the ASP preset
  6. Create user accounts          →  one per staff member, by role
  7. Import existing data          →  traders, buyers, or legacy system data
  8. Hand-off checklist            →  verify, brief the customer, done
```

Plan for **30–60 minutes** for a standard setup (longer if importing legacy data).

---

## Before you start — checklist

Collect this from the customer first:

- [ ] **Company legal details** — trade name, legal name, PAN, CIN (or partnership name), FSSAI/SBL numbers.
- [ ] **GSTINs and addresses** — for each state they operate in (Kerala and/or Tamil Nadu).
- [ ] **Bank details** — account numbers and IFSC for each state.
- [ ] **Rates** — commission %, handling %, GST %, TDS %, TCS %, deductions, gunny rate, etc.
- [ ] **Logo** — PNG/JPG for the invoice letterhead.
- [ ] **Which features** they need — Bills? Debit notes? Pooling? Sister company? Price check?
- [ ] **Staff list** — who needs access and at what role.
- [ ] **Existing data** — do they have a trader/buyer list (Excel) or a legacy system to migrate?

---

## Step 1 — Deploy a new install

Pick the target that fits the customer. Each needs a **persistent location for `data/`** so the database
survives restarts.

### Option A — Railway / cloud (recommended for hosted customers)

1. Create a **new** Railway project/service for this customer (do not reuse another customer's).
2. Deploy from the repo. The `Dockerfile` (sql.js variant) needs no native build.
3. **Add a persistent volume** mounted where the data lives, and point `SPICE_DATA_DIR` at it.
   > Without a persistent volume the database (and license) is wiped on every redeploy.
4. Set environment variables (see table below).

### Option B — Windows desktop (for on-premise customers)

1. Build the installer: `npm run build:win`.
2. Install on the customer's PC. Data lives under the app's data directory locally.

> Before building auto-updating desktop releases, fill in the real GitHub `owner`/`repo` in
> `package.json` → `build.publish` (currently placeholders).

### Required environment variables

| Variable | Set it to | Why |
|----------|-----------|-----|
| `LICENSE_SECRET` | **Your secret signing key** | **Critical.** Must be the *same value* here and on your license-signing machine, or tokens won't validate. Keep it private. |
| `SPICE_DATA_DIR` | Path to the persistent volume | Where `config.db` + backups live |
| `PORT` | e.g. `3001` (or platform default) | HTTP port |
| `LICENSE_TRIAL_DAYS` | `30` (optional) | Trial length, only read on first boot |

---

## Step 2 — First boot (happens automatically)

The very first time the server starts on a fresh `data/` directory, it self-initializes:

| What happens | Result |
|--------------|--------|
| Generates a unique **install ID** (UUID) | The customer's license identity |
| Starts a **30-day trial** | They can log in immediately |
| Creates the schema (all tables) | Empty database, ready to use |
| Seeds **~194 default settings** | Both ISP + ASP defaults present |
| Creates default admin | **`admin` / `admin123`**, flagged to force a password change |

**Verify it's up:** open the app URL. You should see the login page. Log in as `admin` / `admin123` —
you'll be forced to set a new password immediately. **Do this now; the default cannot be reused.**

---

## Step 3 — Issue a license

The customer starts on a 30-day trial. To extend (or to set their paid term):

1. Get their **Install ID**: open `/renew.html` in their app and copy it, or call `GET /api/license/status`.
2. On **your signing machine** (with the same `LICENSE_SECRET`), run:

   ```bash
   LICENSE_SECRET=<your-secret> node tools/license-sign.js --install-id <their-install-id> --days 365
   ```

3. Send the customer the generated **token**.
4. They paste it into `/renew.html` → **Apply**. It takes effect immediately, no restart.

> Renewal works the same way later — they'll see an amber warning pill at ≤7 days and red at ≤3 days.
> Full reference: [LICENSING.md](./LICENSING.md).

---

## Step 4 — Configure the company

Log in as admin/manager and go to **Settings**. Work through the categories with the details you
collected. (Everything here writes to the `company_settings` table; you can also bulk-import via
**Settings → Import**, or `POST /api/company-settings/import`.)

| Settings section | Fill in |
|------------------|---------|
| **Company / Identity** | Trade name, legal name, PAN, CIN *or* partnership name, FSSAI, SBL |
| **Address — Tamil Nadu** | Address, phone, email, GSTIN, branch (used by ISP invoices) |
| **Address — Kerala** | Address, phone, email, GSTIN, branch |
| **Bank details** | Account numbers + IFSC per state |
| **Rates** | Commission %, handling %, GST %, TDS %, TCS %, deductions, gunny rate, transport/insurance |
| **HSN / SAC codes** | Cardamom HSN, gunny HSN, transport/insurance SAC |
| **Season** | Season name + financial-year start/end dates |
| **Invoice settings** | Number prefix, separator, signature text, Tally company names |
| **Feature flags** | Turn on only what they use (Bills, Debit Notes, Pooling, Price Check, etc.) |

**Upload the logo:** Settings → logo upload (or `POST /api/company-settings/logo/ispl`). This is the
letterhead on every invoice/report.

> **Tip:** Set rates and GST/TDS percentages carefully — every invoice and payment figure is computed
> from them. Generate one test invoice afterwards and eyeball the numbers.

---

## Step 5 — (Optional) Set up the sister company (ISP/ASP)

Only if the customer trades under **two identities**:

1. In Settings, use the **Logo Code** dropdown to switch to the **ASP** preset.
2. Enter the sister company's identity (trade name, legal name, PAN/CIN, GSTIN, addresses) — these save
   to the ASP preset.
3. Switch back to **ISP** for normal operation.

Once configured, Kerala-state invoices automatically use ASP and Tamil-Nadu use ISP. A manager can flip
the active company anytime with the same dropdown (requires the `state_toggle` capability).

---

## Step 6 — Create user accounts

Go to **Settings → Users** (admin only). Create one account per staff member and assign the **lowest
role that lets them do their job**:

| Give them… | If they… |
|------------|----------|
| **viewer** | only need to look at data and download reports |
| **lot_entry** | enter lots on the auction floor (often from the mobile app) |
| **operator** | do daily back-office work — invoices, buyers, sellers |
| **manager** | configure settings, revert mistakes, oversee branches |
| **admin** | manage users, do bulk deletes, run backups |

Each new user is created with a temporary password and forced to change it on first login. (Role
capabilities are listed in [ARCHITECTURE.md → Roles](./ARCHITECTURE.md#6-users-roles--permissions).)

---

## Step 7 — Import existing data

Three import paths, depending on what the customer has:

| They have… | Use… | Where |
|------------|------|-------|
| A list of **sellers** in Excel | Traders import | Traders → Import (download the template first) |
| A list of **buyers** in Excel | Buyers import | Buyers → Import (download the template first) |
| A whole **legacy system's data** | Legacy import | Old-data import: **Preview → Verify → Run** (and **Undo** if needed) |

The legacy importer is safe to explore — it previews and verifies before writing, keeps a history, and
can undo a run.

---

## Step 8 — Hand-off checklist

Before you call it done, verify:

- [ ] Logged in successfully; **default admin password changed**.
- [ ] License applied; `/renew.html` shows the correct expiry.
- [ ] Company identity, addresses, GSTINs, and bank details are correct.
- [ ] Logo appears on a generated invoice.
- [ ] Rates produce correct figures on **one test trade → lot → invoice** (then delete the test).
- [ ] Feature flags match what the customer actually uses.
- [ ] (If applicable) ASP sister company configured and switchable.
- [ ] Staff accounts created with appropriate roles.
- [ ] Existing traders/buyers imported and spot-checked.
- [ ] **Backup taken** (Settings → Backup & Restore → download) and stored safely.
- [ ] Customer briefed on: the daily workflow, the mobile lot-entry app, and the license-renewal pill.

---

## Day-2 operations (quick reference)

| Task | How |
|------|-----|
| **Renew a license** | Get install ID → sign a token → customer applies it ([step 3](#step-3--issue-a-license)) |
| **Back up** | Settings → Backup & Restore → download `config.db` (or just copy the file) |
| **Restore** | Settings → Backup & Restore → upload a snapshot |
| **Reset a forgotten admin password** | Set `RESET_ADMIN_PASSWORD` env var, restart |
| **Change rates/settings** | Settings (manager+) — takes effect on the next document generated |
| **White-label look** | `/admin/branding` (developer-gated) — lock a preset, set colors, hide appearance controls |
| **Move a customer to a new host** | Copy their `data/config.db` to the new install's data dir |

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|--------------|-----|
| Login returns **451** | License expired | Issue a renewal token ([step 3](#step-3--issue-a-license)) |
| Applied token is **rejected** | `LICENSE_SECRET` differs between signer and server, or wrong install ID | Make the secret identical; re-copy the install ID |
| Data **disappears** after redeploy | No persistent volume for `data/` | Mount a volume and set `SPICE_DATA_DIR` |
| Server **won't start** ("locked") | Another instance owns the DB (`server.lock`) | Ensure only one server process; remove a stale lock if the old process is truly dead |
| Wrong company on Kerala invoices | ASP preset not configured | Configure the ASP preset ([step 5](#step-5--optional-set-up-the-sister-company-ispasp)) |
| Invoice figures look wrong | Rates / GST / TDS misconfigured | Re-check the Rates section in Settings |

---

*Keep this guide in sync with the app: when onboarding steps or settings change, update the relevant
step above. For the internal architecture, see [ARCHITECTURE.md](./ARCHITECTURE.md).*
