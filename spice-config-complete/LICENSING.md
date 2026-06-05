# Licensing Guide

Time-bombed access control for this app — each deployment runs for **30 days
from first boot**, then the developer issues a signed token to extend it. This
document is the single reference for the licence system: setup, daily
renewals, testing, and troubleshooting.

---

## Table of contents

1. [How it works](#1-how-it-works)
2. [One-time Railway setup](#2-one-time-railway-setup)
3. [Minting a renewal token](#3-minting-a-renewal-token)
4. [The customer renewal flow](#4-the-customer-renewal-flow)
5. [Test playbook](#5-test-playbook) — local + Railway + guardrails
6. [Wipe and bootstrap (rarely needed)](#6-wipe-and-bootstrap-rarely-needed)
7. [Troubleshooting](#7-troubleshooting)
8. [API reference](#8-api-reference)
9. [Remote admin (no shell)](#9-remote-admin-no-shell)
10. [Security notes](#10-security-notes)
11. [File map](#11-file-map)

---

## 1. How it works

### State

Single-row SQLite table `license_state` (see [db.js](db.js)) bootstrapped on
the server's first boot:

| column            | meaning                                                |
|-------------------|--------------------------------------------------------|
| `install_id`      | UUID generated once per install. Binds tokens.         |
| `first_seen_at`   | Server's wall clock the first time it booted.          |
| `expires_at`      | When the licence currently runs out.                   |
| `active_token`    | Most recently applied token (audit trail).             |

On first boot `first_seen_at = now` and `expires_at = now + 30 days`. The row
is **sticky** — subsequent boots load it and the values do **not** reset
(provided the database file persists across deploys).

### Tokens

A licence token is an HMAC-SHA256 signed payload:

```
<base64url(JSON({install_id, expires_at, issued_at, note?}))>.<hex_sig>
```

The signing secret is `LICENSE_SECRET` (env var). The same secret must be set
on both:

- **The developer's laptop** — to mint tokens via the CLI.
- **The Railway deployment** — to verify them at apply time.

### Gates

- `POST /api/login` returns **HTTP 451** with `{error:'license_expired', install_id, expires_at}` once `expires_at < now`. The user is automatically redirected to `/renew.html`.
- `GET /api/license/status` and `POST /api/license/apply` are **auth-free** so a locked-out operator can still post a renewal token from a logged-out state.
- The frontend topbar shows a warning pill at **≤ 7 days** (amber) and **≤ 3 days** (red). After expiry it turns into `⚠ License expired — renew`.

### Threat model

Honest-customer protection only. Anyone with the source code can patch the
gate out. The point of the signature is to stop a customer from forging a
token themselves, not to defeat a determined attacker.

---

## 2. One-time Railway setup

You only do this once per deployment. Skipping any step makes the system
either insecure or impractical to renew.

### Step 1 — Generate a signing secret

On your laptop:

```sh
node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"
```

This prints a 64-character hex string, e.g.

```
7a4f9e23b15c8d6e0a1b2c3d4e5f6071829304a5b6c7d8e9f0a1b2c3d4e5f6071
```

**Save this string in your password manager.** You need the exact same value
on your laptop (to mint tokens) and on Railway (to verify them). If both
copies are lost, no future tokens will verify and the only recovery is to
ship a new app version with a new fallback secret.

### Step 2 — Set `LICENSE_SECRET` on Railway

In **Railway → your service → Variables**:

| key              | value                                |
|------------------|--------------------------------------|
| `LICENSE_SECRET` | (paste the secret from step 1)       |

Save. Railway will trigger a redeploy. After it comes up, the boot log
should NOT contain the `⚠ LICENSE_SECRET env var is not set` warning. If
it still does, the variable isn't being read — check that you added it to
the right service.

### Step 3 — Attach a persistent volume

By default Railway's filesystem resets on every deploy. If `data/config.db`
lives on the ephemeral disk, every push wipes the licence state and the
trial restarts — defeating the licence entirely.

In **Railway → your service → Volumes**, attach a volume mounted at the
directory holding the SQLite file. If `DB_PATH=/data/config.db`, mount the
volume at `/data`. Verify with:

```sh
railway shell
ls -la /data/
exit
```

You should see `config.db` on the volume after the first boot.

### Step 4 — Verify install ID + expiry

```sh
curl https://<your-app>.railway.app/api/license/status
```

Expected response:

```json
{
  "install_id": "8a3f7b91-e5c6-d4a2-098f-1b6e3d7c5a40",
  "first_seen_at": "2026-06-01T09:14:22.317Z",
  "expires_at": "2026-07-01",
  "days_remaining": 30,
  "expired": false,
  "has_token": false,
  "trial_days": 30
}
```

Save the install ID somewhere — that's the canonical identity of this
deployment. Every future token you mint for this customer references it.

### Step 5 — Set the same secret on your laptop

Add to your shell rc (or run before every minting session):

```sh
export LICENSE_SECRET="<the same string you put in Railway Variables>"
```

A quick way to confirm both sides agree: mint a 1-day test token, apply it,
then immediately mint a fresh 30-day one to restore the normal expiry.

---

## 3. Minting a renewal token

### Standard 30-day renewal

```sh
node tools/license-sign.js --install-id <customer-install-id> --days 30
```

Output goes to **two streams**:

- **stderr** prints the human confirmation (install ID, expiry, issued-at, note).
- **stdout** prints the token itself on a single line.

That separation is deliberate — you can pipe the token straight into the
clipboard without picking up the confirmation block.

### Common variants

```sh
# With an audit note — baked into the payload, visible to anyone who
# decodes the token (so keep it short and non-secret).
node tools/license-sign.js \
  --install-id 8a3f7b91-e5c6-d4a2-098f-1b6e3d7c5a40 \
  --days 30 \
  --note "Acme Co — June renewal #INV-2042"

# Specific end date instead of N days.
node tools/license-sign.js \
  --install-id <id> \
  --until 2026-12-31

# Three-month renewal.
node tools/license-sign.js --install-id <id> --days 90

# Copy straight to clipboard (macOS).
node tools/license-sign.js --install-id <id> --days 30 | pbcopy

# Copy to clipboard (Linux + xclip).
node tools/license-sign.js --install-id <id> --days 30 | xclip -selection clipboard

# Save to a file for record-keeping.
node tools/license-sign.js --install-id <id> --days 30 > "tokens/$(date +%F)-acme.token"
```

### CLI options

| flag                  | meaning                                                            | required                 |
|-----------------------|--------------------------------------------------------------------|--------------------------|
| `--install-id <uuid>` | Target install. From customer's `/renew.html` or `/api/license/status`. | yes                      |
| `--days <N>`          | Token expires N calendar days from today.                         | one of `--days`/`--until`|
| `--until <YYYY-MM-DD>`| Explicit calendar end date.                                       | one of `--days`/`--until`|
| `--note <text>`       | Free-form audit note baked into the payload.                       | no                       |
| `-h`, `--help`        | Print usage.                                                       | no                       |

The CLI refuses zero/negative `--days`, malformed `--until`, missing
`--install-id`, and the combination of `--days` + `--until` together.

### Tip: keep a renewal log

Tokens are bound to one install, so you can keep them in version control
or a private file without compromising other deployments:

```
tokens/
  2026-06-01-acme.token       # what you sent
  2026-06-01-acme.note.md     # invoice ref, contact, amount, etc.
  2026-07-01-acme.token
  ...
```

### Delivery to the customer

The token is just text. WhatsApp, email, Slack, SMS — anything that
round-trips the characters intact works. If the channel autoformats
URLs or adds line breaks (some email clients), copy the customer's
displayed string into the renewal page and the server tells you
`bad signature` — re-send via a plaintext channel.

---

## 4. The customer renewal flow

### Before expiry — warning pill

When the licence is within the warning window:

- **≤ 7 days, > 3 days** → amber pill: `⏳ N days left`.
- **≤ 3 days** → red pill: `⚠ N days left`.

Clicking the pill opens `/renew.html`. From there the customer:

1. Clicks **Copy** to grab the install ID.
2. Sends it to you via WhatsApp / email.
3. Receives your token.
4. Pastes it into the textarea.
5. Clicks **Apply renewal token**.
6. Page shows `✓ Token applied. New expiry: <date>` and bounces to login.

### After expiry — hard block

`POST /api/login` returns 451. The login form catches it
([public/index.html:6420](public/index.html#L6420)) and redirects to
`/renew.html?install_id=…&expired=…` with the ID prefilled.

The renewal page works the same way — paste, apply, log in.

### What the customer never has to do

- Type the install ID by hand (Copy button on the page).
- Edit a config file or env var.
- Restart anything (the new expiry takes effect immediately for every
  subsequent request).
- Worry about old tokens — the server refuses any token whose `expires_at`
  is earlier than the current one, so a stale token can't shorten the
  active licence.

---

## 5. Test playbook

Run the four tracks in order. Each one isolates a slice of the system;
don't move to the next until the previous passes. Total time end-to-end:
~15 min.

> Every command below has the URL prefix templated as `$URL` so you can
> run the same block twice — once with `URL=http://localhost:3000`
> (Track A) and once with `URL=https://<your-app>.railway.app` (Track
> B). Or just paste the URL inline.

### Track A — Local sanity check (≈ 5 min)

Proves the code works on your laptop before you touch production. Use a
real (non-fallback) secret so the admin endpoint is exercisable.

#### A1. Start the server

```sh
LICENSE_SECRET=local-test-secret-123 npm start
```

✅ Boot logs do NOT contain `⚠ LICENSE_SECRET env var is not set`.

#### A2. Read the install ID

```sh
export LICENSE_SECRET=local-test-secret-123      # for subsequent steps
export URL=http://localhost:3000
curl -s "$URL/api/license/status" | python3 -m json.tool
```

Save the `install_id` value as `$IID`. ✅ checkpoint: `days_remaining: 30`,
`expired: false`, `has_token: false`.

#### A3. Mint + apply a token (the customer's normal flow)

```sh
TOKEN=$(node tools/license-sign.js --install-id "$IID" --days 60 2>/dev/null)
curl -s -X POST "$URL/api/license/apply" \
  -H 'content-type: application/json' \
  -d "{\"token\":\"$TOKEN\"}" | python3 -m json.tool
```

✅ checkpoint: `ok: true`, response `status.has_token: true`,
`status.days_remaining: 60`.

#### A4. Force a warning state (via the admin endpoint)

```sh
curl -s -X POST "$URL/api/license/admin/set-expiry" \
  -H 'content-type: application/json' \
  -H "X-License-Secret: $LICENSE_SECRET" \
  -d "{\"expires_at\":\"$(date -v+3d +%F 2>/dev/null || date -d '+3 days' +%F)\"}"
```

Open `$URL` in a browser, reload.

✅ checkpoint: topbar shows the red `⚠ 3 days left` pill.

#### A5. Force hard expiry + verify login block

```sh
curl -s -X POST "$URL/api/license/admin/set-expiry" \
  -H 'content-type: application/json' \
  -H "X-License-Secret: $LICENSE_SECRET" \
  -d '{"expires_at":"2020-01-01"}'
```

Reload the browser and try to log in.

✅ checkpoint: page redirects to `/renew.html?install_id=…&expired=2020-01-01`
with the install ID prefilled.

#### A6. Renewal round-trip

```sh
node tools/license-sign.js --install-id "$IID" --days 30 \
  | pbcopy 2>/dev/null \
  || node tools/license-sign.js --install-id "$IID" --days 30 \
       | xclip -selection clipboard
```

Paste into the renewal page → **Apply renewal token**.

✅ checkpoint: green ✓ confirmation, redirect to login, log in works,
topbar pill is hidden.

---

### Track B — Railway end-to-end (≈ 5 min)

Proves the production deploy works. **Skip if Track A failed.**

#### B1. Verify the secret is set on Railway

This is the first call that proves whether your laptop's secret matches
the deploy's secret. If it doesn't, no minted token will verify and no
admin call will succeed.

```sh
export URL=https://<your-app>.railway.app
export LICENSE_SECRET="<the value you stored in your password manager>"
curl -s -X POST "$URL/api/license/admin/set-expiry" \
  -H 'content-type: application/json' \
  -H "X-License-Secret: $LICENSE_SECRET" \
  -d "{\"expires_at\":\"$(date -v+45d +%F 2>/dev/null || date -d '+45 days' +%F)\"}"
```

| Response                                                            | What it means                                       |
|---------------------------------------------------------------------|-----------------------------------------------------|
| `{"ok": true, "status": …}`                                         | ✅ secret matches, ready to test                    |
| `403 admin disabled: LICENSE_SECRET env var is not set on this server` | Env var missing on Railway. Fix §2 step 2.        |
| `403 invalid X-License-Secret`                                       | Your local `$LICENSE_SECRET` ≠ the Railway value.  |
| `403 X-License-Secret header required`                               | The header didn't get sent — quoting / shell typo. |

#### B2. Read the prod install ID

```sh
curl -s "$URL/api/license/status" | python3 -m json.tool
```

Save the prod `install_id` as `$IID_PROD`. ✅ checkpoint: `expires_at`
matches what you just set in B1 (`+45d`).

#### B3. Force warning state on prod

```sh
curl -s -X POST "$URL/api/license/admin/set-expiry" \
  -H 'content-type: application/json' \
  -H "X-License-Secret: $LICENSE_SECRET" \
  -d "{\"expires_at\":\"$(date -v+3d +%F 2>/dev/null || date -d '+3 days' +%F)\"}"
```

Open `$URL` in an incognito window (so it's a fresh page load).

✅ checkpoint: topbar shows red `⚠ 3 days left`.

#### B4. Force hard expiry + verify login block

```sh
curl -s -X POST "$URL/api/license/admin/set-expiry" \
  -H 'content-type: application/json' \
  -H "X-License-Secret: $LICENSE_SECRET" \
  -d '{"expires_at":"2020-01-01"}'
```

Reload, try to log in.

✅ checkpoint: redirect to `/renew.html` with `$IID_PROD` prefilled.

#### B5. Renewal round-trip

```sh
node tools/license-sign.js --install-id "$IID_PROD" --days 30 \
  | pbcopy 2>/dev/null \
  || node tools/license-sign.js --install-id "$IID_PROD" --days 30 \
       | xclip -selection clipboard
```

Paste into the renewal page → Apply.

✅ checkpoint: green ✓ confirmation, log in works, topbar pill hidden.

#### B6. Persistence check (volume sanity)

Trigger a Railway redeploy (push a comment-only change, or click
"Redeploy" in the dashboard). After it boots:

```sh
curl -s "$URL/api/license/status" | python3 -m json.tool
```

✅ checkpoint: `install_id` is unchanged AND `expires_at` still reflects
the 30-day token from B5.

If the install_id changed → the persistent volume isn't attached, your
`data/config.db` lives on ephemeral storage, and the trial would reset
on every deploy. Fix per §2 step 3.

---

### Track C — Guardrails (≈ 3 min)

These all should **fail** with the indicated error. Run against either
local or Railway (they share the same code paths).

| #  | Test                | Setup                                                                                       | Expected response                                                                              |
|----|---------------------|---------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------|
| C1 | Forged signature    | `T=$(node tools/license-sign.js --install-id "$IID" --days 30); FORGED="${T%?}X"` then apply `$FORGED` | `400 Invalid token: bad signature`                                                             |
| C2 | Wrong install ID    | `node tools/license-sign.js --install-id NOT-REAL --days 30` then apply                     | `400 Token was issued for a different install. Send the dev your install ID: <real-id>`        |
| C3 | Past-date token     | `node tools/license-sign.js --install-id "$IID" --until 2020-01-01` then apply              | `400 This token has already expired (2020-01-01). Ask the dev for a fresh one.`                |
| C4 | Shorter expiry      | Current 60d. Apply a 10-day token.                                                          | `400 This token (<near>) would shorten the current license (<later>). Already covered.`         |
| C5 | Admin: no header    | admin curl without `X-License-Secret`                                                       | `403 X-License-Secret header required`                                                          |
| C6 | Admin: wrong secret | admin curl with a wrong header value                                                        | `403 invalid X-License-Secret`                                                                  |
| C7 | Admin: bad date     | Body `{"expires_at":"not-a-date"}`                                                          | `400 expires_at must be YYYY-MM-DD (e.g. 2026-07-31)`                                          |
| C8 | Admin: year typo    | Body `{"expires_at":"1999-12-31"}`                                                          | `400 expires_at year out of range (2020-2099)`                                                  |

Quick example for C1:

```sh
T=$(node tools/license-sign.js --install-id "$IID" --days 30 2>/dev/null)
FORGED="${T%?}X"
curl -s -X POST "$URL/api/license/apply" \
  -H 'content-type: application/json' \
  -d "{\"token\":\"$FORGED\"}"
```

---

### Track D — Restore (do this last, always)

Push expiry back to the customer's actual term so you don't leave the
deploy in a forced-expired state:

```sh
curl -s -X POST "$URL/api/license/admin/set-expiry" \
  -H 'content-type: application/json' \
  -H "X-License-Secret: $LICENSE_SECRET" \
  -d "{\"expires_at\":\"$(date -v+30d +%F 2>/dev/null || date -d '+30 days' +%F)\"}"
```

Or — better, because it exercises the real production path — mint a
real 30-day token and apply it via the renewal page:

```sh
node tools/license-sign.js --install-id "$IID_PROD" --days 30
# → paste the token into /renew.html → Apply.
```

✅ Final check:

```sh
curl -s "$URL/api/license/status" | python3 -m json.tool
```

✅ checkpoint: `expired: false`, `days_remaining: 30`, and (if you used
the token path) `has_token: true`.

---

## 6. Wipe and bootstrap (rarely needed)

If you want to start over with a fresh `install_id` (e.g. you're
testing the install-ID mismatch flow), delete the row and let the
server bootstrap a new one on the next request.

**Locally:**

```sh
sqlite3 data/config.db "DELETE FROM license_state WHERE id = 1;"
# Hit any endpoint — /api/license/status will bootstrap a new row.
```

**On Railway (without shell):** the admin endpoint doesn't support
DELETE intentionally — that would silently change `install_id` and
invalidate every token you've issued to that customer. Use
`railway run sqlite3 …` or temporarily add a DELETE branch to the
admin route if you genuinely need this on prod. Almost no real-world
case requires it; consider whether `set-expiry` covers your need
first.

---

## 7. Troubleshooting

### "Invalid token: bad signature"

The token's HMAC doesn't match. Most common causes:

- **Mismatched `LICENSE_SECRET`** between your laptop and Railway. Re-export
  the right value on your laptop and re-mint.
- **Whitespace damage** during copy/paste (line break inserted, leading
  space added by the email client, etc.). Re-send the token in a
  plaintext channel and have the customer paste exactly as received.
- **Truncation**. Some chat apps cut very long strings; the customer
  pastes only the first half. The token is one line, no spaces — they
  must paste it all.

### "Token was issued for a different install"

You minted with the wrong install ID. The error message includes the
real ID — re-mint against that and re-send.

### "This token has already expired"

You used `--until` with a past date, or sent the customer an old token
saved from a previous renewal. Mint a fresh one.

### "This token would shorten the current license. Already covered."

The customer applied a longer renewal already. No action needed —
they're not expired.

### Customer says "I applied the token but it didn't work"

Have them check `https://<app>/api/license/status` in their browser.
If `expires_at` is updated and `has_token: true`, the token *did* apply —
they may just have an old browser tab open. Refresh.

### Renewal page says "Could not load status"

The server is unreachable. Check the Railway service is up and the
URL is correct. The renewal page itself works from a logged-out state
but still needs the server running.

### Trial keeps resetting after every deploy

The volume isn't attached, so `data/config.db` lives on ephemeral
storage. See [Step 3 above](#step-3---attach-a-persistent-volume).

### Login works even after expiry

Either:
- The `expires_at` is still in the future (check `/api/license/status`).
- The `license` module failed to load on this boot (the gate
  fail-opens rather than bricking the app — check the server logs for
  an error during the licence check).

### `LICENSE_SECRET` warning appears in Railway logs

`process.env.LICENSE_SECRET` is unset or empty. Re-check the variable
in Railway → Variables; it must be named exactly `LICENSE_SECRET` and
attached to the right service.

### I want to change the trial length

Set `LICENSE_TRIAL_DAYS` (env var) before the **first** boot. Existing
installs ignore it because their row already exists.

### I want to change the warning threshold

`LICENSE_WARN_DAYS` env var (default 7). Read by the topbar pill via
`/api/license/status`.

---

## 8. API reference

### `GET /api/license/status`

Auth-free. Returns the current licence state.

```json
{
  "install_id": "8a3f7b91-e5c6-d4a2-098f-1b6e3d7c5a40",
  "first_seen_at": "2026-06-01T09:14:22.317Z",
  "expires_at": "2026-07-01",
  "days_remaining": 27,
  "expired": false,
  "has_token": false,
  "trial_days": 30
}
```

When expired:

```json
{ "...": "...", "expired": true, "days_remaining": -3 }
```

### `POST /api/license/apply`

Auth-free. Body:

```json
{ "token": "<the-signed-token>" }
```

Success (HTTP 200):

```json
{
  "ok": true,
  "status": {
    "install_id": "...",
    "expires_at": "2026-07-01",
    "days_remaining": 30,
    "expired": false,
    "has_token": true,
    "...": "..."
  }
}
```

Failure (HTTP 400):

```json
{ "ok": false, "error": "Invalid token: bad signature" }
```

Possible errors:

- `Invalid token: empty token`
- `Invalid token: malformed token`
- `Invalid token: bad signature`
- `Invalid token: malformed payload`
- `Invalid token: payload missing install_id`
- `Invalid token: payload missing expires_at`
- `Token was issued for a different install. Send the dev your install ID: <id>`
- `This token has already expired (<date>). Ask the dev for a fresh one.`
- `This token (<date>) would shorten the current license (<date>). Already covered.`

### `POST /api/login` — 451 response

When the licence is expired, login returns HTTP **451** instead of 200:

```json
{
  "error": "license_expired",
  "message": "Your access window ended on 2026-04-15. Send the install ID below to your provider to receive a renewal token.",
  "install_id": "8a3f7b91-...",
  "expires_at": "2026-04-15"
}
```

The login form catches this and redirects to `/renew.html`.

---

## 9. Remote admin (no shell)

When you can't open a shell on Railway and you need to read or change
licence state on the live deployment, use the admin endpoint. It's
authenticated by the same `LICENSE_SECRET` you used to sign tokens, so
no extra credential management.

### Why it's safe

- **Refuses to operate when `LICENSE_SECRET` is unset.** The fallback
  development secret in [license.js](license.js) can never double as a
  remote admin key — if you haven't followed [§2 step 2](#step-2--set-license_secret-on-railway)
  and set the env var in Railway, this endpoint returns 403 to every
  call.
- **Constant-time secret compare.** Wrong-secret attempts can't be
  distinguished from each other by response time.
- **Length-mismatch short-circuit.** A header with a different length
  is rejected before the compare so the underlying buffer comparison
  doesn't throw and the secret length isn't leaked.
- **No public surface.** The endpoint is undocumented from the
  customer's point of view — they have no reason to know it exists
  and couldn't use it without `LICENSE_SECRET`.

### Setup on your laptop

Export the secret once per shell session:

```sh
export LICENSE_SECRET="<the exact string you put in Railway Variables>"
```

(You already have this exported for the CLI minting flow — no extra
setup needed.)

### `POST /api/license/admin/set-expiry`

Override `license_state.expires_at` directly. Useful for forcing
warning / expired states during testing, or extending in an emergency
when you can't mint a token right now.

**Headers**

| header              | required | meaning                                                |
|---------------------|----------|--------------------------------------------------------|
| `X-License-Secret`  | yes      | Must equal `LICENSE_SECRET` on the server.             |
| `Content-Type`      | yes      | `application/json`                                     |

**Body**

```json
{ "expires_at": "YYYY-MM-DD" }
```

`expires_at` must be a calendar date in the year range 2020-2099.
Past dates are accepted (that's how you force an expired state).

**Success — HTTP 200**

```json
{
  "ok": true,
  "status": {
    "install_id": "8a3f7b91-...",
    "first_seen_at": "2026-06-01T09:14:22.317Z",
    "expires_at": "2020-01-01",
    "days_remaining": -2342,
    "expired": true,
    "has_token": false,
    "trial_days": 30
  }
}
```

**Failures**

| status | error                                                                          | meaning |
|--------|--------------------------------------------------------------------------------|---------|
| 403    | `admin disabled: LICENSE_SECRET env var is not set on this server`             | The deploy hasn't been wired with the secret yet. Fix per [§2 step 2](#step-2--set-license_secret-on-railway). |
| 403    | `X-License-Secret header required`                                             | You forgot the header.                              |
| 403    | `invalid X-License-Secret`                                                     | Wrong secret. The deploy expects a different value. |
| 400    | `expires_at must be YYYY-MM-DD (e.g. 2026-07-31)`                              | Body missing / malformed.                           |
| 400    | `expires_at year out of range (2020-2099)`                                     | Typo guard.                                         |

### Side effect

A successful `set-expiry` also clears `license_state.active_token`.
That's deliberate — the `active_token` column is the audit trail of
the most recently *applied* token, and a manual override didn't come
from one. If you want token-driven history preserved, mint and apply
through the normal `/api/license/apply` flow instead.

### Recipes

The same scenarios from [§5](#5-testing-on-railway), but without shell access:

```sh
# Read current state from anywhere (no secret needed).
curl https://<app>/api/license/status

# Force expiry to 3 days out → red warning pill.
curl -X POST https://<app>/api/license/admin/set-expiry \
  -H "content-type: application/json" \
  -H "X-License-Secret: $LICENSE_SECRET" \
  -d "{\"expires_at\":\"$(date -v+3d +%F)\"}"     # macOS BSD date

# Linux equivalent:
curl -X POST https://<app>/api/license/admin/set-expiry \
  -H "content-type: application/json" \
  -H "X-License-Secret: $LICENSE_SECRET" \
  -d "{\"expires_at\":\"$(date -d '+3 days' +%F)\"}"

# Hard-expire for the full block test.
curl -X POST https://<app>/api/license/admin/set-expiry \
  -H "content-type: application/json" \
  -H "X-License-Secret: $LICENSE_SECRET" \
  -d '{"expires_at":"2020-01-01"}'

# Restore to ~30 days out.
curl -X POST https://<app>/api/license/admin/set-expiry \
  -H "content-type: application/json" \
  -H "X-License-Secret: $LICENSE_SECRET" \
  -d "{\"expires_at\":\"$(date -v+30d +%F 2>/dev/null || date -d '+30 days' +%F)\"}"
```

Use sparingly — the canonical renewal path is mint-and-apply via
[§3](#3-minting-a-renewal-token) + [§4](#4-the-customer-renewal-flow).
The admin endpoint exists for testing and emergency bumps, not as a
substitute for tokens.

---

## 10. Security notes

### What this protects against

- **Honest customers** who forget to renew or share access with a
  third party past their term. They can't extend themselves without
  contacting you.
- **Casual self-service trial reset** — wiping the DB resets the trial
  but a determined customer who can do that already has shell access
  and can patch the code.

### What this does NOT protect against

- A customer with shell access to the Railway container modifying
  `license.js` or `db.js` directly.
- A customer with the source code patching out the gate locally.
- A customer who clones the repo and runs their own deploy with their
  own `LICENSE_SECRET`.

This is acceptable for the "solo dev → solo customer" model the
licence was designed for.

### Rotating the secret

If `LICENSE_SECRET` ever leaks (e.g. someone took a screenshot of your
terminal):

1. Generate a new 64-char random hex.
2. Update Railway Variables → `LICENSE_SECRET`.
3. Update your laptop's exported value.
4. Mint and apply a fresh token to confirm the new secret works
   end-to-end.

Every previously-issued token immediately fails verification under
the new secret, which is the intended behaviour.

### Where the fallback secret comes from

If `LICENSE_SECRET` is unset, [license.js](license.js) uses
`spice-etrade-license-DEV-ONLY-change-via-LICENSE_SECRET` as the HMAC
key. This is intentionally weak — anyone with the repo can mint
tokens against it. The fallback exists so local development works
without setup, and so a production deploy that forgets to set the
env var still functions (with a loud `⚠ LICENSE_SECRET env var is not
set` warning in the boot log) rather than locking the dev out.

---

## 11. File map

| path                                          | purpose                                                          |
|-----------------------------------------------|------------------------------------------------------------------|
| [license.js](license.js)                      | Token codec + state CRUD. The single source of truth for the licence logic. |
| [tools/license-sign.js](tools/license-sign.js)| Developer-only CLI for minting tokens.                            |
| [db.js](db.js)                                | `license_state` schema migration.                                 |
| [server.js](server.js) (~line 880)            | `/api/license/status`, `/api/license/apply`, `/api/license/admin/set-expiry`, and the 451 gate on `/api/login`. |
| [public/index.html](public/index.html) (~6390)| `_refreshLicenseBadge()` topbar pill + 451 catch in the login flow. |
| [public/renew.html](public/renew.html)        | Standalone renewal page. No dependency on the main app bundle.    |

---

## Quick reference card

| Action                                       | Command                                                                  |
|----------------------------------------------|--------------------------------------------------------------------------|
| Generate a Railway secret                    | `node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"` |
| Read install ID from prod                    | `curl https://<app>/api/license/status`                                  |
| Mint a 30-day token                          | `LICENSE_SECRET=… node tools/license-sign.js --install-id <id> --days 30` |
| Mint with explicit end date                  | `node tools/license-sign.js --install-id <id> --until 2026-12-31`         |
| Mint + copy to clipboard (macOS)             | `node tools/license-sign.js --install-id <id> --days 30 \| pbcopy`        |
| Mint + save to file                          | `node tools/license-sign.js --install-id <id> --days 30 > acme.token`    |
| CLI help                                     | `node tools/license-sign.js --help`                                       |
| **Force expiry remotely** (no shell)         | `curl -X POST https://<app>/api/license/admin/set-expiry -H "content-type: application/json" -H "X-License-Secret: $LICENSE_SECRET" -d '{"expires_at":"2020-01-01"}'` |
| **Restore 30 days remotely** (no shell)      | `curl -X POST https://<app>/api/license/admin/set-expiry -H "content-type: application/json" -H "X-License-Secret: $LICENSE_SECRET" -d "{\"expires_at\":\"$(date -v+30d +%F 2>/dev/null \|\| date -d '+30 days' +%F)\"}"` |
| Force expiry via Railway shell               | `sqlite3 /data/config.db "UPDATE license_state SET expires_at='2020-01-01' WHERE id=1"` |
| Force 3-day warning via Railway shell        | `sqlite3 /data/config.db "UPDATE license_state SET expires_at=date('now','+3 days') WHERE id=1"` |
| Restore 30 days via Railway shell            | `sqlite3 /data/config.db "UPDATE license_state SET expires_at=date('now','+30 days') WHERE id=1"` |
| Wipe licence state (re-bootstrap trial)      | `sqlite3 data/config.db "DELETE FROM license_state WHERE id=1"`           |
