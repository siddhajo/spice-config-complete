# Feature Flags — Wiring Audit

> Generated 2026-06-22. A complete, citation-backed map of every `flag_*` setting:
> whether it is actually wired to working functionality, what it gates, and how it
> behaves in **e-Auction** vs **e-Trade** mode. Re-verify line numbers before relying
> on them — the code moves.

## How feature flags work

- **Definition.** All flags are seeded in [company-config.js](company-config.js) under
  `category: 'flags'` (plus two booking flags under `category: 'booking'`). They appear as
  editable toggles in **Settings → Feature Flags**.
- **Frontend gating.** [`applyFeatureFlags()`](public/index.html#L9713) maps each live flag to a
  `body[data-feat-*]` attribute. CSS rules like `body:not([data-feat-X="1"]) .feat-X{display:none}`
  ([index.html:946-972](public/index.html#L946)) hide tagged elements. Three flags are **AND'ed with
  `business_mode === 'e-Auction'`** so they vanish in e-Trade.
- **Server gating.** Some flags are read directly in [calculations.js](calculations.js),
  [invoice-pdf.js](invoice-pdf.js), [tally-xml.js](tally-xml.js) and [server.js](server.js) to change
  calculations, PDF rendering, or block endpoints.
- **Mode hiding of the toggle itself.** `_MODE_HIDE_KEYS` ([index.html:9613](public/index.html#L9613))
  hides certain flag *toggles* from the Settings screen depending on mode — separate from whether the
  flag does anything.

> ⚠️ **Two important gotchas, repeated below per-flag:**
> 1. Most UI-only flags hide the button/tab but **do NOT disable the backing server endpoint** — the
>    data and API stay reachable.
> 2. Several flags are **seeded and editable but completely dead** — nothing reads them. See
>    [Dead flags](#dead-flags-defined--editable-but-nothing-reads-them).

---

## e-Auction — wired flags

| Flag | Default | Wired to | Layer | Notes |
|---|---|---|---|---|
| `flag_price_list_mapping` | true | "Price List Mapping" sidebar tab ([index.html:2608](public/index.html#L2608)) + Lots-toolbar button ([:3249](public/index.html#L3249)) | Frontend only | AND'ed with e-Auction at [index.html:9730](public/index.html#L9730). Server `/api/price-list/map-*` stays live regardless ([server.js:4711](server.js#L4711)). Toggle hidden in e-Trade. |
| `flag_crop_receipt` | false | "Crop Receipt" Lot Entry field ([index.html:5334](public/index.html#L5334)) + mobile PWA ([mobile-bridge.js:587](mobile-bridge.js#L587)) | Frontend + mobile | AND'ed with e-Auction at [index.html:9725](public/index.html#L9725). Also force-shown by `show_extra_lot_fields`. Server persists `crpt` unconditionally. |
| `flag_reserved_price` | false | "Reserved Price" Lot Entry field ([index.html:5335](public/index.html#L5335)) + mobile PWA ([mobile-bridge.js:588](mobile-bridge.js#L588)) | Frontend + mobile | AND'ed with e-Auction at [index.html:9726](public/index.html#L9726). Server persists `reserved_price` unconditionally so a flag flip never wipes values ([server.js:4318](server.js#L4318)). |

---

## e-Trade — wired flags

| Flag | Default | Wired to | Layer | Notes |
|---|---|---|---|---|
| `flag_dispatch` | true | "Dispatch From" block on the sales-invoice PDF ([invoice-pdf.js:1200](invoice-pdf.js#L1200)) | Server (PDF) | **Only consulted on ASP invoices** (`isASP` = e-Trade + Kerala, [invoice-pdf.js:844](invoice-pdf.js#L844)). For ISP invoices dispatch is always shown. |
| `flag_disc_gst` | false | GST on the trade-credit Discount: lot calc ([calculations.js:258](calculations.js#L258)), debit-note GST ([calculations.js:1208](calculations.js#L1208), [server.js:7689](server.js#L7689)), trade-grid tax columns ([index.html:14380](public/index.html#L14380)) | Server calc + DN render + frontend | Lot-calc path is inside the e-Trade branch. Toggle hidden in e-Auction via `_MODE_HIDE_KEYS` ([index.html:9621](public/index.html#L9621)). |
| `flag_discount_in_prate` | false | Rolls Discount into P_Rate for **Grade 1 lots** ([calculations.js:102](calculations.js#L102), [:713](calculations.js#L713)) + Lots "Set Grade 1" bulk button ([index.html:9749](public/index.html#L9749)) | Server calc + frontend | e-Trade branch only (`ispLotDiscount` returns early otherwise, [calculations.js:708](calculations.js#L708)). Toggle hidden in e-Auction. |
| `flag_print_purchase` | true | "Print Selected Purchase" button in Sales toolbar ([index.html:9699](public/index.html#L9699)) | Frontend only | Shown only when `isASP` (e-Trade + Kerala) **and** flag ON. The purchase-mirror PDF route itself is not flag-gated. |

---

## Both modes — wired flags

These are not guarded by `business_mode`; they work in e-Auction and e-Trade alike.

### Frontend UI gates (sidebar / buttons)

| Flag | Default | Wired to | ⚠️ Endpoint still live? |
|---|---|---|---|
| `flag_bills` | true | Bills of Supply sidebar tab ([index.html:2636](public/index.html#L2636)) + delete-all button ([:2939](public/index.html#L2939)) | Yes — `/api/bills*` stays live |
| `flag_debit_notes` | true | Debit Notes sidebar tab ([index.html:2640](public/index.html#L2640)) + delete-all button ([:2940](public/index.html#L2940)) | Yes — `/api/debit-notes*` stays live |
| `flag_bos_purchase_bill` | true | Swaps Bills toolbar between "Print Selected" purchase-bill ([index.html:3861](public/index.html#L3861)) and "Commission Bill" ([:3862](public/index.html#L3862)); inverted CSS pair ([:955-956](public/index.html#L955)) | n/a (chooses which print action) |
| `flag_debit_note_planter` | true | Planter Debit Notes sidebar tab ([index.html:2644](public/index.html#L2644)) **AND** server 403 gate on generate ([server.js:8064](server.js#L8064), mounted [:8125](server.js#L8125)/[:8218](server.js#L8218)) | **No** — server blocks generate/bulk when OFF |
| `flag_whatsapp` | true | All WhatsApp share buttons ([index.html:3546](public/index.html#L3546), [:4984](public/index.html#L4984), per-row icon) | n/a (client share links) |
| `flag_set_buyer` | true | Lots "Set Buyer" bulk button ([index.html:3240](public/index.html#L3240)) | Yes — `POST /api/lots/bulk-buyer` ([server.js:4534](server.js#L4534)) only checks `requireLotWrite` |

> The Bills / Debit Notes / Planter tabs also carry `state-kerala-only`, an orthogonal **state**
> gate (hidden in Tamil Nadu / ISP context) — independent of the feature flag.

### Server-side calculation / render gates

| Flag | Default | Wired to | Notes |
|---|---|---|---|
| `flag_hsn` | true | HSN/SAC column on invoice PDFs ([invoice-pdf.js:326](invoice-pdf.js#L326), [:1229](invoice-pdf.js#L1229), [:1983](invoice-pdf.js#L1983)) | OFF drops the column; Description spans its width. |
| `flag_bank` | true | Bank-details block on invoice PDF ([invoice-pdf.js:1768](invoice-pdf.js#L1768)) | OFF skips the whole block. |
| `flag_invoice_stripe` | true | Alternating zebra-stripe rows in invoice tables ([invoice-pdf.js:380](invoice-pdf.js#L380), [:917](invoice-pdf.js#L917), [:2037](invoice-pdf.js#L2037)) | Cosmetic only. |
| `flag_tds_purchase` | true | 194Q TDS on purchase invoices: calc ([calculations.js:664](calculations.js#L664)) + Tally voucher block ([tally-xml.js:1765](tally-xml.js#L1765)) | Purchase path. OFF → TDS = 0 and Tally TDS ledger omitted. |
| `flag_wgst` | false | Whether the 194Q TDS basis is with-GST (`total`) or pre-GST (`amount`) ([calculations.js:657](calculations.js#L657), [:665](calculations.js#L665)) | Only matters when `flag_tds_purchase` is ON. |
| `flag_round` | true | Rounds payable/purchase amounts to whole rupees: seller-payment/NEFT calc ([calculations.js:967](calculations.js#L967), [:1109](calculations.js#L1109)), payment export + RTGS threshold ([exports.js:394](exports.js#L394)), per-lot payable distribution ([server.js:8830](server.js#L8830)) | Not mode-guarded, but all call sites are seller-payment / purchase-payment flows. |

### `flag_price_check` (default OFF) — partially wired, soft gate

- **Server side is live.** `pcFlagOn(db)` reads the flag ([server.js:5135](server.js#L5135)). When ON, a passing
  price-check stamps `price_checked_at`; lot mutations clear it to "stale" ([server.js:4353](server.js#L4353) etc.),
  driving an amber banner on the Lots tab ([index.html:20122](public/index.html#L20122)).
- **It is a SOFT stamp + banner, not a hard block** — despite the "transaction gate" label, no endpoint
  rejects a transaction when the price check is stale/pending.
- **Frontend tab gate is DEAD.** `data-feat-price-check` is hardcoded `"1"` on `<body>`
  ([index.html:2422](public/index.html#L2422)) and `applyFeatureFlags()` never toggles it, so the Price Check
  tab is **always visible** regardless of the flag.
- Not mode-guarded (works in both modes in code, though it's auction-oriented in practice).

### Booking alerts (category `booking`) — both modes, server-only

| Flag | Default | Wired to | Companion settings (all read) |
|---|---|---|---|
| `flag_booking_limit` | false | Per-seller booking-limit WhatsApp alerts on lot save: `evaluateBookingLimit()` ([server.js:2253](server.js#L2253), gate [:2255](server.js#L2255)); fires from `POST /api/lots` ([:4364](server.js#L4364)) and `PUT /api/lots/:id` ([:4471](server.js#L4471)); preview `GET /api/booking/status/:auctionId` ([:2515](server.js#L2515)) | `booking_planned_weight_mt`, `booking_soft_pct`, `booking_escalate_pct`, `booking_manager_wa`, `booking_superior_wa` |
| `flag_grade2_limit` | false | Whole-trade Grade-2 share WhatsApp alerts on lot save: `evaluateGrade2Share()` ([server.js:2394](server.js#L2394), gate [:2396](server.js#L2396)); fires from lot create ([:4369](server.js#L4369)) / edit ([:4474](server.js#L4474)); preview `GET /api/booking/grade2-status/:auctionId` ([:2528](server.js#L2528)) | `grade2_soft_pct`, `grade2_escalate_pct`, `grade2_grade_value` (+ reuses manager/superior WA) |

> Not guarded by `business_mode` anywhere — enforcement runs in both modes. The enforcement +
> WhatsApp side is fully server-side and live; no shipped front-end currently surfaces the
> `booking_alert` / `grade2_alert` fields returned by the lot-save responses.

---

## Dead flags (defined & editable, but nothing reads them)

These are seeded into the DB and appear as toggles in **Settings → Feature Flags**, but **no code
reads them** — flipping them does nothing. Candidates for removal or wiring.

| Flag | Label | Why it's dead |
|---|---|---|
| `flag_pooling` | Pooling (Single State) | No reads. Pooler logic is driven by `deduction1` / `deduction1_inclusive` rate fields instead. |
| `flag_sister` | Sister Concern Active | No reads. Sister-concern behavior keys off `business_state === 'KERALA'` / `isASP` ([invoice-pdf.js:818](invoice-pdf.js#L818)). Toggle hidden in e-Auction. |
| `flag_tnpa` | ASP Ship To Address | No reads anywhere. |
| `flag_sample` | Discount in Invoice | No reads; discount logic uses `flag_disc_gst` / `flag_discount_in_prate`. |
| `flag_ship` | Show Ship To Address | **Overridden** — `showShipTo` is hardcoded `isASP ? false : true` ([invoice-pdf.js:843](invoice-pdf.js#L843)); `cfg.flag_ship` is never read. |
| `flag_tds_sales` | TDS on Sales Invoice | No reads. |
| `flag_rtds_inv` | TDS in ASP Purchase | No reads. |
| `flag_debit_note` | Debit Note for Discount | No reads. **Superseded** — DN tax treatment uses `flag_disc_gst`; the DN module/UI uses `flag_debit_notes` / `flag_debit_note_planter`. (Beware: grep "matches" here are just substrings of those two flags.) |
| `flag_dummy` | Allow Dummy Invoices | No reads (dummy-invoice text handled elsewhere). |
| `flag_eway` | ASP eWay Bill / Transport | No reads. |
| `flag_export` | Export Invoices | No reads; `exports-pdf.js` has zero flag refs. |

---

## Quick index

- **e-Auction-only wired:** `flag_price_list_mapping`, `flag_crop_receipt`, `flag_reserved_price`
- **e-Trade-only wired:** `flag_dispatch`, `flag_disc_gst`, `flag_discount_in_prate`, `flag_print_purchase`
- **Both modes wired (UI):** `flag_bills`, `flag_debit_notes`, `flag_bos_purchase_bill`, `flag_debit_note_planter`, `flag_whatsapp`, `flag_set_buyer`
- **Both modes wired (server):** `flag_hsn`, `flag_bank`, `flag_invoice_stripe`, `flag_tds_purchase`, `flag_wgst`, `flag_round`, `flag_price_check` (soft), `flag_booking_limit`, `flag_grade2_limit`
- **Server-enforced (not just UI):** `flag_debit_note_planter` (403), `flag_booking_limit`, `flag_grade2_limit`
- **Dead (11):** `flag_pooling`, `flag_sister`, `flag_tnpa`, `flag_sample`, `flag_ship`, `flag_tds_sales`, `flag_rtds_inv`, `flag_debit_note`, `flag_dummy`, `flag_eway`, `flag_export`
</content>
</invoke>
