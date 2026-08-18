# HANDOFF — Coupa Receipt Filter

Pick-up notes for a fresh Claude Code session. For the full living reference see
[PROJECT-DOCUMENTATION.md](PROJECT-DOCUMENTATION.md); this file is the "what just
happened / how to continue" layer.

## Current state (as of v0.10.0)

- **Single file:** `coupa-receipt-filter.user.js` — a Tampermonkey/Greasemonkey
  userscript for `*.coupahost.com`. One IIFE, no build step.
- **Branch/remote:** work lands on `main`. `origin` =
  `github.com/SexualMoose/coupa-receipt-filter` (**public**). At v0.9.0 `main`
  is pushed and in sync.
- **Distribution:** the `@updateURL`/`@downloadURL` in the script header point at
  a **GitHub gist**, NOT the repo. Installed copies auto-update from the gist.
  - Gist id: `a0de5a5bf56d33abef414b5781bdd984` (owner `SexualMoose`).
  - The repo is backup/history; **the gist is what users actually run.** A commit
    to `main` is not "shipped" until the gist is updated too (see Ship below).
- **Auth:** `gh` is logged in as `SexualMoose` with `gist` + `repo` scopes.

> Public repo → keep employer/tenant-confidential detail and any vulnerability
> specifics OUT of committed docs. Tenant ids already embedded in the script
> (DEFAULT_ACCOUNT, ALL_CATEGORIES) predate this and are the existing baseline.

## What v0.10.0 added (FX)

Panel currency conversions, not the xlsx round trip:

1. **More display currencies.** `TARGETS` grew from `USD/EUR/COP/SGD/TRY/PLN` to
   `USD/SGD/HKD/TWD/KRW/MYR/IDR/VND/EUR/COP/TRY/PLN` (added the SE-/E-Asia set).
   Receipts in those currencies can now be matched by the ±% filter (previously
   `targets[HKD]` was `undefined` → never matched). The xlsx USD-eq path already
   handled any currency the API returns, so no export change was needed.
2. **Currency-aware formatting** (`fmtMoney` + `ZERO_DP_CCY`): VND/IDR/KRW/COP show
   grouped with no decimals (26,137 not 26137.29); everything else keeps 2 dp.
3. **FX fallback + persistence.** Rates were always live (open.er-api.com, 1h cache)
   — nothing was hardcoded/stale. Added robustness: `getRates`/`fetchFxToUSD` now go
   **live → last successful live pull (persisted via `GM_setValue`, else
   `localStorage`, 30-day trust) → baked `FX_FALLBACK` snapshot** (captured
   2026-08-18, cross-verified vs Google Finance/XE/Wise within 0.31%). The panel
   labels fallback rates in red ("offline rates: …"). New grants: `GM_setValue`,
   `GM_getValue`. To refresh `FX_FALLBACK` later, re-fetch
   `https://open.er-api.com/v6/latest/USD` and paste the units-per-USD values.

## What v0.9.0 added

Two features on the xlsx **Upload & Apply** round trip:

1. **Currency editing.** New editable `new_currency` column in the export
   (dropdown). Coupa's PATCH needs a tenant-specific numeric currency **id**, not
   the ISO code, and there's no hard-coded currency catalog — so ids are
   harvested from the user's own lines into a hidden `_currencies` helper sheet
   (code→id), mirroring the existing `_categories` pattern, and translated back
   on upload. Field set: `expense_line[receipt_total_currency_id]`.

2. **Example attendee column.** Every export ends with a placeholder column
   headed by the `EXAMPLE_ATTENDEE_HEADER` sentinel. Left unchanged/deleted →
   ignored on upload. Renamed to `Firstname Lastname` + marked `x` → the uploader
   creates that attendee (`POST /expense_attendees/`) and attaches them, no
   Attendees-sheet edit needed.

## Code map (search these names)

Export (build the workbook):
- `collectCurrencyDirectory(reports)` — unique `{code,id}` from every line's
  `receipt_currency` + `amount_to_receive_currency`.
- `_buildWorkbookInner(...)` — builds the `Lines` sheet. **Column order is
  1-indexed and hard-referenced** in several places; after v0.9.0:
  `1 line_id … 6 currency, 7 new_currency, 8 usd_eq, 9 current_category,
  10 new_category, 11 problems, 12 current_description, 13 new_description,
  14 current_attendees, 15.. attendee cols, LAST = example col`.
  Watch `ATTENDEE_COL_START = 15`, the widths array (14 entries), the dropdown
  data-validation on `G`/`J`, and the conditional-formatting refs
  (`G2:G10000`, `J2:J10000`, attendee range). Hidden sheets: `_categories`,
  `_currencies`.

Import (parse + apply):
- `parseUploadedWorkbook(file)` — name-based `colIdx()` (robust to column
  reordering); `COL_NEW_CURRENCY` is `0` for pre-0.9.0 workbooks (guarded);
  builds `currencyCodeToId` from the `_currencies` sheet; attendee-column loop
  handles (a) sentinel = ignore, (b) known attendee, (c) `(id)`-suffix typo
  guard, (d) new-attendee-from-header via `getOrCreateNewAttendee`.
- `buildLineUpdatePatchBody(line, change)` — the PATCH form body. Currency logic
  lives in the `chgCur`/`hadForeign`/`receiptCurId`/`foreignCurId`/`exchangeRate`/
  `foreignAmt` block at the top.
- `applyUpload(...)` — `isCrossCurrency(line)` guard, `noOp`/`trimmed` diffing,
  `currencySkipped` reporting, first-line confirmation.

## The one non-obvious design decision (READ before touching currency)

Changing only the receipt currency would leave a **cross-currency** line (receipt
currency ≠ reimbursement currency) with a stale `exchange_rate`/`amount_to_receive`
— a financial-integrity bug. We can't recompute Coupa's official FX rate from the
browser, so:

- Currency edits apply **only to single-currency lines**; `buildLineUpdatePatchBody`
  moves the reimbursement (foreign) side to the same new currency at rate 1 so the
  payload is always internally consistent.
- Genuinely cross-currency lines are **refused** (`isCrossCurrency` → `true`) and
  reported to the user; the change is dropped rather than PATCHed.
- The first-line confirmation dialog flags currency changes so the user verifies
  the amount on line 1 before bulk-applying.

**Open uncertainty (untested):** whether Coupa recomputes the reimbursement amount
server-side on a single-currency currency change, or stores the posted values
blindly. Couldn't verify headless. The first-line confirmation is the human gate.
If you can test on the live tenant, that's the thing to confirm; if Coupa does NOT
recompute, revisit whether to also blank/derive the amounts.

## Verify locally (no Node needed)

See [test/README.md](test/README.md). Short version:

```sh
JSC=/System/Library/Frameworks/JavaScriptCore.framework/Versions/A/Helpers/jsc

python3 test/build_roundtrip_test.py
"$JSC" test/roundtrip_test.generated.js
# expect: ==== ALL ASSERTIONS PASSED ==== (40 assertions)

python3 test/build_fx_test.py           # FX: TARGETS, fmtMoney, fallback chain
"$JSC" test/fx_test.generated.js
# expect: ==== ALL ASSERTIONS PASSED ==== (53 assertions)
```

Parse check: `jsc -e 'new Function(readFile("coupa-receipt-filter.user.js")); print("PARSE_OK")'`.

## Ship / publish (both steps — repo commit is NOT enough)

```sh
# 1. bump @version + SCRIPT_VERSION in the script, commit, push
git push origin main

# 2. update the gist (this is what pushes to installed userscripts)
GIST=a0de5a5bf56d33abef414b5781bdd984
python3 -c "import json; c=open('coupa-receipt-filter.user.js').read(); \
open('/tmp/gist.json','w').write(json.dumps({'files':{'coupa-receipt-filter.user.js':{'content':c}}}))"
gh api -X PATCH /gists/$GIST --input /tmp/gist.json \
  --jq '{updated_at, version:(.history[0].version)}'
```

Then regenerate the Word copy if docs changed:
`python3 ../AUDIT-2026-06/md2docx.py PROJECT-DOCUMENTATION.md PROJECT-DOCUMENTATION.docx "Coupa Receipt Filter — Project Documentation"`.

## Open items / ideas

- **FX offline fallback covers only the 12 display currencies.** When the live API
  is down AND no persisted pull exists AND a line is in another currency, the xlsx
  problems export can't compute USD-eq and silently skips that line's `>$25`
  receipt check (still strictly better than pre-0.10.0). Follow-up: have
  `lineProblems` emit an explicit `fx_unknown` flag rather than passing silently.
- Confirm Coupa's server-side recompute behavior on currency change (above).
- No `README.md` in the repo yet (only PROJECT-DOCUMENTATION.md). Global rule
  wants one — cheap follow-up.
- Currency dropdown only offers currencies already used on the user's lines. If a
  needed currency never appears, there's no id to set it. A future option: parse a
  currency catalog from the expense-edit HTML if Coupa embeds one.
- The example-attendee "create from header" path defaults `type_id` to 6
  (`NEW_ATTENDEE_TYPE_ID`, "BDP Employee (manual entry)"); revisit if a different
  attendee type is ever needed.
