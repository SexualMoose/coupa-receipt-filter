# Coupa Receipt Filter

A single-file Tampermonkey/Greasemonkey userscript that augments the Coupa expense UI with bulk-edit tooling: receipt-dialog filtering by ±% across currencies, bulk account assignment, receipt-to-line matching, and an xlsx round-trip for editing categories / descriptions / attendees on draft expense reports.

## Overview / Purpose

This userscript runs on any `*.coupahost.com` page and injects a persistent top-right control panel plus a filter inside Coupa's "Attach a receipt" dialog. It is a personal productivity tool to speed up cleaning and submitting expense reports in a corporate Coupa tenant (BDP / Philadelphia Finance Systems). It performs the same authenticated API calls the Coupa web app makes (using the page's CSRF token and the user's existing session cookies), so it acts entirely as the logged-in user with no separate credentials.

Core capabilities:
- **Receipt-dialog filter** — in the "Attach a receipt" modal, converts the expense-line total into USD/EUR/COP/SGD/TRY using live FX rates and hides wallet receipts outside a ±X% tolerance window.
- **Apply Account to All** — PATCHes every draft expense line whose account does not already match the configured account.
- **Account selector** — type-ahead search against Coupa's own account autocomplete; selection stored in `localStorage`.
- **Match Receipts** — pairs wallet receipts to draft lines using a tiered scoring heuristic (exact amount, ±1% same currency + merchant-token overlap, cross-currency ±12% USD-eq + token overlap) and POSTs `merge_receipt_to_expense_line`.
- **Download Non-Compliant / Export All** — generates an `.xlsx` (via ExcelJS) of draft lines with conditional formatting highlighting problems (missing receipt > $25 USD-eq, missing account/category, gift-meal value-per-attendee > $25).
- **Upload & Apply** — ingests the edited xlsx and PATCHes lines (category, description, attendees), creating new attendees via `/expense_attendees/` when needed. Skips no-op rows and confirms the first change before bulk-applying.

## Status

**Working / actively maintained.** Evidence: clean git tree, `main` branch fully pushed (0 ahead / 0 behind origin), 22 commits with a steady semver progression from v0.3.1 (initial commit, Apr 2026) to v0.8.6 (latest, ~2 weeks before audit). Commit messages are descriptive and feature-scoped. The script is 1,858 lines and includes an in-app help modal that matches the implemented behavior. No obvious dead/half-finished features.

## Technical Requirements

- A userscript manager: **Tampermonkey** (recommended) or Greasemonkey, in a desktop browser.
- An authenticated **Coupa** session on a `*.coupahost.com` tenant.
- Network access to two external services at runtime:
  - `https://open.er-api.com/v6/latest/USD` — free FX rates (declared via `@connect open.er-api.com`; no API key required).
  - `https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js` — ExcelJS, fetched on demand for xlsx generation/parsing.
- No build tooling, no package manager, no compiler. The deliverable IS the `.user.js` file.

## Dependencies

- **ExcelJS 4.4.0** (MIT license) — loaded at runtime from jsDelivr (version-pinned in `EXCELJS_URL`, line 63), not vendored. Loaded via `fetch` + `new Function(...)` eval, with a blob-`<script>` fallback if CSP blocks the eval path.
- **open.er-api.com** ("Exchange Rate API" / open access tier) — runtime HTTP JSON, no SDK bundled.
- Userscript-manager API: `GM_xmlhttpRequest` (granted) used for the cross-origin FX call.
- No npm/`package.json`, no lockfile, no other third-party code.

## Setup Instructions

1. Install Tampermonkey in your browser.
2. Open the raw script and let Tampermonkey prompt to install it, e.g. from this repo:
   `coupa-receipt-filter.user.js`
   or from the published gist referenced in the header:
   `https://gist.githubusercontent.com/SexualMoose/a0de5a5bf56d33abef414b5781bdd984/raw/coupa-receipt-filter.user.js`
3. Confirm the install in Tampermonkey. The `@match https://*.coupahost.com/*` and `@connect open.er-api.com` grants will be requested.
4. (Tenant-specific) Edit the `DEFAULT_ACCOUNT` object near the top of the script to match your own Coupa account. Capture the values by clicking "Choose" on a representative line and watching `POST /accounts/select_dynamic_account` in DevTools → Network. Update `account_id`, `account_type_id`, `display_name`, and `code`.
5. Navigate to your Coupa tenant; the panel appears top-right on expense pages.

## Build & Run

There is no build step. "Running" = having the userscript installed and active while browsing Coupa. Updates: the header declares `@updateURL`/`@downloadURL` pointing at the GitHub gist, so Tampermonkey can auto-update; the in-panel "update" link cache-busts that URL.

## Usage

On a **draft** expense report page:
- Use the top-right panel buttons: Apply Account to All, Match Receipts, Download Non-Compliant, Export All, Upload & Apply, and the account search box (with ↺ reset).
- For the receipt filter: open Coupa's "Attach a receipt" dialog, set the ± tolerance %, and the script hides receipts outside the converted-amount window.
- For the xlsx round-trip: Download Non-Compliant → edit green-header columns (`new_category`, `new_description`, attendee `x` marks) in Excel → Upload & Apply. The hidden `_categories` helper sheet drives the dropdown and name→id resolution; do not delete it.
- Click the "?" / help link in the panel footer for the built-in help modal.

## Architecture

Single IIFE in one file, organized into sections:
- **Config** (lines ~17–120): `DEFAULT_ACCOUNT`, `ALL_CATEGORIES` catalog, attendee type ids, palettes, localStorage keys.
- **ExcelJS loader** (~122–166): lazy `fetch` + `new Function` eval with blob-script fallback.
- **Helpers** (~176–230): amount parsing, FX fetch/cache (`getRates`, 1h TTL), `esc`/`escapeHtml` HTML-escapers.
- **Receipt-dialog filter** (~250–360): reads the expense-line total, converts across `TARGETS`, hides out-of-window wallet lines.
- **Account apply** (~410–540): discovers the draft report, iterates lines, PATCHes accounts via Coupa endpoints using the page CSRF token + `credentials: 'include'`.
- **xlsx export** (~550–960): builds the workbook with conditional formatting and helper sheets.
- **xlsx upload/apply** (~900–1230): parses the workbook, diffs against live line state, PATCHes lines and POSTs new attendees, with a progress overlay and `beforeunload` guard.
- **Receipt matching** (~1240–1430): tokenization + tiered scoring, posts `merge_receipt_to_expense_line`.
- **Account search UI** (~1440–1603): debounced calls to Coupa autocomplete/search endpoints, escaped result rendering.
- **Help modal + panel** (~1605–1858): UI construction.

Data flow: page DOM + Coupa JSON endpoints → in-memory line models → (xlsx out) → user edits → (xlsx in) → PATCH/POST back to Coupa. FX data flows from open.er-api.com into the converter. All Coupa calls are same-origin authenticated requests carrying `X-CSRF-Token` read from the page's `<meta name="csrf-token">`.

## Integrations & Interconnects

- **Coupa** (`*.coupahost.com`) — the host application; the script calls its internal endpoints: `/expenses`, `/expense_reports/:id/edit`, `/expenses/expense_lines/:id`, `/accounts/select_dynamic_account`, `/accounts/autocomplete|search|lookup`, `/expense_attendees/`, `/expenses/wallet/merge_receipt_to_expense_line`.
- **open.er-api.com** — free FX-rate JSON API (USD base).
- **jsDelivr CDN** — serves ExcelJS 4.4.0 at runtime.
- **GitHub Gist** (`gist.githubusercontent.com/SexualMoose/...`) — auto-update source declared in the header.
- No sibling repos, no backend of its own, no hardware.

## Configuration & Secrets

- **No secrets or credentials are stored in the repo.** The script relies entirely on the user's existing Coupa browser session (cookies) and the page CSRF token; it never holds a password, API key, or OAuth secret.
- Tenant-specific config is hardcoded near the top: `DEFAULT_ACCOUNT` (account_id 6222, type 4, a Philadelphia/Finance Systems display name, and account code `US010-26001-999-NONE-70919900`) and an `ALL_CATEGORIES` expense-category catalog. These are internal business identifiers, not credentials — see Security notes.
- User runtime state (selected account) is stored in browser `localStorage` under `__rf_active_account_v1`.

## Testing

No automated tests, no test framework, no CI. Verification is manual against a live Coupa tenant. Given it is a single userscript with no build, this is reasonable, though it means changes are validated only by hand.

## Known Issues / TODO

- No `LICENSE` file (repository is "all rights reserved" by default).
- No `.gitignore` (acceptable here — the repo holds a single source file, no build artifacts).
- No README in the repo; the only documentation is the in-script help modal (and now this file).
- Tenant-coupled: the hardcoded account, account code, and category catalog must be edited for any other Coupa tenant; behavior assumes specific Coupa endpoint shapes that vary across tenants (the code already guards some of this, e.g. account search "may not be available").
- ExcelJS is loaded from a CDN at runtime; offline use or a CDN outage breaks the xlsx features.

## Third-party & Licensing notes

- **No LICENSE file present.** As a public, owner-authored repo this defaults to all-rights-reserved.
- **ExcelJS 4.4.0** — MIT licensed; used via CDN at runtime, not copied into the repo, so no source redistribution obligation. Attribution recommended if redistributed.
- **open.er-api.com** — free/open access tier; review their terms if usage scales.
- No vendored or copied third-party source code; no foreign copyright headers; no GPL/AGPL/CC notices. Not a fork (initial commit is a self-authored userscript).
- **Trademark / brand sensitivity:** the project name, `@match`, and functionality reference **Coupa** (a third-party SaaS product/trademark). It is an unofficial, unaffiliated personal tool that drives Coupa's private/internal endpoints. The internal company name **"BDP"** and a Philadelphia office / GL account code also appear. For a clean-room or publicly shareable version, the owner should: (a) clarify "unofficial, not affiliated with Coupa," (b) remove the employer-specific account code, display name, and "BDP" references, and (c) be mindful that automating another vendor's internal endpoints may violate that vendor's terms of use.

## Security notes

- **No credential leakage.** A full `git grep` for keys/tokens/passwords/private-keys across the tree found only legitimate reads of the page's `X-CSRF-Token` meta tag — there are no committed secrets, `.env`, keystores, or service-account files. Git history is short (22 commits, single file) and shows no secret-like filenames ever added.
- **Internal business-data disclosure (low):** a corporate Coupa account id (6222), GL account code `US010-26001-999-NONE-70919900`, a department/office display name, and the "BDP" company name are committed to a public GitHub repo. Not credentials, but internal info the owner may not want public; remediation is to parameterize/remove them.
- **Runtime remote code execution by design (low/medium):** ExcelJS is fetched from jsDelivr and executed via `new Function(code)` (and a blob-`<script>` fallback). The URL is version-pinned (4.4.0) over HTTPS, which mitigates tampering, but a jsDelivr/CDN compromise would run arbitrary code in the Coupa origin with the user's session. Optionally use Subresource-Integrity-style hashing or bundle ExcelJS.
- **XSS hygiene is good:** all page-derived and remote-derived strings rendered via `innerHTML` (merchant/date/total, account search results) are passed through `esc()`/`escapeHtml()`. No unsanitized interpolation of untrusted data was found.
- **Privileged actions:** the script performs authenticated PATCH/POST mutations (account changes, attendee creation, receipt merges) as the logged-in user. It mitigates accidental bulk damage with first-change confirmations, no-op skipping, and a `beforeunload` warning, but operators should review the affected report before bulk-applying.
- Overall: no critical/high findings. Maximum severity **low**.
