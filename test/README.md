# Tests

No Node in the target environment, so the userscript's export/import logic is
verified by **executing the real functions** (sliced out of
`../coupa-receipt-filter.user.js`) against a small in-memory ExcelJS mock, run
under macOS's built-in JavaScriptCore CLI (`jsc`).

## Run

```sh
# from the repo root
python3 test/build_roundtrip_test.py          # slices real fns -> test/roundtrip_test.generated.js
/System/Library/Frameworks/JavaScriptCore.framework/Versions/A/Helpers/jsc \
  test/roundtrip_test.generated.js
```

Expected tail: `==== ALL ASSERTIONS PASSED ====` (40 assertions).

## What it covers

`build_roundtrip_test.py` extracts the actual `_buildWorkbookInner`,
`parseUploadedWorkbook`, `buildLineUpdatePatchBody`, and their helpers, then a
driver runs an **export → simulated user edits → import → PATCH-body** round trip:

- **Layout** — column order after the v0.9.0 change (`new_currency` at col G,
  downstream shifts, trailing example-attendee column) and the hidden
  `_currencies` code→id map.
- **Currency round trip** — a chosen code (`eur`) resolves back to its id.
- **Example column** — unmodified/ stray-value → ignored; renamed → creates a
  new attendee; directory columns still typo-guarded.
- **PATCH-body consistency** — single-currency lines keep the reimbursement side
  consistent (RUN5–7); the `isCrossCurrency` guard classification (RUN8).

## Parse-only sanity check

```sh
/System/Library/Frameworks/JavaScriptCore.framework/Versions/A/Helpers/jsc \
  -e 'new Function(readFile("coupa-receipt-filter.user.js")); print("PARSE_OK")'
```

`new Function(src)` parses without executing, so the browser globals the script
needs at runtime don't matter — a syntax error throws, anything else prints OK.
