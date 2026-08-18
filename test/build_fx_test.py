#!/usr/bin/env python3
# Extract the REAL FX functions/constants from the userscript and assemble a jsc
# program that exercises: the expanded TARGETS list, currency-aware formatting,
# convertTargets math, and the live->last-live->baked-snapshot fallback chain.
import re, sys, pathlib

HERE = pathlib.Path(__file__).resolve().parent
SRC = (HERE.parent / "coupa-receipt-filter.user.js").read_text()

def slice_fn(name):
    m = re.search(r"\n  (?:async )?function " + re.escape(name) + r"\(", SRC)
    if not m:
        sys.exit("FN NOT FOUND: " + name)
    start = m.start() + 1
    idx = SRC.index("\n  }\n", m.end())
    end = idx + len("\n  }")
    return SRC[start:end]

def slice_between(start_pat, end_pat):
    m = re.search(start_pat, SRC)
    if not m:
        sys.exit("PAT NOT FOUND: " + start_pat)
    start = m.start()
    idx = SRC.index(end_pat, m.end())
    end = idx + len(end_pat)
    return SRC[start:end].strip()

fns = ["fmtMoney", "convertTargets", "saveLastRates", "loadLastRates",
       "getRates", "fetchFxToUSD"]
extracted = "\n\n".join(slice_fn(n) for n in fns)

consts = "\n".join(slice_between(p, e) for (p, e) in [
    (r"\n  const FX_BASE_URL = ", ";"),
    (r"\n  const TARGETS = \[", "];"),
    (r"\n  const FX_TTL_MS = ", ";"),
    (r"\n  const FX_FALLBACK_TTL_MS = ", ";"),
    (r"\n  const ZERO_DP_CCY = new Set\(", ");"),
    (r"\n  const FX_FALLBACK_ASOF = ", ";"),
    (r"\n  const FX_FALLBACK = Object.freeze\(\{", "});"),
    (r"\n  const FX_STORE_KEY = ", ";"),
])

harness = r"""
'use strict';
// ---- constants + functions sliced verbatim from the userscript ----
%(CONSTS)s

// ---- controllable environment stubs jsc lacks ----
let fxCache = null;             // module-level cache getRates() closes over
let fetchFx;                    // stubbed per-test (real one uses GM_xmlhttpRequest)
let __store = {};
globalThis.localStorage = {
  getItem(k){ return Object.prototype.hasOwnProperty.call(__store,k) ? __store[k] : null; },
  setItem(k,v){ __store[k] = String(v); },
  removeItem(k){ delete __store[k]; },
};

%(EXTRACTED)s

// ================= TEST DRIVER =================
let fail = 0;
const assert = (cond, msg) => { if(!cond){ print('ASSERT FAIL: '+msg); fail++; } else { print('ok: '+msg); } };
const approx = (a, b, tol) => isFinite(a) && isFinite(b) && Math.abs(a-b) <= (tol==null?1e-6:tol);

(async () => {
  // ---- TARGETS list ----
  const want = ['SGD','HKD','TWD','KRW','MYR','IDR','VND'];
  want.forEach(c => assert(TARGETS.includes(c), 'TARGETS includes '+c));
  ['USD','EUR','COP','TRY','PLN'].forEach(c => assert(TARGETS.includes(c), 'TARGETS keeps existing '+c));
  assert(new Set(TARGETS).size === TARGETS.length, 'TARGETS has no duplicates');

  // ---- fallback table parity: every displayed currency has a snapshot rate ----
  TARGETS.forEach(c => assert(typeof FX_FALLBACK[c] === 'number' && FX_FALLBACK[c] > 0,
    'FX_FALLBACK covers '+c));
  assert(FX_FALLBACK.USD === 1, 'FX_FALLBACK USD base is 1');

  // ---- convertTargets math ----
  const R = { rates: FX_FALLBACK };
  const from100usd = convertTargets(100, 'USD', R);
  assert(approx(from100usd.USD, 100), 'convert 100 USD -> USD 100');
  assert(approx(from100usd.SGD, 100*FX_FALLBACK.SGD, 1e-4), 'convert 100 USD -> SGD via snapshot');
  assert(approx(from100usd.VND, 100*FX_FALLBACK.VND, 1e-2), 'convert 100 USD -> VND via snapshot');
  const fromVnd = convertTargets(26137.286999, 'VND', R); // exactly 1 USD worth
  assert(approx(fromVnd.USD, 1, 1e-3), 'convert 26,137 VND -> ~1 USD (round-trip)');
  assert(approx(fromVnd.HKD, FX_FALLBACK.HKD, 1e-3), 'convert VND->HKD through USD');
  const bogus = convertTargets(100, 'ZZZ', R);
  assert(!isFinite(bogus.SGD), 'unknown source currency -> NaN targets (shows as dash)');

  // ---- fmtMoney: zero-dp + grouping for large denominations, 2dp otherwise ----
  assert(fmtMoney('VND', 2613728.6999) === '2,613,729', 'VND formats grouped, no decimals');
  assert(fmtMoney('IDR', 17830.686) === '17,831', 'IDR grouped, no decimals');
  assert(fmtMoney('KRW', 1414.42) === '1,414', 'KRW grouped, no decimals');
  assert(fmtMoney('COP', 3137.9) === '3,138', 'COP grouped, no decimals');
  assert(fmtMoney('TWD', 3183.95) === '3,183.95', 'TWD grouped WITH 2 decimals');
  assert(fmtMoney('HKD', 78.45083) === '78.45', 'HKD 2 decimals');
  assert(fmtMoney('EUR', 0.863128) === '0.86', 'EUR 2 decimals');
  assert(fmtMoney('SGD', NaN) === '&mdash;', 'non-finite -> em dash');

  // ---- getRates: live success persists + returns live, no fallback flag ----
  fxCache = null; __store = {};
  fetchFx = async () => ({ rates: { USD:1, HKD:7.10, VND:25000 }, result:'success' });
  let live = await getRates();
  assert(live.__fallback === undefined, 'getRates(success): no fallback flag');
  assert(live.rates.HKD === 7.10, 'getRates(success): returns the live rate');
  assert(JSON.parse(__store[FX_STORE_KEY]).rates.HKD === 7.10, 'getRates(success): persisted last-live rates');

  // ---- getRates: live down, recent last-live present -> uses last-live ----
  fxCache = null; // store still holds the 7.10 pull from above
  fetchFx = async () => { throw new Error('network down'); };
  let deg = await getRates();
  assert(deg.__fallback === 'last-live', 'getRates(down, cached): flagged last-live');
  assert(deg.rates.HKD === 7.10, 'getRates(down, cached): serves the last live pull');

  // ---- getRates: live down, no cache -> baked snapshot ----
  fxCache = null; __store = {};
  let snap = await getRates();
  assert(snap.__fallback === FX_FALLBACK_ASOF, 'getRates(down, empty): flagged with snapshot date');
  assert(snap.rates.HKD === FX_FALLBACK.HKD, 'getRates(down, empty): serves baked snapshot');
  assert(snap.rates.TWD === FX_FALLBACK.TWD && snap.rates.MYR === FX_FALLBACK.MYR,
    'baked snapshot includes the new Asia currencies');

  // ---- getRates caches the fallback briefly so refreshes don't hammer a dead endpoint ----
  fxCache = null; __store = {};
  let downCalls = 0;
  fetchFx = async () => { downCalls++; throw new Error('down'); };
  await getRates();            // caches the snapshot fallback
  await getRates();            // should be served from the fallback cache
  assert(downCalls === 1, 'getRates caches the fallback (no repeat fetch within TTL)');
  fxCache = null;             // reset for later tests

  // ---- loadLastRates staleness: a >30-day-old pull is ignored ----
  __store = {}; saveLastRates({ USD:1, HKD:9.99 });
  const rec = JSON.parse(__store[FX_STORE_KEY]);
  rec.ts = Date.now() - 40*24*3600*1000; // backdate 40 days
  __store[FX_STORE_KEY] = JSON.stringify(rec);
  assert(loadLastRates() === null, 'loadLastRates ignores a >30-day-old cache');

  // ---- fetchFxToUSD: success persists + returns rates map ----
  __store = {};
  globalThis.fetch = async () => ({ json: async () => ({ rates: { USD:1, HKD:7.2 } }) });
  let usd = await fetchFxToUSD();
  assert(usd.HKD === 7.2, 'fetchFxToUSD(success): returns rates map');
  assert(JSON.parse(__store[FX_STORE_KEY]).rates.HKD === 7.2, 'fetchFxToUSD(success): persisted');

  // ---- fetchFxToUSD: failure -> snapshot (keeps the >$25 receipt check alive) ----
  __store = {};
  globalThis.fetch = async () => { throw new Error('down'); };
  let usdFb = await fetchFxToUSD();
  assert(usdFb.HKD === FX_FALLBACK.HKD && usdFb.VND === FX_FALLBACK.VND,
    'fetchFxToUSD(down, empty): falls back to snapshot');

  print(fail ? ('\n==== FAILURES: '+fail+' ====') : '\n==== ALL ASSERTIONS PASSED ====');
  if (fail) throw new Error(fail + ' assertion(s) failed');
})();
"""

out = harness % {"CONSTS": consts, "EXTRACTED": extracted}
pathlib.Path(HERE / "fx_test.generated.js").write_text(out)
print("wrote fx_test.generated.js (%d chars)" % len(out))
