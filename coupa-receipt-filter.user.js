// ==UserScript==
// @name         Coupa Receipt Filter (Attach Receipt dialog, ±% across currencies)
// @namespace    local.tylerkeller
// @version      0.8.4
// @description  Filter the Coupa "Attach a receipt" dialog by ±X%, plus a top-right panel with Apply-Account-to-All, Download-Problems (xlsx with red/yellow row highlights AND conditional formatting on invalid entries), and Upload-and-Apply (description + attendee bulk edit with first-line confirmation + progress bar).
// @match        https://*.coupahost.com/*
// @run-at       document-idle
// @grant        GM_xmlhttpRequest
// @connect      open.er-api.com
// @updateURL    https://gist.githubusercontent.com/SexualMoose/a0de5a5bf56d33abef414b5781bdd984/raw/coupa-receipt-filter.user.js
// @downloadURL  https://gist.githubusercontent.com/SexualMoose/a0de5a5bf56d33abef414b5781bdd984/raw/coupa-receipt-filter.user.js
// ==/UserScript==

(function () {
  'use strict';

  // ---------- config ----------
  // Default account to bulk-apply via the toolbar button.
  // Capture these values from your tenant by clicking "Choose" on a representative
  // line and watching POST /accounts/select_dynamic_account in DevTools Network.
  const DEFAULT_ACCOUNT = Object.freeze({
    account_id: 6222,        // returned id from /accounts/select_dynamic_account
    account_type_id: 4,       // US1
    display_name: 'PHILADELPHIA-Finance Systems & Projects-NONE-Miscellaneous expenses',
    code: 'US010-26001-999-NONE-70919900',
  });
  const ACTIVE_ACCOUNT_LSKEY = '__rf_active_account_v1';
  function getActiveAccount() {
    try {
      const raw = localStorage.getItem(ACTIVE_ACCOUNT_LSKEY);
      if (raw) {
        const obj = JSON.parse(raw);
        if (obj && obj.account_id) return obj;
      }
    } catch {}
    return DEFAULT_ACCOUNT;
  }
  function setActiveAccount(obj) {
    localStorage.setItem(ACTIVE_ACCOUNT_LSKEY, JSON.stringify(obj));
  }
  function resetActiveAccount() {
    localStorage.removeItem(ACTIVE_ACCOUNT_LSKEY);
  }

  const SCRIPT_VERSION = '0.8.4';
  // Palette used to randomize the help-modal accent color each open
  const HELP_PALETTE = [
    { fg: '#1976D2', name: 'blue' },
    { fg: '#388E3C', name: 'green' },
    { fg: '#7B1FA2', name: 'purple' },
    { fg: '#D32F2F', name: 'red' },
    { fg: '#F57C00', name: 'orange' },
    { fg: '#00796B', name: 'teal' },
    { fg: '#C2185B', name: 'pink' },
    { fg: '#303F9F', name: 'indigo' },
    { fg: '#5D4037', name: 'brown' },
    { fg: '#455A64', name: 'blue-grey' },
  ];
  function pickHelpAccent() {
    return HELP_PALETTE[Math.floor(Math.random() * HELP_PALETTE.length)].fg;
  }
  const SCRIPT_UPDATE_URL = 'https://gist.githubusercontent.com/SexualMoose/a0de5a5bf56d33abef414b5781bdd984/raw/coupa-receipt-filter.user.js';
  const EXCELJS_URL = 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js';
  const FX_BASE_URL = 'https://open.er-api.com/v6/latest/USD';
  const PROBLEM_USD_THRESHOLD = 25;
  const GIFT_MEAL_CATEGORY_ID = 85; // "Entertainment (Gift): To Internal Employee - Meal"
  const NEW_ATTENDEE_TYPE_ID = 6;   // "BDP Employee (manual entry)"

  let excelJsPromise = null;
  function getExcelJS() {
    return (typeof window !== 'undefined' && window.ExcelJS)
        || (typeof self !== 'undefined' && self.ExcelJS)
        || (typeof globalThis !== 'undefined' && globalThis.ExcelJS)
        || (typeof unsafeWindow !== 'undefined' && unsafeWindow.ExcelJS);
  }
  function loadExcelJS() {
    const existing = getExcelJS();
    if (existing) return Promise.resolve(existing);
    if (!excelJsPromise) {
      excelJsPromise = (async () => {
        const r = await fetch(EXCELJS_URL);
        const code = await r.text();
        // Wrap so the UMD wrapper finds a usable global, and capture the result.
        try {
          const fn = new Function(code + '\n; return (typeof ExcelJS !== "undefined") ? ExcelJS : (typeof self !== "undefined" && self.ExcelJS);');
          const result = fn();
          if (result) {
            try { window.ExcelJS = result; } catch {}
            try { self.ExcelJS = result; } catch {}
          }
        } catch (e) {
          // Fall through to alternative load
        }
        let resolved = getExcelJS();
        if (resolved) return resolved;
        // Fallback: blob-script tag (works when CSP allows blob:)
        const blob = new Blob([code], { type: 'application/javascript' });
        const url = URL.createObjectURL(blob);
        await new Promise((resolve, reject) => {
          const s = document.createElement('script');
          s.src = url;
          s.onload = resolve;
          s.onerror = () => reject(new Error('ExcelJS blob-script load failed (CSP?)'));
          document.head.appendChild(s);
        });
        URL.revokeObjectURL(url);
        resolved = getExcelJS();
        if (!resolved) throw new Error('ExcelJS loaded but global not set; check Tampermonkey sandbox.');
        return resolved;
      })();
    }
    return excelJsPromise;
  }

  const TARGETS = ['USD', 'EUR', 'COP', 'SGD', 'TRY'];
  const DEFAULT_TOL_PCT = 5;
  const FX_URL = 'https://open.er-api.com/v6/latest/USD';
  const FX_TTL_MS = 60 * 60 * 1000;

  // Dialog text that identifies the modal we filter
  const DIALOG_TITLE_RE = /Attach a receipt/i;

  // ---------- helpers ----------
  const parseAmount = (s) => {
    if (s == null) return NaN;
    let t = String(s).trim().replace(/[^0-9.,\-]/g, '');
    if (t.includes(',') && t.includes('.')) t = t.replace(/,/g, '');
    else if (t.includes(',')) {
      const lc = t.lastIndexOf(',');
      t = (t.length - lc - 1 === 2) ? t.replace(/,/g, '.') : t.replace(/,/g, '');
    }
    return parseFloat(t);
  };

  const fetchFx = () => new Promise((resolve, reject) => {
    if (typeof GM_xmlhttpRequest === 'function') {
      GM_xmlhttpRequest({
        method: 'GET',
        url: FX_URL,
        onload: r => { try { resolve(JSON.parse(r.responseText)); } catch (e) { reject(e); } },
        onerror: reject,
      });
    } else {
      fetch(FX_URL).then(r => r.json()).then(resolve).catch(reject);
    }
  });

  let fxCache = null;
  async function getRates() {
    if (fxCache && Date.now() - fxCache.ts < FX_TTL_MS) return fxCache.data;
    const data = await fetchFx();
    fxCache = { ts: Date.now(), data };
    return data;
  }

  function readTotal() {
    const amt = document.querySelector('input[aria-label="Total Amount"]');
    const cur = document.querySelector('input[aria-label="Total Currency"]');
    if (!amt) return null;
    const value = parseAmount(amt.value);
    const currency = (cur?.value || 'USD').trim().toUpperCase();
    if (!isFinite(value)) return null;
    return { amount: value, currency };
  }

  function readExpenseLineMeta() {
    const merchant = document.querySelector('input[name="merchant"]')?.value?.trim() || '';
    const expenseDate = document.querySelector('input[name="local_expense_date"]')?.value?.trim() || '';
    return { merchant, expenseDate };
  }

  const esc = (s) => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

  function convertTargets(amount, currency, rates) {
    const amtUSD = currency === 'USD' ? amount : amount / (rates.rates[currency] || NaN);
    const out = {};
    TARGETS.forEach(c => out[c] = c === 'USD' ? amtUSD : amtUSD * (rates.rates[c] || NaN));
    return out;
  }

  function collectDialogItems(dialog) {
    return Array.from(dialog.querySelectorAll('li.unattachedReceiptLine')).map(li => ({
      li,
      amount: parseAmount(li.querySelector('.s-receiptAmount')?.textContent),
      currency: (li.querySelector('.currency_code, .s-currencyCode')?.textContent || '').trim().toUpperCase(),
    }));
  }

  // ---------- per-dialog state & UI ----------
  function attachFilterToDialog(dialog) {
    if (dialog.__rfAttached) return;
    dialog.__rfAttached = true;

    // Build bar
    const bar = document.createElement('div');
    bar.className = '__rf_bar';
    bar.style.cssText = [
      'padding:8px',
      'margin:8px',
      'background:#fffbe6',
      'border:1px solid #f0c36d',
      'border-radius:4px',
      'font:12px sans-serif',
    ].join(';');
    bar.innerHTML = `
      <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">
        <strong>Receipt filter</strong>
        <label>&plusmn;<input class="__rf_tol" type="number" value="${DEFAULT_TOL_PCT}" min="0" step="0.1" style="width:48px;">%</label>
        <button class="__rf_apply" type="button">Filter now</button>
        <label><input class="__rf_auto" type="checkbox" checked> auto</label>
        <button class="__rf_clear" type="button">Show all</button>
        <button class="__rf_fx" type="button" title="Refetch FX">&#8635; FX</button>
      </div>
      <div class="__rf_meta" style="margin-top:6px;color:#222;"></div>
      <div class="__rf_targets" style="margin-top:4px;color:#444;"></div>
      <div class="__rf_status" style="margin-top:4px;color:#555;"></div>
    `;

    // Insert bar at top of dialog
    const firstChild = dialog.firstElementChild;
    dialog.insertBefore(bar, firstChild);

    const $ = (sel) => bar.querySelector(sel);

    // Shrink the scrollable body wrapper by the bar's vertical cost so the
    // footer buttons (Attach/Cancel) don't get pushed off the viewport.
    function resizeBodyWrapper() {
      const wrap = dialog.querySelector('.reactModal__bodyWrapper');
      if (!wrap) return;
      if (!wrap.dataset.__rfOrigHeight) {
        wrap.dataset.__rfOrigHeight = getComputedStyle(wrap).height;
      }
      const orig = parseFloat(wrap.dataset.__rfOrigHeight);
      if (!isFinite(orig)) return;
      const bs = getComputedStyle(bar);
      const cost = bar.offsetHeight + parseFloat(bs.marginTop || 0) + parseFloat(bs.marginBottom || 0);
      wrap.style.height = Math.max(200, orig - cost) + 'px';
    }
    resizeBodyWrapper();
    // Re-run on resize and on filter bar content changes (which can reflow the bar height)
    const ro = new ResizeObserver(() => resizeBodyWrapper());
    ro.observe(bar);
    window.addEventListener('resize', resizeBodyWrapper);

    let applying = false;
    function applyFilter(targets, tol) {
      applying = true;
      try {
        let shown = 0, hidden = 0;
        for (const { li, amount, currency } of collectDialogItems(dialog)) {
          const tgt = targets[currency];
          const ok = isFinite(amount) && isFinite(tgt) &&
                     amount >= tgt * (1 - tol) && amount <= tgt * (1 + tol);
          const next = ok ? '' : 'none';
          if (li.style.display !== next) li.style.display = next;
          ok ? shown++ : hidden++;
        }
        return { shown, hidden };
      } finally {
        setTimeout(() => { applying = false; }, 0);
      }
    }

    function clearFilter() {
      applying = true;
      dialog.querySelectorAll('li.unattachedReceiptLine').forEach(li => {
        if (li.style.display) li.style.display = '';
      });
      setTimeout(() => { applying = false; }, 0);
    }

    async function refresh() {
      const total = readTotal();
      const meta = readExpenseLineMeta();
      $('.__rf_meta').innerHTML =
        `<b>Merchant:</b> ${esc(meta.merchant || '—')} ` +
        `&nbsp; <b>Date:</b> ${esc(meta.expenseDate || '—')} ` +
        `&nbsp; <b>Total:</b> ${total ? esc(total.amount + ' ' + total.currency) : '—'}`;
      if (!total) {
        $('.__rf_targets').textContent = 'Could not read Total Amount from the expense line.';
        $('.__rf_status').textContent = '';
        return;
      }
      let rates;
      try { rates = await getRates(); }
      catch (e) {
        $('.__rf_targets').textContent = 'FX fetch failed: ' + e.message;
        return;
      }
      const tgts = convertTargets(total.amount, total.currency, rates);
      const tol = (parseFloat($('.__rf_tol').value) || 0) / 100;
      $('.__rf_targets').innerHTML =
        `<b>Source &rarr;</b> ` +
        TARGETS.map(c => {
          const v = tgts[c];
          if (!isFinite(v)) return `${c}: &mdash;`;
          return `<span style="display:inline-block;margin-right:8px;"><b>${c}</b> ${v.toFixed(2)} <span style="color:#888;">(${(v*(1-tol)).toFixed(2)}&ndash;${(v*(1+tol)).toFixed(2)})</span></span>`;
        }).join('');
      const r = applyFilter(tgts, tol);
      $('.__rf_status').textContent = `${r.shown} shown / ${r.hidden} hidden`;
    }

    // Debounced auto-refresh
    let deb;
    const debounced = () => {
      clearTimeout(deb);
      deb = setTimeout(() => { if ($('.__rf_auto').checked) refresh(); }, 300);
    };

    $('.__rf_apply').addEventListener('click', refresh);
    $('.__rf_clear').addEventListener('click', () => {
      clearFilter();
      $('.__rf_status').textContent = 'cleared';
    });
    $('.__rf_fx').addEventListener('click', async () => {
      fxCache = null;
      await refresh();
    });

    // Re-apply when new receipts get added/removed inside the dialog (e.g., infinite scroll)
    const list = dialog.querySelector('ul') || dialog;
    const mo = new MutationObserver(() => {
      if (applying) return;
      debounced();
    });
    mo.observe(list, { childList: true });

    // Initial apply
    refresh();
  }

  // ---------- bulk account-apply ----------
  // Parse the embedded "var ExpenseReports = [...]" JSON from a Coupa HTML page.
  function parseExpenseReportsFromHtml(html) {
    const start = html.search(/var\s+ExpenseReports\s*=\s*\[/);
    if (start < 0) return null;
    const arrStart = html.indexOf('[', start);
    let depth = 0, i = arrStart, inStr = false, strCh = '', esc = false;
    for (; i < html.length; i++) {
      const c = html[i];
      if (inStr) { if (esc) esc = false; else if (c === '\\') esc = true; else if (c === strCh) inStr = false; }
      else { if (c === '"' || c === '\'') { inStr = true; strCh = c; } else if (c === '[') depth++; else if (c === ']') { depth--; if (depth === 0) { i++; break; } } }
    }
    try { return JSON.parse(html.slice(arrStart, i)); } catch { return null; }
  }

  async function findDraftReportIds() {
    // Try the in-page global first (works without Tampermonkey isolation; only the
    // bundle-context global is reliably accessible via unsafeWindow when granted).
    try {
      // eslint-disable-next-line no-undef
      const w = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;
      if (Array.isArray(w.ExpenseReports) && w.ExpenseReports.length) {
        return w.ExpenseReports.filter(r => r.status === 'draft').map(r => r.id);
      }
    } catch {}
    // Fallback: fetch the listing page and parse the embedded JSON.
    try {
      const r = await fetch('/expenses', { credentials: 'include' });
      const html = await r.text();
      const arr = parseExpenseReportsFromHtml(html);
      if (Array.isArray(arr)) return arr.filter(r => r.status === 'draft').map(r => r.id);
    } catch {}
    return [];
  }

  function fetchReportLines(reportId) {
    return fetch(`/expense_reports/${reportId}/edit`, { credentials: 'include' })
      .then(r => r.text())
      .then(html => {
        const arr = parseExpenseReportsFromHtml(html);
        return arr ? arr.find(p => p.id === reportId) : null;
      });
  }

  function buildAccountPatchBody(line) {
    const u = new URLSearchParams();
    const set = (k, v) => u.append(k, v == null ? '' : String(v));
    set('expense_line[custom_field_3]', line.custom_field_3 ?? '');
    set('expense_line[travel_provider_type]', line.travel_provider_type ?? '');
    set('expense_line[audit_status_id]', line.audit_status_id ?? '');
    set('expense_line[reason]', line.reason ?? '');
    set('expense_line[amount_to_receive]', line.amount_to_receive ?? '');
    {
      const _aa = getActiveAccount();
      set('expense_line[account_id]', _aa.account_id);
      set('expense_line[account_type_id]', _aa.account_type_id);
    }
    set('expense_line[merchant]', line.merchant ?? '');
    set('expense_line[local_expense_date]', line.local_expense_date ?? '');
    set('expense_line[parent_expense_line_id]', line.parent_expense_line_id ?? '');
    set('expense_line[start_date]', line.start_date ?? line.local_expense_date ?? '');
    set('expense_line[end_date]', line.end_date ?? line.local_expense_date ?? '');
    set('expense_line[travel_provider_name]', line.travel_provider_name ?? '');
    set('expense_line[expense_category_id]', line.expense_category_id ?? '');
    set('expense_line[employee_reimbursable]', line.employee_reimbursable ? 'true' : 'false');
    set('expense_line[expense_category_custom_field_1]', line.expense_category_custom_field_1 ?? '');
    set('expense_line[receipt_total_currency_id]', line.receipt_currency?.id ?? '');
    set('expense_line[foreign_currency_id]', line.amount_to_receive_currency?.id ?? '');
    set('expense_line[external_src_name]', line.external_src_name ?? '');
    set('expense_line[exchange_rate]', line.exchange_rate ?? '');
    set('expense_line[description]', line.description ?? '');
    set('expense_line[employee_reimbursable_overridden]', line.employee_reimbursable_overridden ? 'true' : 'false');
    set('expense_line[divisor]', line.divisor ?? '');
    set('expense_line[receipt_total_amount]', line.receipt_amount ?? '');
    set('expense_line[foreign_currency_amount]', line.amount_to_receive ?? '');
    set('expense_line[expense_report_id]', line.expense_report_id ?? '');
    set('expense_line[custom_field_2]', line.custom_field_2 ? 'true' : 'false');
    (line.expense_attendees || []).forEach(a => u.append('expense_line[attendee_ids][]', a.id));
    return u.toString();
  }

  async function runApplyAccountToAll(reportIds, panel) {
    window.__rfAcctRunning = true;
    const statusEl = panel.querySelector('.__rf_acct_status');
    const btn = panel.querySelector('.__rf_apply_account_btn');
    const progressWrap = panel.querySelector('.__rf_progress_wrap');
    const progressBar = panel.querySelector('.__rf_progress_bar');
    const progressText = panel.querySelector('.__rf_progress_text');
    const setProgress = (done, total) => {
      const pct = total > 0 ? Math.min(100, (done / total) * 100) : 0;
      progressBar.style.width = pct.toFixed(1) + '%';
      progressText.textContent = total > 0 ? `${done} / ${total} (${pct.toFixed(0)}%)` : '';
    };
    progressWrap.style.display = 'block';
    progressText.style.display = 'block';
    btn.disabled = true;
    btn.style.opacity = '0.6';
    btn.style.cursor = 'wait';
    btn.textContent = 'Running…';
    // Warn user if they try to navigate while running
    const beforeUnload = (e) => { e.preventDefault(); e.returnValue = ''; };
    window.addEventListener('beforeunload', beforeUnload);

    const csrf = document.querySelector('meta[name="csrf-token"]')?.getAttribute('content');
    const headers = {
      'Accept': 'application/json, text/plain, */*',
      'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
      'X-CSRF-Token': csrf,
      'X-Requested-With': 'XMLHttpRequest',
    };
    let ok = 0, fail = 0, skipped = 0;
    const failures = [];
    statusEl.textContent = 'Fetching reports…';

    try {
      const allLines = [];
      for (const rid of reportIds) {
        try {
          const rpt = await fetchReportLines(rid);
          if (rpt && Array.isArray(rpt.expense_lines)) {
            for (const l of rpt.expense_lines) allLines.push(l);
          }
        } catch (e) {
          failures.push({ report_id: rid, error: 'fetch failed: ' + String(e).slice(0, 100) });
        }
      }

      const needsUpdate = allLines.filter(l => {
        const accounts = Array.isArray(l.accounts) ? l.accounts : [];
        return !accounts.some(a => Number(a.account_id) === Number(getActiveAccount().account_id));
      });
      skipped = allLines.length - needsUpdate.length;
      const total = needsUpdate.length;
      statusEl.textContent = `${skipped} already set, patching ${total}…`;
      setProgress(0, total);

      for (const line of needsUpdate) {
        try {
          const r = await fetch(`/expenses/expense_lines/${line.id}`, {
            method: 'PATCH',
            credentials: 'include',
            headers,
            body: buildAccountPatchBody(line),
          });
          if (r.ok) ok++;
          else { fail++; const txt = await r.text(); failures.push({ line_id: line.id, status: r.status, head: txt.slice(0, 200) }); }
        } catch (e) {
          fail++; failures.push({ line_id: line.id, error: String(e).slice(0, 200) });
        }
        const done = ok + fail;
        setProgress(done, total);
        statusEl.textContent = `ok=${ok} fail=${fail} (skipped ${skipped})`;
        await new Promise(r => setTimeout(r, 120));
      }

      statusEl.innerHTML = `<b>Done.</b> ok=${ok} / fail=${fail} / skipped ${skipped}` +
        (failures.length ? ` <a href="data:application/json;base64,${btoa(JSON.stringify(failures, null, 2))}" download="account-apply-failures.json" style="color:#06c;">download failures</a>` : '');
      progressBar.style.background = fail > 0 ? '#c60' : '#0a7';
    } finally {
      window.__rfAcctRunning = false;
      window.removeEventListener('beforeunload', beforeUnload);
      btn.disabled = false;
      btn.style.opacity = '';
      btn.style.cursor = 'pointer';
      btn.textContent = 'Apply Account to All';
    }
  }

  // ---------- problem detection / xlsx download / upload-and-apply ----------
  async function fetchAllDraftReports() {
    const ids = await findDraftReportIds();
    const reports = [];
    for (const id of ids) {
      try {
        const r = await fetch(`/expense_reports/${id}/edit`, { credentials: 'include' });
        const html = await r.text();
        const arr = parseExpenseReportsFromHtml(html);
        const rpt = arr ? arr.find(p => p.id === id) : null;
        if (rpt) reports.push(rpt);
      } catch {}
    }
    return reports;
  }

  async function fetchFxToUSD() {
    try {
      const r = await fetch(FX_BASE_URL);
      const j = await r.json();
      return j.rates || {};
    } catch { return {}; }
  }

  function parseAmt(s) {
    if (s == null) return NaN;
    let t = String(s).replace(/[^0-9.,\-]/g, '');
    if (t.includes(',') && t.includes('.')) t = t.replace(/,/g, '');
    else if (t.includes(',')) {
      const lc = t.lastIndexOf(',');
      t = (t.length - lc - 1 === 2) ? t.replace(/,/g, '.') : t.replace(/,/g, '');
    }
    return parseFloat(t);
  }

  function lineProblems(line, usdRates) {
    const amt = parseAmt(line.receipt_amount);
    const cur = (line.receipt_currency?.code || '').toUpperCase();
    const rate = cur === 'USD' ? 1 : usdRates[cur];
    const usdEq = isFinite(amt) && rate ? (cur === 'USD' ? amt : amt / rate) : NaN;
    const hasReceipt = (line.expense_artifacts || []).length > 0;
    const hasAccount = Array.isArray(line.accounts) && line.accounts.length > 0;
    const hasCategory = !!line.expense_category_id;
    const attendees = line.expense_attendees || [];
    const isGiftMeal = line.expense_category_id === GIFT_MEAL_CATEGORY_ID;
    const perAttendee = isGiftMeal && isFinite(amt) ? amt / Math.max(attendees.length, 1) : null;

    const problems = [];
    if (isFinite(usdEq) && usdEq > PROBLEM_USD_THRESHOLD && !hasReceipt) problems.push('missing_receipt_>$25');
    if (!hasAccount) problems.push('missing_account');
    if (!hasCategory) problems.push('missing_category');
    if (isGiftMeal && perAttendee != null && perAttendee > PROBLEM_USD_THRESHOLD) problems.push('gift_meal_per_attendee_>$25');
    return { problems, usdEq, isGiftMeal, attendees, perAttendee };
  }

  function collectAttendeeDirectory(reports) {
    const seen = new Map(); // id -> { id, type_id, first_name, last_name }
    reports.forEach(r => (r.expense_lines || []).forEach(l => {
      (l.expense_attendees || []).forEach(a => {
        if (!seen.has(a.id)) {
          seen.set(a.id, { id: a.id, type_id: a.expense_attendee_type_id, first_name: a.first_name || '', last_name: a.last_name || '' });
        }
      });
    }));
    const list = Array.from(seen.values());
    list.sort((a, b) => `${a.first_name} ${a.last_name}`.localeCompare(`${b.first_name} ${b.last_name}`));
    return list;
  }

  async function buildProblemsWorkbook(panel, opts) {
    const onlyProblems = !opts || opts.onlyProblems !== false;
    const status = panel.querySelector('.__rf_acct_status');
    let stage = 'init';
    try {
      stage = 'load ExcelJS';
      status.textContent = 'Loading ExcelJS…';
      const ExcelJS = await loadExcelJS();
      if (!ExcelJS || !ExcelJS.Workbook) throw new Error('ExcelJS not available after load');
      stage = 'fetch reports';
      status.textContent = 'Fetching reports…';
      const [reports, usdRates] = await Promise.all([fetchAllDraftReports(), fetchFxToUSD()]);
      stage = 'analyze lines';
      status.textContent = 'Analyzing lines…';
      return await _buildWorkbookInner(ExcelJS, reports, usdRates, status, panel, onlyProblems);
    } catch (e) {
      const msg = (e && e.message) ? e.message : String(e);
      console.error('[CoupaReceiptFilter] buildProblemsWorkbook failed at stage', stage, e);
      throw new Error(`download failed at "${stage}": ${msg}`);
    }
  }

  // Sweep distinct expense categories actually used on draft lines, with their IDs.
  function collectUsedCategories(reports) {
    const map = new Map();
    reports.forEach(r => (r.expense_lines || []).forEach(l => {
      if (l.expense_category_id && l.expense_category_name) {
        map.set(l.expense_category_name, l.expense_category_id);
      }
    }));
    return Array.from(map.entries())
      .map(([name, id]) => ({ name, id }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }

  async function _buildWorkbookInner(ExcelJS, reports, usdRates, status, panel, onlyProblems) {

    const attendees = collectAttendeeDirectory(reports);
    const usedCategories = collectUsedCategories(reports);
    const YELLOW = 'FFFFEB9C';    // light yellow — only used for gift-meal-per-attendee rows
    const GREEN_HDR = 'FFC8E6C9'; // light green for editable column headers

    const wb = new ExcelJS.Workbook();

    // Hidden helper sheet with valid category names (used by CF formula on Lines).
    const catSh = wb.addWorksheet('_categories', { state: 'veryHidden' });
    catSh.addRow(['name', 'id']);
    usedCategories.forEach(c => catSh.addRow([c.name, c.id]));

    const sh = wb.addWorksheet('Lines');
    // Column order:
    // A line_id, B report_title, C merchant, D date, E amount, F currency, G usd_eq,
    // H current_category, I new_category (EDITABLE+dropdown),
    // J problems, K current_description, L new_description (EDITABLE),
    // M current_attendees, N+ attendee columns (EDITABLE)
    const headers = [
      'line_id', 'report_title', 'merchant', 'date', 'amount', 'currency', 'usd_eq',
      'current_category', 'new_category',
      'problems', 'current_description', 'new_description',
      'current_attendees',
    ].concat(attendees.map(a => `${a.first_name} ${a.last_name} (${a.id})`));
    sh.addRow(headers);
    sh.getRow(1).font = { bold: true };
    sh.views = [{ state: 'frozen', ySplit: 1, xSplit: 4 }];
    [10, 32, 28, 10, 10, 7, 10, 24, 24, 28, 26, 28, 30].forEach((w, i) => {
      sh.getColumn(i + 1).width = w;
    });
    const ATTENDEE_COL_START = 14; // column N
    for (let i = ATTENDEE_COL_START; i <= headers.length; i++) sh.getColumn(i).width = 8;

    // Green headers on editable columns: I (new_category), L (new_description), N..end (attendees)
    const editableHeaderCols = [9, 12]; // I, L
    for (let c = ATTENDEE_COL_START; c <= headers.length; c++) editableHeaderCols.push(c);
    editableHeaderCols.forEach(c => {
      const cell = sh.getRow(1).getCell(c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN_HDR } };
    });

    let problemRows = 0;
    reports.forEach(rpt => {
      (rpt.expense_lines || []).forEach(l => {
        const { problems, usdEq, isGiftMeal, attendees: lineAttendees, perAttendee } = lineProblems(l, usdRates);
        if (onlyProblems && !problems.length) return;
        problemRows++;
        const attIds = new Set(lineAttendees.map(a => a.id));
        const currentAttendees = lineAttendees.map(a => `${a.first_name} ${a.last_name}`).join('; ');
        const row = [
          l.id, rpt.title, l.merchant || '', l.local_expense_date || '',
          parseFloat(l.receipt_amount) || '',
          (l.receipt_currency?.code || '').toUpperCase(),
          isFinite(usdEq) ? Number(usdEq.toFixed(2)) : '',
          l.expense_category_name || '',
          '', // new_category (editable, blank by default)
          problems.join(', '),
          l.description || '',
          '',
          currentAttendees,
        ];
        attendees.forEach(a => row.push(attIds.has(a.id) ? 'x' : ''));
        const r = sh.addRow(row);

        // Only keep YELLOW row tint for gift-meal-per-attendee>$25 rows.
        // (User asked: red ONLY on per-cell CF for new_category + attendee cols, nothing else.)
        const hasYellow = problems.includes('gift_meal_per_attendee_>$25');
        if (hasYellow) {
          r.eachCell({ includeEmpty: false }, cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: YELLOW } };
          });
        }
      });
    });

    // ----- Data validation for new_category dropdown (col I) -----
    // Excel inline dropdown limit is ~255 chars when using a literal list; for safety,
    // reference the hidden _categories sheet's range instead.
    try {
      const catRows = usedCategories.length;
      if (catRows > 0) {
        const lastDataRow = Math.max(2, problemRows + 1);
        for (let r = 2; r <= lastDataRow; r++) {
          sh.getCell(`I${r}`).dataValidation = {
            type: 'list',
            allowBlank: true,
            formulae: [`_categories!$A$2:$A$${catRows + 1}`],
            showErrorMessage: true,
            errorStyle: 'error',
            errorTitle: 'Invalid category',
            error: 'Use the dropdown to pick a valid category. Leave blank to keep current.',
          };
        }
      }
    } catch (e) { console.warn('[CoupaReceiptFilter] new_category dropdown skipped:', e); }

    // ----- Conditional formatting on Lines -----
    // Red on:
    //   - new_category (I): non-empty AND not present in _categories!A:A
    //   - any attendee column (N+): non-empty AND not "x"/"X"
    try {
      const colLetter = (n) => {
        let s = '';
        while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
        return s;
      };
      // CF: new_category invalid (referencing helper sheet)
      sh.addConditionalFormatting({
        ref: 'I2:I10000',
        rules: [{
          type: 'expression',
          priority: 1,
          formulae: ['AND(LEN(I2)>0, COUNTIF(_categories!$A:$A, I2)=0)'],
          style: {
            fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF6B6B' } },
            font: { color: { argb: 'FFFFFFFF' }, bold: true },
          },
        }],
      });
      // CF: attendee columns invalid
      if (headers.length >= ATTENDEE_COL_START) {
        const attCol1Letter = colLetter(ATTENDEE_COL_START);
        const attColLastLetter = colLetter(headers.length);
        const refRange = `${attCol1Letter}2:${attColLastLetter}10000`;
        sh.addConditionalFormatting({
          ref: refRange,
          rules: [{
            type: 'expression',
            priority: 1,
            formulae: [`AND(LEN(${attCol1Letter}2)>0, UPPER(${attCol1Letter}2)<>"X")`],
            style: {
              fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFF6B6B' } },
              font: { color: { argb: 'FFFFFFFF' }, bold: true },
            },
          }],
        });
      }
    } catch (e) { console.warn('[CoupaReceiptFilter] Lines conditional formatting skipped:', e); }

    // Attendees sheet
    const ash = wb.addWorksheet('Attendees');
    ash.addRow(['id', 'type_id', 'first_name', 'last_name']);
    ash.getRow(1).font = { bold: true };
    ash.views = [{ state: 'frozen', ySplit: 1 }];
    [10, 10, 22, 22].forEach((w, i) => { ash.getColumn(i + 1).width = w; });
    attendees.forEach(a => ash.addRow([a.id, a.type_id || NEW_ATTENDEE_TYPE_ID, a.first_name, a.last_name]));
    // Attach instructions as a header-cell comment so it doesn't pollute the data area.
    try {
      ash.getCell('A1').note = {
        texts: [{ text: 'To add a NEW attendee: append a row with id BLANK, type_id blank (defaults to 6), set first_name and last_name. Then add a column header "first last" on the Lines sheet to mark X.' }],
      };
    } catch (e) { /* fall back: no comment */ }

    // (No conditional formatting on Attendees sheet — user requested no red anywhere
    // except per-cell on Lines new_category and attendee columns.)

    status.textContent = `Building xlsx (${problemRows} problem rows)…`;
    const buf = await wb.xlsx.writeBuffer();
    return { blob: new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), problemRows };
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 1000);
  }

  // ---------- upload + apply ----------
  async function parseUploadedWorkbook(file) {
    const ExcelJS = await loadExcelJS();
    const buf = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);

    const linesSheet = wb.getWorksheet('Lines');
    const attendeesSheet = wb.getWorksheet('Attendees');
    if (!linesSheet || !attendeesSheet) throw new Error('Workbook must contain "Lines" and "Attendees" sheets.');

    // Validate + parse Attendees
    const validationErrors = [];
    const attendees = [];
    attendeesSheet.eachRow({ includeEmpty: false }, (row, idx) => {
      if (idx === 1) return;
      const id = row.getCell(1).value;
      const type_id = row.getCell(2).value;
      const first_name = row.getCell(3).value;
      const last_name = row.getCell(4).value;
      // Skip rows that have neither first_name nor last_name — these are not real
      // attendee rows even if some other column got typed into accidentally.
      if (!first_name && !last_name) return;
      // id must be numeric or empty
      if (id != null && id !== '' && isNaN(Number(id))) {
        validationErrors.push(`Attendees!A${idx}: id "${id}" is not numeric`);
      }
      // type_id must be 5 or 6 if present
      if (type_id != null && type_id !== '' && Number(type_id) !== 5 && Number(type_id) !== 6) {
        validationErrors.push(`Attendees!B${idx}: type_id "${type_id}" must be 5 or 6`);
      }
      // If creating a new attendee (no id), must have first_name and last_name
      if ((!id || id === '') && (!first_name || !last_name)) {
        validationErrors.push(`Attendees!A${idx}: new attendee row requires first_name AND last_name`);
      }
      attendees.push({
        id: id ? Number(id) : null,
        type_id: type_id ? Number(type_id) : NEW_ATTENDEE_TYPE_ID,
        first_name: first_name ? String(first_name).trim() : '',
        last_name: last_name ? String(last_name).trim() : '',
      });
    });

    // Build a label->record map for quick lookup
    const labelToAttendee = new Map();
    attendees.forEach(a => {
      const labelWithId = `${a.first_name} ${a.last_name} (${a.id})`;
      const labelNoId = `${a.first_name} ${a.last_name}`;
      if (a.id) labelToAttendee.set(labelWithId, a);
      labelToAttendee.set(labelNoId, a);
    });

    // Parse Lines: for each row, gather changes
    const headers = linesSheet.getRow(1).values.slice(1).map(v => v == null ? '' : String(v));
    const colIdx = (name) => headers.findIndex(h => String(h).trim() === name) + 1;
    const COL_LINE_ID = colIdx('line_id');
    const COL_NEW_DESC = colIdx('new_description');
    const COL_NEW_CAT = colIdx('new_category');
    if (!COL_LINE_ID || !COL_NEW_DESC) throw new Error('Required columns not found.');
    // Attendee columns are anything after current_attendees
    const attendeeColStart = colIdx('current_attendees') + 1;

    // Build category-name -> id map by reading the hidden _categories helper sheet
    // (or fall back to whatever is on the report data if the helper sheet was stripped).
    const categoryNameToId = new Map();
    const helperSheet = wb.getWorksheet('_categories');
    if (helperSheet) {
      helperSheet.eachRow({ includeEmpty: false }, (row, idx) => {
        if (idx === 1) return;
        const name = row.getCell(1).value;
        const id = row.getCell(2).value;
        if (name && id) categoryNameToId.set(String(name).trim(), Number(id));
      });
    }

    const changes = [];
    const colLet = (n) => { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
    linesSheet.eachRow({ includeEmpty: false }, (row, idx) => {
      if (idx === 1) return;
      const lineId = Number(row.getCell(COL_LINE_ID).value);
      if (!lineId) return;
      const newDesc = row.getCell(COL_NEW_DESC).value;
      const newCatRaw = COL_NEW_CAT ? row.getCell(COL_NEW_CAT).value : null;
      let newCategoryId = null;
      let newCategoryName = null;
      if (newCatRaw != null && String(newCatRaw).trim() !== '') {
        const trimmed = String(newCatRaw).trim();
        if (categoryNameToId.has(trimmed)) {
          newCategoryId = categoryNameToId.get(trimmed);
          newCategoryName = trimmed;
        } else {
          validationErrors.push(`Lines!${colLet(COL_NEW_CAT)}${idx}: new_category "${trimmed}" not in dropdown list`);
        }
      }
      const markedAttendees = [];
      for (let c = attendeeColStart; c <= headers.length; c++) {
        const v = row.getCell(c).value;
        if (v == null || v === '') continue;
        const sv = String(v).trim();
        if (sv.toLowerCase() === 'x') {
          const label = headers[c - 1];
          const att = labelToAttendee.get(label);
          if (att) markedAttendees.push(att);
          else validationErrors.push(`Lines!${colLet(c)}${idx}: attendee column header "${label}" not found in Attendees sheet`);
        } else {
          // Anything other than empty or 'x' is an invalid mark
          validationErrors.push(`Lines!${colLet(c)}${idx}: invalid value "${sv}" (must be blank or "x")`);
        }
      }
      if ((newDesc && String(newDesc).trim()) || markedAttendees.length || newCategoryId) {
        changes.push({
          line_id: lineId,
          new_description: newDesc ? String(newDesc).trim() : null,
          new_category_id: newCategoryId,
          new_category_name: newCategoryName,
          attendees: markedAttendees,
        });
      }
    });

    if (validationErrors.length) {
      const err = new Error(`Upload has ${validationErrors.length} invalid cell(s):\n` + validationErrors.slice(0, 10).join('\n') + (validationErrors.length > 10 ? `\n...and ${validationErrors.length - 10} more` : ''));
      err.validationErrors = validationErrors;
      throw err;
    }
    return { attendees, changes };
  }

  async function ensureAttendeeIds(attendees, csrf, statusFn) {
    const toCreate = attendees.filter(a => !a.id && (a.first_name || a.last_name));
    let created = 0;
    for (const a of toCreate) {
      try {
        statusFn(`Creating attendee ${++created}/${toCreate.length}: ${a.first_name} ${a.last_name}…`);
        const body = new URLSearchParams();
        body.append('expense_attendee[expense_attendee_type_id]', a.type_id || NEW_ATTENDEE_TYPE_ID);
        body.append('expense_attendee[first_name]', a.first_name);
        body.append('expense_attendee[last_name]', a.last_name);
        const r = await fetch('/expense_attendees/', {
          method: 'POST', credentials: 'include',
          headers: {
            'Accept': 'application/json, text/plain, */*',
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'X-CSRF-Token': csrf,
            'X-Requested-With': 'XMLHttpRequest',
          },
          body: body.toString(),
        });
        if (r.ok) {
          const j = await r.json();
          a.id = j.id;
        } else {
          throw new Error(`status ${r.status}`);
        }
      } catch (e) {
        a.error = String(e).slice(0, 100);
      }
    }
    return toCreate.filter(a => !a.id);  // names that failed
  }

  function buildLineUpdatePatchBody(line, change) {
    // Merge attendee ids: existing on line + new ones from change
    const existingIds = (line.expense_attendees || []).map(a => a.id);
    const newIds = (change.attendees || []).map(a => a.id).filter(x => x);
    const finalIds = Array.from(new Set([...existingIds, ...newIds]));
    const u = new URLSearchParams();
    const set = (k, v) => u.append(k, v == null ? '' : String(v));
    set('expense_line[custom_field_3]', line.custom_field_3 ?? '');
    set('expense_line[travel_provider_type]', line.travel_provider_type ?? '');
    set('expense_line[audit_status_id]', line.audit_status_id ?? '');
    set('expense_line[reason]', line.reason ?? '');
    set('expense_line[amount_to_receive]', line.amount_to_receive ?? '');
    // preserve line's existing account if any
    const acct = (line.accounts || [])[0];
    set('expense_line[account_id]', acct ? acct.account_id : '');
    set('expense_line[account_type_id]', acct ? acct.account_type_id : '');
    set('expense_line[merchant]', line.merchant ?? '');
    set('expense_line[local_expense_date]', line.local_expense_date ?? '');
    set('expense_line[parent_expense_line_id]', line.parent_expense_line_id ?? '');
    set('expense_line[start_date]', line.start_date ?? line.local_expense_date ?? '');
    set('expense_line[end_date]', line.end_date ?? line.local_expense_date ?? '');
    set('expense_line[travel_provider_name]', line.travel_provider_name ?? '');
    set('expense_line[expense_category_id]', change.new_category_id != null ? change.new_category_id : (line.expense_category_id ?? ''));
    set('expense_line[employee_reimbursable]', line.employee_reimbursable ? 'true' : 'false');
    set('expense_line[expense_category_custom_field_1]', line.expense_category_custom_field_1 ?? '');
    set('expense_line[receipt_total_currency_id]', line.receipt_currency?.id ?? '');
    set('expense_line[foreign_currency_id]', line.amount_to_receive_currency?.id ?? '');
    set('expense_line[external_src_name]', line.external_src_name ?? '');
    set('expense_line[exchange_rate]', line.exchange_rate ?? '');
    set('expense_line[description]', change.new_description != null ? change.new_description : (line.description ?? ''));
    set('expense_line[employee_reimbursable_overridden]', line.employee_reimbursable_overridden ? 'true' : 'false');
    set('expense_line[divisor]', line.divisor ?? '');
    set('expense_line[receipt_total_amount]', line.receipt_amount ?? '');
    set('expense_line[foreign_currency_amount]', line.amount_to_receive ?? '');
    set('expense_line[expense_report_id]', line.expense_report_id ?? '');
    set('expense_line[custom_field_2]', line.custom_field_2 ? 'true' : 'false');
    finalIds.forEach(id => u.append('expense_line[attendee_ids][]', id));
    return u.toString();
  }

  async function patchLine(lineId, body, csrf) {
    return fetch(`/expenses/expense_lines/${lineId}`, {
      method: 'PATCH', credentials: 'include',
      headers: {
        'Accept': 'application/json, text/plain, */*',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'X-CSRF-Token': csrf,
        'X-Requested-With': 'XMLHttpRequest',
      },
      body,
    });
  }

  async function findLineById(reports, lineId) {
    for (const rpt of reports) {
      const l = (rpt.expense_lines || []).find(x => x.id === lineId);
      if (l) return l;
    }
    return null;
  }

  function showFirstLineConfirm(panel, summary) {
    return new Promise((resolve) => {
      const overlay = document.createElement('div');
      overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.4);z-index:100000;display:flex;align-items:center;justify-content:center;';
      overlay.innerHTML = `
        <div style="background:#fff;padding:18px;border-radius:8px;font:13px sans-serif;max-width:520px;box-shadow:0 4px 16px rgba(0,0,0,0.3);">
          <h3 style="margin:0 0 10px 0;color:#06c;">Confirm first applied change</h3>
          <div style="margin-bottom:12px;color:#222;">${summary}</div>
          <div style="margin-bottom:12px;color:#666;font-size:12px;">Open this line in Coupa to verify the change looks right. Then choose:</div>
          <div style="display:flex;gap:8px;justify-content:flex-end;">
            <button class="__rf_confirm_no" type="button" style="padding:6px 14px;border:1px solid #ccc;background:#fff;border-radius:3px;cursor:pointer;">Stop</button>
            <button class="__rf_confirm_yes" type="button" style="padding:6px 14px;border:1px solid #048;background:#06c;color:#fff;border-radius:3px;cursor:pointer;">Apply rest</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);
      overlay.querySelector('.__rf_confirm_yes').addEventListener('click', () => { overlay.remove(); resolve(true); });
      overlay.querySelector('.__rf_confirm_no').addEventListener('click', () => { overlay.remove(); resolve(false); });
    });
  }

  async function applyUpload(panel, file) {
    const status = panel.querySelector('.__rf_acct_status');
    const progressWrap = panel.querySelector('.__rf_progress_wrap');
    const progressBar = panel.querySelector('.__rf_progress_bar');
    const progressText = panel.querySelector('.__rf_progress_text');
    const setProgress = (done, total) => {
      const pct = total > 0 ? Math.min(100, (done / total) * 100) : 0;
      progressBar.style.width = pct.toFixed(1) + '%';
      progressText.textContent = total > 0 ? `${done} / ${total} (${pct.toFixed(0)}%)` : '';
    };
    progressBar.style.background = '#06c';
    progressWrap.style.display = 'block';
    progressText.style.display = 'block';

    const beforeUnload = (e) => { e.preventDefault(); e.returnValue = ''; };
    window.addEventListener('beforeunload', beforeUnload);
    window.__rfAcctRunning = true;

    try {
      status.textContent = 'Parsing workbook…';
      const { attendees, changes } = await parseUploadedWorkbook(file);
      const csrf = document.querySelector('meta[name="csrf-token"]')?.getAttribute('content');
      const failedAtt = await ensureAttendeeIds(attendees, csrf, t => status.textContent = t);
      if (failedAtt.length) {
        status.textContent = `Could not create ${failedAtt.length} attendees; aborting. ` + failedAtt.map(a => `${a.first_name} ${a.last_name}: ${a.error || '?'}`).join('; ');
        return;
      }

      status.textContent = 'Fetching reports for line lookup…';
      const reports = await fetchAllDraftReports();

      if (!changes.length) {
        status.textContent = 'No changes detected in the upload.';
        return;
      }

      // Filter out no-op changes by comparing the spreadsheet's intent against the
      // line's CURRENT live state in Coupa. A "no-op" is a row whose new_description
      // equals the existing description AND whose new_category_id matches the line's
      // current category AND whose marked attendees are all already attached.
      const sameStr = (a, b) => (String(a || '').trim() === String(b || '').trim());
      const noOp = (line, ch) => {
        if (ch.new_description != null && !sameStr(ch.new_description, line.description)) return false;
        if (ch.new_category_id != null && Number(ch.new_category_id) !== Number(line.expense_category_id)) return false;
        const existingIds = new Set((line.expense_attendees || []).map(a => a.id));
        const additions = (ch.attendees || []).map(a => a.id).filter(id => id != null && !existingIds.has(id));
        if (additions.length > 0) return false;
        return true;
      };
      const realChanges = [];
      const skippedNoOp = [];
      for (const ch of changes) {
        const line = await findLineById(reports, ch.line_id);
        if (!line) { realChanges.push(ch); continue; } // let it fail loudly later
        if (noOp(line, ch)) {
          skippedNoOp.push(ch.line_id);
          continue;
        }
        // Build a "trimmed" change that contains only fields that actually changed,
        // so the PATCH body doesn't quietly re-set unchanged values.
        const trimmed = { line_id: ch.line_id };
        if (ch.new_description != null && !sameStr(ch.new_description, line.description)) {
          trimmed.new_description = ch.new_description;
        }
        if (ch.new_category_id != null && Number(ch.new_category_id) !== Number(line.expense_category_id)) {
          trimmed.new_category_id = ch.new_category_id;
          trimmed.new_category_name = ch.new_category_name;
        }
        const existingIds = new Set((line.expense_attendees || []).map(a => a.id));
        trimmed.attendees = (ch.attendees || []).filter(a => a.id != null && !existingIds.has(a.id));
        realChanges.push(trimmed);
      }

      if (!realChanges.length) {
        status.textContent = `No changes to apply — ${skippedNoOp.length} row(s) were already in the desired state.`;
        return;
      }

      // Apply the FIRST real change, ask user to confirm
      const first = realChanges[0];
      const firstLine = await findLineById(reports, first.line_id);
      if (!firstLine) {
        status.textContent = `First line ${first.line_id} not found in any draft report. Aborting.`;
        return;
      }
      status.textContent = `Applying first change (line ${first.line_id})…`;
      let r = await patchLine(first.line_id, buildLineUpdatePatchBody(firstLine, first), csrf);
      if (!r.ok) { status.textContent = `First PATCH failed (status ${r.status}). Aborting.`; return; }
      const summary =
        `<b>Line ${first.line_id}</b> &mdash; ${escapeHtml(firstLine.merchant || '')}<br>` +
        (first.new_category_name ? `&bull; category set to: <i>${escapeHtml(first.new_category_name)}</i><br>` : '') +
        (first.new_description ? `&bull; description set to: <i>${escapeHtml(first.new_description)}</i><br>` : '') +
        (first.attendees && first.attendees.length ? `&bull; attendees added: ${first.attendees.map(a => escapeHtml(a.first_name + ' ' + a.last_name)).join(', ')}<br>` : '') +
        (skippedNoOp.length ? `<div style="color:#888;font-size:11px;margin-top:6px;">(${skippedNoOp.length} other row(s) were already up-to-date and will be skipped.)</div>` : '');
      const ok = await showFirstLineConfirm(panel, summary);
      if (!ok) {
        status.textContent = `Stopped after first line. ${realChanges.length - 1} more were not applied. ${skippedNoOp.length} no-op rows skipped.`;
        return;
      }

      // Apply the rest. Counters align across loop, progress bar, and final tally:
      //   okCount starts at 1 because the first line was already PATCHed successfully.
      //   The progress bar denominator is realChanges.length (includes the first line).
      const rest = realChanges.slice(1);
      const totalToApply = realChanges.length;
      let okCount = 1, failCount = 0;
      const failures = [];
      setProgress(okCount + failCount, totalToApply);
      for (let i = 0; i < rest.length; i++) {
        const ch = rest[i];
        const line = await findLineById(reports, ch.line_id);
        if (!line) { failCount++; failures.push({ line_id: ch.line_id, error: 'not found' }); }
        else {
          try {
            const resp = await patchLine(ch.line_id, buildLineUpdatePatchBody(line, ch), csrf);
            if (resp.ok) okCount++;
            else { failCount++; failures.push({ line_id: ch.line_id, status: resp.status, head: (await resp.text()).slice(0, 200) }); }
          } catch (e) { failCount++; failures.push({ line_id: ch.line_id, error: String(e).slice(0, 200) }); }
        }
        setProgress(okCount + failCount, totalToApply);
        status.textContent = `ok=${okCount} fail=${failCount} skipped=${skippedNoOp.length}`;
        await new Promise(r => setTimeout(r, 120));
      }

      status.innerHTML = `<b>Done.</b> ok=${okCount} / fail=${failCount} / no-op skipped=${skippedNoOp.length}` +
        (failures.length ? ` <a href="data:application/json;base64,${btoa(JSON.stringify(failures, null, 2))}" download="upload-apply-failures.json" style="color:#06c;">download failures</a>` : '');
      progressBar.style.background = failCount > 0 ? '#c60' : '#0a7';
    } catch (e) {
      const msg = (e && e.message ? e.message : String(e));
      if (e && e.validationErrors) {
        const blob = new Blob([e.validationErrors.join('\n')], { type: 'text/plain' });
        const url = URL.createObjectURL(blob);
        status.innerHTML = `Upload aborted: ${e.validationErrors.length} invalid cell(s). <a href="${url}" download="upload-validation-errors.txt" style="color:#06c;">download list</a>`;
      } else {
        status.textContent = 'Error: ' + msg.slice(0, 200);
      }
    } finally {
      window.__rfAcctRunning = false;
      window.removeEventListener('beforeunload', beforeUnload);
    }
  }

  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  // ---------- receipt auto-match ----------
  const MATCH_TOL_SAME_CCY = 0.01;
  const MATCH_TOL_CROSS_CCY = 0.12; // FX spread on corp cards is often 8-12%
  const STOP_WORDS = new Set([
    'the','of','and','llc','ltd','inc','co','corp','sa','sas','bv','gmbh','ag','sl','spa','srl',
    'merchant','copy','customer','receipt','cia','de','la','el','los','las','del','d',
    'mc','mcd','mcdo','mcdonald','mcdonalds','com','www','io','tm','to','at','for','in','on',
    'a','an','as','is','it','no','or','by','be','re','via',
    'order','transaction','transactie','transact','recibo','tienda','bar','cafe','restaurant','restaurante','pizza','pizzeria','food','foods','grill','steakhouse','sushi','kitchen',
    'paris','amsterdam','antwerpen','antwerp','london','madrid','miami','dallas','philadelphia','cartagena','bocagrande','luxembourg','luxenbourg','luxenburgo','watermael',
    'mastercard','visa','card','controle','controlegegevens','airlines','airline','airport','flight','flights',
    'meal','breakfast','lunch','dinner','snack','drink',
  ]);
  function _tokens(s) {
    if (!s) return new Set();
    return new Set(String(s).toLowerCase().replace(/[^a-z0-9 ]+/g, ' ').split(/\s+/).filter(t => t.length >= 3 && !/^\d+$/.test(t) && !STOP_WORDS.has(t)));
  }
  function _overlap(a, b) {
    let n = 0;
    for (const t of a) if (b.has(t)) n++;
    return n;
  }

  function readWalletReceipts() {
    return Array.from(document.querySelectorAll('li.walletLine')).map(li => {
      const img = li.querySelector('img.s-walletReceiptImg');
      const m = img?.src?.match(/expense_lines\/(\d+)\/expense_artifacts\/(\d+)/);
      if (!m) return null;
      const merchant = (li.querySelector('.s-receiptDescription')?.textContent || '').trim();
      const filename = img.alt || '';
      const amountText = (li.querySelector('.s-receiptAmount')?.textContent || '').trim();
      const currency = (li.querySelector('.currency_code')?.textContent || '').trim().toUpperCase();
      const amount = parseAmt(amountText);
      if (!isFinite(amount) || !currency) return null;
      return {
        artifact_id: m[2],
        wallet_line_id: m[1],
        merchant, filename, amount, currency,
        _tokens: new Set([..._tokens(merchant), ..._tokens(filename)]),
      };
    }).filter(Boolean);
  }

  function buildMatches(reports, receipts, usdRates) {
    const usdRate = c => c === 'USD' ? 1 : (usdRates[c] || null);
    const used = new Set();
    const matches = [];
    const lines = [];
    reports.forEach(r => (r.expense_lines || []).forEach(l => {
      if ((l.expense_artifacts || []).length) return; // already has receipt
      const amt = parseAmt(l.receipt_amount);
      const cur = (l.receipt_currency?.code || '').toUpperCase();
      if (!isFinite(amt) || !cur) return;
      lines.push({
        line_id: l.id,
        report_id: r.id,
        merchant: l.merchant || '',
        amount: amt,
        currency: cur,
        date: l.local_expense_date,
        _tokens: _tokens(l.merchant),
      });
    }));
    // Largest first reduces ambiguity
    lines.sort((a, b) => {
      const aUsd = usdRate(a.currency) ? a.amount / usdRate(a.currency) : a.amount;
      const bUsd = usdRate(b.currency) ? b.amount / usdRate(b.currency) : b.amount;
      return bUsd - aUsd;
    });

    for (const line of lines) {
      let best = null, bestScore = -Infinity;
      for (const r of receipts) {
        if (used.has(r.artifact_id)) continue;
        let score = -Infinity;
        let tier = null;
        const sameCcy = r.currency === line.currency;
        const overlap = _overlap(line._tokens, r._tokens);
        if (sameCcy) {
          // Tier A: exact amount → very high
          if (Math.abs(r.amount - line.amount) < 0.01) {
            score = 1000 + overlap * 5;
            tier = 'exact';
          } else if (Math.abs(r.amount - line.amount) <= line.amount * MATCH_TOL_SAME_CCY) {
            // Tier B: same currency, ±1%, must have ≥1 token overlap
            if (overlap >= 1) {
              score = 500 + overlap * 10 - Math.abs(r.amount - line.amount);
              tier = 'close_same_ccy';
            }
          }
        } else {
          // Tier C: cross-currency, FX-window with token overlap required
          const lUsd = usdRate(line.currency) ? line.amount / usdRate(line.currency) : null;
          const rUsd = usdRate(r.currency) ? r.amount / usdRate(r.currency) : null;
          if (lUsd && rUsd) {
            const diffPct = Math.abs(lUsd - rUsd) / Math.max(lUsd, 1);
            if (diffPct <= MATCH_TOL_CROSS_CCY && overlap >= 1) {
              score = 100 + overlap * 20 - diffPct * 100;
              tier = 'cross_ccy';
            }
          }
        }
        if (score > bestScore) { bestScore = score; best = r ? { receipt: r, tier, overlap } : null; }
      }
      if (best && bestScore > -Infinity) {
        used.add(best.receipt.artifact_id);
        matches.push({ line, receipt: best.receipt, tier: best.tier, overlap: best.overlap, score: bestScore });
      }
    }
    return matches;
  }

  async function runMatchReceipts(panel) {
    const status = panel.querySelector('.__rf_acct_status');
    const progressWrap = panel.querySelector('.__rf_progress_wrap');
    const progressBar = panel.querySelector('.__rf_progress_bar');
    const progressText = panel.querySelector('.__rf_progress_text');
    const matchBtn = panel.querySelector('.__rf_match_btn');
    const setProgress = (done, total) => {
      const pct = total > 0 ? Math.min(100, (done / total) * 100) : 0;
      progressBar.style.width = pct.toFixed(1) + '%';
      progressText.textContent = total > 0 ? `${done} / ${total} (${pct.toFixed(0)}%)` : '';
    };
    progressBar.style.background = '#06c';
    progressWrap.style.display = 'block';
    progressText.style.display = 'block';
    matchBtn.disabled = true; matchBtn.style.opacity = '0.6'; matchBtn.style.cursor = 'wait'; matchBtn.textContent = 'Matching…';
    const beforeUnload = (e) => { e.preventDefault(); e.returnValue = ''; };
    window.addEventListener('beforeunload', beforeUnload);
    window.__rfAcctRunning = true;
    try {
      status.textContent = 'Reading wallet & reports…';
      const receipts = readWalletReceipts();
      if (!receipts.length) { status.textContent = 'No wallet receipts found in DOM. Make sure you are on the Expenses page with the wallet visible.'; return; }
      const reports = await fetchAllDraftReports();
      const usdRates = await fetchFxToUSD();
      status.textContent = `${receipts.length} receipts, computing matches…`;
      const matches = buildMatches(reports, receipts, usdRates);
      if (!matches.length) { status.textContent = `No matches found across ${receipts.length} wallet receipts.`; return; }
      const tierCounts = matches.reduce((acc, m) => { acc[m.tier] = (acc[m.tier] || 0) + 1; return acc; }, {});

      const csrf = document.querySelector('meta[name="csrf-token"]')?.getAttribute('content');
      // Apply FIRST match, ask user to confirm
      const first = matches[0];
      status.textContent = `Applying first match (line ${first.line.line_id})…`;
      const r = await fetch('/expenses/wallet/merge_receipt_to_expense_line', {
        method: 'POST', credentials: 'include',
        headers: {
          'Accept': 'application/json, text/plain, */*',
          'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
          'X-CSRF-Token': csrf,
          'X-Requested-With': 'XMLHttpRequest',
        },
        body: `expense_line_id=${encodeURIComponent(first.line.line_id)}&wallet_expense_line_id=${encodeURIComponent(first.receipt.wallet_line_id)}`,
      });
      if (!r.ok) { status.textContent = `First merge failed (status ${r.status}). Aborting.`; return; }
      const summary =
        `<b>${matches.length} candidate matches found.</b><br>` +
        `&bull; Tiers: ${Object.entries(tierCounts).map(([k,v]) => `${k}=${v}`).join(', ')}<br>` +
        `<b>First applied:</b><br>` +
        `&bull; Line <b>${first.line.line_id}</b> &mdash; ${escapeHtml(first.line.merchant.slice(0,40))} ${first.line.amount} ${first.line.currency}<br>` +
        `&bull; Receipt &mdash; ${escapeHtml(first.receipt.merchant.slice(0,40))} ${first.receipt.amount} ${first.receipt.currency} (${escapeHtml(first.receipt.filename.slice(0,40))})<br>` +
        `&bull; Tier: ${first.tier}, overlap=${first.overlap}`;
      const ok = await showFirstLineConfirm(panel, summary);
      if (!ok) {
        status.textContent = `Stopped after first match. ${matches.length - 1} candidates not applied.`;
        return;
      }
      const rest = matches.slice(1);
      const total = matches.length;
      let okCount = 1, failCount = 0;
      const failures = [];
      setProgress(okCount + failCount, total);
      for (const m of rest) {
        try {
          const resp = await fetch('/expenses/wallet/merge_receipt_to_expense_line', {
            method: 'POST', credentials: 'include',
            headers: {
              'Accept': 'application/json, text/plain, */*',
              'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
              'X-CSRF-Token': csrf,
              'X-Requested-With': 'XMLHttpRequest',
            },
            body: `expense_line_id=${encodeURIComponent(m.line.line_id)}&wallet_expense_line_id=${encodeURIComponent(m.receipt.wallet_line_id)}`,
          });
          if (resp.ok) okCount++;
          else { failCount++; failures.push({ line_id: m.line.line_id, artifact_id: m.receipt.artifact_id, status: resp.status, head: (await resp.text()).slice(0, 200) }); }
        } catch (e) { failCount++; failures.push({ line_id: m.line.line_id, error: String(e).slice(0,200) }); }
        setProgress(okCount + failCount, total);
        status.textContent = `ok=${okCount} fail=${failCount}`;
        await new Promise(r => setTimeout(r, 250));
      }
      status.innerHTML = `<b>Match complete.</b> ok=${okCount} / fail=${failCount}` +
        (failures.length ? ` <a href="data:application/json;base64,${btoa(JSON.stringify(failures, null, 2))}" download="match-receipts-failures.json" style="color:#06c;">download failures</a>` : '');
      progressBar.style.background = failCount > 0 ? '#c60' : '#0a7';
    } catch (e) {
      status.textContent = 'Error: ' + ((e && e.message) ? e.message : String(e)).slice(0, 200);
    } finally {
      window.__rfAcctRunning = false;
      window.removeEventListener('beforeunload', beforeUnload);
      matchBtn.disabled = false; matchBtn.style.opacity = ''; matchBtn.style.cursor = 'pointer'; matchBtn.textContent = 'Match Receipts';
    }
  }

  // ---------- account selector ----------
  function refreshAccountDisplay(panel) {
    const a = getActiveAccount();
    const nameEl = panel.querySelector('.__rf_acct_name');
    const codeEl = panel.querySelector('.__rf_acct_code');
    if (nameEl) nameEl.textContent = a.display_name || `(account ${a.account_id})`;
    if (codeEl) codeEl.textContent = `id ${a.account_id}` + (a.code ? ` · ${a.code}` : '');
  }

  // Best-effort autocomplete against Coupa. Different tenants expose the dropdown
  // through different endpoints; we probe a few and use whichever returns JSON.
  let _knownAcctSearchUrl = null;
  async function searchAccountsCoupa(term) {
    if (!term || String(term).trim().length < 2) return [];
    const enc = encodeURIComponent(term);
    const candidates = _knownAcctSearchUrl
      ? [_knownAcctSearchUrl.replace('__TERM__', enc)]
      : [
          `/accounts.json?term=${enc}`,
          `/accounts/autocomplete?term=${enc}`,
          `/accounts/search.json?q=${enc}`,
          `/accounts/lookup?term=${enc}`,
          `/accounts/search?term=${enc}`,
        ];
    for (const url of candidates) {
      try {
        const r = await fetch(url, { credentials: 'include', headers: { Accept: 'application/json', 'X-Requested-With': 'XMLHttpRequest' } });
        if (!r.ok) continue;
        const txt = await r.text();
        let j; try { j = JSON.parse(txt); } catch { continue; }
        const list = Array.isArray(j) ? j : (Array.isArray(j.results) ? j.results : null);
        if (Array.isArray(list) && list.length && (list[0].id != null || list[0].account_id != null)) {
          if (!_knownAcctSearchUrl) _knownAcctSearchUrl = url.replace(enc, '__TERM__');
          return list.map(x => ({
            id: x.id || x.account_id,
            account_type_id: x.account_type_id,
            name: x.name || x.display_name || x.label,
            code: x.code || x.account_code,
          })).filter(x => x.id);
        }
      } catch {}
    }
    return [];
  }

  async function fetchAccountById(id) {
    const idNum = Number(id);
    if (!isFinite(idNum) || idNum <= 0) return null;
    const candidates = [
      `/accounts/${idNum}.json`,
      `/accounts/${idNum}`,
    ];
    for (const url of candidates) {
      try {
        const r = await fetch(url, { credentials: 'include', headers: { Accept: 'application/json', 'X-Requested-With': 'XMLHttpRequest' } });
        if (!r.ok) continue;
        const txt = await r.text();
        let j; try { j = JSON.parse(txt); } catch { continue; }
        if (j && (j.id != null || j.account_id != null)) {
          return {
            id: j.id || j.account_id,
            account_type_id: j.account_type_id,
            name: j.name,
            code: j.code,
          };
        }
      } catch {}
    }
    return null;
  }

  function showAccountEditor(panel) {
    if (document.getElementById('__rf_acct_modal')) return;
    const overlay = document.createElement('div');
    overlay.id = '__rf_acct_modal';
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.45);z-index:100002;display:flex;align-items:center;justify-content:center;font:13px sans-serif;color:#222;';
    const cur = getActiveAccount();
    overlay.innerHTML = `
      <div style="background:#fff;padding:20px 24px;border-radius:8px;min-width:520px;max-width:92vw;max-height:80vh;overflow:auto;box-shadow:0 6px 22px rgba(0,0,0,0.3);">
        <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px;">
          <h2 style="margin:0;color:#06c;">Pick an account</h2>
          <button class="__rf_acct_modal_close" type="button" style="background:transparent;border:none;color:#888;font-size:22px;cursor:pointer;line-height:1;padding:0;">&times;</button>
        </div>
        <div style="margin:6px 0 14px 0;color:#666;font-size:12px;">Currently set: <b>${escapeHtml(cur.display_name || `(id ${cur.account_id})`)}</b> <span style="color:#888;">id ${cur.account_id}</span></div>

        <div style="margin-bottom:14px;">
          <label style="display:block;font-weight:bold;margin-bottom:4px;">Search by name or code</label>
          <input class="__rf_acct_search" type="text" placeholder="type to search…" style="width:100%;padding:6px 8px;border:1px solid #aaa;border-radius:3px;box-sizing:border-box;">
          <div class="__rf_acct_results" style="margin-top:6px;max-height:240px;overflow:auto;border:1px solid #ddd;border-radius:3px;display:none;"></div>
          <div class="__rf_acct_search_status" style="font-size:11px;color:#888;margin-top:4px;"></div>
        </div>

        <div style="margin-bottom:14px;border-top:1px solid #eee;padding-top:14px;">
          <label style="display:block;font-weight:bold;margin-bottom:4px;">Or paste an Account ID</label>
          <div style="display:flex;gap:6px;">
            <input class="__rf_acct_id" type="number" placeholder="e.g. 6222" style="flex:1;padding:6px 8px;border:1px solid #aaa;border-radius:3px;">
            <button class="__rf_acct_use_id" type="button" style="padding:6px 14px;border:1px solid #048;background:#06c;color:#fff;border-radius:3px;cursor:pointer;">Use ID</button>
          </div>
        </div>

        <div style="text-align:right;border-top:1px solid #eee;padding-top:12px;">
          <button class="__rf_acct_modal_cancel" type="button" style="padding:6px 14px;border:1px solid #ccc;background:#fff;border-radius:3px;cursor:pointer;">Cancel</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    const close = () => overlay.remove();
    overlay.querySelector('.__rf_acct_modal_close').addEventListener('click', close);
    overlay.querySelector('.__rf_acct_modal_cancel').addEventListener('click', close);
    overlay.addEventListener('click', e => { if (e.target === overlay) close(); });

    const searchInput = overlay.querySelector('.__rf_acct_search');
    const resultsBox = overlay.querySelector('.__rf_acct_results');
    const searchStatus = overlay.querySelector('.__rf_acct_search_status');
    let deb = null;
    searchInput.addEventListener('input', () => {
      clearTimeout(deb);
      const term = searchInput.value;
      if (!term || term.trim().length < 2) {
        resultsBox.style.display = 'none';
        searchStatus.textContent = '';
        return;
      }
      searchStatus.textContent = 'searching…';
      deb = setTimeout(async () => {
        const list = await searchAccountsCoupa(term);
        if (!list.length) {
          resultsBox.style.display = 'none';
          searchStatus.textContent = `No results (Coupa search endpoint may not be available — paste the Account ID below instead).`;
          return;
        }
        searchStatus.textContent = `${list.length} result(s)`;
        resultsBox.innerHTML = list.slice(0, 50).map(r => {
          return `<div class="__rf_acct_result_row" data-id="${r.id}" data-name="${escapeHtml(r.name || '')}" data-code="${escapeHtml(r.code || '')}" data-type="${r.account_type_id || ''}" style="padding:6px 8px;border-bottom:1px solid #eee;cursor:pointer;">
            <div style="font-weight:bold;">${escapeHtml(r.name || '(unnamed)')}</div>
            <div style="font-size:11px;color:#888;">id ${r.id}${r.code ? ' · ' + escapeHtml(r.code) : ''}</div>
          </div>`;
        }).join('');
        resultsBox.style.display = 'block';
      }, 280);
    });
    resultsBox.addEventListener('click', e => {
      const row = e.target.closest('.__rf_acct_result_row');
      if (!row) return;
      const id = Number(row.dataset.id);
      const name = row.dataset.name;
      const code = row.dataset.code;
      const account_type_id = Number(row.dataset.type) || cur.account_type_id || DEFAULT_ACCOUNT.account_type_id;
      setActiveAccount({ account_id: id, account_type_id, display_name: name, code });
      refreshAccountDisplay(panel);
      close();
    });

    overlay.querySelector('.__rf_acct_use_id').addEventListener('click', async () => {
      const idStr = overlay.querySelector('.__rf_acct_id').value;
      const id = Number(idStr);
      if (!isFinite(id) || id <= 0) { alert('Enter a positive numeric Account ID.'); return; }
      searchStatus.textContent = `Looking up id ${id}…`;
      const acct = await fetchAccountById(id);
      const next = {
        account_id: id,
        account_type_id: (acct && acct.account_type_id) || cur.account_type_id || DEFAULT_ACCOUNT.account_type_id,
        display_name: (acct && acct.name) || `(account ${id})`,
        code: (acct && acct.code) || '',
      };
      setActiveAccount(next);
      refreshAccountDisplay(panel);
      close();
    });
  }

  // ---------- help modal ----------
  function showHelpModal() {
    if (document.getElementById('__rf_help_modal')) return;
    const accent = pickHelpAccent();
    // Console log so the user can verify the script is running v0.8.4+ and a color was picked
    try { console.log(`[CoupaReceiptFilter v${SCRIPT_VERSION}] Help modal accent: ${accent}`); } catch {}
    const overlay = document.createElement('div');
    overlay.id = '__rf_help_modal';
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.45);z-index:100001;display:flex;align-items:center;justify-content:center;font:13px sans-serif;color:#222;';
    overlay.innerHTML = `
      <div style="background:#fff;padding:22px 26px 22px 36px;border-radius:8px;max-width:680px;max-height:88vh;overflow:auto;box-shadow:0 6px 22px rgba(0,0,0,0.3);border-left:10px solid ${accent};border-top:4px solid ${accent};position:relative;">
        <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px;">
          <h2 style="margin:0 0 8px 0;color:${accent};">Coupa Receipt Filter — How it works</h2>
          <button class="__rf_help_close" type="button" style="background:transparent;border:none;color:#888;font-size:22px;line-height:1;cursor:pointer;">&times;</button>
        </div>
        <p style="margin:6px 0 14px 0;color:#666;">A Tampermonkey userscript that bulk-edits Coupa expense reports. Lives at top-right of any Coupa <code>/expenses*</code> page and operates only on <b>draft</b> reports. <span style="color:${accent};font-weight:bold;">v${SCRIPT_VERSION}</span></p>
        <h3 style="margin:14px 0 4px 0;color:${accent};">Buttons</h3>
        <ul style="margin:0 0 0 18px;padding:0;line-height:1.55;">
          <li><b>Apply Account to All</b> — PATCHes every draft line whose <code>accounts[]</code> doesn't already include the configured account (currently id <code>${getActiveAccount().account_id}</code>, ${escapeHtml(getActiveAccount().display_name)}). Skips lines that are already set, so re-running is safe and fast.</li>
          <li><b>Account selector</b> — type to search Coupa accounts (uses Coupa's own autocomplete). The ↺ button reverts to the script default (id <code>${DEFAULT_ACCOUNT.account_id}</code>, ${escapeHtml(DEFAULT_ACCOUNT.display_name)}). The selection is stored in your browser's localStorage and reused by Apply Account to All + the Upload &amp; Apply pipeline.</li>
          <li><b>Match Receipts</b> — scans the wallet sidebar (<code>li.walletLine</code>) and pairs receipts to draft lines that don't yet have one. Tiers (highest score wins): exact same currency + amount → ±1% same currency + ≥1 token of merchant overlap → cross-currency within ±12% USD-equivalent + token overlap. Posts <code>/expenses/wallet/merge_receipt_to_expense_line</code> per match. Confirms the first match before applying the rest.</li>
          <li><b>Download Non-Compliant</b> — generates an .xlsx of every draft line with at least one problem: missing receipt &gt; $25 USD-eq, missing account, missing category, or gift-meal whose value-per-attendee &gt; $25. Editable columns get <span style="background:#C8E6C9;padding:0 4px;">green headers</span>. Invalid cells light <span style="background:#FF6B6B;color:#fff;padding:0 4px;">red</span> (live conditional formatting).</li>
          <li><b>Export All</b> — same xlsx structure but includes every draft line, not just non-compliant ones. Useful for a complete audit pass.</li>
          <li><b>Upload &amp; Apply</b> — opens a file picker, ingests the edited xlsx, and PATCHes the lines. Compares each row against the line's live state and skips no-op rows. Confirms the first applied change before continuing.</li>
        </ul>
        <h3 style="margin:14px 0 4px 0;color:${accent};">Workbook layout (Lines sheet)</h3>
        <ul style="margin:0 0 0 18px;padding:0;line-height:1.55;">
          <li>Read-only context: <code>line_id, report, merchant, date, amount, currency, usd_eq, current_category, problems, current_description, current_attendees</code>.</li>
          <li>Editable (green header): <code>new_category</code> (dropdown of categories you've used), <code>new_description</code> (free text), <code>&lt;person&gt;</code> attendee columns (mark <code>x</code> to add).</li>
          <li>Hidden helper sheet <code>_categories</code> drives both the dropdown and the upload's name→id resolution. Don't delete it.</li>
        </ul>
        <h3 style="margin:14px 0 4px 0;color:${accent};">Workbook layout (Attendees sheet)</h3>
        <ul style="margin:0 0 0 18px;padding:0;line-height:1.55;">
          <li>Existing attendees you've used appear with their Coupa <code>id</code>, <code>type_id</code> (5 = Coupa user, 6 = manual entry), <code>first_name</code>, <code>last_name</code>.</li>
          <li>To add a NEW attendee: append a row, leave <code>id</code> blank, set first &amp; last name. On upload, the script POSTs to <code>/expense_attendees/</code> to create them, then uses the returned id when applying any rows that reference that name.</li>
          <li>To use a new attendee on a line, you must also add a column on Lines named <code>"first last"</code> (without an id) and put <code>x</code> in the row.</li>
        </ul>
        <h3 style="margin:14px 0 4px 0;color:${accent};">Validation &amp; safety</h3>
        <ul style="margin:0 0 0 18px;padding:0;line-height:1.55;">
          <li>Conditional formatting marks invalid cells red live in Excel; the upload validates the same rules and aborts before any PATCH if it finds one (with a downloadable list of bad cells).</li>
          <li>Yellow row tint = informational only (gift-meal value-per-attendee &gt; $25). Doesn't block anything.</li>
          <li>While bulk operations are running, the page warns "Are you sure you want to leave?" if you try to navigate.</li>
          <li>The first PATCH/merge of any bulk action requires confirmation in a modal so you can verify it took before the rest run.</li>
          <li>Counters always reconcile: <code>ok + fail + skipped</code> = total rows the script considered.</li>
        </ul>
        <div style="text-align:right;margin-top:16px;">
          <button class="__rf_help_close2" type="button" style="background:${accent};color:#fff;border:1px solid ${accent};padding:6px 14px;border-radius:3px;cursor:pointer;">Got it</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    const close = () => overlay.remove();
    overlay.querySelector('.__rf_help_close').addEventListener('click', close);
    overlay.querySelector('.__rf_help_close2').addEventListener('click', close);
    overlay.addEventListener('click', e => { if (e.target === overlay) close(); });
  }

  // ---------- persistent top-right account panel ----------
  function isExpensesPage() {
    return /\/expense(?:s|_reports)/i.test(location.pathname);
  }

  function mountAccountPanel() {
    if (!isExpensesPage()) return;
    if (document.getElementById('__rf_acct_panel')) return;
    // Pick a fresh accent color for the entire panel each load
    const accent = pickHelpAccent();
    // Slightly darker shade for borders on solid buttons (just darken hex by 30%)
    const darken = (hex) => {
      const m = hex.replace('#', '');
      const n = parseInt(m, 16);
      const r = Math.max(0, ((n >> 16) & 0xff) - 40);
      const g = Math.max(0, ((n >> 8) & 0xff) - 40);
      const b = Math.max(0, (n & 0xff) - 40);
      return '#' + [r, g, b].map(x => x.toString(16).padStart(2, '0')).join('');
    };
    const accentDark = darken(accent);
    try { console.log(`[CoupaReceiptFilter v${SCRIPT_VERSION}] Panel accent: ${accent}`); } catch {}
    const panel = document.createElement('div');
    panel.id = '__rf_acct_panel';
    panel.style.cssText = [
      'position:fixed',
      'top:10px',
      'right:10px',
      'z-index:99999',
      'padding:8px 10px',
      'background:#fff',
      `border:1px solid ${accent}`,
      'border-radius:6px',
      'font:12px sans-serif',
      'box-shadow:0 2px 8px rgba(0,0,0,.15)',
      'width:170px',
      'box-sizing:border-box',
      'color:#222',
      'word-wrap:break-word',
      'overflow-wrap:break-word',
    ].join(';');
    panel.innerHTML = `
      <div style="display:flex;justify-content:flex-end;">
        <button class="__rf_panel_collapse" type="button" title="Hide" style="background:transparent;border:none;cursor:pointer;color:#888;font-size:14px;line-height:1;padding:0;">&times;</button>
      </div>
      <button class="__rf_apply_account_btn" type="button" style="background:${accent};color:#fff;border:1px solid ${accentDark};padding:5px 8px;border-radius:3px;cursor:pointer;width:100%;white-space:nowrap;">Apply Account to All</button>
      <button class="__rf_match_btn" type="button" style="margin-top:5px;background:${accent};color:#fff;border:1px solid ${accentDark};padding:5px 8px;border-radius:3px;cursor:pointer;width:100%;white-space:nowrap;">Match Receipts</button>
      <button class="__rf_download_problems_btn" type="button" style="margin-top:5px;background:#fff;color:${accent};border:1px solid ${accent};padding:5px 8px;border-radius:3px;cursor:pointer;width:100%;white-space:nowrap;">Export Non-Compliant</button>
      <button class="__rf_export_all_btn" type="button" style="margin-top:5px;background:#fff;color:${accent};border:1px solid ${accent};padding:5px 8px;border-radius:3px;cursor:pointer;width:100%;white-space:nowrap;">Export All</button>
      <button class="__rf_upload_btn" type="button" style="margin-top:5px;background:#fff;color:${accent};border:1px solid ${accent};padding:5px 8px;border-radius:3px;cursor:pointer;width:100%;white-space:nowrap;">Upload &amp; Apply</button>
      <input class="__rf_upload_file" type="file" accept=".xlsx" style="display:none;">
      <div class="__rf_progress_wrap" style="display:none;margin-top:6px;height:10px;background:#eef;border:1px solid #abc;border-radius:4px;overflow:hidden;">
        <div class="__rf_progress_bar" style="height:100%;width:0%;background:${accent};transition:width 200ms ease;"></div>
      </div>
      <div class="__rf_progress_text" style="display:none;margin-top:3px;font-size:10px;color:${accent};text-align:center;font-variant-numeric:tabular-nums;"></div>
      <div class="__rf_acct_block" style="margin-top:6px;font-size:11px;line-height:1.3;word-break:break-word;">
        <div style="display:flex;align-items:center;gap:4px;">
          <b style="flex:0 0 auto;">Account:</b>
          <button class="__rf_acct_edit" type="button" title="Pick a different account" style="flex:0 0 auto;background:transparent;border:none;cursor:pointer;color:${accent};font-size:11px;padding:0;">&#9998;</button>
          <button class="__rf_acct_reset" type="button" title="Revert to script default" style="flex:0 0 auto;background:transparent;border:none;cursor:pointer;color:#888;font-size:13px;padding:0;">&#x21BA;</button>
        </div>
        <div class="__rf_acct_name" style="margin-top:2px;"></div>
        <div class="__rf_acct_code" style="color:#888;"></div>
      </div>
      <div class="__rf_acct_status" style="margin-top:6px;font-size:11px;color:${accent};min-height:14px;word-break:break-word;"></div>
      <div style="margin-top:8px;padding-top:6px;border-top:1px solid #eee;display:flex;align-items:center;justify-content:space-between;">
        <span class="__rf_version" style="font-size:10px;color:#888;">v${SCRIPT_VERSION}</span>
        <span style="display:flex;gap:6px;">
          <button class="__rf_panel_update" type="button" title="Check for / install latest version" style="background:transparent;border:1px solid #888;border-radius:50%;width:20px;height:20px;cursor:pointer;color:${accent};font-size:13px;line-height:1;padding:0;">&#x21BB;</button>
          <button class="__rf_panel_help" type="button" title="How this works" style="background:transparent;border:1px solid #888;border-radius:50%;width:20px;height:20px;cursor:pointer;color:${accent};font-size:11px;font-weight:bold;line-height:1;padding:0;">?</button>
        </span>
      </div>
    `;
    document.body.appendChild(panel);
    panel.querySelector('.__rf_panel_collapse').addEventListener('click', () => panel.remove());
    panel.querySelector('.__rf_apply_account_btn').addEventListener('click', async () => {
      const status = panel.querySelector('.__rf_acct_status');
      if (window.__rfAcctRunning) { status.textContent = 'Already running…'; return; }
      status.textContent = 'Finding draft reports…';
      const draftIds = await findDraftReportIds();
      if (!draftIds.length) {
        status.textContent = 'No draft reports found. Are you signed in to Coupa expenses?';
        return;
      }
      if (!confirm(`Apply account ${getActiveAccount().account_id} (${getActiveAccount().display_name}) to every line in ${draftIds.length} draft reports?\n\nLines that already have this account will be skipped.\n\nThis is a bulk PATCH that may take several minutes.`)) {
        status.textContent = 'cancelled';
        return;
      }
      runApplyAccountToAll(draftIds, panel);
    });

    // Match Receipts
    panel.querySelector('.__rf_match_btn').addEventListener('click', async () => {
      if (window.__rfAcctRunning) {
        panel.querySelector('.__rf_acct_status').textContent = 'Already running…';
        return;
      }
      if (!confirm('Auto-match wallet receipts to draft expense lines?\n\nMatches by:\n• same currency + exact amount\n• same currency + ±1% amount + similar merchant\n• cross-currency + ±12% USD-eq + similar merchant\n\nFirst match needs your confirmation before the rest are applied.')) return;
      runMatchReceipts(panel);
    });

    // Download Non-Compliant
    panel.querySelector('.__rf_download_problems_btn').addEventListener('click', async () => {
      const status = panel.querySelector('.__rf_acct_status');
      try {
        const { blob, problemRows } = await buildProblemsWorkbook(panel, { onlyProblems: true });
        downloadBlob(blob, 'coupa-non-compliant.xlsx');
        status.textContent = `Downloaded ${problemRows} non-compliant rows.`;
      } catch (e) {
        status.textContent = 'Download failed: ' + (e && e.message ? e.message : String(e)).slice(0, 200);
      }
    });
    // Export All
    panel.querySelector('.__rf_export_all_btn').addEventListener('click', async () => {
      const status = panel.querySelector('.__rf_acct_status');
      try {
        const { blob, problemRows } = await buildProblemsWorkbook(panel, { onlyProblems: false });
        downloadBlob(blob, 'coupa-all-draft-lines.xlsx');
        status.textContent = `Downloaded ${problemRows} rows (all draft lines).`;
      } catch (e) {
        status.textContent = 'Download failed: ' + (e && e.message ? e.message : String(e)).slice(0, 200);
      }
    });
    // Help (?)
    panel.querySelector('.__rf_panel_help').addEventListener('click', () => showHelpModal());
    // Update (↻) — opens the @updateURL in a new tab with a cache-buster so
    // GitHub's 5-min CDN cache can't serve a stale revision. Tampermonkey detects
    // the .user.js URL and prompts to install/update.
    panel.querySelector('.__rf_panel_update').addEventListener('click', () => {
      const url = SCRIPT_UPDATE_URL + (SCRIPT_UPDATE_URL.includes('?') ? '&' : '?') + 'cb=' + Date.now();
      window.open(url, '_blank', 'noopener');
    });

    // Account editor (✎)
    panel.querySelector('.__rf_acct_edit').addEventListener('click', () => showAccountEditor(panel));
    // Account reset (↺)
    panel.querySelector('.__rf_acct_reset').addEventListener('click', () => {
      if (!confirm(`Reset account to script default?\n\n${DEFAULT_ACCOUNT.display_name} (id ${DEFAULT_ACCOUNT.account_id})`)) return;
      resetActiveAccount();
      refreshAccountDisplay(panel);
    });
    // Initial render of the account display
    refreshAccountDisplay(panel);

    // Upload + Apply
    const fileInput = panel.querySelector('.__rf_upload_file');
    panel.querySelector('.__rf_upload_btn').addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', async () => {
      const f = fileInput.files && fileInput.files[0];
      if (!f) return;
      fileInput.value = '';
      if (window.__rfAcctRunning) {
        panel.querySelector('.__rf_acct_status').textContent = 'Already running…';
        return;
      }
      await applyUpload(panel, f);
    });
  }

  // ---------- dialog watcher ----------
  function isAttachDialog(el) {
    if (!(el instanceof Element)) return false;
    if (el.getAttribute('role') !== 'dialog') return false;
    // Heading inside dialog
    const heading = el.querySelector('h1, h2, h3, h4, [role="heading"]');
    return !!heading && DIALOG_TITLE_RE.test(heading.textContent);
  }

  function scanForDialogs() {
    document.querySelectorAll('[role="dialog"]').forEach(d => {
      if (isAttachDialog(d)) attachFilterToDialog(d);
    });
  }

  // Light, document-level observer — ONLY listens for added/removed nodes,
  // NOT subtree attribute changes, so it's cheap.
  let bootTimer;
  const bootMo = new MutationObserver(() => {
    clearTimeout(bootTimer);
    bootTimer = setTimeout(scanForDialogs, 200);
  });
  bootMo.observe(document.body, { childList: true, subtree: true });

  // First pass in case the dialog is already open when the script loads.
  scanForDialogs();
  // Mount persistent account panel
  mountAccountPanel();
  // Re-mount on SPA-style URL changes
  let _lastPath = location.pathname;
  setInterval(() => {
    if (location.pathname !== _lastPath) {
      _lastPath = location.pathname;
      const ex = document.getElementById('__rf_acct_panel');
      if (!isExpensesPage() && ex) ex.remove();
      else if (isExpensesPage() && !ex) mountAccountPanel();
    }
  }, 1000);
})();
