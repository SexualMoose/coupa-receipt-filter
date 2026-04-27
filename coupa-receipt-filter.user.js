// ==UserScript==
// @name         Coupa Receipt Filter (Attach Receipt dialog, ±% across currencies)
// @namespace    local.tylerkeller
// @version      0.4.0
// @description  Filter the Coupa "Attach a receipt" dialog to receipts within ±X% of the expense line's Total Amount (USD/EUR/COP/SGD/TRY). Also adds an "Apply Account to All" button that PATCHes every draft expense line with a configured account.
// @match        https://*.coupahost.com/*
// @run-at       document-idle
// @grant        GM_xmlhttpRequest
// @connect      open.er-api.com
// ==/UserScript==

(function () {
  'use strict';

  // ---------- config ----------
  // Default account to bulk-apply via the toolbar button.
  // Capture these values from your tenant by clicking "Choose" on a representative
  // line and watching POST /accounts/select_dynamic_account in DevTools Network.
  const DEFAULT_ACCOUNT = {
    account_id: 6222,        // returned id from /accounts/select_dynamic_account
    account_type_id: 4,       // US1
  };

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
        <span style="border-left:1px solid #c69; margin:0 4px; height:18px;"></span>
        <button class="__rf_apply_account" type="button" title="PATCH every line in every draft report with the configured account" style="background:#06c;color:#fff;border:1px solid #048;padding:3px 8px;border-radius:3px;">Apply Account to All</button>
      </div>
      <div class="__rf_meta" style="margin-top:6px;color:#222;"></div>
      <div class="__rf_targets" style="margin-top:4px;color:#444;"></div>
      <div class="__rf_status" style="margin-top:4px;color:#555;"></div>
      <div class="__rf_acct_status" style="margin-top:4px;color:#06c;"></div>
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
    $('.__rf_apply_account').addEventListener('click', async () => {
      if (window.__rfAcctRunning) {
        $('.__rf_acct_status').textContent = 'Already running…';
        return;
      }
      const draftReportIds = (window.ExpenseReports || [])
        .filter(r => r.status === 'draft')
        .map(r => r.id);
      if (!draftReportIds.length) {
        $('.__rf_acct_status').textContent = 'No draft reports found in this page.';
        return;
      }
      if (!confirm(`Apply account ${DEFAULT_ACCOUNT.account_id} to every line in ${draftReportIds.length} draft reports?\n\nThis is a bulk PATCH that may take several minutes.`)) {
        return;
      }
      runApplyAccountToAll(draftReportIds, $('.__rf_acct_status'));
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
  function fetchReportLines(reportId) {
    return fetch(`/expense_reports/${reportId}/edit`, { credentials: 'include' })
      .then(r => r.text())
      .then(html => {
        const start = html.search(/var\s+ExpenseReports\s*=\s*\[/);
        if (start < 0) return null;
        const arrStart = html.indexOf('[', start);
        let depth = 0, i = arrStart, inStr = false, strCh = '', esc = false;
        for (; i < html.length; i++) {
          const c = html[i];
          if (inStr) { if (esc) esc = false; else if (c === '\\') esc = true; else if (c === strCh) inStr = false; }
          else { if (c === '"' || c === '\'') { inStr = true; strCh = c; } else if (c === '[') depth++; else if (c === ']') { depth--; if (depth === 0) { i++; break; } } }
        }
        return JSON.parse(html.slice(arrStart, i)).find(p => p.id === reportId);
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
    set('expense_line[account_id]', DEFAULT_ACCOUNT.account_id);
    set('expense_line[account_type_id]', DEFAULT_ACCOUNT.account_type_id);
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

  async function runApplyAccountToAll(reportIds, statusEl) {
    window.__rfAcctRunning = true;
    const csrf = document.querySelector('meta[name="csrf-token"]')?.getAttribute('content');
    const headers = {
      'Accept': 'application/json, text/plain, */*',
      'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
      'X-CSRF-Token': csrf,
      'X-Requested-With': 'XMLHttpRequest',
    };
    let ok = 0, fail = 0, skipped = 0, total = 0;
    const failures = [];
    statusEl.textContent = 'Fetching reports…';

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
    total = allLines.length;
    statusEl.textContent = `Patching 0/${total}…`;

    for (const line of allLines) {
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
      if (done % 5 === 0 || done === total) {
        statusEl.textContent = `${done}/${total} ok=${ok} fail=${fail}`;
      }
      await new Promise(r => setTimeout(r, 120));
    }
    window.__rfAcctRunning = false;
    statusEl.innerHTML = `<b>Account apply complete.</b> ok=${ok} / fail=${fail}` +
      (failures.length ? ` <a href="data:application/json;base64,${btoa(JSON.stringify(failures, null, 2))}" download="account-apply-failures.json" style="color:#06c;">download failures</a>` : '');
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
})();
