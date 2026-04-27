// ==UserScript==
// @name         Coupa Receipt Filter (Attach Receipt dialog, ±% across currencies)
// @namespace    local.tylerkeller
// @version      0.4.5
// @description  Filter the Coupa "Attach a receipt" dialog to receipts within ±X% of the expense line's Total Amount (USD/EUR/COP/SGD/TRY). Adds a compact top-right panel with an "Apply Account to All" button + progress bar (skips lines that already have the configured account).
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
  const DEFAULT_ACCOUNT = {
    account_id: 6222,        // returned id from /accounts/select_dynamic_account
    account_type_id: 4,       // US1
    display_name: 'PHILADELPHIA-Finance Systems & Projects-NONE-Miscellaneous expenses',
    code: 'US010-26001-999-NONE-70919900',
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
        return !accounts.some(a => Number(a.account_id) === DEFAULT_ACCOUNT.account_id);
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

  // ---------- persistent top-right account panel ----------
  function isExpensesPage() {
    return /\/expense(?:s|_reports)/i.test(location.pathname);
  }

  function mountAccountPanel() {
    if (!isExpensesPage()) return;
    if (document.getElementById('__rf_acct_panel')) return;
    const panel = document.createElement('div');
    panel.id = '__rf_acct_panel';
    panel.style.cssText = [
      'position:fixed',
      'top:10px',
      'right:10px',
      'z-index:99999',
      'padding:8px 10px',
      'background:#fff',
      'border:1px solid #06c',
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
      <button class="__rf_apply_account_btn" type="button" style="background:#06c;color:#fff;border:1px solid #048;padding:5px 8px;border-radius:3px;cursor:pointer;width:100%;">Apply Account to All</button>
      <div class="__rf_progress_wrap" style="display:none;margin-top:6px;height:10px;background:#eef;border:1px solid #abc;border-radius:4px;overflow:hidden;">
        <div class="__rf_progress_bar" style="height:100%;width:0%;background:#06c;transition:width 200ms ease;"></div>
      </div>
      <div class="__rf_progress_text" style="display:none;margin-top:3px;font-size:10px;color:#06c;text-align:center;font-variant-numeric:tabular-nums;"></div>
      <div style="margin-top:6px;font-size:11px;line-height:1.3;word-break:break-word;">
        <b>Account:</b> ${DEFAULT_ACCOUNT.display_name}<br>
        <span style="color:#888;">id ${DEFAULT_ACCOUNT.account_id} &middot; ${DEFAULT_ACCOUNT.code}</span>
      </div>
      <div class="__rf_acct_status" style="margin-top:6px;font-size:11px;color:#06c;min-height:14px;word-break:break-word;"></div>
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
      if (!confirm(`Apply account ${DEFAULT_ACCOUNT.account_id} (${DEFAULT_ACCOUNT.display_name}) to every line in ${draftIds.length} draft reports?\n\nLines that already have this account will be skipped.\n\nThis is a bulk PATCH that may take several minutes.`)) {
        status.textContent = 'cancelled';
        return;
      }
      runApplyAccountToAll(draftIds, panel);
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
