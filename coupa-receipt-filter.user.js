// ==UserScript==
// @name         Coupa Receipt Filter (Attach Receipt dialog, ±% across currencies)
// @namespace    local.tylerkeller
// @version      0.3.1
// @description  When the Coupa "Attach a receipt" dialog opens, filter receipts to those within ±X% of the expense line's Total Amount (converted to USD/EUR/COP/SGD/TRY). Displays merchant, date, and total on the filter bar. Shrinks the dialog body so the Attach/Cancel buttons remain visible.
// @match        https://*.coupahost.com/*
// @run-at       document-idle
// @grant        GM_xmlhttpRequest
// @connect      open.er-api.com
// ==/UserScript==

(function () {
  'use strict';

  // ---------- config ----------
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
