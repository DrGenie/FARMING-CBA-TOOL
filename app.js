/* app.js */
/* Farming CBA Decision Tool 2 – app.js
   Plain JS single-file app: tabs, Excel upload (trial or CBA template), calculations, export.
   Requires SheetJS (XLSX) from CDN (see index.html). */

(() => {
  'use strict';

  // ---------- Utilities ----------
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  const clamp = (x, a, b) => Math.min(b, Math.max(a, x));

  function toNumber(x) {
    if (x === null || x === undefined) return NaN;
    if (typeof x === 'number') return x;
    if (typeof x === 'string') {
      const s = x.trim();
      if (!s) return NaN;
      // remove commas and $ signs
      const cleaned = s.replace(/[$,]/g, '');
      const n = Number(cleaned);
      return Number.isFinite(n) ? n : NaN;
    }
    return NaN;
  }

  function fmtNum(x, dp = 2) {
    if (!Number.isFinite(x)) return '—';
    return x.toLocaleString(undefined, { minimumFractionDigits: dp, maximumFractionDigits: dp });
  }
  function fmtInt(x) {
    if (!Number.isFinite(x)) return '—';
    return x.toLocaleString(undefined, { maximumFractionDigits: 0 });
  }
  function fmtPct(x, dp = 1) {
    if (!Number.isFinite(x)) return '—';
    return (100 * x).toLocaleString(undefined, { minimumFractionDigits: dp, maximumFractionDigits: dp }) + '%';
  }
  function fmtCur(x, dp = 0) {
    if (!Number.isFinite(x)) return '—';
    return x.toLocaleString(undefined, { style: 'currency', currency: 'AUD', minimumFractionDigits: dp, maximumFractionDigits: dp });
  }
  function safeDiv(a, b) {
    if (!Number.isFinite(a) || !Number.isFinite(b) || b === 0) return NaN;
    return a / b;
  }

  function pvOfStream(amounts, discountRate) {
    // amounts indexed from t=1..n
    let pv = 0;
    for (let t = 1; t <= amounts.length; t++) {
      pv += amounts[t - 1] / Math.pow(1 + discountRate, t);
    }
    return pv;
  }

  function uniqueHeaders(headers) {
    const counts = new Map();
    return headers.map((h) => {
      let key = (h ?? '').toString().trim();
      if (!key) key = 'Unnamed';
      const n = counts.get(key) ?? 0;
      counts.set(key, n + 1);
      return n === 0 ? key : `${key}_${n}`;
    });
  }

  function tsvFromTable(tableEl) {
    const rows = Array.from(tableEl.querySelectorAll('tr'));
    const lines = rows.map((tr) => {
      const cells = Array.from(tr.querySelectorAll('th,td'));
      return cells.map((c) => (c.innerText ?? '').replace(/\s+/g, ' ').trim()).join('\t');
    });
    return lines.join('\n');
  }

  async function copyText(text) {
    try {
      await navigator.clipboard.writeText(text);
      toast('Copied to clipboard.');
    } catch {
      // Fallback
      const ta = document.createElement('textarea');
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      document.body.removeChild(ta);
      toast('Copied to clipboard.');
    }
  }

  // ---------- Default dataset (derived from uploaded Excel: "Data for Lockhart-FA-031225 (1).xlsx") ----------
  // This lets the tool run even before any upload, and provides a reference format.
  const DEFAULT_DATASET = {
    sourceName: 'Data for Lockhart-FA-031225 (1).xlsx (embedded summary)',
    type: 'trial_summary',
    control: 'Control',
    treatments: [
      { name: 'Control', meanYield_t_ha: 6.194611, meanCost_per_ha: 694.6 },
      { name: 'Deep Carbon-coated mineral (CCM)', meanYield_t_ha: 6.063805, meanCost_per_ha: 4109.855 },
      { name: 'Deep Gypsum', meanYield_t_ha: 6.129699, meanCost_per_ha: 1384.855 },
      { name: 'Deep OM (CP1)', meanYield_t_ha: 5.316999, meanCost_per_ha: 17384.855 },
      { name: 'Deep OM (CP1) + Carbon-coated mineral (CCM)', meanYield_t_ha: 5.470641, meanCost_per_ha: 22109.855 },
      { name: 'Deep OM (CP1) + PAM', meanYield_t_ha: 7.226651, meanCost_per_ha: 884.855 },
      { name: 'Deep OM (CP1) + liq. Gypsum (CHT)', meanYield_t_ha: 6.287088, meanCost_per_ha: 17787.507 },
      { name: 'Deep OM + Gypsum (CP2)', meanYield_t_ha: 6.258005, meanCost_per_ha: 24884.855 },
      { name: 'Deep Ripping', meanYield_t_ha: 6.545337, meanCost_per_ha: 884.855 },
      { name: 'Deep liq. Gypsum (CHT)', meanYield_t_ha: 6.236619, meanCost_per_ha: 1234.855 },
      { name: 'Deep liq. NPKS', meanYield_t_ha: 6.460784, meanCost_per_ha: 884.855 },
      { name: 'Surface Silicon', meanYield_t_ha: 6.517313, meanCost_per_ha: 834.855 }
    ]
  };

  // ---------- State ----------
  const state = {
    dataset: DEFAULT_DATASET,
    assumptions: {
      discountRate: 0.05,
      horizonYears: 10,
      benefitYears: 10,
      benefitDecay: 1.0, // 1 = no decay
      pricePerTonne: 350, // AUD per tonne (editable)
      includeOperatingCostDiff: false,
      annualCostDiffPerHa: 0
    },
    results: null
  };

  // ---------- UI: Toast ----------
  let toastTimer = null;
  function toast(msg) {
    const el = $('#toast');
    if (!el) return;
    el.textContent = msg;
    el.classList.add('show');
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.remove('show'), 2400);
  }

  // ---------- Tabs ----------
  function activateTab(tabId) {
    $$('.tab-btn').forEach((b) => b.classList.toggle('active', b.dataset.tab === tabId));
    $$('.tab-panel').forEach((p) => p.classList.toggle('active', p.id === tabId));
  }

  // ---------- Parsing: Detect and import workbook ----------
  function requireXLSX() {
    if (typeof XLSX === 'undefined') {
      throw new Error('SheetJS (XLSX) library not loaded. Check your internet connection or the script tag in index.html.');
    }
  }

  function sheetToRows(ws) {
    // 2D array of rows
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
  }

  function findHeaderRow(rows, requiredHeaders) {
    const req = requiredHeaders.map((h) => h.toLowerCase());
    for (let r = 0; r < Math.min(rows.length, 50); r++) {
      const row = rows[r].map((c) => (c ?? '').toString().trim().toLowerCase());
      const hasAll = req.every((h) => row.includes(h));
      if (hasAll) return r;
    }
    return -1;
  }

  function parseTrialStyleWorkbook(wb) {
    // Expected: a sheet with columns including "Amendment" and "Yield t/ha"
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = sheetToRows(ws);

    const headerRowIndex = findHeaderRow(rows, ['Amendment', 'Yield t/ha']);
    if (headerRowIndex < 0) {
      throw new Error('Could not find a header row containing "Amendment" and "Yield t/ha".');
    }

    const headers = uniqueHeaders(rows[headerRowIndex]);
    const dataRows = rows.slice(headerRowIndex + 1).filter((r) => r.some((v) => (v ?? '').toString().trim() !== ''));

    const idx = (name) => headers.findIndex((h) => h.toLowerCase() === name.toLowerCase());
    const amendmentIdx = idx('Amendment');
    const yieldIdx = headers.findIndex((h) => h.toLowerCase().includes('yield'));
    // Prefer the special "|" column if present; else fall back to "Treatment Input Cost Only /Ha"
    const pipeIdx = headers.findIndex((h) => h.trim() === '|');
    const oneOffIdx = pipeIdx >= 0
      ? pipeIdx
      : headers.findIndex((h) => h.toLowerCase().includes('treatment input cost'));

    if (amendmentIdx < 0 || yieldIdx < 0) {
      throw new Error('Trial format detected, but required columns are missing.');
    }

    const rowsParsed = dataRows.map((r) => {
      const amendment = (r[amendmentIdx] ?? '').toString().trim();
      const y = toNumber(r[yieldIdx]);
      const oneOff = oneOffIdx >= 0 ? toNumber(r[oneOffIdx]) : NaN;
      return { amendment, yield_t_ha: y, oneOffCost_per_ha: oneOff };
    }).filter((x) => x.amendment);

    // group by amendment
    const by = new Map();
    for (const row of rowsParsed) {
      if (!by.has(row.amendment)) by.set(row.amendment, []);
      by.get(row.amendment).push(row);
    }

    const treatments = Array.from(by.entries()).map(([name, arr]) => {
      const ys = arr.map((a) => a.yield_t_ha).filter(Number.isFinite);
      const cs = arr.map((a) => a.oneOffCost_per_ha).filter(Number.isFinite);
      const mean = (xs) => xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : NaN;
      return {
        name,
        meanYield_t_ha: mean(ys),
        meanCost_per_ha: mean(cs)
      };
    });

    // Infer control
    let control = treatments.find((t) => t.name.toLowerCase() === 'control')?.name;
    if (!control) control = treatments[0]?.name ?? 'Control';

    // If the chosen cost column is missing/NaN for many rows, set meanCost to 0 for all and warn via UI
    const costMissing = treatments.filter((t) => !Number.isFinite(t.meanCost_per_ha)).length;
    if (costMissing > 0) {
      toast('Note: Some cost cells are missing in the uploaded file. Costs for those treatments may show as 0.');
      treatments.forEach((t) => {
        if (!Number.isFinite(t.meanCost_per_ha)) t.meanCost_per_ha = 0;
      });
    }

    return {
      sourceName: `${wb.Props?.Title ? wb.Props.Title + ' – ' : ''}${sheetName}`,
      type: 'trial_summary',
      control,
      treatments
    };
  }

  function parseCBATemplateWorkbook(wb) {
    // Expected: sheet with columns Treatment, Year, Benefit, Cost, optional IsControl
    const sheetName = wb.SheetNames.find((n) => n.toLowerCase().includes('cba')) ?? wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const rows = sheetToRows(ws);
    const headerRowIndex = findHeaderRow(rows, ['Treatment', 'Year', 'Benefit', 'Cost']);
    if (headerRowIndex < 0) {
      throw new Error('Could not find a header row containing "Treatment", "Year", "Benefit", and "Cost".');
    }
    const headers = uniqueHeaders(rows[headerRowIndex]);
    const dataRows = rows.slice(headerRowIndex + 1).filter((r) => r.some((v) => (v ?? '').toString().trim() !== ''));

    const col = (label) => headers.findIndex((h) => h.toLowerCase() === label.toLowerCase());

    const iTrt = col('Treatment');
    const iYear = col('Year');
    const iBen = col('Benefit');
    const iCost = col('Cost');
    const iCtrl = headers.findIndex((h) => h.toLowerCase() === 'iscontrol');

    const parsed = dataRows.map((r) => ({
      treatment: (r[iTrt] ?? '').toString().trim(),
      year: Math.trunc(toNumber(r[iYear])),
      benefit: toNumber(r[iBen]),
      cost: toNumber(r[iCost]),
      isControl: iCtrl >= 0 ? ((r[iCtrl] ?? '').toString().trim().toLowerCase() === 'true' || toNumber(r[iCtrl]) === 1) : false
    })).filter((x) => x.treatment && Number.isFinite(x.year));

    // group
    const by = new Map();
    for (const row of parsed) {
      if (!by.has(row.treatment)) by.set(row.treatment, []);
      by.get(row.treatment).push(row);
    }

    const treatments = Array.from(by.entries()).map(([name, arr]) => {
      const isControl = arr.some((a) => a.isControl) || name.toLowerCase() === 'control';
      return {
        name,
        isControl,
        cashflows: arr
          .filter((a) => a.year >= 0)
          .sort((a, b) => a.year - b.year)
          .map((a) => ({ year: a.year, benefit: a.benefit || 0, cost: a.cost || 0 }))
      };
    });

    let control = treatments.find((t) => t.isControl)?.name;
    if (!control) control = treatments.find((t) => t.name.toLowerCase() === 'control')?.name;
    if (!control) control = treatments[0]?.name ?? 'Control';

    return {
      sourceName: sheetName,
      type: 'cba_cashflows',
      control,
      treatments
    };
  }

  function detectAndParseWorkbook(wb) {
    // Try trial, then CBA template
    try {
      return parseTrialStyleWorkbook(wb);
    } catch (e1) {
      try {
        return parseCBATemplateWorkbook(wb);
      } catch (e2) {
        const msg = [
          'Could not recognise this Excel file.',
          '',
          'Tried:',
          '• Trial format: needs columns "Amendment" and "Yield t/ha".',
          '• CBA template: needs columns "Treatment", "Year", "Benefit", "Cost".',
          '',
          `Trial parse error: ${e1.message}`,
          `CBA template parse error: ${e2.message}`
        ].join('\n');
        throw new Error(msg);
      }
    }
  }

  // ---------- Calculations ----------
  function computeResults() {
    const ds = state.dataset;
    const A = state.assumptions;

    const dr = clamp(A.discountRate, 0, 0.5);
    const horizon = Math.max(1, Math.trunc(A.horizonYears));
    const benefitYears = clamp(Math.trunc(A.benefitYears), 0, horizon);
    const decay = clamp(A.benefitDecay, 0, 1.0);
    const price = Math.max(0, A.pricePerTonne);

    if (!ds || !ds.treatments || ds.treatments.length === 0) {
      state.results = null;
      return;
    }

    if (ds.type === 'trial_summary') {
      const control = ds.treatments.find((t) => t.name === ds.control) ?? ds.treatments[0];
      const y0 = control?.meanYield_t_ha ?? 0;
      const c0 = control?.meanCost_per_ha ?? 0;

      const rows = ds.treatments.map((t) => {
        const dy = (t.meanYield_t_ha ?? 0) - y0;
        const incCost = (t.meanCost_per_ha ?? 0) - c0;

        const annualBenefits = [];
        for (let yr = 1; yr <= benefitYears; yr++) {
          annualBenefits.push(dy * price * Math.pow(decay, yr - 1));
        }

        const pvBenefits = pvOfStream(annualBenefits, dr);

        // Costs: treat incCost as year-0 (one-off)
        // Optional: add annual operating cost difference (same each year) if user enables it
        const annualCostDiff = A.includeOperatingCostDiff ? (A.annualCostDiffPerHa || 0) : 0;
        const annualCosts = [];
        for (let yr = 1; yr <= horizon; yr++) annualCosts.push(annualCostDiff);
        const pvAnnualCosts = pvOfStream(annualCosts, dr);

        const pvCosts = incCost + pvAnnualCosts;
        const npv = pvBenefits - pvCosts;
        const bcr = safeDiv(pvBenefits, pvCosts);
        const roi = safeDiv(npv, pvCosts);

        return {
          name: t.name,
          meanYield_t_ha: t.meanYield_t_ha,
          deltaYield_t_ha: dy,
          oneOffCost_per_ha: incCost,
          pvBenefits,
          pvCosts,
          npv,
          bcr,
          roi
        };
      });

      // Ranking: exclude control
      const nonControl = rows.filter((r) => r.name !== ds.control);
      nonControl.sort((a, b) => (b.npv ?? -Infinity) - (a.npv ?? -Infinity));
      const rankMap = new Map(nonControl.map((r, i) => [r.name, i + 1]));
      rows.forEach((r) => { r.rank = (r.name === ds.control) ? '—' : (rankMap.get(r.name) ?? '—'); });

      state.results = {
        kind: 'trial_summary',
        control: ds.control,
        rows
      };
      return;
    }

    if (ds.type === 'cba_cashflows') {
      const rows = ds.treatments.map((t) => {
        const cash = new Map();
        for (const cf of t.cashflows) {
          const yr = Math.trunc(cf.year);
          cash.set(yr, {
            benefit: Number.isFinite(cf.benefit) ? cf.benefit : 0,
            cost: Number.isFinite(cf.cost) ? cf.cost : 0
          });
        }

        const annualBenefits = [];
        const annualCosts = [];
        for (let yr = 1; yr <= horizon; yr++) {
          const v = cash.get(yr) ?? { benefit: 0, cost: 0 };
          annualBenefits.push(v.benefit);
          annualCosts.push(v.cost);
        }

        const pvBenefits = pvOfStream(annualBenefits, dr);
        const pvCosts = pvOfStream(annualCosts, dr) + (cash.get(0)?.cost ?? 0); // allow year-0 cost
        const npv = pvBenefits - pvCosts;
        const bcr = safeDiv(pvBenefits, pvCosts);
        const roi = safeDiv(npv, pvCosts);

        return { name: t.name, pvBenefits, pvCosts, npv, bcr, roi };
      });

      const nonControl = rows.filter((r) => r.name !== ds.control);
      nonControl.sort((a, b) => (b.npv ?? -Infinity) - (a.npv ?? -Infinity));
      const rankMap = new Map(nonControl.map((r, i) => [r.name, i + 1]));
      rows.forEach((r) => { r.rank = (r.name === ds.control) ? '—' : (rankMap.get(r.name) ?? '—'); });

      state.results = { kind: 'cba_cashflows', control: ds.control, rows };
      return;
    }

    state.results = null;
  }

  // ---------- Rendering ----------
  function renderDatasetSummary() {
    const ds = state.dataset;
    $('#dataSource').textContent = ds?.sourceName ?? '—';
    $('#dataFormat').textContent = ds?.type ?? '—';
    $('#controlName').textContent = ds?.control ?? '—';

    // Treatment list
    const list = $('#treatmentList');
    list.innerHTML = '';
    if (!ds?.treatments?.length) {
      list.innerHTML = '<div class="muted">No treatments loaded.</div>';
      return;
    }

    const table = document.createElement('table');
    table.className = 'table';
    const thead = document.createElement('thead');
    const trh = document.createElement('tr');

    if (ds.type === 'trial_summary') {
      trh.innerHTML = `
        <th>Treatment</th>
        <th class="num">Mean yield (t/ha)</th>
        <th class="num">Mean cost (per ha)</th>
      `;
      thead.appendChild(trh);
      table.appendChild(thead);

      const tbody = document.createElement('tbody');
      for (const t of ds.treatments) {
        const tr = document.createElement('tr');
        const cost = Number.isFinite(t.meanCost_per_ha) ? fmtCur(t.meanCost_per_ha, 0) : '—';
        tr.innerHTML = `
          <td>${escapeHtml(t.name)}</td>
          <td class="num">${fmtNum(t.meanYield_t_ha, 3)}</td>
          <td class="num">${cost}</td>
        `;
        if (t.name === ds.control) tr.classList.add('row-accent');
        tbody.appendChild(tr);
      }
      table.appendChild(tbody);
      list.appendChild(table);
      return;
    }

    if (ds.type === 'cba_cashflows') {
      trh.innerHTML = `
        <th>Treatment</th>
        <th class="num">Years provided</th>
        <th class="num">Has year-0 cost</th>
      `;
      thead.appendChild(trh);
      table.appendChild(thead);

      const tbody = document.createElement('tbody');
      for (const t of ds.treatments) {
        const years = (t.cashflows ?? []).map((c) => c.year).filter((y) => Number.isFinite(y));
        const uniqYears = Array.from(new Set(years)).sort((a, b) => a - b);
        const hasY0 = (t.cashflows ?? []).some((c) => Math.trunc(c.year) === 0 && (toNumber(c.cost) || 0) !== 0);

        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${escapeHtml(t.name)}</td>
          <td class="num">${fmtInt(uniqYears.length)}</td>
          <td class="num">${hasY0 ? 'Yes' : 'No'}</td>
        `;
        if (t.name === ds.control) tr.classList.add('row-accent');
        tbody.appendChild(tr);
      }
      table.appendChild(tbody);
      list.appendChild(table);
      return;
    }

    list.innerHTML = '<div class="muted">Loaded, but preview is not available for this data format.</div>';
  }

  function renderAssumptions() {
    $('#discountRate').value = (state.assumptions.discountRate * 100).toString();
    $('#horizonYears').value = state.assumptions.horizonYears.toString();
    $('#benefitYears').value = state.assumptions.benefitYears.toString();
    $('#benefitDecay').value = (state.assumptions.benefitDecay * 100).toString();
    $('#pricePerTonne').value = state.assumptions.pricePerTonne.toString();
    $('#includeOpCost').checked = !!state.assumptions.includeOperatingCostDiff;
    $('#annualOpCost').value = state.assumptions.annualCostDiffPerHa.toString();
    $('#annualOpCost').disabled = !state.assumptions.includeOperatingCostDiff;
  }

  function renderResults() {
    const res = state.results;
    const box = $('#resultsBox');
    box.innerHTML = '';

    if (!res?.rows?.length) {
      box.innerHTML = '<div class="muted">Load data and click “Run analysis”.</div>';
      $('#btnCopyResults').disabled = true;
      $('#btnExportResults').disabled = true;
      return;
    }

    // Vertical comparison table: indicators as rows, treatments as columns (incl control)
    const treatments = res.rows.map((r) => r.name);
    const indicators = [
      { key: 'pvBenefits', label: 'Present value of benefits (PV Benefits)' , fmt: (v) => fmtCur(v, 0) },
      { key: 'pvCosts', label: 'Present value of costs (PV Costs)', fmt: (v) => fmtCur(v, 0) },
      { key: 'npv', label: 'Net present value (NPV = PV Benefits − PV Costs)', fmt: (v) => fmtCur(v, 0) },
      { key: 'bcr', label: 'Benefit–cost ratio (BCR = PV Benefits ÷ PV Costs)', fmt: (v) => Number.isFinite(v) ? fmtNum(v, 2) : '—' },
      { key: 'roi', label: 'Return on investment (ROI = NPV ÷ PV Costs)', fmt: (v) => Number.isFinite(v) ? fmtNum(v, 2) : '—' },
      { key: 'rank', label: 'Ranking (1 = highest NPV)', fmt: (v) => (v ?? '—').toString() }
    ];

    const tableWrap = document.createElement('div');
    tableWrap.className = 'table-wrap';

    const table = document.createElement('table');
    table.className = 'table sticky';

    const thead = document.createElement('thead');
    const trh = document.createElement('tr');
    trh.innerHTML = `<th>Indicator</th>` + treatments.map((t) => `<th class="num">${escapeHtml(t)}</th>`).join('');
    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    for (const ind of indicators) {
      const tr = document.createElement('tr');
      tr.innerHTML = `<th>${escapeHtml(ind.label)}</th>` + res.rows.map((r) => {
        const val = r[ind.key];
        const isControl = r.name === res.control;
        const cls = ['num', isControl ? 'cell-accent' : ''].join(' ').trim();
        return `<td class="${cls}">${ind.fmt(val)}</td>`;
      }).join('');
      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    tableWrap.appendChild(table);

    box.appendChild(tableWrap);

    // Quick notes
    const note = document.createElement('div');
    note.className = 'note';
    const c = res.rows.find((x) => x.name === res.control);
    note.innerHTML = `
      <div><strong>Control:</strong> ${escapeHtml(res.control)}.</div>
      ${c && Number.isFinite(c.meanYield_t_ha) ? `<div><strong>Control mean yield:</strong> ${fmtNum(c.meanYield_t_ha, 3)} t/ha.</div>` : ''}
      <div class="muted">Assumptions: discount rate ${fmtPct(state.assumptions.discountRate)}, horizon ${fmtInt(state.assumptions.horizonYears)} years, benefit duration ${fmtInt(state.assumptions.benefitYears)} years, benefit decay ${fmtPct(1 - state.assumptions.benefitDecay)}, price ${fmtCur(state.assumptions.pricePerTonne, 0)}/t.</div>
    `;
    box.appendChild(note);

    // Enable copy/export
    $('#btnCopyResults').disabled = false;
    $('#btnExportResults').disabled = false;
  }

  function renderSensitivity() {
    const res = state.results;
    const out = $('#sensitivityBox');
    out.innerHTML = '';

    if (!res?.rows?.length) {
      out.innerHTML = '<div class="muted">Run an analysis first.</div>';
      return;
    }

    // simple sensitivity: vary price +/- 20% and show top 5 NPVs
    const basePrice = state.assumptions.pricePerTonne;
    const dr = state.assumptions.discountRate;
    const prices = [0.8, 1.0, 1.2].map((m) => m * basePrice);

    const makeNPVAtPrice = (price) => {
      const A0 = { ...state.assumptions, pricePerTonne: price };
      const ds = state.dataset;
      if (ds.type !== 'trial_summary') return [];
      const control = ds.treatments.find((t) => t.name === ds.control) ?? ds.treatments[0];
      const y0 = control?.meanYield_t_ha ?? 0;
      const c0 = control?.meanCost_per_ha ?? 0;
      const benefitYears = clamp(Math.trunc(A0.benefitYears), 0, Math.max(1, Math.trunc(A0.horizonYears)));
      const decay = clamp(A0.benefitDecay, 0, 1.0);
      const annualCostDiff = A0.includeOperatingCostDiff ? (A0.annualCostDiffPerHa || 0) : 0;
      const horizon = Math.max(1, Math.trunc(A0.horizonYears));

      const rows = ds.treatments.map((t) => {
        const dy = (t.meanYield_t_ha ?? 0) - y0;
        const incCost = (t.meanCost_per_ha ?? 0) - c0;
        const annualBenefits = Array.from({ length: benefitYears }, (_, i) => dy * price * Math.pow(decay, i));
        const pvBenefits = pvOfStream(annualBenefits, dr);
        const pvCosts = incCost + pvOfStream(Array.from({ length: horizon }, () => annualCostDiff), dr);
        return { name: t.name, npv: pvBenefits - pvCosts };
      }).filter((r) => r.name !== ds.control);

      rows.sort((a, b) => b.npv - a.npv);
      return rows.slice(0, 5);
    };

    const card = document.createElement('div');
    card.className = 'card';
    card.innerHTML = `<h3>Price sensitivity (top 5 by NPV)</h3>`;
    const grid = document.createElement('div');
    grid.className = 'grid3';

    for (const p of prices) {
      const top = makeNPVAtPrice(p);
      const col = document.createElement('div');
      col.className = 'mini';
      col.innerHTML = `<div class="mini-title">${fmtCur(p, 0)}/t</div>`;
      if (!top.length) {
        col.innerHTML += `<div class="muted">Not available for this data format.</div>`;
      } else {
        const ul = document.createElement('ol');
        ul.className = 'mini-list';
        for (const r of top) {
          const li = document.createElement('li');
          li.innerHTML = `<span>${escapeHtml(r.name)}</span><span class="num">${fmtCur(r.npv, 0)}</span>`;
          ul.appendChild(li);
        }
        col.appendChild(ul);
      }
      grid.appendChild(col);
    }

    card.appendChild(grid);
    out.appendChild(card);
  }

  function escapeHtml(s) {
    return (s ?? '').toString()
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#039;');
  }

  // ---------- Export ----------
  function buildResultsWorkbook() {
    requireXLSX();
    if (!state.results?.rows?.length) throw new Error('No results to export.');

    const wb = XLSX.utils.book_new();

    // Assumptions sheet
    const A = state.assumptions;
    const assumptionsAOA = [
      ['Parameter', 'Value'],
      ['Data source', state.dataset?.sourceName ?? ''],
      ['Control', state.dataset?.control ?? ''],
      ['Discount rate', A.discountRate],
      ['Horizon (years)', A.horizonYears],
      ['Benefit duration (years)', A.benefitYears],
      ['Benefit decay factor (per year)', A.benefitDecay],
      ['Price ($/t)', A.pricePerTonne],
      ['Include annual operating cost difference', A.includeOperatingCostDiff ? 'Yes' : 'No'],
      ['Annual operating cost difference ($/ha/year)', A.annualCostDiffPerHa]
    ];
    const wsA = XLSX.utils.aoa_to_sheet(assumptionsAOA);
    XLSX.utils.book_append_sheet(wb, wsA, 'Assumptions');

    // Results sheet
    const rows = state.results.rows.map((r) => ({
      Treatment: r.name,
      IsControl: r.name === state.results.control,
      MeanYield_t_ha: r.meanYield_t_ha ?? '',
      DeltaYield_t_ha: r.deltaYield_t_ha ?? '',
      OneOffCost_per_ha: r.oneOffCost_per_ha ?? '',
      PV_Benefits: r.pvBenefits,
      PV_Costs: r.pvCosts,
      NPV: r.npv,
      BCR: r.bcr,
      ROI: r.roi,
      Rank: r.rank
    }));
    const wsT = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, wsT, 'Results');

    return wb;
  }

  function downloadWorkbook(wb, filename) {
    requireXLSX();
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
  }

  function downloadTemplate() {
    requireXLSX();
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ['Treatment', 'IsControl', 'Year', 'Benefit', 'Cost'],
      ['Control', 'TRUE', 0, '', 0],
      ['Control', 'TRUE', 1, 0, 0],
      ['Treatment A', 'FALSE', 0, '', 1000],
      ['Treatment A', 'FALSE', 1, 500, 50],
      ['Treatment A', 'FALSE', 2, 500, 50]
    ]);
    XLSX.utils.book_append_sheet(wb, ws, 'CBA_Input');
    downloadWorkbook(wb, 'FarmingCBA_Template.xlsx');
  }

  // ---------- Event handlers ----------
  function readAssumptionsFromUI() {
    const dr = toNumber($('#discountRate').value) / 100;
    const horizon = Math.trunc(toNumber($('#horizonYears').value));
    const benYears = Math.trunc(toNumber($('#benefitYears').value));
    const decay = toNumber($('#benefitDecay').value) / 100;
    const price = toNumber($('#pricePerTonne').value);
    const includeOp = $('#includeOpCost').checked;
    const opCost = toNumber($('#annualOpCost').value);

    state.assumptions.discountRate = Number.isFinite(dr) ? clamp(dr, 0, 0.5) : state.assumptions.discountRate;
    state.assumptions.horizonYears = Number.isFinite(horizon) ? clamp(horizon, 1, 50) : state.assumptions.horizonYears;
    state.assumptions.benefitYears = Number.isFinite(benYears) ? clamp(benYears, 0, state.assumptions.horizonYears) : state.assumptions.benefitYears;
    state.assumptions.benefitDecay = Number.isFinite(decay) ? clamp(decay, 0, 1) : state.assumptions.benefitDecay;
    state.assumptions.pricePerTonne = Number.isFinite(price) ? clamp(price, 0, 5000) : state.assumptions.pricePerTonne;
    state.assumptions.includeOperatingCostDiff = !!includeOp;
    state.assumptions.annualCostDiffPerHa = Number.isFinite(opCost) ? clamp(opCost, -100000, 100000) : state.assumptions.annualCostDiffPerHa;
  }

  async function handleFileUpload(file) {
    requireXLSX();
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const parsed = detectAndParseWorkbook(wb);
    state.dataset = parsed;

    renderDatasetSummary();
    toast(`Loaded: ${file.name}`);
  }

  function wireUI() {
    // Tabs
    $$('.tab-btn').forEach((btn) => {
      btn.addEventListener('click', () => activateTab(btn.dataset.tab));
    });

    // File upload
    $('#fileInput').addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        $('#uploadStatus').textContent = 'Reading…';
        await handleFileUpload(file);
        $('#uploadStatus').textContent = 'Loaded';
        computeResults();
        renderResults();
        renderSensitivity();
      } catch (err) {
        $('#uploadStatus').textContent = 'Error';
        console.error(err);
        toast(err.message || 'Failed to read file.');
        $('#errorText').textContent = err.message || String(err);
        $('#errorBox').classList.add('show');
        activateTab('tab-data');
      }
    });

    // Download template
    $('#btnTemplate').addEventListener('click', () => {
      try {
        downloadTemplate();
      } catch (err) {
        toast(err.message || 'Template download failed.');
      }
    });

    // Assumptions toggles
    $('#includeOpCost').addEventListener('change', () => {
      $('#annualOpCost').disabled = !$('#includeOpCost').checked;
    });

    // Run analysis
    $('#btnRun').addEventListener('click', () => {
      $('#errorBox').classList.remove('show');
      readAssumptionsFromUI();
      computeResults();
      renderResults();
      renderSensitivity();
      activateTab('tab-results');
      toast('Analysis updated.');
    });

    // Copy results
    $('#btnCopyResults').addEventListener('click', async () => {
      const table = $('#resultsBox table');
      if (!table) return toast('No results table found.');
      await copyText(tsvFromTable(table));
    });

    // Export results
    $('#btnExportResults').addEventListener('click', () => {
      try {
        const wb = buildResultsWorkbook();
        downloadWorkbook(wb, 'FarmingCBA_Results.xlsx');
        toast('Exported Excel file.');
      } catch (err) {
        toast(err.message || 'Export failed.');
      }
    });

    // Reset to embedded sample
    $('#btnResetSample').addEventListener('click', () => {
      state.dataset = DEFAULT_DATASET;
      renderDatasetSummary();
      computeResults();
      renderResults();
      renderSensitivity();
      toast('Loaded embedded sample summary.');
      activateTab('tab-data');
    });

    // Close error box
    $('#btnDismissError').addEventListener('click', () => {
      $('#errorBox').classList.remove('show');
    });
  }

  function init() {
    wireUI();
    renderDatasetSummary();
    renderAssumptions();
    computeResults();
    renderResults();
    renderSensitivity();
    activateTab('tab-overview');
  }

  document.addEventListener('DOMContentLoaded', init);
})();
