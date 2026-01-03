// script.js – Farming CBA Decision Tool 2
// Control-centric results, Excel-first workflow, AI prompt builder

document.addEventListener('DOMContentLoaded', function () {
  'use strict';

  const TOOL_NAME = 'Farming CBA Decision Tool 2';

  // ---------- STATE ----------
  let treatments = [];
  let nextId = 1;
  let currentScenario = 'base';

  // ---------- BASIC HELPERS ----------
  function toNumber(value) {
    if (typeof value === 'number') return value;
    if (value === null || value === undefined) return 0;
    const v = parseFloat(String(value).replace(/,/g, '').trim());
    return isNaN(v) ? 0 : v;
  }

  function formatNumber(value, decimals) {
    const n = Number(value);
    if (!isFinite(n)) return '';
    return n.toLocaleString(undefined, {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals
    });
  }

  function getDiscountRate() {
    const el = document.getElementById('discount-rate');
    if (!el) return 0.0;
    const v = parseFloat(el.value);
    return isNaN(v) ? 0.0 : v / 100.0;
  }

  function getAnalysisYears() {
    const el = document.getElementById('analysis-years');
    if (!el) return 0;
    const v = parseInt(el.value, 10);
    return isNaN(v) || v <= 0 ? 0 : v;
  }

  function getCurrencyLabel() {
    const el = document.getElementById('currency-label');
    if (!el) return '$';
    return el.value || '$';
  }

  function pvFactor(rate, years) {
    // Sum_{t=1..years} [1 / (1+r)^t]
    if (years <= 0) return 0;
    if (rate <= -0.99) return years; // avoid division by zero
    let factor = 0;
    for (let t = 1; t <= years; t++) {
      factor += 1 / Math.pow(1 + rate, t);
    }
    return factor;
  }

  // ---------- SCENARIO ADJUSTMENT ----------
  function getScenarioAdjustedParams(t) {
    let price = toNumber(t.pricePerTonne);
    let annualCost = toNumber(t.annualCostPerHa);

    switch (currentScenario) {
      case 'lowPrice':
        price *= 0.8;
        break;
      case 'highCost':
        annualCost *= 1.2;
        break;
      case 'optimistic':
        price *= 1.2;
        annualCost *= 0.9;
        break;
      case 'base':
      default:
        break;
    }
    return { price, annualCost };
  }

  function getScenarioLabel(key) {
    switch (key) {
      case 'lowPrice':
        return 'Lower commodity prices (−20% price)';
      case 'highCost':
        return 'Higher costs (+20% annual costs)';
      case 'optimistic':
        return 'Optimistic scenario (+20% price, −10% annual costs)';
      case 'base':
      default:
        return 'Base case (values as entered)';
    }
  }

  // ---------- METRICS ----------
  function computeMetricsForTreatment(t) {
    const r = getDiscountRate();
    const years = getAnalysisYears();
    const pf = pvFactor(r, years);

    const areaHa = Math.max(toNumber(t.areaHa), 0);
    const { price, annualCost } = getScenarioAdjustedParams(t);

    const yieldPerHa = toNumber(t.yieldPerHa);
    const otherBenefitsPerHa = toNumber(t.otherAnnualBenefitsPerHa);
    const capitalCost = Math.max(toNumber(t.capitalCost), 0);

    const annualBenefitsPerHa = (yieldPerHa * price) + otherBenefitsPerHa;
    const pvBenefits = areaHa > 0 ? annualBenefitsPerHa * areaHa * pf : 0;
    const pvOperatingCosts = areaHa > 0 ? annualCost * areaHa * pf : 0;
    const pvCosts = capitalCost + pvOperatingCosts;

    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts > 0 ? (pvBenefits / pvCosts) : null;
    const roi = pvCosts > 0 ? (npv / pvCosts) : null;

    return {
      pvBenefits,
      pvCosts,
      npv,
      bcr,
      roi
    };
  }

  function recalcAllBaseMetrics() {
    treatments.forEach(function (t) {
      const m = computeMetricsForTreatment(t);
      t._metrics = m; // base scenario metrics can be overwritten when rendering
    });
  }

  // ---------- BASELINE / CONTROL ----------
  function findControlIndex(arr) {
    if (!arr.length) return -1;
    const idxFlag = arr.findIndex(function (t) {
      return !!t.isControl;
    });
    if (idxFlag >= 0) return idxFlag;
    const idxName = arr.findIndex(function (t) {
      const nm = (t.name || '').toLowerCase();
      return nm.includes('control') || nm.includes('current');
    });
    if (idxName >= 0) return idxName;
    return 0;
  }

  // ---------- SYNC UI <-> STATE ----------
  function syncFromTreatmentsTable() {
    const tbody = document.getElementById('treatments-tbody');
    if (!tbody) return;

    treatments = [];
    nextId = 1;

    const rows = tbody.querySelectorAll('tr');
    rows.forEach(function (tr) {
      const idAttr = tr.getAttribute('data-id');
      const id = idAttr ? parseInt(idAttr, 10) : nextId;

      const nameInput = tr.querySelector('.t-name');
      const isControlInput = tr.querySelector('.t-is-control');
      const areaInput = tr.querySelector('.t-area');
      const capitalInput = tr.querySelector('.t-capital');
      const annualCostInput = tr.querySelector('.t-annual-cost');
      const yieldInput = tr.querySelector('.t-yield');
      const priceInput = tr.querySelector('.t-price');
      const otherInput = tr.querySelector('.t-other-benefits');
      const notesInput = tr.querySelector('.t-notes');

      const name = nameInput ? (nameInput.value || ('Treatment ' + id)) : ('Treatment ' + id);

      treatments.push({
        id: id,
        name: name,
        isControl: isControlInput ? isControlInput.checked : false,
        areaHa: areaInput ? toNumber(areaInput.value) : 0,
        capitalCost: capitalInput ? toNumber(capitalInput.value) : 0,
        annualCostPerHa: annualCostInput ? toNumber(annualCostInput.value) : 0,
        yieldPerHa: yieldInput ? toNumber(yieldInput.value) : 0,
        pricePerTonne: priceInput ? toNumber(priceInput.value) : 0,
        otherAnnualBenefitsPerHa: otherInput ? toNumber(otherInput.value) : 0,
        notes: notesInput ? (notesInput.value || '') : ''
      });

      if (id >= nextId) nextId = id + 1;
    });

    recalcAllBaseMetrics();
  }

  function syncToTreatmentsTable() {
    const tbody = document.getElementById('treatments-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    treatments.forEach(function (t) {
      const tr = document.createElement('tr');
      tr.setAttribute('data-id', String(t.id));

      function makeInputTd(type, className, value, step) {
        const td = document.createElement('td');
        const input = document.createElement('input');
        input.type = type;
        input.className = className;
        if (type === 'number') {
          input.step = step || 'any';
        }
        if (type === 'checkbox') {
          input.checked = !!value;
        } else {
          input.value = value != null ? value : '';
        }
        td.appendChild(input);
        return td;
      }

      // Name
      const nameTd = makeInputTd('text', 't-name', t.name);
      tr.appendChild(nameTd);

      // Is control
      const isControlTd = makeInputTd('checkbox', 't-is-control', t.isControl);
      tr.appendChild(isControlTd);

      // Area
      const areaTd = makeInputTd('number', 't-area', t.areaHa, '0.01');
      tr.appendChild(areaTd);

      // Capital cost
      const capTd = makeInputTd('number', 't-capital', t.capitalCost, '0.01');
      tr.appendChild(capTd);

      // Annual cost
      const annTd = makeInputTd('number', 't-annual-cost', t.annualCostPerHa, '0.01');
      tr.appendChild(annTd);

      // Yield
      const yTd = makeInputTd('number', 't-yield', t.yieldPerHa, '0.01');
      tr.appendChild(yTd);

      // Price
      const pTd = makeInputTd('number', 't-price', t.pricePerTonne, '0.01');
      tr.appendChild(pTd);

      // Other annual benefits
      const oTd = makeInputTd('number', 't-other-benefits', t.otherAnnualBenefitsPerHa, '0.01');
      tr.appendChild(oTd);

      // Notes
      const notesTd = makeInputTd('text', 't-notes', t.notes);
      tr.appendChild(notesTd);

      // Remove button
      const removeTd = document.createElement('td');
      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'btn btn-secondary btn-small remove-treatment';
      removeBtn.textContent = 'Remove';
      removeTd.appendChild(removeBtn);
      tr.appendChild(removeTd);

      tbody.appendChild(tr);
    });
  }

  // ---------- RESULTS MATRIX ----------
  function renderResultsMatrix() {
    const head = document.getElementById('results-matrix-head');
    const body = document.getElementById('results-matrix-body');
    const summaryDiv = document.getElementById('results-summary');
    if (!head || !body) return;
    head.innerHTML = '';
    body.innerHTML = '';

    if (!treatments.length) {
      if (summaryDiv) {
        summaryDiv.textContent = 'Add treatments and click “Update results” to see the matrix.';
      }
      return;
    }

    const metricsList = treatments.map(function (t) {
      return {
        t: t,
        metrics: computeMetricsForTreatment(t)
      };
    });

    const controlIndex = findControlIndex(metricsList.map(function (x) { return x.t; }));
    const controlItem = metricsList[controlIndex];

    const sorted = metricsList.slice().sort(function (a, b) {
      return (b.metrics.npv || 0) - (a.metrics.npv || 0);
    });

    sorted.forEach(function (item, idx) {
      item.rank = idx + 1;
    });

    const scenarioText = getScenarioLabel(currentScenario);
    const positive = sorted.filter(function (x) { return x.metrics.npv > 0; }).length;
    const total = sorted.length;
    const best = sorted[0];

    if (summaryDiv) {
      summaryDiv.textContent =
        scenarioText + '. ' +
        positive + ' of ' + total + ' treatments have positive NPV. ' +
        'Best performer: "' + best.t.name + '" (NPV ' + getCurrencyLabel() + ' ' + formatNumber(best.metrics.npv, 0) + '). ' +
        'Control: "' + controlItem.t.name + '" (NPV ' + getCurrencyLabel() + ' ' + formatNumber(controlItem.metrics.npv, 0) + ').';
    }

    // Header row: indicator vs treatments
    const headerRow = document.createElement('tr');

    const firstTh = document.createElement('th');
    firstTh.textContent = 'Indicator';
    headerRow.appendChild(firstTh);

    sorted.forEach(function (item) {
      const th = document.createElement('th');
      th.textContent = 'Rank ' + item.rank + ': ' + item.t.name;
      if (item.t.id === controlItem.t.id) {
        th.classList.add('col-control');
      }
      headerRow.appendChild(th);
    });

    head.appendChild(headerRow);

    function addRow(label, tooltip, formatter) {
      const tr = document.createElement('tr');
      const labelTd = document.createElement('td');
      labelTd.innerHTML =
        label +
        (tooltip
          ? ' <span class="indicator-tooltip" data-tooltip="' + tooltip.replace(/"/g, '&quot;') + '">?</span>'
          : '');
      tr.appendChild(labelTd);

      sorted.forEach(function (item) {
        const td = document.createElement('td');
        if (item.t.id === controlItem.t.id) {
          td.classList.add('col-control');
        }
        td.innerHTML = formatter(item, controlItem);
        tr.appendChild(td);
      });

      body.appendChild(tr);
    }

    const currency = getCurrencyLabel();

    addRow(
      'PV benefits',
      'Present value of all expected benefits (for example yield × price) across the analysis period.',
      function (item) {
        return currency + ' ' + formatNumber(item.metrics.pvBenefits, 0);
      }
    );

    addRow(
      'PV costs',
      'Present value of capital and recurring costs for this treatment.',
      function (item) {
        return currency + ' ' + formatNumber(item.metrics.pvCosts, 0);
      }
    );

    addRow(
      'Net present value (NPV)',
      'NPV = PV benefits − PV costs. Positive values suggest a net economic gain relative to zero baseline.',
      function (item) {
        return currency + ' ' + formatNumber(item.metrics.npv, 0);
      }
    );

    addRow(
      'Δ NPV vs control',
      'Difference in NPV between this treatment and the control under the current scenario.',
      function (item, control) {
        const delta = item.metrics.npv - control.metrics.npv;
        return currency + ' ' + formatNumber(delta, 0);
      }
    );

    addRow(
      'Benefit–cost ratio (BCR)',
      'BCR = PV benefits ÷ PV costs. Values above 1 indicate that benefits exceed costs in present value terms.',
      function (item) {
        return item.metrics.bcr != null ? formatNumber(item.metrics.bcr, 2) : 'n/a';
      }
    );

    addRow(
      'Return on investment (ROI)',
      'ROI = NPV ÷ PV costs. This can be read as net gain per dollar of PV cost.',
      function (item) {
        return item.metrics.roi != null ? formatNumber(item.metrics.roi, 2) : 'n/a';
      }
    );

    addRow(
      'Rank (by NPV)',
      'Ranking of treatments based on NPV under the selected scenario. 1 = highest NPV.',
      function (item) {
        return String(item.rank);
      }
    );
  }

  // ---------- SNAPSHOTS ----------
  function renderSnapshots() {
    const container = document.getElementById('snapshots-container');
    const scenarioLabelEl = document.getElementById('snapshots-scenario-label');
    if (!container) return;
    container.innerHTML = '';

    if (scenarioLabelEl) {
      scenarioLabelEl.textContent = 'Currently viewing: ' + getScenarioLabel(currentScenario) + '.';
    }

    if (!treatments.length) return;

    const currency = getCurrencyLabel();
    const metricsList = treatments.map(function (t) {
      return {
        t: t,
        metrics: computeMetricsForTreatment(t)
      };
    });

    const sorted = metricsList.slice().sort(function (a, b) {
      return (b.metrics.npv || 0) - (a.metrics.npv || 0);
    });

    sorted.forEach(function (item, idx) {
      const card = document.createElement('div');
      card.className = 'snapshot-card';

      const title = document.createElement('h3');
      title.textContent = 'Rank ' + (idx + 1) + ': ' + item.t.name + (item.t.isControl ? ' (Control)' : '');
      card.appendChild(title);

      const p1 = document.createElement('p');
      p1.textContent = 'NPV: ' + currency + ' ' + formatNumber(item.metrics.npv, 0);
      card.appendChild(p1);

      const p2 = document.createElement('p');
      p2.textContent = 'PV benefits: ' + currency + ' ' + formatNumber(item.metrics.pvBenefits, 0) +
        ' | PV costs: ' + currency + ' ' + formatNumber(item.metrics.pvCosts, 0);
      card.appendChild(p2);

      const p3 = document.createElement('p');
      const bcrText = item.metrics.bcr != null ? formatNumber(item.metrics.bcr, 2) : 'n/a';
      const roiText = item.metrics.roi != null ? formatNumber(item.metrics.roi, 2) : 'n/a';
      p3.textContent = 'BCR: ' + bcrText + ' | ROI: ' + roiText;
      card.appendChild(p3);

      const p4 = document.createElement('p');
      p4.textContent = 'Area: ' + formatNumber(item.t.areaHa, 2) + ' ha';
      card.appendChild(p4);

      if (item.t.notes) {
        const p5 = document.createElement('p');
        p5.textContent = 'Notes: ' + item.t.notes;
        card.appendChild(p5);
      }

      container.appendChild(card);
    });
  }

  // ---------- AI PROMPT ----------
  function buildAIPrompt() {
    const textarea = document.getElementById('ai-prompt-text');
    if (!textarea) return;

    if (!treatments.length) {
      textarea.value =
        'No treatments are currently defined in Farming CBA Decision Tool 2. ' +
        'Please add treatments and calculate results, then refresh this AI prompt.';
      return;
    }

    const r = getDiscountRate();
    const years = getAnalysisYears();
    const currency = getCurrencyLabel();
    const scenario = currentScenario;

    const metricsList = treatments.map(function (t) {
      const m = computeMetricsForTreatment(t);
      return { t, m };
    });

    const controlIndex = findControlIndex(metricsList.map(function (x) { return x.t; }));
    const control = metricsList[controlIndex];

    const sorted = metricsList.slice().sort(function (a, b) {
      return (b.m.npv || 0) - (a.m.npv || 0);
    });

    // Build a JSON-style prompt (not strictly JSON to keep it copy-friendly)
    let prompt = '';
    prompt += 'You are an agricultural economist and farm adviser.\n';
    prompt += 'Your task is to interpret results from a farm cost–benefit analysis tool called "' + TOOL_NAME + '".\n';
    prompt += 'The tool compares multiple treatments against a control using present value (PV) metrics.\n\n';

    prompt += 'CONSTRAINTS:\n';
    prompt += '- Use plain English suitable for a farmer or farm decision maker.\n';
    prompt += '- Do NOT tell the user what to choose, and do NOT set hard thresholds.\n';
    prompt += '- Treat the tool as decision support, not decision making.\n';
    prompt += '- Explain what each indicator means (NPV, PV benefits, PV costs, BCR, ROI).\n';
    prompt += '- Focus on why treatments perform well or poorly and what drives these results.\n';
    prompt += '- For lower BCR or low-performing treatments, suggest realistic, practical ways to improve performance ';
    prompt += '(for example reducing costs, improving yields, achieving better prices, or changing agronomic practices). ';
    prompt += 'Frame these as options for reflection, not rules.\n';
    prompt += '- Always refer back to risk, uncertainty, and fit with the whole-farm plan.\n\n';

    prompt += 'DEFINITIONS (use these verbatim):\n';
    prompt += 'NPV = PV benefits − PV costs. Positive NPV indicates net economic gain relative to a zero baseline for that treatment scenario.\n';
    prompt += 'PV benefits and PV costs are discounted sums over time using the base discount rate.\n';
    prompt += 'BCR = PV benefits ÷ PV costs. Values above 1 imply benefits exceed costs in present value terms.\n';
    prompt += 'ROI = NPV ÷ PV costs. Interpretable as net gain per dollar of PV cost.\n\n';

    prompt += 'GLOBAL ASSUMPTIONS:\n';
    prompt += '- Discount rate (per year): ' + (r * 100).toFixed(2) + '%\n';
    prompt += '- Analysis horizon (years): ' + years + '\n';
    prompt += '- Currency: ' + currency + '\n';
    prompt += '- Scenario: ' + getScenarioLabel(scenario) + '\n\n';

    prompt += 'CONTROL TREATMENT:\n';
    prompt += '- Name: ' + control.t.name + '\n';
    prompt += '- IsControl flag: ' + (control.t.isControl ? 'true' : 'false') + '\n';
    prompt += '- Area (ha): ' + control.t.areaHa + '\n';
    prompt += '- NPV: ' + currency + ' ' + formatNumber(control.m.npv, 0) + '\n';
    prompt += '- PV benefits: ' + currency + ' ' + formatNumber(control.m.pvBenefits, 0) + '\n';
    prompt += '- PV costs: ' + currency + ' ' + formatNumber(control.m.pvCosts, 0) + '\n\n';

    prompt += 'ALL TREATMENTS (sorted by NPV, highest first):\n';
    sorted.forEach(function (item, idx) {
      const delta = item.m.npv - control.m.npv;
      prompt += '- Treatment ' + (idx + 1) + ':\n';
      prompt += '  Name: ' + item.t.name + (item.t.isControl ? ' (Control)' : '') + '\n';
      prompt += '  Area_ha: ' + item.t.areaHa + '\n';
      prompt += '  Yield_t_per_ha: ' + item.t.yieldPerHa + '\n';
      prompt += '  Price_per_tonne: ' + item.t.pricePerTonne + '\n';
      prompt += '  CapitalCost_year0: ' + item.t.capitalCost + '\n';
      prompt += '  AnnualCost_per_ha: ' + item.t.annualCostPerHa + '\n';
      prompt += '  OtherAnnualBenefits_per_ha: ' + item.t.otherAnnualBenefitsPerHa + '\n';
      prompt += '  Notes: ' + (item.t.notes || '') + '\n';
      prompt += '  Metrics_under_scenario:\n';
      prompt += '    PV_benefits: ' + item.m.pvBenefits.toFixed(2) + '\n';
      prompt += '    PV_costs: ' + item.m.pvCosts.toFixed(2) + '\n';
      prompt += '    NPV: ' + item.m.npv.toFixed(2) + '\n';
      prompt += '    Delta_NPV_vs_control: ' + delta.toFixed(2) + '\n';
      prompt += '    BCR: ' + (item.m.bcr != null ? item.m.bcr.toFixed(4) : 'null') + '\n';
      prompt += '    ROI: ' + (item.m.roi != null ? item.m.roi.toFixed(4) : 'null') + '\n';
    });
    prompt += '\n';

    prompt += 'YOUR OUTPUT:\n';
    prompt += '1. Provide a 2–3 page plain-language interpretation of these results.\n';
    prompt += '   - Begin with a short overview of the economic picture (how many treatments look attractive, and how they compare to the control).\n';
    prompt += '   - Explain clearly why the top treatments have stronger NPV, BCR, and ROI (for example higher yield, better prices, lower costs, or a combination).\n';
    prompt += '   - Explain why some treatments underperform and what drives that (high capital cost, high annual costs, weak yield response, low prices, etc.).\n';
    prompt += '   - For each major treatment or group of treatments, describe trade-offs (for example higher NPV but also higher upfront capital and risk).\n';
    prompt += '   - Explicitly state that results are sensitive to yields, prices, costs, discount rate, and analysis horizon.\n';
    prompt += '2. Provide a short section on learning and improvement:\n';
    prompt += '   - For treatments with low BCR or negative NPV, suggest practical ways the farmer could potentially improve outcomes, e.g. reducing costs per hectare, refining amendment rates, improving agronomy, or targeting better prices.\n';
    prompt += '   - Emphasise these are options for reflection and further analysis, not rules or thresholds.\n';
    prompt += '3. End with a reminder that this tool is decision support only, and that choices should consider risk, cash flow, labour, and alignment with the whole-farm strategy.\n';

    textarea.value = prompt;
  }

  // ---------- CSV IMPORT / EXPORT ----------
  function parseCsvLine(line) {
    const cells = [];
    let current = '';
    let insideQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        if (insideQuotes && line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          insideQuotes = !insideQuotes;
        }
      } else if (ch === ',' && !insideQuotes) {
        cells.push(current);
        current = '';
      } else {
        current += ch;
      }
    }
    cells.push(current);
    return cells;
  }

  function parseCsv(text) {
    const lines = text.split(/\r?\n/).filter(function (line) {
      return line.trim() !== '';
    });
    if (!lines.length) return [];
    return lines.map(parseCsvLine);
  }

  function importFromCsvText(text) {
    const rows = parseCsv(text);
    if (!rows.length || rows.length < 2) {
      alert('CSV appears to be empty or has no data rows.');
      return;
    }

    const header = rows[0].map(function (h) { return h.trim(); });
    const nameIdx = header.indexOf('TreatmentName');
    const isControlIdx = header.indexOf('IsControl');
    const areaIdx = header.indexOf('AreaHa');
    const yieldIdx = header.indexOf('Yield_t_per_ha');
    const priceIdx = header.indexOf('Price_per_tonne');
    const capIdx = header.indexOf('CapitalCost_year0');
    const annIdx = header.indexOf('AnnualCost_per_ha');
    const otherIdx = header.indexOf('OtherAnnualBenefits_per_ha');
    const notesIdx = header.indexOf('Notes');

    if (nameIdx === -1 || areaIdx === -1 || yieldIdx === -1 || priceIdx === -1 ||
      capIdx === -1 || annIdx === -1 || otherIdx === -1) {
      alert('CSV is missing one or more required headers.\n' +
        'Required: TreatmentName, IsControl, AreaHa, Yield_t_per_ha, Price_per_tonne, ' +
        'CapitalCost_year0, AnnualCost_per_ha, OtherAnnualBenefits_per_ha, Notes');
      return;
    }

    const imported = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;
      const name = row[nameIdx] != null ? String(row[nameIdx]).trim() : '';
      if (!name) continue;

      const isControlRaw = isControlIdx !== -1 && row[isControlIdx] != null ? String(row[isControlIdx]).trim().toLowerCase() : '';
      const isControl = ['yes', 'y', 'true', '1'].indexOf(isControlRaw) !== -1;

      imported.push({
        id: imported.length + 1,
        name: name,
        isControl: isControl,
        areaHa: areaIdx !== -1 ? toNumber(row[areaIdx]) : 0,
        yieldPerHa: yieldIdx !== -1 ? toNumber(row[yieldIdx]) : 0,
        pricePerTonne: priceIdx !== -1 ? toNumber(row[priceIdx]) : 0,
        capitalCost: capIdx !== -1 ? toNumber(row[capIdx]) : 0,
        annualCostPerHa: annIdx !== -1 ? toNumber(row[annIdx]) : 0,
        otherAnnualBenefitsPerHa: otherIdx !== -1 ? toNumber(row[otherIdx]) : 0,
        notes: notesIdx !== -1 && row[notesIdx] != null ? String(row[notesIdx]) : ''
      });
    }

    if (!imported.length) {
      alert('No usable rows found in CSV.');
      return;
    }

    treatments = imported;
    nextId = imported.length + 1;
    recalcAllBaseMetrics();
    syncToTreatmentsTable();
    updateAllOutputs();
  }

  function exportResultsMatrixCsv() {
    if (!treatments.length) {
      alert('Nothing to export. Please add treatments first.');
      return;
    }

    const metricsList = treatments.map(function (t) {
      return {
        t: t,
        metrics: computeMetricsForTreatment(t)
      };
    });
    const controlIndex = findControlIndex(metricsList.map(function (x) { return x.t; }));
    const control = metricsList[controlIndex];

    const sorted = metricsList.slice().sort(function (a, b) {
      return (b.metrics.npv || 0) - (a.metrics.npv || 0);
    });

    const header = ['Indicator'];
    sorted.forEach(function (item, idx) {
      header.push('Rank ' + (idx + 1) + ': ' + item.t.name);
    });

    const rows = [];
    rows.push(header.join(','));

    const currency = getCurrencyLabel();

    function addRowCsv(label, formatter) {
      const row = [label];
      sorted.forEach(function (item) {
        row.push(formatter(item, control));
      });
      rows.push(row.join(','));
    }

    addRowCsv('PV benefits (' + currency + ')', function (item) {
      return item.metrics.pvBenefits.toFixed(2);
    });

    addRowCsv('PV costs (' + currency + ')', function (item) {
      return item.metrics.pvCosts.toFixed(2);
    });

    addRowCsv('NPV (' + currency + ')', function (item) {
      return item.metrics.npv.toFixed(2);
    });

    addRowCsv('Delta NPV vs control (' + currency + ')', function (item, ctl) {
      return (item.metrics.npv - ctl.metrics.npv).toFixed(2);
    });

    addRowCsv('BCR', function (item) {
      return item.metrics.bcr != null ? item.metrics.bcr.toFixed(4) : '';
    });

    addRowCsv('ROI', function (item) {
      return item.metrics.roi != null ? item.metrics.roi.toFixed(4) : '';
    });

    addRowCsv('Rank (by NPV)', function (item) {
      return String(sorted.findIndex(function (x) { return x.t.id === item.t.id; }) + 1);
    });

    const blob = new Blob([rows.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = TOOL_NAME.replace(/\s+/g, '_').toLowerCase() + '_results_matrix.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function downloadTemplateCsv(blankOnly) {
    const header = [
      'TreatmentName',
      'IsControl',
      'AreaHa',
      'Yield_t_per_ha',
      'Price_per_tonne',
      'CapitalCost_year0',
      'AnnualCost_per_ha',
      'OtherAnnualBenefits_per_ha',
      'Notes'
    ];
    const rows = [];
    rows.push(header.join(','));

    if (!blankOnly && treatments.length) {
      treatments.forEach(function (t) {
        rows.push([
          '"' + String(t.name).replace(/"/g, '""') + '"',
          t.isControl ? 'yes' : 'no',
          t.areaHa,
          t.yieldPerHa,
          t.pricePerTonne,
          t.capitalCost,
          t.annualCostPerHa,
          t.otherAnnualBenefitsPerHa,
          '"' + String(t.notes || '').replace(/"/g, '""') + '"'
        ].join(','));
      });
    } else {
      // one blank example row
      rows.push('"Control",yes,100,0,0,0,0,0,"Example control / current practice row"');
      rows.push('"Treatment 1",no,100,0,0,0,0,0,"Example treatment row"');
    }

    const blob = new Blob([rows.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = TOOL_NAME.replace(/\s+/g, '_').toLowerCase() + '_template.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function downloadAIPromptTxt() {
    const textarea = document.getElementById('ai-prompt-text');
    if (!textarea || !textarea.value) {
      alert('AI prompt is empty. Please click “Refresh AI prompt” first.');
      return;
    }
    const blob = new Blob([textarea.value], { type: 'text/plain;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = TOOL_NAME.replace(/\s+/g, '_').toLowerCase() + '_ai_prompt.txt';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ---------- CORE ACTIONS ----------
  function addBlankTreatment() {
    treatments.push({
      id: nextId++,
      name: 'Treatment ' + (treatments.length + 1),
      isControl: treatments.length === 0, // first row default as control
      areaHa: 100,
      capitalCost: 0,
      annualCostPerHa: 0,
      yieldPerHa: 0,
      pricePerTonne: 0,
      otherAnnualBenefitsPerHa: 0,
      notes: ''
    });
    syncToTreatmentsTable();
  }

  function clearAllTreatments() {
    treatments = [];
    nextId = 1;
    syncToTreatmentsTable();
    renderResultsMatrix();
    renderSnapshots();
    buildAIPrompt();
  }

  function loadSmallSample() {
    treatments = [
      {
        id: 1,
        name: 'Control – current practice',
        isControl: true,
        areaHa: 100,
        capitalCost: 0,
        annualCostPerHa: 250,
        yieldPerHa: 2.5,
        pricePerTonne: 400,
        otherAnnualBenefitsPerHa: 0,
        notes: 'Baseline system using existing machinery and inputs.'
      },
      {
        id: 2,
        name: 'Deep OM (CP1)',
        isControl: false,
        areaHa: 100,
        capitalCost: 16500,
        annualCostPerHa: 260,
        yieldPerHa: 3.0,
        pricePerTonne: 400,
        otherAnnualBenefitsPerHa: 0,
        notes: 'Illustrative soil organic matter amendment based on Faba Beans dataset.'
      },
      {
        id: 3,
        name: 'Deep OM + Gypsum (CP2)',
        isControl: false,
        areaHa: 100,
        capitalCost: 24000,
        annualCostPerHa: 270,
        yieldPerHa: 3.1,
        pricePerTonne: 400,
        otherAnnualBenefitsPerHa: 0,
        notes: 'Illustrative deep OM and gypsum combination.'
      },
      {
        id: 4,
        name: 'Deep Carbon-coated mineral (CCM)',
        isControl: false,
        areaHa: 100,
        capitalCost: 3225,
        annualCostPerHa: 255,
        yieldPerHa: 2.9,
        pricePerTonne: 400,
        otherAnnualBenefitsPerHa: 0,
        notes: 'Illustrative carbon-coated mineral treatment.'
      }
    ];
    nextId = 5;
    recalcAllBaseMetrics();
    syncToTreatmentsTable();
    updateAllOutputs();
  }

  function updateAllOutputs() {
    syncFromTreatmentsTable();
    renderResultsMatrix();
    renderSnapshots();
    buildAIPrompt();
  }

  // ---------- TABS ----------
  function setupTabs() {
    const tabButtons = document.querySelectorAll('[data-tab]');
    const panels = document.querySelectorAll('.tab-panel');

    tabButtons.forEach(function (btn) {
      btn.addEventListener('click', function () {
        const target = btn.getAttribute('data-tab');
        if (!target) return;

        tabButtons.forEach(function (b) { b.classList.remove('active'); });
        panels.forEach(function (panel) {
          if (panel.id === 'tab-' + target) {
            panel.classList.add('active');
          } else {
            panel.classList.remove('active');
          }
        });

        btn.classList.add('active');
      });
    });
  }

  // ---------- EVENT HANDLERS ----------
  function setupEventHandlers() {
    const addBtn = document.getElementById('btn-add-treatment');
    if (addBtn) {
      addBtn.addEventListener('click', function () { addBlankTreatment(); });
    }

    const updateBtn = document.getElementById('btn-update-results');
    if (updateBtn) {
      updateBtn.addEventListener('click', function () { updateAllOutputs(); });
    }

    const clearBtn = document.getElementById('btn-clear-all');
    if (clearBtn) {
      clearBtn.addEventListener('click', function () {
        if (confirm('Clear all treatments and reset the tool?')) {
          clearAllTreatments();
        }
      });
    }

    const sampleBtn = document.getElementById('btn-load-sample');
    if (sampleBtn) {
      sampleBtn.addEventListener('click', function () {
        if (confirm('Replace current inputs with a small sample scenario?')) {
          loadSmallSample();
        }
      });
    }

    const printBtn = document.getElementById('btn-print');
    if (printBtn) {
      printBtn.addEventListener('click', function () {
        window.print();
      });
    }

    const scenarioSelect = document.getElementById('scenario-select');
    if (scenarioSelect) {
      scenarioSelect.addEventListener('change', function () {
        currentScenario = this.value || 'base';
        renderResultsMatrix();
        renderSnapshots();
        buildAIPrompt();
      });
    }

    const treatmentsTable = document.getElementById('treatments-tbody');
    if (treatmentsTable) {
      treatmentsTable.addEventListener('click', function (ev) {
        const target = ev.target;
        if (!(target instanceof HTMLElement)) return;
        if (target.classList.contains('remove-treatment')) {
          const tr = target.closest('tr');
          if (!tr) return;
          const idAttr = tr.getAttribute('data-id');
          const id = idAttr ? parseInt(idAttr, 10) : null;
          if (id != null) {
            treatments = treatments.filter(function (t) { return t.id !== id; });
            syncToTreatmentsTable();
            updateAllOutputs();
          }
        }
      });
    }

    const discountInput = document.getElementById('discount-rate');
    if (discountInput) {
      discountInput.addEventListener('input', function () { updateAllOutputs(); });
    }

    const yearsInput = document.getElementById('analysis-years');
    if (yearsInput) {
      yearsInput.addEventListener('input', function () { updateAllOutputs(); });
    }

    const currencyInput = document.getElementById('currency-label');
    if (currencyInput) {
      currencyInput.addEventListener('input', function () {
        renderResultsMatrix();
        renderSnapshots();
        buildAIPrompt();
      });
    }

    const fileInput = document.getElementById('file-import-csv');
    if (fileInput) {
      fileInput.addEventListener('change', function (ev) {
        const file = ev.target.files && ev.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function (e) {
          const text = e.target.result;
          importFromCsvText(text);
          fileInput.value = '';
        };
        reader.readAsText(file);
      });
    }

    const exportResultsBtn = document.getElementById('btn-export-results-csv');
    if (exportResultsBtn) {
      exportResultsBtn.addEventListener('click', function () { exportResultsMatrixCsv(); });
    }

    const downloadTemplateBtn = document.getElementById('btn-download-template');
    if (downloadTemplateBtn) {
      downloadTemplateBtn.addEventListener('click', function () {
        downloadTemplateCsv(true);
      });
    }

    const downloadCurrentBtn = document.getElementById('btn-download-current');
    if (downloadCurrentBtn) {
      downloadCurrentBtn.addEventListener('click', function () {
        downloadTemplateCsv(false);
      });
    }

    const refreshAIBtn = document.getElementById('btn-refresh-ai-prompt');
    if (refreshAIBtn) {
      refreshAIBtn.addEventListener('click', function () { buildAIPrompt(); });
    }

    const downloadAIBtn = document.getElementById('btn-download-ai-prompt');
    if (downloadAIBtn) {
      downloadAIBtn.addEventListener('click', function () { downloadAIPromptTxt(); });
    }
  }

  // ---------- INIT ----------
  function init() {
    setupTabs();
    setupEventHandlers();
    addBlankTreatment(); // start with one blank row (marked as control)
    renderResultsMatrix();
    renderSnapshots();
    buildAIPrompt();
  }

  init();
});
