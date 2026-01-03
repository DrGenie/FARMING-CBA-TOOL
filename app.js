// script.js – Farming CBA Decision Tool 2 (advanced analysis, scenarios, CSV import)
document.addEventListener('DOMContentLoaded', function () {
  'use strict';

  // --------- STATE ---------
  let treatments = [];
  let nextId = 1;
  let currentScenario = 'base';

  // --------- BASIC HELPERS ---------
  function toNumber(value) {
    if (typeof value === 'number') return value;
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
    if (!el) return null;
    const v = parseFloat(el.value);
    return isNaN(v) ? null : v;
  }

  function getAnalysisYears() {
    const el = document.getElementById('analysis-years');
    if (!el) return null;
    const v = parseFloat(el.value);
    return isNaN(v) ? null : v;
  }

  function getCurrencyLabel() {
    const el = document.getElementById('currency-label');
    if (!el) return '';
    return el.value || '';
  }

  // --------- SCENARIO HANDLING ---------
  function getScenarioLabel(key) {
    switch (key) {
      case 'lowPrice':
        return 'Lower commodity prices (−20% benefits)';
      case 'highCost':
        return 'Higher costs (+20% costs)';
      case 'optimistic':
        return 'Optimistic (benefits +20%, costs −10%)';
      case 'base':
      default:
        return 'Base case (as entered)';
    }
  }

  function adjustForScenario(pvBenefits, pvCosts) {
    let b = pvBenefits;
    let c = pvCosts;

    switch (currentScenario) {
      case 'lowPrice':
        b = b * 0.8;
        break;
      case 'highCost':
        c = c * 1.2;
        break;
      case 'optimistic':
        b = b * 1.2;
        c = c * 0.9;
        break;
      case 'base':
      default:
        break;
    }
    return { pvB: b, pvC: c };
  }

  function calculateBaseMetrics(row) {
    const pvB = toNumber(row.pvBenefits);
    const pvC = toNumber(row.pvCosts);
    const npv = pvB - pvC;
    const bcr = pvC > 0 ? pvB / pvC : null;
    const roi = pvC > 0 ? npv / pvC : null;
    return { npv, bcr, roi };
  }

  function recalcAllBaseMetrics() {
    treatments.forEach(function (t) {
      const m = calculateBaseMetrics(t);
      t.npv = m.npv;
      t.bcr = m.bcr;
      t.roi = m.roi;
    });
  }

  function getScenarioMetrics(t) {
    const adj = adjustForScenario(toNumber(t.pvBenefits), toNumber(t.pvCosts));
    const npv = adj.pvB - adj.pvC;
    const bcr = adj.pvC > 0 ? adj.pvB / adj.pvC : null;
    const roi = adj.pvC > 0 ? npv / adj.pvC : null;
    return {
      pvBenefits: adj.pvB,
      pvCosts: adj.pvC,
      npv: npv,
      bcr: bcr,
      roi: roi
    };
  }

  // --------- SYNC INPUTS <-> STATE ---------
  function syncFromInputs() {
    const tbody = document.querySelector('#inputs-tbody');
    if (!tbody) return;

    treatments = [];
    nextId = 1;

    const rows = tbody.querySelectorAll('tr');
    rows.forEach(function (tr) {
      const idAttr = tr.getAttribute('data-id');
      const id = idAttr ? parseInt(idAttr, 10) : nextId;

      const nameInput = tr.querySelector('.treatment-name');
      const pvBInput = tr.querySelector('.pv-benefits');
      const pvCInput = tr.querySelector('.pv-costs');
      const notesInput = tr.querySelector('.treatment-notes');

      const name = nameInput ? (nameInput.value || ('Treatment ' + id)) : ('Treatment ' + id);
      const pvBenefits = pvBInput ? toNumber(pvBInput.value) : 0;
      const pvCosts = pvCInput ? toNumber(pvCInput.value) : 0;
      const notes = notesInput ? (notesInput.value || '') : '';

      treatments.push({
        id: id,
        name: name,
        pvBenefits: pvBenefits,
        pvCosts: pvCosts,
        notes: notes
      });

      if (id >= nextId) {
        nextId = id + 1;
      }
    });

    recalcAllBaseMetrics();
  }

  function syncToInputs() {
    const tbody = document.querySelector('#inputs-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    treatments.forEach(function (t) {
      const tr = document.createElement('tr');
      tr.setAttribute('data-id', String(t.id));

      const nameTd = document.createElement('td');
      const nameInput = document.createElement('input');
      nameInput.type = 'text';
      nameInput.className = 'treatment-name';
      nameInput.value = t.name;
      nameTd.appendChild(nameInput);

      const pvBTd = document.createElement('td');
      const pvBInput = document.createElement('input');
      pvBInput.type = 'number';
      pvBInput.step = 'any';
      pvBInput.className = 'pv-benefits';
      pvBInput.value = t.pvBenefits;
      pvBTd.appendChild(pvBInput);

      const pvCTd = document.createElement('td');
      const pvCInput = document.createElement('input');
      pvCInput.type = 'number';
      pvCInput.step = 'any';
      pvCInput.className = 'pv-costs';
      pvCInput.value = t.pvCosts;
      pvCTd.appendChild(pvCInput);

      const notesTd = document.createElement('td');
      const notesInput = document.createElement('input');
      notesInput.type = 'text';
      notesInput.className = 'treatment-notes';
      notesInput.value = t.notes || '';
      notesTd.appendChild(notesInput);

      const removeTd = document.createElement('td');
      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'btn btn-secondary btn-small remove-treatment';
      removeBtn.textContent = 'Remove';
      removeTd.appendChild(removeBtn);

      tr.appendChild(nameTd);
      tr.appendChild(pvBTd);
      tr.appendChild(pvCTd);
      tr.appendChild(notesTd);
      tr.appendChild(removeTd);

      tbody.appendChild(tr);
    });
  }

  // --------- BASELINE DETECTION ---------
  function findBaselineIndex(arr) {
    if (!arr.length) return -1;
    // Prefer treatments whose name suggests control or current practice
    const idx = arr.findIndex(function (t) {
      const name = (t.name || '').toLowerCase();
      return name.includes('control') || name.includes('current');
    });
    if (idx >= 0) return idx;
    // Fall back to first treatment
    return 0;
  }

  // --------- RESULTS RENDERING ---------
  function renderResultsTable() {
    const tbody = document.querySelector('#results-tbody');
    const summaryDiv = document.getElementById('results-summary');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (!treatments.length) {
      if (summaryDiv) {
        summaryDiv.textContent = 'Add treatments and click “Update results” to see the ranked matrix.';
      }
      return;
    }

    const metricsList = treatments.map(function (t) {
      return {
        t: t,
        metrics: getScenarioMetrics(t)
      };
    });

    const baselineIndex = findBaselineIndex(metricsList.map(function (x) { return x.t; }));
    const baseline = metricsList[baselineIndex];
    const baselineMetrics = baseline.metrics;

    const sorted = metricsList.slice().sort(function (a, b) {
      return (b.metrics.npv || 0) - (a.metrics.npv || 0);
    });

    sorted.forEach(function (item, idx) {
      const tr = document.createElement('tr');
      if (item.t.id === baseline.t.id) {
        tr.classList.add('row-baseline');
      }

      const rankTd = document.createElement('td');
      rankTd.textContent = String(idx + 1);
      tr.appendChild(rankTd);

      const nameTd = document.createElement('td');
      nameTd.textContent = item.t.name;
      tr.appendChild(nameTd);

      const pvBTd = document.createElement('td');
      pvBTd.textContent = formatNumber(item.metrics.pvBenefits, 0);
      tr.appendChild(pvBTd);

      const pvCTd = document.createElement('td');
      pvCTd.textContent = formatNumber(item.metrics.pvCosts, 0);
      tr.appendChild(pvCTd);

      const npvTd = document.createElement('td');
      npvTd.textContent = formatNumber(item.metrics.npv, 0);
      tr.appendChild(npvTd);

      const deltaTd = document.createElement('td');
      const delta = item.metrics.npv - baselineMetrics.npv;
      deltaTd.textContent = formatNumber(delta, 0);
      tr.appendChild(deltaTd);

      const bcrTd = document.createElement('td');
      bcrTd.textContent = item.metrics.bcr != null ? formatNumber(item.metrics.bcr, 2) : '';
      tr.appendChild(bcrTd);

      const roiTd = document.createElement('td');
      roiTd.textContent = item.metrics.roi != null ? formatNumber(item.metrics.roi, 2) : '';
      tr.appendChild(roiTd);

      tbody.appendChild(tr);
    });

    if (summaryDiv) {
      const positive = metricsList.filter(function (x) {
        return x.metrics.npv > 0;
      }).length;
      const total = metricsList.length;
      const best = sorted[0];
      const scenarioText = getScenarioLabel(currentScenario);

      summaryDiv.textContent =
        scenarioText + ': ' +
        positive + ' of ' + total + ' treatments have positive NPV. ' +
        'Best performer: "' + best.t.name + '" (NPV ' + formatNumber(best.metrics.npv, 0) + '). ' +
        'Baseline: "' + baseline.t.name + '" (NPV ' + formatNumber(baselineMetrics.npv, 0) + ').';
    }
  }

  // --------- SNAPSHOTS ---------
  function renderSnapshots() {
    const container = document.querySelector('#snapshots-container');
    const scenarioLabelEl = document.getElementById('snapshots-scenario-label');
    if (!container) return;
    container.innerHTML = '';

    if (scenarioLabelEl) {
      scenarioLabelEl.textContent = 'Currently viewing: ' + getScenarioLabel(currentScenario) + '.';
    }

    if (!treatments.length) return;

    const currency = getCurrencyLabel();

    treatments.forEach(function (t) {
      const m = getScenarioMetrics(t);

      const card = document.createElement('div');
      card.className = 'snapshot-card';

      const title = document.createElement('h3');
      title.textContent = t.name;
      card.appendChild(title);

      const p1 = document.createElement('p');
      p1.textContent = 'PV benefits: ' + currency + ' ' + formatNumber(m.pvBenefits, 0);
      card.appendChild(p1);

      const p2 = document.createElement('p');
      p2.textContent = 'PV costs: ' + currency + ' ' + formatNumber(m.pvCosts, 0);
      card.appendChild(p2);

      const p3 = document.createElement('p');
      p3.textContent = 'Net present value (NPV): ' + currency + ' ' + formatNumber(m.npv, 0);
      card.appendChild(p3);

      const p4 = document.createElement('p');
      const bcrText = m.bcr != null ? formatNumber(m.bcr, 2) : 'n/a';
      const roiText = m.roi != null ? formatNumber(m.roi, 2) : 'n/a';
      p4.textContent = 'BCR: ' + bcrText + ' | ROI: ' + roiText;
      card.appendChild(p4);

      if (t.notes) {
        const p5 = document.createElement('p');
        p5.textContent = 'Notes: ' + t.notes;
        card.appendChild(p5);
      }

      container.appendChild(card);
    });
  }

  // --------- INTERPRETATION HELPER ---------
  function renderHelper() {
    const container = document.querySelector('#helper-text');
    if (!container) return;

    if (!treatments.length) {
      container.value = 'Add at least one treatment and run the analysis to see the plain-language summary here.';
      return;
    }

    const metricsList = treatments.map(function (t) {
      return {
        t: t,
        metrics: getScenarioMetrics(t)
      };
    });

    const sorted = metricsList.slice().sort(function (a, b) {
      return (b.metrics.npv || 0) - (a.metrics.npv || 0);
    });

    const best = sorted[0];
    const worst = sorted[sorted.length - 1];

    const baselineIndex = findBaselineIndex(metricsList.map(function (x) { return x.t; }));
    const baseline = metricsList[baselineIndex];

    const dr = getDiscountRate();
    const years = getAnalysisYears();
    const currency = getCurrencyLabel();
    const scenarioText = getScenarioLabel(currentScenario);

    let text = '';
    text += 'This summary helps interpret the economic results for your farm.\n\n';
    if (dr != null && years != null) {
      text += 'Your PV numbers are assumed to be calculated over approximately ' + years +
        ' years using a discount rate of about ' + dr + '% per year. ';
    } else if (dr != null) {
      text += 'Your PV numbers are assumed to use a discount rate of about ' + dr + '% per year. ';
    }
    text += 'Net present value (NPV) shows the overall gain after subtracting costs from benefits across the whole period. ';
    text += 'A positive NPV means the treatment is expected to return more benefits than it costs in today\'s ' +
      (currency ? '(' + currency + ')' : '') + ' terms.\n\n';

    text += 'The current scenario is: ' + scenarioText + '. ';
    text += 'Under this scenario, the treatment with the highest NPV is "' + best.t.name + '". ';
    text += 'Its NPV is ' + currency + ' ' + formatNumber(best.metrics.npv, 0) +
      ', with present value benefits of ' + currency + ' ' + formatNumber(best.metrics.pvBenefits, 0) +
      ' and present value costs of ' + currency + ' ' + formatNumber(best.metrics.pvCosts, 0) + '. ';

    if (best.metrics.bcr != null) {
      text += 'Its benefit–cost ratio (BCR) is ' + formatNumber(best.metrics.bcr, 2) +
        ', meaning that each dollar of cost is expected to return about ' +
        formatNumber(best.metrics.bcr, 2) + ' dollars in benefits. ';
    }

    text += '\n\nThe baseline option is "' + baseline.t.name + '". ';
    text += 'Compared with this baseline, "' + best.t.name + '" has an NPV that is higher by about ' +
      currency + ' ' + formatNumber(best.metrics.npv - baseline.metrics.npv, 0) + '. ';

    if (sorted.length > 1) {
      text += '\n\nFor contrast, the treatment with the lowest NPV in this scenario is "' + worst.t.name + '" ';
      text += 'with an NPV of ' + currency + ' ' + formatNumber(worst.metrics.npv, 0) + '. ';
      if (worst.metrics.bcr != null) {
        text += 'Its BCR is ' + formatNumber(worst.metrics.bcr, 2) + ', which is weaker than the best option. ';
      }
    }

    text += '\n\nThese figures do not tell you what to choose. ';
    text += 'They are decision support numbers to help you think about which options deliver stronger long-term gains ';
    text += 'for the money invested, compared with your current practice. ';
    text += 'You should also consider risk (for example how results might change under dry years or low prices), labour ';
    text += 'requirements, cash flow timing, and how well each treatment fits with your wider farm plan and personal goals.\n';

    container.value = text;
  }

  // --------- CSV IMPORT ---------
  function parseCsvLine(line) {
    const cells = [];
    let current = '';
    let insideQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const ch = line[i];

      if (ch === '"') {
        if (insideQuotes && line[i + 1] === '"') {
          // Escaped quote
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

  function handleCsvText(text) {
    const rows = parseCsv(text);
    if (!rows.length || rows.length < 2) {
      alert('CSV file appears to be empty or has no data rows.');
      return;
    }

    const header = rows[0].map(function (h) {
      return h.trim().toLowerCase();
    });

    function findIndex(matchFn) {
      return header.findIndex(matchFn);
    }

    const idxName = findIndex(function (h) { return h.includes('treatment') || h.includes('name'); });
    const idxBenefits = findIndex(function (h) { return h.includes('benefit'); });
    const idxCosts = findIndex(function (h) { return h.includes('cost'); });
    const idxNotes = findIndex(function (h) { return h.includes('note'); });

    if (idxName === -1 || idxBenefits === -1 || idxCosts === -1) {
      alert('CSV must include columns for treatment name, PV benefits, and PV costs.');
      return;
    }

    const imported = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;

      const rawName = row[idxName] != null ? String(row[idxName]).trim() : '';
      const rawBenefits = row[idxBenefits] != null ? row[idxBenefits] : '';
      const rawCosts = row[idxCosts] != null ? row[idxCosts] : '';
      const rawNotes = idxNotes !== -1 && row[idxNotes] != null ? row[idxNotes] : '';

      if (!rawName && !rawBenefits && !rawCosts && !rawNotes) {
        continue;
      }

      imported.push({
        id: imported.length + 1,
        name: rawName || ('Treatment ' + (imported.length + 1)),
        pvBenefits: toNumber(rawBenefits),
        pvCosts: toNumber(rawCosts),
        notes: String(rawNotes || '')
      });
    }

    if (!imported.length) {
      alert('No usable data rows found in CSV.');
      return;
    }

    treatments = imported;
    nextId = imported.length + 1;
    recalcAllBaseMetrics();
    syncToInputs();
    updateAllOutputs();
  }

  // --------- EXPORT CSV ---------
  function exportCSV() {
    if (!treatments.length) {
      alert('Nothing to export. Please add at least one treatment.');
      return;
    }

    const header = [
      'Treatment',
      'PV_benefits',
      'PV_costs',
      'NPV',
      'BCR',
      'ROI',
      'Notes'
    ];
    const rows = [header.join(',')];

    treatments.forEach(function (t) {
      rows.push([
        '"' + String(t.name).replace(/"/g, '""') + '"',
        toNumber(t.pvBenefits),
        toNumber(t.pvCosts),
        t.npv != null ? t.npv : '',
        t.bcr != null ? t.bcr : '',
        t.roi != null ? t.roi : '',
        '"' + String(t.notes || '').replace(/"/g, '""') + '"'
      ].join(','));
    });

    const blob = new Blob([rows.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'farming_cba_results.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // --------- DEMO & CORE ACTIONS ---------
  function addBlankTreatment() {
    treatments.push({
      id: nextId++,
      name: 'Treatment ' + (treatments.length + 1),
      pvBenefits: 0,
      pvCosts: 0,
      notes: ''
    });
    syncToInputs();
  }

  function clearAll() {
    treatments = [];
    nextId = 1;
    syncToInputs();
    renderResultsTable();
    renderSnapshots();
    renderHelper();
  }

  function loadDemo() {
    treatments = [
      {
        id: 1,
        name: 'Control / Current practice',
        pvBenefits: 0,
        pvCosts: 0,
        notes: 'Baseline for comparison. Existing farm practice without new investment.'
      },
      {
        id: 2,
        name: 'Improved fertiliser program',
        pvBenefits: 480000,
        pvCosts: 260000,
        notes: 'More precise nutrient management with moderate upfront costs.'
      },
      {
        id: 3,
        name: 'Precision irrigation upgrade',
        pvBenefits: 620000,
        pvCosts: 320000,
        notes: 'Higher capital cost but strong gains in yield and water efficiency.'
      },
      {
        id: 4,
        name: 'Drought-resilient seed & soil package',
        pvBenefits: 560000,
        pvCosts: 300000,
        notes: 'Package focused on stabilising yields in dry seasons.'
      },
      {
        id: 5,
        name: 'Mixed precision + drought package',
        pvBenefits: 710000,
        pvCosts: 420000,
        notes: 'Combines irrigation control with drought-resilient seed and soil management.'
      }
    ];
    nextId = 6;
    recalcAllBaseMetrics();
    syncToInputs();
    updateAllOutputs();
  }

  function updateAllOutputs() {
    syncFromInputs();
    renderResultsTable();
    renderSnapshots();
    renderHelper();
  }

  // --------- TABS ---------
  function setupTabs() {
    const tabButtons = document.querySelectorAll('[data-tab]');
    const panels = document.querySelectorAll('.tab-panel');

    if (!tabButtons.length || !panels.length) {
      return;
    }

    tabButtons.forEach(function (btn) {
      btn.addEventListener('click', function () {
        const target = btn.getAttribute('data-tab');
        if (!target) return;

        tabButtons.forEach(function (b) {
          b.classList.remove('active');
        });
        btn.classList.add('active');

        panels.forEach(function (panel) {
          if (panel.id === 'tab-' + target) {
            panel.classList.add('active');
          } else {
            panel.classList.remove('active');
          }
        });
      });
    });
  }

  // --------- EVENT HANDLERS ---------
  function setupEventHandlers() {
    const addBtn = document.getElementById('btn-add-treatment');
    if (addBtn) {
      addBtn.addEventListener('click', function () {
        addBlankTreatment();
      });
    }

    const updateBtn = document.getElementById('btn-update-results');
    if (updateBtn) {
      updateBtn.addEventListener('click', function () {
        updateAllOutputs();
      });
    }

    const clearBtn = document.getElementById('btn-clear-all');
    if (clearBtn) {
      clearBtn.addEventListener('click', function () {
        if (confirm('Clear all treatments and reset the tool?')) {
          clearAll();
        }
      });
    }

    const demoBtn = document.getElementById('btn-load-demo');
    if (demoBtn) {
      demoBtn.addEventListener('click', function () {
        if (confirm('Replace current inputs with the demo farm scenario?')) {
          loadDemo();
        }
      });
    }

    const exportBtn = document.getElementById('btn-export-csv');
    if (exportBtn) {
      exportBtn.addEventListener('click', function () {
        // Use current base metrics for export
        syncFromInputs();
        exportCSV();
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
        // Re-render outputs under new scenario, using latest treatment state
        renderResultsTable();
        renderSnapshots();
        renderHelper();
      });
    }

    const inputsTable = document.getElementById('inputs-tbody');
    if (inputsTable) {
      inputsTable.addEventListener('click', function (ev) {
        const target = ev.target;
        if (!(target instanceof HTMLElement)) return;
        if (target.classList.contains('remove-treatment')) {
          const tr = target.closest('tr');
          if (!tr) return;
          const idAttr = tr.getAttribute('data-id');
          const id = idAttr ? parseInt(idAttr, 10) : null;
          if (id != null) {
            treatments = treatments.filter(function (t) {
              return t.id !== id;
            });
            syncToInputs();
            updateAllOutputs();
          }
        }
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
          handleCsvText(text);
          fileInput.value = '';
        };
        reader.readAsText(file);
      });
    }

    const discountInput = document.getElementById('discount-rate');
    if (discountInput) {
      discountInput.addEventListener('input', function () {
        renderHelper();
      });
    }

    const yearsInput = document.getElementById('analysis-years');
    if (yearsInput) {
      yearsInput.addEventListener('input', function () {
        renderHelper();
      });
    }

    const currencyInput = document.getElementById('currency-label');
    if (currencyInput) {
      currencyInput.addEventListener('input', function () {
        renderSnapshots();
        renderHelper();
      });
    }
  }

  // --------- INIT ---------
  function init() {
    setupTabs();
    setupEventHandlers();
    addBlankTreatment(); // start with one blank row for user
    renderResultsTable();
    renderSnapshots();
    renderHelper();
  }

  init();
});
