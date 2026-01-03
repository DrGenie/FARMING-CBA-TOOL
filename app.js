// script.js – Farming CBA Decision Tool 2 (all tabs and buttons active)
document.addEventListener('DOMContentLoaded', function () {
  'use strict';

  // --------- STATE ---------
  let treatments = [];
  let nextId = 1;

  // --------- HELPERS ---------
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

  function calculateMetrics(row) {
    const pvB = toNumber(row.pvBenefits);
    const pvC = toNumber(row.pvCosts);
    const npv = pvB - pvC;
    const bcr = pvC > 0 ? pvB / pvC : null;
    const roi = pvC > 0 ? npv / pvC : null;
    return { npv, bcr, roi };
  }

  function recalcAll() {
    treatments.forEach(function (t) {
      const m = calculateMetrics(t);
      t.npv = m.npv;
      t.bcr = m.bcr;
      t.roi = m.roi;
    });
  }

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

    recalcAll();
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

  function renderResultsTable() {
    const tbody = document.querySelector('#results-tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (!treatments.length) return;

    const sorted = treatments.slice().sort(function (a, b) {
      return (b.npv || 0) - (a.npv || 0);
    });

    sorted.forEach(function (t, idx) {
      const tr = document.createElement('tr');

      const rankTd = document.createElement('td');
      rankTd.textContent = String(idx + 1);
      tr.appendChild(rankTd);

      const nameTd = document.createElement('td');
      nameTd.textContent = t.name;
      tr.appendChild(nameTd);

      const pvBTd = document.createElement('td');
      pvBTd.textContent = formatNumber(t.pvBenefits, 0);
      tr.appendChild(pvBTd);

      const pvCTd = document.createElement('td');
      pvCTd.textContent = formatNumber(t.pvCosts, 0);
      tr.appendChild(pvCTd);

      const npvTd = document.createElement('td');
      npvTd.textContent = formatNumber(t.npv, 0);
      tr.appendChild(npvTd);

      const bcrTd = document.createElement('td');
      bcrTd.textContent = t.bcr != null ? formatNumber(t.bcr, 2) : '';
      tr.appendChild(bcrTd);

      const roiTd = document.createElement('td');
      roiTd.textContent = t.roi != null ? formatNumber(t.roi, 2) : '';
      tr.appendChild(roiTd);

      tbody.appendChild(tr);
    });
  }

  function renderSnapshots() {
    const container = document.querySelector('#snapshots-container');
    if (!container) return;
    container.innerHTML = '';

    if (!treatments.length) return;

    treatments.forEach(function (t) {
      const card = document.createElement('div');
      card.className = 'snapshot-card';

      const title = document.createElement('h3');
      title.textContent = t.name;
      card.appendChild(title);

      const p1 = document.createElement('p');
      p1.textContent = 'PV benefits: ' + formatNumber(t.pvBenefits, 0);
      card.appendChild(p1);

      const p2 = document.createElement('p');
      p2.textContent = 'PV costs: ' + formatNumber(t.pvCosts, 0);
      card.appendChild(p2);

      const p3 = document.createElement('p');
      p3.textContent = 'Net present value (NPV): ' + formatNumber(t.npv, 0);
      card.appendChild(p3);

      const p4 = document.createElement('p');
      const bcrText = t.bcr != null ? formatNumber(t.bcr, 2) : 'n/a';
      const roiText = t.roi != null ? formatNumber(t.roi, 2) : 'n/a';
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

  function renderHelper() {
    const container = document.querySelector('#helper-text');
    if (!container) return;

    if (!treatments.length) {
      container.value = 'Add at least one treatment and run the analysis to see the plain-language summary here.';
      return;
    }

    const sorted = treatments.slice().sort(function (a, b) {
      return (b.npv || 0) - (a.npv || 0);
    });

    const best = sorted[0];
    const worst = sorted[sorted.length - 1];

    let text = '';
    text += 'This summary helps interpret the economic results for your farm.\n\n';
    text += 'Net present value (NPV) shows the overall gain after subtracting costs from benefits across the whole analysis period. ';
    text += 'A positive NPV means the treatment is expected to return more benefits than it costs in today\'s dollars.\n\n';

    text += 'In your current comparison, the treatment with the highest NPV is "' + best.name + '". ';
    text += 'Its NPV is ' + formatNumber(best.npv, 0) + ', with present value benefits of ' + formatNumber(best.pvBenefits, 0) +
      ' and present value costs of ' + formatNumber(best.pvCosts, 0) + '. ';

    if (best.bcr != null) {
      text += 'Its benefit–cost ratio (BCR) is ' + formatNumber(best.bcr, 2) +
        ', meaning that each dollar of cost is expected to return about ' + formatNumber(best.bcr, 2) + ' dollars in benefits. ';
    }

    if (sorted.length > 1) {
      text += '\n\nFor comparison, the treatment with the lowest NPV is "' + worst.name + '" ';
      text += 'with an NPV of ' + formatNumber(worst.npv, 0) + '. ';
      if (worst.bcr != null) {
        text += 'Its BCR is ' + formatNumber(worst.bcr, 2) + '. ';
      }
    }

    text += '\n\nThese figures do not tell you what to choose. ';
    text += 'They are decision support numbers to help you think about which options deliver stronger long-term gains for the money invested. ';
    text += 'You can also consider risk, labour requirements, and how well each treatment fits with your wider farm plan.\n';

    container.value = text;
  }

  function updateAllOutputs() {
    syncFromInputs();
    renderResultsTable();
    renderSnapshots();
    renderHelper();
  }

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
        notes: 'Higher capital cost but big gains in yield and water efficiency.'
      },
      {
        id: 4,
        name: 'Drought-resilient seed and soil package',
        pvBenefits: 560000,
        pvCosts: 300000,
        notes: 'Balanced package that mainly protects against dry seasons.'
      }
    ];
    nextId = 5;
    recalcAll();
    syncToInputs();
    updateAllOutputs();
  }

  function exportCSV() {
    if (!treatments.length) {
      alert('Nothing to export. Please add at least one treatment.');
      return;
    }

    const header = [
      'Treatment',
      'PV benefits',
      'PV costs',
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
        updateAllOutputs();
        exportCSV();
      });
    }

    const printBtn = document.getElementById('btn-print');
    if (printBtn) {
      printBtn.addEventListener('click', function () {
        window.print();
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
  }

  function init() {
    setupTabs();
    setupEventHandlers();
    addBlankTreatment(); // start with one blank row
    renderHelper();
  }

  init();
});
