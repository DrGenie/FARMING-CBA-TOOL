// app.js
(() => {
  "use strict";

  const DEFAULT_FILE_NAME = "faba_beans_trial_clean_named.tsv";

  const state = {
    rawRows: [],
    headers: [],
    mapping: {
      replicate: null,
      treatment: null,
      control: null,
      yield: null,
      variableCost: null,
      capitalCost: null
    },
    treatmentSummaries: [],
    controlTreatment: null,
    includeTreatments: new Set(),
    settings: {
      grainPrice: 400,
      discountRate: 4,
      horizonYears: 10,
      persistence: 1
    },
    extraBenefits: {},
    extraCosts: {},
    costRecurrence: {},
    scenarios: [],
    charts: {
      npvVsControl: null,
      costBenefit: null,
      scenarios: null
    },
    notes: {
      overview: "",
      results: "",
      benefits: "",
      costs: "",
      simulation: "",
      working: ""
    },
    baseResults: null
  };

  const DOM = {};

  document.addEventListener("DOMContentLoaded", () => {
    cacheDom();
    setupTabs();
    setupButtonsAndInputs();
    loadNotesFromStorage();
    loadDefaultDataset();
  });

  function cacheDom() {
    DOM.tabButtons = document.querySelectorAll(".tab-btn");
    DOM.tabPanels = document.querySelectorAll(".tab-panel");

    DOM.dataSummary = document.getElementById("dataSummary");
    DOM.columnMappingTableCells = document.querySelectorAll("[data-col-purpose]");
    DOM.dataChecksList = document.getElementById("dataChecksList");
    DOM.dataPreviewTable = document.getElementById("dataPreviewTable");

    DOM.dataFileInput = document.getElementById("dataFileInput");
    DOM.dataPasteArea = document.getElementById("dataPasteArea");
    DOM.loadFileBtn = document.getElementById("loadFileBtn");
    DOM.loadPasteBtn = document.getElementById("loadPasteBtn");
    DOM.reloadDefaultBtn = document.getElementById("reloadDefaultBtn");

    DOM.treatmentsConfigTable = document.getElementById("treatmentsConfigTable");
    DOM.costRecurrenceTable = document.getElementById("costRecurrenceTable");
    DOM.yieldSummaryTable = document.getElementById("yieldSummaryTable");
    DOM.benefitConfigTable = document.getElementById("benefitConfigTable");
    DOM.baselineCostTable = document.getElementById("baselineCostTable");
    DOM.costConfigTable = document.getElementById("costConfigTable");

    DOM.leaderboardTable = document.getElementById("leaderboardTable");
    DOM.comparisonTable = document.getElementById("comparisonTable");

    DOM.grainPriceInput = document.getElementById("grainPriceInput");
    DOM.discountRateInput = document.getElementById("discountRateInput");
    DOM.horizonYearsInput = document.getElementById("horizonYearsInput");
    DOM.persistenceInput = document.getElementById("persistenceInput");

    DOM.filterAllBtn = document.getElementById("filterAllBtn");
    DOM.filterTopNPVBtn = document.getElementById("filterTopNPVBtn");
    DOM.filterTopBCRBtn = document.getElementById("filterTopBCRBtn");
    DOM.filterImprovedBtn = document.getElementById("filterImprovedBtn");

    DOM.resultsNotes = document.getElementById("resultsNotes");
    DOM.overviewNotes = document.getElementById("overviewNotes");
    DOM.benefitsNotes = document.getElementById("benefitsNotes");
    DOM.costsNotes = document.getElementById("costsNotes");
    DOM.simulationNotes = document.getElementById("simulationNotes");
    DOM.consolidatedNotes = document.getElementById("consolidatedNotes");
    DOM.notesWorkingArea = document.getElementById("notesWorkingArea");

    DOM.exportComparisonCsvBtn = document.getElementById("exportComparisonCsvBtn");
    DOM.exportLeaderboardCsvBtn = document.getElementById("exportLeaderboardCsvBtn");
    DOM.exportExcelBtn = document.getElementById("exportExcelBtn");
    DOM.exportCleanTsvBtn = document.getElementById("exportCleanTsvBtn");
    DOM.exportTreatmentSummaryCsvBtn = document.getElementById("exportTreatmentSummaryCsvBtn");
    DOM.exportSensitivityCsvBtn = document.getElementById("exportSensitivityCsvBtn");

    DOM.addScenarioBtn = document.getElementById("addScenarioBtn");
    DOM.resetScenariosBtn = document.getElementById("resetScenariosBtn");
    DOM.scenarioTable = document.getElementById("scenarioTable");

    DOM.npvVsControlCanvas = document.getElementById("npvVsControlChart");
    DOM.costBenefitCanvas = document.getElementById("costBenefitChart");
    DOM.scenarioCanvas = document.getElementById("scenarioChart");

    DOM.openTechnicalAppendixBtn = document.getElementById("openTechnicalAppendixBtn");
    DOM.openTechnicalAppendixBtn2 = document.getElementById("openTechnicalAppendixBtn2");

    DOM.toastHost = document.getElementById("toastHost");
    DOM.toastRegion = document.getElementById("toastRegion");

    // AI briefing elements
    DOM.openCopilotBtn = document.getElementById("openCopilotBtn");
    DOM.generateAiPromptBtn = document.getElementById("generateAiPromptBtn");
    DOM.copyCopilotPromptBtn = document.getElementById("copyCopilotPromptBtn");
    DOM.copyChatGptPromptBtn = document.getElementById("copyChatGptPromptBtn");
    DOM.aiCopilotPromptArea = document.getElementById("aiCopilotPromptArea");
    DOM.aiChatGptPromptArea = document.getElementById("aiChatGptPromptArea");
  }

  /* Tabs */

  function setupTabs() {
    DOM.tabButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        const target = btn.getAttribute("data-tab-target");
        if (!target) return;
        setActiveTab(target);
      });
    });
    // Start on overview
    setActiveTab("tab-overview");
  }

  function setActiveTab(panelId) {
    DOM.tabButtons.forEach((btn) => {
      const target = btn.getAttribute("data-tab-target");
      btn.classList.toggle("is-active", target === panelId);
    });
    DOM.tabPanels.forEach((panel) => {
      panel.hidden = panel.id !== panelId;
    });
    if (panelId === "tab-notes") {
      updateConsolidatedNotes();
    }
  }

  /* Setup buttons and inputs */

  function setupButtonsAndInputs() {
    DOM.loadFileBtn.addEventListener("click", handleLoadFile);
    DOM.loadPasteBtn.addEventListener("click", handleLoadPaste);
    DOM.reloadDefaultBtn.addEventListener("click", () => loadDefaultDataset(true));

    DOM.grainPriceInput.addEventListener("change", handleSettingsChange);
    DOM.discountRateInput.addEventListener("change", handleSettingsChange);
    DOM.horizonYearsInput.addEventListener("change", handleSettingsChange);
    DOM.persistenceInput.addEventListener("change", handleSettingsChange);

    DOM.filterAllBtn.addEventListener("click", () => setFilterMode("all"));
    DOM.filterTopNPVBtn.addEventListener("click", () => setFilterMode("topNPV"));
    DOM.filterTopBCRBtn.addEventListener("click", () => setFilterMode("topBCR"));
    DOM.filterImprovedBtn.addEventListener("click", () => setFilterMode("improved"));

    DOM.resultsNotes.addEventListener("input", () => saveNote("results", DOM.resultsNotes.value));
    DOM.overviewNotes.addEventListener("input", () => saveNote("overview", DOM.overviewNotes.value));
    DOM.benefitsNotes.addEventListener("input", () => saveNote("benefits", DOM.benefitsNotes.value));
    DOM.costsNotes.addEventListener("input", () => saveNote("costs", DOM.costsNotes.value));
    DOM.simulationNotes.addEventListener("input", () => saveNote("simulation", DOM.simulationNotes.value));
    DOM.notesWorkingArea.addEventListener("input", () => saveNote("working", DOM.notesWorkingArea.value));

    DOM.exportComparisonCsvBtn.addEventListener("click", exportComparisonCsv);
    DOM.exportLeaderboardCsvBtn.addEventListener("click", exportLeaderboardCsv);
    DOM.exportExcelBtn.addEventListener("click", exportExcelWorkbook);
    DOM.exportCleanTsvBtn.addEventListener("click", exportCleanDatasetTsv);
    DOM.exportTreatmentSummaryCsvBtn.addEventListener("click", exportTreatmentSummaryCsv);
    DOM.exportSensitivityCsvBtn.addEventListener("click", exportSensitivityCsv);

    DOM.addScenarioBtn.addEventListener("click", addScenario);
    DOM.resetScenariosBtn.addEventListener("click", resetScenarios);

    [DOM.openTechnicalAppendixBtn, DOM.openTechnicalAppendixBtn2].forEach((btn) => {
      if (btn) {
        btn.addEventListener("click", () => {
          window.open("technical-appendix.html", "_blank", "noopener");
        });
      }
    });

    if (DOM.openCopilotBtn) {
      DOM.openCopilotBtn.addEventListener("click", () => {
        window.open("https://copilot.microsoft.com", "_blank", "noopener");
      });
    }

    if (DOM.generateAiPromptBtn) {
      DOM.generateAiPromptBtn.addEventListener("click", handleGenerateAiPrompts);
    }
    if (DOM.copyCopilotPromptBtn) {
      DOM.copyCopilotPromptBtn.addEventListener("click", () =>
        copyToClipboard(
          DOM.aiCopilotPromptArea ? DOM.aiCopilotPromptArea.value : "",
          "Prompt for Copilot copied to clipboard."
        )
      );
    }
    if (DOM.copyChatGptPromptBtn) {
      DOM.copyChatGptPromptBtn.addEventListener("click", () =>
        copyToClipboard(
          DOM.aiChatGptPromptArea ? DOM.aiChatGptPromptArea.value : "",
          "Prompt for ChatGPT copied to clipboard."
        )
      );
    }
  }

  /* Data loading */

  function loadDefaultDataset(showToastFlag) {
    fetch(DEFAULT_FILE_NAME)
      .then((resp) => {
        if (!resp.ok) {
          throw new Error("File not found");
        }
        return resp.text();
      })
      .then((text) => {
        parseAndCommitData(text, "Default trial dataset");
        if (showToastFlag) showToast("Default dataset reloaded.", "success");
      })
      .catch((err) => {
        console.error(err);
        showToast("Could not load default dataset: " + err.message, "error");
        if (DOM.dataSummary) {
          DOM.dataSummary.innerHTML =
            "<p>Failed to load <code>" +
            DEFAULT_FILE_NAME +
            "</code>. You can still upload or paste a dataset on this page.</p>";
        }
      });
  }

  function handleLoadFile() {
    const file = DOM.dataFileInput.files[0];
    if (!file) {
      showToast("Please choose a data file to upload.", "error");
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        parseAndCommitData(e.target.result, file.name);
        showToast("Dataset loaded from file.", "success");
      } catch (err) {
        console.error(err);
        showToast("Could not read file: " + err.message, "error");
      }
    };
    reader.readAsText(file);
  }

  function handleLoadPaste() {
    const text = DOM.dataPasteArea.value.trim();
    if (!text) {
      showToast("Please paste data to be loaded.", "error");
      return;
    }
    try {
      parseAndCommitData(text, "Pasted data");
      showToast("Dataset loaded from pasted text.", "success");
    } catch (err) {
      console.error(err);
      showToast("Could not read pasted data: " + err.message, "error");
    }
  }

  function parseAndCommitData(rawText, label) {
    const lines = rawText.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n").filter((l) => l.trim() !== "");
    if (lines.length < 2) {
      throw new Error("Dataset must contain a header row and at least one data row.");
    }

    const delimiter = lines[0].includes("\t") ? "\t" : ",";
    const headers = lines[0].split(delimiter).map((h) => h.trim());
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
      const parts = lines[i].split(delimiter);
      const row = {};
      headers.forEach((h, idx) => {
        const raw = parts[idx] !== undefined ? parts[idx].trim() : "";
        if (raw === "" || raw === "NA" || raw === "NaN" || raw === "?") {
          row[h] = null;
        } else {
          const numeric = Number(raw);
          row[h] = isFinite(numeric) ? numeric : raw;
        }
      });
      rows.push(row);
    }

    state.headers = headers;
    state.rawRows = rows;

    state.mapping = inferColumnMapping(headers);
    state.treatmentSummaries = buildTreatmentSummaries();
    state.controlTreatment = detectControlTreatment();
    state.includeTreatments = new Set(state.treatmentSummaries.map((t) => t.treatment));

    initDefaultsForNewDataset();
    renderAfterDataChange(label);
  }

  function inferColumnMapping(headers) {
    const lower = headers.map((h) => h.toLowerCase());
    const find = (candidates) => {
      for (const cand of candidates) {
        const idx = lower.indexOf(cand.toLowerCase());
        if (idx !== -1) return headers[idx];
      }
      return null;
    };

    return {
      replicate: find(["replicate", "replicate_id", "rep"]),
      treatment: find(["treatment_name", "treatment", "amendment_name"]),
      control: find(["is_control", "control", "control_flag"]),
      yield: find(["yield_t_ha", "yield", "yield_t/ha", "grain_yield_t_ha"]),
      variableCost: find(["total_cost_per_ha_raw", "variable_cost_per_ha", "variable_cost", "total_cost_per_ha"]),
      capitalCost: find([
        "capital_cost_per_ha",
        "capital_cost",
        "capital_total_per_ha",
        "capital_22_header_12_m_front"
      ])
    };
  }

  function buildTreatmentSummaries() {
    const map = new Map();
    const {
      treatment: colTreatment,
      control: colControl,
      yield: colYield,
      variableCost: colVarCost
    } = state.mapping;

    state.rawRows.forEach((row) => {
      const tName = String(row[colTreatment] ?? "").trim() || "Unknown";
      if (!map.has(tName)) {
        map.set(tName, {
          treatment: tName,
          yields: [],
          costs: [],
          isControl: false
        });
      }
      const entry = map.get(tName);
      const yVal = Number(row[colYield]);
      if (isFinite(yVal)) entry.yields.push(yVal);

      const cVal = Number(row[colVarCost]);
      if (isFinite(cVal)) entry.costs.push(cVal);

      const controlVal = row[colControl];
      if (String(controlVal).toLowerCase() === "true" || controlVal === 1) {
        entry.isControl = true;
      }
    });

    const summaries = [];
    for (const [, v] of map.entries()) {
      const meanYield = average(v.yields);
      const meanCost = average(v.costs);
      summaries.push({
        treatment: v.treatment,
        isControl: !!v.isControl,
        meanYield: meanYield,
        meanCost: meanCost
      });
    }
    summaries.sort((a, b) => a.treatment.localeCompare(b.treatment));
    return summaries;
  }

  function detectControlTreatment() {
    let explicit = state.treatmentSummaries.find((t) => t.isControl);
    if (!explicit && state.treatmentSummaries.length > 0) {
      explicit = state.treatmentSummaries.find((t) => /control/i.test(t.treatment)) || state.treatmentSummaries[0];
      explicit.isControl = true;
    }
    return explicit || null;
  }

  function initDefaultsForNewDataset() {
    state.extraBenefits = {};
    state.extraCosts = {};
    state.costRecurrence = {};

    state.treatmentSummaries.forEach((t) => {
      const name = t.treatment;
      state.extraBenefits[name] = { annual: 0, onceOff: 0 };
      state.extraCosts[name] = { seed: 0, labour: 0, machinery: 0, chemicals: 0, other: 0 };
      state.costRecurrence[name] = 0.25;
    });

    state.settings.grainPrice = 400;
    state.settings.discountRate = 4;
    state.settings.horizonYears = 10;
    state.settings.persistence = 1;

    if (DOM.grainPriceInput) DOM.grainPriceInput.value = state.settings.grainPrice;
    if (DOM.discountRateInput) DOM.discountRateInput.value = state.settings.discountRate;
    if (DOM.horizonYearsInput) DOM.horizonYearsInput.value = state.settings.horizonYears;
    if (DOM.persistenceInput) DOM.persistenceInput.value = state.settings.persistence;

    resetScenariosInternal();
  }

  function renderAfterDataChange(label) {
    renderDataSummary(label);
    renderColumnMappingTable();
    runDataChecks();
    renderDataPreview();
    renderConfigTables();
    recomputeAll();
  }

  function renderDataSummary(label) {
    const n = state.rawRows.length;
    const treatments = state.treatmentSummaries.length;
    const controlName = state.controlTreatment ? state.controlTreatment.treatment : "Not detected";
    DOM.dataSummary.innerHTML =
      "<p><strong>Source:</strong> " +
      escapeHtml(label) +
      "</p><p><strong>Rows:</strong> " +
      n +
      " &nbsp; • &nbsp; <strong>Treatments:</strong> " +
      treatments +
      " &nbsp; • &nbsp; <strong>Control:</strong> " +
      escapeHtml(controlName) +
      "</p>";
  }

  function renderColumnMappingTable() {
    DOM.columnMappingTableCells.forEach((cell) => {
      const purpose = cell.getAttribute("data-col-purpose");
      const value = state.mapping[purpose] || "Not found";
      cell.textContent = value;
      if (!state.mapping[purpose]) {
        cell.style.color = "#b00020";
      } else {
        cell.style.color = "";
      }
    });
  }

  function runDataChecks() {
    const messages = [];
    if (!state.mapping.replicate) {
      messages.push("Replicate column not found. Replicate-level reporting may be limited.");
    }
    if (!state.mapping.variableCost) {
      messages.push("Variable cost column not found. Cost-based indicators may be incomplete.");
    }
    if (!state.mapping.capitalCost) {
      messages.push("Capital cost column not found. Capital cost items will be treated as zero unless added in configuration.");
    }

    const yieldCol = state.mapping.yield;
    if (yieldCol) {
      const missingYield = state.rawRows.filter((r) => !isFinite(Number(r[yieldCol]))).length;
      if (missingYield > 0) {
        messages.push(missingYield + " rows have missing or non-numeric yield values.");
      }
    }

    const controlCol = state.mapping.control;
    if (controlCol) {
      const controlRows = state.rawRows.filter((r) => {
        const v = r[controlCol];
        return String(v).toLowerCase() === "true" || v === 1;
      }).length;
      if (controlRows === 0) {
        messages.push("No rows are marked as control in the control flag column.");
      }
    }

    DOM.dataChecksList.innerHTML = "";
    if (messages.length === 0) {
      const li = document.createElement("li");
      li.textContent = "No critical issues detected in the loaded dataset.";
      DOM.dataChecksList.appendChild(li);
    } else {
      messages.forEach((m) => {
        const li = document.createElement("li");
        li.textContent = m;
        DOM.dataChecksList.appendChild(li);
      });
    }
  }

  function renderDataPreview() {
    const table = DOM.dataPreviewTable;
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");
    thead.innerHTML = "";
    tbody.innerHTML = "";
    if (!state.headers.length) return;

    const trHead = document.createElement("tr");
    state.headers.forEach((h) => {
      const th = document.createElement("th");
      th.textContent = h;
      trHead.appendChild(th);
    });
    thead.appendChild(trHead);

    const rowsToShow = state.rawRows.slice(0, 8);
    rowsToShow.forEach((row) => {
      const tr = document.createElement("tr");
      state.headers.forEach((h) => {
        const td = document.createElement("td");
        const val = row[h];
        td.textContent = val === null || val === undefined ? "" : String(val);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
  }

  /* Configuration tables */

  function renderConfigTables() {
    renderTreatmentsConfigTable();
    renderCostRecurrenceTable();
    renderYieldSummaryTable();
    renderBenefitConfigTable();
    renderBaselineCostTable();
    renderCostConfigTable();
  }

  function renderTreatmentsConfigTable() {
    const tbody = DOM.treatmentsConfigTable.querySelector("tbody");
    tbody.innerHTML = "";
    state.treatmentSummaries.forEach((t) => {
      const tr = document.createElement("tr");

      const nameTd = document.createElement("td");
      nameTd.textContent = t.treatment;
      tr.appendChild(nameTd);

      const isControlTd = document.createElement("td");
      isControlTd.textContent =
        state.controlTreatment && state.controlTreatment.treatment === t.treatment ? "Yes (control)" : "No";
      tr.appendChild(isControlTd);

      const setControlTd = document.createElement("td");
      const controlBtn = document.createElement("button");
      controlBtn.type = "button";
      controlBtn.className = "btn-secondary";
      controlBtn.textContent = "Set as control";
      controlBtn.addEventListener("click", () => {
        state.treatmentSummaries.forEach((x) => (x.isControl = x.treatment === t.treatment));
        state.controlTreatment = t;
        showToast("Control treatment set to " + t.treatment + ".", "success");
        recomputeAll();
        renderTreatmentsConfigTable();
      });
      setControlTd.appendChild(controlBtn);
      tr.appendChild(setControlTd);

      const includeTd = document.createElement("td");
      const includeCheckbox = document.createElement("input");
      includeCheckbox.type = "checkbox";
      includeCheckbox.checked = state.includeTreatments.has(t.treatment);
      includeCheckbox.addEventListener("change", () => {
        if (includeCheckbox.checked) state.includeTreatments.add(t.treatment);
        else state.includeTreatments.delete(t.treatment);
        recomputeAll();
      });
      includeTd.appendChild(includeCheckbox);
      tr.appendChild(includeTd);

      const noteTd = document.createElement("td");
      noteTd.textContent = t.isControl ? "Derived from control flag or label." : "";
      tr.appendChild(noteTd);

      tbody.appendChild(tr);
    });
  }

  function renderCostRecurrenceTable() {
    const tbody = DOM.costRecurrenceTable.querySelector("tbody");
    tbody.innerHTML = "";
    state.treatmentSummaries.forEach((t) => {
      const tr = document.createElement("tr");

      const nameTd = document.createElement("td");
      nameTd.textContent = t.treatment;
      tr.appendChild(nameTd);

      const inputTd = document.createElement("td");
      const input = document.createElement("input");
      input.type = "number";
      input.min = "0";
      input.max = "1";
      input.step = "0.05";
      input.value = state.costRecurrence[t.treatment] ?? 0;
      input.addEventListener("change", () => {
        const v = clamp(Number(input.value), 0, 1);
        state.costRecurrence[t.treatment] = isFinite(v) ? v : 0;
        input.value = state.costRecurrence[t.treatment];
        recomputeAll();
      });
      inputTd.appendChild(input);
      tr.appendChild(inputTd);

      tbody.appendChild(tr);
    });
  }

  function renderYieldSummaryTable() {
    const tbody = DOM.yieldSummaryTable.querySelector("tbody");
    tbody.innerHTML = "";
    const control = state.controlTreatment;
    const controlYield = control ? control.meanYield : null;
    state.treatmentSummaries.forEach((t) => {
      const tr = document.createElement("tr");
      const nameTd = document.createElement("td");
      nameTd.textContent = t.treatment;
      tr.appendChild(nameTd);

      const yTd = document.createElement("td");
      yTd.textContent = formatNumber(t.meanYield, 3);
      tr.appendChild(yTd);

      const diffTd = document.createElement("td");
      const diff = controlYield != null ? t.meanYield - controlYield : null;
      diffTd.textContent = controlYield != null ? formatNumber(diff, 3) : "–";
      tr.appendChild(diffTd);

      tbody.appendChild(tr);
    });
  }

  function renderBenefitConfigTable() {
    const tbody = DOM.benefitConfigTable.querySelector("tbody");
    tbody.innerHTML = "";
    state.treatmentSummaries.forEach((t) => {
      const tr = document.createElement("tr");
      const nameTd = document.createElement("td");
      nameTd.textContent = t.treatment;
      tr.appendChild(nameTd);

      const annualTd = document.createElement("td");
      const annualInput = document.createElement("input");
      annualInput.type = "number";
      annualInput.step = "1";
      annualInput.value = state.extraBenefits[t.treatment]?.annual ?? 0;
      annualInput.addEventListener("change", () => {
        const v = Number(annualInput.value) || 0;
        state.extraBenefits[t.treatment].annual = v;
        recomputeAll();
      });
      annualTd.appendChild(annualInput);
      tr.appendChild(annualTd);

      const onceTd = document.createElement("td");
      const onceInput = document.createElement("input");
      onceInput.type = "number";
      onceInput.step = "1";
      onceInput.value = state.extraBenefits[t.treatment]?.onceOff ?? 0;
      onceInput.addEventListener("change", () => {
        const v = Number(onceInput.value) || 0;
        state.extraBenefits[t.treatment].onceOff = v;
        recomputeAll();
      });
      onceTd.appendChild(onceInput);
      tr.appendChild(onceTd);

      const totalTd = document.createElement("td");
      totalTd.textContent = "0";
      tr.appendChild(totalTd);

      tbody.appendChild(tr);
    });
  }

  function renderBaselineCostTable() {
    const tbody = DOM.baselineCostTable.querySelector("tbody");
    tbody.innerHTML = "";
    const control = state.controlTreatment;
    const controlCost = control ? control.meanCost : null;
    state.treatmentSummaries.forEach((t) => {
      const tr = document.createElement("tr");
      const nameTd = document.createElement("td");
      nameTd.textContent = t.treatment;
      tr.appendChild(nameTd);

      const cTd = document.createElement("td");
      cTd.textContent = formatCurrency(t.meanCost);
      tr.appendChild(cTd);

      const diffTd = document.createElement("td");
      const diff = controlCost != null ? t.meanCost - controlCost : null;
      diffTd.textContent = controlCost != null ? formatCurrency(diff) : "–";
      tr.appendChild(diffTd);

      tbody.appendChild(tr);
    });
  }

  function renderCostConfigTable() {
    const tbody = DOM.costConfigTable.querySelector("tbody");
    tbody.innerHTML = "";
    state.treatmentSummaries.forEach((t) => {
      const tr = document.createElement("tr");
      const nameTd = document.createElement("td");
      nameTd.textContent = t.treatment;
      tr.appendChild(nameTd);

      const componentKeys = ["seed", "labour", "machinery", "chemicals", "other"];
      componentKeys.forEach((key) => {
        const td = document.createElement("td");
        const input = document.createElement("input");
        input.type = "number";
        input.step = "1";
        input.value = state.extraCosts[t.treatment]?.[key] ?? 0;
        input.addEventListener("change", () => {
          const v = Number(input.value) || 0;
          state.extraCosts[t.treatment][key] = v;
          recomputeAll();
        });
        td.appendChild(input);
        tr.appendChild(td);
      });

      const totalTd = document.createElement("td");
      totalTd.textContent = "0";
      tr.appendChild(totalTd);

      tbody.appendChild(tr);
    });
  }

  /* Settings */

  function handleSettingsChange() {
    const gp = Number(DOM.grainPriceInput.value);
    const dr = Number(DOM.discountRateInput.value);
    const h = Number(DOM.horizonYearsInput.value);
    const p = Number(DOM.persistenceInput.value);
    state.settings.grainPrice = isFinite(gp) && gp >= 0 ? gp : 0;
    state.settings.discountRate = isFinite(dr) && dr >= 0 ? dr : 0;
    state.settings.horizonYears = isFinite(h) && h >= 1 ? Math.round(h) : 1;
    state.settings.persistence = clamp(isFinite(p) ? p : 1, 0, 1);
    DOM.grainPriceInput.value = state.settings.grainPrice;
    DOM.discountRateInput.value = state.settings.discountRate;
    DOM.horizonYearsInput.value = state.settings.horizonYears;
    DOM.persistenceInput.value = state.settings.persistence;
    recomputeAll();
  }

  /* Economics */

  function recomputeAll() {
    if (!state.rawRows.length || !state.treatmentSummaries.length || !state.controlTreatment) {
      return;
    }
    const baseResults = computeEconomicsForScenario({
      name: "Base",
      settings: {
        ...state.settings
      }
    });
    state.baseResults = baseResults;

    updateBenefitAndCostTotalsFromBase();
    refreshScenarios(baseResults);
    renderResults(baseResults);
    renderFigures(baseResults);
    renderScenarioView();
  }

  function computeEconomicsForScenario(scenario) {
    const results = [];
    const settings = scenario.settings;
    const r = settings.discountRate / 100;
    const n = settings.horizonYears;
    const p = settings.persistence;

    const controlName = state.controlTreatment.treatment;

    state.treatmentSummaries.forEach((t) => {
      if (!state.includeTreatments.has(t.treatment) && t.treatment !== controlName) {
        return;
      }
      const name = t.treatment;
      const baseYield = t.meanYield || 0;
      const baseCost = t.meanCost || 0;

      const extraB = state.extraBenefits[name] || { annual: 0, onceOff: 0 };
      const extraC = state.extraCosts[name] || {
        seed: 0,
        labour: 0,
        machinery: 0,
        chemicals: 0,
        other: 0
      };

      const recurringShare = state.costRecurrence[name] ?? 0.25;

      let pvBenefits = 0;
      for (let year = 0; year < n; year++) {
        const discountFactor = 1 / Math.pow(1 + r, year);
        const persistenceFactor = Math.pow(p, year);
        const annualYieldBenefit = baseYield * settings.grainPrice * persistenceFactor;
        const annualBenefit = annualYieldBenefit + extraB.annual * persistenceFactor;
        if (year === 0) {
          pvBenefits += (annualBenefit + extraB.onceOff) * discountFactor;
        } else {
          pvBenefits += annualBenefit * discountFactor;
        }
      }

      let pvCosts = 0;
      const extraCostTotal = extraC.seed + extraC.labour + extraC.machinery + extraC.chemicals + extraC.other;
      pvCosts += baseCost + extraCostTotal;
      for (let year = 1; year < n; year++) {
        const discountFactor = 1 / Math.pow(1 + r, year);
        const persistenceFactor = Math.pow(p, year);
        const annualCost = (baseCost * recurringShare + extraCostTotal * recurringShare) * persistenceFactor;
        pvCosts += annualCost * discountFactor;
      }

      const npv = pvBenefits - pvCosts;
      const bcr = pvCosts > 0 ? pvBenefits / pvCosts : null;
      const roi = pvCosts > 0 ? (pvBenefits - pvCosts) / pvCosts : null;

      results.push({
        treatment: name,
        isControl: name === controlName,
        pvBenefits,
        pvCosts,
        npv,
        bcr,
        roi
      });
    });

    results.sort((a, b) => (b.npv || 0) - (a.npv || 0));

    let rank = 1;
    results.forEach((rEntry) => {
      rEntry.rank = rank++;
    });

    const controlResult = results.find((rEntry) => rEntry.isControl);
    if (!controlResult) {
      return { scenario, results, control: null };
    }
    const control = controlResult;

    results.forEach((rEntry) => {
      rEntry.deltaNPV = rEntry.npv - control.npv;
      rEntry.deltaCost = rEntry.pvCosts - control.pvCosts;
    });

    return { scenario, results, control };
  }

  function updateBenefitAndCostTotalsFromBase() {
    if (!state.treatmentSummaries.length) return;

    const r = state.settings.discountRate / 100;
    const n = state.settings.horizonYears;
    const p = state.settings.persistence;

    if (DOM.benefitConfigTable) {
      const rows = DOM.benefitConfigTable.querySelectorAll("tbody tr");
      rows.forEach((tr) => {
        const nameCell = tr.cells[0];
        const totalCell = tr.cells[3];
        if (!nameCell || !totalCell) return;
        const name = nameCell.textContent.trim();
        const extraB = state.extraBenefits[name];
        if (!extraB) {
          totalCell.textContent = formatCurrency(0);
          return;
        }
        let pv = 0;
        for (let year = 0; year < n; year++) {
          const discountFactor = 1 / Math.pow(1 + r, year);
          const persistenceFactor = Math.pow(p, year);
          const annual = extraB.annual * persistenceFactor;
          if (year === 0) {
            pv += (annual + extraB.onceOff) * discountFactor;
          } else {
            pv += annual * discountFactor;
          }
        }
        totalCell.textContent = formatCurrency(pv);
      });
    }

    if (DOM.costConfigTable) {
      const rows = DOM.costConfigTable.querySelectorAll("tbody tr");
      rows.forEach((tr) => {
        const nameCell = tr.cells[0];
        const totalCell = tr.cells[6];
        if (!nameCell || !totalCell) return;
        const name = nameCell.textContent.trim();
        const extraC = state.extraCosts[name];
        if (!extraC) {
          totalCell.textContent = formatCurrency(0);
          return;
        }
        const total = extraC.seed + extraC.labour + extraC.machinery + extraC.chemicals + extraC.other;
        totalCell.textContent = formatCurrency(total);
      });
    }
  }

  /* Results rendering */

  let currentFilterMode = "all";

  function setFilterMode(mode) {
    currentFilterMode = mode;
    DOM.filterAllBtn.classList.toggle("is-on", mode === "all");
    DOM.filterTopNPVBtn.classList.toggle("is-on", mode === "topNPV");
    DOM.filterTopBCRBtn.classList.toggle("is-on", mode === "topBCR");
    DOM.filterImprovedBtn.classList.toggle("is-on", mode === "improved");
    if (state.baseResults) {
      renderResults(state.baseResults);
    }
  }

  function filterResultsForDisplay(baseResults) {
    let rows = baseResults.results.slice();
    if (currentFilterMode === "topNPV") {
      rows = rows
        .slice()
        .sort((a, b) => (b.npv || 0) - (a.npv || 0))
        .slice(0, 5);
    } else if (currentFilterMode === "topBCR") {
      rows = rows
        .filter((r) => r.bcr != null)
        .sort((a, b) => (b.bcr || 0) - (a.bcr || 0))
        .slice(0, 5);
    } else if (currentFilterMode === "improved") {
      if (baseResults.control) {
        rows = rows.filter((r) => r.deltaNPV > 0 && !r.isControl);
      }
    }
    return rows;
  }

  function renderResults(baseResults) {
    renderLeaderboard(baseResults);
    renderComparisonTable(baseResults);
  }

  function renderLeaderboard(baseResults) {
    const tbody = DOM.leaderboardTable.querySelector("tbody");
    tbody.innerHTML = "";
    const rows = filterResultsForDisplay(baseResults);

    rows.forEach((r) => {
      const tr = document.createElement("tr");
      const rankTd = document.createElement("td");
      rankTd.textContent = r.rank;
      tr.appendChild(rankTd);

      const nameTd = document.createElement("td");
      nameTd.textContent = r.treatment + (r.isControl ? " (control)" : "");
      tr.appendChild(nameTd);

      const npvTd = document.createElement("td");
      npvTd.textContent = formatCurrency(r.npv);
      tr.appendChild(npvTd);

      const diffTd = document.createElement("td");
      diffTd.textContent = r.isControl ? "–" : formatCurrency(r.deltaNPV);
      tr.appendChild(diffTd);

      const bcrTd = document.createElement("td");
      bcrTd.textContent = r.bcr != null ? formatNumber(r.bcr, 2) : "–";
      tr.appendChild(bcrTd);

      const roiTd = document.createElement("td");
      roiTd.textContent = r.roi != null ? formatPercent(r.roi) : "–";
      tr.appendChild(roiTd);

      const incTd = document.createElement("td");
      incTd.textContent = state.includeTreatments.has(r.treatment) || r.isControl ? "Yes" : "No";
      tr.appendChild(incTd);

      tbody.appendChild(tr);
    });
  }

  function renderComparisonTable(baseResults) {
    const table = DOM.comparisonTable;
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");
    thead.innerHTML = "";
    tbody.innerHTML = "";

    const indicators = [
      { key: "pvBenefits", label: "Total benefits over time (present value)" },
      { key: "pvCosts", label: "Total costs over time (present value)" },
      { key: "npv", label: "Net profit over time (present value)" },
      { key: "bcr", label: "Benefit per dollar spent" },
      { key: "roi", label: "Return on investment" },
      { key: "deltaNPV", label: "Difference in net profit compared with control" },
      { key: "deltaCost", label: "Difference in total cost compared with control" },
      { key: "rank", label: "Overall ranking (1 = best net profit)" }
    ];

    const rowsForDisplay = filterResultsForDisplay(baseResults);
    if (!rowsForDisplay.length) return;

    const trHead = document.createElement("tr");
    const firstTh = document.createElement("th");
    firstTh.textContent = "Indicator";
    trHead.appendChild(firstTh);

    rowsForDisplay.forEach((r) => {
      const th = document.createElement("th");
      th.textContent = r.treatment + (r.isControl ? " (control)" : "");
      trHead.appendChild(th);
    });
    thead.appendChild(trHead);

    indicators.forEach((ind) => {
      const tr = document.createElement("tr");
      const labelTd = document.createElement("td");
      labelTd.textContent = ind.label;
      tr.appendChild(labelTd);

      rowsForDisplay.forEach((r) => {
        const td = document.createElement("td");
        const val = r[ind.key];
        if (["pvBenefits", "pvCosts", "npv", "deltaNPV", "deltaCost"].includes(ind.key)) {
          td.textContent = val == null ? "–" : formatCurrency(val);
        } else if (ind.key === "bcr") {
          td.textContent = val != null ? formatNumber(val, 2) : "–";
        } else if (ind.key === "roi") {
          td.textContent = val != null ? formatPercent(val) : "–";
        } else if (ind.key === "rank") {
          td.textContent = val != null ? String(val) : "–";
        } else {
          td.textContent = val == null ? "–" : String(val);
        }
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });
  }

  /* Figures */

  function renderFigures(baseResults) {
    renderNpvVsControlChart(baseResults);
    renderCostBenefitChart(baseResults);
  }

  function renderNpvVsControlChart(baseResults) {
    if (!DOM.npvVsControlCanvas) return;
    if (typeof Chart === "undefined") return;

    const ctx = DOM.npvVsControlCanvas.getContext("2d");
    if (state.charts.npvVsControl) {
      state.charts.npvVsControl.destroy();
      state.charts.npvVsControl = null;
    }
    const control = baseResults.control;
    if (!control) return;

    const treatments = baseResults.results.filter((r) => !r.isControl);
    const labels = treatments.map((t) => t.treatment);
    const deltas = treatments.map((t) => t.deltaNPV);

    state.charts.npvVsControl = new Chart(ctx, {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: "Net profit gain vs control (per ha)",
            data: deltas
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (context) => "Difference vs control: " + formatCurrency(context.parsed.y)
            }
          }
        },
        scales: {
          x: {
            ticks: { font: { size: 11 } }
          },
          y: {
            ticks: {
              callback: (v) => formatCurrency(v)
            },
            title: {
              display: true,
              text: "Net profit gain vs control (per ha)"
            }
          }
        }
      }
    });
  }

  function renderCostBenefitChart(baseResults) {
    if (!DOM.costBenefitCanvas) return;
    if (typeof Chart === "undefined") return;

    const ctx = DOM.costBenefitCanvas.getContext("2d");
    if (state.charts.costBenefit) {
      state.charts.costBenefit.destroy();
      state.charts.costBenefit = null;
    }

    const rows = baseResults.results.slice().sort((a, b) => a.rank - b.rank);
    const labels = rows.map((r) => (r.isControl ? r.treatment + " (control)" : r.treatment));
    const benefits = rows.map((r) => r.pvBenefits);
    const costs = rows.map((r) => r.pvCosts);

    state.charts.costBenefit = new Chart(ctx, {
      type: "bar",
      data: {
        labels,
        datasets: [
          { label: "Total benefits over time", data: benefits },
          { label: "Total costs over time", data: costs }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          tooltip: {
            callbacks: {
              label: (context) => context.dataset.label + ": " + formatCurrency(context.parsed.y)
            }
          }
        },
        scales: {
          x: {
            ticks: { font: { size: 11 } }
          },
          y: {
            ticks: {
              callback: (v) => formatCurrency(v)
            },
            title: {
              display: true,
              text: "Present value (per ha)"
            }
          }
        }
      }
    });
  }

  /* Scenarios */

  function resetScenariosInternal() {
    state.scenarios = [
      {
        id: makeId(),
        name: "Base",
        settings: { ...state.settings },
        lockedBase: true,
        results: null
      }
    ];
  }

  function refreshScenarios(baseResults) {
    if (!state.scenarios.length) {
      resetScenariosInternal();
    }
    state.scenarios.forEach((sc) => {
      if (sc.lockedBase) {
        sc.settings = { ...state.settings };
      }
      sc.results = computeEconomicsForScenario(sc);
    });
  }

  function renderScenarioView() {
    renderScenarioTable();
    renderScenarioChart();
  }

  function renderScenarioTable() {
    const tbody = DOM.scenarioTable.querySelector("tbody");
    tbody.innerHTML = "";
    state.scenarios.forEach((sc) => {
      const tr = document.createElement("tr");

      const nameTd = document.createElement("td");
      const nameInput = document.createElement("input");
      nameInput.type = "text";
      nameInput.value = sc.name;
      nameInput.disabled = sc.lockedBase;
      nameInput.addEventListener("change", () => {
        sc.name = nameInput.value || sc.name;
        renderScenarioChart();
      });
      nameTd.appendChild(nameInput);
      tr.appendChild(nameTd);

      const gpTd = document.createElement("td");
      const gpInput = document.createElement("input");
      gpInput.type = "number";
      gpInput.step = "1";
      gpInput.value = sc.settings.grainPrice;
      gpInput.disabled = sc.lockedBase;
      gpInput.addEventListener("change", () => {
        sc.settings.grainPrice = Number(gpInput.value) || 0;
        sc.results = computeEconomicsForScenario(sc);
        renderScenarioView();
      });
      gpTd.appendChild(gpInput);
      tr.appendChild(gpTd);

      const drTd = document.createElement("td");
      const drInput = document.createElement("input");
      drInput.type = "number";
      drInput.step = "0.1";
      drInput.value = sc.settings.discountRate;
      drInput.disabled = sc.lockedBase;
      drInput.addEventListener("change", () => {
        sc.settings.discountRate = Number(drInput.value) || 0;
        sc.results = computeEconomicsForScenario(sc);
        renderScenarioView();
      });
      drTd.appendChild(drInput);
      tr.appendChild(drTd);

      const hTd = document.createElement("td");
      const hInput = document.createElement("input");
      hInput.type = "number";
      hInput.step = "1";
      hInput.value = sc.settings.horizonYears;
      hInput.disabled = sc.lockedBase;
      hInput.addEventListener("change", () => {
        sc.settings.horizonYears = Math.max(1, Math.round(Number(hInput.value) || 1));
        sc.results = computeEconomicsForScenario(sc);
        renderScenarioView();
      });
      hTd.appendChild(hInput);
      tr.appendChild(hTd);

      const pTd = document.createElement("td");
      const pInput = document.createElement("input");
      pInput.type = "number";
      pInput.step = "0.05";
      pInput.value = sc.settings.persistence;
      pInput.disabled = sc.lockedBase;
      pInput.addEventListener("change", () => {
        const v = clamp(Number(pInput.value) || 0, 0, 1);
        sc.settings.persistence = v;
        pInput.value = v;
        sc.results = computeEconomicsForScenario(sc);
        renderScenarioView();
      });
      pTd.appendChild(pInput);
      tr.appendChild(pTd);

      const bestTd = document.createElement("td");
      const control = sc.results.control;
      let best = null;
      let gain = null;
      if (control) {
        best = sc.results.results
          .filter((r) => !r.isControl)
          .slice()
          .sort((a, b) => (b.npv || 0) - (a.npv || 0))[0];
        if (best) {
          gain = best.npv - control.npv;
        }
      }
      bestTd.textContent = best ? best.treatment : "Not available";
      tr.appendChild(bestTd);

      const gainTd = document.createElement("td");
      gainTd.textContent = gain != null ? formatCurrency(gain) : "–";
      tr.appendChild(gainTd);

      const delTd = document.createElement("td");
      if (sc.lockedBase) {
        delTd.textContent = "–";
      } else {
        const delBtn = document.createElement("button");
        delBtn.type = "button";
        delBtn.className = "btn-ghost";
        delBtn.textContent = "Remove";
        delBtn.addEventListener("click", () => {
          state.scenarios = state.scenarios.filter((s) => s.id !== sc.id);
          renderScenarioView();
        });
        delTd.appendChild(delBtn);
      }
      tr.appendChild(delTd);

      tbody.appendChild(tr);
    });
  }

  function renderScenarioChart() {
    if (!DOM.scenarioCanvas) return;
    if (typeof Chart === "undefined") return;

    const ctx = DOM.scenarioCanvas.getContext("2d");
    if (state.charts.scenarios) {
      state.charts.scenarios.destroy();
      state.charts.scenarios = null;
    }

    const labels = [];
    const gains = [];

    state.scenarios.forEach((sc) => {
      const control = sc.results.control;
      let best = null;
      let gain = 0;
      if (control) {
        best = sc.results.results
          .filter((r) => !r.isControl)
          .slice()
          .sort((a, b) => (b.npv || 0) - (a.npv || 0))[0];
        if (best) {
          gain = best.npv - control.npv;
        }
      }
      labels.push(sc.name);
      gains.push(gain);
    });

    state.charts.scenarios = new Chart(ctx, {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: "Net profit gain of best treatment vs control (per ha)",
            data: gains
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          tooltip: {
            callbacks: {
              label: (context) => formatCurrency(context.parsed.y)
            }
          }
        },
        scales: {
          y: {
            ticks: {
              callback: (v) => formatCurrency(v)
            }
          }
        }
      }
    });
  }

  function addScenario() {
    const name = prompt("Enter a name for the new scenario:", "Scenario " + (state.scenarios.length + 1));
    if (!name) return;
    const sc = {
      id: makeId(),
      name: name,
      settings: { ...state.settings },
      lockedBase: false,
      results: null
    };
    sc.results = computeEconomicsForScenario(sc);
    state.scenarios.push(sc);
    renderScenarioView();
    showToast("Scenario added.", "success");
  }

  function resetScenarios() {
    if (!window.confirm("Reset scenarios to a single base scenario?")) return;
    resetScenariosInternal();
    refreshScenarios(state.baseResults || { scenario: null, results: [], control: null });
    renderScenarioView();
  }

  /* Notes */

  function saveNote(key, value) {
    state.notes[key] = value;
    try {
      localStorage.setItem("fabaCBA_notes_" + key, value);
    } catch (e) {
      // ignore
    }
    if (key !== "working") {
      updateConsolidatedNotes();
    }
  }

  function loadNotesFromStorage() {
    ["overview", "results", "benefits", "costs", "simulation", "working"].forEach((key) => {
      try {
        const v = localStorage.getItem("fabaCBA_notes_" + key);
        if (v != null) {
          state.notes[key] = v;
        }
      } catch (e) {
        // ignore
      }
    });
    if (DOM.overviewNotes) DOM.overviewNotes.value = state.notes.overview || "";
    if (DOM.resultsNotes) DOM.resultsNotes.value = state.notes.results || "";
    if (DOM.benefitsNotes) DOM.benefitsNotes.value = state.notes.benefits || "";
    if (DOM.costsNotes) DOM.costsNotes.value = state.notes.costs || "";
    if (DOM.simulationNotes) DOM.simulationNotes.value = state.notes.simulation || "";
    if (DOM.notesWorkingArea) DOM.notesWorkingArea.value = state.notes.working || "";
    updateConsolidatedNotes();
  }

  function updateConsolidatedNotes() {
    if (!DOM.consolidatedNotes) return;
    const sections = [
      { label: "Overview notes", key: "overview" },
      { label: "Results notes", key: "results" },
      { label: "Benefits notes", key: "benefits" },
      { label: "Costs notes", key: "costs" },
      { label: "Simulation notes", key: "simulation" }
    ];
    const parts = [];
    sections.forEach((s) => {
      const text = (state.notes[s.key] || "").trim();
      if (text) {
        parts.push(s.label + ":\n" + text + "\n");
      }
    });
    DOM.consolidatedNotes.value = parts.join("\n");
  }

  /* Exports */

  function exportComparisonCsv() {
    if (!state.baseResults) return;
    const indicators = [
      { key: "pvBenefits", label: "Total benefits over time (present value)" },
      { key: "pvCosts", label: "Total costs over time (present value)" },
      { key: "npv", label: "Net profit over time (present value)" },
      { key: "bcr", label: "Benefit per dollar spent" },
      { key: "roi", label: "Return on investment" },
      { key: "deltaNPV", label: "Difference in net profit compared with control" },
      { key: "deltaCost", label: "Difference in total cost compared with control" },
      { key: "rank", label: "Overall ranking (1 = best net profit)" }
    ];
    const rows = filterResultsForDisplay(state.baseResults);
    if (!rows.length) return;

    const header = ["Indicator"].concat(rows.map((r) => r.treatment + (r.isControl ? " (control)" : "")));
    const lines = [header];

    indicators.forEach((ind) => {
      const line = [ind.label];
      rows.forEach((r) => {
        const val = r[ind.key];
        if (["pvBenefits", "pvCosts", "npv", "deltaNPV", "deltaCost"].includes(ind.key)) {
          line.push(round(val));
        } else if (ind.key === "bcr") {
          line.push(val != null ? round(val, 3) : "");
        } else if (ind.key === "roi") {
          line.push(val != null ? round(val * 100, 2) + "%" : "");
        } else {
          line.push(val != null ? String(val) : "");
        }
      });
      lines.push(line);
    });

    downloadCsv(lines, "comparison_to_control.csv");
    showToast("Comparison table downloaded.", "success");
  }

  function exportLeaderboardCsv() {
    if (!state.baseResults) return;
    const rows = filterResultsForDisplay(state.baseResults);
    const header = [
      "Rank",
      "Treatment",
      "Net profit over time",
      "Difference vs control",
      "Benefit per dollar spent",
      "Return on investment",
      "Included in comparisons"
    ];
    const lines = [header];

    rows.forEach((r) => {
      lines.push([
        r.rank,
        r.treatment + (r.isControl ? " (control)" : ""),
        round(r.npv),
        r.isControl ? "" : round(r.deltaNPV),
        r.bcr != null ? round(r.bcr, 3) : "",
        r.roi != null ? round(r.roi * 100, 2) + "%" : "",
        state.includeTreatments.has(r.treatment) || r.isControl ? "Yes" : "No"
      ]);
    });

    downloadCsv(lines, "leaderboard.csv");
    showToast("Leaderboard downloaded.", "success");
  }

  function exportExcelWorkbook() {
    if (typeof XLSX === "undefined") {
      showToast("Excel export library is not available in this browser. CSV exports remain available.", "error");
      return;
    }
    if (!state.baseResults) return;

    const wb = XLSX.utils.book_new();

    const compRows = [];
    const indicators = [
      { key: "pvBenefits", label: "Total benefits over time (present value)" },
      { key: "pvCosts", label: "Total costs over time (present value)" },
      { key: "npv", label: "Net profit over time (present value)" },
      { key: "bcr", label: "Benefit per dollar spent" },
      { key: "roi", label: "Return on investment" },
      { key: "deltaNPV", label: "Difference in net profit compared with control" },
      { key: "deltaCost", label: "Difference in total cost compared with control" },
      { key: "rank", label: "Overall ranking (1 = best net profit)" }
    ];
    const rows = filterResultsForDisplay(state.baseResults);
    const compHeader = ["Indicator"].concat(rows.map((r) => r.treatment + (r.isControl ? " (control)" : "")));
    compRows.push(compHeader);
    indicators.forEach((ind) => {
      const line = [ind.label];
      rows.forEach((r) => {
        const val = r[ind.key];
        line.push(val);
      });
      compRows.push(line);
    });
    const wsComp = XLSX.utils.aoa_to_sheet(compRows);
    XLSX.utils.book_append_sheet(wb, wsComp, "Comparison");

    const tsHeader = ["Treatment", "Is control?", "Mean yield (t/ha)", "Mean cost per ha"];
    const tsRows = [tsHeader];
    state.treatmentSummaries.forEach((t) => {
      tsRows.push([t.treatment, t.isControl ? "Yes" : "No", t.meanYield, t.meanCost]);
    });
    const wsTreat = XLSX.utils.aoa_to_sheet(tsRows);
    XLSX.utils.book_append_sheet(wb, wsTreat, "Treatments");

    XLSX.writeFile(wb, "faba_beans_cba_results.xlsx");
    showToast("Excel workbook downloaded.", "success");
  }

  function exportCleanDatasetTsv() {
    if (!state.rawRows.length) return;
    const lines = [];
    lines.push(state.headers.join("\t"));
    state.rawRows.forEach((row) => {
      const cells = state.headers.map((h) => {
        const v = row[h];
        return v === null || v === undefined ? "" : String(v);
      });
      lines.push(cells.join("\t"));
    });
    const blob = new Blob([lines.join("\n")], { type: "text/tab-separated-values;charset=utf-8" });
    downloadBlob(blob, "faba_beans_trial_clean_named.tsv");
    showToast("Cleaned dataset downloaded.", "success");
  }

  function exportTreatmentSummaryCsv() {
    const lines = [["Treatment", "Is control?", "Mean yield (t/ha)", "Mean cost per ha"]];
    state.treatmentSummaries.forEach((t) => {
      lines.push([t.treatment, t.isControl ? "Yes" : "No", round(t.meanYield, 4), round(t.meanCost, 4)]);
    });
    downloadCsv(lines, "treatment_summary.csv");
    showToast("Treatment summary downloaded.", "success");
  }

  function exportSensitivityCsv() {
    if (!state.scenarios.length) return;
    const lines = [["Scenario", "Best treatment", "Net profit gain vs control (per ha)"]];
    state.scenarios.forEach((sc) => {
      const control = sc.results.control;
      let best = null;
      let gain = null;
      if (control) {
        best = sc.results.results
          .filter((r) => !r.isControl)
          .slice()
          .sort((a, b) => (b.npv || 0) - (a.npv || 0))[0];
        if (best) gain = best.npv - control.npv;
      }
      lines.push([sc.name, best ? best.treatment : "", gain != null ? round(gain) : ""]);
    });
    downloadCsv(lines, "scenario_results.csv");
    showToast("Scenario results downloaded.", "success");
  }

  /* AI briefing prompts */

  function handleGenerateAiPrompts() {
    if (!state.baseResults || !state.baseResults.control) {
      showToast("Results are not yet available. Please ensure the dataset has loaded and results are visible.", "error");
      return;
    }
    const prompts = buildAiBriefingPrompts();
    if (DOM.aiCopilotPromptArea) {
      DOM.aiCopilotPromptArea.value = prompts.copilotPrompt;
    }
    if (DOM.aiChatGptPromptArea) {
      DOM.aiChatGptPromptArea.value = prompts.chatGptPrompt;
    }
    showToast("AI briefing prompts generated for Copilot and ChatGPT.", "success");
  }

  function buildAiBriefingPrompts() {
    const base = state.baseResults;
    const control = base.control;
    const allResults = base.results.slice().sort((a, b) => a.rank - b.rank);
    const treatmentsForTable = [];

    if (control) {
      treatmentsForTable.push(control);
    }
    allResults.forEach((r) => {
      if (!r.isControl) treatmentsForTable.push(r);
    });

    const maxCols = 5;
    const cols = treatmentsForTable.slice(0, maxCols);

    function plainCurrency(v) {
      if (!isFinite(v)) return "NA";
      const rounded = round(v);
      return String(rounded);
    }

    function plainNumber(v, decimals) {
      if (!isFinite(v)) return "NA";
      return String(round(v, decimals));
    }

    function plainPercent(v) {
      if (!isFinite(v)) return "NA";
      return String(round(v * 100, 2));
    }

    const headerLine =
      "Indicator | " + cols.map((c) => (c.isControl ? c.treatment + " (control)" : c.treatment)).join(" | ");

    function indicatorLine(label, getter, formatter) {
      const cells = cols.map((c) => formatter(getter(c)));
      return label + " | " + cells.join(" | ");
    }

    const tableLines = [];
    tableLines.push(headerLine);
    tableLines.push(
      indicatorLine("Total benefits over time (present value per ha)", (r) => r.pvBenefits, plainCurrency)
    );
    tableLines.push(
      indicatorLine("Total costs over time (present value per ha)", (r) => r.pvCosts, plainCurrency)
    );
    tableLines.push(
      indicatorLine("Net profit over time (present value per ha)", (r) => r.npv, plainCurrency)
    );
    tableLines.push(
      indicatorLine("Benefit per dollar spent", (r) => r.bcr, (v) => plainNumber(v, 3))
    );
    tableLines.push(
      indicatorLine("Return on investment percent", (r) => r.roi, plainPercent)
    );
    tableLines.push(
      indicatorLine("Difference in net profit vs control (per ha)", (r) => r.deltaNPV, plainCurrency)
    );
    tableLines.push(
      indicatorLine("Difference in total cost vs control (per ha)", (r) => r.deltaCost, plainCurrency)
    );
    tableLines.push(
      indicatorLine("Ranking by net profit (1 is best)", (r) => r.rank, (v) => (v == null ? "NA" : String(v)))
    );

    const scenarioLines = [];
    scenarioLines.push("Scenario name | Grain price | Discount rate percent | Time horizon years | Persistence | Best treatment | Gain vs control (net profit per ha)");
    state.scenarios.forEach((sc) => {
      const controlSc = sc.results.control;
      let best = null;
      let gain = null;
      if (controlSc) {
        best = sc.results.results
          .filter((r) => !r.isControl)
          .slice()
          .sort((a, b) => (b.npv || 0) - (a.npv || 0))[0];
        if (best) {
          gain = best.npv - controlSc.npv;
        }
      }
      scenarioLines.push(
        [
          sc.name,
          plainCurrency(sc.settings.grainPrice),
          plainNumber(sc.settings.discountRate, 2),
          plainNumber(sc.settings.horizonYears, 0),
          plainNumber(sc.settings.persistence, 2),
          best ? best.treatment : "NA",
          gain != null ? plainCurrency(gain) : "NA"
        ].join(" | ")
      );
    });

    const controlName = control ? control.treatment : "Control";
    const controlNpv = control ? plainCurrency(control.npv) : "NA";

    const topTreatment = allResults.find((r) => !r.isControl) || null;
    const topName = topTreatment ? topTreatment.treatment : "NA";
    const topNpv = topTreatment ? plainCurrency(topTreatment.npv) : "NA";
    const topDelta = topTreatment ? plainCurrency(topTreatment.deltaNPV) : "NA";

    const overviewNotes = (state.notes.overview || "").trim();
    const resultsNotes = (state.notes.results || "").trim();
    const benefitsNotes = (state.notes.benefits || "").trim();
    const costsNotes = (state.notes.costs || "").trim();
    const simulationNotes = (state.notes.simulation || "").trim();

    const contextNotesParts = [];
    if (overviewNotes) contextNotesParts.push("Overview notes from the user:\n" + overviewNotes);
    if (resultsNotes) contextNotesParts.push("Results notes from the user:\n" + resultsNotes);
    if (benefitsNotes) contextNotesParts.push("Benefits notes from the user:\n" + benefitsNotes);
    if (costsNotes) contextNotesParts.push("Costs notes from the user:\n" + costsNotes);
    if (simulationNotes) contextNotesParts.push("Simulation notes from the user:\n" + simulationNotes);
    const contextNotesBlock = contextNotesParts.join("\n\n");

    const copilotPrompt =
      "You are assisting with a policy style report for a mixed audience of farmers, farm advisers, and government policy staff.\n\n" +
      "Purpose of the report:\n" +
      "Prepare a detailed yet accessible written report of about three to five pages that presents and interprets the results of a faba beans soil and nutrient management trial.\n" +
      "The report must focus on economic performance relative to a control treatment and should be suitable for use in extension documents and policy briefings.\n\n" +
      "Key writing instructions:\n" +
      "Write in clear and plain English and avoid technical economic jargon.\n" +
      "Do not use bullet points. Use paragraphs and short subheadings instead.\n" +
      "Write for three audiences at once: farmers, advisers, and policy staff.\n" +
      "Highlight which treatments improve net profit compared with the control and why.\n" +
      "Include interpretation of uncertainty and sensitivity to assumptions.\n\n" +
      "Structure the report in the following order using headings and subheadings:\n" +
      "1. Executive summary\n" +
      "2. Background and purpose of the trial\n" +
      "3. Data and methods in plain language\n" +
      "4. Main economic results comparing each treatment with the control\n" +
      "5. Scenario and sensitivity analysis\n" +
      "6. Practical implications for farmers and advisers\n" +
      "7. Policy implications for government and extension programs\n" +
      "8. Limitations and cautions\n" +
      "9. Key messages and next steps\n\n" +
      "Present the quantitative results using the tables provided below. Reproduce these tables in the report in a clean and readable way so that the report is self contained.\n\n" +
      "Trial and analysis context (plain text):\n" +
      "The trial compares several soil and nutrient treatments for faba beans against a control treatment that represents current practice on the trial site.\n" +
      "All results are expressed per hectare and reported as present values over time.\n" +
      "The control treatment is named: " +
      controlName +
      ".\n" +
      "The net profit over time for the control is approximately: " +
      controlNpv +
      " currency units per hectare.\n" +
      "The top ranked treatment by net profit is: " +
      topName +
      " with net profit of approximately " +
      topNpv +
      " per hectare and an improvement over the control of about " +
      topDelta +
      " per hectare.\n\n" +
      "Main comparison table of economic indicators by treatment (per hectare, present values):\n" +
      tableLines.join("\n") +
      "\n\n" +
      "Scenario and sensitivity table (how rankings respond to different assumptions):\n" +
      scenarioLines.join("\n") +
      "\n\n" +
      (contextNotesBlock
        ? "Additional contextual notes supplied by the user. Use these to inform interpretation, but do not treat them as data:\n" +
          contextNotesBlock +
          "\n\n"
        : "") +
      "Please perform the following tasks when writing the report:\n" +
      "Explain how cost benefit analysis is used here in plain language that a non specialist can understand.\n" +
      "Describe clearly how each treatment compares with the control in terms of total benefits, total costs, net profit, benefit per dollar spent, and return on investment.\n" +
      "Explain why some treatments may have higher costs but also higher profits, and how this should be interpreted by farmers and advisers.\n" +
      "Discuss the scenario results and explain what they mean for sensitivity to grain prices, discount rates, and persistence of treatment effects.\n" +
      "Provide concrete examples of how a farmer might use these results when choosing treatments on farm.\n" +
      "Provide concrete examples of how a policy maker might use these results when designing incentives or extension programs.\n" +
      "Make sure the report can be read on its own without needing to see the original dataset or this prompt.\n";

    const chatGptPrompt =
      "You are helping to write a detailed policy style report that explains the results of a faba beans soil and nutrient management trial.\n\n" +
      "Audience:\n" +
      "Farmers, farm advisers, and policy or extension staff who are comfortable with numbers but are not specialists in economics.\n\n" +
      "Goal:\n" +
      "Write a three to five page equivalent report in clear English that explains the cost and benefit results of the trial.\n" +
      "Focus on how each treatment compares with the control treatment in terms of total benefits, costs, and net profit over time.\n\n" +
      "Important style requirements:\n" +
      "Avoid bullet points and instead use structured paragraphs and headings.\n" +
      "Use short sections and subheadings similar to a policy brief.\n" +
      "Explain technical ideas such as discounting, net present value, and benefit per dollar spent using plain language.\n\n" +
      "Use the following quantitative information to build the report. You should reproduce the tables in the report so that it is self contained.\n\n" +
      "Main comparison table of economic indicators by treatment (per hectare, present values):\n" +
      tableLines.join("\n") +
      "\n\n" +
      "Scenario and sensitivity table:\n" +
      scenarioLines.join("\n") +
      "\n\n" +
      "Core context and headline findings to emphasise:\n" +
      "The control treatment is named " +
      controlName +
      " and has a net profit over time of about " +
      controlNpv +
      " per hectare.\n" +
      "The highest ranked treatment by net profit is " +
      topName +
      " with a net profit of about " +
      topNpv +
      " per hectare, which is approximately " +
      topDelta +
      " per hectare better than the control.\n\n" +
      (contextNotesBlock
        ? "User supplied notes that may contain relevant interpretation. Use them as context, but focus the report on the quantitative results above:\n" +
          contextNotesBlock +
          "\n\n"
        : "") +
      "Please structure the report with the following sections:\n" +
      "Executive summary that explains the main findings in one or two paragraphs.\n" +
      "Background and description of the trial in plain language.\n" +
      "Data and method description with emphasis on how yields and costs were turned into present value benefits and costs.\n" +
      "Detailed results section comparing each treatment with the control, drawing on the main comparison table.\n" +
      "Scenario and sensitivity section that explains the scenario table and what it implies about robustness of the results.\n" +
      "Practical recommendations for farmers and advisers, focusing on how to interpret differences between treatments.\n" +
      "Policy implications for government and extension programs.\n" +
      "Limitations and uncertainties, including data limitations and assumptions about prices, discount rates, and persistence.\n" +
      "A short concluding section that summarises the most important messages.\n\n" +
      "Make sure the report is clear, coherent, and suitable for inclusion in a decision aid or extension document.\n";

    return {
      copilotPrompt,
      chatGptPrompt
    };
  }

  /* Toast */

  function showToast(message, type = "info") {
    if (!DOM.toastHost) return;
    const div = document.createElement("div");
    div.className = "toast" + (type === "success" ? " toast-success" : type === "error" ? " toast-error" : "");
    const content = document.createElement("div");
    const title = document.createElement("div");
    title.className = "toast-title";
    title.textContent = type === "success" ? "Done" : type === "error" ? "Issue" : "Notice";
    const body = document.createElement("p");
    body.className = "toast-body";
    body.textContent = message;
    content.appendChild(title);
    content.appendChild(body);
    div.appendChild(content);
    DOM.toastHost.appendChild(div);
    DOM.toastRegion.textContent = message;
    setTimeout(() => {
      if (div.parentNode === DOM.toastHost) {
        DOM.toastHost.removeChild(div);
      }
    }, 4000);
  }

  /* Clipboard helpers */

  function copyToClipboard(text, successMessage) {
    if (!text) {
      showToast("Nothing to copy yet. Please generate the prompt first.", "error");
      return;
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard
        .writeText(text)
        .then(() => {
          showToast(successMessage, "success");
        })
        .catch(() => {
          fallbackCopy(text, successMessage);
        });
    } else {
      fallbackCopy(text, successMessage);
    }
  }

  function fallbackCopy(text, successMessage) {
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.left = "-9999px";
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand("copy");
      showToast(successMessage, "success");
    } catch (e) {
      showToast("Copy to clipboard is not available in this browser.", "error");
    }
    ta.remove();
  }

  /* Utilities */

  function average(arr) {
    if (!arr || !arr.length) return 0;
    const sum = arr.reduce((a, b) => a + b, 0);
    return sum / arr.length;
  }

  function clamp(v, min, max) {
    return Math.min(max, Math.max(min, v));
  }

  function formatNumber(v, decimals) {
    if (!isFinite(v)) return "–";
    return v.toFixed(decimals);
  }

  function formatCurrency(v) {
    if (!isFinite(v)) return "–";
    const sign = v < 0 ? "-" : "";
    const abs = Math.abs(v);
    return sign + "$" + abs.toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  }

  function formatPercent(v) {
    if (!isFinite(v)) return "–";
    return (v * 100).toFixed(1) + "%";
  }

  function round(v, decimals = 0) {
    if (!isFinite(v)) return "";
    const f = Math.pow(10, decimals);
    return Math.round(v * f) / f;
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function downloadCsv(rows, filename) {
    const csv = rows
      .map((r) =>
        r
          .map((cell) => {
            const s = cell == null ? "" : String(cell);
            if (/[",\n]/.test(s)) {
              return '"' + s.replace(/"/g, '""') + '"';
            }
            return s;
          })
          .join(",")
      )
      .join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    downloadBlob(blob, filename);
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function makeId() {
    return Math.random().toString(36).slice(2, 10);
  }
})();
