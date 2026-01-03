// script.js
// Farming CBA Decision Tool 2
// Robust, farmer-friendly logic: uses PV inputs from Excel, calculates NPV, BCR, ROI,
// and keeps snapshot, Excel export, and AI helper text in sync.

(function () {
  const state = {
    farmName: "",
    timeHorizon: null,
    discountRate: null,
    currency: "AUD",
    notes: "",
    treatments: []
  };

  const dom = {};

  function init() {
    cacheDom();
    bindEvents();
    initDefaultState();
    renderAll();
  }

  function cacheDom() {
    dom.tabButtons = document.querySelectorAll(".tab-button");
    dom.tabPanels = document.querySelectorAll(".tab-panel");

    // Settings
    dom.farmName = document.getElementById("farmName");
    dom.timeHorizon = document.getElementById("timeHorizon");
    dom.discountRate = document.getElementById("discountRate");
    dom.currency = document.getElementById("currency");
    dom.notes = document.getElementById("notes");

    // Treatments table
    dom.treatmentsTbody = document.getElementById("treatmentsTbody");
    dom.addTreatmentBtn = document.getElementById("addTreatmentBtn");

    // Results
    dom.refreshResultsBtn = document.getElementById("refreshResultsBtn");
    dom.resultsTbody = document.getElementById("resultsTbody");
    dom.headlineSummary = document.getElementById("headlineSummary");

    // Snapshot
    dom.printSnapshotBtn = document.getElementById("printSnapshotBtn");
    dom.snapshotContext = document.getElementById("snapshotContext");
    dom.snapshotHeadline = document.getElementById("snapshotHeadline");
    dom.snapshotTbody = document.getElementById("snapshotTbody");

    // Excel export
    dom.copyTsvBtn = document.getElementById("copyTsvBtn");
    dom.excelExportArea = document.getElementById("excelExportArea");

    // AI helper
    dom.refreshAiHelperBtn = document.getElementById("refreshAiHelperBtn");
    dom.aiHelperArea = document.getElementById("aiHelperArea");
  }

  function bindEvents() {
    // Tabs
    dom.tabButtons.forEach((btn) => {
      btn.addEventListener("click", () => switchTab(btn.dataset.tabTarget));
    });

    // Settings
    dom.farmName.addEventListener("input", () => {
      state.farmName = dom.farmName.value.trim();
      updateSnapshotContext();
      updateAiHelper();
    });

    dom.timeHorizon.addEventListener("input", () => {
      state.timeHorizon = safeNumber(dom.timeHorizon.value);
      updateSnapshotContext();
      updateAiHelper();
    });

    dom.discountRate.addEventListener("input", () => {
      state.discountRate = safeNumber(dom.discountRate.value);
      updateSnapshotContext();
      updateAiHelper();
    });

    dom.currency.addEventListener("input", () => {
      state.currency = dom.currency.value.trim() || "AUD";
      renderAll();
    });

    dom.notes.addEventListener("input", () => {
      state.notes = dom.notes.value.trim();
      updateSnapshotContext();
      updateAiHelper();
    });

    // Treatments
    dom.addTreatmentBtn.addEventListener("click", () => {
      addTreatment("New treatment", false);
      renderTreatmentsTable();
      renderAllResults();
    });

    dom.treatmentsTbody.addEventListener("input", handleTreatmentInput);
    dom.treatmentsTbody.addEventListener("click", handleTreatmentClick);

    // Results
    dom.refreshResultsBtn.addEventListener("click", () => {
      renderAllResults();
    });

    // Snapshot
    dom.printSnapshotBtn.addEventListener("click", () => {
      switchTab("snapshot");
      // Slight delay to ensure tab is visible before print
      setTimeout(() => window.print(), 100);
    });

    // Excel export
    dom.copyTsvBtn.addEventListener("click", copyTsvToClipboard);

    // AI helper
    dom.refreshAiHelperBtn.addEventListener("click", updateAiHelper);
  }

  function initDefaultState() {
    // Reasonable defaults
    state.timeHorizon = 20;
    state.discountRate = 7;
    state.currency = "AUD";

    dom.timeHorizon.value = state.timeHorizon;
    dom.discountRate.value = state.discountRate;
    dom.currency.value = state.currency;

    // Default treatments: one control + two treatments
    addTreatment("Control", true);
    addTreatment("Treatment A", false);
    addTreatment("Treatment B", false);
  }

  function addTreatment(name, isControl) {
    const id = Date.now().toString() + Math.random().toString().slice(2);
    state.treatments.push({
      id,
      name: name || "",
      isControl: !!isControl,
      pvMarket: 0,
      pvSavings: 0,
      pvRisk: 0,
      pvEnv: 0,
      pvCosts: 0
    });
  }

  /* ---------- Event handlers ---------- */

  function handleTreatmentInput(event) {
    const target = event.target;
    const rowIndex = parseInt(target.getAttribute("data-row-index"), 10);
    const field = target.getAttribute("data-field");

    if (Number.isNaN(rowIndex) || !field || !state.treatments[rowIndex]) {
      return;
    }

    const t = state.treatments[rowIndex];

    if (field === "name") {
      t.name = target.value.trim();
    } else if (field === "role") {
      const newRole = target.value;
      if (newRole === "control") {
        // Ensure only one control
        state.treatments.forEach((row, idx) => {
          row.isControl = idx === rowIndex;
        });
      } else {
        t.isControl = false;
        if (!state.treatments.some((row) => row.isControl)) {
          // Keep at least first row as control if user unsets all
          state.treatments[0].isControl = true;
        }
      }
    } else {
      const numeric = safeNumber(target.value);
      if (field === "pvMarket") t.pvMarket = numeric;
      if (field === "pvSavings") t.pvSavings = numeric;
      if (field === "pvRisk") t.pvRisk = numeric;
      if (field === "pvEnv") t.pvEnv = numeric;
      if (field === "pvCosts") t.pvCosts = numeric;
    }

    renderAllResults();
  }

  function handleTreatmentClick(event) {
    const btn = event.target.closest(".btn-remove");
    if (!btn) return;

    const rowIndex = parseInt(btn.getAttribute("data-row-index"), 10);
    if (Number.isNaN(rowIndex) || !state.treatments[rowIndex]) return;

    // Do not allow removing the only control row
    const isControl = state.treatments[rowIndex].isControl;
    if (isControl) {
      const otherControls = state.treatments.filter((t) => t.isControl);
      if (otherControls.length === 1 && state.treatments.length > 1) {
        // Try to move control flag to next row instead of blocking removal
        state.treatments.splice(rowIndex, 1);
        if (state.treatments.length > 0) {
          state.treatments[0].isControl = true;
        }
      } else {
        state.treatments.splice(rowIndex, 1);
      }
    } else {
      state.treatments.splice(rowIndex, 1);
    }

    if (state.treatments.length === 0) {
      addTreatment("Control", true);
    }

    renderTreatmentsTable();
    renderAllResults();
  }

  /* ---------- Tabs ---------- */

  function switchTab(tabName) {
    dom.tabButtons.forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.tabTarget === tabName);
    });

    dom.tabPanels.forEach((panel) => {
      const id = panel.id.replace("tab-", "");
      panel.classList.toggle("active", id === tabName);
    });

    // Keep Excel export and AI prompt current when visiting those tabs
    if (tabName === "excel") {
      updateExcelExport();
    } else if (tabName === "ai-helper") {
      updateAiHelper();
    } else if (tabName === "snapshot") {
      updateSnapshot();
    }
  }

  /* ---------- Rendering ---------- */

  function renderAll() {
    renderTreatmentsTable();
    renderAllResults();
    updateSnapshot();
    updateExcelExport();
    updateAiHelper();
  }

  function renderTreatmentsTable() {
    const rows = state.treatments.map((t, idx) => {
      const roleValue = t.isControl ? "control" : "treatment";
      return `
        <tr>
          <td>
            <select data-row-index="${idx}" data-field="role">
              <option value="control"${roleValue === "control" ? " selected" : ""}>Control</option>
              <option value="treatment"${roleValue === "treatment" ? " selected" : ""}>Treatment</option>
            </select>
          </td>
          <td>
            <input type="text" data-row-index="${idx}" data-field="name"
                   value="${escapeHtml(t.name)}" placeholder="e.g. Precision lime spreading" />
          </td>
          <td>
            <input type="number" step="0.01" data-row-index="${idx}" data-field="pvMarket"
                   value="${formatInputValue(t.pvMarket)}" />
          </td>
          <td>
            <input type="number" step="0.01" data-row-index="${idx}" data-field="pvSavings"
                   value="${formatInputValue(t.pvSavings)}" />
          </td>
          <td>
            <input type="number" step="0.01" data-row-index="${idx}" data-field="pvRisk"
                   value="${formatInputValue(t.pvRisk)}" />
          </td>
          <td>
            <input type="number" step="0.01" data-row-index="${idx}" data-field="pvEnv"
                   value="${formatInputValue(t.pvEnv)}" />
          </td>
          <td>
            <input type="number" step="0.01" data-row-index="${idx}" data-field="pvCosts"
                   value="${formatInputValue(t.pvCosts)}" />
          </td>
          <td>
            <button type="button" class="btn btn-secondary btn-remove" data-row-index="${idx}">
              ×
            </button>
          </td>
        </tr>
      `;
    });

    dom.treatmentsTbody.innerHTML = rows.join("");
  }

  function renderAllResults() {
    const computed = computeAllResults();

    renderHeadlineSummary(computed);
    renderResultsTable(computed);
    updateSnapshot();
    updateExcelExport();
    updateAiHelper();
  }

  function renderHeadlineSummary(computed) {
    const { control, treatments } = computed;
    if (!treatments.length) {
      dom.headlineSummary.innerHTML = "";
      return;
    }

    const currency = state.currency || "AUD";

    const nonControl = treatments.filter((t) => !t.isControl && t.valid);
    const bestByNPV = nonControl.reduce(
      (best, t) => (best === null || t.npv > best.npv ? t : best),
      null
    );
    const bestByROI = nonControl.reduce(
      (best, t) => (best === null || t.roi > best.roi ? t : best),
      null
    );
    const bestByBCR = nonControl.reduce(
      (best, t) => (best === null || t.bcr > best.bcr ? t : best),
      null
    );

    const pills = [];

    if (control && control.valid) {
      pills.push(`
        <div class="summary-pill">
          <span>Control NPV:</span>
          <strong>${formatCurrency(control.npv, currency)}</strong>
        </div>
      `);
    }

    if (bestByNPV) {
      pills.push(`
        <div class="summary-pill">
          <span>Highest NPV (among treatments):</span>
          <strong>${escapeHtml(bestByNPV.name || "Treatment")}</strong>
          <span>${formatCurrency(bestByNPV.npv, currency)}</span>
        </div>
      `);
    }

    if (bestByBCR) {
      pills.push(`
        <div class="summary-pill">
          <span>Highest BCR (among treatments):</span>
          <strong>${escapeHtml(bestByBCR.name || "Treatment")}</strong>
          <span>${formatRatio(bestByBCR.bcr)}</span>
        </div>
      `);
    }

    if (bestByROI) {
      pills.push(`
        <div class="summary-pill">
          <span>Highest ROI (among treatments):</span>
          <strong>${escapeHtml(bestByROI.name || "Treatment")}</strong>
          <span>${formatPercent(bestByROI.roi * 100)}</span>
        </div>
      `);
    }

    dom.headlineSummary.innerHTML = pills.join("");
  }

  function renderResultsTable(computed) {
    const { control, treatments } = computed;
    const currency = state.currency || "AUD";

    const rows = treatments.map((t) => {
      const deltaNPV = control && control.valid ? t.npv - control.npv : null;
      const deltaPVb = control && control.valid ? t.pvBenefitsTotal - control.pvBenefitsTotal : null;
      const deltaPVc = control && control.valid ? t.pvCosts - control.pvCosts : null;

      return `
        <tr>
          <td>${t.isControl ? "Control" : "Treatment"}</td>
          <td>${escapeHtml(t.name || "")}</td>
          <td>${t.valid ? formatCurrency(t.pvBenefitsTotal, currency) : "–"}</td>
          <td>${t.valid ? formatCurrency(t.pvCosts, currency) : "–"}</td>
          <td>${t.valid ? formatCurrency(t.npv, currency) : "–"}</td>
          <td>${t.valid && t.bcr !== null ? formatRatio(t.bcr) : "–"}</td>
          <td>${t.valid && t.roi !== null ? formatPercent(t.roi * 100) : "–"}</td>
          <td>${t.valid && deltaNPV !== null ? formatCurrency(deltaNPV, currency) : "–"}</td>
          <td>${t.valid && deltaPVb !== null ? formatCurrency(deltaPVb, currency) : "–"}</td>
          <td>${t.valid && deltaPVc !== null ? formatCurrency(deltaPVc, currency) : "–"}</td>
        </tr>
      `;
    });

    dom.resultsTbody.innerHTML = rows.join("");
  }

  /* ---------- Snapshot ---------- */

  function updateSnapshotContext() {
    const parts = [];

    if (state.farmName) {
      parts.push(`<p><strong>Farm / project:</strong> ${escapeHtml(state.farmName)}</p>`);
    }

    const horizon = Number.isFinite(state.timeHorizon) && state.timeHorizon > 0
      ? `${state.timeHorizon} year${state.timeHorizon === 1 ? "" : "s"}`
      : null;
    const dr = Number.isFinite(state.discountRate)
      ? `${state.discountRate.toFixed(1)}%`
      : null;

    if (horizon || dr) {
      const bits = [];
      if (horizon) bits.push(`time horizon ${horizon}`);
      if (dr) bits.push(`discount rate ${dr}`);
      parts.push(`<p><strong>Excel settings:</strong> ${escapeHtml(bits.join(", "))}</p>`);
    }

    if (state.notes) {
      parts.push(`<p><strong>Notes:</strong> ${escapeHtml(state.notes)}</p>`);
    }

    dom.snapshotContext.innerHTML = parts.join("") || "<p class=\"muted\">Add farm name, time horizon, discount rate, and notes on the Inputs tab.</p>";
  }

  function updateSnapshot() {
    updateSnapshotContext();

    const computed = computeAllResults();
    const { treatments } = computed;
    const currency = state.currency || "AUD";

    // Headline section
    const nonControl = treatments.filter((t) => !t.isControl && t.valid);
    const control = computed.control;

    let headlineHtml = "";

    if (!treatments.length || !treatments.some((t) => t.valid)) {
      headlineHtml = "<p class=\"muted\">Enter PV benefits and PV costs for at least one treatment on the Inputs tab.</p>";
    } else {
      const bestNPV = nonControl.reduce(
        (best, t) => (best === null || t.npv > best.npv ? t : best),
        null
      );

      if (control && control.valid) {
        headlineHtml += `
          <p>
            The control has an NPV of <strong>${formatCurrency(control.npv, currency)}</strong>,
            with total PV benefits of <strong>${formatCurrency(control.pvBenefitsTotal, currency)}</strong>
            and total PV costs of <strong>${formatCurrency(control.pvCosts, currency)}</strong>.
          </p>
        `;
      }

      if (bestNPV) {
        const delta = control && control.valid ? bestNPV.npv - control.npv : null;
        headlineHtml += `
          <p>
            Among the non-control treatments, <strong>${escapeHtml(bestNPV.name || "one treatment")}</strong>
            currently has the highest NPV at <strong>${formatCurrency(bestNPV.npv, currency)}</strong>.
            ${delta !== null ? `This is ${formatCurrency(delta, currency)} higher than the control under the PV figures entered.` : ""}
          </p>
        `;
      }

      if (!bestNPV && control && control.valid) {
        headlineHtml += `
          <p>
            Only the control currently has complete PV data. Add PV benefits and costs for other treatments to compare them here.
          </p>
        `;
      }
    }

    dom.snapshotHeadline.innerHTML = headlineHtml;

    // Table
    const rows = treatments.map((t) => {
      return `
        <tr>
          <td>${t.isControl ? "Control" : "Treatment"}</td>
          <td>${escapeHtml(t.name || "")}</td>
          <td>${t.valid ? formatCurrency(t.pvBenefitsTotal, currency) : "–"}</td>
          <td>${t.valid ? formatCurrency(t.pvCosts, currency) : "–"}</td>
          <td>${t.valid ? formatCurrency(t.npv, currency) : "–"}</td>
          <td>${t.valid && t.bcr !== null ? formatRatio(t.bcr) : "–"}</td>
          <td>${t.valid && t.roi !== null ? formatPercent(t.roi * 100) : "–"}</td>
        </tr>
      `;
    });

    dom.snapshotTbody.innerHTML = rows.join("");
  }

  /* ---------- Excel export ---------- */

  function updateExcelExport() {
    const computed = computeAllResults();
    const { treatments } = computed;
    const currency = state.currency || "AUD";

    const header = [
      "role",
      "treatment_name",
      "pv_benefits_market",
      "pv_benefits_cost_savings",
      "pv_benefits_risk_resilience",
      "pv_benefits_environment_other",
      "pv_benefits_total",
      "pv_costs_total",
      "npv",
      "bcr",
      "roi",
      "currency"
    ].join("\t");

    const lines = treatments.map((t) => {
      const line = [
        t.isControl ? "control" : "treatment",
        t.name || "",
        numberOrEmpty(t.pvMarket),
        numberOrEmpty(t.pvSavings),
        numberOrEmpty(t.pvRisk),
        numberOrEmpty(t.pvEnv),
        numberOrEmpty(t.pvBenefitsTotal),
        numberOrEmpty(t.pvCosts),
        t.valid ? numberOrEmpty(t.npv) : "",
        t.valid && t.bcr !== null ? numberOrEmpty(t.bcr) : "",
        t.valid && t.roi !== null ? numberOrEmpty(t.roi) : "",
        currency
      ];
      return line.join("\t");
    });

    const tsv = [header].concat(lines).join("\n");
    dom.excelExportArea.value = tsv;
  }

  function copyTsvToClipboard() {
    const text = dom.excelExportArea.value;
    if (!text) return;

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).catch(() => {
        fallbackCopyText(text);
      });
    } else {
      fallbackCopyText(text);
    }
  }

  function fallbackCopyText(text) {
    const temp = document.createElement("textarea");
    temp.value = text;
    temp.style.position = "fixed";
    temp.style.left = "-9999px";
    document.body.appendChild(temp);
    temp.select();
    try {
      document.execCommand("copy");
    } catch (e) {
      // Ignore
    }
    document.body.removeChild(temp);
  }

  /* ---------- AI helper ---------- */

  function updateAiHelper() {
    const computed = computeAllResults();
    const { treatments, control } = computed;
    const currency = state.currency || "AUD";

    const lines = [];

    lines.push("You are interpreting results from a farm cost–benefit analysis tool called \"Farming CBA Decision Tool 2\".");
    lines.push("Use plain language suitable for a farmer or on-farm manager. Avoid jargon. Focus on what drives results and what could be changed.");
    lines.push("Do not tell the farmer which option to choose and do not apply rigid cut-off rules. Treat this as decision support only.");
    lines.push("");

    if (state.farmName) {
      lines.push(`Farm or project name: ${state.farmName}`);
    }
    if (Number.isFinite(state.timeHorizon)) {
      lines.push(`Time horizon used in the Excel PV calculations: ${state.timeHorizon} years.`);
    }
    if (Number.isFinite(state.discountRate)) {
      lines.push(`Discount rate used in the Excel PV calculations: ${state.discountRate.toFixed(1)}% per year.`);
    }
    lines.push(`Currency for all PV figures: ${currency}.`);
    if (state.notes) {
      lines.push(`Notes from the user about assumptions/context: ${state.notes}`);
    }
    lines.push("");

    lines.push("Definitions:");
    lines.push("• PV benefits and PV costs are discounted sums from the Excel workbook.");
    lines.push("• NPV = PV benefits − PV costs. Positive NPV indicates a net economic gain above zero baseline for that treatment.");
    lines.push("• BCR = PV benefits ÷ PV costs. BCR above 1 means PV benefits exceed PV costs.");
    lines.push("• ROI = NPV ÷ PV costs. This is the net gain per dollar of PV cost.");
    lines.push("");

    lines.push("For each option below, report:");
    lines.push("1. PV benefits, PV costs, NPV, BCR, and ROI.");
    lines.push("2. How much higher or lower the NPV, PV benefits, and PV costs are compared with the control.");
    lines.push("3. What seems to drive PV benefits (yield, cost savings, risk/resilience, environmental/other) and PV costs.");
    lines.push("4. Key trade-offs and what might improve each option (e.g. lowering costs, boosting benefits, reducing downside risk).");
    lines.push("");

    treatments.forEach((t) => {
      const label = t.isControl ? "control" : "treatment";
      lines.push(`Option: ${t.name || "(no name)"} [${label}]`);
      lines.push(`  PV benefits (market): ${numberOrEmpty(t.pvMarket)}`);
      lines.push(`  PV benefits (cost savings): ${numberOrEmpty(t.pvSavings)}`);
      lines.push(`  PV benefits (risk/resilience): ${numberOrEmpty(t.pvRisk)}`);
      lines.push(`  PV benefits (environment/other): ${numberOrEmpty(t.pvEnv)}`);
      lines.push(`  PV benefits (total): ${t.valid ? numberOrEmpty(t.pvBenefitsTotal) : ""}`);
      lines.push(`  PV costs (total): ${t.valid ? numberOrEmpty(t.pvCosts) : ""}`);
      lines.push(`  NPV: ${t.valid ? numberOrEmpty(t.npv) : ""}`);
      lines.push(`  BCR: ${t.valid && t.bcr !== null ? numberOrEmpty(t.bcr) : ""}`);
      lines.push(`  ROI: ${t.valid && t.roi !== null ? numberOrEmpty(t.roi) : ""}`);

      if (control && control.valid) {
        const deltaNPV = t.npv - control.npv;
        const deltaPVb = t.pvBenefitsTotal - control.pvBenefitsTotal;
        const deltaPVc = t.pvCosts - control.pvCosts;
        lines.push(`  Difference in NPV versus control: ${t.valid ? numberOrEmpty(deltaNPV) : ""}`);
        lines.push(`  Difference in PV benefits versus control: ${t.valid ? numberOrEmpty(deltaPVb) : ""}`);
        lines.push(`  Difference in PV costs versus control: ${t.valid ? numberOrEmpty(deltaPVc) : ""}`);
      }
      lines.push("");
    });

    lines.push("Write a two to three page narrative (around 1,200–1,800 words).");
    lines.push("Structure it around: (a) overall economic picture, (b) what drives PV benefits and PV costs, (c) comparison between options, (d) practical considerations and uncertainties.");
    lines.push("Use cautious language and explain that results depend on the underlying assumptions and PV inputs.");

    dom.aiHelperArea.value = lines.join("\n");
  }

  /* ---------- Core calculations ---------- */

  function computeAllResults() {
    const treatments = state.treatments.map((t) => {
      const pvBenefitsTotal = [t.pvMarket, t.pvSavings, t.pvRisk, t.pvEnv]
        .map((x) => (Number.isFinite(x) ? x : 0))
        .reduce((a, b) => a + b, 0);
      const pvCosts = Number.isFinite(t.pvCosts) ? t.pvCosts : 0;
      const valid = pvBenefitsTotal !== 0 || pvCosts !== 0;

      const npv = pvBenefitsTotal - pvCosts;
      const bcr = pvCosts > 0 ? pvBenefitsTotal / pvCosts : null;
      const roi = pvCosts > 0 ? npv / pvCosts : null;

      return {
        ...t,
        pvBenefitsTotal,
        pvCosts,
        valid,
        npv,
        bcr,
        roi
      };
    });

    let control = treatments.find((t) => t.isControl);
    if (!control && treatments.length > 0) {
      control = treatments[0];
      control.isControl = true;
    }

    return { treatments, control };
  }

  /* ---------- Helpers ---------- */

  function safeNumber(value) {
    if (value === null || value === undefined) return NaN;
    if (typeof value === "number") return value;
    const n = Number(String(value).replace(/,/g, "").trim());
    return Number.isFinite(n) ? n : NaN;
  }

  function formatCurrency(x, currency) {
    if (!Number.isFinite(x)) return "–";
    const symbol = currency || "AUD";
    const abs = Math.abs(x);
    const decimals = abs >= 100000 ? 0 : 0;
    const formatted = abs.toLocaleString("en-AU", {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals
    });
    const sign = x < 0 ? "-" : "";
    return `${sign}${symbol} ${formatted}`;
  }

  function formatRatio(x) {
    if (!Number.isFinite(x)) return "–";
    return x.toFixed(2);
  }

  function formatPercent(x) {
    if (!Number.isFinite(x)) return "–";
    return `${x.toFixed(1)}%`;
  }

  function formatInputValue(x) {
    if (!Number.isFinite(x) || x === 0) return "";
    return String(x);
  }

  function numberOrEmpty(x) {
    if (!Number.isFinite(x)) return "";
    return String(x);
  }

  function escapeHtml(str) {
    if (str === null || str === undefined) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  // Initialise
  document.addEventListener("DOMContentLoaded", init);
})();
