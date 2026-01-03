// script.js : Farming CBA Decision Tool 2
document.addEventListener("DOMContentLoaded", () => {
  /* ----------------------- STATE ----------------------- */
  const state = {
    base: {
      areaHa: 100,
      horizonYears: 10,
      discountRatePct: 7,
      defaultPricePerTonne: 450
    },
    treatments: [
      // Sample placeholders – replace via Excel with your full Faba bean dataset
      createTreatment("Control (baseline)", true, 2.5, null, 0, 150, 80, 60, 90, 20, 0),
      createTreatment("Deep OM (CP1)", false, 3.0, null, 1650, 170, 90, 70, 100, 25, 0),
      createTreatment("Deep OM + Gypsum (CP2)", false, 3.3, null, 2400, 180, 95, 70, 105, 30, 0),
      createTreatment("Deep Carbon-coated mineral (CCM)", false, 3.1, null, 3225, 175, 90, 70, 100, 25, 0)
    ]
  };

  function createTreatment(
    name,
    isControl,
    yieldTPerHa,
    pricePerTonne,
    capitalCost,
    cSeed,
    cChem,
    cLab,
    cMach,
    cOther,
    extraBenefits
  ) {
    return {
      id: cryptoRandomId(),
      name,
      isControl,
      yieldTPerHa: toNumberOrZero(yieldTPerHa),
      pricePerTonne:
        pricePerTonne === null || pricePerTonne === undefined
          ? ""
          : toNumberOrZero(pricePerTonne),
      capitalCostPerHa: toNumberOrZero(capitalCost),
      costSeedPerHa: toNumberOrZero(cSeed),
      costChemPerHa: toNumberOrZero(cChem),
      costLabourPerHa: toNumberOrZero(cLab),
      costMachPerHa: toNumberOrZero(cMach),
      costOtherPerHa: toNumberOrZero(cOther),
      extraBenefitsPerHa: toNumberOrZero(extraBenefits)
    };
  }

  function cryptoRandomId() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    return "t_" + Math.random().toString(36).slice(2);
  }

  function toNumberOrZero(v) {
    const x = Number(v);
    return Number.isFinite(x) ? x : 0;
  }

  /* ----------------------- DOM ELEMENTS ----------------------- */
  const tabButtons = Array.from(document.querySelectorAll(".tab-button"));
  const tabPanels = Array.from(document.querySelectorAll(".tab-panel"));

  const inputArea = document.getElementById("input-area");
  const inputHorizon = document.getElementById("input-horizon");
  const inputDiscount = document.getElementById("input-discount");
  const inputPrice = document.getElementById("input-price");

  const treatmentsTableBody = document.querySelector("#treatments-table tbody");
  const resultsTableContainer = document.getElementById("results-table-container");
  const summaryBlock = document.getElementById("summary-block");

  const btnDownloadTemplate = document.getElementById("btn-download-template");
  const btnDownloadScenario = document.getElementById("btn-download-scenario");
  const inputUpload = document.getElementById("input-upload");
  const btnAddTreatment = document.getElementById("btn-add-treatment");
  const btnExportExcel = document.getElementById("btn-export-excel");
  const btnPrintPdf = document.getElementById("btn-print-pdf");
  const btnPrintPdf2 = document.getElementById("btn-print-pdf-2");

  const btnRefreshPrompt = document.getElementById("btn-refresh-prompt");
  const btnCopyPrompt = document.getElementById("btn-copy-prompt");
  const btnDownloadPrompt = document.getElementById("btn-download-prompt");
  const aiPromptTextarea = document.getElementById("ai-prompt");

  /* ----------------------- TAB HANDLING ----------------------- */
  tabButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const tab = btn.dataset.tab;
      tabButtons.forEach((b) => b.classList.toggle("active", b === btn));
      tabPanels.forEach((p) => p.classList.toggle("active", p.dataset.tab === tab));
    });
  });

  /* ----------------------- BASE SETTINGS ----------------------- */
  function syncBaseInputsFromState() {
    inputArea.value = state.base.areaHa;
    inputHorizon.value = state.base.horizonYears;
    inputDiscount.value = state.base.discountRatePct;
    inputPrice.value = state.base.defaultPricePerTonne;
  }

  function updateBaseFromInputs() {
    state.base.areaHa = Math.max(1, toNumberOrZero(inputArea.value));
    state.base.horizonYears = Math.max(1, toNumberOrZero(inputHorizon.value));
    state.base.discountRatePct = Math.max(0, toNumberOrZero(inputDiscount.value));
    state.base.defaultPricePerTonne = Math.max(0, toNumberOrZero(inputPrice.value));
  }

  [inputArea, inputHorizon, inputDiscount, inputPrice].forEach((el) => {
    el.addEventListener("change", () => {
      updateBaseFromInputs();
      recomputeAndRender();
    });
  });

  syncBaseInputsFromState();

  /* ----------------------- TREATMENTS TABLE ----------------------- */

  function renderTreatmentsTable() {
    treatmentsTableBody.innerHTML = "";
    state.treatments.forEach((t, index) => {
      const tr = document.createElement("tr");
      tr.dataset.id = t.id;

      tr.innerHTML = `
        <td>
          <input type="text" data-field="name" value="${escapeHtml(t.name)}" />
        </td>
        <td style="text-align:center;">
          <input type="checkbox" data-field="isControl" ${t.isControl ? "checked" : ""} />
        </td>
        <td>
          <input type="number" step="0.01" data-field="yieldTPerHa" value="${t.yieldTPerHa}" />
        </td>
        <td>
          <input type="number" step="1" data-field="pricePerTonne" value="${t.pricePerTonne}" />
        </td>
        <td>
          <input type="number" step="1" data-field="capitalCostPerHa" value="${t.capitalCostPerHa}" />
        </td>
        <td>
          <input type="number" step="1" data-field="costSeedPerHa" value="${t.costSeedPerHa}" />
        </td>
        <td>
          <input type="number" step="1" data-field="costChemPerHa" value="${t.costChemPerHa}" />
        </td>
        <td>
          <input type="number" step="1" data-field="costLabourPerHa" value="${t.costLabourPerHa}" />
        </td>
        <td>
          <input type="number" step="1" data-field="costMachPerHa" value="${t.costMachPerHa}" />
        </td>
        <td>
          <input type="number" step="1" data-field="costOtherPerHa" value="${t.costOtherPerHa}" />
        </td>
        <td class="total-cost-cell"></td>
        <td>
          <input type="number" step="1" data-field="extraBenefitsPerHa" value="${t.extraBenefitsPerHa}" />
        </td>
        <td class="actions-cell">
          <button type="button" class="btn btn-ghost btn-small" data-action="remove" data-index="${index}">
            ✕
          </button>
        </td>
      `;
      treatmentsTableBody.appendChild(tr);
    });

    updateTotalCostCells();
  }

  function updateTotalCostCells() {
    Array.from(treatmentsTableBody.querySelectorAll("tr")).forEach((tr) => {
      const id = tr.dataset.id;
      const t = state.treatments.find((x) => x.id === id);
      if (!t) return;
      const total =
        t.costSeedPerHa +
        t.costChemPerHa +
        t.costLabourPerHa +
        t.costMachPerHa +
        t.costOtherPerHa;
      const cell = tr.querySelector(".total-cost-cell");
      cell.textContent = formatMoney(total);
    });
  }

  treatmentsTableBody.addEventListener("change", (e) => {
    const target = e.target;
    const tr = target.closest("tr");
    if (!tr) return;
    const id = tr.dataset.id;
    const t = state.treatments.find((x) => x.id === id);
    if (!t) return;

    const field = target.dataset.field;
    if (!field) return;

    if (field === "isControl") {
      // allow multiple controls in Excel, but here we keep one "main" for clarity
      t.isControl = target.checked;
      if (t.isControl) {
        // ensure at least one control; do not uncheck others automatically
      }
    } else if (field === "name") {
      t.name = target.value || "";
    } else {
      t[field] = toNumberOrZero(target.value);
    }

    updateTotalCostCells();
    recomputeAndRender();
  });

  treatmentsTableBody.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-action='remove']");
    if (!btn) return;
    const index = Number(btn.dataset.index);
    if (Number.isFinite(index)) {
      state.treatments.splice(index, 1);
      if (state.treatments.length === 0) {
        state.treatments.push(
          createTreatment("Control", true, 2.5, null, 0, 150, 80, 60, 90, 20, 0)
        );
      }
      renderTreatmentsTable();
      recomputeAndRender();
    }
  });

  btnAddTreatment.addEventListener("click", () => {
    state.treatments.push(
      createTreatment("New treatment", false, 2.5, "", 0, 150, 80, 60, 90, 20, 0)
    );
    renderTreatmentsTable();
    recomputeAndRender();
  });

  renderTreatmentsTable();

  /* ----------------------- ECONOMICS ----------------------- */

  function computeResults() {
    const areaHa = state.base.areaHa;
    const T = state.base.horizonYears;
    const rPct = state.base.discountRatePct;
    const defaultPrice = state.base.defaultPricePerTonne;

    const r = rPct / 100;
    const annuityFactor = r > 0 ? (1 - Math.pow(1 + r, -T)) / r : T;

    const results = state.treatments.map((t) => {
      const price = t.pricePerTonne === "" ? defaultPrice : toNumberOrZero(t.pricePerTonne);
      const annualGrossBenefitPerHa = t.yieldTPerHa * price + t.extraBenefitsPerHa;
      const annualCostPerHa =
        t.costSeedPerHa +
        t.costChemPerHa +
        t.costLabourPerHa +
        t.costMachPerHa +
        t.costOtherPerHa;

      const pvBenefitsPerHa = annualGrossBenefitPerHa * annuityFactor;
      const pvCostsPerHa = annualCostPerHa * annuityFactor + t.capitalCostPerHa;
      const npvPerHa = pvBenefitsPerHa - pvCostsPerHa;
      const bcr = pvCostsPerHa > 0 ? pvBenefitsPerHa / pvCostsPerHa : null;
      const roi = pvCostsPerHa > 0 ? npvPerHa / pvCostsPerHa : null;

      const pvBenefitsFarm = pvBenefitsPerHa * areaHa;
      const pvCostsFarm = pvCostsPerHa * areaHa;
      const npvFarm = npvPerHa * areaHa;

      return {
        treatmentId: t.id,
        name: t.name,
        isControl: !!t.isControl,
        pvBenefitsPerHa,
        pvCostsPerHa,
        npvPerHa,
        bcr,
        roi,
        pvBenefitsFarm,
        pvCostsFarm,
        npvFarm
      };
    });

    // Identify control
    let control = results.find((r) => r.isControl) || results[0];
    // Differences vs control
    results.forEach((r) => {
      r.diffNpvPerHa = r.npvPerHa - control.npvPerHa;
      r.diffNpvFarm = r.npvFarm - control.npvFarm;
    });

    // Ranking by NPV per ha (descending)
    const sortedByNpv = [...results].sort((a, b) => b.npvPerHa - a.npvPerHa);
    sortedByNpv.forEach((r, i) => {
      r.rank = i + 1;
    });

    // Propagate ranks back to original
    results.forEach((r) => {
      const ranked = sortedByNpv.find((x) => x.treatmentId === r.treatmentId);
      r.rank = ranked ? ranked.rank : null;
    });

    return { results, controlId: control.treatmentId };
  }

  /* ----------------------- RESULTS RENDERING ----------------------- */

  function renderResultsTable(computed) {
    const { results, controlId } = computed;

    if (!results.length) {
      resultsTableContainer.innerHTML = "<p>No treatments defined.</p>";
      return;
    }

    const controlResults = results.find((r) => r.treatmentId === controlId) || results[0];
    const others = results.filter((r) => r.treatmentId !== controlId);
    const othersSorted = [...others].sort((a, b) => a.rank - b.rank);
    const ordered = [controlResults, ...othersSorted];

    const headerCols = ordered
      .map(
        (r) =>
          `<th class="${r.treatmentId === controlId ? "col-control" : ""}">
             ${escapeHtml(r.name)}
           </th>`
      )
      .join("");

    function metricRow(labelHtml, key, formatter, isDiff = false) {
      const cells = ordered
        .map((r) => {
          let val;
          if (isDiff) {
            val = r[key];
          } else {
            val = r[key];
          }
          if (val === null || val === undefined || Number.isNaN(val)) {
            return `<td class="${r.treatmentId === controlId ? "col-control" : ""}">–</td>`;
          }
          const formatted = formatter(val);
          const cls =
            val > 0 ? "positive" : val < 0 ? "negative" : "neutral";
          return `<td class="${r.treatmentId === controlId ? "col-control " : ""}${cls}">${formatted}</td>`;
        })
        .join("");
      return `<tr>
        <th>${labelHtml}</th>
        ${cells}
      </tr>`;
    }

    const tableHtml = `
      <table class="results-table">
        <thead>
          <tr>
            <th>Indicator</th>
            ${headerCols}
          </tr>
        </thead>
        <tbody>
          ${metricRow(
            `<span class="indicator-label">
              PV benefits ($/ha)
              <span class="tooltip-icon" data-tooltip="Present value of all benefits per hectare over the time horizon, including yield plus any extra monetised benefits, discounted at the base rate.">?</span>
            </span>`,
            "pvBenefitsPerHa",
            formatMoney
          )}
          ${metricRow(
            `<span class="indicator-label">
              PV costs ($/ha)
              <span class="tooltip-icon" data-tooltip="Present value of all capital and annual costs per hectare over the time horizon, discounted at the base rate.">?</span>
            </span>`,
            "pvCostsPerHa",
            formatMoney
          )}
          ${metricRow(
            `<span class="indicator-label">
              NPV ($/ha)
              <span class="tooltip-icon" data-tooltip="Net present value per hectare: PV benefits minus PV costs. Positive values indicate a net economic gain relative to a zero baseline for that treatment.">?</span>
            </span>`,
            "npvPerHa",
            formatMoney
          )}
          ${metricRow(
            `<span class="indicator-label">
              NPV difference vs control ($/ha)
              <span class="tooltip-icon" data-tooltip="Difference in NPV per hectare between each treatment and the control. Positive values indicate the treatment outperforms the control economically.">?</span>
            </span>`,
            "diffNpvPerHa",
            formatMoney,
            true
          )}
          ${metricRow(
            `<span class="indicator-label">
              Benefit–cost ratio (BCR)
              <span class="tooltip-icon" data-tooltip="PV benefits divided by PV costs. Values above 1 mean benefits exceed costs in present value terms.">?</span>
            </span>`,
            "bcr",
            (v) => v.toFixed(2)
          )}
          ${metricRow(
            `<span class="indicator-label">
              Return on investment (ROI)
              <span class="tooltip-icon" data-tooltip="NPV divided by PV costs. Interpretable as net gain per dollar of PV cost.">?</span>
            </span>`,
            "roi",
            (v) => (v * 100).toFixed(1) + "%"
          )}
          ${metricRow(
            `<span class="indicator-label">
              Rank (by NPV/ha)
              <span class="tooltip-icon" data-tooltip="Ranking of treatments based on NPV per hectare, from highest (1) to lowest. Control is included in the ranking but kept in the first column for comparison.">?</span>
            </span>`,
            "rank",
            (v) => v
          )}
        </tbody>
      </table>
    `;

    resultsTableContainer.innerHTML = tableHtml;
  }

  function renderSummary(computed) {
    const { results } = computed;
    if (!results.length) {
      summaryBlock.textContent = "No treatments defined.";
      return;
    }
    const best = [...results].sort((a, b) => b.npvPerHa - a.npvPerHa)[0];
    const areaHa = state.base.areaHa;
    const bestNpvFarm = best.npvPerHa * areaHa;

    summaryBlock.innerHTML = `
      <div class="summary-badge">
        Best NPV per hectare: ${escapeHtml(best.name)}
      </div>
      <p>
        On the current assumptions (farm area ${areaHa.toLocaleString()} ha,
        time horizon ${state.base.horizonYears} years, discount rate
        ${state.base.discountRatePct}%, and default grain price
        $${state.base.defaultPricePerTonne.toLocaleString()}/t),
        <strong>${escapeHtml(best.name)}</strong> has the highest net present value
        per hectare among the treatments entered.
      </p>
      <p>
        Its NPV is approximately <strong>${formatMoney(best.npvPerHa)}/ha</strong>.
        Scaled up to the whole farm, this corresponds to an NPV of around
        <strong>${formatMoney(bestNpvFarm)}</strong> over the analysis period.
      </p>
      <p>
        Use the cost components and the AI helper to explore what drives this result,
        how it compares with the control, and what practical changes might improve
        under-performing treatments.
      </p>
    `;
  }

  /* ----------------------- AI PROMPT ----------------------- */

  function buildAiPrompt(computed) {
    const { results, controlId } = computed;
    if (!results.length) return "No treatments defined in the current scenario.";

    const control = results.find((r) => r.treatmentId === controlId) || results[0];

    const lines = [];

    lines.push(
      "You are interpreting results from a farm cost–benefit analysis tool called \"Farming CBA Decision Tool 2\"."
    );
    lines.push(
      "Use plain language suitable for a farmer or on-farm manager. Avoid jargon. Focus on what drives results and what could be changed."
    );
    lines.push("");
    lines.push("Context and instructions:");
    lines.push(
      "- Treat this as decision support only. Do not tell the farmer what to choose and do not impose rules or thresholds."
    );
    lines.push(
      "- Explain what net present value (NPV), present value (PV) of benefits and costs, benefit–cost ratio (BCR) and return on investment (ROI) mean in practical terms."
    );
    lines.push(
      "- Show trade-offs between treatments, especially compared with the control. Explain why some perform better or worse."
    );
    lines.push(
      "- When a treatment has a low BCR or negative NPV, suggest realistic ways performance could improve (e.g. reduce costs, increase yield, improve prices, change agronomic practices), framed as options for reflection rather than instructions."
    );
    lines.push("");

    lines.push("Scenario settings:");
    lines.push(`- Farm area: ${state.base.areaHa} hectares`);
    lines.push(`- Time horizon: ${state.base.horizonYears} years`);
    lines.push(`- Discount rate: ${state.base.discountRatePct}% per year`);
    lines.push(
      `- Default grain price used where not specified: $${state.base.defaultPricePerTonne}/tonne`
    );
    lines.push("");

    lines.push("Definitions for you to use consistently:");
    lines.push(
      "- NPV = PV benefits − PV costs. Positive NPV indicates net economic gain relative to a zero baseline for that treatment scenario."
    );
    lines.push(
      "- PV benefits and PV costs are discounted sums over time using the base discount rate."
    );
    lines.push(
      "- BCR = PV benefits ÷ PV costs. Values above 1 imply benefits exceed costs in present value terms."
    );
    lines.push(
      "- ROI = NPV ÷ PV costs. Interpretable as net gain per dollar of PV cost."
    );
    lines.push(
      "- The control is shown alongside treatments for direct comparison; always relate results back to the control."
    );
    lines.push("");

    lines.push("Economic results per hectare (and whole-farm, given the farm area):");

    results
      .slice()
      .sort((a, b) => a.rank - b.rank)
      .forEach((r) => {
        lines.push("");
        lines.push(
          `Treatment: ${r.name}${r.treatmentId === controlId ? " (CONTROL)" : ""}`
        );
        lines.push(`- Rank by NPV/ha: ${r.rank}`);
        lines.push(`- PV benefits per ha: ${formatMoney(r.pvBenefitsPerHa)}`);
        lines.push(`- PV costs per ha: ${formatMoney(r.pvCostsPerHa)}`);
        lines.push(`- NPV per ha: ${formatMoney(r.npvPerHa)}`);
        lines.push(
          `- NPV per ha relative to control: ${formatMoney(r.diffNpvPerHa)}`
        );
        if (r.bcr != null) {
          lines.push(`- BCR: ${r.bcr.toFixed(2)}`);
        } else {
          lines.push("- BCR: not defined (PV costs are zero or missing)");
        }
        if (r.roi != null) {
          lines.push(`- ROI: ${(r.roi * 100).toFixed(1)}%`);
        } else {
          lines.push("- ROI: not defined (PV costs are zero or missing)");
        }
        lines.push(
          `- Whole-farm NPV over the horizon (given area): ${formatMoney(
            r.npvFarm
          )}`
        );
      });

    lines.push("");
    lines.push(
      "Your task: write a two to three page interpretation (around 1200–1800 words)."
    );
    lines.push("Focus on the following:");
    lines.push(
      "1. Explain, in simple terms, what drives PV benefits and PV costs for these treatments (yield, price, capital costs, annual variable costs, and any extra benefits)."
    );
    lines.push(
      "2. Compare each treatment to the control, highlighting where the treatment gains or loses economically and why."
    );
    lines.push(
      "3. For treatments with weak performance (low BCR or negative NPV), discuss practical improvement options such as reducing certain cost components, improving agronomy to lift yields, or targeting better prices."
    );
    lines.push(
      "4. Emphasise uncertainty and that results depend on the assumptions. Encourage the farmer to test alternative prices, yields and cost structures rather than presenting any single treatment as the automatic answer."
    );
    lines.push(
      "5. Keep the tone supportive and exploratory. The goal is to help the farmer understand the numbers and think through what they might want to change, not to tell them what to do."
    );

    return lines.join("\n");
  }

  function refreshPrompt() {
    const computed = computeResults();
    const prompt = buildAiPrompt(computed);
    aiPromptTextarea.value = prompt;
  }

  btnRefreshPrompt.addEventListener("click", refreshPrompt);

  btnCopyPrompt.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(aiPromptTextarea.value);
      alert("Prompt copied to clipboard.");
    } catch {
      alert("Copy failed. Please select and copy the text manually.");
    }
  });

  btnDownloadPrompt.addEventListener("click", () => {
    downloadTextFile(aiPromptTextarea.value, "Farming_CBA_AI_prompt.txt");
  });

  /* ----------------------- EXCEL HANDLERS ----------------------- */

  const TEMPLATE_HEADERS = [
    "TreatmentName",
    "IsControl",
    "Yield_t_per_ha",
    "Price_per_tonne",
    "CapitalCost_Year0_per_ha",
    "Cost_SeedFert_per_ha",
    "Cost_Chemicals_per_ha",
    "Cost_Labour_per_ha",
    "Cost_Machinery_per_ha",
    "Cost_Other_per_ha",
    "ExtraBenefits_per_ha"
  ];

  function treatmentsToSheetData() {
    const rows = [TEMPLATE_HEADERS];
    state.treatments.forEach((t) => {
      rows.push([
        t.name,
        t.isControl ? "yes" : "no",
        t.yieldTPerHa,
        t.pricePerTonne === "" ? "" : t.pricePerTonne,
        t.capitalCostPerHa,
        t.costSeedPerHa,
        t.costChemPerHa,
        t.costLabourPerHa,
        t.costMachPerHa,
        t.costOtherPerHa,
        t.extraBenefitsPerHa
      ]);
    });
    return rows;
  }

  btnDownloadTemplate.addEventListener("click", () => {
    const wb = XLSX.utils.book_new();
    const rows = [TEMPLATE_HEADERS];
    // leave empty rows; user pastes full Faba bean table mapped to these columns
    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, "Treatments");
    XLSX.writeFile(wb, "Farming_CBA_Treatments_Template.xlsx");
  });

  btnDownloadScenario.addEventListener("click", () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(treatmentsToSheetData());
    XLSX.utils.book_append_sheet(wb, ws, "Treatments");
    XLSX.writeFile(wb, "Farming_CBA_Treatments_CurrentScenario.xlsx");
  });

  inputUpload.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target.result;
        const wb = XLSX.read(data, { type: "binary" });
        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

        if (!json.length) {
          alert("No data found in the first sheet.");
          return;
        }

        const headers = Object.keys(json[0]);
        const missing = TEMPLATE_HEADERS.filter((h) => !headers.includes(h));
        if (missing.length) {
          alert(
            "The uploaded file is missing some required columns:\n" +
              missing.join(", ") +
              "\n\nPlease use the template headers exactly."
          );
          return;
        }

        const newTreatments = json.map((row) =>
          createTreatment(
            String(row.TreatmentName || "").trim() || "Untitled treatment",
            String(row.IsControl || "")
              .toLowerCase()
              .startsWith("y"),
            row.Yield_t_per_ha,
            row.Price_per_tonne === "" ? "" : row.Price_per_tonne,
            row.CapitalCost_Year0_per_ha,
            row.Cost_SeedFert_per_ha,
            row.Cost_Chemicals_per_ha,
            row.Cost_Labour_per_ha,
            row.Cost_Machinery_per_ha,
            row.Cost_Other_per_ha,
            row.ExtraBenefits_per_ha
          )
        );

        if (!newTreatments.length) {
          alert("No valid treatment rows found.");
          return;
        }

        state.treatments = newTreatments;
        renderTreatmentsTable();
        recomputeAndRender();
        alert(
          "Treatments imported successfully. All rows in the sheet are now included in the comparison."
        );
      } catch (err) {
        console.error(err);
        alert("There was a problem reading this file. Please check the format.");
      } finally {
        inputUpload.value = "";
      }
    };
    reader.readAsBinaryString(file);
  });

  btnExportExcel.addEventListener("click", () => {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Scenario
    const scenarioData = [
      ["Scenario", "Value"],
      ["Farm area (ha)", state.base.areaHa],
      ["Time horizon (years)", state.base.horizonYears],
      ["Discount rate (%)", state.base.discountRatePct],
      ["Default grain price ($/t)", state.base.defaultPricePerTonne]
    ];
    const wsScenario = XLSX.utils.aoa_to_sheet(scenarioData);
    XLSX.utils.book_append_sheet(wb, wsScenario, "Scenario");

    // Sheet 2: Treatments (inputs)
    const wsTreatments = XLSX.utils.aoa_to_sheet(treatmentsToSheetData());
    XLSX.utils.book_append_sheet(wb, wsTreatments, "Treatments");

    // Sheet 3: Results
    const computed = computeResults();
    const header = [
      "TreatmentName",
      "IsControl",
      "Rank_by_NPV_per_ha",
      "PV_Benefits_per_ha",
      "PV_Costs_per_ha",
      "NPV_per_ha",
      "NPV_diff_vs_control_per_ha",
      "BCR",
      "ROI",
      "NPV_whole_farm"
    ];
    const rows = [header];
    computed.results.forEach((r) => {
      rows.push([
        r.name,
        r.isControl ? "yes" : "no",
        r.rank,
        r.pvBenefitsPerHa,
        r.pvCostsPerHa,
        r.npvPerHa,
        r.diffNpvPerHa,
        r.bcr,
        r.roi,
        r.npvFarm
      ]);
    });
    const wsResults = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, wsResults, "Results");

    XLSX.writeFile(wb, "Farming_CBA_Results.xlsx");
  });

  /* ----------------------- PRINT / PDF ----------------------- */

  function triggerPrint() {
    window.print();
  }

  btnPrintPdf.addEventListener("click", triggerPrint);
  btnPrintPdf2.addEventListener("click", triggerPrint);

  /* ----------------------- UTILITIES ----------------------- */

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function formatMoney(x) {
    if (!Number.isFinite(x)) return "–";
    const sign = x < 0 ? "-" : "";
    const v = Math.abs(x);
    return (
      sign +
      "$" +
      v.toLocaleString(undefined, {
        maximumFractionDigits: v >= 1000 ? 0 : 2,
        minimumFractionDigits: 0
      })
    );
  }

  function downloadTextFile(text, filename) {
    const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function recomputeAndRender() {
    const computed = computeResults();
    renderResultsTable(computed);
    renderSummary(computed);
    // Keep prompt reasonably up to date but allow manual refresh as well
    aiPromptTextarea.value = buildAiPrompt(computed);
  }

  // Initial render
  recomputeAndRender();
});
