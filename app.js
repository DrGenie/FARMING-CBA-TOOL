(function () {
  var state = {
    base: {
      areaHa: 100,
      horizonYears: 10,
      discountRatePct: 7,
      defaultPricePerTonne: 450
    },
    treatments: []
  };

  var nextId = 1;
  function newId() {
    var id = "t" + nextId;
    nextId += 1;
    return id;
  }

  function toNumberOrZero(v) {
    var x = Number(v);
    return isFinite(x) ? x : 0;
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function formatMoney(x) {
    if (!isFinite(x)) return "–";
    var sign = x < 0 ? "-" : "";
    var v = Math.abs(x);
    return (
      sign +
      "$" +
      v.toLocaleString(undefined, {
        maximumFractionDigits: v >= 1000 ? 0 : 2,
        minimumFractionDigits: 0
      })
    );
  }

  function createTreatment(name, isControl, yieldTPerHa, pricePerTonne,
                           capitalCostPerHa, costSeedPerHa, costChemPerHa,
                           costLabourPerHa, costMachPerHa, costOtherPerHa,
                           extraBenefitsPerHa) {
    return {
      id: newId(),
      name: name,
      isControl: !!isControl,
      yieldTPerHa: toNumberOrZero(yieldTPerHa),
      pricePerTonne: pricePerTonne === "" ? "" : toNumberOrZero(pricePerTonne),
      capitalCostPerHa: toNumberOrZero(capitalCostPerHa),
      costSeedPerHa: toNumberOrZero(costSeedPerHa),
      costChemPerHa: toNumberOrZero(costChemPerHa),
      costLabourPerHa: toNumberOrZero(costLabourPerHa),
      costMachPerHa: toNumberOrZero(costMachPerHa),
      costOtherPerHa: toNumberOrZero(costOtherPerHa),
      extraBenefitsPerHa: toNumberOrZero(extraBenefitsPerHa)
    };
  }

  /* DOM references */

  var tabButtons = Array.prototype.slice.call(document.querySelectorAll(".tab-button"));
  var tabPanels = Array.prototype.slice.call(document.querySelectorAll(".tab-panel"));

  var inputArea = document.getElementById("input-area");
  var inputHorizon = document.getElementById("input-horizon");
  var inputDiscount = document.getElementById("input-discount");
  var inputPrice = document.getElementById("input-price");

  var inputUpload = document.getElementById("input-upload");
  var uploadStatus = document.getElementById("upload-status");

  var treatmentsTableBody = document.querySelector("#treatments-table tbody");
  var btnAddTreatment = document.getElementById("btn-add-treatment");

  var resultsTableContainer = document.getElementById("results-table-container");
  var summaryBlock = document.getElementById("summary-block");

  var btnRefreshPrompt = document.getElementById("btn-refresh-prompt");
  var btnCopyPrompt = document.getElementById("btn-copy-prompt");
  var btnDownloadPrompt = document.getElementById("btn-download-prompt");
  var aiPromptTextarea = document.getElementById("ai-prompt");

  var btnExportExcel = document.getElementById("btn-export-excel");
  var btnPrint = document.getElementById("btn-print");
  var btnPrint2 = document.getElementById("btn-print-2");

  /* Tabs */

  function activateTab(tabName) {
    for (var i = 0; i < tabButtons.length; i++) {
      var b = tabButtons[i];
      var active = b.getAttribute("data-tab") === tabName;
      if (active) {
        b.classList.add("active");
      } else {
        b.classList.remove("active");
      }
    }
    for (var j = 0; j < tabPanels.length; j++) {
      var p = tabPanels[j];
      var activeP = p.getAttribute("data-tab") === tabName;
      if (activeP) {
        p.classList.add("active");
      } else {
        p.classList.remove("active");
      }
    }
  }

  for (var i = 0; i < tabButtons.length; i++) {
    tabButtons[i].addEventListener("click", function (e) {
      var tabName = e.currentTarget.getAttribute("data-tab");
      activateTab(tabName);
    });
  }

  /* Base settings */

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

  function attachBaseListeners() {
    var baseInputs = [inputArea, inputHorizon, inputDiscount, inputPrice];
    for (var i = 0; i < baseInputs.length; i++) {
      baseInputs[i].addEventListener("change", function () {
        updateBaseFromInputs();
        recomputeAndRender();
      });
    }
  }

  /* Treatments table */

  function renderTreatmentsTable() {
    treatmentsTableBody.innerHTML = "";
    for (var i = 0; i < state.treatments.length; i++) {
      var t = state.treatments[i];
      var tr = document.createElement("tr");
      tr.setAttribute("data-id", t.id);

      var totalAnnual =
        t.costSeedPerHa +
        t.costChemPerHa +
        t.costLabourPerHa +
        t.costMachPerHa +
        t.costOtherPerHa;

      tr.innerHTML =
        '<td>' +
        '<input type="text" data-field="name" value="' + escapeHtml(t.name) + '" />' +
        "</td>" +
        '<td style="text-align:center;">' +
        '<input type="checkbox" data-field="isControl"' + (t.isControl ? " checked" : "") + " />" +
        "</td>" +
        '<td><input type="number" step="0.01" data-field="yieldTPerHa" value="' + t.yieldTPerHa + '" /></td>' +
        '<td><input type="number" step="1" data-field="pricePerTonne" value="' + (t.pricePerTonne === "" ? "" : t.pricePerTonne) + '" /></td>' +
        '<td><input type="number" step="1" data-field="capitalCostPerHa" value="' + t.capitalCostPerHa + '" /></td>' +
        '<td><input type="number" step="1" data-field="costSeedPerHa" value="' + t.costSeedPerHa + '" /></td>' +
        '<td><input type="number" step="1" data-field="costChemPerHa" value="' + t.costChemPerHa + '" /></td>' +
        '<td><input type="number" step="1" data-field="costLabourPerHa" value="' + t.costLabourPerHa + '" /></td>' +
        '<td><input type="number" step="1" data-field="costMachPerHa" value="' + t.costMachPerHa + '" /></td>' +
        '<td><input type="number" step="1" data-field="costOtherPerHa" value="' + t.costOtherPerHa + '" /></td>' +
        '<td class="total-cost-cell">' + formatMoney(totalAnnual) + "</td>" +
        '<td><input type="number" step="1" data-field="extraBenefitsPerHa" value="' + t.extraBenefitsPerHa + '" /></td>' +
        '<td class="actions-cell">' +
        '<button type="button" class="btn btn-ghost btn-small" data-action="remove" data-index="' + i + '">✕</button>' +
        "</td>";

      treatmentsTableBody.appendChild(tr);
    }
  }

  function refreshTotalCostCell(tr, t) {
    var annual =
      t.costSeedPerHa +
      t.costChemPerHa +
      t.costLabourPerHa +
      t.costMachPerHa +
      t.costOtherPerHa;
    var cell = tr.querySelector(".total-cost-cell");
    if (cell) {
      cell.textContent = formatMoney(annual);
    }
  }

  treatmentsTableBody.addEventListener("change", function (e) {
    var target = e.target;
    var tr = target.closest("tr");
    if (!tr) return;
    var id = tr.getAttribute("data-id");
    var field = target.getAttribute("data-field");
    if (!field) return;

    var t = null;
    for (var i = 0; i < state.treatments.length; i++) {
      if (state.treatments[i].id === id) {
        t = state.treatments[i];
        break;
      }
    }
    if (!t) return;

    if (field === "isControl") {
      t.isControl = target.checked;
    } else if (field === "name") {
      t.name = target.value || "";
    } else {
      t[field] = toNumberOrZero(target.value);
    }

    refreshTotalCostCell(tr, t);
    recomputeAndRender();
  });

  treatmentsTableBody.addEventListener("click", function (e) {
    var btn = e.target.closest("button[data-action='remove']");
    if (!btn) return;
    var idx = Number(btn.getAttribute("data-index"));
    if (!isFinite(idx)) return;
    state.treatments.splice(idx, 1);
    if (state.treatments.length === 0) {
      state.treatments.push(
        createTreatment("Control", true, 2.5, "", 0, 150, 80, 60, 90, 20, 0)
      );
    }
    renderTreatmentsTable();
    recomputeAndRender();
  });

  btnAddTreatment.addEventListener("click", function () {
    state.treatments.push(
      createTreatment("New treatment", false, 2.5, "", 0, 150, 80, 60, 90, 20, 0)
    );
    renderTreatmentsTable();
    recomputeAndRender();
  });

  /* Excel upload using actual dataset */

  inputUpload.addEventListener("change", function (e) {
    var file = e.target.files[0];
    if (!file) return;
    uploadStatus.textContent = "Reading file…";
    var reader = new FileReader();
    reader.onload = function (evt) {
      try {
        var data = evt.target.result;
        var wb = XLSX.read(data, { type: "binary" });

        var sheetName = wb.SheetNames[0];
        var ws = wb.Sheets[sheetName];
        var rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

        if (!rows || !rows.length) {
          uploadStatus.textContent = "No data found in the first sheet.";
          return;
        }

        var firstRow = rows[0];
        var keys = Object.keys(firstRow);
        var requiredCols = ["Amendment", "Yield t/ha", "Treatment Input Cost Only /Ha"];
        var missing = [];
        for (var i = 0; i < requiredCols.length; i++) {
          if (keys.indexOf(requiredCols[i]) === -1) {
            missing.push(requiredCols[i]);
          }
        }
        if (missing.length) {
          uploadStatus.textContent =
            "Missing required columns: " + missing.join(", ") +
            ". Please keep your actual data exactly as is; the tool reads these names literally.";
          return;
        }

        var byAmend = {};
        for (var r = 0; r < rows.length; r++) {
          var row = rows[r];
          var name = String(row["Amendment"] || "").trim();
          if (!name) continue;
          var y = toNumberOrZero(row["Yield t/ha"]);
          var c = toNumberOrZero(row["Treatment Input Cost Only /Ha"]);

          if (!byAmend[name]) {
            byAmend[name] = {
              name: name,
              yields: [],
              costs: []
            };
          }
          byAmend[name].yields.push(y);
          byAmend[name].costs.push(c);
        }

        var newTreatments = [];
        var hasControl = false;
        for (var name in byAmend) {
          if (!byAmend.hasOwnProperty(name)) continue;
          var agg = byAmend[name];
          var sumY = 0;
          var sumC = 0;
          for (var k = 0; k < agg.yields.length; k++) sumY += agg.yields[k];
          for (var k2 = 0; k2 < agg.costs.length; k2++) sumC += agg.costs[k2];
          var avgY = agg.yields.length ? sumY / agg.yields.length : 0;
          var avgC = agg.costs.length ? sumC / agg.costs.length : 0;

          var isControl = name.toLowerCase().indexOf("control") !== -1;
          if (isControl) hasControl = true;

          // Put the aggregated cost into "Other" by default; user can re-allocate.
          var t = createTreatment(
            name,
            isControl,
            avgY,
            "",
            0,
            0,
            0,
            0,
            0,
            avgC,
            0
          );
          newTreatments.push(t);
        }

        if (!newTreatments.length) {
          uploadStatus.textContent =
            "The dataset was read but no rows with an Amendment, Yield t/ha and Treatment Input Cost Only /Ha were found.";
          return;
        }

        if (!hasControl) {
          uploadStatus.textContent =
            "Data loaded. No treatment name contains 'Control'; please tick the control row manually on the treatments tab.";
        } else {
          uploadStatus.textContent =
            "Data loaded. Treatments have been grouped by Amendment and averaged across replicates.";
        }

        state.treatments = newTreatments;
        renderTreatmentsTable();
        recomputeAndRender();
        activateTab("treatments");
      } catch (err) {
        console.error(err);
        uploadStatus.textContent = "There was a problem reading this file. Please check that it is a valid Excel file.";
      } finally {
        inputUpload.value = "";
      }
    };
    reader.readAsBinaryString(file);
  });

  /* Economics */

  function computeResults() {
    var areaHa = state.base.areaHa;
    var T = state.base.horizonYears;
    var rPct = state.base.discountRatePct;
    var defaultPrice = state.base.defaultPricePerTonne;

    var r = rPct / 100;
    var annuityFactor = r > 0 ? (1 - Math.pow(1 + r, -T)) / r : T;

    var results = [];
    for (var i = 0; i < state.treatments.length; i++) {
      var t = state.treatments[i];
      var price =
        t.pricePerTonne === "" || t.pricePerTonne === null || t.pricePerTonne === undefined
          ? defaultPrice
          : toNumberOrZero(t.pricePerTonne);

      var annualGrossBenefitPerHa = t.yieldTPerHa * price + t.extraBenefitsPerHa;
      var annualCostPerHa =
        t.costSeedPerHa +
        t.costChemPerHa +
        t.costLabourPerHa +
        t.costMachPerHa +
        t.costOtherPerHa;

      var pvBenefitsPerHa = annualGrossBenefitPerHa * annuityFactor;
      var pvCostsPerHa = annualCostPerHa * annuityFactor + t.capitalCostPerHa;
      var npvPerHa = pvBenefitsPerHa - pvCostsPerHa;
      var bcr = pvCostsPerHa > 0 ? pvBenefitsPerHa / pvCostsPerHa : null;
      var roi = pvCostsPerHa > 0 ? npvPerHa / pvCostsPerHa : null;

      var pvBenefitsFarm = pvBenefitsPerHa * areaHa;
      var pvCostsFarm = pvCostsPerHa * areaHa;
      var npvFarm = npvPerHa * areaHa;

      results.push({
        treatmentId: t.id,
        name: t.name,
        isControl: !!t.isControl,
        pvBenefitsPerHa: pvBenefitsPerHa,
        pvCostsPerHa: pvCostsPerHa,
        npvPerHa: npvPerHa,
        bcr: bcr,
        roi: roi,
        pvBenefitsFarm: pvBenefitsFarm,
        pvCostsFarm: pvCostsFarm,
        npvFarm: npvFarm
      });
    }

    if (!results.length) {
      return { results: [], controlId: null };
    }

    var control = null;
    for (var j = 0; j < results.length; j++) {
      if (results[j].isControl) {
        control = results[j];
        break;
      }
    }
    if (!control) control = results[0];

    for (var k = 0; k < results.length; k++) {
      results[k].diffNpvPerHa = results[k].npvPerHa - control.npvPerHa;
      results[k].diffNpvFarm = results[k].npvFarm - control.npvFarm;
    }

    var sortedByNpv = results.slice().sort(function (a, b) {
      return b.npvPerHa - a.npvPerHa;
    });
    for (var idx = 0; idx < sortedByNpv.length; idx++) {
      sortedByNpv[idx].rank = idx + 1;
    }
    for (var rIdx = 0; rIdx < results.length; rIdx++) {
      var rid = results[rIdx].treatmentId;
      for (var sIdx = 0; sIdx < sortedByNpv.length; sIdx++) {
        if (sortedByNpv[sIdx].treatmentId === rid) {
          results[rIdx].rank = sortedByNpv[sIdx].rank;
          break;
        }
      }
    }

    return { results: results, controlId: control.treatmentId };
  }

  function renderResultsTable(computed) {
    var results = computed.results;
    var controlId = computed.controlId;

    if (!results.length) {
      resultsTableContainer.innerHTML = "<p>No treatments defined.</p>";
      return;
    }

    var controlRes = null;
    for (var i = 0; i < results.length; i++) {
      if (results[i].treatmentId === controlId) {
        controlRes = results[i];
        break;
      }
    }
    if (!controlRes) controlRes = results[0];

    var others = [];
    for (var j = 0; j < results.length; j++) {
      if (results[j].treatmentId !== controlRes.treatmentId) {
        others.push(results[j]);
      }
    }
    others.sort(function (a, b) {
      return a.rank - b.rank;
    });
    var ordered = [controlRes].concat(others);

    function makeRow(labelHtml, key, formatter) {
      var cells = "";
      for (var c = 0; c < ordered.length; c++) {
        var r = ordered[c];
        var val = r[key];
        if (val === null || val === undefined || !isFinite(val)) {
          cells +=
            '<td class="' +
            (r.treatmentId === controlId ? "col-control " : "") +
            '">–</td>';
        } else {
          var cls;
          if (val > 0) cls = "positive";
          else if (val < 0) cls = "negative";
          else cls = "neutral";
          cells +=
            '<td class="' +
            (r.treatmentId === controlId ? "col-control " : "") +
            cls +
            '">' +
            formatter(val) +
            "</td>";
        }
      }
      return "<tr><th>" + labelHtml + "</th>" + cells + "</tr>";
    }

    var header = "";
    for (var h = 0; h < ordered.length; h++) {
      var r = ordered[h];
      header +=
        '<th class="' +
        (r.treatmentId === controlId ? "col-control" : "") +
        '">' +
        escapeHtml(r.name) +
        "</th>";
    }

    var html =
      '<table class="results-table">' +
      "<thead><tr><th>Indicator</th>" +
      header +
      "</tr></thead><tbody>" +
      makeRow(
        '<span class="indicator-label">PV benefits ($/ha)' +
        '<span class="tooltip-icon" data-tooltip="Present value of all benefits per hectare over the time horizon, including yield and extra monetised benefits, discounted at the base rate.">?</span>' +
        "</span>",
        "pvBenefitsPerHa",
        formatMoney
      ) +
      makeRow(
        '<span class="indicator-label">PV costs ($/ha)' +
        '<span class="tooltip-icon" data-tooltip="Present value of all capital and annual costs per hectare over the time horizon, discounted at the base rate.">?</span>' +
        "</span>",
        "pvCostsPerHa",
        formatMoney
      ) +
      makeRow(
        '<span class="indicator-label">NPV ($/ha)' +
        '<span class="tooltip-icon" data-tooltip="Net present value per hectare: PV benefits minus PV costs. Positive values indicate a net economic gain relative to a zero baseline for that treatment.">?</span>' +
        "</span>",
        "npvPerHa",
        formatMoney
      ) +
      makeRow(
        '<span class="indicator-label">NPV difference vs control ($/ha)' +
        '<span class="tooltip-icon" data-tooltip="Difference in NPV per hectare between each treatment and the control. Positive values indicate the treatment outperforms the control economically.">?</span>' +
        "</span>",
        "diffNpvPerHa",
        formatMoney
      ) +
      makeRow(
        '<span class="indicator-label">Benefit–cost ratio (BCR)' +
        '<span class="tooltip-icon" data-tooltip="PV benefits divided by PV costs. Values above 1 mean benefits exceed costs in present value terms.">?</span>' +
        "</span>",
        "bcr",
        function (v) {
          return v.toFixed(2);
        }
      ) +
      makeRow(
        '<span class="indicator-label">Return on investment (ROI)' +
        '<span class="tooltip-icon" data-tooltip="NPV divided by PV costs. Interpretable as net gain per dollar of PV cost.">?</span>' +
        "</span>",
        "roi",
        function (v) {
          return (v * 100).toFixed(1) + "%";
        }
      ) +
      makeRow(
        '<span class="indicator-label">Rank (by NPV/ha)' +
        '<span class="tooltip-icon" data-tooltip="Ranking of treatments based on NPV per hectare, from highest (1) to lowest, with the control included.">?</span>' +
        "</span>",
        "rank",
        function (v) {
          return v;
        }
      ) +
      "</tbody></table>";

    resultsTableContainer.innerHTML = html;
  }

  function renderSummary(computed) {
    var results = computed.results;
    if (!results.length) {
      summaryBlock.textContent = "No treatments defined.";
      return;
    }
    var best = results.slice().sort(function (a, b) {
      return b.npvPerHa - a.npvPerHa;
    })[0];
    var areaHa = state.base.areaHa;
    var bestNpvFarm = best.npvPerHa * areaHa;

    summaryBlock.innerHTML =
      '<div class="summary-badge">Best NPV per hectare: ' +
      escapeHtml(best.name) +
      "</div>" +
      "<p>On the current assumptions (farm area " +
      areaHa.toLocaleString() +
      " ha, time horizon " +
      state.base.horizonYears +
      " years, discount rate " +
      state.base.discountRatePct +
      "%, default grain price $" +
      state.base.defaultPricePerTonne.toLocaleString() +
      "/t), <strong>" +
      escapeHtml(best.name) +
      "</strong> has the highest NPV per hectare among the treatments entered.</p>" +
      "<p>Its NPV is approximately <strong>" +
      formatMoney(best.npvPerHa) +
      "/ha</strong>. Scaled to the whole farm, this corresponds to about <strong>" +
      formatMoney(bestNpvFarm) +
      "</strong> over the analysis period.</p>" +
      "<p>Use the cost components and the AI helper to explore what drives this result, how it compares with the control, and what practical changes might improve under-performing treatments.</p>";
  }

  /* AI prompt */

  function buildAiPrompt(computed) {
    var results = computed.results;
    var controlId = computed.controlId;
    if (!results.length) {
      return "No treatments defined in the current scenario.";
    }

    var control = null;
    for (var i = 0; i < results.length; i++) {
      if (results[i].treatmentId === controlId) {
        control = results[i];
        break;
      }
    }
    if (!control) control = results[0];

    var lines = [];

    lines.push('You are interpreting results from a farm cost–benefit analysis tool called "Farming CBA Decision Tool 2".');
    lines.push("Use plain language suitable for a farmer or on-farm manager. Avoid jargon. Focus on what drives results and what could be changed.");
    lines.push("");
    lines.push("Context and instructions:");
    lines.push("- Treat this as decision support only. Do not tell the farmer what to choose and do not impose rules or thresholds.");
    lines.push("- Explain what net present value (NPV), present value (PV) of benefits and costs, benefit–cost ratio (BCR) and return on investment (ROI) mean in practical terms.");
    lines.push("- Show trade-offs between treatments, especially compared with the control. Explain why some perform better or worse.");
    lines.push("- When a treatment has a low BCR or negative NPV, suggest realistic ways performance could improve (for example reduce costs, increase yield, improve prices or change agronomic practices), framed as options rather than instructions.");
    lines.push("");
    lines.push("Scenario settings:");
    lines.push("- Farm area: " + state.base.areaHa + " hectares");
    lines.push("- Time horizon: " + state.base.horizonYears + " years");
    lines.push("- Discount rate: " + state.base.discountRatePct + "% per year");
    lines.push("- Default grain price where not specified: $" + state.base.defaultPricePerTonne + " per tonne");
    lines.push("");
    lines.push("Definitions to apply consistently:");
    lines.push("- NPV = PV benefits − PV costs. Positive NPV indicates net economic gain relative to a zero baseline for that treatment scenario.");
    lines.push("- PV benefits and PV costs are discounted sums over time using the base discount rate.");
    lines.push("- BCR = PV benefits ÷ PV costs. Values above 1 imply benefits exceed costs in present value terms.");
    lines.push("- ROI = NPV ÷ PV costs. Interpretable as net gain per dollar of PV cost.");
    lines.push("- The control is shown alongside treatments for direct comparison; always relate results back to the control.");
    lines.push("");
    lines.push("Economic results per hectare and whole-farm:");

    var sorted = results.slice().sort(function (a, b) {
      return a.rank - b.rank;
    });
    for (var j = 0; j < sorted.length; j++) {
      var r = sorted[j];
      lines.push("");
      lines.push("Treatment: " + r.name + (r.treatmentId === controlId ? " (CONTROL)" : ""));
      lines.push("- Rank by NPV/ha: " + r.rank);
      lines.push("- PV benefits per ha: " + formatMoney(r.pvBenefitsPerHa));
      lines.push("- PV costs per ha: " + formatMoney(r.pvCostsPerHa));
      lines.push("- NPV per ha: " + formatMoney(r.npvPerHa));
      lines.push("- NPV per ha relative to control: " + formatMoney(r.diffNpvPerHa));
      if (r.bcr !== null && r.bcr !== undefined && isFinite(r.bcr)) {
        lines.push("- BCR: " + r.bcr.toFixed(2));
      } else {
        lines.push("- BCR: not defined (PV costs are zero or missing)");
      }
      if (r.roi !== null && r.roi !== undefined && isFinite(r.roi)) {
        lines.push("- ROI: " + (r.roi * 100).toFixed(1) + "%");
      } else {
        lines.push("- ROI: not defined (PV costs are zero or missing)");
      }
      lines.push("- Whole-farm NPV over the horizon: " + formatMoney(r.npvFarm));
    }

    lines.push("");
    lines.push("Your task is to write a two to three page interpretation (around 1200–1800 words). Focus on:");
    lines.push("1. Explaining in simple terms what drives PV benefits and PV costs for these treatments (yields, prices, capital costs, annual variable costs and any extra benefits).");
    lines.push("2. Comparing each treatment to the control and highlighting where it gains or loses economically and why.");
    lines.push("3. For treatments with weak performance (low BCR or negative NPV), discussing practical improvement options such as reducing particular cost components, improving agronomy to lift yields or targeting better prices.");
    lines.push("4. Emphasising uncertainty and that results depend on the assumptions. Encourage the farmer to test alternative prices, yields and cost structures rather than treating any single treatment as the automatic answer.");
    lines.push("5. Keeping the tone supportive and exploratory. The goal is to help the farmer understand the numbers and think through what they might want to change, not to instruct them.");

    return lines.join("\n");
  }

  function refreshPrompt() {
    var computed = computeResults();
    aiPromptTextarea.value = buildAiPrompt(computed);
  }

  btnRefreshPrompt.addEventListener("click", function () {
    refreshPrompt();
  });

  btnCopyPrompt.addEventListener("click", function () {
    try {
      navigator.clipboard.writeText(aiPromptTextarea.value);
    } catch (e) {
      alert("Copy to clipboard may not be available. Please select and copy manually.");
    }
  });

  btnDownloadPrompt.addEventListener("click", function () {
    var blob = new Blob([aiPromptTextarea.value], { type: "text/plain;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "Farming_CBA_AI_prompt.txt";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  /* Export results to Excel */

  function treatmentsToSheetData() {
    var rows = [];
    rows.push([
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
    ]);
    for (var i = 0; i < state.treatments.length; i++) {
      var t = state.treatments[i];
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
    }
    return rows;
  }

  btnExportExcel.addEventListener("click", function () {
    var wb = XLSX.utils.book_new();

    var scenarioData = [
      ["Scenario", "Value"],
      ["Farm area (ha)", state.base.areaHa],
      ["Time horizon (years)", state.base.horizonYears],
      ["Discount rate (%)", state.base.discountRatePct],
      ["Default grain price ($/t)", state.base.defaultPricePerTonne]
    ];
    var wsScenario = XLSX.utils.aoa_to_sheet(scenarioData);
    XLSX.utils.book_append_sheet(wb, wsScenario, "Scenario");

    var wsTreatments = XLSX.utils.aoa_to_sheet(treatmentsToSheetData());
    XLSX.utils.book_append_sheet(wb, wsTreatments, "Treatments");

    var computed = computeResults();
    var header = [
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
    var rows = [header];
    for (var i = 0; i < computed.results.length; i++) {
      var r = computed.results[i];
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
    }
    var wsResults = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, wsResults, "Results");

    XLSX.writeFile(wb, "Farming_CBA_Results.xlsx");
  });

  /* Print */

  function triggerPrint() {
    window.print();
  }

  btnPrint.addEventListener("click", triggerPrint);
  btnPrint2.addEventListener("click", triggerPrint);

  /* Recompute and render */

  function recomputeAndRender() {
    var computed = computeResults();
    renderResultsTable(computed);
    renderSummary(computed);
    aiPromptTextarea.value = buildAiPrompt(computed);
  }

  function initDefaultsIfEmpty() {
    if (!state.treatments.length) {
      state.treatments = [
        createTreatment("Control (placeholder)", true, 2.5, "", 0, 150, 80, 60, 90, 20, 0),
        createTreatment("Deep OM (placeholder)", false, 3.0, "", 1650, 170, 90, 70, 100, 25, 0)
      ];
    }
  }

  function init() {
    syncBaseInputsFromState();
    attachBaseListeners();
    initDefaultsIfEmpty();
    renderTreatmentsTable();
    recomputeAndRender();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
