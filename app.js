// Farming CBA Tool — Newcastle Business School (with Cover tab and farm-styled UI)
(() => {
  // ---------- MODEL ----------
  const model = {
    project: {
      name: "Nitrogen Optimization Trial",
      analysts: "Farm Econ Team",
      organisation: "Newcastle Business School, The University of Newcastle",
      contactEmail: "frank.agbola@newcastle.edu.au",
      summary:
        "Test fertilizer strategies to raise wheat yield and protein across 500 ha over 5 years.",
      lastUpdated: new Date().toISOString().slice(0, 10),
      goal: "Increase yield by 10% and protein by 0.5 p.p. on 500 ha within 3 years.",
      withProject: "Adopt optimized nitrogen timing and rates; improved management on 500 ha.",
      withoutProject:
        "Business-as-usual fertilization; yield/protein unchanged; rising costs."
    },
    time: {
      startYear: new Date().getFullYear(),
      years: 10,
      discBase: 7,
      discLow: 4,
      discHigh: 10,
      mirrFinance: 6,
      mirrReinvest: 4
    },
    outputs: [
      { id: uid(), name: "Yield", unit: "t/ha", value: 300, source: "Input Directly" },
      { id: uid(), name: "Protein", unit: "%-point", value: 12, source: "Input Directly" },
      { id: uid(), name: "Moisture", unit: "%-point", value: -5, source: "Input Directly" },
      { id: uid(), name: "Biomass", unit: "t/ha", value: 40, source: "Input Directly" }
    ],
    treatments: [
      {
        id: uid(),
        name: "Optimized N (Rate+Timing)",
        area: 300,
        adoption: 0.8,
        deltas: {},
        annualCost: 45,
        capitalCost: 5000,
        constrained: true,
        source: "Farm Trials"
      },
      {
        id: uid(),
        name: "Slow-Release N",
        area: 200,
        adoption: 0.7,
        deltas: {},
        annualCost: 25,
        capitalCost: 0,
        constrained: true,
        source: "ABARES"
      }
    ],
    benefits: [
      {
        id: uid(),
        label: "Reduced recurring costs (energy/water)",
        category: "C4",
        frequency: "Annual",
        startYear: new Date().getFullYear(),
        endYear: new Date().getFullYear() + 4,
        year: new Date().getFullYear(),
        unitValue: 0,
        quantity: 0,
        abatement: 0,
        annualAmount: 15000,
        growthPct: 0,
        linkAdoption: true,
        linkRisk: true,
        p0: 0,
        p1: 0,
        consequence: 0,
        notes: "Project-wide OPEX saving"
      },
      {
        id: uid(),
        label: "Reduced risk of quality downgrades",
        category: "C7",
        frequency: "Annual",
        startYear: new Date().getFullYear(),
        endYear: new Date().getFullYear() + 9,
        year: new Date().getFullYear(),
        unitValue: 0,
        quantity: 0,
        abatement: 0,
        annualAmount: 0,
        growthPct: 0,
        linkAdoption: true,
        linkRisk: false,
        p0: 0.10,
        p1: 0.07,
        consequence: 120000,
        notes: ""
      },
      {
        id: uid(),
        label: "Soil asset value uplift (carbon/structure)",
        category: "C6",
        frequency: "Once",
        startYear: new Date().getFullYear(),
        endYear: new Date().getFullYear(),
        year: new Date().getFullYear() + 5,
        unitValue: 0,
        quantity: 0,
        abatement: 0,
        annualAmount: 50000,
        growthPct: 0,
        linkAdoption: false,
        linkRisk: true,
        p0: 0,
        p1: 0,
        consequence: 0,
        notes: ""
      }
    ],
    otherCosts: [
      {
        id: uid(),
        label: "Project Mgmt & M&E",
        type: "annual",
        annual: 20000,
        startYear: new Date().getFullYear(),
        endYear: new Date().getFullYear() + 4,
        capital: 0,
        year: new Date().getFullYear(),
        constrained: true
      }
    ],
    adoption: { base: 0.9, low: 0.6, high: 1.0 },
    risk: {
      base: 0.15,
      low: 0.05,
      high: 0.30,
      tech: 0.05,
      nonCoop: 0.04,
      socio: 0.02,
      fin: 0.03,
      man: 0.02
    },
    sim: {
      n: 1000,
      targetBCR: 2,
      bcrMode: "all",
      seed: null,
      results: { npv: [], bcr: [] },
      details: []
    }
  };

  function initTreatmentDeltas() {
    model.treatments.forEach(t => {
      model.outputs.forEach(o => {
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      });
    });
  }
  initTreatmentDeltas();

  // ---------- UTIL ----------
  function uid() { return Math.random().toString(36).slice(2, 10); }
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const fmt = n => (isFinite(n) ? (Math.abs(n) >= 1000 ? n.toLocaleString(undefined, { maximumFractionDigits: 0 }) : n.toLocaleString(undefined, { maximumFractionDigits: 2 })) : "—");
  const money = n => (isFinite(n) ? "$" + fmt(n) : "—");
  const percent = n => (isFinite(n) ? fmt(n) + "%" : "—");
  const slug = s => (s || "project").toLowerCase().replace(/[^a-z0-9]+/g,"_").replace(/^_|_$/g,"");
  const annuityFactor = (N, rPct) => { const r = rPct / 100; return r === 0 ? N : (1 - Math.pow(1 + r, -N)) / r; };
  const esc = s => (s ?? "").toString().replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));

  function saveWorkbook(filename, wb) {
    try { if (window.XLSX && typeof XLSX.writeFile === "function") { XLSX.writeFile(wb, filename, { compression: true }); return; } }
    catch(_) {}
    try {
      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array", compression: true });
      const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = filename;
      document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove();},300);
    } catch (err) { alert("Download failed. Ensure SheetJS script is loaded.\n\n"+(err?.message||err)); }
  }

  // ---------- CASHFLOWS ----------
  function buildCashflows({ forRate = model.time.discBase, adoptMul = model.adoption.base, risk = model.risk.base }) {
    const N = model.time.years;
    const baseYear = model.time.startYear;

    const benefitByYear = new Array(N + 1).fill(0);
    const costByYear = new Array(N + 1).fill(0);
    const constrainedCostByYear = new Array(N + 1).fill(0);

    // Treatments × Outputs
    let annualBenefit = 0, treatAnnualCost = 0, treatConstrAnnualCost = 0;
    let treatCapitalY0 = 0, treatConstrCapitalY0 = 0;

    model.treatments.forEach(t => {
      const adopt = clamp(t.adoption * adoptMul, 0, 1);
      let valuePerHa = 0;
      model.outputs.forEach(o => {
        const delta = Number(t.deltas[o.id]) || 0;
        const v = Number(o.value) || 0;
        valuePerHa += delta * v;
      });
      const benefit = valuePerHa * (Number(t.area) || 0) * (1 - clamp(risk, 0, 1)) * adopt;
      const opCost = (Number(t.annualCost) || 0) * (Number(t.area) || 0);
      const cap = Number(t.capitalCost) || 0;

      annualBenefit += benefit;
      treatAnnualCost += opCost;
      treatCapitalY0 += cap;

      if (t.constrained) {
        treatConstrAnnualCost += opCost;
        treatConstrCapitalY0 += cap;
      }
    });

    costByYear[0] += treatCapitalY0;
    constrainedCostByYear[0] += treatConstrCapitalY0;
    for (let t = 1; t <= N; t++) {
      benefitByYear[t] += annualBenefit;
      costByYear[t] += treatAnnualCost;
      constrainedCostByYear[t] += treatConstrAnnualCost;
    }

    // Other project costs
    const otherAnnualByYear = new Array(N + 1).fill(0);
    const otherConstrAnnualByYear = new Array(N + 1).fill(0);
    let otherCapitalY0 = 0, otherConstrCapitalY0 = 0;

    model.otherCosts.forEach(c => {
      if (c.type === "annual") {
        const a = Number(c.annual) || 0;
        const sy = Number(c.startYear) || baseYear;
        const ey = Number(c.endYear) || sy;
        for (let y = sy; y <= ey; y++) {
          const idx = y - baseYear + 1;
          if (idx >= 1 && idx <= N) {
            otherAnnualByYear[idx] += a;
            if (c.constrained) otherConstrAnnualByYear[idx] += a;
          }
        }
      } else if (c.type === "capital") {
        const cap = Number(c.capital) || 0;
        const cy = Number(c.year) || baseYear;
        const idx = cy - baseYear;
        if (idx === 0) {
          otherCapitalY0 += cap;
          if (c.constrained) otherConstrCapitalY0 += cap;
        } else if (idx > 0 && idx <= N) {
          costByYear[idx] += cap;
          if (c.constrained) constrainedCostByYear[idx] += cap;
        }
      }
    });

    costByYear[0] += otherCapitalY0;
    constrainedCostByYear[0] += otherConstrCapitalY0;
    for (let t = 1; t <= N; t++) {
      costByYear[t] += otherAnnualByYear[t];
      constrainedCostByYear[t] += otherConstrAnnualByYear[t];
    }

    // Additional benefits
    const extra = additionalBenefitsSeries(N, baseYear, adoptMul, risk);
    for (let i = 0; i < extra.length; i++) benefitByYear[i] += extra[i];

    const cf = new Array(N + 1).fill(0).map((_, i) => (benefitByYear[i] - costByYear[i]));
    const annualGM = annualBenefit - treatAnnualCost;
    return { benefitByYear, costByYear, constrainedCostByYear, cf, annualGM };
  }

  function additionalBenefitsSeries(N, baseYear, adoptMul, risk) {
    const series = new Array(N + 1).fill(0);
    model.benefits.forEach(b => {
      const cat = String(b.category || "").toUpperCase();
      const linkA = !!b.linkAdoption, linkR = !!b.linkRisk;
      const A = linkA ? clamp(adoptMul, 0, 1) : 1;
      const R = linkR ? (1 - clamp(risk, 0, 1)) : 1;
      const g = Number(b.growthPct) || 0;

      const addAnnual = (yearIndex, baseAmount, tFromStart) => {
        const grown = baseAmount * Math.pow(1 + g / 100, tFromStart);
        if (yearIndex >= 1 && yearIndex <= N) series[yearIndex] += grown * A * R;
      };
      const addOnce = (absYear, amount) => {
        const idx = absYear - baseYear + 1;
        if (idx >= 0 && idx <= N) series[idx] += amount * A * R;
      };

      const sy = Number(b.startYear) || baseYear;
      const ey = Number(b.endYear) || sy;
      const yr = Number(b.year) || sy;

      if (b.frequency === "Once" || cat === "C6") {
        const amount = Number(b.annualAmount) || 0;
        addOnce(yr, amount);
        return;
      }

      for (let y = sy; y <= ey; y++) {
        const idx = y - baseYear + 1;
        const tFromStart = y - sy;
        let amt = 0;
        switch (cat) {
          case "C1":
          case "C2":
          case "C3": {
            const v = Number(b.unitValue) || 0;
            const q = Number(cat === "C3" ? b.abatement : b.quantity) || 0;
            amt = v * q;
            break;
          }
          case "C4":
          case "C5":
          case "C8":
            amt = Number(b.annualAmount) || 0;
            break;
          case "C7": {
            const p0 = Number(b.p0) || 0, p1 = Number(b.p1) || 0;
            const c = Number(b.consequence) || 0;
            amt = Math.max(p0 - p1, 0) * c;
            break;
          }
          default:
            amt = 0;
        }
        addAnnual(idx, amt, tFromStart);
      }
    });
    return series;
  }

  function presentValue(series, ratePct) {
    let pv = 0;
    for (let t = 0; t < series.length; t++) pv += series[t] / Math.pow(1 + ratePct / 100, t);
    return pv;
  }

  function computeAll(rate, adoptMul, risk, bcrMode) {
    const { benefitByYear, costByYear, constrainedCostByYear, cf, annualGM } =
      buildCashflows({ forRate: rate, adoptMul, risk });
    const pvBenefits = presentValue(benefitByYear, rate);
    const pvCosts = presentValue(costByYear, rate);
    const pvCostsConstrained = presentValue(constrainedCostByYear, rate);

    const npv = pvBenefits - pvCosts;
    const denom = (bcrMode === "constrained") ? pvCostsConstrained : pvCosts;
    const bcr = denom > 0 ? pvBenefits / denom : NaN;

    const irrVal = irr(cf);
    const mirrVal = mirr(cf, model.time.mirrFinance, model.time.mirrReinvest);
    const roi = pvCosts > 0 ? ((pvBenefits - pvCosts) / pvCosts) * 100 : NaN;
    const profitMargin = benefitByYear[1] > 0 ? (annualGM / benefitByYear[1]) * 100 : NaN;
    const pb = payback(cf, rate);

    return { pvBenefits, pvCosts, pvCostsConstrained, npv, bcr, irrVal, mirrVal, roi, annualGM, profitMargin, paybackYears: pb, cf };
  }

  function irr(cf) {
    const hasPos = cf.some(v => v > 0);
    const hasNeg = cf.some(v => v < 0);
    if (!hasPos || !hasNeg) return NaN;
    let lo = -0.99, hi = 5.0;
    const npvAt = r => cf.reduce((acc, v, t) => acc + v / Math.pow(1 + r, t), 0);
    let nLo = npvAt(lo), nHi = npvAt(hi);
    if (nLo * nHi > 0) { for (let k = 0; k < 20 && nLo * nHi > 0; k++) { hi *= 1.5; nHi = npvAt(hi); } if (nLo * nHi > 0) return NaN; }
    for (let i = 0; i < 80; i++) {
      const mid = (lo + hi) / 2; const nMid = npvAt(mid);
      if (Math.abs(nMid) < 1e-8) return mid * 100;
      if (nLo * nMid <= 0) { hi = mid; nHi = nMid; } else { lo = mid; nLo = nMid; }
    }
    return ((lo + hi) / 2) * 100;
  }

  function mirr(cf, financeRatePct, reinvestRatePct) {
    const n = cf.length - 1;
    const fr = financeRatePct / 100, rr = reinvestRatePct / 100;
    let pvNeg = 0, fvPos = 0;
    for (let t = 0; t <= n; t++) {
      const v = cf[t];
      if (v < 0) pvNeg += v / Math.pow(1 + fr, t);
      if (v > 0) fvPos += v * Math.pow(1 + rr, n - t);
    }
    if (pvNeg === 0) return NaN;
    const mirrVal = Math.pow(-fvPos / pvNeg, 1 / n) - 1;
    return mirrVal * 100;
  }

  function payback(cf, ratePct) {
    let cum = 0;
    for (let t = 0; t < cf.length; t++) {
      cum += cf[t] / Math.pow(1 + ratePct / 100, t);
      if (cum >= 0) return t;
    }
    return null;
  }

  // ---------- DOM ----------
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));
  const num = sel => +(document.querySelector(sel)?.value || 0);
  const setVal = (sel, text) => (document.querySelector(sel).textContent = text);

  function switchTab(target){
    $$("#tabs button").forEach(b =>
      b.classList.toggle("active", b.dataset.tab === target)
    );
    $$(".tab-panel").forEach(p =>
      p.classList.toggle("show", p.id === `tab-${target}`)
    );
    if (target === "distribution") drawHists();
    if (target === "report") calcAndRender();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function initTabs() {
    document.addEventListener("click", (e) => {
      const navBtn = e.target.closest("#tabs button[data-tab]");
      const jumpBtn = e.target.closest("[data-tab-jump]");
      const target = navBtn?.dataset.tab || jumpBtn?.dataset.tabJump;
      if (!target) return;
      switchTab(target);
    });
  }

  function initActions(){
    document.addEventListener("click", (e) => {
      if (e.target.closest("#recalc, #getResults, [data-action='recalc']")) {
        e.preventDefault(); calcAndRender();
      }
      if (e.target.closest("#runSim, [data-action='run-sim']")) {
        e.preventDefault(); runSimulation();
      }
    });
  }

  // ---------- BIND + RENDER FORMS ----------
  function setBasicsFieldsFromModel() {
    $("#projectName").value = model.project.name || "";
    $("#analystNames").value = model.project.analysts || "";
    $("#projectSummary").value = model.project.summary || "";
    $("#lastUpdated").value = model.project.lastUpdated || "";
    $("#projectGoal").value = model.project.goal || "";
    $("#withProject").value = model.project.withProject || "";
    $("#withoutProject").value = model.project.withoutProject || "";
    $("#organisation").value = model.project.organisation || "";
    $("#contactEmail").value = model.project.contactEmail || "";

    $("#startYear").value = model.time.startYear;
    $("#years").value = model.time.years;
    $("#discBase").value = model.time.discBase;
    $("#discLow").value = model.time.discLow;
    $("#discHigh").value = model.time.discHigh;
    $("#mirrFinance").value = model.time.mirrFinance;
    $("#mirrReinvest").value = model.time.mirrReinvest;

    $("#adoptBase").value = model.adoption.base;
    $("#adoptLow").value = model.adoption.low;
    $("#adoptHigh").value = model.adoption.high;

    $("#riskBase").value = model.risk.base;
    $("#riskLow").value = model.risk.low;
    $("#riskHigh").value = model.risk.high;
    $("#rTech").value = model.risk.tech;
    $("#rNonCoop").value = model.risk.nonCoop;
    $("#rSocio").value = model.risk.socio;
    $("#rFin").value = model.risk.fin;
    $("#rMan").value = model.risk.man;

    $("#simN").value = model.sim.n;
    $("#targetBCR").value = model.sim.targetBCR;
    $("#bcrMode").value = model.sim.bcrMode;
    $("#simBcrTargetLabel").textContent = model.sim.targetBCR;
  }

  function bindBasics() {
    setBasicsFieldsFromModel();

    $("#calcCombinedRisk").addEventListener("click", () => {
      const r = 1 - (1 - num("#rTech")) * (1 - num("#rNonCoop")) * (1 - num("#rSocio")) * (1 - num("#rFin")) * (1 - num("#rMan"));
      $("#combinedRiskOut").textContent = `Combined: ${(r * 100).toFixed(2)}%`;
      $("#riskBase").value = r.toFixed(3);
      model.risk.base = r;
      calcAndRender();
    });

    $("#addCost").addEventListener("click", () => {
      const c = { id: uid(), label: "New Cost", type: "annual", annual: 0, startYear: model.time.startYear, endYear: model.time.startYear, capital: 0, year: model.time.startYear, constrained: true };
      model.otherCosts.push(c);
      renderCosts();
      calcAndRender();
    });

    document.addEventListener("input", e => {
      const id = e.target.id;
      if (!id) return;
      switch (id) {
        case "projectName": model.project.name = e.target.value; break;
        case "analystNames": model.project.analysts = e.target.value; break;
        case "projectSummary": model.project.summary = e.target.value; break;
        case "lastUpdated": model.project.lastUpdated = e.target.value; break;
        case "projectGoal": model.project.goal = e.target.value; break;
        case "withProject": model.project.withProject = e.target.value; break;
        case "withoutProject": model.project.withoutProject = e.target.value; break;
        case "organisation": model.project.organisation = e.target.value; break;
        case "contactEmail": model.project.contactEmail = e.target.value; break;

        case "startYear": model.time.startYear = +e.target.value; break;
        case "years": model.time.years = +e.target.value; break;
        case "discBase": model.time.discBase = +e.target.value; break;
        case "discLow": model.time.discLow = +e.target.value; break;
        case "discHigh": model.time.discHigh = +e.target.value; break;
        case "mirrFinance": model.time.mirrFinance = +e.target.value; break;
        case "mirrReinvest": model.time.mirrReinvest = +e.target.value; break;

        case "adoptBase": model.adoption.base = +e.target.value; break;
        case "adoptLow": model.adoption.low = +e.target.value; break;
        case "adoptHigh": model.adoption.high = +e.target.value; break;

        case "riskBase": model.risk.base = +e.target.value; break;
        case "riskLow": model.risk.low = +e.target.value; break;
        case "riskHigh": model.risk.high = +e.target.value; break;
        case "rTech": model.risk.tech = +e.target.value; break;
        case "rNonCoop": model.risk.nonCoop = +e.target.value; break;
        case "rSocio": model.risk.socio = +e.target.value; break;
        case "rFin": model.risk.fin = +e.target.value; break;
        case "rMan": model.risk.man = +e.target.value; break;

        case "simN": model.sim.n = +e.target.value; break;
        case "targetBCR": model.sim.targetBCR = +e.target.value; $("#simBcrTargetLabel").textContent = e.target.value; break;
        case "bcrMode": model.sim.bcrMode = e.target.value; break;
        case "randSeed": model.sim.seed = e.target.value ? +e.target.value : null; break;
      }
      calcAndRenderDebounced();
    });

    $("#saveProject").addEventListener("click", () => {
      const data = JSON.stringify(model, null, 2);
      downloadFile(`cba_${(model.project.name || "project").replace(/\s+/g, "_")}.json`, data, "application/json");
    });
    $("#loadProject").addEventListener("click", () => $("#loadFile").click());
    $("#loadFile").addEventListener("change", async e => {
      const file = e.target.files?.[0]; if (!file) return;
      const text = await file.text();
      try {
        const obj = JSON.parse(text);
        Object.assign(model, obj);
        initTreatmentDeltas();
        renderAll();
        setBasicsFieldsFromModel();
        calcAndRender();
      } catch {
        alert("Invalid JSON file.");
      } finally { e.target.value = ""; }
    });

    $("#exportCsv").addEventListener("click", exportAllCsv);
    $("#exportCsvFoot").addEventListener("click", exportAllCsv);
    $("#exportPdf").addEventListener("click", exportPdf);
    $("#exportPdfFoot").addEventListener("click", exportPdf);

    $("#parseExcel").addEventListener("click", handleParseExcel);
    $("#importExcel").addEventListener("click", commitExcelToModel);

    $("#downloadTemplate").addEventListener("click", downloadExcelTemplate);
    $("#downloadSample").addEventListener("click", downloadSampleDataset);

    // Make Start go to Project (cover stays available)
    $("#startBtn")?.addEventListener("click", () => switchTab("project"));
  }

  // ---------- RENDERERS ----------
  function renderOutputs() {
    const root = $("#outputsList"); root.innerHTML = "";
    model.outputs.forEach(o => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>Output: ${esc(o.name)}</h4>
        <div class="row-6">
          <div class="field"><label>Name</label><input value="${esc(o.name)}" data-k="name" data-id="${o.id}"/></div>
          <div class="field"><label>Unit</label><input value="${esc(o.unit)}" data-k="unit" data-id="${o.id}"/></div>
          <div class="field"><label>Value ($/unit)</label><input type="number" step="0.01" value="${o.value}" data-k="value" data-id="${o.id}"/></div>
          <div class="field"><label>Source</label>
            <select data-k="source" data-id="${o.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"].map(s => `<option ${s===o.source?"selected":""}>${s}</option>`).join("")}
            </select>
          </div>
          <div class="field"><label>&nbsp;</label><button class="danger" data-del-output="${o.id}">Remove</button></div>
        </div>
        <div class="kv"><small class="muted">id:</small> <code>${o.id}</code></div>
      `;
      root.appendChild(el);
    });
    root.addEventListener("input", onOutputEdit);
    root.addEventListener("click", onOutputDelete);
  }
  function onOutputEdit(e) {
    const k = e.target.dataset.k, id = e.target.dataset.id; if (!k || !id) return;
    const o = model.outputs.find(x => x.id === id); if (!o) return;
    if (k === "value") o[k] = +e.target.value; else o[k] = e.target.value;
    model.treatments.forEach(t => { if (!(id in t.deltas)) t.deltas[id] = 0; });
    renderTreatments(); renderDatabaseTags(); calcAndRenderDebounced();
  }
  function onOutputDelete(e) {
    const id = e.target.dataset.delOutput; if (!id) return;
    if (!confirm("Remove this output metric?")) return;
    model.outputs = model.outputs.filter(o => o.id !== id);
    model.treatments.forEach(t => delete t.deltas[id]);
    renderOutputs(); renderTreatments(); renderDatabaseTags(); calcAndRender();
  }
  $("#addOutput")?.addEventListener("click", () => {
    const id = uid();
    model.outputs.push({ id, name: "Custom Output", unit: "unit", value: 0, source: "Input Directly" });
    model.treatments.forEach(t => (t.deltas[id] = 0));
    renderOutputs(); renderTreatments(); renderDatabaseTags();
  });

  function renderTreatments() {
    const root = $("#treatmentsList"); root.innerHTML = "";
    model.treatments.forEach(t => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>🚜 Treatment: ${esc(t.name)}</h4>
        <div class="row">
          <div class="field"><label>Name</label><input value="${esc(t.name)}" data-tk="name" data-id="${t.id}" /></div>
          <div class="field"><label>Area (ha)</label><input type="number" step="0.01" value="${t.area}" data-tk="area" data-id="${t.id}" /></div>
          <div class="field"><label>Adoption (0–1)</label><input type="number" min="0" max="1" step="0.01" value="${t.adoption}" data-tk="adoption" data-id="${t.id}" /></div>
          <div class="field"><label>Source</label>
            <select data-tk="source" data-id="${t.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"].map(s => `<option ${s===t.source?"selected":""}>${s}</option>`).join("")}
            </select>
          </div>
          <div class="field"><label>Annual Cost ($/ha)</label><input type="number" step="0.01" value="${t.annualCost}" data-tk="annualCost" data-id="${t.id}" /></div>
          <div class="field"><label>Capital Cost ($, y0)</label><input type="number" step="0.01" value="${t.capitalCost}" data-tk="capitalCost" data-id="${t.id}" /></div>
          <div class="field"><label>Constrained?</label>
            <select data-tk="constrained" data-id="${t.id}">
              <option value="true" ${t.constrained?"selected":""}>Yes</option>
              <option value="false" ${!t.constrained?"selected":""}>No</option>
            </select>
          </div>
          <div class="field"><label>&nbsp;</label><button class="danger" data-del-treatment="${t.id}">Remove</button></div>
        </div>
        <h5>Output Deltas (per ha)</h5>
        <div class="row">
          ${model.outputs.map(o => `
            <div class="field">
              <label>${esc(o.name)} (${esc(o.unit)})</label>
              <input type="number" step="0.0001" value="${t.deltas[o.id] ?? 0}" data-td="${o.id}" data-id="${t.id}" />
            </div>
          `).join("")}
        </div>
        <div class="kv"><small class="muted">id:</small> <code>${t.id}</code></div>
      `;
      root.appendChild(el);
    });
    root.oninput = e => {
      const id = e.target.dataset.id; if (!id) return;
      const t = model.treatments.find(x => x.id === id); if (!t) return;
      const tk = e.target.dataset.tk;
      if (tk) {
        if (tk === "constrained") t[tk] = e.target.value === "true";
        else if (tk === "name" || tk === "source") t[tk] = e.target.value;
        else t[tk] = +e.target.value;
      }
      const td = e.target.dataset.td;
      if (td) t.deltas[td] = +e.target.value;
      calcAndRenderDebounced();
    };
    root.addEventListener("click", e => {
      const id = e.target.dataset.delTreatment; if (!id) return;
      if (!confirm("Remove this treatment?")) return;
      model.treatments = model.treatments.filter(x => x.id !== id);
      renderTreatments(); renderDatabaseTags(); calcAndRender();
    });
  }

  function renderBenefits() {
    const root = $("#benefitsList"); root.innerHTML = "";
    const rowFor = b => `
      <div class="item">
        <h4>🌱 ${esc(b.label || "Benefit")}</h4>
        <div class="row-6">
          <div class="field"><label>Label</label><input value="${esc(b.label||"")}" data-bk="label" data-id="${b.id}" /></div>
          <div class="field"><label>Category</label>
            <select data-bk="category" data-id="${b.id}">
              ${["C1","C2","C3","C4","C5","C6","C7","C8"].map(c=>`<option ${c===b.category?"selected":""}>${c}</option>`).join("")}
            </select>
          </div>
          <div class="field"><label>Frequency</label>
            <select data-bk="frequency" data-id="${b.id}">
              <option ${b.frequency==="Annual"?"selected":""}>Annual</option>
              <option ${b.frequency==="Once"?"selected":""}>Once</option>
            </select>
          </div>
          <div class="field"><label>Start year</label><input type="number" value="${b.startYear||model.time.startYear}" data-bk="startYear" data-id="${b.id}" /></div>
          <div class="field"><label>End year</label><input type="number" value="${b.endYear||model.time.startYear}" data-bk="endYear" data-id="${b.id}" /></div>
          <div class="field"><label>Once year</label><input type="number" value="${b.year||model.time.startYear}" data-bk="year" data-id="${b.id}" /></div>
        </div>

        <div class="row-6">
          <div class="field"><label>Unit value ($)</label><input type="number" step="0.01" value="${b.unitValue||0}" data-bk="unitValue" data-id="${b.id}" /></div>
          <div class="field"><label>Quantity</label><input type="number" step="0.01" value="${b.quantity||0}" data-bk="quantity" data-id="${b.id}" /></div>
          <div class="field"><label>Abatement</label><input type="number" step="0.01" value="${b.abatement||0}" data-bk="abatement" data-id="${b.id}" /></div>
          <div class="field"><label>Annual amount ($)</label><input type="number" step="0.01" value="${b.annualAmount||0}" data-bk="annualAmount" data-id="${b.id}" /></div>
          <div class="field"><label>Growth (%/yr)</label><input type="number" step="0.01" value="${b.growthPct||0}" data-bk="growthPct" data-id="${b.id}" /></div>
          <div class="field"><label>Link adoption?</label>
            <select data-bk="linkAdoption" data-id="${b.id}">
              <option value="true" ${b.linkAdoption?"selected":""}>Yes</option>
              <option value="false" ${!b.linkAdoption?"selected":""}>No</option>
            </select>
          </div>
          <div class="field"><label>Link risk?</label>
            <select data-bk="linkRisk" data-id="${b.id}">
              <option value="true" ${b.linkRisk?"selected":""}>Yes</option>
              <option value="false" ${!b.linkRisk?"selected":""}>No</option>
            </select>
          </div>
        </div>

        <div class="row-6">
          <div class="field"><label>P0 (baseline prob)</label><input type="number" step="0.001" value="${b.p0||0}" data-bk="p0" data-id="${b.id}" /></div>
          <div class="field"><label>P1 (with-project prob)</label><input type="number" step="0.001" value="${b.p1||0}" data-bk="p1" data-id="${b.id}" /></div>
          <div class="field"><label>Consequence ($)</label><input type="number" step="0.01" value="${b.consequence||0}" data-bk="consequence" data-id="${b.id}" /></div>
          <div class="field"><label>Notes</label><input value="${esc(b.notes||"")}" data-bk="notes" data-id="${b.id}" /></div>
          <div class="field"><label>&nbsp;</label><button class="danger" data-del-benefit="${b.id}">Remove</button></div>
        </div>
      </div>
    `;

    model.benefits.forEach(b => {
      const el = document.createElement("div");
      el.innerHTML = rowFor(b);
      root.appendChild(el.firstElementChild);
    });

    root.oninput = e => {
      const id = e.target.dataset.id; if (!id) return;
      const b = model.benefits.find(x => x.id === id); if (!b) return;
      const k = e.target.dataset.bk;
      if (!k) return;
      if (k === "label" || k === "category" || k === "frequency" || k === "notes") b[k] = e.target.value;
      else if (k === "linkAdoption" || k === "linkRisk") b[k] = e.target.value === "true";
      else b[k] = +e.target.value;
      calcAndRenderDebounced();
    };
    root.addEventListener("click", e => {
      const id = e.target.dataset.delBenefit; if (!id) return;
      if (!confirm("Remove this benefit item?")) return;
      model.benefits = model.benefits.filter(x => x.id !== id);
      renderBenefits(); calcAndRender();
    });
  }
  $("#addBenefit")?.addEventListener("click", () => {
    model.benefits.push({
      id: uid(), label: "New Benefit", category: "C4", frequency: "Annual",
      startYear: model.time.startYear, endYear: model.time.startYear, year: model.time.startYear,
      unitValue: 0, quantity: 0, abatement: 0, annualAmount: 0, growthPct: 0,
      linkAdoption: true, linkRisk: true, p0: 0, p1: 0, consequence: 0, notes: ""
    });
    renderBenefits();
  });

  function renderCosts() {
    const root = $("#costsList"); root.innerHTML = "";
    model.otherCosts.forEach(c => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>💰 Cost Item: ${esc(c.label)}</h4>
        <div class="row-6">
          <div class="field"><label>Label</label><input value="${esc(c.label)}" data-ck="label" data-id="${c.id}" /></div>
          <div class="field"><label>Type</label>
            <select data-ck="type" data-id="${c.id}">
              <option value="annual" ${c.type==="annual"?"selected":""}>Annual</option>
              <option value="capital" ${c.type==="capital"?"selected":""}>Capital</option>
            </select>
          </div>
          <div class="field"><label>Annual ($/yr)</label><input type="number" step="0.01" value="${c.annual ?? 0}" data-ck="annual" data-id="${c.id}" /></div>
          <div class="field"><label>Start Year</label><input type="number" value="${c.startYear ?? model.time.startYear}" data-ck="startYear" data-id="${c.id}" /></div>
          <div class="field"><label>End Year</label><input type="number" value="${c.endYear ?? model.time.startYear}" data-ck="endYear" data-id="${c.id}" /></div>
          <div class="field"><label>Capital ($)</label><input type="number" step="0.01" value="${c.capital ?? 0}" data-ck="capital" data-id="${c.id}" /></div>
          <div class="field"><label>Capital Year</label><input type="number" value="${c.year ?? model.time.startYear}" data-ck="year" data-id="${c.id}" /></div>
          <div class="field"><label>Constrained?</label>
            <select data-ck="constrained" data-id="${c.id}">
              <option value="true" ${c.constrained?"selected":""}>Yes</option>
              <option value="false" ${!c.constrained?"selected":""}>No</option>
            </select>
          </div>
          <div class="field"><label>&nbsp;</label><button class="danger" data-del-cost="${c.id}">Remove</button></div>
        </div>
      `;
      root.appendChild(el);
    });
    root.oninput = e => {
      const id = e.target.dataset.id, k = e.target.dataset.ck; if (!id || !k) return;
      const c = model.otherCosts.find(x => x.id === id); if (!c) return;
      if (k === "label" || k === "type") c[k] = e.target.value;
      else if (k === "constrained") c[k] = e.target.value === "true";
      else c[k] = +e.target.value;
      calcAndRenderDebounced();
    };
    root.addEventListener("click", e => {
      const id = e.target.dataset.delCost; if (!id) return;
      if (!confirm("Remove this cost item?")) return;
      model.otherCosts = model.otherCosts.filter(x => x.id !== id);
      renderCosts(); calcAndRender();
    });
  }

  function renderDatabaseTags() {
    const outRoot = $("#dbOutputs"); outRoot.innerHTML = "";
    model.outputs.forEach(o => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <div class="row-2">
          <div class="field"><label>${esc(o.name)} (${esc(o.unit)})</label></div>
          <div class="field">
            <label>Source</label>
            <select data-db-out="${o.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"].map(s => `<option ${s===o.source?"selected":""}>${s}</option>`).join("")}
            </select>
          </div>
        </div>`;
      outRoot.appendChild(el);
    });
    outRoot.onchange = e => {
      const id = e.target.dataset.dbOut;
      const o = model.outputs.find(x => x.id === id);
      if (o) o.source = e.target.value;
    };

    const tRoot = $("#dbTreatments"); tRoot.innerHTML = "";
    model.treatments.forEach(t => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <div class="row-2">
          <div class="field"><label>${esc(t.name)}</label></div>
          <div class="field">
            <label>Source</label>
            <select data-db-t="${t.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"].map(s => `<option ${s===t.source?"selected":""}>${s}</option>`).join("")}
            </select>
          </div>
        </div>`;
      tRoot.appendChild(el);
    });
    tRoot.onchange = e => {
      const id = e.target.dataset.dbT;
      const t = model.treatments.find(x => x.id === id);
      if (t) t.source = e.target.value;
    };
  }

  function renderTreatmentSummary(rate, adoptMul, risk) {
    const root = $("#treatmentSummary"); root.innerHTML = "";
    model.treatments.forEach(t => {
      let valuePerHa = 0;
      model.outputs.forEach(o => (valuePerHa += (Number(t.deltas[o.id]) || 0) * (Number(o.value) || 0)));
      const adopt = clamp(t.adoption * adoptMul, 0, 1);
      const annualBen = valuePerHa * (Number(t.area) || 0) * (1 - clamp(risk, 0, 1)) * adopt;
      const annualCost = (Number(t.annualCost) || 0) * (Number(t.area) || 0);
      const cap = Number(t.capitalCost) || 0;
      const pvBen = annualBen * annuityFactor(model.time.years, rate);
      const pvCost = cap + annualCost * annuityFactor(model.time.years, rate);
      const bcr = pvCost > 0 ? pvBen / pvCost : NaN;
      const npv = pvBen - pvCost;

      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <div class="row-6">
          <div class="field"><label>Treatment</label><div class="metric"><div class="value">${esc(t.name)}</div></div></div>
          <div class="field"><label>Area</label><div class="metric"><div class="value">${fmt(t.area)} ha</div></div></div>
          <div class="field"><label>Adoption</label><div class="metric"><div class="value">${fmt(adopt)}</div></div></div>
          <div class="field"><label>Annual Benefit</label><div class="metric"><div class="value">${money(annualBen)}</div></div></div>
          <div class="field"><label>Annual Cost</label><div class="metric"><div class="value">${money(annualCost)}</div></div></div>
          <div class="field"><label>PV Benefit / PV Cost</label><div class="metric"><div class="value">${money(pvBen)} / ${money(pvCost)}</div></div></div>
          <div class="field"><label>BCR</label><div class="metric"><div class="value">${isFinite(bcr)?fmt(bcr):"—"}</div></div></div>
          <div class="field"><label>NPV</label><div class="metric"><div class="value">${money(npv)}</div></div></div>
        </div>`;
      root.appendChild(el);
    });
  }

  function renderAll() {
    renderOutputs();
    renderTreatments();
    renderBenefits();
    renderDatabaseTags();
    renderCosts();
  }

  // ---------- MAIN CALC / REPORT ----------
  function calcAndRender() {
    const rate = model.time.discBase;
    const adoptMul = model.adoption.base;
    const risk = model.risk.base;

    const { pvBenefits, pvCosts, npv, bcr, irrVal, mirrVal, roi, annualGM, profitMargin, paybackYears } =
      computeAll(rate, adoptMul, risk, model.sim.bcrMode);

    setVal("#pvBenefits", money(pvBenefits));
    setVal("#pvCosts", money(pvCosts));
    const npvEl = $("#npv");
    npvEl.textContent = money(npv);
    npvEl.className = "value " + (npv >= 0 ? "positive" : "negative");
    setVal("#bcr", isFinite(bcr) ? fmt(bcr) : "—");
    setVal("#irr", isFinite(irrVal) ? percent(irrVal) : "—");
    setVal("#mirr", isFinite(mirrVal) ? percent(mirrVal) : "—");
    setVal("#roi", isFinite(roi) ? percent(roi) : "—");
    setVal("#grossMargin", money(annualGM));
    setVal("#profitMargin", isFinite(profitMargin) ? percent(profitMargin) : "—");
    setVal("#payback", paybackYears != null ? paybackYears : "Not reached");

    renderTreatmentSummary(rate, adoptMul, risk);
    $("#simBcrTargetLabel").textContent = model.sim.targetBCR;
  }

  let debTimer = null;
  function calcAndRenderDebounced() {
    clearTimeout(debTimer);
    debTimer = setTimeout(calcAndRender, 120);
  }

  // ---------- MONTE CARLO ----------
  function rng(seed) {
    let t = (seed || Math.floor(Math.random() * 2 ** 31)) >>> 0;
    return () => {
      t += 0x6D2B79F5;
      let x = t;
      x = Math.imul(x ^ (x >>> 15), 1 | x);
      x ^= x + Math.imul(x ^ (x >>> 7), 61 | x);
      return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
    };
  }
  function triangular(r, a, c, b) {
    const F = (c - a) / (b - a);
    if (r < F) return a + Math.sqrt(r * (b - a) * (c - a));
    return b - Math.sqrt((1 - r) * (b - a) * (b - c));
  }

  async function runSimulation() {
    $("#simStatus").textContent = "Running…";
    await new Promise(r => setTimeout(r));
    const N = model.sim.n;
    const seed = model.sim.seed;
    const rand = rng(seed ?? undefined);

    const discLow = model.time.discLow, discBase = model.time.discBase, discHigh = model.time.discHigh;
    const adoptLow = model.adoption.low, adoptBase = model.adoption.base, adoptHigh = model.adoption.high;
    const riskLow = model.risk.low, riskBase = model.risk.base, riskHigh = model.risk.high;

    const npvs = new Array(N);
    const bcrs = new Array(N);
    const details = [];

    for (let i = 0; i < N; i++) {
      const r1 = rand(), r2 = rand(), r3 = rand();
      const disc = triangular(r1, discLow, discBase, discHigh);
      const adoptMul = clamp(triangular(r2, adoptLow, adoptBase, adoptHigh), 0, 1);
      const risk = clamp(triangular(r3, riskLow, riskBase, riskHigh), 0, 1);

      const { pvBenefits, pvCosts, bcr, npv } = computeAll(disc, adoptMul, risk, model.sim.bcrMode);

      npvs[i] = npv;
      bcrs[i] = bcr;
      details.push({ run: i + 1, discount: disc, adoption: adoptMul, risk, pvBenefits, pvCosts, npv, bcr });
    }

    model.sim.results = { npv: npvs, bcr: bcrs };
    model.sim.details = details;
    $("#simStatus").textContent = "Done.";
    renderSimulationResults();
    drawHists();
  }

  function renderSimulationResults() {
    const { npv, bcr } = model.sim.results;
    if (!npv?.length) return;
    const sortedNpv = [...npv].sort((a, b) => a - b);
    const validBcr = bcr.filter(x => isFinite(x));
    const sortedBcr = [...validBcr].sort((a, b) => a - b);
    const N = npv.length, NB = sortedBcr.length;

    const stats = arr => ({
      min: arr[0],
      max: arr[arr.length - 1],
      mean: arr.reduce((a, c) => a + c, 0) / arr.length,
      median: arr.length ? (arr[Math.floor((arr.length - 1) / 2)] + arr[Math.ceil((arr.length - 1) / 2)]) / 2 : NaN
    });

    const sN = stats(sortedNpv);
    const sB = stats(sortedBcr.length ? sortedBcr : [NaN]);

    setVal("#simNpvMin", money(sN.min));
    setVal("#simNpvMax", money(sN.max));
    setVal("#simNpvMean", money(sN.mean));
    setVal("#simNpvMedian", money(sN.median));
    const pN = npv.filter(x => x > 0).length / N * 100;
    setVal("#simNpvProb", fmt(pN) + "%");

    setVal("#simBcrMin", isFinite(sB.min) ? fmt(sB.min) : "—");
    setVal("#simBcrMax", isFinite(sB.max) ? fmt(sB.max) : "—");
    setVal("#simBcrMean", isFinite(sB.mean) ? fmt(sB.mean) : "—");
    setVal("#simBcrMedian", isFinite(sB.median) ? fmt(sB.median) : "—");
    const pB1 = NB ? validBcr.filter(x => x > 1).length / NB * 100 : 0;
    setVal("#simBcrProb1", fmt(pB1) + "%");
    const tgt = model.sim.targetBCR;
    const pBt = NB ? validBcr.filter(x => x > tgt).length / NB * 100 : 0;
    setVal("#simBcrProbTarget", fmt(pBt) + "%");
  }

  function drawHist(canvasId, data, bins = 24, labelFmt = v => v.toFixed(0)) {
    const canvas = document.getElementById(canvasId);
    if (!canvas || !data?.length) return;
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    const min = Math.min(...data), max = Math.max(...data);
    const padL = 54, padR = 14, padT = 10, padB = 34;
    const W = canvas.width - padL - padR, H = canvas.height - padT - padB;

    const counts = new Array(bins).fill(0);
    const span = (max - min) || 1e-9;
    data.forEach(v => {
      let idx = Math.floor(((v - min) / span) * bins);
      if (idx < 0) idx = 0;
      if (idx >= bins) idx = bins - 1;
      counts[idx]++;
    });
    const maxC = Math.max(...counts) || 1;

    // axes
    ctx.strokeStyle = "#3c6a52";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(padL, padT);
    ctx.lineTo(padL, padT + H);
    ctx.lineTo(padL + W, padT + H);
    ctx.stroke();

    // bars
    for (let i = 0; i < bins; i++) {
      const x = padL + (i * W) / bins + 1;
      const h = (counts[i] / maxC) * (H - 2);
      const y = padT + H - h;
      ctx.fillStyle = "rgba(116, 209, 140, 0.45)";
      ctx.fillRect(x, y, (W / bins) - 2, h);
    }

    // x labels
    ctx.fillStyle = "#c9efd6";
    ctx.font = "12px system-ui";
    ctx.textAlign = "center";
    const lbls = [min, (min + max) / 2, max];
    [0, 0.5, 1].forEach((p, i) => {
      const x = padL + p * W;
      ctx.fillText(labelFmt(lbls[i]), x, padT + H + 20);
    });
  }
  function drawHists() {
    const { npv, bcr } = model.sim.results;
    if (npv?.length) drawHist("histNpv", npv, 24, v => money(v));
    if (bcr?.length) drawHist("histBcr", bcr.filter(x => isFinite(x)), 24, v => v.toFixed(2));
  }

  // ---------- EXPORTS ----------
  function toCsv(rows) {
    return rows.map(r => r.map(v => {
      const s = (v ?? "").toString();
      if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
      return s;
    }).join(",")).join("\n");
  }
  function downloadFile(filename, text, mime="text/csv") {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([text], { type: mime }));
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function buildSummaryForCsv() {
    const rate = model.time.discBase;
    const adoptMul = model.adoption.base;
    const risk = model.risk.base;
    const all = computeAll(rate, adoptMul, risk, model.sim.bcrMode);

    return {
      meta: {
        name: model.project.name,
        analysts: model.project.analysts,
        organisation: model.project.organisation,
        contact: model.project.contactEmail,
        updated: model.project.lastUpdated
      },
      params: {
        startYear: model.time.startYear,
        years: model.time.years,
        discountBase: model.time.discBase,
        discountLow: model.time.discLow,
        discountHigh: model.time.discHigh,
        mirrFinance: model.time.mirrFinance,
        mirrReinvest: model.time.mirrReinvest,
        adoptionBase: model.adoption.base,
        riskBase: model.risk.base,
        bcrMode: model.sim.bcrMode
      },
      results: all
    };
  }

  function exportAllCsv() {
    const s = buildSummaryForCsv();
    const summaryRows = [
      ["Project", s.meta.name],
      ["Analysts", s.meta.analysts],
      ["Organisation", s.meta.organisation],
      ["Contact", s.meta.contact],
      ["Last Updated", s.meta.updated],
      [],
      ["Start Year", s.params.startYear],
      ["Years", s.params.years],
      ["Discount Rate (Base)", s.params.discountBase],
      ["Discount Rate (Low)", s.params.discountLow],
      ["Discount Rate (High)", s.params.discountHigh],
      ["MIRR Finance %", s.params.mirrFinance],
      ["MIRR Reinvest %", s.params.mirrReinvest],
      ["Adoption Multiplier", s.params.adoptionBase],
      ["Risk (overall)", s.params.riskBase],
      ["BCR Mode", s.params.bcrMode],
      [],
      ["PV Benefits", s.results.pvBenefits],
      ["PV Costs", s.results.pvCosts],
      ["NPV", s.results.npv],
      ["BCR", s.results.bcr],
      ["IRR %", s.results.irrVal],
      ["MIRR %", s.results.mirrVal],
      ["ROI %", s.results.roi],
      ["Gross Margin (annual)", s.results.annualGM],
      ["Gross Profit Margin %", s.results.profitMargin],
      ["Payback (years)", s.results.paybackYears ?? "Not reached"]
    ];
    downloadFile(`cba_summary_${slug(s.meta.name)}.csv`, toCsv(summaryRows));

    const treatHeader = ["Treatment","Area(ha)","Adoption","Annual Benefit","Annual Cost","PV Benefit","PV Cost","BCR","NPV"];
    const treatRows = [treatHeader];
    model.treatments.forEach(t => {
      let valuePerHa = 0;
      model.outputs.forEach(o => valuePerHa += ((+t.deltas[o.id]||0) * (+o.value||0)));
      const adopt = clamp(t.adoption * model.adoption.base, 0, 1);
      const annBen = valuePerHa * (t.area||0) * (1 - model.risk.base) * adopt;
      const annCost = (t.annualCost||0) * (t.area||0);
      const pvB = annBen * annuityFactor(model.time.years, model.time.discBase);
      const pvC = (t.capitalCost||0) + annCost * annuityFactor(model.time.years, model.time.discBase);
      const bcr = pvC>0 ? pvB/pvC : "";
      const npv = pvB - pvC;
      treatRows.push([t.name, t.area, adopt, annBen, annCost, pvB, pvC, bcr, npv]);
    });
    downloadFile(`cba_treatments_${slug(s.meta.name)}.csv`, toCsv(treatRows));

    const benRows = [["Label","Category","Frequency","StartYear","EndYear","Year","UnitValue","Quantity","Abatement","AnnualAmount","GrowthPct","LinkAdoption","LinkRisk","P0","P1","Consequence","Notes"]];
    model.benefits.forEach(b => benRows.push([b.label,b.category,b.frequency,b.startYear,b.endYear,b.year,b.unitValue,b.quantity,b.abatement,b.annualAmount,b.growthPct,b.linkAdoption,b.linkRisk,b.p0,b.p1,b.consequence,b.notes]));
    downloadFile(`cba_benefits_${slug(s.meta.name)}.csv`, toCsv(benRows));

    const outRows = [["Output","Unit","$/unit","Source","Id"]];
    model.outputs.forEach(o => outRows.push([o.name,o.unit,o.value,o.source,o.id]));
    downloadFile(`cba_outputs_${slug(s.meta.name)}.csv`, toCsv(outRows));

    const { npv, bcr } = model.sim.results;
    if (npv?.length) {
      const validBcr = bcr.filter(x => isFinite(x));
      const stats = arr => {
        const a = [...arr].sort((x,y)=>x-y);
        const N = a.length;
        const med = (a[Math.floor((N-1)/2)] + a[Math.ceil((N-1)/2)]) / 2;
        const mean = a.reduce((u,v)=>u+v,0) / N;
        return { min:a[0], max:a[N-1], mean, median:med };
      };
      const sN = stats(npv);
      const sB = validBcr.length ? stats(validBcr) : {min:"",max:"",mean:"",median:""};
      const pN = npv.filter(x => x > 0).length / npv.length * 100;
      const tgt = model.sim.targetBCR;
      const pB1 = validBcr.filter(x => x > 1).length / (validBcr.length||1) * 100;
      const pBt = validBcr.filter(x => x > tgt).length / (validBcr.length||1) * 100;

      const simRows = [
        ["N", model.sim.n],
        ["BCR Mode", model.sim.bcrMode],
        ["NPV Min", sN.min],["NPV Max", sN.max],["NPV Mean", sN.mean],["NPV Median", sN.median],["Pr(NPV>0)%", pN],
        [],
        ["BCR Min", sB.min],["BCR Max", sB.max],["BCR Mean", sB.mean],["BCR Median", sB.median],["Pr(BCR>1)%", pB1],["Pr(BCR>Target)%", pBt]
      ];
      downloadFile(`cba_simulation_summary_${slug(s.meta.name)}.csv`, toCsv(simRows));

      const rawRows = [["run","discount","adoption","risk","pvBenefits","pvCosts","npv","bcr"]];
      model.sim.details.forEach(d => rawRows.push([d.run,d.discount,d.adoption,d.risk,d.pvBenefits,d.pvCosts,d.npv,d.bcr]));
      downloadFile(`cba_simulation_raw_${slug(s.meta.name)}.csv`, toCsv(rawRows));
    } else {
      const simRows = [["N", model.sim.n],["BCR Mode", model.sim.bcrMode],["Note","Run Monte Carlo to populate results."]];
      downloadFile(`cba_simulation_summary_${slug(s.meta.name)}.csv`, toCsv(simRows));
    }
  }

  function exportPdf() {
    drawHists();

    const npvCan = document.getElementById("histNpv");
    const bcrCan = document.getElementById("histBcr");
    const npvImg = (npvCan && npvCan.width) ? npvCan.toDataURL("image/png") : null;
    const bcrImg = (bcrCan && bcrCan.width) ? bcrCan.toDataURL("image/png") : null;

    const s = buildSummaryForCsv();

    const trRows = model.treatments.map(t => {
      let valuePerHa = 0;
      model.outputs.forEach(o => valuePerHa += ((+t.deltas[o.id]||0) * (+o.value||0)));
      const adopt = clamp(t.adoption * model.adoption.base, 0, 1);
      const annBen = valuePerHa * (t.area||0) * (1 - model.risk.base) * adopt;
      const annCost = (t.annualCost||0) * (t.area||0);
      const pvB = annBen * annuityFactor(model.time.years, model.time.discBase);
      const pvC = (t.capitalCost||0) + annCost * annuityFactor(model.time.years, model.time.discBase);
      const bcr = pvC>0 ? pvB/pvC : NaN;
      const npv = pvB - pvC;
      return `<tr>
        <td>${esc(t.name)}</td><td>${fmt(t.area)}</td><td>${fmt(adopt)}</td>
        <td>${money(annBen)}</td><td>${money(annCost)}</td>
        <td>${money(pvB)}</td><td>${money(pvC)}</td>
        <td>${isFinite(bcr)?fmt(bcr):"—"}</td><td>${money(npv)}</td>
      </tr>`;
    }).join("");

    const benRows = model.benefits.map(b => `
      <tr><td>${esc(b.label)}</td><td>${b.category}</td><td>${b.frequency}</td>
      <td>${b.startYear||""}</td><td>${b.endYear||""}</td><td>${b.year||""}</td>
      <td>${b.unitValue||""}</td><td>${b.quantity||""}</td><td>${b.abatement||""}</td>
      <td>${b.annualAmount||""}</td><td>${b.growthPct||""}</td>
      <td>${b.linkAdoption?"Yes":"No"}</td><td>${b.linkRisk?"Yes":"No"}</td>
      <td>${b.p0||""}</td><td>${b.p1||""}</td><td>${b.consequence||""}</td>
      <td>${esc(b.notes||"")}</td></tr>
    `).join("");

    const now = new Date().toLocaleString();

    const win = window.open("", "_blank");
    win.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>${esc(s.meta.name)} — CBA Report</title>
      <style>${getPrintCss()}</style></head><body>
      <div class="print-report">
        <div class="print-header">
          <div class="print-logo">🚜🌾</div>
          <div>
            <h1>${esc(s.meta.name)} <span class="print-badge">CBA Report</span></h1>
            <div class="print-small muted">Generated ${esc(now)}</div>
          </div>
        </div>

        <div class="print-cols">
          <div>
            <h2>Project</h2>
            <table>
              <tr><th>Analysts</th><td>${esc(s.meta.analysts)}</td></tr>
              <tr><th>Organisation</th><td>${esc(s.meta.organisation)}</td></tr>
              <tr><th>Contact</th><td><a href="mailto:${esc(s.meta.contact)}">${esc(s.meta.contact)}</a></td></tr>
              <tr><th>Last Updated</th><td>${esc(s.meta.updated)}</td></tr>
            </table>
          </div>
          <div>
            <h2>Parameters</h2>
            <table>
              <tr><th>Start Year</th><td>${s.params.startYear}</td></tr>
              <tr><th>Years</th><td>${s.params.years}</td></tr>
              <tr><th>Discount (L/B/H)</th><td>${s.params.discountLow}% / ${s.params.discountBase}% / ${s.params.discountHigh}%</td></tr>
              <tr><th>MIRR (Finance/Reinvest)</th><td>${s.params.mirrFinance}% / ${s.params.mirrReinvest}%</td></tr>
              <tr><th>Adoption Multiplier</th><td>${s.params.adoptionBase}</td></tr>
              <tr><th>Risk (overall)</th><td>${s.params.riskBase}</td></tr>
              <tr><th>BCR Mode</th><td>${esc(s.params.bcrMode)}</td></tr>
            </table>
          </div>
        </div>

        <h2>Economic Indicators</h2>
        <table>
          <tr><th>PV Benefits</th><td>${money(s.results.pvBenefits)}</td></tr>
          <tr><th>PV Costs</th><td>${money(s.results.pvCosts)}</td></tr>
          <tr><th>NPV</th><td>${money(s.results.npv)}</td></tr>
          <tr><th>BCR</th><td>${isFinite(s.results.bcr)?fmt(s.results.bcr):"—"}</td></tr>
          <tr><th>IRR</th><td>${isFinite(s.results.irrVal)?percent(s.results.irrVal):"—"}</td></tr>
          <tr><th>MIRR</th><td>${isFinite(s.results.mirrVal)?percent(s.results.mirrVal):"—"}</td></tr>
          <tr><th>ROI</th><td>${isFinite(s.results.roi)?percent(s.results.roi):"—"}</td></tr>
          <tr><th>Gross Margin (annual)</th><td>${money(s.results.annualGM)}</td></tr>
          <tr><th>Gross Profit Margin</th><td>${isFinite(s.results.profitMargin)?percent(s.results.profitMargin):"—"}</td></tr>
          <tr><th>Payback (years)</th><td>${s.results.paybackYears ?? "Not reached"}</td></tr>
        </table>

        <h2>Treatments</h2>
        <table>
          <thead><tr>
            <th>Treatment</th><th>Area</th><th>Adoption</th><th>Annual Benefit</th><th>Annual Cost</th>
            <th>PV Benefit</th><th>PV Cost</th><th>BCR</th><th>NPV</th>
          </tr></thead>
          <tbody>${trRows}</tbody>
        </table>

        <h2>Additional Benefits</h2>
        <table>
          <thead><tr>
            <th>Label</th><th>Cat</th><th>Freq</th><th>Start</th><th>End</th><th>Year</th>
            <th>UnitValue</th><th>Qty</th><th>Abatement</th><th>Annual</th><th>Growth%</th>
            <th>Adopt?</th><th>Risk?</th><th>P0</th><th>P1</th><th>Consequence</th><th>Notes</th>
          </tr></thead>
          <tbody>${benRows}</tbody>
        </table>

        <h2>Simulation Highlights</h2>
        <div class="print-cols">
          <div>${npvImg ? `<img src="${npvImg}" style="width:100%;border:1px solid #ddd;border-radius:8px" />` : "<div class='muted'>NPV histogram not available.</div>"}</div>
          <div>${bcrImg ? `<img src="${bcrImg}" style="width:100%;border:1px solid #ddd;border-radius:8px" />` : "<div class='muted'>BCR histogram not available.</div>"}</div>
        </div>

        <hr />
        <div class="print-small muted">
          Newcastle Business School • The University of Newcastle • Contact: <a href="mailto:${esc(model.project.contactEmail)}">${esc(model.project.contactEmail)}</a>
        </div>
      </div>
    </body></html>`);
    win.document.close();
    setTimeout(() => { win.focus(); win.print(); }, 300);
  }

  function getPrintCss() {
    return `
      body{background:#fff;margin:0}
      .print-report{font:13px/1.5 system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#111;padding:24px;max-width:900px;margin:0 auto}
      .print-header{display:grid;grid-template-columns:60px 1fr;gap:12px;align-items:center;margin-bottom:6px}
      .print-logo{width:60px;height:60px;border-radius:14px;display:grid;place-items:center;font-size:26px;
        background:linear-gradient(135deg,#9be2ad,#ffd77e);border:1px solid #eee}
      .print-report h1{font-size:20px;margin:0 0 6px}
      .print-report h2{font-size:16px;margin:14px 0 6px}
      .print-report table{border-collapse:collapse;width:100%;margin:6px 0}
      .print-report th,.print-report td{border:1px solid #ddd;padding:6px;text-align:left}
      .print-report .muted{color:#555}
      .print-cols{display:grid;grid-template-columns:1fr 1fr;gap:12px}
      .print-small{font-size:12px}
      .print-badge{display:inline-block;border:1px solid #ddd;border-radius:8px;padding:2px 8px;margin-left:6px}
      @media print {.print-report img{page-break-inside:avoid}}
    `;
  }

  // ---------- EXCEL: TEMPLATE + PARSE + SAMPLE ----------
  let parsedExcel = null;

  async function handleParseExcel() {
    const file = $("#excelFile").files?.[0];
    const status = $("#loadStatus");
    const alertBox = $("#validation");
    const preview = $("#preview");
    parsedExcel = null;
    alertBox.classList.remove("show");
    alertBox.innerHTML = "";
    preview.innerHTML = "";
    $("#importExcel").disabled = true;

    if (!file) { status.textContent = "Select an Excel/CSV file first."; return; }
    status.textContent = "Parsing…";

    try {
      const buf = await file.arrayBuffer();
      let wb;
      if (file.name.toLowerCase().endsWith(".csv")) {
        const csvTxt = new TextDecoder().decode(new Uint8Array(buf));
        const ws = XLSX.utils.csv_to_sheet(csvTxt);
        wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, $("#csvSheetType").value || "Outputs");
      } else {
        wb = XLSX.read(buf, { type: "array" });
      }

      const getSheet = (name) => {
        const s = wb.Sheets[name];
        if (!s) return null;
        return XLSX.utils.sheet_to_json(s, { defval: "", raw: true });
      };

      const outputs = getSheet("Outputs") || [];
      const treatments = getSheet("Treatments") || [];
      const otherCosts = getSheet("OtherCosts") || [];
      const time = getSheet("Time") || [];
      const adoptRisk = getSheet("AdoptionRisk") || [];
      const project = getSheet("Project") || [];
      const benefits = getSheet("Benefits") || [];

      const issues = [];
      const needCols = (rows, cols, label) => {
        if (!rows.length) { issues.push(`Sheet <b>${label}</b> is empty.`); return; }
        const lower = Object.keys(rows[0]).map(k => k.toLowerCase());
        const miss = cols.filter(c => !lower.includes(c.toLowerCase()));
        if (miss.length) issues.push(`Missing in <b>${label}</b>: ${miss.join(", ")}`);
      };

      if (!outputs.length) issues.push("Sheet <b>Outputs</b> is missing.");
      else needCols(outputs, ["Name","Unit","ValuePerUnit","Source"], "Outputs");

      if (!treatments.length) issues.push("Sheet <b>Treatments</b> is missing.");
      else needCols(treatments, ["Name","Area","Adoption","AnnualCostPerHa","CapitalCostY0","Constrained","Source"], "Treatments");

      if (otherCosts.length) needCols(otherCosts, ["Label","Type","Constrained"], "OtherCosts");
      if (time.length) needCols(time, ["StartYear","Years","DiscLow","DiscBase","DiscHigh","MIRR_Finance","MIRR_Reinvest"], "Time");
      if (adoptRisk.length) needCols(adoptRisk, ["AdoptLow","AdoptBase","AdoptHigh","RiskLow","RiskBase","RiskHigh"], "AdoptionRisk");
      if (benefits.length) needCols(benefits, ["Label","Category","Frequency","StartYear","EndYear","Year","UnitValue","Quantity","Abatement","AnnualAmount","GrowthPct","LinkAdoption","LinkRisk","P0","P1","Consequence","Notes"], "Benefits");

      const showTable = (title, rows) => {
        if (!rows?.length) return "";
        const cols = Object.keys(rows[0]);
        const head = cols.map(c => `<th>${esc(c)}</th>`).join("");
        const body = rows.slice(0, 10).map(r => `<tr>${cols.map(c=>`<td>${esc(r[c])}</td>`).join("")}</tr>`).join("");
        return `<h4>${esc(title)} (${rows.length} rows)</h4>
          <div class="preview"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
      };

      $("#preview").innerHTML =
        showTable("Outputs", outputs) +
        showTable("Treatments", treatments) +
        (otherCosts.length ? showTable("OtherCosts", otherCosts) : "") +
        (time.length ? showTable("Time", time) : "") +
        (adoptRisk.length ? showTable("AdoptionRisk", adoptRisk) : "") +
        (benefits.length ? showTable("Benefits", benefits) : "") +
        (project.length ? showTable("Project", project) : "");

      if (issues.length) {
        alertBox.innerHTML = `<div class="title">Please fix the following:</div><ul>${issues.map(i=>`<li>${i}</li>`).join("")}</ul>`;
        alertBox.classList.add("show");
        status.textContent = "Parsed with issues.";
      } else {
        status.textContent = "Parsed successfully. Review preview, then click “Import to Tool”.";
        parsedExcel = { outputs, treatments, otherCosts, time, adoptRisk, project, benefits };
        $("#importExcel").disabled = false;
      }
    } catch (err) {
      $("#validation").innerHTML = `<div class="title">Error</div><div>${esc(err.message || String(err))}</div>`;
      $("#validation").classList.add("show");
      $("#loadStatus").textContent = "Parsing failed.";
    }
  }

  function commitExcelToModel() {
    if (!parsedExcel) return;
    const { outputs, treatments, otherCosts, time, adoptRisk, project, benefits } = parsedExcel;

    const outMap = new Map();
    model.outputs = outputs.map(r => {
      const id = uid();
      outMap.set(String(r.Name).toLowerCase(), id);
      return {
        id,
        name: String(r.Name || "").trim() || "Output",
        unit: String(r.Unit || "").trim() || "unit",
        value: +r.ValuePerUnit || 0,
        source: String(r.Source || "Input Directly")
      };
    });

    model.treatments = treatments.map(r => {
      const t = {
        id: uid(),
        name: String(r.Name || "Treatment"),
        area: +r.Area || 0,
        adoption: clamp(+r.Adoption || 0, 0, 1),
        annualCost: +r.AnnualCostPerHa || 0,
        capitalCost: +r.CapitalCostY0 || 0,
        constrained: String(r.Constrained || "Yes").toLowerCase().startsWith("y") || String(r.Constrained).toLowerCase() === "true",
        source: String(r.Source || "Input Directly"),
        deltas: {}
      };
      Object.keys(r).forEach(k => {
        if (k.toLowerCase().startsWith("delta:")) {
          const oname = k.slice(6).trim().toLowerCase();
          if (!outMap.has(oname)) {
            const id = uid();
            model.outputs.push({ id, name: k.slice(6).trim(), unit: "unit", value: 0, source: "Input Directly" });
            outMap.set(oname, id);
          }
          const oid = outMap.get(oname);
          t.deltas[oid] = +r[k] || 0;
        }
      });
      model.outputs.forEach(o => { if (!(o.id in t.deltas)) t.deltas[o.id] = 0; });
      return t;
    });

    model.otherCosts = (otherCosts || []).map(r => ({
      id: uid(),
      label: String(r.Label || "Cost"),
      type: (String(r.Type || "Annual").toLowerCase().startsWith("c") ? "capital" : "annual"),
      annual: +r.Annual || 0,
      startYear: +r.StartYear || model.time.startYear,
      endYear: +r.EndYear || model.time.startYear,
      capital: +r.Capital || 0,
      year: +r.Year || model.time.startYear,
      constrained: String(r.Constrained || "Yes").toLowerCase().startsWith("y") || String(r.Constrained).toLowerCase() === "true"
    }));

    if ((time || []).length) {
      const t = time[0];
      model.time.startYear = +t.StartYear || model.time.startYear;
      model.time.years = +t.Years || model.time.years;
      model.time.discLow = +t.DiscLow || model.time.discLow;
      model.time.discBase = +t.DiscBase || model.time.discBase;
      model.time.discHigh = +t.DiscHigh || model.time.discHigh;
      model.time.mirrFinance = +t.MIRR_Finance || model.time.mirrFinance;
      model.time.mirrReinvest = +t.MIRR_Reinvest || model.time.mirrReinvest;
    }

    if ((adoptRisk || []).length) {
      const a = adoptRisk[0];
      model.adoption.low = clamp(+a.AdoptLow || model.adoption.low, 0, 1);
      model.adoption.base = clamp(+a.AdoptBase || model.adoption.base, 0, 1);
      model.adoption.high = clamp(+a.AdoptHigh || model.adoption.high, 0, 1);
      model.risk.low = clamp(+a.RiskLow || model.risk.low, 0, 1);
      model.risk.base = clamp(+a.RiskBase || model.risk.base, 0, 1);
      model.risk.high = clamp(+a.RiskHigh || model.risk.high, 0, 1);
      if ("Tech" in a) model.risk.tech = clamp(+a.Tech || model.risk.tech, 0, 1);
      if ("NonCoop" in a) model.risk.nonCoop = clamp(+a.NonCoop || model.risk.nonCoop, 0, 1);
      if ("Socio" in a) model.risk.socio = clamp(+a.Socio || model.risk.socio, 0, 1);
      if ("Fin" in a) model.risk.fin = clamp(+a.Fin || model.risk.fin, 0, 1);
      if ("Man" in a) model.risk.man = clamp(+a.Man || model.risk.man, 0, 1);
    }

    if ((project || []).length) {
      const p = project[0];
      model.project.name = String(p.Name || model.project.name);
      model.project.analysts = String(p.Analysts || model.project.analysts);
      model.project.organisation = String(p.Organisation || model.project.organisation);
      model.project.contactEmail = String(p.ContactEmail || model.project.contactEmail);
      model.project.summary = String(p.Summary || model.project.summary);
      model.project.goal = String(p.Goal || model.project.goal);
      model.project.withProject = String(p.WithProject || model.project.withProject);
      model.project.withoutProject = String(p.WithoutProject || model.project.withoutProject);
      model.project.lastUpdated = String(p.LastUpdated || model.project.lastUpdated);
    }

    if ((benefits || []).length) {
      model.benefits = benefits.map(r => ({
        id: uid(),
        label: String(r.Label || "Benefit"),
        category: String(r.Category || "C4").toUpperCase(),
        frequency: String(r.Frequency || "Annual"),
        startYear: +r.StartYear || model.time.startYear,
        endYear: +r.EndYear || model.time.startYear,
        year: +r.Year || model.time.startYear,
        unitValue: +r.UnitValue || 0,
        quantity: +r.Quantity || 0,
        abatement: +r.Abatement || 0,
        annualAmount: +r.AnnualAmount || 0,
        growthPct: +r.GrowthPct || 0,
        linkAdoption: String(r.LinkAdoption||"Yes").toLowerCase().startsWith("y") || String(r.LinkAdoption).toLowerCase()==="true",
        linkRisk: String(r.LinkRisk||"Yes").toLowerCase().startsWith("y") || String(r.LinkRisk).toLowerCase()==="true",
        p0: +r.P0 || 0,
        p1: +r.P1 || 0,
        consequence: +r.Consequence || 0,
        notes: String(r.Notes || "")
      }));
    }

    initTreatmentDeltas();
    renderAll();
    setBasicsFieldsFromModel();
    calcAndRender();

    $("#loadStatus").textContent = "Imported into the tool.";
    $("#importExcel").disabled = true;
    switchTab("report");
  }

  function downloadExcelTemplate() {
    const status = $("#loadStatus");
    status.textContent = "Building Excel template…";
    if (!window.XLSX) { alert("Excel library not loaded."); status.textContent = ""; return; }

    const wb = XLSX.utils.book_new();

    const outputs = [
      ["Name","Unit","ValuePerUnit","Source"],
      ["Yield","t/ha",300,"Input Directly"],
      ["Protein","%-point",12,"ABARES"],
      ["Moisture","%-point",-5,"GRDC"],
      ["Biomass","t/ha",40,"Farm Trials"]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(outputs), "Outputs");

    const tr = [
      ["Name","Area","Adoption","AnnualCostPerHa","CapitalCostY0","Constrained","Source","Delta:Yield","Delta:Protein","Delta:Moisture","Delta:Biomass"],
      ["Optimized N (Rate+Timing)",300,0.8,45,5000,"Yes","Farm Trials",0.2,0.5,0,0.3],
      ["Slow-Release N",200,0.7,25,0,"Yes","ABARES",0.1,0.2,0,0.1]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(tr), "Treatments");

    const oc = [
      ["Label","Type","Annual","StartYear","EndYear","Capital","Year","Constrained"],
      ["Project Mgmt & M&E","Annual",20000,new Date().getFullYear(),new Date().getFullYear()+4,"","", "Yes"]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(oc), "OtherCosts");

    const time = [
      ["StartYear","Years","DiscLow","DiscBase","DiscHigh","MIRR_Finance","MIRR_Reinvest"],
      [new Date().getFullYear(),10,4,7,10,6,4]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(time), "Time");

    const ar = [
      ["AdoptLow","AdoptBase","AdoptHigh","RiskLow","RiskBase","RiskHigh","Tech","NonCoop","Socio","Fin","Man"],
      [0.6,0.9,1.0,0.05,0.15,0.3,0.05,0.04,0.02,0.03,0.02]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ar), "AdoptionRisk");

    const proj = [
      ["Name","Analysts","Organisation","ContactEmail","Summary","Goal","WithProject","WithoutProject","LastUpdated"],
      ["Nitrogen Optimization Trial","Farm Econ Team","Newcastle Business School, The University of Newcastle","frank.agbola@newcastle.edu.au","Test fertilizer strategies to raise wheat yield and protein across 500 ha over 5 years.","Increase yield by 10% and protein by 0.5 p.p. on 500 ha within 3 years.","Adopt optimized nitrogen timing and rates; improved management on 500 ha.","Business-as-usual fertilization; yield/protein unchanged; rising costs.", new Date().toISOString().slice(0,10)]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(proj), "Project");

    const ben = [
      ["Label","Category","Frequency","StartYear","EndYear","Year","UnitValue","Quantity","Abatement","AnnualAmount","GrowthPct","LinkAdoption","LinkRisk","P0","P1","Consequence","Notes"],
      ["Reduced recurring costs (energy/water)","C4","Annual",new Date().getFullYear(),new Date().getFullYear()+4,"","","","",15000,0,"Yes","Yes","","","","Project-wide OPEX saving"],
      ["Reduced risk of quality downgrades","C7","Annual",new Date().getFullYear(),new Date().getFullYear()+9,"","","","",0,0,"Yes","No",0.10,0.07,120000,""],
      ["Soil asset value uplift (carbon/structure)","C6","Once","","",new Date().getFullYear()+5,"","","",50000,0,"No","Yes","","","",""]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ben), "Benefits");

    saveWorkbook("CBA_Farming_Template.xlsx", wb);
    status.textContent = "Template download started.";
  }

  function downloadSampleDataset() {
    const status = $("#loadStatus");
    status.textContent = "Building sample dataset…";
    if (!window.XLSX) { alert("Excel library not loaded."); status.textContent = ""; return; }

    const wb = XLSX.utils.book_new();

    const outputs = [
      ["Name","Unit","ValuePerUnit","Source"],
      ["Yield","t/ha",330,"ABARES"],
      ["Protein","%-point",15,"GRDC"],
      ["Moisture","%-point",-6,"GRDC"],
      ["Biomass","t/ha",38,"Farm Trials"]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(outputs), "Outputs");

    const tr = [
      ["Name","Area","Adoption","AnnualCostPerHa","CapitalCostY0","Constrained","Source","Delta:Yield","Delta:Protein","Delta:Moisture","Delta:Biomass"],
      ["Optimized N (Rate+Timing)",300,0.8,45,5000,"Yes","Farm Trials",0.22,0.5,0,0.3],
      ["Slow-Release N",200,0.7,25,0,"Yes","ABARES",0.12,0.2,0,0.1],
      ["Foliar Micronutrients",150,0.6,18,1500,"No","Plant Farm",0.08,0.15,0,0.05],
      ["Improved Irrigation Scheduling",100,0.5,30,8000,"Yes","Farm Trials",0.1,0.1,-0.2,0.2]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(tr), "Treatments");

    const oc = [
      ["Label","Type","Annual","StartYear","EndYear","Capital","Year","Constrained"],
      ["Project Mgmt & M&E","Annual",22000,new Date().getFullYear(),new Date().getFullYear()+4,"","", "Yes"],
      ["Baseline Monitoring","Annual",8000,new Date().getFullYear(),new Date().getFullYear()+2,"","", "No"],
      ["Weather Station Upgrade","Capital","","","","12000",new Date().getFullYear(),"No"]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(oc), "OtherCosts");

    const time = [
      ["StartYear","Years","DiscLow","DiscBase","DiscHigh","MIRR_Finance","MIRR_Reinvest"],
      [new Date().getFullYear(),12,4,7,10,6,4]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(time), "Time");

    const ar = [
      ["AdoptLow","AdoptBase","AdoptHigh","RiskLow","RiskBase","RiskHigh","Tech","NonCoop","Socio","Fin","Man"],
      [0.5,0.85,1.0,0.04,0.14,0.28,0.05,0.04,0.03,0.03,0.02]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ar), "AdoptionRisk");

    const proj = [
      ["Name","Analysts","Organisation","ContactEmail","Summary","Goal","WithProject","WithoutProject","LastUpdated"],
      ["Integrated Nutrient & Water Strategy","Ag Econ Team","Newcastle Business School, The University of Newcastle","frank.agbola@newcastle.edu.au","Evaluate multiple nutrient and irrigation practices across 750 ha.","Lift gross margin by ≥8% with NPV > 0 within 5 years.","Implement four practices with staged adoption and monitoring.","Maintain current practices; incremental yield trends only.", new Date().toISOString().slice(0,10)]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(proj), "Project");

    const ben = [
      ["Label","Category","Frequency","StartYear","EndYear","Year","UnitValue","Quantity","Abatement","AnnualAmount","GrowthPct","LinkAdoption","LinkRisk","P0","P1","Consequence","Notes"],
      ["Reduced diesel use (field passes)","C4","Annual",new Date().getFullYear(),new Date().getFullYear()+6,"","","","",12000,0,"Yes","Yes","","","",""],
      ["Reduced risk of crop downgrades","C7","Annual",new Date().getFullYear(),new Date().getFullYear()+11,"","","","",0,0,"Yes","No",0.12,0.08,150000,""],
      ["Soil organic matter uplift","C6","Once","","",new Date().getFullYear()+7,"","","",80000,0,"No","Yes","","","",""]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ben), "Benefits");

    saveWorkbook("CBA_Farming_Sample.xlsx", wb);
    status.textContent = "Sample dataset download started.";
  }

  // ---------- INIT ----------
  function init() {
    initTreatmentDeltas();
    initTabs();
    initActions();
    bindBasics();
    renderAll();
    calcAndRender();
    // Start on Cover tab by default (HTML already shows it)
  }
  document.addEventListener("DOMContentLoaded", init);
})();
