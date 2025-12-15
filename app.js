// Farming CBA Tool - Newcastle Business School
// Fully upgraded script with working tabs, CBA, simulation, Excel import/exports, and Copilot helper

(() => {
  "use strict";

  // ---------- CONSTANTS ----------
  const DEFAULT_DISCOUNT_SCHEDULE = [
    { label: "2025-2034", from: 2025, to: 2034, low: 2, base: 4, high: 6 },
    { label: "2035-2044", from: 2035, to: 2044, low: 4, base: 7, high: 10 },
    { label: "2045-2054", from: 2045, to: 2054, low: 4, base: 7, high: 10 },
    { label: "2055-2064", from: 2055, to: 2064, low: 3, base: 6, high: 9 },
    { label: "2065-2074", from: 2065, to: 2074, low: 2, base: 5, high: 8 }
  ];

  const horizons = [5, 10, 15, 20, 25];

  // ---------- MODEL ----------
  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

  const model = {
    project: {
      name: "Nitrogen Optimization Trial",
      lead: "Project Lead",
      analysts: "Farm Econ Team",
      team: "",
      organisation: "Newcastle Business School, The University of Newcastle",
      contactEmail: "frank.agbola@newcastle.edu.au",
      contactPhone: "",
      summary:
        "Test fertilizer strategies to raise wheat yield and protein across 500 ha over 5 years.",
      objectives: "",
      activities: "",
      stakeholders: "",
      lastUpdated: new Date().toISOString().slice(0, 10),
      goal: "Increase yield by 10% and protein by 0.5 percentage points on 500 ha within 3 years.",
      withProject: "Adopt optimized nitrogen timing and rates with improved management over 500 ha.",
      withoutProject:
        "Business as usual fertilization with unchanged yield and protein and rising costs."
    },
    time: {
      startYear: new Date().getFullYear(),
      projectStartYear: new Date().getFullYear(),
      years: 10,
      discBase: 7,
      discLow: 4,
      discHigh: 10,
      mirrFinance: 6,
      mirrReinvest: 4,
      discountSchedule: JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE))
    },
    outputsMeta: {
      systemType: "single",
      assumptions: ""
    },
    outputs: [
      { id: uid(), name: "Yield", unit: "t/ha", value: 300, source: "Input Directly" },
      { id: uid(), name: "Protein", unit: "percentage point", value: 12, source: "Input Directly" },
      { id: uid(), name: "Moisture", unit: "percentage point", value: -5, source: "Input Directly" },
      { id: uid(), name: "Biomass", unit: "t/ha", value: 40, source: "Input Directly" }
    ],
    treatments: [
      {
        id: uid(),
        name: "Optimized N (Rate + Timing)",
        area: 300,
        adoption: 1,
        deltas: {},
        labourCost: 45,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: 5000,
        constrained: true,
        source: "Farm Trials",
        isControl: false,
        notes: ""
      },
      {
        id: uid(),
        name: "Slow release N",
        area: 200,
        adoption: 1,
        deltas: {},
        labourCost: 25,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "ABARES",
        isControl: false,
        notes: ""
      }
    ],
    benefits: [
      {
        id: uid(),
        label: "Reduced recurring costs (energy and water)",
        category: "C4",
        theme: "Cost savings",
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
        notes: "Project wide operating cost saving"
      },
      {
        id: uid(),
        label: "Reduced risk of quality downgrades",
        category: "C7",
        theme: "Risk reduction",
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
        label: "Soil asset value uplift (carbon and structure)",
        category: "C6",
        theme: "Soil carbon",
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
        label: "Project management and monitoring and evaluation",
        type: "annual",
        category: "Capital",
        annual: 20000,
        startYear: new Date().getFullYear(),
        endYear: new Date().getFullYear() + 4,
        capital: 50000,
        year: new Date().getFullYear(),
        constrained: true,
        depMethod: "declining",
        depLife: 5,
        depRate: 30
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
      details: [],
      variationPct: 20,
      varyOutputs: true,
      varyTreatCosts: true,
      varyInputCosts: false
    }
  };

  let parsedExcel = null;

  function initTreatmentDeltas() {
    model.treatments.forEach(t => {
      model.outputs.forEach(o => {
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      });
      // Backward compatibility mapping for labour cost and adoption
      if (typeof t.labourCost === "undefined") {
        t.labourCost = Number(t.annualCost || 0) || 0;
      }
      if (typeof t.materialsCost === "undefined") t.materialsCost = 0;
      if (typeof t.servicesCost === "undefined") t.servicesCost = 0;
      if (typeof t.adoption !== "number" || isNaN(t.adoption)) t.adoption = 1;
      delete t.annualCost;
    });
  }
  initTreatmentDeltas();

  // ---------- UTIL ----------
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const fmt = n =>
    isFinite(n)
      ? Math.abs(n) >= 1000
        ? n.toLocaleString(undefined, { maximumFractionDigits: 0 })
        : n.toLocaleString(undefined, { maximumFractionDigits: 2 })
      : "n/a";
  const money = n => (isFinite(n) ? "$" + fmt(n) : "n/a");
  const percent = n => (isFinite(n) ? fmt(n) + "%" : "n/a");
  const slug = s =>
    (s || "project")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_|_$/g, "");
  const annuityFactor = (N, rPct) => {
    const r = rPct / 100;
    return r === 0 ? N : (1 - Math.pow(1 + r, -N)) / r;
  };
  const esc = s =>
    (s ?? "")
      .toString()
      .replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  function rng(seed) {
    let t = (seed || Math.floor(Math.random() * 2 ** 31)) >>> 0;
    return () => {
      t += 0x6d2b79f5;
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

  function showToast(message) {
    const root = document.getElementById("toast-root") || document.body;
    const toast = document.createElement("div");
    toast.className = "toast";
    toast.textContent = message;
    root.appendChild(toast);
    void toast.offsetWidth;
    toast.classList.add("show");
    setTimeout(() => {
      toast.classList.remove("show");
      setTimeout(() => toast.remove(), 200);
    }, 3500);
  }

  // ---------- CASHFLOWS ----------
  function additionalBenefitsSeries(N, baseYear, adoptMul, risk) {
    const series = new Array(N + 1).fill(0);
    model.benefits.forEach(b => {
      const cat = String(b.category || "").toUpperCase();
      const linkA = !!b.linkAdoption;
      const linkR = !!b.linkRisk;
      const A = linkA ? clamp(adoptMul, 0, 1) : 1;
      const R = linkR ? 1 - clamp(risk, 0, 1) : 1;
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
            const p0 = Number(b.p0) || 0;
            const p1 = Number(b.p1) || 0;
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

  function buildCashflows({ forRate = model.time.discBase, adoptMul = model.adoption.base, risk = model.risk.base }) {
    const N = model.time.years;
    const baseYear = model.time.startYear;

    const benefitByYear = new Array(N + 1).fill(0);
    const costByYear = new Array(N + 1).fill(0);
    const constrainedCostByYear = new Array(N + 1).fill(0);

    let annualBenefit = 0;
    let treatAnnualCost = 0;
    let treatConstrAnnualCost = 0;
    let treatCapitalY0 = 0;
    let treatConstrCapitalY0 = 0;

    model.treatments.forEach(t => {
      const adopt = clamp(adoptMul, 0, 1);
      let valuePerHa = 0;
      model.outputs.forEach(o => {
        const delta = Number(t.deltas[o.id]) || 0;
        const v = Number(o.value) || 0;
        valuePerHa += delta * v;
      });
      const area = Number(t.area) || 0;
      const benefit = valuePerHa * area * (1 - clamp(risk, 0, 1)) * adopt;

      const annualCostPerHa =
        (Number(t.materialsCost) || 0) +
        (Number(t.servicesCost) || 0) +
        (Number(t.labourCost) || 0);
      const opCost = annualCostPerHa * area;
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

    const otherAnnualByYear = new Array(N + 1).fill(0);
    const otherConstrAnnualByYear = new Array(N + 1).fill(0);
    let otherCapitalY0 = 0;
    let otherConstrCapitalY0 = 0;

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

    const extra = additionalBenefitsSeries(N, baseYear, adoptMul, risk);
    for (let i = 0; i < extra.length; i++) benefitByYear[i] += extra[i];

    const cf = new Array(N + 1).fill(0).map((_, i) => benefitByYear[i] - costByYear[i]);
    const annualGM = annualBenefit - treatAnnualCost;
    return { benefitByYear, costByYear, constrainedCostByYear, cf, annualGM };
  }

  function presentValue(series, ratePct) {
    let pv = 0;
    for (let t = 0; t < series.length; t++) {
      pv += series[t] / Math.pow(1 + ratePct / 100, t);
    }
    return pv;
  }

  function irr(cf) {
    const hasPos = cf.some(v => v > 0);
    const hasNeg = cf.some(v => v < 0);
    if (!hasPos || !hasNeg) return NaN;
    let lo = -0.99;
    let hi = 5.0;
    const npvAt = r => cf.reduce((acc, v, t) => acc + v / Math.pow(1 + r, t), 0);
    let nLo = npvAt(lo);
    let nHi = npvAt(hi);
    if (nLo * nHi > 0) {
      for (let k = 0; k < 20 && nLo * nHi > 0; k++) {
        hi *= 1.5;
        nHi = npvAt(hi);
      }
      if (nLo * nHi > 0) return NaN;
    }
    for (let i = 0; i < 80; i++) {
      const mid = (lo + hi) / 2;
      const nMid = npvAt(mid);
      if (Math.abs(nMid) < 1e-8) return mid * 100;
      if (nLo * nMid <= 0) {
        hi = mid;
        nHi = nMid;
      } else {
        lo = mid;
        nLo = nMid;
      }
    }
    return ((lo + hi) / 2) * 100;
  }

  function mirr(cf, financeRatePct, reinvestRatePct) {
    const n = cf.length - 1;
    const fr = financeRatePct / 100;
    const rr = reinvestRatePct / 100;
    let pvNeg = 0;
    let fvPos = 0;
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

  function computeAll(rate, adoptMul, risk, bcrMode) {
    const { benefitByYear, costByYear, constrainedCostByYear, cf, annualGM } =
      buildCashflows({ forRate: rate, adoptMul, risk });
    const pvBenefits = presentValue(benefitByYear, rate);
    const pvCosts = presentValue(costByYear, rate);
    const pvCostsConstrained = presentValue(constrainedCostByYear, rate);

    const npv = pvBenefits - pvCosts;
    const denom = bcrMode === "constrained" ? pvCostsConstrained : pvCosts;
    const bcr = denom > 0 ? pvBenefits / denom : NaN;

    const irrVal = irr(cf);
    const mirrVal = mirr(cf, model.time.mirrFinance, model.time.mirrReinvest);
    const roi = pvCosts > 0 ? ((pvBenefits - pvCosts) / pvCosts) * 100 : NaN;
    const profitMargin = benefitByYear[1] > 0 ? (annualGM / benefitByYear[1]) * 100 : NaN;
    const pb = payback(cf, rate);

    return {
      pvBenefits,
      pvCosts,
      pvCostsConstrained,
      npv,
      bcr,
      irrVal,
      mirrVal,
      roi,
      annualGM,
      profitMargin,
      paybackYears: pb,
      cf,
      benefitByYear,
      costByYear
    };
  }

  // ---------- DOM HELPERS ----------
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));
  const num = sel => +(document.querySelector(sel)?.value || 0);
  const setVal = (sel, text) => {
    const el = document.querySelector(sel);
    if (el) el.textContent = text;
  };

  // ---------- TABS ----------
  function switchTab(target) {
    if (!target) return;

    const navEls = $$("[data-tab],[data-tab-target],[data-tab-jump]");
    navEls.forEach(el => {
      const key = el.dataset.tab || el.dataset.tabTarget || el.dataset.tabJump;
      const isActive = key === target;
      el.classList.toggle("active", isActive);
      el.setAttribute("aria-selected", isActive ? "true" : "false");
    });

    const panels = $$(".tab-panel");
    panels.forEach(p => {
      const key = p.dataset.tabPanel || (p.id ? p.id.replace(/^tab-/, "") : "");
      const match =
        key === target ||
        p.id === target ||
        p.id === "tab-" + target;

      const show = !!match;
      p.classList.toggle("active", show);
      p.classList.toggle("show", show);
      p.hidden = !show;
      p.setAttribute("aria-hidden", show ? "false" : "true");
      p.style.display = show ? "" : "none";
    });

    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function initTabs() {
    document.addEventListener("click", e => {
      const el = e.target.closest("[data-tab],[data-tab-target],[data-tab-jump]");
      if (!el) return;
      const target =
        el.dataset.tab ||
        el.dataset.tabTarget ||
        el.dataset.tabJump;
      if (!target) return;
      e.preventDefault();
      switchTab(target);
      showToast(`Switched to ${target} tab.`);
    });

    const activeNav =
      document.querySelector("[data-tab].active, [data-tab-target].active, [data-tab-jump].active") ||
      document.querySelector("[data-tab], [data-tab-target], [data-tab-jump]");
    if (activeNav) {
      const target =
        activeNav.dataset.tab ||
        activeNav.dataset.tabTarget ||
        activeNav.dataset.tabJump;
      if (target) {
        switchTab(target);
        return;
      }
    }

    const firstPanel = document.querySelector(".tab-panel");
    if (firstPanel) {
      const key =
        firstPanel.dataset.tabPanel ||
        (firstPanel.id ? firstPanel.id.replace(/^tab-/, "") : "");
      if (key) switchTab(key);
    }
  }

  function initActions() {
    document.addEventListener("click", e => {
      const el = e.target.closest("#recalc, #getResults, [data-action='recalc']");
      if (!el) return;
      e.preventDefault();
      e.stopPropagation();
      calcAndRender();
      showToast("Base case economic indicators recalculated.");
    });

    document.addEventListener("click", e => {
      const el = e.target.closest("#runSim, [data-action='run-sim']");
      if (!el) return;
      e.preventDefault();
      e.stopPropagation();
      runSimulation();
    });
  }

  // ---------- BIND + RENDER FORMS ----------
  function setBasicsFieldsFromModel() {
    if ($("#projectName")) $("#projectName").value = model.project.name || "";
    if ($("#projectLead")) $("#projectLead").value = model.project.lead || "";
    if ($("#analystNames")) $("#analystNames").value = model.project.analysts || "";
    if ($("#projectTeam")) $("#projectTeam").value = model.project.team || "";
    if ($("#projectSummary")) $("#projectSummary").value = model.project.summary || "";
    if ($("#projectObjectives")) $("#projectObjectives").value = model.project.objectives || "";
    if ($("#projectActivities")) $("#projectActivities").value = model.project.activities || "";
    if ($("#stakeholderGroups")) $("#stakeholderGroups").value = model.project.stakeholders || "";
    if ($("#lastUpdated")) $("#lastUpdated").value = model.project.lastUpdated || "";
    if ($("#projectGoal")) $("#projectGoal").value = model.project.goal || "";
    if ($("#withProject")) $("#withProject").value = model.project.withProject || "";
    if ($("#withoutProject")) $("#withoutProject").value = model.project.withoutProject || "";
    if ($("#organisation")) $("#organisation").value = model.project.organisation || "";
    if ($("#contactEmail")) $("#contactEmail").value = model.project.contactEmail || "";
    if ($("#contactPhone")) $("#contactPhone").value = model.project.contactPhone || "";

    if ($("#startYear")) $("#startYear").value = model.time.startYear;
    if ($("#projectStartYear")) $("#projectStartYear").value = model.time.projectStartYear || model.time.startYear;
    if ($("#years")) $("#years").value = model.time.years;
    if ($("#discBase")) $("#discBase").value = model.time.discBase;
    if ($("#discLow")) $("#discLow").value = model.time.discLow;
    if ($("#discHigh")) $("#discHigh").value = model.time.discHigh;
    if ($("#mirrFinance")) $("#mirrFinance").value = model.time.mirrFinance;
    if ($("#mirrReinvest")) $("#mirrReinvest").value = model.time.mirrReinvest;

    if ($("#adoptBase")) $("#adoptBase").value = model.adoption.base;
    if ($("#adoptLow")) $("#adoptLow").value = model.adoption.low;
    if ($("#adoptHigh")) $("#adoptHigh").value = model.adoption.high;

    if ($("#riskBase")) $("#riskBase").value = model.risk.base;
    if ($("#riskLow")) $("#riskLow").value = model.risk.low;
    if ($("#riskHigh")) $("#riskHigh").value = model.risk.high;
    if ($("#rTech")) $("#rTech").value = model.risk.tech;
    if ($("#rNonCoop")) $("#rNonCoop").value = model.risk.nonCoop;
    if ($("#rSocio")) $("#rSocio").value = model.risk.socio;
    if ($("#rFin")) $("#rFin").value = model.risk.fin;
    if ($("#rMan")) $("#rMan").value = model.risk.man;

    if ($("#simN")) $("#simN").value = model.sim.n;
    if ($("#targetBCR")) $("#targetBCR").value = model.sim.targetBCR;
    if ($("#bcrMode")) $("#bcrMode").value = model.sim.bcrMode;
    if ($("#simBcrTargetLabel")) $("#simBcrTargetLabel").textContent = model.sim.targetBCR;

    if ($("#simVarPct")) $("#simVarPct").value = String(model.sim.variationPct || 20);
    if ($("#simVaryOutputs")) $("#simVaryOutputs").value = model.sim.varyOutputs ? "true" : "false";
    if ($("#simVaryTreatCosts")) $("#simVaryTreatCosts").value = model.sim.varyTreatCosts ? "true" : "false";
    if ($("#simVaryInputCosts")) $("#simVaryInputCosts").value = model.sim.varyInputCosts ? "true" : "false";

    if ($("#systemType")) $("#systemType").value = model.outputsMeta.systemType || "single";
    if ($("#outputAssumptions")) $("#outputAssumptions").value = model.outputsMeta.assumptions || "";

    const sched = model.time.discountSchedule || DEFAULT_DISCOUNT_SCHEDULE;
    $$("input[data-disc-period]").forEach(inp => {
      const idx = +inp.dataset.discPeriod;
      const scenario = inp.dataset.scenario;
      const row = sched[idx];
      if (!row) return;
      let v = "";
      if (scenario === "low") v = row.low;
      else if (scenario === "base") v = row.base;
      else if (scenario === "high") v = row.high;
      inp.value = v ?? "";
    });
  }

  function bindBasics() {
    setBasicsFieldsFromModel();
    initActions();

    const calcRiskBtn = $("#calcCombinedRisk");
    if (calcRiskBtn) {
      calcRiskBtn.addEventListener("click", e => {
        e.stopPropagation();
        const r =
          1 -
          (1 - num("#rTech")) *
            (1 - num("#rNonCoop")) *
            (1 - num("#rSocio")) *
            (1 - num("#rFin")) *
            (1 - num("#rMan"));
        if ($("#combinedRiskOut")) $("#combinedRiskOut").textContent = "Combined: " + (r * 100).toFixed(2) + "%";
        if ($("#riskBase")) $("#riskBase").value = r.toFixed(3);
        model.risk.base = r;
        calcAndRender();
        showToast("Combined risk updated from component risks.");
      });
    }

    const addCostBtn = $("#addCost");
    if (addCostBtn) {
      addCostBtn.addEventListener("click", e => {
        e.stopPropagation();
        const c = {
          id: uid(),
          label: "New cost",
          type: "annual",
          category: "Services",
          annual: 0,
          startYear: model.time.startYear,
          endYear: model.time.startYear,
          capital: 0,
          year: model.time.startYear,
          constrained: true,
          depMethod: "none",
          depLife: 5,
          depRate: 30
        };
        model.otherCosts.push(c);
        renderCosts();
        calcAndRender();
        showToast("New cost item added.");
      });
    }

    document.addEventListener("input", e => {
      const t = e.target;
      if (!t) return;

      if (t.dataset && t.dataset.discPeriod !== undefined) {
        const idx = +t.dataset.discPeriod;
        const scenario = t.dataset.scenario;
        if (!model.time.discountSchedule) {
          model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
        }
        const row = model.time.discountSchedule[idx];
        if (row && scenario) {
          const val = +t.value;
          if (scenario === "low") row.low = val;
          else if (scenario === "base") row.base = val;
          else if (scenario === "high") row.high = val;
          calcAndRenderDebounced();
        }
        return;
      }

      const id = t.id;
      if (!id) return;
      switch (id) {
        case "projectName":
          model.project.name = t.value;
          break;
        case "projectLead":
          model.project.lead = t.value;
          break;
        case "analystNames":
          model.project.analysts = t.value;
          break;
        case "projectTeam":
          model.project.team = t.value;
          break;
        case "projectSummary":
          model.project.summary = t.value;
          break;
        case "projectObjectives":
          model.project.objectives = t.value;
          break;
        case "projectActivities":
          model.project.activities = t.value;
          break;
        case "stakeholderGroups":
          model.project.stakeholders = t.value;
          break;
        case "lastUpdated":
          model.project.lastUpdated = t.value;
          break;
        case "projectGoal":
          model.project.goal = t.value;
          break;
        case "withProject":
          model.project.withProject = t.value;
          break;
        case "withoutProject":
          model.project.withoutProject = t.value;
          break;
        case "organisation":
          model.project.organisation = t.value;
          break;
        case "contactEmail":
          model.project.contactEmail = t.value;
          break;
        case "contactPhone":
          model.project.contactPhone = t.value;
          break;

        case "startYear":
          model.time.startYear = +t.value;
          break;
        case "projectStartYear":
          model.time.projectStartYear = +t.value;
          break;
        case "years":
          model.time.years = +t.value;
          break;
        case "discBase":
          model.time.discBase = +t.value;
          break;
        case "discLow":
          model.time.discLow = +t.value;
          break;
        case "discHigh":
          model.time.discHigh = +t.value;
          break;
        case "mirrFinance":
          model.time.mirrFinance = +t.value;
          break;
        case "mirrReinvest":
          model.time.mirrReinvest = +t.value;
          break;

        case "adoptBase":
          model.adoption.base = +t.value;
          break;
        case "adoptLow":
          model.adoption.low = +t.value;
          break;
        case "adoptHigh":
          model.adoption.high = +t.value;
          break;

        case "riskBase":
          model.risk.base = +t.value;
          break;
        case "riskLow":
          model.risk.low = +t.value;
          break;
        case "riskHigh":
          model.risk.high = +t.value;
          break;
        case "rTech":
          model.risk.tech = +t.value;
          break;
        case "rNonCoop":
          model.risk.nonCoop = +t.value;
          break;
        case "rSocio":
          model.risk.socio = +t.value;
          break;
        case "rFin":
          model.risk.fin = +t.value;
          break;
        case "rMan":
          model.risk.man = +t.value;
          break;

        case "simN":
          model.sim.n = +t.value;
          break;
        case "targetBCR":
          model.sim.targetBCR = +t.value;
          if ($("#simBcrTargetLabel")) $("#simBcrTargetLabel").textContent = t.value;
          break;
        case "bcrMode":
          model.sim.bcrMode = t.value;
          break;
        case "randSeed":
          model.sim.seed = t.value ? +t.value : null;
          break;

        case "simVarPct":
          model.sim.variationPct = +t.value || 20;
          break;
        case "simVaryOutputs":
          model.sim.varyOutputs = t.value === "true";
          break;
        case "simVaryTreatCosts":
          model.sim.varyTreatCosts = t.value === "true";
          break;
        case "simVaryInputCosts":
          model.sim.varyInputCosts = t.value === "true";
          break;

        case "systemType":
          model.outputsMeta.systemType = t.value;
          break;
        case "outputAssumptions":
          model.outputsMeta.assumptions = t.value;
          break;
      }
      calcAndRenderDebounced();
    });

    const saveProjectBtn = $("#saveProject");
    if (saveProjectBtn) {
      saveProjectBtn.addEventListener("click", e => {
        e.stopPropagation();
        const data = JSON.stringify(model, null, 2);
        downloadFile(
          "cba_" + (model.project.name || "project").replace(/\s+/g, "_") + ".json",
          data,
          "application/json"
        );
        showToast("Project JSON downloaded.");
      });
    }

    const loadProjectBtn = $("#loadProject");
    const loadFileInput = $("#loadFile");
    if (loadProjectBtn && loadFileInput) {
      loadProjectBtn.addEventListener("click", e => {
        e.stopPropagation();
        loadFileInput.click();
      });
      loadFileInput.addEventListener("change", async e => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        const text = await file.text();
        try {
          const obj = JSON.parse(text);
          Object.assign(model, obj);

          // Backward compatibility for treatments after JSON load
          if (Array.isArray(model.treatments)) {
            model.treatments.forEach(t => {
              if (typeof t.labourCost === "undefined") {
                t.labourCost = Number(t.annualCost || 0) || 0;
              }
              if (typeof t.materialsCost === "undefined") t.materialsCost = 0;
              if (typeof t.servicesCost === "undefined") t.servicesCost = 0;
              if (typeof t.adoption !== "number" || isNaN(t.adoption)) t.adoption = 1;
              delete t.annualCost;
            });
          }

          if (!model.time.discountSchedule) {
            model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
          }
          initTreatmentDeltas();
          renderAll();
          setBasicsFieldsFromModel();
          calcAndRender();
          showToast("Project JSON loaded and applied.");
        } catch (err) {
          alert("Invalid JSON file.");
          console.error(err);
        } finally {
          e.target.value = "";
        }
      });
    }

    const exportCsvBtn = $("#exportCsv");
    const exportCsvFootBtn = $("#exportCsvFoot");
    if (exportCsvBtn) exportCsvBtn.addEventListener("click", e => {
      e.stopPropagation();
      exportAllCsv();
    });
    if (exportCsvFootBtn) exportCsvFootBtn.addEventListener("click", e => {
      e.stopPropagation();
      exportAllCsv();
    });

    const exportPdfBtn = $("#exportPdf");
    const exportPdfFootBtn = $("#exportPdfFoot");
    if (exportPdfBtn)
      exportPdfBtn.addEventListener("click", e => {
        e.stopPropagation();
        exportPdf();
        showToast("Print dialog opened for PDF export.");
      });
    if (exportPdfFootBtn)
      exportPdfFootBtn.addEventListener("click", e => {
        e.stopPropagation();
        exportPdf();
        showToast("Print dialog opened for PDF export.");
      });

    const parseExcelBtn = $("#parseExcel");
    const importExcelBtn = $("#importExcel");
    if (parseExcelBtn)
      parseExcelBtn.addEventListener("click", e => {
        e.stopPropagation();
        handleParseExcel();
      });
    if (importExcelBtn)
      importExcelBtn.addEventListener("click", e => {
        e.stopPropagation();
        commitExcelToModel();
      });

    const downloadTemplateBtn = $("#downloadTemplate");
    const downloadSampleBtn = $("#downloadSample");
    if (downloadTemplateBtn)
      downloadTemplateBtn.addEventListener("click", e => {
        e.stopPropagation();
        downloadExcelTemplate();
      });
    if (downloadSampleBtn)
      downloadSampleBtn.addEventListener("click", e => {
        e.stopPropagation();
        downloadSampleDataset();
      });

    const startBtn = $("#startBtn");
    if (startBtn) {
      startBtn.addEventListener("click", e => {
        e.stopPropagation();
        switchTab("project");
        showToast("Welcome. Start with the Project tab.");
      });
    }

    const openCopilotBtns = $$("#openCopilot");
    if (openCopilotBtns.length) {
      openCopilotBtns.forEach(btn => {
        btn.addEventListener("click", e => {
          e.preventDefault();
          e.stopPropagation();
          handleOpenCopilotClick();
        });
      });
    }
  }

  // Initialise add buttons
  function initAddButtons() {
    const addOutputBtn = $("#addOutput");
    if (addOutputBtn) {
      addOutputBtn.addEventListener("click", e => {
        e.stopPropagation();
        const id = uid();
        model.outputs.push({
          id,
          name: "Custom output",
          unit: "unit",
          value: 0,
          source: "Input Directly"
        });
        model.treatments.forEach(t => {
          t.deltas[id] = 0;
        });
        renderOutputs();
        renderTreatments();
        renderDatabaseTags();
        calcAndRender();
        showToast("New output metric added.");
      });
    }

    const addTreatmentBtn = $("#addTreatment");
    if (addTreatmentBtn) {
      addTreatmentBtn.addEventListener("click", e => {
        e.stopPropagation();
        if (model.treatments.length >= 64) {
          alert("Maximum of 64 treatments reached.");
          return;
        }
        const t = {
          id: uid(),
          name: "New treatment",
          area: 0,
          adoption: 1,
          deltas: {},
          labourCost: 0,
          materialsCost: 0,
          servicesCost: 0,
          capitalCost: 0,
          constrained: true,
          source: "Input Directly",
          isControl: false,
          notes: ""
        };
        model.outputs.forEach(o => {
          t.deltas[o.id] = 0;
        });
        model.treatments.push(t);
        renderTreatments();
        renderDatabaseTags();
        calcAndRender();
        showToast("New treatment added.");
      });
    }

    const addBenefitBtn = $("#addBenefit");
    if (addBenefitBtn) {
      addBenefitBtn.addEventListener("click", e => {
        e.stopPropagation();
        model.benefits.push({
          id: uid(),
          label: "New benefit",
          category: "C4",
          theme: "Other",
          frequency: "Annual",
          startYear: model.time.startYear,
          endYear: model.time.startYear,
          year: model.time.startYear,
          unitValue: 0,
          quantity: 0,
          abatement: 0,
          annualAmount: 0,
          growthPct: 0,
          linkAdoption: true,
          linkRisk: true,
          p0: 0,
          p1: 0,
          consequence: 0,
          notes: ""
        });
        renderBenefits();
        calcAndRender();
        showToast("New benefit item added.");
      });
    }
  }

  // ---------- RENDERERS ----------
  function renderOutputs() {
    const root = $("#outputsList");
    if (!root) return;
    root.innerHTML = "";
    model.outputs.forEach(o => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>Output: ${esc(o.name)}</h4>
        <div class="row-6">
          <div class="field"><label>Name</label><input value="${esc(o.name)}" data-k="name" data-id="${o.id}" /></div>
          <div class="field"><label>Unit</label><input value="${esc(o.unit)}" data-k="unit" data-id="${o.id}" /></div>
          <div class="field"><label>Value ($/unit)</label><input type="number" step="0.01" value="${o.value}" data-k="value" data-id="${o.id}" /></div>
          <div class="field"><label>Source</label>
            <select data-k="source" data-id="${o.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"]
                .map(s => `<option ${s === o.source ? "selected" : ""}>${s}</option>`)
                .join("")}
            </select>
          </div>
          <div class="field"><label>&nbsp;</label><button class="btn small danger" data-del-output="${o.id}">Remove</button></div>
        </div>
        <div class="kv"><small class="muted">id:</small> <code>${o.id}</code></div>
      `;
      root.appendChild(el);
    });
    root.oninput = onOutputEdit;
    root.onclick = onOutputDelete;
  }

  function onOutputEdit(e) {
    const k = e.target.dataset.k;
    const id = e.target.dataset.id;
    if (!k || !id) return;
    const o = model.outputs.find(x => x.id === id);
    if (!o) return;
    if (k === "value") o[k] = +e.target.value;
    else o[k] = e.target.value;
    model.treatments.forEach(t => {
      if (!(id in t.deltas)) t.deltas[id] = 0;
    });
    renderTreatments();
    renderDatabaseTags();
    calcAndRenderDebounced();
  }

  function onOutputDelete(e) {
    const id = e.target.dataset.delOutput;
    if (!id) return;
    if (!confirm("Remove this output metric?")) return;
    model.outputs = model.outputs.filter(o => o.id !== id);
    model.treatments.forEach(t => {
      delete t.deltas[id];
    });
    renderOutputs();
    renderTreatments();
    renderDatabaseTags();
    calcAndRender();
    showToast("Output metric removed.");
  }

  function renderTreatments() {
    const root = $("#treatmentsList");
    if (!root) return;
    root.innerHTML = "";
    const list = [...model.treatments];
    list.forEach(t => {
      const materials = Number(t.materialsCost) || 0;
      const services = Number(t.servicesCost) || 0;
      const labour = Number(t.labourCost) || 0;
      const totalPerHa = materials + services + labour;
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>Treatment: ${esc(t.name)}</h4>
        <div class="row">
          <div class="field"><label>Name</label><input value="${esc(t.name)}" data-tk="name" data-id="${t.id}" /></div>
          <div class="field"><label>Area (ha)</label><input type="number" step="0.01" value="${t.area}" data-tk="area" data-id="${t.id}" /></div>
          <div class="field"><label>Source</label>
            <select data-tk="source" data-id="${t.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"]
                .map(s => `<option ${s === t.source ? "selected" : ""}>${s}</option>`)
                .join("")}
            </select>
          </div>
          <div class="field"><label>Control vs treatment?</label>
            <select data-tk="isControl" data-id="${t.id}">
              <option value="treatment" ${!t.isControl ? "selected" : ""}>Treatment</option>
              <option value="control" ${t.isControl ? "selected" : ""}>Control</option>
            </select>
          </div>
          <div class="field"><label>&nbsp;</label><button class="btn small danger" data-del-treatment="${t.id}">Remove</button></div>
        </div>
        <div class="row-6">
          <div class="field"><label>Materials cost ($/ha)</label><input type="number" step="0.01" value="${t.materialsCost || 0}" data-tk="materialsCost" data-id="${t.id}" /></div>
          <div class="field"><label>Services cost ($/ha)</label><input type="number" step="0.01" value="${t.servicesCost || 0}" data-tk="servicesCost" data-id="${t.id}" /></div>
          <div class="field"><label>Labour cost ($/ha)</label><input type="number" step="0.01" value="${t.labourCost || 0}" data-tk="labourCost" data-id="${t.id}" /></div>
          <div class="field"><label>Total cost ($/ha)</label><input type="number" step="0.01" value="${totalPerHa}" readonly data-total-cost="${t.id}" /></div>
          <div class="field"><label>Capital cost ($, year 0)</label><input type="number" step="0.01" value="${t.capitalCost}" data-tk="capitalCost" data-id="${t.id}" /></div>
          <div class="field"><label>Constrained?</label>
            <select data-tk="constrained" data-id="${t.id}">
              <option value="true" ${t.constrained ? "selected" : ""}>Yes</option>
              <option value="false" ${!t.constrained ? "selected" : ""}>No</option>
            </select>
          </div>
        </div>
        <div class="field">
          <label>Notes (for example implementation details or control definition)</label>
          <textarea data-tk="notes" data-id="${t.id}" rows="2">${esc(t.notes || "")}</textarea>
        </div>
        <h5>Output deltas (per ha)</h5>
        <div class="row">
          ${model.outputs
            .map(
              o => `
            <div class="field">
              <label>${esc(o.name)} (${esc(o.unit)})</label>
              <input type="number" step="0.0001" value="${t.deltas[o.id] ?? 0}" data-td="${o.id}" data-id="${t.id}" />
            </div>
          `
            )
            .join("")}
        </div>
        <div class="kv"><small class="muted">id:</small> <code>${t.id}</code></div>
      `;
      root.appendChild(el);
    });
    root.oninput = e => {
      const id = e.target.dataset.id;
      if (!id) return;
      const t = model.treatments.find(x => x.id === id);
      if (!t) return;
      const tk = e.target.dataset.tk;
      if (tk) {
        if (tk === "constrained") {
          t[tk] = e.target.value === "true";
        } else if (tk === "name" || tk === "source" || tk === "notes") {
          t[tk] = e.target.value;
        } else if (tk === "isControl") {
          const val = e.target.value === "control";
          model.treatments.forEach(tt => {
            tt.isControl = false;
          });
          if (val) t.isControl = true;
          renderTreatments();
          calcAndRenderDebounced();
          showToast(`Control treatment set to ${t.name}.`);
          return;
        } else {
          t[tk] = +e.target.value;
        }

        if (tk === "materialsCost" || tk === "servicesCost" || tk === "labourCost") {
          const container = e.target.closest(".item");
          if (container) {
            const mats = Number(
              container.querySelector(`input[data-tk="materialsCost"][data-id="${id}"]`)?.value || 0
            );
            const serv = Number(
              container.querySelector(`input[data-tk="servicesCost"][data-id="${id}"]`)?.value || 0
            );
            const lab = Number(
              container.querySelector(`input[data-tk="labourCost"][data-id="${id}"]`)?.value || 0
            );
            const totalField = container.querySelector(`input[data-total-cost="${id}"]`);
            if (totalField) totalField.value = mats + serv + lab;
          }
        }
      }
      const td = e.target.dataset.td;
      if (td) t.deltas[td] = +e.target.value;
      calcAndRenderDebounced();
    };
    root.addEventListener("click", e => {
      const id = e.target.dataset.delTreatment;
      if (!id) return;
      if (!confirm("Remove this treatment?")) return;
      model.treatments = model.treatments.filter(x => x.id !== id);
      renderTreatments();
      renderDatabaseTags();
      calcAndRender();
      showToast("Treatment removed.");
    });
  }

  function renderBenefits() {
    const root = $("#benefitsList");
    if (!root) return;
    root.innerHTML = "";
    const THEMES = [
      "Soil chemical",
      "Soil physical",
      "Soil biological",
      "Soil carbon",
      "Soil pH by depth",
      "Soil nutrients by depth",
      "Soil properties by treatment",
      "Cost savings",
      "Water retention",
      "Risk reduction",
      "Other"
    ];
    model.benefits.forEach(b => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>Benefit: ${esc(b.label || "Benefit")}</h4>
        <div class="row-6">
          <div class="field"><label>Label</label><input value="${esc(b.label || "")}" data-bk="label" data-id="${b.id}" /></div>
          <div class="field"><label>Category</label>
            <select data-bk="category" data-id="${b.id}">
              ${["C1","C2","C3","C4","C5","C6","C7","C8"]
                .map(c => `<option ${c === b.category ? "selected" : ""}>${c}</option>`)
                .join("")}
            </select>
          </div>
          <div class="field"><label>Benefit type</label>
            <select data-bk="theme" data-id="${b.id}">
              ${THEMES.map(th => `<option ${th === (b.theme || "") ? "selected" : ""}>${th}</option>`).join("")}
            </select>
          </div>
          <div class="field"><label>Frequency</label>
            <select data-bk="frequency" data-id="${b.id}">
              <option ${b.frequency === "Annual" ? "selected" : ""}>Annual</option>
              <option ${b.frequency === "Once" ? "selected" : ""}>Once</option>
            </select>
          </div>
          <div class="field"><label>Start year</label><input type="number" value="${b.startYear || model.time.startYear}" data-bk="startYear" data-id="${b.id}" /></div>
          <div class="field"><label>End year</label><input type="number" value="${b.endYear || model.time.startYear}" data-bk="endYear" data-id="${b.id}" /></div>
        </div>

        <div class="row-6">
          <div class="field"><label>Once year</label><input type="number" value="${b.year || model.time.startYear}" data-bk="year" data-id="${b.id}" /></div>
          <div class="field"><label>Unit value ($)</label><input type="number" step="0.01" value="${b.unitValue || 0}" data-bk="unitValue" data-id="${b.id}" /></div>
          <div class="field"><label>Quantity</label><input type="number" step="0.01" value="${b.quantity || 0}" data-bk="quantity" data-id="${b.id}" /></div>
          <div class="field"><label>Abatement</label><input type="number" step="0.01" value="${b.abatement || 0}" data-bk="abatement" data-id="${b.id}" /></div>
          <div class="field"><label>Annual amount ($)</label><input type="number" step="0.01" value="${b.annualAmount || 0}" data-bk="annualAmount" data-id="${b.id}" /></div>
          <div class="field"><label>Growth (% per year)</label><input type="number" step="0.01" value="${b.growthPct || 0}" data-bk="growthPct" data-id="${b.id}" /></div>
        </div>

        <div class="row-6">
          <div class="field"><label>Link adoption?</label>
            <select data-bk="linkAdoption" data-id="${b.id}">
              <option value="true" ${b.linkAdoption ? "selected" : ""}>Yes</option>
              <option value="false" ${!b.linkAdoption ? "selected" : ""}>No</option>
            </select>
          </div>
          <div class="field"><label>Link risk?</label>
            <select data-bk="linkRisk" data-id="${b.id}">
              <option value="true" ${b.linkRisk ? "selected" : ""}>Yes</option>
              <option value="false" ${!b.linkRisk ? "selected" : ""}>No</option>
            </select>
          </div>
          <div class="field"><label>P0 (baseline probability)</label><input type="number" step="0.001" value="${b.p0 || 0}" data-bk="p0" data-id="${b.id}" /></div>
          <div class="field"><label>P1 (with project probability)</label><input type="number" step="0.001" value="${b.p1 || 0}" data-bk="p1" data-id="${b.id}" /></div>
          <div class="field"><label>Consequence ($)</label><input type="number" step="0.01" value="${b.consequence || 0}" data-bk="consequence" data-id="${b.id}" /></div>
          <div class="field"><label>Notes</label><input value="${esc(b.notes || "")}" data-bk="notes" data-id="${b.id}" /></div>
          <div class="field"><label>&nbsp;</label><button class="btn small danger" data-del-benefit="${b.id}">Remove</button></div>
        </div>
      `;
      root.appendChild(el);
    });

    root.oninput = e => {
      const id = e.target.dataset.id;
      if (!id) return;
      const b = model.benefits.find(x => x.id === id);
      if (!b) return;
      const k = e.target.dataset.bk;
      if (!k) return;
      if (["label", "category", "frequency", "notes", "theme"].includes(k)) b[k] = e.target.value;
      else if (k === "linkAdoption" || k === "linkRisk") b[k] = e.target.value === "true";
      else b[k] = +e.target.value;
      calcAndRenderDebounced();
    };
    root.addEventListener("click", e => {
      const id = e.target.dataset.delBenefit;
      if (!id) return;
      if (!confirm("Remove this benefit item?")) return;
      model.benefits = model.benefits.filter(x => x.id !== id);
      renderBenefits();
      calcAndRender();
      showToast("Benefit item removed.");
    });
  }

  function renderCosts() {
    const root = $("#costsList");
    if (!root) return;
    root.innerHTML = "";
    model.otherCosts.forEach(c => {
      const isCapitalCategory = c.category === "Capital";
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>Cost item: ${esc(c.label)}</h4>
        <div class="row-6">
          <div class="field"><label>Label</label><input value="${esc(c.label)}" data-ck="label" data-id="${c.id}" /></div>
          <div class="field"><label>Type</label>
            <select data-ck="type" data-id="${c.id}">
              <option value="annual" ${c.type === "annual" ? "selected" : ""}>Annual</option>
              <option value="capital" ${c.type === "capital" ? "selected" : ""}>Capital</option>
            </select>
          </div>
          <div class="field"><label>Category</label>
            <select data-ck="category" data-id="${c.id}">
              <option ${c.category === "Capital" ? "selected" : ""}>Capital</option>
              <option ${c.category === "Labour" ? "selected" : ""}>Labour</option>
              <option ${c.category === "Materials" ? "selected" : ""}>Materials</option>
              <option ${c.category === "Services" ? "selected" : ""}>Services</option>
            </select>
          </div>
          <div class="field"><label>Annual ($ per year)</label><input type="number" step="0.01" value="${c.annual ?? 0}" data-ck="annual" data-id="${c.id}" /></div>
          <div class="field"><label>Start year</label><input type="number" value="${c.startYear ?? model.time.startYear}" data-ck="startYear" data-id="${c.id}" /></div>
          <div class="field"><label>End year</label><input type="number" value="${c.endYear ?? model.time.startYear}" data-ck="endYear" data-id="${c.id}" /></div>
        </div>
        <div class="row-6">
          <div class="field"><label>Capital ($)</label><input type="number" step="0.01" value="${c.capital ?? 0}" data-ck="capital" data-id="${c.id}" /></div>
          <div class="field"><label>Capital year</label><input type="number" value="${c.year ?? model.time.startYear}" data-ck="year" data-id="${c.id}" /></div>
          ${
            isCapitalCategory
              ? `
          <div class="field"><label>Depreciation method</label>
            <select data-ck="depMethod" data-id="${c.id}">
              <option value="none" ${c.depMethod === "none" ? "selected" : ""}>None</option>
              <option value="straight" ${c.depMethod === "straight" ? "selected" : ""}>Straight line</option>
              <option value="declining" ${c.depMethod === "declining" ? "selected" : ""}>Declining balance</option>
            </select>
          </div>
          <div class="field"><label>Life (years)</label><input type="number" step="1" min="1" value="${c.depLife || 5}" data-ck="depLife" data-id="${c.id}" /></div>
          <div class="field"><label>Declining rate (% per year)</label><input type="number" step="1" value="${c.depRate || 30}" data-ck="depRate" data-id="${c.id}" /></div>
          `
              : `
          <div class="field"><label>Depreciation method</label>
            <select data-ck="depMethod" data-id="${c.id}">
              <option value="none" selected>None</option>
            </select>
          </div>
          <div class="field"><label>Life (years)</label><input type="number" step="1" min="1" value="${c.depLife || 5}" data-ck="depLife" data-id="${c.id}" disabled /></div>
          <div class="field"><label>Declining rate (% per year)</label><input type="number" step="1" value="${c.depRate || 30}" data-ck="depRate" data-id="${c.id}" disabled /></div>
          `
          }
          <div class="field"><label>Constrained?</label>
            <select data-ck="constrained" data-id="${c.id}">
              <option value="true" ${c.constrained ? "selected" : ""}>Yes</option>
              <option value="false" ${!c.constrained ? "selected" : ""}>No</option>
            </select>
          </div>
          <div class="field"><label>&nbsp;</label><button class="btn small danger" data-del-cost="${c.id}">Remove</button></div>
        </div>
      `;
      root.appendChild(el);
    });
    root.oninput = e => {
      const id = e.target.dataset.id;
      const k = e.target.dataset.ck;
      if (!id || !k) return;
      const c = model.otherCosts.find(x => x.id === id);
      if (!c) return;
      if (["label", "type", "category", "depMethod"].includes(k)) {
        c[k] = e.target.value;
        if (k === "category" && c.category !== "Capital") {
          c.depMethod = "none";
        }
      } else if (k === "constrained") c[k] = e.target.value === "true";
      else c[k] = +e.target.value;
      calcAndRenderDebounced();
    };
    root.addEventListener("click", e => {
      const id = e.target.dataset.delCost;
      if (!id) return;
      if (!confirm("Remove this cost item?")) return;
      model.otherCosts = model.otherCosts.filter(x => x.id !== id);
      renderCosts();
      calcAndRender();
      showToast("Cost item removed.");
    });
  }

  function renderDatabaseTags() {
    const outRoot = $("#dbOutputs");
    if (outRoot) {
      outRoot.innerHTML = "";
      model.outputs.forEach(o => {
        const el = document.createElement("div");
        el.className = "item";
        el.innerHTML = `
          <div class="row-2">
            <div class="field"><label>${esc(o.name)} (${esc(o.unit)})</label></div>
            <div class="field">
              <label>Source</label>
              <select data-db-out="${o.id}">
                ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"]
                  .map(s => `<option ${s === o.source ? "selected" : ""}>${s}</option>`)
                  .join("")}
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
    }

    const tRoot = $("#dbTreatments");
    if (tRoot) {
      tRoot.innerHTML = "";
      model.treatments.forEach(t => {
        const el = document.createElement("div");
        el.className = "item";
        el.innerHTML = `
          <div class="row-2">
            <div class="field"><label>${esc(t.name)}</label></div>
            <div class="field">
              <label>Source</label>
              <select data-db-t="${t.id}">
                ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"]
                  .map(s => `<option ${s === t.source ? "selected" : ""}>${s}</option>`)
                  .join("")}
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
  }

  function computeSingleTreatmentMetrics(t, rate, years, adoptMul, risk) {
    let valuePerHa = 0;
    model.outputs.forEach(o => {
      valuePerHa += (Number(t.deltas[o.id]) || 0) * (Number(o.value) || 0);
    });
    const adopt = clamp(adoptMul, 0, 1);
    const area = Number(t.area) || 0;
    const annualBen = valuePerHa * area * (1 - clamp(risk, 0, 1)) * adopt;
    const annualCostPerHa =
      (Number(t.materialsCost) || 0) +
      (Number(t.servicesCost) || 0) +
      (Number(t.labourCost) || 0);
    const annualCost = annualCostPerHa * area;
    const cap = Number(t.capitalCost) || 0;
    const pvBen = annualBen * annuityFactor(years, rate);
    const pvCost = cap + annualCost * annuityFactor(years, rate);
    const bcr = pvCost > 0 ? pvBen / pvCost : NaN;
    const npv = pvBen - pvCost;
    const cf = new Array(years + 1).fill(0);
    cf[0] = -cap;
    for (let i = 1; i <= years; i++) cf[i] = annualBen - annualCost;
    const irrVal = irr(cf);
    const mirrVal = mirr(cf, model.time.mirrFinance, model.time.mirrReinvest);
    const roi = pvCost > 0 ? (npv / pvCost) * 100 : NaN;
    const gm = annualBen - annualCost;
    const gpm = annualBen > 0 ? (gm / annualBen) * 100 : NaN;
    const pb = payback(cf, rate);
    return { pvBen, pvCost, bcr, npv, irrVal, mirrVal, roi, gm, gpm, pb };
  }

  function renderTreatmentSummary(rate, adoptMul, risk) {
    const root = $("#treatmentSummary");
    if (!root) return;
    root.innerHTML = "";
    const list = [...model.treatments]
      .map(t => {
        const m = computeSingleTreatmentMetrics(
          t,
          rate,
          model.time.years,
          adoptMul,
          risk
        );
        return { t, m };
      })
      .sort((a, b) => {
        const bcrA = isFinite(a.m.bcr) ? a.m.bcr : -Infinity;
        const bcrB = isFinite(b.m.bcr) ? b.m.bcr : -Infinity;
        return bcrB - bcrA;
      });

    list.forEach((entry, idx) => {
      const t = entry.t;
      const m = entry.m;
      const adopt = clamp(adoptMul, 0, 1);
      const area = Number(t.area) || 0;
      const annualCostPerHa =
        (Number(t.materialsCost) || 0) +
        (Number(t.servicesCost) || 0) +
        (Number(t.labourCost) || 0);
      const annualCost = annualCostPerHa * area;
      const annualBen = m.gm + annualCost;
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <div class="row-6">
          <div class="field"><label>Rank</label><div class="metric"><div class="value">${idx + 1}</div></div></div>
          <div class="field"><label>Treatment</label><div class="metric"><div class="value">${esc(t.name)}${t.isControl ? " (Control)" : ""}</div></div></div>
          <div class="field"><label>Area</label><div class="metric"><div class="value">${fmt(area)} ha</div></div></div>
          <div class="field"><label>Adoption (multiplier)</label><div class="metric"><div class="value">${fmt(adopt)}</div></div></div>
          <div class="field"><label>Annual benefit</label><div class="metric"><div class="value">${money(annualBen)}</div></div></div>
          <div class="field"><label>Annual cost</label><div class="metric"><div class="value">${money(annualCost)}</div></div></div>
          <div class="field"><label>Present value of benefits</label><div class="metric"><div class="value">${money(m.pvBen)}</div></div></div>
          <div class="field"><label>Present value of costs</label><div class="metric"><div class="value">${money(m.pvCost)}</div></div></div>
          <div class="field"><label>BCR</label><div class="metric"><div class="value">${isFinite(m.bcr) ? fmt(m.bcr) : "n/a"}</div></div></div>
          <div class="field"><label>NPV</label><div class="metric"><div class="value">${money(m.npv)}</div></div></div>
          <div class="field"><label>IRR</label><div class="metric"><div class="value">${isFinite(m.irrVal) ? percent(m.irrVal) : "n/a"}</div></div></div>
          <div class="field"><label>Payback</label><div class="metric"><div class="value">${m.pb != null ? m.pb : "Not reached"}</div></div></div>
        </div>`;
      root.appendChild(el);
    });
  }

  function computeControlAndTreatmentGroupMetrics(rate, adoptMul, risk) {
    const years = model.time.years;
    const control = model.treatments.find(t => t.isControl);
    let controlMetrics = null;
    if (control) {
      controlMetrics = computeSingleTreatmentMetrics(control, rate, years, adoptMul, risk);
    }
    const combined = {
      area: 0,
      adoption: 1,
      deltas: {},
      materialsCost: 0,
      servicesCost: 0,
      labourCost: 0,
      capitalCost: 0,
      annualCost: 0
    };
    model.outputs.forEach(o => {
      combined.deltas[o.id] = 0;
    });
    model.treatments.forEach(t => {
      if (t.isControl) return;
      combined.capitalCost += Number(t.capitalCost) || 0;
      const area = Number(t.area) || 0;
      const adopt = clamp(adoptMul, 0, 1);
      combined.area += area;
      const annualCostPerHa =
        (Number(t.materialsCost) || 0) +
        (Number(t.servicesCost) || 0) +
        (Number(t.labourCost) || 0);
      combined.annualCost += annualCostPerHa * area;
      model.outputs.forEach(o => {
        combined.deltas[o.id] += (Number(t.deltas[o.id]) || 0) * (area * adopt);
      });
    });
    let treatMetrics = null;
    if (
      combined.capitalCost ||
      combined.annualCost ||
      Object.values(combined.deltas).some(v => v !== 0)
    ) {
      const temp = {
        id: "treatGroup",
        name: "Treatment group",
        area: combined.area || 1,
        adoption: 1,
        deltas: {},
        labourCost: 0,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: combined.capitalCost
      };
      model.outputs.forEach(o => {
        const perHaDelta = combined.area ? combined.deltas[o.id] / combined.area : 0;
        temp.deltas[o.id] = perHaDelta;
      });
      const annualCostPerHa = combined.area ? combined.annualCost / combined.area : 0;
      temp.labourCost = annualCostPerHa;
      treatMetrics = computeSingleTreatmentMetrics(
        temp,
        rate,
        years,
        1,
        risk
      );
    }
    return { controlMetrics, treatMetrics };
  }

  function renderDepreciationSummary() {
    const root = $("#depSummary");
    if (!root) return;
    root.innerHTML = "";
    const N = model.time.years;
    const baseYear = model.time.startYear;
    const totalPerYear = new Array(N + 1).fill(0);
    const rows = [];

    model.otherCosts.forEach(c => {
      if (c.category !== "Capital") return;
      const method = c.depMethod || "none";
      const cost = Number(c.capital) || 0;
      if (method === "none" || !cost) return;
      const life = Math.max(1, Number(c.depLife) || 5);
      const rate = Number(c.depRate) || 30;
      const startIndex = (Number(c.year) || baseYear) - baseYear;
      const sched = [];
      if (method === "straight") {
        const annual = cost / life;
        for (let i = 0; i < life; i++) {
          const idx = startIndex + i;
          if (idx >= 0 && idx <= N) {
            sched[idx] = (sched[idx] || 0) + annual;
            totalPerYear[idx] += annual;
          }
        }
      } else if (method === "declining") {
        let book = cost;
        for (let i = 0; i < life; i++) {
          const dep = (book * rate) / 100;
          const idx = startIndex + i;
          if (idx >= 0 && idx <= N) {
            sched[idx] = (sched[idx] || 0) + dep;
            totalPerYear[idx] += dep;
          }
          book -= dep;
          if (book <= 0) break;
        }
      }
      const firstDep = sched.find(v => v > 0) || 0;
      rows.push({
        label: c.label,
        method: method === "straight" ? "Straight line" : "Declining balance",
        life,
        rate: method === "declining" ? rate : "",
        firstDep
      });
    });

    if (!rows.length) {
      const p = document.createElement("p");
      p.className = "small muted";
      p.textContent =
        "No capital items with depreciation configured. Set the depreciation method on capital costs to see a schedule.";
      root.appendChild(p);
      return;
    }

    const tbl = document.createElement("table");
    tbl.className = "dep-table summary-table";
    tbl.innerHTML = `
      <thead>
        <tr>
          <th>Cost item</th>
          <th>Method</th>
          <th>Life (years)</th>
          <th>Rate (% per year)</th>
          <th>Approximate first year depreciation ($)</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            r => `
          <tr>
            <td>${esc(r.label)}</td>
            <td>${esc(r.method)}</td>
            <td>${r.life}</td>
            <td>${r.rate || ""}</td>
            <td>${money(r.firstDep)}</td>
          </tr>
        `
          )
          .join("")}
      </tbody>
    `;
    root.appendChild(tbl);
  }

  function renderAll() {
    renderOutputs();
    renderTreatments();
    renderBenefits();
    renderDatabaseTags();
    renderCosts();
  }

  // ---------- MAIN CALC / REPORT ----------
  function renderTimeProjections(benefitByYear, costByYear, rate) {
    const tblBody = $("#timeProjectionTable tbody");
    if (!tblBody) return;
    tblBody.innerHTML = "";
    const maxYears = model.time.years;
    const npvSeries = [];
    const usedHorizons = [];

    horizons.forEach(H => {
      const h = Math.min(H, maxYears);
      if (h <= 0) return;
      const b = benefitByYear.slice(0, h + 1);
      const c = costByYear.slice(0, h + 1);
      const pvB = presentValue(b, rate);
      const pvC = presentValue(c, rate);
      const npv = pvB - pvC;
      const bcr = pvC > 0 ? pvB / pvC : NaN;
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${h}</td>
        <td>${money(pvB)}</td>
        <td>${money(pvC)}</td>
        <td>${money(npv)}</td>
        <td>${isFinite(bcr) ? fmt(bcr) : "n/a"}</td>
      `;
      tblBody.appendChild(tr);
      npvSeries.push(npv);
      usedHorizons.push(h);
    });

    drawTimeSeries("timeNpvChart", usedHorizons, npvSeries);
  }

  let debTimer = null;
  function calcAndRenderDebounced() {
    clearTimeout(debTimer);
    debTimer = setTimeout(calcAndRender, 120);
  }

  function calcAndRender() {
    const rate = model.time.discBase;
    const adoptMul = model.adoption.base;
    const risk = model.risk.base;

    const all = computeAll(rate, adoptMul, risk, model.sim.bcrMode);

    setVal("#pvBenefits", money(all.pvBenefits));
    setVal("#pvCosts", money(all.pvCosts));
    const npvEl = $("#npv");
    if (npvEl) {
      npvEl.textContent = money(all.npv);
      npvEl.className = "value " + (all.npv >= 0 ? "positive" : "negative");
    }
    setVal("#bcr", isFinite(all.bcr) ? fmt(all.bcr) : "n/a");
    setVal("#irr", isFinite(all.irrVal) ? percent(all.irrVal) : "n/a");
    setVal("#mirr", isFinite(all.mirrVal) ? percent(all.mirrVal) : "n/a");
    setVal("#roi", isFinite(all.roi) ? percent(all.roi) : "n/a");
    setVal("#grossMargin", money(all.annualGM));
    setVal("#profitMargin", isFinite(all.profitMargin) ? percent(all.profitMargin) : "n/a");
    setVal("#payback", all.paybackYears != null ? all.paybackYears : "Not reached");

    renderTreatmentSummary(rate, adoptMul, risk);

    const { controlMetrics, treatMetrics } = computeControlAndTreatmentGroupMetrics(
      rate,
      adoptMul,
      risk
    );
    if (controlMetrics) {
      setVal("#pvBenefitsControl", money(controlMetrics.pvBen));
      setVal("#pvCostsControl", money(controlMetrics.pvCost));
      setVal("#npvControl", money(controlMetrics.npv));
      setVal("#bcrControl", isFinite(controlMetrics.bcr) ? fmt(controlMetrics.bcr) : "n/a");
      setVal("#irrControl", isFinite(controlMetrics.irrVal) ? percent(controlMetrics.irrVal) : "n/a");
      setVal("#roiControl", isFinite(controlMetrics.roi) ? percent(controlMetrics.roi) : "n/a");
      setVal("#paybackControl", controlMetrics.pb != null ? controlMetrics.pb : "Not reached");
      setVal("#gmControl", money(controlMetrics.gm));
    } else {
      [
        "#pvBenefitsControl",
        "#pvCostsControl",
        "#npvControl",
        "#bcrControl",
        "#irrControl",
        "#roiControl",
        "#paybackControl",
        "#gmControl"
      ].forEach(sel => setVal(sel, "No control selected"));
    }
    if (treatMetrics) {
      setVal("#pvBenefitsTreat", money(treatMetrics.pvBen));
      setVal("#pvCostsTreat", money(treatMetrics.pvCost));
      setVal("#npvTreat", money(treatMetrics.npv));
      setVal("#bcrTreat", isFinite(treatMetrics.bcr) ? fmt(treatMetrics.bcr) : "n/a");
      setVal("#irrTreat", isFinite(treatMetrics.irrVal) ? percent(treatMetrics.irrVal) : "n/a");
      setVal("#roiTreat", isFinite(treatMetrics.roi) ? percent(treatMetrics.roi) : "n/a");
      setVal("#paybackTreat", treatMetrics.pb != null ? treatMetrics.pb : "Not reached");
      setVal("#gmTreat", money(treatMetrics.gm));
    } else {
      [
        "#pvBenefitsTreat",
        "#pvCostsTreat",
        "#npvTreat",
        "#bcrTreat",
        "#irrTreat",
        "#roiTreat",
        "#paybackTreat",
        "#gmTreat"
      ].forEach(sel => setVal(sel, "n/a"));
    }

    renderTimeProjections(all.benefitByYear, all.costByYear, rate);
    renderDepreciationSummary();

    setVal("#simBcrTargetLabel", model.sim.targetBCR);
  }

  // ---------- MONTE CARLO ----------
  async function runSimulation() {
    const status = $("#simStatus");
    if (status) status.textContent = "Running ...";
    await new Promise(r => setTimeout(r));

    const N = model.sim.n;
    const seed = model.sim.seed;
    const rand = rng(seed ?? undefined);

    const discLow = model.time.discLow;
    const discBase = model.time.discBase;
    const discHigh = model.time.discHigh;
    const adoptLow = model.adoption.low;
    const adoptBase = model.adoption.base;
    const adoptHigh = model.adoption.high;
    const riskLow = model.risk.low;
    const riskBase = model.risk.base;
    const riskHigh = model.risk.high;

    const npvs = new Array(N);
    const bcrs = new Array(N);
    const details = [];
    const varPct = (model.sim.variationPct || 0) / 100;

    for (let i = 0; i < N; i++) {
      const r1 = rand();
      const r2 = rand();
      const r3 = rand();
      const disc = triangular(r1, discLow, discBase, discHigh);
      const adoptMul = clamp(triangular(r2, adoptLow, adoptBase, adoptHigh), 0, 1);
      const risk = clamp(triangular(r3, riskLow, riskBase, riskHigh), 0, 1);

      const shockOutputs = model.sim.varyOutputs ? 1 + (rand() * 2 * varPct - varPct) : 1;
      const shockTreatCosts = model.sim.varyTreatCosts ? 1 + (rand() * 2 * varPct - varPct) : 1;
      const shockInputCosts = model.sim.varyInputCosts ? 1 + (rand() * 2 * varPct - varPct) : 1;

      const origOutValues = model.outputs.map(o => o.value);
      const origTreatLab = model.treatments.map(t => t.labourCost || 0);
      const origTreatMat = model.treatments.map(t => t.materialsCost || 0);
      const origTreatServ = model.treatments.map(t => t.servicesCost || 0);
      const origTreatCap = model.treatments.map(t => t.capitalCost);
      const origOcAnn = model.otherCosts.map(c => c.annual);
      const origOcCap = model.otherCosts.map(c => c.capital);

      try {
        if (model.sim.varyOutputs) {
          model.outputs.forEach((o, idx) => {
            o.value = origOutValues[idx] * shockOutputs;
          });
        }
        if (model.sim.varyTreatCosts) {
          model.treatments.forEach((t, idx) => {
            t.labourCost = origTreatLab[idx] * shockTreatCosts;
            t.materialsCost = origTreatMat[idx] * shockTreatCosts;
            t.servicesCost = origTreatServ[idx] * shockTreatCosts;
          });
        }
        if (model.sim.varyInputCosts) {
          model.otherCosts.forEach((c, idx) => {
            c.annual = origOcAnn[idx] * shockInputCosts;
            c.capital = origOcCap[idx] * shockInputCosts;
          });
        }

        const { pvBenefits, pvCosts, bcr, npv } = computeAll(
          disc,
          adoptMul,
          risk,
          model.sim.bcrMode
        );
        npvs[i] = npv;
        bcrs[i] = bcr;
        details.push({
          run: i + 1,
          discount: disc,
          adoption: adoptMul,
          risk,
          pvBenefits,
          pvCosts,
          npv,
          bcr
        });
      } finally {
        model.outputs.forEach((o, idx) => {
          o.value = origOutValues[idx];
        });
        model.treatments.forEach((t, idx) => {
          t.labourCost = origTreatLab[idx];
          t.materialsCost = origTreatMat[idx];
          t.servicesCost = origTreatServ[idx];
          t.capitalCost = origTreatCap[idx];
        });
        model.otherCosts.forEach((c, idx) => {
          c.annual = origOcAnn[idx];
          c.capital = origOcCap[idx];
        });
      }
    }

    model.sim.results = { npv: npvs, bcr: bcrs };
    model.sim.details = details;
    if (status) status.textContent = "Done.";
    renderSimulationResults();
    drawHists();
    showToast("Monte Carlo simulation completed. Review results in the Simulation tab.");
  }

  function renderSimulationResults() {
    const { npv, bcr } = model.sim.results;
    if (!npv || !npv.length) return;
    const sortedNpv = [...npv].sort((a, b) => a - b);
    const validBcr = bcr.filter(x => isFinite(x));
    const sortedBcr = [...validBcr].sort((a, b) => a - b);
    const N = npv.length;
    const NB = sortedBcr.length;

    const stats = arr => ({
      min: arr[0],
      max: arr[arr.length - 1],
      mean: arr.reduce((a, c) => a + c, 0) / arr.length,
      median: arr.length
        ? (arr[Math.floor((arr.length - 1) / 2)] + arr[Math.ceil((arr.length - 1) / 2)]) / 2
        : NaN
    });

    const sN = stats(sortedNpv);
    const sB = sortedBcr.length ? stats(sortedBcr) : { min: NaN, max: NaN, mean: NaN, median: NaN };

    setVal("#simNpvMin", money(sN.min));
    setVal("#simNpvMax", money(sN.max));
    setVal("#simNpvMean", money(sN.mean));
    setVal("#simNpvMedian", money(sN.median));
    const pN = (npv.filter(x => x > 0).length / N) * 100;
    setVal("#simNpvProb", fmt(pN) + "%");

    setVal("#simBcrMin", isFinite(sB.min) ? fmt(sB.min) : "n/a");
    setVal("#simBcrMax", isFinite(sB.max) ? fmt(sB.max) : "n/a");
    setVal("#simBcrMean", isFinite(sB.mean) ? fmt(sB.mean) : "n/a");
    setVal("#simBcrMedian", isFinite(sB.median) ? fmt(sB.median) : "n/a");
    const pB1 = NB ? (validBcr.filter(x => x > 1).length / NB) * 100 : 0;
    setVal("#simBcrProb1", fmt(pB1) + "%");
    const tgt = model.sim.targetBCR;
    const pBt = NB ? (validBcr.filter(x => x > tgt).length / NB) * 100 : 0;
    setVal("#simBcrProbTarget", fmt(pBt) + "%");
  }

  function drawHist(canvasId, data, bins = 24, labelFmt = v => v.toFixed(0), titleText = "") {
    const canvas = document.getElementById(canvasId);
    if (!canvas || !data || !data.length) return;
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    const min = Math.min(...data);
    const max = Math.max(...data);
    const padL = 54;
    const padR = 14;
    const padT = 24;
    const padB = 40;
    const W = canvas.width - padL - padR;
    const H = canvas.height - padT - padB;

    const counts = new Array(bins).fill(0);
    const span = max - min || 1e-9;
    data.forEach(v => {
      let idx = Math.floor(((v - min) / span) * bins);
      if (idx < 0) idx = 0;
      if (idx >= bins) idx = bins - 1;
      counts[idx]++;
    });
    const maxC = Math.max(...counts) || 1;

    ctx.strokeStyle = "#3c6a52";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(padL, padT);
    ctx.lineTo(padL, padT + H);
    ctx.lineTo(padL + W, padT + H);
    ctx.stroke();

    for (let i = 0; i < bins; i++) {
      const x = padL + (i * W) / bins + 1;
      const h = (counts[i] / maxC) * (H - 2);
      const y = padT + H - h;
      ctx.fillStyle = "rgba(116, 209, 140, 0.65)";
      ctx.fillRect(x, y, W / bins - 2, h);
    }

    ctx.fillStyle = "#c9efd6";
    ctx.font = "12px system-ui";
    ctx.textAlign = "center";
    const lbls = [min, (min + max) / 2, max];
    [0, 0.5, 1].forEach((p, i) => {
      const x = padL + p * W;
      ctx.fillText(labelFmt(lbls[i]), x, padT + H + 22);
    });

    ctx.textAlign = "center";
    if (titleText) {
      ctx.fillText(titleText, padL + W / 2, padT - 8);
    }
    ctx.save();
    ctx.translate(12, padT + H / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.fillText("Frequency", 0, 0);
    ctx.restore();
  }

  function drawHists() {
    const { npv, bcr } = model.sim.results;
    if (npv && npv.length) drawHist("histNpv", npv, 24, v => v.toFixed(0), "Distribution of NPV");
    if (bcr && bcr.length)
      drawHist(
        "histBcr",
        bcr.filter(x => isFinite(x)),
        24,
        v => v.toFixed(2),
        "Distribution of BCR"
      );
  }

  function drawTimeSeries(canvasId, xs, ys) {
    const canvas = document.getElementById(canvasId);
    if (!canvas || !xs.length || !ys.length) return;
    const ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    const padL = 60;
    const padR = 16;
    const padT = 24;
    const padB = 40;
    const W = canvas.width - padL - padR;
    const H = canvas.height - padT - padB;

    const minX = Math.min(...xs);
    const maxX = Math.max(...xs);
    const minY = Math.min(...ys);
    const maxY = Math.max(...ys);
    const yMin = Math.min(minY, 0);
    const yMax = Math.max(maxY, 0) || 1;

    const xScale = v => padL + ((v - minX) / (maxX - minX || 1)) * W;
    const yScale = v => padT + H - ((v - yMin) / (yMax - yMin || 1)) * H;

    ctx.strokeStyle = "#3c6a52";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(padL, padT);
    ctx.lineTo(padL, padT + H);
    ctx.lineTo(padL + W, padT + H);
    ctx.stroke();

    const zeroY = yScale(0);
    ctx.strokeStyle = "rgba(200,220,210,0.6)";
    ctx.setLineDash([4, 4]);
    ctx.beginPath();
    ctx.moveTo(padL, zeroY);
    ctx.lineTo(padL + W, zeroY);
    ctx.stroke();
    ctx.setLineDash([]);

    ctx.strokeStyle = "rgba(116,209,140,0.9)";
    ctx.lineWidth = 2;
    ctx.beginPath();
    xs.forEach((xv, idx) => {
      const x = xScale(xv);
      const y = yScale(ys[idx]);
      if (idx === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });
    ctx.stroke();

    ctx.fillStyle = "#c9efd6";
    ctx.font = "12px system-ui";
    ctx.textAlign = "center";
    xs.forEach(xv => {
      const x = xScale(xv);
      ctx.fillText(String(xv), x, padT + H + 20);
    });
    ctx.textAlign = "right";
    ctx.fillText(money(yMax), padL - 6, yScale(yMax) + 4);
    ctx.fillText(money(0), padL - 6, zeroY + 4);
    ctx.fillText(money(yMin), padL - 6, yScale(yMin) + 4);

    ctx.textAlign = "center";
    ctx.fillText("Years of analysis", padL + W / 2, padT + H + 36);
    ctx.save();
    ctx.translate(16, padT + H / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.fillText("Net present value", 0, 0);
    ctx.restore();
    ctx.fillText("NPV by horizon", padL + W / 2, padT - 8);
  }

  // ---------- EXPORTS ----------
  function toCsv(rows) {
    return rows
      .map(r =>
        r
          .map(v => {
            const s = (v ?? "").toString();
            if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
            return s;
          })
          .join(",")
      )
      .join("\n");
  }

  function downloadFile(filename, text, mime = "text/csv") {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([text], { type: mime }));
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function downloadBinaryFile(filename, data, mime) {
    const a = document.createElement("a");
    const blob = new Blob([data], { type: mime });
    a.href = URL.createObjectURL(blob);
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
        lead: model.project.lead,
        analysts: model.project.analysts,
        team: model.project.team,
        organisation: model.project.organisation,
        contact: model.project.contactEmail,
        phone: model.project.contactPhone,
        updated: model.project.lastUpdated
      },
      params: {
        startYear: model.time.startYear,
        projectStartYear: model.time.projectStartYear,
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
      ["Project lead", s.meta.lead],
      ["Analysts", s.meta.analysts],
      ["Project team", s.meta.team],
      ["Organisation", s.meta.organisation],
      ["Contact email", s.meta.contact],
      ["Contact phone", s.meta.phone],
      ["Last updated", s.meta.updated],
      [],
      ["Analysis start year", s.params.startYear],
      ["Project start year", s.params.projectStartYear],
      ["Years", s.params.years],
      ["Discount rate (base)", s.params.discountBase],
      ["Discount rate (low)", s.params.discountLow],
      ["Discount rate (high)", s.params.discountHigh],
      ["MIRR finance %", s.params.mirrFinance],
      ["MIRR reinvest %", s.params.mirrReinvest],
      ["Adoption multiplier", s.params.adoptionBase],
      ["Risk (overall)", s.params.riskBase],
      ["BCR mode", s.params.bcrMode],
      [],
      ["Present value of benefits", s.results.pvBenefits],
      ["Present value of costs", s.results.pvCosts],
      ["NPV", s.results.npv],
      ["BCR", s.results.bcr],
      ["IRR %", s.results.irrVal],
      ["MIRR %", s.results.mirrVal],
      ["ROI %", s.results.roi],
      ["Gross margin (annual)", s.results.annualGM],
      ["Gross profit margin %", s.results.profitMargin],
      ["Payback (years)", s.results.paybackYears ?? "Not reached"]
    ];
    downloadFile("cba_summary_" + slug(s.meta.name) + ".csv", toCsv(summaryRows));

    const treatHeader = [
      "Treatment",
      "Area (ha)",
      "Adoption multiplier",
      "Annual benefit",
      "Annual cost",
      "PV benefit",
      "PV cost",
      "BCR",
      "NPV"
    ];
    const treatRows = [treatHeader];
    const rate = model.time.discBase;
    const adoptMul = model.adoption.base;
    const risk = model.risk.base;
    model.treatments.forEach(t => {
      const m = computeSingleTreatmentMetrics(
        t,
        rate,
        model.time.years,
        adoptMul,
        risk
      );
      const adopt = clamp(adoptMul, 0, 1);
      const area = Number(t.area) || 0;
      const annualCostPerHa =
        (Number(t.materialsCost) || 0) +
        (Number(t.servicesCost) || 0) +
        (Number(t.labourCost) || 0);
      const annualCost = annualCostPerHa * area;
      const annualBen = m.gm + annualCost;
      treatRows.push([
        t.name + (t.isControl ? " (Control)" : ""),
        area,
        adopt,
        annualBen,
        annualCost,
        m.pvBen,
        m.pvCost,
        m.bcr,
        m.npv
      ]);
    });
    downloadFile("cba_treatments_" + slug(s.meta.name) + ".csv", toCsv(treatRows));

    const benRows = [
      [
        "Label",
        "Category",
        "BenefitType",
        "Frequency",
        "StartYear",
        "EndYear",
        "Year",
        "UnitValue",
        "Quantity",
        "Abatement",
        "AnnualAmount",
        "GrowthPct",
        "LinkAdoption",
        "LinkRisk",
        "P0",
        "P1",
        "Consequence",
        "Notes"
      ]
    ];
    model.benefits.forEach(b =>
      benRows.push([
        b.label,
        b.category,
        b.theme || "",
        b.frequency,
        b.startYear,
        b.endYear,
        b.year,
        b.unitValue,
        b.quantity,
        b.abatement,
        b.annualAmount,
        b.growthPct,
        b.linkAdoption,
        b.linkRisk,
        b.p0,
        b.p1,
        b.consequence,
        b.notes
      ])
    );
    downloadFile("cba_benefits_" + slug(s.meta.name) + ".csv", toCsv(benRows));

    const outRows = [["Output", "Unit", "$ per unit", "Source", "Id"]];
    model.outputs.forEach(o => outRows.push([o.name, o.unit, o.value, o.source, o.id]));
    downloadFile("cba_outputs_" + slug(s.meta.name) + ".csv", toCsv(outRows));

    showToast("CSV summary and detail exports downloaded.");
  }

  function exportPdf() {
    window.print();
  }

  // ---------- EXCEL IMPORT / EXPORT ----------
  function ensureXlsxAvailable() {
    if (typeof XLSX === "undefined") {
      showToast("Excel library is not loaded. Include XLSX in the HTML to use Excel features.");
      return false;
    }
    return true;
  }

  function createExcelInput() {
    let input = document.getElementById("excelFile");
    if (!input) {
      input = document.createElement("input");
      input.type = "file";
      input.accept = ".xlsx,.xls,.csv";
      input.id = "excelFile";
      input.style.display = "none";
      document.body.appendChild(input);
    }
    return input;
  }

  function handleParseExcel() {
    if (!ensureXlsxAvailable()) return;
    const input = createExcelInput();
    input.onchange = e => {
      const file = e.target.files && e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = evt => {
        try {
          const data = evt.target.result;
          const wb = XLSX.read(data, { type: "array" });
          const res = {};
          const getSheet = name =>
            wb.Sheets[name] ? XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: "" }) : [];

          res.meta = getSheet("Meta");
          res.outputs = getSheet("Outputs");
          res.treatments = getSheet("Treatments");
          res.benefits = getSheet("Benefits");
          res.costs = getSheet("Costs");

          parsedExcel = res;
          const nOut = res.outputs.length || 0;
          const nTreat = res.treatments.length || 0;
          const nBen = res.benefits.length || 0;
          const nCost = res.costs.length || 0;
          showToast(
            `Excel parsed successfully: ${nOut} outputs, ${nTreat} treatments, ${nBen} benefits, ${nCost} cost items ready to import.`
          );
        } catch (err) {
          console.error(err);
          showToast("Could not parse Excel file. Check that the template structure is correct.");
        } finally {
          e.target.value = "";
        }
      };
      reader.readAsArrayBuffer(file);
    };
    input.click();
  }

  function commitExcelToModel() {
    if (!parsedExcel) {
      showToast("No parsed Excel data found. Click Parse Excel first.");
      return;
    }
    const data = parsedExcel;
    const now = new Date().getFullYear();

    // Meta (optional)
    if (data.meta && data.meta.length) {
      const mrow = data.meta[0];
      if (mrow.ProjectName) model.project.name = mrow.ProjectName;
      if (mrow.ProjectLead) model.project.lead = mrow.ProjectLead;
      if (mrow.Analysts) model.project.analysts = mrow.Analysts;
      if (mrow.Team) model.project.team = mrow.Team;
      if (mrow.Organisation) model.project.organisation = mrow.Organisation;
      if (mrow.ContactEmail) model.project.contactEmail = mrow.ContactEmail;
      if (mrow.ContactPhone) model.project.contactPhone = mrow.ContactPhone;
      if (mrow.Summary) model.project.summary = mrow.Summary;
      model.project.lastUpdated = new Date().toISOString().slice(0, 10);
    }

    // Outputs
    if (data.outputs && data.outputs.length) {
      model.outputs = data.outputs.map(row => ({
        id: uid(),
        name: row.Name || row.Output || "Output",
        unit: row.Unit || "unit",
        value:
          Number(row.ValuePerUnit ?? row.Value ?? row["$ per unit"] ?? row.Price ?? 0) || 0,
        source: row.Source || "Input Directly"
      }));
    }

    // Treatments
    const outputNameToId = {};
    model.outputs.forEach(o => {
      outputNameToId[o.name.toLowerCase()] = o.id;
    });

    if (data.treatments && data.treatments.length) {
      model.treatments = data.treatments.map(row => {
        const isControlString = String(row.IsControl ?? row.Control ?? "").toLowerCase();
        const t = {
          id: uid(),
          name: row.Name || "Treatment",
          area: Number(row.AreaHa ?? row["Area (ha)"] ?? 0) || 0,
          adoption: 1,
          deltas: {},
          labourCost:
            Number(
              row.LabourCostPerHa ??
                row.AnnualCostPerHa ??
                row["AnnualCostPerHa"] ??
                0
            ) || 0,
          materialsCost: Number(row.MaterialsCostPerHa ?? 0) || 0,
          servicesCost: Number(row.ServicesCostPerHa ?? 0) || 0,
          capitalCost: Number(row.CapitalCost ?? row.Capital ?? 0) || 0,
          constrained: String(row.Constrained ?? "Yes").toLowerCase().startsWith("y"),
          source: row.Source || "Input Directly",
          isControl:
            isControlString.startsWith("y") ||
            isControlString === "control" ||
            row.IsControl === 1,
          notes: row.Notes || ""
        };

        Object.keys(row).forEach(key => {
          const m = key.match(/^Delta[_:\s]*(.+)$/i);
          if (m) {
            const outName = m[1].trim().toLowerCase();
            const id = outputNameToId[outName];
            if (id) {
              t.deltas[id] = Number(row[key] || 0) || 0;
            }
          }
        });

        model.outputs.forEach(o => {
          if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
        });
        return t;
      });
    } else {
      initTreatmentDeltas();
    }

    // Benefits
    if (data.benefits && data.benefits.length) {
      model.benefits = data.benefits.map(row => ({
        id: uid(),
        label: row.Label || "Benefit",
        category: row.Category || "C4",
        theme: row.BenefitType || row.Theme || "Other",
        frequency: row.Frequency || "Annual",
        startYear: Number(row.StartYear || now) || now,
        endYear: Number(row.EndYear || now) || now,
        year: Number(row.Year || row.OnceYear || now) || now,
        unitValue: Number(row.UnitValue || 0) || 0,
        quantity: Number(row.Quantity || 0) || 0,
        abatement: Number(row.Abatement || 0) || 0,
        annualAmount: Number(row.AnnualAmount || 0) || 0,
        growthPct: Number(row.GrowthPct || 0) || 0,
        linkAdoption:
          String(row.LinkAdoption ?? "true").toLowerCase().startsWith("t") ||
          String(row.LinkAdoption ?? "true").toLowerCase().startsWith("y"),
        linkRisk:
          String(row.LinkRisk ?? "true").toLowerCase().startsWith("t") ||
          String(row.LinkRisk ?? "true").toLowerCase().startsWith("y"),
        p0: Number(row.P0 || 0) || 0,
        p1: Number(row.P1 || 0) || 0,
        consequence: Number(row.Consequence || 0) || 0,
        notes: row.Notes || ""
      }));
    }

    // Costs
    if (data.costs && data.costs.length) {
      model.otherCosts = data.costs.map(row => ({
        id: uid(),
        label: row.Label || "Cost item",
        type: (row.Type || "annual").toString().toLowerCase() === "capital" ? "capital" : "annual",
        category: row.Category || "Services",
        annual: Number(row.Annual || 0) || 0,
        startYear: Number(row.StartYear || now) || now,
        endYear: Number(row.EndYear || now) || now,
        capital: Number(row.Capital || 0) || 0,
        year: Number(row.Year || now) || now,
        constrained:
          String(row.Constrained ?? "true").toLowerCase().startsWith("t") ||
          String(row.Constrained ?? "true").toLowerCase().startsWith("y"),
        depMethod: (row.DepMethod || "none").toString().toLowerCase(),
        depLife: Number(row.DepLife || 5) || 5,
        depRate: Number(row.DepRate || 30) || 30
      }));
    }

    initTreatmentDeltas();
    renderAll();
    setBasicsFieldsFromModel();
    calcAndRender();
    showToast("Excel data imported into the model. Results updated.");
    parsedExcel = null;
  }

  function buildTemplateWorkbook(includeSampleRows) {
    const wb = XLSX.utils.book_new();

    const metaRows = [
      [
        "ProjectName",
        "ProjectLead",
        "Analysts",
        "Team",
        "Organisation",
        "ContactEmail",
        "ContactPhone",
        "Summary"
      ]
    ];
    if (includeSampleRows) {
      metaRows.push([
        model.project.name,
        model.project.lead,
        model.project.analysts,
        model.project.team,
        model.project.organisation,
        model.project.contactEmail,
        model.project.contactPhone,
        model.project.summary
      ]);
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(metaRows), "Meta");

    const outHeader = ["Name", "Unit", "ValuePerUnit", "Source"];
    const outAoA = [outHeader];
    if (includeSampleRows) {
      model.outputs.forEach(o => {
        outAoA.push([o.name, o.unit, o.value, o.source]);
      });
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(outAoA), "Outputs");

    const baseTreatColumns = [
      "Name",
      "AreaHa",
      "MaterialsCostPerHa",
      "ServicesCostPerHa",
      "LabourCostPerHa",
      "CapitalCost",
      "IsControl",
      "Constrained",
      "Source",
      "Notes"
    ];
    const deltaCols = model.outputs.map(o => "Delta_" + o.name);
    const treatHeader = baseTreatColumns.concat(deltaCols);
    const treatAoA = [treatHeader];
    if (includeSampleRows) {
      model.treatments.forEach(t => {
        const baseRow = [
          t.name,
          t.area,
          t.materialsCost || 0,
          t.servicesCost || 0,
          t.labourCost || 0,
          t.capitalCost || 0,
          t.isControl ? "Control" : "Treatment",
          t.constrained ? "Yes" : "No",
          t.source,
          t.notes || ""
        ];
        const deltasRow = model.outputs.map(o => t.deltas[o.id] ?? 0);
        treatAoA.push(baseRow.concat(deltasRow));
      });
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(treatAoA), "Treatments");

    const benHeader = [
      "Label",
      "Category",
      "BenefitType",
      "Frequency",
      "StartYear",
      "EndYear",
      "Year",
      "UnitValue",
      "Quantity",
      "Abatement",
      "AnnualAmount",
      "GrowthPct",
      "LinkAdoption",
      "LinkRisk",
      "P0",
      "P1",
      "Consequence",
      "Notes"
    ];
    const benAoA = [benHeader];
    if (includeSampleRows) {
      model.benefits.forEach(b => {
        benAoA.push([
          b.label,
          b.category,
          b.theme || "",
          b.frequency,
          b.startYear,
          b.endYear,
          b.year,
          b.unitValue,
          b.quantity,
          b.abatement,
          b.annualAmount,
          b.growthPct,
          b.linkAdoption ? "Yes" : "No",
          b.linkRisk ? "Yes" : "No",
          b.p0,
          b.p1,
          b.consequence,
          b.notes || ""
        ]);
      });
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(benAoA), "Benefits");

    const costHeader = [
      "Label",
      "Type",
      "Category",
      "Annual",
      "StartYear",
      "EndYear",
      "Capital",
      "Year",
      "DepMethod",
      "DepLife",
      "DepRate",
      "Constrained"
    ];
    const costAoA = [costHeader];
    if (includeSampleRows) {
      model.otherCosts.forEach(c => {
        costAoA.push([
          c.label,
          c.type,
          c.category,
          c.annual,
          c.startYear,
          c.endYear,
          c.capital,
          c.year,
          c.depMethod,
          c.depLife,
          c.depRate,
          c.constrained ? "Yes" : "No"
        ]);
      });
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(costAoA), "Costs");

    return wb;
  }

  function downloadExcelTemplate() {
    if (!ensureXlsxAvailable()) return;
    const wb = buildTemplateWorkbook(false);
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadBinaryFile("farming_cba_template.xlsx", wbout, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Excel template downloaded. Use this structure for imports.");
  }

  function downloadSampleDataset() {
    if (!ensureXlsxAvailable()) return;
    const wb = buildTemplateWorkbook(true);
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadBinaryFile(
      "farming_cba_sample_australia.xlsx",
      wbout,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    showToast("Sample Excel dataset downloaded using a realistic Australian nitrogen trial scenario.");
  }

  // ---------- COPILOT HELPER ----------
  async function copyToClipboard(text) {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(text);
        return true;
      }
    } catch (_) {}
    try {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.style.position = "fixed";
      ta.style.opacity = "0";
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      document.execCommand("copy");
      ta.remove();
      return true;
    } catch (_) {
      return false;
    }
  }

  function buildScenarioForCopilot() {
    const summary = buildSummaryForCsv();
    const base = summary.results;

    const rate = model.time.discBase;
    const adoptMul = model.adoption.base;
    const risk = model.risk.base;

    const treatmentSummaries = model.treatments.slice(0, 10).map(t => {
      const m = computeSingleTreatmentMetrics(
        t,
        rate,
        model.time.years,
        adoptMul,
        risk
      );
      return {
        name: t.name,
        areaHa: t.area,
        adoptionMultiplier: model.adoption.base,
        isControl: !!t.isControl,
        pvBenefits: m.pvBen,
        pvCosts: m.pvCost,
        npv: m.npv,
        bcr: m.bcr,
        irr: m.irrVal,
        paybackYears: m.pb
      };
    });

    const npvArr = model.sim.results.npv || [];
    const bcrArr = model.sim.results.bcr || [];
    let simSummary = null;
    if (npvArr.length) {
      const sortedNpv = [...npvArr].sort((a, b) => a - b);
      const N = sortedNpv.length;
      const meanNpv = sortedNpv.reduce((a, c) => a + c, 0) / N;
      const medianNpv =
        (sortedNpv[Math.floor((N - 1) / 2)] + sortedNpv[Math.ceil((N - 1) / 2)]) / 2;
      const npvMin = sortedNpv[0];
      const npvMax = sortedNpv[sortedNpv.length - 1];
      const probPosNpv = npvArr.filter(v => v > 0).length / N;

      const validBcr = bcrArr.filter(x => isFinite(x));
      let bcrMin = null;
      let bcrMax = null;
      let bcrMean = null;
      let bcrMedian = null;
      let probBcrAbove1 = null;
      let probBcrAboveTarget = null;
      if (validBcr.length) {
        const sortedBcr = [...validBcr].sort((a, b) => a - b);
        const NB = sortedBcr.length;
        bcrMin = sortedBcr[0];
        bcrMax = sortedBcr[NB - 1];
        bcrMean = sortedBcr.reduce((a, c) => a + c, 0) / NB;
        bcrMedian =
          (sortedBcr[Math.floor((NB - 1) / 2)] + sortedBcr[Math.ceil((NB - 1) / 2)]) / 2;
        probBcrAbove1 = validBcr.filter(x => x > 1).length / NB;
        probBcrAboveTarget = validBcr.filter(x => x > model.sim.targetBCR).length / NB;
      }

      simSummary = {
        runs: N,
        discountLow: model.time.discLow,
        discountBase: model.time.discBase,
        discountHigh: model.time.discHigh,
        adoptionLow: model.adoption.low,
        adoptionBase: model.adoption.base,
        adoptionHigh: model.adoption.high,
        riskLow: model.risk.low,
        riskBase: model.risk.base,
        riskHigh: model.risk.high,
        npvMin,
        npvMax,
        npvMean: meanNpv,
        npvMedian: medianNpv,
        probNpvPositive: probPosNpv,
        bcrMin,
        bcrMax,
        bcrMean,
        bcrMedian,
        probBcrAbove1,
        probBcrAboveTarget: probBcrAboveTarget,
        targetBCR: model.sim.targetBCR
      };
    }

    return {
      project: summary.meta,
      parameters: summary.params,
      baseCase: {
        pvBenefits: base.pvBenefits,
        pvCosts: base.pvCosts,
        npv: base.npv,
        bcr: base.bcr,
        irr: base.irrVal,
        mirr: base.mirrVal,
        roi: base.roi,
        annualGrossMargin: base.annualGM,
        profitMargin: base.profitMargin,
        paybackYears: base.paybackYears
      },
      treatmentSummaries,
      simulationSummary: simSummary
    };
  }

  function handleOpenCopilotClick() {
    calcAndRender();
    const scenario = buildScenarioForCopilot();
    let scenarioJson = JSON.stringify(scenario, null, 2);
    const maxJsonChars = 5000;
    if (scenarioJson.length > maxJsonChars) {
      scenarioJson = scenarioJson.slice(0, maxJsonChars) + "\n... (scenario truncated to fit character limits)";
    }

    const promptText =
`You are an agricultural economics assistant helping to interpret a cost benefit analysis for a farming project.

Using the JSON scenario below, write a clear and accessible policy briefing for decision makers in agriculture and natural resource management in Australia. Structure the briefing in well separated sections with short headings, such as: Project context, Methods and assumptions, Base case economic results, Risk and uncertainty, Adoption and implementation issues, and Policy conclusions.

Explain in plain language what the project does, why it matters, and how the cost benefit analysis was carried out. Describe the main economic indicators including present value of benefits and costs, net present value, benefit cost ratio, internal rate of return, modified internal rate of return, return on investment, annual gross margin, profit margin, and payback period. Interpret the results for each indicator in words that a nontechnical reader can understand.

If simulation results are available, summarise what they show about the probability that net present value is positive, the probability that the benefit cost ratio is above 1, and the probability that it is above the target threshold. Explain what these probabilities mean in practical terms.

Discuss how adoption, risk, and implementation constraints could influence the results, and outline any important distributional or implementation considerations that policy makers should be aware of.

Conclude with a clear recommendation on whether the project appears economically attractive, under what conditions it is most attractive, and what further evidence or monitoring might be useful.

Write only in paragraphs with no bullet points or numbered lists. Use professional but lay friendly language. Aim for the level of detail that would correspond to roughly three to five pages of text. You may include self contained tables if they help to present the main numbers clearly.

SCENARIO JSON:
\`\`\`json
${scenarioJson}
\`\`\`
`;

    copyToClipboard(promptText).then(ok => {
      if (ok) {
        showToast("Copilot briefing prompt copied. A new Copilot tab will open.");
      } else {
        showToast("Could not copy automatically. Use the Copilot tab to copy the text manually.");
      }
      window.open("https://copilot.microsoft.com/", "_blank");
    });
  }

  // ---------- INIT ----------
  document.addEventListener("DOMContentLoaded", () => {
    initTabs();
    bindBasics();
    initAddButtons();
    renderAll();
    calcAndRender();
  });
})();
