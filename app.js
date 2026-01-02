// Farming CBA Decision Tool 2 — Newcastle Business School
// app.js (fully upgraded)
// Focus: robust CBA engine, Excel-first workflow, vertical results table (rows=metrics, cols=treatments incl control),
// clean exports (Excel/CSV/PDF), and AI-assisted interpretation prompt generation (non-prescriptive, learning-oriented).

(() => {
  "use strict";

  // =========================
  // 0) Naming + Versioning
  // =========================
  const TOOL_NAME = "Farming CBA Decision Tool 2";
  const TOOL_SHORT = "Farming CBA Decision Tool 2";
  const ORG_NAME = "Newcastle Business School";

  // =========================
  // 1) Constants + Defaults
  // =========================
  const DEFAULT_DISCOUNT_SCHEDULE = [
    { label: "2025-2034", from: 2025, to: 2034, low: 2, base: 4, high: 6 },
    { label: "2035-2044", from: 2035, to: 2044, low: 4, base: 7, high: 10 },
    { label: "2045-2054", from: 2045, to: 2054, low: 4, base: 7, high: 10 },
    { label: "2055-2064", from: 2055, to: 2064, low: 3, base: 6, high: 9 },
    { label: "2065-2074", from: 2065, to: 2074, low: 2, base: 5, high: 8 }
  ];

  const HORIZONS = [5, 10, 15, 20, 25];

  const OUTPUT_SOURCES = ["Farm Trials", "Plant Farm", "ABARES", "GRDC", "Input Directly"];
  const TREATMENT_SOURCES = ["Farm Trials", "Plant Farm", "ABARES", "GRDC", "Input Directly"];

  // =========================
  // 2) ID + Utility
  // =========================
  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

  const esc = s =>
    (s ?? "")
      .toString()
      .replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  function fmtNumber(n, maxFrac = 2) {
    if (!isFinite(n)) return "n/a";
    const abs = Math.abs(n);
    if (abs >= 1000000) return n.toLocaleString(undefined, { maximumFractionDigits: 0 });
    if (abs >= 1000) return n.toLocaleString(undefined, { maximumFractionDigits: 0 });
    return n.toLocaleString(undefined, { maximumFractionDigits: maxFrac });
  }
  const money = n => (isFinite(n) ? "$" + fmtNumber(n, 0) : "n/a");
  const money2 = n => (isFinite(n) ? "$" + fmtNumber(n, 2) : "n/a");
  const percent = n => (isFinite(n) ? fmtNumber(n, 2) + "%" : "n/a");

  const slug = s =>
    (s || "project")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_|_$/g, "");

  function parseNumber(value) {
    if (value === null || value === undefined || value === "") return NaN;
    if (typeof value === "number") return value;
    const cleaned = String(value).replace(/[\$,]/g, "").trim();
    if (!cleaned) return NaN;
    const n = parseFloat(cleaned);
    return Number.isFinite(n) ? n : NaN;
  }

  // Deterministic RNG for simulation
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

  function triangular(u, a, c, b) {
    const F = (c - a) / (b - a);
    if (u < F) return a + Math.sqrt(u * (b - a) * (c - a));
    return b - Math.sqrt((1 - u) * (b - a) * (b - c));
  }

  // =========================
  // 3) Toast
  // =========================
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

  // =========================
  // 4) Model (State)
  // =========================
  const model = {
    meta: {
      toolName: TOOL_NAME,
      version: "2.0",
      created: new Date().toISOString()
    },
    project: {
      name: "Faba bean soil amendment trial",
      lead: "Project lead",
      analysts: "Farm economics team",
      team: "Trial team",
      organisation: `${ORG_NAME}, The University of Newcastle`,
      contactEmail: "",
      contactPhone: "",
      summary:
        "Applied faba bean trial comparing deep ripping, organic matter, gypsum and fertiliser treatments against a control.",
      objectives: "Quantify yield and gross margin impacts of alternative soil amendment strategies.",
      activities:
        "Establish replicated field plots, collect plot-level yield and cost data, and summarise trial-wide economics.",
      stakeholders: "Producers, agronomists, government agencies, research partners.",
      lastUpdated: new Date().toISOString().slice(0, 10),
      goal:
        "Identify soil amendment packages that deliver higher yields and acceptable returns after accounting for additional costs.",
      withProject:
        "Growers adopt higher performing amendment packages on suitable soils and validate performance through monitoring.",
      withoutProject:
        "Growers continue baseline practice and rely on partial evidence without a transparent farm-scale CBA comparison."
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
      { id: uid(), name: "Grain yield", unit: "t/ha", value: 450, source: "Input Directly" },
      { id: uid(), name: "Screenings", unit: "percentage point", value: -20, source: "Input Directly" },
      { id: uid(), name: "Protein", unit: "percentage point", value: 10, source: "Input Directly" }
    ],
    treatments: [
      {
        id: uid(),
        name: "Control (no amendment)",
        area: 100,
        adoption: 1,
        deltas: {},
        labourCost: 40,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "Farm Trials",
        isControl: true,
        notes: "Baseline practice without deep soil amendment."
      },
      {
        id: uid(),
        name: "Deep organic matter CP1",
        area: 100,
        adoption: 1,
        deltas: {},
        labourCost: 60,
        materialsCost: 16500,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "Farm Trials",
        isControl: false,
        notes: "Deep incorporation of organic matter at CP1 rate."
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
        consequence: 120000,
        notes: "Project wide operating cost saving"
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
      high: 0.3,
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
      variationPct: 20,
      varyOutputs: true,
      varyTreatCosts: true,
      varyInputCosts: false
    }
  };

  // Ensure deltas exist for all outputs
  function initTreatmentDeltas() {
    model.treatments.forEach(t => {
      if (!t.deltas || typeof t.deltas !== "object") t.deltas = {};
      model.outputs.forEach(o => {
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      });
      if (typeof t.labourCost === "undefined") t.labourCost = Number(t.annualCost || 0) || 0;
      if (typeof t.materialsCost === "undefined") t.materialsCost = 0;
      if (typeof t.servicesCost === "undefined") t.servicesCost = 0;
      if (typeof t.adoption !== "number" || isNaN(t.adoption)) t.adoption = 1;
      delete t.annualCost;
    });
  }
  initTreatmentDeltas();

  // =========================
  // 5) DOM Helpers
  // =========================
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));
  const num = sel => +(document.querySelector(sel)?.value || 0);

  const setText = (sel, text) => {
    const el = document.querySelector(sel);
    if (el) el.textContent = text;
  };

  // =========================
  // 6) Tabs
  // =========================
  function switchTab(target) {
    if (!target) return;

    const navEls = $$("[data-tab],[data-tab-target],[data-tab-jump]");
    navEls.forEach(el => {
      const key = el.dataset.tab || el.dataset.tabTarget || el.dataset.tabJump;
      const isActive = key === target;
      el.classList.toggle("active", isActive);
      if (el.hasAttribute("aria-selected")) el.setAttribute("aria-selected", isActive ? "true" : "false");
    });

    const panels = $$(".tab-panel");
    panels.forEach(p => {
      const key = p.dataset.tabPanel || (p.id ? p.id.replace(/^tab-/, "") : "");
      const match = key === target || p.id === target || p.id === "tab-" + target;
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
      const target = el.dataset.tab || el.dataset.tabTarget || el.dataset.tabJump;
      if (!target) return;
      e.preventDefault();
      switchTab(target);
    });

    const activeNav =
      document.querySelector("[data-tab].active, [data-tab-target].active, [data-tab-jump].active") ||
      document.querySelector("[data-tab], [data-tab-target], [data-tab-jump]");
    if (activeNav) {
      const target = activeNav.dataset.tab || activeNav.dataset.tabTarget || activeNav.dataset.tabJump;
      if (target) switchTab(target);
      return;
    }

    const firstPanel = document.querySelector(".tab-panel");
    if (firstPanel) {
      const key = firstPanel.dataset.tabPanel || (firstPanel.id ? firstPanel.id.replace(/^tab-/, "") : "");
      if (key) switchTab(key);
    }
  }

  // =========================
  // 7) CBA Core
  // =========================
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

  function additionalBenefitsSeries(N, baseYear, adoptMul, risk) {
    const series = new Array(N + 1).fill(0);

    model.benefits.forEach(b => {
      const cat = String(b.category || "").toUpperCase();
      const linkA = !!b.linkAdoption;
      const linkR = !!b.linkRisk;
      const A = linkA ? clamp(adoptMul, 0, 1) : 1;
      const R = linkR ? 1 - clamp(risk, 0, 1) : 1;
      const g = Number(b.growthPct) || 0;

      const addAnnual = (idx, baseAmount, tFromStart) => {
        const grown = baseAmount * Math.pow(1 + g / 100, tFromStart);
        if (idx >= 1 && idx <= N) series[idx] += grown * A * R;
      };
      const addOnce = (absYear, amount) => {
        const idx = absYear - baseYear + 1;
        if (idx >= 0 && idx <= N) series[idx] += amount * A * R;
      };

      const sy = Number(b.startYear) || baseYear;
      const ey = Number(b.endYear) || sy;
      const onceYear = Number(b.year) || sy;

      if (b.frequency === "Once" || cat === "C6") {
        addOnce(onceYear, Number(b.annualAmount) || 0);
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

  // Whole-project cashflows (all treatments combined)
  function buildProjectCashflows({ adoptMul = model.adoption.base, risk = model.risk.base }) {
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

      const annualCostPerHa = (Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0) + (Number(t.labourCost) || 0);
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

    // Other costs
    let otherCapitalY0 = 0;
    let otherConstrCapitalY0 = 0;
    const otherAnnualByYear = new Array(N + 1).fill(0);
    const otherConstrAnnualByYear = new Array(N + 1).fill(0);

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

    const cf = new Array(N + 1).fill(0).map((_, i) => benefitByYear[i] - costByYear[i]);
    const annualGM = annualBenefit - treatAnnualCost;

    return { benefitByYear, costByYear, constrainedCostByYear, cf, annualGM };
  }

  // Per-treatment cashflows (treatment-only, excluding shared otherCosts and shared additional benefits by default)
  function buildTreatmentCashflows(treatment, { adoptMul = model.adoption.base, risk = model.risk.base, includeShared = false } = {}) {
    const N = model.time.years;
    const baseYear = model.time.startYear;

    const benefitByYear = new Array(N + 1).fill(0);
    const costByYear = new Array(N + 1).fill(0);
    const constrainedCostByYear = new Array(N + 1).fill(0);

    const adopt = clamp(adoptMul, 0, 1);
    const area = Number(treatment.area) || 0;

    let valuePerHa = 0;
    model.outputs.forEach(o => {
      const delta = Number(treatment.deltas[o.id]) || 0;
      const v = Number(o.value) || 0;
      valuePerHa += delta * v;
    });

    const annualBenefit = valuePerHa * area * (1 - clamp(risk, 0, 1)) * adopt;

    const annualCostPerHa = (Number(treatment.materialsCost) || 0) + (Number(treatment.servicesCost) || 0) + (Number(treatment.labourCost) || 0);
    const annualCost = annualCostPerHa * area;
    const cap = Number(treatment.capitalCost) || 0;

    costByYear[0] += cap;
    if (treatment.constrained) constrainedCostByYear[0] += cap;

    for (let y = 1; y <= N; y++) {
      benefitByYear[y] += annualBenefit;
      costByYear[y] += annualCost;
      if (treatment.constrained) constrainedCostByYear[y] += annualCost;
    }

    if (includeShared) {
      // Allocate shared items by "exposure weight": area * adoptMul, across all treatments.
      const weights = model.treatments.map(tt => (Number(tt.area) || 0) * clamp(adoptMul, 0, 1));
      const denom = weights.reduce((a, b) => a + b, 0) || 1;
      const w = ((Number(treatment.area) || 0) * clamp(adoptMul, 0, 1)) / denom;

      // Allocate otherCosts (annual and capital)
      model.otherCosts.forEach(c => {
        if (c.type === "annual") {
          const a = (Number(c.annual) || 0) * w;
          const sy = Number(c.startYear) || baseYear;
          const ey = Number(c.endYear) || sy;
          for (let year = sy; year <= ey; year++) {
            const idx = year - baseYear + 1;
            if (idx >= 1 && idx <= N) {
              costByYear[idx] += a;
              if (c.constrained) constrainedCostByYear[idx] += a;
            }
          }
        } else if (c.type === "capital") {
          const cap2 = (Number(c.capital) || 0) * w;
          const cy = Number(c.year) || baseYear;
          const idx = cy - baseYear;
          if (idx >= 0 && idx <= N) {
            costByYear[idx] += cap2;
            if (c.constrained) constrainedCostByYear[idx] += cap2;
          }
        }
      });

      // Allocate additional benefits
      const extra = additionalBenefitsSeries(N, baseYear, adoptMul, risk).map(v => v * w);
      for (let i = 0; i < extra.length; i++) benefitByYear[i] += extra[i];
    }

    const cf = new Array(N + 1).fill(0).map((_, i) => benefitByYear[i] - costByYear[i]);
    const annualGM = annualBenefit - annualCost;

    return { benefitByYear, costByYear, constrainedCostByYear, cf, annualGM };
  }

  function computeMetricsFromCashflows({ benefitByYear, costByYear, constrainedCostByYear, cf, annualGM }, ratePct, bcrMode) {
    const pvBenefits = presentValue(benefitByYear, ratePct);
    const pvCosts = presentValue(costByYear, ratePct);
    const pvCostsConstrained = presentValue(constrainedCostByYear, ratePct);

    const npv = pvBenefits - pvCosts;
    const denom = bcrMode === "constrained" ? pvCostsConstrained : pvCosts;
    const bcr = denom > 0 ? pvBenefits / denom : NaN;

    const irrVal = irr(cf);
    const mirrVal = mirr(cf, model.time.mirrFinance, model.time.mirrReinvest);
    const roi = pvCosts > 0 ? (npv / pvCosts) * 100 : NaN;

    const annualBenefitsApprox = benefitByYear[1] || 0;
    const profitMargin = annualBenefitsApprox > 0 ? (annualGM / annualBenefitsApprox) * 100 : NaN;
    const pb = payback(cf, ratePct);

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
      paybackYears: pb
    };
  }

  // =========================
  // 8) Rendering: Project + Inputs
  // =========================
  function setBasicsFieldsFromModel() {
    // project fields
    if ($("#projectName")) $("#projectName").value = model.project.name || "";
    if ($("#projectLead")) $("#projectLead").value = model.project.lead || "";
    if ($("#analystNames")) $("#analystNames").value = model.project.analysts || "";
    if ($("#projectTeam")) $("#projectTeam").value = model.project.team || "";
    if ($("#organisation")) $("#organisation").value = model.project.organisation || "";
    if ($("#lastUpdated")) $("#lastUpdated").value = model.project.lastUpdated || "";
    if ($("#contactEmail")) $("#contactEmail").value = model.project.contactEmail || "";
    if ($("#contactPhone")) $("#contactPhone").value = model.project.contactPhone || "";
    if ($("#projectSummary")) $("#projectSummary").value = model.project.summary || "";
    if ($("#projectGoal")) $("#projectGoal").value = model.project.goal || "";
    if ($("#withProject")) $("#withProject").value = model.project.withProject || "";
    if ($("#withoutProject")) $("#withoutProject").value = model.project.withoutProject || "";
    if ($("#projectObjectives")) $("#projectObjectives").value = model.project.objectives || "";
    if ($("#projectActivities")) $("#projectActivities").value = model.project.activities || "";
    if ($("#stakeholderGroups")) $("#stakeholderGroups").value = model.project.stakeholders || "";

    // time + settings
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

    if ($("#systemType")) $("#systemType").value = model.outputsMeta.systemType || "single";
    if ($("#outputAssumptions")) $("#outputAssumptions").value = model.outputsMeta.assumptions || "";

    // discount schedule table inputs
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

    // simulation controls
    if ($("#simN")) $("#simN").value = model.sim.n;
    if ($("#targetBCR")) $("#targetBCR").value = model.sim.targetBCR;
    if ($("#bcrMode")) $("#bcrMode").value = model.sim.bcrMode;
    if ($("#simBcrTargetLabel")) $("#simBcrTargetLabel").textContent = model.sim.targetBCR;

    if ($("#simVarPct")) $("#simVarPct").value = String(model.sim.variationPct || 20);
    if ($("#simVaryOutputs")) $("#simVaryOutputs").value = model.sim.varyOutputs ? "true" : "false";
    if ($("#simVaryTreatCosts")) $("#simVaryTreatCosts").value = model.sim.varyTreatCosts ? "true" : "false";
    if ($("#simVaryInputCosts")) $("#simVaryInputCosts").value = model.sim.varyInputCosts ? "true" : "false";
  }

  function bindBasics() {
    setBasicsFieldsFromModel();

    // Combine risk button
    const calcRiskBtn = $("#calcCombinedRisk");
    if (calcRiskBtn) {
      calcRiskBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        const r =
          1 -
          (1 - num("#rTech")) *
            (1 - num("#rNonCoop")) *
            (1 - num("#rSocio")) *
            (1 - num("#rFin")) *
            (1 - num("#rMan"));
        model.risk.base = clamp(r, 0, 1);
        if ($("#riskBase")) $("#riskBase").value = model.risk.base.toFixed(3);
        const out = $("#combinedRiskOut");
        if (out) out.innerHTML = `<div class="label">Combined risk</div><div class="value">${percent(model.risk.base * 100)}</div>`;
        calcAndRenderDebounced();
        showToast("Combined risk updated.");
      });
    }

    // Global input binding
    document.addEventListener("input", e => {
      const t = e.target;
      if (!t) return;

      // Discount schedule inputs
      if (t.dataset && t.dataset.discPeriod !== undefined) {
        const idx = +t.dataset.discPeriod;
        const scenario = t.dataset.scenario;
        if (!model.time.discountSchedule) model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
        const row = model.time.discountSchedule[idx];
        if (row && scenario) {
          const val = +t.value;
          if (scenario === "low") row.low = val;
          else if (scenario === "base") row.base = val;
          else if (scenario === "high") row.high = val;
        }
        calcAndRenderDebounced();
        return;
      }

      const id = t.id;
      if (!id) return;

      switch (id) {
        case "projectName":
          model.project.name = t.value;
          updateBrandingText();
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
        case "organisation":
          model.project.organisation = t.value;
          break;
        case "lastUpdated":
          model.project.lastUpdated = t.value;
          break;
        case "contactEmail":
          model.project.contactEmail = t.value;
          break;
        case "contactPhone":
          model.project.contactPhone = t.value;
          break;
        case "projectSummary":
          model.project.summary = t.value;
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
        case "projectObjectives":
          model.project.objectives = t.value;
          break;
        case "projectActivities":
          model.project.activities = t.value;
          break;
        case "stakeholderGroups":
          model.project.stakeholders = t.value;
          break;

        case "startYear":
          model.time.startYear = +t.value;
          break;
        case "projectStartYear":
          model.time.projectStartYear = +t.value;
          break;
        case "years":
          model.time.years = Math.max(1, +t.value || 1);
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
          model.adoption.base = clamp(+t.value, 0, 1);
          break;
        case "adoptLow":
          model.adoption.low = clamp(+t.value, 0, 1);
          break;
        case "adoptHigh":
          model.adoption.high = clamp(+t.value, 0, 1);
          break;

        case "riskBase":
          model.risk.base = clamp(+t.value, 0, 1);
          break;
        case "riskLow":
          model.risk.low = clamp(+t.value, 0, 1);
          break;
        case "riskHigh":
          model.risk.high = clamp(+t.value, 0, 1);
          break;

        case "rTech":
          model.risk.tech = clamp(+t.value, 0, 1);
          break;
        case "rNonCoop":
          model.risk.nonCoop = clamp(+t.value, 0, 1);
          break;
        case "rSocio":
          model.risk.socio = clamp(+t.value, 0, 1);
          break;
        case "rFin":
          model.risk.fin = clamp(+t.value, 0, 1);
          break;
        case "rMan":
          model.risk.man = clamp(+t.value, 0, 1);
          break;

        case "simN":
          model.sim.n = Math.max(100, +t.value || 100);
          break;
        case "targetBCR":
          model.sim.targetBCR = +t.value || 2;
          if ($("#simBcrTargetLabel")) $("#simBcrTargetLabel").textContent = model.sim.targetBCR;
          break;
        case "bcrMode":
          model.sim.bcrMode = t.value;
          break;
        case "randSeed":
          model.sim.seed = t.value ? +t.value : null;
          break;

        case "simVarPct":
          model.sim.variationPct = Math.max(0, +t.value || 0);
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

    // Start button
    const startBtn = $("#startBtn");
    if (startBtn) {
      startBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        switchTab("project");
        showToast("Start with Project setup.");
      });
    }
    const startBtnDup = $("#startBtn-duplicate");
    if (startBtnDup) {
      startBtnDup.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        switchTab("project");
      });
    }

    // Recalculate
    const recalcBtn = $("#recalc");
    if (recalcBtn) {
      recalcBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        calcAndRender();
        showToast("Results recalculated.");
      });
    }

    // Run simulation
    const runSimBtn = $("#runSim");
    if (runSimBtn) {
      runSimBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        runSimulation();
      });
    }

    // Save/load project JSON
    const saveProjectBtn = $("#saveProject");
    if (saveProjectBtn) {
      saveProjectBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        const data = JSON.stringify(model, null, 2);
        downloadFile(`${slug(model.project.name)}_${slug(TOOL_SHORT)}.json`, data, "application/json");
        showToast("Project file downloaded.");
      });
    }

    const loadProjectBtn = $("#loadProject");
    const loadFileInput = $("#loadFile");
    if (loadProjectBtn && loadFileInput) {
      loadProjectBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        loadFileInput.click();
      });

      loadFileInput.addEventListener("change", async e => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        const text = await file.text();
        try {
          const obj = JSON.parse(text);
          // shallow merge, preserving our defaults if missing
          Object.assign(model, obj);
          if (!model.meta) model.meta = { toolName: TOOL_NAME, version: "2.0", created: new Date().toISOString() };
          model.meta.toolName = TOOL_NAME;

          if (!model.time.discountSchedule) model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
          if (!model.sim) model.sim = { n: 1000, targetBCR: 2, bcrMode: "all", seed: null, results: { npv: [], bcr: [] } };

          initTreatmentDeltas();
          renderAll();
          setBasicsFieldsFromModel();
          calcAndRender();
          showToast("Project loaded.");
        } catch (err) {
          alert("Invalid project JSON.");
          console.error(err);
        } finally {
          e.target.value = "";
        }
      });
    }

    // Exports: CSV / PDF
    const exportCsvBtn = $("#exportCsv");
    const exportCsvFootBtn = $("#exportCsvFoot");
    const exportPdfBtn = $("#exportPdf");
    const exportPdfFootBtn = $("#exportPdfFoot");

    if (exportCsvBtn) exportCsvBtn.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); exportResultsCsv(); });
    if (exportCsvFootBtn) exportCsvFootBtn.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); exportResultsCsv(); });

    if (exportPdfBtn) exportPdfBtn.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); exportPdf(); });
    if (exportPdfFootBtn) exportPdfFootBtn.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); exportPdf(); });

    // Excel tab: template/download/parse/commit
    const downloadTemplateBtn = $("#downloadTemplate");
    if (downloadTemplateBtn) downloadTemplateBtn.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); downloadExcelTemplate({ mode: "blank" }); });

    const downloadSampleBtn = $("#downloadSample");
    if (downloadSampleBtn) downloadSampleBtn.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); downloadExcelTemplate({ mode: "current" }); });

    const parseExcelBtn = $("#parseExcel");
    if (parseExcelBtn) parseExcelBtn.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); handleParseExcel(); });

    const importExcelBtn = $("#importExcel");
    if (importExcelBtn) importExcelBtn.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); commitExcelToModel(); });

    // Copilot helper
    const openCopilotBtn = $("#openCopilot");
    if (openCopilotBtn) openCopilotBtn.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); handleCopyAIPrompt(); });
  }

  // Debounce for frequent edits
  let _calcTimer = null;
  function calcAndRenderDebounced() {
    clearTimeout(_calcTimer);
    _calcTimer = setTimeout(() => calcAndRender(), 250);
  }

  // =========================
  // 9) Render: Outputs / Treatments / Benefits / Costs / DB tags
  // =========================
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
          <div class="field"><label>Name</label><input value="${esc(o.name)}" data-ok="name" data-id="${o.id}" /></div>
          <div class="field"><label>Unit</label><input value="${esc(o.unit)}" data-ok="unit" data-id="${o.id}" /></div>
          <div class="field"><label>Value ($/unit)</label><input type="number" step="0.01" value="${Number(o.value) || 0}" data-ok="value" data-id="${o.id}" /></div>
          <div class="field"><label>Source</label>
            <select data-ok="source" data-id="${o.id}">
              ${OUTPUT_SOURCES.map(s => `<option ${s === o.source ? "selected" : ""}>${esc(s)}</option>`).join("")}
            </select>
          </div>
          <div class="field"><label>&nbsp;</label><button class="btn small danger" data-del-output="${o.id}">Remove</button></div>
        </div>
        <div class="kv"><small class="muted">id:</small> <code>${esc(o.id)}</code></div>
      `;
      root.appendChild(el);
    });

    root.oninput = e => {
      const id = e.target.dataset.id;
      const k = e.target.dataset.ok;
      if (!id || !k) return;
      const o = model.outputs.find(x => x.id === id);
      if (!o) return;
      if (k === "value") o.value = +e.target.value;
      else o[k] = e.target.value;

      model.treatments.forEach(t => {
        if (!(id in t.deltas)) t.deltas[id] = 0;
      });

      renderTreatments(); // keeps deltas aligned
      renderDatabaseTags();
      calcAndRenderDebounced();
    };

    root.onclick = e => {
      const id = e.target.dataset.delOutput;
      if (!id) return;
      if (!confirm("Remove this output metric?")) return;
      model.outputs = model.outputs.filter(o => o.id !== id);
      model.treatments.forEach(t => { delete t.deltas[id]; });
      renderOutputs();
      renderTreatments();
      renderDatabaseTags();
      calcAndRender();
      showToast("Output removed.");
    };
  }

  function renderTreatments() {
    const root = $("#treatmentsList");
    if (!root) return;
    root.innerHTML = "";

    model.treatments.forEach(t => {
      const mats = Number(t.materialsCost) || 0;
      const serv = Number(t.servicesCost) || 0;
      const lab = Number(t.labourCost) || 0;
      const totalPerHa = mats + serv + lab;

      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>Treatment: ${esc(t.name)}</h4>
        <div class="row">
          <div class="field"><label>Name</label><input value="${esc(t.name)}" data-tk="name" data-id="${t.id}" /></div>
          <div class="field"><label>Area (ha)</label><input type="number" step="0.01" value="${Number(t.area) || 0}" data-tk="area" data-id="${t.id}" /></div>
          <div class="field"><label>Source</label>
            <select data-tk="source" data-id="${t.id}">
              ${TREATMENT_SOURCES.map(s => `<option ${s === t.source ? "selected" : ""}>${esc(s)}</option>`).join("")}
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
          <div class="field"><label>Materials cost ($/ha)</label><input type="number" step="0.01" value="${mats}" data-tk="materialsCost" data-id="${t.id}" /></div>
          <div class="field"><label>Services cost ($/ha)</label><input type="number" step="0.01" value="${serv}" data-tk="servicesCost" data-id="${t.id}" /></div>
          <div class="field"><label>Labour cost ($/ha)</label><input type="number" step="0.01" value="${lab}" data-tk="labourCost" data-id="${t.id}" /></div>
          <div class="field"><label>Total cost ($/ha)</label><input type="number" step="0.01" value="${totalPerHa}" readonly data-total-cost="${t.id}" /></div>
          <div class="field"><label>Capital cost ($, year 0)</label><input type="number" step="0.01" value="${Number(t.capitalCost) || 0}" data-tk="capitalCost" data-id="${t.id}" /></div>
          <div class="field"><label>Constrained?</label>
            <select data-tk="constrained" data-id="${t.id}">
              <option value="true" ${t.constrained ? "selected" : ""}>Yes</option>
              <option value="false" ${!t.constrained ? "selected" : ""}>No</option>
            </select>
          </div>
        </div>

        <div class="field">
          <label>Notes (definition, implementation details, or control definition)</label>
          <textarea data-tk="notes" data-id="${t.id}" rows="2">${esc(t.notes || "")}</textarea>
        </div>

        <h5>Output deltas (per ha)</h5>
        <div class="row">
          ${model.outputs
            .map(
              o => `
            <div class="field">
              <label>${esc(o.name)} (${esc(o.unit)})</label>
              <input type="number" step="0.0001" value="${Number(t.deltas[o.id] ?? 0)}" data-td="${o.id}" data-id="${t.id}" />
            </div>
          `
            )
            .join("")}
        </div>

        <div class="kv"><small class="muted">id:</small> <code>${esc(t.id)}</code></div>
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
        if (tk === "constrained") t.constrained = e.target.value === "true";
        else if (tk === "isControl") {
          const val = e.target.value === "control";
          model.treatments.forEach(tt => { tt.isControl = false; });
          if (val) t.isControl = true;
          renderTreatments();
          calcAndRenderDebounced();
          showToast(`Control set to: ${t.name}`);
          return;
        } else if (tk === "name" || tk === "source" || tk === "notes") t[tk] = e.target.value;
        else t[tk] = +e.target.value;

        if (tk === "materialsCost" || tk === "servicesCost" || tk === "labourCost") {
          const container = e.target.closest(".item");
          if (container) {
            const mats = Number(container.querySelector(`input[data-tk="materialsCost"][data-id="${id}"]`)?.value || 0);
            const serv = Number(container.querySelector(`input[data-tk="servicesCost"][data-id="${id}"]`)?.value || 0);
            const lab = Number(container.querySelector(`input[data-tk="labourCost"][data-id="${id}"]`)?.value || 0);
            const totalField = container.querySelector(`input[data-total-cost="${id}"]`);
            if (totalField) totalField.value = mats + serv + lab;
          }
        }
      }

      const td = e.target.dataset.td;
      if (td) t.deltas[td] = +e.target.value;

      renderDatabaseTags();
      calcAndRenderDebounced();
    };

    root.onclick = e => {
      const id = e.target.dataset.delTreatment;
      if (!id) return;
      if (!confirm("Remove this treatment?")) return;
      model.treatments = model.treatments.filter(x => x.id !== id);
      if (!model.treatments.some(t => t.isControl) && model.treatments.length) model.treatments[0].isControl = true;
      renderTreatments();
      renderDatabaseTags();
      calcAndRender();
      showToast("Treatment removed.");
    };
  }

  function renderBenefits() {
    const root = $("#benefitsList");
    if (!root) return;

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

    root.innerHTML = "";
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
                .map(c => `<option ${c === b.category ? "selected" : ""}>${esc(c)}</option>`)
                .join("")}
            </select>
          </div>
          <div class="field"><label>Benefit type</label>
            <select data-bk="theme" data-id="${b.id}">
              ${THEMES.map(th => `<option ${th === (b.theme || "") ? "selected" : ""}>${esc(th)}</option>`).join("")}
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
          <div class="field"><label>Unit value ($)</label><input type="number" step="0.01" value="${Number(b.unitValue) || 0}" data-bk="unitValue" data-id="${b.id}" /></div>
          <div class="field"><label>Quantity</label><input type="number" step="0.01" value="${Number(b.quantity) || 0}" data-bk="quantity" data-id="${b.id}" /></div>
          <div class="field"><label>Abatement</label><input type="number" step="0.01" value="${Number(b.abatement) || 0}" data-bk="abatement" data-id="${b.id}" /></div>
          <div class="field"><label>Annual amount ($)</label><input type="number" step="0.01" value="${Number(b.annualAmount) || 0}" data-bk="annualAmount" data-id="${b.id}" /></div>
          <div class="field"><label>Growth (% per year)</label><input type="number" step="0.01" value="${Number(b.growthPct) || 0}" data-bk="growthPct" data-id="${b.id}" /></div>
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
          <div class="field"><label>P0 (baseline probability)</label><input type="number" step="0.001" value="${Number(b.p0) || 0}" data-bk="p0" data-id="${b.id}" /></div>
          <div class="field"><label>P1 (with project probability)</label><input type="number" step="0.001" value="${Number(b.p1) || 0}" data-bk="p1" data-id="${b.id}" /></div>
          <div class="field"><label>Consequence ($)</label><input type="number" step="0.01" value="${Number(b.consequence) || 0}" data-bk="consequence" data-id="${b.id}" /></div>
          <div class="field"><label>Notes</label><input value="${esc(b.notes || "")}" data-bk="notes" data-id="${b.id}" /></div>
          <div class="field"><label>&nbsp;</label><button class="btn small danger" data-del-benefit="${b.id}">Remove</button></div>
        </div>
      `;
      root.appendChild(el);
    });

    root.oninput = e => {
      const id = e.target.dataset.id;
      const k = e.target.dataset.bk;
      if (!id || !k) return;
      const b = model.benefits.find(x => x.id === id);
      if (!b) return;

      if (["label", "category", "frequency", "notes", "theme"].includes(k)) b[k] = e.target.value;
      else if (k === "linkAdoption" || k === "linkRisk") b[k] = e.target.value === "true";
      else b[k] = +e.target.value;

      calcAndRenderDebounced();
    };

    root.onclick = e => {
      const id = e.target.dataset.delBenefit;
      if (!id) return;
      if (!confirm("Remove this benefit item?")) return;
      model.benefits = model.benefits.filter(x => x.id !== id);
      renderBenefits();
      calcAndRender();
      showToast("Benefit removed.");
    };
  }

  function renderCosts() {
    const root = $("#costsList");
    if (!root) return;
    root.innerHTML = "";

    model.otherCosts.forEach(c => {
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
          <div class="field"><label>Annual ($ per year)</label><input type="number" step="0.01" value="${Number(c.annual) || 0}" data-ck="annual" data-id="${c.id}" /></div>
          <div class="field"><label>Start year</label><input type="number" value="${Number(c.startYear) || model.time.startYear}" data-ck="startYear" data-id="${c.id}" /></div>
          <div class="field"><label>End year</label><input type="number" value="${Number(c.endYear) || model.time.startYear}" data-ck="endYear" data-id="${c.id}" /></div>
        </div>

        <div class="row-6">
          <div class="field"><label>Capital ($)</label><input type="number" step="0.01" value="${Number(c.capital) || 0}" data-ck="capital" data-id="${c.id}" /></div>
          <div class="field"><label>Capital year</label><input type="number" value="${Number(c.year) || model.time.startYear}" data-ck="year" data-id="${c.id}" /></div>
          <div class="field"><label>Constrained?</label>
            <select data-ck="constrained" data-id="${c.id}">
              <option value="true" ${c.constrained ? "selected" : ""}>Yes</option>
              <option value="false" ${!c.constrained ? "selected" : ""}>No</option>
            </select>
          </div>
          <div class="field"><label>Depreciation method</label>
            <select data-ck="depMethod" data-id="${c.id}">
              <option value="none" ${c.depMethod === "none" ? "selected" : ""}>None</option>
              <option value="straight" ${c.depMethod === "straight" ? "selected" : ""}>Straight line</option>
              <option value="declining" ${c.depMethod === "declining" ? "selected" : ""}>Declining balance</option>
            </select>
          </div>
          <div class="field"><label>Depreciation life (years)</label><input type="number" value="${Number(c.depLife) || 5}" data-ck="depLife" data-id="${c.id}" /></div>
          <div class="field"><label>Declining rate (% p.a.)</label><input type="number" step="0.1" value="${Number(c.depRate) || 30}" data-ck="depRate" data-id="${c.id}" /></div>
        </div>

        <div class="row-2">
          <div class="field"><button class="btn small danger" data-del-cost="${c.id}">Remove</button></div>
          <div class="field small muted">Tip: Use annual for repeated costs and capital for one-off investments.</div>
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

      if (k === "label" || k === "type" || k === "category" || k === "depMethod") c[k] = e.target.value;
      else if (k === "constrained") c.constrained = e.target.value === "true";
      else c[k] = +e.target.value;

      calcAndRenderDebounced();
    };

    root.onclick = e => {
      const id = e.target.dataset.delCost;
      if (!id) return;
      if (!confirm("Remove this cost item?")) return;
      model.otherCosts = model.otherCosts.filter(x => x.id !== id);
      renderCosts();
      calcAndRender();
      showToast("Cost removed.");
    };
  }

  function renderDatabaseTags() {
    const outRoot = $("#dbOutputs");
    const trRoot = $("#dbTreatments");
    if (!outRoot || !trRoot) return;

    outRoot.innerHTML = "";
    model.outputs.forEach(o => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>${esc(o.name)}</h4>
        <div class="small muted">Source: ${esc(o.source || "n/a")} · Unit: ${esc(o.unit || "")} · Value: ${money2(Number(o.value) || 0)}</div>
      `;
      outRoot.appendChild(el);
    });

    trRoot.innerHTML = "";
    model.treatments.forEach(t => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>${esc(t.name)} ${t.isControl ? "<span class='badge'>Control</span>" : ""}</h4>
        <div class="small muted">Source: ${esc(t.source || "n/a")} · Area: ${fmtNumber(Number(t.area) || 0)} ha</div>
      `;
      trRoot.appendChild(el);
    });
  }

  function initAddButtons() {
    const addOutputBtn = $("#addOutput");
    if (addOutputBtn) {
      addOutputBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        const id = uid();
        model.outputs.push({ id, name: "Custom output", unit: "unit", value: 0, source: "Input Directly" });
        model.treatments.forEach(t => { t.deltas[id] = 0; });
        renderOutputs();
        renderTreatments();
        renderDatabaseTags();
        calcAndRender();
        showToast("Output added.");
      });
    }

    const addTreatmentBtn = $("#addTreatment");
    if (addTreatmentBtn) {
      addTreatmentBtn.addEventListener("click", e => {
        e.preventDefault();
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
        model.outputs.forEach(o => { t.deltas[o.id] = 0; });
        model.treatments.push(t);
        renderTreatments();
        renderDatabaseTags();
        calcAndRender();
        showToast("Treatment added.");
      });
    }

    const addBenefitBtn = $("#addBenefit");
    if (addBenefitBtn) {
      addBenefitBtn.addEventListener("click", e => {
        e.preventDefault();
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
        showToast("Benefit added.");
      });
    }

    const addCostBtn = $("#addCost");
    if (addCostBtn) {
      addCostBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        model.otherCosts.push({
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
        });
        renderCosts();
        calcAndRender();
        showToast("Cost added.");
      });
    }
  }

  // =========================
  // 10) Results: Snapshot + Vertical Table (rows metrics, cols treatments)
  // =========================
  let resultsCache = {
    project: null,
    perTreatment: null,
    control: null,
    includeShared: false
  };

  function ensureResultsScaffold() {
    const tab = $("#tab-results .card");
    if (!tab) return;

    // Insert new V2 results block near top of results tab.
    if (!$("#v2ResultsBlock")) {
      const block = document.createElement("div");
      block.id = "v2ResultsBlock";
      block.className = "card subtle";
      block.innerHTML = `
        <h3>${esc(TOOL_SHORT)} — snapshot comparison</h3>
        <p class="small muted">
          This view compares every treatment against the flagged control using a vertical table:
          indicators are rows and treatments are columns. Use the toggle below to optionally allocate shared project-wide
          benefits and costs across treatments.
        </p>
        <div class="row-3">
          <div class="field">
            <label>Comparison mode</label>
            <select id="v2IncludeShared">
              <option value="false">Treatment-only (recommended for clarity)</option>
              <option value="true">Allocate shared project-wide items across treatments (advanced)</option>
            </select>
            <div class="small muted">
              Treatment-only uses outputs and treatment costs; shared items remain in the whole-project summary below.
            </div>
          </div>
          <div class="field">
            <label>Copy-friendly outputs</label>
            <button id="v2CopyTable" class="btn small">Copy results table</button>
            <button id="v2ExportXlsx" class="btn small ghost">Export Excel (XLSX)</button>
          </div>
          <div class="field">
            <label>Headline</label>
            <div id="v2Headline" class="small muted">Not calculated yet.</div>
          </div>
        </div>

        <div class="table-scroll" style="margin-top:12px;">
          <div id="v2ResultsTableWrap"></div>
        </div>

        <div class="small muted" style="margin-top:10px;">
          Notes: NPV is PV(benefits) minus PV(costs). BCR is PV(benefits) divided by PV(costs) (or constrained PV costs if selected in Simulation tab).
          ROI is NPV as a percentage of PV(costs). Ranking is by BCR by default.
        </div>
      `;

      // Place immediately before the existing "Whole project summary" header
      const anchor = tab.querySelector("h3") || tab.firstChild;
      tab.insertBefore(block, anchor);

      // Bind new controls
      const sel = $("#v2IncludeShared");
      if (sel) {
        sel.value = resultsCache.includeShared ? "true" : "false";
        sel.addEventListener("change", () => {
          resultsCache.includeShared = sel.value === "true";
          calcAndRender();
        });
      }
      const copyBtn = $("#v2CopyTable");
      if (copyBtn) copyBtn.addEventListener("click", () => copyResultsTableToClipboard());
      const xlsxBtn = $("#v2ExportXlsx");
      if (xlsxBtn) xlsxBtn.addEventListener("click", () => exportResultsXlsx());
    }
  }

  function computePerTreatmentResults({ includeShared }) {
    const rate = model.time.discBase;
    const bcrMode = model.sim.bcrMode || "all";
    const adoptMul = model.adoption.base;
    const risk = model.risk.base;

    const control = model.treatments.find(t => t.isControl) || model.treatments[0] || null;
    if (!control) return { control: null, rows: [] };

    const tMetrics = [];

    for (const t of model.treatments) {
      const flows = buildTreatmentCashflows(t, { adoptMul, risk, includeShared });
      const m = computeMetricsFromCashflows(flows, rate, bcrMode);
      tMetrics.push({ t, flows, m });
    }

    const controlEntry = tMetrics.find(x => x.t.id === control.id) || tMetrics[0];
    const ctrlM = controlEntry.m;

    // Ranking rule: BCR, then NPV as tiebreak
    const ranked = [...tMetrics].sort((a, b) => {
      const bcrA = isFinite(a.m.bcr) ? a.m.bcr : -Infinity;
      const bcrB = isFinite(b.m.bcr) ? b.m.bcr : -Infinity;
      if (bcrB !== bcrA) return bcrB - bcrA;
      const npvA = isFinite(a.m.npv) ? a.m.npv : -Infinity;
      const npvB = isFinite(b.m.npv) ? b.m.npv : -Infinity;
      return npvB - npvA;
    });
    const rankMap = new Map();
    ranked.forEach((x, i) => rankMap.set(x.t.id, i + 1));

    // Build table rows (metrics as rows)
    const cols = tMetrics.map(x => x.t);

    const rowSpecs = [
      { key: "pvBenefits", label: "Present value of benefits", fmt: money },
      { key: "pvCosts", label: "Present value of costs", fmt: money },
      { key: "npv", label: "Net present value", fmt: money },
      { key: "bcr", label: "Benefit–cost ratio", fmt: v => (isFinite(v) ? fmtNumber(v, 3) : "n/a") },
      { key: "roi", label: "Return on investment (ROI)", fmt: v => percent(v) },
      { key: "irrVal", label: "Internal rate of return (IRR)", fmt: v => percent(v) },
      { key: "mirrVal", label: "Modified IRR (MIRR)", fmt: v => percent(v) },
      { key: "paybackYears", label: "Payback period (years)", fmt: v => (v === null || v === undefined ? "n/a" : String(v)) },
      { key: "__rank__", label: "Ranking (by BCR)", fmt: v => String(v) },
      { key: "__npvDiff__", label: "NPV difference vs control", fmt: money },
      { key: "__bcrDiff__", label: "BCR difference vs control", fmt: v => (isFinite(v) ? fmtNumber(v, 3) : "n/a") }
    ];

    const table = {
      columns: cols,
      rows: rowSpecs.map(r => {
        const values = cols.map(colT => {
          const entry = tMetrics.find(x => x.t.id === colT.id);
          if (!entry) return null;

          if (r.key === "__rank__") return rankMap.get(colT.id) || null;

          if (r.key === "__npvDiff__") {
            const v = entry.m.npv - ctrlM.npv;
            return v;
          }
          if (r.key === "__bcrDiff__") {
            const v = entry.m.bcr - ctrlM.bcr;
            return v;
          }

          return entry.m[r.key];
        });
        return { ...r, values };
      })
    };

    return { control: controlEntry, tMetrics, table, rankMap };
  }

  function renderTreatmentComparisonTable(perTreatment) {
    const wrap = $("#v2ResultsTableWrap");
    if (!wrap) return;

    if (!perTreatment || !perTreatment.table || !perTreatment.control) {
      wrap.innerHTML = `<div class="small muted">No treatments found. Add at least one treatment and mark a control.</div>`;
      return;
    }

    const controlId = perTreatment.control.t.id;

    const cols = perTreatment.table.columns;
    const rows = perTreatment.table.rows;

    // Table HTML
    const head = `
      <thead>
        <tr>
          <th style="position:sticky;left:0;background:var(--panel, #fff);z-index:3;">Indicator</th>
          ${cols
            .map(c => {
              const badge = c.id === controlId ? " <span class='badge'>Control</span>" : "";
              return `<th>${esc(c.name)}${badge}</th>`;
            })
            .join("")}
        </tr>
      </thead>
    `;

    const body = `
      <tbody>
        ${rows
          .map(r => {
            const tds = r.values
              .map((v, j) => {
                const isCtrl = cols[j].id === controlId;
                let cell = r.fmt ? r.fmt(v) : (v ?? "n/a");
                // Light emphasis: control column and negative NPVs
                const className = [
                  isCtrl ? "cell-control" : "",
                  r.key === "npv" && isFinite(v) && v < 0 ? "cell-bad" : "",
                  r.key === "__npvDiff__" && isFinite(v) && v < 0 ? "cell-bad" : "",
                  r.key === "bcr" && isFinite(v) && v < 1 ? "cell-warn" : ""
                ]
                  .filter(Boolean)
                  .join(" ");
                return `<td class="${className}">${esc(cell)}</td>`;
              })
              .join("");
            return `
              <tr>
                <th style="position:sticky;left:0;background:var(--panel, #fff);z-index:2;">${esc(r.label)}</th>
                ${tds}
              </tr>
            `;
          })
          .join("")}
      </tbody>
    `;

    wrap.innerHTML = `
      <table class="summary-table" id="v2ResultsTable">
        ${head}
        ${body}
      </table>
    `;
  }

  function renderHeadline(perTreatment) {
    const el = $("#v2Headline");
    if (!el) return;

    if (!perTreatment || !perTreatment.tMetrics || perTreatment.tMetrics.length === 0) {
      el.textContent = "Not available.";
      return;
    }

    const controlId = perTreatment.control?.t?.id;
    const nonControl = perTreatment.tMetrics.filter(x => x.t.id !== controlId);

    const bestByBcr = [...nonControl].sort((a, b) => (b.m.bcr || -Infinity) - (a.m.bcr || -Infinity))[0];
    const bestByNpv = [...nonControl].sort((a, b) => (b.m.npv || -Infinity) - (a.m.npv || -Infinity))[0];

    const worstByNpv = [...nonControl].sort((a, b) => (a.m.npv || Infinity) - (b.m.npv || Infinity))[0];

    const parts = [];
    if (bestByBcr) parts.push(`Highest BCR: ${bestByBcr.t.name} (BCR ${fmtNumber(bestByBcr.m.bcr, 3)}, NPV ${money(bestByBcr.m.npv)})`);
    if (bestByNpv) parts.push(`Highest NPV: ${bestByNpv.t.name} (NPV ${money(bestByNpv.m.npv)}, BCR ${fmtNumber(bestByNpv.m.bcr, 3)})`);
    if (worstByNpv) parts.push(`Lowest NPV: ${worstByNpv.t.name} (NPV ${money(worstByNpv.m.npv)})`);

    el.textContent = parts.join(" · ");
  }

  // =========================
  // 11) Existing Results Widgets + Tables
  // =========================
  function renderWholeProjectSummary(projectMetrics) {
    if (!projectMetrics) return;

    setText("#pvBenefits", money(projectMetrics.pvBenefits));
    setText("#pvCosts", money(projectMetrics.pvCosts));
    setText("#npv", money(projectMetrics.npv));
    setText("#bcr", isFinite(projectMetrics.bcr) ? fmtNumber(projectMetrics.bcr, 3) : "n/a");

    setText("#irr", percent(projectMetrics.irrVal));
    setText("#mirr", percent(projectMetrics.mirrVal));
    setText("#roi", percent(projectMetrics.roi));
    setText("#payback", projectMetrics.paybackYears === null ? "n/a" : String(projectMetrics.paybackYears));
    setText("#grossMargin", money(projectMetrics.annualGM));
    setText("#profitMargin", percent(projectMetrics.profitMargin));
  }

  function renderTimeProjection(projectFlows) {
    const table = $("#timeProjectionTable tbody");
    if (!table || !projectFlows) return;
    table.innerHTML = "";

    const rate = model.time.discBase;
    const bcrMode = model.sim.bcrMode || "all";

    const base = projectFlows;
    const N = base.benefitByYear.length - 1;

    const points = [];

    for (const h of HORIZONS) {
      const hh = Math.min(h, N);
      const b = base.benefitByYear.slice(0, hh + 1);
      const c = base.costByYear.slice(0, hh + 1);
      const cc = base.constrainedCostByYear.slice(0, hh + 1);
      const cf = b.map((v, i) => v - c[i]);

      const pvB = presentValue(b, rate);
      const pvC = presentValue(c, rate);
      const pvCC = presentValue(cc, rate);
      const npv = pvB - pvC;
      const denom = bcrMode === "constrained" ? pvCC : pvC;
      const bcr = denom > 0 ? pvB / denom : NaN;

      points.push({ h: hh, npv });

      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${hh}</td>
        <td>${money(pvB)}</td>
        <td>${money(pvC)}</td>
        <td>${money(npv)}</td>
        <td>${isFinite(bcr) ? fmtNumber(bcr, 3) : "n/a"}</td>
      `;
      table.appendChild(tr);
    }

    // basic chart on canvas
    const canvas = $("#timeNpvChart");
    if (canvas) drawLineChart(canvas, points.map(p => ({ x: p.h, y: p.npv })), "NPV by horizon");
  }

  // =========================
  // 12) Charts (no external libs)
  // =========================
  function clearCanvas(ctx, w, h) {
    ctx.clearRect(0, 0, w, h);
  }

  function drawLineChart(canvas, points, title) {
    if (!canvas || !points || points.length === 0) return;
    const ctx = canvas.getContext("2d");
    const w = canvas.width;
    const h = canvas.height;
    clearCanvas(ctx, w, h);

    const pad = 36;
    const xs = points.map(p => p.x);
    const ys = points.map(p => p.y);
    const xmin = Math.min(...xs);
    const xmax = Math.max(...xs);
    const ymin = Math.min(...ys);
    const ymax = Math.max(...ys);

    const xScale = x => pad + ((x - xmin) / (xmax - xmin || 1)) * (w - 2 * pad);
    const yScale = y => h - pad - ((y - ymin) / (ymax - ymin || 1)) * (h - 2 * pad);

    // axes
    ctx.beginPath();
    ctx.moveTo(pad, pad);
    ctx.lineTo(pad, h - pad);
    ctx.lineTo(w - pad, h - pad);
    ctx.stroke();

    // title
    ctx.font = "12px sans-serif";
    ctx.fillText(title || "", pad, 16);

    // line
    ctx.beginPath();
    points.forEach((p, i) => {
      const x = xScale(p.x);
      const y = yScale(p.y);
      if (i === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });
    ctx.stroke();

    // markers
    points.forEach(p => {
      const x = xScale(p.x);
      const y = yScale(p.y);
      ctx.beginPath();
      ctx.arc(x, y, 3, 0, Math.PI * 2);
      ctx.fill();
    });

    // labels (x)
    ctx.fillText(String(xmin), xScale(xmin), h - 12);
    ctx.fillText(String(xmax), xScale(xmax) - 10, h - 12);
  }

  function drawHistogram(canvas, values, title) {
    if (!canvas || !values || values.length === 0) return;
    const ctx = canvas.getContext("2d");
    const w = canvas.width;
    const h = canvas.height;
    clearCanvas(ctx, w, h);

    const pad = 36;
    const clean = values.filter(v => isFinite(v));
    if (!clean.length) return;

    const minV = Math.min(...clean);
    const maxV = Math.max(...clean);
    const bins = 18;
    const counts = new Array(bins).fill(0);
    const step = (maxV - minV) / bins || 1;

    clean.forEach(v => {
      const idx = Math.min(bins - 1, Math.max(0, Math.floor((v - minV) / step)));
      counts[idx]++;
    });

    const maxC = Math.max(...counts) || 1;

    // axes
    ctx.beginPath();
    ctx.moveTo(pad, pad);
    ctx.lineTo(pad, h - pad);
    ctx.lineTo(w - pad, h - pad);
    ctx.stroke();

    ctx.font = "12px sans-serif";
    ctx.fillText(title || "", pad, 16);

    // bars
    const barW = (w - 2 * pad) / bins;
    counts.forEach((c, i) => {
      const x = pad + i * barW;
      const bh = (c / maxC) * (h - 2 * pad);
      const y = h - pad - bh;
      ctx.fillRect(x, y, Math.max(1, barW - 2), bh);
    });

    ctx.fillText(fmtNumber(minV, 2), pad, h - 12);
    ctx.fillText(fmtNumber(maxV, 2), w - pad - 40, h - 12);
  }

  // =========================
  // 13) Simulation
  // =========================
  function runSimulation() {
    const status = $("#simStatus");
    if (status) status.textContent = "Running simulation…";

    // read sim settings from controls (if present)
    if ($("#simN")) model.sim.n = Math.max(100, +$("#simN").value || 100);
    if ($("#targetBCR")) model.sim.targetBCR = +$("#targetBCR").value || 2;
    if ($("#bcrMode")) model.sim.bcrMode = $("#bcrMode").value || "all";
    if ($("#randSeed")) model.sim.seed = $("#randSeed").value ? +$("#randSeed").value : null;

    if ($("#simVarPct")) model.sim.variationPct = Math.max(0, +$("#simVarPct").value || 0);
    if ($("#simVaryOutputs")) model.sim.varyOutputs = $("#simVaryOutputs").value === "true";
    if ($("#simVaryTreatCosts")) model.sim.varyTreatCosts = $("#simVaryTreatCosts").value === "true";
    if ($("#simVaryInputCosts")) model.sim.varyInputCosts = $("#simVaryInputCosts").value === "true";

    const N = model.sim.n;
    const seed = model.sim.seed;
    const u = rng(seed);

    const baseDisc = model.time.discBase;
    const discLow = model.time.discLow;
    const discHigh = model.time.discHigh;

    const adoptBase = model.adoption.base;
    const adoptLow = model.adoption.low;
    const adoptHigh = model.adoption.high;

    const riskBase = model.risk.base;
    const riskLow = model.risk.low;
    const riskHigh = model.risk.high;

    const pct = (model.sim.variationPct || 0) / 100;

    // capture base values for perturbations
    const baseOutputs = model.outputs.map(o => ({ id: o.id, value: Number(o.value) || 0 }));
    const baseTreatCosts = model.treatments.map(t => ({
      id: t.id,
      labourCost: Number(t.labourCost) || 0,
      materialsCost: Number(t.materialsCost) || 0,
      servicesCost: Number(t.servicesCost) || 0,
      capitalCost: Number(t.capitalCost) || 0
    }));
    const baseOtherCosts = model.otherCosts.map(c => ({
      id: c.id,
      annual: Number(c.annual) || 0,
      capital: Number(c.capital) || 0
    }));

    const npvArr = [];
    const bcrArr = [];

    const target = model.sim.targetBCR;

    for (let i = 0; i < N; i++) {
      const disc = triangular(u(), discLow, baseDisc, discHigh);
      const adoptMul = triangular(u(), adoptLow, adoptBase, adoptHigh);
      const risk = triangular(u(), riskLow, riskBase, riskHigh);

      // perturb selected inputs
      if (model.sim.varyOutputs) {
        baseOutputs.forEach(b => {
          const o = model.outputs.find(x => x.id === b.id);
          if (!o) return;
          const mult = 1 + (u() * 2 - 1) * pct;
          o.value = b.value * mult;
        });
      }

      if (model.sim.varyTreatCosts) {
        baseTreatCosts.forEach(b => {
          const t = model.treatments.find(x => x.id === b.id);
          if (!t) return;
          const mult = 1 + (u() * 2 - 1) * pct;
          t.labourCost = b.labourCost * mult;
          t.materialsCost = b.materialsCost * mult;
          t.servicesCost = b.servicesCost * mult;
          t.capitalCost = b.capitalCost * mult;
        });
      }

      if (model.sim.varyInputCosts) {
        baseOtherCosts.forEach(b => {
          const c = model.otherCosts.find(x => x.id === b.id);
          if (!c) return;
          const mult = 1 + (u() * 2 - 1) * pct;
          c.annual = b.annual * mult;
          c.capital = b.capital * mult;
        });
      }

      // compute project metrics
      const flows = buildProjectCashflows({ adoptMul, risk });
      const metrics = computeMetricsFromCashflows(flows, disc, model.sim.bcrMode);

      npvArr.push(metrics.npv);
      bcrArr.push(metrics.bcr);

      // restore base values quickly by writing back at end of each run
      if (model.sim.varyOutputs) baseOutputs.forEach(b => { const o = model.outputs.find(x => x.id === b.id); if (o) o.value = b.value; });
      if (model.sim.varyTreatCosts) baseTreatCosts.forEach(b => { const t = model.treatments.find(x => x.id === b.id); if (t) { t.labourCost = b.labourCost; t.materialsCost = b.materialsCost; t.servicesCost = b.servicesCost; t.capitalCost = b.capitalCost; } });
      if (model.sim.varyInputCosts) baseOtherCosts.forEach(b => { const c = model.otherCosts.find(x => x.id === b.id); if (c) { c.annual = b.annual; c.capital = b.capital; } });
    }

    model.sim.results = { npv: npvArr, bcr: bcrArr };

    // summary stats
    const npvClean = npvArr.filter(v => isFinite(v));
    const bcrClean = bcrArr.filter(v => isFinite(v));

    const stat = arr => {
      const a = [...arr].sort((x, y) => x - y);
      const n = a.length;
      const mean = a.reduce((s, v) => s + v, 0) / (n || 1);
      const median = n ? a[Math.floor(n / 2)] : NaN;
      return {
        min: n ? a[0] : NaN,
        max: n ? a[n - 1] : NaN,
        mean,
        median
      };
    };

    const sNpv = stat(npvClean);
    const sBcr = stat(bcrClean);

    setText("#simNpvMin", money(sNpv.min));
    setText("#simNpvMax", money(sNpv.max));
    setText("#simNpvMean", money(sNpv.mean));
    setText("#simNpvMedian", money(sNpv.median));
    setText("#simNpvProb", npvClean.length ? percent((npvClean.filter(v => v > 0).length / npvClean.length) * 100) : "n/a");

    setText("#simBcrMin", isFinite(sBcr.min) ? fmtNumber(sBcr.min, 3) : "n/a");
    setText("#simBcrMax", isFinite(sBcr.max) ? fmtNumber(sBcr.max, 3) : "n/a");
    setText("#simBcrMean", isFinite(sBcr.mean) ? fmtNumber(sBcr.mean, 3) : "n/a");
    setText("#simBcrMedian", isFinite(sBcr.median) ? fmtNumber(sBcr.median, 3) : "n/a");
    setText("#simBcrProb1", bcrClean.length ? percent((bcrClean.filter(v => v > 1).length / bcrClean.length) * 100) : "n/a");
    setText("#simBcrProbTarget", bcrClean.length ? percent((bcrClean.filter(v => v > target).length / bcrClean.length) * 100) : "n/a");

    const histNpv = $("#histNpv");
    const histBcr = $("#histBcr");
    if (histNpv) drawHistogram(histNpv, npvClean, "NPV distribution");
    if (histBcr) drawHistogram(histBcr, bcrClean, "BCR distribution");

    if (status) status.textContent = `Simulation complete (${N.toLocaleString()} runs).`;
    showToast("Simulation complete.");
  }

  // =========================
  // 14) Exports: CSV / XLSX / PDF / Copy
  // =========================
  function downloadFile(filename, text, mime) {
    const blob = new Blob([text], { type: mime || "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(url);
      a.remove();
    }, 0);
  }

  function exportPdf() {
    // Simple: print current page. CSS should control print layout.
    window.print();
  }

  function buildResultsMatrixForExport(perTreatment) {
    if (!perTreatment || !perTreatment.table) return null;
    const cols = perTreatment.table.columns;
    const rows = perTreatment.table.rows;

    const header = ["Indicator", ...cols.map(c => c.name)];
    const matrix = [header];

    rows.forEach(r => {
      const row = [r.label];
      r.values.forEach(v => {
        let out = v;
        if (r.key === "pvBenefits" || r.key === "pvCosts" || r.key === "npv" || r.key === "__npvDiff__") out = isFinite(v) ? Number(v) : "";
        else if (r.key === "roi" || r.key === "irrVal" || r.key === "mirrVal") out = isFinite(v) ? Number(v) : "";
        else if (r.key === "bcr" || r.key === "__bcrDiff__") out = isFinite(v) ? Number(v) : "";
        else if (r.key === "__rank__") out = v ?? "";
        else if (r.key === "paybackYears") out = v ?? "";
        row.push(out);
      });
      matrix.push(row);
    });

    return matrix;
  }

  function exportResultsCsv() {
    ensureResultsScaffold();
    if (!resultsCache.perTreatment || !resultsCache.perTreatment.table) {
      showToast("No results available to export.");
      return;
    }

    const matrix = buildResultsMatrixForExport(resultsCache.perTreatment);
    if (!matrix) return;

    const csv = matrix
      .map(row =>
        row
          .map(v => {
            const s = v === null || v === undefined ? "" : String(v);
            const safe = s.includes(",") || s.includes('"') || s.includes("\n") ? `"${s.replace(/"/g, '""')}"` : s;
            return safe;
          })
          .join(",")
      )
      .join("\n");

    const filename = `${slug(model.project.name)}_${slug(TOOL_SHORT)}_results.csv`;
    downloadFile(filename, csv, "text/csv");
    showToast("CSV exported.");
  }

  function exportResultsXlsx() {
    if (typeof window.XLSX === "undefined") {
      alert("XLSX library not found. The page includes SheetJS via CDN, but it may be blocked in your environment.");
      return;
    }
    ensureResultsScaffold();
    if (!resultsCache.perTreatment || !resultsCache.perTreatment.table) {
      showToast("No results available to export.");
      return;
    }

    const wb = window.XLSX.utils.book_new();

    // Sheet 1: Results table
    const matrix = buildResultsMatrixForExport(resultsCache.perTreatment);
    const ws1 = window.XLSX.utils.aoa_to_sheet(matrix);
    window.XLSX.utils.book_append_sheet(wb, ws1, "Results");

    // Sheet 2: Inputs summary (lightweight)
    const inputs = [
      ["Tool", TOOL_NAME],
      ["Project name", model.project.name],
      ["Organisation", model.project.organisation],
      ["Last updated", model.project.lastUpdated],
      ["Analysis start year", model.time.startYear],
      ["Years of analysis", model.time.years],
      ["Discount rate (base, %)", model.time.discBase],
      ["Adoption (base)", model.adoption.base],
      ["Risk (base)", model.risk.base],
      ["Include shared allocation", resultsCache.includeShared ? "Yes" : "No"],
      ["BCR denominator", model.sim.bcrMode || "all"]
    ];
    const ws2 = window.XLSX.utils.aoa_to_sheet(inputs);
    window.XLSX.utils.book_append_sheet(wb, ws2, "Scenario");

    // Sheet 3: Outputs
    const outRows = [["Output name", "Unit", "Value ($/unit)", "Source", "Output ID"]];
    model.outputs.forEach(o => outRows.push([o.name, o.unit, Number(o.value) || 0, o.source || "", o.id]));
    const ws3 = window.XLSX.utils.aoa_to_sheet(outRows);
    window.XLSX.utils.book_append_sheet(wb, ws3, "Outputs");

    // Sheet 4: Treatments + deltas
    const tHeader = ["Treatment name", "Control?", "Area (ha)", "Materials ($/ha)", "Services ($/ha)", "Labour ($/ha)", "Capital ($ y0)", "Constrained?", "Source", "Notes", "Treatment ID"];
    model.outputs.forEach(o => tHeader.push(`Delta: ${o.name} (${o.unit}) [per ha]`));

    const tRows = [tHeader];
    model.treatments.forEach(t => {
      const row = [
        t.name,
        t.isControl ? "Yes" : "No",
        Number(t.area) || 0,
        Number(t.materialsCost) || 0,
        Number(t.servicesCost) || 0,
        Number(t.labourCost) || 0,
        Number(t.capitalCost) || 0,
        t.constrained ? "Yes" : "No",
        t.source || "",
        t.notes || "",
        t.id
      ];
      model.outputs.forEach(o => row.push(Number(t.deltas[o.id] ?? 0)));
      tRows.push(row);
    });
    const ws4 = window.XLSX.utils.aoa_to_sheet(tRows);
    window.XLSX.utils.book_append_sheet(wb, ws4, "Treatments");

    const filename = `${slug(model.project.name)}_${slug(TOOL_SHORT)}_results.xlsx`;
    window.XLSX.writeFile(wb, filename);
    showToast("Excel exported.");
  }

  async function copyResultsTableToClipboard() {
    const table = $("#v2ResultsTable");
    if (!table) {
      showToast("No table to copy.");
      return;
    }

    // Copy as TSV for Excel/Word friendly paste
    const rows = Array.from(table.querySelectorAll("tr")).map(tr =>
      Array.from(tr.children)
        .map(td => td.textContent.replace(/\s+/g, " ").trim())
        .join("\t")
    );
    const tsv = rows.join("\n");

    try {
      await navigator.clipboard.writeText(tsv);
      showToast("Results table copied (tab-separated). Paste into Excel or Word.");
    } catch (err) {
      // fallback
      const ta = document.createElement("textarea");
      ta.value = tsv;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      ta.remove();
      showToast("Results table copied.");
    }
  }

  // =========================
  // 15) Excel-first Workflow (Template + Parse + Commit)
  // =========================
  let parsedExcel = null;

  function downloadExcelTemplate({ mode = "blank" } = {}) {
    if (typeof window.XLSX === "undefined") {
      alert("XLSX library not found. The page includes SheetJS via CDN, but it may be blocked in your environment.");
      return;
    }

    const wb = window.XLSX.utils.book_new();

    const metaRows = [
      ["Tool", TOOL_NAME],
      ["Project name", mode === "current" ? model.project.name : ""],
      ["Organisation", mode === "current" ? model.project.organisation : ""],
      ["Last updated", mode === "current" ? model.project.lastUpdated : ""],
      ["Analysis start year", mode === "current" ? model.time.startYear : ""],
      ["Years of analysis", mode === "current" ? model.time.years : ""],
      ["Discount rate base (%)", mode === "current" ? model.time.discBase : ""],
      ["Adoption base (0–1)", mode === "current" ? model.adoption.base : ""],
      ["Risk base (0–1)", mode === "current" ? model.risk.base : ""],
      ["Notes", "Edit the other sheets. Keep column headers unchanged. IDs help match rows on import."]
    ];
    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(metaRows), "Metadata");

    const outputsRows = [["Output name", "Unit", "Value ($/unit)", "Source", "Output ID"]];
    if (mode === "current") {
      model.outputs.forEach(o => outputsRows.push([o.name, o.unit, Number(o.value) || 0, o.source || "", o.id]));
    } else {
      outputsRows.push(["Grain yield", "t/ha", 0, "Input Directly", ""]);
    }
    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(outputsRows), "Outputs");

    const tHeader = ["Treatment name", "Control? (Yes/No)", "Area (ha)", "Materials ($/ha)", "Services ($/ha)", "Labour ($/ha)", "Capital ($ y0)", "Constrained? (Yes/No)", "Source", "Notes", "Treatment ID"];
    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet([tHeader]), "Treatments");

    // Deltas sheet: long format (preferred for robustness)
    const dHeader = ["Treatment ID", "Treatment name", "Output ID", "Output name", "Delta (per ha)"];
    const dRows = [dHeader];
    if (mode === "current") {
      model.treatments.forEach(t => {
        model.outputs.forEach(o => {
          dRows.push([t.id, t.name, o.id, o.name, Number(t.deltas[o.id] ?? 0)]);
        });
      });
    }
    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(dRows), "TreatmentDeltas");

    const bHeader = [
      "Benefit label",
      "Category (C1–C8)",
      "Theme",
      "Frequency (Annual/Once)",
      "Start year",
      "End year",
      "Once year",
      "Unit value ($)",
      "Quantity",
      "Abatement",
      "Annual amount ($)",
      "Growth (%/yr)",
      "Link adoption? (Yes/No)",
      "Link risk? (Yes/No)",
      "P0",
      "P1",
      "Consequence ($)",
      "Notes",
      "Benefit ID"
    ];
    const bRows = [bHeader];
    if (mode === "current") {
      model.benefits.forEach(b => {
        bRows.push([
          b.label,
          b.category,
          b.theme,
          b.frequency,
          b.startYear,
          b.endYear,
          b.year,
          Number(b.unitValue) || 0,
          Number(b.quantity) || 0,
          Number(b.abatement) || 0,
          Number(b.annualAmount) || 0,
          Number(b.growthPct) || 0,
          b.linkAdoption ? "Yes" : "No",
          b.linkRisk ? "Yes" : "No",
          Number(b.p0) || 0,
          Number(b.p1) || 0,
          Number(b.consequence) || 0,
          b.notes || "",
          b.id
        ]);
      });
    }
    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(bRows), "Benefits");

    const cHeader = [
      "Cost label",
      "Type (annual/capital)",
      "Category",
      "Annual ($/yr)",
      "Start year",
      "End year",
      "Capital ($)",
      "Capital year",
      "Constrained? (Yes/No)",
      "Dep method (none/straight/declining)",
      "Dep life (years)",
      "Dep rate (%/yr)",
      "Cost ID"
    ];
    const cRows = [cHeader];
    if (mode === "current") {
      model.otherCosts.forEach(c => {
        cRows.push([
          c.label,
          c.type,
          c.category,
          Number(c.annual) || 0,
          Number(c.startYear) || model.time.startYear,
          Number(c.endYear) || model.time.startYear,
          Number(c.capital) || 0,
          Number(c.year) || model.time.startYear,
          c.constrained ? "Yes" : "No",
          c.depMethod || "none",
          Number(c.depLife) || 5,
          Number(c.depRate) || 30,
          c.id
        ]);
      });
    }
    window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(cRows), "OtherCosts");

    const filename = `${slug(model.project.name)}_${slug(TOOL_SHORT)}_${mode === "blank" ? "blank_template" : "scenario_template"}.xlsx`;
    window.XLSX.writeFile(wb, filename);
    showToast("Excel template downloaded.");
  }

  function handleParseExcel() {
    if (typeof window.XLSX === "undefined") {
      alert("XLSX library not found. The page includes SheetJS via CDN, but it may be blocked in your environment.");
      return;
    }

    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".xlsx,.xls";
    input.style.display = "none";
    document.body.appendChild(input);

    input.addEventListener("change", async () => {
      const file = input.files && input.files[0];
      input.remove();
      if (!file) return;

      try {
        const data = await file.arrayBuffer();
        const wb = window.XLSX.read(data, { type: "array" });

        const sheets = wb.SheetNames || [];
        const required = ["Metadata", "Outputs", "Treatments", "TreatmentDeltas"];
        const missing = required.filter(s => !sheets.includes(s));
        if (missing.length) {
          alert(`Missing required sheet(s): ${missing.join(", ")}. Use the tool template.`);
          return;
        }

        const toJson = name => window.XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: "" });

        const meta = toJson("Metadata");
        const outputs = toJson("Outputs");
        const treatments = toJson("Treatments");
        const deltas = toJson("TreatmentDeltas");
        const benefits = sheets.includes("Benefits") ? toJson("Benefits") : [];
        const otherCosts = sheets.includes("OtherCosts") ? toJson("OtherCosts") : [];

        // Lightweight validation
        const outOk = outputs.some(r => String(r["Output name"] || "").trim());
        const trOk = treatments.some(r => String(r["Treatment name"] || "").trim());
        if (!outOk || !trOk) {
          alert("Outputs and Treatments sheets must contain at least one row with a name.");
          return;
        }

        parsedExcel = { meta, outputs, treatments, deltas, benefits, otherCosts, wbInfo: { name: file.name, sheets } };
        showToast(`Excel parsed: ${file.name}. Review then click "Apply parsed Excel data to the model".`);
      } catch (err) {
        console.error(err);
        alert("Could not parse the Excel workbook.");
      }
    });

    input.click();
  }

  function commitExcelToModel() {
    if (!parsedExcel) {
      alert("No parsed Excel data found. Click 'Parse Excel file' first.");
      return;
    }

    // Build outputs (prefer IDs if provided, else create)
    const outputs = [];
    parsedExcel.outputs.forEach(r => {
      const name = String(r["Output name"] || "").trim();
      if (!name) return;
      const unit = String(r["Unit"] || "").trim();
      const value = parseNumber(r["Value ($/unit)"]);
      const source = String(r["Source"] || "Input Directly").trim() || "Input Directly";
      const id = String(r["Output ID"] || "").trim() || uid();
      outputs.push({ id, name, unit, value: Number.isFinite(value) ? value : 0, source });
    });

    if (!outputs.length) {
      alert("No valid outputs found in Outputs sheet.");
      return;
    }

    // Treatments (prefer IDs)
    const treatments = [];
    parsedExcel.treatments.forEach(r => {
      const name = String(r["Treatment name"] || "").trim();
      if (!name) return;

      const id = String(r["Treatment ID"] || "").trim() || uid();
      const isControl = String(r["Control? (Yes/No)"] || "").trim().toLowerCase() === "yes";
      const area = parseNumber(r["Area (ha)"]);
      const materialsCost = parseNumber(r["Materials ($/ha)"]);
      const servicesCost = parseNumber(r["Services ($/ha)"]);
      const labourCost = parseNumber(r["Labour ($/ha)"]);
      const capitalCost = parseNumber(r["Capital ($ y0)"]);
      const constrained = String(r["Constrained? (Yes/No)"] || "").trim().toLowerCase() !== "no";
      const source = String(r["Source"] || "Input Directly").trim() || "Input Directly";
      const notes = String(r["Notes"] || "").trim();

      treatments.push({
        id,
        name,
        area: Number.isFinite(area) ? area : 0,
        adoption: 1,
        deltas: {},
        labourCost: Number.isFinite(labourCost) ? labourCost : 0,
        materialsCost: Number.isFinite(materialsCost) ? materialsCost : 0,
        servicesCost: Number.isFinite(servicesCost) ? servicesCost : 0,
        capitalCost: Number.isFinite(capitalCost) ? capitalCost : 0,
        constrained,
        source,
        isControl,
        notes
      });
    });

    if (!treatments.length) {
      alert("No valid treatments found in Treatments sheet.");
      return;
    }

    // Ensure exactly one control (if none, set first)
    const controls = treatments.filter(t => t.isControl);
    if (controls.length === 0) treatments[0].isControl = true;
    if (controls.length > 1) {
      // keep first control only
      let seen = false;
      treatments.forEach(t => {
        if (t.isControl) {
          if (!seen) seen = true;
          else t.isControl = false;
        }
      });
    }

    // Map deltas
    // Prefer mapping by Treatment ID + Output ID, else by names if IDs missing.
    const tById = new Map(treatments.map(t => [t.id, t]));
    const oById = new Map(outputs.map(o => [o.id, o]));
    const tByName = new Map(treatments.map(t => [t.name.toLowerCase(), t]));
    const oByName = new Map(outputs.map(o => [o.name.toLowerCase(), o]));

    // initialise all deltas to 0
    treatments.forEach(t => { outputs.forEach(o => { t.deltas[o.id] = 0; }); });

    parsedExcel.deltas.forEach(r => {
      const tId = String(r["Treatment ID"] || "").trim();
      const tName = String(r["Treatment name"] || "").trim().toLowerCase();
      const oId = String(r["Output ID"] || "").trim();
      const oName = String(r["Output name"] || "").trim().toLowerCase();
      const d = parseNumber(r["Delta (per ha)"]);

      const t = tId ? tById.get(tId) : tByName.get(tName);
      const o = oId ? oById.get(oId) : oByName.get(oName);
      if (!t || !o) return;
      if (!Number.isFinite(d)) return;
      t.deltas[o.id] = d;
    });

    // Benefits (optional)
    const benefits = [];
    if (parsedExcel.benefits && parsedExcel.benefits.length) {
      parsedExcel.benefits.forEach(r => {
        const label = String(r["Benefit label"] || "").trim();
        if (!label) return;
        benefits.push({
          id: String(r["Benefit ID"] || "").trim() || uid(),
          label,
          category: String(r["Category (C1–C8)"] || "C4").trim() || "C4",
          theme: String(r["Theme"] || "Other").trim() || "Other",
          frequency: String(r["Frequency (Annual/Once)"] || "Annual").trim() || "Annual",
          startYear: parseNumber(r["Start year"]) || model.time.startYear,
          endYear: parseNumber(r["End year"]) || model.time.startYear,
          year: parseNumber(r["Once year"]) || model.time.startYear,
          unitValue: parseNumber(r["Unit value ($)"]) || 0,
          quantity: parseNumber(r["Quantity"]) || 0,
          abatement: parseNumber(r["Abatement"]) || 0,
          annualAmount: parseNumber(r["Annual amount ($)"]) || 0,
          growthPct: parseNumber(r["Growth (%/yr)"]) || 0,
          linkAdoption: String(r["Link adoption? (Yes/No)"] || "Yes").trim().toLowerCase() === "yes",
          linkRisk: String(r["Link risk? (Yes/No)"] || "Yes").trim().toLowerCase() === "yes",
          p0: parseNumber(r["P0"]) || 0,
          p1: parseNumber(r["P1"]) || 0,
          consequence: parseNumber(r["Consequence ($)"]) || 0,
          notes: String(r["Notes"] || "").trim()
        });
      });
    }

    // Other costs (optional)
    const otherCosts = [];
    if (parsedExcel.otherCosts && parsedExcel.otherCosts.length) {
      parsedExcel.otherCosts.forEach(r => {
        const label = String(r["Cost label"] || "").trim();
        if (!label) return;
        otherCosts.push({
          id: String(r["Cost ID"] || "").trim() || uid(),
          label,
          type: String(r["Type (annual/capital)"] || "annual").trim() || "annual",
          category: String(r["Category"] || "Services").trim() || "Services",
          annual: parseNumber(r["Annual ($/yr)"]) || 0,
          startYear: parseNumber(r["Start year"]) || model.time.startYear,
          endYear: parseNumber(r["End year"]) || model.time.startYear,
          capital: parseNumber(r["Capital ($)"]) || 0,
          year: parseNumber(r["Capital year"]) || model.time.startYear,
          constrained: String(r["Constrained? (Yes/No)"] || "Yes").trim().toLowerCase() === "yes",
          depMethod: String(r["Dep method (none/straight/declining)"] || "none").trim() || "none",
          depLife: parseNumber(r["Dep life (years)"]) || 5,
          depRate: parseNumber(r["Dep rate (%/yr)"]) || 30
        });
      });
    }

    // Metadata (optional: update some headline fields)
    const meta = parsedExcel.meta || [];
    const pickMeta = key => {
      const row = meta.find(r => String(r["Tool"] || "").trim() === key || String(r[0] || "").trim() === key);
      return row ? (row[1] || row["Project name"] || "") : "";
    };
    // Avoid fragile meta parsing; keep current project unless metadata has values.
    // (User can edit project tab directly.)

    // Commit to model
    model.outputs = outputs;
    model.treatments = treatments;
    model.benefits = benefits.length ? benefits : model.benefits;
    model.otherCosts = otherCosts.length ? otherCosts : model.otherCosts;

    initTreatmentDeltas();
    renderAll();
    calcAndRender();
    parsedExcel = null;

    showToast("Excel data applied to the model.");
  }

  // =========================
  // 16) AI-assisted interpretation prompt
  // =========================
  function buildAIPrompt() {
    ensureResultsScaffold();
    if (!resultsCache.perTreatment || !resultsCache.perTreatment.table) {
      return `${TOOL_NAME}\n\nNo results are currently available. Please calculate results first.`;
    }

    const includeShared = resultsCache.includeShared;
    const rate = model.time.discBase;
    const adopt = model.adoption.base;
    const risk = model.risk.base;
    const bcrMode = model.sim.bcrMode || "all";

    const controlName = resultsCache.perTreatment.control?.t?.name || "Control";

    // Extract table as TSV (LLM-friendly)
    const tableEl = $("#v2ResultsTable");
    const tsv = tableEl
      ? Array.from(tableEl.querySelectorAll("tr"))
          .map(tr => Array.from(tr.children).map(td => td.textContent.replace(/\s+/g, " ").trim()).join("\t"))
          .join("\n")
      : "";

    // Identify low performers for tailored improvement suggestions
    const controlId = resultsCache.perTreatment.control?.t?.id;
    const nonControl = resultsCache.perTreatment.tMetrics.filter(x => x.t.id !== controlId);

    const lowBcr = nonControl
      .filter(x => isFinite(x.m.bcr) && x.m.bcr < 1)
      .sort((a, b) => (a.m.bcr || Infinity) - (b.m.bcr || Infinity))
      .slice(0, 5);

    const negNpv = nonControl
      .filter(x => isFinite(x.m.npv) && x.m.npv < 0)
      .sort((a, b) => (a.m.npv || Infinity) - (b.m.npv || Infinity))
      .slice(0, 5);

    const prompt = `
You are assisting with interpretation of an agricultural cost–benefit analysis produced by ${TOOL_NAME}. 
Write in plain English suitable for farmers and extension officers. Do not prescribe a single “correct” decision. 
Explain what the indicators mean and why treatments perform differently. Highlight trade-offs and what drives results.

CONTEXT
Project: ${model.project.name}
Organisation: ${model.project.organisation}
Analysis start year: ${model.time.startYear}
Years of analysis: ${model.time.years}
Discount rate (base): ${rate}%
Adoption (base): ${adopt} (0–1)
Risk (base): ${risk} (0–1)
BCR denominator: ${bcrMode}
Comparison mode: ${includeShared ? "Shared project-wide items allocated across treatments (advanced)" : "Treatment-only (outputs + treatment costs); shared items kept in project totals"}

CONTROL
The flagged control treatment is: ${controlName}.
Interpretation should compare each treatment to this control, and note that the table also includes differences vs control.

RESULTS TABLE (tab-separated; rows are indicators, columns are treatments)
${tsv}

YOUR OUTPUT MUST FOLLOW THIS STRUCTURE
1) One-paragraph overview that summarises which treatments look stronger and weaker economically under the current assumptions, without telling the user what to choose.
2) A short explanation of each key indicator: PV benefits, PV costs, NPV, BCR, ROI, and payback. Keep each explanation to one or two sentences.
3) A treatment comparison section that:
   - Identifies the top two treatments (by BCR and by NPV if different) and briefly explains what is driving their performance (benefits, costs, or both).
   - Identifies the weakest two treatments (especially those with negative NPV or low BCR) and explains what drives underperformance (high costs, small yield response, risk, adoption, or discounting).
4) Learning and improvement suggestions (non-prescriptive). For any treatment with low BCR or negative NPV, suggest realistic ways it might be improved in practice, using language like “could consider” or “might explore”, such as:
   - reducing input, labour, or service costs
   - targeting the practice to the right paddocks or soil constraints
   - improving implementation quality to raise yield response
   - adjusting timing or application rate
   - improving marketing/pricing assumptions where justified
   - reducing risk through monitoring, trials, or staged adoption
   Make clear these are levers to explore, not rules or thresholds imposed by the tool.
5) A short “what to check next” section that suggests sensitivity checks the user can run inside the tool (discount rate, adoption, risk, output prices, and treatment costs), and how those parameters might matter.

ADDITIONAL FLAGS FOR YOU TO PAY ATTENTION TO
Treatments with BCR < 1 (benefits lower than costs under current assumptions): ${lowBcr.length ? lowBcr.map(x => x.t.name).join(", ") : "None flagged"}
Treatments with NPV < 0 under current assumptions: ${negNpv.length ? negNpv.map(x => x.t.name).join(", ") : "None flagged"}

STYLE RULES
No bullet lists unless absolutely necessary. Prefer short paragraphs. Use cautious, non-prescriptive wording.
`.trim();

    return prompt;
  }

  async function handleCopyAIPrompt() {
    const prompt = buildAIPrompt();
    const preview = $("#copilotPreview");
    if (preview) preview.value = prompt;

    try {
      await navigator.clipboard.writeText(prompt);
      showToast("AI interpretation prompt copied. Paste into Copilot or ChatGPT.");
    } catch (err) {
      const ta = document.createElement("textarea");
      ta.value = prompt;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      ta.remove();
      showToast("AI interpretation prompt copied.");
    }
  }

  // =========================
  // 17) Main Calculation + Render
  // =========================
  function calcAndRender() {
    ensureResultsScaffold();

    // Basic sanity: ensure one control
    if (!model.treatments.some(t => t.isControl) && model.treatments.length) model.treatments[0].isControl = true;
    // Ensure deltas exist
    initTreatmentDeltas();

    const rate = model.time.discBase;
    const bcrMode = model.sim.bcrMode || "all";

    // Whole project
    const projectFlows = buildProjectCashflows({ adoptMul: model.adoption.base, risk: model.risk.base });
    const projectMetrics = computeMetricsFromCashflows(projectFlows, rate, bcrMode);
    resultsCache.project = { flows: projectFlows, metrics: projectMetrics };

    renderWholeProjectSummary(projectMetrics);
    renderTimeProjection(projectFlows);

    // Per treatment
    const perTreatment = computePerTreatmentResults({ includeShared: resultsCache.includeShared });
    resultsCache.perTreatment = perTreatment;

    renderTreatmentComparisonTable(perTreatment);
    renderHeadline(perTreatment);

    // Update copilot preview opportunistically
    const preview = $("#copilotPreview");
    if (preview) preview.value = buildAIPrompt();

    // Keep legacy sections reasonably consistent: control/treatment group cards and ranking list
    renderLegacyControlAndGroup(perTreatment);
    renderLegacyRankingList(perTreatment);

    // Depreciation summary (informational)
    renderDepreciationSummary();
  }

  function renderLegacyControlAndGroup(perTreatment) {
    // This keeps the existing “Control” and “Combined treatment group” panels populated,
    // but the primary comparison is now the vertical table.

    const controlEntry = perTreatment?.control;
    if (controlEntry) {
      setText("#pvBenefitsControl", money(controlEntry.m.pvBenefits));
      setText("#pvCostsControl", money(controlEntry.m.pvCosts));
      setText("#npvControl", money(controlEntry.m.npv));
      setText("#bcrControl", isFinite(controlEntry.m.bcr) ? fmtNumber(controlEntry.m.bcr, 3) : "n/a");
      setText("#irrControl", percent(controlEntry.m.irrVal));
      setText("#roiControl", percent(controlEntry.m.roi));
      setText("#paybackControl", controlEntry.m.paybackYears === null ? "n/a" : String(controlEntry.m.paybackYears));
      setText("#gmControl", money(controlEntry.m.annualGM));
    }

    if (perTreatment?.tMetrics && perTreatment.tMetrics.length) {
      const controlId = controlEntry?.t?.id;
      const nonControl = perTreatment.tMetrics.filter(x => x.t.id !== controlId);

      // Combine non-control treatments by summing cashflows (treatment-only or shared-allocated depending on mode)
      if (nonControl.length) {
        const N = model.time.years;
        const b = new Array(N + 1).fill(0);
        const c = new Array(N + 1).fill(0);
        const cc = new Array(N + 1).fill(0);
        nonControl.forEach(x => {
          x.flows.benefitByYear.forEach((v, i) => (b[i] += v));
          x.flows.costByYear.forEach((v, i) => (c[i] += v));
          x.flows.constrainedCostByYear.forEach((v, i) => (cc[i] += v));
        });
        const cf = b.map((v, i) => v - c[i]);
        const flows = { benefitByYear: b, costByYear: c, constrainedCostByYear: cc, cf, annualGM: NaN };
        const m = computeMetricsFromCashflows(flows, model.time.discBase, model.sim.bcrMode || "all");

        setText("#pvBenefitsTreat", money(m.pvBenefits));
        setText("#pvCostsTreat", money(m.pvCosts));
        setText("#npvTreat", money(m.npv));
        setText("#bcrTreat", isFinite(m.bcr) ? fmtNumber(m.bcr, 3) : "n/a");
        setText("#irrTreat", percent(m.irrVal));
        setText("#roiTreat", percent(m.roi));
        setText("#paybackTreat", m.paybackYears === null ? "n/a" : String(m.paybackYears));
        setText("#gmTreat", isFinite(m.annualGM) ? money(m.annualGM) : "n/a");
      }
    }
  }

  function renderLegacyRankingList(perTreatment) {
    const root = $("#treatmentSummary");
    if (!root) return;

    root.innerHTML = "";
    if (!perTreatment || !perTreatment.tMetrics || perTreatment.tMetrics.length === 0) {
      root.innerHTML = `<div class="small muted">No ranking available.</div>`;
      return;
    }

    const controlId = perTreatment.control?.t?.id;

    const ranked = [...perTreatment.tMetrics].sort((a, b) => {
      const bcrA = isFinite(a.m.bcr) ? a.m.bcr : -Infinity;
      const bcrB = isFinite(b.m.bcr) ? b.m.bcr : -Infinity;
      if (bcrB !== bcrA) return bcrB - bcrA;
      const npvA = isFinite(a.m.npv) ? a.m.npv : -Infinity;
      const npvB = isFinite(b.m.npv) ? b.m.npv : -Infinity;
      return npvB - npvA;
    });

    ranked.forEach((x, i) => {
      const div = document.createElement("div");
      div.className = "item";
      const isCtrl = x.t.id === controlId;
      const diffNpv = isCtrl ? 0 : x.m.npv - perTreatment.control.m.npv;

      div.innerHTML = `
        <h4>${i + 1}. ${esc(x.t.name)} ${isCtrl ? "<span class='badge'>Control</span>" : ""}</h4>
        <div class="row-4">
          <div class="field"><label>BCR</label><div class="metric"><div class="value">${isFinite(x.m.bcr) ? fmtNumber(x.m.bcr, 3) : "n/a"}</div></div></div>
          <div class="field"><label>NPV</label><div class="metric"><div class="value">${money(x.m.npv)}</div></div></div>
          <div class="field"><label>ROI</label><div class="metric"><div class="value">${percent(x.m.roi)}</div></div></div>
          <div class="field"><label>ΔNPV vs control</label><div class="metric"><div class="value">${money(diffNpv)}</div></div></div>
        </div>
      `;
      root.appendChild(div);
    });
  }

  function renderDepreciationSummary() {
    const root = $("#depSummary");
    if (!root) return;

    const items = model.otherCosts.filter(c => c.category === "Capital" && (Number(c.capital) || 0) > 0 && c.depMethod && c.depMethod !== "none");
    if (!items.length) {
      root.innerHTML = `<div class="small muted">No capital items with depreciation settings.</div>`;
      return;
    }

    const rows = items.map(c => {
      const cap = Number(c.capital) || 0;
      const life = Number(c.depLife) || 5;
      const rate = Number(c.depRate) || 30;
      const method = c.depMethod;

      // Simple summary amounts (informational only)
      let year1 = 0;
      if (method === "straight") year1 = cap / life;
      else if (method === "declining") year1 = (rate / 100) * cap;

      return `<tr>
        <td>${esc(c.label)}</td>
        <td>${esc(method)}</td>
        <td>${fmtNumber(life, 0)}</td>
        <td>${percent(rate)}</td>
        <td>${money(cap)}</td>
        <td>${money(year1)}</td>
      </tr>`;
    });

    root.innerHTML = `
      <div class="table-scroll">
        <table class="summary-table">
          <thead>
            <tr>
              <th>Capital item</th>
              <th>Method</th>
              <th>Life (years)</th>
              <th>Declining rate</th>
              <th>Capital value</th>
              <th>Indicative year-1 depreciation</th>
            </tr>
          </thead>
          <tbody>${rows.join("")}</tbody>
        </table>
      </div>
      <div class="small muted">Depreciation is shown for documentation and learning. Present value calculations use the entered costs as cashflows, not accounting depreciation.</div>
    `;
  }

  // =========================
  // 18) Branding consistency (Tool 2 naming)
  // =========================
  function updateBrandingText() {
    // Header title
    const brandTitle = document.querySelector(".brand-title");
    if (brandTitle) brandTitle.textContent = TOOL_NAME;

    // Footer
    const footerLeft = document.querySelector(".app-footer .footer-left .small");
    if (footerLeft) footerLeft.textContent = `${TOOL_SHORT}, ${ORG_NAME}`;

    // Document title
    document.title = `${TOOL_NAME} - ${ORG_NAME}`;
  }

  // =========================
  // 19) Render All + Init
  // =========================
  function renderAll() {
    updateBrandingText();
    renderOutputs();
    renderTreatments();
    renderBenefits();
    renderCosts();
    renderDatabaseTags();
    ensureResultsScaffold();
  }

  function init() {
    updateBrandingText();
    initTabs();
    bindBasics();
    initAddButtons();
    renderAll();
    calcAndRender();
  }

  // Run
  document.addEventListener("DOMContentLoaded", init);
})();
