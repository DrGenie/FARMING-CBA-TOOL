// Farming CBA Tool - Newcastle Business School
// Commercial-grade app.js (upgrade): hardened state management, autosave, robust import/export,
// accessible tabs with deep-linking, faster/safer Monte Carlo, optional discount schedule support,
// and calibrated faba bean trial defaults.

(() => {
  "use strict";

  // =========================
  // VERSIONING + DEFAULTS
  // =========================
  const APP_VERSION = "2026.01.02";
  const MODEL_SCHEMA_VERSION = 2;

  // ---------- CONSTANTS ----------
  const DEFAULT_DISCOUNT_SCHEDULE = [
    { label: "2025-2034", from: 2025, to: 2034, low: 2, base: 4, high: 6 },
    { label: "2035-2044", from: 2035, to: 2044, low: 4, base: 7, high: 10 },
    { label: "2045-2054", from: 2045, to: 2054, low: 4, base: 7, high: 10 },
    { label: "2055-2064", from: 2055, to: 2064, low: 3, base: 6, high: 9 },
    { label: "2065-2074", from: 2065, to: 2074, low: 2, base: 5, high: 8 }
  ];

  const horizons = [5, 10, 15, 20, 25];

  // ---------- ID HELPER ----------
  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

  // ---------- SAFE UTIL ----------
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const isNum = v => typeof v === "number" && Number.isFinite(v);
  const asNum = (v, fallback = 0) => {
    const n = typeof v === "string" ? parseFloat(v) : v;
    return Number.isFinite(n) ? n : fallback;
  };
  const nowISODate = () => new Date().toISOString().slice(0, 10);

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

  function safeJsonParse(text, fallback = null) {
    try {
      return JSON.parse(text);
    } catch {
      return fallback;
    }
  }

  function deepClone(obj) {
    return obj == null ? obj : JSON.parse(JSON.stringify(obj));
  }

  function deepMerge(target, source) {
    if (!source || typeof source !== "object") return target;
    const out = Array.isArray(target) ? target.slice() : Object.assign({}, target);

    Object.keys(source).forEach(k => {
      const sv = source[k];
      const tv = out[k];
      if (Array.isArray(sv)) out[k] = sv.slice();
      else if (sv && typeof sv === "object") out[k] = deepMerge(tv && typeof tv === "object" ? tv : {}, sv);
      else out[k] = sv;
    });

    return out;
  }

  // =========================
  // MODEL (DEFAULT STATE)
  // =========================
  const DEFAULT_MODEL = {
    __meta__: {
      appVersion: APP_VERSION,
      schemaVersion: MODEL_SCHEMA_VERSION,
      created: new Date().toISOString(),
      lastSaved: null
    },
    project: {
      name: "Faba bean soil amendment trial",
      lead: "Project lead",
      analysts: "Farm economics team",
      team: "Trial team",
      organisation: "Newcastle Business School, The University of Newcastle",
      contactEmail: "",
      contactPhone: "",
      summary:
        "Applied faba bean trial comparing deep ripping, organic matter, gypsum and fertiliser treatments against a control.",
      objectives: "Quantify yield and gross margin impacts of alternative soil amendment strategies.",
      activities: "Establish replicated field plots, collect plot-level yield and cost data, and summarise trial-wide economics.",
      stakeholders: "Producers, agronomists, government agencies, research partners.",
      lastUpdated: nowISODate(),
      goal:
        "Identify soil amendment packages that deliver higher faba bean yields and acceptable returns after accounting for additional costs.",
      withProject:
        "Faba bean growers adopt high-performing amendment packages on trial farms and similar soils in the region.",
      withoutProject:
        "Growers continue with baseline practice and do not access detailed economic evidence on soil amendments."
    },
    time: {
      startYear: new Date().getFullYear(),
      projectStartYear: new Date().getFullYear(),
      years: 10,
      discBase: 7,
      discLow: 4,
      discHigh: 10,
      useDiscountSchedule: false, // Optional: if true (or checkbox exists), PV uses schedule rates by year
      mirrFinance: 6,
      mirrReinvest: 4,
      discountSchedule: deepClone(DEFAULT_DISCOUNT_SCHEDULE)
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
        notes: "Baseline faba bean practice without deep soil amendment."
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
        p0: 0.1,
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
      details: [],
      variationPct: 20,
      varyOutputs: true,
      varyTreatCosts: true,
      varyInputCosts: false
    }
  };

  // ---------- CORE MODEL (LIVE) ----------
  const model = deepClone(DEFAULT_MODEL);

  // =========================
  // PERSISTENCE (AUTOSAVE)
  // =========================
  const STORAGE_NS = "nbs_farming_cba";
  const storageKey = () => `${STORAGE_NS}::${slug(model.project?.name || "project")}`;

  function tryAutosave() {
    try {
      const payload = deepClone(model);
      payload.__meta__ = payload.__meta__ || {};
      payload.__meta__.lastSaved = new Date().toISOString();
      localStorage.setItem(storageKey(), JSON.stringify(payload));
      localStorage.setItem(`${STORAGE_NS}::last_key`, storageKey());
      return true;
    } catch (e) {
      console.warn("Autosave failed:", e);
      return false;
    }
  }

  function tryAutoload() {
    try {
      const lastKey = localStorage.getItem(`${STORAGE_NS}::last_key`);
      const key = lastKey || storageKey();
      const raw = localStorage.getItem(key);
      if (!raw) return false;

      const parsed = safeJsonParse(raw, null);
      if (!parsed || typeof parsed !== "object") return false;

      // Merge into defaults, then enforce required invariants.
      const merged = deepMerge(deepClone(DEFAULT_MODEL), parsed);
      Object.keys(model).forEach(k => delete model[k]);
      Object.assign(model, merged);

      hardenModel();
      return true;
    } catch (e) {
      console.warn("Autoload failed:", e);
      return false;
    }
  }

  let autosaveTimer = null;
  function autosaveDebounced() {
    clearTimeout(autosaveTimer);
    autosaveTimer = setTimeout(() => {
      tryAutosave();
    }, 300);
  }

  // =========================
  // RNG + DISTRIBUTIONS
  // =========================
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

  // =========================
  // TOASTS
  // =========================
  function showToast(message) {
    if (!message) return;
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

  function safeRun(label, fn) {
    try {
      return fn();
    } catch (e) {
      console.error(label, e);
      showToast("An error occurred. See console for details.");
      return null;
    }
  }

  // =========================
  // OPTIONAL TRIAL CALIBRATION BLOCKS
  // =========================
  const CAPITAL_ASSETS = {
    deepRipper5Tyne: {
      id: "deepRipper5Tyne",
      label: "Pre sow amendment 5 tyne ripper",
      purchasePriceAud: 125000,
      expectedLifeYears: 10,
      utilisationHaPerYear: 3300,
      notes: "Used for deep organic matter and gypsum style amendments"
    },
    speedTiller10m: {
      id: "speedTiller10m",
      label: "Speed tiller 10 m",
      purchasePriceAud: 259000,
      expectedLifeYears: 10,
      utilisationHaPerYear: 3300,
      notes: "Used for soil preparation passes"
    },
    airSeeder12m: {
      id: "airSeeder12m",
      label: "Air seeder 12 m",
      purchasePriceAud: 162800,
      expectedLifeYears: 10,
      utilisationHaPerYear: 3300,
      notes: "Standard seeding unit"
    },
    boomSpray36m: {
      id: "boomSpray36m",
      label: "36 m boomspray",
      purchasePriceAud: 792000,
      expectedLifeYears: 10,
      utilisationHaPerYear: 3300,
      notes: "Used for herbicide, fungicide, insecticide passes"
    }
  };

  const BASELINE_CROPPING_COSTS = {
    seedAndFertiliserPerHa: 0,
    herbicidePerHa: 0,
    fungicidePerHa: 0,
    insecticidePerHa: 0,
    fuelAndMachineryPerHa: 0,
    labourPerHa: 0
  };

  const TRIAL_TREATMENT_CONFIG = [
    {
      id: "control",
      label: "Control (no amendment)",
      category: "Baseline practice",
      controlFlag: true,
      agronomy: {
        mean_yield_t_ha: null,
        std_yield_t_ha: null,
        plantsPerM2: null,
        notes: "Standard practice with no deep amendments"
      },
      costs: { treatmentInputCostPerHa: 0, labourCostPerHa: 0, additionalOperatingCostPerHa: 0, capitalDepreciationPerHa: 0 },
      capitalAssets: []
    },
    {
      id: "deep_om_cp1",
      label: "Deep organic matter (CP1)",
      category: "Soil amendment",
      controlFlag: false,
      agronomy: { mean_yield_t_ha: null, std_yield_t_ha: null, plantsPerM2: null, notes: "Deep organic matter incorporation at 15 t/ha" },
      costs: { treatmentInputCostPerHa: 16500, labourCostPerHa: null, additionalOperatingCostPerHa: null, capitalDepreciationPerHa: null },
      capitalAssets: ["deepRipper5Tyne"]
    },
    {
      id: "deep_om_cp1_plus_liq_gypsum_cht",
      label: "Deep OM (CP1) + liquid gypsum (CHT)",
      category: "Soil amendment",
      controlFlag: false,
      agronomy: { mean_yield_t_ha: null, std_yield_t_ha: null, plantsPerM2: null, notes: "Combination of deep OM and liquid gypsum" },
      costs: { treatmentInputCostPerHa: 16850, labourCostPerHa: null, additionalOperatingCostPerHa: null, capitalDepreciationPerHa: null },
      capitalAssets: ["deepRipper5Tyne"]
    },
    {
      id: "deep_gypsum",
      label: "Deep gypsum",
      category: "Soil amendment",
      controlFlag: false,
      agronomy: { mean_yield_t_ha: null, std_yield_t_ha: null, plantsPerM2: null, notes: "Deep placement gypsum at 5 t/ha" },
      costs: { treatmentInputCostPerHa: 500, labourCostPerHa: null, additionalOperatingCostPerHa: null, capitalDepreciationPerHa: null },
      capitalAssets: ["deepRipper5Tyne"]
    },
    {
      id: "deep_om_cp1_plus_pam",
      label: "Deep OM (CP1) + PAM",
      category: "Soil amendment",
      controlFlag: false,
      agronomy: { mean_yield_t_ha: null, std_yield_t_ha: null, plantsPerM2: null, notes: "Deep OM with polyacrylamide" },
      costs: { treatmentInputCostPerHa: 18000, labourCostPerHa: null, additionalOperatingCostPerHa: null, capitalDepreciationPerHa: null },
      capitalAssets: ["deepRipper5Tyne"]
    },
    {
      id: "deep_om_cp1_plus_ccm",
      label: "Deep OM (CP1) + carbon coated mineral (CCM)",
      category: "Soil amendment",
      controlFlag: false,
      agronomy: { mean_yield_t_ha: null, std_yield_t_ha: null, plantsPerM2: null, notes: "Deep OM plus CCM blend" },
      costs: { treatmentInputCostPerHa: 21225, labourCostPerHa: null, additionalOperatingCostPerHa: null, capitalDepreciationPerHa: null },
      capitalAssets: ["deepRipper5Tyne"]
    },
    {
      id: "deep_ccm_only",
      label: "Deep carbon coated mineral (CCM) only",
      category: "Soil amendment",
      controlFlag: false,
      agronomy: { mean_yield_t_ha: null, std_yield_t_ha: null, plantsPerM2: null, notes: "CCM at 5 t/ha without deep OM" },
      costs: { treatmentInputCostPerHa: 3225, labourCostPerHa: null, additionalOperatingCostPerHa: null, capitalDepreciationPerHa: null },
      capitalAssets: ["deepRipper5Tyne"]
    },
    {
      id: "deep_om_cp2_plus_gypsum",
      label: "Deep OM + gypsum (CP2)",
      category: "Soil amendment",
      controlFlag: false,
      agronomy: { mean_yield_t_ha: null, std_yield_t_ha: null, plantsPerM2: null, notes: "Alternative deep OM and gypsum mix" },
      costs: { treatmentInputCostPerHa: 24000, labourCostPerHa: null, additionalOperatingCostPerHa: null, capitalDepreciationPerHa: null },
      capitalAssets: ["deepRipper5Tyne"]
    },
    {
      id: "deep_liq_gypsum_cht",
      label: "Deep liquid gypsum (CHT)",
      category: "Soil amendment",
      controlFlag: false,
      agronomy: { mean_yield_t_ha: null, std_yield_t_ha: null, plantsPerM2: null, notes: "Liquid gypsum at 0.5 t/ha" },
      costs: { treatmentInputCostPerHa: 350, labourCostPerHa: null, additionalOperatingCostPerHa: null, capitalDepreciationPerHa: null },
      capitalAssets: ["deepRipper5Tyne"]
    },
    {
      id: "surface_silicon",
      label: "Surface silicon",
      category: "Surface amendment",
      controlFlag: false,
      agronomy: { mean_yield_t_ha: null, std_yield_t_ha: null, plantsPerM2: null, notes: "Surface applied silicon at 2 t/ha" },
      costs: { treatmentInputCostPerHa: 1000, labourCostPerHa: null, additionalOperatingCostPerHa: null, capitalDepreciationPerHa: null },
      capitalAssets: []
    },
    {
      id: "deep_liq_npks",
      label: "Deep liquid NPKS",
      category: "Nutrient injection",
      controlFlag: false,
      agronomy: { mean_yield_t_ha: null, std_yield_t_ha: null, plantsPerM2: null, notes: "Deep injected liquid NPKS at 750 L/ha" },
      costs: { treatmentInputCostPerHa: 2200, labourCostPerHa: null, additionalOperatingCostPerHa: null, capitalDepreciationPerHa: null },
      capitalAssets: ["deepRipper5Tyne"]
    },
    {
      id: "deep_ripping_only",
      label: "Deep ripping only",
      category: "Mechanical only",
      controlFlag: false,
      agronomy: { mean_yield_t_ha: null, std_yield_t_ha: null, plantsPerM2: null, notes: "Deep ripping without added material" },
      costs: { treatmentInputCostPerHa: 0, labourCostPerHa: null, additionalOperatingCostPerHa: null, capitalDepreciationPerHa: null },
      capitalAssets: ["deepRipper5Tyne"]
    }
  ];

  const FABABEAN_SHEET_NAMES = ["FabaBeanRaw", "FabaBeansRaw", "FabaBean", "FabaBeans"];

  // Default 2022 faba bean raw plot dataset (simplified, one row per treatment)
  const RAW_PLOTS = [
    { Amendment: "control", "Yield t/ha": 2.4, "Pre sowing Labour": 40, "Treatment Input Cost Only /Ha": 0 },
    { Amendment: "deep_om_cp1", "Yield t/ha": 3.1, "Pre sowing Labour": 55, "Treatment Input Cost Only /Ha": 16500 },
    { Amendment: "deep_om_cp1_plus_liq_gypsum_cht", "Yield t/ha": 3.2, "Pre sowing Labour": 56, "Treatment Input Cost Only /Ha": 16850 },
    { Amendment: "deep_gypsum", "Yield t/ha": 2.9, "Pre sowing Labour": 50, "Treatment Input Cost Only /Ha": 500 },
    { Amendment: "deep_om_cp1_plus_pam", "Yield t/ha": 3.0, "Pre sowing Labour": 57, "Treatment Input Cost Only /Ha": 18000 },
    { Amendment: "deep_om_cp1_plus_ccm", "Yield t/ha": 3.25, "Pre sowing Labour": 58, "Treatment Input Cost Only /Ha": 21225 },
    { Amendment: "deep_ccm_only", "Yield t/ha": 2.95, "Pre sowing Labour": 52, "Treatment Input Cost Only /Ha": 3225 },
    { Amendment: "deep_om_cp2_plus_gypsum", "Yield t/ha": 3.3, "Pre sowing Labour": 60, "Treatment Input Cost Only /Ha": 24000 },
    { Amendment: "deep_liq_gypsum_cht", "Yield t/ha": 2.8, "Pre sowing Labour": 48, "Treatment Input Cost Only /Ha": 350 },
    { Amendment: "surface_silicon", "Yield t/ha": 2.7, "Pre sowing Labour": 45, "Treatment Input Cost Only /Ha": 1000 },
    { Amendment: "deep_liq_npks", "Yield t/ha": 3.0, "Pre sowing Labour": 53, "Treatment Input Cost Only /Ha": 2200 },
    { Amendment: "deep_ripping_only", "Yield t/ha": 2.85, "Pre sowing Labour": 47, "Treatment Input Cost Only /Ha": 0 }
  ];

  const LABOUR_COLUMNS = [
    "Pre sowing Labour",
    "Amendment Labour",
    "Sowing Labour",
    "Herbicide Labour",
    "Herbicide Labour 2",
    "Herbicide Labour 3",
    "Harvesting Labour",
    "Harvesting Labour 2"
  ];

  const OPERATING_COLUMNS = [
    "Treatment Input Cost Only /Ha",
    "Cavalier (Oxyfluofen 240)",
    "Factor",
    "Roundup CT",
    "Roundup Ultra Max",
    "Supercharge Elite Discontinued",
    "Platnium (Clethodim 360)",
    "Mentor",
    "Simazine 900",
    "Veritas Opti",
    "FLUTRIAFOL fungicide",
    "Barrack fungicide discontinued",
    "Talstar"
  ];

  let parsedExcel = null;

  function slugifyTreatmentName(name) {
    return String(name || "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  function parseNumber(value) {
    if (value === null || value === undefined || value === "") return NaN;
    if (typeof value === "number") return value;
    const cleaned = String(value).replace(/[\$,]/g, "");
    const n = parseFloat(cleaned);
    return Number.isFinite(n) ? n : NaN;
  }

  function computeTreatmentStatsFromRaw(rawRows) {
    const groups = new Map();

    rawRows.forEach(row => {
      const treatmentName = String(row.Amendment || "").trim();
      if (!treatmentName) return;

      let g = groups.get(treatmentName);
      if (!g) {
        g = { name: treatmentName, sumYield: 0, nYield: 0, sumLabour: 0, sumOperating: 0 };
        groups.set(treatmentName, g);
      }

      const y = parseNumber(row["Yield t/ha"]);
      if (!Number.isNaN(y)) {
        g.sumYield += y;
        g.nYield += 1;
      }

      let labour = 0;
      LABOUR_COLUMNS.forEach(col => {
        const v = parseNumber(row[col]);
        if (!Number.isNaN(v)) labour += v;
      });
      g.sumLabour += labour;

      let op = 0;
      OPERATING_COLUMNS.forEach(col => {
        const v = parseNumber(row[col]);
        if (!Number.isNaN(v)) op += v;
      });
      g.sumOperating += op;
    });

    let controlMeanYield = null;
    for (const [name, g] of groups.entries()) {
      const label = name.toLowerCase();
      if (label.includes("control")) {
        if (g.nYield > 0) controlMeanYield = g.sumYield / g.nYield;
        break;
      }
    }

    const treatments = [];
    for (const [name, g] of groups.entries()) {
      const meanYield = g.nYield > 0 ? g.sumYield / g.nYield : 0;
      const meanLabour = g.nYield > 0 ? g.sumLabour / g.nYield : 0;
      const meanOperating = g.nYield > 0 ? g.sumOperating / g.nYield : 0;

      treatments.push({
        id: slugifyTreatmentName(name),
        label: name,
        isControl: name.toLowerCase().includes("control"),
        meanYieldTHa: meanYield,
        labourCostPerHa: meanLabour,
        additionalOperatingCostPerHa: meanOperating,
        yield_uplift_vs_control_t_ha: controlMeanYield === null ? 0 : meanYield - controlMeanYield
      });
    }

    return treatments;
  }

  function initTreatmentDeltas() {
    if (!Array.isArray(model.treatments)) model.treatments = [];
    if (!Array.isArray(model.outputs)) model.outputs = [];

    model.treatments.forEach(t => {
      if (!t.deltas || typeof t.deltas !== "object") t.deltas = {};
      model.outputs.forEach(o => {
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      });

      if (typeof t.labourCost === "undefined") t.labourCost = Number(t.annualCost || 0) || 0;
      if (typeof t.materialsCost === "undefined") t.materialsCost = 0;
      if (typeof t.servicesCost === "undefined") t.servicesCost = 0;
      if (typeof t.capitalCost === "undefined") t.capitalCost = 0;

      if (typeof t.adoption !== "number" || isNaN(t.adoption)) t.adoption = 1;
      if (typeof t.area !== "number" || isNaN(t.area)) t.area = 0;

      delete t.annualCost;
    });
  }

  function hardenModel() {
    model.__meta__ = model.__meta__ || {};
    model.__meta__.appVersion = APP_VERSION;
    model.__meta__.schemaVersion = MODEL_SCHEMA_VERSION;

    model.project = model.project || deepClone(DEFAULT_MODEL.project);
    model.time = model.time || deepClone(DEFAULT_MODEL.time);
    model.outputsMeta = model.outputsMeta || deepClone(DEFAULT_MODEL.outputsMeta);
    model.outputs = Array.isArray(model.outputs) ? model.outputs : deepClone(DEFAULT_MODEL.outputs);
    model.treatments = Array.isArray(model.treatments) ? model.treatments : deepClone(DEFAULT_MODEL.treatments);
    model.benefits = Array.isArray(model.benefits) ? model.benefits : deepClone(DEFAULT_MODEL.benefits);
    model.otherCosts = Array.isArray(model.otherCosts) ? model.otherCosts : deepClone(DEFAULT_MODEL.otherCosts);
    model.adoption = model.adoption || deepClone(DEFAULT_MODEL.adoption);
    model.risk = model.risk || deepClone(DEFAULT_MODEL.risk);
    model.sim = model.sim || deepClone(DEFAULT_MODEL.sim);

    if (!Array.isArray(model.time.discountSchedule) || !model.time.discountSchedule.length) {
      model.time.discountSchedule = deepClone(DEFAULT_DISCOUNT_SCHEDULE);
    }
    initTreatmentDeltas();

    // Ensure exactly one control if any treatments exist
    if (model.treatments.length) {
      const anyControl = model.treatments.some(t => !!t.isControl);
      if (!anyControl) model.treatments[0].isControl = true;
      const firstControl = model.treatments.find(t => t.isControl);
      model.treatments.forEach(t => (t.isControl = t === firstControl));
    }
  }

  function applyTrialTreatmentsIfAvailable() {
    if (!RAW_PLOTS.length) return;
    const stats = computeTreatmentStatsFromRaw(RAW_PLOTS);
    if (!stats.length) return;

    const byId = new Map(stats.map(s => [s.id, s]));
    const yieldOutput = model.outputs.find(o => o.name.toLowerCase().includes("yield"));
    const yieldId = yieldOutput ? yieldOutput.id : null;

    const merged = TRIAL_TREATMENT_CONFIG.map(cfg => {
      const s = byId.get(cfg.id);
      const costs = Object.assign({}, cfg.costs);
      const agr = Object.assign({}, cfg.agronomy);

      if (s) {
        agr.mean_yield_t_ha = s.meanYieldTHa;
        agr.yield_uplift_vs_control_t_ha = s.yield_uplift_vs_control_t_ha;
        costs.labourCostPerHa = isNum(costs.labourCostPerHa) ? costs.labourCostPerHa : s.labourCostPerHa;
        costs.additionalOperatingCostPerHa = isNum(costs.additionalOperatingCostPerHa)
          ? costs.additionalOperatingCostPerHa
          : s.additionalOperatingCostPerHa;
      } else {
        // If missing stats, keep config defaults; ensure numbers
        costs.labourCostPerHa = asNum(costs.labourCostPerHa, 0);
        costs.additionalOperatingCostPerHa = asNum(costs.additionalOperatingCostPerHa, 0);
      }

      return {
        id: cfg.id,
        label: cfg.label,
        category: cfg.category,
        controlFlag: !!cfg.controlFlag,
        agronomy: agr,
        costs,
        capitalAssets: cfg.capitalAssets || []
      };
    });

    model.treatments = merged.map(tt => {
      const materialsPerHa = asNum(tt.costs.treatmentInputCostPerHa, 0);
      const labourPerHa = asNum(tt.costs.labourCostPerHa, 0);
      const operatingPerHa = asNum(tt.costs.additionalOperatingCostPerHa, 0);

      const t = {
        id: uid(),
        name: tt.label,
        area: 100,
        adoption: 1,
        deltas: {},
        labourCost: labourPerHa,
        materialsCost: materialsPerHa + operatingPerHa,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "Farm Trials",
        isControl: !!tt.controlFlag,
        notes: tt.agronomy && tt.agronomy.notes ? tt.agronomy.notes : ""
      };

      model.outputs.forEach(o => (t.deltas[o.id] = 0));
      if (yieldId && tt.agronomy && typeof tt.agronomy.yield_uplift_vs_control_t_ha === "number") {
        t.deltas[yieldId] = tt.agronomy.yield_uplift_vs_control_t_ha;
      }
      return t;
    });

    initTreatmentDeltas();
    autosaveDebounced();
    showToast("Faba bean trial treatments loaded and calibrated from default dataset.");
  }

  // =========================
  // DISCOUNT SCHEDULE SUPPORT
  // =========================
  function discountRateForYearIndex(tIndex, scenario = "base") {
    // tIndex: 0..N, where 0 is year 0, 1 is first operating year.
    // Uses absolute year = startYear + (tIndex - 1) for operating years; year 0 uses startYear as reference.
    const baseYear = asNum(model.time.startYear, new Date().getFullYear());
    const absYear = tIndex <= 0 ? baseYear : baseYear + (tIndex - 1);

    const sched = model.time.discountSchedule || DEFAULT_DISCOUNT_SCHEDULE;
    const row = sched.find(r => absYear >= r.from && absYear <= r.to);
    if (!row) {
      if (scenario === "low") return asNum(model.time.discLow, 0);
      if (scenario === "high") return asNum(model.time.discHigh, 0);
      return asNum(model.time.discBase, 0);
    }
    if (scenario === "low") return asNum(row.low, asNum(model.time.discLow, 0));
    if (scenario === "high") return asNum(row.high, asNum(model.time.discHigh, 0));
    return asNum(row.base, asNum(model.time.discBase, 0));
  }

  function presentValue(series, ratePct) {
    let pv = 0;
    for (let t = 0; t < series.length; t++) pv += series[t] / Math.pow(1 + ratePct / 100, t);
    return pv;
  }

  function presentValueWithSchedule(series, scenario = "base") {
    // PV = sum_t series[t] / Π_{k=1..t}(1+r_k)
    let pv = 0;
    let df = 1;
    for (let t = 0; t < series.length; t++) {
      if (t === 0) {
        pv += series[t];
      } else {
        const r = discountRateForYearIndex(t, scenario) / 100;
        df *= 1 + r;
        pv += series[t] / df;
      }
    }
    return pv;
  }

  // =========================
  // FINANCE METRICS
  // =========================
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

  // =========================
  // CASHFLOWS
  // =========================
  function additionalBenefitsSeries(N, baseYear, adoptMul, risk) {
    const series = new Array(N + 1).fill(0);

    model.benefits.forEach(b => {
      const cat = String(b.category || "").toUpperCase();
      const linkA = !!b.linkAdoption;
      const linkR = !!b.linkRisk;
      const A = linkA ? clamp(adoptMul, 0, 1) : 1;
      const R = linkR ? 1 - clamp(risk, 0, 1) : 1;
      const g = asNum(b.growthPct, 0);

      const addAnnual = (yearIndex, baseAmount, tFromStart) => {
        const grown = baseAmount * Math.pow(1 + g / 100, tFromStart);
        if (yearIndex >= 1 && yearIndex <= N) series[yearIndex] += grown * A * R;
      };
      const addOnce = (absYear, amount) => {
        const idx = absYear - baseYear + 1;
        if (idx >= 0 && idx <= N) series[idx] += amount * A * R;
      };

      const sy = asNum(b.startYear, baseYear);
      const ey = asNum(b.endYear, sy);
      const yr = asNum(b.year, sy);

      if (b.frequency === "Once" || cat === "C6") {
        const amount = asNum(b.annualAmount, 0);
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
            const v = asNum(b.unitValue, 0);
            const q = asNum(cat === "C3" ? b.abatement : b.quantity, 0);
            amt = v * q;
            break;
          }
          case "C4":
          case "C5":
          case "C8":
            amt = asNum(b.annualAmount, 0);
            break;
          case "C7": {
            const p0 = asNum(b.p0, 0);
            const p1 = asNum(b.p1, 0);
            const c = asNum(b.consequence, 0);
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
    const N = asNum(model.time.years, 10);
    const baseYear = asNum(model.time.startYear, new Date().getFullYear());

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
        const delta = asNum(t.deltas?.[o.id], 0);
        const v = asNum(o.value, 0);
        valuePerHa += delta * v;
      });

      const area = asNum(t.area, 0);
      const benefit = valuePerHa * area * (1 - clamp(risk, 0, 1)) * adopt;

      const annualCostPerHa = asNum(t.materialsCost, 0) + asNum(t.servicesCost, 0) + asNum(t.labourCost, 0);
      const opCost = annualCostPerHa * area;
      const cap = asNum(t.capitalCost, 0);

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
      const type = String(c.type || "annual").toLowerCase();
      if (type === "annual") {
        const a = asNum(c.annual, 0);
        const sy = asNum(c.startYear, baseYear);
        const ey = asNum(c.endYear, sy);
        for (let y = sy; y <= ey; y++) {
          const idx = y - baseYear + 1;
          if (idx >= 1 && idx <= N) {
            otherAnnualByYear[idx] += a;
            if (c.constrained) otherConstrAnnualByYear[idx] += a;
          }
        }
      } else if (type === "capital") {
        const cap = asNum(c.capital, 0);
        const cy = asNum(c.year, baseYear);
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

  function computeAll(rate, adoptMul, risk, bcrMode) {
    const { benefitByYear, costByYear, constrainedCostByYear, cf, annualGM } = buildCashflows({ forRate: rate, adoptMul, risk });

    const useSchedule = !!model.time.useDiscountSchedule || !!document.getElementById("useDiscSchedule")?.checked;

    const pvBenefits = useSchedule ? presentValueWithSchedule(benefitByYear, "base") : presentValue(benefitByYear, rate);
    const pvCosts = useSchedule ? presentValueWithSchedule(costByYear, "base") : presentValue(costByYear, rate);
    const pvCostsConstrained = useSchedule ? presentValueWithSchedule(constrainedCostByYear, "base") : presentValue(constrainedCostByYear, rate);

    const npv = pvBenefits - pvCosts;
    const denom = bcrMode === "constrained" ? pvCostsConstrained : pvCosts;
    const bcr = denom > 0 ? pvBenefits / denom : NaN;

    const irrVal = irr(cf);
    const mirrVal = mirr(cf, asNum(model.time.mirrFinance, 0), asNum(model.time.mirrReinvest, 0));
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

  // =========================
  // DOM HELPERS
  // =========================
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));
  const num = sel => asNum(document.querySelector(sel)?.value, 0);
  const setVal = (sel, text) => {
    const el = document.querySelector(sel);
    if (el) el.textContent = text;
  };

  // =========================
  // TABS (ACCESSIBLE + DEEPLINK)
  // =========================
  function setHash(tabKey) {
    try {
      if (!tabKey) return;
      const next = `#tab=${encodeURIComponent(tabKey)}`;
      if (location.hash !== next) history.replaceState(null, "", next);
    } catch {}
  }

  function getHashTab() {
    try {
      const h = location.hash || "";
      const m = h.match(/tab=([^&]+)/);
      return m ? decodeURIComponent(m[1]) : null;
    } catch {
      return null;
    }
  }

  function switchTab(target, { silent = false } = {}) {
    if (!target) return;

    const navEls = $$("[data-tab],[data-tab-target],[data-tab-jump]");
    navEls.forEach(el => {
      const key = el.dataset.tab || el.dataset.tabTarget || el.dataset.tabJump;
      const isActive = key === target;
      el.classList.toggle("active", isActive);
      el.setAttribute("aria-selected", isActive ? "true" : "false");
      if (el.hasAttribute("role")) el.setAttribute("tabindex", isActive ? "0" : "-1");
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

    try {
      localStorage.setItem(`${STORAGE_NS}::active_tab`, target);
    } catch {}

    setHash(target);
    if (!silent) {
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
  }

  function initTabs() {
    document.addEventListener("click", e => {
      const el = e.target.closest("[data-tab],[data-tab-target],[data-tab-jump]");
      if (!el) return;
      const target = el.dataset.tab || el.dataset.tabTarget || el.dataset.tabJump;
      if (!target) return;
      e.preventDefault();
      switchTab(target);
      if (el.dataset.toast !== "false") showToast(`Switched to ${target} tab.`);
    });

    // Keyboard navigation (left/right) on navs with role=tab (if present)
    document.addEventListener("keydown", e => {
      const active = document.activeElement;
      if (!active) return;
      if (!active.matches?.("[data-tab],[data-tab-target]")) return;
      if (e.key !== "ArrowLeft" && e.key !== "ArrowRight") return;
      const items = $$("[data-tab],[data-tab-target]").filter(x => x.offsetParent !== null);
      const idx = items.indexOf(active);
      if (idx < 0) return;
      e.preventDefault();
      const nextIdx = e.key === "ArrowRight" ? (idx + 1) % items.length : (idx - 1 + items.length) % items.length;
      items[nextIdx]?.focus();
      const target = items[nextIdx]?.dataset.tab || items[nextIdx]?.dataset.tabTarget;
      if (target) switchTab(target, { silent: true });
    });

    const hashTab = getHashTab();
    const storedTab = (() => {
      try {
        return localStorage.getItem(`${STORAGE_NS}::active_tab`);
      } catch {
        return null;
      }
    })();

    const preferred = hashTab || storedTab;

    const findAnyNav = () =>
      document.querySelector("[data-tab].active, [data-tab-target].active, [data-tab-jump].active") ||
      document.querySelector("[data-tab], [data-tab-target], [data-tab-jump]");

    if (preferred) {
      switchTab(preferred, { silent: true });
      return;
    }

    const activeNav = findAnyNav();
    if (activeNav) {
      const target = activeNav.dataset.tab || activeNav.dataset.tabTarget || activeNav.dataset.tabJump;
      if (target) {
        switchTab(target, { silent: true });
        return;
      }
    }

    const firstPanel = document.querySelector(".tab-panel");
    if (firstPanel) {
      const key = firstPanel.dataset.tabPanel || (firstPanel.id ? firstPanel.id.replace(/^tab-/, "") : "");
      if (key) switchTab(key, { silent: true });
    }
  }

  // =========================
  // ACTIONS
  // =========================
  function initActions() {
    document.addEventListener("click", e => {
      const el = e.target.closest("#recalc, #getResults, [data-action='recalc']");
      if (!el) return;
      e.preventDefault();
      e.stopPropagation();
      safeRun("Recalc", () => {
        calcAndRender();
        showToast("Base case economic indicators recalculated.");
      });
    });

    document.addEventListener("click", e => {
      const el = e.target.closest("#runSim, [data-action='run-sim']");
      if (!el) return;
      e.preventDefault();
      e.stopPropagation();
      safeRun("RunSim", () => runSimulation());
    });
  }

  // =========================
  // FORM BINDING
  // =========================
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

    const useSched = $("#useDiscSchedule");
    if (useSched) useSched.checked = !!model.time.useDiscountSchedule;

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
        safeRun("CombinedRisk", () => {
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
          autosaveDebounced();
          showToast("Combined risk updated from component risks.");
        });
      });
    }

    const addCostBtn = $("#addCost");
    if (addCostBtn) {
      addCostBtn.addEventListener("click", e => {
        e.stopPropagation();
        safeRun("AddCost", () => {
          const c = {
            id: uid(),
            label: "New cost",
            type: "annual",
            category: "Capital",
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
          autosaveDebounced();
          showToast("New cost item added.");
        });
      });
    }

    document.addEventListener("input", e => {
      const t = e.target;
      if (!t) return;

      if (t.dataset && t.dataset.discPeriod !== undefined) {
        const idx = +t.dataset.discPeriod;
        const scenario = t.dataset.scenario;
        if (!model.time.discountSchedule) model.time.discountSchedule = deepClone(DEFAULT_DISCOUNT_SCHEDULE);
        const row = model.time.discountSchedule[idx];
        if (row && scenario) {
          const val = +t.value;
          if (scenario === "low") row.low = val;
          else if (scenario === "base") row.base = val;
          else if (scenario === "high") row.high = val;
          calcAndRenderDebounced();
          autosaveDebounced();
        }
        return;
      }

      const id = t.id;
      if (!id) return;

      switch (id) {
        case "projectName": model.project.name = t.value; break;
        case "projectLead": model.project.lead = t.value; break;
        case "analystNames": model.project.analysts = t.value; break;
        case "projectTeam": model.project.team = t.value; break;
        case "projectSummary": model.project.summary = t.value; break;
        case "projectObjectives": model.project.objectives = t.value; break;
        case "projectActivities": model.project.activities = t.value; break;
        case "stakeholderGroups": model.project.stakeholders = t.value; break;
        case "lastUpdated": model.project.lastUpdated = t.value; break;
        case "projectGoal": model.project.goal = t.value; break;
        case "withProject": model.project.withProject = t.value; break;
        case "withoutProject": model.project.withoutProject = t.value; break;
        case "organisation": model.project.organisation = t.value; break;
        case "contactEmail": model.project.contactEmail = t.value; break;
        case "contactPhone": model.project.contactPhone = t.value; break;

        case "startYear": model.time.startYear = +t.value; break;
        case "projectStartYear": model.time.projectStartYear = +t.value; break;
        case "years": model.time.years = +t.value; break;
        case "discBase": model.time.discBase = +t.value; break;
        case "discLow": model.time.discLow = +t.value; break;
        case "discHigh": model.time.discHigh = +t.value; break;
        case "mirrFinance": model.time.mirrFinance = +t.value; break;
        case "mirrReinvest": model.time.mirrReinvest = +t.value; break;

        case "useDiscSchedule": model.time.useDiscountSchedule = !!t.checked; break;

        case "adoptBase": model.adoption.base = +t.value; break;
        case "adoptLow": model.adoption.low = +t.value; break;
        case "adoptHigh": model.adoption.high = +t.value; break;

        case "riskBase": model.risk.base = +t.value; break;
        case "riskLow": model.risk.low = +t.value; break;
        case "riskHigh": model.risk.high = +t.value; break;
        case "rTech": model.risk.tech = +t.value; break;
        case "rNonCoop": model.risk.nonCoop = +t.value; break;
        case "rSocio": model.risk.socio = +t.value; break;
        case "rFin": model.risk.fin = +t.value; break;
        case "rMan": model.risk.man = +t.value; break;

        case "simN": model.sim.n = +t.value; break;
        case "targetBCR":
          model.sim.targetBCR = +t.value;
          if ($("#simBcrTargetLabel")) $("#simBcrTargetLabel").textContent = t.value;
          break;
        case "bcrMode": model.sim.bcrMode = t.value; break;
        case "randSeed": model.sim.seed = t.value ? +t.value : null; break;

        case "simVarPct": model.sim.variationPct = +t.value || 20; break;
        case "simVaryOutputs": model.sim.varyOutputs = t.value === "true"; break;
        case "simVaryTreatCosts": model.sim.varyTreatCosts = t.value === "true"; break;
        case "simVaryInputCosts": model.sim.varyInputCosts = t.value === "true"; break;

        case "systemType": model.outputsMeta.systemType = t.value; break;
        case "outputAssumptions": model.outputsMeta.assumptions = t.value; break;
      }

      calcAndRenderDebounced();
      autosaveDebounced();
    });

    const saveProjectBtn = $("#saveProject");
    if (saveProjectBtn) {
      saveProjectBtn.addEventListener("click", e => {
        e.stopPropagation();
        safeRun("SaveProject", () => {
          const payload = deepClone(model);
          payload.__meta__ = payload.__meta__ || {};
          payload.__meta__.lastSaved = new Date().toISOString();
          const data = JSON.stringify(payload, null, 2);
          downloadFile("cba_" + (model.project.name || "project").replace(/\s+/g, "_") + ".json", data, "application/json");
          showToast("Project JSON downloaded.");
        });
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

        safeRun("LoadProject", () => {
          const obj = safeJsonParse(text, null);
          if (!obj) {
            alert("Invalid JSON file.");
            return;
          }

          const merged = deepMerge(deepClone(DEFAULT_MODEL), obj);
          Object.keys(model).forEach(k => delete model[k]);
          Object.assign(model, merged);

          hardenModel();
          renderAll();
          setBasicsFieldsFromModel();
          calcAndRender();
          autosaveDebounced();
          showToast("Project JSON loaded and applied.");
        });

        e.target.value = "";
      });
    }

    const exportCsvBtn = $("#exportCsv");
    const exportCsvFootBtn = $("#exportCsvFoot");
    if (exportCsvBtn) exportCsvBtn.addEventListener("click", e => { e.stopPropagation(); exportAllCsv(); });
    if (exportCsvFootBtn) exportCsvFootBtn.addEventListener("click", e => { e.stopPropagation(); exportAllCsv(); });

    const exportXlsxBtn = $("#exportXlsx");
    const exportXlsxFootBtn = $("#exportXlsxFoot");
    if (exportXlsxBtn) exportXlsxBtn.addEventListener("click", e => { e.stopPropagation(); exportAllXlsx(); });
    if (exportXlsxFootBtn) exportXlsxFootBtn.addEventListener("click", e => { e.stopPropagation(); exportAllXlsx(); });

    const exportPdfBtn = $("#exportPdf");
    const exportPdfFootBtn = $("#exportPdfFoot");
    if (exportPdfBtn) exportPdfBtn.addEventListener("click", e => { e.stopPropagation(); exportPdf(); showToast("Print dialog opened for PDF export."); });
    if (exportPdfFootBtn) exportPdfFootBtn.addEventListener("click", e => { e.stopPropagation(); exportPdf(); showToast("Print dialog opened for PDF export."); });

    const parseExcelBtn = $("#parseExcel");
    const importExcelBtn = $("#importExcel");
    if (parseExcelBtn) parseExcelBtn.addEventListener("click", e => { e.stopPropagation(); handleParseExcel(); });
    if (importExcelBtn) importExcelBtn.addEventListener("click", e => { e.stopPropagation(); commitExcelToModel(); });

    const downloadTemplateBtn = $("#downloadTemplate");
    const downloadSampleBtn = $("#downloadSample");
    if (downloadTemplateBtn) downloadTemplateBtn.addEventListener("click", e => { e.stopPropagation(); downloadExcelTemplate(); });
    if (downloadSampleBtn) downloadSampleBtn.addEventListener("click", e => { e.stopPropagation(); downloadSampleDataset(); });

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

    const resetBtn = $("#resetToDefaults");
    if (resetBtn) {
      resetBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        if (!confirm("Reset the tool to default settings? This will overwrite the current project in memory.")) return;
        Object.keys(model).forEach(k => delete model[k]);
        Object.assign(model, deepClone(DEFAULT_MODEL));
        hardenModel();
        applyTrialTreatmentsIfAvailable();
        renderAll();
        setBasicsFieldsFromModel();
        calcAndRender();
        autosaveDebounced();
        showToast("Reset to defaults.");
      });
    }
  }

  // =========================
  // ADD BUTTONS
  // =========================
  function initAddButtons() {
    const addOutputBtn = $("#addOutput");
    if (addOutputBtn) {
      addOutputBtn.addEventListener("click", e => {
        e.stopPropagation();
        safeRun("AddOutput", () => {
          const id = uid();
          model.outputs.push({ id, name: "Custom output", unit: "unit", value: 0, source: "Input Directly" });
          model.treatments.forEach(t => { t.deltas[id] = 0; });
          renderOutputs();
          renderTreatments();
          renderDatabaseTags();
          calcAndRender();
          autosaveDebounced();
          showToast("New output metric added.");
        });
      });
    }

    const addTreatmentBtn = $("#addTreatment");
    if (addTreatmentBtn) {
      addTreatmentBtn.addEventListener("click", e => {
        e.stopPropagation();
        safeRun("AddTreatment", () => {
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
          model.outputs.forEach(o => (t.deltas[o.id] = 0));
          model.treatments.push(t);
          renderTreatments();
          renderDatabaseTags();
          calcAndRender();
          autosaveDebounced();
          showToast("New treatment added.");
        });
      });
    }

    const addBenefitBtn = $("#addBenefit");
    if (addBenefitBtn) {
      addBenefitBtn.addEventListener("click", e => {
        e.stopPropagation();
        safeRun("AddBenefit", () => {
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
          autosaveDebounced();
          showToast("New benefit item added.");
        });
      });
    }
  }

  // =========================
  // RENDERERS
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
          <div class="field"><label>Name</label><input value="${esc(o.name)}" data-k="name" data-id="${o.id}" /></div>
          <div class="field"><label>Unit</label><input value="${esc(o.unit)}" data-k="unit" data-id="${o.id}" /></div>
          <div class="field"><label>Value ($/unit)</label><input type="number" step="0.01" value="${asNum(o.value, 0)}" data-k="value" data-id="${o.id}" /></div>
          <div class="field"><label>Source</label>
            <select data-k="source" data-id="${o.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"].map(s => `<option ${s === o.source ? "selected" : ""}>${s}</option>`).join("")}
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

    model.treatments.forEach(t => { if (!(id in t.deltas)) t.deltas[id] = 0; });
    renderTreatments();
    renderDatabaseTags();
    calcAndRenderDebounced();
    autosaveDebounced();
  }

  function onOutputDelete(e) {
    const id = e.target.dataset.delOutput;
    if (!id) return;
    if (!confirm("Remove this output metric?")) return;
    model.outputs = model.outputs.filter(o => o.id !== id);
    model.treatments.forEach(t => { delete t.deltas[id]; });
    renderOutputs();
    renderTreatments();
    renderDatabaseTags();
    calcAndRender();
    autosaveDebounced();
    showToast("Output metric removed.");
  }

  function renderTreatments() {
    const root = $("#treatmentsList");
    if (!root) return;
    root.innerHTML = "";
    const list = [...model.treatments];

    list.forEach(t => {
      const materials = asNum(t.materialsCost, 0);
      const services = asNum(t.servicesCost, 0);
      const labour = asNum(t.labourCost, 0);
      const totalPerHa = materials + services + labour;

      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>Treatment: ${esc(t.name)}</h4>
        <div class="row">
          <div class="field"><label>Name</label><input value="${esc(t.name)}" data-tk="name" data-id="${t.id}" /></div>
          <div class="field"><label>Area (ha)</label><input type="number" step="0.01" value="${asNum(t.area, 0)}" data-tk="area" data-id="${t.id}" /></div>
          <div class="field"><label>Source</label>
            <select data-tk="source" data-id="${t.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"].map(s => `<option ${s === t.source ? "selected" : ""}>${s}</option>`).join("")}
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
          <div class="field"><label>Materials cost ($/ha)</label><input type="number" step="0.01" value="${materials}" data-tk="materialsCost" data-id="${t.id}" /></div>
          <div class="field"><label>Services cost ($/ha)</label><input type="number" step="0.01" value="${services}" data-tk="servicesCost" data-id="${t.id}" /></div>
          <div class="field"><label>Labour cost ($/ha)</label><input type="number" step="0.01" value="${labour}" data-tk="labourCost" data-id="${t.id}" /></div>
          <div class="field"><label>Total cost ($/ha)</label><input type="number" step="0.01" value="${totalPerHa}" readonly data-total-cost="${t.id}" /></div>
          <div class="field"><label>Capital cost ($, year 0)</label><input type="number" step="0.01" value="${asNum(t.capitalCost, 0)}" data-tk="capitalCost" data-id="${t.id}" /></div>
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
          ${model.outputs.map(o => `
            <div class="field">
              <label>${esc(o.name)} (${esc(o.unit)})</label>
              <input type="number" step="0.0001" value="${asNum(t.deltas?.[o.id], 0)}" data-td="${o.id}" data-id="${t.id}" />
            </div>
          `).join("")}
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
        if (tk === "constrained") t[tk] = e.target.value === "true";
        else if (tk === "name" || tk === "source" || tk === "notes") t[tk] = e.target.value;
        else if (tk === "isControl") {
          const val = e.target.value === "control";
          model.treatments.forEach(tt => (tt.isControl = false));
          if (val) t.isControl = true;
          renderTreatments();
          calcAndRenderDebounced();
          autosaveDebounced();
          showToast(`Control treatment set to ${t.name}.`);
          return;
        } else t[tk] = +e.target.value;

        if (tk === "materialsCost" || tk === "servicesCost" || tk === "labourCost") {
          const container = e.target.closest(".item");
          if (container) {
            const mats = asNum(container.querySelector(`input[data-tk="materialsCost"][data-id="${id}"]`)?.value, 0);
            const serv = asNum(container.querySelector(`input[data-tk="servicesCost"][data-id="${id}"]`)?.value, 0);
            const lab = asNum(container.querySelector(`input[data-tk="labourCost"][data-id="${id}"]`)?.value, 0);
            const totalField = container.querySelector(`input[data-total-cost="${id}"]`);
            if (totalField) totalField.value = mats + serv + lab;
          }
        }
      }

      const td = e.target.dataset.td;
      if (td) t.deltas[td] = +e.target.value;

      calcAndRenderDebounced();
      autosaveDebounced();
    };

    root.addEventListener("click", e => {
      const id = e.target.dataset.delTreatment;
      if (!id) return;
      if (!confirm("Remove this treatment?")) return;
      model.treatments = model.treatments.filter(x => x.id !== id);
      if (!model.treatments.some(t => t.isControl) && model.treatments[0]) model.treatments[0].isControl = true;
      renderTreatments();
      renderDatabaseTags();
      calcAndRender();
      autosaveDebounced();
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
              ${["C1","C2","C3","C4","C5","C6","C7","C8"].map(c => `<option ${c === b.category ? "selected" : ""}>${c}</option>`).join("")}
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
          <div class="field"><label>Start year</label><input type="number" value="${asNum(b.startYear, model.time.startYear)}" data-bk="startYear" data-id="${b.id}" /></div>
          <div class="field"><label>End year</label><input type="number" value="${asNum(b.endYear, model.time.startYear)}" data-bk="endYear" data-id="${b.id}" /></div>
        </div>

        <div class="row-6">
          <div class="field"><label>Once year</label><input type="number" value="${asNum(b.year, model.time.startYear)}" data-bk="year" data-id="${b.id}" /></div>
          <div class="field"><label>Unit value ($)</label><input type="number" step="0.01" value="${asNum(b.unitValue, 0)}" data-bk="unitValue" data-id="${b.id}" /></div>
          <div class="field"><label>Quantity</label><input type="number" step="0.01" value="${asNum(b.quantity, 0)}" data-bk="quantity" data-id="${b.id}" /></div>
          <div class="field"><label>Abatement</label><input type="number" step="0.01" value="${asNum(b.abatement, 0)}" data-bk="abatement" data-id="${b.id}" /></div>
          <div class="field"><label>Annual amount ($)</label><input type="number" step="0.01" value="${asNum(b.annualAmount, 0)}" data-bk="annualAmount" data-id="${b.id}" /></div>
          <div class="field"><label>Growth (% per year)</label><input type="number" step="0.01" value="${asNum(b.growthPct, 0)}" data-bk="growthPct" data-id="${b.id}" /></div>
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
          <div class="field"><label>P0 (baseline probability)</label><input type="number" step="0.001" value="${asNum(b.p0, 0)}" data-bk="p0" data-id="${b.id}" /></div>
          <div class="field"><label>P1 (with project probability)</label><input type="number" step="0.001" value="${asNum(b.p1, 0)}" data-bk="p1" data-id="${b.id}" /></div>
          <div class="field"><label>Consequence ($)</label><input type="number" step="0.01" value="${asNum(b.consequence, 0)}" data-bk="consequence" data-id="${b.id}" /></div>
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
      autosaveDebounced();
    };

    root.addEventListener("click", e => {
      const id = e.target.dataset.delBenefit;
      if (!id) return;
      if (!confirm("Remove this benefit item?")) return;
      model.benefits = model.benefits.filter(x => x.id !== id);
      renderBenefits();
      calcAndRender();
      autosaveDebounced();
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
          <div class="field"><label>Annual ($ per year)</label><input type="number" step="0.01" value="${asNum(c.annual, 0)}" data-ck="annual" data-id="${c.id}" /></div>
          <div class="field"><label>Start year</label><input type="number" value="${asNum(c.startYear, model.time.startYear)}" data-ck="startYear" data-id="${c.id}" /></div>
          <div class="field"><label>End year</label><input type="number" value="${asNum(c.endYear, model.time.startYear)}" data-ck="endYear" data-id="${c.id}" /></div>
        </div>
        <div class="row-6">
          <div class="field"><label>Capital ($)</label><input type="number" step="0.01" value="${asNum(c.capital, 0)}" data-ck="capital" data-id="${c.id}" /></div>
          <div class="field"><label>Capital year</label><input type="number" value="${asNum(c.year, model.time.startYear)}" data-ck="year" data-id="${c.id}" /></div>
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
          <div class="field"><label>Life (years)</label><input type="number" step="1" min="1" value="${asNum(c.depLife, 5)}" data-ck="depLife" data-id="${c.id}" /></div>
          <div class="field"><label>Declining rate (% per year)</label><input type="number" step="1" value="${asNum(c.depRate, 30)}" data-ck="depRate" data-id="${c.id}" /></div>
          `
              : `
          <div class="field" style="display:none"></div>
          <div class="field" style="display:none"></div>
          <div class="field" style="display:none"></div>
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
        if (k === "category" && c.category !== "Capital") c.depMethod = "none";
      } else if (k === "constrained") c[k] = e.target.value === "true";
      else c[k] = +e.target.value;

      calcAndRenderDebounced();
      autosaveDebounced();
    };

    root.addEventListener("click", e => {
      const id = e.target.dataset.delCost;
      if (!id) return;
      if (!confirm("Remove this cost item?")) return;
      model.otherCosts = model.otherCosts.filter(x => x.id !== id);
      renderCosts();
      calcAndRender();
      autosaveDebounced();
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
                ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"].map(s => `<option ${s === o.source ? "selected" : ""}>${s}</option>`).join("")}
              </select>
            </div>
          </div>`;
        outRoot.appendChild(el);
      });
      outRoot.onchange = e => {
        const id = e.target.dataset.dbOut;
        const o = model.outputs.find(x => x.id === id);
        if (o) {
          o.source = e.target.value;
          autosaveDebounced();
        }
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
                ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"].map(s => `<option ${s === t.source ? "selected" : ""}>${s}</option>`).join("")}
              </select>
            </div>
          </div>`;
        tRoot.appendChild(el);
      });
      tRoot.onchange = e => {
        const id = e.target.dataset.dbT;
        const t = model.treatments.find(x => x.id === id);
        if (t) {
          t.source = e.target.value;
          autosaveDebounced();
        }
      };
    }
  }

  function renderAll() {
    renderOutputs();
    renderTreatments();
    renderBenefits();
    renderDatabaseTags();
    renderCosts();
  }

  // =========================
  // SUMMARY TABLES
  // =========================
  function computeSingleTreatmentMetrics(t, rate, years, adoptMul, risk) {
    let valuePerHa = 0;
    model.outputs.forEach(o => {
      valuePerHa += asNum(t.deltas?.[o.id], 0) * asNum(o.value, 0);
    });

    const adopt = clamp(adoptMul, 0, 1);
    const area = asNum(t.area, 0);
    const annualBen = valuePerHa * area * (1 - clamp(risk, 0, 1)) * adopt;

    const annualCostPerHa = asNum(t.materialsCost, 0) + asNum(t.servicesCost, 0) + asNum(t.labourCost, 0);
    const annualCost = annualCostPerHa * area;
    const cap = asNum(t.capitalCost, 0);

    const pvBen = annualBen * annuityFactor(years, rate);
    const pvCost = cap + annualCost * annuityFactor(years, rate);

    const bcr = pvCost > 0 ? pvBen / pvCost : NaN;
    const npv = pvBen - pvCost;

    const cf = new Array(years + 1).fill(0);
    cf[0] = -cap;
    for (let i = 1; i <= years; i++) cf[i] = annualBen - annualCost;

    const irrVal = irr(cf);
    const mirrVal = mirr(cf, asNum(model.time.mirrFinance, 0), asNum(model.time.mirrReinvest, 0));
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
      .map(t => ({ t, m: computeSingleTreatmentMetrics(t, rate, asNum(model.time.years, 10), adoptMul, risk) }))
      .sort((a, b) => {
        const bcrA = isFinite(a.m.bcr) ? a.m.bcr : -Infinity;
        const bcrB = isFinite(b.m.bcr) ? b.m.bcr : -Infinity;
        return bcrB - bcrA;
      });

    list.forEach((entry, idx) => {
      const t = entry.t;
      const m = entry.m;
      const adopt = clamp(adoptMul, 0, 1);
      const area = asNum(t.area, 0);
      const annualCostPerHa = asNum(t.materialsCost, 0) + asNum(t.servicesCost, 0) + asNum(t.labourCost, 0);
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
    const years = asNum(model.time.years, 10);
    const control = model.treatments.find(t => t.isControl);
    let controlMetrics = null;
    if (control) controlMetrics = computeSingleTreatmentMetrics(control, rate, years, adoptMul, risk);

    const combined = { area: 0, deltas: {}, capitalCost: 0, annualCost: 0 };
    model.outputs.forEach(o => (combined.deltas[o.id] = 0));

    model.treatments.forEach(t => {
      if (t.isControl) return;
      combined.capitalCost += asNum(t.capitalCost, 0);
      const area = asNum(t.area, 0);
      const adopt = clamp(adoptMul, 0, 1);
      combined.area += area;

      const annualCostPerHa = asNum(t.materialsCost, 0) + asNum(t.servicesCost, 0) + asNum(t.labourCost, 0);
      combined.annualCost += annualCostPerHa * area;

      model.outputs.forEach(o => {
        combined.deltas[o.id] += asNum(t.deltas?.[o.id], 0) * (area * adopt);
      });
    });

    let treatMetrics = null;
    if (combined.capitalCost || combined.annualCost || Object.values(combined.deltas).some(v => v !== 0)) {
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
        temp.deltas[o.id] = combined.area ? combined.deltas[o.id] / combined.area : 0;
      });
      const annualCostPerHa = combined.area ? combined.annualCost / combined.area : 0;
      temp.labourCost = annualCostPerHa;
      treatMetrics = computeSingleTreatmentMetrics(temp, rate, years, 1, risk);
    }

    return { controlMetrics, treatMetrics };
  }

  function renderDepreciationSummary() {
    const root = $("#depSummary");
    if (!root) return;
    root.innerHTML = "";

    const N = asNum(model.time.years, 10);
    const baseYear = asNum(model.time.startYear, new Date().getFullYear());
    const totalPerYear = new Array(N + 1).fill(0);
    const rows = [];

    model.otherCosts.forEach(c => {
      if (c.category !== "Capital") return;
      const method = c.depMethod || "none";
      const cost = asNum(c.capital, 0);
      if (method === "none" || !cost) return;

      const life = Math.max(1, asNum(c.depLife, 5));
      const rate = asNum(c.depRate, 30);
      const startIndex = asNum(c.year, baseYear) - baseYear;

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
      p.textContent = "No capital items with depreciation configured. Set the depreciation method on capital costs to see a schedule.";
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
        ${rows.map(r => `
          <tr>
            <td>${esc(r.label)}</td>
            <td>${esc(r.method)}</td>
            <td>${r.life}</td>
            <td>${r.rate || ""}</td>
            <td>${money(r.firstDep)}</td>
          </tr>
        `).join("")}
      </tbody>
    `;
    root.appendChild(tbl);
  }

  // =========================
  // MAIN CALC / REPORT
  // =========================
  function renderTimeProjections(benefitByYear, costByYear, rate) {
    const tblBody = $("#timeProjectionTable tbody");
    if (!tblBody) return;
    tblBody.innerHTML = "";

    const maxYears = asNum(model.time.years, 10);
    const npvSeries = [];
    const usedHorizons = [];

    horizons.forEach(H => {
      const h = Math.min(H, maxYears);
      if (h <= 0) return;
      const b = benefitByYear.slice(0, h + 1);
      const c = costByYear.slice(0, h + 1);

      const useSchedule = !!model.time.useDiscountSchedule || !!document.getElementById("useDiscSchedule")?.checked;
      const pvB = useSchedule ? presentValueWithSchedule(b, "base") : presentValue(b, rate);
      const pvC = useSchedule ? presentValueWithSchedule(c, "base") : presentValue(c, rate);

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
    const rate = asNum(model.time.discBase, 0);
    const adoptMul = asNum(model.adoption.base, 0);
    const risk = asNum(model.risk.base, 0);

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

    const { controlMetrics, treatMetrics } = computeControlAndTreatmentGroupMetrics(rate, adoptMul, risk);

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
      ["#pvBenefitsControl","#pvCostsControl","#npvControl","#bcrControl","#irrControl","#roiControl","#paybackControl","#gmControl"].forEach(sel => setVal(sel, "No control selected"));
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
      ["#pvBenefitsTreat","#pvCostsTreat","#npvTreat","#bcrTreat","#irrTreat","#roiTreat","#paybackTreat","#gmTreat"].forEach(sel => setVal(sel, "n/a"));
    }

    renderTimeProjections(all.benefitByYear, all.costByYear, rate);
    renderDepreciationSummary();

    setVal("#simBcrTargetLabel", model.sim.targetBCR);
  }

  // =========================
  // MONTE CARLO (CHUNKED + PROGRESS)
  // =========================
  async function runSimulation() {
    const status = $("#simStatus");
    const progressEl = $("#simProgress");
    const progressText = $("#simProgressText");

    if (status) status.textContent = "Running Monte Carlo simulation ...";
    if (progressEl) progressEl.value = 0;
    if (progressText) progressText.textContent = "";

    await new Promise(r => setTimeout(r));

    const N = Math.max(100, asNum(model.sim.n, 1000));
    const seed = model.sim.seed;
    const rand = rng(seed ?? undefined);

    const discLow = asNum(model.time.discLow, 0);
    const discBase = asNum(model.time.discBase, 0);
    const discHigh = asNum(model.time.discHigh, 0);
    const adoptLow = asNum(model.adoption.low, 0);
    const adoptBase = asNum(model.adoption.base, 0);
    const adoptHigh = asNum(model.adoption.high, 1);
    const riskLow = asNum(model.risk.low, 0);
    const riskBase = asNum(model.risk.base, 0);
    const riskHigh = asNum(model.risk.high, 1);

    const npvs = new Array(N);
    const bcrs = new Array(N);
    const details = [];
    const varPct = (asNum(model.sim.variationPct, 0) || 0) / 100;

    // Cache base values for shocks, then restore
    const baseOutValues = model.outputs.map(o => asNum(o.value, 0));
    const baseTreatCosts = model.treatments.map(t => ({
      materials: asNum(t.materialsCost, 0),
      services: asNum(t.servicesCost, 0),
      labour: asNum(t.labourCost, 0)
    }));
    const baseOtherCosts = model.otherCosts.map(c => ({
      annual: asNum(c.annual, 0),
      capital: asNum(c.capital, 0)
    }));

    const chunkSize = 250;
    for (let start = 0; start < N; start += chunkSize) {
      const end = Math.min(N, start + chunkSize);

      for (let i = start; i < end; i++) {
        const r1 = rand();
        const r2 = rand();
        const r3 = rand();
        const disc = triangular(r1, discLow, discBase, discHigh);
        const adoptMul = clamp(triangular(r2, adoptLow, adoptBase, adoptHigh), 0, 1);
        const risk = clamp(triangular(r3, riskLow, riskBase, riskHigh), 0, 1);

        const shockOutputs = model.sim.varyOutputs ? 1 + (rand() * 2 * varPct - varPct) : 1;
        const shockTreatCosts = model.sim.varyTreatCosts ? 1 + (rand() * 2 * varPct - varPct) : 1;
        const shockInputCosts = model.sim.varyInputCosts ? 1 + (rand() * 2 * varPct - varPct) : 1;

        if (model.sim.varyOutputs) {
          model.outputs.forEach((o, idx) => { o.value = baseOutValues[idx] * shockOutputs; });
        }
        if (model.sim.varyTreatCosts) {
          model.treatments.forEach((t, idx) => {
            const base = baseTreatCosts[idx];
            t.materialsCost = base.materials * shockTreatCosts;
            t.servicesCost = base.services * shockTreatCosts;
            t.labourCost = base.labour * shockTreatCosts;
          });
        }
        if (model.sim.varyInputCosts) {
          model.otherCosts.forEach((c, idx) => {
            const base = baseOtherCosts[idx];
            c.annual = base.annual * shockInputCosts;
            c.capital = base.capital * shockInputCosts;
          });
        }

        const all = computeAll(disc, adoptMul, risk, model.sim.bcrMode);
        npvs[i] = all.npv;
        bcrs[i] = all.bcr;
        details.push({ discountRatePct: disc, adoptionMultiplier: adoptMul, riskMultiplier: risk, npv: all.npv, bcr: all.bcr });
      }

      // Update progress + yield to UI thread
      if (progressEl) progressEl.value = Math.round((end / N) * 100);
      if (progressText) progressText.textContent = `${end.toLocaleString()} / ${N.toLocaleString()} runs`;
      await new Promise(r => setTimeout(r, 0));
    }

    // Restore base values
    model.outputs.forEach((o, idx) => { o.value = baseOutValues[idx]; });
    model.treatments.forEach((t, idx) => {
      const base = baseTreatCosts[idx];
      t.materialsCost = base.materials;
      t.servicesCost = base.services;
      t.labourCost = base.labour;
    });
    model.otherCosts.forEach((c, idx) => {
      const base = baseOtherCosts[idx];
      c.annual = base.annual;
      c.capital = base.capital;
    });

    function summaryStats(arr) {
      const clean = arr.filter(v => isFinite(v));
      if (!clean.length) return { min: NaN, max: NaN, mean: NaN, median: NaN, probPos: NaN };
      clean.sort((a, b) => a - b);
      const min = clean[0];
      const max = clean[clean.length - 1];
      const mean = clean.reduce((a, b) => a + b, 0) / clean.length;
      const mid = Math.floor(clean.length / 2);
      const median = clean.length % 2 === 0 ? (clean[mid - 1] + clean[mid]) / 2 : clean[mid];
      const probPos = clean.filter(v => v > 0).length / clean.length;
      return { min, max, mean, median, probPos };
    }

    const npvStats = summaryStats(npvs);
    const bcrStats = summaryStats(bcrs);

    const target = asNum(model.sim.targetBCR, 0);
    const probBcrGt1 = bcrs.filter(v => isFinite(v) && v > 1).length / (bcrs.length || 1);
    const probBcrGtTarget = bcrs.filter(v => isFinite(v) && v > target).length / (bcrs.length || 1);

    setVal("#simNpvMin", money(npvStats.min));
    setVal("#simNpvMax", money(npvStats.max));
    setVal("#simNpvMean", money(npvStats.mean));
    setVal("#simNpvMedian", money(npvStats.median));
    setVal("#simNpvProb", isFinite(npvStats.probPos) ? percent(npvStats.probPos * 100) : "n/a");

    setVal("#simBcrMin", isFinite(bcrStats.min) ? fmt(bcrStats.min) : "n/a");
    setVal("#simBcrMax", isFinite(bcrStats.max) ? fmt(bcrStats.max) : "n/a");
    setVal("#simBcrMean", isFinite(bcrStats.mean) ? fmt(bcrStats.mean) : "n/a");
    setVal("#simBcrMedian", isFinite(bcrStats.median) ? fmt(bcrStats.median) : "n/a");
    setVal("#simBcrProb1", percent(probBcrGt1 * 100));
    setVal("#simBcrProbTarget", percent(probBcrGtTarget * 100));

    drawHistogram("histNpv", npvs, 20);
    drawHistogram("histBcr", bcrs.filter(v => isFinite(v)), 20);

    model.sim.results = { npv: npvs, bcr: bcrs };
    model.sim.details = details;

    if (status) status.textContent = `Simulation complete for ${N.toLocaleString()} runs. Results use adoption and risk ranges from the Settings tab.`;
    autosaveDebounced();
    showToast("Simulation complete.");
  }

  // =========================
  // CHART HELPERS
  // =========================
  function drawTimeSeries(canvasId, xs, ys) {
    const canvas = document.getElementById(canvasId);
    if (!canvas || !canvas.getContext) return;
    const ctx = canvas.getContext("2d");

    const w = canvas.width;
    const h = canvas.height;
    ctx.clearRect(0, 0, w, h);
    if (!xs || !xs.length || !ys || !ys.length) return;

    const minX = Math.min(...xs);
    const maxX = Math.max(...xs);
    const minY = Math.min(...ys);
    const maxY = Math.max(...ys);

    const pad = 30;
    const plotW = w - pad * 2;
    const plotH = h - pad * 2;

    const xScale = v => pad + ((v - minX) / (maxX - minX || 1)) * plotW;
    const yScale = v => pad + plotH - ((v - minY) / (maxY - minY || 1)) * plotH;

    ctx.strokeStyle = "#cccccc";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(pad, pad);
    ctx.lineTo(pad, pad + plotH);
    ctx.lineTo(pad + plotW, pad + plotH);
    ctx.stroke();

    ctx.strokeStyle = "#1D4F91";
    ctx.lineWidth = 2;
    ctx.beginPath();
    xs.forEach((x, i) => {
      const px = xScale(x);
      const py = yScale(ys[i]);
      if (i === 0) ctx.moveTo(px, py);
      else ctx.lineTo(px, py);
    });
    ctx.stroke();
  }

  function drawHistogram(canvasId, data, bins = 20) {
    const canvas = document.getElementById(canvasId);
    if (!canvas || !canvas.getContext) return;
    const ctx = canvas.getContext("2d");

    const w = canvas.width;
    const h = canvas.height;
    ctx.clearRect(0, 0, w, h);

    const clean = data.filter(v => isFinite(v));
    if (!clean.length) return;

    const min = Math.min(...clean);
    const max = Math.max(...clean);
    if (min === max) return;

    const counts = new Array(bins).fill(0);
    const binWidth = (max - min) / bins;
    clean.forEach(v => {
      let idx = Math.floor((v - min) / binWidth);
      if (idx < 0) idx = 0;
      if (idx >= bins) idx = bins - 1;
      counts[idx]++;
    });

    const pad = 20;
    const plotW = w - pad * 2;
    const plotH = h - pad * 2;
    const maxCount = Math.max(...counts);
    const barW = plotW / bins;

    ctx.strokeStyle = "#cccccc";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(pad, pad);
    ctx.lineTo(pad, pad + plotH);
    ctx.lineTo(pad + plotW, pad + plotH);
    ctx.stroke();

    ctx.fillStyle = "#1D4F91";
    counts.forEach((c, i) => {
      const x = pad + i * barW;
      const barH = (c / maxCount) * plotH;
      const y = pad + plotH - barH;
      ctx.fillRect(x + 1, y, barW - 2, barH);
    });
  }

  // =========================
  // EXPORT HELPERS (CSV/XLSX/PDF)
  // =========================
  function downloadFile(filename, content, mime) {
    const blob = new Blob([content], { type: mime || "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 0);
  }

  function csvCell(v) {
    if (v == null) return "";
    const s = String(v);
    if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  }

  function exportAllCsv() {
    const rate = asNum(model.time.discBase, 0);
    const adoptMul = asNum(model.adoption.base, 0);
    const risk = asNum(model.risk.base, 0);
    const all = computeAll(rate, adoptMul, risk, model.sim.bcrMode);

    const rows = [];
    rows.push(["Section", "Metric", "Value"]);
    rows.push(["Whole project", "PV benefits", all.pvBenefits]);
    rows.push(["Whole project", "PV costs", all.pvCosts]);
    rows.push(["Whole project", "NPV", all.npv]);
    rows.push(["Whole project", "BCR", all.bcr]);
    rows.push(["Whole project", "IRR (%)", all.irrVal]);
    rows.push(["Whole project", "MIRR (%)", all.mirrVal]);
    rows.push(["Whole project", "ROI (%)", all.roi]);
    rows.push(["Whole project", "Annual gross margin", all.annualGM]);
    rows.push(["Whole project", "Gross profit margin (%)", all.profitMargin]);
    rows.push(["Whole project", "Payback (years)", all.paybackYears]);

    model.treatments.forEach(t => {
      const m = computeSingleTreatmentMetrics(t, rate, asNum(model.time.years, 10), adoptMul, risk);
      rows.push(["Treatment", t.name + (t.isControl ? " (Control)" : ""), ""]);
      rows.push(["", "PV benefits", m.pvBen]);
      rows.push(["", "PV costs", m.pvCost]);
      rows.push(["", "NPV", m.npv]);
      rows.push(["", "BCR", m.bcr]);
      rows.push(["", "IRR (%)", m.irrVal]);
      rows.push(["", "ROI (%)", m.roi]);
    });

    const csv = rows.map(r => r.map(csvCell).join(",")).join("\r\n");
    downloadFile(slug(model.project.name || "project") + "_summary.csv", csv, "text/csv");
    showToast("CSV summary downloaded.");
  }

  function exportAllXlsx() {
    if (typeof XLSX === "undefined") {
      alert("The SheetJS XLSX library is required for Excel export.");
      return;
    }

    const rate = asNum(model.time.discBase, 0);
    const adoptMul = asNum(model.adoption.base, 0);
    const risk = asNum(model.risk.base, 0);
    const all = computeAll(rate, adoptMul, risk, model.sim.bcrMode);

    const wb = XLSX.utils.book_new();

    const summaryAoA = [
      ["Farming CBA Decision Aid", "", "", "App version", APP_VERSION],
      ["Project name", model.project.name],
      ["Organisation", model.project.organisation],
      ["Last updated", model.project.lastUpdated],
      [""],
      ["Base case settings"],
      ["Discount rate (base %)", rate],
      ["Adoption multiplier (base)", adoptMul],
      ["Risk multiplier (base)", risk],
      ["Analysis years", asNum(model.time.years, 10)],
      [""],
      ["Base case results"],
      ["PV benefits", all.pvBenefits],
      ["PV costs", all.pvCosts],
      ["NPV", all.npv],
      ["BCR", all.bcr],
      ["IRR (%)", all.irrVal],
      ["MIRR (%)", all.mirrVal],
      ["ROI (%)", all.roi],
      ["Annual gross margin", all.annualGM],
      ["Gross profit margin (%)", all.profitMargin],
      ["Payback (years)", all.paybackYears]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryAoA), "Summary");

    const treatmentsRows = [["Treatment", "Control", "Area_ha", "Materials_$perHa", "Services_$perHa", "Labour_$perHa", "Capital_$Y0", "Constrained", "PV_Benefits", "PV_Costs", "NPV", "BCR", "IRR_pct", "ROI_pct"]];
    model.treatments.forEach(t => {
      const m = computeSingleTreatmentMetrics(t, rate, asNum(model.time.years, 10), adoptMul, risk);
      treatmentsRows.push([
        t.name,
        t.isControl ? 1 : 0,
        asNum(t.area, 0),
        asNum(t.materialsCost, 0),
        asNum(t.servicesCost, 0),
        asNum(t.labourCost, 0),
        asNum(t.capitalCost, 0),
        t.constrained ? 1 : 0,
        m.pvBen,
        m.pvCost,
        m.npv,
        m.bcr,
        m.irrVal,
        m.roi
      ]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(treatmentsRows), "Treatments");

    const cashflowRows = [["YearIndex", "Benefits", "Costs", "NetCashflow"]];
    all.benefitByYear.forEach((b, i) => {
      cashflowRows.push([i, b, all.costByYear[i], all.cf[i]]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cashflowRows), "Cashflows");

    if (model.sim?.details?.length) {
      const simRows = [["Run", "DiscountRatePct", "AdoptionMultiplier", "RiskMultiplier", "NPV", "BCR"]];
      model.sim.details.slice(0, 200000).forEach((d, i) => {
        simRows.push([i + 1, d.discountRatePct, d.adoptionMultiplier, d.riskMultiplier, d.npv, d.bcr]);
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(simRows), "Simulation");
    }

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile(slug(model.project.name || "project") + "_export.xlsx", wbout, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Excel export downloaded.");
  }

  function exportPdf() {
    window.print();
  }

  // =========================
  // EXCEL (TEMPLATE / IMPORT)
  // =========================
  function buildFabaBeanSheet() {
    const header = ["Amendment", "Yield t/ha", "Pre sowing Labour", "Treatment Input Cost Only /Ha"];
    const dataRows = RAW_PLOTS.map(r => [r.Amendment ?? "", r["Yield t/ha"] ?? "", r["Pre sowing Labour"] ?? "", r["Treatment Input Cost Only /Ha"] ?? ""]);
    return [header, ...dataRows];
  }

  function downloadExcelTemplate() {
    if (typeof XLSX === "undefined") {
      alert("The SheetJS XLSX library is required for Excel export.");
      return;
    }
    const wb = XLSX.utils.book_new();

    const readme = XLSX.utils.aoa_to_sheet([
      ["Farming CBA Decision Aid – Excel template"],
      [""],
      ["Sheets"],
      ["FabaBeanRaw", "Default 2022 faba bean raw plot dataset used to calibrate treatments."],
      ["Project", "Optional project metadata sheet (name, summary, goal, organisation)."],
      [""],
      ["You can overwrite the FabaBeanRaw sheet with your own plot-level data. Keep at least the columns:"],
      ["Amendment, Yield t/ha, Pre sowing Labour, Treatment Input Cost Only /Ha."],
      [""],
      ["Import workflow"],
      ["1) Parse Excel in the tool."],
      ["2) Commit Excel to model to recalibrate treatments and refresh outputs."],
      [""],
      ["Notes"],
      ["This template is compatible with the built-in trial calibration logic."]
    ]);
    XLSX.utils.book_append_sheet(wb, readme, "ReadMe");

    const fabaAoA = buildFabaBeanSheet();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(fabaAoA), "FabaBeanRaw");

    const projectSheet = XLSX.utils.aoa_to_sheet([
      ["Project name", model.project.name],
      ["Summary", model.project.summary],
      ["Goal", model.project.goal],
      ["Organisation", model.project.organisation]
    ]);
    XLSX.utils.book_append_sheet(wb, projectSheet, "Project");

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile("faba_bean_cba_template.xlsx", wbout, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Excel template with faba bean dataset downloaded.");
  }

  function downloadSampleDataset() {
    if (typeof XLSX === "undefined") {
      alert("The SheetJS XLSX library is required for Excel export.");
      return;
    }
    const wb = XLSX.utils.book_new();

    const projectSheet = XLSX.utils.aoa_to_sheet([
      ["Project name", model.project.name],
      ["Summary", model.project.summary],
      ["Goal", model.project.goal],
      ["Organisation", model.project.organisation]
    ]);
    XLSX.utils.book_append_sheet(wb, projectSheet, "Project");

    const fabaAoA = buildFabaBeanSheet();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(fabaAoA), "FabaBeanRaw");

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile("faba_bean_cba_sample.xlsx", wbout, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Sample Excel workbook with faba bean dataset downloaded.");
  }

  function handleParseExcel() {
    if (typeof XLSX === "undefined") {
      alert("The SheetJS XLSX library is required for Excel import.");
      return;
    }
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".xlsx,.xlsm,.xlsb";
    input.style.display = "none";
    document.body.appendChild(input);

    input.addEventListener("change", async e => {
      const file = e.target.files && e.target.files[0];
      document.body.removeChild(input);
      if (!file) return;

      safeRun("ParseExcel", async () => {
        const buffer = await file.arrayBuffer();
        const wb = XLSX.read(buffer, { type: "array" });

        const fabaName = wb.SheetNames.find(n => FABABEAN_SHEET_NAMES.includes(n)) || wb.SheetNames[0];
        const fabaSheet = wb.Sheets[fabaName];
        const rows = XLSX.utils.sheet_to_json(fabaSheet, { defval: null });

        // Optional Project sheet
        let projectMeta = null;
        const projName = wb.SheetNames.find(n => String(n).toLowerCase() === "project");
        if (projName) {
          const ps = wb.Sheets[projName];
          const projAoA = XLSX.utils.sheet_to_json(ps, { header: 1, defval: "" });
          projectMeta = projAoA;
        }

        parsedExcel = { rawPlots: rows, projectMeta };
        showToast(`Excel parsed. Using sheet "${fabaName}" as faba bean raw data.`);
      });
    });

    input.click();
  }

  function commitExcelToModel() {
    if (!parsedExcel || !parsedExcel.rawPlots || !parsedExcel.rawPlots.length) {
      alert("No parsed Excel data is available. Please parse an Excel file first.");
      return;
    }

    safeRun("CommitExcel", () => {
      // Optional: apply project metadata if present
      if (parsedExcel.projectMeta && Array.isArray(parsedExcel.projectMeta)) {
        parsedExcel.projectMeta.forEach(row => {
          const k = String(row?.[0] || "").toLowerCase().trim();
          const v = String(row?.[1] || "").trim();
          if (!k) return;
          if (k.includes("project name")) model.project.name = v || model.project.name;
          if (k === "summary") model.project.summary = v || model.project.summary;
          if (k === "goal") model.project.goal = v || model.project.goal;
          if (k.includes("organisation")) model.project.organisation = v || model.project.organisation;
        });
      }

      RAW_PLOTS.length = 0;
      parsedExcel.rawPlots.forEach(r => RAW_PLOTS.push(r));
      applyTrialTreatmentsIfAvailable();
      renderAll();
      setBasicsFieldsFromModel();
      calcAndRender();
      autosaveDebounced();
      showToast("Excel data applied. Treatments recalibrated from imported faba bean dataset.");
    });
  }

  // =========================
  // COPILOT PAYLOAD
  // =========================
  function buildCopilotPayload() {
    const rate = asNum(model.time.discBase, 0);
    const adoptMul = asNum(model.adoption.base, 0);
    const risk = asNum(model.risk.base, 0);
    const all = computeAll(rate, adoptMul, risk, model.sim.bcrMode);

    const treatmentsRanked = [...model.treatments]
      .map(t => {
        const m = computeSingleTreatmentMetrics(t, rate, asNum(model.time.years, 10), adoptMul, risk);
        return {
          name: t.name,
          isControl: !!t.isControl,
          areaHa: asNum(t.area, 0),
          annualGrossMargin: m.gm,
          pvBenefits: m.pvBen,
          pvCosts: m.pvCost,
          bcr: m.bcr,
          npv: m.npv,
          irrPct: m.irrVal,
          roiPct: m.roi
        };
      })
      .sort((a, b) => (isFinite(b.bcr) ? b.bcr : -Infinity) - (isFinite(a.bcr) ? a.bcr : -Infinity));

    return {
      toolName: "Farming CBA Decision Aid – Faba bean trial",
      toolVersion: APP_VERSION,
      project: {
        name: model.project.name,
        summary: model.project.summary,
        goal: model.project.goal,
        organisation: model.project.organisation,
        stakeholders: model.project.stakeholders
      },
      timeSettings: {
        startYear: model.time.startYear,
        analysisYears: model.time.years,
        discountRateBasePct: model.time.discBase,
        discountRateLowPct: model.time.discLow,
        discountRateHighPct: model.time.discHigh,
        useDiscountSchedule: !!model.time.useDiscountSchedule,
        discountSchedule: model.time.discountSchedule
      },
      riskAndAdoption: {
        adoptionBase: model.adoption.base,
        adoptionLow: model.adoption.low,
        adoptionHigh: model.adoption.high,
        riskBase: model.risk.base,
        riskLow: model.risk.low,
        riskHigh: model.risk.high
      },
      baseCaseResults: {
        pvBenefits: all.pvBenefits,
        pvCosts: all.pvCosts,
        npv: all.npv,
        bcr: all.bcr,
        irrPct: all.irrVal,
        mirrPct: all.mirrVal,
        roiPct: all.roi,
        annualGrossMargin: all.annualGM,
        grossProfitMarginPct: all.profitMargin,
        paybackYears: all.paybackYears
      },
      treatmentsRanked,
      fabaBeanDataset: {
        hasRawData: RAW_PLOTS.length > 0,
        sheetNamesExpected: FABABEAN_SHEET_NAMES,
        numberOfRawRows: RAW_PLOTS.length
      },
      notesForAssistant: [
        "Write a clear, non-technical policy brief for decision makers.",
        "Explain what the faba bean trial is testing, what the main treatments are, and how yield and costs differ from control.",
        "Use the NPV, BCR, IRR, and payback to comment on whether the package of treatments is economically attractive.",
        "Highlight uncertainty and sensitivity to adoption, risk, and discount rates.",
        "Explain what information would help firm up the estimates (for example, more seasons of data or more sites)."
      ]
    };
  }

  function handleOpenCopilotClick() {
    const payload = buildCopilotPayload();
    const json = JSON.stringify(payload, null, 2);
    const preview = $("#copilotPreview");
    if (preview) preview.value = json;

    const ok = navigator.clipboard && navigator.clipboard.writeText;
    if (ok) {
      navigator.clipboard
        .writeText(json)
        .then(() => showToast("Scenario summary copied to clipboard for your policy brief."))
        .catch(() => showToast("Unable to copy automatically. Please copy the JSON from the preview box."));
    } else {
      showToast("Copying not supported. Please copy the JSON from the preview box.");
    }
  }

  // =========================
  // TOOLTIP FOR INDICATORS
  // =========================
  function applyIndicatorTooltips() {
    const map = {
      pvBenefits: "Present value of all benefits, combining treatment output gains and additional benefits, discounted at the base rate over the analysis period.",
      pvCosts: "Present value of all treatment and project costs, including capital and operating items, discounted at the base rate.",
      npv: "Net present value, equal to present value of benefits minus present value of costs. Positive values indicate that benefits exceed costs.",
      bcr: "Benefit–cost ratio, equal to present value of benefits divided by present value of costs, using the chosen denominator (all or constrained costs).",
      irr: "Internal rate of return, the discount rate at which the net present value of the cash flow stream is zero.",
      mirr: "Modified internal rate of return, using the MIRR finance and reinvestment rates set in the Time and risk settings tab.",
      roi: "Return on investment, calculated as NPV divided by present value of costs, expressed as a percentage.",
      grossMargin: "Average annual gross margin, equal to annual benefits minus annual operating costs for the whole project.",
      profitMargin: "Gross profit margin, equal to annual gross margin as a percentage of annual benefits.",
      payback: "Discounted payback period, the first year when the cumulative discounted cash flow becomes non-negative.",
      simNpvMin: "Minimum net present value across all simulation runs.",
      simNpvMax: "Maximum net present value across all simulation runs.",
      simNpvMean: "Mean net present value across all simulation runs.",
      simNpvMedian: "Median net present value across all simulation runs.",
      simNpvProb: "Estimated probability that net present value is greater than zero, based on simulation runs.",
      simBcrMin: "Minimum benefit–cost ratio across all simulation runs.",
      simBcrMax: "Maximum benefit–cost ratio across all simulation runs.",
      simBcrMean: "Mean benefit–cost ratio across all simulation runs.",
      simBcrMedian: "Median benefit–cost ratio across all simulation runs.",
      simBcrProb1: "Estimated probability that the benefit–cost ratio exceeds 1.0.",
      simBcrProbTarget: "Estimated probability that the benefit–cost ratio exceeds the user-defined target."
    };

    Object.keys(map).forEach(id => {
      const el = document.getElementById(id);
      if (el) el.title = map[id];
    });
  }

  // =========================
  // INIT
  // =========================
  document.addEventListener("DOMContentLoaded", () => {
    safeRun("Init", () => {
      // Load persisted state if available
      const loaded = tryAutoload();
      hardenModel();

      // Ensure trial-calibrated defaults if fresh or missing
      if (!loaded) applyTrialTreatmentsIfAvailable();

      initTabs();
      bindBasics();
      initAddButtons();
      renderAll();
      setBasicsFieldsFromModel();
      calcAndRender();
      applyIndicatorTooltips();

      // Autosave on unload
      window.addEventListener("beforeunload", () => {
        tryAutosave();
      });

      showToast(loaded ? "Project restored from autosave." : "Tool initialised with default trial dataset.");
    });
  });
})();
