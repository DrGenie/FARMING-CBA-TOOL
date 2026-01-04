// Farming CBA Decision Tool 2 — Newcastle Business School
// Commercial-grade, control-centric CBA tool with Excel-first workflow.
// Default dataset is loaded THROUGH THE SAME parser/validator pipeline as an uploaded Excel file.
// All tabs/buttons work. Results open by default with a sticky, scrollable Comparison-to-Control table,
// leaderboard, exports (Excel/CSV/PDF print), robust missing-value handling ("?" treated as missing),
// and non-prescriptive AI prompt generation based only on computed results.

(() => {
  "use strict";

  // =========================
  // 0) HARD REQUIREMENTS
  // =========================
  // - Default dataset must behave identically to an uploaded Excel file: same parser, same validator, same commit.
  // - Missing or "?" values are treated as missing (NaN), never zero; flagged in validation; excluded from calculations.
  // - Results view is control-centric and opens by default; comparison table visible without scrolling.
  // - Capital cost is year-0 (not discounted); annual variable costs are discounted.
  // - Control is a single baseline; if multiple controls are flagged, we keep the first and flag validation.
  // - Tool name used consistently: "Farming CBA Decision Tool 2".

  // =========================
  // 1) CONSTANTS / DEFAULTS
  // =========================
  const TOOL_NAME = "Farming CBA Decision Tool 2";
  const ORG_NAME = "Newcastle Business School, The University of Newcastle";

  const DEFAULT_DISCOUNT_SCHEDULE = [
    { label: "2025-2034", from: 2025, to: 2034, low: 2, base: 4, high: 6 },
    { label: "2035-2044", from: 2035, to: 2044, low: 4, base: 7, high: 10 },
    { label: "2045-2054", from: 2045, to: 2054, low: 4, base: 7, high: 10 },
    { label: "2055-2064", from: 2055, to: 2064, low: 3, base: 6, high: 9 },
    { label: "2065-2074", from: 2065, to: 2074, low: 2, base: 5, high: 8 }
  ];

  const horizons = [5, 10, 15, 20, 25];

  // Expected sheet names for Excel-first workflow (single source of truth)
  const FABABEAN_SHEET_NAMES = ["FabaBeanRaw", "FabaBeansRaw", "FabaBean", "FabaBeans"];

  // Raw dataset schema columns (must be preserved; blanks and "?" preserved)
  const RAW_SCHEMA_COLUMNS = [
    "Amendment",
    "Yield t/ha",
    "Pre sowing Labour",
    "Treatment Input Cost Only /Ha",
    // Optional columns still preserved if present in Excel:
    "Amendment Labour",
    "Sowing Labour",
    "Herbicide Labour",
    "Herbicide Labour 2",
    "Herbicide Labour 3",
    "Harvesting Labour",
    "Harvesting Labour 2",
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

  // =========================
  // 2) DEFAULT DATASET (EMBEDDED)
  // =========================
  // This is your CURRENT embedded default dataset. It is loaded by building an in-memory workbook
  // and passing it through the SAME parser/validator/commit pipeline used for uploaded Excel files.
  // IMPORTANT: Values (including blanks/"?") are preserved EXACTLY as raw strings as provided.
  const DEFAULT_RAW_PLOTS = [
    { Amendment: "control", "Yield t/ha": 2.4, "Pre sowing Labour": 40, "Treatment Input Cost Only /Ha": 0 },
    { Amendment: "deep_om_cp1", "Yield t/ha": 3.1, "Pre sowing Labour": 55, "Treatment Input Cost Only /Ha": 16500 },
    {
      Amendment: "deep_om_cp1_plus_liq_gypsum_cht",
      "Yield t/ha": 3.2,
      "Pre sowing Labour": 56,
      "Treatment Input Cost Only /Ha": 16850
    },
    { Amendment: "deep_gypsum", "Yield t/ha": 2.9, "Pre sowing Labour": 50, "Treatment Input Cost Only /Ha": 500 },
    {
      Amendment: "deep_om_cp1_plus_pam",
      "Yield t/ha": 3.0,
      "Pre sowing Labour": 57,
      "Treatment Input Cost Only /Ha": 18000
    },
    {
      Amendment: "deep_om_cp1_plus_ccm",
      "Yield t/ha": 3.25,
      "Pre sowing Labour": 58,
      "Treatment Input Cost Only /Ha": 21225
    },
    {
      Amendment: "deep_ccm_only",
      "Yield t/ha": 2.95,
      "Pre sowing Labour": 52,
      "Treatment Input Cost Only /Ha": 3225
    },
    {
      Amendment: "deep_om_cp2_plus_gypsum",
      "Yield t/ha": 3.3,
      "Pre sowing Labour": 60,
      "Treatment Input Cost Only /Ha": 24000
    },
    {
      Amendment: "deep_liq_gypsum_cht",
      "Yield t/ha": 2.8,
      "Pre sowing Labour": 48,
      "Treatment Input Cost Only /Ha": 350
    },
    {
      Amendment: "surface_silicon",
      "Yield t/ha": 2.7,
      "Pre sowing Labour": 45,
      "Treatment Input Cost Only /Ha": 1000
    },
    {
      Amendment: "deep_liq_npks",
      "Yield t/ha": 3.0,
      "Pre sowing Labour": 53,
      "Treatment Input Cost Only /Ha": 2200
    },
    {
      Amendment: "deep_ripping_only",
      "Yield t/ha": 2.85,
      "Pre sowing Labour": 47,
      "Treatment Input Cost Only /Ha": 0
    }
  ];

  // =========================
  // 3) TRIAL CONFIG (CALIBRATION)
  // =========================
  // Capital assets (optional; still available)
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

  // Trial treatments to calibrate to model.treatments
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
      costs: {
        treatmentInputCostPerHa: 0,
        labourCostPerHa: 0,
        additionalOperatingCostPerHa: 0,
        capitalDepreciationPerHa: 0
      },
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

  // =========================
  // 4) MODEL (APP STATE)
  // =========================
  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

  const model = {
    project: {
      name: "Faba bean soil amendment trial",
      lead: "Project lead",
      analysts: "Farm economics team",
      team: "Trial team",
      organisation: ORG_NAME,
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
        "Identify soil amendment packages that deliver higher faba bean yields and acceptable returns after accounting for additional costs.",
      withProject:
        "Growers adopt high-performing amendment packages on trial farms and similar soils in the region.",
      withoutProject:
        "Growers continue baseline practice and do not access detailed economic evidence on soil amendments."
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
    outputsMeta: { systemType: "single", assumptions: "" },
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
    risk: { base: 0.15, low: 0.05, high: 0.3, tech: 0.05, nonCoop: 0.04, socio: 0.02, fin: 0.03, man: 0.02 },
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
    },
    // Data pipeline state
    dataPipeline: {
      lastLoadedSource: "default",
      lastLoadedSheet: "FabaBeanRaw",
      validation: { ok: true, issues: [], stats: {} },
      formulaInfo: { found: false, cells: [] },
      rawRowsPreserved: [] // exact raw rows (including "?" / blanks)
    }
  };

  function initTreatmentDeltas() {
    model.treatments.forEach(t => {
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

  // =========================
  // 5) UTILITIES
  // =========================
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

  const esc = s =>
    (s ?? "")
      .toString()
      .replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const annuityFactor = (N, rPct) => {
    const r = rPct / 100;
    return r === 0 ? N : (1 - Math.pow(1 + r, -N)) / r;
  };

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
    }, 2800);
  }

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

  // STRICT numeric parsing:
  // - preserves raw value elsewhere
  // - treats "" and "?" and null/undefined as missing => NaN
  // - strips $, commas
  function parseNumberStrict(value) {
    if (value === null || value === undefined) return NaN;
    const s = String(value).trim();
    if (!s || s === "?") return NaN;
    if (typeof value === "number") return Number.isFinite(value) ? value : NaN;
    const cleaned = s.replace(/[\$,]/g, "");
    const n = parseFloat(cleaned);
    return Number.isFinite(n) ? n : NaN;
  }

  function slugifyTreatmentName(name) {
    return String(name || "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  // =========================
  // 6) EXCEL-FIRST PIPELINE
  // =========================
  // One single pipeline for BOTH default and uploaded:
  //   Workbook -> choose sheet -> preserve raw rows -> validate -> commit -> recalc -> render.
  //
  // Formula detection:
  // - We scan chosen sheet cells for .f (formula) and record them.
  // - We use stored values (.v) that SheetJS provides; if a formula cell has no value, it becomes missing and flagged.

  function buildDefaultWorkbookFromEmbeddedData() {
    if (typeof XLSX === "undefined") return null;

    // Create workbook with ReadMe + FabaBeanRaw sheets
    const wb = XLSX.utils.book_new();

    const readmeAoA = [
      [TOOL_NAME + " — Default Excel Dataset (embedded)"],
      [""],
      ["Purpose"],
      ["This workbook mirrors the default dataset that loads automatically in the web tool."],
      ["You can edit the FabaBeanRaw sheet and upload it back into the tool. The tool will parse and validate it identically to the default."],
      [""],
      ["Important rules"],
      ["- Use '?' for unknown values. The tool treats '?' as missing, flags it, and excludes it from calculations that need it."],
      ["- Do not delete required columns. Extra columns are preserved if present."],
      [""],
      ["Sheets"],
      ["FabaBeanRaw", "Raw plot-level or treatment-level rows used to calibrate treatments."]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(readmeAoA), "ReadMe");

    // Build FabaBeanRaw AoA with schema-first header to match template structure
    const header = [...RAW_SCHEMA_COLUMNS];
    const rowsAoA = DEFAULT_RAW_PLOTS.map(r => header.map(h => (h in r ? r[h] : "")));
    const aoa = [header, ...rowsAoA];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "FabaBeanRaw");

    return wb;
  }

  function scanSheetForFormulas(sheet) {
    const out = { found: false, cells: [] };
    if (!sheet) return out;
    // Sheet cells are keys like "A1", "B2", plus "!ref"
    Object.keys(sheet).forEach(k => {
      if (k[0] === "!") return;
      const cell = sheet[k];
      if (cell && cell.f) {
        out.found = true;
        out.cells.push({ cell: k, formula: String(cell.f) });
      }
    });
    return out;
  }

  function chooseFabaSheetName(wb) {
    const match = wb.SheetNames.find(n => FABABEAN_SHEET_NAMES.includes(n));
    return match || wb.SheetNames[0];
  }

  function workbookToPreservedRows(wb, sheetName) {
    const sheet = wb.Sheets[sheetName];
    if (!sheet) return { rows: [], sheetName, formulaInfo: { found: false, cells: [] } };

    const formulaInfo = scanSheetForFormulas(sheet);

    // Preserve raw values: defval keeps blanks; raw:true keeps raw cell stored values; we also keep strings as-is.
    const rows = XLSX.utils.sheet_to_json(sheet, {
      defval: "",
      raw: true
    });

    // Ensure required columns exist (even if blank), without discarding any extra columns.
    const preserved = rows.map(r => {
      const rr = Object.assign({}, r);
      RAW_SCHEMA_COLUMNS.forEach(col => {
        if (!(col in rr)) rr[col] = "";
      });
      return rr;
    });

    return { rows: preserved, sheetName, formulaInfo };
  }

  function validateRawRows(rawRows) {
    const issues = [];
    const stats = {
      nRows: rawRows.length,
      missingYield: 0,
      missingAmendment: 0,
      missingLabour: 0,
      missingInputCost: 0
    };

    if (!rawRows.length) {
      issues.push({ level: "error", code: "NO_ROWS", message: "No rows found in the raw dataset sheet." });
      return { ok: false, issues, stats };
    }

    rawRows.forEach((r, idx) => {
      const rowNo = idx + 2; // + header row
      const amend = String(r["Amendment"] ?? "").trim();
      if (!amend) {
        stats.missingAmendment++;
        issues.push({
          level: "error",
          code: "MISSING_AMENDMENT",
          message: `Row ${rowNo}: "Amendment" is blank (required).`
        });
      }

      const y = parseNumberStrict(r["Yield t/ha"]);
      if (!isFinite(y)) {
        stats.missingYield++;
        issues.push({
          level: "warn",
          code: "MISSING_YIELD",
          message: `Row ${rowNo}: "Yield t/ha" is missing or non-numeric ("${String(r["Yield t/ha"] ?? "")}"). Treated as missing.`
        });
      }

      // Labour: sum across known columns; if all missing, warn
      let anyLab = false;
      let sumLab = 0;
      LABOUR_COLUMNS.forEach(c => {
        const v = parseNumberStrict(r[c]);
        if (isFinite(v)) {
          anyLab = true;
          sumLab += v;
        }
      });
      if (!anyLab) {
        stats.missingLabour++;
        issues.push({
          level: "warn",
          code: "MISSING_LABOUR",
          message: `Row ${rowNo}: labour columns are all missing/blank. Labour cost treated as missing for calibration.`
        });
      }

      const ic = parseNumberStrict(r["Treatment Input Cost Only /Ha"]);
      if (!isFinite(ic)) {
        stats.missingInputCost++;
        issues.push({
          level: "warn",
          code: "MISSING_INPUT_COST",
          message: `Row ${rowNo}: "Treatment Input Cost Only /Ha" is missing or non-numeric ("${String(
            r["Treatment Input Cost Only /Ha"] ?? ""
          )}"). Treated as missing.`
        });
      }
    });

    // Check for at least one control-like row in Amendment
    const hasControl = rawRows.some(r => String(r["Amendment"] || "").toLowerCase().includes("control"));
    if (!hasControl) {
      issues.push({
        level: "warn",
        code: "NO_CONTROL_ROW",
        message:
          'No row has "Amendment" containing the word "control". The tool will still calibrate, but uplift vs control may be unavailable.'
      });
    }

    const ok = !issues.some(x => x.level === "error");
    return { ok, issues, stats };
  }

  function computeTreatmentStatsFromRaw(rawRows) {
    // Group by Amendment; preserve missing as missing; do not treat missing as zero.
    const groups = new Map();

    rawRows.forEach(row => {
      const treatmentName = String(row["Amendment"] || "").trim();
      if (!treatmentName) return;

      let g = groups.get(treatmentName);
      if (!g) {
        g = { name: treatmentName, yieldVals: [], labourVals: [], opVals: [] };
        groups.set(treatmentName, g);
      }

      const y = parseNumberStrict(row["Yield t/ha"]);
      if (isFinite(y)) g.yieldVals.push(y);

      let labour = 0;
      let anyLab = false;
      LABOUR_COLUMNS.forEach(col => {
        const v = parseNumberStrict(row[col]);
        if (isFinite(v)) {
          labour += v;
          anyLab = true;
        }
      });
      if (anyLab) g.labourVals.push(labour);

      let op = 0;
      let anyOp = false;
      OPERATING_COLUMNS.forEach(col => {
        const v = parseNumberStrict(row[col]);
        if (isFinite(v)) {
          op += v;
          anyOp = true;
        }
      });
      if (anyOp) g.opVals.push(op);
    });

    const mean = arr => (arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : NaN);

    let controlMeanYield = NaN;
    for (const [name, g] of groups.entries()) {
      if (name.toLowerCase().includes("control")) {
        controlMeanYield = mean(g.yieldVals);
        break;
      }
    }

    const out = [];
    for (const [name, g] of groups.entries()) {
      const meanYield = mean(g.yieldVals);
      const meanLabour = mean(g.labourVals);
      const meanOperating = mean(g.opVals);

      out.push({
        id: slugifyTreatmentName(name),
        label: name,
        isControl: name.toLowerCase().includes("control"),
        meanYieldTHa: meanYield,
        labourCostPerHa: meanLabour,
        additionalOperatingCostPerHa: meanOperating,
        yield_uplift_vs_control_t_ha: isFinite(controlMeanYield) && isFinite(meanYield) ? meanYield - controlMeanYield : NaN
      });
    }
    return out;
  }

  function commitRawRowsToModel(rawRows, sourceLabel) {
    // 1) Store preserved rows in pipeline state (exact raw values)
    model.dataPipeline.rawRowsPreserved = rawRows.map(r => Object.assign({}, r));
    model.dataPipeline.lastLoadedSource = sourceLabel || "uploaded";

    // 2) Calibrate treatments from raw rows (stats)
    const stats = computeTreatmentStatsFromRaw(rawRows);
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

        // If missing in config, use stats; missing stats stay missing => treated as 0 only when applied to model
        costs.labourCostPerHa =
          typeof costs.labourCostPerHa === "number" && !isNaN(costs.labourCostPerHa) ? costs.labourCostPerHa : s.labourCostPerHa;

        costs.additionalOperatingCostPerHa =
          typeof costs.additionalOperatingCostPerHa === "number" && !isNaN(costs.additionalOperatingCostPerHa)
            ? costs.additionalOperatingCostPerHa
            : s.additionalOperatingCostPerHa;
      }

      return { id: cfg.id, label: cfg.label, category: cfg.category, controlFlag: !!cfg.controlFlag, agronomy: agr, costs, capitalAssets: cfg.capitalAssets || [] };
    });

    // 3) Write model.treatments
    model.treatments = merged.map(tt => {
      const materialsPerHa = Number(tt.costs.treatmentInputCostPerHa || 0);

      // Missing labour/operating in stats should not break: treat missing as 0 but VALIDATION flags missing.
      const labourPerHa = isFinite(tt.costs.labourCostPerHa) ? Number(tt.costs.labourCostPerHa) : 0;
      const operatingPerHa = isFinite(tt.costs.additionalOperatingCostPerHa) ? Number(tt.costs.additionalOperatingCostPerHa) : 0;

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

      if (yieldId && tt.agronomy) {
        const uplift = tt.agronomy.yield_uplift_vs_control_t_ha;
        if (typeof uplift === "number" && isFinite(uplift)) t.deltas[yieldId] = uplift;
        else t.deltas[yieldId] = 0; // missing uplift handled by validation; keep computations stable
      }
      return t;
    });

    initTreatmentDeltas();
  }

  async function loadWorkbookThroughPipeline({ wb, sourceLabel }) {
    const sheetName = chooseFabaSheetName(wb);
    const { rows, formulaInfo } = workbookToPreservedRows(wb, sheetName);
    const validation = validateRawRows(rows);

    model.dataPipeline.lastLoadedSheet = sheetName;
    model.dataPipeline.validation = validation;
    model.dataPipeline.formulaInfo = formulaInfo;

    // Update validation UI immediately
    renderValidationPanel();

    if (!validation.ok) {
      showToast("Data loaded, but there are validation errors. Fix the Excel sheet and re-upload.");
      // Still commit what we can if there are rows; do not break tool
      if (rows.length) commitRawRowsToModel(rows, sourceLabel);
      return;
    }

    commitRawRowsToModel(rows, sourceLabel);
    showToast(sourceLabel === "default" ? "Default dataset loaded via Excel pipeline." : "Excel uploaded and applied via Excel pipeline.");
  }

  // =========================
  // 7) CASHFLOWS / METRICS
  // =========================
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
        addOnce(yr, Number(b.annualAmount) || 0);
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

  function presentValue(series, ratePct) {
    let pv = 0;
    for (let t = 0; t < series.length; t++) pv += series[t] / Math.pow(1 + ratePct / 100, t);
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

      const annualCostPerHa = (Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0) + (Number(t.labourCost) || 0);
      const opCost = annualCostPerHa * area;
      const cap = Number(t.capitalCost) || 0; // year-0

      annualBenefit += benefit;
      treatAnnualCost += opCost;
      treatCapitalY0 += cap;

      if (t.constrained) {
        treatConstrAnnualCost += opCost;
        treatConstrCapitalY0 += cap;
      }
    });

    // Capital treated as year-0 (NOT discounted)
    costByYear[0] += treatCapitalY0;
    constrainedCostByYear[0] += treatConstrCapitalY0;

    // Annual operating costs
    for (let t = 1; t <= N; t++) {
      benefitByYear[t] += annualBenefit;
      costByYear[t] += treatAnnualCost;
      constrainedCostByYear[t] += treatConstrAnnualCost;
    }

    // Other costs (annual + capital)
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

    // Additional benefits
    const extra = additionalBenefitsSeries(N, baseYear, adoptMul, risk);
    for (let i = 0; i < extra.length; i++) benefitByYear[i] += extra[i];

    const cf = new Array(N + 1).fill(0).map((_, i) => benefitByYear[i] - costByYear[i]);
    const annualGM = annualBenefit - treatAnnualCost;
    return { benefitByYear, costByYear, constrainedCostByYear, cf, annualGM };
  }

  function computeAll(rate, adoptMul, risk, bcrMode) {
    const { benefitByYear, costByYear, constrainedCostByYear, cf, annualGM } = buildCashflows({ forRate: rate, adoptMul, risk });

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

    return { pvBenefits, pvCosts, pvCostsConstrained, npv, bcr, irrVal, mirrVal, roi, annualGM, profitMargin, paybackYears: pb, cf, benefitByYear, costByYear };
  }

  function computeSingleTreatmentMetrics(t, rate, years, adoptMul, risk) {
    let valuePerHa = 0;
    model.outputs.forEach(o => {
      valuePerHa += (Number(t.deltas[o.id]) || 0) * (Number(o.value) || 0);
    });

    const adopt = clamp(adoptMul, 0, 1);
    const area = Number(t.area) || 0;

    const annualBen = valuePerHa * area * (1 - clamp(risk, 0, 1)) * adopt;

    const annualCostPerHa = (Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0) + (Number(t.labourCost) || 0);
    const annualCost = annualCostPerHa * area;

    const cap = Number(t.capitalCost) || 0; // year-0
    const pvBen = annualBen * annuityFactor(years, rate);
    const pvCost = cap + annualCost * annuityFactor(years, rate); // cap not discounted

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

  // =========================
  // 8) CONTROL-CENTRIC RESULTS TABLE
  // =========================
  function getControlTreatment() {
    const controls = model.treatments.filter(t => !!t.isControl);
    if (!controls.length) return null;
    // Enforce single control baseline
    return controls[0];
  }

  function computeComparisonToControl(rate, adoptMul, risk) {
    const years = model.time.years;
    const control = getControlTreatment();
    const treatments = model.treatments.slice();

    const metricsById = new Map();
    treatments.forEach(t => {
      metricsById.set(t.id, computeSingleTreatmentMetrics(t, rate, years, adoptMul, risk));
    });

    const controlM = control ? metricsById.get(control.id) : null;

    // Ranking: by NPV (descending). If NPV missing => bottom.
    const ranked = treatments
      .map(t => {
        const m = metricsById.get(t.id);
        return { t, m, npvSort: isFinite(m?.npv) ? m.npv : -Infinity };
      })
      .sort((a, b) => b.npvSort - a.npvSort);

    const rankById = new Map();
    ranked.forEach((x, i) => rankById.set(x.t.id, i + 1));

    const indicators = [
      {
        key: "pvBen",
        label: "Present Value (PV) of Benefits",
        format: money,
        deltaPctMeaningful: true,
        tooltip: "Discounted value of benefits over the analysis period (including output gains and additional benefits)."
      },
      {
        key: "pvCost",
        label: "Present Value (PV) of Costs",
        format: money,
        deltaPctMeaningful: true,
        tooltip: "Discounted value of all costs. Capital is counted in year 0 (not discounted); annual costs are discounted."
      },
      {
        key: "npv",
        label: "Net Present Value (NPV)",
        format: money,
        deltaPctMeaningful: false,
        tooltip: "PV Benefits − PV Costs. Positive means benefits exceed costs, compared with zero for that option."
      },
      {
        key: "bcr",
        label: "Benefit–Cost Ratio (BCR)",
        format: v => (isFinite(v) ? fmt(v) : "n/a"),
        deltaPctMeaningful: true,
        tooltip: "PV Benefits ÷ PV Costs. Interpretable as dollars of benefit per dollar of cost (present value)."
      },
      {
        key: "roi",
        label: "Return on Investment (ROI)",
        format: v => (isFinite(v) ? percent(v) : "n/a"),
        deltaPctMeaningful: true,
        tooltip: "NPV ÷ PV Costs, expressed as a percentage (net gain per dollar of PV cost)."
      },
      {
        key: "rank",
        label: "Ranking (by NPV)",
        format: v => (isFinite(v) ? String(v) : "n/a"),
        deltaPctMeaningful: false,
        tooltip: "Ranking across treatments using NPV (higher NPV ranks higher). This is descriptive, not a rule."
      }
    ];

    const table = {
      control,
      controlM,
      treatments: treatments.filter(t => !t.isControl),
      indicators,
      metricsById,
      rankById
    };

    return table;
  }

  function calcDeltaAbs(controlVal, treatVal) {
    if (!isFinite(controlVal) || !isFinite(treatVal)) return NaN;
    return treatVal - controlVal;
  }

  function calcDeltaPct(controlVal, treatVal) {
    if (!isFinite(controlVal) || !isFinite(treatVal)) return NaN;
    if (controlVal === 0) return NaN; // percentage not meaningful
    return ((treatVal - controlVal) / Math.abs(controlVal)) * 100;
  }

  function renderLeaderboard(comp) {
    const root = document.getElementById("leaderboard");
    if (!root) return;

    const rate = model.time.discBase;
    const adoptMul = model.adoption.base;
    const risk = model.risk.base;

    const control = comp.control;
    const controlM = comp.controlM;

    const rows = model.treatments
      .map(t => {
        const m = comp.metricsById.get(t.id);
        const rank = comp.rankById.get(t.id) || null;
        const dNpv = controlM ? calcDeltaAbs(controlM.npv, m.npv) : NaN;
        const dCost = controlM ? calcDeltaAbs(controlM.pvCost, m.pvCost) : NaN;
        return { t, m, rank, dNpv, dCost };
      })
      .sort((a, b) => (a.rank || 9999) - (b.rank || 9999));

    // Snapshot: show all, but allow quick filter toggles without hiding permanently (toggle only)
    const mode = (document.querySelector("input[name='lbFilter']:checked")?.value || "all").toLowerCase();

    let filtered = rows;
    if (mode === "topnpv") filtered = rows.slice().sort((a, b) => (b.m?.npv || -Infinity) - (a.m?.npv || -Infinity)).slice(0, 5);
    if (mode === "topbcr") filtered = rows.slice().sort((a, b) => (b.m?.bcr || -Infinity) - (a.m?.bcr || -Infinity)).slice(0, 5);
    if (mode === "improve") filtered = rows.filter(r => isFinite(r.dNpv) && r.dNpv > 0);

    root.innerHTML = `
      <div class="leaderboard-head">
        <div class="leaderboard-title">
          <h3>Snapshot leaderboard</h3>
          <div class="muted small">Base case: discount ${fmt(rate)}%, adoption ${fmt(adoptMul)}, risk ${fmt(risk)}.</div>
        </div>

        <div class="leaderboard-filters" role="group" aria-label="Leaderboard filters">
          <label class="chip"><input type="radio" name="lbFilter" value="all" ${mode === "all" ? "checked" : ""}> Show all</label>
          <label class="chip"><input type="radio" name="lbFilter" value="topnpv" ${mode === "topnpv" ? "checked" : ""}> Top 5 by NPV</label>
          <label class="chip"><input type="radio" name="lbFilter" value="topbcr" ${mode === "topbcr" ? "checked" : ""}> Top 5 by BCR</label>
          <label class="chip"><input type="radio" name="lbFilter" value="improve" ${mode === "improve" ? "checked" : ""}> Improvements vs control</label>
        </div>
      </div>

      <div class="table-wrap compact">
        <table class="table leaderboard-table">
          <thead>
            <tr>
              <th class="sticky-col">Rank</th>
              <th>Treatment</th>
              <th>NPV</th>
              <th>BCR</th>
              <th>PV Costs</th>
              <th>ΔNPV vs Control</th>
              <th>ΔPV Cost vs Control</th>
            </tr>
          </thead>
          <tbody>
            ${filtered
              .map(r => {
                const isCtrl = !!r.t.isControl;
                const name = esc(r.t.name + (isCtrl ? " (Control)" : ""));
                const rank = r.rank != null ? r.rank : "n/a";
                const npv = money(r.m?.npv);
                const bcr = isFinite(r.m?.bcr) ? fmt(r.m.bcr) : "n/a";
                const pvc = money(r.m?.pvCost);
                const dnpv = isFinite(r.dNpv) ? money(r.dNpv) : "n/a";
                const dcost = isFinite(r.dCost) ? money(r.dCost) : "n/a";
                const cls = isCtrl ? "row-control" : "";
                return `
                  <tr class="${cls}">
                    <td class="sticky-col">${rank}</td>
                    <td>${name}</td>
                    <td class="${(r.m?.npv ?? 0) >= 0 ? "pos" : "neg"}">${npv}</td>
                    <td>${bcr}</td>
                    <td>${pvc}</td>
                    <td class="${(r.dNpv ?? 0) >= 0 ? "pos" : "neg"}">${dnpv}</td>
                    <td class="${(r.dCost ?? 0) <= 0 ? "pos" : "neg"}">${dcost}</td>
                  </tr>
                `;
              })
              .join("")}
          </tbody>
        </table>
      </div>
    `;

    // Re-bind filter radios
    root.querySelectorAll("input[name='lbFilter']").forEach(inp => {
      inp.addEventListener("change", () => renderLeaderboard(comp));
    });
  }

  function renderComparisonTable(comp) {
    const root = document.getElementById("comparisonTable");
    if (!root) return;

    const control = comp.control;
    const controlM = comp.controlM;

    // Build table with grouped columns:
    // - First sticky column: indicator names
    // - Control column: value only
    // - For each treatment: Value + Δ (abs) + Δ (% where meaningful)
    const treatments = comp.treatments.slice();

    const headerTopCells = [
      `<th class="sticky-col" rowspan="2">Indicator</th>`,
      `<th class="sticky-head control-col" rowspan="2">Control (baseline)</th>`
    ];

    const headerSubCells = [
      // indicator column has no subheader
      // control has no subheader
    ];

    treatments.forEach(t => {
      headerTopCells.push(`<th class="sticky-head" colspan="3">${esc(t.name)}</th>`);
      headerSubCells.push(`<th class="subhead">Value</th>`);
      headerSubCells.push(`<th class="subhead">Δ vs Control</th>`);
      headerSubCells.push(`<th class="subhead">Δ %</th>`);
    });

    const bodyRows = comp.indicators
      .map(ind => {
        const labelHtml = `
          <div class="indicator-label">
            <span>${esc(ind.label)}</span>
            <button type="button" class="tip" data-tooltip="${esc(ind.tooltip)}" aria-label="Explain ${esc(ind.label)}">i</button>
          </div>
        `;

        const controlVal =
          ind.key === "rank" ? (control ? comp.rankById.get(control.id) : NaN) : (controlM ? controlM[ind.key] : NaN);

        const controlCell = ind.key === "rank" ? (isFinite(controlVal) ? String(controlVal) : "n/a") : ind.format(controlVal);

        const tds = treatments
          .map(t => {
            const m = comp.metricsById.get(t.id);
            const treatVal = ind.key === "rank" ? comp.rankById.get(t.id) : (m ? m[ind.key] : NaN);

            const dAbs = controlM && ind.key !== "rank" ? calcDeltaAbs(controlVal, treatVal) : NaN;
            const dPct = controlM && ind.key !== "rank" ? calcDeltaPct(controlVal, treatVal) : NaN;

            const vCell = ind.key === "rank" ? (isFinite(treatVal) ? String(treatVal) : "n/a") : ind.format(treatVal);
            const dAbsCell = ind.key === "rank" ? "—" : (isFinite(dAbs) ? ind.format(dAbs) : "n/a");

            const pctMeaningful = ind.deltaPctMeaningful && isFinite(dPct);
            const dPctCell = ind.key === "rank" ? "—" : (pctMeaningful ? percent(dPct) : "n/a");

            const clsV =
              ind.key === "npv"
                ? (isFinite(treatVal) && treatVal >= 0 ? "pos" : "neg")
                : "";

            const clsAbs =
              ind.key === "pvCost"
                ? (isFinite(dAbs) && dAbs <= 0 ? "pos" : "neg") // lower cost vs control is "good" (pos)
                : (isFinite(dAbs) && dAbs >= 0 ? "pos" : "neg");

            return `
              <td class="${clsV}">${vCell}</td>
              <td class="${isFinite(dAbs) ? clsAbs : ""}">${dAbsCell}</td>
              <td>${dPctCell}</td>
            `;
          })
          .join("");

        return `
          <tr>
            <td class="sticky-col">${labelHtml}</td>
            <td class="control-col">${controlCell}</td>
            ${tds}
          </tr>
        `;
      })
      .join("");

    root.innerHTML = `
      <div class="table-actions">
        <div class="table-actions-left">
          <button class="btn primary" id="exportResultsXlsx">Export comparison to Excel</button>
          <button class="btn ghost" id="exportResultsCsv">Export comparison to CSV</button>
          <button class="btn ghost" id="printResults">Print / Save as PDF</button>
        </div>
        <div class="table-actions-right small muted">
          Sticky headers enabled. Scroll horizontally if you have many treatments.
        </div>
      </div>

      <div class="table-wrap grid">
        <table class="table compare-table" id="compareTable">
          <thead>
            <tr>${headerTopCells.join("")}</tr>
            <tr>
              ${"" /* indicator+control already rowspan=2 */ }
              ${headerSubCells.join("")}
            </tr>
          </thead>
          <tbody>
            ${bodyRows}
          </tbody>
        </table>
      </div>
    `;

    // Bind exports
    const btnX = document.getElementById("exportResultsXlsx");
    if (btnX) btnX.addEventListener("click", () => exportComparisonToExcel(comp));

    const btnC = document.getElementById("exportResultsCsv");
    if (btnC) btnC.addEventListener("click", () => exportComparisonToCsv(comp));

    const btnP = document.getElementById("printResults");
    if (btnP) btnP.addEventListener("click", () => window.print());

    // Tooltip logic (simple popover)
    wireTooltips(root);
  }

  function wireTooltips(scopeEl) {
    const tips = scopeEl.querySelectorAll(".tip[data-tooltip]");
    tips.forEach(btn => {
      btn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        const msg = btn.getAttribute("data-tooltip") || "";
        showToast(msg);
      });
    });
  }

  function exportComparisonToCsv(comp) {
    const control = comp.control;
    const controlM = comp.controlM;
    const treatments = comp.treatments.slice();

    const header = ["Indicator", "Control (baseline)"];
    treatments.forEach(t => {
      header.push(`${t.name} — Value`);
      header.push(`${t.name} — Δ vs Control`);
      header.push(`${t.name} — Δ %`);
    });

    const rows = [header];

    comp.indicators.forEach(ind => {
      const controlVal = ind.key === "rank" ? (control ? comp.rankById.get(control.id) : NaN) : (controlM ? controlM[ind.key] : NaN);
      const controlCell = ind.key === "rank" ? (isFinite(controlVal) ? String(controlVal) : "") : (isFinite(controlVal) ? controlVal : "");

      const row = [ind.label, controlCell];

      treatments.forEach(t => {
        const m = comp.metricsById.get(t.id);
        const treatVal = ind.key === "rank" ? comp.rankById.get(t.id) : (m ? m[ind.key] : NaN);

        const dAbs = controlM && ind.key !== "rank" ? calcDeltaAbs(controlVal, treatVal) : NaN;
        const dPct = controlM && ind.key !== "rank" ? calcDeltaPct(controlVal, treatVal) : NaN;

        row.push(isFinite(treatVal) ? treatVal : "");
        row.push(ind.key === "rank" ? "" : (isFinite(dAbs) ? dAbs : ""));
        row.push(ind.key === "rank" ? "" : (ind.deltaPctMeaningful && isFinite(dPct) ? dPct : ""));
      });

      rows.push(row);
    });

    const csv = rows
      .map(r => r.map(x => (x == null ? "" : String(x).replace(/"/g, '""'))).map(x => `"${x}"`).join(","))
      .join("\r\n");

    downloadFile(`${slug(model.project.name)}_comparison_to_control.csv`, csv, "text/csv");
    showToast("Comparison CSV downloaded.");
  }

  function exportComparisonToExcel(comp) {
    if (typeof XLSX === "undefined") {
      alert("The SheetJS XLSX library is required for Excel export.");
      return;
    }

    const control = comp.control;
    const controlM = comp.controlM;
    const treatments = comp.treatments.slice();

    const header1 = ["Indicator", "Control (baseline)"];
    treatments.forEach(t => {
      header1.push(`${t.name} — Value`);
      header1.push(`${t.name} — Δ vs Control`);
      header1.push(`${t.name} — Δ %`);
    });

    const data = [header1];

    comp.indicators.forEach(ind => {
      const controlVal = ind.key === "rank" ? (control ? comp.rankById.get(control.id) : NaN) : (controlM ? controlM[ind.key] : NaN);

      const row = [ind.label, ind.key === "rank" ? (isFinite(controlVal) ? controlVal : "") : (isFinite(controlVal) ? controlVal : "")];

      treatments.forEach(t => {
        const m = comp.metricsById.get(t.id);
        const treatVal = ind.key === "rank" ? comp.rankById.get(t.id) : (m ? m[ind.key] : NaN);
        const dAbs = controlM && ind.key !== "rank" ? calcDeltaAbs(controlVal, treatVal) : NaN;
        const dPct = controlM && ind.key !== "rank" ? calcDeltaPct(controlVal, treatVal) : NaN;

        row.push(isFinite(treatVal) ? treatVal : "");
        row.push(ind.key === "rank" ? "" : (isFinite(dAbs) ? dAbs : ""));
        row.push(ind.key === "rank" ? "" : (ind.deltaPctMeaningful && isFinite(dPct) ? dPct : ""));
      });

      data.push(row);
    });

    const wb = XLSX.utils.book_new();
    const metaAoA = [
      [TOOL_NAME],
      ["Project", model.project.name],
      ["Organisation", model.project.organisation],
      ["Export date", new Date().toISOString()],
      ["Discount rate (base, %)", model.time.discBase],
      ["Years", model.time.years],
      ["Adoption (base)", model.adoption.base],
      ["Risk (base)", model.risk.base]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(metaAoA), "Meta");

    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), "Comparison_to_Control");

    // Add a sheet with raw data used (preserved)
    const raw = model.dataPipeline.rawRowsPreserved || [];
    const cols = Array.from(new Set(raw.flatMap(r => Object.keys(r))));
    const rawAoA = [cols, ...raw.map(r => cols.map(c => (c in r ? r[c] : "")))];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rawAoA), "Raw_Data_Used");

    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile(`${slug(model.project.name)}_comparison_to_control.xlsx`, out, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Comparison Excel downloaded.");
  }

  // =========================
  // 9) AI PROMPT (NON-PRESCRIPTIVE)
  // =========================
  function buildAIPromptFromResults(comp) {
    // Plain-English, structured prompt based only on computed results.
    // No decision rules, no telling user what to choose.
    const rate = model.time.discBase;
    const years = model.time.years;
    const adopt = model.adoption.base;
    const risk = model.risk.base;

    const control = comp.control;
    const controlM = comp.controlM;

    const rows = model.treatments
      .map(t => {
        const m = comp.metricsById.get(t.id);
        const rank = comp.rankById.get(t.id);
        const dNpv = controlM ? calcDeltaAbs(controlM.npv, m.npv) : NaN;
        const dCost = controlM ? calcDeltaAbs(controlM.pvCost, m.pvCost) : NaN;
        return {
          name: t.name,
          isControl: !!t.isControl,
          rank,
          pvBenefits: m.pvBen,
          pvCosts: m.pvCost,
          npv: m.npv,
          bcr: m.bcr,
          roi: m.roi,
          dNpv,
          dCost
        };
      })
      .sort((a, b) => (a.rank || 9999) - (b.rank || 9999));

    const best = rows.find(r => !r.isControl) || null;
    const worst = rows.slice().reverse().find(r => !r.isControl) || null;

    const lines = [];
    lines.push(`You are interpreting results from a farm cost–benefit analysis tool called "${TOOL_NAME}".`);
    lines.push(`Use plain language suitable for a farmer or on-farm manager. Avoid jargon. Treat this as decision support only.`);
    lines.push(`Do NOT tell the user which treatment to choose. Do NOT impose rules or thresholds (for example do not say "always choose BCR > 1").`);
    lines.push("");
    lines.push("Context (base case settings):");
    lines.push(`- Analysis horizon: ${years} years`);
    lines.push(`- Discount rate: ${fmt(rate)}%`);
    lines.push(`- Adoption multiplier: ${fmt(adopt)}`);
    lines.push(`- Risk multiplier: ${fmt(risk)} (benefits scaled by 1−risk where linked)`);
    lines.push("");
    lines.push("Definitions (use consistently):");
    lines.push(`- NPV = PV Benefits − PV Costs. Positive NPV indicates net economic gain relative to that option's cost stream.`);
    lines.push(`- PV Benefits = discounted value of benefits over time (including yield-driven revenue changes and any additional benefits).`);
    lines.push(`- PV Costs = discounted value of costs over time. Capital cost is a year-0 cost (not discounted). Annual costs are discounted.`);
    lines.push(`- BCR = PV Benefits ÷ PV Costs.`);
    lines.push(`- ROI = NPV ÷ PV Costs (expressed as a percentage).`);
    lines.push("");
    lines.push("Task:");
    lines.push("Write a farmer-facing interpretation (about 1–2 pages) that:");
    lines.push("1) Explains what is driving differences compared with the control (benefit changes vs cost changes).");
    lines.push("2) Highlights trade-offs (for example higher PV benefits but also higher PV costs).");
    lines.push("3) Flags uncertainty and data gaps (missing values were treated as missing, not zero).");
    lines.push("4) For treatments with negative NPV or low BCR, suggest realistic improvement paths framed as options (cost reduction, yield improvement, commodity price, timing, agronomic adjustments), without telling the user what to choose.");
    lines.push("");
    lines.push("Results table (copy into your reasoning; do not invent numbers):");

    // Add compact numeric block
    rows.forEach(r => {
      const tag = r.isControl ? "CONTROL" : "TREATMENT";
      lines.push(
        `- [${tag}] ${r.name}: Rank=${r.rank ?? "n/a"}, PV Benefits=${isFinite(r.pvBenefits) ? money(r.pvBenefits) : "n/a"}, PV Costs=${isFinite(r.pvCosts) ? money(r.pvCosts) : "n/a"}, NPV=${isFinite(r.npv) ? money(r.npv) : "n/a"}, BCR=${isFinite(r.bcr) ? fmt(r.bcr) : "n/a"}, ROI=${isFinite(r.roi) ? percent(r.roi) : "n/a"}`
      );
      if (!r.isControl && controlM) {
        lines.push(
          `  Compared with control: ΔNPV=${isFinite(r.dNpv) ? money(r.dNpv) : "n/a"}, ΔPV Costs=${isFinite(r.dCost) ? money(r.dCost) : "n/a"}`
        );
      }
    });

    lines.push("");
    if (best && worst) {
      lines.push("Focus points (do not turn these into prescriptions):");
      lines.push(
        `- A higher-ranked option appears to be "${best.name}" mainly because NPV is higher than other options. Explain whether that comes from higher benefits, lower costs, or both.`
      );
      lines.push(
        `- A lower-ranked option appears to be "${worst.name}". Explain what is pulling it down (costs, benefits, or both) and what could realistically change that.`
      );
    }

    lines.push("");
    lines.push("Output format:");
    lines.push("Use short paragraphs with clear headings. Avoid technical language where possible.");

    return lines.join("\n");
  }

  function renderAIPanel(comp) {
    const ta = document.getElementById("aiPrompt");
    if (!ta) return;
    ta.value = buildAIPromptFromResults(comp);

    const copyBtn = document.getElementById("copyAIPrompt");
    if (copyBtn) {
      copyBtn.onclick = async () => {
        const txt = ta.value || "";
        try {
          if (navigator.clipboard && navigator.clipboard.writeText) {
            await navigator.clipboard.writeText(txt);
            showToast("AI prompt copied to clipboard.");
          } else {
            ta.select();
            document.execCommand("copy");
            showToast("AI prompt copied.");
          }
        } catch {
          showToast("Unable to copy automatically. Please copy from the text box.");
        }
      };
    }
  }

  // =========================
  // 10) UI / DOM HELPERS
  // =========================
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));
  const num = sel => +(document.querySelector(sel)?.value || 0);

  function setVal(sel, text) {
    const el = document.querySelector(sel);
    if (el) el.textContent = text;
  }

  // =========================
  // 11) TABS
  // =========================
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
      const match = key === target || p.id === target || p.id === "tab-" + target;
      const show = !!match;
      p.classList.toggle("active", show);
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
  }

  // =========================
  // 12) RENDERERS: FORMS
  // =========================
  function setBasicsFieldsFromModel() {
    if ($("#appTitle")) $("#appTitle").textContent = TOOL_NAME;
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
    if ($("#randSeed")) $("#randSeed").value = model.sim.seed ?? "";

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
      inp.value = scenario === "low" ? row.low : scenario === "high" ? row.high : row.base;
    });

    // Data pipeline summary
    renderValidationPanel();
  }

  function renderValidationPanel() {
    const box = document.getElementById("validationPanel");
    if (!box) return;

    const v = model.dataPipeline.validation || { ok: true, issues: [], stats: {} };
    const f = model.dataPipeline.formulaInfo || { found: false, cells: [] };

    const issuesHtml = (v.issues || [])
      .slice(0, 50)
      .map(it => `<li class="${it.level}"><strong>${esc(it.level.toUpperCase())}</strong> — ${esc(it.message)}</li>`)
      .join("");

    const more = (v.issues || []).length > 50 ? `<div class="small muted">Showing first 50 issues. Export or fix in Excel for full resolution.</div>` : "";

    const formulaHtml = f.found
      ? `<div class="warn-box small">
          <strong>Excel formulas detected.</strong>
          The tool uses stored cell values from the workbook. If your workbook relies on formulas without stored values, those cells may become missing.
          <details>
            <summary>Show detected formula cells</summary>
            <ul class="small">${f.cells.slice(0, 50).map(c => `<li><code>${esc(c.cell)}</code>: <code>${esc(c.formula)}</code></li>`).join("")}</ul>
          </details>
        </div>`
      : `<div class="ok-box small">No Excel formulas detected in the loaded sheet.</div>`;

    box.innerHTML = `
      <div class="validation-head">
        <div>
          <div class="small muted">Loaded source</div>
          <div><strong>${esc(model.dataPipeline.lastLoadedSource || "unknown")}</strong></div>
        </div>
        <div>
          <div class="small muted">Sheet used</div>
          <div><strong>${esc(model.dataPipeline.lastLoadedSheet || "unknown")}</strong></div>
        </div>
        <div>
          <div class="small muted">Status</div>
          <div><strong class="${v.ok ? "pos" : "neg"}">${v.ok ? "OK" : "Issues found"}</strong></div>
        </div>
      </div>

      ${formulaHtml}

      <div class="validation-stats small">
        <div class="stat"><span>Rows:</span> <strong>${fmt(v.stats?.nRows ?? 0)}</strong></div>
        <div class="stat"><span>Missing yields:</span> <strong>${fmt(v.stats?.missingYield ?? 0)}</strong></div>
        <div class="stat"><span>Missing labour:</span> <strong>${fmt(v.stats?.missingLabour ?? 0)}</strong></div>
        <div class="stat"><span>Missing input cost:</span> <strong>${fmt(v.stats?.missingInputCost ?? 0)}</strong></div>
      </div>

      <details class="validation-issues" ${v.ok ? "" : "open"}>
        <summary>Validation messages</summary>
        <ul class="issues-list">${issuesHtml || `<li class="ok">No validation messages.</li>`}</ul>
        ${more}
      </details>
    `;
  }

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
              ${["Farm Trials", "Plant Farm", "ABARES", "GRDC", "Input Directly"]
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
    root.oninput = e => {
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
      calcAndRenderDebounced();
    };
    root.onclick = e => {
      const id = e.target.dataset.delOutput;
      if (!id) return;
      if (!confirm("Remove this output metric?")) return;
      model.outputs = model.outputs.filter(o => o.id !== id);
      model.treatments.forEach(t => delete t.deltas[id]);
      renderOutputs();
      renderTreatments();
      calcAndRender();
      showToast("Output metric removed.");
    };
  }

  function renderTreatments() {
    const root = $("#treatmentsList");
    if (!root) return;
    root.innerHTML = "";

    model.treatments.forEach(t => {
      const materials = Number(t.materialsCost) || 0;
      const services = Number(t.servicesCost) || 0;
      const labour = Number(t.labourCost) || 0;

      // Strict requirement: capital cost appears before total cost ($/ha)
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
              ${["Farm Trials", "Plant Farm", "ABARES", "GRDC", "Input Directly"]
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
          <div class="field"><label>Materials cost ($/ha, annual)</label><input type="number" step="0.01" value="${t.materialsCost || 0}" data-tk="materialsCost" data-id="${t.id}" /></div>
          <div class="field"><label>Services cost ($/ha, annual)</label><input type="number" step="0.01" value="${t.servicesCost || 0}" data-tk="servicesCost" data-id="${t.id}" /></div>
          <div class="field"><label>Labour cost ($/ha, annual)</label><input type="number" step="0.01" value="${t.labourCost || 0}" data-tk="labourCost" data-id="${t.id}" /></div>

          <div class="field"><label>Capital cost ($, year 0)</label><input type="number" step="0.01" value="${t.capitalCost || 0}" data-tk="capitalCost" data-id="${t.id}" /></div>

          <div class="field"><label>Total cost ($/ha, annual)</label><input type="number" step="0.01" value="${totalPerHa}" readonly data-total-cost="${t.id}" /></div>

          <div class="field"><label>Constrained?</label>
            <select data-tk="constrained" data-id="${t.id}">
              <option value="true" ${t.constrained ? "selected" : ""}>Yes</option>
              <option value="false" ${!t.constrained ? "selected" : ""}>No</option>
            </select>
          </div>
        </div>

        <div class="field">
          <label>Notes</label>
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
        if (tk === "constrained") t[tk] = e.target.value === "true";
        else if (tk === "name" || tk === "source" || tk === "notes") t[tk] = e.target.value;
        else if (tk === "isControl") {
          const val = e.target.value === "control";
          // Enforce single control
          model.treatments.forEach(tt => (tt.isControl = false));
          if (val) t.isControl = true;
          renderTreatments();
          calcAndRenderDebounced();
          showToast(`Control baseline set to: ${t.name}`);
          return;
        } else t[tk] = +e.target.value;

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

      calcAndRenderDebounced();
    };

    root.onclick = e => {
      const id = e.target.dataset.delTreatment;
      if (!id) return;
      if (!confirm("Remove this treatment?")) return;
      model.treatments = model.treatments.filter(x => x.id !== id);
      renderTreatments();
      calcAndRender();
      showToast("Treatment removed.");
    };
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
          <div class="field"><label>P0</label><input type="number" step="0.001" value="${b.p0 || 0}" data-bk="p0" data-id="${b.id}" /></div>
          <div class="field"><label>P1</label><input type="number" step="0.001" value="${b.p1 || 0}" data-bk="p1" data-id="${b.id}" /></div>
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

    root.onclick = e => {
      const id = e.target.dataset.delBenefit;
      if (!id) return;
      if (!confirm("Remove this benefit item?")) return;
      model.benefits = model.benefits.filter(x => x.id !== id);
      renderBenefits();
      calcAndRender();
      showToast("Benefit item removed.");
    };
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
          <div class="field"><label>Annual ($/year)</label><input type="number" step="0.01" value="${c.annual ?? 0}" data-ck="annual" data-id="${c.id}" /></div>
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
            <div class="field"><label>Declining rate (%/year)</label><input type="number" step="1" value="${c.depRate || 30}" data-ck="depRate" data-id="${c.id}" /></div>
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

      renderCosts(); // keep depreciation fields consistent
      calcAndRenderDebounced();
    };

    root.onclick = e => {
      const id = e.target.dataset.delCost;
      if (!id) return;
      if (!confirm("Remove this cost item?")) return;
      model.otherCosts = model.otherCosts.filter(x => x.id !== id);
      renderCosts();
      calcAndRender();
      showToast("Cost item removed.");
    };
  }

  function renderAll() {
    renderOutputs();
    renderTreatments();
    renderBenefits();
    renderCosts();
    setBasicsFieldsFromModel();
  }

  // =========================
  // 13) RESULTS SUMMARY (WHOLE PROJECT)
  // =========================
  function renderWholeProjectSummary(all) {
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
    setVal("#payback", all.paybackYears != null ? String(all.paybackYears) : "Not reached");
  }

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

  // =========================
  // 14) CHARTS
  // =========================
  function drawTimeSeries(canvasId, xs, ys) {
    const canvas = document.getElementById(canvasId);
    if (!canvas || !canvas.getContext) return;
    const ctx = canvas.getContext("2d");
    const w = canvas.width;
    const h = canvas.height;
    ctx.clearRect(0, 0, w, h);
    if (!xs?.length || !ys?.length) return;

    const minX = Math.min(...xs);
    const maxX = Math.max(...xs);
    const minY = Math.min(...ys);
    const maxY = Math.max(...ys);

    const pad = 28;
    const plotW = w - pad * 2;
    const plotH = h - pad * 2;

    const xScale = v => pad + ((v - minX) / (maxX - minX || 1)) * plotW;
    const yScale = v => pad + plotH - ((v - minY) / (maxY - minY || 1)) * plotH;

    ctx.strokeStyle = "#cfd6e4";
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

    ctx.strokeStyle = "#cfd6e4";
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
  // 15) DEPRECIATION SUMMARY
  // =========================
  function renderDepreciationSummary() {
    const root = $("#depSummary");
    if (!root) return;
    root.innerHTML = "";

    const N = model.time.years;
    const baseYear = model.time.startYear;
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
          if (idx >= 0 && idx <= N) sched[idx] = (sched[idx] || 0) + annual;
        }
      } else if (method === "declining") {
        let book = cost;
        for (let i = 0; i < life; i++) {
          const dep = (book * rate) / 100;
          const idx = startIndex + i;
          if (idx >= 0 && idx <= N) sched[idx] = (sched[idx] || 0) + dep;
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
      root.innerHTML = `<p class="small muted">No capital items with depreciation configured. Set depreciation for capital costs to see a schedule summary.</p>`;
      return;
    }

    root.innerHTML = `
      <div class="table-wrap compact">
        <table class="table dep-table">
          <thead>
            <tr>
              <th>Cost item</th>
              <th>Method</th>
              <th>Life (years)</th>
              <th>Rate</th>
              <th>Approx. first-year depreciation</th>
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
                <td>${esc(String(r.rate || ""))}</td>
                <td>${money(r.firstDep)}</td>
              </tr>
            `
              )
              .join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  // =========================
  // 16) MAIN CALC / RENDER
  // =========================
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
    renderWholeProjectSummary(all);

    // Control-centric results
    const comp = computeComparisonToControl(rate, adoptMul, risk);
    renderLeaderboard(comp);
    renderComparisonTable(comp);
    renderAIPanel(comp);

    // Time projections
    renderTimeProjections(all.benefitByYear, all.costByYear, rate);

    // Depreciation summary
    renderDepreciationSummary();

    // Update validation panel
    renderValidationPanel();
  }

  // =========================
  // 17) SIMULATION (MONTE CARLO)
  // =========================
  async function runSimulation() {
    const status = $("#simStatus");
    if (status) status.textContent = "Running Monte Carlo simulation ...";
    await new Promise(r => setTimeout(r));

    const N = Math.max(100, Number(model.sim.n) || 1000);
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

    const baseOutValues = model.outputs.map(o => Number(o.value) || 0);
    const baseTreatCosts = model.treatments.map(t => ({
      materials: Number(t.materialsCost) || 0,
      services: Number(t.servicesCost) || 0,
      labour: Number(t.labourCost) || 0,
      capital: Number(t.capitalCost) || 0
    }));
    const baseOtherCosts = model.otherCosts.map(c => ({ annual: Number(c.annual) || 0, capital: Number(c.capital) || 0 }));

    for (let i = 0; i < N; i++) {
      const disc = triangular(rand(), discLow, discBase, discHigh);
      const adoptMul = clamp(triangular(rand(), adoptLow, adoptBase, adoptHigh), 0, 1);
      const risk = clamp(triangular(rand(), riskLow, riskBase, riskHigh), 0, 1);

      const shockOutputs = model.sim.varyOutputs ? 1 + (rand() * 2 * varPct - varPct) : 1;
      const shockTreatCosts = model.sim.varyTreatCosts ? 1 + (rand() * 2 * varPct - varPct) : 1;
      const shockInputCosts = model.sim.varyInputCosts ? 1 + (rand() * 2 * varPct - varPct) : 1;

      if (model.sim.varyOutputs) model.outputs.forEach((o, idx) => (o.value = baseOutValues[idx] * shockOutputs));

      if (model.sim.varyTreatCosts) {
        model.treatments.forEach((t, idx) => {
          const base = baseTreatCosts[idx];
          t.materialsCost = base.materials * shockTreatCosts;
          t.servicesCost = base.services * shockTreatCosts;
          t.labourCost = base.labour * shockTreatCosts;
          // Capital stays as entered (year-0). If you want to vary capital too, toggle it explicitly in UI later.
          t.capitalCost = base.capital;
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

    // Restore base values
    model.outputs.forEach((o, idx) => (o.value = baseOutValues[idx]));
    model.treatments.forEach((t, idx) => {
      const base = baseTreatCosts[idx];
      t.materialsCost = base.materials;
      t.servicesCost = base.services;
      t.labourCost = base.labour;
      t.capitalCost = base.capital;
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

    const target = Number(model.sim.targetBCR) || 0;
    const probBcrGt1 = bcrs.filter(v => isFinite(v) && v > 1).length / bcrs.length || 0;
    const probBcrGtTarget = bcrs.filter(v => isFinite(v) && v > target).length / bcrs.length || 0;

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

    if (status) status.textContent = `Simulation complete for ${N.toLocaleString()} runs.`;
    showToast("Simulation complete.");
  }

  // =========================
  // 18) EXPORTS (PROJECT JSON / PDF)
  // =========================
  function exportProjectJson() {
    const data = JSON.stringify(model, null, 2);
    downloadFile(`cba_${slug(model.project.name)}.json`, data, "application/json");
    showToast("Project JSON downloaded.");
  }

  function exportPdf() {
    window.print();
  }

  // =========================
  // 19) EXCEL DOWNLOAD / UPLOAD
  // =========================
  function downloadExcelTemplate() {
    if (typeof XLSX === "undefined") {
      alert("The SheetJS XLSX library is required for Excel export.");
      return;
    }
    const wb = buildDefaultWorkbookFromEmbeddedData();
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile(`${slug(model.project.name)}_template.xlsx`, out, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Excel template downloaded.");
  }

  function downloadCurrentDatasetWorkbook() {
    if (typeof XLSX === "undefined") {
      alert("The SheetJS XLSX library is required for Excel export.");
      return;
    }

    const wb = XLSX.utils.book_new();

    const readmeAoA = [
      [TOOL_NAME],
      ["This workbook contains the raw dataset currently loaded in the tool (preserved exactly)."],
      [""],
      ["How to use:"],
      ["1) Edit the FabaBeanRaw sheet in Excel."],
      ["2) Upload it using Data import → Upload Excel."],
      ["3) The tool will validate and recalibrate treatments automatically."]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(readmeAoA), "ReadMe");

    const raw = model.dataPipeline.rawRowsPreserved || [];
    const cols = Array.from(new Set(raw.flatMap(r => Object.keys(r))));
    const aoa = [cols, ...raw.map(r => cols.map(c => (c in r ? r[c] : "")))];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "FabaBeanRaw");

    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile(`${slug(model.project.name)}_current_dataset.xlsx`, out, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Current dataset workbook downloaded.");
  }

  async function handleUploadExcel(file) {
    if (typeof XLSX === "undefined") {
      alert("The SheetJS XLSX library is required for Excel import.");
      return;
    }
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array", cellFormula: true, cellNF: true, cellText: true });

      await loadWorkbookThroughPipeline({ wb, sourceLabel: "uploaded" });
      renderAll();
      calcAndRender();
    } catch (err) {
      console.error(err);
      alert("Error reading Excel file. Please check the file and try again.");
    }
  }

  // =========================
  // 20) EVENT BINDINGS
  // =========================
  function bindGlobalButtons() {
    const startBtn = $("#startBtn");
    if (startBtn) startBtn.addEventListener("click", () => switchTab("project"));

    const saveBtn = $("#saveProject");
    if (saveBtn) saveBtn.addEventListener("click", exportProjectJson);

    const loadBtn = $("#loadProject");
    const loadFileInput = $("#loadFile");
    if (loadBtn && loadFileInput) {
      loadBtn.addEventListener("click", () => loadFileInput.click());
      loadFileInput.addEventListener("change", async e => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        const text = await file.text();
        try {
          const obj = JSON.parse(text);
          Object.assign(model, obj);
          if (!model.time.discountSchedule) model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
          initTreatmentDeltas();
          renderAll();
          calcAndRender();
          showToast("Project JSON loaded and applied.");
        } catch (err) {
          console.error(err);
          alert("Invalid JSON file.");
        } finally {
          e.target.value = "";
        }
      });
    }

    const downloadTemplateBtn = $("#downloadTemplate");
    if (downloadTemplateBtn) downloadTemplateBtn.addEventListener("click", downloadExcelTemplate);

    const downloadCurrentBtn = $("#downloadCurrentDataset");
    if (downloadCurrentBtn) downloadCurrentBtn.addEventListener("click", downloadCurrentDatasetWorkbook);

    const uploadInput = $("#uploadExcelFile");
    if (uploadInput) {
      uploadInput.addEventListener("change", async e => {
        const file = e.target.files && e.target.files[0];
        await handleUploadExcel(file);
        e.target.value = "";
      });
    }

    const recalcBtns = $$("#recalc, #getResults, [data-action='recalc']");
    recalcBtns.forEach(btn => btn.addEventListener("click", e => {
      e.preventDefault();
      calcAndRender();
      showToast("Results recalculated.");
    }));

    const exportPdfBtn = $("#exportPdf");
    if (exportPdfBtn) exportPdfBtn.addEventListener("click", exportPdf);

    const runSimBtn = $("#runSim");
    if (runSimBtn) runSimBtn.addEventListener("click", runSimulation);

    const addOutputBtn = $("#addOutput");
    if (addOutputBtn) addOutputBtn.addEventListener("click", () => {
      const id = uid();
      model.outputs.push({ id, name: "Custom output", unit: "unit", value: 0, source: "Input Directly" });
      model.treatments.forEach(t => (t.deltas[id] = 0));
      renderOutputs();
      renderTreatments();
      calcAndRender();
      showToast("New output metric added.");
    });

    const addTreatmentBtn = $("#addTreatment");
    if (addTreatmentBtn) addTreatmentBtn.addEventListener("click", () => {
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
      calcAndRender();
      showToast("New treatment added.");
    });

    const addBenefitBtn = $("#addBenefit");
    if (addBenefitBtn) addBenefitBtn.addEventListener("click", () => {
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

    const addCostBtn = $("#addCost");
    if (addCostBtn) addCostBtn.addEventListener("click", () => {
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
      showToast("New cost item added.");
    });

    const calcRiskBtn = $("#calcCombinedRisk");
    if (calcRiskBtn) calcRiskBtn.addEventListener("click", () => {
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
      showToast("Combined risk updated.");
    });

    // Central input binding for settings/project fields
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
          else if (scenario === "high") row.high = val;
          else row.base = val;
          calcAndRenderDebounced();
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
        case "targetBCR": model.sim.targetBCR = +t.value; break;
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
    });
  }

  // =========================
  // 21) INIT (DEFAULT LOAD THROUGH PIPELINE)
  // =========================
  document.addEventListener("DOMContentLoaded", async () => {
    initTabs();
    bindGlobalButtons();

    // Default dataset must be parsed through the same pipeline as uploads.
    if (typeof XLSX !== "undefined") {
      const wb = buildDefaultWorkbookFromEmbeddedData();
      if (wb) {
        await loadWorkbookThroughPipeline({ wb, sourceLabel: "default" });
      } else {
        // Fallback: still commit raw rows; tool remains usable, but export/import require XLSX
        model.dataPipeline.validation = validateRawRows(DEFAULT_RAW_PLOTS.map(r => Object.assign({}, r)));
        model.dataPipeline.rawRowsPreserved = DEFAULT_RAW_PLOTS.map(r => Object.assign({}, r));
        commitRawRowsToModel(DEFAULT_RAW_PLOTS.map(r => Object.assign({}, r)), "default");
      }
    } else {
      // If XLSX not loaded, we cannot honor Excel-first fully; still keep tool usable
      model.dataPipeline.validation = validateRawRows(DEFAULT_RAW_PLOTS.map(r => Object.assign({}, r)));
      model.dataPipeline.rawRowsPreserved = DEFAULT_RAW_PLOTS.map(r => Object.assign({}, r));
      commitRawRowsToModel(DEFAULT_RAW_PLOTS.map(r => Object.assign({}, r)), "default");
      showToast("Warning: XLSX library not found. Excel import/export disabled.");
    }

    // Render UI and compute
    renderAll();
    calcAndRender();

    // Results tab must open by default
    switchTab("results");
  });
})();

/* APP.JS PATCH (paste into your existing app.js near the end, before the closing of your IIFE)
   This module is self-contained and only assumes:
   - your existing showToast(message) exists (it does in your app.js paste)
   - your existing downloadFile(name, content, mime) exists (if not, this patch provides a fallback)
*/
(function attachFabaBeansTrialCBA(){
  "use strict";

  // ---------- FALLBACKS ----------
  const hasToast = typeof window.showToast === "function";
  const toast = (m)=>{ try{ (hasToast? window.showToast : alert)(m);}catch(_){ /* ignore */ } };

  const downloadFileFallback = (filename, content, mime) => {
    const blob = new Blob([content], { type: mime || "application/octet-stream" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(()=>{ URL.revokeObjectURL(a.href); a.remove(); }, 0);
  };
  const downloadFile = (typeof window.downloadFile === "function") ? window.downloadFile : downloadFileFallback;

  const $ = (sel)=>document.querySelector(sel);

  // ---------- DEFAULT SCENARIOS ----------
  const DEFAULT = {
    prices: [300,350,400,450,500],
    discounts: [0.05,0.07,0.10],
    horizon: 10,
    persistence: {
      "1yr_only": [1,0,0,0,0,0,0,0,0,0],
      "3yr_decay": [1,0.5,0.25,0,0,0,0,0,0,0],
      "5yr_decay": [1,0.8,0.6,0.4,0.2,0,0,0,0,0],
      "10yr_constant": [1,1,1,1,1,1,1,1,1,1]
    }
  };

  // ---------- STATE ----------
  const state = {
    rawText: "",
    dictText: "",
    rows: [],
    columns: [],
    dict: null,
    validation: null,
    committed: false,
    replicateStats: null,
    treatmentStats: null,
    config: new Map(), // key: treatmentKey -> { include:boolean, recurrence:'one_off'|'annual'|'custom', customCostPath?:number[] }
    scenario: {
      horizon: DEFAULT.horizon,
      prices: [...DEFAULT.prices],
      discounts: [...DEFAULT.discounts],
      persistence: JSON.parse(JSON.stringify(DEFAULT.persistence)),
      selectedPrice: DEFAULT.prices[3],
      selectedDiscount: DEFAULT.discounts[1],
      selectedPersistence: "10yr_constant",
      scale: 1
    },
    filters: { mode: "all" }
  };

  // ---------- DICTIONARY PARSE ----------
  function parseCSVLine(line){
    const out = [];
    let cur = "";
    let inQ = false;
    for(let i=0;i<line.length;i++){
      const ch = line[i];
      if(ch === '"' ){
        if(inQ && line[i+1] === '"'){ cur += '"'; i++; }
        else inQ = !inQ;
      } else if(ch === "," && !inQ){
        out.push(cur);
        cur = "";
      } else {
        cur += ch;
      }
    }
    out.push(cur);
    return out.map(s=>s.trim());
  }

  function parseDictionaryCSV(text){
    const lines = (text||"").replace(/\r\n/g,"\n").replace(/\r/g,"\n").split("\n").filter(l=>l.trim().length);
    if(!lines.length) return null;
    const header = parseCSVLine(lines[0]);
    const idx = (name)=>header.findIndex(h=>h.toLowerCase().trim()===name.toLowerCase().trim());
    const colIdx = idx("column_name");
    const labelIdx = idx("label");
    const typeIdx = idx("type");
    const reqIdx = idx("required");
    const tipIdx = idx("tooltip");
    const roleIdx = idx("role");
    const unitIdx = idx("unit");

    const dict = {
      byCol: new Map(),
      rawHeader: header
    };

    for(let i=1;i<lines.length;i++){
      const row = parseCSVLine(lines[i]);
      const col = (colIdx>=0? row[colIdx] : row[0])?.trim();
      if(!col) continue;
      const meta = {
        column_name: col,
        label: (labelIdx>=0? row[labelIdx] : "") || col,
        type: (typeIdx>=0? row[typeIdx] : "") || "",
        required: (reqIdx>=0? row[reqIdx] : "") || "",
        tooltip: (tipIdx>=0? row[tipIdx] : "") || "",
        role: (roleIdx>=0? row[roleIdx] : "") || "",
        unit: (unitIdx>=0? row[unitIdx] : "") || ""
      };
      dict.byCol.set(col, meta);
    }
    return dict;
  }

  function colLabel(col){
    const d = state.dict?.byCol?.get(col);
    return d?.label || col;
  }
  function colTip(col){
    const d = state.dict?.byCol?.get(col);
    return d?.tooltip || "";
  }

  // ---------- DATA PARSE ----------
  function splitLines(text){
    return (text||"").replace(/\r\n/g,"\n").replace(/\r/g,"\n").split("\n");
  }

  function sniffDelimiter(text){
    const head = splitLines(text).slice(0,3).join("\n");
    const tabs = (head.match(/\t/g)||[]).length;
    const commas = (head.match(/,/g)||[]).length;
    return (tabs>=commas) ? "\t" : ",";
  }

  function parseDelimited(text){
    const delim = sniffDelimiter(text);
    const lines = splitLines(text).filter(l=>l.trim().length);
    if(!lines.length) return { columns:[], rows:[] };

    const header = (delim === "\t") ? lines[0].split("\t") : parseCSVLine(lines[0]);
    const columns = header.map(h=>h.trim()).filter(h=>h.length);

    // Guard against unnamed columns
    const unnamed = columns.filter(c=>/^unnamed/i.test(c) || c==="" );
    if(unnamed.length){
      // keep them (no dropping) but flag loudly later
    }

    const rows = [];
    for(let i=1;i<lines.length;i++){
      const parts = (delim === "\t") ? lines[i].split("\t") : parseCSVLine(lines[i]);
      const obj = {};
      for(let c=0;c<columns.length;c++){
        obj[columns[c]] = (parts[c] ?? "").trim();
      }
      rows.push(obj);
    }
    return { columns, rows };
  }

  function isMissing(v){
    if(v === null || v === undefined) return true;
    const s = String(v).trim();
    return s === "" || s === "?" ;
  }

  function toNumberOrNaN(v){
    if(isMissing(v)) return NaN;
    if(typeof v === "number") return Number.isFinite(v) ? v : NaN;
    const cleaned = String(v).trim().replace(/[\$,]/g,"");
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : NaN;
  }

  // ---------- REQUIRED COLUMN RESOLUTION ----------
  // Use dictionary roles if present, else fall back to the explicit names given in your spec.
  function resolveColumns(){
    const cols = new Set(state.columns);
    const byRole = new Map();
    if(state.dict?.byCol){
      for(const [c,m] of state.dict.byCol.entries()){
        const role = (m.role||"").trim();
        if(role) byRole.set(role, c);
      }
    }

    const pick = (role, fallbackName)=>{
      const a = byRole.get(role);
      if(a && cols.has(a)) return a;
      if(fallbackName && cols.has(fallbackName)) return fallbackName;
      return null;
    };

    const resolved = {
      replicate_id: pick("replicate_id","replicate_id"),
      treatment_id: pick("treatment_id","treatment_id"),
      amendment_name: pick("amendment_name","amendment_name"),
      yield_t_ha: pick("yield_t_ha","yield_t_ha"),
      total_cost_per_ha_raw: pick("total_cost_per_ha_raw","total_cost_per_ha_raw"),
      cost_amendment_input_per_ha_raw: pick("cost_amendment_input_per_ha_raw","cost_amendment_input_per_ha_raw")
    };

    return resolved;
  }

  // ---------- VALIDATION + COST RECONSTRUCTION ----------
  function reconstructCosts(rows, col){
    const out = [];
    const scalingHits = new Map(); // amendment -> {count, rawMin, rawMax, rawMean, adjMin, adjMax, adjMean}
    const amendCol = col.amendment_name;

    let sumRaw=0, sumAdj=0, nRaw=0, nAdj=0;
    for(const r of rows){
      const rawTotal = toNumberOrNaN(r[col.total_cost_per_ha_raw]);
      const rawAmend = toNumberOrNaN(r[col.cost_amendment_input_per_ha_raw]);
      const amendName = String(r[amendCol] ?? "").trim();

      let adjAmend = rawAmend;
      let scaled = false;
      if(Number.isFinite(rawAmend) && rawAmend > 1000){
        adjAmend = rawAmend / 100;
        scaled = true;
      }
      let total = rawTotal;
      if(Number.isFinite(rawTotal) && Number.isFinite(rawAmend) && Number.isFinite(adjAmend)){
        total = rawTotal - rawAmend + adjAmend;
      } else {
        // if missing parts, total becomes NaN (explicitly missing) rather than silently using rawTotal
        total = NaN;
      }

      out.push({ rawTotal, rawAmend, adjAmend, total, scaled, amendName });

      if(scaled && amendName){
        let s = scalingHits.get(amendName);
        if(!s){
          s = { count:0, rawMin:+Infinity, rawMax:-Infinity, rawSum:0, adjMin:+Infinity, adjMax:-Infinity, adjSum:0 };
          scalingHits.set(amendName, s);
        }
        s.count++;
        if(Number.isFinite(rawAmend)){ s.rawMin=Math.min(s.rawMin,rawAmend); s.rawMax=Math.max(s.rawMax,rawAmend); s.rawSum += rawAmend; }
        if(Number.isFinite(adjAmend)){ s.adjMin=Math.min(s.adjMin,adjAmend); s.adjMax=Math.max(s.adjMax,adjAmend); s.adjSum += adjAmend; }
      }

      if(Number.isFinite(rawAmend)){ sumRaw += rawAmend; nRaw++; }
      if(Number.isFinite(adjAmend)){ sumAdj += adjAmend; nAdj++; }
    }

    const scalingSummary = [];
    for(const [k,v] of scalingHits.entries()){
      scalingSummary.push({
        amendment_name: k,
        count: v.count,
        raw_min: (v.rawMin===+Infinity? NaN : v.rawMin),
        raw_max: (v.rawMax===-Infinity? NaN : v.rawMax),
        raw_mean: v.count ? (v.rawSum / v.count) : NaN,
        adj_min: (v.adjMin===+Infinity? NaN : v.adjMin),
        adj_max: (v.adjMax===-Infinity? NaN : v.adjMax),
        adj_mean: v.count ? (v.adjSum / v.count) : NaN
      });
    }
    scalingSummary.sort((a,b)=>b.count-a.count);

    return { reconstructed: out, scalingSummary, overall: { raw_mean: nRaw? sumRaw/nRaw : NaN, adj_mean: nAdj? sumAdj/nAdj : NaN } };
  }

  function validateAndSummarise(){
    const resolved = resolveColumns();
    const required = Object.entries(resolved).filter(([,v])=>!v).map(([k])=>k);

    const issues = [];
    const unnamedCols = state.columns.filter(c=>/^unnamed/i.test(c) || c==="");
    if(unnamedCols.length){
      issues.push({ level:"bad", msg:`The dataset includes unnamed columns (${unnamedCols.slice(0,3).join(", ")}). The tool will keep them but you should fix headers to avoid ambiguity.` });
    }
    if(required.length){
      issues.push({ level:"bad", msg:`Missing required columns: ${required.join(", ")}. The faba-beans CBA engine requires these fields.` });
    }

    // Missingness scan for key numeric fields
    const keyNumeric = [resolved.yield_t_ha, resolved.total_cost_per_ha_raw, resolved.cost_amendment_input_per_ha_raw].filter(Boolean);
    const missCounts = {};
    for(const c of keyNumeric) missCounts[c] = 0;

    // replicate controls availability
    const repMap = new Map(); // rep -> { n:0, nCtrl:0, ctrlY:[], ctrlC:[] }
    const rec = reconstructCosts(state.rows, resolved);

    for(let i=0;i<state.rows.length;i++){
      const r = state.rows[i];
      const rep = String(r[resolved.replicate_id] ?? "").trim();
      const tid = toNumberOrNaN(r[resolved.treatment_id]);
      const amend = String(r[resolved.amendment_name] ?? "").trim();

      const isCtrl = (amend && amend.toLowerCase() === "control") || (Number.isFinite(tid) && tid === 1);

      const y = toNumberOrNaN(r[resolved.yield_t_ha]);
      const c = rec.reconstructed[i].total;

      for(const cN of keyNumeric){
        const v = (cN === resolved.yield_t_ha) ? y :
                  (cN === resolved.total_cost_per_ha_raw) ? toNumberOrNaN(r[resolved.total_cost_per_ha_raw]) :
                  toNumberOrNaN(r[resolved.cost_amendment_input_per_ha_raw]);
        if(!Number.isFinite(v)) missCounts[cN] += 1;
      }

      if(!rep){
        continue;
      }
      let g = repMap.get(rep);
      if(!g){
        g = { n:0, nCtrl:0, ctrlY:[], ctrlC:[] };
        repMap.set(rep,g);
      }
      g.n++;
      if(isCtrl){
        g.nCtrl++;
        if(Number.isFinite(y)) g.ctrlY.push(y);
        if(Number.isFinite(c)) g.ctrlC.push(c);
      }
    }

    const repIssues = [];
    let repsWithNoCtrl = 0;
    for(const [rep,g] of repMap.entries()){
      if(g.nCtrl === 0){
        repsWithNoCtrl++;
        repIssues.push(rep);
      }
    }
    if(repsWithNoCtrl){
      issues.push({ level:"bad", msg:`Some replicates have no control plots and will be excluded from delta computations: ${repIssues.join(", ")}.` });
    } else {
      issues.push({ level:"good", msg:`All replicates contain control plots. Replicate-specific baselines can be computed safely.` });
    }

    const missText = keyNumeric.map(c=>`${colLabel(c)}: ${missCounts[c]}`).join(" | ");
    issues.push({ level: (Object.values(missCounts).some(n=>n>0) ? "warn" : "good"), msg:`Missing or non-numeric values in key numeric fields (excluded from means): ${missText}.` });

    state.validation = { resolved, issues, missCounts, repMap, scaling: rec.scalingSummary };
    state.replicateStats = { repMap, repsWithNoCtrl, repIssues };
    state.committed = false;
    renderDiagnostics();
    renderDataChecks();
    renderPreviewTable();
    const rowsEl = $("#faba-rows"), colsEl = $("#faba-cols"), repsEl = $("#faba-reps");
    if(rowsEl) rowsEl.textContent = String(state.rows.length);
    if(colsEl) colsEl.textContent = String(state.columns.length);
    if(repsEl) repsEl.textContent = String(repMap.size);

    const ok = required.length === 0 && state.rows.length > 0;
    const commitBtn = $("#faba-commit");
    if(commitBtn) commitBtn.disabled = !ok;
    toast(ok ? "Validation complete. Ready to commit." : "Validation found blocking issues. Fix required columns first.");
    return ok;
  }

  // ---------- REPLICATE BASELINES + TREATMENT STATS ----------
  function mean(arr){
    const c = arr.filter(v=>Number.isFinite(v));
    if(!c.length) return NaN;
    return c.reduce((a,b)=>a+b,0)/c.length;
  }
  function sd(arr){
    const c = arr.filter(v=>Number.isFinite(v));
    if(c.length<2) return NaN;
    const m = mean(c);
    const v = c.reduce((a,b)=>a+(b-m)*(b-m),0)/(c.length-1);
    return Math.sqrt(v);
  }

  function commitDataset(){
    if(!state.validation) {
      toast("Validate first.");
      return;
    }
    const col = state.validation.resolved;
    const rec = reconstructCosts(state.rows, col);

    // compute replicate baselines
    const repBase = new Map(); // rep -> { y0, c0 }
    for(const [rep,g] of state.validation.repMap.entries()){
      const y0 = mean(g.ctrlY);
      const c0 = mean(g.ctrlC);
      repBase.set(rep, { y0, c0, ok: Number.isFinite(y0) && Number.isFinite(c0) && g.nCtrl>0 });
    }

    // build plot-level deltas
    const plot = [];
    for(let i=0;i<state.rows.length;i++){
      const r = state.rows[i];
      const rep = String(r[col.replicate_id] ?? "").trim();
      const tid = toNumberOrNaN(r[col.treatment_id]);
      const amend = String(r[col.amendment_name] ?? "").trim();
      const isCtrl = (amend && amend.toLowerCase() === "control") || (Number.isFinite(tid) && tid === 1);

      const y = toNumberOrNaN(r[col.yield_t_ha]);
      const c = rec.reconstructed[i].total;

      const base = repBase.get(rep);
      const hasBase = base?.ok;

      const dy = (hasBase && Number.isFinite(y)) ? (y - base.y0) : NaN;
      const dc = (hasBase && Number.isFinite(c)) ? (c - base.c0) : NaN;

      plot.push({
        replicate_id: rep,
        treatment_id: tid,
        amendment_name: amend,
        isControl: isCtrl,
        yield_t_ha: y,
        total_cost_per_ha: c,
        delta_yield_t_ha: dy,
        delta_cost_per_ha: dc,
        _rep_ok: !!hasBase
      });
    }

    // treatment-level stats (excluding controls from treatment list; but keep a control summary too)
    const groups = new Map(); // key -> {name,isControl, ys, cs, dys, dcs, n, nDelta}
    function keyOf(p){
      // robust key: treatment_id if finite else amendment_name
      if(Number.isFinite(p.treatment_id)) return `tid:${p.treatment_id}`;
      return `name:${(p.amendment_name||"").toLowerCase()}`;
    }

    for(const p of plot){
      const k = keyOf(p);
      let g = groups.get(k);
      if(!g){
        g = { key:k, name: p.amendment_name || k, isControl: !!p.isControl, ys:[], cs:[], dys:[], dcs:[], n:0, nDelta:0 };
        groups.set(k,g);
      }
      g.n++;
      if(Number.isFinite(p.yield_t_ha)) g.ys.push(p.yield_t_ha);
      if(Number.isFinite(p.total_cost_per_ha)) g.cs.push(p.total_cost_per_ha);
      if(p._rep_ok){
        if(Number.isFinite(p.delta_yield_t_ha)) g.dys.push(p.delta_yield_t_ha);
        if(Number.isFinite(p.delta_cost_per_ha)) g.dcs.push(p.delta_cost_per_ha);
        if(Number.isFinite(p.delta_yield_t_ha) || Number.isFinite(p.delta_cost_per_ha)) g.nDelta++;
      }
      // if any row indicates control, keep as control
      if(p.isControl) g.isControl = true;
      // best effort name
      if(p.amendment_name && p.amendment_name.trim().length) g.name = p.amendment_name.trim();
    }

    const stats = [];
    for(const g of groups.values()){
      stats.push({
        key: g.key,
        name: g.isControl ? "Control" : g.name,
        isControl: g.isControl,
        n_plots: g.n,
        mean_yield_t_ha: mean(g.ys),
        sd_yield_t_ha: sd(g.ys),
        mean_cost_per_ha: mean(g.cs),
        sd_cost_per_ha: sd(g.cs),
        mean_delta_yield_t_ha: mean(g.dys),
        sd_delta_yield_t_ha: sd(g.dys),
        mean_delta_cost_per_ha: mean(g.dcs),
        sd_delta_cost_per_ha: sd(g.dcs),
        n_delta_plots: g.dys.length || g.dcs.length ? Math.max(g.dys.length, g.dcs.length) : 0
      });
    }

    // normalise control label
    const control = stats.find(s=>s.isControl) || null;
    for(const s of stats){
      if(s.isControl) s.name = "Control";
    }

    // initialise configuration per treatment (control read-only)
    const sorted = stats.slice().sort((a,b)=>{
      if(a.isControl && !b.isControl) return -1;
      if(!a.isControl && b.isControl) return 1;
      return a.name.localeCompare(b.name);
    });

    for(const s of sorted){
      if(!state.config.has(s.key)){
        state.config.set(s.key, {
          include: true,
          recurrence: s.isControl ? "one_off" : "one_off",
          customCostPath: null
        });
      }
    }

    state.treatmentStats = { plot, stats: sorted, controlKey: control?.key || null, scaling: rec.scalingSummary };
    state.committed = true;

    // enable UI
    const enableIds = [
      "#faba-export-clean","#faba-apply-config","#faba-view-summary","#faba-export-treatments","#faba-export-grid",
      "#faba-filter-topnpv","#faba-filter-topbcr","#faba-filter-improve","#faba-filter-all",
      "#faba-copy-prompt","#faba-copy-json","#faba-save-scenario"
    ];
    for(const id of enableIds){
      const el = $(id);
      if(el) el.disabled = false;
    }

    renderConfigTable();
    renderGlance();
    recalcAndRender();
    toast("Dataset committed. Results and all tabs updated.");
  }

  // ---------- PV ENGINE ----------
  function adjustVectorToHorizon(vec, T){
    const v = Array.isArray(vec) ? vec.slice() : [];
    const clean = v.map(x=>Number(x)).map(x=>Number.isFinite(x)?x:0);
    if(clean.length === T) return { vec: clean, warn: "" };
    if(clean.length > T) return { vec: clean.slice(0,T), warn: `Persistence vector longer than horizon; truncated to ${T}.` };
    // extend with zeros
    const out = clean.concat(new Array(T - clean.length).fill(0));
    return { vec: out, warn: `Persistence vector shorter than horizon; extended with zeros to ${T}.` };
  }

  function pvBenefits(deltaY, price, r, fVec){
    let pv = 0;
    for(let t=1;t<=fVec.length;t++){
      pv += (deltaY * price * fVec[t-1]) / Math.pow(1+r, t);
    }
    return pv;
  }
  function pvCosts(deltaC, r, T, recurrence, customPath){
    if(!Number.isFinite(deltaC)) return NaN;
    if(recurrence === "annual"){
      let pv = 0;
      for(let t=1;t<=T;t++){
        pv += deltaC / Math.pow(1+r, t);
      }
      return pv;
    }
    if(recurrence === "custom" && Array.isArray(customPath)){
      const { vec } = adjustVectorToHorizon(customPath, T);
      let pv = 0;
      for(let t=1;t<=T;t++){
        pv += (deltaC * vec[t-1]) / Math.pow(1+r, t);
      }
      return pv;
    }
    // one_off upfront
    return deltaC;
  }

  function computeScenarioSlice(price, r, persistenceName){
    if(!state.treatmentStats) return null;
    const T = Math.max(1, Math.floor(Number(state.scenario.horizon) || 10));
    const pers = state.scenario.persistence[persistenceName];
    const adj = adjustVectorToHorizon(pers, T);

    const control = state.treatmentStats.stats.find(s=>s.isControl);
    const controlLevel = {
      pv_benefits: 0,
      pv_costs: 0,
      npv: 0,
      bcr: NaN,
      roi: NaN,
      rank: 0
    };

    const results = [];
    for(const s of state.treatmentStats.stats){
      const cfg = state.config.get(s.key) || { include:true, recurrence:"one_off", customCostPath:null };
      const include = !!cfg.include || s.isControl;

      const dy = s.isControl ? 0 : s.mean_delta_yield_t_ha;
      const dc = s.isControl ? 0 : s.mean_delta_cost_per_ha;

      const pvb = s.isControl ? 0 : pvBenefits(dy, price, r, adj.vec);
      const pvc = s.isControl ? 0 : pvCosts(dc, r, T, cfg.recurrence, cfg.customCostPath);

      const npv = (Number.isFinite(pvb) && Number.isFinite(pvc)) ? (pvb - pvc) : NaN;

      const bcr = (Number.isFinite(pvb) && Number.isFinite(pvc) && pvc > 0) ? (pvb / pvc) : NaN;
      const roi = (Number.isFinite(npv) && Number.isFinite(pvc) && pvc > 0) ? (npv / pvc) : NaN;

      results.push({
        key: s.key,
        name: s.isControl ? "Control (baseline)" : s.name,
        isControl: s.isControl,
        include,
        recurrence: s.isControl ? "one_off" : cfg.recurrence,
        pv_benefits: pvb,
        pv_costs: pvc,
        npv,
        bcr,
        roi
      });
    }

    // ranking (only included, non-control)
    const rankable = results.filter(x=>!x.isControl && x.include && Number.isFinite(x.npv)).sort((a,b)=>b.npv-a.npv);
    const rankMap = new Map();
    for(let i=0;i<rankable.length;i++) rankMap.set(rankable[i].key, i+1);

    for(const r0 of results){
      if(r0.isControl){ r0.rank = 0; continue; }
      r0.rank = rankMap.has(r0.key) ? rankMap.get(r0.key) : null;
    }

    // control for deltas: always 0 baseline
    const ctl = results.find(x=>x.isControl) || null;

    for(const r0 of results){
      const baseNPV = ctl ? 0 : 0;
      const basePVC = ctl ? 0 : 0;
      const basePVB = ctl ? 0 : 0;

      const dNpv = Number.isFinite(r0.npv) ? (r0.npv - baseNPV) : NaN;
      const dPvc = Number.isFinite(r0.pv_costs) ? (r0.pv_costs - basePVC) : NaN;
      const dPvb = Number.isFinite(r0.pv_benefits) ? (r0.pv_benefits - basePVB) : NaN;

      r0.delta_npv = dNpv;
      r0.delta_pv_costs = dPvc;
      r0.delta_pv_benefits = dPvb;

      r0.delta_npv_pct = (Number.isFinite(baseNPV) && baseNPV !== 0 && Number.isFinite(dNpv)) ? (dNpv/baseNPV) : NaN;
      r0.delta_pv_costs_pct = (Number.isFinite(basePVC) && basePVC !== 0 && Number.isFinite(dPvc)) ? (dPvc/basePVC) : NaN;
      r0.delta_pv_benefits_pct = (Number.isFinite(basePVB) && basePVB !== 0 && Number.isFinite(dPvb)) ? (dPvb/basePVB) : NaN;
    }

    return { T, price, r, persistenceName, persistenceWarn: adj.warn, results };
  }

  function computeFullSensitivityGrid(){
    if(!state.treatmentStats) return [];
    const out = [];
    const T = Math.max(1, Math.floor(Number(state.scenario.horizon) || 10));
    for(const price of state.scenario.prices){
      for(const r of state.scenario.discounts){
        for(const pn of Object.keys(state.scenario.persistence)){
          const slice = computeScenarioSlice(Number(price), Number(r), pn);
          if(!slice) continue;
          for(const res of slice.results){
            if(res.isControl) continue;
            out.push({
              price_aud_per_t: Number(price),
              discount_rate: Number(r),
              persistence: pn,
              horizon_years: T,
              treatment: res.name,
              include: res.include,
              recurrence: res.recurrence,
              pv_benefits_per_ha: res.pv_benefits,
              pv_costs_per_ha: res.pv_costs,
              npv_per_ha: res.npv,
              bcr: res.bcr,
              roi: res.roi,
              rank_by_npv: res.rank
            });
          }
        }
      }
    }
    return out;
  }

  // ---------- RENDERING ----------
  const fmt = (n)=> (Number.isFinite(n) ? n.toLocaleString(undefined,{ maximumFractionDigits: 2 }) : "Not applicable");
  const money = (n)=> (Number.isFinite(n) ? ("$"+n.toLocaleString(undefined,{ maximumFractionDigits: 2 })) : "Not applicable");

  function diagItem(level, text){
    return `<div class="diag-item ${level}">${text}</div>`;
  }

  function renderDiagnostics(){
    const root = $("#faba-diagnostics");
    if(!root) return;
    const v = state.validation;
    if(!v){
      root.innerHTML = diagItem("warn","No data loaded yet. Upload or paste the dataset, then validate.");
      return;
    }
    root.innerHTML = v.issues.map(x=>diagItem(x.level, x.msg)).join("");
  }

  function renderDataChecks(){
    const root = $("#faba-data-checks");
    if(!root) return;
    const s = state.validation?.scaling || [];
    if(!s.length){
      root.innerHTML = diagItem("good","No amendments triggered the scaling rule (raw amendment input cost per hectare above 1000).");
      return;
    }
    const rows = s.slice(0,30).map(x=>{
      return `<div class="diag-item warn">
        <div><strong>${x.amendment_name}</strong> scaled in ${x.count} plot(s).</div>
        <div class="small muted">Raw mean ${money(x.raw_mean)} (range ${money(x.raw_min)} to ${money(x.raw_max)}). Adjusted mean ${money(x.adj_mean)} (range ${money(x.adj_min)} to ${money(x.adj_max)}).</div>
      </div>`;
    }).join("");
    root.innerHTML = rows + (s.length>30 ? diagItem("warn",`Showing first 30 amendments. Total amendments scaled: ${s.length}.`) : "");
  }

  function renderPreviewTable(){
    const tbl = $("#faba-preview");
    if(!tbl) return;
    const thead = tbl.querySelector("thead");
    const tbody = tbl.querySelector("tbody");
    if(!thead || !tbody) return;

    tbody.innerHTML = "";
    thead.innerHTML = "";

    if(!state.rows.length || !state.columns.length){
      thead.innerHTML = "<tr><th>No data</th></tr>";
      return;
    }
    const cols = state.columns.slice(0, Math.min(8, state.columns.length));
    const header = cols.map(c=>{
      const tip = colTip(c);
      return `<th ${tip? `data-tip="${escapeHtml(tip)}"`:""}>${escapeHtml(colLabel(c))}</th>`;
    }).join("");
    thead.innerHTML = `<tr>${header}</tr>`;

    for(const r of state.rows.slice(0,10)){
      const row = cols.map(c=>`<td>${escapeHtml(String(r[c] ?? ""))}</td>`).join("");
      tbody.innerHTML += `<tr>${row}</tr>`;
    }
  }

  function escapeHtml(s){
    return (s??"").toString().replace(/[&<>"']/g, c=>({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[c]));
  }

  function renderConfigTable(){
    const tbody = $("#faba-config tbody");
    if(!tbody || !state.treatmentStats) return;
    tbody.innerHTML = "";

    for(const s of state.treatmentStats.stats){
      const cfg = state.config.get(s.key) || { include:true, recurrence:"one_off", customCostPath:null };
      const isControl = s.isControl;

      const include = isControl ? true : !!cfg.include;
      const recur = isControl ? "one_off" : (cfg.recurrence || "one_off");

      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(isControl ? "Control (baseline)" : s.name)}</td>
        <td>${isControl ? '<span class="pill">Yes</span>' : ''}</td>
        <td>
          <input type="checkbox" data-faba-include="${escapeHtml(s.key)}" ${include ? "checked":""} ${isControl ? "disabled":""} />
        </td>
        <td>
          <select data-faba-recur="${escapeHtml(s.key)}" ${isControl ? "disabled":""}>
            <option value="one_off" ${recur==="one_off"?"selected":""}>One-off upfront</option>
            <option value="annual" ${recur==="annual"?"selected":""}>Annual</option>
            <option value="custom" ${recur==="custom"?"selected":""}>Custom</option>
          </select>
        </td>
        <td class="right mono">${Number(s.n_plots||0).toLocaleString()}</td>
      `;
      tbody.appendChild(tr);
    }

    tbody.onchange = (e)=>{
      const incKey = e.target?.dataset?.fabaInclude;
      const recKey = e.target?.dataset?.fabaRecur;

      if(incKey){
        const cfg = state.config.get(incKey) || { include:true, recurrence:"one_off", customCostPath:null };
        cfg.include = !!e.target.checked;
        state.config.set(incKey, cfg);
        toast("Configuration updated.");
        recalcAndRender();
        renderGlance();
        return;
      }
      if(recKey){
        const cfg = state.config.get(recKey) || { include:true, recurrence:"one_off", customCostPath:null };
        cfg.recurrence = String(e.target.value || "one_off");
        if(cfg.recurrence === "custom" && !Array.isArray(cfg.customCostPath)){
          // default custom path: treat like annual by default
          cfg.customCostPath = new Array(Math.max(1, Math.floor(Number(state.scenario.horizon)||10))).fill(1);
        }
        state.config.set(recKey, cfg);
        toast("Cost recurrence updated.");
        recalcAndRender();
        renderGlance();
      }
    };
  }

  function renderGlance(){
    const root = $("#faba-glance");
    if(!root) return;

    const repCount = state.replicateStats?.repMap?.size || 0;
    const repsNoCtrl = state.replicateStats?.repsWithNoCtrl || 0;
    const nRows = state.rows.length;
    const stats = state.treatmentStats?.stats || [];
    const nTrt = stats.filter(s=>!s.isControl).length;
    const nCtrl = stats.filter(s=>s.isControl).length;

    root.innerHTML = `
      <div><span class="k">Price</span><span class="v">${money(Number(state.scenario.selectedPrice))} per tonne</span></div>
      <div><span class="k">Discount rate</span><span class="v">${(Number(state.scenario.selectedDiscount)*100).toFixed(2)}%</span></div>
      <div><span class="k">Persistence</span><span class="v">${escapeHtml(state.scenario.selectedPersistence)}</span></div>
      <div><span class="k">Horizon</span><span class="v">${Math.max(1,Math.floor(Number(state.scenario.horizon)||10))} years</span></div>
      <div><span class="k">Plots</span><span class="v">${nRows.toLocaleString()}</span></div>
      <div><span class="k">Replicates</span><span class="v">${repCount.toLocaleString()}${repsNoCtrl? ` (missing controls: ${repsNoCtrl})`:""}</span></div>
      <div><span class="k">Treatments</span><span class="v">${nTrt.toLocaleString()}</span></div>
      <div><span class="k">Controls</span><span class="v">${nCtrl.toLocaleString()}</span></div>
    `;
    const trtEl = $("#faba-trt");
    if(trtEl) trtEl.textContent = String(nTrt);
  }

  function renderScenarioSelectors(){
    const priceSel = $("#faba-price");
    const discSel = $("#faba-discount");
    const persSel = $("#faba-persist");
    const persJson = $("#faba-persist-json");
    if(persJson){
      persJson.value = JSON.stringify(state.scenario.persistence, null, 2);
    }

    if(priceSel){
      priceSel.innerHTML = state.scenario.prices.map(p=>`<option value="${p}" ${Number(p)===Number(state.scenario.selectedPrice)?"selected":""}>${p}</option>`).join("");
    }
    if(discSel){
      discSel.innerHTML = state.scenario.discounts.map(d=>`<option value="${d}" ${Number(d)===Number(state.scenario.selectedDiscount)?"selected":""}>${(Number(d)*100).toFixed(2)}%</option>`).join("");
    }
    if(persSel){
      persSel.innerHTML = Object.keys(state.scenario.persistence).map(k=>`<option value="${escapeHtml(k)}" ${k===state.scenario.selectedPersistence?"selected":""}>${escapeHtml(k)}</option>`).join("");
    }
  }

  function renderLeaderboard(slice){
    const root = $("#faba-leaderboard");
    if(!root) return;
    const scale = Number(state.scenario.scale)||1;

    const rows = slice.results
      .filter(x=>!x.isControl)
      .filter(x=>state.filters.mode==="all" ? true : state.filters.mode==="improve" ? (Number.isFinite(x.npv) && x.npv>0) : true);

    let shown = rows.slice();
    if(state.filters.mode==="topnpv"){
      shown = shown.filter(x=>x.include).sort((a,b)=>(Number.isFinite(b.npv)?b.npv:-Infinity)-(Number.isFinite(a.npv)?a.npv:-Infinity)).slice(0,5);
    } else if(state.filters.mode==="topbcr"){
      shown = shown.filter(x=>x.include).sort((a,b)=>(Number.isFinite(b.bcr)?b.bcr:-Infinity)-(Number.isFinite(a.bcr)?a.bcr:-Infinity)).slice(0,5);
    } else if(state.filters.mode==="improve"){
      shown = shown.filter(x=>x.include && Number.isFinite(x.npv) && x.npv>0);
    } else {
      shown = shown.sort((a,b)=>{
        const ra = (a.rank==null)? 1e9 : a.rank;
        const rb = (b.rank==null)? 1e9 : b.rank;
        return ra-rb;
      });
    }

    root.innerHTML = shown.map(x=>{
      const npv = x.npv*scale;
      const pvb = x.pv_benefits*scale;
      const pvc = x.pv_costs*scale;
      const rank = (x.rank==null) ? "Not ranked" : String(x.rank);
      const cls = (Number.isFinite(x.npv) && x.npv>=0) ? "good" : "bad";
      return `
        <div class="lb-row">
          <div><div class="h">Rank</div><div class="b">${rank}</div></div>
          <div><div class="h">Treatment</div><div class="b">${escapeHtml(x.name)}</div></div>
          <div><div class="h">Net present value</div><div class="b mono ${cls}">${money(npv)}</div></div>
          <div><div class="h">Net present value for 100 hectares</div><div class="b mono">${money(x.npv*100)}</div></div>
          <div><div class="h">Net present value for 3300 hectares</div><div class="b mono">${money(x.npv*3300)}</div></div>
          <div><div class="h">Present value of benefits</div><div class="b mono">${money(pvb)}</div></div>
          <div><div class="h">Present value of costs</div><div class="b mono">${money(pvc)}</div></div>
        </div>
      `;
    }).join("");

    if(!shown.length){
      root.innerHTML = `<div class="diag-item warn">No treatments match the current filter.</div>`;
    }
  }

  function renderComparisonTable(slice){
    const table = $("#faba-compare");
    if(!table) return;
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");
    if(!thead || !tbody) return;

    const scale = Number(state.scenario.scale)||1;
    const control = slice.results.find(x=>x.isControl);

    // choose displayed treatments based on filter
    let treatments = slice.results.filter(x=>!x.isControl);
    if(state.filters.mode==="topnpv"){
      treatments = treatments.filter(x=>x.include).sort((a,b)=>(Number.isFinite(b.npv)?b.npv:-Infinity)-(Number.isFinite(a.npv)?a.npv:-Infinity)).slice(0,5);
    } else if(state.filters.mode==="topbcr"){
      treatments = treatments.filter(x=>x.include).sort((a,b)=>(Number.isFinite(b.bcr)?b.bcr:-Infinity)-(Number.isFinite(a.bcr)?a.bcr:-Infinity)).slice(0,5);
    } else if(state.filters.mode==="improve"){
      treatments = treatments.filter(x=>x.include && Number.isFinite(x.npv) && x.npv>0);
    } else {
      treatments = treatments.sort((a,b)=>{
        const ra = (a.rank==null)? 1e9 : a.rank;
        const rb = (b.rank==null)? 1e9 : b.rank;
        return ra-rb;
      });
    }

    const cols = [control].concat(treatments).filter(Boolean);

    const colHeads = cols.map(x=>{
      const tip = x.isControl ? "Baseline. All treatment results are incremental relative to the replicate-specific control means." : `Cost recurrence: ${x.recurrence}. Included in ranking: ${x.include ? "Yes":"No"}.`;
      return `<th data-tip="${escapeHtml(tip)}">${escapeHtml(x.name)}</th>`;
    }).join("");

    // For each treatment, add a delta vs control column
    const deltaHeads = treatments.map(x=>{
      return `<th data-tip="Difference from control under the selected scenario. Percent is shown only where meaningful.">Δ vs control</th>`;
    }).join("");

    thead.innerHTML = `<tr><th>Indicator</th>${colHeads}${deltaHeads}</tr>`;

    const indicators = [
      { key:"pv_benefits", label:"Present value of benefits", fmt:(v)=>money(v*scale), good:(v)=>Number.isFinite(v) && v>=0 },
      { key:"pv_costs", label:"Present value of costs", fmt:(v)=>money(v*scale), good:(v)=>Number.isFinite(v) && v<=0 ? false : true },
      { key:"npv", label:"Net present value", fmt:(v)=>money(v*scale), good:(v)=>Number.isFinite(v) && v>=0 },
      { key:"bcr", label:"Benefit–cost ratio", fmt:(v)=> (Number.isFinite(v)? fmt(v) : "Not applicable"), good:(v)=>Number.isFinite(v) && v>=1 },
      { key:"roi", label:"Return on investment", fmt:(v)=> (Number.isFinite(v)? fmt(v) : "Not applicable"), good:(v)=>Number.isFinite(v) && v>=0 },
      { key:"rank", label:"Rank by net present value", fmt:(v)=> (v==null? "Not ranked" : String(v)), good:(_)=>true }
    ];

    tbody.innerHTML = "";
    for(const ind of indicators){
      const row = document.createElement("tr");
      let tds = `<td class="mono"><strong>${escapeHtml(ind.label)}</strong></td>`;

      for(const c of cols){
        const v = c[ind.key];
        const cls = c.isControl ? "neu" : (ind.good(v) ? "good" : "bad");
        tds += `<td class="mono"><span class="val ${cls}">${ind.fmt(v)}</span></td>`;
      }

      // delta columns
      for(const t of treatments){
        let deltaAbs = NaN, deltaPct = NaN, base = 0;
        if(ind.key === "npv"){ deltaAbs = t.delta_npv; base = 0; }
        if(ind.key === "pv_costs"){ deltaAbs = t.delta_pv_costs; base = 0; }
        if(ind.key === "pv_benefits"){ deltaAbs = t.delta_pv_benefits; base = 0; }

        // Only percent where meaningful (non-zero baseline); baseline is 0 here because the PVs are incremental vs control by construction.
        // Therefore show only absolute deltas and label percent as Not applicable.
        const absTxt = Number.isFinite(deltaAbs) ? money(deltaAbs*scale) : "Not applicable";
        const pctTxt = "Not applicable";

        let cls = "neu";
        if(ind.key === "pv_costs" && Number.isFinite(deltaAbs)){
          // lower costs favourable
          cls = deltaAbs <= 0 ? "good" : "bad";
        } else if((ind.key === "npv" || ind.key === "pv_benefits") && Number.isFinite(deltaAbs)){
          cls = deltaAbs >= 0 ? "good" : "bad";
        }

        if(ind.key === "npv" || ind.key === "pv_costs" || ind.key === "pv_benefits"){
          tds += `<td class="mono"><span class="val ${cls}">${absTxt}</span><div class="small muted">${pctTxt}</div></td>`;
        } else {
          tds += `<td class="mono"><span class="val neu">Not applicable</span></td>`;
        }
      }

      row.innerHTML = tds;
      tbody.appendChild(row);
    }
  }

  function buildMeaningText(slice){
    const root = $("#faba-meaning");
    if(!root) return;

    const price = Number(slice.price);
    const rPct = (Number(slice.r)*100);
    const pers = slice.persistenceName;

    const trt = slice.results.filter(x=>!x.isControl && x.include && Number.isFinite(x.npv));
    const top = trt.slice().sort((a,b)=>b.npv-a.npv)[0] || null;
    const bottom = trt.slice().sort((a,b)=>a.npv-b.npv)[0] || null;

    let txt = "";
    txt += `Under the selected settings, the tool compares each amendment package to the replicate-specific control plots. The selected grain price is ${money(price)} per tonne, the discount rate is ${rPct.toFixed(2)} percent, the yield persistence pattern is ${pers}, and the time horizon is ${slice.T} years. `;

    if(slice.persistenceWarn){
      txt += `A persistence adjustment was applied because the persistence vector did not match the horizon. ${slice.persistenceWarn} `;
    }

    if(!trt.length){
      txt += `No treatments are currently eligible for ranking under this slice, either because they are excluded in Configuration or because there is insufficient numeric data to compute net present value. `;
      root.textContent = txt;
      return;
    }

    txt += `The strongest performer by net present value in this slice is ${top.name}. Its net present value is ${money(top.npv)} per hectare, which reflects present value of benefits of ${money(top.pv_benefits)} per hectare and present value of costs of ${money(top.pv_costs)} per hectare. `;
    txt += `The weakest performer in this slice is ${bottom.name}, with net present value of ${money(bottom.npv)} per hectare. `;

    const drivers = (x)=>{
      if(!Number.isFinite(x.pv_benefits) || !Number.isFinite(x.pv_costs)) return "";
      if(x.pv_benefits <= 0 && x.pv_costs > 0) return "low or negative yield benefit combined with added cost";
      if(x.pv_benefits > 0 && x.pv_costs <= 0) return "benefits with little or negative added cost";
      if(x.pv_benefits > x.pv_costs) return "benefits that outweigh costs";
      return "costs that outweigh benefits";
    };

    txt += `In practical terms, treatments look better than control when the additional yield benefit persists for more years or when costs are one-off rather than repeated annually. In this slice, ${top.name} is mainly driven by ${drivers(top)}, while ${bottom.name} is mainly driven by ${drivers(bottom)}. `;
    txt += `If grain price increases, the value of any positive yield change rises proportionally, which can move treatments up the ranking. If the persistence pattern decays quickly, only treatments with strong first-year yield benefits tend to remain attractive. If you switch a treatment’s cost recurrence from one-off to annual, present value of costs rises over time and net present value can fall substantially. `;

    root.textContent = txt;
  }

  function buildAI(slice){
    const promptEl = $("#faba-ai-prompt");
    if(!promptEl) return;

    const scale = Number(state.scenario.scale)||1;

    const ranked = slice.results.filter(x=>!x.isControl && x.include).slice().sort((a,b)=>{
      const na = Number.isFinite(a.npv) ? a.npv : -Infinity;
      const nb = Number.isFinite(b.npv) ? b.npv : -Infinity;
      return nb-na;
    });

    const top = ranked.slice(0,5);
    const worst = ranked.slice(-3);

    const summaryLine = (x)=> {
      return `${x.name} has net present value of ${money(x.npv*scale)} with present value of benefits of ${money(x.pv_benefits*scale)} and present value of costs of ${money(x.pv_costs*scale)} under the selected recurrence setting.`;
    };

    const p = Number(slice.price);
    const rPct = (Number(slice.r)*100);
    const pers = slice.persistenceName;
    const T = slice.T;

    let text = "";
    text += `Write a farmer-facing policy brief about a faba bean soil amendment trial. Use plain language and do not use bullet points or abbreviations. Use short paragraphs. The brief must be based only on the results provided below and must not introduce new assumptions. `;
    text += `The analysis compares each amendment treatment to replicate-specific control plots. The selected scenario settings are a grain price of ${money(p)} per tonne, a discount rate of ${rPct.toFixed(2)} percent, a yield persistence pattern of ${pers}, and a time horizon of ${T} years. Costs are handled using the recurrence setting shown for each treatment. `;
    text += `Explain what is driving differences across treatments by separating benefits and costs, and explain how conclusions change if grain price, persistence, or cost recurrence changes. Include uncertainty and data quality caveats, especially where replicates lack controls or numeric fields have missing values. `;
    text += `Now use these computed results. `;

    if(top.length){
      text += `The top treatments by net present value in this slice are as follows. `;
      for(const x of top){
        text += summaryLine(x) + " ";
      }
    }
    if(worst.length){
      text += `Treatments that perform worst in this slice include the following. `;
      for(const x of worst){
        text += summaryLine(x) + " ";
      }
    }

    text += `Also include a section that translates results to total impacts for 100 hectares and for 3300 hectares using the same per-hectare net present values. `;
    promptEl.value = text;

    // JSON
    const payload = {
      scenario: {
        price_aud_per_t: p,
        discount_rate: Number(slice.r),
        persistence: pers,
        horizon_years: T,
        scale: Number(state.scenario.scale)||1
      },
      treatments: slice.results.map(x=>({
        name: x.name,
        is_control: x.isControl,
        include: x.include,
        recurrence: x.recurrence,
        pv_benefits_per_ha: x.pv_benefits,
        pv_costs_per_ha: x.pv_costs,
        npv_per_ha: x.npv,
        bcr: x.bcr,
        roi: x.roi,
        rank_by_npv: x.rank,
        delta_npv_per_ha: x.delta_npv,
        delta_pv_benefits_per_ha: x.delta_pv_benefits,
        delta_pv_costs_per_ha: x.delta_pv_costs
      }))
    };
    state._latestJSON = payload;
  }

  function recalcAndRender(){
    if(!state.committed || !state.treatmentStats){
      return;
    }
    const slice = computeScenarioSlice(
      Number(state.scenario.selectedPrice),
      Number(state.scenario.selectedDiscount),
      String(state.scenario.selectedPersistence)
    );
    if(!slice) return;

    renderGlance();
    renderLeaderboard(slice);
    renderComparisonTable(slice);
    buildMeaningText(slice);
    buildAI(slice);
  }

  // ---------- EXPORTS ----------
  function toTSV(rows, columns){
    const header = columns.join("\t");
    const lines = [header];
    for(const r of rows){
      const line = columns.map(c=>String(r[c]??"")).join("\t");
      lines.push(line);
    }
    return lines.join("\n");
  }
  function toCSV(rows, columns){
    const esc = (v)=>{
      const s = String(v??"");
      if(/[",\n]/.test(s)) return `"${s.replace(/"/g,'""')}"`;
      return s;
    };
    const header = columns.map(esc).join(",");
    const lines = [header];
    for(const r of rows){
      lines.push(columns.map(c=>esc(r[c])).join(","));
    }
    return lines.join("\n");
  }

  function exportCleanedTSV(){
    if(!state.rows.length) return;
    // cleaned means: same columns, values preserved; missing markers remain as blanks; we do not drop columns
    const cleaned = state.rows.map(r=>{
      const o = {};
      for(const c of state.columns){
        const v = r[c];
        o[c] = (isMissing(v) ? "" : v);
      }
      return o;
    });
    downloadFile("faba_beans_cleaned_export.tsv", toTSV(cleaned, state.columns), "text/tab-separated-values");
    toast("Cleaned dataset exported as TSV.");
  }

  function exportTreatmentSummaryCSV(){
    if(!state.treatmentStats) return;
    const cols = [
      "treatment","is_control","n_plots","mean_yield_t_ha","sd_yield_t_ha","mean_cost_per_ha","sd_cost_per_ha",
      "mean_delta_yield_t_ha","sd_delta_yield_t_ha","mean_delta_cost_per_ha","sd_delta_cost_per_ha"
    ];
    const rows = state.treatmentStats.stats.map(s=>({
      treatment: s.isControl ? "Control" : s.name,
      is_control: s.isControl ? 1 : 0,
      n_plots: s.n_plots,
      mean_yield_t_ha: s.mean_yield_t_ha,
      sd_yield_t_ha: s.sd_yield_t_ha,
      mean_cost_per_ha: s.mean_cost_per_ha,
      sd_cost_per_ha: s.sd_cost_per_ha,
      mean_delta_yield_t_ha: s.isControl ? 0 : s.mean_delta_yield_t_ha,
      sd_delta_yield_t_ha: s.sd_delta_yield_t_ha,
      mean_delta_cost_per_ha: s.isControl ? 0 : s.mean_delta_cost_per_ha,
      sd_delta_cost_per_ha: s.sd_delta_cost_per_ha
    }));
    downloadFile("faba_beans_treatment_summary.csv", toCSV(rows, cols), "text/csv");
    toast("Treatment summary exported as CSV.");
  }

  function exportSensitivityGridCSV(){
    const grid = computeFullSensitivityGrid();
    const cols = [
      "price_aud_per_t","discount_rate","persistence","horizon_years","treatment","include","recurrence",
      "pv_benefits_per_ha","pv_costs_per_ha","npv_per_ha","bcr","roi","rank_by_npv"
    ];
    downloadFile("faba_beans_sensitivity_grid.csv", toCSV(grid, cols), "text/csv");
    toast("Sensitivity grid exported as CSV.");
  }

  // ---------- SCENARIO SAVE/LOAD ----------
  const SCEN_KEY = "faba_cba_scenario_v1";
  function saveScenario(){
    const payload = {
      scenario: state.scenario,
      config: Array.from(state.config.entries())
    };
    localStorage.setItem(SCEN_KEY, JSON.stringify(payload));
    toast("Scenario saved to local storage.");
  }
  function loadScenario(){
    const raw = localStorage.getItem(SCEN_KEY);
    if(!raw){
      toast("No saved scenario found.");
      return;
    }
    try{
      const obj = JSON.parse(raw);
      if(obj?.scenario){
        state.scenario = obj.scenario;
      }
      if(Array.isArray(obj?.config)){
        state.config = new Map(obj.config);
      }
      // normalise persistence json if missing
      if(!state.scenario.persistence || typeof state.scenario.persistence !== "object"){
        state.scenario.persistence = JSON.parse(JSON.stringify(DEFAULT.persistence));
      }
      renderScenarioSelectors();
      renderConfigTable();
      renderGlance();
      recalcAndRender();
      toast("Scenario loaded.");
    }catch(e){
      toast("Saved scenario could not be read.");
    }
  }

  // ---------- EVENTS ----------
  async function readFileToText(input, assign){
    const file = input?.files?.[0];
    if(!file) return;
    const text = await file.text();
    assign(text);
  }

  function applyScenarioLists(){
    const pList = ($("#faba-price-list")?.value || "").split(",").map(s=>Number(s.trim())).filter(n=>Number.isFinite(n));
    const dList = ($("#faba-discount-list")?.value || "").split(",").map(s=>Number(s.trim())).filter(n=>Number.isFinite(n));
    let pers = null;
    try{
      pers = JSON.parse($("#faba-persist-json")?.value || "{}");
    }catch(_){
      pers = null;
    }
    const persOk = pers && typeof pers === "object" && Object.keys(pers).length>0 && Object.values(pers).every(v=>Array.isArray(v));
    if(!pList.length || !dList.length || !persOk){
      toast("Scenario lists are invalid. Ensure prices, discount rates, and persistence JSON are valid.");
      return;
    }
    state.scenario.prices = pList;
    state.scenario.discounts = dList;
    state.scenario.persistence = pers;
    // reset selections if invalid
    if(!state.scenario.prices.includes(Number(state.scenario.selectedPrice))) state.scenario.selectedPrice = state.scenario.prices[0];
    if(!state.scenario.discounts.includes(Number(state.scenario.selectedDiscount))) state.scenario.selectedDiscount = state.scenario.discounts[0];
    if(!Object.keys(state.scenario.persistence).includes(state.scenario.selectedPersistence)) state.scenario.selectedPersistence = Object.keys(state.scenario.persistence)[0];
    renderScenarioSelectors();
    recalcAndRender();
    toast("Scenario lists applied.");
  }

  function bindUI(){
    // Import
    const file = $("#faba-file");
    const dictFile = $("#faba-dict-file");
    const paste = $("#faba-paste");
    const dictPaste = $("#faba-dict-paste");
    const validateBtn = $("#faba-validate");
    const commitBtn = $("#faba-commit");

    if(file){
      file.addEventListener("change", ()=>readFileToText(file, (t)=>{ state.rawText=t; const p=parseDelimited(t); state.columns=p.columns; state.rows=p.rows; renderPreviewTable(); toast("Dataset loaded. Validate next."); }));
    }
    if(dictFile){
      dictFile.addEventListener("change", ()=>readFileToText(dictFile, (t)=>{ state.dictText=t; state.dict=parseDictionaryCSV(t); toast("Dictionary loaded."); renderPreviewTable(); }));
    }
    if(validateBtn){
      validateBtn.addEventListener("click", ()=>{
        // choose paste over file if paste present
        const pasted = (paste?.value||"").trim();
        if(pasted){
          state.rawText = pasted;
          const p = parseDelimited(pasted);
          state.columns = p.columns;
          state.rows = p.rows;
        }
        const dp = (dictPaste?.value||"").trim();
        if(dp){
          state.dictText = dp;
          state.dict = parseDictionaryCSV(dp);
        }
        validateAndSummarise();
      });
    }
    if(commitBtn){
      commitBtn.addEventListener("click", ()=>commitDataset());
    }

    const exportClean = $("#faba-export-clean");
    if(exportClean) exportClean.addEventListener("click", exportCleanedTSV);

    // Scenario controls
    const horizon = $("#faba-horizon");
    const priceSel = $("#faba-price");
    const discSel = $("#faba-discount");
    const persSel = $("#faba-persist");
    const scaleSel = $("#faba-scale");

    if(horizon){
      horizon.addEventListener("input", ()=>{
        const T = Math.max(1, Math.floor(Number(horizon.value)||10));
        state.scenario.horizon = T;
        // safety: adjust custom cost paths lengths lazily
        toast("Time horizon updated.");
        recalcAndRender();
        renderGlance();
      });
    }
    if(priceSel){
      priceSel.addEventListener("change", ()=>{
        state.scenario.selectedPrice = Number(priceSel.value);
        toast("Price scenario updated.");
        recalcAndRender();
        renderGlance();
      });
    }
    if(discSel){
      discSel.addEventListener("change", ()=>{
        state.scenario.selectedDiscount = Number(discSel.value);
        toast("Discount rate updated.");
        recalcAndRender();
        renderGlance();
      });
    }
    if(persSel){
      persSel.addEventListener("change", ()=>{
        state.scenario.selectedPersistence = String(persSel.value);
        toast("Persistence pattern updated.");
        recalcAndRender();
        renderGlance();
      });
    }
    if(scaleSel){
      scaleSel.addEventListener("change", ()=>{
        state.scenario.scale = Number(scaleSel.value)||1;
        toast("Result scaling updated.");
        recalcAndRender();
      });
    }

    const applyListsBtn = $("#faba-apply-scenario-lists");
    if(applyListsBtn) applyListsBtn.addEventListener("click", applyScenarioLists);

    // Filters
    const topNpv = $("#faba-filter-topnpv");
    const topBcr = $("#faba-filter-topbcr");
    const improve = $("#faba-filter-improve");
    const all = $("#faba-filter-all");

    if(topNpv) topNpv.addEventListener("click", ()=>{ state.filters.mode="topnpv"; toast("Filter applied."); recalcAndRender(); });
    if(topBcr) topBcr.addEventListener("click", ()=>{ state.filters.mode="topbcr"; toast("Filter applied."); recalcAndRender(); });
    if(improve) improve.addEventListener("click", ()=>{ state.filters.mode="improve"; toast("Filter applied."); recalcAndRender(); });
    if(all) all.addEventListener("click", ()=>{ state.filters.mode="all"; toast("Filter cleared."); recalcAndRender(); });

    // Config buttons
    const applyCfg = $("#faba-apply-config");
    const viewSum = $("#faba-view-summary");
    if(applyCfg) applyCfg.addEventListener("click", ()=>{ toast("Configuration applied."); recalcAndRender(); renderGlance(); });
    if(viewSum) viewSum.addEventListener("click", ()=>{ toast("Results summary refreshed."); recalcAndRender(); });

    // Exports
    const expTrt = $("#faba-export-treatments");
    const expGrid = $("#faba-export-grid");
    if(expTrt) expTrt.addEventListener("click", exportTreatmentSummaryCSV);
    if(expGrid) expGrid.addEventListener("click", exportSensitivityGridCSV);

    // Scenario save/load
    const saveBtn = $("#faba-save-scenario");
    const loadBtn = $("#faba-load-scenario");
    if(saveBtn) saveBtn.addEventListener("click", saveScenario);
    if(loadBtn) loadBtn.addEventListener("click", loadScenario);

    // Copy buttons
    const copyPrompt = $("#faba-copy-prompt");
    const copyJson = $("#faba-copy-json");
    if(copyPrompt){
      copyPrompt.addEventListener("click", async ()=>{
        const val = $("#faba-ai-prompt")?.value || "";
        await navigator.clipboard.writeText(val);
        toast("Briefing prompt copied.");
      });
    }
    if(copyJson){
      copyJson.addEventListener("click", async ()=>{
        const txt = JSON.stringify(state._latestJSON || {}, null, 2);
        await navigator.clipboard.writeText(txt);
        toast("Results JSON copied.");
      });
    }

    // Init scenario selectors
    renderScenarioSelectors();
    renderGlance();
  }

  // ---------- INIT ----------
  document.addEventListener("DOMContentLoaded", ()=>{
    // Only bind if the panel exists
    if($("#tab-faba") || $("#faba-validate")){
      // seed persistence JSON into textarea if present
      const persJson = $("#faba-persist-json");
      if(persJson && !persJson.value.trim()){
        persJson.value = JSON.stringify(state.scenario.persistence, null, 2);
      }
      bindUI();
    }
  });

})();
