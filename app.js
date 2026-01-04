// Farming CBA Tool - Newcastle Business School
// Fully upgraded script with working tabs, robust import pipeline (upload + paste TSV/CSV + dictionary parsing),
// replicate-specific control baselines, plot-level deltas, treatment summaries with missing-safe stats,
// control-centric Results (leaderboard + comparison-to-control grid + filters + narrative + charts),
// discounted CBA engine + sensitivity grid (price, discount, persistence, recurrence),
// scenario save/load to localStorage, exports (cleaned TSV, summaries CSV, sensitivity CSV, Excel workbook if available),
// AI Briefing (copy-ready narrative prompt with no bullets, no em dash, no abbreviations) + Copy Results JSON,
// Technical Appendix tab opens in a new browser tab, and bottom-right toasts for every major action.

(() => {
  "use strict";

  // =========================
  // 0) CONSTANTS + DEFAULTS
  // =========================
  const DEFAULT_DISCOUNT_SCHEDULE = [
    { label: "2025-2034", from: 2025, to: 2034, low: 2, base: 4, high: 6 },
    { label: "2035-2044", from: 2035, to: 2044, low: 4, base: 7, high: 10 },
    { label: "2045-2054", from: 2045, to: 2054, low: 4, base: 7, high: 10 },
    { label: "2055-2064", from: 2055, to: 2064, low: 3, base: 6, high: 9 },
    { label: "2065-2074", from: 2065, to: 2074, low: 2, base: 5, high: 8 }
  ];

  const horizons = [5, 10, 15, 20, 25];

  const STORAGE_KEYS = {
    scenarios: "farming_cba_scenarios_v1",
    activeScenario: "farming_cba_active_scenario_v1"
  };

  // Default sensitivity grids
  const DEFAULT_SENS_PRICE = [300, 350, 400, 450, 500, 550, 600];
  const DEFAULT_SENS_DISC = [2, 4, 7, 10, 12];
  const DEFAULT_SENS_PERSIST = [1, 2, 3, 5, 7, 10];
  const DEFAULT_SENS_RECURRENCE = [1, 2, 3, 4, 5, 7, 10, 0]; // 0 = once only at year 0

  // Built-in example trial data (tab-separated) so the tool has a working dataset on first load.
  // Replace this constant with your full trial TSV if you want your own dataset to be the default.
  const DEFAULT_TRIAL_DATA_TSV = [
    "Treatment\tReplicate\tPlot\tYield t/ha\tPlot area (ha)\tLabour cost $/ha\tMaterials cost $/ha\tServices cost $/ha\tis_control",
    "Control\tR1\tP1\t2.00\t0.10\t10\t30\t15\t1",
    "Control\tR1\tP2\t2.10\t0.10\t10\t30\t15\t1",
    "Deep ripping\tR1\tP3\t2.60\t0.10\t14\t38\t18\t0",
    "Organic matter\tR1\tP4\t2.55\t0.10\t16\t42\t18\t0",
    "Gypsum\tR1\tP5\t2.40\t0.10\t12\t32\t16\t0",
    "Control\tR2\tP6\t1.90\t0.10\t10\t30\t15\t1",
    "Control\tR2\tP7\t2.00\t0.10\t10\t30\t15\t1",
    "Deep ripping\tR2\tP8\t2.45\t0.10\t14\t38\t18\t0",
    "Organic matter\tR2\tP9\t2.35\t0.10\t16\t42\t18\t0",
    "Gypsum\tR2\tP10\t2.20\t0.10\t12\t32\t16\t0"
  ].join("\n");

  // =========================
  // 1) ID + UTIL
  // =========================
  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

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

  function parseNumber(value) {
    if (value === null || value === undefined) return NaN;
    if (typeof value === "number") return Number.isFinite(value) ? value : NaN;
    const s = String(value).trim();
    if (!s || s === "?" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a") return NaN;
    const cleaned = s.replace(/[\$,]/g, "");
    const n = parseFloat(cleaned);
    return Number.isFinite(n) ? n : NaN;
  }

  function isBlank(v) {
    return v === null || v === undefined || (typeof v === "string" && v.trim() === "") || v === "?";
  }

  function median(arr) {
    const a = arr.filter(v => Number.isFinite(v)).slice().sort((x, y) => x - y);
    if (!a.length) return NaN;
    const mid = Math.floor(a.length / 2);
    return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
  }

  function mean(arr) {
    const a = arr.filter(v => Number.isFinite(v));
    if (!a.length) return NaN;
    return a.reduce((s, v) => s + v, 0) / a.length;
  }

  function sd(arr) {
    const a = arr.filter(v => Number.isFinite(v));
    if (a.length < 2) return NaN;
    const m = mean(a);
    const v = a.reduce((s, x) => s + (x - m) * (x - m), 0) / (a.length - 1);
    return Math.sqrt(v);
  }

  function quantile(arr, q) {
    const a = arr.filter(v => Number.isFinite(v)).slice().sort((x, y) => x - y);
    if (!a.length) return NaN;
    const pos = (a.length - 1) * q;
    const base = Math.floor(pos);
    const rest = pos - base;
    if (a[base + 1] === undefined) return a[base];
    return a[base] + rest * (a[base + 1] - a[base]);
  }

  function iqrOutlierFlags(arr) {
    const a = arr.filter(v => Number.isFinite(v));
    if (a.length < 8) return { low: NaN, high: NaN, outliers: 0 };
    const q1 = quantile(a, 0.25);
    const q3 = quantile(a, 0.75);
    const iqr = q3 - q1;
    const low = q1 - 1.5 * iqr;
    const high = q3 + 1.5 * iqr;
    const outliers = a.filter(v => v < low || v > high).length;
    return { low, high, outliers };
  }

  function annuityFactor(N, rPct) {
    const r = rPct / 100;
    return r === 0 ? N : (1 - Math.pow(1 + r, -N)) / r;
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

  function ensureToastRoot() {
    if (document.getElementById("toast-root")) return;
    const div = document.createElement("div");
    div.id = "toast-root";
    div.setAttribute("aria-live", "polite");
    div.setAttribute("aria-atomic", "true");
    document.body.appendChild(div);
  }

  function showToast(message) {
    ensureToastRoot();
    const root = document.getElementById("toast-root") || document.body;
    const toast = document.createElement("div");
    toast.className = "toast";
    toast.innerHTML = `<div class="t-title">Notice</div><div class="t-body">${esc(message)}</div>`;
    root.appendChild(toast);
    void toast.offsetWidth;
    setTimeout(() => {
      toast.style.opacity = "0";
      toast.style.transform = "translateY(4px)";
      setTimeout(() => toast.remove(), 200);
    }, 3500);
  }

  // =========================
  // 2) MODEL (kept, extended)
  // =========================
  const model = {
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
      lastUpdated: new Date().toISOString().slice(0, 10),
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
      mirrFinance: 6,
      mirrReinvest: 4,
      discountSchedule: JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE))
    },
    outputsMeta: {
      systemType: "single",
      assumptions: ""
    },
    outputs: [
      { id: uid(), name: "Grain yield", unit: "t/ha", value: 450, source: "Input directly" }
    ],
    treatments: [
      {
        id: uid(),
        name: "Control (baseline)",
        area: 100,
        adoption: 1,
        deltas: {},
        labourCost: 0,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "Farm trials",
        isControl: true,
        notes: "Control definition is taken from the uploaded dataset where available.",
        recurrenceYears: 0
      }
    ],
    benefits: [],
    otherCosts: [],
    adoption: { base: 1.0, low: 0.6, high: 1.0 },
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

  function initTreatmentDeltas() {
    model.treatments.forEach(t => {
      model.outputs.forEach(o => {
        if (!t.deltas) t.deltas = {};
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      });
      if (typeof t.labourCost === "undefined") t.labourCost = 0;
      if (typeof t.materialsCost === "undefined") t.materialsCost = 0;
      if (typeof t.servicesCost === "undefined") t.servicesCost = 0;
      if (typeof t.capitalCost === "undefined") t.capitalCost = 0;
      if (typeof t.adoption !== "number" || isNaN(t.adoption)) t.adoption = 1;
      if (typeof t.recurrenceYears !== "number" || isNaN(t.recurrenceYears)) t.recurrenceYears = 0;
    });
  }
  initTreatmentDeltas();

  // =========================
  // 3) STATE FOR DATASET + SCENARIOS
  // =========================
  const state = {
    dataset: {
      sourceName: "",
      rawText: "",
      rows: [],
      dictionary: null,
      schema: null,
      derived: {
        cleanedRows: [],
        checks: [],
        replicateBaselines: new Map(),
        plotDeltas: [],
        treatmentSummary: [],
        controlKey: null
      },
      committedAt: null
    },
    config: {
      persistenceYears: 5,
      sensPrice: DEFAULT_SENS_PRICE.slice(),
      sensDiscount: DEFAULT_SENS_DISC.slice(),
      sensPersistence: DEFAULT_SENS_PERSIST.slice(),
      sensRecurrence: DEFAULT_SENS_RECURRENCE.slice()
    },
    results: {
      perTreatmentBaseCase: [],
      sensitivityGrid: [],
      lastComputedAt: null,
      currentFilter: "all"
    }
  };

  // =========================
  // 4) CSV/TSV + DICTIONARY PARSING
  // =========================
  function parseDelimited(text, delimiter) {
    const rows = [];
    let i = 0;
    const len = text.length;
    let field = "";
    let row = [];
    let inQuotes = false;
    while (i < len) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          const next = text[i + 1];
          if (next === '"') {
            field += '"';
            i += 2;
            continue;
          } else {
            inQuotes = false;
            i++;
            continue;
          }
        } else {
          field += ch;
          i++;
          continue;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
          i++;
          continue;
        }
        if (ch === delimiter) {
          row.push(field);
          field = "";
          i++;
          continue;
        }
        if (ch === "\r") {
          i++;
          continue;
        }
        if (ch === "\n") {
          row.push(field);
          rows.push(row);
          row = [];
          field = "";
          i++;
          continue;
        }
        field += ch;
        i++;
      }
    }
    row.push(field);
    rows.push(row);
    while (rows.length && rows[rows.length - 1].every(c => String(c ?? "").trim() === "")) rows.pop();
    return rows;
  }

  function detectDelimiter(text) {
    const firstLine = (text || "").split(/\n/).find(l => l.trim().length > 0) || "";
    const tabCount = (firstLine.match(/\t/g) || []).length;
    const commaCount = (firstLine.match(/,/g) || []).length;
    if (tabCount >= commaCount && tabCount > 0) return "\t";
    if (commaCount > 0) return ",";
    const tabs = (text.match(/\t/g) || []).length;
    const commas = (text.match(/,/g) || []).length;
    if (tabs >= commas && tabs > 0) return "\t";
    return ",";
  }

  function normaliseHeader(h) {
    return String(h ?? "")
      .trim()
      .replace(/\s+/g, " ")
      .replace(/[^\S\r\n]+/g, " ");
  }

  function headersToObjects(table) {
    if (!table.length) return [];
    const header = table[0].map(normaliseHeader);
    const out = [];
    for (let r = 1; r < table.length; r++) {
      const row = table[r];
      const obj = {};
      for (let c = 0; c < header.length; c++) {
        const key = header[c] || `col_${c + 1}`;
        obj[key] = row[c] ?? "";
      }
      out.push(obj);
    }
    return out;
  }

  function looksLikeDictionaryHeader(cols) {
    const lower = cols.map(c => String(c || "").toLowerCase());
    const hasVar = lower.some(c => c.includes("variable") || c.includes("field") || c === "name" || c.includes("column"));
    const hasDesc = lower.some(c => c.includes("description") || c.includes("label") || c.includes("notes") || c.includes("definition"));
    return hasVar && hasDesc;
  }

  function splitDictionaryAndDataFromText(rawText) {
    const text = String(rawText || "");
    const chunks = text.split(/\n{2,}/g).map(s => s.trim()).filter(Boolean);
    if (chunks.length < 2) return { dictText: null, dataText: text };

    let dictIdx = -1;
    let dataIdx = -1;

    for (let i = 0; i < chunks.length; i++) {
      const del = detectDelimiter(chunks[i]);
      const tbl = parseDelimited(chunks[i], del);
      if (!tbl.length) continue;
      const head = tbl[0] || [];
      if (looksLikeDictionaryHeader(head) && tbl.length >= 2) {
        dictIdx = i;
        break;
      }
    }

    if (dictIdx >= 0) {
      for (let j = dictIdx + 1; j < chunks.length; j++) {
        const del = detectDelimiter(chunks[j]);
        const tbl = parseDelimited(chunks[j], del);
        if (!tbl.length) continue;
        const head = tbl[0].map(h => String(h || "").toLowerCase());
        const wide = (tbl[0] || []).length >= 6;
        const hasYield = head.some(h => h.includes("yield"));
        const hasTreat = head.some(h => h.includes("treatment") || h.includes("amend") || h.includes("variant"));
        const hasRep = head.some(h => h.includes("rep") || h.includes("block") || h.includes("replicate"));
        if (wide || (hasYield && hasTreat) || (hasYield && hasRep)) {
          dataIdx = j;
          break;
        }
      }
    }

    if (dictIdx >= 0 && dataIdx >= 0) {
      return { dictText: chunks[dictIdx], dataText: chunks[dataIdx] };
    }
    return { dictText: null, dataText: text };
  }

  function parseDictionaryText(dictText) {
    if (!dictText) return null;
    const del = detectDelimiter(dictText);
    const tbl = parseDelimited(dictText, del);
    if (!tbl.length) return null;
    const objs = headersToObjects(tbl);

    const keys = Object.keys(objs[0] || {});
    const lowerKeys = keys.map(k => k.toLowerCase());

    const varKey =
      keys[lowerKeys.findIndex(k => k.includes("variable") || k.includes("field") || k === "name" || k.includes("column"))] ||
      keys[0];
    const descKey =
      keys[lowerKeys.findIndex(k => k.includes("description") || k.includes("label") || k.includes("definition") || k.includes("notes"))] ||
      keys[Math.min(1, keys.length - 1)] ||
      keys[0];

    const roleKey =
      keys[lowerKeys.findIndex(k => k.includes("role") || k.includes("type") || k.includes("category") || k.includes("domain"))] ||
      null;

    const unitKey = keys[lowerKeys.findIndex(k => k.includes("unit"))] || null;

    const dict = new Map();
    objs.forEach(r => {
      const v = String(r[varKey] ?? "").trim();
      if (!v) return;
      dict.set(v, {
        variable: v,
        description: String(r[descKey] ?? "").trim(),
        role: roleKey ? String(r[roleKey] ?? "").trim() : "",
        unit: unitKey ? String(r[unitKey] ?? "").trim() : ""
      });
    });

    return { rows: objs, map: dict };
  }

  function inferSchema(rows, dictionary) {
    const headers = rows.length ? Object.keys(rows[0]) : [];

    function bestHeader(cands) {
      const lower = headers.map(h => h.toLowerCase());
      for (const c of cands) {
        const idx = lower.findIndex(h => h === c || h.includes(c));
        if (idx >= 0) return headers[idx];
      }
      return null;
    }

    const dictRoleMatch = role => {
      if (!dictionary || !dictionary.map) return null;
      for (const [k, meta] of dictionary.map.entries()) {
        const r = String(meta.role || "").toLowerCase();
        if (r.includes(role)) {
          const idx = headers.findIndex(h => h.trim() === k.trim());
          if (idx >= 0) return headers[idx];
        }
      }
      return null;
    };

    const treatmentCol =
      dictRoleMatch("treatment") ||
      bestHeader(["amendment", "treatment", "variant", "package", "option", "arm"]);
    const replicateCol =
      dictRoleMatch("replicate") ||
      bestHeader(["replicate", "rep", "block", "trial block", "replication"]);
    const plotCol = dictRoleMatch("plot") || bestHeader(["plot", "plot id", "plotid", "plot_no", "plot number"]);
    const controlFlagCol = dictRoleMatch("control") || bestHeader(["is_control", "control", "baseline"]);
    const yieldCol =
      dictRoleMatch("yield") ||
      bestHeader(["yield t/ha", "yield", "grain yield", "yield_tha", "yield (t/ha)", "yield t/ha"]);

    const costCols = headers.filter(h => {
      const s = h.toLowerCase();
      const isCosty =
        s.includes("cost") || s.includes("labour") || s.includes("labor") || s.includes("input") || s.includes("fert") ||
        s.includes("herb") || s.includes("fung") || s.includes("insect") || s.includes("fuel") || s.includes("machinery") ||
        s.includes("spray") || s.includes("seed");
      const isClearlyNotCost =
        s.includes("yield") || s.includes("protein") || s.includes("screen") || s.includes("moist") || s.includes("rep") ||
        s.includes("plot") || s.includes("treatment") || s.includes("amend");
      return isCosty && !isClearlyNotCost;
    });

    const plotAreaCol = bestHeader(["plot area", "plot_area", "area (ha)", "area_ha", "plot_ha", "ha"]);

    return {
      headers,
      treatmentCol,
      replicateCol,
      plotCol,
      controlFlagCol,
      yieldCol,
      costCols,
      plotAreaCol
    };
  }

  function normaliseTreatmentKey(v) {
    return String(v ?? "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ")
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  function detectControlKey(rows, schema) {
    if (!rows.length) return null;
    const tCol = schema.treatmentCol;
    const cCol = schema.controlFlagCol;

    if (cCol) {
      for (const r of rows) {
        const v = r[cCol];
        const s = String(v ?? "").trim().toLowerCase();
        const truthy = s === "1" || s === "true" || s === "yes" || s === "y";
        if (truthy && tCol && !isBlank(r[tCol])) return normaliseTreatmentKey(r[tCol]);
      }
    }

    if (tCol) {
      const counts = new Map();
      for (const r of rows) {
        const tv = String(r[tCol] ?? "");
        if (!tv.trim()) continue;
        const key = normaliseTreatmentKey(tv);
        const low = tv.toLowerCase();
        const isCtrl = low.includes("control") || low.includes("baseline") || low.includes("check");
        if (isCtrl) counts.set(key, (counts.get(key) || 0) + 1);
      }
      if (counts.size) {
        let best = null;
        let bestN = -1;
        for (const [k, n] of counts.entries()) {
          if (n > bestN) {
            best = k;
            bestN = n;
          }
        }
        return best;
      }
    }
    return null;
  }

  function costPerHaFromRow(row, schema, col) {
    const raw = parseNumber(row[col]);
    if (!Number.isFinite(raw)) return NaN;

    const h = String(col || "").toLowerCase();
    const looksPerHa = h.includes("/ha") || h.includes("per ha") || h.includes("per_ha") || h.includes("ha)");
    if (looksPerHa) return raw;

    if (schema.plotAreaCol) {
      const a = parseNumber(row[schema.plotAreaCol]);
      if (Number.isFinite(a) && a > 0) return raw / a;
    }
    return raw;
  }

  function computeDerivedFromDataset(rows, schema) {
    const checks = [];
    const derived = {
      cleanedRows: [],
      checks,
      replicateBaselines: new Map(),
      plotDeltas: [],
      treatmentSummary: [],
      controlKey: null
    };

    if (!rows.length) {
      checks.push({ code: "NO_ROWS", severity: "error", message: "No data rows found after parsing.", count: 0, detail: "" });
      return derived;
    }

    if (!schema.treatmentCol)
      checks.push({ code: "NO_TREATMENT_COL", severity: "error", message: "Treatment column not found.", count: 0, detail: "" });
    if (!schema.replicateCol)
      checks.push({
        code: "NO_REPLICATE_COL",
        severity: "warn",
        message: "Replicate column not found. Replicate-specific baselines will fall back to overall control mean.",
        count: 0,
        detail: ""
      });
    if (!schema.yieldCol)
      checks.push({ code: "NO_YIELD_COL", severity: "error", message: "Yield column not found.", count: 0, detail: "" });

    const controlKey = detectControlKey(rows, schema);
    derived.controlKey = controlKey;

    if (!controlKey)
      checks.push({
        code: "NO_CONTROL_DETECTED",
        severity: "error",
        message: "Control treatment could not be detected. Provide an is_control column or ensure the control label includes the word control.",
        count: 0,
        detail: ""
      });

    const cleaned = rows.map((r, idx) => {
      const treatVal = schema.treatmentCol ? r[schema.treatmentCol] : "";
      const repVal = schema.replicateCol ? r[schema.replicateCol] : "";
      const plotVal = schema.plotCol ? r[schema.plotCol] : "";
      const y = schema.yieldCol ? parseNumber(r[schema.yieldCol]) : NaN;

      const tKey = normaliseTreatmentKey(treatVal);
      const repKey = schema.replicateCol ? String(repVal ?? "").trim() : "";
      const pKey = schema.plotCol ? String(plotVal ?? "").trim() : String(idx + 1);

      let isControl = false;
      if (schema.controlFlagCol) {
        const v = String(r[schema.controlFlagCol] ?? "").trim().toLowerCase();
        isControl = v === "1" || v === "true" || v === "yes" || v === "y";
      }
      if (!isControl && controlKey && tKey === controlKey) isControl = true;

      const costByCol = {};
      (schema.costCols || []).forEach(c => {
        costByCol[c] = costPerHaFromRow(r, schema, c);
      });

      return {
        __rowIndex: idx,
        treatment: String(treatVal ?? "").trim(),
        treatmentKey: tKey,
        replicate: repKey,
        plot: pKey,
        isControl,
        yield: y,
        costsPerHa: costByCol,
        original: r
      };
    });

    derived.cleanedRows = cleaned;

    const missingYield = cleaned.filter(r => !Number.isFinite(r.yield)).length;
    if (missingYield) {
      checks.push({
        code: "MISSING_YIELD",
        severity: "warn",
        message: "Some rows have missing yield values. These are excluded from yield summaries.",
        count: missingYield,
        detail: ""
      });
    }

    const negYield = cleaned.filter(r => Number.isFinite(r.yield) && r.yield < 0).length;
    if (negYield) {
      checks.push({
        code: "NEGATIVE_YIELD",
        severity: "warn",
        message: "Some rows have negative yield values. Check units or data entry.",
        count: negYield,
        detail: ""
      });
    }

    const reps = new Map();
    const overallCtrlY = [];
    const overallCtrlCosts = new Map();
    (schema.costCols || []).forEach(c => overallCtrlCosts.set(c, []));

    cleaned.forEach(r => {
      if (!r.isControl) return;
      if (Number.isFinite(r.yield)) overallCtrlY.push(r.yield);
      (schema.costCols || []).forEach(c => {
        const v = r.costsPerHa[c];
        if (Number.isFinite(v)) overallCtrlCosts.get(c).push(v);
      });

      const repKey = schema.replicateCol ? (r.replicate || "__NO_REP__") : "__NO_REP__";
      if (!reps.has(repKey)) {
        const m = new Map();
        (schema.costCols || []).forEach(c => m.set(c, []));
        reps.set(repKey, { ctrlY: [], ctrlCostsByCol: m });
      }
      const entry = reps.get(repKey);
      if (Number.isFinite(r.yield)) entry.ctrlY.push(r.yield);
      (schema.costCols || []).forEach(c => {
        const v = r.costsPerHa[c];
        if (Number.isFinite(v)) entry.ctrlCostsByCol.get(c).push(v);
      });
    });

    const overallCtrlMeanYield = mean(overallCtrlY);
    if (!Number.isFinite(overallCtrlMeanYield)) {
      checks.push({
        code: "CONTROL_YIELD_MISSING",
        severity: "error",
        message: "Control yields are missing. Cannot compute deltas.",
        count: 0,
        detail: ""
      });
    }

    const replicateBaselines = new Map();
    for (const [repKey, entry] of reps.entries()) {
      const yMean = mean(entry.ctrlY);
      const costsMean = {};
      (schema.costCols || []).forEach(c => {
        costsMean[c] = mean(entry.ctrlCostsByCol.get(c) || []);
      });
      replicateBaselines.set(repKey, {
        yieldMean: Number.isFinite(yMean) ? yMean : overallCtrlMeanYield,
        costsMeanByCol: costsMean
      });
    }
    derived.replicateBaselines = replicateBaselines;

    if (schema.replicateCol) {
      const allRepKeys = new Set(cleaned.map(r => r.replicate || "__MISSING_REP__"));
      let repsNoCtrl = 0;
      allRepKeys.forEach(k => {
        const has = replicateBaselines.has(k);
        if (!has) repsNoCtrl++;
      });
      if (repsNoCtrl) {
        checks.push({
          code: "REPS_WITHOUT_CONTROL",
          severity: "warn",
          message: "Some replicates have no control rows. Their baselines fall back to overall control mean.",
          count: repsNoCtrl,
          detail: ""
        });
      }
    }

    const plotDeltas = cleaned.map(r => {
      const repKey = schema.replicateCol ? (r.replicate || "__NO_REP__") : "__NO_REP__";
      const base = replicateBaselines.get(repKey) || { yieldMean: overallCtrlMeanYield, costsMeanByCol: {} };
      const dy = Number.isFinite(r.yield) && Number.isFinite(base.yieldMean) ? r.yield - base.yieldMean : NaN;
      const dCosts = {};
      (schema.costCols || []).forEach(c => {
        const v = r.costsPerHa[c];
        const b = base.costsMeanByCol ? base.costsMeanByCol[c] : NaN;
        dCosts[c] = Number.isFinite(v) && Number.isFinite(b) ? v - b : NaN;
      });
      return {
        ...r,
        controlYieldMeanRep: base.yieldMean,
        deltaYield: dy,
        deltaCostsPerHa: dCosts
      };
    });
    derived.plotDeltas = plotDeltas;

    const byTreat = new Map();
    plotDeltas.forEach(r => {
      if (!r.treatmentKey) return;
      if (!byTreat.has(r.treatmentKey)) {
        byTreat.set(r.treatmentKey, {
          treatmentKey: r.treatmentKey,
          treatmentLabel: r.treatment || r.treatmentKey,
          isControl: !!r.isControl,
          n: 0,
          yield: [],
          deltaYield: [],
          costsByCol: {},
          deltaCostsByCol: {}
        });
        (schema.costCols || []).forEach(c => {
          byTreat.get(r.treatmentKey).costsByCol[c] = [];
          byTreat.get(r.treatmentKey).deltaCostsByCol[c] = [];
        });
      }
      const g = byTreat.get(r.treatmentKey);
      g.n++;
      if (Number.isFinite(r.yield)) g.yield.push(r.yield);
      if (Number.isFinite(r.deltaYield)) g.deltaYield.push(r.deltaYield);

      (schema.costCols || []).forEach(c => {
        const v = r.costsPerHa[c];
        const dv = r.deltaCostsPerHa[c];
        if (Number.isFinite(v)) g.costsByCol[c].push(v);
        if (Number.isFinite(dv)) g.deltaCostsByCol[c].push(dv);
      });
    });

    const summaries = [];
    for (const [, g] of byTreat.entries()) {
      const dy = g.deltaYield;
      const y = g.yield;
      const out = {
        treatmentKey: g.treatmentKey,
        treatmentLabel: g.treatmentLabel,
        isControl: g.isControl,
        nRows: g.n,
        nYield: y.filter(Number.isFinite).length,
        yieldMean: mean(y),
        yieldSD: sd(y),
        deltaYieldMean: mean(dy),
        deltaYieldSD: sd(dy),
        deltaYieldMedian: median(dy),
        costs: {},
        deltaCosts: {}
      };
      (schema.costCols || []).forEach(c => {
        out.costs[c] = { mean: mean(g.costsByCol[c] || []), sd: sd(g.costsByCol[c] || []) };
        out.deltaCosts[c] = { mean: mean(g.deltaCostsByCol[c] || []), sd: sd(g.deltaCostsByCol[c] || []) };
      });
      summaries.push(out);
    }

    const dyAll = plotDeltas.map(r => r.deltaYield).filter(v => Number.isFinite(v));
    const outFlags = iqrOutlierFlags(dyAll);
    if (Number.isFinite(outFlags.outliers) && outFlags.outliers > 0) {
      checks.push({
        code: "YIELD_DELTA_OUTLIERS",
        severity: "warn",
        message: "Some yield deltas are outliers by IQR rule. Check plots or consider robustness.",
        count: outFlags.outliers,
        detail:
          Number.isFinite(outFlags.low) && Number.isFinite(outFlags.high)
            ? `IQR bounds are ${fmt(outFlags.low)} to ${fmt(outFlags.high)} t/ha.`
            : ""
      });
    }

    const lowN = summaries.filter(s => !s.isControl && (s.nYield || 0) < 2).length;
    if (lowN) {
      checks.push({
        code: "LOW_REPLICATION",
        severity: "warn",
        message: "Some treatments have fewer than 2 yield observations. Means are unstable.",
        count: lowN,
        detail: ""
      });
    }

    let missingCostCells = 0;
    (schema.costCols || []).forEach(c => {
      plotDeltas.forEach(r => {
        if (!Number.isFinite(r.costsPerHa[c]) && !isBlank(r.original[c])) missingCostCells += 1;
      });
    });
    if (missingCostCells) {
      checks.push({
        code: "NON_NUMERIC_COSTS",
        severity: "warn",
        message: "Some cost cells are non-numeric. They are treated as missing and excluded from cost summaries.",
        count: missingCostCells,
        detail: ""
      });
    }

    derived.treatmentSummary = summaries;
    return derived;
  }

  // =========================
  // 5) CALIBRATE MODEL FROM DATASET SUMMARY
  // =========================
  function ensureYieldOutput() {
    let out = model.outputs.find(o => String(o.name || "").toLowerCase().includes("yield"));
    if (!out) {
      out = { id: uid(), name: "Grain yield", unit: "t/ha", value: 450, source: "Input directly" };
      model.outputs.unshift(out);
    }
    return out;
  }

  function extractIncrementalCostsFromSummary(summary, schema) {
    let labour = 0;
    let services = 0;
    let materials = 0;

    const cols = schema.costCols || [];
    cols.forEach(c => {
      const dv = summary.deltaCosts && summary.deltaCosts[c] ? summary.deltaCosts[c].mean : NaN;
      if (!Number.isFinite(dv)) return;
      const h = String(c || "").toLowerCase();
      const isLab = h.includes("labour") || h.includes("labor") || h.includes("hours") || h.includes("wage");
      const isServ = h.includes("contract") || h.includes("service") || h.includes("hire") || h.includes("machinery") || h.includes("fuel") || h.includes("spray");
      if (isLab) labour += dv;
      else if (isServ) services += dv;
      else materials += dv;
    });

    return { labour, services, materials };
  }

  function applyDatasetToModel() {
    const derived = state.dataset.derived;
    const schema = state.dataset.schema;

    if (!derived || !derived.treatmentSummary || !derived.treatmentSummary.length) {
      showToast("No derived treatment summary available to apply.");
      return;
    }

    const yieldOut = ensureYieldOutput();
    const yieldId = yieldOut.id;

    const controlSummary =
      derived.treatmentSummary.find(s => s.isControl) ||
      (derived.controlKey ? derived.treatmentSummary.find(s => s.treatmentKey === derived.controlKey) : null);

    const controlName = controlSummary ? controlSummary.treatmentLabel : "Control (baseline)";
    const currentControl = model.treatments.find(t => t.isControl) || model.treatments[0];
    const farmArea = currentControl ? Number(currentControl.area) || 100 : 100;

    const newTreatments = [];
    newTreatments.push({
      id: uid(),
      name: controlName || "Control (baseline)",
      area: farmArea,
      adoption: 1,
      deltas: { [yieldId]: 0 },
      labourCost: 0,
      materialsCost: 0,
      servicesCost: 0,
      capitalCost: 0,
      constrained: true,
      source: "Imported dataset",
      isControl: true,
      notes: "Control is defined by the dataset control flag or by the treatment label.",
      recurrenceYears: 0
    });

    derived.treatmentSummary
      .filter(s => !s.isControl)
      .sort((a, b) => {
        const A = Number.isFinite(a.deltaYieldMean) ? a.deltaYieldMean : -Infinity;
        const B = Number.isFinite(b.deltaYieldMean) ? b.deltaYieldMean : -Infinity;
        return B - A;
      })
      .forEach(s => {
        const incCosts = extractIncrementalCostsFromSummary(s, schema);
        const t = {
          id: uid(),
          name: s.treatmentLabel || s.treatmentKey,
          area: farmArea,
          adoption: 1,
          deltas: { [yieldId]: Number.isFinite(s.deltaYieldMean) ? s.deltaYieldMean : 0 },
          labourCost: Number.isFinite(incCosts.labour) ? incCosts.labour : 0,
          materialsCost: Number.isFinite(incCosts.materials) ? incCosts.materials : 0,
          servicesCost: Number.isFinite(incCosts.services) ? incCosts.services : 0,
          capitalCost: 0,
          constrained: true,
          source: "Imported dataset",
          isControl: false,
          notes: "Incremental values are computed as replicate-specific deltas relative to the control mean within each replicate.",
          recurrenceYears: 0
        };
        newTreatments.push(t);
      });

    model.treatments = newTreatments;
    initTreatmentDeltas();
    showToast("Dataset applied. Treatments calibrated from replicate-specific deltas versus control.");
  }

  // =========================
  // 6) CBA ENGINE
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

  function presentValue(series, ratePct) {
    let pv = 0;
    for (let t = 0; t < series.length; t++) {
      pv += series[t] / Math.pow(1 + ratePct / 100, t);
    }
    return pv;
  }

  function getGrainPrice() {
    const el = document.getElementById("grainPrice");
    const v = el ? parseNumber(el.value) : NaN;
    if (Number.isFinite(v)) return v;
    const yieldOut = ensureYieldOutput();
    const p = Number(yieldOut.value) || 0;
    return Number.isFinite(p) ? p : 0;
  }

  function getPersistenceYears() {
    const el = document.getElementById("persistenceYears");
    const v = el ? parseNumber(el.value) : NaN;
    if (Number.isFinite(v) && v >= 0) return Math.floor(v);
    return Math.max(0, Math.floor(state.config.persistenceYears || 0));
  }

  function getRecurrenceYears(t) {
    const v = Number(t.recurrenceYears);
    return Number.isFinite(v) ? Math.max(0, Math.floor(v)) : 0;
  }

  function buildTreatmentCashflowsVsControl(t, opts) {
    const years = Math.max(0, Math.floor(opts.years));
    const disc = Number(opts.discountRatePct) || 0;
    const price = Number(opts.pricePerTonne) || 0;
    const persistence = Math.max(0, Math.floor(opts.persistenceYears));
    const recurrence = Math.max(0, Math.floor(opts.recurrenceYears));
    const adoption = clamp(Number(opts.adoptionMultiplier) || 1, 0, 1);
    const risk = clamp(Number(opts.riskMultiplier) || 0, 0, 1);

    const area = Number(t.area) || 0;
    const yieldOut = ensureYieldOutput();
    const yieldDelta = Number(t.deltas && t.deltas[yieldOut.id]) || 0;

    const benefitByYear = new Array(years + 1).fill(0);
    for (let y = 1; y <= years; y++) {
      if (persistence === 0 || y > persistence) {
        benefitByYear[y] = 0;
      } else {
        const annual = yieldDelta * price * area * adoption * (1 - risk);
        benefitByYear[y] = annual;
      }
    }

    const costByYear = new Array(years + 1).fill(0);
    const perHaApplicationCost =
      (Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0) + (Number(t.labourCost) || 0);

    const cap0 = Number(t.capitalCost) || 0;
    costByYear[0] += cap0;

    if (!t.isControl) {
      costByYear[0] += perHaApplicationCost * area;
      if (recurrence > 0) {
        for (let y = recurrence; y <= years; y += recurrence) {
          costByYear[y] += perHaApplicationCost * area;
        }
      }
    }

    const cf = new Array(years + 1).fill(0).map((_, i) => benefitByYear[i] - costByYear[i]);

    const pvBenefits = presentValue(benefitByYear, disc);
    const pvCosts = presentValue(costByYear, disc);
    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts > 0 ? pvBenefits / pvCosts : NaN;
    const roi = pvCosts > 0 ? (npv / pvCosts) * 100 : NaN;

    const irrVal = irr(cf);
    const mirrVal = mirr(cf, model.time.mirrFinance, model.time.mirrReinvest);
    const pb = payback(cf, disc);

    return {
      treatmentId: t.id,
      treatmentName: t.name,
      isControl: !!t.isControl,
      areaHa: area,
      pricePerTonne: price,
      discountRatePct: disc,
      persistenceYears: persistence,
      recurrenceYears: recurrence,
      adoptionMultiplier: adoption,
      riskMultiplier: risk,
      perHaApplicationCost,
      pvBenefits,
      pvCosts,
      npv,
      bcr,
      roiPct: roi,
      irrPct: irrVal,
      mirrPct: mirrVal,
      paybackYears: pb,
      benefitByYear,
      costByYear,
      cf
    };
  }

  function computeBaseCaseResultsVsControl() {
    const price = getGrainPrice();
    const disc = Number(model.time.discBase) || 0;
    const years = Math.max(0, Math.floor(model.time.years || 0));
    const persistence = getPersistenceYears();
    const adopt = clamp(Number(model.adoption.base) || 1, 0, 1);
    const risk = clamp(Number(model.risk.base) || 0, 0, 1);

    const control = model.treatments.find(x => x.isControl) || null;

    const results = model.treatments.map(t =>
      buildTreatmentCashflowsVsControl(t, {
        pricePerTonne: price,
        discountRatePct: disc,
        years,
        persistenceYears: persistence,
        recurrenceYears: getRecurrenceYears(t),
        adoptionMultiplier: adopt,
        riskMultiplier: risk
      })
    );

    const ranked = results
      .filter(r => !r.isControl)
      .slice()
      .sort((a, b) => {
        const A = Number.isFinite(a.npv) ? a.npv : -Infinity;
        const B = Number.isFinite(b.npv) ? b.npv : -Infinity;
        return B - A;
      })
      .map((r, i) => ({ ...r, rankByNpv: i + 1 }));

    const out = results.map(r => {
      if (r.isControl) return { ...r, rankByNpv: null };
      const rr = ranked.find(x => x.treatmentId === r.treatmentId);
      return rr ? rr : { ...r, rankByNpv: null };
    });

    state.results.perTreatmentBaseCase = out;
    state.results.lastComputedAt = new Date().toISOString();

    const totalBenefit = new Array(years + 1).fill(0);
    const totalCost = new Array(years + 1).fill(0);
    out.forEach(r => {
      if (r.isControl) return;
      for (let y = 0; y <= years; y++) {
        totalBenefit[y] += r.benefitByYear[y] || 0;
        totalCost[y] += r.costByYear[y] || 0;
      }
    });
    const pvB = presentValue(totalBenefit, disc);
    const pvC = presentValue(totalCost, disc);
    const total = {
      pvBenefits: pvB,
      pvCosts: pvC,
      npv: pvB - pvC,
      bcr: pvC > 0 ? pvB / pvC : NaN
    };

    return { control, perTreatment: out, projectTotal: total };
  }

  function computeSensitivityGrid() {
    const priceGrid = (state.config.sensPrice || DEFAULT_SENS_PRICE).slice();
    const discGrid = (state.config.sensDiscount || DEFAULT_SENS_DISC).slice();
    const persistGrid = (state.config.sensPersistence || DEFAULT_SENS_PERSIST).slice();
    const recurGrid = (state.config.sensRecurrence || DEFAULT_SENS_RECURRENCE).slice();

    const years = Math.max(0, Math.floor(model.time.years || 0));
    const adopt = clamp(Number(model.adoption.base) || 1, 0, 1);
    const risk = clamp(Number(model.risk.base) || 0, 0, 1);

    const treatments = model.treatments.filter(t => !t.isControl);

    const grid = [];
    treatments.forEach(t => {
      const baseRec = getRecurrenceYears(t);
      recurGrid.forEach(rec => {
        const recurrenceYears = Number.isFinite(rec) ? Math.max(0, Math.floor(rec)) : baseRec;
        persistGrid.forEach(persistenceYears => {
          discGrid.forEach(discountRatePct => {
            priceGrid.forEach(pricePerTonne => {
              const r = buildTreatmentCashflowsVsControl(t, {
                pricePerTonne,
                discountRatePct,
                years,
                persistenceYears,
                recurrenceYears,
                adoptionMultiplier: adopt,
                riskMultiplier: risk
              });
              grid.push({
                treatment: t.name,
                treatmentId: t.id,
                pricePerTonne,
                discountRatePct,
                persistenceYears,
                recurrenceYears,
                pvBenefits: r.pvBenefits,
                pvCosts: r.pvCosts,
                npv: r.npv,
                bcr: r.bcr,
                roiPct: r.roiPct
              });
            });
          });
        });
      });
    });

    state.results.sensitivityGrid = grid;
    showToast("Sensitivity grid computed.");
    return grid;
  }

  // =========================
  // 7) RESULTS RENDERING
  // =========================
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));

  function classifyDelta(val) {
    if (!Number.isFinite(val)) return "";
    if (val > 0) return "pos";
    if (val < 0) return "neg";
    return "zero";
  }

  function filterTreatments(results, mode) {
    const list = results.filter(r => !r.isControl);
    if (!mode || mode === "all") return list;

    if (mode === "top5_npv") {
      return list
        .slice()
        .sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity))
        .slice(0, 5);
    }
    if (mode === "top5_bcr") {
      return list
        .slice()
        .sort((a, b) => (Number.isFinite(b.bcr) ? b.bcr : -Infinity) - (Number.isFinite(a.bcr) ? a.bcr : -Infinity))
        .slice(0, 5);
    }
    if (mode === "improve_only") {
      return list.filter(r => Number.isFinite(r.npv) && r.npv > 0);
    }
    return list;
  }

  function renderLeaderboard(perTreatment, filterMode) {
    const root =
      document.getElementById("resultsLeaderboard") ||
      document.getElementById("leaderboard") ||
      document.getElementById("treatmentLeaderboard");
    if (!root) return;

    const list = filterTreatments(perTreatment, filterMode)
      .slice()
      .sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity));

    root.innerHTML = "";

    if (!list.length) {
      root.innerHTML = `<p class="small muted">No treatments available. Import a dataset or review configuration.</p>`;
      return;
    }

    const table = document.createElement("table");
    table.className = "summary-table leaderboard-table";
    table.innerHTML = `
      <thead>
        <tr>
          <th>Rank</th>
          <th>Treatment</th>
          <th>NPV</th>
          <th>BCR</th>
          <th>PV benefits</th>
          <th>PV costs</th>
        </tr>
      </thead>
      <tbody>
        ${list
          .map((r, i) => {
            const rank = i + 1;
            const npvCls = classifyDelta(r.npv);
            const bcrText = Number.isFinite(r.bcr) ? fmt(r.bcr) : "n/a";
            return `
              <tr>
                <td>${rank}</td>
                <td>${esc(r.treatmentName)}</td>
                <td class="${npvCls}">${money(r.npv)}</td>
                <td>${bcrText}</td>
                <td>${money(r.pvBenefits)}</td>
                <td>${money(r.pvCosts)}</td>
              </tr>
            `;
          })
          .join("")}
      </tbody>
    `;
    root.appendChild(table);
  }

  function formatIndicatorValue(key, r, isControl) {
    if (key === "pvBenefits") return money(isControl ? 0 : r.pvBenefits);
    if (key === "pvCosts") return money(isControl ? 0 : r.pvCosts);
    if (key === "npv") return money(isControl ? 0 : r.npv);
    if (key === "bcr") return isControl ? "n/a" : (Number.isFinite(r.bcr) ? fmt(r.bcr) : "n/a");
    if (key === "roiPct") return isControl ? "n/a" : (Number.isFinite(r.roiPct) ? percent(r.roiPct) : "n/a");
    if (key === "rankByNpv") return isControl ? "" : (r.rankByNpv != null ? String(r.rankByNpv) : "");
    if (key === "deltaNpv") return isControl ? "" : money(r.npv);
    if (key === "deltaPvCost") return isControl ? "" : money(r.pvCosts);
    return "";
  }

  function formatDeltaValue(key, r) {
    if (key === "pvBenefits") return money(r.pvBenefits);
    if (key === "pvCosts") return money(r.pvCosts);
    if (key === "npv") return money(r.npv);
    if (key === "bcr") return Number.isFinite(r.bcr) ? fmt(r.bcr) : "n/a";
    if (key === "roiPct") return Number.isFinite(r.roiPct) ? percent(r.roiPct) : "n/a";
    if (key === "rankByNpv") return r.rankByNpv != null ? String(r.rankByNpv) : "";
    if (key === "deltaNpv") return money(r.npv);
    if (key === "deltaPvCost") return money(r.pvCosts);
    return "";
  }

  function classifyIndicatorCell(key, r, mode) {
    if (key === "pvCosts" || key === "deltaPvCost") {
      const v = r.pvCosts;
      if (!Number.isFinite(v)) return mode === "delta" ? "delta-cell" : "value-cell";
      const cls = v > 0 ? "neg" : v < 0 ? "pos" : "zero";
      return `${mode}-cell ${cls}`;
    }
    if (key === "npv" || key === "deltaNpv") {
      const cls = classifyDelta(r.npv);
      return `${mode}-cell ${cls}`;
    }
    if (key === "pvBenefits") {
      const cls = classifyDelta(r.pvBenefits);
      return `${mode}-cell ${cls}`;
    }
    if (key === "bcr") {
      const cls = Number.isFinite(r.bcr) ? (r.bcr >= 1 ? "pos" : "neg") : "";
      return `${mode}-cell ${cls}`;
    }
    return `${mode}-cell`;
  }

  function renderComparisonToControl(perTreatment, filterMode) {
    const root =
      document.getElementById("comparisonToControl") ||
      document.getElementById("comparisonTable") ||
      document.getElementById("comparisonGrid") ||
      document.getElementById("resultsComparison");

    if (!root) return;

    const control = perTreatment.find(r => r.isControl) || null;
    const treatments = filterTreatments(perTreatment, filterMode)
      .slice()
      .sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity));

    root.innerHTML = "";

    if (!control || !treatments.length) {
      root.innerHTML = `<p class="small muted">Comparison table will appear here once there is a control and at least one treatment.</p>`;
      return;
    }

    const indicators = [
      { key: "pvBenefits", label: "PV benefits" },
      { key: "pvCosts", label: "PV costs" },
      { key: "npv", label: "NPV" },
      { key: "bcr", label: "BCR" },
      { key: "roiPct", label: "ROI" },
      { key: "rankByNpv", label: "Rank by NPV" },
      { key: "deltaNpv", label: "Δ NPV vs control" },
      { key: "deltaPvCost", label: "Δ PV cost vs control" }
    ];

    const colHeaders = [];
    colHeaders.push({ type: "control", name: control ? control.treatmentName : "Control (baseline)" });
    treatments.forEach(t => {
      colHeaders.push({ type: "treatment", name: t.treatmentName, id: t.treatmentId });
      colHeaders.push({ type: "delta", name: "Δ", id: t.treatmentId });
    });

    const table = document.createElement("table");
    table.className = "summary-table c2c-table";

    const thead = document.createElement("thead");
    thead.innerHTML = `
      <tr>
        <th class="sticky-col">Indicator</th>
        ${colHeaders
          .map(h => {
            if (h.type === "control") return `<th class="control-col">${esc(h.name)} (baseline)</th>`;
            if (h.type === "treatment") return `<th>${esc(h.name)}</th>`;
            return `<th class="delta-col">Δ vs control</th>`;
          })
          .join("")}
      </tr>
    `;

    const tbody = document.createElement("tbody");

    indicators.forEach(ind => {
      const tr = document.createElement("tr");
      const first = document.createElement("td");
      first.className = "sticky-col";
      first.textContent = ind.label;
      tr.appendChild(first);

      const tdControl = document.createElement("td");
      tdControl.className = "control-col";
      tdControl.textContent = formatIndicatorValue(ind.key, control, true);
      tr.appendChild(tdControl);

      treatments.forEach(t => {
        const td = document.createElement("td");
        td.textContent = formatIndicatorValue(ind.key, t, false);
        td.className = classifyIndicatorCell(ind.key, t, "value");
        tr.appendChild(td);

        const tdD = document.createElement("td");
        tdD.textContent = formatDeltaValue(ind.key, t);
        tdD.className = classifyIndicatorCell(ind.key, t, "delta");
        tr.appendChild(tdD);
      });

      tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);

    const wrap = document.createElement("div");
    wrap.className = "c2c-wrap";
    wrap.appendChild(table);

    root.appendChild(wrap);
  }

  function renderResultsNarrative(perTreatment, filterMode) {
    const root =
      document.getElementById("resultsNarrative") ||
      document.getElementById("whatThisMeans") ||
      document.getElementById("resultsWhatThisMeans");

    if (!root) return;

    const treatments = filterTreatments(perTreatment, filterMode)
      .slice()
      .sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity));

    const top = treatments[0] || null;
    const worst = treatments.slice().reverse().find(x => Number.isFinite(x.npv)) || null;

    const price = getGrainPrice();
    const disc = Number(model.time.discBase) || 0;
    const years = Math.floor(model.time.years || 0);
    const persistence = getPersistenceYears();
    const adopt = clamp(Number(model.adoption.base) || 1, 0, 1);
    const risk = clamp(Number(model.risk.base) || 0, 0, 1);

    const parts = [];
    parts.push(
      `This results panel compares each treatment against the control baseline using discounted cashflows over ${years} years. The grain price used in the base case is ${money(
        price
      )} per tonne, the discount rate is ${fmt(disc)} percent per year, the assumed persistence of yield effects is ${persistence} years, the adoption multiplier is ${fmt(
        adopt
      )}, and the risk multiplier reduces benefits by ${fmt(risk)} as a proportion.`
    );

    if (top) {
      parts.push(
        `The strongest base case result by net present value is ${top.treatmentName}. Its present value of benefits is ${money(
          top.pvBenefits
        )}, its present value of costs is ${money(top.pvCosts)}, and its net present value is ${money(
          top.npv
        )}. This result is driven by the combination of yield uplift against the control and the incremental costs applied under the recurrence setting for that treatment.`
      );
    }

    if (worst && top && worst.treatmentId !== top.treatmentId) {
      parts.push(
        `A weaker base case result is ${worst.treatmentName}. Its present value of benefits is ${money(
          worst.pvBenefits
        )}, its present value of costs is ${money(worst.pvCosts)}, and its net present value is ${money(
          worst.npv
        )}. This pattern usually reflects either a smaller yield uplift compared with the control, higher incremental costs, or both.`
      );
    }

    const improves = treatments.filter(r => Number.isFinite(r.npv) && r.npv > 0).length;
    const total = treatments.length;
    if (total) {
      parts.push(
        `Under the current assumptions, ${improves} of ${total} treatments have a positive net present value relative to the control. This does not decide anything by itself, but it highlights which packages are more sensitive to costs and grain price assumptions.`
      );
    }

    root.textContent = parts.join("\n\n");
  }

  // =========================
  // 8) DATA CHECKS RENDERING
  // =========================
  function renderDataChecks() {
    const root =
      document.getElementById("dataChecks") ||
      document.getElementById("dataChecksList") ||
      document.getElementById("checksPanel");
    if (!root) return;

    const checks = state.dataset.derived && state.dataset.derived.checks ? state.dataset.derived.checks : [];
    if (!checks.length) {
      root.innerHTML = `<p class="small muted">No data checks triggered. If you have imported a dataset, this means required columns were found and core summaries could be computed.</p>`;
      return;
    }

    const rows = checks.slice().sort((a, b) => {
      const sevRank = s => (s === "error" ? 0 : s === "warn" ? 1 : 2);
      return sevRank(a.severity) - sevRank(b.severity);
    });

    root.innerHTML = `
      <table class="summary-table checks-table">
        <thead>
          <tr>
            <th>Severity</th>
            <th>Check</th>
            <th>Count</th>
            <th>Summary</th>
          </tr>
        </thead>
        <tbody>
          ${rows
            .map(r => {
              const sev = String(r.severity || "info").toUpperCase();
              const cls = r.severity === "error" ? "neg" : r.severity === "warn" ? "warn" : "zero";
              return `
                <tr>
                  <td class="${cls}">${esc(sev)}</td>
                  <td><code>${esc(r.code || "")}</code></td>
                  <td>${Number.isFinite(r.count) ? fmt(r.count) : ""}</td>
                  <td>${esc(r.message || "")}${r.detail ? " " + esc(r.detail) : ""}</td>
                </tr>
              `;
            })
            .join("")}
        </tbody>
      </table>
    `;
  }

  // =========================
  // 9) EXPORTS
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

  function toCsv(rows) {
    return rows
      .map(r =>
        r
          .map(x => {
            const s = x == null ? "" : String(x);
            const needs = s.includes(",") || s.includes('"') || s.includes("\n") || s.includes("\r");
            const safe = s.replace(/"/g, '""');
            return needs ? `"${safe}"` : safe;
          })
          .join(",")
      )
      .join("\r\n");
  }

  function exportCleanedDatasetTsv() {
    const derived = state.dataset.derived;
    if (!derived || !derived.plotDeltas || !derived.plotDeltas.length) {
      showToast("No cleaned dataset to export.");
      return;
    }
    const schema = state.dataset.schema;
    const rows = derived.plotDeltas;

    const costCols = schema && schema.costCols ? schema.costCols.slice() : [];
    const header = [
      "treatment",
      "treatment_key",
      "replicate",
      "plot",
      "is_control",
      "yield",
      "control_yield_mean_replicate",
      "delta_yield",
      ...costCols.map(c => `cost_per_ha:${c}`),
      ...costCols.map(c => `delta_cost_per_ha:${c}`)
    ];

    const lines = [header.join("\t")];
    rows.forEach(r => {
      const out = [
        r.treatment || "",
        r.treatmentKey || "",
        r.replicate || "",
        r.plot || "",
        r.isControl ? "1" : "0",
        Number.isFinite(r.yield) ? r.yield : "",
        Number.isFinite(r.controlYieldMeanRep) ? r.controlYieldMeanRep : "",
        Number.isFinite(r.deltaYield) ? r.deltaYield : ""
      ];

      costCols.forEach(c => out.push(Number.isFinite(r.costsPerHa[c]) ? r.costsPerHa[c] : ""));
      costCols.forEach(c => out.push(Number.isFinite(r.deltaCostsPerHa[c]) ? r.deltaCostsPerHa[c] : ""));

      lines.push(out.join("\t"));
    });

    const name = slug(model.project.name || "project");
    downloadFile(`${name}_cleaned_dataset.tsv`, lines.join("\n"), "text/tab-separated-values");
    showToast("Cleaned dataset TSV downloaded.");
  }

  function exportTreatmentSummaryCsv() {
    const derived = state.dataset.derived;
    if (!derived || !derived.treatmentSummary || !derived.treatmentSummary.length) {
      showToast("No treatment summary to export.");
      return;
    }

    const schema = state.dataset.schema;
    const costCols = schema && schema.costCols ? schema.costCols.slice() : [];

    const rows = [];
    rows.push([
      "treatment",
      "is_control",
      "n_yield",
      "yield_mean",
      "yield_sd",
      "delta_yield_mean",
      "delta_yield_sd",
      "delta_yield_median",
      ...costCols.map(c => `delta_cost_mean:${c}`),
      ...costCols.map(c => `delta_cost_sd:${c}`)
    ]);

    derived.treatmentSummary
      .slice()
      .sort((a, b) => (b.isControl ? 1 : 0) - (a.isControl ? 1 : 0))
      .forEach(s => {
        const row = [
          s.treatmentLabel || s.treatmentKey,
          s.isControl ? "1" : "0",
          Number.isFinite(s.nYield) ? s.nYield : "",
          Number.isFinite(s.yieldMean) ? s.yieldMean : "",
          Number.isFinite(s.yieldSD) ? s.yieldSD : "",
          Number.isFinite(s.deltaYieldMean) ? s.deltaYieldMean : "",
          Number.isFinite(s.deltaYieldSD) ? s.deltaYieldSD : "",
          Number.isFinite(s.deltaYieldMedian) ? s.deltaYieldMedian : ""
        ];
        costCols.forEach(c => row.push(Number.isFinite(s.deltaCosts?.[c]?.mean) ? s.deltaCosts[c].mean : ""));
        costCols.forEach(c => row.push(Number.isFinite(s.deltaCosts?.[c]?.sd) ? s.deltaCosts[c].sd : ""));
        rows.push(row);
      });

    const csv = toCsv(rows);
    const name = slug(model.project.name || "project");
    downloadFile(`${name}_treatment_summary.csv`, csv, "text/csv");
    showToast("Treatment summary CSV downloaded.");
  }

  function exportSensitivityGridCsv() {
    const grid = state.results.sensitivityGrid || [];
    if (!grid.length) {
      showToast("No sensitivity grid to export. Run sensitivity first.");
      return;
    }

    const rows = [];
    rows.push([
      "treatment",
      "price_per_tonne",
      "discount_rate_pct",
      "persistence_years",
      "recurrence_years",
      "pv_benefits",
      "pv_costs",
      "npv",
      "bcr",
      "roi_pct"
    ]);

    grid.forEach(g => {
      rows.push([
        g.treatment,
        g.pricePerTonne,
        g.discountRatePct,
        g.persistenceYears,
        g.recurrenceYears,
        g.pvBenefits,
        g.pvCosts,
        g.npv,
        g.bcr,
        g.roiPct
      ]);
    });

    const csv = toCsv(rows);
    const name = slug(model.project.name || "project");
    downloadFile(`${name}_sensitivity_grid.csv`, csv, "text/csv");
    showToast("Sensitivity grid CSV downloaded.");
  }

  function exportWorkbookIfAvailable() {
    if (typeof XLSX === "undefined") {
      showToast("Excel export requires the XLSX library.");
      return;
    }

    const wb = XLSX.utils.book_new();
    const name = slug(model.project.name || "project");

    const settingsAoA = [
      ["Project name", model.project.name],
      ["Analysis years", model.time.years],
      ["Discount rate base percent", model.time.discBase],
      ["Grain price per tonne", getGrainPrice()],
      ["Persistence years", getPersistenceYears()],
      ["Adoption multiplier", model.adoption.base],
      ["Risk multiplier", model.risk.base],
      ["Dataset source", state.dataset.sourceName || ""],
      ["Dataset committed at", state.dataset.committedAt || ""]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(settingsAoA), "Settings");

    const derived = state.dataset.derived;
    if (derived && derived.plotDeltas && derived.plotDeltas.length) {
      const schema = state.dataset.schema;
      const costCols = schema && schema.costCols ? schema.costCols.slice() : [];
      const rows = derived.plotDeltas.map(r => {
        const obj = {
          treatment: r.treatment,
          treatment_key: r.treatmentKey,
          replicate: r.replicate,
          plot: r.plot,
          is_control: r.isControl ? 1 : 0,
          yield: Number.isFinite(r.yield) ? r.yield : null,
          control_yield_mean_replicate: Number.isFinite(r.controlYieldMeanRep) ? r.controlYieldMeanRep : null,
          delta_yield: Number.isFinite(r.deltaYield) ? r.deltaYield : null
        };
        costCols.forEach(c => {
          obj[`cost_per_ha:${c}`] = Number.isFinite(r.costsPerHa[c]) ? r.costsPerHa[c] : null;
        });
        costCols.forEach(c => {
          obj[`delta_cost_per_ha:${c}`] = Number.isFinite(r.deltaCostsPerHa[c]) ? r.deltaCostsPerHa[c] : null;
        });
        return obj;
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "CleanedData");
    }

    if (derived && derived.treatmentSummary && derived.treatmentSummary.length) {
      const schema = state.dataset.schema;
      const costCols = schema && schema.costCols ? schema.costCols.slice() : [];
      const rows = derived.treatmentSummary.map(s => {
        const obj = {
          treatment: s.treatmentLabel || s.treatmentKey,
          is_control: s.isControl ? 1 : 0,
          n_yield: s.nYield,
          yield_mean: s.yieldMean,
          yield_sd: s.yieldSD,
          delta_yield_mean: s.deltaYieldMean,
          delta_yield_sd: s.deltaYieldSD,
          delta_yield_median: s.deltaYieldMedian
        };
        costCols.forEach(c => {
          obj[`delta_cost_mean:${c}`] = s.deltaCosts?.[c]?.mean ?? null;
          obj[`delta_cost_sd:${c}`] = s.deltaCosts?.[c]?.sd ?? null;
        });
        return obj;
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "TreatmentSummary");
    }

    const base = state.results.perTreatmentBaseCase || [];
    if (base.length) {
      const rows = base.map(r => ({
        treatment: r.treatmentName,
        is_control: r.isControl ? 1 : 0,
        area_ha: r.areaHa,
        recurrence_years: r.recurrenceYears,
        price_per_tonne: r.pricePerTonne,
        discount_rate_pct: r.discountRatePct,
        persistence_years: r.persistenceYears,
        pv_benefits: r.pvBenefits,
        pv_costs: r.pvCosts,
        npv: r.npv,
        bcr: r.bcr,
        roi_pct: r.roiPct,
        irr_pct: r.irrPct,
        mirr_pct: r.mirrPct,
        payback_years: r.paybackYears
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "BaseCaseResults");
    }

    const grid = state.results.sensitivityGrid || [];
    if (grid.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(grid), "SensitivityGrid");
    }

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile(`${name}_workbook.xlsx`, wbout, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Excel workbook downloaded.");
  }

  // =========================
  // 10) AI BRIEFING
  // =========================
  function buildResultsJsonPayload() {
    const derived = state.dataset.derived || null;
    const schema = state.dataset.schema || null;

    return {
      toolName: "Farming CBA Decision Tool",
      project: model.project,
      time: model.time,
      config: {
        grainPricePerTonne: getGrainPrice(),
        persistenceYears: getPersistenceYears(),
        adoptionMultiplier: model.adoption.base,
        riskMultiplier: model.risk.base
      },
      dataset: {
        sourceName: state.dataset.sourceName,
        committedAt: state.dataset.committedAt,
        schema,
        checks: derived ? derived.checks : [],
        treatmentSummary: derived ? derived.treatmentSummary : []
      },
      results: {
        baseCasePerTreatment: state.results.perTreatmentBaseCase,
        sensitivityGridCount: (state.results.sensitivityGrid || []).length
      }
    };
  }

  function buildAiBriefingText() {
    const base = state.results.perTreatmentBaseCase || [];
    const treatments = base
      .filter(r => !r.isControl)
      .slice()
      .sort((a, b) => (b.npv || -Infinity) - (a.npv || -Infinity));

    const years = Math.floor(model.time.years || 0);
    const price = getGrainPrice();
    const disc = Number(model.time.discBase) || 0;
    const persistence = getPersistenceYears();
    const adopt = clamp(Number(model.adoption.base) || 1, 0, 1);
    const risk = clamp(Number(model.risk.base) || 0, 0, 1);

    const derived = state.dataset.derived;
    const checks = derived && derived.checks ? derived.checks : [];

    const top = treatments[0] || null;

    const p = [];
    p.push(
      `Write a decision brief in clear plain language for an on farm manager. Use full sentences and paragraphs only. Do not use bullet points. Do not use an em dash. Do not use abbreviations.`
    );

    p.push(
      `The analysis compares each soil amendment treatment against the control baseline using discounted cashflows over ${years} years. The grain price used in the base case is ${money(
        price
      )} per tonne and the discount rate is ${fmt(disc)} percent per year. Yield effects are assumed to persist for ${persistence} years after application. The adoption multiplier is ${fmt(
        adopt
      )} and the risk multiplier reduces benefits by ${fmt(risk)} as a proportion.`
    );

    if (derived && derived.treatmentSummary && derived.treatmentSummary.length) {
      const nTreat = derived.treatmentSummary.filter(s => !s.isControl).length;
      const nRows = derived.plotDeltas ? derived.plotDeltas.length : 0;
      p.push(
        `The underlying dataset includes ${nRows} plot level records and ${nTreat} non control treatments. Treatment effects are computed using replicate specific control baselines, meaning each plot is compared with the control mean within the same replicate before averaging.`
      );
    }

    if (checks.length) {
      const errs = checks.filter(c => c.severity === "error").length;
      const warns = checks.filter(c => c.severity === "warn").length;
      p.push(
        `Data checks were run after import. There are ${errs} error level checks and ${warns} warning level checks. Summarise the most important implications for interpretation and sensitivity, without recommending a single option.`
      );
    } else {
      p.push(`Data checks were run after import and no triggers were recorded in the checks panel.`);
    }

    if (top) {
      p.push(
        `In the base case, the strongest treatment by net present value is ${top.treatmentName}. Its present value of benefits is ${money(
          top.pvBenefits
        )}, its present value of costs is ${money(top.pvCosts)}, and its net present value is ${money(
          top.npv
        )}. Explain in practical terms what is driving this result, focusing on yield uplift against the control and incremental costs under the recurrence assumption for that treatment.`
      );
    }

    if (treatments.length) {
      const pos = treatments.filter(t => Number.isFinite(t.npv) && t.npv > 0).length;
      p.push(
        `Across all treatments, ${pos} have a positive net present value relative to the control under the base case. Explain what this means in terms of trade offs. Do not instruct the user to choose anything.`
      );
    }

    p.push(
      `Include a section that explains the meaning of net present value, present value of benefits, present value of costs, benefit cost ratio, and return on investment in farmer facing terms.`
    );
    p.push(
      `Include a section that explains sensitivity, focusing on grain price, discount rate, persistence of yield effects, and recurrence of costs. Explain which inputs are likely to change and how that could move results.`
    );
    p.push(
      `Include a section that lists practical options to improve weaker treatments without giving a directive, such as reducing cost items, changing timing, improving establishment, or verifying the yield response with additional seasons or sites.`
    );

    const resultsBlock = {
      project: {
        name: model.project.name,
        years,
        discountRatePct: disc,
        grainPricePerTonne: price,
        persistenceYears: persistence,
        adoptionMultiplier: adopt,
        riskMultiplier: risk
      },
      treatments: treatments.slice(0, 12).map(t => ({
        name: t.treatmentName,
        recurrenceYears: t.recurrenceYears,
        pvBenefits: t.pvBenefits,
        pvCosts: t.pvCosts,
        npv: t.npv,
        bcr: t.bcr,
        roiPct: t.roiPct,
        irrPct: t.irrPct,
        paybackYears: t.paybackYears
      })),
      dataChecks: checks.slice(0, 12)
    };

    p.push(`Use the following computed results as the only quantitative basis for the brief.`);
    p.push(JSON.stringify(resultsBlock, null, 2));
    return p.join("\n\n");
  }

  function renderAiBriefing() {
    const text = buildAiBriefingText();
    const box = document.getElementById("aiBriefingText") || document.getElementById("copilotPreview");
    if (box && "value" in box) box.value = text;

    const jsonBox = document.getElementById("resultsJson") || document.getElementById("aiResultsJson");
    if (jsonBox && "value" in jsonBox) {
      const payload = buildResultsJsonPayload();
      jsonBox.value = JSON.stringify(payload, null, 2);
    }
  }

  async function copyToClipboard(text, successMsg, failMsg) {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(text);
        showToast(successMsg);
      } else {
        throw new Error("Clipboard API unavailable");
      }
    } catch {
      showToast(failMsg);
    }
  }

  // =========================
  // 11) IMPORT PIPELINE
  // =========================
  function renderImportSummary() {
    const el =
      document.getElementById("importSummary") ||
      document.getElementById("dataImportSummary") ||
      document.getElementById("importStatus");
    if (!el) return;

    const rows = state.dataset.rows || [];
    const schema = state.dataset.schema;
    const derived = state.dataset.derived;

    if (!rows.length || !schema) {
      el.textContent = "No dataset parsed yet.";
      return;
    }

    const parts = [];
    parts.push(`Rows parsed: ${rows.length.toLocaleString()}.`);
    parts.push(`Treatment column: ${schema.treatmentCol || "not found"}.`);
    parts.push(`Replicate column: ${schema.replicateCol || "not found"}.`);
    parts.push(`Yield column: ${schema.yieldCol || "not found"}.`);
    parts.push(`Cost columns: ${(schema.costCols || []).length.toLocaleString()}.`);
    if (derived && derived.controlKey) parts.push(`Detected control key: ${derived.controlKey}.`);
    if (state.dataset.committedAt) parts.push(`Committed at: ${state.dataset.committedAt}.`);
    el.textContent = parts.join(" ");
  }

  function setHeaderDatasetBadge() {
    const badge = document.getElementById("headerDatasetName");
    if (!badge) return;
    if (!state.dataset.rows.length) {
      badge.textContent = "None loaded";
      return;
    }
    const name = state.dataset.sourceName || "Imported dataset";
    const committed = state.dataset.committedAt ? " (committed)" : " (staged only)";
    badge.textContent = name + committed;
  }

  function parseAndStageFromText(rawText, sourceName) {
    const split = splitDictionaryAndDataFromText(rawText);
    const dict = parseDictionaryText(split.dictText);
    const dataText = split.dataText;

    const del = detectDelimiter(dataText);
    const tbl = parseDelimited(dataText, del);
    const rows = headersToObjects(tbl);

    state.dataset.sourceName = sourceName || "";
    state.dataset.rawText = rawText || "";
    state.dataset.dictionary = dict;
    state.dataset.rows = rows;
    state.dataset.schema = inferSchema(rows, dict);
    state.dataset.derived = computeDerivedFromDataset(rows, state.dataset.schema, dict);
    state.dataset.committedAt = null;

    renderImportSummary();
    renderDataChecks();
    setHeaderDatasetBadge();
    showToast("Dataset parsed and staged. Review Data Checks, then commit.");
  }

  function commitStagedDataset() {
    const derived = state.dataset.derived;
    const schema = state.dataset.schema;

    const errors = derived && derived.checks ? derived.checks.filter(c => c.severity === "error") : [];
    if (errors.length) {
      showToast("Cannot commit. Fix the error level data checks first.");
      renderDataChecks();
      return;
    }
    if (!schema || !schema.treatmentCol || !schema.yieldCol) {
      showToast("Cannot commit. Missing required columns.");
      renderDataChecks();
      return;
    }

    state.dataset.committedAt = new Date().toISOString();
    applyDatasetToModel();

    renderAll();
    setBasicsFieldsFromModel();
    calcAndRender();
    renderControlCentricResults();
    renderAiBriefing();
    setHeaderDatasetBadge();

    showToast("Dataset committed. Results updated.");
  }

  function initImportBindings() {
    const fileInput =
      document.getElementById("dataFile") ||
      document.getElementById("datasetFile") ||
      document.getElementById("uploadData") ||
      document.getElementById("uploadFile") ||
      document.getElementById("trialFile");

    const pasteBox =
      document.getElementById("dataPaste") ||
      document.getElementById("datasetPaste") ||
      document.getElementById("pasteData") ||
      document.getElementById("pasteBox");

    const parseBtn =
      document.getElementById("parseData") ||
      document.getElementById("parseImport") ||
      document.getElementById("parseDataset");

    const commitBtn =
      document.getElementById("commitData") ||
      document.getElementById("commitImport") ||
      document.getElementById("importCommit");

    if (fileInput) {
      fileInput.addEventListener("change", async e => {
        const f = e.target.files && e.target.files[0];
        if (!f) return;
        const text = await f.text();
        parseAndStageFromText(text, f.name);
        showToast("File loaded and parsed.");
        e.target.value = "";
      });
    }

    if (parseBtn && pasteBox) {
      parseBtn.addEventListener("click", e => {
        e.preventDefault();
        const text = String(pasteBox.value || "");
        if (!text.trim()) {
          showToast("Paste data is empty.");
          return;
        }
        parseAndStageFromText(text, "pasted_text");
      });
    }

    if (commitBtn) {
      commitBtn.addEventListener("click", e => {
        e.preventDefault();
        commitStagedDataset();
      });
    }

    document.addEventListener("click", e => {
      const btn = e.target.closest("[data-action]");
      if (!btn) return;
      const act = btn.getAttribute("data-action");
      if (!act) return;

      if (act === "parse-import") {
        e.preventDefault();
        const box =
          document.getElementById("dataPaste") ||
          document.getElementById("datasetPaste") ||
          document.getElementById("pasteData") ||
          document.getElementById("pasteBox");
        const text = box ? String(box.value || "") : "";
        if (!text.trim()) {
          showToast("Paste data is empty.");
          return;
        }
        parseAndStageFromText(text, "pasted_text");
        return;
      }

      if (act === "commit-import") {
        e.preventDefault();
        commitStagedDataset();
        return;
      }

      if (act === "export-cleaned-tsv") {
        e.preventDefault();
        exportCleanedDatasetTsv();
        return;
      }

      if (act === "export-treatment-summary-csv") {
        e.preventDefault();
        exportTreatmentSummaryCsv();
        return;
      }

      if (act === "run-sensitivity") {
        e.preventDefault();
        computeSensitivityGrid();
        renderSensitivitySummary();
        return;
      }

      if (act === "export-sensitivity-csv") {
        e.preventDefault();
        exportSensitivityGridCsv();
        return;
      }

      if (act === "export-workbook") {
        e.preventDefault();
        exportWorkbookIfAvailable();
        return;
      }

      if (act === "copy-ai-briefing") {
        e.preventDefault();
        const box = document.getElementById("aiBriefingText") || document.getElementById("copilotPreview");
        const txt = box && "value" in box ? String(box.value || "") : "";
        if (!txt.trim()) {
          showToast("AI briefing text is empty.");
          return;
        }
        copyToClipboard(txt, "AI briefing text copied.", "Unable to copy. Please copy manually.");
        return;
      }

      if (act === "copy-results-json") {
        e.preventDefault();
        const payload = buildResultsJsonPayload();
        const txt = JSON.stringify(payload, null, 2);
        copyToClipboard(txt, "Results JSON copied.", "Unable to copy. Please copy manually.");
        return;
      }
    });
  }

  // =========================
  // 12) CONFIG / SCENARIOS
  // =========================
  function setBasicsFieldsFromModel() {
    const m = model;

    const assignVal = (id, v) => {
      const el = document.getElementById(id);
      if (!el) return;
      if (el.type === "date") {
        el.value = v || "";
      } else {
        el.value = v != null ? v : "";
      }
    };

    assignVal("projectName", m.project.name);
    assignVal("organisation", m.project.organisation);
    assignVal("projectLead", m.project.lead);
    assignVal("analystNames", m.project.analysts);
    assignVal("projectTeam", m.project.team);
    assignVal("lastUpdated", m.project.lastUpdated);
    assignVal("projectSummary", m.project.summary);
    assignVal("projectObjectives", m.project.objectives);
    assignVal("projectActivities", m.project.activities);
    assignVal("stakeholderGroups", m.project.stakeholders);
    assignVal("contactEmail", m.project.contactEmail);
    assignVal("contactPhone", m.project.contactPhone);
    assignVal("projectGoal", m.project.goal);
    assignVal("withProject", m.project.withProject);
    assignVal("withoutProject", m.project.withoutProject);

    assignVal("years", m.time.years);
    assignVal("discBase", m.time.discBase);
    assignVal("grainPrice", getGrainPrice());
    assignVal("persistenceYears", state.config.persistenceYears);
    assignVal("mirrFinance", m.time.mirrFinance);
    assignVal("mirrReinvest", m.time.mirrReinvest);
    assignVal("startYear", m.time.startYear);
    assignVal("projectStartYear", m.time.projectStartYear);

    assignVal("adoptBase", m.adoption.base);
    assignVal("adoptLow", m.adoption.low);
    assignVal("adoptHigh", m.adoption.high);

    assignVal("riskBase", m.risk.base);
    assignVal("riskLow", m.risk.low);
    assignVal("riskHigh", m.risk.high);
    assignVal("rTech", m.risk.tech);
    assignVal("rNonCoop", m.risk.nonCoop);
    assignVal("rSocio", m.risk.socio);
    assignVal("rFin", m.risk.fin);
    assignVal("rMan", m.risk.man);

    const schedInputs = $$("input[data-disc-period][data-scenario]");
    schedInputs.forEach(inp => {
      const p = parseInt(inp.getAttribute("data-disc-period"), 10);
      const s = inp.getAttribute("data-scenario");
      const period = m.time.discountSchedule[p];
      if (!period) return;
      const key = s === "low" ? "low" : s === "high" ? "high" : "base";
      inp.value = period[key];
    });
  }

  function syncModelFromBasicsFields() {
    const getVal = id => {
      const el = document.getElementById(id);
      if (!el) return null;
      if (el.type === "number") return parseNumber(el.value);
      return el.value;
    };

    model.project.name = getVal("projectName") || model.project.name;
    model.project.organisation = getVal("organisation") || model.project.organisation;
    model.project.lead = getVal("projectLead") || model.project.lead;
    model.project.analysts = getVal("analystNames") || model.project.analysts;
    model.project.team = getVal("projectTeam") || model.project.team;
    const lu = getVal("lastUpdated");
    if (lu) model.project.lastUpdated = lu;
    model.project.summary = getVal("projectSummary") || model.project.summary;
    model.project.objectives = getVal("projectObjectives") || model.project.objectives;
    model.project.activities = getVal("projectActivities") || model.project.activities;
    model.project.stakeholders = getVal("stakeholderGroups") || model.project.stakeholders;
    model.project.contactEmail = getVal("contactEmail") || "";
    model.project.contactPhone = getVal("contactPhone") || "";
    model.project.goal = getVal("projectGoal") || model.project.goal;
    model.project.withProject = getVal("withProject") || model.project.withProject;
    model.project.withoutProject = getVal("withoutProject") || model.project.withoutProject;

    const years = parseNumber(getVal("years"));
    if (Number.isFinite(years) && years >= 0) model.time.years = Math.floor(years);

    const discBase = parseNumber(getVal("discBase"));
    if (Number.isFinite(discBase)) model.time.discBase = discBase;

    const grainPrice = parseNumber(getVal("grainPrice"));
    if (Number.isFinite(grainPrice)) {
      const yieldOut = ensureYieldOutput();
      yieldOut.value = grainPrice;
    }

    const persistence = parseNumber(getVal("persistenceYears"));
    if (Number.isFinite(persistence) && persistence >= 0) state.config.persistenceYears = Math.floor(persistence);

    const mirrFinance = parseNumber(getVal("mirrFinance"));
    if (Number.isFinite(mirrFinance)) model.time.mirrFinance = mirrFinance;
    const mirrReinvest = parseNumber(getVal("mirrReinvest"));
    if (Number.isFinite(mirrReinvest)) model.time.mirrReinvest = mirrReinvest;

    const startYear = parseNumber(getVal("startYear"));
    if (Number.isFinite(startYear)) model.time.startYear = Math.floor(startYear);
    const projStartYear = parseNumber(getVal("projectStartYear"));
    if (Number.isFinite(projStartYear)) model.time.projectStartYear = Math.floor(projStartYear);

    const adoptBase = parseNumber(getVal("adoptBase"));
    if (Number.isFinite(adoptBase)) model.adoption.base = adoptBase;
    const adoptLow = parseNumber(getVal("adoptLow"));
    if (Number.isFinite(adoptLow)) model.adoption.low = adoptLow;
    const adoptHigh = parseNumber(getVal("adoptHigh"));
    if (Number.isFinite(adoptHigh)) model.adoption.high = adoptHigh;

    const riskBase = parseNumber(getVal("riskBase"));
    if (Number.isFinite(riskBase)) model.risk.base = riskBase;
    const riskLow = parseNumber(getVal("riskLow"));
    if (Number.isFinite(riskLow)) model.risk.low = riskLow;
    const riskHigh = parseNumber(getVal("riskHigh"));
    if (Number.isFinite(riskHigh)) model.risk.high = riskHigh;
    const rTech = parseNumber(getVal("rTech"));
    if (Number.isFinite(rTech)) model.risk.tech = rTech;
    const rNonCoop = parseNumber(getVal("rNonCoop"));
    if (Number.isFinite(rNonCoop)) model.risk.nonCoop = rNonCoop;
    const rSocio = parseNumber(getVal("rSocio"));
    if (Number.isFinite(rSocio)) model.risk.socio = rSocio;
    const rFin = parseNumber(getVal("rFin"));
    if (Number.isFinite(rFin)) model.risk.fin = rFin;
    const rMan = parseNumber(getVal("rMan"));
    if (Number.isFinite(rMan)) model.risk.man = rMan;

    const schedInputs = $$("input[data-disc-period][data-scenario]");
    schedInputs.forEach(inp => {
      const p = parseInt(inp.getAttribute("data-disc-period"), 10);
      const s = inp.getAttribute("data-scenario");
      const val = parseNumber(inp.value);
      if (!Number.isFinite(val)) return;
      if (!model.time.discountSchedule[p]) return;
      const key = s === "low" ? "low" : s === "high" ? "high" : "base";
      model.time.discountSchedule[p][key] = val;
    });
  }

  function computeCombinedRisk() {
    const rTech = parseNumber(document.getElementById("rTech")?.value);
    const rNon = parseNumber(document.getElementById("rNonCoop")?.value);
    const rSoc = parseNumber(document.getElementById("rSocio")?.value);
    const rFin = parseNumber(document.getElementById("rFin")?.value);
    const rMan = parseNumber(document.getElementById("rMan")?.value);

    const parts = [rTech, rNon, rSoc, rFin, rMan].filter(v => Number.isFinite(v) && v >= 0 && v <= 1);
    if (!parts.length) return null;

    let product = 1;
    parts.forEach(p => {
      product *= 1 - p;
    });
    const combined = 1 - product;
    return combined;
  }

  function initConfigBindings() {
    const configInputs = [
      "projectName",
      "organisation",
      "projectLead",
      "analystNames",
      "projectTeam",
      "lastUpdated",
      "projectSummary",
      "projectObjectives",
      "projectActivities",
      "stakeholderGroups",
      "contactEmail",
      "contactPhone",
      "projectGoal",
      "withProject",
      "withoutProject",
      "years",
      "discBase",
      "grainPrice",
      "persistenceYears",
      "mirrFinance",
      "mirrReinvest",
      "startYear",
      "projectStartYear",
      "adoptBase",
      "adoptLow",
      "adoptHigh",
      "riskBase",
      "riskLow",
      "riskHigh",
      "rTech",
      "rNonCoop",
      "rSocio",
      "rFin",
      "rMan"
    ];

    configInputs.forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      el.addEventListener("change", () => {
        syncModelFromBasicsFields();
        calcAndRender();
      });
      el.addEventListener("blur", () => {
        syncModelFromBasicsFields();
        calcAndRender();
      });
    });

    const schedInputs = $$("input[data-disc-period][data-scenario]");
    schedInputs.forEach(inp => {
      inp.addEventListener("change", () => {
        syncModelFromBasicsFields();
        calcAndRender();
      });
    });

    const calcRiskBtn = document.getElementById("calcCombinedRisk");
    if (calcRiskBtn) {
      calcRiskBtn.addEventListener("click", () => {
        const combined = computeCombinedRisk();
        const out = document.getElementById("combinedRiskOut");
        if (!Number.isFinite(combined)) {
          if (out) out.textContent = "Enter risk multipliers between 0 and 1 to calculate combined risk.";
          return;
        }
        model.risk.base = combined;
        const baseEl = document.getElementById("riskBase");
        if (baseEl) baseEl.value = combined.toFixed(2);
        if (out) out.textContent = `Combined risk multiplier set to ${combined.toFixed(2)}.`;
        calcAndRender();
      });
    }

    const jumpBtn = document.querySelector("[data-tab-jump='results']");
    if (jumpBtn) {
      jumpBtn.addEventListener("click", () => {
        syncModelFromBasicsFields();
        calcAndRender();
        activateTab("results");
      });
    }

    const saveBtn = document.getElementById("saveScenario");
    const loadBtn = document.getElementById("loadScenario");
    const select = document.getElementById("scenarioSelect");

    if (saveBtn) {
      saveBtn.addEventListener("click", () => {
        saveScenario();
      });
    }
    if (loadBtn) {
      loadBtn.addEventListener("click", () => {
        const id = select ? select.value : "";
        if (!id) {
          showToast("Select a saved scenario to load.");
          return;
        }
        loadScenario(id);
      });
    }

    refreshScenarioSelect();
  }

  function readScenariosFromStorage() {
    try {
      const raw = localStorage.getItem(STORAGE_KEYS.scenarios);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      return parsed;
    } catch {
      return [];
    }
  }

  function writeScenariosToStorage(list) {
    try {
      localStorage.setItem(STORAGE_KEYS.scenarios, JSON.stringify(list || []));
    } catch {
      // ignore
    }
  }

  function refreshScenarioSelect() {
    const select = document.getElementById("scenarioSelect");
    if (!select) return;
    const scenarios = readScenariosFromStorage();
    const activeId = localStorage.getItem(STORAGE_KEYS.activeScenario) || "";

    select.innerHTML = `<option value="">Select a saved scenario</option>`;
    scenarios.forEach(s => {
      const opt = document.createElement("option");
      opt.value = s.id;
      opt.textContent = s.name;
      if (s.id === activeId) opt.selected = true;
      select.appendChild(opt);
    });
  }

  function captureScenarioPayload() {
    return {
      id: uid(),
      name: "",
      savedAt: new Date().toISOString(),
      model: JSON.parse(JSON.stringify(model)),
      stateConfig: JSON.parse(JSON.stringify(state.config)),
      datasetMeta: {
        sourceName: state.dataset.sourceName,
        committedAt: state.dataset.committedAt
      }
    };
  }

  function saveScenario() {
    const nameInput = document.getElementById("scenarioName");
    const label = nameInput && nameInput.value.trim() ? nameInput.value.trim() : "";
    const payload = captureScenarioPayload();
    payload.name = label || `Scenario (${new Date().toLocaleString()})`;

    const list = readScenariosFromStorage();
    list.push(payload);
    writeScenariosToStorage(list);
    localStorage.setItem(STORAGE_KEYS.activeScenario, payload.id);
    refreshScenarioSelect();
    showToast("Scenario saved.");
  }

  function loadScenario(id) {
    const list = readScenariosFromStorage();
    const s = list.find(x => x.id === id);
    if (!s) {
      showToast("Scenario not found.");
      return;
    }

    Object.assign(model, s.model);
    Object.assign(state.config, s.stateConfig || state.config);

    setBasicsFieldsFromModel();
    renderOutputsList();
    renderTreatmentsList();
    calcAndRender();
    renderImportSummary();
    renderDataChecks();
    renderSensitivitySummary();
    refreshScenarioSelect();
    showToast("Scenario loaded.");
  }

  // =========================
  // 13) OUTPUTS + TREATMENTS UI
  // =========================
  function renderOutputsList() {
    const root = document.getElementById("outputsList");
    if (!root) return;

    if (!model.outputs.length) {
      root.innerHTML = `<p class="small muted">No outputs defined yet.</p>`;
      return;
    }

    const html = model.outputs
      .map(
        o => `
      <div class="list-item output-row" data-output-id="${esc(o.id)}">
        <div class="title">
          <strong>${esc(o.name)}</strong>
          <span class="small muted">Unit: ${esc(o.unit || "")}</span>
        </div>
        <div class="kv">
          <div class="k">Name</div>
          <div><input type="text" data-field="name" value="${esc(o.name)}"></div>

          <div class="k">Unit</div>
          <div><input type="text" data-field="unit" value="${esc(o.unit || "")}"></div>

          <div class="k">Value</div>
          <div><input type="number" step="0.01" data-field="value" value="${o.value != null ? o.value : ""}"></div>

          <div class="k">Source</div>
          <div><input type="text" data-field="source" value="${esc(o.source || "")}"></div>
        </div>
      </div>
    `
      )
      .join("");

    root.innerHTML = `<div class="list">${html}</div>`;
  }

  function renderTreatmentsList() {
    const root = document.getElementById("treatmentsList");
    if (!root) return;

    if (!model.treatments.length) {
      root.innerHTML = `<p class="small muted">No treatments defined yet. Import a dataset to calibrate treatments.</p>`;
      return;
    }

    const yieldOut = ensureYieldOutput();
    const html = model.treatments
      .map(t => {
        const isControl = t.isControl;
        const recurrence = getRecurrenceYears(t);
        const dy = t.deltas ? t.deltas[yieldOut.id] : 0;
        return `
        <div class="list-item treatment-row" data-treatment-id="${esc(t.id)}">
          <div class="title">
            <strong>${esc(t.name)}</strong>
            <span class="small muted">${isControl ? "Control (baseline)" : "Treatment"} · Area ${fmt(
          t.area || 0
        )} ha</span>
          </div>
          <div class="kv">
            <div class="k">Area (ha)</div>
            <div><input type="number" step="1" data-field="area" value="${t.area != null ? t.area : ""}"></div>

            <div class="k">Adoption multiplier</div>
            <div><input type="number" step="0.01" min="0" max="1" data-field="adoption" value="${
              t.adoption != null ? t.adoption : ""
            }"></div>

            <div class="k">Recurrence (years; 0 = once at year 0)</div>
            <div><input type="number" step="1" min="0" data-field="recurrenceYears" value="${recurrence}"></div>

            <div class="k">Yield delta vs control (t/ha)</div>
            <div><input type="number" step="0.01" data-field="deltaYield" value="${dy != null ? dy : ""}"></div>

            <div class="k">Labour cost per ha</div>
            <div><input type="number" step="0.01" data-field="labourCost" value="${
              t.labourCost != null ? t.labourCost : ""
            }"></div>

            <div class="k">Materials cost per ha</div>
            <div><input type="number" step="0.01" data-field="materialsCost" value="${
              t.materialsCost != null ? t.materialsCost : ""
            }"></div>

            <div class="k">Services cost per ha</div>
            <div><input type="number" step="0.01" data-field="servicesCost" value="${
              t.servicesCost != null ? t.servicesCost : ""
            }"></div>

            <div class="k">Capital cost (year 0)</div>
            <div><input type="number" step="0.01" data-field="capitalCost" value="${
              t.capitalCost != null ? t.capitalCost : ""
            }"></div>

            <div class="k">Notes</div>
            <div><textarea data-field="notes" rows="2">${esc(t.notes || "")}</textarea></div>
          </div>
        </div>
      `;
      })
      .join("");

    root.innerHTML = `<div class="list">${html}</div>`;
  }

  function initOutputsTreatmentsBindings() {
    const outputsRoot = document.getElementById("outputsList");
    if (outputsRoot) {
      outputsRoot.addEventListener("input", e => {
        const row = e.target.closest(".output-row");
        if (!row) return;
        const id = row.getAttribute("data-output-id");
        const field = e.target.getAttribute("data-field");
        if (!id || !field) return;
        const out = model.outputs.find(o => o.id === id);
        if (!out) return;

        if (field === "value") {
          const v = parseNumber(e.target.value);
          out.value = Number.isFinite(v) ? v : out.value;
        } else {
          out[field] = e.target.value;
        }
        calcAndRender();
      });
    }

    const treatRoot = document.getElementById("treatmentsList");
    if (treatRoot) {
      treatRoot.addEventListener("input", e => {
        const row = e.target.closest(".treatment-row");
        if (!row) return;
        const id = row.getAttribute("data-treatment-id");
        const field = e.target.getAttribute("data-field");
        if (!id || !field) return;
        const t = model.treatments.find(x => x.id === id);
        if (!t) return;

        if (field === "notes") {
          t.notes = e.target.value;
        } else if (field === "deltaYield") {
          const v = parseNumber(e.target.value);
          const yieldOut = ensureYieldOutput();
          if (!t.deltas) t.deltas = {};
          t.deltas[yieldOut.id] = Number.isFinite(v) ? v : t.deltas[yieldOut.id];
        } else if (field === "area") {
          const v = parseNumber(e.target.value);
          t.area = Number.isFinite(v) ? v : t.area;
        } else if (field === "adoption") {
          const v = parseNumber(e.target.value);
          t.adoption = Number.isFinite(v) ? clamp(v, 0, 1) : t.adoption;
        } else if (field === "recurrenceYears") {
          const v = parseNumber(e.target.value);
          t.recurrenceYears = Number.isFinite(v) ? Math.max(0, Math.floor(v)) : t.recurrenceYears;
        } else if (field === "labourCost" || field === "materialsCost" || field === "servicesCost" || field === "capitalCost") {
          const v = parseNumber(e.target.value);
          t[field] = Number.isFinite(v) ? v : t[field];
        }

        calcAndRender();
      });
    }
  }

  // =========================
  // 14) SENSITIVITY SUMMARY
  // =========================
  function renderSensitivitySummary() {
    const root = document.getElementById("sensitivitySummary");
    if (!root) return;

    const grid = state.results.sensitivityGrid || [];
    if (!grid.length) {
      root.innerHTML = `<p class="small muted">No sensitivity grid has been run yet. Use “Run sensitivity” to compute a grid across grain price, discount rate, persistence and recurrence.</p>`;
      return;
    }

    const n = grid.length;
    const treatments = Array.from(new Set(grid.map(g => g.treatment)));
    const prices = Array.from(new Set(grid.map(g => g.pricePerTonne))).sort((a, b) => a - b);
    const discs = Array.from(new Set(grid.map(g => g.discountRatePct))).sort((a, b) => a - b);

    const best = grid
      .slice()
      .filter(g => Number.isFinite(g.npv))
      .sort((a, b) => b.npv - a.npv)[0];

    const worst = grid
      .slice()
      .filter(g => Number.isFinite(g.npv))
      .sort((a, b) => a.npv - b.npv)[0];

    const parts = [];
    parts.push(
      `The current sensitivity grid contains ${fmt(
        n
      )} combinations of grain price, discount rate, persistence, and recurrence across ${treatments.length} treatments.`
    );
    parts.push(
      `Grain price scenarios range from ${money(prices[0])} per tonne to ${money(
        prices[prices.length - 1]
      )} per tonne. Discount rate scenarios range from ${fmt(discs[0])} percent to ${fmt(discs[discs.length - 1])} percent.`
    );
    if (best && worst) {
      parts.push(
        `Across all scenarios, the highest net present value occurs for ${best.treatment} with NPV ${money(
          best.npv
        )}. The lowest net present value occurs for ${worst.treatment} with NPV ${money(
          worst.npv
        )}. These extremes help show how sensitive results are to price and discount assumptions.`
      );
    }

    root.textContent = parts.join(" ");
  }

  // =========================
  // 15) RESULTS FILTERS + CHARTS
  // =========================
  function initResultsFilterBindings() {
    const select = document.getElementById("resultsFilter");
    if (select) {
      select.addEventListener("change", () => {
        state.results.currentFilter = select.value || "all";
        renderControlCentricResults();
      });
    }

    const btnTopNpv = document.getElementById("filterTopNpv");
    const btnTopBcr = document.getElementById("filterTopBcr");
    const btnImprove = document.getElementById("filterImproveOnly");
    const btnAll = document.getElementById("filterShowAll");

    if (btnTopNpv) {
      btnTopNpv.addEventListener("click", () => {
        state.results.currentFilter = "top5_npv";
        if (select) select.value = "top5_npv";
        renderControlCentricResults();
      });
    }
    if (btnTopBcr) {
      btnTopBcr.addEventListener("click", () => {
        state.results.currentFilter = "top5_bcr";
        if (select) select.value = "top5_bcr";
        renderControlCentricResults();
      });
    }
    if (btnImprove) {
      btnImprove.addEventListener("click", () => {
        state.results.currentFilter = "improve_only";
        if (select) select.value = "improve_only";
        renderControlCentricResults();
      });
    }
    if (btnAll) {
      btnAll.addEventListener("click", () => {
        state.results.currentFilter = "all";
        if (select) select.value = "all";
        renderControlCentricResults();
      });
    }
  }

  function ensureChartsCard() {
    const page = document.querySelector("#tab-results .page");
    if (!page) return null;
    let card = document.getElementById("resultsChartsCard");
    if (!card) {
      card = document.createElement("div");
      card.id = "resultsChartsCard";
      card.className = "card";
      card.innerHTML = `
        <h2>Visual summary</h2>
        <p class="small muted">
          Bars show how each treatment compares with the control in terms of net present value and the split between benefits and costs.
          Hover tooltips are not required. Read the labels below each chart.
        </p>
        <canvas id="npvChart" height="220"></canvas>
        <canvas id="pvChart" height="220" style="margin-top:16px;"></canvas>
      `;
      page.appendChild(card);
    }
    return card;
  }

  function renderCharts(perTreatment, filterMode) {
    const card = ensureChartsCard();
    if (!card) return;

    const list = filterTreatments(perTreatment, filterMode)
      .slice()
      .sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity))
      .slice(0, 8); // top 8 for readability

    const npvCanvas = document.getElementById("npvChart");
    const pvCanvas = document.getElementById("pvChart");
    if (!npvCanvas || !pvCanvas) return;

    const dpr = window.devicePixelRatio || 1;
    const width = card.clientWidth - 32;
    const chartWidth = Math.max(300, width);

    [npvCanvas, pvCanvas].forEach(c => {
      c.style.width = chartWidth + "px";
      c.width = chartWidth * dpr;
      c.height = 220 * dpr;
    });

    const ctxN = npvCanvas.getContext("2d");
    const ctxP = pvCanvas.getContext("2d");
    ctxN.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctxP.setTransform(dpr, 0, 0, dpr, 0, 0);

    ctxN.clearRect(0, 0, chartWidth, 220);
    ctxP.clearRect(0, 0, chartWidth, 220);

    if (!list.length) {
      ctxN.fillStyle = "rgba(255,255,255,0.75)";
      ctxN.font = "12px system-ui, sans-serif";
      ctxN.fillText("Charts will appear here once there are treatments with results.", 10, 30);
      ctxP.fillStyle = "rgba(255,255,255,0.75)";
      ctxP.font = "12px system-ui, sans-serif";
      ctxP.fillText("Charts will appear here once there are treatments with results.", 10, 30);
      return;
    }

    // NPV bar chart
    const margin = { top: 10, right: 10, bottom: 50, left: 60 };
    const w = chartWidth - margin.left - margin.right;
    const h = 220 - margin.top - margin.bottom;

    const npvs = list.map(r => r.npv || 0);
    const maxAbs = Math.max(...npvs.map(v => Math.abs(v)), 1);
    const scaleY = v => h / 2 - (v / (maxAbs * 1.1)) * (h / 2);

    ctxN.save();
    ctxN.translate(margin.left, margin.top);

    ctxN.strokeStyle = "rgba(255,255,255,0.25)";
    ctxN.beginPath();
    ctxN.moveTo(0, scaleY(0));
    ctxN.lineTo(w, scaleY(0));
    ctxN.stroke();

    const barWidth = w / (list.length * 1.5);
    list.forEach((r, i) => {
      const x = i * (w / list.length) + (w / list.length - barWidth) / 2;
      const y0 = scaleY(0);
      const y1 = scaleY(r.npv || 0);
      const height = Math.abs(y1 - y0);

      ctxN.fillStyle = (r.npv || 0) >= 0 ? "rgba(52,211,153,0.80)" : "rgba(251,113,133,0.80)";
      ctxN.fillRect(x, Math.min(y0, y1), barWidth, height);

      ctxN.fillStyle = "rgba(255,255,255,0.85)";
      ctxN.font = "11px system-ui, sans-serif";
      const label = r.treatmentName.length > 12 ? r.treatmentName.slice(0, 11) + "…" : r.treatmentName;
      ctxN.save();
      ctxN.translate(x + barWidth / 2, h + 12);
      ctxN.rotate(-Math.PI / 4);
      ctxN.textAlign = "left";
      ctxN.fillText(label, 0, 0);
      ctxN.restore();
    });

    ctxN.fillStyle = "rgba(255,255,255,0.85)";
    ctxN.font = "12px system-ui, sans-serif";
    ctxN.textAlign = "left";
    ctxN.fillText("Net present value relative to control", 0, 12);
    ctxN.restore();

    // PV benefits vs PV costs stacked bars
    ctxP.save();
    ctxP.translate(margin.left, margin.top);

    const pvVals = list.flatMap(r => [r.pvBenefits || 0, r.pvCosts || 0]);
    const maxPv = Math.max(...pvVals.map(v => Math.abs(v)), 1);
    const scaleY2 = v => h - (v / (maxPv * 1.1)) * h;

    list.forEach((r, i) => {
      const x = i * (w / list.length) + (w / list.length - barWidth) / 2;
      const baseY = h;

      const bHeight = h - scaleY2(r.pvBenefits || 0);
      const cHeight = h - scaleY2(r.pvCosts || 0);

      ctxP.fillStyle = "rgba(125,211,252,0.85)";
      ctxP.fillRect(x, baseY - bHeight, barWidth, bHeight);

      ctxP.fillStyle = "rgba(96,165,250,0.65)";
      ctxP.fillRect(x, baseY - bHeight - cHeight, barWidth, cHeight);

      ctxP.fillStyle = "rgba(255,255,255,0.85)";
      ctxP.font = "11px system-ui, sans-serif";
      const label = r.treatmentName.length > 12 ? r.treatmentName.slice(0, 11) + "…" : r.treatmentName;
      ctxP.save();
      ctxP.translate(x + barWidth / 2, h + 12);
      ctxP.rotate(-Math.PI / 4);
      ctxP.textAlign = "left";
      ctxP.fillText(label, 0, 0);
      ctxP.restore();
    });

    ctxP.fillStyle = "rgba(255,255,255,0.85)";
    ctxP.font = "12px system-ui, sans-serif";
    ctxP.textAlign = "left";
    ctxP.fillText("Present value of benefits (light) and costs (darker) per treatment", 0, 12);
    ctxP.restore();
  }

  // =========================
  // 16) CONTROL-CENTRIC RESULTS WRAPPER
  // =========================
  function renderControlCentricResults() {
    const { perTreatment } = { perTreatment: state.results.perTreatmentBaseCase || [] };
    if (!perTreatment.length) {
      const lb = document.getElementById("resultsLeaderboard");
      const c2c = document.getElementById("comparisonToControl");
      const nar = document.getElementById("resultsNarrative");
      if (lb) lb.innerHTML = `<p class="small muted">Results will appear here once a dataset has been committed.</p>`;
      if (c2c) c2c.innerHTML = `<p class="small muted">Comparison table will appear here once a dataset has been committed.</p>`;
      if (nar) nar.textContent = "";
      return;
    }
    const filterMode = state.results.currentFilter || "all";
    renderLeaderboard(perTreatment, filterMode);
    renderComparisonToControl(perTreatment, filterMode);
    renderResultsNarrative(perTreatment, filterMode);
    renderCharts(perTreatment, filterMode);
  }

  function calcAndRender() {
    computeBaseCaseResultsVsControl();
    renderControlCentricResults();
    renderAiBriefing();
  }

  function renderAll() {
    renderImportSummary();
    renderDataChecks();
    renderOutputsList();
    renderTreatmentsList();
    renderSensitivitySummary();
    setHeaderDatasetBadge();
  }

  // =========================
  // 17) TABS + TECHNICAL APPENDIX NEW TAB
  // =========================
  function activateTab(name) {
    const tabs = $$(".tab");
    const panels = $$(".tab-panel");

    tabs.forEach(btn => {
      const t = btn.getAttribute("data-tab");
      const active = t === name;
      btn.classList.toggle("active", active);
      btn.setAttribute("aria-selected", active ? "true" : "false");
    });

    panels.forEach(panel => {
      const t = panel.getAttribute("data-tab-panel");
      const active = t === name;
      panel.hidden = !active;
      panel.setAttribute("aria-hidden", active ? "false" : "true");
    });
  }

  function openTechnicalAppendixWindow() {
    const panel = document.getElementById("tab-appendix");
    if (!panel) return;
    const win = window.open("", "_blank");
    if (!win) {
      showToast("Popup blocked. Allow popups for this site to open the Technical Appendix in a new tab.");
      return;
    }

    const html = `
      <!doctype html>
      <html lang="en">
      <head>
        <meta charset="utf-8" />
        <title>Technical Appendix — Farming CBA Decision Tool</title>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <link rel="stylesheet" href="styles.css" />
      </head>
      <body>
        <div class="app-shell">
          <header class="app-header">
            <div class="brand">
              <div class="brand-mark" aria-hidden="true">NBS</div>
              <div class="brand-text">
                <div class="brand-title">Technical Appendix</div>
                <div class="brand-subtitle">Farming CBA Decision Tool</div>
              </div>
            </div>
          </header>
          <main class="tab-panels">
            <section class="tab-panel active">
              ${panel.innerHTML}
            </section>
          </main>
        </div>
      </body>
      </html>
    `;
    win.document.open();
    win.document.write(html);
    win.document.close();
    win.focus();
  }

  function initTabs() {
    const tabs = $$(".tab");
    tabs.forEach(btn => {
      btn.addEventListener("click", () => {
        const t = btn.getAttribute("data-tab");
        if (!t) return;
        if (t === "appendix") {
          openTechnicalAppendixWindow();
          activateTab(t); // also show in main app for accessibility
        } else {
          activateTab(t);
        }
      });
    });

    activateTab("results");
  }

  // =========================
  // 18) DEFAULT DATASET LOADER
  // =========================
  function loadDefaultDatasetIfEmpty() {
    if (state.dataset.rows && state.dataset.rows.length) return;
    const txt = DEFAULT_TRIAL_DATA_TSV && DEFAULT_TRIAL_DATA_TSV.trim();
    if (!txt) return;
    parseAndStageFromText(txt, "Built-in example trial dataset");
    commitStagedDataset();
  }

  // =========================
  // 19) INIT
  // =========================
  document.addEventListener("DOMContentLoaded", () => {
    initTabs();
    initImportBindings();
    initConfigBindings();
    initOutputsTreatmentsBindings();
    initResultsFilterBindings();

    setBasicsFieldsFromModel();
    loadDefaultDatasetIfEmpty();
    renderAll();
    calcAndRender();

    window.addEventListener("resize", () => {
      const per = state.results.perTreatmentBaseCase || [];
      if (!per.length) return;
      const filterMode = state.results.currentFilter || "all";
      renderCharts(per, filterMode);
    });
  });
})();
