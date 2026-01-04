// app.js
// Farming CBA Tool - Newcastle Business School
// Production-grade, control-centric CBA decision support with:
// - Default trial data loaded automatically (via local fetch, same pipeline as upload/paste) and replaceable via import.
// - Robust import pipeline: upload + paste TSV/CSV + dictionary splitting/parsing, schema inference, validation, staging, commit.
// - Replicate-specific control baselines, plot-level deltas, treatment summaries with missing-safe stats.
// - Discounted CBA engine + sensitivity grid + exports (TSV/CSV/XLSX when available).
// - Clear mobile-first layouts and responsive, informative plots.
// - Technical Appendix opens in a new tab (generated from current dataset + settings).
// - AI briefing prompt + results JSON, copy actions, and toasts for major actions.

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

  const STORAGE_KEYS = {
    scenarios: "farming_cba_scenarios_v2",
    activeScenario: "farming_cba_active_scenario_v2",
    appendixPayload: "farming_cba_appendix_payload_v2"
  };

  const DEFAULT_SENS_PRICE = [300, 350, 400, 450, 500, 550, 600];
  const DEFAULT_SENS_DISC = [2, 4, 7, 10, 12];
  const DEFAULT_SENS_PERSIST = [1, 2, 3, 5, 7, 10];
  const DEFAULT_SENS_RECURRENCE = [1, 2, 3, 4, 5, 7, 10, 0]; // 0 = once only at year 0

  // Default trial data candidates (place these files alongside index.html on GitHub Pages)
  const DEFAULT_DATA_CANDIDATES = [
    "faba_beans_trial_clean_named.tsv",
    "trial_data.tsv",
    "data.tsv"
  ];

  // Optional dictionary candidates (if present, improves schema inference and appendix detail)
  const DEFAULT_DICT_CANDIDATES = [
    "faba_beans_trial_data_dictionary_FULL.csv",
    "data_dictionary.csv",
    "dictionary.csv"
  ];

  // =========================
  // 1) UTILITIES
  // =========================
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));

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

  const esc = s =>
    (s ?? "")
      .toString()
      .replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const slug = s =>
    (s || "project")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_|_$/g, "");

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

  function mean(arr) {
    const a = arr.filter(v => Number.isFinite(v));
    if (!a.length) return NaN;
    return a.reduce((s, v) => s + v, 0) / a.length;
  }

  function median(arr) {
    const a = arr.filter(v => Number.isFinite(v)).slice().sort((x, y) => x - y);
    if (!a.length) return NaN;
    const mid = Math.floor(a.length / 2);
    return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
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

  // =========================
  // 2) TOASTS
  // =========================
  function ensureToastRoot() {
    if (document.getElementById("toast-root")) return;
    const div = document.createElement("div");
    div.id = "toast-root";
    div.setAttribute("aria-live", "polite");
    div.setAttribute("aria-atomic", "true");
    document.body.appendChild(div);
  }

  function showToast(title, body, tone) {
    ensureToastRoot();
    const root = document.getElementById("toast-root") || document.body;
    const toast = document.createElement("div");
    toast.className = "toast" + (tone ? " " + tone : "");
    toast.innerHTML = `
      <div class="t-title">${esc(title || "Update")}</div>
      ${body ? `<div class="t-body">${esc(body)}</div>` : ""}
    `;
    root.appendChild(toast);
    setTimeout(() => toast.remove(), 4200);
  }

  // =========================
  // 3) MODEL
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
        "Growers adopt high-performing amendment packages on trial farms and similar soils in the region.",
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
    outputs: [
      { id: uid(), name: "Grain yield", unit: "t/ha", value: 450, source: "Input Directly" }
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
        source: "Farm Trials",
        isControl: true,
        notes: "Control definition is taken from the dataset where available.",
        recurrenceYears: 0
      }
    ],
    adoption: { base: 1.0, low: 0.6, high: 1.0 },
    risk: { base: 0.15, low: 0.05, high: 0.3, tech: 0.05, nonCoop: 0.04, socio: 0.02, fin: 0.03, man: 0.02 }
  };

  function ensureYieldOutput() {
    let out = model.outputs.find(o => String(o.name || "").toLowerCase().includes("yield"));
    if (!out) {
      out = { id: uid(), name: "Grain yield", unit: "t/ha", value: 450, source: "Input Directly" };
      model.outputs.unshift(out);
    }
    return out;
  }

  function initTreatmentDeltas() {
    const out = ensureYieldOutput();
    model.treatments.forEach(t => {
      if (!t.deltas) t.deltas = {};
      if (!(out.id in t.deltas)) t.deltas[out.id] = 0;
      if (typeof t.labourCost === "undefined") t.labourCost = 0;
      if (typeof t.materialsCost === "undefined") t.materialsCost = 0;
      if (typeof t.servicesCost === "undefined") t.servicesCost = 0;
      if (typeof t.capitalCost === "undefined") t.capitalCost = 0;
      if (typeof t.area === "undefined") t.area = 100;
      if (typeof t.recurrenceYears === "undefined") t.recurrenceYears = 0;
    });
  }
  initTreatmentDeltas();

  // =========================
  // 4) STATE
  // =========================
  const state = {
    dataset: {
      sourceName: "",
      rawText: "",
      rows: [],
      dictionary: null,
      schema: null,
      staged: null,
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
      lastComputedAt: null
    },
    ui: {
      resultsFilter: "all",
      activeTab: "results"
    }
  };

  // =========================
  // 5) PARSING (CSV/TSV + DICTIONARY SPLIT)
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
    return String(h ?? "").trim().replace(/\s+/g, " ").replace(/[^\S\r\n]+/g, " ");
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
      dictRoleMatch("treatment") || bestHeader(["amendment", "treatment", "variant", "package", "option", "arm"]);
    const replicateCol =
      dictRoleMatch("replicate") || bestHeader(["replicate", "rep", "block", "trial block", "replication"]);
    const plotCol = dictRoleMatch("plot") || bestHeader(["plot", "plot id", "plotid", "plot_no", "plot number"]);
    const controlFlagCol = dictRoleMatch("control") || bestHeader(["is_control", "control", "baseline"]);
    const yieldCol = dictRoleMatch("yield") || bestHeader(["yield t/ha", "yield", "grain yield", "yield_tha", "yield (t/ha)"]);

    const costCols = headers.filter(h => {
      const s = h.toLowerCase();
      const isCosty =
        s.includes("cost") || s.includes("labour") || s.includes("labor") || s.includes("input") || s.includes("fert") ||
        s.includes("herb") || s.includes("fung") || s.includes("insect") || s.includes("fuel") || s.includes("machinery") ||
        s.includes("spray") || s.includes("seed") || s.includes("gyro") || s.includes("gypsum");
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
          if (n > bestN) { best = k; bestN = n; }
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

    if (!schema.treatmentCol) checks.push({ code: "NO_TREATMENT_COL", severity: "error", message: "Treatment column not found.", count: 0, detail: "" });
    if (!schema.replicateCol) checks.push({ code: "NO_REPLICATE_COL", severity: "warn", message: "Replicate column not found. Baselines fall back to overall control mean.", count: 0, detail: "" });
    if (!schema.yieldCol) checks.push({ code: "NO_YIELD_COL", severity: "error", message: "Yield column not found.", count: 0, detail: "" });

    const controlKey = detectControlKey(rows, schema);
    derived.controlKey = controlKey;
    if (!controlKey) checks.push({ code: "NO_CONTROL_DETECTED", severity: "error", message: "Control treatment could not be detected. Provide an is_control column or ensure the label includes the word control.", count: 0, detail: "" });

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
      (schema.costCols || []).forEach(c => { costByCol[c] = costPerHaFromRow(r, schema, c); });

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
    if (missingYield) checks.push({ code: "MISSING_YIELD", severity: "warn", message: "Some rows have missing yield values. Excluded from yield summaries.", count: missingYield, detail: "" });

    const negYield = cleaned.filter(r => Number.isFinite(r.yield) && r.yield < 0).length;
    if (negYield) checks.push({ code: "NEGATIVE_YIELD", severity: "warn", message: "Some rows have negative yield values. Check units or data entry.", count: negYield, detail: "" });

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
      checks.push({ code: "CONTROL_YIELD_MISSING", severity: "error", message: "Control yields are missing. Cannot compute deltas.", count: 0, detail: "" });
    }

    const replicateBaselines = new Map();
    for (const [repKey, entry] of reps.entries()) {
      const yMean = mean(entry.ctrlY);
      const costsMean = {};
      (schema.costCols || []).forEach(c => { costsMean[c] = mean(entry.ctrlCostsByCol.get(c) || []); });
      replicateBaselines.set(repKey, {
        yieldMean: Number.isFinite(yMean) ? yMean : overallCtrlMeanYield,
        costsMeanByCol: costsMean
      });
    }
    derived.replicateBaselines = replicateBaselines;

    if (schema.replicateCol) {
      const allRepKeys = new Set(cleaned.map(r => r.replicate || "__MISSING_REP__"));
      let repsNoCtrl = 0;
      allRepKeys.forEach(k => { if (!replicateBaselines.has(k)) repsNoCtrl++; });
      if (repsNoCtrl) {
        checks.push({
          code: "REPS_WITHOUT_CONTROL",
          severity: "warn",
          message: "Some replicates have no control rows. Baselines fall back to overall control mean.",
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
      return { ...r, controlYieldMeanRep: base.yieldMean, deltaYield: dy, deltaCostsPerHa: dCosts };
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
        detail: Number.isFinite(outFlags.low) && Number.isFinite(outFlags.high) ? `IQR bounds are ${fmt(outFlags.low)} to ${fmt(outFlags.high)} t/ha.` : ""
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
  // 6) IMPORT PIPELINE (STAGE + COMMIT)
  // =========================
  function stageDatasetFromText(rawText, sourceName) {
    const raw = String(rawText || "");
    const { dictText, dataText } = splitDictionaryAndDataFromText(raw);

    const dict = parseDictionaryText(dictText);
    const del = detectDelimiter(dataText);
    const tbl = parseDelimited(dataText, del);
    const objs = headersToObjects(tbl);

    const schema = inferSchema(objs, dict);
    const derived = computeDerivedFromDataset(objs, schema);

    state.dataset.staged = {
      sourceName: sourceName || "Staged dataset",
      rawText: raw,
      rows: objs,
      dictionary: dict,
      schema,
      derived
    };

    renderDataChecks(derived);
    renderImportSummary(derived, schema, state.dataset.staged.sourceName, true);

    const errs = derived.checks.filter(c => c.severity === "error").length;
    showToast("Dataset staged", errs ? `Staged with ${errs} error-level checks. Fix before commit.` : "Staged and ready to commit.", errs ? "warn" : "good");
  }

  function commitStagedDataset() {
    const staged = state.dataset.staged;
    if (!staged) {
      showToast("Nothing to commit", "Parse and stage a dataset first.", "warn");
      return;
    }
    const errs = staged.derived.checks.filter(c => c.severity === "error").length;
    if (errs) {
      showToast("Commit blocked", "Resolve error-level data checks, then stage again.", "bad");
      return;
    }

    state.dataset.sourceName = staged.sourceName;
    state.dataset.rawText = staged.rawText;
    state.dataset.rows = staged.rows;
    state.dataset.dictionary = staged.dictionary;
    state.dataset.schema = staged.schema;
    state.dataset.derived = staged.derived;
    state.dataset.committedAt = new Date().toISOString();

    applyDatasetToModel();
    recomputeAndRenderAll("Committed dataset");

    renderImportSummary(state.dataset.derived, state.dataset.schema, state.dataset.sourceName, false);
    showToast("Dataset committed", "Treatments and results updated.", "good");
  }

  async function tryFetchFirstExisting(paths) {
    for (const p of paths) {
      try {
        const res = await fetch(p, { cache: "no-store" });
        if (res.ok) {
          const txt = await res.text();
          if (String(txt || "").trim().length > 0) return { path: p, text: txt };
        }
      } catch (e) {}
    }
    return null;
  }

  async function loadDefaultTrialDataIfAvailable() {
    setHeaderDataset("Loading default trial data…");

    const dataHit = await tryFetchFirstExisting(DEFAULT_DATA_CANDIDATES);
    if (!dataHit) {
      setHeaderDataset("None loaded");
      renderImportStatus("No default trial file found. You can still upload or paste your dataset in Data Import.");
      showToast("Default trial data not found", "Place a TSV/CSV file next to index.html (for example faba_beans_trial_clean_named.tsv).", "warn");
      return;
    }

    const dictHit = await tryFetchFirstExisting(DEFAULT_DICT_CANDIDATES);

    let combinedText = dataHit.text;
    if (dictHit && dictHit.text) {
      combinedText = dictHit.text.trim() + "\n\n" + dataHit.text.trim();
    }

    stageDatasetFromText(combinedText, `Default trial data (${dataHit.path})`);
    commitStagedDataset();

    setHeaderDataset(`Default: ${dataHit.path}`);
    showToast("Default trial data loaded", "Ready to use immediately. You can replace it via Data Import.", "good");
  }

  async function stageDatasetFromUploadFile(file) {
    if (!file) {
      showToast("No file selected", "Choose a TSV, CSV, or Excel file.", "warn");
      return;
    }
    const name = file.name || "Uploaded file";
    const lower = name.toLowerCase();

    // Excel support if XLSX is available
    if ((lower.endsWith(".xlsx") || lower.endsWith(".xls")) && typeof XLSX !== "undefined") {
      try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const sheetName = wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
        const tsv = aoa.map(r => r.map(x => String(x ?? "")).join("\t")).join("\n");
        stageDatasetFromText(tsv, `Upload: ${name}`);
        showToast("Excel staged", "Review checks and commit when ready.", "good");
        return;
      } catch (e) {
        showToast("Excel parse failed", "Could not read the workbook. Try exporting as TSV/CSV and upload again.", "bad");
        return;
      }
    }

    // Text formats
    try {
      const txt = await file.text();
      stageDatasetFromText(txt, `Upload: ${name}`);
      showToast("File staged", "Review checks and commit when ready.", "good");
    } catch (e) {
      showToast("Upload failed", "Could not read the selected file.", "bad");
    }
  }

  // =========================
  // 7) APPLY DATASET TO MODEL (calibrate treatments)
  // =========================
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
      const isServ =
        h.includes("contract") || h.includes("service") || h.includes("hire") ||
        h.includes("machinery") || h.includes("fuel") || h.includes("spray") || h.includes("tractor");
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
      showToast("No summary to apply", "Import and commit a dataset first.", "warn");
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
        newTreatments.push({
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
          notes: "Incremental values are replicate-specific deltas relative to the control mean within each replicate.",
          recurrenceYears: 0
        });
      });

    model.treatments = newTreatments;
    initTreatmentDeltas();
  }

  // =========================
  // 8) CBA ENGINE
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
      if (nLo * nMid <= 0) { hi = mid; nHi = nMid; }
      else { lo = mid; nLo = nMid; }
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
    for (let t = 0; t < series.length; t++) pv += series[t] / Math.pow(1 + ratePct / 100, t);
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
      if (persistence === 0 || y > persistence) benefitByYear[y] = 0;
      else benefitByYear[y] = yieldDelta * price * area * adoption * (1 - risk);
    }

    const costByYear = new Array(years + 1).fill(0);
    const perHaApplicationCost =
      (Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0) + (Number(t.labourCost) || 0);

    const cap0 = Number(t.capitalCost) || 0;
    costByYear[0] += cap0;

    if (!t.isControl) {
      costByYear[0] += perHaApplicationCost * area;
      if (recurrence > 0) {
        for (let y = recurrence; y <= years; y += recurrence) costByYear[y] += perHaApplicationCost * area;
      }
    }

    const cf = new Array(years + 1).fill(0).map((_, i) => benefitByYear[i] - costByYear[i]);

    const pvBenefits = presentValue(benefitByYear, disc);
    const pvCosts = presentValue(costByYear, disc);
    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts > 0 ? pvBenefits / pvCosts : NaN;
    const roi = pvCosts > 0 ? (npv / pvCosts) * 100 : NaN;

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
      irrPct: irr(cf),
      mirrPct: mirr(cf, model.time.mirrFinance, model.time.mirrReinvest),
      paybackYears: payback(cf, disc),
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
      .sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity))
      .map((r, i) => ({ ...r, rankByNpv: i + 1 }));

    const out = results.map(r => {
      if (r.isControl) return { ...r, rankByNpv: null };
      const rr = ranked.find(x => x.treatmentId === r.treatmentId);
      return rr ? rr : { ...r, rankByNpv: null };
    });

    state.results.perTreatmentBaseCase = out;
    state.results.lastComputedAt = new Date().toISOString();
    return out;
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
    showToast("Sensitivity computed", "Grid is ready. Use export for full detail.", "good");
    return grid;
  }

  // =========================
  // 9) RESULTS RENDERING (tables + narrative)
  // =========================
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
    if (mode === "improve_only") return list.filter(r => Number.isFinite(r.npv) && r.npv > 0);
    return list;
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

  function formatDeltaValue(key, control, r) {
    if (!control) return "";
    if (key === "pvBenefits") return money(r.pvBenefits - control.pvBenefits);
    if (key === "pvCosts") return money(r.pvCosts - control.pvCosts);
    if (key === "npv") return money(r.npv - control.npv);
    if (key === "bcr") {
      if (!Number.isFinite(r.bcr) || !Number.isFinite(control.bcr)) return "n/a";
      return fmt(r.bcr - control.bcr);
    }
    if (key === "roiPct") {
      if (!Number.isFinite(r.roiPct) || !Number.isFinite(control.roiPct)) return "n/a";
      return percent(r.roiPct - control.roiPct);
    }
    if (key === "rankByNpv") return r.rankByNpv != null ? String(r.rankByNpv) : "";
    if (key === "deltaNpv") return money(r.npv - control.npv);
    if (key === "deltaPvCost") return money(r.pvCosts - control.pvCosts);
    return "";
  }

  function classifyIndicatorCell(key, control, r) {
    const v = (key === "pvBenefits") ? (r.pvBenefits - (control ? control.pvBenefits : 0))
            : (key === "pvCosts") ? (r.pvCosts - (control ? control.pvCosts : 0))
            : (key === "npv") ? (r.npv - (control ? control.npv : 0))
            : (key === "deltaNpv") ? (r.npv - (control ? control.npv : 0))
            : (key === "deltaPvCost") ? (r.pvCosts - (control ? control.pvCosts : 0))
            : NaN;

    if (key === "pvCosts" || key === "deltaPvCost") {
      if (!Number.isFinite(v)) return "";
      return v < 0 ? "pos" : v > 0 ? "neg" : "zero"; // lower costs are better
    }
    if (key === "pvBenefits") {
      if (!Number.isFinite(v)) return "";
      return v > 0 ? "pos" : v < 0 ? "neg" : "zero";
    }
    if (key === "npv" || key === "deltaNpv") {
      if (!Number.isFinite(v)) return "";
      return v > 0 ? "pos" : v < 0 ? "neg" : "zero";
    }
    if (key === "bcr") {
      if (!Number.isFinite(r.bcr)) return "";
      return r.bcr >= 1 ? "pos" : "neg";
    }
    return "";
  }

  function renderLeaderboard(perTreatment, filterMode) {
    const root = document.getElementById("resultsLeaderboard");
    if (!root) return;

    const list = filterTreatments(perTreatment, filterMode).slice().sort((a, b) => (b.npv || -Infinity) - (a.npv || -Infinity));
    root.innerHTML = "";

    if (!list.length) {
      root.innerHTML = `<p class="small muted">No treatments to rank yet. Import and commit a dataset, or check that control detection worked.</p>`;
      return;
    }

    const table = document.createElement("table");
    table.className = "summary-table";
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
        ${list.map((r, i) => {
          const cls = classifyDelta(r.npv);
          const bcrText = Number.isFinite(r.bcr) ? fmt(r.bcr) : "n/a";
          return `
            <tr>
              <td>${i + 1}</td>
              <td>${esc(r.treatmentName)}</td>
              <td class="${cls}">${money(r.npv)}</td>
              <td>${bcrText}</td>
              <td>${money(r.pvBenefits)}</td>
              <td>${money(r.pvCosts)}</td>
            </tr>
          `;
        }).join("")}
      </tbody>
    `;
    const wrap = document.createElement("div");
    wrap.className = "table-wrap";
    wrap.appendChild(table);
    root.appendChild(wrap);
  }

  function renderComparisonToControl(perTreatment, filterMode) {
    const root = document.getElementById("comparisonToControl");
    if (!root) return;

    const control = perTreatment.find(r => r.isControl) || null;
    const treatments = filterTreatments(perTreatment, filterMode).slice().sort((a, b) => (b.npv || -Infinity) - (a.npv || -Infinity));

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

    root.innerHTML = "";

    if (!control || !treatments.length) {
      root.innerHTML = `<p class="small muted">Results will appear here once a dataset is committed and the control baseline is detected.</p>`;
      return;
    }

    const table = document.createElement("table");
    table.className = "comparison-table";

    const thead = document.createElement("thead");
    thead.innerHTML = `
      <tr>
        <th class="sticky-col">Indicator</th>
        <th class="control-col">${esc(control.treatmentName)} (baseline)</th>
        ${treatments.map(t => `<th>${esc(t.treatmentName)}</th><th>Δ vs control</th>`).join("")}
      </tr>
    `;

    const tbody = document.createElement("tbody");
    indicators.forEach(ind => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td class="sticky-col">${esc(ind.label)}</td>
        <td class="control-col">${esc(formatIndicatorValue(ind.key, control, true))}</td>
        ${treatments.map(t => {
          const val = formatIndicatorValue(ind.key, t, false);
          const del = formatDeltaValue(ind.key, control, t);
          const clsV = classifyIndicatorCell(ind.key, control, t);
          const clsD = classifyIndicatorCell(ind.key, control, t);
          return `<td class="${clsV}">${esc(val)}</td><td class="${clsD}">${esc(del)}</td>`;
        }).join("")}
      `;
      tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);

    const wrap = document.createElement("div");
    wrap.className = "comparison-wrap";
    wrap.appendChild(table);
    root.appendChild(wrap);
  }

  function renderResultsNarrative(perTreatment, filterMode) {
    const root = document.getElementById("resultsNarrative");
    if (!root) return;

    const treatments = filterTreatments(perTreatment, filterMode).slice().sort((a, b) => (b.npv || -Infinity) - (a.npv || -Infinity));
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
      `These results compare each treatment against the control baseline using discounted cashflows over ${years} years. The grain price in the base case is ${money(price)} per tonne and the discount rate is ${fmt(disc)} percent per year. Yield effects are assumed to persist for ${persistence} years after application. Adoption is applied as a multiplier of ${fmt(adopt)} and risk reduces benefits by ${fmt(risk)} as a proportion.`
    );

    if (top) {
      parts.push(
        `${top.treatmentName} is the strongest result by net present value under the current assumptions. Its present value of benefits is ${money(top.pvBenefits)}, its present value of costs is ${money(top.pvCosts)}, and its net present value is ${money(top.npv)}. This reflects the yield uplift compared with the control, valued at grain price, relative to the incremental costs (and whether those costs repeat under the recurrence setting).`
      );
    }

    if (worst && top && worst.treatmentId !== top.treatmentId) {
      parts.push(
        `${worst.treatmentName} is a weaker result under the current assumptions. Its present value of benefits is ${money(worst.pvBenefits)}, its present value of costs is ${money(worst.pvCosts)}, and its net present value is ${money(worst.npv)}. This usually happens when the yield uplift is small, incremental costs are high, or the yield effect does not persist long enough to pay back the costs in present value terms.`
      );
    }

    const improves = treatments.filter(r => Number.isFinite(r.npv) && r.npv > 0).length;
    const total = treatments.length;
    if (total) {
      parts.push(
        `Under the current assumptions, ${improves} of ${total} treatments show a positive net present value relative to the control. This does not decide a choice. It highlights which treatments are more dependent on grain price, cost control, and how long the yield lift lasts.`
      );
    } else {
      parts.push(`No non-control treatments are available to compare under the current filter. Try “Show all”.`);
    }

    root.textContent = parts.join("\n\n");
  }

  function renderBaseCaseBadge() {
    const el = document.getElementById("baseCaseBadge");
    if (!el) return;
    const years = Math.floor(model.time.years || 0);
    const price = getGrainPrice();
    const disc = Number(model.time.discBase) || 0;
    const persistence = getPersistenceYears();
    const adopt = clamp(Number(model.adoption.base) || 1, 0, 1);
    const risk = clamp(Number(model.risk.base) || 0, 0, 1);
    el.textContent = `Horizon ${years} years. Price ${money(price)} per tonne. Discount ${fmt(disc)} percent. Persistence ${persistence} years. Adoption ${fmt(adopt)}. Risk ${fmt(risk)}.`;
  }

  // =========================
  // 10) DATA CHECKS + IMPORT SUMMARY RENDERING
  // =========================
  function renderImportStatus(text) {
    const el = document.getElementById("importStatus");
    if (el) el.textContent = text || "";
  }

  function renderImportSummary(derived, schema, sourceName, isStaged) {
    const box = document.getElementById("importSummary");
    if (box) {
      const errs = derived.checks.filter(c => c.severity === "error").length;
      const warns = derived.checks.filter(c => c.severity === "warn").length;
      const nRows = derived.plotDeltas ? derived.plotDeltas.length : 0;
      const nTreat = derived.treatmentSummary ? derived.treatmentSummary.filter(s => !s.isControl).length : 0;
      box.textContent =
        `${isStaged ? "Staged" : "Committed"}: ${sourceName || "dataset"} · ` +
        `${nRows} rows · ${nTreat} treatments · ${errs} errors · ${warns} warnings.`;
    }

    const hdr = document.getElementById("headerDatasetName");
    if (hdr) hdr.textContent = (state.dataset.sourceName || (isStaged ? sourceName : "None loaded")) || "None loaded";

    const status = document.getElementById("importStatus");
    if (status) {
      const errs = derived.checks.filter(c => c.severity === "error").length;
      const warns = derived.checks.filter(c => c.severity === "warn").length;
      const schemaBits = [];
      if (schema) {
        if (schema.treatmentCol) schemaBits.push(`treatment: ${schema.treatmentCol}`);
        if (schema.replicateCol) schemaBits.push(`replicate: ${schema.replicateCol}`);
        if (schema.yieldCol) schemaBits.push(`yield: ${schema.yieldCol}`);
      }
      status.textContent = `${isStaged ? "Staged" : "Committed"} dataset · ${errs} errors · ${warns} warnings · ${schemaBits.join(" · ")}`;
    }
  }

  function renderDataChecks(derived) {
    const root = document.getElementById("dataChecks");
    if (!root) return;

    const checks = derived && derived.checks ? derived.checks : [];
    if (!checks.length) {
      root.innerHTML = `<p class="small muted">No data checks triggered.</p>`;
      return;
    }

    const rows = checks.slice().sort((a, b) => {
      const sevRank = s => (s === "error" ? 0 : s === "warn" ? 1 : 2);
      return sevRank(a.severity) - sevRank(b.severity);
    });

    root.innerHTML = `
      <div class="table-wrap">
        <table class="summary-table">
          <thead>
            <tr>
              <th>Severity</th>
              <th>Check</th>
              <th>Count</th>
              <th>Summary</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(r => {
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
            }).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function setHeaderDataset(name) {
    const el = document.getElementById("headerDatasetName");
    if (el) el.textContent = name || "None loaded";
  }

  // =========================
  // 11) EXPORTS
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
        r.map(x => {
          const s = x == null ? "" : String(x);
          const needs = s.includes(",") || s.includes('"') || s.includes("\n") || s.includes("\r");
          const safe = s.replace(/"/g, '""');
          return needs ? `"${safe}"` : safe;
        }).join(",")
      )
      .join("\r\n");
  }

  function exportCleanedDatasetTsv() {
    const derived = state.dataset.derived;
    if (!derived || !derived.plotDeltas || !derived.plotDeltas.length) {
      showToast("No cleaned dataset", "Import and commit a dataset first.", "warn");
      return;
    }
    const schema = state.dataset.schema;
    const rows = derived.plotDeltas;
    const costCols = schema && schema.costCols ? schema.costCols.slice() : [];
    const header = [
      "treatment","treatment_key","replicate","plot","is_control","yield","control_yield_mean_replicate","delta_yield",
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
    showToast("Exported", "Cleaned dataset TSV downloaded.", "good");
  }

  function exportTreatmentSummaryCsv() {
    const derived = state.dataset.derived;
    if (!derived || !derived.treatmentSummary || !derived.treatmentSummary.length) {
      showToast("No summary", "Import and commit a dataset first.", "warn");
      return;
    }
    const schema = state.dataset.schema;
    const costCols = schema && schema.costCols ? schema.costCols.slice() : [];
    const rows = [];
    rows.push([
      "treatment","is_control","n_yield","yield_mean","yield_sd","delta_yield_mean","delta_yield_sd","delta_yield_median",
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
    showToast("Exported", "Treatment summary CSV downloaded.", "good");
  }

  function exportSensitivityGridCsv() {
    const grid = state.results.sensitivityGrid || [];
    if (!grid.length) {
      showToast("No sensitivity grid", "Run sensitivity first.", "warn");
      return;
    }
    const rows = [];
    rows.push([
      "treatment","price_per_tonne","discount_rate_pct","persistence_years","recurrence_years",
      "pv_benefits","pv_costs","npv","bcr","roi_pct"
    ]);
    grid.forEach(g => {
      rows.push([
        g.treatment, g.pricePerTonne, g.discountRatePct, g.persistenceYears, g.recurrenceYears,
        g.pvBenefits, g.pvCosts, g.npv, g.bcr, g.roiPct
      ]);
    });
    const csv = toCsv(rows);
    const name = slug(model.project.name || "project");
    downloadFile(`${name}_sensitivity_grid.csv`, csv, "text/csv");
    showToast("Exported", "Sensitivity grid CSV downloaded.", "good");
  }

  function exportWorkbookIfAvailable() {
    if (typeof XLSX === "undefined") {
      showToast("Excel export unavailable", "XLSX library did not load.", "warn");
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
        costCols.forEach(c => { obj[`cost_per_ha:${c}`] = Number.isFinite(r.costsPerHa[c]) ? r.costsPerHa[c] : null; });
        costCols.forEach(c => { obj[`delta_cost_per_ha:${c}`] = Number.isFinite(r.deltaCostsPerHa[c]) ? r.deltaCostsPerHa[c] : null; });
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
    if (grid.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(grid), "SensitivityGrid");

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile(`${name}_workbook.xlsx`, wbout, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Exported", "Excel workbook downloaded.", "good");
  }

  // =========================
  // 12) OUTPUTS + TREATMENTS EDIT UI
  // =========================
  function renderOutputs() {
    const root = document.getElementById("outputsList");
    if (!root) return;
    const yieldOut = ensureYieldOutput();

    root.innerHTML = `
      <div class="table-wrap">
        <table class="summary-table">
          <thead>
            <tr>
              <th>Output</th>
              <th>Unit</th>
              <th>Value used for pricing</th>
              <th>Notes</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>${esc(yieldOut.name)}</td>
              <td>${esc(yieldOut.unit)}</td>
              <td style="min-width:220px">
                <input id="outputYieldPrice" type="number" step="1" value="${esc(yieldOut.value)}" />
              </td>
              <td class="small muted">Value is interpreted as dollars per tonne when yield deltas are t per hectare.</td>
            </tr>
          </tbody>
        </table>
      </div>
    `;

    const inp = document.getElementById("outputYieldPrice");
    if (inp) {
      inp.addEventListener("input", () => {
        const v = parseNumber(inp.value);
        if (Number.isFinite(v)) {
          yieldOut.value = v;
          const gp = document.getElementById("grainPrice");
          if (gp && !String(gp.value || "").trim()) gp.value = String(v);
          recomputeAndRenderAll("Updated grain price");
        }
      });
    }
  }

  function renderTreatments() {
    const root = document.getElementById("treatmentsList");
    if (!root) return;

    const yieldOut = ensureYieldOutput();
    root.innerHTML = "";

    const list = document.createElement("div");
    list.className = "list";

    model.treatments.forEach(t => {
      const isCtrl = !!t.isControl;
      const card = document.createElement("div");
      card.className = "list-item";

      const dy = Number(t.deltas?.[yieldOut.id] || 0);
      const cost = (Number(t.labourCost) || 0) + (Number(t.servicesCost) || 0) + (Number(t.materialsCost) || 0);

      card.innerHTML = `
        <div class="title">
          <strong>${esc(t.name)}</strong>
          <span class="badge">${isCtrl ? "Control" : "Treatment"}</span>
        </div>

        <div class="kv">
          <div class="k">Area (ha)</div>
          <div class="v"><input type="number" step="1" data-treat-field="area" data-treat-id="${esc(t.id)}" value="${esc(t.area)}" ${isCtrl ? "disabled" : ""} /></div>

          <div class="k">Yield delta (t/ha)</div>
          <div class="v"><input type="number" step="0.01" data-treat-field="yieldDelta" data-treat-id="${esc(t.id)}" value="${esc(dy)}" ${isCtrl ? "disabled" : ""} /></div>

          <div class="k">Labour cost ($/ha)</div>
          <div class="v"><input type="number" step="0.01" data-treat-field="labourCost" data-treat-id="${esc(t.id)}" value="${esc(t.labourCost)}" ${isCtrl ? "disabled" : ""} /></div>

          <div class="k">Services cost ($/ha)</div>
          <div class="v"><input type="number" step="0.01" data-treat-field="servicesCost" data-treat-id="${esc(t.id)}" value="${esc(t.servicesCost)}" ${isCtrl ? "disabled" : ""} /></div>

          <div class="k">Materials cost ($/ha)</div>
          <div class="v"><input type="number" step="0.01" data-treat-field="materialsCost" data-treat-id="${esc(t.id)}" value="${esc(t.materialsCost)}" ${isCtrl ? "disabled" : ""} /></div>

          <div class="k">Capital cost (year 0, $)</div>
          <div class="v"><input type="number" step="1" data-treat-field="capitalCost" data-treat-id="${esc(t.id)}" value="${esc(t.capitalCost)}" ${isCtrl ? "disabled" : ""} /></div>

          <div class="k">Recurrence (years)</div>
          <div class="v"><input type="number" step="1" min="0" data-treat-field="recurrenceYears" data-treat-id="${esc(t.id)}" value="${esc(t.recurrenceYears)}" ${isCtrl ? "disabled" : ""} /></div>

          <div class="k">Incremental cost ($/ha per application)</div>
          <div class="v"><span class="small muted">${isCtrl ? "Baseline" : money(cost)}</span></div>
        </div>

        <div class="divider"></div>
        <div class="small muted">${esc(t.notes || "")}</div>
      `;

      list.appendChild(card);
    });

    root.appendChild(list);

    $$("#treatmentsList input[data-treat-field]").forEach(inp => {
      inp.addEventListener("input", () => {
        const id = inp.getAttribute("data-treat-id");
        const field = inp.getAttribute("data-treat-field");
        const t = model.treatments.find(x => x.id === id);
        if (!t || t.isControl) return;

        const v = parseNumber(inp.value);
        if (field === "area") t.area = Number.isFinite(v) ? Math.max(0, v) : t.area;
        else if (field === "labourCost") t.labourCost = Number.isFinite(v) ? v : t.labourCost;
        else if (field === "servicesCost") t.servicesCost = Number.isFinite(v) ? v : t.servicesCost;
        else if (field === "materialsCost") t.materialsCost = Number.isFinite(v) ? v : t.materialsCost;
        else if (field === "capitalCost") t.capitalCost = Number.isFinite(v) ? v : t.capitalCost;
        else if (field === "recurrenceYears") t.recurrenceYears = Number.isFinite(v) ? Math.max(0, Math.floor(v)) : t.recurrenceYears;
        else if (field === "yieldDelta") {
          const out = ensureYieldOutput();
          if (!t.deltas) t.deltas = {};
          t.deltas[out.id] = Number.isFinite(v) ? v : t.deltas[out.id];
        }

        recomputeAndRenderAll("Updated treatment");
      });
    });
  }

  // =========================
  // 13) AI BRIEFING + RESULTS JSON
  // =========================
  function buildResultsPayload() {
    const base = state.results.perTreatmentBaseCase || [];
    const treatments = base.filter(r => !r.isControl).slice().sort((a, b) => (b.npv || -Infinity) - (a.npv || -Infinity));
    const years = Math.floor(model.time.years || 0);
    const price = getGrainPrice();
    const disc = Number(model.time.discBase) || 0;
    const persistence = getPersistenceYears();
    const adopt = clamp(Number(model.adoption.base) || 1, 0, 1);
    const risk = clamp(Number(model.risk.base) || 0, 0, 1);

    const derived = state.dataset.derived;
    const checks = derived && derived.checks ? derived.checks : [];

    return {
      meta: {
        generatedAt: new Date().toISOString(),
        datasetSource: state.dataset.sourceName || "",
        committedAt: state.dataset.committedAt || ""
      },
      settings: {
        years,
        grainPricePerTonne: price,
        discountRatePct: disc,
        persistenceYears: persistence,
        adoptionMultiplier: adopt,
        riskMultiplier: risk
      },
      treatments: treatments.map(t => ({
        name: t.treatmentName,
        recurrenceYears: t.recurrenceYears,
        pvBenefits: t.pvBenefits,
        pvCosts: t.pvCosts,
        npv: t.npv,
        bcr: t.bcr,
        roiPct: t.roiPct,
        irrPct: t.irrPct,
        mirrPct: t.mirrPct,
        paybackYears: t.paybackYears
      })),
      dataChecks: checks
    };
  }

  function buildAiBriefingText() {
    const payload = buildResultsPayload();
    const t = payload.treatments || [];
    const top = t[0] || null;

    const p = [];
    p.push("Write a decision brief in clear plain language for an on farm manager. Use full sentences and paragraphs only. Do not use bullet points. Do not use an em dash. Do not use abbreviations.");
    p.push(
      `The analysis compares each soil amendment treatment against the control baseline using discounted cashflows over ${payload.settings.years} years. The grain price used in the base case is ${money(payload.settings.grainPricePerTonne)} per tonne and the discount rate is ${fmt(payload.settings.discountRatePct)} percent per year. Yield effects are assumed to persist for ${payload.settings.persistenceYears} years after application. The adoption multiplier is ${fmt(payload.settings.adoptionMultiplier)} and the risk multiplier reduces benefits by ${fmt(payload.settings.riskMultiplier)} as a proportion.`
    );

    if (payload.dataChecks && payload.dataChecks.length) {
      const errs = payload.dataChecks.filter(c => c.severity === "error").length;
      const warns = payload.dataChecks.filter(c => c.severity === "warn").length;
      p.push(`Data checks were run after import. There are ${errs} error level checks and ${warns} warning level checks. Summarise implications for interpretation and sensitivity without recommending a single option.`);
    } else {
      p.push("Data checks were run after import and no triggers were recorded in the checks panel.");
    }

    if (top) {
      p.push(
        `In the base case, the strongest treatment by net present value is ${top.name}. Its present value of benefits is ${money(top.pvBenefits)}, its present value of costs is ${money(top.pvCosts)}, and its net present value is ${money(top.npv)}. Explain what is driving this result in practical terms, focusing on yield uplift against the control and incremental costs under the recurrence assumption.`
      );
    }

    const pos = t.filter(x => Number.isFinite(x.npv) && x.npv > 0).length;
    if (t.length) p.push(`Across all treatments, ${pos} have a positive net present value relative to the control under the base case. Explain what this means in terms of trade offs. Do not instruct the user to choose anything.`);

    p.push("Explain the meaning of net present value, present value of benefits, present value of costs, benefit cost ratio, and return on investment in farmer facing terms.");
    p.push("Explain sensitivity focusing on grain price, discount rate, persistence of yield effects, and recurrence of costs. Explain which inputs are likely to change and how that could move results.");
    p.push("List practical options to improve weaker treatments without giving a directive, such as reducing cost items, changing timing, improving establishment, or verifying the yield response with additional seasons or sites.");
    p.push("Use the following computed results as the only quantitative basis for the brief.");
    p.push(JSON.stringify(payload, null, 2));
    return p.join("\n\n");
  }

  function renderAiTab() {
    const a = document.getElementById("aiBriefingText");
    const j = document.getElementById("resultsJson");
    if (a) a.value = buildAiBriefingText();
    if (j) j.value = JSON.stringify(buildResultsPayload(), null, 2);
  }

  async function copyToClipboard(text) {
    try {
      await navigator.clipboard.writeText(String(text || ""));
      showToast("Copied", "Copied to clipboard.", "good");
    } catch (e) {
      showToast("Copy failed", "Your browser blocked clipboard access. Select and copy manually.", "warn");
    }
  }

  // =========================
  // 14) PLOTS (responsive canvas)
  // =========================
  function getCanvas(id) {
    const c = document.getElementById(id);
    if (!c) return null;
    const rect = c.getBoundingClientRect();
    const dpr = Math.max(1, window.devicePixelRatio || 1);
    const w = Math.max(320, Math.floor(rect.width || c.width || 800));
    const h = Math.max(240, Math.floor((c.height / c.width) * w) || c.height || 360);
    c.width = Math.floor(w * dpr);
    c.height = Math.floor(h * dpr);
    const ctx = c.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    return { c, ctx, w, h };
  }

  function clearPlot(p) {
    p.ctx.clearRect(0, 0, p.w, p.h);
  }

  function drawAxes(p, margin) {
    const { ctx, w, h } = p;
    ctx.save();
    ctx.globalAlpha = 0.55;
    ctx.strokeStyle = "white";
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(margin.l, margin.t);
    ctx.lineTo(margin.l, h - margin.b);
    ctx.lineTo(w - margin.r, h - margin.b);
    ctx.stroke();
    ctx.restore();
  }

  function drawText(ctx, text, x, y, align, size, alpha) {
    ctx.save();
    ctx.fillStyle = "white";
    ctx.globalAlpha = alpha == null ? 0.9 : alpha;
    ctx.font = `${size || 12}px system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial`;
    ctx.textAlign = align || "left";
    ctx.textBaseline = "middle";
    ctx.fillText(text, x, y);
    ctx.restore();
  }

  function plotNpvBar() {
    const p = getCanvas("chartNpv");
    if (!p) return;

    const base = state.results.perTreatmentBaseCase || [];
    const list = base.filter(r => !r.isControl).slice().sort((a, b) => (b.npv || -Infinity) - (a.npv || -Infinity)).slice(0, 12);

    clearPlot(p);

    const { ctx, w, h } = p;
    const margin = { l: 56, r: 18, t: 22, b: 36 };
    drawAxes(p, margin);

    if (!list.length) {
      drawText(ctx, "No treatments to plot yet.", margin.l + 10, h / 2, "left", 14, 0.9);
      return;
    }

    const values = list.map(d => d.npv).filter(Number.isFinite);
    const minV = Math.min(0, ...values);
    const maxV = Math.max(0, ...values);
    const span = (maxV - minV) || 1;

    const plotW = w - margin.l - margin.r;
    const plotH = h - margin.t - margin.b;
    const barH = plotH / list.length;

    // zero line
    const x0 = margin.l + ((0 - minV) / span) * plotW;
    ctx.save();
    ctx.strokeStyle = "white";
    ctx.globalAlpha = 0.35;
    ctx.beginPath();
    ctx.moveTo(x0, margin.t);
    ctx.lineTo(x0, margin.t + plotH);
    ctx.stroke();
    ctx.restore();

    list.forEach((d, i) => {
      const y = margin.t + i * barH + barH * 0.15;
      const bh = barH * 0.7;
      const v = Number.isFinite(d.npv) ? d.npv : 0;
      const x = margin.l + ((Math.min(v, 0) - minV) / span) * plotW;
      const bw = (Math.abs(v) / span) * plotW;

      ctx.save();
      ctx.globalAlpha = 0.85;
      ctx.fillStyle = v >= 0 ? "rgba(46, 204, 113, 0.95)" : "rgba(231, 76, 60, 0.95)";
      ctx.fillRect(x, y, bw, bh);
      ctx.restore();

      drawText(ctx, d.treatmentName.length > 22 ? d.treatmentName.slice(0, 22) + "…" : d.treatmentName, 8, y + bh / 2, "left", 11, 0.85);
      drawText(ctx, money(v), margin.l + ((v - minV) / span) * plotW + (v >= 0 ? 6 : -6), y + bh / 2, v >= 0 ? "left" : "right", 11, 0.9);
    });

    drawText(ctx, "Top treatments by NPV (base case)", margin.l, 12, "left", 12, 0.85);
  }

  function plotCashflowLine() {
    const p = getCanvas("chartCashflow");
    if (!p) return;

    const sel = document.getElementById("cashflowTreatmentSelect");
    const chosenId = sel ? sel.value : "";
    const base = state.results.perTreatmentBaseCase || [];
    const chosen = base.find(r => r.treatmentId === chosenId) || base.find(r => !r.isControl) || null;

    clearPlot(p);
    const { ctx, w, h } = p;
    const margin = { l: 56, r: 18, t: 22, b: 36 };
    drawAxes(p, margin);

    if (!chosen) {
      drawText(ctx, "No treatment available to plot cashflows.", margin.l + 10, h / 2, "left", 14, 0.9);
      return;
    }

    const series = chosen.cf || [];
    const years = series.length ? series.length - 1 : 0;
    if (!years) {
      drawText(ctx, "Cashflow series is empty.", margin.l + 10, h / 2, "left", 14, 0.9);
      return;
    }

    const plotW = w - margin.l - margin.r;
    const plotH = h - margin.t - margin.b;

    const minV = Math.min(0, ...series.filter(Number.isFinite));
    const maxV = Math.max(0, ...series.filter(Number.isFinite));
    const span = (maxV - minV) || 1;

    const xAt = t => margin.l + (t / years) * plotW;
    const yAt = v => margin.t + (1 - ((v - minV) / span)) * plotH;

    // zero line
    const y0 = yAt(0);
    ctx.save();
    ctx.strokeStyle = "white";
    ctx.globalAlpha = 0.35;
    ctx.beginPath();
    ctx.moveTo(margin.l, y0);
    ctx.lineTo(margin.l + plotW, y0);
    ctx.stroke();
    ctx.restore();

    // line
    ctx.save();
    ctx.strokeStyle = "white";
    ctx.globalAlpha = 0.9;
    ctx.lineWidth = 2;
    ctx.beginPath();
    series.forEach((v, t) => {
      const x = xAt(t);
      const y = yAt(Number.isFinite(v) ? v : 0);
      if (t === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });
    ctx.stroke();
    ctx.restore();

    // points
    ctx.save();
    ctx.fillStyle = "white";
    ctx.globalAlpha = 0.9;
    series.forEach((v, t) => {
      const x = xAt(t);
      const y = yAt(Number.isFinite(v) ? v : 0);
      ctx.beginPath();
      ctx.arc(x, y, 3, 0, Math.PI * 2);
      ctx.fill();
    });
    ctx.restore();

    drawText(ctx, `Cashflow vs control: ${chosen.treatmentName}`, margin.l, 12, "left", 12, 0.85);
    drawText(ctx, `IRR ${Number.isFinite(chosen.irrPct) ? percent(chosen.irrPct) : "n/a"} · Payback ${chosen.paybackYears == null ? "n/a" : chosen.paybackYears + " years"}`, margin.l, h - 12, "left", 11, 0.75);
  }

  function renderCashflowSelect() {
    const sel = document.getElementById("cashflowTreatmentSelect");
    if (!sel) return;
    const base = state.results.perTreatmentBaseCase || [];
    const opts = base.filter(r => !r.isControl).slice().sort((a, b) => (b.npv || -Infinity) - (a.npv || -Infinity));
    sel.innerHTML = opts.map(o => `<option value="${esc(o.treatmentId)}">${esc(o.treatmentName)}</option>`).join("");
    sel.onchange = () => plotCashflowLine();
  }

  // =========================
  // 15) CONFIG + SCENARIOS
  // =========================
  function snapshotScenario() {
    return {
      id: uid(),
      name: "",
      savedAt: new Date().toISOString(),
      model: JSON.parse(JSON.stringify(model)),
      config: JSON.parse(JSON.stringify(state.config)),
      ui: { resultsFilter: state.ui.resultsFilter }
    };
  }

  function applyScenario(s) {
    if (!s || !s.model) return;
    // shallow replace into existing objects to keep references stable
    Object.assign(model.project, s.model.project || {});
    Object.assign(model.time, s.model.time || {});
    model.outputs = Array.isArray(s.model.outputs) ? s.model.outputs : model.outputs;
    model.treatments = Array.isArray(s.model.treatments) ? s.model.treatments : model.treatments;
    model.adoption = s.model.adoption || model.adoption;
    model.risk = s.model.risk || model.risk;

    state.config = s.config ? s.config : state.config;
    if (s.ui && s.ui.resultsFilter) state.ui.resultsFilter = s.ui.resultsFilter;

    initTreatmentDeltas();
    syncConfigInputsFromModel();
    renderOutputs();
    renderTreatments();
    recomputeAndRenderAll("Loaded scenario");
  }

  function loadScenariosFromStorage() {
    try {
      const raw = localStorage.getItem(STORAGE_KEYS.scenarios);
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr : [];
    } catch (e) {
      return [];
    }
  }

  function saveScenariosToStorage(arr) {
    try {
      localStorage.setItem(STORAGE_KEYS.scenarios, JSON.stringify(arr || []));
    } catch (e) {}
  }

  function getActiveScenarioId() {
    try {
      return localStorage.getItem(STORAGE_KEYS.activeScenario) || "";
    } catch (e) {
      return "";
    }
  }

  function setActiveScenarioId(id) {
    try {
      localStorage.setItem(STORAGE_KEYS.activeScenario, id || "");
    } catch (e) {}
  }

  function renderScenarioDropdown() {
    const sel = document.getElementById("scenarioSelect");
    if (!sel) return;
    const scenarios = loadScenariosFromStorage();
    const active = getActiveScenarioId();
    sel.innerHTML = `<option value="">Select a saved scenario…</option>` + scenarios.map(s => {
      const tag = (s.id === active) ? " (active)" : "";
      return `<option value="${esc(s.id)}">${esc((s.name || "Scenario") + tag)}</option>`;
    }).join("");
  }

  function handleSaveScenario() {
    const nameInp = document.getElementById("scenarioName");
    const name = nameInp ? String(nameInp.value || "").trim() : "";
    if (!name) {
      showToast("Scenario name required", "Enter a name, then save.", "warn");
      return;
    }
    const scenarios = loadScenariosFromStorage();
    const s = snapshotScenario();
    s.name = name;

    scenarios.unshift(s);
    saveScenariosToStorage(scenarios);
    setActiveScenarioId(s.id);
    renderScenarioDropdown();
    showToast("Scenario saved", "Stored in this browser.", "good");
  }

  function handleLoadScenario() {
    const sel = document.getElementById("scenarioSelect");
    const id = sel ? sel.value : "";
    if (!id) {
      showToast("Choose a scenario", "Select a saved scenario first.", "warn");
      return;
    }
    const scenarios = loadScenariosFromStorage();
    const s = scenarios.find(x => x.id === id);
    if (!s) {
      showToast("Scenario not found", "It may have been cleared from browser storage.", "bad");
      return;
    }
    setActiveScenarioId(s.id);
    renderScenarioDropdown();
    applyScenario(s);
    showToast("Scenario loaded", "All settings and treatments updated.", "good");
  }

  function handleDeleteScenario() {
    const sel = document.getElementById("scenarioSelect");
    const id = sel ? sel.value : "";
    if (!id) {
      showToast("Choose a scenario", "Select a saved scenario first.", "warn");
      return;
    }
    const scenarios = loadScenariosFromStorage().filter(s => s.id !== id);
    saveScenariosToStorage(scenarios);
    if (getActiveScenarioId() === id) setActiveScenarioId("");
    renderScenarioDropdown();
    showToast("Scenario deleted", "Removed from this browser.", "good");
  }

  // =========================
  // 16) TECHNICAL APPENDIX
  // =========================
  function buildAppendixPayload() {
    const derived = state.dataset.derived || null;
    const schema = state.dataset.schema || null;
    const base = state.results.perTreatmentBaseCase || [];
    const grid = state.results.sensitivityGrid || [];
    return {
      meta: {
        generatedAt: new Date().toISOString(),
        datasetSource: state.dataset.sourceName || "",
        committedAt: state.dataset.committedAt || ""
      },
      project: model.project,
      time: model.time,
      settings: {
        grainPricePerTonne: getGrainPrice(),
        persistenceYears: getPersistenceYears(),
        adoptionMultiplier: model.adoption.base,
        riskMultiplier: model.risk.base
      },
      schema,
      dataChecks: derived && derived.checks ? derived.checks : [],
      treatmentSummary: derived && derived.treatmentSummary ? derived.treatmentSummary : [],
      baseCaseResults: base,
      sensitivityGrid: grid
    };
  }

  function openTechnicalAppendix() {
    try {
      localStorage.setItem(STORAGE_KEYS.appendixPayload, JSON.stringify(buildAppendixPayload()));
    } catch (e) {}
    window.open("technical-appendix.html", "_blank", "noopener,noreferrer");
    showToast("Technical appendix", "Opened in a new tab.", "good");
  }

  // =========================
  // 17) TAB NAVIGATION
  // =========================
  function setActiveTab(tabId) {
    state.ui.activeTab = tabId;
    const tabs = $$(".tab-btn");
    const panels = $$(".tab-panel");
    tabs.forEach(b => b.classList.toggle("active", b.getAttribute("data-tab") === tabId));
    panels.forEach(p => p.classList.toggle("active", p.id === `tab-${tabId}`));

    // refresh plots when entering results
    if (tabId === "results") {
      requestAnimationFrame(() => {
        plotNpvBar();
        renderCashflowSelect();
        plotCashflowLine();
      });
    }
    if (tabId === "ai") renderAiTab();
  }

  // =========================
  // 18) CONFIG INPUT SYNC
  // =========================
  function syncConfigInputsFromModel() {
    const setVal = (id, v) => {
      const el = document.getElementById(id);
      if (el) el.value = (v == null ? "" : String(v));
    };

    setVal("yearsHorizon", model.time.years);
    setVal("discountRateBase", model.time.discBase);
    setVal("adoptionMultiplier", model.adoption.base);
    setVal("riskMultiplier", model.risk.base);

    const gp = document.getElementById("grainPrice");
    if (gp && !String(gp.value || "").trim()) gp.value = String(getGrainPrice());

    setVal("persistenceYears", state.config.persistenceYears);

    // sensitivity lists
    setVal("sensPrice", (state.config.sensPrice || []).join(", "));
    setVal("sensDiscount", (state.config.sensDiscount || []).join(", "));
    setVal("sensPersistence", (state.config.sensPersistence || []).join(", "));
    setVal("sensRecurrence", (state.config.sensRecurrence || []).join(", "));
  }

  function parseListNumbers(text) {
    const s = String(text || "");
    const parts = s.split(/[,;\s]+/g).map(x => x.trim()).filter(Boolean);
    const nums = parts.map(parseNumber).filter(Number.isFinite);
    return nums;
  }

  function bindConfigInputs() {
    const on = (id, fn) => {
      const el = document.getElementById(id);
      if (el) el.addEventListener("input", fn);
    };

    on("grainPrice", () => recomputeAndRenderAll("Updated grain price"));
    on("yearsHorizon", () => {
      const v = parseNumber($("#yearsHorizon")?.value);
      if (Number.isFinite(v) && v > 0) model.time.years = Math.max(1, Math.floor(v));
      recomputeAndRenderAll("Updated horizon");
    });
    on("discountRateBase", () => {
      const v = parseNumber($("#discountRateBase")?.value);
      if (Number.isFinite(v)) model.time.discBase = v;
      recomputeAndRenderAll("Updated discount rate");
    });
    on("adoptionMultiplier", () => {
      const v = parseNumber($("#adoptionMultiplier")?.value);
      if (Number.isFinite(v)) model.adoption.base = clamp(v, 0, 1);
      recomputeAndRenderAll("Updated adoption multiplier");
    });
    on("riskMultiplier", () => {
      const v = parseNumber($("#riskMultiplier")?.value);
      if (Number.isFinite(v)) model.risk.base = clamp(v, 0, 1);
      recomputeAndRenderAll("Updated risk multiplier");
    });
    on("persistenceYears", () => {
      const v = parseNumber($("#persistenceYears")?.value);
      if (Number.isFinite(v) && v >= 0) state.config.persistenceYears = Math.floor(v);
      recomputeAndRenderAll("Updated persistence");
    });

    const updateSens = () => {
      const p = parseListNumbers($("#sensPrice")?.value);
      const d = parseListNumbers($("#sensDiscount")?.value);
      const pe = parseListNumbers($("#sensPersistence")?.value);
      const r = parseListNumbers($("#sensRecurrence")?.value);
      if (p.length) state.config.sensPrice = p;
      if (d.length) state.config.sensDiscount = d;
      if (pe.length) state.config.sensPersistence = pe.map(x => Math.max(0, Math.floor(x)));
      if (r.length) state.config.sensRecurrence = r.map(x => Math.max(0, Math.floor(x)));
      showToast("Sensitivity settings updated", "Ready to compute sensitivity grid.", "good");
    };

    const btn = document.getElementById("btnUpdateSensitivitySettings");
    if (btn) btn.addEventListener("click", updateSens);
  }

  // =========================
  // 19) SENSITIVITY SUMMARY RENDER
  // =========================
  function renderSensitivitySummary() {
    const root = document.getElementById("sensitivitySummary");
    if (!root) return;
    const grid = state.results.sensitivityGrid || [];
    if (!grid.length) {
      root.innerHTML = `<p class="small muted">Run sensitivity to populate results.</p>`;
      return;
    }

    // simple summary: for each treatment, fraction of grid points with NPV>0
    const byTreat = new Map();
    grid.forEach(g => {
      if (!byTreat.has(g.treatmentId)) byTreat.set(g.treatmentId, { name: g.treatment, n: 0, pos: 0, npv: [] });
      const x = byTreat.get(g.treatmentId);
      x.n += 1;
      if (Number.isFinite(g.npv) && g.npv > 0) x.pos += 1;
      if (Number.isFinite(g.npv)) x.npv.push(g.npv);
    });

    const rows = Array.from(byTreat.values())
      .map(x => ({
        name: x.name,
        sharePositive: x.n ? (x.pos / x.n) * 100 : NaN,
        medianNpv: median(x.npv),
        meanNpv: mean(x.npv)
      }))
      .sort((a, b) => (Number.isFinite(b.sharePositive) ? b.sharePositive : -Infinity) - (Number.isFinite(a.sharePositive) ? a.sharePositive : -Infinity))
      .slice(0, 12);

    root.innerHTML = `
      <div class="table-wrap">
        <table class="summary-table">
          <thead>
            <tr>
              <th>Treatment</th>
              <th>Share of scenarios with NPV &gt; 0</th>
              <th>Median NPV</th>
              <th>Mean NPV</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map(r => `
              <tr>
                <td>${esc(r.name)}</td>
                <td>${Number.isFinite(r.sharePositive) ? percent(r.sharePositive) : "n/a"}</td>
                <td>${money(r.medianNpv)}</td>
                <td>${money(r.meanNpv)}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
      <p class="small muted">This table summarises how often each treatment remains positive across the full sensitivity grid. It is a stability indicator only.</p>
    `;
  }

  // =========================
  // 20) RESULTS FILTER UI
  // =========================
  function bindResultsFilters() {
    const map = {
      btnFilterAll: "all",
      btnFilterTopNpv: "top5_npv",
      btnFilterTopBcr: "top5_bcr",
      btnFilterImprove: "improve_only"
    };
    Object.keys(map).forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      el.addEventListener("click", () => {
        state.ui.resultsFilter = map[id];
        recomputeAndRenderAll("Updated results filter");
      });
    });
  }

  // =========================
  // 21) FULL RENDER PIPELINE
  // =========================
  function recomputeAndRenderAll(reason) {
    computeBaseCaseResultsVsControl();
    renderBaseCaseBadge();
    renderOutputs();
    renderTreatments();

    const base = state.results.perTreatmentBaseCase || [];
    renderLeaderboard(base, state.ui.resultsFilter);
    renderComparisonToControl(base, state.ui.resultsFilter);
    renderResultsNarrative(base, state.ui.resultsFilter);
    renderAiTab();

    renderCashflowSelect();
    plotNpvBar();
    plotCashflowLine();

    const stamp = document.getElementById("lastUpdatedStamp");
    if (stamp) stamp.textContent = state.results.lastComputedAt ? `Updated ${new Date(state.results.lastComputedAt).toLocaleString()}` : "";

    if (reason) showToast("Updated", reason, "good");
  }

  // =========================
  // 22) IMPORT UI BINDINGS
  // =========================
  function bindImportUi() {
    const paste = document.getElementById("pasteData");
    const btnStagePaste = document.getElementById("btnStagePaste");
    const btnCommit = document.getElementById("btnCommitStaged");
    const btnExportClean = document.getElementById("btnExportCleaned");
    const btnExportSummary = document.getElementById("btnExportSummary");
    const btnExportWorkbook = document.getElementById("btnExportWorkbook");

    if (btnStagePaste && paste) {
      btnStagePaste.addEventListener("click", () => {
        const txt = String(paste.value || "");
        if (!txt.trim()) {
          showToast("Paste is empty", "Paste TSV or CSV content first.", "warn");
          return;
        }
        stageDatasetFromText(txt, "Pasted dataset");
      });
    }

    if (btnCommit) btnCommit.addEventListener("click", () => commitStagedDataset());

    const fileInput = document.getElementById("fileUpload");
    const btnStageUpload = document.getElementById("btnStageUpload");
    if (btnStageUpload) {
      btnStageUpload.addEventListener("click", async () => {
        const f = fileInput && fileInput.files ? fileInput.files[0] : null;
        await stageDatasetFromUploadFile(f);
      });
    }
    if (fileInput) {
      fileInput.addEventListener("change", async () => {
        const f = fileInput.files ? fileInput.files[0] : null;
        if (f) await stageDatasetFromUploadFile(f);
      });
    }

    const btnLoadDefault = document.getElementById("btnLoadDefault");
    if (btnLoadDefault) btnLoadDefault.addEventListener("click", () => loadDefaultTrialDataIfAvailable());

    if (btnExportClean) btnExportClean.addEventListener("click", exportCleanedDatasetTsv);
    if (btnExportSummary) btnExportSummary.addEventListener("click", exportTreatmentSummaryCsv);
    if (btnExportWorkbook) btnExportWorkbook.addEventListener("click", exportWorkbookIfAvailable);
  }

  // =========================
  // 23) SENSITIVITY UI BINDINGS
  // =========================
  function bindSensitivityUi() {
    const btnRun = document.getElementById("btnRunSensitivity");
    if (btnRun) btnRun.addEventListener("click", () => {
      computeSensitivityGrid();
      renderSensitivitySummary();
      renderAiTab();
    });

    const btnExport = document.getElementById("btnExportSensitivity");
    if (btnExport) btnExport.addEventListener("click", exportSensitivityGridCsv);
  }

  // =========================
  // 24) AI UI BINDINGS
  // =========================
  function bindAiUi() {
    const btnCopyBrief = document.getElementById("btnCopyAiBrief");
    const btnCopyJson = document.getElementById("btnCopyResultsJson");
    const a = document.getElementById("aiBriefingText");
    const j = document.getElementById("resultsJson");

    if (btnCopyBrief) btnCopyBrief.addEventListener("click", () => copyToClipboard(a ? a.value : buildAiBriefingText()));
    if (btnCopyJson) btnCopyJson.addEventListener("click", () => copyToClipboard(j ? j.value : JSON.stringify(buildResultsPayload(), null, 2)));
  }

  // =========================
  // 25) APPENDIX UI BINDINGS
  // =========================
  function bindAppendixUi() {
    const btn = document.getElementById("btnOpenAppendix");
    if (btn) btn.addEventListener("click", openTechnicalAppendix);
  }

  // =========================
  // 26) SCENARIO UI BINDINGS
  // =========================
  function bindScenarioUi() {
    const btnSave = document.getElementById("btnSaveScenario");
    const btnLoad = document.getElementById("btnLoadScenario");
    const btnDelete = document.getElementById("btnDeleteScenario");

    if (btnSave) btnSave.addEventListener("click", handleSaveScenario);
    if (btnLoad) btnLoad.addEventListener("click", handleLoadScenario);
    if (btnDelete) btnDelete.addEventListener("click", handleDeleteScenario);

    renderScenarioDropdown();
  }

  // =========================
  // 27) TAB UI BINDINGS
  // =========================
  function bindTabs() {
    $$(".tab-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const t = btn.getAttribute("data-tab");
        if (t) setActiveTab(t);
      });
    });
  }

  // =========================
  // 28) INITIALISE
  // =========================
  function initEmptyPanels() {
    renderImportStatus("No dataset loaded yet. Load default data or import your trial file.");
    renderImportSummary({ checks: [], plotDeltas: [], treatmentSummary: [] }, { treatmentCol: null, replicateCol: null, yieldCol: null }, "None loaded", false);
    renderDataChecks({ checks: [] });
    renderOutputs();
    renderTreatments();
    computeBaseCaseResultsVsControl();
    renderBaseCaseBadge();
    renderLeaderboard(state.results.perTreatmentBaseCase, state.ui.resultsFilter);
    renderComparisonToControl(state.results.perTreatmentBaseCase, state.ui.resultsFilter);
    renderResultsNarrative(state.results.perTreatmentBaseCase, state.ui.resultsFilter);
    renderAiTab();
    renderSensitivitySummary();
  }

  async function init() {
    ensureToastRoot();

    bindTabs();
    bindImportUi();
    bindConfigInputs();
    bindResultsFilters();
    bindSensitivityUi();
    bindAiUi();
    bindAppendixUi();
    bindScenarioUi();

    syncConfigInputsFromModel();
    initEmptyPanels();

    // restore active scenario (if exists)
    const scenarios = loadScenariosFromStorage();
    const activeId = getActiveScenarioId();
    const active = activeId ? scenarios.find(s => s.id === activeId) : null;
    if (active) {
      applyScenario(active);
      showToast("Restored scenario", active.name || "Active scenario", "good");
    } else {
      // load default dataset if present
      await loadDefaultTrialDataIfAvailable();
    }

    // set initial tab
    setActiveTab(state.ui.activeTab || "results");

    // responsive plots
    window.addEventListener("resize", () => {
      plotNpvBar();
      plotCashflowLine();
    });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
