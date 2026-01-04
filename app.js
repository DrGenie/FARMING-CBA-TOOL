// Farming CBA Tool - Newcastle Business School
// Upgraded app.js: TSV/CSV import + paste pipeline, dictionary parsing, validation + Data Checks,
// replicate-specific control baselines, plot deltas, treatment summaries, discounted CBA with persistence + recurrence,
// sensitivity grid, scenario save/load to localStorage, Results comparison-to-control grid + leaderboard + narrative,
// exports (clean TSV, summary CSV, sensitivity CSV, workbook if XLSX available), AI Briefing prompt + Results JSON copy,
// and bottom-right toasts on every major action.

(() => {
  "use strict";

  // =========================
  // 0) CONSTANTS & UTILITIES
  // =========================
  const APP_KEY = "farming_cba_tool_v3";
  const LS_SCENARIOS_KEY = `${APP_KEY}:scenarios`;
  const LS_AUTOSAVE_KEY = `${APP_KEY}:autosave`;
  const LS_LAST_ACTIVE_TAB_KEY = `${APP_KEY}:active_tab`;

  const DEFAULT_SENSITIVITY = {
    priceMultipliers: [0.8, 0.9, 1.0, 1.1, 1.2],
    discountRatesPct: [4, 7, 10],
    persistenceYears: [1, 3, 5, 10],
    recurrenceYears: [0, 1, 2, 3, 5] // 0 = once at year 0
  };

  const DEFAULT_MODEL_CONFIG = {
    persistenceYearsBase: 5,
    effectTailMode: "step", // "step" means full delta until persistenceYears, then zero
    costTiming: "year0", // "year0" or "annual"
    recurrenceByTreatmentId: {}, // { [treatmentId]: { recurrenceYears: number } }
    notes: ""
  };

  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const isFiniteNumber = v => typeof v === "number" && Number.isFinite(v);
  const esc = s =>
    (s ?? "")
      .toString()
      .replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const fmt = n =>
    isFiniteNumber(n)
      ? Math.abs(n) >= 1000
        ? n.toLocaleString(undefined, { maximumFractionDigits: 0 })
        : n.toLocaleString(undefined, { maximumFractionDigits: 4 })
      : "n/a";

  const money = n => (isFiniteNumber(n) ? "$" + (Math.abs(n) >= 1000 ? n.toLocaleString(undefined, { maximumFractionDigits: 0 }) : n.toLocaleString(undefined, { maximumFractionDigits: 2 })) : "n/a");

  const percent = n => (isFiniteNumber(n) ? (Math.abs(n) >= 100 ? n.toLocaleString(undefined, { maximumFractionDigits: 1 }) : n.toLocaleString(undefined, { maximumFractionDigits: 2 })) + "%" : "n/a");

  function uid() {
    return Math.random().toString(36).slice(2, 10) + "_" + Date.now().toString(36);
  }

  function safeTrim(s) {
    return (s ?? "").toString().trim();
  }

  function parseNumber(x) {
    if (x === null || x === undefined) return NaN;
    if (typeof x === "number") return Number.isFinite(x) ? x : NaN;
    const s = String(x).trim();
    if (!s) return NaN;
    if (s === "?" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a" || s.toLowerCase() === "null") return NaN;
    const cleaned = s.replace(/[\$,]/g, "");
    const v = parseFloat(cleaned);
    return Number.isFinite(v) ? v : NaN;
  }

  function mean(arr) {
    let s = 0;
    let n = 0;
    for (const v of arr) {
      if (isFiniteNumber(v)) {
        s += v;
        n += 1;
      }
    }
    return n ? s / n : NaN;
  }

  function sd(arr) {
    const m = mean(arr);
    if (!isFiniteNumber(m)) return NaN;
    let s2 = 0;
    let n = 0;
    for (const v of arr) {
      if (isFiniteNumber(v)) {
        const d = v - m;
        s2 += d * d;
        n += 1;
      }
    }
    return n > 1 ? Math.sqrt(s2 / (n - 1)) : NaN;
  }

  function median(arr) {
    const clean = arr.filter(isFiniteNumber).slice().sort((a, b) => a - b);
    if (!clean.length) return NaN;
    const mid = Math.floor(clean.length / 2);
    return clean.length % 2 ? clean[mid] : (clean[mid - 1] + clean[mid]) / 2;
  }

  function quantile(arr, q) {
    const clean = arr.filter(isFiniteNumber).slice().sort((a, b) => a - b);
    if (!clean.length) return NaN;
    const pos = (clean.length - 1) * q;
    const base = Math.floor(pos);
    const rest = pos - base;
    if (clean[base + 1] === undefined) return clean[base];
    return clean[base] + rest * (clean[base + 1] - clean[base]);
  }

  function annuityFactor(N, rPct) {
    const r = rPct / 100;
    return r === 0 ? N : (1 - Math.pow(1 + r, -N)) / r;
  }

  function presentValue(series, ratePct) {
    let pv = 0;
    const r = ratePct / 100;
    for (let t = 0; t < series.length; t++) {
      pv += series[t] / Math.pow(1 + r, t);
    }
    return pv;
  }

  function irr(cf) {
    const hasPos = cf.some(v => v > 0);
    const hasNeg = cf.some(v => v < 0);
    if (!hasPos || !hasNeg) return NaN;

    const npvAt = r => {
      let s = 0;
      for (let t = 0; t < cf.length; t++) s += cf[t] / Math.pow(1 + r, t);
      return s;
    };

    let lo = -0.99;
    let hi = 5.0;
    let nLo = npvAt(lo);
    let nHi = npvAt(hi);

    if (nLo * nHi > 0) {
      for (let k = 0; k < 30 && nLo * nHi > 0; k++) {
        hi *= 1.5;
        nHi = npvAt(hi);
      }
      if (nLo * nHi > 0) return NaN;
    }

    for (let i = 0; i < 100; i++) {
      const mid = (lo + hi) / 2;
      const nMid = npvAt(mid);
      if (Math.abs(nMid) < 1e-9) return mid * 100;
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

  function paybackDiscounted(cf, ratePct) {
    const r = ratePct / 100;
    let cum = 0;
    for (let t = 0; t < cf.length; t++) {
      cum += cf[t] / Math.pow(1 + r, t);
      if (cum >= 0) return t;
    }
    return null;
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

  function safeSetText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
  }

  function safeSetValue(id, value) {
    const el = document.getElementById(id);
    if (el) el.value = value;
  }

  function safeGetValue(id) {
    const el = document.getElementById(id);
    return el ? el.value : "";
  }

  function safeGetNumber(id) {
    return parseNumber(safeGetValue(id));
  }

  function safeOnClick(id, handler) {
    const el = document.getElementById(id);
    if (!el) return false;
    el.addEventListener("click", e => {
      e.preventDefault();
      e.stopPropagation();
      handler(e);
    });
    return true;
  }

  function safeOnInput(id, handler) {
    const el = document.getElementById(id);
    if (!el) return false;
    el.addEventListener("input", handler);
    return true;
  }

  // =========================
  // 1) MODEL (kept compatible)
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
        "Applied trial comparing soil amendment and mechanical treatments against a control using replicated field plots.",
      objectives: "Quantify yield and gross margin impacts of soil amendment strategies.",
      activities: "Establish replicated plots, collect plot-level yield and cost data, and summarise economics.",
      stakeholders: "Producers, agronomists, government agencies, research partners.",
      lastUpdated: new Date().toISOString().slice(0, 10),
      goal:
        "Identify soil amendment packages that deliver higher yields and acceptable returns after accounting for additional costs.",
      withProject:
        "Growers consider adopting high-performing amendment packages on similar soils using evidence from the trial.",
      withoutProject:
        "Growers continue baseline practice without access to detailed economic evidence on amendments."
    },
    time: {
      startYear: new Date().getFullYear(),
      years: 10,
      discBase: 7,
      discLow: 4,
      discHigh: 10
    },
    outputs: [
      // Output value is used as the grain price per tonne for yield revenue.
      { id: uid(), name: "Grain yield", unit: "t/ha", value: 450, source: "Input Directly" }
    ],
    treatments: [
      // Will be populated from committed dataset.
      {
        id: uid(),
        name: "Control (baseline)",
        area: 100,
        adoption: 1,
        deltas: {}, // per ha delta vs control (for yield, this is delta t/ha)
        labourCost: 0,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "Farm Trials",
        isControl: true,
        notes: "Baseline"
      }
    ],
    config: JSON.parse(JSON.stringify(DEFAULT_MODEL_CONFIG)),
    sensitivity: JSON.parse(JSON.stringify(DEFAULT_SENSITIVITY))
  };

  // Ensure deltas exist for outputs
  function initTreatmentDeltas() {
    for (const t of model.treatments) {
      if (!t.deltas || typeof t.deltas !== "object") t.deltas = {};
      for (const o of model.outputs) {
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      }
      if (!isFiniteNumber(t.area)) t.area = parseNumber(t.area) || 0;
      if (!isFiniteNumber(t.adoption)) t.adoption = 1;
      if (!isFiniteNumber(t.labourCost)) t.labourCost = 0;
      if (!isFiniteNumber(t.materialsCost)) t.materialsCost = 0;
      if (!isFiniteNumber(t.servicesCost)) t.servicesCost = 0;
      if (!isFiniteNumber(t.capitalCost)) t.capitalCost = 0;
      if (typeof t.isControl !== "boolean") t.isControl = false;
      if (typeof t.constrained !== "boolean") t.constrained = true;
    }
  }
  initTreatmentDeltas();

  // =========================
  // 2) DATA STATE
  // =========================
  const state = {
    importStage: {
      dataText: "",
      dictText: "",
      sourceLabel: "",
      parsed: null, // { headers, rows, delimiter, mapping, dictionary }
      checks: []
    },
    dataset: {
      committed: false,
      sourceLabel: "",
      committedAt: null,
      headers: [],
      rows: [], // array of objects
      delimiter: "\t",
      dictionary: null, // { fields: [...], byName: Map }
      mapping: null, // column mapping used for computations
      checks: [],
      derived: {
        plotDeltas: [], // array of { idx, treatment, replicate, yield, controlYield, deltaYield, costs... }
        summaries: [], // per treatment summary
        controlByReplicate: new Map()
      }
    },
    ui: {
      resultsFilter: "all" // all | topNpv | topBcr | improvements
    },
    computed: {
      baseCase: null,
      perTreatment: [],
      comparisonGrid: null,
      sensitivityGrid: null
    }
  };

  // =========================
  // 3) PARSING: CSV/TSV + DICTIONARY
  // =========================
  function normaliseNewlines(text) {
    return String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  }

  function detectDelimiter(text) {
    const sample = normaliseNewlines(text).slice(0, 20000);
    const lines = sample.split("\n").filter(l => l.trim().length > 0).slice(0, 50);
    const countChar = ch => {
      let c = 0;
      for (const line of lines) c += (line.split(ch).length - 1);
      return c;
    };
    const tabs = countChar("\t");
    const commas = countChar(",");
    const semis = countChar(";");
    if (tabs >= commas && tabs >= semis) return "\t";
    if (commas >= semis) return ",";
    return ";";
  }

  // Robust CSV/DSV parser with quotes
  function parseDSV(text, delimiter) {
    const out = { headers: [], rows: [], errors: [] };
    const src = normaliseNewlines(text);
    const lines = src.split("\n").filter(l => l.length > 0);
    if (!lines.length) {
      out.errors.push("No data lines found.");
      return out;
    }

    const parseLine = line => {
      const fields = [];
      let cur = "";
      let inQ = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (inQ) {
          if (ch === '"') {
            if (line[i + 1] === '"') {
              cur += '"';
              i += 1;
            } else {
              inQ = false;
            }
          } else {
            cur += ch;
          }
        } else {
          if (ch === '"') inQ = true;
          else if (ch === delimiter) {
            fields.push(cur);
            cur = "";
          } else {
            cur += ch;
          }
        }
      }
      fields.push(cur);
      return fields;
    };

    // Find first non-empty header line
    let headerLineIdx = -1;
    for (let i = 0; i < lines.length; i++) {
      const raw = lines[i].trim();
      if (!raw) continue;
      headerLineIdx = i;
      break;
    }
    if (headerLineIdx < 0) {
      out.errors.push("Header line not found.");
      return out;
    }

    const headersRaw = parseLine(lines[headerLineIdx]).map(h => safeTrim(h));
    const headers = headersRaw.map((h, i) => (h ? h : `col_${i + 1}`));
    out.headers = headers;

    for (let i = headerLineIdx + 1; i < lines.length; i++) {
      const line = lines[i];
      if (!line || !line.trim()) continue;
      const parts = parseLine(line);
      const row = {};
      for (let c = 0; c < headers.length; c++) {
        row[headers[c]] = parts[c] !== undefined ? parts[c] : "";
      }
      out.rows.push(row);
    }

    return out;
  }

  // Split combined TXT that contains a dictionary section then data section
  function splitCombinedTextIfNeeded(text) {
    const src = normaliseNewlines(text);
    const upper = src.toUpperCase();

    const dictIdx = upper.search(/\bDATA\s+DICTIONARY\b/);
    const dataIdx = upper.search(/\bDATA\s+SECTION\b/);
    const dataIdx2 = upper.search(/\bBEGIN\s+DATA\b/);
    const dataIdx3 = upper.search(/\bDATA\b\s*$/m);

    // If no strong markers, return as data only.
    if (dictIdx === -1) return { dictText: "", dataText: src };

    // Identify best data start marker after dict
    let start = -1;
    const candidates = [dataIdx, dataIdx2, dataIdx3].filter(i => i !== -1 && i > dictIdx);
    if (candidates.length) start = Math.min(...candidates);
    if (start === -1) return { dictText: src, dataText: "" };

    const dictText = src.slice(dictIdx, start).replace(/^.*\bDATA\s+DICTIONARY\b[^\n]*\n/i, "");
    const dataText = src.slice(start).replace(/^.*\b(DATA\s+SECTION|BEGIN\s+DATA|DATA)\b[^\n]*\n/i, "");

    return { dictText: dictText.trim(), dataText: dataText.trim() };
  }

  function parseDictionaryCSV(dictText) {
    const text = normaliseNewlines(dictText || "");
    if (!text.trim()) return null;

    const delim = detectDelimiter(text);
    const parsed = parseDSV(text, delim);
    if (parsed.errors.length) return null;
    const headersLower = parsed.headers.map(h => h.toLowerCase());

    const getCol = candidates => {
      for (const c of candidates) {
        const idx = headersLower.indexOf(c);
        if (idx >= 0) return parsed.headers[idx];
      }
      return null;
    };

    const nameCol = getCol(["variable", "variable_name", "name", "field", "column"]);
    const labelCol = getCol(["label", "description", "field_label"]);
    const typeCol = getCol(["type", "data_type"]);
    const unitCol = getCol(["unit", "units"]);
    const missingCol = getCol(["missing", "missing_values", "na_values"]);
    const levelCol = getCol(["levels", "coding", "codes"]);

    const fields = [];
    for (const r of parsed.rows) {
      const name = safeTrim(nameCol ? r[nameCol] : "");
      if (!name) continue;
      fields.push({
        name,
        label: safeTrim(labelCol ? r[labelCol] : ""),
        type: safeTrim(typeCol ? r[typeCol] : ""),
        unit: safeTrim(unitCol ? r[unitCol] : ""),
        missing: safeTrim(missingCol ? r[missingCol] : ""),
        levels: safeTrim(levelCol ? r[levelCol] : "")
      });
    }
    const byName = new Map(fields.map(f => [f.name, f]));
    return { fields, byName };
  }

  function canonicaliseHeader(h) {
    return safeTrim(h).toLowerCase().replace(/\s+/g, " ").replace(/[^\w\s\/\-]/g, "").replace(/\s/g, "_");
  }

  function buildColumnMapping(headers, dictionary) {
    const canon = headers.map(h => ({ raw: h, c: canonicaliseHeader(h) }));
    const byCanon = new Map(canon.map(x => [x.c, x.raw]));
    const byRawLower = new Map(headers.map(h => [h.toLowerCase(), h]));

    const findByPatterns = patterns => {
      for (const p of patterns) {
        for (const x of canon) {
          if (p.test(x.c) || p.test(x.raw.toLowerCase())) return x.raw;
        }
      }
      return null;
    };

    // Primary structural columns
    const treatmentCol =
      findByPatterns([/(^|_)treatment(_|$)/, /amendment/, /treatment_name/, /treatmentlabel/]) ||
      (byCanon.get("amendment") || null);

    const replicateCol =
      findByPatterns([/(^|_)replicate(_|$)/, /block/, /rep(_|$)/, /trial_rep/]) || null;

    const plotCol =
      findByPatterns([/(^|_)plot(_|$)/, /plot_id/, /plotid/, /subplot/, /unit_id/, /row_id/]) || null;

    const yearCol =
      findByPatterns([/(^|_)year(_|$)/, /season/, /harvest_year/, /trial_year/]) || null;

    const isControlCol =
      findByPatterns([/(^|_)is_control(_|$)/, /control_flag/, /control\?/]) || null;

    const plotAreaCol =
      findByPatterns([/(^|_)plot_area(_|$)/, /plot_area_ha/, /area_plot/, /plot_size/, /plot_ha/]) || null;

    // Yield
    const yieldCol =
      findByPatterns([/(^|_)yield(_|$)/, /yield_t\/ha/, /yield_tha/, /grain_yield/, /yield_kg_ha/]) || null;

    // Costs (try to identify per-ha costs)
    const labourCostCol =
      findByPatterns([/labour/, /labor/, /pre_sowing_labour/, /amendment_labour/]) || null;

    const inputCostCol =
      findByPatterns([/treatment_input/, /input_cost/, /materials_cost/, /treatment_cost/, /amendment_cost/]) || null;

    const servicesCostCol =
      findByPatterns([/services_cost/, /contractor/, /machinery/, /fuel/, /operating_cost/]) || null;

    const capitalCostCol =
      findByPatterns([/capital_cost/, /capex/, /purchase_cost/, /asset_cost/]) || null;

    // Unit hints from headers / dictionary (for cost scaling)
    const unitHint = col => {
      if (!col) return "";
      const f = dictionary && dictionary.byName ? dictionary.byName.get(col) : null;
      const h = col.toLowerCase();
      const u = (f && (f.unit || f.label || "")) ? (f.unit || f.label || "") : "";
      if (/\b\/ha\b/i.test(h) || /\bper_ha\b/i.test(h) || /\bper ha\b/i.test(h) || /\b\/ha\b/i.test(u)) return "per_ha";
      if (/\b\/plot\b/i.test(h) || /\bper_plot\b/i.test(h) || /\bper plot\b/i.test(h) || /\bplot\b/i.test(u)) return "per_plot";
      return "";
    };

    return {
      treatmentCol,
      replicateCol,
      plotCol,
      yearCol,
      isControlCol,
      plotAreaCol,
      yieldCol,
      labourCostCol,
      inputCostCol,
      servicesCostCol,
      capitalCostCol,
      unitHints: {
        labourCost: unitHint(labourCostCol),
        inputCost: unitHint(inputCostCol),
        servicesCost: unitHint(servicesCostCol),
        capitalCost: unitHint(capitalCostCol),
        yield: unitHint(yieldCol)
      },
      headers,
      byCanon,
      byRawLower
    };
  }

  // Cost scaling rule (implemented deterministically, using unit hints if available):
  // - If value is expressed per hectare, use it as perHaCost.
  // - Otherwise, if plotArea is available and positive, treat value as per plot and divide by plotArea to get perHaCost.
  // - If plotArea is missing or non-positive, treat value as already per hectare and flag a check.
  function scaleToPerHa(value, unitHint, plotAreaHa) {
    const v = parseNumber(value);
    if (!isFiniteNumber(v)) return { perHa: NaN, assumed: "" };

    if (unitHint === "per_ha") return { perHa: v, assumed: "" };
    if (unitHint === "per_plot") {
      if (isFiniteNumber(plotAreaHa) && plotAreaHa > 0) return { perHa: v / plotAreaHa, assumed: "" };
      return { perHa: v, assumed: "Assumed per hectare because plot area was missing." };
    }

    // Infer from common header patterns embedded in the numeric field itself is not available here,
    // so default to per ha unless plot area exists and values look large relative to typical per ha.
    if (isFiniteNumber(plotAreaHa) && plotAreaHa > 0) {
      // Heuristic: if v is very large, it is likely total per plot, convert; otherwise keep as per ha.
      // This is only a fallback; the Data Checks panel will report when unit hints are missing.
      if (Math.abs(v) > 5000) return { perHa: v / plotAreaHa, assumed: "Inferred per plot and converted using plot area." };
    }
    return { perHa: v, assumed: "" };
  }

  // =========================
  // 4) VALIDATION & DATA CHECKS
  // =========================
  function buildDataChecks(parsed, mapping, dictionary) {
    const checks = [];

    const rows = parsed && parsed.rows ? parsed.rows : [];
    const headers = parsed && parsed.headers ? parsed.headers : [];

    const add = (id, severity, label, count, summary) => {
      checks.push({ id, severity, label, count, summary });
    };

    add("rows_total", "info", "Rows read", rows.length, rows.length ? "Data rows were read successfully." : "No data rows were found.");
    add("cols_total", "info", "Columns detected", headers.length, headers.length ? "Columns were detected from the header row." : "No columns were detected.");

    const missingCols = [];
    if (!mapping.treatmentCol) missingCols.push("Treatment");
    if (!mapping.yieldCol) missingCols.push("Yield");
    if (missingCols.length) {
      add("missing_required_cols", "error", "Missing required columns", missingCols.length, "Missing: " + missingCols.join(", ") + ".");
    }

    // Missing treatments
    if (mapping.treatmentCol) {
      let miss = 0;
      for (const r of rows) {
        if (!safeTrim(r[mapping.treatmentCol])) miss++;
      }
      if (miss) add("missing_treatment", "error", "Missing treatment label", miss, "Some rows have no treatment label.");
      else add("missing_treatment", "ok", "Missing treatment label", 0, "All rows have a treatment label.");
    }

    // Missing yield
    if (mapping.yieldCol) {
      let miss = 0;
      let nonNum = 0;
      for (const r of rows) {
        const v = r[mapping.yieldCol];
        if (!safeTrim(v)) miss++;
        else if (!isFiniteNumber(parseNumber(v))) nonNum++;
      }
      if (miss) add("missing_yield", "error", "Missing yield values", miss, "Some rows have missing yield values.");
      else add("missing_yield", "ok", "Missing yield values", 0, "No missing yield values detected.");
      if (nonNum) add("nonnumeric_yield", "error", "Non-numeric yield values", nonNum, "Some yield values could not be parsed as numbers.");
      else add("nonnumeric_yield", "ok", "Non-numeric yield values", 0, "All yield values are numeric or empty.");
    }

    // Replicate coverage
    if (mapping.replicateCol) {
      let miss = 0;
      for (const r of rows) {
        if (!safeTrim(r[mapping.replicateCol])) miss++;
      }
      if (miss) add("missing_replicate", "warn", "Missing replicate identifier", miss, "Some rows have missing replicate identifiers.");
      else add("missing_replicate", "ok", "Missing replicate identifier", 0, "All rows have replicate identifiers.");
    } else {
      add("missing_replicate_col", "warn", "Replicate column not detected", 1, "Replicate-specific control baselines will fall back to a global control baseline.");
    }

    // Plot id duplicates
    if (mapping.plotCol) {
      const seen = new Map();
      let dups = 0;
      for (const r of rows) {
        const k = safeTrim(r[mapping.plotCol]);
        if (!k) continue;
        seen.set(k, (seen.get(k) || 0) + 1);
      }
      for (const [, c] of seen.entries()) if (c > 1) dups += (c - 1);
      if (dups) add("duplicate_plot_ids", "warn", "Duplicate plot identifiers", dups, "Some plot identifiers appear more than once.");
      else add("duplicate_plot_ids", "ok", "Duplicate plot identifiers", 0, "No duplicate plot identifiers detected.");
    }

    // Plot area availability when costs exist without explicit per-ha units
    const costCols = [mapping.labourCostCol, mapping.inputCostCol, mapping.servicesCostCol, mapping.capitalCostCol].filter(Boolean);
    if (costCols.length) {
      const needArea =
        (mapping.unitHints.labourCost !== "per_ha") ||
        (mapping.unitHints.inputCost !== "per_ha") ||
        (mapping.unitHints.servicesCost !== "per_ha") ||
        (mapping.unitHints.capitalCost !== "per_ha");

      if (needArea && !mapping.plotAreaCol) {
        add(
          "missing_plot_area",
          "warn",
          "Plot area column not detected",
          1,
          "Cost scaling used a per-hectare assumption where units were unclear."
        );
      } else {
        add("missing_plot_area", "ok", "Plot area column not detected", 0, "Plot area is available or cost units are clearly per hectare.");
      }
    }

    // Dictionary alignment
    if (dictionary && dictionary.fields && dictionary.fields.length) {
      let missingInData = 0;
      for (const f of dictionary.fields) {
        if (!headers.includes(f.name)) missingInData++;
      }
      if (missingInData) add("dict_missing_vars", "warn", "Dictionary variables not present in data", missingInData, "Some dictionary fields do not match data columns.");
      else add("dict_missing_vars", "ok", "Dictionary variables not present in data", 0, "Dictionary fields match data columns.");
    }

    return checks;
  }

  function renderDataChecks(checks) {
    const listEl = document.getElementById("dataChecksList");
    const summaryEl = document.getElementById("dataChecksSummary");

    if (summaryEl) {
      const errs = checks.filter(c => c.severity === "error").reduce((a, c) => a + (c.count || 0), 0);
      const warns = checks.filter(c => c.severity === "warn").reduce((a, c) => a + (c.count || 0), 0);
      summaryEl.textContent = `Errors: ${errs.toLocaleString()}  Warnings: ${warns.toLocaleString()}  Checks: ${checks.length.toLocaleString()}`;
    }

    if (!listEl) return;

    listEl.innerHTML = "";
    for (const c of checks) {
      const row = document.createElement("div");
      row.className = `check ${c.severity}`;
      row.innerHTML = `
        <div class="check-head">
          <div class="check-title">${esc(c.label)}</div>
          <div class="check-count">${esc(String(c.count ?? ""))}</div>
        </div>
        <div class="check-summary">${esc(c.summary || "")}</div>
      `;
      listEl.appendChild(row);
    }
  }

  // =========================
  // 5) DERIVATIONS: CONTROL BASELINES, PLOT DELTAS, SUMMARIES
  // =========================
  function detectControlTreatmentName(rows, mapping) {
    const tcol = mapping.treatmentCol;
    if (!tcol) return null;

    // If explicit isControl column exists, use it.
    if (mapping.isControlCol) {
      // Return a special marker; control rows will be identified by the boolean column.
      return "__CONTROL_BY_FLAG__";
    }

    // Otherwise infer by name
    const counts = new Map();
    for (const r of rows) {
      const t = safeTrim(r[tcol]);
      if (!t) continue;
      counts.set(t, (counts.get(t) || 0) + 1);
    }
    // Prefer the most frequent label containing "control"
    let best = null;
    let bestCount = -1;
    for (const [name, cnt] of counts.entries()) {
      const lc = name.toLowerCase();
      if (lc.includes("control")) {
        if (cnt > bestCount) {
          best = name;
          bestCount = cnt;
        }
      }
    }
    // If none, return the most frequent label
    if (!best && counts.size) {
      for (const [name, cnt] of counts.entries()) {
        if (cnt > bestCount) {
          best = name;
          bestCount = cnt;
        }
      }
    }
    return best;
  }

  function rowIsControl(row, mapping, inferredControlName) {
    if (!mapping.treatmentCol) return false;
    const tname = safeTrim(row[mapping.treatmentCol]);
    if (!tname) return false;

    if (mapping.isControlCol) {
      const v = safeTrim(row[mapping.isControlCol]).toLowerCase();
      if (v === "1" || v === "true" || v === "yes" || v === "y") return true;
      if (v === "0" || v === "false" || v === "no" || v === "n") return false;
      // Fall back to name check if flag is unparseable
    }

    if (inferredControlName && inferredControlName !== "__CONTROL_BY_FLAG__") {
      return tname === inferredControlName;
    }
    return tname.toLowerCase().includes("control");
  }

  function computeControlBaselines(rows, mapping, inferredControlName) {
    const replCol = mapping.replicateCol;
    const ycol = mapping.yieldCol;

    const controlByRep = new Map(); // key -> { yieldMean, costMeans }
    const globalControl = { yields: [], labour: [], input: [], services: [], capital: [], plotAreas: [] };

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!rowIsControl(r, mapping, inferredControlName)) continue;

      const repl = replCol ? safeTrim(r[replCol]) : "__NO_REPLICATE__";
      const y = parseNumber(ycol ? r[ycol] : NaN);

      const plotAreaHa = mapping.plotAreaCol ? parseNumber(r[mapping.plotAreaCol]) : NaN;

      const labourScaled = mapping.labourCostCol ? scaleToPerHa(r[mapping.labourCostCol], mapping.unitHints.labourCost, plotAreaHa) : { perHa: NaN, assumed: "" };
      const inputScaled = mapping.inputCostCol ? scaleToPerHa(r[mapping.inputCostCol], mapping.unitHints.inputCost, plotAreaHa) : { perHa: NaN, assumed: "" };
      const servicesScaled = mapping.servicesCostCol ? scaleToPerHa(r[mapping.servicesCostCol], mapping.unitHints.servicesCost, plotAreaHa) : { perHa: NaN, assumed: "" };
      const capitalScaled = mapping.capitalCostCol ? scaleToPerHa(r[mapping.capitalCostCol], mapping.unitHints.capitalCost, plotAreaHa) : { perHa: NaN, assumed: "" };

      if (!controlByRep.has(repl)) controlByRep.set(repl, { yields: [], labour: [], input: [], services: [], capital: [] });

      const g = controlByRep.get(repl);
      if (isFiniteNumber(y)) g.yields.push(y);
      if (isFiniteNumber(labourScaled.perHa)) g.labour.push(labourScaled.perHa);
      if (isFiniteNumber(inputScaled.perHa)) g.input.push(inputScaled.perHa);
      if (isFiniteNumber(servicesScaled.perHa)) g.services.push(servicesScaled.perHa);
      if (isFiniteNumber(capitalScaled.perHa)) g.capital.push(capitalScaled.perHa);

      if (isFiniteNumber(y)) globalControl.yields.push(y);
      if (isFiniteNumber(labourScaled.perHa)) globalControl.labour.push(labourScaled.perHa);
      if (isFiniteNumber(inputScaled.perHa)) globalControl.input.push(inputScaled.perHa);
      if (isFiniteNumber(servicesScaled.perHa)) globalControl.services.push(servicesScaled.perHa);
      if (isFiniteNumber(capitalScaled.perHa)) globalControl.capital.push(capitalScaled.perHa);
      if (isFiniteNumber(plotAreaHa)) globalControl.plotAreas.push(plotAreaHa);
    }

    // Convert arrays to means
    const result = new Map();
    for (const [rep, g] of controlByRep.entries()) {
      result.set(rep, {
        yieldMean: mean(g.yields),
        labourPerHa: mean(g.labour),
        inputPerHa: mean(g.input),
        servicesPerHa: mean(g.services),
        capitalPerHa: mean(g.capital)
      });
    }

    const global = {
      yieldMean: mean(globalControl.yields),
      labourPerHa: mean(globalControl.labour),
      inputPerHa: mean(globalControl.input),
      servicesPerHa: mean(globalControl.services),
      capitalPerHa: mean(globalControl.capital)
    };

    return { byReplicate: result, global };
  }

  function computePlotDeltas(rows, mapping, baselines, inferredControlName) {
    const tcol = mapping.treatmentCol;
    const replCol = mapping.replicateCol;
    const ycol = mapping.yieldCol;

    const deltas = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const treatment = safeTrim(tcol ? r[tcol] : "");
      const replicate = replCol ? safeTrim(r[replCol]) : "__NO_REPLICATE__";
      const isControl = rowIsControl(r, mapping, inferredControlName);

      const plotAreaHa = mapping.plotAreaCol ? parseNumber(r[mapping.plotAreaCol]) : NaN;
      const y = parseNumber(ycol ? r[ycol] : NaN);

      const base = baselines.byReplicate.get(replicate) || baselines.global;
      const controlYield = base ? base.yieldMean : NaN;

      const deltaYield = isFiniteNumber(y) && isFiniteNumber(controlYield) ? (y - controlYield) : NaN;

      const labourScaled = mapping.labourCostCol ? scaleToPerHa(r[mapping.labourCostCol], mapping.unitHints.labourCost, plotAreaHa) : { perHa: NaN, assumed: "" };
      const inputScaled = mapping.inputCostCol ? scaleToPerHa(r[mapping.inputCostCol], mapping.unitHints.inputCost, plotAreaHa) : { perHa: NaN, assumed: "" };
      const servicesScaled = mapping.servicesCostCol ? scaleToPerHa(r[mapping.servicesCostCol], mapping.unitHints.servicesCost, plotAreaHa) : { perHa: NaN, assumed: "" };
      const capitalScaled = mapping.capitalCostCol ? scaleToPerHa(r[mapping.capitalCostCol], mapping.unitHints.capitalCost, plotAreaHa) : { perHa: NaN, assumed: "" };

      deltas.push({
        idx: i,
        treatment,
        replicate,
        isControl,
        plotId: mapping.plotCol ? safeTrim(r[mapping.plotCol]) : "",
        year: mapping.yearCol ? safeTrim(r[mapping.yearCol]) : "",
        plotAreaHa: isFiniteNumber(plotAreaHa) ? plotAreaHa : NaN,
        yieldTHa: isFiniteNumber(y) ? y : NaN,
        controlYieldTHa: isFiniteNumber(controlYield) ? controlYield : NaN,
        deltaYieldTHa: isFiniteNumber(deltaYield) ? deltaYield : NaN,
        labourCostPerHa: isFiniteNumber(labourScaled.perHa) ? labourScaled.perHa : NaN,
        inputCostPerHa: isFiniteNumber(inputScaled.perHa) ? inputScaled.perHa : NaN,
        servicesCostPerHa: isFiniteNumber(servicesScaled.perHa) ? servicesScaled.perHa : NaN,
        capitalCostPerHa: isFiniteNumber(capitalScaled.perHa) ? capitalScaled.perHa : NaN,
        scalingAssumptions: [labourScaled.assumed, inputScaled.assumed, servicesScaled.assumed, capitalScaled.assumed].filter(Boolean)
      });
    }
    return deltas;
  }

  function computeTreatmentSummaries(plotDeltas) {
    const groups = new Map();
    for (const d of plotDeltas) {
      const name = d.treatment || "";
      if (!name) continue;
      if (!groups.has(name)) {
        groups.set(name, {
          treatment: name,
          isControl: !!d.isControl,
          yields: [],
          deltaYields: [],
          labour: [],
          input: [],
          services: [],
          capital: [],
          missingYield: 0,
          missingDelta: 0,
          missingCosts: 0,
          n: 0
        });
      }
      const g = groups.get(name);
      g.n += 1;
      if (isFiniteNumber(d.yieldTHa)) g.yields.push(d.yieldTHa); else g.missingYield += 1;
      if (isFiniteNumber(d.deltaYieldTHa)) g.deltaYields.push(d.deltaYieldTHa); else g.missingDelta += 1;

      const costParts = [d.labourCostPerHa, d.inputCostPerHa, d.servicesCostPerHa, d.capitalCostPerHa];
      if (costParts.every(x => !isFiniteNumber(x))) g.missingCosts += 1;
      if (isFiniteNumber(d.labourCostPerHa)) g.labour.push(d.labourCostPerHa);
      if (isFiniteNumber(d.inputCostPerHa)) g.input.push(d.inputCostPerHa);
      if (isFiniteNumber(d.servicesCostPerHa)) g.services.push(d.servicesCostPerHa);
      if (isFiniteNumber(d.capitalCostPerHa)) g.capital.push(d.capitalCostPerHa);

      if (d.isControl) g.isControl = true;
    }

    const summaries = [];
    for (const g of groups.values()) {
      const yieldMean = mean(g.yields);
      const yieldSd = sd(g.yields);
      const deltaMean = mean(g.deltaYields);
      const deltaSd = sd(g.deltaYields);

      const labourMean = mean(g.labour);
      const inputMean = mean(g.input);
      const servicesMean = mean(g.services);
      const capitalMean = mean(g.capital);

      summaries.push({
        treatment: g.treatment,
        isControl: g.isControl,
        n: g.n,
        yieldMeanTHa: yieldMean,
        yieldSdTHa: yieldSd,
        deltaYieldMeanTHa: deltaMean,
        deltaYieldSdTHa: deltaSd,
        labourCostPerHa: labourMean,
        inputCostPerHa: inputMean,
        servicesCostPerHa: servicesMean,
        capitalCostPerHa: capitalMean,
        missingYield: g.missingYield,
        missingDelta: g.missingDelta,
        missingCostsRows: g.missingCosts
      });
    }

    // Rank control first, then others by deltaYieldMean descending
    summaries.sort((a, b) => {
      if (a.isControl && !b.isControl) return -1;
      if (!a.isControl && b.isControl) return 1;
      const da = isFiniteNumber(a.deltaYieldMeanTHa) ? a.deltaYieldMeanTHa : -Infinity;
      const db = isFiniteNumber(b.deltaYieldMeanTHa) ? b.deltaYieldMeanTHa : -Infinity;
      return db - da;
    });

    return summaries;
  }

  function enrichChecksWithControlCoverage(checks, mapping, rows, inferredControlName) {
    if (!mapping.treatmentCol || !mapping.yieldCol) return checks;

    // replicate control coverage
    if (!mapping.replicateCol) {
      // global only
      const controlRows = rows.filter(r => rowIsControl(r, mapping, inferredControlName)).length;
      if (!controlRows) {
        checks.push({
          id: "no_control_rows",
          severity: "error",
          label: "No control rows detected",
          count: 1,
          summary: "No control rows were detected, so deltas and comparisons to control cannot be computed reliably."
        });
      } else {
        checks.push({
          id: "no_control_rows",
          severity: "ok",
          label: "No control rows detected",
          count: 0,
          summary: "Control rows were detected."
        });
      }
      return checks;
    }

    const replCol = mapping.replicateCol;
    const repls = new Map();
    for (const r of rows) {
      const rep = safeTrim(r[replCol]) || "__MISSING__";
      if (!repls.has(rep)) repls.set(rep, { total: 0, control: 0 });
      const obj = repls.get(rep);
      obj.total += 1;
      if (rowIsControl(r, mapping, inferredControlName)) obj.control += 1;
    }
    let missingControlReps = 0;
    const examples = [];
    for (const [rep, obj] of repls.entries()) {
      if (rep === "__MISSING__") continue;
      if (obj.control === 0) {
        missingControlReps += 1;
        if (examples.length < 8) examples.push(rep);
      }
    }
    if (missingControlReps) {
      checks.push({
        id: "replicates_without_control",
        severity: "warn",
        label: "Replicates without control plots",
        count: missingControlReps,
        summary: "Some replicates have no control plot rows. Deltas in those replicates use the global control baseline. Examples: " + examples.join(", ") + "."
      });
    } else {
      checks.push({
        id: "replicates_without_control",
        severity: "ok",
        label: "Replicates without control plots",
        count: 0,
        summary: "All replicates have at least one control plot row."
      });
    }

    return checks;
  }

  // =========================
  // 6) COMMIT DATASET INTO TOOL MODEL
  // =========================
  function commitParsedToState(parsed, dictionary, mapping, sourceLabel) {
    const checks = buildDataChecks(parsed, mapping, dictionary);
    const inferredControlName = detectControlTreatmentName(parsed.rows, mapping);
    enrichChecksWithControlCoverage(checks, mapping, parsed.rows, inferredControlName);

    const baselines = computeControlBaselines(parsed.rows, mapping, inferredControlName);
    const plotDeltas = computePlotDeltas(parsed.rows, mapping, baselines, inferredControlName);
    const summaries = computeTreatmentSummaries(plotDeltas);

    state.dataset.committed = true;
    state.dataset.sourceLabel = sourceLabel || "Imported dataset";
    state.dataset.committedAt = new Date().toISOString();
    state.dataset.headers = parsed.headers.slice();
    state.dataset.rows = parsed.rows.slice();
    state.dataset.delimiter = parsed.delimiter;
    state.dataset.dictionary = dictionary;
    state.dataset.mapping = mapping;
    state.dataset.checks = checks;
    state.dataset.derived.controlByReplicate = baselines.byReplicate;
    state.dataset.derived.plotDeltas = plotDeltas;
    state.dataset.derived.summaries = summaries;

    // Push into the tool model: treatments and yield deltas
    const yieldOutput = model.outputs.find(o => o.name.toLowerCase().includes("yield")) || model.outputs[0];
    const yieldId = yieldOutput ? yieldOutput.id : null;

    // Choose control
    let controlSummary = summaries.find(s => s.isControl);
    if (!controlSummary && summaries.length) controlSummary = summaries[0];

    // Rebuild treatments list
    const rebuilt = [];
    for (const s of summaries) {
      const tId = uid();
      const isControl = controlSummary ? s.treatment === controlSummary.treatment : !!s.isControl;
      const t = {
        id: tId,
        name: s.treatment,
        area: 100,
        adoption: 1,
        deltas: {},
        labourCost: isFiniteNumber(s.labourCostPerHa) ? s.labourCostPerHa : 0,
        materialsCost: isFiniteNumber(s.inputCostPerHa) ? s.inputCostPerHa : 0,
        servicesCost: isFiniteNumber(s.servicesCostPerHa) ? s.servicesCostPerHa : 0,
        capitalCost: isFiniteNumber(s.capitalCostPerHa) ? s.capitalCostPerHa : 0,
        constrained: true,
        source: "Imported dataset",
        isControl: !!isControl,
        notes: ""
      };
      for (const o of model.outputs) t.deltas[o.id] = 0;

      if (yieldId) {
        const d = isControl ? 0 : (isFiniteNumber(s.deltaYieldMeanTHa) ? s.deltaYieldMeanTHa : 0);
        t.deltas[yieldId] = d;
      }
      rebuilt.push(t);

      // Ensure recurrence config exists
      if (!model.config.recurrenceByTreatmentId[tId]) {
        model.config.recurrenceByTreatmentId[tId] = { recurrenceYears: 0 };
      }
    }

    // If we have a known control, keep it as the single control in UI model.
    rebuilt.forEach(tt => (tt.isControl = false));
    if (controlSummary) {
      const idx = rebuilt.findIndex(x => x.name === controlSummary.treatment);
      if (idx >= 0) rebuilt[idx].isControl = true;
    } else if (rebuilt.length) {
      rebuilt[0].isControl = true;
    }

    model.treatments = rebuilt;
    initTreatmentDeltas();

    // Update UI renderers (if present)
    renderDataChecks(checks);
    renderAllSafe();
    calcAndRenderAllSafe();

    // Persist autosave
    autosaveToLocalStorage();

    showToast("Dataset committed. Treatments, control baseline, plot deltas, and summaries updated.");
  }

  // =========================
  // 7) DISCOUNTED CBA ENGINE (WITH PERSISTENCE + RECURRENCE)
  // =========================
  function getGrainPricePerTonne() {
    const out = model.outputs.find(o => o.name.toLowerCase().includes("yield")) || model.outputs[0];
    const v = out ? parseNumber(out.value) : NaN;
    return isFiniteNumber(v) ? v : 0;
  }

  function getControlFromModel() {
    return model.treatments.find(t => t.isControl) || null;
  }

  function getControlYieldMeanFromDataset() {
    // Prefer dataset summaries if available
    const sums = state.dataset.derived.summaries || [];
    const ctrl = sums.find(s => s.isControl) || null;
    if (ctrl && isFiniteNumber(ctrl.yieldMeanTHa)) return ctrl.yieldMeanTHa;
    // Otherwise infer from model: delta for control is 0, but need level. Use NaN then.
    return NaN;
  }

  function getTreatmentYieldMeanFromDataset(treatmentName) {
    const sums = state.dataset.derived.summaries || [];
    const s = sums.find(x => x.treatment === treatmentName);
    return s && isFiniteNumber(s.yieldMeanTHa) ? s.yieldMeanTHa : NaN;
  }

  function effectFactor(yearIndex, persistenceYears, tailMode) {
    // yearIndex is 1..N for benefits
    const p = Math.max(0, Math.floor(persistenceYears || 0));
    if (p <= 0) return 0;
    if (tailMode === "step") return yearIndex <= p ? 1 : 0;

    // Fallback: step
    return yearIndex <= p ? 1 : 0;
  }

  function buildTreatmentSeriesAbsolute(t, control, opts) {
    const years = opts.years;
    const rate = opts.ratePct;
    const price = opts.pricePerT;
    const persistenceYears = opts.persistenceYears;
    const tailMode = opts.tailMode;
    const costTiming = opts.costTiming;

    const area = isFiniteNumber(t.area) ? t.area : 0;
    const adoption = clamp(isFiniteNumber(t.adoption) ? t.adoption : 1, 0, 1);

    // Determine control yield level and delta
    const controlYield = opts.controlYieldTHa;
    const deltaYield = opts.deltaYieldTHa;

    // Costs per ha (from model fields)
    const labourPerHa = isFiniteNumber(t.labourCost) ? t.labourCost : 0;
    const materialsPerHa = isFiniteNumber(t.materialsCost) ? t.materialsCost : 0;
    const servicesPerHa = isFiniteNumber(t.servicesCost) ? t.servicesCost : 0;
    const capitalPerHa = isFiniteNumber(t.capitalCost) ? t.capitalCost : 0;

    const recurrenceYears = (opts.recurrenceYears !== undefined && opts.recurrenceYears !== null) ? Math.max(0, Math.floor(opts.recurrenceYears)) : 0;

    // Build series
    const benefits = new Array(years + 1).fill(0);
    const costs = new Array(years + 1).fill(0);

    // Year 0 capital always at year 0
    costs[0] += capitalPerHa * area * adoption;

    // Operational costs: either year 0 (application) with recurrence, or annual
    const opCostPerHa = (labourPerHa + materialsPerHa + servicesPerHa);

    if (costTiming === "annual") {
      for (let y = 1; y <= years; y++) {
        costs[y] += opCostPerHa * area * adoption;
      }
    } else {
      // Apply at year 0 and repeat every recurrenceYears if recurrenceYears > 0
      const applyAt = y0 => {
        if (y0 >= 0 && y0 <= years) costs[y0] += opCostPerHa * area * adoption;
      };
      if (recurrenceYears === 0) {
        applyAt(0);
      } else {
        for (let y = 0; y <= years; y += recurrenceYears) applyAt(y);
      }
    }

    // Benefits: each year 1..years, baseline control yield plus persistent delta component
    for (let y = 1; y <= years; y++) {
      const eff = effectFactor(y, persistenceYears, tailMode);
      const yld = (isFiniteNumber(controlYield) ? controlYield : 0) + (isFiniteNumber(deltaYield) ? deltaYield : 0) * eff;
      benefits[y] += yld * price * area * adoption;
    }

    const cashflow = benefits.map((b, i) => b - costs[i]);

    const pvBenefits = presentValue(benefits, rate);
    const pvCosts = presentValue(costs, rate);
    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts > 0 ? pvBenefits / pvCosts : NaN;
    const roi = pvCosts > 0 ? (npv / pvCosts) * 100 : NaN;
    const irrPct = irr(cashflow);
    const payback = paybackDiscounted(cashflow, rate);

    return {
      benefits,
      costs,
      cashflow,
      pvBenefits,
      pvCosts,
      npv,
      bcr,
      roi,
      irrPct,
      paybackYears: payback
    };
  }

  function computePerTreatmentResultsBaseCase() {
    const years = Math.max(1, Math.floor(model.time.years || 10));
    const rate = parseNumber(model.time.discBase);
    const price = getGrainPricePerTonne();

    const control = getControlFromModel();
    const ctrlYield = getControlYieldMeanFromDataset();

    // Fallback: if dataset is not committed, derive control yield from current default RAW-like treatment names
    const controlYieldFallback = isFiniteNumber(ctrlYield) ? ctrlYield : 0;

    const persistenceYears = Math.max(0, Math.floor(model.config.persistenceYearsBase || years));
    const tailMode = model.config.effectTailMode || "step";
    const costTiming = model.config.costTiming || "year0";

    const out = [];
    for (const t of model.treatments) {
      const recCfg = model.config.recurrenceByTreatmentId[t.id] || { recurrenceYears: 0 };
      const recurrenceYears = Math.max(0, Math.floor(parseNumber(recCfg.recurrenceYears) || 0));

      const deltaYield = (() => {
        const yOut = model.outputs.find(o => o.name.toLowerCase().includes("yield")) || model.outputs[0];
        if (!yOut) return 0;
        const d = parseNumber(t.deltas[yOut.id]);
        return isFiniteNumber(d) ? d : 0;
      })();

      // Use dataset yield mean when available for absolute yield of the treatment (delta based), but benefits are computed from control+delta
      const res = buildTreatmentSeriesAbsolute(t, control, {
        years,
        ratePct: isFiniteNumber(rate) ? rate : 0,
        pricePerT: isFiniteNumber(price) ? price : 0,
        persistenceYears,
        tailMode,
        costTiming,
        recurrenceYears,
        controlYieldTHa: controlYieldFallback,
        deltaYieldTHa: deltaYield
      });

      out.push({ treatmentId: t.id, name: t.name, isControl: !!t.isControl, recurrenceYears, ...res });
    }

    // Ranking: by NPV descending, excluding control for rank display but keeping in table
    const nonControl = out.filter(x => !x.isControl).slice().sort((a, b) => (b.npv - a.npv));
    const rankMap = new Map();
    nonControl.forEach((x, i) => rankMap.set(x.treatmentId, i + 1));
    for (const r of out) r.rank = r.isControl ? 0 : (rankMap.get(r.treatmentId) || null);

    return out;
  }

  function computeComparisonGrid(perTreatment) {
    const control = perTreatment.find(x => x.isControl) || null;

    const cols = [];
    if (control) cols.push({ key: control.treatmentId, label: "Control (baseline)", isControl: true });
    for (const t of perTreatment) {
      if (t.isControl) continue;
      cols.push({ key: t.treatmentId, label: t.name, isControl: false });
      cols.push({ key: t.treatmentId + ":delta", label: "Delta vs control", isDelta: true, baseKey: t.treatmentId });
    }

    const get = (t, metric) => (t && isFiniteNumber(t[metric]) ? t[metric] : NaN);
    const delta = (t, metric) => {
      if (!control) return NaN;
      const a = get(t, metric);
      const b = get(control, metric);
      return isFiniteNumber(a) && isFiniteNumber(b) ? (a - b) : NaN;
    };

    const rows = [
      { key: "pvBenefits", label: "Present value of benefits", fmt: money, kind: "level" },
      { key: "pvCosts", label: "Present value of costs", fmt: money, kind: "level" },
      { key: "npv", label: "Net present value", fmt: money, kind: "level" },
      { key: "bcr", label: "Benefit cost ratio", fmt: x => (isFiniteNumber(x) ? fmt(x) : "n/a"), kind: "ratio" },
      { key: "roi", label: "Return on investment", fmt: x => (isFiniteNumber(x) ? percent(x) : "n/a"), kind: "pct" },
      { key: "irrPct", label: "Internal rate of return", fmt: x => (isFiniteNumber(x) ? percent(x) : "n/a"), kind: "pct" },
      { key: "paybackYears", label: "Discounted payback year", fmt: x => (x === null || x === undefined ? "Not reached" : String(x)), kind: "int" },
      { key: "rank", label: "Rank by net present value", fmt: x => (x === null || x === undefined ? "" : String(x)), kind: "int" },
      { key: "deltaNpv", label: "Delta net present value vs control", fmt: money, kind: "delta", compute: (t) => delta(t, "npv") },
      { key: "deltaPvCosts", label: "Delta present value of costs vs control", fmt: money, kind: "delta", compute: (t) => delta(t, "pvCosts") }
    ];

    const matrix = rows.map(r => {
      const row = { rowKey: r.key, label: r.label, cells: [] };
      for (const c of cols) {
        if (c.isDelta) {
          const base = perTreatment.find(x => x.treatmentId === c.baseKey);
          // For delta columns: show delta for primary metrics rows, blank for precomputed delta rows
          if (r.kind === "delta" && r.compute) row.cells.push({ value: r.compute(base), mode: "delta" });
          else if (r.kind === "level" || r.kind === "ratio" || r.kind === "pct" || r.kind === "int") row.cells.push({ value: delta(base, r.key), mode: "delta" });
          else row.cells.push({ value: NaN, mode: "delta" });
        } else {
          const base = perTreatment.find(x => x.treatmentId === c.key);
          if (r.compute) row.cells.push({ value: r.compute(base), mode: "level" });
          else row.cells.push({ value: base ? base[r.key] : NaN, mode: "level" });
        }
      }
      return row;
    });

    return { cols, rows, matrix, controlKey: control ? control.treatmentId : null };
  }

  // =========================
  // 8) SENSITIVITY GRID
  // =========================
  function computeSensitivityGrid() {
    const years = Math.max(1, Math.floor(model.time.years || 10));
    const controlYield = getControlYieldMeanFromDataset();
    const ctrlYield = isFiniteNumber(controlYield) ? controlYield : 0;

    const basePrice = getGrainPricePerTonne();
    const priceMultipliers = (model.sensitivity && model.sensitivity.priceMultipliers) ? model.sensitivity.priceMultipliers.slice() : DEFAULT_SENSITIVITY.priceMultipliers.slice();
    const discountRatesPct = (model.sensitivity && model.sensitivity.discountRatesPct) ? model.sensitivity.discountRatesPct.slice() : DEFAULT_SENSITIVITY.discountRatesPct.slice();
    const persistenceYearsList = (model.sensitivity && model.sensitivity.persistenceYears) ? model.sensitivity.persistenceYears.slice() : DEFAULT_SENSITIVITY.persistenceYears.slice();
    const recurrenceYearsList = (model.sensitivity && model.sensitivity.recurrenceYears) ? model.sensitivity.recurrenceYears.slice() : DEFAULT_SENSITIVITY.recurrenceYears.slice();

    const tailMode = model.config.effectTailMode || "step";
    const costTiming = model.config.costTiming || "year0";

    const yOut = model.outputs.find(o => o.name.toLowerCase().includes("yield")) || model.outputs[0];
    const yieldId = yOut ? yOut.id : null;

    const grid = [];
    for (const t of model.treatments) {
      const deltaYield = yieldId ? parseNumber(t.deltas[yieldId]) : 0;
      const dy = isFiniteNumber(deltaYield) ? deltaYield : 0;

      for (const pm of priceMultipliers) {
        const price = (isFiniteNumber(basePrice) ? basePrice : 0) * pm;
        for (const dr of discountRatesPct) {
          const rate = parseNumber(dr);
          for (const py of persistenceYearsList) {
            const pYears = Math.max(0, Math.floor(parseNumber(py) || 0));
            for (const ry of recurrenceYearsList) {
              const recYears = Math.max(0, Math.floor(parseNumber(ry) || 0));

              const res = buildTreatmentSeriesAbsolute(t, getControlFromModel(), {
                years,
                ratePct: isFiniteNumber(rate) ? rate : 0,
                pricePerT: price,
                persistenceYears: pYears,
                tailMode,
                costTiming,
                recurrenceYears: recYears,
                controlYieldTHa: ctrlYield,
                deltaYieldTHa: dy
              });

              grid.push({
                treatment: t.name,
                isControl: !!t.isControl,
                priceMultiplier: pm,
                pricePerTonne: price,
                discountRatePct: rate,
                persistenceYears: pYears,
                recurrenceYears: recYears,
                pvBenefits: res.pvBenefits,
                pvCosts: res.pvCosts,
                npv: res.npv,
                benefitCostRatio: res.bcr,
                returnOnInvestmentPct: res.roi
              });
            }
          }
        }
      }
    }

    return grid;
  }

  function renderSensitivityGrid(grid) {
    const root = document.getElementById("sensitivityGrid");
    const status = document.getElementById("sensitivityStatus");
    if (status) status.textContent = grid && grid.length ? `Sensitivity grid updated. Rows: ${grid.length.toLocaleString()}.` : "Sensitivity grid is empty.";

    if (!root) return;
    root.innerHTML = "";
    if (!grid || !grid.length) {
      root.innerHTML = `<p class="small muted">No sensitivity results to display.</p>`;
      return;
    }

    // Compact table with first N rows, full export available
    const maxShow = 120;
    const show = grid.slice(0, maxShow);

    const tbl = document.createElement("table");
    tbl.className = "summary-table";
    tbl.innerHTML = `
      <thead>
        <tr>
          <th>Treatment</th>
          <th>Price multiplier</th>
          <th>Discount rate</th>
          <th>Persistence years</th>
          <th>Recurrence years</th>
          <th>Net present value</th>
          <th>Benefit cost ratio</th>
          <th>Return on investment</th>
        </tr>
      </thead>
      <tbody>
        ${show.map(r => `
          <tr>
            <td>${esc(r.treatment)}${r.isControl ? " (Control)" : ""}</td>
            <td>${fmt(r.priceMultiplier)}</td>
            <td>${fmt(r.discountRatePct)}%</td>
            <td>${esc(String(r.persistenceYears))}</td>
            <td>${r.recurrenceYears === 0 ? "Once at year 0" : esc(String(r.recurrenceYears))}</td>
            <td>${money(r.npv)}</td>
            <td>${isFiniteNumber(r.benefitCostRatio) ? fmt(r.benefitCostRatio) : "n/a"}</td>
            <td>${isFiniteNumber(r.returnOnInvestmentPct) ? percent(r.returnOnInvestmentPct) : "n/a"}</td>
          </tr>
        `).join("")}
      </tbody>
    `;
    root.appendChild(tbl);

    if (grid.length > maxShow) {
      const note = document.createElement("p");
      note.className = "small muted";
      note.textContent = `Showing ${maxShow.toLocaleString()} of ${grid.length.toLocaleString()} rows. Use the sensitivity export for the full grid.`;
      root.appendChild(note);
    }
  }

  // =========================
  // 9) RESULTS RENDERING: LEADERBOARD + GRID + NARRATIVE + FILTERS
  // =========================
  function applyResultsFilter(perTreatment) {
    const nonControl = perTreatment.filter(x => !x.isControl);

    if (state.ui.resultsFilter === "topNpv") {
      const top = nonControl.slice().sort((a, b) => (b.npv - a.npv)).slice(0, 5);
      const ids = new Set(top.map(x => x.treatmentId));
      return perTreatment.filter(x => x.isControl || ids.has(x.treatmentId));
    }

    if (state.ui.resultsFilter === "topBcr") {
      const top = nonControl
        .slice()
        .sort((a, b) => ((isFiniteNumber(b.bcr) ? b.bcr : -Infinity) - (isFiniteNumber(a.bcr) ? a.bcr : -Infinity)))
        .slice(0, 5);
      const ids = new Set(top.map(x => x.treatmentId));
      return perTreatment.filter(x => x.isControl || ids.has(x.treatmentId));
    }

    if (state.ui.resultsFilter === "improvements") {
      const control = perTreatment.find(x => x.isControl) || null;
      if (!control) return perTreatment;
      return perTreatment.filter(x => x.isControl || (isFiniteNumber(x.npv) && isFiniteNumber(control.npv) && (x.npv - control.npv) > 0));
    }

    return perTreatment;
  }

  function renderLeaderboard(perTreatment) {
    const root = document.getElementById("resultsLeaderboard");
    if (!root) return;

    const control = perTreatment.find(x => x.isControl) || null;
    const items = perTreatment
      .filter(x => !x.isControl)
      .slice()
      .sort((a, b) => (b.npv - a.npv))
      .map(x => {
        const dNpv = control ? (x.npv - control.npv) : NaN;
        return { ...x, deltaNpv: dNpv };
      });

    root.innerHTML = "";
    if (!items.length) {
      root.innerHTML = `<p class="small muted">No treatments available to rank.</p>`;
      return;
    }

    const tbl = document.createElement("table");
    tbl.className = "summary-table";
    tbl.innerHTML = `
      <thead>
        <tr>
          <th>Rank</th>
          <th>Treatment</th>
          <th>Net present value</th>
          <th>Delta vs control</th>
          <th>Benefit cost ratio</th>
          <th>Return on investment</th>
        </tr>
      </thead>
      <tbody>
        ${items.map(x => {
          const cls = isFiniteNumber(x.deltaNpv) ? (x.deltaNpv >= 0 ? "pos" : "neg") : "";
          return `
            <tr>
              <td>${x.rank ?? ""}</td>
              <td>${esc(x.name)}</td>
              <td>${money(x.npv)}</td>
              <td class="${cls}">${isFiniteNumber(x.deltaNpv) ? money(x.deltaNpv) : "n/a"}</td>
              <td>${isFiniteNumber(x.bcr) ? fmt(x.bcr) : "n/a"}</td>
              <td>${isFiniteNumber(x.roi) ? percent(x.roi) : "n/a"}</td>
            </tr>
          `;
        }).join("")}
      </tbody>
    `;
    root.appendChild(tbl);
  }

  function renderComparisonGrid(grid) {
    const root = document.getElementById("comparisonGrid");
    if (!root) return;

    root.innerHTML = "";
    if (!grid || !grid.matrix || !grid.cols || !grid.matrix.length) {
      root.innerHTML = `<p class="small muted">No comparison grid available.</p>`;
      return;
    }

    const table = document.createElement("table");
    table.className = "comparison-grid summary-table";

    const thead = document.createElement("thead");
    thead.innerHTML = `
      <tr>
        <th class="sticky-col">Indicator</th>
        ${grid.cols.map(c => {
          const cls = c.isControl ? "sticky-control" : (c.isDelta ? "delta-col" : "");
          return `<th class="${cls}">${esc(c.label)}</th>`;
        }).join("")}
      </tr>
    `;
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    for (let r = 0; r < grid.matrix.length; r++) {
      const rowDef = grid.rows[r];
      const tr = document.createElement("tr");
      tr.innerHTML = `<td class="sticky-col">${esc(rowDef.label)}</td>`;

      for (let c = 0; c < grid.cols.length; c++) {
        const col = grid.cols[c];
        const cell = grid.matrix[r].cells[c];
        const v = cell ? cell.value : NaN;

        let cls = "";
        if (col.isControl) cls += " sticky-control";
        if (col.isDelta || cell.mode === "delta") {
          if (isFiniteNumber(v)) cls += (v >= 0 ? " pos" : " neg");
          else cls += " muted";
        }

        let text = "";
        if (rowDef.key === "bcr") text = isFiniteNumber(v) ? fmt(v) : "n/a";
        else if (rowDef.key === "roi" || rowDef.key === "irrPct") text = isFiniteNumber(v) ? percent(v) : "n/a";
        else if (rowDef.key === "paybackYears") text = (v === null || v === undefined || Number.isNaN(v)) ? "Not reached" : String(v);
        else if (rowDef.key === "rank") text = (v === null || v === undefined || Number.isNaN(v)) ? "" : String(v);
        else text = rowDef.fmt ? rowDef.fmt(v) : (isFiniteNumber(v) ? fmt(v) : "n/a");

        tr.innerHTML += `<td class="${cls.trim()}">${esc(text)}</td>`;
      }

      tbody.appendChild(tr);
    }
    table.appendChild(tbody);
    root.appendChild(table);
  }

  function renderWhatThisMeans(perTreatment) {
    const root = document.getElementById("resultsNarrative");
    if (!root) return;

    const control = perTreatment.find(x => x.isControl) || null;
    const nonControl = perTreatment.filter(x => !x.isControl).slice().sort((a, b) => (b.npv - a.npv));

    if (!control || !nonControl.length) {
      root.textContent = "A control baseline and at least one treatment are needed to interpret the comparison results.";
      return;
    }

    const top = nonControl[0];
    const bottom = nonControl[nonControl.length - 1];

    const deltaTopNpv = top.npv - control.npv;
    const deltaTopCost = top.pvCosts - control.pvCosts;
    const deltaTopBen = top.pvBenefits - control.pvBenefits;

    const deltaBottomNpv = bottom.npv - control.npv;

    const years = Math.max(1, Math.floor(model.time.years || 10));
    const price = getGrainPricePerTonne();
    const rate = parseNumber(model.time.discBase);

    const text =
      `The table compares each treatment against the control baseline using discounted results over ${years} years at a discount rate of ${isFiniteNumber(rate) ? fmt(rate) : "n/a"} percent and a grain price of ${money(price)} per tonne. ` +
      `The control is a reference point. Differences against control are shown in the delta columns. ` +
      `The highest ranked treatment by net present value is ${top.name}. Its net present value is ${money(top.npv)} compared with the control at ${money(control.npv)}, which is a difference of ${money(deltaTopNpv)}. ` +
      `That difference is explained by changes in present value of benefits of ${money(deltaTopBen)} and changes in present value of costs of ${money(deltaTopCost)}. ` +
      `A positive difference means the treatment generates a larger discounted net return than the control under the current persistence and recurrence settings. ` +
      `The lowest ranked treatment by net present value is ${bottom.name}. Its net present value is ${money(bottom.npv)} compared with the control, a difference of ${money(deltaBottomNpv)}. ` +
      `If a treatment has a weaker result, the most direct levers are a larger and more persistent yield lift, a lower application or operating cost, a lower recurrence frequency, or a setting where benefits are valued at a higher price and discounted less heavily.`;

    root.textContent = text;
  }

  function bindResultsFilters() {
    const bind = (id, filter, label) => {
      const el = document.getElementById(id);
      if (!el) return;
      el.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        state.ui.resultsFilter = filter;
        calcAndRenderAllSafe();
        showToast(label);
      });
    };

    bind("filterAll", "all", "Results filter set to show all treatments.");
    bind("filterTopNpv", "topNpv", "Results filter set to top five treatments by net present value.");
    bind("filterTopBcr", "topBcr", "Results filter set to top five treatments by benefit cost ratio.");
    bind("filterImprovements", "improvements", "Results filter set to treatments with improved net present value compared with control.");
  }

  // =========================
  // 10) EXPORTS
  // =========================
  function exportCleanedDatasetTSV() {
    if (!state.dataset.committed || !state.dataset.headers.length) {
      showToast("No committed dataset is available to export.");
      return;
    }
    const headers = state.dataset.headers.slice();
    const rows = state.dataset.rows.slice();

    const tsvLines = [];
    tsvLines.push(headers.map(h => h.replace(/\t/g, " ")).join("\t"));
    for (const r of rows) {
      const line = headers.map(h => {
        const v = r[h];
        const s = v === null || v === undefined ? "" : String(v);
        return s.replace(/\t/g, " ");
      }).join("\t");
      tsvLines.push(line);
    }

    const fname = (model.project.name || "project").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");
    downloadFile(`${fname}_cleaned_dataset.tsv`, tsvLines.join("\n"), "text/tab-separated-values");
    showToast("Cleaned dataset TSV downloaded.");
  }

  function exportTreatmentSummaryCSV() {
    const perTreatment = state.computed.perTreatment && state.computed.perTreatment.length ? state.computed.perTreatment : computePerTreatmentResultsBaseCase();
    const control = perTreatment.find(x => x.isControl) || null;

    const rows = [];
    rows.push([
      "Treatment",
      "IsControl",
      "RankByNpv",
      "PvBenefits",
      "PvCosts",
      "NetPresentValue",
      "BenefitCostRatio",
      "ReturnOnInvestmentPct",
      "InternalRateOfReturnPct",
      "PaybackYear",
      "DeltaNpvVsControl",
      "DeltaPvCostsVsControl"
    ]);

    for (const t of perTreatment) {
      const dNpv = control ? (t.npv - control.npv) : NaN;
      const dCost = control ? (t.pvCosts - control.pvCosts) : NaN;
      rows.push([
        t.name,
        t.isControl ? "1" : "0",
        t.rank == null ? "" : String(t.rank),
        isFiniteNumber(t.pvBenefits) ? String(t.pvBenefits) : "",
        isFiniteNumber(t.pvCosts) ? String(t.pvCosts) : "",
        isFiniteNumber(t.npv) ? String(t.npv) : "",
        isFiniteNumber(t.bcr) ? String(t.bcr) : "",
        isFiniteNumber(t.roi) ? String(t.roi) : "",
        isFiniteNumber(t.irrPct) ? String(t.irrPct) : "",
        t.paybackYears == null ? "" : String(t.paybackYears),
        isFiniteNumber(dNpv) ? String(dNpv) : "",
        isFiniteNumber(dCost) ? String(dCost) : ""
      ]);
    }

    const csv = rows
      .map(r => r.map(x => `"${String(x ?? "").replace(/"/g, '""')}"`).join(","))
      .join("\r\n");

    const fname = (model.project.name || "project").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");
    downloadFile(`${fname}_treatment_summary.csv`, csv, "text/csv");
    showToast("Treatment summary CSV downloaded.");
  }

  function exportSensitivityGridCSV() {
    const grid = state.computed.sensitivityGrid && state.computed.sensitivityGrid.length ? state.computed.sensitivityGrid : computeSensitivityGrid();
    const rows = [];
    rows.push([
      "Treatment",
      "IsControl",
      "PriceMultiplier",
      "PricePerTonne",
      "DiscountRatePct",
      "PersistenceYears",
      "RecurrenceYears",
      "PvBenefits",
      "PvCosts",
      "NetPresentValue",
      "BenefitCostRatio",
      "ReturnOnInvestmentPct"
    ]);

    for (const r of grid) {
      rows.push([
        r.treatment,
        r.isControl ? "1" : "0",
        String(r.priceMultiplier),
        String(r.pricePerTonne),
        String(r.discountRatePct),
        String(r.persistenceYears),
        String(r.recurrenceYears),
        isFiniteNumber(r.pvBenefits) ? String(r.pvBenefits) : "",
        isFiniteNumber(r.pvCosts) ? String(r.pvCosts) : "",
        isFiniteNumber(r.npv) ? String(r.npv) : "",
        isFiniteNumber(r.benefitCostRatio) ? String(r.benefitCostRatio) : "",
        isFiniteNumber(r.returnOnInvestmentPct) ? String(r.returnOnInvestmentPct) : ""
      ]);
    }

    const csv = rows
      .map(r => r.map(x => `"${String(x ?? "").replace(/"/g, '""')}"`).join(","))
      .join("\r\n");

    const fname = (model.project.name || "project").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");
    downloadFile(`${fname}_sensitivity_grid.csv`, csv, "text/csv");
    showToast("Sensitivity grid CSV downloaded.");
  }

  function exportWorkbookIfFeasible() {
    if (typeof XLSX === "undefined") {
      showToast("Workbook export is not available because the XLSX library was not detected.");
      return;
    }

    const wb = XLSX.utils.book_new();

    // Cleaned data
    if (state.dataset.committed && state.dataset.headers.length) {
      const aoa = [state.dataset.headers.slice()];
      for (const r of state.dataset.rows) {
        aoa.push(state.dataset.headers.map(h => (r[h] ?? "")));
      }
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(wb, ws, "CleanedData");
    }

    // Treatment summary
    const perTreatment = state.computed.perTreatment && state.computed.perTreatment.length ? state.computed.perTreatment : computePerTreatmentResultsBaseCase();
    {
      const control = perTreatment.find(x => x.isControl) || null;
      const aoa = [[
        "Treatment",
        "IsControl",
        "RankByNpv",
        "PvBenefits",
        "PvCosts",
        "NetPresentValue",
        "BenefitCostRatio",
        "ReturnOnInvestmentPct",
        "InternalRateOfReturnPct",
        "PaybackYear",
        "DeltaNpvVsControl",
        "DeltaPvCostsVsControl"
      ]];
      for (const t of perTreatment) {
        const dNpv = control ? (t.npv - control.npv) : NaN;
        const dCost = control ? (t.pvCosts - control.pvCosts) : NaN;
        aoa.push([
          t.name,
          t.isControl ? 1 : 0,
          t.rank ?? "",
          t.pvBenefits ?? "",
          t.pvCosts ?? "",
          t.npv ?? "",
          t.bcr ?? "",
          t.roi ?? "",
          t.irrPct ?? "",
          t.paybackYears ?? "",
          isFiniteNumber(dNpv) ? dNpv : "",
          isFiniteNumber(dCost) ? dCost : ""
        ]);
      }
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(wb, ws, "TreatmentSummary");
    }

    // Sensitivity
    {
      const grid = state.computed.sensitivityGrid && state.computed.sensitivityGrid.length ? state.computed.sensitivityGrid : computeSensitivityGrid();
      const aoa = [[
        "Treatment",
        "IsControl",
        "PriceMultiplier",
        "PricePerTonne",
        "DiscountRatePct",
        "PersistenceYears",
        "RecurrenceYears",
        "PvBenefits",
        "PvCosts",
        "NetPresentValue",
        "BenefitCostRatio",
        "ReturnOnInvestmentPct"
      ]];
      for (const r of grid) {
        aoa.push([
          r.treatment,
          r.isControl ? 1 : 0,
          r.priceMultiplier,
          r.pricePerTonne,
          r.discountRatePct,
          r.persistenceYears,
          r.recurrenceYears,
          r.pvBenefits,
          r.pvCosts,
          r.npv,
          r.benefitCostRatio,
          r.returnOnInvestmentPct
        ]);
      }
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(wb, ws, "SensitivityGrid");
    }

    // Data checks
    {
      const checks = state.dataset.checks || [];
      const aoa = [["Severity", "Label", "Count", "Summary"]];
      for (const c of checks) aoa.push([c.severity, c.label, c.count, c.summary]);
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(wb, ws, "DataChecks");
    }

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const fname = (model.project.name || "project").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");
    downloadFile(`${fname}_workbook.xlsx`, wbout, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Workbook export downloaded.");
  }

  // =========================
  // 11) AI BRIEFING TAB
  // =========================
  function buildResultsJsonObject() {
    const perTreatment = state.computed.perTreatment && state.computed.perTreatment.length ? state.computed.perTreatment : computePerTreatmentResultsBaseCase();
    const comparison = state.computed.comparisonGrid || computeComparisonGrid(perTreatment);

    const grid = state.computed.sensitivityGrid || null;

    const datasetSummary = (() => {
      if (!state.dataset.committed) return { committed: false };
      const sums = state.dataset.derived.summaries || [];
      const treatments = sums.map(s => ({
        treatment: s.treatment,
        isControl: !!s.isControl,
        n: s.n,
        yieldMeanTHa: s.yieldMeanTHa,
        yieldSdTHa: s.yieldSdTHa,
        deltaYieldMeanTHa: s.deltaYieldMeanTHa,
        deltaYieldSdTHa: s.deltaYieldSdTHa,
        labourCostPerHa: s.labourCostPerHa,
        inputCostPerHa: s.inputCostPerHa,
        servicesCostPerHa: s.servicesCostPerHa,
        capitalCostPerHa: s.capitalCostPerHa,
        missingYield: s.missingYield,
        missingDelta: s.missingDelta,
        missingCostsRows: s.missingCostsRows
      }));
      return {
        committed: true,
        sourceLabel: state.dataset.sourceLabel,
        committedAt: state.dataset.committedAt,
        rows: state.dataset.rows.length,
        columns: state.dataset.headers.length,
        mapping: state.dataset.mapping,
        checks: state.dataset.checks,
        treatments
      };
    })();

    return {
      tool: "Farming CBA Decision Tool",
      project: model.project,
      time: model.time,
      configuration: model.config,
      grainPricePerTonne: getGrainPricePerTonne(),
      baseCase: {
        perTreatment: perTreatment.map(t => ({
          name: t.name,
          isControl: t.isControl,
          rankByNpv: t.rank,
          pvBenefits: t.pvBenefits,
          pvCosts: t.pvCosts,
          netPresentValue: t.npv,
          benefitCostRatio: t.bcr,
          returnOnInvestmentPct: t.roi,
          internalRateOfReturnPct: t.irrPct,
          paybackYear: t.paybackYears,
          recurrenceYears: t.recurrenceYears
        })),
        comparisonGrid: {
          columns: comparison.cols,
          rows: comparison.rows
        }
      },
      sensitivityGrid: grid,
      datasetSummary
    };
  }

  function buildAiNarrativePrompt() {
    const years = Math.max(1, Math.floor(model.time.years || 10));
    const rate = parseNumber(model.time.discBase);
    const price = getGrainPricePerTonne();
    const persistence = Math.max(0, Math.floor(model.config.persistenceYearsBase || years));
    const costTiming = model.config.costTiming === "annual" ? "annual costs" : "application costs at year zero";
    const tailMode = model.config.effectTailMode === "step" ? "a step persistence assumption" : "a persistence assumption";

    const perTreatment = state.computed.perTreatment && state.computed.perTreatment.length ? state.computed.perTreatment : computePerTreatmentResultsBaseCase();
    const control = perTreatment.find(x => x.isControl) || null;
    const ranked = perTreatment.filter(x => !x.isControl).slice().sort((a, b) => (b.npv - a.npv));

    const top = ranked[0] || null;
    const mid = ranked.length ? ranked[Math.floor(ranked.length / 2)] : null;
    const low = ranked.length ? ranked[ranked.length - 1] : null;

    const controlText = control
      ? `The control baseline has a present value of benefits of ${money(control.pvBenefits)}, a present value of costs of ${money(control.pvCosts)}, and a net present value of ${money(control.npv)}.`
      : `A control baseline is not clearly identified in the current results.`;

    const topText = top && control
      ? `The highest net present value treatment is ${top.name}. Its net present value is ${money(top.npv)} and the difference compared with the control is ${money(top.npv - control.npv)}. Its benefit cost ratio is ${isFiniteNumber(top.bcr) ? fmt(top.bcr) : "not available"} and its return on investment is ${isFiniteNumber(top.roi) ? percent(top.roi) : "not available"}.`
      : "";

    const midText = mid && control
      ? `A mid ranked treatment is ${mid.name}. Its net present value is ${money(mid.npv)} and the difference compared with the control is ${money(mid.npv - control.npv)}.`
      : "";

    const lowText = low && control
      ? `The lowest net present value treatment is ${low.name}. Its net present value is ${money(low.npv)} and the difference compared with the control is ${money(low.npv - control.npv)}.`
      : "";

    // No bullets, no em dash, no abbreviations.
    const prompt =
      `Write a farmer facing interpretation and an internal technical note using the results provided in the JSON that will be pasted after this prompt. ` +
      `Use full terms and avoid abbreviations. Do not use bullet points. Do not use em dash punctuation. ` +
      `Explain what drives differences between treatments and the control baseline, focusing on yield changes, the timing and size of costs, recurrence settings, and the persistence assumption. ` +
      `The analysis horizon is ${years} years. The discount rate is ${isFiniteNumber(rate) ? fmt(rate) : "not available"} percent. The grain price is ${money(price)} per tonne. ` +
      `The persistence setting assumes the yield difference lasts for ${persistence} years under ${tailMode}. The cost timing uses ${costTiming}. ` +
      `${controlText} ` +
      `${topText} ` +
      `${midText} ` +
      `${lowText} ` +
      `When discussing treatments with weaker results, describe realistic ways the result could improve, such as lower application cost, less frequent recurrence, stronger yield lift, more persistent yield lift, or a context with a higher grain price or a lower discount rate. ` +
      `Do not tell the user what to choose. Do not impose decision rules or thresholds. ` +
      `Include a short section on uncertainty and data limitations and state which additional data would most improve confidence, such as additional seasons, more sites, or better measurement of costs and plot areas.`;

    return prompt;
  }

  function renderAiBriefing() {
    const promptEl = document.getElementById("aiBriefingPrompt");
    const jsonEl = document.getElementById("resultsJson");
    if (promptEl) promptEl.value = buildAiNarrativePrompt();
    if (jsonEl) jsonEl.value = JSON.stringify(buildResultsJsonObject(), null, 2);
  }

  function copyToClipboard(text, successToast, failToast) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(
        () => showToast(successToast),
        () => showToast(failToast)
      );
    } else {
      try {
        const ta = document.createElement("textarea");
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
        showToast(successToast);
      } catch {
        showToast(failToast);
      }
    }
  }

  // =========================
  // 12) SCENARIO SAVE/LOAD (localStorage)
  // =========================
  function loadScenariosFromLocalStorage() {
    try {
      const raw = localStorage.getItem(LS_SCENARIOS_KEY);
      if (!raw) return {};
      const obj = JSON.parse(raw);
      return obj && typeof obj === "object" ? obj : {};
    } catch {
      return {};
    }
  }

  function saveScenariosToLocalStorage(scenarios) {
    try {
      localStorage.setItem(LS_SCENARIOS_KEY, JSON.stringify(scenarios));
      return true;
    } catch {
      return false;
    }
  }

  function autosaveToLocalStorage() {
    try {
      const payload = {
        model,
        dataset: {
          committed: state.dataset.committed,
          sourceLabel: state.dataset.sourceLabel,
          committedAt: state.dataset.committedAt,
          headers: state.dataset.headers,
          rows: state.dataset.rows,
          delimiter: state.dataset.delimiter,
          dictionary: state.dataset.dictionary,
          mapping: state.dataset.mapping,
          checks: state.dataset.checks
        }
      };
      localStorage.setItem(LS_AUTOSAVE_KEY, JSON.stringify(payload));
    } catch {
      // ignore
    }
  }

  function restoreAutosaveFromLocalStorage() {
    try {
      const raw = localStorage.getItem(LS_AUTOSAVE_KEY);
      if (!raw) return false;
      const payload = JSON.parse(raw);
      if (!payload || typeof payload !== "object") return false;

      if (payload.model && typeof payload.model === "object") {
        // Shallow assign safe fields
        if (payload.model.project) Object.assign(model.project, payload.model.project);
        if (payload.model.time) Object.assign(model.time, payload.model.time);
        if (Array.isArray(payload.model.outputs)) model.outputs = payload.model.outputs;
        if (Array.isArray(payload.model.treatments)) model.treatments = payload.model.treatments;
        if (payload.model.config) model.config = payload.model.config;
        if (payload.model.sensitivity) model.sensitivity = payload.model.sensitivity;
      }

      initTreatmentDeltas();

      if (payload.dataset && payload.dataset.committed) {
        state.dataset.committed = true;
        state.dataset.sourceLabel = payload.dataset.sourceLabel || "";
        state.dataset.committedAt = payload.dataset.committedAt || null;
        state.dataset.headers = payload.dataset.headers || [];
        state.dataset.rows = payload.dataset.rows || [];
        state.dataset.delimiter = payload.dataset.delimiter || "\t";
        state.dataset.dictionary = payload.dataset.dictionary || null;
        state.dataset.mapping = payload.dataset.mapping || null;
        state.dataset.checks = payload.dataset.checks || [];

        // Recompute derived items deterministically from committed rows, if mapping exists.
        if (state.dataset.mapping && state.dataset.rows.length) {
          const inferredControlName = detectControlTreatmentName(state.dataset.rows, state.dataset.mapping);
          const baselines = computeControlBaselines(state.dataset.rows, state.dataset.mapping, inferredControlName);
          state.dataset.derived.controlByReplicate = baselines.byReplicate;
          state.dataset.derived.plotDeltas = computePlotDeltas(state.dataset.rows, state.dataset.mapping, baselines, inferredControlName);
          state.dataset.derived.summaries = computeTreatmentSummaries(state.dataset.derived.plotDeltas);
        }
      }

      showToast("Autosaved scenario restored.");
      return true;
    } catch {
      return false;
    }
  }

  function renderScenarioList() {
    const sel = document.getElementById("scenarioSelect");
    if (!sel) return;
    const scenarios = loadScenariosFromLocalStorage();
    const names = Object.keys(scenarios).sort((a, b) => a.localeCompare(b));
    sel.innerHTML = `<option value="">Select a saved scenario</option>` + names.map(n => `<option value="${esc(n)}">${esc(n)}</option>`).join("");
  }

  function saveScenario(name) {
    const nm = safeTrim(name);
    if (!nm) {
      showToast("Scenario name is required to save.");
      return;
    }
    const scenarios = loadScenariosFromLocalStorage();
    scenarios[nm] = {
      savedAt: new Date().toISOString(),
      model,
      dataset: {
        committed: state.dataset.committed,
        sourceLabel: state.dataset.sourceLabel,
        committedAt: state.dataset.committedAt,
        headers: state.dataset.headers,
        rows: state.dataset.rows,
        delimiter: state.dataset.delimiter,
        dictionary: state.dataset.dictionary,
        mapping: state.dataset.mapping,
        checks: state.dataset.checks
      }
    };
    if (saveScenariosToLocalStorage(scenarios)) {
      renderScenarioList();
      showToast("Scenario saved to local storage.");
    } else {
      showToast("Scenario could not be saved to local storage.");
    }
  }

  function loadScenario(name) {
    const nm = safeTrim(name);
    if (!nm) {
      showToast("Select a scenario to load.");
      return;
    }
    const scenarios = loadScenariosFromLocalStorage();
    const s = scenarios[nm];
    if (!s) {
      showToast("Scenario not found.");
      return;
    }
    try {
      // Restore model
      if (s.model && typeof s.model === "object") {
        if (s.model.project) Object.assign(model.project, s.model.project);
        if (s.model.time) Object.assign(model.time, s.model.time);
        model.outputs = Array.isArray(s.model.outputs) ? s.model.outputs : model.outputs;
        model.treatments = Array.isArray(s.model.treatments) ? s.model.treatments : model.treatments;
        model.config = s.model.config ? s.model.config : model.config;
        model.sensitivity = s.model.sensitivity ? s.model.sensitivity : model.sensitivity;
      }
      initTreatmentDeltas();

      // Restore dataset
      if (s.dataset && s.dataset.committed) {
        state.dataset.committed = true;
        state.dataset.sourceLabel = s.dataset.sourceLabel || "";
        state.dataset.committedAt = s.dataset.committedAt || null;
        state.dataset.headers = s.dataset.headers || [];
        state.dataset.rows = s.dataset.rows || [];
        state.dataset.delimiter = s.dataset.delimiter || "\t";
        state.dataset.dictionary = s.dataset.dictionary || null;
        state.dataset.mapping = s.dataset.mapping || null;
        state.dataset.checks = s.dataset.checks || [];

        if (state.dataset.mapping && state.dataset.rows.length) {
          const inferredControlName = detectControlTreatmentName(state.dataset.rows, state.dataset.mapping);
          const baselines = computeControlBaselines(state.dataset.rows, state.dataset.mapping, inferredControlName);
          state.dataset.derived.controlByReplicate = baselines.byReplicate;
          state.dataset.derived.plotDeltas = computePlotDeltas(state.dataset.rows, state.dataset.mapping, baselines, inferredControlName);
          state.dataset.derived.summaries = computeTreatmentSummaries(state.dataset.derived.plotDeltas);
        }
      } else {
        state.dataset.committed = false;
        state.dataset.headers = [];
        state.dataset.rows = [];
        state.dataset.dictionary = null;
        state.dataset.mapping = null;
        state.dataset.checks = [];
        state.dataset.derived.plotDeltas = [];
        state.dataset.derived.summaries = [];
      }

      renderAllSafe();
      renderDataChecks(state.dataset.checks || []);
      calcAndRenderAllSafe();
      autosaveToLocalStorage();
      showToast("Scenario loaded from local storage.");
    } catch {
      showToast("Scenario could not be loaded.");
    }
  }

  function deleteScenario(name) {
    const nm = safeTrim(name);
    if (!nm) {
      showToast("Select a scenario to delete.");
      return;
    }
    const scenarios = loadScenariosFromLocalStorage();
    if (!scenarios[nm]) {
      showToast("Scenario not found.");
      return;
    }
    delete scenarios[nm];
    if (saveScenariosToLocalStorage(scenarios)) {
      renderScenarioList();
      showToast("Scenario deleted.");
    } else {
      showToast("Scenario could not be deleted.");
    }
  }

  // =========================
  // 13) CONFIGURATION TAB: RECURRENCE PER TREATMENT
  // =========================
  function renderRecurrenceConfig() {
    const root = document.getElementById("recurrenceConfig");
    if (!root) return;

    root.innerHTML = "";
    if (!model.treatments.length) {
      root.innerHTML = `<p class="small muted">No treatments available.</p>`;
      return;
    }

    const wrap = document.createElement("div");
    wrap.className = "recurrence-wrap";

    // Global settings (persistence and cost timing) if corresponding inputs exist elsewhere, we still render here
    const global = document.createElement("div");
    global.className = "item";
    global.innerHTML = `
      <h4>Persistence and cost timing</h4>
      <div class="row-6">
        <div class="field">
          <label>Persistence years</label>
          <input id="cfgPersistenceYears" type="number" min="0" step="1" value="${esc(String(model.config.persistenceYearsBase || model.time.years || 10))}" />
        </div>
        <div class="field">
          <label>Cost timing</label>
          <select id="cfgCostTiming">
            <option value="year0" ${model.config.costTiming === "year0" ? "selected" : ""}>Costs applied at year zero</option>
            <option value="annual" ${model.config.costTiming === "annual" ? "selected" : ""}>Costs applied annually</option>
          </select>
        </div>
        <div class="field">
          <label>Effect tail mode</label>
          <select id="cfgTailMode">
            <option value="step" ${model.config.effectTailMode === "step" ? "selected" : ""}>Step persistence then zero</option>
          </select>
        </div>
        <div class="field"><label>&nbsp;</label><button id="cfgApplyGlobal" class="btn small">Apply</button></div>
      </div>
      <p class="small muted">Persistence affects the yield difference relative to control. Recurrence affects how often the non-capital costs are applied.</p>
    `;
    wrap.appendChild(global);

    const tbl = document.createElement("table");
    tbl.className = "summary-table";
    tbl.innerHTML = `
      <thead>
        <tr>
          <th>Treatment</th>
          <th>Recurrence</th>
        </tr>
      </thead>
      <tbody>
        ${model.treatments.map(t => {
          const cfg = model.config.recurrenceByTreatmentId[t.id] || { recurrenceYears: 0 };
          const ry = Math.max(0, Math.floor(parseNumber(cfg.recurrenceYears) || 0));
          return `
            <tr>
              <td>${esc(t.name)}${t.isControl ? " (Control)" : ""}</td>
              <td>
                <select data-rec-tid="${esc(t.id)}">
                  <option value="0" ${ry === 0 ? "selected" : ""}>Once at year zero</option>
                  <option value="1" ${ry === 1 ? "selected" : ""}>Every year</option>
                  <option value="2" ${ry === 2 ? "selected" : ""}>Every 2 years</option>
                  <option value="3" ${ry === 3 ? "selected" : ""}>Every 3 years</option>
                  <option value="5" ${ry === 5 ? "selected" : ""}>Every 5 years</option>
                </select>
              </td>
            </tr>
          `;
        }).join("")}
      </tbody>
    `;
    wrap.appendChild(tbl);

    root.appendChild(wrap);

    // Bind global apply
    const applyBtn = document.getElementById("cfgApplyGlobal");
    const pEl = document.getElementById("cfgPersistenceYears");
    const cEl = document.getElementById("cfgCostTiming");
    const tEl = document.getElementById("cfgTailMode");

    if (applyBtn && pEl && cEl && tEl) {
      applyBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        const p = Math.max(0, Math.floor(parseNumber(pEl.value) || 0));
        model.config.persistenceYearsBase = p;
        model.config.costTiming = cEl.value === "annual" ? "annual" : "year0";
        model.config.effectTailMode = tEl.value === "step" ? "step" : "step";
        calcAndRenderAllSafe();
        autosaveToLocalStorage();
        showToast("Configuration updated.");
      });
    }

    // Bind per-treatment recurrence dropdowns (event delegation on root is avoided; bind directly to each select present now)
    const selects = root.querySelectorAll("select[data-rec-tid]");
    selects.forEach(sel => {
      sel.addEventListener("change", e => {
        const tid = sel.getAttribute("data-rec-tid");
        const v = Math.max(0, Math.floor(parseNumber(sel.value) || 0));
        if (!model.config.recurrenceByTreatmentId[tid]) model.config.recurrenceByTreatmentId[tid] = { recurrenceYears: 0 };
        model.config.recurrenceByTreatmentId[tid].recurrenceYears = v;
        calcAndRenderAllSafe();
        autosaveToLocalStorage();
        showToast("Recurrence setting updated.");
      });
    });
  }

  // =========================
  // 14) IMPORT PIPELINE: UPLOAD + PASTE
  // =========================
  function renderImportStatus(text) {
    safeSetText("importStatus", text);
  }

  function parseDataAndDictionary(dataText, dictText, sourceLabel) {
    const split = splitCombinedTextIfNeeded(dataText);
    const finalDictText = (dictText && dictText.trim()) ? dictText : split.dictText;
    const finalDataText = split.dataText || dataText;

    const dictionary = parseDictionaryCSV(finalDictText);

    // Parse data
    const delim = detectDelimiter(finalDataText);
    const parsed = parseDSV(finalDataText, delim);
    parsed.delimiter = delim;

    const mapping = buildColumnMapping(parsed.headers, dictionary);

    const checks = buildDataChecks(parsed, mapping, dictionary);
    renderDataChecks(checks);

    state.importStage.dataText = finalDataText;
    state.importStage.dictText = finalDictText;
    state.importStage.sourceLabel = sourceLabel || "";
    state.importStage.parsed = { ...parsed, mapping, dictionary };
    state.importStage.checks = checks;

    // Optional preview elements
    const prev = document.getElementById("importPreview");
    if (prev) {
      const cols = parsed.headers.length;
      const rows = parsed.rows.length;
      prev.textContent = `Preview: ${rows.toLocaleString()} rows and ${cols.toLocaleString()} columns. Detected delimiter: ${delim === "\t" ? "tab" : delim}.`;
    }

    renderImportStatus("Data parsed. Review Data Checks, then commit to update treatments and results.");
    showToast("Import parsed.");
  }

  function handleUploadParse() {
    // Use existing parseExcel button id as the entry point, but accept TSV/CSV/TXT and Excel.
    const accept = ".tsv,.csv,.txt,.xlsx,.xlsm,.xlsb,.xls";
    const input = document.createElement("input");
    input.type = "file";
    input.accept = accept;
    input.multiple = true;
    input.style.display = "none";
    document.body.appendChild(input);

    input.addEventListener("change", async e => {
      const files = Array.from(input.files || []);
      document.body.removeChild(input);
      if (!files.length) return;

      // Identify dictionary file by filename heuristics
      let dictFile = null;
      let dataFile = null;

      for (const f of files) {
        const n = f.name.toLowerCase();
        if (n.includes("dictionary") && (n.endsWith(".csv") || n.endsWith(".tsv") || n.endsWith(".txt"))) dictFile = f;
      }
      // Prefer non-dictionary as data file
      dataFile = files.find(f => f !== dictFile) || files[0];

      try {
        let dictText = "";
        if (dictFile) dictText = await dictFile.text();

        // If Excel, try XLSX; else read as text.
        const lower = dataFile.name.toLowerCase();
        if ((lower.endsWith(".xlsx") || lower.endsWith(".xlsm") || lower.endsWith(".xlsb") || lower.endsWith(".xls")) && typeof XLSX !== "undefined") {
          const buf = await dataFile.arrayBuffer();
          const wb = XLSX.read(buf, { type: "array" });
          const sheetName = wb.SheetNames[0];
          const sheet = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

          // Convert to TSV for parser consistency
          const headers = rows.length ? Object.keys(rows[0]) : [];
          const tsvLines = [];
          tsvLines.push(headers.join("\t"));
          for (const r of rows) {
            tsvLines.push(headers.map(h => String(r[h] ?? "").replace(/\t/g, " ")).join("\t"));
          }

          parseDataAndDictionary(tsvLines.join("\n"), dictText, dataFile.name);
          renderImportStatus(`Excel parsed from sheet "${sheetName}". Ready to commit.`);
          showToast("Excel file parsed.");
          return;
        }

        const dataText = await dataFile.text();
        parseDataAndDictionary(dataText, dictText, dataFile.name);
        showToast("Text dataset parsed.");
      } catch (err) {
        console.error(err);
        renderImportStatus("Import failed. The file could not be parsed.");
        showToast("Import failed.");
      }
    });

    input.click();
  }

  function handleCommitImportStage() {
    if (!state.importStage.parsed || !state.importStage.parsed.rows) {
      showToast("No parsed data is available to commit.");
      return;
    }
    const parsed = state.importStage.parsed;
    commitParsedToState(
      { headers: parsed.headers, rows: parsed.rows, delimiter: parsed.delimiter },
      parsed.dictionary,
      parsed.mapping,
      state.importStage.sourceLabel || "Imported dataset"
    );
  }

  function handleParsePaste() {
    const ta = document.getElementById("pasteData");
    if (!ta) {
      showToast("Paste area was not found.");
      return;
    }
    const text = ta.value || "";
    if (!text.trim()) {
      showToast("Paste data is empty.");
      return;
    }

    // Optional dictionary paste area
    const dt = document.getElementById("pasteDictionary");
    const dictText = dt ? (dt.value || "") : "";

    parseDataAndDictionary(text, dictText, "Pasted data");
    showToast("Pasted data parsed.");
  }

  function handleCommitPaste() {
    handleCommitImportStage();
  }

  // =========================
  // 15) BASE UI RENDERERS (compatible with existing IDs)
  // =========================
  function renderOutputs() {
    const root = document.getElementById("outputsList");
    if (!root) return;

    root.innerHTML = "";
    for (const o of model.outputs) {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>Output: ${esc(o.name)}</h4>
        <div class="row-6">
          <div class="field"><label>Name</label><input id="out_name_${esc(o.id)}" value="${esc(o.name)}" /></div>
          <div class="field"><label>Unit</label><input id="out_unit_${esc(o.id)}" value="${esc(o.unit)}" /></div>
          <div class="field"><label>Value ($ per unit)</label><input id="out_value_${esc(o.id)}" type="number" step="0.01" value="${esc(String(o.value))}" /></div>
          <div class="field"><label>Source</label>
            <select id="out_source_${esc(o.id)}">
              ${["Farm Trials", "Plant Farm", "ABARES", "GRDC", "Input Directly"].map(s => `<option ${s === o.source ? "selected" : ""}>${esc(s)}</option>`).join("")}
            </select>
          </div>
        </div>
      `;
      root.appendChild(el);

      const nameId = `out_name_${o.id}`;
      const unitId = `out_unit_${o.id}`;
      const valId = `out_value_${o.id}`;
      const srcId = `out_source_${o.id}`;

      safeOnInput(nameId, e => {
        o.name = e.target.value;
        calcAndRenderAllSafe();
        autosaveToLocalStorage();
      });
      safeOnInput(unitId, e => {
        o.unit = e.target.value;
        calcAndRenderAllSafe();
        autosaveToLocalStorage();
      });
      safeOnInput(valId, e => {
        o.value = parseNumber(e.target.value);
        calcAndRenderAllSafe();
        autosaveToLocalStorage();
        showToast("Output value updated.");
      });
      const srcEl = document.getElementById(srcId);
      if (srcEl) {
        srcEl.addEventListener("change", e => {
          o.source = e.target.value;
          autosaveToLocalStorage();
          showToast("Output source updated.");
        });
      }
    }
  }

  function renderTreatments() {
    const root = document.getElementById("treatmentsList");
    if (!root) return;

    root.innerHTML = "";
    for (const t of model.treatments) {
      const mats = isFiniteNumber(t.materialsCost) ? t.materialsCost : 0;
      const serv = isFiniteNumber(t.servicesCost) ? t.servicesCost : 0;
      const lab = isFiniteNumber(t.labourCost) ? t.labourCost : 0;
      const total = mats + serv + lab;

      const yOut = model.outputs.find(o => o.name.toLowerCase().includes("yield")) || model.outputs[0];
      const yDelta = yOut ? parseNumber(t.deltas[yOut.id]) : 0;

      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>Treatment: ${esc(t.name)}</h4>
        <div class="row">
          <div class="field"><label>Name</label><input id="tr_name_${esc(t.id)}" value="${esc(t.name)}" /></div>
          <div class="field"><label>Area (ha)</label><input id="tr_area_${esc(t.id)}" type="number" step="0.01" value="${esc(String(t.area))}" /></div>
          <div class="field"><label>Control vs treatment</label>
            <select id="tr_isControl_${esc(t.id)}">
              <option value="treatment" ${!t.isControl ? "selected" : ""}>Treatment</option>
              <option value="control" ${t.isControl ? "selected" : ""}>Control</option>
            </select>
          </div>
        </div>
        <div class="row-6">
          <div class="field"><label>Materials cost ($ per ha)</label><input id="tr_mats_${esc(t.id)}" type="number" step="0.01" value="${esc(String(mats))}" /></div>
          <div class="field"><label>Services cost ($ per ha)</label><input id="tr_serv_${esc(t.id)}" type="number" step="0.01" value="${esc(String(serv))}" /></div>
          <div class="field"><label>Labour cost ($ per ha)</label><input id="tr_lab_${esc(t.id)}" type="number" step="0.01" value="${esc(String(lab))}" /></div>
          <div class="field"><label>Total cost ($ per ha)</label><input id="tr_total_${esc(t.id)}" type="number" step="0.01" value="${esc(String(total))}" readonly /></div>
          <div class="field"><label>Capital cost ($ per ha, year 0)</label><input id="tr_cap_${esc(t.id)}" type="number" step="0.01" value="${esc(String(t.capitalCost || 0))}" /></div>
          <div class="field"><label>Yield delta (t per ha)</label><input id="tr_yield_${esc(t.id)}" type="number" step="0.0001" value="${esc(String(isFiniteNumber(yDelta) ? yDelta : 0))}" /></div>
        </div>
        <div class="field">
          <label>Notes</label>
          <textarea id="tr_notes_${esc(t.id)}" rows="2">${esc(t.notes || "")}</textarea>
        </div>
      `;
      root.appendChild(el);

      safeOnInput(`tr_name_${t.id}`, e => {
        t.name = e.target.value;
        calcAndRenderAllSafe();
        autosaveToLocalStorage();
      });
      safeOnInput(`tr_area_${t.id}`, e => {
        t.area = parseNumber(e.target.value) || 0;
        calcAndRenderAllSafe();
        autosaveToLocalStorage();
      });

      const ctrlSel = document.getElementById(`tr_isControl_${t.id}`);
      if (ctrlSel) {
        ctrlSel.addEventListener("change", e => {
          const v = e.target.value === "control";
          if (v) {
            for (const tt of model.treatments) tt.isControl = false;
            t.isControl = true;
          } else {
            t.isControl = false;
          }
          calcAndRenderAllSafe();
          renderTreatments();
          renderRecurrenceConfig();
          autosaveToLocalStorage();
          showToast("Control treatment updated.");
        });
      }

      const updateTotal = () => {
        const matsEl = document.getElementById(`tr_mats_${t.id}`);
        const servEl = document.getElementById(`tr_serv_${t.id}`);
        const labEl = document.getElementById(`tr_lab_${t.id}`);
        const totalEl = document.getElementById(`tr_total_${t.id}`);
        if (matsEl && servEl && labEl && totalEl) {
          const mm = parseNumber(matsEl.value) || 0;
          const ss = parseNumber(servEl.value) || 0;
          const ll = parseNumber(labEl.value) || 0;
          totalEl.value = String(mm + ss + ll);
        }
      };

      safeOnInput(`tr_mats_${t.id}`, e => {
        t.materialsCost = parseNumber(e.target.value) || 0;
        updateTotal();
        calcAndRenderAllSafe();
        autosaveToLocalStorage();
      });
      safeOnInput(`tr_serv_${t.id}`, e => {
        t.servicesCost = parseNumber(e.target.value) || 0;
        updateTotal();
        calcAndRenderAllSafe();
        autosaveToLocalStorage();
      });
      safeOnInput(`tr_lab_${t.id}`, e => {
        t.labourCost = parseNumber(e.target.value) || 0;
        updateTotal();
        calcAndRenderAllSafe();
        autosaveToLocalStorage();
      });
      safeOnInput(`tr_cap_${t.id}`, e => {
        t.capitalCost = parseNumber(e.target.value) || 0;
        calcAndRenderAllSafe();
        autosaveToLocalStorage();
      });
      safeOnInput(`tr_yield_${t.id}`, e => {
        const yOut2 = model.outputs.find(o => o.name.toLowerCase().includes("yield")) || model.outputs[0];
        if (yOut2) t.deltas[yOut2.id] = parseNumber(e.target.value) || 0;
        calcAndRenderAllSafe();
        autosaveToLocalStorage();
        showToast("Yield delta updated.");
      });
      safeOnInput(`tr_notes_${t.id}`, e => {
        t.notes = e.target.value;
        autosaveToLocalStorage();
      });
    }
  }

  function renderAllSafe() {
    renderOutputs();
    renderTreatments();
    renderRecurrenceConfig();
    renderScenarioList();
  }

  // =========================
  // 16) CALC + RENDER PIPELINE
  // =========================
  let calcDebounceTimer = null;

  function calcAndRenderAllSafeDebounced() {
    clearTimeout(calcDebounceTimer);
    calcDebounceTimer = setTimeout(calcAndRenderAllSafe, 120);
  }

  function calcAndRenderAllSafe() {
    // Base case per treatment
    const base = computePerTreatmentResultsBaseCase();
    const filtered = applyResultsFilter(base);

    state.computed.perTreatment = base;
    state.computed.comparisonGrid = computeComparisonGrid(filtered);

    // Render standard headline metrics if present (use control vs best treatment)
    const control = base.find(x => x.isControl) || null;
    const best = base.filter(x => !x.isControl).slice().sort((a, b) => (b.npv - a.npv))[0] || null;

    if (best) {
      safeSetText("pvBenefits", money(best.pvBenefits));
      safeSetText("pvCosts", money(best.pvCosts));
      safeSetText("npv", money(best.npv));
      safeSetText("bcr", isFiniteNumber(best.bcr) ? fmt(best.bcr) : "n/a");
      safeSetText("roi", isFiniteNumber(best.roi) ? percent(best.roi) : "n/a");
      safeSetText("irr", isFiniteNumber(best.irrPct) ? percent(best.irrPct) : "n/a");
      safeSetText("payback", best.paybackYears == null ? "Not reached" : String(best.paybackYears));
    }

    if (control) {
      safeSetText("pvBenefitsControl", money(control.pvBenefits));
      safeSetText("pvCostsControl", money(control.pvCosts));
      safeSetText("npvControl", money(control.npv));
      safeSetText("bcrControl", isFiniteNumber(control.bcr) ? fmt(control.bcr) : "n/a");
      safeSetText("roiControl", isFiniteNumber(control.roi) ? percent(control.roi) : "n/a");
      safeSetText("irrControl", isFiniteNumber(control.irrPct) ? percent(control.irrPct) : "n/a");
      safeSetText("paybackControl", control.paybackYears == null ? "Not reached" : String(control.paybackYears));
    }

    renderLeaderboard(filtered);
    renderComparisonGrid(state.computed.comparisonGrid);
    renderWhatThisMeans(base);
    renderAiBriefing();

    autosaveToLocalStorage();
  }

  // =========================
  // 17) TABS (bind via an ID container if present)
  // =========================
  function initTabsIdOnly() {
    const nav = document.getElementById("tabNav");
    if (!nav) return;

    const panels = Array.from(document.querySelectorAll(".tab-panel"));

    const showTab = key => {
      // nav buttons expected to have data-tab-target, but binding stays on tabNav id
      const buttons = Array.from(nav.querySelectorAll("[data-tab],[data-tab-target],[data-tab-jump]"));
      buttons.forEach(b => {
        const k = b.getAttribute("data-tab") || b.getAttribute("data-tab-target") || b.getAttribute("data-tab-jump");
        const active = k === key;
        b.classList.toggle("active", active);
        b.setAttribute("aria-selected", active ? "true" : "false");
      });

      panels.forEach(p => {
        const k = p.getAttribute("data-tab-panel") || (p.id ? p.id.replace(/^tab-/, "") : "");
        const match = k === key || p.id === key || p.id === "tab-" + key;
        p.classList.toggle("active", !!match);
        p.hidden = !match;
        p.setAttribute("aria-hidden", match ? "false" : "true");
        p.style.display = match ? "" : "none";
      });

      try { localStorage.setItem(LS_LAST_ACTIVE_TAB_KEY, key); } catch {}
    };

    nav.addEventListener("click", e => {
      const target = e.target.closest("[data-tab],[data-tab-target],[data-tab-jump]");
      if (!target) return;
      const key = target.getAttribute("data-tab") || target.getAttribute("data-tab-target") || target.getAttribute("data-tab-jump");
      if (!key) return;
      e.preventDefault();
      e.stopPropagation();
      showTab(key);
      showToast(`Switched to ${key} tab.`);
    });

    const saved = (() => {
      try { return localStorage.getItem(LS_LAST_ACTIVE_TAB_KEY) || ""; } catch { return ""; }
    })();

    if (saved) showTab(saved);
    else {
      const first = nav.querySelector("[data-tab],[data-tab-target],[data-tab-jump]");
      if (first) {
        const key = first.getAttribute("data-tab") || first.getAttribute("data-tab-target") || first.getAttribute("data-tab-jump");
        if (key) showTab(key);
      }
    }
  }

  // =========================
  // 18) BINDINGS (ID-ONLY)
  // =========================
  function bindProjectFields() {
    safeSetValue("projectName", model.project.name || "");
    safeSetValue("projectLead", model.project.lead || "");
    safeSetValue("analystNames", model.project.analysts || "");
    safeSetValue("projectTeam", model.project.team || "");
    safeSetValue("projectSummary", model.project.summary || "");
    safeSetValue("projectObjectives", model.project.objectives || "");
    safeSetValue("projectActivities", model.project.activities || "");
    safeSetValue("stakeholderGroups", model.project.stakeholders || "");
    safeSetValue("lastUpdated", model.project.lastUpdated || "");
    safeSetValue("projectGoal", model.project.goal || "");
    safeSetValue("withProject", model.project.withProject || "");
    safeSetValue("withoutProject", model.project.withoutProject || "");
    safeSetValue("organisation", model.project.organisation || "");
    safeSetValue("contactEmail", model.project.contactEmail || "");
    safeSetValue("contactPhone", model.project.contactPhone || "");

    safeSetValue("startYear", model.time.startYear);
    safeSetValue("years", model.time.years);
    safeSetValue("discBase", model.time.discBase);
    safeSetValue("discLow", model.time.discLow);
    safeSetValue("discHigh", model.time.discHigh);

    // Inputs
    safeOnInput("projectName", e => { model.project.name = e.target.value; autosaveToLocalStorage(); calcAndRenderAllSafeDebounced(); });
    safeOnInput("projectLead", e => { model.project.lead = e.target.value; autosaveToLocalStorage(); });
    safeOnInput("analystNames", e => { model.project.analysts = e.target.value; autosaveToLocalStorage(); });
    safeOnInput("projectTeam", e => { model.project.team = e.target.value; autosaveToLocalStorage(); });
    safeOnInput("projectSummary", e => { model.project.summary = e.target.value; autosaveToLocalStorage(); });
    safeOnInput("projectObjectives", e => { model.project.objectives = e.target.value; autosaveToLocalStorage(); });
    safeOnInput("projectActivities", e => { model.project.activities = e.target.value; autosaveToLocalStorage(); });
    safeOnInput("stakeholderGroups", e => { model.project.stakeholders = e.target.value; autosaveToLocalStorage(); });
    safeOnInput("lastUpdated", e => { model.project.lastUpdated = e.target.value; autosaveToLocalStorage(); });
    safeOnInput("projectGoal", e => { model.project.goal = e.target.value; autosaveToLocalStorage(); });
    safeOnInput("withProject", e => { model.project.withProject = e.target.value; autosaveToLocalStorage(); });
    safeOnInput("withoutProject", e => { model.project.withoutProject = e.target.value; autosaveToLocalStorage(); });
    safeOnInput("organisation", e => { model.project.organisation = e.target.value; autosaveToLocalStorage(); });
    safeOnInput("contactEmail", e => { model.project.contactEmail = e.target.value; autosaveToLocalStorage(); });
    safeOnInput("contactPhone", e => { model.project.contactPhone = e.target.value; autosaveToLocalStorage(); });

    safeOnInput("startYear", e => { model.time.startYear = Math.floor(parseNumber(e.target.value) || model.time.startYear); autosaveToLocalStorage(); calcAndRenderAllSafeDebounced(); });
    safeOnInput("years", e => { model.time.years = Math.max(1, Math.floor(parseNumber(e.target.value) || model.time.years)); autosaveToLocalStorage(); calcAndRenderAllSafeDebounced(); });
    safeOnInput("discBase", e => { model.time.discBase = parseNumber(e.target.value) || model.time.discBase; autosaveToLocalStorage(); calcAndRenderAllSafeDebounced(); });
    safeOnInput("discLow", e => { model.time.discLow = parseNumber(e.target.value) || model.time.discLow; autosaveToLocalStorage(); });
    safeOnInput("discHigh", e => { model.time.discHigh = parseNumber(e.target.value) || model.time.discHigh; autosaveToLocalStorage(); });

    safeOnClick("recalc", () => { calcAndRenderAllSafe(); showToast("Results recalculated."); });
    safeOnClick("getResults", () => { calcAndRenderAllSafe(); showToast("Results updated."); });

    // Import buttons (existing IDs)
    safeOnClick("parseExcel", () => { handleUploadParse(); });
    safeOnClick("importExcel", () => { handleCommitImportStage(); });

    // Paste pipeline (optional IDs)
    safeOnClick("parsePaste", () => { handleParsePaste(); });
    safeOnClick("commitPaste", () => { handleCommitPaste(); });

    // Exports (existing and optional IDs)
    safeOnClick("exportCsv", () => { exportTreatmentSummaryCSV(); });
    safeOnClick("exportCsvFoot", () => { exportTreatmentSummaryCSV(); });

    safeOnClick("exportCleanTsv", () => { exportCleanedDatasetTSV(); });
    safeOnClick("exportTreatSummary", () => { exportTreatmentSummaryCSV(); });
    safeOnClick("exportSensitivity", () => { exportSensitivityGridCSV(); });
    safeOnClick("exportWorkbook", () => { exportWorkbookIfFeasible(); });

    // PDF print buttons if present
    safeOnClick("exportPdf", () => { window.print(); showToast("Print dialog opened."); });
    safeOnClick("exportPdfFoot", () => { window.print(); showToast("Print dialog opened."); });

    // Sensitivity run if present
    safeOnClick("runSensitivity", () => {
      const grid = computeSensitivityGrid();
      state.computed.sensitivityGrid = grid;
      renderSensitivityGrid(grid);
      renderAiBriefing();
      autosaveToLocalStorage();
      showToast("Sensitivity grid computed.");
    });

    // AI briefing copy buttons (optional IDs)
    safeOnClick("copyAiPrompt", () => {
      const prompt = buildAiNarrativePrompt();
      copyToClipboard(prompt, "AI briefing prompt copied.", "Copy failed. Please copy from the prompt box.");
    });

    safeOnClick("copyResultsJson", () => {
      const json = JSON.stringify(buildResultsJsonObject(), null, 2);
      copyToClipboard(json, "Results JSON copied.", "Copy failed. Please copy from the JSON box.");
    });

    // Scenarios (optional IDs)
    safeOnClick("scenarioSave", () => {
      const nm = safeGetValue("scenarioName");
      saveScenario(nm);
    });

    safeOnClick("scenarioLoad", () => {
      const sel = document.getElementById("scenarioSelect");
      const nm = sel ? sel.value : safeGetValue("scenarioName");
      loadScenario(nm);
    });

    safeOnClick("scenarioDelete", () => {
      const sel = document.getElementById("scenarioSelect");
      const nm = sel ? sel.value : safeGetValue("scenarioName");
      deleteScenario(nm);
    });

    // Scenario select immediate load (optional)
    const scenarioSelect = document.getElementById("scenarioSelect");
    if (scenarioSelect) {
      scenarioSelect.addEventListener("change", e => {
        const nm = e.target.value;
        if (nm) {
          safeSetValue("scenarioName", nm);
          showToast("Scenario selected.");
        }
      });
    }
  }

  // =========================
  // 19) DEFAULT DATA (committed through pipeline)
  // =========================
  function buildDefaultRawTSV() {
    // Minimal default dataset consistent with the existing raw plot structure
    const header = ["Amendment", "Replicate", "Plot", "Yield t/ha", "Pre sowing Labour", "Treatment Input Cost Only /Ha"];
    const rows = [
      ["control", "1", "C1", "2.4", "40", "0"],
      ["deep_om_cp1", "1", "T1", "3.1", "55", "16500"],
      ["deep_om_cp1_plus_liq_gypsum_cht", "1", "T2", "3.2", "56", "16850"],
      ["deep_gypsum", "1", "T3", "2.9", "50", "500"],
      ["deep_om_cp1_plus_pam", "1", "T4", "3.0", "57", "18000"],
      ["deep_om_cp1_plus_ccm", "1", "T5", "3.25", "58", "21225"],
      ["deep_ccm_only", "1", "T6", "2.95", "52", "3225"],
      ["deep_om_cp2_plus_gypsum", "1", "T7", "3.3", "60", "24000"],
      ["deep_liq_gypsum_cht", "1", "T8", "2.8", "48", "350"],
      ["surface_silicon", "1", "T9", "2.7", "45", "1000"],
      ["deep_liq_npks", "1", "T10", "3.0", "53", "2200"],
      ["deep_ripping_only", "1", "T11", "2.85", "47", "0"]
    ];
    return header.join("\t") + "\n" + rows.map(r => r.join("\t")).join("\n");
  }

  function commitDefaultIfNoAutosave() {
    if (state.dataset.committed) return;

    const defaultData = buildDefaultRawTSV();
    const dict = [
      "variable,label,type,unit",
      "Amendment,Treatment label,string,",
      "Replicate,Replicate identifier,string,",
      "Plot,Plot identifier,string,",
      "\"Yield t/ha\",Yield in tonnes per hectare,number,t/ha",
      "\"Pre sowing Labour\",Pre sowing labour cost per hectare,number,$/ha",
      "\"Treatment Input Cost Only /Ha\",Treatment input cost per hectare,number,$/ha"
    ].join("\n");

    try {
      parseDataAndDictionary(defaultData, dict, "Embedded default dataset");
      handleCommitImportStage();
      showToast("Default dataset loaded.");
    } catch {
      // If parsing fails, keep model as-is
    }
  }

  // =========================
  // 20) INITIALISE
  // =========================
  document.addEventListener("DOMContentLoaded", () => {
    // Attempt restore autosave first
    const restored = restoreAutosaveFromLocalStorage();

    // Tabs: bind through a known nav container ID only
    initTabsIdOnly();

    // Base bindings
    bindProjectFields();
    bindResultsFilters();

    // If no autosave restored and no committed dataset, commit default via pipeline
    if (!restored) {
      commitDefaultIfNoAutosave();
    }

    // Render baseline UI
    renderAllSafe();
    renderDataChecks(state.dataset.checks || []);
    calcAndRenderAllSafe();

    // Sensitivity preview (only if user runs it)
    const sensRoot = document.getElementById("sensitivityGrid");
    if (sensRoot) {
      const note = document.createElement("p");
      note.className = "small muted";
      note.textContent = "Run the sensitivity grid to populate this table, then export the full grid if needed.";
      sensRoot.innerHTML = "";
      sensRoot.appendChild(note);
    }

    renderImportStatus(state.dataset.committed ? "A dataset is committed. You can re-import at any time." : "Import a dataset to begin.");
    renderScenarioList();
  });
})();
