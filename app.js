/* app.js
   Farming CBA Decision Tool 2 mesfin Newcastle Business School
   Upgrades implemented:
   - Unmistakable “Control vs all treatments” comparison table (sticky Indicator + Control columns, sticky header, horizontal scroll)
   - Leaderboard above table (click to focus/highlight a treatment column)
   - Excel-first “schema as contract” parsing + validation report + strict fail/clear errors + calibration summary
   - Auditable calculations: per-indicator “Calculation details” drawer + reconciliation checks
   - Treatments tab cost build-up (Capital first; operating line-items; computed totals drive all downstream calculations)
   - Simulations tab: one-way sensitivity + break-even tools + scenario sets (saved)
   - AI tab: structured copy-paste prompt (JSON + instructions), exportable
   - Exports: Excel (Results/Inputs/Assumptions/Simulations/Audit/AI), Print-ready PDF via browser
   - Governance: tool version + audit trail + reset + restore last successful upload
*/

(() => {
  "use strict";

  // ----------------------------
  // Tool identity (embedded in exports, audit, AI packs)
  // ----------------------------
  const TOOL = Object.freeze({
    name: "Farming CBA Decision Tool 2",
    organisation: "Newcastle Business School, The University of Newcastle",
    version: "2.0.0 (2026-01-03)"
  });

  // ----------------------------
  // Utilities
  // ----------------------------
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const uid = () => Math.random().toString(36).slice(2, 10);

  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

  const esc = s =>
    (s ?? "")
      .toString()
      .replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const isFiniteNum = n => typeof n === "number" && Number.isFinite(n);

  const fmt = n =>
    isFiniteNum(n)
      ? Math.abs(n) >= 1000
        ? n.toLocaleString(undefined, { maximumFractionDigits: 0 })
        : n.toLocaleString(undefined, { maximumFractionDigits: 2 })
      : "n/a";

  const money = n => (isFiniteNum(n) ? "$" + fmt(n) : "n/a");

  const ratio = n => (isFiniteNum(n) ? fmt(n) : "n/a");

  const percent = n => (isFiniteNum(n) ? fmt(n) + "%" : "n/a");

  const nowISO = () => new Date().toISOString();

  function showToast(message) {
    const root = $("#toast-root") || document.body;
    const toast = document.createElement("div");
    toast.className = "toast";
    toast.textContent = message;
    root.appendChild(toast);
    void toast.offsetWidth;
    toast.classList.add("show");
    setTimeout(() => {
      toast.classList.remove("show");
      setTimeout(() => toast.remove(), 180);
    }, 3200);
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

  function parseNumberLoose(value) {
    if (value === null || value === undefined || value === "") return NaN;
    if (typeof value === "number") return Number.isFinite(value) ? value : NaN;

    const s0 = String(value).trim();
    if (!s0) return NaN;

    // Percent values like "7%" or "0.07"
    const isPct = /%$/.test(s0);
    const cleaned = s0
      .replace(/\s+/g, "")
      .replace(/[\$,]/g, "")
      .replace(/^\((.*)\)$/, "-$1"); // (123) => -123

    const n = parseFloat(cleaned);
    if (!Number.isFinite(n)) return NaN;

    return isPct ? n / 100 : n;
  }

  function normaliseKey(k) {
    return String(k || "")
      .toLowerCase()
      .trim()
      .replace(/\s+/g, " ")
      .replace(/[_\-]+/g, " ")
      .replace(/[^\w\s/%().]+/g, "");
  }

  function slugify(s) {
    return (s || "project")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_|_$/g, "");
  }

  // Discount factors cache: key = `${rate}|${N}`
  const PV_CACHE = new Map();
  function discountFactors(ratePct, N) {
    const key = `${ratePct}|${N}`;
    if (PV_CACHE.has(key)) return PV_CACHE.get(key);
    const r = ratePct / 100;
    const df = new Array(N + 1);
    for (let t = 0; t <= N; t++) df[t] = 1 / Math.pow(1 + r, t);
    PV_CACHE.set(key, df);
    return df;
  }

  function presentValue(series, ratePct) {
    const N = series.length - 1;
    const df = discountFactors(ratePct, N);
    let pv = 0;
    for (let t = 0; t <= N; t++) pv += (Number(series[t]) || 0) * df[t];
    return pv;
  }

  // ----------------------------
  // Default dataset (used as baseline + “Reset to default dataset”)
  // ----------------------------
  const DEFAULT_SHEET_CANDIDATES = ["FabaBeanRaw", "FabaBeansRaw", "FabaBean", "FabaBeans", "Data", "Sheet1"];

  // A minimal default dataset; tool accepts richer uploads with many extra columns.
  const DEFAULT_RAW_ROWS = [
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

  // ----------------------------
  // Model
  // ----------------------------
  const model = {
    tool: { ...TOOL },
    project: {
      name: "Faba bean soil amendment trial",
      lead: "Project lead",
      analysts: "Farm economics team",
      team: "Trial team",
      organisation: TOOL.organisation,
      contactEmail: "",
      contactPhone: "",
      summary:
        "Applied faba bean trial comparing soil amendment and ripping treatments against a control.",
      objectives:
        "Quantify yield and gross margin impacts of alternative soil amendment strategies.",
      activities:
        "Establish replicated plots, collect yield and cost data, and summarise economics.",
      stakeholders: "Producers, agronomists, government agencies, research partners.",
      lastUpdated: new Date().toISOString().slice(0, 10)
    },
    time: {
      startYear: new Date().getFullYear(),
      years: 10,
      discBase: 7,
      discLow: 4,
      discHigh: 10
    },
    adoption: { base: 1.0, low: 0.6, high: 1.0 },
    risk: { base: 0.15, low: 0.05, high: 0.3 },
    // Outputs used to compute benefits: benefit_per_ha = Σ(delta_output * value_per_unit)
    outputs: [
      { id: "out_yield", name: "Grain yield uplift", unit: "t/ha", value: 450, source: "Input Directly" }
    ],
    outputsMeta: {
      assumptions:
        "Yield value uses grain price ($/t). Add other outputs (quality, soil services) as needed by assigning values."
    },
    // Each treatment is an *alternative* compared to control (incremental vs control)
    // Cost build-up drives all calculations; do not duplicate totals elsewhere.
    treatments: [],
    // Optional shared items (applied to *non-control* treatments by default)
    benefitsShared: [],
    costsShared: [],
    // Simulations
    sim: {
      oneWay: {
        grainPrice: 450,
        yieldDeltaMultiplier: 1.0,
        discountRate: 7,
        adoption: 1.0,
        risk: 0.15,
        operatingCostMultiplier: 1.0,
        capitalCostMultiplier: 1.0
      },
      scenarios: [], // saved scenario sets
      monteCarlo: { n: 1000, variationPct: 20, seed: null, results: null }
    },
    // Data provenance + audit
    data: {
      currentSource: { kind: "default", name: "Built-in default dataset", timestamp: nowISO() },
      lastSuccessfulUpload: null,
      rawRowsCount: 0,
      treatmentsCount: 0,
      replicateRule: "Within each Amendment/Treatment, numeric columns are averaged; deltas are computed versus control means."
    },
    audit: [] // [{time, action, details}]
  };

  // ----------------------------
  // Audit
  // ----------------------------
  const LS_LAST_SUCCESS = "farming_cba_last_success_v2";
  const LS_SCENARIOS = "farming_cba_scenarios_v2";

  function logEvent(action, details = {}) {
    model.audit.unshift({ time: nowISO(), action, details });
    renderAudit();
  }

  // ----------------------------
  // Excel schema contract + validation
  // ----------------------------
  const EXCEL_SCHEMA = Object.freeze({
    required: {
      treatment: ["amendment", "treatment", "treatment name", "treatmentname", "option", "arm"],
      yield: ["yield t/ha", "yield", "grain yield", "yield (t/ha)", "yield_t_ha", "yield tha"]
    },
    optional: {
      isControl: ["iscontrol", "control", "baseline", "control flag", "control?"],
      areaHa: ["area", "area ha", "area (ha)", "hectares", "ha"],
      adoption: ["adoption", "adoption rate", "implementation rate"],
      capital: ["capital", "capital cost", "capital cost ($)", "capex", "year 0 capital", "capital y0"]
    },
    // Classification heuristics (all numeric columns are used: cost or output; non-numeric kept as metadata and exported)
    costKeywords: [
      "cost", "labour", "labor", "herbicide", "fungicide", "insecticide", "seed", "fert",
      "fertiliser", "fertilizer", "fuel", "machinery", "chemical", "spray", "amendment",
      "input", "service", "contract", "freight", "water", "energy"
    ],
    outputKeywords: ["protein", "screen", "quality", "carbon", "soil", "emission", "runoff", "erosion"]
  });

  function findColumnByAliases(headers, aliases) {
    const normHeaders = headers.map(h => ({ raw: h, norm: normaliseKey(h) }));
    for (const a of aliases) {
      const target = normaliseKey(a);
      const hit = normHeaders.find(h => h.norm === target);
      if (hit) return hit.raw;
    }
    // fall back: “contains” match
    for (const a of aliases) {
      const target = normaliseKey(a);
      const hit = normHeaders.find(h => h.norm.includes(target));
      if (hit) return hit.raw;
    }
    return null;
  }

  function classifyNumericColumn(colName) {
    const n = normaliseKey(colName);
    if (EXCEL_SCHEMA.required.yield.map(normaliseKey).includes(n)) return "output";
    const isCost = EXCEL_SCHEMA.costKeywords.some(k => n.includes(k));
    if (isCost) return "cost";
    const isOut = EXCEL_SCHEMA.outputKeywords.some(k => n.includes(k));
    if (isOut) return "output";
    // Default: treat as output (value defaults to 0 until user sets it)
    return "output";
  }

  function guessCostCategory(colName) {
    const n = normaliseKey(colName);
    if (n.includes("labour") || n.includes("labor")) return "Labour";
    if (n.includes("service") || n.includes("contract")) return "Services";
    if (n.includes("fuel") || n.includes("machinery")) return "Services";
    return "Materials";
  }

  function buildSchemaReportHTML(report) {
    const { ok, errors, warnings, mapping, calibration } = report;

    const li = arr => arr.map(x => `<li>${esc(x)}</li>`).join("");
    const kv = obj =>
      Object.entries(obj || {})
        .map(([k, v]) => `<div class="kvrow"><div class="k">${esc(k)}</div><div class="v">${esc(v)}</div></div>`)
        .join("");

    return `
      <div class="report ${ok ? "ok" : "bad"}">
        <div class="report-head">
          <div class="badge ${ok ? "good" : "warn"}">${ok ? "Validation passed" : "Validation failed"}</div>
          <div class="small muted">${esc(TOOL.name)} • ${esc(TOOL.version)}</div>
        </div>

        <div class="report-grid">
          <div class="card">
            <h4>Schema mapping</h4>
            <div class="kv">${kv(mapping)}</div>
          </div>
          <div class="card">
            <h4>Calibration summary</h4>
            <div class="kv">${kv(calibration)}</div>
          </div>
        </div>

        ${errors.length ? `<div class="card danger"><h4>Errors (must fix)</h4><ul>${li(errors)}</ul></div>` : ""}
        ${warnings.length ? `<div class="card"><h4>Warnings (check)</h4><ul>${li(warnings)}</ul></div>` : ""}

        <div class="small muted">
          Strict reliability: the tool will only apply an upload if validation fully passes (no silent partial imports).
        </div>
      </div>
    `;
  }

  function validateAndCanonicaliseRows(jsonRows) {
    const headers = jsonRows.length ? Object.keys(jsonRows[0]) : [];
    const errors = [];
    const warnings = [];

    if (!jsonRows.length) {
      return {
        ok: false,
        report: {
          ok: false,
          errors: ["No rows found in the selected sheet."],
          warnings: [],
          mapping: {},
          calibration: {}
        },
        canonical: null
      };
    }

    const colTreatment = findColumnByAliases(headers, EXCEL_SCHEMA.required.treatment);
    const colYield = findColumnByAliases(headers, EXCEL_SCHEMA.required.yield);

    const colIsControl = findColumnByAliases(headers, EXCEL_SCHEMA.optional.isControl);
    const colArea = findColumnByAliases(headers, EXCEL_SCHEMA.optional.areaHa);
    const colAdopt = findColumnByAliases(headers, EXCEL_SCHEMA.optional.adoption);
    const colCapital = findColumnByAliases(headers, EXCEL_SCHEMA.optional.capital);

    if (!colTreatment) errors.push(`Missing required column for Treatment/Amendment. Accepted aliases include: ${EXCEL_SCHEMA.required.treatment.join(", ")}.`);
    if (!colYield) errors.push(`Missing required column for Yield (t/ha). Accepted aliases include: ${EXCEL_SCHEMA.required.yield.join(", ")}.`);

    const mapping = {
      Treatment: colTreatment || "(missing)",
      "Yield (t/ha)": colYield || "(missing)",
      "Is control? (optional)": colIsControl || "(not provided)",
      "Area ha (optional)": colArea || "(not provided)",
      "Adoption (optional)": colAdopt || "(not provided)",
      "Capital cost Y0 (optional)": colCapital || "(not provided)"
    };

    // Identify all other columns
    const usedBase = new Set([colTreatment, colYield, colIsControl, colArea, colAdopt, colCapital].filter(Boolean));
    const otherCols = headers.filter(h => !usedBase.has(h));

    // Determine which are numeric enough to be used
    const numericCols = [];
    const metaCols = [];

    for (const c of otherCols) {
      let okNum = 0;
      let seen = 0;
      for (const r of jsonRows) {
        const v = r[c];
        if (v === null || v === undefined || v === "") continue;
        seen++;
        const n = parseNumberLoose(v);
        if (!Number.isNaN(n)) okNum++;
      }
      if (seen === 0) {
        // empty column in practice
        metaCols.push(c);
        continue;
      }
      const share = okNum / Math.max(1, seen);
      if (share >= 0.6) numericCols.push(c);
      else metaCols.push(c);
      if (share < 0.9) warnings.push(`Column "${c}" has mixed values (numeric parse success ${(share * 100).toFixed(0)}%). The tool will auto-coerce common formats; check for stray text.`);
    }

    // Canonical rows: every row preserved; numeric values coerced; all numeric columns used (cost or output).
    const canonical = [];
    let badTreatment = 0;
    let badYield = 0;

    for (let i = 0; i < jsonRows.length; i++) {
      const r = jsonRows[i];

      const tName = String(r[colTreatment] ?? "").trim();
      if (!tName) badTreatment++;

      const y = parseNumberLoose(r[colYield]);
      if (Number.isNaN(y)) badYield++;

      const isControl =
        colIsControl
          ? (() => {
              const v = r[colIsControl];
              const s = String(v ?? "").toLowerCase().trim();
              if (typeof v === "boolean") return v;
              if (s === "1" || s === "true" || s === "yes" || s === "y") return true;
              if (s === "0" || s === "false" || s === "no" || s === "n") return false;
              return null;
            })()
          : null;

      const areaHa = colArea ? parseNumberLoose(r[colArea]) : NaN;
      const adoption = colAdopt ? parseNumberLoose(r[colAdopt]) : NaN;
      const capital = colCapital ? parseNumberLoose(r[colCapital]) : NaN;

      const vars = {};
      // Use ALL numeric columns
      for (const c of numericCols) {
        const n = parseNumberLoose(r[c]);
        vars[c] = Number.isNaN(n) ? null : n;
      }

      // Keep meta columns for provenance and export
      const meta = {};
      for (const c of metaCols) meta[c] = r[c] ?? null;

      canonical.push({
        __rowIndex: i + 2, // Excel-like line number if header is row 1
        treatment: tName,
        yield_t_ha: Number.isNaN(y) ? null : y,
        isControl,
        areaHa: Number.isNaN(areaHa) ? null : areaHa,
        adoption: Number.isNaN(adoption) ? null : adoption,
        capital_y0: Number.isNaN(capital) ? null : capital,
        vars,
        meta
      });
    }

    if (badTreatment > 0) errors.push(`Found ${badTreatment} row(s) with blank Treatment/Amendment values.`);
    if (badYield > 0) errors.push(`Found ${badYield} row(s) with non-numeric Yield values (after auto-coercion).`);

    // If required columns are missing or core data invalid, fail hard.
    const ok = errors.length === 0;

    const calibration = {
      "Rows read": String(canonical.length),
      "Detected numeric columns (used)": String(numericCols.length),
      "Detected meta columns (kept for provenance)": String(metaCols.length),
      "Aggregation rule": model.data.replicateRule
    };

    return {
      ok,
      report: { ok, errors, warnings, mapping, calibration },
      canonical: ok ? { rows: canonical, numericCols, metaCols } : null
    };
  }

  function aggregateToTreatments(canonical) {
    const rows = canonical.rows;

    // Identify control: priority order:
    // 1) explicit isControl==true
    // 2) treatment name contains 'control' or 'baseline'
    let controlName = null;
    const explicit = rows.find(r => r.isControl === true && r.treatment);
    if (explicit) controlName = explicit.treatment;

    if (!controlName) {
      const implied = rows.find(r => /control|baseline/i.test(r.treatment || ""));
      if (implied) controlName = implied.treatment;
    }

    // If still not found: first unique treatment becomes control (but flagged loudly)
    const uniqTreats = Array.from(new Set(rows.map(r => r.treatment).filter(Boolean)));
    if (!controlName && uniqTreats.length) controlName = uniqTreats[0];

    const groups = new Map(); // name -> {n, sums: {yield, capital, vars...}, counts...}
    const allVars = canonical.numericCols.slice(); // used numeric columns

    function initGroup(name) {
      const sums = { yield: 0, capital: 0 };
      const counts = { yield: 0, capital: 0 };
      const sumsVars = {};
      const countsVars = {};
      for (const c of allVars) {
        sumsVars[c] = 0;
        countsVars[c] = 0;
      }
      return { name, n: 0, sums, counts, sumsVars, countsVars, metaSamples: [] };
    }

    for (const r of rows) {
      const name = r.treatment;
      if (!name) continue;

      if (!groups.has(name)) groups.set(name, initGroup(name));
      const g = groups.get(name);
      g.n += 1;

      if (r.yield_t_ha !== null) {
        g.sums.yield += r.yield_t_ha;
        g.counts.yield += 1;
      }
      if (r.capital_y0 !== null) {
        g.sums.capital += r.capital_y0;
        g.counts.capital += 1;
      }

      for (const c of allVars) {
        const v = r.vars[c];
        if (v !== null && v !== undefined) {
          g.sumsVars[c] += v;
          g.countsVars[c] += 1;
        }
      }

      // Keep a few meta samples
      if (g.metaSamples.length < 3) g.metaSamples.push(r.meta || {});
    }

    const means = new Map(); // name -> {yield, capital, varsMeans...}
    for (const [name, g] of groups.entries()) {
      const m = {
        name,
        n: g.n,
        yield: g.counts.yield ? g.sums.yield / g.counts.yield : null,
        capital: g.counts.capital ? g.sums.capital / g.counts.capital : null,
        vars: {}
      };
      for (const c of allVars) {
        m.vars[c] = g.countsVars[c] ? g.sumsVars[c] / g.countsVars[c] : null;
      }
      means.set(name, m);
    }

    const controlMeans = means.get(controlName) || null;

    return {
      controlName,
      controlMeans,
      means,
      allVars
    };
  }

  // ----------------------------
  // Build tool model from aggregated upload (strictly uses ALL numeric columns)
  // ----------------------------
  function applyAggregatedDataToModel(agg, opts = {}) {
    const { controlName, controlMeans, means, allVars } = agg;
    const forceArea = opts.areaHa ?? null;

    if (!controlMeans) throw new Error("Unable to identify a control treatment (baseline).");

    // Outputs:
    // - Yield uplift always mapped to out_yield
    // - All other numeric columns: classified as cost or output
    const costCols = [];
    const outputCols = [];

    for (const c of allVars) {
      const kind = classifyNumericColumn(c);
      if (kind === "cost") costCols.push(c);
      else outputCols.push(c);
    }

    // Ensure outputs exist for outputCols (besides yield)
    const existingByName = new Map(model.outputs.map(o => [normaliseKey(o.name), o]));
    const extraOutputs = [];

    for (const c of outputCols) {
      // Avoid duplicate “yield” style columns; those should remain in yield.
      const n = normaliseKey(c);
      if (n.includes("yield")) continue;

      const outName = c;
      const key = normaliseKey(outName);
      if (existingByName.has(key)) continue;

      extraOutputs.push({
        id: "out_" + slugify(outName).slice(0, 28) + "_" + uid(),
        name: outName,
        unit: "unit",
        value: 0,
        source: "Imported (value=0 until set)"
      });
    }

    // Keep grain yield uplift output first
    model.outputs = [
      model.outputs.find(o => o.id === "out_yield") || { id: "out_yield", name: "Grain yield uplift", unit: "t/ha", value: 450, source: "Input Directly" },
      ...model.outputs.filter(o => o.id !== "out_yield"),
      ...extraOutputs
    ];

    const yieldOutput = model.outputs.find(o => o.id === "out_yield");

    // Build treatments as alternatives
    const treatments = [];

    // Control baseline: by definition incremental deltas are 0
    treatments.push({
      id: uid(),
      name: "Control (baseline)" + (controlName && !/control/i.test(controlName) ? ` — ${controlName}` : ""),
      isControl: true,
      areaHa: forceArea ?? 100,
      adoption: 1,
      capitalCostY0: 0,
      costItems: costCols.map(c => ({
        id: uid(),
        label: c,
        category: guessCostCategory(c),
        valuePerHa: 0
      })),
      deltas: Object.fromEntries(model.outputs.map(o => [o.id, 0])),
      provenance: { mean: controlMeans, controlName, appliedAt: nowISO() }
    });

    // Other treatments: delta = treatment mean - control mean (yield + all numeric vars)
    const names = Array.from(means.keys()).filter(n => n !== controlName);
    names.sort((a, b) => a.localeCompare(b));

    for (const name of names) {
      const m = means.get(name);
      const area = forceArea ?? 100;

      const deltas = Object.fromEntries(model.outputs.map(o => [o.id, 0]));
      // Yield uplift:
      if (yieldOutput) {
        const yT = m.yield ?? 0;
        const yC = controlMeans.yield ?? 0;
        deltas[yieldOutput.id] = (yT - yC) || 0;
      }

      // Other outputs:
      for (const o of model.outputs) {
        if (o.id === "out_yield") continue;
        const col = o.name;
        if (m.vars[col] === null || controlMeans.vars[col] === null) continue;
        deltas[o.id] = (Number(m.vars[col]) || 0) - (Number(controlMeans.vars[col]) || 0);
      }

      // Costs:
      const costItems = costCols.map(c => {
        const vT = m.vars[c];
        const vC = controlMeans.vars[c];
        const dv = (vT === null || vC === null) ? 0 : (Number(vT) || 0) - (Number(vC) || 0);
        return {
          id: uid(),
          label: c,
          category: guessCostCategory(c),
          valuePerHa: dv
        };
      });

      // Capital Y0 delta:
      const capT = m.capital ?? 0;
      const capC = controlMeans.capital ?? 0;
      const dCap = (capT - capC) || 0;

      treatments.push({
        id: uid(),
        name: humaniseTreatmentName(name),
        isControl: false,
        areaHa: area,
        adoption: 1,
        capitalCostY0: dCap,
        costItems,
        deltas,
        provenance: { mean: m, controlName, appliedAt: nowISO() }
      });
    }

    model.treatments = treatments;

    // Data summary
    model.data.rawRowsCount = (controlMeans?.n || 0) + names.reduce((acc, nm) => acc + (means.get(nm)?.n || 0), 0);
    model.data.treatmentsCount = model.treatments.length;

    // Keep sim defaults aligned to base settings
    model.sim.oneWay.grainPrice = model.outputs.find(o => o.id === "out_yield")?.value ?? model.sim.oneWay.grainPrice;
    model.sim.oneWay.discountRate = model.time.discBase;
    model.sim.oneWay.adoption = model.adoption.base;
    model.sim.oneWay.risk = model.risk.base;

    logEvent("Applied dataset to model", {
      control: controlName,
      treatments: model.treatments.length,
      outputs: model.outputs.length,
      costColumnsUsed: (model.treatments[0]?.costItems?.length || 0)
    });
  }

  function humaniseTreatmentName(s) {
    const t = String(s || "").trim();
    if (!t) return "Treatment";
    if (t.toLowerCase() === "control") return "Control";
    return t
      .replace(/_/g, " ")
      .replace(/\s+/g, " ")
      .replace(/\bcp(\d+)\b/gi, "CP$1")
      .replace(/\bnpks\b/gi, "NPKS")
      .replace(/\bcht\b/gi, "CHT")
      .trim();
  }

  // ----------------------------
  // Shared benefits/costs (optional; applied to non-control treatments)
  // ----------------------------
  function buildSharedBenefitsSeries(N, baseYear, adoption, risk) {
    // Kept minimal by default; user can add in UI.
    const s = new Array(N + 1).fill(0);
    for (const b of model.benefitsShared) {
      const freq = String(b.frequency || "Annual");
      const sy = Number(b.startYear) || baseYear;
      const ey = Number(b.endYear) || sy;
      const onceY = Number(b.year) || sy;
      const g = Number(b.growthPct) || 0;

      const A = b.linkAdoption ? clamp(adoption, 0, 1) : 1;
      const R = b.linkRisk ? (1 - clamp(risk, 0, 1)) : 1;

      const addAnnual = (y, amt, tFromStart) => {
        const idx = y - baseYear + 1;
        if (idx >= 1 && idx <= N) s[idx] += amt * Math.pow(1 + g / 100, tFromStart) * A * R;
      };
      const addOnce = (y, amt) => {
        const idx = y - baseYear + 1;
        if (idx >= 0 && idx <= N) s[idx] += amt * A * R;
      };

      if (freq === "Once") {
        addOnce(onceY, Number(b.amount) || 0);
        continue;
      }

      for (let y = sy; y <= ey; y++) addAnnual(y, Number(b.amount) || 0, y - sy);
    }
    return s;
  }

  function buildSharedCostsSeries(N, baseYear) {
    const s = new Array(N + 1).fill(0);
    for (const c of model.costsShared) {
      const type = String(c.type || "annual");
      if (type === "capital") {
        const y = Number(c.year) || baseYear;
        const idx = y - baseYear;
        if (idx >= 0 && idx <= N) s[idx] += Number(c.amount) || 0;
      } else {
        const sy = Number(c.startYear) || baseYear;
        const ey = Number(c.endYear) || sy;
        for (let y = sy; y <= ey; y++) {
          const idx = y - baseYear + 1;
          if (idx >= 1 && idx <= N) s[idx] += Number(c.amount) || 0;
        }
      }
    }
    return s;
  }

  // ----------------------------
  // Core calculations (auditable, series-based)
  // ----------------------------
  function computeTreatmentScenarioMetrics(treatment, settings) {
    const N = Number(settings.years);
    const baseYear = Number(settings.startYear);
    const r = Number(settings.discountRate);
    const adoption = clamp(Number(settings.adoption), 0, 1);
    const risk = clamp(Number(settings.risk), 0, 1);

    const isControl = !!treatment.isControl;

    // Benefits per ha from outputs
    let benefitPerHa = 0;
    for (const o of model.outputs) {
      const delta = Number(treatment.deltas?.[o.id]) || 0;
      const val = Number(o.value) || 0;
      benefitPerHa += delta * val;
    }

    // Apply one-way multipliers if provided
    benefitPerHa *= Number(settings.yieldDeltaMultiplier ?? 1);

    // Operating costs per ha
    const operatingPerHaRaw = (treatment.costItems || []).reduce((acc, it) => acc + (Number(it.valuePerHa) || 0), 0);
    const operatingPerHa = operatingPerHaRaw * Number(settings.operatingCostMultiplier ?? 1);

    // Capital cost Y0
    const capY0 = (Number(treatment.capitalCostY0) || 0) * Number(settings.capitalCostMultiplier ?? 1);

    const area = Number(treatment.areaHa) || 0;

    // Incremental analysis convention:
    // - Control baseline: treat incremental series as 0 (benefits=0, costs=0)
    // - Treatment: use deltas already defined vs control; apply adoption & risk
    const benefitByYear = new Array(N + 1).fill(0);
    const costByYear = new Array(N + 1).fill(0);

    if (!isControl) {
      const annualBenefit = benefitPerHa * area * adoption * (1 - risk);
      const annualCost = operatingPerHa * area * adoption;

      costByYear[0] += capY0;
      for (let y = 1; y <= N; y++) {
        benefitByYear[y] += annualBenefit;
        costByYear[y] += annualCost;
      }

      // Shared items apply to any non-control treatment
      const sharedB = buildSharedBenefitsSeries(N, baseYear, adoption, risk);
      const sharedC = buildSharedCostsSeries(N, baseYear);
      for (let i = 0; i <= N; i++) {
        benefitByYear[i] += sharedB[i] || 0;
        costByYear[i] += sharedC[i] || 0;
      }
    }

    const pvBenefits = presentValue(benefitByYear, r);
    const pvCosts = presentValue(costByYear, r);
    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts > 0 ? pvBenefits / pvCosts : NaN;
    const roi = pvCosts > 0 ? npv / pvCosts : NaN; // ratio

    // Reconciliation checks
    const checks = {
      pvBenefits: { lhs: pvBenefits, rhs: presentValue(benefitByYear, r) },
      pvCosts: { lhs: pvCosts, rhs: presentValue(costByYear, r) },
      npv: { lhs: npv, rhs: pvBenefits - pvCosts },
      bcr: { lhs: bcr, rhs: pvCosts > 0 ? pvBenefits / pvCosts : NaN },
      roi: { lhs: roi, rhs: pvCosts > 0 ? (pvBenefits - pvCosts) / pvCosts : NaN }
    };

    return {
      pvBenefits,
      pvCosts,
      npv,
      bcr,
      roi,
      series: { benefitByYear, costByYear },
      components: {
        benefitPerHa,
        operatingPerHa,
        capitalY0: capY0,
        areaHa: area,
        adoption,
        risk
      },
      checks
    };
  }

  function computeComparisonDataset(settings, filterMode = "all") {
    const treatments = model.treatments.slice();
    const control = treatments.find(t => t.isControl) || treatments[0];

    const metrics = treatments.map(t => ({
      t,
      m: computeTreatmentScenarioMetrics(t, settings)
    }));

    const controlEntry = metrics.find(x => x.t.id === control.id) || metrics[0];
    const c = controlEntry.m;

    // Rank by NPV (excluding control)
    const nonControl = metrics.filter(x => !x.t.isControl);
    const byNpv = nonControl
      .slice()
      .sort((a, b) => (b.m.npv ?? -Infinity) - (a.m.npv ?? -Infinity))
      .map((x, i) => ({ id: x.t.id, rank: i + 1 }));
    const npvRankMap = new Map(byNpv.map(x => [x.id, x.rank]));

    // Rank by BCR (excluding control)
    const byBcr = nonControl
      .slice()
      .sort((a, b) => (Number.isFinite(b.m.bcr) ? b.m.bcr : -Infinity) - (Number.isFinite(a.m.bcr) ? a.m.bcr : -Infinity))
      .map((x, i) => ({ id: x.t.id, rank: i + 1 }));
    const bcrRankMap = new Map(byBcr.map(x => [x.id, x.rank]));

    // Build derived deltas vs control
    const rows = metrics.map(x => {
      const m = x.m;
      const d = {
        pvBenefits: m.pvBenefits - c.pvBenefits,
        pvCosts: m.pvCosts - c.pvCosts,
        npv: m.npv - c.npv,
        bcr: (Number.isFinite(m.bcr) && Number.isFinite(c.bcr)) ? (m.bcr - c.bcr) : NaN,
        roi: (Number.isFinite(m.roi) && Number.isFinite(c.roi)) ? (m.roi - c.roi) : NaN
      };
      return {
        id: x.t.id,
        name: x.t.name,
        isControl: x.t.isControl,
        metrics: m,
        delta: d,
        rankByNpv: x.t.isControl ? 0 : (npvRankMap.get(x.t.id) || null),
        rankByBcr: x.t.isControl ? 0 : (bcrRankMap.get(x.t.id) || null)
      };
    });

    // Filtering (always keep control)
    let kept = rows.slice();
    if (filterMode === "top5_npv") {
      const top = kept
        .filter(r => !r.isControl)
        .sort((a, b) => (b.metrics.npv ?? -Infinity) - (a.metrics.npv ?? -Infinity))
        .slice(0, 5)
        .map(r => r.id);
      kept = kept.filter(r => r.isControl || top.includes(r.id));
    } else if (filterMode === "top5_bcr") {
      const top = kept
        .filter(r => !r.isControl)
        .sort((a, b) => (Number.isFinite(b.metrics.bcr) ? b.metrics.bcr : -Infinity) - (Number.isFinite(a.metrics.bcr) ? a.metrics.bcr : -Infinity))
        .slice(0, 5)
        .map(r => r.id);
      kept = kept.filter(r => r.isControl || top.includes(r.id));
    } else if (filterMode === "improve_only") {
      kept = kept.filter(r => r.isControl || (r.delta.npv > 0));
    }

    // Keep control first
    kept.sort((a, b) => (a.isControl === b.isControl) ? 0 : (a.isControl ? -1 : 1));

    return { rows: kept, control: rows.find(r => r.isControl) || rows[0] };
  }

  // ----------------------------
  // Results UI: leaderboard + comparison table + calc drawer
  // ----------------------------
  let resultsFilterMode = "all";
  let focusedTreatmentId = null;

  const INDICATORS = [
    {
      key: "pvBenefits",
      label: "PV Benefits",
      unit: "$",
      format: money,
      formula: {
        english: "PV Benefits equals the discounted sum of annual benefit amounts over the analysis horizon.",
        algebra: "PV(B) = Σₜ Bₜ / (1 + r)ᵗ"
      }
    },
    {
      key: "pvCosts",
      label: "PV Costs",
      unit: "$",
      format: money,
      formula: {
        english: "PV Costs equals year-0 capital plus the discounted sum of annual operating costs over the horizon (plus any shared costs).",
        algebra: "PV(C) = C₀ + Σₜ Cₜ / (1 + r)ᵗ"
      }
    },
    {
      key: "npv",
      label: "NPV",
      unit: "$",
      format: money,
      formula: {
        english: "NPV equals PV Benefits minus PV Costs. Positive values mean net gain compared with control.",
        algebra: "NPV = PV(B) − PV(C)"
      }
    },
    {
      key: "bcr",
      label: "BCR",
      unit: "",
      format: ratio,
      formula: {
        english: "BCR equals PV Benefits divided by PV Costs.",
        algebra: "BCR = PV(B) / PV(C)"
      }
    },
    {
      key: "roi",
      label: "ROI",
      unit: "",
      format: n => (isFiniteNum(n) ? fmt(n) : "n/a"),
      formula: {
        english: "ROI equals NPV divided by PV Costs (net gain per dollar of PV cost).",
        algebra: "ROI = NPV / PV(C)"
      }
    },
    {
      key: "rank",
      label: "Rank (by NPV)",
      unit: "",
      format: n => (n ? String(n) : "—"),
      formula: {
        english: "Rank orders treatments by NPV (highest NPV is rank 1). Control is baseline.",
        algebra: "Rank = order_desc(NPV)"
      }
    }
  ];

  function renderResults() {
    const settings = currentScenarioSettings();
    const ds = computeComparisonDataset(settings, resultsFilterMode);

    renderLeaderboard(ds);
    renderComparisonTable(ds);
    renderResultsMeta(ds);
  }

  function renderResultsMeta(ds) {
    const el = $("#resultsMeta");
    if (!el) return;

    const s = currentScenarioSettings();
    el.innerHTML = `
      <div class="meta-strip">
        <div class="pill"><strong>${esc(TOOL.name)}</strong> <span class="muted">v${esc(TOOL.version)}</span></div>
        <div class="pill">Years: <strong>${esc(String(s.years))}</strong></div>
        <div class="pill">Discount: <strong>${esc(String(s.discountRate))}%</strong></div>
        <div class="pill">Adoption: <strong>${esc((s.adoption * 100).toFixed(0))}%</strong></div>
        <div class="pill">Risk: <strong>${esc((s.risk * 100).toFixed(0))}%</strong></div>
        <div class="pill">Filter: <strong>${esc(filterLabel(resultsFilterMode))}</strong></div>
      </div>
    `;
  }

  function filterLabel(mode) {
    if (mode === "top5_npv") return "Top 5 by NPV";
    if (mode === "top5_bcr") return "Top 5 by BCR";
    if (mode === "improve_only") return "Only improvements vs control";
    return "Show all";
  }

  function renderLeaderboard(ds) {
    const root = $("#resultsLeaderboard");
    if (!root) return;

    const rows = ds.rows.filter(r => !r.isControl);

    // Rank uses NPV by default (as requested). Show ΔNPV vs Control, BCR, PV Cost, PV Benefit.
    const sorted = rows.slice().sort((a, b) => (b.metrics.npv ?? -Infinity) - (a.metrics.npv ?? -Infinity));

    // Default focus: best NPV if nothing focused
    if (!focusedTreatmentId && sorted.length) focusedTreatmentId = sorted[0].id;

    root.innerHTML = `
      <div class="leaderboard-head">
        <h3>Leaderboard (click a treatment to focus)</h3>
        <div class="small muted">One row per treatment • Δ values are versus Control (baseline)</div>
      </div>

      <div class="leaderboard-wrap" role="region" aria-label="Leaderboard">
        <table class="leaderboard">
          <thead>
            <tr>
              <th>Rank</th>
              <th>Treatment</th>
              <th>ΔNPV vs Control</th>
              <th>BCR</th>
              <th>PV Cost</th>
              <th>PV Benefit</th>
            </tr>
          </thead>
          <tbody>
            ${sorted
              .map((r, i) => {
                const isActive = r.id === focusedTreatmentId;
                const dnpv = r.delta.npv;
                const dClass = dnpv > 0 ? "pos" : dnpv < 0 ? "neg" : "";
                return `
                  <tr class="${isActive ? "active" : ""}" data-focus-treatment="${esc(r.id)}" tabindex="0">
                    <td>${i + 1}</td>
                    <td><strong>${esc(r.name)}</strong></td>
                    <td class="${dClass}">${money(dnpv)}</td>
                    <td>${Number.isFinite(r.metrics.bcr) ? fmt(r.metrics.bcr) : "n/a"}</td>
                    <td>${money(r.metrics.pvCosts)}</td>
                    <td>${money(r.metrics.pvBenefits)}</td>
                  </tr>
                `;
              })
              .join("")}
          </tbody>
        </table>
      </div>
    `;

    // Click/focus handlers
    root.onclick = e => {
      const tr = e.target.closest("[data-focus-treatment]");
      if (!tr) return;
      focusedTreatmentId = tr.dataset.focusTreatment;
      renderResults();
      focusColumnIntoView(focusedTreatmentId);
      logEvent("Focused treatment in results", { treatmentId: focusedTreatmentId });
    };

    root.onkeydown = e => {
      if (e.key !== "Enter" && e.key !== " ") return;
      const tr = e.target.closest("[data-focus-treatment]");
      if (!tr) return;
      e.preventDefault();
      focusedTreatmentId = tr.dataset.focusTreatment;
      renderResults();
      focusColumnIntoView(focusedTreatmentId);
    };
  }

  function pctDelta(delta, base) {
    if (!isFiniteNum(delta) || !isFiniteNum(base) || Math.abs(base) < 1e-9) return null;
    return (delta / base) * 100;
  }

  function formatDeltaMoney(delta, base) {
    const p = pctDelta(delta, base);
    return `
      <div class="d-abs">${money(delta)}</div>
      <div class="d-pct muted">${p === null ? "—" : (p >= 0 ? "+" : "") + fmt(p) + "%"}</div>
    `;
  }

  function formatDeltaRatio(delta, base) {
    // ratio deltas: show absolute; percent usually not meaningful; still compute if base != 0
    const p = pctDelta(delta, base);
    return `
      <div class="d-abs">${isFiniteNum(delta) ? (delta >= 0 ? "+" : "") + fmt(delta) : "n/a"}</div>
      <div class="d-pct muted">${p === null ? "—" : (p >= 0 ? "+" : "") + fmt(p) + "%"}</div>
    `;
  }

  function renderComparisonTable(ds) {
    const wrap = $("#comparisonTableWrap");
    const table = $("#comparisonTable");
    if (!wrap || !table) return;

    const rows = ds.rows;
    const control = ds.control;

    // Identify the order of treatments (control first)
    const treatments = rows;

    // Build header rows:
    // Row 1: Indicator | Control | (Treatment name colspan=2 for each treatment)
    // Row 2: (blank)   | Value   | Value | Δ vs Control   ...
    const h1 = `
      <tr>
        <th class="sticky-col sticky-ind" rowspan="2">Indicator</th>
        <th class="sticky-col sticky-ctl" rowspan="2">Control (baseline)</th>
        ${treatments
          .filter(r => !r.isControl)
          .map(r => `<th class="treat-head ${r.id === focusedTreatmentId ? "focused" : ""}" colspan="2" data-colhead="${esc(r.id)}">${esc(r.name)}</th>`)
          .join("")}
      </tr>
    `;

    const h2 = `
      <tr>
        ${treatments
          .filter(r => !r.isControl)
          .map(r => `
            <th class="subhead ${r.id === focusedTreatmentId ? "focused" : ""}" data-subhead="${esc(r.id)}">Value</th>
            <th class="subhead ${r.id === focusedTreatmentId ? "focused" : ""}" data-subhead="${esc(r.id)}">Δ vs Control</th>
          `)
          .join("")}
      </tr>
    `;

    const body = INDICATORS.map(ind => {
      const key = ind.key;

      // Control value
      const cVal =
        key === "rank"
          ? "—"
          : ind.format(control.metrics[key]);

      // For rank, use rankByNpv for each treatment
      const tCells = treatments
        .filter(r => !r.isControl)
        .map(r => {
          const isFocused = r.id === focusedTreatmentId;
          const val = key === "rank" ? (r.rankByNpv ? String(r.rankByNpv) : "—") : ind.format(r.metrics[key]);
          const delta =
            key === "pvBenefits"
              ? formatDeltaMoney(r.delta.pvBenefits, control.metrics.pvBenefits)
              : key === "pvCosts"
              ? formatDeltaMoney(r.delta.pvCosts, control.metrics.pvCosts)
              : key === "npv"
              ? formatDeltaMoney(r.delta.npv, control.metrics.npv)
              : key === "bcr"
              ? formatDeltaRatio(r.delta.bcr, control.metrics.bcr)
              : key === "roi"
              ? formatDeltaRatio(r.delta.roi, control.metrics.roi)
              : `<div class="d-abs muted">—</div>`;

          return `
            <td class="val ${isFocused ? "focused" : ""}" data-col="${esc(r.id)}">${val}</td>
            <td class="delta ${isFocused ? "focused" : ""}" data-col="${esc(r.id)}">${delta}</td>
          `;
        })
        .join("");

      return `
        <tr data-indicator="${esc(key)}">
          <td class="sticky-col sticky-ind">
            <div class="ind-cell">
              <div class="ind-title">${esc(ind.label)}</div>
              <button class="btn tiny ghost" type="button" data-open-calc="${esc(key)}" aria-label="Calculation details for ${esc(ind.label)}">Details</button>
            </div>
          </td>
          <td class="sticky-col sticky-ctl ctl-val">${cVal}</td>
          ${tCells}
        </tr>
      `;
    }).join("");

    table.innerHTML = `
      <thead>${h1}${h2}</thead>
      <tbody>${body}</tbody>
    `;

    // Details drawer handlers
    table.onclick = e => {
      const btn = e.target.closest("[data-open-calc]");
      if (!btn) return;
      openCalcDetails(btn.dataset.openCalc, ds);
    };
  }

  function focusColumnIntoView(treatmentId) {
    const wrap = $("#comparisonTableWrap");
    if (!wrap) return;
    const head = $(`[data-colhead="${CSS.escape(treatmentId)}"]`);
    if (!head) return;

    const wrapRect = wrap.getBoundingClientRect();
    const headRect = head.getBoundingClientRect();
    const delta = headRect.left - wrapRect.left - 24;

    wrap.scrollLeft += delta;
  }

  function openCalcDetails(indicatorKey, ds) {
    const drawer = $("#calcDrawer");
    const body = $("#calcDrawerBody");
    const title = $("#calcDrawerTitle");
    if (!drawer || !body || !title) return;

    const ind = INDICATORS.find(x => x.key === indicatorKey);
    if (!ind) return;

    const focus = ds.rows.find(r => r.id === focusedTreatmentId) || ds.rows.find(r => !r.isControl) || ds.rows[0];
    const control = ds.control;

    // Use focused treatment metrics for reconciliation view
    const m = focus.metrics;
    const checks = focus.metrics ? (computeTreatmentScenarioMetrics(
      model.treatments.find(t => t.id === focus.id) || model.treatments[0],
      currentScenarioSettings()
    ).checks) : null;

    const tol = 1e-6;
    const chkRow = (k, label) => {
      const c = checks?.[k];
      if (!c) return "";
      const diff = (Number(c.lhs) || 0) - (Number(c.rhs) || 0);
      const ok = Number.isFinite(diff) ? Math.abs(diff) <= tol : false;
      return `
        <tr>
          <td>${esc(label)}</td>
          <td>${isFiniteNum(c.lhs) ? fmt(c.lhs) : "n/a"}</td>
          <td>${isFiniteNum(c.rhs) ? fmt(c.rhs) : "n/a"}</td>
          <td class="${ok ? "pos" : "neg"}">${ok ? "PASS" : "CHECK"}</td>
          <td>${isFiniteNum(diff) ? fmt(diff) : "n/a"}</td>
        </tr>
      `;
    };

    title.textContent = `Calculation details — ${ind.label}`;

    const cVal = indicatorKey === "rank" ? 0 : control.metrics[indicatorKey];
    const tVal = indicatorKey === "rank" ? (focus.rankByNpv || 0) : focus.metrics[indicatorKey];

    body.innerHTML = `
      <div class="calc-block">
        <div class="calc-formula">
          <div class="badge">Formula</div>
          <p><strong>Plain English:</strong> ${esc(ind.formula.english)}</p>
          <p><strong>Algebra:</strong> <code>${esc(ind.formula.algebra)}</code></p>
        </div>

        <div class="calc-example">
          <div class="badge">Example values (current scenario)</div>
          <div class="kv">
            <div class="kvrow"><div class="k">Focused treatment</div><div class="v"><strong>${esc(focus.name)}</strong></div></div>
            <div class="kvrow"><div class="k">Control (baseline)</div><div class="v"><strong>${esc(control.name)}</strong></div></div>
            <div class="kvrow"><div class="k">${esc(ind.label)} — Control</div><div class="v">${indicatorKey === "rank" ? "—" : esc(ind.format(cVal))}</div></div>
            <div class="kvrow"><div class="k">${esc(ind.label)} — Treatment</div><div class="v">${indicatorKey === "rank" ? esc(String(tVal || "—")) : esc(ind.format(tVal))}</div></div>
          </div>
        </div>
      </div>

      <div class="calc-block">
        <div class="badge">Reconciliation checks (automatic)</div>
        <div class="small muted">
          These checks confirm the accounting identities: NPV = PV(B) − PV(C), ROI = NPV / PV(C), BCR = PV(B) / PV(C), and PVs equal discounted sums.
        </div>
        <div class="table-wrap">
          <table class="mini">
            <thead>
              <tr>
                <th>Check</th>
                <th>LHS</th>
                <th>RHS</th>
                <th>Status</th>
                <th>Diff</th>
              </tr>
            </thead>
            <tbody>
              ${chkRow("pvBenefits", "PV Benefits")}
              ${chkRow("pvCosts", "PV Costs")}
              ${chkRow("npv", "NPV identity")}
              ${chkRow("bcr", "BCR identity")}
              ${chkRow("roi", "ROI identity")}
            </tbody>
          </table>
        </div>
      </div>
    `;

    drawer.classList.add("open");
    $("#calcDrawerClose")?.focus();
    logEvent("Opened calculation details drawer", { indicator: indicatorKey, focusedTreatmentId });
  }

  function closeCalcDrawer() {
    $("#calcDrawer")?.classList.remove("open");
  }

  // ----------------------------
  // Simulations (one-way + break-even + scenario sets + Monte Carlo minimal)
  // ----------------------------
  function currentScenarioSettings() {
    // Base scenario uses model.time + model.adoption/risk + one-way sliders (if user is in Simulations tab)
    // Results tab uses these “current settings” (transparent and auditable).
    const ow = model.sim.oneWay;
    const priceOut = model.outputs.find(o => o.id === "out_yield");
    if (priceOut) priceOut.value = Number(ow.grainPrice) || priceOut.value;

    return {
      startYear: model.time.startYear,
      years: Number(ow.years ?? model.time.years),
      discountRate: Number(ow.discountRate ?? model.time.discBase),
      adoption: Number(ow.adoption ?? model.adoption.base),
      risk: Number(ow.risk ?? model.risk.base),
      grainPrice: Number(ow.grainPrice ?? (priceOut?.value || 450)),
      yieldDeltaMultiplier: Number(ow.yieldDeltaMultiplier ?? 1),
      operatingCostMultiplier: Number(ow.operatingCostMultiplier ?? 1),
      capitalCostMultiplier: Number(ow.capitalCostMultiplier ?? 1)
    };
  }

  function renderSimulations() {
    renderOneWayControls();
    renderBreakEven();
    renderScenarioSets();
    renderResults(); // keep results synced with current scenario settings
  }

  function renderOneWayControls() {
    const root = $("#oneWayControls");
    if (!root) return;

    const ow = model.sim.oneWay;

    root.innerHTML = `
      <div class="grid-2">
        <div class="card">
          <h3>One-way sensitivity (fast sliders)</h3>
          <div class="small muted">These sliders update Results immediately. They do not overwrite your uploaded data; they adjust assumptions.</div>

          <div class="row-3">
            <div class="field">
              <label>Grain price ($/t)</label>
              <input id="owGrainPrice" type="number" step="1" value="${esc(String(ow.grainPrice))}">
            </div>
            <div class="field">
              <label>Yield uplift multiplier</label>
              <input id="owYieldMult" type="number" step="0.01" value="${esc(String(ow.yieldDeltaMultiplier))}">
            </div>
            <div class="field">
              <label>Discount rate (%)</label>
              <input id="owDisc" type="number" step="0.1" value="${esc(String(ow.discountRate))}">
            </div>
          </div>

          <div class="row-3">
            <div class="field">
              <label>Adoption / implementation rate (0–1)</label>
              <input id="owAdopt" type="number" step="0.01" min="0" max="1" value="${esc(String(ow.adoption))}">
            </div>
            <div class="field">
              <label>Risk multiplier (0–1)</label>
              <input id="owRisk" type="number" step="0.01" min="0" max="1" value="${esc(String(ow.risk))}">
            </div>
            <div class="field">
              <label>Horizon (years)</label>
              <input id="owYears" type="number" step="1" min="1" value="${esc(String(ow.years ?? model.time.years))}">
            </div>
          </div>

          <div class="row-3">
            <div class="field">
              <label>Operating cost multiplier</label>
              <input id="owOpMult" type="number" step="0.01" value="${esc(String(ow.operatingCostMultiplier))}">
            </div>
            <div class="field">
              <label>Capital cost multiplier</label>
              <input id="owCapMult" type="number" step="0.01" value="${esc(String(ow.capitalCostMultiplier))}">
            </div>
            <div class="field">
              <label>&nbsp;</label>
              <button class="btn" id="owApply">Apply</button>
            </div>
          </div>
        </div>

        <div class="card">
          <h3>Break-even tools (decision support)</h3>
          <div class="small muted">These show “what would need to change” for selected treatments, without recommending a choice.</div>
          <div id="breakEvenBox"></div>
        </div>
      </div>
    `;

    $("#owApply")?.addEventListener("click", () => {
      ow.grainPrice = parseNumberLoose($("#owGrainPrice")?.value) || ow.grainPrice;
      ow.yieldDeltaMultiplier = parseNumberLoose($("#owYieldMult")?.value) || ow.yieldDeltaMultiplier;
      ow.discountRate = parseNumberLoose($("#owDisc")?.value) || ow.discountRate;
      ow.adoption = clamp(parseNumberLoose($("#owAdopt")?.value) || ow.adoption, 0, 1);
      ow.risk = clamp(parseNumberLoose($("#owRisk")?.value) || ow.risk, 0, 1);
      ow.years = Math.max(1, Math.floor(parseNumberLoose($("#owYears")?.value) || model.time.years));
      ow.operatingCostMultiplier = parseNumberLoose($("#owOpMult")?.value) || ow.operatingCostMultiplier;
      ow.capitalCostMultiplier = parseNumberLoose($("#owCapMult")?.value) || ow.capitalCostMultiplier;

      logEvent("Applied one-way sensitivity settings", { ...ow });
      renderSimulations();
      showToast("Scenario settings applied and Results updated.");
    });
  }

  function solveForBreakEvenPrice(treatment, baseSettings) {
    // Solve grain price that makes NPV = 0 vs control, holding everything else fixed.
    const outYield = model.outputs.find(o => o.id === "out_yield");
    if (!outYield) return null;

    const t = treatment;
    if (t.isControl) return null;

    // NPV is approximately linear in grain price via yield delta contribution.
    // We compute NPV at price p0 and p1 and linearly interpolate.
    const p0 = Math.max(1, Number(baseSettings.grainPrice) || 450);
    const p1 = p0 * 1.25;

    const original = outYield.value;
    outYield.value = p0;
    const m0 = computeTreatmentScenarioMetrics(t, baseSettings).npv;
    outYield.value = p1;
    const m1 = computeTreatmentScenarioMetrics(t, baseSettings).npv;
    outYield.value = original;

    if (!isFiniteNum(m0) || !isFiniteNum(m1) || Math.abs(m1 - m0) < 1e-9) return null;

    const pStar = p0 + (0 - m0) * (p1 - p0) / (m1 - m0);
    if (!isFiniteNum(pStar)) return null;
    return pStar;
  }

  function solveForBreakEvenYieldMult(treatment, baseSettings) {
    // Solve yield multiplier that makes BCR = 1, holding costs fixed.
    const t = treatment;
    if (t.isControl) return null;

    const f = (mult) => {
      const s = { ...baseSettings, yieldDeltaMultiplier: mult };
      return computeTreatmentScenarioMetrics(t, s).bcr;
    };

    // Simple bracket search
    let lo = 0;
    let hi = 5;
    const target = 1;

    let fLo = f(lo);
    let fHi = f(hi);

    if (!Number.isFinite(fLo) || !Number.isFinite(fHi)) return null;
    if ((fLo - target) * (fHi - target) > 0) return null;

    for (let i = 0; i < 60; i++) {
      const mid = (lo + hi) / 2;
      const fMid = f(mid);
      if (!Number.isFinite(fMid)) return null;
      if (Math.abs(fMid - target) < 1e-5) return mid;
      if ((fLo - target) * (fMid - target) <= 0) {
        hi = mid;
        fHi = fMid;
      } else {
        lo = mid;
        fLo = fMid;
      }
    }
    return (lo + hi) / 2;
  }

  function solveForCostReductionToReachRoi0(treatment, baseSettings) {
    // Solve operating cost multiplier that makes ROI = 0 (i.e., NPV = 0 since PV(C)>0), holding benefits fixed.
    const t = treatment;
    if (t.isControl) return null;

    const f = (opMult) => {
      const s = { ...baseSettings, operatingCostMultiplier: opMult };
      return computeTreatmentScenarioMetrics(t, s).npv;
    };

    // Want NPV = 0; costs up -> NPV down; costs down -> NPV up.
    let lo = 0;
    let hi = 2.5;
    let fLo = f(lo);
    let fHi = f(hi);

    if (!isFiniteNum(fLo) || !isFiniteNum(fHi)) return null;

    // Ensure bracket
    if (fLo < 0 && fHi < 0) return null;
    if (fLo > 0 && fHi > 0) return 1; // already >=0 at current (approx)

    for (let i = 0; i < 60; i++) {
      const mid = (lo + hi) / 2;
      const fMid = f(mid);
      if (!isFiniteNum(fMid)) return null;
      if (Math.abs(fMid) < 1e-5) return mid;
      if (fLo * fMid <= 0) {
        hi = mid;
        fHi = fMid;
      } else {
        lo = mid;
        fLo = fMid;
      }
    }
    return (lo + hi) / 2;
  }

  function renderBreakEven() {
    const box = $("#breakEvenBox");
    const out = $("#breakEvenBox") || $("#breakEvenBox");
    const host = $("#breakEvenBox") ? $("#breakEvenBox").closest(".card") : null;
    const inner = $("#breakEvenBox");
    const root = $("#breakEvenBox");

    const container = $("#breakEvenBox") || $("#breakEvenBox");
    const breakBox = $("#breakEvenBox");
    const breakEvenTarget = $("#breakEvenBox");

    const wrap = $("#breakEvenBox");
    const be = $("#breakEvenBox");

    const rootBox = $("#breakEvenBox");
    const target = $("#breakEvenBox");

    const r = $("#breakEvenBox");
    const p = $("#breakEvenBox");

    const rootEl = $("#breakEvenBox");
    if (!rootEl) return;

    const treatments = model.treatments.filter(t => !t.isControl);
    if (!treatments.length) {
      rootEl.innerHTML = `<div class="small muted">No treatments available.</div>`;
      return;
    }

    // pick focused as default
    const focus = model.treatments.find(t => t.id === focusedTreatmentId && !t.isControl) || treatments[0];
    const settings = currentScenarioSettings();

    const pStar = solveForBreakEvenPrice(focus, settings);
    const yStar = solveForBreakEvenYieldMult(focus, settings);
    const cStar = solveForCostReductionToReachRoi0(focus, settings);

    rootEl.innerHTML = `
      <div class="field">
        <label>Selected treatment</label>
        <select id="beTreatment">
          ${treatments.map(t => `<option value="${esc(t.id)}" ${t.id === focus.id ? "selected" : ""}>${esc(t.name)}</option>`).join("")}
        </select>
      </div>

      <div class="be-grid">
        <div class="be-item">
          <div class="be-title">Break-even grain price</div>
          <div class="be-val">${pStar === null ? "n/a" : money(pStar)}</div>
          <div class="be-note muted">Grain price that makes NPV ≈ 0 (vs control), holding costs and other settings fixed.</div>
        </div>
        <div class="be-item">
          <div class="be-title">Yield uplift multiplier for BCR = 1</div>
          <div class="be-val">${yStar === null ? "n/a" : fmt(yStar)}</div>
          <div class="be-note muted">Scale factor on yield deltas needed to reach BCR ≈ 1 at current costs.</div>
        </div>
        <div class="be-item">
          <div class="be-title">Operating cost multiplier for ROI = 0</div>
          <div class="be-val">${cStar === null ? "n/a" : fmt(cStar)}</div>
          <div class="be-note muted">Cost scaling needed so NPV ≈ 0 at current benefits (lower is “needs costs to fall”).</div>
        </div>
      </div>
    `;

    $("#beTreatment")?.addEventListener("change", e => {
      const id = e.target.value;
      focusedTreatmentId = id;
      renderSimulations();
      focusColumnIntoView(id);
      logEvent("Changed break-even treatment selection", { treatmentId: id });
    });
  }

  function loadSavedScenarios() {
    try {
      const raw = localStorage.getItem(LS_SCENARIOS);
      if (!raw) return;
      const arr = JSON.parse(raw);
      if (Array.isArray(arr)) model.sim.scenarios = arr;
    } catch {}
  }

  function saveScenarios() {
    try {
      localStorage.setItem(LS_SCENARIOS, JSON.stringify(model.sim.scenarios));
    } catch {}
  }

  function renderScenarioSets() {
    const root = $("#scenarioSets");
    if (!root) return;

    const s = currentScenarioSettings();

    root.innerHTML = `
      <div class="card">
        <div class="row space-between">
          <div>
            <h3>Scenario sets</h3>
            <div class="small muted">Save named scenarios (e.g., “Dry year”, “High fuel cost”) and compare instantly.</div>
          </div>
          <div class="row">
            <button class="btn" id="saveScenario">Save current scenario</button>
          </div>
        </div>

        <div class="row-2">
          <div class="field">
            <label>Scenario name</label>
            <input id="scenarioName" placeholder="e.g., Dry year" />
          </div>
          <div class="field">
            <label>&nbsp;</label>
            <div class="small muted">Saved scenarios are stored in your browser (local storage).</div>
          </div>
        </div>

        <div class="table-wrap">
          <table class="mini">
            <thead>
              <tr>
                <th>Name</th>
                <th>Years</th>
                <th>Discount</th>
                <th>Price</th>
                <th>Adoption</th>
                <th>Risk</th>
                <th>Op cost ×</th>
                <th>Cap cost ×</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              ${
                model.sim.scenarios.length
                  ? model.sim.scenarios
                      .map((sc, i) => `
                        <tr>
                          <td><strong>${esc(sc.name || "Scenario")}</strong></td>
                          <td>${esc(String(sc.years))}</td>
                          <td>${esc(String(sc.discountRate))}%</td>
                          <td>${money(sc.grainPrice)}</td>
                          <td>${esc((sc.adoption * 100).toFixed(0))}%</td>
                          <td>${esc((sc.risk * 100).toFixed(0))}%</td>
                          <td>${esc(fmt(sc.operatingCostMultiplier))}</td>
                          <td>${esc(fmt(sc.capitalCostMultiplier))}</td>
                          <td class="row">
                            <button class="btn tiny" data-apply-scn="${i}">Apply</button>
                            <button class="btn tiny danger" data-del-scn="${i}">Delete</button>
                          </td>
                        </tr>
                      `)
                      .join("")
                  : `<tr><td colspan="9" class="muted">No saved scenarios yet.</td></tr>`
              }
            </tbody>
          </table>
        </div>
      </div>
    `;

    $("#saveScenario")?.addEventListener("click", () => {
      const name = ($("#scenarioName")?.value || "").trim() || `Scenario ${model.sim.scenarios.length + 1}`;
      const snap = { ...s, name };
      model.sim.scenarios.unshift(snap);
      saveScenarios();
      logEvent("Saved scenario", { name, settings: snap });
      renderScenarioSets();
      showToast("Scenario saved.");
    });

    root.onclick = e => {
      const apply = e.target.closest("[data-apply-scn]");
      if (apply) {
        const idx = +apply.dataset.applyScn;
        const sc = model.sim.scenarios[idx];
        if (!sc) return;
        Object.assign(model.sim.oneWay, {
          grainPrice: sc.grainPrice,
          yieldDeltaMultiplier: sc.yieldDeltaMultiplier,
          discountRate: sc.discountRate,
          adoption: sc.adoption,
          risk: sc.risk,
          years: sc.years,
          operatingCostMultiplier: sc.operatingCostMultiplier,
          capitalCostMultiplier: sc.capitalCostMultiplier
        });
        logEvent("Applied saved scenario", { name: sc.name });
        renderSimulations();
        showToast(`Applied scenario: ${sc.name}`);
        return;
      }

      const del = e.target.closest("[data-del-scn]");
      if (del) {
        const idx = +del.dataset.delScn;
        const sc = model.sim.scenarios[idx];
        if (!sc) return;
        model.sim.scenarios.splice(idx, 1);
        saveScenarios();
        logEvent("Deleted scenario", { name: sc.name });
        renderScenarioSets();
        showToast("Scenario deleted.");
      }
    };
  }

  // ----------------------------
  // AI / Copilot structured prompt pack
  // ----------------------------
  function buildComparisonJSON(ds) {
    const treatments = ds.rows.filter(r => !r.isControl);

    return {
      control: {
        name: ds.control.name,
        pvBenefits: ds.control.metrics.pvBenefits,
        pvCosts: ds.control.metrics.pvCosts,
        npv: ds.control.metrics.npv,
        bcr: ds.control.metrics.bcr,
        roi: ds.control.metrics.roi
      },
      treatments: treatments.map(r => ({
        name: r.name,
        pvBenefits: r.metrics.pvBenefits,
        pvCosts: r.metrics.pvCosts,
        npv: r.metrics.npv,
        bcr: r.metrics.bcr,
        roi: r.metrics.roi,
        deltaVsControl: {
          pvBenefits: r.delta.pvBenefits,
          pvCosts: r.delta.pvCosts,
          npv: r.delta.npv
        },
        rankByNpv: r.rankByNpv
      }))
    };
  }

  function buildDriversPerTreatment() {
    // Biggest drivers: cost components (absolute PV contribution proxies) + yield benefit component
    // We report per-ha components so it's interpretable.
    const outYield = model.outputs.find(o => o.id === "out_yield");
    const price = Number(outYield?.value) || 0;

    return model.treatments
      .filter(t => !t.isControl)
      .map(t => {
        const costs = (t.costItems || [])
          .map(it => ({ label: it.label, perHa: Number(it.valuePerHa) || 0 }))
          .sort((a, b) => Math.abs(b.perHa) - Math.abs(a.perHa))
          .slice(0, 8);

        const yDelta = Number(t.deltas?.[outYield?.id] || 0);
        const yieldValuePerHa = yDelta * price;

        return {
          treatment: t.name,
          yieldUplift_t_ha: yDelta,
          yieldValue_perHa: yieldValuePerHa,
          topCostComponents_perHa: costs
        };
      });
  }

  function buildAiPrompt() {
    const settings = currentScenarioSettings();
    const ds = computeComparisonDataset(settings, resultsFilterMode);
    const comp = buildComparisonJSON(ds);
    const drivers = buildDriversPerTreatment();

    const payload = {
      tool: { name: TOOL.name, version: TOOL.version, organisation: TOOL.organisation },
      scenario: {
        areaNote: "Each treatment is evaluated as an alternative compared to Control (baseline). Values are incremental versus control.",
        years: settings.years,
        discountRatePct: settings.discountRate,
        grainPricePerTonne: settings.grainPrice,
        adoptionRate: settings.adoption,
        riskMultiplier: settings.risk,
        operatingCostMultiplier: settings.operatingCostMultiplier,
        capitalCostMultiplier: settings.capitalCostMultiplier
      },
      comparisonToControl: comp,
      biggestDrivers: drivers,
      instructions: [
        "Write in plain language suitable for a farmer or on-farm manager.",
        "Explain what PV Benefits, PV Costs, NPV, BCR, and ROI mean in practical terms.",
        "Compare each treatment against the Control (baseline) and describe what drives the differences.",
        "Do NOT recommend a choice and do NOT impose decision rules or thresholds.",
        "For treatments with negative ΔNPV or low BCR, suggest improvement options framed as possibilities (reduce costs, improve yield, improve price, improve implementation efficiency, agronomy options).",
        "Keep the tone decision-support and highlight uncertainty and sensitivity."
      ]
    };

    const prompt =
      `You are interpreting results from ${TOOL.name} (version ${TOOL.version}).\n` +
      `Use the JSON below as the only source of numbers. Do not invent values.\n\n` +
      `JSON:\n` +
      JSON.stringify(payload, null, 2);

    return { prompt, payload };
  }

  function renderAI() {
    const box = $("#aiBox");
    if (!box) return;

    const { prompt } = buildAiPrompt();
    box.innerHTML = `
      <div class="card">
        <div class="row space-between">
          <div>
            <h3>AI / Copilot prompt pack (structured)</h3>
            <div class="small muted">This generates a copy-paste prompt (not prose) with scenario settings, the full comparison-to-control table, and the biggest drivers.</div>
          </div>
          <div class="row">
            <button class="btn" id="aiCopy">Copy prompt</button>
            <button class="btn" id="aiExport">Export AI brief pack</button>
          </div>
        </div>

        <div class="field">
          <label>Prompt</label>
          <textarea id="aiPrompt" rows="18">${esc(prompt)}</textarea>
        </div>
      </div>
    `;

    $("#aiCopy")?.addEventListener("click", async () => {
      const txt = $("#aiPrompt")?.value || "";
      try {
        await navigator.clipboard.writeText(txt);
        showToast("AI prompt copied to clipboard.");
        logEvent("Copied AI prompt", {});
      } catch {
        showToast("Copy failed. Please select and copy manually.");
      }
    });

    $("#aiExport")?.addEventListener("click", () => {
      const { payload } = buildAiPrompt();
      const pack = {
        tool: TOOL,
        exportedAt: nowISO(),
        payload
      };
      downloadFile(`${slugify(model.project.name)}_ai_brief_pack.json`, JSON.stringify(pack, null, 2), "application/json");
      showToast("AI brief pack exported.");
      logEvent("Exported AI brief pack", {});
    });
  }

  // ----------------------------
  // Exports (Excel + Print PDF)
  // ----------------------------
  function exportExcel() {
    if (typeof XLSX === "undefined") {
      alert("Excel export requires the SheetJS (XLSX) library. Ensure the XLSX script is loaded.");
      return;
    }

    const settings = currentScenarioSettings();
    const ds = computeComparisonDataset(settings, resultsFilterMode);

    // Results sheet (table-like)
    const header1 = ["Indicator", "Control (baseline)"];
    const header2 = ["", ""];
    const treatCols = ds.rows.filter(r => !r.isControl);

    for (const r of treatCols) {
      header1.push(r.name, r.name);
      header2.push("Value", "Δ vs Control");
    }

    const rows = [header1, header2];

    for (const ind of INDICATORS) {
      const key = ind.key;
      const line = [ind.label];

      // control value
      line.push(
        key === "rank" ? "" : ds.control.metrics[key]
      );

      for (const r of treatCols) {
        const val = key === "rank" ? (r.rankByNpv || "") : r.metrics[key];
        let d = "";
        if (key === "pvBenefits") d = `${r.delta.pvBenefits} (${pctDelta(r.delta.pvBenefits, ds.control.metrics.pvBenefits) ?? ""}%)`;
        if (key === "pvCosts") d = `${r.delta.pvCosts} (${pctDelta(r.delta.pvCosts, ds.control.metrics.pvCosts) ?? ""}%)`;
        if (key === "npv") d = `${r.delta.npv} (${pctDelta(r.delta.npv, ds.control.metrics.npv) ?? ""}%)`;
        if (key === "bcr") d = `${r.delta.bcr}`;
        if (key === "roi") d = `${r.delta.roi}`;
        if (key === "rank") d = "";
        line.push(val, d);
      }
      rows.push(line);
    }

    const wsResults = XLSX.utils.aoa_to_sheet(rows);

    // Inputs sheet: treatments cost build-up + outputs
    const inputRows = [];
    inputRows.push(["Tool", TOOL.name]);
    inputRows.push(["Version", TOOL.version]);
    inputRows.push(["Exported at", nowISO()]);
    inputRows.push([]);
    inputRows.push(["Scenario years", settings.years]);
    inputRows.push(["Discount rate (%)", settings.discountRate]);
    inputRows.push(["Grain price ($/t)", settings.grainPrice]);
    inputRows.push(["Adoption", settings.adoption]);
    inputRows.push(["Risk", settings.risk]);
    inputRows.push(["Operating cost multiplier", settings.operatingCostMultiplier]);
    inputRows.push(["Capital cost multiplier", settings.capitalCostMultiplier]);
    inputRows.push([]);
    inputRows.push(["Outputs (value per unit)"]);
    inputRows.push(["Output", "Unit", "Value", "Source"]);
    model.outputs.forEach(o => inputRows.push([o.name, o.unit, o.value, o.source || ""]));
    inputRows.push([]);
    inputRows.push(["Treatments (incremental vs control)"]);
    inputRows.push(["Treatment", "IsControl", "AreaHa", "CapitalCostY0", "OperatingCostPerHa (sum)"]);
    model.treatments.forEach(t => {
      const op = (t.costItems || []).reduce((a, it) => a + (Number(it.valuePerHa) || 0), 0);
      inputRows.push([t.name, t.isControl ? 1 : 0, t.areaHa, t.capitalCostY0, op]);
    });
    inputRows.push([]);
    inputRows.push(["Treatment cost build-up (per ha deltas vs control)"]);
    inputRows.push(["Treatment", "Cost item", "Category", "ValuePerHa"]);
    model.treatments.forEach(t => {
      (t.costItems || []).forEach(ci => inputRows.push([t.name, ci.label, ci.category, ci.valuePerHa]));
    });

    const wsInputs = XLSX.utils.aoa_to_sheet(inputRows);

    // Assumptions sheet
    const wsAssum = XLSX.utils.aoa_to_sheet([
      ["Project name", model.project.name],
      ["Organisation", model.project.organisation],
      ["Summary", model.project.summary],
      ["Assumptions (outputs)", model.outputsMeta.assumptions],
      [],
      ["Incremental convention", "Treatments are evaluated as alternatives; values are incremental versus Control (baseline)."],
      ["Aggregation rule", model.data.replicateRule]
    ]);

    // Simulations sheet (saved scenarios)
    const scnRows = [["Saved scenarios"], ["Name", "Years", "Discount", "Price", "Adoption", "Risk", "OpCost×", "CapCost×"]];
    model.sim.scenarios.forEach(sc => scnRows.push([sc.name, sc.years, sc.discountRate, sc.grainPrice, sc.adoption, sc.risk, sc.operatingCostMultiplier, sc.capitalCostMultiplier]));
    const wsSim = XLSX.utils.aoa_to_sheet(scnRows);

    // Audit sheet
    const audRows = [["Audit log (most recent first)"], ["Time", "Action", "Details"]];
    model.audit.forEach(a => audRows.push([a.time, a.action, JSON.stringify(a.details || {})]));
    const wsAudit = XLSX.utils.aoa_to_sheet(audRows);

    // AI sheet
    const { prompt } = buildAiPrompt();
    const wsAI = XLSX.utils.aoa_to_sheet([
      ["AI / Copilot prompt"],
      ["Tool", TOOL.name],
      ["Version", TOOL.version],
      ["Exported at", nowISO()],
      [],
      ["Prompt"],
      [prompt]
    ]);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsResults, "Results_Comparison");
    XLSX.utils.book_append_sheet(wb, wsInputs, "Inputs");
    XLSX.utils.book_append_sheet(wb, wsAssum, "Assumptions");
    XLSX.utils.book_append_sheet(wb, wsSim, "Simulations");
    XLSX.utils.book_append_sheet(wb, wsAudit, "Audit");
    XLSX.utils.book_append_sheet(wb, wsAI, "AI_Prompt");

    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile(`${slugify(model.project.name)}_export.xlsx`, out, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Excel export created.");
    logEvent("Exported Excel workbook", {});
  }

  function exportPdf() {
    // Print pipeline: CSS includes print layout and a “condensed mode”
    window.print();
    logEvent("Opened print dialog", {});
  }

  // ----------------------------
  // Treatments tab: cost build-up (capital first; operating line-items; totals drive calcs)
  // ----------------------------
  function renderTreatments() {
    const root = $("#treatmentsList");
    if (!root) return;

    root.innerHTML = "";

    model.treatments.forEach(t => {
      const op = (t.costItems || []).reduce((a, it) => a + (Number(it.valuePerHa) || 0), 0);
      const opFarm = op * (Number(t.areaHa) || 0);
      const cap = Number(t.capitalCostY0) || 0;

      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>${esc(t.name)} ${t.isControl ? `<span class="badge">Control (baseline)</span>` : ""}</h4>

        <div class="row-4">
          <div class="field">
            <label>Name</label>
            <input data-tk="name" data-id="${esc(t.id)}" value="${esc(t.name)}">
          </div>
          <div class="field">
            <label>Area (ha)</label>
            <input type="number" step="0.01" data-tk="areaHa" data-id="${esc(t.id)}" value="${esc(String(t.areaHa ?? 0))}">
          </div>
          <div class="field">
            <label>Control vs treatment?</label>
            <select data-tk="isControl" data-id="${esc(t.id)}">
              <option value="control" ${t.isControl ? "selected" : ""}>Control (baseline)</option>
              <option value="treatment" ${!t.isControl ? "selected" : ""}>Treatment</option>
            </select>
          </div>
          <div class="field">
            <label>Adoption (0–1)</label>
            <input type="number" step="0.01" min="0" max="1" data-tk="adoption" data-id="${esc(t.id)}" value="${esc(String(t.adoption ?? 1))}">
          </div>
        </div>

        <div class="card subtle">
          <h5>Cost build-up (incremental vs control)</h5>

          <div class="row-3">
            <div class="field">
              <label>Capital cost ($, year 0)</label>
              <input type="number" step="0.01" data-tk="capitalCostY0" data-id="${esc(t.id)}" value="${esc(String(t.capitalCostY0 ?? 0))}">
            </div>

            <div class="field">
              <label>Total operating cost ($/ha)</label>
              <input type="number" step="0.01" value="${esc(String(op))}" readonly>
            </div>

            <div class="field">
              <label>Whole-farm operating total ($/year)</label>
              <input type="number" step="0.01" value="${esc(String(opFarm))}" readonly>
            </div>
          </div>

          <div class="table-wrap">
            <table class="mini">
              <thead>
                <tr>
                  <th>Operating cost component</th>
                  <th>Category</th>
                  <th>$ / ha</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                ${(t.costItems || []).map(ci => `
                  <tr>
                    <td><input data-ci-k="label" data-ci-id="${esc(ci.id)}" data-tid="${esc(t.id)}" value="${esc(ci.label)}"></td>
                    <td>
                      <select data-ci-k="category" data-ci-id="${esc(ci.id)}" data-tid="${esc(t.id)}">
                        ${["Materials", "Services", "Labour"].map(cat => `<option ${ci.category === cat ? "selected" : ""}>${cat}</option>`).join("")}
                      </select>
                    </td>
                    <td><input type="number" step="0.01" data-ci-k="valuePerHa" data-ci-id="${esc(ci.id)}" data-tid="${esc(t.id)}" value="${esc(String(ci.valuePerHa ?? 0))}"></td>
                    <td><button class="btn tiny danger" data-del-ci="${esc(ci.id)}" data-tid="${esc(t.id)}">Remove</button></td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>

          <div class="row">
            <button class="btn tiny" data-add-ci="${esc(t.id)}">Add operating cost line</button>
          </div>
        </div>

        <div class="card subtle">
          <h5>Output deltas (per ha, incremental vs control)</h5>
          <div class="row">
            ${model.outputs.map(o => `
              <div class="field">
                <label>${esc(o.name)} (${esc(o.unit)})</label>
                <input type="number" step="0.0001" data-od="${esc(o.id)}" data-tid="${esc(t.id)}" value="${esc(String(t.deltas?.[o.id] ?? 0))}">
              </div>
            `).join("")}
          </div>
        </div>
      `;
      root.appendChild(el);
    });

    root.oninput = e => {
      const tid = e.target.dataset.id || e.target.dataset.tid;
      if (!tid) return;

      const t = model.treatments.find(x => x.id === tid);
      if (!t) return;

      // Treatment-level fields
      const tk = e.target.dataset.tk;
      if (tk) {
        if (tk === "name") t.name = e.target.value;
        else if (tk === "areaHa") t.areaHa = +e.target.value;
        else if (tk === "adoption") t.adoption = clamp(+e.target.value, 0, 1);
        else if (tk === "capitalCostY0") t.capitalCostY0 = +e.target.value;
        else if (tk === "isControl") {
          const isCtl = e.target.value === "control";
          model.treatments.forEach(tt => (tt.isControl = false));
          t.isControl = isCtl;
        }
        renderTreatments();
        renderResults();
        return;
      }

      // Cost item fields
      const ciId = e.target.dataset.ciId;
      const ciKey = e.target.dataset.ciK;
      if (ciId && ciKey) {
        const ci = (t.costItems || []).find(x => x.id === ciId);
        if (!ci) return;
        if (ciKey === "label") ci.label = e.target.value;
        else if (ciKey === "category") ci.category = e.target.value;
        else if (ciKey === "valuePerHa") ci.valuePerHa = +e.target.value;
        renderTreatments();
        renderResults();
        return;
      }

      // Output deltas
      const outId = e.target.dataset.od;
      if (outId) {
        t.deltas = t.deltas || {};
        t.deltas[outId] = +e.target.value;
        renderResults();
      }
    };

    root.onclick = e => {
      const add = e.target.closest("[data-add-ci]");
      if (add) {
        const tid = add.dataset.addCi;
        const t = model.treatments.find(x => x.id === tid);
        if (!t) return;
        t.costItems = t.costItems || [];
        t.costItems.push({ id: uid(), label: "New operating cost", category: "Materials", valuePerHa: 0 });
        logEvent("Added operating cost line", { treatmentId: tid });
        renderTreatments();
        renderResults();
        return;
      }

      const del = e.target.closest("[data-del-ci]");
      if (del) {
        const tid = del.dataset.tid;
        const cid = del.dataset.delCi;
        const t = model.treatments.find(x => x.id === tid);
        if (!t) return;
        t.costItems = (t.costItems || []).filter(x => x.id !== cid);
        logEvent("Removed operating cost line", { treatmentId: tid, costItemId: cid });
        renderTreatments();
        renderResults();
        return;
      }
    };
  }

  // ----------------------------
  // Outputs tab
  // ----------------------------
  function renderOutputs() {
    const root = $("#outputsList");
    if (!root) return;

    root.innerHTML = `
      <div class="card">
        <div class="row space-between">
          <div>
            <h3>Outputs (benefit drivers)</h3>
            <div class="small muted">Benefits per hectare are calculated as Σ(delta × value). Set values to monetise outputs beyond yield.</div>
          </div>
          <div><button class="btn" id="addOutput">Add output</button></div>
        </div>

        <div class="table-wrap">
          <table class="mini">
            <thead>
              <tr>
                <th>Output</th>
                <th>Unit</th>
                <th>Value ($ per unit)</th>
                <th>Source / note</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              ${model.outputs.map(o => `
                <tr>
                  <td><input data-ok="name" data-oid="${esc(o.id)}" value="${esc(o.name)}"></td>
                  <td><input data-ok="unit" data-oid="${esc(o.id)}" value="${esc(o.unit)}"></td>
                  <td><input type="number" step="0.01" data-ok="value" data-oid="${esc(o.id)}" value="${esc(String(o.value ?? 0))}"></td>
                  <td><input data-ok="source" data-oid="${esc(o.id)}" value="${esc(o.source || "")}"></td>
                  <td>${o.id === "out_yield" ? `<span class="muted">required</span>` : `<button class="btn tiny danger" data-del-out="${esc(o.id)}">Remove</button>`}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `;

    root.oninput = e => {
      const oid = e.target.dataset.oid;
      const k = e.target.dataset.ok;
      if (!oid || !k) return;
      const o = model.outputs.find(x => x.id === oid);
      if (!o) return;
      if (k === "value") o.value = +e.target.value;
      else o[k] = e.target.value;

      // keep one-way price in sync with yield output
      if (o.id === "out_yield") model.sim.oneWay.grainPrice = o.value;

      renderTreatments();
      renderResults();
    };

    root.onclick = e => {
      if (e.target.id === "addOutput") {
        const id = "out_" + uid();
        model.outputs.push({ id, name: "New output", unit: "unit", value: 0, source: "Input Directly" });
        // add to all treatments deltas
        model.treatments.forEach(t => {
          t.deltas = t.deltas || {};
          t.deltas[id] = 0;
        });
        logEvent("Added output", { outputId: id });
        renderOutputs();
        renderTreatments();
        renderResults();
        return;
      }

      const del = e.target.closest("[data-del-out]");
      if (del) {
        const id = del.dataset.delOut;
        model.outputs = model.outputs.filter(o => o.id !== id);
        model.treatments.forEach(t => {
          if (t.deltas) delete t.deltas[id];
        });
        logEvent("Removed output", { outputId: id });
        renderOutputs();
        renderTreatments();
        renderResults();
      }
    };
  }

  // ----------------------------
  // Project & Settings (minimal but functional)
  // ----------------------------
  function renderProject() {
    const root = $("#projectBox");
    if (!root) return;

    root.innerHTML = `
      <div class="card">
        <h3>Project details</h3>
        <div class="row-2">
          <div class="field"><label>Project name</label><input id="pName" value="${esc(model.project.name)}"></div>
          <div class="field"><label>Organisation</label><input id="pOrg" value="${esc(model.project.organisation)}"></div>
        </div>
        <div class="row-2">
          <div class="field"><label>Lead</label><input id="pLead" value="${esc(model.project.lead)}"></div>
          <div class="field"><label>Analysts</label><input id="pAnalysts" value="${esc(model.project.analysts)}"></div>
        </div>
        <div class="field"><label>Summary</label><textarea id="pSummary" rows="4">${esc(model.project.summary)}</textarea></div>
        <div class="row">
          <button class="btn" id="saveJson">Download project JSON</button>
        </div>
      </div>
    `;

    root.oninput = e => {
      if (e.target.id === "pName") model.project.name = e.target.value;
      if (e.target.id === "pOrg") model.project.organisation = e.target.value;
      if (e.target.id === "pLead") model.project.lead = e.target.value;
      if (e.target.id === "pAnalysts") model.project.analysts = e.target.value;
      if (e.target.id === "pSummary") model.project.summary = e.target.value;
    };

    $("#saveJson")?.addEventListener("click", () => {
      downloadFile(`${slugify(model.project.name)}.json`, JSON.stringify(model, null, 2), "application/json");
      showToast("Project JSON downloaded.");
      logEvent("Downloaded project JSON", {});
    });
  }

  function renderSettings() {
    const root = $("#settingsBox");
    if (!root) return;

    root.innerHTML = `
      <div class="card">
        <h3>Base settings</h3>
        <div class="row-3">
          <div class="field"><label>Start year</label><input id="sStart" type="number" value="${esc(String(model.time.startYear))}"></div>
          <div class="field"><label>Default horizon (years)</label><input id="sYears" type="number" min="1" value="${esc(String(model.time.years))}"></div>
          <div class="field"><label>Base discount rate (%)</label><input id="sDisc" type="number" step="0.1" value="${esc(String(model.time.discBase))}"></div>
        </div>

        <div class="row-2">
          <div class="field"><label>Adoption (base)</label><input id="sAdopt" type="number" step="0.01" min="0" max="1" value="${esc(String(model.adoption.base))}"></div>
          <div class="field"><label>Risk (base)</label><input id="sRisk" type="number" step="0.01" min="0" max="1" value="${esc(String(model.risk.base))}"></div>
        </div>

        <div class="row">
          <button class="btn" id="applyBase">Apply to scenario sliders</button>
        </div>
      </div>
    `;

    $("#applyBase")?.addEventListener("click", () => {
      model.time.startYear = +($("#sStart")?.value || model.time.startYear);
      model.time.years = Math.max(1, +($("#sYears")?.value || model.time.years));
      model.time.discBase = +($("#sDisc")?.value || model.time.discBase);
      model.adoption.base = clamp(+($("#sAdopt")?.value || model.adoption.base), 0, 1);
      model.risk.base = clamp(+($("#sRisk")?.value || model.risk.base), 0, 1);

      Object.assign(model.sim.oneWay, {
        discountRate: model.time.discBase,
        years: model.time.years,
        adoption: model.adoption.base,
        risk: model.risk.base
      });

      logEvent("Updated base settings", {});
      renderSimulations();
      showToast("Base settings applied to current scenario.");
    });
  }

  // ----------------------------
  // Import / Export tab: schema validation report + strict apply + reset/restore
  // ----------------------------
  let stagedImport = null; // {fileName, sheetName, canonical, reportHTML, agg}

  async function parseExcelFile(file) {
    if (typeof XLSX === "undefined") {
      alert("Excel import requires the SheetJS (XLSX) library. Ensure the XLSX script is loaded.");
      return;
    }

    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array" });

    const sheetName = wb.SheetNames.find(n => DEFAULT_SHEET_CANDIDATES.includes(n)) || wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];
    const jsonRows = XLSX.utils.sheet_to_json(sheet, { defval: null });

    const validated = validateAndCanonicaliseRows(jsonRows);

    const reportEl = $("#excelReport");
    if (reportEl) reportEl.innerHTML = buildSchemaReportHTML(validated.report);

    if (!validated.ok) {
      stagedImport = null;
      logEvent("Excel validation failed", { file: file.name, sheet: sheetName, errors: validated.report.errors });
      showToast("Upload failed validation. See the error list.");
      return;
    }

    const agg = aggregateToTreatments(validated.canonical);
    stagedImport = {
      fileName: file.name,
      sheetName,
      canonical: validated.canonical,
      report: validated.report,
      agg
    };

    // Add calibration details into the report box (post-aggregation)
    const extra = {
      "Control identified as": agg.controlName,
      "Treatments found": String(agg.means.size),
      "Numeric columns used (cost/output)": String(agg.allVars.length)
    };
    validated.report.calibration = { ...(validated.report.calibration || {}), ...extra };
    if (reportEl) reportEl.innerHTML = buildSchemaReportHTML(validated.report);

    logEvent("Excel parsed and staged (validation passed)", { file: file.name, sheet: sheetName, control: agg.controlName });
    showToast("Excel parsed and validated. Ready to apply.");
  }

  function applyStagedImport() {
    if (!stagedImport) {
      alert("No validated upload staged. Please parse an Excel file first.");
      return;
    }

    try {
      applyAggregatedDataToModel(stagedImport.agg);
      model.data.currentSource = { kind: "upload", name: stagedImport.fileName, sheet: stagedImport.sheetName, timestamp: nowISO() };

      // Save last success for restore
      model.data.lastSuccessfulUpload = {
        fileName: stagedImport.fileName,
        sheetName: stagedImport.sheetName,
        timestamp: nowISO(),
        controlName: stagedImport.agg.controlName,
        canonical: {
          numericCols: stagedImport.canonical.numericCols,
          metaCols: stagedImport.canonical.metaCols,
          // keep the canonical rows (can be large; but user asked robustness; still OK for typical files)
          rows: stagedImport.canonical.rows
        }
      };

      try {
        localStorage.setItem(LS_LAST_SUCCESS, JSON.stringify(model.data.lastSuccessfulUpload));
      } catch {}

      stagedImport = null;

      renderAll();
      showToast("Upload applied successfully. Results updated.");
      logEvent("Applied Excel upload (success)", { source: model.data.currentSource });
    } catch (err) {
      console.error(err);
      alert("Upload could not be applied: " + (err?.message || "Unknown error"));
      logEvent("Failed applying staged upload", { error: String(err?.message || err) });
    }
  }

  function restoreLastSuccessfulUpload() {
    try {
      const raw = localStorage.getItem(LS_LAST_SUCCESS);
      if (!raw) {
        alert("No previous successful upload found in this browser.");
        return;
      }
      const saved = JSON.parse(raw);
      if (!saved || !saved.canonical || !Array.isArray(saved.canonical.rows)) {
        alert("Saved upload record is incomplete.");
        return;
      }

      const canonical = {
        rows: saved.canonical.rows,
        numericCols: saved.canonical.numericCols || [],
        metaCols: saved.canonical.metaCols || []
      };
      const agg = aggregateToTreatments(canonical);
      applyAggregatedDataToModel(agg);

      model.data.currentSource = { kind: "restore", name: saved.fileName || "Last upload", sheet: saved.sheetName, timestamp: nowISO() };
      model.data.lastSuccessfulUpload = saved;

      renderAll();
      showToast("Restored last successful upload.");
      logEvent("Restored last successful upload", { fileName: saved.fileName, sheetName: saved.sheetName });
    } catch (err) {
      console.error(err);
      alert("Restore failed: " + (err?.message || "Unknown error"));
      logEvent("Restore failed", { error: String(err?.message || err) });
    }
  }

  function resetToDefaultDataset() {
    // Create a canonical structure from DEFAULT_RAW_ROWS and apply it.
    const headers = Object.keys(DEFAULT_RAW_ROWS[0] || {});
    const colTreatment = findColumnByAliases(headers, EXCEL_SCHEMA.required.treatment) || "Amendment";
    const colYield = findColumnByAliases(headers, EXCEL_SCHEMA.required.yield) || "Yield t/ha";
    const otherCols = headers.filter(h => h !== colTreatment && h !== colYield);

    const canonical = {
      rows: DEFAULT_RAW_ROWS.map((r, i) => {
        const vars = {};
        for (const c of otherCols) {
          const n = parseNumberLoose(r[c]);
          vars[c] = Number.isNaN(n) ? null : n;
        }
        return {
          __rowIndex: i + 2,
          treatment: String(r[colTreatment] ?? "").trim(),
          yield_t_ha: parseNumberLoose(r[colYield]),
          isControl: /control/i.test(String(r[colTreatment] ?? "")) || String(r[colTreatment] ?? "").toLowerCase() === "control",
          areaHa: null,
          adoption: null,
          capital_y0: null,
          vars,
          meta: {}
        };
      }),
      numericCols: otherCols,
      metaCols: []
    };

    const agg = aggregateToTreatments(canonical);
    applyAggregatedDataToModel(agg);

    model.data.currentSource = { kind: "default", name: "Built-in default dataset", timestamp: nowISO() };

    renderAll();
    showToast("Reset to default dataset.");
    logEvent("Reset to default dataset", {});
  }

  function renderImportExport() {
    const root = $("#importBox");
    if (!root) return;

    root.innerHTML = `
      <div class="card">
        <div class="row space-between">
          <div>
            <h3>Excel-first workflow (schema as contract)</h3>
            <div class="small muted">
              Upload must fully pass validation or it will not be applied. The tool will auto-coerce common formats (currency symbols, commas, blanks).
              All numeric columns are used (classified as costs or outputs); non-numeric columns are kept as provenance and exported.
            </div>
          </div>
          <div class="row">
            <button class="btn" id="parseExcel">Parse & validate</button>
            <button class="btn" id="applyExcel">Apply upload</button>
          </div>
        </div>

        <div class="row">
          <div class="field">
            <label>Current data source</label>
            <div class="pill">${esc(model.data.currentSource.kind)} • ${esc(model.data.currentSource.name)} • ${esc(model.data.currentSource.timestamp)}</div>
          </div>
        </div>

        <div id="excelReport" class="mt"></div>

        <div class="row mt space-between">
          <div class="row">
            <button class="btn" id="downloadTemplate">Download Excel template</button>
            <button class="btn" id="exportExcel">Export Excel</button>
            <button class="btn" id="exportPdf">Print / PDF</button>
          </div>
          <div class="row">
            <button class="btn danger" id="resetDefault">Reset to default dataset</button>
            <button class="btn" id="restoreLast">Restore last successful upload</button>
          </div>
        </div>
      </div>
    `;

    $("#parseExcel")?.addEventListener("click", async () => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".xlsx,.xlsm,.xlsb";
      input.onchange = async e => {
        const file = e.target.files?.[0];
        if (!file) return;
        try {
          await parseExcelFile(file);
        } catch (err) {
          console.error(err);
          alert("Parse failed: " + (err?.message || "Unknown error"));
          logEvent("Excel parse error", { error: String(err?.message || err) });
        } finally {
          input.value = "";
        }
      };
      input.click();
    });

    $("#applyExcel")?.addEventListener("click", () => applyStagedImport());
    $("#resetDefault")?.addEventListener("click", () => resetToDefaultDataset());
    $("#restoreLast")?.addEventListener("click", () => restoreLastSuccessfulUpload());
    $("#exportExcel")?.addEventListener("click", () => exportExcel());
    $("#exportPdf")?.addEventListener("click", () => exportPdf());
    $("#downloadTemplate")?.addEventListener("click", () => downloadExcelTemplate());
  }

  function downloadExcelTemplate() {
    if (typeof XLSX === "undefined") {
      alert("Excel template generation requires the SheetJS (XLSX) library. Ensure the XLSX script is loaded.");
      return;
    }

    // Template is generated from schema; includes a ReadMe sheet.
    const wb = XLSX.utils.book_new();

    const readme = XLSX.utils.aoa_to_sheet([
      [`${TOOL.name} — Excel template (schema contract)`],
      [`Version: ${TOOL.version}`],
      [""],
      ["Required columns (at least these)"],
      ["- Amendment (or Treatment)"],
      ["- Yield t/ha (or Yield)"],
      [""],
      ["Optional columns"],
      ["- Area (ha)"],
      ["- Adoption"],
      ["- Capital cost (year 0)"],
      ["- IsControl (TRUE/FALSE)"],
      [""],
      ["All other numeric columns will be used automatically:"],
      ["- If a column name looks like a cost (labour, herbicide, fertiliser, fuel, etc.), it becomes an operating cost component ($/ha)."],
      ["- Otherwise, it becomes an output delta. Assign a $/unit value for that output inside the tool to monetise it."],
      [""],
      ["Reliability rule"],
      ["- Upload must fully validate or it will not be applied. You will see a clear error list."],
      [""],
      ["Recommended sheet name"],
      ["- Use 'FabaBeanRaw' or 'Data'."]
    ]);
    XLSX.utils.book_append_sheet(wb, readme, "ReadMe");

    const header = ["Amendment", "Yield t/ha", "Pre sowing Labour", "Treatment Input Cost Only /Ha"];
    const data = [header, ...DEFAULT_RAW_ROWS.map(r => [r.Amendment, r["Yield t/ha"], r["Pre sowing Labour"], r["Treatment Input Cost Only /Ha"]])];
    const wsData = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, wsData, "FabaBeanRaw");

    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile("farming_cba_template.xlsx", out, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Excel template downloaded.");
    logEvent("Downloaded Excel template", {});
  }

  // ----------------------------
  // Audit tab
  // ----------------------------
  function renderAudit() {
    const root = $("#auditBox");
    if (!root) return;

    const rows = model.audit.slice(0, 200);

    root.innerHTML = `
      <div class="card">
        <div class="row space-between">
          <div>
            <h3>Audit log</h3>
            <div class="small muted">Tracks file uploads, validation outcomes, assumption changes, exports, and focus actions.</div>
          </div>
          <div class="row">
            <button class="btn tiny" id="auditExport">Export audit JSON</button>
          </div>
        </div>

        <div class="table-wrap">
          <table class="mini">
            <thead>
              <tr><th>Time</th><th>Action</th><th>Details</th></tr>
            </thead>
            <tbody>
              ${
                rows.length
                  ? rows.map(a => `
                      <tr>
                        <td class="muted">${esc(a.time)}</td>
                        <td><strong>${esc(a.action)}</strong></td>
                        <td><code class="code">${esc(JSON.stringify(a.details || {}))}</code></td>
                      </tr>
                    `).join("")
                  : `<tr><td colspan="3" class="muted">No events yet.</td></tr>`
              }
            </tbody>
          </table>
        </div>
      </div>
    `;

    $("#auditExport")?.addEventListener("click", () => {
      downloadFile(`${slugify(model.project.name)}_audit.json`, JSON.stringify({ tool: TOOL, exportedAt: nowISO(), audit: model.audit }, null, 2), "application/json");
      showToast("Audit log exported.");
      logEvent("Exported audit log JSON", {});
    });
  }

  // ----------------------------
  // Tabs + accessibility
  // ----------------------------
  function switchTab(target) {
    if (!target) return;

    const tabs = $$("[data-tab-target]");
    const panels = $$(".tab-panel");

    tabs.forEach(t => {
      const on = t.dataset.tabTarget === target;
      t.classList.toggle("active", on);
      t.setAttribute("aria-selected", on ? "true" : "false");
      t.tabIndex = on ? 0 : -1;
    });

    panels.forEach(p => {
      const on = p.dataset.tabPanel === target;
      p.hidden = !on;
      p.classList.toggle("active", on);
    });

    // do not force scroll in print contexts
    if (!document.body.classList.contains("print-mode")) window.scrollTo({ top: 0, behavior: "smooth" });

    // Re-render on tab switches so panels are always live
    if (target === "results") renderResults();
    if (target === "simulations") renderSimulations();
    if (target === "ai") renderAI();
    if (target === "import") renderImportExport();
    if (target === "audit") renderAudit();
  }

  function initTabs() {
    document.addEventListener("click", e => {
      const btn = e.target.closest("[data-tab-target]");
      if (!btn) return;
      e.preventDefault();
      switchTab(btn.dataset.tabTarget);
    });

    // Keyboard navigation for tablist (WCAG-friendly)
    const tablist = $("#mainTabs");
    if (tablist) {
      tablist.addEventListener("keydown", e => {
        const tabs = $$("[data-tab-target]", tablist);
        const idx = tabs.findIndex(t => t.classList.contains("active"));
        if (idx < 0) return;

        if (e.key === "ArrowRight" || e.key === "ArrowLeft") {
          e.preventDefault();
          const next = e.key === "ArrowRight" ? (idx + 1) % tabs.length : (idx - 1 + tabs.length) % tabs.length;
          const target = tabs[next].dataset.tabTarget;
          switchTab(target);
          tabs[next].focus();
        }
      });
    }

    // Default open Results tab (snapshot usability)
    switchTab("results");
  }

  // ----------------------------
  // Results filter controls
  // ----------------------------
  function initResultsControls() {
    $("#filterTopNpv")?.addEventListener("click", () => {
      resultsFilterMode = "top5_npv";
      renderResults();
      showToast("Filter applied: Top 5 by NPV.");
      logEvent("Results filter changed", { mode: resultsFilterMode });
    });
    $("#filterTopBcr")?.addEventListener("click", () => {
      resultsFilterMode = "top5_bcr";
      renderResults();
      showToast("Filter applied: Top 5 by BCR.");
      logEvent("Results filter changed", { mode: resultsFilterMode });
    });
    $("#filterImprove")?.addEventListener("click", () => {
      resultsFilterMode = "improve_only";
      renderResults();
      showToast("Filter applied: Only improvements vs control.");
      logEvent("Results filter changed", { mode: resultsFilterMode });
    });
    $("#filterAll")?.addEventListener("click", () => {
      resultsFilterMode = "all";
      renderResults();
      showToast("Filter applied: Show all.");
      logEvent("Results filter changed", { mode: resultsFilterMode });
    });

    $("#resultsExportExcel")?.addEventListener("click", () => exportExcel());
    $("#resultsPrint")?.addEventListener("click", () => exportPdf());

    $("#calcDrawerClose")?.addEventListener("click", () => closeCalcDrawer());
    $("#calcDrawerBackdrop")?.addEventListener("click", () => closeCalcDrawer());
    document.addEventListener("keydown", e => {
      if (e.key === "Escape") closeCalcDrawer();
    });
  }

  // ----------------------------
  // Render all
  // ----------------------------
  function renderAll() {
    renderProject();
    renderSettings();
    renderOutputs();
    renderTreatments();
    renderResults();
    renderImportExport();
    renderAudit();
  }

  // ----------------------------
  // Init
  // ----------------------------
  function initFromDefault() {
    resetToDefaultDataset();
    loadSavedScenarios();

    // Try restore last success automatically if present (but do not force)
    // Users can click restore explicitly; auto-restore can be surprising.
    logEvent("Tool loaded", { version: TOOL.version });
  }

  function init() {
    initTabs();
    initResultsControls();
    initFromDefault();
    renderAll();
    showToast(`${TOOL.name} loaded (v${TOOL.version}).`);
  }

  // ----------------------------
  // Start
  // ----------------------------
  document.addEventListener("DOMContentLoaded", init);
})();
