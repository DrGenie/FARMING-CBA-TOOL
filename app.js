// app.js
// Farming CBA Decision Tool 2
// Excel-first, schema-validated import; Control vs Treatments comparison; auditable formulas & reconciliation;
// cost build-up with computed totals; simulations with break-even + scenario sets; AI prompt pack; exports; audit trail.

(() => {
  "use strict";

  const TOOL_VERSION = "2.0.0";
  const STORAGE_KEYS = {
    lastUpload: "farming_cba2_last_successful_upload_v1",
    scenarios: "farming_cba2_saved_scenarios_v1",
    audit: "farming_cba2_audit_v1",
    printMode: "farming_cba2_print_mode_v1"
  };

  // ---------- DOM ----------
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));

  // ---------- UX ----------
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

  function esc(s) {
    return (s ?? "")
      .toString()
      .replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const isFiniteNum = v => Number.isFinite(v) && !Number.isNaN(v);

  function fmt(n, digits = 2) {
    if (!isFiniteNum(n)) return "n/a";
    return n.toLocaleString(undefined, { maximumFractionDigits: digits });
  }
  function money(n) {
    if (!isFiniteNum(n)) return "n/a";
    const abs = Math.abs(n);
    const digits = abs >= 1000 ? 0 : 2;
    return "$" + n.toLocaleString(undefined, { maximumFractionDigits: digits });
  }
  function percent(n) {
    if (!isFiniteNum(n)) return "n/a";
    return fmt(n, 2) + "%";
  }

  // ---------- AUDIT ----------
  let auditLog = [];
  function audit(action, details) {
    const item = {
      ts: new Date().toISOString(),
      action: String(action || "").slice(0, 120),
      details: String(details || "").slice(0, 500)
    };
    auditLog.unshift(item);
    auditLog = auditLog.slice(0, 500);
    try {
      localStorage.setItem(STORAGE_KEYS.audit, JSON.stringify(auditLog));
    } catch {}
    renderAudit();
  }

  function loadAudit() {
    try {
      const raw = localStorage.getItem(STORAGE_KEYS.audit);
      if (raw) auditLog = JSON.parse(raw) || [];
    } catch {
      auditLog = [];
    }
  }

  function renderAudit() {
    const tbody = $("#auditTable tbody");
    if (!tbody) return;
    tbody.innerHTML = auditLog
      .map(
        r => `
        <tr>
          <td class="small">${esc(new Date(r.ts).toLocaleString())}</td>
          <td>${esc(r.action)}</td>
          <td class="small muted">${esc(r.details)}</td>
        </tr>`
      )
      .join("");
  }

  // ---------- MODEL ----------
  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

  // Default dataset (one row per treatment; in real uploads, may be replicated plot rows)
  const DEFAULT_ROWS = [
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

  const model = {
    toolName: "Farming CBA Decision Tool 2",
    project: {
      name: "Faba bean soil amendment trial",
      organisation: "Newcastle Business School, The University of Newcastle",
      lastUpdated: new Date().toISOString().slice(0, 10)
    },
    settings: {
      farmAreaHa: 100,
      years: 10,
      discountRatePct: 7,
      adoption: 0.9,
      risk: 0.15,
      reconTolerance: 10
    },
    outputs: [
      { id: "out_yield", name: "Grain yield", unit: "t/ha", unitValue: 450, source: "Input Directly" }
    ],
    treatments: [],
    importMeta: {
      active: { source: "default", fileName: "Default dataset", sheetName: "Default", rowsRead: DEFAULT_ROWS.length, appliedAt: new Date().toISOString() },
      mapping: null,
      columnStats: null,
      aggregated: null
    }
  };

  // ---------- TABS + ACCESSIBILITY ----------
  function switchTab(target) {
    if (!target) return;
    $$("[data-tab]").forEach(btn => {
      const active = btn.dataset.tab === target;
      btn.classList.toggle("active", active);
      btn.setAttribute("aria-selected", active ? "true" : "false");
      btn.tabIndex = active ? 0 : -1;
    });
    $$(".tab-panel").forEach(p => {
      const key = p.dataset.tabPanel;
      const show = key === target;
      p.classList.toggle("active", show);
      p.hidden = !show;
      p.setAttribute("aria-hidden", show ? "false" : "true");
    });
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function initTabs() {
    document.addEventListener("click", e => {
      const jump = e.target.closest("[data-tab-jump]");
      if (jump) {
        e.preventDefault();
        switchTab(jump.dataset.tabJump);
        return;
      }
      const btn = e.target.closest("[data-tab]");
      if (!btn) return;
      e.preventDefault();
      switchTab(btn.dataset.tab);
    });

    // Keyboard navigation
    $(".tabs")?.addEventListener("keydown", e => {
      const tabs = $$("[data-tab]");
      const idx = tabs.findIndex(t => t.classList.contains("active"));
      if (idx < 0) return;
      if (e.key === "ArrowRight" || e.key === "ArrowLeft") {
        e.preventDefault();
        const next = e.key === "ArrowRight" ? (idx + 1) % tabs.length : (idx - 1 + tabs.length) % tabs.length;
        tabs[next].focus();
        switchTab(tabs[next].dataset.tab);
      }
    });
  }

  // ---------- MATH ----------
  const annuityFactor = (N, rPct) => {
    const r = rPct / 100;
    return r === 0 ? N : (1 - Math.pow(1 + r, -N)) / r;
  };

  // Robust numeric parsing (currency, commas, blanks, parentheses)
  function parseNumber(value) {
    if (value === null || value === undefined || value === "") return NaN;
    if (typeof value === "number") return value;
    let s = String(value).trim();
    if (!s) return NaN;
    // parentheses negative
    const neg = /^\(.*\)$/.test(s);
    if (neg) s = s.slice(1, -1);
    // strip currency symbols and commas
    s = s.replace(/[$£€,\s]/g, "");
    // percentages (interpret as numeric percent, not fraction)
    if (/%$/.test(s)) s = s.replace(/%$/, "");
    const n = parseFloat(s);
    if (!Number.isFinite(n)) return NaN;
    return neg ? -n : n;
  }

  function canonCol(s) {
    return String(s || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ")
      .replace(/[^a-z0-9 %/()._-]+/g, "");
  }

  // ---------- IMPORT SCHEMA ----------
  // Roles:
  // - treatment: treatment identifier/name
  // - output: output level (e.g., yield)
  // - cost_labour: cost component
  // - cost_operating: cost component
  // - ignore: (allowed but discouraged)
  const ROLE = {
    TREATMENT: "treatment",
    OUTPUT: "output",
    COST_LABOUR: "cost_labour",
    COST_OPERATING: "cost_operating",
    IGNORE: "ignore"
  };

  const REQUIRED = {
    treatment: {
      label: "Treatment name",
      aliases: ["amendment", "treatment", "treatment name", "name", "option", "arm"]
    },
    yield: {
      label: "Yield (t/ha)",
      aliases: ["yield t/ha", "yield", "grain yield", "yield (t/ha)", "yield t ha", "t/ha"]
    }
  };

  function detectMappingFromHeaders(headers) {
    const hcanon = headers.map(h => ({ raw: h, canon: canonCol(h) }));
    let treatmentCol = null;
    let yieldCol = null;

    for (const h of hcanon) {
      if (!treatmentCol && REQUIRED.treatment.aliases.some(a => h.canon === canonCol(a))) treatmentCol = h.raw;
      if (!yieldCol && REQUIRED.yield.aliases.some(a => h.canon === canonCol(a))) yieldCol = h.raw;
    }

    const roles = {};
    headers.forEach(h => {
      const c = canonCol(h);
      if (treatmentCol && h === treatmentCol) roles[h] = { role: ROLE.TREATMENT, outputId: null };
      else if (yieldCol && h === yieldCol) roles[h] = { role: ROLE.OUTPUT, outputId: "out_yield" };
      else {
        // heuristic: labour vs operating
        if (c.includes("labour")) roles[h] = { role: ROLE.COST_LABOUR, outputId: null };
        else roles[h] = { role: ROLE.COST_OPERATING, outputId: null };
      }
    });

    // If treatment/yield not found, attempt fuzzy contains matches
    if (!treatmentCol) {
      const cand = headers.find(h => canonCol(h).includes("amendment") || canonCol(h).includes("treatment"));
      if (cand) roles[cand] = { role: ROLE.TREATMENT, outputId: null };
    }
    if (!yieldCol) {
      const cand = headers.find(h => canonCol(h).includes("yield"));
      if (cand) roles[cand] = { role: ROLE.OUTPUT, outputId: "out_yield" };
    }

    return roles;
  }

  function buildColumnStats(rows, headers, roles) {
    const stats = {};
    headers.forEach(h => {
      let ok = 0, bad = 0, blank = 0;
      rows.forEach(r => {
        const v = r[h];
        if (v === null || v === undefined || v === "") { blank++; return; }
        if (roles[h]?.role === ROLE.TREATMENT) { ok++; return; }
        const n = parseNumber(v);
        if (Number.isNaN(n)) bad++; else ok++;
      });
      const total = rows.length || 1;
      stats[h] = { ok, bad, blank, rate: ok / total };
    });
    return stats;
  }

  function validateImport(rows, headers, roles) {
    const errors = [];

    const treatmentCols = headers.filter(h => roles[h]?.role === ROLE.TREATMENT);
    const yieldCols = headers.filter(h => roles[h]?.role === ROLE.OUTPUT && roles[h]?.outputId === "out_yield");
    if (!treatmentCols.length) errors.push("Missing required treatment column (e.g., 'Amendment' or 'Treatment').");
    if (!yieldCols.length) errors.push("Missing required yield column (e.g., 'Yield t/ha').");

    // Check for empty treatment names
    const tcol = treatmentCols[0];
    if (tcol) {
      const empties = rows.filter(r => !String(r[tcol] ?? "").trim()).length;
      if (empties > 0) errors.push(`Found ${empties} rows with blank treatment names in column "${tcol}".`);
    }

    // Numeric parse success thresholds for numeric columns (exclude treatment col)
    const colStats = buildColumnStats(rows, headers, roles);
    headers.forEach(h => {
      const role = roles[h]?.role;
      if (role === ROLE.TREATMENT) return;
      // If parse success is low, report
      const st = colStats[h];
      if (st && st.bad > 0 && st.rate < 0.9) {
        errors.push(`Column "${h}" has many non-numeric cells (${st.bad}/${rows.length}). Check formatting.`);
      }
    });

    return { ok: errors.length === 0, errors, colStats };
  }

  function aggregateRows(rows, headers, roles) {
    const tcol = headers.find(h => roles[h]?.role === ROLE.TREATMENT);
    if (!tcol) return null;

    // Build outputs list from roles (always include out_yield)
    const outputCols = headers.filter(h => roles[h]?.role === ROLE.OUTPUT);
    const costLabCols = headers.filter(h => roles[h]?.role === ROLE.COST_LABOUR);
    const costOpCols = headers.filter(h => roles[h]?.role === ROLE.COST_OPERATING);

    const groups = new Map();
    rows.forEach(r => {
      const name = String(r[tcol] ?? "").trim();
      if (!name) return;

      if (!groups.has(name)) {
        groups.set(name, {
          name,
          n: 0,
          sums: {},
          counts: {}
        });
      }
      const g = groups.get(name);
      g.n++;

      // numeric aggregation for every non-treatment column (no omissions)
      headers.forEach(h => {
        if (h === tcol) return;
        const role = roles[h]?.role;
        if (role === ROLE.IGNORE) return;

        // For OUTPUT and COST roles, try parseNumber; blank -> treated as 0 but counted
        const n = parseNumber(r[h]);
        const val = Number.isNaN(n) ? 0 : n;

        g.sums[h] = (g.sums[h] || 0) + val;
        g.counts[h] = (g.counts[h] || 0) + 1;
      });
    });

    const treatments = [];
    for (const [name, g] of groups.entries()) {
      const means = {};
      Object.keys(g.sums).forEach(h => {
        const c = g.counts[h] || 1;
        means[h] = g.sums[h] / c;
      });

      treatments.push({
        name,
        nReplicates: g.n,
        outputLevels: Object.fromEntries(outputCols.map(h => [h, means[h] ?? 0])),
        labourCosts: Object.fromEntries(costLabCols.map(h => [h, means[h] ?? 0])),
        operatingCosts: Object.fromEntries(costOpCols.map(h => [h, means[h] ?? 0]))
      });
    }

    // Detect control by name
    const control = treatments.find(t => canonCol(t.name).includes("control"));
    return {
      treatmentCol: tcol,
      outputCols,
      costLabCols,
      costOpCols,
      treatments,
      controlName: control ? control.name : null
    };
  }

  // ---------- APPLY IMPORT TO MODEL ----------
  function ensureOutputsFromMapping(agg) {
    // Always keep out_yield; add additional outputs if user classified more as OUTPUT
    // In this build, we store outputs by "Excel column name", but map to output ids.
    // We keep output id stable by slugging the column name.
    const existing = new Map(model.outputs.map(o => [o.id, o]));
    // Guarantee yield exists
    if (!existing.has("out_yield")) {
      model.outputs.unshift({ id: "out_yield", name: "Grain yield", unit: "t/ha", unitValue: 450, source: "Input Directly" });
    }

    // Add any non-yield output columns as separate outputs
    (agg.outputCols || []).forEach(col => {
      // col might be yield too; skip if mapped to out_yield by override logic
      const isYield = canonCol(col).includes("yield") || canonCol(col) === canonCol("Yield t/ha");
      if (isYield) return;

      const id = "out_" + canonCol(col).replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 32);
      if (!existing.has(id)) {
        model.outputs.push({
          id,
          name: col,
          unit: "unit/ha",
          unitValue: 0,
          source: "Imported"
        });
      }
    });
  }

  function applyAggregatedToModel(agg, meta) {
    ensureOutputsFromMapping(agg);

    // Determine which column is yield in the aggregated outputLevels: best match using aliases
    const yieldCol = (agg.outputCols || []).find(c => REQUIRED.yield.aliases.some(a => canonCol(c) === canonCol(a))) ||
                     (agg.outputCols || []).find(c => canonCol(c).includes("yield")) ||
                     (agg.outputCols || [])[0];

    const controlName = agg.controlName || agg.treatments[0]?.name || "Control";
    const farmArea = Number(model.settings.farmAreaHa) || 0;

    // Create treatments with cost build-up line items
    model.treatments = agg.treatments.map(t => {
      const isControl = t.name === controlName;
      const opItems = [];

      // labour costs as items
      Object.keys(t.labourCosts || {}).forEach(k => {
        opItems.push({ id: uid(), label: k, valuePerHa: Number(t.labourCosts[k]) || 0, category: "Labour" });
      });

      // operating costs as items
      Object.keys(t.operatingCosts || {}).forEach(k => {
        opItems.push({ id: uid(), label: k, valuePerHa: Number(t.operatingCosts[k]) || 0, category: "Operating" });
      });

      // Output levels (per ha). Map yield column to out_yield and other output cols to their output ids by name.
      const outputsPerHa = {};
      model.outputs.forEach(o => { outputsPerHa[o.id] = 0; });

      // Yield
      outputsPerHa["out_yield"] = Number((t.outputLevels || {})[yieldCol]) || 0;

      // Additional outputs (if any output cols were reclassified). Match by column name to output id slug.
      (agg.outputCols || []).forEach(col => {
        if (col === yieldCol) return;
        const id = "out_" + canonCol(col).replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 32);
        if (outputsPerHa[id] !== undefined) outputsPerHa[id] = Number((t.outputLevels || {})[col]) || 0;
      });

      return {
        id: uid(),
        name: t.name,
        isControl,
        areaHa: farmArea,
        adoption: 1,
        capitalCostY0: 0,
        opCostItems: opItems, // $/ha line items
        outputsPerHa,          // per ha levels
        nReplicates: t.nReplicates || 1,
        notes: ""
      };
    });

    // Ensure exactly one control; if none detected, force first as control
    if (!model.treatments.some(t => t.isControl) && model.treatments.length) {
      model.treatments[0].isControl = true;
    } else if (model.treatments.filter(t => t.isControl).length > 1) {
      const first = model.treatments.find(t => t.isControl);
      model.treatments.forEach(t => { t.isControl = (t === first); });
    }

    model.importMeta.active = meta;
    model.importMeta.aggregated = agg;
    audit("Import applied", `${meta.fileName} · sheet=${meta.sheetName} · rows=${meta.rowsRead} · treatments=${model.treatments.length}`);

    // Persist last successful upload meta for restore
    try {
      localStorage.setItem(STORAGE_KEYS.lastUpload, JSON.stringify({
        meta,
        agg,
        outputs: model.outputs
      }));
    } catch {}

    renderAll();
    showToast("Excel import applied. Results updated.");
  }

  function restoreLastSuccessfulUpload() {
    try {
      const raw = localStorage.getItem(STORAGE_KEYS.lastUpload);
      if (!raw) {
        showToast("No saved upload found in this browser.");
        return;
      }
      const obj = JSON.parse(raw);
      if (!obj?.agg || !obj?.meta) {
        showToast("Saved upload is invalid.");
        return;
      }
      if (Array.isArray(obj.outputs) && obj.outputs.length) model.outputs = obj.outputs;
      applyAggregatedToModel(obj.agg, obj.meta);
      audit("Restore last upload", obj.meta.fileName || "unknown");
    } catch {
      showToast("Unable to restore last upload.");
    }
  }

  // ---------- RESULTS CALCULATIONS ----------
  function computeTreatmentTotals(t) {
    const opPerHa = (t.opCostItems || []).reduce((s, it) => s + (Number(it.valuePerHa) || 0), 0);
    const opFarm = opPerHa * (Number(t.areaHa) || 0);
    return { opPerHa, opFarm };
  }

  function computeAnnualBenefit(t, outputsOverride = null, unitValuesOverride = null, adoptionOverride = null, riskOverride = null) {
    const area = Number(t.areaHa) || 0;
    const adopt = clamp(adoptionOverride ?? model.settings.adoption, 0, 1);
    const risk = clamp(riskOverride ?? model.settings.risk, 0, 1);

    const outputsPerHa = outputsOverride || t.outputsPerHa || {};
    let perHaValue = 0;
    model.outputs.forEach(o => {
      const level = Number(outputsPerHa[o.id]) || 0;
      const unitValue = Number((unitValuesOverride && unitValuesOverride[o.id] != null) ? unitValuesOverride[o.id] : o.unitValue) || 0;
      perHaValue += level * unitValue;
    });

    return perHaValue * area * adopt * (1 - risk);
  }

  function computeMetricsForTreatment(t, scenario = null) {
    const years = Number(scenario?.years ?? model.settings.years) || 1;
    const r = Number(scenario?.discountRatePct ?? model.settings.discountRatePct) || 0;
    const adopt = clamp(scenario?.adoption ?? model.settings.adoption, 0, 1);
    const risk = clamp(scenario?.risk ?? model.settings.risk, 0, 1);

    const { opFarm } = computeTreatmentTotals(t);
    const cap = Number(t.capitalCostY0) || 0;

    const annualBenefit = computeAnnualBenefit(t, scenario?.outputsOverride || null, scenario?.unitValuesOverride || null, adopt, risk);
    const annualCost = opFarm; // per-year operating

    const af = annuityFactor(years, r);
    const pvBenefits = annualBenefit * af;
    const pvCosts = cap + annualCost * af;
    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts > 0 ? pvBenefits / pvCosts : NaN;
    const roi = pvCosts > 0 ? (npv / pvCosts) * 100 : NaN;

    return { pvBenefits, pvCosts, npv, bcr, roi, annualBenefit, annualCost, capY0: cap, years, r, adopt, risk };
  }

  function computeAllMetrics(scenario = null) {
    const control = model.treatments.find(t => t.isControl) || model.treatments[0] || null;
    const list = model.treatments.map(t => ({ t, m: computeMetricsForTreatment(t, scenario) }));

    const controlEntry = control ? list.find(x => x.t === control) : null;
    const c = controlEntry ? controlEntry.m : null;

    // Rank by NPV (descending), excluding control from ranking (but still displayed)
    const ranked = list
      .filter(x => !x.t.isControl)
      .slice()
      .sort((a, b) => {
        const A = isFiniteNum(a.m.npv) ? a.m.npv : -Infinity;
        const B = isFiniteNum(b.m.npv) ? b.m.npv : -Infinity;
        return B - A;
      });

    const rankMap = new Map();
    ranked.forEach((x, i) => rankMap.set(x.t.id, i + 1));

    const withDelta = list.map(x => {
      const d = {};
      if (c) {
        d.pvBenefits = x.m.pvBenefits - c.pvBenefits;
        d.pvCosts = x.m.pvCosts - c.pvCosts;
        d.npv = x.m.npv - c.npv;
        d.bcr = (isFiniteNum(x.m.bcr) && isFiniteNum(c.bcr)) ? (x.m.bcr - c.bcr) : NaN;
        d.roi = (isFiniteNum(x.m.roi) && isFiniteNum(c.roi)) ? (x.m.roi - c.roi) : NaN;

        // % deltas where meaningful (avoid divide by zero; if control is 0, leave null)
        const pct = (val, base) => (isFiniteNum(base) && Math.abs(base) > 1e-12) ? (val / base) * 100 : null;
        d.pctBenefits = pct(d.pvBenefits, c.pvBenefits);
        d.pctCosts = pct(d.pvCosts, c.pvCosts);
        d.pctNpv = pct(d.npv, c.npv);
      }
      return {
        t: x.t,
        m: x.m,
        rank: x.t.isControl ? null : (rankMap.get(x.t.id) || null),
        delta: d
      };
    });

    return { control, controlMetrics: c, rows: withDelta };
  }

  // ---------- RESULTS UI (LEADERBOARD + COMPARISON TABLE) ----------
  const RESULTS_FILTER = { mode: "all" }; // all | topnpv | topbcr | improve
  let focusedTreatmentId = null;

  function applyFilter(rows) {
    const controlRow = rows.find(r => r.t.isControl);
    const others = rows.filter(r => !r.t.isControl);

    if (RESULTS_FILTER.mode === "topnpv") {
      const sorted = others.slice().sort((a, b) => (b.m.npv - a.m.npv));
      return [controlRow, ...sorted.slice(0, 5)].filter(Boolean);
    }
    if (RESULTS_FILTER.mode === "topbcr") {
      const sorted = others.slice().sort((a, b) => {
        const A = isFiniteNum(a.m.bcr) ? a.m.bcr : -Infinity;
        const B = isFiniteNum(b.m.bcr) ? b.m.bcr : -Infinity;
        return B - A;
      });
      return [controlRow, ...sorted.slice(0, 5)].filter(Boolean);
    }
    if (RESULTS_FILTER.mode === "improve") {
      const improved = others.filter(r => isFiniteNum(r.delta?.npv) && r.delta.npv > 0);
      return [controlRow, ...improved].filter(Boolean);
    }
    return rows;
  }

  function renderLeaderboard(allRows) {
    const tbody = $("#leaderboard tbody");
    if (!tbody) return;

    const control = allRows.find(r => r.t.isControl);
    const others = allRows.filter(r => !r.t.isControl);

    const sorted = others.slice().sort((a, b) => {
      const A = isFiniteNum(a.m.npv) ? a.m.npv : -Infinity;
      const B = isFiniteNum(b.m.npv) ? b.m.npv : -Infinity;
      return B - A;
    });

    tbody.innerHTML = sorted
      .map(r => {
        const deltaNpv = control ? (r.m.npv - control.m.npv) : NaN;
        return `
          <tr data-focus="${esc(r.t.id)}" class="${focusedTreatmentId === r.t.id ? "focus" : ""}">
            <td class="center">${r.rank ?? ""}</td>
            <td><b>${esc(r.t.name)}</b></td>
            <td class="num">${money(deltaNpv)}</td>
            <td class="num">${isFiniteNum(r.m.bcr) ? fmt(r.m.bcr, 2) : "n/a"}</td>
            <td class="num">${money(r.m.pvCosts)}</td>
            <td class="num">${money(r.m.pvBenefits)}</td>
          </tr>`;
      })
      .join("");

    tbody.onclick = e => {
      const tr = e.target.closest("tr[data-focus]");
      if (!tr) return;
      const id = tr.dataset.focus;
      focusedTreatmentId = (focusedTreatmentId === id) ? null : id;
      renderResults(); // re-render to apply highlight
      setTimeout(() => focusTreatmentColumn(id), 0);
    };
  }

  function indicatorDefinitions() {
    const years = Number(model.settings.years) || 1;
    const r = Number(model.settings.discountRatePct) || 0;
    return [
      {
        key: "pvBenefits",
        label: "PV Benefits",
        helpPlain:
          "PV Benefits is the discounted sum of annual benefits over the time horizon. In this tool, annual benefits are computed from per-hectare output levels (e.g., yield) multiplied by unit values (e.g., grain price), then multiplied by farm area, adoption, and (1 − risk).",
        helpMath:
          `AnnualBenefits = (Σ_o [ Output_o(per ha) × UnitValue_o ]) × Area × Adoption × (1 − Risk)\n` +
          `PV(Benefits) = Σ_{t=1..${years}} AnnualBenefits / (1 + r)^t, where r = ${r}%`,
        recon: "pvBenefits"
      },
      {
        key: "pvCosts",
        label: "PV Costs",
        helpPlain:
          "PV Costs is the discounted sum of annual operating costs plus any year-0 capital cost. Operating costs are computed from the cost build-up line-items ($/ha) times farm area.",
        helpMath:
          `AnnualCosts = (Σ_k CostItem_k($/ha)) × Area\n` +
          `PV(Costs) = CapitalCost(Y0) + Σ_{t=1..${years}} AnnualCosts / (1 + r)^t`,
        recon: "pvCosts"
      },
      {
        key: "npv",
        label: "NPV",
        helpPlain: "NPV is PV(Benefits) minus PV(Costs).",
        helpMath: "NPV = PV(Benefits) − PV(Costs)",
        recon: "npv"
      },
      {
        key: "bcr",
        label: "BCR",
        helpPlain: "BCR is PV(Benefits) divided by PV(Costs).",
        helpMath: "BCR = PV(Benefits) ÷ PV(Costs)",
        recon: "bcr"
      },
      {
        key: "roi",
        label: "ROI",
        helpPlain: "ROI is NPV divided by PV(Costs), expressed as a percentage.",
        helpMath: "ROI (%) = 100 × NPV ÷ PV(Costs)",
        recon: "roi"
      },
      {
        key: "rank",
        label: "Rank",
        helpPlain: "Rank orders treatments (excluding control) by NPV (highest to lowest).",
        helpMath: "Rank = sort by NPV descending",
        recon: "rank"
      },
      {
        key: "deltaNpv",
        label: "ΔNPV vs Control",
        helpPlain: "ΔNPV is NPV(treatment) minus NPV(control).",
        helpMath: "ΔNPV = NPV_t − NPV_control",
        recon: "deltaNpv"
      },
      {
        key: "deltaPvCost",
        label: "ΔPV Cost vs Control",
        helpPlain: "ΔPV Cost is PV Costs(treatment) minus PV Costs(control).",
        helpMath: "ΔPV Cost = PV(Costs)_t − PV(Costs)_control",
        recon: "deltaPvCost"
      }
    ];
  }

  function renderComparisonTable(allRows) {
    const table = $("#comparisonTable");
    if (!table) return;

    const filtered = applyFilter(allRows);
    const control = filtered.find(r => r.t.isControl) || allRows.find(r => r.t.isControl) || null;

    // Build header: sticky indicator col + control col + per-treatment group (value, Δ, Δ%)
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");
    thead.innerHTML = "";
    tbody.innerHTML = "";

    const treatments = filtered.filter(r => r.t); // includes control
    const nonControl = treatments.filter(r => !r.t.isControl);

    // Header row
    const tr1 = document.createElement("tr");
    tr1.innerHTML = `<th class="sticky-left">Indicator</th>` +
      `<th class="control-col">Control (baseline)</th>` +
      nonControl
        .map(r => {
          const focused = focusedTreatmentId === r.t.id ? " highlight" : "";
          return `
            <th class="${focused}" colspan="3" data-colgroup="${esc(r.t.id)}">
              ${esc(r.t.name)}
            </th>`;
        })
        .join("");
    thead.appendChild(tr1);

    // Subheader row (value / Δ / Δ%)
    const tr2 = document.createElement("tr");
    tr2.innerHTML = `<th class="sticky-left"> </th>` +
      `<th class="control-col">Value</th>` +
      nonControl
        .map(r => {
          const focused = focusedTreatmentId === r.t.id ? " highlight" : "";
          return `
            <th class="${focused}">Value</th>
            <th class="${focused}">Δ vs control</th>
            <th class="${focused}">Δ%</th>`;
        })
        .join("");
    thead.appendChild(tr2);

    const defs = indicatorDefinitions();

    defs.forEach(def => {
      const tr = document.createElement("tr");

      const indicatorCell = document.createElement("td");
      indicatorCell.className = "sticky-left body";
      indicatorCell.innerHTML = `<button class="linklike" data-ind="${esc(def.key)}" type="button">${esc(def.label)}</button>`;
      tr.appendChild(indicatorCell);

      // Control value
      const cVal = renderCellValue(def.key, control, control);
      const tdC = document.createElement("td");
      tdC.className = "num control-col";
      tdC.innerHTML = cVal;
      tr.appendChild(tdC);

      // Each treatment group
      nonControl.forEach(r => {
        const valHtml = renderCellValue(def.key, r, control);
        const delHtml = renderCellDelta(def.key, r, control);
        const pctHtml = renderCellPct(def.key, r, control);

        const cls = focusedTreatmentId === r.t.id ? "num highlight" : "num";

        const tdV = document.createElement("td");
        tdV.className = cls;
        tdV.innerHTML = valHtml;

        const tdD = document.createElement("td");
        tdD.className = cls;
        tdD.innerHTML = `<span class="delta">${delHtml}</span>`;

        const tdP = document.createElement("td");
        tdP.className = cls;
        tdP.innerHTML = `<span class="delta">${pctHtml}</span>`;

        tr.appendChild(tdV);
        tr.appendChild(tdD);
        tr.appendChild(tdP);
      });

      tbody.appendChild(tr);
    });

    // Click indicator to open drawer
    tbody.onclick = e => {
      const btn = e.target.closest("button[data-ind]");
      if (!btn) return;
      openCalcDrawer(btn.dataset.ind, allRows);
    };
  }

  function renderCellValue(key, row, controlRow) {
    if (!row) return "—";
    if (key === "rank") return row.rank != null ? String(row.rank) : (row.t.isControl ? "Baseline" : "");
    if (key === "deltaNpv") {
      if (!controlRow || row.t.isControl) return row.t.isControl ? "0" : "n/a";
      return money(row.m.npv - controlRow.m.npv);
    }
    if (key === "deltaPvCost") {
      if (!controlRow || row.t.isControl) return row.t.isControl ? "0" : "n/a";
      return money(row.m.pvCosts - controlRow.m.pvCosts);
    }
    const v = row.m[key];
    if (key === "bcr") return isFiniteNum(v) ? fmt(v, 2) : "n/a";
    if (key === "roi") return isFiniteNum(v) ? percent(v) : "n/a";
    if (key === "npv" || key === "pvBenefits" || key === "pvCosts") return money(v);
    return isFiniteNum(v) ? fmt(v) : "n/a";
  }

  function renderCellDelta(key, row, controlRow) {
    if (!row || !controlRow) return "—";
    if (row.t.isControl) return "—";
    if (key === "rank") return "—";
    if (key === "deltaNpv") return "—";
    if (key === "deltaPvCost") return "—";

    const c = controlRow.m;
    const v = row.m[key];

    if (key === "bcr" || key === "roi") {
      const dv = (isFiniteNum(v) && isFiniteNum(c[key])) ? (v - c[key]) : NaN;
      return isFiniteNum(dv) ? fmt(dv, 2) : "n/a";
    }
    if (key === "npv" || key === "pvBenefits" || key === "pvCosts") {
      const dv = (isFiniteNum(v) && isFiniteNum(c[key])) ? (v - c[key]) : NaN;
      return isFiniteNum(dv) ? money(dv) : "n/a";
    }
    const dv = (isFiniteNum(v) && isFiniteNum(c[key])) ? (v - c[key]) : NaN;
    return isFiniteNum(dv) ? fmt(dv) : "n/a";
  }

  function renderCellPct(key, row, controlRow) {
    if (!row || !controlRow) return "—";
    if (row.t.isControl) return "—";
    if (key === "rank" || key === "deltaNpv" || key === "deltaPvCost") return "—";

    const c = controlRow.m;
    const v = row.m[key];
    const base = c[key];

    if (!isFiniteNum(v) || !isFiniteNum(base) || Math.abs(base) < 1e-12) return "n/a";
    const pct = ((v - base) / base) * 100;
    return percent(pct);
  }

  function focusTreatmentColumn(tid) {
    const table = $("#comparisonTable");
    if (!table) return;
    // Remove highlight classes
    $$(`#comparisonTable [data-colgroup]`).forEach(th => th.classList.remove("highlight"));
    // Add highlight to the colgroup header
    const th = table.querySelector(`[data-colgroup="${CSS.escape(tid)}"]`);
    if (th) th.classList.add("highlight");

    // Scroll to that header if possible
    const wrap = $("#comparisonWrap");
    if (wrap && th) {
      const left = th.offsetLeft;
      wrap.scrollTo({ left: Math.max(0, left - 120), behavior: "smooth" });
    }
  }

  // ---------- CALC DETAILS DRAWER + RECONCILIATION ----------
  function openCalcDrawer(indicatorKey, allRows) {
    const defs = indicatorDefinitions();
    const def = defs.find(d => d.key === indicatorKey);
    if (!def) return;

    const overlay = $("#drawerOverlay");
    const drawer = $("#calcDrawer");
    if (!drawer || !overlay) return;

    $("#drawerTitle").textContent = `Calculation Details: ${def.label}`;
    $("#drawerSubtitle").textContent = `Applies across treatments using the current Settings.`;
    $("#formulaPlain").textContent = def.helpPlain;
    $("#formulaMath").textContent = def.helpMath;

    renderReconTable(indicatorKey, allRows);

    overlay.hidden = false;
    drawer.hidden = false;
    drawer.setAttribute("aria-hidden", "false");
    overlay.onclick = closeDrawer;
    $("#drawerClose").onclick = closeDrawer;
    document.addEventListener("keydown", escCloseDrawer, { once: true });
  }

  function escCloseDrawer(e) {
    if (e.key === "Escape") closeDrawer();
  }

  function closeDrawer() {
    const overlay = $("#drawerOverlay");
    const drawer = $("#calcDrawer");
    if (!drawer || !overlay) return;
    overlay.hidden = true;
    drawer.hidden = true;
    drawer.setAttribute("aria-hidden", "true");
  }

  function renderReconTable(indicatorKey, allRows) {
    const tol = Number(model.settings.reconTolerance) || 10;

    const thead = $("#reconTable thead");
    const tbody = $("#reconTable tbody");
    if (!thead || !tbody) return;

    const rows = allRows.map(r => {
      const pvB = r.m.pvBenefits;
      const pvC = r.m.pvCosts;
      const npv = r.m.npv;
      const bcr = r.m.bcr;
      const roi = r.m.roi;

      const calcNpv = pvB - pvC;
      const calcBcr = pvC > 0 ? pvB / pvC : NaN;
      const calcRoi = pvC > 0 ? (calcNpv / pvC) * 100 : NaN;

      const diffNpv = isFiniteNum(npv) && isFiniteNum(calcNpv) ? (npv - calcNpv) : NaN;
      const diffBcr = isFiniteNum(bcr) && isFiniteNum(calcBcr) ? (bcr - calcBcr) : NaN;
      const diffRoi = isFiniteNum(roi) && isFiniteNum(calcRoi) ? (roi - calcRoi) : NaN;

      const okNpv = !isFiniteNum(diffNpv) ? true : Math.abs(diffNpv) <= tol;
      const okBcr = !isFiniteNum(diffBcr) ? true : Math.abs(diffBcr) <= 1e-6;
      const okRoi = !isFiniteNum(diffRoi) ? true : Math.abs(diffRoi) <= 1e-6;

      return {
        name: r.t.name + (r.t.isControl ? " (Control)" : ""),
        pvB, pvC, npv, bcr, roi,
        calcNpv, calcBcr, calcRoi,
        okNpv, okBcr, okRoi,
        diffNpv, diffBcr, diffRoi
      };
    });

    thead.innerHTML = `
      <tr>
        <th>Treatment</th>
        <th class="num">PV(B)</th>
        <th class="num">PV(C)</th>
        <th class="num">NPV (reported)</th>
        <th class="num">NPV check (PV(B)-PV(C))</th>
        <th class="num">Flag</th>
        <th class="num">BCR check</th>
        <th class="num">ROI check</th>
      </tr>
    `;

    tbody.innerHTML = rows.map(r => {
      const flag = (!r.okNpv || !r.okBcr || !r.okRoi) ? "⚠︎" : "OK";
      return `
        <tr>
          <td>${esc(r.name)}</td>
          <td class="num">${money(r.pvB)}</td>
          <td class="num">${money(r.pvC)}</td>
          <td class="num">${money(r.npv)}</td>
          <td class="num">${money(r.calcNpv)} <span class="muted small">(${isFiniteNum(r.diffNpv) ? (r.diffNpv >= 0 ? "+" : "") + money(r.diffNpv).replace("$","$") : "n/a"} diff)</span></td>
          <td class="num">${flag}</td>
          <td class="num">${isFiniteNum(r.calcBcr) ? fmt(r.calcBcr, 4) : "n/a"} <span class="muted small">${isFiniteNum(r.diffBcr) ? ("(" + (r.diffBcr >= 0 ? "+" : "") + fmt(r.diffBcr, 6) + ")") : ""}</span></td>
          <td class="num">${isFiniteNum(r.calcRoi) ? percent(r.calcRoi) : "n/a"} <span class="muted small">${isFiniteNum(r.diffRoi) ? ("(" + (r.diffRoi >= 0 ? "+" : "") + fmt(r.diffRoi, 6) + ")") : ""}</span></td>
        </tr>
      `;
    }).join("");
  }

  // ---------- TREATMENTS UI (COST BUILD-UP) ----------
  function renderTreatments() {
    const root = $("#treatmentsList");
    if (!root) return;

    const outputs = model.outputs;

    root.innerHTML = model.treatments.map(t => {
      const totals = computeTreatmentTotals(t);
      const opPerHa = totals.opPerHa;
      const opFarm = totals.opFarm;
      const area = Number(t.areaHa) || 0;

      const costRows = (t.opCostItems || []).map(it => `
        <div class="row" style="gap:8px">
          <div class="field" style="flex:1">
            <label>Line item</label>
            <input value="${esc(it.label)}" data-cost-label="${esc(t.id)}" data-item="${esc(it.id)}" />
          </div>
          <div class="field" style="width:180px">
            <label>$/ha</label>
            <input type="number" step="0.01" value="${Number(it.valuePerHa) || 0}" data-cost-val="${esc(t.id)}" data-item="${esc(it.id)}" />
          </div>
          <div class="field" style="width:150px">
            <label>Category</label>
            <select data-cost-cat="${esc(t.id)}" data-item="${esc(it.id)}">
              <option value="Operating" ${it.category === "Operating" ? "selected" : ""}>Operating</option>
              <option value="Labour" ${it.category === "Labour" ? "selected" : ""}>Labour</option>
            </select>
          </div>
          <div class="field" style="width:110px">
            <label>&nbsp;</label>
            <button class="btn danger" type="button" data-del-cost="${esc(t.id)}" data-item="${esc(it.id)}">Remove</button>
          </div>
        </div>
      `).join("");

      const outputRows = outputs.map(o => `
        <div class="field">
          <label>${esc(o.name)} (${esc(o.unit)}) per ha</label>
          <input type="number" step="0.0001" value="${Number((t.outputsPerHa || {})[o.id]) || 0}" data-out="${esc(t.id)}" data-outid="${esc(o.id)}" />
        </div>
      `).join("");

      return `
        <div class="card" data-treat="${esc(t.id)}">
          <div class="row between">
            <h3>${esc(t.name)} ${t.isControl ? '<span class="badge">Control (baseline)</span>' : ""}</h3>
            <div class="row">
              <button class="btn" type="button" data-set-control="${esc(t.id)}">Set as control</button>
              <button class="btn danger" type="button" data-del-treatment="${esc(t.id)}">Remove</button>
            </div>
          </div>

          <div class="row-3">
            <div class="field">
              <label>Treatment name</label>
              <input value="${esc(t.name)}" data-tname="${esc(t.id)}" />
            </div>
            <div class="field">
              <label>Area (ha)</label>
              <input type="number" step="0.01" value="${area}" data-area="${esc(t.id)}" />
            </div>
            <div class="field">
              <label>Replications aggregated</label>
              <input value="${t.nReplicates || 1}" readonly />
            </div>
          </div>

          <div class="row-3">
            <div class="field">
              <label>Capital cost ($, year 0)</label>
              <input type="number" step="0.01" value="${Number(t.capitalCostY0) || 0}" data-cap="${esc(t.id)}" />
            </div>
            <div class="field">
              <label>Total operating cost ($/ha)</label>
              <input value="${fmt(opPerHa, 2)}" readonly />
            </div>
            <div class="field">
              <label>Whole-farm operating cost ($/year)</label>
              <input value="${money(opFarm)}" readonly />
            </div>
          </div>

          <div class="callout">
            <div><b>Cost build-up</b></div>
            <div class="muted small">Edit line items; totals update automatically and flow through all PV / ROI calculations.</div>
          </div>

          <div>${costRows || '<div class="muted small">No operating cost items found. Add one below.</div>'}</div>

          <div class="row">
            <button class="btn" type="button" data-add-cost="${esc(t.id)}">Add operating cost line</button>
          </div>

          <div class="callout">
            <div><b>Outputs (benefit drivers)</b></div>
            <div class="muted small">Per-hectare output levels for this treatment (used with unit values from Settings).</div>
          </div>

          <div class="row-3">${outputRows}</div>

          <div class="field">
            <label>Notes</label>
            <textarea rows="2" data-notes="${esc(t.id)}">${esc(t.notes || "")}</textarea>
          </div>
        </div>
      `;
    }).join("");

    root.onclick = e => {
      const del = e.target.closest("[data-del-treatment]");
      if (del) {
        const id = del.dataset.delTreatment;
        if (!confirm("Remove this treatment?")) return;
        model.treatments = model.treatments.filter(t => t.id !== id);
        if (!model.treatments.some(t => t.isControl) && model.treatments.length) model.treatments[0].isControl = true;
        audit("Treatment removed", id);
        renderAll();
        return;
      }
      const setC = e.target.closest("[data-set-control]");
      if (setC) {
        const id = setC.dataset.setControl;
        model.treatments.forEach(t => t.isControl = (t.id === id));
        audit("Control set", model.treatments.find(t => t.id === id)?.name || id);
        renderAll();
        return;
      }
      const addCost = e.target.closest("[data-add-cost]");
      if (addCost) {
        const id = addCost.dataset.addCost;
        const t = model.treatments.find(x => x.id === id);
        if (!t) return;
        t.opCostItems = t.opCostItems || [];
        t.opCostItems.push({ id: uid(), label: "New cost line", valuePerHa: 0, category: "Operating" });
        audit("Cost line added", t.name);
        renderAll();
        return;
      }
      const delCost = e.target.closest("[data-del-cost]");
      if (delCost) {
        const tid = delCost.dataset.delCost;
        const itemId = delCost.dataset.item;
        const t = model.treatments.find(x => x.id === tid);
        if (!t) return;
        t.opCostItems = (t.opCostItems || []).filter(it => it.id !== itemId);
        audit("Cost line removed", t.name);
        renderAll();
        return;
      }
    };

    root.oninput = e => {
      const tid =
        e.target.dataset.tname ||
        e.target.dataset.area ||
        e.target.dataset.cap ||
        e.target.dataset.notes ||
        e.target.dataset.costLabel ||
        e.target.dataset.costVal ||
        e.target.dataset.costCat ||
        e.target.dataset.out;

      if (!tid) return;
      const t = model.treatments.find(x => x.id === tid);
      if (!t) return;

      if (e.target.dataset.tname) t.name = e.target.value;
      if (e.target.dataset.area) t.areaHa = +e.target.value;
      if (e.target.dataset.cap) t.capitalCostY0 = +e.target.value;
      if (e.target.dataset.notes) t.notes = e.target.value;

      if (e.target.dataset.costLabel || e.target.dataset.costVal || e.target.dataset.costCat) {
        const itemId = e.target.dataset.item;
        const it = (t.opCostItems || []).find(x => x.id === itemId);
        if (it) {
          if (e.target.dataset.costLabel) it.label = e.target.value;
          if (e.target.dataset.costVal) it.valuePerHa = +e.target.value;
          if (e.target.dataset.costCat) it.category = e.target.value;
        }
      }

      if (e.target.dataset.out) {
        const outId = e.target.dataset.outid;
        t.outputsPerHa = t.outputsPerHa || {};
        t.outputsPerHa[outId] = +e.target.value;
      }

      renderResults();
    };
  }

  // ---------- OUTPUTS UI ----------
  function renderOutputs() {
    const tbody = $("#outputsTable tbody");
    if (!tbody) return;

    tbody.innerHTML = model.outputs.map(o => `
      <tr>
        <td><input value="${esc(o.name)}" data-out-name="${esc(o.id)}" /></td>
        <td><input value="${esc(o.unit)}" data-out-unit="${esc(o.id)}" /></td>
        <td><input type="number" step="0.01" value="${Number(o.unitValue) || 0}" data-out-val="${esc(o.id)}" /></td>
        <td><input value="${esc(o.source || "")}" data-out-src="${esc(o.id)}" /></td>
      </tr>
    `).join("");

    tbody.oninput = e => {
      const id =
        e.target.dataset.outName ||
        e.target.dataset.outUnit ||
        e.target.dataset.outVal ||
        e.target.dataset.outSrc;
      if (!id) return;

      const o = model.outputs.find(x => x.id === id);
      if (!o) return;

      if (e.target.dataset.outName) o.name = e.target.value;
      if (e.target.dataset.outUnit) o.unit = e.target.value;
      if (e.target.dataset.outVal) o.unitValue = +e.target.value;
      if (e.target.dataset.outSrc) o.source = e.target.value;

      renderResults();
    };
  }

  // ---------- SETTINGS UI ----------
  function bindSettings() {
    $("#farmArea").value = model.settings.farmAreaHa;
    $("#years").value = model.settings.years;
    $("#discRate").value = model.settings.discountRatePct;
    $("#adoption").value = model.settings.adoption;
    $("#risk").value = model.settings.risk;
    $("#tol").value = String(model.settings.reconTolerance);

    const on = () => {
      model.settings.farmAreaHa = +$("#farmArea").value;
      model.settings.years = +$("#years").value;
      model.settings.discountRatePct = +$("#discRate").value;
      model.settings.adoption = +$("#adoption").value;
      model.settings.risk = +$("#risk").value;
      model.settings.reconTolerance = +$("#tol").value;

      // Apply farm area to treatments by default (Excel-first contract, but allow per-treatment overrides)
      model.treatments.forEach(t => {
        if (!isFiniteNum(t.areaHa) || t.areaHa === 0) t.areaHa = model.settings.farmAreaHa;
      });

      renderTreatments();
      renderResults();
    };

    ["farmArea","years","discRate","adoption","risk","tol"].forEach(id => {
      $("#" + id).addEventListener("input", on);
    });

    $("#btnRecalc")?.addEventListener("click", () => {
      renderAll();
      showToast("Recalculated.");
    });
  }

  // ---------- RESULTS RENDER ----------
  function renderResults() {
    const computed = computeAllMetrics(null);
    renderLeaderboard(computed.rows);
    renderComparisonTable(computed.rows);
    renderBreakevenTable(); // based on current simulation sliders (if any)
  }

  // ---------- FILTER BUTTONS ----------
  function bindResultsFilters() {
    $("#filterAll").onclick = () => { RESULTS_FILTER.mode = "all"; renderResults(); };
    $("#filterTopNpv").onclick = () => { RESULTS_FILTER.mode = "topnpv"; renderResults(); };
    $("#filterTopBcr").onclick = () => { RESULTS_FILTER.mode = "topbcr"; renderResults(); };
    $("#filterImprove").onclick = () => { RESULTS_FILTER.mode = "improve"; renderResults(); };
  }

  // ---------- EXCEL TEMPLATE (SCHEMA-BASED) ----------
  function downloadTemplate() {
    if (typeof XLSX === "undefined") {
      alert("SheetJS (XLSX) is required.");
      return;
    }
    const wb = XLSX.utils.book_new();

    const readme = XLSX.utils.aoa_to_sheet([
      ["Farming CBA Decision Tool 2 – Excel template (schema-based)"],
      [""],
      ["Required columns (aliases accepted):"],
      ["- Treatment name", "Amendment | Treatment | Treatment name | Name"],
      ["- Yield (t/ha)", "Yield t/ha | Yield | Grain yield"],
      [""],
      ["All other numeric columns are treated as cost components by default (no omissions)."],
      ["If your sheet includes non-cost numeric outcomes (e.g., protein), reclassify them as OUTPUT in the mapping table after upload."],
      [""],
      ["Aggregation rule:"],
      ["Rows with the same treatment name are aggregated using the mean for numeric columns; replicate counts are recorded in the audit log."],
      [""],
      ["Units and formatting:"],
      ["- Costs can be entered as numbers or formatted currency (e.g., $16,850.00)."],
      ["- Blank numeric cells are treated as 0; the mapping report will flag heavy non-numeric cells."],
    ]);
    XLSX.utils.book_append_sheet(wb, readme, "Read Me");

    // Example sheet using current treatment set (if present), else default
    const baseRows = (model.importMeta.active.source === "default") ? DEFAULT_ROWS : (model.importMeta.aggregated?.treatments || []).map(t => ({
      Amendment: t.name,
      "Yield t/ha": (t.outputLevels && t.outputLevels["Yield t/ha"]) ? t.outputLevels["Yield t/ha"] : ""
    }));
    const sample = XLSX.utils.json_to_sheet(baseRows);
    XLSX.utils.book_append_sheet(wb, sample, "Data");

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile("farming_cba2_template.xlsx", wbout, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    audit("Template download", "Schema-based Excel template");
  }

  // ---------- EXCEL PARSE + MAPPING UI ----------
  let parsed = null; // { fileName, sheetName, rows, headers, roles, colStats, validation }

  function openFilePicker() {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".xlsx,.xlsm,.xlsb";
    input.style.display = "none";
    document.body.appendChild(input);
    input.addEventListener("change", async e => {
      const file = e.target.files && e.target.files[0];
      document.body.removeChild(input);
      if (!file) return;
      await parseExcelFile(file);
    });
    input.click();
  }

  async function parseExcelFile(file) {
    if (typeof XLSX === "undefined") {
      alert("SheetJS (XLSX) is required for Excel import.");
      return;
    }

    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });

      // Choose first sheet that contains required columns (if any), else first sheet.
      let chosen = wb.SheetNames[0];
      let chosenRows = null;
      let chosenHeaders = null;

      for (const sName of wb.SheetNames) {
        const sheet = wb.Sheets[sName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });
        const headers = rows.length ? Object.keys(rows[0]) : [];
        const roles = detectMappingFromHeaders(headers);
        const validation = validateImport(rows, headers, roles);
        if (validation.ok) {
          chosen = sName;
          chosenRows = rows;
          chosenHeaders = headers;
          break;
        }
      }

      if (!chosenRows) {
        const sheet = wb.Sheets[chosen];
        chosenRows = XLSX.utils.sheet_to_json(sheet, { defval: null });
        chosenHeaders = chosenRows.length ? Object.keys(chosenRows[0]) : [];
      }

      const roles = detectMappingFromHeaders(chosenHeaders);
      const validation = validateImport(chosenRows, chosenHeaders, roles);
      const colStats = validation.colStats;

      parsed = {
        fileName: file.name,
        sheetName: chosen,
        rows: chosenRows,
        headers: chosenHeaders,
        roles,
        colStats,
        validation
      };

      renderMappingUI();
      $("#btnDiscardParsed").disabled = false;
      $("#btnApplyImport").disabled = !validation.ok;

      audit("Excel parsed", `${file.name} · sheet=${chosen} · rows=${chosenRows.length}`);
      showToast(`Excel parsed. Using sheet "${chosen}".`);
    } catch (err) {
      console.error(err);
      alert("Error parsing Excel file.");
    }
  }

  function renderMappingUI() {
    const summary = $("#validationSummary");
    const tbody = $("#mappingTable tbody");
    const errorsBox = $("#validationErrors");
    const cal = $("#calibrationSummary");
    if (!tbody || !summary) return;

    if (!parsed) {
      summary.textContent = "No file parsed yet.";
      tbody.innerHTML = "";
      errorsBox.classList.remove("show");
      errorsBox.innerHTML = "";
      return;
    }

    const { rows, headers, roles, colStats, validation } = parsed;

    const errs = validation.errors || [];
    summary.innerHTML = `
      <div><b>File:</b> ${esc(parsed.fileName)} · <b>Sheet:</b> ${esc(parsed.sheetName)} · <b>Rows:</b> ${rows.length}</div>
      <div class="muted small">Import status: ${validation.ok ? "<b>Ready to apply</b>" : "<b>Blocked by errors</b>"}</div>
    `;

    // Errors
    if (errs.length) {
      errorsBox.classList.add("show");
      errorsBox.innerHTML = `<ul>${errs.map(e => `<li>${esc(e)}</li>`).join("")}</ul>`;
    } else {
      errorsBox.classList.remove("show");
      errorsBox.innerHTML = "";
    }

    // Mapping table
    tbody.innerHTML = headers.map(h => {
      const role = roles[h]?.role || ROLE.COST_OPERATING;
      const detected =
        role === ROLE.TREATMENT ? "Treatment name" :
        role === ROLE.OUTPUT ? "Output (benefit driver)" :
        role === ROLE.COST_LABOUR ? "Cost (labour)" :
        role === ROLE.COST_OPERATING ? "Cost (operating)" :
        "Ignored";

      const st = colStats[h];
      const rate = st ? Math.round(st.rate * 100) : 0;

      return `
        <tr>
          <td>${esc(h)}</td>
          <td>${esc(detected)}</td>
          <td class="num">${rate}%</td>
          <td>
            <select data-role="${esc(h)}">
              <option value="${ROLE.TREATMENT}" ${role === ROLE.TREATMENT ? "selected" : ""}>Treatment name</option>
              <option value="${ROLE.OUTPUT}" ${role === ROLE.OUTPUT ? "selected" : ""}>Output (benefit driver)</option>
              <option value="${ROLE.COST_LABOUR}" ${role === ROLE.COST_LABOUR ? "selected" : ""}>Cost (labour)</option>
              <option value="${ROLE.COST_OPERATING}" ${role === ROLE.COST_OPERATING ? "selected" : ""}>Cost (operating)</option>
              <option value="${ROLE.IGNORE}" ${role === ROLE.IGNORE ? "selected" : ""}>Ignore</option>
            </select>
          </td>
        </tr>
      `;
    }).join("");

    // Override change -> revalidate
    tbody.onchange = e => {
      const sel = e.target.closest("select[data-role]");
      if (!sel) return;
      const col = sel.dataset.role;
      const role = sel.value;
      parsed.roles[col] = parsed.roles[col] || {};
      parsed.roles[col].role = role;

      // If role becomes OUTPUT and it's yield-ish, map it to out_yield implicitly
      if (role === ROLE.OUTPUT) {
        const isYield = canonCol(col).includes("yield") || REQUIRED.yield.aliases.some(a => canonCol(col) === canonCol(a));
        parsed.roles[col].outputId = isYield ? "out_yield" : null;
      } else {
        parsed.roles[col].outputId = null;
      }

      parsed.validation = validateImport(parsed.rows, parsed.headers, parsed.roles);
      parsed.colStats = buildColumnStats(parsed.rows, parsed.headers, parsed.roles);

      // Update apply button and re-render summary/error list quickly
      $("#btnApplyImport").disabled = !parsed.validation.ok;
      renderMappingUI();
    };

    // Calibration summary (preview)
    if (cal) {
      const agg = aggregateRows(rows, headers, roles);
      const controlName = agg?.controlName || "Not detected";
      const repCount = agg ? agg.treatments.reduce((s, t) => s + (t.nReplicates || 1), 0) : 0;

      cal.innerHTML = `
        <div><span class="k">Rows read</span><span class="v">${rows.length}</span></div>
        <div><span class="k">Treatments</span><span class="v">${agg ? agg.treatments.length : "—"}</span></div>
        <div><span class="k">Replications aggregated</span><span class="v">${agg ? repCount : "—"}</span></div>
        <div><span class="k">Control detected</span><span class="v">${esc(controlName)}</span></div>
        <div><span class="k">Sheet used</span><span class="v">${esc(parsed.sheetName)}</span></div>
      `;
    }
  }

  function applyParsedImport() {
    if (!parsed) return;
    const { rows, headers, roles, validation } = parsed;
    if (!validation.ok) {
      showToast("Cannot apply import until errors are fixed.");
      return;
    }

    const agg = aggregateRows(rows, headers, roles);
    if (!agg || !agg.treatments || !agg.treatments.length) {
      showToast("Import produced no treatments.");
      return;
    }

    const meta = {
      source: "upload",
      fileName: parsed.fileName,
      sheetName: parsed.sheetName,
      rowsRead: rows.length,
      appliedAt: new Date().toISOString()
    };

    applyAggregatedToModel(agg, meta);
    parsed = null;
    renderMappingUI();
    $("#btnApplyImport").disabled = true;
    $("#btnDiscardParsed").disabled = true;
    audit("Upload set as default", meta.fileName);
  }

  function discardParsed() {
    parsed = null;
    renderMappingUI();
    $("#btnApplyImport").disabled = true;
    $("#btnDiscardParsed").disabled = true;
    showToast("Parsed file discarded.");
  }

  // ---------- SIMULATIONS (SLIDERS + BREAK-EVEN + SCENARIOS) ----------
  const simState = {
    grainPriceMult: 1,
    yieldMult: 1,
    costMult: 1,
    discountRatePct: null,
    adoption: null,
    risk: null
  };

  function bindSimControls() {
    const sPrice = $("#simGrainPrice");
    const sYield = $("#simYield");
    const sCosts = $("#simCosts");
    const sDisc = $("#simDisc");
    const sAdopt = $("#simAdopt");
    const sRisk = $("#simRisk");

    if (!sPrice) return;

    const syncLabels = () => {
      $("#simGrainPriceVal").textContent = fmt(simState.grainPriceMult, 2) + "×";
      $("#simYieldVal").textContent = fmt(simState.yieldMult, 2) + "×";
      $("#simCostsVal").textContent = fmt(simState.costMult, 2) + "×";
      $("#simDiscVal").textContent = fmt(simState.discountRatePct ?? model.settings.discountRatePct, 2) + "%";
      $("#simAdoptVal").textContent = fmt(simState.adoption ?? model.settings.adoption, 2);
      $("#simRiskVal").textContent = fmt(simState.risk ?? model.settings.risk, 2);
    };

    const reset = () => {
      simState.grainPriceMult = 1;
      simState.yieldMult = 1;
      simState.costMult = 1;
      simState.discountRatePct = model.settings.discountRatePct;
      simState.adoption = model.settings.adoption;
      simState.risk = model.settings.risk;

      sPrice.value = String(simState.grainPriceMult);
      sYield.value = String(simState.yieldMult);
      sCosts.value = String(simState.costMult);
      sDisc.value = String(simState.discountRatePct);
      sAdopt.value = String(simState.adoption);
      sRisk.value = String(simState.risk);

      syncLabels();
      renderBreakevenTable();
      showToast("Simulation sliders reset to base.");
    };

    // Init
    sPrice.value = "1.00";
    sYield.value = "1.00";
    sCosts.value = "1.00";
    sDisc.value = String(model.settings.discountRatePct);
    sAdopt.value = String(model.settings.adoption);
    sRisk.value = String(model.settings.risk);
    reset();

    const onInput = () => {
      simState.grainPriceMult = +sPrice.value;
      simState.yieldMult = +sYield.value;
      simState.costMult = +sCosts.value;
      simState.discountRatePct = +sDisc.value;
      simState.adoption = +sAdopt.value;
      simState.risk = +sRisk.value;
      syncLabels();
    };

    [sPrice, sYield, sCosts, sDisc, sAdopt, sRisk].forEach(el => el.addEventListener("input", onInput));

    $("#btnSimReset").onclick = reset;
    $("#btnSimApplyToView").onclick = () => {
      renderBreakevenTable();
      showToast("Simulation outputs updated.");
    };

    $("#btnSaveScenario").onclick = () => saveScenario();
  }

  function buildSimulationScenario() {
    // Override unit values: apply grain price multiplier to yield output only.
    const unitValuesOverride = {};
    model.outputs.forEach(o => {
      unitValuesOverride[o.id] = Number(o.unitValue) || 0;
    });
    unitValuesOverride["out_yield"] = (Number(unitValuesOverride["out_yield"]) || 0) * simState.grainPriceMult;

    // Override outputs: apply yield multiplier to all outputsPerHa (including yield)
    // (This is intentionally broad and transparent; user can interpret as “good/bad year” scaling.)
    const outputsOverrideByTid = new Map();
    model.treatments.forEach(t => {
      const o = {};
      model.outputs.forEach(out => {
        o[out.id] = (Number((t.outputsPerHa || {})[out.id]) || 0) * simState.yieldMult;
      });
      outputsOverrideByTid.set(t.id, o);
    });

    // Override costs: multiply operating costs by costMult
    const costMult = simState.costMult;

    return {
      years: model.settings.years,
      discountRatePct: simState.discountRatePct,
      adoption: simState.adoption,
      risk: simState.risk,
      unitValuesOverride,
      outputsOverrideByTid,
      costMult
    };
  }

  function computeMetricsForTreatmentUnderScenario(t, scenarioObj) {
    // Apply scenario outputs override
    const outputsOverride = scenarioObj.outputsOverrideByTid.get(t.id);

    // Apply scenario unit values override
    const unitValuesOverride = scenarioObj.unitValuesOverride;

    // Apply scenario cost multiplier (operating costs only)
    const tAdj = { ...t, opCostItems: (t.opCostItems || []).map(it => ({ ...it, valuePerHa: (Number(it.valuePerHa) || 0) * scenarioObj.costMult })) };

    return computeMetricsForTreatment(tAdj, {
      years: scenarioObj.years,
      discountRatePct: scenarioObj.discountRatePct,
      adoption: scenarioObj.adoption,
      risk: scenarioObj.risk,
      outputsOverride,
      unitValuesOverride
    });
  }

  function renderBreakevenTable() {
    const tbody = $("#breakevenTable tbody");
    if (!tbody) return;

    const scenarioObj = buildSimulationScenario();
    const control = model.treatments.find(t => t.isControl) || model.treatments[0];
    if (!control) {
      tbody.innerHTML = `<tr><td colspan="4" class="muted">No treatments available.</td></tr>`;
      return;
    }

    // Control metrics under scenario
    const cM = computeMetricsForTreatmentUnderScenario(control, scenarioObj);

    const rows = model.treatments
      .filter(t => !t.isControl)
      .map(t => {
        const baseM = computeMetricsForTreatmentUnderScenario(t, scenarioObj);

        // Break-even 1: grain price multiplier for ΔNPV=0 vs control (bisection on yield unit value multiplier)
        const priceBE = findBreakEvenPriceMultiplier(t, control, scenarioObj);

        // Break-even 2: yield multiplier for BCR=1 (bisection on outputs multiplier)
        const yieldBE = findBreakEvenYieldMultiplierForBcr1(t, scenarioObj);

        // Break-even 3: cost reduction to reach ROI=0 (solve for cost multiplier where NPV=0)
        const costRed = findCostReductionForRoi0(t, scenarioObj);

        return {
          name: t.name,
          priceBE,
          yieldBE,
          costRed,
          baseDeltaNpv: baseM.npv - cM.npv
        };
      })
      .sort((a, b) => b.baseDeltaNpv - a.baseDeltaNpv);

    tbody.innerHTML = rows.map(r => `
      <tr>
        <td><b>${esc(r.name)}</b></td>
        <td class="num">${r.priceBE != null ? money(r.priceBE) : "n/a"}</td>
        <td class="num">${r.yieldBE != null ? fmt(r.yieldBE, 2) + "×" : "n/a"}</td>
        <td class="num">${r.costRed != null ? fmt(r.costRed, 1) + "%" : "n/a"}</td>
      </tr>
    `).join("");
  }

  function yieldUnitValueBase() {
    const y = model.outputs.find(o => o.id === "out_yield");
    return y ? Number(y.unitValue) || 0 : 0;
  }

  function findBreakEvenPriceMultiplier(t, control, scenarioObj) {
    // Solve for grain price ($/t) that makes ΔNPV = 0 vs control, holding everything else in scenario fixed.
    const basePrice = yieldUnitValueBase();
    if (basePrice <= 0) return null;

    const lo = 0; // $/t
    let hi = basePrice * 10; // expand if needed

    const f = (price) => {
      const unitValuesOverride = { ...scenarioObj.unitValuesOverride, out_yield: price };
      const sc = { ...scenarioObj, unitValuesOverride };

      const cM = computeMetricsForTreatmentUnderScenario(control, sc);
      const tM = computeMetricsForTreatmentUnderScenario(t, sc);
      return (tM.npv - cM.npv);
    };

    let fLo = f(lo);
    let fHi = f(hi);

    // If already positive at lo (rare), return lo
    if (isFiniteNum(fLo) && fLo >= 0) return lo;

    // Expand hi until sign change or limit
    let tries = 0;
    while (tries < 12 && (!isFiniteNum(fHi) || fLo * fHi > 0)) {
      hi *= 1.6;
      fHi = f(hi);
      tries++;
    }
    if (!isFiniteNum(fHi) || fLo * fHi > 0) return null;

    // Bisection
    let a = lo, b = hi;
    for (let i = 0; i < 60; i++) {
      const mid = (a + b) / 2;
      const fm = f(mid);
      if (!isFiniteNum(fm)) return null;
      if (Math.abs(fm) < 1e-6) return mid;
      if (fLo * fm <= 0) { b = mid; fHi = fm; }
      else { a = mid; fLo = fm; }
    }
    return (a + b) / 2;
  }

  function findBreakEvenYieldMultiplierForBcr1(t, scenarioObj) {
    // Solve for yield multiplier that makes BCR = 1 (holding costs).
    const f = (mult) => {
      const outputsOverrideByTid = new Map();
      model.treatments.forEach(tt => {
        const o = {};
        model.outputs.forEach(out => {
          o[out.id] = (Number((tt.outputsPerHa || {})[out.id]) || 0) * mult;
        });
        outputsOverrideByTid.set(tt.id, o);
      });
      const sc = { ...scenarioObj, outputsOverrideByTid };
      const m = computeMetricsForTreatmentUnderScenario(t, sc);
      return (m.bcr - 1);
    };

    let lo = 0.1, hi = 5.0;
    let fLo = f(lo), fHi = f(hi);
    if (!isFiniteNum(fLo) || !isFiniteNum(fHi)) return null;

    if (fLo >= 0) return lo;
    let tries = 0;
    while (tries < 10 && fLo * fHi > 0) {
      hi *= 1.6;
      fHi = f(hi);
      tries++;
      if (!isFiniteNum(fHi)) return null;
    }
    if (fLo * fHi > 0) return null;

    for (let i = 0; i < 60; i++) {
      const mid = (lo + hi) / 2;
      const fm = f(mid);
      if (!isFiniteNum(fm)) return null;
      if (Math.abs(fm) < 1e-6) return mid;
      if (fLo * fm <= 0) { hi = mid; fHi = fm; }
      else { lo = mid; fLo = fm; }
    }
    return (lo + hi) / 2;
  }

  function findCostReductionForRoi0(t, scenarioObj) {
    // ROI=0 implies NPV=0 (when PV(C)>0). Solve for operating cost multiplier m such that NPV=0.
    const f = (costMult) => {
      const sc = { ...scenarioObj, costMult };
      const m = computeMetricsForTreatmentUnderScenario(t, sc);
      return m.npv; // want 0
    };

    let lo = 0.0, hi = 2.0;
    let fLo = f(lo), fHi = f(hi);
    if (!isFiniteNum(fLo) || !isFiniteNum(fHi)) return null;

    // If NPV already >=0 at hi, then no reduction needed (could even increase)
    if (fHi >= 0) return 0;

    // If NPV <0 even at costMult=0, not solvable by cost reduction
    if (fLo < 0) return null;

    // Bisection to find m where NPV=0 between lo and hi
    for (let i = 0; i < 60; i++) {
      const mid = (lo + hi) / 2;
      const fm = f(mid);
      if (!isFiniteNum(fm)) return null;
      if (Math.abs(fm) < 1e-6) {
        const reductionPct = (1 - mid) * 100;
        return clamp(reductionPct, 0, 100);
      }
      if (fLo * fm <= 0) { hi = mid; fHi = fm; }
      else { lo = mid; fLo = fm; }
    }
    const m = (lo + hi) / 2;
    const reductionPct = (1 - m) * 100;
    return clamp(reductionPct, 0, 100);
  }

  // Saved scenarios
  function loadScenarios() {
    try {
      const raw = localStorage.getItem(STORAGE_KEYS.scenarios);
      return raw ? (JSON.parse(raw) || []) : [];
    } catch { return []; }
  }
  function saveScenarios(list) {
    try { localStorage.setItem(STORAGE_KEYS.scenarios, JSON.stringify(list)); } catch {}
  }

  function saveScenario() {
    const name = prompt("Scenario name (e.g., Dry year, High fuel cost, Lower market price):");
    if (!name) return;
    const s = {
      id: uid(),
      name: name.trim().slice(0, 60),
      createdAt: new Date().toISOString(),
      sliders: { ...simState },
      settingsSnapshot: { ...model.settings }
    };
    const list = loadScenarios();
    list.unshift(s);
    saveScenarios(list.slice(0, 30));
    renderScenarioList();
    audit("Scenario saved", s.name);
    showToast("Scenario saved.");
  }

  function renderScenarioList() {
    const root = $("#scenarioList");
    if (!root) return;
    const list = loadScenarios();
    if (!list.length) {
      root.innerHTML = `<div class="muted small">No saved scenarios yet.</div>`;
      return;
    }
    root.innerHTML = list.map(s => `
      <div class="scenario">
        <div class="name">${esc(s.name)}</div>
        <div class="meta">${esc(new Date(s.createdAt).toLocaleString())}</div>
        <div class="meta">Price× ${fmt(s.sliders.grainPriceMult,2)} · Yield× ${fmt(s.sliders.yieldMult,2)} · Cost× ${fmt(s.sliders.costMult,2)} · r ${fmt(s.sliders.discountRatePct,2)}%</div>
        <div class="actions">
          <button class="btn" type="button" data-load-scn="${esc(s.id)}">Load</button>
          <button class="btn danger" type="button" data-del-scn="${esc(s.id)}">Delete</button>
        </div>
      </div>
    `).join("");

    root.onclick = e => {
      const loadBtn = e.target.closest("[data-load-scn]");
      const delBtn = e.target.closest("[data-del-scn]");
      if (loadBtn) {
        const id = loadBtn.dataset.loadScn;
        const scn = list.find(x => x.id === id);
        if (!scn) return;
        // Apply sliders
        Object.assign(simState, scn.sliders);

        // Push to UI controls
        $("#simGrainPrice").value = String(simState.grainPriceMult);
        $("#simYield").value = String(simState.yieldMult);
        $("#simCosts").value = String(simState.costMult);
        $("#simDisc").value = String(simState.discountRatePct);
        $("#simAdopt").value = String(simState.adoption);
        $("#simRisk").value = String(simState.risk);
        // Trigger label refresh
        $("#simGrainPrice").dispatchEvent(new Event("input"));
        renderBreakevenTable();
        audit("Scenario loaded", scn.name);
        showToast("Scenario loaded.");
        return;
      }
      if (delBtn) {
        const id = delBtn.dataset.delScn;
        const kept = list.filter(x => x.id !== id);
        saveScenarios(kept);
        renderScenarioList();
        audit("Scenario deleted", id);
        showToast("Scenario deleted.");
      }
    };
  }

  // ---------- AI PROMPT ----------
  function topDriversForTreatment(t) {
    // Costs: top 3 line items by $/ha
    const costItems = (t.opCostItems || []).slice().sort((a, b) => (Number(b.valuePerHa)||0) - (Number(a.valuePerHa)||0)).slice(0, 3);
    // Benefits: top outputs by contribution per ha
    const out = model.outputs.map(o => {
      const level = Number((t.outputsPerHa || {})[o.id]) || 0;
      const v = Number(o.unitValue) || 0;
      return { name: o.name, unit: o.unit, level, unitValue: v, perHaValue: level * v };
    }).sort((a, b) => b.perHaValue - a.perHaValue).slice(0, 3);

    return {
      topCostsPerHa: costItems.map(x => ({ label: x.label, category: x.category, valuePerHa: Number(x.valuePerHa) || 0 })),
      topBenefitsPerHa: out.map(x => ({ output: x.name, unit: x.unit, levelPerHa: x.level, unitValue: x.unitValue, valuePerHa: x.perHaValue }))
    };
  }

  function buildComparisonCompactJSON() {
    const computed = computeAllMetrics(null);
    const control = computed.control;
    const rows = computed.rows;

    const out = rows.map(r => ({
      treatment: r.t.name,
      isControl: r.t.isControl,
      rank: r.rank,
      pvBenefits: r.m.pvBenefits,
      pvCosts: r.m.pvCosts,
      npv: r.m.npv,
      bcr: r.m.bcr,
      roiPct: r.m.roi,
      deltaVsControl: control ? {
        pvBenefits: r.m.pvBenefits - computed.controlMetrics.pvBenefits,
        pvCosts: r.m.pvCosts - computed.controlMetrics.pvCosts,
        npv: r.m.npv - computed.controlMetrics.npv
      } : null
    }));

    return { control: control ? control.name : null, rows: out };
  }

  function buildAIPromptText() {
    const computed = computeAllMetrics(null);
    const compact = buildComparisonCompactJSON();

    const drivers = model.treatments.map(t => ({
      treatment: t.name,
      isControl: t.isControl,
      drivers: topDriversForTreatment(t)
    }));

    const promptObj = {
      tool: model.toolName,
      version: TOOL_VERSION,
      scenarioSettings: { ...model.settings },
      importMeta: { ...model.importMeta.active },
      comparisonToControl: compact,
      treatmentDrivers: drivers,
      instructions: [
        "Write in plain language for a farmer or on-farm manager. Avoid jargon.",
        "Explain PV Benefits, PV Costs, NPV, BCR, ROI in practical terms.",
        "Compare each treatment to the control. Highlight whether a treatment wins by lifting benefits, cutting costs, or both.",
        "Do not recommend a choice and do not impose rules or thresholds.",
        "For low BCR or negative ΔNPV vs control, suggest improvement options framed as possibilities (reduce costs, improve yield, improve price outcomes, improve implementation efficiency, agronomy options).",
        "Note uncertainty and what assumptions drive results."
      ]
    };

    return JSON.stringify(promptObj, null, 2);
  }

  function bindAI() {
    $("#btnBuildPrompt").onclick = () => {
      const text = buildAIPromptText();
      $("#aiPrompt").value = text;
      audit("AI prompt built", "Prompt generated");
      showToast("AI prompt generated.");
    };

    $("#btnCopyPrompt").onclick = async () => {
      const text = $("#aiPrompt").value || buildAIPromptText();
      $("#aiPrompt").value = text;
      try {
        await navigator.clipboard.writeText(text);
        audit("AI prompt copied", "Clipboard");
        showToast("Copied to clipboard.");
      } catch {
        showToast("Copy failed. You can manually copy from the box.");
      }
    };

    $("#btnExportAIPack").onclick = () => {
      const text = $("#aiPrompt").value || buildAIPromptText();
      exportExcelWorkbook({ includeAIPromptOnly: true, aiPromptText: text });
      audit("AI brief pack exported", "Excel");
    };
  }

  // ---------- EXPORTS ----------
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

  function exportExcelWorkbook(opts = {}) {
    if (typeof XLSX === "undefined") {
      alert("SheetJS (XLSX) is required for Excel export.");
      return;
    }
    const wb = XLSX.utils.book_new();

    const computed = computeAllMetrics(null);
    const control = computed.control;
    const rows = computed.rows;

    // Results sheet: tidy comparison table
    const indicators = [
      { k: "pvBenefits", label: "PV Benefits", fmt: money },
      { k: "pvCosts", label: "PV Costs", fmt: money },
      { k: "npv", label: "NPV", fmt: money },
      { k: "bcr", label: "BCR", fmt: v => (isFiniteNum(v) ? fmt(v, 4) : "") },
      { k: "roi", label: "ROI (%)", fmt: v => (isFiniteNum(v) ? fmt(v, 4) : "") },
      { k: "rank", label: "Rank", fmt: v => (v == null ? "" : String(v)) },
      { k: "deltaNpv", label: "ΔNPV vs Control", fmt: money },
      { k: "deltaPvCost", label: "ΔPV Cost vs Control", fmt: money }
    ];

    const header = ["Indicator", "Control (baseline)"];
    const nonControl = rows.filter(r => !r.t.isControl);
    nonControl.forEach(r => {
      header.push(r.t.name);
      header.push(r.t.name + " Δ vs control");
      header.push(r.t.name + " Δ%");
    });

    const aoa = [header];
    indicators.forEach(ind => {
      const row = [ind.label];
      const c = rows.find(r => r.t.isControl) || null;
      row.push(c ? (ind.k === "deltaNpv" || ind.k === "deltaPvCost" ? ind.fmt(0) : ind.k === "rank" ? "Baseline" : ind.fmt(c.m[ind.k] ?? c[ind.k])) : "");

      nonControl.forEach(r => {
        // value
        let v;
        if (ind.k === "deltaNpv") v = control ? (r.m.npv - computed.controlMetrics.npv) : NaN;
        else if (ind.k === "deltaPvCost") v = control ? (r.m.pvCosts - computed.controlMetrics.pvCosts) : NaN;
        else if (ind.k === "rank") v = r.rank;
        else v = r.m[ind.k];

        row.push(ind.fmt(v));

        // delta
        let d = "";
        if (control && ind.k !== "rank" && ind.k !== "deltaNpv" && ind.k !== "deltaPvCost") {
          const base = computed.controlMetrics[ind.k];
          if (isFiniteNum(v) && isFiniteNum(base)) {
            const dv = v - base;
            d = (ind.k === "bcr" || ind.k === "roi") ? fmt(dv, 4) : money(dv);
          }
        }
        row.push(d);

        // delta %
        let p = "";
        if (control && ind.k !== "rank" && ind.k !== "deltaNpv" && ind.k !== "deltaPvCost") {
          const base = computed.controlMetrics[ind.k];
          if (isFiniteNum(v) && isFiniteNum(base) && Math.abs(base) > 1e-12) {
            p = fmt(((v - base) / base) * 100, 2) + "%";
          }
        }
        row.push(p);
      });

      aoa.push(row);
    });

    const wsResults = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, wsResults, "Results");

    if (!opts.includeAIPromptOnly) {
      // Inputs sheet
      const inputs = [];
      inputs.push(["Treatment", "IsControl", "Area(ha)", "CapitalCostY0"]);
      // add outputs columns
      model.outputs.forEach(o => inputs[0].push(`${o.name} (${o.unit}) per ha`));
      // add cost items collapsed (total op cost/ha + list)
      inputs[0].push("Total operating cost ($/ha)");
      inputs[0].push("Operating cost line items (label: $/ha)");

      model.treatments.forEach(t => {
        const totals = computeTreatmentTotals(t);
        const row = [t.name, t.isControl ? 1 : 0, t.areaHa, t.capitalCostY0];
        model.outputs.forEach(o => row.push(Number((t.outputsPerHa || {})[o.id]) || 0));
        row.push(totals.opPerHa);
        row.push((t.opCostItems || []).map(it => `${it.label}: ${Number(it.valuePerHa)||0}`).join(" | "));
        inputs.push(row);
      });

      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(inputs), "Inputs");

      // Assumptions
      const ass = [
        ["Tool", model.toolName],
        ["Version", TOOL_VERSION],
        ["Import source", model.importMeta.active.source],
        ["Import file", model.importMeta.active.fileName],
        ["Sheet", model.importMeta.active.sheetName],
        ["Rows read", model.importMeta.active.rowsRead],
        ["Applied at", model.importMeta.active.appliedAt],
        [""],
        ["Farm area (ha)", model.settings.farmAreaHa],
        ["Time horizon (years)", model.settings.years],
        ["Discount rate (% p.a.)", model.settings.discountRatePct],
        ["Adoption", model.settings.adoption],
        ["Risk", model.settings.risk]
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ass), "Assumptions");

      // Simulations (saved scenarios)
      const scn = loadScenarios();
      const scnAoa = [["Scenario", "CreatedAt", "PriceMult", "YieldMult", "CostMult", "DiscountRate", "Adoption", "Risk"]];
      scn.forEach(s => {
        scnAoa.push([s.name, s.createdAt, s.sliders.grainPriceMult, s.sliders.yieldMult, s.sliders.costMult, s.sliders.discountRatePct, s.sliders.adoption, s.sliders.risk]);
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(scnAoa), "Simulations");

      // Audit log
      const aud = [["Time", "Action", "Details"]];
      auditLog.slice().reverse().forEach(a => aud.push([a.ts, a.action, a.details]));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aud), "AuditLog");
    }

    // AI Prompt
    const promptText = opts.aiPromptText || ($("#aiPrompt")?.value || buildAIPromptText());
    const wsAI = XLSX.utils.aoa_to_sheet([["AI_Prompt_JSON"], [promptText]]);
    XLSX.utils.book_append_sheet(wb, wsAI, "AI_Prompt");

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile("farming_cba2_export.xlsx", wbout, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Excel exported.");
  }

  function bindExports() {
    $("#btnExportExcel").onclick = () => {
      exportExcelWorkbook();
      audit("Excel export", "Workbook exported");
    };
    $("#btnPrintFull").onclick = () => {
      try { localStorage.setItem(STORAGE_KEYS.printMode, "full"); } catch {}
      audit("Print", "Results (full)");
      window.print();
    };
    $("#btnPrintCondensed").onclick = () => {
      try { localStorage.setItem(STORAGE_KEYS.printMode, "condensed"); } catch {}
      audit("Print", "Results (condensed)");
      // In this simple build, condensed uses same print but user can hide by CSS if extended; we keep audit.
      window.print();
    };
  }

  // ---------- HOME/HEADER BUTTONS ----------
  function resetToDefaultDataset() {
    const headers = Object.keys(DEFAULT_ROWS[0] || {});
    const roles = detectMappingFromHeaders(headers);
    const validation = validateImport(DEFAULT_ROWS, headers, roles);
    if (!validation.ok) {
      alert("Default dataset failed validation, which should not happen.");
      return;
    }
    const agg = aggregateRows(DEFAULT_ROWS, headers, roles);
    const meta = { source: "default", fileName: "Default dataset", sheetName: "Default", rowsRead: DEFAULT_ROWS.length, appliedAt: new Date().toISOString() };
    applyAggregatedToModel(agg, meta);
    audit("Reset to default dataset", "Default dataset applied");
  }

  // ---------- DATA TAB BINDINGS ----------
  function bindDataTab() {
    $("#btnChooseExcel").onclick = openFilePicker;
    $("#btnApplyImport").onclick = applyParsedImport;
    $("#btnDiscardParsed").onclick = discardParsed;
    $("#btnDownloadTemplate").onclick = downloadTemplate;
  }

  // ---------- RESULTS TAB STYLES (linklike buttons) ----------
  function injectLinklikeStyle() {
    const style = document.createElement("style");
    style.textContent = `
      .linklike{
        border:none;
        background:transparent;
        color:#fff;
        font-weight:900;
        padding:0;
        cursor:pointer;
        text-align:left;
      }
      .linklike:focus{outline:none; box-shadow: var(--focus); border-radius:8px}
    `;
    document.head.appendChild(style);
  }

  // ---------- RENDER ALL ----------
  function renderAll() {
    $("#toolVersionBadge").textContent = "v" + TOOL_VERSION;
    renderOutputs();
    renderTreatments();
    renderResults();
    renderScenarioList();
    renderAudit();
  }

  // ---------- INIT ----------
  function init() {
    loadAudit();
    initTabs();
    injectLinklikeStyle();

    // Header actions
    $("#resetToDefault").onclick = () => resetToDefaultDataset();
    $("#restoreLastUpload").onclick = () => restoreLastSuccessfulUpload();
    $("#btnClearAudit").onclick = () => {
      if (!confirm("Clear audit log in this browser?")) return;
      auditLog = [];
      try { localStorage.removeItem(STORAGE_KEYS.audit); } catch {}
      renderAudit();
      showToast("Audit log cleared.");
    };

    // Buttons
    bindDataTab();
    bindSettings();
    bindResultsFilters();
    bindSimControls();
    bindAI();
    bindExports();

    // Add output / treatment
    $("#btnAddOutput").onclick = () => {
      const id = "out_" + uid();
      model.outputs.push({ id, name: "New output", unit: "unit/ha", unitValue: 0, source: "Input Directly" });
      // Add to treatments
      model.treatments.forEach(t => {
        t.outputsPerHa = t.outputsPerHa || {};
        t.outputsPerHa[id] = 0;
      });
      audit("Output added", id);
      renderAll();
    };

    $("#btnAddTreatment").onclick = () => {
      const farmArea = Number(model.settings.farmAreaHa) || 0;
      const t = {
        id: uid(),
        name: "New treatment",
        isControl: false,
        areaHa: farmArea,
        adoption: 1,
        capitalCostY0: 0,
        opCostItems: [{ id: uid(), label: "New cost line", valuePerHa: 0, category: "Operating" }],
        outputsPerHa: Object.fromEntries(model.outputs.map(o => [o.id, 0])),
        nReplicates: 1,
        notes: ""
      };
      model.treatments.push(t);
      audit("Treatment added", t.name);
      renderAll();
    };

    // Apply default dataset at startup (and treat it as active default)
    resetToDefaultDataset();

    // Map UI state
    renderMappingUI();

    // Print mode hint (optional, for future extension)
    try {
      const pm = localStorage.getItem(STORAGE_KEYS.printMode);
      if (pm) audit("Print mode set (remembered)", pm);
    } catch {}
  }

  document.addEventListener("DOMContentLoaded", init);
})();
