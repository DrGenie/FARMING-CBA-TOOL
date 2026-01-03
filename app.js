// app.js
/* Fully functional Farming CBA Decision Aid (vanilla JS + SheetJS):
   - Active tabs (Import, Mapping, Assumptions, Results, Cashflows, Sensitivity, Export, Help)
   - Robust Excel read (auto header detection + dedupe headers)
   - Flexible mapping for treatment/yield/cost
   - Per-treatment overrides: adoption, yield multiplier, cost multiplier, cost timing
   - Full CBA: PV benefits, PV costs, NPV, BCR, ROI, ranking; control included in vertical table
   - Cashflow tables + simple canvas charts
   - Sensitivity (one-way) + XLSX export
   - Clean Excel export (multiple sheets)
   ASCII-only.
*/

/* =========================
   Embedded Lockhart sample
   ========================= */
const EMBEDDED_LOCKHART_ROWS = [
  {"Plot":1,"Rep":1,"Trt":12,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)","Practice Change":"Crop 1","Yield t/ha":7.029229293617021,"Protein":23.2,"|":17945.488764568763},
  {"Plot":2,"Rep":1,"Trt":6,"Amendment":"Deep OM (CP1)","Practice Change":"Crop 1","Yield t/ha":6.539273035489362,"Protein":23.6,"|":24884.884058984914},
  {"Plot":3,"Rep":1,"Trt":13,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)+PAM","Practice Change":"Crop 1","Yield t/ha":6.54757540287234,"Protein":23.7,"|":18463.88888888889},
  {"Plot":4,"Rep":1,"Trt":9,"Amendment":"Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.37207183687234,"Protein":24.7,"|":1012.5633802816902},
  {"Plot":5,"Rep":1,"Trt":2,"Amendment":"Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.667165176319149,"Protein":23.9,"|":912.1951219512195},
  {"Plot":6,"Rep":1,"Trt":11,"Amendment":"Lime+Deep OM+Gypsum CHT+PAM","Practice Change":"Crop 1","Yield t/ha":7.199593337872341,"Protein":23.3,"|":976.8292682926829},
  {"Plot":7,"Rep":1,"Trt":5,"Amendment":"Lime +Deep OM (CP1)","Practice Change":"Crop 1","Yield t/ha":6.614808249489362,"Protein":23.4,"|":1033.6585365853657},
  {"Plot":8,"Rep":1,"Trt":15,"Amendment":"Liquid Gypsum (CHT)+PAM","Practice Change":"Crop 1","Yield t/ha":6.958392845957447,"Protein":23.1,"|":782.9268292682927},
  {"Plot":9,"Rep":1,"Trt":10,"Amendment":"Lime+Deep OM+Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.280082541361702,"Protein":23.7,"|":1059.349593495935},
  {"Plot":10,"Rep":1,"Trt":4,"Amendment":"Lime + Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.839204480808511,"Protein":23.8,"|":908.9430894308944},
  {"Plot":11,"Rep":1,"Trt":1,"Amendment":"Lime only","Practice Change":"Crop 1","Yield t/ha":7.142857142857143,"Protein":23.3,"|":790.650406504065},
  {"Plot":12,"Rep":1,"Trt":14,"Amendment":"Liquid Gypsum (CHT)","Practice Change":"Crop 1","Yield t/ha":7.421177186276596,"Protein":23.4,"|":748.780487804878},
  {"Plot":13,"Rep":1,"Trt":7,"Amendment":"Lime+Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.35011252212766,"Protein":23.7,"|":1024.5934959349595},
  {"Plot":14,"Rep":1,"Trt":3,"Amendment":"Lime + Gypsum CHT+PAM","Practice Change":"Crop 1","Yield t/ha":7.35667234112766,"Protein":23.2,"|":923.1707317073171},
  {"Plot":15,"Rep":1,"Trt":8,"Amendment":"Lime+Gypsum CHT+Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.269241385957447,"Protein":23.9,"|":1062.439024390244},
  {"Plot":16,"Rep":1,"Trt":0,"Amendment":"Control","Practice Change":"Crop 1","Yield t/ha":7.626622926382979,"Protein":23.5,"|":694.6341463414634},

  {"Plot":17,"Rep":2,"Trt":10,"Amendment":"Lime+Deep OM+Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":8.16927218306383,"Protein":23.5,"|":1059.349593495935},
  {"Plot":18,"Rep":2,"Trt":14,"Amendment":"Liquid Gypsum (CHT)","Practice Change":"Crop 1","Yield t/ha":7.467401828978723,"Protein":23.3,"|":748.780487804878},
  {"Plot":19,"Rep":2,"Trt":3,"Amendment":"Lime + Gypsum CHT+PAM","Practice Change":"Crop 1","Yield t/ha":7.751529506382978,"Protein":23.1,"|":923.1707317073171},
  {"Plot":20,"Rep":2,"Trt":7,"Amendment":"Lime+Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.771685019574468,"Protein":23.7,"|":1024.5934959349595},
  {"Plot":21,"Rep":2,"Trt":11,"Amendment":"Lime+Deep OM+Gypsum CHT+PAM","Practice Change":"Crop 1","Yield t/ha":7.373296617021276,"Protein":23.5,"|":976.8292682926829},
  {"Plot":22,"Rep":2,"Trt":4,"Amendment":"Lime + Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.794740369574468,"Protein":23.4,"|":908.9430894308944},
  {"Plot":23,"Rep":2,"Trt":1,"Amendment":"Lime only","Practice Change":"Crop 1","Yield t/ha":7.830657711489362,"Protein":23.2,"|":790.650406504065},
  {"Plot":24,"Rep":2,"Trt":5,"Amendment":"Lime +Deep OM (CP1)","Practice Change":"Crop 1","Yield t/ha":7.082754630638297,"Protein":23.3,"|":1033.6585365853657},
  {"Plot":25,"Rep":2,"Trt":8,"Amendment":"Lime+Gypsum CHT+Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.722016223404255,"Protein":23.6,"|":1062.439024390244},
  {"Plot":26,"Rep":2,"Trt":2,"Amendment":"Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.417663186382979,"Protein":23.2,"|":912.1951219512195},
  {"Plot":27,"Rep":2,"Trt":15,"Amendment":"Liquid Gypsum (CHT)+PAM","Practice Change":"Crop 1","Yield t/ha":7.268999694468085,"Protein":23.1,"|":782.9268292682927},
  {"Plot":28,"Rep":2,"Trt":9,"Amendment":"Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.640707778085106,"Protein":23.5,"|":1012.5633802816902},
  {"Plot":29,"Rep":2,"Trt":13,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)+PAM","Practice Change":"Crop 1","Yield t/ha":6.601502030425532,"Protein":23.1,"|":18463.88888888889},
  {"Plot":30,"Rep":2,"Trt":6,"Amendment":"Deep OM (CP1)","Practice Change":"Crop 1","Yield t/ha":6.842683569361702,"Protein":23.5,"|":17787.777777777777},
  {"Plot":31,"Rep":2,"Trt":12,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)","Practice Change":"Crop 1","Yield t/ha":7.006188607446808,"Protein":23.2,"|":17848.538011695906},
  {"Plot":32,"Rep":2,"Trt":0,"Amendment":"Control","Practice Change":"Crop 1","Yield t/ha":7.482897900425532,"Protein":23.2,"|":694.6341463414634},

  {"Plot":33,"Rep":3,"Trt":7,"Amendment":"Lime+Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.322008111702128,"Protein":23.4,"|":1024.5934959349595},
  {"Plot":34,"Rep":3,"Trt":3,"Amendment":"Lime + Gypsum CHT+PAM","Practice Change":"Crop 1","Yield t/ha":7.048594062978723,"Protein":23.5,"|":923.1707317073171},
  {"Plot":35,"Rep":3,"Trt":1,"Amendment":"Lime only","Practice Change":"Crop 1","Yield t/ha":7.468294498723404,"Protein":23.5,"|":790.650406504065},
  {"Plot":36,"Rep":3,"Trt":14,"Amendment":"Liquid Gypsum (CHT)","Practice Change":"Crop 1","Yield t/ha":7.208078594468085,"Protein":23.5,"|":748.780487804878},
  {"Plot":37,"Rep":3,"Trt":10,"Amendment":"Lime+Deep OM+Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.82302918893617,"Protein":23.5,"|":1059.349593495935},
  {"Plot":38,"Rep":3,"Trt":5,"Amendment":"Lime +Deep OM (CP1)","Practice Change":"Crop 1","Yield t/ha":6.866323363404255,"Protein":23.7,"|":1033.6585365853657},
  {"Plot":39,"Rep":3,"Trt":11,"Amendment":"Lime+Deep OM+Gypsum CHT+PAM","Practice Change":"Crop 1","Yield t/ha":7.176539666595744,"Protein":23.4,"|":976.8292682926829},
  {"Plot":40,"Rep":3,"Trt":15,"Amendment":"Liquid Gypsum (CHT)+PAM","Practice Change":"Crop 1","Yield t/ha":7.064219107659574,"Protein":23.4,"|":782.9268292682927},
  {"Plot":41,"Rep":3,"Trt":12,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)","Practice Change":"Crop 1","Yield t/ha":6.519162115744681,"Protein":23.4,"|":17945.488764568763},
  {"Plot":42,"Rep":3,"Trt":8,"Amendment":"Lime+Gypsum CHT+Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.05956678787234,"Protein":23.7,"|":1062.439024390244},
  {"Plot":43,"Rep":3,"Trt":6,"Amendment":"Deep OM (CP1)","Practice Change":"Crop 1","Yield t/ha":6.396029749361702,"Protein":23.7,"|":17787.777777777777},
  {"Plot":44,"Rep":3,"Trt":9,"Amendment":"Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.268999694468085,"Protein":23.6,"|":1012.5633802816902},
  {"Plot":45,"Rep":3,"Trt":2,"Amendment":"Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.120229763404255,"Protein":23.7,"|":912.1951219512195},
  {"Plot":46,"Rep":3,"Trt":4,"Amendment":"Lime + Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.462494905957446,"Protein":23.4,"|":908.9430894308944},
  {"Plot":47,"Rep":3,"Trt":13,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)+PAM","Practice Change":"Crop 1","Yield t/ha":6.435858070212766,"Protein":23.6,"|":18463.88888888889},
  {"Plot":48,"Rep":3,"Trt":0,"Amendment":"Control","Practice Change":"Crop 1","Yield t/ha":7.269686206808511,"Protein":23.5,"|":694.6341463414634}
];

/* =========================
   State
   ========================= */
const state = {
  source: null, // {kind:'excel'|'embedded', name:string}
  raw: [],
  columns: [],
  map: {
    treatment: null,
    baseline: null,
    yield: null,
    cost: null,
    optional1: null
  },
  assumptions: {
    areaHa: 100,
    horizonYears: 10,
    discountRatePct: 7,
    priceYear1: 450,
    priceGrowthPct: 0,
    yieldScale: 1,
    costScale: 1,

    benefitStartYear: 1,
    effectDurationYears: 5,
    decay: "linear", // none|linear|exp
    halfLifeYears: 2,

    costTimingDefault: "y1_only" // y1_only|y0_only|annual_duration|annual_horizon|split_50_50
  },
  // treatmentOverrides[name] = {enabled, adoption, yieldMult, costMult, costTiming}
  treatmentOverrides: {},
  ui: {
    activeTab: "import",
    rankBy: "npv",
    view: "whole_farm",
    cashflowTreatment: null
  },
  cache: {
    trialSummary: null,        // per-treatment means and deltas
    cbaPerTreatment: null,     // per-treatment cba results + cashflows
    lastSensitivity: null
  }
};

/* =========================
   DOM helpers
   ========================= */
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

function escapeHtml(s){
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

function toNum(v){
  if (v === null || v === undefined) return NaN;
  if (typeof v === "number") return v;
  if (typeof v === "string"){
    const s = v.replace(/[, ]+/g, "").replace(/^\$/,"");
    const n = Number(s);
    return Number.isFinite(n) ? n : NaN;
  }
  return NaN;
}

function fmtNumber(x, digits=2){
  if (x === null || x === undefined || Number.isNaN(x)) return "—";
  const n = Number(x);
  if (!Number.isFinite(n)) return "—";
  return n.toLocaleString(undefined, {maximumFractionDigits: digits, minimumFractionDigits: digits});
}

function fmtInt(x){
  if (x === null || x === undefined || Number.isNaN(x)) return "—";
  const n = Number(x);
  if (!Number.isFinite(n)) return "—";
  return n.toLocaleString(undefined, {maximumFractionDigits: 0});
}

function fmtMoney(x, digits=0){
  if (x === null || x === undefined || Number.isNaN(x)) return "—";
  const n = Number(x);
  if (!Number.isFinite(n)) return "—";
  const abs = Math.abs(n);
  const s = abs.toLocaleString(undefined, {minimumFractionDigits: digits, maximumFractionDigits: digits});
  return (n < 0 ? "-$" : "$") + s;
}

function toast(msg, ms=2200){
  const t = $("#toast");
  const tt = $("#toastText");
  tt.textContent = msg;
  t.hidden = false;
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => { t.hidden = true; }, ms);
}

/* =========================
   Tabs
   ========================= */
function setActiveTab(tab){
  state.ui.activeTab = tab;
  $$(".tab").forEach(btn => {
    const on = btn.dataset.tab === tab;
    btn.classList.toggle("is-active", on);
    btn.setAttribute("aria-selected", on ? "true" : "false");
  });
  $$(".panel").forEach(p => {
    const on = p.dataset.panel === tab;
    p.classList.toggle("is-active", on);
  });
}

/* =========================
   Excel import
   ========================= */
function isEmptyCell(v){
  return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
}

function densestRowIndex(rows2d, maxScan=60){
  const lim = Math.min(rows2d.length, maxScan);
  let best = {idx: 0, score: -1};
  for (let i=0; i<lim; i++){
    const row = rows2d[i] || [];
    const score = row.reduce((acc, v) => acc + (isEmptyCell(v) ? 0 : 1), 0);
    if (score > best.score){
      best = {idx: i, score};
    }
  }
  return best.idx;
}

function dedupeHeaders(headers){
  const seen = new Map();
  return headers.map(h => {
    const name = (h === null || h === undefined) ? "" : String(h).trim();
    const base = name === "" ? "col" : name;
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count === 0 ? base : `${base}__${count}`;
  });
}

function sheetToRows2D(wb){
  const firstSheet = wb.SheetNames[0];
  const ws = wb.Sheets[firstSheet];
  return XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: null});
}

function rows2DToObjects(rows2d){
  const headerIdx = densestRowIndex(rows2d, 60);
  const headersRaw = rows2d[headerIdx] || [];
  const headers = dedupeHeaders(headersRaw);

  const out = [];
  for (let r=headerIdx+1; r<rows2d.length; r++){
    const row = rows2d[r] || [];
    const nonEmpty = row.reduce((acc,v)=>acc + (isEmptyCell(v)?0:1), 0);
    if (nonEmpty === 0) continue;

    const obj = {};
    for (let c=0; c<headers.length; c++){
      obj[headers[c]] = (c < row.length) ? row[c] : null;
    }

    // Drop very sparse/footer-ish rows
    const filled = Object.values(obj).reduce((acc,v)=>acc + (isEmptyCell(v)?0:1), 0);
    if (filled < Math.max(3, Math.floor(headers.length*0.03))) continue;

    out.push(obj);
  }

  return {rows: out, columns: headers};
}

function unique(arr){
  const s = new Set(arr);
  return Array.from(s);
}

function detectDefaultMappings(rows, columns){
  const cols = columns || (rows[0] ? Object.keys(rows[0]) : []);

  // treatment: prefer Amendment, then Treatment/Trt
  const treatmentKey =
    (cols.includes("Amendment") ? "Amendment" : null) ||
    (cols.includes("Treatment") ? "Treatment" : null) ||
    (cols.includes("Trt") ? "Trt" : null) ||
    (cols.find(c => /treat|amend/i.test(c)) || null);

  // yield: prefer "Yield t/ha", else first starting with Yield
  let yieldKey = null;
  if (cols.includes("Yield t/ha")) yieldKey = "Yield t/ha";
  else yieldKey = cols.find(c => /^yield\b/i.test(c)) || null;

  // cost: prefer "|" else cost-like column with broad numeric coverage
  let costKey = cols.includes("|") ? "|" : null;
  if (!costKey){
    const candidates = cols.filter(c => /cost|\$|\/ha|per ha|ha\b/i.test(c));
    const scored = candidates.map(c => {
      const nums = rows.map(r => toNum(r[c])).filter(v => Number.isFinite(v));
      const coverage = nums.length / Math.max(1, rows.length);
      const p50 = nums.length ? percentile(nums, 50) : -Infinity;
      return {c, coverage, p50};
    }).filter(x => x.coverage >= 0.5);
    scored.sort((a,b) => (b.p50 - a.p50) || (b.coverage - a.coverage));
    costKey = scored[0]?.c || null;
  }

  // optional: Protein if exists
  const optional1 =
    (cols.includes("Protein") ? "Protein" : null) ||
    (cols.find(c => /protein|quality/i.test(c)) || null);

  // baseline: prefer "Control" treatment label
  let baseline = null;
  if (treatmentKey){
    const groups = unique(rows.map(r => String(r[treatmentKey] ?? "").trim()).filter(Boolean));
    baseline = groups.find(g => g.toLowerCase() === "control") || groups[0] || null;
  }

  return {treatmentKey, yieldKey, costKey, optional1, baseline};
}

function percentile(arr, p){
  if (!arr.length) return NaN;
  const a = arr.slice().sort((x,y)=>x-y);
  const idx = (p/100)*(a.length-1);
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  if (lo === hi) return a[lo];
  const t = idx - lo;
  return a[lo]*(1-t) + a[hi]*t;
}

async function importExcelArrayBuffer(buf, name="workbook.xlsx"){
  if (!window.XLSX) throw new Error("XLSX library missing. Check index.html include.");
  const wb = XLSX.read(buf, {type:"array"});
  const rows2d = sheetToRows2D(wb);
  const {rows, columns} = rows2DToObjects(rows2d);

  state.raw = rows;
  state.columns = columns;
  state.source = {kind:"excel", name};

  const d = detectDefaultMappings(rows, columns);
  state.map.treatment = d.treatmentKey;
  state.map.yield = d.yieldKey;
  state.map.cost = d.costKey;
  state.map.optional1 = d.optional1;
  state.map.baseline = d.baseline;

  initTreatmentOverrides(); // create defaults for each treatment
  invalidateCache();
  toast(`Imported: ${name}`);
  renderAll();
}

function loadEmbedded(){
  state.raw = EMBEDDED_LOCKHART_ROWS.map(r => ({...r}));
  state.columns = dedupeHeaders(Object.keys(state.raw[0] || {}));
  state.source = {kind:"embedded", name:"Lockhart (embedded sample)"};

  const d = detectDefaultMappings(state.raw, state.columns);
  state.map.treatment = d.treatmentKey;
  state.map.yield = d.yieldKey;
  state.map.cost = d.costKey;
  state.map.optional1 = d.optional1;
  state.map.baseline = d.baseline;

  initTreatmentOverrides();
  invalidateCache();
  toast("Loaded embedded sample");
  renderAll();
}

async function tryLoadBundled(){
  const candidates = ["./data.xlsx", "./Data for Lockhart-FA-031225 (1).xlsx"];
  for (const url of candidates){
    try{
      const res = await fetch(url);
      if (!res.ok) continue;
      const buf = await res.arrayBuffer();
      await importExcelArrayBuffer(buf, url.split("/").pop());
      return true;
    } catch(e){
      // keep trying
    }
  }
  return false;
}

/* =========================
   Treatment overrides
   ========================= */
function initTreatmentOverrides(){
  state.treatmentOverrides = {};
  const tkey = state.map.treatment;
  if (!tkey || !state.raw.length) return;

  const names = unique(state.raw.map(r => String(r[tkey] ?? "").trim()).filter(Boolean)).sort((a,b)=>a.localeCompare(b));
  for (const n of names){
    state.treatmentOverrides[n] = {
      enabled: true,
      adoption: 1,      // 0..1
      yieldMult: 1,     // multiplier on incremental yield
      costMult: 1,      // multiplier on incremental cost
      costTiming: "default" // default|y1_only|y0_only|annual_duration|annual_horizon|split_50_50
    };
  }
  // Ensure baseline enabled + adoption=1 by default
  if (state.map.baseline && state.treatmentOverrides[state.map.baseline]){
    state.treatmentOverrides[state.map.baseline].enabled = true;
    state.treatmentOverrides[state.map.baseline].adoption = 1;
  }
}

function invalidateCache(){
  state.cache.trialSummary = null;
  state.cache.cbaPerTreatment = null;
  // sensitivity cache kept until overwritten
}

/* =========================
   Core computations
   ========================= */
function mean(arr){
  const xs = (arr || []).filter(v => Number.isFinite(v));
  if (!xs.length) return NaN;
  return xs.reduce((a,b)=>a+b,0)/xs.length;
}

function getTreatmentNames(){
  const tkey = state.map.treatment;
  if (!tkey || !state.raw.length) return [];
  return unique(state.raw.map(r => String(r[tkey] ?? "").trim()).filter(Boolean)).sort((a,b)=>a.localeCompare(b));
}

function buildTrialSummary(){
  if (state.cache.trialSummary) return state.cache.trialSummary;

  const rows = state.raw;
  const tkey = state.map.treatment;
  const ykey = state.map.yield;
  const ckey = state.map.cost;
  if (!rows.length || !tkey || !ykey || !ckey) return [];

  const yScale = Number(state.assumptions.yieldScale) || 1;
  const cScale = Number(state.assumptions.costScale) || 1;

  const buckets = new Map();
  for (const r of rows){
    const name = String(r[tkey] ?? "").trim();
    if (!name) continue;
    const y = toNum(r[ykey]) * yScale;
    const c = toNum(r[ckey]) * cScale;
    if (!buckets.has(name)) buckets.set(name, {name, n:0, yields:[], costs:[]});
    const b = buckets.get(name);
    b.n += 1;
    if (Number.isFinite(y)) b.yields.push(y);
    if (Number.isFinite(c)) b.costs.push(c);
  }

  const out = [];
  for (const b of buckets.values()){
    out.push({
      name: b.name,
      n: b.n,
      yield_mean: mean(b.yields),
      cost_mean: mean(b.costs)
    });
  }
  out.sort((a,b)=>a.name.localeCompare(b.name));

  const baseName = state.map.baseline;
  const base = out.find(x => x.name === baseName) || null;

  for (const r of out){
    r.delta_yield = base ? (r.yield_mean - base.yield_mean) : NaN;
    r.delta_cost  = base ? (r.cost_mean  - base.cost_mean)  : NaN;
  }

  state.cache.trialSummary = out;
  return out;
}

function priceForYear(year){
  // year starts at 1..T
  const p1 = Number(state.assumptions.priceYear1) || 0;
  const g = (Number(state.assumptions.priceGrowthPct) || 0) / 100;
  return p1 * Math.pow(1 + g, Math.max(0, year - 1));
}

function discountFactor(t){
  // t in years from 0
  const r = (Number(state.assumptions.discountRatePct) || 0) / 100;
  return 1 / Math.pow(1 + r, t);
}

function benefitFactorByYear(year){
  // year is 1..T
  const start = Math.max(1, Math.floor(Number(state.assumptions.benefitStartYear) || 1));
  const D = Math.max(1, Math.floor(Number(state.assumptions.effectDurationYears) || 1));
  const decay = state.assumptions.decay || "linear";

  const idx = year - start; // 0 at start year
  if (idx < 0) return 0;
  if (idx >= D) return 0;

  if (decay === "none") return 1;

  if (decay === "linear"){
    // 1 at start, to ~0 at end of duration
    if (D === 1) return 1;
    return Math.max(0, 1 - (idx / (D - 1)));
  }

  if (decay === "exp"){
    const hl = Math.max(0.1, Number(state.assumptions.halfLifeYears) || 2);
    const lambda = Math.log(2) / hl;
    return Math.exp(-lambda * idx);
  }

  return 1;
}

function resolveCostTiming(name){
  const ov = state.treatmentOverrides[name];
  if (!ov) return state.assumptions.costTimingDefault;
  const ct = ov.costTiming || "default";
  if (ct === "default") return state.assumptions.costTimingDefault;
  return ct;
}

function buildCashflowsForTreatment(treatmentName){
  const summary = buildTrialSummary();
  const baseName = state.map.baseline;
  const base = summary.find(x => x.name === baseName);
  const tr = summary.find(x => x.name === treatmentName);
  if (!base || !tr) return null;

  const ov = state.treatmentOverrides[treatmentName] || {enabled:true, adoption:1, yieldMult:1, costMult:1};
  const enabled = !!ov.enabled;
  const adoption = clamp01(Number(ov.adoption));
  const yMult = Number(ov.yieldMult) || 1;
  const cMult = Number(ov.costMult) || 1;

  const area = Number(state.assumptions.areaHa) || 0;
  const T = Math.max(1, Math.floor(Number(state.assumptions.horizonYears) || 1));

  // Incremental deltas (per ha) vs baseline
  const dy_perha = (tr.delta_yield || 0) * yMult;
  const dc_perha = (tr.delta_cost  || 0) * cMult;

  // Scale to whole farm via adoption + area
  const scale = (state.ui.view === "per_ha") ? 1 : (area * adoption);

  // benefits start in year 1..T (year0 has no yield benefit)
  const benefits = Array(T+1).fill(0); // index = year, year 0..T
  const costs = Array(T+1).fill(0);

  for (let y=1; y<=T; y++){
    const bf = benefitFactorByYear(y);
    const p = priceForYear(y);
    const b = dy_perha * p * bf;
    benefits[y] = b * scale;
  }

  // Cost timing options apply to incremental cost stream
  const timing = resolveCostTiming(treatmentName);

  if (timing === "y1_only"){
    costs[1] = dc_perha * scale;
  } else if (timing === "y0_only"){
    costs[0] = dc_perha * scale;
  } else if (timing === "annual_duration"){
    const start = 1; // apply costs from year1
    for (let y=1; y<=T; y++){
      const bf = benefitFactorByYear(y);
      if (bf > 0) costs[y] = dc_perha * scale;
    }
  } else if (timing === "annual_horizon"){
    for (let y=1; y<=T; y++){
      costs[y] = dc_perha * scale;
    }
  } else if (timing === "split_50_50"){
    costs[0] = 0.5 * dc_perha * scale;
    costs[1] += 0.5 * dc_perha * scale;
  } else {
    // fallback
    costs[1] = dc_perha * scale;
  }

  // If disabled, set adoption to 0 effectually (for whole-farm view); per-ha view still shows per-ha
  if (!enabled && state.ui.view !== "per_ha"){
    for (let y=0; y<=T; y++){
      benefits[y] = 0;
      costs[y] = 0;
    }
  }

  const net = benefits.map((b,i)=>b - costs[i]);

  // PV
  let pvB = 0, pvC = 0, npv = 0;
  for (let t=0; t<=T; t++){
    const df = discountFactor(t);
    pvB += benefits[t]*df;
    pvC += costs[t]*df;
    npv += net[t]*df;
  }
  const bcr = (pvC === 0) ? (pvB === 0 ? NaN : Infinity) : (pvB / pvC);
  const roi = (pvC === 0) ? NaN : (npv / pvC);

  return {
    name: treatmentName,
    baseline: baseName,
    enabled: enabled,
    adoption: adoption,
    yieldMult: yMult,
    costMult: cMult,
    timing: timing,
    dy_perha: dy_perha,
    dc_perha: dc_perha,
    benefits,
    costs,
    net,
    pvBenefits: pvB,
    pvCosts: pvC,
    npv: npv,
    bcr: bcr,
    roi: roi
  };
}

function clamp01(x){
  if (!Number.isFinite(x)) return 0;
  return Math.max(0, Math.min(1, x));
}

function buildCbaAllTreatments(){
  if (state.cache.cbaPerTreatment) return state.cache.cbaPerTreatment;

  const names = getTreatmentNames();
  const base = state.map.baseline;
  if (!base || !names.length) return [];

  // Ensure baseline exists
  const ordered = [base, ...names.filter(n => n !== base)];
  const out = [];
  for (const n of ordered){
    const cf = buildCashflowsForTreatment(n);
    if (cf) out.push(cf);
  }
  state.cache.cbaPerTreatment = out;
  return out;
}

/* =========================
   Rendering
   ========================= */
function renderAll(){
  renderDatasetPill();
  renderImportKPIs();
  renderColumns();
  renderPreviewTable();

  renderMappingSelectors();
  renderTreatmentsConfigTable();

  renderAssumptionsControls();
  renderSanity();

  renderResults();
  renderCashflows();
  renderSensitivity(); // last-run display if present
  renderHalfLifeVisibility();
}

function renderDatasetPill(){
  const pill = $("#datasetPill");
  if (!state.raw.length){
    pill.textContent = "No data loaded";
    pill.style.borderColor = "rgba(255,255,255,0.12)";
    return;
  }
  const src = state.source?.name || "Dataset";
  pill.textContent = `${src} · ${state.raw.length} rows`;
  pill.style.borderColor = "rgba(125,211,252,0.35)";
}

function renderImportKPIs(){
  $("#kpiRows").textContent = fmtInt(state.raw.length);
  const tkey = state.map.treatment;
  const treatments = tkey ? getTreatmentNames() : [];
  $("#kpiTreatments").textContent = fmtInt(treatments.length);
  $("#kpiBaseline").textContent = state.map.baseline || "—";
  $("#kpiKeys").textContent = (state.map.yield && state.map.cost) ? `${state.map.yield} | ${state.map.cost}` : "—";
}

function renderColumns(){
  const box = $("#columnsChips");
  box.innerHTML = "";
  if (!state.columns.length){
    box.innerHTML = `<div class="muted">No columns yet.</div>`;
    return;
  }
  const frag = document.createDocumentFragment();
  for (const c of state.columns){
    const d = document.createElement("div");
    d.className = "chip";
    d.textContent = c;
    frag.appendChild(d);
  }
  box.appendChild(frag);
}

function renderPreviewTable(){
  const wrap = $("#previewTableWrap");
  const note = $("#previewNote");
  if (!state.raw.length){
    note.hidden = false;
    wrap.hidden = true;
    wrap.innerHTML = "";
    return;
  }
  note.hidden = true;
  wrap.hidden = false;

  const cols = (state.columns || Object.keys(state.raw[0] || {})).slice(0, 10);
  const rows = state.raw.slice(0, 8);

  let html = `<table><thead><tr>`;
  for (const c of cols) html += `<th>${escapeHtml(c)}</th>`;
  html += `</tr></thead><tbody>`;
  for (const r of rows){
    html += `<tr>`;
    for (const c of cols){
      const v = r[c];
      html += `<td class="${typeof v === "number" ? "mono" : ""}">${escapeHtml(v)}</td>`;
    }
    html += `</tr>`;
  }
  html += `</tbody></table>`;
  wrap.innerHTML = html;
}

function fillSelect(sel, options, value, allowEmpty=false, emptyLabel="—"){
  sel.innerHTML = "";
  if (allowEmpty){
    const opt0 = document.createElement("option");
    opt0.value = "";
    opt0.textContent = emptyLabel;
    sel.appendChild(opt0);
  }
  for (const o of options){
    const opt = document.createElement("option");
    opt.value = o;
    opt.textContent = o;
    if (o === value) opt.selected = true;
    sel.appendChild(opt);
  }
  sel.disabled = options.length === 0;
}

function renderMappingSelectors(){
  const cols = state.columns || [];
  fillSelect($("#mapTreatment"), cols, state.map.treatment, false);
  fillSelect($("#mapYield"), cols, state.map.yield, true, "—");
  fillSelect($("#mapCost"), cols, state.map.cost, true, "—");
  fillSelect($("#mapOptional1"), cols, state.map.optional1 || "", true, "—");

  const tnames = getTreatmentNames();
  fillSelect($("#mapBaseline"), tnames, state.map.baseline, false);

  // Also populate cashflow/sensitivity treatment selectors
  fillSelect($("#cashflowTreatment"), tnames, state.ui.cashflowTreatment || state.map.baseline || "", false);
  fillSelect($("#sensTreatment"), tnames, $("#sensTreatment").value || state.map.baseline || "", false);
}

function renderTreatmentsConfigTable(){
  const wrap = $("#treatmentsConfigWrap");
  const tnames = getTreatmentNames();
  if (!tnames.length || !state.map.baseline){
    wrap.innerHTML = `<div class="muted">Load data and set a treatment mapping first.</div>`;
    return;
  }

  const costTimingOptions = [
    {v:"default", label:"Default (from Assumptions)"},
    {v:"y1_only", label:"Year 1 only"},
    {v:"y0_only", label:"Year 0 only"},
    {v:"annual_duration", label:"Annual during effect duration"},
    {v:"annual_horizon", label:"Annual for full horizon"},
    {v:"split_50_50", label:"Split 50% Year 0, 50% Year 1"}
  ];

  let html = `<table>
    <thead><tr>
      <th>Treatment</th>
      <th>Enable</th>
      <th>Adoption (0-1)</th>
      <th>Yield mult</th>
      <th>Cost mult</th>
      <th>Cost timing</th>
    </tr></thead>
    <tbody>`;

  for (const name of tnames){
    const ov = state.treatmentOverrides[name] || {enabled:true, adoption:1, yieldMult:1, costMult:1, costTiming:"default"};
    const isBase = name === state.map.baseline;
    const badge = isBase ? ` <span class="badge">Baseline</span>` : "";
    html += `<tr>
      <td>${escapeHtml(name)}${badge}</td>
      <td class="mono">
        <input type="checkbox" data-ov="enabled" data-name="${escapeHtml(name)}" ${ov.enabled ? "checked":""} />
      </td>
      <td><input class="tinput" type="number" step="0.05" min="0" max="1" data-ov="adoption" data-name="${escapeHtml(name)}" value="${escapeHtml(String(ov.adoption))}" /></td>
      <td><input class="tinput" type="number" step="0.05" data-ov="yieldMult" data-name="${escapeHtml(name)}" value="${escapeHtml(String(ov.yieldMult))}" /></td>
      <td><input class="tinput" type="number" step="0.05" data-ov="costMult" data-name="${escapeHtml(name)}" value="${escapeHtml(String(ov.costMult))}" /></td>
      <td>
        <select class="tselect" data-ov="costTiming" data-name="${escapeHtml(name)}">
          ${costTimingOptions.map(o => `<option value="${o.v}" ${o.v===ov.costTiming ? "selected":""}>${escapeHtml(o.label)}</option>`).join("")}
        </select>
      </td>
    </tr>`;
  }

  html += `</tbody></table>`;
  wrap.innerHTML = html;

  // Bind events for inline edits
  wrap.querySelectorAll("input[data-ov], select[data-ov]").forEach(el => {
    el.addEventListener("change", (e) => {
      const name = e.target.getAttribute("data-name");
      const key = e.target.getAttribute("data-ov");
      if (!state.treatmentOverrides[name]) state.treatmentOverrides[name] = {enabled:true, adoption:1, yieldMult:1, costMult:1, costTiming:"default"};

      if (e.target.type === "checkbox"){
        state.treatmentOverrides[name][key] = e.target.checked;
      } else if (key === "costTiming"){
        state.treatmentOverrides[name][key] = e.target.value;
      } else {
        state.treatmentOverrides[name][key] = Number(e.target.value);
      }

      // Keep baseline enabled
      if (name === state.map.baseline){
        state.treatmentOverrides[name].enabled = true;
      }

      invalidateCache();
      renderResults();
      renderCashflows();
    });
  });
}

function renderAssumptionsControls(){
  $("#assumpArea").value = state.assumptions.areaHa;
  $("#assumpHorizon").value = state.assumptions.horizonYears;
  $("#assumpDiscount").value = state.assumptions.discountRatePct;
  $("#assumpPrice").value = state.assumptions.priceYear1;
  $("#assumpPriceGrowth").value = state.assumptions.priceGrowthPct;
  $("#assumpYieldScale").value = String(state.assumptions.yieldScale);
  $("#assumpCostScale").value = String(state.assumptions.costScale);

  $("#assumpBenefitStart").value = state.assumptions.benefitStartYear;
  $("#assumpDuration").value = state.assumptions.effectDurationYears;
  $("#assumpDecay").value = state.assumptions.decay;
  $("#assumpHalfLife").value = state.assumptions.halfLifeYears;
  $("#assumpCostTimingDefault").value = state.assumptions.costTimingDefault;
}

function renderHalfLifeVisibility(){
  const decay = state.assumptions.decay;
  $("#halfLifeRow").hidden = decay !== "exp";
}

function renderSanity(){
  const box = $("#sanityBox");
  if (!state.raw.length || !state.map.yield || !state.map.cost){
    box.innerHTML = `<div class="muted">Load data and set yield/cost mapping to see diagnostics.</div>`;
    return;
  }

  const yScale = Number(state.assumptions.yieldScale) || 1;
  const cScale = Number(state.assumptions.costScale) || 1;

  const ys = state.raw.map(r => toNum(r[state.map.yield]) * yScale).filter(Number.isFinite);
  const cs = state.raw.map(r => toNum(r[state.map.cost]) * cScale).filter(Number.isFinite);

  const yP50 = percentile(ys, 50), yP5 = percentile(ys, 5), yP95 = percentile(ys, 95);
  const cP50 = percentile(cs, 50), cP5 = percentile(cs, 5), cP95 = percentile(cs, 95);

  const warnings = [];
  if (!Number.isFinite(yP50) || yP50 <= 0) warnings.push("Yield median is non-positive; check yield mapping/scaling.");
  if (!Number.isFinite(cP50)) warnings.push("Cost median invalid; check cost mapping/scaling.");
  if (Number.isFinite(yP95) && yP95 > 50) warnings.push("Yield P95 looks very high for t/ha; consider yield scale.");
  if (Number.isFinite(cP95) && cP95 > 100000) warnings.push("Cost P95 looks extremely high; confirm the cost column is correct.");

  const warnHtml = warnings.length ? `
    <div class="sanity-item">
      <div class="sanity-item__title">Warnings</div>
      <div class="sanity-item__text">${warnings.map(w => `• ${escapeHtml(w)}`).join("<br>")}</div>
    </div>` : "";

  box.innerHTML = `
    <div class="sanity-item">
      <div class="sanity-item__title">Detected mapping</div>
      <div class="sanity-item__text">
        Treatment: <code>${escapeHtml(state.map.treatment || "—")}</code><br>
        Baseline: <code>${escapeHtml(state.map.baseline || "—")}</code><br>
        Yield: <code>${escapeHtml(state.map.yield || "—")}</code><br>
        Cost: <code>${escapeHtml(state.map.cost || "—")}</code>
      </div>
    </div>

    <div class="sanity-item">
      <div class="sanity-item__title">Yield distribution (scaled)</div>
      <div class="sanity-item__text">P5=${escapeHtml(fmtNumber(yP5,3))}, P50=${escapeHtml(fmtNumber(yP50,3))}, P95=${escapeHtml(fmtNumber(yP95,3))}</div>
    </div>

    <div class="sanity-item">
      <div class="sanity-item__title">Cost distribution (scaled)</div>
      <div class="sanity-item__text">P5=${escapeHtml(fmtMoney(cP5,0))}, P50=${escapeHtml(fmtMoney(cP50,0))}, P95=${escapeHtml(fmtMoney(cP95,0))}</div>
    </div>

    ${warnHtml}
  `;
}

function renderTrialSummary(){
  const wrap = $("#trialSummaryWrap");
  const sum = buildTrialSummary();
  if (!sum.length){
    wrap.innerHTML = `<div class="muted">Load data and confirm mappings.</div>`;
    return;
  }

  let html = `<table>
    <thead><tr>
      <th>Treatment</th>
      <th>N</th>
      <th>Yield mean (t/ha)</th>
      <th>Cost mean (/ha)</th>
      <th>Δ Yield vs baseline</th>
      <th>Δ Cost vs baseline</th>
    </tr></thead><tbody>`;

  for (const r of sum){
    const isBase = r.name === state.map.baseline;
    const badge = isBase ? ` <span class="badge">Baseline</span>` : "";
    html += `<tr>
      <td>${escapeHtml(r.name)}${badge}</td>
      <td class="mono">${fmtInt(r.n)}</td>
      <td class="mono">${fmtNumber(r.yield_mean,3)}</td>
      <td class="mono">${fmtMoney(r.cost_mean,0)}</td>
      <td class="mono">${fmtNumber(r.delta_yield,3)}</td>
      <td class="mono">${fmtMoney(r.delta_cost,0)}</td>
    </tr>`;
  }
  html += `</tbody></table>`;
  wrap.innerHTML = html;
}

function verticalCbaTable(cbaList){
  // Indicators as rows, treatments as columns. Control included.
  const view = state.ui.view; // whole_farm|per_ha
  const names = cbaList.map(x => x.name);

  const rows = [
    {label:"PV benefits", key:"pvBenefits"},
    {label:"PV costs", key:"pvCosts"},
    {label:"NPV", key:"npv"},
    {label:"BCR", key:"bcr"},
    {label:"ROI", key:"roi"},
    {label:"Rank (by current metric)", key:"rank"}
  ];

  // compute rank
  const metric = state.ui.rankBy;
  const scored = cbaList.map(x => {
    const v = (metric === "npv") ? x.npv : (metric === "bcr" ? x.bcr : x.roi);
    return {name:x.name, v};
  });

  // Rank: highest is 1; handle NaN/Infinity by pushing to end
  scored.sort((a,b) => {
    const av = a.v, bv = b.v;
    const aBad = !Number.isFinite(av);
    const bBad = !Number.isFinite(bv);
    if (aBad && bBad) return a.name.localeCompare(b.name);
    if (aBad) return 1;
    if (bBad) return -1;
    return (bv - av);
  });

  const rankMap = {};
  for (let i=0; i<scored.length; i++) rankMap[scored[i].name] = i+1;

  let html = `<table><thead><tr><th>Indicator</th>`;
  for (const n of names){
    const isBase = n === state.map.baseline;
    html += `<th>${escapeHtml(n)}${isBase ? "<br><span class='muted'>(baseline)</span>":""}</th>`;
  }
  html += `</tr></thead><tbody>`;

  for (const r of rows){
    html += `<tr><td>${escapeHtml(r.label)}</td>`;
    for (const t of cbaList){
      let v = "";
      if (r.key === "rank"){
        v = rankMap[t.name] ? String(rankMap[t.name]) : "—";
      } else if (r.key === "bcr" || r.key === "roi"){
        v = fmtNumber(t[r.key], 3);
      } else {
        v = fmtMoney(t[r.key], 0);
      }
      html += `<td class="mono">${escapeHtml(v)}</td>`;
    }
    html += `</tr>`;
  }
  html += `</tbody></table>`;
  return html;
}

function renderResults(){
  // Guard
  const cbaWrap = $("#cbaVerticalWrap");
  const chart = $("#npvChart");
  if (!cbaWrap || !chart) return;

  if (!state.raw.length || !state.map.treatment || !state.map.baseline || !state.map.yield || !state.map.cost){
    cbaWrap.innerHTML = `<div class="muted">Load data and confirm mappings (treatment, baseline, yield, cost).</div>`;
    $("#trialSummaryWrap").innerHTML = "";
    drawEmpty(chart, "Load data to see NPV chart");
    return;
  }

  const cba = buildCbaAllTreatments();
  if (!cba.length){
    cbaWrap.innerHTML = `<div class="muted">Could not compute results. Check mappings and assumptions.</div>`;
    drawEmpty(chart, "No results yet");
    return;
  }

  // Render tables
  cbaWrap.innerHTML = verticalCbaTable(cba);
  renderTrialSummary();

  // Chart: NPV bars ordered by rank metric
  drawNpvChart(chart, cba);
}

function renderCashflows(){
  const tableWrap = $("#cashflowsTableWrap");
  const chart = $("#cashflowChart");
  if (!tableWrap || !chart) return;

  if (!state.raw.length || !state.map.baseline){
    tableWrap.innerHTML = `<div class="muted">Load data first.</div>`;
    drawEmpty(chart, "Load data to see cashflows");
    return;
  }

  const tnames = getTreatmentNames();
  if (!tnames.length) return;

  // Choose current
  const sel = $("#cashflowTreatment");
  const current = sel.value || state.ui.cashflowTreatment || state.map.baseline;
  state.ui.cashflowTreatment = current;

  const cf = buildCashflowsForTreatment(current);
  if (!cf){
    tableWrap.innerHTML = `<div class="muted">Select a treatment.</div>`;
    drawEmpty(chart, "No cashflows");
    return;
  }

  const T = Math.max(1, Math.floor(Number(state.assumptions.horizonYears) || 1));
  let html = `<table>
    <thead><tr>
      <th>Year</th>
      <th>Price ($/t)</th>
      <th>Benefit factor</th>
      <th>Benefits</th>
      <th>Costs</th>
      <th>Net</th>
      <th>Discount factor</th>
      <th>PV net</th>
    </tr></thead><tbody>`;

  for (let y=0; y<=T; y++){
    const price = (y === 0) ? "—" : fmtNumber(priceForYear(y), 2);
    const bf = (y === 0) ? "—" : fmtNumber(benefitFactorByYear(y), 3);
    const df = fmtNumber(discountFactor(y), 4);
    const pvnet = cf.net[y] * discountFactor(y);
    html += `<tr>
      <td class="mono">${y}</td>
      <td class="mono">${escapeHtml(price)}</td>
      <td class="mono">${escapeHtml(bf)}</td>
      <td class="mono">${escapeHtml(fmtMoney(cf.benefits[y],0))}</td>
      <td class="mono">${escapeHtml(fmtMoney(cf.costs[y],0))}</td>
      <td class="mono">${escapeHtml(fmtMoney(cf.net[y],0))}</td>
      <td class="mono">${escapeHtml(df)}</td>
      <td class="mono">${escapeHtml(fmtMoney(pvnet,0))}</td>
    </tr>`;
  }
  html += `</tbody></table>`;
  tableWrap.innerHTML = html;

  $("#cfPvBenefits").textContent = fmtMoney(cf.pvBenefits, 0);
  $("#cfPvCosts").textContent = fmtMoney(cf.pvCosts, 0);
  $("#cfNpv").textContent = fmtMoney(cf.npv, 0);
  $("#cfBcr").textContent = Number.isFinite(cf.bcr) ? fmtNumber(cf.bcr, 3) : "—";

  drawCashflowChart(chart, cf);
}

function renderSensitivity(){
  const wrap = $("#sensitivityWrap");
  if (!wrap) return;

  if (!state.cache.lastSensitivity){
    wrap.innerHTML = `<div class="muted">Run sensitivity to see results.</div>`;
    return;
  }
  wrap.innerHTML = sensitivityTableHtml(state.cache.lastSensitivity);
}

/* =========================
   Charts (simple canvas)
   ========================= */
function drawEmpty(canvas, msg){
  const ctx = canvas.getContext("2d");
  const w = canvas.width, h = canvas.height;
  ctx.clearRect(0,0,w,h);
  ctx.fillStyle = "rgba(255,255,255,0.72)";
  ctx.font = "16px system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial";
  ctx.fillText(msg, 16, 28);
}

function drawNpvChart(canvas, cbaList){
  const ctx = canvas.getContext("2d");
  const w = canvas.width, h = canvas.height;
  ctx.clearRect(0,0,w,h);

  // order by rank metric
  const metric = state.ui.rankBy;
  const items = cbaList.slice().map(x => {
    const v = (metric === "npv") ? x.npv : (metric === "bcr" ? x.bcr : x.roi);
    return {name:x.name, v, npv:x.npv};
  });

  items.sort((a,b) => {
    const av = a.v, bv = b.v;
    const aBad = !Number.isFinite(av);
    const bBad = !Number.isFinite(bv);
    if (aBad && bBad) return a.name.localeCompare(b.name);
    if (aBad) return 1;
    if (bBad) return -1;
    return (bv - av);
  });

  // Use NPV for bars; baseline included
  const values = items.map(x => x.npv).filter(Number.isFinite);
  const minV = Math.min(0, ...values);
  const maxV = Math.max(0, ...values);

  const padL = 50, padR = 10, padT = 18, padB = 90;
  const plotW = w - padL - padR;
  const plotH = h - padT - padB;

  // axis
  ctx.strokeStyle = "rgba(255,255,255,0.18)";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(padL, padT);
  ctx.lineTo(padL, padT + plotH);
  ctx.lineTo(padL + plotW, padT + plotH);
  ctx.stroke();

  const n = items.length;
  const gap = 6;
  const barW = Math.max(6, (plotW - gap*(n-1)) / n);

  function yScale(v){
    if (maxV === minV) return padT + plotH/2;
    return padT + (maxV - v) * (plotH / (maxV - minV));
  }
  const y0 = yScale(0);

  // bars
  for (let i=0; i<n; i++){
    const x = padL + i*(barW + gap);
    const v = items[i].npv;
    const y = yScale(v);
    const top = Math.min(y, y0);
    const bh = Math.abs(y0 - y);

    ctx.fillStyle = (items[i].name === state.map.baseline) ? "rgba(167,243,208,0.55)" : "rgba(125,211,252,0.55)";
    ctx.fillRect(x, top, barW, bh);

    // labels
    ctx.save();
    ctx.translate(x + barW/2, padT + plotH + 10);
    ctx.rotate(-Math.PI/3);
    ctx.fillStyle = "rgba(255,255,255,0.75)";
    ctx.font = "12px ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas";
    ctx.textAlign = "right";
    ctx.fillText(items[i].name.length > 18 ? (items[i].name.slice(0,18)+"…") : items[i].name, 0, 0);
    ctx.restore();
  }

  // title
  ctx.fillStyle = "rgba(255,255,255,0.86)";
  ctx.font = "14px system-ui, -apple-system, Segoe UI, Roboto";
  ctx.fillText(`NPV chart (rank by ${metric.toUpperCase()})`, 16, 18);
}

function drawCashflowChart(canvas, cf){
  const ctx = canvas.getContext("2d");
  const w = canvas.width, h = canvas.height;
  ctx.clearRect(0,0,w,h);

  const T = cf.net.length - 1;
  const values = cf.net.slice(0).filter(Number.isFinite);
  const minV = Math.min(0, ...values);
  const maxV = Math.max(0, ...values);

  const padL = 50, padR = 10, padT = 18, padB = 50;
  const plotW = w - padL - padR;
  const plotH = h - padT - padB;

  ctx.strokeStyle = "rgba(255,255,255,0.18)";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(padL, padT);
  ctx.lineTo(padL, padT + plotH);
  ctx.lineTo(padL + plotW, padT + plotH);
  ctx.stroke();

  function yScale(v){
    if (maxV === minV) return padT + plotH/2;
    return padT + (maxV - v) * (plotH / (maxV - minV));
  }
  const y0 = yScale(0);

  const n = T + 1;
  const gap = 6;
  const barW = Math.max(8, (plotW - gap*(n-1)) / n);

  for (let i=0; i<n; i++){
    const x = padL + i*(barW + gap);
    const v = cf.net[i];
    const y = yScale(v);
    const top = Math.min(y, y0);
    const bh = Math.abs(y0 - y);

    ctx.fillStyle = (v >= 0) ? "rgba(167,243,208,0.55)" : "rgba(252,165,165,0.45)";
    ctx.fillRect(x, top, barW, bh);

    ctx.fillStyle = "rgba(255,255,255,0.70)";
    ctx.font = "11px ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas";
    ctx.textAlign = "center";
    ctx.fillText(String(i), x + barW/2, padT + plotH + 16);
  }

  ctx.fillStyle = "rgba(255,255,255,0.86)";
  ctx.font = "14px system-ui, -apple-system, Segoe UI, Roboto";
  ctx.fillText(`Net cashflow by year: ${cf.name}`, 16, 18);
}

/* =========================
   Sensitivity
   ========================= */
function computeMetric(cf, metric){
  if (metric === "npv") return cf.npv;
  if (metric === "bcr") return cf.bcr;
  if (metric === "roi") return cf.roi;
  return cf.npv;
}

function cloneAssumptions(){
  return JSON.parse(JSON.stringify(state.assumptions));
}

function runSensitivity(){
  const treatment = $("#sensTreatment").value;
  const metric = $("#sensMetric").value;

  const pricePct = Number($("#sensPricePct").value || 0) / 100;
  const costPct = Number($("#sensCostPct").value || 0) / 100;
  const yieldPct = Number($("#sensYieldPct").value || 0) / 100;
  const discPts = Number($("#sensDiscountPts").value || 0);
  const durDelta = Math.floor(Number($("#sensDurationDelta").value || 0));

  // base
  const baseAssump = cloneAssumptions();
  const baseOv = JSON.parse(JSON.stringify(state.treatmentOverrides));

  function evalWith(modFn){
    const savedA = state.assumptions;
    const savedO = state.treatmentOverrides;

    state.assumptions = cloneAssumptions();
    state.treatmentOverrides = JSON.parse(JSON.stringify(baseOv));

    modFn();

    invalidateCache();
    const cf = buildCashflowsForTreatment(treatment);

    // restore
    state.assumptions = savedA;
    state.treatmentOverrides = savedO;
    invalidateCache();

    return cf ? computeMetric(cf, metric) : NaN;
  }

  const baseVal = evalWith(() => {});

  const res = [
    {
      driver: "Price",
      low: evalWith(() => { state.assumptions.priceYear1 = baseAssump.priceYear1 * (1 - pricePct); }),
      base: baseVal,
      high: evalWith(() => { state.assumptions.priceYear1 = baseAssump.priceYear1 * (1 + pricePct); })
    },
    {
      driver: "Incremental cost multiplier",
      low: evalWith(() => { state.treatmentOverrides[treatment].costMult = (baseOv[treatment].costMult || 1) * (1 - costPct); }),
      base: baseVal,
      high: evalWith(() => { state.treatmentOverrides[treatment].costMult = (baseOv[treatment].costMult || 1) * (1 + costPct); })
    },
    {
      driver: "Incremental yield multiplier",
      low: evalWith(() => { state.treatmentOverrides[treatment].yieldMult = (baseOv[treatment].yieldMult || 1) * (1 - yieldPct); }),
      base: baseVal,
      high: evalWith(() => { state.treatmentOverrides[treatment].yieldMult = (baseOv[treatment].yieldMult || 1) * (1 + yieldPct); })
    },
    {
      driver: "Discount rate",
      low: evalWith(() => { state.assumptions.discountRatePct = Math.max(0, baseAssump.discountRatePct - discPts); }),
      base: baseVal,
      high: evalWith(() => { state.assumptions.discountRatePct = Math.max(0, baseAssump.discountRatePct + discPts); })
    },
    {
      driver: "Effect duration",
      low: evalWith(() => { state.assumptions.effectDurationYears = Math.max(1, baseAssump.effectDurationYears - durDelta); }),
      base: baseVal,
      high: evalWith(() => { state.assumptions.effectDurationYears = Math.max(1, baseAssump.effectDurationYears + durDelta); })
    }
  ];

  const payload = {
    treatment,
    metric,
    inputs: {pricePct, costPct, yieldPct, discPts, durDelta},
    rows: res
  };

  state.cache.lastSensitivity = payload;
  renderSensitivity();
  toast("Sensitivity updated");
}

function sensitivityTableHtml(payload){
  const metric = payload.metric;
  const isRatio = (metric === "bcr" || metric === "roi");

  let html = `<table>
    <thead><tr>
      <th>Driver</th>
      <th>Low</th>
      <th>Base</th>
      <th>High</th>
    </tr></thead><tbody>`;

  for (const r of payload.rows){
    const f = (x) => isRatio ? fmtNumber(x, 3) : fmtMoney(x, 0);
    html += `<tr>
      <td>${escapeHtml(r.driver)}</td>
      <td class="mono">${escapeHtml(f(r.low))}</td>
      <td class="mono">${escapeHtml(f(r.base))}</td>
      <td class="mono">${escapeHtml(f(r.high))}</td>
    </tr>`;
  }

  html += `</tbody></table>`;
  return html;
}

/* =========================
   Exports
   ========================= */
function toCsv(rows, columns){
  const cols = columns || (rows[0] ? Object.keys(rows[0]) : []);
  const esc = (v) => {
    const s = (v === null || v === undefined) ? "" : String(v);
    if (/[",\n]/.test(s)) return `"${s.replaceAll('"','""')}"`;
    return s;
  };
  const lines = [];
  lines.push(cols.map(esc).join(","));
  for (const r of rows){
    lines.push(cols.map(c => esc(r[c])).join(","));
  }
  return lines.join("\n");
}

function downloadBlob(filename, blob){
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function downloadText(filename, text, mime="text/plain"){
  downloadBlob(filename, new Blob([text], {type: mime}));
}

function exportVerticalCsv(){
  const cba = buildCbaAllTreatments();
  if (!cba.length) return toast("Nothing to export");
  // Build a 2D representation: first column indicator, then each treatment
  const metric = state.ui.rankBy;
  const scored = cba.map(x => ({name:x.name, v: (metric==="npv"?x.npv:(metric==="bcr"?x.bcr:x.roi))}))
    .sort((a,b)=> {
      const aBad = !Number.isFinite(a.v);
      const bBad = !Number.isFinite(b.v);
      if (aBad && bBad) return a.name.localeCompare(b.name);
      if (aBad) return 1;
      if (bBad) return -1;
      return (b.v - a.v);
    });
  const rankMap = {};
  for (let i=0; i<scored.length; i++) rankMap[scored[i].name] = i+1;

  const indicators = [
    {Indicator:"PV benefits", key:"pvBenefits"},
    {Indicator:"PV costs", key:"pvCosts"},
    {Indicator:"NPV", key:"npv"},
    {Indicator:"BCR", key:"bcr"},
    {Indicator:"ROI", key:"roi"},
    {Indicator:`Rank (by ${metric.toUpperCase()})`, key:"rank"}
  ];

  const cols = ["Indicator", ...cba.map(x => x.name)];
  const rows = indicators.map(ind => {
    const obj = {Indicator: ind.Indicator};
    for (const t of cba){
      let v = "";
      if (ind.key === "rank") v = rankMap[t.name] || "";
      else v = t[ind.key];
      obj[t.name] = v;
    }
    return obj;
  });

  downloadText("cba_vertical.csv", toCsv(rows, cols), "text/csv");
}

function exportTrialSummaryCsv(){
  const sum = buildTrialSummary();
  if (!sum.length) return toast("Nothing to export");
  const cols = ["name","n","yield_mean","cost_mean","delta_yield","delta_cost"];
  downloadText("trial_summary.csv", toCsv(sum, cols), "text/csv");
}

function exportStateJson(){
  const payload = {
    source: state.source,
    columns: state.columns,
    map: state.map,
    assumptions: state.assumptions,
    treatmentOverrides: state.treatmentOverrides,
    raw: state.raw
  };
  downloadText("tool_state.json", JSON.stringify(payload, null, 2), "application/json");
}

function exportSensitivityXlsx(){
  if (!state.cache.lastSensitivity) return toast("Run sensitivity first");
  if (!window.XLSX) return toast("XLSX library missing");

  const wb = XLSX.utils.book_new();
  const s = state.cache.lastSensitivity;

  const meta = [
    {key:"Treatment", value:s.treatment},
    {key:"Metric", value:s.metric},
    {key:"Price change +/-", value:(s.inputs.pricePct*100)+"%"},
    {key:"Cost mult change +/-", value:(s.inputs.costPct*100)+"%"},
    {key:"Yield mult change +/-", value:(s.inputs.yieldPct*100)+"%"},
    {key:"Discount rate change +/- (pp)", value:s.inputs.discPts},
    {key:"Duration change +/- (years)", value:s.inputs.durDelta}
  ];

  const wsMeta = XLSX.utils.json_to_sheet(meta);
  XLSX.utils.book_append_sheet(wb, wsMeta, "Sensitivity_Meta");

  const ws = XLSX.utils.json_to_sheet(s.rows);
  XLSX.utils.book_append_sheet(wb, ws, "Sensitivity");

  const out = XLSX.write(wb, {bookType:"xlsx", type:"array"});
  downloadBlob("sensitivity.xlsx", new Blob([out], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
}

function exportFullWorkbookXlsx(){
  if (!window.XLSX) return toast("XLSX library missing");
  if (!state.raw.length) return toast("Load data first");

  const wb = XLSX.utils.book_new();

  // Assumptions
  const assumpRows = Object.keys(state.assumptions).map(k => ({key:k, value: state.assumptions[k]}));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(assumpRows), "Assumptions");

  // Trial means
  const trial = buildTrialSummary();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(trial), "TrialMeans");

  // CBA vertical
  const cba = buildCbaAllTreatments();
  if (cba.length){
    const metric = state.ui.rankBy;
    const scored = cba.map(x => ({name:x.name, v: (metric==="npv"?x.npv:(metric==="bcr"?x.bcr:x.roi))}))
      .sort((a,b)=> {
        const aBad = !Number.isFinite(a.v);
        const bBad = !Number.isFinite(b.v);
        if (aBad && bBad) return a.name.localeCompare(b.name);
        if (aBad) return 1;
        if (bBad) return -1;
        return (b.v - a.v);
      });
    const rankMap = {};
    for (let i=0; i<scored.length; i++) rankMap[scored[i].name] = i+1;

    const indicators = [
      {Indicator:"PV benefits", key:"pvBenefits"},
      {Indicator:"PV costs", key:"pvCosts"},
      {Indicator:"NPV", key:"npv"},
      {Indicator:"BCR", key:"bcr"},
      {Indicator:"ROI", key:"roi"},
      {Indicator:`Rank (by ${metric.toUpperCase()})`, key:"rank"}
    ];
    const cols = ["Indicator", ...cba.map(x => x.name)];
    const rows = indicators.map(ind => {
      const obj = {Indicator: ind.Indicator};
      for (const t of cba){
        obj[t.name] = (ind.key === "rank") ? (rankMap[t.name] || "") : t[ind.key];
      }
      return obj;
    });
    const wsVert = XLSX.utils.json_to_sheet(rows, {header: cols});
    XLSX.utils.book_append_sheet(wb, wsVert, "CBA_Vertical");
  }

  // Cashflows (all treatments)
  const cashRows = [];
  const T = Math.max(1, Math.floor(Number(state.assumptions.horizonYears) || 1));
  for (const t of cba){
    for (let y=0; y<=T; y++){
      cashRows.push({
        treatment: t.name,
        year: y,
        price: (y===0 ? "" : priceForYear(y)),
        benefit_factor: (y===0 ? "" : benefitFactorByYear(y)),
        benefits: t.benefits[y],
        costs: t.costs[y],
        net: t.net[y],
        discount_factor: discountFactor(y),
        pv_net: t.net[y] * discountFactor(y)
      });
    }
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cashRows), "Cashflows_All");

  // Sensitivity (last run)
  if (state.cache.lastSensitivity){
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.cache.lastSensitivity.rows), "Sensitivity_Last");
  }

  // Raw data
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.raw), "RawData");

  const out = XLSX.write(wb, {bookType:"xlsx", type:"array"});
  downloadBlob("farming_cba_tool_output.xlsx", new Blob([out], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
}

/* =========================
   Events
   ========================= */
function bindEvents(){
  // Tabs
  $$(".tab").forEach(btn => btn.addEventListener("click", () => setActiveTab(btn.dataset.tab)));

  // Import
  $("#fileInput").addEventListener("change", async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    try{
      const buf = await f.arrayBuffer();
      await importExcelArrayBuffer(buf, f.name);
      setActiveTab("mapping");
    } catch(err){
      console.error(err);
      toast(`Import failed: ${err.message || err}`);
    }
  });

  $("#btnLoadBundledXlsx").addEventListener("click", async () => {
    try{
      const ok = await tryLoadBundled();
      if (!ok) toast("Bundled XLSX not found. Use Import or embedded sample.");
      else setActiveTab("mapping");
    } catch(err){
      console.error(err);
      toast(`Bundled load failed: ${err.message || err}`);
    }
  });

  $("#btnLoadEmbedded").addEventListener("click", () => {
    loadEmbedded();
    setActiveTab("mapping");
  });

  $("#btnReset").addEventListener("click", resetApp);

  $("#btnGoMapping").addEventListener("click", () => setActiveTab("mapping"));
  $("#btnGoResults").addEventListener("click", () => setActiveTab("results"));

  $("#btnGoAssumptions").addEventListener("click", () => setActiveTab("assumptions"));
  $("#btnGoImport").addEventListener("click", () => setActiveTab("import"));
  $("#btnGoResults2").addEventListener("click", () => setActiveTab("results"));

  // Mapping selectors
  $("#mapTreatment").addEventListener("change", (e) => {
    state.map.treatment = e.target.value;
    // baseline options depend on treatment key
    const d = detectDefaultMappings(state.raw, state.columns);
    // if changed, re-init overrides
    initTreatmentOverrides();
    // set baseline to control if exists
    state.map.baseline = d.baseline || state.map.baseline;
    invalidateCache();
    renderAll();
  });

  $("#mapBaseline").addEventListener("change", (e) => {
    state.map.baseline = e.target.value;
    // keep baseline enabled
    if (state.treatmentOverrides[state.map.baseline]){
      state.treatmentOverrides[state.map.baseline].enabled = true;
      state.treatmentOverrides[state.map.baseline].adoption = 1;
    }
    invalidateCache();
    renderAll();
  });

  $("#mapYield").addEventListener("change", (e) => {
    state.map.yield = e.target.value || null;
    invalidateCache();
    renderAll();
  });

  $("#mapCost").addEventListener("change", (e) => {
    state.map.cost = e.target.value || null;
    invalidateCache();
    renderAll();
  });

  $("#mapOptional1").addEventListener("change", (e) => {
    state.map.optional1 = e.target.value || null;
    renderAll();
  });

  // Assumptions inputs
  const bindAssump = (id, key, parser=(x)=>x) => {
    $(id).addEventListener("input", (e) => {
      state.assumptions[key] = parser(e.target.value);
      invalidateCache();
      renderAll();
    });
  };

  bindAssump("#assumpArea", "areaHa", (v)=>Number(v));
  bindAssump("#assumpHorizon", "horizonYears", (v)=>Math.max(1, Math.floor(Number(v||1))));
  bindAssump("#assumpDiscount", "discountRatePct", (v)=>Math.max(0, Number(v)));
  bindAssump("#assumpPrice", "priceYear1", (v)=>Math.max(0, Number(v)));
  bindAssump("#assumpPriceGrowth", "priceGrowthPct", (v)=>Number(v));

  $("#assumpYieldScale").addEventListener("change", (e)=>{ state.assumptions.yieldScale = Number(e.target.value||1); invalidateCache(); renderAll(); });
  $("#assumpCostScale").addEventListener("change", (e)=>{ state.assumptions.costScale = Number(e.target.value||1); invalidateCache(); renderAll(); });

  bindAssump("#assumpBenefitStart", "benefitStartYear", (v)=>Math.max(1, Math.floor(Number(v||1))));
  bindAssump("#assumpDuration", "effectDurationYears", (v)=>Math.max(1, Math.floor(Number(v||1))));
  $("#assumpDecay").addEventListener("change", (e)=>{ state.assumptions.decay = e.target.value; invalidateCache(); renderAll(); });
  bindAssump("#assumpHalfLife", "halfLifeYears", (v)=>Math.max(0.1, Number(v||2)));
  $("#assumpCostTimingDefault").addEventListener("change", (e)=>{ state.assumptions.costTimingDefault = e.target.value; invalidateCache(); renderAll(); });

  $("#btnResetAssumptions").addEventListener("click", () => {
    state.assumptions = {
      areaHa: 100,
      horizonYears: 10,
      discountRatePct: 7,
      priceYear1: 450,
      priceGrowthPct: 0,
      yieldScale: 1,
      costScale: 1,
      benefitStartYear: 1,
      effectDurationYears: 5,
      decay: "linear",
      halfLifeYears: 2,
      costTimingDefault: "y1_only"
    };
    invalidateCache();
    renderAll();
    toast("Assumptions reset");
  });

  // Results controls
  $("#resultMetric").addEventListener("change", (e)=>{ state.ui.rankBy = e.target.value; renderResults(); });
  $("#resultView").addEventListener("change", (e)=>{ state.ui.view = e.target.value; invalidateCache(); renderAll(); });

  $("#btnGoCashflows").addEventListener("click", ()=> setActiveTab("cashflows"));

  // Cashflows
  $("#cashflowTreatment").addEventListener("change", ()=>{ renderCashflows(); });
  $("#btnBackToResults").addEventListener("click", ()=> setActiveTab("results"));

  // Sensitivity
  $("#btnRunSensitivity").addEventListener("click", runSensitivity);
  $("#btnResetSensitivity").addEventListener("click", () => {
    $("#sensPricePct").value = 10;
    $("#sensCostPct").value = 10;
    $("#sensYieldPct").value = 10;
    $("#sensDiscountPts").value = 2;
    $("#sensDurationDelta").value = 2;
    toast("Sensitivity inputs reset");
  });
  $("#btnSensitivityExportXlsx").addEventListener("click", exportSensitivityXlsx);

  // Export
  $("#btnExportXlsx").addEventListener("click", exportFullWorkbookXlsx);
  $("#btnExportVerticalCsv").addEventListener("click", exportVerticalCsv);
  $("#btnExportTrialSummaryCsv").addEventListener("click", exportTrialSummaryCsv);
  $("#btnExportStateJson").addEventListener("click", exportStateJson);
}

/* =========================
   Reset + boot
   ========================= */
function resetApp(){
  state.source = null;
  state.raw = [];
  state.columns = [];
  state.map = {treatment:null, baseline:null, yield:null, cost:null, optional1:null};
  state.assumptions = {
    areaHa: 100,
    horizonYears: 10,
    discountRatePct: 7,
    priceYear1: 450,
    priceGrowthPct: 0,
    yieldScale: 1,
    costScale: 1,
    benefitStartYear: 1,
    effectDurationYears: 5,
    decay: "linear",
    halfLifeYears: 2,
    costTimingDefault: "y1_only"
  };
  state.treatmentOverrides = {};
  state.ui = {activeTab:"import", rankBy:"npv", view:"whole_farm", cashflowTreatment:null};
  state.cache = {trialSummary:null, cbaPerTreatment:null, lastSensitivity: state.cache.lastSensitivity || null};

  const fi = $("#fileInput");
  if (fi) fi.value = "";
  toast("Reset complete");
  renderAll();
  setActiveTab("import");
}

(function boot(){
  bindEvents();

  // Sensitivity defaults
  $("#sensPricePct").value = 10;
  $("#sensCostPct").value = 10;
  $("#sensYieldPct").value = 10;
  $("#sensDiscountPts").value = 2;
  $("#sensDurationDelta").value = 2;

  renderAll();
  setActiveTab("import");

  // Auto-load: try bundled; otherwise embedded
  (async () => {
    try{
      const ok = await tryLoadBundled();
      if (!ok) loadEmbedded();
      // stay on import; data ready
    } catch(e){
      loadEmbedded();
    }
  })();
})();
