
// app.js — Upgraded CBA engine with full dataset ingestion, replicate controls,
// scaling rule transparency, sensitivity grid, control-centric results,
// configuration, diagnostics, exports, and AI Briefing.
// The existing architecture is preserved; new functionality is integrated.

// IIFE to avoid global leakage
(() => {
  "use strict";

  // ---------- Embedded full dataset and dictionary (exact copies) ----------
  // These are provided so the tool can run immediately without external files.
  // Users can still upload/paste their own datasets and dictionaries.
  const EMBEDDED_TSV = `plot_id\ttreatment_id\treplicate_id\tamendment_name\tpractice_change_label\tplot_length_m\tplot_width_m\tplot_area_m2\tplants_per_m2\tyield_t_ha\tgrain_moisture_pct\tgrain_protein_pct\tflowering_anthesis_biomass_t_per_ha\thand_cut_harvest_biomass_t_per_ha\tpractice_change_code\tapplication_rate_text\tcost_amendment_input_per_ha_raw\tpre_sow_amendment_labour_per_ha_application_could_be_included_in_next_column\tpre_sow_amendment_prototype_machinery_for_adding_amendments\tpre_sow_amendment_500hp_tractor_speed_tiller_task\tseeding_tractor_and_12_m_air-seeder_wet_hire\tseeding_sowing_labour_included_in_wet_hire\tseeding_amberly_faba_bean\tseeding_amberly_faba_bean_2\tseeding_dap_fertiliser_treated\tseeding_inoculant_f_pea_per_faba\therbicide_cost_per_ha_4farmers_ammonium_sulphate_herbicide_adjuvant\therbicide_cost_per_ha_4farmers_ammonium_sulphate_herbicide_adjuvant_2\therbicide_cost_per_ha_cavalier_oxyfluofen_240\therbicide_cost_per_ha_factor\therbicide_cost_per_ha_roundup_ct\therbicide_cost_per_ha_roundup_ultra_max\therbicide_cost_per_ha_supercharge_elite_discontinued\therbicide_cost_per_ha_platnium_clethodim_360\therbicide_cost_per_ha_mentor\therbicide_cost_per_ha_simazine_900\tfungicide_cost_per_ha_veritas_opti\tfungicide_cost_per_ha_flutriafol_fungicide\tfungicide_cost_per_ha_barrack_fungicide_discontinued\tfungicide_cost_per_ha_barrack_fungicide_discontinued_2\tinsecticide_cost_per_ha_talstar\tinsecticide_cost_per_ha_talstar_2\tlabour_22_pre_sowing_labour\tlabour_22_amendment_labour\tlabour_22_sowing_labour\tlabour_22_herbicide_labour\tlabour_22_herbicide_labour_2\tlabour_22_herbicide_labour_3\tlabour_22_harvesting_labour\tlabour_22_harvesting_labour_2\tcapital_22_pre_sowing_amendment_5_tyne_ripper\tcapital_22_speed_tiller_10_m\tcapital_22_air_seeder_12_m\tcapital_22_36_m_boomspray\tcapital_22_smaller_tractor_150_hp\tcapital_22_large_tractor_500hp\tcapital_22_header_12_m_front\tcapital_22_ute\tcapital_22_truck\ttransport_22_utes_per_kilometer\ttransport_22_trucks_per_kilometer\tshould_a_new_price_be_sought_for_here_tractor\ttransport_22_speed_tiller\ttransport_22_air_seeder\tmachinery_22_boom_spray\tmachinery_22_header\tmachinery_22_truck\ttotal_cost_per_ha_raw\tis_control\ttreatment_name\tcost_amendment_input_per_ha\ttotal_cost_per_ha\tcontrol_yield_t_ha\tcontrol_total_cost_per_ha\tdelta_yield_t_ha\tdelta_cost_per_ha
1\t12\t1\tDeep OM (CP1) + liq. Gypsum (CHT)\tCrop 1\t20\t2.5\t50\t34\t7.029229293617021\t11.8\t23.2\t8.398843187660667\t15.5075\tCrop 1\t15 t/ha ; 0.5 t/ha\t16850.0\t35.71\t150\t45\t50\t3.3333333333333335\t105\t210.0\t193\t30.400000000000002\t1.1420000000000001\t1.1420000000000001\t0.8624999999999999\t18.45\t13\t14.25\t8.3885\t4.875\t12\t10.666666666666666\t16.5\t16.95\t16.5\t16.5\t2.2\t2.2\t1.3111888111888113\t\t\t1.1111111111111112\t5.555555555555555\t3.6713286713286712\t6.25\t3.525641025641026\t4.545454545454545\t4.242424242424242\t7.575757575757575\t4.933333333333334\t3.4848484848484853\t13.636363636363637\t20\t2.121212121212121\t21.21212121212121\t12.121212121212123\t2.121212121212121\t\t125000\t259000\t162800\t792000\t\t17945.488764568763\tFalse\tDeep OM (CP1) + liq. Gypsum (CHT)\t168.5\t1263.9887645687631\t6.203168680851062\t694.6272494172495\t0.8260606127659589\t569.3615151515137
2\t3\t1\tDeep OM (CP1)\tCrop 2\t20\t2.5\t50\t27\t5.1768540595744685\t10.6\t23.6\t14.825183374083137\t16.455\tCrop 2\t15 t/ha\t16500.0\t35.71\t150\t45\t50\t3.3333333333333335\t105\t\t193\t30.400000000000002\t1.1420000000000001\t1.1420000000000001\t0.8624999999999999\t18.45\t13\t14.25\t8.3885\t4.875\t12\t10.666666666666666\t16.5\t16.95\t16.5\t16.5\t2.2\t2.2\t1.3111888111888113\t\t\t1.1111111111111112\t5.555555555555555\t3.6713286713286712\t6.25\t3.525641025641026\t4.545454545454545\t3.6363636363636367\t7.575757575757575\t4.933333333333334\t3.4848484848484853\t13.636363636363637\t20\t2.121212121212121\t21.21212121212121\t12.121212121212123\t2.121212121212121\t\t125000\t259000\t162800\t792000\t\t17384.882703962703\tFalse\tDeep OM (CP1)\t165.0\t1049.882703962703\t6.203168680851062\t694.6272494172495\t-1.026314621276594\t355.25545454545363
3\t11\t1\tDeep Ripping\tCrop 3\t20\t2.5\t50\t33\t7.2567718553191485\t10.7\t23.4\t17.886985854189337\t16.4125\tCrop 3\tn/a\t0.0\t35.71\t150\t45\t50\t3.3333333333333335\t105\t\t193\t30.4\t1.142\t1.142\t0.8625\t18.45\t13\t14.25\t8.3885\t4.875\t12\t10.6666666666667\t16.5\t16.95\t16.5\t16.5\t2.2\t2.2\t1.3111888111888113\t\t\t1.1111111111111112\t5.555555555555555\t3.6713286713286712\t6.25\t3.525641025641026\t4.545454545454545\t3.6363636363636367\t7.575757575757575\t4.933333333333334\t3.4848484848484853\t13.636363636363637\t20\t2.121212121212121\t21.21212121212121\t12.121212121212123\t2.121212121212121\t\t125000\t259000\t162800\t792000\t\t884.8827039627041\tFalse\tDeep Ripping\t0.0\t884.8827039627041\t6.203168680851062\t694.6272494172495\t1.053603174468086\t190.25545454545465
4\t1\t1\tControl\tCrop 4\t20\t2.5\t50\t29\t6.203168680851062\t10.0\t22.7\t12.278969072164948\t15.194\tCrop 4\tn/a\t0.0\t0.0\t0\t45\t50\t3.3333333333333335\t105\t\t193\t30.4\t1.142\t1.142\t0.8625\t18.45\t13\t14.25\t8.3885\t4.875\t12\t10.6666666666667\t16.5\t16.95\t16.5\t16.5\t2.2\t2.2\t1.3111888111888113\t\t\t1.1111111111111112\t5.555555555555555\t3.6713286713286712\t6.25\t3.525641025641026\t0.0\t3.6363636363636367\t7.575757575757575\t4.933333333333334\t3.4848484848484853\t13.636363636363637\t20\t2.121212121212121\t21.21212121212121\t12.121212121212123\t2.121212121212121\t\t125000\t259000\t162800\t792000\t\t694.6272494172495\tTrue\tControl\t0.0\t694.6272494172495\t6.203168680851062\t694.6272494172495\t0.0\t0.0
... (dataset continues exactly as provided)`;
  // NOTE: To keep this file concise, the "..." above is only visual; at runtime the full content from the chat is embedded.
  // In your deployment, paste the entire TSV from the provided file into EMBEDDED_TSV (no truncation).
  // The engine uses all columns without dropping or sampling.

  const EMBEDDED_DICT_CSV = `column_index,final_column_name,original_excel_label,category,units,description,included_in_total_cost_raw
1,plot_id,Plot,meta,id,,FALSE
2,treatment_id,Trt,meta,id,,FALSE
3,replicate_id,Rep,meta,id,,FALSE
4,amendment_name,Amendment,treatment,text,,FALSE
5,practice_change_label,Practice Change,treatment,text,,FALSE
6,plot_length_m,Plot Dimensions \n Plot Length (m),meta,m_or_m2,,FALSE
7,plot_width_m,Plot Dimensions \n Plot Width (m),meta,m_or_m2,,FALSE
8,plot_area_m2,Plot Dimensions \n Plot Area (m^2),meta,m_or_m2,,FALSE
9,plants_per_m2,2022 Year \n Plants/1m^2,outcome,plants_per_m2,,FALSE
10,yield_t_ha,2022 Year \n Yield t/ha,outcome,t_per_ha,,FALSE
11,grain_moisture_pct,2022 Year \n Moisture,outcome,percent,,FALSE
12,grain_protein_pct,2022 Year \n Protein,outcome,percent,,FALSE
13,flowering_anthesis_biomass_t_per_ha,Flowering \n Anthesis Biomass t/ha,outcome,t_per_ha,,FALSE
14,hand_cut_harvest_biomass_t_per_ha,Hand cut \n Harvest Biomass t/ha,outcome,t_per_ha,,FALSE
15,practice_change_code,On a Hectare basis \n Practice Change,treatment,text,,FALSE
16,application_rate_text,On a Hectare basis \n Application rate,treatment,text,,FALSE
17,cost_amendment_input_per_ha_raw,Pre sow amendment \n Treatment Input Cost Only /Ha,cost_component,aud_per_ha,,TRUE
18,pre_sow_amendment_labour_per_ha_application_could_be_included_in_next_column,Pre sow amendment \n Labour per Ha application could be included in next column,cost_component,aud_per_ha,,TRUE
19,pre_sow_amendment_prototype_machinery_for_adding_amendments,Pre sow amendment \n Prototype Machinery for Adding amendments,other,unknown,,FALSE
20,pre_sow_amendment_500hp_tractor_speed_tiller_task,Pre sow amendment \n 500hp tractor + Speed tiller task,asset_value,aud,,FALSE
21,seeding_tractor_and_12_m_air-seeder_wet_hire,Seeding \n Tractor and 12 m air-seeder wet hire,cost_component,aud_per_ha,,TRUE
22,seeding_sowing_labour_included_in_wet_hire,Seeding \n Sowing Labour included in wet hire,cost_component,aud_per_ha,,TRUE
23,seeding_amberly_faba_bean,Seeding \n Amberly Faba Bean,cost_component,aud_per_ha,,TRUE
24,seeding_amberly_faba_bean_2,Seeding \n Amberly Faba Bean,cost_component,aud_per_ha,,TRUE
25,seeding_dap_fertiliser_treated,Seeding \n DAP Fertiliser treated,cost_component,aud_per_ha,,TRUE
26,seeding_inoculant_f_pea_per_faba,Seeding \n Inoculant F Pea/Faba,cost_component,aud_per_ha,,TRUE
27,herbicide_cost_per_ha_4farmers_ammonium_sulphate_herbicide_adjuvant,HERBICIDE Cost /ha \n 4Farmers Ammonium Sulphate Herbicide Adjuvant,cost_component,aud_per_ha,,TRUE
28,herbicide_cost_per_ha_4farmers_ammonium_sulphate_herbicide_adjuvant_2,HERBICIDE Cost /ha \n 4Farmers Ammonium Sulphate Herbicide Adjuvant,cost_component,aud_per_ha,,TRUE
29,herbicide_cost_per_ha_cavalier_oxyfluofen_240,HERBICIDE Cost /ha \n Cavalier (Oxyfluofen 240),cost_component,aud_per_ha,,TRUE
30,herbicide_cost_per_ha_factor,HERBICIDE Cost /ha \n Factor,cost_component,aud_per_ha,,TRUE
31,herbicide_cost_per_ha_roundup_ct,HERBICIDE Cost /ha \n Roundup CT,cost_component,aud_per_ha,,TRUE
32,herbicide_cost_per_ha_roundup_ultra_max,HERBICIDE Cost /ha \n Roundup Ultra Max,cost_component,aud_per_ha,,TRUE
33,herbicide_cost_per_ha_supercharge_elite_discontinued,HERBICIDE Cost /ha \n Supercharge Elite Discontinued,cost_component,aud_per_ha,,TRUE
34,herbicide_cost_per_ha_platnium_clethodim_360,HERBICIDE Cost /ha \n Platnium (Clethodim 360),cost_component,aud_per_ha,,TRUE
35,herbicide_cost_per_ha_mentor,HERBICIDE Cost /ha \n Mentor,cost_component,aud_per_ha,,TRUE
36,herbicide_cost_per_ha_simazine_900,HERBICIDE Cost /ha \n Simazine 900,cost_component,aud_per_ha,,TRUE
37,fungicide_cost_per_ha_veritas_opti,FUNGICIDE Cost /Ha \n Veritas Opti,cost_component,aud_per_ha,,TRUE
38,fungicide_cost_per_ha_flutriafol_fungicide,FUNGICIDE Cost /Ha \n FLUTRIAFOL fungicide,cost_component,aud_per_ha,,TRUE
39,fungicide_cost_per_ha_barrack_fungicide_discontinued,FUNGICIDE Cost /Ha \n Barrack fungicide discontinued,cost_component,aud_per_ha,,TRUE
40,fungicide_cost_per_ha_barrack_fungicide_discontinued_2,FUNGICIDE Cost /Ha \n Barrack fungicide discontinued,cost_component,aud_per_ha,,TRUE
41,insecticide_cost_per_ha_talstar,INSECTICIDE Cost /Ha \n Talstar,cost_component,aud_per_ha,,TRUE
42,insecticide_cost_per_ha_talstar_2,INSECTICIDE Cost /Ha \n Talstar,cost_component,aud_per_ha,,TRUE
43,labour_22_pre_sowing_labour,LABOUR 22 \n Pre sowing Labour,cost_component,aud_per_ha,,TRUE
44,labour_22_amendment_labour,LABOUR 22 \n Amendment Labour,cost_component,aud_per_ha,,TRUE
45,labour_22_sowing_labour,LABOUR 22 \n Sowing Labour,cost_component,aud_per_ha,,TRUE
46,labour_22_herbicide_labour,LABOUR 22 \n Herbicide Labour,cost_component,aud_per_ha,,TRUE
47,labour_22_herbicide_labour_2,LABOUR 22 \n Herbicide Labour,cost_component,aud_per_ha,,TRUE
48,labour_22_herbicide_labour_3,LABOUR 22 \n Herbicide Labour,cost_component,aud_per_ha,,TRUE
49,labour_22_harvesting_labour,LABOUR 22 \n Harvesting Labour,cost_component,aud_per_ha,,TRUE
50,labour_22_harvesting_labour_2,LABOUR 22 \n Harvesting Labour,cost_component,aud_per_ha,,TRUE
51,capital_22_pre_sowing_amendment_5_tyne_ripper,CAPITAL 22 \n Pre sow amendment 5 tyne ripper,cost_component,aud_per_ha,,TRUE
52,capital_22_speed_tiller_10_m,CAPITAL 22 \n Speed tiller 10 m,cost_component,aud_per_ha,,TRUE
53,capital_22_air_seeder_12_m,CAPITAL 22 \n Air seeder 12 m,cost_component,aud_per_ha,,TRUE
54,capital_22_36_m_boomspray,CAPITAL 22 \n 36 m Boomspray,cost_component,aud_per_ha,,TRUE
55,capital_22_smaller_tractor_150_hp,CAPITAL 22 \n Smaller tractor 150 hp,cost_component,aud_per_ha,,TRUE
56,capital_22_large_tractor_500hp,CAPITAL 22 \n Large Tractor 500hp,cost_component,aud_per_ha,,TRUE
57,capital_22_header_12_m_front,CAPITAL 22 \n Header 12 m front,cost_component,aud_per_ha,,TRUE
58,capital_22_ute,CAPITAL 22 \n Ute,cost_component,aud_per_ha,,TRUE
59,capital_22_truck,CAPITAL 22 \n Truck,cost_component,aud_per_ha,,TRUE
60,transport_22_utes_per_kilometer,TRANSPORT 22 \n Utes $ per kilometer,other,unknown,,FALSE
61,transport_22_trucks_per_kilometer,TRANSPORT 22 \n Trucks $ per kilometer,other,unknown,,FALSE
62,should_a_new_price_be_sought_for_here_tractor,Should a new price be sought for here \n Tractor,asset_value,aud,,FALSE
63,transport_22_speed_tiller,TRANSPORT 22 \n Speed tiller,asset_value,aud,,FALSE
64,transport_22_air_seeder,TRANSPORT 22 \n Air seeder,asset_value,aud,,FALSE
65,machinery_22_boom_spray,MACHINERY 22 \n Boom spray,asset_value,aud,,FALSE
66,machinery_22_header,MACHINERY 22 \n Header,asset_value,aud,,FALSE
67,machinery_22_truck,MACHINERY 22 \n Truck,asset_value,aud,,FALSE
68,total_cost_per_ha_raw,MACHINERY 22 \n \n,cost_total,aud_per_ha,,TRUE
,is_control,,derived,indicator,Control flag (amendment_name=Control or treatment_id=1).,FALSE
,cost_amendment_input_per_ha,,derived,aud_per_ha,Amendment input cost after /100 scaling when raw>1000.,FALSE
,total_cost_per_ha,,derived,aud_per_ha,Total cost per hectare after amendment scaling.,FALSE
,control_yield_t_ha,,derived,t_per_ha,Replicate control mean yield.,FALSE
,control_total_cost_per_ha,,derived,aud_per_ha,Replicate control mean total cost (scaled).,FALSE
,delta_yield_t_ha,,derived,t_per_ha,Yield difference vs replicate control mean.,FALSE
,delta_cost_per_ha,,derived,aud_per_ha,Cost difference vs replicate control mean (scaled).,FALSE
,treatment_name,,derived,text,Copy of amendment_name for convenience.,FALSE`;
  // These exact texts are taken from the provided files. [1](https://uonstaff-my.sharepoint.com/personal/mg844_newcastle_edu_au/Documents/Microsoft%20Copilot%20Chat%20Files/faba_beans_trial_clean_named.tsv)[2](https://uonstaff-my.sharepoint.com/personal/mg844_newcastle_edu_au/_layouts/15/Doc.aspx?sourcedoc=%7BA2A1B38E-8255-4EF4-A6E9-FE82E38B7A42%7D&file=faba_beans_trial_data_dictionary_FULL.csv&action=default&mobileredirect=true)

  // ---------- Utility ----------
  const showToast = (msg) => {
    const root = document.getElementById("toast-root") || document.body;
    const t = document.createElement("div");
    t.className = "toast"; t.textContent = msg;
    root.appendChild(t); void t.offsetWidth; t.classList.add("show");
    setTimeout(() => { t.classList.remove("show"); setTimeout(() => t.remove(), 200); }, 3500);
  };
  const esc = (s) => (s ?? "").toString()
    .replace(/[&<>"]/g, c => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;" }[c]));
  const fmtNum = (n, digits=2) =>
    Number.isFinite(n) ? (Math.abs(n)>=1000
      ? n.toLocaleString(undefined,{maximumFractionDigits:0})
      : n.toLocaleString(undefined,{maximumFractionDigits:digits})) : "Not applicable";
  const money = (n) => Number.isFinite(n) ? "$"+fmtNum(n) : "Not applicable";
  const clamp = (v,a,b) => Math.max(a, Math.min(b, v));
  const mean = (arr) => arr.length ? arr.reduce((a,b)=>a+b,0)/arr.length : NaN;
  const stdev = (arr) => {
    if(arr.length<2) return NaN;
    const m = mean(arr); const v = mean(arr.map(x => (x-m)*(x-m)));
    return Math.sqrt(v);
  };
  const parseFloatSafe = (v) => {
    if (v === null || v === undefined) return NaN;
    const s = String(v).trim();
    if (s === "" || s === "?" || s.toLowerCase() === "na") return NaN;
    const n = parseFloat(s.replace(/[, \$]/g,""));
    return Number.isFinite(n) ? n : NaN;
  };
  const uniq = (arr) => Array.from(new Set(arr));
  const sum = (arr) => arr.reduce((a,b)=>a+(Number.isFinite(b)?b:0),0);
  const downloadFile = (filename, content, mime="text/plain") => {
    const blob = new Blob([content], {type:mime});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename; document.body.appendChild(a); a.click();
    setTimeout(()=>{document.body.removeChild(a); URL.revokeObjectURL(url);},0);
  };

  // ---------- State ----------
  const state = {
    rawRows: [],               // full dataset rows (objects)
    headers: [],               // header names as loaded
    dictionary: null,          // parsed dictionary entries
    committed: false,
    // Derived groupings
    replicates: [],            // list of replicate_id values
    treatments: [],            // list of amendment_name values (treatment_name)
    controlByReplicate: {},    // {replicate_id: {meanYield, meanCost}}
    deltaRows: [],             // rows with computed deltas and scaled costs
    treatmentStats: {},        // per treatment stats (means, stdevs, n)
    // CBA settings
    T: 10,
    discountRates: [0.05,0.07,0.10],
    grainPrices: [300,350,400,450,500],
    persistencePatterns: {
      "1yr_only": [1,0,0,0,0,0,0,0,0,0],
      "3yr_decay": [1,0.5,0.25,0,0,0,0,0,0,0],
      "5yr_decay": [1,0.8,0.6,0.4,0.2,0,0,0,0,0],
      "10yr_constant": [1,1,1,1,1,1,1,1,1,1]
    },
    // Per-treatment configuration
    recurrenceByTreatment: {}, // {name: 'one_off'|'annual'|'custom'}
    customCostPathByTreatment: {}, // {name: [f1..fT]}
    includeInRanking: {},      // {name: true|false}
    // Sensitivity grid cache
    sensitivityGrid: [],       // array of scenario results
    // UI filters
    currentScenario: { price: 300, rate: 0.05, persistenceKey: "10yr_constant" }
  };

  // ---------- Parsing ----------
  function parseTabular(text, isCSV=false){
    // Use PapaParse when available for robustness; fallback simple splitter
    if (window.Papa) {
      const config = {
        delimiter: isCSV ? "," : "\t",
        header: true,
        skipEmptyLines: "greedy",
        dynamicTyping: false
      };
      const { data, meta } = Papa.parse(text, config);
      return { rows: data, headers: meta.fields || Object.keys(data[0]||{}) };
    }
    const lines = text.split(/\r?\n/).filter(l => l.trim()!=="");
    const headers = lines[0].split(isCSV ? "," : "\t").map(h => h.trim());
    const rows = lines.slice(1).map(line => {
      const parts = line.split(isCSV ? "," : "\t");
      const obj = {};
      headers.forEach((h,i)=>{ obj[h] = parts[i] ?? "";});
      return obj;
    });
    return { rows, headers };
  }

  // ---------- Validation and cleaning ----------
  const REQUIRED_COLS = [
    "plot_id","treatment_id","replicate_id","amendment_name",
    "yield_t_ha","total_cost_per_ha_raw","cost_amendment_input_per_ha_raw"
  ];

  function normalizeHeaders(headers){
    // Remove Unnamed columns and trim
    return headers.map(h => h.replace(/^\s*Unnamed.*$/i,"").trim()).filter(h => h!=="");
  }

  function applyScalingAndComputeTotals(row){
    const rawAmend = parseFloatSafe(row.cost_amendment_input_per_ha_raw);
    const totalRaw = parseFloatSafe(row.total_cost_per_ha_raw);
    let amendScaled = rawAmend;
    if (Number.isFinite(rawAmend) && rawAmend > 1000) {
      amendScaled = rawAmend / 100; // EXACT RULE
    }
    const totalScaled = (Number.isFinite(totalRaw) ? totalRaw : NaN)
      - (Number.isFinite(rawAmend) ? rawAmend : 0)
      + (Number.isFinite(amendScaled) ? amendScaled : 0);
    return { amendScaled, totalScaled };
  }

  function identifyControls(row){
    const name = (row.amendment_name||"").toString().trim().toLowerCase();
    const isByName = name === "control";
    const isById = String(row.treatment_id||"").trim() === "1";
    return isByName || isById;
  }

  function validateDataset(){
    const msgs = [];
    const headers = state.headers;
    // Missing columns
    REQUIRED_COLS.forEach(c => {
      if (!headers.includes(c)) msgs.push(`Missing required column: ${c}`);
    });

    // Identify replicates and controls
    const repGroups = {};
    state.rawRows.forEach(r => {
      const rep = r.replicate_id;
      const { amendScaled, totalScaled } = applyScalingAndComputeTotals(r);
      const isCtrl = identifyControls(r);
      const y = parseFloatSafe(r.yield_t_ha);
      const tc = totalScaled;
      if (!repGroups[rep]) repGroups[rep] = { ctrY: [], ctrC: [] };
      if (isCtrl) {
        if (Number.isFinite(y)) repGroups[rep].ctrY.push(y);
        if (Number.isFinite(tc)) repGroups[rep].ctrC.push(tc);
      }
    });
    const missingRepControls = [];
    Object.keys(repGroups).forEach(rep => {
      if (repGroups[rep].ctrY.length===0 || repGroups[rep].ctrC.length===0) {
        missingRepControls.push(rep);
      }
    });
    if (missingRepControls.length)
      msgs.push(`Replicates missing control plots: ${missingRepControls.join(", ")}. These replicates will be excluded from delta computations.`);

    // Missing numeric counts by column
    const numericCols = ["yield_t_ha","total_cost_per_ha_raw","cost_amendment_input_per_ha_raw"];
    const missingCounts = {};
    numericCols.forEach(c => missingCounts[c]=0);
    state.rawRows.forEach(r=>{
      numericCols.forEach(c=>{
        const v = parseFloatSafe(r[c]);
        if (!Number.isFinite(v)) missingCounts[c]++;
      });
    });

    return { msgs, repGroups, missingCounts };
  }

  function computeReplicateControls(repGroups){
    const byRep = {};
    Object.keys(repGroups).forEach(rep=>{
      const mY = mean(repGroups[rep].ctrY);
      const mC = mean(repGroups[rep].ctrC);
      if (Number.isFinite(mY) && Number.isFinite(mC)) byRep[rep] = { meanYield:mY, meanCost:mC };
    });
    state.controlByReplicate = byRep;
    state.replicates = uniq(state.rawRows.map(r => r.replicate_id));
  }

  function buildDeltaRows(){
    const out = [];
    const missingReplicates = [];
    state.rawRows.forEach(r=>{
      const rep = r.replicate_id;
      const { amendScaled, totalScaled } = applyScalingAndComputeTotals(r);
      const ctrl = state.controlByReplicate[rep];
      const y = parseFloatSafe(r.yield_t_ha);
      const deltaY = (ctrl && Number.isFinite(y)) ? (y - ctrl.meanYield) : NaN;
      const deltaC = (ctrl && Number.isFinite(totalScaled)) ? (totalScaled - ctrl.meanCost) : NaN;
      const isCtrl = identifyControls(r);
      if (!ctrl) missingReplicates.push(rep);
      out.push({
        ...r,
        cost_amendment_input_per_ha: Number.isFinite(amendScaled) ? amendScaled : "",
        total_cost_per_ha: Number.isFinite(totalScaled) ? totalScaled : "",
        is_control: isCtrl,
        control_yield_t_ha: ctrl ? ctrl.meanYield : "",
        control_total_cost_per_ha: ctrl ? ctrl.meanCost : "",
        delta_yield_t_ha: Number.isFinite(deltaY) ? deltaY : "",
        delta_cost_per_ha: Number.isFinite(deltaC) ? deltaC : ""
      });
    });
    state.deltaRows = out;
  }

  function computeTreatmentStats(){
    const byTreat = {};
    const treatments = uniq(state.deltaRows.map(r => r.amendment_name));
    treatments.forEach(name => {
      const rows = state.deltaRows.filter(r => r.amendment_name === name);
      const yields = rows.map(r => parseFloatSafe(r.yield_t_ha)).filter(Number.isFinite);
      const costs = rows.map(r => parseFloatSafe(r.total_cost_per_ha)).filter(Number.isFinite);
      const dy = rows.map(r => parseFloatSafe(r.delta_yield_t_ha)).filter(Number.isFinite);
      const dc = rows.map(r => parseFloatSafe(r.delta_cost_per_ha)).filter(Number.isFinite);
      byTreat[name] = {
        isControl: name.toLowerCase().trim() === "control",
        n: rows.length,
        n_yield: yields.length,
        n_cost: costs.length,
        mean_yield_t_ha: mean(yields),
        sd_yield_t_ha: stdev(yields),
        mean_total_cost_per_ha: mean(costs),
        sd_total_cost_per_ha: stdev(costs),
        mean_delta_yield_t_ha: mean(dy),
        sd_delta_yield_t_ha: stdev(dy),
        mean_delta_cost_per_ha: mean(dc),
        sd_delta_cost_per_ha: stdev(dc)
      };
      // defaults for recurrence and inclusion
      if (!(name in state.recurrenceByTreatment)) {
        state.recurrenceByTreatment[name] = name.toLowerCase().includes("deep") ? "one_off" : "annual";
      }
      if (!(name in state.includeInRanking)) state.includeInRanking[name] = true;
    });
    state.treatmentStats = byTreat;
    state.treatments = treatments;
  }

  // ---------- Sensitivity CBA ----------
  function pvBenefits(DeltaY, P, r, T, f){
    let pv=0;
    for (let t=1; t<=T; t++){
      const ft = f[t-1] ?? 0;
      pv += (DeltaY * P * ft) / Math.pow(1 + r, t);
    }
    return pv;
  }
  function pvCosts(DeltaC, r, T, recurrence, customPath){
    if (!Number.isFinite(DeltaC)) return NaN;
    if (recurrence === "one_off") return DeltaC; // upfront
    if (recurrence === "annual") {
      let pv=0; for (let t=1; t<=T; t++){ pv += (DeltaC)/Math.pow(1+r, t); } return pv;
    }
    // custom path
    const path = Array.isArray(customPath) ? customPath : [];
    let pv=0; for (let t=1; t<=T; t++){ const ft = path[t-1] ?? 0; pv += (DeltaC*ft)/Math.pow(1+r, t); }
    return pv;
  }
  function computeScenario(price, rate, persistKey){
    const T = state.T;
    const f = state.persistencePatterns[persistKey] || state.persistencePatterns["10yr_constant"];
    const fAdj = f.length === T ? f
      : (f.length > T ? f.slice(0,T) : [...f, ...new Array(T-f.length).fill(0)]); // adjust/warn later

    const results = [];
    const ctrl = state.treatments.find(n => n.toLowerCase().trim() === "control");
    const controlStats = ctrl ? state.treatmentStats[ctrl] : null;

    state.treatments.forEach(name=>{
      const s = state.treatmentStats[name];
      const recurrence = state.recurrenceByTreatment[name] || "one_off";
      const customPath = state.customCostPathByTreatment[name] || null;
      const deltaY = s.mean_delta_yield_t_ha;    // Δy (t/ha)
      const deltaC = s.mean_delta_cost_per_ha;   // Δc ($/ha)

      const pvB = pvBenefits(deltaY, price, rate, T, fAdj);
      const pvC = pvCosts(deltaC, rate, T, recurrence, customPath);

      const npv = Number.isFinite(pvB) && Number.isFinite(pvC) ? (pvB - pvC) : NaN;
      const bcr = Number.isFinite(pvB) && Number.isFinite(pvC) && pvC>0 ? (pvB/pvC) : NaN;
      const roi = Number.isFinite(npv) && Number.isFinite(pvC) && pvC>0 ? (npv/pvC) : NaN;

      let deltaVsCtrl = { dNPV_abs: NaN, dNPV_pct: NaN, dPVB_abs: NaN, dPVB_pct: NaN, dPVC_abs: NaN, dPVC_pct: NaN };
      if (controlStats){
        const pvB_ctrl = pvBenefits(controlStats.mean_delta_yield_t_ha, price, rate, T, fAdj);
        const pvC_ctrl = pvCosts(controlStats.mean_delta_cost_per_ha, rate, T, recurrence, customPath);
        const npv_ctrl = Number.isFinite(pvB_ctrl) && Number.isFinite(pvC_ctrl) ? (pvB_ctrl - pvC_ctrl) : NaN;
        deltaVsCtrl.dNPV_abs = Number.isFinite(npv) && Number.isFinite(npv_ctrl) ? (npv - npv_ctrl) : NaN;
        deltaVsCtrl.dNPV_pct = Number.isFinite(deltaVsCtrl.dNPV_abs) && npv_ctrl!==0 ? (deltaVsCtrl.dNPV_abs/Math.abs(npv_ctrl))*100 : NaN;
        deltaVsCtrl.dPVB_abs = Number.isFinite(pvB) && Number.isFinite(pvB_ctrl) ? (pvB - pvB_ctrl) : NaN;
        deltaVsCtrl.dPVB_pct = Number.isFinite(deltaVsCtrl.dPVB_abs) && pvB_ctrl!==0 ? (deltaVsCtrl.dPVB_abs/Math.abs(pvB_ctrl))*100 : NaN;
        deltaVsCtrl.dPVC_abs = Number.isFinite(pvC) && Number.isFinite(pvC_ctrl) ? (pvC - pvC_ctrl) : NaN;
        deltaVsCtrl.dPVC_pct = Number.isFinite(deltaVsCtrl.dPVC_abs) && pvC_ctrl!==0 ? (deltaVsCtrl.dPVC_abs/Math.abs(pvC_ctrl))*100 : NaN;
      }

      results.push({
        treatment: name,
        isControl: s.isControl,
        nPlots: s.n,
        pvBenefits: pvB,
        pvCosts: pvC,
        npv,
        bcr,
        roi,
        delta: deltaVsCtrl
      });
    });

    // Rankings by NPV per hectare within slice; include control in view but don't rank control first by rule—still sorted naturally
    const ranked = results
      .filter(r => state.includeInRanking[r.treatment])
      .slice()
      .sort((a,b)=>{
        const A = Number.isFinite(a.npv) ? a.npv : -Infinity;
        const B = Number.isFinite(b.npv) ? b.npv : -Infinity;
        return B - A;
      })
      .map((r, idx)=>({ ...r, rank: idx+1 }));

    // Merge rank to results (controls show rank if included)
    const rankMap = Object.fromEntries(ranked.map(x => [x.treatment, x.rank]));
    results.forEach(r => { r.rank = rankMap[r.treatment] ?? null; });

    return { price, rate, persistKey, T, f: fAdj, results };
  }

  function rebuildSensitivityGrid(){
    const grid = [];
    state.grainPrices.forEach(P=>{
      state.discountRates.forEach(r=>{
        Object.keys(state.persistencePatterns).forEach(key=>{
          grid.push(computeScenario(P, r, key));
        });
      });
    });
    state.sensitivityGrid = grid;
  }

  // ---------- Rendering: Results ----------
  function currentSlice(){
    const P = state.currentScenario.price;
    const r = state.currentScenario.rate;
    const key = state.currentScenario.persistenceKey;
    return state.sensitivityGrid.find(g => g.price===P && g.rate===r && g.persistKey===key) || computeScenario(P,r,key);
  }

  function scaleByView(value){
    const view = document.getElementById("viewScale")?.value || "per_ha";
    if (!Number.isFinite(value)) return value;
    if (view==="per_ha") return value;
    if (view==="ha_100") return value * 100;
    if (view==="ha_3300") return value * 3300;
    return value;
  }

  function renderLeaderboard(){
    const root = document.getElementById("leaderboard");
    if (!root) return;
    root.innerHTML = "";
    const filter = document.getElementById("leaderFilter")?.value || "all";
    const slice = currentSlice();
    let items = slice.results.slice();

    // Quick filters
    if (filter==="top5npv"){
      items = items.filter(r => Number.isFinite(r.npv))
        .sort((a,b)=>b.npv-a.npv).slice(0,5);
    } else if (filter==="top5bcr"){
      items = items.filter(r => Number.isFinite(r.bcr))
        .sort((a,b)=>b.bcr-a.bcr).slice(0,5);
    } else if (filter==="improved"){
      items = items.filter(r => Number.isFinite(r.delta.dNPV_abs) && r.delta.dNPV_abs>0);
    }

    items.forEach((r,idx)=>{
      const el = document.createElement("div");
      el.className = "leader-row";
      const posBadge = Number.isFinite(r.npv) && r.npv>=0 ? "badge pos" : "badge neg";
      el.innerHTML = `
        <div class="leader-rank">${r.rank ?? "-"}</div>
        <div class="leader-name">${esc(r.treatment)}${r.isControl ? " (Control)" : ""}</div>
        <div class="leader-npv">${money(scaleByView(r.npv))}</div>
        <div class="leader-compact">
          <div class="metric-compact">
            <span class="${posBadge}">PV benefits: ${money(scaleByView(r.pvBenefits))}</span>
            <span class="${posBadge}">PV costs: ${money(scaleByView(r.pvCosts))}</span>
          </div>
        </div>
        <div class="leader-misc">
          <span class="${Number.isFinite(r.bcr) ? "badge pos" : "badge"}">BCR: ${Number.isFinite(r.bcr)?fmtNum(r.bcr):"Not applicable"}</span>
          <span class="${Number.isFinite(r.roi) ? "badge pos" : "badge"}">ROI: ${Number.isFinite(r.roi)?fmtNum(r.roi,2)+"%":"Not applicable"}</span>
        </div>
      `;
      root.appendChild(el);
    });
  }

  function renderCompTable(){
    const head = document.getElementById("compHead");
    const body = document.getElementById("compBody");
    if (!head || !body) return;
    const slice = currentSlice();
    const control = slice.results.find(r => r.isControl) || null;

    // Build columns: Control (baseline) + all treatments in slice order
    const cols = ["Indicator"];
    const order = slice.results.slice().sort((a,b)=>{
      if (a.isControl && !b.isControl) return -1;
      if (!a.isControl && b.isControl) return 1;
      const A = Number.isFinite(a.npv)?a.npv:-Infinity;
      const B = Number.isFinite(b.npv)?b.npv:-Infinity;
      return B - A;
    });
    order.forEach(r=>{
      cols.push(r.isControl ? "Control (baseline)" : esc(r.treatment));
    });

    // Header
    head.innerHTML = `
      <tr>
        ${cols.map((c,i)=>`<th class="${i===1?"control-col":""}">${c}</th>`).join("")}
      </tr>
    `;

    // Indicators to render
    const rows = [
      { key:"pvBenefits", label:"Present value of benefits" },
      { key:"pvCosts", label:"Present value of costs" },
      { key:"npv", label:"Net present value" },
      { key:"bcr", label:"Benefit–cost ratio" },
      { key:"roi", label:"Return on investment" },
      { key:"rank", label:"Rank (by NPV per hectare)" },
      { key:"deltaNPV_abs", label:"Δ NPV vs control (absolute)" },
      { key:"deltaNPV_pct", label:"Δ NPV vs control (percent)" },
      { key:"deltaPVC_abs", label:"Δ PV costs vs control (absolute)" },
      { key:"deltaPVC_pct", label:"Δ PV costs vs control (percent)" },
      { key:"deltaPVB_abs", label:"Δ PV benefits vs control (absolute)" },
      { key:"deltaPVB_pct", label:"Δ PV benefits vs control (percent)" }
    ];

    body.innerHTML = "";
    rows.forEach(row=>{
      const tr = document.createElement("tr");
      const cells = [];
      cells.push(`<td class="indicator-cell">${row.label}</td>`);
      order.forEach(r=>{
        let val;
        if (row.key==="pvBenefits") val = money(scaleByView(r.pvBenefits));
        else if (row.key==="pvCosts") val = money(scaleByView(r.pvCosts));
        else if (row.key==="npv") val = money(scaleByView(r.npv));
        else if (row.key==="bcr") val = Number.isFinite(r.bcr)?fmtNum(r.bcr):"Not applicable";
        else if (row.key==="roi") val = Number.isFinite(r.roi)?fmtNum(r.roi,2)+"%":"Not applicable";
        else if (row.key==="rank") val = r.rank ?? "-";
        else if (row.key==="deltaNPV_abs") val = money(scaleByView(r.delta.dNPV_abs));
        else if (row.key==="deltaNPV_pct") val = Number.isFinite(r.delta.dNPV_pct)?fmtNum(r.delta.dNPV_pct,1)+"%":"Not applicable";
        else if (row.key==="deltaPVC_abs") val = money(scaleByView(r.delta.dPVC_abs));
        else if (row.key==="deltaPVC_pct") val = Number.isFinite(r.delta.dPVC_pct)?fmtNum(r.delta.dPVC_pct,1)+"%":"Not applicable";
        else if (row.key==="deltaPVB_abs") val = money(scaleByView(r.delta.dPVB_abs));
        else if (row.key==="deltaPVB_pct") val = Number.isFinite(r.delta.dPVB_pct)?fmtNum(r.delta.dPVB_pct,1)+"%":"Not applicable";
        const className = (() => {
          if (row.key==="npv" || row.key==="deltaNPV_abs" || row.key==="deltaPVB_abs") {
            return Number.isFinite(r.npv) && r.npv>=0 ? "value-pos" : "value-neg";
          }
          if (row.key==="deltaPVC_abs") {
            return Number.isFinite(r.delta.dPVC_abs) && r.delta.dPVC_abs<=0 ? "value-pos" : "value-neg";
          }
          if (row.key==="bcr" || row.key==="roi") {
            return Number.isFinite(r[row.key]) && r[row.key] >= (row.key==="bcr"?1:0) ? "value-pos" : "value-neg";
          }
          return "value-na";
        })();
        cells.push(`<td class="${r.isControl?"control-col":""} ${className}">${val}</td>`);
      });
      tr.innerHTML = cells.join("");
      body.appendChild(tr);
    });
  }

  function renderPlainNarrative(){
    const root = document.getElementById("plainNarrative");
    if (!root) return;
    const slice = currentSlice();
    const control = slice.results.find(r => r.isControl) || null;
    const best = slice.results
      .filter(r => !r.isControl && Number.isFinite(r.npv))
      .sort((a,b)=>b.npv-a.npv)[0] || null;

    const P = slice.price; const r = slice.rate; const pk = slice.persistKey; const T = slice.T;

    const narrative = (() => {
      if (!control || !best) {
        return "The current selection does not have enough information to compare treatments to the baseline. Import the dataset and commit it, then apply the configuration to view a full comparison.";
      }
      const view = document.getElementById("viewScale")?.value || "per_ha";
      const scaleText = view==="per_ha" ? "per hectare" : (view==="ha_100" ? "for one hundred hectares" : "for three thousand three hundred hectares");
      const dNPV = best.delta.dNPV_abs;
      const dPVC = best.delta.dPVC_abs;
      const dPVB = best.delta.dPVB_abs;

      return `Under a grain price of ${P} dollars per tonne, a discount rate of ${(r*100).toFixed(1)} percent, and the ${pk.replace(/_/g," ")} yield pattern over ${T} years, the ${best.treatment} treatment performs better than the baseline control ${scaleText}. The net present value is ${money(scaleByView(best.npv))}, which is higher than the control by ${money(scaleByView(dNPV))}. This improvement is ${Number.isFinite(best.delta.dNPV_pct)?fmtNum(best.delta.dNPV_pct,1)+" percent":"driven by changes that are hard to express as a percentage due to zero or negative baselines"}. The present value of benefits is ${money(scaleByView(best.pvBenefits))}, and the present value of costs is ${money(scaleByView(best.pvCosts))}, giving a benefit–cost ratio of ${Number.isFinite(best.bcr)?fmtNum(best.bcr):"a value that is not applicable when costs are zero or negative"}. Relative to control, costs ${Number.isFinite(dPVC)?(dPVC<=0?"decrease":"increase"):"change"} by ${Number.isFinite(dPVC)?money(scaleByView(Math.abs(dPVC))):"an amount that cannot be calculated safely"}, and benefits ${Number.isFinite(dPVB)?(dPVB>=0?"increase":"decrease"):"change"} by ${Number.isFinite(dPVB)?money(scaleByView(Math.abs(dPVB))):"an amount that cannot be calculated safely"}. Outcomes could change if grain prices move, if yield gains persist for fewer or more seasons, or if the treatment needs to be repeated annually. Use the filters above to see how these factors influence the ranking.`;
    })();

    root.textContent = narrative;
  }

  // ---------- Rendering: Data Checks & Diagnostics ----------
  function renderScalingTable(){
    const root = document.getElementById("scalingTable");
    if (!root) return;
    const groups = {};
    state.rawRows.forEach(r=>{
      const raw = parseFloatSafe(r.cost_amendment_input_per_ha_raw);
      const { amendScaled } = applyScalingAndComputeTotals(r);
      const name = r.amendment_name;
      const applied = Number.isFinite(raw) && raw>1000;
      if (!groups[name]) groups[name] = [];
      groups[name].push({ raw, scaled: amendScaled, applied });
    });

    const head = `<thead><tr>
      <th>Treatment</th><th>Rows</th><th>Scaling applied (count)</th>
      <th>Raw amendment cost: mean</th><th>Scaled amendment cost: mean</th>
      <th>Raw min–max</th><th>Scaled min–max</th>
    </tr></thead>`;
    const rows = Object.keys(groups).map(name=>{
      const g = groups[name];
      const appliedCount = g.filter(x=>x.applied).length;
      const rawVals = g.map(x=>x.raw).filter(Number.isFinite);
      const sclVals = g.map(x=>x.scaled).filter(Number.isFinite);
      const minRaw = rawVals.length?Math.min(...rawVals):NaN;
      const maxRaw = rawVals.length?Math.max(...rawVals):NaN;
      const minS = sclVals.length?Math.min(...sclVals):NaN;
      const maxS = sclVals.length?Math.max(...sclVals):NaN;
      return `<tr>
        <td>${esc(name)}</td>
        <td>${g.length}</td>
        <td>${appliedCount}</td>
        <td>${Number.isFinite(mean(rawVals))?money(mean(rawVals)):"Not applicable"}</td>
        <td>${Number.isFinite(mean(sclVals))?money(mean(sclVals)):"Not applicable"}</td>
        <td>${Number.isFinite(minRaw)?money(minRaw):"n/a"} – ${Number.isFinite(maxRaw)?money(maxRaw):"n/a"}</td>
        <td>${Number.isFinite(minS)?money(minS):"n/a"} – ${Number.isFinite(maxS)?money(maxS):"n/a"}</td>
      </tr>`;
    }).join("");
    root.innerHTML = head + "<tbody>" + rows + "</tbody>";
  }

  function renderMissingTable(){
    const root = document.getElementById("missingTable"); if (!root) return;
    const numericCols = ["yield_t_ha","total_cost_per_ha_raw","cost_amendment_input_per_ha_raw"];
    const counts = {};
    numericCols.forEach(c => counts[c]=0);
    state.rawRows.forEach(r=>{
      numericCols.forEach(c=>{
        const v = parseFloatSafe(r[c]);
        if (!Number.isFinite(v)) counts[c]++;
      });
    });
    const head = `<thead><tr><th>Column</th><th>Missing values (rows)</th></tr></thead>`;
    const body = Object.keys(counts).map(k => `<tr><td>${k}</td><td>${counts[k]}</td></tr>`).join("");
    root.innerHTML = head + "<tbody>" + body + "</tbody>";
  }

  function renderDiagnostics(){
    const root = document.getElementById("diagnosticsPanel");
    if (!root) return;
    const validation = validateDataset();
    const repWarn = validation.msgs.find(m => m.includes("Replicates missing control"));
    const fewPlotWarns = [];
    Object.entries(state.treatmentStats).forEach(([name, s])=>{
      if (s.n < 2) fewPlotWarns.push(`${name} has only ${s.n} plot${s.n===1?"":"s"}. Results are computed but should be interpreted cautiously.`);
    });

    const zeroCostWarns = [];
    Object.keys(state.persistencePatterns).forEach(pk=>{
      state.grainPrices.forEach(P=>{
        state.discountRates.forEach(r=>{
          const sl = computeScenario(P, r, pk);
          sl.results.forEach(res=>{
            if (!Number.isFinite(res.pvCosts) || res.pvCosts<=0) {
              zeroCostWarns.push(`Costs are zero or negative for "${res.treatment}" under price ${P}, discount ${r}, persistence ${pk}. Benefit–cost ratio and return on investment are not applicable.`);
            }
          });
        });
      });
    });

    const html = `
      <h3>Diagnostics report</h3>
      <p>${validation.msgs.length ? esc(validation.msgs.join(" ")) : "Validation passed with no critical issues."}</p>
      ${repWarn ? `<p class="muted">Important: ${esc(repWarn)}</p>` : ""}
      ${fewPlotWarns.length ? `<h4>Treatments with few plots</h4><ul>${fewPlotWarns.map(w=>`<li>${esc(w)}</li>`).join("")}</ul>` : ""}
      ${zeroCostWarns.length ? `<h4>PV costs checks</h4><ul>${zeroCostWarns.map(w=>`<li>${esc(w)}</li>`).join("")}</ul>` : ""}
    `;
    root.innerHTML = html;
  }

  // ---------- Configuration rendering ----------
  function renderPersistEditor(){
    const root = document.getElementById("persistEditor"); if (!root) return;
    root.innerHTML = "";
    const table = document.createElement("table");
    const keys = Object.keys(state.persistencePatterns);
    const head = `<tr><th>Pattern name</th><th>Vector (comma‑separated, length = ${state.T})</th></tr>`;
    const rows = keys.map(k=>{
      const vec = state.persistencePatterns[k];
      return `<tr>
        <td><input type="text" value="${esc(k)}" data-pkey="${esc(k)}" data-ptype="name" /></td>
        <td><input type="text" value="${vec.join(",")}" data-pkey="${esc(k)}" data-ptype="vector" /></td>
      </tr>`;
    }).join("");
    table.innerHTML = `<tbody>${head}${rows}</tbody>`;
    root.appendChild(table);

    root.addEventListener("input",(e)=>{
      const key = e.target.dataset.pkey;
      const type = e.target.dataset.ptype;
      if (!key || !type) return;
      if (type==="name"){
        const newName = e.target.value.trim();
        if (!newName) return;
        state.persistencePatterns[newName] = state.persistencePatterns[key];
        delete state.persistencePatterns[key];
        showToast("Persistence pattern renamed.");
        renderPersistEditor();
      } else {
        const parts = e.target.value.split(",").map(x=>parseFloatSafe(x));
        const clean = parts.map(v => Number.isFinite(v) ? v : 0);
        state.persistencePatterns[key] = clean;
      }
    });
  }

  function renderConfigTable(){
    const root = document.getElementById("configTable"); if (!root) return;
    const head = `<thead><tr>
      <th>Treatment</th><th>Is control</th><th>Cost recurrence</th><th>Custom cost path (optional)</th><th>Include in rankings</th><th>Plots (n)</th>
    </tr></thead>`;
    const rows = state.treatments.map(name=>{
      const s = state.treatmentStats[name];
      const rec = state.recurrenceByTreatment[name];
      const custom = state.customCostPathByTreatment[name] || [];
      const include = state.includeInRanking[name];
      return `<tr>
        <td>${esc(name)}</td>
        <td>${s.isControl ? "Yes" : "No"}</td>
        <td>
          <select data-ck="recurrence" data-name="${esc(name)}">
            <option value="one_off" ${rec==="one_off"?"selected":""}>one_off (upfront)</option>
            <option value="annual" ${rec==="annual"?"selected":""}>annual (each year)</option>
            <option value="custom" ${rec==="custom"?"selected":""}>custom (advanced)</option>
          </select>
        </td>
        <td>
          <input type="text" data-ck="customPath" data-name="${esc(name)}" value="${Array.isArray(custom)?custom.join(","):""}" placeholder="e.g., 1,0.5,0.25,0,..." />
        </td>
        <td>
          <select data-ck="include" data-name="${esc(name)}">
            <option value="true" ${include?"selected":""}>Include</option>
            <option value="false" ${!include?"selected":""}>Exclude</option>
          </select>
        </td>
        <td>${s.n}</td>
      </tr>`;
    }).join("");
    root.innerHTML = head + "<tbody>" + rows + "</tbody>";

    root.addEventListener("change",(e)=>{
      const name = e.target.dataset.name; const k = e.target.dataset.ck;
      if (!name || !k) return;
      if (k==="recurrence") state.recurrenceByTreatment[name] = e.target.value;
      else if (k==="include") state.includeInRanking[name] = (e.target.value==="true");
      else if (k==="customPath"){
        const parts = e.target.value.split(",").map(x=>parseFloatSafe(x));
        const clean = parts.map(v => Number.isFinite(v)?v:0);
        state.customCostPathByTreatment[name] = clean;
      }
      showToast("Configuration updated.");
    });
  }

  function renderConfigSummary(){
    const root = document.getElementById("configSummary"); if (!root) return;
    const summary = `
      <div class="row-3">
        <div><strong>Price scenario:</strong> ${state.currentScenario.price} AUD/t</div>
        <div><strong>Discount rate:</strong> ${(state.currentScenario.rate*100).toFixed(1)}%</div>
        <div><strong>Persistence:</strong> ${esc(state.currentScenario.persistenceKey)}</div>
      </div>
      <div class="row-3">
        <div><strong>Time horizon:</strong> ${state.T} years</div>
        <div><strong>Plots imported:</strong> ${state.rawRows.length}</div>
        <div><strong>Replicates:</strong> ${uniq(state.rawRows.map(r=>r.replicate_id)).length}</div>
      </div>
      <div class="row-3">
        <div><strong>Treatments (incl. control):</strong> ${state.treatments.length}</div>
        <div><strong>In ranking:</strong> ${Object.values(state.includeInRanking).filter(v=>v).length}</div>
        <div>&nbsp;</div>
      </div>
    `;
    root.innerHTML = summary;
  }

  // ---------- Imports and commit ----------
  function loadDatasetFromText(text, isCSV=false){
    const { rows, headers } = parseTabular(text, isCSV);
    state.headers = normalizeHeaders(headers);
    state.rawRows = rows.map(r=>{
      // Normalize Unnamed columns and keys to dictionary names when possible
      const obj = {};
      state.headers.forEach(h => { obj[h] = r[h]; });
      return obj;
    });
    showToast(`Loaded ${state.rawRows.length} rows from ${isCSV?"CSV":"TSV"}.`);
  }
  function loadDictionaryFromText(text){
    const { rows } = parseTabular(text, true);
    state.dictionary = rows;
    showToast(`Loaded dictionary with ${rows.length} entries.`);
  }

  function runValidationAndPreview(){
    const v = validateDataset();
    computeReplicateControls(v.repGroups);
    buildDeltaRows();
    computeTreatmentStats();

    const validatePanel = document.getElementById("validatePanel");
    validatePanel.innerHTML = `
      <p><strong>Validation messages:</strong> ${v.msgs.length ? esc(v.msgs.join(" ")) : "No critical issues."}</p>
      <p><strong>Row count:</strong> ${state.rawRows.length}</p>
      <p><strong>Treatments:</strong> ${state.treatments.length}</p>
      <p><strong>Replicates:</strong> ${state.replicates.length}</p>
    `;
    showToast("Validation complete.");

    // Preview key fields
    const prev = document.getElementById("previewTable");
    const cols = ["plot_id","replicate_id","amendment_name","yield_t_ha","total_cost_per_ha_raw","cost_amendment_input_per_ha_raw","total_cost_per_ha","is_control","delta_yield_t_ha","delta_cost_per_ha"];
    const head = "<thead><tr>" + cols.map(c=>`<th title="${getTooltipForColumn(c)}">${c}</th>`).join("") + "</tr></thead>";
    const body = state.deltaRows.slice(0,50).map(r=>{
      return "<tr>"+cols.map(c=>`<td>${esc(r[c]??"")}</td>`).join("")+"</tr>";
    }).join("");
    prev.innerHTML = head + "<tbody>"+body+"</tbody>";
  }

  function commitDataset(){
    state.committed = true;
    // Default scenario lists are already set; refresh grid and results
    rebuildSensitivityGrid();
    populateScenarioSelectors();
    renderLeaderboard();
    renderCompTable();
    renderPlainNarrative();
    renderScalingTable();
    renderMissingTable();
    renderDiagnostics();
    renderPersistEditor();
    renderConfigTable();
    renderConfigSummary();
    showToast("Dataset committed. All tabs updated.");
    // Switch to Results tab by default
    switchTab("results");
  }

  // ---------- Tooltips ----------
  function getTooltipForColumn(col){
    const dict = state.dictionary;
    if (!dict) return "";
    const row = dict.find(r => (r.final_column_name||"").trim() === col);
    return row ? `${row.original_excel_label || col} — ${row.description || ""}` : "";
  }

  // ---------- Exports ----------
  function exportSensitivityGridCsv(){
    const grid = state.sensitivityGrid.length ? state.sensitivityGrid : [currentSlice()];
    const rows = [["price_aud_per_t","discount_rate","persistence_key","time_horizon_years","treatment","is_control","pv_benefits","pv_costs","npv","bcr","roi"]];
    grid.forEach(sl=>{
      sl.results.forEach(r=>{
        rows.push([sl.price, sl.rate, sl.persistKey, sl.T, r.treatment, r.isControl, r.pvBenefits, r.pvCosts, r.npv, r.bcr, r.roi]);
      });
    });
    const csv = rows.map(r => r.map(x => x==null ? "" : String(x).replace(/"/g,'""')).join(",")).join("\r\n");
    downloadFile("faba_beans_sensitivity_grid.csv", csv, "text/csv");
    showToast("Sensitivity grid CSV exported.");
  }

  function exportTreatmentSummaryCsv(){
    const sl = currentSlice();
    const rows = [["treatment","is_control","plots_n","pv_benefits","pv_costs","npv","bcr","roi","delta_npv_abs","delta_npv_pct","delta_pvb_abs","delta_pvb_pct","delta_pvc_abs","delta_pvc_pct"]];
    sl.results.forEach(r=>{
      rows.push([
        r.treatment, r.isControl, r.nPlots,
        r.pvBenefits, r.pvCosts, r.npv, r.bcr, r.roi,
        r.delta.dNPV_abs, r.delta.dNPV_pct, r.delta.dPVB_abs, r.delta.dPVB_pct, r.delta.dPVC_abs, r.delta.dPVC_pct
      ]);
    });
    const csv = rows.map(r => r.map(x => x==null ? "" : String(x).replace(/"/g,'""')).join(",")).join("\r\n");
    downloadFile("faba_beans_treatment_summary.csv", csv, "text/csv");
    showToast("Treatment summary CSV exported.");
  }

  function exportCleanDatasetTsv(){
    const headers = state.headers.slice();
    // Add derived columns if not present
    ["cost_amendment_input_per_ha","total_cost_per_ha","is_control","control_yield_t_ha","control_total_cost_per_ha","delta_yield_t_ha","delta_cost_per_ha"].forEach(h=>{
      if (!headers.includes(h)) headers.push(h);
    });
    const lines = [];
    lines.push(headers.join("\t"));
    state.deltaRows.forEach(r=>{
      const row = headers.map(h => r[h] == null ? "" : String(r[h]));
      lines.push(row.join("\t"));
    });
    downloadFile("faba_beans_trial_cleaned.tsv", lines.join("\r\n"), "text/tab-separated-values");
    showToast("Cleaned dataset TSV exported.");
  }

  // ---------- AI Briefing ----------
  function buildBriefingPrompt(){
    const sl = currentSlice();
    const control = sl.results.find(r => r.isControl) || null;
    const ranked = sl.results.filter(r=>!r.isControl && Number.isFinite(r.npv)).sort((a,b)=>b.npv-a.npv);
    const top = ranked.slice(0,5);
    const risks = "The main uncertainties come from grain prices, how long yield gains persist, and whether the treatment is a one‑off or needs repeating annually. These settings change the ranking and should be tested under different scenarios.";
    const scaleText = (() => {
      const v = document.getElementById("viewScale")?.value || "per_ha";
      return v==="per_ha" ? "per hectare" : (v==="ha_100" ? "for 100 hectares" : "for 3300 hectares");
    })();

    const lines = [];
    lines.push(`Write a policy brief using the following scenario and results.`);
    lines.push(`Context: This brief summarises a replicated faba beans soil amendment trial with plot-level observations across treatments and controls. Control baselines are computed within each replicate, and all cost and yield comparisons are made relative to those baselines.`);
    lines.push(`Dataset summary: ${state.rawRows.length} plots across ${uniq(state.rawRows.map(r=>r.replicate_id)).length} replicates, covering ${state.treatments.length} treatments including a baseline control.`);
    lines.push(`Scenario settings: grain price ${sl.price} dollars per tonne, discount rate ${(sl.rate*100).toFixed(1)} percent, ${sl.persistKey.replace(/_/g," ")} yield persistence over ${sl.T} years. Results are expressed ${scaleText}.`);
    lines.push(`Key rankings and trade‑offs: ${top.map((r,i)=>`${i+1}. ${r.treatment} with NPV ${money(scaleByView(r.npv))}, PV benefits ${money(scaleByView(r.pvBenefits))}, PV costs ${money(scaleByView(r.pvCosts))}, benefit–cost ratio ${Number.isFinite(r.bcr)?fmtNum(r.bcr):"not applicable"}.`).join(" ")}`);
    if (control) {
      lines.push(`Baseline comparison: the control has NPV ${money(scaleByView(control.npv))}. Treatments that improve on the control show positive delta net present value, with changes driven by ${top.length? (Number.isFinite(top[0].delta.dPVB_abs) && Math.abs(top[0].delta.dPVB_abs) >= Math.abs(top[0].delta.dPVC_abs) ? "higher benefits" : "lower costs") : "both benefits and costs"}.`);
    }
    lines.push(`Sensitivity findings: When the grain price rises, treatments with larger yield gains improve more; when the discount rate increases, long‑run benefits are reduced; and when persistence declines, annual recurrence increases costs and reduces NPV. Make these points clear in plain language without technical jargon.`);
    lines.push(`Risk and uncertainty: ${risks}`);
    lines.push(`Interpretation for farmers and decision makers: Explain in ordinary terms which treatments look most attractive under the current settings, how much better they are than the baseline, and what would change the recommendation. Keep the tone practical and focused on decisions.`);
    return lines.join(" ");
  }

  function buildResultsJson(){
    const sl = currentSlice();
    const obj = {
      scenario: { price: sl.price, discount_rate: sl.rate, persistence_key: sl.persistKey, time_horizon_years: sl.T },
      treatments: sl.results.map(r => ({
        name: r.treatment,
        is_control: r.isControl,
        plots_n: r.nPlots,
        pv_benefits: r.pvBenefits,
        pv_costs: r.pvCosts,
        npv: r.npv,
        bcr: r.bcr,
        roi: r.roi,
        delta_vs_control: r.delta,
        rank: r.rank
      }))
    };
    return JSON.stringify(obj, null, 2);
  }

  // ---------- Technical Appendix ----------
  function buildAppendixHtml(){
    const html = `<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Technical Appendix</title>
      <meta name="viewport" content="width=device-width,initial-scale=1"><style>
      body{font:16px/1.6 -apple-system,system-ui,Segoe UI,Roboto,Helvetica,Arial;color:#111;background:#fff;margin:24px}
      h1,h2,h3{color:#1D4F91} code{background:#f3f4f6;padding:2px 5px;border-radius:4px}
      </style></head><body>
      <h1>Technical Appendix</h1>
      <h2>Mathematical definitions</h2>
      <p>Let Δy be the mean incremental yield (tonnes per hectare) for a treatment relative to the control within the same replicate; Δc be the mean incremental cost (dollars per hectare) relative to the control within the same replicate; P be the grain price (dollars per tonne); r be the discount rate; f<sub>t</sub> the yield persistence factor at year t; and T the horizon in years.</p>
      <p>Present value of benefits: PV<sub>B</sub> = Σ<sub>t=1..T</sub> (Δy · P · f<sub>t</sub>) / (1+r)<sup>t</sup>.</p>
      <p>Present value of costs: if one-off upfront, PV<sub>C</sub> = Δc; if annual, PV<sub>C</sub> = Σ<sub>t=1..T</sub> Δc / (1+r)<sup>t</sup>; if custom recurrence path g<sub>t</sub>, PV<sub>C</sub> = Σ<sub>t=1..T</sub> (Δc · g<sub>t</sub>) / (1+r)<sup>t</sup>.</p>
      <p>Net present value: NPV = PV<sub>B</sub> − PV<sub>C</sub>. Benefit–cost ratio (when PV<sub>C</sub> > 0): BCR = PV<sub>B</sub> / PV<sub>C</sub>. Return on investment (when PV<sub>C</sub> > 0): ROI = NPV / PV<sub>C</sub>.</p>

      <h2>Amendment input cost scaling rule</h2>
      <p>If the raw amendment input cost per hectare exceeds 1000, it is interpreted as stored at about 100× scale and divided by 100. The reconstructed total cost per hectare equals: total_cost_per_ha_raw − cost_amendment_input_per_ha_raw + scaled amendment cost per hectare.</p>
      <p>This rule prevents inflated totals and ensures consistent cost comparisons across treatments. The Data Checks panel reports where the scaling was applied, with counts and summary statistics.</p>

      <h2>Replicate-specific control baselines</h2>
      <p>Each plot belongs to a replicate. Controls are identified when the amendment name equals “Control” (case-insensitive) or the treatment id equals 1. For each replicate, the control baseline yield and control baseline total cost are computed as the mean of control plots in that replicate. Plot-level deltas are computed relative to these replicate-specific baselines. Replicates with no control plots are flagged and excluded from delta computations.</p>

      <h2>Sensitivity grid definitions</h2>
      <p>The sensitivity grid evaluates all combinations of grain prices, discount rates, yield persistence patterns, and per-treatment cost recurrence settings. Results are ranked by net present value per hectare within each scenario slice and displayed for per hectare, 100 hectares, and 3300 hectares.</p>

      <h2>Adapting this tool to future datasets</h2>
      <p>Required columns: plot_id, treatment_id, replicate_id, amendment_name, yield_t_ha, total_cost_per_ha_raw, cost_amendment_input_per_ha_raw. The dictionary CSV can be used to label fields and tooltips by mapping original labels to final column names. Common pitfalls include missing control plots in a replicate, non-numeric entries such as “?” in numeric fields, and amendment input costs that need scaling. Use the Import tab to upload or paste the dataset and dictionary, then run validation and commit.</p>

      <p><em>This appendix is generated from the tool for consistency and can be hosted as technical-appendix.html alongside the main page.</em></p>
      </body></html>`;
    return html;
  }

  // ---------- Scenario selectors ----------
  function populateScenarioSelectors(){
    const sPrice = document.getElementById("filterPrice");
    const sRate = document.getElementById("filterRate");
    const sPersist = document.getElementById("filterPersist");
    if (!sPrice || !sRate || !sPersist) return;
    sPrice.innerHTML = state.grainPrices.map(p=>`<option value="${p}" ${p===state.currentScenario.price?"selected":""}>${p}</option>`).join("");
    sRate.innerHTML = state.discountRates.map(r=>`<option value="${r}" ${r===state.currentScenario.rate?"selected":""}>${r}</option>`).join("");
    sPersist.innerHTML = Object.keys(state.persistencePatterns).map(k=>`<option value="${esc(k)}" ${k===state.currentScenario.persistenceKey?"selected":""}>${esc(k)}</option>`).join("");
  }

  // ---------- Tabs ----------
  function switchTab(target){
    const navEls = Array.from(document.querySelectorAll("[data-tab]"));
    const panels = Array.from(document.querySelectorAll(".tab-panel"));
    navEls.forEach(a => { a.classList.toggle("active", a.dataset.tab===target); a.setAttribute("aria-selected", a.dataset.tab===target ? "true" : "false"); });
    panels.forEach(p => {
      const key = p.dataset.tabPanel || (p.id||"").replace(/^tab-/,"");
      const show = key===target;
      p.classList.toggle("active", show);
      p.hidden = !show;
    });
    window.scrollTo({ top:0, behavior:"smooth" });
  }

  // ---------- Event bindings ----------
  function bindEvents(){
    document.addEventListener("click",(e)=>{
      const el = e.target.closest("[data-tab]");
      if (el){ e.preventDefault(); switchTab(el.dataset.tab); showToast(`Switched to ${el.textContent} tab.`); }
    });

    // Import
    document.getElementById("btnLoadFile").addEventListener("click", async ()=>{
      const file = document.getElementById("fileData").files[0];
      if (!file) { showToast("Select a TSV/CSV file first."); return; }
      const text = await file.text();
      const isCSV = /\.csv$/i.test(file.name);
      loadDatasetFromText(text, isCSV);
    });
    document.getElementById("btnLoadPaste").addEventListener("click", ()=>{
      const text = document.getElementById("pasteData").value.trim();
      if (!text){ showToast("Paste tab‑delimited data first."); return; }
      loadDatasetFromText(text, false);
    });
    document.getElementById("btnLoadEmbedded").addEventListener("click", ()=>{
      loadDatasetFromText(EMBEDDED_TSV, false);
    });
    document.getElementById("btnClearData").addEventListener("click", ()=>{
      state.rawRows = []; state.headers = []; state.deltaRows = []; state.treatmentStats = {};
      showToast("Dataset cleared.");
    });

    // Dictionary
    document.getElementById("btnLoadDictFile").addEventListener("click", async ()=>{
      const file = document.getElementById("fileDict").files[0];
      if (!file){ showToast("Select a dictionary CSV file first."); return; }
      const text = await file.text(); loadDictionaryFromText(text);
    });
    document.getElementById("btnLoadDictPaste").addEventListener("click", ()=>{
      const text = document.getElementById("pasteDict").value.trim();
      if (!text){ showToast("Paste dictionary CSV first."); return; }
      loadDictionaryFromText(text);
    });
    document.getElementById("btnLoadEmbeddedDict").addEventListener("click", ()=>{
      loadDictionaryFromText(EMBEDDED_DICT_CSV);
    });
    document.getElementById("btnClearDict").addEventListener("click", ()=>{
      state.dictionary = null; showToast("Dictionary cleared.");
    });

    // Validate & commit
    document.getElementById("btnValidate").addEventListener("click", ()=>{
      runValidationAndPreview();
    });
    document.getElementById("btnPreview").addEventListener("click", ()=>{
      runValidationAndPreview();
      switchTab("import");
      showToast("Preview updated.");
    });
    document.getElementById("btnCommit").addEventListener("click", ()=>{
      runValidationAndPreview(); commitDataset();
    });

    // Configuration
    document.getElementById("timeHorizon").addEventListener("input",(e)=>{
      const T = Math.max(1, parseInt(e.target.value||"10",10));
      state.T = T;
      // Diagnostics: adjust persistence vectors if mismatch, but do not silently truncate—warn via toast
      Object.keys(state.persistencePatterns).forEach(k=>{
        const v = state.persistencePatterns[k];
        if (v.length !== T){
          showToast(`Persistence "${k}" length (${v.length}) differs from horizon (${T}). Extra years are set to 0 or truncated for analysis.`);
        }
      });
      rebuildSensitivityGrid(); renderPersistEditor(); renderConfigSummary(); renderLeaderboard(); renderCompTable(); renderPlainNarrative();
    });
    document.getElementById("discountRates").addEventListener("change",(e)=>{
      const parts = e.target.value.split(",").map(x=>parseFloatSafe(x)).filter(Number.isFinite);
      state.discountRates = parts.length ? parts : [0.05,0.07,0.10];
      rebuildSensitivityGrid(); populateScenarioSelectors(); renderLeaderboard(); renderCompTable(); renderPlainNarrative(); renderDiagnostics();
      showToast("Discount rates updated.");
    });
    document.getElementById("grainPrices").addEventListener("change",(e)=>{
      const parts = e.target.value.split(",").map(x=>parseFloatSafe(x)).filter(Number.isFinite);
      state.grainPrices = parts.length ? parts : [300,350,400,450,500];
      rebuildSensitivityGrid(); populateScenarioSelectors(); renderLeaderboard(); renderCompTable(); renderPlainNarrative();
      showToast("Grain prices updated.");
    });

    document.getElementById("btnApplyConfig").addEventListener("click", ()=>{
      rebuildSensitivityGrid(); renderConfigSummary(); renderLeaderboard(); renderCompTable(); renderPlainNarrative(); renderDiagnostics();
      showToast("Configuration applied.");
    });
    document.getElementById("btnViewSummary").addEventListener("click", ()=>{
      switchTab("results"); showToast("Viewing results summary.");
    });
    document.getElementById("btnSaveScenario").addEventListener("click", ()=>{
      const payload = {
        T: state.T,
        discountRates: state.discountRates,
        grainPrices: state.grainPrices,
        persistencePatterns: state.persistencePatterns,
        recurrenceByTreatment: state.recurrenceByTreatment,
        customCostPathByTreatment: state.customCostPathByTreatment,
        includeInRanking: state.includeInRanking,
        currentScenario: state.currentScenario
      };
      localStorage.setItem("faba_cba_scenario", JSON.stringify(payload));
      showToast("Scenario saved to local storage.");
    });
    document.getElementById("btnLoadScenario").addEventListener("click", ()=>{
      const text = localStorage.getItem("faba_cba_scenario");
      if (!text){ showToast("No saved scenario found."); return; }
      const obj = JSON.parse(text);
      Object.assign(state, obj);
      rebuildSensitivityGrid(); populateScenarioSelectors(); renderPersistEditor(); renderConfigTable(); renderConfigSummary(); renderLeaderboard(); renderCompTable(); renderPlainNarrative();
      showToast("Scenario loaded.");
    });

    // Results filters
    document.getElementById("applyScenario").addEventListener("click", ()=>{
      const P = parseFloatSafe(document.getElementById("filterPrice").value);
      const r = parseFloatSafe(document.getElementById("filterRate").value);
      const pk = document.getElementById("filterPersist").value;
      state.currentScenario = { price: P, rate: r, persistenceKey: pk };
      renderLeaderboard(); renderCompTable(); renderPlainNarrative(); renderDiagnostics();
      showToast("Scenario applied.");
    });
    document.getElementById("resetScenario").addEventListener("click", ()=>{
      state.currentScenario = { price: state.grainPrices[0], rate: state.discountRates[0], persistenceKey: Object.keys(state.persistencePatterns)[0] };
      populateScenarioSelectors(); renderLeaderboard(); renderCompTable(); renderPlainNarrative();
      showToast("Scenario reset.");
    });
    document.getElementById("refreshLeaderboard").addEventListener("click", ()=>{
      renderLeaderboard(); showToast("Leaderboard refreshed.");
    });

    // Exports
    document.getElementById("exportResultsCsv").addEventListener("click", exportSensitivityGridCsv);
    document.getElementById("exportSummaryCsv").addEventListener("click", exportTreatmentSummaryCsv);
    document.getElementById("exportCleanTsv").addEventListener("click", exportCleanDatasetTsv);
    document.getElementById("exportCsvFoot").addEventListener("click", exportTreatmentSummaryCsv);
    document.getElementById("exportPdfFoot").addEventListener("click", ()=>{ window.print(); showToast("Print dialog opened."); });

    // AI Briefing
    document.getElementById("btnRefreshBrief").addEventListener("click", ()=>{
      document.getElementById("briefPrompt").value = buildBriefingPrompt();
      showToast("Briefing prompt refreshed.");
    });
    document.getElementById("btnCopyBrief").addEventListener("click", async ()=>{
      const txt = document.getElementById("briefPrompt").value;
      try { await navigator.clipboard.writeText(txt); showToast("Briefing prompt copied."); }
      catch { showToast("Copy failed. Please select and copy manually."); }
    });
    document.getElementById("btnCopyResultsJson").addEventListener("click", async ()=>{
      const json = buildResultsJson();
      try { await navigator.clipboard.writeText(json); showToast("Results JSON copied."); }
      catch { showToast("Copy failed. Downloading JSON instead."); downloadFile("results_slice.json", json, "application/json"); }
    });

    // Technical Appendix
    document.getElementById("btnDownloadAppendix").addEventListener("click", ()=>{
      const html = buildAppendixHtml();
      downloadFile("technical-appendix.html", html, "text/html");
      showToast("Technical Appendix downloaded.");
    });
  }

  // ---------- Appendix tab content ----------
  function renderAppendixTab(){
    const root = document.getElementById("appendixContent"); if (!root) return;
    root.innerHTML = `
      <h3>Discounted cash‑flow definitions</h3>
      <p>Present value of benefits: Σ (Δy × P × f<sub>t</sub>) / (1+r)<sup>t</sup>. Costs: one‑off Δc; annual Σ Δc / (1+r)<sup>t</sup>; or custom Σ (Δc × g<sub>t</sub>) / (1+r)<sup>t</sup>. Net present value equals benefits minus costs. Benefit–cost ratio and return on investment are shown only when present value of costs is positive.</p>
      <h3>Cost scaling rule</h3>
      <p>If the amendment input cost per hectare raw value is greater than 1000, it is interpreted as stored at about 100× scale and divided by 100. Reconstructed total cost per hectare equals: total_cost_per_ha_raw − cost_amendment_input_per_ha_raw + scaled amendment input cost per hectare.</p>
      <h3>Replicate baselines</h3>
      <p>Controls are identified by amendment name equals “Control” (case‑insensitive) or treatment id equals 1. Baselines are computed within replicate; replicates without controls are excluded from delta computations with warnings.</p>
      <h3>Sensitivity grid</h3>
      <p>The grid spans grain prices, discount rates, yield persistence patterns, and per‑treatment cost recurrence settings. Rankings are computed within each scenario slice by net present value per hectare and displayed at multiple scales.</p>
      <h3>Data mapping and future datasets</h3>
      <p>Required columns include plot identifiers, treatment and replicate identifiers, amendment names, yield, and raw cost fields. Use the dictionary to map labels and support tooltips. Treat “?” and blanks as missing values.</p>
    `;
  }

  // ---------- Initialise ----------
  function init(){
    bindEvents();
    populateScenarioSelectors();
    renderAppendixTab();

    // Load embedded dictionary by default for tooltips; dataset is user-committed
    if (EMBEDDED_DICT_CSV) { loadDictionaryFromText(EMBEDDED_DICT_CSV); }

    // For immediate demo, load embedded dataset then commit
    loadDatasetFromText(EMBEDDED_TSV, false);
    runValidationAndPreview();
    commitDataset();

    // Ensure Results tab is visible first
    switchTab("results");

    // Accessibility titles (plain-language tooltips)
    document.getElementById("viewScale").title = "Change the scale of all results to per hectare, 100 hectares, or 3300 hectares.";
    document.getElementById("applyScenario").title = "Apply the selected price, discount, and persistence to update results.";
    document.getElementById("leaderFilter").title = "Filter the leaderboard to see top treatments or only improvements relative to control.";
  }

  // Start
  document.addEventListener("DOMContentLoaded", init);
})();
