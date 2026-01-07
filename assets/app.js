/* NKR-KA KYC Tool — Multi-page static app (GitHub Pages)
   Pages: index.html (redirect), upload.html, dashboard.html, actions.html, data.html
   Storage: localStorage (dataset + edits + theme)
*/

const STORAGE_KEY = "kyc_dataset_v1";
const THEME_KEY = "kyc_theme";

const EXPECTED_HEADERS = [
  "Rgn Sl No","Dvn Sl No","sol_id","Office","Division","Account No","cif_id","acct_name",
  "schm_code","acct_opn_date","last_any_tran_date","Status","Consignment number",
  "Date of submission to CPC","Scan/Upload status","Omissions/Rejections"
];

const ACTION_COLS = [
  "Division","Office","sol_id","Account No","Date of submission to CPC",
  "Scan/Upload status","Consignment number","Omissions/Rejections","Flags"
];

const el = (id) => document.getElementById(id);
function pageName(){ return document.body?.dataset?.page || ""; }

/* ---------------- Theme ---------------- */

function applyThemeFromStorage(){
  const t = localStorage.getItem(THEME_KEY) || "dark";
  document.documentElement.setAttribute("data-theme", t);
}
function toggleTheme(){
  const cur = document.documentElement.getAttribute("data-theme") || "dark";
  const next = cur === "dark" ? "light" : "dark";
  localStorage.setItem(THEME_KEY, next);
  document.documentElement.setAttribute("data-theme", next);
}

/* ---------------- Alerts / chips ---------------- */

function showAlert(msg, type="ok"){
  const box = el("alertBox");
  if(!box) return;
  box.classList.remove("hidden","alert--ok","alert--warn","alert--bad");
  box.classList.add(`alert--${type}`);
  box.textContent = msg;
}
function clearAlert(){
  const box = el("alertBox");
  if(!box) return;
  box.classList.add("hidden");
  box.textContent = "";
}

function ensureNotBlankMessage(){
  // If alertBox exists, ensure it’s not “hidden forever” on pages with no data
  const box = el("alertBox");
  if(!box) return;
  box.classList.remove("hidden");
}

function setDataChipText(){
  const chip = el("dataChip");
  if(!chip) return;
  if(!rawRows.length) chip.textContent = "No data loaded";
  else chip.textContent = `Rows: ${rawRows.length}`;
}

/* ---------------- Storage ---------------- */

function saveDataset(rows){
  localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
}
function loadDataset(){
  try{
    const s = localStorage.getItem(STORAGE_KEY);
    if(!s) return [];
    const rows = JSON.parse(s);
    return Array.isArray(rows) ? rows : [];
  }catch{
    return [];
  }
}
function resetDataset(){
  localStorage.removeItem(STORAGE_KEY);
}

/* ---------------- Normalization / parsing ---------------- */

function normalizeValue(v){
  if(v === null || v === undefined) return "";
  if(typeof v === "string") return v.trim();
  return String(v).trim();
}
function normHeader(s){
  return String(s ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/\r?\n/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}
function validateHeaderRow(headerRow){
  const detected = new Set(headerRow.map(normHeader).filter(Boolean));
  const missing = EXPECTED_HEADERS.filter(h => !detected.has(normHeader(h)));
  return { ok: missing.length===0, missing };
}
function countHeaderMatches(rowValues){
  const set = new Set(rowValues.map(normHeader).filter(Boolean));
  let m=0;
  for(const h of EXPECTED_HEADERS) if(set.has(normHeader(h))) m++;
  return m;
}
function findHeaderRowIndex(rowsAOA, scanRows=30){
  let bestIdx=-1, bestScore=-1;
  const max = Math.min(rowsAOA.length, scanRows);
  for(let i=0;i<max;i++){
    const row = (rowsAOA[i]||[]).map(v=>normalizeValue(v)).filter(v=>v!=="");
    if(!row.length) continue;
    const score = countHeaderMatches(row);
    if(score>bestScore){ bestScore=score; bestIdx=i; }
  }
  if(bestIdx>=0 && bestScore>=10){
    return { headerRowIndex: bestIdx, headerRowValues: (rowsAOA[bestIdx]||[]).map(v=>normalizeValue(v)) };
  }
  return null;
}

function parseAnyDate(v){
  if(v === null || v === undefined || v === "") return null;

  if(typeof v === "number" && typeof XLSX !== "undefined"){
    const d = XLSX.SSF.parse_date_code(v);
    if(!d) return null;
    return new Date(d.y, d.m-1, d.d);
  }

  const s = String(v).trim();
  if(!s) return null;

  if(/^\d{4}-\d{2}-\d{2}$/.test(s)){
    const d = new Date(s+"T00:00:00");
    return isNaN(d.getTime()) ? null : d;
  }

  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if(m){
    let dd = parseInt(m[1],10);
    let mm = parseInt(m[2],10);
    let yy = parseInt(m[3],10);
    if(yy < 100) yy += 2000;
    const d = new Date(yy, mm-1, dd);
    return isNaN(d.getTime()) ? null : d;
  }

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
function fmtDateISO(d){
  if(!d) return "";
  const y=d.getFullYear();
  const m=String(d.getMonth()+1).padStart(2,"0");
  const day=String(d.getDate()).padStart(2,"0");
  return `${y}-${m}-${day}`;
}
function startOfDay(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function inRange(d, from, to){
  if(!d) return false;
  const t = startOfDay(d).getTime();
  if(from && t < startOfDay(from).getTime()) return false;
  if(to && t > startOfDay(to).getTime()) return false;
  return true;
}
function hasText(v){ return String(v||"").trim().length>0; }
function toLower(s){ return String(s||"").trim().toLowerCase(); }
function isDoneScan(v){
  const s = toLower(v);
  return ["done","completed","complete","uploaded","scanned","ok","yes"].some(k=>s.includes(k));
}

/* ---------------- Template download (Data + Instructions) ---------------- */

function buildInstructionsAOA(){
  const rows=[];
  rows.push(["NKR KYC Standard Template — Instructions",""]);
  rows.push(["How to use","1) Don’t rename headers  2) Fill data in KYC_Data from row 2  3) Save .xlsx  4) Upload"]);
  rows.push(["Date format","Use DD-MM-YYYY or DD/MM/YYYY (recommended)."]);
  rows.push(["Scan/Upload status","Use: Done / Pending (recommended)."]);
  rows.push(["Omissions/Rejections","Write brief reason; leave blank if none."]);
  rows.push(["",""]);
  rows.push(["Column guide",""]);
  for(const h of EXPECTED_HEADERS) rows.push([h,"(Fill as applicable)"]);
  return rows;
}

function downloadStandardTemplate(){
  if(typeof XLSX === "undefined"){
    showAlert("Cannot generate template because XLSX library did not load.", "bad");
    return;
  }
  const wsData = XLSX.utils.aoa_to_sheet([EXPECTED_HEADERS, EXPECTED_HEADERS.map(()=> "")]);
  wsData["!cols"] = EXPECTED_HEADERS.map(h => ({ wch: Math.max(12, Math.min(32, String(h).length + 2)) }));
  wsData["!freeze"] = { xSplit: 0, ySplit: 1 };

  const wsIns = XLSX.utils.aoa_to_sheet(buildInstructionsAOA());
  wsIns["!cols"] = [{ wch: 28 }, { wch: 80 }];
  wsIns["!freeze"] = { xSplit: 0, ySplit: 1 };

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsData, "KYC_Data");
  XLSX.utils.book_append_sheet(wb, wsIns, "Instructions");

  XLSX.writeFile(wb, "NKR_KYC_Standard_Template.xlsx");
  showAlert("Template downloaded. Fill ‘KYC_Data’ and upload.", "ok");
}

/* ---------------- Upload & import ---------------- */

let rawRows = [];
let filteredRows = [];
let actionRows = [];
let charts = { trend: null, division: null, scan: null };

async function handleFile(file){
  clearAlert();

  if(typeof XLSX === "undefined"){
    showAlert("Upload failed: XLSX library not loaded. Please check CDN access.", "bad");
    return;
  }

  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type:"array", cellDates:true });

  const sheetName = wb.SheetNames.includes("KYC_Data") ? "KYC_Data" : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];

  const rowsAOA = XLSX.utils.sheet_to_json(ws, { header:1, raw:true, defval:"" });
  if(!rowsAOA.length){
    showAlert("No rows found in the sheet.", "bad");
    return;
  }

  const found = findHeaderRowIndex(rowsAOA, 30);
  if(!found){
    showAlert("Header row not found in top 30 rows. Use the Standard Template.", "bad");
    return;
  }

  const v = validateHeaderRow(found.headerRowValues);
  if(!v.ok){
    showAlert(`Header validation failed. Missing: ${v.missing.join(", ")}. Use Standard Template.`, "bad");
    return;
  }

  const trimmed = rowsAOA.slice(found.headerRowIndex);
  const tempWs = XLSX.utils.aoa_to_sheet(trimmed);
  const json = XLSX.utils.sheet_to_json(tempWs, { defval:"" });

  rawRows = json.map(r=>{
    const obj={};
    for(const h of EXPECTED_HEADERS) obj[h]=normalizeValue(r[h]);
    obj.__dates={
      "Date of submission to CPC": parseAnyDate(obj["Date of submission to CPC"]),
      "acct_opn_date": parseAnyDate(obj["acct_opn_date"]),
      "last_any_tran_date": parseAnyDate(obj["last_any_tran_date"])
    };
    return obj;
  });

  saveDataset(rawRows);
  setDataChipText();
  showAlert(`Imported ${rawRows.length} rows successfully.`, "ok");
}

/* ---------------- Filters (dashboard) ---------------- */

function uniq(arr){ return [...new Set(arr)]; }

function fillSelect(selectEl, values){
  if(!selectEl) return;
  selectEl.innerHTML = "";
  const optAll = document.createElement("option");
  optAll.value=""; optAll.textContent="All";
  selectEl.appendChild(optAll);
  for(const v of values){
    const opt = document.createElement("option");
    opt.value=v; opt.textContent=v;
    selectEl.appendChild(opt);
  }
}

function autoSetDateRange(){
  const dates = rawRows.map(r=>r.__dates["Date of submission to CPC"]).filter(Boolean).sort((a,b)=>a-b);
  if(!dates.length) return;
  const max = dates[dates.length-1];
  const min = dates[0];
  const from = new Date(max.getFullYear(), max.getMonth(), max.getDate()-30);
  if(el("fromDate")) el("fromDate").value = fmtDateISO(from < min ? min : from);
  if(el("toDate")) el("toDate").value = fmtDateISO(max);
}

function populateFilters(){
  fillSelect(el("divisionFilter"), uniq(rawRows.map(r=>r["Division"]).filter(Boolean)).sort());
  fillSelect(el("officeFilter"), uniq(rawRows.map(r=>r["Office"]).filter(Boolean)).sort());
  fillSelect(el("statusFilter"), uniq(rawRows.map(r=>r["Status"]).filter(Boolean)).sort());
  fillSelect(el("scanFilter"), uniq(rawRows.map(r=>r["Scan/Upload status"]).filter(Boolean)).sort());

  const divisionFilter = el("divisionFilter");
  if(divisionFilter){
    divisionFilter.onchange = () => {
      const dvn = divisionFilter.value;
      const scoped = dvn ? rawRows.filter(r=>r["Division"]===dvn) : rawRows;
      fillSelect(el("officeFilter"), uniq(scoped.map(r=>r["Office"]).filter(Boolean)).sort());
    };
  }
}

function getFilters(){
  const dateBasis = el("dateBasis")?.value || "Date of submission to CPC";
  const from = el("fromDate")?.value ? new Date(el("fromDate").value+"T00:00:00") : null;
  const to = el("toDate")?.value ? new Date(el("toDate").value+"T00:00:00") : null;

  return {
    viewMode: el("viewMode")?.value || "review",
    dateBasis,
    from,to,
    division: el("divisionFilter")?.value || "",
    office: el("officeFilter")?.value || "",
    status: el("statusFilter")?.value || "",
    scan: el("scanFilter")?.value || ""
  };
}

function filterRows(f){
  return rawRows.filter(r=>{
    if(f.division && r["Division"]!==f.division) return false;
    if(f.office && r["Office"]!==f.office) return false;
    if(f.status && r["Status"]!==f.status) return false;
    if(f.scan && r["Scan/Upload status"]!==f.scan) return false;

    const d = r.__dates[f.dateBasis];
    if((f.from||f.to) && !inRange(d, f.from, f.to)) return false;

    return true;
  });
}

/* ---------------- KPIs + charts (same as before) ---------------- */

function countDuplicates(values){
  const m=new Map();
  for(const v of values) m.set(v,(m.get(v)||0)+1);
  let dup=0;
  for(const [,c] of m.entries()) if(c>1) dup += (c-1);
  return dup;
}

function ageingBuckets(rows){
  const now = new Date();
  const pend = rows.filter(r=>r.__dates["Date of submission to CPC"] && !isDoneScan(r["Scan/Upload status"]));
  const buckets = { "0–2 days":0, "3–7 days":0, "8–15 days":0, ">15 days":0 };
  for(const r of pend){
    const d = r.__dates["Date of submission to CPC"];
    const diff = Math.floor((startOfDay(now)-startOfDay(d)) / (1000*60*60*24));
    if(diff<=2) buckets["0–2 days"]++;
    else if(diff<=7) buckets["3–7 days"]++;
    else if(diff<=15) buckets["8–15 days"]++;
    else buckets[">15 days"]++;
  }
  return { pendCount: pend.length, buckets };
}

function topNCount(rows, keyFn, n=5){
  const m=new Map();
  for(const r of rows){
    const k = normalizeValue(keyFn(r));
    if(!k) continue;
    m.set(k,(m.get(k)||0)+1);
  }
  return [...m.entries()].map(([k,v])=>({k,v})).sort((a,b)=>b.v-a.v).slice(0,n);
}

function escapeHtml(s){
  return String(s).replace(/[&<>"']/g, (m)=>({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;" }[m]));
}

function renderKPIs(rows, f){
  const kpiGrid = el("kpiGrid");
  const meta = el("kpiMeta");
  if(!kpiGrid) return;

  const total = rows.length;
  const submitted = rows.filter(r=>r.__dates["Date of submission to CPC"]).length;
  const scanDone = rows.filter(r=>isDoneScan(r["Scan/Upload status"])).length;
  const scanPending = rows.filter(r=>r.__dates["Date of submission to CPC"] && !isDoneScan(r["Scan/Upload status"])).length;

  const missingConsign = rows.filter(r=>r.__dates["Date of submission to CPC"] && !hasText(r["Consignment number"])).length;
  const omissions = rows.filter(r=>hasText(r["Omissions/Rejections"])).length;
  const missingCif = rows.filter(r=>!hasText(r["cif_id"])).length;
  const missingName = rows.filter(r=>!hasText(r["acct_name"])).length;

  const dupAcct = countDuplicates(rows.map(r=>r["Account No"]).filter(Boolean));
  const dupCons = countDuplicates(rows.map(r=>r["Consignment number"]).filter(Boolean));

  const age = ageingBuckets(rows);
  const late15 = age.buckets[">15 days"] || 0;

  const topScheme = topNCount(rows, r=>r["schm_code"], 1);
  const topSchemeText = topScheme[0] ? `${topScheme[0].k} (${topScheme[0].v})` : "—";

  if(meta){
    const rangeText = (f.from||f.to) ? `${f.from?fmtDateISO(f.from):"…"} → ${f.to?fmtDateISO(f.to):"…"}` : "All dates";
    meta.textContent = `Rows: ${total} • View: ${f.viewMode.toUpperCase()} • Date basis: ${f.dateBasis} • Range: ${rangeText}`;
  }

  const submittedPct = total ? (submitted/total)*100 : 0;
  const scanDonePct = total ? (scanDone/total)*100 : 0;
  const omissionRate = total ? (omissions/total)*100 : 0;

  const kpis = [
    {label:"Total records", value: total, sub:"After filters"},
    {label:"Submitted to CPC", value: submitted, sub:`${submittedPct.toFixed(1)}% of total`},
    {label:"Scan/Upload Done", value: scanDone, sub:`${scanDonePct.toFixed(1)}% of total`},
    {label:"Pending Scan (submitted)", value: scanPending, sub:"Backlog KPI"},

    {label:"Missing Consignment (submitted)", value: missingConsign, sub:"Dispatch tracking gap"},
    {label:"Omissions/Rejections", value: omissions, sub:`Rate: ${omissionRate.toFixed(1)}%`},
    {label:"Pending > 15 days", value: late15, sub:"SLA breach indicator"},
    {label:"Top Scheme Code", value: topSchemeText, sub:"Most frequent scheme"},

    {label:"Missing CIF", value: missingCif, sub:"Data quality"},
    {label:"Missing Account Name", value: missingName, sub:"Data quality"},
    {label:"Duplicate Account No", value: dupAcct, sub:"Possible duplicates"},
    {label:"Duplicate Consignment No", value: dupCons, sub:"Dup dispatch risk"},

    {label:"Ageing buckets (pending)", value: age.pendCount, sub:`0–2:${age.buckets["0–2 days"]} • 3–7:${age.buckets["3–7 days"]} • 8–15:${age.buckets["8–15 days"]} • >15:${age.buckets[">15 days"]}`},
  ];

  kpiGrid.innerHTML = "";
  for(const k of kpis){
    const d=document.createElement("div");
    d.className="kpi";
    d.innerHTML = `
      <div class="kpi__label">${escapeHtml(k.label)}</div>
      <div class="kpi__value">${escapeHtml(String(k.value))}</div>
      <div class="kpi__sub">${escapeHtml(k.sub||"")}</div>
    `;
    kpiGrid.appendChild(d);
  }
}

function countBy(arr, keyFn){
  const map=new Map();
  for(const x of arr){
    const k=keyFn(x);
    map.set(k,(map.get(k)||0)+1);
  }
  return map;
}
function groupBy(rows, keyFn){
  const m=new Map();
  for(const r of rows){
    const k = normalizeValue(keyFn(r)) || "(Blank)";
    if(!m.has(k)) m.set(k, []);
    m.get(k).push(r);
  }
  return m;
}

function buildOrUpdateChart(existing, canvasId, type, data){
  const c = document.getElementById(canvasId);
  if(!c || typeof Chart === "undefined") return null;

  if(existing){
    existing.data = data;
    existing.update();
    return existing;
  }

  return new Chart(c, {
    type,
    data,
    options:{
      responsive:true,
      plugins:{ legend:{ labels:{ color:"rgba(255,255,255,.85)" } } },
      scales: type==="doughnut" ? {} : {
        x:{ ticks:{ color:"rgba(255,255,255,.75)" }, grid:{ color:"rgba(255,255,255,.06)" } },
        y:{ ticks:{ color:"rgba(255,255,255,.75)" }, grid:{ color:"rgba(255,255,255,.06)" } }
      }
    }
  });
}

function renderCharts(rows, f){
  if(typeof Chart === "undefined") return;

  const basis = f.dateBasis;
  const dated = rows.map(r => r.__dates[basis] ? fmtDateISO(r.__dates[basis]) : null).filter(Boolean);
  const trendMap = countBy(dated, x=>x);
  const trendLabels = [...trendMap.keys()].sort();
  const trendValues = trendLabels.map(l=>trendMap.get(l));

  charts.trend = buildOrUpdateChart(charts.trend, "trendChart", "line", {
    labels: trendLabels,
    datasets:[{ label:`Count by day (${basis})`, data:trendValues, tension:0.25 }]
  });

  const byDiv = groupBy(rows, r=>r["Division"] || "(Blank)");
  const divArr = [];
  for(const [div, list] of byDiv.entries()){
    const sub = list.filter(r=>r.__dates["Date of submission to CPC"]).length;
    const pend = list.filter(r=>r.__dates["Date of submission to CPC"] && !isDoneScan(r["Scan/Upload status"])).length;
    const pct = sub ? (pend/sub)*100 : 0;
    divArr.push({ div, pct:+pct.toFixed(2) });
  }
  divArr.sort((a,b)=>b.pct-a.pct);
  const top12 = divArr.slice(0,12);

  charts.division = buildOrUpdateChart(charts.division, "divisionChart", "bar", {
    labels: top12.map(x=>x.div),
    datasets:[{ label:"Pending Scan % (top 12)", data: top12.map(x=>x.pct) }]
  });

  const done = rows.filter(r=>isDoneScan(r["Scan/Upload status"])).length;
  const pending = rows.filter(r=>hasText(r["Scan/Upload status"]) && !isDoneScan(r["Scan/Upload status"])).length;
  const blank = rows.filter(r=>!hasText(r["Scan/Upload status"])).length;

  charts.scan = buildOrUpdateChart(charts.scan, "scanChart", "doughnut", {
    labels:["Done","Pending","Blank"],
    datasets:[{ label:"Scan/Upload status", data:[done,pending,blank] }]
  });
}

/* ---------------- Tables + dual scrollbars ---------------- */

function buildTable(tableEl, columns, rows){
  const thead = tableEl.querySelector("thead");
  const tbody = tableEl.querySelector("tbody");
  thead.innerHTML=""; tbody.innerHTML="";

  const trh=document.createElement("tr");
  for(const c of columns){
    const th=document.createElement("th");
    th.textContent=c;
    trh.appendChild(th);
  }
  thead.appendChild(trh);

  rows.forEach((r, idx)=>{
    const tr=document.createElement("tr");
    tr.dataset.rowIndex = String(idx);
    for(const c of columns){
      const td=document.createElement("td");
      td.textContent = normalizeValue(r[c]);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });
}

function syncTopBottomScroll(topScrollEl, topInnerEl, wrapEl){
  if(!topScrollEl || !topInnerEl || !wrapEl) return;

  const updateWidth = () => {
    topInnerEl.style.width = wrapEl.scrollWidth + "px";
  };
  updateWidth();
  window.addEventListener("resize", updateWidth);

  let lock=false;
  topScrollEl.addEventListener("scroll", ()=>{
    if(lock) return;
    lock=true;
    wrapEl.scrollLeft = topScrollEl.scrollLeft;
    lock=false;
  });
  wrapEl.addEventListener("scroll", ()=>{
    if(lock) return;
    lock=true;
    topScrollEl.scrollLeft = wrapEl.scrollLeft;
    lock=false;
  });
}

function filterTableBody(tableEl, q){
  const rows = [...tableEl.querySelectorAll("tbody tr")];
  rows.forEach(tr=>{
    const text = tr.innerText.toLowerCase();
    tr.style.display = text.includes(q) ? "" : "none";
  });
}

/* ---------------- Action items ---------------- */

function buildActionItems(rows){
  const out=[];
  for(const r of rows){
    const submitted = !!r.__dates["Date of submission to CPC"];
    const pendingScan = submitted && !isDoneScan(r["Scan/Upload status"]);
    const missingCons = submitted && !hasText(r["Consignment number"]);
    const hasOmission = hasText(r["Omissions/Rejections"]);
    const missingCif = !hasText(r["cif_id"]);
    const missingName = !hasText(r["acct_name"]);

    if(pendingScan || missingCons || hasOmission || missingCif || missingName){
      out.push({
        "Division": r["Division"],
        "Office": r["Office"],
        "sol_id": r["sol_id"],
        "Account No": r["Account No"],
        "Date of submission to CPC": r["Date of submission to CPC"],
        "Scan/Upload status": r["Scan/Upload status"],
        "Consignment number": r["Consignment number"],
        "Omissions/Rejections": r["Omissions/Rejections"],
        "Flags": [
          pendingScan ? "Pending Scan" : null,
          missingCons ? "Missing Consignment" : null,
          hasOmission ? "Omission/Rejection" : null,
          missingCif ? "Missing CIF" : null,
          missingName ? "Missing Name" : null
        ].filter(Boolean).join(" | ")
      });
    }
  }
  return out;
}

/* ---------------- Export ---------------- */

function csvCell(v){
  const s = String(v ?? "").replace(/"/g,'""');
  return /[",\n]/.test(s) ? `"${s}"` : s;
}

function downloadCsv(filename, cols, rows){
  if(!rows.length){
    showAlert("Nothing to export (0 rows).","warn");
    return;
  }
  const csv = [
    cols.join(","),
    ...rows.map(r => cols.map(c=>csvCell(r[c])).join(","))
  ].join("\n");

  const blob = new Blob([csv], {type:"text/csv;charset=utf-8"});
  const url = URL.createObjectURL(blob);
  const a=document.createElement("a");
  a.href=url;
  a.download=filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* ---------------- Page initializers ---------------- */

function commonWire(){
  applyThemeFromStorage();
  const tgl = el("themeToggle");
  if(tgl) tgl.addEventListener("click", toggleTheme);

  rawRows = loadDataset();
  setDataChipText();
}

function initUploadPage(){
  commonWire();

  el("btnDownloadTemplate")?.addEventListener("click", downloadStandardTemplate);
  el("btnGoDashboard")?.addEventListener("click", ()=> window.location.href="dashboard.html");

  el("fileInput")?.addEventListener("change", async (e)=>{
    const file = e.target.files?.[0];
    if(!file) return;
    try{
      await handleFile(file);
      setDataChipText();
    }catch(err){
      console.error(err);
      showAlert("Error reading file. Please ensure it is a valid Excel file.","bad");
    }
  });
}

function initDashboardPage(){
  commonWire();

  if(!rawRows.length){
    ensureNotBlankMessage();
    showAlert("No dataset loaded. Please go to Upload page and import first.", "warn");
    return;
  }

  populateFilters();
  autoSetDateRange();

  const apply = ()=>{
    const f = getFilters();
    filteredRows = filterRows(f);
    renderKPIs(filteredRows, f);
    renderCharts(filteredRows, f);
  };

  el("btnApply")?.addEventListener("click", apply);

  el("btnReset")?.addEventListener("click", ()=>{
    resetDataset();
    rawRows = [];
    filteredRows = [];
    setDataChipText();
    ensureNotBlankMessage();
    showAlert("Dataset reset. Please go to Upload page to import again.", "ok");
  });

  apply();
}

function initActionsPage(){
  commonWire();

  if(!rawRows.length){
    ensureNotBlankMessage();
    showAlert("No dataset loaded. Please go to Upload page and import first.", "warn");
    return;
  }

  actionRows = buildActionItems(rawRows);
  const table = el("actionsTable");
  if(table) buildTable(table, ACTION_COLS, actionRows);

  if(el("actionsSummary")) el("actionsSummary").textContent = `Action items: ${actionRows.length}`;

  el("actionsSearch")?.addEventListener("input", ()=>{
    const q = el("actionsSearch").value.toLowerCase().trim();
    filterTableBody(table, q);
  });

  el("btnDownloadActionsCsv")?.addEventListener("click", ()=>{
    downloadCsv(`kyc_action_items_${new Date().toISOString().slice(0,10)}.csv`, ACTION_COLS, actionRows);
  });

  syncTopBottomScroll(el("actionsTopScroll"), el("actionsTopInner"), el("actionsWrap"));
}

function initDataPage(){
  // Data page initialization remains from your previous version
  commonWire();

  if(!rawRows.length){
    ensureNotBlankMessage();
    showAlert("No dataset loaded. Please go to Upload page and import first.", "warn");
    return;
  }

  // If you already have the data-entry modal code in your current app.js,
  // keep it as-is. This hardened file focuses on ensuring pages never appear blank.
  // The table & scrollbars still work on data page if dataset exists.

  const table = el("dataTable");
  if(table){
    buildTable(table, EXPECTED_HEADERS, rawRows);
    if(el("dataSummary")) el("dataSummary").textContent = `Rows: ${rawRows.length}`;
    syncTopBottomScroll(el("dataTopScroll"), el("dataTopInner"), el("dataWrap"));

    el("dataSearch")?.addEventListener("input", ()=>{
      const q = el("dataSearch").value.toLowerCase().trim();
      filterTableBody(table, q);
    });

    el("btnDownloadCsv")?.addEventListener("click", ()=>{
      downloadCsv(`kyc_data_${new Date().toISOString().slice(0,10)}.csv`, EXPECTED_HEADERS, rawRows);
    });
  } else {
    ensureNotBlankMessage();
    showAlert("Data table container not found on this page. Please ensure data.html is updated.", "bad");
  }
}

/* ---------------- Boot ---------------- */

(function boot(){
  const p = pageName();

  // If a page forgot data-page, show a warning instead of staying blank
  if(!p){
    commonWire();
    ensureNotBlankMessage();
    showAlert("Page identifier missing (data-page). Please use the latest HTML files provided.", "bad");
    return;
  }

  if(p === "upload") initUploadPage();
  if(p === "dashboard") initDashboardPage();
  if(p === "actions") initActionsPage();
  if(p === "data") initDataPage();
})();
