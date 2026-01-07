/* NKR-KA KYC Tool — Browser-only static app (GitHub Pages ready)
   Includes:
   1) Download Standard Template (prevents header mismatch)
   2) Auto-detect header row (handles title rows / merged headings)
*/

const EXPECTED_HEADERS = [
  "Rgn Sl No",
  "Dvn Sl No",
  "sol_id",
  "Office",
  "Division",
  "Account No",
  "cif_id",
  "acct_name",
  "schm_code",
  "acct_opn_date",
  "last_any_tran_date",
  "Status",
  "Consignment number",
  "Date of submission to CPC",
  "Scan/Upload status",
  "Omissions/Rejections"
];

const el = (id) => document.getElementById(id);

let rawRows = [];
let filteredRows = [];
let charts = { trend: null, division: null, scan: null };

// Safety check: if the XLSX library didn't load, file upload will fail.
if (typeof XLSX === "undefined") {
  console.error("XLSX library not loaded. Check script tag / CDN / network policy.");
}

/* ---------- UI helpers ---------- */

function showAlert(msg, type = "ok") {
  const box = el("alertBox");
  box.classList.remove("hidden", "alert--ok", "alert--warn", "alert--bad");
  box.classList.add(`alert--${type}`);
  box.textContent = msg;
}

function clearAlert() {
  const box = el("alertBox");
  box.classList.add("hidden");
  box.textContent = "";
}

function setDataChip(text) {
  el("dataChip").textContent = text;
}

function setVisible(id, visible) {
  const node = el(id);
  if (!node) return;
  node.classList.toggle("hidden", !visible);
}

function normalizeValue(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v.trim();
  return String(v).trim();
}

/* ---------- Header normalization + auto-detect ---------- */

function normHeader(s) {
  return String(s ?? "")
    .replace(/\u00A0/g, " ")     // NBSP -> space
    .replace(/\r?\n/g, " ")      // newlines -> space
    .replace(/\s+/g, " ")        // collapse spaces
    .trim()
    .toLowerCase();
}

function countHeaderMatches(rowValues) {
  const set = new Set(rowValues.map(normHeader).filter(Boolean));
  let matched = 0;
  for (const h of EXPECTED_HEADERS) {
    if (set.has(normHeader(h))) matched++;
  }
  return matched;
}

/**
 * Finds the header row index by scanning top N rows
 * Returns {headerRowIndex, headerRowValues} or null
 */
function findHeaderRowIndex(rowsAOA, scanRows = 30) {
  let bestIdx = -1;
  let bestScore = -1;

  const max = Math.min(rowsAOA.length, scanRows);
  for (let i = 0; i < max; i++) {
    const row = (rowsAOA[i] || []).map(v => normalizeValue(v)).filter(v => v !== "");
    if (!row.length) continue;

    const score = countHeaderMatches(row);
    if (score > bestScore) {
      bestScore = score;
      bestIdx = i;
    }
  }

  // Require a minimum match threshold (prevents accidental matches)
  // We expect all 16, but accept >= 10 to allow cases where some headers are blank/merged.
  if (bestIdx >= 0 && bestScore >= 10) {
    return {
      headerRowIndex: bestIdx,
      headerRowValues: (rowsAOA[bestIdx] || []).map(v => normalizeValue(v))
    };
  }
  return null;
}

function validateHeaderRow(actualHeaderRow) {
  const detectedSet = new Set(actualHeaderRow.map(normHeader).filter(Boolean));
  const missing = EXPECTED_HEADERS.filter(h => !detectedSet.has(normHeader(h)));
  return { ok: missing.length === 0, missing };
}

/* ---------- Date parsing ---------- */

function parseAnyDate(v) {
  if (v === null || v === undefined || v === "") return null;

  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    return new Date(d.y, d.m - 1, d.d);
  }

  const s = String(v).trim();
  if (!s) return null;

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const d = new Date(s + "T00:00:00");
    return isNaN(d.getTime()) ? null : d;
  }

  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    let dd = parseInt(m[1], 10);
    let mm = parseInt(m[2], 10);
    let yy = parseInt(m[3], 10);
    if (yy < 100) yy += 2000;
    const d = new Date(yy, mm - 1, dd);
    return isNaN(d.getTime()) ? null : d;
  }

  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function fmtDateISO(d) {
  if (!d) return "";
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function startOfDay(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function inRange(d, from, to) {
  if (!d) return false;
  const t = startOfDay(d).getTime();
  if (from && t < startOfDay(from).getTime()) return false;
  if (to && t > startOfDay(to).getTime()) return false;
  return true;
}

function toLower(s) {
  return String(s || "").trim().toLowerCase();
}

function isDoneScan(v) {
  const s = toLower(v);
  return ["done", "completed", "complete", "uploaded", "scanned", "ok", "yes"].some(k => s.includes(k));
}

function hasText(v) {
  return String(v || "").trim().length > 0;
}

function uniq(arr) {
  return [...new Set(arr)];
}

function countBy(rows, keyFn) {
  const map = new Map();
  for (const r of rows) {
    const k = keyFn(r);
    map.set(k, (map.get(k) || 0) + 1);
  }
  return map;
}

function groupBy(rows, keyFn) {
  const m = new Map();
  for (const r of rows) {
    const k = normalizeValue(keyFn(r)) || "(Blank)";
    if (!m.has(k)) m.set(k, []);
    m.get(k).push(r);
  }
  return m;
}

function topNCount(rows, keyFn, n = 5) {
  const m = new Map();
  for (const r of rows) {
    const k = normalizeValue(keyFn(r));
    if (!k) continue;
    m.set(k, (m.get(k) || 0) + 1);
  }
  return [...m.entries()].map(([k, v]) => ({ k, v })).sort((a, b) => b.v - a.v).slice(0, n);
}

function countDuplicates(values) {
  const m = new Map();
  for (const v of values) m.set(v, (m.get(v) || 0) + 1);
  let dup = 0;
  for (const [, c] of m.entries()) if (c > 1) dup += (c - 1);
  return dup;
}

/* ---------- Standard Template Download ---------- */

function downloadStandardTemplate() {
  if (typeof XLSX === "undefined") {
    showAlert("Cannot generate template because XLSX library did not load.", "bad");
    return;
  }

  // Row 1 = headers; Row 2 = blank sample row
  const aoa = [
    EXPECTED_HEADERS,
    EXPECTED_HEADERS.map(() => "")
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Make header bold + freeze top row (best-effort; styling support varies)
  // Column widths (optional)
  ws["!cols"] = EXPECTED_HEADERS.map(h => ({ wch: Math.max(12, Math.min(32, h.length + 2)) }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "KYC_Data");

  const filename = `NKR_KYC_Standard_Template.xlsx`;
  XLSX.writeFile(wb, filename);

  showAlert("Standard template downloaded. Fill it and upload for processing.", "ok");
}

/* ---------- Upload & Parse ---------- */

async function handleFile(file) {
  clearAlert();
  rawRows = [];
  filteredRows = [];

  if (typeof XLSX === "undefined") {
    showAlert(
      "Upload failed because the XLSX library did not load. Please check internet/CDN access or use the offline/local library option.",
      "bad"
    );
    return;
  }

  const arrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const firstSheetName = wb.SheetNames[0];
  const ws = wb.Sheets[firstSheetName];

  // Read sheet as AOA to locate header row anywhere in top section
  const rowsAOA = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });

  if (!rowsAOA.length) {
    showAlert("No rows found in the first sheet.", "bad");
    return;
  }

  const found = findHeaderRowIndex(rowsAOA, 30);
  if (!found) {
    showAlert(
      "Header row not found in the top 30 rows. Please use 'Download Standard Template' and paste your data into it.",
      "bad"
    );
    return;
  }

  const headerRowIndex = found.headerRowIndex;
  const headerRow = found.headerRowValues;

  // Validate headers strictly (after normalization)
  const v = validateHeaderRow(headerRow);
  if (!v.ok) {
    showAlert(
      `Header validation failed. Missing headers: ${v.missing.join(", ")}. ` +
      `Please use 'Download Standard Template' to avoid mismatch.`,
      "bad"
    );
    return;
  }

  // Convert to JSON using the detected header row as keys
  // We rebuild a "sub-sheet" that starts at the header row
  const trimmedAOA = rowsAOA.slice(headerRowIndex);
  const tempWs = XLSX.utils.aoa_to_sheet(trimmedAOA);
  const json = XLSX.utils.sheet_to_json(tempWs, { defval: "" });

  // Normalize into canonical keys exactly as EXPECTED_HEADERS
  rawRows = json.map((r) => {
    const obj = {};
    for (const h of EXPECTED_HEADERS) obj[h] = normalizeValue(r[h]);

    obj.__dates = {
      "Date of submission to CPC": parseAnyDate(obj["Date of submission to CPC"]),
      "acct_opn_date": parseAnyDate(obj["acct_opn_date"]),
      "last_any_tran_date": parseAnyDate(obj["last_any_tran_date"])
    };
    return obj;
  });

  if (!rawRows.length) {
    showAlert("No data rows found after the header row.", "bad");
    return;
  }

  populateFilters(rawRows);

  setDataChip(`Loaded: ${file.name} • Rows: ${rawRows.length}`);
  showAlert(
    `Template validated successfully. Loaded ${rawRows.length} rows. (Header row detected at line ${headerRowIndex + 1})`,
    "ok"
  );

  setVisible("filtersPanel", true);
  setVisible("dashPanel", true);

  autoSetDateRange();
  activateTab("kpis");
  applyFiltersAndRender();
}

function autoSetDateRange() {
  const dates = rawRows
    .map(r => r.__dates["Date of submission to CPC"])
    .filter(Boolean)
    .sort((a, b) => a - b);

  if (!dates.length) return;

  const max = dates[dates.length - 1];
  const min = dates[0];
  const from = new Date(max.getFullYear(), max.getMonth(), max.getDate() - 30);

  el("fromDate").value = fmtDateISO(from < min ? min : from);
  el("toDate").value = fmtDateISO(max);
}

function populateFilters(rows) {
  const divisions = uniq(rows.map(r => r["Division"]).filter(Boolean)).sort();
  const offices = uniq(rows.map(r => r["Office"]).filter(Boolean)).sort();
  const status = uniq(rows.map(r => r["Status"]).filter(Boolean)).sort();
  const scan = uniq(rows.map(r => r["Scan/Upload status"]).filter(Boolean)).sort();

  fillSelect(el("divisionFilter"), divisions);
  fillSelect(el("officeFilter"), offices);
  fillSelect(el("statusFilter"), status);
  fillSelect(el("scanFilter"), scan);

  el("divisionFilter").onchange = () => {
    const dvn = el("divisionFilter").value;
    const scoped = dvn ? rows.filter(r => r["Division"] === dvn) : rows;
    const newOffices = uniq(scoped.map(r => r["Office"]).filter(Boolean)).sort();
    fillSelect(el("officeFilter"), newOffices, true);
  };
}

function fillSelect(selectEl, values, keepAll = false) {
  selectEl.innerHTML = "";

  const optAll = document.createElement("option");
  optAll.value = "";
  optAll.textContent = "All";
  selectEl.appendChild(optAll);

  for (const v of values) {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    selectEl.appendChild(opt);
  }

  if (keepAll) selectEl.value = "";
}

/* ---------- Filtering & Rendering ---------- */

function getCurrentFilters() {
  const dateBasis = el("dateBasis").value;
  const from = el("fromDate").value ? new Date(el("fromDate").value + "T00:00:00") : null;
  const to = el("toDate").value ? new Date(el("toDate").value + "T00:00:00") : null;

  return {
    viewMode: el("viewMode").value,
    dateBasis,
    from,
    to,
    division: el("divisionFilter").value,
    office: el("officeFilter").value,
    status: el("statusFilter").value,
    scan: el("scanFilter").value
  };
}

function filterRows(rows, f) {
  return rows.filter(r => {
    if (f.division && r["Division"] !== f.division) return false;
    if (f.office && r["Office"] !== f.office) return false;
    if (f.status && r["Status"] !== f.status) return false;
    if (f.scan && r["Scan/Upload status"] !== f.scan) return false;

    const d = r.__dates[f.dateBasis];
    if ((f.from || f.to) && !inRange(d, f.from, f.to)) return false;

    return true;
  });
}

function applyFiltersAndRender() {
  if (!rawRows.length) return;

  const f = getCurrentFilters();
  filteredRows = filterRows(rawRows, f);

  renderKPIs(filteredRows, f);
  renderQuality(filteredRows);
  renderAgeing(filteredRows);
  renderDataTable(filteredRows);
  renderActionItems(filteredRows);
  renderCharts(filteredRows, f);

  el("dataSummary").textContent = `Showing ${filteredRows.length} rows in Data table`;
}

/* ---------- KPI / Cards ---------- */

function renderKPIs(rows, f) {
  const total = rows.length;

  const submitted = rows.filter(r => r.__dates["Date of submission to CPC"]).length;
  const missingConsignment = rows.filter(r => !hasText(r["Consignment number"]) && r.__dates["Date of submission to CPC"]).length;

  const doneScan = rows.filter(r => isDoneScan(r["Scan/Upload status"])).length;
  const pendingScan = rows.filter(r => !isDoneScan(r["Scan/Upload status"]) && r.__dates["Date of submission to CPC"]).length;

  const omissions = rows.filter(r => hasText(r["Omissions/Rejections"])).length;
  const omissionRate = total ? (omissions / total) * 100 : 0;

  const uniqSol = uniq(rows.map(r => r["sol_id"]).filter(Boolean)).length;
  const uniqAcct = uniq(rows.map(r => r["Account No"]).filter(Boolean)).length;

  const missingCif = rows.filter(r => !hasText(r["cif_id"])).length;
  const missingName = rows.filter(r => !hasText(r["acct_name"])).length;

  const schemeTop = topNCount(rows, r => r["schm_code"], 5);
  const modeLabel = f.viewMode.toUpperCase();

  const kpis = [
    { label: `${modeLabel} • Total Rows`, value: total, sub: "After filters" },
    { label: "Submitted (has submission date)", value: submitted, sub: "Based on ‘Date of submission to CPC’" },
    { label: "Scan/Upload Done", value: doneScan, sub: "Auto-detected synonyms: done/completed/uploaded..." },
    { label: "Pending Scan/Upload", value: pendingScan, sub: "Submitted but not ‘Done’" },

    { label: "Missing Consignment", value: missingConsignment, sub: "Among submitted" },
    { label: "Omissions/Rejections", value: omissions, sub: `Rate: ${omissionRate.toFixed(1)}%` },
    { label: "Unique SOL IDs", value: uniqSol, sub: "Coverage KPI" },
    { label: "Unique Account Nos", value: uniqAcct, sub: "De-dup KPI" },

    { label: "Missing CIF", value: missingCif, sub: "Data quality" },
    { label: "Missing Account Name", value: missingName, sub: "Data quality" },
    { label: "Top Schemes", value: schemeTop[0] ? schemeTop[0].k : "—", sub: schemeTop.map(x => `${x.k}: ${x.v}`).join(" • ") || "No scheme codes found" },
    { label: "Division Count", value: uniq(rows.map(r => r["Division"]).filter(Boolean)).length, sub: "Divisions in filtered set" }
  ];

  const grid = el("kpiGrid");
  grid.innerHTML = "";
  for (const k of kpis) {
    const d = document.createElement("div");
    d.className = "kpi";
    d.innerHTML = `
      <div class="kpi__label">${escapeHtml(k.label)}</div>
      <div class="kpi__value">${escapeHtml(String(k.value))}</div>
      <div class="kpi__sub">${escapeHtml(k.sub || "")}</div>
    `;
    grid.appendChild(d);
  }
}

function renderQuality(rows) {
  const invalidSubmission = rows.filter(r => r["Date of submission to CPC"] && !r.__dates["Date of submission to CPC"]).length;
  const invalidOpen = rows.filter(r => r["acct_opn_date"] && !r.__dates["acct_opn_date"]).length;
  const invalidLast = rows.filter(r => r["last_any_tran_date"] && !r.__dates["last_any_tran_date"]).length;

  const dupAccounts = countDuplicates(rows.map(r => r["Account No"]).filter(Boolean));
  const dupConsign = countDuplicates(rows.map(r => r["Consignment number"]).filter(Boolean));

  const missingAny = rows.filter(r => {
    const must = ["sol_id", "Office", "Division", "Account No"];
    return must.some(k => !hasText(r[k]));
  }).length;

  el("qualityBox").innerHTML = `
    <ul>
      <li>Invalid dates — Submission: <b>${invalidSubmission}</b>, Open: <b>${invalidOpen}</b>, Last Tran: <b>${invalidLast}</b></li>
      <li>Duplicate Account No occurrences: <b>${dupAccounts}</b></li>
      <li>Duplicate Consignment No occurrences: <b>${dupConsign}</b></li>
      <li>Rows missing one or more core identifiers (sol_id/Office/Division/Account No): <b>${missingAny}</b></li>
    </ul>
  `;
}

function renderAgeing(rows) {
  const now = new Date();
  const pend = rows.filter(r => r.__dates["Date of submission to CPC"] && !isDoneScan(r["Scan/Upload status"]));

  const buckets = { "0–2 days": 0, "3–7 days": 0, "8–15 days": 0, ">15 days": 0 };
  for (const r of pend) {
    const d = r.__dates["Date of submission to CPC"];
    const diffDays = Math.floor((startOfDay(now) - startOfDay(d)) / (1000 * 60 * 60 * 24));
    if (diffDays <= 2) buckets["0–2 days"]++;
    else if (diffDays <= 7) buckets["3–7 days"]++;
    else if (diffDays <= 15) buckets["8–15 days"]++;
    else buckets[">15 days"]++;
  }

  el("ageingBox").innerHTML = `
    <div>Pending scan cases: <b>${pend.length}</b></div>
    <ul>
      ${Object.entries(buckets).map(([k,v]) => `<li>${k}: <b>${v}</b></li>`).join("")}
    </ul>
  `;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (m) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;"
  }[m]));
}

/* ---------- Tables ---------- */

function buildTable(tableEl, columns, rows) {
  const thead = tableEl.querySelector("thead");
  const tbody = tableEl.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  const trh = document.createElement("tr");
  for (const c of columns) {
    const th = document.createElement("th");
    th.textContent = c;
    trh.appendChild(th);
  }
  thead.appendChild(trh);

  for (const r of rows) {
    const tr = document.createElement("tr");
    for (const c of columns) {
      const td = document.createElement("td");
      td.textContent = normalizeValue(r[c]);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  const dataSearch = el("dataSearch");
  if (tableEl.id === "dataTable") {
    dataSearch.oninput = () => {
      const q = dataSearch.value.toLowerCase().trim();
      filterTableBody(tableEl, q);
    };
  }
}

function filterTableBody(tableEl, q) {
  const tbody = tableEl.querySelector("tbody");
  const rows = [...tbody.querySelectorAll("tr")];
  for (const tr of rows) {
    const text = tr.innerText.toLowerCase();
    tr.style.display = text.includes(q) ? "" : "none";
  }
}

function renderDataTable(rows) {
  buildTable(el("dataTable"), EXPECTED_HEADERS, rows);
  el("dataSummary").textContent = `Showing ${rows.length} rows in Data table`;
}

function renderActionItems(rows) {
  const actions = [];

  for (const r of rows) {
    const submitted = !!r.__dates["Date of submission to CPC"];
    const pendingScan = submitted && !isDoneScan(r["Scan/Upload status"]);
    const missingConsignment = submitted && !hasText(r["Consignment number"]);
    const hasOmission = hasText(r["Omissions/Rejections"]);
    const missingCif = !hasText(r["cif_id"]);
    const missingName = !hasText(r["acct_name"]);

    if (pendingScan || missingConsignment || hasOmission || missingCif || missingName) {
      actions.push({
        "Division": r["Division"],
        "Office": r["Office"],
        "sol_id": r["sol_id"],
        "Account No": r["Account No"],
        "Submission Date": r["Date of submission to CPC"],
        "Scan/Upload status": r["Scan/Upload status"],
        "Consignment number": r["Consignment number"],
        "Omissions/Rejections": r["Omissions/Rejections"],
        "Flags": [
          pendingScan ? "Pending Scan" : null,
          missingConsignment ? "Missing Consignment" : null,
          hasOmission ? "Omission/Rejection" : null,
          missingCif ? "Missing CIF" : null,
          missingName ? "Missing Name" : null
        ].filter(Boolean).join(" | ")
      });
    }
  }

  const cols = ["Division","Office","sol_id","Account No","Submission Date","Scan/Upload status","Consignment number","Omissions/Rejections","Flags"];
  buildTable(el("actionsTable"), cols, actions);

  el("actionsSummary").textContent = `Action items: ${actions.length} (Pending scan / missing consignment / omissions / missing CIF/name)`;
  el("actionsSearch").oninput = () => {
    const q = el("actionsSearch").value.toLowerCase().trim();
    filterTableBody(el("actionsTable"), q);
  };
}

/* ---------- Charts ---------- */

function buildOrUpdateChart(existing, canvasId, type, data) {
  const ctx = el(canvasId);
  if (!ctx) return null;

  if (existing) {
    existing.data = data;
    existing.update();
    return existing;
  }

  return new Chart(ctx, {
    type,
    data,
    options: {
      responsive: true,
      plugins: { legend: { labels: { color: "rgba(255,255,255,0.85)" } } },
      scales: type === "doughnut" ? {} : {
        x: { ticks: { color: "rgba(255,255,255,0.75)" }, grid: { color: "rgba(255,255,255,0.06)" } },
        y: { ticks: { color: "rgba(255,255,255,0.75)" }, grid: { color: "rgba(255,255,255,0.06)" } }
      }
    }
  });
}

function renderCharts(rows, f) {
  const basis = f.dateBasis;
  const dated = rows
    .map(r => r.__dates[basis] ? { d: fmtDateISO(r.__dates[basis]) } : null)
    .filter(Boolean);

  const trendMap = countBy(dated, x => x.d);
  const trendLabels = [...trendMap.keys()].sort();
  const trendValues = trendLabels.map(l => trendMap.get(l));

  const byDiv = groupBy(rows, r => r["Division"] || "(Blank)");
  const divLabels = [];
  const divVals = [];
  for (const [div, list] of byDiv.entries()) {
    const submitted = list.filter(r => r.__dates["Date of submission to CPC"]).length || 0;
    const pend = list.filter(r => r.__dates["Date of submission to CPC"] && !isDoneScan(r["Scan/Upload status"])).length || 0;
    const pct = submitted ? (pend / submitted) * 100 : 0;
    divLabels.push(div);
    divVals.push(+pct.toFixed(2));
  }

  const zipped = divLabels.map((d,i) => ({d, v: divVals[i]})).sort((a,b) => b.v - a.v).slice(0, 12);
  const divLabels2 = zipped.map(x => x.d);
  const divVals2 = zipped.map(x => x.v);

  const done = rows.filter(r => isDoneScan(r["Scan/Upload status"])).length;
  const pending = rows.filter(r => hasText(r["Scan/Upload status"]) && !isDoneScan(r["Scan/Upload status"])).length;
  const blank = rows.filter(r => !hasText(r["Scan/Upload status"])).length;

  charts.trend = buildOrUpdateChart(charts.trend, "trendChart", "line", {
    labels: trendLabels,
    datasets: [{ label: `Count by day (${basis})`, data: trendValues, tension: 0.25 }]
  });

  charts.division = buildOrUpdateChart(charts.division, "divisionChart", "bar", {
    labels: divLabels2,
    datasets: [{ label: "Pending Scan % (top 12)", data: divVals2 }]
  });

  charts.scan = buildOrUpdateChart(charts.scan, "scanChart", "doughnut", {
    labels: ["Done", "Pending", "Blank"],
    datasets: [{ label: "Scan/Upload status", data: [done, pending, blank] }]
  });
}

/* ---------- Tabs ---------- */

function activateTab(name) {
  const tabs = document.querySelectorAll(".tab");
  const panes = {
    kpis: el("tab-kpis"),
    charts: el("tab-charts"),
    actions: el("tab-actions"),
    data: el("tab-data")
  };
  tabs.forEach(t => t.classList.toggle("active", t.dataset.tab === name));
  Object.entries(panes).forEach(([k, node]) => node.classList.toggle("hidden", k !== name));
}

/* ---------- Export ---------- */

function csvCell(v) {
  const s = String(v ?? "").replace(/"/g, '""');
  if (/[",\n]/.test(s)) return `"${s}"`;
  return s;
}

function downloadCsv(rows) {
  if (!rows.length) {
    showAlert("Nothing to export (0 rows after filters).", "warn");
    return;
  }

  const cols = EXPECTED_HEADERS;
  const csv = [
    cols.join(","),
    ...rows.map(r => cols.map(c => csvCell(r[c])).join(","))
  ].join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = `kyc_export_${new Date().toISOString().slice(0,10)}.csv`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* ---------- Events ---------- */

function wireEvents() {
  el("fileInput").addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      await handleFile(file);
    } catch (err) {
      console.error(err);

      if (typeof XLSX === "undefined") {
        showAlert(
          "Upload failed because the XLSX library did not load. Please check internet/CDN access or use the offline/local library option.",
          "bad"
        );
        return;
      }

      showAlert("Error reading file. Please ensure it is a valid Excel file.", "bad");
    }
  });

  el("btnDownloadTemplate").addEventListener("click", () => downloadStandardTemplate());

  el("btnApply").addEventListener("click", () => applyFiltersAndRender());

  el("btnReset").addEventListener("click", () => {
    rawRows = [];
    filteredRows = [];
    clearAlert();
    setDataChip("No data loaded");
    el("fileInput").value = "";
    setVisible("filtersPanel", false);
    setVisible("dashPanel", false);

    if (charts.trend) charts.trend.destroy();
    if (charts.division) charts.division.destroy();
    if (charts.scan) charts.scan.destroy();
    charts = { trend: null, division: null, scan: null };

    showAlert("Reset done. Please upload the template again.", "ok");
  });

  el("btnDownloadCsv").addEventListener("click", () => downloadCsv(filteredRows));

  el("btnPrintReview").addEventListener("click", () => {
    activateTab("kpis");
    window.print();
  });

  el("viewMode").addEventListener("change", () => applyFiltersAndRender());

  document.querySelectorAll(".tab").forEach(btn => {
    btn.addEventListener("click", () => activateTab(btn.dataset.tab));
  });
}

wireEvents();
document.querySelector('.tab[data-tab="kpis"]').classList.add("active");
