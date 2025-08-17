// ---------------- i18n ----------------
const I18N = {
  en: {
    "app.title": "Duplicates Auditor",
    "header.title": "Duplicates Auditor",
    "header.subtitle": "Unpaid & Paid CSV compare",
    "steps.title": "Getting started",
    "steps.1": "Upload <strong>Unpaid</strong> CSV",
    "steps.2": "Upload <strong>Paid</strong> CSV",
    "steps.3": "Adjust settings (optional)",
    "steps.4": "Run detection",
    "settings.dateWindow": "Date window (days)",
    "settings.amountTol": "Amount tolerance ($)",
    "settings.advanced": "Advanced options",
    "settings.vendorNorm": "Normalize vendor names (ignore punctuation, LLC/Inc/Co/Corp, “THE”; “&” → “AND”)",
    "upload.unpaid": "Unpaid CSV",
    "upload.paid": "Paid CSV",
    "upload.drop": "Drop CSV here or click to select",
    "upload.hintsUnpaid": "Required: EntryDate • VendorName • Amount • PropertyName",
    "upload.hintsPaid": "Required: entryDate • payeeNameRaw • amount • buildingName",
    "upload.showMap": "Show mapping",
    "results.title": "Results",
    "results.unpaid": "Duplicates: Unpaid",
    "results.paid": "Duplicates: Paid",
    "results.cross": "Paid vs Unpaid",
    "rows.compact": "Compact",
    "rows.all": "All rows",
    "cols.basic": "Basic",
    "cols.all": "All columns",
    "filter.placeholder": "Filter vendor/building…",
    "actions.run": "Run detection",
    "actions.export": "Download CSV",
    "progress.preparing": "Preparing...",
    "progress.processing": "Processing...",
    "progress.matching": "Matching duplicates...",
    "progress.crossMatch": "Cross-matching paid vs unpaid...",
    "progress.updating": "Updating results...",
    "progress.complete": "Complete!",
    "pager.prev": "‹ Prev",
    "pager.next": "Next ›",
    "pager.showing": "Showing",
    "pager.of": "of",
    "pager.groups": "groups",
    "footer.brand": "Duplicates Auditor"
  },
  zh: {
    "app.title": "重复检测工具",
    "header.title": "重复检测工具",
    "header.subtitle": "对比未付与已付 CSV",
    "steps.title": "快速开始",
    "steps.1": "上传 <strong>未付</strong> CSV",
    "steps.2": "上传 <strong>已付</strong> CSV",
    "steps.3": "调整设置（可选）",
    "steps.4": "开始检测",
    "settings.dateWindow": "日期窗口（天）",
    "settings.amountTol": "金额容差（$）",
    "settings.advanced": "高级选项",
    "settings.vendorNorm": "归一化供应商名称（忽略标点、LLC/Inc/Co/Corp、“THE”；“&”→“AND”）",
    "upload.unpaid": "未付 CSV",
    "upload.paid": "已付 CSV",
    "upload.drop": "将 CSV 拖入此处或点击选择",
    "upload.hintsUnpaid": "必需列：EntryDate • VendorName • Amount • PropertyName",
    "upload.hintsPaid": "必需列：entryDate • payeeNameRaw • amount • buildingName",
    "upload.showMap": "显示映射",
    "results.title": "结果",
    "results.unpaid": "重复项：未付",
    "results.paid": "重复项：已付",
    "results.cross": "已付 vs 未付",
    "rows.compact": "紧凑",
    "rows.all": "所有行",
    "cols.basic": "基础",
    "cols.all": "全部列",
    "filter.placeholder": "筛选 供应商/楼宇…",
    "actions.run": "开始检测",
    "actions.export": "下载 CSV",
    "progress.preparing": "准备中...",
    "progress.processing": "处理中...",
    "progress.matching": "正在匹配重复项...",
    "progress.crossMatch": "正在交叉匹配已付与未付...",
    "progress.updating": "正在更新结果...",
    "progress.complete": "完成！",
    "pager.prev": "‹ 上一页",
    "pager.next": "下一页 ›",
    "pager.showing": "显示",
    "pager.of": "共",
    "pager.groups": "组",
    "footer.brand": "重复检测工具"
  }
};
let currentLang = "en";
function t(key){ const dict = I18N[currentLang] || I18N.en; return dict[key] || I18N.en[key] || key; }
function applyI18n(){
  // Text nodes
  document.querySelectorAll("[data-i18n]").forEach(el=>{
    const key = el.getAttribute("data-i18n");
    const htmlSafe = ["steps.1","steps.2","steps.3","steps.4"].includes(key);
    if(htmlSafe) el.innerHTML = t(key);
    else el.textContent = t(key);
  });
  // Placeholder attributes
  document.querySelectorAll("[data-i18n-placeholder]").forEach(el=>{
    const key = el.getAttribute("data-i18n-placeholder");
    el.setAttribute("placeholder", t(key));
  });
  // Document title and language toggle label
  document.title = t("app.title");
  const langLabel = document.getElementById("lang-label");
  if(langLabel) langLabel.textContent = currentLang === "en" ? "中文" : "EN";
}

// ---------------- Tabs/sections not needed; single tool app ----------------

// ---------------- Column contracts ----------------
const CONTRACTS = {
  unpaid: {
    name: "Unpaid",
    required: {
      building: ["PropertyName","Building","buildingName","Property","Property Name","Building Name"],
      vendor:   ["VendorName","Vendor","Payee","PayeeName","payee","payeeName"],
      date:     ["EntryDate","DueDate","InvoiceDate","Date","TxnDate","Transaction Date"],
      amount:   ["Amount","AmountDue","Balance","Total","OpenBalance"],
    },
    optional: {
      reference:["RefNo","InvoiceNo","Invoice #","Ref #","referenceNumber","reference","CheckNo","checkNo"],
      memo:     ["Memo","Description","Notes","Note","Details"]
    }
  },
  paid: {
    name: "Paid",
    required: {
      building: ["buildingName","PropertyName","Building","Property","Property Name","Building Name"],
      vendor:   ["payeeNameRaw","VendorName","Vendor","Payee","PayeeName","payee","payeeName","vendorNameRaw"],
      date:     ["entryDate","PostedDate","ClearedDate","PaymentDate","Date","TxnDate","Transaction Date","postedOn"],
      amount:   ["amount","Amount","AmountOut","Amount Out","Total"],
    },
    optional: {
      reference:["CheckNo","RefNo","Reference #","Check #","ChequeNo","referenceNumber","reference"],
      memo:     ["memo","Memo","Description","Notes","Note","Details"]
    }
  }
};

// ---------------- Helpers ----------------
const el = sel => document.querySelector(sel);
const els = sel => Array.from(document.querySelectorAll(sel));
function normalizeName(name, enabled = true){
  if(name==null) return "";
  let v = String(name);
  if(!enabled) return v.trim().toUpperCase();
  v = v.toUpperCase().replace(/&/g,"AND").replace(/[^A-Z0-9 ]+/g," ").replace(/\b(LLC|INC|CO|CORP|THE)\b/g,"").replace(/\s+/g," ").trim();
  return v;
}
function normalizeBuilding(b){ return (b==null?"":String(b)).trim().toUpperCase(); }
function parseMoney(a){
  if(a==null || a==="") return null;
  const s = String(a).replace(/[\$,]/g,"").trim();
  const n = Number(s);
  return Number.isFinite(n) ? Math.abs(n) : null;
}
function parseDate(d){
  if(d==null) return null;
  const s = String(d).trim();
  if(!s) return null;
  const patterns = [
    /^(\d{4})-(\d{1,2})-(\d{1,2})$/,
    /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/,
    /^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/,
    /^(\d{1,2})-(\d{1,2})-(\d{2,4})$/,
    /^(\d{4})(\d{2})(\d{2})$/
  ];
  let y,m,da;
  for(const re of patterns){
    const mth = s.match(re);
    if(!mth) continue;
    if(re===patterns[0]) { y=+mth[1]; m=+mth[2]; da=+mth[3]; }
    else if(re===patterns[4]){ y=+mth[1]; m=+mth[2]; da=+mth[3]; }
    else { const a=+mth[1], b=+mth[2], c=+mth[3]; y=(mth[3].length===2)?(2000+c):c; m=a; da=b; }
    const dt = new Date(y, m-1, da);
    if(!isNaN(dt)) return dt;
  }
  const dt = new Date(s);
  return isNaN(dt) ? null : dt;
}
function dateDiffDays(a,b){
  if(!a || !b) return Infinity;
  const ms = Math.abs(a.getTime() - b.getTime());
  return Math.floor(ms/86400000);
}

// ---------------- CSV mapping ----------------
function pickColumn(headers, candidates){
  if(!headers || !headers.length) return null;
  const lower = headers.map(h=>h.toLowerCase());
  for(const cand of candidates){
    const idx = lower.indexOf(cand.toLowerCase());
    if(idx !== -1) return headers[idx];
  }
  const norm = s => s.toLowerCase().replace(/[\s_\-/#.]+/g,"").replace(/[()]/g,"");
  const nheaders = headers.map(norm);
  for(const cand of candidates.map(norm)){
    const idx = nheaders.indexOf(cand);
    if(idx !== -1) return headers[idx];
  }
  return null;
}
function mapRecords(rows, headers, contract){
  const picked = {};
  for(const [k, aliases] of Object.entries(contract.required)) picked[k] = pickColumn(headers, aliases);
  for(const [k, aliases] of Object.entries(contract.optional)) picked[k] = pickColumn(headers, aliases);
  const missing = Object.entries(contract.required).filter(([k])=>!picked[k]).map(([k])=>k);
  return { mapping: picked, missing };
}
function extractRecords(rows, mapping, kind, normalizeVendors){
  return rows.map((r,i)=>{
    const buildingRaw = r[mapping.building];
    const vendorRaw   = r[mapping.vendor];
    const dateRaw     = r[mapping.date];
    const amountRaw   = r[mapping.amount];
    const refRaw      = mapping.reference ? r[mapping.reference] : null;
    const memoRaw     = mapping.memo ? r[mapping.memo] : null;
    const amount = parseMoney(amountRaw);
    return {
      __row: i+2,
      building: normalizeBuilding(buildingRaw), buildingRaw,
      vendor: normalizeName(vendorRaw, normalizeVendors), vendorRaw,
      date: parseDate(dateRaw), dateRaw,
      amount, amountRaw,
      reference: refRaw==null ? "" : String(refRaw),
      memo: memoRaw==null ? "" : String(memoRaw),
      kind
    };
  }).filter(r=> r.amount!=null && r.date!=null && r.vendor && r.building);
}

// ---------------- Dedup logic ----------------
function makeKey(rec, amountTolCents){
  const amt = Math.round(rec.amount*100);
  const bucket = Math.round(amt/amountTolCents);
  return rec.building + "|" + rec.vendor + "|" + bucket;
}
function detectDuplicates(records, daysTol, amountTol){
  const amountTolCents = Math.max(1, Math.round(amountTol*100));
  const map = new Map();
  for(const rec of records){
    const k = makeKey(rec, amountTolCents);
    if(!map.has(k)) map.set(k, []);
    map.get(k).push(rec);
  }
  const dups=[];
  for(const group of map.values()){
    if(group.length<2) continue;
    group.sort((a,b)=> a.date-b.date);
    for(let i=0;i<group.length;i++){
      for(let j=i+1;j<group.length;j++){
        const a=group[i], b=group[j];
        if(dateDiffDays(a.date,b.date) <= daysTol){ dups.push([a,b]); }
      }
    }
  }
  return dups;
}
function crossMatch(unpaid, paid, daysTol){
  const index = new Map();
  const add = rec=>{
    const amt = rec.amount==null?null:Math.round(rec.amount*100);
    const key = rec.building + "|" + rec.vendor + "|" + amt;
    if(!index.has(key)) index.set(key, []);
    index.get(key).push(rec);
  };
  unpaid.forEach(add);
  const matches=[];
  for(const p of paid){
    const amt = p.amount==null?null:Math.round(p.amount*100);
    const key = p.building + "|" + p.vendor + "|" + amt;
    const candidates = index.get(key) || [];
    for(const u of candidates){ if(dateDiffDays(u.date,p.date) <= daysTol) matches.push([u,p]); }
  }
  return matches;
}

// ---------------- Export helpers ----------------
function pairsToRows(pairs, sideA="A", sideB="B"){
  const out=[];
  for(let i=0;i<pairs.length;i++){
    const [a,b]=pairs[i];
    const shape = (rec, side)=>({
      pair:i+1, side,
      source: rec.kind,
      row: rec.__row,
      building: rec.buildingRaw,
      vendor: rec.vendorRaw,
      date: rec.dateRaw,
      amount: rec.amountRaw,
      reference: rec.reference,
      memo: rec.memo
    });
    out.push(shape(a, sideA)); out.push(shape(b, sideB));
  }
  return out;
}
function downloadCSV(filename, rows){
  if(!rows.length){ alert("No rows to export."); return; }
  const head = Object.keys(rows[0]||{});
  const lines = [head.join(",")].concat(rows.map(r=> head.map(k=> JSON.stringify(r[k] ?? "")).join(",")));
  const blob = new Blob([lines.join("\n")], {type:"text/csv;charset=utf-8;"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 1200);
}

// ---------------- State ----------------
let unpaidRows=null, unpaidHeaders=null, unpaidMap=null, unpaidRecords=null, unpaidMissing=[];
let paidRows=null,   paidHeaders=null,   paidMap=null,   paidRecords=null,   paidMissing=[];
let lastResults = { dupsUnpaid: [], dupsPaid: [], cross: [] };
let lastParams = { daysTol: 3, amountTol: 1 };

const viewState = {
  unpaid: { cols:"basic", rows:"compact", filter:"", expanded:new Set(), page:1, pageSize:10 },
  paid:   { cols:"basic", rows:"compact", filter:"", expanded:new Set(), page:1, pageSize:10 },
  cross:  { cols:"basic", rows:"compact", filter:"", expanded:new Set(), page:1, pageSize:10 },
};

// ---------------- Grouping + rendering (with pagination) ----------------
const BASIC_COLS = ["__row","buildingRaw","vendorRaw","dateRaw","amountRaw"];
const EXTRA_COLS = ["reference","memo","kind"];
function fmtAmount(a){
  const n = parseMoney(a);
  return n==null ? (a ?? "") : n.toLocaleString(undefined,{style:"currency",currency:"USD"});
}
function recordId(r){ return `${r.kind}|${r.__row}|${r.amountRaw}|${r.dateRaw}`; }
function keyFor(rec, amountTolCents, forCross=false){
  if(forCross){ return `${rec.building}|${rec.vendor}|${Math.round(rec.amount*100)}`; }
  return makeKey(rec, amountTolCents);
}
function buildGroups(pairs, amountTolCents, forCross=false){
  const graphs = new Map();
  const ensureGraph = k=>{
    if(!graphs.has(k)) graphs.set(k, {nodes:new Map(), adj:new Map()});
    return graphs.get(k);
  };
  for(const [a,b] of pairs){
    const gkey = keyFor(a, amountTolCents, forCross);
    const G = ensureGraph(gkey);
    const ida = recordId(a), idb = recordId(b);
    if(!G.nodes.has(ida)) G.nodes.set(ida, a);
    if(!G.nodes.has(idb)) G.nodes.set(idb, b);
    if(!G.adj.has(ida)) G.adj.set(ida,new Set());
    if(!G.adj.has(idb)) G.adj.set(idb,new Set());
    G.adj.get(ida).add(idb);
    G.adj.get(idb).add(ida);
  }
  const groups=[];
  for(const [gkey, G] of graphs){
    const seen = new Set();
    for(const id of G.nodes.keys()){
      if(seen.has(id)) continue;
      const compIds=[], q=[id]; seen.add(id);
      while(q.length){
        const cur=q.shift(); compIds.push(cur);
        const nbrs = G.adj.get(cur)||new Set();
        for(const nb of nbrs){ if(!seen.has(nb)){ seen.add(nb); q.push(nb); } }
      }
      const recs = compIds.map(x=>G.nodes.get(x)).sort((a,b)=> a.date-b.date || (a.amount-b.amount));
      groups.push({ id: `${gkey}#${groups.length+1}`, gkey, records: recs });
    }
  }
  return groups.sort((a,b)=>{
    const av=a.records[0], bv=b.records[0];
    return (av.vendor.localeCompare(bv.vendor)) || (av.building.localeCompare(bv.building));
  });
}

function renderPager(scope, total){
  const state = viewState[scope];
  const totalPages = Math.max(1, Math.ceil(total / state.pageSize));
  // Ensure page in range
  if(state.page > totalPages) state.page = totalPages;
  if(state.page < 1) state.page = 1;

  const start = total ? (state.page-1)*state.pageSize + 1 : 0;
  const end   = Math.min(total, state.page*state.pageSize);

  const host = document.getElementById(`pager-${scope}`);
  host.innerHTML = `
    <button class="page-btn" data-scope="${scope}" data-dir="prev" ${state.page<=1?'disabled':''}>${t("pager.prev")}</button>
    <span class="page-info">${t("pager.showing")} ${start}–${end} ${t("pager.of")} ${total} ${t("pager.groups")}</span>
    <button class="page-btn" data-scope="${scope}" data-dir="next" ${state.page>=totalPages?'disabled':''}>${t("pager.next")}</button>
    <label class="page-info" for="ps-${scope}">|</label>
    <select id="ps-${scope}" class="page-size" data-scope="${scope}">
      <option value="5"  ${state.pageSize===5?'selected':''}>5</option>
      <option value="10" ${state.pageSize===10?'selected':''}>10</option>
      <option value="25" ${state.pageSize===25?'selected':''}>25</option>
    </select>
  `;
}

function renderSection(containerId, pairs, scope, forCross=false){
  const container = document.getElementById(containerId);
  const amountTolCents = Math.max(1, Math.round(lastParams.amountTol*100));
  const groups = buildGroups(pairs, amountTolCents, forCross);
  const state = viewState[scope];

  // Filter
  const filter = (state.filter||"").trim().toLowerCase();
  const filtered = filter
    ? groups.filter(g => g.records.some(r =>
        (r.vendorRaw||"").toLowerCase().includes(filter) ||
        (r.buildingRaw||"").toLowerCase().includes(filter)))
    : groups;

  // Pagination
  const total = filtered.length;
  const totalPages = Math.max(1, Math.ceil(total / state.pageSize));
  if(state.page > totalPages) state.page = totalPages;
  if(state.page < 1) state.page = 1;
  const startIdx = (state.page-1)*state.pageSize;
  const endIdx   = Math.min(total, startIdx + state.pageSize);
  const pageGroups = filtered.slice(startIdx, endIdx);

  // Column/row modes
  const rowsMode = state.rows;
  const colsMode = state.cols;
  const colsAll = [...BASIC_COLS, ...EXTRA_COLS];
  const tableHeadHtml = `<tr>${["#",...colsAll].map(c=>`<th class="${EXTRA_COLS.includes(c)?"col-extra":""}">${c}</th>`).join("")}</tr>`;

  const groupHtml = pageGroups.map(g=>{
    const all = g.records;
    const showAll = rowsMode==="all" || state.expanded.has(g.id);
    const toShow = showAll ? all : all.slice(0,2);

    const head = all[0] || {};
    const bld = head.buildingRaw || "";
    const ven = head.vendorRaw || "";
    const amt = head.amountRaw ? fmtAmount(head.amountRaw) : "";
    const start = all[0]?.dateRaw || "";
    const end   = all[all.length-1]?.dateRaw || "";

    const tableRows = [];
    for(const r of toShow){
      const cells = colsAll.map(c=>{
        const cls = EXTRA_COLS.includes(c) ? "col-extra" : "";
        const val = (c==="amountRaw") ? fmtAmount(r[c]) : (r[c] ?? "");
        return `<td class="${cls}">${val}</td>`;
      }).join("");
      tableRows.push(`<tr><td class="muted">${r.__row}</td>${cells}</tr>`);
    }

    const remaining = all.length - toShow.length;
    const moreBtn = remaining>0
      ? `<button class="link-btn show-toggle" data-scope="${scope}" data-group="${g.id}" aria-expanded="false">${t("rows.all").replace("All rows","Show "+remaining+" more")}</button>`
      : (showAll && all.length>2
          ? `<button class="link-btn show-toggle" data-scope="${scope}" data-group="${g.id}" aria-expanded="true">Collapse</button>`
          : "");

    return `
      <div class="dup-group ${colsMode==="basic"?"cols-basic":"cols-all"}" data-group="${g.id}">
        <div class="dup-head">
          <div class="dup-meta">
            <span class="kv"><span class="k">Vendor</span> ${ven || "—"}</span>
            <span class="kv"><span class="k">Building</span> ${bld || "—"}</span>
            <span class="kv"><span class="k">Amount</span> ${amt || "≈"}</span>
            <span class="kv"><span class="k">Dates</span> ${start || "—"} → ${end || "—"}</span>
          </div>
          <div class="group-actions">
            <span class="count-chip">${all.filter(r=>r.kind==="Unpaid").length} Unpaid • ${all.filter(r=>r.kind==="Paid").length} Paid</span>
            ${moreBtn}
          </div>
        </div>
        <div class="table-wrap">
          <table class="table">
            <thead>${tableHeadHtml}</thead>
            <tbody>${tableRows.join("") || `<tr><td class="muted">No matches</td></tr>`}</tbody>
          </table>
        </div>
      </div>
    `;
  }).join("");

  container.innerHTML = groupHtml || `<div class="muted small">No matches</div>`;
  renderPager(scope, total);
}

// ---------------- Results bridge ----------------
function renderResults(dupsUnpaid, dupsPaid, cross){
  lastResults = { dupsUnpaid, dupsPaid, cross };
  document.getElementById("count-unpaid").textContent = dupsUnpaid.length;
  document.getElementById("count-paid").textContent   = dupsPaid.length;
  document.getElementById("count-cross").textContent  = cross.length;

  renderSection("dups-unpaid", dupsUnpaid, "unpaid", false);
  renderSection("dups-paid",   dupsPaid,   "paid",   false);
  renderSection("cross",       cross,      "cross",  true);
}

// ---------------- Progress ----------------
function showProgress(){
  el("#progress-container").classList.remove("hidden");
  el("#run").disabled = true;
  el("#run").textContent = t("progress.processing");
}
function hideProgress(){
  el("#progress-container").classList.add("hidden");
  el("#run").disabled = false;
  el("#run").textContent = t("actions.run");
}
function updateProgress(percent, labelKey="progress.preparing"){
  el("#progress-fill").style.width = `${Math.max(0, Math.min(100, percent))}%`;
  el("#progress-text").textContent = `${Math.round(Math.max(0, Math.min(100, percent)))}%`;
  el("#progress-label").textContent = t(labelKey);
}

// ---------------- Readiness ----------------
function computeReadinessIssues(){
  const issues=[];
  if(!unpaidRows) issues.push("Upload Unpaid CSV.");
  if(!paidRows)   issues.push("Upload Paid CSV.");

  if(unpaidRows){
    if(unpaidMissing.length) issues.push("Unpaid: missing required columns → "+unpaidMissing.join(", "));
    const n = (unpaidRecords&&unpaidRecords.length)||0;
    if(n===0) issues.push("Unpaid: 0 usable rows (need building, vendor, date, amount).");
  }
  if(paidRows){
    if(paidMissing.length) issues.push("Paid: missing required columns → "+paidMissing.join(", "));
    const n = (paidRecords&&paidRecords.length)||0;
    if(n===0) issues.push("Paid: 0 usable rows (need building, vendor, date, amount).");
  }
  return issues;
}
function updateReadinessUI(){
  const elx = document.getElementById("readiness");
  const issues = computeReadinessIssues();
  elx.innerHTML = issues.length ? ("Why the button is disabled:<br>• " + issues.join("<br>• ")) : "";
}
function isReady(){ return computeReadinessIssues().length===0; }
function updateRunButton(){
  const btn = document.getElementById("run"); if(!btn) return;
  const ready = isReady();
  btn.disabled = !ready;
  btn.setAttribute("aria-disabled", String(!ready));
  btn.title = ready ? "" : "Upload both CSVs and satisfy required fields";
  updateReadinessUI();
}

// ---------------- Controls wiring ----------------
function wireDropzone(id, inputId, onFile){
  const dz = document.getElementById(id);
  const fileInput = document.getElementById(inputId);
  if(!dz || !fileInput) return;

  const openPicker = () => fileInput.click();
  dz.addEventListener("click", openPicker);
  dz.addEventListener("keydown", e=>{ if(e.key==="Enter"||e.key===" ") openPicker(); });

  ["dragenter","dragover"].forEach(evt=> dz.addEventListener(evt, e=>{
    e.preventDefault(); e.stopPropagation(); dz.classList.add("drag");
  }));
  ["dragleave","drop"].forEach(evt=> dz.addEventListener(evt, e=>{
    e.preventDefault(); e.stopPropagation(); dz.classList.remove("drag");
  }));
  dz.addEventListener("drop", e=>{
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if(f) onFile(f);
  });
  fileInput.addEventListener("change", e=>{
    const f = e.target.files && e.target.files[0];
    if(f) onFile(f);
  });
}

function initResultControls(){
  // Filters
  ["unpaid","paid","cross"].forEach(scope=>{
    const input = document.getElementById(`filter-${scope}`);
    if(input){
      let tmr=null;
      input.addEventListener("input", ()=>{
        clearTimeout(tmr);
        tmr=setTimeout(()=>{
          viewState[scope].filter = input.value || "";
          viewState[scope].page = 1;  // reset to first page when filtering
          rerender(scope);
        }, 150);
      });
    }
  });

  // Segmented toggles
  document.querySelectorAll(".segmented .seg").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      const scope = btn.dataset.scope;
      const kind  = btn.dataset.kind;   // 'rows' | 'cols'
      const value = btn.dataset.value;  // 'compact' | 'all'  OR  'basic' | 'all'
      const siblings = btn.parentElement.querySelectorAll(".seg");
      siblings.forEach(b=>{
        if(b.dataset.kind===kind){
          b.classList.toggle("active", b===btn);
          b.setAttribute("aria-pressed", String(b===btn));
        }
      });
      viewState[scope][kind] = value;
      viewState[scope].page = 1; // reset page when view changes
      rerender(scope);
    });
  });

  // "Show more / Collapse" per-group
  ["dups-unpaid","dups-paid","cross"].forEach(containerId=>{
    const elx = document.getElementById(containerId);
    elx.addEventListener("click", e=>{
      const btn = e.target.closest(".show-toggle");
      if(!btn) return;
      const scope = btn.dataset.scope;
      const gid   = btn.dataset.group;
      const set = viewState[scope].expanded;
      if(set.has(gid)) set.delete(gid); else set.add(gid);
      rerender(scope);
    });
  });

  // Pager delegation
  ["pager-unpaid","pager-paid","pager-cross"].forEach(pid=>{
    const pe = document.getElementById(pid);
    pe.addEventListener("click", e=>{
      const b = e.target.closest(".page-btn");
      if(!b) return;
      const scope = b.dataset.scope;
      if(b.dataset.dir==="prev") viewState[scope].page = Math.max(1, viewState[scope].page-1);
      if(b.dataset.dir==="next") viewState[scope].page = viewState[scope].page+1;
      rerender(scope);
    });
    pe.addEventListener("change", e=>{
      const sel = e.target.closest(".page-size");
      if(!sel) return;
      const scope = sel.dataset.scope;
      viewState[scope].pageSize = parseInt(sel.value,10);
      viewState[scope].page = 1;
      rerender(scope);
    });
  });
}

function rerender(scope){
  if(scope==="unpaid") renderSection("dups-unpaid", lastResults.dupsUnpaid, "unpaid", false);
  else if(scope==="paid") renderSection("dups-paid", lastResults.dupsPaid, "paid", false);
  else if(scope==="cross") renderSection("cross", lastResults.cross, "cross", true);
}

// ---------------- Detection run ----------------
function runDetection(){
  const issues = computeReadinessIssues();
  if(issues.length){ alert('Cannot run yet:\n- ' + issues.join('\n- ')); return; }

  const daysTol   = parseInt(document.getElementById("days-window").value || "3", 10);
  const amountTol = parseFloat(document.getElementById("amount-window").value || "1");
  lastParams = { daysTol, amountTol };

  showProgress(); updateProgress(0, "progress.preparing");
  setTimeout(()=>{
    updateProgress(35, "progress.matching");
    const dupsUnpaid = detectDuplicates(unpaidRecords, daysTol, amountTol);
    const dupsPaid   = detectDuplicates(paidRecords,   daysTol, amountTol);
    updateProgress(65, "progress.crossMatch");
    const cross      = crossMatch(unpaidRecords, paidRecords, daysTol);
    updateProgress(90, "progress.updating");
    renderResults(dupsUnpaid, dupsPaid, cross);
    updateProgress(100, "progress.complete");
    setTimeout(hideProgress, 700);
  }, 50);
}

// ---------------- Init ----------------
document.addEventListener("DOMContentLoaded", ()=>{
  document.getElementById("year").textContent = new Date().getFullYear();

  // Language toggle
  document.getElementById("lang-toggle").addEventListener("click", ()=>{
    currentLang = (currentLang === "en" ? "zh" : "en");
    applyI18n();
    // Also re-render pagers to update labels
    ["unpaid","paid","cross"].forEach(scope=> renderPager(scope, 0)); // dummy to refresh buttons text
    // And rerender sections to refresh "Collapse/Show" text
    rerender("unpaid"); rerender("paid"); rerender("cross");
  });
  applyI18n();

  document.getElementById("unpaid-mapping-toggle").addEventListener("click", ()=> document.getElementById("unpaid-info").classList.toggle("hidden"));
  document.getElementById("paid-mapping-toggle").addEventListener("click",  ()=> document.getElementById("paid-info").classList.toggle("hidden"));
  document.getElementById("run").addEventListener("click", runDetection);

  initResultControls();
  updateRunButton();

  // Uploader wiring
  wireDropzone("dz-unpaid", "file-unpaid", file=>{
    Papa.parse(file, {
      header:true, skipEmptyLines:true, dynamicTyping:false,
      complete: results=>{
        unpaidRows = results.data; unpaidHeaders = results.meta.fields || [];
        const {mapping, missing} = mapRecords(unpaidRows, unpaidHeaders, CONTRACTS.unpaid);
        unpaidMap = mapping; unpaidMissing = missing;
        const normalizeVendors = document.getElementById("norm-vendor").checked;
        unpaidRecords = extractRecords(unpaidRows, mapping, "Unpaid", normalizeVendors);
        showMapping("unpaid-info", CONTRACTS.unpaid.name, mapping, missing, unpaidHeaders);
        const dzTitle = document.querySelector("#dz-unpaid .dz-title");
        if(dzTitle){ dzTitle.textContent = `${file.name} – ${unpaidRows.length} rows loaded`; }
        updateRunButton();
      },
      error: err => alert("Failed to parse CSV: "+err)
    });
  });

  wireDropzone("dz-paid", "file-paid", file=>{
    Papa.parse(file, {
      header:true, skipEmptyLines:true, dynamicTyping:false,
      complete: results=>{
        paidRows = results.data; paidHeaders = results.meta.fields || [];
        const {mapping, missing} = mapRecords(paidRows, paidHeaders, CONTRACTS.paid);
        paidMap = mapping; paidMissing = missing;
        const normalizeVendors = document.getElementById("norm-vendor").checked;
        paidRecords = extractRecords(paidRows, mapping, "Paid", normalizeVendors);
        showMapping("paid-info", CONTRACTS.paid.name, mapping, missing, paidHeaders);
        const dzTitle = document.querySelector("#dz-paid .dz-title");
        if(dzTitle){ dzTitle.textContent = `${file.name} – ${paidRows.length} rows loaded`; }
        updateRunButton();
      },
      error: err => alert("Failed to parse CSV: "+err)
    });
  });

  // Export buttons (export all pairs, not only current page)
  document.getElementById("export-unpaid").addEventListener("click", ()=>{
    const rows = pairsToRows(lastResults.dupsUnpaid,"A","B");
    rows.length ? downloadCSV("duplicates_unpaid.csv", rows) : alert("No duplicates found in Unpaid section.");
  });
  document.getElementById("export-paid").addEventListener("click", ()=>{
    const rows = pairsToRows(lastResults.dupsPaid,"A","B");
    rows.length ? downloadCSV("duplicates_paid.csv", rows) : alert("No duplicates found in Paid section.");
  });
  document.getElementById("export-cross").addEventListener("click", ()=>{
    const rows = pairsToRows(lastResults.cross,"Unpaid","Paid");
    rows.length ? downloadCSV("paid_vs_unpaid.csv", rows) : alert("No matches found between Paid and Unpaid.");
  });
});

// --------------- Mapping UI ---------------
function showMapping(elId, contractName, mapping, missing, headers){
  const el = document.getElementById(elId);
  const rows = Object.entries(mapping).map(([k,v])=> `<tr><td>${k}</td><td><code>${v||""}</code></td></tr>`).join("");
  const miss = (missing && missing.length)
    ? `<div style="color:#b91c1c"><strong>Missing: ${missing.join(", ")}</strong></div>`
    : `<div style="color:#047857">Ready</div>`;
  el.innerHTML = `
    <div><strong>Detected mapping</strong> · <em>${contractName}</em></div>
    ${miss}
    <div class="table-wrap" style="margin-top:8px">
      <table class="table">
        <thead><tr><th>Field</th><th>Column</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    <div class="muted small" style="margin-top:6px">Headers: ${headers.map(h=>`<code>${h}</code>`).join(" ")}</div>
  `;
}