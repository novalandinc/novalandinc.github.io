/* =========================================================
   Duplicates Auditor — app.js
   (Group-level horizontal scroll; no per-cell scrollers)
   ========================================================= */

/* =============== i18n (kept minimal; EN default) =============== */
let currentLang = "en";
function applyI18n(){
  document.title = "Duplicates Auditor";
  document.getElementById("lang-label").textContent = currentLang === "en" ? "中文" : "EN";
}

/* =============== Column contracts =============== */
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

/* =============== Helpers =============== */
const el = sel => document.querySelector(sel);
function normalizeName(name, enabled = true){
  if(name==null) return "";
  let v = String(name);
  if(!enabled) return v.trim().toUpperCase();
  v = v.toUpperCase()
       .replace(/&/g,"AND")
       .replace(/[^A-Z0-9 ]+/g," ")
       .replace(/\b(LLC|INC|CO|CORP|THE)\b/g,"")
       .replace(/\s+/g," ").trim();
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
function escapeHtml(s){
  return String(s)
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;");
}

/* =============== Mapping & extraction =============== */
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

/* =============== Duplicate detection =============== */
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

/* =============== Export helpers =============== */
function fmtAmount(a){
  const n = parseMoney(a);
  return n==null ? (a ?? "") : n.toLocaleString(undefined,{style:"currency",currency:"USD"});
}
const BASIC_COLS = ["__row","buildingRaw","vendorRaw","dateRaw","amountRaw"];
const EXTRA_COLS = ["reference","memo","kind"];
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

/* =============== Grouping + rendering (shared) =============== */
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
  if(state.page > totalPages) state.page = totalPages;
  if(state.page < 1) state.page = 1;

  const start = total ? (state.page-1)*state.pageSize + 1 : 0;
  const end   = Math.min(total, state.page*state.pageSize);

  const host = document.getElementById(`pager-${scope}`);
  host.innerHTML = `
    <button class="page-btn" data-scope="${scope}" data-dir="prev" ${state.page<=1?'disabled':''}>‹ Prev</button>
    <span class="page-info">Showing ${start}–${end} of ${total} groups</span>
    <button class="page-btn" data-scope="${scope}" data-dir="next" ${state.page>=totalPages?'disabled':''}>Next ›</button>
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

  const filter = (state.filter||"").trim().toLowerCase();
  const filtered = filter
    ? groups.filter(g => g.records.some(r =>
        (r.vendorRaw||"").toLowerCase().includes(filter) ||
        (r.buildingRaw||"").toLowerCase().includes(filter)))
    : groups;

  const total = filtered.length;
  const totalPages = Math.max(1, Math.ceil(total / state.pageSize));
  if(state.page > totalPages) state.page = totalPages;
  if(state.page < 1) state.page = 1;
  const startIdx = (state.page-1)*state.pageSize;
  const endIdx   = Math.min(total, startIdx + state.pageSize);
  const pageGroups = filtered.slice(startIdx, endIdx);

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
        const isExtra = EXTRA_COLS.includes(c);
        const raw = c==="amountRaw" ? fmtAmount(r[c]) : (r[c] ?? "");
        const tdClass = `${isExtra?"col-extra":""}`;
        // Title attribute provides quick full-text on hover, too
        const titleAttr = `title="${escapeHtml(r[c] ?? "")}"`;
        return `<td class="${tdClass}" ${titleAttr}>${escapeHtml(raw)}</td>`;
      }).join("");
      tableRows.push(`<tr><td class="muted">${r.__row}</td>${cells}</tr>`);
    }

    const remaining = all.length - toShow.length;
    const moreBtn = remaining>0
      ? `<button class="link-btn show-toggle" data-scope="${scope}" data-group="${g.id}" aria-expanded="false">Show ${remaining} more</button>`
      : (showAll && all.length>2
          ? `<button class="link-btn show-toggle" data-scope="${scope}" data-group="${g.id}" aria-expanded="true">Collapse</button>`
          : "");

    return `
      <div class="dup-group ${colsMode==="basic"?"cols-basic":"cols-all"}" data-group="${g.id}">
        <div class="dup-head">
          <div class="dup-meta">
            <span class="kv"><span class="k">Vendor</span> ${escapeHtml(ven || "—")}</span>
            <span class="kv"><span class="k">Building</span> ${escapeHtml(bld || "—")}</span>
            <span class="kv"><span class="k">Amount</span> ${escapeHtml(amt || "≈")}</span>
            <span class="kv"><span class="k">Dates</span> ${escapeHtml(start || "—")} → ${escapeHtml(end || "—")}</span>
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

/* =============== State =============== */
let mode = "single"; // 'single' | 'legacy'

// Single-mode
let singleRows=null, singleHeaders=null, singleMap=null, singleRecords=null, singleMissing=[], singleContract="";

// Legacy-mode
let unpaidRows=null, unpaidHeaders=null, unpaidMap=null, unpaidRecords=null, unpaidMissing=[];
let paidRows=null,   paidHeaders=null,   paidMap=null,   paidRecords=null,   paidMissing=[];

let lastResults = { single: [], dupsUnpaid: [], dupsPaid: [], cross: [] };
let lastParams = { daysTol: 3, amountTol: 1 };

const viewState = {
  single: { cols:"basic", rows:"compact", filter:"", expanded:new Set(), page:1, pageSize:10 },
  unpaid: { cols:"basic", rows:"compact", filter:"", expanded:new Set(), page:1, pageSize:10 },
  paid:   { cols:"basic", rows:"compact", filter:"", expanded:new Set(), page:1, pageSize:10 },
  cross:  { cols:"basic", rows:"compact", filter:"", expanded:new Set(), page:1, pageSize:10 },
};

/* =============== Mode UI & readiness =============== */
function setMode(newMode){
  if(mode === newMode) return;
  mode = newMode;

  const tabSingle = el("#tab-single");
  const tabLegacy = el("#tab-legacy");
  tabSingle.classList.toggle("active", mode==="single");
  tabLegacy.classList.toggle("active", mode==="legacy");
  tabSingle.setAttribute("aria-selected", String(mode==="single"));
  tabLegacy.setAttribute("aria-selected", String(mode==="legacy"));

  el("#panel-single").classList.toggle("hidden", mode!=="single");
  el("#panel-legacy").classList.toggle("hidden", mode!=="legacy");

  const steps = el("#steps");
  if(mode==="single"){
    steps.innerHTML = `
      <li>Upload <strong>one</strong> CSV</li>
      <li>Adjust settings (optional)</li>
      <li>Run detection</li>
    `;
  }else{
    steps.innerHTML = `
      <li>Upload <strong>Unpaid</strong> CSV</li>
      <li>Upload <strong>Paid</strong> CSV</li>
      <li>Adjust settings (optional)</li>
      <li>Run detection</li>
    `;
  }
  updateRunButton();
}
function computeReadinessIssues(){
  const issues=[];
  if(mode==="single"){
    if(!singleRows) issues.push("Upload a CSV.");
    if(singleRows){
      if(singleMissing.length) issues.push("Missing required columns → "+singleMissing.join(", "));
      const n = (singleRecords&&singleRecords.length)||0;
      if(n===0) issues.push("0 usable rows (need building, vendor, date, amount).");
    }
  }else{
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
  const btn = document.getElementById("run");
  const ready = isReady();
  btn.disabled = !ready;
  btn.setAttribute("aria-disabled", String(!ready));
  btn.title = ready ? "" : (mode==="single" ? "Upload a CSV and satisfy required fields" : "Upload both CSVs and satisfy required fields");
  updateReadinessUI();
}

/* =============== Controls wiring =============== */
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

/* =============== View wiring (filters, toggles, pager) =============== */
function initResultControls(){
  ["single","unpaid","paid","cross"].forEach(scope=>{
    const input = document.getElementById(`filter-${scope}`);
    if(input){
      let tmr=null;
      input.addEventListener("input", ()=>{
        clearTimeout(tmr);
        tmr=setTimeout(()=>{
          viewState[scope].filter = input.value || "";
          viewState[scope].page = 1;
          rerender(scope);
        }, 150);
      });
    }
  });

  document.querySelectorAll(".segmented .seg").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      const scope = btn.dataset.scope;
      const kind  = btn.dataset.kind;
      const value = btn.dataset.value;
      const siblings = btn.parentElement.querySelectorAll(".seg");
      siblings.forEach(b=>{
        if(b.dataset.kind===kind){
          b.classList.toggle("active", b===btn);
          b.setAttribute("aria-pressed", String(b===btn));
        }
      });
      viewState[scope][kind] = value;
      viewState[scope].page = 1;
      rerender(scope);
    });
  });

  ["dups-single","dups-unpaid","dups-paid","cross"].forEach(containerId=>{
    const elx = document.getElementById(containerId);
    if(!elx) return;
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

  ["pager-single","pager-unpaid","pager-paid","pager-cross"].forEach(pid=>{
    const pe = document.getElementById(pid);
    if(!pe) return;
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
  if(scope==="single") renderSection("dups-single", lastResults.single, "single", false);
  else if(scope==="unpaid") renderSection("dups-unpaid", lastResults.dupsUnpaid, "unpaid", false);
  else if(scope==="paid") renderSection("dups-paid", lastResults.dupsPaid, "paid", false);
  else if(scope==="cross") renderSection("cross", lastResults.cross, "cross", true);
}

/* =============== Single-mode: auto-detect mapping =============== */
function chooseBestContract(headers, rows){
  const tries = [
    ["unpaid", CONTRACTS.unpaid],
    ["paid", CONTRACTS.paid]
  ].map(([key, c])=>{
    const {mapping, missing} = mapRecords(rows, headers, c);
    const normalizeVendors = document.getElementById("norm-vendor").checked;
    const recs = extractRecords(rows, mapping, c.name, normalizeVendors);
    return { key, contract:c, mapping, missing, usable: recs.length, records: recs };
  });

  tries.sort((a,b)=>{
    if(a.missing.length !== b.missing.length) return a.missing.length - b.missing.length;
    return b.usable - a.usable;
  });

  return tries[0];
}

/* =============== Progress helpers =============== */
function showProgress(){
  el("#progress-container").classList.remove("hidden");
  el("#run").disabled = true;
  el("#run").textContent = "Processing...";
}
function hideProgress(){
  el("#progress-container").classList.add("hidden");
  el("#run").disabled = false;
  el("#run").textContent = "Run detection";
}
function updateProgress(percent, label="Preparing..."){
  el("#progress-fill").style.width = `${Math.max(0, Math.min(100, percent))}%`;
  el("#progress-text").textContent = `${Math.round(Math.max(0, Math.min(100, percent)))}%`;
  el("#progress-label").textContent = label;
}

/* =============== Run detection (mode aware) =============== */
function runDetection(){
  const issues = computeReadinessIssues();
  if(issues.length){ alert('Cannot run yet:\n- ' + issues.join('\n- ')); return; }

  const daysTol   = parseInt(document.getElementById("days-window").value || "3", 10);
  const amountTol = parseFloat(document.getElementById("amount-window").value || "1");
  lastParams = { daysTol, amountTol };

  showProgress(); updateProgress(0, "Preparing...");
  setTimeout(()=>{
    if(mode==="single"){
      updateProgress(40, "Matching duplicates...");
      const dups = detectDuplicates(singleRecords, daysTol, amountTol);
      updateProgress(85, "Updating results...");
      lastResults.single = dups;
      document.getElementById("count-single").textContent = dups.length;
      renderSection("dups-single", dups, "single", false);
    }else{
      updateProgress(35, "Matching duplicates...");
      const dupsUnpaid = detectDuplicates(unpaidRecords, daysTol, amountTol);
      const dupsPaid   = detectDuplicates(paidRecords,   daysTol, amountTol);
      updateProgress(65, "Cross-matching paid vs unpaid...");
      const cross      = crossMatch(unpaidRecords, paidRecords, daysTol);
      updateProgress(90, "Updating results...");
      lastResults.dupsUnpaid = dupsUnpaid;
      lastResults.dupsPaid   = dupsPaid;
      lastResults.cross      = cross;
      document.getElementById("count-unpaid").textContent = dupsUnpaid.length;
      document.getElementById("count-paid").textContent   = dupsPaid.length;
      document.getElementById("count-cross").textContent  = cross.length;
      renderSection("dups-unpaid", dupsUnpaid, "unpaid", false);
      renderSection("dups-paid",   dupsPaid,   "paid",   false);
      renderSection("cross",       cross,      "cross",  true);
    }
    updateProgress(100, "Complete!");
    setTimeout(hideProgress, 700);
  }, 50);
}

/* =============== Mapping panel helper =============== */
function showMapping(elId, contractName, mapping, missing, headers){
  const host = document.getElementById(elId);
  const rows = Object.entries(mapping).map(([k,v])=> `<tr><td>${k}</td><td><code>${v||""}</code></td></tr>`).join("");
  const miss = (missing && missing.length)
    ? `<div style="color:#b91c1c"><strong>Missing: ${missing.join(", ")}</strong></div>`
    : `<div style="color:#047857">Ready</div>`;
  host.innerHTML = `
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

/* =============== Boot =============== */
document.addEventListener("DOMContentLoaded", ()=>{
  document.getElementById("year").textContent = new Date().getFullYear();

  document.getElementById("lang-toggle").addEventListener("click", ()=>{
    currentLang = (currentLang === "en" ? "zh" : "en");
    applyI18n();
  });
  applyI18n();

  document.getElementById("tab-single").addEventListener("click", ()=> setMode("single"));
  document.getElementById("tab-legacy").addEventListener("click", ()=> setMode("legacy"));
  setMode("single");

  document.getElementById("run").addEventListener("click", runDetection);

  initResultControls();
  updateRunButton();

  /* Single uploader */
  wireDropzone("dz-single", "file-single", file=>{
    Papa.parse(file, {
      header:true, skipEmptyLines:true, dynamicTyping:false,
      complete: results=>{
        singleRows = results.data; singleHeaders = results.meta.fields || [];
        const best = chooseBestContract(singleHeaders, singleRows);
        singleContract = best.contract.name;
        singleMap = best.mapping; singleMissing = best.missing; singleRecords = best.records;

        showMapping("single-info", `Auto • ${best.contract.name}`, best.mapping, best.missing, singleHeaders);
        const dzTitle = document.querySelector("#dz-single .dz-title");
        if(dzTitle){ dzTitle.textContent = `${file.name} – ${singleRows.length} rows loaded (${best.contract.name})`; }
        updateRunButton();
      },
      error: err => alert("Failed to parse CSV: "+err)
    });
  });
  document.getElementById("single-mapping-toggle").addEventListener("click", ()=> document.getElementById("single-info").classList.toggle("hidden"));

  /* Legacy uploaders */
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

  // Exports
  document.getElementById("export-single").addEventListener("click", ()=>{
    const rows = pairsToRows(lastResults.single,"A","B");
    rows.length ? downloadCSV("duplicates.csv", rows) : alert("No duplicates found.");
  });
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
