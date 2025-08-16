// ---------------- i18n ----------------
const I18N = {
  en: {
    "header.title": "Internal Tools",
    "header.subtitle": "Lightweight utilities for operations",
    "tabs.home": "Home", "tabs.payments": "Payments",
    "home.headline": "Welcome to Internal Tools",
    "home.tagline": "A simple hub for day-to-day operations. Start with the Payments Duplicate Auditor below.",
    "home.open": "Open",
    "home.payments.title": "Payments Duplicate Auditor",
    "home.payments.desc": "Detect duplicate outgoing payments and open charges across CSV exports.",
    "payments.sidebar.title": "Duplicate Auditor",
    "payments.steps.1": "Upload Unpaid Bills CSV",
    "payments.steps.2": "Upload Paid Bills CSV",
    "payments.steps.3": "Adjust settings (optional)",
    "payments.steps.4": "Run detection",
    "payments.detection.howItWorks": "How duplicate detection works",
    "payments.detection.criteria": "Detection criteria:",
    "payments.detection.building": "Same building/property",
    "payments.detection.vendor": "Same vendor (normalized)",
    "payments.detection.amount": "Same amount (within tolerance)",
    "payments.detection.date": "Same date (within window)",
    "payments.settings.dateWindow": "Date window (days)",
    "payments.settings.amountTol": "Amount tolerance ($)",
    "payments.settings.advanced": "Advanced options",
    "payments.settings.vendorNorm": "Vendor normalization ignores punctuation, LLC/Inc/Co/Corp, \"THE\", and treats \"&\" as \"AND\".",
    "payments.settings.signNote": "Paid amounts are treated as positive outflows automatically.",
    "payments.settings.dateFormats": "Supports multiple date formats like MM/DD/YYYY, DD-MM-YYYY, YYYY-MM-DD, MM.DD.YYYY, YYYYMMDD, and more. Handles 2-digit years automatically.",
    "payments.run": "Run detection",
    "payments.upload.unpaid": "Unpaid Bills",
    "payments.upload.paid": "Paid Bills",
    "payments.upload.drop": "Drop CSV here or click to select",
    "payments.upload.hintsUnpaid": "Required columns: EntryDate • VendorName • Amount • PropertyName",
    "payments.upload.hintsPaid": "Required columns: entryDate • payeeNameRaw • amount • buildingName",
    "payments.tabs.unpaid": "Duplicates: Unpaid",
    "payments.tabs.paid": "Duplicates: Paid",
    "payments.tabs.cross": "Paid vs Unpaid",
    "payments.export": "Download CSV",
    "progress.preparing": "Preparing...",
    "progress.processing": "Processing...",
    "progress.normalizing": "Normalizing data...",
    "progress.matching": "Matching duplicates...",
    "progress.crossMatch": "Cross-matching paid vs unpaid...",
    "progress.updating": "Updating results...",
    "progress.complete": "Complete!",
    "footer.company": "Novaland, Inc."
  },
  zh: {
    "header.title": "内部工具",
    "header.subtitle": "面向运营的轻量工具",
    "tabs.home": "主页", "tabs.payments": "付款",
    "home.headline": "欢迎使用内部工具",
    "home.tagline": "用于日常工作的简易工具集合。可从下方的付款重复检测开始。",
    "home.open": "打开",
    "home.payments.title": "付款重复检测",
    "home.payments.desc": "在多个 CSV 导出中检测重复的付款和未付费用。",
    "payments.sidebar.title": "重复检测",
    "payments.steps.1": "上传未付账单 CSV",
    "payments.steps.2": "上传已付账单 CSV",
    "payments.steps.3": "调整设置（可选）",
    "payments.steps.4": "开始检测",
    "payments.detection.howItWorks": "重复检测的原理",
    "payments.detection.criteria": "检测依据：",
    "payments.detection.building": "相同物业/楼宇",
    "payments.detection.vendor": "相同供应商（归一化）",
    "payments.detection.amount": "相同金额（容差内）",
    "payments.detection.date": "相同日期（时间窗内）",
    "payments.settings.dateWindow": "日期窗口（天）",
    "payments.settings.amountTol": "金额容差（$）",
    "payments.settings.advanced": "高级选项",
    "payments.settings.vendorNorm": "供应商名归一化会忽略标点、LLC/Inc/Co/Corp、“THE”，并将“&”视作“AND”。",
    "payments.settings.signNote": "已付金额会自动视为支出（正值）。",
    "payments.settings.dateFormats": "支持多种日期格式：MM/DD/YYYY、DD-MM-YYYY、YYYY-MM-DD、MM.DD.YYYY、YYYYMMDD 等；自动处理两位年份。",
    "payments.run": "开始检测",
    "payments.upload.unpaid": "未付账单",
    "payments.upload.paid": "已付账单",
    "payments.upload.drop": "拖拽 CSV 到此处或点击选择",
    "payments.upload.hintsUnpaid": "必需列：EntryDate • VendorName • Amount • PropertyName",
    "payments.upload.hintsPaid": "必需列：entryDate • payeeNameRaw • amount • buildingName",
    "payments.tabs.unpaid": "重复：未付",
    "payments.tabs.paid": "重复：已付",
    "payments.tabs.cross": "已付 vs 未付",
    "payments.export": "下载 CSV",
    "progress.preparing": "准备中...",
    "progress.processing": "处理中...",
    "progress.normalizing": "正在归一化数据...",
    "progress.matching": "正在匹配重复项...",
    "progress.crossMatch": "正在交叉匹配已付 vs 未付...",
    "progress.updating": "正在更新结果...",
    "progress.complete": "完成！",
    "footer.company": "Novaland, Inc."
  }
};
let currentLang = "en";
function t(key){ const dict = I18N[currentLang] || I18N.en; return dict[key] || I18N.en[key] || key; }
function applyI18n(){
  document.querySelectorAll("[data-i18n]").forEach(el=>{
    const key = el.getAttribute("data-i18n");
    const dict = I18N[currentLang] || I18N.en;
    if(key && dict[key]) el.textContent = dict[key];
  });
  document.getElementById("lang-label").textContent = currentLang === "en" ? "中文" : "EN";
}

// ---------------- Tabs (fixed) ----------------
function initNavTabs(){
  // Main nav (Home/Payments)
  document.querySelectorAll('nav[role="tablist"] .tab').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      const target = btn.dataset.tab;
      setActivePanel(target);
      // update active state in nav
      document.querySelectorAll('nav[role="tablist"] .tab').forEach(b=>{
        b.classList.toggle('active', b===btn);
        b.setAttribute('aria-selected', String(b===btn));
      });
    });
  });

  // Results tabs inside Payments
  document.querySelectorAll('#payments .tabs .tab.small').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      setActiveResultsTab(btn.dataset.tab);
      document.querySelectorAll('#payments .tabs .tab.small').forEach(b=>{
        b.classList.toggle('active', b===btn);
      });
    });
  });

  // “Open” button from Home
  document.querySelectorAll('[data-open]').forEach(btn=>{
    btn.addEventListener('click', ()=> setActivePanel(btn.dataset.open));
  });
}
function setActivePanel(id){
  document.querySelectorAll('.panel').forEach(p=>{
    const active = p.id === id;
    p.classList.toggle('active', active);
    p.setAttribute('aria-hidden', String(!active));
  });
}
function setActiveResultsTab(id){
  document.querySelectorAll('.results').forEach(el=> el.classList.add('hidden'));
  document.getElementById(id).classList.remove('hidden');
}

// ---------------- Column contracts ----------------
const CONTRACTS = {
  unpaid: {
    name: "Unpaid Bills",
    required: {
      building: ['PropertyName','Building','buildingName','Property','Property Name','Building Name'],
      vendor:   ['VendorName','Vendor','Payee','PayeeName','payee','payeeName'],
      date:     ['EntryDate','DueDate','InvoiceDate','Date','TxnDate','Transaction Date'],
      amount:   ['Amount','AmountDue','Balance','Total','OpenBalance'],
    },
    optional: {
      reference:['RefNo','InvoiceNo','Invoice #','Ref #','referenceNumber','reference','CheckNo','checkNo'],
      memo:     ['Memo','Description','Notes','Note','Details']
    }
  },
  paid: {
    name: "Paid Bills",
    required: {
      building: ['buildingName','PropertyName','Building','Property','Property Name','Building Name'],
      vendor:   ['payeeNameRaw','VendorName','Vendor','Payee','PayeeName','payee','payeeName','vendorNameRaw'],
      date:     ['entryDate','PostedDate','ClearedDate','PaymentDate','Date','TxnDate','Transaction Date','postedOn'],
      amount:   ['amount','Amount','AmountOut','Amount Out','Total'],
    },
    optional: {
      reference:['CheckNo','RefNo','Reference #','Check #','ChequeNo','referenceNumber','reference'],
      memo:     ['memo','Memo','Description','Notes','Note','Details']
    }
  }
};

// ---------------- Helpers ----------------
function pickColumn(headers, candidates){
  if(!headers || !headers.length) return null;
  const lower = headers.map(h=>h.toLowerCase());
  for(const cand of candidates){
    const idx = lower.indexOf(cand.toLowerCase());
    if(idx !== -1) return headers[idx];
  }
  const norm = s => s.toLowerCase().replace(/[\s_\-/#.]+/g,'').replace(/[()]/g,'');
  const nheaders = headers.map(norm);
  for(const cand of candidates.map(norm)){
    const idx = nheaders.indexOf(cand);
    if(idx !== -1) return headers[idx];
  }
  return null;
}
function normalizeName(name){
  if(name==null) return '';
  let v=String(name).toUpperCase();
  v=v.replace(/&/g,'AND').replace(/[^A-Z0-9 ]+/g,' ').replace(/\b(LLC|INC|CO|CORP|THE)\b/g,'').replace(/\s+/g,' ').trim();
  return v;
}
function normalizeBuilding(b){ return (b==null?'':String(b)).trim().toUpperCase(); }
function parseMoney(a){
  if(a==null || a==='') return null;
  const s=String(a).replace(/[\$,]/g,'').trim();
  const n=Number(s);
  return isNaN(n)?null:Math.abs(n);
}
function parseDate(d){
  if(!d && d!==0) return null;
  const s = String(d).trim();
  if(!s) return null;
  const parts = [
    /^(\d{4})-(\d{1,2})-(\d{1,2})$/,           // 2024-07-01
    /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/,       // 7/1/24 or 07/01/2024
    /^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/,       // 07.01.2024
    /^(\d{1,2})-(\d{1,2})-(\d{2,4})$/,         // 07-01-2024
    /^(\d{4})(\d{2})(\d{2})$/,                 // 20240701
  ];
  let y,m,da;
  for(const re of parts){
    const mm = s.match(re);
    if(mm){
      if(re===parts[0]){ y=+mm[1]; m=+mm[2]; da=+mm[3]; }
      else if(re===parts[1]||re===parts[2]||re===parts[3]){
        const a=+mm[1], b=+mm[2], c=+mm[3]; y = (mm[3].length===2) ? (2000+c) : c; m=a; da=b;
      } else if(re===parts[4]){ y=+mm[1]; m=+mm[2]; da=+mm[3]; }
      const dt = new Date(y, m-1, da);
      if(!isNaN(dt)) return dt;
    }
  }
  const dt = new Date(s);
  return isNaN(dt)?null:dt;
}
function dateDiffDays(a,b){
  if(!a || !b) return Infinity;
  const ms=Math.abs(a.getTime()-b.getTime());
  return Math.floor(ms/86400000);
}

// ---------------- CSV parsing ----------------
function parseCSV(file, cb){
  if (typeof Papa === 'undefined') {
    alert('CSV parser not loaded. Please refresh and try again.');
    return;
  }
  Papa.parse(file, {
    header: true, skipEmptyLines: true, dynamicTyping: false,
    complete: results => cb(results.data, results.meta.fields || []),
    error: err => { console.error('CSV parse error', err); alert('Failed to parse CSV: '+err); }
  });
}

function mapRecords(rows, headers, contract){
  const picked = {};
  for(const [k, aliases] of Object.entries(contract.required)){ picked[k]=pickColumn(headers, aliases); }
  for(const [k, aliases] of Object.entries(contract.optional)){ picked[k]=pickColumn(headers, aliases); }
  const missing = Object.entries(contract.required).filter(([k])=>!picked[k]).map(([k])=>k);
  return { mapping: picked, missing };
}

function extractRecords(rows, mapping, kind){
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
      vendor: normalizeName(vendorRaw),         vendorRaw,
      date: parseDate(dateRaw),                 dateRaw,
      amount, amountRaw,
      reference: refRaw==null ? '' : String(refRaw),
      memo: memoRaw==null ? '' : String(memoRaw),
      kind
    };
  }).filter(r=> r.amount!=null && r.date!=null && r.vendor && r.building);
}

// ---------------- Dedup logic ----------------
function makeKey(rec, amountTolCents){
  const amt = Math.round(rec.amount*100);
  const bucket = Math.round(amt/amountTolCents);
  return rec.building + '|' + rec.vendor + '|' + bucket;
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
    const key = rec.building + '|' + rec.vendor + '|' + amt;
    if(!index.has(key)) index.set(key, []);
    index.get(key).push(rec);
  };
  unpaid.forEach(add);
  const matches=[];
  for(const p of paid){
    const amt = p.amount==null?null:Math.round(p.amount*100);
    const key = p.building + '|' + p.vendor + '|' + amt;
    const candidates = index.get(key) || [];
    for(const u of candidates){ if(dateDiffDays(u.date,p.date) <= daysTol) matches.push([u,p]); }
  }
  return matches;
}

// ---------------- Export ----------------
function downloadCSV(filename, rows){
  const head = Object.keys(rows[0]||{});
  const lines = [head.join(',')].concat(rows.map(r=> head.map(k=> JSON.stringify(r[k] ?? '')).join(',')));
  const blob = new Blob([lines.join('\n')], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 2000);
}

// ---------------- State ----------------
let unpaidRows=null, unpaidHeaders=null, unpaidMap=null, unpaidRecords=null, unpaidMissing=[];
let paidRows=null,   paidHeaders=null,   paidMap=null,   paidRecords=null,   paidMissing=[];

// Compute & show readiness reasons
function computeReadinessIssues(){
  const issues=[];
  if(!unpaidRows) issues.push('Upload Unpaid CSV.');
  if(!paidRows)   issues.push('Upload Paid CSV.');

  if(unpaidRows){
    if(unpaidMissing.length) issues.push('Unpaid: missing required columns → '+unpaidMissing.join(', '));
    const n = (unpaidRecords&&unpaidRecords.length)||0;
    if(n===0) issues.push('Unpaid: 0 usable rows after filtering (need building, vendor, date, amount).');
  }
  if(paidRows){
    if(paidMissing.length) issues.push('Paid: missing required columns → '+paidMissing.join(', '));
    const n = (paidRecords&&paidRecords.length)||0;
    if(n===0) issues.push('Paid: 0 usable rows after filtering (need building, vendor, date, amount).');
  }
  return issues;
}
function updateReadinessUI(){
  const el = document.getElementById('readiness');
  const issues = computeReadinessIssues();
  if(issues.length){
    el.innerHTML = 'Why the button is disabled:<br>• ' + issues.join('<br>• ');
  }else{
    el.textContent = '';
  }
}
function isReady(){ return computeReadinessIssues().length===0; }
function updateRunButton(){
  const btn = document.getElementById('run'); if(!btn) return;
  const ready = isReady();
  btn.disabled = !ready;
  btn.setAttribute('aria-disabled', String(!ready));
  btn.title = ready ? '' : 'Upload both CSVs and satisfy required fields';
  updateReadinessUI();
}

// ---------------- UI rendering ----------------
function showMapping(elId, contractName, mapping, missing, headers){
  const el = document.getElementById(elId);
  const rows = Object.entries(mapping).map(([k,v])=> `<tr><td>${k}</td><td><code>${v||''}</code></td></tr>`).join('');
  const miss = (missing && missing.length)
    ? `<div style="color:#ef4444"><strong>Missing: ${missing.join(', ')}</strong></div>`
    : `<div style="color:#22c55e">Ready</div>`;
  el.innerHTML = `
    <div><strong>Detected mapping</strong> · <em>${contractName}</em></div>
    ${miss}
    <table class="mapping table">
      <thead><tr><th>Field</th><th>Column</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>
    <div class="muted small">Headers: ${headers.map(h=>`<code>${h}</code>`).join(' ')}</div>
  `;
}

function showProgress(){
  document.getElementById('progress-container').classList.remove('hidden');
  document.getElementById('run').disabled = true;
  document.getElementById('run').textContent = t('progress.processing');
}
function hideProgress(){
  document.getElementById('progress-container').classList.add('hidden');
  document.getElementById('run').disabled = false;
  document.getElementById('run').textContent = t('payments.run');
}
function updateProgress(percent, labelKey='progress.preparing'){
  document.getElementById('progress-fill').style.width = `${percent}%`;
  document.getElementById('progress-text').textContent = `${Math.round(percent)}%`;
  document.getElementById('progress-label').textContent = t(labelKey);
}

function renderResults(dupsUnpaid, dupsPaid, cross){
  function table(pairs){
    const cols = ['__row','buildingRaw','vendorRaw','dateRaw','amountRaw','reference','memo'];
    const head = `<tr>${['#',...cols].map(c=>`<th>${c}</th>`).join('')}</tr>`;
    const rows = pairs.map(([a,b],i)=>`
      <tr><td class="muted">${i+1}</td>${cols.map(c=>`<td>${a[c]??''}</td>`).join('')}</tr>
      <tr><td></td>${cols.map(c=>`<td>${b[c]??''}</td>`).join('')}</tr>
    `).join('');
    return `<table class="table">${head}${rows || '<tr><td class="muted">No matches</td></tr>'}</table>`;
  }
  document.getElementById('dups-unpaid').innerHTML = table(dupsUnpaid);
  document.getElementById('dups-paid').innerHTML   = table(dupsPaid);
  document.getElementById('cross').innerHTML       = table(cross);
}

// ---------------- Run detection ----------------
function runDetection(){
  const issues = computeReadinessIssues();
  if(issues.length){
    alert('Cannot run yet:\n- ' + issues.join('\n- '));
    return;
  }
  const daysTol   = parseInt(document.getElementById('days-window').value || '3', 10);
  const amountTol = parseFloat(document.getElementById('amount-window').value || '1');

  showProgress(); updateProgress(0, 'progress.preparing');
  setTimeout(()=>{
    updateProgress(35, 'progress.matching');
    const dupsUnpaid = detectDuplicates(unpaidRecords, daysTol, amountTol);
    const dupsPaid   = detectDuplicates(paidRecords,   daysTol, amountTol);
    updateProgress(65, 'progress.crossMatch');
    const cross      = crossMatch(unpaidRecords, paidRecords, daysTol);
    updateProgress(90, 'progress.updating');
    renderResults(dupsUnpaid, dupsPaid, cross);
    updateProgress(100, 'progress.complete');
    setTimeout(hideProgress, 700);
  }, 50);
}

// ---------------- Dropzone wiring ----------------
function wireDropzone(id, inputId, onFile){
  const dz = document.getElementById(id);
  const fileInput = document.getElementById(inputId);
  if(!dz || !fileInput) return;

  const openPicker = () => fileInput.click();
  dz.addEventListener('click', openPicker);
  dz.addEventListener('keydown', e=>{ if(e.key==='Enter'||e.key===' ') openPicker(); });

  ['dragenter','dragover'].forEach(evt=> dz.addEventListener(evt, e=>{
    e.preventDefault(); e.stopPropagation(); dz.classList.add('drag');
  }));
  ['dragleave','drop'].forEach(evt=> dz.addEventListener(evt, e=>{
    e.preventDefault(); e.stopPropagation(); dz.classList.remove('drag');
  }));
  dz.addEventListener('drop', e=>{
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if(f) onFile(f);
  });
  fileInput.addEventListener('change', e=>{
    const f = e.target.files && e.target.files[0];
    if(f) onFile(f);
  });
}

// ---------------- Init ----------------
document.addEventListener('DOMContentLoaded', ()=>{
  document.getElementById('year').textContent = new Date().getFullYear();
  document.getElementById('lang-toggle').addEventListener('click', ()=>{ currentLang = (currentLang==='en'?'zh':'en'); applyI18n(); });
  applyI18n();

  initNavTabs();
  setActivePanel('payments'); // optional: land on Payments automatically? comment out if not desired
  updateRunButton();

  document.getElementById('unpaid-mapping-toggle').addEventListener('click', ()=> document.getElementById('unpaid-info').classList.toggle('hidden'));
  document.getElementById('paid-mapping-toggle').addEventListener('click',  ()=> document.getElementById('paid-info').classList.toggle('hidden'));
  document.getElementById('run').addEventListener('click', runDetection);

  wireDropzone('dz-unpaid', 'file-unpaid', file=>{
    parseCSV(file, (rows, headers)=>{
      unpaidRows = rows; unpaidHeaders = headers;
      const {mapping, missing} = mapRecords(rows, headers, CONTRACTS.unpaid);
      unpaidMap = mapping; unpaidMissing = missing;
      unpaidRecords = extractRecords(rows, mapping, 'Unpaid');
      showMapping('unpaid-info', CONTRACTS.unpaid.name, mapping, missing, headers);
      const dzTitle = document.querySelector('#dz-unpaid .dz-title');
      if(dzTitle){ dzTitle.textContent = `${file.name} – ${rows.length} rows loaded`; }
      updateRunButton();
      console.log('UNPAID usable rows:', unpaidRecords.length, 'of', rows.length);
    });
  });

  wireDropzone('dz-paid', 'file-paid', file=>{
    parseCSV(file, (rows, headers)=>{
      paidRows = rows; paidHeaders = headers;
      const {mapping, missing} = mapRecords(rows, headers, CONTRACTS.paid);
      paidMap = mapping; paidMissing = missing;
      paidRecords = extractRecords(rows, mapping, 'Paid');
      showMapping('paid-info', CONTRACTS.paid.name, mapping, missing, headers);
      const dzTitle = document.querySelector('#dz-paid .dz-title');
      if(dzTitle){ dzTitle.textContent = `${file.name} – ${rows.length} rows loaded`; }
      updateRunButton();
      console.log('PAID usable rows:', paidRecords.length, 'of', rows.length);
    });
  });
});