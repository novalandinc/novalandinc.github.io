/* =========================================================
   app.js — Smart exports; dynamic labels; bolded modal targets; Desc off by default
   ========================================================= */

(function(){
  const $  = (sel, root=document) => root.querySelector(sel);
  const $$ = (sel, root=document) => Array.from(root.querySelectorAll(sel));
  const on = (el, evt, fn) => el && el.addEventListener(evt, fn);
  const html = String.raw;

  /* ---------- utilities ---------- */
  const DATE_HEADER_CANDS = [
    'date','post date','posting date','transaction date','posted date','value date','book date','processing date'
  ];
  const CARD_HEADER_CANDS = [
    'card','card number','card #','card#','acct','account','account number','acct number','last 4','last4','card last 4','cardlast4'
  ];
  const DESC_HEADER_CANDS = [
    'description','merchant','memo','details','narrative','payee','counterparty'
  ];
  const AMOUNT_HEADER_CANDS = [
    'amount','amt','transaction amount','amount (usd)','debit','credit','value','payment','charge','withdrawal','deposit'
  ];

  const normalizeHeader = s => String(s||'').toLowerCase().replace(/[\s_\-/#.]+/g,' ').trim();
  function esc(s){
    return String(s ?? '')
      .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;").replace(/'/g,"&#39;");
  }
  function toDateAny(v){
    if (v instanceof Date) return v;
    if (typeof v === 'number') {
      if (v > 20000 && v < 80000) {
        const d = new Date(Date.UTC(1899, 11, 30));
        d.setUTCDate(d.getUTCDate() + Math.round(v));
        return d;
      }
      const d2 = new Date(v);
      return isNaN(d2) ? null : d2;
    }
    const s = String(v ?? '').trim();
    if (!s) return null;
    const d = new Date(s);
    if (!isNaN(d)) return d;
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (m){ let [_, mm, dd, yy] = m; yy = (+yy<100)?2000+(+yy):+yy; return new Date(+yy,(+mm)-1,+dd); }
    return null;
  }
  function daysDiff(a, b){
    const da = toDateAny(a), db = toDateAny(b);
    if (!da || !db) return Number.POSITIVE_INFINITY;
    return Math.abs((da - db) / (1000*60*60*24));
  }
  function withinDays(a, b, tol){ return daysDiff(a,b) <= (tol ?? 0); }

  function coerceNumber(v){
    if (v === null || v === undefined) return null;
    let s = String(v).trim();
    if (!s) return null;
    const unicodeMinus = /[\u2212\u2012\u2013\u2014\u2010]/g;
    s = s.replace(unicodeMinus, '-');
    const parenNeg = /^\((.*)\)$/;
    const isParen = parenNeg.test(s);
    if (isParen) s = s.replace(parenNeg, '$1');
    s = s.replace(/[\s,]/g,'').replace(/\$/g,'');
    if (s === '' || s === '-' ) return null;
    const n = Number(s);
    if (isNaN(n)) return null;
    return isParen ? -n : n;
  }
  const cents = n => n === null ? null : Math.round(Math.abs(n) * 100);

  const normalizeDesc = s => String(s ?? '').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim().replace(/\s+/g,' ');
  const normalizeDescTokens = s => normalizeDesc(s).split(' ').filter(Boolean).map(tok=>{
    if (/\d/.test(tok)){
      const digits = tok.replace(/\D+/g,'');
      const last = digits.slice(-6);
      return last || tok;
    }
    return tok;
  });

  function descMatches(a, b, fuzzy=false){
    const A = normalizeDesc(a), B = normalizeDesc(b);
    if (!fuzzy) return A === B;
    if (!A || !B) return A === B;
    if (A.includes(B) || B.includes(A)) return true;
    const ta = normalizeDescTokens(a);
    const tb = normalizeDescTokens(b);
    const setA = new Set(ta);
    const setB = new Set(tb);
    const inter = ta.filter(t => setB.has(t)).length;
    const union = new Set([...ta, ...tb]).size || 1;
    const jacc = inter / union;
    if (jacc >= 0.66) return true;
    const da = ta.filter(t=>/\d/.test(t)).join('');
    const db = tb.filter(t=>/\d/.test(t)).join('');
    if (da && db && da.slice(-6) === db.slice(-6)) return true;
    return false;
  }

  const cardKey = s => {
    const digits = String(s ?? '').replace(/\D+/g,'');
    if (!digits) return '';
    return digits.length > 4 ? digits.slice(-4) : digits;
  };

  function detectByHeader(headers, candidates){
    if (!headers || !headers.length) return null;
    const norm = headers.map(h => normalizeHeader(h));
    for (const cand of candidates){
      const idx = norm.indexOf(normalizeHeader(cand));
      if (idx !== -1) return headers[idx];
    }
    for (let i=0;i<headers.length;i++){
      const h = norm[i];
      if (candidates.some(c => h.includes(normalizeHeader(c)))) return headers[i];
    }
    return null;
  }

  function detectDateColumnsByData(rows, headers){
    if (!rows.length) return [];
    const scores = headers.map(h => {
      let ok = 0, nonEmpty=0;
      for (let i=0;i<rows.length;i++){
        const v = rows[i][h];
        if (v === undefined || v === null || String(v).trim()==='') continue;
        nonEmpty++;
        if (toDateAny(v)) ok++;
        if (nonEmpty >= 60) break;
      }
      return { h, ratio: nonEmpty ? ok/nonEmpty : 0, nonEmpty };
    });
    return scores.filter(s => s.nonEmpty>=5 && s.ratio >= 0.7).map(s => s.h);
  }

  /* ---------- modal helpers ---------- */
  function chooseColumnModal(headers, whichLabelHTML){
    return new Promise((resolve) => {
      const dlg = $('#column-modal');
      const contentLabel = whichLabelHTML || '';
      if (!dlg){
        const choice = window.prompt(`Choose a column for ${contentLabel.replace(/<[^>]*>/g,'')}:\n\n- ${headers.join('\n- ')}`);
        if (!choice) return resolve(null);
        const match = headers.find(h => h.toLowerCase() === choice.toLowerCase());
        return resolve(match || null);
      }
      $('#modal-title').textContent = 'Choose column';
      // allow bold for the exact target (Card, Date, etc.)
      $('#modal-which').innerHTML = contentLabel;
      const sel = $('#modal-select');
      sel.innerHTML = headers.map(h => `<option value="${esc(h)}">${esc(h)}</option>`).join('');
      dlg.returnValue = '';
      dlg.showModal();
      const onOk = () => { dlg.close('ok'); resolve(sel.value || null); cleanup(); };
      const onCancel = () => { dlg.close('cancel'); resolve(null); cleanup(); };
      function cleanup(){
        $('#modal-ok').removeEventListener('click', onOk);
        $('#modal-cancel').removeEventListener('click', onCancel);
      }
      $('#modal-ok').addEventListener('click', onOk);
      $('#modal-cancel').addEventListener('click', onCancel);
    });
  }
  function chooseSheetModal(sheets, whichFile='the Excel file'){
    return new Promise((resolve)=>{
      const choice = window.prompt(`Multiple sheets detected in ${whichFile}. Type the sheet name to use:\n\n- ${sheets.join('\n- ')}`);
      if (!choice) return resolve(null);
      const m = sheets.find(n => n.toLowerCase() === choice.toLowerCase());
      resolve(m || null);
    });
  }

  /* ---------- parsing CSV or Excel ---------- */
  function readFileAsArrayBuffer(file){
    return new Promise((resolve,reject)=>{
      const r = new FileReader();
      r.onerror = () => reject(r.error);
      r.onload = () => resolve(r.result);
      r.readAsArrayBuffer(file);
    });
  }
  async function parseFile(file){
    const name = (file?.name || '').toLowerCase();
    if (name.endsWith('.csv')){
      const results = await new Promise((resolve,reject)=>{
        if (!window.Papa) return reject(new Error('PapaParse not found'));
        Papa.parse(file, { header:true, dynamicTyping:false, skipEmptyLines:'greedy',
          complete: results => resolve(results), error: err => reject(err) });
      });
      return { rows: results?.data ?? [], headers: results?.meta?.fields ?? Object.keys((results?.data||[{}])[0]) };
    }
    if (!window.XLSX) throw new Error('XLSX library not found');
    const ab = await readFileAsArrayBuffer(file);
    const wb = XLSX.read(ab, { type: 'array' });
    let sheetName = wb.SheetNames[0];
    if (wb.SheetNames.length > 1){
      const pick = await chooseSheetModal(wb.SheetNames, file.name);
      if (pick) sheetName = pick;
    }
    const ws = wb.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
    const headers = Object.keys(json[0] || {});
    return { rows: json, headers };
  }

  function isExcelFile(file){
    const n = (file?.name || '').toLowerCase();
    return /\.xlsx?$/.test(n);
  }
  function stripExt(name='file.xlsx'){
    const i = name.lastIndexOf('.'); return i>0 ? name.slice(0,i) : name;
  }
  const stem = (name='File') => {
    const base = stripExt(name);
    return base.length > 20 ? base.slice(0,20) + '…' : base;
  };

  function downloadCSV(filename, rows){
    const csv = rows.map(r => r.map(v => `"${String(v??'').replace(/"/g,'""')}"`).join(',')).join('\n');
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
    document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(a.href); a.remove();}, 100);
  }

  // ExcelJS-based XLSX with highlighted rows
  async function exportHighlightedXlsx(filename, headers, rows, highlightIdxSet, sheetName='Data'){
    if (!window.ExcelJS){ alert('ExcelJS library not loaded.'); return; }
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(sheetName);
    // Header
    ws.addRow(headers);
    const header = ws.getRow(1);
    header.font = { bold: true };
    ws.views = [{ state: 'frozen', ySplit: 1 }];

    // Data rows
    rows.forEach((r, i) => {
      const vals = headers.map(h => r[h]);
      const row = ws.addRow(vals);
      if (highlightIdxSet && highlightIdxSet.has(i)){
        row.eachCell(cell => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE68A' } };
        });
      }
    });

    // Auto filter + widths
    const colCount = headers.length;
    if (colCount){
      const lastColLetter = (n => {
        let s=''; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s;
      })(colCount);
      ws.autoFilter = { from: 'A1', to: `${lastColLetter}1` };
      for (let c=1;c<=colCount;c++){
        let w = headers[c-1] ? String(headers[c-1]).length : 10;
        for (let r=2;r<=rows.length+1;r++){
          const v = ws.getRow(r).getCell(c).value;
          const l = v==null ? 0 : String(v).length;
          if (l+2 > w) w = Math.min(l+2, 60);
        }
        ws.getColumn(c).width = Math.max(10, Math.min(60, w));
      }
    }

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
    document.body.appendChild(a); a.click(); setTimeout(()=>{URL.revokeObjectURL(a.href); a.remove();}, 100);
  }

  /* ---------- dropzone UI ---------- */
  function setDropzoneUploaded(zoneEl, fileName, inputEl){
    if (!zoneEl) return;
    zoneEl.classList.add('success');
    zoneEl.innerHTML = `✅ Uploaded<br><span class="filename">${esc(fileName)}</span>
      <button class="upload-new" type="button">Upload New File</button>`;
    const btn = zoneEl.querySelector('.upload-new');
    if (btn && inputEl){
      btn.addEventListener('click', (e)=>{
        e.stopPropagation();
        inputEl.click();
      });
    }
  }

  function wireDropzone(zoneId, inputId){
    const dz = document.getElementById(zoneId);
    const input = document.getElementById(inputId);
    if (!dz || !input) return;
    ['dragenter','dragover'].forEach(evt => on(dz, evt, e => { e.preventDefault(); dz.classList.add('drag'); }));
    ['dragleave','drop'].forEach(evt => on(dz, evt, e => { e.preventDefault(); dz.classList.remove('drag'); }));
    on(dz, 'drop', (e) => { const f=e.dataTransfer?.files?.[0]; if (f){ input.files=e.dataTransfer.files; input.dispatchEvent(new Event('change',{bubbles:true})); } });
    on(dz, 'click', () => input.click());
    on(dz, 'keydown', (e)=>{ if (e.key==='Enter' || e.key===' ') { e.preventDefault(); input.click(); } });
    on(input, 'change', (e)=>{
      const f = e.target.files?.[0];
      if (f) setDropzoneUploaded(dz, f.name, input);
    });
  }

  /* ---------- navigation ---------- */
  function showPanel(id){
    ['panel-home','panel-duplicates','panel-compare'].forEach(pid => {
      const el = document.getElementById(pid);
      if (!el) return;
      if (pid === id) el.classList.remove('hidden'); else el.classList.add('hidden');
    });
  }
  function initNav(){
    const goDup = ()=>showPanel('panel-duplicates');
    const goCmp = ()=>showPanel('panel-compare');
    on($('#nav-duplicates'),'click',goDup);
    on($('#nav-compare'),'click',goCmp);
    on($('#nav-duplicates-card'),'click',goDup);
    on($('#nav-duplicates-card'),'keydown',e=>{ if(e.key==='Enter'||e.key===' ') {e.preventDefault();goDup();}});
    on($('#nav-compare-card'),'click',goCmp);
    on($('#nav-compare-card'),'keydown',e=>{ if(e.key==='Enter'||e.key===' ') {e.preventDefault();goCmp();}});
    on($('#back-home-1'),'click',()=>showPanel('panel-home'));
    on($('#back-home-2'),'click',()=>showPanel('panel-home'));
  }

  /* ---------- dynamic A/B naming ---------- */
  const nameState = {
    aStem: 'A',
    bStem: 'B',
    paidStem: 'A',
    unpaidStem: 'B'
  };
  function updateABLabels(aStem, bStem){
    nameState.aStem = aStem || 'A';
    nameState.bStem = bStem || 'B';
    // badges
    const ba = $('#badge-a-to-b'); if (ba) ba.textContent = `${nameState.aStem}▸${nameState.bStem}`;
    const bb = $('#badge-b-to-a'); if (bb) bb.textContent = `${nameState.bStem}▸${nameState.aStem}`;
    // export buttons
    const exA = $('#export-a-not-b'); if (exA) exA.textContent = `Export ${nameState.aStem}▸${nameState.bStem}`;
    const exB = $('#export-b-not-a'); if (exB) exB.textContent = `Export ${nameState.bStem}▸${nameState.aStem}`;
    // filter placeholders
    const fa = $('#filter-a-not-b'); if (fa) fa.placeholder = `Filter rows in ${nameState.aStem} not in ${nameState.bStem}`;
    const fb = $('#filter-b-not-a'); if (fb) fb.placeholder = `Filter rows in ${nameState.bStem} not in ${nameState.aStem}`;
    // dropzone text (compare)
    const dza = $('#dz-a'); if (dza && dza.classList.contains('success')===false) dza.textContent = `File ${nameState.aStem} (CSV/XLSX/XLS)`;
    const dzb = $('#dz-b'); if (dzb && dzb.classList.contains('success')===false) dzb.textContent = `File ${nameState.bStem} (CSV/XLSX/XLS)`;
  }
  function updateOverlapLabel(paidStem, unpaidStem){
    nameState.paidStem = paidStem || 'A';
    nameState.unpaidStem = unpaidStem || 'B';
    const ex = $('#export-overlap'); if (ex) ex.textContent = `Export Overlap ${nameState.paidStem}∩${nameState.unpaidStem}`;
  }
  const makeWhichLabel = (stemName, exact) => `${esc(stemName)}: <b>${esc(exact)}</b>`;

  /* =========================================================
     Duplicates Auditor (+ export)
     ========================================================= */
  function renderTable(container, headers, rows, title){
    if (!container) return;
    if (!headers || !headers.length) headers = Object.keys(rows[0] || {});
    container.innerHTML = html`<div class="card">
      <div class="card-title">${esc(title || 'Results')}</div>
      <div class="scroll-x">
        <table class="table compact">
          <thead><tr><th class="num">#</th>${headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead>
          <tbody>
            ${rows.map((r,i)=>`<tr><td class="num">${i+2}</td>${headers.map(h=>`<td>${esc(r[h])}</td>`).join('')}</tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
  }

  const singleState = { file: null, rows: [], headers: [], dupIdxs: new Set() };
  async function runSingle(){
    const status = $('#status-single');
    const file = $('#file-single')?.files?.[0];
    const days = Math.max(0, +($('#days-single')?.value||0));
    if (!file){ if(status) status.textContent='Please choose a file.'; return; }
    if (status) status.textContent='Parsing…';
    const parsed = await parseFile(file);
    const rows = parsed.rows; const headers = parsed.headers;
    let dateCol = detectByHeader(headers, DATE_HEADER_CANDS);
    let txnCol  = detectByHeader(headers, ['transaction number','transaction id','trans #','txn #','reference','ref','id','trace number','check #','check number']);
    const dateCands = [...new Set([...(detectDateColumnsByData(rows, headers)), ...headers.filter(h=>DATE_HEADER_CANDS.some(c=>normalizeHeader(h).includes(normalizeHeader(c))))])];
    if (dateCands.length >= 2) dateCol = await chooseColumnModal(dateCands, makeWhichLabel(stem(file.name),'Date'));
    if (!dateCol)  dateCol = await chooseColumnModal(headers, makeWhichLabel(stem(file.name),'Date'));
    if (!txnCol)   txnCol  = await chooseColumnModal(headers, makeWhichLabel(stem(file.name),'Transaction #'));
    if (!dateCol || !txnCol){ if(status) status.textContent='Required columns not set.'; return; }
    const seen = new Map();
    const dups = [];
    const dupIdxs = new Set();
    rows.forEach((r, i)=>{
      const k = String(r[txnCol]??'').trim();
      const d = r[dateCol];
      if (!seen.has(k)) seen.set(k, []);
      const arr = seen.get(k);
      const hit = arr.find(x => withinDays(x.d, d, days));
      if (hit){ dups.push(r); dupIdxs.add(i); } else { arr.push({d, idx:i}); }
    });
    singleState.file = file; singleState.rows = rows; singleState.headers = headers; singleState.dupIdxs = dupIdxs;

    if (status) status.textContent = `Found ${dups.length} potential duplicates.`;
    renderTable($('#out-single'), headers, dups, 'Duplicate rows');

    const btn = $('#export-dups');
    if (btn){
      btn.disabled = dups.length === 0;
      btn.onclick = async () => {
        if (isExcelFile(singleState.file)){
          const fn = `${stripExt(singleState.file.name)}__DUPLICATE_ROWS_HIGHLIGHTED.xlsx`;
          await exportHighlightedXlsx(fn, headers, rows, dupIdxs, 'Data');
        } else {
          if (!dups.length) return;
          const csvRows = [headers, ...rows.filter((_,i)=>dupIdxs.has(i)).map(r=>headers.map(h=>r[h]))];
          downloadCSV(`${stripExt(singleState.file.name)}__duplicates_only.csv`, csvRows);
        }
      };
    }
  }

  const crossState = { fileA: null, rowsA: [], headersA: [], overlapIdxs: new Set(), paidName:'', unpaidName:'' };
  async function runCross(){
    const status = $('#status-cross');
    const fa = $('#file-paid')?.files?.[0];
    const fb = $('#file-unpaid')?.files?.[0];
    const days = Math.max(0, +($('#days-cross')?.value||0));
    if (!fa || !fb){ if(status) status.textContent='Choose both files.'; return; }
    if (status) status.textContent='Parsing…';

    updateOverlapLabel(stem(fa.name), stem(fb.name));

    const [pa, pb] = await Promise.all([parseFile(fa), parseFile(fb)]);
    const rowsA = pa.rows; const rowsB = pb.rows;
    const headersA = pa.headers; const headersB = pb.headers;
    let dateA = detectByHeader(headersA, DATE_HEADER_CANDS);
    let dateB = detectByHeader(headersB, DATE_HEADER_CANDS);
    let txnA  = detectByHeader(headersA, ['transaction number','transaction id','trans #','txn #','reference','ref','id','trace number','check #','check number']);
    let txnB  = detectByHeader(headersB, ['transaction number','transaction id','trans #','txn #','reference','ref','id','trace number','check #','check number']);
    const candA = [...new Set([...(detectDateColumnsByData(rowsA, headersA)), ...headersA.filter(h=>DATE_HEADER_CANDS.some(c=>normalizeHeader(h).includes(normalizeHeader(c))))])];
    const candB = [...new Set([...(detectDateColumnsByData(rowsB, headersB)), ...headersB.filter(h=>DATE_HEADER_CANDS.some(c=>normalizeHeader(h).includes(normalizeHeader(c))))])];
    if (candA.length >= 2) dateA = await chooseColumnModal(candA, makeWhichLabel(stem(fa.name),'Date'));
    if (candB.length >= 2) dateB = await chooseColumnModal(candB, makeWhichLabel(stem(fb.name),'Date'));
    if (!dateA) dateA = await chooseColumnModal(headersA, makeWhichLabel(stem(fa.name),'Date'));
    if (!dateB) dateB = await chooseColumnModal(headersB, makeWhichLabel(stem(fb.name),'Date'));
    if (!txnA)  txnA  = await chooseColumnModal(headersA, makeWhichLabel(stem(fa.name),'Transaction #'));
    if (!txnB)  txnB  = await chooseColumnModal(headersB, makeWhichLabel(stem(fb.name),'Transaction #'));
    if (!dateA || !dateB || !txnA || !txnB){ if(status) status.textContent='Required columns not set.'; return; }

    const byTxnB = new Map();
    rowsB.forEach(r=>{
      const k = String(r[txnB]??'').trim();
      if (!byTxnB.has(k)) byTxnB.set(k, []);
      byTxnB.get(k).push(r);
    });
    const overlap = [];
    const overlapIdxs = new Set();
    rowsA.forEach((r, i)=>{
      const k = String(r[txnA]??'').trim();
      const cands = byTxnB.get(k)||[];
      if (cands.find(b=>withinDays(r[dateA], b[dateB], days))){ overlap.push(r); overlapIdxs.add(i); }
    });
    crossState.fileA = fa; crossState.rowsA = rowsA; crossState.headersA = headersA; crossState.overlapIdxs = overlapIdxs;
    crossState.paidName = fa.name; crossState.unpaidName = fb.name;

    if (status) status.textContent = `Found ${overlap.length} overlaps.`;
    renderTable($('#out-cross'), headersA, overlap, `Rows in ${stem(fa.name)} that overlap ${stem(fb.name)}`);

    const btn = $('#export-overlap');
    if (btn){
      btn.disabled = overlap.length === 0;
      btn.onclick = async () => {
        if (isExcelFile(crossState.fileA)){
          const fn = `${stripExt(crossState.fileA.name)}__OVERLAP_ROWS_HIGHLIGHTED.xlsx`;
          await exportHighlightedXlsx(fn, headersA, rowsA, overlapIdxs, 'Data');
        } else {
          if (!overlap.length) return;
          const csvRows = [headersA, ...rowsA.filter((_,i)=>overlapIdxs.has(i)).map(r=>headersA.map(h=>r[h]))];
          downloadCSV(`${stripExt(crossState.fileA.name)}__overlap_only.csv`, csvRows);
        }
      };
    }
  }

  /* =========================================================
     Compare (+ dynamic names + exports)
     ========================================================= */
  const compareState = {
    onlyA: [], headersA: [], fileAName: '', rowsA: [], fileA: null,
    onlyB: [], headersB: [], fileBName: '', rowsB: [], fileB: null,
    a: { page: 1, size: 25 },
    b: { page: 1, size: 25 }
  };

  function renderGridPaged(container, headers, items, page, size){
    if (!container) return;
    if (!headers || !headers.length) headers = Object.keys(items?.[0]?.r || {});
    const total = items.length;
    const maxPage = Math.max(1, Math.ceil(total / size));
    const clampedPage = Math.min(Math.max(1, page), maxPage);
    const start = (clampedPage-1)*size;
    const end = Math.min(start+size, total);
    const slice = items.slice(start, end);
    const head = html`<thead><tr><th class="num">#</th>${headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead>`;
    const body = html`<tbody>${slice.map(({r,idx})=>`<tr><td class="num">${idx+2}</td>${headers.map(h=>`<td>${esc(r[h])}</td>`).join('')}</tr>`).join('')}</tbody>`;
    container.innerHTML = html`<div class="scroll-x"><table class="table compact">${head}${body}</table></div>
      <div class="muted" style="padding:8px 4px;">Showing ${total? start+1:0}–${end} of ${total}</div>`;
    return { page: clampedPage, maxPage, total, start, end };
  }
  function syncPagination(which){
    if (which==='a'){
      const info = renderGridPaged($('#grid-a-not-b'), compareState.headersA, compareState.onlyA, compareState.a.page, compareState.a.size);
      $('#a-pageinfo').textContent = `Page ${info.page} / ${info.maxPage}`;
      $('#total-a-not-b').textContent = `Total: ${info.total}`;
    } else {
      const info = renderGridPaged($('#grid-b-not-a'), compareState.headersB, compareState.onlyB, compareState.b.page, compareState.b.size);
      $('#b-pageinfo').textContent = `Page ${info.page} / ${info.maxPage}`;
      $('#total-b-not-a').textContent = `Total: ${info.total}`;
    }
  }
  function reflectMapping(which, cols){
    const mapEl = which === 'A' ? $('#map-a') : $('#map-b');
    if (!mapEl) return;
    const { date, amount, desc, card } = cols;
    const parts = [];
    if (date) parts.push(`Date: <code>${esc(date)}</code>`);
    if (amount) parts.push(`Amount: <code>${esc(amount)}</code>`);
    if (desc) parts.push(`Description: <code>${esc(desc)}</code>`);
    if (card) parts.push(`Card: <code>${esc(card)}</code>`);
    mapEl.innerHTML = parts.join(' • ');
  }
  function wireFilter(input, which){
    const grid = which==='a' ? $('#grid-a-not-b') : $('#grid-b-not-a');
    on(input,'input',()=>{
      const q = input.value.toLowerCase();
      $$('tbody tr', grid).forEach(tr => {
        const t = tr.textContent.toLowerCase();
        tr.style.display = t.includes(q) ? '' : 'none';
      });
    });
  }

  async function runCompare(){
    const status = $('#compare-readiness');
    const fileA = $('#file-a')?.files?.[0];
    const fileB = $('#file-b')?.files?.[0];
    const days = Math.max(0, +($('#compare-days')?.value||0));
    const useDesc = $('#use-desc')?.checked;
    const fuzzyDesc = $('#fuzzy-desc')?.checked;
    if (!fileA || !fileB){ if(status) status.textContent='Please select File A and File B.'; return; }
    if (status) status.textContent='Parsing…';

    const aStem = stem(fileA.name);
    const bStem = stem(fileB.name);
    updateABLabels(aStem, bStem);

    const [pa, pb] = await Promise.all([parseFile(fileA), parseFile(fileB)]);
    const rowsA = pa.rows; const rowsB = pb.rows;
    const headersA = pa.headers; const headersB = pb.headers;
    if (!rowsA.length || !rowsB.length){ if(status) status.textContent='One or both files are empty.'; return; }

    // Detect columns and force prompt if multiple date candidates (with bolded exact field and name stems)
    let dateA = detectByHeader(headersA, DATE_HEADER_CANDS);
    let dateB = detectByHeader(headersB, DATE_HEADER_CANDS);
    const dateCandsA = [...new Set([...(detectDateColumnsByData(rowsA, headersA)), ...headersA.filter(h=>DATE_HEADER_CANDS.some(c=>normalizeHeader(h).includes(normalizeHeader(c))))])];
    const dateCandsB = [...new Set([...(detectDateColumnsByData(rowsB, headersB)), ...headersB.filter(h=>DATE_HEADER_CANDS.some(c=>normalizeHeader(h).includes(normalizeHeader(c))))])];
    if (dateCandsA.length >= 2) dateA = await chooseColumnModal(dateCandsA, makeWhichLabel(aStem,'Date'));
    if (dateCandsB.length >= 2) dateB = await chooseColumnModal(dateCandsB, makeWhichLabel(bStem,'Date'));
    if (!dateA) dateA = await chooseColumnModal(headersA, makeWhichLabel(aStem,'Date'));
    if (!dateB) dateB = await chooseColumnModal(headersB, makeWhichLabel(bStem,'Date'));

    let amountA = detectByHeader(headersA, AMOUNT_HEADER_CANDS);
    let amountB = detectByHeader(headersB, AMOUNT_HEADER_CANDS);
    if (!amountA) amountA = await chooseColumnModal(headersA, makeWhichLabel(aStem,'Amount'));
    if (!amountB) amountB = await chooseColumnModal(headersB, makeWhichLabel(bStem,'Amount'));

    let descA = useDesc ? (detectByHeader(headersA, DESC_HEADER_CANDS) || await chooseColumnModal(headersA, makeWhichLabel(aStem,'Description'))) : null;
    let descB = useDesc ? (detectByHeader(headersB, DESC_HEADER_CANDS) || await chooseColumnModal(headersB, makeWhichLabel(bStem,'Description'))) : null;

    let cardA = detectByHeader(headersA, CARD_HEADER_CANDS) || await chooseColumnModal(headersA, makeWhichLabel(aStem,'Card'));
    let cardB = detectByHeader(headersB, CARD_HEADER_CANDS) || await chooseColumnModal(headersB, makeWhichLabel(bStem,'Card'));

    if (!dateA || !dateB || !amountA || !amountB || !cardA || !cardB){
      if (status) status.textContent = 'Missing required columns (Date, Amount, Card).';
      return;
    }
    if (useDesc && (!descA || !descB)){
      if (status) status.textContent = 'Description matching enabled but Description columns not set.';
      return;
    }

    reflectMapping('A', {date:dateA, amount:amountA, desc:descA, card:cardA});
    reflectMapping('B', {date:dateB, amount:amountB, desc:descB, card:cardB});

    // Info banner for opposite signs (still matching on |amount|)
    let evaluated=0, opposite=0;
    rowsA.slice(0,2000).forEach(ra=>{
      const a = coerceNumber(ra[amountA]);
      if (a===null) return;
      const match = rowsB.find(rb => cardKey(rb[cardB])===cardKey(ra[cardA])
        && (!useDesc || descMatches(rb[descB], ra[descA], fuzzyDesc))
        && cents(coerceNumber(rb[amountB]))===cents(a));
      if (match){
        evaluated++;
        const b = coerceNumber(match[amountB]);
        if (b!==null && Math.sign(a) === -Math.sign(b)) opposite++;
      }
    });
    const banner = $('#sign-flip-banner');
    if (evaluated>0 && (opposite/evaluated) >= 0.7){
      if (banner){ banner.classList.remove('hidden'); banner.textContent=`Most matched amounts have opposite signs. Comparison uses |amount| for matching.`; }
    } else if (banner){ banner.classList.add('hidden'); banner.textContent=''; }

    // Build B index key: Card-last4 + (optional Desc) + |Amount| cents
    const idxB = new Map();
    rowsB.forEach((rb, i)=>{
      const keyParts = [cardKey(rb[cardB])];
      if (useDesc) keyParts.push(fuzzyDesc ? normalizeDescTokens(rb[descB]).join(' ') : normalizeDesc(rb[descB]));
      keyParts.push(String(cents(coerceNumber(rb[amountB]))));
      const key = keyParts.join('|');
      if (!idxB.has(key)) idxB.set(key, []);
      idxB.get(key).push({ row: rb, idx: i });
    });

    // 1-1 pairing with closest date within tolerance
    const usedB = new Set();
    const matchedA = new Set();
    const matchedB = new Set();

    rowsA.forEach((ra, i)=>{
      const keyParts = [cardKey(ra[cardA])];
      if (useDesc) keyParts.push(fuzzyDesc ? normalizeDescTokens(ra[descA]).join(' ') : normalizeDesc(ra[descA]));
      keyParts.push(String(cents(coerceNumber(ra[amountA]))));
      const key = keyParts.join('|');

      const cands = (idxB.get(key) || []).filter(c => !usedB.has(c.idx));
      if (!cands.length) return;
      let best = null, bestDiff = Infinity;
      for (const c of cands){
        const diff = daysDiff(ra[dateA], c.row[dateB]);
        if (diff <= days && diff < bestDiff){
          best = c; bestDiff = diff;
        }
      }
      if (best){
        matchedA.add(i);
        matchedB.add(best.idx);
        usedB.add(best.idx);
      }
    });

    const onlyA = rowsA.map((r,i)=>({r,idx:i})).filter(x=>!matchedA.has(x.idx));
    const onlyB = rowsB.map((r,i)=>({r,idx:i})).filter(x=>!matchedB.has(x.idx));

    // store state
    compareState.onlyA = onlyA;
    compareState.onlyB = onlyB;
    compareState.headersA = headersA;
    compareState.headersB = headersB;
    compareState.fileAName = fileA.name;
    compareState.fileBName = fileB.name;
    compareState.rowsA = rowsA;
    compareState.rowsB = rowsB;
    compareState.fileA = fileA;
    compareState.fileB = fileB;
    compareState.a.page = 1; compareState.b.page = 1;

    // exports
    const idxAset = new Set(onlyA.map(x=>x.idx));
    const idxBset = new Set(onlyB.map(x=>x.idx));

    const btnA = $('#export-a-not-b');
    if (btnA){
      btnA.onclick = async () => {
        if (!onlyA.length) return;
        if (isExcelFile(compareState.fileA)){
          const fn = `${stripExt(compareState.fileAName)}__MISMATCH_ROWS_HIGHLIGHTED.xlsx`;
          await exportHighlightedXlsx(fn, headersA, rowsA, idxAset, 'Data');
        } else {
          const headers = Object.keys(onlyA[0].r||{});
          const csvRows = [headers, ...onlyA.map(({r})=>headers.map(h=>r[h]))];
          downloadCSV(`rows_in_${compareState.fileAName}_not_in_${compareState.fileBName}.csv`, csvRows);
        }
      };
    }
    const btnB = $('#export-b-not-a');
    if (btnB){
      btnB.onclick = async () => {
        if (!onlyB.length) return;
        if (isExcelFile(compareState.fileB)){
          const fn = `${stripExt(compareState.fileBName)}__MISMATCH_ROWS_HIGHLIGHTED.xlsx`;
          await exportHighlightedXlsx(fn, headersB, rowsB, idxBset, 'Data');
        } else {
          const headers = Object.keys(onlyB[0].r||{});
          const csvRows = [headers, ...onlyB.map(({r})=>headers.map(h=>r[h]))];
          downloadCSV(`rows_in_${compareState.fileBName}_not_in_${compareState.fileAName}.csv`, csvRows);
        }
      };
    }

    wireFilter($('#filter-a-not-b'), 'a');
    wireFilter($('#filter-b-not-a'), 'b');

    // totals + initial render
    $('#total-a-not-b').textContent = `Total: ${onlyA.length}`;
    $('#total-b-not-a').textContent = `Total: ${onlyB.length}`;
    $('#a-pagesize').value = String(compareState.a.size);
    $('#b-pagesize').value = String(compareState.b.size);
    syncPagination('a');
    syncPagination('b');

    if (status) status.textContent = `Done. ${onlyA.length} ${aStem}→${bStem} misses; ${onlyB.length} ${bStem}→${aStem} misses.`;
  }

  /* ---------- boot ---------- */
  function boot(){
    initNav();
    ['single','paid','unpaid','a','b'].forEach(tag => wireDropzone(`dz-${tag}`, `file-${tag}`));
    on($('#run-single'),'click',()=>{ runSingle().catch(e=>{$('#status-single').textContent='Error: '+e.message;}); });
    on($('#run-cross'),'click',()=>{ runCross().catch(e=>{$('#status-cross').textContent='Error: '+e.message;}); });
    on($('#run-compare'),'click',()=>{ runCompare().catch(e=>{$('#compare-readiness').textContent='Error: '+e.message;}); });
    on($('#file-a'),'change',e=>{ const f=e.target.files?.[0]; if(f){ updateABLabels(stem(f.name), nameState.bStem); $('#compare-readiness').textContent = `A: ${f.name}`; } });
    on($('#file-b'),'change',e=>{ const f=e.target.files?.[0]; if(f){ updateABLabels(nameState.aStem, stem(f.name)); const t=$('#compare-readiness'); t.textContent = (t.textContent? t.textContent+' • ' : '') + `B: ${f.name}`; } });
    on($('#file-paid'),'change',e=>{ const f=e.target.files?.[0]; if(f){ updateOverlapLabel(stem(f.name), nameState.unpaidStem); } });
    on($('#file-unpaid'),'change',e=>{ const f=e.target.files?.[0]; if(f){ updateOverlapLabel(nameState.paidStem, stem(f.name)); } });
    const yearEl=$('#year'); if(yearEl) yearEl.textContent=new Date().getFullYear();
    on($('#a-prev'),'click',()=>{ compareState.a.page--; syncPagination('a'); });
    on($('#a-next'),'click',()=>{ compareState.a.page++; syncPagination('a'); });
    on($('#b-prev'),'click',()=>{ compareState.b.page--; syncPagination('b'); });
    on($('#b-next'),'click',()=>{ compareState.b.page++; syncPagination('b'); });
    on($('#a-pagesize'),'change',e=>{ compareState.a.size=+e.target.value; compareState.a.page=1; syncPagination('a'); });
    on($('#b-pagesize'),'change',e=>{ compareState.b.size=+e.target.value; compareState.b.page=1; syncPagination('b'); });
  }
  if (document.readyState === 'loading'){ document.addEventListener('DOMContentLoaded', boot); } else { boot(); }
})();