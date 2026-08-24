
let files=[],masterRows=[],masterHeaders=[],results=[],aliases={},finalReady=false,masterLoaded=false,identityMeta=[];
const $=id=>document.getElementById(id); const tabs=document.querySelectorAll('.tab');
function toast(s){$('toast').textContent=s;$('toast').style.display='block';setTimeout(()=>$('toast').style.display='none',5000)}
function tab(n){tabs.forEach(b=>b.classList.toggle('active',b.dataset.tab===n));['filesTab','reviewTab','outputTab','summaryTab'].forEach(x=>$(x).classList.toggle('hidden',x!==n))}
tabs.forEach(b=>b.onclick=()=>tab(b.dataset.tab));
function progress(p,stage,detail=''){$('bar').style.width=p+'%';$('pct').textContent=Math.round(p)+'%';$('status').textContent=stage;$('stage').textContent=stage;$('detail').textContent=detail}
function updateCount(){$('fileCount').textContent=`${files.length} File${files.length===1?'':'s'} Uploaded`;$('dropHint').textContent=files.length?files.map(f=>f.name).join(' • '):'No files selected';$('processBtn').disabled=!files.length||!masterLoaded;$('fileKpi').textContent=`${validatedCount()} / ${files.length} validated`}
function validatedCount(){return files.filter((f,i)=>{const h=$('hq-'+i),s=$('cust-'+i);return h&&h.value&&s&&s.value}).length}
function renderFileRows(){ $('filesBody').innerHTML=files.map((f,i)=>`<tr><td>${i+1}</td><td><b>${esc(f.name)}</b></td><td><select id="hq-${i}"><option value="">Detecting HQ…</option></select></td><td><select id="cust-${i}"><option value="">Detecting stockist…</option></select></td><td id="conf-${i}">—</td><td id="rows-${i}">—</td><td id="sec-${i}">—</td><td id="clo-${i}">—</td><td id="st-${i}"><span class="pill info">Reading PDF header…</span></td></tr>`).join('')}
function esc(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function normCustomer(s){s=String(s||'').toUpperCase().replace(/\.[ ]*/g,' ');s=s.replace(/[-_/]+/g,' ');s=s.replace(/\b(M\/S|MS|MESSRS|AKOLA|NAGPUR|CHA|HEADQUARTER|HQ|CORPORATION|CORP|COMPANY|DISTRIBUTOR|DISTRIBUTORS|STOCKIST|AGENCIES|AGENCY|PHARMA|PHARMACEUTICALS)\b/g,' ');s=s.replace(/\b([A-Z])\s+(?=[A-Z](?:\s|$))/g,'$1');return s.replace(/[^A-Z0-9 ]+/g,' ').replace(/\s+/g,' ').trim()}
function ratio(a,b){a=normCustomer(a);b=normCustomer(b);if(!a||!b)return 0;if(a===b)return 100;if(a.includes(b)||b.includes(a))return 96;const A=a.split(' '),B=b.split(' ');let common=0;A.forEach(x=>{if(B.includes(x))common++});return Math.min(99,Math.round((common/Math.max(A.length,B.length))*100))}
function skuNorm(s){s=String(s||'').toUpperCase().replace(/’/g,"'").replace(/CILNIREM/g,'CILNIKEM').replace(/(?<=\d)SMG\b/g,'5MG');s=s.replace(/\b(TABLETS?|TABS?|CAPSULES?|CAPS?|ORAL|ER)\b/g,' ');s=s.replace(/\b\d+(?:X1|X|TAB|TABS|,S|'S|S)\b/g,' ');return s.replace(/[^A-Z0-9./+]+/g,'')}
function dedupeMaster(rows){
  const seen=new Set(),out=[];
  for(const r of (rows||[])){
    const h=String(r.HQ||'').trim();
    const c=String(r['CUSTOMER NAME']||'').trim();
    const s=String(r['SKU NAME']||'').trim();
    if(!h||!c||!s) continue;
    const key=skuNorm(h)+'|'+skuNorm(c)+'|'+skuNorm(s);
    if(seen.has(key)) continue;
    seen.add(key); out.push(r);
  }
  return out;
}
function loadMasterBuffer(buf){const wb=XLSX.read(buf,{type:'array'});const ws=wb.Sheets[wb.SheetNames[0]];const data=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});masterHeaders=data[0].map(x=>String(x||'').trim());const required=['HQ','CUSTOMER NAME','SKU NAME','PTS'];if(!required.every(h=>masterHeaders.includes(h)))throw Error('Master must contain HQ, CUSTOMER NAME, SKU NAME and PTS columns');masterRows=data.slice(1).map(r=>{const o={};masterHeaders.forEach((h,i)=>o[h]=r[i]??'');o.PTS=Number(String(o.PTS).replace(/,/g,''))||0;return o}).filter(r=>r.HQ&&r['CUSTOMER NAME']&&r['SKU NAME']);
  masterRows=dedupeMaster(masterRows);masterLoaded=true;populateHQ();$('masterStatus').textContent=`Loaded ${masterRows.length.toLocaleString()} master rows`}
function loadBundledMaster(){try{if(!window.BUNDLED_MASTER_ROWS||!window.BUNDLED_MASTER_HEADERS)throw Error('Embedded master unavailable');masterHeaders=window.BUNDLED_MASTER_HEADERS;masterRows=dedupeMaster(window.BUNDLED_MASTER_ROWS.map(r=>({...r,PTS:Number(r.PTS)||0})));masterLoaded=true;populateHQ();$('masterStatus').textContent=`Loaded ${masterRows.length.toLocaleString()} bundled master rows`;updateCount();}catch(e){masterLoaded=false;$('masterStatus').textContent='Bundled master unavailable — upload your Excel';toast('Default master could not be loaded. Please upload your master Excel.')}}
function populateHQ(){const h=[...new Set(masterRows.map(r=>r.HQ))].sort();$('hq').innerHTML='<option value="">Auto / All HQ</option>'+h.map(x=>`<option>${esc(x)}</option>`).join('')}
function stockistsForHQ(hq){return [...new Set(masterRows.filter(r=>String(r.HQ)===String(hq)).map(r=>r['CUSTOMER NAME']))].sort()}
function hqOptions(selected=''){const hs=[...new Set(masterRows.map(r=>r.HQ))].filter(Boolean).sort();return '<option value="">Select HQ</option>'+hs.map(h=>`<option value="${esc(h)}" ${String(h)===String(selected)?'selected':''}>${esc(h)}</option>`).join('')}
function setStockistOptions(i,selected=''){const h=$('hq-'+i)?.value||'';const c=$('cust-'+i);if(!c)return;const list=stockistsForHQ(h);c.innerHTML='<option value="">Select stockist</option>'+list.map(x=>`<option value="${esc(x)}">${esc(x)}</option>`).join('');if(selected && list.includes(selected))c.value=selected}
async function identifyLocal(file){
  const ext=(file.name.split('.').pop()||'').toLowerCase();
  const supported=['pdf','xlsx','xlsm','xls','txt','html','htm'];
  if(!supported.includes(ext))
    return {customer:'',hq:'',confidence:0,match_status:'UNSUPPORTED',source:'Manual selection',stockists:[]};

  try{
    const r=await fetch('/api/identify?filename='+encodeURIComponent(file.name),{
      method:'POST',
      headers:{'Content-Type':'application/octet-stream'},
      body:file
    });
    const text=await r.text();
    if(!r.ok) throw Error(text||`API returned ${r.status}`);
    const d=JSON.parse(text);
    return {...d,stockists:stockistsForHQ(d.hq||'')};
  }catch(e){
    return {customer:'',hq:'',confidence:0,match_status:'IDENTIFICATION ERROR',source:'API error',error:e.message,stockists:[]};
  }
}

async function validateFiles(){
  if(!masterLoaded||!files.length)return;
  progress(3,'Reading statements','Reading HQ and Stockist from each statement…');
  renderFileRows();identityMeta=[];let done=0;
  for(let i=0;i<files.length;i++){
    const x=await identifyLocal(files[i]);identityMeta[i]=x;
    const h=$('hq-'+i),c=$('cust-'+i);
    h.innerHTML=hqOptions(x.hq||'');
    if(x.hq){h.value=x.hq;setStockistOptions(i,x.customer||'')}else{setStockistOptions(i,'')}
    h.onchange=()=>{setStockistOptions(i,'');$('conf-'+i).innerHTML='<span class="pill warn">Manual HQ</span>';$('st-'+i).innerHTML='<span class="pill warn">Select stockist</span>';updateCount()};
    c.onchange=()=>{if(c.value){$('st-'+i).innerHTML='<span class="pill ok">✓ Validated</span>'}else{$('st-'+i).innerHTML='<span class="pill warn">Select stockist</span>'}updateCount()};
    const conf=Number(x.confidence)||0;
    $('conf-'+i).innerHTML=`<span class="pill ${conf>=90?'ok':conf>=70?'warn':'bad'}">${conf.toFixed(0)}%</span>`;
    if(x.match_status==='EXACT'||x.match_status==='HIGH CONFIDENCE'){
      $('st-'+i).innerHTML='<span class="pill ok">✓ Identity matched</span>';
    }else if(x.hq){
      $('st-'+i).innerHTML='<span class="pill warn">HQ detected — verify stockist</span>';
    }else{
      $('st-'+i).innerHTML='<span class="pill warn">Manual HQ/Stockist required</span>';
    }
    done++;progress(5+done/files.length*20,'Statement identity validation',`${done}/${files.length} statements checked`);await new Promise(r=>setTimeout(r,0));
  }
  updateCount();
}
$('browse').onclick=()=>$('files').click();
$('files').onchange=e=>{const incoming=[...e.target.files].filter(f=>!files.some(x=>x.name===f.name&&x.size===f.size));files.push(...incoming);e.target.value='';updateCount();validateFiles()};
$('drop').ondragover=e=>{e.preventDefault();$('drop').classList.add('drag')};$('drop').ondragleave=()=>$('drop').classList.remove('drag');$('drop').ondrop=e=>{e.preventDefault();$('drop').classList.remove('drag');const incoming=[...e.dataTransfer.files].filter(f=>!files.some(x=>x.name===f.name&&x.size===f.size));files.push(...incoming);updateCount();validateFiles()};
$('master').onchange=async e=>{const f=e.target.files[0];if(!f)return;try{progress(1,'Loading master','Reading master Excel…');loadMasterBuffer(await f.arrayBuffer());if(files.length)await validateFiles();progress(100,'Master ready',`${masterRows.length.toLocaleString()} master rows loaded`)}catch(err){toast(err.message);progress(0,'Master error',err.message)}};
function assignments(){const a={};files.forEach((f,i)=>a[f.name]=$('cust-'+i)?.value||'');return a}
function poolFor(customer,hq){
  const seen=new Set(),out=[];
  masterRows.filter(r=>r.HQ===hq&&r['CUSTOMER NAME']===customer).forEach(r=>{
    const key=skuNorm(r['SKU NAME']);
    if(seen.has(key)) return;
    seen.add(key);
    out.push({HQ:r.HQ,CUSTOMER_NAME:r['CUSTOMER NAME'],SKU_NAME:r['SKU NAME'],PTS:r.PTS});
  });
  return out;
}
async function analyzeFile(i,aliasesForRun){const f=files[i],customer=$('cust-'+i).value,hq=$('hq-'+i).value;let last='Unknown API error';for(let attempt=1;attempt<=2;attempt++){const ctl=new AbortController();const timer=setTimeout(()=>ctl.abort(),45000);try{const r=await fetch('/api/analyze?filename='+encodeURIComponent(f.name)+'&hq='+encodeURIComponent(hq)+'&customer='+encodeURIComponent(customer)+'&threshold='+encodeURIComponent($('threshold').value),{method:'POST',headers:{'Content-Type':'application/octet-stream','X-Aliases':JSON.stringify(aliasesForRun||{})},body:f,signal:ctl.signal});clearTimeout(timer);const text=await r.text();if(!r.ok)throw Error(text||`API returned ${r.status}`);const data=JSON.parse(text);return {...data,hq,customer}}catch(e){clearTimeout(timer);last=e.name==='AbortError'?'API timed out after 45 seconds':(e?.message||String(e));if(attempt<2)await new Promise(r=>setTimeout(r,600))}}throw Error(last)}
async function processStatements(aliasesForRun={}){if(!files.length)return toast('Upload statements first');if(validatedCount()!==files.length)return toast('Please validate/select a stockist for every file.');$('processBtn').disabled=true;$('applyAll').disabled=true;results=[];progress(25,'Reading statements','Processing statements individually for faster Vercel requests…');let done=0;const queue=[...files.keys()];const concurrency=Math.min(4,queue.length);async function worker(){while(queue.length){const i=queue.shift();$('st-'+i).innerHTML='<span class="pill info">Reading…</span>';try{const res=await analyzeFile(i,aliasesForRun);results[i]=res;$('rows-'+i).textContent=res.rows;$('sec-'+i).textContent=Math.round(res.sec_total);$('clo-'+i).textContent=Math.round(res.close_total);$('st-'+i).innerHTML=`<span class="pill ${res.reviews.length?'warn':'ok'}">${res.reviews.length?res.reviews.length+' Review':'Ready'}</span>`}catch(e){results[i]={file:files[i].name,hq:$('hq-'+i).textContent,customer:$('cust-'+i).value,rows:0,sec_total:0,close_total:0,matched:[],reviews:[],error:e.message};$('st-'+i).innerHTML='<span class="pill bad">Error</span>';toast(`${files[i].name}: ${e.message}`)}done++;progress(25+done/files.length*60,'Reading & matching statements',`${done}/${files.length} files complete`);}}
await Promise.all(Array.from({length:concurrency},worker));buildReview();$('processBtn').disabled=false;}
$('processBtn').onclick=()=>processStatements(aliases);
function buildReview(){
  /*
   * IMPORTANT:
   * Keep EVERY statement occurrence as a separate review row.
   *
   * A stockist statement can be arranged in product blocks/pages, e.g.
   * Voglikem -> Cilnikem -> ... -> Voglikem -> Cilnikem.
   * Repeated SKU occurrences must NOT be combined because SEC/CLO would
   * otherwise be double-counted.
   */
  const items=[];
  const reviewSeen=new Set();

  results.forEach((r,ri)=>{
    (r.reviews||[]).forEach(x=>{
      const source=String(x.source_sku||'').trim();
      if(!source) return;

      const key=skuNorm(r.hq)+'|'+skuNorm(r.customer)+'|'+skuNorm(source);
      if(reviewSeen.has(key)) return;
      reviewSeen.add(key);

      items.push({
        ri,r,x,
        auto: Number(x.confidence)>=82 && !!(x.master_sku||x.suggestion)
      });
    });
  });

  // Stable row identity. This lets two identical source SKUs be corrected
  // independently without changing the other occurrence.
  items.forEach((o,k)=>o.rowId='sku-row-'+k);

  $('oReview').textContent=items.length;

  if(!items.length){
    $('reviewBody').innerHTML=
      '<tr><td colspan="8"><span class="pill bad">No SKU rows were extracted. Check the statement format.</span></td></tr>';
    $('applyAll').disabled=true;
    tab('reviewTab');
    progress(90,'SKU validation ready','No SKU rows extracted.');
    return;
  }

  window.__reviewItems=items;

  $('reviewBody').innerHTML=items.map((o,k)=>{
    const pool=poolFor(o.r.customer,o.r.hq);
    const suggested=o.x.master_sku||o.x.suggestion||'';
    const confidence=Number(o.x.confidence)||0;
    const autoValidated=o.auto && confidence>=82 && !!suggested;
    const listId='sku-list-'+k;

    const options=pool.map(v=>
      `<option value="${esc(v.SKU_NAME)}"></option>`
    ).join('');

    const initialValue=autoValidated?suggested:'';
    const sourceKey=skuNorm(o.x.source_sku);

    return `<tr
      data-review-row="1"
      data-row-id="${o.rowId}"
      data-source-key="${esc(sourceKey)}"
      data-hq-key="${esc(skuNorm(o.r.hq))}"
      data-customer-key="${esc(skuNorm(o.r.customer))}"
      data-confidence="${confidence}"
      data-auto-validated="${autoValidated?'1':'0'}"
      data-ignored="0">

      <td>${esc(o.r.file)}</td>
      <td><b>${esc(o.x.source_sku)}</b></td>
      <td>${o.x.sec}</td>
      <td>${o.x.close}</td>
      <td>${esc(suggested||'—')}</td>
      <td><span class="pill ${confidence>=90?'ok':confidence>=70?'warn':'bad'}">${confidence}%</span></td>
      <td><span class="pill ${autoValidated?'ok':'warn'} review-status">
        ${autoValidated?'✓ AUTO MATCHED':'REVIEW'}
      </span></td>
      <td>
        <div style="display:flex;gap:6px;align-items:center;min-width:260px">
          <input
            class="reviewSelect"
            id="rv-${k}"
            list="${listId}"
            value="${esc(initialValue)}"
            placeholder="Type SKU name or IGN…"
            autocomplete="off"
            spellcheck="false"
            aria-label="Select Master SKU for ${esc(o.x.source_sku)}">
          <datalist id="${listId}">
            ${options}
            <option value="IGNORE — Not Our Product"></option>
          </datalist>
        </div>
        <div class="small" style="margin-top:4px">
          Type first letters to search • type <b>IGN</b> to Ignore
        </div>
      </td>
    </tr>`;
  }).join('');

  items.forEach((o,k)=>{
    const input=$('rv-'+k);
    if(!input) return;

    const row=input.closest('tr');
    const suggested=o.x.master_sku||o.x.suggestion||'';
    const autoValidated=o.auto && Number(o.x.confidence)>=82 && !!suggested;

    function refresh(){
      updateApplyState();
      setTimeout(()=>{
        if(typeof window.__refreshSkuActionScan==='function')
          window.__refreshSkuActionScan();
      },10);
    }

    function exactMaster(value){
      const v=String(value||'').trim();
      const pool=poolFor(o.r.customer,o.r.hq);

      return pool.find(m=>skuNorm(m.SKU_NAME)===skuNorm(v))?.SKU_NAME||'';
    }

    function applyRowValue(raw,fromUser=true){
      let value=String(raw||'').trim();

      // Fast Ignore shortcut.
      if(/^ign/i.test(value) || /^🚫?\s*ignore/i.test(value)){
        input.value='IGNORE — Not Our Product';
        row.dataset.ignored='1';
        row.dataset.autoValidated='0';
        row.dataset.userValidated='1';

        // Only THIS occurrence is ignored.
        o.x.__ignoredByUser=true;
        o.x.__reviewValue='';

        row.querySelector('.review-status').textContent='IGNORED';
        row.querySelector('.review-status').className='pill bad review-status';

        refresh();
        return;
      }

      if(!value){
        row.dataset.ignored='0';
        row.dataset.autoValidated=autoValidated?'1':'0';
        row.dataset.userValidated=autoValidated?'0':'0';
        o.x.__ignoredByUser=false;
        o.x.__reviewValue='';

        row.querySelector('.review-status').textContent=
          autoValidated?'✓ AUTO MATCHED':'REVIEW';
        row.querySelector('.review-status').className=
          'pill '+(autoValidated?'ok':'warn')+' review-status';

        refresh();
        return;
      }

      const exact=exactMaster(value);

      if(!exact){
        /*
         * Allow fast keyboard typing, but don't accept an SKU that isn't
         * actually in the detected HQ/stockist Master SKU pool.
         */
        if(fromUser){
          const q=normAction(value);
          const pool=poolFor(o.r.customer,o.r.hq);
          const matches=pool
            .filter(m=>normAction(m.SKU_NAME).includes(q))
            .slice(0,8);

          if(matches.length===1){
            input.value=matches[0].SKU_NAME;
            applyRowValue(matches[0].SKU_NAME,false);
            return;
          }

          toast(matches.length
            ? 'Select the required Master SKU from the suggestions.'
            : 'No matching Master SKU found for this stockist.');
        }

        return;
      }

      /*
       * CORRECTION IS ROW-SPECIFIC.
       * Do NOT write aliases[source SKU] here because two occurrences of
       * the same source SKU must be independently editable.
       */
      input.value=exact;
      row.dataset.ignored='0';
      row.dataset.autoValidated='0';
      row.dataset.userValidated='1';
      row.dataset.confidence='100';

      o.x.__ignoredByUser=false;
      o.x.__reviewValue=exact;
      o.x.confidence=100;
      o.x.master_sku=exact;

      row.querySelector('.review-status').textContent='✓ VALIDATED';
      row.querySelector('.review-status').className='pill ok review-status';

      refresh();
    }

    input.addEventListener('change',()=>applyRowValue(input.value,true));

    input.addEventListener('input',()=>{
      const value=input.value.trim();

      /*
       * IMPORTANT: do NOT run the duplicate/low filter while the user is
       * typing. In Duplicate / <85% views, deleting one character used to
       * make the row lose its flag and disappear before the user could finish
       * selecting the replacement SKU. The row is rescanned only after a
       * real selection/change.
       */
      if(/^ign$/i.test(value)){
        applyRowValue('IGNORE — Not Our Product',false);
      }

      updateApplyState();
    });

    if(autoValidated && suggested){
      input.value=suggested;
      o.x.__reviewValue=suggested;
    }
  });

  updateApplyState();
  tab('reviewTab');

  const manual=items.filter(o=>!(o.auto && Number(o.x.confidence)>=82 && !!(o.x.master_sku||o.x.suggestion))).length;

  progress(
    90,
    'SKU validation ready',
    `${items.length} statement SKU occurrences loaded • ${manual} need review • duplicate occurrences kept separate`
  );

  setTimeout(()=>{
    if(typeof window.__refreshSkuActionScan==='function')
      window.__refreshSkuActionScan();
  },50);
}

function isSkuResolved(row){
  if(!row) return true;
  if(row.dataset.ignored==='1') return true;

  // AUTO MATCHED rows >=85% are already accepted.
  if(row.dataset.autoValidated==='1') return true;

  // Any manual correction is accepted immediately.
  if(row.dataset.userValidated==='1') return true;

  const confidence=Number(row.dataset.confidence)||0;

  // >=85% with a suggested Master SKU is optional — no manual action.
  if(confidence>=85){
    const input=row.querySelector('.reviewSelect');
    if(input && input.value.trim()) return true;

    // Treat the API suggestion as already accepted even if the input
    // was rendered blank in a filtered/rebuilt view.
    const item=(window.__reviewItems||[]).find(
      o=>o.rowId===row.dataset.rowId
    );
    if(item && (item.x.master_sku||item.x.suggestion))
      return true;

    return true;
  }

  // Only <85% rows without a manual decision require action.
  const input=row.querySelector('.reviewSelect');
  return !!(input && input.value.trim());
}

function updateApplyState(){
  const rows=[...document.querySelectorAll('#reviewBody tr[data-review-row]')];

  const pending=rows.filter(row=>!isSkuResolved(row));

  $('applyAll').disabled=pending.length>0 || !results.length;
  $('applyAll').textContent=pending.length
    ? `Validate Remaining SKUs (${pending.length})`
    : 'Apply All Validated SKUs';
}

$('applyAll').onclick=async()=>{
  const rows=[...document.querySelectorAll('#reviewBody tr[data-review-row]')];
  const items=window.__reviewItems||[];

  const pending=rows.filter(row=>!isSkuResolved(row));

  if(pending.length){
    toast(`${pending.length} SKU(s) below 85% still need validation or Ignore.`);
    return;
  }

  progress(
    92,
    'Applying validated SKUs',
    'Applying corrections without combining statement quantities…'
  );

  $('applyAll').disabled=true;

  try{
    rows.forEach(row=>{
      const item=items.find(o=>o.rowId===row.dataset.rowId);
      if(!item) return;

      const input=row.querySelector('.reviewSelect');
      const value=String(input?.value||'').trim();
      const confidence=Number(row.dataset.confidence)||0;

      if(row.dataset.ignored==='1' || /^ign/i.test(value)){
        item.x.__ignoredByUser=true;
        return;
      }

      item.x.__ignoredByUser=false;

      /*
       * >=85% suggestion:
       * no manual selection is required. Preserve the API suggestion.
       */
      if(confidence>=85 && !value){
        const suggested=item.x.master_sku||item.x.suggestion||'';
        if(suggested){
          item.x.master_sku=suggested;
          item.x.__reviewValue=suggested;
        }
      }else if(value){
        item.x.master_sku=value;
        item.x.__reviewValue=value;
        item.x.confidence=100;
      }

      // Any resolved review row becomes matched.
      const idx=(item.r.reviews||[]).indexOf(item.x);
      if(idx>=0) item.r.reviews.splice(idx,1);

      if(!(item.r.matched||[]).includes(item.x))
        item.r.matched.push(item.x);
    });

    results.forEach(r=>{
      r.matched=(r.matched||[]).filter(x=>!x.__ignoredByUser);
      r.reviews=(r.reviews||[]).filter(x=>!x.__ignoredByUser);
    });

    /*
     * IMPORTANT:
     * Remaining reviews are only genuine unresolved <85% rows.
     * High-confidence API matches are never sent back for validation.
     */
    const remaining=results.flatMap(r=>r.reviews||[]).filter(x=>{
      const c=Number(x.confidence)||0;
      return c<85 && !x.__ignoredByUser;
    });

    if(remaining.length){
      buildReview();
      toast(`${remaining.length} SKU(s) below 85% still need validation.`);
      tab('reviewTab');
      updateApplyState();
      return;
    }

    finalReady=true;
    renderOutput();
    tab('outputTab');

    progress(
      100,
      'Compilation complete',
      'All >=85% matches accepted automatically; only required corrections applied.'
    );

    $('download').disabled=false;

  }catch(e){
    toast(e.message);
    progress(0,'Compilation error',e.message);
    $('applyAll').disabled=false;
  }
};

function getCanonicalMatchedRows(){
  /*
   * One statement product = one source SKU for a given HQ + Stockist.
   * Exact repeated source SKU copies are ignored WITHOUT adding quantities.
   * Different source SKUs mapping to the same master SKU remain separate so
   * the SKU Review duplicate/collision logic can flag them.
   */
  const seen=new Set();
  const rows=[];

  results.forEach(r=>{
    (r.matched||[]).forEach(x=>{
      if(x.__ignoredByUser) return;

      const source=String(x.source_sku||'').trim();
      const key=skuNorm(r.hq)+'|'+skuNorm(r.customer)+'|'+skuNorm(source);

      if(seen.has(key)) return;
      seen.add(key);
      rows.push({r,x});
    });
  });

  return rows;
}

function renderOutput(){let rows=[],sv=0,cv=0,m=0;getCanonicalMatchedRows().forEach(({r,x})=>{const sec=Number(x.sec)||0,clo=Number(x.close)||0,pts=Number(x.pts)||0;const a=sec*pts,b=clo*pts;sv+=a;cv+=b;m++;rows.push([r.hq,r.customer,x.master_sku,pts,sec,a,clo,b])});$('outBody').innerHTML=rows.slice(0,2000).map(a=>'<tr>'+a.map(v=>`<td>${esc(typeof v==='number'?v.toLocaleString('en-IN',{maximumFractionDigits:2}):v)}</td>`).join('')+'</tr>').join('');$('oSec').textContent=sv.toLocaleString('en-IN',{maximumFractionDigits:2});$('oClo').textContent=cv.toLocaleString('en-IN',{maximumFractionDigits:2});$('oMatched').textContent=m;$('oFiles').textContent=files.length;renderSummary()}
function renderSummary(){const stock={},hq={},prod={};const days=30;getCanonicalMatchedRows().forEach(({r,x})=>{const pts=Number(x.pts)||0,sec=Number(x.sec)||0,clo=Number(x.close)||0,sv=sec*pts,cv=clo*pts;const sk=r.hq+'|'+r.customer;stock[sk]??={hq:r.hq,c:r.customer,su:0,sv:0,cu:0,cv:0};Object.assign(stock[sk],{su:stock[sk].su+sec,sv:stock[sk].sv+sv,cu:stock[sk].cu+clo,cv:stock[sk].cv+cv});hq[r.hq]??={su:0,sv:0,cu:0,cv:0};Object.assign(hq[r.hq],{su:hq[r.hq].su+sec,sv:hq[r.hq].sv+sv,cu:hq[r.hq].cu+clo,cv:hq[r.hq].cv+cv});prod[x.master_sku]??={pts,su:0,sv:0,cu:0,cv:0};Object.assign(prod[x.master_sku],{su:prod[x.master_sku].su+sec,sv:prod[x.master_sku].sv+sv,cu:prod[x.master_sku].cu+clo,cv:prod[x.master_sku].cv+cv})});let html='<h3>Stockist-wise</h3><div class="tablewrap"><table><tr><th>HQ</th><th>Stockist</th><th>SEC Value</th><th>Closing Value</th><th>Inventory Days</th><th>Status</th></tr>';Object.values(stock).forEach(x=>{const d=x.sv?x.cv/x.sv*days:9999;html+=`<tr><td>${esc(x.hq)}</td><td>${esc(x.c)}</td><td>${x.sv.toFixed(2)}</td><td>${x.cv.toFixed(2)}</td><td>${d.toFixed(1)}</td><td>${x.sv===0&&x.cv>0?'NON-MOVING':d>40?'>40 DAYS':'NORMAL'}</td></tr>`});html+='</table></div><h3>HQ-wise</h3><div class="tablewrap"><table><tr><th>HQ</th><th>SEC Value</th><th>Closing Value</th><th>Inventory Days</th><th>Status</th></tr>';Object.entries(hq).forEach(([k,x])=>{const d=x.sv?x.cv/x.sv*days:9999;html+=`<tr><td>${esc(k)}</td><td>${x.sv.toFixed(2)}</td><td>${x.cv.toFixed(2)}</td><td>${d.toFixed(1)}</td><td>${x.sv===0&&x.cv>0?'NON-MOVING':d>40?'>40 DAYS':'NORMAL'}</td></tr>`});html+='</table></div><h3>Product Analysis</h3><div class="tablewrap"><table><tr><th>SKU</th><th>PTS</th><th>SEC Value</th><th>Closing Value</th><th>Inventory Days</th><th>Status</th></tr>';Object.entries(prod).sort((a,b)=>b[1].cv-a[1].cv).forEach(([k,x])=>{const d=x.sv?x.cv/x.sv*days:9999;html+=`<tr><td>${esc(k)}</td><td>${x.pts}</td><td>${x.sv.toFixed(2)}</td><td>${x.cv.toFixed(2)}</td><td>${d.toFixed(1)}</td><td>${x.sv===0&&x.cv>0?'NON-MOVING':d>40?'>40 DAYS':'NORMAL'}</td></tr>`});html+='</table></div>';

const groupMap={};
getCanonicalMatchedRows().forEach(({r,x})=>{
  const sku=String(x.master_sku||'');
  const n=skuNorm(sku);
  let g=productGroupNameFinal(sku);
  groupMap[g]??={group:g,skus:new Set(),sv:0,cv:0,su:0,cu:0};
  const pts=Number(x.pts)||0,sec=Number(x.sec)||0,clo=Number(x.close)||0;
  groupMap[g].skus.add(sku); groupMap[g].su+=sec; groupMap[g].sv+=sec*pts; groupMap[g].cu+=clo; groupMap[g].cv+=clo*pts;
});
const groupOrder=['Empanorm Group','Cilnikem Group','Glucoryl M','Glucoryl MV','Rosukem Group','Linapil Group','Voglikem Group','Glucoryl','Other'];
html+='<h3>SKU Group Analysis</h3><div class="tablewrap"><table><tr><th>Product Group</th><th>SKU Count</th><th>SEC Units</th><th>SEC Value</th><th>CLO Units</th><th>Closing Value</th><th>Inventory Days</th><th>Status</th></tr>';
groupOrder.forEach(k=>{const x=groupMap[k];if(!x)return;const d=x.sv?x.cv/x.sv*days:(x.cv>0?9999:0);html+=`<tr><td>${esc(x.group)}</td><td>${x.skus.size}</td><td>${x.su}</td><td>${x.sv.toFixed(2)}</td><td>${x.cu}</td><td>${x.cv.toFixed(2)}</td><td>${d===9999?'∞':d.toFixed(1)}</td><td>${x.sv===0&&x.cv>0?'NON-MOVING':d>40?'>40 DAYS':'NORMAL'}</td></tr>`});
html+='</table></div>';$('summary').innerHTML=html}
function makeWorkbook(){const agg=new Map();getCanonicalMatchedRows().forEach(({r,x})=>{const key=r.hq+'||'+r.customer+'||'+x.master_sku;const v=agg.get(key)||{sec:0,clo:0};v.sec+=Number(x.sec)||0;v.clo+=Number(x.close)||0;agg.set(key,v)});const idx={};masterHeaders.forEach((h,i)=>idx[h]=i);const outHeaders=[...masterHeaders.filter(Boolean)];['SEC UNITS','SEC VALUE','CLO UNITS','CLO VALUE'].forEach(h=>{if(!outHeaders.includes(h))outHeaders.push(h)});const out=[outHeaders];masterRows.forEach(mr=>{const v=agg.get(mr.HQ+'||'+mr['CUSTOMER NAME']+'||'+mr['SKU NAME'])||{sec:0,clo:0};const row=outHeaders.map(h=>{if(h==='SEC UNITS')return v.sec;if(h==='SEC VALUE')return Math.round(v.sec*mr.PTS*100)/100;if(h.trim()==='CLO UNITS')return v.clo;if(h==='CLO VALUE')return Math.round(v.clo*mr.PTS*100)/100;return mr[h]??''});out.push(row)});const wb=XLSX.utils.book_new();const ws=XLSX.utils.aoa_to_sheet(out);XLSX.utils.book_append_sheet(wb,ws,'MAIN FILE');
const stock=new Map(),hqs=new Map(),prod=new Map(),days=31;for(const [key,v] of agg){const [hq,cust,sku]=key.split('||');const mr=masterRows.find(r=>r.HQ===hq&&r['CUSTOMER NAME']===cust&&r['SKU NAME']===sku);const pts=mr?.PTS||0,sv=v.sec*pts,cv=v.clo*pts;let a=stock.get(hq+'||'+cust)||{hq,cust,su:0,sv:0,cu:0,cv:0};a.su+=v.sec;a.sv+=sv;a.cu+=v.clo;a.cv+=cv;stock.set(hq+'||'+cust,a);let h=hqs.get(hq)||{hq,su:0,sv:0,cu:0,cv:0};h.su+=v.sec;h.sv+=sv;h.cu+=v.clo;h.cv+=cv;hqs.set(hq,h);let p=prod.get(sku)||{sku,pts,su:0,sv:0,cu:0,cv:0};p.su+=v.sec;p.sv+=sv;p.cu+=v.clo;p.cv+=cv;prod.set(sku,p)}
function inv(x){return x.sv?x.cv/x.sv*days:(x.cv>0?9999:0)}function stat(x){const d=inv(x);return x.sv===0&&x.cv>0?'NON-MOVING':d>40?'>40 DAYS':'NORMAL'}
const ss=[['HQ','STOCKIST','SEC UNITS','SEC VALUE','CLO UNITS','CLO VALUE','INVENTORY DAYS','STATUS'],...Array.from(stock.values()).sort((a,b)=>a.hq.localeCompare(b.hq)||a.cust.localeCompare(b.cust)).map(x=>[x.hq,x.cust,x.su,Math.round(x.sv*100)/100,x.cu,Math.round(x.cv*100)/100,Math.round(inv(x)*10)/10,stat(x)])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(ss),'STOCKIST SUMMARY');
const hh=[['HQ','SEC UNITS','SEC VALUE','CLO UNITS','CLO VALUE','INVENTORY DAYS','STATUS'],...Array.from(hqs.values()).sort((a,b)=>a.hq.localeCompare(b.hq)).map(x=>[x.hq,x.su,Math.round(x.sv*100)/100,x.cu,Math.round(x.cv*100)/100,Math.round(inv(x)*10)/10,stat(x)])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(hh),'HQ SUMMARY');
const pp=[['SKU','PTS','SEC UNITS','SEC VALUE','CLO UNITS','CLO VALUE','INVENTORY DAYS','STATUS'],...Array.from(prod.values()).sort((a,b)=>b.cv-a.cv).map(x=>[x.sku,x.pts,x.su,Math.round(x.sv*100)/100,x.cu,Math.round(x.cv*100)/100,Math.round(inv(x)*10)/10,stat(x)])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(pp),'PRODUCT ANALYSIS');

const groupRules=[
  ['EMPANORM GROUP',s=>productGroupNameFinal(s)==='EMPANORM GROUP'],
  ['CILNIKEM GROUP',s=>productGroupNameFinal(s)==='CILNIKEM GROUP'],
  ['GLUCORYL M GROUP',s=>productGroupNameFinal(s)==='GLUCORYL M GROUP'],
  ['GLUCORYL MV GROUP',s=>productGroupNameFinal(s)==='GLUCORYL MV GROUP'],
  ['ROSUKEM GROUP',s=>productGroupNameFinal(s)==='ROSUKEM GROUP'],
  ['LINAPIL GROUP',s=>productGroupNameFinal(s)==='LINAPIL GROUP']
];
function skuGroup(sku){
  const s=skuNorm(sku);
  for(const [name,test] of groupRules) if(test(s)) return name;
  return 'OTHER / REMAINING';
}
const groups=new Map();
for(const p of prod.values()){
  const g=skuGroup(p.sku);
  const a=groups.get(g)||{group:g,skuSet:new Set(),su:0,sv:0,cu:0,cv:0};
  a.skuSet.add(p.sku);
  a.su+=p.su; a.sv+=p.sv; a.cu+=p.cu; a.cv+=p.cv;
  groups.set(g,a);
}
const groupOrder=['EMPANORM GROUP','CILNIKEM GROUP','GLUCORYL M GROUP','GLUCORYL MV GROUP','ROSUKEM GROUP','LINAPIL GROUP','OTHER / REMAINING'];
const ga=[['PRODUCT GROUP','SKU COUNT','SEC UNITS','SEC VALUE','CLO UNITS','CLO VALUE','INVENTORY DAYS','STATUS','SEC VALUE %']];
for(const name of groupOrder){
  const x=groups.get(name);
  if(!x) continue;
  const d=inv(x);
  ga.push([
    x.group,
    x.skuSet.size,
    x.su,
    Math.round(x.sv*100)/100,
    x.cu,
    Math.round(x.cv*100)/100,
    Math.round(d*10)/10,
    stat(x),
    0
  ]);
}
const totalGroupSec=Array.from(groups.values()).reduce((s,x)=>s+x.sv,0);
for(let i=1;i<ga.length;i++) ga[i][8]=totalGroupSec?Math.round((ga[i][3]/totalGroupSec)*10000)/100:0;
XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(ga),'SKU GROUP ANALYSIS');

/* ===== CRITICAL INVENTORY: >40 DAYS ===== */
const criticalRows=[];
const hqMap=new Map();

const canonicalCritical=getCanonicalMatchedRows();
canonicalCritical.forEach(({r,x})=>{
  const hq=r.hq||'UNMAPPED HQ';
  const customer=r.customer||r.file||'Unknown Stockist';
  const sku=String(x.master_sku||x.sku||'').trim();
  const sec=Number(x.sec)||0;
  const clo=Number(x.close)||0;
  const pts=Number(x.pts)||0;
  const secVal=sec*pts;
  const cloVal=clo*pts;

  /* Consistent 30-day inventory basis across portal and Excel */
  const days=sec>0 ? (clo/sec)*30 : (clo>0?9999:0);

  if(days>40){
    const h=hqMap.get(hq)||{
      hq:hq,stockists:new Set(),skuSet:new Set(),
      rows:0,cloUnits:0,cloValue:0,secValue:0,maxDays:0
    };

    h.stockists.add(customer);
    h.skuSet.add(sku);
    h.rows++;
    h.cloUnits+=clo;
    h.cloValue+=cloVal;
    h.secValue+=secVal;
    h.maxDays=Math.max(h.maxDays,days);
    hqMap.set(hq,h);

    criticalRows.push([
      hq,customer,sku,sec,secVal,clo,cloVal,
      Math.round(days*10)/10,
      days>60?'CRITICAL - >60 DAYS':'HIGH - >40 DAYS'
    ]);
  }
});

criticalRows.sort((a,b)=>b[7]-a[7]);((a,b)=>b[7]-a[7]);

const criticalSheetData=[
  ['HQ','Stockist','SKU','SEC Units','SEC Value','Closing Units','Closing Value','Inventory Days','Priority'],
  ...criticalRows
];
const criticalWS=XLSX.utils.aoa_to_sheet(criticalSheetData);
XLSX.utils.book_append_sheet(wb,criticalWS,'>40 DAYS CRITICAL');

const hqFocusData=[
  ['HQ','Critical SKU Count','Stockist Count','SKU Count','Closing Units','Closing Value','SEC Value','Highest Inventory Days','Focus Priority'],
  ...Array.from(hqMap.values())
    .sort((a,b)=>b.cloValue-a.cloValue || b.maxDays-a.maxDays)
    .map(h=>[
      h.hq,h.rows,h.stockists.size,h.skuSet.size,
      h.cloUnits,Math.round(h.cloValue*100)/100,
      Math.round(h.secValue*100)/100,
      Math.round(h.maxDays*10)/10,
      h.maxDays>60?'1 - IMMEDIATE FOCUS':'2 - PRIORITY FOCUS'
    ])
];
const hqFocusWS=XLSX.utils.aoa_to_sheet(hqFocusData);
XLSX.utils.book_append_sheet(wb,hqFocusWS,'HQ FOCUS PRIORITY');

/* Highlight >40-day inventory rows in Excel */
function highlightInventory(ws,daysCol){
  if(!ws || !ws['!ref']) return;
  const rg=XLSX.utils.decode_range(ws['!ref']);
  for(let R=1;R<=rg.e.r;R++){
    const days=Number(ws[XLSX.utils.encode_cell({r:R,c:daysCol})]?.v)||0;
    if(days>40){
      const fill=days>60?'FFC7CE':'FFF2CC';
      for(let C=0;C<=rg.e.c;C++){
        const cell=ws[XLSX.utils.encode_cell({r:R,c:C})];
        if(ws[cell]){
          ws[cell].s={
            fill:{patternType:'solid',fgColor:{rgb:fill}},
            font:{bold:true}
          };
        }
      }
    }
  }
}

/* PRODUCT ANALYSIS: Inventory Days is column G = index 6 */
highlightInventory(wb.Sheets['PRODUCT ANALYSIS'],6);

/* >40 DAYS CRITICAL: Inventory Days is column H = index 7 */
highlightInventory(criticalWS,7);

/* HQ focus: Highest Inventory Days is column H = index 7 */
highlightInventory(hqFocusWS,7);

return wb}
$('download').onclick=()=>{if(!finalReady)return toast('Validate SKUs first');progress(98,'Building Excel','Filling SEC/CLO units and PTS values…');const wb=makeWorkbook();XLSX.writeFile(wb,`Stock_Statement_Compiled_${$('month').value||'MAY26'}.xlsx`);progress(100,'Excel ready','Downloaded successfully.')}
$('clear').onclick=()=>{files=[];results=[];aliases={};identityMeta=[];finalReady=false;$('files').value='';$('master').value='';$('filesBody').innerHTML='';$('reviewBody').innerHTML='';$('outBody').innerHTML='';$('summary').innerHTML='';$('download').disabled=true;updateCount();progress(0,'Ready','Waiting for files');tab('filesTab')}
async function checkApi(){try{const r=await fetch('/api/health',{cache:'no-store'});const d=await r.json();if(!r.ok||!d.ok)throw Error('Health check failed');$('apiStatus').innerHTML='<span class="pill ok">API: Online</span>'}catch(e){$('apiStatus').innerHTML='<span class="pill bad">API: Offline — deploy the project with the /api folder</span>';}}checkApi();loadBundledMaster();updateCount();
  /* ===== SKU ACTION HIGHLIGHTS v6 — OCCURRENCE SAFE / NO BLINK ===== */
(function(){
  let currentFilter='all';

  function normAction(v){
    return String(v||'').trim().toUpperCase()
      .replace(/[^A-Z0-9]+/g,' ')
      .replace(/\s+/g,' ')
      .trim();
  }

  function getRowMaster(row){
    const input=row.querySelector('.reviewSelect');
    if(!input) return '';

    const value=String(input.value||'').trim();

    if(!value || /^🚫?\s*IGNORE/i.test(value))
      return '';

    return normAction(value);
  }

  function scanSkuActions(){
    const body=document.getElementById('reviewBody');
    const banner=document.getElementById('skuActionBanner');
    if(!body || !banner) return;

    const rows=[...body.querySelectorAll('tr[data-review-row]')];

    if(!rows.length){
      banner.className='sku-action-banner';
      banner.textContent='No SKU rows available for review.';
      return;
    }

    /*
     * IMPORTANT:
     * Exact source-SKU repeats have already been removed by buildReview().
     * Therefore a source repeat is NEVER called a duplicate here.
     *
     * TRUE DUPLICATE = two DIFFERENT source SKUs for the same
     * HQ + Stockist currently pointing to the SAME Master SKU.
     *
     * Example:
     *   EMPANORM 10   -> EMPANORM L 10
     *   EMPANORM L 10 -> EMPANORM L 10
     * These are a genuine collision.
     */
    const masterGroups=new Map();

    rows.forEach(row=>{
      row.classList.remove('sku-row-action','sku-row-duplicate');
      row.querySelectorAll('.sku-flag').forEach(e=>e.remove());

      row.dataset.skuDuplicate='0';
      row.dataset.skuCollision='0';
      row.dataset.skuLow='0';
      row.dataset.skuAction='0';

      const source=row.dataset.sourceKey||'';
      const hq=row.dataset.hqKey||'';
      const customer=row.dataset.customerKey||'';
      const master=getRowMaster(row);

      if(!source || !master) return;

      const masterKey=hq+'|'+customer+'|'+master;

      if(!masterGroups.has(masterKey))
        masterGroups.set(masterKey,[]);

      masterGroups.get(masterKey).push({
        row,
        source
      });
    });

    const collisionRows=new Set();

    masterGroups.forEach(group=>{
      const uniqueSources=new Set(group.map(x=>x.source));

      // Only different source SKUs mapping to one Master SKU.
      if(uniqueSources.size>1){
        group.forEach(x=>collisionRows.add(x.row));
      }
    });

    let actionCount=0;
    let lowCount=0;
    let duplicateCount=0;

    rows.forEach(row=>{
      const confidence=Number(row.dataset.confidence)||0;
      const isCollision=collisionRows.has(row);
      const isLow=confidence<85 && row.dataset.userValidated!=='1';

      if(isCollision){
        row.dataset.skuDuplicate='1';
        row.dataset.skuCollision='1';
        row.dataset.skuAction='1';
        row.classList.add('sku-row-duplicate');
        duplicateCount++;

        const cell=row.children[6];
        if(cell){
          const f=document.createElement('span');
          f.className='sku-flag collision';
          f.textContent='⚠ DUPLICATE MASTER SKU';
          cell.appendChild(f);
        }
      }

      if(isLow){
        row.dataset.skuLow='1';
        row.dataset.skuAction='1';
        row.classList.add('sku-row-action');
        lowCount++;

        const cell=row.children[6];
        if(cell){
          const f=document.createElement('span');
          f.className='sku-flag low';
          f.textContent='<85% REVIEW';
          cell.appendChild(f);
        }
      }

      if(isCollision || isLow)
        actionCount++;
    });

    const ac=document.getElementById('skuActionCount');
    const lc=document.getElementById('skuLowCount');
    const dc=document.getElementById('skuDupCount');

    if(ac) ac.textContent=actionCount;
    if(lc) lc.textContent=lowCount;
    if(dc) dc.textContent=duplicateCount;

    if(actionCount){
      const collisionCount=[...rows]
        .filter(r=>r.dataset.skuCollision==='1').length;

      banner.className='sku-action-banner';
      banner.innerHTML=
        `<span>⚠ ACTION REQUIRED: ${actionCount} SKU row(s) remaining</span>`+
        `<span>• ${duplicateCount} duplicate master row(s)</span>`+
        `<span>• ${collisionCount} collision row(s)</span>`+
        `<span>• ${lowCount} below 85%</span>`;
    }else{
      banner.className='sku-action-banner ok';
      banner.textContent='✓ All duplicate/collision and <85% action items are cleared.';
    }

    applySkuFilter();
  }


  function applySkuFilter(){
    const body=document.getElementById('reviewBody');
    if(!body) return;

    const rows=[...body.querySelectorAll('tr[data-review-row]')];

    rows.forEach(row=>{
      let show=true;

      if(currentFilter==='duplicate'){
        show=row.dataset.skuCollision==='1';
      }else if(currentFilter==='low'){
        show=row.dataset.skuLow==='1';
      }else if(currentFilter==='action'){
        show=row.dataset.skuAction==='1';
      }else{
        show=true;
      }

      row.style.display=show?'':'none';
    });

    document.querySelectorAll('[data-sku-filter]').forEach(btn=>{
      btn.classList.toggle(
        'active',
        (btn.dataset.skuFilter||'all')===currentFilter
      );
    });
  }

window.__refreshSkuActionScan=scanSkuActions;

  document.addEventListener('input',e=>{
    if(e.target.closest('.reviewSelect'))
      setTimeout(scanSkuActions,20);
  });

  document.addEventListener('change',e=>{
    if(e.target.closest('.reviewSelect'))
      setTimeout(scanSkuActions,20);
  });

  document.addEventListener('click',e=>{
    const btn=e.target.closest('[data-sku-filter]');

    if(btn){
      currentFilter=btn.dataset.skuFilter||'all';
      applySkuFilter();
    }

    if(e.target.closest('#reviewTab'))
      setTimeout(scanSkuActions,50);
  });

  const body=document.getElementById('reviewBody');

  if(body){
    const observer=new MutationObserver(()=>{
      setTimeout(scanSkuActions,20);
    });

    observer.observe(body,{childList:true,subtree:true});
  }

  setTimeout(scanSkuActions,200);
})();

/* ===== SKU SEARCHABLE DROPDOWN v2 — STABLE IN FILTERED VIEWS ===== */
(function(){
  let openMenu=null;
  let openInput=null;

  function closeMenu(){
    if(openMenu){ openMenu.remove(); openMenu=null; }
    openInput=null;
  }

  function positionMenu(menu,input){
    const r=input.getBoundingClientRect();
    menu.style.left=Math.max(6,r.left)+'px';
    menu.style.top=(r.bottom+4)+'px';
    menu.style.width=Math.max(280,r.width)+'px';
  }

  function getPool(input){
    const row=input.closest('tr');
    if(!row || typeof poolFor!=='function') return [];
    const customer=row.dataset.customerKey||'';
    const hq=row.dataset.hqKey||'';
    // dataset keys are normalized; find the original row values from reviewItems.
    const item=(window.__reviewItems||[]).find(o=>o.rowId===row.dataset.rowId);
    if(item) return poolFor(item.r.customer,item.r.hq);
    return [];
  }

  function showMenu(input,query){
    closeMenu();
    const menu=document.createElement('div');
    menu.className='sku-combo-menu';
    const pool=getPool(input);
    const q=String(query||'').trim().toUpperCase();
    const seen=new Set();
    const options=[];

    pool.forEach(m=>{
      const name=String(m.SKU_NAME||'').trim();
      const key=skuNorm(name);
      if(!name || seen.has(key)) return;
      seen.add(key);
      if(!q || skuNorm(name).includes(skuNorm(q))) options.push(name);
    });
    options.push('IGNORE — Not Our Product');

    if(!options.length){
      const empty=document.createElement('div');
      empty.className='sku-combo-empty';
      empty.textContent='No matching Master SKU';
      menu.appendChild(empty);
    }else{
      options.slice(0,150).forEach(name=>{
        const item=document.createElement('div');
        item.className='sku-combo-option';
        item.textContent=name;
        item.title=name;
        item.addEventListener('mousedown',e=>{
          e.preventDefault();
          input.value=name;
          input.dispatchEvent(new Event('change',{bubbles:true}));
          closeMenu();
          input.focus();
        });
        menu.appendChild(item);
      });
    }

    document.body.appendChild(menu);
    positionMenu(menu,input);
    openMenu=menu;
    openInput=input;
  }

  function initOne(input){
    if(!input || input.dataset.comboReady==='1') return;
    input.dataset.comboReady='1';
    input.removeAttribute('list');

    const wrap=input.parentElement;
    if(wrap){
      wrap.classList.add('sku-combo-wrap');
      let btn=wrap.querySelector('.sku-combo-arrow');
      if(!btn){
        btn=document.createElement('button');
        btn.type='button';
        btn.className='sku-combo-arrow';
        btn.tabIndex=-1;
        btn.textContent='▼';
        btn.title='Show all Master SKUs';
        wrap.appendChild(btn);
        btn.addEventListener('mousedown',e=>{
          e.preventDefault();
          showMenu(input,'');
          input.focus();
        });
      }
    }

    input.addEventListener('focus',()=>{
      // Always show the complete list first. User can then type to filter.
      showMenu(input,'');
    });
    input.addEventListener('click',()=>{
      if(openInput!==input) showMenu(input,'');
      else positionMenu(openMenu,input);
    });
    input.addEventListener('input',()=>{
      // Typing filters the menu only; it never filters/hides the review row.
      showMenu(input,input.value);
    });

    input.addEventListener('keydown',e=>{
      if(e.key==='Escape'){ closeMenu(); return; }
      if(e.key==='ArrowDown' && openMenu){
        const first=openMenu.querySelector('.sku-combo-option');
        if(first){e.preventDefault();first.focus();}
      }
    });
  }

  function initAll(){
    document.querySelectorAll('.reviewSelect').forEach(initOne);
  }

  document.addEventListener('mousedown',e=>{
    if(openMenu && !e.target.closest('.sku-combo-menu') && e.target!==openInput && !e.target.closest('.sku-combo-arrow'))
      closeMenu();
  });
  window.addEventListener('resize',()=>{if(openMenu&&openInput)positionMenu(openMenu,openInput)});
  window.addEventListener('scroll',()=>{if(openMenu&&openInput)positionMenu(openMenu,openInput)},true);

  const body=document.getElementById('reviewBody');
  if(body){
    const observer=new MutationObserver(()=>setTimeout(initAll,0));
    observer.observe(body,{childList:true,subtree:true});
  }
  setTimeout(initAll,100);
})();

/* ===== PRODUCT GROUP MATCHING FIX v2 ===== */
function productGroupName(raw){
  const s=String(raw||'')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g,' ')
    .replace(/\s+/g,' ')
    .trim();

  if(s.startsWith('EMPANORM')) return 'EMPANORM GROUP';
  if(s.startsWith('CILNIKEM')) return 'CILNIKEM GROUP';

  /* MV must be tested BEFORE M */
  if(s.startsWith('GLUCORYL MV') || /\bGLUCORYL\s*MV\b/.test(s))
    return 'GLUCORYL MV GROUP';

  if(s.startsWith('GLUCORYL M') || /\bGLUCORYL\s*M\b/.test(s))
    return 'GLUCORYL M GROUP';

  if(s.startsWith('ROSUKEM')) return 'ROSUKEM GROUP';
  if(s.startsWith('LINAPIL')) return 'LINAPIL GROUP';

  return 'OTHER / REMAINING';
}


/* ===== FINAL PRODUCT GROUP RULES ===== */
function productGroupNameFinal(s){
  const raw=String(s||'').toUpperCase()
    .replace(/’/g,"'")
    .replace(/[^A-Z0-9./+\-]+/g,' ')
    .replace(/\s+/g,' ')
    .trim();

  // IMPORTANT: check MV before M so every MV SKU stays in Glucoryl MV.
  if(/\bGLUCORYL\s*[- ]?MV\b/.test(raw) || /\bGLUCORYL\s*[- ]?MV\d/.test(raw))
    return 'Glucoryl MV';

  // Glucoryl M family, including M1/M2/M3/M4 Forte, M 0.5, etc.
  if(/\bGLUCORYL\s*[- ]?M(?:\d|[ .-]?0\.5\b)/.test(raw))
    return 'Glucoryl M';

  if(/\bGLUCORYL\b/.test(raw))
    return 'Glucoryl';

  if(/\bEMPANORM\b/.test(raw)) return 'Empanorm Group';
  if(/\bCILNIKEM\b/.test(raw)) return 'Cilnikem Group';
  if(/\bROSUKEM\b/.test(raw)) return 'Rosukem Group';
  if(/\bLINAPIL\b/.test(raw)) return 'Linapil Group';
  if(/\bVOGLIKEM\b/.test(raw)) return 'Voglikem Group';

  return 'Other';
}

