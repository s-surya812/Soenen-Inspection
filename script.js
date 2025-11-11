/* Soenen Excel → Interactive form (replicates your filled format) */
const $ = sel => document.querySelector(sel);

/* UI handles */
const fileInput     = $('#fileInput');
const parseExcelBtn = $('#parseExcelBtn');
const buildFormBtn  = $('#buildFormBtn');
const exportPdfBtn  = $('#exportPdfBtn');
const headerOut     = $('#headerOut');
const hdrPart       = $('#hdrPart');
const hdrFsmSpec    = $('#hdrFsmSpec');
const formArea      = $('#formArea');
const warnings      = $('#warnings');

/* Data stores */
let workbook = null;
let rows = [];                 // main table rows (up to 45)
let header = {};               // header/meta
let extremes = {};             // bottom “Dimension verification” block

/* Helpers */
const show = (msg,type="info")=>{
  warnings.textContent = msg;
  warnings.style.padding = '8px';
  warnings.style.borderRadius = '6px';
  warnings.className = type==="warn" ? 'panel warn' : 'panel';
};
const clearMsg = ()=>{ warnings.textContent=""; warnings.className=''; };
const round = (n,d=2)=>Math.round(n*10**d)/10**d;
const isEmptyRow = r => !r || r.every(c => (c==null || (""+c).trim()===""));

/* -------------- Excel → JSON extraction --------------- */
/**
 * We do position-based extraction that matches your “final” file:
 * Rows:
 *   - Part/Rev/Hand on row ~5: cell B5 “PART NUMBER / LEVEL / HAND : P / L / H”
 *   - KB/PC row ~8
 *   - Spec blocks “ROOT WIDTH OF FSM”, “FSM LENGTH”, “TOTAL HOLES COUNT”
 *   - Main table starts at row with headings (Si. No., Press, Sel ID, Ref, X-axis, Spec (Y or Z axis), Spec Dia, Value from Hole edge, Actual Dia, Actual Y or Z, Offset)
 *   - Up to 45 rows (show blanks if fewer)
 *   - Bottom extremes block (Web PWX/PW/PWM, PFF, PFB columns)
 *
 * If any label shifts slightly, we also search by text.
 */
function parseExcel(workbook){
  const ws = workbook.Sheets[workbook.SheetNames[0]];
  const get = addr => (ws[addr] ? ws[addr].v : null);
  const findText = (needle)=>{
    const A = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
    if(!A) return null;
    for(let R=A.s.r; R<=A.e.r; R++){
      for(let C=A.s.c; C<=A.e.c; C++){
        const cell = ws[XLSX.utils.encode_cell({r:R,c:C})];
        if(!cell || typeof cell.v!=='string') continue;
        if(cell.v.toString().toLowerCase().includes(needle.toLowerCase()))
          return {r:R,c:C,addr:XLSX.utils.encode_cell({r:R,c:C}),value:cell.v};
      }
    }
    return null;
  };

  /* Header: Part / Level / Hand */
  // Prefer finding the label
  let pLoc = findText('PART NUMBER / LEVEL / HAND');
  if(pLoc){
    // Value is likely to the right in same row (string like "FH140813 / O2 / LH")
    const vCell = XLSX.utils.encode_cell({r:pLoc.r,c:pLoc.c+1});
    const val = get(vCell) || '';
    const parts = val.toString().split('/').map(s=>s.trim());
    header.partNumber = parts[0] || '';
    header.revision   = parts[1] || '';
    header.hand       = parts[2] || '';
  }else{
    header.partNumber=''; header.revision=''; header.hand='';
  }

  /* Holes count */
  const hcLoc = findText('TOTAL HOLES COUNT');
  if(hcLoc){
    header.totalHoles = get(XLSX.utils.encode_cell({r:hcLoc.r,c:hcLoc.c+2})) ?? '';
  }

  /* Root width Spec */
  const rwLoc = findText('ROOT WIDTH OF FSM');
  if(rwLoc){
    const specCell = XLSX.utils.encode_cell({r:rwLoc.r,c:rwLoc.c+2});
    header.rootWidthSpec = (get(specCell)||'').toString().replace(/[^\d.]/g,'');
  }

  /* FSM length Spec */
  const flLoc = findText('FSM LENGTH');
  if(flLoc){
    const specCell = XLSX.utils.encode_cell({r:flLoc.r,c:flLoc.c+2});
    header.fsmLengthSpec = (get(specCell)||'').toString().replace(/[^\d.]/g,'');
  }

  /* KB & PC Code Spec and placeholders for Act */
  const kbLoc = findText('KB & PC Code');
  if(kbLoc){
    const specCell = XLSX.utils.encode_cell({r:kbLoc.r,c:kbLoc.c+2});
    const specVal = (get(specCell)||'').toString();
    // expected like "Spec- 1 / 0"
    const m = specVal.match(/(\d+)\s*\/\s*(\d+)/);
    header.kbSpec = m ? m[1] : '';
    header.pcSpec = m ? m[2] : '';
  }

  /* Matrix used */
  const matLoc = findText('Matrix used');
  header.matrix = matLoc ? (get(XLSX.utils.encode_cell({r:matLoc.r,c:matLoc.c+1}))||'') : '';

  /* Inspectors: we’ll just leave editable fields for 2 names */
  header.inspectors = ['', ''];

  /* Top header view */
  hdrPart.textContent   = `PART NUMBER / LEVEL / HAND : ${header.partNumber||'—'} / ${header.revision||'—'} / ${header.hand||'—'}`;
  hdrFsmSpec.textContent= `FSM LENGTH : ${header.fsmLengthSpec||'—'} mm`;

  /* ---------- Main table ---------- */
  // Find heading row by "Si. No." text
  const headLoc = findText('Si.');
  let startRow = headLoc ? headLoc.r+1 : 12; // data starts next row
  rows = [];
  for(let r = startRow; r < startRow+60; r++){
    // columns: A..K (0..10)
    const rowObj = {
      sl: get(`A${r+1}`),
      press: get(`B${r+1}`),
      sel: get(`C${r+1}`),
      ref: get(`D${r+1}`),
      x: get(`E${r+1}`),
      specYZ: get(`F${r+1}`),
      specDia: get(`G${r+1}`),
      actEdge: get(`H${r+1}`),   // not used in offset logic, but kept
      actDia: get(`I${r+1}`),    // editable
      actYZ: get(`J${r+1}`),     // editable
      offset: get(`K${r+1}`)     // will be computed live
    };
    const specCols = [rowObj.press,rowObj.sel,rowObj.ref,rowObj.x,rowObj.specYZ,rowObj.specDia];
    const allNull = [rowObj.sl,...specCols,rowObj.actEdge,rowObj.actDia,rowObj.actYZ].every(v=>v==null||v==='');
    if(allNull) break; // stop at first full empty
    rows.push(rowObj);
    if(rows.length>=45) break; // cap at 45
  }

  /* ----------- Extremes (bottom block) ----------- */
  // We search for the block title
  const exLoc = findText('Dimension verification for holes at');
  if(exLoc){
    // Spec columns appear a few rows below exLoc.r
    // Build a small helper to read by (row, col) offsets relative to the found block
    const R = exLoc.r;
    const cSpecLeft  = exLoc.c + 5; // First holes of FSM - Spec
    const cSpecRight = exLoc.c + 9; // Last holes of FSM - Spec
    // Web rows order: PWX, PW, PWM; then PFF (Top Flange), PFB (Bottom Flange)
    const map = [
      {key:'web_pwx', r:R+3},
      {key:'web_pw',  r:R+4},
      {key:'web_pwm', r:R+5},
      {key:'pff',     r:R+6},
      {key:'pfb',     r:R+7},
    ];
    extremes = {};
    map.forEach(m=>{
      extremes[m.key] = {
        first_spec : get(XLSX.utils.encode_cell({r:m.r,c:cSpecLeft})) ?? '',
        last_spec  : get(XLSX.utils.encode_cell({r:m.r,c:cSpecRight})) ?? ''
      };
    });
  }

  headerOut.textContent = JSON.stringify({header, rows:rows.length, extremes}, null, 2);
}

/* -------------- Build the on-screen form --------------- */
function makeInput({id, value='', placeholder='', readOnly=false, cls='input-small', size, step}={}){
  const el = document.createElement('input');
  el.type = 'text';
  el.className = cls;
  if(id) el.id=id;
  if(value!==undefined && value!==null) el.value = value;
  if(placeholder) el.placeholder = placeholder;
  if(readOnly) el.readOnly = true;
  if(size) el.size = size;
  if(step) el.step = step;
  return el;
}
function td(inner){const td=document.createElement('td'); td.append(inner); return td;}
function cell(inner){const d=document.createElement('div'); d.textContent=inner??''; return d;}

function buildHeaderBlock(){
  const blk = document.createElement('div'); blk.className='form-block';
  blk.innerHTML = `<strong>Header / Mandatory fields</strong>`;

  const r1 = document.createElement('div'); r1.className='form-row';
  // FSM Serial
  const c1 = document.createElement('div'); c1.className='col';
  c1.innerHTML = `<div class="small">FSM Serial Number:</div>`;
  c1.append(makeInput({id:'fsmSerial', placeholder:'Enter FSM Serial', cls:'input'}));
  // Inspectors (two fields)
  const c2 = document.createElement('div'); c2.className='col';
  c2.innerHTML = `<div class="small">Inspectors:</div>`;
  c2.append(makeInput({id:'insp1',placeholder:'Inspector 1',cls:'inline'}));
  c2.append(makeInput({id:'insp2',placeholder:'Inspector 2',cls:'inline'}));
  r1.append(c1,c2);

  const r2 = document.createElement('div'); r2.className='form-row';
  // KB & PC
  const c3 = document.createElement('div'); c3.className='col';
  c3.innerHTML = `<div class="small">KB & PC Code (Spec / Act)</div>`;
  c3.append(makeInput({id:'kbSpec',value:header.kbSpec||'',readOnly:true,cls:'inline',size:3}));
  c3.append(document.createTextNode(' / '));
  c3.append(makeInput({id:'kbAct',placeholder:'Act',cls:'inline',size:3}));
  c3.append(document.createTextNode(' — '));
  c3.append(makeInput({id:'pcSpec',value:header.pcSpec||'',readOnly:true,cls:'inline',size:3}));
  c3.append(document.createTextNode(' / '));
  c3.append(makeInput({id:'pcAct',placeholder:'Act',cls:'inline',size:3}));

  // Root width + FSM length
  const c4 = document.createElement('div'); c4.className='col';
  c4.innerHTML = `<div class="small">ROOT WIDTH OF FSM (Spec / Act) mm  |  FSM LENGTH (Spec / Act) mm</div>`;
  // root
  c4.append(makeInput({id:'rootSpec',value:header.rootWidthSpec||'',readOnly:true,cls:'inline',size:6}));
  c4.append(makeInput({id:'rootAct',placeholder:'Act',cls:'inline',size:6}));
  // fsm
  c4.append(makeInput({id:'fsmSpec',value:header.fsmLengthSpec||'',readOnly:true,cls:'inline',size:8}));
  c4.append(makeInput({id:'fsmAct',placeholder:'Act',cls:'inline',size:8}));
  r2.append(c3,c4);

  blk.append(r1,r2);
  formArea.append(blk);
}

function buildMainTable(){
  const blk = document.createElement('div'); blk.className='form-block';
  blk.innerHTML = `<strong>Main Inspection Table</strong>`;
  const table = document.createElement('table'); table.className='table';
  table.innerHTML = `
    <thead>
      <tr>
        <th>Sl No</th><th>Press</th><th>Sel ID</th><th>Ref</th>
        <th>X-axis</th><th>Spec (Y or Z)</th><th>Spec Dia</th>
        <th>Value from Hole edge (Act)</th><th>Actual Dia</th><th>Actual Y or Z</th>
        <th>Offset</th><th>Result</th>
      </tr>
    </thead>
  `;
  const tb = document.createElement('tbody');

  const toNum = v => {
    if(v==null||v==='') return '';
    const s = (''+v).replace(',', '.');
    return isFinite(+s) ? +s : (''+v);
  };

  const rowCount = Math.max(rows.length, 45);
  for(let i=0;i<rowCount;i++){
    const r = rows[i] || {};
    const tr = document.createElement('tr');

    const sl     = makeInput({value:r.sl??(i+1),readOnly:true});
    const press  = makeInput({value:r.press??'',readOnly:true});
    const sel    = makeInput({value:r.sel??'',readOnly:true});
    const ref    = makeInput({value:r.ref??'',readOnly:true});
    const xaxis  = makeInput({value:r.x??'',readOnly:true});
    const specYZ = makeInput({value:r.specYZ??'',readOnly:true});
    const specDia= makeInput({value:r.specDia??'',readOnly:true});
    const valEdge= makeInput({value:r.actEdge??''});                    // free field
    const actDia = makeInput({value:r.actDia??''});                     // editable
    const actYZ  = makeInput({value:r.actYZ??''});                      // editable
    const offsetCell = cell('');
    const resultCell = cell('');

    const recalc = ()=>{
      // parse dia (handle slots "9*12" or "9x12")
      let specDiaVal = (specDia.value||'').toString().trim();
      let slot = false, slotH=null, slotV=null;
      if(/^\s*\d+(\.\d+)?\s*[*xX]\s*\d+(\.\d+)?\s*$/.test(specDiaVal)){
        const mm = specDiaVal.toLowerCase().split(/[x*]/).map(s=>+s.trim());
        slot = true;
        slotH = mm[1];   // width
        slotV = mm[0];   // height → used for offset tolerance as per your rule
      }
      const sDia = slot ? slotV : toNum(specDiaVal);
      const sYZ  = toNum(specYZ.value);
      const aYZ  = toNum(actYZ.value);
      const aDia = toNum(actDia.value);
      const x    = toNum(xaxis.value);
      const fsm  = +($('#fsmSpec')?.value || header.fsmLengthSpec || 0);

      // offset = Spec(Y/Z) + (Spec Dia / 2) (use vertical for slot)
      let offset = (sYZ && sDia) ? ( +sYZ + (+sDia)/2 ) : '';
      offsetCell.textContent = offset==='' ? '' : round(offset,2);

      // tolerances
      // Dia tolerance: -0 to +0.5 (with special ≤10.7 → +0.4)
      let diaOk = true;
      if(aDia!=='' && sDia!==''){
        const up = (sDia<=10.7) ? 0.4 : 0.5;
        diaOk = (aDia >= sDia) && (aDia <= sDia + up);
      }else diaOk=false;

      // Offset tol: ±1 mm; edges (≤200 or ≥FSM-200) → ±2 mm
      let tol = 1;
      if(isFinite(x) && isFinite(fsm) && (x<=200 || x>=fsm-200)) tol = 2;

      let offOk = true;
      if(aYZ!=='' && offset!==''){
        offOk = Math.abs(aYZ - offset) <= tol;
      }else offOk=false;

      // result
      if(diaOk && offOk){
        resultCell.textContent='OK';
        resultCell.className='ok-cell';
      }else{
        resultCell.textContent='NOK';
        resultCell.className='nok-cell';
      }
    };

    [actDia, actYZ].forEach(el=>el.addEventListener('input', recalc));
    recalc();

    [sl,press,sel,ref,xaxis,specYZ,specDia,valEdge,actDia,actYZ].forEach(inp=>{
      tr.appendChild(td(inp));
    });
    tr.appendChild(td(offsetCell));
    tr.appendChild(td(resultCell));
    tb.appendChild(tr);
  }

  table.appendChild(tb);
  blk.appendChild(table);
  formArea.appendChild(blk);
}

function buildExtremesBlock(){
  const blk = document.createElement('div'); blk.className='form-block';
  blk.innerHTML = `<strong>Dimension verification for holes at extreme ends of FSM</strong>`;
  const table = document.createElement('table'); table.className='table';
  table.innerHTML = `
    <thead>
      <tr>
        <th rowspan="2">X Axis</th>
        <th rowspan="2">Web / Flange</th>
        <th colspan="3">First holes of FSM - from front end</th>
        <th colspan="3">Last holes of FSM - from rear end</th>
      </tr>
      <tr>
        <th>Spec</th><th>Actual</th><th>offset</th>
        <th>Spec</th><th>Actual</th><th>offset</th>
      </tr>
    </thead>
  `;
  const tb = document.createElement('tbody');

  const rowsMeta = [
    ['Web','PWX','web_pwx'],
    ['Web','PW','web_pw'],
    ['Web','PWM','web_pwm'],
    ['Top Flange','PFF','pff'],
    ['Bottom Flange','PFB','pfb']
  ];

  rowsMeta.forEach(([grp,label,key])=>{
    const tr = document.createElement('tr');
    tr.appendChild(td(cell(grp)));
    tr.appendChild(td(cell(label)));
    const fSpec = makeInput({value: extremes[key]?.first_spec ?? '', readOnly:true});
    const fAct  = makeInput({placeholder:'/'});
    const fOff  = cell('');
    const lSpec = makeInput({value: extremes[key]?.last_spec ?? '', readOnly:true});
    const lAct  = makeInput({placeholder:'/'});
    const lOff  = cell('');

    const recalc = ()=>{
      const fS = parseFloat((fSpec.value||'').toString());
      const fA = parseFloat((fAct.value||'').toString());
      fOff.textContent = (!isNaN(fS) && !isNaN(fA)) ? round(fA - fS,2) : '';
      const lS = parseFloat((lSpec.value||'').toString());
      const lA = parseFloat((lAct.value||'').toString());
      lOff.textContent = (!isNaN(lS) && !isNaN(lA)) ? round(lA - lS,2) : '';
    };
    [fAct,lAct].forEach(e=>e.addEventListener('input',recalc));
    recalc();

    [fSpec,fAct,fOff,lSpec,lAct,lOff].forEach(v=>tr.appendChild(td(v)));
    tb.appendChild(tr);
  });

  table.appendChild(tb);
  blk.appendChild(table);
  formArea.appendChild(blk);
}

function buildVisualChecks(){
  const blk = document.createElement('div'); blk.className='form-block';
  blk.innerHTML = `<strong>Visual / Binary Checks</strong>`;

  const grid = document.createElement('table'); grid.className='table';
  grid.innerHTML = `
    <thead><tr>
      <th>Punch Break</th><th>Length Variation</th><th>Slug mark</th><th>Machine Error (YES/NO)</th>
    </tr></thead>
  `;
  const tr1 = document.createElement('tr');
  const mk = ()=>{
    const yes = document.createElement('button'); yes.textContent='OK / YES';
    const no  = document.createElement('button'); no.textContent ='NOK / NO';
    yes.className='input-small'; no.className='input-small';
    yes.onclick=()=>{yes.className='ok-cell input-small'; no.className='input-small';};
    no.onclick =()=>{no.className='nok-cell input-small'; yes.className='input-small';};
    const d=document.createElement('div'); d.append(yes,' ',no); return d;
  };
  for(let i=0;i<4;i++){ tr1.appendChild(td(mk())); }
  grid.appendChild(tr1);

  const grid2 = document.createElement('table'); grid2.className='table';
  grid2.innerHTML = `
    <thead><tr>
      <th>Radius crack</th><th>Holes Burr</th><th>Line mark</th><th>Pit Mark</th>
    </tr></thead>
  `;
  const tr2 = document.createElement('tr');
  for(let i=0;i<4;i++){ tr2.appendChild(td(mk())); }
  grid2.appendChild(tr2);

  blk.append(grid, grid2);
  formArea.appendChild(blk);
}

function buildRemarks(){
  const blk = document.createElement('div'); blk.className='form-block';
  blk.innerHTML = `<strong>Remarks / Details of issue</strong>`;
  const ta = document.createElement('textarea'); ta.id='remarks'; ta.placeholder='Enter remarks (optional)';
  blk.appendChild(ta);
  formArea.appendChild(blk);
}

/* -------------- Export to PDF --------------- */
async function exportPDF(){
  exportPdfBtn.disabled = true; exportPdfBtn.textContent = 'Generating...';
  // clone area with header for nice capture
  const container = document.createElement('div');
  container.style.width = '900px';
  container.style.padding = '12px';
  const hdr = document.querySelector('header').cloneNode(true);
  const body = formArea.cloneNode(true);
  container.append(hdr, body);
  container.style.position='fixed'; container.style.left='-2000px';
  document.body.appendChild(container);

  try{
    const canvas = await html2canvas(container, {scale:2, useCORS:true});
    const img = canvas.toDataURL('image/png');
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF('p','mm','a4');
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();
    const imgProps = pdf.getImageProperties(img);
    const w = pageW - 20;
    const h = (imgProps.height * w) / imgProps.width;
    pdf.addImage(img, 'PNG', 10, 10, w, h);
    const now = new Date();
    const part = header.partNumber || 'PART';
    const rev  = header.revision   || 'X';
    const hand = header.hand       || 'H';
    const fn = `${part}_${rev}_${hand}_Inspection_${now.toISOString().slice(0,10)}.pdf`;
    pdf.save(fn);
    show('PDF exported: '+fn);
  }catch(e){
    alert('PDF export failed: '+e.message);
  }finally{
    document.body.removeChild(container);
    exportPdfBtn.disabled = false; exportPdfBtn.textContent = 'Export to PDF';
  }
}

/* -------------- Wiring --------------- */
parseExcelBtn.addEventListener('click', async ()=>{
  const f = fileInput.files?.[0];
  if(!f){ alert('Please choose the final Excel file (.xlsx).'); return; }
  try{
    const data = await f.arrayBuffer();
    workbook = XLSX.read(data, {type:'array'});
    parseExcel(workbook);
    buildFormBtn.disabled = false;
    show('Excel parsed. Click "Build Form".');
  }catch(e){
    console.error(e);
    alert('Failed to parse Excel: '+e.message);
  }
});

buildFormBtn.addEventListener('click', ()=>{
  if(!workbook){ show('Parse Excel first.', 'warn'); return; }
  formArea.innerHTML = '';
  clearMsg();
  buildHeaderBlock();
  buildMainTable();
  buildExtremesBlock();
  buildVisualChecks();
  buildRemarks();
  exportPdfBtn.disabled = false;
  show('Form built. You can now fill Actuals and export to PDF.');
});

exportPdfBtn.addEventListener('click', exportPDF);
