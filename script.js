/* Soenen Excel → Interactive form (fully replicates filled format) */

const $ = sel => document.querySelector(sel);
const fileInput      = $('#fileInput');
const parseExcelBtn  = $('#parseExcelBtn');
const buildFormBtn   = $('#buildFormBtn');
const exportPdfBtn   = $('#exportPdfBtn');
const headerOut      = $('#headerOut');
const formArea       = $('#formArea');
const warnings       = $('#warnings');
const hdrPart        = $('#hdrPart');
const hdrFsmSpec     = $('#hdrFsmSpec');

let rows2D = [];          // 2D array of sheet
let parsed  = {};         // header & meta
let tableRows = [];       // up to 45
let extreme = {};         // bottom "dimension verification" blocks

/* Utils */
const show = (m,t="info")=>{
  warnings.style.padding="8px";warnings.style.borderRadius="6px";warnings.style.marginTop="6px";
  warnings.textContent=m;
  if(t==="warn"){warnings.style.background="#fff6cc";warnings.style.color="#ff8c00";}
  else{warnings.style.background="#e8f4ff";warnings.style.color="#005fcc";}
};
const clearMsg=()=>{warnings.textContent="";warnings.removeAttribute('style');};
const round=(n,d=2)=>Math.round(n*10**d)/10**d;

/* Find a row index that has all labels */
function findRowIdxByLabels(labels){
  return rows2D.findIndex(r => {
    const s = r.map(v=>String(v||'').toLowerCase());
    return labels.every(lbl => s.some(x=>x.includes(lbl)));
  });
}
/* Find a cell by regex anywhere; return {r,c,val} */
function findCellByRegex(rx){
  for(let r=0;r<rows2D.length;r++){
    for(let c=0;c<rows2D[r].length;c++){
      const val = rows2D[r][c];
      if(val!=null && rx.test(String(val))) return {r,c,val};
    }
  }
  return null;
}

/* Parse slot/diameter spec: Option A (first = Height/Vertical) */
function parseSpecDia(raw){
  const s = String(raw||'').trim();
  const m = s.match(/^(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)/);
  if(m) return {isSlot:true, H:parseFloat(m[1]), W:parseFloat(m[2]), raw:s};
  const d = parseFloat(s);
  return {isSlot:false, dia: isFinite(d)?d:NaN, raw:s};
}

/* Dia tolerances */
function diaHoleOK(spec, act){
  if(!isFinite(spec) || !isFinite(act)) return null;
  if(spec <= 10.7) return (act >= spec && act <= spec + 0.4);
  if(spec >= 11.7) return (act >= spec && act <= spec + 0.5);
  return (act >= spec && act <= spec + 0.5); // 10.7..11.7 → +0.5 per your note
}
/* Slot tolerances: -0 / +0.5 both H & W */
function diaSlotOK(spec, act){ // spec,act single side
  if(!isFinite(spec) || !isFinite(act)) return null;
  return (act >= spec && act <= spec + 0.5);
}
/* Offset tolerance by X & FSM length */
function offsetTol(x, fsmLen){
  if(!isFinite(x)||!isFinite(fsmLen)) return 1;
  return (x<=200 || x>=fsmLen-200) ? 1.5 : 1;
}

/* Excel → 2D array and parse content */
parseExcelBtn.addEventListener('click', async ()=>{
  const f = fileInput.files?.[0];
  if(!f){ alert('Please choose the final macro sheet (.xlsx)'); return; }

  const buf = await f.arrayBuffer();
  const wb = XLSX.read(buf, {type:'array'});
  const sheet = wb.Sheets[wb.SheetNames[0]];
  rows2D = XLSX.utils.sheet_to_json(sheet, {header:1, raw:true});

  // ---- HEADER extraction (label-based, robust) ----
  parsed = {
    partNumber: null, revision: null, hand: null,
    totalHoles: null, rootWidth: null, fsmLength: null,
    kbSpec: null, pcSpec: null, matrix: null
  };

  // Part / Level / Hand
  const partCell = findCellByRegex(/PART\s*NUMBER\s*\/\s*LEVEL\s*\/\s*HAND/i);
  if(partCell){
    const row = rows2D[partCell.r];
    // find "X / Y / Z" on same row
    const val = row.find((v,idx)=> idx>partCell.c && v && /\/\s*.+\s*\/.+/.test(String(v)));
    if(val){
      const m = String(val).match(/([A-Z0-9\-_]+)\s*\/\s*([A-Z0-9\-_]+)\s*\/\s*([A-Z]+)/i);
      if(m){ parsed.partNumber=m[1]; parsed.revision=m[2]; parsed.hand=m[3]; }
    }
  }

  // TOTAL HOLES COUNT
  const thc = findCellByRegex(/TOTAL\s*HOLES\s*COUNT/i);
  if(thc){
    const row = rows2D[thc.r];
    const n = row.slice(thc.c).find(v=>/\d+/.test(String(v||'')));
    if(n) parsed.totalHoles = parseInt(n,10);
  }

  // ROOT WIDTH OF FSM : Spec- N mm
  const rw = findCellByRegex(/ROOT\s*WIDTH\s*OF\s*FSM/i);
  if(rw){
    const line = rows2D[rw.r].slice(rw.c).join(' ');
    const m = String(line).match(/Spec[-\s]*([0-9.]+)/i);
    if(m) parsed.rootWidth = parseFloat(m[1]);
  }

  // FSM LENGTH : Spec- N mm
  const fl = findCellByRegex(/FSM\s*LENGTH/i);
  if(fl){
    const line = rows2D[fl.r].slice(fl.c).join(' ');
    const m = String(line).match(/Spec[-\s]*([0-9.]+)/i);
    if(m) parsed.fsmLength = parseFloat(m[1]);
  }

  // KB & PC Code : Spec- a / b
  const kbpc = findCellByRegex(/KB\s*&\s*PC\s*Code/i);
  if(kbpc){
    const line = rows2D[kbpc.r].slice(kbpc.c).join(' ');
    const m = String(line).match(/Spec[-\s]*([0-9]+)\s*\/\s*([0-9]+)/i);
    if(m){ parsed.kbSpec=m[1]; parsed.pcSpec=m[2]; }
  }

  // Matrix used:
  const mat = findCellByRegex(/Matrix\s*used/i);
  if(mat){
    const row = rows2D[mat.r];
    const val = row.slice(mat.c).find(v=>v && v!==':');
    if(val) parsed.matrix = String(val);
  }

  // ---- MAIN TABLE (find header row with "X-axis" & "Spec Dia") ----
  const headerIdx = findRowIdxByLabels(['x-axis','spec dia']);
  if(headerIdx === -1){ show('Could not find main table header. Check the template text.', 'warn'); return; }

  const headerRow = rows2D[headerIdx];
  const colIndex = (label) => headerRow.findIndex(v => String(v||'').toLowerCase().includes(label));

  const map = {
    sl    : colIndex('sl'),
    press : colIndex('press'),
    selid : colIndex('sel id'),
    ref   : colIndex('ref'),
    xaxis : colIndex('x-axis'),
    specyz: colIndex('spec (y'),     // "Spec (Y or Z axis)"
    specdia: colIndex('spec dia'),
    valEdge: colIndex('value from hole edge'), // for completeness (usually blank in source)
    actDia : colIndex('actual dia'),
    actYZ  : colIndex('actual y or z'),
    offset : colIndex('offset')
  };

  tableRows = [];
  for(let r=headerIdx+1; r<rows2D.length && tableRows.length<45; r++){
    const row = rows2D[r] || [];
    const cols = [map.sl,map.press,map.selid,map.ref,map.xaxis,map.specyz,map.specdia]
      .map(ci => ci>=0 ? row[ci] : '');

    if(cols.every(v => String(v||'').trim()==='')) break;

    tableRows.push({
      sl: cols[0], press: cols[1], selid: cols[2], ref: cols[3],
      x: parseFloat(cols[4]),
      specYZ: parseFloat(cols[5]),
      specDiaRaw: cols[6]
    });
  }

  // ---- EXTREME ENDS (Dimension verification) ----
  extreme = {
    // order: Web PWX, Web PW, Web PWM, Top Flange PFF, Bottom Flange PFB
    rows: ['PWX','PW','PWM','PFF','PFB'],
    first: {spec:[], actual:[], offset:[]},
    last : {spec:[], actual:[], offset:[]}
  };
  const dimIdx = findRowIdxByLabels(['dimension verification','extreme ends of fsm']);
  if(dimIdx !== -1){
    // Locate headers "First holes of FSM - from front end" and "Last holes ..."
    const hdrRow = rows2D[dimIdx+1] || [];
    // We’ll scan following ~10 rows to fetch 5 lines for X Axis set
    let start = dimIdx;
    // find "X Axis" label row
    const xAxisRowIdx = rows2D.findIndex((r,i)=> i>dimIdx && r.some(v=>String(v||'').toLowerCase().includes('x axis')));
    if(xAxisRowIdx !== -1){
      // After "X Axis" line, the next 5 lines correspond to PWX,PW,PWM,PFF,PFB (per your sample)
      for(let i=0;i<5;i++){
        const r = rows2D[xAxisRowIdx+1+i] || [];
        // try to find 3 numbers on "first holes" side and 3 numbers on "last holes" side
        const nums = r.map(v=>String(v||''));
        // naive numeric scan: pick left block (first holes) and right block (last holes)
        const onlyNums = nums.map(s=> s.match(/^-?\d+(\.\d+)?$/) ? parseFloat(s) : null);
        // heuristic: pick first number trio and last number trio
        const left = onlyNums.filter(n=>n!==null).slice(0,3);
        const right= onlyNums.filter(n=>n!==null).slice(-3);
        extreme.first.spec.push(left[0] ?? null);
        extreme.first.actual.push(left[1] ?? null);
        extreme.first.offset.push(left[2] ?? null);
        extreme.last.spec.push(right[0] ?? null);
        extreme.last.actual.push(right[1] ?? null);
        extreme.last.offset.push(right[2] ?? null);
      }
    }
  }

  // Update sticky header
  if(parsed.partNumber){
    hdrPart.textContent = `PART NUMBER / LEVEL / HAND : ${parsed.partNumber} / ${parsed.revision||'—'} / ${parsed.hand||'—'}`;
  }
  if(parsed.fsmLength) hdrFsmSpec.textContent = `FSM LENGTH : ${parsed.fsmLength} mm`;

  headerOut.textContent = JSON.stringify({
    header: parsed,
    firstRowsPreview: tableRows.slice(0,5),
    extremePreview: extreme
  }, null, 2);

  buildFormBtn.disabled = false;
  clearMsg();
});

/* Build full form (all blocks) */
buildFormBtn.addEventListener('click', ()=>{
  formArea.innerHTML = "";

  // ===== Header / Mandatory =====
  const blkH = block('Header / Mandatory fields');

  // KB & PC + Inspectors + sample times/matrix + totals + root width + fsm length + retake/patrol
  const row1 = row();
  const c1 = col(`
    <div class="small">FSM Serial Number:</div>
  `);
  c1.append(input({placeholder:'Enter FSM Serial No', id:'fsmSerial'}));

  const c2 = col(`
    <div class="small">Inspectors:</div>
  `);
  c2.append(input({placeholder:'Inspector 1', id:'inspector1', className:'input inline'}),
            input({placeholder:'Inspector 2', id:'inspector2', className:'input inline'}));
  row1.append(c1,c2); blkH.append(row1);

  const row2 = row();
  const c3 = col(`<div class="small">KB &amp; PC Code (Spec / Act)</div>`);
  c3.append(input({value:parsed.kbSpec||'', readOnly:true, className:'input inline'}),
           text(' / '),
           input({placeholder:'KB Act', id:'kbAct', className:'input inline'}),
           text('    '),
           input({value:parsed.pcSpec||'', readOnly:true, className:'input inline'}),
           text(' / '),
           input({placeholder:'PC Act', id:'pcAct', className:'input inline'}));

  const c4 = col(`<div class="small">Matrix used:</div>`);
  c4.append(input({value:parsed.matrix||'', id:'matrixUsed'}));
  row2.append(c3,c4); blkH.append(row2);

  const row3 = row();
  const c5 = col(`<div class="small">SAMPLE GIVEN TIME :</div>`);
  c5.append(input({placeholder:'', className:'input inline'}), text(' AM/PM '),
            input({placeholder:'', className:'input inline'}), text(' SAMPLE CLEARED TIME : '),
            input({placeholder:'', className:'input inline'}), text(' AM/PM'));
  const c6 = col(`<div class="small">TOTAL HOLES COUNT :</div>`);
  c6.append(input({value:parsed.totalHoles||'', id:'holesAct'}));
  row3.append(c5,c6); blkH.append(row3);

  const row4 = row();
  const c7 = col(`<div class="small">ROOT WIDTH OF FSM (Spec / Act) mm</div>`);
  c7.append(input({value:parsed.rootWidth||'', readOnly:true, className:'input inline', id:'rootSpec'}),
            input({placeholder:'Act', className:'input inline', id:'rootAct'}));
  const c8 = col(`<div class="small">FSM LENGTH (Spec / Act) mm</div>`);
  c8.append(input({value:parsed.fsmLength||'', readOnly:true, className:'input inline', id:'fsmSpec'}),
            input({placeholder:'Act', className:'input inline', id:'fsmAct'}));
  row4.append(c7,c8); blkH.append(row4);

  const row5 = row();
  const c9 = col(`<div class="small">Retake / Patrol</div>`);
  c9.append(input({placeholder:'Retake', className:'input inline'}),
            input({placeholder:'Patrol', className:'input inline'}));
  row5.append(c9); blkH.append(row5);

  formArea.appendChild(blkH);

  // ===== Main Inspection Table =====
  const blkT = block('Main Inspection Table');
  const table = document.createElement('table'); table.className='table';
  table.innerHTML = `
    <thead><tr>
      <th>Sl No</th><th>Press</th><th>Sel ID</th><th>Ref</th><th>X-axis</th>
      <th>Spec (Y or Z)</th><th>Spec Dia</th>
      <th>Value from Hole edge (Act)</th>
      <th>Actual Dia / Slot (H × W)</th>
      <th>Actual Y or Z</th><th>Offset (Actual - Spec)</th><th>Result</th>
    </tr></thead>
  `;
  const tb = document.createElement('tbody');

  tableRows.forEach((r, idx)=>{
    const tr = document.createElement('tr');

    const ro = v => roCell(v);
    tr.append(td(ro(r.sl)));
    tr.append(td(ro(r.press)));
    tr.append(td(ro(r.selid)));
    tr.append(td(ro(r.ref)));
    tr.append(td(ro(isFinite(r.x)?r.x:'')));
    tr.append(td(ro(isFinite(r.specYZ)?r.specYZ:'')));
    tr.append(td(ro(String(r.specDiaRaw ?? ''))));

    const valEdge = input({className:'input'});   // free numeric
    const spec = parseSpecDia(r.specDiaRaw);

    let diaWrap = document.createElement('div');
    if(spec.isSlot){
      diaWrap.append(
        input({className:'input inline', placeholder:`H (spec ${spec.H})`, 'data-role':'slotH'}),
        text(' × '),
        input({className:'input inline', placeholder:`W (spec ${spec.W})`, 'data-role':'slotW'})
      );
    }else{
      diaWrap.append(input({className:'input inline', placeholder:`Dia (spec ${spec.dia||''})`, 'data-role':'holeDia'}));
    }

    const actYZ = input({className:'input'});
    const offDiv = document.createElement('div');
    const resDiv = document.createElement('div');

    const recalc = ()=>{
      const fsmLen = parseFloat($('#fsmSpec')?.value) || parsed.fsmLength || NaN;
      const offset = (isFinite(parseFloat(actYZ.value)) && isFinite(r.specYZ))
        ? (parseFloat(actYZ.value) - r.specYZ) : NaN;
      offDiv.textContent = isFinite(offset) ? round(offset,2) : '';

      let diaOK = null;
      if(spec.isSlot){
        const aH = parseFloat(diaWrap.querySelector('[data-role="slotH"]')?.value);
        const aW = parseFloat(diaWrap.querySelector('[data-role="slotW"]')?.value);
        const okH = diaSlotOK(spec.H, aH);
        const okW = diaSlotOK(spec.W, aW);
        diaOK = (okH===null || okW===null) ? null : (okH && okW);
      }else{
        const aD = parseFloat(diaWrap.querySelector('[data-role="holeDia"]')?.value);
        diaOK = diaHoleOK(spec.dia, aD);
      }

      const tol = offsetTol(r.x, fsmLen);
      const offOK = isFinite(offset) ? (Math.abs(offset) <= tol) : null;

      let allOk = null;
      if(diaOK===null || offOK===null) allOk = null;
      else allOk = (diaOK && offOK);

      resDiv.textContent = (allOk===null)?'':(allOk?'OK':'NOK');
      resDiv.className   = (allOk===null)?'':(allOk?'ok-cell':'nok-cell');
    };

    [valEdge, actYZ].forEach(i=> i.addEventListener('input', recalc));
    diaWrap.querySelectorAll('input').forEach(i=> i.addEventListener('input', recalc));

    tr.append(td(valEdge));
    tr.append(td(diaWrap));
    tr.append(td(actYZ));
    tr.append(td(offDiv));
    tr.append(td(resDiv));

    tb.appendChild(tr);
  });

  table.appendChild(tb);
  blkT.appendChild(table);
  formArea.appendChild(blkT);

  // ===== Dimension verification (extreme ends) =====
  const blkD = block('Dimension verification for holes at extreme ends of FSM');
  const sub = document.createElement('table'); sub.className='table';
  sub.innerHTML = `
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
  const sbody = document.createElement('tbody');
  const labels = ['PWX','PW','PWM','PFF','PFB'];
  labels.forEach((lab, i)=>{
    const tr = document.createElement('tr');
    tr.append(td(roCell(i<3?'Web':(i===3?'Top Flange':'Bottom Flange'))));
    tr.append(td(roCell(lab)));
    // FIRST
    const fSpec = roCell(extractNum(extreme.first.spec[i]));
    const fAct  = input({className:'input inline', placeholder:'/'});
    const fOff  = roCell(extractNum(extreme.first.offset[i]));
    // LAST
    const lSpec = roCell(extractNum(extreme.last.spec[i]));
    const lAct  = input({className:'input inline', placeholder:'/'});
    const lOff  = roCell(extractNum(extreme.last.offset[i]));

    tr.append(td(fSpec)); tr.append(td(fAct)); tr.append(td(fOff));
    tr.append(td(lSpec)); tr.append(td(lAct)); tr.append(td(lOff));
    sbody.appendChild(tr);
  });
  sub.appendChild(sbody);
  blkD.appendChild(sub);
  formArea.appendChild(blkD);

  // ===== Visual / Binary checks =====
  const blkV = block('Visual / Binary Checks');
  const grid = document.createElement('div'); grid.className='binary';
  const items = [
    'Part no location','Profile cutting','Kink Bending','Machine Error (YES / NO)',
    'Punch Break','Length Variation','Slug mark','Part No. Legibility',
    'Radius crack','Holes Burr','Line mark','Pit Mark'
  ];
  // show as pairs (OK / NOK)
  items.forEach(label=>{
    const wrap = document.createElement('div'); wrap.style.margin='6px 0';
    const span = document.createElement('span'); span.textContent = label + ' : ';
    const ok = document.createElement('button'); ok.textContent='OK/YES';
    const nok = document.createElement('button'); nok.textContent='NOK/NO';
    ok.addEventListener('click',()=>{ ok.classList.add('on-ok'); nok.classList.remove('on-nok'); });
    nok.addEventListener('click',()=>{ nok.classList.add('on-nok'); ok.classList.remove('on-ok'); });
    wrap.append(span, ok, nok);
    grid.appendChild(wrap);
  });
  blkV.appendChild(grid);
  formArea.appendChild(blkV);

  // ===== Remarks + Signatures =====
  const blkR = block('Remarks / Details of issue');
  const ta = document.createElement('textarea'); ta.className='input'; ta.style.minHeight='90px';
  blkR.appendChild(ta);
  const rowS = row();
  const left = col(`<div class="small">Prodn Incharge</div>`); left.append(input({className:'input', id:'prodIncharge'}));
  const right= col(`<div class="small">QC Inspector</div>`);  right.append(input({className:'input', id:'qcInspector'}));
  rowS.append(left,right); blkR.append(rowS);
  formArea.appendChild(blkR);

  exportPdfBtn.disabled = false;
  show('Form created. Enter actuals and export to PDF.');
});

/* Export to PDF */
exportPdfBtn.addEventListener('click', async ()=>{
  exportPdfBtn.disabled = true; exportPdfBtn.textContent='Generating PDF...';
  try{
    const clone = document.createElement('div');
    clone.style.width='900px'; clone.style.padding='12px';
    const hdr = document.querySelector('header').cloneNode(true);
    clone.appendChild(hdr);
    const fa = formArea.cloneNode(true);
    fa.querySelectorAll('input, textarea, button').forEach(el=>{
      if(el.tagName==='BUTTON'){ el.remove(); return; }
      const s = document.createElement('span');
      s.textContent = el.value || el.placeholder || '';
      el.parentNode && el.parentNode.replaceChild(s, el);
    });
    clone.appendChild(fa);
    clone.style.position='fixed'; clone.style.left='-2000px';
    document.body.appendChild(clone);

    const canvas = await html2canvas(clone, {scale:2, useCORS:true});
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF('p','mm','a4');
    const pageW = pdf.internal.pageSize.getWidth();
    const imgW = pageW - 20;
    const imgH = canvas.height * imgW / canvas.width;
    pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 10, 10, imgW, imgH);
    const d = new Date();
    const fname = `${parsed.partNumber||'PART'}_${parsed.revision||'X'}_${parsed.hand||'H'}_Inspection_${d.toISOString().slice(0,10)}.pdf`;
    pdf.save(fname);
    show('PDF exported: '+fname);
  }catch(e){
    console.error(e);
    alert('PDF export failed: '+e.message);
  }finally{
    exportPdfBtn.disabled=false; exportPdfBtn.textContent='Export to PDF';
  }
});

/* DOM helpers */
function block(title){
  const d = document.createElement('div'); d.className='form-block';
  const h = document.createElement('div'); h.className='block-title'; h.textContent = title;
  d.appendChild(h); return d;
}
function row(){ const d=document.createElement('div'); d.className='form-row'; return d; }
function col(html){ const d=document.createElement('div'); d.className='col'; if(html) d.innerHTML=html; return d; }
function input({value="", placeholder="", className="input", readOnly=false, id, ...attrs}={}){
  const i=document.createElement('input'); i.type='text'; i.value=value??""; i.placeholder=placeholder; i.className=className; i.readOnly=!!readOnly; if(id) i.id=id;
  Object.entries(attrs).forEach(([k,v])=> i.setAttribute(k, v));
  return i;
}
function text(t){ return document.createTextNode(t); }
function roCell(v){ const d=document.createElement('div'); d.textContent = v ?? ""; return d; }
function td(inner){ const td=document.createElement('td'); if(inner instanceof Node) td.appendChild(inner); else td.textContent=inner??""; return td; }
function extractNum(v){ return (v==null||v==="--") ? '' : (isFinite(parseFloat(v))?String(v):''); }
