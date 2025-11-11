/* Excel → Form (Option A for slots: first number = vertical) */
const fileInput = document.getElementById('fileInput');
const parseExcelBtn = document.getElementById('parseExcelBtn');
const buildFormBtn   = document.getElementById('buildFormBtn');
const exportPdfBtn   = document.getElementById('exportPdfBtn');
const headerOut = document.getElementById('headerOut');
const formArea  = document.getElementById('formArea');
const warnings  = document.getElementById('warnings');
const hdrPart   = document.getElementById('hdrPart');
const hdrFsmSpec= document.getElementById('hdrFsmSpec');

let workbook, rows, parsed = {};
let tableRows = []; // up to 45

/* Helpers */
const show = (m, type="info")=>{
  warnings.style.padding="8px"; warnings.style.borderRadius="6px"; warnings.style.marginTop="6px";
  warnings.textContent=m;
  if(type==="warn"){warnings.style.background="#fff6cc";warnings.style.color="#ff8c00";}
  else{warnings.style.background="#e8f4ff";warnings.style.color="#005fcc";}
};
const clearMsg=()=>{warnings.textContent="";warnings.removeAttribute('style');};
const round=(n,d=2)=>Math.round(n*10**d)/10**d;

/* Slot parse — Option A */
function parseSpecDia(cell) {
  if(cell==null) return {isSlot:false, dia:NaN, v:NaN, h:NaN, raw:""};
  const raw = String(cell).trim();
  const m = raw.match(/^(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)/);
  if(m){
    // Option A: first number is vertical
    return {isSlot:true, v:parseFloat(m[1]), h:parseFloat(m[2]), raw};
  }
  const d = parseFloat(raw);
  return {isSlot:false, dia:isFinite(d)?d:NaN, v:NaN, h:NaN, raw};
}

/* Dia tolerance */
function diaOk(spec, actual){
  if(isNaN(spec) || isNaN(actual)) return null; // not decidable
  if (spec <= 10.7) return (actual >= spec && actual <= spec + 0.4);
  if (spec >= 11.7) return (actual >= spec && actual <= spec + 0.5);
  // between 10.7 and 11.7 → +0.5 (your note)
  return (actual >= spec && actual <= spec + 0.5);
}
/* Slot tolerance: -0 / +0.5 for both height & width */
const slotOk=(s,a)=> (isFinite(s)&&isFinite(a)) ? (a>=s && a<=s+0.5) : null;

/* Offset tolerance vs X position */
function offsetTol(x, fsmLen){
  if(!isFinite(x)||!isFinite(fsmLen)) return 1;
  if(x<=200 || x>=fsmLen-200) return 1.5;
  return 1;
}

/* Parse Excel to 2D array of strings */
parseExcelBtn.addEventListener('click', async ()=>{
  const f = fileInput.files?.[0];
  if(!f){ alert("Select the Excel file (.xlsx)"); return; }

  const data = await f.arrayBuffer();
  workbook = XLSX.read(data, {type:'array'});
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  rows = XLSX.utils.sheet_to_json(sheet, {header:1, raw:true}); // 2D array

  // Extract header info by searching strings
  parsed = {
    partNumber:null, revision:null, hand:null,
    totalHoles:null, rootWidth:null, fsmLength:null,
    kbSpec:null, pcSpec:null, matrix:null
  };

  for(const r of rows){
    const line = (r.join(' ')||'').toString();

    if(/PART\s*NUMBER\s*\/\s*LEVEL\s*\/\s*HAND/i.test(line)){
      // The value is usually in the next non-empty cell(s) in that row
      const txt = r.find((c,idx)=> idx>0 && c && typeof c ==='string' && /[A-Z0-9]/i.test(c));
      if(txt){
        // e.g. "FH140813 / O2 / LH"
        const m = txt.match(/([A-Z0-9\-_]+)\s*\/\s*([A-Z0-9\-_]+)\s*\/\s*([A-Z]+)/i);
        if(m){ parsed.partNumber=m[1]; parsed.revision=m[2]; parsed.hand=m[3]; }
      }
    }
    if(/TOTAL\s*HOLES\s*COUNT/i.test(line)){
      const n = String(line).match(/COUNT\s*[:\-]?\s*(\d+)/i);
      if(n) parsed.totalHoles = parseInt(n[1],10);
    }
    if(/ROOT\s*WIDTH\s*OF\s*FSM/i.test(line)){
      const m = String(line).match(/Spec[-\s]*([0-9.]+)/i);
      if(m) parsed.rootWidth = parseFloat(m[1]);
    }
    if(/FSM\s*LENGTH/i.test(line)){
      const m = String(line).match(/Spec[-\s]*([0-9.]+)/i);
      if(m) parsed.fsmLength = parseFloat(m[1]);
    }
    if(/KB\s*&\s*PC\s*Code/i.test(line)){
      const m = String(line).match(/Spec[-\s]*([0-9]+)\s*\/\s*([0-9]+)/i);
      if(m){ parsed.kbSpec = m[1]; parsed.pcSpec = m[2]; }
    }
    if(/Matrix\s*used/i.test(line)){
      const m = String(line).match(/Matrix\s*used\s*[:\-]?\s*([A-Za-z0-9]+)/i);
      if(m) parsed.matrix = m[1];
    }
  }

  // Find table header row (contains "Spec Dia")
  let headerIdx = rows.findIndex(r=> r.some(c => String(c||'').toLowerCase().includes('spec dia')));
  if(headerIdx === -1){
    show("Could not find table header. Please ensure the template matches.", "warn");
  }

  // Map columns by title
  const headerRow = rows[headerIdx] || [];
  const findCol = (label)=> headerRow.findIndex(c => String(c||'').toLowerCase().includes(label));
  const colMap = {
    sl: findCol('sl'),
    press: findCol('press'),
    selid: findCol('sel id'),
    ref: findCol('ref'),
    x: findCol('x-axis'),
    specYZ: findCol('spec') ,       // "Spec (Y or Z axis)"
    specDia: findCol('spec dia'),
    // the rest will be built later (value from hole edge, actual dia, actual yz, offset)
  };

  // Collect up to 45 rows or until first fully-empty line
  tableRows = [];
  for(let i = headerIdx+1; i < rows.length && tableRows.length < 45; i++){
    const r = rows[i] || [];
    const firstCells = [colMap.sl,colMap.press,colMap.selid,colMap.ref,colMap.x,colMap.specYZ,colMap.specDia]
      .map(ci => ci>=0 ? r[ci] : '')
      .map(v => (v===undefined || v===null) ? '' : v);

    // stop if completely empty
    if(firstCells.every(v => String(v).trim()==='')) break;

    tableRows.push({
      sl: firstCells[0],
      press: firstCells[1],
      selid: firstCells[2],
      ref: firstCells[3],
      x: parseFloat(firstCells[4]) || NaN,
      specYZ: parseFloat(firstCells[5]),
      specDiaRaw: firstCells[6] // could be "9*12"
    });
  }

  headerOut.textContent = JSON.stringify({header:parsed, rows:tableRows.slice(0,5)}, null, 2);
  if(parsed.partNumber) hdrPart.textContent = `PART NUMBER / LEVEL / HAND : ${parsed.partNumber} / ${parsed.revision||'—'} / ${parsed.hand||'—'}`;
  if(parsed.fsmLength) hdrFsmSpec.textContent = `FSM LENGTH : ${parsed.fsmLength} mm`;
  buildFormBtn.disabled = false;
  clearMsg();
});

/* Build form */
buildFormBtn.addEventListener('click', ()=>{
  formArea.innerHTML = "";
  clearMsg();

  // --- Header / Mandatory ---
  const blkH = document.createElement('div');
  blkH.className = "form-block";
  blkH.innerHTML = `<strong>Header / Mandatory fields</strong>`;
  const r1 = document.createElement('div'); r1.className='form-row';

  const c1 = document.createElement('div'); c1.className='col';
  c1.innerHTML = `<div class="small">KB &amp; PC Code (Spec / Act)</div>`;
  const kbSpec = mkInput({value:parsed.kbSpec||"", readOnly:true, className:'input-inline'});
  const kbAct  = mkInput({placeholder:'KB Act', className:'input-inline', id:'kbAct'});
  const pcSpec = mkInput({value:parsed.pcSpec||"", readOnly:true, className:'input-inline'});
  const pcAct  = mkInput({placeholder:'PC Act', className:'input-inline', id:'pcAct'});
  c1.append(kbSpec, document.createTextNode(' / '), kbAct, document.createTextNode('    '),
            pcSpec, document.createTextNode(' / '), pcAct);

  const c2 = document.createElement('div'); c2.className='col';
  c2.innerHTML = `<div class="small">ROOT WIDTH OF FSM (Spec / Act) mm  |  FSM LENGTH (Spec / Act) mm</div>`;
  const rootSpec = mkInput({value:parsed.rootWidth||"", readOnly:true, className:'input-inline', id:'rootSpec'});
  const rootAct  = mkInput({placeholder:'Act', className:'input-inline', id:'rootAct'});
  const fsmSpec  = mkInput({value:parsed.fsmLength||"", readOnly:true, className:'input-inline', id:'fsmSpec'});
  const fsmAct   = mkInput({placeholder:'Act', className:'input-inline', id:'fsmAct'});
  c2.append(rootSpec, document.createTextNode(' '), rootAct,
            document.createTextNode('    '), fsmSpec, document.createTextNode(' '), fsmAct);

  r1.append(c1,c2);
  blkH.append(r1);
  formArea.appendChild(blkH);

  // --- Main Inspection Table ---
  const blkT = document.createElement('div'); blkT.className='form-block';
  blkT.innerHTML = `<strong>Main Inspection Table</strong>`;
  const table = document.createElement('table'); table.className='table';
  table.innerHTML = `
    <thead><tr>
      <th>Sl No</th><th>Press</th><th>Sel ID</th><th>Ref</th>
      <th>X-axis</th><th>Spec (Y or Z)</th><th>Spec Dia</th>
      <th>Value from Hole edge (Act)</th>
      <th>Actual Dia / Slot (H x W)</th>
      <th>Actual Y or Z</th><th>Offset (Actual - Spec)</th><th>Result</th>
    </tr></thead>
  `;
  const tb = document.createElement('tbody');

  tableRows.forEach((r, idx)=>{
    const tr = document.createElement('tr');

    // read-only first 7
    const ro = v => roCell(v);
    tr.append(tdWrap(ro(r.sl)));
    tr.append(tdWrap(ro(r.press)));
    tr.append(tdWrap(ro(r.selid)));
    tr.append(tdWrap(ro(r.ref)));
    tr.append(tdWrap(ro(isFinite(r.x)?r.x:"")));
    tr.append(tdWrap(ro(isFinite(r.specYZ)?r.specYZ:"")));

    // spec dia cell (read-only text; show raw like "9*12")
    tr.append(tdWrap(ro(r.specDiaRaw)));

    // Value from hole edge (Act)
    const valEdge = mkInput({className:'input', placeholder:''});

    // Actual dia / slot
    const specDia = parseSpecDia(r.specDiaRaw);
    let diaContainer = document.createElement('div');

    if(specDia.isSlot){
      // Two inputs: Height (vertical), Width (horizontal)
      const dH = mkInput({className:'input-inline', placeholder:`H (spec ${specDia.v})`});
      const dW = mkInput({className:'input-inline', placeholder:`W (spec ${specDia.h})`});
      dH.dataset.type="slotH"; dW.dataset.type="slotW";
      diaContainer.append(dH, document.createTextNode(' × '), dW);
    }else{
      const d = mkInput({className:'input-inline', placeholder:`Dia (spec ${specDia.dia||''})`});
      d.dataset.type="holeDia";
      diaContainer.append(d);
    }

    // Actual Y/Z, Offset, Result
    const actYZ = mkInput({className:'input'});
    const offTD = document.createElement('div');
    const resTD = document.createElement('div');

    // events → recalc
    const onChange = ()=> {
      const fsmLen = parseFloat(document.getElementById('fsmSpec')?.value) || parsed.fsmLength || NaN;
      const x = r.x;
      const specY = r.specYZ;
      const actualY = parseFloat(actYZ.value);

      // Offset = Actual Y/Z - Spec Y/Z   (header says: Actual - Spec)
      let offset = (isFinite(actualY) && isFinite(specY)) ? (actualY - specY) : NaN;
      offTD.textContent = isFinite(offset)? round(offset,2) : '';

      // Dia checks
      let diaOK = null;
      if(specDia.isSlot){
        const dH = diaContainer.querySelector('[data-type="slotH"]');
        const dW = diaContainer.querySelector('[data-type="slotW"]');
        const aH = parseFloat(dH.value), aW = parseFloat(dW.value);
        const okH = slotOk(specDia.v, aH);
        const okW = slotOk(specDia.h, aW);
        diaOK = (okH === null || okW === null) ? null : (okH && okW);
      }else{
        const d = diaContainer.querySelector('[data-type="holeDia"]');
        const a = parseFloat(d.value);
        diaOK = diaOk(specDia.dia, a);
      }

      // Offset tolerance
      const tol = offsetTol(x, fsmLen);
      const offOK = (isFinite(offset)) ? (Math.abs(offset) <= tol) : null;

      // Result
      let allOk = null;
      if(diaOK === null || offOK === null) allOk = null;
      else allOk = (diaOK && offOK);

      resTD.textContent = (allOk===null) ? '' : (allOk ? 'OK':'NOK');
      resTD.className = (allOk===null) ? '' : (allOk ? 'ok-cell':'nok-cell');
    };

    // wire
    [valEdge, actYZ].forEach(i=> i.addEventListener('input', onChange));
    diaContainer.querySelectorAll('input').forEach(i=> i.addEventListener('input', onChange));

    tr.append(tdWrap(valEdge));
    tr.append(tdWrap(diaContainer));
    tr.append(tdWrap(actYZ));
    tr.append(tdWrap(offTD));
    tr.append(tdWrap(resTD));

    tb.appendChild(tr);
  });

  table.appendChild(tb);
  blkT.appendChild(table);
  formArea.appendChild(blkT);

  // Enable export
  exportPdfBtn.disabled = false;
  show("Editable form ready. Enter Actuals, then Export to PDF.");
});

/* Export to PDF */
exportPdfBtn.addEventListener('click', async ()=>{
  exportPdfBtn.disabled = true;
  exportPdfBtn.textContent = "Generating PDF...";
  try{
    const clone = document.createElement('div');
    clone.style.width="900px"; clone.style.padding="12px";
    // header
    const hdr = document.querySelector('header').cloneNode(true);
    clone.appendChild(hdr);

    // content clone (inputs → spans)
    const fa = formArea.cloneNode(true);
    fa.querySelectorAll('input').forEach(n=>{
      const s = document.createElement('span');
      s.textContent = n.value || (n.placeholder || '');
      n.parentNode && n.parentNode.replaceChild(s, n);
    });
    clone.appendChild(fa);

    clone.style.position="fixed"; clone.style.left="-2000px";
    document.body.appendChild(clone);

    const canvas = await html2canvas(clone, {scale:2, useCORS:true});
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF('p','mm','a4');
    const pageW = pdf.internal.pageSize.getWidth();
    const imgW = pageW - 20;
    const imgH = canvas.height * imgW / canvas.width;
    const img = canvas.toDataURL('image/png');
    pdf.addImage(img, 'PNG', 10, 10, imgW, imgH);
    const now = new Date();
    const fname = `${(parsed.partNumber||'PART')}_${(parsed.revision||'X')}_${(parsed.hand||'H')}_Inspection_${now.toISOString().slice(0,10)}.pdf`;
    pdf.save(fname);
    show("PDF exported: " + fname);
  }catch(e){
    console.error(e);
    alert("PDF export failed: " + e.message);
  }finally{
    exportPdfBtn.disabled = false;
    exportPdfBtn.textContent = "Export to PDF";
  }
});

/* Tiny DOM helpers */
function mkInput({value="", placeholder="", className="input", readOnly=false, id}={}){
  const i = document.createElement('input');
  i.type = 'text'; i.value = value ?? ""; i.placeholder = placeholder;
  i.className = className; i.readOnly = !!readOnly;
  if(id) i.id = id;
  return i;
}
function roCell(text){
  const div = document.createElement('div'); div.textContent = text ?? "";
  return div;
}
function tdWrap(inner){
  const td = document.createElement('td');
  if(inner instanceof Node) td.appendChild(inner); else td.textContent = inner ?? "";
  return td;
}
