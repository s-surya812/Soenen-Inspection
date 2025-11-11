/* Excel-driven Soenen Inspection */
const fileInput = document.getElementById("fileInput");
const parseExcelBtn = document.getElementById("parseExcelBtn");
const generateFormBtn = document.getElementById("generateFormBtn");
const exportPdfBtn = document.getElementById("exportPdfBtn");
const headerOut = document.getElementById("headerOut");
const warnings = document.getElementById("warnings");
const formArea = document.getElementById("formArea");
const hdrPart = document.getElementById("hdrPart");
const hdrFsmSpec = document.getElementById("hdrFsmSpec");

let excelFile = null;
let parsedHeader = {
  partNumber: null, revision: null, hand: null,
  fsmLength: null, rootWidth: null, kbSpec:null, pcSpec:null,
  holesSpec: null
};
let specRows = [];      // [{sl, press, selId, ref, x, specYZ, specDia}, ...] length up to 45
let firstHoleSpecs = []; // 5
let lastHoleSpecs = [];  // 5

parseExcelBtn.addEventListener("click", async () => {
  if (!fileInput.files?.length){ alert("Please choose an Excel file (.xlsx)."); return; }
  excelFile = fileInput.files[0];
  parseExcelBtn.disabled = true; parseExcelBtn.textContent = "Parsing...";
  try {
    const data = await excelFile.arrayBuffer();
    const wb = XLSX.read(data, { type: "array" });

    // Prefer "Check sheet format" if found, else the first sheet
    const sheetName = wb.SheetNames.find(n => /check\s*sheet\s*format/i.test(n)) || wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];

    // Convert to 2D array
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
    // Try also to read "Filled sheet example" to infer some samples (optional)
    // const wsFilled = wb.SheetNames.find(n => /filled/i.test(n)) ? wb.Sheets[wb.SheetNames.find(n => /filled/i.test(n))] : null;

    // Parse header block & table
    parseHeaderBlock(rows);
    parseMainTable(rows);
    parseFirstLastHoles(rows);

    // Reflect in UI
    updateHeaderUI();
    headerOut.textContent =
      JSON.stringify({ sheet: sheetName, parsedHeader, firstHoleSpecs, lastHoleSpecs, tablePreview: specRows.slice(0,5) }, null, 2);

    // Enable form creation
    generateFormBtn.disabled = false;
    showMessage("Excel parsed. Click 'Create Editable Form' to build the form.", "info");
  } catch (e) {
    console.error(e);
    alert("Failed to parse Excel: " + e.message);
  } finally {
    parseExcelBtn.disabled = false; parseExcelBtn.textContent = "Load & Parse Excel";
  }
});

function updateHeaderUI(){
  const pn = parsedHeader.partNumber || "—";
  const rv = parsedHeader.revision || "—";
  const hd = parsedHeader.hand || "—";
  const fsm = parsedHeader.fsmLength ?? "—";
  hdrPart.textContent = `PART NUMBER / LEVEL / HAND : ${pn} / ${rv} / ${hd}`;
  hdrFsmSpec.textContent = `FSM LENGTH : ${fsm} mm`;
}

function norm(s){ return String(s||"").trim(); }
function eq(a,b){ return norm(a).toLowerCase()===norm(b).toLowerCase(); }
function includes(hay, needle){ return norm(hay).toLowerCase().includes(norm(needle).toLowerCase()); }

function parseHeaderBlock(rows){
  // scan first ~40 rows for labels
  for(let r=0;r<Math.min(rows.length,80);r++){
    const line = rows[r].map(norm).join(" ");
    // Part / Level / Hand (robust)
    if (/part\s*number/i.test(line) && /level/i.test(line) && /hand/i.test(line)){
      const flat = line.replace(/\s+/g," ");
      // e.g., "PART NUMBER / LEVEL / HAND : 776929 / C / LH"
      const m = flat.match(/hand\s*:?\s*(.+)$/i);
      if (m){
        const tail = m[1].split("/");
        parsedHeader.partNumber = norm(tail[0]||"");
        parsedHeader.revision   = norm(tail[1]||"");
        parsedHeader.hand       = norm(tail[2]||"");
      }
    }
    // FSM LENGTH
    if (includes(line,"fsm length")){
      const num = (line.match(/([0-9]+(?:\.[0-9]+)?)/) || [])[1];
      if (num) parsedHeader.fsmLength = parseFloat(num);
    }
    // ROOT WIDTH
    if (includes(line,"root width")){
      const num = (line.match(/([0-9]+(?:\.[0-9]+)?)/) || [])[1];
      if (num) parsedHeader.rootWidth = parseFloat(num);
    }
    // KB & PC spec e.g. "KB & PC Code : Spec- 5 / 1"
    if (includes(line,"kb") && includes(line,"pc")){
      const mm = line.match(/spec[^0-9]*([0-9]+)\s*\/\s*([0-9]+)/i);
      if (mm){ parsedHeader.kbSpec = mm[1]; parsedHeader.pcSpec = mm[2]; }
    }
    // Holes Count (Spec)
    if (includes(line,"total holes") || includes(line,"holes count")){
      const num = (line.match(/([0-9]{1,3})/) || [])[1];
      if (num) parsedHeader.holesSpec = Number(num);
    }
  }
}

function parseMainTable(rows){
  // Find header row of main table
  // Expected headers (order): Sl No | Press | Sel ID | Ref | X-axis | Spec (Y or Z) | Spec Dia
  let headerRowIdx = -1;
  for (let r=0;r<rows.length;r++){
    const line = rows[r].map(norm).join(" ").toLowerCase();
    if (includes(line,"sl") && includes(line,"press") && includes(line,"sel") && includes(line,"ref") &&
        includes(line,"x") && (includes(line,"y or z")||includes(line,"y/z")) && includes(line,"spec") && includes(line,"dia")){
      headerRowIdx = r; break;
    }
  }
  if (headerRowIdx===-1){
    // fallback: try a looser match on "Spec Dia"
    for (let r=0;r<rows.length;r++){
      const line = rows[r].map(norm).join(" ").toLowerCase();
      if (includes(line,"spec dia")) { headerRowIdx = r; break; }
    }
  }
  if (headerRowIdx===-1){
    console.warn("Main table header not found — building blank 45 rows.");
    specRows = [];
    for (let i=1;i<=45;i++){
      specRows.push({sl:i,press:"",selId:"",ref:"",x:"",specYZ:"",specDia:""});
    }
    return;
  }

  // Map columns from header row
  const hdr = rows[headerRowIdx].map(norm);
  const col = (label) => hdr.findIndex(h => includes(h,label));

  const idxSl     = col("sl");
  const idxPress  = col("press");
  const idxSel    = (col("sel id")!==-1)? col("sel id") : col("sel");
  const idxRef    = col("ref");
  const idxX      = col("x");
  const idxSpecYZ = (col("y or z")!==-1)? col("y or z") : (col("y/z")!==-1? col("y/z") : col("spec (y"));
  const idxDia    = col("spec dia");

  const start = headerRowIdx + 1;
  const collected = [];
  for (let r=start; r<rows.length && collected.length<45; r++){
    const row = rows[r];
    const sl = row[idxSl] ?? "";
    const press = row[idxPress] ?? "";
    const selId = row[idxSel] ?? "";
    const ref   = row[idxRef] ?? "";
    const x     = row[idxX] ?? "";
    const yz    = row[idxSpecYZ] ?? "";
    const dia   = row[idxDia] ?? "";

    const allEmpty = [sl,press,selId,ref,x,yz,dia].every(v => norm(v)==="");
    if (allEmpty) {
      // still push blank to keep row count to 45 (allow gaps)
      collected.push({sl: collected.length+1, press:"", selId:"", ref:"", x:"", specYZ:"", specDia:""});
      continue;
    }
    // Normalize numbers
    const toNum = v => {
      const n = parseFloat(String(v).replace(/[^\d.]/g,""));
      return isNaN(n) ? "" : n;
    }
    collected.push({
      sl: (Number(sl) || collected.length+1),
      press: norm(press),
      selId: norm(selId),
      ref: norm(ref),
      x: toNum(x),
      specYZ: toNum(yz),
      specDia: toNum(dia)
    });
  }

  // Pad up to 45 rows
  while (collected.length < 45) {
    collected.push({sl: collected.length+1, press:"", selId:"", ref:"", x:"", specYZ:"", specDia:""});
  }
  specRows = collected.slice(0,45);
}

function parseFirstLastHoles(rows){
  // Look for "First holes" section: take next 5 numeric cells as specs
  firstHoleSpecs = [];
  lastHoleSpecs  = [];

  const pick5NumbersAfter = (anchorRegex) => {
    let found = -1;
    for (let r=0; r<rows.length; r++){
      const line = rows[r].map(norm).join(" ");
      if (anchorRegex.test(line)) { found = r; break; }
    }
    if (found === -1) return [];
    const nums = [];
    for (let r=found; r<Math.min(found+12, rows.length); r++){
      for (const c of rows[r]){
        const m = String(c).match(/^([0-9]+(?:\.[0-9]+)?)$/);
        if (m) { nums.push(parseFloat(m[1])); if (nums.length===5) return nums; }
      }
    }
    return nums.slice(0,5);
  };

  firstHoleSpecs = pick5NumbersAfter(/first\s*holes/i);
  lastHoleSpecs  = pick5NumbersAfter(/last\s*holes/i);

  // Pad to 5 if fewer
  while(firstHoleSpecs.length<5) firstHoleSpecs.push("");
  while(lastHoleSpecs.length<5) lastHoleSpecs.push("");
}

/* ===== Build Form ===== */
generateFormBtn.addEventListener("click", () => {
  formArea.innerHTML = "";
  warnings.innerHTML = "";

  buildHeaderInputs();
  buildMainTable();
  buildFirstHolesSection();
  buildLastHolesSection();
  buildRootAndFlangeSection();
  buildPartNoLocation();
  buildBinaryChecks();
  buildRemarksAndSign();
  attachValidationListeners();

  exportPdfBtn.disabled = false;
  showMessage("Editable form created. Fill Actual values and mandatory fields. Then export to PDF.", "info");
});

/* Small input factory */
function makeInput(attrs = {}) {
  const inp = document.createElement("input");
  inp.type = attrs.type || "text";
  inp.value = attrs.value ?? "";
  inp.className = attrs.className || "input-small";
  if (attrs.placeholder) inp.placeholder = attrs.placeholder;
  if (attrs.step) inp.step = attrs.step;
  if (attrs.readOnly) inp.readOnly = true;
  if (attrs.size) inp.size = attrs.size;
  if (attrs.id) inp.id = attrs.id;
  return inp;
}

/* ===== Header Block ===== */
function buildHeaderInputs(){
  const blk = document.createElement("div");
  blk.className="form-block";
  blk.innerHTML = `<strong>Header / Mandatory fields</strong>`;

  // Row 1: Serial + Inspectors
  const row1 = document.createElement("div"); row1.className="form-row";
  const c1 = document.createElement("div"); c1.className="col";
  c1.innerHTML = `<div class="small">FSM Serial Number:</div>`;
  c1.appendChild(makeInput({ id:"fsmSerial", className:"input-small", placeholder:"Enter FSM Serial No" }));

  const c2 = document.createElement("div"); c2.className="col";
  c2.innerHTML = `<div class="small">INSPECTORS:</div>`;
  c2.append(
    makeInput({ id:"inspector1", className:"input-small", placeholder:"Inspector 1" }),
    document.createTextNode(" "),
    makeInput({ id:"inspector2", className:"input-small", placeholder:"Inspector 2" })
  );
  row1.append(c1,c2); blk.appendChild(row1);

  // Row 2: Holes + Matrix
  const row2 = document.createElement("div"); row2.className="form-row";
  const c3 = document.createElement("div"); c3.className="col";
  c3.innerHTML = `<div class="small">TOTAL HOLES COUNT (Spec / Act):</div>`;
  c3.append(
    makeInput({ id:"holesSpec", className:"input-small", readOnly:true, value: parsedHeader.holesSpec ?? "" }),
    document.createTextNode(" "),
    makeInput({ id:"holesAct", className:"input-small", placeholder:"Act (mandatory)" }),
  );
  const c4 = document.createElement("div"); c4.className="col";
  c4.innerHTML = `<div class="small">Matrix used:</div>`;
  c4.appendChild(makeInput({ id:"matrixUsed", className:"input-small", placeholder:"Matrix" }));
  row2.append(c3,c4); blk.appendChild(row2);

  // Row 3: KB & PC (Spec readonly, Act editable; compact)
  const row3 = document.createElement("div"); row3.className="form-row";
  const kbcol = document.createElement("div"); kbcol.className="col";
  kbcol.innerHTML = `<div class="small">KB & PC Code (Spec / Act)</div>`;
  const kbSpec = makeInput({ id:"kbSpec", className:"input-small", readOnly:true, size:3, value: parsedHeader.kbSpec ?? "" });
  const kbAct  = makeInput({ id:"kbAct",  className:"input-small", size:3, placeholder:"Act" });
  const pcSpec = makeInput({ id:"pcSpec", className:"input-small", readOnly:true, size:3, value: parsedHeader.pcSpec ?? "" });
  const pcAct  = makeInput({ id:"pcAct",  className:"input-small", size:3, placeholder:"Act" });
  kbcol.append(kbSpec, document.createTextNode(" / "), kbAct, document.createTextNode("    "), pcSpec, document.createTextNode(" / "), pcAct);
  row3.appendChild(kbcol); blk.appendChild(row3);

  // Row 4: Root Width & FSM Length (Spec readonly, Act editable)
  const row4 = document.createElement("div"); row4.className="form-row";

  const rwcol = document.createElement("div"); rwcol.className="col";
  rwcol.innerHTML = `<div class="small">ROOT WIDTH OF FSM (Spec / Act) mm</div>`;
  rwcol.append(
    makeInput({ id:"rootSpec", className:"input-small", readOnly:true, value: parsedHeader.rootWidth ?? "", size:6 }),
    document.createTextNode(" "),
    makeInput({ id:"rootAct", className:"input-small", step:"0.01", placeholder:"Act (mm)", size:6 })
  );

  const fsmcol = document.createElement("div"); fsmcol.className="col";
  fsmcol.innerHTML = `<div class="small">FSM LENGTH (Spec / Act) mm</div>`;
  fsmcol.append(
    makeInput({ id:"fsmSpec", className:"input-small", readOnly:true, value: parsedHeader.fsmLength ?? "", size:6 }),
    document.createTextNode(" "),
    makeInput({ id:"fsmAct", className:"input-small", placeholder:"Act (mm)", size:6 })
  );

  row4.append(rwcol, fsmcol); blk.appendChild(row4);

  formArea.appendChild(blk);
}

/* ===== Main Table ===== */
function buildMainTable(){
  const blk = document.createElement("div"); blk.className="form-block";
  blk.innerHTML = `<strong>Main Inspection Table (1–45)</strong>`;
  const table = document.createElement("table"); table.className="table";
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
  const tbody = document.createElement("tbody");

  for(let i=0;i<45;i++){
    const r = specRows[i] || { sl:i+1, press:"", selId:"", ref:"", x:"", specYZ:"", specDia:"" };
    const tr = document.createElement("tr");

    // Non-editable Spec columns
    tr.appendChild(tdText(r.sl));
    tr.appendChild(tdText(r.press));
    tr.appendChild(tdText(r.selId));
    tr.appendChild(tdText(r.ref));
    tr.appendChild(tdText(r.x));
    tr.appendChild(tdText(r.specYZ));
    tr.appendChild(tdText(r.specDia));

    // Editable Actual columns
    const valEdge = makeInput({ className:"input-small" });
    const actualDia = makeInput({ className:"input-small" });
    const actualYZ  = makeInput({ className:"input-small" });
    const offsetDiv = document.createElement("div");
    const resultDiv = document.createElement("div");

    [valEdge, actualDia, actualYZ].forEach(inp => inp.addEventListener("input", () => recalcRowAndMark(tr)));

    tr.appendChild(tdWrap(valEdge));
    tr.appendChild(tdWrap(actualDia));
    tr.appendChild(tdWrap(actualYZ));
    tr.appendChild(tdWrap(offsetDiv));
    tr.appendChild(tdWrap(resultDiv));

    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  blk.appendChild(table);
  formArea.appendChild(blk);
}
function tdText(v){ const td=document.createElement("td"); td.textContent=(v===0?0:(v||"")); return td; }
function tdWrap(el){ const td=document.createElement("td"); td.appendChild(el); return td; }

function recalcRowAndMark(tr){
  const tds = tr.querySelectorAll("td");

  // Read spec values (non-editable text)
  const x      = parseFloat(tds[4].textContent||"");
  const specYZ = parseFloat(tds[5].textContent||"");
  const specDia= parseFloat(tds[6].textContent||"");

  // Actuals
  const valEdge = parseFloat(tds[7].querySelector("input")?.value || "");
  const actDia  = parseFloat(tds[8].querySelector("input")?.value || "");
  const actYZ   = parseFloat(tds[9].querySelector("input")?.value || "");

  const offsetNode = tds[10];
  const resultNode = tds[11];

  // Offset = Spec(Y/Z) + (Spec Dia)/2
  let offset = (isFinite(specYZ) && isFinite(specDia)) ? (specYZ + specDia/2) : null;
  offsetNode.textContent = (offset==null? "" : round(offset,2));

  // Tolerances:
  // Offset tol: ±1 mm normally; ±2 mm for x<=200 or x >= (FSM Length - 200)
  const fsmLen = parseFloat(document.getElementById("fsmSpec")?.value || parsedHeader.fsmLength || "");
  let tol = 1;
  if (isFinite(x) && isFinite(fsmLen)){
    if (x<=200 || x >= (fsmLen - 200)) tol = 1.5; // as per your earlier spec ±1.5 near ends
  }

  // Dia tolerance (upper only)
  let diaOk = true;
  if (isFinite(specDia) && isFinite(actDia)){
    if (specDia <= 10.7) diaOk = (actDia <= specDia + 0.4);
    else if (specDia >= 11.7) diaOk = (actDia <= specDia + 0.5);
    else diaOk = (actDia <= specDia + 0.5);
  } else diaOk = false;

  // Offset check: |ActYZ - Offset| <= tol
  let offsetOk = true;
  if (isFinite(offset) && isFinite(actYZ)) {
    offsetOk = Math.abs(actYZ - offset) <= tol;
  } else offsetOk = false;

  const allOk = (diaOk && offsetOk);

  // Color
  if (allOk){
    resultNode.textContent = "OK"; resultNode.className = "ok-cell";
    [tds[8],tds[9],tds[10]].forEach(td=>td.classList.remove("nok-cell"));
  } else {
    resultNode.textContent = "NOK"; resultNode.className = "nok-cell";
    if (!diaOk) tds[8].classList.add("nok-cell");
    if (!offsetOk) tds[9].classList.add("nok-cell");
  }
}

/* ===== First/Last Holes ===== */
function buildFirstHolesSection(){
  const blk=document.createElement("div"); blk.className="form-block";
  blk.innerHTML = `<strong>First holes of FSM - from front end</strong>`;
  const table = document.createElement("table"); table.className="table";
  table.innerHTML = `<thead><tr><th>Spec</th><th>Actual (Value from Hole edge)</th><th>Actual (Dia)</th><th>Offset</th><th>Result</th></tr></thead>`;
  const tb = document.createElement("tbody");

  for (let i=0;i<5;i++){
    const tr = document.createElement("tr");
    const spec = firstHoleSpecs[i] ?? "";
    const specTd = document.createElement("td"); specTd.textContent = spec;

    const valEdge = makeInput({className:"input-small", placeholder:"Edge"});
    const actDia  = makeInput({className:"input-small", placeholder:"Dia"});
    const offsetTd = document.createElement("td");
    const resultTd = document.createElement("td");

    const recalc = ()=>{
      const s = parseFloat(String(spec));
      const v = parseFloat(valEdge.value);
      const d = parseFloat(actDia.value);
      offsetTd.textContent = (isFinite(s)&&isFinite(d)) ? round(s + d/2, 2) : "";
      if (isFinite(s)&&isFinite(v)){
        const ok = Math.abs(v - s) <= 1; // ±1 mm
        resultTd.textContent = ok?"OK":"NOK";
        resultTd.className = ok?"ok-cell":"nok-cell";
      } else { resultTd.textContent = ""; resultTd.className=""; }
    };
    valEdge.addEventListener("input", recalc);
    actDia.addEventListener("input", recalc);

    tr.appendChild(specTd);
    tr.appendChild(tdWrap(valEdge));
    tr.appendChild(tdWrap(actDia));
    tr.appendChild(offsetTd);
    tr.appendChild(resultTd);
    tb.appendChild(tr);
  }
  table.appendChild(tb); blk.appendChild(table); formArea.appendChild(blk);
}

function buildLastHolesSection(){
  const blk=document.createElement("div"); blk.className="form-block";
  blk.innerHTML = `<strong>Last holes of FSM - from rear end</strong>`;
  const table = document.createElement("table"); table.className="table";
  table.innerHTML = `<thead><tr><th>Spec</th><th>Actual (Value from Hole edge)</th><th>Actual (Dia)</th><th>Offset</th><th>Result</th></tr></thead>`;
  const tb = document.createElement("tbody");

  for (let i=0;i<5;i++){
    const tr = document.createElement("tr");
    const spec = lastHoleSpecs[i] ?? "";
    const specTd = document.createElement("td"); specTd.textContent = spec;

    const valEdge = makeInput({className:"input-small", placeholder:"Edge"});
    const actDia  = makeInput({className:"input-small", placeholder:"Dia"});
    const offsetTd = document.createElement("td");
    const resultTd = document.createElement("td");

    const recalc = ()=>{
      const s = parseFloat(String(spec));
      const v = parseFloat(valEdge.value);
      const d = parseFloat(actDia.value);
      offsetTd.textContent = (isFinite(s)&&isFinite(d)) ? round(s + d/2, 2) : "";
      if (isFinite(s)&&isFinite(v)){
        const ok = Math.abs(v - s) <= 1; // ±1 mm
        resultTd.textContent = ok?"OK":"NOK";
        resultTd.className = ok?"ok-cell":"nok-cell";
      } else { resultTd.textContent = ""; resultTd.className=""; }
    };
    valEdge.addEventListener("input", recalc);
    actDia.addEventListener("input", recalc);

    tr.appendChild(specTd);
    tr.appendChild(tdWrap(valEdge));
    tr.appendChild(tdWrap(actDia));
    tr.appendChild(offsetTd);
    tr.appendChild(resultTd);
    tb.appendChild(tr);
  }
  table.appendChild(tb); blk.appendChild(table); formArea.appendChild(blk);
}

/* ===== Root / Flange / Web ===== */
function buildRootAndFlangeSection(){
  const blk=document.createElement("div"); blk.className="form-block";
  blk.innerHTML = `<strong>Root / Flange / Web (reference)</strong>`;
  const tbl=document.createElement("table"); tbl.className="table";
  tbl.innerHTML = `
    <thead><tr>
      <th>Root Width of FSM (Spec)</th><th>Root Width Act</th><th>Top Flange</th><th>Web PWM</th><th>Bottom Flange</th>
    </tr></thead>
  `;
  const tb=document.createElement("tbody");
  const tr=document.createElement("tr");
  const specTd=document.createElement("td"); specTd.textContent = document.getElementById("rootSpec")?.value || parsedHeader.rootWidth || "";
  const actIn=makeInput({className:"input-small", id:"rootAct_flange", placeholder:"Act (mm)"});
  actIn.addEventListener("input",()=>{
    const s=parseFloat(specTd.textContent||"");
    const a=parseFloat(actIn.value||"");
    if (isFinite(s)&&isFinite(a)){
      if (Math.abs(a - s) > 1){ actIn.classList.add("nok-cell"); }
      else { actIn.classList.remove("nok-cell"); }
    } else { actIn.classList.remove("nok-cell"); }
  });
  const top=document.createElement("td"); top.textContent="Top Flange: PWM (info)";
  const web=document.createElement("td"); web.textContent="Web PWM (info)";
  const bot=document.createElement("td"); bot.textContent="Bottom Flange (info)";
  tr.append(specTd, tdWrap(actIn), top, web, bot);
  tb.appendChild(tr); tbl.appendChild(tb); blk.appendChild(tbl); formArea.appendChild(blk);
}

/* ===== Part no location ===== */
function buildPartNoLocation(){
  const blk=document.createElement("div"); blk.className="form-block";
  blk.innerHTML = `<strong>Part no location (Spec / Act)</strong>`;
  const tbl=document.createElement("table"); tbl.className="table";
  tbl.innerHTML = `<thead><tr><th>Spec (mm)</th><th>Act (mm)</th><th>Result</th></tr></thead>`;
  const tb=document.createElement("tbody");
  const tr=document.createElement("tr");
  const specTd=tdWrap(makeInput({className:"input-small"}));
  const actIn = makeInput({className:"input-small"}); const actTd=tdWrap(actIn);
  const resTd=document.createElement("td");
  actIn.addEventListener("input",()=>{
    const s=parseFloat(specTd.querySelector("input").value||"");
    const a=parseFloat(actIn.value||"");
    if (isFinite(s)&&isFinite(a)){
      const ok = Math.abs(a - s) <= 5;
      resTd.textContent = ok?"OK":"NOK";
      resTd.className = ok?"ok-cell":"nok-cell";
    } else { resTd.textContent=""; resTd.className=""; }
  });
  tr.append(specTd, actTd, resTd); tb.appendChild(tr); tbl.appendChild(tb); blk.appendChild(tbl); formArea.appendChild(blk);
}

/* ===== Binary checks ===== */
function buildBinaryChecks(){
  const blk=document.createElement("div"); blk.className="form-block";
  blk.innerHTML = `<strong>Visual / Binary Checks</strong>`;
  const grid=document.createElement("div");
  grid.style.display='grid'; grid.style.gridTemplateColumns='repeat(4,1fr)'; grid.style.gap='8px';
  const labels=["Punch Break","Radius crack","Length Variation","Holes Burr","Slug mark","Line mark","Part No. Legibility","Pit Mark","Machine Error (YES/NO)"];
  labels.forEach(lbl=>{
    const cell=document.createElement("div"); cell.className="small";
    const title=document.createElement("div"); title.textContent=lbl;
    const okBtn=document.createElement("button"); okBtn.textContent="OK/YES"; okBtn.className="input-small";
    const nokBtn=document.createElement("button"); nokBtn.textContent="NOK/NO"; nokBtn.className="input-small";
    okBtn.addEventListener("click",()=>{ okBtn.classList.add("ok-cell-binary"); nokBtn.classList.remove("nok-cell-binary"); checkMandatoryBeforeExport(); });
    nokBtn.addEventListener("click",()=>{ nokBtn.classList.add("nok-cell-binary"); okBtn.classList.remove("ok-cell-binary"); checkMandatoryBeforeExport(); });
    cell.append(title, okBtn, nokBtn); grid.appendChild(cell);
  });
  blk.appendChild(grid); formArea.appendChild(blk);
}

/* ===== Remarks & signature ===== */
function buildRemarksAndSign(){
  const blk=document.createElement("div"); blk.className="form-block";
  blk.innerHTML = `<strong>Remarks / Details of issue</strong>`;
  const ta=document.createElement("textarea"); ta.id="remarks"; ta.placeholder="Enter remarks (optional)";
  blk.appendChild(ta);

  const row=document.createElement("div"); row.className="form-row";
  const col=document.createElement("div"); col.className="col";
  col.innerHTML = `<div class="small">Prodn Incharge (Shift Executive) - mandatory</div>`;
  const prod=makeInput({id:"prodIncharge", className:"input-small", placeholder:"Enter shift executive name"});
  prod.addEventListener("input", checkMandatoryBeforeExport);
  col.appendChild(prod); row.appendChild(col);
  blk.appendChild(row);
  formArea.appendChild(blk);
}

/* ===== Mandatory checks ===== */
function attachValidationListeners(){
  ["fsmSerial","holesAct","matrixUsed","prodIncharge"].forEach(id=>{
    document.getElementById(id)?.addEventListener("input", checkMandatoryBeforeExport);
  });
  document.getElementById("kbAct")?.addEventListener("input", checkMandatoryBeforeExport);
  document.getElementById("pcAct")?.addEventListener("input", checkMandatoryBeforeExport);
}
function checkMandatoryBeforeExport(){
  const fsmSerial = document.getElementById("fsmSerial")?.value.trim();
  const holesAct  = document.getElementById("holesAct")?.value.trim();
  const matrix    = document.getElementById("matrixUsed")?.value.trim();
  const prod      = document.getElementById("prodIncharge")?.value.trim();
  const kbAct     = document.getElementById("kbAct")?.value.trim();
  const pcAct     = document.getElementById("pcAct")?.value.trim();
  const missing=[];
  if (!fsmSerial) missing.push("FSM Serial Number");
  if (!holesAct) missing.push("Total Holes Count (Act)");
  if (!matrix) missing.push("Matrix used");
  if (!prod) missing.push("Prodn Incharge");
  if (!kbAct || !pcAct) missing.push("KB/PC Act");
  if (missing.length){
    exportPdfBtn.disabled = true;
    showMessage("Please fill mandatory fields: " + missing.join(", "), "warn");
    return false;
  }
  exportPdfBtn.disabled = false;
  clearMessage();
  return true;
}

/* ===== Export PDF ===== */
exportPdfBtn.addEventListener("click", async ()=>{
  if (!checkMandatoryBeforeExport()) return;
  const hasNok = document.querySelectorAll(".nok-cell, .nok-cell-binary").length > 0;
  if (hasNok && !confirm("Some fields are NOK. Proceed to export?")) return;

  exportPdfBtn.disabled = true; exportPdfBtn.textContent = "Generating PDF...";
  await generatePdfFromForm();
  exportPdfBtn.disabled = false; exportPdfBtn.textContent = "Export to PDF";
});

async function generatePdfFromForm(){
  const clone = document.createElement("div");
  clone.style.width="900px"; clone.style.padding="12px"; clone.style.position="fixed"; clone.style.left="-2000px";
  const headerClone = document.createElement("div"); headerClone.innerHTML = document.querySelector("header").innerHTML;
  clone.appendChild(headerClone);

  const fa = formArea.cloneNode(true);
  fa.querySelectorAll("input, textarea, button").forEach(n=>{
    if (n.tagName==="INPUT"){
      const s=document.createElement("span"); s.textContent = n.value; n.parentNode?.replaceChild(s,n);
    } else if (n.tagName==="TEXTAREA"){
      const d=document.createElement("div"); d.textContent = n.value; d.style.whiteSpace="pre-wrap"; n.parentNode?.replaceChild(d,n);
    } else if (n.tagName==="BUTTON"){ n.remove(); }
  });
  clone.appendChild(fa);
  document.body.appendChild(clone);

  try{
    const canvas = await html2canvas(clone, {scale:2, useCORS:true});
    const img = canvas.toDataURL("image/png");
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF("p","mm","a4");
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();
    const imgProps = pdf.getImageProperties(img);
    const pdfW = pageW - 20;
    const pdfH = (imgProps.height * pdfW) / imgProps.width;
    pdf.addImage(img, "PNG", 10, 10, pdfW, pdfH);
    pdf.setFontSize(9);
    const now = new Date();
    pdf.text(`Generated: ${now.toLocaleString()}`, 14, pageH - 10);

    const part = parsedHeader.partNumber || "PART";
    const rev  = parsedHeader.revision || "X";
    const hand = parsedHeader.hand || "H";
    const shortDate = now.toISOString().slice(0,10);
    const filename = `${part}_${rev}_${hand}_Inspection_${shortDate}.pdf`;
    pdf.save(filename);
    showMessage("PDF exported: " + filename, "info");
  } catch(e){
    console.error(e);
    alert("PDF export failed: " + e.message);
  } finally {
    document.body.removeChild(clone);
  }
}

/* ===== Utils ===== */
function showMessage(msg, type="info"){
  warnings.style.padding="8px"; warnings.style.borderRadius="6px"; warnings.style.marginTop="6px";
  warnings.textContent = msg;
  if (type==="warn"){ warnings.style.background="#fff6cc"; warnings.style.color="#ff8c00"; }
  else { warnings.style.background="#e8f4ff"; warnings.style.color="#005fcc"; }
}
function clearMessage(){ warnings.innerHTML=""; warnings.style.background=""; warnings.style.color=""; }
function round(n,dec){ return Math.round(n*Math.pow(10,dec))/Math.pow(10,dec); }
