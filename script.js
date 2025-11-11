/* ================= Soenen Phase 2 - PDF -> Editable Form (with Slot Logic) ================ */
/* Drop-in replacement for script.js */

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.9.179/pdf.worker.min.js";

const fileInput = document.getElementById("fileInput");
const extractBtn = document.getElementById("extractBtn");
const pdfViewer = document.getElementById("pdfViewer");
const headerOut = document.getElementById("headerOut");
const generateFormBtn = document.getElementById("generateFormBtn");
const formArea = document.getElementById("formArea");
const warnings = document.getElementById("warnings");
const exportPdfBtn = document.getElementById("exportPdfBtn");
const hdrPart = document.getElementById("hdrPart");
const hdrFsmSpec = document.getElementById("hdrFsmSpec");

let pdfFile = null;
let lastExtractedText = "";
let parsedHeader = {};      // partNumber, revision, hand, fsmLength, rootWidth, kbSpec, pcSpec
let parsedTableRows = [];   // tokens per row for main table

/* ---------------- Utilities ---------------- */
const round = (n, d) => Math.round(n * 10**d) / 10**d;
function showMessage(msg, type="info") {
  warnings.style.padding = "8px";
  warnings.style.borderRadius = "6px";
  warnings.style.marginTop = "6px";
  warnings.textContent = msg;
  if (type==="warn") { warnings.style.background = "#fff6cc"; warnings.style.color="#ff8c00"; }
  else { warnings.style.background = "#e8f4ff"; warnings.style.color="#005fcc"; }
}
function clearMessage(){ warnings.innerHTML=""; warnings.style.background=""; warnings.style.color=""; }

/* ---------------- Header display ---------------- */
function updateHeaderFields(h) {
  if (!h) return;
  const pn = h.partNumber || "—";
  const rev = h.revision || "—";
  const hand = h.hand || "—";
  const fsm = h.fsmLength ?? "—";
  if (hdrPart) hdrPart.textContent = `PART NUMBER / LEVEL / HAND : ${pn} / ${rev} / ${hand}`;
  if (hdrFsmSpec) hdrFsmSpec.textContent = `FSM LENGTH : ${fsm} mm`;
}

/* ---------------- PDF load & preview ---------------- */
fileInput.addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  pdfFile = f;
  const buffer = await pdfFile.arrayBuffer();
  await renderPDF(new Uint8Array(buffer));
  lastExtractedText = "";
  headerOut.textContent = "PDF loaded — click 'Load & Parse PDF'.";
  generateFormBtn.disabled = true;
  exportPdfBtn.disabled = true;
});

extractBtn.addEventListener("click", async () => {
  if (!pdfFile) { alert("Please select a PDF first."); return; }
  extractBtn.disabled = true;
  extractBtn.textContent = "Parsing...";
  try {
    const buf = await pdfFile.arrayBuffer();
    const text = await extractTextFromPDF(new Uint8Array(buf));
    lastExtractedText = text;
    parsedHeader = parseHeader(text);
    headerOut.textContent = JSON.stringify(parsedHeader, null, 2);
    updateHeaderFields(parsedHeader);
    parsedTableRows = detectTableLines(text); // best-effort
    generateFormBtn.disabled = false;
    showMessage("Parsed PDF. Click 'Create Editable Form' to build the inspection form.", "info");
  } catch (err) {
    console.error(err);
    alert("Error while parsing PDF: " + err.message);
  } finally {
    extractBtn.disabled = false;
    extractBtn.textContent = "Load & Parse PDF";
  }
});

async function renderPDF(bytes) {
  pdfViewer.innerHTML = "";
  try {
    const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
    const page = await pdf.getPage(1);
    const vp = page.getViewport({ scale: 1.1 });
    const canvas = document.createElement("canvas");
    canvas.width = vp.width; canvas.height = vp.height;
    const ctx = canvas.getContext("2d");
    await page.render({ canvasContext: ctx, viewport: vp }).promise;
    pdfViewer.appendChild(canvas);
  } catch {
    pdfViewer.textContent = "Unable to render preview.";
  }
}

async function extractTextFromPDF(bytes) {
  const doc = await pdfjsLib.getDocument({ data: bytes }).promise;
  let out = "";
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    out += content.items.map(it => it.str).join(" ") + "\n---PAGE_BREAK---\n";
  }
  return out;
}

/* ---------------- Parsing helpers ---------------- */
function parseHeader(t) {
  const h = { partNumber:null, revision:null, hand:null, date:null, formatNo:null, rootWidth:null, fsmLength:null, kbSpec:null, pcSpec:null };
  let m;

  m = t.match(/PART\s*NUMBER\s*\/\s*LEVEL\s*\/\s*HAND\s*[:\s\-]*([A-Z0-9_\-]+)\s*\/\s*([A-Z0-9_\-]+)\s*\/\s*([A-Z]+)/i);
  if (m) { h.partNumber = m[1]; h.revision = m[2]; h.hand = m[3]; }
  if (!h.partNumber) {
    m = t.match(/PART\s*NUMBER\s*[:\s\-]*([A-Z0-9_\-]+)/i);
    if (m) h.partNumber = m[1];
  }
  m = t.match(/Date\s*[:\-\s]*([0-3]?\d[-\/][A-Za-z]{3,}[-\/]\d{4})/i);
  if (m) h.date = m[1];

  m = t.match(/ROOT\s*WIDTH\s*OF\s*FSM\s*[:\-\s]*Spec[-\s]*([0-9.]+)/i);
  if (m) h.rootWidth = parseFloat(m[1]);
  m = t.match(/FSM\s*LENGTH\s*[:\-\s]*Spec[-\s]*([0-9.]+)/i);
  if (m) h.fsmLength = parseFloat(m[1]);

  m = t.match(/KB\s*&\s*PC\s*Code\s*[:\-]*\s*Spec[-\s]*([0-9]+)\s*\/\s*([0-9]+)/i);
  if (m) { h.kbSpec = m[1]; h.pcSpec = m[2]; }

  m = t.match(/FORMAT\s*NO\.?\s*[:\-]*\s*([A-Z0-9\/\s\-\_]+)/i);
  if (m) h.formatNo = m[1].trim();
  return h;
}

function detectTableLines(fullText) {
  if (!fullText) return [];
  const lines = fullText
    .split(/\n|---PAGE_BREAK---/).map(s => s.replace(/\s{2,}/g, " ").trim()).filter(Boolean);
  const rows = [];
  for (const line of lines) {
    // loose catch: row starting with number
    if (/^\d+(\.|-|)\s+/.test(line)) rows.push(line);
  }
  // fallback grab-bag if too few
  if (rows.length < 10) {
    const m = fullText.match(/\b\d+\s+[A-Z0-9\-]+\s+[A-Z0-9\-]+\s+\w+\s+\d+(?:\.\d+)?\b/g);
    if (m) rows.push(...m);
  }
  // tokenize
  return rows.map(r => r.split(/\s+/));
}

/* ---------------- Small DOM helpers ---------------- */
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
function makeCellText(text) {
  const el = document.createElement("div");
  el.textContent = text ?? "";
  return el;
}
function tdWrap(inner) {
  const td = document.createElement("td");
  if (inner instanceof HTMLElement) td.appendChild(inner);
  else td.textContent = inner ?? "";
  return td;
}

/* ---------------- Slot detection & parsing ---------------- */
function parseSlotSpec(specStr) {
  if (!specStr) return null;
  const s = String(specStr).toLowerCase().replace(/\s/g, "");
  // patterns: 9*12, 9x12, 9×12
  const m = s.match(/^(\d+(?:\.\d+)?)[\*x×](\d+(?:\.\d+)?)$/);
  if (!m) return null;
  const a = parseFloat(m[1]), b = parseFloat(m[2]);
  if (!isFinite(a) || !isFinite(b)) return null;
  // you said vertical height = 9, width = 12 (common), but be safe:
  const height = Math.min(a,b);
  const width  = Math.max(a,b);
  return { height, width };
}

/* ---------------- Build form ---------------- */
generateFormBtn.addEventListener("click", () => {
  formArea.innerHTML = "";
  warnings.innerHTML = "";

  buildHeaderInputs();
  buildMainTable();                 // 45 rows, slot-aware
  buildFirstHolesSection();
  buildLastHolesSection();
  buildRootAndFlangeSection();
  buildPartNoLocation();
  buildBinaryChecks();
  buildRemarksAndSign();
  attachValidationListeners();

  exportPdfBtn.disabled = false;
  showMessage("Editable form created. Fill actual values and use 'Export to PDF' to generate the report.", "info");
});

/* ---------- Header inputs (mandatory fields) ---------- */
function buildHeaderInputs() {
  const blk = document.createElement("div");
  blk.className = "form-block";
  blk.innerHTML = `<strong>Header / Mandatory fields</strong>`;

  const row1 = document.createElement("div"); row1.className="form-row";
  const c1 = document.createElement("div"); c1.className="col";
  c1.innerHTML = `<div class="small">FSM Serial Number:</div>`;
  c1.appendChild(makeInput({ id:"fsmSerial", className:"input-small", placeholder:"Enter FSM Serial No" }));

  const c2 = document.createElement("div"); c2.className="col";
  c2.innerHTML = `<div class="small">INSPECTORS:</div>`;
  const i1 = makeInput({ id:"inspector1", className:"input-small", placeholder:"Inspector 1" });
  const i2 = makeInput({ id:"inspector2", className:"input-small", placeholder:"Inspector 2" });
  c2.append(i1, document.createTextNode(" "), i2);

  row1.append(c1,c2); blk.appendChild(row1);

  const row2 = document.createElement("div"); row2.className="form-row";
  const c3 = document.createElement("div"); c3.className="col";
  c3.innerHTML = `<div class="small">TOTAL HOLES COUNT (Spec / Act):</div>`;
  const holesSpec = makeInput({ id:"holesSpec", className:"input-small", placeholder:"Spec (if known)" });
  const holesAct = makeInput({ id:"holesAct", className:"input-small", placeholder:"Act (mandatory)" });
  c3.append(holesSpec, document.createTextNode(" "), holesAct);

  const c4 = document.createElement("div"); c4.className="col";
  c4.innerHTML = `<div class="small">Matrix used:</div>`;
  c4.appendChild(makeInput({ id:"matrixUsed", className:"input-small", placeholder:"Matrix" }));
  row2.append(c3,c4); blk.appendChild(row2);

  // KB & PC line (Spec readonly)
  const row3 = document.createElement("div"); row3.className="form-row";
  const kbcol = document.createElement("div"); kbcol.className="col";
  kbcol.innerHTML = `<div class="small">KB & PC Code (Spec / Act)</div>`;
  const kbSpec = makeInput({ id:"kbSpec", className:"input-small", readOnly:true, size:3, value: parsedHeader.kbSpec ?? ""});
  const kbAct  = makeInput({ id:"kbAct",  className:"input-small", size:3, placeholder:"Act" });
  const pcSpec = makeInput({ id:"pcSpec", className:"input-small", readOnly:true, size:3, value: parsedHeader.pcSpec ?? ""});
  const pcAct  = makeInput({ id:"pcAct",  className:"input-small", size:3, placeholder:"Act" });
  kbcol.append(kbSpec, document.createTextNode(" / "), kbAct, document.createTextNode("    "),
               pcSpec, document.createTextNode(" / "), pcAct);
  row3.appendChild(kbcol); blk.appendChild(row3);

  // Root width & FSM length (Spec readonly, Act editable)
  const row4 = document.createElement("div"); row4.className="form-row";
  const rwcol = document.createElement("div"); rwcol.className="col";
  rwcol.innerHTML = `<div class="small">ROOT WIDTH OF FSM (Spec / Act) mm</div>`;
  rwcol.append(
    makeInput({ id:"rootSpec", className:"input-small", readOnly:true, value: parsedHeader.rootWidth ?? "", size:6 }),
    document.createTextNode(" "),
    makeInput({ id:"rootAct",  className:"input-small", placeholder:"Act (mm)", step:"0.01", size:6 })
  );
  const fsmcol = document.createElement("div"); fsmcol.className="col";
  fsmcol.innerHTML = `<div class="small">FSM LENGTH (Spec / Act) mm</div>`;
  fsmcol.append(
    makeInput({ id:"fsmSpec", className:"input-small", readOnly:true, value: parsedHeader.fsmLength ?? "", size:6 }),
    document.createTextNode(" "),
    makeInput({ id:"fsmAct",  className:"input-small", placeholder:"Act (mm)", size:6 })
  );
  row4.append(rwcol, fsmcol); blk.appendChild(row4);

  formArea.appendChild(blk);
}

/* ---------- Main table (45 rows, slot-aware) ---------- */
function buildMainTable() {
  const blk = document.createElement("div");
  blk.className = "form-block";
  blk.innerHTML = `<strong>Main Inspection Table</strong>`;

  const table = document.createElement("table");
  table.className = "table";
  const thead = document.createElement("thead");
  thead.innerHTML = `<tr>
    <th>Sl No</th><th>Press</th><th>Sel ID</th><th>Ref</th><th>X-axis</th><th>Spec (Y or Z)</th>
    <th>Spec Dia / Slot</th><th>Value from Hole edge (Act)</th>
    <th>Actual Dia / Slot Height</th><th>Slot Width (if slot)</th>
    <th>Actual Y or Z</th><th>Offset</th><th>Result</th>
  </tr>`;
  table.appendChild(thead);
  const tbody = document.createElement("tbody");

  const TOTAL_ROWS = 45;
  const rows = [];
  for (let i = 0; i < TOTAL_ROWS; i++) {
    rows.push(parsedTableRows[i] || []); // undefined rows → blanks
  }

  rows.forEach((tokens, idx) => {
    const tr = document.createElement("tr");

    // tokens mapping (best-effort)
    const press  = tokens[1] ?? "";   // readonly
    const selID  = tokens[2] ?? "";   // readonly
    const ref    = tokens[3] ?? "";   // readonly
    const xVal   = tokens[4] ?? "";   // editable? requirement says non-editable for col 1–7, so readonly here
    const specYZ = tokens[5] ?? "";   // readonly
    const specDiaStr = tokens[6] ?? "";// readonly

    const slNoCell  = makeCellText(idx+1);
    const pressCell = makeCellText(press);
    const selCell   = makeCellText(selID);
    const refCell   = makeCellText(ref);

    const xaxis = makeInput({ className:"input-small", value:xVal });  // set readonly to enforce non-editable
    xaxis.readOnly = true;

    const specYZInp = makeInput({ className:"input-small", value: specYZ });
    specYZInp.readOnly = true;

    const specDiaInp = makeInput({ className:"input-small", value: specDiaStr });
    specDiaInp.readOnly = true;

    const valEdge = makeInput({ className:"input-small", placeholder:"Edge" });

    // SLOT vs DIA inputs
    const slot = parseSlotSpec(specDiaStr);
    let actDiaOrHeight, slotWidthInp = null;
    if (slot) {
      // col 9 → Act Height (used for offset), col 10 → Act Width
      actDiaOrHeight = makeInput({ className:"input-small", placeholder:"Act Height" });
      slotWidthInp   = makeInput({ className:"input-small", placeholder:"Act Width" });
    } else {
      actDiaOrHeight = makeInput({ className:"input-small", placeholder:"Actual Dia" });
    }

    const actualYZ = makeInput({ className:"input-small", placeholder:"Actual Y/Z" });

    const offsetCell = makeCellText("");
    const resultCell = makeCellText("");

    // Recalc on any change of relevant inputs
    const recalc = () => recalcRowAndMark({
      tr,
      xaxis,
      specYZInp,
      specDiaInp,
      valEdge,
      actDiaOrHeight,
      slotWidthInp,
      actualYZ,
      offsetCell,
      resultCell
    });

    [valEdge, actDiaOrHeight, actualYZ, xaxis].forEach(el => el && el.addEventListener('input', recalc));
    if (slotWidthInp) slotWidthInp.addEventListener('input', recalc);

    tr.appendChild(tdWrap(slNoCell));
    tr.appendChild(tdWrap(pressCell));
    tr.appendChild(tdWrap(selCell));
    tr.appendChild(tdWrap(refCell));
    tr.appendChild(tdWrap(xaxis));
    tr.appendChild(tdWrap(specYZInp));
    tr.appendChild(tdWrap(specDiaInp));
    tr.appendChild(tdWrap(valEdge));
    tr.appendChild(tdWrap(actDiaOrHeight));
    tr.appendChild(tdWrap(slotWidthInp ? slotWidthInp : makeCellText("")));
    tr.appendChild(tdWrap(actualYZ));
    tr.appendChild(tdWrap(offsetCell));
    tr.appendChild(tdWrap(resultCell));

    // mark spec columns (1–7) as visually locked
    [xaxis, specYZInp, specDiaInp].forEach(inp => inp.readOnly = true);

    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  blk.appendChild(table);
  formArea.appendChild(blk);
}

/* ---------- Row calculation (slot-aware) ---------- */
function recalcRowAndMark(ctx) {
  const {
    xaxis, specYZInp, specDiaInp, valEdge,
    actDiaOrHeight, slotWidthInp, actualYZ, offsetCell, resultCell
  } = ctx;

  const x = parseFloat(xaxis.value);
  const specYZ = parseFloat(specYZInp.value);
  const specDiaStr = specDiaInp.value;
  const specSlot = parseSlotSpec(specDiaStr);

  const fsmSpecVal = parseFloat(document.getElementById('fsmSpec')?.value || parsedHeader.fsmLength || NaN);

  // Tolerance band by X & FSM
  let tol = 1;
  if (!isNaN(x) && !isNaN(fsmSpecVal)) {
    if (x <= 200 || x >= (fsmSpecVal - 200)) tol = 1.5; // per your rule: ±1.5 edge zones
  }

  // Compute offset
  let offset = null;
  let sizeOK = true;

  if (specSlot) {
    const actH = parseFloat(actDiaOrHeight.value);
    const actW = parseFloat(slotWidthInp?.value);
    // specSlot.height=9, width=12 typically
    // Slot tolerance: 0 to +0.5 on both
    if (!isNaN(actH) && actH > specSlot.height + 0.5) sizeOK = false;
    if (!isNaN(actW) && actW > specSlot.width + 0.5) sizeOK = false;
    if (!isNaN(specYZ) && !isNaN(actH)) offset = specYZ + (actH/2);
  } else {
    const specDia = parseFloat(specDiaStr);
    const actDia = parseFloat(actDiaOrHeight.value);
    if (!isNaN(specDia) && !isNaN(actDia)) {
      // Dia tolerance: <=10.7 → +0.4 ; >=11.7 → +0.5 ; else +0.5
      const up = (specDia <= 10.7) ? 0.4 : 0.5;
      if (!(actDia >= specDia - 0 && actDia <= specDia + up)) sizeOK = false;
      if (!isNaN(specYZ)) offset = specYZ + (actDia/2);
    } else {
      sizeOK = false;
    }
  }

  offsetCell.textContent = (offset == null || isNaN(offset)) ? "" : String(round(offset,2));

  // Offset check vs Actual Y/Z
  const actYZ = parseFloat(actualYZ.value);
  let offsetOK = true;
  if (!isNaN(offset) && !isNaN(actYZ)) {
    const dev = Math.abs(actYZ - offset);
    offsetOK = dev <= tol;
  } else {
    offsetOK = false;
  }

  const allOk = sizeOK && offsetOK;

  // Paint result
  resultCell.textContent = allOk ? "OK" : "NOK";
  resultCell.className = allOk ? "ok-cell" : "nok-cell";
}

/* ---------- First/Last holes quick blocks (kept from earlier) ---------- */
function buildFirstHolesSection() {
  const blk = document.createElement("div"); blk.className="form-block";
  blk.innerHTML = `<strong>First holes of FSM - from front end</strong>`;
  const table = document.createElement('table'); table.className='table';
  table.innerHTML = `<thead><tr><th>Spec</th><th>Actual (Value from Hole edge)</th><th>Actual (Dia)</th><th>Offset</th><th>Result</th></tr></thead>`;
  const tb = document.createElement('tbody');
  const specs = extractFirstHoleSpecs(lastExtractedText) || ["","","","",""];
  for (let i=0;i<5;i++){
    const tr = document.createElement('tr');
    const specTd = document.createElement('td'); specTd.textContent = specs[i] || "";
    const valEdge = makeInput({className:"input-small"}); valEdge.placeholder="Edge";
    const actDia = makeInput({className:"input-small"}); actDia.placeholder="Dia";
    const offsetTd = document.createElement('td');
    const resultTd = document.createElement('td');
    [valEdge, actDia].forEach(inp => inp.addEventListener('input', ()=>{
      const spec = parseFloat(specTd.textContent||NaN);
      const dia = parseFloat(actDia.value||NaN);
      let offset = (isNaN(spec) || isNaN(dia)) ? "" : round(spec + (dia/2), 2);
      offsetTd.textContent = offset;
      const val = parseFloat(valEdge.value||NaN);
      if (!isNaN(val) && !isNaN(spec)) {
        if (Math.abs(val - spec) > 1) { valEdge.classList.add('nok-cell'); resultTd.textContent = "NOK"; resultTd.className="nok-cell"; }
        else { valEdge.classList.remove('nok-cell'); resultTd.textContent = "OK"; resultTd.className="ok-cell"; }
      } else { resultTd.textContent = ""; resultTd.className = ""; }
    }) );
    tr.appendChild(specTd);
    tr.appendChild(tdWrap(valEdge));
    tr.appendChild(tdWrap(actDia));
    tr.appendChild(offsetTd);
    tr.appendChild(resultTd);
    tb.appendChild(tr);
  }
  table.appendChild(tb);
  blk.appendChild(table);
  formArea.appendChild(blk);
}
function extractFirstHoleSpecs(text) {
  if (!text) return null;
  const matches = text.match(/([0-9]{1,4}(?:\.[0-9]+)?)/g);
  if (!matches) return null;
  const smalls = matches.map(Number).filter(n=>n>0 && n<5000);
  return smalls.slice(0,5).map(n=>n.toString());
}
function buildLastHolesSection() {
  const blk = document.createElement("div"); blk.className="form-block";
  blk.innerHTML = `<strong>Last holes of FSM - from rear end</strong>`;
  const table = document.createElement('table'); table.className='table';
  table.innerHTML = `<thead><tr><th>Spec</th><th>Actual (Value from Hole edge)</th><th>Actual (Dia)</th><th>Offset</th><th>Result</th></tr></thead>`;
  const tb = document.createElement('tbody');
  const specs = extractFirstHoleSpecs(lastExtractedText) || ["","","","",""];
  for (let i=0;i<5;i++){
    const tr = document.createElement('tr');
    const specTd = document.createElement('td'); specTd.textContent = specs[i] || "";
    const valEdge = makeInput({className:"input-small"}); valEdge.placeholder="Edge";
    const actDia = makeInput({className:"input-small"}); actDia.placeholder="Dia";
    const offsetTd = document.createElement('td');
    const resultTd = document.createElement('td');
    [valEdge, actDia].forEach(inp => inp.addEventListener('input', ()=>{
      const spec = parseFloat(specTd.textContent||NaN);
      const dia = parseFloat(actDia.value||NaN);
      let offset = (isNaN(spec) || isNaN(dia)) ? "" : round(spec + (dia/2), 2);
      offsetTd.textContent = offset;
      const val = parseFloat(valEdge.value||NaN);
      if (!isNaN(val) && !isNaN(spec)) {
        if (Math.abs(val - spec) > 1) { valEdge.classList.add('nok-cell'); resultTd.textContent = "NOK"; resultTd.className="nok-cell"; }
        else { valEdge.classList.remove('nok-cell'); resultTd.textContent = "OK"; resultTd.className="ok-cell"; }
      } else { resultTd.textContent = ""; resultTd.className = ""; }
    }) );
    tr.appendChild(specTd);
    tr.appendChild(tdWrap(valEdge));
    tr.appendChild(tdWrap(actDia));
    tr.appendChild(offsetTd);
    tr.appendChild(resultTd);
    tb.appendChild(tr);
  }
  table.appendChild(tb);
  blk.appendChild(table);
  formArea.appendChild(blk);
}

/* ---------- Root/flange quick ref ---------- */
function buildRootAndFlangeSection() {
  const blk = document.createElement('div'); blk.className='form-block';
  blk.innerHTML = `<strong>Root / Flange / Web (reference)</strong>`;
  const tbl = document.createElement('table'); tbl.className='table';
  tbl.innerHTML = `<thead><tr><th>Root Width Spec</th><th>Root Width Act</th><th>Top Flange</th><th>Web PWM</th><th>Bottom Flange</th></tr></thead>`;
  const tb = document.createElement('tbody');
  const tr = document.createElement('tr');

  const rootSpecTd = document.createElement('td'); rootSpecTd.textContent = document.getElementById('rootSpec')?.value || "";
  const rootActInput = makeInput({className:'input-small', id:'rootAct_flange', placeholder:'Act (mm)'});
  rootActInput.addEventListener('input', ()=> {
    const spec = parseFloat(rootSpecTd.textContent||NaN);
    const act = parseFloat(rootActInput.value||NaN);
    if (!isNaN(spec) && !isNaN(act)) {
      if (Math.abs(act - spec) > 1) rootActInput.classList.add('nok-cell');
      else rootActInput.classList.remove('nok-cell');
    }
  });
  const topFl = document.createElement('td'); topFl.textContent = "Top Flange: PWM (info)";
  const webPW = document.createElement('td'); webPW.textContent = "Web PWM (info)";
  const bottomFl = document.createElement('td'); bottomFl.textContent = "Bottom Flange (info)";

  tr.append(rootSpecTd, tdWrap(rootActInput), topFl, webPW, bottomFl);
  tb.appendChild(tr); tbl.appendChild(tb); blk.appendChild(tbl); formArea.appendChild(blk);
}

/* ---------- Part no location ---------- */
function buildPartNoLocation() {
  const blk = document.createElement('div'); blk.className='form-block';
  blk.innerHTML = `<strong>Part no location (Spec / Act)</strong>`;
  const table = document.createElement('table'); table.className='table';
  table.innerHTML = `<thead><tr><th>Spec (mm)</th><th>Act (mm)</th><th>Result</th></tr></thead>`;
  const tb = document.createElement('tbody');
  const tr = document.createElement('tr');
  const specTd = document.createElement('td'); const specInp = makeInput({className:'input-small'}); specTd.appendChild(specInp);
  const actTd = document.createElement('td'); const actInp = makeInput({className:'input-small'}); actTd.appendChild(actInp);
  const resTd = document.createElement('td');
  actInp.addEventListener('input', ()=> {
    const s = parseFloat(specInp.value||NaN); const a = parseFloat(actInp.value||NaN);
    if (!isNaN(s) && !isNaN(a)) {
      if (Math.abs(a - s) > 5) { actInp.classList.add('nok-cell'); resTd.textContent="NOK"; resTd.className='nok-cell'; }
      else { actInp.classList.remove('nok-cell'); resTd.textContent="OK"; resTd.className='ok-cell'; }
    } else { resTd.textContent=""; resTd.className=''; }
  });
  tr.append(specTd, actTd, resTd); tb.appendChild(tr); table.appendChild(tb);
  blk.appendChild(table); formArea.appendChild(blk);
}

/* ---------- Binary checks ---------- */
function buildBinaryChecks() {
  const blk = document.createElement('div'); blk.className='form-block';
  blk.innerHTML = `<strong>Visual / Binary Checks</strong>`;
  const grid = document.createElement('div');
  grid.style.display='grid'; grid.style.gridTemplateColumns='repeat(4,1fr)'; grid.style.gap='8px';
  const labels = ["Punch Break","Radius crack","Length Variation","Holes Burr","Slug mark","Line mark","Part No. Legibility","Pit Mark","Machine Error (YES/NO)"];
  labels.forEach(lbl=>{
    const cell = document.createElement('div'); cell.className='small';
    const labelEl = document.createElement('div'); labelEl.textContent = lbl;
    const okBtn = document.createElement('button'); okBtn.textContent='OK/YES'; okBtn.className='input-small';
    const nokBtn = document.createElement('button'); nokBtn.textContent='NOK/NO';  nokBtn.className='input-small';
    okBtn.addEventListener('click', ()=> { okBtn.classList.add('ok-cell-binary'); nokBtn.classList.remove('nok-cell-binary'); });
    nokBtn.addEventListener('click', ()=> { nokBtn.classList.add('nok-cell-binary'); okBtn.classList.remove('ok-cell-binary'); });
    cell.append(labelEl, okBtn, nokBtn);
    grid.appendChild(cell);
  });
  blk.appendChild(grid); formArea.appendChild(blk);
}

/* ---------- Remarks & sign ---------- */
function buildRemarksAndSign() {
  const blk = document.createElement('div'); blk.className='form-block';
  blk.innerHTML = `<strong>Remarks / Details of issue</strong>`;
  const ta = document.createElement('textarea'); ta.id='remarks'; ta.placeholder = "Enter remarks (optional)";
  blk.appendChild(ta);

  const signRow = document.createElement('div'); signRow.className='form-row';
  const col1 = document.createElement('div'); col1.className='col';
  col1.innerHTML = `<div class="small">Prodn Incharge (Shift Executive) - mandatory</div>`;
  const prodIn = makeInput({className:"input-small", id:"prodIncharge", placeholder:"Enter shift executive name"});
  col1.appendChild(prodIn); signRow.appendChild(col1);
  blk.appendChild(signRow); formArea.appendChild(blk);
}

/* ---------- Validation listeners ---------- */
function attachValidationListeners() {
  ["fsmSerial","holesAct","matrixUsed","prodIncharge"].forEach(id=>{
    const el = document.getElementById(id); if (el) el.addEventListener('input', checkMandatoryBeforeExport);
  });
  document.getElementById('kbAct')?.addEventListener('input', checkMandatoryBeforeExport);
  document.getElementById('pcAct')?.addEventListener('input', checkMandatoryBeforeExport);
  document.getElementById('rootAct')?.addEventListener('input', checkMandatoryBeforeExport);
  document.getElementById('fsmAct')?.addEventListener('input', checkMandatoryBeforeExport);
}

function checkMandatoryBeforeExport() {
  const fsmSerial = document.getElementById('fsmSerial')?.value || "";
  const holesAct = document.getElementById('holesAct')?.value || "";
  const matrixUsed = document.getElementById('matrixUsed')?.value || "";
  const prod = document.getElementById('prodIncharge')?.value || "";
  let missing = [];
  if (!fsmSerial.trim()) missing.push("FSM Serial Number");
  if (!holesAct.trim()) missing.push("Total Holes Count (Act)");
  if (!matrixUsed.trim()) missing.push("Matrix used");
  if (!prod.trim()) missing.push("Prodn Incharge");
  if (missing.length) {
    showMessage("Please fill mandatory fields: " + missing.join(", "), "warn");
    exportPdfBtn.disabled = true; return false;
  }
  const kbAct = document.getElementById('kbAct')?.value || "";
  const pcAct = document.getElementById('pcAct')?.value || "";
  if (!kbAct.trim() || !pcAct.trim()) {
    showMessage("Please fill KB & PC Act values.", "warn");
    exportPdfBtn.disabled = true; return false;
  }
  exportPdfBtn.disabled = false; clearMessage(); return true;
}

/* ---------- Export PDF ---------- */
exportPdfBtn.addEventListener('click', async () => {
  if (!checkMandatoryBeforeExport()) return;
  const hasNok = document.querySelectorAll('.nok-cell, .nok-cell-binary').length > 0;
  if (hasNok && !confirm("Some fields are out of spec (NOK). Proceed to export anyway?")) return;

  exportPdfBtn.disabled = true; exportPdfBtn.textContent = "Generating PDF...";
  await generatePdfFromForm();
  exportPdfBtn.disabled = false; exportPdfBtn.textContent = "Export to PDF";
});

async function generatePdfFromForm() {
  const clone = document.createElement('div');
  clone.style.width = "900px"; clone.style.padding = "12px";
  const title = document.createElement('div'); title.innerHTML = document.querySelector('header').innerHTML;
  clone.appendChild(title);
  const fa = formArea.cloneNode(true);

  fa.querySelectorAll('input, textarea, button').forEach(n=>{
    if (n.tagName === 'INPUT') {
      const span = document.createElement('span'); span.textContent = n.value;
      n.parentNode && n.parentNode.replaceChild(span, n);
    } else if (n.tagName === 'TEXTAREA') {
      const span = document.createElement('div'); span.textContent = n.value; span.style.whiteSpace="pre-wrap";
      n.parentNode && n.parentNode.replaceChild(span, n);
    } else if (n.tagName === 'BUTTON') { n.remove(); }
  });

  clone.appendChild(fa);
  clone.style.position = "fixed"; clone.style.left = "-2000px";
  document.body.appendChild(clone);

  try {
    const canvas = await html2canvas(clone, {scale: 2, useCORS: true});
    const imgData = canvas.toDataURL('image/png');
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF('p','mm','a4');
    const pageWidth = pdf.internal.pageSize.getWidth();
    const imgProps = pdf.getImageProperties(imgData);
    const pdfWidth = pageWidth - 20;
    const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
    pdf.addImage(imgData, 'PNG', 10, 10, pdfWidth, pdfHeight);
    const stamp = new Date().toLocaleString();
    pdf.setFontSize(9);
    pdf.text(`Generated: ${stamp}`, 14, pdf.internal.pageSize.getHeight() - 10);

    const part = parsedHeader.partNumber || "PART";
    const rev  = parsedHeader.revision || "X";
    const hand = parsedHeader.hand || "H";
    const shortDate = new Date().toISOString().slice(0,10);
    pdf.save(`${part}_${rev}_${hand}_Inspection_${shortDate}.pdf`);
    showMessage("PDF exported.", "info");
  } catch (err) {
    console.error(err);
    alert("PDF export failed: " + err.message);
  } finally {
    document.body.removeChild(clone);
  }
}
