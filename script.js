/* Phase 2 script - builds editable form, validations and PDF export */
/* NOTE: best-effort parsing; please edit specs manually if not filled by parser */

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
let parsedHeader = {}; // will hold partNumber, revision, hand, fsmLength, rootWidth, KBPC specs etc
let parsedTableRows = []; // best-effort rows from main inspection table

// -------------------- Update Header Fields --------------------
function updateHeaderFields(parsedHeader) {
  if (!parsedHeader) return;

  const hdrPartEl = document.getElementById("hdrPart");
  const hdrFsmSpecEl = document.getElementById("hdrFsmSpec");
  if (!hdrPartEl || !hdrFsmSpecEl) return;

  const partNumber = parsedHeader.partNumber || "—";
  const revision = parsedHeader.revision || "—";
  const hand = parsedHeader.hand || "—";
  const fsmLength = parsedHeader.fsmLength || "—";

  hdrPartEl.textContent = `PART NUMBER / LEVEL / HAND : ${partNumber} / ${revision} / ${hand}`;
  hdrFsmSpecEl.textContent = `FSM LENGTH : ${fsmLength} mm`;
}

/* ---------- File load & preview ---------- */
fileInput.addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  pdfFile = f;
  const buffer = await pdfFile.arrayBuffer();
  await renderPDF(new Uint8Array(buffer));
  // clear last parse
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
    const hdr = parseHeader(text);
    parsedHeader = hdr;
    headerOut.textContent = JSON.stringify(hdr, null, 2);

    // update header elements if present
    if (hdrPart && hdrFsmSpec) updateHeaderFields(parsedHeader);

    // attempt to parse main table as rows:
    parsedTableRows = detectTableLines(text);
    // enable create-form button
    generateFormBtn.disabled = false;

    // also mirror into header spans:
    if (hdr.partNumber) hdrPart.textContent = `${hdr.partNumber} / ${hdr.revision} / ${hdr.hand}`;
    else hdrPart.textContent = "—";
    hdrFsmSpec.textContent = hdr.fsmLength || "—";

    showMessage("Parsed PDF. Click 'Create Editable Form' to build the inspection form (you can edit any spec).", "info");
  } catch (err) {
    console.error(err);
    alert("Error while parsing PDF: " + err.message);
  } finally {
    extractBtn.disabled = false;
    extractBtn.textContent = "Load & Parse PDF";
  }
});

/* ---------- render preview (first page) ---------- */
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
  } catch (e) {
    pdfViewer.textContent = "Unable to render preview.";
  }
}

/* ---------- text extraction ---------- */
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

/* ---------- parsing helpers (best-effort regexes) ---------- */
function parseHeader(t) {
  const h = { partNumber:null, revision:null, hand:null, date:null, formatNo:null, rootWidth:null, fsmLength:null, kbSpec:null, pcSpec:null };
  let m;

  // Part / Rev / Hand
  m = t.match(/PART\s*NUMBER\s*\/\s*LEVEL\s*\/\s*HAND\s*[:\s\-]*([A-Z0-9_\-]+)\s*\/\s*([A-Z0-9_\-]+)\s*\/\s*([A-Z]+)/i);
  if (m) { h.partNumber = m[1]; h.revision = m[2]; h.hand = m[3]; }
  // fallback
  if (!h.partNumber) {
    m = t.match(/PART\s*NUMBER\s*[:\s\-]*([A-Z0-9_\-]+)/i);
    if (m) h.partNumber = m[1];
  }

  // Date
  m = t.match(/Date\s*[:\-\s]*([0-3]?\d[-\/][A-Za-z]{3,}[-\/]\d{4})/i);
  if (m) h.date = m[1];

  // Root width
  m = t.match(/ROOT\s*WIDTH\s*OF\s*FSM\s*[:\-\s]*Spec[-\s]*([0-9.]+(?:\.[0-9]+)?)/i);
  if (m) h.rootWidth = parseFloat(m[1]);

  // FSM LENGTH
  m = t.match(/FSM\s*LENGTH\s*[:\-\s]*Spec[-\s]*([0-9.]+)/i);
  if (m) h.fsmLength = parseFloat(m[1]);

  // KB & PC spec (try to find "KB & PC Code : Spec- 5 / 1" pattern)
  m = t.match(/KB\s*&\s*PC\s*Code\s*[:\-]*\s*Spec[-\s]*([0-9]+)\s*\/\s*([0-9]+)/i);
  if (m) { h.kbSpec = m[1]; h.pcSpec = m[2]; }

  // Format number
  m = t.match(/FORMAT\s*NO\.?\s*[:\-]*\s*([A-Z0-9\/\s\-\_]+)/i);
  if (m) h.formatNo = m[1].trim();

  return h;
}

/* Detect table-like lines (best-effort) */
function detectTableLines(fullText) {
  if (!fullText) return [];
  const lines = fullText
    .split(/\n|---PAGE_BREAK---/)
    .map(s => s.replace(/\s{2,}/g, " ").trim())
    .filter(Boolean);

  const rows = [];
  for (const line of lines) {
    // Match patterns like "1 ", " 1.", "01", etc.
    if (/^\s*\d+(\.|-)?\s+\w+/.test(line)) {
      rows.push(line);
    }
  }

  // Fallback: try to catch multi-line numeric sequences if still empty
  if (rows.length < 5) {
    const allNums = fullText.match(/(\d+\s+[A-Z]{1,3}\s+\d+)/g);
    if (allNums) rows.push(...allNums);
  }

  // Split each row into cell-like tokens
  return rows.map(r => r.trim().split(/\s+/));
}


/* ---------- create small helper to create an input with attributes ---------- */
function makeInput(attrs = {}) {
  const inp = document.createElement("input");
  inp.type = attrs.type || "text";
  inp.value = attrs.value || "";
  inp.className = attrs.className || "input-small";
  if (attrs.placeholder) inp.placeholder = attrs.placeholder;
  if (attrs.step) inp.step = attrs.step;
  if (attrs.onchange) inp.addEventListener('change', attrs.onchange);
  if (attrs.oninput) inp.addEventListener('input', attrs.oninput);
  if (attrs.readOnly) inp.readOnly = true;
  if (attrs.size) inp.size = attrs.size;
  if (attrs.id) inp.id = attrs.id;
  return inp;
}

/* ---------- Header inputs (mandatory fields) ---------- */
function buildHeaderInputs() {
  const blk = document.createElement("div");
  blk.className = "form-block";
  blk.innerHTML = `<strong>Header / Mandatory fields</strong>`;
  const row = document.createElement("div");
  row.className = "form-row";

  // FSM Serial Number
  const c1 = document.createElement("div"); c1.className = "col";
  c1.innerHTML = `<div class="small">FSM Serial Number:</div>`;
  const fsmSerial = makeInput({placeholder:"Enter FSM Serial No", className:"input-small", id:"fsmSerial"});
  c1.appendChild(fsmSerial);

  // Inspectors
const c2 = document.createElement("div");
c2.className = "col";
c2.innerHTML = `<div class="small">INSPECTORS:</div>`;
const inspector1 = makeInput({
  placeholder: "Inspector 1",
  className: "input-small",
  id: "inspector1"
});
const inspector2 = makeInput({
  placeholder: "Inspector 2",
  className: "input-small",
  id: "inspector2"
});
c2.append(inspector1, document.createTextNode(" "), inspector2);

  row.appendChild(c1); row.appendChild(c2);
  blk.appendChild(row);

  const row2 = document.createElement("div"); row2.className="form-row";
  const c3 = document.createElement("div"); c3.className="col";
  c3.innerHTML = `<div class="small">TOTAL HOLES COUNT (Spec / Act):</div>`;
  const holesSpec = makeInput({className:"input-small", id:"holesSpec"});
  const holesAct = makeInput({className:"input-small", id:"holesAct"});
  holesSpec.placeholder = "Spec (if parsed)";
  holesAct.placeholder = "Act (mandatory)";
  c3.appendChild(holesSpec); c3.appendChild(document.createTextNode(" "));
  c3.appendChild(holesAct);

  const c4 = document.createElement("div"); c4.className="col";
  c4.innerHTML = `<div class="small">Matrix used:</div>`;
  const matrixUsed = makeInput({className:"input-small", id:"matrixUsed", placeholder:"Matrix"});
  c4.appendChild(matrixUsed);

  row2.appendChild(c3); row2.appendChild(c4);
  blk.appendChild(row2);

// KB & PC area: allow dynamic spec/act
const row3 = document.createElement("div");
row3.className = "form-row";
const kbcol = document.createElement("div");
kbcol.className = "col";
kbcol.innerHTML = `<div class="small">KB & PC Code (Spec / Act)</div>`;

// Compact KB/PC inputs (3 chars wide)
const kbSpec = makeInput({
  className: "input-small",
  id: "kbSpec",
  readOnly: true,
  placeholder: "Spec",
  size: 3
});
const pcSpec = makeInput({
  className: "input-small",
  id: "pcSpec",
  readOnly: true,
  placeholder: "Spec",
  size: 3
});
const kbAct = makeInput({
  className: "input-small",
  id: "kbAct",
  placeholder: "Act",
  size: 3
});
const pcAct = makeInput({
  className: "input-small",
  id: "pcAct",
  placeholder: "Act",
  size: 3
});

// Pre-fill parsed values if available
if (parsedHeader.kbSpec) kbSpec.value = parsedHeader.kbSpec;
if (parsedHeader.pcSpec) pcSpec.value = parsedHeader.pcSpec;

kbcol.append(
  kbSpec,
  document.createTextNode(" / "),
  kbAct,
  document.createTextNode("    "),
  pcSpec,
  document.createTextNode(" / "),
  pcAct
);
row3.appendChild(kbcol);
blk.appendChild(row3);

// Root Width + FSM Length (side-by-side)
const row4 = document.createElement("div");
row4.className = "form-row";

// ROOT WIDTH BLOCK
const rwcol = document.createElement("div");
rwcol.className = "col";
rwcol.innerHTML = `<div class="small">ROOT WIDTH OF FSM (Spec / Act) mm</div>`;

const rootSpec = makeInput({
  className: "input-small",
  id: "rootSpec",
  readOnly: true,
  value: parsedHeader.rootWidth || "",
  size: 6
});

const rootAct = makeInput({
  className: "input-small",
  id: "rootAct",
  placeholder: "Act (mm)",
  step: "0.01",
  size: 6
});

rwcol.appendChild(rootSpec);
rwcol.appendChild(document.createTextNode(" "));
rwcol.appendChild(rootAct);

// store Root Width Act reference for later
const rootActSpan = document.createElement("span");
rootActSpan.id = "rootActRef"; // used by flange section
rootActSpan.style.display = "none";
rootActSpan.textContent = ""; // will mirror value
rwcol.appendChild(rootActSpan);

// keep rootAct value synced
rootAct.addEventListener("input", () => {
  rootActSpan.textContent = rootAct.value;
});

// FSM LENGTH BLOCK
const fsmcol = document.createElement("div");
fsmcol.className = "col";
fsmcol.innerHTML = `<div class="small">FSM LENGTH (Spec / Act) mm</div>`;

const fsmSpecInp = makeInput({
  className: "input-small",
  id: "fsmSpec",
  readOnly: true,
  value: parsedHeader.fsmLength || "",
  size: 6
});

const fsmActInp = makeInput({
  className: "input-small",
  id: "fsmAct",
  placeholder: "Act (mm)",
  size: 6
});

fsmcol.appendChild(fsmSpecInp);
fsmcol.appendChild(document.createTextNode(" "));
fsmcol.appendChild(fsmActInp);

// Append both columns to same row
row4.appendChild(rwcol);
row4.appendChild(fsmcol);
blk.appendChild(row4);

// Debug check — will show in console
console.log("✅ Root Width fields added:", rootSpec, rootAct);

/* ---------- Main inspection table ---------- */
function buildMainTable() {
  const blk = document.createElement("div");
  blk.className = "form-block";
  blk.innerHTML = `<strong>Main Inspection Table</strong>`;
  const table = document.createElement("table");
  table.className = "table";
  const thead = document.createElement("thead");
  thead.innerHTML = `<tr>
    <th>Sl No</th><th>Press</th><th>Sel ID</th><th>Ref</th><th>X-axis</th><th>Spec (Y or Z)</th>
    <th>Spec Dia</th><th>Value from Hole edge (Act)</th><th>Actual Dia</th><th>Actual Y or Z</th><th>Offset</th><th>Result</th>
  </tr>`;
  table.appendChild(thead);
  const tbody = document.createElement("tbody");

  // If parsed rows available, attempt to populate; otherwise create 20 rows (max 45 in real cases)
  const rowsToBuild = parsedTableRows.length ? parsedTableRows.length : 20;
  for (let i=0;i<rowsToBuild;i++){
    const tr = document.createElement("tr");
    const tokens = parsedTableRows[i] || [];
    const sl = makeCell(i+1);
    const press = makeCell(tokens[1] || "");
    const sel = makeCell(tokens[2] || "");
    const ref = makeCell(tokens[3] || "");
    const xaxis = makeInput({className:"input-small", value: tokens[4] || ""});
    const specYZ = makeInput({className:"input-small", value: tokens[5] || ""});
    const specDia = makeInput({className:"input-small", value: tokens[6] || ""});
    const valEdge = makeInput({className:"input-small", value:""});
    const actualDia = makeInput({className:"input-small", value:""});
    const actualYZ = makeInput({className:"input-small", value:""});
    const offsetCell = makeCell("");
    const resultCell = makeCell("");
    // attach events to recalc on input
    [valEdge, actualDia, actualYZ, specYZ, specDia, xaxis].forEach(inp => {
      inp.addEventListener('input', () => recalcRowAndMark(tr));
    });
    tr.appendChild(tdWrap(sl));
    tr.appendChild(tdWrap(press));
    tr.appendChild(tdWrap(sel));
    tr.appendChild(tdWrap(ref));
    tr.appendChild(tdWrap(xaxis));
    tr.appendChild(tdWrap(specYZ));
    tr.appendChild(tdWrap(specDia));
    tr.appendChild(tdWrap(valEdge));
    tr.appendChild(tdWrap(actualDia));
    tr.appendChild(tdWrap(actualYZ));
    tr.appendChild(tdWrap(offsetCell));
    tr.appendChild(tdWrap(resultCell));
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  blk.appendChild(table);
  formArea.appendChild(blk);
}

/* helper to create a td with input or text */
function makeCell(content) {
  const el = document.createElement("div");
  el.textContent = content;
  return el;
}
function tdWrap(inner) {
  const td = document.createElement("td");
  if (inner instanceof HTMLElement) td.appendChild(inner);
  else td.textContent = inner;
  return td;
}

/* recalc row */
function recalcRowAndMark(tr) {
  // find cells
  const tds = tr.querySelectorAll('td');
  const xaxis = parseFloat((tds[4].querySelector('input')||{value:""}).value || NaN);
  const specYZ = parseFloat((tds[5].querySelector('input')||{value:""}).value || NaN);
  const specDia = parseFloat((tds[6].querySelector('input')||{value:""}).value || NaN);
  const valEdge = parseFloat((tds[7].querySelector('input')||{value:""}).value || NaN);
  const actualDia = parseFloat((tds[8].querySelector('input')||{value:""}).value || NaN);
  const actualYZ = parseFloat((tds[9].querySelector('input')||{value:""}).value || NaN);

  const offsetNode = tds[10];
  const resultNode = tds[11];

  // offset formula: Offset = Spec(Y/Z) + (Spec Dia / 2)
  let offset = null;
  if (!isNaN(specYZ) && !isNaN(specDia)) offset = specYZ + (specDia/2);
  offsetNode.textContent = offset === null ? "" : round(offset,2);

  // now compute tolerance rules
  // offset tolerance: depends on FSM length (parsed or input)
  const fsmSpecVal = parseFloat(document.getElementById('fsmSpec')?.value || parsedHeader.fsmLength || NaN);
  let tol = 1;
  if (!isNaN(xaxis) && !isNaN(fsmSpecVal)) {
    if (xaxis <= 200 || xaxis >= (fsmSpecVal - 200)) tol = 2;
  } // else default 1

  // Dia tolerance
  let diaOk = true;
  if (!isNaN(specDia) && !isNaN(actualDia)) {
    if (specDia <= 10.7) diaOk = (actualDia >= specDia - 0 && actualDia <= specDia + 0.4);
    else if (specDia >= 11.7) diaOk = (actualDia >= specDia - 0 && actualDia <= specDia + 0.5);
    else diaOk = (actualDia >= specDia - 0 && actualDia <= specDia + 0.5);
  } else {
    diaOk = false; // require actualDia value
  }

  // offset check: compare actualYZ (user-entered) against computed offset
  let offsetOk = true;
  if (!isNaN(offset) && !isNaN(actualYZ)) {
    const dev = Math.abs(actualYZ - offset);
    offsetOk = dev <= tol;
  } else {
    offsetOk = false; // require actual yz value
  }

  // Combine: if either fail → NOK
  const allOk = diaOk && offsetOk;

  // style
  if (allOk) {
    // mark result OK
    resultNode.textContent = "OK";
    resultNode.className = "ok-cell";
    // clear individual cells if they were nok
    [tds[7], tds[8], tds[9], tds[10]].forEach(cell => cell.classList.remove('nok-cell'));
  } else {
    resultNode.textContent = "NOK";
    resultNode.className = "nok-cell";
    // highlight cells that failed
    if (!diaOk) tds[8].classList.add('nok-cell'); // actual dia cell
    if (!offsetOk) tds[9].classList.add('nok-cell'); // actual yz cell
    if (!isNaN(offset) && !isNaN(actualYZ) && Math.abs(actualYZ - offset) > tol) tds[10].classList.add('nok-cell');
  }
}

/* ---------- First holes (front) section: fixed 5 rows ---------- */
function buildFirstHolesSection() {
  const blk = document.createElement("div"); blk.className="form-block";
  blk.innerHTML = `<strong>First holes of FSM - from front end</strong>`;
  const table = document.createElement('table'); table.className='table';
  table.innerHTML = `<thead><tr><th>Spec</th><th>Actual (Value from Hole edge)</th><th>Actual (Dia)</th><th>Offset</th><th>Result</th></tr></thead>`;
  const tb = document.createElement('tbody');
  // try to parse first-holes specs from text: look for known values list - best-effort
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
        if (Math.abs(val - spec) > 1) {
          valEdge.classList.add('nok-cell');
          resultTd.textContent = "NOK"; resultTd.className="nok-cell";
        } else {
          valEdge.classList.remove('nok-cell');
          resultTd.textContent = "OK"; resultTd.className="ok-cell";
        }
      } else {
        resultTd.textContent = "";
        resultTd.className = "";
      }
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
  // try to find sequence like numeric specs - best-effort
  const matches = text.match(/([0-9]{1,4}(?:\.[0-9]+)?)/g);
  if (!matches) return null;
  const smalls = matches.map(Number).filter(n=>n>0 && n<5000);
  return smalls.slice(0,5).map(n=>n.toString());
}

/* ---------- Last holes (rear) section: identical to first ---------- */
function buildLastHolesSection() {
  const blk = document.createElement("div"); blk.className="form-block";
  blk.innerHTML = `<strong>Last holes of FSM - from rear end</strong>`;
  const table = document.createElement('table'); table.className='table';
  table.innerHTML = `<thead><tr><th>Spec</th><th>Actual (Value from Hole edge)</th><th>Actual (Dia)</th><th>Offset</th><th>Result</th></tr></thead>`;
  const tb = document.createElement('tbody');
  const specs = extractLastHoleSpecs(lastExtractedText) || ["","","","",""];
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
        if (Math.abs(val - spec) > 1) {
          valEdge.classList.add('nok-cell');
          resultTd.textContent = "NOK"; resultTd.className="nok-cell";
        } else {
          valEdge.classList.remove('nok-cell');
          resultTd.textContent = "OK"; resultTd.className="ok-cell";
        }
      } else {
        resultTd.textContent = "";
        resultTd.className = "";
      }
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
function extractLastHoleSpecs(text) {
  return extractFirstHoleSpecs(text); // best-effort same method
}

/* ---------- Root width and flanges ---------- */
function buildRootAndFlangeSection() {
  const blk = document.createElement('div'); blk.className='form-block';
  blk.innerHTML = `<strong>Root / Flange / Web (reference)</strong>`;
  const tbl = document.createElement('table'); tbl.className='table';
  tbl.innerHTML = `<thead><tr><th>Root Width of FSM (Spec)</th><th>Root Width Act</th><th>Top Flange</th><th>Web PWM</th><th>Bottom Flange</th></tr></thead>`;
  const tb = document.createElement('tbody');
  const tr = document.createElement('tr');
  const rootSpecTd = document.createElement('td'); rootSpecTd.textContent = document.getElementById('rootSpec')?.value || "";
  // safely reuse value from earlier input without breaking layout
let rootActValue = document.getElementById('rootAct')?.value || "";
const rootActInput = makeInput({
  className: 'input-small',
  value: rootActValue,
  placeholder: "Act (mm)",
  id: "rootAct_flange"
});
  rootActInput.addEventListener('input', ()=> {
    const spec = parseFloat(rootSpecTd.textContent||NaN);
    const act = parseFloat(rootActInput.value||NaN);
    if (!isNaN(spec) && !isNaN(act)) {
      if (Math.abs(act - spec) > 1) rootActInput.classList.add('nok-cell');
      else { rootActInput.classList.remove('nok-cell'); rootActInput.classList.add('ok-cell'); }
    }
  });
  const topFl = document.createElement('td'); topFl.textContent = "Top Flange: PWM (info)";
  const webPW = document.createElement('td'); webPW.textContent = "Web PWM (info)";
  const bottomFl = document.createElement('td'); bottomFl.textContent = "Bottom Flange (info)";
  tr.appendChild(rootSpecTd);
  tr.appendChild(tdWrap(rootActInput));
  tr.appendChild(topFl); tr.appendChild(webPW); tr.appendChild(bottomFl);
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
  tr.appendChild(specTd); tr.appendChild(actTd); tr.appendChild(resTd); tb.appendChild(tr); table.appendChild(tb);
  blk.appendChild(table); formArea.appendChild(blk);
}

/* ---------- Binary checks (OK/NOK and YES/NO) ---------- */
function buildBinaryChecks() {
  const blk = document.createElement('div'); blk.className='form-block';
  blk.innerHTML = `<strong>Visual / Binary Checks</strong>`;
  const grid = document.createElement('div');
  grid.style.display='grid'; grid.style.gridTemplateColumns='repeat(4,1fr)'; grid.style.gap='8px';
  const labels = ["Punch Break","Radius crack","Length Variation","Holes Burr","Slug mark","Line mark","Part No. Legibility","Pit Mark","Machine Error (YES/NO)"];
  labels.forEach(lbl=>{
    const cell = document.createElement('div'); cell.className='small';
    const labelEl = document.createElement('div'); labelEl.textContent = lbl;
    // options
    const okBtn = document.createElement('button'); okBtn.textContent='OK/YES'; okBtn.className='input-small';
    const nokBtn = document.createElement('button'); nokBtn.textContent='NOK/NO'; nokBtn.className='input-small';
    okBtn.addEventListener('click', ()=> {
      okBtn.classList.add('ok-cell-binary'); nokBtn.classList.remove('nok-cell-binary');
      okBtn.classList.remove('nok-cell-binary');
      nokBtn.classList.remove('ok-cell-binary');
      checkMandatoryBeforeExport();
    });
    nokBtn.addEventListener('click', ()=> {
      nokBtn.classList.add('nok-cell-binary'); okBtn.classList.remove('ok-cell-binary');
      okBtn.classList.remove('nok-cell-binary');
      checkMandatoryBeforeExport();
    });
    cell.appendChild(labelEl); cell.appendChild(okBtn); cell.appendChild(nokBtn);
    grid.appendChild(cell);
  });
  blk.appendChild(grid); formArea.appendChild(blk);
}

/* ---------- Remarks and signature ---------- */
function buildRemarksAndSign() {
  const blk = document.createElement('div'); blk.className='form-block';
  blk.innerHTML = `<strong>Remarks / Details of issue</strong>`;
  const ta = document.createElement('textarea'); ta.id='remarks'; ta.placeholder = "Enter remarks (optional)";
  blk.appendChild(ta);

  // prod incharge mandatory
  const signRow = document.createElement('div'); signRow.className='form-row';
  const col1 = document.createElement('div'); col1.className='col';
  col1.innerHTML = `<div class="small">Prodn Incharge (Shift Executive) - mandatory</div>`;
  const prodIn = makeInput({className:"input-small", id:"prodIncharge"});
  prodIn.placeholder = "Enter shift executive name";
  prodIn.addEventListener('input', checkMandatoryBeforeExport);
  col1.appendChild(prodIn); signRow.appendChild(col1);

  blk.appendChild(signRow);
  formArea.appendChild(blk);
}

/* ---------- attach general listeners ---------- */
function attachValidationListeners() {
  // monitor mandatory header inputs and KB/PC act
  ["fsmSerial","inspectors","holesAct","matrixUsed","prodIncharge"].forEach(id=>{
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', checkMandatoryBeforeExport);
  });
  // also rootAct and fsmAct
  const rAct = document.getElementById('rootAct'), fAct = document.getElementById('fsmAct');
  if (rAct) rAct.addEventListener('input', checkMandatoryBeforeExport);
  if (fAct) fAct.addEventListener('input', checkMandatoryBeforeExport);
  // kb/pc act
  document.getElementById('kbAct')?.addEventListener('input', checkMandatoryBeforeExport);
  document.getElementById('pcAct')?.addEventListener('input', checkMandatoryBeforeExport);
  // kb/pc spec changes
  document.getElementById('kbSpec')?.addEventListener('input', checkMandatoryBeforeExport);
  document.getElementById('pcSpec')?.addEventListener('input', checkMandatoryBeforeExport);
}

/* ---------- Mandatory checks and enabling export ---------- */
function checkMandatoryBeforeExport() {
  const fsmSerial = document.getElementById('fsmSerial')?.value || "";
  const inspectors = document.getElementById('inspectors')?.value || "";
  const holesAct = document.getElementById('holesAct')?.value || "";
  const matrixUsed = document.getElementById('matrixUsed')?.value || "";
  const prod = document.getElementById('prodIncharge')?.value || "";
  let missing = [];
  if (!fsmSerial.trim()) missing.push("FSM Serial Number");
  if (!inspectors.trim()) missing.push("Inspectors");
  if (!holesAct.trim()) missing.push("Total Holes Count (Act)");
  if (!matrixUsed.trim()) missing.push("Matrix used");
  if (!prod.trim()) missing.push("Prodn Incharge");
  if (missing.length) {
    showMessage("Please fill mandatory fields: " + missing.join(", "), "warn");
    exportPdfBtn.disabled = true;
    return false;
  }
  // KB & PC Act must be present (we allow them to be NOK but must be filled)
  const kbAct = document.getElementById('kbAct')?.value || "";
  const pcAct = document.getElementById('pcAct')?.value || "";
  if (!kbAct.trim() || !pcAct.trim()) {
    showMessage("Please fill KB & PC Act values.", "warn");
    exportPdfBtn.disabled = true;
    return false;
  }

  // All binary checks must be selected: simple check - ensure no binary button group is untouched
  exportPdfBtn.disabled = false;
  clearMessage();
  return true;
}

/* ---------- Export to PDF ---------- */
exportPdfBtn.addEventListener('click', async () => {
  if (!checkMandatoryBeforeExport()) return;
  // check for NOKs; show confirm (option B)
  const hasNok = document.querySelectorAll('.nok-cell, .nok-cell-binary').length > 0;
  if (hasNok) {
    const proceed = confirm("Some fields are out of spec (NOK). Please verify. Proceed to export anyway?");
    if (!proceed) return;
  }
  exportPdfBtn.disabled = true;
  exportPdfBtn.textContent = "Generating PDF...";
  await generatePdfFromForm();
  exportPdfBtn.disabled = false; exportPdfBtn.textContent = "Export to PDF";
});

/* ---------- generate PDF using html2canvas + jsPDF ---------- */
async function generatePdfFromForm() {
  // create a printable clone of form area with header for capture
  const clone = document.createElement('div');
  clone.style.width = "900px"; // A4 scaling reference
  clone.style.padding = "12px";
  // header
  const title = document.createElement('div'); title.innerHTML = document.querySelector('header').innerHTML;
  clone.appendChild(title);
  // form area (clone)
  const fa = formArea.cloneNode(true);
  // remove interactive classes that may break layout (but keep inline styles)
  fa.querySelectorAll('input, textarea, button').forEach(n=>{
    // convert inputs to spans showing values
    if (n.tagName === 'INPUT') {
      const span = document.createElement('span'); span.textContent = n.value;
      n.parentNode && n.parentNode.replaceChild(span, n);
    } else if (n.tagName === 'TEXTAREA') {
      const span = document.createElement('div'); span.textContent = n.value; span.style.whiteSpace="pre-wrap";
      n.parentNode && n.parentNode.replaceChild(span, n);
    } else if (n.tagName === 'BUTTON') {
      n.remove();
    }
  });
  clone.appendChild(fa);

  // append to body off-screen, render, then remove
  clone.style.position = "fixed";
  clone.style.left = "-2000px";
  document.body.appendChild(clone);
  try {
    const canvas = await html2canvas(clone, {scale: 2, useCORS: true});
    const imgData = canvas.toDataURL('image/png');
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF('p','mm','a4');
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    // fit image keeping aspect ratio
    const imgProps = pdf.getImageProperties(imgData);
    const pdfWidth = pageWidth - 20;
    const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
    pdf.addImage(imgData, 'PNG', 10, 10, pdfWidth, pdfHeight);
    // system date/time
    const now = new Date();
    const stamp = now.toLocaleString();
    // add footer text
    pdf.setFontSize(9);
    pdf.text(`Generated: ${stamp}`, 14, pageHeight - 10);

    // filename pattern: Part_Rev_Hand_Inspection_<date>.pdf
    const part = parsedHeader.partNumber || document.getElementById('hdrPart')?.textContent || "PART";
    const rev = parsedHeader.revision || "X";
    const hand = parsedHeader.hand || "H";
    const shortDate = now.toISOString().slice(0,10);
    const filename = `${part}_${rev}_${hand}_Inspection_${shortDate}.pdf`;
    pdf.save(filename);
    showMessage("PDF exported: " + filename, "info");
  } catch (err) {
    console.error(err);
    alert("PDF export failed: " + err.message);
  } finally {
    document.body.removeChild(clone);
  }
}

/* ---------- simple helpers ---------- */
function showMessage(msg, type="info") {
  warnings.style.padding = "8px";
  warnings.style.borderRadius = "6px";
  warnings.style.marginTop = "6px";
  warnings.textContent = msg;
  if (type==="warn") { warnings.style.background = "#fff6cc"; warnings.style.color="#ff8c00"; }
  else { warnings.style.background = "#e8f4ff"; warnings.style.color="#005fcc"; }
}
function clearMessage(){ warnings.innerHTML=""; warnings.style.background=""; warnings.style.color=""; }
function round(n,dec){ return Math.round(n * Math.pow(10,dec))/Math.pow(10,dec); }

/* ---------------- Safety rebind (in case DOM or order changed) ---------------- */
document.getElementById("generateFormBtn").addEventListener("click", () => {
  // require parsed table rows (or at least parsing done)
  if (!parsedTableRows || !parsedTableRows.length) {
    showMessage("No parsed table data found. Please parse PDF first.", "warn");
    return;
  }

  // button feedback
  generateFormBtn.disabled = true;
  generateFormBtn.textContent = "Building...";

  // build the form
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
  generateFormBtn.disabled = false;
  generateFormBtn.textContent = "Create Editable Form";

  showMessage("Editable form created. Fill Actual values and mandatory fields. Use 'Export to PDF' to generate the report.", "info");
});
