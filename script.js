/* script.js - Phase 2 (complete replacement)
   - Render PDF preview (pdf.js)
   - Extract text for best-effort parsing
   - Build editable form with up to 45 table rows (prefill first 7 columns readonly)
   - Validations + export to PDF (html2canvas + jsPDF)
*/

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
let parsedHeader = {};       // parsed header object
let parsedTableRows = [];    // parsed table data (array of arrays/tokens)
const MAX_TABLE_ROWS = 45;   // keep this exact so form always supports up to 45 rows

/* ---------------------- Helpers ---------------------- */
function showMessage(msg, type = "info") {
  warnings.style.display = "block";
  warnings.style.padding = "8px";
  warnings.style.borderRadius = "6px";
  warnings.style.marginTop = "6px";
  warnings.textContent = msg;
  if (type === "warn") {
    warnings.style.background = "#fff6cc";
    warnings.style.color = "#ff8c00";
  } else {
    warnings.style.background = "#e8f4ff";
    warnings.style.color = "#005fcc";
  }
}
function clearMessage() {
  warnings.style.display = "none";
  warnings.textContent = "";
  warnings.style.background = "";
  warnings.style.color = "";
}
function round(n, dec = 2) {
  if (typeof n !== "number" || isNaN(n)) return n;
  return Math.round(n * Math.pow(10, dec)) / Math.pow(10, dec);
}

/* ---------------------- PDF load / preview / extract ---------------------- */
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
  clearMessage();
});

extractBtn.addEventListener("click", async () => {
  if (!pdfFile) {
    alert("Please select a PDF first.");
    return;
  }
  extractBtn.disabled = true;
  extractBtn.textContent = "Parsing...";
  clearMessage();
  try {
    const buf = await pdfFile.arrayBuffer();
    const text = await extractTextFromPDF(new Uint8Array(buf));
    lastExtractedText = text;
    parsedHeader = parseHeader(text);
    headerOut.textContent = JSON.stringify(parsedHeader, null, 2);

    // update header HUD
    updateHeaderFields(parsedHeader);

    // detect table rows (best-effort)
    parsedTableRows = detectTableLines(text);
    // enable create form button after parsing
    generateFormBtn.disabled = false;
    showMessage("Parsed PDF. Click 'Create Editable Form' to build the inspection form (you can edit any spec).", "info");
  } catch (err) {
    console.error("Parsing failed:", err);
    alert("Error while parsing PDF: " + (err.message || err));
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
  } catch (e) {
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

/* ---------------------- Parsing helpers ---------------------- */
function updateHeaderFields(hdr) {
  if (!hdr) return;
  if (hdrPart) {
    const partNumber = hdr.partNumber || "—";
    const revision = hdr.revision || "—";
    const hand = hdr.hand || "—";
    hdrPart.textContent = `PART NUMBER / LEVEL / HAND : ${partNumber} / ${revision} / ${hand}`;
  }
  if (hdrFsmSpec) {
    hdrFsmSpec.textContent = `FSM LENGTH : ${hdr.fsmLength || "—"} mm`;
  }
}

function parseHeader(text) {
  const h = {
    partNumber: null, revision: null, hand: null, date: null, formatNo: null,
    rootWidth: null, fsmLength: null, kbSpec: null, pcSpec: null, holesCount: null
  };
  let m;

  // Part / Rev / Hand (best-effort)
  m = text.match(/PART\s*NUMBER\s*\/\s*LEVEL\s*\/\s*HAND\s*[:\s\-]*([A-Z0-9_\-]+)\s*\/\s*([A-Z0-9_\-]+)\s*\/\s*([A-Z0-9\-]+)/i);
  if (m) { h.partNumber = m[1]; h.revision = m[2]; h.hand = m[3]; }

  // fallback to part only
  if (!h.partNumber) {
    m = text.match(/PART\s*NUMBER\s*[:\s\-]*([A-Z0-9_\-]+)/i);
    if (m) h.partNumber = m[1];
  }

  // Date (best-effort)
  m = text.match(/Date\s*[:\-\s]*([0-3]?\d[-\/][A-Za-z]{3,}[-\/]\d{2,4})/i);
  if (m) h.date = m[1];

  // root width (try several patterns)
  m = text.match(/ROOT\s*WIDTH\s*OF\s*FSM\s*[:\-\s]*[Ss]pec[:\s\-]*([0-9.]+)/i) ||
      text.match(/ROOT\s*WIDTH\s*OF\s*FSM\s*[:\-\s]*([0-9.]+)/i);
  if (m) h.rootWidth = parseFloat(m[1]);

  // fsm length
  m = text.match(/FSM\s*LENGTH\s*[:\-\s]*[Ss]pec[:\s\-]*([0-9.]+)/i) || text.match(/FSM\s*LENGTH\s*[:\-\s]*([0-9.]+)/i);
  if (m) h.fsmLength = parseFloat(m[1]);

  // KB & PC spec pattern like "KB & PC Code : Spec- 5 / 1" (best-effort)
  m = text.match(/KB\s*&\s*PC\s*Code\s*[:\-\s]*\s*Spec[-\s]*([0-9]+)\s*\/\s*([0-9]+)/i);
  if (m) { h.kbSpec = m[1]; h.pcSpec = m[2]; }

  // holes count (Total holes count)
  m = text.match(/TOTAL\s*HOLES\s*COUNT\s*[:\-\s]*([0-9]+)/i);
  if (m) h.holesCount = parseInt(m[1], 10);

  // format no
  m = text.match(/FORMAT\s*NO\.?\s*[:\-]*\s*([A-Z0-9\/\s\-\_]+)/i);
  if (m) h.formatNo = (m[1] || "").trim();

  return h;
}

/* Very small best-effort table detector:
   We try to find sequences that look like: "<index> <Press> <SelID> <Ref> <X-axis> <Spec> <SpecDia>"
   If we can't find, keep parsedTableRows empty (we still build blank rows later).
*/
function detectTableLines(fullText) {
  if (!fullText) return [];
  const rows = [];
  // Strategy:
  //  - split fullText into lines and try to identify lines that start with a row index (1..45)
  //  - for each such line, attempt to extract tokens (numbers & short strings)
  const rawLines = fullText.split(/\n|---PAGE_BREAK---/).map(l => l.trim()).filter(Boolean);
  for (const line of rawLines) {
    // Trim repeated whitespace
    const cleaned = line.replace(/\s{2,}/g, " ");
    // If starts with a number and has at least 6 tokens, consider it a table row candidate
    if (/^\s*\d+(\.|:)?\s+/.test(cleaned)) {
      // remove possible leading "1." etc
      const tokens = cleaned.replace(/^\s*\d+(\.|:)?\s*/, "").split(/\s+/);
      // tokens may include press (PW/PWX/etc), sel id, ref (letter), x axis number, spec number, spec dia.
      // We'll collect the first 7 logical values from tokens (some may be non-numeric).
      const extracted = [];
      // attempt to pick tokens that look like the expected columns in order:
      // Press (word), SelID (number), Ref (letter), X-axis (number), Spec (number), SpecDia (number)
      // We'll iterate tokens and push appropriate ones
      let idx = 0;
      for (let i = 0; i < tokens.length && extracted.length < 7; i++) {
        const tk = tokens[i];
        // Accept Press if alphabetic or contains PW/PF etc
        if (extracted.length === 0 && /^[A-Za-z]{1,4}$/.test(tk) || /^[A-Za-z]{1,4}\d*$/.test(tk)) {
          extracted.push(tk);
          continue;
        }
        // Sel ID: small number
        if (extracted.length === 0 && /^\d+$/.test(tk)) { extracted.push(tk); continue; }
        if (extracted.length === 1 && /^\d+$/.test(tk)) { extracted.push(tk); continue; }
        // Ref: letter(s)
        if (extracted.length === 2 && /^[A-Za-z]$/.test(tk)) { extracted.push(tk); continue; }
        // X-axis: number with possible decimals
        if (extracted.length === 3 && /^[0-9]+(?:\.[0-9]+)?$/.test(tk)) { extracted.push(tk); continue; }
        // Spec (Y or Z) numeric
        if (extracted.length === 4 && /^[0-9]+(?:\.[0-9]+)?$/.test(tk)) { extracted.push(tk); continue; }
        // Spec Dia numeric
        if (extracted.length === 5 && /^[0-9]+(?:\.[0-9]+)?$/.test(tk)) { extracted.push(tk); continue; }
        // fallback: push anything if position missing
        if (extracted.length < 7) extracted.push(tk);
      }
      // If we constructed something that looks reasonable (>= 4 tokens), push it
      if (extracted.length >= 4) {
        // Prepend the detected row index if present at start
        const idxMatch = cleaned.match(/^\s*(\d+)(?:\.|:)?/);
        const sl = idxMatch ? idxMatch[1] : (rows.length + 1).toString();
        rows.push([sl, ...(extracted.slice(0, 6))]); // keep 7 values total (sl + 6 tokens)
      }
    }
  }

  // If nothing found, return empty array and allow manual building
  return rows;
}

/* ---------------------- UI input factory ---------------------- */
function makeInput(attrs = {}) {
  const inp = document.createElement("input");
  inp.type = attrs.type || "text";
  if (attrs.value !== undefined && attrs.value !== null) inp.value = attrs.value;
  inp.className = attrs.className || "input-small";
  if (attrs.placeholder) inp.placeholder = attrs.placeholder;
  if (attrs.step) inp.step = attrs.step;
  if (attrs.onchange) inp.addEventListener("change", attrs.onchange);
  if (attrs.oninput) inp.addEventListener("input", attrs.oninput);
  if (attrs.readOnly) inp.readOnly = true;
  if (attrs.size) inp.size = attrs.size;
  if (attrs.id) inp.id = attrs.id;
  return inp;
}

/* ---------------------- Build header inputs ---------------------- */
function buildHeaderInputs() {
  const blk = document.createElement("div");
  blk.className = "form-block";
  blk.innerHTML = `<strong>Header / Mandatory fields</strong>`;
  // Row 1: FSM Serial (left) + Inspectors (right)
  const row = document.createElement("div"); row.className = "form-row";
  // FSM Serial
  const c1 = document.createElement("div"); c1.className = "col";
  c1.innerHTML = `<div class="small">FSM Serial Number:</div>`;
  const fsmSerial = makeInput({ placeholder: "Enter FSM Serial No", id: "fsmSerial" });
  c1.appendChild(fsmSerial);
  // Inspectors (allow two)
  const c2 = document.createElement("div"); c2.className = "col";
  c2.innerHTML = `<div class="small">Inspectors:</div>`;
  const inspector1 = makeInput({ placeholder: "Inspector 1", id: "inspector1" });
  const inspector2 = makeInput({ placeholder: "Inspector 2", id: "inspector2" });
  c2.appendChild(inspector1);
  c2.appendChild(document.createTextNode(" "));
  c2.appendChild(inspector2);

  row.appendChild(c1);
  row.appendChild(c2);
  blk.appendChild(row);

  // Row 2: Total holes count / Matrix used
  const row2 = document.createElement("div"); row2.className = "form-row";
  const c3 = document.createElement("div"); c3.className = "col";
  c3.innerHTML = `<div class="small">TOTAL HOLES COUNT (Spec / Act):</div>`;
  const holesSpec = makeInput({ id: "holesSpec", readOnly: false, placeholder: "Spec (if parsed)" });
  const holesAct = makeInput({ id: "holesAct", placeholder: "Act (mandatory)" });
  holesSpec.value = parsedHeader.holesCount || "";
  c3.appendChild(holesSpec);
  c3.appendChild(document.createTextNode(" "));
  c3.appendChild(holesAct);

  const c4 = document.createElement("div"); c4.className = "col";
  c4.innerHTML = `<div class="small">Matrix used:</div>`;
  const matrixUsed = makeInput({ id: "matrixUsed", placeholder: "Matrix" });
  c4.appendChild(matrixUsed);

  row2.appendChild(c3); row2.appendChild(c4);
  blk.appendChild(row2);

  // Row 3: KB & PC compact
  const row3 = document.createElement("div"); row3.className = "form-row";
  const kbcol = document.createElement("div"); kbcol.className = "col";
  kbcol.innerHTML = `<div class="small">KB & PC Code (Spec / Act)</div>`;

  const kbSpec = makeInput({ id: "kbSpec", readOnly: true, value: parsedHeader.kbSpec || "", size: 3 });
  const kbAct = makeInput({ id: "kbAct", placeholder: "Act", size: 3 });
  const pcSpec = makeInput({ id: "pcSpec", readOnly: true, value: parsedHeader.pcSpec || "", size: 3 });
  const pcAct = makeInput({ id: "pcAct", placeholder: "Act", size: 3 });

  kbcol.appendChild(kbSpec);
  kbcol.appendChild(document.createTextNode(" / "));
  kbcol.appendChild(kbAct);
  kbcol.appendChild(document.createTextNode("    "));
  kbcol.appendChild(pcSpec);
  kbcol.appendChild(document.createTextNode(" / "));
  kbcol.appendChild(pcAct);
  row3.appendChild(kbcol);
  blk.appendChild(row3);

  // Row 4: Root Width + FSM length (spec readonly, act editable)
  const row4 = document.createElement("div"); row4.className = "form-row";

  const rwcol = document.createElement("div"); rwcol.className = "col";
  rwcol.innerHTML = `<div class="small">ROOT WIDTH OF FSM (Spec / Act) mm</div>`;
  const rootSpec = makeInput({ id: "rootSpec", readOnly: true, value: parsedHeader.rootWidth || "" });
  const rootAct = makeInput({ id: "rootAct", placeholder: "Act (mm)" });
  rwcol.appendChild(rootSpec);
  rwcol.appendChild(document.createTextNode(" "));
  rwcol.appendChild(rootAct);

  const fsmcol = document.createElement("div"); fsmcol.className = "col";
  fsmcol.innerHTML = `<div class="small">FSM LENGTH (Spec / Act) mm</div>`;
  const fsmSpecInp = makeInput({ id: "fsmSpec", readOnly: true, value: parsedHeader.fsmLength || "" });
  const fsmActInp = makeInput({ id: "fsmAct", placeholder: "Act (mm)" });
  fsmcol.appendChild(fsmSpecInp);
  fsmcol.appendChild(document.createTextNode(" "));
  fsmcol.appendChild(fsmActInp);

  row4.appendChild(rwcol);
  row4.appendChild(fsmcol);
  blk.appendChild(row4);

  // append to form area
  formArea.appendChild(blk);
}

/* ---------------------- Build Main Inspection Table ---------------------- */
function buildMainTable() {
  const blk = document.createElement("div");
  blk.className = "form-block";
  blk.innerHTML = `<strong>Main Inspection Table</strong>`;

  const table = document.createElement("table");
  table.className = "table";
  const thead = document.createElement("thead");
  thead.innerHTML = `<tr>
    <th>Sl No.</th><th>Press</th><th>Sel ID</th><th>Ref</th><th>X-axis</th><th>Spec (Y or Z)</th>
    <th>Spec Dia</th><th>Value from Hole edge (Act)</th><th>Actual Dia</th><th>Actual Y or Z</th><th>Offset</th><th>Result</th>
  </tr>`;
  table.appendChild(thead);

  const tbody = document.createElement("tbody");

  // rows: up to MAX_TABLE_ROWS (45). If parsedTableRows has data, use it to prefill the first columns
  for (let r = 0; r < MAX_TABLE_ROWS; r++) {
    const tr = document.createElement("tr");

    // Sl No: always present and readonly
    const slTd = document.createElement("td");
    const slInput = makeInput({ value: (r + 1), readOnly: true });
    slTd.appendChild(slInput);
    tr.appendChild(slTd);

    // Determine parsed tokens (if available)
    const parsedRow = parsedTableRows[r] || [];
    // parsedRow expected shape: [sl, press, selid, ref, xaxis, specYZ, specDia] (some fields may be missing)

    // Helper to create readonly prefilled cell or blank input
    const makePrefilledCell = (prefillVal) => {
      const td = document.createElement("td");
      const inp = makeInput({ value: prefillVal !== undefined ? prefillVal : "", readOnly: true });
      // visually mark readonly inputs to be compact
      inp.classList.add("input-small");
      td.appendChild(inp);
      return td;
    };

    // Press (col 2) - readonly
    const pressTd = makePrefilledCell(parsedRow[1] || "");
    tr.appendChild(pressTd);

    // Sel ID (col 3) - readonly
    const selTd = makePrefilledCell(parsedRow[2] || "");
    tr.appendChild(selTd);

    // Ref (col 4) - readonly
    const refTd = makePrefilledCell(parsedRow[3] || "");
    tr.appendChild(refTd);

    // X-axis (col 5) - readonly (prefill numeric value if parsed)
    const xaxisTd = makePrefilledCell(parsedRow[4] || "");
    tr.appendChild(xaxisTd);

    // Spec (Y or Z) (col 6) - readonly
    const specYZTd = makePrefilledCell(parsedRow[5] || "");
    tr.appendChild(specYZTd);

    // Spec Dia (col 7) - readonly
    const specDiaTd = makePrefilledCell(parsedRow[6] || "");
    tr.appendChild(specDiaTd);

    // From here onward: editable inputs (Act fields)
    const valEdgeTd = document.createElement("td");
    const valEdgeInp = makeInput({ placeholder: "Value", className: "input-small" });
    valEdgeTd.appendChild(valEdgeInp);
    tr.appendChild(valEdgeTd);

    const actualDiaTd = document.createElement("td");
    const actualDiaInp = makeInput({ placeholder: "Dia", className: "input-small" });
    actualDiaTd.appendChild(actualDiaInp);
    tr.appendChild(actualDiaTd);

    const actualYZTd = document.createElement("td");
    const actualYZInp = makeInput({ placeholder: "Y / Z", className: "input-small" });
    actualYZTd.appendChild(actualYZInp);
    tr.appendChild(actualYZTd);

    // Offset and Result columns (computed)
    const offsetTd = document.createElement("td");
    offsetTd.textContent = "";
    tr.appendChild(offsetTd);

    const resultTd = document.createElement("td");
    resultTd.textContent = "";
    tr.appendChild(resultTd);

    // Attach input events to recalc
    [valEdgeInp, actualDiaInp, actualYZInp].forEach(inp => {
      inp.addEventListener("input", () => recalcRowAndMark(tr));
    });

    tbody.appendChild(tr);
  } // end rows

  table.appendChild(tbody);
  blk.appendChild(table);
  formArea.appendChild(blk);
}

/* Recalc logic per row (keeps earlier rules) */
function recalcRowAndMark(tr) {
  const tds = tr.querySelectorAll("td");

  // Read spec values from the readonly inputs in the row
  // Col indices mapping:
  // 0 Sl, 1 Press, 2 SelID, 3 Ref, 4 X-axis, 5 SpecYZ, 6 SpecDia, 7 ValEdge, 8 ActualDia, 9 ActualYZ, 10 Offset, 11 Result
  const specYZ = parseFloat((tds[5].querySelector("input") || {}).value || NaN);
  const specDia = parseFloat((tds[6].querySelector("input") || {}).value || NaN);
  const valEdge = parseFloat((tds[7].querySelector("input") || {}).value || NaN);
  const actualDia = parseFloat((tds[8].querySelector("input") || {}).value || NaN);
  const actualYZ = parseFloat((tds[9].querySelector("input") || {}).value || NaN);
  const xaxis = parseFloat((tds[4].querySelector("input") || {}).value || NaN);

  const offsetCell = tds[10];
  const resultCell = tds[11];

  // compute offset: specYZ + specDia/2 (if both present)
  let offset = null;
  if (!isNaN(specYZ) && !isNaN(specDia)) {
    offset = specYZ + (specDia / 2);
    offsetCell.textContent = round(offset, 2);
  } else {
    offsetCell.textContent = "";
  }

  // tolerance based on FSM length: pick parsed fsm length or value from fsmSpec input
  const fsmSpecVal = parseFloat(document.getElementById("fsmSpec")?.value || parsedHeader.fsmLength || NaN);
  let tol = 1;
  if (!isNaN(xaxis) && !isNaN(fsmSpecVal)) {
    if (xaxis <= 200 || xaxis >= (fsmSpecVal - 200)) tol = 2;
  }

  // Dia tolerance calculation (keep old logic)
  let diaOk = true;
  if (!isNaN(specDia) && !isNaN(actualDia)) {
    if (specDia <= 10.7) diaOk = (actualDia >= specDia - 0 && actualDia <= specDia + 0.4);
    else if (specDia >= 11.7) diaOk = (actualDia >= specDia - 0 && actualDia <= specDia + 0.5);
    else diaOk = (actualDia >= specDia - 0 && actualDia <= specDia + 0.5);
  } else {
    diaOk = false;
  }

  // offset/yz tolerance
  let offsetOk = true;
  if (!isNaN(offset) && !isNaN(actualYZ)) {
    offsetOk = Math.abs(actualYZ - offset) <= tol;
  } else {
    offsetOk = false;
  }

  const allOk = diaOk && offsetOk;

  // Reset styles before applying
  [tds[7], tds[8], tds[9], tds[10], tds[11]].forEach(cell => {
    cell.classList.remove("nok-cell");
    cell.classList.remove("ok-cell");
  });

  if (allOk) {
    resultCell.textContent = "OK";
    resultCell.classList.add("ok-cell");
  } else {
    resultCell.textContent = "NOK";
    resultCell.classList.add("nok-cell");
    if (!diaOk) tds[8].classList.add("nok-cell");
    if (!offsetOk) tds[9].classList.add("nok-cell");
    if (!isNaN(offset) && !isNaN(actualYZ) && Math.abs(actualYZ - offset) > tol) tds[10].classList.add("nok-cell");
  }
}

/* ---------------------- Other sections (First/Last holes) ---------------------- */
function buildFirstHolesSection() {
  const blk = document.createElement("div"); blk.className = "form-block";
  blk.innerHTML = `<strong>First holes of FSM - from front end</strong>`;
  const table = document.createElement("table"); table.className = "table";
  table.innerHTML = `<thead><tr><th>Spec</th><th>Actual (Value from Hole edge)</th><th>Actual (Dia)</th><th>Offset</th><th>Result</th></tr></thead>`;
  const tb = document.createElement("tbody");

  // best-effort parse a few numbers; if not found produce blanks
  const specs = extractFirstHoleSpecs(lastExtractedText) || ["", "", "", "", ""];
  for (let i = 0; i < 5; i++) {
    const tr = document.createElement("tr");
    const specTd = document.createElement("td"); specTd.textContent = specs[i] || "";
    const valEdge = makeInput({ className: "input-small" });
    const actDia = makeInput({ className: "input-small" });
    const offTd = document.createElement("td");
    const resTd = document.createElement("td");
    [valEdge, actDia].forEach(inp => {
      inp.addEventListener("input", () => {
        const spec = parseFloat(specTd.textContent || NaN);
        const dia = parseFloat(actDia.value || NaN);
        let offset = (isNaN(spec) || isNaN(dia)) ? "" : round(spec + (dia / 2), 2);
        offTd.textContent = offset;
        const val = parseFloat(valEdge.value || NaN);
        if (!isNaN(val) && !isNaN(spec)) {
          if (Math.abs(val - spec) > 1) {
            valEdge.classList.add("nok-cell");
            resTd.textContent = "NOK"; resTd.className = "nok-cell";
          } else {
            valEdge.classList.remove("nok-cell");
            resTd.textContent = "OK"; resTd.className = "ok-cell";
          }
        } else {
          resTd.textContent = "";
          resTd.className = "";
        }
      });
    });
    tr.appendChild(specTd);
    tr.appendChild(tdWrap(valEdge));
    tr.appendChild(tdWrap(actDia));
    tr.appendChild(offTd);
    tr.appendChild(resTd);
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
  const smalls = matches.map(Number).filter(n => n > 0 && n < 5000);
  return smalls.slice(0, 5).map(n => n.toString());
}
function buildLastHolesSection() {
  const blk = document.createElement("div"); blk.className = "form-block";
  blk.innerHTML = `<strong>Last holes of FSM - from rear end</strong>`;
  const table = document.createElement("table"); table.className = "table";
  table.innerHTML = `<thead><tr><th>Spec</th><th>Actual (Value from Hole edge)</th><th>Actual (Dia)</th><th>Offset</th><th>Result</th></tr></thead>`;
  const tb = document.createElement("tbody");
  const specs = extractFirstHoleSpecs(lastExtractedText) || ["", "", "", "", ""];
  for (let i = 0; i < 5; i++) {
    const tr = document.createElement("tr");
    const specTd = document.createElement("td"); specTd.textContent = specs[i] || "";
    const valEdge = makeInput({ className: "input-small" });
    const actDia = makeInput({ className: "input-small" });
    const offTd = document.createElement("td");
    const resTd = document.createElement("td");
    [valEdge, actDia].forEach(inp => {
      inp.addEventListener("input", () => {
        const spec = parseFloat(specTd.textContent || NaN);
        const dia = parseFloat(actDia.value || NaN);
        let offset = (isNaN(spec) || isNaN(dia)) ? "" : round(spec + (dia / 2), 2);
        offTd.textContent = offset;
        const val = parseFloat(valEdge.value || NaN);
        if (!isNaN(val) && !isNaN(spec)) {
          if (Math.abs(val - spec) > 1) {
            valEdge.classList.add("nok-cell");
            resTd.textContent = "NOK"; resTd.className = "nok-cell";
          } else {
            valEdge.classList.remove("nok-cell");
            resTd.textContent = "OK"; resTd.className = "ok-cell";
          }
        } else {
          resTd.textContent = "";
          resTd.className = "";
        }
      });
    });
    tr.appendChild(specTd);
    tr.appendChild(tdWrap(valEdge));
    tr.appendChild(tdWrap(actDia));
    tr.appendChild(offTd);
    tr.appendChild(resTd);
    tb.appendChild(tr);
  }
  table.appendChild(tb);
  blk.appendChild(table);
  formArea.appendChild(blk);
}

/* ---------------------- Root / Flange table (reads rootAct value from header input) ---------------------- */
function buildRootAndFlangeSection() {
  const blk = document.createElement("div"); blk.className = "form-block";
  blk.innerHTML = `<strong>Root / Flange / Web (reference)</strong>`;
  const tbl = document.createElement("table"); tbl.className = "table";
  tbl.innerHTML = `<thead><tr><th>Root Width Spec</th><th>Root Width Act</th><th>Top Flange</th><th>Web PWM</th><th>Bottom Flange</th></tr></thead>`;
  const tb = document.createElement("tbody");
  const tr = document.createElement("tr");

  const rootSpecTd = document.createElement("td");
  rootSpecTd.textContent = document.getElementById("rootSpec")?.value || "";

  const rootActInput = makeInput({ id: "rootAct_flange", className: "input-small", placeholder: "Act (mm)", value: document.getElementById("rootAct")?.value || "" });
  // sync rootAct header input -> flange input
  const headerRootAct = document.getElementById("rootAct");
  if (headerRootAct) {
    headerRootAct.addEventListener("input", () => {
      rootActInput.value = headerRootAct.value;
      // trigger validation class change
      rootActInput.dispatchEvent(new Event('input'));
    });
  }
  // validate difference > 1 -> NOK class
  rootActInput.addEventListener("input", () => {
    const spec = parseFloat(rootSpecTd.textContent || NaN);
    const act = parseFloat(rootActInput.value || NaN);
    if (!isNaN(spec) && !isNaN(act)) {
      if (Math.abs(act - spec) > 1) {
        rootActInput.classList.add("nok-cell");
      } else {
        rootActInput.classList.remove("nok-cell");
      }
    }
  });

  const topFl = document.createElement("td"); topFl.textContent = "Top Flange: PWM (info)";
  const webPW = document.createElement("td"); webPW.textContent = "Web PWM (info)";
  const bottomFl = document.createElement("td"); bottomFl.textContent = "Bottom Flange (info)";

  tr.appendChild(rootSpecTd);
  tr.appendChild(tdWrap(rootActInput));
  tr.appendChild(topFl);
  tr.appendChild(webPW);
  tr.appendChild(bottomFl);

  tb.appendChild(tr); tbl.appendChild(tb); blk.appendChild(tbl); formArea.appendChild(blk);
}

/* ---------------------- Part no location, binary checks, remarks & sign ---------------------- */
function buildPartNoLocation() {
  const blk = document.createElement("div"); blk.className = "form-block";
  blk.innerHTML = `<strong>Part no location (Spec / Act)</strong>`;
  const table = document.createElement("table"); table.className = "table";
  table.innerHTML = `<thead><tr><th>Spec (mm)</th><th>Act (mm)</th><th>Result</th></tr></thead>`;
  const tb = document.createElement("tbody");
  const tr = document.createElement("tr");
  const specTd = document.createElement("td"); const specInp = makeInput({ className: "input-small" }); specTd.appendChild(specInp);
  const actTd = document.createElement("td"); const actInp = makeInput({ className: "input-small" }); actTd.appendChild(actInp);
  const resTd = document.createElement("td");
  actInp.addEventListener("input", () => {
    const s = parseFloat(specInp.value || NaN);
    const a = parseFloat(actInp.value || NaN);
    if (!isNaN(s) && !isNaN(a)) {
      if (Math.abs(a - s) > 5) { actInp.classList.add("nok-cell"); resTd.textContent = "NOK"; resTd.className = "nok-cell"; }
      else { actInp.classList.remove("nok-cell"); resTd.textContent = "OK"; resTd.className = "ok-cell"; }
    } else { resTd.textContent = ""; resTd.className = ""; }
  });
  tr.appendChild(specTd); tr.appendChild(actTd); tr.appendChild(resTd); tb.appendChild(tr);
  table.appendChild(tb); blk.appendChild(table); formArea.appendChild(blk);
}

function buildBinaryChecks() {
  const blk = document.createElement("div"); blk.className = "form-block";
  blk.innerHTML = `<strong>Visual / Binary Checks</strong>`;
  const grid = document.createElement("div"); grid.style.display = "grid"; grid.style.gridTemplateColumns = "repeat(4,1fr)"; grid.style.gap = "8px";
  const labels = ["Punch Break", "Radius crack", "Length Variation", "Holes Burr", "Slug mark", "Line mark", "Part No. Legibility", "Pit Mark", "Machine Error (YES/NO)"];
  labels.forEach(lbl => {
    const cell = document.createElement("div"); cell.className = "small";
    const labelEl = document.createElement("div"); labelEl.textContent = lbl;
    const okBtn = document.createElement("button"); okBtn.textContent = "OK/YES"; okBtn.className = "input-small";
    const nokBtn = document.createElement("button"); nokBtn.textContent = "NOK/NO"; nokBtn.className = "input-small";
    okBtn.addEventListener("click", () => {
      okBtn.classList.add("ok-cell-binary"); nokBtn.classList.remove("nok-cell-binary");
      nokBtn.classList.remove("ok-cell-binary"); checkMandatoryBeforeExport();
    });
    nokBtn.addEventListener("click", () => {
      nokBtn.classList.add("nok-cell-binary"); okBtn.classList.remove("ok-cell-binary"); checkMandatoryBeforeExport();
    });
    cell.appendChild(labelEl); cell.appendChild(okBtn); cell.appendChild(nokBtn);
    grid.appendChild(cell);
  });
  blk.appendChild(grid); formArea.appendChild(blk);
}

function buildRemarksAndSign() {
  const blk = document.createElement("div"); blk.className = "form-block";
  blk.innerHTML = `<strong>Remarks / Details of issue</strong>`;
  const ta = document.createElement("textarea"); ta.id = "remarks"; ta.placeholder = "Enter remarks (optional)";
  ta.style.width = "100%"; ta.style.minHeight = "60px";
  blk.appendChild(ta);

  const signRow = document.createElement("div"); signRow.className = "form-row";
  const col1 = document.createElement("div"); col1.className = "col";
  col1.innerHTML = `<div class="small">Prodn Incharge (Shift Executive) - mandatory</div>`;
  const prodIn = makeInput({ id: "prodIncharge", className: "input-small", placeholder: "Enter shift executive name" });
  prodIn.addEventListener("input", checkMandatoryBeforeExport);
  col1.appendChild(prodIn); signRow.appendChild(col1);
  blk.appendChild(signRow);
  formArea.appendChild(blk);
}

/* ---------------------- Utility wrappers ---------------------- */
function tdWrap(inner) {
  const td = document.createElement("td");
  if (inner instanceof HTMLElement) td.appendChild(inner);
  else td.textContent = inner;
  return td;
}

/* ---------------------- Attach validation listeners & mandatory checks ---------------------- */
function attachValidationListeners() {
  // header & mandatory inputs
  const idsToWatch = ["fsmSerial", "holesAct", "matrixUsed", "prodIncharge"];
  idsToWatch.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("input", checkMandatoryBeforeExport);
  });
  // inspector1 or inspector2
  document.getElementById("inspector1")?.addEventListener("input", checkMandatoryBeforeExport);
  document.getElementById("inspector2")?.addEventListener("input", checkMandatoryBeforeExport);
  // KB/PC act
  document.getElementById("kbAct")?.addEventListener("input", checkMandatoryBeforeExport);
  document.getElementById("pcAct")?.addEventListener("input", checkMandatoryBeforeExport);
  // rootAct / fsmAct
  document.getElementById("rootAct")?.addEventListener("input", checkMandatoryBeforeExport);
  document.getElementById("fsmAct")?.addEventListener("input", checkMandatoryBeforeExport);
}

function checkMandatoryBeforeExport() {
  const fsmSerial = document.getElementById("fsmSerial")?.value || "";
  const instr1 = document.getElementById("inspector1")?.value || "";
  const instr2 = document.getElementById("inspector2")?.value || "";
  const holesAct = document.getElementById("holesAct")?.value || "";
  const matrixUsed = document.getElementById("matrixUsed")?.value || "";
  const prod = document.getElementById("prodIncharge")?.value || "";
  let missing = [];
  if (!fsmSerial.trim()) missing.push("FSM Serial Number");
  if (!instr1.trim() && !instr2.trim()) missing.push("Inspectors");
  if (!holesAct.trim()) missing.push("Total Holes Count (Act)");
  if (!matrixUsed.trim()) missing.push("Matrix used");
  if (!prod.trim()) missing.push("Prodn Incharge");
  if (missing.length) {
    showMessage("Please fill mandatory fields: " + missing.join(", "), "warn");
    exportPdfBtn.disabled = true;
    return false;
  }
  // KB & PC Act must be present
  const kbAct = document.getElementById("kbAct")?.value || "";
  const pcAct = document.getElementById("pcAct")?.value || "";
  if (!kbAct.trim() || !pcAct.trim()) {
    showMessage("Please fill KB & PC Act values.", "warn");
    exportPdfBtn.disabled = true;
    return false;
  }

  exportPdfBtn.disabled = false;
  clearMessage();
  return true;
}

/* ---------------------- PDF export ---------------------- */
exportPdfBtn.addEventListener("click", async () => {
  if (!checkMandatoryBeforeExport()) return;
  const hasNok = document.querySelectorAll(".nok-cell, .nok-cell-binary").length > 0;
  if (hasNok) {
    const proceed = confirm("Some fields are out of spec (NOK). Please verify. Proceed to export anyway?");
    if (!proceed) return;
  }
  exportPdfBtn.disabled = true;
  exportPdfBtn.textContent = "Generating PDF...";
  try {
    await generatePdfFromForm();
  } catch (err) {
    console.error(err);
    alert("PDF export failed: " + (err.message || err));
  } finally {
    exportPdfBtn.disabled = false;
    exportPdfBtn.textContent = "Export to PDF";
  }
});

async function generatePdfFromForm() {
  const clone = document.createElement("div");
  clone.style.width = "900px";
  clone.style.padding = "12px";
  const title = document.createElement("div");
  title.innerHTML = document.querySelector("header")?.innerHTML || "";
  clone.appendChild(title);

  // clone form area and replace inputs with spans
  const fa = formArea.cloneNode(true);
  fa.querySelectorAll("input, textarea, button").forEach(n => {
    if (n.tagName === "INPUT") {
      const span = document.createElement("span");
      span.textContent = n.value;
      n.parentNode && n.parentNode.replaceChild(span, n);
    } else if (n.tagName === "TEXTAREA") {
      const span = document.createElement("div");
      span.textContent = n.value;
      span.style.whiteSpace = "pre-wrap";
      n.parentNode && n.parentNode.replaceChild(span, n);
    } else if (n.tagName === "BUTTON") {
      n.remove();
    }
  });

  clone.appendChild(fa);
  clone.style.position = "fixed";
  clone.style.left = "-2000px";
  document.body.appendChild(clone);

  try {
    const canvas = await html2canvas(clone, { scale: 2, useCORS: true });
    const imgData = canvas.toDataURL("image/png");
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF("p", "mm", "a4");
    const pageWidth = pdf.internal.pageSize.getWidth();
    const imgProps = pdf.getImageProperties(imgData);
    const pdfWidth = pageWidth - 20;
    const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
    pdf.addImage(imgData, "PNG", 10, 10, pdfWidth, pdfHeight);
    const now = new Date();
    pdf.setFontSize(9);
    pdf.text(`Generated: ${now.toLocaleString()}`, 14, pdf.internal.pageSize.getHeight() - 10);

    const part = parsedHeader.partNumber || "PART";
    const rev = parsedHeader.revision || "X";
    const hand = parsedHeader.hand || "H";
    const shortDate = now.toISOString().slice(0, 10);
    const filename = `${part}_${rev}_${hand}_Inspection_${shortDate}.pdf`;
    pdf.save(filename);
    showMessage("PDF exported: " + filename, "info");
  } finally {
    document.body.removeChild(clone);
  }
}

/* ---------------------- Build form main entry (safe rebind) ---------------------- */
/* NOTE: We allow building the form once user has loaded the PDF (parsing preferred but not mandatory).
   If you want to enforce parsing before building, change the first if-check to: if (!lastExtractedText) { showMessage(...); return; }
*/
generateFormBtn.addEventListener("click", () => {
  // allow building even if parsedTableRows empty; prefer to show blanks than to block user
  if (!pdfFile && !lastExtractedText) {
    showMessage("Please load a PDF first.", "warn");
    return;
  }

  generateFormBtn.disabled = true;
  generateFormBtn.textContent = "Building...";

  try {
    formArea.innerHTML = "";
    clearMessage();

    // sections (order preserved)
    buildHeaderInputs();           // header & mandatory fields
    buildMainTable();              // main table - up to 45 rows prefilled for columns 1..7 readonly
    buildFirstHolesSection();      // first holes (front)
    buildLastHolesSection();       // last holes (rear)
    buildRootAndFlangeSection();   // root/flange (syncs with header rootAct)
    buildPartNoLocation();
    buildBinaryChecks();
    buildRemarksAndSign();

    attachValidationListeners();
    exportPdfBtn.disabled = true; // must satisfy mandatory to enable
    showMessage("Editable form created. Fill Actual values and mandatory fields. Use 'Export to PDF' to generate the report.", "info");

  } catch (err) {
    console.error("Build form failed:", err);
    showMessage("Build failed: " + (err.message || err), "warn");
  } finally {
    generateFormBtn.disabled = false;
    generateFormBtn.textContent = "Create Editable Form";
  }
});
