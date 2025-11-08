/* ===== SOENEN INSPECTION PHASE 2 SCRIPT ===== */

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.9.179/pdf.worker.min.js";

const fileInput = document.getElementById("fileInput");
const extractBtn = document.getElementById("extractBtn");
const pdfViewer = document.getElementById("pdfViewer");
const headerOut = document.getElementById("headerOut");
const generateFormBtn = document.getElementById("generateFormBtn");
const formArea = document.getElementById("formArea");
const exportPdfBtn = document.getElementById("exportPdfBtn");
const warnings = document.getElementById("warnings");

let pdfFile = null;
let lastExtractedText = "";
let parsedHeader = {};
let parsedTableRows = [];

/* ---------- Load PDF and preview ---------- */
fileInput.addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  pdfFile = f;
  const buffer = await pdfFile.arrayBuffer();
  await renderPDF(new Uint8Array(buffer));
  headerOut.textContent = "PDF loaded — click 'Load & Parse PDF'.";
  generateFormBtn.disabled = true;
  exportPdfBtn.disabled = true;
});

/* ---------- Parse PDF ---------- */
extractBtn.addEventListener("click", async () => {
  if (!pdfFile) {
    alert("Please select a PDF first.");
    return;
  }
  extractBtn.disabled = true;
  extractBtn.textContent = "Parsing...";
  try {
    const buf = await pdfFile.arrayBuffer();
    const text = await extractTextFromPDF(new Uint8Array(buf));
    lastExtractedText = text;
    parsedHeader = parseHeader(text);
    parsedTableRows = detectTableLines(text);
    headerOut.textContent = JSON.stringify(parsedHeader, null, 2);
    showMessage("Parsed successfully. Click 'Create Editable Form'.", "info");
    generateFormBtn.disabled = false;
  } catch (err) {
    console.error(err);
    alert("Error while parsing PDF: " + err.message);
  } finally {
    extractBtn.disabled = false;
    extractBtn.textContent = "Load & Parse PDF";
  }
});

/* ---------- Render first page preview ---------- */
async function renderPDF(bytes) {
  pdfViewer.innerHTML = "";
  try {
    const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
    const page = await pdf.getPage(1);
    const vp = page.getViewport({ scale: 1.1 });
    const canvas = document.createElement("canvas");
    canvas.width = vp.width;
    canvas.height = vp.height;
    const ctx = canvas.getContext("2d");
    await page.render({ canvasContext: ctx, viewport: vp }).promise;
    pdfViewer.appendChild(canvas);
  } catch {
    pdfViewer.textContent = "Unable to render preview.";
  }
}

/* ---------- Extract text ---------- */
async function extractTextFromPDF(bytes) {
  const doc = await pdfjsLib.getDocument({ data: bytes }).promise;
  let text = "";
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map((it) => it.str).join(" ") + "\n";
  }
  return text;
}

/* ---------- Parse header ---------- */
function parseHeader(t) {
  const h = {};
  let m;

  m = t.match(/PART\s*NUMBER\s*[:\-]?\s*([A-Z0-9]+)/i);
  if (m) h.partNumber = m[1];

  m = t.match(/ROOT\s*WIDTH.*?([0-9.]+)\s*mm/i);
  if (m) h.rootWidth = m[1];

  m = t.match(/FSM\s*LENGTH.*?([0-9.]+)\s*mm/i);
  if (m) h.fsmLength = m[1];

  m = t.match(/KB.*?Spec.*?(\d+)\s*\/\s*(\d+)/i);
  if (m) {
    h.kbSpec = m[1];
    h.pcSpec = m[2];
  }

  m = t.match(/FORMAT\s*NO.*?([A-Z0-9\/\s\-]+)/i);
  if (m) h.formatNo = m[1].trim();

  return h;
}

/* ---------- Detect table rows ---------- */
function detectTableLines(fullText) {
  const lines = fullText
    .split(/\n/)
    .map((s) => s.trim())
    .filter((s) => s && /^\d+\s/.test(s));
  return lines.map((l) => l.split(/\s+/));
}

/* ---------- Input helper ---------- */
function makeInput(attrs = {}) {
  const inp = document.createElement("input");
  inp.type = "text";
  Object.assign(inp, attrs);
  return inp;
}

/* ---------- Build Form ---------- */
generateFormBtn.addEventListener("click", () => {
  formArea.innerHTML = "";
  buildHeaderInputs();
  buildMainTable();
  exportPdfBtn.disabled = false;
  showMessage("Editable form created. Fill values and export.", "info");
});

/* ---------- Header / Mandatory fields ---------- */
function buildHeaderInputs() {
  const blk = document.createElement("div");
  blk.className = "form-block";
  blk.innerHTML = `<strong>Header / Mandatory fields</strong>`;

  // FSM Serial + Inspectors
  const row = document.createElement("div");
  row.className = "form-row";

  const c1 = document.createElement("div");
  c1.className = "col";
  c1.innerHTML = `<div class="small">FSM Serial Number:</div>`;
  const fsmSerial = makeInput({
    id: "fsmSerial",
    className: "input-small",
    placeholder: "Enter FSM Serial",
  });
  c1.appendChild(fsmSerial);

  const c2 = document.createElement("div");
  c2.className = "col";
  c2.innerHTML = `<div class="small">Inspectors:</div>`;
  const inspector1 = makeInput({
    id: "inspector1",
    className: "input-small",
    placeholder: "Inspector 1",
  });
  const inspector2 = makeInput({
    id: "inspector2",
    className: "input-small",
    placeholder: "Inspector 2",
  });
  c2.append(inspector1, inspector2);
  row.append(c1, c2);
  blk.appendChild(row);

  // KB/PC + Root/FSM
  const row2 = document.createElement("div");
  row2.className = "form-row";
  const kbcol = document.createElement("div");
  kbcol.className = "col";
  kbcol.innerHTML = `<div class="small">KB & PC Code (Spec / Act)</div>`;

  const kbSpec = makeInput({
    className: "input-small",
    id: "kbSpec",
    value: parsedHeader.kbSpec || "",
    readOnly: true,
  });
  const kbAct = makeInput({ className: "input-small", id: "kbAct" });
  const pcSpec = makeInput({
    className: "input-small",
    id: "pcSpec",
    value: parsedHeader.pcSpec || "",
    readOnly: true,
  });
  const pcAct = makeInput({ className: "input-small", id: "pcAct" });
  kbcol.append(
    kbSpec,
    document.createTextNode(" / "),
    kbAct,
    document.createTextNode("    "),
    pcSpec,
    document.createTextNode(" / "),
    pcAct
  );

  const rwcol = document.createElement("div");
  rwcol.className = "col";
  rwcol.innerHTML = `<div class="small">ROOT WIDTH OF FSM (Spec / Act) mm | FSM LENGTH (Spec / Act) mm</div>`;
  const rootSpec = makeInput({
    id: "rootSpec",
    className: "input-small",
    value: parsedHeader.rootWidth || "",
    readOnly: true,
  });
  const rootAct = makeInput({ id: "rootAct", className: "input-small" });
  const fsmSpec = makeInput({
    id: "fsmSpec",
    className: "input-small",
    value: parsedHeader.fsmLength || "",
    readOnly: true,
  });
  const fsmAct = makeInput({ id: "fsmAct", className: "input-small" });
  rwcol.append(rootSpec, rootAct, fsmSpec, fsmAct);

  row2.append(kbcol, rwcol);
  blk.appendChild(row2);

  formArea.appendChild(blk);
}

/* ---------- Main inspection table ---------- */
function buildMainTable() {
  const blk = document.createElement("div");
  blk.className = "form-block";
  blk.innerHTML = `<strong>Main Inspection Table</strong>`;

  const table = document.createElement("table");
  table.className = "table";

  const thead = document.createElement("thead");
  thead.innerHTML = `
    <tr>
      <th>Sl No</th><th>Press</th><th>Sel ID</th><th>Ref</th>
      <th>X-axis</th><th>Spec (Y or Z)</th><th>Spec Dia</th>
      <th>Value from Hole edge (Act)</th><th>Actual Dia</th>
      <th>Actual Y or Z</th><th>Offset (Actual - Spec)</th><th>Result</th>
    </tr>`;
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  const rowsToBuild = parsedTableRows.length ? parsedTableRows.length : 20;

  for (let i = 0; i < rowsToBuild; i++) {
    const tr = document.createElement("tr");
    const tokens = parsedTableRows[i] || [];

    const makePrefilled = (v) => {
      const inp = makeInput({ className: "input-small", value: v || "", readOnly: true });
      inp.style.background = "#f0f0f0";
      return inp;
    };

    const sl = makePrefilled(i + 1);
    const press = makePrefilled(tokens[1] || "");
    const sel = makePrefilled(tokens[2] || "");
    const ref = makePrefilled(tokens[3] || "");
    const xaxis = makePrefilled(tokens[4] || "");
    const specYZ = makePrefilled(tokens[5] || "");
    const specDia = makePrefilled(tokens[6] || "");

    const valEdge = makeInput({ className: "input-small" });
    const actDia = makeInput({ className: "input-small" });
    const actYZ = makeInput({ className: "input-small" });
    const offset = document.createElement("td");
    const result = document.createElement("td");

    [valEdge, actDia, actYZ].forEach((inp) =>
      inp.addEventListener("input", () => recalcRow(tr))
    );

    [sl, press, sel, ref, xaxis, specYZ, specDia, valEdge, actDia, actYZ].forEach((v) =>
      tr.appendChild(tdWrap(v))
    );
    tr.append(offset, result);
    tbody.appendChild(tr);
  }

  table.appendChild(tbody);
  blk.appendChild(table);
  formArea.appendChild(blk);
}

function tdWrap(el) {
  const td = document.createElement("td");
  td.appendChild(el);
  return td;
}

function recalcRow(tr) {
  const tds = tr.querySelectorAll("td");
  const specYZ = parseFloat((tds[5].querySelector("input") || {}).value || NaN);
  const specDia = parseFloat((tds[6].querySelector("input") || {}).value || NaN);
  const actDia = parseFloat((tds[8].querySelector("input") || {}).value || NaN);
  const actYZ = parseFloat((tds[9].querySelector("input") || {}).value || NaN);

  const offsetCell = tds[10];
  const resultCell = tds[11];
  let offset = NaN;
  if (!isNaN(actYZ) && !isNaN(specYZ)) offset = actYZ - specYZ;
  offsetCell.textContent = isNaN(offset) ? "" : offset.toFixed(2);

  let ok = true;
  if (!isNaN(specDia) && !isNaN(actDia)) {
    ok = actDia >= specDia - 0.2 && actDia <= specDia + 0.5;
  }

  if (ok) {
    resultCell.textContent = "OK";
    resultCell.className = "ok-cell";
  } else {
    resultCell.textContent = "NOK";
    resultCell.className = "nok-cell";
  }
}

/* ---------- Export to PDF ---------- */
exportPdfBtn.addEventListener("click", async () => {
  exportPdfBtn.textContent = "Generating...";
  const canvas = await html2canvas(formArea, { scale: 2 });
  const img = canvas.toDataURL("image/png");
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF("p", "mm", "a4");
  const width = pdf.internal.pageSize.getWidth() - 20;
  const height = (canvas.height * width) / canvas.width;
  pdf.addImage(img, "PNG", 10, 10, width, height);
  pdf.save("inspection_report.pdf");
  exportPdfBtn.textContent = "Export to PDF";
});

/* ---------- Helper Messages ---------- */
function showMessage(msg, type) {
  warnings.textContent = msg;
  warnings.style.padding = "6px";
  warnings.style.borderRadius = "6px";
  warnings.style.marginTop = "6px";
  warnings.style.background =
    type === "warn" ? "#fff6cc" : "#e8f4ff";
  warnings.style.color = type === "warn" ? "#ff8c00" : "#005fcc";
}
