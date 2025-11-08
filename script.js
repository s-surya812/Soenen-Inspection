/* ===========================================
   SOENEN INSPECTION PHASE 2 – COMPLETE SCRIPT
   =========================================== */

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
let parsedHeader = {};
let parsedTableRows = [];

/* ---------- HEADER UPDATE ---------- */
function updateHeaderFields(parsedHeader) {
  const partNumber = parsedHeader.partNumber || "—";
  const revision = parsedHeader.revision || "—";
  const hand = parsedHeader.hand || "—";
  const fsmLength = parsedHeader.fsmLength || "—";
  hdrPart.textContent = `PART NUMBER / LEVEL / HAND : ${partNumber} / ${revision} / ${hand}`;
  hdrFsmSpec.textContent = `FSM LENGTH : ${fsmLength} mm`;
}

/* ---------- PDF LOAD ---------- */
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

/* ---------- PARSE PDF ---------- */
extractBtn.addEventListener("click", async () => {
  if (!pdfFile) return alert("Please select a PDF first.");
  extractBtn.disabled = true;
  extractBtn.textContent = "Parsing...";
  try {
    const buf = await pdfFile.arrayBuffer();
    const text = await extractTextFromPDF(new Uint8Array(buf));
    lastExtractedText = text;
    parsedHeader = parseHeader(text);
    headerOut.textContent = JSON.stringify(parsedHeader, null, 2);
    updateHeaderFields(parsedHeader);
    parsedTableRows = detectTableLines(text);
    generateFormBtn.disabled = false;
    showMessage("Parsed PDF successfully. Click 'Create Editable Form'.", "info");
  } catch (err) {
    console.error(err);
    alert("Error parsing PDF: " + err.message);
  } finally {
    extractBtn.disabled = false;
    extractBtn.textContent = "Load & Parse PDF";
  }
});

/* ---------- PDF PREVIEW ---------- */
async function renderPDF(bytes) {
  pdfViewer.innerHTML = "";
  try {
    const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
    const page = await pdf.getPage(1);
    const vp = page.getViewport({ scale: 1.1 });
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    canvas.width = vp.width;
    canvas.height = vp.height;
    await page.render({ canvasContext: ctx, viewport: vp }).promise;
    pdfViewer.appendChild(canvas);
  } catch {
    pdfViewer.textContent = "Unable to render preview.";
  }
}

/* ---------- TEXT EXTRACTION ---------- */
async function extractTextFromPDF(bytes) {
  const doc = await pdfjsLib.getDocument({ data: bytes }).promise;
  let out = "";
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    out += content.items.map((it) => it.str).join(" ") + "\n---PAGE_BREAK---\n";
  }
  return out;
}

/* ---------- HEADER PARSING ---------- */
function parseHeader(t) {
  const h = {
    partNumber: null,
    revision: null,
    hand: null,
    date: null,
    formatNo: null,
    rootWidth: null,
    fsmLength: null,
    kbSpec: null,
    pcSpec: null,
  };
  let m;
  m = t.match(/PART\s*NUMBER\s*\/\s*LEVEL\s*\/\s*HAND\s*[:\s\-]*([A-Z0-9_\-]+)\s*\/\s*([A-Z0-9_\-]+)\s*\/\s*([A-Z]+)/i);
  if (m) [h.partNumber, h.revision, h.hand] = [m[1], m[2], m[3]];
  m = t.match(/ROOT\s*WIDTH\s*OF\s*FSM\s*[:\-\s]*Spec[-\s]*([0-9.]+)/i);
  if (m) h.rootWidth = parseFloat(m[1]);
  m = t.match(/FSM\s*LENGTH\s*[:\-\s]*Spec[-\s]*([0-9.]+)/i);
  if (m) h.fsmLength = parseFloat(m[1]);
  m = t.match(/KB\s*&\s*PC\s*Code\s*[:\-]*\s*Spec[-\s]*([0-9]+)\s*\/\s*([0-9]+)/i);
  if (m) [h.kbSpec, h.pcSpec] = [m[1], m[2]];
  m = t.match(/FORMAT\s*NO\.?\s*[:\-]*\s*([A-Z0-9\/\s\-\_]+)/i);
  if (m) h.formatNo = m[1].trim();
  return h;
}

/* ---------- TABLE LINE DETECTION ---------- */
function detectTableLines(fullText) {
  if (!fullText) return [];
  const lines = fullText
    .split(/\n|---PAGE_BREAK---/)
    .map((s) => s.replace(/\s{2,}/g, " ").trim())
    .filter(Boolean);
  const rows = [];
  for (const line of lines) {
    if (/^\s*\d+(\.|-)?\s+\w+/.test(line)) rows.push(line);
  }
  return rows.map((r) => r.trim().split(/\s+/));
}

/* ---------- INPUT CREATOR ---------- */
function makeInput(attrs = {}) {
  const inp = document.createElement("input");
  inp.type = attrs.type || "text";
  inp.value = attrs.value || "";
  inp.className = attrs.className || "input-small";
  if (attrs.placeholder) inp.placeholder = attrs.placeholder;
  if (attrs.step) inp.step = attrs.step;
  if (attrs.readOnly) inp.readOnly = true;
  if (attrs.size) inp.size = attrs.size;
  if (attrs.id) inp.id = attrs.id;
  if (attrs.oninput) inp.addEventListener("input", attrs.oninput);
  return inp;
}

/* ---------- HEADER BLOCK ---------- */
function buildHeaderInputs() {
  const blk = document.createElement("div");
  blk.className = "form-block";
  blk.innerHTML = `<strong>Header / Mandatory fields</strong>`;

  // Row 1: FSM Serial + Inspectors
  const row1 = document.createElement("div");
  row1.className = "form-row";
  const c1 = document.createElement("div");
  c1.className = "col";
  c1.innerHTML = `<div class="small">FSM Serial Number:</div>`;
  c1.appendChild(makeInput({ id: "fsmSerial", placeholder: "Enter FSM Serial No" }));

  const c2 = document.createElement("div");
  c2.className = "col";
  c2.innerHTML = `<div class="small">INSPECTORS:</div>`;
  c2.appendChild(makeInput({ id: "inspector1", placeholder: "Inspector 1" }));
  c2.appendChild(document.createTextNode(" "));
  c2.appendChild(makeInput({ id: "inspector2", placeholder: "Inspector 2" }));
  row1.append(c1, c2);
  blk.appendChild(row1);

  // Row 2: Holes + Matrix
  const row2 = document.createElement("div");
  row2.className = "form-row";
  const c3 = document.createElement("div");
  c3.className = "col";
  c3.innerHTML = `<div class="small">TOTAL HOLES COUNT (Spec / Act):</div>`;
  c3.append(
    makeInput({ id: "holesSpec", placeholder: "Spec" }),
    document.createTextNode(" "),
    makeInput({ id: "holesAct", placeholder: "Act" })
  );
  const c4 = document.createElement("div");
  c4.className = "col";
  c4.innerHTML = `<div class="small">Matrix Used:</div>`;
  c4.appendChild(makeInput({ id: "matrixUsed", placeholder: "Matrix" }));
  row2.append(c3, c4);
  blk.appendChild(row2);

  // Row 3: KB/PC
  const row3 = document.createElement("div");
  row3.className = "form-row";
  const kbcol = document.createElement("div");
  kbcol.className = "col";
  kbcol.innerHTML = `<div class="small">KB & PC Code (Spec / Act)</div>`;
  const kbSpec = makeInput({ id: "kbSpec", readOnly: true, value: parsedHeader.kbSpec || "", size: 3 });
  const kbAct = makeInput({ id: "kbAct", placeholder: "Act", size: 3 });
  const pcSpec = makeInput({ id: "pcSpec", readOnly: true, value: parsedHeader.pcSpec || "", size: 3 });
  const pcAct = makeInput({ id: "pcAct", placeholder: "Act", size: 3 });
  kbcol.append(kbSpec, document.createTextNode(" / "), kbAct, document.createTextNode("  "), pcSpec, document.createTextNode(" / "), pcAct);
  row3.append(kbcol);
  blk.appendChild(row3);

  // Row 4: Root width + FSM length
  const row4 = document.createElement("div");
  row4.className = "form-row";
  const rwcol = document.createElement("div");
  rwcol.className = "col";
  rwcol.innerHTML = `<div class="small">ROOT WIDTH OF FSM (Spec / Act) mm</div>`;
  const rootSpec = makeInput({ id: "rootSpec", readOnly: true, value: parsedHeader.rootWidth || "" });
  const rootAct = makeInput({ id: "rootAct", placeholder: "Act (mm)" });
  rwcol.append(rootSpec, document.createTextNode(" "), rootAct);

  const fsmcol = document.createElement("div");
  fsmcol.className = "col";
  fsmcol.innerHTML = `<div class="small">FSM LENGTH (Spec / Act) mm</div>`;
  const fsmSpec = makeInput({ id: "fsmSpec", readOnly: true, value: parsedHeader.fsmLength || "" });
  const fsmAct = makeInput({ id: "fsmAct", placeholder: "Act (mm)" });
  fsmcol.append(fsmSpec, document.createTextNode(" "), fsmAct);
  row4.append(rwcol, fsmcol);
  blk.appendChild(row4);

  formArea.appendChild(blk);
}

/* ---------- MAIN TABLE ---------- */
function buildMainTable() {
  const blk = document.createElement("div");
  blk.className = "form-block";
  blk.innerHTML = `<strong>Main Inspection Table</strong>`;
  const table = document.createElement("table");
  table.className = "table";
  table.innerHTML = `
    <thead>
      <tr>
        <th>Sl No</th><th>Press</th><th>Sel ID</th><th>Ref</th>
        <th>X-axis</th><th>Spec (Y or Z)</th><th>Spec Dia</th>
        <th>Value from Hole edge (Act)</th><th>Actual Dia</th><th>Actual Y or Z</th><th>Offset</th><th>Result</th>
      </tr>
    </thead>
  `;
  const tb = document.createElement("tbody");
  for (let i = 0; i < 20; i++) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${i + 1}</td><td></td><td></td><td></td>`;
    for (let j = 0; j < 8; j++) {
      const td = document.createElement("td");
      td.appendChild(makeInput());
      tr.appendChild(td);
    }
    tb.appendChild(tr);
  }
  table.appendChild(tb);
  blk.appendChild(table);
  formArea.appendChild(blk);
}

/* ---------- FIRST / LAST HOLES, FLANGE, ETC. ---------- */
function buildExtraSections() {
  const blocks = [
    "First Holes of FSM - Front End",
    "Last Holes of FSM - Rear End",
    "Root / Flange / Web (reference)",
    "Part No. Location (Spec / Act)",
    "Visual / Binary Checks",
    "Remarks / Sign Section",
  ];
  blocks.forEach((b) => {
    const blk = document.createElement("div");
    blk.className = "form-block";
    blk.innerHTML = `<strong>${b}</strong>`;
    formArea.appendChild(blk);
  });
}

/* ---------- BUILD FORM ---------- */
generateFormBtn.addEventListener("click", () => {
  if (!parsedHeader) return alert("Please parse a PDF first.");
  formArea.innerHTML = "";
  buildHeaderInputs();
  buildMainTable();
  buildExtraSections();
  showMessage("Editable form created successfully.", "info");
  exportPdfBtn.disabled = false;
});

/* ---------- EXPORT TO PDF ---------- */
exportPdfBtn.addEventListener("click", async () => {
  const { jsPDF } = window.jspdf;
  const clone = formArea.cloneNode(true);
  clone.style.width = "900px";
  document.body.appendChild(clone);
  const canvas = await html2canvas(clone, { scale: 2 });
  const imgData = canvas.toDataURL("image/png");
  const pdf = new jsPDF("p", "mm", "a4");
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const imgProps = pdf.getImageProperties(imgData);
  const pdfWidth = pageWidth - 20;
  const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
  pdf.addImage(imgData, "PNG", 10, 10, pdfWidth, pdfHeight);
  pdf.save("Inspection_Form.pdf");
  document.body.removeChild(clone);
  showMessage("✅ PDF exported successfully.", "info");
});

/* ---------- MESSAGE ---------- */
function showMessage(msg, type = "info") {
  warnings.textContent = msg;
  warnings.style.padding = "8px";
  warnings.style.borderRadius = "6px";
  warnings.style.marginTop = "6px";
  warnings.style.background = type === "warn" ? "#fff6cc" : "#e8f4ff";
  warnings.style.color = type === "warn" ? "#ff8c00" : "#005fcc";
}
