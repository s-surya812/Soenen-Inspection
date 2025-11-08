const fileInput = document.getElementById('fileInput');
const extractBtn = document.getElementById('extractBtn');
const pdfViewer = document.getElementById('pdfViewer');
const rawTextEl = document.getElementById('rawText');
const headerOut = document.getElementById('headerOut');
const tablePreview = document.getElementById('tablePreview');

let currentPDFBuffer = null;
pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.9.179/pdf.worker.min.js';

fileInput.addEventListener('change', async e => {
  const f = e.target.files[0];
  if (!f) return;
  const arrayBuffer = await f.arrayBuffer();
  currentPDFBuffer = arrayBuffer;
  renderPDF(arrayBuffer);
});

extractBtn.addEventListener('click', async () => {
  if (!currentPDFBuffer) {
    alert('Please choose a PDF first.');
    return;
  }
  extractBtn.disabled = true;
  extractBtn.textContent = 'Extracting...';
  try {
    const text = await extractTextFromPDF(currentPDFBuffer.slice(0));
    showRawText(text);
    const header = parseHeader(text);
    headerOut.textContent = JSON.stringify(header, null, 2);
    const rows = detectTableLines(text);
    showTableLines(rows);
  } catch (err) {
    alert('Error: ' + err.message);
    console.error(err);
  }
  extractBtn.disabled = false;
  extractBtn.textContent = 'Extract Text & Parse';
});

async function renderPDF(buf) {
  pdfViewer.innerHTML = '';
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const p1 = await pdf.getPage(1);
  const vp = p1.getViewport({ scale: 1.2 });
  const canvas = document.createElement('canvas');
  canvas.width = vp.width; canvas.height = vp.height;
  const ctx = canvas.getContext('2d');
  await p1.render({ canvasContext: ctx, viewport: vp }).promise;
  pdfViewer.appendChild(canvas);
}

async function extractTextFromPDF(buf) {
  const doc = await pdfjsLib.getDocument({ data: buf }).promise;
  let out = '';
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    out += content.items.map(it => it.str).join(' ') + '\n---PAGE_BREAK---\n';
  }
  return out;
}

function showRawText(t) {
  const lines = t.split(/\n|---PAGE_BREAK---/).map(s => s.trim()).filter(Boolean);
  rawTextEl.textContent = lines.slice(0, 200).join('\n');
}

function parseHeader(t) {
  const h = {};
  let m = t.match(/PART\s*NUMBER.*?:\s*([A-Z0-9_]+)\s*\/\s*([A-Z0-9]+)\s*\/\s*([A-Z]+)/i);
  if (m) { h.partNumber = m[1]; h.revision = m[2]; h.hand = m[3]; }
  m = t.match(/Date\s*[:\-]\s*([0-9]{1,2}[-/][A-Za-z]{3,}[-/]\d{4})/i);
  if (m) h.date = m[1];
  m = t.match(/FORMAT\s*NO\.?\s*[:\-]*\s*([A-Z0-9\-_]+)/i);
  if (m) h.formatNo = m[1];
  m = t.match(/ROOT\s*WIDTH.*?Spec[-\s]*([0-9.]+)/i);
  if (m) h.rootWidth = m[1];
  m = t.match(/FSM\s*LENGTH.*?Spec[-\s]*([0-9.]+)/i);
  if (m) h.fsmLength = m[1];
  return h;
}

function detectTableLines(t) {
  const lines = t.split(/\n|---PAGE_BREAK---/).map(s => s.replace(/\s{2,}/g, ' ').trim()).filter(Boolean);
  return lines.filter(l => /^\d+\b/.test(l) || /\bPWX?\b|\bPWM\b|\bPFB\b|\bPFF\b/i.test(l));
}

function showTableLines(rows) {
  if (!rows.length) {
    tablePreview.textContent = 'No table rows found.';
    return;
  }
  const ol = document.createElement('ol');
  ol.style.paddingLeft = '18px';
  rows.forEach(r => {
    const li = document.createElement('li');
    li.textContent = r;
    ol.appendChild(li);
  });
  tablePreview.innerHTML = '';
  tablePreview.appendChild(ol);
}
