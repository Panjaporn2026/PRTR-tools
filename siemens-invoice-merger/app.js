// ══════════════════════════════════════════════════════
//  SIEMENS Invoice Merger — merges Signed Invoice + PO + Unsigned Invoice + a grayscale-rasterized
//  duplicate-x2 copy of the unsigned invoice's first page into one A4-normalized PDF, entirely
//  client-side via pdf-lib.js (assembly) and pdf.js (invoice-number OCR-free text read + rasterize).
// ══════════════════════════════════════════════════════

const { PDFDocument } = PDFLib;
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const state = { signed: null, po: null, unsigned: null, resultBytes: null, resultName: 'output' };

function setupDropzone(zoneId, inputId, fileLabelId, key) {
  const zone = document.getElementById(zoneId);
  const input = document.getElementById(inputId);
  const label = document.getElementById(fileLabelId);

  zone.addEventListener('click', () => input.click());
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0], key, label);
  });
  input.addEventListener('change', e => {
    if (e.target.files.length) handleFile(e.target.files[0], key, label);
  });
}

function handleFile(file, key, label) {
  if (file.type !== 'application/pdf' && !file.name.toLowerCase().endsWith('.pdf')) {
    setStatus('กรุณาเลือกไฟล์ PDF เท่านั้น', 'err');
    return;
  }
  state[key] = file;
  label.textContent = file.name;
  label.classList.add('set');
  updateMergeBtn();

  if (key === 'signed') {
    detectInvoiceNumber(file);
  }
}

// Try to auto-detect the invoice number ("Invoice No. : 526013741") from the PDF's text layer,
// falling back to pulling a 9-digit run out of the filename itself.
async function detectInvoiceNumber(file) {
  const nameInput = document.getElementById('outname');
  try {
    const buf = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
    const page = await pdf.getPage(1);
    const textContent = await page.getTextContent();
    const fullText = textContent.items.map(it => it.str).join(' ');

    let invoiceNo = null;
    const labeled = fullText.match(/Invoice\s*No\.?\s*:?\s*(\d{6,})/i);
    if (labeled) {
      invoiceNo = labeled[1];
    } else {
      const fromFilename = file.name.match(/(\d{9})/);
      if (fromFilename) invoiceNo = fromFilename[1];
    }

    if (invoiceNo && !nameInput.value.trim()) {
      nameInput.value = invoiceNo;
      setStatus(`ตรวจพบเลขที่ invoice: ${invoiceNo} (แก้ไขชื่อไฟล์ได้ถ้าไม่ตรง)`, 'ok');
    } else if (!invoiceNo) {
      setStatus('ไม่สามารถอ่านเลขที่ invoice จากไฟล์ได้อัตโนมัติ กรุณาระบุชื่อไฟล์เอง', '');
    }
  } catch (err) {
    console.warn('detectInvoiceNumber failed', err);
  }
}

function updateMergeBtn() {
  const btn = document.getElementById('mergeBtn');
  btn.disabled = !(state.signed && state.po && state.unsigned);
}

function setStatus(msg, cls) {
  const el = document.getElementById('status');
  el.textContent = msg;
  el.className = 'status-box' + (cls ? ' ' + cls : '');
}

function showWarn(msg) {
  const box = document.getElementById('warnBox');
  if (!msg) { box.classList.remove('show'); box.textContent = ''; return; }
  box.textContent = msg;
  box.classList.add('show');
}

setupDropzone('dz-signed', 'input-signed', 'file-signed', 'signed');
setupDropzone('dz-po', 'input-po', 'file-po', 'po');
setupDropzone('dz-unsigned', 'input-unsigned', 'file-unsigned', 'unsigned');

const A4_WIDTH = 595.28;
const A4_HEIGHT = 841.89;

async function fileToArrayBuffer(file) {
  return await file.arrayBuffer();
}

// Force every page of a PDFDocument's pages (already embedded into target) to A4 by scaling content.
function forcePageToA4(page) {
  const { width, height } = page.getSize();
  if (Math.abs(width - A4_WIDTH) < 2 && Math.abs(height - A4_HEIGHT) < 2) return;
  const scaleX = A4_WIDTH / width;
  const scaleY = A4_HEIGHT / height;
  const scale = Math.min(scaleX, scaleY);
  page.scaleContent(scale, scale);
  page.setSize(A4_WIDTH, A4_HEIGHT);
}

// Rasterize page 1 of a PDF file (via pdf.js) to a grayscale PNG data URL sized to fit A4.
async function rasterizeFirstPageGrayscale(arrayBuffer) {
  const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer.slice(0) });
  const pdf = await loadingTask.promise;
  const page = await pdf.getPage(1);

  const targetDpi = 150;
  const scale = targetDpi / 72;
  const viewport = page.getViewport({ scale });

  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  const ctx = canvas.getContext('2d');

  await page.render({ canvasContext: ctx, viewport }).promise;

  // Convert to grayscale
  const imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const d = imgData.data;
  for (let i = 0; i < d.length; i += 4) {
    const gray = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
    d[i] = d[i + 1] = d[i + 2] = gray;
  }
  ctx.putImageData(imgData, 0, 0);

  const dataUrl = canvas.toDataURL('image/png');
  return { dataUrl, pxWidth: canvas.width, pxHeight: canvas.height };
}

document.getElementById('mergeBtn').addEventListener('click', async () => {
  showWarn('');
  document.getElementById('result').classList.remove('show');
  const btn = document.getElementById('mergeBtn');
  btn.disabled = true;
  setStatus('กำลังรวมไฟล์...', 'info');
  try {
    const [signedBuf, poBuf, unsignedBuf] = await Promise.all([
      fileToArrayBuffer(state.signed),
      fileToArrayBuffer(state.po),
      fileToArrayBuffer(state.unsigned),
    ]);

    const outDoc = await PDFDocument.create();

    // 1. Signed invoice
    const signedSrc = await PDFDocument.load(signedBuf);
    const signedPages = await outDoc.copyPages(signedSrc, signedSrc.getPageIndices());
    signedPages.forEach(p => { outDoc.addPage(p); forcePageToA4(p); });

    // 2. PO (all pages)
    const poSrc = await PDFDocument.load(poBuf);
    const poPages = await outDoc.copyPages(poSrc, poSrc.getPageIndices());
    poPages.forEach(p => { outDoc.addPage(p); forcePageToA4(p); });

    // 3. Unsigned invoice
    const unsignedSrc = await PDFDocument.load(unsignedBuf);
    const unsignedPages = await outDoc.copyPages(unsignedSrc, unsignedSrc.getPageIndices());
    unsignedPages.forEach(p => { outDoc.addPage(p); forcePageToA4(p); });

    // 4. BW page from first page of unsigned invoice, duplicated x2, forced to A4
    setStatus('กำลังสร้างหน้าขาวดำ...', 'info');
    const { dataUrl, pxWidth, pxHeight } = await rasterizeFirstPageGrayscale(unsignedBuf);
    const pngBytes = await (await fetch(dataUrl)).arrayBuffer();
    const pngImage = await outDoc.embedPng(pngBytes);

    // Fit image into A4 page, preserving aspect ratio, centered
    const imgAspect = pxWidth / pxHeight;
    const a4Aspect = A4_WIDTH / A4_HEIGHT;
    let drawWidth, drawHeight;
    if (imgAspect > a4Aspect) {
      drawWidth = A4_WIDTH;
      drawHeight = A4_WIDTH / imgAspect;
    } else {
      drawHeight = A4_HEIGHT;
      drawWidth = A4_HEIGHT * imgAspect;
    }
    const offsetX = (A4_WIDTH - drawWidth) / 2;
    const offsetY = (A4_HEIGHT - drawHeight) / 2;

    for (let i = 0; i < 2; i++) {
      const bwPage = outDoc.addPage([A4_WIDTH, A4_HEIGHT]);
      bwPage.drawImage(pngImage, {
        x: offsetX, y: offsetY, width: drawWidth, height: drawHeight,
      });
    }

    const outBytes = await outDoc.save();
    state.resultBytes = outBytes;

    const rawName = document.getElementById('outname').value.trim();
    const finalName = (rawName || 'merged_invoice') + '.pdf';
    state.resultName = finalName;

    document.getElementById('resultName').textContent = finalName;
    document.querySelector('#result .rsub').textContent = `รวมไฟล์เสร็จแล้ว — ${outDoc.getPageCount()} หน้า, ขนาด A4 ทุกหน้า`;
    document.getElementById('result').classList.add('show');
    setStatus(`รวมไฟล์สำเร็จ (${outDoc.getPageCount()} หน้า)`, 'ok');
  } catch (err) {
    console.error(err);
    setStatus('เกิดข้อผิดพลาด: ' + err.message, 'err');
    showWarn('หากปัญหายังคงอยู่ ลองตรวจสอบว่าไฟล์ PDF ทั้ง 3 ไฟล์ไม่ได้เสียหายหรือถูกล็อกด้วยรหัสผ่าน');
  } finally {
    btn.disabled = false;
  }
});

document.getElementById('resetBtn').addEventListener('click', () => {
  // Clear selected files
  state.signed = null;
  state.po = null;
  state.unsigned = null;
  state.resultBytes = null;
  state.resultName = 'output';

  // Reset file inputs so the same file can be re-selected later if needed
  document.getElementById('input-signed').value = '';
  document.getElementById('input-po').value = '';
  document.getElementById('input-unsigned').value = '';

  // Reset dropzone labels
  const resetLabel = (id, text) => {
    const el = document.getElementById(id);
    el.textContent = text;
    el.classList.remove('set');
  };
  resetLabel('file-signed', 'ยังไม่ได้เลือกไฟล์');
  resetLabel('file-po', 'ยังไม่ได้เลือกไฟล์');
  resetLabel('file-unsigned', 'ยังไม่ได้เลือกไฟล์');

  // Reset filename field
  document.getElementById('outname').value = '';

  // Reset result panel and status/warning
  document.getElementById('result').classList.remove('show');
  document.getElementById('resultName').textContent = '-';
  setStatus('');
  showWarn('');

  updateMergeBtn();
});

document.getElementById('downloadBtn').addEventListener('click', () => {
  if (!state.resultBytes) return;
  const blob = new Blob([state.resultBytes], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = state.resultName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});
