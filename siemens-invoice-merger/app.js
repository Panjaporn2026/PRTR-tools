// ══════════════════════════════════════════════════════
//  SIEMENS Invoice Merger — drop in any mix of signed/unsigned invoice + PO PDFs across multiple
//  invoices at once; the tool reads each file's own content to classify it and auto-match files
//  into per-invoice groups (by invoice number and PO number), then batch-merges every complete
//  group into an A4-normalized PDF (+ grayscale duplicate-x2 BW page) and zips all results.
//  Entirely client-side via pdf-lib.js (assembly), pdf.js (text read + rasterize) and JSZip.
// ══════════════════════════════════════════════════════

const { PDFDocument } = PDFLib;
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const A4_WIDTH = 595.28;
const A4_HEIGHT = 841.89;

// ---- Global state ----
// allFiles: [{file, kind: 'invoice'|'po'|'unknown', invoiceNo, poRef, orderNo, isSigned, text}]
let allFiles = [];
// groups: Map<invoiceNo, {invoiceNo, poRef, signed: fileEntry|null, unsigned: fileEntry|null, po: fileEntry|null}>
let groups = new Map();
let unmatchedFiles = [];
let processedResults = []; // {name, bytes}

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

// ---- Dropzone (multi-file) ----
const dz = document.getElementById('dz-all');
const inputAll = document.getElementById('input-all');
dz.addEventListener('click', () => inputAll.click());
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag'); });
dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
dz.addEventListener('drop', e => {
  e.preventDefault();
  dz.classList.remove('drag');
  if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
});
inputAll.addEventListener('change', e => {
  if (e.target.files.length) handleFiles(e.target.files);
});

async function handleFiles(fileList) {
  const pdfFiles = Array.from(fileList).filter(f =>
    f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf')
  );
  if (!pdfFiles.length) {
    setStatus('กรุณาเลือกไฟล์ PDF เท่านั้น', 'err');
    return;
  }
  setStatus(`กำลังอ่านไฟล์ ${pdfFiles.length} ไฟล์...`);
  document.getElementById('fileCountPill').style.display = 'none';

  const entries = [];
  for (const file of pdfFiles) {
    const entry = await analyzeFile(file);
    entries.push(entry);
  }
  allFiles = allFiles.concat(entries);

  document.getElementById('fileCountPill').textContent = `รับไฟล์แล้วทั้งหมด ${allFiles.length} ไฟล์`;
  document.getElementById('fileCountPill').style.display = 'inline-block';

  buildGroups();
  renderGroups();
  setStatus(`อ่านไฟล์เสร็จแล้ว (${pdfFiles.length} ไฟล์ใหม่)`, 'ok');
}

// Extract text + classify a single PDF file.
async function analyzeFile(file) {
  const entry = {
    file, name: file.name,
    kind: 'unknown', // 'invoice' | 'po'  ('invoice' covers both real invoices and debit memos)
    docType: null,   // 'INVOICE' | 'DEBIT MEMO' — for display only, doesn't affect matching
    invoiceNo: null,
    poRef: null,     // PO number mentioned inside an invoice/memo
    orderNo: null,   // Order No. inside a PO file
    isSigned: false,
    text: '',
  };
  try {
    const buf = await file.arrayBuffer();

    // --- Signature detection at the raw PDF byte level ---
    // Count /ByteRange occurrences (see countByteRangeOccurrences for why a raw count, not a
    // simple present/absent check, is needed). The THRESHOLD for "signed" differs by document
    // template, discovered from real signed/unsigned pairs of each type:
    //   - INVOICE template bakes in one EMPTY /ByteRange placeholder even when unsigned, so an
    //     unsigned invoice already has 1; a genuinely signed one has 2 (placeholder + real sig).
    //   - DEBIT MEMO template has no such placeholder: unsigned memos have 0, a genuinely signed
    //     one has 1 (just the real signature, no baked-in empty slot).
    // The doc type (from the label text below) decides which threshold applies; until the type
    // is known further down, byteRangeCount is stashed and isSigned is finalized afterward.
    const byteRangeCount = countByteRangeOccurrences(buf);

    const pdf = await pdfjsLib.getDocument({ data: buf.slice(0) }).promise;
    const page = await pdf.getPage(1);
    const textContent = await page.getTextContent();
    const fullText = textContent.items.map(it => it.str).join(' ');
    entry.text = fullText;

    // NOTE on layout: these invoice PDFs render a hidden/leading data-row BEFORE the visible
    // labeled layout — e.g. "Siemens Energy Limited 526017993 31 Jul 2026 ... Invoice No. Date
    // Ref. no. ...". The label "Invoice No." and its actual value are NOT adjacent in the
    // extracted text (pdf.js gives text in draw order, and the value is drawn separately from
    // the label). So "Invoice No.\s*:?\s*(\d+)" style regexes reliably fail on these files —
    // the invoice number has to be pulled from that leading data row instead.
    const purchaseOrderTitle = /Purchase\s*order/i.test(fullText);
    const orderMatch = fullText.match(/Order\s*No\.?\s*:?\s*(\d{6,})/i);
    // Tolerant of pdf.js inserting stray spaces between text runs, e.g. "PO . 4510261185",
    // "PO.4510261185", "PO 4510261185".
    const poRefMatch = fullText.match(/PO\s*\.?\s*(\d{6,})/i);

    // Two different label/value layouts have been observed across real PRTR PDFs — neither
    // is reliably "labels immediately followed by their own value" in pdf.js's draw-order text:
    //  1. INVOICE layout: a hidden leading data-row up front ("Siemens Energy Limited
    //     526017993 31 Jul 2026 ..."), with the "Invoice No." label appearing much later,
    //     unattached to its value.
    //  2. DEBIT MEMO layout: labels are grouped together first ("Memo no. Date Ref. no. : : :"),
    //     then ALL their values follow as one run right after ("626000842 20 Jul 2026 ...").
    // Both patterns need their own regex; relying on only one silently drops the other doc type.
    const leadingRowMatch = fullText.match(/^\s*\S.{0,40}?\b(\d{9})\b/);
    const labeledMatch = fullText.match(/Invoice\s*No\.?\s*:?\s*(\d{6,})/i);
    // "Memo no." label followed (possibly much later in the joined string, past other label
    // words and ":" separators) by the first 9-digit run — that first run is always the memo no.
    const memoMatch = fullText.match(/Memo\s*no\.?\s*(?:Date\s*)?(?:Ref\.?\s*no\.?\s*)?:?\s*:?\s*:?\s*(\d{9})/i);
    // Fallback: pull the invoice/memo number straight out of the filename (PRTR convention:
    // "<invoiceOrMemoNo><SMcode>.pdf", e.g. "526017993SM289.pdf", "626000842SM110.pdf").
    const filenameInvoiceMatch = file.name.match(/^(\d{9})/);

    if (orderMatch && purchaseOrderTitle) {
      entry.kind = 'po';
      entry.orderNo = orderMatch[1];
    } else if (leadingRowMatch || labeledMatch || memoMatch || filenameInvoiceMatch) {
      entry.kind = 'invoice';
      // PRTR issues either an INVOICE or a DEBIT MEMO for the same workflow (same customer
      // signature requirement, same PO reference, same file-naming convention) — treat both
      // as the same matchable document type, just remember which one it was for display.
      entry.docType = /DEBIT\s*MEMO/i.test(fullText) ? 'DEBIT MEMO' : 'INVOICE';
      entry.invoiceNo = (labeledMatch && labeledMatch[1])
        || (memoMatch && memoMatch[1])
        || (leadingRowMatch && leadingRowMatch[1])
        || (filenameInvoiceMatch && filenameInvoiceMatch[1]);
      if (poRefMatch) entry.poRef = poRefMatch[1];
    }

    // Apply the doc-type-specific signed threshold now that docType is known.
    // DEBIT MEMO has no baked-in placeholder -> signed threshold is >=1.
    // INVOICE (and anything unrecognized, safer default) has a baked-in placeholder -> >=2.
    if (entry.docType === 'DEBIT MEMO') {
      entry.isSigned = byteRangeCount >= 1;
    } else {
      entry.isSigned = byteRangeCount >= 2;
    }

    // Secondary text-based signal (works when the caption IS plain text)
    if (!entry.isSigned && /Digitally signed by/i.test(fullText)) {
      entry.isSigned = true;
    }
  } catch (err) {
    console.warn('analyzeFile failed for', file.name, err);
  }

  // Filename fallbacks (only used when content-based detection found nothing)
  if (entry.kind === 'unknown') {
    if (/PO[_\s]/i.test(file.name) || /purchase.?order/i.test(file.name)) {
      entry.kind = 'po';
      const m = file.name.match(/(\d{7,})/);
      if (m) entry.orderNo = m[1];
    } else {
      entry.kind = 'invoice';
      const m = file.name.match(/(\d{9})/);
      if (m) entry.invoiceNo = m[1];
    }
  }
  if (!entry.isSigned && /signed/i.test(file.name) && !/unsigned/i.test(file.name)) {
    entry.isSigned = true;
  }

  return entry;
}

// Count how many times "/ByteRange" appears in the raw PDF bytes.
//
// IMPORTANT: different PRTR document templates embed a different number of /ByteRange
// occurrences even in their UNSIGNED state:
//   - INVOICE template bakes in ONE empty /ByteRange + /Sig placeholder even when unsigned
//     (the signature widget exists on the page layout regardless of whether anyone signed it).
//   - DEBIT MEMO template has no such placeholder at all — unsigned memos have ZERO.
// So a simple "/ByteRange present" check can't tell signed from unsigned for either type, and a
// single fixed threshold doesn't work across both types either. This function only returns the
// raw count; analyzeFile() applies the correct doc-type-specific threshold once it knows whether
// the file is an INVOICE or a DEBIT MEMO (see the isSigned assignment further up).
function countByteRangeOccurrences(arrayBuffer) {
  try {
    const bytes = new Uint8Array(arrayBuffer);
    const chunkSize = 2_000_000;
    const overlap = 64; // covers the "/ByteRange" marker (10 chars) plus margin, so a marker
                         // split across a chunk boundary is never double-counted or missed.
    let byteRangeCount = 0;
    let tailStr = '';
    for (let offset = 0; offset < bytes.length; offset += chunkSize) {
      const end = Math.min(offset + chunkSize, bytes.length);
      const slice = bytes.subarray(offset, end);
      let str = tailStr;
      for (let i = 0; i < slice.length; i++) str += String.fromCharCode(slice[i]);

      const matches = str.match(/\/ByteRange/g);
      const matchCount = matches ? matches.length : 0;
      // Since tailStr overlaps with the end of the previous chunk, subtract matches that were
      // already counted there (tailStr alone can contain at most one partial/whole marker).
      const tailMatches = tailStr.match(/\/ByteRange/g);
      const tailMatchCount = tailMatches ? tailMatches.length : 0;
      byteRangeCount += matchCount - tailMatchCount;

      tailStr = str.slice(-overlap);
    }
    return byteRangeCount;
  } catch (err) {
    console.warn('countByteRangeOccurrences failed', err);
    return 0;
  }
}

function buildGroups() {
  groups = new Map();
  unmatchedFiles = [];

  const invoiceFiles = allFiles.filter(f => f.kind === 'invoice');
  const poFiles = allFiles.filter(f => f.kind === 'po');

  // Group invoice files (signed + unsigned) by invoiceNo
  for (const f of invoiceFiles) {
    if (!f.invoiceNo) { unmatchedFiles.push(f); continue; }
    if (!groups.has(f.invoiceNo)) {
      groups.set(f.invoiceNo, { invoiceNo: f.invoiceNo, docType: f.docType || 'INVOICE', poRef: f.poRef || null, signed: null, unsigned: null, po: null });
    }
    const g = groups.get(f.invoiceNo);
    if (!g.poRef && f.poRef) g.poRef = f.poRef;
    if (f.docType) g.docType = f.docType;
    if (f.isSigned) {
      if (!g.signed) g.signed = f; else unmatchedFiles.push(f);
    } else {
      if (!g.unsigned) g.unsigned = f; else unmatchedFiles.push(f);
    }
  }

  // Match PO files to groups by poRef === orderNo.
  // One PO can legitimately back multiple invoices/debit memos (shared PO number),
  // so a PO file is reused across every matching group rather than being claimed once.
  const usedPoFiles = new Set();
  for (const g of groups.values()) {
    if (!g.poRef) continue;
    const match = poFiles.find(p => p.orderNo === g.poRef);
    if (match) {
      g.po = match;
      usedPoFiles.add(match);
    }
  }
  for (const p of poFiles) {
    if (!usedPoFiles.has(p)) unmatchedFiles.push(p);
  }
}

function renderGroups() {
  const groupsCard = document.getElementById('groupsCard');
  const processCard = document.getElementById('processCard');
  const tbody = document.getElementById('groupsBody');
  const unmatchedBox = document.getElementById('unmatchedBox');

  if (groups.size === 0 && unmatchedFiles.length === 0) {
    groupsCard.style.display = 'none';
    processCard.style.display = 'none';
    return;
  }
  groupsCard.style.display = 'block';
  processCard.style.display = groups.size > 0 ? 'block' : 'none';

  tbody.innerHTML = '';
  const sortedKeys = Array.from(groups.keys()).sort();

  // Count how many invoice groups reference each PO number, to flag shared POs.
  const poRefCounts = {};
  for (const g of groups.values()) {
    if (g.poRef) poRefCounts[g.poRef] = (poRefCounts[g.poRef] || 0) + 1;
  }

  for (const key of sortedKeys) {
    const g = groups.get(key);
    const complete = g.signed && g.unsigned && g.po;
    const tr = document.createElement('tr');

    const chip = (ok, label) => `<span class="file-chip${ok ? '' : ' missing'}"><span class="dot"></span>${label}</span>`;
    const shared = g.poRef && poRefCounts[g.poRef] > 1;

    const typeLabel = g.docType === 'DEBIT MEMO' ? 'Memo' : 'Invoice';
    tr.innerHTML = `
      <td class="inv-key">${g.invoiceNo}<div style="font-weight:400; font-size:11px; color:#90a4ae; margin-top:2px;">${typeLabel}</div></td>
      <td>${g.poRef ? g.poRef + (shared ? ' <span style="color:#e65100; font-size:11px;">(ใช้ร่วมกัน ' + poRefCounts[g.poRef] + ' invoice)</span>' : '') : '<span style="color:#c62828">ไม่พบ</span>'}</td>
      <td>
        ${chip(!!g.signed, g.signed ? 'Signed ✓' : 'Signed ✗')}
        ${chip(!!g.unsigned, g.unsigned ? 'Unsigned ✓' : 'Unsigned ✗')}
        ${chip(!!g.po, g.po ? 'PO ✓' : 'PO ✗')}
      </td>
      <td class="grp-status ${complete ? 'ready' : 'incomplete'}">${complete ? 'พร้อมประมวลผล' : 'ไฟล์ไม่ครบ'}</td>
    `;
    tbody.appendChild(tr);
  }

  if (unmatchedFiles.length > 0) {
    unmatchedBox.style.display = 'block';
    unmatchedBox.innerHTML = `<b>ไฟล์ที่จับคู่ไม่ได้ (${unmatchedFiles.length}):</b><br>` +
      unmatchedFiles.map(f => `• ${f.name} ${f.invoiceNo ? `(อ่านได้ invoice no. ${f.invoiceNo} ซ้ำกับไฟล์อื่น — signed/unsigned เกินคู่)` : f.orderNo ? `(PO ${f.orderNo} ไม่ตรงกับ invoice ใดเลย)` : '(ไม่สามารถระบุประเภทได้)'}`).join('<br>');
  } else {
    unmatchedBox.style.display = 'none';
  }
}

// ---- A4 forcing ----
function forcePageToA4(page) {
  const { width, height } = page.getSize();
  if (Math.abs(width - A4_WIDTH) < 2 && Math.abs(height - A4_HEIGHT) < 2) return;
  const scale = Math.min(A4_WIDTH / width, A4_HEIGHT / height);
  page.scaleContent(scale, scale);
  page.setSize(A4_WIDTH, A4_HEIGHT);
}

async function rasterizeFirstPageGrayscale(arrayBuffer) {
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer.slice(0) }).promise;
  const page = await pdf.getPage(1);
  const scale = 150 / 72;
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  const ctx = canvas.getContext('2d');
  await page.render({ canvasContext: ctx, viewport }).promise;
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

async function mergeOneGroup(g) {
  const [signedBuf, poBuf, unsignedBuf] = await Promise.all([
    g.signed.file.arrayBuffer(),
    g.po.file.arrayBuffer(),
    g.unsigned.file.arrayBuffer(),
  ]);

  const outDoc = await PDFDocument.create();

  const signedSrc = await PDFDocument.load(signedBuf);
  (await outDoc.copyPages(signedSrc, signedSrc.getPageIndices())).forEach(p => { outDoc.addPage(p); forcePageToA4(p); });

  const poSrc = await PDFDocument.load(poBuf);
  (await outDoc.copyPages(poSrc, poSrc.getPageIndices())).forEach(p => { outDoc.addPage(p); forcePageToA4(p); });

  const unsignedSrc = await PDFDocument.load(unsignedBuf);
  (await outDoc.copyPages(unsignedSrc, unsignedSrc.getPageIndices())).forEach(p => { outDoc.addPage(p); forcePageToA4(p); });

  const { dataUrl, pxWidth, pxHeight } = await rasterizeFirstPageGrayscale(unsignedBuf);
  const pngBytes = await (await fetch(dataUrl)).arrayBuffer();
  const pngImage = await outDoc.embedPng(pngBytes);

  const imgAspect = pxWidth / pxHeight;
  const a4Aspect = A4_WIDTH / A4_HEIGHT;
  let drawWidth, drawHeight;
  if (imgAspect > a4Aspect) { drawWidth = A4_WIDTH; drawHeight = A4_WIDTH / imgAspect; }
  else { drawHeight = A4_HEIGHT; drawWidth = A4_HEIGHT * imgAspect; }
  const offsetX = (A4_WIDTH - drawWidth) / 2;
  const offsetY = (A4_HEIGHT - drawHeight) / 2;

  for (let i = 0; i < 2; i++) {
    const bwPage = outDoc.addPage([A4_WIDTH, A4_HEIGHT]);
    bwPage.drawImage(pngImage, { x: offsetX, y: offsetY, width: drawWidth, height: drawHeight });
  }

  return await outDoc.save();
}

function renderProgressList(readyKeys) {
  const list = document.getElementById('progressList');
  list.innerHTML = '';
  for (const key of readyKeys) {
    const row = document.createElement('div');
    row.className = 'progress-row';
    row.id = `prog-${key}`;
    row.innerHTML = `<div class="pname">${key}.pdf</div><div class="pstate wait" id="pstate-${key}">รอคิว</div>`;
    list.appendChild(row);
  }
}

function setProgressState(key, state, label) {
  const el = document.getElementById(`pstate-${key}`);
  if (!el) return;
  el.className = 'pstate ' + state;
  el.textContent = label;
}

document.getElementById('processBtn').addEventListener('click', async () => {
  showWarn('');
  document.getElementById('resultSummary').classList.remove('show');
  processedResults = [];

  const readyEntries = Array.from(groups.values()).filter(g => g.signed && g.unsigned && g.po);
  if (!readyEntries.length) {
    setStatus('ไม่มีชุดไฟล์ที่ครบพร้อมประมวลผล', 'err');
    return;
  }

  const btn = document.getElementById('processBtn');
  btn.disabled = true;
  setStatus(`กำลังประมวลผล ${readyEntries.length} invoice...`);
  renderProgressList(readyEntries.map(g => g.invoiceNo));

  let successCount = 0;
  let failCount = 0;

  for (const g of readyEntries) {
    setProgressState(g.invoiceNo, 'run', 'กำลังรวม...');
    try {
      const bytes = await mergeOneGroup(g);
      processedResults.push({ name: `${g.invoiceNo}.pdf`, bytes });
      setProgressState(g.invoiceNo, 'done', 'เสร็จแล้ว');
      successCount++;
    } catch (err) {
      console.error(`Failed merging ${g.invoiceNo}`, err);
      setProgressState(g.invoiceNo, 'fail', 'ผิดพลาด');
      failCount++;
    }
  }

  btn.disabled = false;
  if (failCount > 0) {
    showWarn(`มี ${failCount} invoice ที่ประมวลผลไม่สำเร็จ กรุณาตรวจสอบไฟล์ต้นฉบับ (อาจเสียหายหรือถูกล็อกด้วยรหัสผ่าน)`);
  }
  setStatus(`ประมวลผลเสร็จสิ้น: สำเร็จ ${successCount} / ล้มเหลว ${failCount}`, failCount ? 'err' : 'ok');

  if (successCount > 0) {
    const summary = document.getElementById('resultSummary');
    document.getElementById('resultText').innerHTML = `รวมไฟล์สำเร็จ <b>${successCount}</b> invoice พร้อมดาวน์โหลด`;
    summary.classList.add('show');
  }
});

document.getElementById('downloadZipBtn').addEventListener('click', async () => {
  if (!processedResults.length) return;
  const zip = new JSZip();
  for (const r of processedResults) {
    zip.file(r.name, r.bytes);
  }
  const blob = await zip.generateAsync({ type: 'blob' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'siemens_invoices.zip';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

document.getElementById('resetBtn').addEventListener('click', () => {
  allFiles = [];
  groups = new Map();
  unmatchedFiles = [];
  processedResults = [];
  document.getElementById('input-all').value = '';
  document.getElementById('fileCountPill').style.display = 'none';
  document.getElementById('groupsCard').style.display = 'none';
  document.getElementById('processCard').style.display = 'none';
  document.getElementById('progressList').innerHTML = '';
  document.getElementById('resultSummary').classList.remove('show');
  setStatus('');
  showWarn('');
});
