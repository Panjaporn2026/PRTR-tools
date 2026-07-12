// ══════════════════════════════════════════════════════
//  UI wiring: template upload -> GL_Invoice upload (multi) -> Detail-of-Invoice upload ->
//  option pickers (grouping key / Reference No. / no-VAT income accounts, all real-detected,
//  never hardcoded) -> preview/confirm -> fill -> verify -> optional password lock -> download.
// ══════════════════════════════════════════════════════

var state = {
  templateBuf: null, templateWb: null, templateSst: null,
  lookups: null, // { costIncomeMap, seriesTable, vatCodeTable, paymentGroupMap, customerMap }
  glParsedFiles: [], // [{sourceLabel, colIndexByName, dataRows, headerRowNum}]
  detailAoa: null, detailParsed: null,
  candidateKeyColumns: [],
  groupingKeyCol: null, refNoCol: null,
  mappedRows: null, unmappedRows: null,
  incomeAccountsPresent: [], noVatIncomeAccounts: new Set(),
  invoices: null,
  resultBytes: null
};

function $(id) { return document.getElementById(id); }
function setStatus(el, html, cls) {
  el.className = 'mapping-status show' + (cls ? ' ' + cls : '');
  el.innerHTML = html;
}
function enableSection(id) { $(id).classList.remove('disabled'); }

async function readFileBuf(file) {
  return await file.arrayBuffer();
}

// SheetJS-based aoa reader for arbitrary, messy real-world exports (GL_Invoice/Detail of
// Invoice) -- same technique doi-tools/rename-lock's readDetailOfInvoiceMeta already uses for
// this file family (label/value split cells, Date-typed cells, decoy legend rows).
function sheetToAoa(buf) {
  var wb = XLSX.read(buf, { type: 'array', cellDates: true });
  var ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
}

// ── Step 1: DTW template ─────────────────────────────────────────────────────────────────────

function wireDropzone(dropEl, inputEl, onFiles, multi) {
  dropEl.addEventListener('click', function () { inputEl.click(); });
  dropEl.addEventListener('dragover', function (e) { e.preventDefault(); dropEl.classList.add('dragover'); });
  dropEl.addEventListener('dragleave', function () { dropEl.classList.remove('dragover'); });
  dropEl.addEventListener('drop', function (e) {
    e.preventDefault(); dropEl.classList.remove('dragover');
    onFiles(multi ? e.dataTransfer.files : [e.dataTransfer.files[0]]);
  });
  inputEl.addEventListener('change', function (e) { onFiles(e.target.files); e.target.value = ''; });
}

wireDropzone($('templateDrop'), $('templateInput'), function (files) { loadTemplate(files[0]); }, false);

async function loadTemplate(file) {
  if (!file) return;
  var statusEl = $('templateStatus');
  try {
    setStatus(statusEl, '⏳ กำลังอ่านไฟล์ template...', '');
    var buf = await readFileBuf(file);
    var wb = await loadWorkbook(buf);
    var sst = await getSharedStrings(wb);
    var kumueGrid = await readSheetGrid(wb, 'คู่มือ');
    var outsourceGrid = await readSheetGrid(wb, 'OUTSOURCE');
    var costIncomeMap = buildCostIncomeMap(kumueGrid.aoa);
    var seriesTable = buildSeriesTable(kumueGrid.aoa);
    var vatCodeTable = buildVatCodeTable(kumueGrid.aoa);
    var paymentGroupMap = buildPaymentGroupMap(kumueGrid.aoa);
    var customerMap = buildCustomerLookup(outsourceGrid.aoa);

    state.templateBuf = buf;
    state.templateWb = wb;
    state.templateSst = sst;
    state.lookups = { costIncomeMap: costIncomeMap, seriesTable: seriesTable, vatCodeTable: vatCodeTable, paymentGroupMap: paymentGroupMap, customerMap: customerMap };

    setStatus(statusEl, '✅ อ่าน template สำเร็จ — พบตาราง Cost→Income ' + costIncomeMap.size + ' รายการ, ลูกค้าใน OUTSOURCE ' + customerMap.size + ' ราย', 'ms-ok');
    enableSection('glSection');
  } catch (err) {
    setStatus(statusEl, '❌ ' + err.message, 'ms-err');
  }
}

// ── Step 2: GL_Invoice (multi) ───────────────────────────────────────────────────────────────

wireDropzone($('glDrop'), $('glInput'), function (files) { addGLFiles(files); }, true);

function renderGLFileList() {
  var wrap = $('glFileList');
  wrap.innerHTML = state.glParsedFiles.map(function (pf, i) {
    return '<div class="file-row"><span>📄</span><span class="row-orig">' + esc(pf.sourceLabel) + '</span>' +
      '<button class="row-remove" onclick="removeGLFile(' + i + ')">✕</button></div>';
  }).join('');
}
function removeGLFile(i) {
  state.glParsedFiles.splice(i, 1);
  renderGLFileList();
  refreshKeyColumnPickers();
}

function addGLFiles(files) {
  var statusEl = $('glStatus');
  var errors = [];
  Promise.all(Array.prototype.map.call(files, function (file) {
    return readFileBuf(file).then(function (buf) {
      var aoa = sheetToAoa(buf);
      return parseGLInvoiceAoa(aoa, file.name);
    }).catch(function (err) { errors.push(file.name + ': ' + err.message); return null; });
  })).then(function (parsed) {
    parsed.filter(Boolean).forEach(function (pf) { state.glParsedFiles.push(pf); });
    renderGLFileList();
    if (errors.length) setStatus(statusEl, '⚠️ ' + errors.join(' / '), 'ms-warn');
    else setStatus(statusEl, '✅ อัปโหลดแล้ว ' + state.glParsedFiles.length + ' ไฟล์', 'ms-ok');
    if (state.glParsedFiles.length) {
      enableSection('detailSection');
      refreshKeyColumnPickers();
    }
  });
}

function refreshKeyColumnPickers() {
  if (!state.glParsedFiles.length) { $('optionsSection').classList.add('disabled'); return; }
  var candidates;
  try {
    candidates = detectCandidateKeyColumns(state.glParsedFiles);
  } catch (e) { setStatus($('glStatus'), '❌ ' + e.message, 'ms-err'); return; }
  state.candidateKeyColumns = candidates;
  renderRadioChips('groupingKeyList', 'groupingKey', candidates, function (val) {
    state.groupingKeyCol = val;
    onKeyColumnsChanged();
  });
  renderRadioChips('refNoList', 'refNoCol', candidates, function (val) {
    state.refNoCol = val;
    onKeyColumnsChanged();
  });
  enableSection('optionsSection');
}

function renderRadioChips(containerId, name, options, onChange) {
  var el = $(containerId);
  if (!options.length) { el.innerHTML = '<span style="color:#c62828;font-size:12px">ไม่พบคอลัมน์ที่ใช้เป็นตัวเลือกได้ในไฟล์ GL_Invoice ที่อัปโหลด</span>'; return; }
  el.innerHTML = options.map(function (opt, i) {
    return '<label class="an-col-chip"><input type="radio" name="' + name + '" value="' + esc(opt) + '"' + (i === 0 ? ' checked' : '') + '> ' + esc(opt) + '</label>';
  }).join('');
  Array.prototype.forEach.call(el.querySelectorAll('input'), function (input) {
    input.addEventListener('change', function () { if (input.checked) onChange(input.value); });
  });
  if (options.length) onChange(options[0]);
}

function onKeyColumnsChanged() {
  if (!state.groupingKeyCol || !state.refNoCol) return;
  try {
    var merged = mergeGLInvoiceFiles(state.glParsedFiles);
    var expense = filterExpenseRows(merged);
    var result = mapCostToIncome(expense, state.lookups.costIncomeMap);
    state.mappedRows = result.mapped;
    state.unmappedRows = result.unmapped;
    var accounts = new Map();
    result.mapped.forEach(function (e) {
      var cur = accounts.get(e.incomeAccount) || { incomeAccount: e.incomeAccount, incomeName: e.incomeName, amount: 0 };
      cur.amount += Number(e.row.get('Amount')) || 0;
      accounts.set(e.incomeAccount, cur);
    });
    state.incomeAccountsPresent = Array.from(accounts.values()).sort(function (a, b) { return b.amount - a.amount; });
    renderNoVatChecklist();
    enableSection('previewSection');
  } catch (e) {
    setStatus($('glStatus'), '❌ ' + e.message, 'ms-err');
  }
}

function renderNoVatChecklist() {
  var wrap = $('noVatWrap'), list = $('noVatList');
  if (!state.incomeAccountsPresent.length) { wrap.style.display = 'none'; return; }
  wrap.style.display = '';
  list.innerHTML = state.incomeAccountsPresent.map(function (a) {
    var checked = state.noVatIncomeAccounts.has(a.incomeAccount) ? ' checked' : '';
    return '<label class="an-col-chip"><input type="checkbox" value="' + esc(a.incomeAccount) + '"' + checked + ' onchange="toggleNoVatAccount(this)"> ' +
      esc(a.incomeAccount) + ' ' + esc(a.incomeName) + ' (' + a.amount.toLocaleString('en-US', { maximumFractionDigits: 2 }) + ')</label>';
  }).join('');
}
function toggleNoVatAccount(input) {
  if (input.checked) state.noVatIncomeAccounts.add(input.value);
  else state.noVatIncomeAccounts.delete(input.value);
}

// ── Step 3: Detail of Invoice ─────────────────────────────────────────────────────────────────

wireDropzone($('detailDrop'), $('detailInput'), function (files) { loadDetail(files[0]); }, false);

async function loadDetail(file) {
  if (!file) return;
  var statusEl = $('detailStatus');
  try {
    setStatus(statusEl, '⏳ กำลังอ่านไฟล์...', '');
    var buf = await readFileBuf(file);
    var aoa = sheetToAoa(buf);
    state.detailAoa = aoa;
    state.detailParsed = parseDetailOfInvoice(aoa);
    var msg = '✅ อ่านไฟล์สำเร็จ — พบพนักงาน ' + state.detailParsed.byAltId.size + ' ราย';
    if (!state.detailParsed.hasTotalColumn) msg += ' (ไม่พบคอลัมน์ยอดรวมต่อคน — จะเช็คได้แค่สถานะ Pending ไม่เช็คยอด reconcile)';
    setStatus(statusEl, msg, 'ms-ok');
  } catch (err) {
    setStatus(statusEl, '❌ ' + err.message + ' — ยังสามารถทำงานต่อได้โดยข้ามการเช็ค Pending/reconcile', 'ms-warn');
    state.detailParsed = null;
  }
}

// ── Step 5: preview ───────────────────────────────────────────────────────────────────────────

$('btnPreview').addEventListener('click', function () {
  if (!state.mappedRows) { alert('กรุณาอัปโหลดและเลือกตัวเลือกให้ครบก่อน'); return; }
  try {
    state.invoices = buildInvoiceGroups(state.mappedRows, state.groupingKeyCol, state.refNoCol, state.noVatIncomeAccounts);
    renderPreview();
    enableSection('runSection');
    $('btnRun').disabled = false;
  } catch (e) {
    $('previewBody').innerHTML = '<div class="warn-box err">❌ ' + esc(e.message) + '</div>';
  }
});

function renderPreview() {
  var body = $('previewBody');
  var invoices = state.invoices;
  var totalVat = invoices.filter(function (i) { return !i.isNoVat; }).reduce(function (s, i) { return s + i.totalBeforeVat; }, 0);
  var totalNoVat = invoices.filter(function (i) { return i.isNoVat; }).reduce(function (s, i) { return s + i.totalBeforeVat; }, 0);

  var html = '<div class="stat-row">' +
    '<div class="stat-card"><b>' + invoices.length + '</b>จำนวน invoice</div>' +
    '<div class="stat-card"><b>' + totalVat.toLocaleString('en-US', { maximumFractionDigits: 2 }) + '</b>ยอดก่อน Vat (VAT invoices)</div>' +
    '<div class="stat-card"><b>' + totalNoVat.toLocaleString('en-US', { maximumFractionDigits: 2 }) + '</b>ยอดก่อน Vat (no-VAT invoices)</div>' +
    (state.unmappedRows && state.unmappedRows.length ? '<div class="stat-card"><b>' + state.unmappedRows.length + '</b>แถวที่ไม่พบใน Cost→Income (ตัดออกแล้ว)</div>' : '') +
    '</div>';

  if (state.detailParsed) {
    var pending = findPossiblyPending(invoices, state.detailParsed);
    if (pending.length) {
      html += '<div class="warn-box">⚠️ พบ ' + pending.length + ' invoice ที่ไม่พบ "' + esc(state.groupingKeyCol) + '" ในไฟล์ Detail of Invoice เลย (อาจเป็น Pending) — กรุณาตรวจสอบก่อนกรอกไฟล์จริง: ' +
        pending.slice(0, 20).map(function (i) { return esc(i.groupKey); }).join(', ') + (pending.length > 20 ? ' ...' : '') + '</div>';
    }
  } else {
    html += '<div class="warn-box">⚠️ ไม่มีข้อมูล Detail of Invoice สำหรับเช็ค Pending/reconcile — ระบบจะกรอกทุก invoice ที่คำนวณได้โดยไม่ตัดใครออก</div>';
  }

  html += '<table class="preview-table"><thead><tr><th>Key</th><th>Ref No.</th><th>ชื่อ</th><th>ประเภท</th><th class="num">ยอดก่อน Vat</th></tr></thead><tbody>';
  invoices.forEach(function (inv) {
    html += '<tr><td>' + esc(inv.groupKey) + '</td><td>' + esc(inv.refNo) + '</td><td>' + esc(inv.name) + '</td><td>' +
      (inv.isNoVat ? 'ไม่มี Vat' : 'มี Vat') + '</td><td class="num">' + inv.totalBeforeVat.toLocaleString('en-US', { maximumFractionDigits: 2 }) + '</td></tr>';
  });
  html += '</tbody></table>';
  body.innerHTML = html;
}

// ── Step 6: run + verify + download ─────────────────────────────────────────────────────────────

$('pwEye').addEventListener('click', function () {
  var input = $('pwInput');
  input.type = input.type === 'password' ? 'text' : 'password';
});

$('btnRun').addEventListener('click', async function () {
  var msgEl = $('statusMsg');
  $('btnRun').disabled = true;
  try {
    msgEl.textContent = '⏳ กำลังกรอกไฟล์...';
    var headPath = state.templateWb.sheets['Head'];
    var linePath = state.templateWb.sheets['Line'];
    var headXmlOrig = await getSheetXml(state.templateWb, 'Head');
    var lineXmlOrig = await getSheetXml(state.templateWb, 'Line');

    var result = fillHeadLine(headXmlOrig, lineXmlOrig, state.invoices, state.lookups, state.templateSst);

    var headCheck = verifyWrites(result.headXml, result.headWrites);
    var lineCheck = verifyWrites(result.lineXml, result.lineWrites);
    if (!headCheck.ok || !lineCheck.ok) {
      msgEl.innerHTML = '<div class="warn-box err">❌ ตรวจสอบพบข้อผิดพลาดก่อนสร้างไฟล์ — ไม่สร้างไฟล์ให้ กรุณาแจ้งผู้ดูแล<br>' +
        JSON.stringify(headCheck.mismatches.concat(lineCheck.mismatches)).slice(0, 500) + '</div>';
      $('btnRun').disabled = false;
      return;
    }

    var enc = new TextEncoder();
    var zipFiles = [];
    var names = Object.keys(state.templateWb.entries);
    for (var i = 0; i < names.length; i++) {
      var name = names[i];
      var entry = state.templateWb.entries[name];
      if (name === headPath) { zipFiles.push({ name: name, data: enc.encode(result.headXml) }); continue; }
      if (name === linePath) { zipFiles.push({ name: name, data: enc.encode(result.lineXml) }); continue; }
      var bytes = await decompressEntryBytes(entry, state.templateBuf);
      zipFiles.push({ name: name, data: bytes });
    }
    zipFiles = await applyRecalcFixes(zipFiles, enc);
    var zipBytes = buildZip(zipFiles);

    var finalBytes = zipBytes;
    var pw = $('pwInput').value;
    if (pw) {
      msgEl.textContent = '⏳ กำลังล็อกรหัสผ่าน...';
      var wbPop = await XlsxPopulate.fromDataAsync(zipBytes);
      var blob = await wbPop.outputAsync({ password: pw });
      finalBytes = new Uint8Array(await blob.arrayBuffer());
    }

    state.resultBytes = finalBytes;
    var dlBlob = new Blob([finalBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var dlUrl = URL.createObjectURL(dlBlob);
    var dlBtn = $('btnDownload');
    dlBtn.href = dlUrl;
    dlBtn.download = 'DTW_filled_' + Date.now() + '.xlsx';
    dlBtn.classList.add('show');

    var warnHtml = result.warnings.length
      ? '<div class="warn-box">⚠️ มีจุดที่ควรตรวจสอบก่อนใช้ไฟล์จริง (' + result.warnings.length + ' รายการ):<br>' + result.warnings.map(esc).join('<br>') + '</div>'
      : '';
    msgEl.innerHTML = '<div class="warn-box" style="background:#e8f5e9;border-color:#a5d6a7;color:#1b5e20">✅ กรอกไฟล์สำเร็จ ตรวจสอบความถูกต้องแล้ว (verifyWrites ผ่านทุกจุด)</div>' + warnHtml;
  } catch (err) {
    msgEl.innerHTML = '<div class="warn-box err">❌ ' + esc(err.message) + '</div>';
  } finally {
    $('btnRun').disabled = false;
  }
});

function esc(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
