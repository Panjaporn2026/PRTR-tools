// ══════════════════════════════════════════════════════
//  UI wiring for the Health Insurance Reconcile tool.
//  See formula.js / reconcile.js / employee.js / verify.js / excel.js for the actual logic --
//  this file is state + rendering + the top-level pipeline that calls into them.
// ══════════════════════════════════════════════════════

var state = {
  reconcileFile: null,
  reconcileWb: null,
  reconcileGrid: null, // { aoa, rowXmlByNum, sheetXml }
  existingIds: null, // Set
  invoiceFiles: [], // { file, wb, grid, monthLabel, altIdCol, bundled: {status,col,header,sum,...}, sums, ackColumn }
  monthsInScope: {}, // monthLabel -> bool
  candidates: [], // new-employee candidates (from employee.js)
  confirmedNew: {}, // altId -> { templateRow, ... }
  planCodeCols: null,
};

function dbg(msg) {
  var el = document.getElementById('dbgBox');
  if (!el) return;
  el.style.display = 'block';
  el.textContent += '[' + new Date().toLocaleTimeString() + '] ' + msg + '\n';
  el.scrollTop = 1e9;
  console.log(msg);
}
function setStatus(msg, cls) {
  var el = document.getElementById('statusBox');
  el.textContent = msg || '';
  el.className = 'status-box' + (cls ? ' ' + cls : '');
  el.style.display = msg ? 'block' : 'none';
}

// ── Reconcile file ──────────────────────────────────────
document.getElementById('reconcileInput').addEventListener('change', function (e) {
  if (e.target.files[0]) loadReconcileFile(e.target.files[0]);
});
var reconcileDrop = document.getElementById('reconcileDrop');
reconcileDrop.addEventListener('dragover', function (e) { e.preventDefault(); reconcileDrop.classList.add('drag'); });
reconcileDrop.addEventListener('dragleave', function () { reconcileDrop.classList.remove('drag'); });
reconcileDrop.addEventListener('drop', function (e) {
  e.preventDefault(); reconcileDrop.classList.remove('drag');
  if (e.dataTransfer.files[0]) loadReconcileFile(e.dataTransfer.files[0]);
});

async function loadReconcileFile(file) {
  try {
    setStatus('⏳ กำลังอ่านไฟล์ Reconcile...', 'info');
    var buf = await file.arrayBuffer();
    var wb = await loadWorkbook(buf);
    var grid = await readSheetGrid(wb, 'Detail Reconcile');
    var altIdCol = colLettersToIndex(RECONCILE_ALT_ID_COL);
    var ids = new Set();
    for (var r = RECONCILE_HEADER_ROW; r < grid.aoa.length; r++) {
      var v = (grid.aoa[r] || [])[altIdCol];
      if (v != null && String(v).trim() !== '') ids.add(normAltId(v));
    }
    state.reconcileFile = file;
    state.reconcileWb = wb;
    state.reconcileGrid = grid;
    state.existingIds = ids;

    // Plan-code reference sheets, for Step 4 validation.
    var tl = await readSheetGrid(wb, 'TL2025');
    var pru = await readSheetGrid(wb, 'PRU2024');
    state.tl2025Keys = (tl.aoa || []).slice(6).map(function (r) { return r && r[0] != null ? String(r[0]) : null; }).filter(Boolean);
    state.pru2024Keys = (pru.aoa || []).slice(2).map(function (r) { return r && r[0] != null ? String(r[0]) : null; }).filter(Boolean);

    document.getElementById('reconcileDrop').classList.add('done');
    document.getElementById('reconcileDrop').querySelector('.dz-hint').textContent = '✅ ' + file.name + ' (' + ids.size + ' คนในไฟล์แล้ว)';
    setStatus('✅ โหลดไฟล์ Reconcile สำเร็จ — พบพนักงาน ' + ids.size + ' คน', 'ok');
    dbg('Reconcile loaded: ' + ids.size + ' existing Alternate IDs, sheets: ' + Object.keys(wb.sheets).join(', '));
    render();
  } catch (err) {
    setStatus('❌ อ่านไฟล์ Reconcile ไม่ได้: ' + err.message, 'err');
    dbg('ERROR loading reconcile: ' + err.message);
  }
}

// ── Invoice files ────────────────────────────────────────
document.getElementById('invoiceInput').addEventListener('change', function (e) {
  if (e.target.files.length) Array.from(e.target.files).forEach(loadInvoiceFile);
});
var invoiceDrop = document.getElementById('invoiceDrop');
invoiceDrop.addEventListener('dragover', function (e) { e.preventDefault(); invoiceDrop.classList.add('drag'); });
invoiceDrop.addEventListener('dragleave', function () { invoiceDrop.classList.remove('drag'); });
invoiceDrop.addEventListener('drop', function (e) {
  e.preventDefault(); invoiceDrop.classList.remove('drag');
  Array.from(e.dataTransfer.files).forEach(loadInvoiceFile);
});

async function loadInvoiceFile(file) {
  try {
    setStatus('⏳ กำลังอ่าน ' + file.name + '...', 'info');
    var buf = await file.arrayBuffer();
    var wb = await loadWorkbook(buf);
    var grid = await readSheetGrid(wb, 'Detail of Invoice');
    var altIdCol = findAlternateIdColumn(grid.aoa, INVOICE_HEADER_ROW - 1);
    if (altIdCol < 0) throw new Error('ไม่พบคอลัมน์ Alternate ID ในแถวที่ ' + INVOICE_HEADER_ROW);
    var bundled = findBundledHealthColumn(grid.aoa, INVOICE_HEADER_ROW - 1, altIdCol);
    var monthLabel = detectInvoiceMonth(grid.aoa) || file.name;

    var entry = {
      file: file, wb: wb, grid: grid, altIdCol: altIdCol, monthLabel: monthLabel,
      bundled: bundled, acknowledged: false, sums: null,
    };
    state.invoiceFiles.push(entry);
    if (state.monthsInScope[monthLabel] === undefined) state.monthsInScope[monthLabel] = true;
    dbg('Invoice loaded: ' + file.name + ' -> month ' + monthLabel + ', bundled-col status: ' + bundled.status);
    setStatus('✅ โหลด ' + file.name + ' สำเร็จ', 'ok');
    render();
  } catch (err) {
    setStatus('❌ อ่าน ' + file.name + ' ไม่ได้: ' + err.message, 'err');
    dbg('ERROR loading invoice ' + file.name + ': ' + err.message);
  }
}

// Invoice month: look for a date value in the first ~16 rows / first 8 cols (confirmed pattern:
// row 13 holds the invoice month date in the real files) rather than trusting the filename alone.
function detectInvoiceMonth(aoa) {
  for (var r = 0; r < Math.min(16, aoa.length); r++) {
    var row = aoa[r] || [];
    for (var c = 0; c < Math.min(8, row.length); c++) {
      var v = row[c];
      if (typeof v === 'number' && v > 40000 && v < 60000) return excelSerialToMonthKey(v);
    }
  }
  return null;
}

function pickBundledColumn(fileIdx, col) {
  var entry = state.invoiceFiles[fileIdx];
  var cand = entry.bundled.candidates.find(function (c) { return c.col === col; });
  entry.bundled = useBundledColumn(entry.grid.aoa, INVOICE_HEADER_ROW - 1, entry.altIdCol, cand.col, cand.header);
  render();
}
function acknowledgeColumn(fileIdx) {
  state.invoiceFiles[fileIdx].acknowledged = true;
  render();
}
function removeInvoiceFile(fileIdx) {
  state.invoiceFiles.splice(fileIdx, 1);
  render();
}

// ── Render ───────────────────────────────────────────────
function render() {
  renderInvoiceList();
  renderMonthChecklist();
  renderActionRow();
}

function renderInvoiceList() {
  var el = document.getElementById('invoiceList');
  el.innerHTML = state.invoiceFiles.map(function (entry, i) {
    var b = entry.bundled;
    var body = '';
    if (b.status === 'ok') {
      body = '<div class="col-card ' + (entry.acknowledged ? 'ok' : 'warn') + '">'
        + '<div><b>คอลัมน์ที่จับคู่:</b> ' + colIndexToLetters(b.col) + ' — "' + esc(b.header) + '"</div>'
        + '<div>ผลรวม: ' + fmt(b.sum) + ' (ตัวอย่าง: ' + b.sampleValues.map(fmt).join(', ') + ')</div>'
        + (entry.acknowledged ? '' : '<button class="btn-ack" onclick="acknowledgeColumn(' + i + ')">✓ ยืนยันว่าคอลัมน์นี้ถูกต้อง</button>')
        + '</div>';
    } else {
      body = '<div class="col-card err"><b>' + (b.status === 'not_found' ? '⚠ ไม่พบคอลัมน์ที่ตรงเงื่อนไข' : '⚠ พบหลายคอลัมน์ที่เข้าเงื่อนไข — กรุณาเลือกเอง') + '</b>'
        + (b.candidates || []).map(function (c) {
          return '<div class="col-candidate"><label><input type="radio" name="bcol' + i + '" onclick="pickBundledColumn(' + i + ',' + c.col + ')">'
            + colIndexToLetters(c.col) + ' — "' + esc(c.header) + '" (ผลรวม ' + fmt(c.sum) + ', ตัวอย่าง ' + c.sampleValues.map(fmt).join(', ') + ')</label></div>';
        }).join('')
        + '</div>';
    }
    var dupWarn = '';
    if (b.status === 'ok') {
      var s = sumAmountsByAltId(entry.grid.aoa, INVOICE_HEADER_ROW - 1, entry.altIdCol, b.col);
      entry.sums = s;
      if (s.duplicates.length) {
        dupWarn = '<div class="dup-warn">⚠ พบรหัสพนักงานซ้ำ ' + s.duplicates.length + ' คนในเดือนนี้ (รวมยอดแบบมีเครื่องหมายบวก/ลบให้แล้ว): '
          + s.duplicates.map(function (d) { return d.altId + ' (' + d.rows.length + ' แถว)'; }).join(', ') + '</div>';
      }
    }
    return '<div class="invoice-card">'
      + '<div class="invoice-card-top"><b>' + esc(entry.file.name) + '</b> <span class="month-tag">' + esc(entry.monthLabel) + '</span>'
      + '<button class="row-remove" onclick="removeInvoiceFile(' + i + ')">✕</button></div>'
      + body + dupWarn + '</div>';
  }).join('') || '<div class="empty-hint">ยังไม่มีไฟล์ invoice</div>';
}

function renderMonthChecklist() {
  var el = document.getElementById('monthChecklist');
  var months = Object.keys(state.monthsInScope);
  el.innerHTML = months.map(function (m) {
    return '<label class="month-cb"><input type="checkbox" ' + (state.monthsInScope[m] ? 'checked' : '')
      + ' onchange="toggleMonth(\'' + m + '\',this.checked)"> ' + esc(m) + '</label>';
  }).join('') || '<span class="muted">อัปโหลดไฟล์ invoice ก่อนเพื่อเลือกเดือน</span>';
}
function toggleMonth(m, checked) { state.monthsInScope[m] = checked; render(); }

function renderActionRow() {
  var ready = state.reconcileGrid && state.invoiceFiles.length
    && state.invoiceFiles.every(function (f) { return f.bundled.status === 'ok' && f.acknowledged; });
  document.getElementById('btnStart').disabled = !ready;
}

// ── Pipeline: Steps 1-3 (existing employees) then Step 4 (candidates) ──────────────────────
async function startProcess() {
  setStatus('⏳ กำลังประมวลผล...', 'info');
  var inScope = state.invoiceFiles.filter(function (f) { return state.monthsInScope[f.monthLabel]; });
  var monthlyAmounts = inScope.map(function (f) {
    return { monthLabel: f.monthLabel, amounts: f.sums.amounts };
  });

  // Existing employees: build the CR:CW write-list.
  var monthCols = findInvoiceMonthColumns(state.reconcileGrid.aoa); // 0-based col indices, CR..CW
  var monthColByLabel = mapMonthColumnsToLabels(state.reconcileGrid.aoa, monthCols);
  state.writePlan = [];
  for (var r = RECONCILE_HEADER_ROW; r < state.reconcileGrid.aoa.length; r++) {
    var row = state.reconcileGrid.aoa[r];
    if (!row) continue;
    var altId = normAltId(row[colLettersToIndex(RECONCILE_ALT_ID_COL)]);
    if (!altId) continue;
    inScope.forEach(function (f) {
      var col = monthColByLabel[f.monthLabel];
      if (col === undefined) return; // this reconcile file has no column for this month yet
      var amt = f.sums.amounts[altId] || 0;
      state.writePlan.push({ rowNum: r + 1, col: col, altId: altId, monthLabel: f.monthLabel, amount: amt });
    });
  }

  // New-employee candidates.
  var candidates = findNewEmployeeCandidates(monthlyAmounts, state.existingIds, inScope.map(function (f) {
    return { identityByAltId: buildIdentitySnapshot(f) };
  }));
  flagNearMissIds(candidates, state.existingIds);
  state.candidates = candidates;

  renderWritePreview();
  renderCandidates();
  document.getElementById('reviewCard').style.display = 'block';
  setStatus('✅ วิเคราะห์เสร็จแล้ว — ตรวจสอบผลด้านล่างก่อนดาวน์โหลด', 'ok');
}

function mapMonthColumnsToLabels(aoa, monthCols) {
  var headerRow = aoa[RECONCILE_HEADER_ROW - 1] || [];
  var map = {};
  monthCols.forEach(function (col) {
    var v = headerRow[col];
    if (typeof v === 'number') map[excelSerialToMonthKey(v)] = col;
  });
  return map;
}

function buildIdentitySnapshot(invoiceEntry) {
  var aoa = invoiceEntry.grid.aoa;
  var header = aoa[INVOICE_HEADER_ROW - 1] || [];
  function findCol(kw) {
    for (var c = 0; c < header.length; c++) {
      if (normHeader(header[c]).indexOf(kw) >= 0) return c;
    }
    return -1;
  }
  var companyCol = 3; // confirmed constant cell area (row 12 D12 in real file) -- read separately
  var nameCol = findCol('NAME (THAI)') >= 0 ? findCol('NAME (THAI)') : findCol('THAI');
  var posCol = findCol('POSITION');
  var empIdCol = findCol('EMP ID') >= 0 ? findCol('EMP ID') : findCol('รหัสพนักงาน');
  var out = {};
  for (var r = INVOICE_HEADER_ROW; r < aoa.length; r++) {
    var row = aoa[r];
    if (!row) continue;
    var altId = normAltId(row[invoiceEntry.altIdCol]);
    if (!altId) continue;
    out[altId] = {
      company: (aoa[11] || [])[3] || '', // D12 in the real file (0-based row 11, col 3)
      altId: row[invoiceEntry.altIdCol],
      empId: empIdCol >= 0 ? row[empIdCol] : '',
      thaiName: nameCol >= 0 ? row[nameCol] : '',
      position: posCol >= 0 ? row[posCol] : '',
    };
  }
  return out;
}

function renderWritePreview() {
  var el = document.getElementById('writePreview');
  var byEmployee = {};
  state.writePlan.forEach(function (w) {
    (byEmployee[w.altId] || (byEmployee[w.altId] = [])).push(w);
  });
  var rows = Object.keys(byEmployee).map(function (altId) {
    var cells = byEmployee[altId].map(function (w) {
      return '<span class="mo-amt">' + esc(w.monthLabel) + ': ' + fmt(w.amount) + '</span>';
    }).join(' ');
    return '<tr><td>' + esc(altId) + '</td><td>' + cells + '</td></tr>';
  }).join('');
  el.innerHTML = '<table><thead><tr><th>Alternate ID</th><th>จำนวนเงินที่จะกรอกต่อเดือน</th></tr></thead><tbody>' + rows + '</tbody></table>';
}

function renderCandidates() {
  var el = document.getElementById('candidateList');
  if (!state.candidates.length) { el.innerHTML = '<div class="empty-hint">ไม่พบพนักงานใหม่ในช่วงเดือนที่เลือก</div>'; return; }
  var templateOptions = buildTemplateRowOptions();
  el.innerHTML = state.candidates.map(function (cand, i) {
    var nearMissWarn = cand.nearMiss
      ? '<div class="nearmiss-warn">⚠ อาจเป็นรหัสเดิมที่มีอยู่แล้ว: ' + esc(cand.nearMiss) + ' — ตรวจสอบก่อนเพิ่มเป็นคนใหม่</div>' : '';
    var months = cand.months.map(function (m) { return esc(m.monthLabel) + ': ' + fmt(m.amount); }).join(', ');
    return '<div class="candidate-card">'
      + '<label class="cand-check"><input type="checkbox" onchange="toggleCandidate(' + i + ',this.checked)"> '
      + '<b>' + esc(cand.altId) + '</b>' + (cand.identity ? ' — ' + esc(cand.identity.thaiName || '') : '') + '</label>'
      + '<div class="bundle-label">ยอดที่เจอ (รวม Life/Accident/TPD — ไม่ใช่ Health อย่างเดียว): ' + months + '</div>'
      + nearMissWarn
      + '<div class="tpl-picker" id="tplPicker' + i + '" style="display:none">'
      + 'ใช้แถวต้นแบบ: <select onchange="setCandidateTemplate(' + i + ',this.value)">' + templateOptions + '</select>'
      + '<div id="tplStatus' + i + '"></div></div>'
      + '</div>';
  }).join('');
}

function buildTemplateRowOptions() {
  var opts = ['<option value="">-- เลือก --</option>'];
  for (var r = RECONCILE_HEADER_ROW; r < state.reconcileGrid.aoa.length; r++) {
    var row = state.reconcileGrid.aoa[r];
    if (!row) continue;
    var altId = row[colLettersToIndex(RECONCILE_ALT_ID_COL)];
    var name = row[colLettersToIndex('E')];
    if (!altId) continue;
    opts.push('<option value="' + (r + 1) + '">' + esc((r + 1) + ': ' + altId + ' - ' + (name || '')) + '</option>');
  }
  return opts.join('');
}

function toggleCandidate(i, checked) {
  var cand = state.candidates[i];
  document.getElementById('tplPicker' + i).style.display = checked ? 'block' : 'none';
  if (!checked) delete state.confirmedNew[cand.altId];
}

function setCandidateTemplate(i, rowNumStr) {
  var cand = state.candidates[i];
  var statusEl = document.getElementById('tplStatus' + i);
  if (!rowNumStr) { delete state.confirmedNew[cand.altId]; statusEl.innerHTML = ''; return; }
  var templateRow = parseInt(rowNumStr, 10);
  var rowXml = state.reconcileGrid.rowXmlByNum[templateRow];
  var sharedMap = buildSharedFormulaMap(state.reconcileGrid.sheetXml);
  var cells = parseRowCells(rowXml, sharedMap);

  var errors = templateRowErrors(cells);
  var planCodeCols = { F: getCellDisplayValue(cells, 'F', state.reconcileGrid.aoa[templateRow - 1]), H: getCellDisplayValue(cells, 'H', state.reconcileGrid.aoa[templateRow - 1]) };
  var planProblems = validatePlanCodes(
    { F: state.reconcileGrid.aoa[templateRow - 1][colLettersToIndex('F')], H: state.reconcileGrid.aoa[templateRow - 1][colLettersToIndex('H')] },
    state.tl2025Keys, state.pru2024Keys
  );

  var msgs = [];
  if (errors.length) msgs.push('<div class="tpl-error">⚠ แถวต้นแบบนี้มีสูตรที่ error อยู่แล้ว (' + errors.map(function (e) { return e.col; }).join(', ') + ') — ไม่ควรใช้เป็นต้นแบบ</div>');
  if (planProblems.length) msgs.push('<div class="tpl-error">⚠ รหัสแผนของแถวต้นแบบไม่พบใน sheet อ้างอิง: ' + planProblems.map(function (p) { return p.col + '=' + p.code + ' (' + p.sheet + ')'; }).join(', ') + '</div>');
  statusEl.innerHTML = msgs.join('') || '<div class="tpl-ok">✅ แถวต้นแบบใช้ได้</div>';

  if (!errors.length && !planProblems.length) {
    state.confirmedNew[cand.altId] = { candidateIdx: i, templateRow: templateRow, cells: cells, planCodeCols: {
      F: state.reconcileGrid.aoa[templateRow - 1][colLettersToIndex('F')],
      G: state.reconcileGrid.aoa[templateRow - 1][colLettersToIndex('G')],
      H: state.reconcileGrid.aoa[templateRow - 1][colLettersToIndex('H')],
    } };
  } else {
    delete state.confirmedNew[cand.altId];
  }
}

function getCellDisplayValue(cells, col, aoaRow) {
  return aoaRow ? aoaRow[colLettersToIndex(col)] : (cells[col] ? cells[col].value : null);
}

// ── helpers ──────────────────────────────────────────────
function colIndexToLetters(idx) {
  var s = '', n = idx + 1;
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function fmt(n) { return (typeof n === 'number') ? n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : String(n); }

// ── Generate + download ─────────────────────────────────
function cellXmlFor(col, rowNum, spec) {
  var sAttr = spec.style ? ' s="' + spec.style + '"' : '';
  if (spec.blank) return '<c r="' + col + rowNum + '"' + sAttr + '/>';
  if (spec.formula !== undefined && spec.formula !== null) {
    return '<c r="' + col + rowNum + '"' + sAttr + '><f>' + esc(spec.formula) + '</f></c>'; // no <v> -- see formula.js
  }
  if (spec.override) {
    if (typeof spec.override.literal === 'number') {
      return '<c r="' + col + rowNum + '"' + sAttr + '><v>' + spec.override.literal + '</v></c>';
    }
    return '<c r="' + col + rowNum + '"' + sAttr + ' t="inlineStr"><is><t>' + esc(spec.override.literal) + '</t></is></c>';
  }
  if (spec.literal !== undefined) {
    if (typeof spec.literal === 'number') return '<c r="' + col + rowNum + '"' + sAttr + '><v>' + spec.literal + '</v></c>';
    return '<c r="' + col + rowNum + '"' + sAttr + ' t="inlineStr"><is><t>' + esc(spec.literal) + '</t></is></c>';
  }
  return '<c r="' + col + rowNum + '"' + sAttr + '/>';
}

// Replace (or insert, in correct column order) one cell within a row's raw inner XML.
function setCellInRowXml(rowInnerXml, col, rowNum, newCellXml) {
  var addr = col + rowNum;
  var existingRe = new RegExp('<c\\s+r="' + addr + '"[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)');
  if (existingRe.test(rowInnerXml)) return rowInnerXml.replace(existingRe, newCellXml);
  var targetIdx = colLettersToIndex(col);
  var cellRe = /<c\s+r="([A-Z]+)\d+"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g, m, insertAt = rowInnerXml.length;
  while ((m = cellRe.exec(rowInnerXml)) !== null) {
    if (colLettersToIndex(m[1]) > targetIdx) { insertAt = m.index; break; }
  }
  return rowInnerXml.slice(0, insertAt) + newCellXml + rowInnerXml.slice(insertAt);
}

async function generateOutput() {
  try {
    setStatus('⏳ กำลังสร้างไฟล์...', 'info');
    var enc = new TextEncoder();
    var sheetXml = state.reconcileGrid.sheetXml;
    var sdStart = sheetXml.indexOf('<sheetData');
    var sdTagEnd = sheetXml.indexOf('>', sdStart) + 1;
    var sdEnd = sheetXml.indexOf('</sheetData>');
    var pre = sheetXml.slice(0, sdTagEnd);
    var sdBody = sheetXml.slice(sdTagEnd, sdEnd);
    var post = sheetXml.slice(sdEnd);

    var verifyWritesList = [];

    // 1) Existing-employee Invoice fills.
    var byRow = {};
    state.writePlan.forEach(function (w) { (byRow[w.rowNum] || (byRow[w.rowNum] = [])).push(w); });
    Object.keys(byRow).forEach(function (rowNumStr) {
      var rowNum = parseInt(rowNumStr, 10);
      var rowRe = new RegExp('(<row\\b[^>]*\\br="' + rowNum + '"[^>]*>)([\\s\\S]*?)(<\\/row>)');
      var m = rowRe.exec(sdBody);
      if (!m) return;
      var inner = m[2];
      byRow[rowNum].forEach(function (w) {
        var col = colIndexToLetters(w.col);
        inner = setCellInRowXml(inner, col, rowNum, '<c r="' + col + rowNum + '"><v>' + w.amount + '</v></c>');
        verifyWritesList.push({ addr: col + rowNum, expected: w.amount });
      });
      sdBody = sdBody.slice(0, m.index) + m[1] + inner + m[3] + sdBody.slice(m.index + m[0].length);
    });

    // 2) New-employee rows, appended after the last existing row.
    var maxRow = 0;
    var rowNumRe = /<row\b[^>]*\br="(\d+)"/g, rm;
    while ((rm = rowNumRe.exec(sdBody)) !== null) maxRow = Math.max(maxRow, parseInt(rm[1], 10));

    var monthCols = findInvoiceMonthColumns(state.reconcileGrid.aoa);
    var monthColByLabel = mapMonthColumnsToLabels(state.reconcileGrid.aoa, monthCols);
    var inScope = state.invoiceFiles.filter(function (f) { return state.monthsInScope[f.monthLabel]; });

    var newRowsXml = [];
    Object.keys(state.confirmedNew).forEach(function (altId) {
      maxRow++;
      var conf = state.confirmedNew[altId];
      var cand = state.candidates[conf.candidateIdx];
      var amountsByCol = {};
      cand.months.forEach(function (mo) {
        var col = monthColByLabel[mo.monthLabel];
        if (col !== undefined) amountsByCol[colIndexToLetters(col)] = mo.amount;
      });
      var overrides = buildNewRowOverrides(
        cand.identity || {}, conf.planCodeCols,
        monthCols.map(colIndexToLetters), amountsByCol
      );
      var cellSpecs = buildCopiedRowCells(conf.cells, conf.templateRow, maxRow, overrides);
      var cellsXml = Object.keys(cellSpecs).sort(function (a, b) { return colLettersToIndex(a) - colLettersToIndex(b); })
        .map(function (col) { return cellXmlFor(col, maxRow, cellSpecs[col]); }).join('');
      newRowsXml.push('<row r="' + maxRow + '">' + cellsXml + '</row>');
      Object.keys(amountsByCol).forEach(function (col) {
        verifyWritesList.push({ addr: col + maxRow, expected: amountsByCol[col] });
      });
    });

    var newSheetXml = pre + sdBody + newRowsXml.join('') + post;

    var check = verifyWrites(newSheetXml, verifyWritesList);
    if (!check.ok) {
      dbg('VERIFY MISMATCH: ' + JSON.stringify(check.mismatches));
      setStatus('❌ ตรวจสอบพบข้อผิดพลาดก่อนสร้างไฟล์ — ดูรายละเอียดใน log ด้านล่าง กรุณาอย่าใช้ไฟล์นี้', 'err');
      return;
    }

    var sheetKey = state.reconcileWb.sheets['Detail Reconcile'];
    var names = Object.keys(state.reconcileWb.entries);
    var zipFiles = [];
    for (var i = 0; i < names.length; i++) {
      var nm = names[i];
      zipFiles.push({
        name: nm,
        data: nm === sheetKey ? enc.encode(newSheetXml) : await decompressEntryBytes(state.reconcileWb.entries[nm], state.reconcileWb.buf)
      });
    }
    zipFiles = await applyRecalcFixes(zipFiles, enc);
    var resultBytes = buildZip(zipFiles);

    var blob = new Blob([resultBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = state.reconcileFile.name.replace(/\.xlsx$/i, '') + '_updated.xlsx';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 3000);

    setStatus('✅ สร้างไฟล์สำเร็จ — ดาวน์โหลดแล้ว กรุณาเปิดด้วย Excel จริงและตรวจคอลัมน์ Actual Cost ของพนักงานใหม่ก่อนใช้งานจริง', 'ok');
    dbg('Generated output: ' + resultBytes.length + ' bytes, ' + newRowsXml.length + ' new row(s), ' + Object.keys(byRow).length + ' existing row(s) touched');
  } catch (err) {
    setStatus('❌ สร้างไฟล์ไม่สำเร็จ: ' + err.message, 'err');
    dbg('ERROR generating output: ' + err.message + '\n' + err.stack);
  }
}

render();
