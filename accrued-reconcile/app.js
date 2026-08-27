// ══════════════════════════════════════════════════════
//  Data for Reverse Accrued — merges multiple JE export CSVs into one "Reverse Income" CSV
//  (net debit/credit reversed and summed per project+account+ref3+memo), then compares the
//  resulting income total per project against the "Accrue Income" / "Accrue Minus Income" sheets
//  read directly from the uploaded Accrued (all).xlsx to flag any project whose numbers don't
//  match. XLSX reading reuses this portal's shared zip/xlsx primitives (excel.js).
// ══════════════════════════════════════════════════════

var GL_INCOME_SET = new Set([
  '41120102', '41120102A', '41120103', '41120104', '41120105', '41120106', '41120107', '41120108',
  '41120109', '41120110', '41120111', '41120112', '41120113', '41120114', '41120115', '41120116',
  '41120117', '41120118', '41120119', '41120120', '41120121', '41120122', '41120101',
  '21710101A', '21710103A', '21710102'
]);

var state = { csvFiles: [], xlsxFile: null, mergedRows: null, mergedHeader1: null, mergedHeader2: null, tableData: null, fileJE: null };

function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function setStatus(msg, cls) {
  var el = document.getElementById('statusBox');
  el.textContent = msg || '';
  el.className = 'status-box' + (cls ? ' ' + cls : '');
  el.style.display = msg ? 'block' : 'none';
}

// ── CSV parsing (handles quoted fields with embedded commas/quotes) ──────────────────────────
function parseCsvLine(line) {
  var out = [], cur = '', inQuotes = false;
  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (inQuotes) {
      if (ch === '"') { if (line[i + 1] === '"') { cur += '"'; i++; } else { inQuotes = false; } }
      else cur += ch;
    } else {
      if (ch === '"') inQuotes = true;
      else if (ch === ',') { out.push(cur); cur = ''; }
      else cur += ch;
    }
  }
  out.push(cur);
  return out;
}
function parseCsv(text) {
  return text.split(/\r\n|\n|\r/).filter(function (l) { return l.length > 0; }).map(parseCsvLine);
}
function readFileAsText(file) {
  return new Promise(function (resolve, reject) {
    var r = new FileReader();
    r.onload = function () { resolve(r.result); };
    r.onerror = reject;
    r.readAsText(file);
  });
}
function round6(n) { return Math.round(n * 1e6) / 1e6; }
function numCell(v) {
  if (v === undefined || v === null || v === '') return 0;
  var n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}
function formatMoney(n) { return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

// ── Merge JE export CSVs: reverse debit/credit, net by project+account+ref3+memo ─────────────
// Row shape (post-header): batchnum, oldLineId, account, debit, credit, linememo, project,
// projcode, ref3. linememo ends in "_MMYY" (or similar) -- extracted as the reversed memo's month
// tag ("Reverse Income_MMYY"). Once the Accrued (all).xlsx is loaded, a JE ref (read from that
// workbook's "JE Accrue" column) is appended: "Reverse Income_MMYY (JExxxxxxx)" — one JE per
// uploaded CSV file (1 file = 1 JE), not per project, since a single JE can span several
// projects. Each output row's `fileIdx` (tracked via rowFileIdx / fileProjects below) records
// which uploaded file it came from so the caller can resolve that file's JE and tag it.
async function mergeCsvFiles(files) {
  var header1 = null, header2 = null;
  var net = new Map();
  var orderedKeys = [];
  var fileProjects = files.map(function () { return []; });

  for (var fi = 0; fi < files.length; fi++) {
    var text = await readFileAsText(files[fi]);
    var rows = parseCsv(text);
    if (rows.length < 3) continue;
    if (!header1) { header1 = rows[0]; header2 = rows[1]; }
    var dataRows = rows.slice(2);
    for (var ri = 0; ri < dataRows.length; ri++) {
      var row = dataRows[ri];
      if (row.every(function (c) { return c.trim() === ''; })) continue;
      while (row.length < 9) row.push('');
      var account = row[2], debitStr = row[3], creditStr = row[4], linememo = row[5], project = row[6], projcode = row[7], ref3 = row[8];
      if (project && fileProjects[fi].indexOf(project) === -1) fileProjects[fi].push(project);
      var m = /_(\d{4})\s*$/.exec(linememo.trim());
      var monthCode = m ? m[1] : linememo.trim().split('_').pop();
      var newMemo = 'Reverse Income_' + monthCode;
      var d = debitStr.trim() ? parseFloat(debitStr) : 0;
      var c = creditStr.trim() ? parseFloat(creditStr) : 0;
      var newDebit = c, newCredit = d;
      var key = [project, projcode, account, ref3, newMemo].join('|||');
      if (!net.has(key)) { net.set(key, { debit: 0, credit: 0, project: project, projcode: projcode, account: account, ref3: ref3, memo: newMemo, fileIdx: fi }); orderedKeys.push(key); }
      var entry = net.get(key);
      entry.debit += newDebit;
      entry.credit += newCredit;
    }
  }

  var outRows = [];
  var outFileIdx = [];
  var lineId = 0;
  orderedKeys.forEach(function (key) {
    var e = net.get(key);
    var netAmt = e.debit - e.credit;
    lineId++;
    var debitVal = '', creditVal = '';
    if (netAmt >= 0) debitVal = round6(netAmt); else creditVal = round6(-netAmt);
    outRows.push(['1', String(lineId), e.account, debitVal, creditVal, e.memo, e.project, e.projcode, e.ref3]);
    outFileIdx.push(e.fileIdx);
  });

  return { header1: header1, header2: header2, rows: outRows, rowFileIdx: outFileIdx, fileProjects: fileProjects };
}

function sumIncomeByProjectFromCsv(mergedRows) {
  var byProject = {};
  mergedRows.forEach(function (r) {
    var account = r[2], debitStr = r[3], creditStr = r[4], project = r[6];
    if (!GL_INCOME_SET.has(account)) return;
    var d = debitStr !== '' ? parseFloat(debitStr) : 0;
    var c = creditStr !== '' ? parseFloat(creditStr) : 0;
    byProject[project] = (byProject[project] || 0) + (c - d);
  });
  Object.keys(byProject).forEach(function (k) { byProject[k] = -byProject[k]; });
  return byProject;
}

// Reads the "Accrue Income" / "Accrue Minus Income" sheet directly out of the uploaded Accrued
// (all).xlsx: rows 1-3 are header (title/label/GL code), real data starts row 4 (aoa index 3).
// Column B (index 1) = Project, columns E-AD (index 4-29) = the income GL-code columns to sum.
function sumIncomeFromSheetGrid(aoa) {
  var byProject = {};
  if (!aoa) return byProject;
  for (var r = 3; r < aoa.length; r++) {
    var row = aoa[r];
    if (!row) continue;
    var proj = row[1];
    if (proj == null || String(proj).trim() === '') continue;
    if (proj === 'Row Labels' || proj === 'Grand Total') continue;
    var s = 0;
    for (var c = 4; c <= 29; c++) s += numCell(row[c]);
    byProject[proj] = (byProject[proj] || 0) + s;
  }
  return byProject;
}

// Finds the column index of a header cell (searched across the 3 header rows, aoa index 0-2)
// whose trimmed text matches `label` case-insensitively. Returns -1 if not found — used so the
// "JE Accrue" lookup below never hardcodes a column position.
function findHeaderCol(aoa, label) {
  if (!aoa) return -1;
  var target = label.trim().toLowerCase();
  for (var r = 0; r < 3 && r < aoa.length; r++) {
    var row = aoa[r];
    if (!row) continue;
    for (var c = 0; c < row.length; c++) {
      if (row[c] != null && String(row[c]).trim().toLowerCase() === target) return c;
    }
  }
  return -1;
}

// Reads the "JE Accrue" column (found dynamically by header text, see findHeaderCol) out of the
// same Accrue Income / Accrue Minus Income sheet grid, and returns { project: "JExxxxxxx" } —
// one JE ref per project (first non-empty value found wins). Returns {} if the column isn't
// present in this sheet, so callers degrade gracefully instead of erroring.
function sumJEByProjectFromSheetGrid(aoa) {
  var byProject = {};
  if (!aoa) return byProject;
  var jeCol = findHeaderCol(aoa, 'JE Accrue');
  if (jeCol === -1) return byProject;
  for (var r = 3; r < aoa.length; r++) {
    var row = aoa[r];
    if (!row) continue;
    var proj = row[1];
    if (proj == null || String(proj).trim() === '') continue;
    if (proj === 'Row Labels' || proj === 'Grand Total') continue;
    if (byProject[proj]) continue;
    var v = row[jeCol];
    if (v == null || String(v).trim() === '') continue;
    var s = String(v).trim();
    byProject[proj] = /^je/i.test(s) ? s : ('JE' + s);
  }
  return byProject;
}

// ── UI wiring ──────────────────────────────────────────────────────────────────────────────
function wireDropzone(dzId, inputId, onFiles) {
  var dz = document.getElementById(dzId);
  var input = document.getElementById(inputId);
  dz.addEventListener('click', function (e) { if (e.target.tagName !== 'INPUT') input.click(); });
  input.addEventListener('change', function () { if (input.files.length) onFiles(input.files); });
  ['dragover', 'dragenter'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('drag'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('drag'); }); });
  dz.addEventListener('drop', function (e) { if (e.dataTransfer.files.length) onFiles(e.dataTransfer.files); });
}

function renderCsvFileList() {
  var el = document.getElementById('csvFileList');
  if (!state.csvFiles.length) { el.innerHTML = ''; document.getElementById('csvCountLabel').textContent = 'เลือกได้หลายไฟล์'; return; }
  document.getElementById('csvCountLabel').textContent = 'เลือกแล้ว ' + state.csvFiles.length + ' ไฟล์';
  el.innerHTML = state.csvFiles.map(function (f, i) {
    var jeTag = '';
    if (state.fileJE && state.fileJE.length === state.csvFiles.length) {
      jeTag = state.fileJE[i]
        ? '<span style="margin-left:8px;color:#0a7a3d;font-weight:600;">→ ' + esc_(state.fileJE[i]) + '</span>'
        : '<span style="margin-left:8px;color:#999;">→ ไม่พบ JE</span>';
    }
    return '<div class="file-row"><span class="idx">' + (i + 1) + '</span><span class="fname">' + esc_(f.name) + '</span>' + jeTag +
      '<button class="rm" data-i="' + i + '" title="ลบไฟล์">✕</button></div>';
  }).join('');
  Array.prototype.forEach.call(el.querySelectorAll('.rm'), function (btn) {
    btn.addEventListener('click', function () {
      state.csvFiles.splice(parseInt(btn.getAttribute('data-i'), 10), 1);
      state.fileJE = null;
      renderCsvFileList();
    });
  });
}

wireDropzone('dropCsv', 'csvInput', function (files) {
  state.csvFiles = state.csvFiles.concat(Array.prototype.slice.call(files).filter(function (f) { return /\.csv$/i.test(f.name); }));
  state.fileJE = null;
  renderCsvFileList();
});
wireDropzone('dropXlsx', 'xlsxInput', function (files) {
  state.xlsxFile = files[0];
  document.getElementById('xlsxLabel').textContent = state.xlsxFile.name;
});

document.getElementById('btnRun').addEventListener('click', runReconcile);
document.getElementById('btnDownload').addEventListener('click', downloadMergedCsv);
document.getElementById('searchBox').addEventListener('input', function () { renderTable(this.value); });

async function runReconcile() {
  if (!state.csvFiles.length) { setStatus('กรุณาเลือกไฟล์ JE (.csv) อย่างน้อย 1 ไฟล์', 'err'); return; }
  if (!state.xlsxFile) { setStatus('กรุณาเลือกไฟล์ Accrued (all).xlsx', 'err'); return; }

  setStatus('⏳ กำลังประมวลผล...', 'info');
  document.getElementById('resultSection').style.display = 'none';

  try {
    var merged = await mergeCsvFiles(state.csvFiles);
    state.mergedRows = merged.rows;
    state.mergedHeader1 = merged.header1;
    state.mergedHeader2 = merged.header2;

    var revByProject = sumIncomeByProjectFromCsv(state.mergedRows);

    var xlsxBuf = await state.xlsxFile.arrayBuffer();
    var wb = await loadWorkbook(xlsxBuf);
    if (!wb.sheets['Accrue Income'] && !wb.sheets['Accrue Minus Income']) {
      throw new Error('ไม่พบชีต "Accrue Income" หรือ "Accrue Minus Income" ในไฟล์นี้ — ตรวจสอบว่าเป็นไฟล์ Accrued (all).xlsx ที่ถูกต้อง');
    }
    var allByProject = {};
    var jeByProject = {};
    if (wb.sheets['Accrue Income']) {
      var incomeGrid = (await readSheetGrid(wb, 'Accrue Income')).aoa;
      var s1 = sumIncomeFromSheetGrid(incomeGrid);
      Object.keys(s1).forEach(function (k) { allByProject[k] = (allByProject[k] || 0) + s1[k]; });
      var je1 = sumJEByProjectFromSheetGrid(incomeGrid);
      Object.keys(je1).forEach(function (k) { if (!jeByProject[k]) jeByProject[k] = je1[k]; });
    }
    if (wb.sheets['Accrue Minus Income']) {
      var minusGrid = (await readSheetGrid(wb, 'Accrue Minus Income')).aoa;
      var s2 = sumIncomeFromSheetGrid(minusGrid);
      Object.keys(s2).forEach(function (k) { allByProject[k] = (allByProject[k] || 0) + s2[k]; });
      var je2 = sumJEByProjectFromSheetGrid(minusGrid);
      Object.keys(je2).forEach(function (k) { if (!jeByProject[k]) jeByProject[k] = je2[k]; });
    }

    // Resolve one JE ref per uploaded CSV file (1 file = 1 JE): take the first non-empty
    // jeByProject match among the projects that appear in that file. Every merged row inherits
    // its originating file's JE ref, e.g. "Reverse Income_0326" -> "Reverse Income_0326
    // (JE1234567)". Files whose projects have no JE match keep the plain memo.
    var fileJE = merged.fileProjects.map(function (projList) {
      for (var i = 0; i < projList.length; i++) {
        if (jeByProject[projList[i]]) return jeByProject[projList[i]];
      }
      return null;
    });
    state.fileJE = fileJE;
    state.mergedRows.forEach(function (row, idx) {
      var je = fileJE[merged.rowFileIdx[idx]];
      if (je) row[5] = row[5] + ' (' + je + ')';
    });
    renderCsvFileList();

    var allProjects = Array.from(new Set(Object.keys(revByProject).concat(Object.keys(allByProject)))).sort();
    var tableData = [];
    var matchCount = 0, totalRev = 0, totalAll = 0;
    allProjects.forEach(function (p) {
      var rv = revByProject[p] || 0;
      var av = allByProject[p] || 0;
      var diff = Math.round((rv - av) * 100) / 100;
      if (diff === 0) matchCount++;
      totalRev += rv;
      totalAll += av;
      tableData.push({ project: p, rev: rv, all: av, diff: diff });
    });
    state.tableData = tableData;

    document.getElementById('emptyHint').style.display = 'none';
    document.getElementById('resultSection').style.display = 'block';
    document.getElementById('statRow').innerHTML =
      '<div class="stat-box"><div class="stat-num c-blue">' + matchCount + '/' + allProjects.length + '</div><div class="stat-label">โปรเจคที่ตรงกัน</div></div>' +
      '<div class="stat-box"><div class="stat-num c-green">' + formatMoney(totalRev) + '</div><div class="stat-label">ยอดรวม Reverse Income (CSV)</div></div>' +
      '<div class="stat-box"><div class="stat-num c-orange">' + formatMoney(totalAll) + '</div><div class="stat-label">ยอดรวม Accrued (all)</div></div>';

    document.getElementById('searchBox').value = '';
    renderTable('');
    setStatus('✅ เสร็จสิ้น: อัปโหลดไฟล์ JE ทั้งหมด ' + state.csvFiles.length + ' ไฟล์ รวมได้ ' + state.mergedRows.length + ' บรรทัด เทียบกับ ' + state.xlsxFile.name, 'ok');
  } catch (err) {
    setStatus('❌ เกิดข้อผิดพลาด: ' + err.message, 'err');
  }
}

function renderTable(filterText) {
  var data = state.tableData || [];
  var q = (filterText || '').trim().toLowerCase();
  var filtered = q ? data.filter(function (r) { return r.project.toLowerCase().indexOf(q) !== -1; }) : data;

  var rowsHtml = filtered.map(function (r) {
    var mismatch = r.diff !== 0;
    return '<tr' + (mismatch ? ' class="row-mismatch"' : '') + '>' +
      '<td>' + esc_(r.project) + '</td>' +
      '<td>' + formatMoney(r.rev) + '</td>' +
      '<td>' + formatMoney(r.all) + '</td>' +
      '<td>' + (mismatch ? '<span class="badge b-mismatch">' + formatMoney(r.diff) + '</span>' : '<span class="badge b-match">ตรงกัน</span>') + '</td>' +
      '</tr>';
  }).join('');

  document.getElementById('tableWrap').innerHTML =
    '<div class="detail-table-wrap"><table class="detail-table"><thead><tr>' +
    '<th>Project</th><th>Reverse Income (CSV)</th><th>Accrued (all)</th><th>Diff</th>' +
    '</tr></thead><tbody>' + rowsHtml + '</tbody></table></div>' +
    '<div class="result-footer">แสดง ' + filtered.length + ' จาก ' + data.length + ' โปรเจค</div>';
}

function downloadMergedCsv() {
  if (!state.mergedRows) return;
  var lines = [state.mergedHeader1.join(','), state.mergedHeader2.join(',')];
  state.mergedRows.forEach(function (row) {
    lines.push(row.map(function (v) {
      var s = String(v);
      return s.indexOf(',') !== -1 ? '"' + s + '"' : s;
    }).join(','));
  });
  var csvContent = lines.join('\r\n');
  var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = 'Reverse Income Merged.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(function () { URL.revokeObjectURL(url); }, 3000);
}
