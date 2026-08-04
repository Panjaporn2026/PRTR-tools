// ══════════════════════════════════════════════════════
//  Fee Table Rate Check — client-side port of the fee-table-rate-check Claude Skill's
//  compare_fee_table.py. Compares a per-project client "Fee table" workbook against the
//  "★ MASTER Service Fee table -- ALL ★.xlsx" reference over columns S through MQ, for the row(s)
//  matching that project. Read-only against both files -- this tool never writes back into either
//  uploaded workbook, it only builds a brand-new summary .xlsx.
// ══════════════════════════════════════════════════════

var PROJECT_COL = 2;   // C (0-based) -- project name
var AGREEMENT_COL = 3; // D -- Agreement No.
var CONDITION_COL = 7; // H -- Condition details ("PRTR - SERVICE FEE" / "CLIENT - SERVICE FEE")
var GROUP_ROW = 1;     // row 2 (0-based) -- category label
var PAYCODE_ROW = 2;   // row 3 -- SAP pay code
var DETAIL_ROW = 3;    // row 4 -- short label
var MASTER_SHEET_NAME = '2. Fee table - Ipop 21.07';
var FIRST_COL = 'S', LAST_COL = 'MQ';

var state = { masterWb: null, clientWb: null, masterFile: null, clientFile: null, pendingPairing: null };

function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function normText(s) { return String(s == null ? '' : s).trim(); }
function setStatus(msg, cls) {
  var el = document.getElementById('statusBox');
  el.textContent = msg || '';
  el.className = 'status-box' + (cls ? ' ' + cls : '');
  el.style.display = msg ? 'block' : 'none';
}
function isBlank(v) { return v === null || v === undefined || String(v).trim() === ''; }

function cellVal(ws, r, c) {
  var addr = XLSX.utils.encode_cell({ r: r, c: c });
  var cell = ws[addr];
  return cell ? (cell.v === undefined ? null : cell.v) : null;
}

// ── Sheet selection ─────────────────────────────────────────────────────────────────────────
// Client files use varying sheet names (Sheet1, "2. Fee table - Ipop 21.07", etc.) -- there's
// normally exactly one sheet with data, so just take the first one.
function findClientSheet(wb) { return wb.Sheets[wb.SheetNames[0]]; }

// The Master workbook has 30+ tabs (working scratch sheets, per-person copies, old versions...).
// The one with real project rows is named MASTER_SHEET_NAME as of this writing; if that ever
// changes, fall back to whichever sheet has the most populated Agreement No. (column D) cells.
function findMasterSheet(wb) {
  if (wb.SheetNames.indexOf(MASTER_SHEET_NAME) !== -1) return { ws: wb.Sheets[MASTER_SHEET_NAME], name: MASTER_SHEET_NAME };
  var best = null, bestCount = -1, bestName = null;
  wb.SheetNames.forEach(function (name) {
    var ws = wb.Sheets[name];
    var ref = ws['!ref'];
    if (!ref) return;
    var range = XLSX.utils.decode_range(ref);
    var count = 0;
    for (var r = range.s.r; r <= range.e.r; r++) { if (!isBlank(cellVal(ws, r, AGREEMENT_COL))) count++; }
    if (count > bestCount) { best = ws; bestCount = count; bestName = name; }
  });
  return { ws: best, name: bestName };
}

function sheetRowCount(ws) {
  var ref = ws['!ref'];
  return ref ? XLSX.utils.decode_range(ref).e.r + 1 : 0;
}

// Client fee tables put their data row(s) starting at row 5 (0-based r=4). Return every row from
// there onward that has a project name in column C.
function findClientRows(ws) {
  var maxR = sheetRowCount(ws);
  var rows = [];
  for (var r = 4; r < maxR; r++) { if (!isBlank(cellVal(ws, r, PROJECT_COL))) rows.push(r); }
  return rows;
}

// Search the master sheet for rows matching the project name (substring, case-insensitive).
// Prefer an exact Agreement No. match; fall back to name-only matches so the caller can present
// them for the user to pick from rather than guessing.
function findMasterRows(ws, project, agreement) {
  var maxR = sheetRowCount(ws);
  var nameMatches = [], exactMatches = [];
  var projU = normText(project).toUpperCase();
  for (var r = 0; r < maxR; r++) {
    var c3 = cellVal(ws, r, PROJECT_COL);
    var c4 = cellVal(ws, r, AGREEMENT_COL);
    if (c3 && projU && normText(c3).toUpperCase().indexOf(projU) !== -1) {
      var entry = { row: r, project: c3, agreement: c4 };
      nameMatches.push(entry);
      if (agreement && c4 && normText(c4) === normText(agreement)) exactMatches.push(entry);
    }
  }
  return { exact: exactMatches, nameOnly: nameMatches };
}

// Some projects split into two data rows -- a "PRTR - SERVICE FEE" row and a "CLIENT - SERVICE
// FEE" row -- both sharing the same project name and agreement number. Pair client rows to master
// candidates using column H ("Condition details") first; fall back to matching by position
// (both lists top-to-bottom) when that label doesn't disambiguate cleanly.
function pairRows(wsClient, wsMaster, clientRows, masterCandidates) {
  if (clientRows.length === 1 && masterCandidates.length >= 1) {
    return [[clientRows[0], masterCandidates[0].row]];
  }
  var sortedClient = clientRows.slice().sort(function (a, b) { return a - b; });
  var usedMaster = {};
  var pairs = [];
  sortedClient.forEach(function (crow) {
    var ccond = cellVal(wsClient, crow, CONDITION_COL);
    var match = null;
    if (!isBlank(ccond)) {
      for (var i = 0; i < masterCandidates.length; i++) {
        var mrow = masterCandidates[i].row;
        if (usedMaster[mrow]) continue;
        var mcond = cellVal(wsMaster, mrow, CONDITION_COL);
        if (!isBlank(mcond) && normText(mcond).toUpperCase() === normText(ccond).toUpperCase()) { match = mrow; break; }
      }
    }
    if (match === null) {
      var remaining = masterCandidates.map(function (m) { return m.row; }).filter(function (m) { return !usedMaster[m]; }).sort(function (a, b) { return a - b; });
      var idx = sortedClient.indexOf(crow);
      if (idx < remaining.length) match = remaining[idx];
    }
    if (match !== null) { usedMaster[match] = true; pairs.push([crow, match]); }
  });
  return pairs;
}

function fmtVal(v) {
  if (isBlank(v)) return '(ว่าง)';
  if (typeof v === 'number') return (v * 100).toFixed(2).replace(/\.00$/, '') + '%';
  return String(v);
}

// Diffs every column from FIRST_COL to LAST_COL between the matched client/master rows. Header
// info (group/paycode/detail) can live on either sheet -- some client files leave a column's
// header blank even though the sheet nominally extends that far (a template never filled out that
// wide); fall back to the Master's header so the summary still reads sensibly, and flag it.
function diffColumns(wsClient, wsMaster, clientRow, masterRow) {
  var s = XLSX.utils.decode_col(FIRST_COL), e = XLSX.utils.decode_col(LAST_COL);
  var diffs = [];
  for (var c = s; c <= e; c++) {
    var vc = cellVal(wsClient, clientRow, c);
    var vm = cellVal(wsMaster, masterRow, c);
    if (vc === vm) continue;
    if (isBlank(vc) && isBlank(vm)) continue; // both empty in different flavors (null vs '') -- not a real diff
    var group = cellVal(wsClient, GROUP_ROW, c); if (isBlank(group)) group = cellVal(wsMaster, GROUP_ROW, c);
    var paycode = cellVal(wsClient, PAYCODE_ROW, c); if (isBlank(paycode)) paycode = cellVal(wsMaster, PAYCODE_ROW, c);
    var detail = cellVal(wsClient, DETAIL_ROW, c); if (isBlank(detail)) detail = cellVal(wsMaster, DETAIL_ROW, c);
    var note = '';
    if (isBlank(cellVal(wsClient, DETAIL_ROW, c))) note = 'ไฟล์ client ไม่มีคอลัมน์นี้ในหัวตาราง (header ว่าง)';
    diffs.push({
      col: XLSX.utils.encode_col(c), group: group, paycode: paycode, detail: detail,
      clientVal: vc, masterVal: vm, note: note
    });
  }
  return { diffs: diffs, totalChecked: e - s + 1 };
}

// ── Dropzone wiring ─────────────────────────────────────────────────────────────────────────
function wireDropzone(dzId, inputId, onFile) {
  var dz = document.getElementById(dzId), input = document.getElementById(inputId);
  dz.addEventListener('click', function (e) { if (e.target.tagName !== 'INPUT') input.click(); });
  input.addEventListener('change', function () { if (input.files.length) onFile(input.files[0]); });
  ['dragover', 'dragenter'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('drag'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('drag'); }); });
  dz.addEventListener('drop', function (e) { if (e.dataTransfer.files.length) onFile(e.dataTransfer.files[0]); });
}

function maybeEnableRun() {
  document.getElementById('btnRun').disabled = !(state.masterWb && state.clientWb);
}

wireDropzone('dropMaster', 'masterInput', function (file) { loadWorkbookFile(file, 'master'); });
wireDropzone('dropClient', 'clientInput', function (file) { loadWorkbookFile(file, 'client'); });

async function loadWorkbookFile(file, kind) {
  setStatus('⏳ กำลังอ่านไฟล์ ' + file.name + ' ...', 'info');
  try {
    var buf = await file.arrayBuffer();
    var wb = XLSX.read(buf, { type: 'array', cellDates: true });
    if (kind === 'master') {
      state.masterWb = wb; state.masterFile = file;
      document.getElementById('dropMaster').classList.add('done');
      document.getElementById('masterLabel').textContent = '✓ ' + file.name;
    } else {
      state.clientWb = wb; state.clientFile = file;
      document.getElementById('dropClient').classList.add('done');
      document.getElementById('clientLabel').textContent = '✓ ' + file.name;
    }
    setStatus('✅ อ่านไฟล์ ' + file.name + ' สำเร็จ', 'ok');
    maybeEnableRun();
  } catch (err) {
    setStatus('❌ อ่านไฟล์ ' + file.name + ' ไม่สำเร็จ: ' + err.message, 'err');
  }
}

document.getElementById('btnRun').addEventListener('click', runComparison);
document.getElementById('btnConfirmCandidate').addEventListener('click', confirmCandidateSelection);

function runComparison() {
  document.getElementById('candidateCard').style.display = 'none';
  document.getElementById('resultCard').style.display = 'none';
  try {
    var wsClient = findClientSheet(state.clientWb);
    var masterSheetInfo = findMasterSheet(state.masterWb);
    var wsMaster = masterSheetInfo.ws;
    if (!wsMaster) throw new Error('ไม่พบชีตข้อมูลใน Master (ไม่พบชีต "' + MASTER_SHEET_NAME + '" และไม่มีชีตอื่นที่มีคอลัมน์ Agreement No.)');

    var clientRows = findClientRows(wsClient);
    if (!clientRows.length) throw new Error('ไม่พบแถวข้อมูลในไฟล์ client (คาดว่าจะมีชื่อโปรเจคในคอลัมน์ C ตั้งแต่แถว 5 เป็นต้นไป)');

    var project = cellVal(wsClient, clientRows[0], PROJECT_COL);
    var agreement = cellVal(wsClient, clientRows[0], AGREEMENT_COL);
    var found = findMasterRows(wsMaster, project, agreement);
    var candidates = found.exact.length ? found.exact : found.nameOnly;
    if (!candidates.length) {
      throw new Error('ไม่พบแถวใน Master ที่ตรงกับโปรเจค "' + project + '" (Agreement No. "' + (agreement || '-') + '") — กรุณาตรวจสอบชื่อโปรเจคในไฟล์ client');
    }

    var pairs = pairRows(wsClient, wsMaster, clientRows, candidates);
    if (pairs.length !== clientRows.length || candidates.length > 1 && found.exact.length === 0) {
      // Ambiguous (multiple name-only candidates with different agreements, or pairing came up
      // short) -- let the user pick the right master row(s) rather than guessing.
      showCandidatePicker(wsClient, wsMaster, clientRows, candidates, project, agreement);
      return;
    }

    finishComparison(wsClient, wsMaster, pairs, masterSheetInfo.name);
  } catch (err) {
    setStatus('❌ ' + err.message, 'err');
  }
}

function showCandidatePicker(wsClient, wsMaster, clientRows, candidates, project, agreement) {
  state.pendingPairing = { wsClient: wsClient, wsMaster: wsMaster, clientRows: clientRows };
  var box = document.getElementById('candidateBody');
  var html = '<p style="margin-bottom:10px;">พบหลายแถวใน Master ที่ชื่อโปรเจคตรงกับ "' + esc_(project) +
    '" แต่เลข Agreement No. ไม่ตรงกับไฟล์ client ("' + esc_(agreement || '-') + '") — กรุณาเลือกแถวที่ถูกต้อง' +
    (clientRows.length > 1 ? ' (เลือกให้ครบ ' + clientRows.length + ' แถว ตามจำนวนแถวในไฟล์ client)' : '') + '</p>';
  clientRows.forEach(function (crow, i) {
    var ccond = cellVal(wsClient, crow, CONDITION_COL);
    html += '<div class="candidate-card"><b>แถว client #' + (i + 1) + '</b> (Condition: ' + esc_(ccond || '-') + ')<br>';
    candidates.forEach(function (cand, ci) {
      var mcond = cellVal(wsMaster, cand.row, CONDITION_COL);
      html += '<label style="margin-top:4px;"><input type="radio" name="cand_' + i + '" value="' + cand.row + '"' + (ci === 0 ? ' checked' : '') + '> ' +
        'แถว Master ' + (cand.row + 1) + ' — ' + esc_(cand.agreement || '-') + ' (Condition: ' + esc_(mcond || '-') + ')</label><br>';
    });
    html += '</div>';
  });
  box.innerHTML = html;
  document.getElementById('candidateCard').style.display = 'block';
  document.getElementById('mainEmptyHint').style.display = 'none';
  setStatus('⚠️ ต้องเลือกแถว Master ให้ถูกต้องก่อน', 'warn');
}

function confirmCandidateSelection() {
  var p = state.pendingPairing;
  if (!p) return;
  var pairs = [];
  p.clientRows.forEach(function (crow, i) {
    var sel = document.querySelector('input[name="cand_' + i + '"]:checked');
    if (sel) pairs.push([crow, parseInt(sel.value, 10)]);
  });
  if (!pairs.length) { setStatus('⚠️ กรุณาเลือกอย่างน้อย 1 แถว', 'warn'); return; }
  document.getElementById('candidateCard').style.display = 'none';
  var masterSheetInfo = findMasterSheet(state.masterWb);
  finishComparison(p.wsClient, p.wsMaster, pairs, masterSheetInfo.name);
}

var lastResults = null;

function finishComparison(wsClient, wsMaster, pairs, masterSheetName) {
  var results = pairs.map(function (pair) {
    var crow = pair[0], mrow = pair[1];
    var project = cellVal(wsClient, crow, PROJECT_COL);
    var agreement = cellVal(wsClient, crow, AGREEMENT_COL);
    var d = diffColumns(wsClient, wsMaster, crow, mrow);
    return { clientRow: crow, masterRow: mrow, project: project, agreement: agreement, diffs: d.diffs, totalChecked: d.totalChecked };
  });
  lastResults = results;
  renderResult(results, masterSheetName);
  setStatus('✅ เปรียบเทียบเสร็จแล้ว', 'ok');
}

function renderResult(results, masterSheetName) {
  var totalDiffs = results.reduce(function (s, r) { return s + r.diffs.length; }, 0);
  var totalChecked = results.reduce(function (s, r) { return s + r.totalChecked; }, 0);
  var html = '<h2>📄 ผลการเปรียบเทียบ</h2>';
  html += '<div class="stat-row">' +
    '<div class="stat-box"><div class="stat-num c-blue">' + results.length + '</div><div class="stat-label">แถวที่เทียบ</div></div>' +
    '<div class="stat-box"><div class="stat-num c-green">' + totalChecked + '</div><div class="stat-label">คอลัมน์ที่ตรวจสอบ (รวม)</div></div>' +
    '<div class="stat-box"><div class="stat-num ' + (totalDiffs ? 'c-red' : 'c-green') + '">' + totalDiffs + '</div><div class="stat-label">รายการที่ต่าง/เพิ่มจาก Master</div></div>' +
    '</div>';
  html += '<div class="result-footer">Master sheet ที่ใช้: "' + esc_(masterSheetName) + '"</div>';

  results.forEach(function (r, i) {
    html += '<div class="row-block">';
    html += '<div class="row-block-title">แถว ' + (i + 1) + ': ' + esc_(r.project) + ' — Agreement No.: ' + esc_(r.agreement || '-') +
      ' (client แถว ' + (r.clientRow + 1) + ' ↔ master แถว ' + (r.masterRow + 1) + ') — ' + r.diffs.length + ' รายการต่าง</div>';
    if (r.diffs.length) {
      html += '<div class="detail-table-wrap"><table class="detail-table"><thead><tr>' +
        '<th>คอลัมน์</th><th>หมวดหมู่</th><th>Pay Code</th><th>รายละเอียด</th><th>ค่าในไฟล์ Client</th><th>ค่าใน Master</th><th>สถานะ</th></tr></thead><tbody>';
      r.diffs.forEach(function (d) {
        var status, badgeClass;
        if (d.note) { status = d.note; badgeClass = 'b-note'; }
        else if (isBlank(d.masterVal)) { status = 'เพิ่มข้อมูลใหม่ (Master ไม่มี)'; badgeClass = 'b-new'; }
        else { status = 'ค่าต่างจาก Master'; badgeClass = 'b-diff'; }
        html += '<tr><td>' + esc_(d.col) + '</td><td>' + esc_(d.group || '-') + '</td><td>' + esc_(d.paycode || '-') + '</td><td>' + esc_(d.detail || '-') + '</td>' +
          '<td class="num">' + esc_(fmtVal(d.clientVal)) + '</td><td class="num">' + esc_(fmtVal(d.masterVal)) + '</td>' +
          '<td><span class="badge ' + badgeClass + '">' + esc_(status) + '</span></td></tr>';
      });
      html += '</tbody></table></div>';
    } else {
      html += '<div class="status-box ok" style="display:block">✅ ไม่พบความแตกต่าง — ตรงกับ Master ทั้งหมด</div>';
    }
    html += '</div>';
  });

  html += '<button class="btn-outline" id="btnDownloadSummary">⬇ ดาวน์โหลดไฟล์สรุป Excel</button>';

  var card = document.getElementById('resultCard');
  card.innerHTML = html;
  card.style.display = 'block';
  document.getElementById('mainEmptyHint').style.display = 'none';
  document.getElementById('btnDownloadSummary').addEventListener('click', downloadSummary);
}

// ── Excel summary export (ExcelJS -- brand-new workbook, no template to preserve) ──────────────
async function downloadSummary() {
  var btn = document.getElementById('btnDownloadSummary');
  btn.disabled = true; btn.textContent = '⏳ กำลังสร้างไฟล์...';
  try {
    var wb = new ExcelJS.Workbook();
    var headerFont = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    var headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    var normalFont = { name: 'Arial', size: 10 };
    var diffFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
    var thin = { style: 'thin', color: { argb: 'FFB7B7B7' } };
    var border = { top: thin, left: thin, right: thin, bottom: thin };

    lastResults.forEach(function (r, i) {
      var ws = wb.addWorksheet(('Summary_row' + (r.clientRow + 1)).slice(0, 31));
      ws.getCell('A1').value = 'สรุปการเปรียบเทียบ Rate ค่าบริการ (คอลัมน์ ' + FIRST_COL + ' ถึง ' + LAST_COL + ')';
      ws.getCell('A1').font = { name: 'Arial', size: 14, bold: true };
      ws.getCell('A2').value = 'โปรเจค: ' + r.project + ' | Agreement No.: ' + (r.agreement || '-');
      ws.getCell('A2').font = { name: 'Arial', size: 10, italic: true };
      ws.getCell('A3').value = 'พบข้อมูลที่แตกต่าง/เพิ่มเติมจาก Master ทั้งหมด ' + r.diffs.length + ' รายการ (จากทั้งหมด ' + r.totalChecked + ' ช่องที่ตรวจสอบ)';
      ws.getCell('A3').font = { name: 'Arial', size: 10, bold: true };

      var headers = ['คอลัมน์', 'หมวดหมู่ (แถว 2)', 'Pay Code (แถว 3)', 'รายละเอียด (แถว 4)', 'ค่าในไฟล์ Client', 'ค่าใน Master', 'สถานะ / หมายเหตุ'];
      var headerRow = ws.getRow(5);
      headers.forEach(function (h, ci) {
        var cell = headerRow.getCell(ci + 1);
        cell.value = h; cell.font = headerFont; cell.fill = headerFill; cell.border = border;
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      });

      var rIdx = 6;
      r.diffs.forEach(function (d) {
        var status;
        if (d.note) status = d.note;
        else if (isBlank(d.masterVal)) status = 'เพิ่มข้อมูลใหม่ (Master ไม่มี)';
        else status = 'ค่าต่างจาก Master';
        var vals = [d.col, d.group || '', d.paycode || '', d.detail || '', fmtVal(d.clientVal), fmtVal(d.masterVal), status];
        var row = ws.getRow(rIdx);
        vals.forEach(function (v, ci) {
          var cell = row.getCell(ci + 1);
          cell.value = v; cell.font = normalFont; cell.border = border; cell.fill = diffFill;
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        });
        rIdx++;
      });

      var widths = [10, 26, 14, 16, 20, 14, 30];
      widths.forEach(function (w, ci) { ws.getColumn(ci + 1).width = w; });
      ws.views = [{ state: 'frozen', ySplit: 5 }];
    });

    var buf = await wb.xlsx.writeBuffer();
    var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = 'fee_table_comparison_summary.xlsx';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 3000);
    btn.textContent = '✅ ดาวน์โหลดแล้ว';
    setTimeout(function () { btn.disabled = false; btn.textContent = '⬇ ดาวน์โหลดไฟล์สรุป Excel'; }, 2500);
  } catch (err) {
    btn.disabled = false; btn.textContent = '⬇ ดาวน์โหลดไฟล์สรุป Excel';
    alert('เกิดข้อผิดพลาด: ' + err.message);
  }
}
