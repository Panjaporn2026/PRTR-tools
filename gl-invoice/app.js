// ══════════════════════════════════════════════════════
//  UI wiring for the GL Invoice Processing tool. See sheetmodel.js / styles.js / functions.js for
//  the actual logic -- this file is state + rendering + the top-level pipeline that calls into them.
// ══════════════════════════════════════════════════════

var FUNCTIONS = [
  { id: 'sso750', label: '1. SSO PRTR 750', multi: false,
    desc: 'Grouping = E51110102 & Paycode = T2A3 หรือ TZ74\n• Amount > 750 → Amount ใหม่ = Amount − 750\n• Amount ≤ 750 → Amount ใหม่ = 0\nแถวอื่นๆ ทั้งหมด ไม่เปลี่ยน' },
  { id: 'ssoIntroduceBy', label: '2. SSO Introduce by', multi: false,
    desc: 'เงื่อนไขหลัก: Grouping = E51110102 & Paycode = T2A3 หรือ TZ74\n• Introduce By = PRTR → Amount ใหม่ = 0\n• Introduce By = CLNT → ไม่เปลี่ยน\nแถวที่ไม่ตรงเงื่อนไข ไม่เปลี่ยน' },
  { id: 'duplicate', label: '3. Duplicate', multi: false,
    desc: 'เพิ่ม/อัพเดทแถว EXPENSE ต่อท้ายทุกคน (unique NAME):\nPaycode Code = EXPENSE, Paycode Name = ค่าใช้จ่าย\nAccount = 51110122, Grouping = E51110122, Amount = (ว่าง)\nแถว EXPENSE ทุกแถว — ตัวอักษรสีแดง' },
  { id: 'sso750PlusDuplicate', label: '4. SSO PRTR 750 + Duplicate EXPENSE', multi: false,
    desc: 'ขั้นที่ 1: ปรับ Amount เหมือนฟังก์ชัน 1\nขั้นที่ 2: เพิ่ม/อัพเดทแถว EXPENSE เหมือนฟังก์ชัน 3 (สีแดง)' },
  { id: 'ssoIntroduceByPlusDuplicate', label: '5. SSO Introduce by + Duplicate EXPENSE', multi: false,
    desc: 'ขั้นที่ 1: ปรับ Amount เหมือนฟังก์ชัน 2\nขั้นที่ 2: เพิ่ม/อัพเดทแถว EXPENSE เหมือนฟังก์ชัน 3 (สีแดง)' },
  { id: 'removeSso', label: '6. Remove SSO', multi: false,
    desc: 'Grouping = E51110102 & Paycode = T2A3 หรือ TZ74\n• บรรทัดที่ตรงเงื่อนไข → ลบออกทั้งบรรทัด\n• บรรทัดอื่นทั้งหมด → คงเดิม' },
  { id: 'merge', label: '7. Merge', multi: true,
    desc: 'โยนไฟล์ที่ 1 แล้วโยนไฟล์ที่ 2 แล้วโยนไฟล์ที่ 3-6 ตามลำดับ\nระบบจะนำข้อมูล (หลัง header) ของไฟล์ที่ 2-6 ต่อท้ายไฟล์ที่ 1\nHeader ของ output ยึดตามไฟล์ที่ 1 ทั้งหมด รูปแบบของไฟล์ห้ามเปลี่ยนแปลง' },
  { id: 'changeHeader', label: '8. Change Header', multi: false,
    desc: 'ลบแถวที่ 4 และ 5 ออก + เปลี่ยนชื่อ Column:\n• แถว Header เลื่อนขึ้น 2 แถว\n• Period → Calendar Group\n• Paycode Code → PIN Name' }
];

var FN_META = {
  sso750: { icon: '💰', title: 'สรุปผล SSO PRTR 750' },
  ssoIntroduceBy: { icon: '💰', title: 'สรุปผล SSO Introduce by' },
  duplicate: { icon: '➕', title: 'สรุปผล Duplicate — แถว EXPENSE' },
  sso750PlusDuplicate: { icon: '💰➕', title: 'สรุปผล SSO PRTR 750 + Duplicate EXPENSE' },
  ssoIntroduceByPlusDuplicate: { icon: '💰➕', title: 'สรุปผล SSO Introduce by + Duplicate EXPENSE' },
  removeSso: { icon: '🗑️', title: 'สรุปผล Remove SSO — บรรทัดที่ถูกลบ' },
  merge: { icon: '🔗', title: 'สรุปผล Merge — รวมไฟล์' },
  changeHeader: { icon: '📝', title: 'สรุปผล Change Header' }
};

var state = { fnId: FUNCTIONS[0].id, files: [], resultBytes: null, resultBaseName: null, processedAt: null, sourceLabel: null };

function dbg(msg) { console.log(msg); }
function setStatus(msg, cls) {
  var el = document.getElementById('statusBox');
  el.textContent = msg || '';
  el.className = 'status-box' + (cls ? ' ' + cls : '');
  el.style.display = msg ? 'block' : 'none';
}
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function fmtAmt(v) { return (v == null || v === '') ? '-' : Number(v).toFixed(2); }

function currentFn() { return FUNCTIONS.find(function (f) { return f.id === state.fnId; }); }

function renderSidebar() {
  var el = document.getElementById('sidebar');
  el.innerHTML = FUNCTIONS.map(function (f) {
    return '<button class="fn-btn' + (f.id === state.fnId ? ' active' : '') + '" onclick="selectFunction(\'' + f.id + '\')">' + esc_(f.label) + '</button>';
  }).join('');
}
function renderFnDesc() {
  document.getElementById('fnDesc').textContent = currentFn().desc;
}
function renderDropzoneHint() {
  var fn = currentFn();
  document.getElementById('dzHint').textContent = fn.multi ? '.xlsx (เลือกได้หลายไฟล์ เรียงตามลำดับที่ต้องการ)' : '.xlsx';
  document.getElementById('fileInput').multiple = !!fn.multi;
}
function renderFileList() {
  var el = document.getElementById('fileList');
  var fn = currentFn();
  if (!fn.multi || !state.files.length) { el.innerHTML = ''; document.getElementById('btnProcess').style.display = 'none'; return; }
  el.innerHTML = state.files.map(function (f, i) {
    return '<div class="file-row"><span class="idx">' + (i + 1) + '</span><span class="fname">' + esc_(f.name) + '</span>' +
      '<button class="rm" onclick="removeFile(' + i + ')" title="ลบ">✕</button></div>';
  }).join('');
  var btn = document.getElementById('btnProcess');
  btn.style.display = 'inline-block';
  btn.disabled = state.files.length < 2;
}

function resetOutputUI() {
  state.resultBytes = null;
  state.resultBaseName = null;
  var box = document.getElementById('summaryBox');
  box.style.display = 'none';
  box.innerHTML = '';
  setStatus('', '');
}

function resetForNewFile() {
  state.files = [];
  resetOutputUI();
  renderFileList();
  document.getElementById('fileInput').value = '';
  var dz = document.getElementById('dropzone');
  if (dz) dz.style.display = '';
}

function selectFunction(id) {
  state.fnId = id;
  state.files = [];
  resetOutputUI();
  renderSidebar();
  renderFnDesc();
  renderDropzoneHint();
  renderFileList();
  document.getElementById('fileInput').value = '';
}
function removeFile(i) {
  state.files.splice(i, 1);
  renderFileList();
}

// ── Drop zone wiring ─────────────────────────────────────────────────────
var dropzone = document.getElementById('dropzone');
dropzone.addEventListener('dragover', function (e) { e.preventDefault(); dropzone.classList.add('drag'); });
dropzone.addEventListener('dragleave', function () { dropzone.classList.remove('drag'); });
dropzone.addEventListener('drop', function (e) {
  e.preventDefault(); dropzone.classList.remove('drag');
  handleIncomingFiles(e.dataTransfer.files);
});
document.getElementById('fileInput').addEventListener('change', function (e) {
  handleIncomingFiles(e.target.files);
  e.target.value = '';
});

async function handleIncomingFiles(fileList) {
  var files = Array.from(fileList).filter(function (f) { return /\.xlsx$|\.xlsm$/i.test(f.name); });
  if (!files.length) { setStatus('❌ รองรับเฉพาะไฟล์ .xlsx / .xlsm', 'err'); return; }
  var fn = currentFn();
  resetOutputUI();

  if (fn.multi) {
    state.files = state.files.concat(files);
    renderFileList();
    setStatus('✅ เพิ่มไฟล์แล้ว (' + state.files.length + ' ไฟล์) — เรียงลำดับถูกต้องหรือยัง? กด "เริ่มรวมไฟล์" เมื่อพร้อม', 'info');
    return;
  }

  state.files = [files[0]];
  await runProcessing();
}

document.getElementById('btnProcess').addEventListener('click', runProcessing);

async function runProcessing() {
  var fn = currentFn();
  try {
    setStatus('⏳ กำลังประมวลผล...', 'info');
    resetOutputUI();

    var result;
    if (fn.multi) {
      var bufs = [], names = state.files.map(function (f) { return f.name; });
      for (var i = 0; i < state.files.length; i++) bufs.push(await state.files[i].arrayBuffer());
      result = await runMergeFunction(bufs, names);
      state.resultBaseName = state.files[0].name.replace(/\.[^.]+$/, '');
      state.sourceLabel = state.files.map(function (f) { return f.name; }).join(', ');
    } else {
      var buf = await state.files[0].arrayBuffer();
      result = await runSingleFileFunction(fn.id, buf);
      state.resultBaseName = state.files[0].name.replace(/\.[^.]+$/, '');
      state.sourceLabel = state.files[0].name;
    }

    if (!result.ok) {
      renderErrorSummary(result.summary);
      setStatus('❌ พบข้อผิดพลาด — ดูรายละเอียดด้านล่าง ไม่ได้สร้างไฟล์ผลลัพธ์', 'err');
      return;
    }

    state.resultBytes = result.outputBytes;
    state.processedAt = new Date();
    renderSummary(fn, result.summary);
    setStatus('✅ ประมวลผลสำเร็จ — ตรวจสอบสรุปผลด้านล่างแล้วดาวน์โหลดไฟล์', 'ok');
    dbg('Processed function=' + fn.id + ' summary=' + JSON.stringify(result.summary));
  } catch (err) {
    setStatus('❌ ประมวลผลไม่สำเร็จ: ' + err.message, 'err');
    dbg('ERROR: ' + err.message + '\n' + (err.stack || ''));
  }
}

function renderErrorSummary(summary) {
  var box = document.getElementById('summaryBox');
  box.style.display = 'block';
  var lines = (summary && summary.errors) || [];
  box.innerHTML = lines.map(function (l) { return '<div class="err-line">⚠ ' + esc_(l) + '</div>'; }).join('');
}

// ── Result card building blocks ───────────────────────────────────────────
function statCardHtml(s) {
  return '<div class="stat-box"><div class="stat-num c-' + s.color + '">' + esc_(s.value) + '</div>' +
    '<div class="stat-label">' + esc_(s.label) + '</div></div>';
}
function statRowHtml(stats) {
  return '<div class="stat-row">' + stats.map(statCardHtml).join('') + '</div>';
}
function noteBadge(note) {
  var cls = 'b-neutral';
  if (/ลบ/.test(note)) cls = 'b-removed';
  else if (/เพิ่ม/.test(note)) cls = 'b-added';
  else if (/อัพเดท/.test(note)) cls = 'b-updated';
  return '<span class="badge ' + cls + '">' + esc_(note) + '</span>';
}
function detailTableHtml(columns, rows) {
  if (!rows || !rows.length) return '';
  var thead = '<tr>' + columns.map(function (c) { return '<th>' + esc_(c.label) + '</th>'; }).join('') + '</tr>';
  var tbody = rows.map(function (r) {
    var rowClass = r._rowClass ? ' class="' + r._rowClass + '"' : '';
    return '<tr' + rowClass + '>' + columns.map(function (c) { return '<td>' + c.render(r) + '</td>'; }).join('') + '</tr>';
  }).join('');
  return '<div class="detail-table-wrap"><table class="detail-table"><thead>' + thead + '</thead><tbody>' + tbody + '</tbody></table></div>';
}

var SSO_ADJUST_COLUMNS = [
  { label: 'แถวที่ (เดิม)', render: function (r) { return esc_(r.row); } },
  { label: 'Paycode', render: function (r) { return esc_(r.paycode); } },
  { label: 'Paycode Name', render: function (r) { return esc_(r.paycodeName); } },
  { label: 'Amount เดิม', render: function (r) { return esc_(fmtAmt(r.amountBefore)); } },
  { label: 'Amount ใหม่', render: function (r) { return esc_(fmtAmt(r.amountAfter)); } },
  { label: 'หมายเหตุ', render: function (r) { return noteBadge(r.note); } }
];
var SSO_INTRODUCE_COLUMNS = [
  { label: 'แถวที่ (เดิม)', render: function (r) { return esc_(r.row); } },
  { label: 'Paycode', render: function (r) { return esc_(r.paycode); } },
  { label: 'Paycode Name', render: function (r) { return esc_(r.paycodeName); } },
  { label: 'Introduce By', render: function (r) { return esc_(r.introduceBy); } },
  { label: 'Amount เดิม', render: function (r) { return esc_(fmtAmt(r.amountBefore)); } },
  { label: 'Amount ใหม่', render: function (r) { return esc_(fmtAmt(r.amountAfter)); } },
  { label: 'หมายเหตุ', render: function (r) { return noteBadge(r.note); } }
];
var EXPENSE_COLUMNS = [
  { label: 'แถวที่', render: function (r) { return esc_(r.row); } },
  { label: 'ชื่อ', render: function (r) { return esc_(r.name); } },
  { label: 'EMP ID', render: function (r) { return esc_(r.empId); } },
  { label: 'การดำเนินการ', render: function (r) { return noteBadge(r.action); } }
];
var REMOVE_SSO_COLUMNS = [
  { label: 'แถวที่ (เดิม)', render: function (r) { return esc_(r.row); } },
  { label: 'Paycode', render: function (r) { return esc_(r.paycode); } },
  { label: 'Paycode Name', render: function (r) { return esc_(r.paycodeName); } },
  { label: 'Amount เดิม', render: function (r) { return esc_(fmtAmt(r.amount)); } },
  { label: 'หมายเหตุ', render: function (r) { return noteBadge(r.note); } }
];
var MERGE_COLUMNS = [
  { label: 'ไฟล์ที่', render: function (r) { return esc_(r.fileIndex); } },
  { label: 'ชื่อไฟล์', render: function (r) { return esc_(r.fileName); } },
  { label: 'จำนวนแถวที่เพิ่ม', render: function (r) { return esc_(r.rowsAppended); } },
  { label: 'หมายเหตุ', render: function (r) { return noteBadge(r.note); } }
];
var CHANGE_HEADER_COLUMNS = [
  { label: 'คอลัมน์', render: function (r) { return esc_(r.column); } },
  { label: 'ชื่อเดิม', render: function (r) { return esc_(r.oldLabel); } },
  { label: 'ชื่อใหม่', render: function (r) { return esc_(r.newLabel); } }
];

function buildResultBody(fn, summary) {
  var html = '';
  if (fn.id === 'sso750') {
    html += statRowHtml([
      { value: summary.matched, label: 'พบแถวที่ตรงเงื่อนไข', color: 'blue' },
      { value: summary.reduced, label: 'ลด Amount (> 750)', color: 'green' },
      { value: summary.zeroed, label: 'ปรับเป็น 0 (≤ 750)', color: 'orange' }
    ]);
    html += detailTableHtml(SSO_ADJUST_COLUMNS, summary.details);
  } else if (fn.id === 'ssoIntroduceBy') {
    html += statRowHtml([
      { value: summary.matched, label: 'พบแถวที่ตรงเงื่อนไข', color: 'blue' },
      { value: summary.zeroedPRTR, label: 'ปรับเป็น 0 (Introduce By = PRTR)', color: 'red' },
      { value: summary.unchangedCLNT, label: 'ไม่เปลี่ยน (CLNT)', color: 'green' }
    ]);
    html += detailTableHtml(SSO_INTRODUCE_COLUMNS, summary.details);
  } else if (fn.id === 'duplicate') {
    html += statRowHtml([
      { value: summary.added, label: 'เพิ่มแถวใหม่', color: 'green' },
      { value: summary.updated, label: 'อัพเดทแถวเดิม', color: 'blue' }
    ]);
    html += detailTableHtml(EXPENSE_COLUMNS, summary.details);
  } else if (fn.id === 'sso750PlusDuplicate' || fn.id === 'ssoIntroduceByPlusDuplicate') {
    var isPrtr = fn.id === 'sso750PlusDuplicate';
    html += statRowHtml(isPrtr ? [
      { value: summary.matched, label: 'พบแถวที่ตรงเงื่อนไข SSO', color: 'blue' },
      { value: summary.reduced, label: 'ลด Amount (> 750)', color: 'green' },
      { value: summary.zeroed, label: 'ปรับเป็น 0 (≤ 750)', color: 'orange' }
    ] : [
      { value: summary.matched, label: 'พบแถวที่ตรงเงื่อนไข SSO', color: 'blue' },
      { value: summary.zeroedPRTR, label: 'ปรับเป็น 0 (PRTR)', color: 'red' },
      { value: summary.unchangedCLNT, label: 'ไม่เปลี่ยน (CLNT)', color: 'green' }
    ]);
    html += '<div class="result-sub">ผลปรับ SSO</div>';
    html += detailTableHtml(isPrtr ? SSO_ADJUST_COLUMNS : SSO_INTRODUCE_COLUMNS, summary.ssoDetails);
    html += statRowHtml([
      { value: summary.added, label: 'เพิ่มแถว EXPENSE ใหม่', color: 'green' },
      { value: summary.updated, label: 'อัพเดทแถว EXPENSE เดิม', color: 'blue' }
    ]);
    html += '<div class="result-sub">แถว EXPENSE</div>';
    html += detailTableHtml(EXPENSE_COLUMNS, summary.expenseDetails);
  } else if (fn.id === 'removeSso') {
    html += statRowHtml([
      { value: summary.removed, label: 'บรรทัดที่ลบออก', color: 'red' },
      { value: summary.countT2A3, label: 'Paycode T2A3', color: 'green' },
      { value: summary.countTZ74, label: 'Paycode TZ74', color: 'orange' }
    ]);
    html += detailTableHtml(REMOVE_SSO_COLUMNS, summary.details);
  } else if (fn.id === 'merge') {
    html += statRowHtml([
      { value: summary.filesAppended, label: 'ไฟล์ที่รวมสำเร็จ', color: 'blue' },
      { value: summary.rowsAppended, label: 'แถวที่เพิ่มเข้ามา', color: 'green' },
      { value: (summary.rejected || []).length, label: 'ไฟล์ที่ถูกปฏิเสธ', color: 'red' }
    ]);
    html += detailTableHtml(MERGE_COLUMNS, summary.details);
    if (summary.rejected && summary.rejected.length) {
      html += summary.rejected.map(function (r) {
        return '<div class="err-line">⚠ ไฟล์ที่ ' + esc_(r.fileIndex) + ': ' + esc_(r.reason) + '</div>';
      }).join('');
    }
  } else if (fn.id === 'changeHeader') {
    html += statRowHtml([
      { value: summary.rowsDeleted, label: 'แถวที่ลบออก', color: 'red' },
      { value: (summary.renamed || []).length, label: 'คอลัมน์ที่เปลี่ยนชื่อ', color: 'blue' }
    ]);
    html += detailTableHtml(CHANGE_HEADER_COLUMNS, summary.details);
  }
  return html;
}

function renderSummary(fn, summary) {
  var box = document.getElementById('summaryBox');
  box.style.display = 'block';
  var meta = FN_META[fn.id] || { icon: '📄', title: 'สรุปผล' };
  var processedAtStr = state.processedAt ? state.processedAt.toLocaleString('th-TH') : '-';

  var html = '<div class="result-title">' + meta.icon + ' ' + esc_(meta.title) + '</div>';
  html += buildResultBody(fn, summary);
  html += '<div class="result-footer">ไฟล์ต้นฉบับ: ' + esc_(state.sourceLabel) + ' | ประมวลผลเมื่อ: ' + esc_(processedAtStr) + '</div>';
  html += '<div class="result-actions">' +
    '<button class="btn-main" id="btnDownload">⬇️ ดาวน์โหลดไฟล์ที่ประมวลผลแล้ว</button>' +
    '<button class="btn-outline" id="btnReprocess">🔄 ประมวลผลไฟล์ใหม่</button>' +
    '</div>';

  box.innerHTML = html;

  document.getElementById('btnDownload').addEventListener('click', downloadResult);
  document.getElementById('btnReprocess').addEventListener('click', resetForNewFile);
}

function downloadResult() {
  if (!state.resultBytes) return;
  var blob = new Blob([state.resultBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = (state.resultBaseName || 'GL_Invoice') + '_' + state.fnId + '.xlsx';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(function () { URL.revokeObjectURL(url); }, 3000);
}

// init
renderSidebar();
renderFnDesc();
renderDropzoneHint();
