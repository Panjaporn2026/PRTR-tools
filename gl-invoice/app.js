// ══════════════════════════════════════════════════════
//  UI wiring for the GL Invoice Processing tool. See sheetmodel.js / styles.js / functions.js for
//  the actual logic -- this file is state + rendering + the top-level pipeline that calls into them.
// ══════════════════════════════════════════════════════

var FUNCTIONS = [
  { id: 'duplicate', label: '1. Duplicate', multi: false,
    desc: 'เพิ่ม/อัพเดทแถวต่อท้ายทุกคน (unique NAME) ตามที่เลือกไว้ด้านล่าง\nแถวที่เพิ่ม/อัพเดททุกแถว — ตัวอักษรสีแดง\nPaycode Code = EXPENSE, Paycode Name = ค่าใช้จ่าย\nAccount = 51110116, Grouping = E51110116, Amount = (ว่าง)' },
  { id: 'merge', label: '2. Merge', multi: true,
    desc: 'โยนไฟล์ที่ 1 แล้วโยนไฟล์ที่ 2 แล้วโยนไฟล์ที่ 3-6 ตามลำดับ\nระบบจะนำข้อมูล (หลัง header) ของไฟล์ที่ 2-6 ต่อท้ายไฟล์ที่ 1\nHeader ของ output ยึดตามไฟล์ที่ 1 ทั้งหมด รูปแบบของไฟล์ห้ามเปลี่ยนแปลง' },
  { id: 'changeHeaderDynamic', label: '3. Change Header (ไฟล์รูปแบบใหม่)', multi: false,
    desc: 'สำหรับไฟล์รูปแบบใหม่ที่ระบบเพิ่มแถว metadata มา (Start-End Period, Payment Date)\nลบแถว metadata ทั้งหมดตั้งแต่แถวที่ 4 จนถึงก่อนแถว Header (หาอัตโนมัติ) + เปลี่ยนชื่อ Column:\n• Period → Calendar Group\n• Paycode Code → PIN Name' },
  { id: 'deleteGL', label: '4.ลบบรรทัด EWF GL 51110129', multi: true, batch: true, minFiles: 1, processLabel: '▶ เริ่มลบบรรทัด',
    desc: 'ลบทุกแถวที่ Account = 51110129 แล้วเลื่อนแถวถัดไปขึ้นมาแทน\nข้อมูลอื่น รูปแบบ สี และ Head Count ไม่เปลี่ยน\nเลือกได้หลายไฟล์ ได้ไฟล์แยกเหมือนเดิม ชื่อไฟล์เดิม (หลายไฟล์ดาวน์โหลดเป็น ZIP)\nไฟล์ที่ไม่มี EWF GL 51110129 จะได้ไฟล์เดิมกลับมาโดยไม่แก้ไขอะไร' },
  // panel:'split' = ฟังก์ชันนี้มีหน้าจอของตัวเอง (split.js) ไม่ใช้กล่องอัปโหลด/สรุปผลร่วมของฟังก์ชัน 1-4
  { id: 'split', label: '5. แยกไฟล์ตามคอลัมน์', multi: false, panel: 'split',
    desc: 'อัปโหลดไฟล์ GL Invoice แล้วติ๊กคอลัมน์ที่ต้องการใช้แยก ระบบจะสร้างไฟล์ใหม่ต่อค่าที่ไม่ซ้ำกัน\nโดยคงข้อมูลและรูปแบบเดิมไว้ทุกอย่าง ยกเว้นช่อง Head Count ที่ปรับตามจำนวนพนักงานในไฟล์นั้น' }
];

var FN_META = {
  duplicate: { icon: '➕', title: 'สรุปผล Duplicate — แถวที่เพิ่ม/อัพเดท' },
  merge: { icon: '🔗', title: 'สรุปผล Merge — รวมไฟล์' },
  changeHeaderDynamic: { icon: '📝', title: 'สรุปผล Change Header (ไฟล์รูปแบบใหม่)' },
  deleteGL: { icon: '🗑️', title: 'สรุปผล ลบบรรทัด EWF GL 51110129' }
};

var state = { fnId: FUNCTIONS[0].id, files: [], resultBytes: null, resultBaseName: null, processedAt: null, sourceLabel: null, extraLineTypes: ['EXPENSE'] };

// Selectable line types for the Duplicate step. EXPENSE is first and checked by default (matches
// the tool's original always-on behavior); the other three are opt-in, unchecked by default since
// most months don't need them.
var EXTRA_LINE_TYPE_OPTIONS = [
  { key: 'EXPENSE', label: 'ค่าใช้จ่าย (EXPENSE)', sub: 'Paycode EXPENSE · Account 51110116 · Grouping E51110116', checkedByDefault: true },
  { key: 'PROVIDENT_FUND_REFUND', label: 'คืนเงินสมทบกองทุนสำรองเลี้ยงชีพนายจ้าง', sub: 'Paycode T208 · Account 51110103 · Grouping E51110103' },
  { key: 'ACCIDENT_REFUND', label: 'หักค่าประกันอุบัติเหตุ ค่าใช้จ่ายลูกค้า (คืนค่าประกันซื่อสัตย์)', sub: 'Paycode AC CL D AC · Account 51110104 · Grouping E51110104' },
  { key: 'OTHER_INCOME', label: 'Other Income', sub: 'Paycode OTHER INCOME · Account 51110108 · Grouping E51110108' }
];
var FN_IDS_WITH_EXTRA_LINE_OPTIONS = ['duplicate'];

function dbg(msg) { console.log(msg); }
function setStatus(msg, cls) {
  var el = document.getElementById('statusBox');
  el.textContent = msg || '';
  el.className = 'status-box' + (cls ? ' ' + cls : '');
  el.style.display = msg ? 'block' : 'none';
}
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

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
function toggleExtraLineType(key, checked) {
  var i = state.extraLineTypes.indexOf(key);
  if (checked && i === -1) state.extraLineTypes.push(key);
  else if (!checked && i !== -1) state.extraLineTypes.splice(i, 1);
}
function renderExtraLineOptions() {
  var el = document.getElementById('extraLineOptions');
  if (FN_IDS_WITH_EXTRA_LINE_OPTIONS.indexOf(state.fnId) === -1) { el.innerHTML = ''; el.style.display = 'none'; return; }
  el.style.display = 'block';
  el.innerHTML = '<div class="extra-line-title">➕ เลือกแถวที่ต้องการเพิ่ม/อัพเดท</div>' +
    EXTRA_LINE_TYPE_OPTIONS.map(function (o) {
      var checked = state.extraLineTypes.indexOf(o.key) !== -1 ? ' checked' : '';
      return '<label class="extra-line-opt">' +
        '<input type="checkbox" data-key="' + o.key + '"' + checked + '>' +
        '<span><b>' + esc_(o.label) + '</b><br><span class="extra-line-sub">' + esc_(o.sub) + '</span></span>' +
        '</label>';
    }).join('');
  Array.prototype.forEach.call(el.querySelectorAll('input[type=checkbox]'), function (cb) {
    cb.addEventListener('change', function () { toggleExtraLineType(cb.getAttribute('data-key'), cb.checked); });
  });
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
  btn.textContent = fn.processLabel || '▶ เริ่มรวมไฟล์';
  btn.disabled = state.files.length < (fn.minFiles || 2);
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

function renderPanels() {
  var isSplit = currentFn().panel === 'split';
  document.getElementById('genericPanel').style.display = isSplit ? 'none' : '';
  document.getElementById('splitPanel').style.display = isSplit ? '' : 'none';
}
function selectFunction(id) {
  state.fnId = id;
  state.files = [];
  state.extraLineTypes = EXTRA_LINE_TYPE_OPTIONS.filter(function (o) { return o.checkedByDefault; }).map(function (o) { return o.key; });
  resetOutputUI();
  renderSidebar();
  renderFnDesc();
  renderExtraLineOptions();
  renderDropzoneHint();
  renderFileList();
  renderPanels();
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
    setStatus(fn.batch
      ? '✅ เพิ่มไฟล์แล้ว (' + state.files.length + ' ไฟล์) — เพิ่มไฟล์ได้อีก หรือกด "เริ่มลบบรรทัด" เมื่อพร้อม'
      : '✅ เพิ่มไฟล์แล้ว (' + state.files.length + ' ไฟล์) — เรียงลำดับถูกต้องหรือยัง? กด "เริ่มรวมไฟล์" เมื่อพร้อม', 'info');
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
    state.resultFileName = null;
    if (fn.batch) {
      result = await runDeleteGLBatch(state.files);
    } else if (fn.multi) {
      var bufs = [], names = state.files.map(function (f) { return f.name; });
      for (var i = 0; i < state.files.length; i++) bufs.push(await state.files[i].arrayBuffer());
      result = await runMergeFunction(bufs, names);
      state.resultBaseName = state.files[0].name.replace(/\.[^.]+$/, '');
      state.sourceLabel = state.files.map(function (f) { return f.name; }).join(', ');
    } else {
      var buf = await state.files[0].arrayBuffer();
      result = await runSingleFileFunction(fn.id, buf, state.extraLineTypes);
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
    return '<tr' + rowClass + '>' + columns.map(function (c) {
      return '<td' + (c.wrap ? ' class="wrap-cell"' : '') + '>' + c.render(r) + '</td>';
    }).join('') + '</tr>';
  }).join('');
  return '<div class="detail-table-wrap"><table class="detail-table"><thead>' + thead + '</thead><tbody>' + tbody + '</tbody></table></div>';
}

var EXPENSE_COLUMNS = [
  { label: 'แถวที่', render: function (r) { return esc_(r.row); } },
  { label: 'ชื่อ', render: function (r) { return esc_(r.name); } },
  { label: 'EMP ID', render: function (r) { return esc_(r.empId); } },
  { label: 'Paycode Code', render: function (r) { return esc_(r.paycodeCode); } },
  { label: 'Paycode Name', render: function (r) { return esc_(r.paycodeName); } },
  { label: 'Account', render: function (r) { return esc_(r.account); } },
  { label: 'Grouping', render: function (r) { return esc_(r.grouping); } },
  { label: 'การดำเนินการ', render: function (r) { return noteBadge(r.action); } }
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

function fmtAmount(n) { return typeof n === 'number' ? n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : esc_(n); }
var DELETE_GL_FILE_COLUMNS = [
  { label: 'ชื่อไฟล์', wrap: true, render: function (r) { return esc_(r.fileName); } },
  { label: 'แถวที่ลบ', render: function (r) { return esc_(r.count); } },
  { label: 'ยอด Amount ที่ลบ', render: function (r) { return fmtAmount(r.amount); } },
  { label: 'หมายเหตุ', render: function (r) { return noteBadge(r.note); } }
];
var DELETE_GL_ROW_COLUMNS = [
  { label: 'ไฟล์', wrap: true, render: function (r) { return esc_(r.fileName); } },
  { label: 'แถวที่ (เดิม)', render: function (r) { return esc_(r.row); } },
  { label: 'ชื่อ', render: function (r) { return esc_(r.name); } },
  { label: 'EMP ID', render: function (r) { return esc_(r.empId); } },
  { label: 'Paycode Name', wrap: true, render: function (r) { return esc_(r.paycodeName); } },
  { label: 'Grouping', render: function (r) { return esc_(r.grouping); } },
  { label: 'Amount', render: function (r) { return fmtAmount(r.amount); } }
];

// Runs function 4 over every uploaded file. Each file keeps its own original name; 1 file is
// downloaded as-is, 2+ files as one ZIP. Files that can't be read are reported and left out.
async function runDeleteGLBatch(files) {
  var summary = { files: [], rows: [], rejected: [], totalDeleted: 0 };
  var outs = [];
  for (var i = 0; i < files.length; i++) {
    var f = files[i];
    setStatus('⏳ กำลังประมวลผล ' + (i + 1) + '/' + files.length + ': ' + f.name, 'info');
    try {
      var r = await runDeleteGLFunction(await f.arrayBuffer());
      var s = r.summary;
      outs.push({ name: f.name, data: r.outputBytes });
      summary.totalDeleted += s.deleted.length;
      summary.files.push({ fileName: f.name, count: s.deleted.length, amount: s.amount,
        note: s.unchanged ? 'ไม่พบ EWF GL 51110129 (ไฟล์เดิม)' : 'ลบแล้ว' });
      s.deleted.forEach(function (d) { d.fileName = f.name; summary.rows.push(d); });
    } catch (err) {
      summary.rejected.push({ fileName: f.name, reason: 'อ่านไฟล์ไม่สำเร็จ: ' + err.message });
    }
  }
  if (!outs.length) return { ok: false, summary: { errors: summary.rejected.map(function (r) { return r.fileName + ': ' + r.reason; }) } };
  state.sourceLabel = files.map(function (f) { return f.name; }).join(', ');
  if (outs.length === 1) {
    state.resultFileName = outs[0].name;
    return { ok: true, summary: summary, outputBytes: outs[0].data };
  }
  var d = new Date(), ymd = d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
  state.resultFileName = 'GL_Invoice_ลบGL51110129_' + ymd + '.zip';
  return { ok: true, summary: summary, outputBytes: buildZip(outs) };
}

function buildResultBody(fn, summary) {
  var html = '';
  if (fn.id === 'duplicate') {
    html += statRowHtml([
      { value: summary.added, label: 'เพิ่มแถวใหม่', color: 'green' },
      { value: summary.updated, label: 'อัพเดทแถวเดิม', color: 'blue' }
    ]);
    html += detailTableHtml(EXPENSE_COLUMNS, summary.details);
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
  } else if (fn.id === 'deleteGL') {
    html += statRowHtml([
      { value: summary.files.length, label: 'ไฟล์ที่ประมวลผล', color: 'blue' },
      { value: summary.totalDeleted, label: 'แถวที่ลบออก', color: 'red' },
      { value: summary.rejected.length, label: 'ไฟล์ที่อ่านไม่ได้', color: 'orange' }
    ]);
    html += '<div class="result-sub">แยกตามไฟล์</div>' + detailTableHtml(DELETE_GL_FILE_COLUMNS, summary.files);
    if (summary.rows.length) html += '<div class="result-sub">แถวที่ลบ</div>' + detailTableHtml(DELETE_GL_ROW_COLUMNS, summary.rows);
    if (summary.rejected.length) {
      html += summary.rejected.map(function (r) {
        return '<div class="err-line">⚠ ' + esc_(r.fileName) + ': ' + esc_(r.reason) + ' (ไม่ได้ใส่ในไฟล์ผลลัพธ์)</div>';
      }).join('');
    }
  } else if (fn.id === 'changeHeaderDynamic') {
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
  var isZip = !!state.resultFileName && /\.zip$/i.test(state.resultFileName);
  var blob = new Blob([state.resultBytes], { type: isZip ? 'application/zip' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = state.resultFileName || ((state.resultBaseName || 'GL_Invoice') + '_' + state.fnId + '.xlsx');
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(function () { URL.revokeObjectURL(url); }, 3000);
}

// init
renderSidebar();
renderFnDesc();
renderExtraLineOptions();
renderDropzoneHint();
