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

var state = { fnId: FUNCTIONS[0].id, files: [], resultBytes: null, resultBaseName: null };

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
  document.getElementById('summaryBox').style.display = 'none';
  document.getElementById('btnDownload').style.display = 'none';
  setStatus('', '');
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
    document.getElementById('summaryBox').style.display = 'none';
    document.getElementById('btnDownload').style.display = 'none';

    var result;
    if (fn.multi) {
      var bufs = [];
      for (var i = 0; i < state.files.length; i++) bufs.push(await state.files[i].arrayBuffer());
      result = await runMergeFunction(bufs);
      state.resultBaseName = state.files[0].name.replace(/\.[^.]+$/, '');
    } else {
      var buf = await state.files[0].arrayBuffer();
      result = await runSingleFileFunction(fn.id, buf);
      state.resultBaseName = state.files[0].name.replace(/\.[^.]+$/, '');
    }

    if (!result.ok) {
      renderErrorSummary(result.summary);
      setStatus('❌ พบข้อผิดพลาด — ดูรายละเอียดด้านล่าง ไม่ได้สร้างไฟล์ผลลัพธ์', 'err');
      return;
    }

    state.resultBytes = result.outputBytes;
    renderSummary(fn, result.summary);
    document.getElementById('btnDownload').style.display = 'inline-block';
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

function renderSummary(fn, summary) {
  var box = document.getElementById('summaryBox');
  box.style.display = 'block';
  var lines = [];
  if (fn.id === 'sso750') {
    lines.push('พบแถวที่ตรงเงื่อนไข: ' + summary.matched + ' แถว');
    lines.push('ลด Amount (Amount > 750): ' + summary.reduced + ' แถว');
    lines.push('ปรับ Amount เป็น 0 (Amount ≤ 750): ' + summary.zeroed + ' แถว');
  } else if (fn.id === 'ssoIntroduceBy') {
    lines.push('พบแถวที่ตรงเงื่อนไข: ' + summary.matched + ' แถว');
    lines.push('ปรับ Amount เป็น 0 (Introduce By = PRTR): ' + summary.zeroedPRTR + ' แถว');
    lines.push('ไม่เปลี่ยน (Introduce By = CLNT): ' + summary.unchangedCLNT + ' แถว');
  } else if (fn.id === 'duplicate') {
    lines.push('เพิ่มแถว EXPENSE ใหม่: ' + summary.added + ' คน');
    lines.push('อัพเดทแถว EXPENSE ที่มีอยู่แล้ว: ' + summary.updated + ' คน');
  } else if (fn.id === 'sso750PlusDuplicate' || fn.id === 'ssoIntroduceByPlusDuplicate') {
    lines.push('พบแถวที่ตรงเงื่อนไข SSO: ' + summary.matched + ' แถว');
    lines.push('เพิ่มแถว EXPENSE ใหม่: ' + summary.added + ' คน');
    lines.push('อัพเดทแถว EXPENSE ที่มีอยู่แล้ว: ' + summary.updated + ' คน');
  } else if (fn.id === 'removeSso') {
    lines.push('ลบแถวที่ตรงเงื่อนไข: ' + summary.removed + ' แถว');
  } else if (fn.id === 'merge') {
    lines.push('รวมไฟล์สำเร็จ: ' + summary.filesAppended + ' ไฟล์');
    lines.push('จำนวนแถวที่เพิ่มเข้ามา: ' + summary.rowsAppended + ' แถว');
    if (summary.rejected && summary.rejected.length) {
      summary.rejected.forEach(function (r) { lines.push('⚠ ไฟล์ที่ ' + r.fileIndex + ': ' + r.reason); });
    }
  } else if (fn.id === 'changeHeader') {
    lines.push('ลบแถว 4 และ 5 สำเร็จ — header ย้ายไปแถว ' + summary.newHeaderRow);
    lines.push('เปลี่ยนชื่อคอลัมน์: ' + summary.renamed.join(', '));
  }
  box.innerHTML = lines.map(function (l) { return '<div>' + esc_(l) + '</div>'; }).join('');
}

document.getElementById('btnDownload').addEventListener('click', function () {
  if (!state.resultBytes) return;
  var blob = new Blob([state.resultBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url; a.download = (state.resultBaseName || 'GL_Invoice') + '_' + state.fnId + '.xlsx';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(function () { URL.revokeObjectURL(url); }, 3000);
});

// init
renderSidebar();
renderFnDesc();
renderDropzoneHint();
