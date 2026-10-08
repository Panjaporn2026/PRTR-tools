// ══════════════════════════════════════════════════════
//  Function 6 (ปลดรหัสไฟล์ Excel): UI of the user's "Excel Unlock" tool, embedded in GL Invoice
//  Processing. Script body is the user's verbatim; the only change is that $() looks up ids with
//  the ul_ prefix, because #unlockPanel shares the page with the other functions' panels (which
//  already use ids like fileInput / fileList / drop / status / result / btnZip). Decryption itself
//  is officecrypto.js (verbatim).
// ══════════════════════════════════════════════════════
(function () {
  'use strict';
  var $ = function (id) { return document.getElementById('ul_' + id); }; // ids prefixed ul_ inside #unlockPanel
  var state = { files: [], results: null, busy: false };

  function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function fmtSize(n) { return n >= 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB'; }
  function status(t) { $('status').textContent = t || ''; }
  function warn(t) { $('warn').textContent = t || ''; $('warn').classList.toggle('show', !!t); }

  // ---- files ----
  var drop = $('drop');
  $('fileInput').addEventListener('change', function (e) { addFiles(e.target.files); e.target.value = ''; });
  ['dragenter', 'dragover'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('is-over'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('is-over'); }); });
  drop.addEventListener('drop', function (e) { addFiles(e.dataTransfer.files); });
  $('btnClear').addEventListener('click', function () { state.files = []; resetResults(); renderFiles(); });

  function addFiles(list) {
    var skipped = [];
    Array.prototype.forEach.call(list || [], function (f) {
      if (!/\.(xlsx|xlsm)$/i.test(f.name)) { skipped.push(f.name); return; }
      state.files = state.files.filter(function (x) { return x.name !== f.name; }); // same name -> replace
      state.files.push(f);
    });
    resetResults(); renderFiles();
    warn(skipped.length ? 'ข้ามไฟล์ที่ไม่ใช่ .xlsx / .xlsm: ' + skipped.join(', ') : '');
  }
  function renderFiles() {
    $('fileList').innerHTML = state.files.map(function (f, i) {
      return '<div class="file-row"><span class="fname">' + esc(f.name) + '</span><span class="fsize">' + fmtSize(f.size) +
        '</span><button type="button" class="icon-btn" data-rm="' + i + '" aria-label="เอาไฟล์ ' + esc(f.name) + ' ออก">✕</button></div>';
    }).join('');
    $('listActions').hidden = !state.files.length;
    $('fileCount').textContent = state.files.length + ' ไฟล์';
    refreshRun();
  }
  $('fileList').addEventListener('click', function (e) {
    var b = e.target.closest('[data-rm]'); if (!b) return;
    state.files.splice(+b.dataset.rm, 1); resetResults(); renderFiles();
  });

  // ---- passwords ----
  var showPw = false;
  function addPwRow(value) {
    var row = document.createElement('div'); row.className = 'pw-row';
    row.innerHTML = '<span class="pw-idx"></span><input type="' + (showPw ? 'text' : 'password') + '" autocomplete="off" placeholder="รหัสผ่าน">' +
      '<button type="button" class="icon-btn" aria-label="ลบรหัสนี้">✕</button>';
    row.querySelector('input').value = value || '';
    row.querySelector('input').addEventListener('input', function () { resetResults(); refreshRun(); });
    // pasting several lines into one box splits them into separate passwords
    row.querySelector('input').addEventListener('paste', function (e) {
      var t = (e.clipboardData || window.clipboardData).getData('text');
      if (!/[\r\n]/.test(t)) return;
      e.preventDefault();
      var parts = t.split(/\r?\n/).filter(function (x) { return x !== ''; });
      this.value = parts.shift() || '';
      parts.forEach(function (p) { addPwRow(p); });
      resetResults(); refreshRun();
    });
    row.querySelector('button').addEventListener('click', function () {
      if ($('pwList').children.length > 1) row.remove(); else row.querySelector('input').value = '';
      numberPw(); resetResults(); refreshRun();
    });
    $('pwList').appendChild(row); numberPw();
    return row;
  }
  function numberPw() { Array.prototype.forEach.call($('pwList').children, function (r, i) { r.querySelector('.pw-idx').textContent = (i + 1) + '.'; }); }
  function passwords() { return Array.prototype.map.call($('pwList').querySelectorAll('input'), function (i) { return i.value; }).filter(function (v) { return v !== ''; }); }
  $('btnAddPw').addEventListener('click', function () { addPwRow('').querySelector('input').focus(); });
  $('btnShowPw').addEventListener('click', function () {
    showPw = !showPw;
    $('pwList').querySelectorAll('input').forEach(function (i) { i.type = showPw ? 'text' : 'password'; });
    $('btnShowPw').textContent = showPw ? 'ซ่อนรหัส' : 'แสดงรหัส';
  });
  addPwRow('');

  // ---- run ----
  function refreshRun() {
    var n = state.files.length, p = passwords().length;
    $('btnRun').disabled = state.busy || !n || !p;
    $('summary').textContent = !n || !p ? 'เลือกไฟล์และใส่รหัสอย่างน้อย 1 รหัส' : n + ' ไฟล์, ' + p + ' รหัส';
  }
  function resetResults() { state.results = null; $('result').innerHTML = ''; $('dlActions').hidden = true; status(''); }

  var KIND_TEXT = { xls: 'เป็นไฟล์ .xls รุ่นเก่า ยังไม่รองรับ', unsupported: 'รูปแบบการเข้ารหัสนี้ยังไม่รองรับ', notexcel: 'ไม่ใช่ไฟล์ Excel หรือไฟล์เสียหาย' };

  $('btnRun').addEventListener('click', async function () {
    if (state.busy) return;
    if (typeof XLSX === 'undefined' || !XLSX.CFB || !window.crypto || !crypto.subtle) { status('โหลดตัวอ่านไฟล์ไม่สำเร็จ ตรวจการเชื่อมต่ออินเทอร์เน็ตแล้วรีเฟรชหน้านี้'); return; }
    state.busy = true; refreshRun(); warn('');
    var pws = passwords(), lastOk = -1, results = state.files.map(function (f) { return { file: f, state: 'wait' }; });
    state.results = results; render();
    for (var i = 0; i < results.length; i++) {
      var r = results[i];
      try {
        var bytes = new Uint8Array(await r.file.arrayBuffer());
        var x = OfficeCrypto.inspect(bytes);
        if (x.kind === 'plain') { r.state = 'plain'; r.data = bytes; }
        else if (x.kind !== 'agile' && x.kind !== 'standard') { r.state = 'bad'; r.reason = KIND_TEXT[x.kind]; }
        else {
          r.state = 'bad'; r.reason = 'ไม่มีรหัสที่ใส่ไว้ตรงกับไฟล์นี้';
          // try the password that worked for the previous file first: batches usually share one
          var order = pws.map(function (_, j) { return j; });
          if (lastOk >= 0) order = [lastOk].concat(order.filter(function (j) { return j !== lastOk; }));
          for (var t = 0; t < order.length; t++) {
            var k = order[t];
            status('ไฟล์ ' + (i + 1) + '/' + results.length + ': กำลังลองรหัส (' + (t + 1) + '/' + pws.length + ')');
            await new Promise(function (res) { setTimeout(res, 0); }); // let the page repaint between tries
            var out = await OfficeCrypto.tryPassword(x, pws[k]);
            if (out) { r.state = 'ok'; r.data = out; r.pwIndex = k + 1; lastOk = k; break; }
          }
        }
      } catch (e) { r.state = 'bad'; r.reason = 'อ่านไฟล์ไม่สำเร็จ: ' + (e && e.message || e); }
      render();
    }
    var ok = results.filter(function (r) { return r.data; }).length;
    status('เสร็จแล้ว: ใช้งานได้ ' + ok + ' จาก ' + results.length + ' ไฟล์');
    $('dlActions').hidden = !ok;
    state.busy = false; refreshRun();
  });

  function render() {
    var rows = state.results.map(function (r, i) {
      var badge = r.state === 'ok' ? '<span class="badge b-ok">ปลดรหัสแล้ว (รหัสที่ ' + r.pwIndex + ')</span>'
        : r.state === 'plain' ? '<span class="badge b-plain">ไม่ได้ล็อกอยู่แล้ว (ไฟล์เดิม)</span>'
        : r.state === 'wait' ? '<span class="badge b-wait">รอ</span>'
        : '<span class="badge b-bad">ไม่สำเร็จ</span> ' + esc(r.reason);
      return '<tr><td class="fn">' + esc(r.file.name) + '</td><td>' + badge + '</td><td class="num">' +
        (r.data ? '<button type="button" class="ds-btn ds-btn--ghost" data-dl="' + i + '">ดาวน์โหลด</button>' : '') + '</td></tr>';
    }).join('');
    $('result').innerHTML = '<div class="ds-tablewrap" style="margin-top:16px"><table class="ds-table"><thead><tr><th>ชื่อไฟล์</th><th>ผล</th><th class="num"></th></tr></thead><tbody>' + rows + '</tbody></table></div>';
  }

  // ---- download ----
  var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  function save(name, blob) {
    var url = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 5000);
  }
  function mimeFor(name) { return /\.xlsm$/i.test(name) ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : XLSX_MIME; }
  $('result').addEventListener('click', function (e) {
    var b = e.target.closest('[data-dl]'); if (!b) return;
    var r = state.results[+b.dataset.dl]; save(r.file.name, new Blob([r.data], { type: mimeFor(r.file.name) }));
  });
  $('btnZip').addEventListener('click', async function () {
    var ok = state.results.filter(function (r) { return r.data; });
    var z = new JSZip(); ok.forEach(function (r) { z.file(r.file.name, r.data); });
    status('กำลังสร้าง ZIP…');
    var d = new Date(), ymd = d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
    save('Excel_Unlocked_' + ymd + '.zip', await z.generateAsync({ type: 'blob', compression: 'STORE' }));
    status('ดาวน์โหลด ZIP แล้ว (' + ok.length + ' ไฟล์)');
  });
  $('btnEach').addEventListener('click', async function () {
    var ok = state.results.filter(function (r) { return r.data; });
    for (var i = 0; i < ok.length; i++) { save(ok[i].file.name, new Blob([ok[i].data], { type: mimeFor(ok[i].file.name) })); await new Promise(function (res) { setTimeout(res, 600); }); }
    status('ดาวน์โหลดแล้ว ' + ok.length + ' ไฟล์ (ถ้าเบราว์เซอร์ถามเรื่องดาวน์โหลดหลายไฟล์ ให้กดอนุญาต)');
  });
})();
