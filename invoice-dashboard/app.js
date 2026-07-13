(function () {
  'use strict';

  var REQUIRED_HEADERS = ['Document No', 'Status', 'Action'];
  var BUCKET_DEFS = [
    { key: 'error', label: 'ส่งแล้ว Error / ไม่สำเร็จ', cls: 'c-red' },
    { key: 'not_sent', label: 'ยังไม่ได้ส่ง', cls: 'c-orange' },
    { key: 'dont_send', label: 'เลือกไม่ส่ง (Don\'t Send)', cls: 'c-purple' },
    { key: 'sent_opened', label: 'ส่งแล้ว — เปิดอ่านแล้ว', cls: 'c-green' },
    { key: 'sent_not_opened', label: 'ส่งแล้ว — ยังไม่เปิดอ่าน', cls: 'c-blue' },
    { key: 'unknown', label: 'ไม่ทราบสถานะ (ตรวจสอบ)', cls: 'c-gray' }
  ];
  // "In Progress"/"Not Send" = ระบบยังไม่ได้ส่ง (คิวรอ/pending)
  // "Don't Send" = พนักงานกดเลือกเองว่าไม่ต้องส่ง (deliberate, ไม่ใช่ pending) จึงแยกเป็นคนละหมวด
  var NOT_SENT_STATUSES = ["In Progress", "Not Send"];
  var DONT_SEND_STATUSES = ["Don't Send"];
  var INTERNAL_EMAIL = 'osinvoiceteam@prtr.com';
  var DATE_RANGE_DEFS = [
    { key: 'all', label: 'ทั้งหมด' },
    { key: 'today', label: 'วันนี้', days: 0 },
    { key: '7d', label: '7 วันล่าสุด', days: 7 },
    { key: '1m', label: '1 เดือนล่าสุด', days: 30 },
    { key: '2m', label: '2 เดือนล่าสุด', days: 60 }
  ];
  var STORAGE_KEY = 'invoiceDashboard.v1';

  var state = {
    files: [],       // [{name, rows:[record]}]
    records: [],      // merged flat records
    activeBucket: 'all',
    searchText: '',
    dateRange: 'all',
    lastLoadedAt: null   // Date -- when the current data was uploaded/read (or restored from a prior upload)
  };

  // SheetJS (cellDates:true) parses Excel date serials into UTC-midnight Date objects (e.g. "13
  // Jul 2026" becomes 2026-07-13T00:00:00Z) specifically to dodge DST ambiguity -- so every date
  // in this file must be read/compared using the UTC getters, never the local ones. Reading with
  // local getters in a timezone behind UTC (which is most of them) silently displays/filters every
  // date one day too early (confirmed: a file spanning 1 Jun - 13 Jul showed as 31 May - 12 Jul).
  function utcMidnightToday() {
    var now = new Date();
    return new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
  }

  function filterByDateRange(records) {
    var def = DATE_RANGE_DEFS.filter(function (d) { return d.key === state.dateRange; })[0];
    if (!def || def.days == null) return records;
    var cutoff = utcMidnightToday();
    cutoff.setUTCDate(cutoff.getUTCDate() - def.days);
    return records.filter(function (r) {
      return r.docDate instanceof Date && r.docDate >= cutoff;
    });
  }

  function esc_(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function findHeaderRow(aoa, mustHave) {
    var matches = [];
    for (var r = 0; r < aoa.length; r++) {
      var row = aoa[r] || [];
      var texts = row.map(function (v) { return (v == null ? '' : String(v)).trim(); });
      var ok = mustHave.every(function (h) { return texts.indexOf(h) !== -1; });
      if (ok) matches.push(r);
    }
    if (matches.length === 0) {
      throw new Error('ไม่พบแถวหัวตาราง (ต้องมีคอลัมน์ ' + mustHave.join(', ') + ') ในไฟล์นี้ — โครงสร้างไฟล์อาจไม่ใช่ไฟล์ Export จาก getinvoice.net');
    }
    if (matches.length > 1) {
      throw new Error('พบแถวหัวตารางที่ตรงเงื่อนไขมากกว่า 1 แถว (แถว ' + matches.map(function (m) { return m + 1; }).join(', ') + ') — กรุณาตรวจสอบไฟล์');
    }
    return matches[0];
  }

  function colIndex(headerRow, name) {
    for (var c = 0; c < headerRow.length; c++) {
      if ((headerRow[c] == null ? '' : String(headerRow[c]).trim()) === name) return c;
    }
    return -1;
  }

  // SheetJS's cellDates date parsing (at least for this workbook/version) doesn't land exactly on
  // midnight UTC -- observed consistently ~7 hours short (e.g. "13 Jul 2026" parses to
  // 2026-07-12T16:59:56Z), which happens to equal this org's UTC+7 offset almost to the second.
  // Document Date has no meaningful time-of-day component in the source data, so round to the
  // nearest whole UTC day rather than depend on exactly which direction/magnitude that skew runs.
  function normalizeToUTCDay(v) {
    if (!(v instanceof Date) || isNaN(v.getTime())) return v;
    return new Date(Math.round(v.getTime() / 86400000) * 86400000);
  }

  function toNumber(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    var n = parseFloat(String(v).replace(/,/g, ''));
    return isNaN(n) ? 0 : n;
  }

  function toDateLabel(v) {
    if (v == null || v === '') return '—';
    if (v instanceof Date) {
      var d = v.getUTCDate(), m = v.getUTCMonth() + 1, y = v.getUTCFullYear();
      return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
    }
    return String(v);
  }

  function splitEmails(emailAddress) {
    if (!emailAddress) return [];
    return String(emailAddress).split(/;|\s+and\s+/i).map(function (s) { return s.trim(); }).filter(Boolean);
  }

  // Action is a comma-separated per-send-event list (open/none/bounce); Email Address is a
  // ;-separated recipient list. They line up 1:1 only when the counts match (a doc that was
  // resent produces more action events than recipients) -- pairing is only trusted then.
  function classify(status, action, emailAddress) {
    var st = (status == null ? '' : String(status)).trim();
    var tokens = (action == null ? '' : String(action)).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    var hasBounce = tokens.indexOf('bounce') !== -1;
    var hasOpen = tokens.indexOf('open') !== -1;

    var hasRealOpenDespiteBounce = false;
    if (hasBounce && hasOpen) {
      var emails = splitEmails(emailAddress);
      if (emails.length === tokens.length) {
        for (var i = 0; i < tokens.length; i++) {
          if (tokens[i] === 'open' && emails[i].toLowerCase() !== INTERNAL_EMAIL) {
            hasRealOpenDespiteBounce = true;
            break;
          }
        }
      }
    }

    if (st === 'Failure') return 'error';
    if (hasBounce && !hasRealOpenDespiteBounce) return 'error';
    if (DONT_SEND_STATUSES.indexOf(st) !== -1) return 'dont_send';
    if (NOT_SENT_STATUSES.indexOf(st) !== -1) return 'not_sent';
    if (st === 'Success') return hasOpen ? 'sent_opened' : 'sent_not_opened';
    return 'unknown';
  }

  function parseWorkbookFile(file) {
    return file.arrayBuffer().then(function (buf) {
      var wb = XLSX.read(buf, { type: 'array', cellDates: true });
      var records = [];
      wb.SheetNames.forEach(function (sheetName) {
        var ws = wb.Sheets[sheetName];
        var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
        var headerRowIdx;
        try {
          headerRowIdx = findHeaderRow(aoa, REQUIRED_HEADERS);
        } catch (e) {
          return; // this sheet isn't the export table sheet; skip silently, only fail if NO sheet matches (checked by caller)
        }
        var headerRow = aoa[headerRowIdx];
        var idx = {
          docNo: colIndex(headerRow, 'Document No'),
          type: colIndex(headerRow, 'Type'),
          docDate: colIndex(headerRow, 'Document Date'),
          orderNo: colIndex(headerRow, 'Order No'),
          total: colIndex(headerRow, 'Total'),
          buyerName: colIndex(headerRow, 'Buyer Name'),
          email: colIndex(headerRow, 'Email Address'),
          status: colIndex(headerRow, 'Status'),
          sentDate: colIndex(headerRow, 'Sent Date'),
          action: colIndex(headerRow, 'Action'),
          error: colIndex(headerRow, 'Error')
        };
        for (var r = headerRowIdx + 1; r < aoa.length; r++) {
          var row = aoa[r];
          if (!row || row[idx.docNo] == null || row[idx.docNo] === '') continue;
          var status = idx.status >= 0 ? row[idx.status] : null;
          var action = idx.action >= 0 ? row[idx.action] : null;
          var email = idx.email >= 0 ? row[idx.email] : '';
          records.push({
            docNo: row[idx.docNo],
            type: idx.type >= 0 ? row[idx.type] : '',
            docDate: idx.docDate >= 0 ? normalizeToUTCDay(row[idx.docDate]) : null,
            orderNo: idx.orderNo >= 0 ? row[idx.orderNo] : '',
            total: idx.total >= 0 ? toNumber(row[idx.total]) : 0,
            buyerName: idx.buyerName >= 0 ? row[idx.buyerName] : '',
            email: email,
            status: status,
            sentDate: idx.sentDate >= 0 ? row[idx.sentDate] : null,
            action: action,
            error: idx.error >= 0 ? row[idx.error] : '',
            bucket: classify(status, action, email),
            sourceFile: file.name
          });
        }
      });
      if (records.length === 0) {
        throw new Error('ไม่พบแถวข้อมูลที่ใช้ได้ในไฟล์ "' + file.name + '"');
      }
      return records;
    });
  }

  // ---- UI wiring ----
  var dz = document.getElementById('dropzone');
  var fileInput = document.getElementById('fileInput');
  var fileListEl = document.getElementById('fileList');
  var statusBox = document.getElementById('statusBox');
  var dashboard = document.getElementById('dashboard');

  function setStatus(msg, kind) {
    statusBox.style.display = 'block';
    statusBox.className = 'status-box ' + (kind || 'info');
    statusBox.textContent = msg;
  }
  function clearStatus() {
    statusBox.style.display = 'none';
  }

  function renderFileList() {
    if (!state.files.length) { fileListEl.innerHTML = ''; return; }
    fileListEl.innerHTML = state.files.map(function (f, i) {
      return '<div class="file-row"><span class="idx">' + (i + 1) + '</span>' +
        '<span class="fname">' + esc_(f.name) + ' (' + f.rows.length + ' แถว)</span>' +
        '<button class="rm" data-i="' + i + '" title="ลบไฟล์">✕</button></div>';
    }).join('');
    Array.prototype.forEach.call(fileListEl.querySelectorAll('.rm'), function (btn) {
      btn.addEventListener('click', function () {
        state.files.splice(parseInt(btn.getAttribute('data-i'), 10), 1);
        rebuildRecords();
      });
    });
  }

  function saveToStorage() {
    state.lastLoadedAt = new Date();
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ files: state.files, savedAt: state.lastLoadedAt.toISOString() }));
    } catch (e) {
      // localStorage full/blocked (private mode) -- non-critical, just skip persistence
    }
  }

  function loadFromStorage() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      var data = JSON.parse(raw);
      if (!data || !data.files) return null;
      data.files.forEach(function (f) {
        f.rows.forEach(function (r) {
          if (r.docDate) r.docDate = new Date(r.docDate);
          if (r.sentDate) r.sentDate = new Date(r.sentDate);
        });
      });
      return data;
    } catch (e) {
      return null;
    }
  }

  function applyFiles() {
    state.records = [];
    state.files.forEach(function (f) { state.records = state.records.concat(f.rows); });
    renderFileList();
    if (state.records.length) {
      renderDashboard();
    } else {
      dashboard.style.display = 'none';
    }
  }

  function rebuildRecords() {
    saveToStorage(); // sets state.lastLoadedAt -- must happen before applyFiles() renders it
    applyFiles();
  }

  function handleFiles(fileList) {
    var files = Array.prototype.slice.call(fileList).filter(function (f) {
      return /\.xlsx$/i.test(f.name);
    });
    if (!files.length) {
      setStatus('รองรับเฉพาะไฟล์ .xlsx (ไฟล์ที่ Export จากหน้า Send to Buyer ของ getinvoice.net)', 'err');
      return;
    }
    setStatus('กำลังอ่านไฟล์...', 'info');
    var chain = Promise.resolve();
    files.forEach(function (file) {
      chain = chain.then(function () {
        return parseWorkbookFile(file).then(function (records) {
          state.files.push({ name: file.name, rows: records });
        });
      });
    });
    chain.then(function () {
      clearStatus();
      rebuildRecords();
    }).catch(function (err) {
      setStatus('เกิดข้อผิดพลาด: ' + err.message, 'err');
    });
  }

  dz.addEventListener('click', function (e) {
    if (e.target.tagName !== 'INPUT') fileInput.click();
  });
  fileInput.addEventListener('change', function () {
    if (fileInput.files.length) handleFiles(fileInput.files);
    fileInput.value = '';
  });
  ['dragover', 'dragenter'].forEach(function (ev) {
    dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('drag'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('drag'); });
  });
  dz.addEventListener('drop', function (e) {
    if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
  });

  function fmtNum(n) {
    return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function fmtInt(n) {
    return n.toLocaleString('en-US');
  }

  function computeCounts(records) {
    var counts = {};
    var amounts = {};
    BUCKET_DEFS.forEach(function (b) { counts[b.key] = 0; amounts[b.key] = 0; });
    records.forEach(function (r) {
      counts[r.bucket] = (counts[r.bucket] || 0) + 1;
      amounts[r.bucket] = (amounts[r.bucket] || 0) + r.total;
    });
    return { counts: counts, amounts: amounts };
  }

  function getDateRangeLabel(records) {
    var times = records.filter(function (r) { return r.docDate instanceof Date; }).map(function (r) { return r.docDate.getTime(); });
    if (!times.length) return null;
    return toDateLabel(new Date(Math.min.apply(null, times))) + ' – ' + toDateLabel(new Date(Math.max.apply(null, times)));
  }

  function sortNewestFirst(records) {
    return records.slice().sort(function (a, b) {
      var ta = a.docDate instanceof Date ? a.docDate.getTime() : -Infinity;
      var tb = b.docDate instanceof Date ? b.docDate.getTime() : -Infinity;
      return tb - ta;
    });
  }

  function downloadExcel(filteredSorted, agg, totalCount, totalAmount, rangeLabel) {
    var bucketLabelMap = {};
    BUCKET_DEFS.forEach(function (b) { bucketLabelMap[b.key] = b; });

    var summaryAoa = [
      ['Invoice Send Status Dashboard'],
      ['สร้างเมื่อ', new Date().toLocaleString('th-TH')],
      ['ช่วงข้อมูล (Document Date)', rangeLabel || '—'],
      ['ตัวกรองที่ใช้', 'หมวด: ' + (state.activeBucket === 'all' ? 'ทั้งหมด' : (bucketLabelMap[state.activeBucket] || {}).label) +
        ' | ช่วงวันที่: ' + (DATE_RANGE_DEFS.filter(function (d) { return d.key === state.dateRange; })[0] || {}).label +
        (state.searchText ? ' | ค้นหา: ' + state.searchText : '')],
      [],
      ['หมวด', 'จำนวน (ใบ)', 'ยอดรวม (บาท)'],
      ['ทั้งหมด', totalCount, totalAmount]
    ];
    BUCKET_DEFS.forEach(function (b) {
      if (b.key === 'unknown' && !agg.counts.unknown) return;
      summaryAoa.push([b.label, agg.counts[b.key] || 0, agg.amounts[b.key] || 0]);
    });

    var detailAoa = [['Document No', 'Type', 'Document Date', 'Buyer Name', 'Email Address', 'Total', 'Status', 'Action', 'หมวด']];
    filteredSorted.forEach(function (r) {
      var b = bucketLabelMap[r.bucket];
      detailAoa.push([r.docNo, r.type, toDateLabel(r.docDate), r.buyerName, r.email, r.total, r.status, r.action, b ? b.label : r.bucket]);
    });

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryAoa), 'สรุป');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(detailAoa), 'รายละเอียด');
    var now = new Date();
    var stamp = now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0') +
      '-' + String(now.getHours()).padStart(2, '0') + String(now.getMinutes()).padStart(2, '0');
    XLSX.writeFile(wb, 'invoice-status-dashboard-' + stamp + '.xlsx');
  }

  function renderDashboard() {
    dashboard.style.display = 'block';
    var records = filterByDateRange(state.records);
    var agg = computeCounts(records);
    var totalCount = records.length;
    var totalAmount = records.reduce(function (s, r) { return s + r.total; }, 0);

    var dateChipsHtml = '<div class="chip-row">' +
      DATE_RANGE_DEFS.map(function (d) {
        var active = state.dateRange === d.key ? ' active' : '';
        return '<button class="chip date-chip' + active + '" data-range="' + d.key + '">' + esc_(d.label) + '</button>';
      }).join('') +
      '</div>';
    var dateNote = records.length !== state.records.length ?
      '<div class="result-footer">กรองตาม Document Date: แสดง ' + fmtInt(records.length) + ' จาก ' + fmtInt(state.records.length) + ' แถวทั้งหมดที่อัปโหลด</div>' : '';

    var overallRangeLabel = getDateRangeLabel(state.records);
    var uploadedAtLabel = state.lastLoadedAt instanceof Date ? state.lastLoadedAt.toLocaleString('th-TH') : null;
    var rangeHtml = overallRangeLabel ?
      '<div class="data-range-note">📅 ข้อมูลใบแจ้งหนี้ (Document Date) ตั้งแต่ <b>' + esc_(overallRangeLabel) + '</b> — เรียงจากล่าสุดไปเก่าสุด' +
      (uploadedAtLabel ? '<br>📤 อัปโหลดไฟล์เข้าอ่านเมื่อ: <b>' + esc_(uploadedAtLabel) + '</b>' : '') +
      '</div>' : '';

    var statCardsHtml = '<div class="stat-row">' +
      '<button class="stat-box' + (state.activeBucket === 'all' ? ' active' : '') + '" data-bucket="all">' +
      '<div class="stat-num c-blue">' + fmtInt(totalCount) + '</div>' +
      '<div class="stat-label">ใบแจ้งหนี้ทั้งหมด<br>' + fmtNum(totalAmount) + ' บาท</div></button>' +
      BUCKET_DEFS.filter(function (b) { return b.key !== 'unknown' || agg.counts.unknown > 0; }).map(function (b) {
        return '<button class="stat-box' + (state.activeBucket === b.key ? ' active' : '') + '" data-bucket="' + b.key + '">' +
          '<div class="stat-num ' + b.cls + '">' + fmtInt(agg.counts[b.key]) + '</div>' +
          '<div class="stat-label">' + esc_(b.label) + '<br>' + fmtNum(agg.amounts[b.key]) + ' บาท</div></button>';
      }).join('') +
      '</div>';

    var searchHtml = '<input type="text" id="searchBox" class="search-box" placeholder="ค้นหา Document No / ชื่อผู้ซื้อ / อีเมล" value="' + esc_(state.searchText) + '">';

    var filtered = records.filter(function (r) {
      if (state.activeBucket !== 'all' && r.bucket !== state.activeBucket) return false;
      if (state.searchText) {
        var hay = (String(r.docNo) + ' ' + String(r.buyerName) + ' ' + String(r.email)).toLowerCase();
        if (hay.indexOf(state.searchText.toLowerCase()) === -1) return false;
      }
      return true;
    });
    filtered = sortNewestFirst(filtered);

    var bucketLabelMap = {};
    BUCKET_DEFS.forEach(function (b) { bucketLabelMap[b.key] = b; });

    var tableRowsHtml = filtered.slice(0, 500).map(function (r) {
      var b = bucketLabelMap[r.bucket];
      return '<tr><td>' + esc_(r.docNo) + '</td><td>' + esc_(r.type) + '</td>' +
        '<td>' + toDateLabel(r.docDate) + '</td>' +
        '<td class="wrap-cell">' + esc_(r.buyerName) + '</td>' +
        '<td class="wrap-cell">' + esc_(r.email) + '</td>' +
        '<td>' + fmtNum(r.total) + '</td>' +
        '<td>' + esc_(r.status) + '</td>' +
        '<td>' + esc_(r.action) + '</td>' +
        '<td><span class="badge ' + b.cls + '-bg">' + esc_(b.label) + '</span></td></tr>';
    }).join('');

    var moreNote = filtered.length > 500 ? '<div class="result-footer">แสดง 500 จาก ' + filtered.length + ' แถวที่ตรงเงื่อนไข — ใช้ช่องค้นหาเพื่อกรองเพิ่มเติม</div>' : '';

    var downloadHtml = '<button class="btn-main" id="btnDownloadExcel">⬇ ดาวน์โหลด Excel</button>';

    dashboard.innerHTML =
      rangeHtml +
      dateChipsHtml +
      dateNote +
      statCardsHtml +
      '<div class="search-row">' + searchHtml + downloadHtml + '</div>' +
      '<div class="detail-table-wrap"><table class="detail-table"><thead><tr>' +
      '<th>Document No</th><th>Type</th><th>Document Date</th><th>Buyer Name</th><th>Email Address</th>' +
      '<th>Total</th><th>Status</th><th>Action</th><th>หมวด</th></tr></thead><tbody>' +
      tableRowsHtml + '</tbody></table></div>' + moreNote;

    var btnDownload = document.getElementById('btnDownloadExcel');
    if (btnDownload) {
      btnDownload.addEventListener('click', function () {
        downloadExcel(filtered, agg, totalCount, totalAmount, overallRangeLabel);
      });
    }

    Array.prototype.forEach.call(dashboard.querySelectorAll('.date-chip'), function (btn) {
      btn.addEventListener('click', function () {
        state.dateRange = btn.getAttribute('data-range');
        renderDashboard();
      });
    });
    Array.prototype.forEach.call(dashboard.querySelectorAll('.stat-box[data-bucket]'), function (btn) {
      btn.addEventListener('click', function () {
        state.activeBucket = btn.getAttribute('data-bucket');
        renderDashboard();
      });
    });
    var searchBox = document.getElementById('searchBox');
    if (searchBox) {
      searchBox.addEventListener('input', function () {
        state.searchText = searchBox.value;
        renderDashboard();
      });
      // restore focus/cursor after re-render
      searchBox.focus();
      searchBox.setSelectionRange(searchBox.value.length, searchBox.value.length);
    }
  }

  // ---- Restore last-uploaded data on load, so the dashboard always shows something ----
  // until the user uploads a newer file (persisted client-side only, via localStorage).
  (function initFromStorage() {
    var persisted = loadFromStorage();
    if (!persisted || !persisted.files || !persisted.files.length) return;
    state.files = persisted.files;
    state.lastLoadedAt = persisted.savedAt ? new Date(persisted.savedAt) : null;
    applyFiles();
    if (persisted.savedAt) {
      setStatus('แสดงข้อมูลล่าสุดที่เคยอัปโหลดไว้ (บันทึกเมื่อ ' + new Date(persisted.savedAt).toLocaleString('th-TH') + ') — อัปโหลดไฟล์ใหม่เพื่ออัปเดตข้อมูล', 'info');
    }
  })();
})();
