(function () {
  'use strict';

  var REQUIRED_HEADERS = ['Document No', 'Status', 'Action'];
  var BUCKET_DEFS = [
    { key: 'error', label: 'ส่งแล้ว Error / ไม่สำเร็จ', cls: 'c-red' },
    { key: 'not_sent', label: 'ยังไม่ได้ส่ง', cls: 'c-orange' },
    { key: 'sent_opened', label: 'ส่งแล้ว — เปิดอ่านแล้ว', cls: 'c-green' },
    { key: 'sent_not_opened', label: 'ส่งแล้ว — ยังไม่เปิดอ่าน', cls: 'c-blue' },
    { key: 'unknown', label: 'ไม่ทราบสถานะ (ตรวจสอบ)', cls: 'c-gray' }
  ];
  var NOT_SENT_STATUSES = ["In Progress", "Don't Send", "Not Send"];
  var DATE_RANGE_DEFS = [
    { key: 'all', label: 'ทั้งหมด' },
    { key: 'today', label: 'วันนี้', days: 0 },
    { key: '7d', label: '7 วันล่าสุด', days: 7 },
    { key: '1m', label: '1 เดือนล่าสุด', days: 30 },
    { key: '2m', label: '2 เดือนล่าสุด', days: 60 }
  ];

  var state = {
    files: [],       // [{name, rows:[record]}]
    records: [],      // merged flat records
    activeBucket: 'all',
    searchText: '',
    dateRange: 'all'
  };

  function startOfDay(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function filterByDateRange(records) {
    var def = DATE_RANGE_DEFS.filter(function (d) { return d.key === state.dateRange; })[0];
    if (!def || def.days == null) return records;
    var cutoff = startOfDay(new Date());
    cutoff.setDate(cutoff.getDate() - def.days);
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

  function toNumber(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    var n = parseFloat(String(v).replace(/,/g, ''));
    return isNaN(n) ? 0 : n;
  }

  function toDateLabel(v) {
    if (v == null || v === '') return '—';
    if (v instanceof Date) {
      var d = v.getDate(), m = v.getMonth() + 1, y = v.getFullYear();
      return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
    }
    return String(v);
  }

  function classify(status, action) {
    var st = (status == null ? '' : String(status)).trim();
    var tokens = (action == null ? '' : String(action)).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    var hasBounce = tokens.indexOf('bounce') !== -1;
    var hasOpen = tokens.indexOf('open') !== -1;

    if (st === 'Failure' || hasBounce) return 'error';
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
          records.push({
            docNo: row[idx.docNo],
            type: idx.type >= 0 ? row[idx.type] : '',
            docDate: idx.docDate >= 0 ? row[idx.docDate] : null,
            orderNo: idx.orderNo >= 0 ? row[idx.orderNo] : '',
            total: idx.total >= 0 ? toNumber(row[idx.total]) : 0,
            buyerName: idx.buyerName >= 0 ? row[idx.buyerName] : '',
            email: idx.email >= 0 ? row[idx.email] : '',
            status: status,
            sentDate: idx.sentDate >= 0 ? row[idx.sentDate] : null,
            action: action,
            error: idx.error >= 0 ? row[idx.error] : '',
            bucket: classify(status, action),
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

  function rebuildRecords() {
    state.records = [];
    state.files.forEach(function (f) { state.records = state.records.concat(f.rows); });
    renderFileList();
    if (state.records.length) {
      renderDashboard();
    } else {
      dashboard.style.display = 'none';
    }
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

    var statCardsHtml = '<div class="stat-row">' +
      '<div class="stat-box"><div class="stat-num c-blue">' + fmtInt(totalCount) + '</div>' +
      '<div class="stat-label">ใบแจ้งหนี้ทั้งหมด<br>' + fmtNum(totalAmount) + ' บาท</div></div>' +
      BUCKET_DEFS.filter(function (b) { return b.key !== 'unknown' || agg.counts.unknown > 0; }).map(function (b) {
        return '<div class="stat-box"><div class="stat-num ' + b.cls + '">' + fmtInt(agg.counts[b.key]) + '</div>' +
          '<div class="stat-label">' + esc_(b.label) + '<br>' + fmtNum(agg.amounts[b.key]) + ' บาท</div></div>';
      }).join('') +
      '</div>';

    var chipsHtml = '<div class="chip-row">' +
      ['all'].concat(BUCKET_DEFS.map(function (b) { return b.key; }))
        .filter(function (key) { return key === 'all' || key !== 'unknown' || agg.counts.unknown > 0; })
        .map(function (key) {
          var label = key === 'all' ? 'ทั้งหมด (' + totalCount + ')' :
            (function () { var b = BUCKET_DEFS.filter(function (x) { return x.key === key; })[0]; return b.label + ' (' + agg.counts[key] + ')'; })();
          var active = state.activeBucket === key ? ' active' : '';
          return '<button class="chip' + active + '" data-bucket="' + key + '">' + esc_(label) + '</button>';
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

    dashboard.innerHTML =
      dateChipsHtml +
      dateNote +
      statCardsHtml +
      chipsHtml +
      '<div class="search-row">' + searchHtml + '</div>' +
      '<div class="detail-table-wrap"><table class="detail-table"><thead><tr>' +
      '<th>Document No</th><th>Type</th><th>Document Date</th><th>Buyer Name</th><th>Email Address</th>' +
      '<th>Total</th><th>Status</th><th>Action</th><th>หมวด</th></tr></thead><tbody>' +
      tableRowsHtml + '</tbody></table></div>' + moreNote;

    Array.prototype.forEach.call(dashboard.querySelectorAll('.date-chip'), function (btn) {
      btn.addEventListener('click', function () {
        state.dateRange = btn.getAttribute('data-range');
        renderDashboard();
      });
    });
    Array.prototype.forEach.call(dashboard.querySelectorAll('.chip:not(.date-chip)'), function (btn) {
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
})();
