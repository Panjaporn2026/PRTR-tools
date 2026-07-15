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

  // Extra buckets that only appear once an Invoice list.xlsx is also uploaded, cross-referencing
  // its "Send email by" column against the getinvoice.net export (matched by Document No).
  var ENRICHED_BUCKET_DEFS = [
    { key: 'get_invoice_missing', label: 'Get-invoice — ไม่พบข้อมูลในระบบ', cls: 'c-red' },
    { key: 'outlook_manual', label: 'ต้องส่งทาง Outlook (ไม่มีระบบตรวจสอบอัตโนมัติ)', cls: 'c-teal' },
    { key: 'no_need_email', label: 'ไม่ต้องส่งอีเมล', cls: 'c-gray' },
    { key: 'unknown_channel', label: 'ไม่ทราบช่องทาง (ตรวจสอบ)', cls: 'c-gray' }
  ];

  var state = {
    files: [],       // [{name, rows:[record]}] -- getinvoice.net Send to Buyer exports
    records: [],      // merged flat records (from files)
    invoiceListFiles: [],   // [{name, rows:[record]}] -- Invoice list.xlsx (optional, master list incl. Outlook)
    invoiceListRecords: [], // merged flat records (from invoiceListFiles)
    enrichedRecords: [],    // cross-referenced records shown when invoiceListRecords is non-empty
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

  var INVOICE_LIST_REQUIRED_HEADERS = ['Inv No./ CN No.', 'Send email by'];

  // "Get-invoice" / "Get invoice" (typo variant seen in real data) both -> get_invoice.
  // "Outlook" / "Outlook (HR onsite)" both -> outlook (no per-recipient tracking exists for either).
  // Anything else (blank, unrecognized text) -> unknown_channel, surfaced rather than silently guessed.
  function classifyChannel(raw) {
    var s = (raw == null ? '' : String(raw)).trim();
    var compact = s.replace(/[-\s]/g, '').toLowerCase();
    if (compact === 'getinvoice') return 'get_invoice';
    if (/^outlook/i.test(s)) return 'outlook';
    if (/no need email/i.test(s)) return 'no_need';
    return 'unknown';
  }

  function normalizeDocNo(v) {
    return v == null ? '' : String(v).trim();
  }

  function parseInvoiceListFile(file) {
    return file.arrayBuffer().then(function (buf) {
      var wb = XLSX.read(buf, { type: 'array', cellDates: true });
      var records = [];
      wb.SheetNames.forEach(function (sheetName) {
        var ws = wb.Sheets[sheetName];
        var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
        var headerRowIdx;
        try {
          headerRowIdx = findHeaderRow(aoa, INVOICE_LIST_REQUIRED_HEADERS);
        } catch (e) {
          return; // not the Invoice list sheet; skip silently, only fail if NO sheet matches
        }
        var headerRow = aoa[headerRowIdx];
        var idx = {
          docType: colIndex(headerRow, 'Doc Type'),
          docNo: colIndex(headerRow, 'Inv No./ CN No.'),
          project: colIndex(headerRow, 'Project'),
          customerName: colIndex(headerRow, 'Customer Name'),
          docDate: colIndex(headerRow, 'Doc Date'),
          dueDate: colIndex(headerRow, 'Due Date'),
          invoiceAmt: colIndex(headerRow, 'Invoice Amt'),
          referenceNo: colIndex(headerRow, 'Reference No.'),
          sendEmailBy: colIndex(headerRow, 'Send email by')
        };
        for (var r = headerRowIdx + 1; r < aoa.length; r++) {
          var row = aoa[r];
          if (!row || row[idx.docNo] == null || row[idx.docNo] === '') continue;
          var sendEmailByRaw = idx.sendEmailBy >= 0 ? row[idx.sendEmailBy] : null;
          records.push({
            docNo: row[idx.docNo],
            docType: idx.docType >= 0 ? row[idx.docType] : '',
            project: idx.project >= 0 ? row[idx.project] : '',
            customerName: idx.customerName >= 0 ? row[idx.customerName] : '',
            docDate: idx.docDate >= 0 ? normalizeToUTCDay(row[idx.docDate]) : null,
            dueDate: idx.dueDate >= 0 ? row[idx.dueDate] : null,
            invoiceAmt: idx.invoiceAmt >= 0 ? toNumber(row[idx.invoiceAmt]) : 0,
            referenceNo: idx.referenceNo >= 0 ? row[idx.referenceNo] : '',
            sendEmailByRaw: sendEmailByRaw,
            channel: classifyChannel(sendEmailByRaw),
            sourceFile: file.name
          });
        }
      });
      if (records.length === 0) {
        throw new Error('ไม่พบแถวข้อมูลที่ใช้ได้ในไฟล์ "' + file.name + '" — ตรวจสอบว่าเป็นไฟล์ Invoice list.xlsx ที่ถูกต้อง (ต้องมีคอลัมน์ Inv No./ CN No. และ Send email by)');
      }
      return records;
    });
  }

  // Cross-references Invoice list rows (the master/complete set, incl. Outlook) against the
  // getinvoice.net records (matched by Document No) to produce one unified row per Invoice list
  // entry. Only runs once an Invoice list file is uploaded; otherwise the plain getinvoice.net-only
  // view (state.records) is used as-is.
  function buildEnrichedRecords(getinvoiceRecords, invoiceListRecords) {
    var byDocNo = {};
    getinvoiceRecords.forEach(function (r) { byDocNo[normalizeDocNo(r.docNo)] = r; });
    return invoiceListRecords.map(function (r) {
      var matched = r.channel === 'get_invoice' ? byDocNo[normalizeDocNo(r.docNo)] : null;
      var bucket, status, action, email;
      if (r.channel === 'get_invoice') {
        if (matched) {
          bucket = matched.bucket; status = matched.status; action = matched.action; email = matched.email;
        } else {
          bucket = 'get_invoice_missing'; status = '(ไม่พบใน Get-invoice)'; action = ''; email = '';
        }
      } else if (r.channel === 'outlook') {
        bucket = 'outlook_manual'; status = r.sendEmailByRaw; action = ''; email = '';
      } else if (r.channel === 'no_need') {
        bucket = 'no_need_email'; status = r.sendEmailByRaw; action = ''; email = '';
      } else {
        bucket = 'unknown_channel'; status = r.sendEmailByRaw; action = ''; email = '';
      }
      return {
        docNo: r.docNo, type: r.docType, docDate: r.docDate, orderNo: r.referenceNo,
        total: r.invoiceAmt, buyerName: r.customerName, email: email, status: status, action: action,
        project: r.project, sendEmailBy: r.sendEmailByRaw, bucket: bucket, sourceFile: r.sourceFile
      };
    });
  }

  function currentBucketDefs() {
    return state.invoiceListRecords.length ? BUCKET_DEFS.concat(ENRICHED_BUCKET_DEFS) : BUCKET_DEFS;
  }

  // ---- UI wiring ----
  var dz = document.getElementById('dropzone');
  var fileInput = document.getElementById('fileInput');
  var fileListEl = document.getElementById('fileList');
  var dz2 = document.getElementById('dropzone2');
  var fileInput2 = document.getElementById('fileInput2');
  var fileListEl2 = document.getElementById('fileList2');
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

  function renderInvoiceListFileList() {
    if (!state.invoiceListFiles.length) { fileListEl2.innerHTML = ''; return; }
    fileListEl2.innerHTML = state.invoiceListFiles.map(function (f, i) {
      return '<div class="file-row"><span class="idx">' + (i + 1) + '</span>' +
        '<span class="fname">' + esc_(f.name) + ' (' + f.rows.length + ' แถว)</span>' +
        '<button class="rm" data-i="' + i + '" title="ลบไฟล์">✕</button></div>';
    }).join('');
    Array.prototype.forEach.call(fileListEl2.querySelectorAll('.rm'), function (btn) {
      btn.addEventListener('click', function () {
        state.invoiceListFiles.splice(parseInt(btn.getAttribute('data-i'), 10), 1);
        rebuildRecords();
      });
    });
  }

  function saveToStorage() {
    state.lastLoadedAt = new Date();
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        files: state.files,
        invoiceListFiles: state.invoiceListFiles,
        savedAt: state.lastLoadedAt.toISOString()
      }));
    } catch (e) {
      // localStorage full/blocked (private mode) -- non-critical, just skip persistence
    }
  }

  function reviveDates(files) {
    files.forEach(function (f) {
      f.rows.forEach(function (r) {
        if (r.docDate) r.docDate = new Date(r.docDate);
        if (r.sentDate) r.sentDate = new Date(r.sentDate);
        if (r.dueDate) r.dueDate = new Date(r.dueDate);
      });
    });
  }

  function loadFromStorage() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      var data = JSON.parse(raw);
      if (!data || !data.files) return null;
      reviveDates(data.files);
      if (data.invoiceListFiles) reviveDates(data.invoiceListFiles);
      return data;
    } catch (e) {
      return null;
    }
  }

  function applyFiles() {
    state.records = [];
    state.files.forEach(function (f) { state.records = state.records.concat(f.rows); });
    state.invoiceListRecords = [];
    state.invoiceListFiles.forEach(function (f) { state.invoiceListRecords = state.invoiceListRecords.concat(f.rows); });
    state.enrichedRecords = state.invoiceListRecords.length ? buildEnrichedRecords(state.records, state.invoiceListRecords) : [];
    renderFileList();
    renderInvoiceListFileList();
    if (state.records.length || state.invoiceListRecords.length) {
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

  function handleInvoiceListFiles(fileList) {
    var files = Array.prototype.slice.call(fileList).filter(function (f) {
      return /\.xlsx$/i.test(f.name);
    });
    if (!files.length) {
      setStatus('รองรับเฉพาะไฟล์ .xlsx (ไฟล์ Invoice list.xlsx)', 'err');
      return;
    }
    setStatus('กำลังอ่านไฟล์ Invoice list...', 'info');
    var chain = Promise.resolve();
    files.forEach(function (file) {
      chain = chain.then(function () {
        return parseInvoiceListFile(file).then(function (records) {
          state.invoiceListFiles.push({ name: file.name, rows: records });
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

  dz2.addEventListener('click', function (e) {
    if (e.target.tagName !== 'INPUT') fileInput2.click();
  });
  fileInput2.addEventListener('change', function () {
    if (fileInput2.files.length) handleInvoiceListFiles(fileInput2.files);
    fileInput2.value = '';
  });
  ['dragover', 'dragenter'].forEach(function (ev) {
    dz2.addEventListener(ev, function (e) { e.preventDefault(); dz2.classList.add('drag'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    dz2.addEventListener(ev, function (e) { e.preventDefault(); dz2.classList.remove('drag'); });
  });
  dz2.addEventListener('drop', function (e) {
    if (e.dataTransfer.files.length) handleInvoiceListFiles(e.dataTransfer.files);
  });

  function fmtNum(n) {
    return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function fmtInt(n) {
    return n.toLocaleString('en-US');
  }

  function computeCounts(records, bucketDefs) {
    var counts = {};
    var amounts = {};
    bucketDefs.forEach(function (b) { counts[b.key] = 0; amounts[b.key] = 0; });
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

  function downloadExcel(filteredSorted, agg, totalCount, totalAmount, rangeLabel, bucketDefs, enrichedMode) {
    var bucketLabelMap = {};
    bucketDefs.forEach(function (b) { bucketLabelMap[b.key] = b; });

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
    bucketDefs.forEach(function (b) {
      if ((b.key === 'unknown' || b.key === 'unknown_channel') && !agg.counts[b.key]) return;
      summaryAoa.push([b.label, agg.counts[b.key] || 0, agg.amounts[b.key] || 0]);
    });

    var detailHeader = enrichedMode ?
      ['Document No', 'Type', 'Project', 'Document Date', 'Buyer Name', 'Email Address', 'Total', 'Status', 'Action', 'Send email by', 'หมวด'] :
      ['Document No', 'Type', 'Document Date', 'Buyer Name', 'Email Address', 'Total', 'Status', 'Action', 'หมวด'];
    var detailAoa = [detailHeader];
    filteredSorted.forEach(function (r) {
      var b = bucketLabelMap[r.bucket];
      var row = enrichedMode ?
        [r.docNo, r.type, r.project, toDateLabel(r.docDate), r.buyerName, r.email, r.total, r.status, r.action, r.sendEmailBy, b ? b.label : r.bucket] :
        [r.docNo, r.type, toDateLabel(r.docDate), r.buyerName, r.email, r.total, r.status, r.action, b ? b.label : r.bucket];
      detailAoa.push(row);
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
    var enrichedMode = state.invoiceListRecords.length > 0;
    var baseRecords = enrichedMode ? state.enrichedRecords : state.records;
    var bucketDefs = currentBucketDefs();
    var records = filterByDateRange(baseRecords);
    var agg = computeCounts(records, bucketDefs);
    var totalCount = records.length;
    var totalAmount = records.reduce(function (s, r) { return s + r.total; }, 0);

    var dateChipsHtml = '<div class="chip-row">' +
      DATE_RANGE_DEFS.map(function (d) {
        var active = state.dateRange === d.key ? ' active' : '';
        return '<button class="chip date-chip' + active + '" data-range="' + d.key + '">' + esc_(d.label) + '</button>';
      }).join('') +
      '</div>';
    var dateNote = records.length !== baseRecords.length ?
      '<div class="result-footer">กรองตาม Document Date: แสดง ' + fmtInt(records.length) + ' จาก ' + fmtInt(baseRecords.length) + ' แถวทั้งหมดที่อัปโหลด</div>' : '';

    var overallRangeLabel = getDateRangeLabel(baseRecords);
    var uploadedAtLabel = state.lastLoadedAt instanceof Date ? state.lastLoadedAt.toLocaleString('th-TH') : null;
    var modeNote = enrichedMode ?
      '<br>🔀 โหมดตรวจครบทุกช่องทาง (อ้างอิงจาก Invoice list.xlsx) — จับคู่กับไฟล์ Send to Buyer ด้วย Document No' : '';
    var rangeHtml = overallRangeLabel ?
      '<div class="data-range-note">📅 ข้อมูลใบแจ้งหนี้ (Document Date) ตั้งแต่ <b>' + esc_(overallRangeLabel) + '</b> — เรียงจากล่าสุดไปเก่าสุด' +
      (uploadedAtLabel ? '<br>📤 อัปโหลดไฟล์เข้าอ่านเมื่อ: <b>' + esc_(uploadedAtLabel) + '</b>' : '') +
      modeNote +
      '</div>' : '';

    var statCardsHtml = '<div class="stat-row">' +
      '<button class="stat-box' + (state.activeBucket === 'all' ? ' active' : '') + '" data-bucket="all">' +
      '<div class="stat-num c-blue">' + fmtInt(totalCount) + '</div>' +
      '<div class="stat-label">ใบแจ้งหนี้ทั้งหมด<br>' + fmtNum(totalAmount) + ' บาท</div></button>' +
      bucketDefs.filter(function (b) { return (b.key !== 'unknown' && b.key !== 'unknown_channel') || agg.counts[b.key] > 0; }).map(function (b) {
        return '<button class="stat-box' + (state.activeBucket === b.key ? ' active' : '') + '" data-bucket="' + b.key + '">' +
          '<div class="stat-num ' + b.cls + '">' + fmtInt(agg.counts[b.key]) + '</div>' +
          '<div class="stat-label">' + esc_(b.label) + '<br>' + fmtNum(agg.amounts[b.key]) + ' บาท</div></button>';
      }).join('') +
      '</div>';

    var searchPlaceholder = enrichedMode ?
      'ค้นหา Document No / Project / ชื่อผู้ซื้อ / อีเมล' :
      'ค้นหา Document No / ชื่อผู้ซื้อ / อีเมล';
    var searchHtml = '<input type="text" id="searchBox" class="search-box" placeholder="' + esc_(searchPlaceholder) + '" value="' + esc_(state.searchText) + '">';

    var filtered = records.filter(function (r) {
      if (state.activeBucket !== 'all' && r.bucket !== state.activeBucket) return false;
      if (state.searchText) {
        var hay = (String(r.docNo) + ' ' + String(r.buyerName) + ' ' + String(r.email) + ' ' + String(r.project || '')).toLowerCase();
        if (hay.indexOf(state.searchText.toLowerCase()) === -1) return false;
      }
      return true;
    });
    filtered = sortNewestFirst(filtered);

    var bucketLabelMap = {};
    bucketDefs.forEach(function (b) { bucketLabelMap[b.key] = b; });

    var tableRowsHtml = filtered.slice(0, 500).map(function (r) {
      var b = bucketLabelMap[r.bucket];
      return '<tr><td>' + esc_(r.docNo) + '</td><td>' + esc_(r.type) + '</td>' +
        (enrichedMode ? '<td class="wrap-cell">' + esc_(r.project) + '</td>' : '') +
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

    var tableHeaderHtml = '<th>Document No</th><th>Type</th>' +
      (enrichedMode ? '<th>Project</th>' : '') +
      '<th>Document Date</th><th>Buyer Name</th><th>Email Address</th>' +
      '<th>Total</th><th>Status</th><th>Action</th><th>หมวด</th>';

    dashboard.innerHTML =
      rangeHtml +
      dateChipsHtml +
      dateNote +
      statCardsHtml +
      '<div class="search-row">' + searchHtml + downloadHtml + '</div>' +
      '<div class="detail-table-wrap"><table class="detail-table"><thead><tr>' +
      tableHeaderHtml + '</tr></thead><tbody>' +
      tableRowsHtml + '</tbody></table></div>' + moreNote;

    var btnDownload = document.getElementById('btnDownloadExcel');
    if (btnDownload) {
      btnDownload.addEventListener('click', function () {
        downloadExcel(filtered, agg, totalCount, totalAmount, overallRangeLabel, bucketDefs, enrichedMode);
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
    if (!persisted || (!persisted.files || !persisted.files.length) && (!persisted.invoiceListFiles || !persisted.invoiceListFiles.length)) return;
    state.files = persisted.files || [];
    state.invoiceListFiles = persisted.invoiceListFiles || [];
    state.lastLoadedAt = persisted.savedAt ? new Date(persisted.savedAt) : null;
    applyFiles();
    if (persisted.savedAt) {
      setStatus('แสดงข้อมูลล่าสุดที่เคยอัปโหลดไว้ (บันทึกเมื่อ ' + new Date(persisted.savedAt).toLocaleString('th-TH') + ') — อัปโหลดไฟล์ใหม่เพื่ออัปเดตข้อมูล', 'info');
    }
  })();
})();
