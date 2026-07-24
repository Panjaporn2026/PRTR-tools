// ══════════════════════════════════════════════════════
//  Kick-off Meeting Note filler — client-side form that writes into the "Master for inovice"
//  sheet of a "★KICK OFF MEETING NOTE - <Client>.xlsx" workbook.
//
//  This workbook's dropdowns are Excel 2010+ extended data validation (x14:dataValidation inside
//  <extLst>). Any full parse-model-reserialize approach (like gl-invoice's row model) risks
//  dropping that block on rewrite, so this tool never reserializes the sheet at all -- it does
//  targeted string-level <c r="REF">...</c> replacement directly on the raw sheetXml, exactly
//  mirroring the project's own safe_fill_xlsx.py Claude Skill script, and leaves every other byte
//  of the file (styles, other sheets, the validation extList itself) completely untouched.
//
//  Row numbers are never hardcoded: every field is looked up by its own column-A label text at
//  load time (same "never guess" convention as gl-invoice's findHeaderRow), and every dropdown's
//  option list is read from the actual x14:dataValidation formula1 range declared in THIS file,
//  not a hardcoded range -- so a future template revision that shifts rows still works.
// ══════════════════════════════════════════════════════

var SHEET_NAME = 'Master for inovice';
var DV_SHEET_NAME = 'Data validation';

function normText(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
}

// Column-C "how to fill this" instructional hints this template pre-fills next to a few B-column
// inputs (Sales/HR/Recruit owner email, address sub-fields). These are instructions aimed at the
// human, not real data, so a cell holding one of them counts as blank/fillable -- same list as the
// project's safe_fill_xlsx.py Claude Skill script, which hit this exact case first.
var HINT_TEXTS = [
  'ใส่ชื่อจริง + ชื่อเล่น sales',
  'ใส่ชื่อจริง + ชื่อเล่น hr',
  'ใส่ชื่อจริง + ชื่อเล่น recruit',
  'ใส่ถึงอำเภอ',
  'จังหวัด',
  'รหัสไปรษณีย์',
  'ประเทศ'
];

// Placeholder text this template uses for "nothing entered yet" -- a cell holding one of these
// counts as blank/fillable, never as a real answer to protect.
function isPlaceholderText(s) {
  var t = normText(s).toLowerCase();
  if (t === '' || t === '-') return true;
  if (/^-?\s*select\s*-?$/.test(t)) return true; // "- select -", "-select-", "- select-", etc.
  if (HINT_TEXTS.indexOf(t) !== -1) return true;
  return false;
}

// ── Column-A label search (row is never hardcoded) ────────────────────────────────────────────
function findRowByLabel(aoa, label) {
  var target = normText(label);
  var matches = [];
  for (var r = 0; r < aoa.length; r++) {
    var row = aoa[r];
    if (row && normText(row[0]) === target) matches.push(r + 1);
  }
  if (matches.length === 0) throw new Error('ไม่พบแถวที่มีข้อความ "' + label + '" ในคอลัมน์ A ของชีต "' + SHEET_NAME + '"');
  if (matches.length > 1) throw new Error('พบแถวที่มีข้อความ "' + label + '" มากกว่า 1 แถว (แถว ' + matches.join(', ') + ')');
  return matches[0];
}

// ── Parse x14:dataValidation blocks from the Master sheet's own raw XML ───────────────────────
// Returns { CELLREF: { sheetName, col, r1, r2 } }. When more than one validation targets the same
// cell (a real quirk seen in this template for one field), the LAST one declared wins.
function parseDataValidationMap(sheetXml) {
  var map = {};
  var blockRe = /<x14:dataValidation\b[\s\S]*?<\/x14:dataValidation>/g, bm;
  while ((bm = blockRe.exec(sheetXml)) !== null) {
    var block = bm[0];
    var fM = /<xm:f>([\s\S]*?)<\/xm:f>/.exec(block);
    var sqM = /<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(block);
    if (!fM || !sqM) continue;
    var rangeM = /^'?([^'!]+)'?!\$([A-Z]+)\$(\d+):\$[A-Z]+\$(\d+)$/.exec(fM[1].trim());
    if (!rangeM) continue; // formula1 isn't a plain range reference (e.g. an inline list) -- skip
    var info = { sheetName: rangeM[1], col: rangeM[2], r1: parseInt(rangeM[3], 10), r2: parseInt(rangeM[4], 10) };
    sqM[1].trim().split(/\s+/).forEach(function (ref) { map[ref] = info; });
  }
  return map;
}

function readDropdownOptions(dvAoa, info) {
  var colIdx = colLettersToIndex(info.col);
  var out = [];
  for (var r = info.r1; r <= info.r2; r++) {
    var row = dvAoa[r - 1];
    var v = row ? row[colIdx] : null;
    if (v != null && String(v).trim() !== '') out.push(String(v));
  }
  return out;
}

// ── Field definitions (grouped for the form) ───────────────────────────────────────────────────
// col: which column holds the value for this label's row ('B' unless noted). type drives the
// rendered input. dropdown: true means resolve options via parseDataValidationMap at load time.
var SECTIONS = [
  { title: '1. ข้อมูลทั่วไป', fields: [
    { key: 'date', label: 'Date of Kick Off Meeting', col: 'B', type: 'text', placeholder: 'DD-MM-YY' },
    { key: 'sap', label: 'SAP Name (Maximum 20 Digits)', col: 'B', type: 'text', hint: 'ใช้เป็นชื่อไฟล์ output ด้วย (★KICK OFF MEETING NOTE - <SAP Name>.xlsx)' },
    { key: 'crmLink', label: 'PRTR Link (Back Office CRM)', col: 'B', type: 'text' },
    { key: 'category', label: 'Project Categories (Y26)', col: 'B', type: 'select', dropdown: true },
    { key: 'pic', label: 'OS Invoice PIC', col: 'B', type: 'select', dropdown: true },
    { key: 'salesTeam', label: 'Sales Team', col: 'B', type: 'select', dropdown: true },
    { key: 'salesName', label: 'Sales Manager', col: 'B', type: 'text', overrideLabel: 'Sales Manager (ชื่อจริง + ชื่อเล่น)' },
    { key: 'salesEmail', label: 'Sales Manager', col: 'C', type: 'text', overrideLabel: 'Sales Manager Email' },
    { key: 'hrTeam', label: 'HR Team', col: 'B', type: 'select', dropdown: true },
    { key: 'hrName', label: 'HR Project Owner', col: 'B', type: 'text', overrideLabel: 'HR Project Owner (ชื่อจริง + ชื่อเล่น)' },
    { key: 'hrEmail', label: 'HR Project Owner', col: 'C', type: 'text', overrideLabel: 'HR Project Owner Email' },
    { key: 'recruitName', label: 'Recruitment Project Owner', col: 'B', type: 'text', overrideLabel: 'Recruitment Project Owner (ชื่อจริง + ชื่อเล่น)' },
    { key: 'recruitEmail', label: 'Recruitment Project Owner', col: 'C', type: 'text', overrideLabel: 'Recruitment Project Owner Email' }
  ] },
  { title: '2. ข้อมูลบริษัทลูกค้า', fields: [
    { key: 'companyName', label: 'Company Name', col: 'B', type: 'text' },
    { key: 'group', label: 'Group', col: 'B', type: 'select', dropdown: true },
    { key: 'industry', label: 'Industry', col: 'B', type: 'select', dropdown: true },
    { key: 'taxId', label: 'Tax ID', col: 'B', type: 'text' },
    { key: 'branch', label: 'Branch', col: 'B', type: 'select', dropdown: true },
    { key: 'taxBranch', label: 'Tax Branch', col: 'B', type: 'select', dropdown: true },
    { key: 'addr1', label: '- Building/Floor/Room', col: 'B', type: 'text' },
    { key: 'addr2', label: '- City', col: 'B', type: 'text' },
    { key: 'addr3', label: '- Zip Code', col: 'B', type: 'text' },
    { key: 'bank', label: 'Payment Run - Bank', col: 'B', type: 'select', dropdown: true },
    { key: 'account', label: 'Payment Run - Account Number', col: 'B', type: 'select', dropdown: true },
    { key: 'doiTemplate', label: 'Detail of Invoice Teamplate', col: 'B', type: 'select', dropdown: true },
    { key: 'invoiceFormat', label: 'Invoice Format', col: 'B', type: 'select', dropdown: true },
    { key: 'sendEmailBy', label: 'Submit Invoice via Email By', col: 'B', type: 'select', dropdown: true },
    { key: 'submitBySystem', label: 'Submit Invoice by System', col: 'B', type: 'select', dropdown: true },
    { key: 'sendOriginalBy', label: 'Original Send by (If Client Need Original)', col: 'B', type: 'select', dropdown: true },
    { key: 'billingAddr', label: 'Billing Delivery Address (If Client Need Original)', col: 'B', type: 'textarea', hint: 'กรอกเฉพาะกรณี "Original Send by" ต้องส่งเอกสารจริง' }
  ] },
  { title: '3. ผู้ติดต่อสำหรับใบแจ้งหนี้ (เห็นเงินเดือนพนักงาน)', fields: [
    { key: 'contactFirst', label: '- First Name', col: 'B', type: 'text', hint: 'เก็บคำนำหน้าอย่าง "Khun" ไว้ในชื่อจริงด้วย เช่น "Khun Sasivan"' },
    { key: 'contactLast', label: '- Last Name', col: 'B', type: 'text' },
    { key: 'contactPosition', label: '- Position', col: 'B', type: 'text' },
    { key: 'contactEmail', label: '- Email', col: 'B', type: 'text' },
    { key: 'contactMobile', label: '- Mobile Phone No.', col: 'B', type: 'text' },
    { key: 'contactPhone', label: '- Company Phone No.', col: 'B', type: 'text', placeholder: '- ถ้าไม่มี' }
  ] },
  { title: '4. Payroll Details', fields: [
    { key: 'staffType', label: "Staff's Type", col: 'B', type: 'text' },
    { key: 'salaryCutoff', label: 'Salary Cut Off', col: 'B', type: 'text' },
    { key: 'salaryPayDate', label: 'Salary Pay Date', col: 'B', type: 'text' },
    { key: 'variableCutoff', label: 'Variable Cut Off', col: 'B', type: 'text' },
    { key: 'variablePayDate', label: 'Variable Pay Date', col: 'B', type: 'text' }
  ] },
  { title: '5. B2B Details', fields: [
    { key: 'b2bNo', label: 'B2B Contract No.', col: 'B', type: 'text' },
    { key: 'b2bStart', label: 'B2B Start Date', col: 'B', type: 'text', placeholder: 'DD-MM-YY', id: 'b2bStart' },
    { key: 'b2bEnd', label: 'B2B End Date', col: 'B', type: 'text', placeholder: 'DD-MM-YY' },
    { key: 'b2bStatus', label: 'Status B2B', col: 'B', type: 'select', dropdown: true },
    { key: 'b2bRemark', label: 'Remark for B2B', col: 'B', type: 'textarea', labelPrefix: true },
    { key: 'feeCondition', label: 'Service Fee Conditions', col: 'B', type: 'select', dropdown: true },
    { key: 'ssfCost', label: 'SSF Cost', col: 'B', type: 'select', dropdown: true },
    { key: 'minFee', label: 'Minimum Fee', col: 'B', type: 'text' },
    { key: 'implementFee', label: 'Implement Fee', col: 'B', type: 'text', labelPrefix: true },
    { key: 'monthlyFee', label: 'Monthly Fee', col: 'B', type: 'text' }
  ] },
  { title: '6. Invoice Submission Process', fields: [
    { key: 'invoiceType', label: 'Type of Invoice', col: 'B', type: 'select', dropdown: true },
    { key: 'deposit', label: 'Deposit', col: 'B', type: 'select', dropdown: true },
    { key: 'depositDetail', label: 'Deposit Amount / Date / Due Date', col: 'B', type: 'text' },
    { key: 'startDate', label: 'Start Date', col: 'B', type: 'text', placeholder: 'DD-MM-YY', id: 'startDate', readonly: true, hint: 'คำนวณอัตโนมัติ = วันที่ 1 ของเดือนถัดจาก B2B Start Date' },
    { key: 'marginMonth', label: 'Margin Month Start', col: 'B', type: 'text', placeholder: 'MMM-YY', id: 'marginMonth', readonly: true, hint: 'คำนวณอัตโนมัติ = เดือนเดียวกับ Start Date ด้านบน' },
    { key: 'paymentTerms', label: 'Payment Terms (From B2B)', col: 'B', type: 'textarea', labelPrefix: true },
    { key: 'firstInvoiceDate', label: '1st Invoice Date', col: 'B', type: 'text', placeholder: 'DD-MM-YY' },
    { key: 'paymentDays', label: 'Payment Terms (Days)', col: 'B', type: 'select', dropdown: true },
    { key: 'firstInvoiceDue', label: '1st Invoice Due Date', col: 'B', type: 'text', placeholder: 'DD-MM-YY' },
    { key: 'clientCycle', label: 'Client Payment Cycle', col: 'B', type: 'textarea', labelPrefix: true },
    { key: 'requiredVendor', label: 'Required Register Vendor', col: 'B', type: 'text', labelPrefix: true },
    { key: 'requiredPo', label: 'Required P/O', col: 'B', type: 'text', labelPrefix: true },
    { key: 'poTypeFreq', label: 'PO Type (Monthly/Quarterly/Yearly/Onetime)', col: 'B', type: 'text' },
    { key: 'poTypeBasis', label: 'PO Type (Actual/ Estimated)', col: 'B', type: 'text' },
    { key: 'specialInstructions', label: 'Special Instructions', col: 'B', type: 'textarea' }
  ] }
];

// labelPrefix fields: the actual column-A text has extra trailing content (Thai hint / "(Yes/No)"
// suffix etc.) that varies, so match by prefix instead of exact equality.
function findRowByLabelFlexible(aoa, def) {
  if (!def.labelPrefix) return findRowByLabel(aoa, def.label);
  var target = normText(def.label);
  var matches = [];
  for (var r = 0; r < aoa.length; r++) {
    var row = aoa[r];
    if (row && normText(row[0]).indexOf(target) === 0) matches.push(r + 1);
  }
  if (matches.length === 0) throw new Error('ไม่พบแถวที่ขึ้นต้นด้วย "' + def.label + '" ในคอลัมน์ A');
  if (matches.length > 1) throw new Error('พบแถวที่ขึ้นต้นด้วย "' + def.label + '" มากกว่า 1 แถว (แถว ' + matches.join(', ') + ')');
  return matches[0];
}

// ── State ───────────────────────────────────────────────────────────────────────────────────
var state = { file: null, wb: null, sheetXml: null, sst: null, aoa: null, dvAoa: null, dvMap: null, resolved: null };

function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function setStatus(msg, cls) {
  var el = document.getElementById('statusBox');
  el.textContent = msg || '';
  el.className = 'status-box' + (cls ? ' ' + cls : '');
  el.style.display = msg ? 'block' : 'none';
}

var dz = document.getElementById('dropzone');
var fileInput = document.getElementById('fileInput');
var formSection = document.getElementById('formSection');
var resultSection = document.getElementById('resultSection');

dz.addEventListener('click', function (e) { if (e.target.tagName !== 'INPUT') fileInput.click(); });
fileInput.addEventListener('change', function () { if (fileInput.files.length) handleFile(fileInput.files[0]); });
['dragover', 'dragenter'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('drag'); }); });
['dragleave', 'drop'].forEach(function (ev) { dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('drag'); }); });
dz.addEventListener('drop', function (e) { if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]); });

async function handleFile(file) {
  if (!/\.xlsx$/i.test(file.name)) { setStatus('รองรับเฉพาะไฟล์ .xlsx', 'err'); return; }
  setStatus('⏳ กำลังอ่านไฟล์...', 'info');
  formSection.style.display = 'none';
  resultSection.style.display = 'none';
  try {
    var buf = await file.arrayBuffer();
    var wb = await loadWorkbook(buf);
    if (!wb.sheets[SHEET_NAME]) throw new Error('ไม่พบชีต "' + SHEET_NAME + '" ในไฟล์นี้ — ตรวจสอบว่าเป็นไฟล์ ★KICK OFF MEETING NOTE ที่ถูกต้อง');
    if (!wb.sheets[DV_SHEET_NAME]) throw new Error('ไม่พบชีต "' + DV_SHEET_NAME + '" ในไฟล์นี้');
    var sheetXml = await getSheetXml(wb, SHEET_NAME);
    var sst = await getSharedStrings(wb);
    var aoa = parseGridFromXml(sheetXml, sst);
    var dvXml = await getSheetXml(wb, DV_SHEET_NAME);
    var dvAoa = parseGridFromXml(dvXml, sst);
    var dvMap = parseDataValidationMap(sheetXml);

    // Resolve every field's row + current value up front, so a missing label fails loudly now
    // rather than silently later.
    var resolved = {};
    SECTIONS.forEach(function (section) {
      section.fields.forEach(function (def) {
        var row = findRowByLabelFlexible(aoa, def);
        var cellRef = def.col + row;
        var colIdx = colLettersToIndex(def.col);
        var currentVal = (aoa[row - 1] || [])[colIdx];
        var options = null;
        if (def.dropdown) {
          var info = dvMap[cellRef];
          if (!info) throw new Error('ไม่พบ dropdown (data validation) สำหรับเซลล์ ' + cellRef + ' (' + def.label + ')');
          options = readDropdownOptions(dvAoa, info);
        }
        resolved[def.key] = { def: def, row: row, cellRef: cellRef, currentVal: currentVal, options: options };
      });
    });

    state = { file: file, wb: wb, sheetXml: sheetXml, sst: sst, aoa: aoa, dvAoa: dvAoa, dvMap: dvMap, resolved: resolved };
    renderForm();
    setStatus('✅ อ่านไฟล์สำเร็จ — กรอกข้อมูลด้านล่างแล้วกด "สร้างไฟล์"', 'ok');
  } catch (err) {
    setStatus('❌ ' + err.message, 'err');
  }
}

function renderForm() {
  var html = '';
  SECTIONS.forEach(function (section) {
    html += '<div class="form-section"><div class="form-section-title">' + esc_(section.title) + '</div><div class="form-grid">';
    section.fields.forEach(function (def) {
      var r = state.resolved[def.key];
      var label = esc_(def.overrideLabel || def.label);
      var already = !isPlaceholderText(r.currentVal) && r.currentVal != null && String(r.currentVal).trim() !== '';
      var idAttr = def.id ? ' id="fld_' + def.id + '"' : '';
      html += '<div class="form-field' + (def.type === 'textarea' ? ' wide' : '') + '">';
      html += '<label>' + label + ' <span class="cell-tag">' + r.cellRef + '</span>' +
        (already ? '<span class="badge-existing">มีข้อมูลอยู่แล้ว</span>' : '') + '</label>';
      if (def.hint) html += '<div class="field-hint">' + esc_(def.hint) + '</div>';
      if (def.type === 'select') {
        html += '<select data-key="' + def.key + '"' + idAttr + (def.readonly ? ' disabled' : '') + '>';
        html += '<option value="">— ไม่กรอก —</option>';
        (r.options || []).forEach(function (opt) {
          if (isPlaceholderText(opt)) return;
          // Never pre-select the sheet's existing value here -- an untouched <select> must submit
          // as "" (the blank sentinel option) exactly like an untouched text input does, so a
          // field the user never opened never gets reported as "skipped" noise below.
          html += '<option value="' + esc_(opt) + '">' + esc_(opt) + '</option>';
        });
        html += '</select>';
      } else if (def.type === 'textarea') {
        html += '<textarea data-key="' + def.key + '"' + idAttr + ' rows="3" placeholder="' + esc_(def.placeholder || '') + '">' +
          (already ? '' : '') + '</textarea>';
      } else {
        html += '<input type="text" data-key="' + def.key + '"' + idAttr + (def.readonly ? ' readonly' : '') +
          ' placeholder="' + esc_(def.placeholder || '') + '" value="">';
      }
      if (already) html += '<div class="field-current">ค่าปัจจุบัน: ' + esc_(r.currentVal) + '</div>';
      html += '</div>';
    });
    html += '</div></div>';
  });
  html += '<button class="btn-main" id="btnGenerate">⬇️ สร้างไฟล์</button>';
  formSection.innerHTML = html;
  formSection.style.display = 'block';

  var b2bStartEl = document.getElementById('fld_b2bStart');
  if (b2bStartEl) b2bStartEl.addEventListener('input', updateDerivedDates);
  document.getElementById('btnGenerate').addEventListener('click', generateFile);
}

var MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function updateDerivedDates() {
  var raw = document.getElementById('fld_b2bStart').value.trim();
  var m = /^(\d{1,2})-(\d{1,2})-(\d{2})$/.exec(raw);
  var startEl = document.getElementById('fld_startDate');
  var marginEl = document.getElementById('fld_marginMonth');
  if (!m) { startEl.value = ''; marginEl.value = ''; return; }
  var yy = parseInt(m[3], 10), mm = parseInt(m[2], 10);
  var fullYear = 2000 + yy;
  var nextMonthIdx = mm; // 0-based next month (mm is 1-based current month, so mm == next month's 0-based index)
  var nextYear = fullYear;
  if (nextMonthIdx > 11) { nextMonthIdx = 0; nextYear += 1; }
  var dd = '01';
  var mm2 = String(nextMonthIdx + 1).padStart(2, '0');
  var yy2 = String(nextYear % 100).padStart(2, '0');
  startEl.value = dd + '-' + mm2 + '-' + yy2;
  marginEl.value = MONTH_ABBR[nextMonthIdx] + '-' + yy2;
}

// ── Cell-level XML surgery (JS port of scripts/safe_fill_xlsx.py) ─────────────────────────────
function getCellText(sheetXml, ref, sst) {
  var selfM = new RegExp('<c r="' + ref + '"[^>]*/>').exec(sheetXml);
  if (selfM) return '';
  var fullM = new RegExp('<c r="' + ref + '"([^>]*)>([\\s\\S]*?)</c>').exec(sheetXml);
  if (!fullM) return null;
  var attrs = fullM[1], inner = fullM[2];
  if (inner.trim() === '') return '';
  var tM = /\bt="([^"]*)"/.exec(attrs);
  var t = tM ? tM[1] : null;
  if (t === 's') {
    var idxM = /<v>(\d+)<\/v>/.exec(inner);
    if (idxM) { var idx = parseInt(idxM[1], 10); return sst[idx] != null ? sst[idx] : ''; }
    return '';
  }
  if (t === 'inlineStr') {
    var texts = [], tRe = /<t[^>]*>([\s\S]*?)<\/t>/g, tm;
    while ((tm = tRe.exec(inner)) !== null) texts.push(tm[1]);
    return decodeXmlEntities(texts.join(''));
  }
  var vM = /<v>([\s\S]*?)<\/v>/.exec(inner);
  return vM ? vM[1] : inner;
}

function isFillableCell(sheetXml, ref, sst) {
  var text = getCellText(sheetXml, ref, sst);
  if (text === null) return true; // cell doesn't exist in the row's XML at all -> safe to add
  return isPlaceholderText(text);
}

function fillCellXml(sheetXml, ref, value) {
  var valEsc = esc_(value).replace(/&quot;/g, '&amp;quot;').replace(/&amp;quot;/g, '&quot;'); // esc_ already escapes &, keep quot simple
  valEsc = escXmlText(value);
  var selfRe = new RegExp('<c r="' + ref + '"([^>]*)/>');
  var fullRe = new RegExp('<c r="' + ref + '"([^>]*)>([\\s\\S]*?)</c>');
  function stripT(attrs) { return attrs.replace(/\s*t="[^"]*"/, ''); }
  if (selfRe.test(sheetXml)) {
    return sheetXml.replace(selfRe, function (_, attrs) {
      return '<c r="' + ref + '"' + stripT(attrs) + ' t="inlineStr"><is><t xml:space="preserve">' + valEsc + '</t></is></c>';
    });
  }
  if (fullRe.test(sheetXml)) {
    return sheetXml.replace(fullRe, function (_, attrs) {
      return '<c r="' + ref + '"' + stripT(attrs) + ' t="inlineStr"><is><t xml:space="preserve">' + valEsc + '</t></is></c>';
    });
  }
  return null; // cell ref not found in the sheet at all -- caller treats as skipped
}

function escXmlText(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ── Generate output ─────────────────────────────────────────────────────────────────────────
async function generateFile() {
  try {
    setStatus('⏳ กำลังสร้างไฟล์...', 'info');
    var inputs = formSection.querySelectorAll('[data-key]');
    var values = {};
    inputs.forEach(function (el) { values[el.getAttribute('data-key')] = el.value; });

    var sheetXml = state.sheetXml;
    var applied = [], skipped = [], emptyInput = [];

    Object.keys(state.resolved).forEach(function (key) {
      var r = state.resolved[key];
      var val = (values[key] || '').trim();
      if (!val) { emptyInput.push(r); return; }
      if (!isFillableCell(sheetXml, r.cellRef, state.sst)) {
        skipped.push(r);
        return;
      }
      var updated = fillCellXml(sheetXml, r.cellRef, val);
      if (updated === null) { skipped.push(r); return; }
      sheetXml = updated;
      applied.push(r);
    });

    if (!applied.length) {
      setStatus('⚠️ ไม่มีช่องไหนถูกกรอกเลย — กรุณากรอกข้อมูลอย่างน้อย 1 ช่องก่อนกด "สร้างไฟล์"', 'err');
      return;
    }

    var outputBytes = await buildOutputBytes(state.wb, sheetXml);

    // Sanity check: dropdown count in the sheet must be byte-for-byte unaffected.
    var beforeCount = (state.sheetXml.match(/x14:dataValidation\b/g) || []).length;
    var afterCount = (sheetXml.match(/x14:dataValidation\b/g) || []).length;
    if (beforeCount !== afterCount) {
      throw new Error('ตรวจพบว่าจำนวน dropdown เปลี่ยนไป (' + beforeCount + ' -> ' + afterCount + ') — ยกเลิกการสร้างไฟล์เพื่อความปลอดภัย');
    }

    var sapVal = (values.sap || '').trim() || normText(getCellText(sheetXml, state.resolved.sap.cellRef, state.sst)) || 'ไม่ระบุ SAP Name';
    var fname = '★KICK OFF MEETING NOTE - ' + sapVal + '.xlsx';

    renderResult(outputBytes, fname, applied, skipped, beforeCount);
    setStatus('✅ สร้างไฟล์สำเร็จ', 'ok');
  } catch (err) {
    setStatus('❌ ' + err.message, 'err');
  }
}

async function buildOutputBytes(wb, newSheetXml) {
  var enc = new TextEncoder();
  var sheetPath = wb.sheets[SHEET_NAME];
  var names = Object.keys(wb.entries);
  var zipFiles = [];
  for (var i = 0; i < names.length; i++) {
    var nm = names[i];
    var data = (nm === sheetPath) ? enc.encode(newSheetXml) : await decompressEntryBytes(wb.entries[nm], wb.buf);
    zipFiles.push({ name: nm, data: data });
  }
  return buildZip(zipFiles);
}

function renderResult(bytes, fname, applied, skipped, dvCount) {
  var html = '<div class="result-title">📋 สรุปผล</div>';
  html += '<div class="result-sub">กรอกสำเร็จ (' + applied.length + ' ช่อง)</div>';
  html += '<div class="detail-table-wrap"><table class="detail-table"><thead><tr><th>Cell</th><th>ฟิลด์</th></tr></thead><tbody>' +
    applied.map(function (r) { return '<tr><td>' + esc_(r.cellRef) + '</td><td>' + esc_(r.def.overrideLabel || r.def.label) + '</td></tr>'; }).join('') +
    '</tbody></table></div>';
  if (skipped.length) {
    html += '<div class="result-sub">ข้ามไป เพราะเซลล์มีข้อมูลอยู่แล้ว (' + skipped.length + ' ช่อง)</div>';
    html += '<div class="detail-table-wrap"><table class="detail-table"><thead><tr><th>Cell</th><th>ฟิลด์</th></tr></thead><tbody>' +
      skipped.map(function (r) { return '<tr><td>' + esc_(r.cellRef) + '</td><td>' + esc_(r.def.overrideLabel || r.def.label) + '</td></tr>'; }).join('') +
      '</tbody></table></div>';
  }
  html += '<div class="result-footer">Dropdown (x14:dataValidation) ในไฟล์: ' + dvCount + ' รายการ — ตรวจสอบแล้วว่าไม่เปลี่ยนแปลง</div>';
  html += '<div class="result-actions"><button class="btn-main" id="btnDownload">⬇️ ดาวน์โหลด ' + esc_(fname) + '</button></div>';
  resultSection.innerHTML = html;
  resultSection.style.display = 'block';
  document.getElementById('btnDownload').addEventListener('click', function () {
    var blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = fname;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 3000);
  });
}
