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
  var mainHintReset = document.getElementById('mainEmptyHint');
  if (mainHintReset) mainHintReset.style.display = 'block';
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
  var mainHint = document.getElementById('mainEmptyHint');
  if (mainHint) mainHint.style.display = 'none';

  var b2bStartEl = document.getElementById('fld_b2bStart');
  if (b2bStartEl) b2bStartEl.addEventListener('input', updateDerivedDates);
  document.getElementById('btnGenerate').addEventListener('click', generateFile);

  // If a CRM PDF was already parsed (uploaded before or after the xlsx, order doesn't matter),
  // apply it now that the form fields actually exist in the DOM.
  if (crmParsedData) applyCrmDataToForm(crmParsedData);
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

// XML 1.0 forbids most control characters outright (anything outside #x9 | #xA | #xD |
// [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]) -- not just the &/</> that need escaping.
// A stray NUL or other control byte written into a cell produces a .xlsx whose sheet XML doesn't
// even parse, which Excel then either refuses to open or silently "repairs" by dropping content
// (confirmed real-world case: pdf.js's text extraction occasionally emits a literal U+0000 inside
// an otherwise normal Thai word, e.g. a CRM PDF's "...เลื่อนขึ้น..." -- that one invisible byte was
// enough to corrupt the whole output file). Strip these before they ever reach the XML, regardless
// of whether the value came from the CRM PDF auto-fill or was typed/pasted by hand.
function stripInvalidXmlChars(s) {
  return String(s == null ? '' : s).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/g, '');
}

function escXmlText(s) {
  return stripInvalidXmlChars(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
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

// Other sheets in this workbook (e.g. "SAP & BP") hold formulas that reference cells on "Master
// for inovice" -- writing new values into that sheet's raw XML doesn't retouch those formula
// cells' own cached <v>, so without this fix Excel keeps showing their PRE-edit cached result
// (0 / "- Select -") until the user manually forces a recalc on every single one (confirmed
// real-world case: a filled-in Kick-off file's "SAP & BP" sheet still showed stale values).
// Dropping calcChain.xml (a now-stale calculation-order cache) and forcing fullCalcOnLoad makes
// Excel recompute every formula in the workbook the moment it opens -- the same fix already used
// in accrued-income/app.js and doi-tools/index.html for the identical stale-formula symptom.
async function buildOutputBytes(wb, newSheetXml) {
  var enc = new TextEncoder();
  var sheetPath = wb.sheets[SHEET_NAME];
  var names = Object.keys(wb.entries);
  var zipFiles = [];
  for (var i = 0; i < names.length; i++) {
    var nm = names[i];
    if (nm === 'xl/calcChain.xml') continue;
    var data;
    if (nm === sheetPath) {
      data = enc.encode(newSheetXml);
    } else if (nm === '[Content_Types].xml') {
      var ct = await decompressEntry(wb.entries[nm], wb.buf);
      ct = ct.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, '');
      data = enc.encode(ct);
    } else if (nm === 'xl/_rels/workbook.xml.rels') {
      var rels = await decompressEntry(wb.entries[nm], wb.buf);
      rels = rels.replace(/<Relationship[^>]*Target="calcChain\.xml"[^>]*\/>/, '');
      data = enc.encode(rels);
    } else if (nm === 'xl/workbook.xml') {
      var wbXml = await decompressEntry(wb.entries[nm], wb.buf);
      if (/<calcPr\b[^>]*\/>/.test(wbXml)) {
        wbXml = wbXml.replace(/<calcPr\b([^>]*)\/>/, function (_, attrs) {
          return /fullCalcOnLoad=/.test(attrs)
            ? '<calcPr' + attrs.replace(/fullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"') + '/>'
            : '<calcPr' + attrs + ' fullCalcOnLoad="1"/>';
        });
      } else {
        wbXml = wbXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
      }
      data = enc.encode(wbXml);
    } else {
      data = await decompressEntryBytes(wb.entries[nm], wb.buf);
    }
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

// ══════════════════════════════════════════════════════
//  CRM PDF auto-fill — reads a printed/exported "Project Intro Details" PDF from PRTR's
//  Backoffice CRM and pre-fills this form's inputs the same way a person copy-pasting from the
//  live CRM page would, following the same label -> Excel-cell mapping as the
//  prtr-crm-kickoff-invoice Claude Skill (references/field_mapping.md). This never touches the
//  xlsx directly -- it only sets .value on the already-rendered form fields, so the normal
//  "only overwrite genuinely blank cells" review/generate flow above is completely unchanged.
//
//  pdf.js's getTextContent() returns each page's text with no line breaks and (per real-world
//  testing against an actual exported CRM PDF) no space between one field's value and the next
//  field's label either -- e.g. "...Industry :ผลิตภัณฑ์ก่อสร้างContract Type :Outsource...". So
//  every field's value has to be bounded by searching for the START of the NEXT known label in
//  the whole document, not a line break or fixed delimiter. CRM_LABELS is deliberately built from
//  every label observed across a full real 7-page export (not guessed), including labels this
//  tool never fills, purely so they still work as boundary markers for the ones it does fill.
// ══════════════════════════════════════════════════════

var crmParsedData = null; // set once a PDF has been parsed; survives xlsx (re)uploads

// Every label text this CRM template prints, in no particular order -- order does not need to
// match the PDF's own layout, only each label's OWN occurrence position (found independently via
// indexOf) matters for boundary detection. Longer/more specific labels that share a prefix with a
// shorter one (e.g. "1st Invoice Month (Please Specify DueDate)" vs "1st Invoice Month") are
// listed as their own distinct literal strings -- indexOf on the full literal string finds each
// one's own real occurrence, so there's no ambiguity between the two.
var CRM_LABELS = [
  'Date of Issue', 'Sales Manager', 'Recruit Manager', 'HR Manager',
  'Sap Code', 'SAP Code - Sub Group', 'Business Unit', 'Industry', 'Contract Type', 'Staff Type',
  'Company Name', 'Tax ID', 'Head Office/Branch No', 'Branch No',
  'Company Address', 'Company Billing Delivery Address',
  'Detail of contact person for invoice', 'Detail of contact person for this project',
  'Full Name', 'Position', 'Email', 'Mobile Phone', 'Company Phone',
  'Salary Details', 'Payrise After Probation Remark', 'Payrise After Probation', 'Pay Group',
  'Reimbursement', 'Contact Period', 'Started Date', 'Finished Date', 'Day of Work',
  'Hours of Work-In', 'Time-In Note', 'Hours of Work-Out', 'Time-Out Note', 'Break Hour(minute)',
  'Shift Pattern', 'Over Time', 'Continue Service Year (Only TransferCase)', 'Probation Period On Day',
  'Provident Fund Employer Rate %', 'Provident Fund Employee Rate %', 'Provident Fund Type', 'Provident Fund',
  'Health Insurance Plan', 'Health Insurance Cost By', 'Accident Insurance Plan', 'Accident Insurance Cost By',
  'Employee Bond Insurance Plan', 'Employee Bond Insurance Cost By',
  'Annual Leave Effective Date', 'Annual Leave', 'Business Leave', 'Any Deduction (Absence/Late)',
  'Other HR Requirement',
  'Variable Cut-Off Date (OT, Incentive, Commission)', 'Salary Cut-Off Date',
  'Salary Pay Date', 'Variable Pay Date',
  'B2B Contract No.', 'B2B Start Date', 'B2B End Date', 'Type of Invoice',
  'Deposit : Staff Income (No. of Deposit Month)', 'Deposit : Staff Income',
  '1st Invoice Month (Please Specify DueDate)', '1st Invoice Month',
  'Date of Client Received Invoice', 'Credit Terms (Please copy from B2B)', 'Client Payment Cycle',
  'Penalty Accrued', 'Media Pay Via', 'Required P/O', 'PO Type', 'PO Issue By',
  'Special Instructions', 'Deduction',
  'Interest Budget(In Days)', 'Estimated Payroll Size (No.Staff)', 'Estimated Payroll Size(Amount)',
  'Late Payment InterestCharge Amount(%)', 'Charge Type', 'Create By'
];

// Pure section-header text this template prints between fields -- never followed by a colon (so
// they can't be found the same way as a real "Label :value" field), but still needed as boundary
// markers so a real field's value doesn't bleed into this filler text before the next real label
// (confirmed real-world case: "Staff Type" bled into "Office & OthersClient Company1." because
// "Client Company" wasn't tracked as a boundary at all).
var CRM_SECTION_HEADERS = [
  'Details of Client', 'Client Company', 'Details of Contact Person',
  'Details of Hiring & Payroll', 'Pre Employee Process',
  'Details for Invoices Part A', 'Details for Invoices Part B', 'Attach additional files'
];

async function extractPdfAllText(file) {
  if (typeof pdfjsLib === 'undefined') throw new Error('โหลด pdf.js ไม่สำเร็จ (ตรวจสอบการเชื่อมต่ออินเทอร์เน็ต)');
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  var buf = await file.arrayBuffer();
  var pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  var text = '';
  for (var p = 1; p <= pdf.numPages; p++) {
    var page = await pdf.getPage(p);
    var content = await page.getTextContent();
    text += content.items.map(function (i) { return i.str; }).join('');
  }
  // pdf.js occasionally emits a literal control character (e.g. U+0000) inside an otherwise
  // normal word when extracting certain Thai glyphs/ligatures -- strip these at the source so
  // every downstream value (displayed in the form, or later written into the xlsx) is clean.
  return stripInvalidXmlChars(text);
}

// Finds each CRM_LABELS entry's own occurrence within `text`. A label is only accepted where it's
// immediately followed by (optional whitespace then) a colon -- some labels are literal substrings
// of a longer label (e.g. "Branch No" inside "Head Office/Branch No", "Deduction" inside "Any
// Deduction (Absence/Late)"), and a plain indexOf() would lock onto that embedded false match
// instead of the real standalone "Branch No :" occurrence further down, confirmed against a real
// exported CRM PDF. Keeps searching past a false match (no colon right after it) for the next
// occurrence of the same literal string.
// "Branch No" is also the literal tail of the longer label "Head Office/Branch No" -- and that
// longer label's OWN colon comes right after "...Branch No" too, so the colon-lookahead check alone
// can't tell the two apart (confirmed real-world case). The one distinguishing signal available is
// what comes immediately before the match: the false (embedded) occurrence is always preceded by
// "/", which a genuine standalone "Branch No" label never is.
function isEmbeddedBranchNo(text, idx) {
  return idx > 0 && text[idx - 1] === '/';
}

function findLabelPositions(text, labels) {
  var out = [];
  labels.forEach(function (label) {
    var searchFrom = 0;
    while (true) {
      var idx = text.indexOf(label, searchFrom);
      if (idx === -1) break;
      var afterStart = idx + label.length;
      var colonOk = /^\s*:/.test(text.slice(afterStart, afterStart + 6));
      var isFalseMatch = label === 'Branch No' && isEmbeddedBranchNo(text, idx);
      if (colonOk && !isFalseMatch) {
        out.push({ label: label, start: idx, end: afterStart });
        break;
      }
      searchFrom = idx + 1;
    }
  });
  return out;
}

// Section headers are boundary-only -- no colon follows them, so a plain first-match indexOf is
// enough (none of CRM_SECTION_HEADERS is a literal substring of another header or of a CRM_LABELS
// entry, unlike the value-bearing labels above).
function findHeaderPositions(text, headers) {
  var out = [];
  headers.forEach(function (h) {
    var idx = text.indexOf(h);
    if (idx !== -1) out.push({ label: h, start: idx, end: idx + h.length });
  });
  return out;
}

// Full boundary set for one text blob: every value-bearing label this tool cares about, PLUS every
// pure section-header marker, merged and sorted by position -- extractLabelValue() doesn't care
// which list a given boundary came from, only where it starts.
function buildCrmPositions(text) {
  var positions = findLabelPositions(text, CRM_LABELS).concat(findHeaderPositions(text, CRM_SECTION_HEADERS));
  positions.sort(function (a, b) { return a.start - b.start; });
  return positions;
}

// Value = whatever sits right after "<label>:" (space before/after the colon both optional -- real
// examples show both "Label :Value" and "Label(...):Value" with no space) up to the start of
// whichever known label occurs soonest afterward, or end-of-text if none does.
function extractLabelValue(text, positions, label) {
  var m = null;
  for (var i = 0; i < positions.length; i++) { if (positions[i].label === label) { m = positions[i]; break; } }
  if (!m) return null;
  var after = text.slice(m.end);
  var colonM = /^\s*:\s*/.exec(after);
  if (!colonM) return null;
  var valueStart = m.end + colonM[0].length;
  var nextStart = text.length;
  positions.forEach(function (p) { if (p.start >= valueStart && p.start < nextStart) nextStart = p.start; });
  return text.slice(valueStart, nextStart).trim();
}

function crmVal(text, positions, label) {
  var v = extractLabelValue(text, positions, label);
  return v == null ? '' : v;
}

// "Detail of contact person for invoice (...)" and "Detail of contact person for this project"
// both have their own Full Name/Position/Email/Mobile Phone/Company Phone underneath -- this tool
// only fills the INVOICE contact's fields, so extraction is scoped to just that section's text
// rather than matching whichever "Full Name" occurs first globally.
function sliceBetween(text, startLabel, endLabel) {
  var s = text.indexOf(startLabel);
  if (s === -1) return '';
  var e = endLabel ? text.indexOf(endLabel, s + startLabel.length) : -1;
  if (e === -1) e = text.length;
  return text.slice(s, e);
}

// Values this CRM template prints when a field genuinely has no real answer -- must never be
// treated as real contact data (confirmed real-world case: the invoice contact section showed
// "Full Name: As line manager", "Email: XX@bluescope.com" (masked), "Mobile Phone: 000-000-0000").
function isCrmJunkValue(v) {
  var t = normText(v).toLowerCase();
  if (t === '' || t === '-') return true;
  if (t === 'as line manager') return true;
  if (/^000-000-0000$/.test(t)) return true;
  if (/^xx@/i.test(t)) return true;
  return false;
}

function pad2(n) { return String(n).length < 2 ? '0' + n : String(n); }

// CRM shows dates as DD/MM/YYYY; the Excel sheet wants DD-MM-YY (see field_mapping.md).
function reformatCrmDate(v) {
  var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(normText(v));
  if (!m) return null;
  return pad2(m[1]) + '-' + pad2(m[2]) + '-' + m[3].slice(-2);
}

// "Normal (Actual)" -> "Normal" -- the dropdown only has the bare word (see field_mapping.md).
function stripParenthetical(v) {
  return normText(v).replace(/\s*\([^)]*\)\s*$/, '').trim();
}

// Splits "89/12 Moo 3, ... District, Province 12150, Thailand" into
// { building: '...up to district', city: 'Province', zip: '12150' }. Returns null if the address
// doesn't end in "<words> <5-digit zip>[, Thailand]" the way every real example so far has.
function splitCrmAddress(v) {
  var s = normText(v).replace(/,\s*Thailand\s*$/i, '');
  var m = /^(.*),\s*([A-Za-zก-๙.\s]+?)\s+(\d{5})$/.exec(s);
  if (!m) return null;
  return { building: m[1].trim(), city: m[2].trim(), zip: m[3] };
}

// Best-effort Thai -> English match against the Industry dropdown's fixed 25-item list (see
// field_mapping.md). Most-specific phrase checked first so e.g. "ผลิตภัณฑ์ก่อสร้าง" (construction
// PRODUCTS) doesn't fall through to the broader "ก่อสร้าง" (construction, i.e. CONSTRUCTION & REAL
// ESTATE) match. Returns { value, confident } -- confident:false means "OTHERS" was used as a
// fallback and the user should double check what CRM actually said.
var INDUSTRY_TH_MAP = [
  { pats: ['ผลิตภัณฑ์ก่อสร้าง'], value: 'CONSTRUCTION PRODUCTS' },
  { pats: ['เกษตร'], value: 'AGRICULTURE PRODUCTS' },
  { pats: ['ยานยนต์', 'การบิน'], value: 'AUTOMOTIVE & AVIATION' },
  { pats: ['ธนาคาร', 'การเงิน', 'สถาบันการเงิน'], value: 'BANKING & FINANCIAL SERVICES' },
  { pats: ['เคมีภัณฑ์', 'พลาสติก', 'กระดาษ'], value: 'CHEMICALS/PLASTICS/PAPER' },
  { pats: ['ก่อสร้าง', 'อสังหาริมทรัพย์'], value: 'CONSTRUCTION & REAL ESTATE' },
  { pats: ['อิเล็กทรอนิกส์', 'มือถือ'], value: 'CONSUMER ELECTRONICS/ MOBILE' },
  { pats: ['สินค้าอุปโภคบริโภค'], value: 'CONSUMER GOODS' },
  { pats: ['แฟชั่น', 'สิ่งทอ'], value: 'FASHION & TEXTILE' },
  { pats: ['อาหารและเครื่องดื่ม', 'อาหาร'], value: 'FOOD & BEVERAGE' },
  { pats: ['สุขภาพ', 'ความงาม', 'เครื่องสำอาง'], value: 'HEALTH/BEAUTY/COSMETIC' },
  { pats: ['เทคโนโลยีสารสนเทศ', 'ไอที'], value: 'INFORMATION TECHNOLOGY (IT)' },
  { pats: ['ประกันภัย'], value: 'INSURANCE' },
  { pats: ['โลจิสติกส์', 'ขนส่ง'], value: 'LOGISTICS/FREIGHT FORWARDING' },
  { pats: ['เครื่องจักร', 'ชิ้นส่วนอุตสาหกรรม'], value: 'MACHINERY/INDUSTRIAL PARTS & COMPONENTS' },
  { pats: ['องค์กรไม่แสวงหาผลกำไร'], value: 'NON-PROFIT ORGANIZATION' },
  { pats: ['น้ำมัน', 'พลังงาน', 'ก๊าซ'], value: 'OIL & GAS/ ENERGY' },
  { pats: ['อีคอมเมิร์ซ', 'ออนไลน์'], value: 'ONLINE & E-COMMERCE' },
  { pats: ['เภสัชกรรม', 'ยา'], value: 'PHARMACEUTICAL' },
  { pats: ['โทรคมนาคม'], value: 'TELECOMMUNICATION' },
  { pats: ['ท่องเที่ยว', 'โรงแรม'], value: 'TOURIST & HOSPITALITY' },
  { pats: ['สาธารณูปโภค', 'รีไซเคิล'], value: 'UTILITIES & RECYCLING' },
  { pats: ['ค้าส่ง', 'ค้าปลีก'], value: 'WHOLESALE/RETAIL' },
  { pats: ['บริการ'], value: 'SERVICES' }
];
function translateIndustry(thaiText) {
  var t = normText(thaiText);
  if (!t) return { value: '', confident: false };
  for (var i = 0; i < INDUSTRY_TH_MAP.length; i++) {
    var entry = INDUSTRY_TH_MAP[i];
    for (var j = 0; j < entry.pats.length; j++) {
      if (t.indexOf(entry.pats[j]) !== -1) return { value: entry.value, confident: true };
    }
  }
  return { value: 'OTHERS', confident: false };
}

// ── Main parse: PDF full text -> { values: {key: value}, flags: [message, ...] } ──────────────
function parseCrmPdfText(text) {
  var positions = buildCrmPositions(text);
  var values = {};
  var flags = [];

  var sap = crmVal(text, positions, 'Sap Code');
  if (sap) values.sap = sap;

  var salesEmail = crmVal(text, positions, 'Sales Manager');
  if (salesEmail) { values.salesEmail = salesEmail; flags.push('Sales Manager: CRM มีแค่อีเมล (' + salesEmail + ') ใส่ไว้ในช่อง Email แล้ว — ชื่อจริง/ชื่อเล่นกรุณากรอกเอง'); }
  var hrEmail = crmVal(text, positions, 'HR Manager');
  if (hrEmail) { values.hrEmail = hrEmail; flags.push('HR Project Owner: CRM มีแค่อีเมล (' + hrEmail + ') ใส่ไว้ในช่อง Email แล้ว — ชื่อจริง/ชื่อเล่นกรุณากรอกเอง'); }
  var recruitEmail = crmVal(text, positions, 'Recruit Manager');
  if (recruitEmail) { values.recruitEmail = recruitEmail; flags.push('Recruitment Project Owner: CRM มีแค่อีเมล (' + recruitEmail + ') ใส่ไว้ในช่อง Email แล้ว — ชื่อจริง/ชื่อเล่นกรุณากรอกเอง'); }

  var companyName = crmVal(text, positions, 'Company Name');
  if (companyName) values.companyName = companyName;
  var taxId = crmVal(text, positions, 'Tax ID');
  if (taxId) values.taxId = taxId;
  var branch = crmVal(text, positions, 'Head Office/Branch No');
  if (branch) values.branch = branch;

  var industryRaw = crmVal(text, positions, 'Industry');
  if (industryRaw) {
    var ind = translateIndustry(industryRaw);
    values.industry = ind.value;
    if (!ind.confident) flags.push('Industry: CRM ระบุ "' + industryRaw + '" ไม่พบตัวเลือกที่ตรงชัดเจน — ใส่ "OTHERS" ไว้ก่อน กรุณาตรวจสอบ/เลือกใหม่');
  }

  var staffType = crmVal(text, positions, 'Staff Type');
  if (staffType) values.staffType = staffType;

  var addrRaw = crmVal(text, positions, 'Company Address');
  if (addrRaw) {
    var addr = splitCrmAddress(addrRaw);
    if (addr) { values.addr1 = addr.building; values.addr2 = addr.city; values.addr3 = addr.zip; }
    else flags.push('Company Address: รูปแบบที่อยู่จาก CRM ไม่ตรงกับที่คาด ("' + addrRaw + '") — กรุณาแยก Building/City/Zip เอง');
  }
  flags.push('Billing Delivery Address: ต้องเช็ค "Original Send by" ในหน้า CRM ก่อนว่าต้องส่งเอกสารจริงหรือไม่ — ระบบไม่ auto-fill ช่องนี้ให้');

  // Invoice contact -- scoped to avoid the second "Full Name" (project contact) block below it.
  var invBlock = sliceBetween(text, 'Detail of contact person for invoice', 'Detail of contact person for this project');
  if (invBlock) {
    var invPositions = findLabelPositions(invBlock, ['Full Name', 'Position', 'Email', 'Mobile Phone', 'Company Phone']);
    var fullName = crmVal(invBlock, invPositions, 'Full Name');
    var position = crmVal(invBlock, invPositions, 'Position');
    var email = crmVal(invBlock, invPositions, 'Email');
    var mobile = crmVal(invBlock, invPositions, 'Mobile Phone');
    var phone = crmVal(invBlock, invPositions, 'Company Phone');

    if (!isCrmJunkValue(fullName)) {
      var nm = /^Khun\s+(\S+)\s+(.+)$/i.exec(fullName.trim());
      if (nm) { values.contactFirst = 'Khun ' + nm[1]; values.contactLast = nm[2]; }
      else { values.contactFirst = fullName.trim(); values.contactLast = '-'; }
    } else if (fullName) {
      flags.push('ผู้ติดต่อใบแจ้งหนี้ (First/Last Name): CRM ระบุว่า "' + fullName + '" (ใช้ข้อมูลเดียวกับ Line Manager) — ไม่มีชื่อจริงให้ดึง กรุณากรอกเอง');
    }
    if (!isCrmJunkValue(position)) values.contactPosition = position;
    if (!isCrmJunkValue(email)) values.contactEmail = email;
    else if (email) flags.push('ผู้ติดต่อใบแจ้งหนี้ (Email): CRM แสดงเป็นอีเมลปกปิด ("' + email + '") — กรุณากรอกอีเมลจริงเอง');
    if (!isCrmJunkValue(mobile)) values.contactMobile = mobile;
    if (phone) values.contactPhone = normText(phone); // "-" is itself a valid value for this field
  }

  // Salary/Variable Cut-Off Date: real CRM export renders these two adjacent date-range cells as
  // one squished string (e.g. "1-30 1-30") when both fields sit right next to Salary Pay Date with
  // no other label between them -- split by whitespace into the two positions in that case.
  var salaryCutoff = crmVal(text, positions, 'Salary Cut-Off Date');
  var variableCutoff = crmVal(text, positions, 'Variable Cut-Off Date (OT, Incentive, Commission)');
  if (!salaryCutoff && variableCutoff) {
    var pair = variableCutoff.split(/\s+/);
    if (pair.length === 2) { salaryCutoff = pair[0]; variableCutoff = pair[1]; }
  }
  if (salaryCutoff) values.salaryCutoff = salaryCutoff;
  if (variableCutoff) values.variableCutoff = variableCutoff;

  var salaryPayDate = crmVal(text, positions, 'Salary Pay Date');
  if (salaryPayDate) values.salaryPayDate = salaryPayDate;
  var variablePayDate = crmVal(text, positions, 'Variable Pay Date');
  if (variablePayDate) values.variablePayDate = variablePayDate;

  var b2bNo = crmVal(text, positions, 'B2B Contract No.');
  if (b2bNo) values.b2bNo = b2bNo;
  var b2bStartRaw = crmVal(text, positions, 'B2B Start Date');
  var b2bStart = reformatCrmDate(b2bStartRaw);
  if (b2bStart) values.b2bStart = b2bStart;
  else if (b2bStartRaw) flags.push('B2B Start Date: รูปแบบวันที่จาก CRM ไม่ตรงที่คาด ("' + b2bStartRaw + '") — กรุณากรอกเอง (DD-MM-YY)');
  var b2bEndRaw = crmVal(text, positions, 'B2B End Date');
  var b2bEnd = reformatCrmDate(b2bEndRaw);
  if (b2bEnd) values.b2bEnd = b2bEnd;
  else if (b2bEndRaw) flags.push('B2B End Date: รูปแบบวันที่จาก CRM ไม่ตรงที่คาด ("' + b2bEndRaw + '") — กรุณากรอกเอง (DD-MM-YY)');

  var invoiceTypeRaw = crmVal(text, positions, 'Type of Invoice');
  if (invoiceTypeRaw) values.invoiceType = stripParenthetical(invoiceTypeRaw);

  var firstInvRaw = crmVal(text, positions, '1st Invoice Month');
  var firstInv = reformatCrmDate(firstInvRaw);
  if (firstInv) values.firstInvoiceDate = firstInv;
  else if (firstInvRaw) flags.push('1st Invoice Date: รูปแบบวันที่จาก CRM ไม่ตรงที่คาด ("' + firstInvRaw + '") — กรุณากรอกเอง (DD-MM-YY)');
  var firstInvDueRaw = crmVal(text, positions, '1st Invoice Month (Please Specify DueDate)');
  var firstInvDue = reformatCrmDate(firstInvDueRaw);
  if (firstInvDue) values.firstInvoiceDue = firstInvDue;
  else if (firstInvDueRaw) flags.push('1st Invoice Due Date: รูปแบบวันที่จาก CRM ไม่ตรงที่คาด ("' + firstInvDueRaw + '") — กรุณากรอกเอง (DD-MM-YY)');

  var paymentTerms = crmVal(text, positions, 'Credit Terms (Please copy from B2B)');
  if (paymentTerms) values.paymentTerms = paymentTerms;
  var clientCycle = crmVal(text, positions, 'Client Payment Cycle');
  if (clientCycle) values.clientCycle = clientCycle;

  var mediaPay = crmVal(text, positions, 'Media Pay Via');
  if (mediaPay && !isCrmJunkValue(mediaPay)) { values.bank = mediaPay; flags.push('Payment Run - Bank: เดา mapping จาก CRM "Media Pay Via" (' + mediaPay + ') กรุณาตรวจสอบว่าตรงกับตัวเลือกจริง'); }

  var requiredPo = crmVal(text, positions, 'Required P/O');
  if (requiredPo) values.requiredPo = requiredPo;

  var specialInstructions = crmVal(text, positions, 'Special Instructions');
  if (specialInstructions) values.specialInstructions = specialInstructions;

  flags.push('Payment Terms (Days): ไม่ auto-fill ตามกฎเสมอ — กรุณาเลือกเองจาก dropdown');

  return { values: values, flags: flags };
}

// ── Apply parsed CRM data onto the already-rendered form ───────────────────────────────────────
// Only ever sets a field that's currently blank, so re-uploading a PDF (or uploading it after the
// user already typed something) never clobbers a manual edit.
function applyCrmDataToForm(data) {
  if (!formSection || formSection.style.display === 'none') return; // form not rendered yet
  var filled = [];
  Object.keys(data.values).forEach(function (key) {
    var val = data.values[key];
    if (val == null || val === '') return;
    var el = formSection.querySelector('[data-key="' + key + '"]');
    if (!el) return;
    if (el.tagName === 'SELECT') {
      if (el.value) return; // already has a selection
      var opts = Array.from(el.options);
      var match = opts.find(function (o) { return o.value.toLowerCase() === String(val).toLowerCase(); }) ||
        opts.find(function (o) { return o.value && o.value.toLowerCase().indexOf(String(val).toLowerCase()) !== -1; });
      if (match) { el.value = match.value; filled.push(key); }
      else data.flags.push((el.previousElementSibling ? '' : '') + 'ไม่พบตัวเลือกที่ตรงกับ "' + val + '" ในช่อง ' + key + ' — กรุณาเลือกเอง');
    } else {
      if (el.value) return; // already has a value (user typed, or filled by an earlier PDF)
      el.value = val;
      filled.push(key);
    }
  });

  var b2bStartEl = document.getElementById('fld_b2bStart');
  if (b2bStartEl && b2bStartEl.value) updateDerivedDates();

  renderCrmStatus(filled.length, data.flags);
}

function renderCrmStatus(filledCount, flags) {
  var box = document.getElementById('crmPdfStatus');
  if (!box) return;
  var uniqueFlags = flags.filter(function (f, i) { return flags.indexOf(f) === i; });
  var html = '📎 <b>อ่าน PDF จาก CRM แล้ว</b> — กรอกอัตโนมัติ ' + filledCount + ' ช่อง';
  if (uniqueFlags.length) {
    html += '<ul>' + uniqueFlags.map(function (f) { return '<li>' + esc_(f) + '</li>'; }).join('') + '</ul>';
  }
  box.innerHTML = html;
  box.style.display = 'block';
}

// ── PDF dropzone wiring ─────────────────────────────────────────────────────────────────────────
var crmPdfDz = document.getElementById('crmPdfDropzone');
var crmPdfInput = document.getElementById('crmPdfInput');
if (crmPdfDz && crmPdfInput) {
  crmPdfDz.addEventListener('click', function (e) { if (e.target.tagName !== 'INPUT') crmPdfInput.click(); });
  crmPdfInput.addEventListener('change', function () { if (crmPdfInput.files.length) handleCrmPdf(crmPdfInput.files[0]); });
  ['dragover', 'dragenter'].forEach(function (ev) { crmPdfDz.addEventListener(ev, function (e) { e.preventDefault(); crmPdfDz.classList.add('drag'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { crmPdfDz.addEventListener(ev, function (e) { e.preventDefault(); crmPdfDz.classList.remove('drag'); }); });
  crmPdfDz.addEventListener('drop', function (e) { if (e.dataTransfer.files.length) handleCrmPdf(e.dataTransfer.files[0]); });
}

async function handleCrmPdf(file) {
  if (!/\.pdf$/i.test(file.name)) { setStatus('รองรับเฉพาะไฟล์ .pdf สำหรับข้อมูล CRM', 'err'); return; }
  var box = document.getElementById('crmPdfStatus');
  if (box) { box.style.display = 'block'; box.innerHTML = '⏳ กำลังอ่าน PDF...'; }
  try {
    var text = await extractPdfAllText(file);
    crmParsedData = parseCrmPdfText(text);
    applyCrmDataToForm(crmParsedData);
    if (formSection.style.display === 'none') {
      if (box) box.innerHTML = '📎 <b>อ่าน PDF สำเร็จ</b> — รออัปโหลดไฟล์ ★KICK OFF MEETING NOTE.xlsx ด้านบนก่อน แล้วข้อมูลจะถูกกรอกให้อัตโนมัติ';
    }
  } catch (err) {
    if (box) box.innerHTML = '❌ อ่าน PDF ไม่สำเร็จ: ' + esc_(err.message);
  }
}
