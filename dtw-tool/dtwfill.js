// ══════════════════════════════════════════════════════
//  Head/Line write engine. Only Head/Line sheets are ever modified -- every other sheet's zip
//  entry passes through byte-identical (wired in app.js). Confirmed real-file facts this depends
//  on ("Fill in data for DTW.xlsx" + DTW_SIEMENS_June2026_GLmapping.xlsx ground truth):
//  - Head rows 1-3 / Line rows 1-3 are the fixed header block (system/SAP/Thai names + cell
//    comments anchored to row 1) -- NEVER touched.
//  - Head's original sample rows (4-5) and Line's (4-7) hold a handful of genuinely-constant
//    columns pre-filled (bank fields, DocCurrency, DocRate, ObjType, ETAX, SendToRD, Line's
//    Currency/CostingCode) -- these are read from the template's OWN first sample row and copied
//    verbatim into every generated row, never hand-typed, so a future template edit is honored.
//  - Dates are ALWAYS written as literal YYYYMMDD text (inlineStr) -- inline strings ignore any
//    numFmt on the cell's style, so this is correct regardless of the sample cell's own numFmt.
//  - A handful of fields have no fully-automatable source in the template (PayToCode/branch,
//    postal code, free-text remarks/contact) -- best-effort defaults are used but every such
//    field is also pushed onto `warnings` for the user to review before trusting the output.
// ══════════════════════════════════════════════════════

var HEAD_COL = {
  DocNum: 'A', DocType: 'B', DocDate: 'C', DocDueDate: 'D', CardCode: 'E', CardName: 'F',
  NumAtCard: 'G', DocTotal: 'H', DocCurrency: 'I', DocRate: 'J', Comments: 'K',
  PaymentGroupCode: 'L', Series: 'M', TaxDate: 'N', DocObjectCode: 'O', FederalTaxID: 'P',
  PayToCode: 'Q', Address: 'R', OpeningRemarks: 'S', ContactPerson: 'T', Telephone: 'U',
  Email1: 'V', ItemDescription: 'W', Address2: 'X', Bank: 'Y', BankName: 'Z', BankBranch: 'AA',
  AccountType: 'AB', SwiftCode: 'AC', TaxBranch: 'AD', BookMonth: 'AE', BookYear: 'AF',
  Branch: 'AG', ETAXExport: 'AH', SendToRD: 'AI', ZipCode: 'AJ', Country: 'AK'
};
// Columns copied verbatim from the template's own first sample row -- truly constant per skill's
// documented "ไม่เปลี่ยน" fields.
var HEAD_CONSTANT_COLS = ['DocType', 'DocCurrency', 'DocRate', 'DocObjectCode', 'Bank', 'BankName',
  'BankBranch', 'AccountType', 'SwiftCode', 'TaxBranch', 'Branch', 'ETAXExport', 'SendToRD'];

var LINE_COL = {
  ParentKey: 'A', LineNum: 'B', Description: 'C', Price: 'D', PriceAfterVAT: 'E', Currency: 'F',
  AccountCode: 'G', CostingCode: 'H', ProjectCode: 'I', VatGroup: 'J', TaxType: 'K',
  LineTotal: 'L', TaxPercentagePerRow: 'M'
};
var LINE_CONSTANT_COLS = ['Currency', 'CostingCode'];

function round2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }

// Parses a GL_Invoice "Calendar Group"-style period value ("6/2026", "06/2026", a Date object,
// etc.) into { month (1-12), year (4-digit) }. Throws if the shape isn't recognized -- this feeds
// DocDate/BookMonth/BookYear so it must never be guessed silently.
function parsePeriod(periodVal) {
  if (periodVal instanceof Date) return { month: periodVal.getMonth() + 1, year: periodVal.getFullYear() };
  var m = /^(\d{1,2})\s*\/\s*(\d{4})$/.exec(String(periodVal || '').trim());
  if (m) return { month: parseInt(m[1], 10), year: parseInt(m[2], 10) };
  throw new Error('ไม่สามารถอ่านงวด (period) จากค่า "' + periodVal + '" ได้ -- รูปแบบที่รองรับ: M/YYYY หรือวันที่');
}

function pad2(n) { return String(n).padStart(2, '0'); }
function ymd(year, month, day) { return String(year) + pad2(month) + pad2(day); }
function lastDayOfMonth(year, month) { return new Date(year, month, 0).getDate(); }

// Best-effort postal code extraction from a free-text address (last 5-digit run) -- always
// flagged for manual review by the caller, never trusted blindly (mirrors the template's own
// Customer-sheet note "ดูจากบรรทัดบน" for this exact field).
function guessZipFromAddress(address) {
  var m = /(\d{5})\s*$/.exec(String(address || '').trim());
  return m ? m[1] : '';
}

// Builds one Head-row + N Line-row set of values (not yet written to XML) for one invoice group
// (from glinvoice.js's buildInvoiceGroups output), resolving customer/payment-group/series via
// lookups.js. Returns { headValues: {col:value}, lineValues: [{col:value}], warnings: [string] }.
function buildInvoiceRecord(invoice, docNum, lookups, sample) {
  var warnings = [];
  var lookup = lookupCustomer(lookups.customerMap, invoice.invoiceSentTo);
  if (!lookup.found) {
    var suggestion = lookup.candidates.length
      ? ' -- ใกล้เคียง: ' + lookup.candidates.slice(0, 3).map(function (c) { return c.name; }).join(', ')
      : ' -- ไม่พบชื่อใกล้เคียงเลยใน OUTSOURCE';
    warnings.push('ไม่พบลูกค้า "' + invoice.invoiceSentTo + '" ใน OUTSOURCE แบบตรงเป๊ะ (invoice ' + invoice.groupKey + ')' + suggestion);
  }
  var customer = lookup.found ? lookup.customer : { name: invoice.invoiceSentTo, code: '', taxId: '', address: '', paymentDueName: '' };

  var period = parsePeriod(invoice.period);
  var docDate = ymd(period.year, period.month, lastDayOfMonth(period.year, period.month));
  var bookMonth = pad2(period.month) + String(period.year).slice(-2);
  var bookYear = String(period.year);

  var creditDays = resolveCreditTermDays(customer.paymentDueName);
  var docDueDate;
  if (creditDays == null) {
    warnings.push('เงื่อนไข credit term "' + customer.paymentDueName + '" ของลูกค้า "' + customer.name + '" ไม่ใช่รูปแบบ "N Days" -- กรุณาคำนวณ/กรอก Due Date เองสำหรับ invoice ' + invoice.groupKey);
    docDueDate = docDate; // placeholder, flagged above -- never silently computed as a real due date
  } else {
    var due = new Date(period.year, period.month - 1, lastDayOfMonth(period.year, period.month));
    due.setDate(due.getDate() + creditDays);
    docDueDate = ymd(due.getFullYear(), due.getMonth() + 1, due.getDate());
  }

  var paymentGroupCode = lookups.paymentGroupMap.byName.get(String(customer.paymentDueName).toUpperCase());
  if (paymentGroupCode == null) {
    warnings.push('ไม่พบรหัส Group credit term สำหรับ "' + customer.paymentDueName + '" ในชีตคู่มือ (invoice ' + invoice.groupKey + ')');
    paymentGroupCode = '';
  }

  var series = invoice.isNoVat ? SERIES_NO_VAT : SERIES_VAT;
  var vatGroup = invoice.isNoVat ? VATGROUP_DEFAULT_NO_VAT : VATGROUP_DEFAULT_VAT;
  var taxType = invoice.isNoVat ? 'N' : 'Y';
  var vatPercent = invoice.isNoVat ? '0.00' : '7.00';
  var docTotal = invoice.isNoVat
    ? round2(invoice.totalBeforeVat)
    : round2(invoice.lines.reduce(function (s, l) { return s + round2(l.amount * 1.07); }, 0));

  var zip = guessZipFromAddress(customer.address);
  if (!zip) warnings.push('ไม่พบรหัสไปรษณีย์จากที่อยู่ลูกค้า "' + customer.name + '" -- กรุณากรอกเอง (invoice ' + invoice.groupKey + ')');

  var itemDescription = "PRTR Group Public Company Limited's outsourcing service fee as follows:\nMonth: " + MONTH_NAMES_EN[period.month - 1] + ' ' + period.year + '.';

  var headValues = {};
  headValues[HEAD_COL.DocNum] = { type: 'num', value: docNum };
  headValues[HEAD_COL.DocDate] = { type: 'text', value: docDate };
  headValues[HEAD_COL.DocDueDate] = { type: 'text', value: docDueDate };
  headValues[HEAD_COL.CardCode] = { type: 'text', value: customer.code };
  headValues[HEAD_COL.CardName] = { type: 'text', value: customer.name };
  headValues[HEAD_COL.NumAtCard] = { type: 'text', value: invoice.refNo };
  headValues[HEAD_COL.DocTotal] = { type: 'num', value: docTotal };
  headValues[HEAD_COL.PaymentGroupCode] = { type: 'text', value: String(paymentGroupCode) };
  headValues[HEAD_COL.Series] = { type: 'text', value: series };
  headValues[HEAD_COL.TaxDate] = { type: 'text', value: docDate };
  headValues[HEAD_COL.FederalTaxID] = { type: 'text', value: customer.taxId };
  headValues[HEAD_COL.PayToCode] = { type: 'text', value: sample.payToCode || 'Head Office' };
  headValues[HEAD_COL.Address] = { type: 'text', value: customer.address };
  headValues[HEAD_COL.Address2] = { type: 'text', value: customer.address };
  headValues[HEAD_COL.ContactPerson] = { type: 'text', value: invoice.lineManager || '' };
  headValues[HEAD_COL.ItemDescription] = { type: 'text', value: itemDescription };
  headValues[HEAD_COL.BookMonth] = { type: 'text', value: bookMonth };
  headValues[HEAD_COL.BookYear] = { type: 'text', value: bookYear };
  headValues[HEAD_COL.ZipCode] = { type: 'text', value: zip };
  headValues[HEAD_COL.Country] = { type: 'text', value: sample.country || 'TH' };
  headValues[HEAD_COL.Comments] = { type: 'text', value: '' };
  headValues[HEAD_COL.OpeningRemarks] = { type: 'text', value: '' };

  var lineValues = invoice.lines.map(function (line, i) {
    var price = round2(line.amount);
    var priceAfterVat = invoice.isNoVat ? price : round2(price * 1.07);
    var v = {};
    v[LINE_COL.ParentKey] = { type: 'num', value: docNum };
    v[LINE_COL.LineNum] = { type: 'num', value: i };
    v[LINE_COL.Description] = { type: 'text', value: bookMonth };
    v[LINE_COL.Price] = { type: 'num', value: price };
    v[LINE_COL.PriceAfterVAT] = { type: 'num', value: priceAfterVat };
    v[LINE_COL.AccountCode] = { type: 'text', value: line.incomeAccount };
    v[LINE_COL.ProjectCode] = { type: 'text', value: invoice.projectCodeSap };
    v[LINE_COL.VatGroup] = { type: 'text', value: vatGroup };
    v[LINE_COL.TaxType] = { type: 'text', value: taxType };
    v[LINE_COL.LineTotal] = { type: 'num', value: price };
    v[LINE_COL.TaxPercentagePerRow] = { type: 'text', value: vatPercent };
    return v;
  });

  return { headValues: headValues, lineValues: lineValues, warnings: warnings };
}

var MONTH_NAMES_EN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

// Reads the template's own first Head/Line sample row -- used both as the STYLE source (every
// generated cell in a column copies that column's s="N" index from here) and as the VALUE source
// for the genuinely-constant columns.
function readSampleRow(rows, constantCols, colMap, sst) {
  var firstDataRow = rows.find(function (r) { return r.rowNum > 3; });
  if (!firstDataRow) throw new Error('ไม่พบแถวตัวอย่างในชีต (แถวที่ 4 เป็นต้นไป) -- ไฟล์ template อาจผิดรูปแบบ');
  var styleByCol = {}, constantByCol = {};
  Object.keys(colMap).forEach(function (field) {
    var col = colMap[field];
    var cellXml = firstDataRow.cellsByCol[col];
    styleByCol[col] = getCellStyleIndex(cellXml);
    if (constantCols.indexOf(field) >= 0) constantByCol[field] = getCellValue(cellXml, sst);
  });
  return { rowNum: firstDataRow.rowNum, styleByCol: styleByCol, constants: constantByCol };
}

function buildCellsForRow(values, styleByCol, rowNum) {
  var cellsByCol = {};
  Object.keys(values).forEach(function (col) {
    var spec = values[col];
    var addr = col + rowNum;
    var style = styleByCol[col] || 0;
    if (spec.value === '' || spec.value == null) {
      cellsByCol[col] = buildBlankCell(addr, style);
    } else if (spec.type === 'num') {
      cellsByCol[col] = buildLiteralNumberCell(addr, style, spec.value);
    } else {
      cellsByCol[col] = buildInlineStringCell(addr, style, String(spec.value));
    }
  });
  return cellsByCol;
}

// Fills Head/Line sheet XML for the given invoice list. `templateHeadXml`/`templateLineXml` are
// the ORIGINAL (unmodified) sheet XML strings from the uploaded DTW template. Returns
// { headXml, lineXml, warnings, headWrites, lineWrites } where `*Writes` are the {addr, expected}
// lists verify.js needs (one per sheet), and `warnings` is the combined per-invoice manual-review
// list.
function fillHeadLine(templateHeadXml, templateLineXml, invoices, lookups, sst) {
  var headModel = parseSheetRows(templateHeadXml);
  var lineModel = parseSheetRows(templateLineXml);

  var headSample = readSampleRow(headModel.rows, HEAD_CONSTANT_COLS, HEAD_COL, sst);
  var lineSample = readSampleRow(lineModel.rows, LINE_CONSTANT_COLS, LINE_COL, sst);
  var sampleExtra = { payToCode: headSample.constants.PayToCode, country: headSample.constants.Country };

  var allWarnings = [];
  var headWrites = [], lineWrites = [];
  var newHeadRows = [], newLineRows = [];
  var lineRowNum = lineSample.rowNum;

  invoices.forEach(function (invoice, i) {
    var docNum = i + 1;
    var headRowNum = headSample.rowNum + i;
    var rec = buildInvoiceRecord(invoice, docNum, lookups, sampleExtra);
    allWarnings = allWarnings.concat(rec.warnings);

    // Head constant columns come from the sample row's own resolved value, not re-typed.
    HEAD_CONSTANT_COLS.forEach(function (field) {
      rec.headValues[HEAD_COL[field]] = { type: typeof headSample.constants[field] === 'number' ? 'num' : 'text', value: headSample.constants[field] };
    });
    var headCells = buildCellsForRow(rec.headValues, headSample.styleByCol, headRowNum);
    newHeadRows.push({ attrsRest: '', cellsByCol: headCells });
    Object.keys(rec.headValues).forEach(function (col) {
      var v = rec.headValues[col].value;
      if (v !== '' && v != null) headWrites.push({ addr: col + headRowNum, expected: v });
    });

    rec.lineValues.forEach(function (lv) {
      LINE_CONSTANT_COLS.forEach(function (field) {
        lv[LINE_COL[field]] = { type: 'text', value: lineSample.constants[field] };
      });
      var thisLineRowNum = lineRowNum++;
      var lineCells = buildCellsForRow(lv, lineSample.styleByCol, thisLineRowNum);
      newLineRows.push({ attrsRest: '', cellsByCol: lineCells });
      Object.keys(lv).forEach(function (col) {
        var v = lv[col].value;
        if (v !== '' && v != null) lineWrites.push({ addr: col + thisLineRowNum, expected: v });
      });
    });
  });

  // Overwrite the template's existing sample rows in place (same rowNum, so this is a pure
  // content swap), then insertRowsAfter for anything beyond the original sample count.
  applyRowsToModel(headModel, newHeadRows, 4);
  applyRowsToModel(lineModel, newLineRows, 4);

  function applyRowsToModel(model, newRows, startRowNum) {
    var existingIdx = [];
    model.rows.forEach(function (r, idx) { if (r.rowNum >= startRowNum) existingIdx.push(idx); });
    var existingCount = existingIdx.length;
    for (var i = 0; i < newRows.length && i < existingCount; i++) {
      model.rows[existingIdx[i]].cellsByCol = newRows[i].cellsByCol;
    }
    if (newRows.length > existingCount) {
      var extra = newRows.slice(existingCount);
      var lastRowNum = model.rows[model.rows.length - 1].rowNum;
      model.rows = insertRowsAfter(model.rows, lastRowNum, extra);
    } else if (newRows.length < existingCount) {
      // Fewer invoices than sample rows -- delete the unused trailing sample rows so no
      // constant-only placeholder rows leak into the output.
      var toRemove = existingIdx.slice(newRows.length);
      var removeSet = new Set(toRemove.map(function (idx) { return model.rows[idx].rowNum; }));
      model.rows = model.rows.filter(function (r) { return !removeSet.has(r.rowNum); });
    }
  }

  var headExtent = columnExtent(headModel.rows, 1);
  var lineExtent = columnExtent(lineModel.rows, 1);
  var headLastRow = headModel.rows[headModel.rows.length - 1].rowNum;
  var lineLastRow = lineModel.rows[lineModel.rows.length - 1].rowNum;

  var headXml = serializeSheetRows(headModel);
  var lineXml = serializeSheetRows(lineModel);
  headXml = updateDimension(headXml, headExtent.firstCol, headExtent.lastCol, headLastRow);
  lineXml = updateDimension(lineXml, lineExtent.firstCol, lineExtent.lastCol, lineLastRow);

  return { headXml: headXml, lineXml: lineXml, warnings: allWarnings, headWrites: headWrites, lineWrites: lineWrites };
}
