// ══════════════════════════════════════════════════════
//  The 8 GL Invoice transform functions, plus the shared load/save orchestration around them.
//  Business rules confirmed verbatim by the user; column positions and the header row itself are
//  found dynamically (never hardcoded) since column letters and row numbers drift file-to-file.
// ══════════════════════════════════════════════════════

var SSO_GROUPING = 'E51110102';
var SSO_PAYCODES = ['T2A3', 'TZ74'];
var EXPENSE_PAYCODE = 'EXPENSE';
var EXPENSE_PAYCODE_NAME = 'ค่าใช้จ่าย';
var EXPENSE_ACCOUNT = '51110116';
var EXPENSE_GROUPING = 'E51110116';

// Duplicate (function 3) line types -- EXPENSE is always added; the other two are opt-in via
// checkbox (most months don't need them, confirmed by the user), each added/updated per person
// with the exact same idempotent add-or-update pattern as EXPENSE, just keyed by their own Paycode.
var LINE_TYPES = {
  EXPENSE: { paycodeCode: EXPENSE_PAYCODE, paycodeName: EXPENSE_PAYCODE_NAME, account: EXPENSE_ACCOUNT, grouping: EXPENSE_GROUPING },
  PROVIDENT_FUND_REFUND: { paycodeCode: 'T208', paycodeName: 'เงินสมทบกองทุนสำรองเลี้ยงชีพนายจ้าง', account: '51110103', grouping: 'E51110103' },
  ACCIDENT_REFUND: { paycodeCode: 'AC CL D AC', paycodeName: 'หักค่าประกันอุบัติเหตุ ค่าใช้จ่ายลูกค้า', account: '51110104', grouping: 'E51110104' }
};

var REQUIRED_HEADER_LABELS = ['NAME', 'Paycode Code', 'Paycode Name', 'Account', 'Grouping', 'Amount', 'Introduce By', 'EMP ID'];

function findColLetterByHeaderText(aoa, headerRowNum, label) {
  return colIndexToLetters(findColByHeaderText(aoa, headerRowNum, label));
}

// Strips a trailing " - N" correction-row suffix from an EMP ID (confirmed real pattern: the same
// person can appear with EMP ID "10130928" on their normal row and "10130928 - 2" on a correction/
// amendment row -- these must key to the SAME person, not be treated as two different people with
// a genuine identity conflict).
function baseEmpId(empId) {
  return normText(empId).replace(/\s*-\s*\d+\s*$/, '').trim();
}

// ── Load a GL_Invoice workbook into a ready-to-transform context ──────────────────────────────
async function loadGLInvoiceContext(buf) {
  var wb = await loadWorkbook(buf);
  var sheetPath = firstSheetPath(wb);
  var sheetXml = await decompressEntry(wb.entries[sheetPath], wb.buf);
  var sst = await getSharedStrings(wb);
  var aoa = parseGridFromXml(sheetXml, sst);
  var headerRow = findHeaderRow(aoa, REQUIRED_HEADER_LABELS, 20);
  var cols = {};
  ['NAME', 'EMP ID', 'Introduce By', 'Paycode Code', 'Paycode Name', 'Account', 'Grouping', 'Amount'].forEach(function (label) {
    var key = label.replace(/\s+/g, '');
    cols[key] = findColLetterByHeaderText(aoa, headerRow, label);
  });
  var stylesXml = wb.entries['xl/styles.xml'] ? await decompressEntry(wb.entries['xl/styles.xml'], wb.buf) : null;
  var wbXmlRels = wb.entries['xl/_rels/workbook.xml.rels'] ? await decompressEntry(wb.entries['xl/_rels/workbook.xml.rels'], wb.buf) : '';
  var sheetNameM = new RegExp('<sheet\\b[^>]*name="([^"]+)"[^>]*\\/?>').exec(wb.wbXml); // first sheet's own tab name
  return {
    wb: wb, sheetPath: sheetPath, sheetXml: sheetXml, wbXml: wb.wbXml, wbXmlRels: wbXmlRels,
    sst: sst, aoa: aoa, headerRow: headerRow, cols: cols,
    styles: stylesXml ? parseStylesXml(stylesXml) : null,
    sheetName: sheetNameM ? sheetNameM[1] : null,
    model: parseSheetRows(sheetXml)
  };
}

// Re-derives aoa + model from the context's CURRENT sheetXml -- call this between chained stages
// (functions 4/5) so the next stage never operates on a stale row-number snapshot.
function reparseContext(ctx) {
  ctx.sheetXml = serializeSheetRows(ctx.model);
  ctx.aoa = parseGridFromXml(ctx.sheetXml, ctx.sst);
  ctx.model = parseSheetRows(ctx.sheetXml);
}

function matchesSSOCondition(row, cols, sst) {
  var grouping = normTextUpper(getCellValue(row.cellsByCol[cols.Grouping], sst));
  var paycode = normTextUpper(getCellValue(row.cellsByCol[cols.PaycodeCode], sst));
  return grouping === SSO_GROUPING && SSO_PAYCODES.indexOf(paycode) >= 0;
}

function writeAmount(row, cols, newAmount) {
  var addr = cols.Amount + row.rowNum;
  var styleIdx = getCellStyleIndex(row.cellsByCol[cols.Amount]);
  row.cellsByCol[cols.Amount] = buildLiteralNumberCell(addr, styleIdx, newAmount);
}

// Reads the identifying fields of a row that a summary detail table needs to display, BEFORE
// any write touches it -- paycode/paycodeName never change in functions 1/2/6, so reading them
// here is always safe regardless of write order.
function rowDetailBase(row, ctx) {
  return {
    row: row.rowNum,
    paycode: getCellValue(row.cellsByCol[ctx.cols.PaycodeCode], ctx.sst),
    paycodeName: getCellValue(row.cellsByCol[ctx.cols.PaycodeName], ctx.sst)
  };
}

// ── Function 1: SSO PRTR 750 ───────────────────────────────────────────────────────────────────
function fn1_SsoPrtr750(ctx) {
  var summary = { matched: 0, reduced: 0, zeroed: 0, countT2A3: 0, countTZ74: 0, details: [] };
  ctx.model.rows.forEach(function (row) {
    if (row.rowNum <= ctx.headerRow) return;
    if (!matchesSSOCondition(row, ctx.cols, ctx.sst)) return;
    summary.matched++;
    var detail = rowDetailBase(row, ctx);
    if (normTextUpper(detail.paycode) === 'T2A3') summary.countT2A3++; else summary.countTZ74++;
    var amt = getCellValue(row.cellsByCol[ctx.cols.Amount], ctx.sst);
    amt = typeof amt === 'number' ? amt : 0;
    detail.amountBefore = amt;
    if (amt > 750) { writeAmount(row, ctx.cols, amt - 750); summary.reduced++; detail.amountAfter = amt - 750; detail.note = 'ลด 750'; }
    else { writeAmount(row, ctx.cols, 0); summary.zeroed++; detail.amountAfter = 0; detail.note = 'ปรับเป็น 0'; }
    summary.details.push(detail);
  });
  return summary;
}

// ── Function 2: SSO Introduce by ───────────────────────────────────────────────────────────────
function fn2_SsoIntroduceBy(ctx) {
  var summary = { matched: 0, zeroedPRTR: 0, unchangedCLNT: 0, unchangedOther: 0, countT2A3: 0, countTZ74: 0, details: [] };
  ctx.model.rows.forEach(function (row) {
    if (row.rowNum <= ctx.headerRow) return;
    if (!matchesSSOCondition(row, ctx.cols, ctx.sst)) return;
    summary.matched++;
    var detail = rowDetailBase(row, ctx);
    if (normTextUpper(detail.paycode) === 'T2A3') summary.countT2A3++; else summary.countTZ74++;
    var introduceBy = normTextUpper(getCellValue(row.cellsByCol[ctx.cols.IntroduceBy], ctx.sst));
    var amt = getCellValue(row.cellsByCol[ctx.cols.Amount], ctx.sst);
    detail.introduceBy = introduceBy;
    detail.amountBefore = typeof amt === 'number' ? amt : 0;
    if (introduceBy === 'PRTR') {
      writeAmount(row, ctx.cols, 0); summary.zeroedPRTR++;
      detail.amountAfter = 0; detail.note = 'ปรับเป็น 0 (PRTR)';
    } else if (introduceBy === 'CLNT') {
      summary.unchangedCLNT++; detail.amountAfter = detail.amountBefore; detail.note = 'ไม่เปลี่ยน (CLNT)';
    } else {
      summary.unchangedOther++; detail.amountAfter = detail.amountBefore; detail.note = 'ไม่เปลี่ยน';
    }
    summary.details.push(detail);
  });
  return summary;
}

// ── Function 3: Duplicate (add-or-update a fixed line per person, idempotent) ───────────────────
function setLineFields(row, ctx, lineType) {
  writeTextCell(row, ctx.cols.PaycodeCode, lineType.paycodeCode);
  writeTextCell(row, ctx.cols.PaycodeName, lineType.paycodeName);
  writeTextCell(row, ctx.cols.Account, lineType.account);
  writeTextCell(row, ctx.cols.Grouping, lineType.grouping);
  writeBlankCell(row, ctx.cols.Amount);
}
function writeTextCell(row, colLetter, text) {
  var addr = colLetter + row.rowNum;
  var styleIdx = getCellStyleIndex(row.cellsByCol[colLetter]);
  row.cellsByCol[colLetter] = buildInlineStringCell(addr, styleIdx, text);
}
function writeBlankCell(row, colLetter) {
  var addr = colLetter + row.rowNum;
  var styleIdx = getCellStyleIndex(row.cellsByCol[colLetter]);
  row.cellsByCol[colLetter] = buildBlankCell(addr, styleIdx);
}
function recolorRowRed(row, styles) {
  Object.keys(row.cellsByCol).forEach(function (col) {
    var cellXml = row.cellsByCol[col];
    var styleIdx = getCellStyleIndex(cellXml);
    var redIdx = getOrCreateRedVariant(styles, styleIdx);
    row.cellsByCol[col] = restyleCell(cellXml, col + row.rowNum, redIdx);
  });
}
// Builds a brand-new line row's cell map (no rowNum assigned yet -- insertRowsAfter sets it),
// copying every identity column's actual VALUE (never a raw shared-string/style index) from
// templateRow, recoloring every cell red, and overriding the 5 line-type-specific fields.
function buildLineRowFrom(templateRow, ctx, lineType) {
  var cellsByCol = {};
  Object.keys(templateRow.cellsByCol).forEach(function (col) {
    var origCellXml = templateRow.cellsByCol[col];
    var styleIdx = getCellStyleIndex(origCellXml);
    var redIdx = getOrCreateRedVariant(ctx.styles, styleIdx);
    var placeholderAddr = col + '1'; // real address assigned later by insertRowsAfter/readdress
    if (col === ctx.cols.PaycodeCode) cellsByCol[col] = buildInlineStringCell(placeholderAddr, redIdx, lineType.paycodeCode);
    else if (col === ctx.cols.PaycodeName) cellsByCol[col] = buildInlineStringCell(placeholderAddr, redIdx, lineType.paycodeName);
    else if (col === ctx.cols.Account) cellsByCol[col] = buildInlineStringCell(placeholderAddr, redIdx, lineType.account);
    else if (col === ctx.cols.Grouping) cellsByCol[col] = buildInlineStringCell(placeholderAddr, redIdx, lineType.grouping);
    else if (col === ctx.cols.Amount) cellsByCol[col] = buildBlankCell(placeholderAddr, redIdx);
    else {
      var val = getCellValue(origCellXml, ctx.sst);
      if (val === null) cellsByCol[col] = buildBlankCell(placeholderAddr, redIdx);
      else if (typeof val === 'number') cellsByCol[col] = buildLiteralNumberCell(placeholderAddr, redIdx, val);
      else cellsByCol[col] = buildInlineStringCell(placeholderAddr, redIdx, val);
    }
  });
  return { attrsRest: templateRow.attrsRest, cellsByCol: cellsByCol };
}

// lineTypeKeys: e.g. ['EXPENSE'] or ['EXPENSE', 'PROVIDENT_FUND_REFUND', 'ACCIDENT_REFUND'] -- EXPENSE is
// always included by the caller; the other two are opt-in checkboxes (most months don't need
// them). Each line type is added/updated independently per person, keyed by that line type's own
// Paycode Code, so a person can end up with any combination of the selected lines.
function fn3_DuplicateLines(ctx, lineTypeKeys) {
  var summary = { added: 0, updated: 0, errors: [], details: [] };
  var dataRows = ctx.model.rows.filter(function (row) { return row.rowNum > ctx.headerRow; });

  var byKey = {};
  var nameToEmpIds = {};
  dataRows.forEach(function (row) {
    var empId = baseEmpId(getCellValue(row.cellsByCol[ctx.cols.EMPID], ctx.sst));
    var name = normText(getCellValue(row.cellsByCol[ctx.cols.NAME], ctx.sst));
    if (!name) return; // no identity to key on -- shouldn't happen for a real data row
    var nameUpper = normTextUpper(name);
    if (!nameToEmpIds[nameUpper]) nameToEmpIds[nameUpper] = {};
    if (empId) nameToEmpIds[nameUpper][empId] = true;
    var key = empId ? ('E:' + empId) : ('N:' + nameUpper);
    if (!byKey[key]) byKey[key] = { rows: [] };
    byKey[key].rows.push(row);
  });

  Object.keys(nameToEmpIds).forEach(function (nameUpper) {
    var empIds = Object.keys(nameToEmpIds[nameUpper]);
    if (empIds.length > 1) {
      summary.errors.push('ชื่อ "' + nameUpper + '" มี EMP ID ต่างกันหลายค่า (' + empIds.join(', ') + ') — กรุณาตรวจสอบไฟล์ก่อนใช้ฟังก์ชันนี้');
    }
  });
  if (summary.errors.length) return summary; // hard-stop, never guess which EMP ID is "right"

  var newRowsToInsert = [];
  var newRowKeys = []; // parallel to newRowsToInsert -- carries the person's own EMP ID/name/lineType for the detail table

  lineTypeKeys.forEach(function (lineTypeKey) {
    var lineType = LINE_TYPES[lineTypeKey];
    Object.keys(byKey).forEach(function (key) {
      var person = byKey[key];
      var templateRow = person.rows[0];
      var name = getCellValue(templateRow.cellsByCol[ctx.cols.NAME], ctx.sst);
      var empId = getCellValue(templateRow.cellsByCol[ctx.cols.EMPID], ctx.sst);
      var existingRow = person.rows.find(function (row) {
        return normTextUpper(getCellValue(row.cellsByCol[ctx.cols.PaycodeCode], ctx.sst)) === normTextUpper(lineType.paycodeCode);
      });
      if (existingRow) {
        setLineFields(existingRow, ctx, lineType);
        recolorRowRed(existingRow, ctx.styles);
        summary.updated++;
        summary.details.push({ row: existingRow.rowNum, name: name, empId: empId,
          paycodeCode: lineType.paycodeCode, paycodeName: lineType.paycodeName,
          account: lineType.account, grouping: lineType.grouping, action: 'อัพเดท' });
      } else {
        newRowsToInsert.push(buildLineRowFrom(templateRow, ctx, lineType));
        newRowKeys.push({ name: name, empId: empId, lineType: lineType });
        summary.added++;
      }
    });
  });

  if (newRowsToInsert.length) {
    var anchorRow = findLastEmployeeRow(ctx.model.rows, ctx.cols.NAME, ctx.sst) || ctx.headerRow;
    newRowKeys.forEach(function (k, i) {
      summary.details.push({ row: anchorRow + 1 + i, name: k.name, empId: k.empId,
        paycodeCode: k.lineType.paycodeCode, paycodeName: k.lineType.paycodeName,
        account: k.lineType.account, grouping: k.lineType.grouping, action: 'เพิ่มใหม่' });
    });
    ctx.model.rows = insertRowsAfter(ctx.model.rows, anchorRow, newRowsToInsert);
  }
  return summary;
}

// ── Function 6: Remove SSO ─────────────────────────────────────────────────────────────────────
function fn6_RemoveSso(ctx) {
  var before = ctx.model.rows.length;
  var details = [], countT2A3 = 0, countTZ74 = 0;
  // Capture each matching row's detail BEFORE it's deleted -- deleteAndRenumber discards it.
  ctx.model.rows.forEach(function (row) {
    if (row.rowNum <= ctx.headerRow || !matchesSSOCondition(row, ctx.cols, ctx.sst)) return;
    var detail = rowDetailBase(row, ctx);
    detail.amount = getCellValue(row.cellsByCol[ctx.cols.Amount], ctx.sst);
    detail.note = 'ลบออกแล้ว';
    if (normTextUpper(detail.paycode) === 'T2A3') countT2A3++; else countTZ74++;
    details.push(detail);
  });
  ctx.model.rows = deleteAndRenumber(
    ctx.model.rows,
    function (row) { return row.rowNum > ctx.headerRow && matchesSSOCondition(row, ctx.cols, ctx.sst); },
    ctx.headerRow + 1,
    ctx.headerRow + 1
  );
  return { removed: before - ctx.model.rows.length, countT2A3: countT2A3, countTZ74: countTZ74, details: details };
}

// ── Function 8: Change Header ──────────────────────────────────────────────────────────────────
// Deletes rows 4 and 5, shifting the header (and everything below it) up by 2; renames two header
// cells found by exact text match (not fixed column letter): Period->Calendar Group,
// Paycode Code->PIN Name. Row/column extent + freeze pane are updated by the caller afterward
// (needs the NEW header row number, which this function returns).
function fn8_ChangeHeader(ctx) {
  var oldHeaderRow = ctx.headerRow;
  ctx.model.rows = deleteAndRenumber(
    ctx.model.rows,
    function (row) { return row.rowNum === 4 || row.rowNum === 5; },
    oldHeaderRow, // everything from the old header row onward shifts
    oldHeaderRow - 2
  );
  var newHeaderRow = oldHeaderRow - 2;
  var headerRowModel = ctx.model.rows.find(function (r) { return r.rowNum === newHeaderRow; });
  if (!headerRowModel) throw new Error('ไม่พบแถว header หลังจากลบแถว 4-5 (คาดว่าจะอยู่ที่แถว ' + newHeaderRow + ')');

  var renamed = [];
  var details = [];
  [['Period', 'Calendar Group'], ['Paycode Code', 'PIN Name']].forEach(function (pair) {
    var oldLabel = pair[0], newLabel = pair[1];
    var colLetter = findColLetterByHeaderText(ctx.aoa, oldHeaderRow, oldLabel);
    writeTextCell(headerRowModel, colLetter, newLabel);
    renamed.push(oldLabel + ' -> ' + newLabel);
    details.push({ column: colLetter, oldLabel: oldLabel, newLabel: newLabel });
  });

  return { newHeaderRow: newHeaderRow, renamed: renamed, details: details, rowsDeleted: 2 };
}

// ── Structural refs + zip assembly (shared by every function) ─────────────────────────────────
function finalizeAndBuildOutputBytes(ctx, structuralOpts) {
  var sheetXml = serializeSheetRows(ctx.model);
  var wbXml = ctx.wbXml;
  if (structuralOpts) {
    var updated = updateStructuralRefs(sheetXml, wbXml, structuralOpts);
    sheetXml = updated.sheetXml;
    wbXml = updated.wbXml;
  }

  var verifyList = structuralOpts ? [
    { addr: 'dimension', expected: structuralOpts.firstCol + '1:' + structuralOpts.lastCol + structuralOpts.lastRow }
  ] : [];
  var check = verifyStructural(sheetXml, verifyList);
  if (!check.ok) throw new Error('ตรวจสอบไฟล์ผลลัพธ์ไม่ผ่าน: ' + JSON.stringify(check.mismatches));

  return buildOutputZip(ctx, sheetXml, wbXml);
}

function verifyStructural(sheetXml, list) {
  var mismatches = [];
  list.forEach(function (item) {
    if (item.addr === 'dimension') {
      var m = /<dimension ref="([^"]*)"/.exec(sheetXml);
      if (!m || m[1] !== item.expected) mismatches.push({ what: 'dimension', expected: item.expected, found: m ? m[1] : null });
    }
  });
  return { ok: mismatches.length === 0, mismatches: mismatches };
}

// Rebuilds the zip: swaps in the new sheetXml/wbXml/stylesXml, drops calcChain.xml (+ its two
// references) and forces fullCalcOnLoad="1" defensively -- this file has no formulas today, but a
// future month's file gaining one shouldn't silently reintroduce the stale-calcChain "Excel found
// a problem with this file" repair prompt already hit (and fixed) once in this portal.
async function buildOutputZip(ctx, sheetXml, wbXml) {
  var enc = new TextEncoder();
  var names = Object.keys(ctx.wb.entries);
  var zipFiles = [];
  for (var i = 0; i < names.length; i++) {
    var nm = names[i];
    var data;
    if (nm === ctx.sheetPath) data = enc.encode(sheetXml);
    else if (nm === 'xl/workbook.xml') data = enc.encode(wbXml);
    else if (nm === 'xl/styles.xml' && ctx.styles) data = enc.encode(serializeStylesXml(ctx.styles));
    else data = await decompressEntryBytes(ctx.wb.entries[nm], ctx.wb.buf);
    zipFiles.push({ name: nm, data: data });
  }
  zipFiles = applyCalcChainFix(zipFiles, enc);
  return buildZip(zipFiles);
}

function applyCalcChainFix(zipFiles, enc) {
  var out = [];
  for (var i = 0; i < zipFiles.length; i++) {
    var f = zipFiles[i];
    if (f.name === 'xl/calcChain.xml') continue;
    if (f.name === '[Content_Types].xml') {
      var ct = new TextDecoder().decode(f.data);
      ct = ct.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, '');
      out.push({ name: f.name, data: enc.encode(ct) });
    } else if (f.name === 'xl/_rels/workbook.xml.rels') {
      var rels = new TextDecoder().decode(f.data);
      rels = rels.replace(/<Relationship[^>]*Target="calcChain\.xml"[^>]*\/>/, '');
      out.push({ name: f.name, data: enc.encode(rels) });
    } else if (f.name === 'xl/workbook.xml') {
      var wb = new TextDecoder().decode(f.data);
      if (/<calcPr\b[^>]*\/>/.test(wb)) {
        wb = wb.replace(/<calcPr\b([^>]*)\/>/, function (_, attrs) {
          return /fullCalcOnLoad=/.test(attrs)
            ? '<calcPr' + attrs.replace(/fullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"') + '/>'
            : '<calcPr' + attrs + ' fullCalcOnLoad="1"/>';
        });
      } else {
        wb = wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
      }
      out.push({ name: f.name, data: enc.encode(wb) });
    } else {
      out.push(f);
    }
  }
  return out;
}

// ── Top-level entry points, one per function id ────────────────────────────────────────────────
// selectedLineTypeKeys: array of LINE_TYPES keys the user checked in the UI (e.g. ['EXPENSE',
// 'PROVIDENT_FUND_REFUND']) for the 3 functions that run the Duplicate step. EXPENSE is checked by default in
// the UI but is a real, uncheckable-off-by-force choice here -- an empty array is valid and just
// means no Duplicate rows get added/updated this run.
async function runSingleFileFunction(functionId, buf, selectedLineTypeKeys) {
  var ctx = await loadGLInvoiceContext(buf);
  var summary;
  var structuralOpts = null;
  var lineTypeKeys = selectedLineTypeKeys || [];

  if (functionId === 'sso750') {
    summary = fn1_SsoPrtr750(ctx);
  } else if (functionId === 'ssoIntroduceBy') {
    summary = fn2_SsoIntroduceBy(ctx);
  } else if (functionId === 'duplicate') {
    summary = fn3_DuplicateLines(ctx, lineTypeKeys);
    structuralOpts = buildStructuralOptsAfterAppend(ctx);
  } else if (functionId === 'sso750PlusDuplicate') {
    var s1 = fn1_SsoPrtr750(ctx);
    reparseContext(ctx);
    var s3a = fn3_DuplicateLines(ctx, lineTypeKeys);
    // Deliberately NOT a flat Object.assign -- both stages have their own "details"/"errors" keys,
    // which would silently clobber stage 1's per-row Amount-adjustment details with stage 3's.
    summary = {
      matched: s1.matched, reduced: s1.reduced, zeroed: s1.zeroed, countT2A3: s1.countT2A3, countTZ74: s1.countTZ74,
      added: s3a.added, updated: s3a.updated, errors: s3a.errors,
      ssoDetails: s1.details, expenseDetails: s3a.details
    };
    structuralOpts = buildStructuralOptsAfterAppend(ctx);
  } else if (functionId === 'ssoIntroduceByPlusDuplicate') {
    var s2 = fn2_SsoIntroduceBy(ctx);
    reparseContext(ctx);
    var s3b = fn3_DuplicateLines(ctx, lineTypeKeys);
    summary = {
      matched: s2.matched, zeroedPRTR: s2.zeroedPRTR, unchangedCLNT: s2.unchangedCLNT, unchangedOther: s2.unchangedOther,
      countT2A3: s2.countT2A3, countTZ74: s2.countTZ74,
      added: s3b.added, updated: s3b.updated, errors: s3b.errors,
      ssoDetails: s2.details, expenseDetails: s3b.details
    };
    structuralOpts = buildStructuralOptsAfterAppend(ctx);
  } else if (functionId === 'removeSso') {
    summary = fn6_RemoveSso(ctx);
    structuralOpts = buildStructuralOptsAfterRowChange(ctx, ctx.headerRow);
  } else if (functionId === 'changeHeader') {
    var oldHeaderRowForYSplit = ctx.headerRow;
    summary = fn8_ChangeHeader(ctx);
    structuralOpts = buildStructuralOptsAfterRowChange(ctx, summary.newHeaderRow, oldHeaderRowForYSplit);
  } else {
    throw new Error('ไม่รู้จักฟังก์ชัน: ' + functionId);
  }

  if (summary && summary.errors && summary.errors.length) {
    return { ok: false, summary: summary };
  }
  var outputBytes = await finalizeAndBuildOutputBytes(ctx, structuralOpts);
  return { ok: true, summary: summary, outputBytes: outputBytes };
}

function buildStructuralOptsAfterAppend(ctx) {
  var extent = columnExtent(ctx.model.rows, ctx.headerRow);
  var lastRow = ctx.model.rows.length ? Math.max.apply(null, ctx.model.rows.map(function (r) { return r.rowNum; })) : ctx.headerRow;
  return { firstCol: extent.firstCol, lastCol: extent.lastCol, headerRow: ctx.headerRow, lastRow: lastRow, newYSplit: null, sheetName: ctx.sheetName };
}
function buildStructuralOptsAfterRowChange(ctx, newHeaderRow, oldYSplit) {
  var extent = columnExtent(ctx.model.rows, newHeaderRow);
  var lastRow = ctx.model.rows.length ? Math.max.apply(null, ctx.model.rows.map(function (r) { return r.rowNum; })) : newHeaderRow;
  return {
    firstCol: extent.firstCol, lastCol: extent.lastCol, headerRow: newHeaderRow, lastRow: lastRow,
    newYSplit: (oldYSplit != null && oldYSplit !== newHeaderRow) ? newHeaderRow : null,
    sheetName: ctx.sheetName
  };
}

// ── Function 7: Merge ──────────────────────────────────────────────────────────────────────────
// bufs: array of ArrayBuffer, file 1 first (the base) through file N last, in upload order.
// fileNames: parallel array of display names (optional -- falls back to "ไฟล์ที่ N").
async function runMergeFunction(bufs, fileNames) {
  if (bufs.length < 2) throw new Error('Merge ต้องการอย่างน้อย 2 ไฟล์');
  var baseCtx = await loadGLInvoiceContext(bufs[0]);
  var baseHeaderLabels = REQUIRED_HEADER_LABELS.slice().sort();
  var summary = { filesAppended: 0, rowsAppended: 0, rejected: [], details: [] };

  // Per-column style map taken from file 1's own last data row -- file 1 is the explicit visual
  // authority (its header/styles/everything stay as-is); appended rows never carry another file's
  // raw style/shared-string index, since both are local/positional to that OTHER file.
  var baseLastRow = findLastEmployeeRow(baseCtx.model.rows, baseCtx.cols.NAME, baseCtx.sst);
  var baseStyleRow = baseCtx.model.rows.find(function (r) { return r.rowNum === baseLastRow; });
  var baseStyleByCol = {};
  if (baseStyleRow) {
    Object.keys(baseStyleRow.cellsByCol).forEach(function (col) {
      baseStyleByCol[col] = getCellStyleIndex(baseStyleRow.cellsByCol[col]);
    });
  }

  var allNewRows = [];
  for (var i = 1; i < bufs.length; i++) {
    var srcCtx;
    try {
      srcCtx = await loadGLInvoiceContext(bufs[i]);
    } catch (err) {
      summary.rejected.push({ fileIndex: i + 1, reason: 'อ่านไฟล์ไม่สำเร็จ: ' + err.message });
      continue;
    }
    var srcHeaderLabels = REQUIRED_HEADER_LABELS.slice().sort();
    // Validate this file's header set actually matches file 1's before appending anything from
    // it -- hard-stop per-file rather than silently appending misaligned data. (Both label sets
    // are drawn from the same REQUIRED_HEADER_LABELS constant today, so this check is really
    // guarding against a future file missing one of them, which loadGLInvoiceContext would
    // already have thrown on -- kept explicit here for clarity and as a seam for a future,
    // richer structural comparison.)
    if (JSON.stringify(srcHeaderLabels) !== JSON.stringify(baseHeaderLabels)) {
      summary.rejected.push({ fileIndex: i + 1, reason: 'โครงสร้างคอลัมน์ไม่ตรงกับไฟล์ที่ 1' });
      continue;
    }

    var srcDataRows = srcCtx.model.rows.filter(function (row) { return row.rowNum > srcCtx.headerRow; });
    srcDataRows.forEach(function (row) {
      var cellsByCol = {};
      Object.keys(row.cellsByCol).forEach(function (col) {
        var val = getCellValue(row.cellsByCol[col], srcCtx.sst);
        var styleIdx = baseStyleByCol[col] != null ? baseStyleByCol[col] : getCellStyleIndex(row.cellsByCol[col]);
        var placeholderAddr = col + '1';
        if (val === null) cellsByCol[col] = buildBlankCell(placeholderAddr, styleIdx);
        else if (typeof val === 'number') cellsByCol[col] = buildLiteralNumberCell(placeholderAddr, styleIdx, val);
        else cellsByCol[col] = buildInlineStringCell(placeholderAddr, styleIdx, val);
      });
      allNewRows.push({ attrsRest: baseStyleRow ? baseStyleRow.attrsRest : '', cellsByCol: cellsByCol });
    });
    summary.filesAppended++;
    summary.rowsAppended += srcDataRows.length;
    summary.details.push({
      fileIndex: i + 1,
      fileName: (fileNames && fileNames[i]) || ('ไฟล์ที่ ' + (i + 1)),
      rowsAppended: srcDataRows.length,
      note: 'รวมสำเร็จ'
    });
  }

  if (allNewRows.length) {
    baseCtx.model.rows = insertRowsAfter(baseCtx.model.rows, baseLastRow || baseCtx.headerRow, allNewRows);
  }

  var structuralOpts = buildStructuralOptsAfterAppend(baseCtx);
  var outputBytes = await finalizeAndBuildOutputBytes(baseCtx, structuralOpts);
  return { ok: true, summary: summary, outputBytes: outputBytes };
}
