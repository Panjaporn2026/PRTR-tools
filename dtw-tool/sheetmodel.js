// ══════════════════════════════════════════════════════
//  Shared row-level engine the DTW Head/Line write path builds on. Ported from
//  gl-invoice/sheetmodel.js. Confirmed real-file facts this depends on (the blank
//  "Fill in data for DTW.xlsx" template + the completed DTW_SIEMENS_June2026_GLmapping.xlsx
//  ground truth): Head/Line have NO autoFilter and NO freeze pane -- only <dimension> needs
//  keeping in sync after appending rows (unlike gl-invoice's sheet, which has both). Rows 1-3 on
//  both sheets are the fixed system/SAP/Thai header block (carries cell comments anchored to row
//  1 via comments1.xml/comments2.xml) and must never be touched; rows 4+ are sample/data rows.
// ══════════════════════════════════════════════════════

function esc(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function normText(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
}
function normTextUpper(s) {
  return normText(s).toUpperCase();
}

// ── Header / column discovery (never guess silently: hard error on 0 or >1 match) ─────────────

// Find the 1-based row number, within the first `maxRow` rows of aoa, that contains ALL of the
// given exact (normalized) header labels as distinct cell values. Throws if zero or more than one
// row qualifies.
function findHeaderRow(aoa, requiredLabels, maxRow) {
  var requiredUpper = requiredLabels.map(normTextUpper);
  var candidates = [];
  for (var r = 0; r < Math.min(maxRow, aoa.length); r++) {
    var row = aoa[r] || [];
    var foundSet = {};
    for (var c = 0; c < row.length; c++) {
      var v = normTextUpper(row[c]);
      if (v) foundSet[v] = true;
    }
    if (requiredUpper.every(function (lbl) { return foundSet[lbl]; })) candidates.push(r + 1);
  }
  if (candidates.length === 0) {
    throw new Error('ไม่พบแถว header ที่มีคอลัมน์ครบตามที่ต้องการ: ' + requiredLabels.join(', '));
  }
  if (candidates.length > 1) {
    throw new Error('พบแถว header ที่ตรงเงื่อนไขมากกว่า 1 แถว (แถว ' + candidates.join(', ') + ') กรุณาตรวจสอบไฟล์');
  }
  return candidates[0];
}

// Find the 0-based column index of an exact header label within a specific 1-based row of aoa.
// Throws if zero or more than one column matches.
function findColByHeaderText(aoa, headerRowNum, label) {
  var row = aoa[headerRowNum - 1] || [];
  var target = normTextUpper(label);
  var matches = [];
  for (var c = 0; c < row.length; c++) {
    if (normTextUpper(row[c]) === target) matches.push(c);
  }
  if (matches.length === 0) throw new Error('ไม่พบคอลัมน์ "' + label + '" ในแถว header (แถว ' + headerRowNum + ')');
  if (matches.length > 1) throw new Error('พบคอลัมน์ "' + label + '" มากกว่า 1 คอลัมน์ในแถว header (แถว ' + headerRowNum + ')');
  return matches[0];
}

// ── Row model: parse <sheetData> into an addressable, order-preserving structure ──────────────
//
// Each row keeps its own raw per-cell XML strings (not a re-derived value+style split) so any
// cell this tool never touches is carried forward completely byte-identical. Only cells a
// function explicitly needs to change get their raw XML string replaced.

function parseSheetRows(sheetXml) {
  var sdStart = sheetXml.indexOf('<sheetData');
  var sdTagEnd = sheetXml.indexOf('>', sdStart) + 1;
  var sdEnd = sheetXml.indexOf('</sheetData>');
  var pre = sheetXml.slice(0, sdTagEnd);
  var sdBody = sheetXml.slice(sdTagEnd, sdEnd);
  var post = sheetXml.slice(sdEnd); // "</sheetData>" onward

  var rows = [];
  // Handles both self-closing <row .../> (blank rows) and full <row ...>...</row> forms --
  // matching only the full form silently swallows content through to the NEXT row's closing tag
  // when an intervening row is self-closing.
  var rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g, rm;
  while ((rm = rowRe.exec(sdBody)) !== null) {
    var fullAttrs = rm[1];
    var rNumM = /\br="(\d+)"/.exec(fullAttrs);
    if (!rNumM) continue; // malformed row with no r attribute -- skip defensively
    var rowNum = parseInt(rNumM[1], 10);
    var attrsRest = fullAttrs.replace(/\s*\br="\d+"/, ''); // keep original order/spacing of the rest
    var rowInner = rm[2];
    var cellsByCol = {};
    if (rowInner !== undefined) {
      var cellRe = /<c\s+r="([A-Z]+)\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g, cm;
      while ((cm = cellRe.exec(rowInner)) !== null) cellsByCol[cm[1]] = cm[0];
    }
    rows.push({ rowNum: rowNum, attrsRest: attrsRest, cellsByCol: cellsByCol });
  }
  return { pre: pre, post: post, rows: rows };
}

function serializeSheetRows(model) {
  var rowsXml = model.rows.map(function (row) {
    var cellsXml = Object.keys(row.cellsByCol)
      .sort(function (a, b) { return colLettersToIndex(a) - colLettersToIndex(b); })
      .map(function (col) { return readdressCell(row.cellsByCol[col], col + row.rowNum); })
      .join('');
    return '<row r="' + row.rowNum + '"' + row.attrsRest + '>' + cellsXml + '</row>';
  }).join('');
  return model.pre + rowsXml + model.post;
}

// Replaces just the r="..." address on a raw cell XML string with a new address, regardless of
// what the cell's previous address was -- used whenever a row is renumbered.
function readdressCell(rawCellXml, newAddr) {
  return rawCellXml.replace(/^<c\s+r="[A-Z]+\d+"/, '<c r="' + newAddr + '"');
}

// Swaps a cell's s="N" style attribute (adding one if it had none) while leaving its address,
// type, and value completely untouched.
function restyleCell(rawCellXml, newAddr, newStyleIndex) {
  var withAddr = readdressCell(rawCellXml, newAddr);
  if (/\bs="\d+"/.test(withAddr)) return withAddr.replace(/\bs="\d+"/, 's="' + newStyleIndex + '"');
  return withAddr.replace(/^(<c\s+r="[^"]+")/, '$1 s="' + newStyleIndex + '"');
}

// Inserts newRowsData (array of {attrsRest, cellsByCol}, no rowNum yet) immediately after the row
// currently numbered afterRowNum, shifting every row after that point down to make room. This is
// how Head/Line grow beyond their original 2/4 sample rows -- inserting after a specific anchor
// row (the last sample/data row) rather than blindly at the physical end of the array.
function insertRowsAfter(rows, afterRowNum, newRowsData) {
  var idx = rows.findIndex(function (r) { return r.rowNum === afterRowNum; });
  var insertAt = idx >= 0 ? idx + 1 : rows.length;
  var startNum = afterRowNum + 1;
  var before = rows.slice(0, insertAt);
  var after = rows.slice(insertAt);
  var newRows = newRowsData.map(function (data, i) {
    return { rowNum: startNum + i, attrsRest: data.attrsRest, cellsByCol: data.cellsByCol };
  });
  var afterRenumbered = after.map(function (row) {
    return { rowNum: row.rowNum + newRowsData.length, attrsRest: row.attrsRest, cellsByCol: row.cellsByCol };
  });
  return before.concat(newRows, afterRenumbered);
}

// ── Cell read/write helpers ─────────────────────────────────────────────────────────────────

function getCellStyleIndex(rawCellXml) {
  if (!rawCellXml) return 0;
  var m = /\bs="(\d+)"/.exec(rawCellXml);
  return m ? parseInt(m[1], 10) : 0;
}

// Resolves one raw cell XML string's actual value: shared-string index -> the real string,
// inline string -> its text, plain <v> -> a number if parseable else the raw string, no <v> ->
// null (a true blank).
function getCellValue(rawCellXml, sst) {
  if (!rawCellXml) return null;
  var tMatch = /\bt="([^"]+)"/.exec(rawCellXml);
  var type = tMatch ? tMatch[1] : null;
  var innerM = /^<c\b[^>]*>([\s\S]*)<\/c>$/.exec(rawCellXml);
  var inner = innerM ? innerM[1] : '';
  if (type === 's') {
    var vM = /<v>([\s\S]*?)<\/v>/.exec(inner);
    return vM ? (sst[parseInt(vM[1], 10)] != null ? sst[parseInt(vM[1], 10)] : null) : null;
  }
  if (type === 'inlineStr') {
    var isM = /<is>([\s\S]*?)<\/is>/.exec(inner);
    var tM = isM ? /<t[^>]*>([\s\S]*?)<\/t>/.exec(isM[1]) : null;
    return tM ? decodeXmlEntities(tM[1]) : null;
  }
  if (type === 'str') { // formula cached string result
    var vM2 = /<v>([\s\S]*?)<\/v>/.exec(inner);
    return vM2 ? decodeXmlEntities(vM2[1]) : null;
  }
  var vM3 = /<v>([\s\S]*?)<\/v>/.exec(inner);
  if (!vM3) return null;
  var n = parseFloat(vM3[1]);
  return isNaN(n) ? vM3[1] : n;
}

function buildLiteralNumberCell(addr, styleIndex, num) {
  var sAttr = styleIndex ? ' s="' + styleIndex + '"' : '';
  return '<c r="' + addr + '"' + sAttr + '><v>' + num + '</v></c>';
}
function buildInlineStringCell(addr, styleIndex, text) {
  var sAttr = styleIndex ? ' s="' + styleIndex + '"' : '';
  return '<c r="' + addr + '"' + sAttr + ' t="inlineStr"><is><t>' + esc(text) + '</t></is></c>';
}
// A true blank cell -- no <v> at all, not <v></v> or <v>0</v>.
function buildBlankCell(addr, styleIndex) {
  var sAttr = styleIndex ? ' s="' + styleIndex + '"' : '';
  return '<c r="' + addr + '"' + sAttr + '/>';
}

// ── Structural references: Head/Line only need <dimension> kept in sync after appends --
//    confirmed no autoFilter, no freeze pane on either sheet in the real template. ──────────────

function updateDimension(sheetXml, firstCol, lastCol, lastRow) {
  var newDim = firstCol + '1:' + lastCol + lastRow;
  return sheetXml.replace(/<dimension ref="[^"]*"\s*\/>/, '<dimension ref="' + newDim + '"/>');
}

// Last row (by original rowNum, ascending order assumed) whose column-letter cell resolves to a
// non-empty value -- the anchor for "where the existing sample/data rows end" on Head or Line.
function findLastDataRow(rows, colLetter, sst) {
  var last = null;
  rows.forEach(function (row) {
    var cellXml = row.cellsByCol[colLetter];
    if (!cellXml) return;
    var val = getCellValue(cellXml, sst);
    if (val !== null && String(val).trim() !== '') last = row.rowNum;
  });
  return last;
}

// First/last column LETTERS actually used across a row model's cells, derived from a given
// header row (column count never changes for Head/Line).
function columnExtent(rows, headerRowNum) {
  var headerRow = rows.find(function (r) { return r.rowNum === headerRowNum; });
  var cols = headerRow ? Object.keys(headerRow.cellsByCol) : [];
  if (!cols.length) throw new Error('ไม่พบคอลัมน์ใดๆ ในแถว header (แถว ' + headerRowNum + ')');
  var indices = cols.map(colLettersToIndex);
  return { firstCol: colIndexToLetters(Math.min.apply(null, indices)), lastCol: colIndexToLetters(Math.max.apply(null, indices)) };
}
