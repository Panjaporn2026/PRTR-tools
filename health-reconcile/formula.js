// ══════════════════════════════════════════════════════
//  Shared-formula resolver + cross-sheet-safe / dollar-safe row shifter.
//
//  The Detail Reconcile sheet's U:CY block is NOT the simple same-row-only case the sibling
//  accrued-income tool's formula copier assumes. Confirmed against the real workbook:
//   - Most of the block (Y, AD, AI, AJ, AP, AU, AZ, BA, BG, BL, BQ, BR, BT:BZ, CB:CH, CJ:CP, CX,
//     CY) is stored as ONE shared formula per column: a master cell with full text
//     (a cell like Y6 holds t="shared" ref="Y6:Y37" si="4" with the real formula text) and slave
//     cells with only an si reference and no text at all. Reading a slave cell directly yields
//     nothing to copy -- the master's text must be resolved via si first.
//   - Formulas reference OTHER sheets by whole-column ('PRU2024'!A:A) or, in general, could be
//     row-specific ('TL2025'!F25) refs -- these must never have their row shifted, since they
//     point into a different sheet's row-space, unrelated to the employee row being copied.
//   - Formulas mix relative rows ($AI6, $N6 -- column absolute, row relative, MEANT to shift)
//     with fully-absolute refs ($BY$2 -- a fixed header row, must NOT shift).
// ══════════════════════════════════════════════════════

var NUL = String.fromCharCode(0);

// Build a map of si -> { formula: master text, row: the master cell's OWN row number }, by
// scanning the whole sheet body once. Excel always writes the master (the cell holding the
// actual formula text) before any slave that shares its si, so the first occurrence wins.
// The master's row matters as much as its text: a slave cell's formula is anchored to the
// MASTER's row, not to whatever row the slave cell happens to sit on. Reanchoring a slave's
// resolved formula using the slave's own row (instead of the master's) would be a no-op for
// any row-shift regex looking for the slave's row number, since that number never appears in
// the master's text -- silently leaving every shared-formula copy unshifted.
function buildSharedFormulaMap(sheetBodyXml) {
  var map = {};
  var re = /<c\s+r="[A-Z]+(\d+)"[^>]*>\s*<f\s+t="shared"[^>]*\bsi="(\d+)"[^>]*>([\s\S]*?)<\/f>/g;
  var m;
  while ((m = re.exec(sheetBodyXml)) !== null) {
    var row = parseInt(m[1], 10), si = m[2], text = m[3];
    if (!(si in map)) map[si] = { formula: text, row: row };
  }
  return map;
}

// Parse one row's raw XML (the "<row ...>...</row>" inner content) into
// { col: {style, formula, value, type, isError} }.
function parseRowCells(rowInnerXml, sharedMap) {
  var cells = {};
  var re = /<c\s+r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  var m;
  while ((m = re.exec(rowInnerXml)) !== null) {
    var col = m[1];
    var attrs = m[2] || '';
    var inner = m[3] || '';
    var sMatch = /\bs="(\d+)"/.exec(attrs);
    var tMatch = /\st="([^"]+)"/.exec(attrs);
    var vMatch = /<v>([\s\S]*?)<\/v>/.exec(inner);
    var fSharedRef = /<f\s+t="shared"[^>]*\bsi="(\d+)"[^>]*\/>/.exec(inner); // slave, self-closing
    var fFull = /<f\b[^>]*>([\s\S]*?)<\/f>/.exec(inner); // master or plain formula (has text)
    var formula = null;
    var formulaFromRow = null; // row the formula text is actually anchored to, for reanchoring
    if (fFull) {
      formula = fFull[1]; // this row owns its own text (plain formula, or the shared master itself)
    } else if (fSharedRef && sharedMap && sharedMap[fSharedRef[1]] !== undefined) {
      formula = sharedMap[fSharedRef[1]].formula;
      formulaFromRow = sharedMap[fSharedRef[1]].row; // NOT this (slave) cell's own row
    }
    cells[col] = {
      style: sMatch ? sMatch[1] : null,
      type: tMatch ? tMatch[1] : null,
      formula: formula,
      formulaFromRow: formulaFromRow,
      value: vMatch ? vMatch[1] : null,
      isError: !!(tMatch && tMatch[1] === 'e')
    };
  }
  return cells;
}

// Re-anchor a same-sheet formula from fromRow to toRow, leaving cross-sheet references and
// fully-dollar-absolute rows untouched.
//
// Strategy: first mask out every cross-sheet-qualified reference (quoted 'Sheet Name'!... or
// bare SheetName!... token, single cell or range) so the row-shift regex never sees inside a
// sheet name or a foreign-sheet row number. Then shift any <col><row> token in what's left,
// UNLESS the row digit itself is immediately preceded by $ (fully absolute row, e.g. the 2 in
// $BY$2). Finally restore the masked spans unchanged.
function reanchorFormula(formula, fromRow, toRow) {
  var placeholders = [];
  var crossSheetRe = /(?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Z]{1,3}\$?\d*(?::\$?[A-Z]{1,3}\$?\d*)?/g;

  var masked = formula.replace(crossSheetRe, function (match) {
    placeholders.push(match);
    return NUL + (placeholders.length - 1) + NUL;
  });

  var rowShiftRe = new RegExp('(\\$?[A-Z]{1,3})(\\$?)(' + fromRow + ')\\b', 'g');
  masked = masked.replace(rowShiftRe, function (full, colPart, rowDollar) {
    if (rowDollar === '$') return full; // fully-absolute row -- a fixed constant row, do not shift
    return colPart + rowDollar + toRow;
  });

  var restoreRe = new RegExp(NUL + '(\\d+)' + NUL, 'g');
  masked = masked.replace(restoreRe, function (_, idx) { return placeholders[+idx]; });
  return masked;
}

// Build the full set of cell-write instructions for copying templateRow's formula/style
// pattern onto newRow, given a pre-parsed cell map for the template row (from parseRowCells).
// overrides maps column letter -> a plain value/blank instruction that should win instead of
// the template's formula (e.g. identity fields, plan codes, active-days-forced-to-0, real
// Invoice amounts) -- anything not in overrides gets the template's style + re-anchored formula,
// with NO cached value (we cannot evaluate XLOOKUP client-side; an omitted value recalculates
// honestly-blank on open instead of showing the template employee's real number until then).
function buildCopiedRowCells(templateCells, templateRow, newRow, overrides) {
  var out = {};
  Object.keys(templateCells).forEach(function (col) {
    var t = templateCells[col];
    if (overrides && Object.prototype.hasOwnProperty.call(overrides, col)) {
      out[col] = { style: t.style, override: overrides[col] };
      return;
    }
    if (t.formula) {
      // Shared-formula slave cells must reanchor from the MASTER's row, not from the row the
      // template happens to sit on -- see buildSharedFormulaMap/parseRowCells above.
      var fromRow = t.formulaFromRow != null ? t.formulaFromRow : templateRow;
      out[col] = { style: t.style, formula: reanchorFormula(t.formula, fromRow, newRow) };
    } else if (t.value !== null && t.type !== 's') {
      // plain numeric/date literal on the template row (e.g. a fixed "bridge period" figure) --
      // copy as a literal value, matching the template exactly, not derived per-employee.
      out[col] = { style: t.style, literal: t.value };
    } else {
      out[col] = { style: t.style, blank: true };
    }
  });
  return out;
}
