// ══════════════════════════════════════════════════════
//  Steps 1-3: locate the bundled health-insurance total column in a monthly invoice file
//  (column letter drifts every month -- confirmed GY in Sep'25 -> HE in Jan'26, a 149-column
//  shift over 4 months -- so it must be found by header-text search on row 25 every time, never
//  hardcoded), sum duplicate Alternate ID rows (signed), and match against the reconcile sheet.
// ══════════════════════════════════════════════════════

var INVOICE_HEADER_ROW = 25; // "Detail of Invoice" sheet's real working header row
var ALT_ID_HEADER_TEXT = 'ALTERNATE ID';

function normHeader(s) {
  return String(s || '').toUpperCase().replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
}
function normAltId(s) {
  return String(s == null ? '' : s).trim().toUpperCase();
}

// Scan a header row for every column whose normalized text contains all of `mustContain`.
// Returns an array of {col, header} so the caller can force a manual pick when there isn't
// exactly one match -- this drift is the single largest silent-failure surface in the tool.
function findHeaderCandidates(aoa, headerRowIdx, mustContain) {
  var row = aoa[headerRowIdx] || [];
  var out = [];
  for (var c = 0; c < row.length; c++) {
    var h = normHeader(row[c]);
    if (!h) continue;
    var ok = mustContain.every(function (kw) { return h.indexOf(kw) >= 0; });
    if (ok) out.push({ col: c, header: row[c] });
  }
  return out;
}

function findAlternateIdColumn(aoa, headerRowIdx) {
  var candidates = findHeaderCandidates(aoa, headerRowIdx, ['ALTERNATE', 'ID']);
  return candidates.length ? candidates[0].col : -1;
}

// Locate the "HEALTH INSURANCE EMPLOYER TOTAL"-style bundled column. Returns:
//  { status: 'ok', col, header, sum, sampleValues }
//  { status: 'ambiguous'|'not_found', candidates: [...] } -- caller MUST surface this for a
//  manual pick rather than silently guessing.
function findBundledHealthColumn(aoa, headerRowIdx, altIdCol) {
  var strict = findHeaderCandidates(aoa, headerRowIdx, ['INSURANCE', 'TOTAL', 'EMPLOYER']);
  var candidates = strict.length ? strict : findHeaderCandidates(aoa, headerRowIdx, ['INSURANCE', 'TOTAL']);
  if (candidates.length !== 1) {
    // Attach sample data + a running sum for each candidate so a manual picker has context.
    var withSamples = candidates.map(function (c) {
      var samples = [], sum = 0, n = 0;
      for (var r = headerRowIdx + 1; r < aoa.length; r++) {
        var v = (aoa[r] || [])[c.col];
        if (typeof v === 'number') { sum += v; n++; if (samples.length < 3) samples.push(v); }
      }
      return { col: c.col, header: c.header, sum: sum, count: n, sampleValues: samples };
    });
    return { status: candidates.length === 0 ? 'not_found' : 'ambiguous', candidates: withSamples };
  }
  return useBundledColumn(aoa, headerRowIdx, altIdCol, candidates[0].col, candidates[0].header);
}

function useBundledColumn(aoa, headerRowIdx, altIdCol, col, header) {
  var sum = 0, samples = [];
  for (var r = headerRowIdx + 1; r < aoa.length; r++) {
    var v = (aoa[r] || [])[col];
    if (typeof v === 'number') { sum += v; if (samples.length < 3) samples.push(v); }
  }
  return { status: 'ok', col: col, header: header, sum: sum, sampleValues: samples };
}

// Sum the bundled amount per Alternate ID for one month's "Detail of Invoice" sheet.
// Returns { amounts: {altId: signedSum}, duplicates: [{altId, rows, values}], rowCount }.
function sumAmountsByAltId(aoa, headerRowIdx, altIdCol, amountCol) {
  var amounts = {}, occurrences = {};
  for (var r = headerRowIdx + 1; r < aoa.length; r++) {
    var row = aoa[r];
    if (!row) continue;
    var rawId = row[altIdCol];
    if (rawId === undefined || rawId === null || String(rawId).trim() === '') continue;
    var id = normAltId(rawId);
    var amt = row[amountCol];
    amt = typeof amt === 'number' ? amt : 0;
    amounts[id] = (amounts[id] || 0) + amt; // signed sum -- a correction row can be negative
    (occurrences[id] || (occurrences[id] = [])).push({ row: r + 1, value: amt });
  }
  var duplicates = Object.keys(occurrences)
    .filter(function (id) { return occurrences[id].length > 1; })
    .map(function (id) { return { altId: id, rows: occurrences[id] }; });
  return { amounts: amounts, duplicates: duplicates, rowCount: Object.keys(amounts).length };
}

// Detail Reconcile sheet layout (confirmed against the real workbook, header row 5).
var RECONCILE_HEADER_ROW = 5; // 1-based
var RECONCILE_ALT_ID_COL = 'C';

// Month columns in the Invoice section (CR:CW) are identified by their header (row 5) holding an
// actual date -- read from the live file rather than hardcoded, since a new month's column gets
// appended to this range as the workbook is extended over time.
function findInvoiceMonthColumns(aoa) {
  var headerRow = aoa[RECONCILE_HEADER_ROW - 1] || [];
  // "Invoice" section is bounded by CR:CX per the confirmed layout; CX is "Total Invoice" (a
  // formula, must not be touched) and is the first non-date cell after the date-header run.
  var startCol = colLettersToIndex('CR');
  var cols = [];
  for (var c = startCol; c < headerRow.length; c++) {
    var v = headerRow[c];
    if (v instanceof Date || (typeof v === 'number' && v > 40000 && v < 60000)) {
      cols.push(c);
    } else if (cols.length) {
      break; // first non-date cell after the run = Total Invoice (CX) -- stop
    }
  }
  return cols; // 0-based column indices, in order (earliest month first)
}

function excelSerialToMonthKey(serial) {
  var d = new Date(Math.round((serial - 25569) * 86400 * 1000));
  return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0');
}
