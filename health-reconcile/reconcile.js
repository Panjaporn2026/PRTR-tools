// ══════════════════════════════════════════════════════
//  Steps 1-3: locate the header row and the bundled health-insurance total column in a monthly
//  invoice file (both the row and the column letter vary by client/month -- confirmed the bundled
//  column drifts GY in Sep'25 -> HE in Jan'26 within just KOHLER_PC's own files, a 149-column
//  shift over 4 months, and different client projects use a different-height title/cutoff band
//  above the header row entirely -- so both must be found by header-text search every time a
//  file loads, never hardcoded), sum duplicate Alternate ID rows (signed), and match against the
//  reconcile sheet.
// ══════════════════════════════════════════════════════

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

// Known real-world phrasings for the bundled health-insurance column header, confirmed against
// actual client files -- most specific first. KOHLER_PC uses the first phrase; KARCHER uses the
// second, plainer phrase (which is a strict substring of the first -- checking the longer phrase
// first lets a KOHLER-style file resolve via that single, most-specific tier rather than the
// looser one, which would also match itself but adds no value there).
var BUNDLED_HEALTH_PHRASES = [
  'HEALTH INSURANCE EMPLOYER TOTAL',
  'HEALTH INSURANCE EMPLOYER',
];

// Tiered search for the bundled health-insurance column's header. Tries the known exact phrases
// first (most specific -> least specific, matching real KOHLER_PC / KARCHER files exactly), then
// falls back to a looser, order-independent word search for any other client's phrasing that
// doesn't match either known phrase verbatim (e.g. reordered words). Returns the FIRST tier with
// at least one match, so a file matching an earlier (more specific) tier isn't made falsely
// ambiguous by a later, looser tier also matching unrelated columns.
function findBundledHealthCandidatesTiered(aoa, headerRowIdx) {
  for (var i = 0; i < BUNDLED_HEALTH_PHRASES.length; i++) {
    var phraseCandidates = findHeaderCandidates(aoa, headerRowIdx, [BUNDLED_HEALTH_PHRASES[i]]);
    if (phraseCandidates.length) return phraseCandidates;
  }
  var tokenTiers = [
    ['INSURANCE', 'TOTAL', 'EMPLOYER'],
    ['INSURANCE', 'TOTAL'],
    ['INSURANCE'], // loosest -- may also catch unrelated text columns (e.g. "Insurance Plan");
                   // that's fine, the ambiguous-candidate picker below makes the user choose.
  ];
  for (var j = 0; j < tokenTiers.length; j++) {
    var tokenCandidates = findHeaderCandidates(aoa, headerRowIdx, tokenTiers[j]);
    if (tokenCandidates.length) return tokenCandidates;
  }
  return [];
}

// Locate the "Detail of Invoice" sheet's real header row by scanning every row (not a fixed row
// number) for a cell containing "Alternate ID". Different client projects use a different-height
// title/cutoff-date band above the header (KOHLER_PC's real header row happens to be row 25 --
// that number is specific to that one client's template, not a safe assumption once this tool is
// reused across other clients' invoice files).
//
// "Alternate ID" is not guaranteed to appear only once: the real KOHLER_PC file confirmed a
// SECOND, unrelated "Alternate ID" column in a completely different payroll-detail table earlier
// in the same sheet (row 8, a wide salary/OT/deduction export), well before the actual invoice
// summary at row 25. Taking the first match blindly picks the wrong table. Disambiguate using
// what the tool actually needs: the real header row is the one where "Alternate ID" AND a bundled
// health-insurance column appear together -- if exactly one row has both, use it silently; if
// there's only one "Alternate ID" row at all (no ambiguity to resolve), use it even without the
// insurance hint; otherwise (zero matches, or more than one row plausibly qualifies) return -1
// and let the caller surface this rather than guess.
//
// Deliberately conservative here -- checks for the known exact phrases (BUNDLED_HEALTH_PHRASES),
// NOT the column-picker's looser word-order-independent/single-keyword fallback tiers. Confirmed
// the real KOHLER_PC row-8 payroll dump has several stray "insurance"-mentioning columns of its
// own ("Insurance Plan", "Car Insurance (Fixed)", "Fidelity Guarantee Insurance", "Health
// Insurance By EMP Deduct") -- none contain either exact phrase, so they correctly don't qualify
// here, but a looser tier (e.g. "INSURANCE" alone) made BOTH row 8 and row 25 qualify, breaking
// this row-disambiguation itself. A wide payroll table can easily mention "insurance" somewhere
// among 100+ columns without being the real invoice-summary table; the looser tiers are only safe
// once we already know which row we're picking a column from (see findBundledHealthColumn), not
// for deciding which row that is.
function findInvoiceHeaderRow(aoa) {
  var altIdRows = [];
  for (var r = 0; r < aoa.length; r++) {
    if (findHeaderCandidates(aoa, r, ['ALTERNATE', 'ID']).length) altIdRows.push(r);
  }
  if (altIdRows.length === 1) return altIdRows[0];
  var withBundledHint = altIdRows.filter(function (r) {
    return BUNDLED_HEALTH_PHRASES.some(function (phrase) {
      return findHeaderCandidates(aoa, r, [phrase]).length > 0;
    });
  });
  return withBundledHint.length === 1 ? withBundledHint[0] : -1;
}

function findAlternateIdColumn(aoa, headerRowIdx) {
  var candidates = findHeaderCandidates(aoa, headerRowIdx, ['ALTERNATE', 'ID']);
  return candidates.length ? candidates[0].col : -1;
}

// Locate the bundled health-insurance column (see findBundledHealthCandidatesTiered for the
// tiered phrasing it matches). Returns:
//  { status: 'ok', col, header, sum, sampleValues }
//  { status: 'ambiguous'|'not_found', candidates: [...] } -- caller MUST surface this for a
//  manual pick rather than silently guessing.
function findBundledHealthColumn(aoa, headerRowIdx, altIdCol) {
  var candidates = findBundledHealthCandidatesTiered(aoa, headerRowIdx);
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
