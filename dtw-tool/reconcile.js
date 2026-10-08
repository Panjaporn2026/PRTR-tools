// ══════════════════════════════════════════════════════
//  Cross-checks computed invoice totals against the Detail-of-Invoice ground-truth file (skill
//  §4) and surfaces anyone present in GL_Invoice but not confirmed-ready in the ground truth
//  (skill §3 Pending filter). Read-only, never auto-adjusts a mismatch or auto-excludes anyone --
//  every result here is meant to be reviewed by the user in the preview table, per this
//  project's recurring "always ask, never guess" convention (confirmed necessary again this
//  session: the SM566 real-data case, where an automated E-side total legitimately overstated a
//  real invoice because of a matched retro-correction pair a human had to specially handle).
// ══════════════════════════════════════════════════════

var DETAIL_TOTAL_LABEL_CANDIDATES = ['Total Invoice Before Vat', 'Total invoice before vat', 'Total Invoice Before VAT'];

// Finds the Alternate-ID column and (if present) a per-person total column in a Detail-of-Invoice
// aoa. Detail-of-Invoice files sometimes have a decoy legend table above the real data header
// (confirmed in an earlier session), so this searches by header text rather than a fixed row, and
// picks the header row that's immediately followed by a row that actually looks like data.
// Real Detail of Invoice templates carry a legend table near the top (row 7-8) that also has an
// "Alternate ID" title, above the real data table (header at row 25-28 in the DKSH, Toshiba and
// BBGI files tested). The shared findHeaderRow rejects more than one match, which made every such
// file fail here; for this file only, take the LAST matching row that is followed by an actual
// Alternate ID value (the legend sits above the data, never below it).
function findDetailHeaderRow(aoa) {
  var target = normTextUpper('Alternate ID'), candidates = [];
  for (var r = 0; r < Math.min(60, aoa.length); r++) {
    var row = aoa[r] || [];
    for (var c = 0; c < row.length; c++) if (normTextUpper(row[c]) === target) { candidates.push({ rowNum: r + 1, col: c }); break; }
  }
  if (!candidates.length) throw new Error('ไม่พบแถว header ที่มีคอลัมน์ Alternate ID');
  for (var i = candidates.length - 1; i >= 0; i--) {
    var next = aoa[candidates[i].rowNum] || [];
    if (normText(next[candidates[i].col])) return candidates[i].rowNum;
  }
  return candidates[candidates.length - 1].rowNum;
}

function parseDetailOfInvoice(aoa) {
  var headerRow = findDetailHeaderRow(aoa);
  var altCol = findColByHeaderText(aoa, headerRow, 'Alternate ID');
  var totalCol = null;
  for (var i = 0; i < DETAIL_TOTAL_LABEL_CANDIDATES.length && totalCol == null; i++) {
    try { totalCol = findColByHeaderText(aoa, headerRow, DETAIL_TOTAL_LABEL_CANDIDATES[i]); } catch (e) { /* try next candidate */ }
  }
  var byAltId = new Map();
  for (var r = headerRow; r < aoa.length; r++) {
    var row = aoa[r] || [];
    var altId = normText(row[altCol]);
    if (!altId) continue;
    addDetailTotal(byAltId, altId, totalCol != null ? row[totalCol] : null, totalCol != null);
  }
  return { byAltId: byAltId, hasTotalColumn: totalCol != null, aoa: aoa, headerRow: headerRow };
}

// The same Alternate ID can appear on more than one row -- within one file, or across several
// Detail of Invoice files uploaded together (e.g. files split per cost center). Per the user's
// decision, every row of that person counts: totals are SUMMED (previously the last row won).
function addDetailTotal(byAltId, altId, total, hasTotal) {
  var cur = byAltId.get(altId);
  if (!cur) { byAltId.set(altId, { total: hasTotal ? total : null, rows: 1 }); return; }
  cur.rows++;
  if (!hasTotal) return;
  var a = Number(cur.total), b = Number(total);
  if (isNaN(a) && isNaN(b)) return;          // nothing numeric to add up -- leave as is
  cur.total = (isNaN(a) ? 0 : a) + (isNaN(b) ? 0 : b);
}

// Combines several parsed Detail of Invoice files into one ground truth. Totals are only
// comparable if EVERY file has the per-person total column; otherwise the reconcile falls back to
// the Pending (existence) check only, exactly as for a single file without that column.
function mergeDetailOfInvoices(parsedList) {
  var byAltId = new Map();
  parsedList.forEach(function (p) {
    p.byAltId.forEach(function (v, altId) {
      addDetailTotal(byAltId, altId, v.total, p.hasTotalColumn); // v.total is already this file's sum (counts 1 row)
      byAltId.get(altId).rows += (v.rows || 1) - 1;              // plus that file's other rows of this person
    });
  });
  return { byAltId: byAltId, hasTotalColumn: parsedList.length > 0 && parsedList.every(function (p) { return p.hasTotalColumn; }), sources: parsedList };
}

// Compares each invoice's own computed total (before VAT) against the ground truth's per-person
// total, keyed by Alternate ID -- NOT by whatever grouping key the user picked, since the ground
// truth is inherently per-employee. Only meaningful when the grouping key IS Alternate ID or when
// an Alternate-ID-equivalent value is available per group; otherwise this degrades to a Pending
// existence check only (still useful) without a numeric diff.
function reconcileInvoices(invoices, detailOfInvoice, altIdColUsedAsKey) {
  return invoices.map(function (inv) {
    var found = detailOfInvoice.byAltId.has(inv.groupKey);
    var result = { invoice: inv, foundInGroundTruth: found, diff: null };
    if (!found) return result;
    if (!altIdColUsedAsKey || !detailOfInvoice.hasTotalColumn) return result;
    var groundTotal = Number(detailOfInvoice.byAltId.get(inv.groupKey).total);
    if (isNaN(groundTotal)) return result;
    result.groundTotal = groundTotal;
    result.diff = round2(inv.totalBeforeVat - groundTotal);
    return result;
  });
}

// Anyone whose grouping-key value doesn't appear in the ground truth at all is surfaced as
// "possibly Pending" -- never silently dropped from the invoice list. The caller renders this as
// a reviewable warning list; excluding them is an explicit user action, not automatic.
// Besides Alternate ID, the key is also looked up in the Detail-of-Invoice column with the same
// header as the chosen grouping-key column (e.g. "Cost Center 4"), across every uploaded Detail
// file -- a group keyed by a line manager's name (real MICHELIN case: K.Chanpen Chaisutartip)
// never matches any Alternate ID even though its employees are in the Detail file, which
// produced a false Pending warning.
function detailValuesOfColumn(detailOfInvoice, colName) {
  var values = new Set();
  if (!colName) return values;
  (detailOfInvoice.sources || [detailOfInvoice]).forEach(function (src) {
    if (!src.aoa) return;
    var col;
    try { col = findColByHeaderText(src.aoa, src.headerRow, colName); } catch (e) { return; }
    for (var r = src.headerRow; r < src.aoa.length; r++) {
      var v = normText((src.aoa[r] || [])[col]);
      if (v) values.add(v);
    }
  });
  return values;
}
function findPossiblyPending(invoices, detailOfInvoice, groupingKeyCol) {
  var keyValues = detailValuesOfColumn(detailOfInvoice, groupingKeyCol);
  return invoices.filter(function (inv) { return !detailOfInvoice.byAltId.has(inv.groupKey) && !keyValues.has(inv.groupKey); });
}
