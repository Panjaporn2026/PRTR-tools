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
function parseDetailOfInvoice(aoa) {
  var headerRow = findHeaderRow(aoa, ['Alternate ID'], 60);
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
    byAltId.set(altId, { total: totalCol != null ? row[totalCol] : null });
  }
  return { byAltId: byAltId, hasTotalColumn: totalCol != null };
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
function findPossiblyPending(invoices, detailOfInvoice) {
  return invoices.filter(function (inv) { return !detailOfInvoice.byAltId.has(inv.groupKey); });
}
