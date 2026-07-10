// ══════════════════════════════════════════════════════
//  Step 4: adding new employees -- detect candidates, never auto-add. Everything here feeds a
//  review UI where the user explicitly confirms each row; nothing in this file writes anything.
// ══════════════════════════════════════════════════════

// Alternate IDs billed (nonzero) in any in-scope invoice month but absent from the reconcile
// sheet. `monthlyAmounts` is [{ monthLabel, amounts: {altId: signedSum} }, ...] (one entry per
// in-scope month, from reconcile.js's sumAmountsByAltId), `existingIds` is the Set of Alternate
// IDs already in the Detail Reconcile sheet.
function findNewEmployeeCandidates(monthlyAmounts, existingIds, invoiceIdentityByMonth) {
  var candidates = {}; // altId -> { altId, months: [{monthLabel, amount}], name }
  monthlyAmounts.forEach(function (m) {
    Object.keys(m.amounts).forEach(function (altId) {
      var amt = m.amounts[altId];
      if (!amt) return; // only nonzero bundled charges count as "billed"
      if (existingIds.has(altId)) return;
      if (!candidates[altId]) candidates[altId] = { altId: altId, months: [] };
      candidates[altId].months.push({ monthLabel: m.monthLabel, amount: amt });
    });
  });
  Object.keys(candidates).forEach(function (altId) {
    var identity = (invoiceIdentityByMonth || []).slice().reverse()
      .map(function (m) { return m.identityByAltId && m.identityByAltId[altId]; })
      .find(function (x) { return x; });
    candidates[altId].identity = identity || null;
    candidates[altId].nearMiss = null; // filled in by flagNearMissIds
  });
  return Object.values(candidates);
}

// Guard against Excel's leading-zero coercion / rehires: for each brand-new candidate ID, check
// whether a very-similar ID already exists in the reconcile sheet (same digits, different
// leading zeros / same core with a different prefix) and surface it instead of silently treating
// the candidate as certainly new.
function flagNearMissIds(candidates, existingIds) {
  var existingArr = Array.from(existingIds);
  candidates.forEach(function (cand) {
    var bare = cand.altId.replace(/^0+/, '');
    var match = existingArr.find(function (id) {
      return id !== cand.altId && id.replace(/^0+/, '') === bare && bare !== '';
    });
    cand.nearMiss = match || null;
  });
  return candidates;
}

// Confirm the chosen template row's Actual-Cost-by-CLIENT block (BT:BZ, which CX/CY's Total
// Invoice/Differ derive from) isn't already broken -- a broken template here would propagate the
// break silently into every new row copied from it. Deliberately scoped to just this block: the
// real workbook has several legacy per-employee lookup columns (e.g. U:Y "Old price" and BH:BL
// "Change Plan", both keyed on historical date ranges) that read #N/A for the vast majority of
// CURRENT rows by design -- confirmed against all 191 real rows, where BT:BZ/CX/CY are 100% clean
// but those legacy columns are #N/A on ~97% of rows. Treating those as blocking would make almost
// every real row ineligible as a template, defeating the point of the check.
var TEMPLATE_CRITICAL_COLS = ['BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ'];
function templateRowErrors(templateCells) {
  return TEMPLATE_CRITICAL_COLS
    .filter(function (col) { return templateCells[col] && templateCells[col].isError; })
    .map(function (col) { return { col: col, value: templateCells[col].value }; });
}

// Confirm the template row's plan-code values actually exist as lookup keys in the reference
// sheets the Actual-Cost formulas XLOOKUP against, so a new row doesn't inherit a formula that
// will show #N/A the moment it's opened.
function validatePlanCodes(templateRowValues, tl2025Keys, pru2024Keys) {
  var pruCode = templateRowValues.F; // 'PRU2024'!A:A lookup key
  var tlCode = templateRowValues.H; // 'TL2025'!A:A lookup key
  var problems = [];
  if (pruCode != null && pruCode !== '' && pru2024Keys.indexOf(String(pruCode)) < 0) {
    problems.push({ col: 'F', code: pruCode, sheet: 'PRU2024' });
  }
  if (tlCode != null && tlCode !== '' && tl2025Keys.indexOf(String(tlCode)) < 0) {
    problems.push({ col: 'H', code: tlCode, sheet: 'TL2025' });
  }
  return problems; // empty = valid
}

// Build the override map (formula.js's `overrides` argument) for one confirmed new employee row.
// Everything not listed here inherits the template row's style + re-anchored formula.
function buildNewRowOverrides(identity, planCodeCols, invoiceMonthCols, amountsByCol) {
  var overrides = {
    B: { literal: identity.company || '' },
    C: { literal: identity.altId || '' },
    D: { literal: identity.empId || '' },
    E: { literal: identity.thaiName || '' },
    I: { literal: identity.position || '' }
    // K/L/M (effective/last-working dates) intentionally left to the caller if present in the
    // invoice identity snapshot -- not every new hire has a last-working-day yet.
  };
  if (identity.lastWorkingDay) overrides.M = { literal: identity.lastWorkingDay };

  // Plan codes: literal copy from the template (company-wide constant), not per-employee.
  Object.keys(planCodeCols).forEach(function (col) { overrides[col] = { literal: planCodeCols[col] }; });

  // Active-days columns forced to 0 -- matches the sheet's convention for unverified rows; the
  // copied Actual-Cost formulas will naturally evaluate to 0 downstream from this on recalc.
  ['N', 'O', 'P', 'Q', 'R', 'S'].forEach(function (col) { overrides[col] = { literal: 0 }; });

  // Invoice section: the one part of a new row that carries real, computed data.
  invoiceMonthCols.forEach(function (col) {
    overrides[col] = { literal: amountsByCol[col] || 0 };
  });

  return overrides;
}
