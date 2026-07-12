// ══════════════════════════════════════════════════════
//  Read-only lookups built from the DTW template's own sheets (คู่มือ/OUTSOURCE). These sheets
//  are NEVER written to -- their zip entries pass through byte-identical -- we only read their
//  resolved cell values (via excel.js's readSheetGrid, which already resolves cached formula
//  results, shared strings, etc.) to build in-memory maps consumed by glinvoice.js/dtwfill.js.
//
//  Real-file facts confirmed against "Fill in data for DTW.xlsx" (blank template) this session:
//  - คู่มือ D:F starting ~row34 = the actual Cost->Income mapping (D=Income acct, E=Cost acct
//    [the lookup key], F=description). The A:B "OcrCode/OcrName" columns on the SAME rows are an
//    unrelated staff-allocation-code table -- never confuse the two.
//  - คู่มือ E:F rows ~3-8 = Series code/name (only 1304/26-OUT [normal VAT] and 1305/26-MEMO
//    [no-VAT] have ever been seen used by this codebase).
//  - คู่มือ E:F rows ~13-22 = Tax-Code table (OS07="Sales Vat Undue 7%" default for VAT items,
//    ON00="Non Vat" default for no-VAT items).
//  - คู่มือ A:B = GroupNum <-> Payment-Group-Name (e.g. 21 = "100 Days").
//  - OUTSOURCE J ("Payment due") is a formula's CACHED value -- a credit-term NAME like
//    "100 Days", not a number -- must reverse-lookup through คู่มือ A:B for Head!L
//    (PaymentGroupCode), and separately be parsed into an actual day-count for DocDueDate.
// ══════════════════════════════════════════════════════

// ── Cost -> Income mapping (คู่มือ, columns D/E/F) ─────────────────────────────────────────────

function buildCostIncomeMap(aoa) {
  var headerRow = findHeaderRow(aoa, ['Income', 'Cost', 'GL ACCOUNT OUTSOURCE'], 200);
  var incomeCol = findColByHeaderText(aoa, headerRow, 'Income');
  var costCol = findColByHeaderText(aoa, headerRow, 'Cost');
  var nameCol = findColByHeaderText(aoa, headerRow, 'GL ACCOUNT OUTSOURCE');
  var map = new Map();
  var r = headerRow; // 0-based index of the row AFTER the header (header is 1-based headerRow -> aoa[headerRow-1])
  while (r < aoa.length) {
    var row = aoa[r] || [];
    var costVal = normText(row[costCol]);
    var incomeVal = normText(row[incomeCol]);
    if (!costVal && !incomeVal) break; // table ends where both key columns go empty
    if (costVal) map.set(costVal, { income: incomeVal, name: normText(row[nameCol]) });
    r++;
  }
  if (!map.size) throw new Error('อ่านตาราง Cost→Income จากชีต "คู่มือ" ไม่พบข้อมูลเลย กรุณาตรวจสอบไฟล์ template');
  return map;
}

// ── Series lookup table (informational only -- Series itself is rule-based, see resolveSeries) ─

function buildSeriesTable(aoa) {
  var headerRow = findHeaderRow(aoa, ['Series', 'SeriesName'], 40);
  var codeCol = findColByHeaderText(aoa, headerRow, 'Series');
  var nameCol = findColByHeaderText(aoa, headerRow, 'SeriesName');
  var list = [];
  var r = headerRow;
  while (r < aoa.length) {
    var row = aoa[r] || [];
    var code = normText(row[codeCol]);
    if (!code) break;
    list.push({ code: code, name: normText(row[nameCol]) });
    r++;
  }
  return list;
}

// Series is rule-based per the skill's own documented convention: VAT invoices use 1304
// (26-OUT), no-VAT invoices use 1305 (26-MEMO). This codebase has never seen the other 4 codes
// (1302/1303/1306/1307) actually used -- if a caller needs a different Series, that must be an
// explicit user decision surfaced in the UI, never guessed here.
var SERIES_VAT = '1304';
var SERIES_NO_VAT = '1305';

// ── Tax-code table (VatGroup) ───────────────────────────────────────────────────────────────

function buildVatCodeTable(aoa) {
  var headerRow = findHeaderRow(aoa, ['Code', 'Name'], 40);
  var codeCol = findColByHeaderText(aoa, headerRow, 'Code');
  var nameCol = findColByHeaderText(aoa, headerRow, 'Name');
  var list = [];
  var r = headerRow;
  while (r < aoa.length) {
    var row = aoa[r] || [];
    var code = normText(row[codeCol]);
    if (!code) break;
    list.push({ code: code, name: normText(row[nameCol]) });
    r++;
  }
  return list;
}

// Defaults per the skill's documented convention -- confirmed present in every คู่มือ Tax-Code
// table seen so far (OS07 = "Sales Vat Undue 7%", ON00 = "Non Vat").
var VATGROUP_DEFAULT_VAT = 'OS07';
var VATGROUP_DEFAULT_NO_VAT = 'ON00';

// ── Payment-group table (คู่มือ, columns A/B) ───────────────────────────────────────────────────

function buildPaymentGroupMap(aoa) {
  var headerRow = findHeaderRow(aoa, ['GroupNum', 'Payment Group Name'], 10);
  var numCol = findColByHeaderText(aoa, headerRow, 'GroupNum');
  var nameCol = findColByHeaderText(aoa, headerRow, 'Payment Group Name');
  var byNum = new Map(), byName = new Map();
  var r = headerRow;
  while (r < aoa.length) {
    var row = aoa[r] || [];
    var numVal = row[numCol];
    var nameVal = normText(row[nameCol]);
    if ((numVal === null || numVal === undefined || numVal === '') && !nameVal) break;
    var numKey = normText(numVal);
    if (numKey) byNum.set(numKey, nameVal);
    if (nameVal) byName.set(nameVal.toUpperCase(), numKey);
    r++;
  }
  return { byNum: byNum, byName: byName };
}

// ── Customer / business-partner lookup (OUTSOURCE) ─────────────────────────────────────────────

function buildCustomerLookup(aoa) {
  var headerRow = findHeaderRow(aoa, ['BP Code', 'BP Name', 'Tax ID', 'Address', 'GroupNum', 'Payment due'], 20);
  var codeCol = findColByHeaderText(aoa, headerRow, 'BP Code');
  var nameCol = findColByHeaderText(aoa, headerRow, 'BP Name');
  var taxCol = findColByHeaderText(aoa, headerRow, 'Tax ID');
  var addrCol = findColByHeaderText(aoa, headerRow, 'Address');
  var groupCol = findColByHeaderText(aoa, headerRow, 'GroupNum');
  var dueCol = findColByHeaderText(aoa, headerRow, 'Payment due');
  var byName = new Map();
  for (var r = headerRow; r < aoa.length; r++) {
    var row = aoa[r] || [];
    var name = normText(row[nameCol]);
    if (!name) continue;
    var key = name.toUpperCase();
    if (byName.has(key)) continue; // first match wins; duplicates are surfaced by caller if needed
    byName.set(key, {
      name: name,
      code: normText(row[codeCol]),
      taxId: normText(row[taxCol]),
      address: normText(row[addrCol]),
      groupNum: normText(row[groupCol]),
      paymentDueName: normText(row[dueCol])
    });
  }
  return byName;
}

// Exact match only (trim + case-insensitive) -- on miss, returns candidates whose name contains
// the search text (or vice versa) so the UI can show "did you mean...?" instead of silently
// guessing or hard-failing with no lead.
function lookupCustomer(customerMap, name) {
  var key = normText(name).toUpperCase();
  var hit = customerMap.get(key);
  if (hit) return { found: true, customer: hit };
  var candidates = [];
  customerMap.forEach(function (v, k) {
    if (k.indexOf(key) >= 0 || key.indexOf(k) >= 0) candidates.push(v);
  });
  return { found: false, candidates: candidates };
}

// ── Credit-term day-count parsing ──────────────────────────────────────────────────────────────

// Best-effort parse of "N Days"/"N days" into an integer day count. Anything else ("Last day of
// the same month", "Within 3rd of next month", "Every 20th of Month", etc.) returns null and MUST
// surface as a manual-review warning to the user -- never guessed, per the skill's own explicit
// DocDueDate warning (a prior real incident hardcoded a wrong day count here).
function resolveCreditTermDays(termName) {
  var m = /^(\d+)\s*days?$/i.exec(normText(termName));
  return m ? parseInt(m[1], 10) : null;
}
