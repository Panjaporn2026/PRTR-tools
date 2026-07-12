// ══════════════════════════════════════════════════════
//  GL_Invoice ingestion: merge multiple files, filter to Expense-side rows, group into invoices,
//  split VAT vs no-VAT per the user's own selections. Read-only -- GL_Invoice files are never
//  rewritten, only read (via SheetJS in app.js, same technique already proven in
//  doi-tools/rename-lock's readDetailOfInvoiceMeta for messy real-world exports).
//
//  Real-file facts confirmed this session (Siemens GL_Invoice, post-gl-invoice-tool shape):
//  43 columns, header row position not fixed (this sample had it at row 1, but the skill's own
//  file family sometimes has metadata rows before it) -- always found by text search via
//  findHeaderRow/findColByHeaderText from sheetmodel.js, never a hardcoded row/column.
//  `Grouping` values are prefixed D (debit/clearing -- net pay, statutory deductions, NEVER
//  invoiced) or E (expense/cost -- the ONLY rows that get invoiced, skill step 3).
//
//  Per explicit user decision this session: candidate grouping-key/Reference-No. columns are
//  drawn ONLY from columns that actually exist in the uploaded GL_Invoice file(s) -- no join
//  against the Detail-of-Invoice file for a PO column that GL_Invoice itself doesn't carry.
// ══════════════════════════════════════════════════════

var GL_REQUIRED_COLUMNS = ['Alternate ID', 'NAME', 'Project Code SAP', 'Line Manager', 'Account', 'Grouping', 'Amount'];

// Columns worth offering as invoice-grouping-key / Reference-No. picker options, if present.
// Deliberately excludes Amount/Account/Grouping/date-like/free-text columns -- these are never
// meaningful things to group or reference an invoice by.
var GL_CANDIDATE_KEY_COLUMNS = [
  'Alternate ID', 'EMP ID', 'NAME', 'Line Manager', 'Invoice Sent To',
  'Cost Center 1', 'Cost Center 2', 'Cost Center 3', 'Cost Center 4', 'Project Code SAP'
];

function normText(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
}
function normTextUpper(s) { return normText(s).toUpperCase(); }

// Reads one already-parsed aoa (from SheetJS: XLSX.utils.sheet_to_json(ws, {header:1})) into a
// {headerRowNum, colIndexByName, dataRows: aoa-rows-below-header} shape. Hard-throws (via
// findHeaderRow/findColByHeaderText) if the required columns aren't all present in one row.
function parseGLInvoiceAoa(aoa, sourceLabel) {
  var headerRowNum;
  try {
    headerRowNum = findHeaderRow(aoa, GL_REQUIRED_COLUMNS, 30);
  } catch (e) {
    throw new Error('ไฟล์ "' + sourceLabel + '": ' + e.message);
  }
  var headerRow = aoa[headerRowNum - 1] || [];
  var colIndexByName = {};
  for (var c = 0; c < headerRow.length; c++) {
    var label = normText(headerRow[c]);
    if (label) colIndexByName[label] = c;
  }
  var dataRows = aoa.slice(headerRowNum); // 0-based rows after the header
  return { headerRowNum: headerRowNum, colIndexByName: colIndexByName, dataRows: dataRows, sourceLabel: sourceLabel };
}

// Merges N parsed GL_Invoice files into one flat row-object array. Every row becomes
// { sourceFile, get(colName) } -- get() returns '' if that file doesn't have the column at all
// (rather than throwing), so per-file column drift doesn't break ingestion as long as the
// REQUIRED columns are present in every file.
function mergeGLInvoiceFiles(parsedFiles) {
  parsedFiles.forEach(function (pf) {
    GL_REQUIRED_COLUMNS.forEach(function (label) {
      if (!(label in pf.colIndexByName)) {
        throw new Error('ไฟล์ "' + pf.sourceLabel + '" ไม่มีคอลัมน์ที่จำเป็น: "' + label + '"');
      }
    });
  });
  var rows = [];
  parsedFiles.forEach(function (pf) {
    pf.dataRows.forEach(function (rawRow) {
      var row = rawRow || [];
      // Skip fully-blank rows (trailing blank rows are common in real exports).
      var hasAny = row.some(function (v) { return normText(v) !== ''; });
      if (!hasAny) return;
      rows.push({
        sourceFile: pf.sourceLabel,
        get: function (colName) {
          var idx = pf.colIndexByName[colName];
          return idx == null ? '' : row[idx];
        }
      });
    });
  });
  return rows;
}

// Column names present in EVERY merged file (intersection), restricted to the curated candidate
// list -- these become the radio-button options for grouping-key / Reference-No. source. Never
// silently pick one; the caller (app.js) must render this as a user choice.
function detectCandidateKeyColumns(parsedFiles) {
  if (!parsedFiles.length) return [];
  var sets = parsedFiles.map(function (pf) { return new Set(Object.keys(pf.colIndexByName)); });
  return GL_CANDIDATE_KEY_COLUMNS.filter(function (label) {
    return sets.every(function (s) { return s.has(label); });
  });
}

// Keep only Expense/Cost-side rows (Grouping prefixed "E") -- Debit/clearing-side rows
// (net pay, statutory deductions, prefixed "D") are never invoiced, per skill step 3.
function filterExpenseRows(rows) {
  return rows.filter(function (row) {
    var g = normText(row.get('Grouping'));
    return g.charAt(0).toUpperCase() === 'E';
  });
}

// Maps every kept row's Cost Account through the Cost->Income table. Rows whose Account isn't in
// the mapping are NOT silently dropped -- they're returned separately as `unmapped` so the caller
// can surface them as a hard warning (an unmapped cost account means either a data error or a
// mapping-table gap, and this codebase's rule is to ask rather than guess).
function mapCostToIncome(rows, costIncomeMap) {
  var mapped = [], unmapped = [];
  rows.forEach(function (row) {
    var acct = normText(row.get('Account'));
    var m = costIncomeMap.get(acct);
    if (!m) { unmapped.push(row); return; }
    mapped.push({ row: row, incomeAccount: m.income, incomeName: m.name });
  });
  return { mapped: mapped, unmapped: unmapped };
}

// Groups mapped rows by the user-chosen grouping-key column, and within each group splits Line
// items into a normal-VAT invoice and (only if present) a separate no-VAT invoice, per the
// noVatIncomeAccounts the user ticked. Sums Amount per Income account within each split -- this
// directly mirrors the real Siemens ground truth (two Head rows for the same PO: one VAT, one
// no-VAT Reimbursement).
function buildInvoiceGroups(mappedRows, groupingKeyCol, refNoCol, noVatIncomeAccounts) {
  var groups = new Map(); // groupKey -> { key, refNo, rows: [] }
  mappedRows.forEach(function (entry) {
    var key = normText(entry.row.get(groupingKeyCol));
    if (!key) return; // rows with no grouping-key value can't be assigned to an invoice
    if (!groups.has(key)) {
      groups.set(key, {
        key: key,
        refNo: normText(entry.row.get(refNoCol)),
        name: normText(entry.row.get('NAME')),
        invoiceSentTo: normText(entry.row.get('Invoice Sent To')),
        projectCodeSap: normText(entry.row.get('Project Code SAP')),
        period: entry.row.get('Calendar Group'),
        lineManager: normText(entry.row.get('Line Manager')),
        rows: []
      });
    }
    groups.get(key).rows.push(entry);
  });

  var invoices = [];
  groups.forEach(function (g) {
    var vatLines = new Map(), noVatLines = new Map();
    g.rows.forEach(function (entry) {
      var target = noVatIncomeAccounts.has(entry.incomeAccount) ? noVatLines : vatLines;
      var cur = target.get(entry.incomeAccount) || { incomeAccount: entry.incomeAccount, incomeName: entry.incomeName, amount: 0 };
      cur.amount += Number(entry.row.get('Amount')) || 0;
      target.set(entry.incomeAccount, cur);
    });
    function toInvoice(linesMap, isNoVat) {
      if (!linesMap.size) return null;
      var lines = Array.from(linesMap.values()).filter(function (l) { return Math.abs(l.amount) > 1e-9; });
      if (!lines.length) return null;
      var total = lines.reduce(function (s, l) { return s + l.amount; }, 0);
      return {
        groupKey: g.key, refNo: g.refNo, name: g.name, invoiceSentTo: g.invoiceSentTo,
        projectCodeSap: g.projectCodeSap, period: g.period, lineManager: g.lineManager,
        isNoVat: isNoVat, lines: lines, totalBeforeVat: total
      };
    }
    var vatInv = toInvoice(vatLines, false);
    var noVatInv = toInvoice(noVatLines, true);
    if (vatInv) invoices.push(vatInv);
    if (noVatInv) invoices.push(noVatInv);
  });
  return invoices;
}
