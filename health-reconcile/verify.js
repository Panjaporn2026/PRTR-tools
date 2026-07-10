// ══════════════════════════════════════════════════════
//  Verification -- there is no LibreOffice/Excel engine available in-browser, so this replaces
//  the skill's "recalculate headless and scan for errors" step with:
//   1. calcChain.xml removal (it goes stale the instant rows are added/changed -- this is what
//      caused a real "Excel found a problem with this file" repair prompt in the sibling
//      accrued-income tool until this exact fix was applied there) + fullCalcOnLoad so Excel
//      recalculates everything the moment the file is opened.
//   2. The one deterministic check that IS possible without an Excel engine: re-parse the
//      generated output and assert the cells we intentionally wrote hold what we intended --
//      this catches wrong-row/wrong-column bugs mechanically, without needing to evaluate
//      XLOOKUP.
// ══════════════════════════════════════════════════════

// Strip xl/calcChain.xml and its two references, and force fullCalcOnLoad in workbook.xml.
// `zipFiles` is an array of {name, data(Uint8Array)} already built for every other part;
// this returns a new array with those 3 parts patched (or dropped, for calcChain.xml).
async function applyRecalcFixes(zipFiles, enc) {
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
      var wbXml = new TextDecoder().decode(f.data);
      if (/<calcPr\b[^>]*\/>/.test(wbXml)) {
        wbXml = wbXml.replace(/<calcPr\b([^>]*)\/>/, function (_, attrs) {
          return /fullCalcOnLoad=/.test(attrs)
            ? '<calcPr' + attrs.replace(/fullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"') + '/>'
            : '<calcPr' + attrs + ' fullCalcOnLoad="1"/>';
        });
      } else {
        wbXml = wbXml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
      }
      out.push({ name: f.name, data: enc.encode(wbXml) });
    } else {
      out.push(f);
    }
  }
  return out;
}

// Re-parse the just-built sheet XML and confirm every intentional write landed where expected.
// `writes` is [{ addr: 'CR41', expected: number|string }], `sheetXml` is the final sheet XML
// after all edits. Returns { ok: bool, mismatches: [...] }.
function verifyWrites(sheetXml, writes) {
  var mismatches = [];
  writes.forEach(function (w) {
    var m = new RegExp('<c\\s+r="' + w.addr + '"[^>]*>([\\s\\S]*?)<\\/c>|<c\\s+r="' + w.addr + '"[^>]*\\/>')
      .exec(sheetXml);
    if (!m) { mismatches.push({ addr: w.addr, expected: w.expected, found: null }); return; }
    var inner = m[0];
    var vMatch = /<v>([\s\S]*?)<\/v>/.exec(inner);
    var found = vMatch ? vMatch[1] : null;
    var foundNum = found !== null ? parseFloat(found) : null;
    var expectedNum = typeof w.expected === 'number' ? w.expected : parseFloat(w.expected);
    var matches = found !== null && !isNaN(foundNum) && !isNaN(expectedNum)
      ? Math.abs(foundNum - expectedNum) < 0.005
      : found === String(w.expected);
    if (!matches) mismatches.push({ addr: w.addr, expected: w.expected, found: found });
  });
  return { ok: mismatches.length === 0, mismatches: mismatches };
}
