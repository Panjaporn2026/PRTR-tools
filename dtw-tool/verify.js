// ══════════════════════════════════════════════════════
//  Ported from health-reconcile/verify.js. There is no headless Excel recalculation engine
//  available in-browser, so this substitutes for it: re-parse the FINAL sheet XML (after all
//  edits, before zipping) and assert every intended write actually landed. Hard-stop (no
//  download) on any mismatch -- this is the cheapest correctness net available and has already
//  caught real bugs in sibling tools this session.
// ══════════════════════════════════════════════════════

function verifyWrites(sheetXml, writes) {
  var mismatches = [];
  writes.forEach(function (w) {
    var re = new RegExp('<c r="' + w.addr + '"[^>]*(?:/>|>([\\s\\S]*?)</c>)');
    var m = re.exec(sheetXml);
    if (!m) { mismatches.push({ addr: w.addr, expected: w.expected, actual: '(cell not found)' }); return; }
    var inner = m[1] || '';
    var vM = /<t[^>]*>([\s\S]*?)<\/t>|<v>([\s\S]*?)<\/v>/.exec(inner);
    var actualRaw = vM ? (vM[1] !== undefined ? vM[1] : vM[2]) : '';
    var actual = decodeXmlEntities(actualRaw);
    var expNum = typeof w.expected === 'number' ? w.expected : parseFloat(w.expected);
    var actNum = parseFloat(actual);
    var ok;
    if (!isNaN(expNum) && !isNaN(actNum) && typeof w.expected === 'number') {
      ok = Math.abs(expNum - actNum) < 0.005;
    } else {
      ok = String(actual) === String(w.expected);
    }
    if (!ok) mismatches.push({ addr: w.addr, expected: w.expected, actual: actual });
  });
  return { ok: mismatches.length === 0, mismatches: mismatches };
}

// Strips xl/calcChain.xml (+ its [Content_Types].xml override and workbook.xml.rels relationship)
// and forces fullCalcOnLoad="1" on <calcPr> -- defensive even though Head/Line carry no formulas
// today, matching the same posture already used in gl-invoice, in case a future template gains
// one. `zipFiles`: [{name, data:Uint8Array}]. Returns a new array.
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
    } else { out.push(f); }
  }
  return out;
}
