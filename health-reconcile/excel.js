// ══════════════════════════════════════════════════════
//  ZIP + XLSX raw read/write primitives — ported from accrued-income/index.html.
//  Same xlsx-as-zip format, same techniques (no external library needed for this part).
// ══════════════════════════════════════════════════════

function parseZipEntries(buf) {
  var b = new Uint8Array(buf);
  var entries = {};
  var i = 0;
  while (i < b.length - 4) {
    if (b[i]===0x50 && b[i+1]===0x4B && b[i+2]===0x03 && b[i+3]===0x04) {
      var method = b[i+8] | (b[i+9]<<8);
      var csize  = b[i+18]|(b[i+19]<<8)|(b[i+20]<<16)|(b[i+21]<<24);
      var usize  = b[i+22]|(b[i+23]<<8)|(b[i+24]<<16)|(b[i+25]<<24);
      var fnlen  = b[i+26]|(b[i+27]<<8);
      var exlen  = b[i+28]|(b[i+29]<<8);
      var dstart = i + 30 + fnlen + exlen;
      var fname  = new TextDecoder().decode(b.slice(i+30, i+30+fnlen));
      entries[fname] = {method, csize, usize, dstart};
      i = dstart + csize;
    } else { i++; }
  }
  return entries;
}

async function decompressEntry(entry, buf) {
  var raw = new Uint8Array(buf).slice(entry.dstart, entry.dstart + entry.csize);
  if (entry.method === 0) return new TextDecoder().decode(raw);
  var ds = new DecompressionStream('deflate-raw');
  var w = ds.writable.getWriter(), r = ds.readable.getReader();
  w.write(raw); w.close();
  var chunks = [], tot = 0;
  while (true) { var {done,value} = await r.read(); if(done) break; chunks.push(value); tot+=value.length; }
  var out = new Uint8Array(tot), off = 0;
  for (var c of chunks) { out.set(c,off); off+=c.length; }
  return new TextDecoder().decode(out);
}

async function decompressEntryBytes(entry, buf) {
  var raw = new Uint8Array(buf).slice(entry.dstart, entry.dstart + entry.csize);
  if (entry.method === 0) return raw;
  var ds = new DecompressionStream('deflate-raw');
  var w = ds.writable.getWriter(), r = ds.readable.getReader();
  w.write(raw); w.close();
  var chunks = [], tot = 0;
  while (true) { var {done,value} = await r.read(); if(done) break; chunks.push(value); tot+=value.length; }
  var out = new Uint8Array(tot), off = 0;
  for (var c of chunks) { out.set(c,off); off+=c.length; }
  return out;
}

function crc32(data) {
  var c = 0xFFFFFFFF;
  for (var i=0; i<data.length; i++) {
    c ^= data[i];
    for (var j=0; j<8; j++) c = (c&1) ? ((c>>>1)^0xEDB88320) : (c>>>1);
  }
  return (c ^ 0xFFFFFFFF) >>> 0;
}
function u32(n){return [n&0xFF,(n>>8)&0xFF,(n>>16)&0xFF,(n>>24)&0xFF];}

function buildZip(files) {
  var enc = new TextEncoder();
  var locals = [], cds = [], offset = 0;
  for (var fi=0; fi<files.length; fi++) {
    var f = files[fi];
    var nb = enc.encode(f.name);
    var crc = crc32(f.data);
    var lh = new Uint8Array([
      0x50,0x4B,0x03,0x04, 0x14,0x00, 0x00,0x00, 0x00,0x00,
      0x00,0x00, 0x00,0x00
    ].concat(u32(crc), u32(f.data.length), u32(f.data.length), [nb.length&0xFF,(nb.length>>8)&0xFF], [0,0]));
    var lhFull = concatBytes([lh, nb]);
    var cd = new Uint8Array([
      0x50,0x4B,0x01,0x02, 0x14,0x00, 0x14,0x00, 0x00,0x00, 0x00,0x00,
      0x00,0x00, 0x00,0x00
    ].concat(u32(crc), u32(f.data.length), u32(f.data.length), [nb.length&0xFF,(nb.length>>8)&0xFF], [0,0,0,0,0,0,0,0,0,0,0,0], u32(offset)));
    var cdFull = concatBytes([cd, nb]);
    locals.push(lhFull, f.data);
    cds.push(cdFull);
    offset += lhFull.length + f.data.length;
  }
  var cdStart = offset, cdSize = 0;
  for (var c of cds) cdSize += c.length;
  var eocd = new Uint8Array([
    0x50,0x4B,0x05,0x06, 0x00,0x00, 0x00,0x00
  ].concat([files.length&0xFF,(files.length>>8)&0xFF], [files.length&0xFF,(files.length>>8)&0xFF], u32(cdSize), u32(cdStart), [0,0]));
  return concatBytes(locals.concat(cds).concat([eocd]));
}
function concatBytes(parts) {
  var tot = 0;
  for (var p of parts) tot += p.length;
  var out = new Uint8Array(tot), off = 0;
  for (var p of parts) { out.set(p,off); off+=p.length; }
  return out;
}

// Read a workbook's zip + resolve sheet name -> worksheet XML path via workbook.xml + rels.
async function loadWorkbook(buf) {
  var entries = parseZipEntries(buf);
  var wbXml = await decompressEntry(entries['xl/workbook.xml'], buf);
  var relsXml = entries['xl/_rels/workbook.xml.rels'] ? await decompressEntry(entries['xl/_rels/workbook.xml.rels'], buf) : '';
  // Attribute order within a tag is not guaranteed by the OOXML spec -- genuine Excel saves
  // happen to write Id before Target (and name before r:id on <sheet>), but files re-saved by
  // other tools (LibreOffice, openpyxl, other accounting systems a different client project might
  // export from) can write them in a different order. Parse each tag fully, then pull attributes
  // out independently of position, instead of assuming a fixed order.
  var relMap = {};
  var relTagRe = /<Relationship\b[^>]*\/>/g, relTagM;
  while ((relTagM = relTagRe.exec(relsXml)) !== null) {
    var idM = /\bId="([^"]+)"/.exec(relTagM[0]);
    var targetM = /\bTarget="([^"]+)"/.exec(relTagM[0]);
    if (idM && targetM) relMap[idM[1]] = targetM[1];
  }
  var sheets = {};
  var sheetTagRe = /<sheet\b[^>]*\/>/g, sheetTagM;
  while ((sheetTagM = sheetTagRe.exec(wbXml)) !== null) {
    var nameM = /\bname="([^"]+)"/.exec(sheetTagM[0]);
    var ridM = /\br:id="(rId\d+)"/.exec(sheetTagM[0]);
    if (!nameM || !ridM) continue;
    var target = relMap[ridM[1]];
    if (!target) continue;
    var path = target.indexOf('worksheets/') >= 0 ? 'xl/' + target.replace(/^\/?xl\//,'') : 'xl/' + target;
    if (path.indexOf('xl/xl/') === 0) path = path.slice(3);
    sheets[nameM[1]] = path;
  }
  return { buf, entries, sheets, wbXml };
}

async function getSheetXml(wb, sheetName) {
  var path = wb.sheets[sheetName];
  if (!path) throw new Error('ไม่พบชีต "' + sheetName + '"');
  return await decompressEntry(wb.entries[path], wb.buf);
}

async function getSharedStrings(wb) {
  if (wb._sst) return wb._sst;
  var sst = [];
  if (wb.entries['xl/sharedStrings.xml']) {
    var xml = await decompressEntry(wb.entries['xl/sharedStrings.xml'], wb.buf);
    var re = /<si>([\s\S]*?)<\/si>/g, m;
    while ((m = re.exec(xml)) !== null) {
      var textRe = /<t[^>]*>([\s\S]*?)<\/t>/g, tm, joined = '';
      while ((tm = textRe.exec(m[1])) !== null) joined += tm[1];
      sst.push(decodeXmlEntities(joined));
    }
  }
  wb._sst = sst;
  return sst;
}

function decodeXmlEntities(s) {
  return String(s || '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function colLettersToIndex(letters) {
  var col = 0;
  for (var i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64);
  return col - 1; // 0-based
}

// Parse a worksheet's raw XML into a 2D array (aoa[rowIndex][colIndex], both 0-based),
// resolving shared strings (t="s") and inline strings (t="inlineStr"); numeric cells become
// JS numbers, everything else stays a string. Also returns the raw per-row inner XML (needed
// by formula.js's shared-formula resolver) keyed by 1-based row number.
async function readSheetGrid(wb, sheetName) {
  var xml = await getSheetXml(wb, sheetName);
  var sst = await getSharedStrings(wb);
  var aoa = [];
  var rowXmlByNum = {};
  var rowRe = /<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g, rm;
  while ((rm = rowRe.exec(xml)) !== null) {
    var rowNum = parseInt(rm[1], 10);
    rowXmlByNum[rowNum] = rm[2];
    var arr = aoa[rowNum - 1] || (aoa[rowNum - 1] = []);
    var cellRe = /<c\s+r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, cm;
    while ((cm = cellRe.exec(rm[2])) !== null) {
      var colIdx = colLettersToIndex(cm[1]);
      var attrs = cm[2] || '', inner = cm[3] || '';
      var tMatch = /\st="([^"]+)"/.exec(attrs);
      var type = tMatch ? tMatch[1] : null;
      var vMatch = /<v>([\s\S]*?)<\/v>/.exec(inner);
      var value = null;
      if (type === 's' && vMatch) {
        value = sst[parseInt(vMatch[1], 10)];
      } else if (type === 'inlineStr') {
        var isMatch = /<is>([\s\S]*?)<\/is>/.exec(inner);
        var tInner = isMatch ? /<t[^>]*>([\s\S]*?)<\/t>/.exec(isMatch[1]) : null;
        value = decodeXmlEntities(tInner ? tInner[1] : '');
      } else if (vMatch) {
        var n = parseFloat(vMatch[1]);
        value = isNaN(n) ? vMatch[1] : n;
      }
      arr[colIdx] = value;
    }
  }
  return { aoa: aoa, rowXmlByNum: rowXmlByNum, sheetXml: xml };
}
