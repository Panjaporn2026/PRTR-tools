// ══════════════════════════════════════════════════════
//  Style-append helper -- new territory for this codebase (no sibling tool has needed to
//  synthesize new cellXfs/font entries before; every prior tool only ever *carried forward* an
//  existing cell's own s="N" attribute). Used to build "same style but red font" variants for
//  EXPENSE rows (functions 3/4/5). Append-only: inserting a new font/cellXfs entry in the MIDDLE
//  of either array would silently shift every OTHER cell's existing style index across the whole
//  workbook, corrupting formatting that has nothing to do with this tool's own writes.
// ══════════════════════════════════════════════════════

function parseStylesXml(stylesXml) {
  var fontsM = /<fonts\b([^>]*)>([\s\S]*?)<\/fonts>/.exec(stylesXml);
  var fonts = [];
  if (fontsM) {
    var fontRe = /<font>([\s\S]*?)<\/font>/g, fm;
    while ((fm = fontRe.exec(fontsM[2])) !== null) fonts.push(fm[1]);
  }
  var cellXfsM = /<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/.exec(stylesXml);
  var cellXfs = [];
  if (cellXfsM) {
    var xfRe = /<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g, xm;
    while ((xm = xfRe.exec(cellXfsM[2])) !== null) cellXfs.push(xm[0]);
  }
  return {
    stylesXml: stylesXml,
    fonts: fonts,
    cellXfs: cellXfs,
    fontsAttrs: fontsM ? fontsM[1] : '',
    cellXfsAttrs: cellXfsM ? cellXfsM[1] : ''
  };
}

function parseXfAttrs(xfXml) {
  var attrs = {};
  ['numFmtId', 'fontId', 'fillId', 'borderId', 'xfId'].forEach(function (name) {
    var m = new RegExp('\\b' + name + '="(\\d+)"').exec(xfXml);
    attrs[name] = m ? parseInt(m[1], 10) : 0;
  });
  return attrs;
}

function xfSemanticEquals(xfA, xfB) {
  var a = parseXfAttrs(xfA), b = parseXfAttrs(xfB);
  var alignA = /<alignment\b[^>]*\/>/.exec(xfA);
  var alignB = /<alignment\b[^>]*\/>/.exec(xfB);
  return a.numFmtId === b.numFmtId && a.fontId === b.fontId && a.fillId === b.fillId &&
    a.borderId === b.borderId && a.xfId === b.xfId &&
    (alignA ? alignA[0] : null) === (alignB ? alignB[0] : null);
}

// Returns a copy of xfXml with its fontId attribute set to newFontId (adding the attribute, and
// applyFont="1", if the original xf didn't specify a font at all).
function withFontId(xfXml, newFontId) {
  var replaced = /\bfontId="\d+"/.test(xfXml)
    ? xfXml.replace(/\bfontId="\d+"/, 'fontId="' + newFontId + '"')
    : xfXml.replace('<xf ', '<xf fontId="' + newFontId + '" ');
  if (!/\bapplyFont="1"/.test(replaced)) {
    replaced = /\/>\s*$/.test(replaced)
      ? replaced.replace(/\/>\s*$/, ' applyFont="1"/>')
      : replaced.replace(/^(<xf\b[^>]*)>/, '$1 applyFont="1">');
  }
  return replaced;
}

var RED_RGB = 'FFFF0000';

function isRedFont(fontXml) {
  return new RegExp('<color[^/]*rgb="' + RED_RGB + '"').test(fontXml);
}
function isBoldFont(fontXml) {
  return /<b\s*\/>/.test(fontXml);
}
// Compares two fonts ignoring bold and color -- i.e. "is this the same base font family/size".
function sameBaseFont(fontA, fontB) {
  function stripped(f) { return f.replace(/<b\s*\/>/g, '').replace(/<color[^/]*\/>/g, '').replace(/\s+/g, ''); }
  return stripped(fontA) === stripped(fontB);
}

function setFontColorRed(fontXml) {
  if (/<color[^/]*\/>/.test(fontXml)) return fontXml.replace(/<color[^/]*\/>/, '<color rgb="' + RED_RGB + '"/>');
  return fontXml + '<color rgb="' + RED_RGB + '"/>';
}

// Finds an EXISTING red font matching the original font's base family/size and bold-ness exactly
// -- e.g. this template's own fontId=3 (plain red) / fontId=4 (bold red). Returns null (never
// guesses) if no byte-for-byte match exists, so the caller clones the original font instead of
// reusing an unrelated red font that happens to exist in the workbook.
function findMatchingRedFont(styles, origFont, isBold) {
  for (var i = 0; i < styles.fonts.length; i++) {
    var f = styles.fonts[i];
    if (isRedFont(f) && isBoldFont(f) === isBold && sameBaseFont(f, origFont)) return i;
  }
  return null;
}

// Finds (or appends, append-only) a cellXfs index that renders EXACTLY like originalXfIndex but
// with a red font -- same numFmt/fill/border/alignment, font color forced red, bold-ness
// preserved from the original. A missing s="" attribute means style index 0 (Excel's default),
// handled explicitly by the caller passing 0, not null/undefined.
function getOrCreateRedVariant(styles, originalXfIndex) {
  var originalXf = styles.cellXfs[originalXfIndex] || styles.cellXfs[0] || '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
  var origAttrs = parseXfAttrs(originalXf);
  var origFont = styles.fonts[origAttrs.fontId] != null ? styles.fonts[origAttrs.fontId] : (styles.fonts[0] || '<sz val="11"/><color theme="1"/><name val="Calibri"/>');
  var isBold = isBoldFont(origFont);

  var redFontId = findMatchingRedFont(styles, origFont, isBold);
  if (redFontId === null) {
    styles.fonts.push(setFontColorRed(origFont));
    redFontId = styles.fonts.length - 1;
  }

  var candidateXf = withFontId(originalXf, redFontId);
  var existingIdx = -1;
  for (var i = 0; i < styles.cellXfs.length; i++) {
    if (xfSemanticEquals(styles.cellXfs[i], candidateXf)) { existingIdx = i; break; }
  }
  if (existingIdx >= 0) return existingIdx;

  styles.cellXfs.push(candidateXf);
  return styles.cellXfs.length - 1;
}

function serializeStylesXml(styles) {
  var fontsAttrs = /\bcount="\d+"/.test(styles.fontsAttrs)
    ? styles.fontsAttrs.replace(/\bcount="\d+"/, 'count="' + styles.fonts.length + '"')
    : styles.fontsAttrs + ' count="' + styles.fonts.length + '"';
  var cellXfsAttrs = /\bcount="\d+"/.test(styles.cellXfsAttrs)
    ? styles.cellXfsAttrs.replace(/\bcount="\d+"/, 'count="' + styles.cellXfs.length + '"')
    : styles.cellXfsAttrs + ' count="' + styles.cellXfs.length + '"';

  var fontsXml = '<fonts' + fontsAttrs + '>' + styles.fonts.map(function (f) { return '<font>' + f + '</font>'; }).join('') + '</fonts>';
  var cellXfsXml = '<cellXfs' + cellXfsAttrs + '>' + styles.cellXfs.join('') + '</cellXfs>';

  var newXml = styles.stylesXml.replace(/<fonts\b[^>]*>[\s\S]*?<\/fonts>/, fontsXml);
  newXml = newXml.replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/, cellXfsXml);
  return newXml;
}
