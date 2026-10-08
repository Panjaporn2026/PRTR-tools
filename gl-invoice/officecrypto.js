// ══════════════════════════════════════════════════════════════════════════════════════════
//  officecrypto.js — removes the "password to open" from .xlsx / .xlsm files, in the browser.
//
//  A password-protected Excel file is an OLE container holding the ORIGINAL workbook encrypted
//  (MS-OFFCRYPTO / ECMA-376). Decrypting it gives back that original workbook byte-for-byte --
//  nothing is re-saved, so data, formulas and formatting cannot change.
//
//  Supported:  Agile encryption   (Excel 2010+, LibreOffice, most tools; AES + SHA-1/256/384/512)
//              Standard encryption (Excel 2007; AES-128/192/256 + SHA-1)
//  Not supported: legacy .xls (BIFF RC4) -- reported as such, never guessed at.
//
//  Needs: XLSX (SheetJS, for XLSX.CFB to read the OLE container) and WebCrypto (AES).
//  Hashes are implemented here synchronously: key derivation runs 50,000-100,000 hash rounds per
//  password try, which is far faster in plain JS than as 100,000 awaited WebCrypto calls.
// ══════════════════════════════════════════════════════════════════════════════════════════
(function (global) {
  'use strict';

  // ---------- byte helpers ----------
  function concat() {
    var n = 0, i, a;
    for (i = 0; i < arguments.length; i++) n += arguments[i].length;
    var out = new Uint8Array(n), o = 0;
    for (i = 0; i < arguments.length; i++) { a = arguments[i]; out.set(a, o); o += a.length; }
    return out;
  }
  function u32le(v) { return new Uint8Array([v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255]); }
  function readU32(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0; }
  function utf16le(s) {
    var out = new Uint8Array(s.length * 2);
    for (var i = 0; i < s.length; i++) { var c = s.charCodeAt(i); out[i * 2] = c & 255; out[i * 2 + 1] = c >>> 8; }
    return out;
  }
  function b64(s) {
    var bin = atob(s.replace(/\s+/g, '')), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  function equalBytes(a, b) {
    if (a.length !== b.length) return false;
    for (var i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
    return true;
  }
  function fitKey(h, len, fill) { // MS-OFFCRYPTO 2.3.4.11: truncate, or pad with 0x36
    if (h.length >= len) return h.slice(0, len);
    var out = new Uint8Array(len); out.fill(fill); out.set(h); return out;
  }

  // ---------- SHA-1 ----------
  function sha1(msg) {
    var ml = msg.length, nb = ((ml + 8) >> 6) + 1, w = new Int32Array(nb * 16), i;
    for (i = 0; i < ml; i++) w[i >> 2] |= msg[i] << (24 - (i & 3) * 8);
    w[ml >> 2] |= 0x80 << (24 - (ml & 3) * 8);
    w[nb * 16 - 1] = ml * 8;
    var h0 = 0x67452301, h1 = 0xEFCDAB89, h2 = 0x98BADCFE, h3 = 0x10325476, h4 = 0xC3D2E1F0, W = new Int32Array(80);
    for (var blk = 0; blk < nb * 16; blk += 16) {
      for (i = 0; i < 16; i++) W[i] = w[blk + i];
      for (i = 16; i < 80; i++) { var x = W[i - 3] ^ W[i - 8] ^ W[i - 14] ^ W[i - 16]; W[i] = (x << 1) | (x >>> 31); }
      var a = h0, b = h1, c = h2, d = h3, e = h4;
      for (i = 0; i < 80; i++) {
        var f, k;
        if (i < 20) { f = (b & c) | (~b & d); k = 0x5A827999; }
        else if (i < 40) { f = b ^ c ^ d; k = 0x6ED9EBA1; }
        else if (i < 60) { f = (b & c) | (b & d) | (c & d); k = 0x8F1BBCDC; }
        else { f = b ^ c ^ d; k = 0xCA62C1D6; }
        var t = (((a << 5) | (a >>> 27)) + f + e + k + W[i]) | 0;
        e = d; d = c; c = (b << 30) | (b >>> 2); b = a; a = t;
      }
      h0 = (h0 + a) | 0; h1 = (h1 + b) | 0; h2 = (h2 + c) | 0; h3 = (h3 + d) | 0; h4 = (h4 + e) | 0;
    }
    return wordsToBytes([h0, h1, h2, h3, h4]);
  }
  function wordsToBytes(ws) {
    var out = new Uint8Array(ws.length * 4);
    for (var i = 0; i < ws.length; i++) { out[i * 4] = ws[i] >>> 24; out[i * 4 + 1] = (ws[i] >>> 16) & 255; out[i * 4 + 2] = (ws[i] >>> 8) & 255; out[i * 4 + 3] = ws[i] & 255; }
    return out;
  }

  // ---------- SHA-256 ----------
  var K256 = new Int32Array([0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2]);
  function sha256(msg) {
    var ml = msg.length, nb = ((ml + 8) >> 6) + 1, w = new Int32Array(nb * 16), i;
    for (i = 0; i < ml; i++) w[i >> 2] |= msg[i] << (24 - (i & 3) * 8);
    w[ml >> 2] |= 0x80 << (24 - (ml & 3) * 8);
    w[nb * 16 - 1] = ml * 8;
    var H = [0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19].map(function (x) { return x | 0; });
    var W = new Int32Array(64);
    for (var blk = 0; blk < nb * 16; blk += 16) {
      for (i = 0; i < 16; i++) W[i] = w[blk + i];
      for (i = 16; i < 64; i++) {
        var x = W[i - 15], y = W[i - 2];
        var s0 = ((x >>> 7) | (x << 25)) ^ ((x >>> 18) | (x << 14)) ^ (x >>> 3);
        var s1 = ((y >>> 17) | (y << 15)) ^ ((y >>> 19) | (y << 13)) ^ (y >>> 10);
        W[i] = (W[i - 16] + s0 + W[i - 7] + s1) | 0;
      }
      var a = H[0], b = H[1], c = H[2], d = H[3], e = H[4], f = H[5], g = H[6], h = H[7];
      for (i = 0; i < 64; i++) {
        var S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
        var t1 = (h + S1 + ((e & f) ^ (~e & g)) + K256[i] + W[i]) | 0;
        var S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
        var t2 = (S0 + ((a & b) ^ (a & c) ^ (b & c))) | 0;
        h = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
      }
      H[0] = (H[0] + a) | 0; H[1] = (H[1] + b) | 0; H[2] = (H[2] + c) | 0; H[3] = (H[3] + d) | 0;
      H[4] = (H[4] + e) | 0; H[5] = (H[5] + f) | 0; H[6] = (H[6] + g) | 0; H[7] = (H[7] + h) | 0;
    }
    return wordsToBytes(H);
  }

  // ---------- SHA-512 / SHA-384 (64-bit words as hi/lo 32-bit pairs) ----------
  var K512 = [0x428a2f98,0xd728ae22,0x71374491,0x23ef65cd,0xb5c0fbcf,0xec4d3b2f,0xe9b5dba5,0x8189dbbc,0x3956c25b,0xf348b538,0x59f111f1,0xb605d019,0x923f82a4,0xaf194f9b,0xab1c5ed5,0xda6d8118,0xd807aa98,0xa3030242,0x12835b01,0x45706fbe,0x243185be,0x4ee4b28c,0x550c7dc3,0xd5ffb4e2,0x72be5d74,0xf27b896f,0x80deb1fe,0x3b1696b1,0x9bdc06a7,0x25c71235,0xc19bf174,0xcf692694,0xe49b69c1,0x9ef14ad2,0xefbe4786,0x384f25e3,0x0fc19dc6,0x8b8cd5b5,0x240ca1cc,0x77ac9c65,0x2de92c6f,0x592b0275,0x4a7484aa,0x6ea6e483,0x5cb0a9dc,0xbd41fbd4,0x76f988da,0x831153b5,0x983e5152,0xee66dfab,0xa831c66d,0x2db43210,0xb00327c8,0x98fb213f,0xbf597fc7,0xbeef0ee4,0xc6e00bf3,0x3da88fc2,0xd5a79147,0x930aa725,0x06ca6351,0xe003826f,0x14292967,0x0a0e6e70,0x27b70a85,0x46d22ffc,0x2e1b2138,0x5c26c926,0x4d2c6dfc,0x5ac42aed,0x53380d13,0x9d95b3df,0x650a7354,0x8baf63de,0x766a0abb,0x3c77b2a8,0x81c2c92e,0x47edaee6,0x92722c85,0x1482353b,0xa2bfe8a1,0x4cf10364,0xa81a664b,0xbc423001,0xc24b8b70,0xd0f89791,0xc76c51a3,0x0654be30,0xd192e819,0xd6ef5218,0xd6990624,0x5565a910,0xf40e3585,0x5771202a,0x106aa070,0x32bbd1b8,0x19a4c116,0xb8d2d0c8,0x1e376c08,0x5141ab53,0x2748774c,0xdf8eeb99,0x34b0bcb5,0xe19b48a8,0x391c0cb3,0xc5c95a63,0x4ed8aa4a,0xe3418acb,0x5b9cca4f,0x7763e373,0x682e6ff3,0xd6b2b8a3,0x748f82ee,0x5defb2fc,0x78a5636f,0x43172f60,0x84c87814,0xa1f0ab72,0x8cc70208,0x1a6439ec,0x90befffa,0x23631e28,0xa4506ceb,0xde82bde9,0xbef9a3f7,0xb2c67915,0xc67178f2,0xe372532b,0xca273ece,0xea26619c,0xd186b8c7,0x21c0c207,0xeada7dd6,0xcde0eb1e,0xf57d4f7f,0xee6ed178,0x06f067aa,0x72176fba,0x0a637dc5,0xa2c898a6,0x113f9804,0xbef90dae,0x1b710b35,0x131c471b,0x28db77f5,0x23047d84,0x32caab7b,0x40c72493,0x3c9ebe0a,0x15c9bebc,0x431d67c4,0x9c100d4c,0x4cc5d4be,0xcb3e42b6,0x597f299c,0xfc657e2a,0x5fcb6fab,0x3ad6faec,0x6c44198c,0x4a475817].map(function (x) { return x | 0; });
  var IV512 = [0x6a09e667,0xf3bcc908,0xbb67ae85,0x84caa73b,0x3c6ef372,0xfe94f82b,0xa54ff53a,0x5f1d36f1,0x510e527f,0xade682d1,0x9b05688c,0x2b3e6c1f,0x1f83d9ab,0xfb41bd6b,0x5be0cd19,0x137e2179];
  var IV384 = [0xcbbb9d5d,0xc1059ed8,0x629a292a,0x367cd507,0x9159015a,0x3070dd17,0x152fecd8,0xf70e5939,0x67332667,0xffc00b31,0x8eb44a87,0x68581511,0xdb0c2e0d,0x64f98fa7,0x47b5481d,0xbefa4fa4];
  function sha512core(msg, iv, outWords) {
    var ml = msg.length, nb = ((ml + 16) >> 7) + 1, w = new Int32Array(nb * 32), i;
    for (i = 0; i < ml; i++) w[i >> 2] |= msg[i] << (24 - (i & 3) * 8);
    w[ml >> 2] |= 0x80 << (24 - (ml & 3) * 8);
    w[nb * 32 - 1] = ml * 8; w[nb * 32 - 2] = Math.floor(ml / 0x20000000);
    var H = iv.map(function (x) { return x | 0; }), W = new Int32Array(160);
    for (var blk = 0; blk < nb * 32; blk += 32) {
      for (i = 0; i < 32; i++) W[i] = w[blk + i];
      for (i = 32; i < 160; i += 2) {
        var xh = W[i - 30], xl = W[i - 29];
        var s0h = ((xh >>> 1) | (xl << 31)) ^ ((xh >>> 8) | (xl << 24)) ^ (xh >>> 7);
        var s0l = ((xl >>> 1) | (xh << 31)) ^ ((xl >>> 8) | (xh << 24)) ^ ((xl >>> 7) | (xh << 25));
        var yh = W[i - 4], yl = W[i - 3];
        var s1h = ((yh >>> 19) | (yl << 13)) ^ ((yl >>> 29) | (yh << 3)) ^ (yh >>> 6);
        var s1l = ((yl >>> 19) | (yh << 13)) ^ ((yh >>> 29) | (yl << 3)) ^ ((yl >>> 6) | (yh << 26));
        var lo = (W[i - 31] >>> 0) + (s0l >>> 0) + (W[i - 13] >>> 0) + (s1l >>> 0);
        W[i] = (W[i - 32] + s0h + W[i - 14] + s1h + ((lo / 0x100000000) | 0)) | 0;
        W[i + 1] = lo | 0;
      }
      var ah = H[0], al = H[1], bh = H[2], bl = H[3], ch = H[4], cl = H[5], dh = H[6], dl = H[7];
      var eh = H[8], el = H[9], fh = H[10], fl = H[11], gh = H[12], gl = H[13], hh = H[14], hl = H[15];
      for (i = 0; i < 160; i += 2) {
        var S1h = ((eh >>> 14) | (el << 18)) ^ ((eh >>> 18) | (el << 14)) ^ ((el >>> 9) | (eh << 23));
        var S1l = ((el >>> 14) | (eh << 18)) ^ ((el >>> 18) | (eh << 14)) ^ ((eh >>> 9) | (el << 23));
        var chh = (eh & fh) ^ (~eh & gh), chl = (el & fl) ^ (~el & gl);
        var t1l = (hl >>> 0) + (S1l >>> 0) + (chl >>> 0) + (K512[i + 1] >>> 0) + (W[i + 1] >>> 0);
        var t1h = (hh + S1h + chh + K512[i] + W[i] + ((t1l / 0x100000000) | 0)) | 0;
        t1l = t1l | 0;
        var S0h = ((ah >>> 28) | (al << 4)) ^ ((al >>> 2) | (ah << 30)) ^ ((al >>> 7) | (ah << 25));
        var S0l = ((al >>> 28) | (ah << 4)) ^ ((ah >>> 2) | (al << 30)) ^ ((ah >>> 7) | (al << 25));
        var mjh = (ah & bh) ^ (ah & ch) ^ (bh & ch), mjl = (al & bl) ^ (al & cl) ^ (bl & cl);
        var t2l = (S0l >>> 0) + (mjl >>> 0);
        var t2h = (S0h + mjh + ((t2l / 0x100000000) | 0)) | 0;
        t2l = t2l | 0;
        hh = gh; hl = gl; gh = fh; gl = fl; fh = eh; fl = el;
        var nel = (dl >>> 0) + (t1l >>> 0); eh = (dh + t1h + ((nel / 0x100000000) | 0)) | 0; el = nel | 0;
        dh = ch; dl = cl; ch = bh; cl = bl; bh = ah; bl = al;
        var nal = (t1l >>> 0) + (t2l >>> 0); ah = (t1h + t2h + ((nal / 0x100000000) | 0)) | 0; al = nal | 0;
      }
      var vals = [ah, al, bh, bl, ch, cl, dh, dl, eh, el, fh, fl, gh, gl, hh, hl];
      for (i = 0; i < 16; i += 2) {
        var l = (H[i + 1] >>> 0) + (vals[i + 1] >>> 0);
        H[i] = (H[i] + vals[i] + ((l / 0x100000000) | 0)) | 0; H[i + 1] = l | 0;
      }
    }
    return wordsToBytes(H.slice(0, outWords));
  }
  function sha512(m) { return sha512core(m, IV512, 16); }
  function sha384(m) { return sha512core(m, IV384, 12); }

  var HASHES = { SHA1: sha1, 'SHA-1': sha1, SHA256: sha256, 'SHA-256': sha256, SHA384: sha384, 'SHA-384': sha384, SHA512: sha512, 'SHA-512': sha512 };

  // ---------- AES (WebCrypto) without padding ----------
  // WebCrypto's AES-CBC always strips PKCS#7 padding, but Office data has none. Appending one
  // extra block that is the encryption of an empty message (IV = last ciphertext block) makes the
  // data end in a valid padding block, which WebCrypto then strips -- leaving exactly the raw data.
  var subtle = (global.crypto && global.crypto.subtle);
  async function importAes(keyBytes) { return subtle.importKey('raw', keyBytes, { name: 'AES-CBC' }, false, ['encrypt', 'decrypt']); }
  async function aesCbcRaw(key, iv, data) {
    if (!data.length) return new Uint8Array(0);
    if (data.length % 16) data = data.slice(0, data.length - (data.length % 16));
    var last = data.slice(data.length - 16);
    var pad = new Uint8Array(await subtle.encrypt({ name: 'AES-CBC', iv: last }, key, new Uint8Array(0)));
    return new Uint8Array(await subtle.decrypt({ name: 'AES-CBC', iv: iv }, key, concat(data, pad.slice(0, 16))));
  }
  async function aesEcbRaw(key, data) { // ECB = CBC(IV 0) with each block XOR'ed back with the previous ciphertext block
    var out = await aesCbcRaw(key, new Uint8Array(16), data);
    for (var i = 16; i < out.length; i++) out[i] ^= data[i - 16];
    return out;
  }

  // ---------- container ----------
  function isOleContainer(b) { return b.length > 8 && b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0; }
  function isZip(b) { return b.length > 4 && b[0] === 0x50 && b[1] === 0x4B; }
  function readStreams(bytes) {
    var cfb = global.XLSX.CFB.read(bytes, { type: 'array' });
    function get(name) {
      var e = global.XLSX.CFB.find(cfb, name) || global.XLSX.CFB.find(cfb, '/' + name);
      return e && e.content ? new Uint8Array(e.content) : null;
    }
    return { info: get('EncryptionInfo'), pkg: get('EncryptedPackage'), isXls: !!(get('Workbook') || get('Book')) };
  }

  // ---------- Agile ----------
  function attrs(xml, tagRe) {
    var m = tagRe.exec(xml); if (!m) return null;
    var out = {}, re = /([\w:]+)="([^"]*)"/g, a;
    while ((a = re.exec(m[0]))) out[a[1]] = a[2];
    return out;
  }
  var BK_VERIFIER_INPUT = new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]);
  var BK_VERIFIER_VALUE = new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]);
  var BK_KEY_VALUE = new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);

  function parseAgile(info) {
    var xml = new TextDecoder('utf-8').decode(info.slice(8));
    var kd = attrs(xml, /<(?:\w+:)?keyData\b[^>]*>/);
    var ek = attrs(xml, /<(?:\w+:)?encryptedKey\b[^>]*>/);
    if (!kd || !ek) throw new Error('UNSUPPORTED');
    return { kd: kd, ek: ek };
  }
  async function agileDecrypt(info, pkg, password, parsed) {
    var kd = parsed.kd, ek = parsed.ek;
    var H = HASHES[(ek.hashAlgorithm || '').toUpperCase()], Hk = HASHES[(kd.hashAlgorithm || '').toUpperCase()];
    if (!H || !Hk || !/^AES/i.test(ek.cipherAlgorithm || 'AES') || !/^AES/i.test(kd.cipherAlgorithm || 'AES')) throw new Error('UNSUPPORTED');
    var salt = b64(ek.saltValue), spin = parseInt(ek.spinCount, 10), keyLen = parseInt(ek.keyBits, 10) / 8;
    var h = H(concat(salt, utf16le(password)));
    for (var i = 0; i < spin; i++) h = H(concat(u32le(i), h));
    var k1 = await importAes(fitKey(H(concat(h, BK_VERIFIER_INPUT)), keyLen, 0x36));
    var k2 = await importAes(fitKey(H(concat(h, BK_VERIFIER_VALUE)), keyLen, 0x36));
    var vIn = (await aesCbcRaw(k1, salt, b64(ek.encryptedVerifierHashInput))).slice(0, salt.length);
    var vHash = (await aesCbcRaw(k2, salt, b64(ek.encryptedVerifierHashValue))).slice(0, H(vIn).length);
    if (!equalBytes(H(vIn), vHash)) return null; // wrong password
    var k3 = await importAes(fitKey(H(concat(h, BK_KEY_VALUE)), keyLen, 0x36));
    var secret = (await aesCbcRaw(k3, salt, b64(ek.encryptedKeyValue))).slice(0, parseInt(kd.keyBits, 10) / 8);
    var dataKey = await importAes(secret);
    var kdSalt = b64(kd.saltValue), block = parseInt(kd.blockSize, 10) || 16;
    var size = readU32(pkg, 0) + readU32(pkg, 4) * 0x100000000;
    var parts = [], seg = 0;
    for (var off = 8; off < pkg.length; off += 4096, seg++) {
      var iv = fitKey(Hk(concat(kdSalt, u32le(seg))), block, 0x36);
      parts.push(await aesCbcRaw(dataKey, iv, pkg.slice(off, Math.min(off + 4096, pkg.length))));
    }
    return concat.apply(null, parts).slice(0, size);
  }

  // ---------- Standard ----------
  function parseStandard(info) {
    var headerSize = readU32(info, 8), h = 12;
    var algId = readU32(info, h + 8), hashId = readU32(info, h + 12), keyBits = readU32(info, h + 16);
    if ([0x660E, 0x660F, 0x6610].indexOf(algId) < 0 || (hashId !== 0x8004 && hashId !== 0)) throw new Error('UNSUPPORTED'); // RC4 etc.
    var v = 12 + headerSize, saltSize = readU32(info, v);
    var salt = info.slice(v + 4, v + 4 + saltSize);
    var encVerifier = info.slice(v + 4 + saltSize, v + 20 + saltSize);
    var encVerifierHash = info.slice(v + 24 + saltSize, v + 24 + saltSize + 32);
    return { keyBits: keyBits || 128, salt: salt, encVerifier: encVerifier, encVerifierHash: encVerifierHash };
  }
  async function standardDecrypt(info, pkg, password, p) {
    var h = sha1(concat(p.salt, utf16le(password)));
    for (var i = 0; i < 50000; i++) h = sha1(concat(u32le(i), h));
    h = sha1(concat(h, u32le(0)));
    var b1 = new Uint8Array(64).fill(0x36), b2 = new Uint8Array(64).fill(0x5c);
    for (i = 0; i < h.length; i++) { b1[i] ^= h[i]; b2[i] ^= h[i]; }
    var key = await importAes(concat(sha1(b1), sha1(b2)).slice(0, p.keyBits / 8));
    var verifier = await aesEcbRaw(key, p.encVerifier);
    var vHash = (await aesEcbRaw(key, p.encVerifierHash)).slice(0, 20);
    if (!equalBytes(sha1(verifier), vHash)) return null; // wrong password
    var size = readU32(pkg, 0) + readU32(pkg, 4) * 0x100000000;
    return (await aesEcbRaw(key, pkg.slice(8))).slice(0, size);
  }

  // ---------- public ----------
  // inspect(bytes) -> { kind: 'plain' | 'agile' | 'standard' | 'xls' | 'unsupported' | 'notexcel', ... }
  function inspect(bytes) {
    if (isZip(bytes)) return { kind: 'plain' };
    if (!isOleContainer(bytes)) return { kind: 'notexcel' };
    var s;
    try { s = readStreams(bytes); } catch (e) { return { kind: 'notexcel' }; }
    if (!s.info || !s.pkg) return { kind: s.isXls ? 'xls' : 'notexcel' };
    var major = s.info[0] | (s.info[1] << 8), minor = s.info[2] | (s.info[3] << 8);
    try {
      if (major === 4 && minor === 4) return { kind: 'agile', info: s.info, pkg: s.pkg, parsed: parseAgile(s.info) };
      if ((major === 2 || major === 3 || major === 4) && minor === 2) return { kind: 'standard', info: s.info, pkg: s.pkg, parsed: parseStandard(s.info) };
    } catch (e) { /* fall through */ }
    return { kind: 'unsupported' };
  }
  // tryPassword(inspected, password) -> Uint8Array of the original workbook, or null if wrong.
  async function tryPassword(x, password) {
    var out = x.kind === 'agile' ? await agileDecrypt(x.info, x.pkg, password, x.parsed)
      : x.kind === 'standard' ? await standardDecrypt(x.info, x.pkg, password, x.parsed) : null;
    if (out && !isZip(out)) return null; // verifier passed but data isn't a workbook: treat as failure
    return out;
  }

  global.OfficeCrypto = { inspect: inspect, tryPassword: tryPassword, _hash: { sha1: sha1, sha256: sha256, sha384: sha384, sha512: sha512 } };
})(typeof window !== 'undefined' ? window : globalThis);
