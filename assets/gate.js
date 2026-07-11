(function () {
  var STORE_KEY = 'osi_gate_ok';
  var PW_HASH = 'c2b7c9563e1eb80557673d08616f0f9714255fa1417b23a913ba38e73cf7f150';

  function reveal() {
    document.documentElement.style.visibility = 'visible';
  }

  if (sessionStorage.getItem(STORE_KEY) === PW_HASH) {
    reveal();
    return;
  }

  async function sha256(text) {
    var buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
    return Array.from(new Uint8Array(buf)).map(function (b) { return b.toString(16).padStart(2, '0'); }).join('');
  }

  var style = document.createElement('style');
  style.textContent =
    '#osiGateOverlay{position:fixed;inset:0;z-index:2147483647;background:linear-gradient(160deg,#fdf5f6 0%,#f4f5f7 45%,#eef1f4 100%);' +
    'display:flex;align-items:center;justify-content:center;font-family:"IBM Plex Sans Thai","Segoe UI",Tahoma,Arial,sans-serif;}' +
    '.osi-gate-card{background:#fff;border-radius:14px;box-shadow:0 8px 28px rgba(0,0,0,.12);padding:32px 30px;width:100%;max-width:320px;text-align:center;}' +
    '.osi-gate-card img{height:46px;width:auto;margin-bottom:10px;}' +
    '.osi-gate-card h2{font-size:17px;font-weight:700;color:#1a1a2e;margin-bottom:4px;}' +
    '.osi-gate-card p{font-size:12.5px;color:#546e7a;margin-bottom:18px;}' +
    '.osi-gate-card input{width:100%;box-sizing:border-box;border:1px solid #cfd8dc;border-radius:8px;padding:10px 12px;font-size:14px;text-align:center;outline:none;transition:border-color .15s;font-family:inherit;}' +
    '.osi-gate-card input:focus{border-color:#c8102e;}' +
    '.osi-gate-card button{margin-top:12px;width:100%;background:#c8102e;color:#fff;border:none;border-radius:8px;padding:10px;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit;transition:background .15s;}' +
    '.osi-gate-card button:hover{background:#a30d24;}' +
    '.osi-gate-err{min-height:16px;font-size:12px;color:#c62828;margin-top:8px;}';
  document.head.appendChild(style);

  var iconLink = document.querySelector('link[rel="icon"]');
  var iconHref = iconLink ? iconLink.href : '';

  var overlay = document.createElement('div');
  overlay.id = 'osiGateOverlay';
  overlay.innerHTML =
    '<div class="osi-gate-card">' +
    (iconHref ? '<img src="' + iconHref + '" alt="">' : '') +
    '<h2>OSInvoice Tools</h2>' +
    '<p>กรุณาใส่รหัสผ่านทีมเพื่อเข้าใช้งาน</p>' +
    '<input type="password" id="osiGatePw" placeholder="รหัสผ่าน" autocomplete="off">' +
    '<button id="osiGateBtn" type="button">เข้าสู่ระบบ</button>' +
    '<div class="osi-gate-err" id="osiGateErr"></div>' +
    '</div>';
  document.body.appendChild(overlay);
  reveal();

  var input = document.getElementById('osiGatePw');
  var errBox = document.getElementById('osiGateErr');

  function tryUnlock() {
    var val = input.value;
    sha256(val).then(function (hash) {
      if (hash === PW_HASH) {
        sessionStorage.setItem(STORE_KEY, PW_HASH);
        overlay.remove();
      } else {
        errBox.textContent = 'รหัสผ่านไม่ถูกต้อง';
        input.value = '';
        input.focus();
      }
    });
  }

  document.getElementById('osiGateBtn').addEventListener('click', tryUnlock);
  input.addEventListener('keydown', function (e) { if (e.key === 'Enter') tryUnlock(); });
  input.focus();
})();
