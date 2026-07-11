(function () {
  if (document.getElementById('osiPeekCat')) return;

  var wrap = document.createElement('div');
  wrap.id = 'osiPeekCat';
  wrap.setAttribute('aria-hidden', 'true');
  document.body.appendChild(wrap);

  var style = document.createElement('style');
  style.textContent =
    '#osiPeekCat{position:fixed;right:0;top:42%;width:78px;height:96px;pointer-events:none;z-index:3;' +
    'transform:translateX(60%);animation:osiPeek 7s ease-in-out infinite;filter:drop-shadow(-2px 2px 3px rgba(0,0,0,.2));}' +
    '#osiPeekCat svg{display:block;width:100%;height:100%;}' +
    '#osiPeekCat .opc-eye{animation:osiBlink 7s ease-in-out infinite;transform-box:fill-box;transform-origin:50% 50%;}' +
    '@keyframes osiPeek{0%,55%{transform:translateX(60%)}64%,78%{transform:translateX(2%)}100%{transform:translateX(60%)}}' +
    '@keyframes osiBlink{0%,63%,72%,100%{transform:scaleY(1)}67%{transform:scaleY(.12)}}' +
    '@media (max-width:520px){#osiPeekCat{width:58px;height:72px}}';
  document.head.appendChild(style);

  // Returns a full jagged polyline from (x1,y1) to (x2,y2) INCLUDING both endpoints, so callers
  // can splice it straight into a path with no gap/jump at the seam.
  function zigzag(x1, y1, x2, y2, teeth, depth) {
    var dx = x2 - x1, dy = y2 - y1;
    var len = Math.sqrt(dx * dx + dy * dy) || 1;
    var nx = -dy / len, ny = dx / len;
    var pts = [x1.toFixed(1) + ',' + y1.toFixed(1)];
    var steps = teeth * 2;
    for (var i = 1; i < steps; i++) {
      var t = i / steps;
      var px = x1 + dx * t, py = y1 + dy * t;
      var d = (i % 2 === 1) ? depth : 0;
      pts.push((px + nx * d).toFixed(1) + ',' + (py + ny * d).toFixed(1));
    }
    pts.push(x2.toFixed(1) + ',' + y2.toFixed(1));
    return pts.join(' L');
  }

  var ear1 = 'M' + zigzag(24, 6, 12, 50, 3, 4) + ' L40,44 Z';
  var ear2 = 'M' + zigzag(64, 18, 54, 52, 2, 3.5) + ' L86,46 Z';
  var cheek = zigzag(22, 54, 17, 92, 4, 4);
  var head = 'M40,44 C34,44 27,47 22,54 L' + cheek +
    ' C17,110 27,126 44,136 C58,144 74,145 90,140 L120,140 L120,8 L86,46 L64,18 L40,44 Z';

  var svg =
    '<svg viewBox="0 0 120 150" xmlns="http://www.w3.org/2000/svg">' +
    '<path d="' + ear1 + '" fill="#111"/>' +
    '<path d="' + ear2 + '" fill="#111"/>' +
    '<path d="' + head + '" fill="#111"/>' +
    '<g stroke="#2a2a2a" stroke-width="1.6" stroke-linecap="round">' +
    '<line x1="26" y1="86" x2="0" y2="78"/>' +
    '<line x1="25" y1="95" x2="0" y2="95"/>' +
    '<line x1="27" y1="104" x2="2" y2="113"/>' +
    '</g>' +
    '<g class="opc-eye"><circle cx="50" cy="83" r="15" fill="#fff"/><circle cx="52" cy="85" r="8" fill="#111"/></g>' +
    '<g class="opc-eye"><circle cx="80" cy="61" r="10" fill="#fff"/><circle cx="82" cy="62" r="5.5" fill="#111"/></g>' +
    '<path d="M60,98 Q66,106 72,98" fill="none" stroke="#fff" stroke-width="1.3" stroke-linecap="round" opacity=".8"/>' +
    '</svg>';

  wrap.innerHTML = svg;
})();
