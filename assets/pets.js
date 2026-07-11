(function () {
  if (document.getElementById('osiPets')) return;

  var container = document.createElement('div');
  container.id = 'osiPets';
  container.setAttribute('aria-hidden', 'true');
  document.body.appendChild(container);

  var style = document.createElement('style');
  style.textContent =
    '#osiPets{position:fixed;left:0;right:0;bottom:0;height:56px;overflow:hidden;pointer-events:none;z-index:-1;}' +
    '.osi-pet{position:absolute;bottom:4px;}' +
    '.osi-pet-flip{display:inline-block;}' +
    '.osi-pet-fig{display:block;width:46px;height:36px;filter:drop-shadow(0 1px 1.5px rgba(0,0,0,.25));}' +
    '.osi-pet-fig .opl-leg{transform-box:fill-box;transform-origin:50% 0%;}' +
    '.walking .opl-leg-a{animation:osiLegSwing .42s ease-in-out infinite;}' +
    '.walking .opl-leg-b{animation:osiLegSwing .42s ease-in-out infinite reverse;}' +
    '.walking .opl-tail{transform-box:fill-box;transform-origin:100% 50%;animation:osiTailWag .42s ease-in-out infinite;}' +
    '.walking.osi-hop .opl-fig{animation:osiHop .42s ease-in-out infinite;}' +
    '.walking:not(.osi-hop) .opl-fig{animation:osiPetBob .42s ease-in-out infinite;}' +
    '.osi-pet-zzz{position:absolute;top:-8px;right:0;font-size:11px;opacity:0;animation:osiPetZzz 2.2s ease-in-out infinite;}' +
    '@keyframes osiLegSwing{0%,100%{transform:rotate(20deg)}50%{transform:rotate(-20deg)}}' +
    '@keyframes osiTailWag{0%,100%{transform:rotate(-9deg)}50%{transform:rotate(9deg)}}' +
    '@keyframes osiPetBob{0%,100%{transform:translateY(0)}50%{transform:translateY(-2px)}}' +
    '@keyframes osiHop{0%,60%,100%{transform:translateY(0)}30%{transform:translateY(-6px)}}' +
    '@keyframes osiPetZzz{0%{opacity:0;transform:translateY(4px)}35%{opacity:.85}100%{opacity:0;transform:translateY(-11px)}}' +
    '@media (max-width:520px){#osiPets{height:48px}.osi-pet-fig{width:38px;height:30px}}';
  document.head.appendChild(style);

  // Cute "kawaii sticker" palette — chunky dark outline + soft pastel fill, big head, big eyes.
  var PALETTE = {
    cat: { fur: '#f4a95f', dark: '#e08a3a', inner: '#fff0d9', line: '#7a4a1e' },
    dog: { fur: '#f6cd83', dark: '#e0a94f', inner: '#8a5a2a', line: '#7a5230' },
    rabbit: { fur: '#fbf6ef', dark: '#e9dfd0', inner: '#f7bdd0', line: '#8a7968' }
  };

  var L = ' stroke-linejoin="round"';
  function o(color) { return ' stroke="' + color + '" stroke-width="1.1"' + L; }

  function faceDetail(hx, hy, hr, p, closed) {
    // Big sparkly eye + highlight + blush + button nose + tiny smile, kawaii-style.
    var ex = hx + hr * 0.32, ey = hy - hr * 0.12;
    var nx = hx + hr * 0.86, ny = hy + hr * 0.22;
    var cx = hx + hr * 0.05, cy = hy + hr * 0.42;
    var s = '';
    if (closed) {
      s += '<path d="M' + (ex - 1.6) + ',' + ey + ' q1.8,1.6 3.6,0" fill="none" stroke="#2b2b2b" stroke-width="1" stroke-linecap="round"/>';
    } else {
      s += '<circle cx="' + ex + '" cy="' + ey + '" r="1.7" fill="#2b2b2b"/>';
      s += '<circle cx="' + (ex - 0.55) + '" cy="' + (ey - 0.55) + '" r=".55" fill="#fff"/>';
    }
    s += '<ellipse cx="' + cx + '" cy="' + cy + '" rx="2" ry="1.3" fill="#ffb3c1" opacity=".5"/>';
    s += '<circle cx="' + nx + '" cy="' + ny + '" r="1" fill="' + p.line + '"/>';
    s += '<path d="M' + (nx - 1.6) + ',' + (ny + 1.2) + ' q1.6,1.4 3,.1" fill="none" stroke="' + p.line + '" stroke-width=".9" stroke-linecap="round"/>';
    return s;
  }

  function leg(cls, x, y, p) {
    return '<rect class="opl-leg ' + cls + '" x="' + x + '" y="' + y + '" width="3.4" height="6.5" rx="1.7" fill="' + p.dark + '"' + o(p.line) + '/>';
  }

  function ears(kind, hx, hy, hr, p) {
    if (kind === 'cat') {
      return '<path d="M' + (hx - hr * 0.65) + ',' + (hy - hr * 0.55) + ' L' + (hx - hr * 0.5) + ',' + (hy - hr * 1.5) + ' L' + (hx - hr * 0.05) + ',' + (hy - hr * 0.65) + ' Z" fill="' + p.fur + '"' + o(p.line) + '/>' +
        '<path d="M' + (hx + hr * 0.1) + ',' + (hy - hr * 0.65) + ' L' + (hx + hr * 0.5) + ',' + (hy - hr * 1.5) + ' L' + (hx + hr * 0.62) + ',' + (hy - hr * 0.55) + ' Z" fill="' + p.fur + '"' + o(p.line) + '/>';
    }
    if (kind === 'dog') {
      return '<path d="M' + (hx - hr * 0.55) + ',' + (hy - hr * 0.4) + ' Q' + (hx - hr * 1.15) + ',' + (hy) + ' ' + (hx - hr * 0.6) + ',' + (hy + hr * 0.85) + ' Q' + (hx - hr * 0.25) + ',' + (hy + hr * 0.35) + ' ' + (hx - hr * 0.2) + ',' + (hy - hr * 0.5) + ' Z" fill="' + p.dark + '"' + o(p.line) + '/>';
    }
    return '<ellipse cx="' + (hx - hr * 0.3) + '" cy="' + (hy - hr * 1.25) + '" rx="2.3" ry="' + (hr * 0.85) + '" fill="' + p.fur + '" transform="rotate(-12 ' + (hx - hr * 0.3) + ' ' + (hy - hr * 1.25) + ')"' + o(p.line) + '/>' +
      '<ellipse cx="' + (hx - hr * 0.3) + '" cy="' + (hy - hr * 1.1) + '" rx="1.1" ry="' + (hr * 0.55) + '" fill="' + p.inner + '" transform="rotate(-12 ' + (hx - hr * 0.3) + ' ' + (hy - hr * 1.1) + ')"/>' +
      '<ellipse cx="' + (hx + hr * 0.3) + '" cy="' + (hy - hr * 1.3) + '" rx="2.3" ry="' + (hr * 0.9) + '" fill="' + p.fur + '" transform="rotate(10 ' + (hx + hr * 0.3) + ' ' + (hy - hr * 1.3) + ')"' + o(p.line) + '/>' +
      '<ellipse cx="' + (hx + hr * 0.3) + '" cy="' + (hy - hr * 1.15) + '" rx="1.1" ry="' + (hr * 0.6) + '" fill="' + p.inner + '" transform="rotate(10 ' + (hx + hr * 0.3) + ' ' + (hy - hr * 1.15) + ')"/>';
  }

  function earsFlat(kind, hx, hy, p) {
    if (kind === 'cat') {
      return '<path d="M' + (hx - 6) + ',' + (hy - 3) + ' L' + (hx - 11) + ',' + (hy - 5) + ' L' + (hx - 4) + ',' + (hy - 7) + ' Z" fill="' + p.fur + '"' + o(p.line) + '/>';
    }
    if (kind === 'dog') {
      return '<path d="M' + (hx - 4) + ',' + (hy - 3) + ' Q' + (hx - 11) + ',' + (hy - 1) + ' ' + (hx - 9) + ',' + (hy + 4) + ' Q' + (hx - 6) + ',' + (hy + 2) + ' ' + (hx - 3) + ',' + (hy - 4) + ' Z" fill="' + p.dark + '"' + o(p.line) + '/>';
    }
    return '<ellipse cx="' + (hx - 9) + '" cy="' + (hy - 4) + '" rx="7" ry="2.1" fill="' + p.fur + '" transform="rotate(-8 ' + (hx - 9) + ' ' + (hy - 4) + ')"' + o(p.line) + '/>' +
      '<ellipse cx="' + (hx - 10) + '" cy="' + (hy - 1) + '" rx="7.3" ry="2.1" fill="' + p.fur + '" transform="rotate(-4 ' + (hx - 10) + ' ' + (hy - 1) + ')"' + o(p.line) + '/>';
  }

  function tail(kind, x, y, p, cls) {
    var animCls = cls || '';
    if (kind === 'cat') {
      return '<path class="opl-tail ' + animCls + '" d="M' + x + ',' + y + ' Q' + (x - 9) + ',' + (y - 10) + ' ' + (x - 3) + ',' + (y - 17) + '" stroke="' + p.fur + '" stroke-width="3.4" fill="none" stroke-linecap="round"/><path class="opl-tail ' + animCls + '" d="M' + x + ',' + y + ' Q' + (x - 9) + ',' + (y - 10) + ' ' + (x - 3) + ',' + (y - 17) + '" stroke="' + p.line + '" stroke-width="4.4" fill="none" stroke-linecap="round" opacity="0" />';
    }
    if (kind === 'dog') {
      return '<path class="opl-tail ' + animCls + '" d="M' + x + ',' + y + ' Q' + (x - 8) + ',' + (y - 8) + ' ' + (x - 4) + ',' + (y - 13) + '" stroke="' + p.fur + '" stroke-width="3.6" fill="none" stroke-linecap="round"/>';
    }
    return '<circle cx="' + (x - 2) + '" cy="' + (y - 2) + '" r="3.4" fill="#ffffff"' + o(p.line) + '/>';
  }

  // ── Pose builders (big chibi head, small round body) ────
  function walkPose(kind, p) {
    var bx = 17, by = 21, brx = 9, bry = 6;
    var hx = 31, hy = 14, hr = 9.5;
    var svg = '<ellipse cx="' + bx + '" cy="' + by + '" rx="' + brx + '" ry="' + bry + '" fill="' + p.fur + '"' + o(p.line) + '/>';
    svg += leg('opl-leg-a', bx - 3.5, by + 3.5, p);
    svg += leg('opl-leg-b', bx + 5, by + 3.5, p);
    svg += tail(kind, bx - 8, by - 1, p, 'opl-tail');
    svg += '<circle cx="' + hx + '" cy="' + hy + '" r="' + hr + '" fill="' + p.fur + '"' + o(p.line) + '/>';
    svg += ears(kind, hx, hy, hr, p);
    if (kind === 'dog') svg += '<ellipse cx="' + (hx + hr * 0.68) + '" cy="' + (hy + hr * 0.3) + '" rx="3.4" ry="2.5" fill="' + p.dark + '"' + o(p.line) + '/>';
    svg += faceDetail(hx, hy, hr, p, false);
    return svg;
  }

  function sitPose(kind, p) {
    var bx = 16, by = 23, brx = 9, bry = 8.5;
    var hx = 19, hy = 10.5, hr = 9.8;
    var svg = '<ellipse cx="' + bx + '" cy="' + by + '" rx="' + brx + '" ry="' + bry + '" fill="' + p.fur + '"' + o(p.line) + '/>';
    svg += '<ellipse cx="' + (bx - 6) + '" cy="' + (by + 6.5) + '" rx="2.4" ry="1.7" fill="' + p.dark + '"' + o(p.line) + '/>';
    svg += '<ellipse cx="' + (bx + 5) + '" cy="' + (by + 6.5) + '" rx="2.4" ry="1.7" fill="' + p.dark + '"' + o(p.line) + '/>';
    svg += tail(kind, bx - 8, by + 4, p, '');
    svg += '<circle cx="' + hx + '" cy="' + hy + '" r="' + hr + '" fill="' + p.fur + '"' + o(p.line) + '/>';
    svg += ears(kind, hx, hy, hr, p);
    if (kind === 'dog') svg += '<ellipse cx="' + (hx + hr * 0.65) + '" cy="' + (hy + hr * 0.28) + '" rx="3.3" ry="2.4" fill="' + p.dark + '"' + o(p.line) + '/>';
    svg += faceDetail(hx, hy, hr, p, false);
    return svg;
  }

  function liePose(kind, p) {
    var bx = 20, by = 26, brx = 15, bry = 5.5;
    var hx = 36, hy = 21, hr = 8;
    var svg = '<ellipse cx="' + bx + '" cy="' + by + '" rx="' + brx + '" ry="' + bry + '" fill="' + p.fur + '"' + o(p.line) + '/>';
    svg += '<ellipse cx="' + (bx - 11) + '" cy="' + (by + 3.8) + '" rx="3" ry="1.8" fill="' + p.dark + '"' + o(p.line) + '/>';
    svg += tail(kind, bx - 14, by - 1, p, '');
    svg += '<circle cx="' + hx + '" cy="' + hy + '" r="' + hr + '" fill="' + p.fur + '"' + o(p.line) + '/>';
    svg += earsFlat(kind, hx, hy, p);
    if (kind === 'dog') svg += '<ellipse cx="' + (hx + 5.4) + '" cy="' + (hy + 2.1) + '" rx="3" ry="2.1" fill="' + p.dark + '"' + o(p.line) + '/>';
    svg += faceDetail(hx, hy, hr, p, true);
    return svg;
  }

  function buildInner(kind, pose, p) {
    var body = pose === 'walk' ? walkPose(kind, p) : pose === 'sit' ? sitPose(kind, p) : liePose(kind, p);
    return '<svg class="osi-pet-fig" viewBox="0 0 46 36"><g class="opl-fig">' + body + '</g></svg>';
  }

  var PET_DEFS = [
    { kind: 'cat', speed: 38, hop: false },
    { kind: 'dog', speed: 50, hop: false },
    { kind: 'rabbit', speed: 60, hop: true }
  ];

  function maxX() { return Math.max(40, (window.innerWidth || 800) - 46); }

  var pets = PET_DEFS.map(function (def) {
    var wrap = document.createElement('div');
    wrap.className = 'osi-pet';
    var flip = document.createElement('div');
    flip.className = 'osi-pet-flip';
    var pose = document.createElement('div');
    pose.className = 'osi-pet-pose walking' + (def.hop ? ' osi-hop' : '');
    pose.innerHTML = buildInner(def.kind, 'walk', PALETTE[def.kind]);
    flip.appendChild(pose);
    wrap.appendChild(flip);
    container.appendChild(wrap);

    return {
      el: wrap, flip: flip, pose: pose, kind: def.kind, hop: def.hop,
      x: Math.random() * maxX(),
      dir: Math.random() < 0.5 ? 1 : -1,
      speed: def.speed,
      state: 'walk',
      timer: 3 + Math.random() * 5,
      zzz: null
    };
  });

  function setState(p, state) {
    p.state = state;
    p.pose.classList.remove('walking', 'sitting', 'lying');
    if (p.zzz) { p.zzz.remove(); p.zzz = null; }
    p.pose.innerHTML = buildInner(p.kind, state === 'walk' ? 'walk' : state === 'sit' ? 'sit' : 'lie', PALETTE[p.kind]);
    if (state === 'walk') {
      p.pose.classList.add('walking');
      if (p.hop) p.pose.classList.add('osi-hop');
      p.timer = 4 + Math.random() * 5;
      if (Math.random() < 0.4) p.dir *= -1;
    } else if (state === 'sit') {
      p.pose.classList.add('sitting');
      p.timer = 2 + Math.random() * 2.5;
    } else if (state === 'lie') {
      p.pose.classList.add('lying');
      p.timer = 4 + Math.random() * 4;
      var z = document.createElement('div');
      z.className = 'osi-pet-zzz';
      z.textContent = '💤';
      p.flip.appendChild(z);
      p.zzz = z;
    }
  }

  var last = null;
  function frame(ts) {
    if (last === null) last = ts;
    var dt = Math.min((ts - last) / 1000, 0.1);
    last = ts;
    var mx = maxX();

    pets.forEach(function (p) {
      p.timer -= dt;
      if (p.state === 'walk') {
        p.x += p.dir * p.speed * dt;
        if (p.x <= 0) { p.x = 0; p.dir = 1; }
        if (p.x >= mx) { p.x = mx; p.dir = -1; }
        p.flip.style.transform = 'scaleX(' + p.dir + ')';
      }
      if (p.timer <= 0) {
        setState(p, p.state === 'walk' ? (Math.random() < 0.55 ? 'sit' : 'lie') : 'walk');
      }
      p.el.style.left = p.x + 'px';
    });

    requestAnimationFrame(frame);
  }

  pets.forEach(function (p) {
    p.el.style.left = p.x + 'px';
    p.flip.style.transform = 'scaleX(' + p.dir + ')';
  });
  requestAnimationFrame(frame);
})();
