(function () {
  if (document.getElementById('osiPets')) return;

  var base = document.currentScript.src.replace(/pets\.js(\?.*)?$/, '');

  var container = document.createElement('div');
  container.id = 'osiPets';
  container.setAttribute('aria-hidden', 'true');
  document.body.appendChild(container);

  var style = document.createElement('style');
  style.textContent =
    '#osiPets{position:fixed;left:0;right:0;bottom:0;height:64px;overflow:hidden;pointer-events:none;z-index:-1;}' +
    '.osi-pet{position:absolute;bottom:2px;}' +
    '.osi-pet-flip{display:inline-block;}' +
    '.osi-pet-img{display:block;height:60px;width:auto;filter:drop-shadow(0 2px 2px rgba(0,0,0,.15));transform-origin:50% 100%;}' +
    '.osi-pet-img.sitting{animation:osiBreatheSit 1.8s ease-in-out infinite;}' +
    '.osi-pet-img.lying{animation:osiBreatheLie 2.4s ease-in-out infinite;}' +
    '.osi-walk-svg{display:block;height:52px;width:auto;filter:drop-shadow(0 2px 2px rgba(0,0,0,.15));}' +
    '.osi-walk-svg.walk-body{animation:osiBodyBob .32s ease-in-out infinite;}' +
    '.osi-walk-svg .leg-a{transform-box:fill-box;transform-origin:50% 0%;animation:osiLegSwing .32s ease-in-out infinite;}' +
    '.osi-walk-svg .leg-b{transform-box:fill-box;transform-origin:50% 0%;animation:osiLegSwing .32s ease-in-out infinite reverse;}' +
    '.osi-walk-svg .tail{transform-box:fill-box;transform-origin:100% 60%;animation:osiTailWag .32s ease-in-out infinite;}' +
    '.osi-walk-svg.hop-body{animation:osiHopBig .5s cubic-bezier(.3,0,.7,1) infinite;}' +
    '@keyframes osiBodyBob{0%,100%{transform:translateY(0)}50%{transform:translateY(-4px)}}' +
    '@keyframes osiLegSwing{0%,100%{transform:rotate(26deg)}50%{transform:rotate(-26deg)}}' +
    '@keyframes osiTailWag{0%,100%{transform:rotate(-11deg)}50%{transform:rotate(13deg)}}' +
    '@keyframes osiHopBig{0%,100%{transform:translateY(0) scaleY(1)}30%{transform:translateY(-22px) scaleY(1.06)}62%{transform:translateY(-22px) scaleY(1.06)}88%{transform:translateY(0) scaleY(.8)}}' +
    '@keyframes osiBreatheSit{0%,100%{transform:scale(1) rotate(0deg)}50%{transform:scale(1.05,1.07) rotate(3deg)}}' +
    '@keyframes osiBreatheLie{' +
    '0%,100%{transform:rotate(-84deg) translate(-8px,6px) scale(1)}' +
    '50%{transform:rotate(-84deg) translate(-8px,6px) scale(1.04,1.035)}' +
    '}' +
    '@keyframes osiPetZzz{0%{opacity:0;transform:translateY(4px)}35%{opacity:.85}100%{opacity:0;transform:translateY(-11px)}}' +
    '.osi-pet-zzz{position:absolute;top:-10px;right:-2px;font-size:13px;opacity:0;animation:osiPetZzz 2.2s ease-in-out infinite;}' +
    '@media (max-width:520px){#osiPets{height:52px}.osi-pet-img{height:46px}.osi-walk-svg{height:44px}}';
  document.head.appendChild(style);

  function leg(cls, x, y, dark) {
    return '<rect class="' + cls + '" x="' + x + '" y="' + y + '" width="3.2" height="8" rx="1.6" fill="' + dark + '"/>';
  }

  // Golden puppy walk cycle: real alternating legs + wagging tail, matches the sitting art's palette.
  function dogWalkSVG() {
    var fur = '#f3c878', dark = '#dba956';
    return '<svg class="osi-walk-svg walk-body" viewBox="0 0 46 36">' +
      '<ellipse cx="17" cy="21" rx="10" ry="6.5" fill="' + fur + '"/>' +
      leg('leg-a', 13, 25, dark) + leg('leg-b', 22, 25, dark) +
      '<path class="tail" d="M8,19 Q0,12 4,5" stroke="' + fur + '" stroke-width="3.6" fill="none" stroke-linecap="round"/>' +
      '<circle cx="32" cy="13" r="9" fill="' + fur + '"/>' +
      '<path d="M25,10 Q19,15 24,23 Q28,19 28,11 Z" fill="' + dark + '"/>' +
      '<ellipse cx="39" cy="16" rx="3.4" ry="2.5" fill="' + dark + '"/>' +
      '<circle cx="34.5" cy="10.5" r="1.2" fill="#2b2b2b"/>' +
      '<circle cx="34" cy="10" r=".4" fill="#fff"/>' +
      '</svg>';
  }

  // Cream kitten walk cycle: pointy ears instead of floppy, otherwise same rig.
  function catWalkSVG() {
    var fur = '#fbecd6', dark = '#e8cfa3';
    return '<svg class="osi-walk-svg walk-body" viewBox="0 0 46 36">' +
      '<ellipse cx="17" cy="21" rx="10" ry="6.5" fill="' + fur + '"/>' +
      leg('leg-a', 13, 25, dark) + leg('leg-b', 22, 25, dark) +
      '<path class="tail" d="M8,19 Q-1,10 5,3" stroke="' + fur + '" stroke-width="3.4" fill="none" stroke-linecap="round"/>' +
      '<circle cx="32" cy="13" r="9" fill="' + fur + '"/>' +
      '<path d="M26,7 L27,1 L31,6 Z" fill="' + fur + '"/>' +
      '<path d="M33,6 L37,1 L38,7 Z" fill="' + fur + '"/>' +
      '<circle cx="34.5" cy="10.5" r="1.2" fill="#2b2b2b"/>' +
      '<circle cx="34" cy="10" r=".4" fill="#fff"/>' +
      '<circle cx="39.5" cy="13.5" r=".9" fill="#e8a9a0"/>' +
      '</svg>';
  }

  // Rabbits don't stride, they hop — a big vertical arc with tucked legs reads truer to life
  // than trying to fake alternating footsteps.
  function rabbitHopSVG() {
    var fur = '#fbf7f0', inner = '#f6b9c9';
    return '<svg class="osi-walk-svg hop-body" viewBox="0 0 46 36">' +
      '<ellipse cx="18" cy="23" rx="9.5" ry="6" fill="' + fur + '"/>' +
      '<ellipse cx="10" cy="27" rx="3" ry="1.8" fill="#e9dfd0"/>' +
      '<ellipse cx="24" cy="27" rx="3" ry="1.8" fill="#e9dfd0"/>' +
      '<circle cx="7" cy="21" r="3" fill="#fff"/>' +
      '<circle cx="30" cy="14" r="8.5" fill="' + fur + '"/>' +
      '<ellipse cx="26" cy="2" rx="2.1" ry="7" fill="' + fur + '" transform="rotate(-10 26 2)"/>' +
      '<ellipse cx="26" cy="3" rx="1.1" ry="5.2" fill="' + inner + '" transform="rotate(-10 26 3)"/>' +
      '<ellipse cx="33" cy="1.5" rx="2.1" ry="7.3" fill="' + fur + '" transform="rotate(9 33 1.5)"/>' +
      '<ellipse cx="33" cy="2.5" rx="1.1" ry="5.4" fill="' + inner + '" transform="rotate(9 33 2.5)"/>' +
      '<circle cx="32.5" cy="11.5" r="1.2" fill="#2b2b2b"/>' +
      '<circle cx="32" cy="11" r=".4" fill="#fff"/>' +
      '<circle cx="37.5" cy="14.5" r=".9" fill="' + inner + '"/>' +
      '</svg>';
  }

  var PET_DEFS = [
    { restSrc: 'puppy.png', walkSVG: dogWalkSVG, speed: 62 },
    { restSrc: 'kitten.png', walkSVG: catWalkSVG, speed: 54 },
    { restSrc: 'rabbit.png', walkSVG: rabbitHopSVG, speed: 84 }
  ];

  function maxX() { return Math.max(40, (window.innerWidth || 800) - 60); }

  var pets = PET_DEFS.map(function (def) {
    var wrap = document.createElement('div');
    wrap.className = 'osi-pet';
    var flip = document.createElement('div');
    flip.className = 'osi-pet-flip';
    flip.innerHTML = def.walkSVG();
    wrap.appendChild(flip);
    container.appendChild(wrap);

    return {
      el: wrap, flip: flip, restSrc: def.restSrc, walkSVG: def.walkSVG,
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
    if (p.zzz) { p.zzz.remove(); p.zzz = null; }
    if (state === 'walk') {
      p.flip.innerHTML = p.walkSVG();
      p.timer = 4 + Math.random() * 5;
      if (Math.random() < 0.4) p.dir *= -1;
    } else {
      var img = document.createElement('img');
      img.className = 'osi-pet-img ' + (state === 'sit' ? 'sitting' : 'lying');
      img.src = base + p.restSrc;
      img.alt = '';
      p.flip.innerHTML = '';
      p.flip.appendChild(img);
      if (state === 'sit') {
        p.timer = 2 + Math.random() * 2.5;
      } else {
        p.timer = 4 + Math.random() * 4;
        var z = document.createElement('div');
        z.className = 'osi-pet-zzz';
        z.textContent = '💤';
        p.flip.appendChild(z);
        p.zzz = z;
      }
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
