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
    '.osi-pet-img.walking{animation:osiHop .5s ease-in-out infinite;}' +
    '.osi-pet-img.sitting{transform:translateY(0);}' +
    '.osi-pet-img.lying{transform:rotate(-84deg) translate(-8px,6px);transition:transform .3s ease;}' +
    '.osi-pet-zzz{position:absolute;top:-10px;right:-2px;font-size:13px;opacity:0;animation:osiPetZzz 2.2s ease-in-out infinite;}' +
    '@keyframes osiHop{0%,100%{transform:translateY(0)}50%{transform:translateY(-7px)}}' +
    '@keyframes osiPetZzz{0%{opacity:0;transform:translateY(4px)}35%{opacity:.85}100%{opacity:0;transform:translateY(-11px)}}' +
    '@media (max-width:520px){#osiPets{height:52px}.osi-pet-img{height:46px}}';
  document.head.appendChild(style);

  var PET_DEFS = [
    { src: 'puppy.png', speed: 42 },
    { src: 'kitten.png', speed: 36 },
    { src: 'rabbit.png', speed: 58 }
  ];

  function maxX() { return Math.max(40, (window.innerWidth || 800) - 60); }

  var pets = PET_DEFS.map(function (def) {
    var wrap = document.createElement('div');
    wrap.className = 'osi-pet';
    var flip = document.createElement('div');
    flip.className = 'osi-pet-flip';
    var img = document.createElement('img');
    img.className = 'osi-pet-img walking';
    img.src = base + def.src;
    img.alt = '';
    flip.appendChild(img);
    wrap.appendChild(flip);
    container.appendChild(wrap);

    return {
      el: wrap, flip: flip, img: img,
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
    p.img.classList.remove('walking', 'sitting', 'lying');
    if (p.zzz) { p.zzz.remove(); p.zzz = null; }
    if (state === 'walk') {
      p.img.classList.add('walking');
      p.timer = 4 + Math.random() * 5;
      if (Math.random() < 0.4) p.dir *= -1;
    } else if (state === 'sit') {
      p.img.classList.add('sitting');
      p.timer = 2 + Math.random() * 2.5;
    } else if (state === 'lie') {
      p.img.classList.add('lying');
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
