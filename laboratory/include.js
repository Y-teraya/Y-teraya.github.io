{
  let started = false;
  const PRES = /[?&]presenter/.test(location.search);
  if (PRES) document.documentElement.classList.add('presenter');
  const secs = [...document.querySelectorAll('.reveal .slides>section')];
  const hid = i => !!secs[i] && secs[i].classList.contains('hid');
  const visIdx = secs.map((s, i) => i).filter(i => !hid(i));
  const info = secs.map((s, i) => {
    const h = s.querySelector('h2') || s.querySelector('h1');
    const on = s.querySelector('.toc .on');
    const tg = s.querySelector('.top .tag');
    const nt = s.querySelector('aside.notes');
    return { i, hid: hid(i), ch: on ? on.textContent.trim() : '表紙',
      t: (tg ? tg.textContent.trim() + '　' : '') + (h ? h.textContent.replace(/\s+/g, ' ').trim() : '（表紙）'),
      notes: nt ? nt.textContent.trim() : '', txt: s.textContent.replace(/\s+/g, ' ') };
  });
  const nextOf = (i, d) => { let j = i + d; if (!hid(i)) { while (j >= 0 && j < secs.length && hid(j)) j += d; } return (j >= 0 && j < secs.length) ? j : -1; };

  /* ===================== 発表者ツール ===================== */
  if (PRES) {
    const op = window.opener;
    const send = m => { try { op && op.postMessage(Object.assign({ nod26: 1 }, m), '*'); } catch (_) {} };
    const pv = document.getElementById('pv');
    let cur = 0, ret = false, t0 = Date.now();
    const stage = (el, i) => {
      el.innerHTML = '';
      if (i < 0) { el.style.background = '#2a3034'; return; }
      el.style.background = '#fff';
      const w = el.clientWidth || 300;
      el.style.height = (w * 9 / 16) + 'px';
      const cv = document.createElement('div'); cv.className = 'cv'; cv.style.transform = 'scale(' + (w / 1280) + ')';
      const c = secs[i].cloneNode(true); c.classList.remove('present', 'past', 'future'); c.classList.add('pvs'); c.removeAttribute('style');
      c.querySelectorAll('aside.notes').forEach(n => n.remove());
      cv.appendChild(c); el.appendChild(cv);
    };
    const fitNow = () => {
      const now = document.getElementById('now');
      const col = now.parentElement; const avail = col.clientHeight - 90;
      const w = Math.min(col.clientWidth, avail * 16 / 9);
      now.style.width = w + 'px';
    };
    const render = () => {
      fitNow();
      stage(document.getElementById('now'), cur);
      const nx = nextOf(cur, 1);
      stage(document.getElementById('nxt'), nx);
      const k = visIdx.indexOf(cur);
      document.getElementById('nowLab').innerHTML = (cur + 1) + '. ' + info[cur].t + (hid(cur) ? ' <span class="hb">補足</span>' : '');
      document.getElementById('nxtLab').textContent = nx >= 0 ? (nx + 1) + '. ' + info[nx].t : '（最後）';
      document.getElementById('pNotes').textContent = info[cur].notes || '（原稿なし）';
      document.getElementById('pCount').textContent = k >= 0 ? '表示スライド ' + (k + 1) + ' / ' + visIdx.length : '補足スライドを表示中';
      pv.classList.toggle('jumped', ret);
    };
    const tick = () => {
      const s = Math.floor((Date.now() - t0) / 1000);
      const el = document.getElementById('tm');
      el.textContent = String(Math.floor(s / 60)).padStart(2, '0') + ':' + String(s % 60).padStart(2, '0');
      el.classList.toggle('over', s >= 900);
      const d = new Date(); document.getElementById('clock').textContent = '現在時刻 ' + d.getHours() + ':' + String(d.getMinutes()).padStart(2, '0');
    };
    setInterval(tick, 500); tick();
    document.getElementById('tReset').addEventListener('click', () => { t0 = Date.now(); tick(); });
    /* 一覧（サムネイル） */
    const grid = document.getElementById('grid'), gi = document.getElementById('gi'), gq = document.getElementById('gq'), gOnly = document.getElementById('gOnly');
    let built = false;
    const build = () => {
      let last = '', html = '';
      info.forEach(o => {
        if (o.ch !== last) { if (last) html += '</div>'; html += '<h4 data-ch="' + o.ch + '">' + o.ch + '</h4><div class="row2">'; last = o.ch; }
        html += '<div class="th" tabindex="0" data-i="' + o.i + '"><div class="stage" data-s="' + o.i + '"></div><div class="cap"><span class="no">' + (o.i + 1) + '</span><span>' + o.t + '</span>' + (o.hid ? '<span class="hb">補足</span>' : '') + '</div></div>';
      });
      gi.innerHTML = html + '</div>';
      gi.querySelectorAll('.stage[data-s]').forEach(el => stage(el, +el.dataset.s));
      built = true;
    };
    const filter = () => {
      const w = gq.value.trim();
      gi.querySelectorAll('.th').forEach(th => {
        const o = info[+th.dataset.i];
        const ok = (!gOnly.checked || o.hid) && (!w || o.txt.includes(w) || o.ch.includes(w));
        th.style.display = ok ? '' : 'none';
        th.classList.toggle('cur', o.i === cur);
      });
      gi.querySelectorAll('h4').forEach(h => {
        const box = h.nextElementSibling; const any = [...box.children].some(c => c.style.display !== 'none');
        h.style.display = box.style.display = any ? '' : 'none';
      });
    };
    const openGrid = () => { grid.classList.add('open'); if (!built) build(); filter(); gq.focus(); const c = gi.querySelector('.th.cur'); if (c) c.scrollIntoView({ block: 'center' }); };
    const closeGrid = () => grid.classList.remove('open');
    gq.addEventListener('input', filter); gOnly.addEventListener('change', filter);
    gi.addEventListener('click', e => { const th = e.target.closest('.th'); if (th) { send({ t: 'jump', i: +th.dataset.i }); closeGrid(); } });
    gi.addEventListener('keydown', e => { if (e.key === 'Enter') { const th = e.target.closest('.th'); if (th) { send({ t: 'jump', i: +th.dataset.i }); closeGrid(); } } });
    document.getElementById('gClose').addEventListener('click', closeGrid);
    document.getElementById('pGrid').addEventListener('click', openGrid);
    document.getElementById('pPrev').addEventListener('click', () => send({ t: 'step', d: -1 }));
    document.getElementById('pNext').addEventListener('click', () => send({ t: 'step', d: 1 }));
    document.getElementById('pBack').addEventListener('click', () => send({ t: 'back' }));
    addEventListener('keydown', e => {
      if (grid.classList.contains('open')) { if (e.key === 'Escape') { e.preventDefault(); closeGrid(); } return; }
      if (['ArrowRight','ArrowDown','PageDown',' ','Enter'].includes(e.key)) { e.preventDefault(); send({ t: 'step', d: 1 }); }
      else if (['ArrowLeft','ArrowUp','PageUp','Backspace'].includes(e.key)) { e.preventDefault(); send({ t: 'step', d: -1 }); }
      else if (e.key === 'l' || e.key === 'L') { e.preventDefault(); openGrid(); }
      else if (e.key === 'b' || e.key === 'B') { e.preventDefault(); send({ t: 'back' }); }
    });
    addEventListener('message', e => { const m = e.data; if (!m || !m.nod26 || m.t !== 'state') return; cur = m.i; ret = !!m.ret; render(); if (grid.classList.contains('open')) filter(); });
    addEventListener('resize', render);
    window.startDeck = () => {
      if (started) return; started = true;
      document.body.classList.remove('pre');
      if (!op) { document.getElementById('pNotes').textContent = 'スライドの画面で S キーを押して，発表者ツールを開いてください．'; }
      render(); send({ t: 'hello' });
    };
  } else {
  /* ===================== 観客用のスライド ===================== */
    let api = null, retIdx = null, pres = null;
    const cur = () => api.cur();
    const step = d => { const j = nextOf(cur(), d); if (j >= 0) api.go(j); };
    const broadcast = () => { try { if (pres && !pres.closed) pres.postMessage({ nod26: 1, t: 'state', i: cur(), ret: retIdx !== null }, '*'); } catch (_) {} };
    const update = () => {
      const i = cur(), k = visIdx.indexOf(i);
      document.querySelector('#pbar span').style.width = (k >= 0 ? k / (visIdx.length - 1) * 100 : 0) + '%';
      document.getElementById('pnum').textContent = k >= 0 ? (k + 1) + ' / ' + visIdx.length : '補足';
      broadcast();
    };
    const jump = j => { if (retIdx === null) retIdx = cur(); api.go(j); };
    const back = () => { if (retIdx === null) return; const r = retIdx; retIdx = null; api.go(r); };
    const openPresenter = () => {
      if (pres && !pres.closed) { pres.focus(); return; }
      pres = window.open(location.href.split(/[?#]/)[0] + '?presenter', 'nod26presenter', 'width=1400,height=850');
    };
    addEventListener('message', e => {
      const m = e.data; if (!m || !m.nod26 || !api) return;
      if (m.t === 'hello') { if (!pres || pres.closed) pres = e.source; broadcast(); }
      else if (m.t === 'step') step(m.d);
      else if (m.t === 'jump') jump(m.i);
      else if (m.t === 'back') back();
    });
    addEventListener('keydown', e => {
      if (e.key === 's' || e.key === 'S') { e.preventDefault(); e.stopImmediatePropagation(); openPresenter(); }
    }, true);

    /* reveal.js が読めないときの簡易スライド送り */
    const mini = () => {
      document.body.classList.add('mini');
      const box = document.querySelector('.reveal .slides');
      let i = 0;
      const fit = () => { const k = Math.min(innerWidth / 1280, innerHeight / 720) * 0.97; box.style.transform = 'translate(-50%,-50%) scale(' + k + ')'; };
      const show = () => { secs.forEach((s, j) => { s.classList.toggle('present', j === i); s.classList.toggle('past', j < i); s.classList.toggle('future', j > i); }); update(); };
      api = { cur: () => i, go: j => { i = Math.max(0, Math.min(secs.length - 1, j)); show(); } };
      addEventListener('keydown', e => {
        if (['ArrowRight','ArrowDown','PageDown',' ','Enter','n'].includes(e.key)) { e.preventDefault(); step(1); }
        else if (['ArrowLeft','ArrowUp','PageUp','Backspace','p'].includes(e.key)) { e.preventDefault(); step(-1); }
        else if (e.key === 'Home') api.go(visIdx[0]);
        else if (e.key === 'End') api.go(visIdx[visIdx.length - 1]);
        else if (e.key === 'f' || e.key === 'F') { try { document.documentElement.requestFullscreen(); } catch (_) {} }
      });
      let x0 = null;
      addEventListener('touchstart', e => { x0 = e.touches[0].clientX; }, {passive:true});
      addEventListener('touchend', e => { if (x0 === null) return; const dx = e.changedTouches[0].clientX - x0; if (Math.abs(dx) > 40) step(dx < 0 ? 1 : -1); x0 = null; });
      document.querySelector('.reveal').addEventListener('click', e => { if (e.clientX > innerWidth * 0.6) step(1); else if (e.clientX < innerWidth * 0.4) step(-1); });
      addEventListener('resize', fit);
      fit(); show();
      document.body.classList.remove('pre');
    };
    window.startDeck = () => {
      if (started) return; started = true;
      if (typeof Reveal === 'undefined') { mini(); return; }
      try {
        const nx = () => step(1), pv = () => step(-1);
        Reveal.initialize({width:1280,height:720,margin:0.03,center:false,hash:false,controls:false,progress:false,slideNumber:false,transition:'slide',touch:true,viewDistance:80,display:'block',
          keyboard:{39:nx,37:pv,40:nx,38:pv,32:nx,34:nx,33:pv,78:nx,80:pv,13:nx,83:null}});
        let moving = false, lastIdx = 0;
        api = { cur: () => Reveal.getIndices().h, go: j => { moving = true; Reveal.slide(j); moving = false; } };
        Reveal.on('ready', update);
        Reveal.on('slidechanged', e => {
          const i = e.indexh;
          if (!moving && hid(i) && !hid(lastIdx) && Math.abs(i - lastIdx) === 1) {
            const j = nextOf(lastIdx, i > lastIdx ? 1 : -1);
            if (j >= 0) { lastIdx = i; api.go(j); return; }
          }
          lastIdx = i; update();
        });
        document.body.classList.remove('pre');
      } catch (e) { mini(); }
    };
  }
  window.MathJax = {
    tex:{inlineMath:[['\\(','\\)']],displayMath:[['\\[','\\]']]},
    svg:{fontCache:'global'},
    startup:{pageReady(){ return MathJax.startup.defaultPageReady().then(window.startDeck, window.startDeck); }}
  };
  setTimeout(() => window.startDeck(), 4000);
}