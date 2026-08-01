/* 貝葉 — 한문 불전 원문·한국어 대조 열람기
   해시 라우팅만 사용한다(정적 호스팅에서 새로고침해도 깨지지 않음).
     #/                     문헌 목록
     #/w/<id>               대조 열람
     #/w/<id>/<위치표지>     해당 좌표로 이동
     #/s?q=...&scope=all    검색 결과
*/

const BASE = location.pathname.replace(/[^/]*$/, '');
const $ = (s, r = document) => r.querySelector(s);
const el = (t, c, txt) => { const n = document.createElement(t); if (c) n.className = c; if (txt != null) n.textContent = txt; return n; };
const esc = s => s.replace(/[&<>"]/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;' }[c]));
const MARKER = /\[\d{3,4}[abc]\d{2}\]/;
const APP_TAG = /\[(?:\d{1,3}|＊|\*)\]/g;

const state = {
  manifest: null,
  cache: new Map(),
  scope: 'all',
  facet: null,
  view: localStorage.getItem('view2') || 'split',  // split | stack  (기본: 좌우)
  show: localStorage.getItem('show2') || 'both',   // both | cn | ko (기본: 대조)
  fs: Number(localStorage.getItem('fs') || 100),   // 글자 크기 %
  token: 0,
};

const worker = new Worker(BASE + 'assets/js/search.worker.js', { type: 'classic' });
const pending = new Map();
worker.onmessage = e => {
  const { type, token, payload } = e.data;
  const h = pending.get(token);
  if (h && type === 'result') { h(payload); pending.delete(token); }
};
function search(query, ids, scope) {
  const token = ++state.token;
  return new Promise(res => {
    pending.set(token, res);
    worker.postMessage({ type:'search', token, payload:{ query, ids, base:BASE, scope, limit:300 } });
  });
}

/* ─── 데이터 ─────────────────────────────────────── */
async function manifest() {
  if (!state.manifest) {
    state.manifest = await (await fetch(BASE + 'data/manifest.json')).json();
  }
  return state.manifest;
}
async function work(id) {
  if (!state.cache.has(id)) {
    state.cache.set(id, await (await fetch(`${BASE}data/works/${id}.json`)).json());
  }
  return state.cache.get(id);
}

/* ─── 렌더 조각 ──────────────────────────────────── */
function markupCN(text) {
  // 위치표지와 교감 표지를 본문 글자와 구분해 보여 준다
  return esc(text)
    .replace(/^(\[\d{3,4}[abc]\d{2}\])\s*/gm, '<span class="mk">$1</span>')
    .replace(APP_TAG, m => `<span class="app">${m}</span>`);
}

function unitNode(u, wid) {
  const n = el('article', 'unit');
  n.id = 'u' + u.i;
  if (u.m) n.dataset.m = u.m;

  const rail = el('div', 'rail');
  if (u.m) {
    const b = el('button', null, u.m);
    b.title = '이 자리의 주소 복사';
    b.onclick = (ev) => copyAnchor(wid, u, ev);
    rail.append(b);
  } else {
    n.classList.add('nomark');
  }

  const cnbox = el('div', 'cnbox');
  const p = el('p', 'cn');
  p.innerHTML = markupCN(u.cn.join('\n'));
  cnbox.append(p);

  const kobox = el('div', 'kobox');
  if (u.ko && u.ko.length) {
    u.ko.forEach(t => kobox.append(el('p', 'ko', t)));
  } else {
    kobox.append(el('p', 'nokr', '번역 대응 없음'));
  }

  n.append(rail, cnbox, kobox);

  if (u.nt && u.nt.length) {
    const nt = el('div', 'notes');
    u.nt.forEach(t => nt.append(el('p', null, t)));
    n.append(nt);
  }
  return n;
}

function copyAnchor(wid, u, ev) {
  const url = location.origin + location.pathname + `#/w/${wid}/${u.m || 'u' + u.i}`;
  navigator.clipboard?.writeText(url);
  const tip = $('#railtip');
  tip.textContent = '주소를 복사했습니다';
  tip.hidden = false;
  const r = ev.currentTarget.getBoundingClientRect();
  tip.style.left = r.left + r.width / 2 + 'px';
  tip.style.top = r.top + 'px';
  clearTimeout(copyAnchor.t);
  copyAnchor.t = setTimeout(() => tip.hidden = true, 1400);
}

/* ─── 사이드바 ───────────────────────────────────── */
async function renderSide(active) {
  const m = await manifest();
  const ul = $('#workList'); ul.innerHTML = '';
  $('#sideCount').textContent = `${m.works.length}종`;

  let lastCanon = null;
  m.works.forEach(w => {
    // 대장경이 바뀌면 구분선을 넣는다 (대정장 → 속장경 → 만자속장)
    if (w.canon_id !== lastCanon) {
      lastCanon = w.canon_id;
      const h = el('li', 'wl-group', CANON_NAME[w.canon_id] || w.canon_id);
      ul.append(h);
    }
    const li = el('li');
    const a = el('a', 'wl' + (w.id === active ? ' on' : ''));
    a.href = `#/w/${w.id}`;
    a.append(el('div', 'wl-cn', w.title_cn));
    a.append(el('div', 'wl-ko', w.title_ko));

    const src = el('div', 'wl-src');
    src.append(el('span', 'wl-canon', w.canon_label || w.canon));
    src.append(el('span', null, `${w.dynasty} ${w.author_ko}`));
    a.append(src);

    const meta = el('div', 'wl-meta');
    const bar = el('span', 'bar'); const i = el('i');
    i.style.width = Math.round(w.coverage * 100) + '%';
    if (!w.has_translation) i.style.background = 'var(--mist)';
    bar.append(i); meta.append(bar);
    meta.append(el('span', 'wl-pct',
      w.has_translation ? Math.round(w.coverage * 100) + '%' : '원문만'));
    a.append(meta);

    li.append(a); ul.append(li);
  });

  $('#sideFoot').innerHTML =
    `원문 ${m.totals.chars_cn.toLocaleString()}자<br>단위 ${m.totals.units.toLocaleString()}개`;
}

const CANON_NAME = {
  T: '大正新脩大藏經 · 대정신수대장경',
  X: '卍新纂續藏經 · 만신찬속장경',
  L: '乾隆大藏經 · 건륭대장경',
  K: '高麗大藏經 · 고려대장경',
};

/* ─── 화면: 문헌 목록 ────────────────────────────── */
async function viewHome() {
  const m = await manifest();
  const main = $('#main'); main.innerHTML = '';

  const hero = el('div', 'hero');
  hero.innerHTML = `
    <h1>원문과 번역을 같은 좌표에서 읽습니다</h1>
    <p>한문 불전의 원문을 대장경 위치표지 단위로 나누고, 그 자리에 한국어 직역을 붙였습니다.
       한 낱말을 넣으면 원문과 번역을 함께 훑어 그 대목으로 데려다 놓습니다.</p>
    <div class="stat">
      <div><b>${m.totals.works}</b><span>수록 문헌</span></div>
      <div><b>${m.totals.units.toLocaleString()}</b><span>대조 단위</span></div>
      <div><b>${m.totals.chars_cn.toLocaleString()}</b><span>원문 글자</span></div>
      <div><b>${Math.round(m.totals.translated / m.totals.units * 100)}%</b><span>번역 대응</span></div>
    </div>`;
  main.append(hero);

  const grid = el('div', 'grid');
  m.works.forEach(w => {
    const a = el('a', 'card'); a.href = `#/w/${w.id}`;
    a.append(el('h3', null, w.title_cn));
    a.append(el('div', 'ko', w.title_ko));
    a.append(el('div', 'ko2', `${w.dynasty} ${w.author_ko} · ${w.author_cn}`));
    const row = el('div', 'row');
    row.append(el('span', 'wl-canon', w.canon_label || w.canon));
    row.append(el('span', null, `${w.units.toLocaleString()}단위`));
    a.append(row);
    const tl = el('div', 'tagline');
    if (!w.has_translation) tl.append(el('span', 'tag warn', '원문만'));
    else tl.append(el('span', 'tag', `번역 ${Math.round(w.coverage * 100)}%`));
    (w.tags || []).slice(0, 2).forEach(t => tl.append(el('span', 'tag', t)));
    a.append(tl);
    grid.append(a);
  });
  main.append(grid);
  renderSide(null);
}

/* ─── 화면: 대조 열람 ────────────────────────────── */
async function viewWork(id, anchor) {
  const main = $('#main');
  main.innerHTML = '<p class="loading">문헌을 여는 중…</p>';
  renderSide(id);

  const doc = await work(id);
  const w = doc.meta;
  main.innerHTML = '';

  const bar = el('div', 'docbar');
  bar.innerHTML = `<h2>${esc(w.title_cn)}</h2>
    <span class="pos" id="posNow">${w.range ? w.range[0] : ''}</span>
    <span class="spacer"></span>`;
  const seg1 = el('div', 'seg');
  [['both', '대조'], ['cn', '원문'], ['ko', '번역']].forEach(([k, t]) => {
    const b = el('button', state.show === k ? 'on' : '', t);
    b.onclick = () => { state.show = k; localStorage.setItem('show2', k); applyModes(); [...seg1.children].forEach(c => c.classList.remove('on')); b.classList.add('on'); };
    seg1.append(b);
  });
  const seg2 = el('div', 'seg');
  [['stack', '위아래'], ['split', '좌우']].forEach(([k, t]) => {
    const b = el('button', state.view === k ? 'on' : '', t);
    b.onclick = () => { state.view = k; localStorage.setItem('view2', k); applyModes(); [...seg2.children].forEach(c => c.classList.remove('on')); b.classList.add('on'); };
    seg2.append(b);
  });
  if (w.has_translation) {
    const ai = el('span', 'aitag', 'AI 번역');
    ai.title = '이 문헌의 한국어 번역은 AI가 생성했습니다. 인용 전 원문 대조가 필요합니다.';
    bar.append(ai);
  }
  const fsbox = el('div', 'fsbox');
  const minus = el('button', 'fsbtn', '가−');
  minus.title = '글자 작게';
  minus.onclick = () => setFont(state.fs - FS_STEP);
  const now = el('span', 'fsnow'); now.id = 'fsNow';
  const plus = el('button', 'fsbtn', '가＋');
  plus.title = '글자 크게';
  plus.onclick = () => setFont(state.fs + FS_STEP);
  now.onclick = () => setFont(100);
  now.title = '기본 크기로';
  fsbox.append(minus, now, plus);

  bar.append(seg1, seg2, fsbox);
  main.append(bar);
  applyFont();

  const wrap = el('div', 'wrap');

  // 서지
  const info = el('details', 'biblio');
  info.innerHTML = `<summary>서지 · 용어표 · 해제</summary>`;
  const drawer = el('div', 'drawer');
  const tb = el('table', 'meta-table');
  const rows = [
    ['원제', `<span class="cnw">${esc(w.title_cn)}</span>`],
    ['국역명', esc(w.title_ko)],
    ['찬자', `<span class="cnw">${esc(w.author_cn)}</span> · ${esc(w.author_ko)} (${esc(w.dynasty)})`],
    ['대장경', `<span class="cnw">${esc(w.canon)}</span>` +
      (w.canon_ko ? ` <span style="color:var(--ink-3)">— ${esc(w.canon_ko)}</span>` : '') +
      (w.verify ? ' <span class="tag warn">서지 확인 필요</span>' : '')],
    ...(w.note ? [['비고', esc(w.note)]] : []),
    ['수록 범위', w.range ? `${w.range[0]} – ${w.range[1]}` : '—'],
    ['대조 단위', `${w.units.toLocaleString()}개 · 번역 대응 ${Math.round(w.coverage * 100)}%`],
    ['원문 글자', w.chars_cn.toLocaleString() + '자'],
    ['원문 출처', `${esc(w.source || 'CBETA')} 電子佛典集成` +
      (w['底本_publisher'] ? ` · 저본 판권자 <span class="cnw">${esc(w['底本_publisher'])}</span> ©` : '') +
      ' · <a href="#/rights">CC BY-NC-SA 4.0</a>'],
  ];
  tb.innerHTML = rows.map(([k, v]) => `<tr><th>${k}</th><td>${v}</td></tr>`).join('');
  info.append(tb);
  if (doc.glossary?.length) {
    info.append(el('h4', 'sect', '용어 대응표'));
    const g = el('table', 'meta-table');
    g.innerHTML = doc.glossary.map(t =>
      `<tr><th><span class="cnw">${esc(t.cn)}</span></th><td>${esc(t.ko)}${t.note ? ` <span style="color:var(--ink-3)">— ${esc(t.note)}</span>` : ''}</td></tr>`
    ).join('');
    info.append(g);
  }
  if (doc.front?.length) {
    info.append(el('h4', 'sect', '해제 · 편집 안내'));
    const fr = el('div', 'frontnote');
    doc.front.forEach(t => fr.append(el('p', null, t)));
    info.append(fr);
  }
  drawer.append(info);
  bar.append(drawer);

  const chapters = doc.chapters || [];
  const nav = el('div', 'chapnav');
  const body = el('div', 'units');
  if (chapters.length > 1) bar.append(nav);   // 문서 바에 얹어 함께 고정
  wrap.append(body);
  main.append(wrap);

  // 절 표제(권보다 아래 층위)를 단위 사이에 끼워 넣기 위한 지도
  const chapAt = new Set(chapters.map(c => c.i));
  const subAt = new Map();
  (doc.sections || []).forEach(s => {
    if (!chapAt.has(s.i)) subAt.set(s.i, s.t);
  });

  let ci = 0;
  if (anchor) {
    const k = doc.units.findIndex(u =>
      anchor.startsWith('u') ? ('u' + u.i) === anchor : u.m === anchor);
    if (k >= 0) ci = doc.units[k].c || 0;
  }

  function renderChapter(n, jump) {
    ci = Math.max(0, Math.min(n, Math.max(chapters.length - 1, 0)));
    const from = chapters.length ? chapters[ci].i : 0;
    const to = chapters.length && ci + 1 < chapters.length
      ? chapters[ci + 1].i : doc.units.length;

    body.innerHTML = '';
    const frag = document.createDocumentFragment();
    for (let k = from; k < to; k++) {
      const sub = subAt.get(k);
      if (sub) frag.append(el('h3', 'subhead', sub));
      frag.append(unitNode(doc.units[k], id));
    }
    body.append(frag);
    drawNav();
    applyModes();
    watchPosition();
    if (jump) jumpTo(jump);
    else if (chapters.length > 1) main.scrollIntoView({ block: 'start' });
  }

  function drawNav() {
    if (chapters.length < 2) return;
    nav.innerHTML = '';
    const prev = el('button', 'chapbtn', '← 앞 권');
    prev.disabled = ci === 0;
    prev.onclick = () => renderChapter(ci - 1);

    const sel = el('select', 'chapsel');
    chapters.forEach((c, n) => {
      const to = n + 1 < chapters.length ? chapters[n + 1].i : doc.units.length;
      const o = el('option', null, `${c.t}  (${to - c.i}단위)`);
      o.value = String(n);
      if (n === ci) o.selected = true;
      sel.append(o);
    });
    sel.onchange = () => renderChapter(Number(sel.value));

    const next = el('button', 'chapbtn', '뒤 권 →');
    next.disabled = ci >= chapters.length - 1;
    next.onclick = () => renderChapter(ci + 1);

    const count = el('span', 'chapcount', `${ci + 1} / ${chapters.length}`);
    nav.append(prev, sel, next, count);
  }

  // 문서 바 높이가 바뀌면(서지 서랍 여닫기, 줄바꿈) 본문 여백을 맞춘다
  const syncBar = () => {
    const h = bar.getBoundingClientRect().height;
    document.documentElement.style.setProperty('--barh', h + 'px');
  };
  info.addEventListener('toggle', syncBar);
  if (window.ResizeObserver) new ResizeObserver(syncBar).observe(bar);
  syncBar();

  renderChapter(ci, anchor);
}

const FS_MIN = 80, FS_MAX = 160, FS_STEP = 10;
function applyFont() {
  document.documentElement.style.setProperty('--fs', state.fs / 100);
  const o = document.getElementById('fsNow');
  if (o) o.textContent = state.fs + '%';
}
applyFont();

function applyModes() {
  $('#main').classList.toggle('split', state.view === 'split');
  document.querySelector('.wrap')?.classList.toggle('wide', state.view === 'split');
  document.body.classList.toggle('hide-ko', state.show === 'cn');
  document.body.classList.toggle('hide-cn', state.show === 'ko');
}

function setFont(v) {
  state.fs = Math.max(FS_MIN, Math.min(FS_MAX, v));
  localStorage.setItem('fs', String(state.fs));
  applyFont();
}

function jumpTo(anchor) {
  const node = anchor.startsWith('u')
    ? document.getElementById(anchor)
    : $(`.unit[data-m="${CSS.escape(anchor)}"]`);
  if (!node) return;
  node.scrollIntoView({ block: 'center', behavior: 'smooth' });
  node.classList.add('flash');
  setTimeout(() => node.classList.remove('flash'), 2600);
}

let posObserver;
function watchPosition() {
  posObserver?.disconnect();
  const out = $('#posNow'); if (!out) return;
  posObserver = new IntersectionObserver(es => {
    for (const e of es) if (e.isIntersecting && e.target.dataset.m) { out.textContent = e.target.dataset.m; break; }
  }, { rootMargin: '-20% 0px -70% 0px' });
  document.querySelectorAll('.unit').forEach(n => posObserver.observe(n));
}

/* ─── 화면: 검색 ─────────────────────────────────── */
async function viewSearch(q, scope) {
  const m = await manifest();
  const main = $('#main'); main.innerHTML = '';
  renderSide(null);
  $('#q').value = q;

  const head = el('div', 'reshead');
  head.innerHTML = `<h2>‘${esc(q)}’ 찾기</h2><span class="n" id="resN">훑는 중…</span>`;
  main.append(head);

  const facets = el('div', 'facets'); main.append(facets);
  const list = el('div', 'hits'); main.append(list);

  const ids = m.works.map(w => w.id);
  const { hits, counts } = await search(q, ids, scope);

  const total = Object.values(counts).reduce((a, b) => a + b, 0);
  $('#resN').textContent = total ? `${total.toLocaleString()}곳` : '';

  if (!total) {
    list.innerHTML = `<div class="empty"><b>찾은 곳이 없습니다</b>
      다른 표기로 바꿔 보세요. 원문은 이체자를 자동으로 함께 찾습니다(眞·真, 衆·眾).</div>`;
    return;
  }

  const mkFacet = (label, id) => {
    const b = el('button', 'facet' + (state.facet === id ? ' on' : ''), label);
    b.onclick = () => { state.facet = state.facet === id ? null : id; draw(); [...facets.children].forEach(c => c.classList.remove('on')); if (state.facet) b.classList.add('on'); };
    return b;
  };
  facets.append(mkFacet(`전체 ${total}`, null));
  m.works.filter(w => counts[w.id]).forEach(w =>
    facets.append(mkFacet(`${w.title_ko} ${counts[w.id]}`, w.id)));

  function draw() {
    list.innerHTML = '';
    const shown = hits.filter(h => !state.facet || h.work === state.facet);
    shown.forEach(h => {
      const a = el('a', 'hit');
      a.href = `#/w/${h.work}/${h.m || 'u' + h.i}`;
      const top = el('div', 'hit-top');
      const wk = m.works.find(x => x.id === h.work);
      top.innerHTML = `<b>${esc(h.title || h.title_cn || '')}</b>
        <span>${esc(wk ? wk.title_ko : '')}</span>
        <span class="m">${esc(h.m || '—')}</span>
        <span>${h.where === 'cn' ? '원문' : '번역'}</span>`;
      a.append(top);
      const p = el('p', h.where === 'cn' ? 'hit-cn' : 'hit-ko');
      p.innerHTML = h.snip
        ? `${esc(h.snip.pre)}<mark>${esc(h.snip.hit)}</mark>${esc(h.snip.post)}`
        : '';
      a.append(p);
      if (h.alt) {
        const q2 = el('p', h.where === 'cn' ? 'hit-ko' : 'hit-cn');
        q2.textContent = h.alt;
        a.append(q2);
      }
      list.append(a);
    });
    if (hits.length >= 300) {
      list.append(el('p', 'loading', '앞의 300곳만 보여 줍니다. 낱말을 더 좁혀 보세요.'));
    }
  }
  draw();
}

async function viewRights() {
  const m = await manifest();
  const R = m.rights || {};
  const main = $('#main'); main.innerHTML = '';
  renderSide(null);

  const w = el('div', 'wrap rights-page');
  w.innerHTML = `
    <h1 class="rights-h">출처와 이용 조건</h1>

    <section>
      <h2>한문 원문</h2>
      <p>이 사이트의 한문 원문은 <b>${esc(R.source_db || 'CBETA 電子佛典集成')}</b>
         (${esc(R.source_org || '')})에서 가져왔습니다.</p>
      <p>CBETA 版權宣告에 따르면 대정신수대장경·만신찬속장경·역대장경보집 등은
         <b>${esc(R.license || 'CC BY-NC-SA 4.0')}</b>
         (저작자표시–비영리–동일조건변경허락) 조건으로 공개되어 있습니다.
         이 사이트는 그 조건을 따릅니다.</p>
      <ul class="rights-list">
        <li><b>저작자표시</b> — 원문의 출처가 CBETA임을 문헌마다 밝힙니다.</li>
        <li><b>비영리</b> — 이 사이트는 영리 목적으로 운영하지 않으며, 광고를 싣지 않습니다.</li>
        <li><b>동일조건변경허락</b> — 이 사이트의 가공물과 한국어 번역도 같은 조건으로 공개합니다.</li>
        <li><b>실질 내용 무변경</b> — 원문의 글자와 교감 표지를 임의로 고치지 않았습니다.</li>
      </ul>
      <p class="rights-note">저본 판권자: 대정신수대장경 大藏出版株式會社 ©,
         만신찬속장경 株式會社國書刊行會 ©, 건륭대장경 新文豐出版公司.<br>
         CBETA 판본: ${esc(R.cbeta_version || '확인 필요')}<br>
         원문 이용 조건 전문은 <a href="${esc(R.source_url || 'https://cbeta.org/copyright')}"
         target="_blank" rel="noopener">CBETA 版權宣告</a>을 보십시오.</p>
      <p class="rights-warn">CBETA 수록 문헌 가운데 인순법사·여징·태허·연배 등
         근현대 저자의 저작집(類別 B)은 CC 조건이 적용되지 않습니다.
         이 사이트는 해당 문헌을 수록하지 않습니다.</p>
    </section>

    <section>
      <h2>한국어 번역</h2>
      <p>번역문은 생성형 AI로 만든 뒤 검토한 것으로, 한문 원문의 2차적 저작물입니다.
         원문과 같은 <b>${esc(R.license || 'CC BY-NC-SA 4.0')}</b> 조건으로 공개합니다.</p>
      <p class="rights-warn">직역을 원칙으로 삼았으나 오역·누락·과잉 해석이 있을 수 있습니다.
         학술적 인용에 앞서 반드시 한문 원문과 대조하여 직접 검토하십시오.
         판본 대조는 별도로 수행하지 않았습니다.</p>
    </section>

    <section>
      <h2>이 사이트를 인용할 때</h2>
      <p class="rights-cite">한민수 편, 『신한글대장경』, 동명대학교.
         원문 출처: CBETA 電子佛典集成 (CC BY-NC-SA 4.0). 열람일: <span id="today"></span>.</p>
      <p>문헌 단위를 지목하려면 열람 화면 왼쪽의 위치표지를 누르십시오.
         해당 대목의 주소가 복사됩니다.</p>
    </section>

    <section>
      <h2>문의와 정정</h2>
      <p>원문 입력 오류, 번역 오류, 서지 오류를 발견하시면 알려 주십시오.
         원문 자체의 오류는 CBETA에 직접 알리는 편이 빠릅니다.</p>
      <p class="rights-note">한민수 · 동명대학교</p>
    </section>`;
  main.append(w);
  const t = new Date();
  const d = w.querySelector('#today');
  if (d) d.textContent = `${t.getFullYear()}. ${t.getMonth() + 1}. ${t.getDate()}.`;
}

/* ─── 라우터 ─────────────────────────────────────── */
function route() {
  const h = location.hash.slice(1) || '/';
  const [path, qs] = h.split('?');
  const params = new URLSearchParams(qs || '');
  const seg = path.split('/').filter(Boolean);
  $('#sidebar').classList.remove('open');
  document.querySelectorAll('.tabbar a,.tabbar button').forEach(b => b.classList.remove('on'));

  if (seg[0] === 'w' && seg[1]) {
    document.querySelector('[data-tab="view"]')?.classList.add('on');
    return viewWork(seg[1], seg[2]);
  }
  if (seg[0] === 'rights') {
    return viewRights();
  }
  if (seg[0] === 's') {
    document.querySelector('[data-tab="search"]')?.classList.add('on');
    const q = params.get('q') || '';
    state.scope = params.get('scope') || 'all';
    syncScope();
    return q ? viewSearch(q, state.scope) : viewHome();
  }
  document.querySelector('[data-tab="home"]')?.classList.add('on');
  viewHome();
}

function syncScope() {
  document.querySelectorAll('#scopeSeg button').forEach(b =>
    b.classList.toggle('on', b.dataset.scope === state.scope));
}

/* ─── 이벤트 ─────────────────────────────────────── */
$('#searchForm').addEventListener('submit', e => {
  e.preventDefault();
  const q = $('#q').value.trim();
  if (!q) return;
  state.facet = null;
  location.hash = `#/s?q=${encodeURIComponent(q)}&scope=${state.scope}`;
});
$('#scopeSeg').addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b) return;
  state.scope = b.dataset.scope; syncScope();
  const q = $('#q').value.trim();
  if (q) location.hash = `#/s?q=${encodeURIComponent(q)}&scope=${state.scope}`;
});
document.querySelector('[data-tab="search"]').addEventListener('click', () => $('#q').focus());
document.querySelector('[data-tab="view"]').addEventListener('click', () => {
  $('#sidebar').classList.toggle('open');
});
document.addEventListener('keydown', e => {
  if (e.key === '/' && document.activeElement !== $('#q')) { e.preventDefault(); $('#q').focus(); }
});

window.addEventListener('hashchange', route);
route();

// 오프라인 열람
if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
  navigator.serviceWorker.register(BASE + 'sw.js').catch(() => {});
}
