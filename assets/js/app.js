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
  variants: true,   // 이체자를 접어 함께 찾을지
  facet: null,
  view: localStorage.getItem('view2') || 'split',  // split | stack  (기본: 좌우)
  show: localStorage.getItem('show2') || 'both',   // both | cn | ko (기본: 대조)
  fs: Number(localStorage.getItem('fs') || 100),   // 글자 크기 %
  perPage: Number(localStorage.getItem('perPage') || 10),  // 검색 결과 표시 개수
  token: 0,
};

const worker = new Worker(BASE + 'assets/js/search.worker.js', { type: 'classic' });
const pending = new Map();
worker.onmessage = e => {
  const { type, token, payload } = e.data;
  const h = pending.get(token);
  if (h && type === 'result') { h(payload); pending.delete(token); }
};
function search(query, ids, scope, variants = true) {
  const token = ++state.token;
  return new Promise(res => {
    pending.set(token, res);
    worker.postMessage({ type:'search', token, payload:{ query, ids, base:BASE, scope, variants, limit:0 } });
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
/* ─── 검색어 강조 ─────────────────────────────────
   검색과 같은 기준(이체자 접기·구두점 무시)으로 자리를 찾은 뒤,
   원문 글자 위치로 되돌려 표시한다. 그래야 "一心。" 처럼 사이에
   구두점이 낀 경우에도 제대로 잡힌다. */
const VAR_FOLD = {
  '万':'萬',
  '与':'與',
  '両':'兩',
  '两':'兩',
  '严':'嚴',
  '个':'個',
  '为':'為',
  '么':'麼',
  '义':'義',
  '乐':'樂',
  '亀':'龜',
  '于':'於',
  '亙':'亘',
  '从':'從',
  '众':'眾',
  '会':'會',
  '体':'體',
  '兎':'兔',
  '兪':'俞',
  '关':'關',
  '円':'圓',
  '冊':'册',
  '写':'寫',
  '冩':'寫',
  '冲':'沖',
  '决':'決',
  '况':'況',
  '净':'淨',
  '凈':'淨',
  '凉':'涼',
  '処':'處',
  '刹':'剎',
  '剏':'創',
  '剑':'劍',
  '剣':'劍',
  '劒':'劍',
  '勅':'敕',
  '勑':'敕',
  '医':'醫',
  '华':'華',
  '卽':'即',
  '厳':'嚴',
  '収':'收',
  '发':'發',
  '变':'變',
  '号':'號',
  '噐':'器',
  '国':'國',
  '増':'增',
  '壊':'壞',
  '壌':'壤',
  '声':'聲',
  '売':'賣',
  '处':'處',
  '変':'變',
  '妬':'妒',
  '姊':'姉',
  '学':'學',
  '实':'實',
  '実':'實',
  '対':'對',
  '寿':'壽',
  '尽':'盡',
  '峯':'峰',
  '嶋':'島',
  '嶽':'岳',
  '巖':'岩',
  '巻':'卷',
  '师':'師',
  '带':'帶',
  '帯':'帶',
  '帰':'歸',
  '幷':'并',
  '广':'廣',
  '広':'廣',
  '庄':'莊',
  '应':'應',
  '廻':'迴',
  '廼':'迺',
  '弃':'棄',
  '弥':'彌',
  '归':'歸',
  '当':'當',
  '徧':'遍',
  '徳':'德',
  '徴':'徵',
  '怜':'憐',
  '恆':'恒',
  '恶':'惡',
  '恼':'惱',
  '悦':'悅',
  '悩':'惱',
  '悪':'惡',
  '慙':'慚',
  '慜':'愍',
  '战':'戰',
  '戦':'戰',
  '户':'戶',
  '戸':'戶',
  '担':'擔',
  '拝':'拜',
  '挿':'插',
  '掛':'挂',
  '摂':'攝',
  '摄':'攝',
  '摠':'總',
  '擡':'抬',
  '擧':'舉',
  '敍':'敘',
  '敎':'教',
  '数':'數',
  '斉':'齊',
  '斎':'齋',
  '断':'斷',
  '旛':'幡',
  '无':'無',
  '旣':'既',
  '旧':'舊',
  '昬':'昏',
  '昼':'晝',
  '显':'顯',
  '朶':'朵',
  '杂':'雜',
  '来':'來',
  '栄':'榮',
  '栢':'柏',
  '栰':'筏',
  '棱':'稜',
  '検':'檢',
  '楽':'樂',
  '歎':'嘆',
  '歩':'步',
  '歯':'齒',
  '歳':'歲',
  '歴':'歷',
  '殻':'殼',
  '毎':'每',
  '气':'氣',
  '気':'氣',
  '氷':'冰',
  '没':'沒',
  '涙':'淚',
  '渇':'渴',
  '渉':'涉',
  '温':'溫',
  '湼':'涅',
  '潅':'灌',
  '灋':'法',
  '灯':'燈',
  '灵':'靈',
  '点':'點',
  '烦':'煩',
  '煖':'暖',
  '燄':'焰',
  '燐':'憐',
  '爲':'為',
  '牀':'床',
  '独':'獨',
  '现':'現',
  '瑠':'琉',
  '瑯':'琅',
  '璢':'琉',
  '甁':'瓶',
  '甎':'磚',
  '甞':'嘗',
  '畞':'畝',
  '畧':'略',
  '疊':'叠',
  '疎':'疏',
  '疣':'肬',
  '発':'發',
  '皁':'皂',
  '皃':'貌',
  '盋':'缽',
  '盖':'蓋',
  '盗':'盜',
  '眞':'真',
  '睠':'眷',
  '砕':'碎',
  '碍':'礙',
  '礼':'禮',
  '祕':'秘',
  '祗':'祇',
  '禄':'祿',
  '禅':'禪',
  '禱':'祷',
  '称':'稱',
  '稟':'禀',
  '稲':'稻',
  '窃':'竊',
  '窗':'窓',
  '竜':'龍',
  '竝':'並',
  '竪':'豎',
  '笋':'筍',
  '笼':'籠',
  '筯':'箸',
  '篭':'籠',
  '籖':'籤',
  '粛':'肅',
  '経':'經',
  '絶':'絕',
  '継':'繼',
  '綵':'彩',
  '緑':'綠',
  '縁':'緣',
  '縦':'縱',
  '繊':'纖',
  '繋':'繫',
  '纤':'纖',
  '经':'經',
  '继':'繼',
  '罸':'罰',
  '羣':'群',
  '翛':'倏',
  '翫':'玩',
  '聟':'壻',
  '聡':'聰',
  '聨':'聯',
  '聪':'聰',
  '肃':'肅',
  '脇':'脅',
  '脱':'脫',
  '臈':'臘',
  '舎':'舍',
  '舘':'館',
  '舩':'船',
  '荅':'答',
  '荘':'莊',
  '荣':'榮',
  '药':'藥',
  '莭':'節',
  '莵':'菟',
  '菴':'庵',
  '萨':'薩',
  '葢':'蓋',
  '蔵':'藏',
  '蕐':'華',
  '薬':'藥',
  '薰':'熏',
  '蘂':'蕊',
  '蘓':'蘇',
  '虗':'虛',
  '虚':'虛',
  '蛮':'蠻',
  '衆':'眾',
  '衞':'衛',
  '裏':'裡',
  '裵':'裴',
  '覈':'核',
  '覔':'覓',
  '覚':'覺',
  '覩':'睹',
  '観':'觀',
  '观':'觀',
  '觉':'覺',
  '触':'觸',
  '訶':'呵',
  '註':'注',
  '証':'證',
  '詶':'酬',
  '説':'說',
  '読':'讀',
  '諠':'喧',
  '謌':'歌',
  '讁':'謫',
  '讎':'讐',
  '论':'論',
  '证':'證',
  '说':'說',
  '读':'讀',
  '谿':'溪',
  '貍':'狸',
  '貮':'貳',
  '賔':'賓',
  '賖':'賒',
  '賛':'贊',
  '賸':'剩',
  '贒':'賢',
  '赖':'賴',
  '踈':'疏',
  '踪':'蹤',
  '踰':'逾',
  '踴':'踊',
  '蹟':'跡',
  '躰':'體',
  '軆':'體',
  '軽':'輕',
  '輙':'輒',
  '转':'轉',
  '轻':'輕',
  '辞':'辭',
  '边':'邊',
  '迹':'跡',
  '逈':'迥',
  '逹':'達',
  '遯':'遁',
  '遲':'遅',
  '遶':'繞',
  '邨':'村',
  '郞':'郎',
  '鄕':'鄉',
  '酧':'酬',
  '醎':'鹹',
  '釈':'釋',
  '释':'釋',
  '釼':'劍',
  '鈆':'鉛',
  '鉄':'鐵',
  '鉢':'缽',
  '銕':'鐵',
  '鍊':'煉',
  '鎻':'鎖',
  '鑒':'鑑',
  '铁':'鐵',
  '閒':'間',
  '閙':'鬧',
  '関':'關',
  '阯':'址',
  '陜':'陝',
  '随':'隨',
  '隷':'隸',
  '雑':'雜',
  '霊':'靈',
  '靑':'青',
  '静':'靜',
  '韈':'襪',
  '頬':'頰',
  '頺':'頹',
  '頼':'賴',
  '顋':'腮',
  '顚':'顛',
  '飜':'翻',
  '飡':'餐',
  '飮':'飲',
  '餝':'飾',
  '餧':'餵',
  '饍':'膳',
  '駆':'驅',
  '騐':'驗',
  '験':'驗',
  '驱':'驅',
  '髄':'髓',
  '髣':'仿',
  '髴':'彿',
  '鬂':'鬢',
  '鬪':'鬥',
  '鵞':'鵝',
  '麤':'麁',
  '麪':'麵',
  '麽':'麼',
  '黄':'黃',
  '黒':'黑',
  '黙':'默',
  '鼔':'鼓',
  '齐':'齊',
  '齢':'齡',
  '齿':'齒',
  '龙':'龍',
  '龟':'龜'
};
const DROP_CH = /[\s。，、．・？！：；「」『』（）()〔〕【】\[\]“”‘’·…—　]/;

function foldCh(ch) { return state.variants ? (VAR_FOLD[ch] || ch) : ch; }

function normQuery(qs, kind) {
  let out = '';
  for (const ch of qs.normalize('NFKC')) {
    if (DROP_CH.test(ch)) continue;
    out += kind === 'cn' ? foldCh(ch) : ch.toLowerCase();
  }
  return out;
}

/** 원문 문자열에서 검색어가 놓인 구간들을 찾는다 → [[시작, 끝], …] */
function findRanges(raw, kind, needle) {
  if (!needle) return [];
  const skip = new Uint8Array(raw.length);
  if (kind === 'cn') {
    for (const re of [/\[\d{3,4}[abc]\d{2}\]/g, /\[(?:\d{1,3}|＊|\*)\]/g]) {
      let m;
      while ((m = re.exec(raw)) !== null) {
        for (let i = m.index; i < m.index + m[0].length; i++) skip[i] = 1;
      }
    }
  }
  let norm = '';
  const map = [];
  for (let i = 0; i < raw.length; i++) {
    if (skip[i]) continue;
    const ch = raw[i].normalize('NFKC');
    if (DROP_CH.test(ch)) continue;
    const n = kind === 'cn' ? foldCh(ch) : ch.toLowerCase();
    for (let k = 0; k < n.length; k++) map.push(i);
    norm += n;
  }
  const out = [];
  let from = 0;
  for (;;) {
    const p = norm.indexOf(needle, from);
    if (p < 0) break;
    const a = map[p];
    const b = (map[p + needle.length - 1] ?? a) + 1;
    out.push([a, b]);
    from = p + needle.length;
  }
  return out;
}

/** 구간에 <mark>를 씌운 HTML을 만든다 */
function markHTML(raw, ranges) {
  if (!ranges.length) return esc(raw);
  let html = '', at = 0;
  for (const [a, b] of ranges) {
    html += esc(raw.slice(at, a)) + '<mark class="hit-mark">' + esc(raw.slice(a, b)) + '</mark>';
    at = b;
  }
  return html + esc(raw.slice(at));
}

function decorateCN(html) {
  return html
    .replace(/^(\[\d{3,4}[abc]\d{2}\])\s*/gm, '<span class="mk">$1</span>')
    .replace(APP_TAG, m => `<span class="app">${m}</span>`);
}

/* 번역 문단 앞머리의 위치표지를 화면에서 걷어낸다.
   저본마다 표지를 붙인 것도 있고 안 붙인 것도 있어 들쭉날쭉하다.
   자리는 왼쪽 레일이 이미 알려 주므로 번역문에서는 지운다.
   ([0570a04] 하나든 [0119b14]–[0119b15] 범위든 모두 해당) */
const RE_KO_LEAD = /^\s*\[\d{3,4}[abc]\d{2}\](?:\s*[–—~-]\s*\[?\d{3,4}[abc]\d{2}\]?)*\s*/;
function stripKoMarker(t) {
  return String(t).replace(RE_KO_LEAD, '');
}

function markupCN(text) {
  // 위치표지와 교감 표지를 본문 글자와 구분해 보여 준다
  return decorateCN(esc(text));
}

// 도판 설명(있으면). 문헌을 열 때 채워 둔다.
let FIGCAP = {};

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
    // 짧은 줄(권두 표제·찬자 등)만 표제답게 조판한다
    if (u.cn.join('').length <= 80) n.classList.add('titleish');
  }

  const cnbox = el('div', 'cnbox');
  const p = el('p', 'cn');
  p.innerHTML = markupCN(u.cn.join('\n'));
  cnbox.append(p);

  const kobox = el('div', 'kobox');
  if (u.ko && u.ko.length) {
    u.ko.forEach(t => kobox.append(el('p', 'ko', stripKoMarker(t))));
  } else {
    kobox.append(el('p', 'nokr', '번역 대응 없음'));
  }

  n.append(rail, cnbox, kobox);

  // 저본에 실린 도판(결계도 등). 원문·번역 어느 쪽만 볼 때도 보이도록
  // 단위 아래에 따로 둔다.
  if (u.fig && u.fig.length) {
    const figs = el('div', 'figs');
    u.fig.forEach(name => {
      const fig = el('figure', 'fig');
      const img = el('img');
      img.src = `${BASE}assets/figures/${wid}/${name}`;
      img.alt = '저본 도판';
      img.loading = 'lazy';
      fig.append(img, el('figcaption', null, FIGCAP[name] || '저본 도판'));
      figs.append(fig);
    });
    n.append(figs);
  }

  if (u.nt && u.nt.length) {
    const nt = el('div', 'notes');    u.nt.forEach(t => {
      // 한불전 교감주는 요지만 내고, 자세한 내용은 손을 얹거나
      // 눌렀을 때 쪽지로 편다. (t 가 문자열이면 예전처럼 한 줄로)
      if (typeof t === 'string') { nt.append(el('p', null, t)); return; }
      const p = el('p', 'note-more');
      p.append(el('span', 'note-gist', t.t));
      const box = el('span', 'note-detail');
      (t.d || []).forEach(x => box.append(el('span', 'note-line', x)));
      p.append(box);
      p.tabIndex = 0;
      p.setAttribute('role', 'button');
      p.setAttribute('aria-expanded', 'false');
      p.addEventListener('click', e => {
        // 이 클릭은 '바깥을 눌렀다'로 세지 않는다.
        e.stopPropagation();
        const wasOpen = p.classList.contains('open');
        closeNote();
        if (wasOpen) {
          // 쪽지 안쪽을 눌러 닫은 경우, 커서가 그대로 있으면
          // hover 규칙이 즉시 다시 편다. 커서가 떠날 때까지만 막는다.
          p.classList.add('hover-off');
        } else {
          openNote = p;
          p.classList.add('open');
          p.classList.remove('hover-off');
          p.setAttribute('aria-expanded', 'true');
        }
      });
      p.addEventListener('keydown', e => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); p.click(); }
      });
      p.addEventListener('mouseleave', () => p.classList.remove('hover-off'));
      nt.append(p);
    });
    n.append(nt);
  }
  return n;
}

/* 교감주 쪽지(팝오버)
   한 번에 하나만 연다. 화면 어디를 누르든 — 쪽지 안쪽이든 바깥이든 —
   닫힌다. Esc 와 화면 넘김으로도 닫힌다. 웹에서 흔히 쓰는 방식이다. */
let openNote = null;

function closeNote() {
  if (!openNote) return;
  openNote.classList.remove('open');
  openNote.setAttribute('aria-expanded', 'false');
  openNote = null;
}

document.addEventListener('click', closeNote);
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeNote(); });
window.addEventListener('scroll', closeNote, true);
window.addEventListener('hashchange', closeNote);

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
    // 시대가 서로 다른 저자가 겹친 문헌은 author_short 로 따로 적는다
    src.append(el('span', null, w.author_short || `${w.dynasty} ${w.author_ko}`));
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
  B: '大藏經補編 · 대장경보편',
  HB: '韓國佛敎全書 · 한국불교전서',
  BJ: '韓國佛敎全書 · 한국불교전서',
  ZW: '藏外佛敎文獻 · 장외불교문헌',
};

/* ─── 화면: 문헌 목록 ────────────────────────────── */
async function viewHome() {
  const m = await manifest();
  const main = $('#main'); main.innerHTML = '';

  const hero = el('div', 'hero');
  hero.innerHTML = `
    <h1>천년의 문장,<br class="mo-only"> 오늘의 한글로</h1>
    <p>한문 불교문헌을 온전히 읽고 이해할 수 있도록,<br>
       신한글대장경이 새로운 번역의 길을 엽니다.</p>
    <div class="stat">
      <div><b>${m.totals.works}</b><span>수록 문헌</span></div>
    </div>`;
  main.append(hero);

  const grid = el('div', 'grid');
  m.works.forEach(w => {
    const a = el('a', 'card'); a.href = `#/w/${w.id}`;
    a.append(el('h3', null, w.title_cn));
    a.append(el('div', 'ko', w.title_ko));
    // 찬자 표기가 여느 꼴에 맞지 않는 문헌은 registry 의 author_line 을 그대로 쓴다
    a.append(el('div', 'ko2',
      w.author_line || `${w.dynasty} ${w.author_ko} · ${w.author_cn}`));
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
async function viewWork(id, anchor, hit) {
  const main = $('#main');
  main.innerHTML = '<p class="loading">문헌을 여는 중…</p>';
  renderSide(id);

  const doc = await work(id);
  FIGCAP = (doc.meta && doc.meta.figcap) || {};
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
    ['찬자', w.author_line
      ? esc(w.author_line)
      : `<span class="cnw">${esc(w.author_cn)}</span> · ${esc(w.author_ko)} (${esc(w.dynasty)})`],
    ['대장경', `<span class="cnw">${esc(w.canon)}</span>` +
      (w.canon_ko ? ` <span style="color:var(--ink-3)">— ${esc(w.canon_ko)}</span>` : '') +
      (w.verify ? ' <span class="tag warn">서지 확인 필요</span>' : '')],
    ...(w.note ? [['비고', esc(w.note)]] : []),
    ['수록 범위', w.range ? `${w.range[0]} – ${w.range[1]}` : '—'],
    ['대조 단위', `${w.units.toLocaleString()}개 · 번역 대응 ${Math.round(w.coverage * 100)}%`],
    ['원문 글자', w.chars_cn.toLocaleString() + '자'],
    ['원문 출처', (w.source === 'KABC'
        ? 'KABC 한국불교전서 · 동국대학교 불교학술원'
        : `${esc(w.source || 'CBETA')} 電子佛典集成`) +
      (w['底本_publisher'] ? ` · 저본 판권자 <span class="cnw">${esc(w['底本_publisher'])}</span> ©` : '') +
      (w.source === 'KABC'
        ? ' · <a href="#/rights">CC BY-NC-SA</a>'
        : ' · <a href="#/rights">CC BY-NC-SA 4.0</a>')],
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
  if (hit && hit.unit) {
    const k = doc.units.findIndex(u => u.i === hit.unit);
    if (k >= 0) ci = doc.units[k].c || 0;
  }
  if (!ci && anchor) {
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
    if (jump) {
      // 검색으로 들어왔으면 검색어가 놓인 자리까지 데려간다
      const target = hit && hit.q
        ? doc.units.find(u => u.i === hit.unit) ||
          doc.units.find(u => (u.m === jump) || ('u' + u.i === jump))
        : null;
      if (target) {
        const node = document.getElementById('u' + target.i);
        const mark = spotlight(node, target, hit.q, hit.where);
        if (mark) {
          scrollToTarget(mark, true);
          node.classList.add('flash');
          setTimeout(() => node.classList.remove('flash'), 2600);
        } else {
          jumpTo(jump);
        }
      } else {
        jumpTo(jump);
      }
    }
    else if (chapters.length > 1) window.scrollTo({ top: 0, behavior: 'auto' });
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

/** 검색으로 들어왔을 때: 해당 단위에서 검색어를 표시하고 그 자리로 옮긴다 */
function spotlight(node, unit, q, where) {
  if (!node || !q) return null;
  let first = null;

  const paint = (elems, raws, kind) => {
    elems.forEach((el, n) => {
      const raw = raws[n];
      if (raw == null) return;
      const ranges = findRanges(raw, kind, normQuery(q, kind));
      if (!ranges.length) return;
      const html = markHTML(raw, ranges);
      el.innerHTML = kind === 'cn' ? decorateCN(html) : html;
      if (!first) first = el.querySelector('.hit-mark');
    });
  };

  if (where !== 'ko') paint([...node.querySelectorAll('.cn')], [unit.cn.join('\n')], 'cn');
  const koRaw = (unit.ko || []).map(stripKoMarker);
  if (where !== 'cn') paint([...node.querySelectorAll('.ko')], koRaw, 'ko');

  // 어느 쪽이라 지정됐어도 못 찾으면 반대쪽도 훑는다
  if (!first) {
    paint([...node.querySelectorAll('.cn')], [unit.cn.join('\n')], 'cn');
    paint([...node.querySelectorAll('.ko')], koRaw, 'ko');
  }
  return first;
}

function setFont(v) {
  state.fs = Math.max(FS_MIN, Math.min(FS_MAX, v));
  localStorage.setItem('fs', String(state.fs));
  applyFont();
}

/** 고정 막대에 가리지 않게 목표를 화면에 올린다.
 *  모바일에서는 글꼴이 늦게 실려 본문 높이가 나중에 바뀐다. 그래서
 *  한 번 옮긴 뒤에도 자리를 다시 재어 어긋난 만큼 보정한다. */
function barOffset() {
  const top = document.querySelector('.topbar')?.getBoundingClientRect().height || 60;
  const bar = document.querySelector('.docbar')?.getBoundingClientRect().height || 0;
  return top + bar + 14;
}

function scrollToTarget(el, smooth) {
  if (!el) return;
  const place = (behavior) => {
    const y = window.scrollY + el.getBoundingClientRect().top - barOffset() - 8;
    window.scrollTo({ top: Math.max(0, y), behavior });
  };
  const ready = document.fonts?.ready || Promise.resolve();
  ready.then(() => {
    requestAnimationFrame(() => {
      place(smooth ? 'smooth' : 'auto');
      // 글꼴·줄바꿈이 뒤늦게 바뀌면 다시 맞춘다
      let tries = 0;
      const fix = () => {
        const off = el.getBoundingClientRect().top - barOffset();
        if (Math.abs(off) > 24 && tries < 8) {
          tries++;
          place('auto');
          setTimeout(fix, 160);
        }
      };
      setTimeout(fix, smooth ? 420 : 160);
    });
  });
}

function jumpTo(anchor) {
  const node = anchor.startsWith('u')
    ? document.getElementById(anchor)
    : $(`.unit[data-m="${CSS.escape(anchor)}"]`);
  if (!node) return;
  scrollToTarget(node, true);
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
  head.innerHTML = `<h2>‘${esc(q)}’ 찾기</h2>`
    + `<span class="n" id="resN">검색 중…</span>`
    + `<span class="range" id="resRange"></span>`;
  const perBox = el('div', 'perbox');
  head.append(perBox);
  main.append(head);

  const facets = el('div', 'facets'); main.append(facets);
  const list = el('div', 'hits'); main.append(list);

  const ids = m.works.map(w => w.id);
  const { hits, counts } = await search(q, ids, scope, state.variants);

  const total = Object.values(counts).reduce((a, b) => a + b, 0);
  $('#resN').textContent = total ? `${total.toLocaleString()}곳` : '';

  if (!total) {
    list.innerHTML = `<div class="empty"><b>찾은 곳이 없습니다</b>
      ${state.variants ? '다른 표기로 바꿔 보세요. 이체자는 함께 찾고 있습니다(卽·即, 眞·真, 敎·教, 說·説, 观·觀).' : '이체자를 포함한 검색 기능이 꺼져 있어 <br class="br-m">원본 글자 그대로만 찾습니다.<br>검색칸의 <b>이체자</b> 버튼을 누른 뒤 재검색 해보세요.'}</div>`;
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

  // 결과가 수천 곳에 이를 수 있으므로 쪽 단위로 나눠 보여 준다.
  const PER_OPTIONS = [10, 20, 30, 50, 100];
  let page = 1, shown = [];
  const pager = el('nav', 'pager');
  pager.setAttribute('aria-label', '검색 결과 쪽 이동');

  function hitNode(h) {
    const a = el('a', 'hit');
    a.href = `#/w/${h.work}/${h.m || 'u' + h.i}`
      + `?q=${encodeURIComponent(q)}&in=${h.where}&u=${h.i}`;
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
    return a;
  }

  function drawPerBox() {
    perBox.innerHTML = '';
    const sel = el('select', 'persel');
    sel.title = '한 쪽에 보여 줄 개수';
    PER_OPTIONS.forEach(n => {
      const o = el('option', null, `${n}개씩 출력`);
      o.value = String(n);
      if (n === state.perPage) o.selected = true;
      sel.append(o);
    });
    sel.onchange = () => {
      state.perPage = Number(sel.value);
      localStorage.setItem('perPage', String(state.perPage));
      page = 1;
      render();
    };
    perBox.append(sel);
  }

  function go(n) {
    page = n;
    render();
    const y = head.getBoundingClientRect().top + window.scrollY - 80;
    window.scrollTo({ top: Math.max(0, y), behavior: 'smooth' });
  }

  function drawPager(pages) {
    pager.innerHTML = '';
    if (pages <= 1) return;

    const btn = (label, target, cls) => {
      const b = el('button', 'pgbtn' + (cls ? ' ' + cls : ''), label);
      if (target === page) b.classList.add('on');
      if (target < 1 || target > pages || target === page) {
        if (cls) b.disabled = true;
      }
      b.onclick = () => go(Math.min(pages, Math.max(1, target)));
      return b;
    };

    const BLOCK = 10;                        // 한 번에 보여 줄 쪽 번호 수
    const from = Math.floor((page - 1) / BLOCK) * BLOCK + 1;
    const to = Math.min(from + BLOCK - 1, pages);

    if (from > 1) {
      pager.append(btn('«', 1, 'edge'));
      pager.append(btn('‹', from - 1, 'edge'));
    }
    for (let n = from; n <= to; n++) pager.append(btn(String(n), n));
    if (to < pages) {
      pager.append(btn('›', to + 1, 'edge'));
      pager.append(btn('»', pages, 'edge'));
    }
  }

  function render() {
    const pages = Math.max(1, Math.ceil(shown.length / state.perPage));
    if (page > pages) page = pages;
    const from = (page - 1) * state.perPage;
    const slice = shown.slice(from, from + state.perPage);

    list.innerHTML = '';
    const frag = document.createDocumentFragment();
    slice.forEach(h => frag.append(hitNode(h)));
    list.append(frag);

    drawPager(pages);
    if (!pager.isConnected) main.append(pager);

    const info = document.getElementById('resRange');
    if (info) {
      info.textContent = shown.length
        ? `${(from + 1).toLocaleString()}–${(from + slice.length).toLocaleString()} / ${shown.length.toLocaleString()}곳`
        : '';
    }
  }

  function draw() {
    shown = hits.filter(h => !state.facet || h.work === state.facet);
    page = 1;
    drawPerBox();
    render();
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
         (${esc(R.source_org || '')})와
         <b>KABC 한국불교전서</b>(동국대학교 불교학술원)에서 가져왔습니다.
         문헌마다 어느 쪽에서 왔는지는 서지에 밝혀 두었습니다.</p>
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
         만신찬속장경 株式會社國書刊行會 ©, 건륭대장경 新文豐出版公司,
         한국불교전서 東國大學校出版部 ©.<br>
         CBETA 판본: ${esc(R.cbeta_version || '확인 필요')}<br>
         원문 이용 조건 전문은 <a href="${esc(R.source_url || 'https://cbeta.org/copyright')}"
         target="_blank" rel="noopener">CBETA 版權宣告</a>을 보십시오.</p>
      <p class="rights-note">KABC 불교기록문화유산 아카이브는 누리집 아래쪽에
         <b>저작자표시–비영리–동일조건변경허락(CC BY-NC-SA)</b> 표시를 걸어 두었습니다.
         이 사이트가 따르는 조건과 같습니다. 다만 버전 번호는 밝혀져 있지 않습니다.
         서비스 주체는 동국대학교 불교학술원 불교기록문화유산 아카이브(ABC) 사업단이며,
         저본 『한국불교전서』의 판권은 동국대학교출판부에 있습니다.</p>
      <p class="rights-warn">CBETA 수록 문헌 가운데 인순법사·여징·태허·연배 등
         근현대 저자의 저작집(類別 B)은 CC 조건이 적용되지 않습니다.
         이 사이트는 해당 문헌을 수록하지 않습니다.</p>
    </section>

    <section>
      <h2>다른 전자 대장경을 쓸 때</h2>
      <p>기관마다 조건이 다릅니다. 이 사이트는 재배포가 허용된 CBETA와,
        보호기간이 끝난 원문만 옮겨 온 KABC 한국불교전서를 수록합니다.</p>
      <table class="src-table">
        <tr><th>CBETA</th>
            <td><b class="ok">수록 가능</b><br>
                CC BY-NC-SA 4.0. 비영리·출처 명시·동일조건 공개를 지키면 재배포할 수 있습니다.</td></tr>
        <tr><th>SAT</th>
            <td><b class="no">수록 불가</b><br>
                利用条件 제3조가 인터넷·기타 매체를 통한 재배포를 당분간 금지합니다.
                저작권법상 인용은 별개이므로, 링크를 걸거나 필요한 대목만 인용하는 방식으로 쓰십시오.</td></tr>
        <tr><th>KABC<br><span class="src-sub">한국불교전서</span></th>
            <td><b class="ok">원문 한정 수록</b><br>
                CC BY-NC-SA 표시를 걸고 있으며, 한문 원문 자체는 저작권 보호기간이 끝난 고전입니다.
                한국어 번역은 KABC의 역주와 무관하게 이 사이트에서 따로 생성한 것입니다.
                KABC 편집자의 교감 판단은 요지를 한국어로 옮겨 각주 앞머리에 두고,
                근거가 되는 교감문은 쪽지 안쪽에 출처를 밝혀 함께 보였습니다.
                이용하실 때 출처를 함께 밝혀 주십시오.</td></tr>
      </table>
      <p class="rights-note">여기 적은 것은 각 기관이 공개한 이용 조건을 정리한 것이며 법률 자문이 아닙니다.
        조건은 바뀔 수 있으니 수록 전에 해당 기관의 최신 고지를 확인하십시오.</p>
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
         원문 출처: CBETA 電子佛典集成 (CC BY-NC-SA 4.0) · KABC 한국불교전서
         (동국대학교 불교학술원, CC BY-NC-SA). 열람일: <span id="today"></span>.</p>
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
    return viewWork(seg[1], seg[2], {
      q: params.get('q') || '',
      where: params.get('in') || '',
      unit: Number(params.get('u') || 0),
    });
  }
  if (seg[0] === 'rights') {
    return viewRights();
  }
  if (seg[0] === 's') {
    document.querySelector('[data-tab="search"]')?.classList.add('on');
    const q = params.get('q') || '';
    state.scope = params.get('scope') || 'all';
    state.variants = params.get('var') !== '0';
    syncScope();
    return q ? viewSearch(q, state.scope) : viewHome();
  }
  document.querySelector('[data-tab="home"]')?.classList.add('on');
  viewHome();
}

function syncScope() {
  document.querySelectorAll('#scopeSeg button').forEach(b =>
    b.classList.toggle('on', b.dataset.scope === state.scope));
  const vb = $('#varBtn');
  if (vb) vb.checked = state.variants;
}

/** 지금 검색어로 다시 찾아간다. */
function rerunSearch() {
  const q = $('#q').value.trim();
  if (!q) return;
  location.hash = `#/s?q=${encodeURIComponent(q)}&scope=${state.scope}`
                + (state.variants ? '' : '&var=0');
}

/* ─── 이벤트 ─────────────────────────────────────── */
$('#searchForm').addEventListener('submit', e => {
  e.preventDefault();
  const q = $('#q').value.trim();
  if (!q) return;
  state.facet = null;
  rerunSearch();
});
$('#scopeSeg').addEventListener('click', e => {
  const b = e.target.closest('button'); if (!b) return;
  state.scope = b.dataset.scope; syncScope();
  rerunSearch();
});
$('#varBtn')?.addEventListener('change', e => {
  state.variants = e.target.checked;
  rerunSearch();
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
