/* 검색 워커
   본문 스캔을 메인 스레드에서 떼어낸다. 현재는 선형 탐색이지만
   말뭉치가 커지면 이 파일만 교체하면 된다(→ 이분 인덱스 또는 서버 API). */

const MARKER = /\[\d{3,4}[abc]\d{2}\]/g;
const APP = /\[(?:\d{1,3}|＊|\*)\]/g;
const PUNCT = /[\s。，、．・？！：；「」『』（）()〔〕【】\[\]“”‘’·…—　]/g;

// 판본 간 흔들리는 글자 접기 (필요할 때 추가)
const VAR = {
  '眞':'真','衆':'眾','众':'眾','觉':'覺','说':'說','为':'為','无':'無','与':'與',
  '体':'體','万':'萬','来':'來','实':'實','义':'義','经':'經','论':'論','号':'號',
  '当':'當','从':'從','学':'學','断':'斷','边':'邊','转':'轉','显':'顯','现':'現',
  '应':'應','处':'處','随':'隨','点':'點','师':'師','净':'淨','凈':'淨','烦':'煩',
  '恼':'惱','萨':'薩','刹':'剎','弥':'彌','广':'廣','严':'嚴','华':'華','会':'會',
  '个':'個','于':'於','后':'後','裏':'裡','么':'麼','并':'並','余':'餘','觀':'觀','观':'觀'
};

function fold(s){
  let out = '';
  for (const ch of s) out += (VAR[ch] || ch);
  return out;
}
function normCN(s){
  return fold(s.normalize('NFKC').replace(MARKER,'').replace(APP,'').replace(PUNCT,''));
}
function normKO(s){
  return s.normalize('NFKC').toLowerCase().replace(/\s+/g,'');
}

const store = new Map(); // id -> {meta, rows:[{i,m,cn,ko,ncn,nko,cnRaw,koRaw}]}

async function load(id, base){
  if (store.has(id)) return store.get(id);
  const res = await fetch(`${base}data/works/${id}.json`);
  const doc = await res.json();
  const rows = doc.units.map(u => {
    const cnRaw = u.cn.join('\n');
    const koRaw = (u.ko || []).join('\n');
    return { i:u.i, m:u.m, cnRaw, koRaw, ncn:normCN(cnRaw), nko:normKO(koRaw) };
  });
  const rec = { meta: doc.meta, rows };
  store.set(id, rec);
  return rec;
}

/* 정규화 문자열의 위치를 원문 위치로 되돌리기 위한 지도.
   위치표지·교감표지는 통째로 건너뛰어야 하므로 먼저 구간을 표시해 둔다. */
const PUNCT1 = /[\s。，、．・？！：；「」『』（）()〔〕【】\[\]“”‘’·…—　]/;

function buildMap(raw, kind){
  const skip = new Uint8Array(raw.length);
  if (kind === 'cn'){
    for (const re of [MARKER, APP]){
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(raw)) !== null){
        for (let i = m.index; i < m.index + m[0].length; i++) skip[i] = 1;
        if (m[0].length === 0) re.lastIndex++;
      }
    }
  }
  const map = []; let norm = '';
  for (let i = 0; i < raw.length; i++){
    if (skip[i]) continue;
    const ch = raw[i];
    let n;
    if (kind === 'cn'){
      const s = ch.normalize('NFKC');
      n = PUNCT1.test(s) ? '' : fold(s);
    } else {
      n = /\s/.test(ch) ? '' : ch.normalize('NFKC').toLowerCase();
    }
    for (let k = 0; k < n.length; k++) map.push(i);
    norm += n;
  }
  return { norm, map };
}

function kwic(raw, kind, needle, pad){
  const { norm, map } = buildMap(raw, kind);
  const p = norm.indexOf(needle);
  if (p < 0) return null;
  const s = map[p] ?? 0;
  const e = (map[p + needle.length - 1] ?? s) + 1;
  const a = Math.max(0, s - pad), b = Math.min(raw.length, e + pad);
  return {
    pre: (a > 0 ? '…' : '') + raw.slice(a, s),
    hit: raw.slice(s, e),
    post: raw.slice(e, b) + (b < raw.length ? '…' : '')
  };
}

self.onmessage = async (ev) => {
  const { type, payload, token } = ev.data;

  if (type === 'warm'){
    for (const id of payload.ids){ try { await load(id, payload.base); } catch(e){} }
    self.postMessage({ type:'warm-done', token });
    return;
  }

  if (type === 'search'){
    const { query, ids, base, scope, limit } = payload;
    const qcn = normCN(query), qko = normKO(query);
    if (!qcn && !qko){ self.postMessage({ type:'result', token, payload:{ hits:[], counts:{} } }); return; }

    const hits = [], counts = {};
    for (const id of ids){
      let rec;
      try { rec = await load(id, base); } catch(e){ continue; }
      let c = 0;
      for (const r of rec.rows){
        let where = null;
        if (scope !== 'ko' && qcn && r.ncn.includes(qcn)) where = 'cn';
        else if (scope !== 'cn' && qko && r.nko.includes(qko)) where = 'ko';
        if (!where) continue;
        c++;
        if (!limit || hits.length < limit){
          const snip = where === 'cn'
            ? kwic(r.cnRaw, 'cn', qcn, 22)
            : kwic(r.koRaw, 'ko', qko, 34);
          hits.push({
            work:id, title:rec.meta.title_cn, titleKo:rec.meta.title_ko,
            i:r.i, m:r.m, where, snip,
            alt: where === 'cn'
              ? (r.koRaw ? r.koRaw.slice(0,90) : '')
              : (r.cnRaw ? r.cnRaw.replace(MARKER,'').slice(0,60) : '')
          });
        }
      }
      counts[id] = c;
      self.postMessage({ type:'progress', token, payload:{ id, count:c } });
    }
    self.postMessage({ type:'result', token, payload:{ hits, counts } });
  }
};
