/* 검색 워커
   본문 스캔을 메인 스레드에서 떼어낸다. 현재는 선형 탐색이지만
   말뭉치가 커지면 이 파일만 교체하면 된다(→ 이분 인덱스 또는 서버 API). */

const MARKER = /\[\d{3,4}[abc]\d{2}\]/g;
const APP = /\[(?:\d{1,3}|＊|\*)\]/g;
const PUNCT = /[\s。，、．・？！：；「」『』（）()〔〕【】\[\]“”‘’·…—　]/g;

// 판본 간 흔들리는 글자 접기 (필요할 때 추가)
const VAR = {
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

function fold(s){
  let out = '';
  for (const ch of s) out += (VAR[ch] || ch);
  return out;
}
/* 이체자 접기는 검색할 때 켜고 끌 수 있다.
   끄면 판본이 쓴 글자 그대로만 맞춰 본다(문헌학적 대조용). */
function normCN(s, folding = true){
  const t = s.normalize('NFKC').replace(MARKER,'').replace(APP,'').replace(PUNCT,'');
  return folding ? fold(t) : t;
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
    return { i:u.i, m:u.m, cnRaw, koRaw,
             ncn:normCN(cnRaw, true),   // 이체자를 접은 색인
             ncnX:normCN(cnRaw, false), // 판본 글자 그대로의 색인
             nko:normKO(koRaw) };
  });
  const rec = { meta: doc.meta, rows };
  store.set(id, rec);
  return rec;
}

/* 정규화 문자열의 위치를 원문 위치로 되돌리기 위한 지도.
   위치표지·교감표지는 통째로 건너뛰어야 하므로 먼저 구간을 표시해 둔다. */
const PUNCT1 = /[\s。，、．・？！：；「」『』（）()〔〕【】\[\]“”‘’·…—　]/;

function buildMap(raw, kind, folding = true){
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
      n = PUNCT1.test(s) ? '' : (folding ? fold(s) : s);
    } else {
      n = /\s/.test(ch) ? '' : ch.normalize('NFKC').toLowerCase();
    }
    for (let k = 0; k < n.length; k++) map.push(i);
    norm += n;
  }
  return { norm, map };
}

function kwic(raw, kind, needle, pad, folding = true){
  const { norm, map } = buildMap(raw, kind, folding);
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
    const folding = payload.variants !== false;
    const qcn = normCN(query, folding), qko = normKO(query);
    if (!qcn && !qko){ self.postMessage({ type:'result', token, payload:{ hits:[], counts:{} } }); return; }

    const hits = [], counts = {};
    for (const id of ids){
      let rec;
      try { rec = await load(id, base); } catch(e){ continue; }
      let c = 0;
      for (const r of rec.rows){
        let where = null;
        const idx = folding ? r.ncn : r.ncnX;
        if (scope !== 'ko' && qcn && idx.includes(qcn)) where = 'cn';
        else if (scope !== 'cn' && qko && r.nko.includes(qko)) where = 'ko';
        if (!where) continue;
        c++;
        if (!limit || hits.length < limit){
          const snip = where === 'cn'
            ? kwic(r.cnRaw, 'cn', qcn, 22, folding)
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
