#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
藏經 대조본 인게스트 파이프라인
------------------------------------------------------------------
입력:  sources/<work_id>/원문.txt        (CBETA/SAT/KABC 복사본, 선택)
       sources/<work_id>/번역.docx       (한국어 대조 번역, 선택)
       sources/works.yml 대신 tools/registry.json 에 서지 메타 기술
출력:  data/works/<work_id>.json         (열람·검색용 정규화 단위)
       data/manifest.json                (문헌 목록 + 통계)

핵심 설계
  1) 정렬 키는 '위치표지'(CBETA 行標, 예: [0287b17]) 이다.
     CBETA/SAT/KABC 가 공유하는 유일한 안정 좌표이므로 이것을 단위 ID로 삼는다.
  2) docx 는 문서마다 스타일 이름이 제각각이므로(Source Text / Chinese Source /
     대조 원문 / Normal ...) 스타일에 의존하지 않는다. 대신 '문자 스크립트 비율'로
     한문 단락과 한글 단락을 판정한다. 어떤 형식의 번역 docx 가 들어와도 동작한다.
  3) 위치표지가 없는 docx 는 한문 단락의 앞머리를 원문 TXT 와 대조하여 정렬한다.
"""

import json, os, re, sys, unicodedata, hashlib
from pathlib import Path

try:
    import docx  # python-docx
except ImportError:
    docx = None

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "sources"
OUT = ROOT / "data"

# ── 위치표지: [0287b17] / [1089a06] 형태 ──────────────────────────────
RE_MARKER = re.compile(r"\[(\d{3,4}[abc]\d{2})\]")
# 교감 표지: [1] [＊] [12] 등 (본문 아님)
RE_APPARATUS = re.compile(r"\[(?:\d{1,3}|＊|\*)\]")
RE_CJK = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
RE_HANGUL = re.compile(r"[\uac00-\ud7a3]")

# 번역 docx 에서 반복되는 '라벨' 단락 (내용 아님)
LABELS = {
    "원문", "한문 원문", "한국어 번역", "한국어 직역", "번역", "직역",
    "【註】", "【本文】", "【科文】", "【科層】",
    "【논서 번역】", "【주석서 번역】", "【번역 각주】", "번역 각주",
    # 칸 이름을 한문으로 다는 문서도 있다
    "原文", "譯文", "譯註", "校勘", "校勘 · 譯註", "校勘·譯註", "註釋",
}
# ── 표제(heading) 분류 ────────────────────────────────────────────────
# 편집 안내·용어표·검증 보고 등 '해제' 성격의 표제. 이 아래의 한국어 문단은
# 번역이 아니므로 본문 단위에 붙이지 않는다.
RE_APPARATUS_HEAD = re.compile(
    r"^[A-Z]\.\s|"
    r"문헌\s*정보|편집|안내|범례|용어\s*(?:표|대응)|검증|검수|선독|판본\s*정보|"
    r"교정\s*(?:기록|보고|사항)|정리\s*(?:기록|보고|사항)|작업\s*(?:기록|보고)|"
    r"(?:정리|편집|번역|작업|교열|대조|처리|표기|배열)\s*(?:기준|방침|원칙|규칙|방법)|"
    r"(?:최종|전체|일괄)\s*(?:정리|교열|점검|확인)|"
    r"(?:편집|번역|대조|원문)\s*(?:검증|검수|확인)\s*(?:요약|결과|보고)?|"
    r"검증\s*요약|검수\s*요약|"
    r"번역상|쟁점|한계|읽는\s*법|과단\s*계층|부록|참고\s*문헌|대역\s*목차"
)
# 「1. …」 「2.3 …」 처럼 번호를 매긴 표제는 번역 docx 가 스스로 붙인
# 문서 차례이지 저본의 권·품이 아니다. 저본의 표제는 「제1 발심」·「권제1」·
# 「初篇」 꼴이라 이 규칙에 걸리지 않는다.
RE_DOC_OUTLINE = re.compile(r"^\d{1,2}(?:\.\d{1,2})*[.、]\s|^\d{1,2}(?:\.\d{1,2})+\s")
# 단위 라벨: 위치표지·문단 번호·「원문/번역」 같은 꼬리표만으로 이루어진 줄.
# 한글이 섞여 있어 번역문으로 오인되기 쉬우므로 본문에 들어가기 전에 걸러 낸다.
#   예) "[1206c22]  문단 001 · 원문" / "문단 003 · 한국어 직역" / "005. 번역"
RE_UNIT_LABEL = re.compile(
    r"^\s*(?:\[\d{3,4}[abc]\d{2}\])?\s*"
    r"(?:(?:문단|단락|단위|제)\s*\d+\s*(?:항|번|단)?)?\s*"
    r"(?:\d{1,4}\s*[.)]?)?\s*"
    r"(?:[·・|/:：\-—~]\s*)?"
    r"(?:한문\s*)?(?:한국어\s*)?"
    r"(?:원문|본문|번역|직역|대조|원문·번역)?\s*$"
)
LABEL_WORDS = ("원문", "번역", "직역", "본문", "대조", "문단", "단락", "단위")

# 단락 앞머리에 붙는 꼬리표: "원문  大乘起信論…" / "번역  『대승기신론소』…"
# 이 두 글자 때문에 짧은 한문 줄이 번역으로 오인되므로 먼저 떼어 낸다.
RE_SIDE_PREFIX = re.compile(
    r"^(원문|한문\s*원문|본문|번역|직역|한국어\s*직역|한국어\s*번역)\s*[:：·|]?\s+"
)


def strip_side(t: str):
    """앞머리 꼬리표를 떼고, 그것이 가리키던 쪽('cn'/'ko'/None)을 함께 돌려준다."""
    m = RE_SIDE_PREFIX.match(t)
    if not m:
        return t, None
    head = re.sub(r"\s", "", m.group(1))
    side = "cn" if head in ("원문", "한문원문", "본문") else "ko"
    return t[m.end():].strip(), side


def is_unit_label(t: str) -> bool:
    """내용 없는 단위 라벨인가. 실제 번역문을 잘못 지우지 않도록 조건을 좁게 둔다."""
    t = t.strip()
    if not t or len(t) > 40:
        return False
    if not any(w in t for w in LABEL_WORDS):
        return False
    return bool(RE_UNIT_LABEL.match(t))


# 개별 단위마다 붙는 표제(구조가 아님)
RE_UNIT_HEAD = re.compile(
    r"^(?:문단\s*\d+|\d{2,4}\.\s|권두\s*표제|\[\d{3,4}[abc]\d{2}\])"
)
# 라벨 성격의 표제 — 무시하되 단위를 끊지 않는다
RE_NOTE_HEAD = re.compile(r"(?:^|\s)(?:주|주석|각주|교감)$")
# 각주 묶음의 칸 이름. '교감·번역 메모'처럼 스타일 구분 없이 라벨 문단만으로
# 각주 시작을 알리는 docx 가 있어, 이 줄 아래는 각주로 모은다.
# 각주 묶음의 칸 이름. 라벨의 '끝 낱말'이 각주 계열이어야 한다.
#   주석·교감 / 교감 / 교감·번역 메모  → 각주 칸
#   【주석서 번역】 / 교감 표시 규칙     → 각주 칸이 아님
RE_NOTE_LABEL_HEAD = re.compile(
    r"^(?=.{1,16}$)[^\[\]\n。.]*"
    r"(?:메모|주기|비고|각주|주석|교감|교정)"
    r"\s*[】\]）)]?\s*[:：]?$")
# 다만 '교감 표시 규칙'처럼 안내 표제인 것은 각주 칸 이름이 아니다.
RE_NOTE_LABEL_NOT = re.compile(r"규칙|원칙|방식|기준|방법|안내|범례|일러두기")

# 줄 전체가 줄표(—·―·–)로 감싸인 짧은 문단 = 번역본 제작 쪽의 작업 표지.
# 예) '— 원문 전체 위치표지 대응 재배열 완료 —'
# 본문이나 번역이 아니므로 버린다. 안쪽에 줄표가 또 있으면 본문일 수 있어
# 제외하고, 한문 원문 줄과 겹치지 않도록 길이도 제한한다.
RE_WORKLOG = re.compile(r"^[—―–]\s*[^—―–]{2,40}\s*[—―–]$")
# 칸 이름에 이어짐 꼬리가 붙는 docx 가 있다.
#   '원문 · 계속' · '한국어 직역 (계속)' · '원문 계속'
# 가운뎃점으로 잇기도 하고 괄호로 묶기도 하므로 둘 다 떼고 칸 이름을 견준다.
RE_LABEL_TAIL = re.compile(
    r"\s*(?:[·・|/／\-–—]\s*)?[(（]?\s*"
    r"(?:계속|이어짐|이어서|continued|cont\.?)\s*[)）]?\s*$", re.I)


# 도판 자리를 말로 대신한 번역 문단. 도판을 띄우면 필요 없어진다.
RE_FIG_CAPTION = re.compile(
    r"(?:문자\s*본문|본문\s*문자)[^。.\n]{0,24}?(?:없|비어)"
    r"|이\s*(?:위치|자리)에[^。.\n]{0,24}?(?:도판|그림|圖|계도|界圖)"
    r"[^。.\n]{0,12}?(?:배치|삽입|들어)")


def label_key(t: str) -> str:
    return RE_LABEL_TAIL.sub("", t).rstrip("：: ").strip()


LABEL_HEADS = {"원문", "원문 목차", "한국어 목차", "한국어 직역", "한국어 번역",
               "번역", "직역"}
# 원문 안의 권 표제 (원문 TXT 만 있는 문헌의 분권 검출용)
RE_JUAN_LINE = re.compile(
    r"^.{0,24}?卷(?:第[一二三四五六七八九十百]+|[上中下])"
    r"(?:之[上中下]|[上中下])?(?:[(（].{0,12}?[)）])?\s*$"
)

# '주석'·'각주' 로 시작하더라도 그것이 칸 이름일 때만 각주로 본다.
# ('주석 [용어] …' 은 각주, '주석 가운데 먼저 …' 는 본문 번역이다)
NOTE_PREFIX = re.compile(
    r"^\s*(?:(?:주석|각주)\s*\d*\s*(?=[\[【(（:：|·]|$)"
    r"|교감[·\s]|번역 각주|【번역 각주】"
    r"|\[구두 불확실\]|\[판독 불확실\]|\[구문 불확실\])"
)


# ── 문자 판정 ────────────────────────────────────────────────────────
def hangul_ratio(s: str) -> float:
    letters = [c for c in s if unicodedata.category(c).startswith("L")]
    if not letters:
        return 0.0
    return sum(1 for c in letters if RE_HANGUL.match(c)) / len(letters)


def is_source_line(s: str) -> bool:
    """한문 원문 단락인가.

    번역문에는 반드시 한글이 섞이므로 한글 비율로 가른다.
    卷首·卷尾·諦 처럼 두세 글자짜리 원문 조각도 놓치지 않도록
    길이 조건은 최소한으로만 둔다."""
    body = RE_MARKER.sub("", s).strip()
    body = re.sub(r"^[【〔\[(（]+|[】〕\])）]+$", "", body).strip()
    if len(body) < 2:
        return False
    if not RE_CJK.search(body):
        return False
    return hangul_ratio(body) < 0.12


# ── 검색용 정규화(이체자 폴딩) ────────────────────────────────────────
# CBETA·SAT·KABC 판본 간 흔들리는 글자만 최소로 접는다. 필요할 때 추가하면 된다.
VARIANTS = {
    # ① 한국에서 쓰는 자형 → CBETA·SAT 판본의 자형
    #    (한글 글꼴로 입력한 검색어가 원문과 어긋나는 것을 막는다)
    "卽": "即", "敎": "教", "眞": "真", "爲": "為", "衆": "眾", "敍": "敘",
    "兪": "俞", "靑": "青", "絶": "絕", "戸": "戶", "户": "戶", "幷": "并",
    "擧": "舉", "恆": "恒", "冊": "册", "郞": "郎", "祕": "秘", "麤": "麁",
    "况": "況", "决": "決", "畧": "略", "凈": "淨", "裏": "裡",
    "鑒": "鑑", "慙": "慚", "飜": "翻", "遲": "遅", "旣": "既", "飮": "飲",
    "衞": "衛", "羣": "群", "菴": "庵", "翫": "玩", "皁": "皂", "牀": "床",
    "氷": "冰", "朶": "朵", "擡": "抬", "掛": "挂", "覩": "睹", "踰": "逾",
    "顚": "顛", "鬪": "鬥", "姊": "姉", "妬": "妒", "嶽": "岳", "疊": "叠",
    "甎": "磚", "筯": "箸", "禱": "祷", "稟": "禀", "剏": "創", "勑": "敕",
    "勅": "敕", "軆": "體", "釼": "劍", "劒": "劍", "閒": "間", "餧": "餵",
    "髣": "仿", "賔": "賓", "栰": "筏", "灋": "法", "噐": "器", "瑯": "琅",
    "逈": "迥", "盋": "缽", "鉢": "缽", "蘂": "蕊",

    # ② 일본 신자체 → 판본의 자형 (SAT 계열 자료를 함께 찾기 위함)
    "歴": "歷", "虚": "虛", "温": "溫", "説": "說", "増": "增", "徳": "德",
    "対": "對", "実": "實", "経": "經", "処": "處", "帰": "歸", "独": "獨",
    "変": "變", "数": "數", "関": "關", "両": "兩", "称": "稱", "証": "證",
    "釈": "釋", "寿": "壽", "薬": "藥", "礼": "禮", "発": "發", "国": "國",
    "気": "氣", "楽": "樂", "辞": "辭", "悪": "惡", "盗": "盜", "頼": "賴",
    "黄": "黃", "斉": "齊", "竜": "龍", "亀": "龜", "尽": "盡", "旧": "舊",
    "声": "聲", "触": "觸", "読": "讀", "鉄": "鐵", "雑": "雜", "霊": "靈",
    "黙": "默", "斎": "齋", "歯": "齒", "禅": "禪", "静": "靜", "黒": "黑",
    "摂": "攝", "継": "繼", "繊": "纖", "聡": "聰", "粛": "肅", "臈": "臘",
    "軽": "輕", "駆": "驅", "篭": "籠", "栄": "榮", "荘": "莊", "剣": "劍",
    "円": "圓", "戦": "戰", "帯": "帶", "広": "廣", "厳": "嚴", "覚": "覺",
    "観": "觀", "壊": "壞", "頬": "頰", "巻": "卷", "昼": "晝",
    "写": "寫", "壌": "壤", "収": "收", "拝": "拜", 
    # ③ 중국 간체 → 번체 (간체로 입력한 검색어도 함께 찾는다)
    "众": "眾", "觉": "覺", "说": "說", "为": "為", "无": "無", "与": "與",
    "体": "體", "万": "萬", "来": "來", "实": "實", "义": "義", "经": "經",
    "论": "論", "号": "號", "当": "當", "从": "從", "学": "學", "断": "斷",
    "边": "邊", "转": "轉", "显": "顯", "现": "現", "应": "應", "处": "處",
    "随": "隨", "点": "點", "师": "師", "净": "淨", "烦": "煩", "恼": "惱",
    "萨": "薩", "刹": "剎", "弥": "彌", "广": "廣", "严": "嚴", "华": "華",
    "会": "會", "个": "個", "于": "於", "么": "麼",      "观": "觀", "释": "釋", "灵": "靈", "杂": "雜", "赖": "賴",
    "龙": "龍", "龟": "龜", "齿": "齒", "关": "關", "两": "兩", "国": "國",
    "气": "氣", "乐": "樂", "恶": "惡", "旧": "舊", "尽": "盡", "独": "獨",
    "变": "變", "数": "數", "归": "歸", "证": "證", "称": "稱", "药": "藥",
    "礼": "禮", "发": "發", "齐": "齊", "黄": "黃", "声": "聲", "读": "讀",
    "铁": "鐵", "禅": "禪", "静": "靜", "摄": "攝", "继": "繼",
    "纤": "纖", "聪": "聰", "肃": "肅", "轻": "輕", "驱": "驅", "笼": "籠",
    "荣": "榮", "庄": "莊", "剑": "劍", "战": "戰", "带": "帶", "壊": "壞",
    # ④ 저본 안에서 실제로 흔들리는 글자
    #    (같은 말뭉치 안에 두 자형이 함께 쓰여, 접지 않으면 검색이 갈린다)
    "徧": "遍", "疎": "疏", "踈": "疏", "薰": "熏", "竝": "並", "歎": "嘆",
    "廻": "迴", "訶": "呵", "葢": "蓋", "盖": "蓋", "祗": "祇", "註": "注",
    "灯": "燈", "弃": "棄", "竪": "豎", "甞": "嘗", "蹟": "跡", "迹": "跡",
    "詶": "酬", "酧": "酬", "慜": "愍", "峯": "峰", "凉": "涼", "燄": "焰",
    "昬": "昏", "煖": "暖", "覔": "覓", "怜": "憐", "皃": "貌", "遶": "繞",
    "瑠": "琉", "璢": "琉", "舩": "船", "巖": "岩", "醎": "鹹", "餝": "飾",
    "罸": "罰", "綵": "彩", "輙": "輒", "鍊": "煉", "谿": "溪", "閙": "鬧",
    "鎻": "鎖", "諠": "喧", "踪": "蹤", "陜": "陝", "賸": "剩", "旛": "幡",
    "冲": "沖", "踴": "踊", "脇": "脅", "飡": "餐", "遯": "遁", "亙": "亘",
    "栢": "柏", "髴": "彿", "鵞": "鵝", "賖": "賒", "貍": "狸", "頺": "頹",
    "隷": "隸", "麪": "麵", "舘": "館", "廼": "迺", "嶋": "島", "鈆": "鉛",
    "阯": "址", "笋": "筍", "騐": "驗", "験": "驗", "躰": "體", "荅": "答",
    "摠": "總", "麽": "麼", "碍": "礙", "虗": "虛", "覈": "核", "讎": "讐",
    "棱": "稜", "賛": "贊",

    # ⑤ 저본에는 한 자형만 있으나 다른 판본·입력기에서 흔히 쓰는 짝
    "縁": "緣", "悩": "惱", "蔵": "藏", "脱": "脫", "蕐": "華", "逹": "達",
    "贒": "賢", "舎": "舍", "繋": "繫", "徴": "徵", "没": "沒", "縦": "縱",
    "莭": "節", "歳": "歲", "睠": "眷", "医": "醫", "銕": "鐵", "渉": "涉",
    "悦": "悅", "甁": "瓶", "歩": "步", "毎": "每", "鼔": "鼓", "兎": "兔",
    "冩": "寫", "渇": "渴", "燐": "憐", "蘓": "蘇", "邨": "村", "売": "賣",
    "担": "擔", "潅": "灌", "謌": "歌", "砕": "碎", "検": "檢", "窃": "竊",
    "鄕": "鄉", "髄": "髓", "涙": "淚", "疣": "肬", "稲": "稻", "禄": "祿",
    "莵": "菟", "緑": "綠", "殻": "殼", "窗": "窓", "裵": "裴", "挿": "插",
    "聨": "聯", "翛": "倏", "鬂": "鬢", "讁": "謫", "齢": "齡", "韈": "襪",
    "顋": "腮", "貮": "貳", "聟": "壻", "蛮": "蠻", "畞": "畝", "籖": "籤",
    "湼": "涅", "饍": "膳",
}


# CBETA 조자(組字) 표기 — '[狦-(狂-王)]' · '[匚@出]' · '[○@(內-入+人)]'
RE_CBETA_GLYPH = re.compile(r"\[[^\[\]]{2,24}\]")


def norm_search(s: str) -> str:
    s = unicodedata.normalize("NFKC", s)
    s = RE_MARKER.sub("", s)
    s = RE_APPARATUS.sub("", s)
    # 같은 글자를 저본은 확장한자로, 번역본은 조자(組字) 표기 '[狦-(狂-王)]' 로
    # 적는 일이 있다. 대조에 방해가 되므로 둘 다 지우고 견준다.
    s = RE_CBETA_GLYPH.sub("", s)
    s = "".join(c for c in s if ord(c) < 0x20000)
    s = "".join(VARIANTS.get(c, c) for c in s)
    s = re.sub(r"[\s。，、．・？！：；「」『』（）()〔〕【】\[\]“”‘’·…—　]", "", s)
    return s


def head_key(s: str, n: int = 14) -> str:
    """정렬용 앞머리 지문."""
    return norm_search(s)[:n]


# ── TXT 파서 ─────────────────────────────────────────────────────────
def parse_txt(path: Path):
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8", "cp949", "utf-16"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError(f"인코딩 판별 실패: {path}")

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    blocks = [b.strip() for b in re.split(r"\n\s*\n", text) if b.strip()]

    units, order = [], 0
    for b in blocks:
        # 한 블록 안에 위치표지가 여러 개면 표지 단위로 다시 쪼갠다
        parts = re.split(r"(?=\[\d{3,4}[abc]\d{2}\])", b)
        parts = [p.strip() for p in parts if p.strip()]
        for p in parts:
            m = RE_MARKER.match(p)
            order += 1
            units.append({
                "i": order,
                "m": m.group(1) if m else None,
                "cn": re.sub(r"[ \t]+", " ", p).strip(),
            })
    return units


# ── DOCX 파서 ────────────────────────────────────────────────────────
def classify_head(text: str, style: str) -> str:
    """표제를 apparatus / unit / label / structure 로 가른다."""
    t = text.strip()
    if RE_NOTE_HEAD.search(t):
        return "note"
    if t in LABEL_HEADS or label_key(t) in LABEL_HEADS:
        return "label"
    if RE_UNIT_HEAD.match(t):
        return "unit"
    if RE_APPARATUS_HEAD.search(t) or RE_DOC_OUTLINE.match(t):
        return "apparatus"
    return "structure"


def head_level(style: str) -> int:
    m = re.search(r"(\d+)", style or "")
    return int(m.group(1)) if m else 2


def parse_source_docx(path: Path):
    """원문만 담긴 docx(회본 등)를 원문 단위로 읽는다.
    한문 단락만 취하고, 편집 범례·과단 표제 같은 한국어 줄은 버린다."""
    if docx is None:
        raise RuntimeError("python-docx 가 필요합니다: pip install python-docx")
    d = open_docx(path)
    units, order, cur_marker = [], 0, None
    for p in d.paragraphs:
        txt = re.sub(r"[ \t]+", " ", p.text).strip()
        if not txt or not is_source_line(txt):
            continue
        # 【科層】·【科文】·【本文】·【註】 같은 편집 표지로 시작하는 줄은 구조 표시이지
        # 대장경 본문이 아니므로 원문 단위로 세지 않는다
        if re.match(r"^【(?:科層|科文|本文|註|論|疏)】", txt):
            continue
        pieces = [x.strip() for x in
                  re.split(r"(?=\[\d{3,4}[abc]\d{2}\])", txt) if x.strip()]
        for piece in pieces:
            pm = RE_MARKER.match(piece)
            if pm:
                cur_marker = pm.group(1)
            order += 1
            units.append({"i": order, "m": pm.group(1) if pm else cur_marker,
                          "cn": piece})
    return units


PURE_MARKER = re.compile(r"^\[(\d{3,4}[abc]\d{2})\]$")


# ── docx 문단 펴기 ──────────────────────────────────────────────
# 번역 docx 가운데는 한 문단 안에 줄바꿈만으로
#   [0377a05] ⏎ 원문 ⏎ <한문> ⏎ 한국어 번역 ⏎ <번역>
# 을 몰아 담는 형식이 있다. 이대로는 원문과 번역이 한 덩어리로 붙어
# 짝이 지어지지 않으므로, 줄마다 따로 선 문단인 것처럼 펴 준다.
INNER_LABEL = re.compile(r"^(?:원문|한국어(?:\s*(?:직역|번역))?|번역|직역)$")


class _Para:
    """python-docx 문단과 같은 모양(text·style.name)을 가진 대역."""

    __slots__ = ("text", "style")

    def __init__(self, text, style_name):
        self.text = text
        self.style = type("S", (), {"name": style_name})()


class _Doc:
    """문단만 펴 놓은 docx 대역. tables 등은 원본을 그대로 넘긴다."""

    def __init__(self, doc, paras):
        self._doc = doc
        self.paragraphs = paras

    def __getattr__(self, name):
        return getattr(self._doc, name)


def open_docx(path: Path):
    d = docx.Document(str(path))
    out, split_any = [], False
    for p in d.paragraphs:
        lines = [x.strip() for x in p.text.split("\n")]
        lines = [x for x in lines if x]
        # 안쪽에 '원문'·'한국어 번역' 같은 칸 이름이 줄로 서 있을 때만 편다
        if len(lines) > 1 and any(INNER_LABEL.match(x) for x in lines):
            split_any = True
            style = (p.style.name or "").strip()
            for x in lines:
                if INNER_LABEL.match(x):
                    continue
                out.append(_Para(x, style))
        else:
            out.append(p)
    return _Doc(d, out) if split_any else d


def is_translation_only(path: Path) -> bool:
    """원문 없이 번역만 담긴 docx 인가.

    한문 원문 단락이 사실상 없고, '[0459a11]' 처럼 표지만 홀로 선 문단이
    여럿이면 번역 전용 형식으로 본다. 이런 문서는 sources/<id>/원문*.txt
    쪽에서 원문을 따로 대야 하며, merge() 의 표지 대응만으로 짝짓는다."""
    d = open_docx(path)
    paras = [re.sub(r"[ \t]+", " ", p.text).strip() for p in d.paragraphs]
    paras = [t for t in paras if t]
    if not paras:
        return False
    n_src = sum(1 for t in paras if is_source_line(t))
    n_anchor = sum(1 for t in paras if PURE_MARKER.match(t))
    return n_anchor >= 20 and n_src < len(paras) * 0.03


def parse_translation_only_docx(path: Path):
    """표지 단독 문단을 앵커로 삼아 번역만 파싱한다(원문 문단 없음).

    문서는 앞부분(해제) → [표지] 번역… [표지] 번역… → 뒷부분(검증 보고 등)
    순서로 구성된다고 본다. 표지가 처음 나오는 순간부터 '본문'으로 보고,
    표지가 나온 뒤에 다시 나오는 1단계 표제(Heading 1)는 본문이 끝나고
    뒷부분(부록)이 시작된 것으로 본다."""
    d = open_docx(path)
    units, front, appendix = [], [], []
    mode = "front"          # front | body | back
    cur = None

    for p in d.paragraphs:
        txt = re.sub(r"[ \t]+", " ", p.text).strip()
        if not txt:
            continue
        style = (p.style.name or "").strip()

        m = PURE_MARKER.match(txt)
        if m:
            if mode == "back":          # 뒷부분 이후에 표지가 다시 나올 리 없지만 방어적으로
                appendix.append(txt)
                continue
            if cur:
                units.append(cur)
            cur = {"m": m.group(1), "cn": [], "ko": [], "nt": []}
            mode = "body"
            continue

        if mode == "body" and style.lower().startswith("heading"):
            if cur:
                units.append(cur)
                cur = None
            mode = "back"

        if mode == "front":
            front.append(txt)
        elif mode == "back":
            appendix.append(txt)
        elif cur is not None:
            is_note = ("audit" in style.lower()) or ("검증" in style) or ("교감" in style)
            (cur["nt"] if is_note else cur["ko"]).append(txt)

    if cur:
        units.append(cur)

    return {"units": units, "tables": [], "front": front,
            "appendix": appendix, "sections": []}


def is_table_aligned(path: Path) -> bool:
    """대조 본문이 표(위치표지 | 원문 | 번역)로 짜인 docx 인가."""
    if docx is None:
        return False
    d = open_docx(path)
    for t in d.tables:
        if len(t.columns) < 3 or len(t.rows) < 3:
            continue
        head = [c.text.strip() for c in t.rows[0].cells]
        joined = " ".join(head)
        # 번역 칸이 없으면 본문 대조표가 아니다.
        # (‘위치표지 | 원문 | 유형 | 처리’ 같은 교감 목록을 본문으로 오인하지 않도록)
        has_ko = any(k in joined for k in ("번역", "한국어", "직역", "국역"))
        if "위치표지" in joined and "원문" in joined and has_ko:
            return True
        # 표제가 없어도 첫 칸이 위치표지 꼴이고, 어느 칸엔가
        # 한국어 본문이 들어 있으면 대조표로 본다
        if PURE_MARKER.match(head[0] if head else ""):
            body = " ".join(c.text for r in t.rows[:6] for c in r.cells)
            if hangul_ratio(body) > 0.15:
                return True
    return False


def parse_table_docx(path: Path):
    """표로 짜인 대조본을 읽는다.

    각 행이 (위치표지, 원문, 번역) 세 칸으로 이루어진 표를 본문으로 삼고,
    두 칸짜리 표는 서지·용어표로 넘긴다. 표 바깥 문단은 해제로 모은다."""
    if docx is None:
        raise RuntimeError("python-docx 가 필요합니다: pip install python-docx")
    d = open_docx(path)

    units, tables, front = [], [], []

    for t in d.tables:
        rows = [[c.text.strip() for c in r.cells] for r in t.rows]
        rows = [r for r in rows if any(r)]
        if not rows:
            continue
        head = " ".join(rows[0])
        body_table = (len(rows[0]) >= 3 and
                      (("위치표지" in head and "원문" in head)
                       or PURE_MARKER.match(rows[0][0])))
        if not body_table:
            tables.append(rows)          # 서지·용어표
            continue

        start = 1 if ("위치표지" in head or "원문" in head) else 0
        for r in rows[start:]:
            mk_cell, cn_cell = r[0].strip(), r[1].strip()
            ko_cell = r[2].strip() if len(r) > 2 else ""
            m = PURE_MARKER.match(mk_cell) or RE_MARKER.search(mk_cell)
            marker = m.group(1) if m else None
            if not cn_cell and not ko_cell:
                continue
            cn = [x.strip() for x in cn_cell.split("\n") if x.strip()]

            # 번역 칸 안에 '주' 줄이 있고 그 아래 번호 각주가 이어지는 형식이 있다.
            # 줄바꿈으로만 쪼개면 각주가 번역문으로 섞이므로 여기서 갈라낸다.
            ko, nt, in_note = [], [], False
            for x in ko_cell.split("\n"):
                x = x.strip()
                if not x:
                    continue
                if len(x) <= 8 and RE_NOTE_HEAD.search(x):
                    in_note = True
                    continue
                (nt if in_note else ko).append(x)
            units.append({"m": marker, "cn": cn, "ko": ko, "nt": nt})

    for p in d.paragraphs:
        txt = re.sub(r"[ \t]+", " ", p.text).strip()
        if txt and not is_source_line(txt):
            front.append(txt)

    return {"units": units, "tables": tables,
            "front": front, "appendix": [], "sections": []}


# ── KABC(한국불교전서) 번역 docx ───────────────────────────────────
# CBETA 계열과 달리 원문 txt 는 15~16자마다 줄이 끊겨 있어 그대로 실을 수
# 없다. 번역 docx 가 이미 문단으로 이어 붙인 원문을 담고 있으므로,
# 이 갈래는 docx 하나만으로 본문을 세운다.
#
# 블록 차례
#   [저본 위치: 卷上 第2張]  →  원문  →  한국어 직역  →  교감주
RE_KABC_LOC = re.compile(r"^\[저본\s*위치\s*[:：]\s*(.+?)\]\s*$")
RE_KABC_JANG = re.compile(
    r"(卷[上中下一二三四五六七八九十\d]*)?\s*第\s*([一二三四五六七八九十百○\d]+)\s*張")
RE_NOTE_START = re.compile(r"^\[(?:교감|KABC|번역자|위치표지|편집)[^\]]*\]")
# 각주에서 요지로 삼을 줄
RE_NOTE_GIST = re.compile(r"^(?:KABC\s*)?교감\s*내용\s*번역\s*[:：]\s*")
# 각주에서 빼는 줄 — KABC 편집자의 교감문을 그대로 옮긴 대목과,
# 본문을 되풀이하는 대목이다.
RE_NOTE_DROP = re.compile(
    # 줄머리에 '[교감 3)]' 같은 표지가 붙어 있어도 함께 본다
    r"^(?:\[[^\]]*\]\s*)?(?:"
    # ① KABC 편집자의 교감문을 그대로 옮긴 줄
    r"KABC\s*(?:교감\s*원문|저본[·\s]*편집)\s*[:：])|"
    # ② 본문·번역을 되풀이하는 줄 (화면에 이미 있는 것)
    r"^(?:\[[^\]]*\]\s*)?(?:저본|이문)[·\s]*(?:추정)?\s*"
    r"(?:적용[^:：]{0,10}|기준[^:：]{0,12}|번역)\s*[:：]")

HAN_NUM = {"○": 0, "零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
           "六": 6, "七": 7, "八": 8, "九": 9}


def han_to_int(s: str):
    """한자 숫자를 아라비아 숫자로. '第一○張'·'第二十八張' 둘 다 읽는다."""
    s = s.strip()
    if s.isdigit():
        return int(s)
    if not s or any(c not in HAN_NUM and c not in "十百" for c in s):
        return None
    # 자릿수 표기(二十八)와 나열 표기(一○) 를 함께 다룬다
    if "十" in s or "百" in s:
        total, cur = 0, 0
        for c in s:
            if c == "十":
                total += (cur or 1) * 10
                cur = 0
            elif c == "百":
                total += (cur or 1) * 100
                cur = 0
            else:
                cur = HAN_NUM.get(c, 0)
        return total + cur
    n = 0
    for c in s:
        n = n * 10 + HAN_NUM.get(c, 0)
    return n


def kabc_marker(text: str):
    """'[저본 위치: 卷上 第一張]' 의 속살을 '卷上第1張' 꼴로 고른다.

    같은 문헌 안에서도 '第一張' 과 '第2張' 이 섞여 나오므로 숫자를
    아라비아로 통일한다. 표지가 없는 문헌은 None 을 돌려준다."""
    if "없음" in text:
        return None
    m = RE_KABC_JANG.search(text)
    if not m:
        return None
    juan, num = m.group(1) or "", m.group(2)
    n = han_to_int(num)
    if n is None:
        return None
    return f"{juan}第{n}張"


def split_kabc_note(lines):
    """각주 한 덩이를 (요지, 상세) 로 가른다.

    화면에는 요지만 내고, 상세는 손을 얹었을 때 뜨게 한다."""
    keep = [x for x in lines if not RE_NOTE_DROP.match(x)]
    gist = ""
    for x in keep:
        if RE_NOTE_GIST.match(x):
            gist = RE_NOTE_GIST.sub("", x).strip()
            break
    if not gist:
        head = keep[0] if keep else (lines[0] if lines else "")
        gist = re.sub(r"^\[[^\]]*\]\s*", "", head).strip() or head
    # 상세는 손을 얹었을 때 뜨는 쪽지다. '저본 독법'처럼 본문을 통째로
    # 되풀이하는 줄이 있어, 읽을 만한 길이로 잘라 둔다.
    # (원문 전체는 이미 본문 칸에 있으므로 잃는 것이 없다)
    detail = []
    for x in keep:
        x = x.strip()
        if not x or x == gist:
            continue
        if len(x) > 140:
            x = x[:138].rstrip() + "…"
        detail.append(x)
    return gist, detail


def parse_kabc_docx(path: Path):
    """KABC 번역 docx 하나로 본문 단위를 세운다."""
    if docx is None:
        raise RuntimeError("python-docx 가 필요합니다: pip install python-docx")
    d = open_docx(path)
    units, front, cur, marker = [], [], None, None
    note_buf, in_body = [], False

    def flush_note():
        nonlocal note_buf
        if cur is not None and note_buf:
            gist, detail = split_kabc_note(note_buf)
            if gist:
                cur["nt"].append({"t": gist, "d": detail} if detail else gist)
        note_buf = []

    for p in d.paragraphs:
        txt = re.sub(r"[ \t]+", " ", p.text).strip()
        if not txt:
            continue
        style = (p.style.name or "").strip()

        loc = RE_KABC_LOC.match(txt)
        if loc:
            flush_note()
            marker = kabc_marker(loc.group(1))
            in_body = True
            continue
        if style == "Source Text":
            flush_note()
            cur = {"i": len(units) + 1, "m": marker, "cn": [txt],
                   "ko": [], "nt": []}
            units.append(cur)
            in_body = True
            continue
        if style == "Translation Text":
            flush_note()
            if cur is not None:
                cur["ko"].append(txt)
            continue
        if style == "Editorial Note":
            if RE_NOTE_START.match(txt) or not note_buf:
                flush_note()
            note_buf.extend(x.strip() for x in txt.split("\n") if x.strip())
            continue
        if not in_body and style not in LABELS:
            front.append(txt)
    flush_note()

    # 위치표지가 아예 없는 문헌은 단위 순번을 표지로 삼는다
    if units and not any(u["m"] for u in units):
        for u in units:
            u["m"] = None
    return {"units": units, "front": front, "appendix": [], "tables": []}


# ── KABC(한국불교전서) 번역본 ─────────────────────────────────
# KABC 저본은 15~16자마다 줄을 끊어 두어 그대로 실을 수 없다.
# 대신 번역 docx 가 문단으로 묶어 둔 원문을 본문으로 삼는다.
# 문단 차례는 다음과 같다.
#   [저본 위치: 卷上 第1張] / 원문 / <한문> / 한국어 직역 / <번역> / 교감주 / <각주>
KABC_LOC = re.compile(r"^\[저본\s*위치\s*[:：]\s*(.*?)\]\s*$")
KABC_LAYER = re.compile(r"^\[문헌\s*층위\s*[:：]")
KABC_LABELS = {"원문", "한국어 직역", "한국어 번역", "교감주", "기타 각주",
               "주석", "번역", "직역"}
KABC_CONT = re.compile(r"^(?:원문|한국어\s*직역|번역)\s*[(（]\s*이어짐\s*[)）]")
# KABC 저본 TXT 는 교감문을 본문 줄 사이에 그대로 끼워 둔다.
# 번역 docx 를 만들 때 그것을 본문으로 잘못 읽으면 교감문이 본문 단위로
# 올라온다. 아래 두 가지 신호로 그런 대목을 붙잡아 앞 단위의 각주로 돌린다.
#   1) 라벨이 그냥 '원문' 이 아니라 'KABC 교감 원문' 처럼 교감을 밝힐 때
#   2) [문헌 층위: … 교감층] 처럼 층위 자체가 교감이라고 밝힐 때
KABC_APPARATUS_LABEL = re.compile(r"(?:교감|편집)\s*(?:원문|내용)|이문\s*원문")
KABC_APPARATUS_LAYER = re.compile(r"교감층|교감\s*주$|편집층")
# 판본기호. 본문에는 결코 나오지 않고 교감문에만 나온다.
RE_PANBON = re.compile(r"\{[底甲乙丙丁戊己編校]\}")
# KABC TXT 는 맨 끝에 출처 한 줄을 붙인다.
#   『기신본말오중』 起信本末五重(ABC, H0320 v12, p.850c01-853b25)
# 본문도 번역도 아니므로 버린다.
RE_ABC_TAIL = re.compile(r"\(\s*ABC\s*,\s*H\d{3,4}\s*v\d+")
# 산문 뒤에 게송이 이어짐을 알리는 말머리
RE_GATHA_CUE = re.compile(r"(?:頌|偈)(?:曰|云|言)")
# 각주에서 화면에 늘 보일 줄(요지)과, 얹었을 때만 보일 줄(상세)을 가른다.
NOTE_SUMMARY_KEYS = ("교감 내용 번역", "KABC 교감 내용 번역", "교감 내용 한국어 번역",
                     "번역", "교감문 번역", "교감 원문 번역", "한국어",
                     "교감 내용", "내용 번역", "한국어 번역", "교감 번역")
CJK_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7,
           "八": 8, "九": 9, "十": 10, "○": 0, "零": 0}


def cjk_int(s: str):
    """'二八' · '一○' · '三十八' 같은 한자 숫자를 아라비아 숫자로."""
    s = s.strip()
    if s.isdigit():
        return int(s)
    if not s or any(c not in CJK_NUM for c in s):
        return None
    if "十" in s:                       # 三十八 · 十六 · 二十
        a, _, b = s.partition("十")
        return (CJK_NUM[a] if a else 1) * 10 + (CJK_NUM[b] if b else 0)
    if len(s) > 1:                      # 一○ · 二八 처럼 자리마다 적은 꼴
        n = 0
        for c in s:
            n = n * 10 + CJK_NUM[c]
        return n
    return CJK_NUM[s]


CJK_TEN = "一二三四五六七八九"


def han_juan(n: int) -> str:
    """1~99 를 '一'·'十'·'一十九' 아닌 '十九' 꼴의 한자 숫자로."""
    if n <= 0:
        return ""
    if n < 10:
        return CJK_TEN[n - 1]
    tens, ones = divmod(n, 10)
    head = "十" if tens == 1 else CJK_TEN[tens - 1] + "十"
    return head + (CJK_TEN[ones - 1] if ones else "")


def norm_juan(juan: str) -> str:
    """'卷一○'처럼 자리마다 적은 권차를 '卷十' 꼴로 고른다.

    '卷上'·'卷下'처럼 숫자가 아닌 권차는 손대지 않는다."""
    if not juan.startswith("卷"):
        return juan
    body = juan[1:]
    if not body or any(c not in CJK_NUM for c in body):
        return juan
    n = cjk_int(body)
    return f"卷{han_juan(n)}" if n else juan


def norm_kabc_loc(raw: str) -> str:
    """저본 위치를 '卷上 第1張' 꼴로 고른다."""
    t = re.sub(r"\s+", " ", raw).strip()
    if not t or "없음" in t:
        return ""
    # 권차에 '卷一○'(권10)처럼 ○ 가 섞여 오므로 ○·零 도 권차로 읽는다.
    m = re.search(r"(卷[上中下一二三四五六七八九十○零\d]*)?\s*第\s*"
                  r"([一二三四五六七八九十○零\d]+)\s*張", t)
    if not m:
        # '序 — 張 위치표지 이전' · '권두 서문·서례·범례' 처럼 설명이 붙은 것은
        # 첫 마디만 남긴다. 좌측 여백이 좁아 길면 본문 위로 넘친다.
        head = re.split(r"\s*[—–-]\s*", t)[0].strip()
        head = re.split(r"\s+", head)[0][:8]
        # '張 표지 이전' 처럼 앞머리에 남을 말이 없으면 '권두'로 적는다
        if head in ("張", "장", "") and "표지" in t:
            return "권두"
        return head
    n = cjk_int(m.group(2))
    juan = norm_juan((m.group(1) or "").strip())
    return f"{juan} 第{n}張".strip() if n is not None else t


# 항목을 줄바꿈이 아니라 두 칸 이상 띄어쓰기로 잇는 docx 가 있다.
#   [교감] 2)干  KABC: 「干」作「于」{甲}  한국어: 甲본에서는 …  저본 독법: …
# 그런 덩이는 항목 이름 앞에서 줄을 나눠 준다.
RE_NOTE_ITEM = re.compile(r" {2,}(?=[가-힣A-Za-z][^:：\n]{0,14}[:：])")
# 항목을 한 칸 띄어쓰기만으로 잇는 docx 도 있다. 그때는 정해진 항목 이름
# 앞에서만 줄을 나눈다. ('KABC:' 는 첫 줄에 남겨 둔다)
RE_NOTE_ITEM1 = re.compile(
    r"\s+(?=(?:저본|이문|교감|한국어|의미|차이|판정|보충)"
    r"[^:：\n\[\]]{0,18}[:：])")
# '[교감] KABC: 「感」底本傍註加「此」。 저본의 「感」 곁주에는…' 처럼 교감 원문
# 뒤에 한국어 풀이가 바로 붙는 꼴. 한국어가 시작되는 자리를 찾는다.
RE_KABC_GIST = re.compile(r"[。．.]\s*")
# 각주 한 덩이의 시작. '[교감] …' 뿐 아니라 '2) [KABC 편집] …' 꼴도 새 각주다.
RE_NOTE_OPEN = re.compile(r"^\s*(?:\d{1,3}\s*[).）]\s*)?\[")


def split_kabc_note(text: str):
    """각주 한 덩이를 (요지, 상세)로 가른다."""
    if "\n" not in text:
        text = RE_NOTE_ITEM.sub("\n", text)
    if "\n" not in text:
        text = RE_NOTE_ITEM1.sub("\n", text)
    lines = [x.strip() for x in text.split("\n") if x.strip()]
    if not lines:
        return "", ""
    head, summary, detail = lines[0], "", []
    for ln in lines[1:]:
        key, _, val = ln.partition(":")
        if not val:
            key, _, val = ln.partition("：")
        # '• 교감 내용 번역: …' 처럼 글머리표를 앞세우는 문서가 있다
        key = re.sub(r"^\s*[•·‣▪◦\-–—*]\s*", "", key)
        if val and key.strip() in NOTE_SUMMARY_KEYS and not summary:
            summary = val.strip()
        else:
            detail.append(ln)
    tag = re.search(r"\[[^\]]{1,16}\]", head[:28])
    if not summary and "KABC" in head:
        # 요지로 올릴 항목이 없을 때, 교감 원문 뒤에 바로 붙은 한국어 풀이를
        # 요지로 세우고 한문 쪽은 상세로 내린다. 마침표마다 끊어 보며
        # 뒤쪽이 한국어 문장이 되는 첫 자리를 고른다.
        for m2 in RE_KABC_GIST.finditer(head):
            tail = head[m2.end():].strip()
            if len(tail) >= 6 and hangul_ratio(tail) > 0.35:
                summary, head = tail, head[:m2.end()].strip()
                break
    if summary:
        # '[교감]' · '2) [KABC 편집]' 앞머리의 표지를 요지에 살려 둔다
        summary = f"{tag.group(0)} {summary}" if tag else summary
        if head:
            detail.insert(0, head)
    else:
        summary = head
    return summary, "\n".join(detail)


def read_docx_tables(path: Path):
    """docx 안의 표를 (행 × 칸) 목록으로 읽는다. 서지·용어표 추출용."""
    if docx is None:
        return []
    d = open_docx(path)
    out = []
    for t in d.tables:
        rows = [[c.text.strip() for c in r.cells] for r in t.rows]
        if rows:
            out.append(rows)
    return out


def is_kabc_docx(path: Path) -> bool:
    if docx is None:
        return False
    d = open_docx(path)
    for p in d.paragraphs[:400]:
        if KABC_LOC.match(p.text.strip()):
            return True
    return False


def parse_kabc_docx(path: Path):
    """KABC 번역본을 읽는다. 원문·번역·각주를 모두 이 파일에서 얻는다."""
    d = open_docx(path)
    units, front, secs = [], [], []
    loc, cont, started = "", False, False
    note_open = False
    apparatus = False        # 지금 읽는 대목이 본문이 아니라 교감층인가
    skip_tail = False        # KABC 출처 표시줄과 그에 딸린 문단을 흘려보낸다

    for p in d.paragraphs:
        raw = p.text
        t = re.sub(r"[ \t]+", " ", raw).strip()
        if not t:
            continue
        style = (p.style.name or "").strip()

        m = KABC_LOC.match(t)
        if m:
            raw = m.group(1)
            loc = norm_kabc_loc(raw)
            # '卷下 第38張 후속 KABC 교감주' 처럼 위치 자체가 교감을 밝히기도 한다
            apparatus = bool(KABC_APPARATUS_LAYER.search(raw)
                             or "교감" in raw)
            skip_tail = False
            continue
        if KABC_LAYER.match(t):
            apparatus = bool(KABC_APPARATUS_LAYER.search(t))
            continue
        if KABC_CONT.match(t):
            cont = True
            continue
        if t in KABC_LABELS:
            note_open = t in ("교감주", "기타 각주", "주석")
            continue
        if KABC_APPARATUS_LABEL.search(t) and len(t) <= 20:
            # 'KABC 교감 원문' · '교감 내용 한국어 번역' 같은 라벨.
            # 이 아래의 원문·번역 문단은 본문이 아니라 교감이다.
            apparatus = True
            continue

        if style == "Source Text":
            if RE_ABC_TAIL.search(t):
                # KABC 출처 표시줄. 본문이 아니므로 버리고,
                # 뒤따르는 번역·검토주도 함께 흘려보낸다.
                skip_tail = True
                continue
            if apparatus and units:
                # 교감문이 본문 자리에 실려 왔다. 앞 단위의 각주로 돌린다.
                units[-1]["nt"].append(t)
                units[-1]["ntd"].append("")
                cont = False
                continue
            if cont and units:
                units[-1]["cn"].append(t)
            else:
                units.append({"i": len(units), "m": loc or None,
                              "cn": [t], "ko": [], "nt": [], "ntd": []})
            cont, started, note_open = False, True, False
            continue
        if style == "Translation Text":
            if skip_tail:
                continue
            if apparatus and units and units[-1]["nt"]:
                # 교감문의 한국어 번역. 요지로 앞에 세우고 원문은 상세로 내린다.
                units[-1]["ntd"][-1] = units[-1]["nt"][-1]
                units[-1]["nt"][-1] = t
                cont = False
                continue
            if units:
                units[-1]["ko"].append(t)
            cont = False
            continue
        # 각주 문단의 스타일 이름은 문서마다 다르다.
        # (Editorial Note · Critical Note · 교감주 …)
        # 다만 'Front Note' 는 본문 앞 해제이므로 각주로 보지 않는다.
        is_note = (style.endswith("Note") or "교감주" in style) and style != "Front Note"
        if is_note and skip_tail:
            continue
        if is_note and units:
            # 대괄호 표지로 시작하면 새 각주, 아니면 앞 각주에 이어진다
            if RE_NOTE_OPEN.match(t) or not units[-1]["nt"]:
                a, b = split_kabc_note(raw.strip())
                units[-1]["nt"].append(a)
                units[-1]["ntd"].append(b)
            else:
                a, b = split_kabc_note(raw.strip())
                joined = "\n".join(x for x in (units[-1]["ntd"][-1], a, b) if x)
                units[-1]["ntd"][-1] = joined
            continue

        if style.startswith("Heading") and started:
            secs.append({"lv": 2 if style.endswith("2") else 1,
                         "t": t, "at": len(units)})
        elif not started:
            front.append(t)

    # 각주는 여러 문단에 나뉘어 오기도 한다. 덩어리를 다 모은 뒤에
    # 요지와 상세로 갈라, 화면이 바로 쓸 수 있는 꼴로 담는다.
    #   문자열       → 예전처럼 한 줄로 보인다
    #   {t:…, d:[…]} → 요지만 보이고 상세는 얹거나 누르면 편다
    for u in units:
        out = []
        for k in range(len(u["nt"])):
            whole = "\n".join(x for x in (u["nt"][k], u["ntd"][k]) if x)
            gist, detail = split_kabc_note(whole)
            lines = [x.strip() for x in detail.split("\n") if x.strip()]
            out.append({"t": gist, "d": lines} if lines else gist)
        u["nt"] = out
        u.pop("ntd", None)

    # 위치표지가 아예 없는 문헌은 문단 순번을 표지로 삼는다
    if not any(u["m"] for u in units):
        for k, u in enumerate(units, 1):
            u["m"] = f"{k}"
    return {"units": units, "front": front, "secs": secs}


# 본문 문장 사이에 낀 도판. 화면에서 그 자리에 그대로 끼워 넣는다.
FIG_TOKEN = "\u27e6fig:{}\u27e7"
# 번역본을 만든 쪽이 앞머리에 붙인 제작 메모. 저본의 글이 아니다.
RE_MAKER_NOTE = re.compile(
    r"CBETA|첨부\s*TXT|전수\s*교열|대응\s*검증|전문\s*완역|전문\s*번역"
    r"|최종\s*복원본|최종\s*완역본|재배열한|직역\s*대조본|대조\s*번역본"
    r"|위치\s*표지를\s*기준|U\s*단위|전수\s*대응|전수\s*처리|전수\s*대조")
RE_FIG_NAME = re.compile(r"\[\s*([A-Za-z]\d{2}p\d{4}_\d{2})\.(?:gif|git|jpg|png)\s*\]",
                         re.I)


def docx_image_names(doc, figdir: Path):
    """docx 안에 박아 둔 그림을 원본 파일 이름으로 되돌린다.

    docx 는 그림을 image1.gif 처럼 이름을 바꿔 담으므로, 바이트를 견주어
    assets/figures/<문헌id>/ 의 원본 이름을 찾는다."""
    import hashlib
    if not figdir.is_dir():
        return {}
    want = {}
    for f in sorted(figdir.iterdir()):
        if f.is_file():
            want[hashlib.md5(f.read_bytes()).hexdigest()] = f.name
    out = {}
    for rid, part in doc.part.related_parts.items():
        try:
            blob = part.blob
        except Exception:
            continue
        name = want.get(hashlib.md5(blob).hexdigest())
        if name:
            out[rid] = name
    if out:
        return out

    # docx 를 만들 때 그림을 png 로 바꿔 담은 문서가 있다. 바이트가 달라
    # 이름을 못 찾으므로, 본문에 나온 차례대로 원본 파일에 맞춘다.
    # (저본의 도판 번호도 본문 차례를 따르므로 둘이 어긋나지 않는다)
    from docx.oxml.ns import qn
    order = []
    for para in doc.paragraphs:
        for node in para._p.iter():
            if node.tag.split("}")[-1] == "blip":
                rid = node.get(qn("r:embed"))
                if rid and rid not in order:
                    order.append(rid)
    names = sorted(want.values())
    if order and len(order) == len(names):
        return dict(zip(order, names))
    return {}


def para_text_with_figs(p, imgmap):
    """문단 글을 읽되, 중간에 낀 그림을 토큰으로 바꿔 자리를 지킨다."""
    if not imgmap:
        return p.text
    from docx.oxml.ns import qn
    buf, seen = [], False
    for node in p._p.iter():
        tag = node.tag.split("}")[-1]
        if tag == "t":
            buf.append(node.text or "")
        elif tag == "blip":
            name = imgmap.get(node.get(qn("r:embed")))
            if name:
                buf.append(FIG_TOKEN.format(name))
                seen = True
    return "".join(buf) if seen else p.text


# ── 신한글대장경 전문번역 표준 DOCX(SHTK 스타일) ─────────────────────────
# 「SHTK Source Text / Translation Text / Note Summary / Note Detail /
#  Position Marker / Block Label」처럼 문단 스타일이 곧 역할을 말해 주는 문서.
# 글자 비율로 원문·번역을 추정하지 않고 스타일을 그대로 따른다.
#   · 「9. 본문」 표제 앞은 해제(front), 본문 뒤의 제1표제(검증 정보 등)는 부록
#   · Note Summary 한 줄 + Note Detail 여러 줄 → 교감 하나 {"t": 요지, "d": [상세…]}
#   · Heading 2·3·4 와 SHTK Subheading 은 절 표제(층위 1~4)
RE_SHTK_BODY = re.compile(r"본문\s*$")
# DOCX 제작 단계의 처리 통계 블록(교감이 아님)
RE_SHTK_LOG = re.compile(
    r"공통으로 존재하는 교감\s*:|표시 단위\s*:\s*\d|미확보 이미지 수\s*:|번역 차이 없음\s*:\s*\d"
    r"|교감 상호참조 표지\s*:\s*\d")
RE_KO_FOOTNUM = re.compile(r"(?<!No)([다.,’”」\)!?])\s?(\d{1,3})(?=[\s‘“「(]|$)")
RE_KO_APP_TAIL = re.compile(r"\s*(?:\[(?:\d{1,2}|＊)\]\s*)+$")
RE_KO_FIG_TAIL = re.compile(r"\s*\[이미지\s*확인\]\s*(?:\u27e6fig:[^\u27e7]+\u27e7\s*)+$")
RE_SHTK_SUMMARY = re.compile(r"^\s*교감\s*요약\s*[|｜]\s*")
RE_SHTK_ITEM = re.compile(r"^\s*[•·\-–]\s*")


def is_shtk_docx(d) -> bool:
    return any((p.style.name or "").startswith("SHTK ") for p in d.paragraphs)


# ── 표준 DOCX 교감 메모를 앱 표기(「[CBETA 교감] 요지」 + 펼침 상세)로 맞추기 ──────
# 문서마다 교감 요약 줄이 「[CBETA 교감 0538001]」「[대정장 원교감주 0544001 복원] 「…」」
# 「[star reference] …」처럼 제각각이면 앱 목록이 들쭉날쭉하다. 번호·원교감 원문은
# 상세로 내리고, 요지 줄은 저본·이문 독법으로 한 줄에 보이게 한다.
RE_NT_CBETA = re.compile(
    r"^\[(CBETA\s*교감|대정장\s*원교감주)\s+(\d{7})[^\]]*\]\s*(.*)$", re.S)
RE_NT_ADD = re.compile(r"^\[CBETA\s*추가\s*교감\s+([^\]·]+?)\s*(?:·\s*([^\]]+))?\]\s*(.*)$", re.S)
RE_NT_STAR = re.compile(r"^\[star reference\]\s*(.*)$", re.S | re.I)
RE_NT_GAIJI = re.compile(r"^\[외자 복원\s*(CB\d+)\]\s*(.*)$", re.S)


RE_NT_LONG = re.compile(r"^(\[[^\]]*(?:교감|관주|외자|서지|주기|간기)[^\]]*\])\s*(.+)$", re.S)
RE_NT_TAILTAG = re.compile(r"\s*\[[^\]\[]*/[^\]\[]*\]\s*$")
RE_NT_SENT = re.compile(r"(?<=[가-힣’”」)】\]\u3400-\u9fff\U00020000-\U0003ffff])\.\s+")


def _nt_sentences(t):
    t = t.replace("`", "").strip()
    return [x.strip() for x in re.split(r"(?<=[다」])\.\s+", t) if x.strip()]


def _nt_gist(d):
    kv = dict(x.split(": ", 1) for x in d if ": " in x)
    a, b = kv.get("저본 독법"), kv.get("이문 독법")
    if not (a and b):
        return None
    # 첫 문장만 쓰고, 뜻풀이 괄호(“…”)는 뺀다
    drop = lambda v: re.sub(r"\s*\(“[^)]*”\)", "", re.split(r"\.\s", v)[0]).strip()
    a, b = drop(a), drop(b)
    m = re.match(r"^(.*?)\s*\(\[([^\]]+)\](?:,\s*(.*))?\)$", b)
    if m:
        return f"저본 {a} → {m.group(2)}본 {m.group(1)}" + (f" ({m.group(3)})" if m.group(3) else "")
    return f"저본 {a} → 이문 {b}"


# ── 여러 표준 DOCX 교감 메모 변형을 「[태그] 요지 + 상세」로 맞추기 ─────────────
# (화엄경탐현기처럼 「[CBETA 원교감] [2] · n=0107002 · 悕＝希【甲】 · 현행 …」로
#  한 줄에 이어 쓴 것, 「[CBETA 교감 0116c07]」처럼 번호만 요지에 둔 것 등)
RE_WIT = re.compile(r"【([^】]+)】")


def _orig_gist(o):
    """대정 원교감 표기(悕＝希【甲】, 〔X〕－【甲】, A＋（B）【甲】)를 한 줄 요지로."""
    o = o.strip().rstrip("＊*").strip()
    wits = "·".join(w for w in RE_WIT.findall(o) if w != "大")
    body = RE_WIT.sub("", o).replace("＊", "").strip()
    mark = ""
    for k, v in (("ィ", " (일본)"), ("ヵ", " (의심)")):
        if k in body:
            body = body.replace(k, ""); mark = v
    w = f"{wits}본" if wits else "이본"
    m = re.fullmatch(r"(.+?)＝(.+)", body)
    if m:
        return f"저본 {m.group(1)} → {w} {m.group(2)}{mark}"
    m = re.fullmatch(r"〔(.+?)〕－", body)
    if m:
        return f"{w}에는 {m.group(1)} 없음{mark}"
    m = re.fullmatch(r"(.+?)＋（(.+?)）", body)
    if m:
        return f"{w}은 {m.group(1)} 뒤에 {m.group(2)}이 있음{mark}"
    m = re.fullmatch(r"（(.+?)）＋(.+)", body)
    if m:
        return f"{w}은 {m.group(2)} 앞에 {m.group(1)}이 있음{mark}"
    return None


RE_NT_DOTS = re.compile(r"^\[(CBETA 원교감(?: · TXT 표지 없음)?)\]\s*(?:\[[\d＊]+\]\s*·\s*)?n=(\d{7})\s*·\s*(.*)$", re.S)
RE_NT_ID = re.compile(
    r"^\[(CBETA (?:교감|현행 추가주|외자(?: 갱신)?|XML 교감 · TXT 표지 없음|현행 star_removed 교감))"
    r"(?:\s+(\d{4}[abc]\d{2}\d*))?\]\s*(?:(\d{4}[abc]\d{2}\d*)\s*·\s*)?(.*)$", re.S)
NT_NOISE = re.compile(r"^(?:이미지 판독:\s*해당 없음.*|판독 확신도:\s*해당 없음\.?|교감 종류:.*)$")
NT_DROP = re.compile(r"^\[CBETA 공식 교감·외자 복원\]$|^\[편집\]\s*(?:\[\d+\](?:,\s*)?)+[은는]?\s*원문의 교감·주기 앵커")


def shtk_variant_note(n):
    t = n if isinstance(n, str) else n["t"]
    d = [] if isinstance(n, str) else list(n["d"])
    if NT_DROP.search(t):
        return None                                   # 출처 한 줄뿐인 빈 블록·앵커 안내
    d = [re.sub(r"^상세:\s*", "", x) for x in d if not NT_NOISE.match(x.strip())]
    m = RE_NT_DOTS.match(t)
    if m:
        parts = [x.strip() for x in m.group(3).split(" · ") if x.strip()]
        dd, orig = [], None
        for x in parts:
            if x.startswith("XML 위치"):
                dd.append("원문 위치: " + x[len("XML 위치"):].strip() + " (원문 TXT에는 표지 없음)")
            elif x.startswith("현행 CBETA:"):
                dd.append(x.replace("현행 CBETA:", "현행 CBETA:", 1))
            elif orig is None:
                orig = x
            else:
                dd.append(x)
        gist = _orig_gist(orig) if orig else None
        head = ["원교감 원문: " + orig] if orig else []
        return {"t": "[CBETA 교감] " + (gist or orig or "교감 " + m.group(2)),
                "d": head + ["교감 번호: " + m.group(2)] + dd + d}
    m = RE_NT_ID.match(t)
    if m:
        kind, nid, rest = m.group(1), m.group(2) or m.group(3), m.group(4).strip()
        rest = re.sub(r"\s*·\s*B단계 판정:\s*\S+", "", rest)
        star = re.search(r"\s*·\s*원 별표 계열 n=(\d+)", rest)
        if star:
            rest = rest[:star.start()]
            d.append("원 별표 계열: 교감 " + star.group(1))
        if kind.startswith("CBETA 외자"):
            kv = {}
            for x in d:
                k2, _, v2 = x.partition(" ")
                kv.setdefault(k2, v2)
            tx = re.search(r"TXT 표기 (\S+)", " ".join(d))
            nf = re.search(r"normalized form (\S+)", " ".join(d))
            # 요지 문장이 이미 있으면(「CB18658의 자형은 …이다」) 그대로 쓰고,
            # 없을 때만 「TXT 표기 = 정규형」을 조립한다
            g = (tx.group(1) + (f" = {nf.group(1)}" if nf and nf.group(1) != "—" else "")) if tx else ""
            tag = "[외자 갱신]" if "갱신" in kind else "[외자]"
            head = rest or g or (d[0] if d else "")
            return {"t": f"{tag} {head}".strip(),
                    "d": (["위치: " + nid] if nid else []) + d}
        if not rest:
            kv = dict(x.split(": ", 1) for x in d if ": " in x)
            a, b = kv.get("저본 독법"), kv.get("이문/현행 독법") or kv.get("이문 독법")
            if a and b:
                rest = f"저본 {RE_WIT.sub('', a)} → 현행 CBETA {RE_WIT.sub('', b).replace('∅', '(없음)')}"
            elif d and "→" in d[0]:
                rest = RE_WIT.sub("", d[0]) + (" (표기 차이)" if any("표기/자형" in x for x in d) else "")
            elif d:
                rest = d[0]
        tag = {"CBETA 현행 추가주": "[CBETA 추가 교감]", "CBETA 현행 star_removed 교감": "[CBETA 추가 교감]",
               "CBETA XML 교감 · TXT 표지 없음": "[CBETA 교감]"}.get(kind, "[CBETA 교감]")
        if kind == "CBETA XML 교감 · TXT 표지 없음":
            rest += " (원문 TXT에는 표지 없음)"
        d = [x for x in d if not x.startswith("교감번호:")]
        return {"t": f"{tag} {rest}".strip(), "d": (["교감 번호: " + nid] if nid else []) + d}
    m = re.match(r"^\[＊ 상호참조( 후보)?\]\s*(.*)$", t, re.S)
    if m:
        body = m.group(2).replace("→ n=", "교감 ").replace(" · ", " — ", 1)
        return (f"[상호참조{' 후보' if m.group(1) else ''}] " + body) if not d else \
            {"t": f"[상호참조{' 후보' if m.group(1) else ''}] " + body, "d": d}
    if isinstance(n, dict):
        return {"t": t, "d": d} if d else t
    return n


def shtk_tidy_note(n):
    if isinstance(n, dict) and not n["d"]:      # 상세 없는 요약 한 줄은 문자열로 다룬다
        n = n["t"]
    if isinstance(n, dict):
        t, d = n["t"], list(n["d"])
        m = RE_NT_CBETA.match(t)
        if m:
            rest = m.group(3).strip()
            d = [("원교감 원문: " + x.split(": ", 1)[1]) if x.startswith("대정장 원교감주: ") else x
                 for x in d]
            orig = [x for x in d if x.startswith("원교감 원문: ")]
            q = re.match(r"「(.*?)」\.?\s*(.*)$", rest, re.S)
            if q:
                orig = orig or ["원교감 원문: " + q.group(1)]
                rest = q.group(2).strip()
            body = [x for x in d if not x.startswith("원교감 원문: ")]
            gist = _nt_gist(d)
            d = orig + ["교감 번호: " + m.group(2)] + \
                (["요지: " + rest] if rest and gist else []) + body
            fall = orig[0].split(": ", 1)[1] if orig else "교감 " + m.group(2)
            return {"t": "[CBETA 교감] " + (gist or rest or fall), "d": d}
        m = RE_NT_ADD.match(t)
        if m:
            head = ["교감 번호: " + m.group(1).strip()]
            if m.group(2):
                head.append("비고: " + m.group(2).strip())
            if m.group(3).strip():
                head.insert(0, "요지: " + m.group(3).strip())
            d = head + d
            return {"t": "[CBETA 추가 교감] " + (_nt_gist(d) or m.group(3).strip()), "d": d}
        return n
    m = RE_NT_CBETA.match(n)
    if m:                                   # 상세 없이 한 줄로 쓴 원교감주
        rest = m.group(3).strip()
        d = ["교감 번호: " + m.group(2)]
        q = re.match(r"「(.*?)」\.?\s*(.*)$", rest, re.S)
        if q:
            d.insert(0, "원교감 원문: " + q.group(1))
            rest = q.group(2)
        ss = _nt_sentences(rest)
        return {"t": "[CBETA 교감] " + (ss[0] if ss else "교감 " + m.group(2)),
                "d": d + ss[1:]}
    for rx, tag in ((RE_NT_STAR, "[상호참조]"), (RE_NT_GAIJI, "[외자]")):
        m = rx.match(n)
        if m:
            body = m.group(m.lastindex)
            ss = _nt_sentences(body)
            d = ss[1:]
            if tag == "[외자]":
                d = ["CBETA 외자 번호: " + m.group(1)] + d
            return {"t": f"{tag} " + (ss[0] if ss else body), "d": d} if d else f"{tag} {body}"
    m = RE_NT_LONG.match(n)
    if m and len(n) > 90:                   # 상세 없이 긴 한 줄로 쓴 교감·주기 메모
        tag, body = m.group(1), m.group(2)
        d = []
        k = re.match(r"(?:교감(?:/주기)?\s*번호|note)\s*(\d{7}\w*)\s*[.:：]\s*", body)
        if k:
            d.append("교감 번호: " + k.group(1))
            body = body[k.end():]
        ss = [x.strip() for x in RE_NT_SENT.split(body) if x.strip()]
        if len(ss) > 1 or d:
            return {"t": f"{tag} {ss[0] if ss else body}", "d": d + ss[1:]}
    return n


# ── CBETA XML 판본 용어 정리 ─────────────────────────────────────────────
# DOCX 제작 때 CBETA 2018판 XML(xml-p5-2018, '구 XML')과 현행 XML(xml-p5,
# '현재 XML')을 대조한 기록이 교감 상세에 남는다. 독자에게는
#   · 두 판이 같다는 줄은 정보가 없으므로 뺀다
#   · '구 XML/현재 XML'은 'CBETA 2018판/현행 CBETA'로 바꾼다
#   · 두 판이 실제로 다를 때의 기록은 그대로 둔다
#   · 출처(파일 경로) 줄은 상세 맨 끝에 둔다
XML_TERMS = [
    (r"구\s*XML과\s*현재\s*XML", "CBETA 2018판과 현행판"),
    (r"현재\s*·\s*구\s*(?:CBETA\s*)?XML", "현행 CBETA와 2018판"),
    (r"구\s*·\s*현재\s*(?:CBETA\s*)?XML", "CBETA 2018판과 현행판"),
    (r"두 XML", "두 판"),
    (r"현재\s*CBETA\s*XML|현행\s*공식\s*XML|현재\s*XML|현행\s*XML", "현행 CBETA"),
    (r"구버전\s*XML|구판\s*XML|구\s*XML", "CBETA 2018판"),
    (r"구판·현행", "2018판·현행"),
    (r"현재판", "현행판"),
]
XML_TERMS = [(re.compile(a), b) for a, b in XML_TERMS]
RE_XML_SAME = re.compile(
    r"^(?:구버전 XML 대응:\s*있음|구\s*XML도 동일|XML 구조:\s*<figure>.*"
    r"|판본 정보:\s*(?:구·현재 XML 모두 동일한 대립을 보인다|구 XML과 현재 XML에 공통으로 존재한다))\.?$")


def _xml_src(v):
    v = re.sub(r"cbeta-org/", "", v)
    v = re.sub(r"현행\s*xml-p5\b", "현행판", v)
    v = re.sub(r"xml-p5-2018", "2018판", v)
    v = re.sub(r"xml-p5\b", "현행판", v)
    v = re.sub(r"\b[TX]/[TX]\d+/([TX]\d+n\d+\w*)\.xml", r"\1", v)
    return v


def xml_terms(t):
    for rx, rep in XML_TERMS:
        t = rx.sub(rep, t)
    return t


def shtk_xml_tidy(n):
    if isinstance(n, str):
        return xml_terms(n)
    d, src = [], []
    for x in n["d"]:
        if RE_XML_SAME.match(x.strip()):
            continue
        x = re.sub(r"\s*·\s*구판·현행 공통", "", x)
        x = re.sub(r"^비고:\s*구판·현행 공통$", "", x)
        if not x.strip():
            continue
        if re.match(r"^[^:]{0,12}출처:", x):
            src.append(_xml_src(xml_terms(x)))
        else:
            d.append(xml_terms(x))
    d += src
    t = xml_terms(n["t"])
    return {"t": t, "d": d} if d else t


RE_NT_ITEM = re.compile(r"^(\d{1,3})\.\s+(.*)$", re.S)


def shtk_split_notes(n):
    """번호 목록을 한 메모에 몰아넣은 블록을 낱낱의 메모로 편다.

    번역문의 「[주 12]」가 그 번호를 가리키므로, 묶여 있으면 찾아가기 어렵다.
    목록 머리말(「… 메모를 보존한다」)은 안내 문장이라 버린다.
    """
    if not isinstance(n, dict):
        return [n]
    items = [RE_NT_ITEM.match(x) for x in n["d"]]
    if len(n["d"]) < 3 or not all(items):
        return [n]
    return [f"[주 {m.group(1)}] {m.group(2).strip()}" for m in items]


# ── 메모 글을 읽을 수 있게 다듬기 ────────────────────────────────────────
# DOCX 제작 단계의 마크다운 백틱, XML 태그, 편집 책임자(resp) 표기가 그대로
# 남아 화면에 나오는 문서가 있다. 뜻을 바꾸지 않는 선에서 걷어 낸다.
RE_NT_XMLTAG = re.compile(r'<g ref="#(CB\d+)"[^>]*>|</?(charDecl|figure|graphic|app|note|lem|rdg)\b[^>]*>')
RE_NT_RESP = re.compile(r"\s*[^.]*#resp[^.]*\.\s*")
RE_NT_TYPE = re.compile(r'type="(\w+)"\s*')
RE_NT_MEMO = re.compile(r"^\[[^\]]*\]\s*(서지|구문|구두|판독|외자|의미|최소 보충|교감)\s*메모\s*[|｜]\s*")
RE_NT_GAIJI_LINE = re.compile(r"^외자[·\s]")


def _nt_clean(t):
    t = t.replace("`", "")
    t = RE_NT_XMLTAG.sub(lambda m: m.group(1) or m.group(2), t)
    t = RE_NT_RESP.sub(" ", t)
    t = RE_NT_TYPE.sub(r"\1 ", t)
    return re.sub(r"\s{2,}", " ", t).strip()


def shtk_readable_note(n):
    """백틱·XML 태그를 걷어 내고, 「… 메모 |」 꼴을 제 꼬리표로 바꾼다."""
    if isinstance(n, str):
        t, d = n, []
    else:
        t, d = n["t"], list(n["d"])
    t = _nt_clean(t)
    m = RE_NT_MEMO.match(t)
    if m:                                   # 「[CBETA 교감] 외자 메모 | …」 → 「[외자] …」
        t = f"[{m.group(1)}] " + t[m.end():].strip()
    elif RE_NT_GAIJI_LINE.match(t):
        t = "[외자] " + re.sub(r"^외자[·\s]문자 원형\s*:\s*", "", t)
    else:
        tag0, sep0, body0 = t.partition("] ")
        if sep0:
            # 「[CBETA 교감] *외자: …」·「… *별표 상호참조: …」는 교감이 아니다
            m2 = re.match(r"[*＊]?\s*(외자|별표\s*상호참조)\s*[:：]\s*", body0)
            if m2:
                name = "외자" if m2.group(1).startswith("외자") else "상호참조"
                t = f"[{name}] " + body0[m2.end():].strip()
    body = t.split("] ", 1)[-1]
    if body.startswith("최소 보충") is False and t.startswith("[최소 보충]") and len(body) < 20:
        t = "[최소 보충] 보충한 말: " + body
    d = [x for x in d if not re.match(r"^(?:상세 해설|상세)\s*[:：]\s*위 요약과 동일", x)]
    head = body.rstrip("… .")
    d = [x for x in (_nt_clean(x) for x in d)
         if x and x != body and x != t
         and not (head and len(head) > 20 and x.split(": ", 1)[-1].startswith(head))]
    return {"t": t, "d": d} if d else t


RE_HAN = re.compile(r"[\u3400-\u9fff\U00020000-\U0003ffff]")
RE_HANGUL = re.compile(r"[가-힣]")
RE_SHTK_SEP = re.compile(r"\s*[|｜·・]\s*")


def shtk_title_pair(txt, txt_keys, skey):
    """「漢文 제목 | 한국어 제목」이나 「漢文 · 한국어」를 나눈다.

    가운뎃점은 한국어 안에서도 쓰이므로, 왼쪽이 원문 TXT 의 제목과
    맞아떨어지는 자리에서만 나눈다.
    """
    for m in RE_SHTK_SEP.finditer(txt):
        left, right = txt[:m.start()].strip(), txt[m.end():].strip()
        if left and right and skey(left) in txt_keys:
            return left, right
    return None, None


def parse_shtk_docx(d, tables, path=None):
    units, front, appendix, secs, pending = [], [], [], [], []
    phase, cur, marker, cur_sec, parent = "front", None, None, None, None
    last_head, in_log, note_list = None, False, False

    # 원문 TXT 에 독립 단락으로 있는 한문 소제목(品題·章題)은 표제이면서 원문이다.
    # 그런 소제목은 원문 단위로 세워야 TXT 와 제자리가 맞고, 바로 뒤의 교감주도
    # 그 소제목에 붙는다. (편집자가 넣은 과단 표제는 TXT 에 없으므로 표제로만 둔다)
    def skey(t):
        return re.sub(r"[\s\u3000]", "", RE_MARKER.sub("", RE_APPARATUS.sub("", t)))
    # 원문·번역 칸에 박힌 인라인 문자 이미지는 ⟦fig:원본파일명⟧ 토큰으로 제자리에 둔다
    imgmap = (docx_image_names(d, ROOT / "assets" / "figures" / path.parent.name)
              if path is not None else {})
    txt_keys = set()
    if path is not None:
        tp = path.parent / "원문.txt"
        if tp.exists():
            # 문단 전체뿐 아니라 그 안의 각 줄도 제목이 될 수 있다
            # (편목처럼 여러 항목이 한 문단에 줄바꿈으로 이어진 저본이 있다)
            txt_keys = set()
            for u in parse_txt(tp):
                whole = "".join(u["cn"])
                txt_keys.add(skey(whole))
                for line in whole.split("\n"):
                    line = RE_MARKER.sub("", line).strip()
                    if 0 < len(line) <= 20:
                        txt_keys.add(skey(line))

    def shtk_detail(t):
        t = RE_SHTK_ITEM.sub("", t, count=1)
        k, sep, v = t.partition(" | ")
        k, v = k.strip(), v.strip()
        # '판본 정보 | 판본: …' 처럼 칸 이름이 값 앞에 한 번 더 붙은 꼴을 정리
        for pre in {k, k.split()[0]}:
            if v.startswith(pre + ":"):
                v = v[len(pre) + 1:].strip()
        return f"{k}: {v}" if sep else t

    def open_unit(cn, ko=None):
        nonlocal cur, pending
        if cur:
            units.append(cur)
        for k in pending:
            secs[k]["i"] = len(units)
        pending = []
        cur = {"m": marker, "cn": [cn], "ko": list(ko or []), "nt": [], "h": cur_sec}

    for p in d.paragraphs:
        style = (p.style.name or "").strip()
        if style != "SHTK Note Detail":
            in_log = False
        raw_p = (para_text_with_figs(p, imgmap)
                 if style in ("SHTK Source Text", "SHTK Translation Text") else p.text)
        txt = re.sub(r"[ \t]+", " ", raw_p).strip()
        if not txt:
            # 글 없이 도판만 담은 문단(저본에서 빠진 글자 그림 등)은
            # 그 자리에 도판 토큰으로 세운다
            fig = para_text_with_figs(p, imgmap).strip()
            if fig and cur is not None and phase == "body":
                cur["cn"].append(fig)
            continue
        is_h1 = style == "Heading 1"
        if phase == "front":
            if is_h1 and RE_SHTK_BODY.search(txt):
                phase = "body"
            elif style != "SHTK TOC Title":
                front.append(txt)
            continue
        if phase == "body" and is_h1:
            phase = "back"
        if phase == "back":
            appendix.append(txt)
            continue

        if style == "SHTK Subheading" or style.startswith("Heading"):
            # 「주」·「제1문 주」 같은 표제 아래의 목록만 역자 주석이다
            note_list = bool(re.fullmatch(r"(?:.{0,12}\s)?(?:주|주석|역주)", txt))
        elif style in ("SHTK Source Text", "SHTK Position Marker"):
            note_list = False
        if style in ("SHTK Subheading", "SHTK Meta"):   # 소제목·권말 제목
            cn_t, ko_t = shtk_title_pair(txt, txt_keys, skey)
            if cn_t is None and style == "SHTK Subheading":
                # 저본 본문의 첫 구절을 소제목으로 올려 쓴 문서가 있다.
                # 「漢文 · 한국어」 꼴이면 그 한문도 원문 단위로 세운다.
                m2 = re.search(r"\s[·・]\s", txt)
                if m2:
                    l, r = txt[:m2.start()].strip(), txt[m2.end():].strip()
                    if l and r and RE_HAN.search(l) and not RE_HANGUL.search(l):
                        open_unit(l, [r])
                        cur["sealed"] = True
                        last_head = None
                        continue
            pair = [cn_t] if cn_t else [txt.strip()]
            if pair[0] and skey(pair[0]) in txt_keys:
                ko = [ko_t] if ko_t else None
                # 「제목 3: 한국어 장제 → 소제목: 한문 장제」 짝이면 바로 앞 제목이
                # 이 장제의 번역이다. 비워 두면 앱에 '번역 대응 없음'이 뜬다.
                if not ko and last_head:
                    ko = [RE_APPARATUS.sub("", last_head).strip()]
                open_unit(pair[0], ko)
                cur["sealed"] = True
                last_head = None
                continue
            if style == "SHTK Verification":
                appendix.append(txt)            # 제작 단계의 검증 판정
                continue
            if style in ("SHTK Gaiji Meta", "SHTK Figure") and cur is not None:
                cur["nt"].append("[외자] " + re.sub(r"^외자[^|]*\|\s*", "", txt).strip())
                continue
            if style == "SHTK Meta" and cur is not None and cur.get("title") and not cur["ko"]:
                cur["ko"].append(txt)               # 품제의 한국어 제목
                continue
            if style == "SHTK Meta":
                # 본문 칸 사이의 안내·검증 요약 따위는 원문도 교감도 아니다
                appendix.append(re.sub(r"^[•·]\s*", "", txt))
                continue
        hm = re.match(r"Heading ([2-4])$", style)
        if hm and RE_SHTK_SEP.search(txt):
            # 「名號品第三 | 명호품 제3」처럼 원문 TXT의 품제(品題)를 제목으로 쓴 경우:
            # 목록 제목은 두 말을 함께 보이고, 품제는 원문 단위로도 세운다
            cn_t, ko_t = shtk_title_pair(txt, txt_keys, skey)
            pair = [cn_t, ko_t]
            if cn_t and skey(pair[0]) in txt_keys:
                hl = int(hm.group(1))
                if hl == 2:
                    parent = None
                elif hl == 3:
                    parent = pair[0]
                secs.append({"hl": hl, "raw": f"{pair[0]} ({pair[1]})", "parent": None})
                pending.append(len(secs) - 1)
                cur_sec = len(secs) - 1
                open_unit(pair[0], [pair[1]])
                cur["sealed"] = True
                last_head = None
                continue
        if hm or style == "SHTK Subheading":
            hl = int(hm.group(1)) if hm else 5
            if hl == 2:
                parent = None
            elif hl == 3:
                parent = txt
            secs.append({"hl": hl, "raw": txt, "parent": parent if hl == 4 else None})
            last_head = txt if hl in (2, 3) else None
            pending.append(len(secs) - 1)
            cur_sec = len(secs) - 1
            if cur is not None:
                cur["sealed"] = True
            # 제목 자체가 원문 TXT의 품제(名號品第三 등)이면 목차이면서 원문 단위이다.
            # 바로 뒤의 SHTK Meta(「명호품(名號品)」 제3)가 그 번역이 된다.
            if hm and txt_keys and skey(txt) in txt_keys:
                open_unit(txt)
                cur["sealed"] = cur["title"] = True
                last_head = None
            continue
        if style == "SHTK Block Label":
            continue
        last_head = None
        if style == "SHTK Position Marker":
            pm = PURE_MARKER.match(txt)
            if pm:
                marker = pm.group(1)
            continue
        if style == "SHTK Source Text":
            mk = RE_MARKER.match(txt)
            if mk:
                marker = mk.group(1)
            # 원문 한 칸에 위치표지가 여럿 들어 있으면(게송·음석이 산문 뒤에 이어질 때)
            # 표지마다 조각으로 나눈다. 조각이 제 표지 자리를 찾아가야 원문 순서가 맞는다.
            parts = [x.strip() for x in re.split(r"(?=\[\d{3,4}[abc]\d{2}\])", txt) if x.strip()]
            if cur and not cur["ko"] and not cur["nt"] and not cur.get("sealed"):
                cur["cn"] += parts
            else:
                open_unit(parts[0])
                cur["cn"] += parts[1:]
            continue
        if cur is None:
            front.append(txt)
            continue
        if style in ("Quote", "Intense Quote"):      # 번역문 안에 인용한 게송 줄
            if cur is not None:
                cur["ko"].append(txt)
                continue
        if re.match(r"List (?:Number|Bullet|Paragraph)", style):
            # 번호 목록은 대개 번역 본문의 열거(장문 아홉 따위)다.
            # 「주」 표제 아래의 목록만 역자 주석으로 돌린다.
            if cur is not None:
                if note_list:
                    cur["nt"].append("[역주] " + txt)
                else:
                    cur["ko"].append(txt)
                continue
        if style == "SHTK Translation Text":
            # 번역 끝에 「[이미지 확인] ⟦fig⟧⟦fig⟧…」로 원문 이미지를 한데 모아 둔 문서가 있다.
            # 이미지 글자는 원문 칸 제자리에 이미 있고 교감 메모에도 하나씩 달려 있으므로,
            # 번역 칸에서는 이 꼬리 묶음을 뺀다. (번역 문장 안의 토큰은 그대로 둔다)
            txt = RE_KO_FIG_TAIL.sub("", txt).rstrip()
            # 번역 끝에 원문 교감표지를 그대로 옮겨 붙인 「… 밝힌다. [1]」은 뺀다
            # (표지는 원문 칸에 있고, 교감은 아래 메모에 달린다)
            txt = RE_KO_APP_TAIL.sub("", txt).rstrip()
            # 번역문에 각주 번호만 덩그러니 박아 둔 문서가 있다(「…뜻이다.60」).
            # 숫자로만 두면 오자로 보이므로 「[주 60]」으로 드러낸다.
            # 번호는 그 권의 교감·번역 메모에 달린 번호 항목을 가리킨다.
            txt = RE_KO_FOOTNUM.sub(r"\1[주 \2]", txt)
            cur["ko"].append(txt)
        elif style == "SHTK Note Summary" and RE_SHTK_LOG.search(txt):
            # 교감이 아니라 DOCX 제작 단계의 처리 통계(몇 건 대조·몇 종 처리 따위)다.
            # 본문 교감 목록에 섞지 않고 부록으로 보낸다. 뒤따르는 상세 줄도 함께.
            appendix.append(RE_SHTK_SUMMARY.sub("", txt, count=1).strip())
            in_log = True
            continue
        elif style == "SHTK Note Detail" and in_log:
            appendix.append(shtk_detail(txt))
            continue
        elif style == "SHTK Note Summary":
            t = RE_SHTK_SUMMARY.sub("", txt, count=1).strip()
            if not t.startswith("["):
                t = "[CBETA 교감] " + t
            cur["nt"].append({"t": t, "d": []})
        elif style == "SHTK Note Detail" and cur["nt"] and isinstance(cur["nt"][-1], dict):
            # 상세 한 문단에 줄바꿈으로 「• 항목 | 내용」 여러 줄을 담은 문서가 있다
            # 줄바꿈으로 「• 항목 | 내용」을 여러 줄 담은 문단만 나눈다.
            # 머리표 없는 줄은 앞 줄이 이어진 것이므로 도로 붙인다.
            chunks = []
            for x in txt.split("\n"):
                x = x.strip()
                if not x:
                    continue
                if chunks and not RE_SHTK_ITEM.match(x):
                    chunks[-1] += " " + x
                else:
                    chunks.append(x)
            cur["nt"][-1]["d"] += [shtk_detail(x) for x in chunks]
        else:
            cur["nt"].append(shtk_detail(txt))
    if cur:
        units.append(cur)

    # 권(목록) 층위 정하기
    #  · 제2표제 권마다 150단위 이하이면: 제2표제=권, 제3·4표제·소제목은 권 안 제목
    #  · 한 권이 너무 크면(위치표지가 적어 쪽으로도 못 나눌 때): 제2~4표제를
    #    모두 권으로 펴고, 제4표제에는 윗 제3표제를 앞에 붙여 헷갈리지 않게 한다
    placed = [s for s in secs if "i" in s]
    top = [s["i"] for s in placed if s["hl"] == 2]
    sizes = [b - a for a, b in zip(top, top[1:] + [len(units)])]
    flat = ((len(top) < 2 or max(sizes) > 150)
            and sum(1 for s in placed if s["hl"] <= 4) <= 120)   # 펴도 목록이 감당될 때만
    for s in secs:
        if flat:
            s["lv"] = 1 if s["hl"] <= 4 else 2
            s["t"] = f"{s['parent']} · {s['raw']}" if s["hl"] == 4 and s["parent"] else s["raw"]
        else:
            s["lv"] = s["hl"] - 1
            s["t"] = s["raw"]
    for u in units:
        out = []
        for n in u["nt"]:
            v = shtk_variant_note(n)
            if v is None:
                continue
            for x in shtk_split_notes(shtk_readable_note(shtk_xml_tidy(shtk_tidy_note(v)))):
                # 요지 끝에 상세 첫 줄을 되풀이한 꼬리(「… · [0175a13] 題 / 顯 …」)는 자른다
                if isinstance(x, dict) and " · " in x["t"] and x["d"] \
                        and x["t"].split(" · ", 1)[1][:12] in x["d"][0]:
                    x = {"t": x["t"].split(" · ", 1)[0].rstrip(), "d": x["d"]}
                out.append(x)
        u["nt"] = out
    for u in units:
        u.pop("sealed", None)
        u.pop("title", None)
        u["h"] = secs[u["h"]]["t"] if u["h"] is not None else None
    sections = [{"lv": s["lv"], "t": s["t"], "i": s["i"]} for s in placed]
    return {"units": units, "tables": tables, "front": front,
            "appendix": appendix, "sections": sections}


RE_CBETA_NOTE_ID = re.compile(r"\[(?:CBETA\s*교감|대정장\s*원교감주)\s*\d{7}")


def parse_docx(path: Path):
    if docx is None:
        raise RuntimeError("python-docx 가 필요합니다: pip install python-docx")
    d = open_docx(path)

    # 1) 표: 서지 정보 / 용어 대응표 회수
    tables = []
    for t in d.tables:
        rows = [[c.text.strip() for c in r.cells] for r in t.rows]
        rows = [r for r in rows if any(r)]
        if rows:
            tables.append(rows)

    # 표준 DOCX(SHTK 스타일)는 스타일이 역할을 정해 주므로 따로 읽는다
    if is_shtk_docx(d):
        return parse_shtk_docx(d, tables, path)

    # 2) 본문 단락 스캔
    blocks, cur_head, cur_marker = [], None, None
    imgmap = docx_image_names(d, ROOT / "assets" / "figures" / path.parent.name)
    # 본문 앞머리에 붙은 '제작 메모'(누가 몇 단위를 교열했다는 따위)는
    # 저본의 글이 아니므로 본문에 세우지 않고 해제로 돌린다.
    body_style = any((q.style.name or "").strip() == "Source Text"
                     for q in d.paragraphs)
    body_open = not body_style
    # 교감표지가 박힌 한문 소제목(Small Meta) 바로 뒤에 그 교감주가 오는 문서가 있다.
    #   第十地受[1]識章  →  교감·번역 메모 | [CBETA 교감 0575001] …
    # 소제목을 표제로만 흘려보내면 교감주가 앞 단위 끝에 붙어 엉뚱한 자리에 뜬다.
    # 원문 TXT 도 이 소제목을 한 단락으로 두므로, 원문 조각으로 세워 제자리에 맞춘다.
    paras = list(d.paragraphs)
    nxt_para = {}
    for k, q in enumerate(paras):
        for r in paras[k + 1:]:
            if r.text.strip():
                nxt_para[k] = r
                break
    for pi, p in enumerate(paras):
        style = (p.style.name or "").strip()
        nq = nxt_para.get(pi)
        # 뒤따르는 메모가 번호 붙은 CBETA/대정장 교감주일 때만 그렇게 한다
        # (다른 문헌의 소제목 처리는 그대로 둔다)
        if (style == "Small Meta" and RE_APPARATUS.search(p.text)
                and is_source_line(p.text.strip()) and nq is not None
                and "note" in (nq.style.name or "").lower()
                and RE_CBETA_NOTE_ID.search(nq.text)):
            style = "Source Text"
        # 문서 표제(Document Title/Subtitle)는 책 이름일 뿐 원문 단락이 아니다.
        # 원문으로 세우면 권두 권제(卷第一)와 대조가 엇갈린다.
        if style in ("Document Title", "Document Subtitle"):
            if p.text.strip():
                blocks.append({"kind": "front", "text": p.text.strip()})
            continue
        # 본문 칸의 스타일 이름은 문서마다 조금씩 다르다
        # (Source Text · Original Text · OriginalText · Translation Text …)
        raw_p = (para_text_with_figs(p, imgmap)
                 if style.replace(" ", "").endswith("Text") else p.text)
        txt = re.sub(r"[ \t]+", " ", raw_p).strip()
        if not txt:
            continue
        if not body_open:
            if style == "Source Text":
                body_open = True
            elif RE_MAKER_NOTE.search(txt):
                blocks.append({"kind": "front", "text": txt})
                continue
        # 도판을 실제로 띄우므로 '[X08p0116_01.gif]' 같은 파일명 표기는 지운다
        txt = RE_FIG_NAME.sub("", txt).strip()
        if not txt:
            continue

        # '주' / '주석' / '각주' / '제1문 주' 는 스타일과 무관하게 각주 시작으로 본다
        if (len(txt) <= 8 and RE_NOTE_HEAD.search(txt)) or (
                RE_NOTE_LABEL_HEAD.match(txt)
                and not RE_NOTE_LABEL_NOT.search(txt)):
            blocks.append({"kind": "head", "hkind": "note", "text": txt,
                           "lv": head_level(style), "m": cur_marker})
            continue
        if txt in LABELS or label_key(txt) in LABELS:
            continue
        # 작업 표지 줄은 본문이 아니다.
        #   — 원문 전체 위치표지 대응 재배열 완료 —
        # 처럼 줄 전체가 줄표로 감싸인 짧은 문단은 번역본을 만든 쪽의
        # 작업 기록이므로 앞 단위의 번역 끝에 붙이지 않고 버린다.
        if RE_WORKLOG.match(txt):
            continue
        # 표지만 홀로 선 문단([0297a11])은 위치 표시일 뿐 본문이 아니다.
        # 걸러 내지 않으면 앞 단위의 번역 끝에 군더더기로 달라붙는다.
        pm = PURE_MARKER.match(txt)
        if pm:
            cur_marker = pm.group(1)
            continue
        if is_unit_label(txt):
            continue

        txt, side = strip_side(txt)
        if not txt:
            continue

        mk = RE_MARKER.search(txt)

        # '第二卷  |  제2권.' 처럼 한 문단에 한문 표제와 그 번역을 나란히
        # 담은 문서가 있다. 표제로 흘려보내면 번역이 사라지므로
        # 원문·번역 두 조각으로 갈라 준다.
        pair = re.split(r"\s*[|｜]\s*", txt)
        if (len(pair) == 2 and style in ("Small Meta", "Unit Label", "End Matter")
                and pair[0] and pair[1]
                and is_source_line(pair[0]) and not is_source_line(pair[1])):
            if mk:
                cur_marker = mk.group(1)
            blocks.append({"kind": "cn", "text": pair[0],
                           "m": cur_marker, "head": cur_head})
            blocks.append({"kind": "ko", "text": pair[1],
                           "m": cur_marker, "head": cur_head})
            continue

        is_head = "heading" in style.lower() or style in (
            "Unit Heading", "Unit Label", "Title", "Subtitle", "Small Meta"
        ) or re.match(r"^(문단\s*\d+|제\s*\d+\s*항|권\s*제?\d+)", txt)

        if is_head:
            kind = classify_head(txt, style)
            if kind == "label":
                continue
            # 한문 표제 바로 뒤에 그 한국어 번역이 오는 문서가 있다.
            # 따로 두면 번역이 사라지므로 앞 표제에 이어 붙인다.
            if (kind == "structure" and blocks
                    and blocks[-1].get("kind") == "head"
                    and blocks[-1].get("hkind") == "structure"
                    and is_source_line(blocks[-1]["text"])
                    and not is_source_line(txt)
                    and " — " not in blocks[-1]["text"]):
                blocks[-1]["text"] += " — " + txt
                cur_head = blocks[-1]["text"]
                continue
            if kind != "apparatus":
                cur_head = txt
                if mk:
                    cur_marker = mk.group(1)
            blocks.append({"kind": "head", "hkind": kind, "text": txt,
                           "lv": head_level(style), "m": cur_marker})
            continue

        if side == "cn" or (side is None and is_source_line(txt)):
            # 한 문단 안에 위치표지가 여럿이면 표지마다 쪼갠다.
            # (docx 가 원문 두 대목을 줄바꿈으로 한 문단에 담는 경우가 있다)
            pieces = [x.strip() for x in
                      re.split(r"(?=\[\d{3,4}[abc]\d{2}\])", txt) if x.strip()]
            for piece in pieces:
                pm = RE_MARKER.match(piece)
                if pm:
                    cur_marker = pm.group(1)
                blocks.append({"kind": "cn", "text": piece,
                               "m": cur_marker, "head": cur_head})
            continue

        kind = "note" if (
            side != "ko" and (
                "note" in style.lower() or "각주" in style or "Editorial" in style
                or NOTE_PREFIX.match(txt))
        ) else "ko"
        blocks.append({"kind": kind, "text": txt, "m": cur_marker, "head": cur_head})

    # 3) 단위 조립
    #    suppress=True 인 동안의 한국어 문단은 '해제'로 보내고 본문에 붙이지 않는다.
    #    한문 단락이 나오면 suppress 는 자동으로 풀린다.
    units, cur = [], None
    front, appendix = [], []
    body_started = False
    suppress = True          # 첫 한문 단락 전까지는 전부 해제
    note_mode = False        # '주' 표제 아래 — 이하 문단은 각주로 모은다
    sections = []            # [{lv, t, i}]  i = 시작 단위 인덱스
    pending_sec = []

    def stash(text):
        (appendix if body_started else front).append(text)

    for b in blocks:
        if b["kind"] == "front":
            front.append(b["text"])
            continue
        if b["kind"] == "head":
            if b["hkind"] == "note":
                note_mode = True
                continue
            if note_mode and b["lv"] >= 3:
                # 주석 묶음 안의 소표제 — 절이 아니라 각주 줄로 둔다
                if cur is not None:
                    cur["nt"].append(b["text"])
                continue
            note_mode = False
            if b["hkind"] == "apparatus":
                if cur:
                    cur["sealed"] = True
                suppress = True
                continue
            if b["hkind"] == "structure":
                if not pending_sec or pending_sec[-1]["t"] != b["text"]:
                    pending_sec.append({"lv": b["lv"], "t": b["text"]})
            if cur is not None:
                cur["sealed"] = True
            continue

        if b["kind"] == "cn":
            suppress = False
            note_mode = False
            if not body_started and b["m"]:
                body_started = True
            if cur and not cur["ko"] and not cur["nt"] and not cur.get("sealed"):
                cur["cn"].append(b["text"])
                if cur["m"] is None:
                    cur["m"] = b["m"]
                continue
            if cur:
                units.append(cur)
            for sec in pending_sec:
                sections.append({**sec, "i": len(units)})
            pending_sec = []
            cur = {"m": b["m"], "cn": [b["text"]], "ko": [], "nt": [],
                   "h": b.get("head")}
            continue

        # 한국어 / 주석
        if suppress or cur is None:
            stash(b["text"])
            continue
        if b["kind"] == "ko" and not note_mode:
            cur["ko"].append(b["text"])
        else:
            cur["nt"].append(b["text"])

    if cur:
        units.append(cur)

    # 원문 자리에 한문이 아니라 한국어 작업 설명이 들어앉은 단위는 본문이 아니다.
    # (번역 docx 말미의 「정리 기준」 같은 대목이 '원문' 라벨 아래 놓이는 일이 있다)
    kept = []
    for u in units:
        if any(RE_CJK.search(x) and hangul_ratio(x) < 0.3 for x in u["cn"]):
            kept.append(u)
        else:
            appendix.extend(u["cn"] + u["ko"] + u["nt"])
    units = kept

    # 원문 여러 문단을 몰아 놓고 번역 여러 문단을 몰아 놓은 docx 는
    # 한 덩이가 통째로 한 단위가 되어 대조가 되지 않는다. 짝을 지어 나눈다.
    units, remap0 = split_paired_blocks(units)
    sections = [{**sec, "i": remap0.get(sec["i"], sec["i"])} for sec in sections]

    # 번역 표지대로 단위를 쪼개면 번호가 밀리므로, 절 표제의 위치도 함께 옮긴다
    units, remap = split_by_ko_markers(units)
    sections = [{**sec, "i": remap.get(sec["i"], sec["i"])} for sec in sections]

    return {"units": units, "tables": tables,
            "front": front, "appendix": appendix, "sections": sections}


RE_NOTE_QUOTE = re.compile(r"[「『]([^」』]{4,60})[」』]")


def split_paired_blocks(units):
    """원문 문단 여럿과 번역 문단 여럿이 한 단위에 몰려 있으면 짝지어 나눈다.

    길장 주석 계열 docx 처럼 원문 수십 문단을 먼저 늘어놓고 그 뒤에 번역
    수십 문단을 같은 순서로 늘어놓는 문서가 있다. 그대로 두면 한 단위가
    만 자를 넘어 대조가 되지 않는다. 원문 조각 수와 번역 조각 수가 꼭
    같을 때만, 순서대로 하나씩 짝지어 별개 단위로 만든다.
    각주는 인용한 구절이 들어 있는 조각에 붙인다.
    쪼개기 전후의 번호 대응표를 함께 돌려준다."""
    out, remap = [], {}
    for old, u in enumerate(units):
        remap[old] = len(out)
        cn, ko = u["cn"], u["ko"]
        if len(cn) < 2 or len(cn) != len(ko):
            out.append(u)
            continue

        pieces = []
        for i, (c, k) in enumerate(zip(cn, ko)):
            mk = RE_MARKER.match(c)
            pieces.append({"m": mk.group(1) if mk else (u["m"] if i == 0 else None),
                           "cn": [c], "ko": [k], "nt": [], "h": u.get("h")})
        keys = [norm_search(c) for c in cn]
        for t in u["nt"]:
            hit = 0
            q = RE_NOTE_QUOTE.search(t)
            if q:
                needle = norm_search(q.group(1))
                for i, key in enumerate(keys):
                    if needle and needle in key:
                        hit = i
                        break
            pieces[hit]["nt"].append(t)
        out.extend(pieces)
    return out, remap


def split_by_ko_markers(units):
    """번역 문단마다 위치표지가 붙어 있으면 그 표지대로 원문과 번역을 함께 나눈다.

    원문 수십 줄을 한 '대조 단위'로 묶어 번역한 문서라도, 번역 문단이
    [0091a06] 같은 표지를 달고 있으면 그 표지에서 다음 표지 직전까지의
    원문 줄을 그 번역 옆에 붙일 수 있다. 이렇게 해야 대조가 줄 단위로 맞는다.
    표지가 없는 번역은 앞 조각에 이어 둔다.
    쪼개기 전후의 번호 대응표(remap)를 함께 돌려준다."""
    out, remap = [], {}
    for old, u in enumerate(units):
        remap[old] = len(out)

        marked = [(RE_MARKER.match(k), k) for k in u["ko"]]
        anchors = [m for m, _ in marked if m]
        if len(anchors) < 2 or len(anchors) < len(u["ko"]) * 0.8:
            out.append(u)
            continue

        # 표지별로 번역을 묶는다
        groups = []
        for m, k in marked:
            if m:
                groups.append([m.group(1), [k]])
            elif groups:
                groups[-1][1].append(k)

        # 각 표지가 원문 어느 줄에서 시작하는지 찾는다
        cn_marks = [RE_MARKER.match(c) for c in u["cn"]]
        n, cursor = len(u["cn"]), 0
        starts = []
        for mk, _ in groups:
            pos = None
            for j in range(cursor, n):
                mm = cn_marks[j]
                if mm and mm.group(1) == mk:
                    pos = j
                    break
            starts.append(cursor if pos is None else pos)
            cursor = starts[-1]

        for gi, (mk, kos) in enumerate(groups):
            s0 = starts[gi]
            s1 = starts[gi + 1] if gi + 1 < len(groups) else n
            if s1 < s0:
                s1 = s0
            cn = u["cn"][s0:s1]
            if gi == 0 and s0 > 0:
                cn = u["cn"][:s0] + cn          # 표제·찬자 등 앞머리는 첫 조각에
            rec = {"m": mk, "cn": cn, "ko": kos, "nt": []}
            if gi == 0 and u.get("h"):
                rec["h"] = u["h"]
            out.append(rec)
        out[-1]["nt"] = u["nt"]                 # 각주는 마지막 조각에
    remap[len(units)] = len(out)
    return out, remap


# ── 정렬(병합) ───────────────────────────────────────────────────────
# ── 각주 줄 나누기 ───────────────────────────────────────────────
# 번역 docx 가운데는 그 단위의 교감·판독 메모를 한 문단에 몰아 쓰고
# 「 / 」로만 항목을 가르는 형식이 있다. 화면에서는 한 덩어리로 뭉쳐
# 보이므로, 항목마다 줄을 나눠 둔다.
RE_NOTE_LABEL = re.compile(
    r"^\s*(?:[^\[\]\n]{0,24}?(?:메모|주기|비고))\s*[|｜:：]\s*")
QUOTE_PAIRS = {"\u2018": "\u2019", "\u201c": "\u201d",
               "\u300c": "\u300d", "\u300e": "\u300f",
               "(": ")", "\uff08": "\uff09",
               "[": "]", "\uff3b": "\uff3d"}
QUOTE_CLOSE = set(QUOTE_PAIRS.values())


def split_note_line(t: str):
    """각주 한 문단을 「 / 」 자리에서 항목마다 끊는다.

    따옴표·괄호 안의 빗금은 건드리지 않는다. 번역문을 인용하며
    ‘첫째 … / 둘째 …’처럼 쓴 자리가 있어서, 이것까지 끊으면
    한 각주가 엉뚱하게 두 줄로 갈라진다."""
    out, buf, stack, i = [], [], [], 0
    while i < len(t):
        ch = t[i]
        if ch in QUOTE_PAIRS:
            stack.append(QUOTE_PAIRS[ch])
        elif stack and ch == stack[-1]:
            stack.pop()
        elif ch in QUOTE_CLOSE and ch in stack:
            while stack and stack.pop() != ch:
                pass
        if (not stack and ch == "/" and i > 0
                and t[i - 1] in " \t" and i + 1 < len(t) and t[i + 1] in " \t"):
            out.append("".join(buf).strip())
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    out.append("".join(buf).strip())
    return [x for x in out if x]


def split_note_items(units):
    """각주 문단에서 칸 이름(‘교감·번역 메모 |’)을 떼고 항목마다 줄을 나눈다."""
    for u in units:
        if not u.get("nt"):
            continue
        out = []
        for t in u["nt"]:
            if isinstance(t, str) and t.startswith("[주 "):
                out.append(t)                 # 번호 붙은 역자 주는 그대로 한 줄
                continue
            if isinstance(t, dict):           # 요지·상세로 이미 나뉜 교감
                out.append(t)
                continue
            body = RE_NOTE_LABEL.sub("", t, count=1).strip()
            if not body:                      # 칸 이름뿐인 줄은 버린다
                continue
            out.extend(split_note_line(body))
        u["nt"] = out
    return units


FIG_MARK = "\u27e6fig:"


def merge(txt_units, dx):
    """원문 TXT 를 정본으로 두고, 번역 docx 단위를 위치표지·본문 대조로 붙인다.

    대응 우선순위
      1) 위치표지 일치
      2) 한문 앞머리 지문 일치
      3) 원문 전체를 한 줄로 이은 뒤 부분문자열 탐색 (docx 가 원문을 잘게 쪼갠 경우)
    여러 docx 단위가 한 TXT 단위에 대응하면 번역을 순서대로 이어 붙인다.
    """
    dx_units = dx["units"] if dx else []

    # docx 가 원문을 사실상 전부 담고 있으면 docx 의 세분 단위를 정본 순서로 삼는다
    if txt_units and dx_units:
        t_all = norm_search(" ".join(u["cn"] for u in txt_units))
        d_all = norm_search(" ".join(" ".join(u["cn"]) for u in dx_units))
        if len(t_all) and len(d_all) / len(t_all) >= 0.93 and len(dx_units) > len(txt_units):
            txt_units = []

    # docx 단독 문헌
    if not txt_units:
        out = []
        for n, u in enumerate(dx_units, 1):
            rec = {"i": n, "m": u["m"], "cn": u["cn"], "ko": u["ko"], "nt": u["nt"]}
            if u.get("h"):
                rec["h"] = u["h"]
            out.append(rec)
        return out, {k: k for k in range(len(dx_units))}

    out = []
    for tu in txt_units:
        rec = {"i": tu["i"], "m": tu["m"], "cn": [tu["cn"]], "ko": [], "nt": []}
        out.append(rec)

    # 한 위치표지에 원문 단락이 여럿 딸리는 일이 흔하다(問·答이 같은 행에서 시작하는 등).
    # 그래서 표지·앞머리 모두 '대기열'로 두고, 한 번 짝지어진 단위는 다시 쓰지 않는다.
    by_marker, by_head = {}, {}
    flat, offsets = [], []
    pos = 0
    for idx, rec in enumerate(out):
        if rec["m"]:
            by_marker.setdefault(rec["m"], []).append(idx)
        by_head.setdefault(head_key(rec["cn"][0]), []).append(idx)
        n = norm_search(rec["cn"][0])
        flat.append(n)
        offsets.append((pos, idx))
        pos += len(n)
    haystack = "".join(flat)
    starts = [o[0] for o in offsets]

    import bisect

    def locate(needle):
        if len(needle) < 6:
            return None
        p = haystack.find(needle)
        if p < 0:
            return None
        return offsets[bisect.bisect_right(starts, p) - 1][1]

    # docx 한 덩어리가 원문 여러 단락에 걸치는 경우가 흔하다.
    # 조각마다 대응 위치를 찾아, 걸친 범위를 한 단위로 합친다.
    taken = set()

    def pick(queue):
        for i in queue:
            if i not in taken:
                return i
        return None

    def resolve(piece):
        # 본문 앞머리가 가장 확실한 단서이므로 먼저 본다
        k = head_key(piece)
        if k and k in by_head:
            i = pick(by_head[k])
            if i is not None:
                return i
        mk = RE_MARKER.search(piece)
        if mk and mk.group(1) in by_marker:
            i = pick(by_marker[mk.group(1)])
            if i is not None:
                return i
        return locate(norm_search(piece)[:24])

    unmatched, dxmap, spans = [], {}, []
    owners = set()
    for dxi, u in enumerate(dx_units):
        if u["cn"]:
            # docx 의 원문 한 조각이 줄바꿈으로 여러 행을 담고 있을 때가 있다.
            # (산문 뒤에 게송이 이어지는 대목이 그렇다.)
            # 원문 TXT 는 그 게송을 따로 떼어 두므로, 행마다 대응을 찾아야
            # 게송 원문이 '번역 대응 없음'으로 떨어져 나가지 않는다.
            # 자리를 정하는 것은 원문 조각의 앞머리다.
            targets = [t for t in (resolve(c) for c in u["cn"]) if t is not None]
            # 시작 자리는 첫 원문 조각이 정한다. 뒤 조각이 우연히 앞쪽 글귀와
            # 맞아떨어져 덩어리 전체가 앞으로 끌려가는 것을 막는다.
            first = resolve(u["cn"][0])
            pieces = list(u["cn"])
            # 아래 둘은 '어디까지 걸치는가'만 넓힌다. 시작 자리는 바꾸지 않는다.
            #  · docx 원문 조각이 줄바꿈으로 여러 행을 담고 있을 때
            #    (산문 뒤에 게송이 이어지는 대목)
            #  · '頌曰' 뒤에 게송을 한 줄로 붙여 왔을 때
            tail = []
            for c in u["cn"]:
                if "\n" in c:
                    for x in c.split("\n"):
                        if len(x.strip()) >= 6:
                            pieces.append(x)
                            t = resolve(x)
                            if t is not None:
                                tail.append(t)
                if RE_GATHA_CUE.search(c):
                    t = locate(norm_search(c)[-24:])
                    if t is not None:
                        tail.append(t)
            if targets:
                # 앞머리가 정한 자리보다 앞으로 끌려가지 않게 막는다
                floor = min(targets)
                targets += [t for t in tail if t >= floor]
            else:
                targets = tail
        else:
            first = None
            # 원문 조각 없이 표지만 있는 번역 닻.
            # 원문 한 줄이 길면 그 줄 하나에 번역 문단이 여럿 달린다.
            # 그러므로 이미 다른 번역이 붙은 자리라도 이어 붙일 수 있어야 한다.
            targets = []
            if u["m"] and u["m"] in by_marker:
                q = by_marker[u["m"]]
                i = pick(q)                  # 아직 비어 있는 자리가 있으면 그쪽
                if i is None:
                    i = q[0]                 # 없으면 같은 표지의 첫 자리에 이어 붙인다
                targets = [i]
        if not targets:
            unmatched.append((dxi, u))
            continue
        # docx 한 단위가 담은 원문 조각 수보다 훨씬 넓은 범위에 걸쳐 있다면
        # 어느 한 조각이 엉뚱한 자리에 붙은 것이다. 그 이상치는 버린다.
        targets.sort()
        owner = first if u["cn"] and first is not None and first in targets else targets[0]
        targets = [t for t in targets if t >= owner] or [owner]
        span_max = len(pieces) * 2 + 10 if u["cn"] else len(u["cn"]) * 2 + 10
        targets = [t for t in targets if t - owner <= span_max]
        last = targets[-1]
        taken.update(targets)
        dxmap[dxi] = owner
        owners.add(owner)
        spans.append((owner, last))
        out[owner]["ko"].extend(u["ko"])
        out[owner]["nt"].extend(u["nt"])
        if u.get("h") and "h" not in out[owner]:
            out[owner]["h"] = u["h"]

    # 다른 번역이 걸려 있지 않은 중간 단락만 앞 단위로 흡수한다
    absorbed = {}
    for owner, last in spans:
        for k in range(owner + 1, last + 1):
            if k not in owners and k not in absorbed:
                absorbed[k] = owner

    # 묶음형 docx 대응
    #   원문 여러 줄(수십 줄)을 한 '대조 단위'로 묶어 번역한 문서에서는
    #   단위 사이에 낀 원문 줄이 어디에도 걸리지 않고 남는다.
    #   이때는 docx 단위를 구간 경계로 삼아, 다음 경계 직전까지를 그 단위에 붙인다.
    if dx_units:
        per_unit = sorted(len(u["cn"]) for u in dx_units if u["cn"])
        if not per_unit:
            per_unit = [0]
        median_cn = per_unit[len(per_unit) // 2]
        if median_cn >= 5:                       # 한 단위가 원문 5줄 이상을 묶는 문서
            bounds = sorted(owners)
            for n, start in enumerate(bounds):
                stop = bounds[n + 1] if n + 1 < len(bounds) else len(out)
                for k in range(start + 1, stop):
                    if k not in owners and k not in absorbed:
                        absorbed[k] = start
    for k in sorted(absorbed):
        out[absorbed[k]]["cn"].extend(out[k]["cn"])

    # 원문 TXT 에 없는 번역 단위(다른 저본·회본 등)는 별권으로 이어 붙인다
    base = len(out)
    for n, (dxi, u) in enumerate(unmatched, 1):
        dxmap[dxi] = len(out)
        out.append({"i": base + n, "m": u["m"], "cn": u["cn"],
                    "ko": u["ko"], "nt": u["nt"], "x": 1})

    final, remap = [], {}
    for idx, rec in enumerate(out):
        if idx in absorbed:
            continue
        remap[idx] = len(final)
        rec["i"] = len(final) + 1
        final.append(rec)
    dxmap = {k: remap.get(v, remap.get(absorbed.get(v, v), 0))
             for k, v in dxmap.items()}

    return final, dxmap


# ── 표에서 용어표 추출 ───────────────────────────────────────────────
def pick_glossary(tables):
    gloss = []
    for rows in tables:
        header = [h.strip() for h in rows[0]]
        joined = " ".join(header)
        if ("원어" in joined or "번역어" in joined or "원문" in joined) and len(header) >= 2:
            hi = 0
            for i, h in enumerate(header):
                if "원어" in h or h == "원문":
                    hi = i
            ki = 1 if hi == 0 else 0
            for i, h in enumerate(header):
                if "번역" in h and i != hi:
                    ki = i
            for r in rows[1:]:
                if len(r) > max(hi, ki) and r[hi] and r[ki]:
                    gloss.append({"cn": r[hi], "ko": r[ki],
                                  "note": r[max(hi, ki) + 1] if len(r) > max(hi, ki) + 1 else ""})
            if gloss:
                return gloss
    return gloss


def pick_biblio(tables):
    bib = {}
    for rows in tables:
        if len(rows[0]) == 2 and rows[0][0] in ("항목", "구분") or (
            len(rows[0]) == 2 and any("문헌명" in r[0] for r in rows)
        ):
            for r in rows:
                if len(r) >= 2 and r[0] and r[1] and r[0] not in ("항목", "구분"):
                    bib.setdefault(r[0], r[1])
    return bib


# ── 실행 ─────────────────────────────────────────────────────────────
def detect_juan_from_source(units):
    """원문 안에 남아 있는 권 표제(…卷第一 / …卷上)로 분권을 잡는다."""
    secs, seen = [], set()
    for idx, u in enumerate(units):
        for line in u["cn"]:
            line = line.strip()
            # 교감 표지와 서명 부분은 빼고 '卷第三' 꼴만 견준다.
            # 같은 권의 여는 줄과 닫는 줄, 서명에 오자가 섞인 줄을
            # 서로 다른 권으로 세지 않기 위함이다.
            key = RE_APPARATUS.sub("", line)
            key = key[key.find("卷"):] if "卷" in key else key
            if RE_JUAN_LINE.match(line) and key not in seen:
                seen.add(key)
                secs.append({"lv": 1, "t": line, "i": idx})
                break
    return secs


def chunk_by_chars(units, start, end, target=26000):
    """단위 수가 적어도 글자 수가 많으면 쪽 경계에서 끊는다.
    묶음형 번역처럼 한 단위가 원문 수십 줄을 담는 문헌용."""
    out, acc, last_page = [], 0, None
    for i in range(start, end):
        u = units[i]
        page = (u.get("m") or "")[:-3]
        if acc >= target and page and page != last_page:
            out.append(i)
            acc = 0
        acc += len(norm_search(" ".join(u["cn"])))
        if page:
            last_page = page
    if out and (end - out[-1]) < 3:
        out.pop()
    return out


def chunk_by_page(units, start, end, target=90):
    """표제가 없는 구간을 위치표지의 쪽 번호로 끊는다.
    쪽이 바뀌는 자리에서만 자르므로 문맥 한가운데가 갈라지지 않는다."""
    breaks = []
    last_page = None
    for i in range(start, end):
        m = units[i].get("m")
        if not m:
            continue
        page = m[:-3]                      # 1206c22 → 1206
        if last_page is not None and page != last_page:
            breaks.append(i)
        last_page = page
    if not breaks:
        return []

    out, prev = [], start
    for b in breaks:
        if b - prev >= target:
            out.append(b)
            prev = b
    if out and (end - out[-1]) < target // 2:   # 꼬리가 너무 짧으면 앞에 붙인다
        out.pop()
    return out


def label_range(units, a, b):
    """구간의 위치표지 범위를 이름으로 삼는다 — 1206c22–1210b04"""
    ms = [units[i]["m"] for i in range(a, b) if units[i].get("m")]
    if not ms:
        return f"{a + 1}–{b}단위"
    return ms[0] if len(ms) == 1 else f"{ms[0]}–{ms[-1]}"


def subdivide(units, chapters, limit=150):
    """지나치게 긴 권을 쪽 단위로 잘게 나눈다."""
    if not chapters:
        return chapters
    out = []
    for n, c in enumerate(chapters):
        end = chapters[n + 1]["i"] if n + 1 < len(chapters) else len(units)
        out.append(c)
        if end - c["i"] <= limit:
            continue
        cuts = chunk_by_page(units, c["i"], end)
        for k, b in enumerate(cuts):
            stop = cuts[k + 1] if k + 1 < len(cuts) else end
            out.append({"lv": c["lv"], "i": b,
                        "t": f"{c['t']} ({k + 2}) {label_range(units, b, stop)}"})
    out.sort(key=lambda x: x["i"])
    return out


def build_sections(sections, n_units):
    """빈 절을 걷어내고, 화면 분할에 쓸 '권(chapter)' 층위를 고른다."""
    secs = [s for s in sections if 0 <= s["i"] < n_units]
    secs.sort(key=lambda x: (x["i"], x["lv"]))

    # 같은 위치에 겹친 표제는 가장 상위만 남긴다
    dedup, last_i = [], None
    for s in secs:
        if s["i"] == last_i:
            continue
        dedup.append(s)
        last_i = s["i"]
    secs = dedup

    # 내용이 없는 절 제거(다음 절과 시작 위치가 같은 경우)
    kept = []
    for n, s in enumerate(secs):
        nxt = secs[n + 1]["i"] if n + 1 < len(secs) else n_units
        if nxt > s["i"]:
            kept.append(s)
    secs = kept
    if not secs:
        return [], []

    # 짧은 문헌은 굳이 나누지 않는다. 절 표제는 본문 안 소제목으로만 쓴다.
    if n_units < 60:
        return secs, []

    # 권 층위 고르기: 2~120권으로 나뉘고 한 권이 평균 3단위 이상인 가장 상위 층위
    # (48권짜리 『유가론기』처럼 권수가 많은 문헌도 권 표제를 살려 쓰기 위함)
    levels = sorted({s["lv"] for s in secs})
    chapter_lv = None
    for lv in levels:
        n = sum(1 for s in secs if s["lv"] <= lv)
        if 2 <= n <= 120 and n_units / n >= 3:
            chapter_lv = lv
            break
    if chapter_lv is None:
        return secs, []

    chapters = [s for s in secs if s["lv"] <= chapter_lv]
    if chapters and chapters[0]["i"] > 0:
        chapters.insert(0, {"lv": chapter_lv, "t": "권두", "i": 0})
    return secs, chapters


# 번역본을 만든 쪽의 작업 기록. 저본에 대한 주석이 아니므로 각주에서 뺀다.
#   '검증 판정: … 감사를 모두 통과하였다.'
#   '[최종교열] 뒤로 한 단위 밀린 번역을 재배열하였다.'
RE_MAKER_LINE = re.compile(
    r"^\s*\[?\s*(?:검증\s*(?:판정|방법|결과)|최종\s*교열|전수\s*교열\s*완료"
    r"|문단\s*매니페스트|작업\s*기록)")
RE_MAKER_ANY = re.compile(
    r"검증\s*(?:판정|방법)\s*[:：|]|감사를\s*모두\s*통과|JSONL|매니페스트"
    r"|전수교열\s*완료|고정\s*U\s*분절|U\s*단위를\s*연속\s*범위"
    r"|내부\s*감사\s*데이터|작업용\s*U\s*번호를\s*노출")


def maker_line(t: str) -> bool:
    return bool(RE_MAKER_LINE.match(t) or RE_MAKER_ANY.search(t))


# 각주 끝에 꼬리표로 붙는 작업 기록. '[최종 전수교열]' 따위.
RE_MAKER_TAG = re.compile(
    r"\s*\[\s*(?:최종\s*전수\s*교열|전수\s*교열|최종\s*교열|전수\s*대조"
    r"|최종\s*검수|교열\s*완료)\s*\]")


def drop_maker(text: str) -> str:
    """한 각주 안에서 작업 기록 토막만 덜어낸다.

    '[구문 불확실] … | [최종교열] 뒤로 한 단위 밀린 번역을 재배열하였다.'
    처럼 저본 주석 뒤에 붙어 오는 일이 많아, 줄과 세로줄로 토막을 갈라
    작업 기록에 해당하는 토막만 뺀다."""
    out = []
    for line in drop_u_jargon(text).split("\n"):
        line = RE_MAKER_TAG.sub("", line)
        segs = [x for x in re.split(r"\s*\|\s*", line) if x.strip()]
        segs = [x for x in segs if not maker_line(x)]
        if segs:
            out.append(" | ".join(segs))
    return "\n".join(out).strip()


# 번역문 안에 대괄호로 끼워 넣은 역주.
#   '[판독·구문 불확실: 원문 何時導有心神…]'  → 설명이 딸린 것은 각주로
#   '[구문 불확실]'                          → 자리만 가리키는 표는 본문에 둔다
# 안쪽에 '[4]' 같은 교감 번호가 들어 있어도 잡히도록 한 겹까지 허용한다.
RE_BRACKET = re.compile(r"\[(?:[^\[\]]|\[[^\[\]]*\])*\]")
RE_MEMO_HEAD = re.compile(r"^\[\s*([^:：\[\]]{2,24})\s*[:：]\s*(.+)\]$", re.S)


def move_reading_memos(units):
    """번역문 끝이나 사이에 끼워 둔 역주를 각주로 옮긴다.

    '[꼬리표: 설명]' 꼴만 옮기고, '[구문 불확실]' 처럼 설명 없이 자리만
    가리키는 표는 그 자리에 그대로 둔다."""
    for u in units:
        moved = []

        def take(mm):
            g = mm.group(0)
            m2 = RE_MEMO_HEAD.match(g)
            if not m2 or len(g) < 12:
                return g
            label, body = m2.group(1).strip(), m2.group(2).strip()
            moved.append(f"[{label}] {body}")
            return ""

        def strip_memos(x):
            n0 = len(moved)
            y = re.sub(r"\s{2,}", " ", RE_BRACKET.sub(take, x)).strip()
            if len(moved) > n0:             # 메모를 뺀 자리에 남은 ' .' 따위를 붙인다
                y = re.sub(r"\s+([.,;:!?。，、])", r"\1", y)
                y = re.sub(r"([,，、;])\1+", r"\1", y)
            return y

        u["ko"] = [x for x in (strip_memos(x) for x in u.get("ko", [])) if x]
        if moved:
            u.setdefault("nt", []).extend(moved)
            if "ntd" in u:
                u["ntd"].extend([""] * len(moved))


# 각주에 남은 제작 쪽 약호 'U'. 화면에서는 '단위'로 부르므로 그렇게 고친다.
#   'U단위' → '단위'  ·  'U 경계' → '단위 경계'  ·  '한 U 안에' → '한 단위 안에'
RE_U_COMPOUND = re.compile(r"(?<![A-Za-z0-9])U\s*(단위|분절)")
RE_U_ALONE = re.compile(r"(?<![A-Za-z0-9])U(?![A-Za-z0-9+])")


def drop_u_jargon(text: str) -> str:
    t = RE_U_COMPOUND.sub(r"\1", str(text))
    t = RE_U_ALONE.sub("단위", t)
    t = re.sub(r"단위\s*(단위|번호|경계|대응|분할|분절)", r"단위 \1", t)
    t = re.sub(r"단위\s+단위", "단위", t)
    return t


def strip_maker_notes(units):
    """각주에서 번역본 제작 기록만 걷어낸다. 저본 주석은 건드리지 않는다.

    각주는 문헌에 따라 두 꼴로 담긴다.
      · nt = ["요지"] + ntd = ["상세"]            (CBETA 갈래)
      · nt = [{"t": 요지, "d": [상세…]}] 또는 ["한 줄"]  (한불전 갈래)
    둘 다 꼴을 그대로 지키면서 손질한다."""
    for u in units:
        nt = u.get("nt")
        if not nt:
            continue

        if any(isinstance(x, dict) for x in nt):      # 한불전 갈래
            out = []
            for a in nt:
                if isinstance(a, dict):
                    t = drop_maker(a.get("t", ""))
                    if not t:
                        continue
                    d = [x for x in (drop_maker(y) for y in a.get("d", [])) if x]
                    out.append({"t": t, "d": d} if d else t)
                else:
                    t = drop_maker(a)
                    if t:
                        out.append(t)
            u["nt"] = out
            continue

        ntd = u.get("ntd")
        if ntd is None:                               # 상세 칸이 없는 꼴
            u["nt"] = [x for x in (drop_maker(a) for a in nt) if x]
            continue

        keep_t, keep_d = [], []
        for a, b in zip(nt, ntd):
            a2 = drop_maker(a)
            if not a2:
                continue                      # 요지가 통째로 작업 기록이었다
            keep_t.append(a2)
            keep_d.append(drop_maker(b))
        u["nt"], u["ntd"] = keep_t, keep_d


# ── 읽는 사람에게 보일 말로 ─────────────────────────────────────────────
# 교감·외자 메모에는 XML 태그 이름, 파일 경로, 내부 식별자 같은 기계용 말이
# 섞여 들어오기 쉽다. 뜻은 그대로 두고 말만 우리말로 바꾼다.
PLAIN_RULES = [
    # 내부 앵커 이름(nkr_note_orig_0440003 따위)은 읽는 사람에게 쓸모가 없다
    (r"[,，]?\s*(?:본문\s*)?앵커\s*nkr_[\w가-힣]+", ""),
    (r"\s*[/,]?\s*nkr_[\w가-힣]+", ""),
    (r"\s*/\s*앵커(?=\s*[:：])", ""),
    (r'<g ref="#(CB\d+)"[^>]*>', r"\1"),
    (r"</?(?:charDecl|figure|graphic|app|note|lem|rdg|anchor|ref)\b[^>]*>", ""),
    (r'[^.;]*resp\s*=\s*"?#?resp\d[^.;]*[.;]?', ""),
    (r"(?:[TXAJ]/)?([TXAJ]\d+n\d+\w*)\s*(?:\(\d+\))?\s*\.xml", r"\1"),
    (r"[\w/]*\.xml", ""),
    (r"(?i)note(?:\s*(?:및|and)\s*app)?\s*n\s*=\s*([0-9A-Za-z]{4,})", r"교감 번호 \1"),
    (r"\bn\s*=\s*([0-9A-Za-z]{4,})", r"교감 번호 \1"),
    (r"(?i)note_star\s*anchor|fx\s*anchor|note\s*target|corresp", "상호참조 표지"),
    (r"(?i)\bfx[TX]?\w*", "상호참조 표지"),
    (r"(?i)charDecl", "CBETA 외자표"),
    (r"(?i)normalized\s*form", "정규형"),
    (r"(?i)PUA[^.;,)]*", ""),
    (r"(?i)unicode", "유니코드"),
    (r"(?i)(?<![A-Za-z])(?:lemma)(?![A-Za-z])", "채택 독법"),
    (r"(?i)(?<![A-Za-z])(?:witness)(?![A-Za-z])", "판본"),
    (r"(?i)(?<![A-Za-z])(?:anchor)(?![A-Za-z])", "표지"),
    (r"(?i)(?<![A-Za-z])(?:app)(?![A-Za-z])", "교감"),
    (r"(?i)(?<![A-Za-z])(?:note)(?![A-Za-z])", "교감주"),
    (r"(?i)(?<![A-Za-z])(?:figure|graphic)(?![A-Za-z])", "도판"),
    (r"사용자(?:\s*제공)?\s*(?:구버전\s*)?TXT|현행\s*TXT|현\s*TXT|구버전\s*TXT|저본\s*TXT", "저본"),
    (r"(?<![A-Za-z])TXT(?![A-Za-z])", "저본"),
    (r"공식\s*XML|현행\s*XML|CBETA\s*XML|XML\s*헤더", "CBETA 자료"),
    (r"(?<![A-Za-z])XML(?![A-Za-z])", "CBETA 자료"),
    (r"(?<![A-Za-z])(?:JSON|DOCX|ZIP)(?![A-Za-z])", ""),
    (r"(?i)\btype\s*=\s*\"?(\w+)\"?", r"\1"),
    (r"<[^<>\n]{1,60}>", ""),
    (r"CBETA\s+CBETA", "CBETA"),
    (r"CBETA 자료\s+CBETA 자료", "CBETA 자료"),
    (r"저본\s+저본", "저본"),
    (r"(?i)(?<![A-Za-z])orig(?![A-Za-z])", "원교감"),
    (r"(?i)(?<![A-Za-z])add(?![A-Za-z])", "추가 교감"),
    (r"(?i)(?<![A-Za-z])mod(?![A-Za-z])", "수정 교감"),
    (r"원교감\s+원교감", "원교감"),
    (r"[;,]?\s*유니코드[^.;,()]*", ""),
    (r"\s*U\+[0-9A-Fa-f]{4,}", ""),
    (r"도판\s*/\s*도판", "도판"),
    (r"\(\s*[,;]\s*", "("),
    (r"\s*,\s*\)", ")"),
    (r"저본를", "저본을"), (r"저본는", "저본은"), (r"저본가", "저본이"), (r"저본와(?=\s|$)", "저본과"),
    (r"\s+([,.;)])", r"\1"),
    (r"^([^()]*)\)", r"\1"),                    # 짝 잃은 닫는 괄호
    (r"\(([^()]*)$", r"\1"),                    # 짝 잃은 여는 괄호
    (r"\s{2,}", " "),
]
PLAIN_RULES = [(re.compile(a), b) for a, b in PLAIN_RULES]


def plain_words(t):
    t = t.replace("`", "")
    for rx, rep in PLAIN_RULES:
        t = rx.sub(rep, t)
    t = re.sub(r"\s{2,}", " ", t)
    t = re.sub(r"\(\s*\)|\[\s*\]|,\s*(?=[.;,])", "", t)
    return t.strip(" ;,")


RE_MIN_ADD = re.compile(r"^\[최소 보충\]\s*(.{1,25})$")


RE_HANGUL = re.compile(r"[가-힣]")
RE_SENT_END = re.compile(r"(?<=[다요]\.)\s+|(?<=[.!?])\s+(?=[가-힣])")


def _ko_gist(d):
    """상세의 저본·이문 독법으로 한 줄 요지를 만든다."""
    kv = {}
    for x in d:
        k, sep, v = x.partition(": ")
        if sep and k not in kv:
            kv[k] = v.strip()
    a = kv.get("저본 독법")
    b = kv.get("이문 독법") or kv.get("교감 제안 독법")
    if not a or not b or "해당 없음" in b:
        return None
    def cut(v):
        v = v.split(" / ")[0]                            # 한 칸에 여러 항목을 이어 쓴 꼴
        v = re.sub(r"[（(][^）)]*[）)]", "", v)          # 괄호 설명은 뺀다
        v = re.sub(r"^(?:저본|현행 CBETA|대정장|사용자)\S*\s*의\s*", "", v).strip()
        v = re.sub(r"\s{2,}", " ", v).strip(" .·")
        return v if len(v) <= 24 else v[:24] + "…"
    m = re.match(r"^(.*?)\s*[（(\[]?\s*【([^】]{1,6})】\s*[)）\]]?$", b)
    wit = m.group(2) if m else None
    b2 = m.group(1) if wit else b
    return f"저본 {cut(a)} → {wit + '본 ' if wit else '이본 '}{cut(b2)}"


def plain_note(n):
    if isinstance(n, str):
        t = plain_words(n)
        if len(t) > 90:      # 긴 한 줄 메모는 요지 + 펼침으로 나눈다
            tag, sep, body = t.partition("] ")
            parts = [x.strip() for x in RE_SENT_END.split(body) if x.strip()]
            if sep and len(parts) > 1:
                head = parts[0] if len(parts[0]) <= 90 else parts[0][:88] + "…"
                return {"t": f"{tag}] {head}", "d": parts[1:] if head == parts[0] else parts}
        m = RE_MIN_ADD.match(t)
        # 「[최소 보충] 대상을」처럼 보충한 말만 떠 있으면 문장으로 적는다
        return f"[최소 보충] 번역에서 보충한 말: {m.group(1)}" if m else t
    d = [x for x in (plain_words(x) for x in n["d"]) if x]
    t = plain_words(n["t"])
    # 요지가 한문만 적혀 있으면(「[CBETA 원교감] 一無品字」) 우리말 한 줄로 바꾸고
    # 한문 교감문은 상세로 내린다
    tag, sep, body = t.partition("] ")
    if sep and not RE_HANGUL.search(body):
        if re.fullmatch(r"CB\d{4,6}", body):   # 외자 번호만 있는 요지
            src = next((x.split(": ", 1)[1] for x in d if x.startswith("저본 표기: ")), "")
            head = f"{tag}] 저본의 외자 {src}".rstrip() if src else f"{tag}] 저본의 외자 {body}"
            return {"t": head, "d": d} if d else head
        g = _ko_gist(d)
        if g:
            if not any(body in x for x in d):
                d = [f"원교감 원문: {body}"] + d
            t = f"{tag}] {g}"
    return {"t": t, "d": d} if d else t


def build_work(entry):
    wid = entry["id"]
    wdir = SRC / wid
    txt_units, dx = [], None

    # 원문은 여러 개일 수 있다. 회본(본문 + 주석서)이 그런 경우로,
    #   원문-1-십이문론.txt / 원문-2-종치의기.txt
    # 처럼 이름을 붙이면 파일명 순서대로 이어 붙인다.
    tps = sorted(list(wdir.glob("원문*.txt")) + list(wdir.glob("원문*.docx")))
    for tp in tps:
        part = parse_source_docx(tp) if tp.suffix == ".docx" else parse_txt(tp)
        base = len(txt_units)
        for u in part:
            u["i"] += base
            u["src"] = tp.stem
        txt_units.extend(part)
    dp = wdir / "번역.docx"
    kabc = entry.get("source") == "KABC"
    if kabc:
        # 한국불교전서 계열은 저본 txt 가 15~16자마다 끊겨 있어 본문으로
        # 쓸 수 없다. 번역 docx 안의 원문이 이미 문단으로 이어져 있으므로
        # 그쪽 하나로 단위를 세운다. (저본 txt 는 대조 근거로만 둔다)
        dx = parse_kabc_docx(dp)
        dx.setdefault("sections", dx.get("secs", []))
        dx.setdefault("appendix", [])
        dx.setdefault("tables", read_docx_tables(dp))
        units, dxmap = dx["units"], {}
    elif dp.exists():
        if is_table_aligned(dp):
            dx = parse_table_docx(dp)
        elif is_translation_only(dp):
            dx = parse_translation_only_docx(dp)
        else:
            dx = parse_docx(dp)

    if not kabc:
        units, dxmap = merge(txt_units, dx)
        units = split_note_items(units)

    # ── 분권 구성 ────────────────────────────────────────────────
    raw_secs = []
    if dx and not kabc:
        for sec in dx["sections"]:
            j = dxmap.get(sec["i"])
            if j is None:
                # 대응 못 찾은 표제는 그 뒤 첫 대응 지점으로 민다
                later = [dxmap[k] for k in sorted(dxmap) if k >= sec["i"]]
                if not later:
                    continue
                j = later[0]
            raw_secs.append({"lv": sec["lv"], "t": sec["t"], "i": j})
    if not raw_secs:
        raw_secs = detect_juan_from_source(units)

    sections, chapters = build_sections(raw_secs, len(units))

    # 번역 docx가 권마다 표제를 달아 두지 않은 문헌은, 권 나눔이 엉뚱하게 잡힌다.
    # (한 권 제목이 뒤 권까지 끌려가 '…(2) (3)'처럼 되풀이된다)
    # 이럴 때는 원문에 남아 있는 권 표제를 따르는 편이 낫다.
    # 번역 docx가 권마다 표제를 달아 두면 그쪽이 한국어 병기까지 있어 낫다.
    # 표제가 거의 없을 때만(=권이 3개 이하로 잡힐 때) 원문 표제로 갈아탄다.
    if len(chapters) <= 3:
        juan = detect_juan_from_source(units)
        if len(juan) >= 4 and len(juan) > len(chapters):
            _, chapters = build_sections(juan, len(units))


    # 단위는 적지만 글자 수가 많은 문헌(묶음형 번역)도 나눈다
    total_chars = sum(len(norm_search(" ".join(u["cn"]))) for u in units)
    if not chapters and total_chars > 40000 and len(units) <= 120:
        cuts = [0] + chunk_by_chars(units, 0, len(units))
        if len(cuts) > 1:
            chapters = [
                {"lv": 1, "i": b,
                 "t": label_range(units, b,
                                  cuts[k + 1] if k + 1 < len(cuts) else len(units))}
                for k, b in enumerate(cuts)]

    # 표제가 없어 권이 안 잡힌 긴 문헌은 위치표지 쪽으로 나눈다
    if not chapters and len(units) > 120:
        cuts = [0] + chunk_by_page(units, 0, len(units))
        if len(cuts) > 1:
            chapters = [
                {"lv": 1, "i": b,
                 "t": label_range(units, b,
                                  cuts[k + 1] if k + 1 < len(cuts) else len(units))}
                for k, b in enumerate(cuts)]

    # 권 하나가 너무 길면 더 잘게
    chapters = subdivide(units, chapters)

    # 각 단위에 소속 권 번호 부여
    if chapters:
        ci, nxt = 0, (chapters[1]["i"] if len(chapters) > 1 else len(units))
        for idx, u in enumerate(units):
            while ci + 1 < len(chapters) and idx >= chapters[ci + 1]["i"]:
                ci += 1
            u["c"] = ci

    # ── 통계 ─────────────────────────────────────────────────────
    n_ko = sum(1 for u in units if u["ko"])
    chars_cn = sum(len(norm_search(" ".join(u["cn"]))) for u in units)
    chars_done = sum(len(norm_search(" ".join(u["cn"]))) for u in units if u["ko"])
    chars_ko = sum(len("".join(u["ko"])) for u in units)
    markers = sorted({u["m"] for u in units if u["m"] and not u.get("x")})

    # 저본에 실린 도판을 그 자리의 단위에 붙인다.
    # registry 의 figures = {"위치표지": "파일명"} 을 따르고,
    # 한 자리에 도판이 여럿이면 {"위치표지": ["파일1", "파일2", …]} 로 적는다.
    # 파일은 assets/figures/<문헌id>/ 에 둔다.
    figs = entry.get("figures") or {}
    if figs:
        first, last = {}, {}
        for u in units:
            first.setdefault(u.get("m"), u)
            last[u.get("m")] = u
        marks = sorted(m for m in first if m)
        flat_figs = []
        for marker, val in sorted(figs.items()):
            names = val if isinstance(val, list) else [val]
            flat_figs.extend((marker, n) for n in names)
        for marker, fname in flat_figs:
            u = first.get(marker)
            if u is None:
                # 저본에서 도판만 놓인 표지는 원문이 비어 단위가 서지 않는다.
                # 그 앞 표지의 마지막 단위 끝에 붙여 저본과 같은 자리에 오게 한다.
                prev = [m for m in marks if m < marker]
                u = last.get(prev[-1]) if prev else None
            if u is None:
                print(f"  ! {wid}: 도판 {fname} 의 자리 [{marker}] 를 찾지 못했습니다")
                continue
            u.setdefault("fig", []).append(fname)
            # 도판을 실제로 띄우므로, 그 자리를 말로 때운 번역 문단
            # ('문자 본문은 없으며, 이 위치에 …가 배치된다')은 군더더기가 된다.
            u["ko"] = [t for t in u["ko"] if not RE_FIG_CAPTION.search(t)]

    # 원문 TXT 가 '[X45p0782_01.gif]' 로 적어 둔 자리에도 그림을 끼운다
    figdir = ROOT / "assets" / "figures" / wid
    if figdir.is_dir():
        # registry 가 자리를 정해 준 도판은 단위 아래에 따로 놓이므로 제외한다
        placed = set()
        for v in (entry.get("figures") or {}).values():
            placed.update(v if isinstance(v, list) else [v])
        have = {f.name for f in figdir.iterdir()
                if f.is_file() and f.name not in placed}

        def put(mm):
            for ext in (".gif", ".png", ".jpg"):
                if mm.group(1) + ext in have:
                    return FIG_TOKEN.format(mm.group(1) + ext)
            return ""

        for u in units:
            u["cn"] = [RE_FIG_NAME.sub(put, x) for x in u["cn"]]
            u["ko"] = [RE_FIG_NAME.sub(put, x) for x in u["ko"]]

    move_reading_memos(units)
    strip_maker_notes(units)

    meta = dict(entry)
    units_marks = [u["m"] for u in units if u.get("m")]
    meta.pop("figures", None)
    meta.update({
        "units": len(units),
        "translated": n_ko,
        "coverage": round(chars_done / chars_cn, 4) if chars_cn else 0,
        "chars_cn": chars_cn,
        "chars_ko": chars_ko,
        # 위치표지는 대개 사전순이 곧 문헌 순서지만, 한불전의 張 표지는
        # 「第9張」이 「第10張」보다 뒤로 밀리므로 본문에 나온 차례를 따른다.
        # (회본처럼 원문을 이어 붙인 문헌은 사전순 쪽이 낫다)
        "range": ([units_marks[0], units_marks[-1]] if kabc and units_marks
                  else ([markers[0], markers[-1]] if markers else None)),
        "has_source": bool(txt_units),
        "has_translation": bool(dx),
        "chapters": len(chapters),
    })
    if dx:
        gl = pick_glossary(dx["tables"])
        meta["biblio_extracted"] = pick_biblio(dx["tables"])
    else:
        gl = []

    # 화면에 나가는 글에서 기계용 말(태그 이름·파일 경로·내부 식별자)을 걷어 낸다
    for u in units:
        u["nt"] = [x for x in (plain_note(n) for n in u["nt"]) if x]

    doc = {
        "meta": meta,
        "glossary": gl,
        "front": [x for x in ((plain_words(f) if isinstance(f, str) else f)
                              for f in ((dx["front"] + dx["appendix"]) if dx else [])) if x],
        "chapters": [{"t": c["t"], "i": c["i"]} for c in chapters],
        "sections": [{"lv": s["lv"], "t": s["t"], "i": s["i"]} for s in sections],
        "units": units,
    }
    (OUT / "works").mkdir(parents=True, exist_ok=True)
    with open(OUT / "works" / f"{wid}.json", "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, separators=(",", ":"))
    return meta


def discover(reg):
    """sources/ 안에 있으나 registry 에 없는 문헌을 자동 편입한다.
    서지를 몰라도 일단 읽히도록 최소 정보만 채워 둔 뒤,
    registry.json 에 항목을 추가하면 그때부터 정식 서지가 적용된다."""
    known = {w["id"] for w in reg["works"]}
    added = []
    if not SRC.exists():
        return added
    for d in sorted(SRC.iterdir()):
        if not d.is_dir() or d.name in known:
            continue
        if not any((d / f).exists() for f in ("원문.txt", "번역.docx")):
            continue
        added.append({
            "id": d.name,
            "title_cn": d.name,
            "title_ko": d.name,
            "author_cn": "", "author_ko": "미상", "dynasty": "",
            "canon": "서지 미기재", "canon_label": "미분류",
            "canon_id": "Z", "canon_vol": 999, "canon_no": 999,
            "canon_ko": "", "collection": "미분류", "tags": [],
            "verify": True,
        })
    return added


def sync_variant_tables(root: Path) -> None:
    """이체자 표를 화면 쪽 js 에도 그대로 옮겨 적는다.

    검색 색인은 여기서(파이썬), 검색어 정규화는 브라우저(js)에서 이루어진다.
    두 표가 어긋나면 원문에 있는 글자를 찾지 못하므로, 표는 이 파일 하나만
    고치고 나머지는 이 함수가 맞춰 준다."""
    body = ",\n  ".join(
        f"'{k}':'{v}'" for k, v in sorted(VARIANTS.items()) if k != v)
    targets = {"assets/js/search.worker.js": "const VAR = {",
               "assets/js/app.js": "const VAR_FOLD = {"}
    for rel, head in targets.items():
        f = root / rel
        if not f.exists():
            continue
        t = f.read_text(encoding="utf-8")
        i = t.find(head)
        if i < 0:
            print(f"  ! {rel}: 이체자 표를 찾지 못했습니다")
            continue
        j = t.index("};", i) + 2
        new = t[:i] + head + "\n  " + body + "\n};" + t[j:]
        if new != t:
            f.write_text(new, encoding="utf-8")
            print(f"  · {rel} 이체자 표 갱신 ({len(VARIANTS)}자)")


def main():
    reg = json.loads((ROOT / "tools" / "registry.json").read_text(encoding="utf-8"))
    fresh = discover(reg)
    if fresh:
        print("registry 에 없는 문헌을 자동 편입합니다 "
              "(tools/registry.json 에 서지를 채워 주세요):")
        for e in fresh:
            print("   +", e["id"])
        reg["works"] = reg["works"] + fresh
    sync_variant_tables(ROOT)

    metas, panbon_warn = [], []
    for entry in reg["works"]:
        m = build_work(entry)
        metas.append(m)
        print(f"  {m['id']:<22} 단위 {m['units']:>5}  번역 {m['translated']:>5}"
              f"  ({m['coverage']*100:5.1f}%)  {m['range']}")
        # 본문에 판본기호({底}{甲}…)가 남아 있으면 교감문을 본문으로
        # 잘못 읽은 것이다. 자동으로 고치지 않고 알리기만 한다.
        wp = ROOT / "data" / "works" / f"{m['id']}.json"
        if wp.exists():
            wd = json.loads(wp.read_text(encoding="utf-8"))
            for u in wd.get("units", []):
                body = "".join(u.get("cn", [])) + "".join(u.get("ko", []))
                if RE_PANBON.search(body):
                    panbon_warn.append((m["id"], u["i"], u.get("m")))
    if panbon_warn:
        print()
        print("⚠ 본문에 판본기호({底}{甲}…)가 남아 있습니다. "
              "교감문을 본문으로 잘못 읽은 자리입니다.")
        print("  번역 docx 쪽을 고쳐 주십시오(자동으로 잘라내지 않습니다):")
        for wid, i, mk in panbon_warn:
            print(f"   · {wid}  단위 {i}  ({mk})")
        print()

    # 대장경 순서(T → X → L, 책 번호, 경 번호)로 정렬해 둔다
    CANON_ORDER = {"T": 0, "X": 1, "L": 2, "K": 3, "B": 4, "ZW": 5, "BJ": 6, "HB": 6}
    metas.sort(key=lambda m: (CANON_ORDER.get(m.get("canon_id", ""), 9),
                              m.get("canon_vol", 0), m.get("canon_no", 0)))
    manifest = {
        "generated": reg.get("version", "1"),
        "site": reg.get("site", {}),
        "rights": reg.get("rights", {}),
        "works": [{k: v for k, v in m.items() if k != "biblio_extracted"} for m in metas],
        "totals": {
            "works": len(metas),
            "units": sum(m["units"] for m in metas),
            "translated": sum(m["translated"] for m in metas),
            "chars_cn": sum(m["chars_cn"] for m in metas),
        },
    }
    with open(OUT / "manifest.json", "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    # 재배포 조건 점검 — 출처마다 조건이 다르므로 빌드할 때마다 확인한다
    srcinfo = reg.get("rights", {}).get("sources", {})
    flagged = {}
    for m in metas:
        src = m.get("source", "CBETA")
        info = srcinfo.get(src, {})
        if info.get("ok") is not True:
            flagged.setdefault(src, []).append(m["id"])
    for src, ids in flagged.items():
        info = srcinfo.get(src, {})
        mark = "재배포 금지" if info.get("ok") is False else "조건 확인 필요"
        print(f"\n[주의] 출처 {src} — {mark}")
        print(f"        {info.get('condition', '이용 조건을 확인하십시오.')}")
        print(f"        해당 문헌: {', '.join(ids)}")

    print(f"\n총 {manifest['totals']['works']}종 · "
          f"{manifest['totals']['units']}단위 · "
          f"원문 {manifest['totals']['chars_cn']:,}자")


if __name__ == "__main__":
    main()
