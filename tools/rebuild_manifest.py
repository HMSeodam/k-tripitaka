#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
manifest 재생성기
------------------------------------------------------------------
data/works/ 안의 JSON 파일만 보고 data/manifest.json 을 다시 만든다.

언제 쓰는가
  · 이미 만들어 둔 대조 JSON 을 data/works/ 에 직접 떨궜을 때
  · 문헌을 하나 지웠을 때
  · 목록이 실제 파일과 어긋난 것 같을 때

ingest.py 를 돌리면 manifest 도 함께 갱신되므로, 원문·번역 원본에서
새로 만드는 경우에는 이 도구를 따로 부를 필요가 없다.

    python3 tools/rebuild_manifest.py
"""

import json
import re
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WORKS = ROOT / "data" / "works"
OUT = ROOT / "data" / "manifest.json"
REG = ROOT / "tools" / "registry.json"

# 대장경 나열 순서 (대정장 → 만신찬속장 → 건륭장 → 고려장 → 보편 → 장외)
CANON_ORDER = {"T": 0, "X": 1, "L": 2, "K": 3, "B": 4, "ZW": 5}

# manifest 에 실을 서지·통계 항목
KEEP = [
    "id", "title_cn", "title_ko", "author_cn", "author_ko", "dynasty",
    "canon", "canon_id", "canon_vol", "canon_no", "canon_label", "canon_ko",
    "source", "底本_publisher",
    "collection", "tags", "verify", "note",
    "units", "translated", "coverage", "chars_cn", "chars_ko",
    "range", "has_source", "has_translation", "chapters",
]


RE_MARKER = re.compile(r"\[(\d{3,4}[abc]\d{2})\]")
RE_APPARATUS = re.compile(r"\[(?:\d{1,3}|＊|\*)\]")
RE_DROP = re.compile(r"[\s。，、．・？！：；「」『』（）()〔〕【】\[\]“”‘’·…—　]")


def body_len(parts):
    """ingest.py 와 같은 기준으로 순수 본문 글자만 센다
    (위치표지·교감표지·구두점 제외). 두 도구의 통계를 일치시키기 위함."""
    t = unicodedata.normalize("NFKC", " ".join(parts))
    t = RE_APPARATUS.sub("", RE_MARKER.sub("", t))
    return len(RE_DROP.sub("", t))


def stats_from_units(doc):
    """meta 가 부실한 JSON 을 위해 본문에서 통계를 직접 뽑는다."""
    units = doc.get("units", [])
    n_ko = sum(1 for u in units if u.get("ko"))
    cn = sum(body_len(u.get("cn", [])) for u in units)
    done = sum(body_len(u.get("cn", [])) for u in units if u.get("ko"))
    ko = sum(len("".join(u.get("ko", []))) for u in units)
    marks = sorted({u["m"] for u in units if u.get("m")})
    return {
        "units": len(units),
        "translated": n_ko,
        "coverage": round(done / cn, 4) if cn else 0,
        "chars_cn": cn,
        "chars_ko": ko,
        "range": [marks[0], marks[-1]] if marks else None,
        "has_source": bool(units),
        "has_translation": n_ko > 0,
        "chapters": len(doc.get("chapters", [])),
    }


def main():
    if not WORKS.exists():
        raise SystemExit(f"{WORKS} 가 없습니다.")

    # registry 가 있으면 서지를 여기서 보완한다
    reg = {}
    site = {"name": "신한글대장경", "tagline": "한문 불전 원문과 한국어 대조 번역"}
    rights = {}
    version = "1"
    if REG.exists():
        r = json.loads(REG.read_text(encoding="utf-8"))
        reg = {w["id"]: w for w in r.get("works", [])}
        site = r.get("site", site)
        rights = r.get("rights", {})
        version = r.get("version", version)

    metas, skipped = [], []
    for f in sorted(WORKS.glob("*.json")):
        wid = f.stem
        try:
            doc = json.loads(f.read_text(encoding="utf-8"))
        except Exception as e:
            skipped.append((wid, f"읽기 실패: {e}"))
            continue
        if "units" not in doc:
            skipped.append((wid, "units 항목이 없습니다"))
            continue

        meta = {"id": wid}
        meta.update(reg.get(wid, {}))       # registry 서지
        meta.update(doc.get("meta", {}))    # JSON 안의 meta 가 우선
        meta.update(stats_from_units(doc))  # 통계는 본문에서 재계산
        meta["id"] = wid

        meta.setdefault("title_cn", wid)
        meta.setdefault("title_ko", wid)
        meta.setdefault("author_ko", "미상")
        meta.setdefault("author_cn", "")
        meta.setdefault("dynasty", "")
        meta.setdefault("canon", "서지 미기재")
        meta.setdefault("canon_label", "미분류")
        meta.setdefault("canon_id", "Z")
        meta.setdefault("canon_vol", 999)
        meta.setdefault("canon_no", 999)
        if meta.get("canon_id") == "Z":
            meta["verify"] = True

        metas.append({k: meta[k] for k in KEEP if k in meta})

    metas.sort(key=lambda m: (CANON_ORDER.get(m.get("canon_id", ""), 9),
                              m.get("canon_vol", 0), m.get("canon_no", 0)))

    manifest = {
        "generated": version,
        "site": site,
        "rights": rights,
        "works": metas,
        "totals": {
            "works": len(metas),
            "units": sum(m["units"] for m in metas),
            "translated": sum(m["translated"] for m in metas),
            "chars_cn": sum(m["chars_cn"] for m in metas),
        },
    }
    OUT.write_text(json.dumps(manifest, ensure_ascii=False, indent=1),
                   encoding="utf-8")

    for m in metas:
        flag = "  ← 서지 미기재" if m.get("canon_id") == "Z" else ""
        print(f"  {m.get('canon_label',''):<14} {m['title_cn'][:26]:<28}"
              f" {m['units']:>5}단위{flag}")
    if skipped:
        print("\n건너뛴 파일:")
        for wid, why in skipped:
            print(f"  ! {wid} — {why}")
    print(f"\n{len(metas)}종 · {manifest['totals']['units']:,}단위 · "
          f"원문 {manifest['totals']['chars_cn']:,}자 → data/manifest.json")


if __name__ == "__main__":
    main()
