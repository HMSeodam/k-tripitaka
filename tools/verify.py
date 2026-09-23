#!/usr/bin/env python3
"""수록 문헌 전수 점검.

data/works/*.json 을 저본(sources/<id>/)·registry 와 맞대어 보고
문헌마다 아래를 센다.

  · 원문 대조    원문 txt 와 앱 원문이 글자까지 같은지 (위치표지·공백 제외)
  · 번역 공백    번역이 붙지 않은 단위
  · 교감 공백    교감표지가 있는데 메모가 없는 단위
  · 도판         깨진 이미지 경로
  · 찌꺼기       본문에 남은 제작 메모·각주 번호 따위

결과는 data/verify-report.json 에 적는다. 커밋마다 이 파일의 차이를 보면
어떤 문헌이 왜 바뀌었는지 드러난다.

    python3 tools/verify.py            # 전체
    python3 tools/verify.py <문헌id>…  # 일부만
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WORKS = ROOT / "data" / "works"
SOURCES = ROOT / "sources"
FIGURES = ROOT / "assets" / "figures"
REPORT = ROOT / "data" / "verify-report.json"

RE_MARK = re.compile(r"\[\d{3,4}[abc]\d{2}\]")
RE_APP = re.compile(r"\[(?:\d{1,3}|＊|\*)\]")
RE_FIG = re.compile(r"\u27e6fig:([^\u27e7]+)\u27e7")
RE_SPACE = re.compile(r"[\s\u3000]")
# 본문에 남으면 안 되는 제작 흔적
RE_JUNK = re.compile(r"구 XML|현재 XML|처리 수\s*:|공통으로 존재하는 교감|미확보 이미지 수")


def flat(txt: str) -> str:
    return RE_SPACE.sub("", RE_MARK.sub("", txt))


def source_text(wid: str):
    """(저본 글, 종류). 종류가 'ref' 면 대조 근거일 뿐 정본이 아니다.

    한국불교전서 계열은 저본 txt 가 15~16자마다 끊겨 있어 번역 docx 쪽을
    본문으로 삼는다. 그런 문헌은 글자 수가 어긋나는 것이 정상이다.
    """
    p = SOURCES / wid / "원문.txt"
    if p.exists():
        return p.read_text(encoding="utf-8"), "base"
    p = SOURCES / wid / "저본.txt"
    if p.exists():
        return p.read_text(encoding="utf-8"), "ref"
    return None, None


def check(wid: str) -> dict:
    d = json.loads((WORKS / f"{wid}.json").read_text(encoding="utf-8"))
    units = d["units"]
    cn = "".join("".join(u["cn"]) for u in units)
    r: dict = {
        "units": len(units),
        "translated": sum(1 for u in units if u["ko"]),
        "notes": sum(len(u["nt"]) for u in units),
        "markers": len({u["m"] for u in units if u["m"]}),
    }

    txt, kind = source_text(wid)
    r["source_kind"] = kind
    if txt is None or kind == "ref":
        r["source_match"] = None          # 정본 txt 가 없는 문헌은 대조 대상이 아니다
    else:
        a, b = flat(txt), flat(RE_FIG.sub("", cn))
        r["source_match"] = a == b
        if a != b:
            r["source_diff_chars"] = abs(len(a) - len(b)) or None

    r["no_translation"] = [u["m"] or f"#{u['i']}" for u in units if not u["ko"]][:20]
    r["no_translation_count"] = sum(1 for u in units if not u["ko"])

    def note_text(n):
        return n if isinstance(n, str) else n["t"] + " " + " ".join(n["d"])

    r["apparatus_without_note"] = sum(
        1 for u in units
        if RE_APP.search("".join(u["cn"]))
        and not any(re.search(r"교감|apparatus|CBETA", note_text(n)) for n in u["nt"])
    )

    figs = {f.name for f in (FIGURES / wid).iterdir()} if (FIGURES / wid).is_dir() else set()
    used = set(RE_FIG.findall(cn)) | {
        g for u in units for g in (u.get("fig") or [])
        if isinstance(g, str)
    }
    r["figures_used"] = len(used)
    r["figures_missing"] = sorted(g for g in used if g not in figs)[:10]
    r["figures_unused"] = len(figs - used)

    r["junk_in_notes"] = sum(
        1 for u in units for n in u["nt"] if RE_JUNK.search(note_text(n))
    )
    r["bare_number_in_ko"] = sum(
        len(re.findall(r"(?<!No)[다.,’”」)!?]\s?\d{1,3}(?=[\s‘“「(]|$)", k))
        for u in units for k in u["ko"]
    )
    return r


def main(argv: list[str]) -> int:
    reg = json.loads((ROOT / "tools" / "registry.json").read_text(encoding="utf-8"))
    ids = argv[1:] or [w["id"] for w in reg["works"]]
    out, flags = {}, []
    for wid in ids:
        if not (WORKS / f"{wid}.json").exists():
            flags.append((wid, "앱 데이터 없음"))
            continue
        r = check(wid)
        out[wid] = r
        if r["source_match"] is False:
            flags.append((wid, f"원문 불일치 ({r.get('source_diff_chars') or '순서'})"))
        if r["apparatus_without_note"]:
            flags.append((wid, f"교감표지만 있고 메모 없음 {r['apparatus_without_note']}단위"))
        if r["figures_missing"]:
            flags.append((wid, f"이미지 없음 {len(r['figures_missing'])}건"))
        if r["junk_in_notes"]:
            flags.append((wid, f"제작 메모 잔류 {r['junk_in_notes']}건"))
        if r["no_translation_count"] > max(5, r["units"] * 0.02):
            flags.append((wid, f"번역 없는 단위 {r['no_translation_count']}개"))
        if r["bare_number_in_ko"]:
            flags.append((wid, f"번역문에 숫자만 남은 곳 {r['bare_number_in_ko']}건"))

    REPORT.write_text(json.dumps(out, ensure_ascii=False, indent=1, sort_keys=True) + "\n",
                      encoding="utf-8")
    print(f"{len(out)}종 점검 → {REPORT.relative_to(ROOT)}")
    for wid, why in flags:
        print(f"  ! {wid:32s} {why}")
    if not flags:
        print("  이상 없음")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
