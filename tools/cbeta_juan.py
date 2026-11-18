#!/usr/bin/env python3
"""CBETA 원본 XML(cbeta-org/xml-p5)에서 문헌마다 권(卷)이 시작하는 행을 뽑아
tools/cbeta_juan.json 에 적는다. ingest.py 가 이 표로 권을 나눈다.

원문 TXT 에는 권 표제가 빠진 문헌이 많다(『대승기신론의기』, 『대종지현문본론』,
『화엄경탐현기』 따위). CBETA XML 에는 <cb:juan n="002" fun="open"> 과 그 앞의
<lb n="0251b07"/> 로 권 경계가 행 번호까지 적혀 있어 이쪽이 가장 확실하다.

새 문헌을 registry.json 에 넣은 뒤 한 번 돌리면 된다(인터넷 필요).
    python3 tools/cbeta_juan.py            # registry 의 CBETA 문헌 가운데 표에 없는 것만
    python3 tools/cbeta_juan.py --all      # 전부 다시
GitHub Actions 는 이 스크립트를 돌리지 않는다. 결과 표(cbeta_juan.json)만 읽는다.
"""
from __future__ import annotations

import json
import re
import sys
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
REG = ROOT / "tools" / "registry.json"
OUT = ROOT / "tools" / "cbeta_juan.json"
RAW = "https://raw.githubusercontent.com/cbeta-org/xml-p5/master/"

RE_ID = re.compile(r"\b(T|X|L|K|B|J|ZW)\s*(\d+)\s*no\.\s*(\d+)\s*([a-zA-Z]?)")


def xml_names(coll: str, vol: int, no: int, suf: str):
    """xml-p5 의 파일 이름 후보: T44n1846.xml, X09n0245.xml, L141n1600.xml …"""
    vw = 3 if coll in ("L",) else 2
    nw = 4
    base = f"{coll}{vol:0{vw}d}n{no:0{nw}d}"
    folder = f"{coll}/{coll}{vol:0{vw}d}/"
    if suf:
        return [folder + base + suf + ".xml"]
    return [folder + base + ".xml"] + [folder + base + c + ".xml" for c in "abcdefgh"]


def fetch(path: str) -> str | None:
    try:
        with urllib.request.urlopen(RAW + path, timeout=60) as r:
            return r.read().decode("utf-8")
    except Exception:
        return None


def lead_text(win: str, n: int = 14) -> str:
    """권이 열린 뒤 첫 본문 몇 글자(권 제목·찬자·주석·교감 빼고).
    ingest.py 가 우리 원문에서 이 글자를 찾아 권 경계가 맞는지 확인한다."""
    t = re.sub(r"<cb:jhead.*?</cb:jhead>|<head\b.*?</head>|<byline.*?</byline>|"
               r"<note\b.*?</note>|<cb:mulu.*?</cb:mulu>|<cb:docNumber.*?</cb:docNumber>|"
               r"<rdg\b.*?</rdg>|<app\b[^>]*>|</app>|<lem\b[^>]*>|</lem>", "", win, flags=re.S)
    t = re.sub(r"<[^>]+>", "", t)
    chars = re.findall(r"[\u3400-\u9fff\uf900-\ufaff]", t)
    return "".join(chars[:n])


def juan_starts(xml: str, coll: str):
    """[{n, lb, head}] — 권이 열리는 자리의 바로 앞 행 번호와 권 제목"""
    out = []
    lbs = [(m.start(), m.group(2)) for m in re.finditer(
        r'<lb\b([^>]*)\bn="(\d{4}[a-z]\d{2})"', xml)
        if 'ed="' not in m.group(1) or f'ed="{coll}"' in m.group(1)]
    k = 0
    for m in re.finditer(r'<cb:juan\b[^>]*>', xml):
        if 'fun="open"' not in m.group(0):
            continue
        nm = re.search(r'\bn="([^"]+)"', m.group(0))
        if not nm:
            continue
        while k + 1 < len(lbs) and lbs[k + 1][0] < m.start():
            k += 1
        lb = lbs[k][1] if lbs and lbs[k][0] < m.start() else None
        win = xml[m.end():m.end() + 6000]
        head = re.search(r"<cb:jhead[^>]*>(.*?)</cb:jhead>", win, re.S)
        h = re.sub(r"<[^>]+>", "", head.group(1)).strip() if head else ""
        out.append({"n": nm.group(1), "lb": lb, "head": h, "lead": lead_text(win)})
    return out


def main(argv):
    reg = json.loads(REG.read_text(encoding="utf-8"))
    table = json.loads(OUT.read_text(encoding="utf-8")) if OUT.exists() else {}
    redo = "--all" in argv
    for w in reg["works"]:
        if w.get("source", "CBETA") != "CBETA":
            continue
        for coll, vol, no, suf in RE_ID.findall(w.get("canon", "")):
            key = f"{coll}{int(vol)}n{int(no)}{suf}"
            if key in table and not redo:
                continue
            parts, files = [], []
            # 한 책이 여러 책(冊)에 걸치면 다음 책에도 같은 번호의 파일이 있다
            for v in range(int(vol), int(vol) + 6):
                got = False
                for name in xml_names(coll, v, int(no), suf):
                    x = fetch(name)
                    if x is None:
                        if got:
                            break
                        continue
                    got = True
                    files.append(name)
                    parts.extend({**j, "vol": v} for j in juan_starts(x, coll))
                    if name.endswith(f"n{int(no):04d}.xml"):
                        break                  # 한 파일짜리
                if not got and v > int(vol):
                    break
            if not files:
                print(f"  ! {w['id']}: {key} XML 을 찾지 못함")
                continue
            table[key] = {"files": files, "juan": parts}
            print(f"  {w['id']:<32} {key:<12} 권 {len(parts)}")
    OUT.write_text(json.dumps(table, ensure_ascii=False, indent=1, sort_keys=True) + "\n",
                   encoding="utf-8")
    print(f"→ {OUT.relative_to(ROOT)} ({len(table)}종)")


if __name__ == "__main__":
    main(sys.argv)
