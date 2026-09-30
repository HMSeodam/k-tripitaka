"""앱이 쓰는 CBETA 외자 대응표(data/gaiji.json)를 만든다.

원문은 CBETA 2016판을 그대로 두므로 유니코드에 없던 글자가 조합식([卄/公/心] 따위)으로 남아 있다.
그 뒤 CBETA가 외자 자료와 보충 글꼴을 갱신해, 많은 조합식이 이제 실제 글자(䓗)로 보일 수 있다.
앱은 이 표를 보고
  · 원문의 조합식에 손을 얹거나 누르면 CBETA 보충 글꼴로 그 글자를 보여 주고,
  · 교감 메모의 외자 번호(CB01470)에도 같은 쪽지를 달며,
  · 메모에 남은 CBETA 사용자 영역(PUA) 코드는 실제 글자나 조합식으로 바꿔 보인다.
원자료: tools/cbeta_gaiji_min.json (CBETA 缺字資料庫 cbeta_gaiji 에서 뽑은 최소판)
data/works/*.json 에 실제로 나오는 것만 담는다. rebuild_manifest.py 가 끝에 부른다.
"""
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DB = ROOT / "tools" / "cbeta_gaiji_min.json"
OUT = ROOT / "data" / "gaiji.json"

RE_COMP = re.compile(r"\[[^\[\]\d\s<>]{2,40}\]")
RE_CB = re.compile(r"CB\d{5}")
RE_PUA = re.compile(r"[\U000F0000-\U000FFFFD]")


def build():
    if not DB.exists():
        return 0
    g = json.loads(DB.read_text(encoding="utf-8"))["g"]
    by_comp = {v[2]: k for k, v in g.items() if v[2]}
    used = {}
    comps = {}
    for p in sorted((ROOT / "data" / "works").glob("*.json")):
        s = p.read_text(encoding="utf-8")
        for m in set(RE_COMP.findall(s)):
            if re.search(r"[-+*/@]", m) and m in by_comp:
                comps[m] = by_comp[m]
                used[by_comp[m]] = g[by_comp[m]]
        for cb in set(RE_CB.findall(s)):
            if cb in g:
                used[cb] = g[cb]
        for ch in set(RE_PUA.findall(s)):
            cb = "CB%05d" % (ord(ch) - 0xF0000)
            if cb in g:
                used[cb] = g[cb]
    doc = {"_note": "CB번호: [유니코드 글자, 정규형, 조합식]. c: 조합식 → CB번호. tools/gaiji_map.py 가 만든다.",
           "g": dict(sorted(used.items())), "c": dict(sorted(comps.items()))}
    OUT.write_text(json.dumps(doc, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    return len(used)


if __name__ == "__main__":
    n = build()
    print(f"외자 대응표 {n}자 → data/gaiji.json")
