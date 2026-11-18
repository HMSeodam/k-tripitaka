#!/usr/bin/env python3
"""검색 색인 만들기 → data/search/ (저장소에는 넣지 않고 배포할 때만 만든다)

첫 검색에서 전체 문헌(수백 MB)을 내려받지 않도록, 글자 조각(1자·2자)마다
그 조각이 든 단위 목록을 미리 적어 둔다. 브라우저는 검색어의 조각이 든
색인 파일 몇 개와, 결과 화면에 보일 단위의 본문 조각만 받는다.

  data/search/meta.json            문헌별 시작 번호·단위 수·본문 조각 경계
  data/search/{cn|ko}{1|2}/{k}.bin 글자 조각 → 단위 번호 목록 (gzip)
  data/search/t/{cn|ko}/{문헌}/{c}.bin  단위 본문 묶음 (gzip, JSON)

정규화(위치표지·교감표지·구두점 빼기, 이체자 접기)는 assets/js/search.worker.js 와
똑같아야 한다. 이체자 표(VAR)는 그 파일에서 그대로 읽어 온다.
"""
import gzip
import json
import re
import shutil
import struct
import sys
import time
import unicodedata
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUT = DATA / "search"
WORKER = ROOT / "assets" / "js" / "search.worker.js"

SHARDS = {"cn1": 128, "cn2": 2048, "ko1": 128, "ko2": 1024}
CHUNK_CHARS = 16000          # 본문 묶음 하나에 담을 원문+번역 글자 수(대략)

MARKER = re.compile(r"\[\d{3,4}[abc]\d{2}\]")
APP = re.compile(r"\[(?:\d{1,3}|＊|\*)\]")
PUNCT = re.compile(r"[\s。，、．・？！：；「」『』（）()〔〕【】\[\]“”‘’·…—　]")


def load_var():
    src = WORKER.read_text(encoding="utf-8")
    body = src[src.index("const VAR = {"):]
    body = body[:body.index("};")]
    return dict(re.findall(r"'(.)'\s*:\s*'(.)'", body))


VAR = load_var()


def norm_cn(s):
    t = PUNCT.sub("", APP.sub("", MARKER.sub("", unicodedata.normalize("NFKC", s))))
    return "".join(VAR.get(c, c) for c in t)


def norm_ko(s):
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", s).lower())


def fnv(key):
    h = 0x811C9DC5
    for b in key.encode("utf-8"):
        h ^= b
        h = (h * 0x01000193) & 0xFFFFFFFF
    return h


def varint(n, out):
    while n >= 0x80:
        out.append((n & 0x7F) | 0x80)
        n >>= 7
    out.append(n)


def write_gz(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "wb") as f:
        f.write(gzip.compress(data, compresslevel=9, mtime=0))


def main():
    t0 = time.time()
    man = json.loads((DATA / "manifest.json").read_text(encoding="utf-8"))
    ids = [w["id"] for w in man["works"]]
    if OUT.exists():
        shutil.rmtree(OUT)
    post = {k: defaultdict(list) for k in SHARDS}
    works = []
    g = 0
    for wid in ids:
        doc = json.loads((DATA / "works" / f"{wid}.json").read_text(encoding="utf-8"))
        units = doc["units"]
        base = g
        # 본문 묶음 나누기
        starts, acc = [0], 0
        cn_chunk, ko_chunk = [], []

        def flush(c):
            write_gz(OUT / "t" / "cn" / wid / f"{c}.bin",
                     json.dumps(cn_chunk, ensure_ascii=False, separators=(",", ":")).encode())
            write_gz(OUT / "t" / "ko" / wid / f"{c}.bin",
                     json.dumps(ko_chunk, ensure_ascii=False, separators=(",", ":")).encode())

        for k, u in enumerate(units):
            cn_raw = "\n".join(u["cn"])
            ko_raw = "\n".join(u.get("ko") or [])
            if acc and acc + len(cn_raw) + len(ko_raw) > CHUNK_CHARS:
                flush(len(starts) - 1)
                cn_chunk, ko_chunk = [], []
                starts.append(k)
                acc = 0
            cn_chunk.append([u.get("i"), u.get("m"), cn_raw])
            ko_chunk.append(ko_raw)
            acc += len(cn_raw) + len(ko_raw)
            for f, t in (("cn", norm_cn(cn_raw)), ("ko", norm_ko(ko_raw))):
                if not t:
                    continue
                for ch in set(t):
                    post[f + "1"][ch].append(g)
                for bg in {t[p:p + 2] for p in range(len(t) - 1)}:
                    post[f + "2"][bg].append(g)
            g += 1
        if units:
            flush(len(starts) - 1)
        works.append([wid, base, len(units), starts])
        print(f"  {wid:24s} {len(units):6d}단위", flush=True)

    sizes = {}
    for name, n_sh in SHARDS.items():
        shards = defaultdict(list)
        for key, lst in post[name].items():
            shards[fnv(key) % n_sh].append(key)
        total = 0
        for s in range(n_sh):
            keys = sorted(shards.get(s, []))
            head, blob = bytearray(), bytearray()
            head += struct.pack("<I", len(keys))
            for key in keys:
                lst = post[name][key]
                start = len(blob)
                last = -1
                for x in lst:                     # 단위 번호는 이미 오름차순
                    varint(x - last, blob)
                    last = x
                kb = key.encode("utf-8")
                head += struct.pack("<B", len(kb)) + kb
                head += struct.pack("<III", len(lst), start, len(blob) - start)
            data = bytes(head + blob)
            write_gz(OUT / name / f"{s}.bin", data)
            total += (OUT / name / f"{s}.bin").stat().st_size
        sizes[name] = total
        post[name] = None

    meta = {
        "v": 1,
        "built": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "units": g,
        "shards": SHARDS,
        "works": works,
    }
    (OUT / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, separators=(",", ":")),
                                   encoding="utf-8")
    tsize = sum(p.stat().st_size for p in (OUT / "t").rglob("*.bin"))
    for f in ("cn", "ko"):
        ps = list((OUT / "t" / f).rglob("*.bin"))
        meta["avg_" + f] = round(sum(p.stat().st_size for p in ps) / max(1, len(ps)))
    (OUT / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, separators=(",", ":")),
                                   encoding="utf-8")
    print(f"검색 색인: {len(works)}종 {g}단위 · 색인 "
          + " ".join(f"{k} {v / 1e6:.1f}MB" for k, v in sizes.items())
          + f" · 본문 묶음 {tsize / 1e6:.1f}MB · {time.time() - t0:.0f}초")


if __name__ == "__main__":
    sys.exit(main())
