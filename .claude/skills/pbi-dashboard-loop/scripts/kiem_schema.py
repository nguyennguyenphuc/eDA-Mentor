# -*- coding: utf-8 -*-
"""Kiem tung file PBIR theo dung `$schema` cua Microsoft.

`pbi report validate` chi kiem cau truc thu muc va cu phap JSON, no KHONG doi
chieu voi luoc do that. Lan truoc duong nay bat duoc 15 loi ma `report validate`
bao xanh, trong do co hai loi lam Power BI Desktop crash khi luu.

Chay:  python kiem_schema.py <TenDuAn>.Report   (hoac duong dan toi thu muc .Report)

Ma thoat:  0 sach  ·  2 co loi
"""

import json
import os
import sys
import urllib.request
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import jsonschema

HERE = Path(os.path.dirname(os.path.abspath(__file__)))
GOC = Path(sys.argv[1]) if len(sys.argv) > 1 else None
if GOC is not None and (GOC / "definition").is_dir():
    GOC = GOC / "definition"

_cache = {}


def luoc_do(url):
    if url not in _cache:
        with urllib.request.urlopen(url, timeout=30) as r:
            _cache[url] = json.loads(r.read().decode("utf-8"))
    return _cache[url]


def main():
    if GOC is None or not GOC.is_dir():
        sys.exit("Dung: python kiem_schema.py <TenDuAn>.Report")
    loi, n = [], 0
    for f in sorted(GOC.rglob("*.json")):
        j = json.load(open(f, encoding="utf-8"))
        url = j.get("$schema")
        if not url:
            continue
        n += 1
        try:
            jsonschema.validate(j, luoc_do(url))
        except jsonschema.ValidationError as e:
            loi.append((f.relative_to(GOC),
                        "/".join(str(x) for x in e.absolute_path), e.message[:120]))
        except Exception as e:
            loi.append((f.relative_to(GOC), "(không tải được lược đồ)", str(e)[:80]))

    print("%d file có $schema, %d lỗi." % (n, len(loi)))
    if n == 0:
        # Cổng không tìm thấy gì để kiểm thì ĐỎ, không phải xanh (cong-nghiem-thu.md).
        print("   x không có file PBIR nào mang $schema dưới %s" % GOC)
        sys.exit(2)
    for duong, cho, msg in loi[:30]:
        print("   x %-52s %s\n       %s" % (duong, cho, msg))
    if len(loi) > 30:
        print("   ... còn %d lỗi nữa" % (len(loi) - 30))
    sys.exit(2 if loi else 0)


if __name__ == "__main__":
    main()
