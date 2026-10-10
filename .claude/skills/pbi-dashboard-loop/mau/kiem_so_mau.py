# -*- coding: utf-8 -*-
"""Cong chan so lieu: so tung measure DAX voi `moc.json` (so tinh doc lap, cung moc cat).

Chep thanh `kiem_so.py` canh loop.py. Desktop phai dang mo <du_an>.pbip da refresh xong.

    python kiem_so.py
Ma thoat:  0 khop  ·  2 lech  ·  3 khong hoi duoc Desktop

Dung sai:
  dem  -> khop tuyet doi (BLANK chap nhan khi ky vong 0)
  tien -> lech <= 0,005
  so   -> lech tuyet doi <= 1e-6; BLANK KHAC 0 (o chua co du lieu phai BLANK, o co du lieu
          ma bang khong phai 0)

Gom nhieu phep vao MOT cau EVALUATE UNION(ROW(...)) de moi lan goi pbi (vai giay khoi
dong) tinh ca lo. Goi DAX qua `--file` vi DAX nhieu dong khong truyen duoc bang -e.

Hai bai hoc da tra gia (Finance, 09/2026):
  - Cong KHONG duoc tu co lai: phep nao khong chay duoc phai tinh la LECH, khong bo qua.
  - Cong chi goi measure theo TEN nen khong bat duoc visual bind hong -- van phai xuat PDF
    va nhin/so diem anh.
Nguon goc: Game/PowerBI/JellyTrail/kiem_so.py (2026-09-29), tong quat hoa.
"""

import json
import os
import subprocess
import sys
import time

for _luong in (sys.stdout, sys.stderr):
    try:
        _luong.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
MOC = os.path.join(HERE, "moc.json")
BAO_CAO = os.path.join(HERE, "review", "kiem_so.txt")
ENV = dict(os.environ, PYTHONIOENCODING="utf-8")
LO = 200


def bieu_thuc(p):
    m = p.get("bieu_thuc") or "[%s]" % p["measure"]
    dk = p.get("loc") or []
    return "CALCULATE(%s, %s)" % (m, ", ".join(dk)) if dk else m


def ten(p):
    return p.get("measure") or p.get("bieu_thuc", "")[:34]


def chay_dax(cau):
    tam = os.path.join(HERE, "review", "_kiem_so.dax")
    os.makedirs(os.path.dirname(tam), exist_ok=True)
    open(tam, "w", encoding="utf-8", newline="\n").write(cau)
    r = subprocess.run(["pbi", "--json", "dax", "execute", "--file", tam], capture_output=True,
                       text=True, encoding="utf-8", errors="replace", env=ENV, timeout=900)
    ra = r.stdout or ""
    i = ra.find("{")
    if r.returncode != 0 or i < 0:
        print("Không chạy được DAX:\n" + (ra + (r.stderr or ""))[-2500:])
        return None
    try:
        return json.loads(ra[i:])["rows"]
    except Exception:
        print("Không đọc được kết quả:\n" + ra[:2000])
        return None


def so(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def khop(ta, mong, loai):
    if mong is None:
        return ta is None
    if ta is None:
        return loai in ("dem", "tien") and abs(mong) < 1e-9
    if loai == "dem":
        return abs(ta - mong) < 1e-9
    if loai == "tien":
        return abs(ta - mong) <= 0.005
    return abs(ta - mong) <= 1e-6


def main():
    if not os.path.exists(MOC):
        print("Chưa có moc.json (số kỳ vọng tính độc lập bằng SQL/pandas).")
        sys.exit(3)
    moc = json.load(open(MOC, encoding="utf-8"))
    phep = moc["phep"]
    t0 = time.time()

    ket_qua = {}
    for dau in range(0, len(phep), LO):
        lo = phep[dau:dau + LO]
        dong = ['    ROW("i", %d, "gt", %s)' % (dau + j, bieu_thuc(p)) for j, p in enumerate(lo)]
        rows = chay_dax("EVALUATE\nUNION(\n" + ",\n".join(dong) + "\n)")
        if rows is None:
            sys.exit(3)
        for r in rows:
            ket_qua[int(so(r.get("[i]")))] = so(r.get("[gt]"))

    lech, theo_nhom = [], {}
    for i, p in enumerate(phep):
        ta = ket_qua.get(i, "THIEU")
        ok = ta != "THIEU" and khop(ta, p["mong"], p["loai"])
        tong, sai = theo_nhom.get(p["nhom"], (0, 0))
        theo_nhom[p["nhom"]] = (tong + 1, sai + (0 if ok else 1))
        if not ok:
            lech.append((p, ta))

    dong = ["Đối chiếu %d phép (moc.json kéo lúc %s, mốc cắt %s) trong %.0fs"
            % (len(phep), moc.get("keo_luc_utc", "?"), moc.get("moc_cat", "?"), time.time() - t0)]
    for nhom in sorted(theo_nhom):
        tong, sai = theo_nhom[nhom]
        dong.append("   %-22s %4d phép  %s" % (nhom, tong, "OK" if not sai else "LỆCH %d" % sai))
    for p, ta in lech:
        dong.append("   x %-22s %-34s %-44s dax=%-22s moc=%s"
                    % (p["nhom"], ten(p), json.dumps(p.get("loc", []), ensure_ascii=False)[:44],
                       ta, p["mong"]))
    os.makedirs(os.path.dirname(BAO_CAO), exist_ok=True)
    open(BAO_CAO, "w", encoding="utf-8").write("\n".join(dong) + "\n")
    print("\n".join(dong[:1 + len(theo_nhom) + 40]))
    if len(lech) > 40:
        print("   ... còn %d phép lệch nữa, xem review/kiem_so.txt" % (len(lech) - 40))
    if lech:
        print("\nLỆCH %d / %d phép." % (len(lech), len(phep)))
        sys.exit(2)
    print("KHỚP TẤT CẢ %d phép." % len(phep))


if __name__ == "__main__":
    main()
