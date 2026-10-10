#!/usr/bin/env python3
"""So hai ban render theo TUNG DIEM ANH va gom cho khac thanh vung.

Vi sao khong nhin bang mat: da co tien le trong du an nay -- doi cach ve duong
o vong 10/14, toi *tuong* duong da bo tron hon, dem diem anh cho thay GIONG HET.
Mat nguoi khong dang tin cho loai thay doi nay.

Vi sao khong chi dem tong: mot con so "3% diem anh khac" khong tra loi duoc cau
hoi quan trong -- khac O DAU. Nen cong cu gom cac diem khac lien nhau thanh
vung chu nhat va in toa do tung vung, de doi chieu voi danh sach cho DUOC PHEP
khac.

Dung cho v2 vs v3: hai ban PHAI khac, va phai khac DUNG o 5 cho da du doan.
Giong het nghia la v3 chua thuc su doi nguon cong thuc; khac o cho khac nghia
la co loi.

  python so_diem_anh.py review/render_r01 review/render_r02
  python so_diem_anh.py A B --nguong 16 --toi-thieu 40
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from PIL import Image, ImageChops


def vung_khac(a: Image.Image, b: Image.Image, nguong: int, o: int) -> tuple[int, list]:
    """Tra ve (so diem khac, danh sach hop bao cua tung cum).

    Gom bang luoi o vuong `o` diem anh thay vi tim cum lien thong tung diem:
    nhanh hon nhieu, va o day ta chi can biet VUNG NAO doi chu khong can duong
    bien chinh xac.
    """
    diff = ImageChops.difference(a.convert("L"), b.convert("L"))
    px = diff.load()
    w, h = diff.size
    o_khac: dict[tuple[int, int], int] = {}
    tong = 0
    for y in range(h):
        for x in range(w):
            if px[x, y] > nguong:
                tong += 1
                k = (x // o, y // o)
                o_khac[k] = o_khac.get(k, 0) + 1
    # Gom o lien nhau (8 huong) thanh cum
    chua_xet = set(o_khac)
    cum = []
    while chua_xet:
        goc = chua_xet.pop()
        hang_doi, thanh_vien = [goc], [goc]
        while hang_doi:
            cx, cy = hang_doi.pop()
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    k = (cx + dx, cy + dy)
                    if k in chua_xet:
                        chua_xet.discard(k)
                        hang_doi.append(k)
                        thanh_vien.append(k)
        xs = [k[0] for k in thanh_vien]
        ys = [k[1] for k in thanh_vien]
        n = sum(o_khac[k] for k in thanh_vien)
        cum.append((n, xs[0] and min(xs) * o or min(xs) * o, min(ys) * o,
                    (max(xs) + 1) * o, (max(ys) + 1) * o))
    cum.sort(reverse=True)
    return tong, cum


def nap_du_doan(p: Path | None):
    """Doc ban du doan. Tra None neu khong co, tuc cong chi in chu khong phan quyet."""
    if p is None:
        return None
    if not p.exists():
        sys.exit(f"Khong thay ban du doan {p}")
    return json.loads(p.read_text(encoding="utf-8"))


def trong_hop(cum, hop) -> bool:
    _, x0, y0, x1, y1 = cum
    a, b, c, d = hop
    return x0 >= a and y0 >= b and x1 <= c and y1 <= d


def phan_quyet(du_doan: dict, thay: dict, tong: int, toi_thieu: int) -> int:
    """Doi MOI cum diem khac phai roi vao mot vung da du doan, va moi vung phai co cum.

    Truoc day ham main() ket thuc bang `return 0` vo dieu kien: cong in ra 34
    nghin diem khac cung xanh y het khi hai ban giong het nhau. Mot cong khong
    the do la mot trinh in bao cao deo nhan cong.

    Hai phia deu phai dung:
      - cum roi ngoai moi vung  -> co thay doi khong ai du doan, tuc co loi;
      - vung khong co cum nao   -> thay doi da du doan KHONG xay ra, tuc ban
        moi chua thuc su khac ban cu o cho le ra phai khac.
    """
    print("\n=== PHAN QUYET theo ban du doan ===\n")
    loi: list[str] = []
    if tong == 0:
        loi.append("hai ban giong het nhau: ban moi chua doi duoc gi")
    for ten_trang, vung in du_doan.items():
        if ten_trang.startswith("_"):
            continue
        cum = [c for c in thay.get(ten_trang, []) if c[0] >= toi_thieu]
        da_dung = [False] * len(vung)
        for c in cum:
            khop = [i for i, v in enumerate(vung) if trong_hop(c, v["hop"])]
            if not khop:
                loi.append(f"{ten_trang}: cum {c[0]} diem o x {c[1]}-{c[3]} "
                           f"y {c[2]}-{c[4]} KHONG nam trong vung du doan nao")
            for i in khop:
                da_dung[i] = True
        for i, v in enumerate(vung):
            if not da_dung[i]:
                loi.append(f"{ten_trang}: vung du doan “{v['ten']}” khong co cum nao")
        print(f"  {ten_trang}: {len(cum)} cum, {sum(da_dung)}/{len(vung)} vung du doan co cum")
    print()
    if loi:
        for x in loi:
            print(f"  LOI  {x}")
        print(f"\n{len(loi)} cho khong khop ban du doan.")
        return 1
    print("cong xanh: moi thay doi deu nam trong vung da du doan, va moi vung deu co.")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("cu")
    ap.add_argument("moi")
    ap.add_argument("--du-doan", type=Path, default=None,
                    help="tep JSON khai cac vung DUOC PHEP khac; khong co thi "
                         "cong chi in chu khong phan quyet")
    ap.add_argument("--nguong", type=int, default=12,
                    help="chenh do xam toi thieu de tinh la khac (mac dinh 12)")
    ap.add_argument("--toi-thieu", type=int, default=30,
                    help="cum duoi so diem nay thi khong in (mac dinh 30)")
    ap.add_argument("--o", type=int, default=16, help="canh o gom cum")
    args = ap.parse_args()

    d_cu, d_moi = Path(args.cu), Path(args.moi)
    trang = sorted(d_cu.glob("page_*.png"))
    if not trang:
        sys.exit(f"Khong thay page_*.png trong {d_cu}")

    du_doan = nap_du_doan(args.du_doan)
    thay: dict[str, list] = {}
    print(f"so {d_cu.name} vs {d_moi.name} · nguong {args.nguong}\n")
    tong_tat_ca = 0
    for p in trang:
        q = d_moi / p.name
        if not q.exists():
            print(f"  {p.name}: THIEU o {d_moi}")
            continue
        a, b = Image.open(p), Image.open(q)
        if a.size != b.size:
            print(f"  {p.name}: KICH THUOC KHAC {a.size} vs {b.size}")
            continue
        n, cum = vung_khac(a, b, args.nguong, args.o)
        thay[p.name] = cum
        tong_tat_ca += n
        pct = 100 * n / (a.size[0] * a.size[1])
        lon = [c for c in cum if c[0] >= args.toi_thieu]
        print(f"  {p.name}: {n:>7} diem khac ({pct:.2f}%), {len(lon)} vung")
        for cnt, x0, y0, x1, y1 in lon[:10]:
            print(f"      {cnt:>6} diem  x {x0:>4}-{x1:<4} y {y0:>4}-{y1:<4}")
    print(f"\ntong {tong_tat_ca} diem khac tren {len(trang)} trang")

    if du_doan is None:
        print("\n(khong co ban du doan: cong nay chi IN chu khong phan quyet)")
        return 0
    return phan_quyet(du_doan, thay, tong_tat_ca, args.toi_thieu)


if __name__ == "__main__":
    raise SystemExit(main())
