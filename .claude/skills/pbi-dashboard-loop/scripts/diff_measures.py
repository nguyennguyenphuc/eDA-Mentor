#!/usr/bin/env python3
"""So THAN DAX cua tung measure giua hai model.

Vi sao phep kiem nay manh hon 608 phep kiem so lieu: neu DAX cua moi measure
GIONG HET va du lieu nguon giong nhau, thi khong the ra so khac. Do la chung
minh, khong phai phep thu mau. Chay duoc ma khong can Power BI Desktop.

Dung sau moi dot ve sinh model: ve sinh dung nghia la KHONG mot bieu thuc nao
doi, chi thuoc tinh sieu du lieu doi.

  python <skill>/scripts/diff_measures.py MoHinh_cu MoHinh_moi
"""
from __future__ import annotations

import os
import argparse
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
PBI_ROOT = Path(os.environ.get("PBI_ROOT") or Path.cwd())  # thu muc chua <Model>.SemanticModel


def load(model: str) -> dict[str, str]:
    """Doc _Measures.tmdl -> {ten measure: than DAX da chuan hoa khoang trang}."""
    p = PBI_ROOT / f"{model}.SemanticModel" / "definition" / "tables" / "_Measures.tmdl"
    if not p.exists():
        sys.exit(f"khong thay {p}")
    out: dict[str, str] = {}
    name: str | None = None
    buf: list[str] = []
    for raw in p.read_text(encoding="utf-8").splitlines():
        m = re.match(r"\tmeasure\s+'?([^'=]+?)'?\s*=(.*)$", raw)
        if m:
            if name:
                out[name] = norm(buf)
            name = m.group(1).strip()
            buf = [m.group(2)]
            continue
        if name is None:
            continue
        # Ket thuc than DAX khi gap thuoc tinh cua measure hoac dong trong.
        if re.match(r"\t\t(formatString|displayFolder|lineageTag|isHidden|description):?", raw):
            out[name] = norm(buf)
            name = None
            buf = []
            continue
        if raw.startswith("\tpartition") or raw.startswith("\tcolumn"):
            out[name] = norm(buf)
            name = None
            buf = []
            continue
        buf.append(raw)
    if name:
        out[name] = norm(buf)
    return out


def norm(lines: list[str]) -> str:
    """Gop ve mot dong, ep khoang trang ve mot dau cach.

    Chi so SEMANTIC cua bieu thuc, khong so cach thut le -- doi thut le khong
    doi ket qua, ma bo sinh co the doi cach thut le mot cach vo hai.
    """
    s = " ".join(lines)
    s = s.replace("```", " ")
    return re.sub(r"\s+", " ", s).strip()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("goc", help="model goc, vd MoHinh_cu")
    ap.add_argument("moi", help="model moi, vd MoHinh_moi")
    ap.add_argument("--bo-qua", action="append", default=[],
                    help="ten measure duoc phep khac (lap lai duoc)")
    a = ap.parse_args()

    A, B = load(a.goc), load(a.moi)
    bo_qua = set(a.bo_qua)

    chi_goc = sorted(set(A) - set(B) - bo_qua)
    chi_moi = sorted(set(B) - set(A) - bo_qua)
    khac = sorted(n for n in set(A) & set(B) if A[n] != B[n] and n not in bo_qua)

    print(f"{a.goc}: {len(A)} measure   |   {a.moi}: {len(B)} measure")
    if bo_qua:
        print(f"bo qua co chu y: {sorted(bo_qua)}")
    print()

    for n in chi_goc:
        print(f"  CHI CO O {a.goc}: {n}")
    for n in chi_moi:
        print(f"  CHI CO O {a.moi}: {n}")
    for n in khac:
        print(f"  THAN DAX KHAC: {n}")
        print(f"     goc: {A[n][:160]}")
        print(f"     moi: {B[n][:160]}")

    xau = len(chi_goc) + len(chi_moi) + len(khac)
    print()
    if xau == 0:
        print(f"QUA: {len(set(A) & set(B) - bo_qua)} measure GIONG HET tung ky tu.")
        print("     DAX giong nhau tren cung du lieu thi khong the ra so khac.")
        return 0
    print(f"DUNG: {xau} cho lech.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
