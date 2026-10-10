#!/usr/bin/env python3
"""Cham semantic model theo tap con quy tac Best Practice Analyzer.

Nguon quy tac: microsoft/Analysis-Services -> BestPracticeRules/BPARules.json
(bo quy tac chinh thuc ma Tabular Editor dung). O day cai lai bang Python vi
may nay khong cai Tabular Editor, va chi cai nhung quy tac doc duoc TU TMDL --
khong cai quy tac can do luong du lieu (cardinality, dung luong cot).

Doc thang TMDL nen KHONG can Power BI Desktop dang mo.

  python <skill>/scripts/bpa_check.py MoHinh
  python <skill>/scripts/bpa_check.py MoHinh_cu MoHinh_moi   # bang truoc/sau
"""
from __future__ import annotations

import os
import argparse
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
PBI_ROOT = Path(os.environ.get("PBI_ROOT") or Path.cwd())  # thu muc chua <Model>.SemanticModel

# Cot ket thuc bang cac hau to nay duoc coi la khoa/ky thuat, khong danh cho
# nguoi doc bao cao.
KHOA_HAU_TO = ("_id", "Order", "Code", "_path")


def doc(model: str) -> dict:
    d = PBI_ROOT / f"{model}.SemanticModel" / "definition"
    if not d.exists():
        sys.exit(f"khong thay {d}")
    m = {"tables": {}, "rel": (d / "relationships.tmdl").read_text(encoding="utf-8")
         if (d / "relationships.tmdl").exists() else "",
         "model": (d / "model.tmdl").read_text(encoding="utf-8")}
    for f in sorted((d / "tables").glob("*.tmdl")):
        m["tables"][f.stem] = f.read_text(encoding="utf-8")
    return m


def cot_cua(txt: str) -> list[dict]:
    """Tach tung khoi 'column' thanh dict thuoc tinh."""
    out = []
    cur = None
    for ln in txt.splitlines():
        mm = re.match(r"\tcolumn\s+'?([^'=\n]+?)'?(\s*=\s*(.*))?$", ln)
        if mm:
            if cur:
                out.append(cur)
            cur = {"ten": mm.group(1).strip(), "calc": bool(mm.group(2)),
                   "props": [], "raw": []}
            continue
        if cur is None:
            continue
        if ln.startswith("\tcolumn") or ln.startswith("\tpartition") \
                or ln.startswith("\thierarchy") or ln.startswith("\tmeasure"):
            out.append(cur)
            cur = None
            continue
        cur["raw"].append(ln.strip())
    if cur:
        out.append(cur)
    return out


def cham(m: dict) -> dict[str, tuple[int, str]]:
    """-> {ma quy tac: (so vi pham, ghi chu)}"""
    r: dict[str, tuple[int, str]] = {}
    tables = m["tables"]

    # --- HIDE_FOREIGN_KEYS + HIDE_FACT_TABLE_COLUMNS -----------------------
    lo = []
    for t, txt in tables.items():
        if t == "_Measures":
            continue
        for c in cot_cua(txt):
            if any(c["ten"].endswith(s) for s in KHOA_HAU_TO) \
                    and "isHidden" not in c["raw"]:
                lo.append(f"{t}[{c['ten']}]")
    r["HIDE_FOREIGN_KEYS"] = (len(lo), ", ".join(lo[:4]) + ("…" if len(lo) > 4 else ""))

    # --- MARK_PRIMARY_KEYS: phia "one" cua quan he nen co isKey ------------
    ben_one = set(re.findall(r"^\ttoColumn:\s*(\S+)\.(\S+)$", m["rel"], re.M))
    thieu = []
    for t, c in sorted(ben_one):
        txt = tables.get(t, "")
        for col in cot_cua(txt):
            if col["ten"] == c and "isKey" not in col["raw"]:
                thieu.append(f"{t}[{c}]")
    r["MARK_PRIMARY_KEYS"] = (len(thieu), ", ".join(thieu[:4]) + ("…" if len(thieu) > 4 else ""))

    # --- REDUCE_NUMBER_OF_CALCULATED_COLUMNS (nguong BPA: > 5) -------------
    n = sum(1 for t, txt in tables.items() for c in cot_cua(txt) if c["calc"])
    r["REDUCE_NUMBER_OF_CALCULATED_COLUMNS"] = (max(0, n - 5), f"{n} cot tinh (nguong 5)")

    # --- REDUCE_USAGE_OF_CALCULATED_TABLES ---------------------------------
    nct = sum(1 for txt in tables.values() if re.search(r"^\tpartition .* = calculated", txt, re.M))
    r["REDUCE_USAGE_OF_CALCULATED_TABLES"] = (nct, f"{nct}/{len(tables)} bang")

    # --- OBJECTS_WITH_NO_DESCRIPTION (bang) --------------------------------
    kho = [t for t, txt in tables.items() if not txt.lstrip().startswith("///")]
    r["OBJECTS_WITH_NO_DESCRIPTION"] = (len(kho), f"{len(kho)}/{len(tables)} bang thieu mo ta")

    # --- PROVIDE_FORMAT_STRING_FOR_MEASURES --------------------------------
    mt = tables.get("_Measures", "")
    khoi = [b for b in re.split(r"\n(?=\tmeasure )", mt) if b.lstrip().startswith("measure ")]
    # Quy tac that cua BPA chi soi measure TRA SO. Kieu tra ve khong doc duoc tu
    # TMDL, nen loai bang dau hieu: tra chuoi ("..." hoac SELECTEDVALUE cot chu
    # hoac noi chuoi bang &) va tra luan ly (ISFILTERED / ISBLANK / || / &&).
    # Khong loai thi cong cu bao 8 vi pham trong khi ca 8 deu la measure chu.
    def tra_so(b: str) -> bool:
        than = b.split("formatString:")[0]
        than = re.sub(r"\tmeasure\s+'?[^'=\n]+?'?\s*=", "", than, count=1).strip()
        if re.match(r'^"', than) or "&" in than:
            return False
        if re.search(r"\b(ISFILTERED|ISCROSSFILTERED)\b", than):
            return False
        if re.search(r"SELECTEDVALUE\s*\(\s*\w+\[(MonthKey|NhanKy)\]", than):
            return False
        return True

    thieu_fmt = [b for b in khoi if "formatString:" not in b and tra_so(b)]
    r["PROVIDE_FORMAT_STRING_FOR_MEASURES"] = (
        len(thieu_fmt),
        f"{len(thieu_fmt)}/{len(khoi)} measure tra so thieu formatString "
        f"(da loai measure tra chu/luan ly)")

    # --- DATECOLUMN_FORMATSTRING -------------------------------------------
    thieu_ngay = []
    for t, txt in tables.items():
        for c in cot_cua(txt):
            if any("dataType: dateTime" == p for p in c["raw"]) \
                    and not any(p.startswith("formatString:") for p in c["raw"]):
                thieu_ngay.append(f"{t}[{c['ten']}]")
    r["DATECOLUMN_FORMATSTRING"] = (len(thieu_ngay),
                                    ", ".join(thieu_ngay[:4]) + ("…" if len(thieu_ngay) > 4 else ""))

    # --- NUMERIC_COLUMN_SUMMARIZE_BY: cot khoa khong duoc tu cong ----------
    xau = []
    for t, txt in tables.items():
        for c in cot_cua(txt):
            if any(c["ten"].endswith(s) for s in KHOA_HAU_TO) \
                    and "summarizeBy: none" not in c["raw"]:
                xau.append(f"{t}[{c['ten']}]")
    r["NUMERIC_COLUMN_SUMMARIZE_BY"] = (len(xau), ", ".join(xau[:4]))

    # --- MODEL_SHOULD_HAVE_A_DATE_TABLE ------------------------------------
    co = any("dataCategory: Time" in txt for txt in tables.values())
    r["MODEL_SHOULD_HAVE_A_DATE_TABLE"] = (0 if co else 1,
                                           "co bang Ngay" if co else "KHONG co bang Ngay danh dau")

    # --- ENSURE_TABLES_HAVE_RELATIONSHIPS ----------------------------------
    trong_rel = set(re.findall(r"^\t(?:from|to)Column:\s*(\S+)\.", m["rel"], re.M))
    roi = [t for t in tables if t not in trong_rel and t != "_Measures"]
    r["ENSURE_TABLES_HAVE_RELATIONSHIPS"] = (len(roi), f"{len(roi)} bang roi: " + ", ".join(sorted(roi)[:5]) + "…")

    # --- PARTITION_NAME_SHOULD_MATCH_TABLE_NAME ----------------------------
    lech = []
    for t, txt in tables.items():
        mm = re.search(r"^\tpartition\s+'?([^'\s=]+)'?\s*=", txt, re.M)
        if mm and mm.group(1) != t:
            lech.append(f"{t} -> {mm.group(1)}")
    r["PARTITION_NAME_SHOULD_MATCH_TABLE_NAME"] = (len(lech), ", ".join(lech[:3]))

    return r


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("models", nargs="+", help="mot model, hoac hai model de so truoc/sau")
    a = ap.parse_args()
    ket = [(mn, cham(doc(mn))) for mn in a.models]
    ma = sorted(ket[0][1])

    w = max(len(x) for x in ma) + 2
    head = f"  {'QUY TAC BPA':<{w}}" + "".join(f"{mn[:18]:>20}" for mn, _ in ket)
    print(head)
    print("  " + "-" * (len(head) - 2))
    tong = [0] * len(ket)
    for code in ma:
        row = f"  {code:<{w}}"
        for i, (_, r) in enumerate(ket):
            n = r[code][0]
            tong[i] += n
            row += f"{('OK' if n == 0 else str(n)):>20}"
        print(row)
        # ghi chu lay tu model cuoi cung
        note = ket[-1][1][code][1]
        if note:
            print(f"  {'':<{w}}    {note}")
    print("  " + "-" * (len(head) - 2))
    print(f"  {'TONG VI PHAM':<{w}}" + "".join(f"{t:>20}" for t in tong))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
