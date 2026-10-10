# -*- coding: utf-8 -*-
"""
Hỏi thẳng engine: measure nào đang HỎNG, và hỏng vì sao.

Vì sao cần: một measure sai cú pháp không bật dialog, không làm refresh fail, không hiện
trong `pbi report validate`. Nó chỉ làm cả MdxScript của model hỏng, rồi MỌI card trên
MỌI trang báo "See details / Fix this", và mọi query DAX trả về đúng một câu vô dụng:

    MdxScript(Model) (3, 25) Failed to resolve name 'SYNTAXERROR'.

Câu đó KHÔNG nói measure nào. DMV `$SYSTEM.TMSCHEMA_MEASURES` thì có cột `ErrorMessage`
ghi rõ từng measure hỏng kèm nguyên văn biểu thức — đó là đường duy nhất lần ra.

Đã bắt được bằng script này (2026-08-23): `KPI` là TỪ KHOÁ DAX nên `MAX(Kpi[total_orders])`
báo "The syntax for 'Kpi' is incorrect" — 15/41 measure chết vì thiếu nháy đơn quanh tên bảng.

Chạy:  python pbi_measures_check.py   (cần `pbi` trên PATH: pbi-cli thật hoặc shim build_pbi_shim.ps1)
Mã thoát: 0 = mọi measure sạch · 2 = có measure hỏng · 3 = không kết nối được Desktop.
"""

import json
import os
import subprocess
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

DMV = ("SELECT [Name],[TableID],[Expression],[ErrorMessage] "
       "FROM $SYSTEM.TMSCHEMA_MEASURES")


def chay(args):
    # PYTHONIOENCODING: pbi-cli in ra console cp1252 nên tên measure tiếng Việt làm nó
    # chết ngay trong lúc in — không phải lỗi kết nối. Ép utf-8 cho tiến trình con.
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    return subprocess.run(args, capture_output=True, text=True, encoding="utf-8",
                          errors="replace", shell=False, env=env)


def main():
    r = chay(["pbi", "--json", "dax", "execute", DMV])
    out = r.stdout or ""
    i = out.find("{")
    if i < 0:
        print("KHÔNG kết nối được Desktop / DMV lỗi:")
        print((out + (r.stderr or ""))[:800])
        return 3
    try:
        rows = json.loads(out[i:]).get("rows", [])
    except Exception as e:
        print("Không đọc được JSON: %s" % e)
        print(out[:800])
        return 3

    hong = [x for x in rows if (x.get("ErrorMessage") or "").strip()]
    print("MEASURE_CHECK: %d measure, %d hỏng." % (len(rows), len(hong)))
    for x in hong:
        msg = " ".join((x.get("ErrorMessage") or "").split())
        print("   x %-28s %s" % (x.get("Name"), msg[:220]))
    if hong:
        print("\nSửa trong build_model.py của dự án (MEASURES), chạy lại build_model.py,")
        print("rồi chạy lại loop.py — KHÔNG sửa tay TMDL, sẽ mất khi sinh lại.")
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
