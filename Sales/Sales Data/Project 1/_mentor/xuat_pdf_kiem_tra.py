# -*- coding: utf-8 -*-
"""Kiểm dashboard Excel tham chiếu bằng LibreOffice (máy sinh tài liệu không có Excel).

1. Mở BẢN SAO _mentor/Nexa_Sales_P1_Doi_chieu_cong_thuc.xlsx (bản công thức thường) bằng LibreOffice (UNO), tính lại toàn bộ công thức.
2. Đổi ô chọn Tháng / Kênh / Quốc gia qua nhiều tổ hợp, đọc số ở sheet Calc và so với pandas (aggregate.load()) —
   cùng nguồn với bản HTML. Lệch > 0,01 USD là lỗi.
3. Đặt lại mặc định, xuất 4 sheet dashboard ra kiem_tra/dashboard_xlsx.pdf (mỗi sheet một trang ngang A4).
4. Ghép từng trang PDF cạnh ảnh tab tương ứng của bản HTML (assets/dash_sN.png) → kiem_tra/so_sanh_sN.png.
Cần: LibreOffice (soffice + python3-uno), pdftoppm (poppler), Pillow.  Chạy: python xuat_pdf_kiem_tra.py
PDF do LibreOffice vẽ nên font và biểu đồ không giống Excel 100 %; dùng để kiểm số và bố cục.
"""
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import pandas as pd
from PIL import Image

import aggregate as A

HERE = Path(__file__).resolve().parent
SRC = A.P1 / "_mentor" / "Nexa_Sales_P1_Doi_chieu_cong_thuc.xlsx"
OUT = HERE / "kiem_tra"
OUT.mkdir(exist_ok=True)


# ------------------------------------------------------------------ số đúng tính bằng pandas
D = A.load()
ms = D["ms"].merge(D["st"][["store_key", "country", "channel_code"]], on="store_key")
ms["country_vi"] = ms.country.map(D["COUNTRY_VI"])
THR = A.REF["threshold"]


def dung(mid, ch, ct):
    x = ms[ms.month_id == mid]
    if ch != "*":
        x = x[x.channel_code == ch]
    if ct != "*":
        x = x[x.country_vi == ct]
    nr = x.ns.sum() - x.ret.sum()
    d = x.fin - (x.ns - x.ret)
    flag = (x.fin != 0) & (d.abs() > THR["abs_usd"]) & ((d / x.fin.abs().where(x.fin != 0)).abs() > THR["pct"])
    return {"nr": nr, "orders": x.orders.sum(), "fin": x.fin.sum(), "delta": x.fin.sum() - nr,
            "dadj": (x.fin - (x.ns_post - x.ret)).sum(), "nflag": int(flag.sum()), "ret": x.ret.sum(),
            "gross": x.gross.sum()}


# ------------------------------------------------------------------ LibreOffice qua UNO
def mo_libreoffice():
    import uno  # noqa: F401  (python3-uno)
    prof = Path(tempfile.mkdtemp(prefix="lo_prof_"))
    port = 2002 + os.getpid() % 500
    p = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", "--nologo",
                          f"-env:UserInstallation={prof.as_uri()}",
                          f"--accept=socket,host=127.0.0.1,port={port};urp;"],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    import uno
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(60):
        try:
            ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(0.5)
    else:
        p.kill()
        sys.exit("Không kết nối được LibreOffice")
    desk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return p, desk, prof


def prop(name, value):
    import uno
    from com.sun.star.beans import PropertyValue
    pv = PropertyValue()
    pv.Name, pv.Value = name, value
    return pv


def main():
    import uno
    tmp = Path(tempfile.mkdtemp()) / "dash.xlsx"
    shutil.copy(SRC, tmp)
    proc, desk, prof = mo_libreoffice()
    try:
        doc = desk.loadComponentFromURL(tmp.as_uri(), "_blank", 0, (prop("Hidden", True),))
        sh = doc.Sheets
        tq, calc = sh.getByName("Tong_quan"), sh.getByName("Calc")
        # vị trí các KPI trong Calc: cột A là tên, cột B là giá trị tháng chọn
        hang = {}
        for r in range(14, 60):
            k = calc.getCellByPosition(0, r).getString()
            if k:
                hang[k] = r
        TH = [("12/2020", "Tất cả", "Tất cả"), ("12/2020", "Cửa hàng (POS)", "Tất cả"), ("12/2020", "Website/app", "Tất cả"),
              ("06/2020", "Tất cả", "Mỹ"), ("02/2020", "Cửa hàng (POS)", "Đức"), ("01/2020", "Tất cả", "Tất cả"),
              ("10/2019", "Tất cả", "Pháp")]
        CH = {"Tất cả": "*", "Cửa hàng (POS)": "POS", "Website/app": "ONLINE"}
        loi, dong = 0, []
        for thang, kenh, qg in TH:
            tq.getCellRangeByName("C5").setString(thang)
            tq.getCellRangeByName("F5").setString(kenh)
            tq.getCellRangeByName("I5").setString(qg)
            doc.calculateAll()
            mid = int(thang[3:]) * 100 + int(thang[:2])
            ok = dung(mid, CH[kenh], "*" if qg == "Tất cả" else qg)
            for k, v in ok.items():
                got = calc.getCellByPosition(1, hang[k]).getValue()
                if abs(got - v) > 0.01:
                    loi += 1
                    dong.append(f"  LỆCH {thang} · {kenh} · {qg} · {k}: Excel {got:,.2f} ≠ pandas {v:,.2f}")
            dong.append(f"  {thang} · {kenh:15} · {qg:7} · Net Revenue {ok['nr']:>12,.2f} · Δ {ok['delta']:>11,.2f} · "
                        f"{ok['nflag']} ô vượt ngưỡng")
        # lỗi công thức trên 4 sheet dashboard
        for name in ("Tong_quan", "Doanh_thu", "Doi_soat", "Hoan_tra"):
            ws = sh.getByName(name)
            for r in range(1, 61):
                for c in range(1, 17):
                    cell = ws.getCellByPosition(c, r)
                    if cell.getType().value == "FORMULA" and cell.getError() != 0:
                        loi += 1
                        dong.append(f"  LỖI CÔNG THỨC {name}!{chr(65 + c)}{r + 1}: mã {cell.getError()}")
        # đặt lại mặc định rồi xuất PDF 4 trang đầu (4 sheet dashboard, mỗi sheet vừa 1 trang)
        tq.getCellRangeByName("C5").setString("12/2020")
        tq.getCellRangeByName("F5").setString("Tất cả")
        tq.getCellRangeByName("I5").setString("Tất cả")
        doc.calculateAll()
        pdf = OUT / "dashboard_xlsx.pdf"
        doc.storeToURL(pdf.as_uri(), (prop("FilterName", "calc_pdf_Export"),
                                      prop("FilterData", uno.Any("[]com.sun.star.beans.PropertyValue",
                                                                 tuple([prop("PageRange", "1-4")])))))
        doc.close(True)
    finally:
        try:
            desk.terminate()
        except Exception:
            pass
        proc.wait(timeout=30)
        shutil.rmtree(prof, ignore_errors=True)
    print("\n".join(dong))
    # ghép trang PDF cạnh ảnh HTML
    subprocess.run(["pdftoppm", "-r", "80", "-png", str(pdf), str(OUT / "xlsx")], check=True)
    for i in range(1, 5):
        a = Image.open(sorted(OUT.glob("xlsx-*.png"))[i - 1])
        b = Image.open(HERE / "assets" / f"dash_s{i}.png")
        h = 900
        a = a.resize((int(a.width * h / a.height), h))
        b = b.resize((int(b.width * h / b.height), h))
        c = Image.new("RGB", (a.width + b.width + 20, h), "white")
        c.paste(a, (0, 0))
        c.paste(b, (a.width + 20, 0))
        c.save(OUT / f"so_sanh_s{i}.png")
    for p in OUT.glob("xlsx-*.png"):
        p.unlink()
    print(("OK" if not loi else f"{loi} LỖI") + f" · PDF: {pdf.relative_to(A.P1)} · ảnh so sánh: kiem_tra/so_sanh_s1..4.png")
    sys.exit(1 if loi else 0)


if __name__ == "__main__":
    main()
