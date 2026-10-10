# -*- coding: utf-8 -*-
"""Dựng Mô hình Dữ liệu trong Excel thật (Windows + Excel 365/2016+, qua COM) đúng các bước học viên làm:
Power Query (13 query trong hoc_vien/excel_build/m) → Data Model → quan hệ (guideline 3.1) → số đo DAX
(measures.dax + measures_nang_cao.dax). Dùng chung cho build_dashboard_excel_com.py.
"""
import re
import time
from pathlib import Path

import win32com.client as win32

P1 = Path(__file__).resolve().parent.parent
KIT = P1 / "_mentor" / "excel_build"                 # code M + DAX gốc, chỉ mentor (không giao học viên)
DATA_DIR = P1 / "hoc_vien"

# query chỉ tạo connection, không nạp Data Model (guideline 2.7)
ONLY_CONNECTION = {"SourceFolder", "finance_raw", "map_store_finance"}

# (bảng nhiều, cột nhiều, bảng một, cột một, active) — guideline 3.1
RELATIONSHIPS = [
    ("fact_sales_line", "delivery_date", "dim_date", "date", False),
    ("fact_sales_line", "order_date", "dim_date", "date", True),
    ("fact_return", "return_date", "dim_date", "date", True),
    ("fact_finance", "posting_date", "dim_date", "date", True),
    ("fact_sales_line", "store_key", "dim_store", "store_key", True),
    ("fact_return", "store_key", "dim_store", "store_key", True),
    ("fact_finance", "store_key", "dim_store", "store_key", True),
    ("fact_sales_line", "product_key", "dim_product", "product_key", True),
    ("fact_return", "product_key", "dim_product", "product_key", True),
    ("fact_sales_line", "promo_code", "dim_promotion", "promo_code", True),
    ("fact_sales_line", "order_status", "dim_order_status", "order_status", True),
    ("dim_store", "channel_code", "dim_channel", "channel_code", True),
]

# định dạng số đo (guideline 3.2): tiền 0 số lẻ, % 1–2 số lẻ
PCT_MEASURES = {"Return Rate", "Δ Finance %", "Tỷ trọng online", "Tỷ lệ chiết khấu", "Return Rate SPLY"}
INT_MEASURES = {"Orders", "Units", "Vượt ngưỡng", "Số cửa hàng vượt ngưỡng", "Cancelled Orders",
                "Pending Orders", "Orders SPLY"}


def read_queries():
    """[(tên, M)] theo thứ tự file; bỏ dòng chú thích // ở đầu file."""
    out = []
    for f in sorted((KIT / "m").glob("*.pq")):
        name = re.sub(r"^\d+_", "", f.stem)
        lines = f.read_text(encoding="utf-8").splitlines()
        while lines and lines[0].lstrip().startswith("//"):
            lines.pop(0)
        out.append((name, "\n".join(lines).strip()))
    return out


def read_measures():
    """[(tên, công thức)] từ measures.dax rồi measures_nang_cao.dax; bỏ chú thích //."""
    out = []
    for fn in ("measures.dax", "measures_nang_cao.dax"):
        cur = None
        for line in (KIT / fn).read_text(encoding="utf-8").splitlines():
            s = line.strip()
            if not s or s.startswith("//"):
                continue
            m = re.match(r"^(\S.*?)\s*:=\s*(.*)$", line)
            if m and not line[0].isspace():
                cur = [m.group(1).strip(), m.group(2)]
                out.append(cur)
            else:
                cur[1] += "\n" + line
    return [(n, f.strip()) for n, f in out]


def wait_ready(xl, timeout=300):
    """Chờ Excel rảnh (sau khi nạp Data Model Excel còn xử lý nền và từ chối lệnh COM: 'Call was rejected')."""
    t = time.time()
    while time.time() - t < timeout:
        try:
            if xl.Ready:
                xl.CalculateUntilAsyncQueriesDone()
                return
        except Exception:
            pass
        time.sleep(1)


def retry(fn, tries=60):
    """Gọi lại lệnh COM khi Excel bận (RPC_E_CALL_REJECTED = -2147418111)."""
    for k in range(tries):
        try:
            return fn()
        except Exception as e:
            if getattr(e, "hresult", None) != -2147418111 and "rejected" not in str(e) or k == tries - 1:
                raise
            time.sleep(1)


def excel():
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    return xl


PLACEHOLDER = 'let Source = #table(type table [tam = text], {}) in Source'


def build_model(wb, source_folder=DATA_DIR, log=print, timeout=900):
    """Nạp query vào Data Model.

    Excel.Workbook() trong Power Query gọi ngược lại Excel; nếu nạp đồng bộ (Connections.Add2 tự làm mới ngay) thì
    Excel đang bận chờ chính nó → treo. Nên: tạo connection với query rỗng (nhanh), sau đó mới gán code M thật và
    RefreshAll ở chế độ nền, chờ đến khi xong."""
    queries = read_queries()
    for name, m in queries:
        if name in ONLY_CONNECTION:
            continue
        wb.Queries.Add(name, PLACEHOLDER)
        wb.Connections.Add2(
            f"Query - {name}", f"Connection to the '{name}' query in the workbook.",
            f"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={name};Extended Properties=\"\"",
            f'"{name}"', 6, True, False)          # 6 = xlCmdTableCollection → bảng mô hình mang đúng tên query
    for name, m in queries:
        if name == "SourceFolder":
            m = '"' + str(source_folder).rstrip("\\") + "\\" + m[m.index('" meta'):]
        if name in ONLY_CONNECTION:
            wb.Queries.Add(name, m)
        else:
            wb.Queries(name).Formula = m
    conns = [wb.Connections(i) for i in range(1, wb.Connections.Count + 1)
             if wb.Connections(i).Name.startswith("Query - ")]
    for c in conns:
        c.OLEDBConnection.BackgroundQuery = True
    t = time.time()
    wb.RefreshAll()
    while True:
        time.sleep(3)
        busy = [c.Name for c in conns if c.OLEDBConnection.Refreshing]
        if not busy:
            break
        if time.time() - t > timeout:
            raise TimeoutError(f"Power Query chưa nạp xong sau {timeout}s: {busy}")
    log(f"  13 query, {len(conns)} bảng vào Data Model: {time.time() - t:.0f}s")

    model = wb.Model
    col = lambda t, c: model.ModelTables(t).ModelTableColumns(c)
    for ft, fc, pt, pc, active in RELATIONSHIPS:
        r = model.ModelRelationships.Add(col(ft, fc), col(pt, pc))
        if not active:
            r.Active = False
    log(f"  {len(RELATIONSHIPS)} quan hệ")

    fmt_money = model.GetModelFormatDecimalNumber(True, 0)
    fmt_pct = model.GetModelFormatPercentageNumber(False, 2)
    fmt_int = model.GetModelFormatWholeNumber(True)
    tbl = model.ModelTables("_Measures")
    ms = read_measures()
    for name, formula in ms:
        fmt = fmt_pct if name in PCT_MEASURES else fmt_int if name in INT_MEASURES else fmt_money
        model.ModelMeasures.Add(name, tbl, formula, fmt)
    log(f"  {len(ms)} số đo DAX")
    return model


def row_counts(model):
    return {model.ModelTables(i).Name: model.ModelTables(i).RecordCount for i in range(1, model.ModelTables.Count + 1)}
