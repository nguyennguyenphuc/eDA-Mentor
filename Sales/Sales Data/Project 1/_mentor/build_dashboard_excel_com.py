# -*- coding: utf-8 -*-
"""Dựng ../Nexa_Sales_P1_Dashboard.xlsx bằng Excel thật (Windows + Excel 365/2016+, qua COM), đúng các bước
guideline Bước 2–5 và slide Phần 4, 6, 7:

  Power Query (13 query M trong _mentor/excel_build/m) → Mô hình Dữ liệu → 12 quan hệ → 21 + 8 số đo DAX
  → 4 slicer dùng chung (Năm, Tháng, Kênh, Quốc gia) → PivotTable trên _pivot → thẻ KPI, câu Điểm chính, tiêu đề
  hành động bằng CUBEVALUE + công thức Excel → biểu đồ thường dựng từ bảng phụ (xám + một màu nhấn, đỏ chỉ cho ⚠).

Vùng ô theo guideline 5.2–5.8: slicer B4:Q6, Điểm chính B7:Q9, thẻ KPI B11:Q14, biểu đồ từ hàng 16.
Cuối script tự đổi slicer qua vài tổ hợp, so số với ref_numbers.json + Instructor Key, rồi xuất PDF kiểm tra.

Chạy:  python build_dashboard_excel_com.py            (≈ 2–3 phút; cần Excel cho Windows, pywin32, openpyxl)
"""
import json
import sys
import time
from pathlib import Path

import openpyxl

import excel_com_model as EM

sys.stdout.reconfigure(encoding="utf-8")
P1 = EM.P1
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else P1 / "Nexa_Sales_P1_Dashboard.xlsx"
PDF_DIR = P1 / "_mentor" / "kiem_tra"
REF = json.loads((P1 / "_mentor" / "ref_numbers.json").read_text(encoding="utf-8"))
THR_ABS, THR_PCT = REF["threshold"]["abs_usd"], REF["threshold"]["pct"]
ASOF = f"{REF['as_of'][8:10]}/{REF['as_of'][5:7]}/{REF['as_of'][:4]}"
DUP = REF["rows"]["fact_sales_line"] - REF["rows"]["fact_sales_line_clean"]
RM = REF["report_month"]

REASON = {"DEFECT": "Lỗi sản phẩm", "CHANGE_MIND": "Đổi ý", "WRONG_ITEM": "Giao nhầm hàng", "LATE_DELIVERY": "Giao trễ"}
CAT_VI = {"Home Appliances": "Gia dụng", "Computers": "Máy tính", "Cameras and camcorders": "Máy ảnh",
          "Cell phones": "Điện thoại", "TV and Video": "TV & Video", "Games and Toys": "Game & Đồ chơi",
          "Audio": "Âm thanh", "Music, Movies and Audio Books": "Nhạc, phim, sách nói"}
COUNTRY_VI = {"Australia": "Úc", "Canada": "Canada", "Germany": "Đức", "France": "Pháp", "Italy": "Ý",
              "Netherlands": "Hà Lan", "United Kingdom": "Anh", "United States": "Mỹ", "Online": "Online"}

FONT = "Segoe UI"
INK, INK2, MUTED, ACCENT, CTX, CTXD, DOWN, CRIT, RULE, UP, GRID = (
    "1F1F1E", "52514E", "8A8984", "2A78D6", "B9B7B0", "8F8D86", "EB6834", "D03B3B", "E4E3DF", "1C5CAB", "ECEBE7")
STORY, CARD, WHITE = "F2F6FC", "FAFAF8", "FFFFFF"
CONN = '"ThisWorkbookDataModel"'
SL = "Slicer_channel_name,Slicer_country"
PCUR = "Slicer_year,Slicer_month"
NUM0, NUM1, DPCT = "#,##0;-#,##0;0", "#,##0.0;-#,##0.0;0.0", "+0.0%;-0.0%;0.0%"
SIGN0 = '+#,##0;-#,##0;0'

# hằng số Excel
xlRowField, xlTabularRow, xlExternal, xlDesc, xlTopCount = 1, 1, 2, 2, 1
xlLine, xlColumnStacked, xlBarStacked, xlWaterfall = 4, 52, 58, 119
xlCategory, xlValue, xlExpression, xlNone, xlCircle = 1, 2, 2, -4142, 8
xlLabelAbove, xlLegendBottom = 0, -4107


def rgb(h):
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


def log(*a):
    print(*a, flush=True)


# ====================================================================== công thức CUBE
def cube(measure, *members):
    return f'CUBEVALUE({CONN},"[Measures].[{measure}]",{",".join(members)})'


def ncube(measure, *members):
    return f"N({cube(measure, *members)})"


def mid(cell):                              # thành viên dim_date[month_id] từ ô chứa số yyyymm
    return f'"[dim_date].[month_id].&["&{cell}&"]"'


CH_POS, CH_ONL = '"[dim_channel].[channel_code].&[POS]"', '"[dim_channel].[channel_code].&[ONLINE]"'


def sp(cell, d=1, pp=False):                # "+12,3%" / "−4,5%" / "–"
    unit = '" điểm %"' if pp else '"%"'
    return f'IF({cell}="","–",IF({cell}>=0,"+","−")&FIXED(ABS({cell})*100,{d})&{unit})'


def pc(cell, d=0):
    return f'IF({cell}="","–",FIXED({cell}*100,{d})&"%")'


def sg(cell):
    return f'IF({cell}>=0.5,"+",IF({cell}<=-0.5,"−",""))&FIXED(ABS({cell}),0)'


# ====================================================================== sheet _pivot: tham số + KPI
class Pv:
    """Ghi các ô trên sheet _pivot; K[key] = hàng của KPI."""

    def __init__(self, ws):
        self.ws = ws
        self.K = {}

    def put(self, addr, v):
        c = self.ws.Range(addr)
        if isinstance(v, str) and v.startswith("="):
            c.Formula = v
        else:
            c.Value = v
        return c

    def ref(self, key, col="C"):
        return f"_pivot!${col}${self.K[key]}"


def build_params(P):
    p = P.put
    p("A1", "THAM SỐ (đọc từ slicer)")
    rows = [
        ("Năm chọn", f"=CUBERANKEDMEMBER({CONN},Slicer_year,1)", "=IFERROR(VALUE(B2),\"\")"),
        ("Tháng chọn", f"=CUBERANKEDMEMBER({CONN},Slicer_month,1)", "=IFERROR(VALUE(B3),\"\")"),
        ("Số năm đang chọn", "=CUBESETCOUNT(Slicer_year)", ""),
        ("Số tháng đang chọn", "=CUBESETCOUNT(Slicer_month)", ""),
        ("Chọn đúng 1 tháng?", "=AND(B4=1,B5=1,C2<>\"\",C3<>\"\")", ""),
        ("month_id chọn", f"=IF(B6,C2*100+C3,{RM})", ""),
        ("month_id tháng trước", "=IF(B6,IF(C3=1,(C2-1)*100+12,B7-1),\"\")", ""),
        ("month_id cùng kỳ", "=IF(B6,B7-100,\"\")", ""),
        ("Nhãn kỳ", "=IF(B6,TEXT(C3,\"00\")&\"/\"&C2,\"nhiều kỳ\")", ""),
        ("Số kênh đang chọn", "=CUBESETCOUNT(Slicer_channel_name)", ""),
        ("Kênh (mục 1)", f"=CUBERANKEDMEMBER({CONN},Slicer_channel_name,1)", ""),
        ("Số quốc gia đang chọn", "=CUBESETCOUNT(Slicer_country)", ""),
        ("Quốc gia (mục 1)", f"=CUBERANKEDMEMBER({CONN},Slicer_country,1)", ""),
        ("Phạm vi (chữ)", '=IF(AND(B11=1,B12<>"All"),IF(B12="Cửa hàng","kênh cửa hàng","kênh online"),"mọi kênh")&IF(AND(B13=1,B14<>"All")," · "&B14,"")', ""),
        ("Có kênh online trong phạm vi", '=IF(AND(OR(B11<>1,B12="All",B12="Website/app"),OR(B13<>1,B14="All",B14="Online")),1,0)', ""),
    ]
    for i, (a, b, c) in enumerate(rows, 2):
        p(f"A{i}", a)
        p(f"B{i}", b)
        if c:
            p(f"C{i}", c)


# KPI: (khoá, số đo, [thành viên thêm], kiểu so sánh) — kiểu: "pct" (%), "pp" (điểm %), "measure" (số đo LM/SPLY), None
KPIS = [
    ("gross", "Gross Sales", [], "pct"), ("disc", "Discounts", [], "pct"), ("ns", "Net Sales", [], "pct"),
    ("ret", "Returns", [], "pct"), ("nr", "Net Revenue", [], "measure"), ("orders", "Orders", [], "pct"),
    ("units", "Units", [], "pct"), ("aov", "AOV", [], "pct"), ("asp", "ASP", [], "pct"),
    ("rr", "Return Rate", [], "pp"), ("fin", "Finance Net Revenue", [], "pct"), ("delta", "Δ Finance", [], None),
    ("dpct", "Δ Finance %", [], None), ("dadj", "Δ sau quy đổi", [], None),
    ("nflag", "Số cửa hàng vượt ngưỡng", [], None), ("nflagpos", "Số cửa hàng vượt ngưỡng", [CH_POS], None),
    ("hq", "HQ Adjustment", [], None), ("canc", "Cancelled Orders", [], "pct"), ("pend", "Pending Orders", [], None),
    ("onl_share", "Tỷ trọng online", [], "pp"), ("disc_g", "Tỷ lệ chiết khấu", [], "pp"),
    ("onl_ns", "Net Sales", [CH_ONL], "pct"), ("pos_ns", "Net Sales", [CH_POS], "pct"),
]


def build_kpis(P):
    p = P.put
    r0 = 20
    for j, h in enumerate(["KPI", "Số đo", "Kỳ chọn", "Tháng trước", "Cùng kỳ", "so LM", "so SPLY"]):
        p(f"{'ABCDEFG'[j]}{r0 - 1}", h)
    for i, (key, m, extra, kind) in enumerate(KPIS):
        r = r0 + i
        P.K[key] = r
        p(f"A{r}", key)
        p(f"B{r}", m)
        p(f"C{r}", "=" + ncube(m, PCUR, *extra, SL))
        if kind == "measure":
            p(f"D{r}", f'=IF($B$6,{ncube(m + " LM", PCUR, *extra, SL)},0)')
            p(f"E{r}", f'=IF($B$6,{ncube(m + " SPLY", PCUR, *extra, SL)},0)')
        elif kind:
            p(f"D{r}", f'=IF($B$6,{ncube(m, mid("$B$8"), *extra, SL)},0)')
            p(f"E{r}", f'=IF($B$6,{ncube(m, mid("$B$9"), *extra, SL)},0)')
        if kind == "pp":
            p(f"F{r}", f'=IF(OR(D{r}=0,C{r}=0),"",C{r}-D{r})')
            p(f"G{r}", f'=IF(OR(E{r}=0,C{r}=0),"",C{r}-E{r})')
        elif kind:
            p(f"F{r}", f'=IF(D{r}=0,"",(C{r}-D{r})/ABS(D{r}))')
            p(f"G{r}", f'=IF(E{r}=0,"",(C{r}-E{r})/ABS(E{r}))')
    r = r0 + len(KPIS)
    extra = [("timing", f"=C{P.K['delta']}-C{P.K['dadj']}", "Δ do online lệch thời điểm = Δ Finance − Δ sau quy đổi"),
             ("pend_all", "=" + ncube("Pending Orders", SL), "Đơn PENDING toàn thời gian, theo kênh/quốc gia"),
             ("dup", DUP, "Dòng trùng đã loại ở Power Query (63.004 → 62.884)")]
    for key, f, note in extra:
        P.K[key] = r
        p(f"A{r}", key)
        p(f"B{r}", note)
        p(f"C{r}", f)
        r += 1
    return r


# chuỗi 24 tháng (cửa sổ kết thúc ở tháng chọn) — CUBEVALUE theo month_id, chỉ nghe slicer Kênh + Quốc gia
W0, WN = 60, 24
WCOL = dict(mid="I", lab="J", nr="K", nr_sel="L", orders="M", orders_sel="N", aov="O", aov_sel="P", dadj="Q",
            flagpos="R", dadj_red="S", dadj_gray="T", rr="U", rr_sel="V", canc="W", pend="X", onl_raw="Y", onl_adj="Z")


def build_trend(P):
    p = P.put
    W = WCOL
    p(f"I{W0 - 2}", "CHUỖI 24 THÁNG (không nghe slicer Năm/Tháng; cửa sổ kết thúc ở tháng chọn)")
    heads = dict(mid="month_id", lab="Tháng", nr="Net Revenue (k)", nr_sel="nhấn", orders="Orders", orders_sel="nhấn",
                 aov="AOV", aov_sel="nhấn", dadj="Δ sau quy đổi (k)", flagpos="Số CH vượt", dadj_red="đỏ", dadj_gray="xám",
                 rr="Return Rate %", rr_sel="nhấn", canc="Huỷ", pend="Đang giao", onl_raw="Δ gốc", onl_adj="Sau quy đổi")
    for k, c in W.items():
        p(f"{c}{W0 - 1}", heads[k])
    for i in range(WN):
        r = W0 + i
        back = WN - 1 - i
        p(f"I{r}", f"=YEAR(DATE(INT($B$7/100),MOD($B$7,100)-{back},1))*100+MONTH(DATE(INT($B$7/100),MOD($B$7,100)-{back},1))")
        p(f"J{r}", f'=RIGHT("0"&MOD(I{r},100),2)&"/"&RIGHT(INT(I{r}/100),2)')
        m = mid(f"$I{r}")
        p(f"M{r}", f"=IF({ncube('Orders', m, SL)}=0,NA(),{ncube('Orders', m, SL)})")
        p(f"K{r}", f"=IF(ISNA(M{r}),NA(),{ncube('Net Revenue', m, SL)}/1000)")
        p(f"O{r}", f"=IF(ISNA(M{r}),NA(),{ncube('AOV', m, SL)})")
        p(f"Q{r}", f"=IF(ISNA(M{r}),NA(),{ncube('Δ sau quy đổi', m, SL)}/1000)")
        p(f"R{r}", "=" + ncube("Số cửa hàng vượt ngưỡng", m, CH_POS, SL))
        p(f"S{r}", f"=IF(ISNA(Q{r}),0,IF(R{r}>0,Q{r},0))")
        p(f"T{r}", f"=IF(ISNA(Q{r}),0,IF(R{r}>0,0,Q{r}))")
        p(f"U{r}", f"=IF(ISNA(M{r}),NA(),{ncube('Return Rate', m, SL)}*100)")
        p(f"W{r}", f"=IF(ISNA(M{r}),NA(),{ncube('Cancelled Orders', m, SL)})")
        p(f"X{r}", f"=IF(ISNA(M{r}),NA(),{ncube('Pending Orders', m, SL)})")
        p(f"Y{r}", f"=IF($B$16=1,{ncube('Δ Finance', m, CH_ONL, SL)}/1000,NA())")
        p(f"Z{r}", f"=IF($B$16=1,{ncube('Δ sau quy đổi', m, CH_ONL, SL)}/1000,NA())")
        last = i == WN - 1
        for a, b in (("L", "K"), ("N", "M"), ("P", "O"), ("V", "U")):
            p(f"{a}{r}", f"={b}{r}" if last else "=NA()")


# ====================================================================== PivotTable từ Mô hình Dữ liệu
PIVOTS = {}


def make_pivot(wb, P, name, addr, rows, measures, sort_by=None, top=None, fmts=None):
    R = EM.retry
    pc_ = R(lambda: wb.PivotCaches().Create(xlExternal, wb.Connections("ThisWorkbookDataModel"), 6))
    pt = R(lambda: pc_.CreatePivotTable(P.ws.Range(addr), name, True, 6))
    pt.ManualUpdate = True
    for r in rows:
        R(lambda: setattr(pt.CubeFields(r), "Orientation", xlRowField))
    for m in measures:
        R(lambda: pt.AddDataField(pt.CubeFields(f"[Measures].[{m}]")))
    pt.RowAxisLayout(xlTabularRow)
    pt.ColumnGrand = False
    pt.RowGrand = False
    pt.HasAutoFormat = False
    pt.ManualUpdate = False
    lvl = rows[0] + "." + rows[0].split(".")[-1]
    pf = pt.PivotFields(lvl)
    names = {}
    for i in range(1, pt.DataFields.Count + 1):
        df = pt.DataFields.Item(i)
        for m in measures:
            if f"[{m}]" in df.SourceName or df.Name.endswith(m):
                names.setdefault(f"[Measures].[{m}]", df)
    if top:
        pf.PivotFilters.Add2(xlTopCount, names[f"[Measures].[{top[0]}]"], top[1])
    if sort_by:
        df = names[f"[Measures].[{sort_by}]"]
        try:
            pf.AutoSort(xlDesc, df.Name)
        except Exception:
            pf.AutoSort(xlDesc, df.SourceName)
    for m, f in (fmts or {}).items():
        names[f"[Measures].[{m}]"].NumberFormat = f
    PIVOTS[name] = pt
    return pt


def pv_range(pt, k_col, n):
    """Ô dữ liệu dòng 1..n, cột k_col (0 = nhãn hàng) của pivot, địa chỉ tuyệt đối trên _pivot."""
    r0 = pt.TableRange1.Row + 1
    c = pt.TableRange1.Column + k_col
    return [f"_pivot!{col_letter(c)}${r0 + i}" for i in range(n)], (r0, c)


def pv_block(pt, k_col, n):
    """Vùng tuyệt đối n dòng dữ liệu của cột k_col trong pivot."""
    r0 = pt.TableRange1.Row + 1
    c = col_letter(pt.TableRange1.Column + k_col)
    return f"_pivot!${c}${r0}:${c}${r0 + n - 1}"


def col_letter(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def ranked_helper(P, pt, top_row, col, n, value_k, scale=1.0, label=None, map_name=None, title=""):
    """Bảng phụ cho cột ngang xếp hạng: tên (Việt hoá), giá trị, nhấn (mục 1), nền (còn lại)."""
    p = P.put
    names, _ = pv_range(pt, 0, n)
    vals, _ = pv_range(pt, value_k, n)
    c = [col_letter(col + j) for j in range(5)]
    p(f"{c[0]}{top_row - 1}", title)
    for i in range(n):
        r = top_row + i
        nm = names[i]
        vi = f'IFERROR(VLOOKUP({nm},{map_name},2,FALSE),{nm})' if map_name else nm
        p(f"{c[0]}{r}", f'=IF({nm}="","",{vi})')
        p(f"{c[1]}{r}", f'=IF({nm}="",0,N({vals[i]})*{scale})')
        p(f"{c[2]}{r}", f"={c[1]}{r}" if i == 0 else "=0")
        p(f"{c[3]}{r}", "=0" if i == 0 else f"={c[1]}{r}")
        if label:
            p(f"{c[4]}{r}", label(r, c))
    return {k: f"_pivot!${c[j]}${top_row}:${c[j]}${top_row + n - 1}" for j, k in enumerate(["name", "val", "acc", "base", "lab"])}


# ====================================================================== trang dashboard
def page(ws, title_txt, n_cards):
    ws.Activate()
    ws.Parent.Windows(1).DisplayGridlines = False
    ws.Parent.Windows(1).DisplayHeadings = False
    ws.Parent.Windows(1).Zoom = 90
    ws.Columns("A").ColumnWidth = 2
    ws.Range("B:Q").ColumnWidth = 11
    ws.Cells.Font.Name = FONT
    ws.Range("B2").Value = title_txt
    ws.Range("B2").Font.Size, ws.Range("B2").Font.Bold = 18, True
    ws.Range("B3").Value = (f"Nexa Home · Sales Project 1 · dữ liệu đến {ASOF} · nguồn: Nexa_Sales_P1_Dataset.xlsx, "
                            "finance_revenue.xlsx (phòng Tài chính – Kế toán) qua Power Query → Mô hình Dữ liệu → DAX · "
                            "USD, chưa thuế GTGT")
    font(ws.Range("B3"), 9, MUTED)
    for r in (4, 5, 6):
        ws.Rows(r).RowHeight = 31
    # dải Điểm chính B7:Q9
    ws.Range("B7").Formula = '="Điểm chính · "&_pivot!$B$10&" · "&_pivot!$B$15'
    font(ws.Range("B7"), 9, ACCENT, True)
    for r, h in ((8, 32), (9, 36)):
        rg = ws.Range(f"B{r}:Q{r}")
        rg.Merge()
        rg.WrapText = True
        rg.VerticalAlignment = -4108
        ws.Rows(r).RowHeight = h
    ws.Range("B7:Q9").Interior.Color = rgb(STORY)
    b = ws.Range("B7:B9").Borders(7)                       # xlEdgeLeft
    b.Weight, b.Color = 4, rgb(ACCENT)
    font(ws.Range("B8"), 13, INK, True)
    font(ws.Range("B9"), 10, INK2)
    # thẻ KPI B11:Q14
    wc = 3 if n_cards == 5 else 4
    cards = []
    for i in range(n_cards):
        c0 = 2 + i * wc
        cards.append(c0)
        box = ws.Range(ws.Cells(11, c0), ws.Cells(14, c0 + wc - 1))
        box.Interior.Color = rgb(CARD)
        for e in (7, 8, 9, 10):
            bb = box.Borders(e)
            bb.LineStyle, bb.Color, bb.Weight = 1, rgb(RULE), 2
        for r in range(11, 15):
            ws.Range(ws.Cells(r, c0), ws.Cells(r, c0 + wc - 1)).Merge()
        font(ws.Cells(11, c0), 9, INK2)
        font(ws.Cells(12, c0), 18, INK, True)
        font(ws.Cells(13, c0), 9, INK2)
        font(ws.Cells(14, c0), 9, INK2)
        ws.Cells(12, c0).HorizontalAlignment = -4131
    ws.Rows(12).RowHeight = 30
    ws.Range("B52").Value = (f"Δ = số Kế toán − dữ liệu đơn; vượt ngưỡng khi |Δ| > {THR_ABS:,.0f} USD và |Δ %| > "
                             f"{THR_PCT * 100:.0f}% ở mức tháng × cửa hàng. Chỉ tính đơn COMPLETED, đã loại {DUP} dòng trùng "
                             "ở Power Query; hàng trả lại tính theo ngày trả. Kênh online không thuộc quốc gia nào. "
                             "Thẻ KPI = CUBEVALUE trên số đo DAX; % thay đổi và câu chữ = công thức Excel.")
    font(ws.Range("B52"), 8, MUTED)
    ws.Range("B52:Q53").Merge()
    ws.Range("B52").WrapText = True
    ws.Range("B52").VerticalAlignment = -4160
    ps = ws.PageSetup
    ps.PrintArea = "$A$1:$Q$53"
    ps.Orientation = 2
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    return cards


def font(rg, size=None, color=None, bold=None, italic=None):
    f = rg.Font
    f.Name = FONT
    if size:
        f.Size = size
    if color:
        f.Color = rgb(color)
    if bold is not None:
        f.Bold = bold
    if italic is not None:
        f.Italic = italic


def so_sanh(cell, lab, pp):
    so = f'FIXED(ABS({cell})*100,2)&" điểm %"' if pp else f'FIXED(ABS({cell})*100,1)&"%"'
    return f'=IF({cell}="","– {lab}",IF({cell}>0,"▲ ",IF({cell}<0,"▼ ","■ "))&{so}&" {lab}")'


def card(ws, c0, label, value, fmt, comps=(), notes=()):
    """comps: [(ô thay đổi, chữ, tăng_là_tốt, điểm_%)]; notes: [công thức/chữ]."""
    ws.Cells(11, c0).Value = label
    v = ws.Cells(12, c0)
    v.Formula = value
    v.NumberFormat = fmt
    lines = [so_sanh(cell, lab, pp) for cell, lab, good, pp in comps] + list(notes)
    for i, t in enumerate(lines[:2]):
        c = ws.Cells(13 + i, c0)
        if t.startswith("="):
            c.Formula = t
        else:
            c.Value = t
    for i, (cell, lab, good, pp) in enumerate(comps[:2]):
        rg = ws.Cells(13 + i, c0)
        a = rg.Address
        for sym, col in (("▲", UP if good else DOWN), ("▼", DOWN if good else UP)):
            fc = rg.FormatConditions.Add(xlExpression, None, f'=LEFT({a},1)="{sym}"')
            fc.Font.Color = rgb(col)
            fc.Font.Bold = True


def red_if(ws, addr, cond):
    fc = ws.Range(addr).FormatConditions.Add(xlExpression, None, "=" + cond)
    fc.Font.Color = rgb(CRIT)
    fc.Font.Bold = True


def title(ws, addr, formula, sub):
    c = ws.Range(addr)
    c.Formula = formula
    font(c, 11, INK, True)
    s = ws.Cells(c.Row + 1, c.Column)
    s.Value = sub
    font(s, 8, MUTED)


# ---------------------------------------------------------------------- biểu đồ
def new_chart(ws, area, ctype):
    rg = ws.Range(area)
    shp = ws.Shapes.AddChart2(-1, ctype, rg.Left + 2, rg.Top + 2, rg.Width - 4, rg.Height - 4)
    ch = shp.Chart
    while ch.SeriesCollection().Count:
        ch.SeriesCollection(1).Delete()
    return ch


def finish(ch, legend=False, value_axis=True, gridlines=True, fmt=None):
    ch.HasTitle = False
    ch.HasLegend = legend
    if legend:
        ch.Legend.Position = xlLegendBottom
    ca = ch.ChartArea.Format
    ca.Line.ForeColor.RGB = rgb(RULE)
    ca.TextFrame2.TextRange.Font.Name = FONT
    ca.TextFrame2.TextRange.Font.Size = 8
    ca.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(MUTED)
    try:
        ax = ch.Axes(xlValue)
        if not value_axis:
            ax.Delete()
        else:
            ax.HasMajorGridlines = gridlines
            if gridlines:
                ax.MajorGridlines.Format.Line.ForeColor.RGB = rgb(GRID)
            ax.Format.Line.Visible = False
            if fmt:
                ax.TickLabels.NumberFormat = fmt
        ch.Axes(xlCategory).Format.Line.ForeColor.RGB = rgb(RULE)
    except Exception:
        pass


def add_series(ch, P, name, values, cats=None):
    s = ch.SeriesCollection().NewSeries()
    s.Name = name
    s.Values = P.ws.Range(values.split("!")[1]) if "!" in values else values
    if cats:
        s.XValues = P.ws.Range(cats.split("!")[1])
    return s


def color_series(s, color, line=False, weight=2.0):
    if line:
        s.Format.Line.ForeColor.RGB = rgb(color)
        s.Format.Line.Weight = weight
        s.MarkerStyle = xlNone
    else:
        s.Format.Fill.ForeColor.RGB = rgb(color)
        s.Format.Line.Visible = False


def labels(s, fmt, color=INK, pos=None):
    s.HasDataLabels = True
    dl = s.DataLabels()
    dl.NumberFormat = fmt
    dl.Format.TextFrame2.TextRange.Font.Size = 8
    dl.Format.TextFrame2.TextRange.Font.Bold = True
    dl.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(color)
    if pos is not None:
        dl.Position = pos


def wcol(k):
    c = WCOL[k]
    return f"_pivot!${c}${W0}:${c}${W0 + WN - 1}"


def line_chart(ws, P, area, keys, colors, names=None, fmt="#,##0", sel=True, legend=False):
    ch = new_chart(ws, area, xlLine)
    for i, (k, col) in enumerate(zip(keys, colors)):
        s = add_series(ch, P, (names or keys)[i], wcol(k), wcol("lab"))
        color_series(s, col, line=True)
    if sel:                                                # chuỗi nhấn: chỉ điểm tháng chọn, marker xanh + nhãn
        s = add_series(ch, P, "Tháng chọn", wcol(keys[-1] + "_sel"), wcol("lab"))
        s.Format.Line.Visible = False
        s.MarkerStyle, s.MarkerSize = xlCircle, 7
        s.MarkerBackgroundColor = s.MarkerForegroundColor = rgb(ACCENT)
        labels(s, fmt, ACCENT, xlLabelAbove)
    ch.Axes(xlCategory).TickLabelSpacing = 3
    ch.Axes(xlCategory).TickLabelPosition = -4134          # nhãn tháng ở đáy, kể cả khi có giá trị âm
    finish(ch, legend=legend, fmt=fmt)
    if legend and sel:
        ch.Legend.LegendEntries(ch.Legend.LegendEntries().Count).Delete()
    return ch


def stacked(ws, P, area, rng, horizontal=True, fmt="#,##0.0", cat_key="name"):
    ch = new_chart(ws, area, xlBarStacked if horizontal else xlColumnStacked)
    for k, col in (("acc", ACCENT), ("base", CTX)):
        s = add_series(ch, P, k, rng[k], rng[cat_key])
        color_series(s, col)
        labels(s, f"{fmt};-{fmt};;")
    ch.ChartGroups(1).GapWidth = 60
    ch.ChartGroups(1).Overlap = 100
    if horizontal:
        ch.Axes(xlCategory).ReversePlotOrder = True
    finish(ch, value_axis=not horizontal, gridlines=not horizontal, fmt=fmt)
    return ch


def month_flag_chart(ws, P, area):
    ch = new_chart(ws, area, xlColumnStacked)
    for k, col in (("dadj_red", CRIT), ("dadj_gray", CTX)):
        s = add_series(ch, P, k, wcol(k), wcol("lab"))
        color_series(s, col)
    ch.ChartGroups(1).GapWidth = 40
    ch.ChartGroups(1).Overlap = 100
    ch.Axes(xlCategory).TickLabelSpacing = 3
    ch.Axes(xlCategory).TickLabelPosition = -4134          # xlTickLabelPositionLow
    finish(ch, fmt="#,##0.0")
    return ch


def waterfall(ws, P, area, cats, vals, colors, total_idx):
    rg = ws.Range(area)
    shp = ws.Shapes.AddChart2(395, xlWaterfall, rg.Left + 2, rg.Top + 2, rg.Width - 4, rg.Height - 4)
    ch = shp.Chart
    ch.SetSourceData(P.ws.Range(f"{cats.split('!')[1].split(':')[0]}:{vals.split('!')[1].split(':')[1]}"))
    ws.Application.CalculateUntilAsyncQueriesDone()
    s = ch.FullSeriesCollection(1)
    n = s.Points().Count
    assert n == len(colors), f"waterfall có {n} điểm, cần {len(colors)}: {P.ws.Range(cats.split('!')[1]).Value} {P.ws.Range(vals.split('!')[1]).Value} {ch.FullSeriesCollection().Count}"
    for i, col in enumerate(colors, 1):
        pt = s.Points().Item(i)
        if i in total_idx:
            pt.IsTotal = True
        pt.Format.Fill.ForeColor.RGB = rgb(col)
    P.ws.Range(vals.split("!")[1]).NumberFormat = "#,##0.0;-#,##0.0;0.0"   # nhãn waterfall lấy định dạng ô nguồn
    try:
        s.HasDataLabels = True
    except Exception:
        pass
    for f in (lambda: setattr(ch, "HasTitle", False), lambda: setattr(ch, "HasLegend", False),
              lambda: setattr(ch.ChartArea.Format.Line.ForeColor, "RGB", rgb(RULE)),
              lambda: setattr(ch.ChartArea.Font, "Name", FONT), lambda: setattr(ch.ChartArea.Font, "Size", 8)):
        try:                                                # Waterfall (Excel 2016) không hỗ trợ mọi thuộc tính cũ
            f()
        except Exception:
            pass
    return ch


def table(ws, r0, c0, head, rows, fmts, merge=None):
    """Bảng hiển thị trỏ vào pivot; merge = {cột: số ô gộp}."""
    merge = merge or {}
    for j, h in enumerate(head):
        c = ws.Cells(r0, c0 + j)
        c.Value = h
        font(c, 9, INK2, True)
        c.HorizontalAlignment = -4131 if j == 0 or fmts[j] == "@" else -4152
        bb = c.Borders(9)
        bb.LineStyle, bb.Color = 1, rgb(RULE)
    for i, cells in enumerate(rows):
        for j, v in enumerate(cells):
            if v is None:
                continue
            c = ws.Cells(r0 + 1 + i, c0 + j)
            c.Formula = v
            font(c, 9, INK)
            c.NumberFormat = fmts[j]
            c.HorizontalAlignment = -4131 if fmts[j] == "@" else -4152
        rg = ws.Range(ws.Cells(r0 + 1 + i, c0), ws.Cells(r0 + 1 + i, c0 + len(head) - 1)).Borders(9)
        rg.LineStyle, rg.Color, rg.Weight = 1, rgb(RULE), 1


# ====================================================================== dựng toàn bộ
def build(xl):
    wb = xl.Workbooks.Add()
    while wb.Worksheets.Count < 6:
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    names = ["Tong_quan", "Doanh_thu", "Doi_soat", "Hoan_tra", "Huong_dan", "_pivot"]
    for i, n in enumerate(names, 1):
        wb.Worksheets(i).Name = n
    wb.Styles("Normal").Font.Name = FONT
    wb.Styles("Normal").Font.Size = 10
    S1, S2, S3, S4, SH, SP = (wb.Worksheets(n) for n in names)
    P = Pv(SP)

    log("1/6 Power Query → Mô hình Dữ liệu → quan hệ → DAX")
    model = EM.build_model(wb, log=log)
    counts = EM.row_counts(model)
    for t, n in (("fact_sales_line", REF["rows"]["fact_sales_line_clean"]), ("fact_return", REF["rows"]["fact_return"]),
                 ("fact_finance", REF["rows"]["finance_revenue"]), ("dim_date", REF["rows"]["dim_date"]),
                 ("dim_store", REF["rows"]["dim_store"]), ("dim_product", REF["rows"]["dim_product"])):
        assert counts[t] == n, f"{t}: {counts[t]} dòng, phải {n}"
    log("  số dòng khớp guideline 2.8:", {k: counts[k] for k in ("fact_sales_line", "fact_return", "fact_finance")})

    EM.wait_ready(xl)
    log("2/6 Slicer")
    src = wb.Connections("ThisWorkbookDataModel")
    SLICERS = [("Slicer_year", "[dim_date].[year]", "Năm", "B4:E6", 6, [f"[dim_date].[year].&[{RM // 100}]"]),
               ("Slicer_month", "[dim_date].[month]", "Tháng", "F4:K6", 12, [f"[dim_date].[month].&[{RM % 100}]"]),
               ("Slicer_channel_name", "[dim_channel].[channel_name]", "Kênh bán", "L4:M6", 1, None),
               ("Slicer_country", "[dim_store].[country]", "Quốc gia", "N4:Q6", 3, None)]
    caches = {}
    for nm, field, cap, area, ncol, default in SLICERS:
        sc = wb.SlicerCaches.Add2(src, field, nm)
        caches[nm] = sc
        if default:
            sc.VisibleSlicerItemsList = default
    # slicer: đặt cỡ lưới trước (cột B–Q rộng 11, hàng 4–6 cao 28) để slicer nằm gọn B4:Q6; cùng cache → dùng chung
    for ws in (S1, S2, S3, S4):
        ws.Columns("A").ColumnWidth = 2
        ws.Range("B:Q").ColumnWidth = 11
        ws.Rows("4:6").RowHeight = 31
    for nm, field, cap, area, ncol, default in SLICERS:
        for ws in (S1, S2, S3, S4):
            rg = ws.Range(area)
            sl = caches[nm].Slicers.Add(ws, f"{field}.[{field.split('.')[-1][1:-1]}]", f"{nm}_{ws.Name}", cap,
                                        rg.Top + 1, rg.Left + 2, rg.Width - 4, rg.Height - 2)
            sl.NumberOfColumns = ncol
            sl.RowHeight = 15
            sl.Style = "SlicerStyleLight1"



    log("3/6 _pivot: tham số, KPI, chuỗi 24 tháng, PivotTable")
    build_params(P)
    build_kpis(P)
    build_trend(P)
    # bảng Việt hoá nhãn (danh mục, quốc gia, lý do trả hàng) + cửa hàng → quốc gia
    P.put("AB1", "VIỆT HOÁ NHÃN")
    r = 2
    for d in (CAT_VI, COUNTRY_VI, REASON):
        for k, v in d.items():
            P.put(f"AB{r}", k)
            P.put(f"AC{r}", v)
            r += 1
    map_end = r - 1
    st = openpyxl.load_workbook(EM.DATA_DIR / "Nexa_Sales_P1_Dataset.xlsx", read_only=True)["dim_store"]
    rows_ = list(st.iter_rows(values_only=True))
    h = rows_[0]
    for row in rows_[1:]:
        P.put(f"AB{r}", row[h.index("store_code")])
        P.put(f"AC{r}", COUNTRY_VI.get(row[h.index("country")], row[h.index("country")]))
        P.put(f"AB{r + 100}", row[h.index("store_name")])
        P.put(f"AC{r + 100}", COUNTRY_VI.get(row[h.index("country")], row[h.index("country")]))
        r += 1
    SP.Range(f"AB2:AC{map_end}").Name = "map_vi"
    SP.Range(f"AB{map_end + 1}:AC{r - 1}").Name = "map_store"
    SP.Range(f"AB{map_end + 101}:AC{r + 99}").Name = "map_store_name"

    # pivot: cột AF trở đi, mỗi pivot một khối
    pv_cat = make_pivot(wb, P, "pvt_danh_muc", "AF3", ["[dim_product].[category]"], ["Net Sales"], sort_by="Net Sales")
    pv_cty = make_pivot(wb, P, "pvt_quoc_gia", "AF20", ["[dim_store].[country]"], ["Net Sales"], sort_by="Net Sales")
    pv_rsn = make_pivot(wb, P, "pvt_ly_do", "AF40", ["[fact_return].[return_reason]"], ["Returns"], sort_by="Returns")
    pv_crr = make_pivot(wb, P, "pvt_rr_danh_muc", "AF55", ["[dim_product].[category]"], ["Return Rate", "Net Sales"],
                        sort_by="Return Rate")
    pv_ch = make_pivot(wb, P, "pvt_kenh", "AF75", ["[dim_channel].[channel_name]"], ["Net Sales", "Orders", "AOV"])
    pv_top = make_pivot(wb, P, "pvt_top10_cua_hang", "AF85", ["[dim_store].[store_name]"], ["Net Sales", "Orders", "AOV"],
                        sort_by="Net Sales", top=("Net Sales", 10))
    pv_rec = make_pivot(wb, P, "pvt_doi_soat_top10", "AF105", ["[dim_store].[store_code]"],
                        ["Net Revenue", "Finance Net Revenue", "Δ Finance", "Δ Finance %", "Δ sau quy đổi", "Vượt ngưỡng",
                         "Δ tuyệt đối"], sort_by="Δ tuyệt đối", top=("Δ tuyệt đối", 10))
    pv_all = make_pivot(wb, P, "pvt_cua_hang_co_so_lieu", "AF125", ["[dim_store].[store_code]"],
                        ["Net Revenue", "Finance Net Revenue"])
    for sc in caches.values():
        for pt in PIVOTS.values():
            sc.PivotTables.AddPivotTable(pt)

    # bảng phụ cho biểu đồ xếp hạng (cột AQ trở đi)
    cat_rng = ranked_helper(P, pv_cat, 4, 43, 8, 1, 1 / 1000, map_name="map_vi", title="Danh mục (nghìn USD)")
    cty_rng = ranked_helper(P, pv_cty, 21, 43, 9, 1, 1 / 1000, map_name="map_vi", title="Quốc gia (nghìn USD)")
    rsn_rng = ranked_helper(P, pv_rsn, 41, 43, 4, 1, 1, map_name="map_vi", title="Lý do trả hàng (USD)")
    crr_rng = ranked_helper(
        P, pv_crr, 56, 43, 8, 1, 100, map_name="map_vi", title="Tỷ lệ trả theo danh mục (%)",
        label=lambda r, c: (f'=IF({c[0]}{r}="","",{c[0]}{r}&" · NS "&FIXED(N(' +
                            pv_range(pv_crr, 2, 8)[0][r - 56] + ')/1000,1)&"k")'))
    crr_rng["cat"] = crr_rng["lab"]

    # đối soát top 10: cờ ⚠ cửa hàng POS (không gồm ONL) và danh sách mã
    codes, _ = pv_range(pv_rec, 0, 10)
    cols = {k: pv_range(pv_rec, j, 10)[0] for j, k in enumerate(
        ["code", "nr", "fin", "delta", "dpct", "dadj", "flag", "abs"])}
    P.put("AQ104", "Top 10 |Δ|: cờ POS · đếm dồn · mã thứ k · Δ của mã thứ k")
    for i in range(10):
        r = 105 + i
        P.put(f"AQ{r}", f'=IF(AND(N({cols["flag"][i]})=1,{cols["code"][i]}<>"ONL",{cols["code"][i]}<>""),1,0)')
        P.put(f"AR{r}", f"=SUM($AQ$105:AQ{r})")
        P.put(f"AS{r}", f'=IFERROR(INDEX({pv_block(pv_rec, 0, 10)},MATCH({i + 1},$AR$105:$AR$114,0)),"")')
        P.put(f"AT{r}", f'=IFERROR(N(INDEX({pv_block(pv_rec, 3, 10)},MATCH({i + 1},$AR$105:$AR$114,0))),0)')
        P.put(f"AU{r}", f'=IF(AQ{r}=1,N({cols["dadj"][i]}),0)')
    P.put("AQ116", "Δ sau quy đổi của các cửa hàng POS vượt ngưỡng")
    P.put("AU116", "=SUM(AU105:AU114)")
    n_all = f'MAX(0,COUNTA(_pivot!$AF$126:$AF$300))'
    # waterfall: Tổng quan (Gross → Net) và Đối soát (phân rã Δ)
    k = P.ref
    wf1 = [("Gross Sales", f"={k('gross')}/1000"), ("Chiết khấu", f"=-{k('disc')}/1000"),
           ("Hàng trả lại", f"=-{k('ret')}/1000"), ("Net Revenue", f"={k('nr')}/1000")]
    wf2 = [("Online lệch thời điểm", f"={k('timing')}/1000"), ("Cửa hàng ⚠", "=_pivot!$AU$116/1000"),
           ("Còn lại", f"=({k('dadj')}-_pivot!$AU$116)/1000"), ("Δ Kế toán − đơn", f"={k('delta')}/1000")]
    for top_, rows_wf in ((130, wf1), (136, wf2)):
        P.put(f"AQ{top_ - 1}", "Waterfall")
        for i, (lab, f) in enumerate(rows_wf):           # nằm ngang: Excel vẽ waterfall thành 1 chuỗi 4 điểm
            P.put(f"{col_letter(43 + i)}{top_}", lab)
            P.put(f"{col_letter(43 + i)}{top_ + 1}", f)

    # ================================================================== Trang 1 · Tổng quan
    log("4/6 Bốn sheet dashboard")
    cL = lambda key: k(key, "F")
    cS = lambda key: k(key, "G")
    cards = page(S1, "1 · Tổng quan: tháng này bán được bao nhiêu, có khớp Kế toán không?", 5)
    S1.Range("B8").Formula = (f'="Net Revenue "&_pivot!$B$10&" đạt "&FIXED({k("nr")}/1000,1)&" nghìn USD — "&{sp(cL("nr"))}'
                              f'&" so với tháng trước, "&{sp(cS("nr"))}&" so với cùng kỳ năm trước."')
    b1 = (f'IF(OR({cS("orders")}="",{cS("aov")}=""),"",IF(ABS(LN(1+{cS("orders")}))>=ABS(LN(1+{cS("aov")})),'
          f'"• Biến động so với cùng kỳ đến chủ yếu từ số đơn ("&{sp(cS("orders"))}&"); giá trị mỗi đơn (AOV) "&{sp(cS("aov"))}&".",'
          f'"• Biến động so với cùng kỳ đến chủ yếu từ giá trị mỗi đơn (AOV "&{sp(cS("aov"))}&"); số đơn "&{sp(cS("orders"))}&"."))')
    b2 = (f'"• Kế toán ghi nhận "&IF({k("delta")}>=0,"cao","thấp")&" hơn dữ liệu đơn "&FIXED(ABS({k("delta")}),0)'
          f'&" USD ("&{sp(k("dpct"), 2)}&" trên số Kế toán); "&{k("nflag")}&" cửa hàng vượt ngưỡng cần giải thích (sheet Doi_soat)."')
    S1.Range("B9").Formula = f"={b1}&CHAR(10)&{b2}"
    card(S1, cards[0], "Net Revenue", f"={k('nr')}/1000", '#,##0.0" nghìn USD"',
         [(cL("nr"), "so với tháng trước", True, False), (cS("nr"), "so với cùng kỳ", True, False)])
    card(S1, cards[1], "Số đơn hoàn tất", f"={k('orders')}", '#,##0" đơn"',
         [(cL("orders"), "so với tháng trước", True, False), (cS("orders"), "so với cùng kỳ", True, False)])
    card(S1, cards[2], "AOV", f"={k('aov')}", '#,##0" USD/đơn"',
         [(cL("aov"), "so với tháng trước", True, False), (cS("aov"), "so với cùng kỳ", True, False)])
    card(S1, cards[3], "Tỷ lệ hàng bán bị trả lại", f"={k('rr')}", "0.00%",
         [(cL("rr"), "so với tháng trước", False, True), (cS("rr"), "so với cùng kỳ", False, True)])
    card(S1, cards[4], "Chênh lệch với Kế toán (Δ)", f"={k('delta')}", '+#,##0" USD";-#,##0" USD";0" USD"', (),
         [f'={sp(k("dpct"), 2)}&" trên số Kế toán"', f'={k("nflag")}&" cửa hàng vượt ngưỡng"'])
    red_if(S1, S1.Cells(12, cards[4]).Address, f"{k('nflag')}>0")
    nr_c, nr_s = k("nr"), k("nr", "E")
    ratio = f'IF({nr_s}=0,"",{nr_c}/{nr_s})'
    title(S1, "B16", f'="Net Revenue "&_pivot!$B$10&" bằng "&{pc(ratio)}&" cùng kỳ năm trước"',
          "Net Revenue theo tháng đặt hàng, nghìn USD · tháng đang chọn tô xanh")
    line_chart(S1, P, "B18:J32", ["nr"], [CTX], fmt="#,##0")
    lost = f'IF({k("gross")}=0,"",({k("disc")}+{k("ret")})/{k("gross")})'
    title(S1, "K16", f'="Chiết khấu và hàng trả lại lấy đi "&{pc(lost, 1)}&" Gross Sales"',
          "Tháng đang chọn, nghìn USD")
    waterfall(S1, P, "K18:Q32", "_pivot!$AQ$130:$AT$130", "_pivot!$AQ$131:$AT$131", [CTXD, DOWN, DOWN, ACCENT], {4})
    title(S1, "B34", f'="Số đơn "&{sp(cS("orders"))}&" so với cùng kỳ"', "Số đơn hoàn tất theo tháng")
    line_chart(S1, P, "B36:E48", ["orders"], [ACCENT], fmt="#,##0")
    title(S1, "F34", f'="AOV "&{sp(cS("aov"))}&" so với cùng kỳ"', "AOV = Net Sales / số đơn, USD")
    line_chart(S1, P, "F36:J48", ["aov"], [CTX], fmt="#,##0")
    title(S1, "K34", f'="Sau quy đổi online, Δ "&_pivot!$B$10&" còn "&{sg(k("dadj"))}&" USD"',
          "Δ sau quy đổi theo tháng, nghìn USD · đỏ = có cửa hàng vượt ngưỡng")
    month_flag_chart(S1, P, "K36:Q48")

    # ================================================================== Trang 2 · Doanh thu
    cards = page(S2, "2 · Doanh thu đến từ đâu: kênh, danh mục, quốc gia, cửa hàng", 4)
    top_cat, top_cty = cat_rng["name"].split(":")[0], cty_rng["name"].split(":")[0]
    top_cat_v, top_cty_v = cat_rng["val"].split(":")[0], cty_rng["val"].split(":")[0]
    sh_cat = pc(f'{top_cat_v}*1000/{k("ns")}')
    sh_cty = pc(f'{top_cty_v}*1000/{k("ns")}')
    S2.Range("B8").Formula = (f'=IF({k("ns")}=0,"Không có doanh thu trong phạm vi đang chọn.",{top_cat}&" dẫn đầu với "&'
                              f'{sh_cat}&" Net Sales "&_pivot!$B$10&"; "&{top_cty}&'
                              f'" là thị trường lớn nhất ("&{sh_cty}&").")')
    b1 = (f'IF(_pivot!$B$16=1,"• Online chiếm "&{pc(k("onl_share"))}&" Net Sales; so với cùng kỳ online "&'
          f'{sp(cS("onl_ns"))}&", cửa hàng "&{sp(cS("pos_ns"))}&".","• Bộ lọc đang chọn không gồm kênh online.")')
    b2 = f'"• Chiết khấu bằng "&{pc(k("disc_g"), 1)}&" Gross Sales tháng này."'
    S2.Range("B9").Formula = f"={b1}&CHAR(10)&{b2}"
    card(S2, cards[0], "Net Sales", f"={k('ns')}/1000", '#,##0.0" nghìn USD"',
         [(cL("ns"), "so với tháng trước", True, False), (cS("ns"), "so với cùng kỳ", True, False)])
    card(S2, cards[1], "Số sản phẩm bán", f"={k('units')}", '#,##0" sản phẩm"',
         [(cL("units"), "so với tháng trước", True, False), (cS("units"), "so với cùng kỳ", True, False)])
    card(S2, cards[2], "Tỷ trọng online (% Net Sales)", f"={k('onl_share')}", "0.0%",
         [(cL("onl_share"), "so với tháng trước", True, True), (cS("onl_share"), "so với cùng kỳ", True, True)])
    card(S2, cards[3], "Chiết khấu / Gross Sales", f"={k('disc_g')}", "0.0%",
         [(cL("disc_g"), "so với tháng trước", False, True), (cS("disc_g"), "so với cùng kỳ", False, True)])
    title(S2, "B16", f'={top_cat}&" chiếm "&{sh_cat}&" Net Sales"',
          "Net Sales tháng đang chọn theo danh mục, nghìn USD")
    stacked(S2, P, "B18:G34", cat_rng)
    title(S2, "H16", f'={top_cty}&" đóng góp "&{sh_cty}&" Net Sales"',
          "Net Sales theo quốc gia, nghìn USD (online là một dòng)")
    stacked(S2, P, "H18:L34", cty_rng)
    chn, _ = pv_range(pv_ch, 0, 2)
    chv = {j: pv_range(pv_ch, j, 2)[0] for j in (1, 2, 3)}
    title(S2, "M16", f'="AOV "&IF({chn[0]}="","–",{chn[0]}&" "&FIXED(N({chv[3][0]}),0))&IF({chn[1]}="","",", "&{chn[1]}&" "&FIXED(N({chv[3][1]}),0))&" USD"',
          "Theo kênh; so với tháng trước (LM) và cùng kỳ (SPLY)")
    rows_ch, rows_cmp = [], []
    for i in range(2):
        nm = chn[i]
        mem = f'"[dim_channel].[channel_name].&["&{nm}&"]"'
        chg = lambda cell: (f'=IF(OR({nm}="",NOT(_pivot!$B$6)),"",IFERROR(N({chv[1][i]})/'
                            f'{ncube("Net Sales", mem, mid(cell), "Slicer_country")}-1,"–"))')
        rows_ch.append([f'=IF({nm}="","",{nm})', f'=IF({nm}="","",N({chv[1][i]}))', f'=IF({nm}="","",N({chv[2][i]}))',
                        f'=IF({nm}="","",N({chv[3][i]}))'])
        rows_cmp.append([f'=IF({nm}="","",{nm})', chg("_pivot!$B$8"), chg("_pivot!$B$9")])
    table(S2, 18, 13, ["Kênh", "Net Sales", "Đơn", "AOV"], rows_ch, ["@", NUM0, NUM0, NUM0])
    table(S2, 22, 13, ["Kênh", "vs LM", "vs SPLY"], rows_cmp, ["@", DPCT, DPCT])
    tn, _ = pv_range(pv_top, 0, 10)
    tv = {j: pv_range(pv_top, j, 10)[0] for j in (1, 2, 3)}
    top10 = f'IF({k("ns")}=0,"",SUM({pv_block(pv_top, 1, 10)})/{k("ns")})'
    title(S2, "B36", f'="10 cửa hàng lớn nhất tạo "&{pc(top10)}&" Net Sales"',
          "Top 10 cửa hàng theo Net Sales tháng đang chọn (PivotTable Top 10 trên _pivot)")
    rows_t = []
    for i in range(10):
        nm = tn[i]
        mem = f'"[dim_store].[store_name].&["&{nm}&"]"'
        chg = lambda cell: (f'=IF(OR({nm}="",NOT(_pivot!$B$6)),"",IFERROR(N({tv[1][i]})/'
                            f'{ncube("Net Sales", mem, mid(cell), SL)}-1,"–"))')
        rows_t.append([f'=IF({nm}="","",{i + 1})', f'=IF({nm}="","",{nm})', None, None,
                       f'=IF({nm}="","",IFERROR(VLOOKUP({nm},map_store_name,2,FALSE),""))', None,
                       f'=IF({nm}="","",N({tv[1][i]}))', None, f'=IF({nm}="","",N({tv[2][i]}))',
                       f'=IF({nm}="","",N({tv[3][i]}))', chg("_pivot!$B$8"), chg("_pivot!$B$9")])
    table(S2, 38, 2, ["#", "Cửa hàng", "", "", "Quốc gia", "", "Net Sales", "", "Đơn", "AOV", "vs LM", "vs SPLY"], rows_t,
          ["0", "@", "@", "@", "@", "@", NUM0, "@", NUM0, NUM0, DPCT, DPCT])
    db = S2.Range("H39:H48").FormatConditions.AddDatabar()
    db.BarColor.Color = rgb(ACCENT)

    # ================================================================== Trang 3 · Đối soát
    cards = page(S3, "3 · Đối soát với Kế toán: lệch bao nhiêu, ở đâu, vì sao?", 5)
    n_flag_pos = "_pivot!$AR$114"
    S3.Range("B8").Formula = (f'="Kế toán "&IF({k("delta")}>=0,"cao","thấp")&" hơn dữ liệu đơn "&FIXED(ABS({k("delta")}),0)&" USD ("&'
                              f'{sp(k("dpct"), 2)}&")."&IF(ABS({k("timing")})>1," "&{sg(k("timing"))}&'
                              f'" USD là lệch thời điểm của kênh online — biến mất khi tính theo ngày giao.","")')
    lst = '_pivot!$AS$105&IF(_pivot!$AS$106<>"",", "&_pivot!$AS$106,"")&IF(_pivot!$AS$107<>"",", "&_pivot!$AS$107,"")'
    b1 = (f'IF({n_flag_pos}>0,"• Còn lại "&{sg(k("dadj"))}&" USD; cần điều tra "&{n_flag_pos}&" cửa hàng vượt ngưỡng: "&{lst}&" ⚠.",'
          f'"• Còn lại "&{sg(k("dadj"))}&" USD, không cửa hàng nào vượt ngưỡng ngoài kênh online.")')
    b2 = (f'IF({k("hq")}<>0,"• Chưa gồm bút toán điều chỉnh HQ-ADJ "&{sg(k("hq"))}&" USD (không thuộc cửa hàng nào).",'
          f'"• Xuống mức ngày ghi sổ (posting_date) ở các cửa hàng có ⚠ để tìm nguyên nhân và số tiền.")')
    S3.Range("B9").Formula = f"={b1}&CHAR(10)&{b2}"
    card(S3, cards[0], "Net Revenue · dữ liệu đơn", f"={k('nr')}/1000", '#,##0.0" nghìn USD"',
         [(cL("nr"), "so với tháng trước", True, False), (cS("nr"), "so với cùng kỳ", True, False)])
    card(S3, cards[1], "Net Revenue · Kế toán", f"={k('fin')}/1000", '#,##0.0" nghìn USD"', (),
         [f'=IF({k("hq")}<>0,"chưa gồm HQ-ADJ "&{sg(k("hq"))}&" USD","không gồm HQ-ADJ")'])
    card(S3, cards[2], "Chênh lệch Δ", f"={k('delta')}", '+#,##0" USD";-#,##0" USD";0" USD"', (),
         [f'={sp(k("dpct"), 2)}&" trên số Kế toán"'])
    card(S3, cards[3], "Δ sau quy đổi online", f"={k('dadj')}", '+#,##0" USD";-#,##0" USD";0" USD"', (),
         ["online tính theo ngày giao"])
    card(S3, cards[4], "Cửa hàng vượt ngưỡng", f'={k("nflag")}&" / "&{n_all}', "@", (),
         [f"|Δ| > {THR_ABS:,.0f} USD và |Δ%| > {THR_PCT * 100:.0f}%", "trên số cửa hàng có số liệu"])
    for c0, cond in ((cards[2], f"{k('nflag')}>0"), (cards[3], f"{n_flag_pos}>0"), (cards[4], f"{k('nflag')}>0")):
        red_if(S3, S3.Cells(12, c0).Address, cond)
    title(S3, "B16", f'=IF(_pivot!$AS$105<>"",_pivot!$AS$105&" là chênh lệch thật lớn nhất ("&{sg("_pivot!$AT$105")}&" USD)","Chênh lệch chủ yếu là lệch thời điểm")',
          "Phân rã Δ tháng đang chọn, nghìn USD · không gồm HQ-ADJ")
    waterfall(S3, P, "B18:G32", "_pivot!$AQ$136:$AT$136", "_pivot!$AQ$137:$AT$137", [CTX, CRIT, CTXD, ACCENT], {4})
    title(S3, "H16", f'="Δ sau quy đổi "&_pivot!$B$10&" là "&{sg(k("dadj"))}&" USD"',
          "Δ sau khi quy online về ngày giao, nghìn USD · tháng có cửa hàng vượt ngưỡng tô đỏ")
    month_flag_chart(S3, P, "H18:Q32")
    title(S3, "B34", f'=IF({k("nflag")}>0,{k("nflag")}&" cửa hàng vượt ngưỡng tháng "&_pivot!$B$10,"Không cửa hàng nào vượt ngưỡng tháng "&_pivot!$B$10)',
          "Top 10 theo Δ tuyệt đối (PivotTable) · ⚠ = |Δ| > ngưỡng USD và |Δ %| > ngưỡng %")
    rows_r = []
    for i in range(10):
        cd = cols["code"][i]
        rows_r.append([f'=IF({cd}="","",{cd}&" · "&IFERROR(VLOOKUP({cd},map_store,2,FALSE),""))', None,
                       f'=IF({cd}="","",ROUND(N({cols["nr"][i]}),2))', f'=IF({cd}="","",ROUND(N({cols["fin"][i]}),2))',
                       f'=IF({cd}="","",ROUND(N({cols["delta"][i]}),2))', f'=IF({cd}="","",ROUND(N({cols["dpct"][i]}),6))',
                       f'=IF({cd}="","",ROUND(N({cols["dadj"][i]}),2))',
                       f'=IF({cd}="","",IF(N({cols["flag"][i]})=1,"⚠ Vượt","trong ngưỡng"))'])
    table(S3, 36, 2, ["Cửa hàng", "", "Đơn hàng", "Kế toán", "Δ", "Δ %", "Δ sau quy đổi", "Cờ"], rows_r,
          ["@", "@", NUM0, NUM0, SIGN0, DPCT, SIGN0, "@"])
    red_if(S3, "B37:I46", '$I37="⚠ Vượt"')
    title(S3, "K34", '=IF(_pivot!$B$16=1,"Δ online về gần 0 khi tính theo ngày giao","Kênh online không thuộc bộ lọc đang chọn")',
          "Δ kênh online theo tháng, nghìn USD · xám = gốc, xanh = sau quy đổi")
    line_chart(S3, P, "K36:Q50", ["onl_raw", "onl_adj"], [CTX, ACCENT], ["Δ gốc", "Sau quy đổi"], fmt="#,##0.0",
               sel=False, legend=True)

    # ================================================================== Trang 4 · Hoàn trả & chất lượng
    cards = page(S4, "4 · Hàng bán bị trả lại và chất lượng dữ liệu", 4)
    top_r, top_rv = rsn_rng["name"].split(":")[0], rsn_rng["val"].split(":")[0]
    S4.Range("B8").Formula = (f'="Hàng trả lại "&_pivot!$B$10&" là "&FIXED({k("ret")},0)&" USD, bằng "&{pc(k("rr"), 2)}&" Net Sales ("&'
                              f'{sp(cS("rr"), 2, True)}&" so với cùng kỳ)."')
    b1 = (f'IF({k("ret")}=0,"• Không có hàng trả lại trong phạm vi đang chọn.","• Lý do lớn nhất: "&{top_r}&" ("&'
          f'{pc(top_rv + "/" + k("ret"))}&" giá trị trả).")')
    b2 = (f'"• Đã loại {DUP} dòng trùng ở Power Query; "&FIXED({k("pend_all")},0)&'
          f'" đơn đang giao (PENDING) không tính doanh thu — toàn thời gian, theo bộ lọc kênh/quốc gia."')
    S4.Range("B9").Formula = f"={b1}&CHAR(10)&{b2}"
    card(S4, cards[0], "Giá trị hàng trả lại", f"={k('ret')}", '#,##0" USD"',
         [(cL("ret"), "so với tháng trước", False, False), (cS("ret"), "so với cùng kỳ", False, False)])
    card(S4, cards[1], "Tỷ lệ hàng bán bị trả lại", f"={k('rr')}", "0.00%",
         [(cL("rr"), "so với tháng trước", False, True), (cS("rr"), "so với cùng kỳ", False, True)])
    card(S4, cards[2], "Đơn huỷ", f"={k('canc')}", '#,##0" đơn"',
         [(cL("canc"), "so với tháng trước", False, False), (cS("canc"), "so với cùng kỳ", False, False)])
    card(S4, cards[3], "Dòng trùng đã loại", f"={k('dup')}", '#,##0" dòng"', (),
         ["ở Power Query (Remove Duplicates)", f'=FIXED({k("pend_all")},0)&" đơn PENDING tại {ASOF}"'])
    title(S4, "B16", f'="Tỷ lệ hàng trả lại "&{pc(k("rr"), 2)}&" — "&IF({k("rr")}>{k("rr", "E")},"cao","thấp")&" hơn cùng kỳ "&'
                     f'IF({cS("rr")}="","–",FIXED(ABS({cS("rr")})*100,2))&" điểm %"', "Returns / Net Sales theo tháng, %")
    line_chart(S4, P, "B18:J32", ["rr"], [CTX], fmt="0.0")
    title(S4, "K16", f'=IF({k("ret")}=0,"Lý do trả hàng",{top_r}&" là lý do trả hàng lớn nhất")',
          "Giá trị trả tháng đang chọn theo lý do, USD")
    stacked(S4, P, "K18:Q32", rsn_rng, fmt="#,##0")
    title(S4, "B34", f'={crr_rng["name"].split(":")[0]}&" có tỷ lệ trả hàng cao nhất ("&FIXED({crr_rng["val"].split(":")[0]},1)&"%)"',
          "Tỷ lệ hàng trả lại theo danh mục tháng đang chọn, % · nhãn kèm Net Sales để đọc cùng quy mô")
    stacked(S4, P, "B36:J48", crr_rng, cat_key="cat")
    title(S4, "K34", f'=FIXED({k("canc")},0)&" đơn huỷ tháng "&_pivot!$B$10', "Số đơn CANCELLED (xanh) và PENDING (xám) theo tháng đặt hàng")
    line_chart(S4, P, "K36:Q48", ["canc", "pend"], [ACCENT, CTX], ["Huỷ", "Đang giao"], sel=False, legend=True)

    # ================================================================== Hướng dẫn + hoàn thiện
    SH.Activate()
    wb.Windows(1).DisplayGridlines = False
    SH.Columns("B").ColumnWidth = 120
    texts = [
        "Dashboard Excel tham chiếu — Sales Project 1 (Nexa Home)",
        "Dựng ĐÚNG các bước của hướng dẫn: Data → Queries & Connections có 13 truy vấn Power Query (code trong "
        "hoc_vien/excel_build/m); Power Pivot → Manage có Mô hình Dữ liệu 10 bảng, 12 quan hệ (delivery_date inactive) và "
        "29 số đo DAX (21 bắt buộc + 8 nâng cao) trong bảng _Measures.",
        "Cách dùng: chọn Năm, Tháng, Kênh bán, Quốc gia bằng 4 slicer ở đầu trang — slicer dùng chung, đổi ở sheet nào cả 4 "
        "sheet cùng đổi. Mặc định 12/2020, mọi kênh.",
        "Thẻ KPI, câu Điểm chính, tiêu đề biểu đồ: ô CUBEVALUE trên sheet _pivot (ẩn) đọc số đo DAX theo slicer; % thay đổi "
        "và câu chữ là công thức Excel (FIXED, IF). Biểu đồ xếp hạng và bảng: PivotTable từ Mô hình Dữ liệu trên _pivot, "
        "nối với cả 4 slicer (Report Connections). Chuỗi 24 tháng: CUBEVALUE theo month_id, chỉ nghe slicer Kênh và Quốc gia.",
        f"Đổi máy: Data → Queries & Connections → mở query SourceFolder, sửa đường dẫn thư mục chứa "
        f"Nexa_Sales_P1_Dataset.xlsx và finance_revenue.xlsx (kết thúc bằng \\), rồi Data → Làm mới Tất cả (Refresh All).",
        f"Quy ước: Δ = số Kế toán − dữ liệu đơn; vượt ngưỡng khi |Δ| > {THR_ABS:,.0f} USD và |Δ%| > {THR_PCT * 100:.0f}% ở "
        "mức tháng × cửa hàng; bút toán HQ-ADJ không thuộc cửa hàng nào.",
        "Sinh lại file: python _mentor/build_dashboard_excel_com.py (Windows + Excel). Bản đối chiếu số dựng bằng công thức "
        "thường: _mentor/Nexa_Sales_P1_Doi_chieu_cong_thuc.xlsx.",
    ]
    for i, t in enumerate(texts, 2):
        c = SH.Range(f"B{i}")
        c.Value = t
        c.WrapText = True
        c.VerticalAlignment = -4160
        font(c, 14 if i == 2 else 10, INK, i == 2)
    SP.Columns("A:B").ColumnWidth = 24
    SP.Visible = 0                                          # xlSheetHidden
    S1.Activate()
    S1.Range("A1").Select()
    return wb, caches


# ====================================================================== kiểm tra số
def errs(ws):
    """Ô lỗi (#N/A, #VALUE!…) trên sheet: COM trả lỗi dưới dạng số nguyên âm lớn."""
    ur = ws.UsedRange
    v = ur.Value
    out = []
    for i, row in enumerate(v or ()):
        for j, x in enumerate(row):
            if isinstance(x, int) and x < -2146820000:
                out.append(ws.Cells(ur.Row + i, ur.Column + j).Address)
    return out


def check(xl, wb, caches):
    P = wb.Worksheets("_pivot")
    S1, S3 = wb.Worksheets("Tong_quan"), wb.Worksheets("Doi_soat")
    K = {}
    for r in range(20, 60):
        v = P.Range(f"A{r}").Value
        if v:
            K[v] = r

    def val(key, col="C"):
        return P.Range(f"{col}{K[key]}").Value

    def sel(year=None, month=None, channel=None, country=None):
        for nm, v, f in (("Slicer_year", year, "[dim_date].[year].&[{}]"), ("Slicer_month", month, "[dim_date].[month].&[{}]"),
                         ("Slicer_channel_name", channel, "[dim_channel].[channel_name].&[{}]"),
                         ("Slicer_country", country, "[dim_store].[country].&[{}]")):
            sc = caches[nm]
            if v is None:
                sc.ClearManualFilter()
            else:
                sc.VisibleSlicerItemsList = [f.format(v)]
        xl.CalculateUntilAsyncQueriesDone()
        xl.Calculate()
        xl.CalculateUntilAsyncQueriesDone()

    errors = []

    def eq(name, got, exp, tol=0.01):
        ok = got is not None and not isinstance(got, str) and abs(got - exp) <= tol
        log(f"  {'OK ' if ok else 'SAI'} {name}: {got} (chuẩn {exp})")
        if not ok:
            errors.append(name)

    sel(2020, 12)
    cur, lm, sply = REF["month"]["current"], REF["month"]["lm"], REF["month"]["sply"]
    for key, rk in (("gross", "gross_sales"), ("disc", "discounts"), ("ns", "net_sales"), ("ret", "returns"),
                    ("nr", "net_revenue"), ("orders", "orders"), ("units", "units")):
        eq(f"{key} 12/2020", val(key), cur[rk])
        eq(f"{key} LM", val(key, "D"), lm[rk])
        eq(f"{key} SPLY", val(key, "E"), sply[rk])
    eq("AOV 12/2020", val("aov"), cur["aov"], 0.01)
    eq("Δ Finance 12/2020", val("delta"), -12536.23)
    eq("Số cửa hàng vượt ngưỡng 12/2020", val("nflag"), 2, 0)
    # lỗi hiển thị trên 4 sheet
    for ws in (wb.Worksheets(n) for n in ("Tong_quan", "Doanh_thu", "Doi_soat", "Hoan_tra")):
        bad = errs(ws)
        if bad:
            errors.append(f"{ws.Name} có ô lỗi {bad[:5]}")
            log(f"  SAI {ws.Name}: ô lỗi {bad[:5]}")
    log("  câu Điểm chính Tong_quan:", S1.Range("B8").Value)
    log("  câu Điểm chính Doi_soat:", S3.Range("B8").Value, "|", S3.Range("B9").Value)

    sel(2020, 6, country="United States")
    eq("06/2020 Mỹ: số cửa hàng vượt ngưỡng", val("nflag"), 2, 0)
    flagged = sorted(P.Range(f"AS{r}").Value for r in (105, 106))
    log("  06/2020 Mỹ: cửa hàng ⚠", flagged)
    if flagged != ["S45", "S54"]:
        errors.append("06/2020 Mỹ phải là S45, S54")
    sel(2020, 2)
    log("  02/2020 mã ⚠ đầu:", P.Range("AS105").Value, P.Range("AT105").Value)
    sel(2020, 12, channel="Website/app")
    eq("12/2020 online: Δ sau quy đổi ≈ 0", val("dadj"), 0, 1)
    sel(2020, 12, channel="Cửa hàng")
    for ws in (wb.Worksheets(n) for n in ("Tong_quan", "Doanh_thu", "Doi_soat", "Hoan_tra")):
        bad = errs(ws)
        if bad:
            errors.append(f"POS: {ws.Name} có ô lỗi {bad[:5]}")
    sel(2020, 12)                                           # trả về mặc định trước khi lưu
    return errors


def export_pdf(wb):
    PDF_DIR.mkdir(exist_ok=True)
    out = []
    for n in ("Tong_quan", "Doanh_thu", "Doi_soat", "Hoan_tra"):
        f = PDF_DIR / f"excel_{n}.pdf"
        wb.Worksheets(n).ExportAsFixedFormat(0, str(f))
        out.append(f)
    return out


def main():
    t = time.time()
    xl = EM.excel()
    wb = None
    try:
        wb, caches = build(xl)
        log("5/6 Kiểm số")
        errors = check(xl, wb, caches)
        wb.Worksheets("Tong_quan").Activate()
        if OUT.exists():
            OUT.unlink()
        wb.SaveAs(str(OUT), 51)
        log("6/6 Lưu", OUT.name, "· xuất PDF:", [p.name for p in export_pdf(wb)])
        if errors:
            raise SystemExit("LỖI: " + "; ".join(errors))
        log(f"Xong sau {time.time() - t:.0f}s, mọi phép kiểm khớp.")
    finally:
        if wb is not None:
            wb.Close(False)
        xl.Quit()


if __name__ == "__main__":
    main()
