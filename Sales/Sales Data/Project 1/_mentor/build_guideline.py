# -*- coding: utf-8 -*-
"""Dựng ../Dashboard_Guideline (dành cho mentor).docx

Khuôn: guideline của Bank Project 1 (giữ nguyên style, khổ giấy, bảng màu) — xoá nội dung cũ, viết nội dung Sales.
Mọi con số đọc từ ref_numbers.json, Instructor Key và chính file dataset => chạy lại sau generate_dataset.py.
Hình: lấy PNG xuất từ các file .drawio trong ../diagrams (sửa sơ đồ ở .drawio, chạy build_diagrams.py rồi chạy lại file này).
"""
import copy
import json
from pathlib import Path

import docx
import pandas as pd
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

HERE = Path(__file__).resolve().parent
P1 = HERE.parent
BANK = P1.parents[2] / "Bank" / "Project 1" / "Dashboard_Guideline (dành cho mentor).docx"
OUT = P1 / "Dashboard_Guideline (dành cho mentor).docx"
PNG = P1 / "diagrams" / "png"
R = json.loads((HERE / "ref_numbers.json").read_text(encoding="utf-8"))
KEY = pd.read_excel(HERE / "Nexa_Sales_P1_Instructor_Key.xlsx", sheet_name=None)
DS = pd.read_excel(P1 / "Nexa_Sales_P1_Dataset.xlsx", sheet_name=None)


def nf(v, d=0):
    s = f"{abs(v):,.{d}f}".replace(",", "§").replace(".", ",").replace("§", ".")
    return ("−" if v < 0 else "") + s


def pct(v, d=1):
    return nf(v * 100, d) + "%"


# ------------------------------------------------------------------ khuôn tài liệu
doc = docx.Document(BANK)
body = doc.element.body
for el in list(body):
    if el.tag != qn("w:sectPr"):
        body.remove(el)
STY = {s.name: s for s in doc.styles}          # tên style pandoc ("Heading 1") — tra trực tiếp, tránh ánh xạ của python-docx
num_part = doc.part.numbering_part.element
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
BULLET_ABS = None
DEC_ABS = None
for an in num_part.findall(W + "abstractNum"):
    fmt = an.find(W + "lvl").find(W + "numFmt").get(W + "val")
    if fmt == "bullet" and BULLET_ABS is None and an.find(W + "lvl").find(W + "lvlText").get(W + "val") == "•":
        BULLET_ABS = an.get(W + "abstractNumId")
    if fmt == "decimal" and DEC_ABS is None:
        DEC_ABS = an.get(W + "abstractNumId")
_next_num = [5000]


def new_num(abs_id):
    _next_num[0] += 1
    n = OxmlElement("w:num")
    n.set(qn("w:numId"), str(_next_num[0]))
    a = OxmlElement("w:abstractNumId")
    a.set(qn("w:val"), abs_id)
    n.append(a)
    ov = OxmlElement("w:lvlOverride")
    ov.set(qn("w:ilvl"), "0")
    so = OxmlElement("w:startOverride")
    so.set(qn("w:val"), "1")
    ov.append(so)
    n.append(ov)
    num_part.append(n)
    return _next_num[0]


def runs(p, text):
    """**đậm** và `code` đơn giản trong một đoạn."""
    import re
    for part in re.split(r"(\*\*[^*]+\*\*|`[^`]+`)", text):
        if not part:
            continue
        if part.startswith("**"):
            p.add_run(part[2:-2]).bold = True
        elif part.startswith("`"):
            r = p.add_run(part[1:-1])
            r.font.name = "Consolas"
            r.font.size = Pt(9.5)
        else:
            p.add_run(part)
    return p


def H1(t): doc.add_paragraph(t, style=STY["Heading 1"])
def H2(t): doc.add_paragraph(t, style=STY["Heading 2"])
def H3(t): doc.add_paragraph(t, style=STY["Heading 3"])
def P(t, first=False): return runs(doc.add_paragraph(style=STY["First Paragraph"] if first else STY["Body Text"]), t)


def LIST(items, numbered=False):
    nid = new_num(DEC_ABS if numbered else BULLET_ABS)
    for it in items:
        p = runs(doc.add_paragraph(style=STY["Compact"]), it)
        pPr = p._p.get_or_add_pPr()
        numPr = OxmlElement("w:numPr")
        il = OxmlElement("w:ilvl"); il.set(qn("w:val"), "0")
        ni = OxmlElement("w:numId"); ni.set(qn("w:val"), str(nid))
        numPr.append(il); numPr.append(ni)
        pPr.insert(0, numPr) if pPr.find(qn("w:pStyle")) is None else pPr.find(qn("w:pStyle")).addnext(numPr)


def CODE(text):
    p = doc.add_paragraph(style=STY["Source Code"])
    for i, line in enumerate(text.strip("\n").split("\n")):
        if i:
            p.add_run().add_break()
        p.add_run(line)


def TABLE(head, rows, widths=None, align=None):
    t = doc.add_table(rows=1 + len(rows), cols=len(head))
    t.style = STY["Table"]
    tblPr = t._tbl.tblPr
    w_ = OxmlElement("w:tblW"); w_.set(qn("w:type"), "pct"); w_.set(qn("w:w"), "5000"); tblPr.append(w_)
    b = OxmlElement("w:tblBorders")
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        e = OxmlElement(f"w:{side}")
        for k, v in [("val", "single"), ("sz", "4"), ("space", "0"), ("color", "BED0B4")]:
            e.set(qn(f"w:{k}"), v)
        b.append(e)
    tblPr.append(b)
    trPr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true"); trPr.append(th)
    for ri, row in enumerate([head] + rows):
        for ci, val in enumerate(row):
            cell = t.rows[ri].cells[ci]
            p = cell.paragraphs[0]
            p.style = STY["Compact"]
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            if align and ri and align[ci] == "r":
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if ri == 0:
                r = p.add_run(str(val)); r.bold = True; r.font.color.rgb = RGBColor(255, 255, 255); r.font.size = Pt(8.5)
                shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), "447A2A")
                cell._tc.get_or_add_tcPr().append(shd)
            else:
                runs(p, str(val))
                for r in p.runs:
                    r.font.size = Pt(9)
    if widths:
        t.autofit = False
        grid = t._tbl.tblGrid
        for ci, (gc, wcm) in enumerate(zip(grid.findall(qn("w:gridCol")), widths)):
            gc.set(qn("w:w"), str(int(wcm * 567)))
            for row in t.rows:
                row.cells[ci].width = Cm(wcm)
    doc.add_paragraph(style=STY["Compact"])
    return t


def FIG(png, caption, width_cm=17.5):
    p = doc.add_paragraph(style=STY["Captioned Figure"])
    p.add_run().add_picture(str(PNG / png), width=Cm(width_cm))
    doc.add_paragraph(caption, style=STY["Image Caption"])


# ------------------------------------------------------------------ số liệu dùng nhiều lần
rows = R["rows"]
cur, lm, sp = R["month"]["current"], R["month"]["lm"], R["month"]["sply"]
ytd, ytd_sp = R["ytd"]["current"], R["ytd"]["sply"]
thr = R["threshold"]
br = R["bridge_2020"]
anom = KEY["planted_anomalies"]
over = KEY["over_threshold"]
over20 = over[over.month_id // 100 == 2020].copy()
adj_cols = [c for c in over20.columns if c.startswith("adj_")]
CAUSE_VI = {"adj_online_timing": "Online lệch thời điểm", "adj_refund_lag": "Hoàn tiền hạch toán trễ",
            "adj_cancelled_booked": "Đơn huỷ vẫn ghi nhận", "adj_missing_upload": "Finance thiếu dữ liệu ngày",
            "adj_vat_included": "Doanh thu gồm VAT", "adj_duplicate_posting": "Bút toán ghi trùng"}
over20["cause"] = over20[adj_cols].abs().idxmax(axis=1).map(CAUSE_VI)
fin_rows = rows["finance_revenue"]
hq20 = float(anom[(anom.type == "E_MANUAL_ADJ") & (pd.to_datetime(anom.date_from).dt.year == 2020)].effect_usd.sum())
n_canc_2020 = int(((anom.type == "A_CANCELLED_BOOKED") & (pd.to_datetime(anom.date_from).dt.year == 2020)).sum())
mstr = lambda m: f"{m % 100:02d}/{m // 100}"

# ------------------------------------------------------------------ NỘI DUNG
H1("Hướng dẫn xây dựng Dashboard Power BI: Kết quả bán hàng và đối soát doanh thu (Nexa Home · Sales Project 1)")
P("Cập nhật: 10/10/2026 · Dataset v" + R["version"] + f" · seed {R['seed']}", first=True)

H2("Giới thiệu và cách dùng tài liệu")
P("Tài liệu này hướng dẫn dựng dashboard Power BI 4 trang (Tổng quan, Doanh thu, Đối soát, Hoàn trả & chất lượng) "
  "từ dataset Nexa Home, rồi dùng chính dashboard đó để đối soát doanh thu với file Finance và viết báo cáo. "
  "Đây là project đầu tiên của nhóm Sales, mức hỗ trợ của mentor khoảng 80 %: mentor đi cùng học viên qua từng bước, "
  "học viên tự làm phần giải thích chênh lệch.", first=True)
P("**Đối tượng.** Mentor và trợ giảng dựng bản tham chiếu; học viên làm theo từng bước. **Bước 6** chứa số đối soát "
  "và đáp án các chênh lệch, phải ẩn khi dùng tài liệu này làm đề.")
P("**Sản phẩm đầu ra.** Một file .pbix gồm mô hình dữ liệu, bộ measure DAX và 4 trang dashboard; kèm báo cáo 8 phần "
  "theo guild_report_template.pdf, trong đó phần chính là bảng giải thích chênh lệch vượt ngưỡng năm 2020.")

H3("Bộ file đầu vào")
TABLE(["File", "Công dụng", "Ai được xem"], [
    ["Nexa_Sales_P1_Dataset.xlsx", "Dữ liệu đơn hàng, hoàn trả, các bảng dim, bảng map mã Finance, định nghĩa KPI, giao dịch mẫu", "Học viên"],
    ["finance_revenue.xlsx", "Doanh thu Finance đã xác nhận theo ngày × mã cửa hàng Finance (bản rút gọn)", "Học viên"],
    ["Nexa_Sales_P1_Dashboard.html", "Bản dashboard tham chiếu mở bằng trình duyệt, dùng để so hình và so số", "Học viên (sau buổi 1) / mentor"],
    ["Nexa_Sales_P1_Dashboard.xlsx", "Dashboard Excel một trang (chọn tháng/kênh, KPI bằng SUMIFS, sheet Doi_soat có cờ) cho ai không dùng Power BI", "Học viên (sau buổi 1) / mentor"],
    ["Slide_Project_1.pptx", "Slide giảng 8 phần; phụ lục mentor đặt ẩn", "Học viên + mentor"],
    ["guild_report_template.pdf", "Khung báo cáo 8 phần", "Học viên"],
    ["diagrams/*.drawio (+ .pdf)", "Sơ đồ pipeline, mô hình, quy trình đối soát, cây KPI; sửa ở .drawio rồi xuất lại", "Học viên + mentor"],
    ["_mentor/Nexa_Sales_P1_Instructor_Key.xlsx", "Đáp án: 31 chênh lệch cài sẵn, cầu nối từng tháng × cửa hàng", "Chỉ mentor"],
    ["_mentor/generate_dataset.py", "Sinh lại toàn bộ dữ liệu từ nguồn GER (seed cố định)", "Chỉ mentor"],
], widths=[5.2, 9, 3.3])

H3("Dữ liệu trong một cái nhìn")
TABLE(["Nội dung", "Giá trị"], [
    ["Bối cảnh", "Nexa Home bán điện máy qua cửa hàng (POS) và website/app; Finance cần biết doanh thu thật và lý do lệch số"],
    ["Nguồn nền", "Global Electronics Retailer (Maven Analytics), giữ nguyên địa lý: 66 cửa hàng ở 8 quốc gia + 1 kênh online"],
    ["Khoảng thời gian", "Đơn hàng 01/01/2016 → 20/02/2021; ngày trích dữ liệu 20/02/2021"],
    ["Quy mô", f"{nf(rows['orders_total'])} đơn · {nf(rows['fact_sales_line'])} dòng bán (gồm 120 dòng trùng) · "
               f"{nf(rows['fact_return'])} lần trả hàng · {nf(fin_rows)} dòng Finance"],
    ["Tiền tệ", "USD, chưa gồm VAT; VAT giả định 10 % chỉ xuất hiện ở file Finance"],
    ["Tháng báo cáo mặc định", f"{mstr(R['report_month'])}; so với tháng trước (LM) và cùng tháng năm trước (SPLY)"],
    ["Phạm vi đối soát", "Năm 2020, mức tháng × cửa hàng × kênh"],
    ["Ngưỡng", f"Vượt ngưỡng khi |Δ| > {nf(thr['abs_usd'])} USD **và** |Δ %| > {pct(thr['pct'], 0)}"],
], widths=[4.5, 13])

H3("Lộ trình 7 bước")
TABLE(["Bước", "Việc chính", "Thời lượng gợi ý"], [
    ["1. Hiểu dữ liệu", "Đọc README, Data_Dictionary, KPI_Definitions, sample_transactions", "30 phút"],
    ["2. Power Query", "Nạp 2 file, đặt kiểu, bỏ dòng trùng, map mã Finance → fact_finance", "45 phút"],
    ["3. Mô hình", "Quan hệ star schema, đánh dấu bảng ngày, quan hệ inactive cho delivery_date", "30 phút"],
    ["4. DAX", "Khoảng 30 measure: doanh thu, đơn hàng, Finance, chênh lệch, so sánh kỳ", "90 phút"],
    ["5. Layout", "Dựng 4 trang theo bảng vị trí visual", "90 phút"],
    ["6. Đối soát", "So số với Bước 6; lập bảng giải thích chênh lệch năm 2020", "60 phút"],
    ["7. Báo cáo", "Viết báo cáo 8 phần theo template", "Ở nhà"],
], widths=[3.5, 11, 3])

H3("Quy ước trong tài liệu")
LIST(["Tên bảng và cột viết dạng `fact_sales_line[order_date]`; tên measure viết dạng `[Net Revenue]`.",
      f"Mọi con số mẫu lấy từ chính dataset, tại tháng báo cáo {mstr(R['report_month'])} nếu không ghi khác.",
      "“LM” là tháng trước (Last Month), “SPLY” là cùng tháng năm trước (Same Period Last Year).",
      "Δ (chênh lệch) luôn tính **Finance − dữ liệu đơn**: Δ dương nghĩa là Finance ghi nhận cao hơn.",
      "Số viết theo kiểu Việt Nam: dấu chấm phân cách nghìn, dấu phẩy thập phân."])
FIG("01_pipeline_du_lieu.png", "Hình 1: Pipeline từ nguồn dữ liệu đến báo cáo (nguồn: diagrams/01_pipeline_du_lieu.drawio)")

# ---------------------------------------------------------------- Bước 1
H2("Bước 1. Hiểu dữ liệu và chỉ tiêu")
P("Dataset gồm 2 bảng fact, 7 bảng dim, 1 bảng map và file Finance riêng. Nắm grain (mỗi dòng là gì) của từng bảng "
  "trước khi viết bất kỳ measure nào.", first=True)
H3("1.1 Các bảng trong dataset")
cnt = {k: len(v) for k, v in DS.items()}
TABLE(["Bảng", "Vai trò", "Grain (mỗi dòng là)", "Khoá", "Số dòng", "Nhóm"], [
    ["fact_sales_line", "Fact", "1 sản phẩm trong 1 đơn", "order_number + line_item", nf(cnt["fact_sales_line"]), "Bắt buộc"],
    ["fact_return", "Fact", "1 lần trả hàng của 1 dòng", "return_id", nf(cnt["fact_return"]), "Bắt buộc"],
    ["finance_revenue (file riêng)", "Fact", "1 ngày × 1 mã cửa hàng Finance", "—", nf(fin_rows), "Bắt buộc"],
    ["dim_date", "Dim", "1 ngày", "date", nf(cnt["dim_date"]), "Bắt buộc"],
    ["dim_store", "Dim", "1 cửa hàng (0 = online)", "store_key", nf(cnt["dim_store"]), "Bắt buộc"],
    ["dim_product", "Dim", "1 sản phẩm", "product_key", nf(cnt["dim_product"]), "Bắt buộc"],
    ["dim_channel", "Dim", "1 kênh", "channel_code", nf(cnt["dim_channel"]), "Bắt buộc"],
    ["dim_promotion", "Dim", "1 chương trình KM", "promo_code", nf(cnt["dim_promotion"]), "Bắt buộc"],
    ["dim_order_status", "Dim", "1 trạng thái đơn", "order_status", nf(cnt["dim_order_status"]), "Bắt buộc"],
    ["map_store_finance", "Map", "1 mã Finance trong 1 khoảng hiệu lực", "fin_store_code", nf(cnt["map_store_finance"]), "Bắt buộc"],
    ["dim_customer", "Dim", "1 khách hàng", "customer_key", nf(cnt["dim_customer"]), "Bổ sung"],
    ["fx_rate", "Tham khảo", "1 ngày × 1 tiền tệ", "date + currency_code", nf(cnt["fx_rate"]), "Bổ sung"],
], widths=[3.6, 1.6, 4.2, 3.6, 1.8, 1.8], align="lllllr")
P("Nhóm “Bắt buộc” đủ để dựng 4 trang và đối soát. Nhóm “Bổ sung” dành cho phần sáng tạo (20 điểm trong rubric), "
  "ví dụ phân tích theo vùng khách hàng hoặc quy đổi tiền tệ.")
H3("1.2 Từ chỉ tiêu nghiệp vụ đến cột dữ liệu")
kpi = DS["KPI_Definitions"]
TABLE(["Chỉ tiêu", "Công thức", "Ghi chú"], [[r.kpi, r.formula, "" if pd.isna(r.scope_note) else r.scope_note]
                                            for r in kpi.itertuples()], widths=[3.6, 6.4, 7.5])
FIG("04_cay_kpi.png", "Hình 2: Cây KPI (nguồn: diagrams/04_cay_kpi.drawio)")
H3("1.3 Chín quy tắc dữ liệu cần nhớ")
LIST([
    "**Chỉ đơn COMPLETED.** CANCELLED (khách huỷ) và PENDING (online chưa giao tại 20/02/2021) không phải doanh thu. "
    f"Có {nf(rows['orders_cancelled'])} đơn huỷ và {nf(rows['orders_pending'])} đơn đang giao.",
    "**Dòng trùng.** 120 dòng xuất hiện 2 lần y hệt; khoá dòng là order_number + line_item. Bỏ trùng trước khi tính.",
    "**Khuyến mãi áp cho cả đơn.** discount_amount đã tính sẵn ở từng dòng = quantity × unit_price × discount_pct.",
    "**Hoàn trả tính theo ngày trả.** Returns lấy fact_return[return_date], không phải ngày đặt của đơn gốc; một dòng có thể trả một phần.",
    "**Online có 2 mốc ngày.** Dữ liệu đơn dùng order_date; Finance ghi nhận online theo delivery_date (trung bình 4–5 ngày sau).",
    "**POS không có delivery_date.** Khách nhận hàng ngay tại quầy, Finance ghi nhận đúng ngày bán.",
    "**Mã Finance khác store_key.** Dùng map_store_finance; cửa hàng 38 (Anh) đổi mã GBR-03 → GBR-03N từ 01/07/2019, "
    "phải map theo valid_from/valid_to. Mã HQ-ADJ không thuộc cửa hàng nào.",
    "**Finance hạch toán hoàn tiền trễ.** refunds trong file Finance có thể lệch 0–3 ngày so với return_date.",
    "**USD, chưa VAT.** Mọi số tiền phía đơn hàng là USD chưa VAT; so với cột net_revenue của Finance, không so với total_incl_vat.",
])
H3("1.4 Cột được phép để trống")
P("fact_sales_line[delivery_date] trống với mọi đơn POS và mọi đơn online chưa COMPLETED; dim_store[square_meters] trống "
  "ở dòng online; dim_promotion[window_from_mmdd]/[window_to_mmdd] trống với NONE và WEB10; map_store_finance[store_key] "
  "trống ở HQ-ADJ. Mọi cột khác phải đầy đủ.", first=True)

# ---------------------------------------------------------------- Bước 2
H2("Bước 2. Nạp và chuẩn bị dữ liệu trong Power Query")
P("Khác với dataset Bank, dữ liệu ở đây **chưa sạch hoàn toàn**: có dòng trùng và file Finance dùng mã riêng. "
  "Hai việc làm sạch này là một phần của bài.", first=True)
H3("2.1 Các việc cần làm")
LIST(["Get data → Excel workbook → Nexa_Sales_P1_Dataset.xlsx; tick 10 bảng: các dim, fact_sales_line, fact_return, "
      "map_store_finance (dim_customer, fx_rate nếu làm phần sáng tạo). Không tick README, Data_Dictionary, QC_Summary.",
      "Get data → Excel workbook → finance_revenue.xlsx, sheet finance_revenue; đặt tên query `finance_raw`.",
      "Đặt kiểu dữ liệu theo mục 2.2 cho từng query.",
      "fact_sales_line: Home → Remove rows → Remove duplicates trên toàn bộ cột (còn đúng "
      f"{nf(rows['fact_sales_line_clean'])} dòng).",
      "Tạo query `fact_finance` từ finance_raw theo mục 2.3 (gắn store_key theo mã và ngày hiệu lực).",
      "Tắt Enable load cho finance_raw và map_store_finance sau khi fact_finance chạy đúng.",
      "Bật View → Column quality và Column profile; đối chiếu số dòng ở mục 2.4."], numbered=True)
H3("2.2 Kiểu dữ liệu cần đặt")
TABLE(["Cột", "Kiểu"], [
    ["order_date, delivery_date, return_date, posting_date, dim_date[date], valid_from, valid_to, open_date", "Date"],
    ["unit_price, discount_amount, unit_cost, return_amount, revenue_recognized, refunds, net_revenue, vat_amount", "Fixed decimal number"],
    ["quantity, return_qty, order_number, line_item, store_key, product_key, customer_key, month_id", "Whole number"],
    ["discount_pct", "Percentage"],
    ["Các cột mã/tên còn lại", "Text"],
], widths=[13.5, 4])
H3("2.3 Gắn store_key cho file Finance")
P("map_store_finance có 2 dòng cho cửa hàng 38 (mã cũ và mã mới). Nối theo mã rồi giữ dòng có posting_date nằm trong "
  "khoảng hiệu lực:", first=True)
CODE("""let
    Src    = finance_raw,
    Merged = Table.NestedJoin(Src, {"fin_store_code"}, map_store_finance, {"fin_store_code"}, "m", JoinKind.LeftOuter),
    Exp    = Table.ExpandTableColumn(Merged, "m", {"store_key", "valid_from", "valid_to"}),
    Valid  = Table.SelectRows(Exp, each [posting_date] >= [valid_from] and [posting_date] <= [valid_to]),
    Out    = Table.RemoveColumns(Valid, {"valid_from", "valid_to"})
in
    Out""")
P(f"Kiểm tra: fact_finance có đúng {nf(fin_rows)} dòng như finance_raw. Các dòng HQ-ADJ có store_key trống; giữ lại để "
  "giải thích ở tổng chuỗi, không gán cho cửa hàng nào.")
H3("2.4 Số dòng cần khớp sau khi nạp")
TABLE(["Query", "Số dòng"], [
    ["fact_sales_line (sau Remove duplicates)", nf(rows["fact_sales_line_clean"])],
    ["fact_return", nf(rows["fact_return"])],
    ["fact_finance", nf(fin_rows)],
    ["dim_date", nf(rows["dim_date"])], ["dim_store", nf(rows["dim_store"])], ["dim_product", nf(rows["dim_product"])],
    ["map_store_finance", nf(rows["map_store_finance"])],
], widths=[12, 5.5], align="lr")
P("Sai khác số dòng thường do quên bỏ trùng, chọn nhầm sheet README, hoặc merge Finance mà không lọc hiệu lực "
  "(fact_finance dư dòng ở cửa hàng 38).")

# ---------------------------------------------------------------- Bước 3
H2("Bước 3. Xây dựng mô hình dữ liệu")
P("Mô hình hình sao: fact ở giữa, dim xung quanh, quan hệ 1 → nhiều và lọc một chiều. "
  "fact_sales_line và fact_finance không nối trực tiếp với nhau; chúng gặp nhau qua dim_date và dim_store.", first=True)
FIG("02_mo_hinh_du_lieu.png", "Hình 3: Mô hình dữ liệu (nguồn: diagrams/02_mo_hinh_du_lieu.drawio)")
H3("3.1 Quan hệ cần tạo")
TABLE(["Từ (1)", "Đến (nhiều)", "Trạng thái"], [
    ["dim_date[date]", "fact_sales_line[order_date]", "Active"],
    ["dim_date[date]", "fact_sales_line[delivery_date]", "Inactive (dùng USERELATIONSHIP)"],
    ["dim_date[date]", "fact_return[return_date]", "Active"],
    ["dim_date[date]", "fact_finance[posting_date]", "Active"],
    ["dim_store[store_key]", "fact_sales_line[store_key]", "Active"],
    ["dim_store[store_key]", "fact_finance[store_key]", "Active"],
    ["dim_product[product_key]", "fact_sales_line[product_key]", "Active"],
    ["dim_promotion[promo_code]", "fact_sales_line[promo_code]", "Active"],
    ["dim_order_status[order_status]", "fact_sales_line[order_status]", "Active"],
    ["dim_channel[channel_code]", "dim_store[channel_code]", "Active (kênh lọc qua cửa hàng)"],
], widths=[5.5, 6.5, 5.5])
P("fact_return không có store_key: thêm cột store_key cho fact_return trong Power Query bằng cách merge với "
  "fact_sales_line theo order_number + line_item, rồi tạo quan hệ dim_store → fact_return. Đây là cách để lọc hoàn trả theo cửa hàng.")
H3("3.2 Cấu hình bắt buộc")
LIST(["**Đánh dấu bảng ngày.** Chọn dim_date → Mark as date table → cột date.",
      "**Sort by column.** dim_date[month_label] sort theo [month_id].",
      "**Ẩn cột kỹ thuật.** Ẩn các khoá ở bảng fact (store_key, product_key, …) để học viên luôn kéo cột từ dim.",
      "**Bảng _Measures.** Mọi measure đặt trong một bảng rỗng, gom theo Display folder ở Bước 4."])
H3("3.3 Ba phép thử mô hình")
P("Chạy trong DAX query view. Nếu ra đúng số, mô hình đã đúng:", first=True)
CODE(f"""EVALUATE
ROW (
    "Đơn COMPLETED 2020", CALCULATE ( DISTINCTCOUNT ( fact_sales_line[order_number] ),
                            fact_sales_line[order_status] = "COMPLETED", dim_date[year] = 2020 ),   -- {nf(ytd['orders'])}
    "Finance 2020",       CALCULATE ( SUM ( fact_finance[net_revenue] ), dim_date[year] = 2020 ), -- {nf(R['finance_year']['2020'], 2)}
    "Dòng Finance online", CALCULATE ( COUNTROWS ( fact_finance ), dim_store[channel_code] = "ONLINE" )
)""")

# ---------------------------------------------------------------- Bước 4
H2("Bước 4. Viết measure DAX")
P("Viết theo đúng thứ tự cây KPI: doanh thu → đơn hàng → Finance → chênh lệch → so sánh kỳ. Mỗi measure mới kiểm "
  "ngay bằng một card trước khi dùng vào measure khác.", first=True)
H3("4.1 Nguyên tắc chung")
LIST(["Dùng DIVIDE thay cho phép chia “/” để tránh lỗi chia cho 0.",
      "Bộ lọc trạng thái đặt **trong** measure (KEEPFILTERS), không đặt ở filter pane, để mọi trang dùng chung một định nghĩa.",
      "Format: tiền `#,0` USD; tỷ lệ `0.00%`; chênh lệch `+#,0;−#,0;0`.",
      "Display folder: 1 Doanh thu · 2 Đơn hàng · 3 Finance & đối soát · 4 So sánh kỳ · 5 Hoàn trả & chất lượng."])
H3("4.2 Nhóm Doanh thu")
CODE("""Gross Sales =
CALCULATE (
    SUMX ( fact_sales_line, fact_sales_line[quantity] * fact_sales_line[unit_price] ),
    KEEPFILTERS ( fact_sales_line[order_status] = "COMPLETED" )
)

Discounts =
CALCULATE ( SUM ( fact_sales_line[discount_amount] ), KEEPFILTERS ( fact_sales_line[order_status] = "COMPLETED" ) )

Net Sales = [Gross Sales] - [Discounts]

Returns = SUM ( fact_return[return_amount] )

Net Revenue = [Net Sales] - [Returns]""")
H3("4.3 Nhóm Đơn hàng")
CODE("""Orders =
CALCULATE ( DISTINCTCOUNT ( fact_sales_line[order_number] ), KEEPFILTERS ( fact_sales_line[order_status] = "COMPLETED" ) )

Units =
CALCULATE ( SUM ( fact_sales_line[quantity] ), KEEPFILTERS ( fact_sales_line[order_status] = "COMPLETED" ) )

AOV = DIVIDE ( [Net Sales], [Orders] )
ASP = DIVIDE ( [Net Sales], [Units] )""")
H3("4.4 Nhóm Finance và đối soát")
CODE(f"""Finance Net Revenue =
CALCULATE ( SUM ( fact_finance[net_revenue] ), NOT ISBLANK ( fact_finance[store_key] ) )   -- bỏ HQ-ADJ

HQ Adjustment = CALCULATE ( SUM ( fact_finance[net_revenue] ), ISBLANK ( fact_finance[store_key] ) )

Δ Finance = [Finance Net Revenue] - [Net Revenue]
Δ Finance % = DIVIDE ( [Δ Finance], ABS ( [Finance Net Revenue] ) )

Net Sales (ngày giao) =                       -- online tính theo delivery_date như Finance
VAR pos = CALCULATE ( [Net Sales], dim_store[channel_code] = "POS" )
VAR onl = CALCULATE ( [Net Sales], dim_store[channel_code] = "ONLINE",
                      USERELATIONSHIP ( dim_date[date], fact_sales_line[delivery_date] ) )
RETURN pos + onl

Δ sau quy đổi = [Finance Net Revenue] - ( [Net Sales (ngày giao)] - [Returns] )

Vượt ngưỡng =                                 -- 1 nếu ô cửa hàng × tháng vượt ngưỡng
VAR d = [Δ Finance]
RETURN IF ( ABS ( d ) > {nf(thr['abs_usd'])} && ABS ( [Δ Finance %] ) > {str(thr['pct']).replace('.', ',')} , 1, 0 )

Số cửa hàng vượt ngưỡng =
SUMX ( VALUES ( dim_store[store_key] ), [Vượt ngưỡng] )""")
P("Lưu ý: [Số cửa hàng vượt ngưỡng] đúng khi bối cảnh lọc là một tháng; ở mức năm, đếm theo từng cặp tháng × cửa hàng "
  "bằng SUMX ( SUMMARIZE ( dim_date, dim_date[month_id] ), SUMX ( VALUES ( dim_store[store_key] ), [Vượt ngưỡng] ) ).")
H3("4.5 Nhóm So sánh kỳ")
CODE("""Net Revenue LM   = CALCULATE ( [Net Revenue], DATEADD ( dim_date[date], -1, MONTH ) )
Net Revenue SPLY = CALCULATE ( [Net Revenue], SAMEPERIODLASTYEAR ( dim_date[date] ) )
Net Revenue Δ% LM   = DIVIDE ( [Net Revenue] - [Net Revenue LM], [Net Revenue LM] )
Net Revenue Δ% SPLY = DIVIDE ( [Net Revenue] - [Net Revenue SPLY], [Net Revenue SPLY] )
Nhãn Δ% LM =
VAR x = [Net Revenue Δ% LM]
RETURN IF ( ISBLANK ( x ), "–", IF ( x >= 0, "▲ ", "▼ " ) & FORMAT ( ABS ( x ), "0.0%" ) & " so với tháng trước" )""")
P("Lặp lại cùng mẫu cho Gross Sales, Orders, AOV, ASP, Finance Net Revenue. Với tỷ lệ (Return Rate) so sánh bằng điểm %, không dùng Δ%.")
H3("4.6 Nhóm Hoàn trả và chất lượng")
CODE("""Return Rate = DIVIDE ( [Returns], [Net Sales] )
Cancelled Orders = CALCULATE ( DISTINCTCOUNT ( fact_sales_line[order_number] ), fact_sales_line[order_status] = "CANCELLED" )
Pending Orders   = CALCULATE ( DISTINCTCOUNT ( fact_sales_line[order_number] ), fact_sales_line[order_status] = "PENDING" )""")
H3("4.7 Năm lỗi hay gặp")
LIST(["**Quên lọc COMPLETED** → Net Sales cao hơn, AOV sai; Δ âm ở mọi cửa hàng.",
      "**Đếm Orders bằng COUNTROWS** → đếm dòng chứ không đếm đơn, AOV thấp bất thường (~1/2,4 giá trị đúng).",
      "**Trừ Returns theo ngày đặt của đơn gốc** → Returns không khớp tháng với Finance.",
      "**So với total_incl_vat** → Δ ≈ 10 % ở mọi ô.",
      "**Merge Finance chỉ theo mã** → cửa hàng 38 mất doanh thu từ 07/2019 (mã GBR-03N không map) hoặc bị nhân đôi."])

# ---------------------------------------------------------------- Bước 5
H2("Bước 5. Dựng 4 trang dashboard")
H3("5.1 Nguyên tắc trình bày: Storytelling with Data")
P("Dashboard dùng để ra quyết định trước cuộc họp chốt số, nên áp dụng 6 bài học của Cole Nussbaumer Knaflic "
  "(Storytelling with Data). Bản HTML và Excel tham chiếu đều làm theo các quy tắc dưới đây; khi chấm phần trải "
  "nghiệm (6 điểm) và lập luận, đối chiếu với bảng này.", first=True)
TABLE(["Bài học", "Áp dụng trong dashboard", "Lỗi hay gặp cần tránh"], [
    ["1. Hiểu bối cảnh", "Mỗi trang mở bằng câu “Điểm chính” (Big Idea) tự sinh theo số: chuyện gì xảy ra, vì sao, cần làm gì.", "Trang chỉ có biểu đồ, người xem phải tự đoán kết luận"],
    ["2. Chọn hình phù hợp", "Đường cho xu hướng; cột ngang xếp hạng cho danh mục/quốc gia; waterfall cho cầu doanh thu và phân rã Δ; bảng cho chi tiết.", "Pie/donut/treemap để so sánh, biểu đồ hai trục y, 3D"],
    ["3. Gỡ rối", "Card phẳng không nền màu, lưới mờ, nhãn trực tiếp ở cuối thanh/điểm, bỏ chú giải khi chỉ có một chuỗi.", "Khung đậm, nền màu ở mọi tiêu đề, số trên mọi điểm"],
    ["4. Dẫn mắt", "Xám cho bối cảnh, một màu nhấn (xanh dương) cho tháng chọn / mục lớn nhất; đỏ chỉ cho ⚠ vượt ngưỡng; ▲▼ luôn kèm chữ.", "Mỗi chuỗi một màu, đỏ/xanh lá cho mọi thay đổi"],
    ["5. Nghĩ như nhà thiết kế", "Lưới thẳng hàng, định dạng số thống nhất (nghìn USD 1 chữ số lẻ), bảng màu an toàn cho người mù màu, tooltip.", "Số lẻ khác nhau giữa các visual, chữ bị cắt"],
    ["6. Kể chuyện", "Tiêu đề hành động (“Máy tính chiếm 43% Net Sales”), mạch trang: Tổng quan → Doanh thu đến từ đâu → Đối soát → Hoàn trả; chân trang ghi nguồn, as-of, quy ước Δ.", "Tiêu đề chỉ mô tả (“Cơ cấu theo danh mục”)"],
], widths=[3.2, 8.8, 5.5])
P("Trong Power BI: tiêu đề hành động làm bằng measure trả về chuỗi (Format → Title → fx → chọn measure), ví dụ "
  "`Tiêu đề danh mục = MAXX(TOPN(1, VALUES(dim_product[category]), [Net Sales]), dim_product[category]) & \" chiếm \" "
  "& FORMAT(DIVIDE(MAXX(TOPN(1, VALUES(dim_product[category]), [Net Sales]), [Net Sales]), [Net Sales]), \"0%\") & \" Net Sales\"`.")
H3("5.2 Khung chung")
LIST(["Canvas 1280 × 720, nền xám rất nhạt #F6F6F4, card trắng không viền. Header trắng: tiêu đề + dòng phụ ghi ngày dữ liệu.",
      "3 slicer đặt trong header, đồng bộ cho cả 4 trang (View → Sync slicers): Tháng báo cáo (dropdown, chọn một), "
      "Kênh bán (Cửa hàng / Website-app), Quốc gia. Tab trang đặt ngay dưới header.",
      "Dải “Điểm chính” (text box hoặc card dùng measure chuỗi) ngay dưới tab, viền trái màu nhấn.",
      "Card KPI: nhãn xám, giá trị đen 28 px, 2 dòng so sánh LM và SPLY có ▲▼; tốt = xanh dương #1C5CAB, xấu = cam #EB6834; "
      "Δ Finance và cờ vượt ngưỡng dùng đỏ #D03B3B kèm ký hiệu ⚠.",
      "Màu chuỗi: xám #B9B7B0 cho bối cảnh, xanh dương #2A78D6 cho điểm nhấn. Font Segoe UI.",
      "Mở Nexa_Sales_P1_Dashboard.html để xem bố cục và màu mẫu."])
VIS = {
    "5.3 Trang Tổng quan": [
        ["Text / card", "Điểm chính", "Measure chuỗi: Net Revenue, % so với LM/SPLY, nguyên nhân chính (số đơn hay AOV), Δ Finance", "16, 64, 1248, 70"],
        ["Card × 5", "KPI", "Net Revenue · Orders · AOV · Return Rate · Δ Finance (kèm số cửa hàng vượt ngưỡng)", "16, 142, 1248, 96"],
        ["Line chart", "Xu hướng", "Net Revenue 24 tháng; màu xám, tháng chọn tô nhấn, ghi chú điểm SPLY", "16, 246, 828, 230"],
        ["Waterfall", "Cầu doanh thu", "Gross Sales → −Discounts → −Returns → Net Revenue", "852, 246, 412, 230"],
        ["Line chart × 2", "Động lực", "Orders (nhấn) và AOV (xám) là 2 visual riêng, cùng trục thời gian — không dùng 2 trục y", "16, 484, 828, 220"],
        ["Column chart", "Đối soát", "Δ sau quy đổi theo tháng; tháng có cửa hàng vượt ngưỡng tô đỏ (conditional format)", "852, 484, 412, 220"]],
    "5.4 Trang Doanh thu đến từ đâu": [
        ["Text / card", "Điểm chính", "Danh mục dẫn đầu, thị trường lớn nhất, tỷ trọng online", "16, 64, 1248, 70"],
        ["Card × 4", "KPI", "Net Sales · Units · Tỷ trọng online · Giảm giá / Gross", "16, 142, 1248, 96"],
        ["Bar (ngang)", "Danh mục", "Net Sales theo category, sắp giảm dần; mục lớn nhất màu nhấn, còn lại xám; nhãn giá trị + %", "16, 246, 408, 300"],
        ["Bar (ngang)", "Quốc gia", "Net Sales theo country (Online là một dòng)", "432, 246, 408, 300"],
        ["Table", "Kênh", "channel_name; Net Sales, Orders, AOV, Δ% LM, Δ% SPLY", "848, 246, 416, 300"],
        ["Table", "Xếp hạng", "Top 10 cửa hàng; data bar trong cột Net Sales; Δ% LM, Δ% SPLY", "16, 554, 1248, 150"]],
    "5.5 Trang Đối soát với Finance": [
        ["Text / card", "Điểm chính", "Δ toàn chuỗi; phần lệch thời điểm online; cửa hàng cần điều tra", "16, 64, 1248, 70"],
        ["Card × 5", "KPI", "Net Revenue đơn · Finance · Δ · Δ sau quy đổi · Số cửa hàng vượt ngưỡng", "16, 142, 1248, 96"],
        ["Waterfall", "Phân rã Δ", "Online lệch thời điểm → từng cửa hàng ⚠ → còn lại → Δ", "16, 246, 408, 260"],
        ["Column chart", "Theo tháng", "Δ sau quy đổi; đỏ = tháng có cửa hàng vượt ngưỡng", "432, 246, 832, 260"],
        ["Table", "Chi tiết", "store_code; Net Revenue, Finance, Δ, Δ%, Δ sau quy đổi, ⚠ Vượt (conditional format); sắp theo |Δ|", "16, 514, 828, 190"],
        ["Line chart", "Online", "Δ online gốc (xám) và sau quy đổi (nhấn)", "852, 514, 412, 190"]],
    "5.6 Trang Hoàn trả & chất lượng dữ liệu": [
        ["Text / card", "Điểm chính", "Hoàn trả, tỷ lệ, lý do lớn nhất, dòng trùng và PENDING đã loại", "16, 64, 1248, 70"],
        ["Card × 4", "KPI", "Returns · Return Rate (điểm %) · Cancelled Orders · Dòng trùng đã loại + Pending Orders", "16, 142, 1248, 96"],
        ["Line chart", "Xu hướng", "Return Rate 24 tháng, tháng chọn tô nhấn", "16, 246, 828, 220"],
        ["Bar (ngang)", "Lý do", "Returns theo return_reason, lý do lớn nhất tô nhấn", "852, 246, 412, 220"],
        ["Bar (ngang)", "Danh mục", "Return Rate theo category, nhãn kèm Net Sales", "16, 474, 828, 230"],
        ["Line chart", "Xu hướng", "Cancelled Orders (nhấn) và Pending Orders (xám) theo tháng", "852, 474, 412, 230"]],
}
for title, rows_ in VIS.items():
    H3(title)
    TABLE(["Loại visual", "Khối", "Trường / measure", "Vị trí x, y, rộng, cao (px)"], rows_, widths=[3, 2.2, 8.5, 3.8])
H3("5.7 Tương tác và hoàn thiện")
LIST(["**Edit interactions.** Biểu đồ xu hướng không bị lọc bởi slicer Tháng (để luôn thấy 24 tháng); dùng measure "
      "có ALL(dim_date) + lọc cửa sổ 24 tháng, hoặc tắt tương tác slicer Tháng với visual đó.",
      "**Drill-through.** Từ bảng đối soát cửa hàng sang một trang chi tiết theo ngày (posting_date) để dò nguyên nhân.",
      "**Tooltip.** Trên bảng đối soát, tooltip hiện Net Sales, Returns, Finance revenue_recognized và refunds tách riêng.",
      "**Tiêu đề và đơn vị.** Tiêu đề visual nêu kết luận; dòng phụ ghi đơn vị (USD, nghìn USD, %, điểm %) và phạm vi.",
      "**Chân trang.** Ghi nguồn, ngày dữ liệu, quy ước Δ và ngưỡng bằng chữ xám nhỏ."])

# ---------------------------------------------------------------- Bước 6
H2("Bước 6. Đối soát số (DÀNH CHO MENTOR — ẩn khi giao đề)")
P(f"Các số dưới đây tính độc lập bằng pandas từ chính file giao học viên và đã khớp với dashboard tham chiếu. "
  f"Tháng báo cáo {mstr(R['report_month'])}, đủ bộ lọc “Tất cả”.", first=True)
H3(f"6.1 KPI tháng {mstr(R['report_month'])}")
def chg(a, b): return pct((a - b) / abs(b)) if b else "–"
krow = lambda name, k, d=0: [name, nf(cur[k], d), nf(lm[k], d), chg(cur[k], lm[k]), nf(sp[k], d), chg(cur[k], sp[k])]
TABLE(["Chỉ tiêu", "12/2020", "11/2020 (LM)", "Δ% LM", "12/2019 (SPLY)", "Δ% SPLY"], [
    krow("Gross Sales (USD)", "gross_sales", 2), krow("Discounts (USD)", "discounts", 2), krow("Net Sales (USD)", "net_sales", 2),
    krow("Returns (USD)", "returns", 2), krow("Net Revenue (USD)", "net_revenue", 2), krow("Orders", "orders"),
    krow("Units", "units"), krow("AOV (USD/đơn)", "aov", 2), krow("ASP (USD/sp)", "asp", 2),
    ["Return Rate", pct(cur["return_rate"], 2), pct(lm["return_rate"], 2), nf((cur["return_rate"] - lm["return_rate"]) * 100, 2) + " điểm %",
     pct(sp["return_rate"], 2), nf((cur["return_rate"] - sp["return_rate"]) * 100, 2) + " điểm %"],
], widths=[3.6, 2.9, 2.9, 1.9, 2.9, 2.0], align="lrrrrr")
fin12 = R["finance_month"][str(R["report_month"])]
P(f"Finance Net Revenue tháng 12/2020 (gồm cả HQ-ADJ {nf(hq20, 2)} USD): {nf(fin12, 2)} USD; chưa gồm HQ-ADJ: "
  f"{nf(fin12 - hq20, 2)} USD. Δ Finance (không gồm HQ-ADJ) = {nf(fin12 - hq20 - cur['net_revenue'], 2)} USD.")
H3("6.2 Lũy kế năm 2020 so với 2019")
yrow = lambda name, k, d=0: [name, nf(ytd[k], d), nf(ytd_sp[k], d), chg(ytd[k], ytd_sp[k])]
TABLE(["Chỉ tiêu", "2020", "2019", "Δ%"], [
    yrow("Gross Sales (USD)", "gross_sales", 2), yrow("Net Sales (USD)", "net_sales", 2), yrow("Returns (USD)", "returns", 2),
    yrow("Net Revenue (USD)", "net_revenue", 2), yrow("Orders", "orders"), yrow("AOV (USD/đơn)", "aov", 2),
    ["Finance Net Revenue (USD, gồm HQ-ADJ)", nf(R["finance_year"]["2020"], 2), nf(R["finance_year"]["2019"], 2),
     chg(R["finance_year"]["2020"], R["finance_year"]["2019"])],
], widths=[6.5, 4, 4, 3], align="lrrr")
H3("6.3 Cầu nối năm 2020: từ dữ liệu đơn đến Finance")
FIG("03_quy_trinh_doi_soat.png", "Hình 4: Quy trình đối soát (nguồn: diagrams/03_quy_trinh_doi_soat.drawio)")
bridge_rows = [["Net Revenue theo dữ liệu đơn (order_date, return_date)", nf(br["orders_net_revenue"], 2)]]
for k, name in CAUSE_VI.items():
    bridge_rows.append([("+ " if br[k] >= 0 else "− ") + name, nf(br[k], 2)])
bridge_rows.append(["= Finance Net Revenue các cửa hàng", nf(br["finance_net"], 2)])
bridge_rows.append(["+ Điều chỉnh HQ-ADJ (31/12/2020)", nf(hq20, 2)])
bridge_rows.append(["= Tổng Finance năm 2020", nf(br["finance_net"] + hq20, 2)])
TABLE(["Khoản", "USD"], bridge_rows, widths=[13, 4.5], align="lr")
P(f"Phần “Online lệch thời điểm” cộng cả năm còn {nf(br['adj_online_timing'], 2)} USD vì đơn đặt cuối 2019 giao đầu 2020 "
  f"nhiều hơn đơn đặt cuối 2020 giao đầu 2021. Đơn huỷ vẫn ghi nhận: {n_canc_2020} đơn trong năm 2020. "
  "Còn lại chưa giải thích = 0.")
H3("6.4 Các ô tháng × cửa hàng vượt ngưỡng năm 2020")
P(f"Có {len(over20)} ô vượt ngưỡng; học viên đạt yêu cầu khi giải thích đúng nguyên nhân chính của mọi ô, "
  "kèm số tiền.", first=True)
TABLE(["Tháng", "Cửa hàng", "Quốc gia", "Đơn hàng", "Finance", "Δ", "Nguyên nhân chính"],
      [[mstr(int(r.month_id)), r.store_code, r.country, nf(r.orders_net_revenue, 0), nf(r.finance_net, 0),
        nf(r.diff_finance_minus_orders, 0), r.cause] for r in over20.itertuples()],
      widths=[2.0, 1.8, 2.7, 2.3, 2.3, 2.1, 4.3], align="lllrrrl")
H3("6.5 Chênh lệch cài sẵn phía Finance (toàn bộ kỳ)")
grp = anom.groupby("type").agg(n=("anomaly_id", "size"), usd=("effect_usd", "sum")).reset_index()
TVI = {"A_CANCELLED_BOOKED": "Đơn huỷ vẫn ghi nhận (K01–K25)", "B_MISSING_UPLOAD": "Finance thiếu 3 ngày (K26)",
       "C_VAT_INCLUDED": "Doanh thu gồm VAT (K27)", "D_DUPLICATE_POSTING": "Bút toán ghi trùng (K28)",
       "E_MANUAL_ADJ": "Điều chỉnh HQ-ADJ (K29–K31)"}
detail = {t: "; ".join(anom[anom.type == t].detail.head(1)) for t in TVI}
TABLE(["Loại", "Số mục", "Ảnh hưởng (USD)", "Mô tả / cách phát hiện"],
      [[TVI[r.type], str(r.n), nf(r.usd, 2), detail[r.type]] for r in grp.itertuples()],
      widths=[4.5, 1.5, 2.8, 8.7], align="lrrl")
P("Nguyên nhân hệ thống (không phải lỗi, nhưng tạo Δ theo ngày/tháng): online ghi nhận theo ngày giao, hoàn tiền hạch "
  "toán trễ 0–3 ngày. Bẫy phía dữ liệu đơn: không bỏ 120 dòng trùng thì Net Sales cao hơn "
  f"{nf(R['systemic']['S3_DUPLICATE_LINES'], 2)} USD (toàn kỳ); tính nhầm đơn PENDING thì cao hơn "
  f"{nf(R['systemic']['S4_PENDING'], 2)} USD.")
H3("6.6 Checklist kiểm tra hành vi")
LIST(["Chọn tháng 06/2020, quốc gia Mỹ: 2 cửa hàng vượt ngưỡng (S54 do VAT, S45 do đơn huỷ được ghi nhận).",
      "Chọn kênh online: “Δ sau quy đổi” gần 0 ở mọi tháng, khác hẳn “Δ gốc”.",
      "Chọn tháng 12/2020: S12 (Pháp) lệch +10.141 USD, đúng bằng một dòng Finance ngày 10/12/2020 bị ghi 2 lần.",
      "Chọn tháng 02/2020: S24 (Đức) Finance thấp hơn 22.169 USD; xuống mức ngày thấy 3 ngày Finance trống.",
      "Bỏ bước Remove duplicates: Net Sales toàn kỳ tăng đúng bằng giá trị 120 dòng trùng.",
      "Đổi slicer Kênh sang “Cửa hàng” rồi chọn Quốc gia: biểu đồ online hiện thông báo không áp dụng, không báo lỗi."])

# ---------------------------------------------------------------- Bước 7
H2("Bước 7. Từ dashboard đến báo cáo")
H3("7.1 Mạch viết")
P("Mỗi phần báo cáo đi theo mạch **Câu hỏi → Phân tích → Bằng chứng (hình, số) → Kết luận**. Báo cáo theo đúng 8 phần "
  "của guild_report_template.pdf; phần IV (Main Analysis) dành cho đối soát.", first=True)
H3("7.2 Câu hỏi gợi ý theo trang")
TABLE(["Trang", "Câu hỏi"], [
    ["Tổng quan", "Tháng 12/2020 doanh nghiệp thực sự bán được bao nhiêu? Vì sao Net Revenue giảm mạnh so với 12/2019?"],
    ["Doanh thu", "Kênh nào, danh mục nào, quốc gia nào đóng góp chính? AOV online khác cửa hàng thế nào?"],
    ["Đối soát", "Chênh lệch với Finance là bao nhiêu? Bao nhiêu phần là lệch thời điểm, bao nhiêu là lỗi thật?"],
    ["Hoàn trả & chất lượng", "Tỷ lệ hoàn trả có bất thường ở danh mục nào? Dữ liệu đơn có vấn đề chất lượng gì?"],
], widths=[3.5, 14])
H3("7.3 Finding, Insight, Recommendation, Limitation")
LIST([f"**Finding** (số liệu): “Năm 2020 Finance ghi nhận cao hơn dữ liệu đơn {nf(br['finance_net'] - br['orders_net_revenue'], 0)} USD "
      f"({pct((br['finance_net'] - br['orders_net_revenue']) / br['finance_net'], 2)}).”",
      "**Insight** (ý nghĩa): “Phần lớn chênh lệch đến từ kênh online do khác mốc ghi nhận; sau khi quy về ngày giao, "
      "chênh lệch còn lại tập trung ở 4 lỗi vận hành có thể sửa.”",
      "**Recommendation** (hành động): “Đồng bộ trạng thái huỷ đơn sang kế toán trong ngày; kiểm tra trùng bút toán trước khi khoá sổ.”",
      "**Limitation** (giới hạn): “Dữ liệu là giả lập trên nền GER; VAT giả định 10 % thống nhất; chưa xét tỷ giá.”"])
H3("7.4 Gợi ý sáng tạo (20 điểm)")
LIST(["Trang chi tiết theo ngày cho một cửa hàng, có drill-through từ bảng đối soát.",
      "Waterfall cầu nối Orders → Finance theo nguyên nhân (dựng bảng nguyên nhân bằng DAX hoặc Power Query).",
      "Phân tích khuyến mãi: Net Sales và AOV theo promo_code; giảm giá có kéo số đơn lên không.",
      "Quy đổi về tiền tệ thanh toán bằng fx_rate; phân tích theo vùng khách hàng bằng dim_customer."])
H3("7.5 Sản phẩm nộp")
LIST(["File .pbix (4 trang + trang sáng tạo nếu có).",
      "Báo cáo PDF 8 phần, có bảng giải thích chênh lệch vượt ngưỡng năm 2020 (tháng, cửa hàng, Δ, nguyên nhân, số tiền).",
      "Phụ lục: danh mục measure, các bước Power Query, ảnh chụp từng trang dashboard."])

# ---------------------------------------------------------------- Phụ lục
H2("Phụ lục A. Danh mục measure")
TABLE(["Folder", "Measure"], [
    ["1 Doanh thu", "Gross Sales · Discounts · Net Sales · Returns · Net Revenue · Net Sales (ngày giao)"],
    ["2 Đơn hàng", "Orders · Units · AOV · ASP"],
    ["3 Finance & đối soát", "Finance Net Revenue · HQ Adjustment · Δ Finance · Δ Finance % · Δ sau quy đổi · Vượt ngưỡng · Số cửa hàng vượt ngưỡng"],
    ["4 So sánh kỳ", "<measure> LM · <measure> SPLY · Δ% LM · Δ% SPLY · Nhãn Δ% (cho Gross Sales, Net Revenue, Orders, AOV, ASP, Finance)"],
    ["5 Hoàn trả & chất lượng", "Return Rate · Return Rate LM/SPLY (điểm %) · Cancelled Orders · Pending Orders"],
], widths=[4, 13.5])
H2("Phụ lục B. Khi cần sinh lại dữ liệu hoặc sửa sơ đồ")
LIST(["Tải Global_Electronics_Retailer.zip từ thư mục Drive của project vào source/ (file .zip không đưa lên git).",
      "Chạy lần lượt trong _mentor/: generate_dataset.py → build_dashboard.py → build_dashboard_xlsx.py → chup_dashboard.js → build_diagrams.py → build_guideline.py → build_slides.py. "
      "Seed cố định nên số liệu không đổi.",
      "Sửa sơ đồ: mở diagrams/*.drawio bằng draw.io (app.diagrams.net hoặc bản desktop), sửa, lưu; rồi chạy "
      "`python build_diagrams.py --chi-xuat` để xuất lại PDF/PNG mà không ghi đè bản đã sửa tay; chạy lại build_guideline.py để cập nhật hình.",
      "Dashboard Power BI không cần sửa DAX khi dữ liệu đổi: chỉ cần Refresh."])

doc.core_properties.title = "Dashboard Guideline – Sales Project 1"
doc.core_properties.author = "EDA Mentor"
doc.save(OUT)
print("saved", OUT.name)
