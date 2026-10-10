# -*- coding: utf-8 -*-
"""Dựng ../Dashboard_Guideline (dành cho mentor).docx

Khuôn: guideline của Bank Project 1 (giữ nguyên style, khổ giấy, bảng màu) — xoá nội dung cũ, viết nội dung Sales.
Mọi con số đọc từ ref_numbers.json, Instructor Key và chính file dataset => chạy lại sau generate_dataset.py.
Hình: lấy PNG xuất từ các file .drawio trong ../diagrams (sửa sơ đồ ở .drawio, chạy build_diagrams.py rồi chạy lại file này).
"""
import copy
import json
import subprocess
import sys
from pathlib import Path

import docx
import pandas as pd
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from pq_thao_tac import DIMS, DOI_TEN, NAP_DATASET, ORDER, STEPS
from thuat_ngu import BANG, bat_dau, chuan_hoa

bat_dau()

HERE = Path(__file__).resolve().parent
P1 = HERE.parent
BANK = P1.parents[2] / "Bank" / "Project 1" / "Dashboard_Guideline (dành cho mentor).docx"
HOC_VIEN = "--hoc-vien" in sys.argv                  # bản học viên: bỏ Bước 6 (đáp án)
OUT = (P1 / "hoc_vien" / "Huong_dan_hoc_vien.docx") if HOC_VIEN else (P1 / "Dashboard_Guideline (dành cho mentor).docx")
PNG = P1 / "diagrams" / "png"
R = json.loads((HERE / "ref_numbers.json").read_text(encoding="utf-8"))
KEY = pd.read_excel(HERE / "Nexa_Sales_P1_Instructor_Key.xlsx", sheet_name=None)
HV = P1 / "hoc_vien"                                  # bộ file giao học viên
DS = pd.read_excel(HV / "Nexa_Sales_P1_Dataset.xlsx", sheet_name=None)


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


def runs(p, text, raw=False):
    """**đậm** và `code` đơn giản trong một đoạn."""
    import re
    for part in re.split(r"(\*\*[^*]+\*\*|`[^`]+`)", text):
        if not part:
            continue
        if part.startswith("**"):
            p.add_run(part[2:-2] if raw else chuan_hoa(part[2:-2])).bold = True
        elif part.startswith("`"):
            r = p.add_run(part[1:-1])
            r.font.name = "Consolas"
            r.font.size = Pt(9.5)
        else:
            p.add_run(part if raw else chuan_hoa(part))
    return p


def H1(t): doc.add_paragraph(chuan_hoa(t), style=STY["Heading 1"])
def H2(t): doc.add_paragraph(chuan_hoa(t), style=STY["Heading 2"])
def H3(t): doc.add_paragraph(chuan_hoa(t), style=STY["Heading 3"])
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


def TABLE(head, rows, widths=None, align=None, raw_cols=()):
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
                r = p.add_run(chuan_hoa(str(val))); r.bold = True; r.font.color.rgb = RGBColor(255, 255, 255); r.font.size = Pt(8.5)
                shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), "447A2A")
                cell._tc.get_or_add_tcPr().append(shd)
            else:
                runs(p, str(val), raw=ci in raw_cols)
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
    doc.add_paragraph(chuan_hoa(caption), style=STY["Image Caption"])


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
PQ = json.loads((HERE / "pq_steps.json").read_text(encoding="utf-8"))
KIT = HERE / "excel_build" if (HERE / "excel_build").exists() else HV / "excel_build"   # DAX gốc (chỉ mentor)
DAX_TXT = (KIT / "measures.dax").read_text(encoding="utf-8")


def dax_sections(txt):
    """[(tiêu đề nhóm, [dòng code measure…])] theo các dòng // ===== trong measures.dax"""
    out, cur_ = [], None
    for line in txt.splitlines():
        if line.startswith("// ====="):
            cur_ = (line[8:].strip(), []); out.append(cur_)
        elif cur_ is not None and line.strip():
            cur_[1].append(line)
    return out


DAXS = dax_sections(DAX_TXT)
DAXS_NC = dax_sections((KIT / "measures_nang_cao.dax").read_text(encoding="utf-8"))
_fin = pd.read_excel(HV / "finance_revenue.xlsx", sheet_name="finance_revenue")
fin_online = int(_fin.fin_store_code.isin(DS["map_store_finance"].query("store_key == 0").fin_store_code).sum())
ten_so_do = lambda secs: [l.split(":=")[0].strip() for _, ls in secs for l in ls if ":=" in l and not l.startswith(" ")]
MEAS, MEAS_NC = ten_so_do(DAXS), ten_so_do(DAXS_NC)
PL = dict(so_do="A", thuat_ngu="B")
def chg(a, b): return pct((a - b) / abs(b)) if b else "–"
def rr_(k): return k["returns"] / k["net_sales"]     # tỷ lệ chưa làm tròn

# ------------------------------------------------------------------ NỘI DUNG
H1(("Hướng dẫn học viên" if HOC_VIEN else "Hướng dẫn xây dựng") + " Dashboard Excel: Kết quả bán hàng và đối soát doanh thu "
   "(Nexa Home · Sales Project 1)")
P("Cập nhật: 10/10/2026 · Dataset v" + R["version"], first=True)

H2("Giới thiệu và cách dùng tài liệu")
P("Tài liệu này hướng dẫn dựng dashboard **Excel** 4 sheet (Tổng quan, Doanh thu, Đối soát, Hoàn trả & chất lượng) "
  "từ dataset Nexa Home, rồi dùng chính dashboard đó để đối soát (đối chiếu số liệu) doanh thu với số của phòng "
  "Tài chính – Kế toán (Finance, file finance_revenue.xlsx) và viết báo cáo. "
  "Quy trình gồm 4 tầng, làm hoàn toàn trên giao diện Excel: **Power Query** nạp và làm sạch → **Data Model** trong Power Pivot nối các bảng → "
  "**DAX** tính chỉ số hiệu suất then chốt (KPI) → **PivotTable, PivotChart, Slicer và CUBEVALUE** trình bày. "
  "Đây là project đầu tiên của nhóm Sales, mức hỗ trợ của mentor khoảng 80 %: mentor đi cùng học viên qua từng bước, "
  "học viên tự làm phần giải thích chênh lệch.", first=True)
if HOC_VIEN:
    P("**Đối tượng.** Học viên làm theo từng bước; mentor hỗ trợ trên lớp. Số kiểm tra và đáp án đối soát do mentor giữ.")
else:
    P("**Đối tượng.** Bản dành cho mentor và trợ giảng. **Bước 6** chứa số đối soát và đáp án các chênh lệch. Học viên "
      "nhận bản riêng `hoc_vien/Huong_dan_hoc_vien.docx`: Bước 6 chỉ có phương pháp (không có số và đáp án).")
P("**Sản phẩm đầu ra.** Một file .xlsx gồm các truy vấn Power Query, Mô hình Dữ liệu có quan hệ và bộ số đo DAX, cùng 4 sheet "
  "dashboard; kèm báo cáo 8 phần theo guild_report_template.pdf, trong đó phần chính là bảng giải thích chênh lệch "
  "vượt ngưỡng năm 2020.")

H3("Yêu cầu phần mềm")
LIST(["**Excel cho Windows**, bản Microsoft 365 hoặc Excel 2016 trở lên. Excel cho Mac và Excel trên web không có "
      "Power Pivot nên không làm được Bước 3–4.",
      "**Bật add-in Power Pivot.** File → Options → Add-ins → ô Manage chọn **COM Add-ins** → Go… → tick "
      "**Microsoft Power Pivot for Excel** → OK. Ribbon có thêm tab **Power Pivot**.",
      "Power Query có sẵn trong tab Data (nhóm Get & Transform Data), không cần cài thêm.",
      "Tên menu ghi theo giao diện tiếng Việt của Microsoft kèm tên tiếng Anh trong ngoặc, vì máy học viên có thể dùng một trong hai."])

H3("Bộ file và ai được xem")
FILES = [
    ["hoc_vien/README.docx", "Đề bài Project 1", "Học viên"],
    ["hoc_vien/Nexa_Sales_P1_Dataset.xlsx", "Dữ liệu đơn hàng, hàng bán bị trả lại, các bảng chiều, bảng map mã Kế toán, định nghĩa KPI, giao dịch mẫu; mỗi sheet là một Excel Table tên tbl_<sheet>", "Học viên"],
    ["hoc_vien/finance_revenue.xlsx", "Doanh thu phòng Tài chính – Kế toán đã ghi sổ, theo ngày × mã cửa hàng của Kế toán (bản rút gọn)", "Học viên"],
    ["hoc_vien/Huong_dan_hoc_vien.docx", "Hướng dẫn từng bước, bản học viên (không có số kiểm tra và đáp án đối soát)", "Học viên"],
    ["hoc_vien/Slide_Project_1_hoc_vien.pptx", "Slide giảng 8 phần, đã bỏ phụ lục đáp án và lời thuyết trình", "Học viên"],
    ["hoc_vien/so_do/*.pdf", "Sơ đồ pipeline, mô hình, quy trình đối soát, cây KPI, pipeline Power Query", "Học viên"],
    ["hoc_vien/guild_report_template.pdf", "Khung báo cáo 8 phần", "Học viên"],
]
if not HOC_VIEN:
    FILES += [
        ["Nexa_Sales_P1_Dashboard.html", "Dashboard tham chiếu (bố cục, màu, câu chữ); có thể gửi học viên sau buổi 1", "Mentor (học viên sau buổi 1)"],
        ["Nexa_Sales_P1_Dashboard.xlsx", "Dashboard Excel tham chiếu dựng đúng các bước của tài liệu này: Power Query, Mô hình Dữ liệu, số đo DAX, 4 slicer, PivotTable, CUBEVALUE", "Mentor (học viên sau buổi 1)"],
        ["Dashboard_Guideline (dành cho mentor).docx", "Tài liệu này, bản đầy đủ có Bước 6 đáp án", "Chỉ mentor"],
        ["Slide_Project_1.pptx", "Slide bản đầy đủ: 4 slide phụ lục đáp án (ẩn) + lời thuyết trình", "Chỉ mentor"],
        ["_mentor/Nexa_Sales_P1_Instructor_Key.xlsx", "Đáp án: 31 chênh lệch cài sẵn, cầu nối từng tháng × cửa hàng", "Chỉ mentor"],
    ]
    TABLE(["File", "Công dụng", "Ai được xem"], FILES, widths=[5.4, 8.8, 3.3])
else:
    TABLE(["File", "Công dụng"], [f[:2] for f in FILES], widths=[5.6, 11.9])

H3("Dữ liệu trong một cái nhìn")
TABLE(["Nội dung", "Giá trị"], [
    ["Bối cảnh", "Nexa Home bán điện máy qua cửa hàng (POS) và website/app; Trưởng phòng Kinh doanh và phòng Tài chính – Kế toán cần thống nhất doanh thu thật và lý do lệch số"],
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
TABLE(["Bước", "Việc chính", "Công cụ Excel", "Thời lượng gợi ý"], [
    ["1. Hiểu dữ liệu", "Đọc README, Data_Dictionary, KPI_Definitions, sample_transactions", "Mở file dataset", "30 phút"],
    ["2. Power Query", f"{len(ORDER)} query: nạp, đặt kiểu, bỏ dòng trùng, ghép bảng, gắn mã Kế toán → fact_finance", "Data → Get Data → From Excel Workbook; Merge Queries", "60 phút"],
    ["3. Data Model", "Quan hệ lược đồ hình sao, Mark as Date Table, quan hệ inactive cho delivery_date", "Power Pivot → Manage → Diagram View", "30 phút"],
    ["4. DAX", f"{len(MEAS)} số đo bắt buộc trong bảng _Measures ({len(MEAS_NC)} số đo nâng cao tuỳ chọn)", "Power Pivot → Measures → New Measure", "60 phút"],
    ["5. Dashboard", "4 sheet theo bảng visual → đối tượng Excel", "PivotTable, PivotChart, Slicer, CUBEVALUE", "120 phút"],
    ["6. Đối soát", ("Lập bảng giải thích chênh lệch năm 2020; mentor đối chiếu số" if HOC_VIEN else "So số với Bước 6; lập bảng giải thích chênh lệch năm 2020"), "Sheet Doi_soat", "60 phút"],
    ["7. Báo cáo", "Viết báo cáo 8 phần theo template", "—", "Ở nhà"],
], widths=[3.2, 7.3, 4.8, 2.2])

H3("Quy ước trong tài liệu")
LIST(["Tên bảng và cột viết dạng `fact_sales_line[order_date]`; tên measure viết dạng `[Net Revenue]`.",
      f"Mọi con số mẫu lấy từ chính dataset, tại tháng báo cáo {mstr(R['report_month'])} nếu không ghi khác.",
      "“LM” là tháng trước (Last Month), “SPLY” là cùng tháng năm trước (Same Period Last Year).",
      "Δ (chênh lệch) luôn tính **số Kế toán − dữ liệu đơn**: Δ dương nghĩa là Kế toán ghi nhận cao hơn.",
      "Tên file, cột và số đo giữ chữ `Finance` (vd `finance_revenue.xlsx`, `[Finance Net Revenue]`, `[Δ Finance]`): đều chỉ số của phòng Tài chính – Kế toán.",
      "Tên chỉ tiêu theo Chuẩn mực kế toán VAS 14: doanh thu bán hàng (Gross Sales) − chiết khấu thương mại (Discounts) − hàng bán bị trả lại (Returns) = doanh thu thuần (Net Revenue). Bảng thuật ngữ đầy đủ ở Phụ lục " + PL["thuat_ngu"] + ".",
      "Số viết theo kiểu Việt Nam: dấu chấm phân cách nghìn, dấu phẩy thập phân."])
FIG("01_pipeline_du_lieu.png", "Hình 1: Pipeline từ nguồn dữ liệu đến báo cáo")

# ---------------------------------------------------------------- Bước 1
H2("Bước 1. Hiểu dữ liệu và chỉ tiêu")
P("Dataset gồm 2 bảng dữ kiện (fact table), 7 bảng chiều (dimension table), 1 bảng map và file Finance riêng. Nắm grain (mỗi dòng là gì) của từng bảng "
  "trước khi viết bất kỳ measure nào.", first=True)
H3("1.1 Các bảng trong dataset")
cnt = {k: len(v) for k, v in DS.items()}
TABLE(["Bảng", "Loại bảng", "Độ chi tiết – grain (mỗi dòng là)", "Khoá", "Số dòng", "Nhóm"], [
    ["fact_sales_line", "Dữ kiện", "1 sản phẩm trong 1 đơn", "order_number + line_item", nf(cnt["fact_sales_line"]), "Bắt buộc"],
    ["fact_return", "Dữ kiện", "1 lần trả hàng của 1 dòng", "return_id", nf(cnt["fact_return"]), "Bắt buộc"],
    ["finance_revenue (file riêng)", "Dữ kiện", "1 ngày ghi sổ × 1 mã cửa hàng của Kế toán", "—", nf(fin_rows), "Bắt buộc"],
    ["dim_date", "Chiều", "1 ngày", "date", nf(cnt["dim_date"]), "Bắt buộc"],
    ["dim_store", "Chiều", "1 cửa hàng (0 = online)", "store_key", nf(cnt["dim_store"]), "Bắt buộc"],
    ["dim_product", "Chiều", "1 sản phẩm", "product_key", nf(cnt["dim_product"]), "Bắt buộc"],
    ["dim_channel", "Chiều", "1 kênh", "channel_code", nf(cnt["dim_channel"]), "Bắt buộc"],
    ["dim_promotion", "Chiều", "1 chương trình KM", "promo_code", nf(cnt["dim_promotion"]), "Bắt buộc"],
    ["dim_order_status", "Chiều", "1 trạng thái đơn", "order_status", nf(cnt["dim_order_status"]), "Bắt buộc"],
    ["map_store_finance", "Map", "1 mã Finance trong 1 khoảng hiệu lực", "fin_store_code", nf(cnt["map_store_finance"]), "Bắt buộc"],
    ["dim_customer", "Chiều", "1 khách hàng", "customer_key", nf(cnt["dim_customer"]), "Bổ sung"],
    ["fx_rate", "Tham khảo", "1 ngày × 1 tiền tệ", "date + currency_code", nf(cnt["fx_rate"]), "Bổ sung"],
], widths=[3.6, 1.6, 4.2, 3.6, 1.8, 1.8], align="lllllr")
P("Nhóm “Bắt buộc” đủ để dựng 4 trang và đối soát. Nhóm “Bổ sung” dành cho phần sáng tạo (20 điểm trong rubric), "
  "ví dụ phân tích theo vùng khách hàng hoặc quy đổi tiền tệ.")
H3("1.2 Từ chỉ tiêu nghiệp vụ đến cột dữ liệu")
kpi = DS["KPI_Definitions"]
TABLE(["Chỉ tiêu", "Công thức", "Ghi chú"], [[r.kpi, r.formula, "" if pd.isna(r.scope_note) else r.scope_note]
                                            for r in kpi.itertuples()], widths=[3.6, 6.4, 7.5])
FIG("04_cay_kpi.png", "Hình 2: Cây KPI")
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
H2("Bước 2. Power Query: nạp và làm sạch bằng thao tác chuột")
P("Dữ liệu ở đây **chưa sạch hoàn toàn**: có dòng trùng, dòng hàng bán bị trả lại chưa biết thuộc cửa hàng nào, và file "
  "Kế toán dùng mã cửa hàng riêng. Ba việc làm sạch này là một phần của bài. Mọi việc ở bước này làm **bằng chuột** "
  "trong cửa sổ Power Query Editor: mỗi thao tác được Excel tự ghi thành một dòng ở khung **Applied Steps** (bên phải). "
  "Bấm vào một bước để xem dữ liệu tại bước đó; làm sai thì bấm dấu ✕ cạnh tên bước để xoá rồi làm lại.", first=True)
FIG("05_power_query_pipeline.png", "Hình 3: Quy trình Power Query — mỗi ô là một Applied Step, số là số dòng sau bước")
P("Đọc Hình 3 theo làn: ô cam là bước **làm đổi số dòng hoặc ghép bảng** (bỏ trùng, ghép bảng, lọc theo ngày hiệu lực) — "
  "chỗ hay sai nhất, phải kiểm số dòng ngay sau bước đó. Mũi tên nét đứt là phụ thuộc giữa các query: fact_return cần "
  "fact_sales_line; fact_finance cần finance_raw và map_store_finance.")

H3("2.1 Chuẩn bị")
LIST(["Chép Nexa_Sales_P1_Dataset.xlsx và finance_revenue.xlsx vào cùng một thư mục, ví dụ C:\\NexaP1.",
      "Mở một workbook Excel trắng, lưu ngay thành Nexa_P1_<HoTen>.xlsx (file nộp).",
      "Tên lệnh: các lệnh chính ghi theo tên tiếng Việt của Microsoft, lần đầu kèm tiếng Anh trong ngoặc (đối chiếu ở "
      "Phụ lục " + PL["thuat_ngu"] + "); lệnh phụ giữ tiếng Anh. Tên Applied Step Excel tự đặt cũng theo "
      "ngôn ngữ giao diện (ví dụ Removed Duplicates).",
      "Đổi máy hoặc đổi thư mục: Data → Get Data → **Data Source Settings** → chọn đường dẫn cũ → **Change Source…** → "
      "Browse tới thư mục mới → Close → Data → Refresh All. Không cần sửa từng query."], numbered=True)

H3("2.2 Thứ tự tạo query")
P("Chỉ cần hai lượt Get Data: lượt 1 lấy cùng lúc mọi Excel Table trong file dataset, lượt 2 lấy file Kế toán. Sau đó "
  "làm sạch từng query theo thứ tự trong bảng — query sau dùng query trước (fact_return cần fact_sales_line đã bỏ trùng; "
  "fact_finance cần finance_raw và map_store_finance).", first=True)
LIST([NAP_DATASET + ". Chọn 9 Table: 6 bảng dim, tbl_fact_sales_line, tbl_fact_return, tbl_map_store_finance.",
      "Đổi tên: " + DOI_TEN + ".",
      "Lượt 2: Home → New Source → File → Excel Workbook → finance_revenue.xlsx → Table tbl_finance_revenue → OK; đổi tên "
      "thành finance_raw."], numbered=True)
TABLE(["#", "Query", "Tạo bằng", "Các bước chính (Applied Steps)", "Nạp (Close & Load To)", "Số dòng"],
      [[str(i), q, cach, buoc, nap, nf(n)] for i, (q, cach, buoc, nap, n) in enumerate(ORDER, 1)],
      widths=[0.9, 3.0, 3.2, 5.5, 3.3, 1.6], align="lllllr")


def pq_block(q, sub, note=None):
    H3(sub)
    if note:
        P(note, first=True)
    TABLE(["#", "Applied Step", "Thao tác chuột", "Số dòng sau bước"],
          [[str(k), stp, how, "—" if n is None else nf(n)] for k, (stp, _, how, n, _) in enumerate(STEPS[q], 1)],
          widths=[0.7, 3.6, 10.9, 2.3], align="lllr")


pq_block("fact_sales_line", "2.3 fact_sales_line: bỏ dòng trùng, thêm cột tiền",
         "Bảng lớn nhất và là nơi duy nhất có dòng trùng. Bỏ trùng **sau** khi đặt kiểu, để hai dòng giống nhau về giá trị "
         "nhưng khác kiểu (ví dụ ngày dạng chữ) không bị coi là khác nhau. Cột tiền làm bằng Add Column → Standard, "
         "không cần gõ công thức.")
P("Kiểm tra: thanh trạng thái của Power Query Editor chỉ hiện “999+ ROWS” và Column profile mặc định chỉ xét 1.000 dòng "
  "đầu. Muốn đếm đủ: chuột phải query → Reference → Transform → Count Rows, xem xong xoá query tạm. Sau Removed "
  f"Duplicates phải còn {nf(rows['fact_sales_line_clean'])} dòng (bỏ {nf(rows['fact_sales_line'] - rows['fact_sales_line_clean'])} dòng trùng).")
pq_block("fact_return", "2.4 fact_return: ghép cửa hàng, kênh, sản phẩm từ dòng bán gốc",
         "fact_return không có store_key nên không lọc được theo cửa hàng, kênh hay danh mục. Ghép (Merge) với "
         "fact_sales_line đã bỏ trùng — mỗi dòng trả hàng ghép đúng một dòng bán — để mang 3 khoá này sang.")
P(f"Kiểm tra: số dòng trước và sau ghép đều là {nf(rows['fact_return'])}; nếu tăng là do ghép với fact_sales_line "
  "**chưa** bỏ trùng. Lọc thử store_key = null: phải không còn dòng nào.")
H3("2.5 Kế toán: finance_raw, map_store_finance → fact_finance")
P("finance_raw và map_store_finance là query trung gian (chỉ tạo kết nối, không nạp vào Mô hình Dữ liệu). File Kế toán "
  "chỉ có mã cửa hàng của Kế toán (fin_store_code); fact_finance gắn store_key bằng bảng map theo **mã và ngày hiệu lực** "
  "(cửa hàng 38 đổi mã GBR-03 → GBR-03N từ 01/07/2019):", first=True)
pq_block("fact_finance", "Các bước của fact_finance")
P(f"Kiểm tra: fact_finance có đúng {nf(fin_rows)} dòng như finance_raw, và đúng 3 dòng HQ-ADJ có store_key trống "
  "(giữ lại để giải thích ở tổng chuỗi, không gán cho cửa hàng nào). Với dữ liệu hiện tại bước Filtered Rows giữ nguyên "
  "số dòng: mã cũ GBR-03 chỉ có bút toán đến 27/06/2019, mã mới GBR-03N từ 01/07/2019. Bước này là chốt an toàn — nếu "
  "Kế toán còn hạch toán mã cũ sau ngày đổi, dòng đó bị loại và số dòng giảm, là tín hiệu cần hỏi lại Kế toán.")
H3("2.6 Các bảng dim và bảng _Measures")
dim_note = {"dim_date": "1 dòng/ngày 2016–2021; month_label dạng 12/2020", "dim_store": "66 cửa hàng + dòng 0 = Online",
            "dim_product": "category, subcategory, brand", "dim_channel": "POS, ONLINE",
            "dim_promotion": "Mã khuyến mãi", "dim_order_status": "COMPLETED, CANCELLED, PENDING"}
TABLE(["Query", "Excel Table nguồn", "Số dòng", "Ghi chú"],
      [[d, "tbl_" + d, nf(n), dim_note[d]] for d, n in DIMS.items()], widths=[3.5, 4, 2, 8], align="llrl")
P("Bảng dim chỉ cần kiểm bước Changed Type: dim_date[date] và month_end_date là Date; các khoá là Whole Number.")
P("Bảng **_Measures** chỉ để chứa số đo (Excel bắt buộc số đo phải thuộc một bảng). Tạo bằng tay:", first=True)
LIST(["Thêm sheet tên _Measures; ô A1 gõ ghi_chu, ô A2 gõ bảng chứa số đo.",
      "Chọn A1:A2 → Insert → Table (tick My table has headers) → tab Table Design → Table Name: _Measures.",
      "Data → From Table/Range → trong Power Query Editor không cần làm gì thêm → Close & Load To (bước 2.7)."],
     numbered=True)
H3("2.7 Nạp vào Mô hình Dữ liệu")
LIST(["Home → **Close & Load To…** (không bấm Close & Load, vì lệnh đó đổ mọi bảng ra sheet).",
      "Chọn **Only Create Connection** và tick **Add this data to the Data Model** cho các dim, fact_sales_line, "
      "fact_return, fact_finance và _Measures.",
      "finance_raw và map_store_finance: chỉ **Only Create Connection**, không tick Mô hình Dữ liệu.",
      "Lỡ nạp sai: Data → Queries & Connections → chuột phải query → Load To… để đổi lại.",
      "Kiểm số dòng: Power Pivot → Manage, mỗi tab bảng ghi số dòng ở góc dưới (Record: 1 of …)."], numbered=True)
H3("2.8 Số dòng cần khớp sau khi nạp")
TABLE(["Query", "Số dòng"], [
    ["fact_sales_line (trước / sau Removed Duplicates)", f"{nf(rows['fact_sales_line'])} → {nf(rows['fact_sales_line_clean'])}"],
    ["fact_return", nf(rows["fact_return"])],
    ["fact_finance (trong đó HQ-ADJ có store_key trống)", f"{nf(fin_rows)} (3)"],
    ["dim_date", nf(rows["dim_date"])], ["dim_store", nf(rows["dim_store"])], ["dim_product", nf(rows["dim_product"])],
    ["map_store_finance (chỉ kết nối)", nf(rows["map_store_finance"])],
], widths=[12, 5.5], align="lr")
P("Sai khác số dòng thường do quên bỏ trùng, chọn nhầm sheet thay vì Table ở Navigator, hoặc ghép fact_return với "
  "fact_sales_line chưa bỏ trùng (fact_return dư dòng).")

# ---------------------------------------------------------------- Bước 3
H2("Bước 3. Data Model trong Power Pivot")
P("Lược đồ hình sao (star schema): bảng dữ kiện ở giữa, bảng chiều xung quanh, quan hệ 1 → nhiều và lọc một chiều. fact_sales_line và fact_finance "
  "không nối trực tiếp với nhau; chúng gặp nhau qua dim_date và dim_store. Mở Power Pivot → **Manage** → Home → "
  "**Diagram View** để thấy mọi bảng đã nạp.", first=True)
FIG("02_mo_hinh_du_lieu.png", "Hình 4: Mô hình dữ liệu")
H3("3.1 Quan hệ cần tạo")
TABLE(["Từ (1)", "Đến (nhiều)", "Trạng thái"], [
    ["dim_date[date]", "fact_sales_line[order_date]", "Active"],
    ["dim_date[date]", "fact_sales_line[delivery_date]", "Inactive (dùng USERELATIONSHIP)"],
    ["dim_date[date]", "fact_return[return_date]", "Active"],
    ["dim_date[date]", "fact_finance[posting_date]", "Active"],
    ["dim_store[store_key]", "fact_sales_line[store_key]", "Active"],
    ["dim_store[store_key]", "fact_return[store_key]", "Active (cột lấy ở Bước 2.4)"],
    ["dim_store[store_key]", "fact_finance[store_key]", "Active"],
    ["dim_product[product_key]", "fact_sales_line[product_key]", "Active"],
    ["dim_product[product_key]", "fact_return[product_key]", "Active (cột lấy ở Bước 2.4)"],
    ["dim_promotion[promo_code]", "fact_sales_line[promo_code]", "Active"],
    ["dim_order_status[order_status]", "fact_sales_line[order_status]", "Active"],
    ["dim_channel[channel_code]", "dim_store[channel_code]", "Active (kênh lọc qua cửa hàng)"],
], widths=[5.5, 6.5, 5.5])
LIST(["**Tạo quan hệ:** trong Diagram View kéo cột khoá ở bảng dim thả lên cột tương ứng ở bảng fact; hoặc Design → "
      "Create Relationship và chọn 2 bảng, 2 cột.",
      "**Quan hệ inactive:** kéo dim_date[date] lên fact_sales_line[delivery_date] lần thứ hai. Vì đã có quan hệ active "
      "giữa hai bảng, Excel tự để quan hệ mới ở trạng thái inactive (đường nét đứt). Kiểm ở Design → Manage Relationships: "
      "cột Active = No.",
      "**Không tạo** quan hệ giữa các bảng fact, và không nối map_store_finance (bảng này chỉ dùng trong Power Query)."])
H3("3.2 Cấu hình bắt buộc")
LIST(["**Mark as Date Table.** Chọn bảng dim_date → Design → Mark as Date Table → cột date. Bắt buộc để DATEADD và "
      "SAMEPERIODLASTYEAR chạy đúng khi lọc theo năm/tháng.",
      "**Sort by Column.** Ở Data View chọn cột dim_date[month_label] → Home → Sort by Column → sort theo month_id; nếu không, "
      "01/2017 đứng trước 12/2016.",
      "**Hide from Client Tools.** Chuột phải các khoá ở bảng fact (store_key, product_key, order_status, promo_code, các cột "
      "ngày) và cột ghi_chu của _Measures → Hide from Client Tools, để học viên luôn kéo trường từ bảng dim.",
      "**Định dạng.** Mỗi measure chọn định dạng ngay khi tạo (Currency 0 chữ số lẻ, Percentage 1–2 chữ số lẻ); "
      "PivotTable và CUBEVALUE dùng định dạng đó."])
H3("3.3 Ba phép thử mô hình (trước khi viết measure)")
P("Insert → PivotTable → **From Data Model** → New Worksheet (đặt tên sheet `_kiem_tra`). Kéo cột số vào Values, Excel "
  "tự tạo measure ngầm (Sum of …); với order_number đổi Value Field Settings → **Distinct Count**. Ra đúng số thì mô hình đã đúng:",
  first=True)
TABLE(["Phép thử", "Rows / Filters", "Values", "Phải ra"], [
    ["Đơn COMPLETED năm 2020", "Filters: dim_order_status[order_status] = COMPLETED; Rows: dim_date[year]",
     "fact_sales_line[order_number] · Distinct Count", f"2020: {nf(ytd['orders'])}"],
    ["Finance năm 2020 (gồm HQ-ADJ)", "Rows: dim_date[year]", "fact_finance[net_revenue] · Sum",
     f"2020: {nf(R['finance_year']['2020'], 2)}"],
    ["Finance gắn đúng kênh", "Rows: dim_channel[channel_code]", "fact_finance[fin_store_code] · Count",
     f"ONLINE: {nf(fin_online)} · POS: {nf(fin_rows - fin_online - 3)} · (blank): 3"],
], widths=[3.6, 6.1, 4.3, 3.5])
P("Phép thử thứ ba ra một dòng (blank) gồm 3 bút toán HQ-ADJ — đúng thiết kế. Nếu (blank) lớn hơn 3, fact_finance chưa gắn "
  "được store_key (xem lại 2.5).")

# ---------------------------------------------------------------- Bước 4
H2("Bước 4. Viết số đo DAX trong Power Pivot")
P(f"Bộ bắt buộc gồm **{len(MEAS)} số đo**, viết theo đúng thứ tự cây KPI: doanh thu → đơn hàng → Kế toán và chênh lệch → "
  "quy đổi online → so sánh kỳ. Mỗi số đo mới kiểm ngay bằng PivotTable ở sheet `_kiem_tra` trước khi dùng vào số đo khác. "
  "Phần trình bày (% thay đổi, câu Điểm chính, tô màu nhấn) **không** viết bằng DAX mà bằng công thức Excel ở Bước 5 – "
  "đúng sở trường của Excel và dễ sửa hơn.", first=True)
H3("4.1 Tạo số đo trong Excel")
LIST(["**Cách 1 — hộp thoại:** Power Pivot → Measures → **New Measure…** → Table name: `_Measures` → Measure name → "
      "Formula (phần sau dấu `:=`) → Check formula → Category/Format → OK.",
      "**Cách 2 — Calculation Area:** Power Pivot → Manage → Data View → mở tab bảng _Measures → bấm một ô trong vùng lưới "
      "dưới bảng → gõ `Tên := công thức` theo các mục 4.3–4.8 → Enter.",
      "Tên số đo phải gõ **đúng như trong tài liệu** (kể cả dấu cách, chữ hoa, ký hiệu Δ), vì các ô CUBEVALUE ở Bước 5 gọi số đo theo tên.",
      "Sửa số đo: Power Pivot → Measures → Manage Measures → Edit."], numbered=True)
H3("4.2 Nguyên tắc chung")
LIST(["Dùng DIVIDE thay cho phép chia “/” để tránh lỗi chia cho 0 (DIVIDE trả về trống).",
      "Bộ lọc trạng thái đặt **trong** số đo (KEEPFILTERS), không đặt ở filter của PivotTable, để mọi sheet dùng chung một định nghĩa.",
      "Định dạng: tiền `#,##0` USD; tỷ lệ `0.0%`; chênh lệch `+#,##0;-#,##0;0`.",
      "Excel không có thư mục hiển thị (display folder) cho số đo: tạo theo thứ tự các mục 4.3–4.7 để danh sách dễ tìm.",
      "Excel không có cửa sổ chạy truy vấn DAX: kiểm số đo bằng PivotTable (Bước 3.3) hoặc một ô CUBEVALUE."])
DAX_SUB = {1: "4.3 Nhóm Doanh thu", 2: "4.4 Nhóm Đơn hàng và giá", 3: "4.5 Nhóm Kế toán (Finance) và đối soát",
           4: "4.6 Quy đổi online về ngày giao", 5: "4.7 So sánh kỳ"}
DAX_NOTE = {
    1: "Gross Sales = doanh thu bán hàng; Discounts = chiết khấu thương mại; Returns = hàng bán bị trả lại; "
       "Net Revenue = doanh thu thuần (VAS 14). gross_amount đã tính ở Power Query nên chỉ cần SUM.",
    3: f"[Vượt ngưỡng] so một ô tháng × cửa hàng với ngưỡng {nf(thr['abs_usd'])} USD và {pct(thr['pct'], 0)}. "
       "[Số cửa hàng vượt ngưỡng] duyệt từng cặp tháng × cửa hàng (CROSSJOIN), nên đúng cả khi chọn một tháng lẫn cả năm. "
       "[Δ tuyệt đối] dùng để xếp bảng cửa hàng theo độ lớn chênh lệch.",
    4: "Kế toán ghi nhận đơn online theo ngày giao. USERELATIONSHIP bật mối quan hệ không hoạt động (inactive) "
       "delivery_date ngay trong số đo; POS giữ ngày đặt hàng.",
    5: "Chỉ chạy đúng khi dim_date đã Đánh dấu là Bảng ngày tháng và slicer chọn **một** năm + **một** tháng. "
       "% thay đổi tính ở ô Excel: `=giá trị / giá trị LM − 1`.",
}
for i, (title, lines) in enumerate(DAXS, 1):
    H3(DAX_SUB[i])
    if i in DAX_NOTE:
        P(DAX_NOTE[i], first=True)
    CODE("\n".join(lines))
H3("4.8 Số đo nâng cao (tuỳ chọn)")
P(f"{len(MEAS_NC)} số đo dưới đây dành cho các thẻ KPI nâng cao và phần sáng tạo: bút toán điều chỉnh HQ-ADJ, "
  "đơn huỷ/đang giao, cơ cấu doanh thu và so sánh cùng kỳ. Dashboard tham chiếu dùng Tỷ trọng online, Tỷ lệ chiết khấu, "
  "Cancelled Orders, Pending Orders ở thẻ KPI sheet Doanh_thu và Hoan_tra. Tạo sau khi bộ bắt buộc đã khớp số.",
  first=True)
CODE("\n".join(l for _, ls in DAXS_NC for l in ls))
H3("4.9 Năm lỗi hay gặp")
LIST(["**Quên lọc COMPLETED** → Net Sales cao hơn, AOV sai; Δ âm ở mọi cửa hàng.",
      "**Đếm Orders bằng COUNTROWS** → đếm dòng chứ không đếm đơn, AOV thấp bất thường (~1/2,4 giá trị đúng).",
      "**Trừ Returns theo ngày đặt của đơn gốc** → hàng bán bị trả lại không khớp tháng với Kế toán.",
      "**So với total_incl_vat** → Δ ≈ 10 % ở mọi ô (cột này đã gồm thuế GTGT).",
      "**Chưa Đánh dấu là Bảng ngày tháng** → các số đo LM/SPLY/YTD trả về trống hoặc sai."])

# ---------------------------------------------------------------- Bước 5
H2("Bước 5. Dựng 4 sheet dashboard")
H3("5.1 Nguyên tắc trình bày: Storytelling with Data")
P("Dashboard dùng để ra quyết định trước cuộc họp chốt số, nên áp dụng 6 bài học của Cole Nussbaumer Knaflic "
  "(Storytelling with Data); khi chấm phần trải nghiệm (6 điểm) và "
  "lập luận, đối chiếu với bảng này. Dashboard tham chiếu (Excel và HTML) làm theo đúng các quy tắc đó.", first=True)
TABLE(["Bài học", "Áp dụng trong dashboard", "Lỗi hay gặp cần tránh"], [
    ["1. Hiểu bối cảnh", "Mỗi sheet mở bằng câu “Điểm chính” (Big Idea) tự sinh theo số: chuyện gì xảy ra, vì sao, cần làm gì.", "Sheet chỉ có biểu đồ, người xem phải tự đoán kết luận"],
    ["2. Chọn hình phù hợp", "Đường cho xu hướng; cột ngang xếp hạng cho danh mục/quốc gia; waterfall cho cầu doanh thu và phân rã Δ; bảng cho chi tiết.", "Pie/donut/treemap để so sánh, biểu đồ hai trục y, 3D"],
    ["3. Gỡ rối", "Tắt gridlines của sheet, bỏ chú giải khi chỉ có một chuỗi, ẩn field button của PivotChart, nhãn trực tiếp ở cuối thanh.", "Khung đậm, nền màu ở mọi tiêu đề, số trên mọi điểm"],
    ["4. Dẫn mắt", "Xám cho bối cảnh, một màu nhấn (xanh dương) cho tháng chọn / mục lớn nhất; đỏ chỉ cho ⚠ vượt ngưỡng; ▲▼ luôn kèm chữ.", "Mỗi chuỗi một màu, đỏ/xanh lá cho mọi thay đổi"],
    ["5. Nghĩ như nhà thiết kế", "Lưới cột thẳng hàng, định dạng số thống nhất, bảng màu an toàn cho người mù màu, font Segoe UI.", "Số lẻ khác nhau giữa các visual, chữ bị cắt"],
    ["6. Kể chuyện", "Tiêu đề hành động (“Máy tính chiếm 43% Net Sales”), mạch sheet: Tổng quan → Doanh thu đến từ đâu → Đối soát → Hoàn trả; chân trang ghi nguồn, as-of, quy ước Δ.", "Tiêu đề chỉ mô tả (“Cơ cấu theo danh mục”)"],
], widths=[3.2, 8.8, 5.5])
P("Trong Excel, tiêu đề hành động là **một ô công thức** đặt ngay trên biểu đồ (xoá Chart Title của biểu đồ), ghép chữ "
  "với số lấy từ các ô CUBEVALUE, ví dụ `=B20&\" chiếm \"&FIXED(C20/C30*100,0)&\"% doanh thu sau chiết khấu\"` "
  "(B20 = tên danh mục lớn nhất, C20 = Net Sales của nó, C30 = tổng).")
H3("5.2 Khung chung: 4 sheet, 1 sheet pivot, 4 slicer")
LIST(["**Sheet.** `Tong_quan`, `Doanh_thu`, `Doi_soat`, `Hoan_tra` cho người xem; `_pivot` chứa mọi PivotTable nguồn, "
      "các ô CUBEVALUE và bảng phụ của biểu đồ (ẩn khi nộp: chuột phải tab → Hide).",
      "**Lưới.** Mỗi sheet dashboard: View → bỏ tick Gridlines và Headings; cột A rộng 2; cột B–Q rộng 11 (≈ 1.280 px); "
      "nền trắng, font Segoe UI. Hàng 2–3: tiêu đề trang và dòng phụ (nguồn, ngày dữ liệu); hàng 4–6: slicer; hàng 7–9: "
      "Điểm chính; hàng 11–14: thẻ KPI; hàng 16 và 34: tiêu đề hành động, biểu đồ từ hàng 18 và 36; hàng 52: chân trang.",
      "**4 slicer từ Mô hình Dữ liệu** (Insert → Slicer → tab Data Model): dim_date[year], dim_date[month], "
      "dim_channel[channel_name], dim_store[country]. Đặt ở vùng B4:Q6 (Năm B4:E6, Tháng F4:K6, Kênh L4:M6, Quốc gia N4:Q6); "
      "tab Slicer → Columns để slicer nằm ngang (Năm 6 cột, Tháng 12 cột, Quốc gia 3 cột). Mặc định chọn 2020 và 12.",
      "**Tên slicer.** Slicer Settings → Name to use in formulas: kiểm đúng `Slicer_year`, `Slicer_month`, "
      "`Slicer_channel_name`, `Slicer_country` (Excel thêm số 1 nếu tạo lặp — sửa lại cho khớp công thức).",
      "**Dùng chung slicer cho 4 sheet.** Copy 4 slicer ở Tong_quan → Paste sang cùng vùng ở 3 sheet còn lại: bản dán dùng "
      "chung bộ lọc, chọn ở sheet nào cũng đổi cả 4 sheet.",
      "**Report Connections.** Chuột phải từng slicer → Report Connections → tick mọi PivotTable trên `_pivot`. Các bảng phụ "
      "24 tháng dùng CUBEVALUE nên không cần nối (mục 5.4).",
      "**Màu.** Xám #B9B7B0 cho bối cảnh; xanh dương #2A78D6 nhấn; tốt = #1C5CAB, xấu = cam #EB6834; đỏ #D03B3B chỉ cho ⚠ vượt ngưỡng.",
      "Bố cục, màu và câu chữ mẫu: dashboard tham chiếu (Excel và HTML) mentor gửi sau buổi 1."])
H3("5.3 Thẻ KPI và câu Điểm chính bằng CUBEVALUE")
P("CUBEVALUE lấy một số đo từ Mô hình Dữ liệu theo đúng lựa chọn hiện tại của các slicer truyền vào. Đặt các ô CUBE "
  "ở một khối riêng trên sheet `_pivot`, rồi thẻ KPI, % thay đổi và câu Điểm chính chỉ là công thức Excel thường trỏ vào "
  "các ô đó. Khối tham số đọc tháng đang chọn từ slicer:", first=True)
CODE('''B2  Năm chọn        =CUBERANKEDMEMBER("ThisWorkbookDataModel",Slicer_year,1)
C2                  =IFERROR(VALUE(B2),"")
B3  Tháng chọn      =CUBERANKEDMEMBER("ThisWorkbookDataModel",Slicer_month,1)
C3                  =IFERROR(VALUE(B3),"")
B6  Đúng 1 tháng?   =AND(CUBESETCOUNT(Slicer_year)=1,CUBESETCOUNT(Slicer_month)=1)
B7  month_id chọn   =C2*100+C3
B8  tháng trước     =IF(C3=1,(C2-1)*100+12,B7-1)
B9  cùng kỳ         =B7-100
B10 Nhãn kỳ         =TEXT(C3,"00")&"/"&C2''')
P("Mỗi chỉ tiêu một hàng: kỳ chọn, tháng trước (LM), cùng kỳ (SPLY). Net Revenue dùng hai số đo DAX [Net Revenue LM], "
  "[Net Revenue SPLY]; các chỉ tiêu khác truyền thẳng tháng cần so bằng thành viên month_id, không phải viết thêm số đo. "
  "Bọc N(…) để ô trống thành 0, công thức % không lỗi:")
CODE('''C24 Net Revenue     =N(CUBEVALUE("ThisWorkbookDataModel","[Measures].[Net Revenue]",
                        Slicer_year,Slicer_month,Slicer_channel_name,Slicer_country))
D24 LM              =N(CUBEVALUE("ThisWorkbookDataModel","[Measures].[Net Revenue LM]",
                        Slicer_year,Slicer_month,Slicer_channel_name,Slicer_country))
C25 Orders          =N(CUBEVALUE("ThisWorkbookDataModel","[Measures].[Orders]",
                        Slicer_year,Slicer_month,Slicer_channel_name,Slicer_country))
D25 Orders LM       =N(CUBEVALUE("ThisWorkbookDataModel","[Measures].[Orders]",
                        "[dim_date].[month_id].&["&$B$8&"]",Slicer_channel_name,Slicer_country))
F24 % so LM         =IF(D24=0,"",(C24-D24)/ABS(D24))

Thẻ KPI (giá trị) : =C24/1000                      (định dạng #,##0.0" nghìn USD")
Dòng so sánh      : =IF(F24="","– so với tháng trước",IF(F24>0,"▲ ","▼ ")&FIXED(ABS(F24)*100,1)&"% so với tháng trước")
Câu Điểm chính    : ="Net Revenue "&B10&" đạt "&FIXED(C24/1000,1)&" nghìn USD — "
                    &IF(F24>=0,"+","−")&FIXED(ABS(F24)*100,1)&"% so với tháng trước."''')
P("Dùng FIXED(số; số_chữ_số_lẻ) thay cho TEXT(số; \"#,##0.0\") khi ghép số vào câu: FIXED tự dùng dấu thập phân của máy, "
  "còn chuỗi định dạng của TEXT đọc khác nhau giữa máy cài định dạng Việt Nam và máy cài định dạng Mỹ.")
LIST(["Màu dòng so sánh: Home → Conditional Formatting → New Rule → Use a formula, ví dụ `=LEFT($B$13,1)=\"▲\"` → chữ "
      "#1C5CAB; `=LEFT($B$13,1)=\"▼\"` → chữ #EB6834 (với Return Rate, hàng trả lại, đơn huỷ thì đảo màu: tăng là xấu).",
      "Dải Điểm chính: B7 là nhãn nhỏ (“Điểm chính · 12/2020 · mọi kênh”); gộp B8:Q8 cho câu chính, gộp B9:Q9 cho hai gạch "
      "đầu dòng nối bằng CHAR(10), bật Wrap Text; viền trái dày màu nhấn, nền #F2F6FC.",
      "Khi người xem chọn nhiều tháng, ô “Đúng 1 tháng?” = FALSE → các ô LM/SPLY để 0 → dòng so sánh hiện “–”, không hiện lỗi."])
H3("5.4 Biểu đồ và bảng: PivotTable, bảng phụ, chart thường")
TABLE(["Loại hình", "Cách dựng trong Excel", "Ghi chú"], [
    ["Đường 24 tháng có điểm nhấn", "Bảng phụ 24 dòng trên _pivot: cột month_id lùi dần từ tháng chọn (dùng DATE), cột giá trị "
     "=CUBEVALUE(…, \"[dim_date].[month_id].&[\"&ô month_id&\"]\", Slicer_channel_name, Slicer_country), cột nhấn chỉ có giá trị ở "
     "dòng cuối (các dòng khác NA()) → Insert → Line thường", "Chỉ truyền slicer Kênh, Quốc gia nên luôn đủ 24 tháng; chuỗi xám, chuỗi nhấn chỉ hiện marker xanh"],
    ["Waterfall (cầu doanh thu, phân rã Δ)", "4 ô nằm ngang trên _pivot (hàng nhãn, hàng giá trị) trỏ vào các ô CUBEVALUE: Gross Sales, "
     "−Discounts, −Returns, Net Revenue → Insert → Waterfall; bấm cột cuối → Format Data Point → Set as total",
     "Excel 2016+; đặt dữ liệu nằm ngang để Excel vẽ thành một chuỗi"],
    ["Cột ngang xếp hạng", "PivotTable trên _pivot: Rows dim_product[category], Values [Net Sales], sort Largest to Smallest, Report "
     "Layout → Tabular, tắt Grand Totals; bảng phụ 4 cột: tên, giá trị, nhấn = giá trị dòng 1, nền = giá trị các dòng sau → Insert → "
     "Stacked Bar thường; Vertical axis → Categories in reverse order", "Dùng cho danh mục, quốc gia, lý do trả hàng, tỷ lệ trả theo danh mục; đổi slicer màu nhấn tự đi theo"],
    ["Cột theo tháng có cờ đỏ", "Bảng phụ 24 tháng: [Δ sau quy đổi] và [Số cửa hàng vượt ngưỡng] chỉ tính cửa hàng POS (thêm thành viên "
     "\"[dim_channel].[channel_code].&[POS]\" vào CUBEVALUE); đỏ = IF(số cửa hàng > 0, Δ, 0), xám = Δ − đỏ → Stacked Column thường",
     "Chuỗi vượt ngưỡng đỏ #D03B3B, chuỗi khác xám"],
    ["Bảng chi tiết", "PivotTable trên _pivot (Rows store_name hoặc store_code; Value Filters → Top 10; sort giảm dần) → bảng trên "
     "sheet dashboard là các ô trỏ vào pivot để giữ bố cục; cột so với LM/SPLY bằng CUBEVALUE thêm thành viên cửa hàng",
     "Conditional Formatting: data bar cho Net Sales, chữ đỏ cho dòng ⚠ Vượt"],
], widths=[3.6, 9.4, 4.5])
LIST(["Mọi biểu đồ: xoá Chart Title (tiêu đề nằm ở ô phía trên), xoá legend khi chỉ một chuỗi, Gap Width 60 %.",
      "PivotTable → Options → bỏ tick **Autofit column widths on update** để bố cục không nhảy khi đổi slicer."])
VIS = {
    "5.5 Sheet Tong_quan": [
        ["Điểm chính", "Ô công thức (gộp ô)", "Net Revenue, % so LM/SPLY; số đơn hay AOV quyết định biến động; Δ Kế toán và số cửa hàng vượt ngưỡng", "B7:Q9"],
        ["5 thẻ KPI", "Ô trỏ vào CUBEVALUE", "Net Revenue · Orders · AOV · Return Rate · Δ Finance (+ % trên số Kế toán, số cửa hàng vượt ngưỡng)", "B11:Q14"],
        ["Xu hướng 24 tháng", "Line + bảng phụ CUBEVALUE", "Net Revenue theo tháng đặt hàng, tháng chọn nhấn", "B16:J32"],
        ["Cầu doanh thu", "Waterfall", "Gross Sales → −Discounts → −Returns → Net Revenue (total)", "K16:Q32"],
        ["Động lực", "2 Line + bảng phụ", "Orders (nhấn) và AOV (xám) — 2 biểu đồ riêng, cùng trục tháng, không dùng 2 trục y", "B34:J48"],
        ["Đối soát theo tháng", "Stacked Column + bảng phụ", "[Δ sau quy đổi] tách đỏ/xám theo số cửa hàng POS vượt ngưỡng", "K34:Q48"]],
    "5.6 Sheet Doanh_thu": [
        ["Điểm chính", "Ô công thức", "Danh mục lớn nhất và tỷ trọng, thị trường lớn nhất; tỷ trọng online; tỷ lệ chiết khấu", "B7:Q9"],
        ["4 thẻ KPI", "Ô trỏ vào CUBEVALUE", "Net Sales · Units · Tỷ trọng online · Chiết khấu / Gross Sales (2 thẻ sau dùng số đo nâng cao [Tỷ trọng online], [Tỷ lệ chiết khấu])", "B11:Q14"],
        ["Danh mục", "Stacked Bar từ PivotTable", "[Net Sales] theo category; mục lớn nhất xanh, còn lại xám", "B16:G34"],
        ["Quốc gia", "Stacked Bar từ PivotTable", "[Net Sales] theo dim_store[country] (Online là một dòng)", "H16:L34"],
        ["Kênh", "Bảng trỏ vào PivotTable", "channel_name; Net Sales, Orders, AOV; so với LM và SPLY bằng CUBEVALUE", "M16:Q34"],
        ["Top 10 cửa hàng", "Bảng trỏ vào PivotTable Top 10", "store_name, quốc gia; Net Sales (data bar), Orders, AOV, so với LM, SPLY", "B36:Q48"]],
    "5.7 Sheet Doi_soat": [
        ["Điểm chính", "Ô công thức", "Δ Finance, phần lệch thời điểm online, Δ sau quy đổi, các cửa hàng ⚠, HQ-ADJ", "B7:Q9"],
        ["5 thẻ KPI", "Ô trỏ vào CUBEVALUE", "Net Revenue · Finance Net Revenue · Δ Finance (% trên số Kế toán) · Δ sau quy đổi · Số cửa hàng vượt ngưỡng / số cửa hàng có số liệu (đỏ khi > 0)", "B11:Q14"],
        ["Phân rã Δ", "Waterfall", "Online lệch thời điểm (= Δ Finance − Δ sau quy đổi) → Cửa hàng ⚠ (Δ sau quy đổi của các cửa hàng POS vượt ngưỡng) → Còn lại → Δ Kế toán − đơn (total)", "B16:G32"],
        ["Theo tháng", "Stacked Column + bảng phụ", "[Δ sau quy đổi] đỏ/xám theo số cửa hàng POS vượt ngưỡng, 24 tháng", "H16:Q32"],
        ["Chi tiết cửa hàng", "Bảng trỏ vào PivotTable Top 10", "store_code; Net Revenue, Finance Net Revenue, Δ Finance, Δ Finance %, Δ sau quy đổi, Vượt ngưỡng; Top 10 theo [Δ tuyệt đối]", "B34:J50"],
        ["Online", "Line + bảng phụ CUBEVALUE", "Thêm thành viên kênh ONLINE: [Δ Finance] (xám) và [Δ sau quy đổi] (xanh) theo tháng", "K34:Q50"]],
    "5.8 Sheet Hoan_tra": [
        ["Điểm chính", "Ô công thức", "Hàng trả lại và Return Rate (điểm % so cùng kỳ); lý do lớn nhất; dòng trùng và đơn PENDING đã loại", "B7:Q9"],
        ["4 thẻ KPI", "Ô trỏ vào CUBEVALUE + ô ghi chú", "Returns · Return Rate · Đơn huỷ ([Cancelled Orders]) · Dòng trùng đã loại ở Power Query (kèm số đơn [Pending Orders])", "B11:Q14"],
        ["Xu hướng", "Line + bảng phụ CUBEVALUE", "Return Rate 24 tháng, tháng chọn nhấn", "B16:J32"],
        ["Lý do", "Stacked Bar từ PivotTable", "Returns theo fact_return[return_reason], lý do lớn nhất xanh", "K16:Q32"],
        ["Danh mục", "Stacked Bar từ PivotTable", "Return Rate theo category; nhãn kèm Net Sales để đọc cùng quy mô", "B34:J48"],
        ["Huỷ và chờ giao", "Line + bảng phụ CUBEVALUE", "[Cancelled Orders] (xanh) và [Pending Orders] (xám) theo tháng đặt hàng", "K34:Q48"]],
}
for title, rows_ in VIS.items():
    H3(title)
    TABLE(["Khối hình", "Đối tượng Excel", "Trường / số đo", "Vùng ô"], rows_, widths=[2.8, 3.6, 9.2, 1.9])
P("Giữ cùng vùng cho slicer, Điểm chính và KPI ở cả 4 sheet để người xem chuyển sheet không bị nhảy mắt.")
H3("5.9 Tương tác và hoàn thiện")
LIST(["**Biểu đồ xu hướng không lọc theo tháng.** Bảng phụ 24 tháng chỉ truyền slicer Kênh và Quốc gia vào CUBEVALUE, "
      "nên luôn đủ 24 tháng; cửa sổ kết thúc ở tháng đang chọn.",
      "**Xem chi tiết.** Bấm đúp một ô giá trị trong PivotTable (Show Details) để Excel đổ ra sheet mới tối đa 1.000 dòng chi "
      "tiết của ô đó — cách dò nguyên nhân một ô cửa hàng × tháng.",
      "**Làm mới.** Data → Refresh All khi đổi đường dẫn nguồn hoặc dữ liệu; CUBEVALUE tự cập nhật theo.",
      "**Tiêu đề và đơn vị.** Ô tiêu đề nêu kết luận; dòng phụ ghi đơn vị (USD, nghìn USD, %, điểm %) và phạm vi.",
      "**Chân trang.** Mỗi sheet một dòng chữ xám nhỏ: nguồn, ngày dữ liệu, quy ước Δ và ngưỡng.",
      "**Khoá bố cục (tuỳ chọn).** Review → Protect Sheet, tick Use PivotTable & PivotChart và Edit objects để người xem vẫn "
      "dùng được slicer."])


# ---------------------------------------------------------------- Bước 6
if not HOC_VIEN:
    H2("Bước 6. Đối soát số (DÀNH CHO MENTOR — ẩn khi giao đề)")
    P(f"Các số dưới đây tính độc lập từ chính file giao học viên và đã khớp với dashboard tham chiếu. "
      f"Tháng báo cáo {mstr(R['report_month'])}, đủ bộ lọc “Tất cả”.", first=True)
    H3(f"6.1 KPI tháng {mstr(R['report_month'])}")
    krow = lambda name, k, d=0: [name, nf(cur[k], d), nf(lm[k], d), chg(cur[k], lm[k]), nf(sp[k], d), chg(cur[k], sp[k])]
    TABLE(["Chỉ tiêu", "12/2020", "11/2020 (LM)", "Δ% LM", "12/2019 (SPLY)", "Δ% SPLY"], [
        krow("Gross Sales (USD)", "gross_sales", 2), krow("Discounts (USD)", "discounts", 2), krow("Net Sales (USD)", "net_sales", 2),
        krow("Returns (USD)", "returns", 2), krow("Net Revenue (USD)", "net_revenue", 2), krow("Orders", "orders"),
        krow("Units", "units"), krow("AOV (USD/đơn)", "aov", 2), krow("ASP (USD/sp)", "asp", 2),
        ["Return Rate", pct(rr_(cur), 2), pct(rr_(lm), 2), nf((rr_(cur) - rr_(lm)) * 100, 2) + " điểm %",
         pct(rr_(sp), 2), nf((rr_(cur) - rr_(sp)) * 100, 2) + " điểm %"],
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
    FIG("03_quy_trinh_doi_soat.png", "Hình 5: Quy trình đối soát")
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
          "Đổi slicer Kênh sang POS rồi chọn Quốc gia: biểu đồ online trống, các ô CUBEVALUE so sánh kỳ hiện “–”, không có ô #N/A hay #VALUE!."])

if HOC_VIEN:
    H2("Bước 6. Đối soát: tự lập bảng giải thích chênh lệch")
    P("Dùng chính dashboard vừa dựng để tìm vì sao số của mình lệch số Kế toán. Mỗi ô tháng × cửa hàng vượt ngưỡng phải "
      "có nguyên nhân và số tiền; tổng các nguyên nhân bằng Δ, phần chưa giải thích bằng 0.", first=True)
    FIG("03_quy_trinh_doi_soat.png", "Hình 5: Quy trình đối soát")
    H3("6.1 Thứ tự dò tìm")
    LIST(["Tổng chuỗi: so [Net Revenue] với [Finance Net Revenue] theo tháng năm 2020.",
          "Theo kênh: kênh online so thêm [Δ sau quy đổi]; nếu Δ về gần 0 thì đó là lệch thời điểm, không phải lỗi.",
          "Theo cửa hàng: bảng Top 10 theo [Δ tuyệt đối], chỉ đào sâu dòng có cờ ⚠ Vượt.",
          "Theo ngày: bấm đúp ô giá trị (Show Details) hoặc lọc finance_revenue.xlsx theo mã cửa hàng và tháng để xem từng "
          "ngày ghi sổ: ngày trống, dòng ghi trùng, số cao bất thường, đơn đã huỷ.",
          "Ghi nguyên nhân và số tiền vào bảng cầu nối bên dưới; làm lại đến khi phần chưa giải thích = 0."], numbered=True)
    H3("6.2 Khung bảng cầu nối năm 2020 (học viên điền)")
    TABLE(["Khoản", "USD", "Bằng chứng"], [
        ["Net Revenue theo dữ liệu đơn (order_date, return_date)", "…", "số đo [Net Revenue], cả năm 2020"],
        ["± Online lệch thời điểm (ngày đặt ↔ ngày giao)", "…", "[Δ Finance] − [Δ sau quy đổi] kênh online"],
        ["± Hoàn tiền ghi sổ trễ", "…", "refunds ↔ return_date"],
        ["± (nguyên nhân 3, 4, … do học viên tìm)", "…", "cửa hàng, ngày, số tiền"],
        ["= Finance Net Revenue các cửa hàng", "…", "số đo [Finance Net Revenue]"],
        ["Phần chưa giải thích", "0", "phải bằng 0"],
    ], widths=[8.5, 2.5, 6.5], align="lrl")

# ---------------------------------------------------------------- Bước 7
H2("Bước 7. Từ dashboard đến báo cáo")
H3("7.1 Mạch viết")
P("Mỗi phần báo cáo đi theo mạch **Câu hỏi → Phân tích → Bằng chứng (hình, số) → Kết luận**. Báo cáo theo đúng 8 phần "
  "của guild_report_template.pdf; phần IV (Main Analysis) dành cho đối soát.", first=True)
H3("7.2 Câu hỏi gợi ý theo trang")
TABLE(["Trang", "Câu hỏi"], [
    ["Tổng quan", "Tháng 12/2020 doanh nghiệp thực sự bán được bao nhiêu? Vì sao Net Revenue giảm mạnh so với 12/2019?"],
    ["Doanh thu", "Kênh nào, danh mục nào, quốc gia nào đóng góp chính? AOV online khác cửa hàng thế nào?"],
    ["Đối soát", "Chênh lệch với số Kế toán là bao nhiêu? Bao nhiêu phần là lệch thời điểm, bao nhiêu là lỗi thật?"],
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
LIST(["Sheet chi tiết theo ngày cho một cửa hàng: PivotTable posting_date × [Finance Net Revenue], [Net Revenue], [Δ Finance], lọc bằng slicer cửa hàng.",
      "Waterfall cầu nối Orders → Finance theo nguyên nhân, tách thêm từng cửa hàng ⚠ (ô CUBEVALUE cho từng cửa hàng hoặc bảng nguyên nhân dựng bằng Power Query).",
      "Phân tích khuyến mãi: Net Sales và AOV theo promo_code; giảm giá có kéo số đơn lên không.",
      "Quy đổi về tiền tệ thanh toán bằng fx_rate; phân tích theo vùng khách hàng bằng dim_customer."])
H3("7.5 Sản phẩm nộp")
LIST(["File Nexa_P1_<HoTen>.xlsx: các truy vấn Power Query, Mô hình Dữ liệu, số đo DAX và 4 sheet dashboard (+ sheet sáng tạo nếu có). "
      "Mentor mở file, đổi đường dẫn nguồn (Data → Get Data → Data Source Settings → Change Source) rồi Data → Refresh All.",
      "Báo cáo PDF 8 phần, có bảng giải thích chênh lệch vượt ngưỡng năm 2020 (tháng, cửa hàng, Δ, nguyên nhân, số tiền).",
      "Phụ lục: danh mục số đo, ảnh chụp Applied Steps của fact_sales_line và fact_finance, ảnh chụp từng sheet dashboard."])

# ---------------------------------------------------------------- Phụ lục
H2(f"Phụ lục {PL['so_do']}. Danh mục số đo")
TABLE(["Nhóm", "Số đo"], [[t.split("(")[0].strip(), " · ".join(ten_so_do([(t, ls)]))] for t, ls in DAXS] +
      [["Nâng cao · " + t.split("·")[-1].strip(), " · ".join(ten_so_do([(t, ls)]))] for t, ls in DAXS_NC],
      widths=[4.5, 13])
H2(f"Phụ lục {PL['thuat_ngu']}. Bảng thuật ngữ")
P("Tên chỉ tiêu doanh thu theo Chuẩn mực kế toán Việt Nam VAS 14 và chế độ kế toán doanh nghiệp Thông tư 99/2025/TT-BTC "
  "(thay Thông tư 200/2014/TT-BTC từ 01/01/2026); tên menu Excel theo trang hỗ trợ tiếng Việt của Microsoft. "
  "Lần đầu xuất hiện ghi tiếng Việt kèm tiếng Anh; tên file, cột và số đo giữ nguyên tiếng Anh.", first=True)
TABLE(list(BANG[0]), [list(r) for r in BANG[1:]], widths=[2.0, 6.3, 5.4, 3.8], raw_cols={2})

doc.core_properties.title = "Dashboard Guideline – Sales Project 1"
doc.core_properties.author = "EDA Mentor"
doc.save(OUT)
print("saved", OUT.relative_to(P1))
if not HOC_VIEN:                                   # dựng luôn bản học viên
    subprocess.run([sys.executable, __file__, "--hoc-vien"], check=True)
