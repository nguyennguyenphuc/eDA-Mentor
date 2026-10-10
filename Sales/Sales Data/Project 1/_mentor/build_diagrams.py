# -*- coding: utf-8 -*-
"""Sinh các sơ đồ .drawio cho Sales Project 1, rồi gọi draw.io desktop xuất PDF + PNG.

Nguồn sửa được: ../diagrams/*.drawio  (mở bằng app.diagrams.net hoặc draw.io desktop, sửa, xuất lại)
Bản xuất     : ../diagrams/*.pdf  (để in / đính kèm) và ../diagrams/png/*.png (để chèn vào docx/pptx)

Chạy:  python build_diagrams.py            # sinh .drawio + xuất
       python build_diagrams.py --chi-xuat # KHÔNG ghi đè .drawio (giữ bản đã sửa tay), chỉ xuất lại PDF/PNG
Cần: draw.io desktop (lệnh `drawio`); trên Linux không màn hình chạy qua `xvfb-run`.
"""
from thuat_ngu import chuan_hoa
import shutil
import subprocess
import sys
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
OUT = HERE.parent / "diagrams"
PNG = OUT / "png"

G, G10, G35, COR, COR10, BLUE, BLUE10, GRAY, TX = ("#447A2A", "#ECF2EA", "#BED0B4", "#993C1D", "#F7EAE5",
                                                   "#185FA5", "#E6EFF8", "#5F5E5A", "#262626")
FONT = "fontFamily=Segoe UI;"


class Dia:
    def __init__(self, name):
        self.name, self.cells, self.n = name, [], 2

    def _id(self):
        self.n += 1
        return f"c{self.n}"

    def box(self, x, y, w, h, text, fill=G10, stroke=G, font=TX, bold=False, size=13, shape="rounded=1;arcSize=8;",
            align="center", parent="1", extra=""):
        text = chuan_hoa(text, ui=False)
        i = self._id()
        st = (f"{shape}whiteSpace=wrap;html=1;fillColor={fill};strokeColor={stroke};fontColor={font};fontSize={size};"
              f"{FONT}align={align};verticalAlign=middle;spacing=6;{'fontStyle=1;' if bold else ''}{extra}")
        self.cells.append(f'<mxCell id="{i}" value="{escape(text, {chr(34): "&quot;"})}" style="{st}" vertex="1" '
                          f'parent="{parent}"><mxGeometry x="{x}" y="{y}" width="{w}" height="{h}" as="geometry"/></mxCell>')
        return i

    def lane(self, x, y, w, h, title, fill, stroke):
        title = chuan_hoa(title, ui=False)
        return self.box(x, y, w, h, title, fill=fill, stroke=stroke, font=stroke, bold=True, size=14,
                        shape="swimlane;horizontal=0;startSize=34;rounded=1;arcSize=2;", extra="swimlaneFillColor=#FFFFFF;")

    def text(self, x, y, w, h, t, size=12, color=GRAY, align="left", bold=False):
        t = chuan_hoa(t, ui=False)
        return self.box(x, y, w, h, t, fill="none", stroke="none", font=color, size=size, align=align, bold=bold,
                        shape="text;")

    def table(self, x, y, w, title, rows, head=G, fill="#FFFFFF", key_rows=()):
        """Bảng kiểu ERD: tiêu đề + danh sách cột (cột khoá in đậm)."""
        rh, th = 22, 30
        i = self._id()
        st = (f"swimlane;fontStyle=1;childLayout=stackLayout;horizontal=1;startSize={th};horizontalStack=0;"
              f"resizeParent=1;collapsible=0;marginBottom=0;html=1;rounded=1;arcSize=6;fillColor={head};"
              f"strokeColor={head};fontColor=#FFFFFF;fontSize=13;{FONT}swimlaneFillColor={fill};")
        self.cells.append(f'<mxCell id="{i}" value="{escape(title)}" style="{st}" vertex="1" parent="1">'
                          f'<mxGeometry x="{x}" y="{y}" width="{w}" height="{th + rh * len(rows) + 4}" as="geometry"/></mxCell>')
        for k, r in enumerate(rows):
            j = self._id()
            b = "fontStyle=1;" if k in key_rows else ""
            self.cells.append(
                f'<mxCell id="{j}" value="{escape(r)}" style="text;html=1;strokeColor=none;fillColor=none;align=left;'
                f'verticalAlign=middle;spacingLeft=8;fontSize=12;{FONT}fontColor={TX};{b}" vertex="1" parent="{i}">'
                f'<mxGeometry y="{th + rh * k}" width="{w}" height="{rh}" as="geometry"/></mxCell>')
        return i

    def edge(self, a, b, label="", color=GRAY, dashed=False, exit=None, entry=None, width=1.6, ortho=True, pts=()):
        i = self._id()
        st = (f"{'edgeStyle=orthogonalEdgeStyle;' if ortho else ''}rounded=1;html=1;endArrow=block;endFill=1;"
              f"strokeColor={color};strokeWidth={width};fontSize=11;{FONT}fontColor={GRAY};labelBackgroundColor=#FFFFFF;"
              f"{'dashed=1;' if dashed else ''}")
        if exit:
            st += f"exitX={exit[0]};exitY={exit[1]};exitDx=0;exitDy=0;"
        if entry:
            st += f"entryX={entry[0]};entryY={entry[1]};entryDx=0;entryDy=0;"
        self.cells.append(f'<mxCell id="{i}" value="{escape(label)}" style="{st}" edge="1" parent="1" source="{a}" '
                          f'target="{b}"><mxGeometry relative="1" as="geometry">' +
                          (('<Array as="points">' + "".join(f'<mxPoint x="{x}" y="{y}"/>' for x, y in pts) + '</Array>')
                           if pts else "") + '</mxGeometry></mxCell>')
        return i

    def xml(self):
        body = "".join(self.cells)
        return ('<mxfile host="drawio" agent="EDA Mentor build_diagrams.py"><diagram name="' + self.name + '" id="' +
                self.name + '"><mxGraphModel dx="1400" dy="900" grid="1" gridSize="10" guides="1" tooltips="1" '
                'connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="1600" pageHeight="900" '
                'background="#FFFFFF" math="0" shadow="0"><root><mxCell id="0"/><mxCell id="1" parent="0"/>' +
                body + "</root></mxGraphModel></diagram></mxfile>\n")


# ------------------------------------------------------------------ 1. pipeline dữ liệu
def d_pipeline():
    d = Dia("01_pipeline_du_lieu")
    d.text(20, 10, 1300, 34, "Sales Project 1 · Quy trình xử lý dữ liệu từ nguồn đến báo cáo", size=20, color=G, bold=True)
    d.lane(20, 60, 1380, 170, "Mentor chuẩn bị", BLUE10, BLUE)
    d.lane(20, 250, 1380, 190, "Học viên · Excel", G10, G)
    d.lane(20, 460, 1380, 130, "Kiểm tra", COR10, COR)
    b = dict(fill="#FFFFFF")
    s1 = d.box(80, 90, 210, 110, "<b>Global Electronics Retailer</b><br>(Maven Analytics)<br>Sales · Products · Stores<br>Customers · Exchange_Rates", stroke=BLUE, **b)
    s2 = d.box(350, 90, 250, 110, "<b>Bổ sung nghiệp vụ bán lẻ</b><br>+ khuyến mại, trạng thái đơn<br>+ hoàn trả, dòng trùng<br>+ sổ doanh thu của Kế toán", stroke=BLUE, **b)
    s3 = d.box(670, 75, 250, 60, "<b>Nexa_Sales_P1_Dataset.xlsx</b><br>7 dim · 2 fact · map · KPI", stroke=G, fill=G10)
    s4 = d.box(670, 150, 250, 60, "<b>finance_revenue.xlsx</b><br>ngày × mã cửa hàng Kế toán", stroke=COR, fill=COR10)
    s5 = d.box(1000, 90, 250, 110, "<b>Tài liệu kèm dataset</b><br>Data_Dictionary<br>KPI_Definitions<br>giao dịch mẫu", stroke=BLUE, fill=BLUE10)
    d.edge(s1, s2); d.edge(s2, s3); d.edge(s2, s4); d.edge(s2, s5, dashed=True, color=BLUE)
    p1 = d.box(80, 290, 200, 120, "<b>1 · Power Query</b><br>nạp 2 file, đặt kiểu<br>bỏ 120 dòng trùng<br>map fin_store_code<br>theo valid_from/to", stroke=G, **b)
    p2 = d.box(320, 290, 200, 120, "<b>2 · Data Model</b><br>Power Pivot<br>lược đồ hình sao<br>Mark as Date Table<br>inactive delivery_date", stroke=G, **b)
    p3 = d.box(560, 290, 200, 120, "<b>3 · DAX (Power Pivot)</b><br>Gross → Net Sales<br>→ Net Revenue<br>Δ Kế toán, cờ ngưỡng<br>21 số đo bắt buộc", stroke=G, **b)
    p4 = d.box(800, 290, 200, 120, "<b>4 · Dashboard 4 sheet</b><br>Slicer dùng chung<br>PivotTable · biểu đồ<br>CUBEVALUE cho thẻ KPI<br>và câu Điểm chính", stroke=G, **b)
    p5 = d.box(1040, 290, 200, 120, "<b>5 · Báo cáo</b><br>8 phần theo<br>guild_report_template<br>giải thích chênh lệch<br>vượt ngưỡng", stroke=G, **b)
    for a, c in [(p1, p2), (p2, p3), (p3, p4), (p4, p5)]:
        d.edge(a, c, color=G)
    d.edge(s3, p1, color=G, exit=(1, 0.5), entry=(0.3, 0), pts=[(945, 105), (945, 238), (140, 238)])
    d.edge(s4, p1, color=COR, exit=(0.5, 1), entry=(0.8, 0), pts=[(795, 262), (240, 262)])
    k1 = d.box(80, 490, 260, 80, "<b>QC_Summary</b><br>số dòng cần khớp sau khi nạp<br>và làm sạch", stroke=COR, **b)
    k2 = d.box(560, 490, 260, 80, "<b>Số kiểm tra của mentor</b><br>KPI tháng 12/2020, LM, SPLY<br>cầu nối năm 2020", stroke=COR, **b)
    k3 = d.box(1000, 490, 260, 80, "<b>Chấm điểm</b><br>thang điểm 40 · 40 · 20<br>đối chiếu đáp án của mentor", stroke=COR, **b)
    d.edge(p1, k1, dashed=True, color=COR, exit=(0.5, 1), entry=(0.5, 0))
    d.edge(p3, k2, dashed=True, color=COR, exit=(0.5, 1), entry=(0.5, 0))
    d.edge(p5, k3, dashed=True, color=COR, exit=(0.5, 1), entry=(0.5, 0))
    return d


# ------------------------------------------------------------------ 2. mô hình dữ liệu
def d_model():
    d = Dia("02_mo_hinh_du_lieu")
    d.text(20, 10, 1300, 34, "Mô hình dữ liệu (star schema) · quan hệ 1 → nhiều, lọc một chiều", size=20, color=G, bold=True)
    f1 = d.table(560, 250, 250, "fact_sales_line", ["order_number + line_item", "order_date", "delivery_date",
                 "store_key · channel_code", "product_key · customer_key", "quantity · unit_price",
                 "promo_code · discount_amount", "order_status"], head="#2B4F18", key_rows=(0,))
    f2 = d.table(560, 600, 250, "fact_return", ["return_id", "order_number + line_item", "return_date",
                 "store_key · product_key (Power Query)", "return_qty · return_amount", "return_reason"], head="#2B4F18", key_rows=(0,))
    dd = d.table(575, 60, 220, "dim_date", ["date", "month_id · month_label", "year · quarter", "has_data"], key_rows=(0,))
    ds = d.table(120, 230, 220, "dim_store", ["store_key", "store_code · store_name", "channel_code · country", "region · open_date"], key_rows=(0,))
    dc = d.table(120, 420, 220, "dim_channel", ["channel_code", "channel_name", "revenue_recognition_rule"], key_rows=(0,))
    dp = d.table(1020, 60, 230, "dim_product", ["product_key", "product_name · brand", "category · subcategory", "list_price_usd · unit_cost_usd"], key_rows=(0,))
    dm = d.table(1020, 250, 230, "dim_promotion", ["promo_code", "promo_name", "discount_pct"], key_rows=(0,))
    do = d.table(1020, 400, 230, "dim_order_status", ["order_status", "status_name", "count_in_revenue"], key_rows=(0,))
    du = d.table(1020, 550, 230, "dim_customer (bổ sung)", ["customer_key", "country · continent"], key_rows=(0,))
    mp = d.table(120, 600, 220, "map_store_finance", ["fin_store_code", "store_key", "valid_from · valid_to"], head=COR, key_rows=(0,))
    fi = d.table(120, 780, 220, "finance_revenue.xlsx", ["posting_date", "fin_store_code", "revenue_recognized · refunds", "net_revenue · vat_amount"], head=COR)
    d.edge(dd, f1, "order_date", color=G, exit=(0.15, 1), entry=(0.2, 0))
    d.edge(dd, f1, "delivery_date (inactive)", color=G, dashed=True, exit=(0.85, 1), entry=(0.85, 0))
    d.edge(dd, f2, "return_date (1 → *)", color=G, exit=(1, 0.5), entry=(1, 0.3), pts=[(900, 115), (900, 650)])
    d.edge(ds, f1, "1 → *", color=G, exit=(1, 0.5), entry=(0, 0.45))
    d.edge(dc, ds, "1 → * (kênh lọc qua cửa hàng)", color=G, exit=(0.5, 0), entry=(0.5, 1))
    d.edge(ds, f2, "1 → *", color=G, exit=(1, 0.85), entry=(0, 0.55), pts=[(450, 334), (450, 691)])
    d.edge(dp, f2, "1 → *", color=G, exit=(1, 0.5), entry=(1, 0.75), pts=[(1290, 121), (1290, 724)])
    d.edge(dp, f1, "1 → *", color=G, exit=(0, 0.5), entry=(1, 0.12))
    d.edge(dm, f1, "1 → *", color=G, exit=(0, 0.5), entry=(1, 0.4))
    d.edge(do, f1, "1 → *", color=G, exit=(0, 0.5), entry=(1, 0.7))
    d.edge(du, f1, "1 → *", color=G, dashed=True, exit=(0, 0.5), entry=(1, 0.92))
    d.edge(f1, f2, "dòng gốc ← trả hàng", color=GRAY, dashed=True, exit=(0.5, 1), entry=(0.5, 0))
    d.edge(fi, mp, "fin_store_code + ngày hiệu lực", color=COR, exit=(0.5, 0), entry=(0.5, 1))
    d.edge(mp, ds, "store_key", color=COR, exit=(0, 0.5), entry=(0, 0.5), pts=[(80, 660), (80, 300)])
    d.text(560, 800, 700, 100, "Ghi chú<br>• Finance nạp qua Power Query, map sang store_key trước khi vào mô hình (bảng fact_finance).<br>"
           "• fact_finance nối dim_date[date] theo posting_date và dim_store[store_key].<br>"
           "• Nét đứt = quan hệ inactive / tuỳ chọn.", size=12)
    return d


# ------------------------------------------------------------------ 3. quy trình đối soát
def d_recon():
    d = Dia("03_quy_trinh_doi_soat")
    d.text(20, 10, 1400, 34, "Quy trình đối soát doanh thu: dữ liệu đơn ↔ Finance (tháng × cửa hàng × kênh)", size=20, color=G, bold=True)
    w = dict(fill="#FFFFFF")
    o1 = d.box(40, 70, 220, 60, "<b>fact_sales_line</b><br>63.004 dòng", stroke=G, fill=G10)
    o2 = d.box(40, 160, 220, 60, "① Bỏ dòng trùng<br>(trùng hoàn toàn mọi cột)", stroke=G, **w)
    o3 = d.box(40, 250, 220, 60, "② Giữ đơn COMPLETED<br>(loại CANCELLED, PENDING)", stroke=G, **w)
    o4 = d.box(40, 340, 220, 70, "③ Gross Sales − Discounts<br>= <b>Net Sales</b>", stroke=G, **w)
    o5 = d.box(40, 440, 220, 70, "④ Net Sales − Returns<br>= <b>Net Revenue</b> (đơn)", stroke=G, fill=G10)
    f1 = d.box(1180, 70, 220, 60, "<b>finance_revenue.xlsx</b><br>ngày × fin_store_code", stroke=COR, fill=COR10)
    f2 = d.box(1180, 160, 220, 80, "⑤ Map fin_store_code → store_key<br>theo valid_from / valid_to<br>tách riêng HQ-ADJ", stroke=COR, **w)
    f3 = d.box(1180, 270, 220, 60, "Σ net_revenue<br>= <b>Net Revenue (Finance)</b>", stroke=COR, fill=COR10)
    c1 = d.box(560, 380, 300, 70, "⑥ Ghép theo tháng × store_key<br>Δ = Finance − Đơn hàng; Δ % = Δ / |Finance|", stroke=GRAY, **w)
    c2 = d.box(610, 490, 200, 100, "Vượt ngưỡng?<br>|Δ| &gt; 200 USD<br>và |Δ%| &gt; 1 %", stroke=GRAY, fill="#FFF8E5", shape="rhombus;")
    ok = d.box(320, 510, 200, 60, "Trong ngưỡng<br>ghi nhận, không cần giải thích", stroke=G, fill=G10)
    e0 = d.box(900, 500, 260, 80, "⑦ Phân rã nguyên nhân<br>xem theo <b>ngày</b> và theo <b>kênh</b>", stroke=COR, fill=COR10)
    causes = [
        ("Online lệch thời điểm", "đơn tính ngày đặt, Finance tính ngày giao → quy đổi bằng delivery_date"),
        ("Hoàn tiền hạch toán trễ", "Δ dương/âm thành cặp ở 2 ngày/tháng liền kề"),
        ("Đơn huỷ vẫn ghi nhận", "Finance &gt; đơn đúng bằng giá trị đơn CANCELLED"),
        ("Thiếu dữ liệu ngày", "Finance = 0 ở ngày cửa hàng có bán"),
        ("Ghi trùng bút toán", "2 dòng Finance giống hệt nhau"),
        ("Doanh thu gồm VAT", "Finance cao hơn ≈ 10 % cả tháng"),
        ("Điều chỉnh HQ-ADJ", "không thuộc cửa hàng, giải thích ở tổng chuỗi"),
    ]
    grp = d.box(890, 610, 600, 40 + 50 * len(causes), "", fill="#FFFFFF", stroke=COR, shape="rounded=1;arcSize=3;dashed=1;")
    d.text(900, 612, 400, 26, "Nguyên nhân cần dò (kiểm từng cái, ghi số tiền)", size=12, color=COR, bold=True)
    for k, (t, s) in enumerate(causes):
        d.box(900, 642 + k * 50, 580, 44, f"<b>{t}</b> · {s}", stroke=COR, size=12, align="left", **w)
    d.edge(e0, grp, color=COR, exit=(0.5, 1), entry=(0.5, 0))
    rep = d.box(560, 640, 300, 70, "⑧ Bảng giải thích chênh lệch<br>số tiền theo từng nguyên nhân · còn lại = 0", stroke=G, fill=G10)
    for a, b in [(o1, o2), (o2, o3), (o3, o4), (o4, o5), (f1, f2), (f2, f3)]:
        d.edge(a, b, color=G if a in (o1, o2, o3, o4) else COR)
    d.edge(o5, c1, color=G, exit=(1, 0.5), entry=(0, 0.5))
    d.edge(f3, c1, color=COR, exit=(0.5, 1), entry=(1, 0.5))
    d.edge(c1, c2, color=GRAY)
    d.edge(c2, ok, "Không", color=G, exit=(0, 0.5), entry=(1, 0.5))
    d.edge(c2, e0, "Có", color=COR, exit=(1, 0.5), entry=(0, 0.5))
    d.edge(e0, rep, color=COR, exit=(0, 0.85), entry=(1, 0.5), pts=[(880, 568), (880, 675)])
    d.text(40, 560, 260, 200, "<b>Mẹo cho học viên</b><br>• Làm ở mức tháng trước, rồi mới xuống ngày.<br>• Lọc kênh online riêng: "
           "nếu Δ biến mất khi dùng delivery_date thì đó là lệch thời điểm, không phải lỗi.<br>"
           "• Ghi số tiền cho từng nguyên nhân; tổng phải bằng Δ.", size=12)
    return d


# ------------------------------------------------------------------ 4. cây KPI
def d_kpi():
    d = Dia("04_cay_kpi")
    d.text(20, 10, 1300, 34, "Cây KPI Project 1: từ Gross Sales đến chênh lệch với Finance", size=20, color=G, bold=True)
    w = dict(fill="#FFFFFF")
    od = d.box(330, 50, 220, 60, "<b>Orders</b><br>DISTINCTCOUNT(order_number)", stroke=BLUE, **w)
    un = d.box(330, 130, 220, 60, "<b>Units</b><br>Σ quantity", stroke=BLUE, **w)
    aov = d.box(620, 50, 260, 60, "<b>AOV</b> = Net Sales / Orders", stroke=BLUE, fill=BLUE10)
    asp = d.box(620, 130, 260, 60, "<b>ASP</b> = Net Sales / Units", stroke=BLUE, fill=BLUE10)
    gs = d.box(40, 230, 220, 70, "<b>Gross Sales</b><br>Σ quantity × unit_price", stroke=G, **w)
    dc = d.box(40, 330, 220, 60, "<b>Discounts</b><br>Σ discount_amount", stroke=COR, **w)
    ns = d.box(330, 270, 220, 70, "<b>Net Sales</b><br>Gross − Discounts", stroke=G, fill=G10)
    rt = d.box(330, 390, 220, 60, "<b>Returns</b><br>Σ return_amount (return_date)", stroke=COR, **w)
    nr = d.box(620, 320, 220, 70, "<b>Net Revenue</b><br>Net Sales − Returns", stroke=G, fill=G10)
    rr = d.box(620, 470, 220, 60, "<b>Return Rate</b><br>Returns / Net Sales", stroke=COR, **w)
    dl = d.box(910, 315, 260, 80, "<b>Δ Finance</b> = Finance − Net Revenue<br><b>Δ %</b> = Δ / |Finance|", stroke=GRAY, **w)
    fn = d.box(930, 460, 220, 70, "<b>Finance Net Revenue</b><br>Σ net_revenue (đã map)", stroke=COR, fill=COR10)
    fl = d.box(1230, 315, 210, 80, "<b>Cờ vượt ngưỡng</b><br>|Δ| &gt; 200 USD và |Δ %| &gt; 1 %", stroke=GRAY, fill="#FFF8E5")
    d.edge(od, aov, color=BLUE); d.edge(un, asp, color=BLUE)
    d.edge(gs, ns, "+", color=G); d.edge(dc, ns, "−", color=COR); d.edge(ns, nr, "+", color=G)
    d.edge(rt, nr, "−", color=COR, exit=(1, 0.5), entry=(0, 0.8))
    d.edge(rt, rr, color=COR, exit=(0.5, 1), entry=(0, 0.5))
    d.edge(nr, dl, color=G); d.edge(fn, dl, color=COR, exit=(0.5, 0), entry=(0.5, 1)); d.edge(dl, fl, color=GRAY)
    d.text(40, 50, 260, 140, "Net Sales là tử số của AOV, ASP và mẫu số của Return Rate. Orders đếm đơn, không đếm dòng.", size=12)
    d.text(40, 560, 1300, 60, "Phạm vi: chỉ đơn COMPLETED, đã bỏ dòng trùng. Mốc thời gian: order_date cho bán hàng, return_date cho hoàn trả. "
           "Đơn vị USD. Δ dương nghĩa là Finance ghi nhận cao hơn dữ liệu đơn.", size=12)
    return d


def d_pq():
    from pq_thao_tac import DIMS, STEPS
    vn = lambda n: f"{n:,}".replace(",", ".")
    d = Dia("05_power_query_pipeline")
    d.text(20, 10, 1800, 34, "Power Query trong Excel: các bước làm bằng chuột · số dòng sau mỗi bước", size=20,
           color=G, bold=True)
    d.text(20, 42, 1800, 24, "Mỗi ô = một Applied Step trong Power Query Editor (tên bước Excel tự đặt · thao tác · số dòng). "
           "Ô cam = bước làm thay đổi số dòng hoặc ghép bảng. Mũi tên đứt = query dùng kết quả của query khác.", size=12)
    W, GAP, X0, NMAX = 190, 22, 70, 7
    XL = X0 + NMAX * (W + GAP)

    def chain(y, q, h=78, off=0):
        ids = []
        for i, (nm, short, _, n, kind) in enumerate(STEPS[q], off):
            lab = f"<b>{nm}</b><br>{short}" + (f"<br><i>{vn(n)} dòng</i>" if n is not None else "")
            hl = kind == "dong"
            ids.append(d.box(X0 + i * (W + GAP), y, W, h, lab, stroke=COR if hl else G,
                             fill=COR10 if hl else "#FFFFFF", size=11))
        for a_, b_ in zip(ids, ids[1:]):
            d.edge(a_, b_, color=GRAY)
        return ids

    def load(y, prev, text="<b>Mô hình Dữ liệu</b><br>Only Create Connection<br>+ Add to Data Model", h=78):
        b_ = d.box(XL, y, 190, h, text, stroke=G, fill=G10, size=11)
        d.edge(prev, b_, color=G)
        return b_

    d.lane(20, 80, XL + 210, 112, "fact_sales_line", G10, G)
    a = chain(100, "fact_sales_line")
    load(100, a[-1])
    d.lane(20, 205, XL + 210, 112, "fact_return", G10, G)
    b = chain(225, "fact_return")
    load(225, b[-1])
    d.edge(a[-1], b[2], "Merge theo order_number + line_item", color=BLUE, dashed=True, exit=(0.5, 1), entry=(0.5, 0))
    d.lane(20, 330, XL + 210, 300, "Kế toán", COR10, COR)
    fr = chain(350, "finance_raw", h=64)
    mp = chain(350, "map_store_finance", h=64, off=3)
    d.text(X0 + 5 * (W + GAP) + 10, 350, 620, 80, "finance_raw và map_store_finance: Close &amp; Load To → <b>Only Create "
           "Connection</b> (không nạp vào Mô hình Dữ liệu, chỉ làm đầu vào cho fact_finance).", size=12)
    ff = chain(530, "fact_finance")
    load(530, ff[-1])
    d.edge(fr[-1], ff[0], "Reference", color=BLUE, dashed=True, exit=(0.5, 1), entry=(0.5, 0),
           pts=[(X0 + W + GAP + W / 2, 470), (X0 + W / 2, 470)])
    d.edge(mp[-1], ff[1], "Merge theo fin_store_code", color=BLUE, dashed=True, exit=(0.5, 1), entry=(0.5, 0),
           pts=[(X0 + 4 * (W + GAP) + W / 2, 500), (X0 + W + GAP + W / 2, 500)])
    d.lane(20, 645, XL + 210, 150, "6 bảng dim", G10, G)
    dm = [d.box(X0 + i * (W + GAP), 665, W, 60, t, stroke=G, fill="#FFFFFF", size=11) for i, t in
          enumerate(["<b>Source · Navigation</b><br>chọn Table tbl_dim_* (cùng lượt)", "<b>Changed Type</b><br>kiểm kiểu cột",
                     "<b>Đổi tên</b><br>bỏ tiền tố tbl_"])]
    for a_, b_ in zip(dm, dm[1:]):
        d.edge(a_, b_, color=GRAY)
    load(665, dm[-1], h=60)
    d.text(X0, 740, 1500, 44, "<b>Số dòng:</b> " + " · ".join(f"{k} {vn(v)}" for k, v in DIMS.items()) +
           " · dim_date: Mark as Date Table trong Power Pivot", size=12)
    return d


DIAGRAMS = [d_pipeline, d_model, d_recon, d_kpi, d_pq]


def export(src):
    exe = shutil.which("drawio") or next((str(x) for x in (Path(r"C:\\Program Files\\draw.io\\draw.io.exe"),) if x.exists()), None)
    if not exe:
        sys.exit("Không thấy lệnh drawio — cài draw.io desktop rồi chạy lại (bản .drawio vẫn mở/sửa được).")
    pre = ["xvfb-run", "-a"] if shutil.which("xvfb-run") and sys.platform.startswith("linux") else []
    base = pre + [exe, "--no-sandbox", "--export"]
    subprocess.run(base + ["--format", "pdf", "--crop", "--output", str(src.with_suffix(".pdf")), str(src)],
                   check=True, capture_output=True)
    subprocess.run(base + ["--format", "png", "--scale", "2", "--border", "10", "--output",
                           str(PNG / (src.stem + ".png")), str(src)], check=True, capture_output=True)


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    PNG.mkdir(exist_ok=True)
    only_export = "--chi-xuat" in sys.argv
    for fn in DIAGRAMS:
        d = fn()
        src = OUT / (d.name + ".drawio")
        if not only_export:
            src.write_text(d.xml(), encoding="utf-8")
        export(src)
        print("xuat", src.name, "->", src.with_suffix(".pdf").name, "+ png")
