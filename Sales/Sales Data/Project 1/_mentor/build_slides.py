# -*- coding: utf-8 -*-
"""
Dựng slide giảng dạy Sales Project 1 – Kết quả bán hàng và đối soát doanh thu.

Chạy lại được:  python3 _mentor/build_slides.py
Đầu vào (chỉ đọc):
  - ../../../Bank/Project 1/Slide_Project_1_v3.pptx   (template: master, layout, theme, logo, icon)
  - _mentor/ref_numbers.json                          (mọi số tham chiếu)
  - _mentor/Nexa_Sales_P1_Instructor_Key.xlsx         (cầu nối, danh sách vượt ngưỡng – phụ lục mentor)
  - Nexa_Sales_P1_Dataset.xlsx, finance_revenue.xlsx  (số cho phần khám phá dữ liệu)
  - diagrams/png/*.png, _mentor/assets/dash_*.png     (sơ đồ draw.io, ảnh chụp dashboard)
Đầu ra: Slide_Project_1.pptx (thư mục Project 1).

Không có con số nào gõ tay: mọi số trên slide đều đọc từ các file trên.
Các slide phụ lục mentor được ẩn (show="0").
"""
import copy
import io
import json
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ------------------------------------------------------------------ đường dẫn
HERE = Path(__file__).resolve().parent            # .../Project 1/_mentor
P1 = HERE.parent                                   # .../Project 1
REPO = P1.parents[2]                               # .../eDA-Mentor
TEMPLATE = REPO / "Bank" / "Project 1" / "Slide_Project_1_v3.pptx"
OUT = P1 / "Slide_Project_1.pptx"
DIAG = P1 / "diagrams" / "png"
ASSETS = HERE / "assets"

# Tên mentor trên trang bìa (None = giữ nguyên tên trên bìa deck Bank).
MENTORS = ["Nguyễn Phúc Nguyên", "Nguyễn Thảo Nhi", "Đinh Thị Thu Hà"]
PROJECT_LABEL = "Sales Project 1"

# ------------------------------------------------------------------ màu & font (theo deck Bank)
FONT = "Times New Roman"
CODE_FONT = "Courier New"
GREEN = RGBColor(0x2E, 0x5B, 0x31)
ORANGE = RGBColor(0xC5, 0x5A, 0x11)
NAVY = RGBColor(0x12, 0x3A, 0x7A)
RED = RGBColor(0xC0, 0x00, 0x00)
TEAL = RGBColor(0x0F, 0x6E, 0x56)
TXT = RGBColor(0x22, 0x22, 0x22)
GRAY = RGBColor(0x59, 0x59, 0x59)
MUTED = RGBColor(0x7F, 0x7F, 0x7F)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
L_GREEN = RGBColor(0xE2, 0xF0, 0xD9)
L_ORANGE = RGBColor(0xFB, 0xE5, 0xD6)
L_BLUE = RGBColor(0xDE, 0xEA, 0xF6)
L_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
L_RED = RGBColor(0xF8, 0xD7, 0xD7)
SUB_BLUE = RGBColor(0x00, 0x70, 0xC0)
TITLE_GREEN = RGBColor(0x00, 0xB0, 0x50)
CODE_BG = RGBColor(0xF6, 0xF8, 0xFA)
CODE_LN = RGBColor(0xC9, 0xD1, 0xD9)
CODE_TX = RGBColor(0x24, 0x29, 0x2F)
CODE_KW = RGBColor(0x05, 0x50, 0xAE)
CODE_COL = RGBColor(0x0F, 0x6E, 0x56)
CODE_STR = RGBColor(0xA3, 0x15, 0x15)
CODE_CMT = RGBColor(0x6A, 0x73, 0x7D)
GRID = RGBColor(0xD9, 0xD9, 0xD9)
CH_GREEN2 = RGBColor(0x8C, 0xB8, 0x6E)

SLIDE_W, SLIDE_H = 13.333, 7.5

# Tên tiếng Việt dùng thống nhất với dashboard (build_dashboard.py)
CAT_VI = {"Home Appliances": "Gia dụng", "Computers": "Máy tính", "Cameras and camcorders": "Máy ảnh",
          "Cell phones": "Điện thoại", "TV and Video": "TV & Video", "Games and Toys": "Game & Đồ chơi",
          "Audio": "Âm thanh", "Music, Movies and Audio Books": "Nhạc, phim, sách nói"}
REASON_VI = {"DEFECT": "Lỗi sản phẩm", "CHANGE_MIND": "Đổi ý", "WRONG_ITEM": "Giao nhầm hàng",
             "LATE_DELIVERY": "Giao trễ"}

# ------------------------------------------------------------------ số liệu
REF = json.loads((HERE / "ref_numbers.json").read_text(encoding="utf-8"))
REPORT_MONTH = REF["report_month"]
RM_LABEL = f"{REPORT_MONTH % 100:02d}/{REPORT_MONTH // 100}"
LM_ID = (REPORT_MONTH - 1) if REPORT_MONTH % 100 > 1 else (REPORT_MONTH - 89)
SPLY_ID = REPORT_MONTH - 100
LM_LABEL = f"{LM_ID % 100:02d}/{LM_ID // 100}"
SPLY_LABEL = f"{SPLY_ID % 100:02d}/{SPLY_ID // 100}"
FOCUS_YEAR = REPORT_MONTH // 100
THR_ABS = REF["threshold"]["abs_usd"]
THR_PCT = REF["threshold"]["pct"]
AS_OF = pd.Timestamp(REF["as_of"])
AS_OF_LABEL = AS_OF.strftime("%d/%m/%Y")


def vn(x, d=0, sign=False):
    """Định dạng số kiểu Việt Nam: 1.234.567,89; số âm dùng dấu − ."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "–"
    s = f"{abs(x):,.{d}f}".replace(",", "#").replace(".", ",").replace("#", ".")
    if x < 0 and float(s.replace(".", "").replace(",", ".")) != 0:
        return "−" + s
    return ("+" + s) if (sign and x > 0) else s


def pct(x, d=1, sign=False):
    return vn(x * 100, d, sign) + "%"


def rrt(k):
    """Tỷ lệ hoàn trả tính lại từ Returns / Net Sales (không dùng giá trị đã làm tròn)."""
    return k["returns"] / k["net_sales"]


def chg(cur, base):
    return (cur - base) / base if base else np.nan


def load_data():
    """Đọc dataset + Finance, trả về dict các bảng tổng hợp dùng cho phần khám phá dữ liệu."""
    X = pd.read_excel(P1 / "Nexa_Sales_P1_Dataset.xlsx",
                      sheet_name=["fact_sales_line", "fact_return", "dim_product", "dim_store",
                                  "dim_promotion", "sample_transactions", "map_store_finance",
                                  "dim_channel", "dim_order_status"])
    raw = X["fact_sales_line"]
    s = raw.drop_duplicates().copy()
    assert len(raw) == REF["rows"]["fact_sales_line"] and len(s) == REF["rows"]["fact_sales_line_clean"]
    s["gross"] = s.quantity * s.unit_price
    s["ns"] = s.gross - s.discount_amount
    s["year"] = s.order_date.dt.year
    s["mid"] = s.order_date.dt.year * 100 + s.order_date.dt.month
    s = s.merge(X["dim_product"][["product_key", "category"]], on="product_key", how="left")
    c = s[s.order_status == "COMPLETED"]
    r = X["fact_return"].merge(c[["order_number", "line_item", "category", "store_key", "channel_code"]],
                               on=["order_number", "line_item"], how="left")
    r["year"] = r.return_date.dt.year
    r["mid"] = r.return_date.dt.year * 100 + r.return_date.dt.month
    D = {"s": s, "c": c, "r": r, "X": X}
    # tháng: net revenue theo dữ liệu đơn (order_date / return_date)
    m_ns = c.groupby("mid").ns.sum()
    m_rt = r.groupby("mid").return_amount.sum()
    D["month_nr"] = m_ns.sub(m_rt, fill_value=0)
    D["month_orders"] = c.groupby("mid").order_number.nunique()
    # kiểm chứng với ref_numbers
    ry = REF["year"][str(FOCUS_YEAR)]
    got = c[c.year == FOCUS_YEAR].ns.sum() - r[r.year == FOCUS_YEAR].return_amount.sum()
    assert abs(got - ry["net_revenue"]) < 0.05, (got, ry["net_revenue"])
    # kênh theo năm
    D["chan_year"] = c.groupby(["year", "channel_code"]).ns.sum().unstack(fill_value=0)
    D["chan_orders"] = c.groupby(["year", "channel_code"]).order_number.nunique().unstack(fill_value=0)
    # danh mục năm trọng tâm
    cy = c[c.year == FOCUS_YEAR]
    ry_ = r[r.year == FOCUS_YEAR]
    cat = pd.DataFrame({"ns": cy.groupby("category").ns.sum(),
                        "ret": ry_.groupby("category").return_amount.sum()}).fillna(0)
    cat["rr"] = cat.ret / cat.ns
    D["cat"] = cat.sort_values("ns", ascending=False)
    # khuyến mãi năm trọng tâm
    promo = cy.groupby("promo_code").agg(gross=("gross", "sum"), disc=("discount_amount", "sum"))
    promo = promo.join(X["dim_promotion"].set_index("promo_code")[["promo_name"]])
    D["promo"] = promo
    D["promo_share_year"] = c.assign(p=c.promo_code != "NONE").groupby(["year", "p"]).gross.sum().unstack()
    # lý do hoàn trả năm trọng tâm
    D["reason"] = ry_.groupby("return_reason").return_amount.sum().sort_values(ascending=False)
    # tháng báo cáo theo kênh
    cm = c[c.mid == REPORT_MONTH]
    D["rm_chan"] = cm.groupby("channel_code").agg(ns=("ns", "sum"), orders=("order_number", "nunique"))
    cl = c[c.mid == LM_ID]
    D["lm_chan"] = cl.groupby("channel_code").agg(ns=("ns", "sum"))
    D["rm_cat"] = cm.groupby("category").ns.sum().sort_values(ascending=False)
    # đơn theo trạng thái
    D["status_orders"] = s.groupby("order_number").order_status.first().value_counts()
    D["sample"] = X["sample_transactions"]
    D["map"] = X["map_store_finance"]
    return D


def load_key():
    K = pd.read_excel(HERE / "Nexa_Sales_P1_Instructor_Key.xlsx",
                      sheet_name=["planted_anomalies", "systemic_causes", "bridge_store_month",
                                  "bridge_month", "over_threshold"])
    return K


# ================================================================== tiện ích vẽ
def I(x):
    return Inches(x)


def _set_font(run, size=14, color=TXT, bold=False, italic=False, font=FONT):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.name = font
    if color is not None:
        f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", font)


def _bullet(p, char="•", color=None, indent=0.2, level_indent=0.0):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(I(indent + level_indent))))
    pPr.set("indent", str(-int(I(indent))))
    for t in ("a:buNone", "a:buChar", "a:buAutoNum", "a:buClr", "a:buFont"):
        for el in pPr.findall(qn(t)):
            pPr.remove(el)
    if color is not None:
        bc = etree.SubElement(pPr, qn("a:buClr"))
        sc = etree.SubElement(bc, qn("a:srgbClr"))
        sc.set("val", str(color))
    bf = etree.SubElement(pPr, qn("a:buFont"))
    bf.set("typeface", "Arial")
    bu = etree.SubElement(pPr, qn("a:buChar"))
    bu.set("char", char)


def _no_bullet(p):
    pPr = p._p.get_or_add_pPr()
    if pPr.find(qn("a:buNone")) is None:
        etree.SubElement(pPr, qn("a:buNone"))


def fill_tf(tf, paras, size=14, color=TXT, bold=False, align="l", font=FONT, space_after=0, line=None):
    """paras: list; mỗi phần tử là str, list các run, hoặc dict(runs=…, bullet=…, size=…, align=…).
    run = str hoặc (text, {size,color,bold,italic,font})."""
    tf.clear()
    first = True
    for para in paras:
        if isinstance(para, tuple):
            para = {"runs": [para]}
        elif not isinstance(para, dict):
            para = {"runs": para}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        al = para.get("align", align)
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[al]
        sa = para.get("space_after", space_after)
        if sa:
            p.space_after = Pt(sa)
        if para.get("space_before"):
            p.space_before = Pt(para["space_before"])
        ln = para.get("line", line)
        if ln:
            p.line_spacing = ln
        if para.get("bullet"):
            b = para["bullet"]
            _bullet(p, char=b if isinstance(b, str) else "•", color=para.get("bullet_color"),
                    indent=para.get("indent", 0.2), level_indent=para.get("level_indent", 0.0))
        else:
            _no_bullet(p)
        runs = para["runs"]
        if isinstance(runs, str):
            runs = [runs]
        psize = para.get("size", size)
        pcolor = para.get("color", color)
        pbold = para.get("bold", bold)
        for rr in runs:
            if isinstance(rr, str):
                rr = (rr, {})
            text, o = rr
            run = p.add_run()
            run.text = text
            _set_font(run, o.get("size", psize), o.get("color", pcolor), o.get("bold", pbold),
                      o.get("italic", para.get("italic", False)), o.get("font", font))
    return tf


def text(slide, x, y, w, h, paras, size=14, color=TXT, bold=False, align="l", anchor="t",
         margin=(0.05, 0.03, 0.05, 0.03), font=FONT, space_after=0, line=None, name=None):
    tb = slide.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [I(m) for m in margin]
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    if isinstance(paras, str):
        paras = [paras]
    fill_tf(tf, paras, size, color, bold, align, font, space_after, line)
    if name:
        tb.name = name
    return tb


def box(slide, x, y, w, h, fill=None, line=None, lw=1.5, shape="round", adj=0.06, paras=None, size=14,
        color=TXT, bold=False, align="c", anchor="m", margin=(0.08, 0.04, 0.08, 0.04), dash=False, name=None,
        line_spacing=None):
    st = {"round": MSO_SHAPE.ROUNDED_RECTANGLE, "rect": MSO_SHAPE.RECTANGLE, "oval": MSO_SHAPE.OVAL,
          "chev": MSO_SHAPE.CHEVRON, "pent": MSO_SHAPE.PENTAGON, "rarrow": MSO_SHAPE.RIGHT_ARROW,
          "darrow": MSO_SHAPE.DOWN_ARROW}[shape]
    sh = slide.shapes.add_shape(st, I(x), I(y), I(w), I(h))
    sh.shadow.inherit = False
    st_el = sh._element.find(qn("p:style"))
    if st_el is not None:                      # bỏ style theme (tránh bóng đổ effectRef)
        sh._element.remove(st_el)
    if shape == "round":
        sh.adjustments[0] = adj
    if shape in ("chev", "pent"):
        sh.adjustments[0] = adj
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(lw)
        if dash:
            ln = sh.line._get_or_add_ln()
            pd_ = etree.SubElement(ln, qn("a:prstDash"))
            pd_.set("val", "dash")
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [I(m) for m in margin]
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    if paras is not None:
        if isinstance(paras, str):
            paras = [paras]
        fill_tf(tf, paras, size, color, bold, align, line=line_spacing)
    if name:
        sh.name = name
    return sh


def arrow(slide, x1, y1, x2, y2, color=GRAY, w=2.0):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, I(x1), I(y1), I(x2), I(y2))
    st_el = c._element.find(qn("p:style"))
    if st_el is not None:
        c._element.remove(st_el)
    c.line.color.rgb = color
    c.line.width = Pt(w)
    ln = c.line._get_or_add_ln()
    te = etree.SubElement(ln, qn("a:tailEnd"))
    te.set("type", "triangle")
    return c


class Icons:
    """Icon PNG lấy từ media của deck Bank (cùng bộ icon, cùng phong cách)."""
    MAP = {"clock_w": 3, "person": 4, "bar_g": 5, "coins_o": 6, "warn_r": 7, "trend_w": 8, "people_b": 9,
           "bank_w": 10, "building_b": 11, "handshake_g": 12, "search_g": 13, "clip_g": 14, "money_o": 15,
           "eye_o": 16, "check_o": 17, "case_w": 18, "usershield_w": 19, "gears_w": 20, "search_b": 21,
           "layers_b": 22, "scale_b": 23, "bar_w": 24, "warn_w": 25, "calc_w": 26, "db_w": 27, "coins_w": 28,
           "grid_w": 29, "pie_g": 30, "doc_o": 31, "cal_g": 32, "layers_o": 33, "case_b": 34, "trend_g": 35,
           "trend_b": 36, "trend_r": 37, "table_g": 38, "pie_o": 39, "tree_t": 40, "bar_b": 41, "table_r": 42,
           "gauge_g": 43, "tree_g": 44, "gauge_b": 45, "pie_b": 46, "table_b": 47, "tree_b": 48, "bar_o": 49,
           "swap_o": 50, "calcheck_w": 51, "refresh_w": 52, "cal_w": 53, "pct_w": 54, "swap_w": 55,
           "ruler_g": 56, "cal_b": 57, "palette_g": 58, "swap_b": 59, "flag_g": 60, "doc_b": 61, "pie_w": 62,
           "table_gray": 63, "pie_r": 64, "search_w": 78, "bulb_w": 79, "flag_w": 80, "route_w": 81,
           "doc_w": 82, "cap": 2}

    def __init__(self, zf):
        self.zf = zf

    def blob(self, name):
        return self.zf.read(f"ppt/media/image{self.MAP[name]}.png")

    def add(self, slide, name, x, y, s):
        return slide.shapes.add_picture(io.BytesIO(self.blob(name)), I(x), I(y), I(s), I(s))

    def circle(self, slide, name, x, y, d, fill, pad=0.22):
        box(slide, x, y, d, d, fill=fill, shape="oval")
        s = d * (1 - 2 * pad)
        return self.add(slide, name, x + (d - s) / 2, y + (d - s) / 2, s)


def picture_fit(slide, path, x, y, w, h, align="c", border=None):
    im = Image.open(path)
    iw, ih = im.size
    r = min(w / iw, h / ih)
    pw, ph = iw * r, ih * r
    px = x + (w - pw) / 2 if align == "c" else x
    py = y + (h - ph) / 2
    pic = slide.shapes.add_picture(str(path), I(px), I(py), I(pw), I(ph))
    if border is not None:
        pic.line.color.rgb = border
        pic.line.width = Pt(1)
    return pic, (px, py, pw, ph)


def crop_image(src, dst, box_frac):
    """Cắt ảnh theo tỉ lệ (l, t, r, b) ∈ [0,1]."""
    im = Image.open(src)
    W, H = im.size
    l, t, r, b = box_frac
    im.crop((int(l * W), int(t * H), int(r * W), int(b * H))).save(dst)
    return dst


def table(slide, x, y, w, h, rows, col_w=None, size=12, header_fill=GREEN, header_color=WHITE, zebra=True,
          align=None, bold_first_col=False, row_h=None, fills=None, colors=None, bolds=None, header_size=None):
    nr, nc = len(rows), len(rows[0])
    gf = slide.shapes.add_table(nr, nc, I(x), I(y), I(w), I(h))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    # bỏ style mặc định
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None:
        sid.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"   # No Style, Table Grid
    tbl.first_row = True
    tbl.horz_banding = False
    if col_w:
        tot = sum(col_w)
        for i, cw in enumerate(col_w):
            tbl.columns[i].width = I(w * cw / tot)
    rh = row_h or h / nr
    for ri in range(nr):
        tbl.rows[ri].height = I(rh)
        for ci in range(nc):
            cell = tbl.cell(ri, ci)
            cell.margin_left = cell.margin_right = I(0.06)
            cell.margin_top = cell.margin_bottom = I(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            v = rows[ri][ci]
            al = (align[ci] if align else ("l" if ci == 0 else "r"))
            if ri == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = header_fill
                fill_tf(cell.text_frame, [v], size=header_size or size, color=header_color, bold=True,
                        align="c" if ci else "l")
            else:
                f = None
                if fills and fills.get((ri, ci)) is not None:
                    f = fills[(ri, ci)]
                elif fills and fills.get((ri, None)) is not None:
                    f = fills[(ri, None)]
                elif zebra and ri % 2 == 0:
                    f = L_GRAY
                if f is None:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = WHITE
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = f
                col = (colors or {}).get((ri, ci)) or (colors or {}).get((ri, None)) or TXT
                b = (bolds or {}).get((ri, ci)) or (bolds or {}).get((ri, None)) or (bold_first_col and ci == 0)
                fill_tf(cell.text_frame, [v] if not isinstance(v, list) else v, size=size, color=col, bold=b,
                        align=al)
            # viền mảnh
            tcPr = cell._tc.get_or_add_tcPr()
            for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                ln = etree.SubElement(tcPr, qn(tag))
                ln.set("w", "6350")
                sf = etree.SubElement(ln, qn("a:solidFill"))
                c = etree.SubElement(sf, qn("a:srgbClr"))
                c.set("val", "BFBFBF")
            # lnX phải đứng trước solidFill trong tcPr → đưa fill xuống cuối
            for fe in tcPr.findall(qn("a:solidFill")):
                tcPr.remove(fe)
                tcPr.append(fe)
    return gf


def code_box(slide, x, y, w, h, code, size=12, title=None):
    """Hộp code DAX có tô màu đơn giản (từ khoá, cột, chuỗi, chú thích)."""
    import re
    sh = box(slide, x, y, w, h, fill=CODE_BG, line=CODE_LN, lw=1, adj=0.03, align="l", anchor="t",
             margin=(0.12, 0.08, 0.1, 0.06))
    kws = {"CALCULATE", "SUMX", "SUM", "DIVIDE", "DISTINCTCOUNT", "VAR", "RETURN", "IF", "ABS", "VALUES",
           "DATEADD", "SAMEPERIODLASTYEAR", "DATESYTD", "USERELATIONSHIP", "KEEPFILTERS", "CROSSJOIN",
           "BLANK", "MONTH", "COUNTROWS", "FILTER", "SELECTEDVALUE", "ISBLANK", "FORMAT", "MAX", "AND",
           "TRUE", "REMOVEFILTERS", "ALL", "ISINSCOPE", "NOT"}
    tok = re.compile(r'(//.*$|"[^"]*"|[A-Za-z_]+\[[^\]]+\]|\[[^\]]+\]|[A-Z][A-Z_]+(?=\s*\()|\b(?:VAR|RETURN|TRUE|BLANK)\b)')
    paras = []
    for line in code.strip("\n").split("\n"):
        runs = []
        pos = 0
        for m in tok.finditer(line):
            if m.start() > pos:
                runs.append((line[pos:m.start()], {}))
            t = m.group(0)
            if t.startswith("//"):
                runs.append((t, {"color": CODE_CMT, "italic": True}))
            elif t.startswith('"'):
                runs.append((t, {"color": CODE_STR}))
            elif "[" in t:
                runs.append((t, {"color": CODE_COL}))
            elif t in kws or t.isupper():
                runs.append((t, {"color": CODE_KW, "bold": True}))
            else:
                runs.append((t, {}))
            pos = m.end()
        if pos < len(line):
            runs.append((line[pos:], {}))
        if not runs:
            runs = [(" ", {})]
        paras.append({"runs": runs})
    fill_tf(sh.text_frame, paras, size=size, color=CODE_TX, font=CODE_FONT, align="l")
    return sh


# ================================================================== khung slide
class Deck:
    def __init__(self):
        self.zf = zipfile.ZipFile(TEMPLATE)
        self.prs = Presentation(str(TEMPLATE))
        self.icons = Icons(self.zf)
        prs = self.prs
        self.layouts = {l.name: l for l in prs.slide_layouts}
        # Lấy trang bìa, logo, ảnh kết từ deck Bank trước khi xoá slide
        s1 = prs.slides[0]
        self.cover_xml = [copy.deepcopy(sh._element) for sh in s1.shapes]
        self.cover_blobs = {}
        for rel in s1.part.rels.values():
            if "image" in rel.reltype:
                self.cover_blobs[rel.rId] = rel.target_part.blob
        s3 = prs.slides[2]
        self.logo_blob = None
        self.sldnum_xml = None
        for sh in s3.shapes:
            if sh.shape_type == 13 and self.logo_blob is None:     # PICTURE: logo AI VIET NAM
                self.logo_blob = sh.image.blob
                self.logo_geom = (sh.left, sh.top, sh.width, sh.height)
            if sh.is_placeholder and sh.placeholder_format.type is not None and "SLIDE_NUMBER" in str(
                    sh.placeholder_format.type):
                self.sldnum_xml = copy.deepcopy(sh._element)
        last = prs.slides[len(prs.slides) - 1]
        self.end_blob = None
        for sh in last.shapes:
            if sh.shape_type == 13:
                self.end_blob = sh.image.blob
        # Xoá toàn bộ slide của deck Bank (bỏ cả quan hệ để part không còn được lưu)
        sldIdLst = prs.slides._sldIdLst
        for sldId in list(sldIdLst):
            prs.part.drop_rel(sldId.rId)
            sldIdLst.remove(sldId)
        self.sections = []      # (tên section, [sldId])
        self.n = 0

    # ---------------------------------------------------------------- slide mới
    def new(self, layout="SECTION_HEADER", section=None):
        s = self.prs.slides.add_slide(self.layouts[layout])
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        self.n += 1
        if section is not None:
            if not self.sections or self.sections[-1][0] != section:
                self.sections.append((section, []))
        self.sections[-1][1].append(s.slide_id)
        return s

    def slide_number(self, s):
        el = copy.deepcopy(self.sldnum_xml)
        s.shapes._spTree.append(el)

    def content(self, section, header, subtitle, takeaway=None, hidden=False):
        """Slide nội dung theo khung deck Bank: tiêu đề phần (xanh lá, 44pt), logo, số trang, phụ đề ❖."""
        s = self.new("SECTION_HEADER", section)
        text(s, 0, 0, SLIDE_W, 0.81, [header], size=40, color=TITLE_GREEN, bold=True, align="c", anchor="m",
             margin=(2.1, 0.02, 0.6, 0.02))
        l, t, w, h = self.logo_geom
        s.shapes.add_picture(io.BytesIO(self.logo_blob), l, t, w, h)
        self.slide_number(s)
        text(s, 0.05, 0.93, 10.3, 0.4,
             [{"runs": [(subtitle, {"bold": True})], "bullet": "❖", "bullet_color": SUB_BLUE, "indent": 0.25}],
             size=15, color=SUB_BLUE)
        if not hidden:
            text(s, 10.4, 0.95, 2.85, 0.35, [(PROJECT_LABEL, {"italic": True})], size=12, color=MUTED, align="r")
        if takeaway:
            self.takeaway(s, takeaway)
        if hidden:
            s._element.set("show", "0")
            box(s, 10.55, 1.0, 2.7, 0.32, fill=RED, shape="round", adj=0.3,
                paras=[("DÀNH CHO MENTOR", {"bold": True})], size=11, color=WHITE)
        return s

    def takeaway(self, s, runs):
        if isinstance(runs, str):
            runs = [runs]
        box(s, 0.38, 6.72, 12.55, 0.5, fill=L_GREEN, shape="rect", paras=[{"runs": runs}], size=17,
            color=TXT, align="c", margin=(0.1, 0.02, 0.1, 0.02))

    # ---------------------------------------------------------------- lưu
    def save(self):
        prs = self.prs
        # Section (PowerPoint 2010+): p14:sectionLst trong extLst của presentation.xml
        P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
        pres = prs.part._element
        extLst = pres.find(qn("p:extLst"))
        if extLst is None:
            extLst = etree.SubElement(pres, qn("p:extLst"))
        for ext in extLst.findall(qn("p:ext")):
            if ext.get("uri") == "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}":
                extLst.remove(ext)
        ext = etree.Element(qn("p:ext"))
        ext.set("uri", "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}")
        sl = etree.SubElement(ext, f"{{{P14}}}sectionLst", nsmap={"p14": P14})
        import uuid
        for name, ids in self.sections:
            se = etree.SubElement(sl, f"{{{P14}}}section")
            se.set("name", name)
            se.set("id", "{" + str(uuid.uuid5(uuid.NAMESPACE_URL, "nexa-p1/" + name)).upper() + "}")
            lst = etree.SubElement(se, f"{{{P14}}}sldIdLst")
            for i in ids:
                e = etree.SubElement(lst, f"{{{P14}}}sldId")
                e.set("id", str(i))
        extLst.insert(0, ext)
        from slide_notes import apply_notes          # phần thuyết trình (speaker notes)
        apply_notes(prs, REF, load_key(), vn, pct)
        cp = prs.core_properties
        cp.title = "Sales Project 1 – Kết quả bán hàng và đối soát doanh thu"
        cp.subject = "EDA Mentorship – Nexa Home"
        cp.author = "EDA Mentorship"
        prs.save(str(OUT))


# ================================================================== biểu đồ
def style_chart(ch, size=12, legend=True, legend_pos=XL_LEGEND_POSITION.TOP):
    ch.has_title = False
    ch.font.name = FONT
    ch.font.size = Pt(size)
    ch.font.color.rgb = GRAY
    ch.has_legend = legend
    if legend:
        ch.legend.position = legend_pos
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(size)
        ch.legend.font.name = FONT


def style_axes(ch, num_fmt='#,##0', vmin=None, vmax=None, major=None, show_val=True):
    va = ch.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = GRID
    va.major_gridlines.format.line.width = Pt(0.75)
    va.format.line.fill.background()
    va.tick_labels.number_format = num_fmt
    va.tick_labels.number_format_is_linked = False
    va.tick_labels.font.size = Pt(11)
    va.tick_labels.font.name = FONT
    if vmin is not None:
        va.minimum_scale = vmin
    if vmax is not None:
        va.maximum_scale = vmax
    if major is not None:
        va.major_unit = major
    if not show_val:
        va.tick_label_position = XL_TICK_LABEL_POSITION.NONE
    ca = ch.category_axis
    ca.format.line.color.rgb = RGBColor(0xA6, 0xA6, 0xA6)
    ca.tick_labels.font.size = Pt(11)
    ca.tick_labels.font.name = FONT
    ca.has_major_gridlines = False


def color_series(ser, color, line=False, width=2.25, marker=False):
    if line:
        ser.format.line.color.rgb = color
        ser.format.line.width = Pt(width)
        ser.smooth = False
        from pptx.enum.chart import XL_MARKER_STYLE
        if marker:
            ser.marker.style = XL_MARKER_STYLE.CIRCLE
            ser.marker.size = 5
            ser.marker.format.fill.solid()
            ser.marker.format.fill.fore_color.rgb = color
            ser.marker.format.line.color.rgb = color
        else:
            ser.marker.style = XL_MARKER_STYLE.NONE
    else:
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = color
        ser.format.line.fill.background()


def point_labels(ser, labels, pos=XL_LABEL_POSITION.OUTSIDE_END, size=11, color=TXT, bold=False, idx=None):
    for i, lab in enumerate(labels):
        if lab is None or (idx is not None and i not in idx):
            continue
        dl = ser.points[i].data_label
        tf = dl.text_frame
        tf.text = lab
        for r in tf.paragraphs[0].runs:
            _set_font(r, size, color, bold)
        if pos is not None:
            dl.position = pos


def chart_title(slide, x, y, w, t, color=GREEN):
    box(slide, x, y, w, 0.36, fill=color, shape="rect", paras=[(t, {"bold": True})], size=14, color=WHITE,
        align="l", margin=(0.12, 0.02, 0.1, 0.02))


def chart_frame(slide, x, y, w, h, title, color=GREEN, caption=None):
    box(slide, x, y, w, h, fill=WHITE, line=color, lw=1.25, shape="rect")
    chart_title(slide, x, y, w, title, color)
    if caption:
        text(slide, x + 0.05, y + h - 0.34, w - 0.1, 0.3, [(caption, {"italic": True})], size=11, color=GRAY,
             align="c")


# ================================================================== các slide
def build():
    D = load_data()
    K = load_key()
    dk = Deck()
    ic = dk.icons
    m, lm, sply = REF["month"]["current"], REF["month"]["lm"], REF["month"]["sply"]
    ytd = REF["ytd"]["current"]
    rows = REF["rows"]

    # ---------------------------------------------------------------- 1. Bìa
    s = dk.new("TITLE", "Mở đầu")
    rid_map = {}
    for rid, blob in dk.cover_blobs.items():
        _, new_rid = s.part.get_or_add_image_part(io.BytesIO(blob))
        rid_map[rid] = new_rid
    names = MENTORS
    for el in dk.cover_xml:
        el = copy.deepcopy(el)
        txt = "".join(el.itertext())
        if "[Icon" in txt:
            continue
        for blip in el.iter(qn("a:blip")):
            old = blip.get(qn("r:embed"))
            if old in rid_map:
                blip.set(qn("r:embed"), rid_map[old])
        s.shapes._spTree.append(el)
    for sh in s.shapes:
        if not sh.has_text_frame:
            continue
        t = sh.text_frame.text
        if "Finance Project" in t:
            p = sh.text_frame.paragraphs[0]
            runs = p.runs
            runs[0].text = PROJECT_LABEL
            for r in runs[1:]:
                r.text = ""
    for sh in s.shapes:
        top = sh.top / 914400 if sh.top is not None else 0
        if 4.0 < top < 6.5:
            sh.top = sh.top + I(0.45)
    text(s, 1.5, 3.62, 10.33, 0.5, [("Kết quả bán hàng và đối soát doanh thu · Nexa Home", {"italic": True})],
         size=22, color=NAVY, align="c")
    if names:
        # ô tên mentor = ô có chữ ở cột trái (dưới nhãn "Mentor:"), không phải nhãn kết thúc bằng ":"
        name_boxes = [sh for sh in s.shapes if sh.has_text_frame and sh.left is not None and sh.left / 914400 < 7
                      and sh.text_frame.text.strip() and not sh.text_frame.text.strip().endswith(":")
                      and sh.top / 914400 > 4.6]
        name_boxes.sort(key=lambda z: z.top)
        assert len(name_boxes) == len(names), [z.text_frame.text for z in name_boxes]
        for sh, nm in zip(name_boxes, names):
            sh.text_frame.paragraphs[0].runs[0].text = nm
            sh.width = max(sh.width, I(3.4))          # tên dài không bị xuống dòng
            sh.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT

    # ---------------------------------------------------------------- 2. Outline
    parts = ["Bối cảnh bán lẻ và đối soát doanh thu", "Đề bài xây dựng dashboard", "Hệ thống chỉ số KPI",
             "Nguồn dữ liệu và data modeling", "Khám phá dữ liệu", "Các mẫu công thức DAX", "Phân tích kết quả",
             "Báo cáo đầu ra và thang chấm điểm"]
    s = dk.new("OBJECT", "Mở đầu")
    box(s, 0, 0, SLIDE_W, 1.24, fill=RGBColor(0x53, 0x81, 0x35), shape="rect",
        paras=[("Outline", {"bold": True})], size=60, color=WHITE)
    for i, ptxt in enumerate(parts):
        col, row = i // 4, i % 4
        x = 0.75 + col * 6.3
        y = 1.85 + row * 1.35
        box(s, x, y, 5.75, 0.95, fill=RGBColor(0xFB, 0xE4, 0xD4), line=RGBColor(0xBF, 0x90, 0x00), lw=1.5,
            adj=0.2, paras=[(ptxt, {"bold": True})], size=19, color=ORANGE)
        box(s, x - 0.28, y - 0.17, 1.5, 0.34, fill=RGBColor(0x70, 0x30, 0xA0), adj=0.2,
            paras=[(f"PHẦN {i + 1}", {"bold": True})], size=12, color=WHITE)

    # ================================================================ PHẦN 1. Bối cảnh
    SEC = "Bối cảnh bán lẻ"
    sec_name = "Phần 1 · Bối cảnh"
    # ---- 3. Tình huống
    s = dk.content(sec_name, SEC, "Tình huống thực tế: còn 10 phút trước cuộc họp với Finance",
                   ["Muốn tin con số doanh thu, trước hết phải ", ("khớp được số của mình với số Finance",
                                                                    {"bold": True, "color": NAVY})])
    box(s, 0.4, 1.55, 3.5, 3.55, fill=L_ORANGE, line=ORANGE, lw=1.5)
    ic.circle(s, "clock_w", 1.6, 1.7, 1.1, ORANGE)
    text(s, 0.55, 2.85, 3.2, 0.8, [[("Còn ", {}), ("10 phút", {"bold": True}), (" trước họp", {})],
                                    "với Finance"], size=18, align="c")
    ic.add(s, "person", 1.75, 3.62, 0.8)
    text(s, 0.55, 4.45, 3.2, 0.6, ["Sales Manager cần trả lời nhanh 3 câu hỏi"], size=14, align="c",
         color=GRAY)
    box(s, 3.95, 3.05, 0.35, 0.45, fill=ORANGE, shape="rarrow")
    qs = [("1", "Quy mô", "Tháng này Nexa Home thực sự bán được bao nhiêu (Net Revenue, số đơn)?", GREEN, "bar_w"),
          ("2", "Chất lượng doanh thu", "Bao nhiêu bị mất vì giảm giá, hoàn trả, đơn huỷ?", ORANGE, "coins_w"),
          ("3", "Khớp số", "Vì sao số của Sales khác số Finance xác nhận? Chênh ở đâu, bao nhiêu?", RED,
           "scale_b")]
    for i, (n, t, d, colr, icn) in enumerate(qs):
        y = 1.55 + i * 1.2
        box(s, 4.35, y, 5.3, 1.1, fill=WHITE, line=colr, lw=1.5)
        box(s, 4.47, y + 0.3, 0.5, 0.5, fill=colr, shape="oval", paras=[(n, {"bold": True})], size=16,
            color=WHITE, margin=(0, 0, 0, 0))
        if icn.endswith("_w"):
            ic.circle(s, icn, 5.08, y + 0.25, 0.6, colr)
        else:
            ic.add(s, icn, 5.1, y + 0.27, 0.56)
        text(s, 5.8, y + 0.06, 3.8, 1.0, [[(t, {"bold": True, "color": colr})], d], size=14, anchor="m")
    box(s, 9.7, 3.05, 0.35, 0.45, fill=GREEN, shape="rarrow")
    box(s, 10.1, 1.55, 2.83, 3.55, fill=L_GREEN, line=GREEN, lw=1.5)
    ic.circle(s, "trend_w", 10.97, 1.7, 1.1, GREEN)
    text(s, 10.2, 2.9, 2.63, 0.75, [[("Data Analyst", {"bold": True})], "bán lẻ"], size=18, align="c")
    text(s, 10.2, 3.65, 2.63, 1.4, ["Tính KPI từ dữ liệu đơn hàng, đối chiếu với Finance và giải thích "
                                    "từng chênh lệch vượt ngưỡng bằng con số"], size=13, align="c", color=GRAY)
    steps = ["Doanh thu hình thành thế nào", "Hai hệ thống ghi nhận khác nhau", "Vai trò DA và Finance",
             "Xây dashboard 4 trang", "Giải thích chênh lệch"]
    cols = [GREEN, RGBColor(0x54, 0x82, 0x35), GREEN, RGBColor(0x54, 0x82, 0x35), NAVY]
    for i, t in enumerate(steps):
        box(s, 0.42 + i * 2.5, 5.45, 2.45, 0.95, fill=cols[i], shape="pent" if i == 0 else "chev", adj=0.25,
            paras=[(t, {"bold": True})], size=13, color=WHITE, margin=(0.3, 0.02, 0.2, 0.02))

    # ---- 4. Nexa Home & dòng doanh thu
    nst = D["X"]["dim_store"]
    n_pos = int((nst.store_key > 0).sum())
    n_ctry = nst[nst.store_key > 0].country.nunique()
    s = dk.content(sec_name, SEC, "Nexa Home bán hàng thế nào và doanh thu đi qua những bước nào?",
                   ["Một đơn hàng chỉ thành ", ("doanh thu", {"bold": True, "color": NAVY}),
                    " khi đã hoàn tất; hoàn trả làm giảm doanh thu ở ", ("tháng trả hàng", {"bold": True,
                                                                                         "color": ORANGE})])
    stats = [(vn(n_pos), "cửa hàng (POS)", f"tại {n_ctry} quốc gia"), ("1", "kênh online", "website/app"),
             (vn(rows["dim_product"]), "sản phẩm", "điện tử, gia dụng"),
             (vn(rows["orders_total"]), "đơn hàng", f"01/2016 → {AS_OF_LABEL}")]
    for i, (v, l1, l2) in enumerate(stats):
        x = 0.4 + i * 3.15
        box(s, x, 1.5, 2.95, 1.25, fill=WHITE, line=[GREEN, ORANGE, NAVY, TEAL][i], lw=1.5)
        text(s, x, 1.55, 2.95, 0.6, [(v, {"bold": True})], size=28, color=[GREEN, ORANGE, NAVY, TEAL][i],
             align="c")
        text(s, x, 2.13, 2.95, 0.6, [[(l1, {"bold": True})], (l2, {"color": GRAY})], size=13, align="c")
    flow = [("cal_w", "1 · Đặt hàng", "order_date\nPOS: trả tiền, nhận hàng ngay", GREEN),
            ("case_w", "2 · Giao hàng", "delivery_date\nchỉ đơn online", TEAL),
            ("check_o", "3 · Hoàn tất", "order_status = COMPLETED\n(huỷ, đang giao: không tính)", ORANGE),
            ("refresh_w", "4 · Hoàn trả", "return_date\ntrừ vào tháng trả hàng", RED),
            ("doc_w", "5 · Finance ghi nhận", "posting_date\nPOS ngày bán · online ngày giao", NAVY)]
    for i, (icn, t, d, colr) in enumerate(flow):
        x = 0.4 + i * 2.55
        box(s, x, 3.05, 2.2, 3.4, fill=WHITE, line=colr, lw=1.5)
        if icn.endswith("_w"):
            ic.circle(s, icn, x + 0.7, 3.2, 0.8, colr)
        else:
            ic.add(s, icn, x + 0.7, 3.2, 0.8)
        text(s, x + 0.05, 4.1, 2.1, 0.45, [(t, {"bold": True, "color": colr})], size=15, align="c")
        text(s, x + 0.08, 4.55, 2.04, 1.8, d.split("\n"), size=13, align="c", color=TXT, space_after=4)
        if i < 4:
            box(s, x + 2.22, 4.45, 0.3, 0.4, fill=GRAY, shape="rarrow")

    # ---- 5. Vì sao phải đối soát + vai trò
    s = dk.content(sec_name, SEC, "Vì sao số đơn hàng khác số Finance và ai làm gì trong đối soát?",
                   ["Chênh lệch ", ("không phải lúc nào cũng là lỗi", {"bold": True, "color": RED}),
                    ": phải tách lệch thời điểm, lệch quy tắc và lỗi thật"])
    box(s, 0.4, 1.5, 6.15, 2.45, fill=L_BLUE, line=NAVY, lw=1.5)
    text(s, 0.55, 1.55, 5.9, 0.4, [("Cùng một giao dịch, hai cách ghi sổ", {"bold": True, "color": NAVY})],
         size=16)
    two = [("Dữ liệu đơn hàng (POS, website)", "Ghi theo ngày đặt hàng, có cả đơn huỷ, đơn đang giao, dòng "
                                               "trùng do lỗi tải dữ liệu", GREEN),
           ("File Finance (finance_revenue.xlsx)", "Ghi theo ngày hạch toán: online khi giao xong, hoàn tiền "
                                                   "khi Finance xử lý; mã cửa hàng riêng", ORANGE)]
    for i, (t, d, colr) in enumerate(two):
        y = 2.0 + i * 0.95
        box(s, 0.55, y, 0.12, 0.8, fill=colr, shape="rect")
        text(s, 0.75, y - 0.02, 5.7, 0.9, [[(t, {"bold": True, "color": colr})], d], size=13)
    box(s, 0.4, 4.1, 6.15, 2.45, fill=L_ORANGE, line=ORANGE, lw=1.5)
    text(s, 0.55, 4.15, 5.9, 0.4, [("Nếu không đối soát thì sao?", {"bold": True, "color": ORANGE})], size=16)
    text(s, 0.55, 4.6, 5.9, 1.9, [
        {"runs": "Sales báo cáo một số, Finance báo cáo số khác → lãnh đạo không biết tin số nào",
         "bullet": True},
        {"runs": "Thưởng doanh số, chỉ tiêu cửa hàng tính trên số sai", "bullet": True},
        {"runs": "Lỗi thật (ghi trùng, thiếu dữ liệu, sai thuế) bị che lấp trong tổng", "bullet": True}],
         size=14, space_after=4)
    roles = [("Data Analyst", "Tính KPI từ dữ liệu đơn · ghép với Finance theo tháng × cửa hàng × kênh · "
                              "lập bảng giải thích chênh lệch", GREEN, "bar_w"),
             ("Finance", "Cung cấp số đã xác nhận · xác nhận nguyên nhân phía sổ sách (bút toán trùng, "
                         "thiếu dữ liệu, VAT)", NAVY, "calc_w"),
             ("Cửa hàng / E-commerce", "Xác nhận trạng thái đơn, ngày giao, hoàn trả khi DA hỏi lại",
              ORANGE, "building_b")]
    text(s, 6.85, 1.5, 6.0, 0.4, [("Ba vai trò trong một vòng đối soát", {"bold": True, "color": GREEN})],
         size=16)
    for i, (t, d, colr, icn) in enumerate(roles):
        y = 2.0 + i * 1.5
        box(s, 6.85, y, 6.1, 1.35, fill=WHITE, line=colr, lw=1.5)
        if icn.endswith("_w"):
            ic.circle(s, icn, 7.0, y + 0.3, 0.75, colr)
        else:
            ic.add(s, icn, 7.05, y + 0.32, 0.7)
        text(s, 7.9, y + 0.05, 4.95, 1.25, [[(t, {"bold": True, "color": colr})], d], size=14, anchor="m")

    # ================================================================ PHẦN 2. Đề bài
    SEC = "Đề bài xây dựng dashboard"
    sec_name = "Phần 2 · Đề bài"
    # ---- 6. Tổng quan đề bài
    s = dk.content(sec_name, SEC, "Hai sản phẩm cần nộp, 3 bộ lọc dùng chung và 4 trang dashboard",
                   ["Mỗi trang có ", ("yêu cầu tối thiểu", {"bold": True, "color": RED}), " và ",
                    ("gợi ý mở rộng", {"bold": True, "color": ORANGE}), " ở các slide sau"])
    box(s, 0.4, 1.5, 6.1, 0.65, fill=L_GREEN, line=GREEN, lw=1.5, align="l",
        paras=[[("1  File Power BI", {"bold": True}), (" – dashboard gồm 4 trang", {})]], size=16,
        margin=(0.7, 0.02, 0.1, 0.02))
    ic.add(s, "pie_g", 0.5, 1.58, 0.5)
    box(s, 6.85, 1.5, 6.1, 0.65, fill=L_ORANGE, line=ORANGE, lw=1.5, align="l",
        paras=[[("2  Báo cáo phân tích", {"bold": True}), (" – 8 phần theo guild_report_template", {})]],
        size=16, margin=(0.7, 0.02, 0.1, 0.02))
    ic.add(s, "doc_o", 6.95, 1.58, 0.5)
    text(s, 0.4, 2.3, 6, 0.35, [("Bộ lọc dùng chung cho cả 4 trang", {"bold": True, "color": GREEN})], size=15)
    slic = [("cal_g", "Tháng báo cáo", f"Chọn một · mặc định {RM_LABEL}"),
            ("swap_o", "Kênh bán", "Tất cả · Cửa hàng (POS) · Website/app"),
            ("building_b", "Quốc gia", f"Tất cả hoặc 1 trong {n_ctry} quốc gia + Online")]
    for i, (icn, t, d) in enumerate(slic):
        x = 0.4 + i * 4.2
        box(s, x, 2.7, 4.0, 0.85, fill=WHITE, line=[GREEN, ORANGE, NAVY][i], lw=1.5)
        ic.add(s, icn, x + 0.12, 2.86, 0.52)
        text(s, x + 0.75, 2.72, 3.2, 0.8, [[(t, {"bold": True, "color": [GREEN, ORANGE, NAVY][i]})], d],
             size=13, anchor="m")
    text(s, 0.4, 3.7, 6, 0.35, [("Bốn trang dashboard", {"bold": True, "color": GREEN})], size=15)
    pages = [("grid_w", "Tổng quan", "6 thẻ KPI · Net Revenue đơn vs Finance · số đơn · Gross → Net", GREEN),
             ("pie_w", "Doanh thu", "2 thẻ · bảng kênh · danh mục · quốc gia · AOV · Top 10 cửa hàng", ORANGE),
             ("scale_b", "Đối soát", "4 thẻ · Δ theo tháng · bảng cửa hàng có cờ · Δ POS và online", NAVY),
             ("refresh_w", "Hoàn trả & chất lượng", "4 thẻ · lý do trả · tỷ lệ theo danh mục · huỷ/đang giao",
              RED)]
    for i, (icn, t, d, colr) in enumerate(pages):
        x = 0.4 + i * 3.15
        box(s, x, 4.1, 2.95, 2.45, fill=WHITE, line=colr, lw=1.5)
        if icn.endswith("_w"):
            ic.circle(s, icn, x + 1.07, 4.2, 0.8, colr)
        else:
            ic.circle(s, "swap_w", x + 1.07, 4.2, 0.8, colr)
        text(s, x + 0.1, 5.05, 2.75, 0.4, [(t, {"bold": True, "color": colr})], size=16, align="c")
        text(s, x + 0.12, 5.45, 2.71, 1.05, [d], size=13, align="c", color=TXT)

    # ---- 7–10. Từng trang dashboard
    def page_req(title, sub, kpis, visuals, question, take, colr):
        s = dk.content(sec_name, SEC, sub, take)
        box(s, 0.4, 1.5, 12.55, 0.55, fill=colr, shape="rect",
            paras=[[(title, {"bold": True}), ("   ·   Câu hỏi trang trả lời: ", {}), (question, {"italic": True})]],
            size=15, color=WHITE, align="l", margin=(0.15, 0.02, 0.1, 0.02))
        # thẻ KPI
        text(s, 0.4, 2.15, 12, 0.35, [("Thẻ KPI (mỗi thẻ: giá trị + so với tháng trước + so với cùng kỳ năm trước)",
                                      {"bold": True, "color": colr})], size=13)
        n = len(kpis)
        wk = (12.55 - 0.15 * (n - 1)) / n
        for i, (k, d) in enumerate(kpis):
            x = 0.4 + i * (wk + 0.15)
            box(s, x, 2.55, wk, 1.0, fill=WHITE, line=colr, lw=1.25)
            text(s, x + 0.1, 2.58, wk - 0.2, 0.95, [[(k, {"bold": True})], (d, {"color": GRAY, "size": 12})],
                 size=14, anchor="m")
        text(s, 0.4, 3.7, 4, 0.35, [("Biểu đồ và bảng", {"bold": True, "color": colr})], size=13)
        nv = len(visuals)
        per_row = 3 if nv > 4 else nv
        nrow = (nv + per_row - 1) // per_row
        hv = (2.45 - 0.15 * (nrow - 1)) / nrow
        wv = (12.55 - 0.15 * (per_row - 1)) / per_row
        for i, (icn, t, d) in enumerate(visuals):
            r_, c_ = i // per_row, i % per_row
            x = 0.4 + c_ * (wv + 0.15)
            y = 4.1 + r_ * (hv + 0.15)
            box(s, x, y, wv, hv, fill=WHITE, line=colr, lw=1.25)
            ic.add(s, icn, x + 0.12, y + (hv - 0.5) / 2, 0.5)
            text(s, x + 0.72, y + 0.03, wv - 0.8, hv - 0.06, [[(t, {"bold": True, "color": colr})], d], size=13,
                 anchor="m")
        return s

    page_req("Trang 1 · Tổng quan", "Trang Tổng quan: 6 thẻ KPI và 3 biểu đồ",
             [("Gross Sales", "nghìn USD"), ("Net Revenue", "nghìn USD"), ("Số đơn hoàn tất", "đơn"),
              ("AOV", "USD/đơn"), ("ASP", "USD/sản phẩm"), ("Chênh lệch Finance", "USD · % · số ô vượt")],
             [("trend_g", "Line chart Net Revenue: đơn vs Finance",
               "24 tháng 01/2019–12/2020, hai đường; không bị slicer Tháng lọc"),
              ("trend_b", "Line chart số đơn hoàn tất", "Theo tháng đặt hàng, cùng 24 tháng"),
              ("bar_g", "Waterfall Gross → Net Revenue", "Tháng đang chọn: Gross − Giảm giá − Hoàn trả")],
             "Tháng này bán được bao nhiêu và có khớp Finance không?",
             ["Tổng quan: ", ("6 thẻ KPI", {"bold": True, "color": ORANGE}), " ở trên, ",
              ("3 biểu đồ", {"bold": True, "color": ORANGE}), " ở dưới"], GREEN)
    page_req("Trang 2 · Doanh thu", "Trang Doanh thu: doanh thu đến từ kênh, danh mục, quốc gia nào?",
             [("Net Sales", "nghìn USD"), ("Số sản phẩm bán (Units)", "sản phẩm")],
             [("table_g", "Bảng doanh thu theo kênh", "Net Sales, tỷ trọng, số đơn, tháng trước, cùng kỳ, thay đổi"),
              ("pie_o", "Pie cơ cấu danh mục", "Net Sales theo category, tối đa 7 lát + Khác"),
              ("tree_t", "Treemap quốc gia", "Net Sales theo country (Online là một ô)"),
              ("bar_o", "Bar AOV theo kênh", "POS và Website/app"),
              ("table_r", "Bảng Top 10 cửa hàng", "Net Sales, số đơn, tháng trước, cùng kỳ, ▲/▼")],
             "Doanh thu tập trung ở đâu, kênh nào tăng?",
             ["Doanh thu: ", ("bảng kênh", {"bold": True, "color": ORANGE}), ", 3 biểu đồ cơ cấu và ",
              ("Top 10 cửa hàng", {"bold": True, "color": ORANGE})], ORANGE)
    page_req("Trang 3 · Đối soát", "Trang Đối soát: số của mình lệch Finance bao nhiêu, ở đâu?",
             [("Net Revenue (đơn)", "nghìn USD"), ("Net Revenue (Finance)", "nghìn USD"),
              ("Chênh lệch Δ", "USD · Δ% · Δ sau quy đổi online"), ("Số ô vượt ngưỡng", "tháng × cửa hàng")],
             [("trend_b", "Line chart Δ theo tháng", "Δ gốc và Δ sau khi quy online về ngày giao"),
              ("table_b", "Bảng đối soát theo cửa hàng", "Đơn, Finance, Δ, Δ%, Δ sau quy đổi, cờ Vượt"),
              ("trend_r", "Line chart Δ kênh POS", "Theo tháng, USD"),
              ("trend_g", "Line chart Δ kênh online", "Gốc và sau quy về delivery_date")],
             "Chênh lệch nằm ở cửa hàng nào, tháng nào, kênh nào?",
             ["Đối soát: ", ("Δ = Finance − dữ liệu đơn", {"bold": True, "color": NAVY}), "; cờ khi |Δ| > ",
              (vn(THR_ABS) + " USD", {"bold": True, "color": RED}), " và |Δ%| > ",
              (pct(THR_PCT, 0), {"bold": True, "color": RED})], NAVY)
    page_req("Trang 4 · Hoàn trả & chất lượng", "Trang Hoàn trả & chất lượng: bao nhiêu doanh thu bị mất?",
             [("Giá trị hoàn trả", "nghìn USD"), ("Tỷ lệ hoàn trả", "% · thay đổi tính bằng điểm %"),
              ("Đơn huỷ", "đơn"), ("Dòng trùng đã loại", "dòng · kèm số đơn PENDING")],
             [("pie_r", "Pie giá trị hoàn trả theo lý do", "Lỗi sản phẩm, giao nhầm, giao trễ, đổi ý"),
              ("table_r", "Bảng tỷ lệ hoàn trả theo danh mục", "Net Sales, hoàn trả, tỷ lệ, tháng trước, cùng kỳ"),
              ("trend_r", "Line chart tỷ lệ hoàn trả", "Theo tháng, %"),
              ("trend_b", "Line chart đơn huỷ và đang giao", "Theo tháng đặt hàng")],
             "Doanh thu có \"sạch\" không: trả hàng, huỷ đơn, dữ liệu lỗi?",
             ["Tỷ lệ hoàn trả tăng là ", ("tín hiệu xấu", {"bold": True, "color": RED}),
              ": hiển thị màu đảo, thay đổi bằng ", ("điểm %", {"bold": True, "color": NAVY})], RED)

    # ---- 11. Checklist
    s = dk.content(sec_name, SEC, "Bảng tự rà soát mức độ hoàn thiện sản phẩm",
                   ["Tự rà soát: ", ("Bắt buộc", {"bold": True, "color": RED}), " → ",
                    ("Nên có", {"bold": True, "color": ORANGE}), " → ", ("Sáng tạo", {"bold": True, "color": GREEN})])
    chk = [("Bắt buộc", RED, [f"Bỏ {vn(REF['duplicates'])} dòng trùng, chỉ tính đơn COMPLETED",
                              "4 trang đủ thẻ KPI, biểu đồ, bảng theo yêu cầu",
                              "3 slicer dùng chung, đồng bộ cả 4 trang",
                              "Thẻ KPI so với tháng trước và cùng tháng năm trước",
                              "Map Finance theo valid_from / valid_to, tách HQ-ADJ",
                              "Bảng đối soát có cờ vượt ngưỡng"]),
           ("Nên có", ORANGE, ["Δ online sau khi quy về delivery_date",
                               "Line chart không bị slicer Tháng lọc (đủ 24 tháng)",
                               "Caption \"Hình N\" cho mỗi biểu đồ, bảng",
                               "Ghi rõ đơn vị: USD, nghìn USD, %, điểm %",
                               "Nút Đặt lại và điều hướng giữa 4 trang",
                               "Màu đảo cho hoàn trả, ▲/▼ ở cột thay đổi"]),
           ("Sáng tạo", GREEN, ["Drill-through \"Chi tiết cửa hàng – ngày\"",
                                "Tooltip Δ theo ngày khi rê chuột vào ô vượt ngưỡng",
                                "Bảng giải thích chênh lệch: tổng nguyên nhân = Δ",
                                "Dùng fx_rate, unit_cost, dim_customer (bảng bổ sung)"])]
    for i, (t, colr, items) in enumerate(chk):
        x = 0.4 + i * 4.2
        box(s, x, 1.5, 4.0, 5.05, fill=WHITE, line=colr, lw=1.5)
        box(s, x, 1.5, 4.0, 0.55, fill=colr, shape="round", adj=0.15, paras=[(t, {"bold": True})], size=17,
            color=WHITE)
        text(s, x + 0.12, 2.15, 3.8, 4.3,
             [{"runs": it, "bullet": "☐", "indent": 0.3, "space_after": 7} for it in items], size=14)

    # ================================================================ PHẦN 3. KPI
    SEC = "Hệ thống chỉ số KPI"
    sec_name = "Phần 3 · KPI"
    # ---- 12. Cây KPI
    s = dk.content(sec_name, SEC, "Cây KPI: từ Gross Sales đến chênh lệch với Finance",
                   ["Hiểu ", ("định nghĩa và thứ tự", {"bold": True, "color": NAVY}),
                    " trước, viết DAX sau: Gross → Net Sales → Net Revenue → Δ Finance"])
    picture_fit(s, DIAG / "04_cay_kpi.png", 0.4, 1.45, 12.55, 5.15)

    # ---- 13. Nhóm doanh thu + ví dụ tính tay
    smp = D["sample"]
    ex = smp[smp["case"] == "POS có khuyến mãi"]
    ex_no = int(ex.order_number.iloc[0])
    ex_g = float((ex.quantity * ex.unit_price).sum())
    ex_d = float(ex.discount_amount.sum())
    ex_promo = ex.promo_code.iloc[0]
    ex_rate = D["X"]["dim_promotion"].set_index("promo_code").discount_pct.get(ex_promo, np.nan)
    s = dk.content(sec_name, SEC, "Nhóm doanh thu: định nghĩa, công thức và một ví dụ tính tay",
                   [("Net Revenue", {"bold": True, "color": ORANGE}),
                    " là con số đem đi đối soát; Net Sales là mẫu số chung của AOV, ASP, tỷ lệ hoàn trả"])
    kdefs = [("Gross Sales", "Σ quantity × unit_price", "Đơn COMPLETED, đã bỏ dòng trùng", GREEN),
             ("Discounts", "Σ discount_amount", "Khuyến mãi theo promo_code", ORANGE),
             ("Net Sales", "Gross Sales − Discounts", "Doanh thu sau giảm giá", GREEN),
             ("Returns", "Σ return_amount theo return_date", "Trừ vào tháng trả hàng, không phải tháng bán", RED),
             ("Net Revenue", "Net Sales − Returns", "Con số so với Finance", NAVY)]
    for i, (k, f, d, colr) in enumerate(kdefs):
        y = 1.5 + i * 1.02
        box(s, 0.4, y, 7.3, 0.9, fill=WHITE, line=colr, lw=1.25)
        box(s, 0.4, y, 1.9, 0.9, fill=colr, paras=[(k, {"bold": True})], size=15, color=WHITE)
        box(s, 2.42, y + 0.1, 2.9, 0.7, fill=L_GRAY, adj=0.1, paras=[(f, {"bold": True})], size=13, color=NAVY)
        text(s, 5.4, y + 0.05, 2.25, 0.8, [d], size=12, anchor="m", color=TXT)
    box(s, 7.95, 1.5, 5.0, 5.05, fill=L_ORANGE, line=ORANGE, lw=1.5)
    box(s, 7.95, 1.5, 5.0, 0.5, fill=ORANGE, adj=0.15,
        paras=[("Ví dụ tính tay (sheet sample_transactions)", {"bold": True})], size=14, color=WHITE)
    text(s, 8.1, 2.08, 4.7, 0.8, [f"Đơn {ex_no} · cửa hàng POS · {len(ex)} dòng · mã {ex_promo} "
                                  f"(giảm {pct(ex_rate, 0)})"], size=13, color=GRAY)
    calc = [("Gross Sales", "Σ quantity × unit_price", vn(ex_g, 2), GREEN),
            ("− Discounts", "Σ discount_amount", vn(ex_d, 2), ORANGE),
            ("= Net Sales", "Gross − Discounts", vn(ex_g - ex_d, 2), NAVY)]
    for i, (a, b_, v, colr) in enumerate(calc):
        y = 2.75 + i * 0.95
        box(s, 8.1, y, 4.7, 0.82, fill=WHITE, line=colr, lw=1.25)
        text(s, 8.2, y + 0.03, 2.4, 0.76, [[(a, {"bold": True, "color": colr})], (b_, {"size": 11, "color": GRAY})],
             size=15, anchor="m")
        text(s, 10.5, y + 0.03, 2.2, 0.76, [(v + " USD", {"bold": True})], size=17, align="r", anchor="m",
             color=colr)
    text(s, 8.1, 5.6, 4.7, 0.9, [f"Kiểm tra: {vn(ex_d, 2)} / {vn(ex_g, 2)} = {pct(ex_d / ex_g, 1)} ≈ "
                                 f"{pct(ex_rate, 0)} khuyến mãi. Đơn này không có hoàn trả nên Net Revenue = Net "
                                 f"Sales."], size=12, color=TXT)

    # ---- 14. Đơn, giá, tỷ lệ + mốc thời gian
    s = dk.content(sec_name, SEC, "Nhóm đơn hàng và giá: Orders, Units, AOV, ASP, Return Rate",
                   ["Orders ", ("đếm đơn, không đếm dòng", {"bold": True, "color": RED}),
                    "; AOV, ASP và tỷ lệ hoàn trả đều chia cho Net Sales"])
    k2 = [("Orders", "DISTINCTCOUNT(order_number)", "Đơn COMPLETED", vn(m["orders"]) + " đơn", NAVY),
          ("Units", "Σ quantity", "Đơn COMPLETED", vn(m["units"]) + " sản phẩm", NAVY),
          ("AOV", "Net Sales / Orders", "Giá trị trung bình mỗi đơn", vn(m["aov"], 2) + " USD", GREEN),
          ("ASP", "Net Sales / Units", "Giá bán trung bình mỗi sản phẩm", vn(m["asp"], 2) + " USD", GREEN),
          ("Return Rate", "Returns / Net Sales", "Tăng là xấu, đo bằng điểm %", pct(m["return_rate"], 2), RED)]
    rws = [["KPI", "Công thức", "Ghi chú", f"Giá trị {RM_LABEL}"]]
    for k, f, d, v, _ in k2:
        rws.append([k, f, d, v])
    table(s, 0.4, 1.5, 7.9, 3.3, rws, col_w=[1.25, 3.05, 2.2, 1.6], size=12, bold_first_col=True,
          align=["l", "l", "l", "r"])
    text(s, 0.4, 4.9, 7.9, 0.35, [(f"Kiểm chứng: AOV = {vn(m['net_sales'], 2)} / {vn(m['orders'])} = "
                                   f"{vn(m['aov'], 2)} USD/đơn", {"italic": True})], size=12, color=GRAY)
    box(s, 8.55, 1.5, 4.4, 5.05, fill=L_BLUE, line=NAVY, lw=1.5)
    box(s, 8.55, 1.5, 4.4, 0.5, fill=NAVY, adj=0.15, paras=[("Bốn mốc thời gian", {"bold": True})], size=15,
        color=WHITE)
    tl = [("Tháng báo cáo", RM_LABEL, "chọn trên slicer"), ("Tháng trước (LM)", LM_LABEL, "so sánh gần nhất"),
          ("Cùng kỳ năm trước (SPLY)", SPLY_LABEL, "loại bỏ yếu tố mùa vụ"),
          ("Lũy kế năm (YTD)", f"01–{RM_LABEL}", "cộng dồn từ tháng 1")]
    for i, (a, b_, c_) in enumerate(tl):
        y = 2.15 + i * 0.95
        box(s, 8.7, y, 0.45, 0.45, fill=NAVY, shape="oval", paras=[(str(i + 1), {"bold": True})], size=14,
            color=WHITE, margin=(0, 0, 0, 0))
        text(s, 9.25, y - 0.08, 3.6, 0.9, [[(a, {"bold": True, "color": NAVY}), ("  " + b_, {"bold": True})],
                                            (c_, {"color": GRAY, "size": 12})], size=14)
    text(s, 8.7, 5.95, 4.15, 0.55, ["Δ% = (Hiện tại − Kỳ so sánh) / Kỳ so sánh"], size=13, bold=True,
         color=NAVY)
    box(s, 0.4, 5.35, 7.9, 1.2, fill=L_ORANGE, line=ORANGE, lw=1.25, align="l", anchor="m",
        paras=[[("Bẫy thường gặp: ", {"bold": True, "color": ORANGE}),
                ("đếm COUNT(order_number) trên bảng dòng → số đơn bị nhân theo số sản phẩm trong đơn; "
                 "tính AOV trên Gross thay vì Net Sales → AOV cao hơn thực tế.", {})]], size=13,
        margin=(0.15, 0.04, 0.15, 0.04))

    # ---- 15. Quy tắc trạng thái + ghi nhận
    so = D["status_orders"]
    stt = D["X"]["dim_order_status"]
    s = dk.content(sec_name, SEC, "Quy tắc trạng thái đơn và quy tắc ghi nhận của Finance",
                   ["Chỉ ", ("COMPLETED", {"bold": True, "color": GREEN}), " vào doanh thu; Finance ghi online theo ",
                    ("ngày giao", {"bold": True, "color": NAVY}), ", hoàn tiền theo ",
                    ("ngày hạch toán", {"bold": True, "color": NAVY})])
    text(s, 0.4, 1.45, 6, 0.35, [("Trạng thái đơn (dim_order_status)", {"bold": True, "color": GREEN})], size=15)
    st_col = {"COMPLETED": GREEN, "CANCELLED": RED, "PENDING": ORANGE}
    for i, rr in enumerate(stt.itertuples()):
        y = 1.9 + i * 1.12
        colr = st_col.get(rr.order_status, NAVY)
        box(s, 0.4, y, 6.1, 1.0, fill=WHITE, line=colr, lw=1.25)
        box(s, 0.4, y, 1.85, 1.0, fill=colr, paras=[[(rr.order_status, {"bold": True})],
                                                     (f"{vn(int(so.get(rr.order_status, 0)))} đơn", {"size": 12})],
            size=14, color=WHITE)
        text(s, 2.35, y + 0.03, 4.1, 0.95, [[("Tính doanh thu: ", {"bold": True}),
                                              (rr.count_in_revenue, {"bold": True, "color": colr})], rr.rule],
             size=12, anchor="m")
    text(s, 0.4, 5.3, 6.1, 1.2, [f"Tổng {vn(rows['orders_total'])} đơn = {vn(int(so.get('COMPLETED', 0)))} + "
                                 f"{vn(int(so.get('CANCELLED', 0)))} + {vn(int(so.get('PENDING', 0)))}. "
                                 f"PENDING là đơn online chưa giao tại ngày trích {AS_OF_LABEL}."],
         size=13, color=GRAY)
    text(s, 6.85, 1.45, 6, 0.35, [("Finance ghi nhận khi nào?", {"bold": True, "color": NAVY})], size=15)
    ch_rule = D["X"]["dim_channel"]
    rec = [(r_.channel_code + " · " + r_.fin_channel, r_.revenue_recognition_rule) for r_ in ch_rule.itertuples()]
    rec.append(("Hoàn tiền", "Finance hạch toán vào ngày xử lý, có thể trễ vài ngày sau return_date"))
    rec.append(("VAT", f"net_revenue chưa gồm VAT; vat_amount = {pct(REF['vat_rate'], 0)} × net_revenue"))
    for i, (a, b_) in enumerate(rec):
        y = 1.9 + i * 1.15
        box(s, 6.85, y, 6.1, 1.03, fill=L_BLUE if i % 2 == 0 else WHITE, line=NAVY, lw=1.0)
        text(s, 7.0, y + 0.03, 5.85, 0.98, [[(a, {"bold": True, "color": NAVY})], b_], size=13, anchor="m")

    # ---- 16. Chênh lệch và ngưỡng
    bsm = K["bridge_store_month"]
    rm_rows = bsm[bsm.month_id == REPORT_MONTH].copy()
    rm_rows["absd"] = rm_rows.diff_finance_minus_orders.abs()
    # ví dụ: 1 ô trong ngưỡng có Δ ≠ 0 và 1 ô vượt ngưỡng (cửa hàng POS) trong tháng báo cáo
    ins = rm_rows[(rm_rows.absd > 0.005) & (~rm_rows.over_threshold) & (rm_rows.store_key > 0)].sort_values(
        "absd", ascending=False).iloc[0]
    out = rm_rows[(rm_rows.over_threshold) & (rm_rows.store_key > 0)].sort_values("absd", ascending=False).iloc[0]
    s = dk.content(sec_name, SEC, "Chênh lệch Δ và ngưỡng cho phép: khi nào phải giải thích?",
                   ["Chỉ ô ", ("tháng × cửa hàng", {"bold": True, "color": NAVY}), " vượt ", ("cả hai", {"bold": True,
                                                                                                   "color": RED}),
                    " điều kiện mới bắt buộc giải thích"])
    fm = [("Δ Finance", "Finance Net Revenue − Net Revenue (đơn)", "Dương: Finance ghi cao hơn dữ liệu đơn"),
          ("Δ %", "Δ / |Finance Net Revenue|", "Chia cho trị tuyệt đối để giữ dấu của Δ"),
          ("Cờ vượt ngưỡng", f"|Δ| > {vn(THR_ABS)} USD  VÀ  |Δ%| > {pct(THR_PCT, 0)}",
           "Xét ở mức tháng × cửa hàng (store_key 0 = online)")]
    for i, (a, b_, c_) in enumerate(fm):
        y = 1.5 + i * 1.12
        box(s, 0.4, y, 7.0, 1.0, fill=WHITE, line=NAVY, lw=1.25)
        box(s, 0.4, y, 1.9, 1.0, fill=NAVY if i < 2 else RED, paras=[(a, {"bold": True})], size=15, color=WHITE)
        text(s, 2.4, y + 0.03, 4.95, 0.95, [[(b_, {"bold": True})], (c_, {"color": GRAY, "size": 12})], size=14,
             anchor="m")
    box(s, 0.4, 4.95, 7.0, 1.6, fill=L_GREEN, line=GREEN, lw=1.25, align="l", anchor="t",
        paras=[[("Vì sao cần cả hai điều kiện?", {"bold": True, "color": GREEN})],
               {"runs": "Cửa hàng lớn: lệch 300 USD có thể chỉ 0,1% → làm tròn, không đáng giải thích",
                "bullet": True},
               {"runs": "Cửa hàng nhỏ: lệch 5% nhưng chỉ 50 USD → không đáng công truy vết", "bullet": True}],
        size=13, margin=(0.15, 0.06, 0.15, 0.04))
    box(s, 7.65, 1.5, 5.3, 5.05, fill=L_ORANGE, line=ORANGE, lw=1.5)
    box(s, 7.65, 1.5, 5.3, 0.5, fill=ORANGE, adj=0.15, paras=[(f"Ví dụ thật, tháng {RM_LABEL}", {"bold": True})],
        size=15, color=WHITE)
    for i, (rr, verdict, colr) in enumerate([(ins, "Trong ngưỡng: ghi nhận, không bắt buộc giải thích", GREEN),
                                             (out, "Vượt ngưỡng: phải tìm nguyên nhân và số tiền", RED)]):
        y = 2.15 + i * 2.15
        box(s, 7.8, y, 5.0, 2.0, fill=WHITE, line=colr, lw=1.25)
        text(s, 7.95, y + 0.05, 4.75, 1.9, [
            [(f"Cửa hàng {rr.store_code} ({rr.country})", {"bold": True, "color": colr})],
            f"Đơn: {vn(rr.orders_net_revenue, 2)} · Finance: {vn(rr.finance_net, 2)}",
            [("Δ = ", {}), (vn(rr.diff_finance_minus_orders, 2, True) + " USD", {"bold": True}),
             (" · Δ% = ", {}), (pct(rr.diff_pct, 1, True), {"bold": True})],
            [(verdict, {"italic": True, "color": colr})]], size=13, space_after=3)

    # ================================================================ PHẦN 4. Dữ liệu
    SEC = "Nguồn dữ liệu và data modeling"
    sec_name = "Phần 4 · Dữ liệu & modeling"
    # ---- 17. Pipeline
    s = dk.content(sec_name, SEC, "Pipeline dữ liệu: từ file nguồn đến dashboard và báo cáo",
                   ["Học viên làm 5 bước: ", ("Power Query → Mô hình → DAX → Dashboard → Báo cáo", {"bold": True,
                                                                                                   "color": GREEN})])
    picture_fit(s, DIAG / "01_pipeline_du_lieu.png", 0.4, 1.45, 12.55, 5.15)

    # ---- 18. Cấu trúc bộ dữ liệu + bảng fact
    s = dk.content(sec_name, SEC, "Cấu trúc bộ dữ liệu Nexa Home và hai bảng fact",
                   ["Mỗi bảng có ", ("hạt dữ liệu (grain)", {"bold": True, "color": RED}),
                    " khác nhau: dòng sản phẩm, lần trả hàng, ngày × mã cửa hàng Finance"])
    stat = [("2", "file", "Dataset + Finance"), ("2 + 1", "bảng fact", "đơn · trả hàng · Finance"),
            ("7", "bảng dim", "ngày, cửa hàng, sản phẩm…"), ("1", "bảng map", "mã Finance → store_key")]
    for i, (v, a, b_) in enumerate(stat):
        x = 0.4 + i * 3.15
        box(s, x, 1.5, 2.95, 1.05, fill=WHITE, line=[GREEN, ORANGE, NAVY, RED][i], lw=1.5)
        text(s, x, 1.5, 1.1, 1.05, [(v, {"bold": True})], size=26, color=[GREEN, ORANGE, NAVY, RED][i],
             align="c", anchor="m")
        text(s, x + 1.05, 1.52, 1.85, 1.0, [[(a, {"bold": True})], (b_, {"color": GRAY, "size": 12})], size=14,
             anchor="m")
    facts = [("fact_sales_line", vn(rows["fact_sales_line"]) + " dòng",
              "Một dòng = một sản phẩm trong một đơn", "order_number · line_item · order_date · delivery_date · "
              "store_key · channel_code · product_key · quantity · unit_price · promo_code · discount_amount · "
              "order_status", f"Có {vn(REF['duplicates'])} dòng trùng hoàn toàn cần loại", GREEN),
             ("fact_return", vn(rows["fact_return"]) + " dòng", "Một dòng = một lần trả hàng của một dòng bán",
              "return_id · order_number · line_item · return_date · return_qty · return_amount · return_reason",
              "Không có store_key: lấy qua order_number + line_item", ORANGE),
             ("finance_revenue", vn(rows["finance_revenue"]) + " dòng",
              "Một dòng = một ngày hạch toán × một mã cửa hàng Finance",
              "posting_date · fin_store_code · fin_channel · entry_type · revenue_recognized · refunds · "
              "net_revenue · vat_amount", "fin_store_code ≠ store_key: phải map", NAVY)]
    for i, (t, n_, g, f, warn, colr) in enumerate(facts):
        x = 0.4 + i * 4.2
        y = 2.75
        box(s, x, y, 4.0, 3.8, fill=WHITE, line=colr, lw=1.5)
        box(s, x, y, 4.0, 0.5, fill=colr, adj=0.15, align="l",
            paras=[[(t, {"bold": True}), ("   " + n_, {"size": 12})]], size=15, color=WHITE,
            margin=(0.12, 0.02, 0.1, 0.02))
        text(s, x + 0.1, y + 0.55, 3.8, 3.2, [
            [("Hạt dữ liệu", {"bold": True, "color": colr})], g,
            {"runs": [("Trường chính", {"bold": True, "color": colr})], "space_before": 6},
            {"runs": f, "size": 12},
            {"runs": [("Lưu ý: ", {"bold": True, "color": RED}), (warn, {})], "space_before": 6}], size=13)

    # ---- 19. Bảng dim
    s = dk.content(sec_name, SEC, "Các bảng dim và bảng hỗ trợ: dùng để lọc, nhóm và quy đổi",
                   ["dim_date phải được ", ("Mark as date table", {"bold": True, "color": NAVY}),
                    " để dùng DATEADD, SAMEPERIODLASTYEAR, DATESYTD"])
    dims = [("dim_date", vn(rows["dim_date"]), "Một dòng = một ngày", "date · month_id · month_label · year · "
                                                                    "has_data", "cal_g"),
            ("dim_store", vn(rows["dim_store"]), "Một dòng = một cửa hàng (0 = Online)",
             "store_key · store_code · channel_code · country · region", "building_b"),
            ("dim_product", vn(rows["dim_product"]), "Một dòng = một sản phẩm",
             "product_key · category · subcategory · brand", "layers_b"),
            ("dim_customer", vn(rows["dim_customer"]), "Bảng bổ sung (sáng tạo)",
             "customer_key · country · continent", "people_b"),
            ("dim_channel", vn(len(D["X"]["dim_channel"])), "POS / ONLINE + quy tắc ghi nhận",
             "channel_code · fin_channel · revenue_recognition_rule", "swap_o"),
            ("dim_promotion", vn(len(D["X"]["dim_promotion"])), "Mã khuyến mãi và % giảm",
             "promo_code · promo_name · discount_pct", "money_o"),
            ("dim_order_status", vn(len(stt)), "Trạng thái và có tính doanh thu không",
             "order_status · count_in_revenue", "check_o"),
            ("map_store_finance", vn(rows["map_store_finance"]), "Mã Finance → store_key theo hiệu lực",
             "fin_store_code · store_key · valid_from · valid_to", "scale_b")]
    for i, (t, n_, g, f, icn) in enumerate(dims):
        r_, c_ = i // 4, i % 4
        x = 0.4 + c_ * 3.15
        y = 1.5 + r_ * 2.6
        colr = RED if t == "map_store_finance" else (MUTED if t == "dim_customer" else GREEN)
        box(s, x, y, 2.95, 2.45, fill=WHITE, line=colr, lw=1.5)
        box(s, x, y, 2.95, 0.5, fill=colr, adj=0.15, paras=[[(t, {"bold": True}), (" · " + n_, {"size": 12})]],
            size=14, color=WHITE, margin=(0.05, 0.02, 0.05, 0.02))
        ic.add(s, icn, x + 0.1, y + 0.62, 0.45)
        text(s, x + 0.62, y + 0.55, 2.28, 0.65, [(g, {"bold": True})], size=12, anchor="m")
        text(s, x + 0.1, y + 1.25, 2.75, 1.15, [f], size=12, color=GRAY)

    # ---- 20. Mô hình
    s = dk.content(sec_name, SEC, "Mô hình dữ liệu dạng star schema",
                   ["Quan hệ ", ("1 → nhiều, lọc một chiều", {"bold": True, "color": GREEN}),
                    "; delivery_date là quan hệ ", ("inactive", {"bold": True, "color": ORANGE}),
                    ", bật bằng USERELATIONSHIP"])
    picture_fit(s, DIAG / "02_mo_hinh_du_lieu.png", 0.4, 1.42, 8.3, 5.2)
    box(s, 8.9, 1.5, 4.05, 5.05, fill=L_BLUE, line=NAVY, lw=1.5)
    box(s, 8.9, 1.5, 4.05, 0.5, fill=NAVY, adj=0.15, paras=[("Cách dựng trong Power BI", {"bold": True})],
        size=15, color=WHITE)
    text(s, 9.0, 2.1, 3.85, 4.4, [
        {"runs": [("fact_sales_line[order_date]", {"bold": True}), (" → dim_date: active", {})], "bullet": True},
        {"runs": [("fact_sales_line[delivery_date]", {"bold": True}), (" → dim_date: inactive", {})],
         "bullet": True},
        {"runs": [("fact_return[return_date]", {"bold": True}), (" → dim_date: active", {})], "bullet": True},
        {"runs": "fact_return: merge store_key, channel_code, product_key từ fact_sales_line trong Power Query",
         "bullet": True},
        {"runs": "fact_finance: map store_key trong Power Query, nối dim_date theo posting_date", "bullet": True},
        {"runs": "Không nối 2 bảng fact trực tiếp với nhau", "bullet": True}], size=13, space_after=6)

    # ---- 21. map_store_finance & bẫy
    mp = D["map"]
    rec_rows = mp[mp.store_key == REF["recode"]["store_key"]]
    hq = mp[mp.store_key.isna()]
    s = dk.content(sec_name, SEC, "map_store_finance: quy đổi mã cửa hàng Finance theo ngày hiệu lực",
                   ["Join chỉ theo fin_store_code là ", ("chưa đủ", {"bold": True, "color": RED}),
                    ": phải thêm điều kiện posting_date nằm trong [valid_from, valid_to]"])
    rws = [["fin_store_code", "store_key", "valid_from", "valid_to", "Ghi chú"]]
    for r_ in pd.concat([mp.head(2), rec_rows, hq]).itertuples():
        rws.append([r_.fin_store_code, "—" if pd.isna(r_.store_key) else str(int(r_.store_key)),
                    r_.valid_from.strftime("%d/%m/%Y"), r_.valid_to.strftime("%d/%m/%Y"),
                    "" if pd.isna(r_.note) else r_.note])
    fills = {}
    for ri in range(1, len(rws)):
        if rws[ri][0] in (REF["recode"]["old"], REF["recode"]["new"]):
            fills[(ri, None)] = L_ORANGE
        if rws[ri][0] == "HQ-ADJ":
            fills[(ri, None)] = L_BLUE
    table(s, 0.4, 1.5, 12.55, 2.5, rws, col_w=[1.6, 1.1, 1.3, 1.3, 6.0], size=13, fills=fills,
          align=["l", "c", "c", "c", "l"], zebra=False)
    cards = [("Bẫy 1 · Đổi mã cửa hàng", ORANGE,
              f"Cửa hàng {REF['recode']['store_key']} dùng mã {REF['recode']['old']} đến 30/06/2019, sau đó "
              f"{REF['recode']['new']}. Map không theo hiệu lực → doanh thu từ 07/2019 thành \"không xác định "
              "cửa hàng\" hoặc bị nhân đôi."),
             ("Bẫy 2 · Dòng HQ-ADJ", NAVY,
              "Bút toán điều chỉnh tại văn phòng chính, store_key trống. Không gán vào cửa hàng nào; giải thích ở "
              "cấp tổng chuỗi."),
             ("Cách làm trong Power Query", GREEN,
              "Merge finance với map theo fin_store_code → lọc posting_date ≥ valid_from và ≤ valid_to → kiểm "
              f"tra số dòng vẫn = {vn(rows['finance_revenue'])}.")]
    for i, (t, colr, d) in enumerate(cards):
        x = 0.4 + i * 4.2
        box(s, x, 4.2, 4.0, 2.35, fill=WHITE, line=colr, lw=1.5)
        box(s, x, 4.2, 4.0, 0.5, fill=colr, adj=0.15, paras=[(t, {"bold": True})], size=15, color=WHITE)
        text(s, x + 0.12, 4.75, 3.76, 1.75, [d], size=13)

    # ---- 22. Quy tắc làm sạch & QC
    s = dk.content(sec_name, SEC, "Quy tắc làm sạch và số dòng phải khớp sau khi nạp (QC_Summary)",
                   ["Chưa khớp số dòng thì ", ("chưa tính KPI", {"bold": True, "color": RED}),
                    ": mọi sai lệch phía sau đều bắt đầu từ bước này"])
    qc = [["Kiểm tra", "Kết quả cần khớp", "Ghi chú"],
          ["fact_sales_line nạp vào", vn(rows["fact_sales_line"]) + " dòng", "Gồm dòng trùng cài sẵn"],
          ["Bỏ dòng trùng hoàn toàn", "− " + vn(REF["duplicates"]) + " dòng",
           "Còn " + vn(rows["fact_sales_line_clean"]) + " dòng; khoá order_number + line_item duy nhất"],
          ["Tổng số đơn", vn(rows["orders_total"]) + " đơn", "Đếm phân biệt order_number"],
          ["Đơn COMPLETED", vn(rows["orders_completed"]) + " đơn", "Chỉ nhóm này vào KPI doanh thu"],
          ["Đơn CANCELLED / PENDING", vn(rows["orders_cancelled"]) + " / " + vn(rows["orders_pending"]) + " đơn",
           "Loại khỏi KPI"],
          ["fact_return", vn(rows["fact_return"]) + " dòng", "Mỗi dòng khớp 1 dòng bán COMPLETED"],
          ["finance_revenue", vn(rows["finance_revenue"]) + " dòng", "Sau khi map vẫn đủ số dòng, không nhân đôi"],
          ["dim_store / dim_product / dim_date", f"{vn(rows['dim_store'])} / {vn(rows['dim_product'])} / "
                                                 f"{vn(rows['dim_date'])}", "dim_date phủ 01/01/2016 → 31/03/2021"]]
    table(s, 0.4, 1.5, 8.4, 5.05, qc, col_w=[2.9, 2.0, 3.5], size=13, align=["l", "r", "l"], bold_first_col=True)
    box(s, 9.05, 1.5, 3.9, 5.05, fill=L_ORANGE, line=ORANGE, lw=1.5)
    box(s, 9.05, 1.5, 3.9, 0.5, fill=ORANGE, adj=0.15, paras=[("Ba quy tắc làm sạch", {"bold": True})], size=15,
        color=WHITE)
    text(s, 9.15, 2.1, 3.7, 4.4, [
        {"runs": [("1. Bỏ trùng", {"bold": True, "color": ORANGE})]},
        {"runs": "Power Query → Remove Duplicates trên toàn bộ cột", "space_after": 8},
        {"runs": [("2. Lọc trạng thái", {"bold": True, "color": ORANGE})]},
        {"runs": "Giữ dữ liệu đủ cột, lọc COMPLETED trong measure để vẫn đếm được đơn huỷ", "space_after": 8},
        {"runs": [("3. Đúng kiểu dữ liệu", {"bold": True, "color": ORANGE})]},
        {"runs": "Ngày là Date, tiền là Decimal, store_key là Whole Number ở mọi bảng"}], size=13)

    # ================================================================ PHẦN 5. EDA
    SEC = "Khám phá dữ liệu"
    sec_name = "Phần 5 · Khám phá dữ liệu"
    years = sorted(REF["year"].keys())
    # ---- 23. Theo năm
    s = dk.content(sec_name, SEC, f"Doanh thu theo năm: dữ liệu đơn và Finance, 2016–{years[-1]}",
                   ["Hai nguồn ", ("gần khớp ở cấp năm", {"bold": True, "color": GREEN}),
                    "; chênh lệch chỉ lộ ra khi xuống tháng × cửa hàng"])
    cd = CategoryChartData()
    cats = [y if y != years[-1] else f"{y}*" for y in years]
    cd.categories = cats
    nr_y = [REF["year"][y]["net_revenue"] / 1e6 for y in years]
    fin_y = [REF["finance_year"][y] / 1e6 for y in years]
    cd.add_series("Net Revenue (dữ liệu đơn)", nr_y)
    cd.add_series("Net Revenue (Finance)", fin_y)
    chart_frame(s, 0.4, 1.5, 8.3, 5.05, "Net Revenue theo năm (triệu USD)",
                caption=f"* {years[-1]}: chỉ đến {AS_OF_LABEL}")
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, I(0.5), I(1.9), I(8.1), I(4.3), cd)
    ch = gf.chart
    style_chart(ch)
    style_axes(ch, num_fmt="0", vmin=0)
    color_series(ch.series[0], GREEN)
    color_series(ch.series[1], CH_GREEN2)
    ch.plots[0].gap_width = 60
    ch.plots[0].overlap = -10
    point_labels(ch.series[0], [vn(v, 1) for v in nr_y], size=11)
    point_labels(ch.series[1], [vn(v, 1) for v in fin_y], size=11, color=GRAY)
    y19, y20 = REF["year"][str(FOCUS_YEAR - 1)], REF["year"][str(FOCUS_YEAR)]
    box(s, 8.95, 1.5, 4.0, 5.05, fill=WHITE, line=GREEN, lw=1.5)
    box(s, 8.95, 1.5, 4.0, 0.5, fill=GREEN, adj=0.15, paras=[("Đọc nhanh", {"bold": True})], size=15, color=WHITE)
    text(s, 9.05, 2.1, 3.8, 4.4, [
        {"runs": [(f"{FOCUS_YEAR}: ", {"bold": True}), (vn(y20["net_revenue"] / 1e6, 2) + " triệu USD", {}),
                  (f", {pct(chg(y20['net_revenue'], y19['net_revenue']), 1, True)} so với {FOCUS_YEAR - 1}",
                   {"bold": True, "color": RED})], "bullet": True, "space_after": 6},
        {"runs": [("Số đơn: ", {"bold": True}), (f"{vn(y19['orders'])} → {vn(y20['orders'])}", {})],
         "bullet": True, "space_after": 6},
        {"runs": [("AOV gần như đi ngang: ", {"bold": True}),
                  (f"{vn(y19['aov'], 0)} → {vn(y20['aov'], 0)} USD", {})], "bullet": True, "space_after": 6},
        {"runs": [(f"Lệch Finance − đơn {FOCUS_YEAR}: ", {"bold": True}),
                  (vn(REF['finance_year'][str(FOCUS_YEAR)] - y20['net_revenue'], 0, True) + " USD", {}),
                  (f" ({pct((REF['finance_year'][str(FOCUS_YEAR)] - y20['net_revenue']) / REF['finance_year'][str(FOCUS_YEAR)], 2, True)})",
                   {})], "bullet": True, "space_after": 6},
        {"runs": [("Câu hỏi: ", {"bold": True, "color": ORANGE}),
                  ("doanh thu giảm do ít đơn hơn hay đơn nhỏ hơn?", {"italic": True})], "space_before": 4}],
         size=14)

    # ---- 24. Theo tháng
    mids = [y * 100 + mm for y in (FOCUS_YEAR - 1, FOCUS_YEAR) for mm in range(1, 13)]
    nr_m = [D["month_nr"].get(i, 0) / 1e3 for i in mids]
    fin_m = [REF["finance_month"][str(i)] / 1e3 for i in mids]
    s = dk.content(sec_name, SEC, f"Xu hướng theo tháng {FOCUS_YEAR - 1}–{FOCUS_YEAR}: hai đường gần như chồng nhau",
                   ["Đỉnh cuối năm, đáy tháng 4; năm ", (str(FOCUS_YEAR), {"bold": True}),
                    " sụt mạnh từ tháng 3 – cần nêu bối cảnh khi viết báo cáo"])
    cd = CategoryChartData()
    cd.categories = [f"{i % 100:02d}/{str(i // 100)[2:]}" for i in mids]
    cd.add_series("Net Revenue (dữ liệu đơn)", nr_m)
    cd.add_series("Net Revenue (Finance)", fin_m)
    chart_frame(s, 0.4, 1.5, 9.0, 5.05, "Net Revenue theo tháng (nghìn USD)",
                caption="Dữ liệu đơn: Net Sales theo order_date − Returns theo return_date. Finance: theo posting_date.")
    gf = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, I(0.5), I(1.9), I(8.8), I(4.3), cd)
    ch = gf.chart
    style_chart(ch)
    style_axes(ch, num_fmt="#,##0", vmin=0)
    color_series(ch.series[0], GREEN, line=True, marker=True)
    color_series(ch.series[1], ORANGE, line=True, width=1.75, marker=False)
    ln = ch.series[1].format.line._get_or_add_ln()
    pdash = etree.SubElement(ln, qn("a:prstDash"))
    pdash.set("val", "dash")
    imax = int(np.argmax(nr_m))
    point_labels(ch.series[0], [vn(v, 0) for v in nr_m], pos=XL_LABEL_POSITION.ABOVE, size=11, bold=True,
                 idx={imax, len(nr_m) - 1})
    rm_i = mids.index(REPORT_MONTH)
    text(s, 9.6, 1.5, 3.35, 0.4, [("Ba điều cần thấy", {"bold": True, "color": GREEN})], size=15)
    obs = [(f"Đỉnh {cd.categories[imax].label if hasattr(cd.categories[imax], 'label') else ''}".strip(),
            f"{vn(nr_m[imax], 0)} nghìn USD"),
           (f"Tháng báo cáo {RM_LABEL}", f"{vn(nr_m[rm_i], 0)} nghìn USD, Finance {vn(fin_m[rm_i], 0)}"),
           ("Online lệch thời điểm", "Đường Finance trễ hơn ở các tháng có nhiều đơn online giao sang tháng sau")]
    for i, (a, b_) in enumerate(obs):
        y = 1.95 + i * 1.55
        box(s, 9.6, y, 3.35, 1.4, fill=[L_GREEN, L_BLUE, L_ORANGE][i], line=[GREEN, NAVY, ORANGE][i], lw=1.25)
        text(s, 9.7, y + 0.05, 3.15, 1.3, [[(a, {"bold": True, "color": [GREEN, NAVY, ORANGE][i]})], b_], size=13,
             anchor="m")

    # ---- 25. Kênh
    cy = D["chan_year"]
    yrs_full = [int(y) for y in years if int(y) <= FOCUS_YEAR]
    pos_share = [cy.loc[y, "POS"] / cy.loc[y].sum() for y in yrs_full]
    onl_share = [1 - v for v in pos_share]
    co = D["chan_orders"]
    aov_pos = cy.loc[FOCUS_YEAR, "POS"] / co.loc[FOCUS_YEAR, "POS"]
    aov_onl = cy.loc[FOCUS_YEAR, "ONLINE"] / co.loc[FOCUS_YEAR, "ONLINE"]
    s = dk.content(sec_name, SEC, "Kênh bán: cửa hàng chiếm phần lớn, online có quy tắc ghi nhận riêng",
                   ["Online chỉ chiếm khoảng ", (f"{pct(onl_share[-1], 0)}", {"bold": True, "color": ORANGE}),
                    " Net Sales nhưng là ", ("nguồn chênh lệch lớn nhất", {"bold": True, "color": RED}),
                    " khi đối soát theo tháng"])
    cd = CategoryChartData()
    cd.categories = [str(y) for y in yrs_full]
    cd.add_series("Cửa hàng (POS)", pos_share)
    cd.add_series("Website/app", onl_share)
    chart_frame(s, 0.4, 1.5, 7.2, 5.05, "Tỷ trọng Net Sales theo kênh")
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED_100, I(0.5), I(1.9), I(7.0), I(4.55), cd)
    ch = gf.chart
    style_chart(ch)
    style_axes(ch, num_fmt="0%", vmin=0, vmax=1)
    color_series(ch.series[0], GREEN)
    color_series(ch.series[1], CH_GREEN2)
    ch.plots[0].gap_width = 55
    point_labels(ch.series[0], [pct(v, 1) for v in pos_share], pos=XL_LABEL_POSITION.CENTER, color=WHITE, bold=True)
    point_labels(ch.series[1], [pct(v, 1) for v in onl_share], pos=XL_LABEL_POSITION.CENTER, color=TXT, bold=True)
    cd = CategoryChartData()
    cd.categories = ["Cửa hàng (POS)", "Website/app"]
    cd.add_series(f"AOV {FOCUS_YEAR}", [aov_pos, aov_onl])
    chart_frame(s, 7.85, 1.5, 5.1, 3.1, f"AOV theo kênh năm {FOCUS_YEAR} (USD/đơn)")
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, I(7.95), I(1.9), I(4.9), I(2.6), cd)
    ch = gf.chart
    style_chart(ch, legend=False)
    style_axes(ch, num_fmt="#,##0", vmin=0, show_val=False)
    ch.value_axis.has_major_gridlines = False
    color_series(ch.series[0], GREEN)
    ch.series[0].points[1].format.fill.solid()
    ch.series[0].points[1].format.fill.fore_color.rgb = CH_GREEN2
    ch.plots[0].gap_width = 50
    point_labels(ch.series[0], [vn(aov_pos, 0), vn(aov_onl, 0)], size=12, bold=True)
    box(s, 7.85, 4.75, 5.1, 1.8, fill=L_ORANGE, line=ORANGE, lw=1.25, align="l", anchor="m",
        paras=[[("Vì sao online quan trọng khi đối soát? ", {"bold": True, "color": ORANGE})],
               "Dữ liệu đơn tính theo ngày đặt, Finance tính theo ngày giao: đơn đặt cuối tháng, giao đầu tháng sau "
               "sẽ \"nhảy\" sang tháng khác."], size=13, margin=(0.15, 0.04, 0.15, 0.04))

    # ---- 26. Danh mục + tỷ lệ hoàn trả
    cat = D["cat"]
    s = dk.content(sec_name, SEC, f"Danh mục năm {FOCUS_YEAR}: ai tạo doanh thu, ai bị trả hàng nhiều?",
                   ["Danh mục ", ("lớn nhất", {"bold": True, "color": GREEN}), " chưa chắc có ",
                    ("tỷ lệ hoàn trả", {"bold": True, "color": RED}), " cao nhất: đọc hai biểu đồ cùng nhau"])
    cats_ = [CAT_VI.get(c_, c_) for c_ in cat.index]
    cd = CategoryChartData()
    cd.categories = cats_[::-1]
    vals = list((cat.ns / 1e3).values)[::-1]
    cd.add_series("Net Sales", vals)
    chart_frame(s, 0.4, 1.5, 6.2, 5.05, f"Net Sales theo danh mục {FOCUS_YEAR} (nghìn USD)")
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, I(0.5), I(1.9), I(6.0), I(4.55), cd)
    ch = gf.chart
    style_chart(ch, legend=False)
    style_axes(ch, show_val=False, vmin=0)
    ch.value_axis.has_major_gridlines = False
    color_series(ch.series[0], GREEN)
    ch.plots[0].gap_width = 45
    sh_ = cat.ns / cat.ns.sum()
    point_labels(ch.series[0], [f"{vn(v, 0)} ({pct(sh_.iloc[len(vals) - 1 - i], 0)})" for i, v in enumerate(vals)],
                 size=11)
    rr_ = cat.rr.sort_values()
    cd = CategoryChartData()
    cd.categories = [CAT_VI.get(c_, c_) for c_ in rr_.index]
    cd.add_series("Tỷ lệ hoàn trả", list(rr_.values))
    chart_frame(s, 6.85, 1.5, 6.1, 5.05, f"Tỷ lệ hoàn trả = Returns / Net Sales, {FOCUS_YEAR}", color=RED)
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, I(6.95), I(1.9), I(5.9), I(4.55), cd)
    ch = gf.chart
    style_chart(ch, legend=False)
    style_axes(ch, show_val=False, vmin=0)
    ch.value_axis.has_major_gridlines = False
    color_series(ch.series[0], RGBColor(0xD9, 0x8C, 0x8C))
    ch.series[0].points[len(rr_) - 1].format.fill.solid()
    ch.series[0].points[len(rr_) - 1].format.fill.fore_color.rgb = RED
    ch.plots[0].gap_width = 45
    point_labels(ch.series[0], [pct(v, 2) for v in rr_.values], size=11)

    # ---- 27. Khuyến mãi + trạng thái đơn
    pr = D["promo"].sort_values("disc", ascending=False)
    pr = pr[pr.disc > 0]
    ps = D["promo_share_year"]
    s = dk.content(sec_name, SEC, f"Khuyến mãi và trạng thái đơn: doanh thu bị \"bào mòn\" ở đâu?",
                   ["Giảm giá năm ", (str(FOCUS_YEAR), {"bold": True}), " = ",
                    (vn(y20["discounts"], 0) + " USD", {"bold": True, "color": ORANGE}), " (",
                    (pct(y20["discounts"] / y20["gross_sales"], 1), {"bold": True}), " Gross Sales); đơn huỷ và "
                                                                                    "đang giao phải loại"])
    cd = CategoryChartData()
    cd.categories = [f"{r_.promo_name} ({ix})" for ix, r_ in pr.iterrows()][::-1]
    pv = list((pr.disc / 1e3).values)[::-1]
    cd.add_series("Giảm giá", pv)
    chart_frame(s, 0.4, 1.5, 6.6, 5.05, f"Giá trị giảm giá theo chương trình {FOCUS_YEAR} (nghìn USD)", color=ORANGE)
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, I(0.5), I(1.9), I(6.4), I(4.55), cd)
    ch = gf.chart
    style_chart(ch, legend=False)
    style_axes(ch, show_val=False, vmin=0)
    ch.value_axis.has_major_gridlines = False
    color_series(ch.series[0], ORANGE)
    ch.plots[0].gap_width = 45
    point_labels(ch.series[0], [vn(v, 1) for v in pv], size=11)
    st_ = D["status_orders"]
    cd = CategoryChartData()
    order = ["COMPLETED", "CANCELLED", "PENDING"]
    cd.categories = order
    cd.add_series("Số đơn", [int(st_.get(o, 0)) for o in order])
    chart_frame(s, 7.25, 1.5, 5.7, 3.1, "Số đơn theo trạng thái, toàn bộ dữ liệu", color=NAVY)
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, I(7.35), I(1.9), I(5.5), I(2.6), cd)
    ch = gf.chart
    style_chart(ch, legend=False)
    style_axes(ch, show_val=False, vmin=0)
    ch.value_axis.has_major_gridlines = False
    color_series(ch.series[0], NAVY)
    for i, colr in enumerate([GREEN, RED, ORANGE]):
        ch.series[0].points[i].format.fill.solid()
        ch.series[0].points[i].format.fill.fore_color.rgb = colr
    ch.plots[0].gap_width = 40
    tot = sum(int(st_.get(o, 0)) for o in order)
    point_labels(ch.series[0], [f"{vn(int(st_.get(o, 0)))} ({pct(int(st_.get(o, 0)) / tot, 1)})" for o in order],
                 size=12, bold=True)
    pshare = ps.loc[FOCUS_YEAR, True] / ps.loc[FOCUS_YEAR].sum()
    box(s, 7.25, 4.75, 5.7, 1.8, fill=L_BLUE, line=NAVY, lw=1.25, align="l", anchor="m",
        paras=[{"runs": [(pct(pshare, 1), {"bold": True, "color": ORANGE}),
                         (f" Gross Sales {FOCUS_YEAR} đến từ dòng có khuyến mãi", {})], "bullet": True},
               {"runs": [("Đơn PENDING đều là đơn online đặt sát ngày trích dữ liệu: tính nhầm → doanh thu cao hơn "
                          "Finance", {})], "bullet": True}],
        size=13, margin=(0.15, 0.04, 0.15, 0.04))

    # ================================================================ PHẦN 6. DAX
    SEC = "Các mẫu công thức DAX"
    sec_name = "Phần 6 · DAX"
    # ---- 28. Bản đồ measure
    s = dk.content(sec_name, SEC, "Bản đồ measure: viết theo thứ tự, measure sau dùng lại measure trước",
                   ["Không viết lại công thức: ", ("[Net Revenue] = [Net Sales] − [Returns]", {"bold": True,
                                                                                               "color": NAVY}),
                    " – sửa một chỗ, đúng mọi nơi"])
    groups = [("1 · Doanh thu", GREEN, ["Gross Sales", "Discounts", "Net Sales", "Returns", "Net Revenue"]),
              ("2 · Đơn & giá", TEAL, ["Orders", "Units", "AOV", "ASP", "Return Rate"]),
              ("3 · Thời gian", ORANGE, ["… LM", "… SPLY", "… YTD", "% vs LM", "% vs SPLY"]),
              ("4 · Finance", NAVY, ["Finance Net Revenue", "Δ Finance", "Δ Finance %", "Cờ vượt ngưỡng",
                                     "Số ô vượt ngưỡng"]),
              ("5 · Quy đổi online", RED, ["Net Sales (ngày giao)", "Net Sales (cơ sở Finance)",
                                           "Δ sau quy đổi"])]
    for i, (t, colr, ms) in enumerate(groups):
        x = 0.4 + i * 2.53
        box(s, x, 1.5, 2.38, 5.05, fill=WHITE, line=colr, lw=1.5)
        box(s, x, 1.5, 2.38, 0.55, fill=colr, adj=0.15, paras=[(t, {"bold": True})], size=15, color=WHITE)
        for j, mm in enumerate(ms):
            box(s, x + 0.15, 2.2 + j * 0.85, 2.08, 0.65, fill=L_GRAY, line=None, adj=0.15,
                paras=[(mm, {"bold": True})], size=13, color=colr)
        if i < 4:
            box(s, x + 2.39, 3.85, 0.14, 0.35, fill=GRAY, shape="rarrow")

    # ---- 29. DAX doanh thu
    s = dk.content(sec_name, SEC, "DAX nhóm doanh thu: từ Gross Sales đến Net Revenue",
                   ["Lọc ", ("order_status = \"COMPLETED\"", {"bold": True, "color": GREEN}),
                    " ngay trong measure; Returns dùng quan hệ ngày trả hàng của fact_return"])
    code_box(s, 0.4, 1.5, 7.6, 5.05, """
Gross Sales =
CALCULATE (
    SUMX ( fact_sales_line,
        fact_sales_line[quantity] * fact_sales_line[unit_price] ),
    fact_sales_line[order_status] = "COMPLETED"
)

Discounts =
CALCULATE (
    SUM ( fact_sales_line[discount_amount] ),
    fact_sales_line[order_status] = "COMPLETED"
)

Net Sales = [Gross Sales] - [Discounts]

Returns = SUM ( fact_return[return_amount] )   // theo return_date

Net Revenue = [Net Sales] - [Returns]
""", size=13)
    box(s, 8.25, 1.5, 4.7, 5.05, fill=WHITE, line=GREEN, lw=1.5)
    box(s, 8.25, 1.5, 4.7, 0.5, fill=GREEN, adj=0.15, paras=[("Cách đọc", {"bold": True})], size=15, color=WHITE)
    text(s, 8.35, 2.1, 4.5, 4.4, [
        {"runs": [("SUMX", {"bold": True}), (" nhân quantity × unit_price từng dòng rồi mới cộng", {})],
         "bullet": True, "space_after": 6},
        {"runs": [("CALCULATE", {"bold": True}), (" thêm điều kiện COMPLETED cho mọi visual", {})],
         "bullet": True, "space_after": 6},
        {"runs": "Returns không lọc trạng thái: mọi dòng trả hàng đều thuộc đơn COMPLETED (QC đã kiểm)",
         "bullet": True, "space_after": 6},
        {"runs": [("Kiểm chứng ", {"bold": True}), (f"{RM_LABEL}: ", {}),
                  (f"Gross {vn(m['gross_sales'], 2)} − Discounts {vn(m['discounts'], 2)} = Net Sales "
                   f"{vn(m['net_sales'], 2)}; − Returns {vn(m['returns'], 2)} = Net Revenue "
                   f"{vn(m['net_revenue'], 2)} USD", {})], "bullet": True}], size=13)

    # ---- 30. DAX đơn, giá, thời gian
    s = dk.content(sec_name, SEC, "DAX nhóm đơn hàng, giá và so sánh thời gian",
                   [("DIVIDE", {"bold": True, "color": NAVY}), " thay cho dấu /; time intelligence cần ",
                    ("dim_date đã Mark as date table", {"bold": True, "color": NAVY})])
    code_box(s, 0.4, 1.5, 6.2, 5.05, """
Orders =
CALCULATE (
    DISTINCTCOUNT ( fact_sales_line[order_number] ),
    fact_sales_line[order_status] = "COMPLETED"
)

Units =
CALCULATE (
    SUM ( fact_sales_line[quantity] ),
    fact_sales_line[order_status] = "COMPLETED"
)

AOV = DIVIDE ( [Net Sales], [Orders] )
ASP = DIVIDE ( [Net Sales], [Units] )
Return Rate = DIVIDE ( [Returns], [Net Sales] )
""", size=13)
    code_box(s, 6.8, 1.5, 6.15, 3.55, """
Net Revenue LM =
CALCULATE ( [Net Revenue],
    DATEADD ( dim_date[date], -1, MONTH ) )

Net Revenue SPLY =
CALCULATE ( [Net Revenue],
    SAMEPERIODLASTYEAR ( dim_date[date] ) )

Net Revenue YTD =
CALCULATE ( [Net Revenue], DATESYTD ( dim_date[date] ) )

% vs LM = DIVIDE ( [Net Revenue] - [Net Revenue LM],
                   [Net Revenue LM] )
""", size=12)
    box(s, 6.8, 5.2, 6.15, 1.35, fill=L_BLUE, line=NAVY, lw=1.25, align="l", anchor="m",
        paras=[[("Kiểm chứng " + RM_LABEL + ": ", {"bold": True, "color": NAVY}),
                (f"Net Revenue {vn(m['net_revenue'], 0)}, LM {vn(lm['net_revenue'], 0)} → "
                 f"{pct(chg(m['net_revenue'], lm['net_revenue']), 1, True)}; SPLY {vn(sply['net_revenue'], 0)} → "
                 f"{pct(chg(m['net_revenue'], sply['net_revenue']), 1, True)}; YTD {vn(ytd['net_revenue'], 0)} USD",
                 {})]], size=13, margin=(0.15, 0.04, 0.15, 0.04))

    # ---- 31. DAX Finance
    s = dk.content(sec_name, SEC, "DAX đối soát: Finance Net Revenue, Δ, Δ% và cờ vượt ngưỡng",
                   ["Cờ xét ở ", ("tháng × cửa hàng", {"bold": True, "color": RED}),
                    ": dùng SUMX qua từng ô thay vì so tổng của cả bảng"])
    code_box(s, 0.4, 1.5, 6.2, 5.05, f"""
Finance Net Revenue =               // bỏ HQ-ADJ
CALCULATE ( SUM ( fact_finance[net_revenue] ),
    NOT ISBLANK ( fact_finance[store_key] ) )

Δ Finance =
[Finance Net Revenue] - [Net Revenue]

Δ Finance % =
DIVIDE ( [Δ Finance], ABS ( [Finance Net Revenue] ) )

Cờ vượt ngưỡng =
IF ( ABS ( [Δ Finance] ) > {int(THR_ABS)}
        && ABS ( [Δ Finance %] ) > {THR_PCT},
    "Vượt" )
""", size=13)
    code_box(s, 6.8, 1.5, 6.15, 3.3, f"""
Số ô vượt ngưỡng =
SUMX (
    CROSSJOIN ( VALUES ( dim_date[month_id] ),
                VALUES ( dim_store[store_key] ) ),
    VAR d = [Δ Finance]
    VAR p = [Δ Finance %]
    RETURN IF ( ABS ( d ) > {int(THR_ABS)} && ABS ( p ) > {THR_PCT}, 1 )
)
""", size=12)
    box(s, 6.8, 4.95, 6.15, 1.6, fill=WHITE, line=NAVY, lw=1.25, align="l", anchor="m",
        paras=[{"runs": [("Context transition: ", {"bold": True, "color": NAVY}),
                         ("measure trong SUMX được tính lại cho từng cặp tháng × cửa hàng", {})], "bullet": True},
               {"runs": [("HQ-ADJ: ", {"bold": True, "color": NAVY}),
                         ("store_key trống nên không thuộc ô nào; báo riêng ở tổng chuỗi", {})], "bullet": True}],
        size=13, margin=(0.12, 0.04, 0.12, 0.04))

    # ---- 32. DAX online theo ngày giao
    rm_b = K["bridge_month"].set_index("month_id").loc[REPORT_MONTH]
    s = dk.content(sec_name, SEC, "DAX quy đổi online về ngày giao bằng USERELATIONSHIP",
                   ["Nếu Δ online ", ("biến mất sau quy đổi", {"bold": True, "color": GREEN}),
                    " thì đó là lệch thời điểm, không phải lỗi dữ liệu"])
    code_box(s, 0.4, 1.5, 7.4, 5.05, """
Net Sales (ngày giao) =
CALCULATE ( [Net Sales],
    USERELATIONSHIP ( fact_sales_line[delivery_date],
                      dim_date[date] ) )

Net Sales (cơ sở Finance) =
VAR pos = CALCULATE ( [Net Sales],
    KEEPFILTERS ( fact_sales_line[channel_code] = "POS" ) )
VAR onl = CALCULATE ( [Net Sales (ngày giao)],
    KEEPFILTERS ( fact_sales_line[channel_code] = "ONLINE" ) )
RETURN pos + onl

Δ sau quy đổi =
[Finance Net Revenue] - ( [Net Sales (cơ sở Finance)] - [Returns] )
""", size=13)
    dlt = rm_b.diff_finance_minus_orders
    dlt_after = dlt - rm_b.adj_online_timing
    box(s, 8.05, 1.5, 4.9, 5.05, fill=WHITE, line=GREEN, lw=1.5)
    box(s, 8.05, 1.5, 4.9, 0.5, fill=GREEN, adj=0.15, paras=[("Cách đọc", {"bold": True})], size=15, color=WHITE)
    text(s, 8.15, 2.1, 4.7, 4.4, [
        {"runs": "USERELATIONSHIP bật quan hệ inactive delivery_date chỉ trong measure này", "bullet": True,
         "space_after": 6},
        {"runs": "KEEPFILTERS giữ lựa chọn của slicer Kênh, không ghi đè", "bullet": True, "space_after": 6},
        {"runs": "POS giữ ngày đặt hàng vì Finance ghi POS theo ngày bán", "bullet": True, "space_after": 6},
        {"runs": [("Ví dụ " + RM_LABEL + ": ", {"bold": True}),
                  (f"Δ toàn chuỗi {vn(dlt, 0)} USD; sau quy đổi online còn {vn(dlt_after, 0)} USD → phần còn lại "
                   "nằm ở kênh cửa hàng", {})], "bullet": True}], size=13)

    # ================================================================ PHẦN 7. Phân tích kết quả
    SEC = "Phân tích kết quả"
    sec_name = "Phần 7 · Phân tích kết quả"
    # ---- 33. Bố cục 4 trang
    s = dk.content(sec_name, SEC, "Bố cục đề xuất 4 trang dashboard (bản tham chiếu của mentor)",
                   ["Bố cục lưới: ", ("thẻ KPI ở trên", {"bold": True, "color": GREEN}), ", biểu đồ ở giữa, ",
                    ("bảng chi tiết ở dưới", {"bold": True, "color": GREEN})])
    pg = [("dash_s1.png", "Tổng quan", "Tháng này bán được bao nhiêu, khớp Finance chưa?", GREEN),
          ("dash_s2.png", "Doanh thu", "Doanh thu đến từ kênh, danh mục, quốc gia nào?", ORANGE),
          ("dash_s3.png", "Đối soát", "Lệch ở cửa hàng nào, tháng nào, kênh nào?", NAVY),
          ("dash_s4.png", "Hoàn trả & chất lượng", "Bao nhiêu doanh thu bị mất, vì sao?", RED)]
    for i, (f, t, q, colr) in enumerate(pg):
        x = 0.4 + i * 3.15
        box(s, x, 1.5, 2.95, 0.45, fill=colr, adj=0.15, paras=[(t, {"bold": True})], size=15, color=WHITE)
        box(s, x, 2.0, 2.95, 3.3, fill=WHITE, line=colr, lw=1.25, shape="rect")
        picture_fit(s, ASSETS / f, x + 0.05, 2.05, 2.85, 3.2)
        box(s, x, 5.45, 2.95, 1.1, fill=[L_GREEN, L_ORANGE, L_BLUE, L_RED][i], line=colr, lw=1.0, align="l",
            paras=[[("Câu hỏi để khai phá", {"bold": True, "color": colr})], (q, {"italic": True})], size=12,
            margin=(0.1, 0.03, 0.08, 0.03))

    # helper cho slide ảnh + câu hỏi + cách đọc
    def result_slide(sub, img, question, reading, take, colr=GREEN, img_w=7.75):
        s = dk.content(sec_name, SEC, sub, take)
        iw_, ih_ = Image.open(img).size
        img_w = min(8.4, 4.95 * iw_ / ih_ + 0.1)
        box(s, 0.4, 1.5, img_w, 5.05, fill=WHITE, line=colr, lw=1.25, shape="rect")
        picture_fit(s, img, 0.45, 1.55, img_w - 0.1, 4.95)
        x = 0.4 + img_w + 0.25
        w = 12.95 - x
        box(s, x, 1.5, w, 0.95, fill=L_BLUE, line=NAVY, lw=1.25, align="l", anchor="m",
            paras=[[("Câu hỏi của đề bài", {"bold": True, "color": NAVY, "size": 12})],
                   (question, {"italic": True, "bold": True})], size=13, margin=(0.12, 0.03, 0.1, 0.03))
        box(s, x, 2.6, w, 3.95, fill=WHITE, line=colr, lw=1.5)
        box(s, x, 2.6, w, 0.45, fill=colr, adj=0.15, paras=[("Cách đọc", {"bold": True})], size=14, color=WHITE)
        text(s, x + 0.08, 3.1, w - 0.16, 3.4, [{"runs": r_, "bullet": True, "space_after": 5} for r_ in reading],
             size=12)
        return s

    rmc = D["rm_chan"]
    lmc = D["lm_chan"]
    # ---- 34. Tổng quan
    fin_rm_store = rm_b.finance_net
    hq_rm = REF["finance_month"][str(REPORT_MONTH)] - fin_rm_store
    n_over_rm = sum(1 for o in REF["over_threshold_2020"] if o["month_id"] == REPORT_MONTH)
    result_slide(f"Trang Tổng quan, tháng {RM_LABEL}: 6 thẻ KPI và 3 biểu đồ", ASSETS / "dash_s1.png",
                 "Tháng này bán được bao nhiêu và có khớp Finance không?",
                 [[("Net Revenue ", {"bold": True}), (vn(m["net_revenue"] / 1e3, 1) + " nghìn USD", {}),
                   (f": {pct(chg(m['net_revenue'], lm['net_revenue']), 1, True)} so với {LM_LABEL}, "
                    f"{pct(chg(m['net_revenue'], sply['net_revenue']), 1, True)} so với {SPLY_LABEL}", {})],
                  [("Số đơn ", {"bold": True}), (f"{vn(m['orders'])} (cùng kỳ {vn(sply['orders'])})", {}),
                   (" nhưng ", {}), ("AOV ", {"bold": True}),
                   (f"{vn(m['aov'], 0)} USD tăng {pct(chg(m['aov'], sply['aov']), 1, True)} → giảm do ít đơn, "
                    "không do đơn nhỏ", {})],
                  [("Gross → Net: ", {"bold": True}),
                   (f"giảm giá {vn(m['discounts'] / 1e3, 1)}, hoàn trả {vn(m['returns'] / 1e3, 1)} nghìn USD", {})],
                  [("Δ Finance ", {"bold": True}), (f"{vn(dlt, 0)} USD ({pct(dlt / fin_rm_store, 2)}), ", {}),
                   (f"{n_over_rm} ô vượt ngưỡng", {"bold": True, "color": RED}),
                   (f"; chưa gồm HQ-ADJ {vn(hq_rm, 0)} USD", {})]],
                 ["Thẻ KPI = ", ("giá trị", {"bold": True, "color": GREEN}), " + ",
                  ("so với tháng trước", {"bold": True, "color": ORANGE}), " + ",
                  ("so với cùng kỳ", {"bold": True, "color": NAVY}), "; một mốc so sánh không đủ kết luận"])
    # ---- 35. Doanh thu
    top_cat = D["rm_cat"]
    result_slide(f"Trang Doanh thu, tháng {RM_LABEL}: kênh, danh mục, quốc gia, cửa hàng", ASSETS / "dash_s2.png",
                 "Doanh thu tập trung ở đâu, kênh nào tăng?",
                 [[("Cửa hàng (POS) ", {"bold": True}),
                   (f"{vn(rmc.loc['POS', 'ns'] / 1e3, 1)} nghìn USD = {pct(rmc.loc['POS', 'ns'] / rmc.ns.sum(), 1)} "
                    f"Net Sales; online {vn(rmc.loc['ONLINE', 'ns'] / 1e3, 1)}", {})],
                  [("Online ", {"bold": True}),
                   (f"{pct(chg(rmc.loc['ONLINE', 'ns'], lmc.loc['ONLINE', 'ns']), 1, True)} so với tháng trước, "
                    f"POS {pct(chg(rmc.loc['POS', 'ns'], lmc.loc['POS', 'ns']), 1, True)}", {})],
                  [(f"{CAT_VI.get(top_cat.index[0], top_cat.index[0])}", {"bold": True}),
                   (f" chiếm {pct(top_cat.iloc[0] / top_cat.sum(), 1)} Net Sales – danh mục lớn nhất", {})],
                  [("AOV theo kênh: ", {"bold": True}),
                   (f"POS {vn(rmc.loc['POS', 'ns'] / rmc.loc['POS', 'orders'], 0)}, online "
                    f"{vn(rmc.loc['ONLINE', 'ns'] / rmc.loc['ONLINE', 'orders'], 0)} USD/đơn", {})],
                  [("Top 10: ", {"bold": True}), ("đọc cột Thay đổi; ô \"–\" là cửa hàng không bán ở kỳ so sánh", {})]],
                 ["Bảng so sánh luôn kèm cột ", ("Thay đổi", {"bold": True, "color": NAVY}), ": ",
                  ("▲ xanh", {"bold": True, "color": GREEN}), " là tăng, ", ("▼ đỏ", {"bold": True, "color": RED}),
                  " là giảm"], colr=ORANGE)
    # ---- 36. Đối soát
    onl_rm = rm_rows[rm_rows.store_key == 0].iloc[0]
    n_store_rm = int((rm_rows.store_key > 0).sum())
    result_slide(f"Trang Đối soát, tháng {RM_LABEL}: Δ ở đâu và có biến mất sau quy đổi không?",
                 ASSETS / "dash_s3.png", "Chênh lệch nằm ở cửa hàng nào, tháng nào, kênh nào?",
                 [[("Δ toàn chuỗi ", {"bold": True}), (f"{vn(dlt, 0)} USD = {pct(dlt / fin_rm_store, 2)} trên Finance",
                                                        {})],
                  [("Online: ", {"bold": True}),
                   (f"Δ {vn(onl_rm.diff_finance_minus_orders, 0)} USD ({pct(onl_rm.diff_pct, 1)}), về 0 sau khi quy "
                    "về ngày giao → lệch thời điểm", {})],
                  [(f"{out.store_code}: ", {"bold": True}),
                   (f"Finance cao hơn {vn(out.diff_finance_minus_orders, 0)} USD ({pct(out.diff_pct, 1)}), không "
                    "đổi sau quy đổi → phải xem từng posting_date", {})],
                  [("Còn lại: ", {"bold": True}),
                   (f"{n_store_rm - 1} cửa hàng có số liệu khác đều trong ngưỡng", {})],
                  [("Mẹo: ", {"bold": True, "color": ORANGE}),
                   ("xếp bảng theo |Δ| giảm dần, chỉ đào sâu các dòng có cờ Vượt", {})]],
                 ["Đối soát theo thứ tự: ", ("tổng chuỗi → kênh → cửa hàng → ngày", {"bold": True, "color": NAVY})],
                 colr=NAVY)
    # ---- 37. Hoàn trả
    crop_dir = ASSETS
    s4_crop = crop_image(ASSETS / "dash_s4.png", crop_dir / "dash_s4_main.png", (0, 0.273, 1, 1))
    reason = D["r"][D["r"].mid == REPORT_MONTH].groupby("return_reason").return_amount.sum().sort_values(
        ascending=False)
    rr_cat_rm = None
    result_slide(f"Trang Hoàn trả & chất lượng, tháng {RM_LABEL}", s4_crop,
                 "Doanh thu có \"sạch\" không: trả hàng, huỷ đơn, dữ liệu lỗi?",
                 [[("Hoàn trả ", {"bold": True}), (f"{vn(m['returns'], 0)} USD, tỷ lệ {pct(m['return_rate'], 2)}", {}),
                   (f" ({vn((rrt(m) - rrt(lm)) * 100, 2, True)} điểm % so với {LM_LABEL}, "
                    f"{vn((rrt(m) - rrt(sply)) * 100, 2, True)} điểm % so với {SPLY_LABEL})", {})],
                  [("Lý do lớn nhất: ", {"bold": True}),
                   (f"{REASON_VI.get(reason.index[0], reason.index[0])} ({pct(reason.iloc[0] / reason.sum(), 1)} giá "
                    "trị trả trong tháng)", {})],
                  [("Danh mục: ", {"bold": True}),
                   ("đọc tỷ lệ cùng quy mô; danh mục nhỏ dễ có tỷ lệ cao do ít đơn", {})],
                  [("Chất lượng dữ liệu: ", {"bold": True}),
                   (f"{vn(REF['duplicates'])} dòng trùng đã loại, {vn(rows['orders_pending'])} đơn PENDING tại ngày "
                    f"trích ({AS_OF_LABEL})", {})]],
                 ["Tỷ lệ hoàn trả đo bằng ", ("điểm %", {"bold": True, "color": NAVY}), "; tăng là ",
                  ("xấu", {"bold": True, "color": RED}), " nên dùng màu đảo"], colr=RED)

    # ---- 38. Quy trình đối soát
    s = dk.content(sec_name, SEC, "Quy trình đối soát 8 bước: từ hai nguồn đến bảng giải thích chênh lệch",
                   ["Mỗi nguyên nhân phải có ", ("số tiền", {"bold": True, "color": RED}),
                    "; tổng các nguyên nhân = Δ, phần chưa giải thích = 0"])
    _, (px_, py_, pw_, ph_) = picture_fit(s, DIAG / "03_quy_trinh_doi_soat.png", 0.4, 1.42, 8.1, 5.2, align="l")
    xr = 0.4 + pw_ + 0.25
    box(s, xr, 1.5, 12.95 - xr, 5.05, fill=L_BLUE, line=NAVY, lw=1.5)
    box(s, xr, 1.5, 12.95 - xr, 0.5, fill=NAVY, adj=0.15, paras=[("Tóm tắt 8 bước", {"bold": True})], size=15,
        color=WHITE)
    steps8 = ["Bỏ dòng trùng", "Giữ đơn COMPLETED", "Gross − Discounts = Net Sales", "Net Sales − Returns = Net Revenue",
              "Map Finance → store_key theo hiệu lực, tách HQ-ADJ", "Ghép tháng × store_key, tính Δ và Δ%",
              "Ô vượt ngưỡng: xem theo ngày và theo kênh", "Bảng giải thích: tổng nguyên nhân = Δ"]
    text(s, xr + 0.1, 2.08, 12.95 - xr - 0.2, 4.4,
         [{"runs": [(f"{i + 1}. ", {"bold": True, "color": NAVY}), (t_, {})], "space_after": 5}
          for i, t_ in enumerate(steps8)], size=13)

    # ---- 39. Cầu nối mẫu (học viên)
    br = REF["bridge_2020"]
    hq_y = REF["finance_year"][str(FOCUS_YEAR)] - br["finance_net"]
    s = dk.content(sec_name, SEC, f"Bảng cầu nối năm {FOCUS_YEAR}: khung học viên cần điền",
                   ["Hai đầu cầu ", ("tính được từ dữ liệu", {"bold": True, "color": GREEN}),
                    "; phần giữa là bài làm: mỗi dòng = một nguyên nhân + số tiền + bằng chứng"])
    tpl = [["Bước", "Số tiền (USD)", "Bằng chứng cần đưa"],
           [f"Net Revenue dữ liệu đơn {FOCUS_YEAR}", vn(br["orders_net_revenue"], 2), "Measure [Net Revenue]"],
           ["± Lệch thời điểm online", "?", "Δ online trước / sau quy về delivery_date"],
           ["± Hoàn tiền hạch toán trễ", "?", "Cặp Δ trái dấu ở 2 tháng liền kề"],
           ["± Nguyên nhân phía Finance (đơn huỷ, thiếu ngày, VAT, ghi trùng…)", "?",
            "Đơn hàng, ngày, posting_date cụ thể"],
           ["= Finance Net Revenue (cửa hàng + online)", vn(br["finance_net"], 2), "Σ net_revenue đã map store_key"],
           ["± Điều chỉnh HQ-ADJ", vn(hq_y, 2), "Dòng fin_store_code = HQ-ADJ"],
           [f"= Finance Net Revenue toàn chuỗi {FOCUS_YEAR}", vn(REF["finance_year"][str(FOCUS_YEAR)], 2),
            "Σ net_revenue cả file"]]
    fills = {(1, None): L_GREEN, (5, None): L_BLUE, (7, None): L_BLUE}
    for ri in (2, 3, 4):
        fills[(ri, 1)] = L_ORANGE
    table(s, 0.4, 1.5, 8.6, 5.05, tpl, col_w=[4.0, 1.7, 3.0], size=13, fills=fills, zebra=False,
          align=["l", "r", "l"], bolds={(1, None): True, (5, None): True, (7, None): True})
    box(s, 9.25, 1.5, 3.7, 5.05, fill=L_ORANGE, line=ORANGE, lw=1.5)
    box(s, 9.25, 1.5, 3.7, 0.5, fill=ORANGE, adj=0.15, paras=[("Tự kiểm tra", {"bold": True})], size=15,
        color=WHITE)
    text(s, 9.35, 2.1, 3.5, 4.4, [
        {"runs": [("Tổng cần giải thích: ", {"bold": True}),
                  (vn(br["finance_net"] - br["orders_net_revenue"], 2, True) + " USD", {"bold": True, "color": RED})],
         "space_after": 8},
        {"runs": "Cộng các dòng ± phải ra đúng số này", "bullet": True, "space_after": 6},
        {"runs": "Lệch thời điểm cộng cả năm gần 0 nhưng từng tháng rất lớn", "bullet": True, "space_after": 6},
        {"runs": "Mỗi nguyên nhân phía Finance: ghi cửa hàng, ngày, số tiền", "bullet": True, "space_after": 6},
        {"runs": "Không \"ép\" phần dư vào một dòng \"khác\"", "bullet": True}], size=13)

    # ---- 40. F-I-R-L
    s = dk.content(sec_name, SEC, "Viết kết luận: Finding – Insight – Recommendation – Limitation",
                   ["Mỗi kết luận có một ", ("con số trên dashboard", {"bold": True, "color": RED}),
                    " làm chỗ dựa; khuyến nghị ", ("tương xứng", {"bold": True, "color": NAVY}), " bằng chứng"])
    firl = [("Finding", NAVY, "search_w", "Dữ liệu cho thấy gì, ở nhóm nào, chênh bao nhiêu, tại ngày nào",
             f"Δ online {RM_LABEL} = {vn(onl_rm.diff_finance_minus_orders, 0)} USD; về 0 khi tính theo ngày giao"),
            ("Insight", ORANGE, "bulb_w", "Ý nghĩa với kinh doanh, không nhắc lại số",
             "Chênh online là lệch thời điểm ghi nhận, không phải doanh thu bị mất"),
            ("Recommendation", GREEN, "flag_w", "Ai làm gì, theo dõi chỉ số nào",
             "Báo cáo tháng dùng Net Revenue theo cơ sở Finance; Finance kiểm tra bút toán ở ô vượt ngưỡng"),
            ("Limitation", RED, "route_w", "Giới hạn có thể ảnh hưởng kết luận và bước kiểm tra tiếp",
             f"Dữ liệu đến {AS_OF_LABEL}: đơn PENDING và hoàn tiền cuối tháng có thể làm số thay đổi")]
    for i, (t, colr, icn, d, exm) in enumerate(firl):
        x = 0.4 + i * 3.15
        box(s, x, 1.5, 2.95, 5.05, fill=WHITE, line=colr, lw=1.5)
        box(s, x, 1.5, 2.95, 0.5, fill=colr, adj=0.15, paras=[(t, {"bold": True})], size=16, color=WHITE)
        ic.circle(s, icn, x + 1.05, 2.15, 0.85, colr)
        text(s, x + 0.1, 3.1, 2.75, 1.2, [d], size=13, align="c")
        box(s, x + 0.1, 4.35, 2.75, 2.05, fill=L_GRAY, adj=0.08, align="l", anchor="t",
            paras=[[("Ví dụ", {"bold": True, "color": colr})], (exm, {"italic": True})], size=12,
            margin=(0.1, 0.06, 0.1, 0.04))

    # ================================================================ PHẦN 8. Đầu ra
    SEC = "Báo cáo đầu ra và thang chấm điểm"
    sec_name = "Phần 8 · Đầu ra & chấm điểm"
    s = dk.content(sec_name, SEC, "Hai sản phẩm cần nộp và cấu trúc báo cáo",
                   ["Báo cáo bám theo ", ("guild_report_template", {"bold": True, "color": NAVY}),
                    ": 8 phần + Phụ lục"])
    box(s, 0.4, 1.5, 4.3, 2.3, fill=L_GREEN, line=GREEN, lw=1.5)
    ic.add(s, "pie_g", 0.55, 1.65, 0.7)
    text(s, 1.35, 1.65, 3.25, 0.7, [("1  File Power BI (.pbix)", {"bold": True, "color": GREEN})], size=17,
         anchor="m")
    text(s, 0.55, 2.45, 4.05, 1.3, ["Dashboard 4 trang: Tổng quan, Doanh thu, Đối soát, Hoàn trả & chất lượng"],
         size=14)
    box(s, 0.4, 3.95, 4.3, 1.5, fill=L_ORANGE, line=ORANGE, lw=1.5)
    ic.add(s, "doc_o", 0.55, 4.1, 0.7)
    text(s, 1.35, 4.1, 3.25, 0.7, [("2  Báo cáo phân tích (PDF)", {"bold": True, "color": ORANGE})], size=17,
         anchor="m")
    text(s, 0.55, 4.85, 4.05, 0.55, ["8 phần kèm Phụ lục bảng cầu nối"], size=14)
    text(s, 0.4, 5.6, 4.3, 1.0, [{"runs": "File mở được và làm mới được", "bullet": "☐", "indent": 0.3},
                                 {"runs": "Số trong báo cáo khớp dashboard", "bullet": "☐", "indent": 0.3},
                                 {"runs": "Bảng cầu nối có phần chưa giải thích = 0", "bullet": "☐", "indent": 0.3}],
         size=12)
    box(s, 4.95, 1.5, 8.0, 0.5, fill=NAVY, adj=0.15, paras=[("Cấu trúc báo cáo", {"bold": True})], size=15,
        color=WHITE)
    rep = [("I", "Problem & Analytical Questions", "Vì sao số Sales ≠ Finance; 2–4 câu hỏi"),
           ("II", "Data & Preparation", "Nguồn, phạm vi, làm sạch, số dòng QC"),
           ("III", "Exploratory Analysis", "Xu hướng, kênh, danh mục, hoàn trả"),
           ("IV", "Main Analysis", "Đối soát tháng × cửa hàng × kênh"),
           ("V", "Key Findings & Insights", "Nguyên nhân + số tiền"),
           ("VI", "Recommendations", "Ai làm gì với Finance, cửa hàng"),
           ("VII", "Limitations & Next Steps", "As-of, PENDING, giả định"),
           ("VIII", "References", "Nguồn dữ liệu, tài liệu"),
           ("+", "Phụ lục", "Bảng cầu nối, danh sách ô vượt ngưỡng")]
    for i, (n_, t, d) in enumerate(rep):
        r_, c_ = i // 3, i % 3
        x = 4.95 + c_ * 2.7
        y = 2.15 + r_ * 1.47
        colr = GREEN if n_ == "+" else NAVY
        box(s, x, y, 2.55, 1.32, fill=WHITE, line=colr, lw=1.25, dash=True)
        box(s, x + 0.08, y + 0.1, 0.45, 0.45, fill=colr, shape="oval", paras=[(n_, {"bold": True})], size=11,
            color=WHITE, margin=(0, 0, 0, 0))
        text(s, x + 0.58, y + 0.04, 1.95, 0.6, [(t, {"bold": True, "color": colr})], size=12, anchor="m")
        text(s, x + 0.08, y + 0.65, 2.4, 0.65, [d], size=11, color=GRAY)

    # ---- rubric
    s = dk.content(sec_name, SEC, "Thang chấm điểm 100 điểm",
                   ["Rubric ", ("100 điểm", {"bold": True, "color": RED}), ": 40 kỹ thuật · 40 báo cáo · 20 sáng tạo"])
    rub = [("40 điểm", "Kỹ thuật", GREEN, L_GREEN, 3.9,
            [("Làm sạch & mô hình (10)", "Bỏ trùng, COMPLETED, map theo hiệu lực, quan hệ đúng"),
             ("DAX (12)", "KPI khớp số kiểm tra; Δ, cờ, quy đổi online"),
             ("Dashboard 4 trang (12)", "Đủ thẻ, biểu đồ, bảng; slicer đồng bộ"),
             ("Trải nghiệm (6)", "Đơn vị, caption, màu, điều hướng")]),
           ("40 điểm", "Báo cáo phân tích", NAVY, L_BLUE, 5.3,
            [("Cấu trúc 8 phần (8)", "Theo template, từ câu hỏi đến kết luận"),
             ("Đối soát (16)", "Giải thích mọi ô vượt ngưỡng; cầu nối còn lại = 0"),
             ("Lập luận (10)", "Finding → Insight → Recommendation có số làm chỗ dựa"),
             ("Giới hạn (6)", "As-of, PENDING, giả định ghi nhận")]),
           ("20 điểm", "Sáng tạo", ORANGE, L_ORANGE, 3.0,
            [("Vượt gợi ý", "Drill-through cửa hàng – ngày, tooltip"),
             ("Bảng bổ sung", "fx_rate, unit_cost, dim_customer"),
             ("Cảnh báo", "Theo dõi ô vượt ngưỡng hằng tháng")])]
    x = 0.4
    for v, t, colr, lf, w, items in rub:
        box(s, x, 1.5, w, 0.95, fill=colr, adj=0.1, paras=[[(v, {"bold": True, "size": 22})], (t, {"size": 13})],
            size=13, color=WHITE)
        box(s, x, 2.55, w, 4.0, fill=lf, line=colr, lw=1.25, adj=0.04)
        paras = []
        for a, b_ in items:
            paras.append({"runs": [(a, {"bold": True, "color": colr})], "bullet": True, "bullet_color": colr})
            paras.append({"runs": b_, "size": 13, "space_after": 8, "indent": 0.2,
                          "bullet": " "})
        text(s, x + 0.1, 2.65, w - 0.2, 3.85, paras, size=14)
        x += w + 0.175

    # ---- kết thúc
    s = dk.new("BLANK", sec_name)
    s.shapes.add_picture(io.BytesIO(dk.end_blob), 0, 0, I(SLIDE_W), I(SLIDE_H))
    text(s, 0.8, 2.6, 6.0, 1.2, [("Hỏi & đáp", {"bold": True})], size=54, color=WHITE)
    text(s, 0.8, 3.75, 6.0, 1.0, [f"{PROJECT_LABEL} · Kết quả bán hàng và đối soát doanh thu"], size=18,
         color=WHITE)
    dk.slide_number(s)

    # ================================================================ PHỤ LỤC MENTOR (ẩn)
    SEC = "Phụ lục mentor"
    sec_name = "Phụ lục mentor – ẩn khi trình chiếu cho học viên"
    # ---- A1. Cầu nối đáp án
    s = dk.content(sec_name, SEC, f"Đáp án cầu nối năm {FOCUS_YEAR} · ẩn khi trình chiếu cho học viên",
                   ["Unexplained = ", ("0 USD", {"bold": True, "color": GREEN}),
                    " ở mọi ô tháng × cửa hàng (Instructor Key, sheet bridge_store_month)"], hidden=True)
    steps_ = [("Net Revenue đơn", br["orders_net_revenue"], "total"),
              ("Online lệch thời điểm", br["adj_online_timing"], "d"),
              ("Hoàn tiền trễ", br["adj_refund_lag"], "d"),
              ("Đơn huỷ vẫn ghi nhận", br["adj_cancelled_booked"], "d"),
              ("Thiếu dữ liệu ngày", br["adj_missing_upload"], "d"),
              ("Doanh thu gồm VAT", br["adj_vat_included"], "d"),
              ("Ghi trùng bút toán", br["adj_duplicate_posting"], "d"),
              ("Finance cửa hàng + online", br["finance_net"], "total"),
              ("HQ-ADJ", hq_y, "d"),
              ("Finance toàn chuỗi", REF["finance_year"][str(FOCUS_YEAR)], "total")]
    # kiểm tra cầu nối khép kín
    run = br["orders_net_revenue"]
    for nm, v, k_ in steps_[1:7]:
        run += v
    assert abs(run - br["finance_net"]) < 0.05
    base, up, down, tot_ = [], [], [], []
    run = 0
    for nm, v, k_ in steps_:
        if k_ == "total":
            run = v
            base.append(0); up.append(0); down.append(0); tot_.append(v / 1e3)
        else:
            lo = min(run, run + v)
            base.append(lo / 1e3)
            up.append(v / 1e3 if v > 0 else 0)
            down.append(-v / 1e3 if v < 0 else 0)
            tot_.append(0)
            run += v
    vmin = (min(b for b in base if b > 0) - 60) // 20 * 20
    vmax = (max(tot_) + 40) // 20 * 20
    cd = CategoryChartData()
    cd.categories = [nm for nm, _, _ in steps_]
    cd.add_series("base", base)
    cd.add_series("Tổng", tot_)
    cd.add_series("Tăng", up)
    cd.add_series("Giảm", down)
    chart_frame(s, 0.4, 1.5, 8.2, 5.05, f"Cầu nối {FOCUS_YEAR}: dữ liệu đơn → Finance (nghìn USD, trục cắt)",
                color=RED)
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, I(0.5), I(1.9), I(8.0), I(4.6), cd)
    ch = gf.chart
    style_chart(ch, legend=False, size=10)
    style_axes(ch, num_fmt="#,##0", vmin=vmin, vmax=vmax)
    ch.category_axis.tick_labels.font.size = Pt(10)
    ch.series[0].format.fill.background()
    ch.series[0].format.line.fill.background()
    color_series(ch.series[1], NAVY)
    color_series(ch.series[2], GREEN)
    color_series(ch.series[3], RED)
    ch.plots[0].gap_width = 40
    ch.plots[0].overlap = 100
    labs_tot = [vn(v, 0) if v else None for v in tot_]
    point_labels(ch.series[1], labs_tot, pos=XL_LABEL_POSITION.INSIDE_END, color=WHITE, size=10, bold=True)
    for ser, vals_, sg in ((ch.series[2], up, 1), (ch.series[3], down, -1)):
        big = {i for i, v in enumerate(vals_) if v >= 8}
        small = {i for i, v in enumerate(vals_) if 0 < v < 8}
        labs = [vn(sg * v, 1, True) if v else None for v in vals_]
        point_labels(ser, labs, pos=XL_LABEL_POSITION.CENTER, size=10, bold=True, color=WHITE, idx=big)
        point_labels(ser, labs, pos=XL_LABEL_POSITION.INSIDE_BASE, size=10, bold=True, color=TXT, idx=small)
    rws = [["Bước", "USD"]]
    for nm, v, k_ in steps_:
        rws.append([nm, vn(v, 2, k_ != "total")])
    bolds = {(i + 1, None): True for i, (_, _, k_) in enumerate(steps_) if k_ == "total"}
    table(s, 8.85, 1.5, 4.1, 5.05, rws, col_w=[2.3, 1.8], size=12, bolds=bolds, zebra=True)

    # ---- A2. Danh sách ô vượt ngưỡng
    ot = K["over_threshold"]
    ot = ot[ot.month_id // 100 == FOCUS_YEAR].copy()
    adj_cols = {"adj_online_timing": "Online lệch thời điểm", "adj_refund_lag": "Hoàn tiền trễ",
                "adj_cancelled_booked": "Đơn huỷ vẫn ghi nhận", "adj_missing_upload": "Thiếu dữ liệu ngày",
                "adj_vat_included": "Doanh thu gồm VAT", "adj_duplicate_posting": "Ghi trùng bút toán"}
    ot["cause"] = ot[list(adj_cols)].abs().idxmax(axis=1).map(adj_cols)
    ot = ot.sort_values(["month_id", "store_code"])
    half = (len(ot) + 1) // 2
    s = dk.content(sec_name, SEC, f"{len(ot)} ô vượt ngưỡng năm {FOCUS_YEAR} và nguyên nhân chính",
                   [f"Toàn bộ dữ liệu có {REF['over_threshold_cells']} ô vượt ngưỡng; năm {FOCUS_YEAR}: ",
                    (f"{len(ot)} ô", {"bold": True, "color": RED}), " – học viên phải giải thích đủ"], hidden=True)
    for k_, part in enumerate([ot.iloc[:half], ot.iloc[half:]]):
        rws = [["Tháng", "Ô", "Δ (USD)", "Δ%", "Nguyên nhân"]]
        for r_ in part.itertuples():
            rws.append([f"{r_.month_id % 100:02d}/{r_.month_id // 100}", r_.store_code,
                        vn(r_.diff_finance_minus_orders, 2, True), pct(r_.diff_pct, 1, True), r_.cause])
        table(s, 0.4 + k_ * 6.35, 1.42, 6.2, 5.2, rws, col_w=[0.9, 0.6, 1.3, 0.9, 2.5], size=10,
              align=["l", "l", "r", "r", "l"], header_size=11)

    # ---- A3. Chênh lệch cài sẵn + nguyên nhân hệ thống
    pa = K["planted_anomalies"]
    pa = pa[pd.to_datetime(pa.date_from).dt.year == FOCUS_YEAR]
    sc = K["systemic_causes"]
    s = dk.content(sec_name, SEC, f"Chênh lệch cài sẵn năm {FOCUS_YEAR} và nguyên nhân hệ thống",
                   ["Đơn huỷ vẫn ghi nhận ", ("nhỏ từng đơn", {"bold": True, "color": ORANGE}),
                    " nhưng cộng lại đáng kể; một số đơn rơi vào ô trong ngưỡng"], hidden=True)
    grp = pa.groupby("type_vi").agg(n=("anomaly_id", "count"), usd=("effect_usd", "sum"),
                                    ids=("anomaly_id", lambda x: f"{x.iloc[0]}–{x.iloc[-1]}" if len(x) > 1 else x.iloc[0]),
                                    stores=("store_code", lambda x: ", ".join(sorted(set(str(v) for v in x if
                                                                                         str(v) != 'nan'))) or "HQ-ADJ"))
    grp = grp.sort_values("usd", key=abs, ascending=False)
    rws = [["Loại (Instructor Key)", "Mã", "Số ca", "Tác động (USD)", "Cửa hàng"]]
    for nm, r_ in grp.iterrows():
        where = r_.stores if len(r_.stores) <= 30 else (r_.stores[:28] + "…")
        rws.append([nm, r_.ids, vn(r_.n), vn(r_.usd, 2, True), where])
    table(s, 0.4, 1.5, 12.55, 0.42 * len(rws), rws, col_w=[3.6, 1.3, 0.9, 1.8, 4.9], size=12,
          align=["l", "l", "c", "r", "l"])
    y0 = 1.5 + 0.42 * len(rws) + 0.25
    rws = [["Nguyên nhân hệ thống / bẫy", "Độ lớn (USD)", "Cách phát hiện"]]
    for r_ in sc.itertuples(index=False):
        rws.append([r_[1], vn(r_[3], 2), r_[4]])
    table(s, 0.4, y0, 12.55, 6.55 - y0, rws, col_w=[4.6, 1.6, 6.3], size=11, align=["l", "r", "l"])

    # ---- A4. Số kiểm tra KPI
    s = dk.content(sec_name, SEC, f"Số kiểm tra KPI cho học viên: {RM_LABEL}, LM, SPLY, YTD",
                   ["Dùng khi chấm phần ", ("DAX (12 điểm)", {"bold": True, "color": GREEN}),
                    ": lệch quá 0,01 USD → kiểm tra bỏ trùng, trạng thái, ngày trả hàng"], hidden=True)
    keys = [("Gross Sales", "gross_sales", 2), ("Discounts", "discounts", 2), ("Net Sales", "net_sales", 2),
            ("Returns", "returns", 2), ("Net Revenue", "net_revenue", 2), ("Orders", "orders", 0),
            ("Units", "units", 0), ("AOV", "aov", 2), ("ASP", "asp", 2), ("Return Rate", "return_rate", None)]
    rws = [["KPI", f"{RM_LABEL}", f"LM {LM_LABEL}", f"SPLY {SPLY_LABEL}", f"YTD {FOCUS_YEAR}",
            f"YTD {FOCUS_YEAR - 1}"]]
    for nm, k_, d in keys:
        row = [nm]
        for src in (m, lm, sply, ytd, REF["ytd"]["sply"]):
            row.append(pct(src[k_], 2) if d is None else vn(src[k_], d))
        rws.append(row)
    rws.append(["Finance Net Revenue (toàn chuỗi)", vn(REF["finance_month"][str(REPORT_MONTH)], 2),
                vn(REF["finance_month"][str(LM_ID)], 2), vn(REF["finance_month"][str(SPLY_ID)], 2),
                vn(REF["finance_year"][str(FOCUS_YEAR)], 2), vn(REF["finance_year"][str(FOCUS_YEAR - 1)], 2)])
    table(s, 0.4, 1.5, 12.55, 5.05, rws, col_w=[2.9, 1.9, 1.9, 1.9, 2.0, 2.0], size=12, bold_first_col=True,
          align=["l", "r", "r", "r", "r", "r"])

    dk.save()
    return dk


if __name__ == "__main__":
    d = build()
    print(f"Đã lưu {OUT} · {d.n} slide · {len(d.sections)} section")
