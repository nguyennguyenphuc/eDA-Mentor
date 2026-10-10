# -*- coding: utf-8 -*-
"""Dựng ../Nexa_Sales_P1_Dashboard.xlsx — dashboard Excel một trang theo Storytelling with Data.

- Dashboard: ô vàng chọn Tháng + Kênh; KPI bằng SUMIFS (so với tháng trước, cùng kỳ); câu "Điểm chính" ghép bằng
  công thức; tiêu đề hành động trên mỗi biểu đồ; 3 biểu đồ gốc (xám + 1 màu nhấn); bảng cửa hàng vượt ngưỡng.
- Doi_soat: tháng × cửa hàng, Δ / Δ% / cờ vượt ngưỡng bằng công thức (lọc được).
- Data_thang, Data_danh_muc: số tổng hợp từ aggregate.load() (đọc lại file giao học viên).
- Calc: cột phụ cho biểu đồ và danh sách chọn (không sửa).
Chỉ dùng hàm Excel 2007 (SUMIFS, COUNTIFS, INDEX, MATCH, LARGE, IFERROR) để Excel và LibreOffice đều tính được.
Chạy:  python build_dashboard_xlsx.py   (rồi tính lại công thức bằng LibreOffice — xem README_mentor.md)
"""
import aggregate as A
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

OUT = A.P1 / "Nexa_Sales_P1_Dashboard.xlsx"
REF = A.REF
D = A.load()
ms, cat, st = D["ms"].copy(), D["cat"].copy(), D["st"]
CAT_VI, COUNTRY_VI = D["CAT_VI"], D["COUNTRY_VI"]

FONT = "Arial"
INK, INK2, MUTED, ACCENT, CTX, CRIT, RULE = "1F1F1E", "52514E", "8A8984", "2A78D6", "B9B7B0", "D03B3B", "E4E3DF"
f = lambda **k: Font(name=FONT, **{"size": 10, "color": INK, **k})
YELLOW = PatternFill("solid", fgColor="FFF2B3")
thin = Side(style="thin", color=RULE)
NUM0, NUM1, NUM2 = '#,##0;−#,##0;0', '#,##0.0;−#,##0.0;0.0', '#,##0.00;−#,##0.00;0.00'
PCT1, PCTS = '0.0%', '+0.0%;−0.0%;0.0%'
PPS = '+0.00" điểm %";−0.00" điểm %";0.00" điểm %"'

wb = Workbook()
wsD = wb.active
wsD.title = "Dashboard"
wsS = wb.create_sheet("Doi_soat")
wsT = wb.create_sheet("Data_thang")
wsC = wb.create_sheet("Data_danh_muc")
wsK = wb.create_sheet("Calc")
wsH = wb.create_sheet("Huong_dan")

# ------------------------------------------------------------------ Data_thang (giá trị, Excel Table)
sti = st.set_index("store_key")
ms = ms.merge(st[["store_key", "store_code", "store_name", "country", "channel_code"]], on="store_key")
ms = ms.sort_values(["month_id", "store_key"]).reset_index(drop=True)
TH = ["month_id", "month_label", "store_key", "store_code", "store_name", "country", "channel", "gross_sales",
      "discounts", "net_sales", "returns", "orders", "units", "finance_net_revenue", "net_sales_delivery_basis",
      "cancelled_orders", "pending_orders", "duplicate_lines"]
wsT.append(TH)
for r in ms.itertuples():
    wsT.append([int(r.month_id), f"{r.month_id % 100:02d}/{r.month_id // 100}", int(r.store_key), r.store_code,
                r.store_name, COUNTRY_VI[r.country], r.channel_code, round(r.gross, 2), round(r.disc, 2),
                round(r.ns, 2), round(r.ret, 2), int(r.orders), int(r.units), round(r.fin, 2), round(r.ns_post, 2),
                int(r.canc), int(r.pend), int(r.dup)])
NT = len(ms) + 1
wsT.add_table(Table(displayName="tbl_Data_thang", ref=f"A1:R{NT}",
                    tableStyleInfo=TableStyleInfo(name="TableStyleLight15", showRowStripes=True)))
for j in range(8, 16):
    for row in wsT.iter_rows(min_row=2, min_col=j, max_col=j):
        row[0].number_format = NUM2 if j not in (12, 13) else NUM0
COL = {h: get_column_letter(i + 1) for i, h in enumerate(TH)}
rng = lambda h: f"Data_thang!${COL[h]}$2:${COL[h]}${NT}"

# ------------------------------------------------------------------ Data_danh_muc (tháng × kênh × danh mục)
cat = cat.merge(st[["store_key", "channel_code"]], on="store_key")
cg = cat.groupby(["month_id", "channel_code", "category"])[["ns", "ret", "units"]].sum().reset_index()
wsC.append(["month_id", "channel", "category", "net_sales", "returns", "units"])
for r in cg.itertuples():
    wsC.append([int(r.month_id), r.channel_code, CAT_VI[r.category], round(r.ns, 2), round(r.ret, 2), int(r.units)])
NC = len(cg) + 1
wsC.add_table(Table(displayName="tbl_Data_danh_muc", ref=f"A1:F{NC}",
                    tableStyleInfo=TableStyleInfo(name="TableStyleLight15", showRowStripes=True)))
crng = lambda c: f"Data_danh_muc!${c}$2:${c}${NC}"

# ------------------------------------------------------------------ Dashboard: tham số
wsD.sheet_view.showGridLines = False
widths = {"A": 2, "B": 38, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15, "H": 11, "I": 12, "J": 12, "K": 12,
          "L": 12, "M": 12, "N": 12, "O": 12, "P": 12}
for k, v in widths.items():
    wsD.column_dimensions[k].width = v
months = sorted(ms.month_id.unique().tolist(), reverse=True)
months = [m for m in months if ms[ms.month_id == m].orders.sum() > 0]
DEF = REF["report_month"]
wsD["B2"] = "Kết quả bán hàng và đối soát doanh thu · Nexa Home"
wsD["B2"].font = f(size=18, bold=True)
wsD["B3"] = (f"Nguồn: Nexa_Sales_P1_Dataset.xlsx, finance_revenue.xlsx · dữ liệu đến "
             f"{REF['as_of'][8:10]}/{REF['as_of'][5:7]}/{REF['as_of'][:4]} · USD, chưa VAT · chỉ đơn COMPLETED, đã loại "
             "dòng trùng · Δ = Finance − dữ liệu đơn · không gồm bút toán HQ-ADJ")
wsD["B3"].font = f(size=9, color=MUTED)
wsD["B5"], wsD["C5"] = "Tháng báo cáo", f"{DEF % 100:02d}/{DEF // 100}"
wsD["E5"], wsD["F5"] = "Kênh bán", "Tất cả"
for c in ("B5", "E5"):
    wsD[c].font = f(color=INK2)
    wsD[c].alignment = Alignment(horizontal="right")
for c in ("C5", "F5"):
    wsD[c].fill, wsD[c].font = YELLOW, f(bold=True, size=11)
    wsD[c].alignment = Alignment(horizontal="center")
wsD["H5"] = "← ô màu vàng: chọn tháng và kênh từ danh sách; mọi số, câu và biểu đồ tự cập nhật"
wsD["H5"].font = f(size=9, color=MUTED, italic=True)
dv1 = DataValidation(type="list", formula1=f"=Calc!$A$2:$A${len(months) + 1}", allow_blank=False)
dv2 = DataValidation(type="list", formula1='"Tất cả,Cửa hàng,Online"', allow_blank=False)
wsD.add_data_validation(dv1)
wsD.add_data_validation(dv2)
dv1.add("C5")
dv2.add("F5")
# tham số tính (chữ xám nhỏ): mã kênh, month_id, tháng trước, cùng kỳ, ngưỡng
wsD["B6"] = "tham số tính (không sửa): kênh · tháng · tháng trước · cùng kỳ · ngưỡng USD · ngưỡng %"
wsD["B6"].font = f(size=8, color=MUTED)
wsD["C6"] = '=IF(F5="Cửa hàng","POS",IF(F5="Online","ONLINE","*"))'
wsD["D6"] = "=INDEX(Calc!$B$2:$B$100,MATCH(C5,Calc!$A$2:$A$100,0))"
wsD["E6"] = "=IF(MOD(D6,100)=1,D6-89,D6-1)"
wsD["F6"] = "=D6-100"
wsD["G6"] = REF["threshold"]["abs_usd"]
wsD["H6"] = REF["threshold"]["pct"]
for c in ("C6", "D6", "E6", "F6", "G6", "H6"):
    wsD[c].font = f(size=8, color=MUTED)
M_, LM_, SP_, CH = "$D$6", "$E$6", "$F$6", "$C$6"


def S(field, mref, ch=CH):
    return f"SUMIFS({rng(field)},{rng('month_id')},{mref},{rng('channel')},{ch})"


def SC(field, mref):          # dùng trong sheet Calc: tham số kênh phải ghi rõ sheet Dashboard
    return S(field, mref, "Dashboard!$C$6")


# ------------------------------------------------------------------ KPI bảng (dòng 13–24)
wsD["B12"], wsD["C12"], wsD["D12"], wsD["E12"], wsD["F12"], wsD["G12"] = (
    "Chỉ tiêu", "Tháng chọn", "Tháng trước", "so với tháng trước", "Cùng kỳ năm trước", "so với cùng kỳ")
for c in "BCDEFG":
    wsD[f"{c}12"].font = f(bold=True, color=INK2)
    wsD[f"{c}12"].border = Border(bottom=Side(style="thin", color=INK2))
    wsD[f"{c}12"].alignment = Alignment(horizontal="left" if c == "B" else "right", wrap_text=True)
KP = [  # (nhãn, công thức theo cột với {m}, định dạng, kiểu so sánh)
    ("Gross Sales (USD)", lambda m, c: "=" + S("gross_sales", m), NUM0, "pct"),
    ("Giảm giá (USD)", lambda m, c: "=" + S("discounts", m), NUM0, "pct"),
    ("Net Sales (USD)", lambda m, c: "=" + S("net_sales", m), NUM0, "pct"),
    ("Hoàn trả (USD)", lambda m, c: "=" + S("returns", m), NUM0, "pct"),
    ("Net Revenue (USD)", lambda m, c: f"={c}15-{c}16", NUM0, "pct"),
    ("Số đơn hoàn tất", lambda m, c: "=" + S("orders", m), NUM0, "pct"),
    ("AOV (USD/đơn)", lambda m, c: f'=IFERROR({c}15/{c}18,"")', NUM0, "pct"),
    ("Tỷ lệ hoàn trả", lambda m, c: f'=IFERROR({c}16/{c}15,"")', '0.00%', "pp"),
    ("Finance Net Revenue (USD)", lambda m, c: "=" + S("finance_net_revenue", m), NUM0, "pct"),
    ("Δ Finance − dữ liệu đơn (USD)", lambda m, c: f"={c}21-{c}17", NUM0, "none"),
    ("Δ sau khi quy online về ngày giao (USD)", lambda m, c: f"={c}21-(" + S("net_sales_delivery_basis", m) + f"-{c}16)", NUM0, "none"),
    ("Số cửa hàng vượt ngưỡng", lambda m, c: f'=COUNTIFS(Doi_soat!$A$2:$A${NT},{m},Doi_soat!$F$2:$F${NT},{CH},Doi_soat!$L$2:$L${NT},"⚠ Vượt")', NUM0, "none"),
]
for i, (lab, fn, fmt, kind) in enumerate(KP):
    r = 13 + i
    wsD[f"B{r}"] = lab
    wsD[f"B{r}"].font = f(bold=lab.startswith(("Net Revenue", "Δ Finance")))
    for col, mref in (("C", M_), ("D", LM_), ("F", SP_)):
        wsD[f"{col}{r}"] = fn(mref, col)
        wsD[f"{col}{r}"].number_format = fmt
        wsD[f"{col}{r}"].font = f(bold=(col == "C"))
    if kind == "pct":
        wsD[f"E{r}"] = f'=IFERROR(C{r}/D{r}-1,"")'
        wsD[f"G{r}"] = f'=IFERROR(C{r}/F{r}-1,"")'
        wsD[f"E{r}"].number_format = wsD[f"G{r}"].number_format = PCTS
    elif kind == "pp":
        wsD[f"E{r}"] = f'=IFERROR(C{r}-D{r},"")'
        wsD[f"G{r}"] = f'=IFERROR(C{r}-F{r},"")'
        wsD[f"E{r}"].number_format = wsD[f"G{r}"].number_format = '+0.00%;−0.00%;0.00%'
    for c in "BCDEFG":
        wsD[f"{c}{r}"].border = Border(bottom=thin)
# màu chữ ▲/▼ cho cột so sánh: tốt = xanh, xấu = cam (Hoàn trả, Tỷ lệ hoàn trả: tăng là xấu)
good_up = "AND(ISNUMBER(E13),E13>0,ROW()<>16,ROW()<>20)"
for colref in ("E13:E20", "G13:G20"):
    c0 = colref[0]
    wsD.conditional_formatting.add(colref, FormulaRule(
        formula=[f'AND(ISNUMBER({c0}13),OR(AND({c0}13>0,ROW()<>16,ROW()<>20),AND({c0}13<0,OR(ROW()=16,ROW()=20))))'],
        font=Font(name=FONT, color=ACCENT, bold=True)))
    wsD.conditional_formatting.add(colref, FormulaRule(
        formula=[f'AND(ISNUMBER({c0}13),OR(AND({c0}13<0,ROW()<>16,ROW()<>20),AND({c0}13>0,OR(ROW()=16,ROW()=20))))'],
        font=Font(name=FONT, color="EB6834", bold=True)))
wsD.conditional_formatting.add("C22:C24", FormulaRule(formula=[f"ABS(C22)>$G$6"], font=Font(name=FONT, color=CRIT, bold=True)))

# ------------------------------------------------------------------ Điểm chính (dòng 8–10)
wsD["B8"] = "ĐIỂM CHÍNH"
wsD["B8"].font = f(bold=True, color=ACCENT, size=9)
sgn = lambda x: f'IF({x}>=0,"+","−")&FIXED(ABS({x})*100,1)&"%"'
wsD["B9"] = ('=IFERROR("Net Revenue "&C5&" đạt "&FIXED(C17/1000,1)&" nghìn USD — "&' + sgn("E17") +
             '&" so với tháng trước, "&' + sgn("G17") + '&" so với cùng kỳ năm trước.","Chọn tháng có dữ liệu.")')
wsD["B10"] = ('=IFERROR("Biến động so với cùng kỳ đến chủ yếu từ "&IF(ABS(LN(1+G18))>=ABS(LN(1+G19)),"số đơn (","giá trị mỗi đơn (AOV ")&'
              'IF(ABS(LN(1+G18))>=ABS(LN(1+G19)),' + sgn("G18") + ',' + sgn("G19") + ')&"); "&'
              'IF(ABS(LN(1+G18))>=ABS(LN(1+G19)),"AOV "&' + sgn("G19") + ',"số đơn "&' + sgn("G18") + ')&".","")')
wsD["B11"] = ('=IFERROR("Finance "&IF(C22>=0,"cao","thấp")&" hơn dữ liệu đơn "&FIXED(ABS(C22),0)&" USD; sau khi quy online về '
              'ngày giao còn "&FIXED(C23,0)&" USD; "&C24&" cửa hàng vượt ngưỡng (bảng bên dưới, sheet Doi_soat).","")')
wsD["B9"].font = f(size=13, bold=True)
wsD["B10"].font = wsD["B11"].font = f(size=10.5, color=INK2)

# ------------------------------------------------------------------ cầu doanh thu (dòng 27–31) + data bar
wsD["B27"] = '="Giảm giá và hoàn trả lấy đi "&IFERROR(FIXED((C14+C16)/C13*100,1),"–")&"% Gross Sales"'
wsD["B27"].font = f(bold=True, size=11)
for i, (lab, frm) in enumerate([("Gross Sales", "=C13"), ("− Giảm giá", "=-C14"), ("− Hoàn trả", "=-C16"),
                                ("Net Revenue", "=C17")]):
    r = 28 + i
    wsD[f"B{r}"], wsD[f"C{r}"] = lab, frm
    wsD[f"C{r}"].number_format = NUM0
    wsD[f"B{r}"].font = f(bold=(lab == "Net Revenue"))
wsD.conditional_formatting.add("C28:C31", DataBarRule(start_type="num", start_value=0, end_type="max", color=CTX,
                                                      showValue=True))

# ------------------------------------------------------------------ Calc: danh sách tháng, cửa sổ 24 tháng, danh mục
wsK["A1"], wsK["B1"] = "Tháng (chọn)", "month_id"
for i, m in enumerate(months):
    wsK.cell(2 + i, 1, f"{m % 100:02d}/{m // 100}")
    wsK.cell(2 + i, 2, int(m))
hdr = ["k", "month_id", "Tháng", "Net Revenue (nghìn USD)", "Số đơn", "Δ sau quy đổi (nghìn USD)"]
for j, h in enumerate(hdr):
    wsK.cell(1, 4 + j, h)
for k in range(24):
    r = 2 + k
    wsK.cell(r, 4, k)
    wsK.cell(r, 5, f"=YEAR(DATE(INT(Dashboard!$D$6/100),MOD(Dashboard!$D$6,100)-23+D{r},1))*100"
                   f"+MONTH(DATE(INT(Dashboard!$D$6/100),MOD(Dashboard!$D$6,100)-23+D{r},1))")
    wsK.cell(r, 6, f'=RIGHT("0"&MOD(E{r},100),2)&"/"&RIGHT(INT(E{r}/100),2)')
    wsK.cell(r, 7, f"=({SC('net_sales', f'E{r}')}-{SC('returns', f'E{r}')})/1000").number_format = NUM1
    wsK.cell(r, 8, f"={SC('orders', f'E{r}')}").number_format = NUM0
    wsK.cell(r, 9, f"=({SC('finance_net_revenue', f'E{r}')}-{SC('net_sales_delivery_basis', f'E{r}')}"
                   f"+{SC('returns', f'E{r}')})/1000").number_format = NUM1
cats = list(CAT_VI.values())
for j, h in enumerate(["Danh mục", "Net Sales (nghìn USD)", "khoá xếp", "Danh mục (xếp)", "Net Sales (xếp)"]):
    wsK.cell(1, 11 + j, h)
for i, cname in enumerate(cats):
    r = 2 + i
    wsK.cell(r, 11, cname)
    wsK.cell(r, 12, f"=SUMIFS({crng('D')},{crng('A')},Dashboard!$D$6,{crng('B')},Dashboard!$C$6,{crng('C')},K{r})/1000").number_format = NUM1
    wsK.cell(r, 13, f"=L{r}+ROW()/1000000000")
    wsK.cell(r, 14, f"=INDEX($K$2:$K$9,MATCH(LARGE($M$2:$M$9,ROW()-1),$M$2:$M$9,0))")
    wsK.cell(r, 15, f"=INDEX($L$2:$L$9,MATCH(LARGE($M$2:$M$9,ROW()-1),$M$2:$M$9,0))").number_format = NUM1
wsK["A30"] = "Sheet tính toán phụ cho Dashboard — không sửa."
wsK["A30"].font = f(italic=True, color=MUTED)

# ------------------------------------------------------------------ Doi_soat (tháng × cửa hàng, công thức)
SH = ["month_id", "Tháng", "Mã CH", "Cửa hàng", "Quốc gia", "Kênh", "Net Revenue đơn (USD)", "Finance (USD)",
      "Δ (USD)", "Δ %", "Δ sau quy đổi online (USD)", "Cờ", "khoá xếp (Dashboard)"]
wsS.append(SH)
for i in range(2, NT + 1):
    wsS.append([f"=Data_thang!A{i}", f"=Data_thang!B{i}", f"=Data_thang!D{i}", f"=Data_thang!E{i}",
                f"=Data_thang!F{i}", f"=Data_thang!G{i}", f"=Data_thang!J{i}-Data_thang!K{i}", f"=Data_thang!N{i}",
                f"=H{i}-G{i}", f'=IF(H{i}=0,IF(I{i}=0,0,1),I{i}/ABS(H{i}))',
                f"=H{i}-(Data_thang!O{i}-Data_thang!K{i})",
                f'=IF(AND(ABS(I{i})>Dashboard!$G$6,ABS(J{i})>Dashboard!$H$6),"⚠ Vượt","")',
                f'=IF(AND(A{i}=Dashboard!$D$6,OR(Dashboard!$C$6="*",F{i}=Dashboard!$C$6),L{i}<>""),ABS(I{i})+ROW()/1000000000,0)'])
    for c in ("G", "H", "I", "K"):
        wsS[f"{c}{i}"].number_format = NUM2
    wsS[f"J{i}"].number_format = PCTS
wsS.add_table(Table(displayName="tbl_Doi_soat", ref=f"A1:M{NT}",
                    tableStyleInfo=TableStyleInfo(name="TableStyleLight15", showRowStripes=True)))
wsS.conditional_formatting.add(f"L2:L{NT}", FormulaRule(formula=['L2<>""'], font=Font(name=FONT, color=CRIT, bold=True)))
wsS.freeze_panes = "A2"
for c, w in zip("ABCDEFGHIJKLM", [10, 9, 8, 34, 10, 9, 20, 16, 14, 9, 22, 10, 12]):
    wsS.column_dimensions[c].width = w
wsS.column_dimensions["M"].hidden = True

# ------------------------------------------------------------------ Dashboard: bảng cửa hàng vượt ngưỡng (dòng 50+)
wsD["B49"] = '=IF(C24=0,"Không cửa hàng nào vượt ngưỡng tháng "&C5,C24&" cửa hàng vượt ngưỡng tháng "&C5&" — cần giải thích nguyên nhân và số tiền")'
wsD["B49"].font = f(bold=True, size=11)
for j, h in enumerate(["Cửa hàng", "Quốc gia", "Mã CH", "Net Revenue đơn", "Finance", "Δ (USD)", "Δ %", "Δ sau quy đổi"]):
    c = wsD.cell(50, 2 + j, h)
    c.font = f(bold=True, color=INK2)
    c.border = Border(bottom=Side(style="thin", color=INK2))
    c.alignment = Alignment(horizontal="left" if j < 3 else "right")
for k in range(1, 11):
    r = 50 + k
    key = f"LARGE(Doi_soat!$M$2:$M${NT},{k})"
    row = f"MATCH({key},Doi_soat!$M$2:$M${NT},0)"
    for j, col in enumerate("DECGHIJK"):
        c = wsD.cell(r, 2 + j, f'=IF({key}=0,"",INDEX(Doi_soat!${col}$2:${col}${NT},{row}))')
        c.number_format = PCTS if col == "J" else NUM0
        c.border = Border(bottom=thin)
    wsD.cell(r, 7).font = f(color=CRIT, bold=True)

# ------------------------------------------------------------------ biểu đồ (xám + 1 màu nhấn; tiêu đề hành động ở ô phía trên)
def style_axes(ch):
    ch.legend = None
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="ECEBE7"))
    for ax in (ch.x_axis, ch.y_axis):
        ax.delete = False
        ax.txPr = None
    ch.x_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill="CFCDC7"))
    ch.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))


wsD["I12"] = '="Net Revenue "&C5&" bằng "&IFERROR(FIXED(C17/F17*100,0),"–")&"% cùng kỳ năm trước"'
wsD["I12"].font = f(bold=True, size=11)
wsD["I13"] = "Net Revenue theo tháng, nghìn USD · 24 tháng đến tháng chọn (điểm xanh)"
wsD["I13"].font = f(size=9, color=MUTED)
lc = LineChart()
lc.add_data(Reference(wsK, min_col=7, min_row=1, max_row=25), titles_from_data=True)
lc.set_categories(Reference(wsK, min_col=6, min_row=2, max_row=25))
s0 = lc.series[0]
s0.graphicalProperties.line.solidFill = CTX
s0.graphicalProperties.line.width = 28000
s0.smooth = False
s0.marker = Marker(symbol="none")
dp = DataPoint(idx=23)
dp.marker = Marker(symbol="circle", size=8)
dp.marker.graphicalProperties = GraphicalProperties(solidFill=ACCENT, ln=LineProperties(solidFill=ACCENT))
s0.dPt.append(dp)
style_axes(lc)
lc.y_axis.number_format = '#,##0'
lc.x_axis.tickLblSkip = 3
lc.height, lc.width = 7.2, 17.5
wsD.add_chart(lc, "I14")

wsD["I27"] = '="Sau khi quy online về ngày giao, Δ "&C5&" còn "&FIXED(C23,0)&" USD"'
wsD["I27"].font = f(bold=True, size=11)
wsD["I28"] = "Δ sau quy đổi online theo tháng, nghìn USD · tháng chọn tô xanh · chi tiết ở sheet Doi_soat"
wsD["I28"].font = f(size=9, color=MUTED)
bc = BarChart()
bc.type = "col"
bc.add_data(Reference(wsK, min_col=9, min_row=1, max_row=25), titles_from_data=True)
bc.set_categories(Reference(wsK, min_col=6, min_row=2, max_row=25))
bc.series[0].graphicalProperties = GraphicalProperties(solidFill=CTX, ln=LineProperties(noFill=True))
dp = DataPoint(idx=23)
dp.graphicalProperties = GraphicalProperties(solidFill=ACCENT, ln=LineProperties(noFill=True))
bc.series[0].dPt.append(dp)
bc.gapWidth = 60
style_axes(bc)
bc.y_axis.number_format = '#,##0.0'
bc.x_axis.tickLblSkip = 3
bc.x_axis.tickLblPos = "low"            # nhãn tháng ở đáy, không đè cột âm
bc.height, bc.width = 7.2, 17.5
wsD.add_chart(bc, "I29")

wsD["B34"] = '=INDEX(Calc!$N$2:$N$9,1)&" chiếm "&IFERROR(FIXED(INDEX(Calc!$O$2:$O$9,1)*1000/C15*100,0),"–")&"% Net Sales"'
wsD["B34"].font = f(bold=True, size=11)
wsD["B35"] = "Net Sales tháng chọn theo danh mục, nghìn USD"
wsD["B35"].font = f(size=9, color=MUTED)
hb = BarChart()
hb.type = "bar"
hb.add_data(Reference(wsK, min_col=15, min_row=1, max_row=9), titles_from_data=True)
hb.set_categories(Reference(wsK, min_col=14, min_row=2, max_row=9))
hb.series[0].graphicalProperties = GraphicalProperties(solidFill=CTX, ln=LineProperties(noFill=True))
dp = DataPoint(idx=0)
dp.graphicalProperties = GraphicalProperties(solidFill=ACCENT, ln=LineProperties(noFill=True))
hb.series[0].dPt.append(dp)
hb.gapWidth = 50
style_axes(hb)
hb.x_axis.scaling.orientation = "maxMin"       # mục lớn nhất ở trên cùng
hb.y_axis.number_format = '#,##0'
hb.height, hb.width = 6.6, 15.5
wsD.add_chart(hb, "B36")

# ------------------------------------------------------------------ Huong_dan
G = [
    ("Mục đích", "Dashboard Excel một trang cho Sales Manager và Finance: tháng này bán được bao nhiêu, vì sao tăng/giảm, "
                 "và số của dữ liệu đơn lệch Finance ở đâu."),
    ("Cách dùng", "Ở sheet Dashboard, chọn Tháng báo cáo và Kênh bán trong 2 ô màu vàng. KPI, câu 'Điểm chính', tiêu đề "
                  "biểu đồ và bảng cửa hàng vượt ngưỡng tự cập nhật. Sheet Doi_soat lọc được theo tháng, cửa hàng, cờ."),
    ("Định nghĩa", "Net Sales = Gross Sales − Giảm giá; Net Revenue = Net Sales − Hoàn trả (theo ngày trả hàng); "
                   "AOV = Net Sales / số đơn; Δ = Finance − dữ liệu đơn; Δ sau quy đổi = online tính theo ngày giao."),
    ("Ngưỡng", f"Ô tháng × cửa hàng vượt ngưỡng khi |Δ| > {REF['threshold']['abs_usd']:.0f} USD và |Δ %| > "
               f"{REF['threshold']['pct']:.0%} (sửa ở Dashboard!G6, H6 nếu Finance đổi quy định)."),
    ("Nguồn", "Data_thang và Data_danh_muc tổng hợp từ Nexa_Sales_P1_Dataset.xlsx và finance_revenue.xlsx bằng "
              "_mentor/build_dashboard_xlsx.py (chỉ đơn COMPLETED, đã loại dòng trùng). Không gồm bút toán HQ-ADJ."),
    ("Trình bày", "Theo Storytelling with Data: mỗi biểu đồ có tiêu đề nêu kết luận; màu xám cho bối cảnh, xanh dương cho "
                  "điều cần nhìn, đỏ chỉ cho chênh lệch vượt ngưỡng; bỏ lưới và khung thừa; nguồn và ngày dữ liệu ở đầu trang."),
    ("Giới hạn", "Dashboard Excel lọc theo kênh, chưa lọc theo quốc gia (xem bản HTML hoặc tự thêm cột vào SUMIFS). "
                 "Excel cần bật tính toán tự động (Formulas → Calculation Options → Automatic)."),
]
wsH.column_dimensions["A"].width, wsH.column_dimensions["B"].width = 16, 110
wsH.append(["Mục", "Nội dung"])
for row in G:
    wsH.append(list(row))
for row in wsH.iter_rows(min_row=1, max_row=len(G) + 1):
    for c in row:
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = f(bold=(c.row == 1 or c.column == 1))

# font Arial cho mọi ô chưa đặt
for ws in (wsT, wsC, wsK):
    for row in ws.iter_rows():
        for c in row:
            if c.font.name != FONT:
                c.font = f(bold=(c.row == 1))
for c in range(1, 19):
    wsT.column_dimensions[get_column_letter(c)].width = 14
wsT.column_dimensions["E"].width = 34
wsD.page_setup.orientation = "landscape"
wsD.page_setup.fitToWidth = 1
wsD.sheet_properties.pageSetUpPr.fitToPage = True
from openpyxl.workbook.properties import CalcProperties
wb.calculation = CalcProperties(fullCalcOnLoad=True)   # Excel tính lại mọi công thức khi mở
wb.properties.title = "Nexa Sales P1 – Dashboard Excel"
wb.properties.creator = "EDA Mentor"
wb.save(OUT)
print("saved", OUT.name, "rows", NT - 1)
