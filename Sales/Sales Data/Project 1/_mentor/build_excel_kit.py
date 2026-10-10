# -*- coding: utf-8 -*-
"""Sinh bộ dựng dashboard bằng Excel: ../excel_build/m/*.pq (Power Query M), ../excel_build/measures.dax (Power Pivot)
và _mentor/pq_steps.json (số dòng sau từng bước M, dùng cho sơ đồ 05, guideline và slide).

Máy dựng tài liệu không có Excel, nên mỗi bước M được "chạy song song" bằng pandas (cùng logic) để lấy số dòng và
assert khớp QC/ref_numbers.json. Mentor nên chạy thử một lần trên Excel Windows (Microsoft 365 / 2016+ có Power Pivot).
Chạy sau generate_dataset.py:  python build_excel_kit.py
"""
import json
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
P1 = HERE.parent
KIT = P1 / "excel_build"
REF = json.loads((HERE / "ref_numbers.json").read_text(encoding="utf-8"))
THR_ABS, THR_PCT = REF["threshold"]["abs_usd"], REF["threshold"]["pct"]

DS, FIN = "Nexa_Sales_P1_Dataset.xlsx", "finance_revenue.xlsx"
T_INT, T_DATE, T_TEXT, T_CUR, T_PCT, T_BOOL = "Int64.Type", "type date", "type text", "Currency.Type", "Percentage.Type", "type logical"
T_NUM = "type number"

TYPES = {
    "fact_sales_line": [("order_number", T_INT), ("line_item", T_INT), ("order_date", T_DATE), ("delivery_date", T_DATE),
                        ("customer_key", T_INT), ("store_key", T_INT), ("channel_code", T_TEXT), ("product_key", T_INT),
                        ("quantity", T_INT), ("unit_price", T_CUR), ("promo_code", T_TEXT), ("discount_amount", T_CUR),
                        ("order_status", T_TEXT), ("currency_code", T_TEXT), ("unit_cost", T_CUR)],
    "fact_return": [("return_id", T_TEXT), ("order_number", T_INT), ("line_item", T_INT), ("return_date", T_DATE),
                    ("return_qty", T_INT), ("return_amount", T_CUR), ("return_reason", T_TEXT)],
    "finance_revenue": [("posting_date", T_DATE), ("fin_store_code", T_TEXT), ("fin_channel", T_TEXT),
                        ("entry_type", T_TEXT), ("revenue_recognized", T_CUR), ("refunds", T_CUR),
                        ("net_revenue", T_CUR), ("vat_amount", T_CUR), ("total_incl_vat", T_CUR)],
    "map_store_finance": [("fin_store_code", T_TEXT), ("store_key", T_INT), ("fin_channel", T_TEXT),
                          ("valid_from", T_DATE), ("valid_to", T_DATE), ("note", T_TEXT)],
    "dim_date": [("date", T_DATE), ("date_key", T_INT), ("year", T_INT), ("quarter", T_TEXT), ("month", T_INT),
                 ("month_id", T_INT), ("month_label", T_TEXT), ("day", T_INT), ("weekday_no", T_INT),
                 ("weekday_name", T_TEXT), ("is_weekend", T_BOOL), ("has_data", T_BOOL), ("month_end_date", T_DATE)],
    "dim_store": [("store_key", T_INT), ("store_code", T_TEXT), ("store_name", T_TEXT), ("channel_code", T_TEXT),
                  ("country", T_TEXT), ("region", T_TEXT), ("state", T_TEXT), ("square_meters", T_INT),
                  ("open_date", T_DATE)],
    "dim_product": [("product_key", T_INT), ("product_name", T_TEXT), ("brand", T_TEXT), ("color", T_TEXT),
                    ("subcategory_key", T_TEXT), ("subcategory", T_TEXT), ("category_key", T_TEXT),
                    ("category", T_TEXT), ("list_price_usd", T_CUR), ("unit_cost_usd", T_CUR)],
    "dim_channel": [("channel_code", T_TEXT), ("channel_name", T_TEXT), ("fin_channel", T_TEXT),
                    ("revenue_recognition_rule", T_TEXT)],
    "dim_promotion": [("promo_code", T_TEXT), ("promo_name", T_TEXT), ("discount_pct", T_PCT),
                      ("window_from_mmdd", T_TEXT), ("window_to_mmdd", T_TEXT), ("channel", T_TEXT)],
    "dim_order_status": [("order_status", T_TEXT), ("status_name", T_TEXT), ("count_in_revenue", T_TEXT),
                         ("rule", T_TEXT)],
}


def types_m(name, indent="        "):
    rows = [f'{{"{c}", {t}}}' for c, t in TYPES[name]]
    return "{\n" + ",\n".join(indent + "    " + r for r in rows) + "\n" + indent + "}"


def src(book, table, sheet=False):
    return (f'    Source = Excel.Workbook(\n        File.Contents(SourceFolder & "{book}"), null, true),\n'
            f'    Navigation = Source{{[Item = "tbl_{table}", Kind = "Table"]}}[Data],\n')


# ------------------------------------------------------------------ các query M (tên = tên bảng trong Data Model)
# Mỗi dòng code ≤ 80 ký tự để đọc được trong Advanced Editor và trên slide.
Q = {}
Q["SourceFolder"] = ('"C:\\NexaP1\\" meta [IsParameterQuery = true, Type = "Text",\n'
                     '    IsParameterQueryRequired = true]',
                     "Tham số: thư mục chứa 2 file dữ liệu (kết thúc bằng dấu \\). Sửa một chỗ này khi đổi máy.")
Q["fact_sales_line"] = ("let\n" + src(DS, "fact_sales_line") +
    "    ChangedType = Table.TransformColumnTypes(Navigation, " + types_m("fact_sales_line") + "),\n"
    "    // 120 dòng trùng hoàn toàn (mọi cột giống nhau) -> giữ 1\n"
    "    RemovedDuplicates = Table.Distinct(ChangedType),\n"
    "    AddedGross = Table.AddColumn(RemovedDuplicates, \"gross_amount\",\n"
    "        each [quantity] * [unit_price], Currency.Type),\n"
    "    AddedNet = Table.AddColumn(AddedGross, \"net_amount\",\n"
    "        each [gross_amount] - [discount_amount], Currency.Type)\n"
    "in\n    AddedNet", "Dòng bán: đặt kiểu, bỏ dòng trùng, thêm cột tiền dòng (gross, net).")
Q["fact_return"] = ("let\n" + src(DS, "fact_return") +
    "    ChangedType = Table.TransformColumnTypes(Navigation, " + types_m("fact_return") + "),\n"
    "    // lấy cửa hàng, kênh, sản phẩm từ dòng bán gốc\n"
    "    MergedSales = Table.NestedJoin(ChangedType, {\"order_number\", \"line_item\"},\n"
    "        fact_sales_line, {\"order_number\", \"line_item\"}, \"sales\",\n"
    "        JoinKind.LeftOuter),\n"
    "    ExpandedSales = Table.ExpandTableColumn(MergedSales, \"sales\",\n"
    "        {\"store_key\", \"channel_code\", \"product_key\"}),\n"
    "    ChangedType2 = Table.TransformColumnTypes(ExpandedSales,\n"
    "        {{\"store_key\", Int64.Type}, {\"channel_code\", type text},\n"
    "         {\"product_key\", Int64.Type}})\n"
    "in\n    ChangedType2", "Hoàn trả: đặt kiểu, ghép store_key/channel/product từ dòng bán gốc.")
Q["finance_raw"] = ("let\n" + src(FIN, "finance_revenue") +
    "    ChangedType = Table.TransformColumnTypes(Navigation, " + types_m("finance_revenue") + ")\n"
    "in\n    ChangedType", "File Finance gốc (chỉ tạo connection, không nạp vào Data Model).")
Q["map_store_finance"] = ("let\n" + src(DS, "map_store_finance") +
    "    ChangedType = Table.TransformColumnTypes(Navigation, " + types_m("map_store_finance") + ")\n"
    "in\n    ChangedType", "Bảng quy đổi mã Finance → store_key (chỉ tạo connection).")
Q["fact_finance"] = ("let\n"
    "    Source = finance_raw,\n"
    "    MergedMap = Table.NestedJoin(Source, {\"fin_store_code\"},\n"
    "        map_store_finance, {\"fin_store_code\"}, \"map\", JoinKind.LeftOuter),\n"
    "    ExpandedMap = Table.ExpandTableColumn(MergedMap, \"map\",\n"
    "        {\"store_key\", \"valid_from\", \"valid_to\"}),\n"
    "    // giữ mã có hiệu lực tại ngày hạch toán (S38: GBR-03 -> GBR-03N từ 01/07/2019)\n"
    "    FilteredValid = Table.SelectRows(ExpandedMap,\n"
    "        each [posting_date] >= [valid_from] and [posting_date] <= [valid_to]),\n"
    "    RemovedColumns = Table.RemoveColumns(FilteredValid, {\"valid_from\", \"valid_to\"}),\n"
    "    // HQ-ADJ: store_key = null -> không thuộc cửa hàng nào\n"
    "    ChangedType = Table.TransformColumnTypes(RemovedColumns,\n"
    "        {{\"store_key\", Int64.Type}})\n"
    "in\n    ChangedType", "Finance đã gắn store_key theo mã và ngày hiệu lực.")
for d in ["dim_date", "dim_store", "dim_product", "dim_channel", "dim_promotion", "dim_order_status"]:
    Q[d] = ("let\n" + src(DS, d) +
            "    ChangedType = Table.TransformColumnTypes(Navigation, " + types_m(d) + ")\n"
            "in\n    ChangedType", "Bảng dim: nạp và đặt kiểu.")
Q["_Measures"] = ('let\n    Source = #table(type table [ghi_chu = text],\n        {{"bảng chứa measure"}})\nin\n    Source',
                  "Bảng một dòng để chứa mọi measure (ẩn cột ghi_chu trong Power Pivot).")
LOAD = {"SourceFolder": "—", "finance_raw": "Only Create Connection", "map_store_finance": "Only Create Connection"}

# ------------------------------------------------------------------ DAX (Power Pivot)
DAX = f"""// Measures cho Data Model của Excel (Power Pivot). Tạo bằng Power Pivot → Measures → New Measure
// (Table name: _Measures), hoặc dán từng dòng "Tên := công thức" vào vùng tính toán (Calculation Area) của bảng _Measures.
// Tên measure giữ nguyên để CUBEVALUE và slide khớp nhau.

// ===== 1 · Doanh thu (chỉ đơn COMPLETED; dòng trùng đã bỏ ở Power Query)
Gross Sales := CALCULATE ( SUM ( fact_sales_line[gross_amount] ), KEEPFILTERS ( fact_sales_line[order_status] = "COMPLETED" ) )
Discounts := CALCULATE ( SUM ( fact_sales_line[discount_amount] ), KEEPFILTERS ( fact_sales_line[order_status] = "COMPLETED" ) )
Net Sales := [Gross Sales] - [Discounts]
Returns := SUM ( fact_return[return_amount] )
Net Revenue := [Net Sales] - [Returns]

// ===== 2 · Đơn hàng và giá
Orders := CALCULATE ( DISTINCTCOUNT ( fact_sales_line[order_number] ), KEEPFILTERS ( fact_sales_line[order_status] = "COMPLETED" ) )
Units := CALCULATE ( SUM ( fact_sales_line[quantity] ), KEEPFILTERS ( fact_sales_line[order_status] = "COMPLETED" ) )
AOV := DIVIDE ( [Net Sales], [Orders] )
ASP := DIVIDE ( [Net Sales], [Units] )
Return Rate := DIVIDE ( [Returns], [Net Sales] )
Giảm giá / Gross := DIVIDE ( [Discounts], [Gross Sales] )
Tỷ trọng online := DIVIDE ( CALCULATE ( [Net Sales], KEEPFILTERS ( fact_sales_line[channel_code] = "ONLINE" ) ), [Net Sales] )
Cancelled Orders := CALCULATE ( DISTINCTCOUNT ( fact_sales_line[order_number] ), fact_sales_line[order_status] = "CANCELLED" )
Pending Orders := CALCULATE ( DISTINCTCOUNT ( fact_sales_line[order_number] ), fact_sales_line[order_status] = "PENDING" )

// ===== 3 · Finance và đối soát
Finance Net Revenue := CALCULATE ( SUM ( fact_finance[net_revenue] ), NOT ISBLANK ( fact_finance[store_key] ) )
HQ Adjustment := CALCULATE ( SUM ( fact_finance[net_revenue] ), ISBLANK ( fact_finance[store_key] ) )
Δ Finance := [Finance Net Revenue] - [Net Revenue]
Δ Finance % := DIVIDE ( [Δ Finance], ABS ( [Finance Net Revenue] ) )
Δ tuyệt đối := ABS ( [Δ Finance] )
Net Sales (ngày giao) := CALCULATE ( [Net Sales], USERELATIONSHIP ( fact_sales_line[delivery_date], dim_date[date] ) )
Net Sales (cơ sở Finance) :=
    VAR pos = CALCULATE ( [Net Sales], KEEPFILTERS ( fact_sales_line[channel_code] = "POS" ) )
    VAR onl = CALCULATE ( [Net Sales (ngày giao)], KEEPFILTERS ( fact_sales_line[channel_code] = "ONLINE" ) )
    RETURN pos + onl
Δ sau quy đổi := [Finance Net Revenue] - ( [Net Sales (cơ sở Finance)] - [Returns] )
Vượt ngưỡng := IF ( ABS ( [Δ Finance] ) > {THR_ABS:.0f} && ABS ( [Δ Finance %] ) > {THR_PCT}, 1, 0 )
Số cửa hàng vượt ngưỡng := SUMX ( CROSSJOIN ( VALUES ( dim_date[month_id] ), VALUES ( dim_store[store_key] ) ), [Vượt ngưỡng] )
Lệch thời điểm online := [Δ Finance] - [Δ sau quy đổi]

// ===== 4 · So sánh kỳ (cần dim_date đã Mark as Date Table; chọn MỘT tháng trên slicer)
Net Revenue LM := CALCULATE ( [Net Revenue], DATEADD ( dim_date[date], -1, MONTH ) )
Net Revenue SPLY := CALCULATE ( [Net Revenue], SAMEPERIODLASTYEAR ( dim_date[date] ) )
Net Revenue YTD := CALCULATE ( [Net Revenue], DATESYTD ( dim_date[date] ) )
Net Revenue % LM := DIVIDE ( [Net Revenue] - [Net Revenue LM], [Net Revenue LM] )
Net Revenue % SPLY := DIVIDE ( [Net Revenue] - [Net Revenue SPLY], [Net Revenue SPLY] )
Net Sales SPLY := CALCULATE ( [Net Sales], SAMEPERIODLASTYEAR ( dim_date[date] ) )
Net Sales % SPLY := DIVIDE ( [Net Sales] - [Net Sales SPLY], [Net Sales SPLY] )
Orders LM := CALCULATE ( [Orders], DATEADD ( dim_date[date], -1, MONTH ) )
AOV LM := CALCULATE ( [AOV], DATEADD ( dim_date[date], -1, MONTH ) )
Orders % LM := DIVIDE ( [Orders] - [Orders LM], [Orders LM] )
AOV % LM := DIVIDE ( [AOV] - [AOV LM], [AOV LM] )
Return Rate LM := CALCULATE ( [Return Rate], DATEADD ( dim_date[date], -1, MONTH ) )
Orders SPLY := CALCULATE ( [Orders], SAMEPERIODLASTYEAR ( dim_date[date] ) )
AOV SPLY := CALCULATE ( [AOV], SAMEPERIODLASTYEAR ( dim_date[date] ) )
Orders % SPLY := DIVIDE ( [Orders] - [Orders SPLY], [Orders SPLY] )
AOV % SPLY := DIVIDE ( [AOV] - [AOV SPLY], [AOV SPLY] )
Return Rate SPLY := CALCULATE ( [Return Rate], SAMEPERIODLASTYEAR ( dim_date[date] ) )

// ===== 5 · Hỗ trợ biểu đồ Excel (tách 2 chuỗi để tô màu nhấn mà không phải tô tay sau mỗi lần đổi slicer)
Net Sales danh mục dẫn đầu := VAR m = MAXX ( ALLSELECTED ( dim_product[category] ), [Net Sales] ) RETURN IF ( [Net Sales] = m, [Net Sales] )
Net Sales danh mục khác := IF ( ISBLANK ( [Net Sales danh mục dẫn đầu] ), [Net Sales] )
Δ tháng có vượt ngưỡng := IF ( [Số cửa hàng vượt ngưỡng] > 0, [Δ sau quy đổi] )
Δ tháng khác := IF ( [Số cửa hàng vượt ngưỡng] = 0, [Δ sau quy đổi] )

// ===== 6 · Chuỗi cho tiêu đề hành động và câu Điểm chính (đưa ra ô bằng CUBEVALUE)
Tháng chọn := IF ( HASONEVALUE ( dim_date[month_label] ), VALUES ( dim_date[month_label] ), "nhiều tháng" )
Điểm chính Tổng quan :=
    "Net Revenue " & [Tháng chọn] & " đạt " & FORMAT ( [Net Revenue] / 1000, "#,##0.0" ) & " nghìn USD — "
    & FORMAT ( [Net Revenue % LM], "+0.0%;-0.0%" ) & " so với tháng trước, "
    & FORMAT ( [Net Revenue % SPLY], "+0.0%;-0.0%" ) & " so với cùng kỳ năm trước."
Điểm chính Động lực :=
    IF ( ABS ( [Orders % SPLY] ) >= ABS ( [AOV % SPLY] ),
        "Biến động so với cùng kỳ đến chủ yếu từ số đơn (" & FORMAT ( [Orders % SPLY], "+0.0%;-0.0%" ) & "); AOV " & FORMAT ( [AOV % SPLY], "+0.0%;-0.0%" ) & ".",
        "Biến động so với cùng kỳ đến chủ yếu từ giá trị mỗi đơn (AOV " & FORMAT ( [AOV % SPLY], "+0.0%;-0.0%" ) & "); số đơn " & FORMAT ( [Orders % SPLY], "+0.0%;-0.0%" ) & "." )
Điểm chính Đối soát :=
    "Finance " & IF ( [Δ Finance] >= 0, "cao", "thấp" ) & " hơn dữ liệu đơn " & FORMAT ( ABS ( [Δ Finance] ), "#,##0" )
    & " USD; sau khi quy online về ngày giao còn " & FORMAT ( [Δ sau quy đổi], "#,##0" ) & " USD; "
    & [Số cửa hàng vượt ngưỡng] & " cửa hàng vượt ngưỡng."
Tiêu đề danh mục :=
    VAR topCat = TOPN ( 1, VALUES ( dim_product[category] ), [Net Sales] )
    RETURN MAXX ( topCat, dim_product[category] ) & " chiếm "
        & FORMAT ( DIVIDE ( CALCULATE ( [Net Sales], topCat ), [Net Sales] ), "0%" ) & " Net Sales"
Tiêu đề hoàn trả :=
    "Tỷ lệ hoàn trả " & FORMAT ( [Return Rate], "0.00%" ) & " — "
    & IF ( [Return Rate] > [Return Rate SPLY], "cao", "thấp" ) & " hơn cùng kỳ "
    & FORMAT ( ABS ( [Return Rate] - [Return Rate SPLY] ) * 100, "0.00" ) & " điểm %"
"""

# ------------------------------------------------------------------ chạy song song bằng pandas: số dòng sau từng bước
X = pd.read_excel(P1 / DS, sheet_name=None)
finance = pd.read_excel(P1 / FIN, sheet_name="finance_revenue")
steps = {}
f0 = X["fact_sales_line"]
f1 = f0.drop_duplicates()
f1 = f1.assign(gross_amount=f1.quantity * f1.unit_price)
f1 = f1.assign(net_amount=f1.gross_amount - f1.discount_amount)
steps["fact_sales_line"] = [("Source", "Excel.Workbook", None), ("Navigation", "tbl_fact_sales_line", len(f0)),
                            ("Changed Type", "Table.TransformColumnTypes", len(f0)),
                            ("Removed Duplicates", "Table.Distinct", len(f1)),
                            ("Added gross_amount", "Table.AddColumn", len(f1)),
                            ("Added net_amount", "Table.AddColumn", len(f1))]
r0 = X["fact_return"]
r1 = r0.merge(f1[["order_number", "line_item", "store_key", "channel_code", "product_key"]],
              on=["order_number", "line_item"], how="left")
assert r1.store_key.notna().all(), "mọi dòng trả hàng phải ghép được dòng bán"
steps["fact_return"] = [("Source", "Excel.Workbook", None), ("Navigation", "tbl_fact_return", len(r0)),
                        ("Changed Type", "Table.TransformColumnTypes", len(r0)),
                        ("Merged Queries", "Table.NestedJoin ← fact_sales_line", len(r1)),
                        ("Expanded sales", "Table.ExpandTableColumn", len(r1)),
                        ("Changed Type2", "Table.TransformColumnTypes", len(r1))]
mp = X["map_store_finance"]
m1 = finance.merge(mp[["fin_store_code", "store_key", "valid_from", "valid_to"]], on="fin_store_code", how="left")
m2 = m1[(pd.to_datetime(m1.posting_date) >= m1.valid_from) & (pd.to_datetime(m1.posting_date) <= m1.valid_to)]
steps["finance_raw"] = [("Source", "Excel.Workbook", None), ("Navigation", "tbl_finance_revenue", len(finance)),
                        ("Changed Type", "Table.TransformColumnTypes", len(finance))]
steps["map_store_finance"] = [("Source", "Excel.Workbook", None), ("Navigation", "tbl_map_store_finance", len(mp)),
                              ("Changed Type", "Table.TransformColumnTypes", len(mp))]
steps["fact_finance"] = [("Source", "= finance_raw", len(finance)),
                         ("Merged Queries", "Table.NestedJoin ← map_store_finance", len(m1)),
                         ("Expanded map", "Table.ExpandTableColumn", len(m1)),
                         ("Filtered Valid", "Table.SelectRows (ngày hiệu lực)", len(m2)),
                         ("Removed Columns", "Table.RemoveColumns", len(m2)),
                         ("Changed Type", "Table.TransformColumnTypes", len(m2))]
dims = {d: len(X[d]) for d in ["dim_date", "dim_store", "dim_product", "dim_channel", "dim_promotion", "dim_order_status"]}
rows = REF["rows"]
assert len(f1) == rows["fact_sales_line_clean"] and len(f0) == rows["fact_sales_line"]
assert len(r1) == rows["fact_return"] and len(m2) == rows["finance_revenue"] == len(finance)
assert int(m2.store_key.isna().sum()) == 3, "3 dòng HQ-ADJ giữ store_key trống"
# Net Revenue 12/2020 theo đúng logic M + DAX (đối chiếu số tham chiếu)
c = f1[(f1.order_status == "COMPLETED") & (pd.to_datetime(f1.order_date).dt.strftime("%Y%m") == str(REF["report_month"]))]
rr = r0[pd.to_datetime(r0.return_date).dt.strftime("%Y%m") == str(REF["report_month"])]
assert abs(c.net_amount.sum() - rr.return_amount.sum() - REF["month"]["current"]["net_revenue"]) < 0.05

# [Vượt ngưỡng] theo đúng logic DAX: ô tháng × store_key, Δ = Finance (bỏ HQ-ADJ) − (Net Sales theo order_date − Returns theo return_date)
ym = lambda s_: pd.to_datetime(s_).dt.year * 100 + pd.to_datetime(s_).dt.month
cc = f1[f1.order_status == "COMPLETED"]
ns = cc.assign(m=ym(cc.order_date)).groupby(["m", "store_key"]).net_amount.sum()
rt = r1.assign(m=ym(r1.return_date)).groupby(["m", "store_key"]).return_amount.sum()
fn = m2[m2.store_key.notna()].assign(m=ym(m2[m2.store_key.notna()].posting_date)).groupby(["m", "store_key"]).net_revenue.sum()
cell = pd.concat({"ns": ns, "rt": rt, "fn": fn}, axis=1).fillna(0)
cell["d"] = cell.fn - (cell.ns - cell.rt)
cell["pct"] = (cell.d / cell.fn.abs().where(cell.fn != 0)).fillna(0)   # DIVIDE(…, 0) = BLANK -> không vượt ngưỡng
cell["flag"] = (cell.d.abs() > REF["threshold"]["abs_usd"]) & (cell.pct.abs() > REF["threshold"]["pct"])
n_flag = int(cell[cell.index.get_level_values(0) // 100 == 2020].flag.sum())
assert n_flag == len(REF["over_threshold_2020"]), n_flag
assert int(cell.flag.sum()) == REF["over_threshold_cells"], int(cell.flag.sum())

# ------------------------------------------------------------------ ghi file
vn = lambda n: f"{n:,}".replace(",", ".")
(KIT / "m").mkdir(parents=True, exist_ok=True)
ORDER = ["SourceFolder", "dim_date", "dim_store", "dim_product", "dim_channel", "dim_promotion", "dim_order_status",
         "fact_sales_line", "fact_return", "finance_raw", "map_store_finance", "fact_finance", "_Measures"]
for i, name in enumerate(ORDER, 1):
    code, desc = Q[name]
    (KIT / "m" / f"{i:02d}_{name}.pq").write_text(f"// {name}: {desc}\n// Nạp: {LOAD.get(name, 'Only Create Connection + Add this data to the Data Model')}\n{code}\n",
                                                  encoding="utf-8")
(KIT / "measures.dax").write_text(DAX, encoding="utf-8")
(KIT / "README.md").write_text(f"""# Bộ dựng dashboard Excel — Sales Project 1

Yêu cầu: Excel cho Windows (Microsoft 365 hoặc 2016+), bật add-in **Microsoft Power Pivot for Excel**
(File → Options → Add-ins → Manage: COM Add-ins → Go… → tick Power Pivot).

1. **Power Query (M)** — Data → Get Data → From Other Sources → Blank Query → Advanced Editor, dán lần lượt các file
   trong `m/` theo số thứ tự, đặt tên query đúng tên file (bỏ số). Sửa `SourceFolder` thành thư mục chứa
   `{DS}` và `{FIN}`.
2. **Nạp** — Home → Close & Load To… → *Only Create Connection* + tick *Add this data to the Data Model*
   (riêng `finance_raw`, `map_store_finance`: chỉ *Only Create Connection*).
3. **Data Model** — Power Pivot → Manage → Diagram View: tạo quan hệ theo guideline mục 3.1, Mark as Date Table cho dim_date.
4. **DAX** — dán các measure trong `measures.dax` vào bảng `_Measures` (Power Pivot → Measures → New Measure).
5. **Dashboard** — 4 sheet: PivotTable/PivotChart từ Data Model, Slicer Tháng/Kênh/Quốc gia nối mọi pivot
   (Report Connections), thẻ KPI và câu Điểm chính bằng CUBEVALUE.

Số dòng phải khớp sau khi nạp: fact_sales_line {vn(rows['fact_sales_line_clean'])} · fact_return {vn(rows['fact_return'])} ·
fact_finance {vn(rows['finance_revenue'])}. Code M và DAX ở đây dùng lại được nguyên vẹn nếu sau này chuyển sang Power BI.
""", encoding="utf-8")
out = {"steps": steps, "dims": dims, "order": ORDER, "load": LOAD,
       "queries": {k: {"code": v[0], "desc": v[1]} for k, v in Q.items()}}
(HERE / "pq_steps.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
print("excel_build:", len(ORDER), "query M,", DAX.count(":="), "measure; số dòng:",
      {k: v[-1][2] for k, v in steps.items()})
