# -*- coding: utf-8 -*-
"""Các bước Power Query làm BẰNG CHUỘT cho từng query — một nguồn cho guideline (Bước 2), slide (sơ đồ pipeline)
và sơ đồ 05. Không có code M: học viên bấm trên giao diện, Excel tự ghi từng thao tác thành một Applied Step.

Mỗi bước: (tên Applied Step Excel tự đặt — giao diện tiếng Anh, nhãn ngắn, thao tác chuột, số dòng sau bước hoặc None, loại)
loại: "nap" = mở nguồn · "kieu" = đặt kiểu · "dong" = đổi số dòng hoặc ghép bảng (phải kiểm số dòng) · "cot" = thêm/bớt cột.
"""
import json
from pathlib import Path

_R = json.loads((Path(__file__).resolve().parent / "ref_numbers.json").read_text(encoding="utf-8"))["rows"]
N_SALES, N_SALES_SACH = _R["fact_sales_line"], _R["fact_sales_line_clean"]
N_RET, N_FIN, N_MAP = _R["fact_return"], _R["finance_revenue"], _R["map_store_finance"]
DIMS = {"dim_date": _R["dim_date"], "dim_store": _R["dim_store"], "dim_product": _R["dim_product"],
        "dim_channel": 2, "dim_promotion": 7, "dim_order_status": 3}

NAP_DATASET = ("Data → Get Data → From File → From Excel Workbook → chọn Nexa_Sales_P1_Dataset.xlsx → cửa sổ Navigator "
               "tick Select multiple items, chọn các **Table** tbl_… (biểu tượng bảng xanh, không chọn sheet cùng tên) → "
               "Transform Data")
DOI_TEN = "Query Settings → Name: đổi tbl_dim_date thành dim_date (bỏ tiền tố tbl_), làm tương tự cho mọi query"

STEPS = {
    "fact_sales_line": [
        ("Source", "mở file nguồn", "Excel mở file nguồn (tự tạo khi chọn file)", None, "nap"),
        ("Navigation", "chọn Table tbl_fact_sales_line", "Chọn Table tbl_fact_sales_line trong Navigator", N_SALES, "nap"),
        ("Changed Type", "đặt kiểu: Date, Whole Number, Fixed Decimal", "Kiểm kiểu từng cột: ngày → Date; tiền (unit_price, discount_amount, unit_cost) → Fixed Decimal "
                         "Number ($); số lượng, khoá → Whole Number; còn lại Text. Bấm biểu tượng kiểu ở đầu cột, chọn "
                         "Replace current", N_SALES, "kieu"),
        ("Removed Duplicates", "Ctrl+A → Remove Rows → Remove Duplicates", "Ctrl+A chọn mọi cột → Home → Remove Rows → Remove Duplicates", N_SALES_SACH, "dong"),
        ("Inserted Multiplication", "quantity × unit_price → gross_amount", "Chọn quantity, giữ Ctrl chọn unit_price → Add Column → Standard → Multiply; "
                                    "bấm đúp tiêu đề cột mới, đổi tên thành gross_amount", N_SALES_SACH, "cot"),
        ("Inserted Subtraction", "gross_amount − discount_amount → net_amount", "Chọn gross_amount trước, giữ Ctrl chọn discount_amount sau → Add Column → Standard → "
                                 "Subtract; đổi tên thành net_amount", N_SALES_SACH, "cot"),
        ("Changed Type1", "2 cột tiền → Fixed Decimal", "Đặt gross_amount, net_amount về Fixed Decimal Number", N_SALES_SACH, "kieu"),
    ],
    "fact_return": [
        ("Source · Navigation", "chọn Table tbl_fact_return", "Chọn Table tbl_fact_return (cùng lượt với các bảng khác)", N_RET, "nap"),
        ("Changed Type", "đặt kiểu cột", "return_date → Date; return_amount → Fixed Decimal Number; khoá → Whole Number", N_RET, "kieu"),
        ("Merged Queries", "Merge với fact_sales_line theo order_number + line_item", "Home → Merge Queries → bảng dưới chọn fact_sales_line; bấm order_number rồi giữ Ctrl bấm "
                           "line_item ở **cả hai** bảng theo cùng thứ tự; Join Kind: Left Outer → OK", N_RET, "dong"),
        ("Expanded fact_sales_line", "lấy store_key, channel_code, product_key", "Bấm nút mở rộng ⇄ ở cột mới → chỉ tick store_key, channel_code, product_key, "
                                     "bỏ tick Use original column name as prefix", N_RET, "cot"),
        ("Changed Type1", "đặt kiểu 3 cột mới", "Đặt kiểu cho 3 cột vừa lấy (Whole Number, Text, Whole Number)", N_RET, "kieu"),
    ],
    "finance_raw": [
        ("Source · Navigation", "chọn Table tbl_finance_revenue", "Get Data → From Excel Workbook → finance_revenue.xlsx → Table tbl_finance_revenue → "
                                "Transform Data", N_FIN, "nap"),
        ("Changed Type", "đặt kiểu cột", "posting_date → Date; các cột tiền → Fixed Decimal Number", N_FIN, "kieu"),
    ],
    "map_store_finance": [
        ("Source · Navigation", "chọn Table tbl_map_store_finance", "Table tbl_map_store_finance (chọn cùng lượt ở Navigator)", N_MAP, "nap"),
        ("Changed Type", "đặt kiểu cột", "valid_from, valid_to → Date; store_key → Whole Number", N_MAP, "kieu"),
    ],
    "fact_finance": [
        ("Source", "Reference từ finance_raw", "Ở khung Queries chuột phải finance_raw → Reference → đổi tên query mới thành fact_finance", N_FIN, "nap"),
        ("Merged Queries", "Merge với map_store_finance theo fin_store_code", "Home → Merge Queries → chọn map_store_finance, bấm cột fin_store_code ở cả hai bảng, Left Outer",
         N_FIN, "dong"),
        ("Expanded map_store_finance", "lấy store_key, valid_from, valid_to", "Mở rộng ⇄ → tick store_key, valid_from, valid_to; bỏ tick tiền tố", N_FIN, "cot"),
        ("Added Conditional Column", "hieu_luc = 1 nếu valid_from ≤ posting_date ≤ valid_to", "Add Column → Conditional Column, tên hieu_luc: If posting_date **is before** "
                                     "valid_from (ô Value chọn Select a column) Then 0; Else If posting_date **is after** "
                                     "valid_to Then 0; Else 1", N_FIN, "cot"),
        ("Filtered Rows", "giữ hieu_luc = 1", "Bấm nút lọc ở cột hieu_luc → bỏ tick 0 (chỉ giữ dòng mã còn hiệu lực ở ngày ghi sổ)", N_FIN, "dong"),
        ("Removed Columns", "bỏ valid_from, valid_to, hieu_luc", "Chọn valid_from, valid_to, hieu_luc → chuột phải → Remove Columns", N_FIN, "cot"),
        ("Changed Type", "store_key → Whole Number", "store_key → Whole Number (3 dòng HQ-ADJ để trống — đúng thiết kế)", N_FIN, "kieu"),
    ],
}

# query, cách tạo, các bước tóm tắt, nạp, số dòng — bảng "Thứ tự tạo query" (Bước 2.2)
ORDER = (
    [(d, "Navigator (lượt 1)", "Source → Navigation → Changed Type", "Connection + Data Model", DIMS[d]) for d in DIMS] +
    [("fact_sales_line", "Navigator (lượt 1)", "Changed Type → Removed Duplicates → Multiply → Subtract", "Connection + Data Model", N_SALES_SACH),
     ("fact_return", "Navigator (lượt 1)", "Changed Type → Merge với fact_sales_line → Expand", "Connection + Data Model", N_RET),
     ("map_store_finance", "Navigator (lượt 1)", "Changed Type", "Only Create Connection", N_MAP),
     ("finance_raw", "Get Data lượt 2 (finance_revenue.xlsx)", "Changed Type", "Only Create Connection", N_FIN),
     ("fact_finance", "Reference từ finance_raw", "Merge với map_store_finance → Expand → Conditional Column → Filter", "Connection + Data Model", N_FIN),
     ("_Measures", "Bảng 1 ô trên sheet → From Table/Range", "Source", "Connection + Data Model", 1)]
)
