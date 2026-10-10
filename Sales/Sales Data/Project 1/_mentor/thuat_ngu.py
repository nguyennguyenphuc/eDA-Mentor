# -*- coding: utf-8 -*-
"""Bảng thuật ngữ dùng chung cho guideline, slide, notes và README dataset.

Căn cứ (tra cứu 10/2026):
- Doanh thu: Chuẩn mực kế toán Việt Nam VAS 14 "Doanh thu và thu nhập khác"; chế độ kế toán doanh nghiệp
  Thông tư 99/2025/TT-BTC (thay Thông tư 200/2014/TT-BTC từ 01/01/2026) – TK 521 "Các khoản giảm trừ doanh thu".
- Tên giao diện Excel: trang hỗ trợ tiếng Việt của Microsoft (support.microsoft.com/vi-vn): Mô hình Dữ liệu,
  Làm mới Tất cả, Đóng & Tải Tới, Đánh dấu là Bảng ngày tháng, Dạng xem Sơ đồ, Quản lý Quan hệ, Định dạng có Điều kiện,
  Đường lưới; PivotTable, PivotChart, Slicer, Power Pivot, Power Query, CUBEVALUE giữ tiếng Anh.
- Phân tích dữ liệu: lược đồ hình sao (star schema), bảng dữ kiện (fact table), bảng chiều (dimension table).
- Vai trò: không có tên chính thức → theo thực tế tuyển dụng: phòng Tài chính – Kế toán, Trưởng phòng Kinh doanh,
  chuyên viên phân tích dữ liệu (DA), cửa hàng trưởng, bộ phận Thương mại điện tử.

Quy tắc: định danh kỹ thuật (tên file, cột, bảng, số đo trong ngoặc vuông hoặc tên số đo tiếng Anh như
"Finance Net Revenue", "Δ Finance") giữ nguyên; chỉ đổi chữ trong văn xuôi.
"""
import re

# (thuật ngữ chuẩn, tiếng Anh, ghi chú/nguồn) — in ra Phụ lục D của guideline và slide thuật ngữ
BANG = [
    ("Nhóm", "Thuật ngữ dùng trong tài liệu", "Tiếng Anh / tên trong file", "Căn cứ"),
    ("Vai trò", "phòng Tài chính – Kế toán (gọi tắt: Kế toán)", "Finance", "thực tế doanh nghiệp"),
    ("Vai trò", "Trưởng phòng Kinh doanh", "Sales Manager", "thực tế doanh nghiệp"),
    ("Vai trò", "chuyên viên phân tích dữ liệu", "Data Analyst (DA)", "thực tế tuyển dụng"),
    ("Vai trò", "cửa hàng trưởng · bộ phận Thương mại điện tử", "Store manager · E-commerce", "thực tế bán lẻ"),
    ("Doanh thu", "doanh thu bán hàng", "Gross Sales", "VAS 14; mẫu B02-DN"),
    ("Doanh thu", "các khoản giảm trừ doanh thu", "Revenue deductions", "TK 521, TT 99/2025/TT-BTC"),
    ("Doanh thu", "chiết khấu thương mại (khuyến mãi giảm giá trên hoá đơn)", "Discounts", "VAS 14"),
    ("Doanh thu", "hàng bán bị trả lại", "Returns", "VAS 14"),
    ("Doanh thu", "doanh thu sau chiết khấu (chỉ tiêu trung gian)", "Net Sales", "chỉ tiêu quản trị"),
    ("Doanh thu", "doanh thu thuần", "Net Revenue", "VAS 14; TT 99/2025/TT-BTC"),
    ("Doanh thu", "ghi nhận doanh thu · ngày ghi sổ", "Revenue recognition · posting_date", "VAS 14; sổ Nhật ký chung"),
    ("Doanh thu", "bút toán điều chỉnh · khoá sổ cuối tháng", "Manual adjustment (HQ-ADJ) · month-end close", "chế độ kế toán"),
    ("Doanh thu", "thuế GTGT", "VAT", "Luật Thuế GTGT"),
    ("Chỉ tiêu", "chỉ số hiệu suất then chốt", "KPI", "Microsoft vi-vn (Power Pivot KPI)"),
    ("Chỉ tiêu", "giá trị đơn hàng trung bình", "AOV", "thực tế bán lẻ"),
    ("Chỉ tiêu", "giá bán bình quân", "ASP", "thực tế bán lẻ"),
    ("Chỉ tiêu", "tỷ lệ hàng bán bị trả lại", "Return Rate", "VAS 14 + thực tế"),
    ("Chỉ tiêu", "so với tháng trước · cùng kỳ năm trước · luỹ kế từ đầu năm", "LM · SPLY · YTD", "mẫu B02-DN (luỹ kế)"),
    ("Chỉ tiêu", "chênh lệch · ngưỡng · đối soát (đối chiếu số liệu)", "Δ · threshold · reconciliation", "thực tế kế toán"),
    ("Excel", "Mô hình Dữ liệu", "Data Model", "Microsoft vi-vn"),
    ("Excel", "số đo", "measure", "Microsoft vi-vn (Power Pivot)"),
    ("Excel", "mối quan hệ · Dạng xem Sơ đồ · Quản lý Quan hệ", "relationship · Diagram View · Manage Relationships", "Microsoft vi-vn"),
    ("Excel", "Đánh dấu là Bảng ngày tháng", "Mark as Date Table", "Microsoft vi-vn"),
    ("Excel", "Đóng & Tải Tới · Làm mới Tất cả", "Close & Load To · Refresh All", "Microsoft vi-vn"),
    ("Excel", "Định dạng có Điều kiện · Đường lưới · bổ trợ", "Conditional Formatting · Gridlines · add-in", "Microsoft vi-vn"),
    ("Excel", "giữ tên tiếng Anh", "PivotTable · PivotChart · Slicer · Power Query · Power Pivot · CUBEVALUE", "Microsoft vi-vn"),
    ("Dữ liệu", "lược đồ hình sao · bảng dữ kiện · bảng chiều", "star schema · fact table · dimension table", "giáo trình CSDL / kho dữ liệu"),
    ("Dữ liệu", "độ chi tiết (mỗi dòng là gì) · làm sạch dữ liệu", "grain · data cleaning", "giáo trình phân tích dữ liệu"),
    ("Trình bày", "kể chuyện bằng dữ liệu", "Storytelling with Data (C. N. Knaflic)", "tên sách giữ tiếng Anh"),
]



def _thay_finance(t):
    def f(m):
        s, e = m.start(), m.end()
        truoc, sau = t[max(0, s - 7):s], t[e:e + 12]
        # giữ nguyên: định danh (Δ Finance, [Finance Net Revenue], cơ sở Finance, fact_finance) và chú thích "(Finance"
        if (truoc.endswith(("Δ ", "cơ sở ", "_", "[", "-", ".", "(")) or sau.startswith((" Net Revenue", "_", " %", "]"))
                or truoc[-1:].isalnum()):
            return m.group(0)
        return "Kế toán"
    return re.sub(r"Finance", f, t)


QUY_TAC = [   # (mẫu regex, thay bằng) — áp cho mọi chế độ
    (r"Sales Manager", "Trưởng phòng Kinh doanh"),
    (r"ngày hạch toán", "ngày ghi sổ"),
    (r"(?<![\w_\[.])measure(?![\w.])", "số đo"),
    (r"(?<![\w_\[.])measures(?![\w.])", "số đo"),
    (r"(?<!\()\badd-in\b", "bổ trợ (add-in)"),
    (r"bảng fact(?![_\w])", "bảng dữ kiện"),
    (r"bảng dim(?![_\w])", "bảng chiều"),
    (r"(?<![\w(])star schema", "lược đồ hình sao"),
    (r"(?<!\()Storytelling with Data(?! \()", "kể chuyện bằng dữ liệu (Storytelling with Data)"),
]
QUY_TAC_UI = [  # tên giao diện: tiếng Việt của Microsoft + tiếng Anh trong ngoặc (guideline, notes)
    (r"(?<![(\w])Mark as Date Table", "Đánh dấu là Bảng ngày tháng (Mark as Date Table)"),
    (r"(?<![(\w])Close & Load To(…)?", "Đóng & Tải Tới (Close & Load To)"),
    (r"(?<![(\w])Refresh All", "Làm mới Tất cả (Refresh All)"),
    (r"(?<![(\w])Diagram View", "Dạng xem Sơ đồ (Diagram View)"),
    (r"(?<![(\w])Manage Relationships", "Quản lý Quan hệ (Manage Relationships)"),
    (r"(?<![(\w])Conditional Formatting", "Định dạng có Điều kiện (Conditional Formatting)"),
    (r"(?<![(\w])Gridlines", "Đường lưới (Gridlines)"),
    (r"(?<![(\w])Calculation Area", "vùng tính toán (Calculation Area)"),
    (r"Data Model", lambda m: m.group(0) if m.string[max(0, m.start() - 5):m.start()].endswith(("the ", "From ", "("))
                               else _lan_dau("Data Model", "Mô hình Dữ liệu")),
]
_DA_GAP = set()


def _lan_dau(en, vi):
    """Lần đầu trong một tài liệu: "tiếng Việt (tiếng Anh)"; các lần sau chỉ tiếng Việt."""
    if en in _DA_GAP:
        return vi
    _DA_GAP.add(en)
    return f"{vi} ({en})"


def bat_dau():
    """Gọi khi bắt đầu một tài liệu mới."""
    _DA_GAP.clear()


def chuan_hoa(t, ui=True):
    """Đổi thuật ngữ trong một đoạn văn xuôi. ui=False: giữ tên giao diện tiếng Anh (slide chật chỗ)."""
    if not t or not isinstance(t, str):
        return t
    t = _thay_finance(t)
    for a, b in QUY_TAC + (QUY_TAC_UI if ui else []):
        t = re.sub(a, b, t)
    # gộp lặp "X (X)" sinh ra khi văn bản đã có sẵn chú thích
    t = re.sub(r"Mô hình Dữ liệu \(Data Model\) \(Data Model\)", "Mô hình Dữ liệu (Data Model)", t)
    return t


if __name__ == "__main__":
    for s in ["Finance ghi nhận online theo ngày giao", "[Finance Net Revenue] và Δ Finance %",
              "file Finance (finance_revenue.xlsx)", "Sales Manager cần 3 câu", "viết measure DAX trong bảng _Measures",
              "Close & Load To… → Add this data to the Data Model", "Mark as Date Table cho dim_date",
              "Net Sales (cơ sở Finance)", "PivotTable → From Data Model"]:
        print(s, "→", chuan_hoa(s))
