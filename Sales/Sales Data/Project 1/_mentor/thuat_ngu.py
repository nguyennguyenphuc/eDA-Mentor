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
    ("Doanh thu", "doanh thu theo giá niêm yết (trước giảm giá)", "Gross Sales", "chỉ tiêu quản trị"),
    ("Doanh thu", "giảm giá khuyến mại trên hoá đơn (chỉ tiêu quản trị; kế toán ghi doanh thu theo giá đã giảm)", "Discounts", "VAS 14; TK 521, TT 99/2025/TT-BTC"),
    ("Doanh thu", "hàng bán bị trả lại", "Returns", "VAS 14"),
    ("Doanh thu", "doanh thu sau chiết khấu (chỉ tiêu trung gian)", "Net Sales", "chỉ tiêu quản trị"),
    ("Doanh thu", "doanh thu thuần", "Net Revenue", "VAS 14; TT 99/2025/TT-BTC"),
    ("Doanh thu", "ghi nhận doanh thu · ngày ghi sổ", "Revenue recognition · posting_date", "VAS 14; sổ Nhật ký chung"),
    ("Doanh thu", "bút toán điều chỉnh · khoá sổ cuối tháng", "Manual adjustment (HQ-ADJ) · month-end close", "chế độ kế toán"),
    ("Doanh thu", "thuế GTGT", "VAT", "Luật Thuế GTGT"),
    ("Chỉ tiêu", "chỉ số hiệu suất then chốt", "KPI", "Microsoft vi-vn (Power Pivot KPI)"),
    ("Chỉ tiêu", "giá trị đơn hàng trung bình", "AOV", "thực tế bán lẻ"),
    ("Chỉ tiêu", "giá bán bình quân", "ASP", "thực tế bán lẻ"),
    ("Chỉ tiêu", "tỷ lệ hàng bán bị trả lại", "Return Rate", "chỉ tiêu quản trị"),
    ("Chỉ tiêu", "so với tháng trước · cùng kỳ năm trước · luỹ kế từ đầu năm", "LM · SPLY · YTD", "mẫu B02-DN (luỹ kế)"),
    ("Chỉ tiêu", "chênh lệch · ngưỡng · đối soát (đối chiếu số liệu)", "Δ · threshold · reconciliation", "thực tế kế toán"),
    ("Excel", "Mô hình Dữ liệu", "Data Model", "Microsoft vi-vn"),
    ("Excel", "số đo", "measure", "Microsoft vi-vn (Power Pivot)"),
    ("Excel", "mối quan hệ · Dạng xem Sơ đồ · Quản lý Quan hệ", "relationship · Diagram View · Manage Relationships", "Microsoft vi-vn"),
    ("Excel", "Đánh dấu là Bảng ngày tháng", "Mark as Date Table", "Microsoft vi-vn"),
    ("Excel", "Đóng & Tải Tới · Làm mới Tất cả", "Close & Load To · Refresh All", "Microsoft vi-vn"),
    ("Excel", "Định dạng có Điều kiện · Đường lưới · bổ trợ", "Conditional Formatting · Gridlines · add-in", "Microsoft vi-vn"),
    ("Excel", "giữ tên tiếng Anh", "PivotTable · PivotChart · Slicer · Power Query · Power Pivot · CUBEVALUE", "Microsoft vi-vn"),
    ("Excel", "Loại bỏ Nội dung trùng lặp · Hợp nhất truy vấn · Cột có Điều kiện", "Remove Duplicates · Merge Queries · Conditional Column", "Microsoft vi-vn"),
    ("Excel", "Các bước Đã áp dụng · Bộ dẫn hướng · Thiết đặt Nguồn Dữ liệu", "Applied Steps · Navigator · Data Source Settings", "Microsoft vi-vn"),
    ("Excel", "Kết nối Báo cáo · Bộ lọc Giá trị · Đặt làm tổng (biểu đồ thác nước)", "Report Connections · Value Filters · Set as total (Waterfall)", "Microsoft vi-vn"),
    ("Dữ liệu", "lược đồ hình sao · bảng dữ kiện (còn gọi bảng sự kiện) · bảng chiều", "star schema · fact table · dimension table", "bài giảng mô hình dữ liệu đa chiều; NEU TIKT1122 (kho dữ liệu, tri thức kinh doanh)"),
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
    # rà theo giáo trình (10/2026): Luật Thương mại "khuyến mại"; VAS 14; trang Microsoft vi-vn; đề cương NEU TIKT1122
    (r"\bkhuyến mãi", "khuyến mại"), (r"\bKhuyến mãi", "Khuyến mại"),
    (r"văn phòng chính", "trụ sở chính"),
    (r"\b([Bb])ảng map\b", lambda m: ("B" if m.group(1) == "B" else "b") + "ảng ánh xạ"),
    (r"phải map\b", "phải ánh xạ"),
    (r"\b[Mm]ap Kế toán theo valid_from ?/ ?valid_to", "Gắn mã cửa hàng Kế toán theo ngày hiệu lực (valid_from, valid_to)"),
    (r"\bHạt dữ liệu", "Độ chi tiết"), (r"\bhạt dữ liệu", "độ chi tiết"),
    (r"^Outline$", "Nội dung"),
    (r"data modeling", "mô hình hoá dữ liệu"),
    (r"Gỡ rối", "Loại bỏ chi tiết thừa"),
    (r"Dẫn mắt người xem", "Hướng sự chú ý của người xem"), (r"Dẫn mắt", "Hướng sự chú ý"),
    (r"(?<![(\w])Big Idea(?!\))", "ý tưởng chính (Big Idea)"),
    (r"\bAs-of\b", "Ngày chốt dữ liệu"), (r"\bas-of\b", "ngày chốt dữ liệu"),
    (r"pie, donut, treemap", "biểu đồ tròn, vành khuyên, dạng cây"),
    (r"Pie/donut/treemap", "Biểu đồ tròn, vành khuyên, dạng cây"),
    (r"(?<!→ )(?<!\()\bWaterfall\b(?! \()", "biểu đồ thác nước"), (r"(?<!→ )\bwaterfall\b", "biểu đồ thác nước"),
    (r"\b[Ll]ine chart thường", "biểu đồ đường"), (r"\bLine chart\b", "Biểu đồ đường"), (r"\bline chart\b", "biểu đồ đường"),
    (r"\bRubric\b", "Thang chấm"), (r"\brubric\b", "thang chấm"),
    (r"Câu hỏi để khai phá", "Câu hỏi trang cần trả lời"),
    (r"\bcaption\b", "chú thích"),
    (r"Δ sau quy đổi online", "Δ sau khi quy online về ngày giao"),
    (r"sau quy đổi online", "sau khi quy online về ngày giao"),
    (r"\b[Qq]uy đổi online", lambda m: ("Q" if m.group(0)[0] == "Q" else "q") + "uy online về ngày giao"),
    (r"\(Finance\)", "(Kế toán)"),
    (r"số của Sales", "số của phòng Kinh doanh"), (r"số Sales ≠ (Finance|Kế toán)", "số phòng Kinh doanh ≠ số Kế toán"),
    (r"Data Analyst bán lẻ", "Chuyên viên phân tích dữ liệu bán lẻ"),
    (r"(?<!manager · )\bE-commerce\b", "Thương mại điện tử"), (r"\be-commerce\b", "thương mại điện tử"),
    (r"(?<![\w_(])VAT(?![\w_])", "thuế GTGT"),
    (r"\bXây dashboard", "Xây dựng dashboard"),
    (r"\bJoin chỉ theo", "Ghép (merge) chỉ theo"),
    (r"\btổng chuỗi", "toàn chuỗi"), (r"\bTổng chuỗi", "Toàn chuỗi"),
    (r"\b[Mm]àu đảo", lambda m: ("Đ" if m.group(0)[0] == "M" else "đ") + "ảo màu tăng/giảm"),
    (r"\bLũy kế", "Luỹ kế"), (r"\blũy kế", "luỹ kế"),
    (r"\bchờ giao", "đang giao"),
    (r"\b[Mm]ark as date table", "Mark as Date Table"),
    (r"quan hệ inactive", "quan hệ không hoạt động (inactive)"),
    (r"(?<![(\w])[Tt]ime intelligence(?!\))", lambda m: ("Hàm" if m.group(0)[0] == "T" else "hàm") + " thời gian (time intelligence)"),
    (r"(?<![(\w])[Cc]ontext transition(?!\))", lambda m: ("Chuyển" if m.group(0)[0] == "C" else "chuyển") + " ngữ cảnh (context transition)"),
    (r"QC đã kiểm", "đã kiểm ở bước kiểm soát chất lượng dữ liệu"),
    (r"\bmọi [Pp]ivot\b", "mọi PivotTable"),
    (r"\((Australia|Canada|Germany|France|Italy|Netherlands|United Kingdom|United States)\)",
     lambda m: "(" + {"Australia": "Úc", "Canada": "Canada", "Germany": "Đức", "France": "Pháp", "Italy": "Ý",
                      "Netherlands": "Hà Lan", "United Kingdom": "Anh", "United States": "Mỹ"}[m.group(1)] + ")"),
    (r"\bMeasure \[", "Số đo ["),
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
    (r"(?<![(\w])Remove Duplicates", lambda m: _lan_dau("Remove Duplicates", "Loại bỏ Nội dung trùng lặp")),
    (r"(?<![(\w])Merge Queries", lambda m: _lan_dau("Merge Queries", "Hợp nhất truy vấn")),
    (r"(?<![(\w])Conditional Column", lambda m: _lan_dau("Conditional Column", "Cột có Điều kiện")),
    (r"(?<![(\w])Data Source Settings", lambda m: _lan_dau("Data Source Settings", "Thiết đặt Nguồn Dữ liệu")),
    (r"(?<![(\w])Change Source", lambda m: _lan_dau("Change Source", "Thay đổi nguồn")),
    (r"(?<![(\w])Applied Steps", lambda m: _lan_dau("Applied Steps", "Các bước Đã áp dụng")),
    (r"(?<![(\w])Navigator", lambda m: _lan_dau("Navigator", "Bộ dẫn hướng")),
    (r"(?<![(\w])Hide from Client Tools", lambda m: _lan_dau("Hide from Client Tools", "Ẩn khỏi Công cụ khách")),
    (r"(?<![(\w])Report Connections", lambda m: _lan_dau("Report Connections", "Kết nối Báo cáo")),
    (r"(?<![(\w])Value Filters", lambda m: _lan_dau("Value Filters", "Bộ lọc Giá trị")),
    (r"(?<![(\w])Slicer Settings", lambda m: _lan_dau("Slicer Settings", "Cài đặt Slicer")),
    (r"(?<![(\w])Set as total", lambda m: _lan_dau("Set as total", "Đặt làm tổng")),
    (r"(?<![(\w])Queries & Connections", lambda m: _lan_dau("Queries & Connections", "Truy vấn & Kết nối")),
    (r"(?<![(\w])Add this data to the Data Model", lambda m: _lan_dau("Add this data to the Data Model", "Thêm dữ liệu này vào Mô hình Dữ liệu")),
    (r"(?<![(\w])Value Field Settings", lambda m: _lan_dau("Value Field Settings", "Thiết đặt Trường Giá trị")),
    (r"(?<![(\w])Data Bars", lambda m: _lan_dau("Data Bars", "Thanh Dữ liệu")),
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
