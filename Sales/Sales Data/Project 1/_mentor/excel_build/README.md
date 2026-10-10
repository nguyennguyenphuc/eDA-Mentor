# Bộ dựng dashboard Excel — Sales Project 1

**Yêu cầu:** Excel cho Windows (Microsoft 365 hoặc Excel 2016 trở lên), bật bổ trợ **Microsoft Power Pivot for Excel**:
File → Options → Add-ins → ô Manage chọn *COM Add-ins* → Go… → tick *Microsoft Power Pivot for Excel*.
Excel cho Mac và Excel trên web không có Power Pivot.

| Thư mục / file | Nội dung |
|---|---|
| `m/01_…pq` → `m/13_…pq` | 13 truy vấn Power Query (M), đánh số theo thứ tự tạo |
| `measures.dax` | 21 số đo (measure) **bắt buộc** |
| `measures_nang_cao.dax` | 8 số đo **nâng cao** (tuỳ chọn, cho phần sáng tạo) |

1. **Power Query (M).** Data → Get Data → From Other Sources → Blank Query → Advanced Editor. Dán lần lượt các file
   trong `m/` theo số thứ tự, đặt tên truy vấn đúng tên file (bỏ số). Sửa tham số `SourceFolder` thành thư mục chứa
   `Nexa_Sales_P1_Dataset.xlsx` và `finance_revenue.xlsx` (kết thúc bằng dấu \).
2. **Nạp dữ liệu.** Home → Đóng & Tải Tới (Close & Load To…) → *Only Create Connection* và tick
   *Add this data to the Data Model*. Riêng `finance_raw`, `map_store_finance`: chỉ *Only Create Connection*.
3. **Mô hình Dữ liệu (Data Model).** Power Pivot → Manage → Dạng xem Sơ đồ (Diagram View): tạo mối quan hệ theo
   hướng dẫn mục 3.1; Đánh dấu là Bảng ngày tháng (Mark as Date Table) cho dim_date.
4. **Số đo DAX.** Tạo 21 số đo trong `measures.dax` ở bảng `_Measures` (Power Pivot → Measures → New Measure).
5. **Dashboard 4 sheet.** PivotTable/PivotChart từ Mô hình Dữ liệu, Slicer Năm/Tháng/Kênh/Quốc gia nối mọi pivot
   (Report Connections); thẻ KPI bằng CUBEVALUE; % thay đổi và câu "Điểm chính" bằng công thức Excel.

Số dòng phải khớp sau khi nạp: fact_sales_line 62.884 · fact_return 2.430 ·
fact_finance 19.400. Khi dữ liệu đổi: Data → Làm mới Tất cả (Refresh All).
