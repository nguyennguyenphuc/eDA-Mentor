# Bộ dựng dashboard Excel — Sales Project 1

Yêu cầu: Excel cho Windows (Microsoft 365 hoặc 2016+), bật add-in **Microsoft Power Pivot for Excel**
(File → Options → Add-ins → Manage: COM Add-ins → Go… → tick Power Pivot).

1. **Power Query (M)** — Data → Get Data → From Other Sources → Blank Query → Advanced Editor, dán lần lượt các file
   trong `m/` theo số thứ tự, đặt tên query đúng tên file (bỏ số). Sửa `SourceFolder` thành thư mục chứa
   `Nexa_Sales_P1_Dataset.xlsx` và `finance_revenue.xlsx`.
2. **Nạp** — Home → Close & Load To… → *Only Create Connection* + tick *Add this data to the Data Model*
   (riêng `finance_raw`, `map_store_finance`: chỉ *Only Create Connection*).
3. **Data Model** — Power Pivot → Manage → Diagram View: tạo quan hệ theo guideline mục 3.1, Mark as Date Table cho dim_date.
4. **DAX** — dán các measure trong `measures.dax` vào bảng `_Measures` (Power Pivot → Measures → New Measure).
5. **Dashboard** — 4 sheet: PivotTable/PivotChart từ Data Model, Slicer Tháng/Kênh/Quốc gia nối mọi pivot
   (Report Connections), thẻ KPI và câu Điểm chính bằng CUBEVALUE.

Số dòng phải khớp sau khi nạp: fact_sales_line 62.884 · fact_return 2.430 ·
fact_finance 19.400. Code M và DAX ở đây dùng lại được nguyên vẹn nếu sau này chuyển sang Power BI.
