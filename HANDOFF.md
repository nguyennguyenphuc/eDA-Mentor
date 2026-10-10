# HANDOFF — eDA-Mentor · Sales Project 1 (Nexa Home)

Cập nhật: 10/10/2026. Đọc file này trước khi làm tiếp Sales Project 1.

## 1. Tóm tắt nhanh
- **Dự án:** bộ tài liệu Sales Project 1 – "Kết quả bán hàng và đối soát doanh thu" (Nexa Home), dựng theo khuôn `Bank/Project 1`.
- **Thư mục:** `Sales/Sales Data/Project 1/`.
- **Nhánh làm việc:** `claude/trusting-fermi-a8clg7`, PR [nguyennguyenphuc/eDA-Mentor#1](https://github.com/nguyennguyenphuc/eDA-Mentor/pull/1). Đã push, **chưa merge**.
- **Trạng thái:**
  - Đủ bộ file giao: dataset, file Finance, guideline, slide 55 trang có notes, dashboard HTML + Excel, bộ code M/DAX, sơ đồ draw.io.
  - Việc lớn còn lại: **dựng thử trên Excel Windows thật** (mục 7).

## 2. Quyết định đã chốt (đừng làm lại)
| Chủ đề | Quyết định |
|---|---|
| Nguồn dữ liệu | **Global Electronics Retailer** (Maven Analytics), tải từ Drive: https://drive.google.com/drive/folders/10yzAuIYdiTNuGKJHkad4mEEM5KApj9I6. Không dùng Contoso (≈2,5 GB, cột không khớp mô tả, không có mã đơn POS để tính AOV) |
| Mức hỗ trợ | Project đầu tiên của nhóm Sales, mentor hỗ trợ ~80 % |
| Mentor trên bìa slide | Nguyễn Phúc Nguyên, Nguyễn Thảo Nhi, Đinh Thị Thu Hà |
| Sơ đồ | Làm bằng **draw.io** (`.drawio` sửa được), rồi mới xuất PDF/PNG |
| Trình bày | Theo *Storytelling with Data*:<br>• câu "Điểm chính" mở mỗi trang<br>• tiêu đề hành động<br>• xám + một màu nhấn #2A78D6<br>• đỏ chỉ dùng cho ⚠ vượt ngưỡng |
| Công cụ dashboard | **Excel cho Windows** (M365 / 2016+, bật Power Pivot): Power Query (M) → Data Model → DAX → 4 sheet (PivotTable, PivotChart, Slicer, CUBEVALUE).<br>Thay hẳn Power BI; chỉ giữ **một dòng** "M/DAX dùng lại được nếu chuyển sang Power BI" |
| Bố cục dashboard | 4 sheet: Tong_quan → Doanh_thu → Doi_soat → Hoan_tra, cộng sheet `_pivot` (ẩn).<br>Slicer **Năm + Tháng**, Kênh, Quốc gia, dùng chung cho cả 4 sheet |
| Quy tắc dữ liệu | • Chỉ đơn COMPLETED; bỏ 120 dòng trùng; Returns theo return_date<br>• Online: Finance ghi nhận theo delivery_date<br>• Δ = Finance − dữ liệu đơn; vượt ngưỡng khi \|Δ\| > 200 USD **và** \|Δ%\| > 1 % (tháng × cửa hàng)<br>• Cửa hàng 38 đổi mã GBR-03 → GBR-03N từ 01/07/2019; dòng HQ-ADJ có store_key trống |
| Đặt tên | Measure `\|Δ\|` đổi thành `Δ tuyệt đối`, vì ký tự `\|` có thể không hợp lệ trong Excel |
| Bảo mật | Không nhận cookie/session ChatGPT của người dùng (đã từ chối) |

## 3. Đã làm (theo commit)
| Commit | Nội dung |
|---|---|
| 80c1bbc | Dữ liệu các project và skill `pbi-dashboard-loop`, `goi-ngoai` |
| 12cc000 | Dataset, file Finance, dashboard HTML, sơ đồ draw.io, guideline mentor |
| 4053cdd, a914913 | Ảnh chụp dashboard cho slide; bộ slide 47 trang; sửa cách đếm đơn PENDING |
| abba1c6 | Tên mentor trên bìa |
| 6fc8f58 | Speaker notes cho mọi slide |
| 689cd91 | Rà soát theo Storytelling with Data:<br>• dashboard HTML làm lại<br>• thêm `Nexa_Sales_P1_Dashboard.xlsx` (SUMIFS)<br>• chuẩn hoá dataset (Excel Table `tbl_*`, Data_Dictionary, fx_rate) |
| d1beea3 | Chuyển sang Excel:<br>• `excel_build/` gồm 13 query M và `measures.dax` 53 measure<br>• sơ đồ 05 pipeline Power Query; sửa sơ đồ 01, 02<br>• guideline viết lại cho Excel<br>• slide 55 trang + notes |

## 4. Bản đồ file
**Giao học viên** (`Sales/Sales Data/Project 1/`):
- `Nexa_Sales_P1_Dataset.xlsx`: 7 dim, 2 fact, map, KPI, giao dịch mẫu; mỗi sheet là Table `tbl_<sheet>`.
- `finance_revenue.xlsx`: số Finance theo ngày × mã cửa hàng Finance.
- `Dashboard_Guideline (dành cho mentor).docx`: Bước 6 là đáp án, phải ẩn khi giao đề.
- `Slide_Project_1.pptx`: 55 slide, 4 slide phụ lục mentor đặt ẩn.
- `excel_build/`: `m/*.pq`, `measures.dax`, `README.md` (các bước dựng trong Excel).
- `Nexa_Sales_P1_Dashboard.html`: bản tham chiếu bố cục và số.
- `Nexa_Sales_P1_Dashboard.xlsx`: file đối chiếu số; **không** phải bản dựng bằng Data Model.
- `diagrams/*.drawio` (+ `.pdf`, `png/`): 01 pipeline, 02 mô hình, 03 đối soát, 04 cây KPI, 05 Power Query.
- `guild_report_template.pdf`, `README.docx`.

**Chỉ mentor** (`_mentor/`):
- Script `generate_dataset.py`, `aggregate.py`, `build_*.py`, `chup_dashboard.js`, `slide_notes.py`.
- `Nexa_Sales_P1_Instructor_Key.xlsx`: đáp án 31 chênh lệch cài sẵn.
- `ref_numbers.json`, `pq_steps.json`, `README_mentor.md`.

**Không có trên git:** `source/Global_Electronics_Retailer.zip` (`*.zip` bị gitignore), tải lại từ Drive.

## 5. Cách chạy lại (trong `_mentor/`)
```
python generate_dataset.py        # seed 2026, in "fingerprint" cuối
python build_dashboard.py         # HTML
python build_dashboard_xlsx.py    # file đối chiếu SUMIFS
python build_excel_kit.py         # excel_build/ + pq_steps.json (assert số dòng, ngưỡng)
cd /opt/node-tools && node <repo>/Sales/Sales\ Data/Project\ 1/_mentor/chup_dashboard.js <assets> <html>   # ảnh cho slide
python build_diagrams.py          # sinh lại .drawio + xuất; dùng --chi-xuat nếu đã sửa .drawio bằng tay
python build_guideline.py
python build_slides.py            # assert DAX trên slide = measures.dax, notes khớp số slide
```
**Môi trường cần có:**
- Python: pandas, numpy, openpyxl, xlsxwriter, python-docx, python-pptx.
- draw.io desktop, chạy qua `xvfb-run -a drawio --no-sandbox --export`.
- LibreOffice: chỉ để render kiểm tra; khi tính lại công thức thì làm trên bản sao.
- Node + Playwright (Chromium có sẵn trong container cloud).

**Lưu ý:**
- Script build giữ cố định khuôn Bank (`Bank/Project 1`) cho docx/pptx.
- Mọi số trong tài liệu đọc từ dữ liệu, không gõ tay.

## 6. Số mốc để kiểm nhanh
- **Tháng báo cáo 12/2020:**
  - Net Revenue **606.125,07** USD
  - Orders **310**
  - Δ Finance (không gồm HQ-ADJ) **−12.536,23** USD
- **Số dòng:**
  - fact_sales_line 63.004 → **62.884** sau khi bỏ trùng
  - fact_return **2.430**
  - fact_finance **19.400** (3 dòng HQ-ADJ trống store_key)
- **Ô vượt ngưỡng:** **32** ô năm 2020, **120** ô toàn kỳ.
- **Dataset:** v1.0, seed 2026, fingerprint `983ceb93b510cb16`.

## 7. Chưa làm / giới hạn còn tồn tại
1. **Code M và DAX chưa chạy trên Excel thật.** Môi trường sinh tài liệu là Linux, không có Excel.
   - Đã kiểm: pandas chạy song song từng bước M, logic vượt ngưỡng, bracket.
   - Cần làm: dựng thử theo `excel_build/README.md`, so với guideline Bước 6. Nếu lỗi thì sửa trong `build_excel_kit.py` rồi chạy lại build_excel_kit → build_guideline → build_slides.
2. **Chưa có workbook mẫu dựng bằng Data Model/PivotTable thật.** openpyxl không tạo được. Nên lưu bản mentor dựng tay làm đáp án dashboard.
3. **Ảnh dashboard ở slide Phần 7 vẫn chụp từ bản HTML.** Thay bằng ảnh 4 sheet Excel thật sau khi dựng xong (hàm `result_slide` và slide "Bố cục đề xuất 4 trang" trong `build_slides.py`).
4. Slide 25 (code M fact_sales_line / fact_return) dùng chữ 9 pt, hơi nhỏ khi chiếu. Có thể tách thành 2 slide.
5. `return_reason` còn mã tiếng Anh (CHANGE_MIND, DEFECT, …) trong Pivot. Có thể thêm cột Việt hoá trong query M fact_return; nhớ cập nhật pq_steps và sơ đồ 05.
6. Waterfall phân rã Δ trong Excel chỉ có 3 thanh (lệch thời điểm online → Δ sau quy đổi → Δ). Bản HTML tách từng cửa hàng; phần tách này đang để làm gợi ý sáng tạo.
7. `Dataset_for_project1n4.docx` (mô tả Contoso, có nhắc Power BI) để nguyên vì là tài liệu gốc.
8. **Skill:**
   - `pbi-dashboard-loop` (Power BI) không dùng cho P1 sau khi đổi sang Excel.
   - `goi-ngoai` / ChatGPT web cần Chrome mở cổng 9222 trên máy người dùng, nên không chạy được trong container cloud.
   - agy (Gemini CLI) cài trong container sẽ mất khi container bị thu hồi.
9. PR #1 chưa review/merge. Sales Project 2–4 chưa làm (`Dataset_for_project1n4.docx` có ý tưởng cho Project 4).

## 8. Gợi ý bước tiếp (theo thứ tự)
1. Mentor dựng thử trên Excel Windows, sửa lỗi M/DAX nếu có, lưu workbook mẫu vào `_mentor/`.
2. Chụp 4 sheet Excel thật, thay ảnh ở slide Phần 7, chạy lại `build_slides.py`.
3. Review và merge PR #1. Commit nối tiếp làm trên nhánh mới tách từ nhánh mặc định.
4. Bắt đầu Sales Project tiếp theo, dùng lại khung script của P1 (`generate_dataset` → `build_*`).
