# HANDOFF — eDA-Mentor · Sales Project 1 (Nexa Home)

Cập nhật: 10/10/2026. Đọc file này trước khi làm tiếp Sales Project 1.

## 1. Tóm tắt nhanh
- **Dự án:** bộ tài liệu Sales Project 1 – "Kết quả bán hàng và đối soát doanh thu" (Nexa Home), dựng theo khuôn `Bank/Project 1`.
- **Thư mục:** `Sales/Sales Data/Project 1/`.
- **Nhánh làm việc:** `claude/trusting-fermi-a8clg7`, PR [nguyennguyenphuc/eDA-Mentor#1](https://github.com/nguyennguyenphuc/eDA-Mentor/pull/1). Đã push, **chưa merge**.
- **Trạng thái:**
  - Đủ bộ file, **tách 2 bộ**: `hoc_vien/` (giao học viên, có zip) và bản mentor (guideline đủ Bước 6, slide 56 trang có
    4 slide đáp án + notes, dashboard tham chiếu HTML + Excel 4 sheet, Instructor Key).
  - DAX rút còn **21 số đo bắt buộc + 8 nâng cao**; thuật ngữ chuẩn hoá theo VAS 14 / TT 99/2025 và giao diện Excel tiếng Việt.
  - **Đã dựng trên Excel Windows thật (10/10/2026):** `Nexa_Sales_P1_Dashboard.xlsx` giờ là bản Power Query + Mô hình
    Dữ liệu + 29 số đo DAX + 4 slicer + PivotTable + CUBEVALUE, sinh bằng `_mentor/build_dashboard_excel_com.py`;
    M và DAX trong `hoc_vien/excel_build/` chạy đúng, không phải sửa. Bản công thức cũ chuyển thành
    `_mentor/Nexa_Sales_P1_Doi_chieu_cong_thuc.xlsx` (đối chiếu số).

## 2. Quyết định đã chốt (đừng làm lại)
| Chủ đề | Quyết định |
|---|---|
| Nguồn dữ liệu | **Global Electronics Retailer** (Maven Analytics), tải từ Drive: https://drive.google.com/drive/folders/10yzAuIYdiTNuGKJHkad4mEEM5KApj9I6. Không dùng Contoso (≈2,5 GB, cột không khớp mô tả, không có mã đơn POS để tính AOV) |
| Mức hỗ trợ | Project đầu tiên của nhóm Sales, mentor hỗ trợ ~80 % |
| Mentor trên bìa slide | Nguyễn Phúc Nguyên, Nguyễn Thảo Nhi, Đinh Thị Thu Hà |
| Sơ đồ | Làm bằng **draw.io** (`.drawio` sửa được), rồi mới xuất PDF/PNG |
| Trình bày | Theo *Storytelling with Data*:<br>• câu "Điểm chính" mở mỗi trang<br>• tiêu đề hành động<br>• xám + một màu nhấn #2A78D6<br>• đỏ chỉ dùng cho ⚠ vượt ngưỡng |
| Công cụ dashboard | **Excel cho Windows** (M365 / 2016+, bật Power Pivot): Power Query (M) → Data Model → DAX → 4 sheet (PivotTable, PivotChart, Slicer, CUBEVALUE).<br>Thay hẳn Power BI; chỉ giữ **một dòng** "M/DAX dùng lại được nếu chuyển sang Power BI".<br>Sales P1 **không có file Power BI** (.pbix/.pbip) — chưa từng dựng; người dùng xác nhận 10/10/2026 giữ Excel. DAX vẫn dùng trong Power Pivot (21 bắt buộc + 8 nâng cao); chỉ phần trình bày (%, câu Điểm chính, tô màu) là công thức Excel. |
| Bố cục dashboard | 4 sheet: Tong_quan → Doanh_thu → Doi_soat → Hoan_tra, cộng sheet `_pivot` (ẩn).<br>Slicer **Năm + Tháng**, Kênh, Quốc gia, dùng chung cho cả 4 sheet |
| Quy tắc dữ liệu | • Chỉ đơn COMPLETED; bỏ 120 dòng trùng; Returns theo return_date<br>• Online: Finance ghi nhận theo delivery_date<br>• Δ = Finance − dữ liệu đơn; vượt ngưỡng khi \|Δ\| > 200 USD **và** \|Δ%\| > 1 % (tháng × cửa hàng)<br>• Cửa hàng 38 đổi mã GBR-03 → GBR-03N từ 01/07/2019; dòng HQ-ADJ có store_key trống |
| Đặt tên | Measure `\|Δ\|` đổi thành `Δ tuyệt đối`, vì ký tự `\|` có thể không hợp lệ trong Excel |
| Số đo DAX | 21 bắt buộc (`measures.dax`) + 8 nâng cao (`measures_nang_cao.dax`). % thay đổi, câu "Điểm chính", tiêu đề hành động, tô màu nhấn làm bằng **công thức Excel** (FIXED, &, IF) trên ô CUBEVALUE — không viết measure chuỗi |
| Thuật ngữ | `_mentor/thuat_ngu.py` (một nguồn): Finance → phòng Tài chính – Kế toán (gọi tắt Kế toán); Sales Manager → Trưởng phòng Kinh doanh; measure → số đo; Gross Sales/Discounts/Returns/Net Revenue = doanh thu bán hàng / chiết khấu thương mại / hàng bán bị trả lại / doanh thu thuần (VAS 14); tên menu Excel = tiếng Việt của Microsoft + tiếng Anh trong ngoặc. Tên file, cột, số đo giữ nguyên. "Đối soát" giữ làm từ chính |
| Học viên xem gì | `hoc_vien/`: đề bài, dataset, finance_revenue, Huong_dan_hoc_vien.docx, Slide_Project_1_hoc_vien.pptx, excel_build/, so_do/, template báo cáo. Dashboard tham chiếu gửi sau buổi 1 (tuỳ mentor) |
| Số trong câu chữ Excel | Dùng `FIXED()` thay `TEXT(x;"#,##0.0")` — FIXED theo dấu thập phân của máy, không vỡ khi máy cài định dạng Việt Nam |
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
| (commit này) | Sửa lỗi Excel đòi "sửa chữa" file; rút DAX 21+8; dashboard Excel tham chiếu 4 sheet khớp HTML + công cụ kiểm (kiem_xlsx.py, xuat_pdf_kiem_tra.py); chuẩn hoá thuật ngữ; tách bộ học viên (hoc_vien/, dong_goi_hoc_vien.py) |

## 4. Bản đồ file (trong `Sales/Sales Data/Project 1/`)
**`hoc_vien/` — giao học viên** (gửi thư mục hoặc `Nexa_Sales_P1_hoc_vien.zip`, zip bị gitignore):
- `README.docx` (đề bài), `Nexa_Sales_P1_Dataset.xlsx`, `finance_revenue.xlsx`, `guild_report_template.pdf`.
- `Huong_dan_hoc_vien.docx`: bản học viên của guideline (Bước 6 chỉ có phương pháp + khung cầu nối trống).
- `Slide_Project_1_hoc_vien.pptx`: 52 slide, đã **xoá hẳn** 4 slide đáp án và mọi notes.
- `excel_build/` (13 file M, measures.dax, measures_nang_cao.dax, README), `so_do/*.pdf`, `DANH_SACH_FILE.md`.

**Gốc `Project 1/` — mentor:**
- `Dashboard_Guideline (dành cho mentor).docx` (có Bước 6 đáp án + Phụ lục B script), `Slide_Project_1.pptx` (56 slide, notes).
- `Nexa_Sales_P1_Dashboard.html` và `Nexa_Sales_P1_Dashboard.xlsx` — bản tham chiếu, có thể gửi học viên sau buổi 1.
  File .xlsx dựng đúng các bước guideline: 13 query Power Query (SourceFolder trỏ `hoc_vien/`), Mô hình Dữ liệu 10 bảng /
  12 quan hệ, 29 số đo DAX, 4 slicer dùng chung (Năm, Tháng, Kênh, Quốc gia; mặc định 12/2020), 8 PivotTable trên `_pivot`
  (ẩn), thẻ KPI/câu Điểm chính bằng CUBEVALUE + công thức Excel; vùng ô: slicer B4:Q6, Điểm chính B7:Q9, KPI B11:Q14.
- `diagrams/*.drawio` (+ pdf, png/): 01 pipeline, 02 mô hình, 03 đối soát, 04 cây KPI, 05 Power Query.

**`_mentor/` — chỉ mentor:** Instructor Key (31 chênh lệch), `Nexa_Sales_P1_Doi_chieu_cong_thuc.xlsx` (bản công thức thường
để so số), `ref_numbers.json`, `pq_steps.json`, script build, `thuat_ngu.py`,
`kiem_xlsx.py`, `xuat_pdf_kiem_tra.py`, `dong_goi_hoc_vien.py`, `kiem_tra/` (PDF + ảnh so sánh Excel ↔ HTML), README_mentor.md.

**Không có trên git:** `source/Global_Electronics_Retailer.zip`, `Nexa_Sales_P1_hoc_vien.zip` (`*.zip` bị gitignore).

## 5. Cách chạy lại (trong `_mentor/`)
```
python generate_dataset.py        # seed 2026, in "fingerprint" cuối → hoc_vien/ dataset + finance
python build_dashboard.py         # HTML
python build_dashboard_xlsx.py    # bản đối chiếu số _mentor/Nexa_Sales_P1_Doi_chieu_cong_thuc.xlsx (công thức thường)
python build_dashboard_excel_com.py  # CHẠY TRÊN WINDOWS + Excel: ../Nexa_Sales_P1_Dashboard.xlsx (PQ + Data Model + DAX),
                                     # tự kiểm số với ref_numbers.json, xuất kiem_tra/excel_*.pdf (~2 phút)
python build_excel_kit.py         # hoc_vien/excel_build/ + pq_steps.json (assert số dòng, ngưỡng, 21+8 số đo)
cd /opt/node-tools && node <repo>/Sales/Sales\ Data/Project\ 1/_mentor/chup_dashboard.js <assets> <html>   # 4 ảnh tab HTML
python build_diagrams.py          # sinh lại .drawio + xuất; dùng --chi-xuat nếu đã sửa .drawio bằng tay
python build_guideline.py         # 2 bản: mentor + hoc_vien/
python build_slides.py            # 2 bản; assert DAX trên slide = file .dax, notes khớp số slide
python dong_goi_hoc_vien.py       # so_do/, DANH_SACH_FILE.md, zip; assert không lộ đáp án
python kiem_xlsx.py               # lỗi làm Excel đòi sửa chữa file
python xuat_pdf_kiem_tra.py       # LibreOffice tính lại bản đối chiếu số, so số 7 tổ hợp lọc, xuất PDF + ảnh so với HTML
```
**Môi trường cần có:**
- Python: pandas, numpy, openpyxl, xlsxwriter, python-docx, python-pptx, Pillow; python3-uno (LibreOffice).
- draw.io desktop, chạy qua `xvfb-run -a drawio --no-sandbox --export`; LibreOffice; poppler (pdftoppm).
- Node + Playwright (Chromium có sẵn trong container cloud).

**Lưu ý:**
- Script build giữ cố định khuôn Bank (`Bank/Project 1`) cho docx/pptx.
- Mọi số trong tài liệu đọc từ dữ liệu, không gõ tay. Thuật ngữ áp tự động qua `thuat_ngu.chuan_hoa()`.

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
1. ~~Code M và DAX chưa chạy trên Excel thật~~ — **xong 10/10/2026** trên Excel 365: 13 query nạp đúng số dòng, 12 quan hệ,
   29 số đo (tên có Δ chạy được), CUBERANKEDMEMBER/CUBESETCOUNT với tên slicer chạy được; số khớp Bước 6 (12/2020, LM,
   SPLY), 06/2020 Mỹ → S45, S54; 02/2020 → S24 −22.168,91; online 12/2020 Δ sau quy đổi ≈ 0. Mark as Date Table không
   tạo được qua COM nhưng DATEADD/SAMEPERIODLASTYEAR vẫn đúng (quan hệ trên cột kiểu Date).
   Kinh nghiệm tự động hoá (ghi trong `excel_com_model.py`): Power Query đọc .xlsx gọi ngược Excel → nạp đồng bộ bị treo;
   phải tạo connection với query rỗng, rồi gán M thật và RefreshAll chạy nền. Add2 dùng CommandText `"tên"` kiểu 6 để bảng
   mô hình mang đúng tên query. Waterfall: dữ liệu đặt nằm ngang. Slicer phải tạo trước khi trang có biểu đồ.
2. **Lỗi "We found a problem with some content" đã sửa tận gốc** (chuỗi "=" thành công thức hỏng; data validation có "="
   thừa; ký tự "−" trong định dạng số; `<definedNames/>` rỗng) và có `kiem_xlsx.py` chặn lại — nhưng **chưa mở bằng Excel
   thật để xác nhận**; nhờ mentor mở 4 file .xlsx trên Excel Windows một lần.
3. Dashboard Excel tham chiếu được kiểm bằng LibreOffice (số khớp HTML ở 7 tổ hợp lọc, không ô lỗi); Excel thật có thể
   hiển thị biểu đồ hơi khác (font, nhãn). Waterfall dựng bằng cột chồng 4 chuỗi (Excel 2016 có loại Waterfall riêng nhưng
   openpyxl không tạo được).
4. ~~Chưa có workbook mẫu dựng bằng Mô hình Dữ liệu~~ — xong (`Nexa_Sales_P1_Dashboard.xlsx`). Còn: slicer hiện mục
   "(blank)" (dòng HQ-ADJ không có cửa hàng, ngày ngoài dim_date) — hành vi chuẩn của Power Pivot; số trong câu chữ theo
   dấu thập phân của máy (máy dựng dùng dấu chấm). Tài liệu/slide cần cập nhật theo file mới: phiên song song đang làm
   (`build_guideline.py`, `build_slides.py`, `slide_notes.py`).
5. Ảnh dashboard ở slide Phần 7 vẫn chụp từ bản HTML → thay bằng ảnh 4 sheet Excel thật sau khi dựng.
6. Slide 25–26 (code M) chữ 9–10,5 pt, hơi nhỏ khi chiếu.
7. `return_reason` còn mã tiếng Anh trong Pivot (CHANGE_MIND…) — có thể thêm cột Việt hoá trong M fact_return.
8. Thuật ngữ vai trò/KPI bán lẻ không có văn bản chính thức (lấy theo thực tế tuyển dụng); chưa đối chiếu được bản in
   giáo trình NEU/UEH/HVTC. `Dataset_for_project1n4.docx` (mô tả Contoso, nhắc Power BI) để nguyên.
9. Skill: `pbi-dashboard-loop` không dùng cho P1; `goi-ngoai`/ChatGPT web cần Chrome cổng 9222 trên máy người dùng.
10. PR #1 chưa review/merge. Sales Project 2–4 chưa làm.

## 8. Gợi ý bước tiếp (theo thứ tự)
1. ~~Dựng thử Power Query + DAX trên Excel Windows~~ — xong. Mentor mở `Nexa_Sales_P1_Dashboard.xlsx`, sửa SourceFolder
   theo máy mình rồi Refresh All để thử.
2. Chụp 4 sheet Excel thật, thay ảnh ở slide Phần 7, chạy lại `build_slides.py`.
3. Review và merge PR #1. Commit nối tiếp làm trên nhánh mới tách từ nhánh mặc định.
4. Bắt đầu Sales Project tiếp theo, dùng lại khung script của P1 (`generate_dataset` → `build_*`).
