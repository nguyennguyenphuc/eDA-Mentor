# Sales Project 1 — thư mục dành cho mentor

> Tình trạng công việc, việc còn dở và bước tiếp theo: xem `HANDOFF.md` ở gốc repo.

Không gửi thư mục này cho học viên: `Nexa_Sales_P1_Instructor_Key.xlsx` là đáp án.

## Học viên được xem file nào
| Ở đâu | File | Ai xem |
|---|---|---|
| `../hoc_vien/` | README.docx (đề bài), Nexa_Sales_P1_Dataset.xlsx, finance_revenue.xlsx, Huong_dan_hoc_vien.docx, Slide_Project_1_hoc_vien.pptx, excel_build/, so_do/*.pdf, guild_report_template.pdf | **Học viên** – gửi nguyên thư mục hoặc `../Nexa_Sales_P1_hoc_vien.zip` |
| `../` | Nexa_Sales_P1_Dashboard.html, Nexa_Sales_P1_Dashboard.xlsx | Mentor; có thể gửi học viên sau buổi 1 |
| `../` | Dashboard_Guideline (dành cho mentor).docx (có Bước 6 đáp án), Slide_Project_1.pptx (có 4 slide đáp án ẩn + lời thuyết trình), diagrams/*.drawio | Chỉ mentor |
| `_mentor/` | Instructor Key, ref_numbers.json, script, kiem_tra/ | Chỉ mentor |

Bản học viên của hướng dẫn và slide **được sinh tự động**: hướng dẫn bỏ Bước 6 đáp án và Phụ lục script; slide xoá hẳn
4 slide đáp án (không chỉ ẩn) và xoá lời thuyết trình. `dong_goi_hoc_vien.py` kiểm lại không lộ đáp án rồi mới đóng zip.

## Nguồn dữ liệu
- Bộ dữ liệu nền là **Global Electronics Retailer** (Maven Analytics).
- File `Global+Electronics+Retailer.zip` nằm trong thư mục Drive của project:
  https://drive.google.com/drive/folders/10yzAuIYdiTNuGKJHkad4mEEM5KApj9I6
- Tải file về và lưu thành `../source/Global_Electronics_Retailer.zip`.
- Repo bỏ qua `*.zip` (xem `.gitignore`), nên file này không có trên git.
- Vì sao không dùng `Contoso.zip`:
  - Đây là bản ContosoRetailDW, giải nén khoảng 2,5 GB.
  - Các cột không khớp với mô tả trong `Dataset_for_project1n4.docx`.
  - FactSales không có mã đơn, nên không tính được AOV cho đơn tại cửa hàng.

## Chạy lại toàn bộ (Python 3 + pandas, numpy, openpyxl, xlsxwriter, python-docx, python-pptx)
```
python generate_dataset.py      # hoc_vien/ dataset + finance_revenue.xlsx, Instructor Key, ref_numbers.json (seed 2026)
python build_dashboard.py       # ../Nexa_Sales_P1_Dashboard.html (số từ aggregate.py)
python build_dashboard_xlsx.py  # ../Nexa_Sales_P1_Dashboard.xlsx — 4 sheet giống HTML, công thức thường (đối chiếu số)
python build_excel_kit.py       # hoc_vien/excel_build/ (code M, 21 số đo bắt buộc + 8 nâng cao) + pq_steps.json; assert số dòng
cd /opt/node-tools && node <repo>/Sales/Sales\ Data/Project\ 1/_mentor/chup_dashboard.js <thư mục assets> <file html>   # ảnh cho slide
python build_diagrams.py        # ../diagrams/*.drawio → .pdf + png/ (draw.io; Linux chạy qua xvfb-run)
python build_guideline.py       # hướng dẫn: bản mentor (../) + bản học viên (hoc_vien/)
python build_slides.py          # slide: bản mentor (../) + bản học viên (hoc_vien/); assert DAX trên slide = file .dax
python dong_goi_hoc_vien.py     # chép sơ đồ PDF vào hoc_vien/so_do, kiểm không lộ đáp án, tạo ../Nexa_Sales_P1_hoc_vien.zip
python kiem_xlsx.py             # kiểm mọi .xlsx không có lỗi làm Excel đòi "sửa chữa" file
python xuat_pdf_kiem_tra.py     # LibreOffice tính lại dashboard Excel, so số với pandas, xuất kiem_tra/*.pdf + ảnh so với HTML
```
- Seed cố định, nên chạy lại cho đúng các con số cũ.
- Dấu vân tay dữ liệu được in ở cuối `generate_dataset.py`.
- Thuật ngữ (Finance → phòng Tài chính – Kế toán, measure → số đo, tên menu Excel tiếng Việt…) nằm ở `thuat_ngu.py`,
  áp tự động khi sinh hướng dẫn, slide, notes, sơ đồ và README dataset; tên file/cột/số đo giữ nguyên.

## Bộ dựng Excel (excel_build/) — giới hạn cần biết
- Học viên dựng dashboard bằng **Excel for Windows** (Microsoft 365 / 2016+, bật bổ trợ Power Pivot): Power Query (M) →
  Mô hình Dữ liệu → 21 số đo DAX bắt buộc (+8 nâng cao) → 4 sheet (PivotTable, PivotChart, Slicer, CUBEVALUE); % thay đổi,
  câu Điểm chính và tô màu nhấn làm bằng công thức Excel. Guideline Bước 2–5 và slide Phần 4, 6, 7.
- Code M và DAX **chưa được chạy thử trên Excel thật** (máy sinh tài liệu là Linux, không có Excel). Đã kiểm:
  từng bước M được chạy song song bằng pandas và khớp số dòng QC (62.884 · 2.430 · 19.400); logic [Vượt ngưỡng]
  khớp Instructor Key (32 ô năm 2020, 120 ô toàn kỳ); Net Revenue 12/2020 khớp số tham chiếu; ngoặc cân.
- **Trước buổi học, mentor dựng thử một lần trên Excel Windows** theo excel_build/README.md và so số với Bước 6.
  Lỗi cú pháp nếu có thì sửa trong build_excel_kit.py rồi chạy lại build_excel_kit → build_guideline → build_slides.

## Kiểm file Excel (không có Excel trên máy sinh tài liệu)
- `kiem_xlsx.py` đọc XML của mọi .xlsx, báo các lỗi làm Excel hiện "We found a problem with some content": chuỗi bắt
  đầu bằng "=" bị ghi thành công thức hỏng, data validation có dấu "=" thừa, định dạng số có ký tự "−" không đặt trong
  ngoặc kép, `<definedNames/>` rỗng, tiêu đề Table lệch. Bản trước có đủ ba lỗi đầu — đã sửa tận gốc ở script sinh file.
- `xuat_pdf_kiem_tra.py` mở **bản sao** dashboard Excel bằng LibreOffice, đổi ô chọn qua 7 tổ hợp tháng/kênh/quốc gia,
  so Net Revenue, đơn, số Kế toán, Δ, Δ sau quy đổi, số ô vượt ngưỡng với pandas; xuất `kiem_tra/dashboard_xlsx.pdf`
  và `kiem_tra/so_sanh_s1..4.png` (trang Excel đặt cạnh tab HTML). PDF do LibreOffice vẽ nên font/biểu đồ không y hệt Excel.
- Ô vàng ở sheet Tong_quan là ô chọn Tháng / Kênh / Quốc gia; 4 sheet dùng chung.

## Sửa sơ đồ
- **Bản gốc:** các sơ đồ lưu trong `../diagrams/*.drawio`.
- **Mở và sửa:** dùng draw.io desktop hoặc app.diagrams.net, sửa xong thì lưu đè.
- **Xuất lại:** chạy `python build_diagrams.py --chi-xuat`. Lệnh này chỉ xuất lại PDF/PNG, không ghi đè `.drawio` đã sửa tay.
- **Cập nhật tài liệu:** chạy lại `build_guideline.py` và `build_slides.py` để docx/pptx nhận hình mới.
- **Lưu ý:** chạy `build_diagrams.py` không kèm cờ sẽ sinh lại `.drawio` từ code. Khi đó mọi chỉnh sửa tay bị mất.

## Chênh lệch cài sẵn (tóm tắt — chi tiết trong Instructor Key)
| Mã | Loại | Ở đâu |
|---|---|---|
| K01–K25 | Đơn CANCELLED vẫn được Kế toán ghi nhận | 25 đơn POS 2019–2020 |
| K26 | Kế toán thiếu 3 ngày doanh thu | S24 (Đức), 02/2020 |
| K27 | Doanh thu ghi gồm VAT 10 % | S54 (Mỹ), cả tháng 06/2020 |
| K28 | Một dòng Kế toán ghi trùng | S12 (Pháp), 10/12/2020 |
| K29–K31 | Bút toán HQ-ADJ không thuộc cửa hàng | 31/12/2018, 2019, 2020 |
| hệ thống | Online ghi nhận theo ngày giao; hoàn tiền ghi sổ trễ 0–3 ngày | mọi tháng |
| bẫy phía đơn | 120 dòng trùng; 29 đơn PENDING; cửa hàng 38 đổi mã Kế toán từ 01/07/2019 | — |
