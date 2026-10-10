# Sales Project 1 — thư mục dành cho mentor

Không gửi thư mục này cho học viên: `Nexa_Sales_P1_Instructor_Key.xlsx` là đáp án.

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
python generate_dataset.py   # dataset, finance_revenue.xlsx, Instructor Key, ref_numbers.json (seed 2026)
python build_dashboard.py    # ../Nexa_Sales_P1_Dashboard.html (đọc lại file giao học viên, assert khớp generator)
python build_diagrams.py     # ../diagrams/*.drawio → .pdf + png/ (cần draw.io desktop; Linux chạy qua xvfb-run)
python build_guideline.py    # ../Dashboard_Guideline (dành cho mentor).docx (khuôn: guideline Bank)
python build_slides.py       # ../Slide_Project_1.pptx (khuôn: slide Bank)
```
- Seed cố định, nên chạy lại cho đúng các con số cũ.
- Dấu vân tay dữ liệu được in ở cuối `generate_dataset.py`.

## Sửa sơ đồ
- **Bản gốc:** các sơ đồ lưu trong `../diagrams/*.drawio`.
- **Mở và sửa:** dùng draw.io desktop hoặc app.diagrams.net, sửa xong thì lưu đè.
- **Xuất lại:** chạy `python build_diagrams.py --chi-xuat`. Lệnh này chỉ xuất lại PDF/PNG, không ghi đè `.drawio` đã sửa tay.
- **Cập nhật tài liệu:** chạy lại `build_guideline.py` và `build_slides.py` để docx/pptx nhận hình mới.
- **Lưu ý:** chạy `build_diagrams.py` không kèm cờ sẽ sinh lại `.drawio` từ code. Khi đó mọi chỉnh sửa tay bị mất.

## Chênh lệch cài sẵn (tóm tắt — chi tiết trong Instructor Key)
| Mã | Loại | Ở đâu |
|---|---|---|
| K01–K25 | Đơn CANCELLED vẫn được Finance ghi nhận | 25 đơn POS 2019–2020 |
| K26 | Finance thiếu 3 ngày doanh thu | S24 (Đức), 02/2020 |
| K27 | Doanh thu ghi gồm VAT 10 % | S54 (Mỹ), cả tháng 06/2020 |
| K28 | Một dòng Finance ghi trùng | S12 (Pháp), 10/12/2020 |
| K29–K31 | Bút toán HQ-ADJ không thuộc cửa hàng | 31/12/2018, 2019, 2020 |
| hệ thống | Online ghi nhận theo ngày giao; hoàn tiền hạch toán trễ 0–3 ngày | mọi tháng |
| bẫy phía đơn | 120 dòng trùng; 29 đơn PENDING; cửa hàng 38 đổi mã Finance từ 01/07/2019 | — |
