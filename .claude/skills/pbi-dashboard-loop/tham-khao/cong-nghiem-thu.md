# Thiết kế cổng nghiệm thu

Cổng là thứ dừng vòng lặp. Một cổng tốt trả lời **một** câu hỏi, chạy rẻ trước đắt sau, và **có thể đỏ**.

## Thứ tự trong một vòng

| Cổng | Trả lời câu hỏi gì | Cần Desktop? | Script |
|---|---|---|---|
| Dấu vân tay nguồn | dữ liệu trước mốc cắt có bị ai sửa không | không | dự án (`kiem_nen.py` + `cutoff.json`: số dòng + md5 từng dòng) |
| Lint DAX tĩnh | measure có sai cú pháp không | không | `scripts/dax_lint.py` gọi trong `build_model.py` |
| Hợp đồng model ↔ report | model có đủ mọi measure/cột report bind vào không | không | dự án (tự kiểm cuối `build_model.py`) |
| `$schema` | PBIR có hợp lệ theo lược đồ Microsoft không | không | `scripts/kiem_schema.py` |
| Hộp thoại chặn | Desktop có mở được dự án không | có | `scripts/pbi_health.ps1` |
| Nạp đủ dòng | Power BI có nạp đúng số dòng nguồn trả về không | có | dự án (đếm dòng sau refresh, so với SQL) |
| Measure hỏng | measure nào lỗi và vì sao | có | `scripts/pbi_measures_check.py` |
| Số liệu | DAX có ra đúng số tính độc lập không | có | `mau/kiem_so_mau.py` + `moc.json` |
| Dải thông báo | PDF có dính dải cảnh báo không | có | `scripts/pbi_ui.ps1` (loop gọi sẵn) |
| Điểm ảnh | hình có khác **đúng chỗ dự đoán** không | không (cần 2 bản render) | `scripts/so_diem_anh.py` |
| Best Practice Analyzer | model có vi phạm BPA nào mới không | không | `scripts/bpa_check.py` |
| Thân DAX | hai bản model có cùng định nghĩa measure không | không | `scripts/diff_measures.py` |

## Số kỳ vọng phải độc lập

`moc.json` tính bằng đường **khác** Power BI: SQL thuần trên server, hoặc pandas trên file nguồn. Tách hai tầng:
script kéo số (`keo.py`) và script dựng mốc (`moc.py`) tự kiểm với số công bố của đề — để khi lệch biết ngay
lệch ở M/DAX hay ở cách hiểu định nghĩa.

Các lớp nên có (từ cổng 606 phép của Finance):

| Lớp | Bắt lỗi gì |
|---|---|
| `0-tukiem` | chính mốc có tự mâu thuẫn không (tổng lát cắt = tổng) |
| `A-kpi` | chỉ tiêu thẻ × các kỳ |
| `B-latcat` | từng phân khúc / nhóm × các kỳ |
| `C-chuoi` | toàn bộ chuỗi thời gian của biểu đồ đường |
| `D-loc` | hành vi khi lọc: đại lượng nào phải BLANK, trang nào không được đổi |
| `E-bieudo` | measure **riêng của biểu đồ** (biên độ, phễu, cơ cấu theo mốc…) — lớp hay bị quên nhất |

Dung sai: đếm khớp tuyệt đối; tiền ≤ 0,005; tỉ lệ ≤ 1e-6. BLANK khác 0 trừ khi định nghĩa nói khác.

## Cổng phải có thể đỏ

Một vòng rà soát đối kháng (Finance, 02/09/2026) tìm ra **năm cổng không thể đỏ**: cổng đếm lệch rồi `return 0`
vô điều kiện; cổng coi tập kết quả rỗng là đạt; chế độ `--offline` in phán quyết về model mà không chạm model;
tự kiểm trả 0 khi không thấy thư mục cần kiểm. Luật:

- Không tìm thấy đối tượng kiểm → **đỏ**, không phải xanh.
- Phép không chạy được → tính là **lệch**, không bỏ qua. Lọc "nếu tên không nằm trong danh sách thì bỏ qua"
  làm cổng âm thầm co từ 31 xuống 26 phép.
- Mỗi cổng mới kèm cờ `--tiem`: tự phá hỏng một bản chép (đổi tên một measure, bỏ một khối…) rồi đòi chính mình
  phải đỏ. Cổng chưa từng đỏ là cổng chưa được chứng minh.

## Những gì cổng số không bắt được

Cổng số gọi measure theo **tên**, không đi qua visual. Ba loại lỗi xanh 100% mà vẫn sai trên hình:

1. Visual bind vào thứ không còn tồn tại (hierarchy tự sinh, measure tưởng thừa).
2. Measure đúng giá trị nhưng **không vừa ô** (chữ bị cắt, hộp ghi chú tràn dòng).
3. Dải cảnh báo của Desktop in vào PDF.

Nên mỗi vòng vẫn phải xuất PDF, cắt PNG và **đọc** (Read tool trên `review/render_png/page_N.png`), và khi đổi
cách vẽ thì đếm điểm ảnh (`so_diem_anh.py review/render_rNN review/render_rMM`) — mắt người từng "thấy" đường đã
bo tròn trong khi 0 điểm ảnh đổi.

Với `so_diem_anh.py`: viết **trước** danh sách vùng dự đoán sẽ khác (`--du-doan`). "0 khác biệt" có thể là thất bại
(thay đổi không ăn); khác ở chỗ ngoài dự đoán là lỗi. Một trang đáng lẽ không đổi mà ra 0 khác biệt là bằng chứng
đường ống render ổn định.

## Lưới toạ độ chung (ý tưởng, chưa đóng gói)

Finance có `kiem_luoi_toa_do.py`: các trang chung khung (dải màu, nút chuyển trang, hàng thẻ, slicer kỳ) phải cùng
toạ độ trong dung sai 2px — lệch 8,8px của nút chuyển trang từng sống sót qua mọi cổng số. Script đóng cứng vai trò
khung của dự án đó nên không chép; viết lại theo khung của dự án mới nếu báo cáo có nhiều trang chung bố cục.
Bản gốc: `D:\Giang-day\Soan bai\Project DA\Finance\PowerBI\tools\kiem_luoi_toa_do.py`.
