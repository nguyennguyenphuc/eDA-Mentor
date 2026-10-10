---
name: pbi-dashboard-loop
description: Dựng hoặc sửa dashboard Power BI (.pbip, PBIR + TMDL sinh bằng script) theo VÒNG LẶP có cổng chặn — sinh model/report, đóng–mở lại Desktop, nạp tuần tự, kiểm measure hỏng, đối chiếu số với mốc tính độc lập, xuất PDF, cắt PNG để đọc, so điểm ảnh giữa hai vòng. Dùng khi người dùng nói "dựng dashboard Power BI", "chạy một vòng", "loop Power BI", "xuất PDF dashboard", "kiểm số dashboard", "so bản render", hoặc bắt đầu dashboard cho một dự án Mentor (Bank, Sales, E1). Kết hợp với các skill power-bi-* (pbi-cli) cho thao tác lẻ và skill goi-ngoai cho việc cơ khí.
---

# Vòng lặp dựng dashboard Power BI

Mọi thay đổi đi qua **một lệnh** `python loop.py`. Cổng xếp **rẻ trước đắt sau**, dừng ở cổng hỏng đầu tiên —
không tốn một vòng xuất-và-nhìn cho bản đã sai số. Gộp từ ba dự án đã chạy thật: SupplyChain, Finance (AIVB
De1 v1–v3, cổng 606 phép), Game Jelly Trail (12 cổng, 2.125 phép).

Đường dẫn trong skill: `<skill>` = thư mục chứa tệp này.

## Bố cục

```
<skill>/scripts/     công cụ dùng chung — gọi tại chỗ, KHÔNG chép vào dự án
  pbi_desktop.py       status | open --pbip X.pbip | close   (đóng không lưu, mở lại, bắt hộp thoại lỗi)
  pbi_health.ps1       hộp thoại modal đang chặn Desktop? chụp + in nguyên văn lỗi
  pbi_ui.ps1           đọc cây UIA của Desktop (dò dải "pending changes"…)
  pbi_measures_check.py  measure nào hỏng và vì sao (DMV TMSCHEMA_MEASURES)
  dax_lint.py          lint DAX tĩnh — import trong build_model.py: lint_measures(MEASURES, cot_theo_bang)
  kiem_schema.py       đối chiếu PBIR với $schema Microsoft: python kiem_schema.py X.Report
  export_pdf.ps1       File > Export > PDF bằng click chuột thật (-OutPdf, -TitleLike '*X*')
  pdf2png.py           PDF -> PNG từng trang
  so_diem_anh.py       so hai thư mục render, gom vùng khác (--du-doan vùng.json)
  bpa_check.py         Best Practice Analyzer trên TMDL (chạy trong thư mục chứa X.SemanticModel)
  diff_measures.py     so thân DAX từng measure giữa hai model
  m_thu.py             thử biểu thức M
  pbi_shim.cs + build_pbi_shim.ps1   dựng `pbi.exe` thay pbi-cli (ADOMD trực tiếp, tự dò cổng)
<skill>/mau/         chép vào dự án
  loop_mau.py  -> loop.py      khung vòng lặp, đọc loop.json
  loop.json                    cấu hình: du_an, bang_nap, truoc_model, kiem_so, scripts
  kiem_so_mau.py -> kiem_so.py cổng số: DAX vs moc.json, gom UNION một lần gọi
  moc.json                     khuôn số kỳ vọng
<skill>/tham-khao/
  bay-pbir.md          bẫy đã trả giá (Desktop, nạp, PBIR, bố cục, xuất PDF, pbi-cli, DAX) — grep, đừng nạp trọn
  cong-nghiem-thu.md   thiết kế cổng: lớp phép kiểm, cổng phải đỏ được, thứ cổng số không bắt
```

## Khởi tạo dự án mới

1. Tạo thư mục `<du-an>/PowerBI/`. Dựng khung `.pbip` (skill `power-bi-report` hoặc mở Desktop tạo rỗng rồi lưu
   dạng project). Tên dự án **không** là tiền tố của dự án khác đang mở.
2. Chép `mau/loop_mau.py` → `loop.py`, `mau/loop.json`, `mau/kiem_so_mau.py` → `kiem_so.py`, `mau/moc.json`.
3. Sửa `loop.json`: `du_an` (tên tệp .pbip không đuôi), `scripts` (đường tương đối từ thư mục dự án tới
   `<skill>/scripts`; với `Mentor/<X>/PowerBI/` là `../../.claude/skills/pbi-dashboard-loop/scripts`),
   `bang_nap` (bảng measure trước, bảng nhỏ trước, bảng nguồn trước bảng có cột tính đọc nó), `truoc_model`, `kiem_so`.
4. Viết `build_model.py` (sinh TMDL + `dax_lint`) và `build_report.py` (sinh PBIR, viết thẳng `visual.json`).
   Bản mẫu đã chạy thật để đọc cách làm:
   - Game: `G:\My Drive\09._PC\Giang-day\Soan bai\Project DA\Game\PowerBI\JellyTrail\build_model.py`, `build_report.py`, `THIET_KE.md`
   - Finance: `D:\Giang-day\Soan bai\Project DA\Finance\PowerBI\tools\build_model_v3.py`, `build_report.py`, `polish.py`, `chrome.py`, `README.md`
5. Viết script sinh `moc.json` bằng đường **độc lập** Power BI (SQL thuần / pandas trên file nguồn).
6. Không có pbi-cli: `powershell -ExecutionPolicy Bypass -File <skill>/scripts/build_pbi_shim.ps1` rồi đặt
   `<skill>/scripts` lên đầu PATH.

## Một vòng

```
python loop.py                 # đủ 12 cổng
python loop.py --no-model      # chỉ đổi trình bày
python loop.py --no-export     # chỉ kiểm số
python loop.py --tu 8          # Desktop đang mở đúng dự án: từ cổng measure trở đi
```

Sau vòng: **đọc** `review/render_png/page_N.png` bằng Read; bản lưu `review/render_rNN/` để
`python <skill>/scripts/so_diem_anh.py review/render_r01 review/render_r02` khi đổi cách vẽ.
Ghi nhật ký vòng (vòng, kết quả cổng, việc sửa sau vòng) vào `BUILD_STATE.md` của dự án như Game đã làm.

## Luật cứng

- **Đóng–mở lại Desktop mỗi vòng** — `pbi report reload` không nạp JSON sửa ngoài đĩa.
- **Không bao giờ gửi Ctrl+S** vào Desktop; không sửa tay TMDL/JSON đã sinh (mất ở lần dựng sau) — sửa script sinh.
- **Không `pbi report set-theme`** (xoá toàn bộ layout không báo lỗi).
- Bước xuất PDF **chiếm chuột ~1 phút**: báo người dùng trước, không chạy khi màn hình khoá.
- Native query mới phải được duyệt tay một lần; sau refresh phải **Apply changes** trước khi xuất.
- Cổng xanh chưa đủ: cổng số không thấy visual hỏng, chữ bị cắt, dải cảnh báo — luôn đọc PNG.
- Cổng mới phải có cách tự chứng minh nó đỏ được (`--tiem`), xem `tham-khao/cong-nghiem-thu.md`.
- Gặp lỗi lạ: grep `tham-khao/bay-pbir.md` trước khi thử đoán (tên thuộc tính visual tra trong `desktop.min.js`
  của bộ cài, không đoán — từng tốn hai vòng ~5 phút mỗi vòng cho `lineChartType`).

## Phối hợp skill khác

- `power-bi-dax`, `power-bi-modeling`, `power-bi-visuals`, `power-bi-report`… (pbi-cli) — truy vấn DAX lẻ, xem
  model, thao tác một visual. Vòng lặp vẫn là nguồn sự thật: thay đổi cuối cùng phải nằm trong script sinh.
- `goi-ngoai` — `agy` cho việc cơ khí trên script dự án (đổi tên hàng loạt, sửa toạ độ visual theo bảng);
  `chatgpt` cho ý kiến thứ hai về bố cục/lời ghi chú (dán mô tả hoặc trích PNG đã đọc).
