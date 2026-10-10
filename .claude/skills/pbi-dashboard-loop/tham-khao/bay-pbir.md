# Bẫy Power BI Desktop / PBIR / pbi-cli đã trả giá

Chắt từ ba dự án SupplyChain (08/2026), Finance AIVB De1 v1–v3 (08–10/2026), Game Jelly Trail (09/2026).
Mỗi bẫy kèm cách nhận ra và cách tránh. **Tra bằng grep, không nạp trọn.**

## 1. Vòng lặp và Desktop

| # | Bẫy | Cách tránh |
|---|---|---|
| D1 | `pbi report reload` (Ctrl+Shift+F5) **không** nạp lại JSON report sửa ngoài đĩa | Đóng KHÔNG LƯU và mở lại Desktop mỗi vòng (`pbi_desktop.py open`) |
| D2 | Ctrl+S trong Desktop ghi đè bản sửa trên đĩa bằng bản cũ trong bộ nhớ | Không bao giờ gửi Ctrl+S vào cửa sổ Desktop |
| D3 | Chạy `build_model` khi Desktop đang mở → dải *"This project's files were changed externally"* in lên **mọi trang PDF** | Sinh lại model thì đóng/mở lại Desktop trước khi xuất |
| D4 | Mỗi lần `taskkill` để lại một `AnalysisServicesWorkspace_*`; chọn nhầm instance thì TMSL bắn một nơi, đếm dòng đọc nơi khác | Chọn workspace theo `msmdsrv.port.txt` **mới nhất** (shim `pbi.exe` đã làm vậy) |
| D5 | Nhiều cửa sổ PBIDesktop cùng mở → export xuất nhầm báo cáo, không báo gì | `export_pdf.ps1 -TitleLike '*TenDuAn*'`; bản không truyền sẽ báo lỗi khi có ≥2 cửa sổ |
| D6 | Tên dự án là tiền tố của nhau (`X` và `X_v3`) → khẳng định tiêu đề cửa sổ khớp nhầm | Đặt tên không lồng nhau |
| D7 | Hộp thoại lỗi modal (`Issues were found`…) khoá cửa sổ chính, mọi bước sau chết im lặng | `pbi_health.ps1` sau mỗi lần mở; nó chụp hộp thoại và in nguyên văn chi tiết |
| D8 | `.ps1` trên ổ Google Drive bị chặn (`AuthorizationManager check failed`) | Luôn gọi `powershell -ExecutionPolicy Bypass -File` |

## 2. Nạp dữ liệu

| # | Bẫy | Cách tránh |
|---|---|---|
| N1 | Refresh cả model song song với nhiều native query PostgreSQL → `0x80040E4E` / *Exception while reading from stream* | Refresh tuần tự từng bảng (`bang_nap` trong `loop.json`), đặt `parallelQueryLoading: false` |
| N2 | Bảng chỉ chứa measure (`_Measures`) **cũng có partition**; bỏ sót khi refresh → dải *"Some of the tables have incomplete or no data"* in vào PDF | Liệt kê nó **đầu tiên** trong `bang_nap` |
| N3 | Bảng có cột tính đọc bảng khác (RELATED…) nạp trước bảng nguồn → sai | Bảng nguồn trước, bảng dùng sau |
| N4 | Tham số M (`PgServer`, `NgayCat`…) gây dải *"pending changes in your queries"* dù giá trị không đổi | Viết thành biểu thức thường thay vì Parameter; loop có cổng `pbi_ui.ps1` dò dải này |
| N5 | Native query mới/đổi → Desktop hỏi duyệt từng khối (modal), refresh treo im lặng tới khi bấm **Run** | Tắt *Require user approval for new native database queries* (Options → Global → Security) — thiết lập theo **từng máy** |
| N6 | Sau refresh còn dải vàng **Apply changes** → PDF dính dải, trang co lại | Bấm Apply changes trước khi xuất |
| N7 | `localSettings.json` của SemanticModel giữ `securityBindingsSignature`; không chép → refresh treo chờ chứng thực, không hiện hộp thoại | Chép kèm khi nhân bản dự án trên cùng máy (DPAPI theo người dùng) |
| N8 | PostgreSQL cert tự ký + *Encrypt connection* bật → lỗi chuỗi tin cậy | Tắt Encrypt connection cho nguồn đó (theo máy) |
| N9 | Navigator (`Source{[Schema=…,Item=…]}`) fold được và không hỏi duyệt; `[Query="…"]` thì không fold | Ưu tiên Navigator; chỉ dùng native query khi phải gộp bảng khổng lồ phía server |
| N10 | Không tạo VIEW trên DB dùng chung với học viên (đường rò đáp án vòng qua kiểm cách ly) | Gộp bằng native query hoặc Power Query |

## 3. PBIR (report JSON)

| # | Bẫy | Cách tránh |
|---|---|---|
| R1 | `pbi report set-theme` làm Desktop bỏ **toàn bộ** layout, canvas trắng, không báo lỗi | Không dùng; đặt màu trực tiếp trên từng visual, gỡ `customTheme` |
| R2 | Tiêu đề visual nằm ở `visualContainerObjects`, không phải `objects` — đặt nhầm thì Desktop lặng lẽ bỏ qua | `objects` = định dạng mức dữ liệu (màu cột), `visualContainerObjects` = khung, tiêu đề |
| R3 | Tên visual của pbi-cli không khớp Desktop: `stackedBarChart`, `funnelChart` không tồn tại → ô trống | Dùng `barChart`, `funnel` |
| R4 | `pbi visual bind` quyết định Measure/Column theo **tên vai trò**: role `Values` của slicer/card/table luôn ghi thành Measure | Sau bind, sửa lại: gì không thuộc bảng measure là cột |
| R5 | `barChart` dùng role `Series`, pbi-cli sinh `Legend` → cả thanh một màu | Đổi sang `Series` |
| R6 | `"active": true` trên projection đầu của role `Values` → `tableEx` chỉ hiện một cột | Chỉ đặt `active` cho Category/Series/Rows/Columns |
| R7 | Tô màu từng nhóm cần `dataPoint` có `selector` dạng `scopeId`; không selector thì mọi nhóm cùng màu | Selector theo giá trị nhóm |
| R8 | Đường một series tô bằng `dataViewWildcard` → chấm tròn mọi điểm, `showMarker=false` không tắt được | Bỏ tô màu, dùng màu đầu của theme |
| R9 | Làm mượt đường: khoá là `lineStyles.lineChartType` (`'linear'`/`'smooth'`/`'step'`), **không** phải `smoothLine` hay `interpolation` | Tên thật tra trong `C:\Program Files\Microsoft Power BI Desktop\bin\WebView2Resources\minerva\scripts\desktop.min.js` thay vì đoán |
| R10 | `sortDefinition` trên visual không ăn khi cột sắp không nằm trong projection | Đặt `sortByColumn` ở tầng model |
| R11 | Nhãn cột cao bị đẩy vào trong với chữ tối; nhãn trắng trong cột hẹp bị cắt "40...." | `labels.labelPosition = 'OutsideEnd'`, đặt đỉnh trục; cột hẹp đổi thành thanh ngang |
| R12 | `pbi report validate` chỉ kiểm cấu trúc thư mục + cú pháp JSON | `kiem_schema.py` đối chiếu `$schema` thật (từng bắt 15 lỗi validate bỏ sót, 2 lỗi làm Desktop crash khi lưu) |
| R13 | Textbox (không có role dữ liệu) — `pbi visual bind` không đặt được | Viết thẳng JSON |

## 4. Bố cục (canvas 1280×720)

- Textbox đệm trong lớn: chân trang < 26px bị cắt, hộp ghi chú 2 dòng cần ≥ 44px (3 dòng ~70px), ô 18–22px không đủ cho chữ 9–10pt.
- Slicer dạng dropdown cần ≥ 56px; tắt `header` mặc định (in tên cột thô) và đặt nhãn qua tiêu đề khung.
- Hằng chiều cao thẻ KPI ở script polish và ô lưới ở build_report **phải bằng nhau** (từng lệch 24px → hàng KPI đè hàng biểu đồ).
- Kiểm chồng lấn bằng số, không nhìn: `max(y + height)` của mọi visual ≤ 720.
- Trục tháng dày (≥ 40 nhãn) phải cuộn ngang → lọc thưa nhãn.
- Thẻ trả về chữ (vd "CHƯA ĐO ĐƯỢC") cần hạ cỡ chữ; tên measure in thành nhãn dưới thẻ nên tên phải đi theo nội dung (đừng để "(tỷ)").

## 5. Xuất PDF

- Power BI **không** hỏi nơi lưu: nó render vào `%LOCALAPPDATA%\Temp\Power BI Desktop\print-job-*\` — `export_pdf.ps1` lấy từ đó.
- UIA và bàn phím không điều khiển được Power BI → click chuột thật theo toạ độ + xác nhận bằng màu pixel. Cửa sổ phải **phóng to** thì toạ độ mới đúng. Không động chuột ~25–60s.
- Lần xuất **đầu tiên** sau khi mở Desktop vẽ trang đang mở thu ~88% → xuất một lần "khởi động" rồi bỏ (loop đã làm).
- Xuất ngay sau khi mở tệp có thể ra số **chưa tính xong** (PDF vẫn đẹp, vẫn đủ trang) → đợi visual tính xong; cổng số không bắt được lỗi này, chỉ so điểm ảnh/nhìn mới bắt.
- Màn hình khoá → không click được. Qua RDP: lọc `LogonUI` theo **session của chính tiến trình**, không hỏi cả máy.
- Đóng Acrobat trước khi xuất (khoá tệp PDF cũ).

## 6. pbi-cli và shim `pbi.exe`

- Multi-line DAX (VAR/RETURN) không truyền được bằng `-e` → `--file` hoặc stdin.
- pbi-cli in ra console cp1252 → tên measure tiếng Việt làm nó chết lúc in. Ép `PYTHONIOENCODING=utf-8` cho tiến trình con.
- Shim phải là `.exe` (subprocess `shell=False` không đọc PATHEXT). Tệp DLL `Microsoft.PowerBI.AdomdClient.dll` nhưng namespace `Microsoft.AnalysisServices.AdomdClient`.
- `conn.Databases` rỗng khi chuỗi kết nối thiếu `Catalog` → hỏi `DBSCHEMA_CATALOGS` rồi `ChangeDatabase`; nhiều catalog thì chỉ rõ bằng `PBI_CATALOG`.
- Shim ghi stdout phải UTF-8 (`Console.OutputEncoding`), không thì mọi chuỗi tiếng Việt thành `?` và cổng **âm thầm co lại** khi lọc theo tên nhóm.
- Đừng thử shim bằng PowerShell 5.1 (ép tham số về ANSI) — thử bằng Python.
- `pbi tmsl '<TMSL>'` (chỉ có ở shim) refresh một bảng không cần tự động hoá giao diện.

## 7. DAX

- `TRUE` / `FALSE` trần không hợp lệ → `TRUE()` / `FALSE()`. Một measure sai cú pháp làm hỏng **cả MdxScript**: mọi card "See details / Fix this", lỗi chỉ nói `Failed to resolve name 'SYNTAXERROR'`. Bắt offline bằng `dax_lint.py`, online bằng `pbi_measures_check.py` (DMV `TMSCHEMA_MEASURES.ErrorMessage`).
- Tên bảng trùng từ khoá DAX (`KPI`) phải có nháy đơn: `'Kpi'[x]`.
- Tham chiếu cột dòng (`DimMonth[MonthKey]`) bên trong `CALCULATE` mất nghĩa vì chuyển ngữ cảnh → bắt vào `VAR` **trước** `CALCULATE`.
- SQL làm tròn (`round(…,4)`), DAX thì không → khớp bản mẫu phải làm tròn cùng chỗ.
- `DIVIDE(BLANK, x)` ra BLANK trong khi bản mẫu ra 0 → `COALESCE(…, 0)` khi định nghĩa yêu cầu.
- Slicer kỳ báo cáo gắn thẳng vào bảng tháng làm sập trục biểu đồ (còn một điểm) → bảng kỳ **rời**, measure tự đọc `SELECTEDVALUE`.
- Star schema làm bộ lọc lan sang trang không muốn → measure trang đó tự gỡ bộ lọc.
- Đường đi nhập nhằng giữa hai bảng: Power BI **tự tắt một cạnh, không báo** → merge trong Power Query.

## 8. Thứ Desktop tự thêm không mặc nhiên là rác

Hierarchy tự sinh (`MonthEnd Hierarchy`) và measure trông như bản sao (`Remaining %`) từng là thứ visual bind vào; xoá khi sinh lại model → "Something's wrong with one or more fields" mà cổng số vẫn xanh. Muốn biết cái nào bỏ được: xuất PDF và so điểm ảnh, không đọc TMDL rồi đoán.
