# -*- coding: utf-8 -*-
"""Phần thuyết trình (speaker notes) cho Slide_Project_1.pptx.

build_slides.py gọi apply_notes(prs, REF, KEY, vn, pct) ngay trước khi lưu.
- Mỗi mục = (đoạn chữ phải có trên slide, nội dung notes). Đoạn chữ dùng để assert notes gắn đúng slide;
  thêm/bớt slide thì sửa danh sách này cho khớp.
- Mọi con số lấy từ ref_numbers.json / Instructor Key, không gõ tay.
- Slide học viên: notes chỉ gợi hướng, không lộ đáp án. 4 slide phụ lục (ẩn): notes nêu đáp án đầy đủ.
Thời lượng gợi ý cả buổi ≈ 195 phút, ghi ở slide đầu mỗi phần.
"""


def build_notes(REF, KEY, vn, pct):
    import json
    from pathlib import Path
    REF_PQ = json.loads((Path(__file__).resolve().parent / "pq_steps.json").read_text(encoding="utf-8"))
    m, lm, sp = REF["month"]["current"], REF["month"]["lm"], REF["month"]["sply"]
    ytd, ytd_sp = REF["ytd"]["current"], REF["ytd"]["sply"]
    rows, thr, br = REF["rows"], REF["threshold"], REF["bridge_2020"]
    rm = REF["report_month"]
    RM, LM_, SP_ = f"{rm % 100:02d}/{rm // 100}", "11/2020", "12/2019"
    if rm != 202012:  # các câu chữ dưới viết cho tháng mặc định 12/2020
        raise ValueError("slide_notes viết cho tháng báo cáo 12/2020; đổi tháng thì sửa lại notes")
    ch = lambda a, b: pct((a - b) / b, 1, sign=True)
    anom = {a["anomaly_id"]: a for a in REF["anomalies"]}
    k26, k27, k28 = anom["K26"], anom["K27"], anom["K28"]
    hq20 = sum(a["effect_usd"] for a in REF["anomalies"]
               if a["type"] == "E_MANUAL_ADJ" and a["date_from"].startswith("2020"))
    fin12 = REF["finance_month"][str(rm)]
    d12 = fin12 - hq20 - m["net_revenue"]
    bsm = KEY["bridge_store_month"]
    cell = lambda sk, mid: bsm[(bsm.store_key == sk) & (bsm.month_id == mid)].iloc[0]
    onl12, s12 = cell(0, rm), cell(int(k28["store_key"]), rm)
    onl12_adj = onl12.finance_net - (onl12.net_sales_post_date - onl12.returns_return_date)
    n_over20 = len(REF["over_threshold_2020"])
    gap20 = br["finance_net"] - br["orders_net_revenue"]
    canc20 = KEY["planted_anomalies"]
    canc20 = canc20[(canc20.type == "A_CANCELLED_BOOKED") & (canc20.date_from.astype(str).str.startswith("2020"))]
    sysc = REF["systemic"]
    yr19, yr20 = REF["year"]["2019"], REF["year"]["2020"]

    N = []
    add = lambda key, txt: N.append((key, " ".join(txt.split())))

    # ------------------------------------------------------------------ mở đầu
    add("EDA MENTORSHIP", f"""
    Giới thiệu buổi học: Project 1 của nhóm Sales Operations, tình huống Nexa Home – chuỗi bán lẻ điện máy bán qua
    cửa hàng (POS) và website/app. Mục tiêu: học viên dựng dashboard Excel 4 sheet (Power Query, Power Pivot, DAX) từ dữ liệu đơn hàng, tính đúng
    Gross Sales, Net Revenue, số đơn, AOV, ASP, rồi đối soát với file doanh thu Finance và giải thích các chênh lệch
    vượt ngưỡng. Đây là project đầu tiên nên mentor hỗ trợ nhiều (khoảng 80 %): đi cùng học viên qua từng bước, học
    viên tự làm phần giải thích chênh lệch. Giới thiệu mentor và quy ước hỏi đáp trong buổi.""")
    add("Outline", """
    Mạch buổi học (≈ 195 phút): hiểu bối cảnh bán lẻ và vì sao phải đối soát (15') → đề bài dashboard (20') → hệ
    thống KPI (20') → dữ liệu, Power Query và mô hình (40') → khám phá dữ liệu (15') → mẫu DAX (35') → dựng dashboard
    Excel, đọc kết quả và lập luận (40') → sản phẩm nộp và chấm điểm (10'). Nhấn mạnh: phần 3 và 4 là nền móng; sai định nghĩa KPI hoặc sai số dòng
    sau khi nạp thì mọi con số phía sau đều sai, kể cả bảng đối soát.""")

    # ------------------------------------------------------------------ phần 1
    add("còn 10 phút trước cuộc họp", f"""
    (Phần 1 · khoảng 15 phút.) Mở bằng tình huống: 10 phút nữa họp với Finance, Sales Manager cần trả lời 3 câu: tháng
    này thực sự bán được bao nhiêu; bao nhiêu bị mất vì giảm giá, hoàn trả, huỷ đơn; và vì sao số của Sales khác số
    Finance đã xác nhận. Vai trò của Data Analyst là trả lời cả ba bằng con số kiểm chứng được. Hỏi lớp: “Nếu số của
    bạn lệch Finance 2 %, bạn tin số nào? Cần gì để thuyết phục người kia?” – dẫn sang ý: phải giải thích được
    chênh lệch đến từ đâu, bao nhiêu tiền.""")
    add("Nexa Home bán hàng thế nào", f"""
    Quy mô dữ liệu: 66 cửa hàng ở 8 quốc gia cộng một kênh online, {vn(rows['dim_product'])} sản phẩm,
    {vn(rows['orders_total'])} đơn từ 01/2016 đến 20/02/2021. Đi qua 5 bước của một đơn: đặt hàng (order_date; POS trả
    tiền và nhận hàng ngay), giao hàng (delivery_date, chỉ online), hoàn tất (chỉ đơn COMPLETED mới là doanh thu), hoàn
    trả (trừ vào tháng trả hàng, không phải tháng bán), Finance ghi nhận (posting_date). Ý chính cần chốt: cùng một
    đơn có thể rơi vào hai tháng khác nhau tuỳ theo mốc ngày mà mỗi bên dùng.""")
    add("Vì sao số đơn hàng khác số Finance", """
    Hai cách ghi sổ cho cùng một giao dịch: dữ liệu đơn ghi theo ngày đặt và còn lẫn đơn huỷ, đơn đang giao, dòng trùng;
    Finance ghi theo ngày hạch toán, online khi giao xong, hoàn tiền khi xử lý, dùng mã cửa hàng riêng. Vì vậy chênh
    lệch không mặc định là lỗi: có lệch thời điểm, lệch quy tắc và lỗi thật. Nếu không đối soát: hai phòng báo hai
    số, thưởng doanh số tính sai, lỗi thật bị che trong tổng. Giới thiệu ba vai trò trong vòng đối soát: DA lập bảng
    chênh lệch, Finance xác nhận phía sổ sách, cửa hàng/e-commerce xác nhận trạng thái đơn khi DA hỏi lại.""")

    # ------------------------------------------------------------------ phần 2
    add("Hai sản phẩm cần nộp", f"""
    (Phần 2 · khoảng 20 phút.) Hai sản phẩm: file Excel .xlsx (Power Query + Data Model + 4 sheet dashboard) và báo
    cáo 8 phần theo guild_report_template. Slicer dùng chung cho cả 4 sheet: Năm + Tháng báo cáo (chọn một, mặc định
    {RM}), Kênh bán, Quốc gia; thẻ KPI đọc slicer qua CUBEVALUE. Bốn sheet (mỗi sheet là một trang) đi theo
    một mạch kể chuyện: chuyện gì xảy ra (Tổng quan) → doanh thu đến từ đâu → lệch Finance ở đâu → doanh thu có sạch
    không. Lưu ý kênh online không thuộc quốc gia nào: khi chọn một quốc gia thì online bị loại khỏi phạm vi.""")
    add("Trang Tổng quan: Điểm chính", """
    Trang 1 trả lời: tháng này bán được bao nhiêu, vì sao tăng/giảm, có khớp Finance không. Mở đầu bằng câu Điểm chính
    ghép bằng công thức Excel từ các ô số. Năm thẻ: Net Revenue, số đơn, AOV, tỷ lệ hoàn trả, chênh lệch Finance. Biểu đồ: đường Net Revenue
    24 tháng màu xám với tháng chọn tô xanh và điểm cùng kỳ được ghi chú; waterfall Gross → Net; hai biểu đồ nhỏ số
    đơn và AOV đặt cạnh nhau (không dùng hai trục trên một biểu đồ) để thấy ngay doanh thu giảm vì ít đơn; và cột Δ sau
    quy đổi theo tháng để dẫn sang trang Đối soát.""")
    add("Trang Doanh thu: doanh thu đến từ", """
    Trang 2 trả lời: doanh thu tập trung ở đâu, kênh nào tăng. Bốn thẻ: Net Sales, Units, tỷ trọng online, giảm giá trên
    Gross. Danh mục và quốc gia vẽ bằng cột ngang xếp hạng: mục lớn nhất tô xanh, còn lại xám, nhãn giá trị và % ở cuối
    thanh. Giải thích vì sao bỏ pie/treemap: mắt so sánh độ dài chính xác hơn góc hay diện tích. Bảng kênh có AOV, so
    với tháng trước và cùng kỳ; Top 10 cửa hàng có thanh nhỏ trong ô.""")
    add("Trang Đối soát: số của mình lệch", f"""
    Trang 3 là trọng tâm. Δ = Finance − dữ liệu đơn; cờ ⚠ khi |Δ| > {vn(thr['abs_usd'])} USD và |Δ %| >
    {pct(thr['pct'], 0)} ở mức tháng × cửa hàng. Biểu đồ phân rã Δ của tháng chọn: phần lệch thời điểm online (xám) →
    từng cửa hàng vượt ngưỡng (đỏ) → phần còn lại. Cột Δ sau quy đổi theo tháng chỉ tô đỏ những tháng có cửa hàng vượt
    ngưỡng. Bảng cửa hàng xếp theo |Δ|. Hỏi lớp: “Vì sao màu đỏ chỉ dùng cho vượt ngưỡng?” – để mắt người xem đi thẳng
    tới chỗ cần điều tra.""")
    add("Trang Hoàn trả & chất lượng", """
    Trang 4 trả lời: doanh thu có “sạch” không. Thẻ giá trị hoàn trả, tỷ lệ hoàn trả (thay đổi đo bằng điểm %), đơn huỷ,
    dòng trùng đã loại kèm số đơn đang giao. Đường tỷ lệ hoàn trả 24 tháng, cột ngang lý do trả hàng và cột ngang tỷ
    lệ hoàn trả theo danh mục, nhãn kèm Net Sales để không kết luận vội với danh mục nhỏ. Màu đảo: tỷ lệ hoàn trả tăng
    là xấu nên ▲ hiển thị màu cam.""")
    add("Bảng tự rà soát", """
    Checklist để học viên tự chấm trước khi nộp, chia 3 mức. Bắt buộc: bỏ 120 dòng trùng, chỉ tính COMPLETED, đủ 4
    sheet, slicer dùng chung, map Finance theo hiệu lực, có cờ vượt ngưỡng. Nên có: Δ sau quy đổi online, line chart đủ 24
    tháng, tiêu đề hành động và câu Điểm chính, đơn vị rõ, xám + một màu nhấn. Sáng tạo: sheet chi tiết cửa hàng – ngày,
    Show Details trên PivotTable, bảng giải thích chênh lệch có tổng nguyên nhân bằng Δ, dùng bảng bổ sung.""")
    add("Nguyên tắc trình bày: Storytelling", """
    Sáu bài học của Cole Nussbaumer Knaflic áp vào dashboard này. (1) Bối cảnh: người xem là Sales Manager và Finance
    trước họp chốt số – viết kết luận thành một câu. (2) Chọn hình: đường cho xu hướng, cột ngang cho xếp hạng, bảng cho
    chi tiết. (3) Gỡ rối: bỏ khung, nền màu, chú giải thừa. (4) Dẫn mắt: xám cho bối cảnh, một màu nhấn, đỏ chỉ cho
    vấn đề. (5) Thiết kế: căn thẳng, số thống nhất, màu an toàn cho người mù màu. (6) Kể chuyện: tiêu đề nêu kết luận,
    nguồn và ngày dữ liệu ở chân trang. Hoạt động: cho học viên sửa một tiêu đề mô tả thành tiêu đề hành động.""")
    # ------------------------------------------------------------------ phần 3
    add("Thuật ngữ dùng trong bài", """
    (Phần 3 · khoảng 20 phút.) Trước khi vào KPI, thống nhất cách gọi. Doanh thu theo Chuẩn mực kế toán VAS 14 và chế độ
    kế toán Thông tư 99/2025/TT-BTC (thay Thông tư 200 từ 01/01/2026): doanh thu bán hàng trừ các khoản giảm trừ doanh thu
    (chiết khấu thương mại, hàng bán bị trả lại) ra doanh thu thuần – trong file và số đo là Gross Sales, Discounts,
    Returns, Net Revenue. "Finance" trong tên file và số đo chính là phòng Tài chính – Kế toán. Tên menu Excel ghi theo
    giao diện tiếng Việt của Microsoft kèm tiếng Anh, vì máy học viên có thể dùng một trong hai.""")
    add("Cây KPI", """
    Đọc cây KPI từ trái sang phải: Gross Sales trừ giảm giá ra Net Sales, trừ hoàn trả ra
    Net Revenue, rồi so với Finance Net Revenue ra Δ và cờ ngưỡng. Nhánh trên: Orders và Units là mẫu số của AOV và
    ASP, với tử số chung là Net Sales. Thông điệp: hiểu định nghĩa và thứ tự trước, viết DAX sau.""")
    add("Nhóm doanh thu: định nghĩa", f"""
    Đi qua 5 chỉ tiêu nhóm doanh thu và phạm vi: chỉ đơn COMPLETED, đã bỏ dòng trùng; Returns tính theo ngày trả hàng.
    Làm ví dụ tính tay ở khung bên phải (sheet sample_transactions): đơn POS dùng mã BF15, Gross 5.509,99 USD, giảm giá
    826,50 USD, đúng 15 %, ra Net Sales 4.683,49 USD; không có hoàn trả nên Net Revenue bằng Net Sales. Hoạt động: cho
    học viên tự mở sheet, lọc order_number này và tính lại bằng Excel trước khi viết DAX. Chốt: Net Revenue là số đem
    đi đối soát.""")
    add("Nhóm đơn hàng và giá", f"""
    Orders đếm phân biệt order_number, không đếm dòng. Tháng {RM}: {vn(m['orders'])} đơn, {vn(m['units'])} sản phẩm,
    AOV {vn(m['aov'], 2)} USD/đơn, ASP {vn(m['asp'], 2)} USD/sản phẩm, tỷ lệ hoàn trả {pct(m['return_rate'], 2)}. Giới
    thiệu bốn mốc thời gian: tháng báo cáo, tháng trước (LM), cùng kỳ năm trước (SPLY) và lũy kế năm (YTD). Bẫy hay
    gặp: COUNT trên bảng dòng làm số đơn tăng gấp nhiều lần; tính AOV trên Gross thay vì Net Sales làm AOV cao hơn
    thực tế.""")
    add("Quy tắc trạng thái đơn", f"""
    Ba trạng thái: COMPLETED {vn(rows['orders_completed'])} đơn (tính doanh thu), CANCELLED {vn(rows['orders_cancelled'])}
    đơn và PENDING {vn(rows['orders_pending'])} đơn (không tính). PENDING là đơn online chưa giao tại ngày trích dữ liệu
    20/02/2021. Phía Finance: POS ghi nhận ngày bán, online ghi nhận ngày giao, hoàn tiền ghi theo ngày xử lý (có thể
    trễ vài ngày), net_revenue chưa gồm VAT 10 %. Hỏi lớp: “Nếu so số của mình với cột total_incl_vat thì chuyện gì
    xảy ra?” – mọi ô sẽ lệch khoảng 10 %.""")
    add("Chênh lệch Δ và ngưỡng", f"""
    Giải thích vì sao cần cả hai điều kiện: cửa hàng lớn lệch 300 USD có thể chỉ 0,1 % – không đáng truy; cửa hàng nhỏ
    lệch 5 % nhưng chỉ 50 USD – cũng không đáng công. Hai ví dụ thật tháng {RM}: một cửa hàng lệch dưới
    {vn(thr['abs_usd'])} USD nên trong ngưỡng; cửa hàng S12 lệch {vn(s12.diff_finance_minus_orders, 0, True)} USD
    ({pct(s12.diff_pct, 1, True)}) nên bắt buộc tìm nguyên nhân. Không nói trước nguyên nhân của S12; để học viên tự
    tìm ở trang Đối soát.""")

    # ------------------------------------------------------------------ phần 4
    add("Pipeline dữ liệu", """
    (Phần 4 · khoảng 40 phút.) Đọc sơ đồ theo ba làn: mentor chuẩn bị dữ liệu (nguồn Global Electronics Retailer),
    học viên làm 5 bước trong Excel (Power Query → Data Model → DAX trong Power Pivot → 4 sheet dashboard → báo cáo),
    và các điểm kiểm tra (QC_Summary, số đối soát, rubric). Nhấn mạnh hai file đầu vào của học viên: dataset và
    finance_revenue.xlsx.""")
    add("Cấu trúc bộ dữ liệu", f"""
    Hai file, ba bảng fact với ba grain khác nhau: fact_sales_line {vn(rows['fact_sales_line'])} dòng, mỗi dòng là một
    sản phẩm trong một đơn; fact_return {vn(rows['fact_return'])} dòng, mỗi dòng một lần trả hàng; finance_revenue
    {vn(rows['finance_revenue'])} dòng, mỗi dòng một ngày × một mã cửa hàng Finance. Ba lưu ý: có 120 dòng trùng hoàn
    toàn cần loại; fact_return không có store_key nên phải lấy qua order_number + line_item; fin_store_code không phải
    store_key nên phải map.""")
    add("Các bảng dim và bảng hỗ trợ", """
    Đi nhanh qua các dim: dim_date (một dòng một ngày, phải Mark as date table), dim_store (store_key 0 là online),
    dim_product, dim_channel kèm quy tắc ghi nhận, dim_promotion, dim_order_status (có tính doanh thu không) và
    map_store_finance. dim_customer và fx_rate là bảng bổ sung cho phần sáng tạo, không bắt buộc. Hỏi lớp: “Muốn so
    tháng này với tháng trước bằng DATEADD thì cần gì?” – cần bảng ngày liên tục và đã đánh dấu.""")
    add("Mô hình dữ liệu dạng star schema", """
    Mô hình hình sao, quan hệ 1 → nhiều, lọc một chiều. dim_date nối fact_sales_line qua order_date (active) và
    delivery_date (inactive, bật bằng USERELATIONSHIP khi cần quy online về ngày giao); nối fact_return qua return_date
    và fact_finance qua posting_date. Hai bảng fact không nối trực tiếp với nhau, chúng gặp nhau qua dim_date và
    dim_store. dim_channel nối vào dim_store để slicer Kênh lọc được cả Finance và hoàn trả. Trong Excel: Power Pivot →
    Manage → Diagram View, kéo khoá từ dim sang fact; kéo dim_date lần hai sang delivery_date thì Excel tự để quan hệ
    đó inactive. Phần làm trong Power Query: merge store_key vào fact_return, map store_key cho Finance.""")
    add("map_store_finance", f"""
    Bảng map có hiệu lực theo ngày. Bẫy 1: cửa hàng {REF['recode']['store_key']} dùng mã {REF['recode']['old']} đến
    30/06/2019, sau đó {REF['recode']['new']}; nếu gõ tay bảng mã theo mã cũ thì doanh thu từ 07/2019 không gắn được
    cửa hàng. Luôn dùng map_store_finance và lọc posting_date trong [valid_from, valid_to] – với dữ liệu hiện tại
    bước lọc giữ nguyên số dòng, nó là chốt an toàn khi Finance còn hạch toán mã cũ sau ngày đổi. Bẫy 2: dòng HQ-ADJ là bút toán ở văn phòng chính, store_key trống,
    không gán cho cửa hàng nào mà giải thích ở cấp tổng chuỗi. Kiểm tra sau merge: số dòng Finance vẫn đúng
    {vn(rows['finance_revenue'])}.""")
    add("Quy tắc làm sạch và số dòng", f"""
    Chưa khớp số dòng thì chưa tính KPI. Sau khi nạp: {vn(rows['fact_sales_line'])} dòng, bỏ 120 dòng trùng còn
    {vn(rows['fact_sales_line_clean'])}; tổng {vn(rows['orders_total'])} đơn trong đó {vn(rows['orders_completed'])}
    COMPLETED. Ba quy tắc: bỏ trùng trên toàn bộ cột; lọc COMPLETED trong measure chứ không xoá dòng (để vẫn đếm được
    đơn huỷ); đặt đúng kiểu dữ liệu (ngày là Date, tiền là Fixed decimal, store_key là Whole Number ở mọi bảng). Hoạt động:
    học viên đối chiếu số dòng của mình với QC_Summary ngay tại lớp.""")
    st = REF_PQ["steps"]
    add("Pipeline Power Query: mỗi query", f"""
    Sơ đồ: mỗi làn là một query, mỗi ô là một Applied Step, số dưới ô là số dòng sau bước. Mọi bước làm bằng chuột
    trong Power Query Editor, Excel tự ghi lại thành Applied Steps. Chỉ cần hai lượt Get Data: lượt 1 chọn cùng lúc
    mọi Table trong file dataset (Navigator → Select multiple items), lượt 2 lấy finance_revenue.xlsx. Làm sạch theo
    thứ tự dim → fact_sales_line → fact_return → finance_raw, map_store_finance → fact_finance vì query sau dùng query
    trước. Ô cam là ba chỗ hay sai: Removed Duplicates, Merged Queries và Filtered Rows – dừng lại kiểm số dòng ngay
    sau các bước này. Đổi máy: Data Source Settings → Change Source rồi Refresh All.""")
    add("Power Query · fact_sales_line và fact_return", f"""
    Demo trực tiếp. fact_sales_line: chọn Table tbl_fact_sales_line (không chọn sheet), kiểm kiểu cột, Ctrl+A rồi
    Remove Duplicates – {vn(rows['fact_sales_line'])} → {vn(rows['fact_sales_line_clean'])} dòng – rồi thêm gross_amount
    bằng Add Column → Standard → Multiply và net_amount bằng Subtract (chọn gross_amount trước). fact_return: Merge
    Queries với fact_sales_line theo hai cột order_number + line_item (giữ Ctrl để chọn, cùng thứ tự ở hai bảng), mở
    rộng lấy store_key, channel_code, product_key; số dòng giữ {vn(rows['fact_return'])}. Hỏi lớp: “Nếu ghép trước khi
    bỏ trùng thì sao?” – dòng trả hàng có thể bị nhân đôi. Hoạt động: học viên tự làm, so tên bước ở khung Applied
    Steps với slide.""")
    add("Power Query · fact_finance: gắn store_key", f"""
    fact_finance tạo bằng Reference từ finance_raw, Merge với map_store_finance theo fin_store_code, mở rộng store_key,
    valid_from, valid_to. Điều kiện ngày hiệu lực làm bằng Conditional Column (ô Value chọn Select a column để so hai
    cột ngày), rồi lọc hieu_luc = 1 và bỏ 3 cột phụ – không cần gõ công thức. Số dòng giữ
    {vn(rows['finance_revenue'])} ở mọi bước; 3 dòng HQ-ADJ có store_key trống là đúng thiết kế. finance_raw và
    map_store_finance chỉ là query trung gian.""")
    add("Nạp vào Data Model và cấu hình", f"""
    Close & Load To… chọn Only Create Connection và tick Add this data to the Data Model – không đổ bảng ra sheet vì
    63 nghìn dòng làm file nặng mà không dùng tới. Hai query trung gian chỉ tạo connection. Trong Power Pivot: tạo
    quan hệ, Mark as Date Table, Sort by Column cho month_label, ẩn khoá ở bảng fact. Excel không có cửa sổ chạy truy
    vấn DAX nên kiểm mô hình bằng PivotTable From Data Model: đơn COMPLETED năm 2020 phải ra {vn(ytd['orders'])}.
    Đây là lúc dừng lại cho cả lớp đối chiếu số dòng trước khi sang DAX.""")

    # ------------------------------------------------------------------ phần 5
    add("Doanh thu theo năm", f"""
    (Phần 5 · khoảng 15 phút.) Ở cấp năm hai nguồn gần khớp. Năm 2020 Net Revenue {vn(yr20['net_revenue'] / 1e6, 2)}
    triệu USD, giảm {pct(-(yr20['net_revenue'] - yr19['net_revenue']) / yr19['net_revenue'], 1)} so với 2019. Số đơn giảm
    từ {vn(ytd_sp['orders'])} xuống {vn(ytd['orders'])} trong khi AOV gần như đi ngang ({vn(ytd_sp['aov'])} →
    {vn(ytd['aov'])} USD). Hỏi lớp: “Doanh thu giảm do ít đơn hơn hay đơn nhỏ hơn?” – đáp án: do ít đơn. Năm 2021 chỉ
    có dữ liệu đến 20/02 nên không so sánh trực tiếp.""")
    add("Xu hướng theo tháng 2019–2020", f"""
    Hai đường gần như chồng nhau; đỉnh cuối năm, đáy tháng 4. Năm 2020 sụt mạnh từ tháng 3: khi viết báo cáo, học viên
    cần nêu bối cảnh thay vì chỉ mô tả số. Tháng báo cáo {RM}: Net Revenue {vn(m['net_revenue'] / 1e3)} nghìn USD.
    Chỉ ra chỗ đường Finance lệch đường dữ liệu đơn ở các tháng có nhiều đơn online giao sang tháng sau – đây là gợi
    ý đầu tiên cho phần đối soát.""")
    add("Kênh bán: cửa hàng chiếm phần lớn", """
    Online chỉ chiếm khoảng một phần năm Net Sales nhưng là nguồn chênh lệch lớn nhất khi đối soát theo tháng, vì dữ
    liệu đơn tính theo ngày đặt còn Finance tính theo ngày giao. Đơn đặt cuối tháng, giao đầu tháng sau sẽ “nhảy” sang
    tháng khác. AOV online thấp hơn cửa hàng. Hỏi lớp: “Cộng cả năm thì lệch thời điểm online còn bao nhiêu?” – gần
    bằng 0, chỉ còn phần ở hai đầu năm.""")
    add("Danh mục năm 2020", """
    Đọc hai biểu đồ cùng nhau: danh mục lớn nhất (máy tính) chưa chắc có tỷ lệ hoàn trả cao nhất. Danh mục nhỏ dễ có
    tỷ lệ cao vì ít đơn, nên khi viết báo cáo phải nêu cả quy mô lẫn tỷ lệ. Gợi ý câu hỏi phân tích cho báo cáo: danh
    mục nào vừa lớn vừa có tỷ lệ hoàn trả tăng.""")
    add("Khuyến mãi và trạng thái đơn", f"""
    Giảm giá năm 2020 là {vn(ytd['discounts'])} USD, khoảng {pct(ytd['discounts'] / ytd['gross_sales'], 1)} Gross Sales;
    chương trình lớn nhất là ưu đãi đặt online. Biểu đồ trạng thái cho thấy phần lớn đơn COMPLETED; đơn huỷ và đơn
    đang giao phải loại. Đơn PENDING đều là đơn online đặt sát ngày trích dữ liệu: nếu tính nhầm, doanh thu phía đơn sẽ
    cao hơn Finance.""")

    # ------------------------------------------------------------------ phần 6
    add("Bản đồ số đo", """
    (Phần 6 · khoảng 35 phút.) Bộ bắt buộc chỉ có 21 số đo, chia 5 nhóm theo thứ tự: doanh thu → đơn & giá → Kế toán
    và chênh lệch → quy đổi online → so sánh kỳ; 8 số đo nâng cao để riêng cho phần sáng tạo. Số đo sau dùng lại số đo trước, ví dụ [Net Revenue] = [Net Sales] − [Returns]: sửa định nghĩa
    ở một chỗ thì đúng ở mọi nơi. Đặt tất cả trong bảng _Measures, tạo theo đúng thứ tự trên slide.
    Mỗi số đo mới kiểm ngay bằng PivotTable hoặc một ô CUBEVALUE trước khi dùng tiếp.""")
    add("Tạo số đo trong Power Pivot", f"""
    Hai cách tạo số đo: hộp thoại Power Pivot → Measures → New Measure (chọn bảng _Measures, đặt tên, gõ công thức,
    chọn định dạng) hoặc gõ “Tên := công thức” vào Calculation Area dưới bảng _Measures. Gõ đúng tên như trên slide vì
    ô CUBEVALUE gọi số đo theo tên. DAX chỉ tính số; % thay đổi và câu Điểm chính là công thức Excel thường (FIXED, &)
    trỏ vào các ô CUBEVALUE; CUBERANKEDMEMBER đọc tháng đang chọn trên slicer.
    Demo trực tiếp tháng {RM}.""")

    add("DAX nhóm doanh thu", f"""
    Gross Sales là doanh thu bán hàng, Discounts là chiết khấu thương mại, Returns là hàng bán bị trả lại, Net Revenue
    là doanh thu thuần. gross_amount (quantity × unit_price) đã tính ở Power Query nên Gross Sales chỉ cần SUM; CALCULATE với
    KEEPFILTERS thêm điều kiện COMPLETED cho mọi pivot mà vẫn giữ lựa chọn của slicer. Cú pháp trong Excel là
    “Tên := công thức”. Returns không lọc trạng thái vì mọi dòng trả hàng đều thuộc đơn COMPLETED (QC đã kiểm). Kiểm chứng tháng
    {RM}: Gross {vn(m['gross_sales'], 2)} − Discounts {vn(m['discounts'], 2)} = Net Sales {vn(m['net_sales'], 2)}; trừ
    Returns {vn(m['returns'], 2)} ra Net Revenue {vn(m['net_revenue'], 2)} USD. Cho học viên gõ và đối chiếu số trong
    lớp.""")
    add("DAX nhóm đơn hàng, giá và so sánh", f"""
    DIVIDE thay cho dấu “/” để tránh lỗi chia cho 0. Time intelligence (DATEADD, SAMEPERIODLASTYEAR, DATESYTD) chỉ
    chạy đúng khi dim_date đã Mark as date table. Kiểm chứng tháng {RM}: Net Revenue {vn(m['net_revenue'])} USD; tháng
    trước {vn(lm['net_revenue'])} ({ch(m['net_revenue'], lm['net_revenue'])}); cùng kỳ {vn(sp['net_revenue'])}
    ({ch(m['net_revenue'], sp['net_revenue'])}); YTD {vn(ytd['net_revenue'])} USD. Với tỷ lệ (Return Rate) so sánh bằng
    điểm %, không dùng %.""")
    add("DAX đối soát", f"""
    Finance Net Revenue loại các dòng HQ-ADJ (store_key trống). Δ = Finance − Net Revenue; Δ % chia cho trị tuyệt đối
    của Finance để giữ dấu. [Vượt ngưỡng] trả về 1/0 cho từng ô tháng × cửa hàng; [Số cửa hàng vượt ngưỡng] dùng
    SUMX qua CROSSJOIN tháng × cửa hàng – context transition tính lại measure cho từng ô, nên đúng cả khi chọn một
    tháng lẫn cả năm. Nhấn mạnh: không so tổng của cả bảng,
    vì các lệch trái dấu sẽ triệt tiêu nhau.""")
    add("DAX quy đổi online về ngày giao", f"""
    USERELATIONSHIP bật quan hệ inactive delivery_date chỉ trong measure này; KEEPFILTERS giữ lựa chọn của slicer Kênh.
    POS giữ ngày đặt hàng vì Finance ghi POS theo ngày bán. Ví dụ tháng {RM}: Δ toàn chuỗi {vn(d12, 0)} USD; sau quy đổi
    online phần còn lại nằm ở kênh cửa hàng. Thông điệp: nếu Δ online biến mất sau quy đổi thì đó là lệch thời điểm,
    không phải lỗi dữ liệu.""")

    # ------------------------------------------------------------------ phần 7
    add("Bố cục đề xuất 4 trang", """
    (Phần 7 · khoảng 40 phút.) Mở trực tiếp file Nexa_Sales_P1_Dashboard.xlsx – dashboard tham chiếu dựng đúng các
    bước của bài (Power Query, Mô hình Dữ liệu, DAX, slicer, PivotTable, CUBEVALUE); học viên dựng lại bố cục này
    theo 2 slide tiếp theo và so số với file đó. Thứ tự đọc mỗi sheet: câu Điểm chính → thẻ KPI → biểu đồ có tiêu đề nêu kết
    luận → bảng chi tiết. Luôn bắt đầu từ câu hỏi của trang, không bắt đầu từ biểu đồ.""")
    add("Từ measure đến dashboard Excel", """
    Bốn đối tượng Excel đủ dựng cả 4 sheet. Slicer từ Mô hình Dữ liệu (Năm, Tháng, Kênh, Quốc gia), copy sang sheet khác
    là dùng chung bộ lọc; Report Connections nối slicer với mọi PivotTable trên _pivot. Biểu đồ 24 tháng lấy số từ
    bảng phụ CUBEVALUE theo month_id, chỉ truyền slicer Kênh và Quốc gia nên luôn đủ 24 tháng.
    CUBEVALUE đưa một số đo ra một ô theo slicer; các ô đó là nguyên liệu cho thẻ KPI, % thay đổi và câu Điểm chính viết
    bằng công thức Excel. PivotTable From Data Model cho bảng xếp hạng và bảng chi tiết (Top 10). Waterfall vẽ bằng
    loại biểu đồ Waterfall của Excel từ các ô CUBEVALUE, cột cuối Set as total. Mẹo tô màu nhấn: bảng phụ hai cột (nhấn = IF là mục lớn
    nhất, nền = phần còn lại) rồi vẽ cột chồng.""")

    add("Bảng dựng 4 sheet", """
    Bảng này là “bản vẽ thi công” của 4 sheet. Đặt mọi PivotTable nguồn, ô CUBEVALUE và bảng phụ ở sheet _pivot (ẩn
    khi nộp); 4 sheet dashboard chỉ chứa slicer, thẻ KPI, biểu đồ và bảng cho người xem. Giữ cùng vùng ô cho câu Điểm chính
    và thẻ KPI ở mọi sheet. Hoạt động: mỗi nhóm dựng sheet Tong_quan trong 20 phút, so số với file
    Nexa_Sales_P1_Dashboard.xlsx ở tháng 12/2020.""")
    add("Trang Tổng quan, tháng", f"""
    Đọc thẻ: Net Revenue {vn(m['net_revenue'] / 1e3, 1)} nghìn USD, {ch(m['net_revenue'], lm['net_revenue'])} so với
    {LM_}, {ch(m['net_revenue'], sp['net_revenue'])} so với {SP_}. Số đơn {vn(m['orders'])} so với {vn(sp['orders'])} cùng
    kỳ, nhưng AOV {vn(m['aov'])} USD tăng {ch(m['aov'], sp['aov'])} → doanh thu giảm vì ít đơn, không phải vì đơn nhỏ.
    Δ Finance {vn(d12)} USD ({pct(d12 / (fin12 - hq20), 2)}), chưa gồm HQ-ADJ {vn(hq20, 0)} USD. Thông điệp: một mốc so
    sánh không đủ để kết luận.""")
    add("Trang Doanh thu, tháng", """
    Đọc câu Điểm chính: máy tính dẫn đầu, Mỹ là thị trường lớn nhất. Cột ngang xếp hạng cho thấy ngay mục lớn nhất (tô
    xanh) mà không cần đọc chú giải. Bảng kênh: cửa hàng chiếm khoảng ba phần tư Net Sales; cả hai kênh tăng mạnh so
    với tháng trước nhưng giảm sâu so với cùng kỳ; AOV cửa hàng cao hơn online. Hỏi lớp: “Tăng so với tháng trước có
    phải tin tốt không khi cùng kỳ giảm hơn 70 %?”""")
    add("Trang Đối soát, tháng", f"""
    Đối soát theo thứ tự: tổng chuỗi → kênh → cửa hàng → ngày. Tháng {RM}: Δ toàn chuỗi {vn(d12)} USD. Kênh online lệch
    {vn(onl12.diff_finance_minus_orders)} USD và về {vn(onl12_adj)} sau khi quy về ngày giao → lệch thời điểm. Cửa hàng
    S12 lệch {vn(s12.diff_finance_minus_orders, 0, True)} USD và không đổi sau quy đổi → phải xuống từng posting_date
    để tìm. Mẹo: xếp bảng theo |Δ| giảm dần, chỉ đào sâu các dòng có cờ Vượt. Không nói đáp án của S12 – để học viên
    tự tìm trong finance_revenue.xlsx.""")
    add("Trang Hoàn trả & chất lượng, tháng", f"""
    Hoàn trả tháng {RM}: {vn(m['returns'])} USD, tỷ lệ {pct(m['return_rate'], 2)}; thay đổi tính bằng điểm % so với
    tháng trước và cùng kỳ. Lý do trả hàng lớn nhất trong tháng là giao nhầm hàng. Khi đọc bảng danh mục, đọc tỷ lệ
    cùng quy mô. Chất lượng dữ liệu: 120 dòng trùng đã loại và {vn(rows['orders_pending'])} đơn PENDING tại ngày trích –
    hai con số này nên xuất hiện trong phần Data & Preparation của báo cáo.""")
    add("Quy trình đối soát 8 bước", """
    Đi qua 8 bước theo sơ đồ: bỏ trùng, giữ COMPLETED, tính Net Sales và Net
    Revenue phía đơn; map Finance theo hiệu lực và tách HQ-ADJ; ghép theo tháng × cửa hàng, tính Δ và Δ %; với ô vượt
    ngưỡng thì xem theo ngày và theo kênh; cuối cùng lập bảng giải thích. Quy tắc vàng: mỗi nguyên nhân phải có số
    tiền, tổng các nguyên nhân bằng Δ, phần chưa giải thích bằng 0. Bảy nguyên nhân cần dò trong khung bên phải sơ đồ
    là gợi ý hướng tìm, không phải danh sách đáp án theo từng ô.""")
    add("Bảng cầu nối năm 2020", f"""
    Khung học viên phải điền. Hai đầu cầu tính được: Net Revenue dữ liệu đơn năm 2020 {vn(br['orders_net_revenue'], 2)} và
    Finance Net Revenue các cửa hàng {vn(br['finance_net'], 2)} USD, tổng cần giải thích {vn(gap20, 2, True)} USD. Phần
    giữa là bài làm: mỗi dòng một nguyên nhân, số tiền và bằng chứng. Gợi ý: lệch thời điểm online cộng cả năm gần 0
    nhưng từng tháng rất lớn; nguyên nhân phía Finance phải ghi rõ cửa hàng, ngày, số tiền. Không “ép” phần dư vào
    một dòng “khác”.""")
    add("Viết kết luận: Finding", """
    Bốn tầng kết luận. Finding: dữ liệu cho thấy gì, ở đâu, bao nhiêu. Insight: ý nghĩa với kinh doanh, không nhắc lại
    số. Recommendation: ai làm gì, theo dõi chỉ số nào. Limitation: giới hạn ảnh hưởng kết luận và bước kiểm tra tiếp.
    Dùng ví dụ trên slide (Δ online về 0 khi tính theo ngày giao → lệch thời điểm, không phải doanh thu bị mất). Hoạt
    động: mỗi nhóm viết một bộ 4 câu cho một ô vượt ngưỡng của chính mình.""")

    # ------------------------------------------------------------------ phần 8
    add("Hai sản phẩm cần nộp và cấu trúc", """
    (Phần 8 · khoảng 10 phút.) Nộp file Excel .xlsx: đổi đường dẫn nguồn (Data Source Settings) rồi Refresh All phải chạy được, cùng báo cáo PDF 8 phần theo guild_report_template,
    phụ lục có bảng cầu nối và danh sách ô vượt ngưỡng. Phần IV (Main Analysis) là đối soát tháng × cửa hàng × kênh;
    phần VII nêu as-of 20/02/2021, đơn PENDING và các giả định ghi nhận. Số trong báo cáo phải khớp dashboard.""")
    add("Thang chấm điểm 100 điểm", """
    Rubric 100 điểm: 40 kỹ thuật (Power Query và mô hình 10, DAX 12, dashboard 4 sheet Excel 12, trải nghiệm 6), 40 báo
    cáo (cấu trúc 8, đối soát 16, lập luận 10, giới hạn 6) và 20 sáng tạo (sheet chi tiết cửa hàng – ngày, bảng bổ sung, cảnh báo ô vượt ngưỡng hằng
    tháng). Phần đối soát 16 điểm: giải thích đủ mọi ô vượt ngưỡng năm 2020 và cầu nối còn lại bằng 0. Nhắc học viên
    dùng checklist ở slide tự rà soát trước khi nộp.""")
    add("Hỏi & đáp", """
    Mở phần hỏi đáp. Câu hỏi thường gặp: (1) Có cần giải thích ô trong ngưỡng không? – Không bắt buộc. (2) HQ-ADJ đặt
    ở đâu? – Ở tổng chuỗi, không gán cửa hàng. (3) Dùng total_incl_vat được không? – Không, so với net_revenue. (4) Có
    được sửa dữ liệu gốc không? – Không; mọi xử lý làm trong Power Query và ghi lại trong báo cáo. Nhắc hạn nộp và kênh
    hỏi bài.""")

    # ------------------------------------------------------------------ phụ lục mentor (ẩn) — được nêu đáp án
    add("Đáp án cầu nối năm 2020", f"""
    CHỈ MENTOR. Cầu nối 2020: dữ liệu đơn {vn(br['orders_net_revenue'], 2)}; lệch thời điểm online
    {vn(br['adj_online_timing'], 2, True)}; hoàn tiền trễ {vn(br['adj_refund_lag'], 2, True)}; đơn huỷ vẫn ghi nhận
    {vn(br['adj_cancelled_booked'], 2, True)} ({len(canc20)} đơn); Finance thiếu dữ liệu {vn(br['adj_missing_upload'], 2, True)}
    (cửa hàng {k26['store_code']}); VAT {vn(br['adj_vat_included'], 2, True)} ({k27['store_code']}, 06/2020); ghi trùng
    {vn(br['adj_duplicate_posting'], 2, True)} ({k28['store_code']}, {'/'.join(reversed(k28['date_from'].split('-')))}); ra Finance cửa hàng
    {vn(br['finance_net'], 2)}; cộng HQ-ADJ {vn(hq20, 2)} ra tổng Finance. Còn lại chưa giải thích bằng 0 ở mọi ô. Khi
    chấm, chấp nhận gộp lệch thời điểm và trễ hoàn tiền thành một dòng nếu có giải thích.""")
    add("32 ô vượt ngưỡng năm 2020", f"""
    CHỈ MENTOR. Năm 2020 có {n_over20} ô vượt ngưỡng (toàn kỳ {REF['over_threshold_cells']}). Mọi ô online (ONL) là
    lệch thời điểm; các cặp ô trái dấu ở hai tháng liền kề của cùng cửa hàng (ví dụ S16 01–02/2020) là hoàn tiền hạch
    toán trễ; các ô dương lẻ ở cửa hàng là đơn huỷ vẫn được ghi nhận; S24 02/2020 âm lớn là thiếu dữ liệu; S54 06/2020
    là VAT; S12 12/2020 là ghi trùng. Gợi ý dẫn dắt khi học viên bí: “Δ của ô này có bằng giá trị một đơn CANCELLED nào
    cùng ngày không?”""")
    add("Chênh lệch cài sẵn năm 2020", f"""
    CHỈ MENTOR. Cách phát hiện từng loại: thiếu dữ liệu ({k26['store_code']}, {vn(k26['effect_usd'], 2)} USD) – Finance
    không có dòng ở ngày cửa hàng có bán; đơn huỷ vẫn ghi nhận – ghép đơn CANCELLED với posting_date cùng cửa hàng;
    ghi trùng ({k28['store_code']}, {vn(k28['effect_usd'], 2, True)} USD) – hai dòng Finance giống hệt nhau; VAT
    ({k27['store_code']}, {vn(k27['effect_usd'], 2, True)} USD) – cả tháng cao hơn khoảng 10 %. Bẫy phía đơn: không bỏ
    trùng thì Net Sales toàn kỳ cao hơn {vn(sysc['S3_DUPLICATE_LINES'], 2)} USD; tính nhầm PENDING thì cao hơn
    {vn(sysc['S4_PENDING'], 2)} USD.""")
    add("Số kiểm tra KPI", f"""
    CHỈ MENTOR. Dùng khi chấm phần DAX (12 điểm). Lệch quá 0,01 USD thì kiểm theo thứ tự: đã bỏ dòng trùng chưa, đã lọc
    COMPLETED chưa, Returns có theo return_date không, Orders có đếm phân biệt không. Số mốc tháng {RM}: Net Revenue
    {vn(m['net_revenue'], 2)}, Orders {vn(m['orders'])}; YTD 2020 Net Revenue {vn(ytd['net_revenue'], 2)} USD.""")
    return N


def apply_notes(prs, REF, KEY, vn, pct):
    notes = build_notes(REF, KEY, vn, pct)
    slides = list(prs.slides)
    assert len(notes) == len(slides), f"{len(notes)} notes cho {len(slides)} slide"
    from thuat_ngu import chuan_hoa
    for i, (s, (key, txt)) in enumerate(zip(slides, notes), 1):
        allt = " ".join(sh.text_frame.text for sh in _walk(s.shapes) if sh.has_text_frame)
        key = chuan_hoa(key, ui=False)
        assert key in allt, f"slide {i}: không thấy '{key}' – notes lệch slide"
        s.notes_slide.notes_text_frame.text = chuan_hoa(txt, ui=False)


def _walk(shapes):
    for sh in shapes:
        if sh.shape_type == 6:  # group
            yield from _walk(sh.shapes)
        else:
            yield sh
