# -*- coding: utf-8 -*-
"""
Lint DAX tĩnh cho danh sách measure trước khi ghi TMDL.

Vì sao cần: một measure sai cú pháp **không** làm Power BI Desktop bật dialog lỗi.
Model vẫn nạp, refresh vẫn chạy xong — nhưng cả MdxScript của model hỏng, nên MỌI card
hiện "See details / Fix this" và mọi query trả về:

    MdxScript(Model) (5, 26) Failed to resolve name 'SYNTAXERROR'.

Thông báo đó KHÔNG nói measure nào sai, nên không có cách nào lần ra từ trong Desktop.
Bắt tại đây, offline, trước khi ghi file, là đường duy nhất rẻ.

Bẫy đã dính thật: DAX không chấp nhận `TRUE` / `FALSE` trần như Excel — bắt buộc
`TRUE()` / `FALSE()`. Năm measure `Daily[is_event] = FALSE` đã làm hỏng cả 41 measure.

Dùng:  loi = lint_measures(MEASURES, {"Daily": {"date", "lead_tb", ...}, ...})
       loi rỗng = sạch.
"""

import re

# Từ khoá DAX đứng trước `[...]` nhưng KHÔNG phải tên bảng.
KEYWORDS = {"VAR", "RETURN", "IN", "NOT", "AND", "OR", "DEFINE", "EVALUATE",
            "ORDER", "BY", "ASC", "DESC", "MEASURE", "COLUMN", "TABLE"}

RE_BARE_BOOL = re.compile(r"\b(TRUE|FALSE)\b(?!\s*\()")
RE_TBL_COL = re.compile(r"([A-Za-z_][A-Za-z0-9_]*)\s*\[([^\]]+)\]")
RE_MEASURE_REF = re.compile(r"(?<![A-Za-z0-9_\]])\[([^\]]+)\]")

# Ký tự VBA format string hợp lệ (đủ cho mọi format đang dùng: #,0.00 · 0.0% · +0.0;-0.0;0.0 · yyyy-mm-dd)
RE_FORMAT = re.compile(r"^[#0,.;%+\-\\/$a-zA-Z\" ]+$")


# Tên biến VAR mà DAX từ chối, tìm ra bằng cách thử tách bạch từng cái.
# `tran` báo "The syntax for 'tran' is incorrect" dù không trùng tên bảng, tên
# cột hay hàm nào; đổi thành `nguong` là sạch. Lỗi không bật dialog, không làm
# refresh hỏng, và `pbi report validate` cũng không thấy — nó chỉ làm hỏng cả
# MdxScript nên MỌI thẻ cùng lúc báo "See details".
TEN_BIEN_CAM = {"tran"}

RE_VAR = re.compile(r"\bVAR\s+([A-Za-z_][A-Za-z0-9_]*)")


def lint_measures(measures, cols_by_table):
    """measures: [(tên, biểu thức, formatString|None)]. cols_by_table: {bảng: set(cột)}.

    Trả về list chuỗi lỗi; rỗng nghĩa là sạch.
    """
    loi = []
    ten_measure = {m[0] for m in measures}

    # Trùng tên measure: TMDL sinh ra hai khối cùng tên, Desktop bật dialog
    # "TMDL objects cannot be merged because both declare the same property: expression"
    # và KHÔNG mở được project. `pbi report validate` không thấy vì lỗi nằm ở tầng model.
    da_gap = set()
    for name, _, _ in measures:
        if name in da_gap:
            loi.append("measure '%s': khai báo trùng tên (TMDL sẽ không nạp được)" % name)
        da_gap.add(name)

    for name, expr, fmt in measures:
        def bao(msg, _n=name):
            loi.append("measure '%s': %s" % (_n, msg))

        for tf in RE_BARE_BOOL.finditer(expr):
            bao("dùng `%s` trần — DAX bắt buộc `%s()`" % (tf.group(1), tf.group(1)))

        for mo, dong, dc in (("(", "ngoặc tròn", ")"), ("[", "ngoặc vuông", "]")):
            if expr.count(mo) != expr.count(dc):
                bao("lệch %s (%d mở / %d đóng)" % (dong, expr.count(mo), expr.count(dc)))
        if expr.count('"') % 2:
            bao("số nháy kép lẻ (%d) — chuỗi chưa đóng" % expr.count('"'))

        for ref in RE_TBL_COL.finditer(expr):
            tbl, col = ref.group(1), ref.group(2)
            if tbl.upper() in KEYWORDS:
                continue
            if tbl not in cols_by_table:
                bao("tham chiếu bảng không có trong model: %s[%s]" % (tbl, col))
            elif col not in cols_by_table[tbl]:
                gan = [c for c in cols_by_table[tbl] if c.lower().startswith(col.lower()[:4])]
                bao("bảng %s không có cột `%s`%s"
                    % (tbl, col, (" — gần nhất: %s" % ", ".join(sorted(gan)[:3])) if gan else ""))

        for m in RE_VAR.finditer(expr):
            if m.group(1).lower() in TEN_BIEN_CAM:
                bao("tên biến `%s` bị DAX từ chối, đổi tên khác" % m.group(1))

        for ref in RE_MEASURE_REF.finditer(expr):
            ten = ref.group(1)
            # `[@x]` là cột mở rộng do ADDCOLUMNS/SELECTCOLUMNS sinh ra trong
            # chính biểu thức, không phải measure. Quy ước tiền tố @ có sẵn
            # trong DAX; ở đây chỉ cần bỏ qua để khỏi báo nhầm.
            if ten.startswith("@"):
                continue
            if ten not in ten_measure:
                bao("tham chiếu measure chưa khai báo: [%s]" % ten)

        if fmt is not None and not RE_FORMAT.match(fmt):
            bao("formatString lạ: %r" % fmt)

    return loi
