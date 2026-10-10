# -*- coding: utf-8 -*-
"""
Chỗ thử M nhanh: chạy một đoạn M trong chính Power BI Desktop đang mở và in kết quả.

Vì sao cần: viết M mà phải sửa TMDL rồi đóng Desktop, mở lại, refresh thì mỗi vòng mất hơn hai
phút. Với vài chục lần thử cho tám hàm thống kê thì không kham nổi. Script này tạo một bảng tạm
trong model ĐANG CHẠY, refresh riêng nó, in ra rồi xoá — mỗi vòng vài giây.

Bảng tạm chỉ nằm trong bộ nhớ của phiên Desktop, KHÔNG được ghi xuống đĩa (Desktop chỉ ghi TMDL
khi lưu project, mà mình không bao giờ gửi Ctrl+S). Nên nó không đụng vào file nguồn.

Hai điều engine bắt buộc, đã xử lý sẵn ở đây:
  1. Bảng KHÔNG tự suy cột từ M. Không khai báo cột thì DAX báo "does not have any columns".
  2. DAX không giữ thứ tự dòng. Nên phải tự đánh số `stt` rồi sắp theo nó.

Dùng:
    python powerbi/scripts/m_thu.py -f duong/dan/doan.m
    echo 'let x = #table({"a"},{{1}}) in x' | python powerbi/scripts/m_thu.py
    python powerbi/scripts/m_thu.py -f doan.m --dong 30 --giu

Đoạn M phải trả về một BẢNG. Muốn xem một giá trị đơn thì bọc lại:
    let v = <biểu thức> in #table({"kq"}, {{v}})

Mã thoát: 0 chạy được · 1 M lỗi · 3 không nối được Desktop.
"""

import argparse
import json
import os
import subprocess
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BANG = "_MThu"
ENV = dict(os.environ, PYTHONIOENCODING="utf-8")

# Bọc đoạn M của người dùng: đổi mọi bảng trả về thành hai cột cố định (stt, dong) để hợp với
# schema tĩnh khai báo bên dưới. Dòng 1 là tên cột, dòng 2 là vạch ngăn, từ dòng 3 là dữ liệu.
BOC = '''let
    _kq = (
%s
    ),
    _ten = Text.Combine(Table.ColumnNames(_kq), " | "),
    _dong = List.Transform(Table.ToRecords(_kq), each Text.Combine(
                List.Transform(Record.FieldValues(_),
                    (v) => if v = null then "(null)" else try Text.From(v) otherwise "?"),
                " | ")),
    _ds = List.Combine({{_ten}, {Text.Repeat("-", 60)}, _dong}),
    _bang = Table.FromList(_ds, Splitter.SplitByNothing(), {"dong"}),
    _co_stt = Table.AddIndexColumn(_bang, "stt", 1, 1, Int64.Type)
in
    Table.SelectColumns(_co_stt, {"stt", "dong"})'''


def pbi(args, stdin=None):
    return subprocess.run(["pbi"] + args, input=stdin, env=ENV, capture_output=True,
                          text=True, encoding="utf-8", errors="replace")


def loi(r):
    return ((r.stdout or "") + (r.stderr or "")).strip()


def xoa_bang():
    pbi(["table", "delete", BANG, "--yes"])
    pbi(["table", "delete", BANG])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-f", "--file", help="file chứa đoạn M (mặc định đọc stdin)")
    ap.add_argument("--giu", action="store_true", help="giữ bảng tạm lại sau khi chạy")
    ap.add_argument("--dong", type=int, default=15, help="số dòng dữ liệu in ra")
    a = ap.parse_args()

    m = (open(a.file, encoding="utf-8").read() if a.file else sys.stdin.read()).rstrip()
    if not m.strip():
        sys.exit("Không có mã M nào.")

    xoa_bang()
    r = pbi(["table", "create", BANG, "--m-expression", "-", "--hidden"], stdin=BOC % m)
    if r.returncode != 0:
        print("TẠO BẢNG HỎNG:")
        print(loi(r)[-2500:])
        return 1

    for ten, kieu in (("stt", "int64"), ("dong", "string")):
        r = pbi(["column", "create", ten, "-t", BANG, "--data-type", kieu,
                 "--source-column", ten])
        if r.returncode != 0:
            print("TẠO CỘT %s HỎNG:" % ten)
            print(loi(r)[-1500:])
            xoa_bang()
            return 1

    r = pbi(["table", "refresh", BANG])
    if r.returncode != 0:
        print("REFRESH HỎNG — đây là nơi lỗi M hiện ra:")
        print(loi(r)[-3000:])
        if not a.giu:
            xoa_bang()
        return 1

    r = pbi(["--json", "dax", "execute",
             "EVALUATE TOPN(%d, '%s', '%s'[stt], ASC) ORDER BY '%s'[stt] ASC"
             % (a.dong + 2, BANG, BANG, BANG)])
    txt = r.stdout or ""
    i = txt.find("{")
    if i < 0:
        print("Không đọc được kết quả:")
        print((txt + (r.stderr or ""))[-2000:])
        if not a.giu:
            xoa_bang()
        return 1

    for row in json.loads(txt[i:]).get("rows", []):
        v = row.get("[dong]")
        if v is None:
            v = [x for k, x in row.items() if "stt" not in k]
            v = v[0] if v else ""
        print(v)

    if not a.giu:
        xoa_bang()
    return 0


if __name__ == "__main__":
    sys.exit(main())
