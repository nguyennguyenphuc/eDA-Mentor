# -*- coding: utf-8 -*-
"""Tổng hợp dùng chung cho dashboard HTML và dashboard Excel.

Đọc lại CHÍNH các file giao học viên (dataset + finance) — là phép tính độc lập với generator.
load() trả về dict: ms (tháng × cửa hàng), cat (tháng × cửa hàng × danh mục), rr (tháng × cửa hàng × lý do trả),
hq (HQ-ADJ theo tháng), st (dim_store), REF và các từ điển tên tiếng Việt.
"""
import json
from functools import lru_cache
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
P1 = HERE.parent
REF = json.loads((HERE / "ref_numbers.json").read_text(encoding="utf-8"))


@lru_cache(maxsize=1)
def load():
    X = pd.read_excel(P1 / "Nexa_Sales_P1_Dataset.xlsx", sheet_name=None)
    fin = pd.read_excel(P1 / "finance_revenue.xlsx", sheet_name="finance_revenue")
    mp = X["map_store_finance"]
    st = X["dim_store"]
    prod = X["dim_product"]

    # ------------------------------------------------------------ tổng hợp phía đơn hàng
    raw = X["fact_sales_line"]
    dup_mask = raw.duplicated(keep="first")
    s = raw[~dup_mask].copy()
    mid = lambda d: d.dt.year * 100 + d.dt.month
    s["month_id"] = mid(s.order_date)
    s["gross"] = s.quantity * s.unit_price
    s["ns"] = s.gross - s.discount_amount
    s = s.merge(prod[["product_key", "category"]], on="product_key")
    c = s[s.order_status == "COMPLETED"]

    ms = c.groupby(["month_id", "store_key"]).agg(gross=("gross", "sum"), disc=("discount_amount", "sum"),
                                                  ns=("ns", "sum"), units=("quantity", "sum"),
                                                  orders=("order_number", "nunique"))
    # Net Sales theo cơ sở ghi nhận của Finance: online theo ngày giao
    c_post = c.assign(pm=np.where(c.store_key == 0, mid(c.delivery_date), c.month_id))
    ms = ms.join(c_post.groupby(["pm", "store_key"]).ns.sum().rename("ns_post").rename_axis(["month_id", "store_key"]),
                 how="outer")
    st_cnt = (s.drop_duplicates("order_number")            # đếm ĐƠN, không đếm dòng
              .groupby(["month_id", "store_key"]).order_status.value_counts().unstack(fill_value=0))
    ms = ms.join(st_cnt.reindex(columns=["CANCELLED", "PENDING"], fill_value=0)
                 .rename(columns={"CANCELLED": "canc", "PENDING": "pend"}), how="outer")
    rawm = raw.assign(month_id=mid(raw.order_date))
    ms = ms.join(rawm[dup_mask].groupby(["month_id", "store_key"]).size().rename("dup"), how="outer")

    r = X["fact_return"].merge(s[["order_number", "line_item", "store_key", "category"]], on=["order_number", "line_item"])
    r["month_id"] = mid(r.return_date)
    ms = ms.join(r.groupby(["month_id", "store_key"]).return_amount.sum().rename("ret"), how="outer")

    # ------------------------------------------------------------ Finance → store_key theo hiệu lực
    fin["posting_date"] = pd.to_datetime(fin.posting_date)
    f = fin.merge(mp[["fin_store_code", "store_key", "valid_from", "valid_to"]], on="fin_store_code", how="left")
    f = f[(f.posting_date >= f.valid_from) & (f.posting_date <= f.valid_to)]
    assert len(f) == len(fin), "mỗi dòng Finance phải map đúng 1 mã hiệu lực"
    f["month_id"] = mid(f.posting_date)
    hq = f[f.store_key.isna()].groupby("month_id").net_revenue.sum()
    fs = f[f.store_key.notna()].assign(store_key=lambda d: d.store_key.astype(int))
    ms = ms.join(fs.groupby(["month_id", "store_key"]).net_revenue.sum().rename("fin"), how="outer")
    ms = ms.fillna(0).reset_index()

    cat = c.groupby(["month_id", "store_key", "category"]).agg(ns=("ns", "sum"), units=("quantity", "sum")).reset_index()
    rcat = r.groupby(["month_id", "store_key", "category"]).return_amount.sum().rename("ret").reset_index()
    cat = cat.merge(rcat, on=["month_id", "store_key", "category"], how="outer").fillna(0)
    rrs = r.groupby(["month_id", "store_key", "return_reason"]).return_amount.sum().rename("ret").reset_index()

    REASON = {"DEFECT": "Lỗi sản phẩm", "CHANGE_MIND": "Đổi ý", "WRONG_ITEM": "Giao nhầm hàng",
              "LATE_DELIVERY": "Giao trễ"}
    CAT_VI = {"Home Appliances": "Gia dụng", "Computers": "Máy tính", "Cameras and camcorders": "Máy ảnh",
              "Cell phones": "Điện thoại", "TV and Video": "TV & Video", "Games and Toys": "Game & Đồ chơi",
              "Audio": "Âm thanh", "Music, Movies and Audio Books": "Nhạc, phim, sách nói"}
    COUNTRY_VI = {"Australia": "Úc", "Canada": "Canada", "Germany": "Đức", "France": "Pháp", "Italy": "Ý",
                  "Netherlands": "Hà Lan", "United Kingdom": "Anh", "United States": "Mỹ", "Online": "Online"}


    cur = ms[ms.month_id == REF["report_month"]]
    assert abs(cur.ns.sum() - cur.ret.sum() - REF["month"]["current"]["net_revenue"]) < 0.05
    assert int(cur.orders.sum()) == REF["month"]["current"]["orders"]
    assert int(ms.pend.sum()) == REF["rows"]["orders_pending"]
    assert int(ms.canc.sum()) == REF["rows"]["orders_cancelled"]
    return dict(ms=ms, cat=cat, rr=rrs, hq=hq, st=st, REASON=REASON, CAT_VI=CAT_VI, COUNTRY_VI=COUNTRY_VI)
