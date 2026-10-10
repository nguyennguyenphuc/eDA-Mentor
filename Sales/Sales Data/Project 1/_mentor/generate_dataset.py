# -*- coding: utf-8 -*-
"""Sinh bộ dữ liệu Sales Project 1 (Nexa Home) từ Global Electronics Retailer.

Đầu vào : ../source/Global_Electronics_Retailer.zip (tải từ Drive, xem _mentor/README_mentor.md)
Đầu ra  : ../Nexa_Sales_P1_Dataset.xlsx      -- giao học viên
          ../finance_revenue.xlsx            -- file Finance rút gọn, giao học viên
          ./Nexa_Sales_P1_Instructor_Key.xlsx -- đáp án chênh lệch, CHỈ mentor
          ./ref_numbers.json                 -- số tham chiếu cho dashboard/guideline/slide

Chạy:  python generate_dataset.py            (seed cố định 2026, chạy lại cho kết quả y hệt)

Mọi số tiền xử lý bằng cent (số nguyên) để cộng dồn không lệch làm tròn.
"""
import hashlib
import io
import json
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 2026
VERSION = "1.0"
AS_OF = pd.Timestamp("2021-02-20")      # ngày trích dữ liệu
REPORT_MONTH = 202012                    # tháng báo cáo mặc định của dashboard
VAT_RATE = 0.10                          # VAT giả định thống nhất 10 %
THRESH_PCT, THRESH_ABS = 0.01, 200.0     # ngưỡng chênh lệch tháng × cửa hàng × kênh

HERE = Path(__file__).resolve().parent
P1 = HERE.parent
SRC = P1 / "source" / "Global_Electronics_Retailer.zip"
OUT_DATA = P1 / "Nexa_Sales_P1_Dataset.xlsx"
OUT_FIN = P1 / "finance_revenue.xlsx"
OUT_KEY = HERE / "Nexa_Sales_P1_Instructor_Key.xlsx"
OUT_REF = HERE / "ref_numbers.json"

rng = np.random.default_rng(SEED)


# ---------------------------------------------------------------- 1. đọc nguồn
def read_src():
    with zipfile.ZipFile(SRC) as z:
        def rd(name):
            return pd.read_csv(io.BytesIO(z.read(name)), encoding="latin1"
                               if name != "Stores.csv" else "utf-8")
        return {n: rd(n + ".csv") for n in
                ["Sales", "Products", "Stores", "Customers", "Exchange_Rates"]}


def money(s):
    return s.astype(str).str.replace(r"[$,\s]", "", regex=True).astype(float)


src = read_src()
s = src["Sales"].rename(columns={
    "Order Number": "order_number", "Line Item": "line_item", "Order Date": "order_date",
    "Delivery Date": "delivery_date", "CustomerKey": "customer_key", "StoreKey": "store_key",
    "ProductKey": "product_key", "Quantity": "quantity", "Currency Code": "currency_code"})
s["order_date"] = pd.to_datetime(s["order_date"], format="%m/%d/%Y")
s["delivery_date"] = pd.to_datetime(s["delivery_date"], format="%m/%d/%Y")

prod = src["Products"].rename(columns={
    "ProductKey": "product_key", "Product Name": "product_name", "Brand": "brand",
    "Color": "color", "Unit Cost USD": "unit_cost", "Unit Price USD": "unit_price",
    "SubcategoryKey": "subcategory_key", "Subcategory": "subcategory",
    "CategoryKey": "category_key", "Category": "category"})
prod["unit_cost"] = money(prod["unit_cost"])
prod["unit_price"] = money(prod["unit_price"])
prod["subcategory_key"] = prod["subcategory_key"].astype(str).str.zfill(4)
prod["category_key"] = prod["category_key"].astype(str).str.zfill(2)

# ---------------------------------------------------------------- 2. dim_store + map Finance
st = src["Stores"].rename(columns={"StoreKey": "store_key", "Country": "country",
                                   "State": "state", "Square Meters": "square_meters",
                                   "Open Date": "open_date"})
st["open_date"] = pd.to_datetime(st["open_date"], format="%m/%d/%Y")
st["channel_code"] = np.where(st.store_key == 0, "ONLINE", "POS")
st = st.sort_values("store_key").reset_index(drop=True)
ISO3 = {"Australia": "AUS", "Canada": "CAN", "Germany": "DEU", "France": "FRA", "Italy": "ITA",
        "Netherlands": "NLD", "United Kingdom": "GBR", "United States": "USA"}
st["store_code"] = np.where(st.store_key == 0, "ONL", "S" + st.store_key.astype(str).str.zfill(2))
st["seq"] = st.groupby("country").cumcount() + 1
st["fin_code"] = np.where(st.store_key == 0, "ECOM-01",
                          st.country.map(ISO3).fillna("") + "-" + st.seq.astype(str).str.zfill(2))
st["store_name"] = np.where(st.store_key == 0, "Nexa Online (website/app)",
                            "Nexa " + st.state + " (" + st.store_code + ")")
st["region"] = st.country.map({"Australia": "APAC", "Canada": "Bắc Mỹ", "United States": "Bắc Mỹ",
                               "Online": "Online"}).fillna("Châu Âu")
dim_store = st[["store_key", "store_code", "store_name", "channel_code", "country", "region",
                "state", "square_meters", "open_date"]].copy()

# một cửa hàng UK đổi mã Finance từ 01/07/2019 -> bảng map có 2 dòng hiệu lực
RECODE_STORE = int(st[(st.country == "United Kingdom")].store_key.iloc[2])
RECODE_DATE = pd.Timestamp("2019-07-01")
old_code = st.loc[st.store_key == RECODE_STORE, "fin_code"].iloc[0]
new_code = old_code + "N"
mp = []
for r in st.itertuples():
    ch = "E-commerce" if r.store_key == 0 else "Retail"
    if r.store_key == RECODE_STORE:
        mp.append([old_code, r.store_key, ch, pd.Timestamp("2016-01-01"), RECODE_DATE - pd.Timedelta(days=1),
                   "Mã cũ, ngừng dùng sau khi đổi hệ thống kế toán cửa hàng"])
        mp.append([new_code, r.store_key, ch, RECODE_DATE, pd.Timestamp("2099-12-31"),
                   "Mã mới từ 01/07/2019"])
    else:
        mp.append([r.fin_code, r.store_key, ch, pd.Timestamp("2016-01-01"), pd.Timestamp("2099-12-31"), ""])
mp.append(["HQ-ADJ", None, "Head office", pd.Timestamp("2016-01-01"), pd.Timestamp("2099-12-31"),
           "Bút toán điều chỉnh tại văn phòng chính, không thuộc cửa hàng nào"])
map_store_finance = pd.DataFrame(mp, columns=["fin_store_code", "store_key", "fin_channel",
                                              "valid_from", "valid_to", "note"])
map_store_finance["store_key"] = map_store_finance["store_key"].astype("Int64")


def fin_code_of(store_key, d):
    if store_key == RECODE_STORE:
        return new_code if d >= RECODE_DATE else old_code
    return st.set_index("store_key").fin_code[store_key]


# ---------------------------------------------------------------- 3. đơn hàng: giá, khuyến mãi, trạng thái
s = s.merge(prod[["product_key", "unit_price", "unit_cost"]], on="product_key", how="left")
s["channel_code"] = np.where(s.store_key == 0, "ONLINE", "POS")

PROMOS = [  # code, tên, %, (tháng-ngày từ, đến), kênh
    ("NONE", "Không khuyến mãi", 0.00, None, "Tất cả"),
    ("NEWYEAR5", "Năm mới", 0.05, ("01-01", "01-15"), "Tất cả"),
    ("SPRING5", "Mùa xuân", 0.05, ("03-15", "04-15"), "Tất cả"),
    ("BTS10", "Back to school", 0.10, ("08-01", "09-15"), "Tất cả"),
    ("BF15", "Black Friday", 0.15, ("11-20", "11-30"), "Tất cả"),
    ("XMAS10", "Giáng sinh", 0.10, ("12-10", "12-24"), "Tất cả"),
    ("WEB10", "Ưu đãi đặt online", 0.10, None, "ONLINE"),
]
orders = s.groupby("order_number").agg(order_date=("order_date", "first"),
                                       channel_code=("channel_code", "first")).reset_index()
md = orders.order_date.dt.strftime("%m-%d")
orders["promo_code"] = "NONE"
u = rng.random(len(orders))
for code, _, _, win, _ in PROMOS:
    if win:
        inwin = (md >= win[0]) & (md <= win[1])
        orders.loc[inwin & (u < 0.60), "promo_code"] = code
web = (orders.promo_code == "NONE") & (orders.channel_code == "ONLINE") & (rng.random(len(orders)) < 0.20)
orders.loc[web, "promo_code"] = "WEB10"
pct = {p[0]: p[2] for p in PROMOS}

# trạng thái: PENDING = online chưa giao tại AS_OF; CANCELLED ~2,5 %; còn lại COMPLETED
s_deliv = s.groupby("order_number").delivery_date.first()
orders["delivery_date"] = orders.order_number.map(s_deliv)
orders["order_status"] = "COMPLETED"
orders.loc[(orders.channel_code == "ONLINE") & (orders.delivery_date > AS_OF), "order_status"] = "PENDING"
cand = orders.index[orders.order_status == "COMPLETED"]
canc = rng.choice(cand, size=int(round(len(orders) * 0.025)), replace=False)
orders.loc[canc, "order_status"] = "CANCELLED"
orders.loc[orders.order_status != "COMPLETED", "delivery_date"] = pd.NaT

s = s.drop(columns="delivery_date").merge(
    orders[["order_number", "promo_code", "order_status", "delivery_date"]], on="order_number")
s["gross_c"] = (s.quantity * (s.unit_price * 100).round().astype(int)).astype(np.int64)
s["disc_c"] = (s.gross_c * s.promo_code.map(pct)).round().astype(np.int64)
s["net_c"] = s.gross_c - s.disc_c

# ---------------------------------------------------------------- 4. hoàn trả (chỉ dòng COMPLETED)
comp = s[s.order_status == "COMPLETED"]
pick = comp.sample(frac=0.04, random_state=SEED)
base = pick.delivery_date.fillna(pick.order_date)
ret = pick[["order_number", "line_item", "quantity", "unit_price", "promo_code", "channel_code",
            "store_key", "net_c"]].copy()
ret["return_date"] = base + pd.to_timedelta(rng.integers(1, 31, len(pick)), unit="D")
ret = ret[ret.return_date <= AS_OF - pd.Timedelta(days=3)].copy()
ret["return_qty"] = [int(rng.integers(1, q + 1)) for q in ret.quantity]
ret["ret_c"] = (ret.net_c * ret.return_qty / ret.quantity).round().astype(np.int64)
reasons = np.array(["DEFECT", "CHANGE_MIND", "WRONG_ITEM", "LATE_DELIVERY"])
ret["return_reason"] = np.where(ret.channel_code == "ONLINE",
                                rng.choice(reasons, len(ret), p=[.35, .3, .15, .2]),
                                rng.choice(reasons[:3], len(ret), p=[.45, .4, .15]))
# Finance hạch toán hoàn tiền trễ 0-3 ngày sau ngày trả hàng
ret["refund_post_date"] = ret.return_date + pd.to_timedelta(
    rng.choice([0, 1, 2, 3], len(ret), p=[.6, .25, .1, .05]), unit="D")
ret = ret.sort_values(["return_date", "order_number", "line_item"]).reset_index(drop=True)
ret["return_id"] = ["R" + str(i).zfill(6) for i in range(1, len(ret) + 1)]

# ---------------------------------------------------------------- 5. lỗi dữ liệu phía đơn hàng: dòng trùng
dup = s[s.order_status == "COMPLETED"].sample(n=120, random_state=SEED + 1)
s_out = pd.concat([s, dup], ignore_index=True).sort_values(
    ["order_date", "order_number", "line_item"], kind="stable").reset_index(drop=True)

# ---------------------------------------------------------------- 6. Finance (sự thật + lỗi cài sẵn)
anom = []  # nhật ký từng chênh lệch cài sẵn
c = s[s.order_status == "COMPLETED"].copy()
c["post_date"] = np.where(c.channel_code == "ONLINE", c.delivery_date, c.order_date)
c["post_date"] = pd.to_datetime(c["post_date"])
rev = c.groupby(["post_date", "store_key"]).net_c.sum().rename("rev_c")
rf = ret.groupby(["refund_post_date", "store_key"]).ret_c.sum().rename("ref_c")
rf.index.names = ["post_date", "store_key"]
fin = pd.concat([rev, rf], axis=1).fillna(0).astype(np.int64).reset_index()
fin["entry_type"] = "SALES"

# A. Đơn đã huỷ vẫn được Finance ghi nhận (huỷ không đồng bộ sang kế toán) — 25 đơn POS 2019-2020
cx = s[(s.order_status == "CANCELLED") & (s.channel_code == "POS") &
       (s.order_date.dt.year.isin([2019, 2020]))]
cx_orders = rng.choice(np.sort(cx.order_number.unique()), 25, replace=False)
cxl = cx[cx.order_number.isin(cx_orders)]
add = cxl.groupby(["order_date", "store_key"]).net_c.sum().reset_index()
for r in add.itertuples():
    m = (fin.post_date == r.order_date) & (fin.store_key == r.store_key)
    if m.any():
        fin.loc[m, "rev_c"] += r.net_c
    else:
        fin.loc[len(fin)] = [r.order_date, r.store_key, r.net_c, 0, "SALES"]
for o, g in cxl.groupby("order_number"):
    anom.append(dict(anomaly_id="", type="A_CANCELLED_BOOKED", store_key=int(g.store_key.iloc[0]),
                     date_from=g.order_date.iloc[0], date_to=g.order_date.iloc[0],
                     channel="POS", effect_c=int(g.net_c.sum()),
                     detail=f"Đơn {o} trạng thái CANCELLED nhưng Finance vẫn ghi nhận doanh thu"))

# B. Thiếu dữ liệu đẩy lên 3 ngày của 1 cửa hàng Đức (lỗi upload POS)
ger = fin[fin.store_key.isin(st[st.country == "Germany"].store_key) & (fin.post_date.dt.year == 2020)
          & (fin.rev_c > 0)].assign(ym=lambda x: x.post_date.dt.to_period("M"))
de, ym = ger.groupby(["store_key", "ym"]).size().idxmax()          # cửa hàng × tháng có nhiều ngày bán nhất
de = int(de)
miss_days = pd.DatetimeIndex(sorted(ger[(ger.store_key == de) & (ger.ym == ym)].nlargest(3, "rev_c").post_date))
m = fin.store_key.eq(de) & fin.post_date.isin(miss_days)
assert m.sum() == 3
eff = int((fin.loc[m, "rev_c"] - fin.loc[m, "ref_c"]).sum())
fin = fin[~m].copy()
anom.append(dict(anomaly_id="", type="B_MISSING_UPLOAD", store_key=int(de), date_from=miss_days[0],
                 date_to=miss_days[-1], channel="POS", effect_c=-eff,
                 detail="Finance thiếu doanh thu 3 ngày " + ", ".join(miss_days.strftime("%d/%m/%Y"))
                 + " do lỗi đẩy dữ liệu POS → Finance thấp hơn"))

# C. Một cửa hàng Mỹ, tháng 06/2020: ghi doanh thu GỒM VAT (nhầm) → cao hơn 10 %
busy = c[(c.channel_code == "POS") & (c.order_date.dt.to_period("M") == "2020-06")]
us = int(busy[busy.store_key.isin(st[st.country == "United States"].store_key)]
         .groupby("store_key").net_c.sum().idxmax())
m = fin.store_key.eq(us) & (fin.post_date.dt.to_period("M") == "2020-06")
extra = (fin.loc[m, "rev_c"] * VAT_RATE).round().astype(np.int64)
fin.loc[m, "rev_c"] += extra
anom.append(dict(anomaly_id="", type="C_VAT_INCLUDED", store_key=us, date_from=pd.Timestamp("2020-06-01"),
                 date_to=pd.Timestamp("2020-06-30"), channel="POS", effect_c=int(extra.sum()),
                 detail="Cả tháng 06/2020 cửa hàng ghi revenue_recognized đã gồm VAT 10 %"))

# D. Ghi trùng 1 ngày (bút toán đẩy 2 lần) — cửa hàng Pháp, ngày bận nhất tháng 12/2020
fr = st[st.country == "France"].store_key.tolist()
cand = fin[fin.store_key.isin(fr) & (fin.post_date.dt.to_period("M") == "2020-12")]
row = cand.loc[cand.rev_c.idxmax()]
fin = pd.concat([fin, row.to_frame().T], ignore_index=True)
anom.append(dict(anomaly_id="", type="D_DUPLICATE_POSTING", store_key=int(row.store_key),
                 date_from=row.post_date, date_to=row.post_date, channel="POS",
                 effect_c=int(row.rev_c - row.ref_c),
                 detail="Cùng một dòng Finance xuất hiện 2 lần (đẩy bút toán 2 lần)"))

# E. Bút toán điều chỉnh HQ-ADJ, không thuộc cửa hàng nào
adj = [("2018-12-31", -125000, "Đảo dự phòng doanh thu cuối năm 2018"),
       ("2019-12-31", 98000, "Điều chỉnh doanh thu voucher hết hạn 2019"),
       ("2020-12-31", -143550, "Đảo dự phòng doanh thu cuối năm 2020")]
for d, v, note in adj:
    fin.loc[len(fin)] = [pd.Timestamp(d), -1, v, 0, "ADJUSTMENT"]
    anom.append(dict(anomaly_id="", type="E_MANUAL_ADJ", store_key=None, date_from=pd.Timestamp(d),
                     date_to=pd.Timestamp(d), channel="HQ", effect_c=v, detail=note))

fin["post_date"] = pd.to_datetime(fin.post_date)
fin["store_key"] = fin.store_key.astype(int)
fin["fin_store_code"] = [("HQ-ADJ" if k == -1 else fin_code_of(k, d))
                         for k, d in zip(fin.store_key, fin.post_date)]
fin["fin_channel"] = np.select([fin.store_key == -1, fin.store_key == 0],
                               ["Head office", "E-commerce"], "Retail")
fin["net_c"] = fin.rev_c - fin.ref_c
fin["vat_c"] = (fin.net_c * VAT_RATE).round().astype(np.int64)
fin = fin.sort_values(["post_date", "fin_store_code"], kind="stable").reset_index(drop=True)
finance = pd.DataFrame({
    "posting_date": fin.post_date.dt.date,
    "fin_store_code": fin.fin_store_code,
    "fin_channel": fin.fin_channel,
    "entry_type": fin.entry_type,
    "revenue_recognized": fin.rev_c / 100,
    "refunds": fin.ref_c / 100,
    "net_revenue": fin.net_c / 100,
    "vat_amount": fin.vat_c / 100,
    "total_incl_vat": (fin.net_c + fin.vat_c) / 100,
})

anom = pd.DataFrame(anom)
anom["anomaly_id"] = ["K" + str(i).zfill(2) for i in range(1, len(anom) + 1)]

# ---------------------------------------------------------------- 7. cầu nối Orders → Finance theo tháng × cửa hàng
c["month_id"] = c.order_date.dt.year * 100 + c.order_date.dt.month
c["post_month"] = c.post_date.dt.year * 100 + c.post_date.dt.month
ret["ret_month"] = ret.return_date.dt.year * 100 + ret.return_date.dt.month
ret["post_month"] = ret.refund_post_date.dt.year * 100 + ret.refund_post_date.dt.month
fin["month_id"] = fin.post_date.dt.year * 100 + fin.post_date.dt.month


def g(df, mcol, val):
    return df.groupby([mcol, "store_key"])[val].sum()


idx_cols = ["month_id", "store_key"]
br = pd.concat({
    "net_sales_order_date": g(c, "month_id", "net_c"),
    "net_sales_post_date": g(c, "post_month", "net_c"),
    "returns_return_date": g(ret, "ret_month", "ret_c"),
    "refunds_post_date": g(ret, "post_month", "ret_c"),
    "finance_net": g(fin[fin.store_key >= 0], "month_id", "net_c"),
}, axis=1).fillna(0).astype(np.int64)
br.index.names = idx_cols
br = br.reset_index()
br["orders_net_revenue"] = br.net_sales_order_date - br.returns_return_date
br["adj_online_timing"] = br.net_sales_post_date - br.net_sales_order_date
br["adj_refund_lag"] = br.returns_return_date - br.refunds_post_date
an = anom[anom.store_key.notna()].copy()
an["month_id"] = an.date_from.dt.year * 100 + an.date_from.dt.month
an["store_key"] = an.store_key.astype(int)
for t in ["A_CANCELLED_BOOKED", "B_MISSING_UPLOAD", "C_VAT_INCLUDED", "D_DUPLICATE_POSTING"]:
    col = "adj_" + t[2:].lower()
    x = an[an.type == t].groupby(idx_cols).effect_c.sum().rename(col)
    br = br.merge(x, on=idx_cols, how="left")
    br[col] = br[col].fillna(0).astype(np.int64)
adj_cols = [k for k in br.columns if k.startswith("adj_")]
br["explained_finance"] = br.orders_net_revenue + br[adj_cols].sum(axis=1)
br["unexplained"] = br.finance_net - br.explained_finance
assert (br.unexplained == 0).all(), br[br.unexplained != 0].head()
br["diff_finance_minus_orders"] = br.finance_net - br.orders_net_revenue
br["diff_pct"] = np.where(br.finance_net != 0, br.diff_finance_minus_orders / br.finance_net.abs(), np.nan)
br["over_threshold"] = (br.diff_finance_minus_orders.abs() / 100 > THRESH_ABS) & \
                       (br.diff_pct.abs() > THRESH_PCT)
br = br.merge(dim_store[["store_key", "store_code", "channel_code", "country"]], on="store_key", how="left")


# ---------------------------------------------------------------- 8. bảng giao học viên
def dt(x):
    return pd.to_datetime(x).dt.date


fact_sales_line = pd.DataFrame({
    "order_number": s_out.order_number, "line_item": s_out.line_item,
    "order_date": dt(s_out.order_date), "delivery_date": dt(s_out.delivery_date),
    "customer_key": s_out.customer_key, "store_key": s_out.store_key,
    "channel_code": s_out.channel_code, "product_key": s_out.product_key,
    "quantity": s_out.quantity, "unit_price": s_out.unit_price,
    "promo_code": s_out.promo_code, "discount_amount": s_out.disc_c / 100,
    "order_status": s_out.order_status, "currency_code": s_out.currency_code,
    "unit_cost": s_out.unit_cost,
})
fact_return = pd.DataFrame({
    "return_id": ret.return_id, "order_number": ret.order_number, "line_item": ret.line_item,
    "return_date": dt(ret.return_date), "return_qty": ret.return_qty,
    "return_amount": ret.ret_c / 100, "return_reason": ret.return_reason,
})

d0, d1 = pd.Timestamp("2016-01-01"), pd.Timestamp("2021-03-31")
dd = pd.DataFrame({"date": pd.date_range(d0, d1)})
dd["date_key"] = dd.date.dt.strftime("%Y%m%d").astype(int)
dd["year"] = dd.date.dt.year
dd["quarter"] = "Q" + dd.date.dt.quarter.astype(str)
dd["month"] = dd.date.dt.month
dd["month_id"] = dd.year * 100 + dd.month
dd["month_label"] = dd.date.dt.strftime("%m/%Y")
dd["day"] = dd.date.dt.day
dd["weekday_no"] = dd.date.dt.weekday + 1
dd["weekday_name"] = dd.weekday_no.map({1: "Thứ 2", 2: "Thứ 3", 3: "Thứ 4", 4: "Thứ 5", 5: "Thứ 6",
                                        6: "Thứ 7", 7: "Chủ nhật"})
dd["is_weekend"] = dd.weekday_no >= 6
dd["month_end_date"] = (dd.date + pd.offsets.MonthEnd(0)).dt.date
dd["has_data"] = dd.date <= AS_OF
dd["date"] = dd.date.dt.date
dim_date = dd

dim_product = prod[["product_key", "product_name", "brand", "color", "subcategory_key", "subcategory",
                    "category_key", "category", "unit_price", "unit_cost"]].rename(
    columns={"unit_price": "list_price_usd", "unit_cost": "unit_cost_usd"})

cu = src["Customers"].rename(columns={"CustomerKey": "customer_key", "Gender": "gender", "City": "city",
                                      "State": "state", "Country": "country", "Continent": "continent",
                                      "Birthday": "birthday"})
cu["birth_year"] = pd.to_datetime(cu.birthday, format="%m/%d/%Y").dt.year
dim_customer = cu[["customer_key", "gender", "birth_year", "city", "state", "country", "continent"]]

dim_channel = pd.DataFrame([
    ["POS", "Cửa hàng", "Retail", "Ghi nhận doanh thu tại ngày đặt hàng (khách thanh toán và nhận hàng ngay)"],
    ["ONLINE", "Website/app", "E-commerce", "Ghi nhận doanh thu tại ngày giao hàng thành công (delivery_date)"],
], columns=["channel_code", "channel_name", "fin_channel", "revenue_recognition_rule"])

dim_promotion = pd.DataFrame([
    [p[0], p[1], p[2], (p[3][0] if p[3] else ""), (p[3][1] if p[3] else ""), p[4]] for p in PROMOS],
    columns=["promo_code", "promo_name", "discount_pct", "window_from_mmdd", "window_to_mmdd", "channel"])

dim_order_status = pd.DataFrame([
    ["COMPLETED", "Hoàn tất", "Có", "Đơn đã thanh toán (POS) hoặc đã giao thành công (online)"],
    ["CANCELLED", "Đã huỷ", "Không", "Khách huỷ hoặc thanh toán thất bại; loại khỏi mọi KPI doanh thu"],
    ["PENDING", "Đang giao", "Không", f"Đơn online chưa giao tại ngày trích dữ liệu {AS_OF:%d/%m/%Y}; "
                                       "chưa đủ điều kiện ghi nhận doanh thu"],
], columns=["order_status", "status_name", "count_in_revenue", "rule"])

kpi = pd.DataFrame([
    ["Gross Sales", "Σ quantity × unit_price", "Đơn COMPLETED, đã bỏ dòng trùng", "USD"],
    ["Discounts", "Σ discount_amount", "Đơn COMPLETED", "USD"],
    ["Net Sales", "Gross Sales − Discounts", "Đơn COMPLETED", "USD"],
    ["Returns", "Σ return_amount theo return_date", "fact_return", "USD"],
    ["Net Revenue", "Net Sales − Returns", "Con số đem đối soát với Finance", "USD"],
    ["Orders", "Đếm phân biệt order_number", "Đơn COMPLETED", "đơn"],
    ["Units", "Σ quantity", "Đơn COMPLETED", "sản phẩm"],
    ["AOV", "Net Sales / Orders", "Giá trị trung bình mỗi đơn", "USD/đơn"],
    ["ASP", "Net Sales / Units", "Giá bán trung bình mỗi sản phẩm", "USD/sp"],
    ["Return Rate", "Returns / Net Sales", "", "%"],
    ["Finance Net Revenue", "Σ net_revenue (finance_revenue.xlsx)", "Sau khi map fin_store_code → store_key", "USD"],
    ["Δ Finance", "Finance Net Revenue − Net Revenue", "Dương = Finance cao hơn dữ liệu đơn", "USD"],
    ["Δ Finance %", "Δ Finance / |Finance Net Revenue|", "", "%"],
    ["Vượt ngưỡng", f"|Δ| > {THRESH_ABS:.0f} USD VÀ |Δ %| > {THRESH_PCT:.0%}",
     "Xét ở mức tháng × cửa hàng; chỉ những ô vượt ngưỡng mới phải giải thích", "cờ"],
], columns=["kpi", "formula", "scope_note", "unit"])

fx = src["Exchange_Rates"].rename(columns={"Date": "date", "Currency": "currency_code", "Exchange": "rate_to_usd"})
fx["date"] = pd.to_datetime(fx.date, format="%m/%d/%Y").dt.date

# giao dịch mẫu: mỗi tình huống một vài đơn
def first_order(mask, n=1):
    return s.loc[mask].order_number.drop_duplicates().head(n).tolist()


samp = []
def addsamp(orders_, case, explain):
    for o in orders_:
        samp.append((o, case, explain))


addsamp(first_order((s.channel_code == "POS") & (s.order_status == "COMPLETED") & (s.promo_code == "NONE")
                    & (s.order_date.dt.year == 2020)), "POS bình thường",
        "Ghi nhận ngày đặt hàng; Gross = Net vì không khuyến mãi")
addsamp(first_order((s.channel_code == "POS") & (s.promo_code == "BF15") & (s.order_status == "COMPLETED")),
        "POS có khuyến mãi", "Net Sales = Gross − discount_amount (15 %)")
x = s[(s.channel_code == "ONLINE") & (s.order_status == "COMPLETED") &
      (s.order_date.dt.month != s.delivery_date.dt.month) & (s.order_date.dt.year == 2020)]
addsamp(x.order_number.drop_duplicates().head(2).tolist(), "Online giao sang tháng sau",
        "Dữ liệu đơn tính vào tháng đặt; Finance ghi nhận tháng giao → chênh lệch thời điểm")
addsamp(first_order((s.order_status == "CANCELLED") & (s.channel_code == "POS")), "Đơn huỷ",
        "Không tính vào doanh thu")
addsamp(first_order(s.order_status == "PENDING"), "Đơn đang giao",
        "Chưa giao tại ngày trích dữ liệu → không tính, Finance cũng chưa ghi nhận")
r_lag = ret[ret.refund_post_date > ret.return_date].head(1)
addsamp(r_lag.order_number.tolist(), "Có hoàn trả, Finance hạch toán trễ",
        "Returns tính theo return_date; Finance trừ theo ngày hạch toán (trễ 1–3 ngày)")
addsamp(dup.order_number.head(1).tolist(), "Dòng bị trùng", "Cùng order_number + line_item xuất hiện 2 lần → bỏ 1")
sample = []
for o, case, ex in samp:
    for r in fact_sales_line[fact_sales_line.order_number == o].drop_duplicates().itertuples(index=False):
        sample.append([case, ex] + list(r))
sample_transactions = pd.DataFrame(sample, columns=["case", "explanation"] + list(fact_sales_line.columns))

# ---------------------------------------------------------------- 9. QC
qc = []
def chk(name, ok, result, rule):
    qc.append([name, result, rule, "PASS" if ok else "FAIL"])


fs_clean = fact_sales_line.drop_duplicates()
chk("Số dòng fact_sales_line", True, len(fact_sales_line), f"= {len(s)} dòng gốc + 120 dòng trùng cài sẵn")
chk("Số dòng trùng hoàn toàn", len(fact_sales_line) - len(fs_clean) == 120,
    len(fact_sales_line) - len(fs_clean), "= 120 (cố ý, học viên phải loại)")
chk("Khoá order_number+line_item duy nhất sau khi bỏ trùng",
    not fs_clean.duplicated(["order_number", "line_item"]).any(), "0 trùng", "0")
chk("FK store_key → dim_store", fact_sales_line.store_key.isin(dim_store.store_key).all(), "100 %", "100 %")
chk("FK product_key → dim_product", fact_sales_line.product_key.isin(dim_product.product_key).all(), "100 %", "100 %")
chk("FK customer_key → dim_customer", fact_sales_line.customer_key.isin(dim_customer.customer_key).all(),
    "100 %", "100 %")
chk("FK promo_code → dim_promotion", fact_sales_line.promo_code.isin(dim_promotion.promo_code).all(), "100 %", "100 %")
chk("FK order_status → dim_order_status", fact_sales_line.order_status.isin(dim_order_status.order_status).all(),
    "100 %", "100 %")
chk("FK fact_return → dòng COMPLETED",
    fact_return.merge(fs_clean[fs_clean.order_status == "COMPLETED"], on=["order_number", "line_item"]).shape[0]
    == len(fact_return), len(fact_return), "mọi dòng trả hàng khớp 1 dòng bán COMPLETED")
chk("return_qty ≤ quantity", (ret.return_qty <= ret.quantity).all(), "đúng", "luôn đúng")
chk("return_date > ngày giao/ngày đặt", (ret.return_date > ret.return_date - pd.Timedelta(days=1)).all(), "đúng", "")
chk("Online COMPLETED có delivery_date", fact_sales_line[(fact_sales_line.channel_code == "ONLINE") &
    (fact_sales_line.order_status == "COMPLETED")].delivery_date.notna().all(), "100 %", "100 %")
chk("POS không có delivery_date", fact_sales_line[fact_sales_line.channel_code == "POS"].delivery_date.isna().all(),
    "100 %", "theo thiết kế nguồn GER")
chk("discount_amount = gross × %KM (±0,01)",
    ((fact_sales_line.quantity * fact_sales_line.unit_price * fact_sales_line.promo_code.map(pct)
      - fact_sales_line.discount_amount).abs() <= 0.011).all(), "đúng", "±0,01 USD")
chk("Tỷ lệ đơn CANCELLED", abs((orders.order_status == "CANCELLED").mean() - 0.025) < 0.002,
    f"{(orders.order_status == 'CANCELLED').mean():.2%}", "≈ 2,5 %")
chk("Tỷ lệ dòng có hoàn trả", 0.03 < len(ret) / len(comp) < 0.05, f"{len(ret) / len(comp):.2%}", "3–5 %")
chk("Mọi fin_store_code có trong map_store_finance",
    finance.fin_store_code.isin(map_store_finance.fin_store_code).all(), "100 %", "100 %")
chk("Finance: net_revenue = revenue_recognized − refunds",
    ((finance.revenue_recognized - finance.refunds - finance.net_revenue).abs() < 0.005).all(), "đúng", "±0,005")
chk("Finance: vat_amount = 10 % net_revenue (±0,01)",
    ((finance.net_revenue * VAT_RATE - finance.vat_amount).abs() <= 0.011).all(), "đúng", "±0,01")
chk("Cầu nối Orders → Finance giải thích hết chênh lệch", (br.unexplained == 0).all(), "0 USD chưa giải thích",
    "= 0 (Instructor Key)")
chk("dim_date phủ toàn bộ ngày giao dịch", fact_sales_line.order_date.min() >= dim_date.date.min()
    and fact_return.return_date.max() <= dim_date.date.max(), f"{dim_date.date.min()} → {dim_date.date.max()}",
    "phủ đủ")
QC = pd.DataFrame(qc, columns=["Kiểm định", "Kết quả", "Ngưỡng / quy tắc", "Trạng thái"])
assert (QC["Trạng thái"] == "PASS").all(), QC[QC["Trạng thái"] != "PASS"]

# ---------------------------------------------------------------- 10. README + Data_Dictionary
README = [
    ("Tiêu đề", f"Nexa Home — Sales Project 1: Kết quả bán hàng và đối soát doanh thu · Dataset v{VERSION} · seed {SEED}"),
    ("Thời gian", f"Đơn hàng 01/01/2016 → 20/02/2021. Ngày trích dữ liệu (as-of): {AS_OF:%d/%m/%Y}."),
    ("Đơn vị", "Mọi số tiền là USD, chưa gồm VAT. VAT 10 % chỉ xuất hiện trong finance_revenue.xlsx."),
    ("Nguồn", "Global Electronics Retailer (Maven Analytics): giữ nguyên 66 cửa hàng ở 8 quốc gia + 1 kênh online "
              "(store_key = 0). Giảm giá, trạng thái đơn, hoàn trả và file Finance là giả lập có kiểm soát."),
    ("Bảng cần nạp", "dim_date, dim_store, dim_product, dim_channel, dim_promotion, dim_order_status, fact_sales_line, "
                     "fact_return, map_store_finance (+ file finance_revenue.xlsx). Mỗi sheet dữ liệu là một Excel Table "
                     "tên tbl_<tên sheet> — trong Power Query chọn Table thay vì Sheet."),
    ("Bảng bổ sung", "dim_customer, fx_rate và các cột unit_cost, currency_code, customer_key: chỉ dùng cho phần sáng tạo."),
    ("Không nạp", "README, Data_Dictionary, QC_Summary, KPI_Definitions, sample_transactions là tài liệu tham khảo."),
    ("Quan hệ", "dim_date[date] → fact_sales_line[order_date] (active) và [delivery_date] (inactive, dùng USERELATIONSHIP); "
                "dim_date[date] → fact_return[return_date]; các khoá store_key, product_key, promo_code, order_status "
                "nối tới dim cùng tên."),
    ("Quy tắc trạng thái", "Chỉ đơn COMPLETED tính vào KPI doanh thu; CANCELLED và PENDING bị loại (xem dim_order_status)."),
    ("Quy tắc Finance", "POS ghi nhận ngày đặt hàng; online ghi nhận ngày giao (delivery_date); hoàn tiền hạch toán theo "
                        "ngày Finance xử lý, có thể trễ vài ngày sau return_date."),
    ("Ngưỡng", f"Ô tháng × cửa hàng chỉ phải giải thích khi |Δ| > {THRESH_ABS:.0f} USD VÀ |Δ %| > {THRESH_PCT:.0%} "
               "(Δ = Finance − dữ liệu đơn)."),
    ("Làm sạch", "Dữ liệu CHƯA sạch hoàn toàn: có dòng trùng cần loại. Đối chiếu số dòng sau khi làm sạch với QC_Summary."),
    ("Thứ tự gợi ý", "Gross Sales → Discounts → Net Sales → Returns → Net Revenue, rồi mới đối soát với Finance."),
    ("Lưu ý", "Bộ dữ liệu còn các nguyên nhân chênh lệch cài sẵn phía Finance; đáp án nằm trong file riêng của mentor."),
]
REQ, OPT = "Bắt buộc theo đề bài", "Bổ sung cho người học"
COLINFO = {  # bảng: {cột: (nhóm, kiểu, đơn vị, ý nghĩa)}
    "fact_sales_line": {
        "order_number": (REQ, "Số nguyên", "", "Mã đơn hàng; đếm phân biệt ra số đơn"),
        "line_item": (REQ, "Số nguyên", "", "Số thứ tự dòng trong đơn; cùng order_number tạo khoá dòng"),
        "order_date": (REQ, "Ngày", "", "Ngày đặt hàng; mốc thời gian chính của dữ liệu đơn"),
        "delivery_date": (REQ, "Ngày", "", "Ngày giao thành công (chỉ đơn online COMPLETED); mốc ghi nhận online của Finance"),
        "customer_key": (OPT, "Số nguyên", "", "Khách hàng đặt đơn → dim_customer"),
        "store_key": (REQ, "Số nguyên", "", "Cửa hàng xử lý đơn → dim_store; 0 = kênh online"),
        "channel_code": (REQ, "Văn bản", "", "POS hoặc ONLINE → dim_channel"),
        "product_key": (REQ, "Số nguyên", "", "Sản phẩm → dim_product"),
        "quantity": (REQ, "Số nguyên", "sản phẩm", "Số lượng bán"),
        "unit_price": (REQ, "Số thập phân", "USD", "Đơn giá niêm yết, chưa VAT"),
        "promo_code": (REQ, "Văn bản", "", "Mã khuyến mãi áp cho cả đơn; NONE = không khuyến mãi"),
        "discount_amount": (REQ, "Số thập phân", "USD", "Tiền giảm của dòng = quantity × unit_price × discount_pct"),
        "order_status": (REQ, "Văn bản", "", "COMPLETED / CANCELLED / PENDING → dim_order_status"),
        "currency_code": (OPT, "Văn bản", "", "Tiền tệ khách thanh toán; số tiền trong bảng đã quy USD"),
        "unit_cost": (OPT, "Số thập phân", "USD", "Giá vốn mỗi sản phẩm (dùng cho Project 4)"),
    },
    "fact_return": {
        "return_id": (REQ, "Văn bản", "", "Mã lần trả hàng"),
        "order_number": (REQ, "Số nguyên", "", "Đơn gốc"),
        "line_item": (REQ, "Số nguyên", "", "Dòng gốc trong đơn"),
        "return_date": (REQ, "Ngày", "", "Ngày khách trả hàng; mốc tính Returns"),
        "return_qty": (REQ, "Số nguyên", "sản phẩm", "Số lượng trả (có thể trả một phần)"),
        "return_amount": (REQ, "Số thập phân", "USD", "Tiền hoàn = giá trị sau giảm giá của phần trả"),
        "return_reason": (OPT, "Văn bản", "", "DEFECT / CHANGE_MIND / WRONG_ITEM / LATE_DELIVERY"),
    },
    "dim_date": {
        "date": (REQ, "Ngày", "", "Khoá ngày, liên tục 01/01/2016–31/03/2021; đánh dấu Mark as date table"),
        "date_key": (REQ, "Số nguyên", "", "Khoá ngày dạng yyyymmdd"),
        "year": (REQ, "Số nguyên", "", "Năm"), "quarter": (REQ, "Văn bản", "", "Quý (Q1–Q4)"),
        "month": (REQ, "Số nguyên", "", "Tháng 1–12"),
        "month_id": (REQ, "Số nguyên", "", "Tháng dạng yyyymm; dùng để sắp xếp month_label"),
        "month_label": (REQ, "Văn bản", "", "Nhãn tháng mm/yyyy"), "day": (REQ, "Số nguyên", "", "Ngày trong tháng"),
        "weekday_no": (REQ, "Số nguyên", "", "Thứ trong tuần, 1 = Thứ 2"),
        "weekday_name": (REQ, "Văn bản", "", "Tên thứ"), "is_weekend": (REQ, "Đúng/Sai", "", "Thứ 7 hoặc Chủ nhật"),
        "has_data": (REQ, "Đúng/Sai", "", "Ngày ≤ ngày trích dữ liệu"),
        "month_end_date": (REQ, "Ngày", "", "Ngày cuối tháng"),
    },
    "dim_store": {
        "store_key": (REQ, "Số nguyên", "", "Khoá cửa hàng; 0 = Nexa Online"),
        "store_code": (REQ, "Văn bản", "", "Mã ngắn hiển thị (S01…S66, ONL)"),
        "store_name": (REQ, "Văn bản", "", "Tên cửa hàng"), "channel_code": (REQ, "Văn bản", "", "POS hoặc ONLINE"),
        "country": (REQ, "Văn bản", "", "Quốc gia; Online với kênh online"),
        "region": (REQ, "Văn bản", "", "Nhóm quốc gia: Bắc Mỹ, Châu Âu, APAC, Online"),
        "state": (REQ, "Văn bản", "", "Bang / vùng"), "square_meters": (OPT, "Số nguyên", "m²", "Diện tích; trống với online"),
        "open_date": (REQ, "Ngày", "", "Ngày mở cửa"),
    },
    "dim_product": {
        "product_key": (REQ, "Số nguyên", "", "Khoá sản phẩm"), "product_name": (REQ, "Văn bản", "", "Tên sản phẩm"),
        "brand": (REQ, "Văn bản", "", "Thương hiệu"), "color": (OPT, "Văn bản", "", "Màu"),
        "subcategory_key": (REQ, "Văn bản", "", "Mã nhóm hàng con"), "subcategory": (REQ, "Văn bản", "", "Nhóm hàng con"),
        "category_key": (REQ, "Văn bản", "", "Mã nhóm hàng"), "category": (REQ, "Văn bản", "", "Nhóm hàng (8 nhóm)"),
        "list_price_usd": (REQ, "Số thập phân", "USD", "Giá niêm yết theo catalogue"),
        "unit_cost_usd": (OPT, "Số thập phân", "USD", "Giá vốn theo catalogue"),
    },
    "dim_customer": {
        "customer_key": (OPT, "Số nguyên", "", "Khoá khách hàng (đã bỏ tên)"), "gender": (OPT, "Văn bản", "", "Giới tính"),
        "birth_year": (OPT, "Số nguyên", "", "Năm sinh"), "city": (OPT, "Văn bản", "", "Thành phố"),
        "state": (OPT, "Văn bản", "", "Bang / vùng"), "country": (OPT, "Văn bản", "", "Quốc gia của khách"),
        "continent": (OPT, "Văn bản", "", "Châu lục"),
    },
    "dim_channel": {
        "channel_code": (REQ, "Văn bản", "", "POS / ONLINE"), "channel_name": (REQ, "Văn bản", "", "Tên hiển thị"),
        "fin_channel": (REQ, "Văn bản", "", "Tên kênh trong file Finance"),
        "revenue_recognition_rule": (REQ, "Văn bản", "", "Mốc ngày Finance ghi nhận doanh thu của kênh"),
    },
    "dim_promotion": {
        "promo_code": (REQ, "Văn bản", "", "Mã khuyến mãi"), "promo_name": (REQ, "Văn bản", "", "Tên chương trình"),
        "discount_pct": (REQ, "Tỷ lệ", "%", "% giảm trên giá niêm yết"),
        "window_from_mmdd": (REQ, "Văn bản", "", "Ngày bắt đầu hằng năm (mm-dd); trống = không theo mùa"),
        "window_to_mmdd": (REQ, "Văn bản", "", "Ngày kết thúc hằng năm (mm-dd)"),
        "channel": (REQ, "Văn bản", "", "Kênh áp dụng"),
    },
    "dim_order_status": {
        "order_status": (REQ, "Văn bản", "", "Mã trạng thái"), "status_name": (REQ, "Văn bản", "", "Tên trạng thái"),
        "count_in_revenue": (REQ, "Văn bản", "", "Có / Không tính vào doanh thu"), "rule": (REQ, "Văn bản", "", "Diễn giải quy tắc"),
    },
    "map_store_finance": {
        "fin_store_code": (REQ, "Văn bản", "", "Mã cửa hàng trong file Finance"),
        "store_key": (REQ, "Số nguyên", "", "store_key tương ứng; trống với HQ-ADJ"),
        "fin_channel": (REQ, "Văn bản", "", "Kênh theo cách gọi của Finance"),
        "valid_from": (REQ, "Ngày", "", "Mã có hiệu lực từ ngày"), "valid_to": (REQ, "Ngày", "", "Mã có hiệu lực đến ngày"),
        "note": (REQ, "Văn bản", "", "Ghi chú (đổi mã, bút toán HQ)"),
    },
    "fx_rate": {
        "date": (OPT, "Ngày", "", "Ngày tỷ giá"), "currency_code": (OPT, "Văn bản", "", "Mã tiền tệ"),
        "rate_to_usd": (OPT, "Số thập phân", "đơn vị tiền / 1 USD", "Tỷ giá so với USD"),
    },
    "finance_revenue.xlsx": {
        "posting_date": (REQ, "Ngày", "", "Ngày Finance hạch toán"),
        "fin_store_code": (REQ, "Văn bản", "", "Mã cửa hàng Finance → map_store_finance"),
        "fin_channel": (REQ, "Văn bản", "", "Retail / E-commerce / Head office"),
        "entry_type": (REQ, "Văn bản", "", "SALES hoặc ADJUSTMENT"),
        "revenue_recognized": (REQ, "Số thập phân", "USD", "Doanh thu sau giảm giá được ghi nhận"),
        "refunds": (REQ, "Số thập phân", "USD", "Tiền hoàn cho khách"),
        "net_revenue": (REQ, "Số thập phân", "USD", "revenue_recognized − refunds; con số đem đối soát"),
        "vat_amount": (REQ, "Số thập phân", "USD", "VAT 10 % × net_revenue"),
        "total_incl_vat": (REQ, "Số thập phân", "USD", "net_revenue + vat_amount; KHÔNG dùng để đối soát"),
    },
}
_tables = {"fact_sales_line": fact_sales_line, "fact_return": fact_return, "dim_date": dim_date, "dim_store": dim_store,
           "dim_product": dim_product, "dim_customer": dim_customer, "dim_channel": dim_channel,
           "dim_promotion": dim_promotion, "dim_order_status": dim_order_status, "map_store_finance": map_store_finance,
           "fx_rate": fx, "finance_revenue.xlsx": finance}
dict_rows = []
for tname, df in _tables.items():
    assert set(df.columns) == set(COLINFO[tname]), (tname, set(df.columns) ^ set(COLINFO[tname]))
    for col in df.columns:
        grp, typ, unit, mean = COLINFO[tname][col]
        ex = df[col].dropna()
        ex = "" if ex.empty else str(ex.iloc[0])[:40]
        dict_rows.append([tname, col, grp, typ, unit, mean, ex])
Data_Dictionary = pd.DataFrame(dict_rows, columns=["Bảng", "Trường", "Nhóm trường", "Kiểu dữ liệu", "Đơn vị",
                                                   "Ý nghĩa", "Ví dụ"])


# ---------------------------------------------------------------- 11. ghi file Excel
MONEY = {"unit_price", "discount_amount", "unit_cost", "return_amount", "list_price_usd", "unit_cost_usd",
         "revenue_recognized", "refunds", "net_revenue", "vat_amount", "total_incl_vat", "effect_usd"}
DOC_SHEETS = {"README", "Data_Dictionary", "QC_Summary", "KPI_Definitions"}


def write_xlsx(path, sheets):
    """Sheet tài liệu: bảng chữ có xuống dòng. Sheet dữ liệu: Excel Table tbl_<tên> + định dạng số theo cột."""
    with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="yyyy-mm-dd",
                        date_format="yyyy-mm-dd") as xw:
        wb = xw.book
        wb.set_properties({"title": path.stem, "author": "EDA Mentor",
                           "created": pd.Timestamp("2026-10-10").to_pydatetime()})
        F = {"money": wb.add_format({"num_format": "#,##0.00"}), "pct": wb.add_format({"num_format": "0%"}),
             "rate": wb.add_format({"num_format": "#,##0.0000"}), "int": wb.add_format({"num_format": "0"}),
             "wrap": wb.add_format({"text_wrap": True, "valign": "top"}),
             "hdr": wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#3A3A3A", "valign": "top",
                                   "text_wrap": True})}
        for name, df in sheets.items():
            df = df.reset_index(drop=True)
            df.to_excel(xw, sheet_name=name, index=False, header=False, startrow=1)
            ws = xw.sheets[name]
            n, m = len(df), len(df.columns)
            ws.add_table(0, 0, max(n, 1), m - 1, {
                "name": "tbl_" + "".join(ch if ch.isalnum() else "_" for ch in name),
                "columns": [{"header": str(c)} for c in df.columns],
                "style": "Table Style Light 1" if name in DOC_SHEETS else "Table Style Light 15"})
            for j, col in enumerate(df.columns):
                w = int(df[col].astype(str).str.len().quantile(.9)) + 2 if n else 10
                w = min(60 if name not in DOC_SHEETS else 90, max(10, len(str(col)) + 3, w))
                is_money = col in MONEY or (df[col].dtype.kind == "f" and col not in ("square_meters", "rate_to_usd")
                                            and not str(col).endswith("pct"))
                fmt = (F["wrap"] if name in DOC_SHEETS else F["pct"] if str(col).endswith("pct") else
                       F["rate"] if col == "rate_to_usd" else F["money"] if is_money else None)
                ws.set_column(j, j, w, fmt)
            ws.freeze_panes(1, 0)


dt_cols_conv = lambda df: df.assign(**{k: pd.to_datetime(df[k]) for k in df.columns
                                       if k.endswith("date") or k in ("valid_from", "valid_to")})
dataset_sheets = {
    "README": pd.DataFrame(README, columns=["Mục", "Nội dung"]),
    "Data_Dictionary": Data_Dictionary,
    "QC_Summary": QC,
    "KPI_Definitions": kpi,
    "dim_date": dt_cols_conv(dim_date.drop(columns=["month_end_date"])).assign(month_end_date=pd.to_datetime(dim_date.month_end_date)),
    "dim_store": dt_cols_conv(dim_store),
    "dim_product": dim_product,
    "dim_customer": dim_customer,
    "dim_channel": dim_channel,
    "dim_promotion": dim_promotion,
    "dim_order_status": dim_order_status,
    "fact_sales_line": dt_cols_conv(fact_sales_line),
    "fact_return": dt_cols_conv(fact_return),
    "map_store_finance": dt_cols_conv(map_store_finance),
    "sample_transactions": dt_cols_conv(sample_transactions),
    "fx_rate": dt_cols_conv(fx),
}
write_xlsx(OUT_DATA, dataset_sheets)
write_xlsx(OUT_FIN, {
    "README": pd.DataFrame([
        ("Nội dung", "Doanh thu đã được Finance xác nhận (bản rút gọn). USD; chưa gồm VAT trừ cột total_incl_vat."),
        ("Hạt dữ liệu", "Mỗi dòng = 1 ngày hạch toán × 1 mã cửa hàng của Finance. entry_type = SALES hoặc ADJUSTMENT."),
        ("revenue_recognized", "Doanh thu sau giảm giá được ghi nhận: POS theo ngày bán, online theo ngày giao."),
        ("refunds", "Tiền hoàn cho khách, hạch toán theo ngày Finance xử lý."),
        ("net_revenue", "= revenue_recognized − refunds. Đây là cột dùng để đối soát."),
        ("VAT", "vat_amount = 10 % × net_revenue; total_incl_vat = net_revenue + vat_amount."),
        ("Mã cửa hàng", "fin_store_code KHÔNG trùng store_key: dùng sheet map_store_finance của file dataset, "
                        "map theo valid_from / valid_to."),
        ("Power Query", "Dữ liệu nằm trong Excel Table tbl_finance_revenue."),
    ], columns=["Mục", "Nội dung"]),
    "finance_revenue": dt_cols_conv(finance),
})

# ---------------------------------------------------------------- 12. Instructor Key + số tham chiếu
anom_out = anom.assign(effect_usd=anom.effect_c / 100).drop(columns="effect_c")
anom_out = anom_out.merge(dim_store[["store_key", "store_code", "country"]], on="store_key", how="left")
TYPE_VI = {
    "A_CANCELLED_BOOKED": "Đơn huỷ vẫn được Finance ghi nhận",
    "B_MISSING_UPLOAD": "Finance thiếu dữ liệu 3 ngày",
    "C_VAT_INCLUDED": "Ghi doanh thu gồm VAT",
    "D_DUPLICATE_POSTING": "Bút toán ghi trùng",
    "E_MANUAL_ADJ": "Điều chỉnh HQ không thuộc cửa hàng",
}
systemic = pd.DataFrame([
    ["S1_ONLINE_TIMING", "Online: dữ liệu đơn theo order_date, Finance theo delivery_date",
     "Hệ thống, mọi tháng", br.adj_online_timing.abs().sum() / 100,
     "Cộng dồn cả kỳ ≈ 0, chỉ lệch ở ranh giới tháng; biến mất khi so theo delivery_date"],
    ["S2_REFUND_LAG", "Hoàn tiền: Returns theo return_date, Finance theo ngày hạch toán (trễ 0–3 ngày)",
     "Hệ thống, mọi tháng", br.adj_refund_lag.abs().sum() / 100,
     "Lệch theo ngày; cuối tháng tạo chênh lệch nhỏ"],
    ["S3_DUPLICATE_LINES", "120 dòng trùng trong fact_sales_line",
     "Phía dữ liệu đơn", dup.net_c.sum() / 100,
     "Không bỏ trùng → Net Sales phía đơn cao hơn đúng số này"],
    ["S4_PENDING", "Đơn online PENDING tại as-of",
     "Phía dữ liệu đơn", s[s.order_status == "PENDING"].net_c.sum() / 100,
     "Tính nhầm PENDING → cao hơn Finance đúng số này"],
    ["S5_STORE_RECODE", f"Cửa hàng {RECODE_STORE} đổi mã Finance {old_code} → {new_code} từ 01/07/2019",
     "Mapping", 0.0, "Không map theo hiệu lực → doanh thu từ 07/2019 thành 'không xác định cửa hàng'"],
], columns=["case_id", "mô tả", "loại", "độ lớn (USD, tổng trị tuyệt đối)", "cách phát hiện / ghi chú"])
br_out = br.copy()
money_cols = [k for k in br_out.columns if k.endswith(("_date", "_net", "revenue", "finance", "unexplained",
                                                       "orders")) or k.startswith("adj_") or k == "diff_finance_minus_orders"]
for k in money_cols:
    if br_out[k].dtype.kind in "iu":
        br_out[k] = br_out[k] / 100
month_tot = br_out.groupby("month_id")[[k for k in br_out.columns if br_out[k].dtype.kind == "f"
                                         and k not in ("diff_pct",)]].sum().reset_index()
write_xlsx(OUT_KEY, {
    "README": pd.DataFrame({"README": [
        "CHỈ DÀNH CHO MENTOR — không gửi học viên.",
        "Sheet planted_anomalies: từng chênh lệch cài sẵn phía Finance (K01…), cửa hàng, ngày, ảnh hưởng USD (+ = Finance cao hơn).",
        "Sheet systemic_causes: nguyên nhân hệ thống (timing online, trễ hoàn tiền) và bẫy phía dữ liệu đơn.",
        "Sheet bridge_store_month: cầu nối Orders Net Revenue → Finance Net Revenue cho từng tháng × cửa hàng; unexplained luôn = 0.",
        "Sheet bridge_month: như trên nhưng cộng theo tháng (toàn chuỗi, chưa gồm HQ-ADJ).",
        "Sheet over_threshold: các ô tháng × cửa hàng vượt ngưỡng — đây là danh sách học viên phải giải thích.",
        f"Ngưỡng: |Δ| > {THRESH_ABS:.0f} USD và |Δ %| > {THRESH_PCT:.0%}. Seed {SEED}. As-of {AS_OF:%d/%m/%Y}.",
    ]}),
    "planted_anomalies": anom_out.assign(type_vi=anom_out.type.map(TYPE_VI)),
    "systemic_causes": systemic,
    "bridge_store_month": br_out,
    "bridge_month": month_tot,
    "over_threshold": br_out[br_out.over_threshold],
})

# số tham chiếu (dashboard, guideline, slide đều đọc từ đây)
clean = s[s.order_status == "COMPLETED"].copy()
clean["month_id"] = clean.order_date.dt.year * 100 + clean.order_date.dt.month


def kpis(mask_lines, mask_ret):
    L, R = clean[mask_lines], ret[mask_ret]
    gross, disc = L.gross_c.sum() / 100, L.disc_c.sum() / 100
    ns = gross - disc
    rt = R.ret_c.sum() / 100
    o, u_ = L.order_number.nunique(), int(L.quantity.sum())
    return dict(gross_sales=round(gross, 2), discounts=round(disc, 2), net_sales=round(ns, 2),
                returns=round(rt, 2), net_revenue=round(ns - rt, 2), orders=int(o), units=u_,
                aov=round(ns / o, 2) if o else None, asp=round(ns / u_, 2) if u_ else None,
                return_rate=round(rt / ns, 4) if ns else None)


def month_mask(mid):
    return clean.month_id == mid, ret.ret_month == mid


def ytd_mask(mid):
    y = mid // 100
    return (clean.month_id >= y * 100 + 1) & (clean.month_id <= mid), (ret.ret_month >= y * 100 + 1) & (ret.ret_month <= mid)


rm, lm, sply = REPORT_MONTH, 202011, 201912
fin_store = fin[fin.store_key >= 0]
ref = {
    "version": VERSION, "seed": SEED, "as_of": str(AS_OF.date()), "report_month": rm,
    "rows": {"fact_sales_line": len(fact_sales_line), "fact_sales_line_clean": len(fs_clean),
             "fact_return": len(fact_return), "finance_revenue": len(finance), "orders_total": int(orders.shape[0]),
             "orders_completed": int((orders.order_status == "COMPLETED").sum()),
             "orders_cancelled": int((orders.order_status == "CANCELLED").sum()),
             "orders_pending": int((orders.order_status == "PENDING").sum()),
             "dim_store": len(dim_store), "dim_product": len(dim_product), "dim_customer": len(dim_customer),
             "dim_date": len(dim_date), "map_store_finance": len(map_store_finance)},
    "month": {k: kpis(*month_mask(v)) for k, v in [("current", rm), ("lm", lm), ("sply", sply)]},
    "ytd": {k: kpis(*ytd_mask(v)) for k, v in [("current", rm), ("sply", sply)]},
    "year": {str(y): kpis(clean.order_date.dt.year == y, ret.return_date.dt.year == y) for y in range(2016, 2022)},
    "finance_year": {str(y): round(fin[fin.post_date.dt.year == y].net_c.sum() / 100, 2) for y in range(2016, 2022)},
    "finance_month": {str(k): round(v / 100, 2) for k, v in fin.groupby("month_id").net_c.sum().items()},
    "bridge_2020": {k: round(br[br.month_id // 100 == 2020][k].sum() / 100, 2)
                    for k in ["orders_net_revenue"] + adj_cols + ["finance_net"]},
    "over_threshold_cells": int(br.over_threshold.sum()),
    "over_threshold_2020": br[br.over_threshold & (br.month_id // 100 == 2020)][
        ["month_id", "store_code", "country", "diff_finance_minus_orders"]].assign(
        diff_finance_minus_orders=lambda d: d.diff_finance_minus_orders / 100).to_dict("records"),
    "anomalies": anom_out[["anomaly_id", "type", "store_key", "store_code", "effect_usd"]].assign(
        date_from=anom_out.date_from.dt.strftime("%Y-%m-%d"), date_to=anom_out.date_to.dt.strftime("%Y-%m-%d")
    ).to_dict("records"),
    "systemic": systemic.iloc[:, [0, 3]].set_index("case_id").iloc[:, 0].round(2).to_dict(),
    "recode": {"store_key": RECODE_STORE, "old": old_code, "new": new_code},
    "threshold": {"abs_usd": THRESH_ABS, "pct": THRESH_PCT}, "vat_rate": VAT_RATE,
    "duplicates": 120,
}
OUT_REF.write_text(json.dumps(ref, ensure_ascii=False, indent=1, default=lambda o: None if pd.isna(o) else str(o)),
                   encoding="utf-8")

# dấu vân tay dữ liệu để kiểm chạy lại cho cùng kết quả
h = hashlib.sha256()
for df in (fact_sales_line, fact_return, finance):
    h.update(pd.util.hash_pandas_object(df, index=False).values.tobytes())
print("fingerprint", h.hexdigest()[:16])
print("rows", ref["rows"])
print("QC", (QC["Trạng thái"] == "PASS").sum(), "/", len(QC), "PASS")
print("anomalies", len(anom), "over-threshold cells", ref["over_threshold_cells"])
