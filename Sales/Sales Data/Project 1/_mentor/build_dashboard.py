# -*- coding: utf-8 -*-
"""Dựng ../Nexa_Sales_P1_Dashboard.html — bản tham chiếu tự chứa (d3 nhúng, dữ liệu tổng hợp sẵn).

Đọc lại CHÍNH các file giao học viên (dataset + finance) nên số trên dashboard là một phép tính
độc lập với generator. Khung CSS và thư viện d3 lấy nguyên từ dashboard Bank Project 1.
Chạy sau generate_dataset.py:  python build_dashboard.py
"""
import json
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
P1 = HERE.parent
BANK = P1.parents[2] / "Bank" / "Project 1" / "AIV_Bank_KHDN_Dashboard.html"
OUT = P1 / "Nexa_Sales_P1_Dashboard.html"
REF = json.loads((HERE / "ref_numbers.json").read_text(encoding="utf-8"))

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
st_cnt = s.groupby(["month_id", "store_key"]).order_status.value_counts().unstack(fill_value=0)
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


def rnd(df, cols, n=2):
    df = df.copy()
    for k in cols:
        df[k] = df[k].round(n)
    return df


money = ["gross", "disc", "ns", "ns_post", "ret", "fin"]
data = {
    "stores": {int(x.store_key): {"code": x.store_code, "name": x.store_name, "ch": x.channel_code,
                                  "ctry": COUNTRY_VI[x.country]} for x in st.itertuples()},
    "ms": rnd(ms, money).astype({k: int for k in ["month_id", "store_key", "units", "orders", "canc", "pend",
                                                  "dup"]}).to_dict("records"),
    "cat": rnd(cat.assign(category=cat.category.map(CAT_VI)), ["ns", "ret"]).astype(
        {"units": int}).to_dict("records"),
    "rr": rnd(rrs.assign(return_reason=rrs.return_reason.map(REASON)), ["ret"]).to_dict("records"),
    "hq": {str(k): round(v, 2) for k, v in hq.items()},
    "asof": REF["as_of"], "def_month": REF["report_month"],
    "thr": REF["threshold"],
}
# đối chiếu với số của generator (mốc tính độc lập)
cur = ms[ms.month_id == REF["report_month"]]
assert abs(cur.ns.sum() - cur.ret.sum() - REF["month"]["current"]["net_revenue"]) < 0.05
assert int(cur.orders.sum()) == REF["month"]["current"]["orders"]

payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))

# ------------------------------------------------------------ ghép HTML
bank = BANK.read_text(encoding="utf-8").split("\n")
css_bank = "\n".join(bank[3:86])           # dòng 4–86: <style> … quy tắc chung (bỏ </style>)
d3_block = "\n".join(bank[147:150])        # dòng 148–150: <script> d3 v7.9.0 </script>
assert css_bank.startswith("<style>") and "d3js.org v7.9.0" in d3_block

CSS_EXTRA = """
#s1{grid-template-columns:repeat(6,minmax(0,1fr));grid-template-rows:14.5cqw 31cqw}
#s1 .w1{grid-row:2;grid-column:span 2}
#s2{grid-template-columns:2fr 2fr 4fr 3fr;grid-template-rows:14.5cqw 29cqw 30cqw}
#s2 .kA{grid-row:1;grid-column:1}#s2 .kB{grid-row:1;grid-column:2}#s2 .tk{grid-row:1;grid-column:3/5}
#s2 .pie{grid-row:2;grid-column:1/3}#s2 .tm{grid-row:2;grid-column:3}#s2 .bar{grid-row:2;grid-column:4}
#s2 .tb{grid-row:3;grid-column:1/-1}
#s3,#s4{grid-template-columns:repeat(4,minmax(0,1fr));grid-template-rows:14.5cqw 32cqw 24cqw}
#s3 .l1,#s4 .l1{grid-row:3;grid-column:1/3}#s3 .l2,#s4 .l2{grid-row:3;grid-column:3/5}
#s3 .mid{grid-template-columns:3.6fr 5.8fr}
#s3tb table.t td,#s3tb table.t th{padding-left:.3cqw;padding-right:.3cqw}
table.t td.flag{color:#fff;background:var(--cor)!important;font-weight:700;text-align:center}
.k .vl.neg{color:var(--cor)}
</style>"""

HTML_BODY = """
<div class="aq" id="aq">
 <div class="aq-canvas" id="aq-canvas">
  <header class="aq-head">
   <div class="aq-ttl"><b>Kết quả bán hàng và đối soát doanh thu</b><span id="aq-sub">Nexa Home · Sales Project 1</span></div>
   <div class="aq-sl"><span>Tháng báo cáo</span><select id="aq-month" aria-label="Chọn tháng báo cáo"></select></div>
   <div class="aq-sl"><span>Kênh bán</span><select id="aq-ch" aria-label="Chọn kênh bán"><option value="ALL">Tất cả</option><option value="POS">Cửa hàng (POS)</option><option value="ONLINE">Website/app</option></select></div>
   <div class="aq-sl"><span>Quốc gia</span><select id="aq-ctry" aria-label="Chọn quốc gia"></select></div>
   <div class="aq-sl"><span aria-hidden="true">&nbsp;</span><button class="aq-rs" id="aq-reset" type="button">Đặt lại</button></div>
   <div class="aq-dt"><span>Dữ liệu đến</span><b id="aq-date">—</b></div>
  </header>

  <section class="aq-page on" id="s1" data-title="Tổng quan">
   <div class="k" id="s1k1"></div><div class="k" id="s1k2"></div><div class="k" id="s1k3"></div>
   <div class="k" id="s1k4"></div><div class="k" id="s1k5"></div><div class="k" id="s1k6"></div>
   <figure class="c w1"><h3>Net Revenue: dữ liệu đơn và Finance</h3><div class="b" id="s1c1"><div class="bx"></div></div><figcaption id="cap1"></figcaption></figure>
   <figure class="c w1"><h3>Số đơn hoàn tất theo tháng</h3><div class="b" id="s1c2"><div class="bx"></div></div><figcaption id="cap2"></figcaption></figure>
   <figure class="c w1"><h3>Từ Gross Sales đến Net Revenue</h3><div class="b" id="s1c3"><div class="bx"></div></div><figcaption id="cap3"></figcaption></figure>
  </section>

  <section class="aq-page" id="s2" data-title="Doanh thu">
   <div class="k kA" id="s2k1"></div><div class="k kB" id="s2k2"></div>
   <figure class="c tk"><h3>Doanh thu theo kênh</h3><div class="b tbl" id="s2tk"></div></figure>
   <figure class="c pie"><h3>Cơ cấu Net Sales theo danh mục</h3><div class="b" id="s2pie"><div class="bx"></div></div><figcaption id="cap4"></figcaption></figure>
   <figure class="c tm"><h3>Net Sales theo quốc gia</h3><div class="b" id="s2tm"><div class="bx"></div></div><figcaption id="cap5"></figcaption></figure>
   <figure class="c bar"><h3>AOV theo kênh</h3><div class="b" id="s2bar"><div class="bx"></div></div><figcaption id="cap6"></figcaption></figure>
   <figure class="c tb"><h3>Top 10 cửa hàng theo Net Sales</h3><div class="b tbl" id="s2tb"></div><figcaption id="cap7"></figcaption></figure>
  </section>

  <section class="aq-page" id="s3" data-title="Đối soát">
   <div class="k" id="s3k1"></div><div class="k" id="s3k2"></div><div class="k" id="s3k3"></div><div class="k" id="s3k4"></div>
   <div class="mid"><figure class="c"><h3>Chênh lệch Finance − dữ liệu đơn theo tháng</h3><div class="b" id="s3c1"><div class="bx"></div></div><figcaption id="cap8"></figcaption></figure>
   <figure class="c tb"><h3>Đối soát theo cửa hàng, xếp theo |Δ|</h3><div class="b tbl" id="s3tb"></div><figcaption id="cap9"></figcaption></figure></div>
   <figure class="c l1"><h3>Δ kênh cửa hàng (POS)</h3><div class="b" id="s3l1"><div class="bx"></div></div><figcaption id="cap10"></figcaption></figure>
   <figure class="c l2"><h3>Δ kênh online, trước và sau quy về ngày giao</h3><div class="b" id="s3l2"><div class="bx"></div></div><figcaption id="cap11"></figcaption></figure>
  </section>

  <section class="aq-page" id="s4" data-title="Hoàn trả & chất lượng">
   <div class="k" id="s4k1"></div><div class="k" id="s4k2"></div><div class="k" id="s4k3"></div><div class="k" id="s4k4"></div>
   <div class="mid"><figure class="c"><h3>Giá trị hoàn trả theo lý do</h3><div class="b" id="s4pie"><div class="bx"></div></div><figcaption id="cap12"></figcaption></figure>
   <figure class="c tb"><h3>Tỷ lệ hoàn trả theo danh mục</h3><div class="b tbl" id="s4tb"></div><figcaption id="cap13"></figcaption></figure></div>
   <figure class="c l1"><h3>Tỷ lệ hoàn trả theo tháng</h3><div class="b" id="s4l1"><div class="bx"></div></div><figcaption id="cap14"></figcaption></figure>
   <figure class="c l2"><h3>Đơn huỷ và đơn đang giao theo tháng</h3><div class="b" id="s4l2"><div class="bx"></div></div><figcaption id="cap15"></figcaption></figure>
  </section>
  <div class="aq-tip" id="aq-tip"></div>
 </div>
 <nav class="aq-tabs" id="aq-tabs" aria-label="Trang báo cáo">
  <button type="button" data-p="s1" class="on">Tổng quan</button><button type="button" data-p="s2">Doanh thu</button><button type="button" data-p="s3">Đối soát</button><button type="button" data-p="s4">Hoàn trả &amp; chất lượng</button>
 </nav>
</div>
</div>
"""

JS = r"""
<script>
(() => {
const D = window.__NX_DATA;
const C = {g:'#447A2A',g90:'#57873F',g80:'#699555',g70:'#7CA26A',g50:'#A2BD95',g35:'#BED0B4',g30:'#C7D7BF',cor:'#993C1D',blue:'#185FA5',purple:'#534AB7',teal:'#0F6E56',amber:'#854F0B',red:'#A32D2D',gray:'#5F5E5A',grid:'#DDDFD9',axis:'#A6A6A6',tx:'#262626'};
const PAL_PIE = ['#2B4F18','#5E9B3A','#A9CF8A','#3E8E6A','#86B85F','#E8D48A','#EBDDB0','#CFE3B6'];
const PAL_TM = ['#2B4F18','#A9CF8A','#447A2A','#E8D48A','#5E9B3A','#CFE3B6','#3E8E6A','#EBDDB0','#86B85F'];
const lum = c => { const o = d3.rgb(c); const f = v => { v /= 255; return v <= .03928 ? v/12.92 : Math.pow((v+.055)/1.055,2.4); }; return .2126*f(o.r)+.7152*f(o.g)+.0722*f(o.b); };
const inkOn = c => lum(c) < .2 ? '#fff' : C.tx;
const nf = (v,d=0) => { if (Math.abs(v) < 0.5*Math.pow(10,-d)) v = 0; return new Intl.NumberFormat('vi-VN',{minimumFractionDigits:d,maximumFractionDigits:d}).format(v).replace('-','−'); };
const $ = (s,r=document) => r.querySelector(s);
const cv = $('#aq-canvas'), tip = $('#aq-tip');
const mAdd = (m,n) => { const t = Math.floor(m/100)*12 + (m%100-1) + n; return Math.floor(t/12)*100 + t%12 + 1; };
const mStr = m => String(m%100).padStart(2,'0') + '/' + Math.floor(m/100);
const mShort = m => String(m%100).padStart(2,'0') + '/' + String(Math.floor(m/100)).slice(2);
const LM = m => mAdd(m,-1), SP = m => mAdd(m,-12);
const DEF = {mid:D.def_month, ch:'ALL', ctry:'ALL'};
let S = {...DEF}, page = 's1';

const ST = D.stores;
const byM = new Map(); for (const r of D.ms) { if (!byM.has(r.month_id)) byM.set(r.month_id,[]); byM.get(r.month_id).push(r); }
const MONTHS = [...byM.keys()].filter(m => (byM.get(m)||[]).some(r => r.orders > 0)).sort((a,b) => a-b);
const CTRY = [...new Set(Object.values(ST).filter(s => s.ch === 'POS').map(s => s.ctry))].sort((a,b) => a.localeCompare(b,'vi'));
const keep = k => { const s = ST[k]; return (S.ch === 'ALL' || s.ch === S.ch) && (S.ctry === 'ALL' || s.ctry === S.ctry || (s.ch === 'ONLINE' && S.ctry === 'ALL')); };
const rowsM = m => (byM.get(m) || []).filter(r => keep(r.store_key));
const sumM = (m,f) => { if (!byM.has(m)) return null; let s = 0; for (const r of rowsM(m)) s += f(r); return s; };
const gross = m => sumM(m,r => r.gross), disc = m => sumM(m,r => r.disc), ns = m => sumM(m,r => r.ns), ret = m => sumM(m,r => r.ret);
const nrev = m => { const a = ns(m), b = ret(m); return a == null ? null : a-b; };
const orders = m => sumM(m,r => r.orders), units = m => sumM(m,r => r.units);
const aov = m => { const o = orders(m); return o ? ns(m)/o : null; }, asp = m => { const u = units(m); return u ? ns(m)/u : null; };
const finv = m => sumM(m,r => r.fin);
const delta = m => { const a = finv(m), b = nrev(m); return a == null ? null : a-b; };
const deltaPct = m => { const a = finv(m), d = delta(m); return a ? d/Math.abs(a) : null; };
const deltaAdj = m => sumM(m,r => r.fin - (r.ns_post - r.ret));
const rrate = m => { const a = ns(m); return a ? ret(m)/a : null; };
const over = r => { const d = r.fin - (r.ns - r.ret), p = r.fin ? d/Math.abs(r.fin) : (d ? 1 : 0); return Math.abs(d) > D.thr.abs_usd && Math.abs(p) > D.thr.pct; };
const flags = m => rowsM(m).filter(over).length;

const dPct = (a,b) => (a == null || b == null || !b) ? null : (a-b)/Math.abs(b);
const dPP = (a,b) => (a == null || b == null) ? null : (a-b)*100;
const arrow = x => x > 0 ? '▲' : x < 0 ? '▼' : '■';
function dtxt(x,kind) { if (x == null) return '–'; return arrow(x) + ' ' + (kind === 'pp' ? nf(Math.abs(x),2) + ' điểm %' : nf(Math.abs(x*100),1) + '%'); }
function dcls(x,goodUp) { if (x == null || x === 0 || goodUp == null) return 'mut'; return ((x > 0) === goodUp) ? 'good' : 'bad'; }
function kpi(el, lab, val, unit, d1, d2, kind, goodUp, l1, l2, neg) {
  el.replaceChildren();
  const mk = (cls,txt) => { const e = document.createElement('div'); e.className = cls; e.textContent = txt; return e; };
  el.append(mk('lb',lab));
  const v = mk('vl' + (neg ? ' neg' : ''), val == null ? '–' : val); if (unit) { const s = document.createElement('small'); s.textContent = unit; v.append(s); } el.append(v);
  for (const [x,l] of [[d1,l1],[d2,l2]]) { if (l == null) continue; const e = mk('dl ' + dcls(x,goodUp), (typeof x === 'string' ? x : dtxt(x,kind)) + ' '); const i = document.createElement('i'); i.textContent = l; e.append(i); el.append(e); }
}
const L_LM = 'so với tháng trước', L_SP = 'so với cùng tháng năm trước';
function cardAmt(el, lab, f, unit, scale, d, goodUp=true) {
  const m = S.mid, v = f(m), a = f(LM(m)), b = f(SP(m));
  kpi(el, lab, v == null ? null : nf(v/scale,d), unit, dPct(v,a), dPct(v,b), 'pct', goodUp, L_LM, L_SP);
}
function cardRatio(el, lab, f, goodUp) {
  const m = S.mid, v = f(m);
  kpi(el, lab, v == null ? null : nf(v*100,2), '%', dPP(v,f(LM(m))), dPP(v,f(SP(m))), 'pp', goodUp, L_LM, L_SP);
}

function showTip(ev, lines) {
  const r = cv.getBoundingClientRect(); tip.replaceChildren();
  lines.forEach((t,i) => { const d = document.createElement(i === 0 ? 'b' : 'div'); d.textContent = t; tip.append(d); });
  tip.style.display = 'block';
  let x = ev.clientX - r.left + 12, y = ev.clientY - r.top + 12;
  const w = tip.offsetWidth, h = tip.offsetHeight;
  if (x + w > r.width - 4) x = ev.clientX - r.left - w - 12;
  if (y + h > r.height - 4) y = ev.clientY - r.top - h - 12;
  tip.style.left = Math.max(2,x) + 'px'; tip.style.top = Math.max(2,y) + 'px';
}
const hideTip = () => { tip.style.display = 'none'; };
function stage(el) {
  const host = el.querySelector('.bx') || el; host.replaceChildren();
  const W = host.clientWidth, H = host.clientHeight; if (W < 30 || H < 30) return null;
  const svg = d3.select(host).append('svg').attr('width','100%').attr('height','100%').attr('viewBox','0 0 ' + W + ' ' + H);
  const fs = parseFloat(getComputedStyle(el).fontSize) || 12;
  return {svg, W, H, fs, host};
}
const trunc = (s,n) => s.length > n ? s.slice(0,Math.max(1,n-1)) + '…' : s;

function lineChart(el, cats, series, o) {
  const st = stage(el); if (!st) return; const {svg,W,H,fs} = st;
  const all = series.flatMap(s => s.vals.filter(v => v != null)); if (!all.length) return;
  let lo = d3.min(all), hi = d3.max(all);
  if (o.zero) { lo = Math.min(0,lo); hi = Math.max(0,hi); const pad = (hi-lo)*0.18 || 1; hi += pad; if (lo < 0) lo -= pad; }
  else { const pad = (hi-lo)*0.2 || Math.abs(hi)*0.05 || 1; lo = lo-pad; hi = hi+pad; if (o.pos) lo = Math.max(0,lo); }
  const labW = series.length > 1 ? Math.max(...series.map(s => s.name.length))*fs*0.62 + 8 : 0;
  const m = {t:fs*1.8, r:Math.max(fs*3.4,labW+10), b:fs*2.1, l:fs*5};
  const x = d3.scalePoint().domain(cats.map((_,i) => i)).range([m.l,W-m.r]).padding(0.3);
  const y = d3.scaleLinear().domain([lo,hi]).nice(4).range([H-m.b,m.t]);
  const g = svg.append('g');
  y.ticks(4).forEach(t => { g.append('line').attr('x1',m.l).attr('x2',W-m.r).attr('y1',y(t)).attr('y2',y(t)).attr('stroke',t === 0 && lo < 0 ? C.axis : C.grid).attr('stroke-width',.9);
    g.append('text').attr('x',m.l-6).attr('y',y(t)+fs*0.35).attr('text-anchor','end').attr('fill',C.gray).attr('font-size',fs).text(o.fmt(t)); });
  const step = Math.ceil(cats.length / Math.max(2,Math.floor((W-m.l-m.r)/(fs*4.4))));
  cats.forEach((c,i) => { const lastI = cats.length-1; if ((i % step === 0 && (lastI - i >= step*0.7 || i === lastI)) || i === lastI) g.append('text').attr('x',x(i)).attr('y',H-m.b+fs*1.4).attr('text-anchor','middle').attr('fill',C.gray).attr('font-size',fs).text(c); });
  g.append('line').attr('x1',m.l).attr('x2',W-m.r).attr('y1',H-m.b).attr('y2',H-m.b).attr('stroke',C.axis);
  const sel = cats.indexOf(o.sel), labs = [];
  series.forEach(s => {
    const pts = s.vals.map((v,i) => v == null ? null : [x(i),y(v)]);
    const ln = d3.line().defined(d => d).x(d => d[0]).y(d => d[1]);
    g.append('path').attr('d',ln(pts)).attr('fill','none').attr('stroke',s.color).attr('stroke-width',2.6).attr('stroke-dasharray',s.dash || null).attr('stroke-linecap','round').attr('stroke-linejoin','round');
    pts.forEach((p,i) => { if (p) g.append('circle').attr('cx',p[0]).attr('cy',p[1]).attr('r',i === sel ? 5 : 2.4).attr('fill',s.color).attr('stroke','#fff').attr('stroke-width',i === sel ? 1.6 : .9); });
    if (sel >= 0 && s.vals[sel] != null && !s.nolab) {
      const cv0 = s.vals[sel], nb = [s.vals[sel-1], s.vals[sel+1]].filter(v => v != null);
      const nbAvg = nb.length ? d3.mean(nb) : cv0, py = y(cv0);
      let below = nbAvg > cv0; if (below && py + fs*1.9 > H-m.b) below = false; if (!below && py - fs*1.6 < 0) below = true;
      g.append('text').attr('x',x(sel)).attr('y',below ? py + fs*1.7 : py - fs*0.95).attr('text-anchor','middle').attr('font-weight',700).attr('font-size',fs).attr('fill',C.tx)
        .attr('stroke','#fff').attr('stroke-width',3.5).attr('stroke-linejoin','round').attr('paint-order','stroke').text(o.fmt(cv0));
    }
    if (series.length > 1) { const lastI = s.vals.map(v => v != null).lastIndexOf(true); if (lastI >= 0) labs.push({x:x(lastI)+8, y:y(s.vals[lastI])+fs*0.35, s}); }
  });
  labs.sort((a,b) => a.y-b.y); for (let i = 1; i < labs.length; i++) if (labs[i].y - labs[i-1].y < fs*1.15) labs[i].y = labs[i-1].y + fs*1.15;
  const ovf = labs.length ? labs[labs.length-1].y - (H-m.b) : 0; if (ovf > 0) labs.forEach(l => l.y -= ovf);
  labs.forEach(l => g.append('text').attr('x',l.x).attr('y',l.y).attr('font-size',fs).attr('font-weight',600).attr('fill',l.s.color).attr('stroke','#fff').attr('stroke-width',3).attr('paint-order','stroke').text(l.s.name));
  if (sel >= 0) g.append('line').attr('x1',x(sel)).attr('x2',x(sel)).attr('y1',m.t).attr('y2',H-m.b).attr('stroke',C.axis).attr('stroke-dasharray','3 3');
  svg.append('rect').attr('x',m.l).attr('y',0).attr('width',W-m.l-m.r).attr('height',H-m.b).attr('fill','transparent')
    .on('pointermove',ev => { const px = ev.clientX - svg.node().getBoundingClientRect().left; let bi = 0, bd = 1e9; cats.forEach((_,i) => { const dd = Math.abs(x(i)-px); if (dd < bd) { bd = dd; bi = i; } });
      showTip(ev,[cats[bi],...series.map(s => s.name + ': ' + (s.vals[bi] == null ? '–' : o.fmt(s.vals[bi]) + (o.unit ? ' ' + o.unit : '')))]); })
    .on('pointerleave',hideTip);
}
function pieChart(el, items, unit, pal) {
  const st = stage(el); if (!st) return; const {svg,W,H,fs} = st;
  const tot = d3.sum(items, d => d.v); if (!tot) return;
  const two = W < 420, lw = two ? Math.min(fs*6.2,W*0.28) : fs*11.5, R = Math.max(12, Math.min((W-2*lw)/2, H/2-fs*2));
  const g = svg.append('g').attr('transform','translate(' + W/2 + ',' + H/2 + ')');
  const arcs = d3.pie().value(d => d.v).sort(null)(items);
  g.selectAll('path').data(arcs).join('path').attr('d',d3.arc().innerRadius(0).outerRadius(R)).attr('fill',(d,i) => pal[i % pal.length]).attr('stroke','#fff').attr('stroke-width',1.5)
    .on('pointermove',(ev,d) => showTip(ev,[d.data.n, nf(d.data.v,0) + ' ' + unit + ' (' + nf(d.data.v/tot*100,1) + '%)'])).on('pointerleave',hideTip);
  arcs.forEach(d => { if (d.endAngle - d.startAngle < 0.12) return; const a = (d.startAngle+d.endAngle)/2, r1 = R+4, r2 = R+fs*1.0; const px = Math.sin(a)*r1, py = -Math.cos(a)*r1, qx = Math.sin(a)*r2, qy = -Math.cos(a)*r2;
    g.append('line').attr('x1',px).attr('y1',py).attr('x2',qx).attr('y2',qy).attr('stroke',C.axis);
    const right = Math.sin(a) >= 0, ax = qx + (right ? 3 : -3);
    const t = g.append('text').attr('x',ax).attr('y',qy+fs*0.35).attr('text-anchor',right ? 'start' : 'end').attr('font-size',fs).attr('fill',C.tx);
    t.append('tspan').attr('font-weight',700).text(nf(d.data.v/tot*100,1) + '%');
    if (two) t.append('tspan').attr('x',ax).attr('dy',fs*1.15).attr('fill',C.gray).text(trunc(d.data.n,11)); else t.append('tspan').attr('fill',C.gray).text(' ' + trunc(d.data.n,14)); });
}
function treemapChart(el, items, unit, pal) {
  const st = stage(el); if (!st) return; const {svg,W,H,fs} = st;
  const its = items.filter(d => d.v > 0).sort((a,b) => b.v-a.v); if (!its.length) return;
  const colOf = new Map(its.map((d,i) => [d.n,pal[i % pal.length]]));
  const root = d3.hierarchy({children:its}).sum(d => d.v).sort((a,b) => b.value-a.value);
  d3.treemap().size([W,H]).paddingInner(2).round(true)(root);
  const tot = d3.sum(its, d => d.v);
  const gs = svg.selectAll('g').data(root.leaves()).join('g').attr('transform',d => 'translate(' + d.x0 + ',' + d.y0 + ')');
  gs.append('rect').attr('width',d => d.x1-d.x0).attr('height',d => d.y1-d.y0).attr('rx',2).attr('fill',d => colOf.get(d.data.n)).attr('stroke','#fff')
    .on('pointermove',(ev,d) => showTip(ev,[d.data.n, nf(d.data.v,0) + ' ' + unit + ' (' + nf(d.data.v/tot*100,1) + '%)'])).on('pointerleave',hideTip);
  gs.each(function(d) { const w = d.x1-d.x0, h = d.y1-d.y0; if (w < fs*4.6 || h < fs*2.5) return;
    const t = d3.select(this).append('text').attr('fill',inkOn(colOf.get(d.data.n))).attr('font-size',fs).attr('pointer-events','none');
    const perLine = Math.max(4,Math.floor((w-10)/(fs*0.56)));
    t.append('tspan').attr('x',6).attr('y',fs*1.35).attr('font-weight',600).text(trunc(d.data.n,perLine));
    if (h > fs*3.8) t.append('tspan').attr('x',6).attr('dy',fs*1.25).text(nf(d.data.v/tot*100,1) + '%'); });
}
function barChart(el, items, unit, pal, fmt) {
  const st = stage(el); if (!st) return; const {svg,W,H,fs} = st;
  fmt = fmt || (v => nf(v,0));
  const m = {t:fs*2, r:8, b:fs*3.2, l:8};
  const lo = Math.min(0, d3.min(items, d => d.v)), hi = Math.max(0, d3.max(items, d => d.v)) || 1;
  const x = d3.scaleBand().domain(items.map(d => d.n)).range([m.l,W-m.r]).padding(0.3);
  const y = d3.scaleLinear().domain([lo*1.14,hi*1.14]).range([H-m.b,m.t]);
  const g = svg.append('g'); g.append('line').attr('x1',m.l).attr('x2',W-m.r).attr('y1',y(0)).attr('y2',y(0)).attr('stroke',C.axis);
  items.forEach((d,i) => { const top = Math.min(y(d.v),y(0)), h = Math.abs(y(d.v)-y(0));
    g.append('rect').attr('x',x(d.n)).attr('y',top).attr('width',x.bandwidth()).attr('height',Math.max(1,h)).attr('rx',2).attr('fill',d.c || pal[i % pal.length])
      .on('pointermove',ev => showTip(ev,[d.n, fmt(d.v) + ' ' + unit])).on('pointerleave',hideTip);
    g.append('text').attr('x',x(d.n)+x.bandwidth()/2).attr('y',d.v >= 0 ? y(d.v)-6 : y(0)-6).attr('text-anchor','middle').attr('font-size',fs).attr('font-weight',700).attr('fill',d.v < 0 ? C.cor : C.tx).text(fmt(d.v));
    const per = Math.floor(x.step()/(fs*0.58)), lab = g.append('text').attr('text-anchor','middle').attr('font-size',fs).attr('fill',C.gray);
    if (d.n.length > per && d.n.includes(' ')) { const sp = d.n.indexOf(' '); lab.append('tspan').attr('x',x(d.n)+x.bandwidth()/2).attr('y',H-m.b+fs*1.3).text(trunc(d.n.slice(0,sp),per)); lab.append('tspan').attr('x',x(d.n)+x.bandwidth()/2).attr('dy',fs*1.1).text(trunc(d.n.slice(sp+1),per)); }
    else lab.attr('x',x(d.n)+x.bandwidth()/2).attr('y',H-m.b+fs*1.4).text(trunc(d.n,per)); });
}
function table(el, cols, rows, groups) {
  el.replaceChildren(); const t = document.createElement('table'); t.className = 't';
  const cg = document.createElement('colgroup'); cols.forEach(c => { const e = document.createElement('col'); e.style.width = c.w; cg.append(e); }); t.append(cg);
  const thd = document.createElement('thead');
  if (groups) { const gr = document.createElement('tr'); groups.forEach(gp => { const e = document.createElement('th'); e.textContent = gp.h; e.colSpan = gp.span; if (gp.h) e.className = 'gc'; gr.append(e); }); thd.append(gr); }
  const hr = document.createElement('tr'); cols.forEach(c => { const e = document.createElement('th'); e.textContent = c.h; if (c.l) e.style.textAlign = 'left'; hr.append(e); }); thd.append(hr); t.append(thd);
  const tb = document.createElement('tbody');
  rows.forEach(r => { const tr = document.createElement('tr'); cols.forEach(c => { const td = document.createElement('td'); const o = c.f(r); td.textContent = o.t; if (o.c) td.className = o.c; if (c.l) td.style.textAlign = 'left'; tr.append(td); }); tb.append(tr); });
  t.append(tb); el.append(t);
}
const cap = (id,t) => { $('#' + id).textContent = t; };
const dc = x => ({t:dtxt(x,'pct'),c:x == null ? 'mut' : (x > 0 ? 'good' : x < 0 ? 'bad' : 'mut')});
const win = () => MONTHS.filter(m => m <= S.mid && m > mAdd(S.mid,-24));
const k$ = v => nf(v/1000,1);   // nghìn USD
const scopeTxt = () => (S.ch === 'ALL' ? 'mọi kênh' : S.ch === 'POS' ? 'kênh cửa hàng' : 'kênh online') + (S.ctry === 'ALL' ? '' : ', ' + S.ctry);

function pageHome() {
  cardAmt($('#s1k1'),'Gross Sales',gross,'nghìn USD',1000,1);
  cardAmt($('#s1k2'),'Net Revenue',nrev,'nghìn USD',1000,1);
  cardAmt($('#s1k3'),'Số đơn hoàn tất',orders,'đơn',1,0);
  cardAmt($('#s1k4'),'AOV',aov,'USD/đơn',1,0);
  cardAmt($('#s1k5'),'ASP',asp,'USD/sản phẩm',1,1);
  const m = S.mid, dv = delta(m), dp = deltaPct(m);
  kpi($('#s1k6'),'Chênh lệch Finance', dv == null ? null : nf(dv/1000,1), 'nghìn USD', dp == null ? '–' : nf(dp*100,2) + '%', flags(m) + ' ô', null, null, 'so với Finance', 'cửa hàng vượt ngưỡng', Math.abs(dv||0) > D.thr.abs_usd);
  const R = win(), cats = R.map(mShort), sel = mShort(m);
  lineChart($('#s1c1'),cats,[{name:'Dữ liệu đơn',color:C.g,vals:R.map(x => nrev(x)/1000)},{name:'Finance',color:C.cor,dash:'5 4',vals:R.map(x => finv(x)/1000),nolab:true,dy:1}],{fmt:v => nf(v,0),unit:'nghìn USD',sel,pos:true});
  lineChart($('#s1c2'),cats,[{name:'Số đơn',color:C.blue,vals:R.map(orders)}],{fmt:v => nf(v,0),unit:'đơn',sel,pos:true});
  const its = [{n:'Gross Sales',v:gross(m)},{n:'Giảm giá',v:-disc(m),c:C.cor},{n:'Net Sales',v:ns(m)},{n:'Hoàn trả',v:-ret(m),c:C.cor},{n:'Net Revenue',v:nrev(m),c:'#2B4F18'}].map(d => ({...d,v:d.v/1000}));
  barChart($('#s1c3'),its,'nghìn USD',['#86B85F'],v => nf(v,1));
  cap('cap1','Hình 1: Net Revenue theo tháng, ' + mStr(R[0]) + '–' + mStr(R[R.length-1]) + ' (nghìn USD), ' + scopeTxt());
  cap('cap2','Hình 2: Số đơn COMPLETED theo tháng đặt hàng');
  cap('cap3','Hình 3: Gross Sales → Net Revenue tháng ' + mStr(m) + ' (nghìn USD)');
}

function pageRevenue() {
  cardAmt($('#s2k1'),'Net Sales',ns,'nghìn USD',1000,1);
  cardAmt($('#s2k2'),'Số sản phẩm bán',units,'sản phẩm',1,0);
  const m = S.mid;
  const chv = (mm,ch,f) => { let s = 0; for (const r of (byM.get(mm)||[])) if (ST[r.store_key].ch === ch && keep(r.store_key)) s += f(r); return s; };
  const rows = [['POS','Cửa hàng'],['ONLINE','Website/app']].filter(c => S.ch === 'ALL' || S.ch === c[0]).map(([k,n]) => {
    const c = chv(m,k,r => r.ns), l = chv(LM(m),k,r => r.ns), s_ = chv(SP(m),k,r => r.ns), o = chv(m,k,r => r.orders);
    return {n, c, l, s:s_, o, aov:o ? c/o : null}; });
  const totc = d3.sum(rows, d => d.c);
  table($('#s2tk'),[{h:'Kênh',w:'17%',l:1,f:r => ({t:r.n})},{h:'Net Sales (nghìn USD)',w:'16%',f:r => ({t:k$(r.c)})},{h:'Tỷ trọng',w:'10%',f:r => ({t:totc ? nf(r.c/totc*100,1) + '%' : '–'})},{h:'Số đơn',w:'10%',f:r => ({t:nf(r.o)})},{h:'Tháng trước',w:'13%',f:r => ({t:k$(r.l)})},{h:'Thay đổi',w:'10%',f:r => dc(dPct(r.c,r.l))},{h:'Cùng kỳ năm trước',w:'14%',f:r => ({t:k$(r.s)})},{h:'Thay đổi',w:'10%',f:r => dc(dPct(r.c,r.s))}],rows);
  const catM = new Map(); for (const r of D.cat) if (r.month_id === m && keep(r.store_key)) catM.set(r.category,(catM.get(r.category)||0)+r.ns);
  pieChart($('#s2pie'),[...catM.entries()].sort((a,b) => b[1]-a[1]).map(e => ({n:e[0],v:e[1]})),'USD',PAL_PIE);
  const cty = new Map(); for (const r of rowsM(m)) { const n = ST[r.store_key].ctry; cty.set(n,(cty.get(n)||0)+r.ns); }
  treemapChart($('#s2tm'),[...cty.entries()].map(e => ({n:e[0],v:e[1]})),'USD',PAL_TM);
  barChart($('#s2bar'),rows.map((r,i) => ({n:r.n,v:r.aov||0,c:i ? '#86B85F' : '#2B4F18'})),'USD/đơn',['#2B4F18']);
  const agg = (mm) => { const o = new Map(); for (const r of rowsM(mm)) o.set(r.store_key,(o.get(r.store_key)||0)+r.ns); return o; };
  const c = agg(m), l = agg(LM(m)), s_ = agg(SP(m)), oc = new Map(rowsM(m).map(r => [r.store_key,r.orders]));
  const top = [...c.entries()].filter(e => e[1] > 0).sort((a,b) => b[1]-a[1]).slice(0,10).map(([k,v],i) => ({i:i+1,n:ST[k].name,ct:ST[k].ctry,c:v,o:oc.get(k)||0,l:l.get(k),s:s_.get(k)}));
  table($('#s2tb'),[{h:'#',w:'4%',l:1,f:r => ({t:String(r.i)})},{h:'Cửa hàng',w:'26%',l:1,f:r => ({t:r.n})},{h:'Quốc gia',w:'9%',l:1,f:r => ({t:r.ct})},{h:'Net Sales (nghìn USD)',w:'12%',f:r => ({t:k$(r.c)})},{h:'Số đơn',w:'8%',f:r => ({t:nf(r.o)})},{h:'Tháng trước',w:'11%',f:r => ({t:r.l == null ? '–' : k$(r.l)})},{h:'Thay đổi',w:'9%',f:r => dc(dPct(r.c,r.l))},{h:'Cùng kỳ năm trước',w:'12%',f:r => ({t:r.s == null ? '–' : k$(r.s)})},{h:'Thay đổi',w:'9%',f:r => dc(dPct(r.c,r.s))}],top);
  cap('cap4','Hình 4: Cơ cấu Net Sales theo danh mục, tháng ' + mStr(m));
  cap('cap5','Hình 5: Net Sales theo quốc gia, tháng ' + mStr(m));
  cap('cap6','Hình 6: Giá trị trung bình mỗi đơn theo kênh (USD)');
  cap('cap7','Hình 7: Top 10 cửa hàng theo Net Sales tháng ' + mStr(m) + ' (nghìn USD)');
}

function pageRecon() {
  const m = S.mid;
  cardAmt($('#s3k1'),'Net Revenue (dữ liệu đơn)',nrev,'nghìn USD',1000,1);
  cardAmt($('#s3k2'),'Net Revenue (Finance)',finv,'nghìn USD',1000,1);
  const dv = delta(m), dp = deltaPct(m), da = deltaAdj(m);
  kpi($('#s3k3'),'Chênh lệch Finance − đơn', dv == null ? null : nf(dv,0), 'USD', dp == null ? '–' : nf(dp*100,2) + '%', da == null ? '–' : nf(da,0) + ' USD', null, null, 'trên Finance', 'sau khi quy online về ngày giao', Math.abs(dv||0) > D.thr.abs_usd);
  const fl = flags(m), n = rowsM(m).filter(r => r.fin || r.ns).length;
  kpi($('#s3k4'),'Cửa hàng vượt ngưỡng', String(fl), 'trên ' + n + ' cửa hàng có số liệu', 'Ngưỡng', '', null, null, '|Δ| > ' + nf(D.thr.abs_usd) + ' USD và |Δ%| > ' + nf(D.thr.pct*100) + '%', null, fl > 0);
  const R = win(), cats = R.map(mShort), sel = mShort(m);
  lineChart($('#s3c1'),cats,[{name:'Δ gốc',color:C.cor,vals:R.map(delta)},{name:'Δ sau quy đổi',color:C.g,vals:R.map(deltaAdj),nolab:true,dy:1}],{fmt:v => nf(v/1000,1),unit:'nghìn USD',sel,zero:true});
  const rows = rowsM(m).filter(r => r.fin || r.ns).map(r => { const nr = r.ns - r.ret, d = r.fin - nr; return {n:ST[r.store_key].code + ' · ' + ST[r.store_key].ctry, ch:ST[r.store_key].ch === 'POS' ? 'POS' : 'Online', nr, fin:r.fin, d, p:r.fin ? d/Math.abs(r.fin) : null, da:r.fin - (r.ns_post - r.ret), fl:over(r)}; })
    .sort((a,b) => Math.abs(b.d)-Math.abs(a.d)).slice(0,12);
  const pc = v => ({t:v == null ? '–' : nf(v*100,1) + '%'});
  table($('#s3tb'),[{h:'Cửa hàng',w:'17%',l:1,f:r => ({t:r.n})},{h:'Kênh',w:'8%',l:1,f:r => ({t:r.ch})},{h:'Đơn hàng',w:'13%',f:r => ({t:nf(r.nr,0)})},{h:'Finance',w:'13%',f:r => ({t:nf(r.fin,0)})},{h:'Δ',w:'12%',f:r => ({t:nf(r.d,0),c:Math.abs(r.d) > D.thr.abs_usd ? 'bad' : 'mut'})},{h:'Δ %',w:'9%',f:r => pc(r.p)},{h:'Δ sau quy đổi',w:'14%',f:r => ({t:nf(r.da,0)})},{h:'Cờ',w:'8%',f:r => r.fl ? {t:'Vượt',c:'flag'} : {t:'',c:'mut'}}],rows,[{h:'',span:2},{h:'Net Revenue (USD)',span:2},{h:'Chênh lệch (USD)',span:4}]);
  const chD = (ch,adj) => R.map(mm => { let s = 0; for (const r of (byM.get(mm)||[])) if (ST[r.store_key].ch === ch && keep(r.store_key)) s += r.fin - ((adj ? r.ns_post : r.ns) - r.ret); return s; });
  if (S.ch === 'ONLINE') { const h = $('#s3l1 .bx'); h.replaceChildren(); const e = document.createElement('div'); e.style.cssText = 'display:flex;height:100%;align-items:center;justify-content:center;color:#5F5E5A'; e.textContent = 'Kênh cửa hàng không thuộc bộ lọc hiện tại.'; h.append(e); }
  else lineChart($('#s3l1'),cats,[{name:'Δ POS',color:C.blue,vals:chD('POS',false)}],{fmt:v => nf(v,0),unit:'USD',sel,zero:true});
  if (S.ch === 'POS' || S.ctry !== 'ALL') { const h = $('#s3l2 .bx'); h.replaceChildren(); const e = document.createElement('div'); e.style.cssText = 'display:flex;height:100%;align-items:center;justify-content:center;color:#5F5E5A;padding:1em;text-align:center'; e.textContent = 'Kênh online không thuộc bộ lọc hiện tại (online không gắn với quốc gia nào).'; h.append(e); }
  else lineChart($('#s3l2'),cats,[{name:'Δ gốc',color:C.cor,vals:chD('ONLINE',false)},{name:'Δ sau quy đổi',color:C.g,vals:chD('ONLINE',true),nolab:true,dy:1}],{fmt:v => nf(v/1000,1),unit:'nghìn USD',sel,zero:true});
  const hq = D.hq[String(m)];
  cap('cap8','Hình 8: Δ = Finance − dữ liệu đơn (nghìn USD); "sau quy đổi" = online tính theo ngày giao');
  cap('cap9','Hình 9: 12 cửa hàng lệch nhiều nhất tháng ' + mStr(m) + (hq ? ' · chưa gồm HQ-ADJ ' + nf(hq,0) + ' USD' : ''));
  cap('cap10','Hình 10: Chênh lệch kênh cửa hàng theo tháng (USD)');
  cap('cap11','Hình 11: Chênh lệch kênh online: gốc và sau quy về ngày giao (nghìn USD)');
}

function pageQuality() {
  const m = S.mid;
  cardAmt($('#s4k1'),'Giá trị hoàn trả',ret,'nghìn USD',1000,1,false);
  cardRatio($('#s4k2'),'Tỷ lệ hoàn trả',rrate,false);
  const canc = mm => sumM(mm,r => r.canc), pend = mm => sumM(mm,r => r.pend);
  kpi($('#s4k3'),'Đơn huỷ', nf(canc(m)||0), 'đơn', dPct(canc(m),canc(LM(m))), dPct(canc(m),canc(SP(m))), 'pct', false, L_LM, L_SP);
  const dupAll = D.ms.filter(r => keep(r.store_key)).reduce((a,r) => a + r.dup, 0), pendAll = D.ms.filter(r => keep(r.store_key)).reduce((a,r) => a + r.pend, 0);
  kpi($('#s4k4'),'Dòng trùng đã loại', nf(dupAll), 'dòng, toàn bộ dữ liệu', nf(pendAll) + ' đơn', null, null, null, 'đang giao (PENDING) tại ngày trích', null);
  const rr = new Map(); for (const r of D.rr) if (r.month_id === m && keep(r.store_key)) rr.set(r.return_reason,(rr.get(r.return_reason)||0)+r.ret);
  pieChart($('#s4pie'),[...rr.entries()].sort((a,b) => b[1]-a[1]).map(e => ({n:e[0],v:e[1]})),'USD',PAL_PIE);
  const cm = (mm) => { const o = new Map(); for (const r of D.cat) if (r.month_id === mm && keep(r.store_key)) { const x = o.get(r.category) || {ns:0,ret:0}; x.ns += r.ns; x.ret += r.ret; o.set(r.category,x); } return o; };
  const c = cm(m), l = cm(LM(m)), s_ = cm(SP(m)), rt = x => x && x.ns ? x.ret/x.ns : null;
  const rows = [...c.entries()].sort((a,b) => b[1].ns-a[1].ns).map(([k,v]) => ({n:k,ns:v.ns,ret:v.ret,r:rt(v),l:rt(l.get(k)),s:rt(s_.get(k))}));
  const pv = v => ({t:v == null ? '–' : nf(v*100,2) + '%'});
  const chg = (a,b) => { const x = (a == null || b == null) ? null : (a-b)*100; return {t:x == null ? '–' : arrow(x) + ' ' + nf(Math.abs(x),2),c:x == null || x === 0 ? 'mut' : (x > 0 ? 'bad' : 'good')}; };
  table($('#s4tb'),[{h:'Danh mục',w:'22%',l:1,f:r => ({t:r.n})},{h:'Net Sales (nghìn USD)',w:'15%',f:r => ({t:k$(r.ns)})},{h:'Hoàn trả (USD)',w:'13%',f:r => ({t:nf(r.ret,0)})},{h:'Tỷ lệ',w:'10%',f:r => pv(r.r)},{h:'Tháng trước',w:'11%',f:r => pv(r.l)},{h:'Thay đổi (điểm %)',w:'12%',f:r => chg(r.r,r.l)},{h:'Cùng kỳ',w:'8%',f:r => pv(r.s)},{h:'Thay đổi',w:'9%',f:r => chg(r.r,r.s)}],rows);
  const R = win(), cats = R.map(mShort), sel = mShort(m);
  lineChart($('#s4l1'),cats,[{name:'Tỷ lệ hoàn trả',color:C.cor,vals:R.map(x => { const v = rrate(x); return v == null ? null : v*100; })}],{fmt:v => nf(v,2),unit:'%',sel,zero:true});
  lineChart($('#s4l2'),cats,[{name:'Đơn huỷ',color:C.cor,vals:R.map(canc)},{name:'Đang giao',color:C.blue,vals:R.map(pend),nolab:true,dy:1}],{fmt:v => nf(v,0),unit:'đơn',sel,zero:true});
  cap('cap12','Hình 12: Giá trị hoàn trả theo lý do, tháng ' + mStr(m) + ' (USD)');
  cap('cap13','Hình 13: Tỷ lệ hoàn trả = Returns / Net Sales theo danh mục, tháng ' + mStr(m));
  cap('cap14','Hình 14: Tỷ lệ hoàn trả theo tháng (%)');
  cap('cap15','Hình 15: Số đơn CANCELLED và PENDING theo tháng đặt hàng');
}

function drawHeader() {
  const sel = $('#aq-month'), R = MONTHS.slice().reverse();
  if (sel.options.length !== R.length) sel.replaceChildren(...R.map(m => new Option(mStr(m),String(m))));
  sel.value = String(S.mid);
  const cs = $('#aq-ctry');
  if (cs.options.length !== CTRY.length + 1) cs.replaceChildren(new Option('Tất cả','ALL'),...CTRY.map(c => new Option(c,c)));
  cs.value = S.ctry; $('#aq-ch').value = S.ch;
  $('#aq-date').textContent = D.asof.split('-').reverse().join('/');
  $('#aq-sub').textContent = 'Nexa Home · Sales Project 1 · Trang ' + $('#' + page).dataset.title;
}
function render() {
  cv.classList.toggle('narrow', cv.clientWidth < 820);
  if (S.ctry !== 'ALL' && S.ch === 'ONLINE') S.ch = 'ALL';
  drawHeader();
  ({s1:pageHome,s2:pageRevenue,s3:pageRecon,s4:pageQuality})[page]();
}
const set = o => { Object.assign(S,o); render(); };
$('#aq-month').addEventListener('change', e => set({mid:Number(e.target.value)}));
$('#aq-ch').addEventListener('change', e => set({ch:e.target.value, ctry:e.target.value === 'ONLINE' ? 'ALL' : S.ctry}));
$('#aq-ctry').addEventListener('change', e => set({ctry:e.target.value, ch:e.target.value !== 'ALL' && S.ch === 'ONLINE' ? 'POS' : S.ch}));
$('#aq-reset').addEventListener('click', () => set({...DEF}));
$('#aq-tabs').addEventListener('click', e => { const b = e.target.closest('button[data-p]'); if (b) go(b.dataset.p); });
function go(id) {
  page = id; document.querySelectorAll('.aq-page').forEach(s => s.classList.toggle('on', s.id === id));
  document.querySelectorAll('#aq-tabs button').forEach(b => b.classList.toggle('on', b.dataset.p === id));
  render();
}
const named = (location.hash || '').slice(1); if (/^s[1-4]$/.test(named)) page = named;
go(page);
let t; if (window.ResizeObserver) new ResizeObserver(() => { clearTimeout(t); t = setTimeout(render,60); }).observe(cv);
})();
</script>
"""

head = ('<!doctype html>\n<html lang="vi"><head><meta charset="utf-8"><meta name="viewport" '
        'content="width=device-width,initial-scale=1"><title>Kết quả bán hàng và đối soát doanh thu – Nexa Home</title>'
        '<style>html,body{margin:0;background:#fff}body{padding:8px}</style></head><body>\n<div id="dash-root">\n')
html = (head + css_bank + CSS_EXTRA + HTML_BODY + d3_block + "\n<script>window.__NX_DATA=" + payload
        + ";</script>\n" + JS + "</body></html>\n")
OUT.write_text(html, encoding="utf-8")
print(OUT.name, f"{OUT.stat().st_size/1e6:.2f} MB", "ms rows", len(data["ms"]), "cat rows", len(data["cat"]))
