# -*- coding: utf-8 -*-
"""Dựng ../Nexa_Sales_P1_Dashboard.html — dashboard tham chiếu tự chứa, thiết kế theo Storytelling with Data.

Nguyên tắc (Cole Nussbaumer Knaflic, "Storytelling with Data"):
  1. Bối cảnh   — khán giả: Sales Manager + Finance trước họp chốt số; mỗi trang một "Điểm chính" viết thành câu.
  2. Chọn hình  — đường cho xu hướng, cột ngang xếp hạng cho danh mục, bảng cho chi tiết; không pie/donut/treemap.
  3. Gỡ rối     — không khung nền màu, lưới mờ, nhãn trực tiếp, ít chữ số lẻ.
  4. Dẫn mắt    — xám cho bối cảnh, MỘT màu nhấn (xanh dương) cho điều cần thấy; đỏ chỉ cho "vượt ngưỡng" (kèm ⚠ + chữ).
  5. Thiết kế   — lưới thẳng hàng, định dạng số thống nhất, màu an toàn cho người mù màu, tooltip khi rê chuột.
  6. Kể chuyện  — tiêu đề hành động tự sinh theo số; chân trang ghi nguồn, as-of, quy ước Δ.
Số liệu: aggregate.load() (đọc lại file giao học viên, assert khớp ref_numbers.json).
Thư viện d3 v7.9.0 lấy nguyên từ dashboard Bank (chỉ thư viện, không dùng lại CSS/bố cục).
"""
import json

import aggregate as A

OUT = A.P1 / "Nexa_Sales_P1_Dashboard.html"
BANK = A.P1.parents[2] / "Bank" / "Project 1" / "AIV_Bank_KHDN_Dashboard.html"
REF = A.REF
D = A.load()
ms, cat, rrs, hq, st = D["ms"], D["cat"], D["rr"], D["hq"], D["st"]


def rnd(df, cols, n=2):
    df = df.copy()
    for k in cols:
        df[k] = df[k].round(n)
    return df


data = {
    "stores": {int(x.store_key): {"code": x.store_code, "name": x.store_name, "ch": x.channel_code,
                                  "ctry": D["COUNTRY_VI"][x.country]} for x in st.itertuples()},
    "ms": rnd(ms, ["gross", "disc", "ns", "ns_post", "ret", "fin"]).astype(
        {k: int for k in ["month_id", "store_key", "units", "orders", "canc", "pend", "dup"]}).to_dict("records"),
    "cat": rnd(cat.assign(category=cat.category.map(D["CAT_VI"])), ["ns", "ret"]).astype(
        {"units": int}).to_dict("records"),
    "rr": rnd(rrs.assign(return_reason=rrs.return_reason.map(D["REASON"])), ["ret"]).to_dict("records"),
    "hq": {str(k): round(v, 2) for k, v in hq.items()},
    "asof": REF["as_of"], "def_month": REF["report_month"], "thr": REF["threshold"],
}
payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))

bank = BANK.read_text(encoding="utf-8").split("\n")
d3_block = "\n".join(bank[147:150])        # <script> d3 v7.9.0 </script>
assert "d3js.org v7.9.0" in d3_block

CSS = r"""
<style>
:root{
  --surface:#ffffff; --panel:#ffffff; --bg:#f6f6f4;
  --ink:#1f1f1e; --ink2:#52514e; --muted:#8a8984; --rule:#e4e3df; --grid:#ecebe7;
  --accent:#2a78d6; --accent-soft:#cde2fb; --context:#b9b7b0; --context-dark:#8f8d86;
  --down:#eb6834; --crit:#d03b3b; --good:#1c5cab;
}
html,body{margin:0;background:var(--bg);color:var(--ink);
  font-family:"Segoe UI","Segoe UI Variable",system-ui,-apple-system,"Helvetica Neue",Arial,sans-serif;font-variant-numeric:tabular-nums}
*{box-sizing:border-box}
.wrap{max-width:1360px;margin:0 auto;padding:12px 16px 24px}
header.top{display:flex;flex-wrap:wrap;align-items:flex-end;gap:12px 20px;padding:4px 0 10px;border-bottom:1px solid var(--rule)}
.brand{flex:1 1 320px;min-width:0}
.brand h1{margin:0;font-size:22px;font-weight:600;letter-spacing:-.2px}
.brand p{margin:2px 0 0;color:var(--ink2);font-size:13px}
.ctl{display:flex;flex-direction:column;gap:3px;font-size:12px;color:var(--ink2)}
.ctl select,.ctl button{font:inherit;font-size:14px;color:var(--ink);background:#fff;border:1px solid #cfcdc7;border-radius:6px;padding:6px 10px;min-width:140px;height:34px}
.ctl button{min-width:0;cursor:pointer}
.ctl button:hover{border-color:var(--ink2)}
nav.tabs{display:flex;gap:4px;margin:12px 0 0;flex-wrap:wrap}
nav.tabs button{font:inherit;font-size:14px;background:transparent;border:0;border-bottom:3px solid transparent;color:var(--ink2);padding:8px 12px;cursor:pointer}
nav.tabs button.on{color:var(--ink);border-bottom-color:var(--accent);font-weight:600}
.page{display:none}.page.on{display:block}
.story{background:var(--panel);border-left:4px solid var(--accent);padding:12px 16px;margin:14px 0 12px;border-radius:0 8px 8px 0}
.story .lbl{font-size:12px;font-weight:600;color:var(--accent);text-transform:uppercase;letter-spacing:.6px}
.story .big{font-size:18px;line-height:1.45;margin:4px 0 2px;font-weight:600}
.story ul{margin:4px 0 0;padding-left:18px;color:var(--ink2);font-size:14px;line-height:1.5}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin-bottom:12px}
.k{background:var(--panel);border-radius:8px;padding:10px 14px}
.k .lb{font-size:13px;color:var(--ink2)}
.k .vl{font-size:28px;font-weight:600;line-height:1.15;margin-top:2px}
.k .vl small{font-size:13px;font-weight:400;color:var(--ink2);margin-left:4px}
.k .dl{font-size:12.5px;color:var(--ink2);margin-top:3px;line-height:1.35}
.k .dl b{font-weight:600}
.k .dl .up{color:var(--good)}.k .dl .down{color:var(--down)}
.k.alert .vl{color:var(--crit)}
.grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}
.card{background:var(--panel);border-radius:8px;padding:12px 14px 10px;min-width:0;display:flex;flex-direction:column}
.card.s2{grid-column:span 2}.card.s3{grid-column:1/-1}
.card h2{margin:0;font-size:15px;font-weight:600;line-height:1.35}
.card .sub{margin:2px 0 6px;font-size:12.5px;color:var(--muted)}
.card .plot{position:relative;flex:1;min-height:240px}
.card .plot.tall{min-height:300px}
.card .plot svg{display:block;width:100%;height:100%;position:absolute;inset:0}
table.t{width:100%;border-collapse:collapse;font-size:13px}
table.t th{font-weight:600;color:var(--ink2);text-align:right;padding:6px 8px;border-bottom:1px solid var(--rule);white-space:nowrap}
table.t td{padding:5px 8px;text-align:right;border-bottom:1px solid var(--grid);white-space:nowrap}
table.t th.l,table.t td.l{text-align:left}
table.t td.flag{color:var(--crit);font-weight:600}
table.t td.mut{color:var(--muted)}
table.t td .bar{display:inline-block;height:9px;background:var(--context);border-radius:0 3px 3px 0;vertical-align:middle;margin-right:6px}
table.t td .bar.hl{background:var(--accent)}
.foot{margin-top:14px;font-size:12px;color:var(--muted);line-height:1.5}
.tip{position:fixed;z-index:50;pointer-events:none;background:#fff;border:1px solid #cfcdc7;border-radius:6px;padding:6px 10px;font-size:13px;box-shadow:0 2px 8px rgba(0,0,0,.08);display:none;max-width:280px}
.tip b{display:block;margin-bottom:2px}
.empty{display:flex;height:100%;align-items:center;justify-content:center;color:var(--muted);font-size:13px;text-align:center;padding:12px}
@media (max-width:900px){.grid{grid-template-columns:1fr}.card.s2,.card.s3{grid-column:auto}.story .big{font-size:16px}}
</style>"""

BODY = r"""
<div class="wrap">
 <header class="top">
  <div class="brand"><h1>Kết quả bán hàng và đối soát doanh thu</h1><p id="sub">Nexa Home · Sales Project 1</p></div>
  <label class="ctl">Tháng báo cáo<select id="f-month"></select></label>
  <label class="ctl">Kênh bán<select id="f-ch"><option value="ALL">Tất cả kênh</option><option value="POS">Cửa hàng (POS)</option><option value="ONLINE">Website/app</option></select></label>
  <label class="ctl">Quốc gia<select id="f-ctry"></select></label>
  <div class="ctl"><span>&nbsp;</span><button id="f-reset" type="button">Đặt lại</button></div>
 </header>
 <nav class="tabs" id="tabs">
  <button data-p="s1" class="on">1 · Tổng quan</button><button data-p="s2">2 · Doanh thu đến từ đâu</button>
  <button data-p="s3">3 · Đối soát với Finance</button><button data-p="s4">4 · Hoàn trả &amp; chất lượng dữ liệu</button>
 </nav>

 <section class="page on" id="s1">
  <div class="story" id="s1-story"></div>
  <div class="kpis" id="s1-k"></div>
  <div class="grid">
   <div class="card s2"><h2 id="t1a"></h2><div class="sub">Net Revenue theo tháng đặt hàng, nghìn USD · tháng đang chọn tô xanh, cùng kỳ năm trước đánh dấu</div><div class="plot tall" id="c1a"></div></div>
   <div class="card"><h2 id="t1b"></h2><div class="sub">Tháng đang chọn, nghìn USD</div><div class="plot tall" id="c1b"></div></div>
   <div class="card"><h2 id="t1c"></h2><div class="sub">Số đơn hoàn tất theo tháng</div><div class="plot" id="c1c"></div></div>
   <div class="card"><h2 id="t1d"></h2><div class="sub">AOV = Net Sales / số đơn, USD</div><div class="plot" id="c1d"></div></div>
   <div class="card"><h2 id="t1e"></h2><div class="sub">Δ sau quy đổi online theo tháng, nghìn USD · đỏ = tháng có cửa hàng vượt ngưỡng · chi tiết ở trang 3</div><div class="plot" id="c1e"></div></div>
  </div>
 </section>

 <section class="page" id="s2">
  <div class="story" id="s2-story"></div>
  <div class="kpis" id="s2-k"></div>
  <div class="grid">
   <div class="card"><h2 id="t2a"></h2><div class="sub">Net Sales tháng đang chọn theo danh mục, nghìn USD</div><div class="plot tall" id="c2a"></div></div>
   <div class="card"><h2 id="t2b"></h2><div class="sub">Net Sales theo quốc gia, nghìn USD (online tính riêng)</div><div class="plot tall" id="c2b"></div></div>
   <div class="card"><h2 id="t2c"></h2><div class="sub">So với tháng trước (LM) và cùng kỳ năm trước (SPLY)</div><div id="c2c"></div></div>
   <div class="card s3"><h2 id="t2d"></h2><div class="sub">Top 10 cửa hàng theo Net Sales tháng đang chọn</div><div id="c2d"></div></div>
  </div>
 </section>

 <section class="page" id="s3">
  <div class="story" id="s3-story"></div>
  <div class="kpis" id="s3-k"></div>
  <div class="grid">
   <div class="card"><h2 id="t3a"></h2><div class="sub">Phân rã Δ tháng đang chọn, USD · không gồm HQ-ADJ</div><div class="plot tall" id="c3a"></div></div>
   <div class="card s2"><h2 id="t3b"></h2><div class="sub">Δ sau khi quy online về ngày giao, nghìn USD · tháng có cửa hàng vượt ngưỡng tô đỏ</div><div class="plot tall" id="c3b"></div></div>
   <div class="card s2"><h2 id="t3c"></h2><div class="sub">Xếp theo |Δ| giảm dần · ⚠ = |Δ| &gt; ngưỡng USD và |Δ %| &gt; ngưỡng %</div><div id="c3c"></div></div>
   <div class="card"><h2 id="t3d"></h2><div class="sub">Δ kênh online theo tháng, nghìn USD</div><div class="plot" id="c3d"></div></div>
  </div>
 </section>

 <section class="page" id="s4">
  <div class="story" id="s4-story"></div>
  <div class="kpis" id="s4-k"></div>
  <div class="grid">
   <div class="card s2"><h2 id="t4a"></h2><div class="sub">Returns / Net Sales theo tháng, %</div><div class="plot" id="c4a"></div></div>
   <div class="card"><h2 id="t4b"></h2><div class="sub">Giá trị hoàn trả tháng đang chọn theo lý do, USD</div><div class="plot" id="c4b"></div></div>
   <div class="card s2"><h2 id="t4c"></h2><div class="sub">Tỷ lệ hoàn trả theo danh mục tháng đang chọn · nhãn kèm Net Sales để đọc cùng quy mô</div><div class="plot tall" id="c4c"></div></div>
   <div class="card"><h2 id="t4d"></h2><div class="sub">Số đơn CANCELLED và PENDING theo tháng đặt hàng</div><div class="plot tall" id="c4d"></div></div>
  </div>
 </section>

 <div class="foot" id="foot"></div>
</div>
<div class="tip" id="tip"></div>
"""

JS = r"""
<script>
(() => {
const D = window.__NX_DATA;
const C = {accent:'#2a78d6', soft:'#9ec5f4', ctx:'#b9b7b0', ctxD:'#8f8d86', down:'#eb6834', crit:'#d03b3b', ink:'#1f1f1e', ink2:'#52514e', muted:'#8a8984', grid:'#ecebe7', rule:'#cfcdc7'};
const $ = s => document.querySelector(s);
const nf = (v,d=0) => { if (v == null || isNaN(v)) return '–'; if (Math.abs(v) < 0.5*Math.pow(10,-d)) v = 0; return new Intl.NumberFormat('vi-VN',{minimumFractionDigits:d,maximumFractionDigits:d}).format(v).replace('-','−'); };
const sg = (v,d=0) => (v >= 0.5*Math.pow(10,-d) ? '+' : '') + nf(v,d);
const k1 = v => nf(v/1000,1);                       // nghìn USD
const pc = (v,d=1) => v == null || !isFinite(v) ? '–' : nf(v*100,d) + '%';
const spc = (v,d=1) => v == null || !isFinite(v) ? '–' : sg(v*100,d) + '%';
const mAdd = (m,n) => { const t = Math.floor(m/100)*12 + (m%100-1) + n; return Math.floor(t/12)*100 + t%12 + 1; };
const mStr = m => String(m%100).padStart(2,'0') + '/' + Math.floor(m/100);
const mShort = m => String(m%100).padStart(2,'0') + '/' + String(Math.floor(m/100)).slice(2);
const LM = m => mAdd(m,-1), SP = m => mAdd(m,-12);
const DEF = {mid:D.def_month, ch:'ALL', ctry:'ALL'};
let S = {...DEF}, page = 's1';
const ST = D.stores;
const byM = new Map(); for (const r of D.ms) { if (!byM.has(r.month_id)) byM.set(r.month_id,[]); byM.get(r.month_id).push(r); }
const MONTHS = [...byM.keys()].filter(m => byM.get(m).some(r => r.orders > 0)).sort((a,b) => a-b);
const CTRY = [...new Set(Object.values(ST).filter(s => s.ch === 'POS').map(s => s.ctry))].sort((a,b) => a.localeCompare(b,'vi'));
const keep = k => { const s = ST[k]; return (S.ch === 'ALL' || s.ch === S.ch) && (S.ctry === 'ALL' || s.ctry === S.ctry); };
const rowsM = m => (byM.get(m) || []).filter(r => keep(r.store_key));
const sumM = (m,f) => { if (!byM.has(m)) return null; let s = 0; for (const r of rowsM(m)) s += f(r); return s; };
const gross = m => sumM(m,r => r.gross), disc = m => sumM(m,r => r.disc), ns = m => sumM(m,r => r.ns), ret = m => sumM(m,r => r.ret);
const nrev = m => { const a = ns(m); return a == null ? null : a - ret(m); };
const orders = m => sumM(m,r => r.orders), units = m => sumM(m,r => r.units);
const aov = m => { const o = orders(m); return o ? ns(m)/o : null; };
const finv = m => sumM(m,r => r.fin);
const delta = m => { const a = finv(m); return a == null ? null : a - nrev(m); };
const deltaAdj = m => sumM(m,r => r.fin - (r.ns_post - r.ret));
const rrate = m => { const a = ns(m); return a ? ret(m)/a : null; };
const dOf = r => r.fin - (r.ns - r.ret);
const over = r => { const d = dOf(r), p = r.fin ? d/Math.abs(r.fin) : (d ? 1 : 0); return Math.abs(d) > D.thr.abs_usd && Math.abs(p) > D.thr.pct; };
const overRows = m => rowsM(m).filter(over);
const chg = (a,b) => (a == null || b == null || !b) ? null : (a-b)/Math.abs(b);
const win = () => MONTHS.filter(m => m <= S.mid && m > mAdd(S.mid,-24));
const scope = () => (S.ch === 'ALL' ? 'mọi kênh' : S.ch === 'POS' ? 'kênh cửa hàng' : 'kênh online') + (S.ctry === 'ALL' ? '' : ' · ' + S.ctry);

// ---------------------------------------------------------------- tooltip
const tip = $('#tip');
const showTip = (ev, lines) => { tip.replaceChildren(); lines.forEach((t,i) => { const e = document.createElement(i ? 'div' : 'b'); e.textContent = t; tip.append(e); });
  tip.style.display = 'block'; const w = tip.offsetWidth, h = tip.offsetHeight;
  let x = ev.clientX + 14, y = ev.clientY + 14; if (x + w > innerWidth - 8) x = ev.clientX - w - 14; if (y + h > innerHeight - 8) y = ev.clientY - h - 14;
  tip.style.left = x + 'px'; tip.style.top = y + 'px'; };
const hideTip = () => { tip.style.display = 'none'; };

// ---------------------------------------------------------------- khung vẽ
function stage(id) { const host = $('#' + id); host.replaceChildren(); const W = host.clientWidth, H = host.clientHeight;
  if (W < 40 || H < 40) return null; const svg = d3.select(host).append('svg').attr('viewBox', `0 0 ${W} ${H}`); return {svg, W, H, fs:12.5}; }
function empty(id, msg) { const h = $('#' + id); h.replaceChildren(); const e = document.createElement('div'); e.className = 'empty'; e.textContent = msg; h.append(e); }

// Đường: xám cho bối cảnh, điểm tháng chọn tô nhấn, nhãn trực tiếp ở điểm chọn và điểm đánh dấu.
function line(id, cats, series, o) {
  const st = stage(id); if (!st) return; const {svg, W, H, fs} = st;
  const all = series.flatMap(s => s.vals).filter(v => v != null); if (!all.length) return empty(id, 'Không có dữ liệu');
  let lo = Math.min(o.zero ? 0 : d3.min(all), d3.min(all)), hi = Math.max(0, d3.max(all));
  if (!o.zero) { const p = (hi - lo) * .15 || 1; lo -= p; hi += p; if (o.pos) lo = Math.max(0, lo); } else { const p = (hi - lo) * .15 || 1; hi += p; if (lo < 0) lo -= p; }
  const m = {t:22, r:o.r || 70, b:24, l:46};
  const x = d3.scalePoint().domain(cats.map((_,i) => i)).range([m.l, W - m.r]).padding(.2);
  const y = d3.scaleLinear().domain([lo, hi]).nice(4).range([H - m.b, m.t]);
  const g = svg.append('g');
  y.ticks(4).forEach(t => { g.append('line').attr('x1', m.l).attr('x2', W - m.r).attr('y1', y(t)).attr('y2', y(t)).attr('stroke', t === 0 ? C.rule : C.grid);
    g.append('text').attr('x', m.l - 6).attr('y', y(t) + 4).attr('text-anchor', 'end').attr('font-size', fs - 1).attr('fill', C.muted).text(o.fmt(t)); });
  const step = Math.max(1, Math.ceil(cats.length / Math.max(2, Math.floor((W - m.l - m.r) / 52))));
  cats.forEach((c,i) => { const last = cats.length - 1; if ((i % step === 0 && last - i >= step * .8) || i === last) g.append('text').attr('x', x(i)).attr('y', H - 6).attr('text-anchor', 'middle').attr('font-size', fs - 1).attr('fill', C.muted).text(c); });
  const sel = cats.indexOf(o.sel), labs = [];
  series.forEach(s => {
    const pts = s.vals.map((v,i) => v == null ? null : [x(i), y(v)]);
    g.append('path').attr('d', d3.line().defined(d => d).x(d => d[0]).y(d => d[1])(pts)).attr('fill', 'none').attr('stroke', s.color).attr('stroke-width', s.w || 2).attr('stroke-dasharray', s.dash || null);
    if (s.main && sel >= 0 && s.vals[sel] != null) {
      g.append('circle').attr('cx', x(sel)).attr('cy', y(s.vals[sel])).attr('r', 5).attr('fill', C.accent).attr('stroke', '#fff').attr('stroke-width', 2);
      g.append('text').attr('x', x(sel)).attr('y', y(s.vals[sel]) - 10).attr('text-anchor', 'middle').attr('font-size', fs).attr('font-weight', 600).attr('fill', C.ink)
        .attr('stroke', '#fff').attr('stroke-width', 4).attr('paint-order', 'stroke').text(o.fmt(s.vals[sel]));
    }
    (s.marks || []).forEach(mk => { if (mk.i < 0 || s.vals[mk.i] == null) return;
      g.append('circle').attr('cx', x(mk.i)).attr('cy', y(s.vals[mk.i])).attr('r', 4).attr('fill', '#fff').attr('stroke', C.ctxD).attr('stroke-width', 2);
      g.append('text').attr('x', x(mk.i)).attr('y', y(s.vals[mk.i]) - 9).attr('text-anchor', 'middle').attr('font-size', fs - 1).attr('fill', C.ink2)
        .attr('stroke', '#fff').attr('stroke-width', 4).attr('paint-order', 'stroke').text(mk.t); });
    if (s.name) { const li = s.vals.map(v => v != null).lastIndexOf(true); if (li >= 0) labs.push({y: y(s.vals[li]) + 4, x: x(li) + 8, s}); }
  });
  labs.sort((a,b) => a.y - b.y); for (let i = 1; i < labs.length; i++) if (labs[i].y - labs[i-1].y < 14) labs[i].y = labs[i-1].y + 14;
  labs.forEach(l => g.append('text').attr('x', l.x).attr('y', l.y).attr('font-size', fs).attr('font-weight', 600).attr('fill', l.s.color === C.ctx ? C.ink2 : l.s.color).text(l.s.name));
  svg.append('rect').attr('x', m.l).attr('y', 0).attr('width', W - m.l - m.r).attr('height', H - m.b).attr('fill', 'transparent')
    .on('pointermove', ev => { const px = ev.offsetX * (W / svg.node().clientWidth); let bi = 0, bd = 1e9; cats.forEach((_,i) => { const dd = Math.abs(x(i) - px); if (dd < bd) { bd = dd; bi = i; } });
      showTip(ev, [cats[bi], ...series.map(s => (s.name || s.tipName || '') + ': ' + (s.vals[bi] == null ? '–' : o.fmt(s.vals[bi]) + (o.unit ? ' ' + o.unit : '')))]); })
    .on('pointerleave', hideTip);
}

// Cột ngang xếp hạng: một màu nhấn cho mục cần thấy, còn lại xám; nhãn giá trị ở cuối thanh.
function hbar(id, items, o) {
  const st = stage(id); if (!st) return; const {svg, W, H, fs} = st;
  if (!items.length) return empty(id, 'Không có dữ liệu');
  const labW = Math.min(W * .42, Math.max(...items.map(d => d.n.length)) * 7.6 + 16), valW = o.valW || 92;
  const bh = Math.min(26, (H - 8) / items.length);
  const x = d3.scaleLinear().domain([0, d3.max(items, d => d.v) || 1]).range([labW, W - valW]);
  items.forEach((d,i) => { const yy = 4 + i * bh, h = bh * .66;
    svg.append('text').attr('x', labW - 8).attr('y', yy + h/2 + 4).attr('text-anchor', 'end').attr('font-size', fs).attr('fill', d.hl ? C.ink : C.ink2).attr('font-weight', d.hl ? 600 : 400).text(d.n);
    svg.append('rect').attr('x', labW).attr('y', yy).attr('width', Math.max(1, x(d.v) - labW)).attr('height', h).attr('rx', 3).attr('fill', d.hl ? C.accent : (d.c || C.ctx))
      .on('pointermove', ev => showTip(ev, [d.n, ...(d.tip || [o.fmt(d.v)])])).on('pointerleave', hideTip);
    svg.append('text').attr('x', x(d.v) + 6).attr('y', yy + h/2 + 4).attr('font-size', fs).attr('fill', C.ink2).attr('font-weight', d.hl ? 600 : 400).text(d.lab || o.fmt(d.v)); });
}

// Cột theo tháng (có âm/dương): xám, cột nổi bật theo d.c; tháng chọn viền nhấn.
function colMonth(id, cats, vals, colors, o) {
  const st = stage(id); if (!st) return; const {svg, W, H, fs} = st;
  const vv = vals.filter(v => v != null); if (!vv.length) return empty(id, 'Không có dữ liệu');
  const m = {t:16, r:10, b:24, l:46};
  const x = d3.scaleBand().domain(cats.map((_,i) => i)).range([m.l, W - m.r]).padding(.25);
  const lo = Math.min(0, d3.min(vv)), hi = Math.max(0, d3.max(vv));
  const y = d3.scaleLinear().domain([lo, hi]).nice(4).range([H - m.b, m.t]);
  y.ticks(4).forEach(t => { svg.append('line').attr('x1', m.l).attr('x2', W - m.r).attr('y1', y(t)).attr('y2', y(t)).attr('stroke', t === 0 ? C.rule : C.grid);
    svg.append('text').attr('x', m.l - 6).attr('y', y(t) + 4).attr('text-anchor', 'end').attr('font-size', fs - 1).attr('fill', C.muted).text(o.fmt(t)); });
  const step = Math.max(1, Math.ceil(cats.length / Math.max(2, Math.floor((W - m.l - m.r) / 52))));
  vals.forEach((v,i) => { if (v == null) return; const yy = Math.min(y(v), y(0)), h = Math.max(1, Math.abs(y(v) - y(0)));
    svg.append('rect').attr('x', x(i)).attr('y', yy).attr('width', x.bandwidth()).attr('height', h).attr('rx', 2).attr('fill', colors[i])
      .attr('stroke', cats[i] === o.sel ? C.ink : 'none').attr('stroke-width', 1.5)
      .on('pointermove', ev => showTip(ev, [cats[i], o.fmt(v) + ' ' + (o.unit || ''), ...(o.extra ? o.extra(i) : [])])).on('pointerleave', hideTip);
    const last = cats.length - 1; if ((i % step === 0 && last - i >= step * .8) || i === last) svg.append('text').attr('x', x(i) + x.bandwidth()/2).attr('y', H - 6).attr('text-anchor', 'middle').attr('font-size', fs - 1).attr('fill', C.muted).text(cats[i]); });
}

// Waterfall: bước 'total' từ 0, bước delta nối tiếp; giảm = cam, tăng = xám đậm, tổng = nhấn.
function waterfall(id, steps, o) {
  const st = stage(id); if (!st) return; const {svg, W, H, fs} = st;
  const m = {t:22, r:8, b:40, l:8};
  let acc = 0; const bars = steps.map(s => { if (s.total) { acc = s.v; return {...s, a:0, b:s.v}; } const a = acc; acc += s.v; return {...s, a, b:acc}; });
  const lo = Math.min(0, d3.min(bars, d => Math.min(d.a, d.b))), hi = Math.max(0, d3.max(bars, d => Math.max(d.a, d.b)));
  const x = d3.scaleBand().domain(bars.map((_,i) => i)).range([m.l, W - m.r]).padding(.3);
  const y = d3.scaleLinear().domain([lo, hi]).nice().range([H - m.b, m.t]);
  svg.append('line').attr('x1', m.l).attr('x2', W - m.r).attr('y1', y(0)).attr('y2', y(0)).attr('stroke', C.rule);
  bars.forEach((d,i) => { const top = y(Math.max(d.a, d.b)), h = Math.max(1, Math.abs(y(d.a) - y(d.b)));
    const fill = d.c || (d.total ? (d.hl ? C.accent : C.ctxD) : d.v < 0 ? C.down : C.ctxD);
    svg.append('rect').attr('x', x(i)).attr('y', top).attr('width', x.bandwidth()).attr('height', h).attr('rx', 3).attr('fill', fill)
      .on('pointermove', ev => showTip(ev, [d.n, o.fmt(d.v) + ' ' + (o.unit || '')])).on('pointerleave', hideTip);
    if (i < bars.length - 1) svg.append('line').attr('x1', x(i) + x.bandwidth()).attr('x2', x(i + 1)).attr('y1', y(d.b)).attr('y2', y(d.b)).attr('stroke', C.ctx).attr('stroke-dasharray', '2 2');
    svg.append('text').attr('x', x(i) + x.bandwidth()/2).attr('y', top - 6).attr('text-anchor', 'middle').attr('font-size', fs).attr('font-weight', 600).attr('fill', d.v < 0 && !d.total && !d.c ? C.down : C.ink).text((d.total ? '' : (d.v > 0 ? '+' : '')) + o.fmt(d.v));
    const words = d.n.split(' '), half = Math.ceil(words.length / 2), lab = svg.append('text').attr('text-anchor', 'middle').attr('font-size', fs - .5).attr('fill', C.ink2);
    [words.slice(0, half).join(' '), words.slice(half).join(' ')].filter(Boolean).forEach((t,j) => lab.append('tspan').attr('x', x(i) + x.bandwidth()/2).attr('y', H - m.b + 15 + j * 13).text(t)); });
}

function table(id, cols, rows) {
  const host = $('#' + id); host.replaceChildren(); const t = document.createElement('table'); t.className = 't';
  const tr = document.createElement('tr'); cols.forEach(c => { const th = document.createElement('th'); th.textContent = c.h; if (c.l) th.className = 'l'; tr.append(th); });
  const th = document.createElement('thead'); th.append(tr); t.append(th); const tb = document.createElement('tbody');
  rows.forEach(r => { const row = document.createElement('tr'); cols.forEach(c => { const td = document.createElement('td'); const o = c.f(r);
    if (o.bar != null) { const b = document.createElement('span'); b.className = 'bar' + (o.hl ? ' hl' : ''); b.style.width = Math.max(2, o.bar * 120) + 'px'; td.append(b); }
    td.append(document.createTextNode(o.t)); if (o.c) td.className = o.c; if (c.l) td.classList.add('l'); row.append(td); }); tb.append(row); });
  t.append(tb); host.append(t);
}

// ---------------------------------------------------------------- thẻ KPI: giá trị + 2 so sánh, chữ ▲▼ kèm màu
function kpis(id, list) { const host = $('#' + id); host.replaceChildren();
  list.forEach(k => { const e = document.createElement('div'); e.className = 'k' + (k.alert ? ' alert' : '');
    const lb = document.createElement('div'); lb.className = 'lb'; lb.textContent = k.lb;
    const vl = document.createElement('div'); vl.className = 'vl'; vl.textContent = k.v; if (k.u) { const s = document.createElement('small'); s.textContent = k.u; vl.append(s); }
    e.append(lb, vl);
    (k.d || []).forEach(([x, lab, goodUp, kind]) => { const dl = document.createElement('div'); dl.className = 'dl';
      if (x == null || !isFinite(x)) { dl.textContent = '– ' + lab; }
      else { const b = document.createElement('b'); const good = goodUp == null ? null : (x > 0) === goodUp; if (good != null && x !== 0) b.className = good ? 'up' : 'down';
        b.textContent = (x > 0 ? '▲ ' : x < 0 ? '▼ ' : '■ ') + (kind === 'pp' ? nf(Math.abs(x)*100,2) + ' điểm %' : nf(Math.abs(x)*100,1) + '%'); dl.append(b, document.createTextNode(' ' + lab)); }
      e.append(dl); });
    (k.note || []).forEach(n => { const dl = document.createElement('div'); dl.className = 'dl'; dl.textContent = n; e.append(dl); });
    host.append(e); }); }

function story(id, big, bullets) { const h = $('#' + id); h.replaceChildren();
  const l = document.createElement('div'); l.className = 'lbl'; l.textContent = 'Điểm chính · ' + mStr(S.mid) + ' · ' + scope();
  const b = document.createElement('div'); b.className = 'big'; b.textContent = big; h.append(l, b);
  if (bullets.length) { const ul = document.createElement('ul'); bullets.forEach(t => { const li = document.createElement('li'); li.textContent = t; ul.append(li); }); h.append(ul); } }
const T = (id, t) => { $('#' + id).textContent = t; };

// ================================================================ Trang 1 · Tổng quan
function p1() {
  const m = S.mid, nr = nrev(m), o = orders(m), a = aov(m), d = delta(m), fl = overRows(m).length;
  const cO = chg(o, orders(SP(m))), cA = chg(a, aov(SP(m))), cN = chg(nr, nrev(SP(m)));
  const driver = (cO == null || cA == null) ? '' : Math.abs(Math.log(1 + cO)) >= Math.abs(Math.log(1 + cA))
    ? `Biến động so với cùng kỳ đến chủ yếu từ số đơn (${spc(cO)}); giá trị mỗi đơn (AOV) ${spc(cA)}.`
    : `Biến động so với cùng kỳ đến chủ yếu từ giá trị mỗi đơn (AOV ${spc(cA)}); số đơn ${spc(cO)}.`;
  story('s1-story', `Net Revenue ${mStr(m)} đạt ${k1(nr)} nghìn USD — ${spc(chg(nr, nrev(LM(m))))} so với tháng trước, ${spc(cN)} so với cùng kỳ năm trước.`,
    [driver, `Finance ghi nhận ${d >= 0 ? 'cao' : 'thấp'} hơn dữ liệu đơn ${nf(Math.abs(d))} USD (${spc(chg(finv(m), nr))}); ${fl} cửa hàng vượt ngưỡng cần giải thích (trang 3).`]);
  kpis('s1-k', [
    {lb:'Net Revenue', v:k1(nr), u:'nghìn USD', d:[[chg(nr, nrev(LM(m))), 'so với tháng trước', true], [cN, 'so với cùng kỳ', true]]},
    {lb:'Số đơn hoàn tất', v:nf(o), u:'đơn', d:[[chg(o, orders(LM(m))), 'so với tháng trước', true], [cO, 'so với cùng kỳ', true]]},
    {lb:'AOV', v:nf(a), u:'USD/đơn', d:[[chg(a, aov(LM(m))), 'so với tháng trước', true], [cA, 'so với cùng kỳ', true]]},
    {lb:'Tỷ lệ hoàn trả', v:pc(rrate(m), 2), d:[[rrate(m) - rrate(LM(m)), 'so với tháng trước', false, 'pp'], [rrate(m) - rrate(SP(m)), 'so với cùng kỳ', false, 'pp']]},
    {lb:'Chênh lệch Finance', v:sg(d), u:'USD', alert: fl > 0, note:[`${spc(chg(finv(m), nr), 2)} so với dữ liệu đơn`, `${fl} cửa hàng vượt ngưỡng`]},
  ]);
  const R = win(), cats = R.map(mShort), sel = mShort(m), spI = R.indexOf(SP(m));
  T('t1a', `Net Revenue ${mStr(m)} bằng ${pc(nr / nrev(SP(m)), 0)} cùng kỳ năm trước`);
  line('c1a', cats, [{vals:R.map(x => nrev(x)/1000), color:C.ctx, w:2.2, main:true, tipName:'Net Revenue', marks: spI >= 0 ? [{i:spI, t:'cùng kỳ ' + k1(nrev(SP(m)))}] : []}], {fmt:v => nf(v,0), unit:'nghìn USD', sel, pos:true, r:16});
  const g = gross(m), dc = disc(m), rt = ret(m);
  T('t1b', `Giảm giá và hoàn trả lấy đi ${pc((dc + rt) / g, 1)} Gross Sales`);
  waterfall('c1b', [{n:'Gross Sales', v:g/1000, total:true}, {n:'Giảm giá', v:-dc/1000}, {n:'Hoàn trả', v:-rt/1000}, {n:'Net Revenue', v:nr/1000, total:true, hl:true}], {fmt:v => nf(v,1), unit:'nghìn USD'});
  T('t1c', `Số đơn ${spc(cO)} so với cùng kỳ`);
  line('c1c', cats, [{vals:R.map(orders), color:C.accent, w:2, main:true, tipName:'Số đơn'}], {fmt:v => nf(v,0), unit:'đơn', sel, pos:true, r:16});
  T('t1d', `AOV ${spc(cA)} so với cùng kỳ`);
  line('c1d', cats, [{vals:R.map(aov), color:C.ctx, w:2, main:true, tipName:'AOV'}], {fmt:v => nf(v,0), unit:'USD/đơn', sel, pos:true, r:16});
  const fmon = R.map(x => rowsM(x).filter(r => r.store_key !== 0 && over(r)).length), dav = R.map(x => deltaAdj(x));
  T('t1e', `Sau khi quy online về ngày giao, Δ ${mStr(m)} còn ${sg(deltaAdj(m))} USD`);
  colMonth('c1e', cats, dav.map(v => v / 1000), fmon.map(v => v > 0 ? C.crit : C.ctx), {fmt:v => nf(v,1), unit:'nghìn USD', sel, extra:i => [fmon[i] + ' cửa hàng vượt ngưỡng']});
}

// ================================================================ Trang 2 · Doanh thu
function p2() {
  const m = S.mid, cats = new Map();
  for (const r of D.cat) if (r.month_id === m && keep(r.store_key)) cats.set(r.category, (cats.get(r.category) || 0) + r.ns);
  const tot = ns(m) || 0, ci = [...cats.entries()].sort((a,b) => b[1] - a[1]);
  const cty = new Map(); for (const r of rowsM(m)) { const n = ST[r.store_key].ctry; cty.set(n, (cty.get(n) || 0) + r.ns); }
  const ct = [...cty.entries()].filter(e => e[1] > 0).sort((a,b) => b[1] - a[1]);
  const chv = (mm, ch, f) => { let s = 0; for (const r of (byM.get(mm) || [])) if (ST[r.store_key].ch === ch && keep(r.store_key)) s += f(r); return s; };
  const onl = chv(m, 'ONLINE', r => r.ns), onlS = chv(SP(m), 'ONLINE', r => r.ns), posS = chv(SP(m), 'POS', r => r.ns), pos = chv(m, 'POS', r => r.ns);
  const top = ci[0];
  story('s2-story', top ? `${top[0]} dẫn đầu với ${pc(top[1] / tot, 0)} Net Sales ${mStr(m)}; ${ct.length ? ct[0][0] + ' là thị trường lớn nhất (' + pc(ct[0][1] / tot, 0) + ')' : ''}.` : 'Không có doanh thu trong phạm vi đang chọn.',
    S.ch === 'ALL' ? [`Online chiếm ${pc(onl / (tot || 1), 0)} Net Sales; so với cùng kỳ online ${spc(chg(onl, onlS))}, cửa hàng ${spc(chg(pos, posS))}.`,
      `Giảm giá bằng ${pc(disc(m) / gross(m), 1)} Gross Sales tháng này.`] : [`Giảm giá bằng ${pc(disc(m) / gross(m), 1)} Gross Sales tháng này.`]);
  kpis('s2-k', [
    {lb:'Net Sales', v:k1(tot), u:'nghìn USD', d:[[chg(tot, ns(LM(m))), 'so với tháng trước', true], [chg(tot, ns(SP(m))), 'so với cùng kỳ', true]]},
    {lb:'Số sản phẩm bán', v:nf(units(m)), u:'sản phẩm', d:[[chg(units(m), units(LM(m))), 'so với tháng trước', true], [chg(units(m), units(SP(m))), 'so với cùng kỳ', true]]},
    {lb:'Tỷ trọng online', v:pc(onl / (tot || 1), 0), note:['phần Net Sales từ website/app']},
    {lb:'Giảm giá / Gross Sales', v:pc(disc(m) / gross(m), 1), d:[[disc(m) / gross(m) - disc(SP(m)) / gross(SP(m)), 'so với cùng kỳ', false, 'pp']]},
  ]);
  T('t2a', top ? `${top[0]} chiếm ${pc(top[1] / tot, 0)} Net Sales` : 'Net Sales theo danh mục');
  hbar('c2a', ci.map((e,i) => ({n:e[0], v:e[1]/1000, hl:i === 0, lab:k1(e[1]) + ' · ' + pc(e[1] / tot, 0), tip:[k1(e[1]) + ' nghìn USD', pc(e[1] / tot, 1) + ' Net Sales']})), {fmt:v => nf(v,1), valW:100});
  T('t2b', ct.length ? `${ct[0][0]} đóng góp ${pc(ct[0][1] / tot, 0)} Net Sales` : 'Net Sales theo quốc gia');
  hbar('c2b', ct.map((e,i) => ({n:e[0], v:e[1]/1000, hl:i === 0, lab:k1(e[1]) + ' · ' + pc(e[1] / tot, 0)})), {fmt:v => nf(v,1), valW:100});
  const rows = [['POS','Cửa hàng'],['ONLINE','Website/app']].filter(c => S.ch === 'ALL' || S.ch === c[0]).filter(c => !(c[0] === 'ONLINE' && S.ctry !== 'ALL')).map(([k,n]) => {
    const c = chv(m,k,r => r.ns), o = chv(m,k,r => r.orders); return {n, c, o, aov:o ? c/o : null, l:chv(LM(m),k,r => r.ns), s:chv(SP(m),k,r => r.ns)}; });
  T('t2c', rows.length === 2 ? `AOV cửa hàng ${nf(rows[0].aov)} USD, online ${nf(rows[1].aov)} USD` : 'Doanh thu theo kênh');
  const dcell = x => ({t:spc(x), c:x == null ? 'mut' : ''});
  table('c2c', [{h:'Kênh', l:1, f:r => ({t:r.n})}, {h:'Net Sales', f:r => ({t:k1(r.c)})}, {h:'Đơn', f:r => ({t:nf(r.o)})}, {h:'AOV', f:r => ({t:nf(r.aov)})},
    {h:'vs LM', f:r => dcell(chg(r.c, r.l))}, {h:'vs SPLY', f:r => dcell(chg(r.c, r.s))}], rows);
  const agg = mm => { const o = new Map(); for (const r of rowsM(mm)) o.set(r.store_key, (o.get(r.store_key) || 0) + r.ns); return o; };
  const cur = agg(m), l = agg(LM(m)), s_ = agg(SP(m)), oc = new Map(rowsM(m).map(r => [r.store_key, r.orders]));
  const t10 = [...cur.entries()].filter(e => e[1] > 0).sort((a,b) => b[1] - a[1]).slice(0, 10), mx = t10.length ? t10[0][1] : 1;
  T('t2d', t10.length ? `10 cửa hàng lớn nhất tạo ${pc(t10.reduce((a,e) => a + e[1], 0) / tot, 0)} Net Sales` : 'Top 10 cửa hàng');
  table('c2d', [{h:'#', l:1, f:r => ({t:String(r.i)})}, {h:'Cửa hàng', l:1, f:r => ({t:ST[r.k].name})}, {h:'Quốc gia', l:1, f:r => ({t:ST[r.k].ctry})},
    {h:'Net Sales (nghìn USD)', l:1, f:r => ({t:k1(r.v), bar:r.v / mx, hl:r.i === 1})}, {h:'Đơn', f:r => ({t:nf(oc.get(r.k) || 0)})},
    {h:'vs LM', f:r => dcell(chg(r.v, l.get(r.k)))}, {h:'vs SPLY', f:r => dcell(chg(r.v, s_.get(r.k)))}], t10.map((e,i) => ({i:i + 1, k:e[0], v:e[1]})));
}

// ================================================================ Trang 3 · Đối soát
function p3() {
  const m = S.mid, d = delta(m), da = deltaAdj(m), rows = rowsM(m).filter(r => r.fin || r.ns);
  const onlR = rows.find(r => r.store_key === 0), timing = onlR ? (onlR.ns_post - onlR.ns) : 0;   // phần Δ do online lệch thời điểm = Δ gốc − Δ sau quy đổi
  const fl = rows.filter(over), posFl = fl.filter(r => r.store_key !== 0).sort((a,b) => Math.abs(dOf(b)) - Math.abs(dOf(a)));
  const hqv = D.hq[String(m)];
  const sub = posFl.slice(0,3).map(r => `${ST[r.store_key].code} (${sg(dOf(r))} USD)`).join(', ');
  story('s3-story', `Finance ${d >= 0 ? 'cao' : 'thấp'} hơn dữ liệu đơn ${nf(Math.abs(d))} USD (${spc(d / Math.abs(finv(m) || 1), 2)}).` +
      (onlR && Math.abs(timing) > 1 ? ` ${sg(timing)} USD là lệch thời điểm của kênh online — biến mất khi tính theo ngày giao.` : ''),
    [posFl.length ? `Còn lại ${sg(da)} USD; cần điều tra ${posFl.length} cửa hàng vượt ngưỡng: ${sub}${posFl.length > 3 ? '…' : ''}.` : `Còn lại ${sg(da)} USD, không cửa hàng nào vượt ngưỡng ngoài kênh online.`,
     hqv ? `Chưa gồm bút toán HQ-ADJ ${sg(hqv)} USD (không thuộc cửa hàng nào).` : 'Xuống mức ngày (posting_date) ở các cửa hàng có ⚠ để tìm nguyên nhân và số tiền.']);
  kpis('s3-k', [
    {lb:'Net Revenue · dữ liệu đơn', v:k1(nrev(m)), u:'nghìn USD'},
    {lb:'Net Revenue · Finance', v:k1(finv(m)), u:'nghìn USD', note: hqv ? ['chưa gồm HQ-ADJ ' + sg(hqv) + ' USD'] : []},
    {lb:'Chênh lệch Δ', v:sg(d), u:'USD', alert: Math.abs(d) > D.thr.abs_usd, note:[spc(d / Math.abs(finv(m) || 1), 2) + ' trên Finance']},
    {lb:'Δ sau quy đổi online', v:sg(da), u:'USD', alert: posFl.length > 0, note:['online tính theo ngày giao']},
    {lb:'Cửa hàng vượt ngưỡng', v:String(fl.length), u:'/ ' + rows.length, alert: fl.length > 0, note:[`|Δ| > ${nf(D.thr.abs_usd)} USD và |Δ%| > ${nf(D.thr.pct*100)}%`]},
  ]);
  const steps = [];
  if (onlR && Math.abs(timing) > .5) steps.push({n:'Online lệch thời điểm', v:timing, c:C.ctx});
  posFl.slice(0,3).forEach(r => steps.push({n:ST[r.store_key].code + ' ⚠', v:dOf(r), c:C.crit}));
  const rest = d - steps.reduce((a,s) => a + s.v, 0); if (Math.abs(rest) > .5) steps.push({n:'Cửa hàng khác', v:rest, c:C.ctxD});
  steps.push({n:'Δ Finance − đơn', v:d, total:true, hl:true});
  T('t3a', posFl.length ? `${ST[posFl[0].store_key].code} là chênh lệch thật lớn nhất (${sg(dOf(posFl[0]))} USD)` : 'Chênh lệch chủ yếu là lệch thời điểm');
  waterfall('c3a', steps, {fmt:v => nf(v/1000,1), unit:'nghìn USD'});
  const R = win(), cats = R.map(mShort), sel = mShort(m);
  const posOver = R.map(x => rowsM(x).filter(r => r.store_key !== 0 && over(r)).length);
  const dav = R.map(x => deltaAdj(x)), rk = [...dav].map(Math.abs).sort((a,b) => b - a).indexOf(Math.abs(deltaAdj(m))) + 1;
  T('t3b', `Δ sau quy đổi ${mStr(m)} là ${sg(deltaAdj(m))} USD — lớn thứ ${rk}/${R.length} tháng gần nhất`);
  colMonth('c3b', cats, R.map(x => deltaAdj(x)/1000), posOver.map(v => v > 0 ? C.crit : C.ctx), {fmt:v => nf(v,1), unit:'nghìn USD', sel, extra:i => [posOver[i] + ' cửa hàng vượt ngưỡng (POS)']});
  const tr = rows.map(r => ({k:r.store_key, nr:r.ns - r.ret, fin:r.fin, d:dOf(r), p:r.fin ? dOf(r)/Math.abs(r.fin) : null, da:r.fin - (r.ns_post - r.ret), fl:over(r)}))
    .sort((a,b) => Math.abs(b.d) - Math.abs(a.d)).slice(0, 10);
  T('t3c', fl.length ? `${fl.length} cửa hàng vượt ngưỡng tháng ${mStr(m)}` : `Không cửa hàng nào vượt ngưỡng tháng ${mStr(m)}`);
  table('c3c', [{h:'Cửa hàng', l:1, f:r => ({t:ST[r.k].code + ' · ' + ST[r.k].ctry})}, {h:'Đơn hàng', f:r => ({t:nf(r.nr)})}, {h:'Finance', f:r => ({t:nf(r.fin)})},
    {h:'Δ', f:r => ({t:sg(r.d), c:r.fl ? 'flag' : ''})}, {h:'Δ %', f:r => ({t:spc(r.p), c:r.fl ? 'flag' : 'mut'})}, {h:'Δ sau quy đổi', f:r => ({t:sg(r.da)})},
    {h:'Cờ', f:r => r.fl ? {t:'⚠ Vượt', c:'flag'} : {t:'trong ngưỡng', c:'mut'}}], tr);
  if (S.ch === 'POS' || S.ctry !== 'ALL') { T('t3d', 'Kênh online không thuộc bộ lọc đang chọn'); empty('c3d', 'Chọn Kênh = Tất cả và Quốc gia = Tất cả để xem kênh online.'); }
  else { const raw = R.map(x => { const r = (byM.get(x) || []).find(z => z.store_key === 0); return r ? dOf(r)/1000 : null; });
    const adj = R.map(x => { const r = (byM.get(x) || []).find(z => z.store_key === 0); return r ? (r.fin - (r.ns_post - r.ret))/1000 : null; });
    T('t3d', 'Δ online về gần 0 khi tính theo ngày giao');
    line('c3d', cats, [{name:'Δ gốc', vals:raw, color:C.ctx, w:2}, {name:'sau quy đổi', vals:adj, color:C.accent, w:2.4, main:true}], {fmt:v => nf(v,1), unit:'nghìn USD', sel, zero:true, r:86}); }
}

// ================================================================ Trang 4 · Hoàn trả & chất lượng
function p4() {
  const m = S.mid, rr = new Map();
  for (const r of D.rr) if (r.month_id === m && keep(r.store_key)) rr.set(r.return_reason, (rr.get(r.return_reason) || 0) + r.ret);
  const ri = [...rr.entries()].sort((a,b) => b[1] - a[1]), rtot = ri.reduce((a,e) => a + e[1], 0);
  const cm = mm => { const o = new Map(); for (const r of D.cat) if (r.month_id === mm && keep(r.store_key)) { const x = o.get(r.category) || {ns:0, ret:0}; x.ns += r.ns; x.ret += r.ret; o.set(r.category, x); } return o; };
  const cc = [...cm(m).entries()].filter(e => e[1].ns > 0).map(([k,v]) => ({n:k, ns:v.ns, rt:v.ret / v.ns})).sort((a,b) => b.rt - a.rt);
  const canc = mm => sumM(mm, r => r.canc), pend = mm => sumM(mm, r => r.pend);
  const dupAll = D.ms.filter(r => keep(r.store_key)).reduce((a,r) => a + r.dup, 0), pendAll = D.ms.filter(r => keep(r.store_key)).reduce((a,r) => a + r.pend, 0);
  story('s4-story', `Hoàn trả ${mStr(m)} là ${nf(ret(m))} USD, bằng ${pc(rrate(m), 2)} Net Sales (${sg((rrate(m) - rrate(SP(m))) * 100, 2)} điểm % so với cùng kỳ).`,
    [ri.length ? `Lý do lớn nhất: ${ri[0][0]} (${pc(ri[0][1] / rtot, 0)} giá trị trả).` : 'Không có hoàn trả trong phạm vi đang chọn.',
     `Trước khi tính KPI phải loại ${nf(dupAll)} dòng trùng và ${nf(pendAll)} đơn đang giao (PENDING) — tính toàn thời gian, theo bộ lọc kênh/quốc gia.`]);
  kpis('s4-k', [
    {lb:'Giá trị hoàn trả', v:nf(ret(m)), u:'USD', d:[[chg(ret(m), ret(LM(m))), 'so với tháng trước', false], [chg(ret(m), ret(SP(m))), 'so với cùng kỳ', false]]},
    {lb:'Tỷ lệ hoàn trả', v:pc(rrate(m), 2), d:[[rrate(m) - rrate(LM(m)), 'so với tháng trước', false, 'pp'], [rrate(m) - rrate(SP(m)), 'so với cùng kỳ', false, 'pp']]},
    {lb:'Đơn huỷ', v:nf(canc(m) || 0), u:'đơn', d:[[chg(canc(m), canc(SP(m))), 'so với cùng kỳ', false]]},
    {lb:'Dòng trùng đã loại', v:nf(dupAll), u:'dòng', note:['toàn thời gian, theo bộ lọc', nf(pendAll) + ' đơn PENDING tại ' + D.asof.split('-').reverse().join('/')]},
  ]);
  const R = win(), cats = R.map(mShort), sel = mShort(m);
  T('t4a', `Tỷ lệ hoàn trả ${pc(rrate(m), 2)} — ${rrate(m) > rrate(SP(m)) ? 'cao' : 'thấp'} hơn cùng kỳ ${nf(Math.abs(rrate(m) - rrate(SP(m))) * 100, 2)} điểm %`);
  line('c4a', cats, [{vals:R.map(x => { const v = rrate(x); return v == null ? null : v * 100; }), color:C.ctx, w:2.2, main:true, tipName:'Tỷ lệ hoàn trả'}], {fmt:v => nf(v,1) + '%', sel, zero:true, r:16});
  T('t4b', ri.length ? `${ri[0][0]} là lý do trả hàng lớn nhất` : 'Lý do trả hàng');
  hbar('c4b', ri.map((e,i) => ({n:e[0], v:e[1], hl:i === 0, lab:nf(e[1]) + ' · ' + pc(e[1] / rtot, 0)})), {fmt:v => nf(v,0), valW:96});
  T('t4c', cc.length ? `${cc[0].n} có tỷ lệ hoàn trả cao nhất (${pc(cc[0].rt, 1)})` : 'Tỷ lệ hoàn trả theo danh mục');
  hbar('c4c', cc.map((e,i) => ({n:e.n, v:e.rt * 100, hl:i === 0, lab:pc(e.rt, 1) + ' · NS ' + k1(e.ns) + 'k', tip:[pc(e.rt, 2) + ' tỷ lệ hoàn trả', 'Net Sales ' + k1(e.ns) + ' nghìn USD']})), {fmt:v => nf(v,1) + '%', valW:120});
  T('t4d', `${nf(canc(m) || 0)} đơn huỷ tháng ${mStr(m)}`);
  line('c4d', cats, [{name:'Huỷ', vals:R.map(canc), color:C.accent, w:2, main:true}, {name:'Đang giao', vals:R.map(pend), color:C.ctx, w:2}], {fmt:v => nf(v,0), unit:'đơn', sel, zero:true, r:76});
}

// ---------------------------------------------------------------- điều khiển
function header() {
  const ms = $('#f-month'), R = MONTHS.slice().reverse();
  if (ms.options.length !== R.length) ms.replaceChildren(...R.map(m => new Option(mStr(m), String(m))));
  ms.value = String(S.mid);
  const cs = $('#f-ctry'); if (cs.options.length !== CTRY.length + 1) cs.replaceChildren(new Option('Tất cả quốc gia', 'ALL'), ...CTRY.map(c => new Option(c, c)));
  cs.value = S.ctry; $('#f-ch').value = S.ch;
  $('#sub').textContent = 'Nexa Home · Sales Project 1 · dữ liệu đến ' + D.asof.split('-').reverse().join('/');
  $('#foot').textContent = 'Nguồn: Nexa_Sales_P1_Dataset.xlsx (đơn hàng, hoàn trả) và finance_revenue.xlsx (Finance). Dữ liệu đến ' + D.asof.split('-').reverse().join('/') +
    '. USD, chưa VAT. Chỉ tính đơn COMPLETED, đã loại dòng trùng; Returns theo ngày trả hàng. Δ = Finance − dữ liệu đơn; vượt ngưỡng khi |Δ| > ' + nf(D.thr.abs_usd) + ' USD và |Δ%| > ' + nf(D.thr.pct*100) + '% ở mức tháng × cửa hàng. Kênh online không thuộc quốc gia nào: khi chọn quốc gia, online bị loại.';
}
function render() { header(); ({s1:p1, s2:p2, s3:p3, s4:p4})[page](); }
const set = o => { Object.assign(S, o); render(); };
$('#f-month').addEventListener('change', e => set({mid:Number(e.target.value)}));
$('#f-ch').addEventListener('change', e => set({ch:e.target.value, ctry:e.target.value === 'ONLINE' ? 'ALL' : S.ctry}));
$('#f-ctry').addEventListener('change', e => set({ctry:e.target.value, ch:e.target.value !== 'ALL' && S.ch === 'ONLINE' ? 'ALL' : S.ch}));
$('#f-reset').addEventListener('click', () => set({...DEF}));
$('#tabs').addEventListener('click', e => { const b = e.target.closest('button[data-p]'); if (b) go(b.dataset.p); });
function go(id) { page = id; document.querySelectorAll('.page').forEach(s => s.classList.toggle('on', s.id === id));
  document.querySelectorAll('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.p === id)); render(); }
const named = (location.hash || '').slice(1); if (/^s[1-4]$/.test(named)) page = named;
go(page);
let t; addEventListener('resize', () => { clearTimeout(t); t = setTimeout(render, 80); });
})();
</script>
"""

html = ('<!doctype html>\n<html lang="vi"><head><meta charset="utf-8"><meta name="viewport" '
        'content="width=device-width,initial-scale=1"><title>Kết quả bán hàng và đối soát – Nexa Home</title>'
        + CSS + '</head><body>\n' + BODY + d3_block + "\n<script>window.__NX_DATA=" + payload + ";</script>\n" + JS
        + "</body></html>\n")
OUT.write_text(html, encoding="utf-8")
print(OUT.name, f"{OUT.stat().st_size/1e6:.2f} MB")
