# -*- coding: utf-8 -*-
"""Mot vong dung dashboard Power BI, chay bang MOT lenh. Khung dung chung.

Chep tep nay (cung `loop.json`) vao thu muc chua `<du_an>.pbip`, sua `loop.json`, roi:

    python loop.py [--no-model] [--no-refresh] [--no-export] [--tu N]

Moi thay doi phai di qua vong nay. Sua `build_model.py` hoac `build_report.py`
xong thi chay lenh tren, roi mo `review/render_png/page_N.png` ma xem.

Cac cong xep theo thu tu loi xuat hien, RE TRUOC DAT SAU, dung ngay o cong
hong dau tien:

   1 truoc_model        cac script kiem du lieu nguon (vd dau van tay DB) -- tuy du an
   2 build_model        sinh TMDL (+ lint DAX tinh bang scripts/dax_lint.py)
   3 build_report       sinh PBIR (+ pbi report validate)
   4 kiem_schema.py     doi chieu $schema cua Microsoft
   5 pbi_desktop.py     dong KHONG LUU, mo lai <du_an>.pbip
   6 pbi_health.ps1     Desktop co bi hop thoai chan khong
   7 nap du lieu        tuan tu tung bang, bang nho truoc (nap song song lam dut ket noi)
   8 pbi_measures_check measure nao hong va vi sao
   9 kiem_so            DAX so voi so ky vong (moc.json tinh bang SQL/Excel) -- tuy du an
  10 dai thong bao      pbi_ui.ps1 tim dai "pending changes"/"incomplete data"
  11 export_pdf.ps1     xuat 2 lan (khoi dong + that), chiem chuot ~1 phut
  12 pdf2png.py         cat PNG + luu ban sao review/render_rNN de so diem anh

**Phai dong va mo lai Desktop moi vong.** `pbi report reload` khong nap lai
JSON sua ngoai dia. Va KHONG BAO GIO gui Ctrl+S vao cua so Desktop: no ghi de
ban sua tren dia bang ban cu trong bo nho.

Tuy chon:
    --no-model     bo qua build_model (chi doi trinh bay)
    --no-refresh   khong nap lai du lieu (chi dung khi Desktop con cache)
    --no-export    dung truoc khi xuat PDF (chi muon kiem so)
    --tu N         bat dau tu cong N (vd --tu 8 khi Desktop dang mo dung du an)

Nguon goc: gop tu Game/PowerBI/JellyTrail/loop.py, Finance/PowerBI/tools/loop_v3.py
va SupplyChain/powerbi/scripts/loop.py (2026-10).
"""

import json
import os
import shutil
import subprocess
import sys
import time

for _luong in (sys.stdout, sys.stderr):
    try:
        _luong.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
REVIEW = os.path.join(HERE, "review")


def doc_cau_hinh():
    p = os.path.join(HERE, "loop.json")
    if not os.path.exists(p):
        sys.exit("Thieu loop.json canh loop.py -- chep tu <skill>/mau/loop.json roi sua.")
    with open(p, encoding="utf-8") as f:
        ch = json.load(f)
    ch.setdefault("bang_nap", [])
    ch.setdefault("truoc_model", [])
    ch.setdefault("kiem_so", [])
    ch.setdefault("build_model", "build_model.py")
    ch.setdefault("build_report", "build_report.py")
    ch.setdefault("xuat_khoi_dong", True)
    ch["python_db"] = ch.get("python_db") or sys.executable
    # thu muc scripts cua skill: tu loop.json, bien PBI_SKILL_SCRIPTS, hoac ../.claude/skills/...
    ch["scripts"] = os.path.abspath(os.path.join(HERE, ch.get("scripts")
                                                 or os.environ.get("PBI_SKILL_SCRIPTS", "")))
    if not os.path.exists(os.path.join(ch["scripts"], "pbi_desktop.py")):
        sys.exit("Khong thay scripts cua skill tai %s -- dat khoa \"scripts\" trong loop.json."
                 % ch["scripts"])
    if not ch.get("du_an"):
        sys.exit("loop.json thieu khoa \"du_an\" (ten tep .pbip, khong duoi).")
    return ch


if "-h" in sys.argv or "--help" in sys.argv:
    print(__doc__)
    sys.exit(0)

CH = doc_cau_hinh()
PBIP = os.path.join(HERE, CH["du_an"] + ".pbip")
SCRIPTS = CH["scripts"]
ENV = dict(os.environ, PYTHONIOENCODING="utf-8", PBI_PBIP=PBIP, PBI_PROJECT=CH["du_an"])


def buoc(n, ten):
    print("\n===== %d. %s %s" % (n, ten, "=" * max(0, 60 - len(ten))), flush=True)


def chay(args, cwd=None, chan=True, cho=None, im=False, ma_bo_qua=()):
    r = subprocess.run(args, cwd=cwd or HERE, env=ENV, capture_output=True,
                       text=True, encoding="utf-8", errors="replace", timeout=cho)
    ra = ((r.stdout or "") + (r.stderr or "")).rstrip()
    if ra and not im:
        print(ra[-3500:], flush=True)
    if chan and r.returncode != 0 and r.returncode not in ma_bo_qua:
        print("\nDỪNG: %s trả mã %d." % (" ".join(os.path.basename(a) for a in args[:3]),
                                         r.returncode))
        sys.exit(r.returncode)
    return r


def ps(duong, *them, **kw):
    return chay(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-STA", "-File",
                 duong] + list(them), **kw)


def chay_script_du_an(muc, bat_buoc=False):
    """muc = {"ten": "kiem_so.py", "mo_ta": "...", "python": "db"|None, "ma_bo_qua": [..]}"""
    duong = os.path.join(HERE, muc["ten"])
    if not os.path.exists(duong):
        if bat_buoc:
            sys.exit("DỪNG: thiếu %s — dự án chưa có script sinh, không dựng được gì." % muc["ten"])
        print("  (chưa có %s — bỏ qua)" % muc["ten"])
        return
    py = CH["python_db"] if muc.get("python") == "db" else sys.executable
    r = chay([py, duong] + list(muc.get("tham_so", [])), ma_bo_qua=tuple(muc.get("ma_bo_qua", ())))
    if r.returncode in muc.get("ma_bo_qua", ()):
        print("  (cảnh báo) %s trả mã %d — bỏ qua theo loop.json." % (muc["ten"], r.returncode))


def dat_ban_render(png_dir):
    """Chep render_png sang render_rNN de so_diem_anh.py so hai vong."""
    so = [int(d[8:]) for d in os.listdir(REVIEW) if d.startswith("render_r") and d[8:].isdigit()]
    dich = os.path.join(REVIEW, "render_r%02d" % (max(so, default=0) + 1))
    shutil.copytree(png_dir, dich)
    return dich


def main():
    co = set(a for a in sys.argv[1:] if a.startswith("--"))
    tu = int(sys.argv[sys.argv.index("--tu") + 1]) if "--tu" in sys.argv else 1
    t0 = time.time()
    os.makedirs(REVIEW, exist_ok=True)
    n = 0

    def cong(ten):
        nonlocal n
        n += 1
        if n < tu:
            return False
        buoc(n, ten)
        return True

    if cong("kiểm dữ liệu nguồn (truoc_model)"):
        for muc in CH["truoc_model"]:
            chay_script_du_an(muc)

    if cong("%s — sinh TMDL, lint DAX tĩnh" % CH["build_model"]):
        if "--no-model" in co:
            print("  (--no-model) bỏ qua")
        else:
            chay_script_du_an({"ten": CH["build_model"]}, bat_buoc=True)

    if cong("%s — sinh PBIR" % CH["build_report"]):
        chay_script_du_an({"ten": CH["build_report"]}, bat_buoc=True)

    if cong("kiem_schema.py — đối chiếu $schema của Microsoft"):
        chay([sys.executable, os.path.join(SCRIPTS, "kiem_schema.py"),
              os.path.join(HERE, CH["du_an"] + ".Report")])

    if cong("đóng KHÔNG LƯU và mở lại Power BI Desktop"):
        chay([sys.executable, os.path.join(SCRIPTS, "pbi_desktop.py"), "open", "--pbip", PBIP],
             cho=400)

    if cong("pbi_health.ps1 — Desktop có bị hộp thoại chặn không"):
        ps(os.path.join(SCRIPTS, "pbi_health.ps1"), "-OutDir", REVIEW)

    if shutil.which("pbi") is None:
        sys.exit("DỪNG: không thấy `pbi` trên PATH — cài pbi-cli hoặc dựng shim bằng scripts/build_pbi_shim.ps1.")
    chay(["pbi", "connect"], im=True, chan=False)

    if cong("nạp dữ liệu, tuần tự từng bảng"):
        if "--no-refresh" in co or not CH["bang_nap"]:
            print("  bỏ qua (%s)" % ("--no-refresh" if "--no-refresh" in co
                                     else "loop.json chưa liệt kê bang_nap"))
        for b in ([] if "--no-refresh" in co else CH["bang_nap"]):
            s = time.time()
            r = chay(["pbi", "table", "refresh", b, "--type", "Full"], chan=False, im=True,
                     cho=1200)
            ok = "refreshed" in ((r.stdout or "") + (r.stderr or ""))
            print("   %-4s %-24s %5.0fs" % ("OK" if ok else "LỖI", b, time.time() - s),
                  flush=True)
            if not ok:
                print(((r.stdout or "") + (r.stderr or ""))[-1500:])
                sys.exit(2)

    if cong("pbi_measures_check.py — measure nào hỏng"):
        chay([sys.executable, os.path.join(SCRIPTS, "pbi_measures_check.py")])

    if cong("kiểm số của dự án (kiem_so)"):
        if not CH["kiem_so"]:
            print("  CẢNH BÁO: loop.json chưa có cổng kiểm số — vòng này chỉ kiểm hình thức.")
        for muc in CH["kiem_so"]:
            chay_script_du_an(muc)

    if cong("dải thông báo trên canvas (in cả vào PDF nếu còn)"):
        r = chay(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                  os.path.join(SCRIPTS, "pbi_ui.ps1"), "-List", "-Match",
                  "pending changes|incomplete or no data|need to be manually refreshed"],
                 chan=False, im=True)
        dai = sorted({l.split("PBI_UI:")[1].split("[")[0].strip()
                      for l in (r.stdout or "").splitlines()
                      if "PBI_UI:" in l and "StatusBar" in l})
        if dai:
            print("  CẢNH BÁO — Desktop đang hiện dải thông báo, PDF sẽ có nó:")
            for l in dai:
                print("   ! " + l)
        else:
            print("  Sạch: không có dải thông báo nào.")

    if "--no-export" not in co and cong("xuất PDF rồi cắt PNG — KHÔNG đụng chuột lúc này"):
        tat_pdf = ["powershell", "-NoProfile", "-Command",
                   "Stop-Process -Name Acrobat,AcroRd32,AcroCEF -Force -ErrorAction SilentlyContinue"]
        xuat = os.path.join(SCRIPTS, "export_pdf.ps1")
        loc = ["-TitleLike", "*%s*" % CH["du_an"], "-DebugPng", os.path.join(REVIEW, "pbi_debug.png")]
        chay(tat_pdf, chan=False, im=True)
        if CH["xuat_khoi_dong"]:
            # Lan Export to PDF DAU TIEN sau khi mo Desktop ve trang dang mo thu nho ~88%;
            # lan thu hai thi dung. Xuat mot lan khoi dong roi bo di (do that 29/09/2026).
            time.sleep(10)
            ps(xuat, "-OutPdf", os.path.join(REVIEW, "_khoi_dong.pdf"), *loc, cho=600, im=True)
            chay(tat_pdf, chan=False, im=True)
        pdf = os.path.join(REVIEW, "render.pdf")
        ps(xuat, "-OutPdf", pdf, *loc, cho=600)
        png = os.path.join(REVIEW, "render_png")
        if os.path.isdir(png):
            shutil.rmtree(png)
        chay([sys.executable, os.path.join(SCRIPTS, "pdf2png.py"), pdf, png])
        print("  bản lưu để so điểm ảnh: %s" % dat_ban_render(png))

    print("\nXONG %d cổng trong %.0f giây." % (n, time.time() - t0))
    print("Mở review/render_png/page_N.png để xem.")


if __name__ == "__main__":
    main()
