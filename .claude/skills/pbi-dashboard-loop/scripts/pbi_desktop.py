#!/usr/bin/env python3
"""Dieu khien Power BI Desktop: mo du an, phat hien hop thoai loi, chup anh.

Nguyen tac phong thu (vi tu dong hoa giao dien la mat xich de gay nhat):
  - KHANG DINH tieu de cua so truoc khi gui bat ky phim nao. Khong khop ->
    thoat ngay, khong go mot phim nao. Go nham vao cua so khac la tai hoa.
  - Moi buoc cho deu co HAN GIO CUNG. Het gio thi bao cao trang thai that,
    khong im lang treo.
  - Bang chung hoan thanh la TRANG THAI THAT (tieu de cua so, file xuat hien
    tren dia), khong phai "da gui phim di roi".

  python pbi_desktop.py status
  python pbi_desktop.py open  --pbip <du-an>/X.pbip   # dong ban cu, mo lai .pbip
  python pbi_desktop.py close

Chon du an: `--pbip` hoac bien moi truong PBI_PBIP (duong dan .pbip). Ten du an
(dung de khang dinh tieu de cua so) = ten tep .pbip, ghi de bang PBI_PROJECT.
Anh chup loi ghi vao <thu muc .pbip>/review/shots.
"""
from __future__ import annotations

import os
import subprocess
import sys
import time
from ctypes import windll
from pathlib import Path

import win32con
import win32gui
import win32process
import win32ui

def _chon_pbip() -> Path | None:
    """--pbip trong argv thang bien PBI_PBIP. Khong co ca hai thi chi dung duoc status/close."""
    if "--pbip" in sys.argv:
        i = sys.argv.index("--pbip")
        if i + 1 < len(sys.argv):
            return Path(sys.argv[i + 1]).resolve()
    v = os.environ.get("PBI_PBIP")
    return Path(v).resolve() if v else None


# Luu y: phep khang dinh tieu de dung "PROJECT in tieu_de", ma ten "X" la TIEN TO
# cua "X_v3" -- mo nham ban cu van khop. Dat ten du an khong la tien to cua nhau.
PBIP = _chon_pbip()
PROJECT = os.environ.get("PBI_PROJECT") or (PBIP.stem if PBIP else "")
SHOTS = (PBIP.parent if PBIP else Path.cwd()) / "review" / "shots"
PBIDESKTOP = Path(r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe")
# Cac tieu de hop thoai loi ma Desktop dung. Gap la dung ngay.
ERROR_TITLES = ("Issues were found", "Couldn't load the model", "Unable to open document")


def visible_windows() -> list[tuple[int, str, str]]:
    out: list[tuple[int, str, str]] = []

    def cb(h, _):
        if win32gui.IsWindowVisible(h):
            t = win32gui.GetWindowText(h)
            if t.strip():
                out.append((h, win32gui.GetClassName(h), t))
        return True

    win32gui.EnumWindows(cb, None)
    return out


def _pbi_pids() -> set[int]:
    out = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq PBIDesktop.exe", "/FO", "CSV", "/NH"],
        capture_output=True, text=True).stdout
    pids = set()
    for line in out.splitlines():
        parts = [c.strip('"') for c in line.split('","')]
        if len(parts) > 1 and parts[1].isdigit():
            pids.add(int(parts[1]))
    return pids


def pbi_windows() -> list[tuple[int, str]]:
    """Nhan dien theo TIEN TRINH, khong theo chu trong tieu de.

    Desktop dat tieu de la ten du an tran (khong co duoi ' - Power BI
    Desktop') khi mo .pbip, nen loc theo chuoi se truot.
    """
    pids = _pbi_pids()
    out = []
    for h, _, t in visible_windows():
        try:
            _, pid = win32process.GetWindowThreadProcessId(h)
        except Exception:
            continue
        if pid in pids:
            out.append((h, t))
    return out


def error_dialog() -> tuple[int, str] | None:
    for h, _, t in visible_windows():
        if any(e in t for e in ERROR_TITLES):
            return h, t
    return None


def shoot(hwnd: int, name: str) -> Path:
    """Chup rieng mot cua so ra PNG."""
    SHOTS.mkdir(parents=True, exist_ok=True)
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    time.sleep(0.4)
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    w, h = max(r - l, 1), max(b - t, 1)
    wdc = win32gui.GetWindowDC(hwnd)
    mfc = win32ui.CreateDCFromHandle(wdc)
    save = mfc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfc, w, h)
    save.SelectObject(bmp)
    # 2 = PW_RENDERFULLCONTENT, bat buoc de bat duoc noi dung ve bang GPU
    windll.user32.PrintWindow(hwnd, save.GetSafeHdc(), 2)
    bmp_path = SHOTS / f"{name}.bmp"
    bmp.SaveBitmapFile(save, str(bmp_path))
    win32gui.DeleteObject(bmp.GetHandle())
    save.DeleteDC()
    mfc.DeleteDC()
    win32gui.ReleaseDC(hwnd, wdc)
    png = SHOTS / f"{name}.png"
    try:
        from PIL import Image

        Image.open(bmp_path).save(png)
        bmp_path.unlink(missing_ok=True)
        return png
    except Exception:
        return bmp_path


def close_all(timeout: float = 60) -> None:
    """Dong moi cua so Power BI Desktop, khong luu."""
    for h, t in pbi_windows():
        win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
    for h, _, t in visible_windows():
        if any(e in t for e in ERROR_TITLES):
            win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
    deadline = time.time() + timeout
    while time.time() < deadline:
        if not pbi_windows():
            break
        time.sleep(1)
    # Con sot thi ket lieu tien trinh (du an nam tren dia, khong mat gi).
    subprocess.run(["taskkill", "/F", "/IM", "PBIDesktop.exe"],
                   capture_output=True, text=True)
    subprocess.run(["taskkill", "/F", "/IM", "msmdsrv.exe"],
                   capture_output=True, text=True)
    time.sleep(2)


def open_project(timeout: float = 240) -> dict:
    """Mo .pbip, cho den khi tieu de mang ten du an hoac hien hop thoai loi."""
    if PBIP is None:
        raise SystemExit("Chua chi .pbip: dung --pbip <tep.pbip> hoac dat PBI_PBIP")
    if not PBIP.exists():
        raise SystemExit(f"Khong thay {PBIP}")
    subprocess.Popen([str(PBIDESKTOP), str(PBIP)])
    deadline = time.time() + timeout
    while time.time() < deadline:
        err = error_dialog()
        if err:
            h, t = err
            png = shoot(h, "loi_mo_du_an")
            return {"ok": False, "ly_do": t, "anh": str(png)}
        for h, t in pbi_windows():
            if PROJECT in t:
                # Phong to NGAY: Export to PDF ve trang dang mo theo kich thuoc khung
                # hien tai. Phong to sat luc xuat thi trang dau bi thu ~88% (do that 29/09/2026).
                try:
                    win32gui.ShowWindow(h, win32con.SW_MAXIMIZE)
                except Exception:
                    pass
                return {"ok": True, "tieu_de": t, "hwnd": h}
        time.sleep(2)
    titles = [t for _, t in pbi_windows()]
    return {"ok": False, "ly_do": f"het {timeout:.0f}s, tieu de hien tai: {titles}"}


def main() -> int:
    cmd = sys.argv[1] if len(sys.argv) > 1 else "status"
    if cmd == "status":
        wins = pbi_windows()
        err = error_dialog()
        print(f"cua so Power BI : {wins if wins else 'khong co'}")
        print(f"hop thoai loi   : {err if err else 'khong co'}")
        if err:
            print(f"anh             : {shoot(err[0], 'hop_thoai_loi')}")
        return 0
    if cmd == "close":
        close_all()
        print("da dong Power BI Desktop")
        return 0
    if cmd == "open":
        # Kiem .pbip TRUOC khi dong: dong roi moi phat hien sai duong dan la giet Desktop
        # cua nguoi dung ma khong mo lai duoc gi.
        if PBIP is None or not PBIP.exists():
            print(f"Khong thay .pbip ({PBIP}) -- dung --pbip <tep.pbip> hoac dat PBI_PBIP")
            return 1
        close_all()
        res = open_project()
        for k, v in res.items():
            print(f"  {k}: {v}")
        return 0 if res.get("ok") else 1
    print(__doc__)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
