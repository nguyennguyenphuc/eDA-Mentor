# -*- coding: utf-8 -*-
"""Goi model NGOAI Claude: ChatGPT web (qua Chrome cua nguoi dung) va Gemini (qua agy CLI).

    python goi_ngoai.py chatgpt "<cau hoi>" [--tep a.md ...] [--ra tra_loi.md] [--kieu hoi|cham]
    python goi_ngoai.py chatgpt --prompt-tep cau_hoi.md [--tep ...] [--ra ...]
    python goi_ngoai.py agy "<viec>" --tep a.py [--tep b.py ...] [--suy-luan]
    python goi_ngoai.py --tu-kiem

Dich:  chatgpt -> chatgpt.com qua `chatgpt_bridge.py` (Chrome mo san o cong 9222, da dang nhap)
       agy     -> Gemini qua Antigravity CLI, lenh doc tu `cau-hinh/agy-cmd.json`:
                  viec code -> 3.8 Flash; `--suy-luan` -> 3.1 Pro

Ra:    stdout: TOM TAT <=10 dong. Claude CHI doc tom tat; nguyen van tra loi nam o tep `--ra`
       (chatgpt, mac dinh `nhat-ky/tra-loi/<thoi diem>.md`) hoac trong chinh cac tep agy da sua.
       Moi lan goi ghi mot dong vao `nhat-ky/goi-ngoai.tsv`.

Ma thoat:  0 xong
           1 agy chay xong nhung tu bao LOI (doc tom tat)
           2 FALLBACK: ben ngoai hong 3 lan -> luong Claude tu lam viec ay (hoac spawn subagent)
           3 CHUA-CAI: thieu agy / thieu Python co websocket-client -> tu lam bang Bash

Nguon goc: rut tu `academic-writing-v3/scripts/giao_viec.py` (2026-10-02): giu `goi_chatgpt`,
`python_cho_cau_noi`, `lam_co_khi`, `nhat_ky`, `fallback`; bo cac vai cham bai giang.
"""
import argparse
import datetime
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import uuid

DAY = os.path.dirname(os.path.abspath(__file__))
GOC = os.path.dirname(DAY)                                   # .../goi-ngoai
NHAT_KY_DIR = os.path.join(GOC, "nhat-ky")
NHAT_KY = os.path.join(NHAT_KY_DIR, "goi-ngoai.tsv")
AGY_CMD = os.path.join(GOC, "cau-hinh", "agy-cmd.json")
CAU_NOI = os.path.join(DAY, "chatgpt_bridge.py")

MAU_AGY = """Bạn làm một việc cơ khí trong một dự án phân tích dữ liệu. Ràng buộc:
- Chỉ sửa các tệp được liệt kê dưới đây. Tệp nào khác thì không được tạo, sửa hay xoá.
- Không đổi câu chữ tiếng Việt trong văn xuôi hay nhãn hiển thị, trừ khi việc yêu cầu rõ.
- Giữ nguyên mã hoá UTF-8 và kiểu xuống dòng của từng tệp. Gạch chéo ngược `\\` phải giữ nguyên.
- Xong việc thì báo kết quả đúng khuôn bên dưới.

VIỆC:
{{VIEC}}

TỆP ĐƯỢC PHÉP ĐỤNG:
{{TEP}}

TRẢ LỜI CUỐI CÙNG ĐÚNG KHUÔN:

###PHIEU###
[KETQUA]
<XONG|LOI> | <một câu tóm tắt>
[TEP]
<đường dẫn> | <đã sửa gì, một dòng>
###HET###
"""


# ------------------------------------------------------------------ tien ich

def doc(p):
    with io.open(p, encoding="utf-8-sig") as f:
        return f.read()


def ghi(p, s):
    d = os.path.dirname(p)
    if d and not os.path.isdir(d):
        os.makedirs(d)
    with io.open(p, "w", encoding="utf-8", newline="\n") as f:
        f.write(s)


def tach_phieu(tra):
    """Khoi ###PHIEU### ... ###HET### -> {MUC: [[cot, ...], ...]}. None neu khong co khoi."""
    m = re.search(r"#{3}\s*PHIEU\s*#{3}(.*?)#{3}\s*HET\s*#{3}", tra or "", re.S | re.I)
    if not m:
        return None
    ra, muc = {}, None
    for ln in m.group(1).split("\n"):
        ln = ln.strip()
        if not ln or ln.startswith("```"):
            continue
        h = re.match(r"^\[([A-Z]+)\]$", ln)
        if h:
            muc = h.group(1)
            ra.setdefault(muc, [])
            continue
        if muc is None:
            continue
        ln = ln.strip("|").strip()
        ra[muc].append([c.strip() for c in ln.split("|")])
    return ra


def nhat_ky(vai, nhan, dich, n_ky_tu, giay, ket_qua):
    moi = not os.path.exists(NHAT_KY)
    if not os.path.isdir(NHAT_KY_DIR):
        os.makedirs(NHAT_KY_DIR)
    with io.open(NHAT_KY, "a", encoding="utf-8", newline="\n") as f:
        if moi:
            f.write("thoi_diem\tvai\tnhan\tdich\tky_tu_prompt\tgiay\tket_qua\n")
        f.write("%s\t%s\t%s\t%s\t%d\t%.0f\t%s\n" % (
            datetime.datetime.now().isoformat(timespec="seconds"), vai, nhan, dich,
            n_ky_tu, giay, ket_qua.replace("\t", " ").replace("\n", " ")[:120]))


def fallback(vai, ly_do):
    print("FALLBACK: %s hong -- %s. Luong Claude tu lam viec nay." % (vai, ly_do))
    sys.exit(2)


def chua_cai(ly_do):
    print("CHUA-CAI: %s" % ly_do)
    sys.exit(3)


def gop_tep(tep):
    phan = []
    for t in tep:
        phan.append("----- %s -----\n%s" % (os.path.basename(t), doc(t)))
    return "\n\n".join(phan)


# ------------------------------------------------------------------ dich: ChatGPT web

_PY_CAU_NOI = []


def python_cho_cau_noi():
    """Python co `websocket` + `requests`. Do 2026-10-02: `python` trong PATH co the la mot venv
    KHONG co websocket -> cau noi chet o dong import. Anaconda thi co."""
    if _PY_CAU_NOI:
        return _PY_CAU_NOI[0]
    ung = [os.environ.get("CHATGPT_PYTHON"), sys.executable,
           os.path.expanduser(r"~\anaconda3\python.exe"), shutil.which("python")]
    for p in ung:
        if p and os.path.exists(p) and subprocess.run(
                [p, "-c", "import websocket, requests"], capture_output=True).returncode == 0:
            _PY_CAU_NOI.append(p)
            return p
    return None


def goi_chatgpt(prompt, nhan, thu_muc, kieu="hoi", lan=3, nghi=20.0, timeout=900, tam_thoi=True):
    """Mot cau hoi = mot luot chat (Temporary Chat mac dinh). Tra (reply, giay) hoac (None, loi)."""
    py = python_cho_cau_noi()
    if py is None:
        chua_cai("khong Python nao co `websocket-client` + `requests` -- dat CHATGPT_PYTHON "
                 "hoac `pip install websocket-client requests`")
    vao = os.path.join(thu_muc, "_gui_%s.jsonl" % nhan)
    ra = os.path.join(thu_muc, "_nhan_%s.jsonl" % nhan)
    ghi(vao, json.dumps({"id": nhan, "prompt": prompt}, ensure_ascii=False) + "\n")
    lenh = [py, CAU_NOI, vao, ra, "--kieu", kieu, "--timeout", str(timeout)]
    if not tam_thoi:
        lenh.append("--khong-tam-thoi")
    t0, loi = time.time(), ""
    for i in range(lan):
        p = subprocess.run(lenh, capture_output=True, text=True, encoding="utf-8", errors="replace")
        for ln in reversed(doc(ra).strip().split("\n") if os.path.exists(ra) else []):
            try:
                bg = json.loads(ln)
            except ValueError:
                continue
            if bg.get("id") == nhan and bg.get("reply") and not bg.get("error"):
                return bg["reply"], time.time() - t0
            if bg.get("id") == nhan:
                loi = bg.get("error") or "reply rong"
                break
        loi = loi or ((p.stderr or "").strip().split("\n") or ["?"])[-1]
        sys.stderr.write("[chatgpt] lan %d/%d hong: %s\n" % (i + 1, lan, loi[:200]))
        if i + 1 < lan:
            time.sleep(nghi)
            # lan sau hoi lai tu dau: bo ban ghi loi de cau noi khong coi la "da xong"
            if os.path.exists(ra):
                os.unlink(ra)
    return None, loi


def lam_chatgpt(a):
    if a.prompt_tep:
        cau = doc(a.prompt_tep)
    elif a.cau_hoi:
        cau = a.cau_hoi
    else:
        sys.exit("LOI: thieu cau hoi (doi so dau) hoac --prompt-tep")
    prompt = cau + ("\n\n" + gop_tep(a.tep) if a.tep else "")
    nhan = "q" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:4]
    tam = tempfile.mkdtemp(prefix="goi_ngoai_")
    reply, giay_hay_loi = goi_chatgpt(prompt, nhan, tam, kieu=a.kieu, timeout=a.timeout,
                                      tam_thoi=not a.khong_tam_thoi)
    if reply is None:
        nhat_ky("chatgpt", a.nhan or nhan, "chatgpt-web", len(prompt), 0, "HONG " + giay_hay_loi)
        fallback("chatgpt", giay_hay_loi[:200])
    shutil.rmtree(tam, ignore_errors=True)
    ra = a.ra or os.path.join(NHAT_KY_DIR, "tra-loi", nhan + ".md")
    ghi(ra, reply.rstrip() + "\n")
    nhat_ky("chatgpt", a.nhan or nhan, "chatgpt-web", len(prompt), giay_hay_loi, "OK")
    dong = [l for l in reply.strip().split("\n") if l.strip()]
    print("chatgpt · %.0fs · %d ky tu -> %s" % (giay_hay_loi, len(reply), ra))
    for l in dong[:8]:
        print("   " + l[:160])
    if len(dong) > 8:
        print("   ... con %d dong, doc tep tren" % (len(dong) - 8))


# ------------------------------------------------------------------ dich: agy (Gemini)

def tim_agy(mau0):
    for p in (os.environ.get("AGY_EXE"), mau0):
        if p and os.path.exists(p):
            return p
    return shutil.which("agy") or shutil.which(os.path.basename(mau0 or "agy"))


def lam_agy(a):
    if not os.path.exists(AGY_CMD):
        chua_cai("chua co %s -- tu chay viec bang Bash" % AGY_CMD)
    cau_hinh = json.loads(doc(AGY_CMD))
    mau = cau_hinh["lenh"]                           # vd ["agy.exe", "-p", "{prompt}", ...]
    exe = tim_agy(mau[0])
    if exe is None:
        chua_cai("khong thay agy (AGY_EXE, %s, PATH) -- tu chay viec bang Bash" % mau[0])
    if not a.tep:
        sys.exit("LOI: agy can it nhat mot --tep (thu muc lam viec = thu muc chung cua cac tep)")
    prompt = MAU_AGY.replace("{{VIEC}}", a.viec).replace(
        "{{TEP}}", "\n".join("- " + os.path.abspath(t) for t in a.tep))
    # agy lay thu muc chay lam workspace -> dat vao thu muc chung cua cac tep duoc phep dung
    td = a.thu_muc or cau_hinh.get("thu_muc") or os.path.commonpath(
        [os.path.dirname(os.path.abspath(t)) for t in a.tep])
    # prompt de NGOAI thu muc lam viec: thu muc ay chi nen chua tep cua viec
    tep_prompt = os.path.join(tempfile.gettempdir(), "_agy_prompt_%s.md" % uuid.uuid4().hex[:8])
    ghi(tep_prompt, prompt)
    kieu = "suy-luan" if a.suy_luan else "code"
    model = (cau_hinh.get("model") or {}).get(kieu, "")
    lenh = [exe] + [x.replace("{prompt}", prompt).replace("{prompt_file}", tep_prompt)
                    .replace("{model}", model) for x in mau[1:]]
    if "--model" in lenh and not model:                    # khong co model thi de agy tu chon
        i = lenh.index("--model")
        del lenh[i:i + 2]
    t0 = time.time()
    try:
        p = subprocess.run(lenh, capture_output=True, text=True, encoding="utf-8", errors="replace",
                           timeout=int(cau_hinh.get("timeout", 1800)), cwd=td,
                           input=prompt if cau_hinh.get("stdin") else None)
    except subprocess.TimeoutExpired:
        nhat_ky("agy", a.nhan, "gemini-agy:" + model, len(prompt), time.time() - t0, "HONG timeout")
        fallback("agy", "het gio %ss" % cau_hinh.get("timeout", 1800))
    giay = time.time() - t0
    tra = p.stdout or ""
    phieu = tach_phieu(tra)
    if p.returncode or phieu is None:
        nhat_ky("agy", a.nhan, "gemini-agy:" + model, len(prompt), giay, "HONG rc=%d" % p.returncode)
        fallback("agy", "agy thoat %d, %s" % (p.returncode,
                 "khong co khoi ###PHIEU###" if phieu is None else (p.stderr or "")[-150:]))
    nhat_ky("agy", a.nhan, "gemini-agy:" + model, len(prompt), giay, "OK")
    kq = (phieu.get("KETQUA") or [["?", ""]])[0]
    print("agy · %s · %.0fs · %s: %s" % (model or "agy mac dinh", giay, kq[0], kq[1] if len(kq) > 1 else ""))
    for h in phieu.get("TEP", [])[:8]:
        print("   %s  %s" % (h[0], h[1] if len(h) > 1 else ""))
    if kq[0].upper() != "XONG":
        sys.exit(1)


# ------------------------------------------------------------------ tu kiem

def tu_kiem():
    """Ca kiem khong gui gi ra ngoai: tach khuon, cau hinh, cong cu san sang."""
    hong = 0

    def ok(ten, dung, ghi_chu="", bat_buoc=True):
        nonlocal hong
        hong += 0 if (dung or not bat_buoc) else 1
        nhan = "DAT " if dung else ("HONG" if bat_buoc else "CHUA")
        print("[tu-kiem] %-30s %s %s" % (ten, nhan, ghi_chu))

    tra = ("noi dung thua\n```\n###PHIEU###\n[KETQUA]\nXONG | da sua 2 tep\n"
           "[TEP]\n| a.py | doi ten ham |\nb.py | them import\n###HET###\n```")
    p = tach_phieu(tra)
    ok("tach_phieu KETQUA", p is not None and p["KETQUA"][0] == ["XONG", "da sua 2 tep"])
    ok("tach_phieu bo | dau-cuoi", p is not None and p["TEP"][0][0] == "a.py")
    ok("tach_phieu thieu khoi", tach_phieu("khong co gi") is None)
    ok("mau agy du cho trong", "{{VIEC}}" in MAU_AGY and "{{TEP}}" in MAU_AGY
       and "###PHIEU###" in MAU_AGY)
    ok("cau noi chatgpt ton tai", os.path.exists(CAU_NOI))
    try:
        ch = json.loads(doc(AGY_CMD))
        ok("agy-cmd.json doc duoc", "{prompt}" in " ".join(ch["lenh"]), AGY_CMD)
        exe = tim_agy(ch["lenh"][0])
        ok("agy san sang", exe is not None, exe or "khong thay -> lenh agy se tra ma 3",
           bat_buoc=False)
    except Exception as e:
        ok("agy-cmd.json doc duoc", False, str(e)[:80])
    py = python_cho_cau_noi()
    ok("Python co websocket-client", py is not None, py or "pip install websocket-client requests",
       bat_buoc=False)
    try:
        import urllib.request
        with urllib.request.urlopen("http://127.0.0.1:9222/json/version", timeout=2) as r:
            trinh = json.loads(r.read().decode("utf-8")).get("Browser", "?")
        ok("Chrome debug cong 9222", True, trinh, bat_buoc=False)
    except Exception:
        ok("Chrome debug cong 9222", False, "chua mo -- xem SKILL.md muc Chuan bi", bat_buoc=False)
    print("KET QUA: %s" % ("tat ca dat." if not hong else "%d HONG" % hong))
    return hong


# ------------------------------------------------------------------ main

def _utf8_console():
    for luong in (sys.stdout, sys.stderr):
        try:
            luong.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def main():
    _utf8_console()
    if "--tu-kiem" in sys.argv:
        sys.exit(1 if tu_kiem() else 0)
    ap = argparse.ArgumentParser(description="Goi ChatGPT web / Gemini (agy) ngoai Claude")
    sub = ap.add_subparsers(dest="dich", required=True)

    sp = sub.add_parser("chatgpt", help="hoi ChatGPT web qua Chrome cong 9222")
    sp.add_argument("cau_hoi", nargs="?")
    sp.add_argument("--prompt-tep", help="doc cau hoi tu tep (prompt dai, nhieu dong)")
    sp.add_argument("--tep", action="append", default=[], help="dan noi dung tep vao sau cau hoi")
    sp.add_argument("--ra", help="tep ghi nguyen van tra loi")
    sp.add_argument("--kieu", choices=("hoi", "cham"), default="hoi",
                    help="cham: bat buoc khoi ###PHIEU###...###HET### trong tra loi")
    sp.add_argument("--timeout", type=int, default=900)
    sp.add_argument("--khong-tam-thoi", action="store_true", help="khong dung Temporary Chat")
    sp.add_argument("--nhan", default="", help="nhan ghi vao nhat ky")

    sp = sub.add_parser("agy", help="giao viec co khi cho Gemini qua Antigravity CLI")
    sp.add_argument("viec")
    sp.add_argument("--tep", action="append", default=[])
    sp.add_argument("--thu-muc", help="thu muc lam viec cua agy (mac dinh: thu muc chung cua --tep)")
    sp.add_argument("--suy-luan", action="store_true",
                    help="viec can suy luan (doc log tim nguyen nhan, dung .drawio) -> Gemini 3.1 Pro; "
                         "mac dinh la viec code -> Gemini 3.8 Flash")
    sp.add_argument("--nhan", default="-", help="nhan ghi vao nhat ky")

    a = ap.parse_args()
    {"chatgpt": lam_chatgpt, "agy": lam_agy}[a.dich](a)


if __name__ == "__main__":
    main()
