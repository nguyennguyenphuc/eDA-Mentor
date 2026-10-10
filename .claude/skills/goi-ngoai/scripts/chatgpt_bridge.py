# -*- coding: utf-8 -*-
"""Cau noi ChatGPT web: day tung doan van sang chatgpt.com, mang ban viet lai ve.

Cach dung:
    python chatgpt_bridge.py --selftest [--live]        # kiem truoc khi chay that
    python chatgpt_bridge.py vao.jsonl ra.jsonl         # chay 1 luong, 1 doan/luot
    python chatgpt_bridge.py vao.jsonl ra.jsonl --lo 8 --workers 3   # nhanh (~8x * 3)

BAN SKILL `goi-ngoai` (chep tu academic-writing-v3 ngay 2026-10-10, ban goc khong sua). Ba kieu:
  --kieu viet-lai (mac dinh) reply la ban viet lai cua doan van -> kiem do dai so voi doan goc.
  --kieu cham   reply la PHIEU CHAM, bat buoc khoi ###PHIEU### ... ###HET###, bo kiem do dai.
  --kieu hoi    hoi dap tu do (goi qua `goi_ngoai.py chatgpt`): chi kiem reply khong rong va
                khong bat dau bang tien to la (chu cua nut bi boc nham).
  Selector theo giao dien 2026-10-02 (bo data-message-author-role) va don tep dinh kem sot
                lai truoc khi go. Dan prompt dai la DUONG CUT, xem ghi chu o `go_prompt`.

Dinh dang vao (JSONL, moi dong mot doan):  {"id": "p001", "prompt": "<van ban>"}
Dinh dang ra  (JSONL, ghi noi tiep):       {"id": "p001", "reply": "...", "elapsed": 8.4}
  Ban ghi bi nghi hong mang thêm truong "error" va lam script exit khac 0.

DUONG GUI MAC DINH LA GUI TRAN: dung doan van, khong mot chu huong dan nao. Da chay that
2026-09-05 tren 48 doan cua bai Power BI v3 theo dung cach nay, 0 loi, va cong A/B ba phieu
cham mu cho thay 6 / giu 16 voi so lieu va thuat ngu khop ban goc 100%. Gui tran giu duoc
dung cai dang can: mot giong thu hai CHUA bi nhiem checklist cua skill. `--brief` chi bat
cho mot lo cu the khi lo ay tra ve van day colon hoac em-dash.
    Dong co khoa "error" duoc coi la CHUA XONG -> lan chay sau tu dong lam lai dung nhung id do.

CHUAN BI MOT LAN (nguoi dung tu lam, script khong tu mo Chrome):
    chrome.exe --remote-debugging-port=9222 ^
        --user-data-dir="%LOCALAPPDATA%\\chatgpt-bridge-profile"
    roi dang nhap chatgpt.com trong cua so do va de tab do mo.
    BAT BUOC dung --user-data-dir rieng: Chrome >=136 tu choi mo port debug tren profile mac
    dinh (may nay dang chay Chrome 152).

Ky thuat (da kiem chung 2026-09-04 tren may nay):
- `websocket-client` 0.58.0: `recv()` KHONG nhan tham so timeout -> phai dat bang `ws.settimeout()`;
  `recv()` tra ve chuoi RONG khi nhan frame CLOSE, khong phai khi tin nhan rong.
- `suppress_origin=True` la BAT BUOC. websocket-client mac dinh gui header `Origin`, va Chrome
  >=111 tra 403 cho handshake DevTools khi co header do.
- KHONG goi `Runtime.enable` / `Page.enable` / bat ky `*.enable` nao. Cac lenh ta dung
  (`Runtime.evaluate`, `Input.insertText`, `Input.dispatchKeyEvent`, `Page.navigate`) deu chay
  duoc khi domain chua bat, va nho vay socket khong co luong su kien rac. Van loc theo `id`.
- `Emulation.setFocusEmulationEnabled` giu `document.hasFocus()` luon dung, nen composer van nhan
  chu khi cua so Chrome bi che hoac thu nho.
- Chen chu: `Input.insertText` di qua duong IME that -> ProseMirror nhan duoc. Xuong dong phai
  dispatch Shift+Enter rieng, KHONG nhet "\\n" vao insertText.
- Biet da tra loi xong = 2 tin hieu AND: nut Stop bien mat VA chu ky noi dung dung yen qua
  `--on-dinh` lan poll. Cong thu ba: so node assistant phai tang so voi truoc luc gui.
- Nut send dung `data-testid="send-button"` va CHI ton tai khi composer da co chu. `--selftest`
  vi vay phai go thu mot ky tu truoc khi kiem, khong thi no bao HONG oan.
- **Writing block**: ChatGPT boc van xuoi trong `data-testid="writing-block-container"`, va header
  DINH cua khoi nay nam BEN TRONG chinh `.markdown`, mang nut "Edit". Doc thang
  `.markdown`.innerText se dinh tien to "Edit" o MOI cau tra loi. Noi dung that o node
  `[class*="_editor"]`. Da do that: 5/5 doan dinh "Edit" truoc khi sua, 0/5 sau khi sua.

Da chay that 2026-09-04: `--selftest --live` dat (5,5s, gui bang button); 5 doan van that
15,7s/doan, 0 loi; colon giua cau 2->0, 2->0, 1->0 o ba doan co loi that. CHUA kiem duoc:
cau chu thong bao rate-limit (chua cham gioi han) va che do gop lo `--lo` tren dulieu that.

DUONG CUT -- DUNG THU LAI:
- backend-api /conversation: can sentinel proof-of-work token do JS cua trang sinh ra, khong tai
  tao duoc ngoai trinh duyet.
- Gan `.innerText` cho composer: KHONG co tac dung. ProseMirror dung document tu MutationObserver,
  khong doc thuoc tinh.
- Dispatch `InputEvent` tong hop mot minh: cung khong chen duoc gi (isTrusted=false, khong sinh
  mutation that). Su kien `paste` tong hop THI duoc, va dang lam duong du phong.
- `time.sleep()` co dinh de doi tra loi: cat cut cau tra loi. Khong dung.
- MCP/tunnel: da can nhac va loai. Xem `chatgpt-style-brief.md` va ke hoach kem theo.

GIOI HAN: script nay phu thuoc DOM cua chatgpt.com va se gay khi OpenAI doi giao dien.
`--selftest` in ra selector nao con song va khop o ung vien thu may.
"""
import io
import os
import re
import sys
import json
import time
import random
import argparse
import threading

import requests
import websocket

# ==================== SELECTOR -- SUA O DAY VA CHI O DAY ====================
# Moi muc la DANH SACH ung vien, thu tu cu the -> tong quat.
SEL = {
    # composer la div contenteditable cua ProseMirror, KHONG phai <textarea>   [2026-09-04]
    "composer": ["#prompt-textarea",
                 "div.ProseMirror[contenteditable='true']",
                 "form div[contenteditable='true']"],
    # KIEM CHUNG 2026-09-04: dung `send-button`, va nut nay CHI ton tai khi composer co chu.
    "send": ['[data-testid="send-button"]', '[data-testid="composer-send-button"]',
             'button[aria-label*="Send prompt" i]', 'button[aria-label*="Send" i]',
             'form button[type="submit"]'],
    # SU HIEN DIEN cua nut stop = dang stream                                  [2026-09-04]
    #
    # SUA 2026-09-16 -- ung vien cu `button[aria-label*="Dung" i]` khop hai muc
    # LICH SU CHAT o thanh ben, vi ten chat "Bien tap noi dung" chua chu "dung" va
    # co `i` bo qua hoa thuong. Hai phan tu ay THUONG TRU, nen dieu kien ket thuc
    # `not p["stop"]` khong bao gio dung va cau noi luon tuong dang stream: mot cau
    # tra loi "OK" mot chu mat 121 giay. Hai luat rut ra:
    #   1. Moi ung vien cua nhom `stop` phai bi RANG BUOC VAO KHUNG SOAN (`form`),
    #      vi nut dung chi song trong do. Khong bao gio de mot selector trong nhom
    #      nay quet ca trang.
    #   2. Chuoi tieng Viet phai viet DUNG DAU. `*="Dung"` khong khop "Dung" ma khop
    #      "noi dung"; muon bat nut that thi phai la "Dung".
    "stop": ['[data-testid="stop-button"]', 'form button[data-testid*="stop" i]',
             'form button[aria-label*="Stop" i]',
             'form button[aria-label*="Dừng" i]'],
    # SUA 2026-10-02 -- giao dien moi BO `data-message-author-role`, `article`, `.markdown`
    # (dem ra 0 ca ba). Moi luot tra loi nam trong mot don vi tim kiem co khoa ket thuc
    # bang ":assistant"; chu nam trong khoi `data-markdown-text-style="assistant-message"`.
    # Do that: --live gui "OK", ChatGPT tra "OK", cau noi doi 75s roi bao khong thay luot moi.
    # Ung vien cu giu lai phia sau phong khi giao dien quay ve. Lay ung vien DAU TIEN co
    # ket qua, khong gop: gop hai kieu se dem doi so luot.
    "assistant": ['[data-content-search-unit-key$=":assistant"]',
                  '[data-message-author-role="assistant"]'],
    "markdown": ['[data-markdown-text-style="assistant-message"]', '.markdown',
                 '[data-message-content]'],
    # KHONG dung `button[aria-label*="New chat" i]`: khop "New chat in <project>" (2026-10-02).
    # chat_moi() nay luon dieu huong; nhom nay chi con cho --selftest bao giao dien con song.
    "newchat": ['[data-testid="create-new-chat-button"]',
                'button[aria-label="New chat" i]', 'a[href="/"]'],
}

# CAN KIEM CHUNG LAI voi UI thuc te lan dau cham gioi han -- day moi la phong doan.
RATE_PAT = re.compile(
    r"(you've (reached|hit) (our|the|your)|reached our limit of messages|plan limit for|"
    r"usage (cap|limit)|rate limit|please try again (later|after)|"
    r"you can (continue|send more) (messages )?(after|at)|"
    r"bạn đã đạt giới hạn|vui lòng thử lại sau|đã hết lượt)", re.I)
ERR_PAT = re.compile(r"(something went wrong|an error occurred|network error|"
                     r"đã có lỗi xảy ra)", re.I)

TACH = "###DOAN %s###"
TACH_RE = re.compile(r"#{2,4}\s*DOAN\s+([A-Za-z0-9_.-]+)\s*#{2,4}", re.I)


class RateLimited(RuntimeError):
    pass


def Q(ten):
    """Lay node dau tien khop theo THU TU UNG VIEN, khong phai thu tu DOM.

    KHONG duoc quay ve `document.querySelector("a, b, c")`: dang do tra ve node dau tien theo
    THU TU TAI LIEU. Da gay loi that 2026-09-04: khi cau tra loi truoc do la mot "writing block"
    (o soan thao co contenteditable, nam TREN composer trong DOM), ung vien
    `div.ProseMirror[contenteditable='true']` khop trung o soan thao ay, nen prompt duoc go vao
    BAI VIET cua ChatGPT chu khong vao composer -- va `go_prompt` van bao thanh cong vi no doc
    lai dung cai node sai do.
    """
    return ("(() => { for (const s of %s) { const e = document.querySelector(s); if (e) return e; }"
            " return null; })()" % json.dumps(SEL[ten]))


Q_thu_tu = Q          # giu ten cu, cung mot nghia


def QA(ten):
    """Nhu `Q` nhung tra MOI node cua ung vien DAU TIEN co ket qua (mang JS)."""
    return ("(() => { for (const s of %s) { const r = document.querySelectorAll(s);"
            " if (r.length) return Array.from(r); } return []; })()" % json.dumps(SEL[ten]))


def QT(ten, goc):
    """Node dau tien BEN TRONG `goc` (bien JS) khop theo thu tu ung vien."""
    return ("(() => { for (const s of %s) { const e = %s.querySelector(s); if (e) return e; }"
            " return null; })()" % (json.dumps(SEL[ten]), goc))


# ------------------------------------------------------------------ CDP

def _http():
    s = requests.Session()
    s.trust_env = False          # bo qua proxy he thong cho 127.0.0.1
    return s


def tim_tab(port, muon="chatgpt.com"):
    try:
        tabs = _http().get("http://127.0.0.1:%d/json" % port, timeout=6).json()
    except requests.exceptions.RequestException as e:
        raise RuntimeError(
            "khong ket noi duoc CDP 127.0.0.1:%d (%s).\n"
            "Chrome phai dang chay voi --remote-debugging-port=%d VA --user-data-dir rieng."
            % (port, e.__class__.__name__, port))
    trang = [t for t in tabs if t.get("type") == "page" and muon in (t.get("url") or "")]
    if not trang:
        mo = ", ".join((t.get("url") or "?")[:60] for t in tabs if t.get("type") == "page")
        raise RuntimeError("khong tim thay tab %s; dang mo: %s" % (muon, mo or "(khong co)"))
    return trang


def mo_tab_moi(port, url="https://chatgpt.com/?temporary-chat=true"):
    s = _http()
    r = s.put("http://127.0.0.1:%d/json/new?%s" % (port, url), timeout=10)
    if r.status_code == 405:
        r = s.get("http://127.0.0.1:%d/json/new?%s" % (port, url), timeout=10)
    if not r.ok:
        raise RuntimeError("khong mo duoc tab moi: http %s" % r.status_code)
    return r.json()


def cdp_connect(target):
    ws = websocket.create_connection(
        target["webSocketDebuggerUrl"], timeout=30,
        suppress_origin=True,                       # BAT BUOC, xem docstring
        http_no_proxy=["127.0.0.1", "localhost"],
        enable_multithread=False)                   # moi luong mot socket rieng
    try:
        cdp_send(ws, "Emulation.setFocusEmulationEnabled", {"enabled": True})
    except RuntimeError:
        pass                                        # khong sao, chi can giu cua so hien
    return ws


_dem = [0]
_dem_khoa = threading.Lock()


def cdp_send(ws, method, params=None, timeout=30):
    with _dem_khoa:
        _dem[0] += 1
        mid = _dem[0]
    ws.send(json.dumps({"id": mid, "method": method, "params": params or {}}))
    ws.settimeout(timeout)
    han = time.time() + timeout
    while True:
        try:
            raw = ws.recv()
        except websocket.WebSocketTimeoutException:
            raise RuntimeError("CDP het gio (%ss) khi goi %s" % (timeout, method))
        if raw == "":
            raise RuntimeError("CDP socket dong khi goi %s (tab bi dong?)" % method)
        msg = json.loads(raw)
        if "id" not in msg:
            if time.time() > han:
                raise RuntimeError("CDP het gio (chi nhan su kien) khi goi %s" % method)
            continue
        if msg["id"] != mid:
            continue
        if "error" in msg:
            raise RuntimeError("CDP %s loi: %s" % (method, msg["error"]))
        return msg.get("result", {})


def cdp_eval(ws, expr, timeout=30):
    r = cdp_send(ws, "Runtime.evaluate",
                 {"expression": expr, "returnByValue": True,
                  "awaitPromise": True, "userGesture": True}, timeout=timeout)
    if r.get("exceptionDetails"):
        d = r["exceptionDetails"]
        raise RuntimeError("JS loi: %s" % ((d.get("exception") or {}).get("description") or d.get("text")))
    return (r.get("result") or {}).get("value")


SHIFT = 8


def cdp_key(ws, key, vk, code, modifiers=0, text=""):
    goc = {"key": key, "code": code, "windowsVirtualKeyCode": vk,
           "nativeVirtualKeyCode": vk, "modifiers": modifiers}
    cdp_send(ws, "Input.dispatchKeyEvent",
             dict(goc, type="keyDown" if text else "rawKeyDown", text=text, unmodifiedText=text))
    cdp_send(ws, "Input.dispatchKeyEvent", dict(goc, type="keyUp"))


# ------------------------------------------------------------- thao tac trang

JS_FOCUS = """(() => {
  const el = %s;
  if (!el) return 'no-composer';
  el.focus();
  const sel = window.getSelection(), r = document.createRange();
  r.selectNodeContents(el); r.collapse(false);
  sel.removeAllRanges(); sel.addRange(r);
  return 'ok';
})()"""

JS_LOGIN = """(() => {
  if (/\\/auth\\/(login|signin)|\\/login/.test(location.href)) return 'login-url';
  const c = %s;
  const t = (document.body.innerText || '').slice(0, 4000);
  if (/(Verify you are human|Checking your browser|Just a moment)/.test(t)) return 'cloudflare';
  if (!c && /(Log in|Sign up|Stay logged out|Đăng nhập)/.test(t)) return 'login-wall';
  return c ? 'ok' : 'no-composer';
})()"""

JS_PASTE = """(() => {
  const el = %s;
  if (!el) return null;
  el.focus();
  document.execCommand('selectAll');
  const dt = new DataTransfer();
  dt.setData('text/plain', %s);
  el.dispatchEvent(new ClipboardEvent('paste', {clipboardData: dt, bubbles: true, cancelable: true}));
  return el.innerText;
})()"""

# KIEM CHUNG 2026-09-04: ChatGPT boc van xuoi trong "writing block"
# (`data-testid="writing-block-container"`), co mot header DINH ben trong chinh `.markdown`
# chua nut "Edit". Doc thang `.markdown`.innerText se dinh tien to "Edit" o MOI cau tra loi,
# va co the dinh ca chu thich o cuoi. Noi dung that nam o node co class chua "_editor".
# Doan chen lai dau nhay nguoc quanh <code> duoi day la LUOI AN TOAN CHUA DUOC THU, doc ky
# truoc khi tin no:
#
# Do that tren bai Game 2026-09-10: 49/49 ban tra ve co 0 dau nhay nguoc, va 35/49 doan bi cong
# A/B danh truot vi "lech thuat ngu". Gia thuyet dau tien la `innerText` cua ban DA RENDER lam
# mat dau nhay. Da vá theo gia thuyet ay roi chay lai d.253 (goc 16 nhay) va d.327 (goc 22
# nhay): van 0 nhay, TRONG KHI cac dinh danh (`liveops_events`, `daily_kpi`) van co mat duoi
# dang chu tron. Tuc nguyen nhan that la CHATGPT viet dinh danh thanh chu tron, khong phai bo
# boc lam mat dinh dang. Gia thuyet dau tien SAI.
#
# Giu lai doan nay vi no dung neu ChatGPT co luc tra ve <code> that, nhung PHAI ghi ro: no chua
# tung chay qua mot ca nao co <code> trong cau tra loi, nen chua duoc tinh la da kiem chung
# (S54: moi muc trong docstring phai co mot ca trong tu_kiem). Dung noi rang no "da chua loi
# lech thuat ngu" — no chua chua gi ca.
#
# Ban sao la bat buoc: sua node that thi lan poll sau doc ra van ban da bi doi, va bo phat hien
# "noi dung dung yen" se hieu nham la con dang stream.
JS_EXTRACT = """(() => {
  const A = __QA_ASSISTANT__;
  const last = A[A.length - 1];
  if (!last) return null;
  const md = __QT_MARKDOWN__ || last;
  const wb = md.querySelector('[data-testid="writing-block-container"]');
  if (wb) {
    const ed = wb.querySelector('[class*="_editor"]');
    if (ed && (ed.innerText || '').trim()) return ed.innerText;
  }
  const ban = md.cloneNode(true);
  ban.querySelectorAll('[data-testid*="writing-block-header"], button, [role="button"]')
     .forEach(e => e.remove());
  ban.querySelectorAll('pre code').forEach(e => {
    e.textContent = '\\n```\\n' + e.textContent + '\\n```\\n';
  });
  ban.querySelectorAll('code').forEach(e => {
    const s = e.textContent || '';
    if (s.indexOf('```') === -1) e.textContent = '`' + s + '`';
  });
  // Ban sao chua gan vao tai lieu nen `innerText` co the tra ve rong; gan tam ngoai man hinh.
  ban.style.position = 'fixed'; ban.style.left = '-99999px'; ban.style.top = '0';
  document.body.appendChild(ban);
  const t = ban.innerText;
  ban.remove();
  return t;
})()""".replace("__QA_ASSISTANT__", QA("assistant")).replace("__QT_MARKDOWN__", QT("markdown", "last"))

JS_BODY_DUOI = "(document.body.innerText || '').slice(-3000)"


def js_probe():
    return """(() => {
  const A = __QA_ASSISTANT__;
  const last = A[A.length - 1] || null;
  const md = last ? (__QT_MARKDOWN__ || last) : null;
  const t = md ? md.innerText : '';
  let h = 0; for (let i = 0; i < t.length; i++) h = (h * 31 + t.charCodeAt(i)) | 0;
  return JSON.stringify({n: A.length, len: t.length, h: h,
                         stop: !!%s, cls: !!document.querySelector('.result-streaming, [data-is-streaming="true"]')});
})()""".replace("__QA_ASSISTANT__", QA("assistant")).replace(
        "__QT_MARKDOWN__", QT("markdown", "last")) % Q("stop")


def kiem_dang_nhap(ws):
    r = cdp_eval(ws, JS_LOGIN % Q("composer"))
    if r == "cloudflare":
        raise RuntimeError("Cloudflare dang chan -- mo tab len bam qua thu thach roi chay lai")
    if r in ("login-url", "login-wall"):
        raise RuntimeError("chua dang nhap chatgpt.com trong profile nay -- dang nhap 1 lan roi chay lai")
    if r != "ok":
        raise RuntimeError("khong thay composer tren trang (DOM doi? chay --selftest)")


def chu_composer(ws):
    return cdp_eval(ws, "(() => { const e = %s; return e ? e.innerText : null; })()" % Q("composer"))


# DUONG CUT 2026-10-02 -- DUNG THU LAI: dan (paste) prompt dai de go cho nhanh. Giao dien moi
# bien doan dan dai thanh TEP DINH KEM "Pasted text.txt"; nut send `disabled` trong luc tep
# tai len, va tep dinh kem SONG QUA "New chat" (6 lan thu de lai 6 tep trong o soan). Go tung
# dong (duoi day) moi la duong chinh; JS_PASTE chi con la du phong khi go rot.
JS_DON_DINH_KEM = """(() => { let n = 0;
  for (const b of document.querySelectorAll('form button[aria-label^="Remove"]')) { b.click(); n++; }
  return n; })()"""


def go_prompt(ws, text):
    if cdp_eval(ws, JS_FOCUS % Q("composer")) != "ok":
        raise RuntimeError("khong tim thay composer -- chay --selftest")
    if cdp_eval(ws, JS_DON_DINH_KEM):
        time.sleep(0.5)
    cdp_eval(ws, "document.execCommand('selectAll'), 1")
    cdp_key(ws, "Backspace", 8, "Backspace")
    dong = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    for i, d in enumerate(dong):
        if i:
            cdp_key(ws, "Enter", 13, "Enter", modifiers=SHIFT, text="\r")   # xuong dong, KHONG gui
        if d:
            cdp_send(ws, "Input.insertText", {"text": d})
    duoc = (chu_composer(ws) or "").strip()
    if not duoc:
        duoc = (cdp_eval(ws, JS_PASTE % (Q("composer"), json.dumps(text))) or "").strip()
    if not duoc:
        raise RuntimeError("composer khong nhan text (ca insertText lan paste deu that bai)")
    if len(duoc) < 0.8 * len(text.strip()):
        raise RuntimeError("composer chi nhan %d/%d ky tu -- nghi bi cat" % (len(duoc), len(text.strip())))


def gui(ws):
    r = cdp_eval(ws, """(() => { const b = %s;
        if (!b) return 'no-button';
        if (b.disabled || b.getAttribute('aria-disabled') === 'true') return 'disabled';
        b.click(); return 'clicked'; })()""" % Q_thu_tu("send"))
    time.sleep(0.8)
    if (chu_composer(ws) or "").strip() == "":
        return "button"
    cdp_eval(ws, JS_FOCUS % Q("composer"))
    cdp_key(ws, "Enter", 13, "Enter", text="\r")
    time.sleep(0.8)
    if (chu_composer(ws) or "").strip() == "":
        return "enter"
    raise RuntimeError("khong gui duoc prompt (nut send: %s) -- chay --selftest" % r)


#: Vong doi gan day co THAY nut stop luc dang stream khong. `--selftest --live`
#: doc bien nay. Cong nay canh CHIEU NGUOC cua loi selector `stop`: neu khong ung
#: vien nao khop luc dang stream thi dieu kien `not p["stop"]` dung ngay tu vong
#: dau, va cau noi co the tra ve ban CUT ma khong bao gi -- dung lop E35.
DA_THAY_STOP = False


def doi_tra_loi(ws, n0, timeout=300.0, poll=0.6, on_dinh=4, an_han=75.0):
    global DA_THAY_STOP
    DA_THAY_STOP = False
    t0 = time.time()
    bat_dau, yen, ky = False, 0, None
    while True:
        p = json.loads(cdp_eval(ws, js_probe(), timeout=20))
        if p["stop"]:
            DA_THAY_STOP = True
        # Nut stop hien = da nhan va dang nghi. Do 2026-10-02 o muc High: prompt 49k ky tu
        # nghi qua 75s ma CHUA hien don vi tra loi -> bao "khong thay luot moi" oan.
        if p["n"] > n0 or p["stop"]:
            bat_dau = True
        elif not bat_dau and time.time() - t0 > an_han:
            raise RuntimeError("khong thay luot tra loi moi sau %.0fs (rate limit? mat mang?)" % an_han)
        if bat_dau:
            hien = (p["n"], p["len"], p["h"])
            yen = yen + 1 if hien == ky else 0
            ky = hien
            if (not p["stop"]) and (not p["cls"]) and p["len"] > 0 and yen >= on_dinh:
                return "ok"
        if time.time() - t0 > timeout:
            return "timeout"
        time.sleep(poll)


def boc_tra_loi(ws):
    t = cdp_eval(ws, JS_EXTRACT) or ""
    t = t.replace("\r\n", "\n").strip()
    return re.sub(r"\n{3,}", "\n\n", t)


def _cho_chat_trong(ws, giay):
    han = time.time() + giay
    while time.time() < han:
        try:
            p = json.loads(cdp_eval(ws, js_probe(), timeout=15))
            san = cdp_eval(ws, "document.readyState === 'complete' && !!%s" % Q("composer"))
            if p["n"] == 0 and san:
                return True
        except RuntimeError:
            pass
        time.sleep(0.7)
    return False


def chat_moi(ws, tam_thoi=True, manh=True):
    # SUA 2026-10-02 -- LUON dieu huong, KHONG bam nut o thanh ben (tham so `manh` giu cho
    # tuong thich, khong con tac dung). Giao dien moi co nut "New chat in <ten project>" cho
    # MOI project; ung vien `button[aria-label*="New chat" i]` khop nut cua project DAU TIEN
    # ("Nghiep vu dau thau"), nen moi luot cham bi tao BEN TRONG project cua nguoi dung.
    # Dieu huong thang toi chatgpt.com/?temporary-chat=true thi ra chat tam ngoai moi project.
    cdp_send(ws, "Page.navigate",
             {"url": "https://chatgpt.com/" + ("?temporary-chat=true" if tam_thoi else "")})
    if not _cho_chat_trong(ws, 45):
        raise RuntimeError("mo chat moi xong van khong thay composer -- chua dang nhap?")
    return "navigate"


def hoi(ws, prompt, timeout=300.0, on_dinh=4, tam_thoi=True, manh=False):
    chat_moi(ws, tam_thoi, manh)
    t0 = time.time()
    go_prompt(ws, prompt)
    cach = gui(ws)
    tt = doi_tra_loi(ws, 0, timeout=timeout, on_dinh=on_dinh)
    tra = boc_tra_loi(ws)
    duoi = cdp_eval(ws, JS_BODY_DUOI) or ""
    if RATE_PAT.search(tra) or RATE_PAT.search(duoi):
        raise RateLimited("ChatGPT bao gioi han: %s" % (tra[:200] or duoi[-200:]))
    if tt == "timeout":
        return tra, {"gui": cach, "elapsed": time.time() - t0, "error": "timeout sau %.0fs" % timeout}
    if not tra:
        return tra, {"gui": cach, "elapsed": time.time() - t0, "error": "tra loi rong"}
    if ERR_PAT.search(tra) and len(tra) < 200:
        return tra, {"gui": cach, "elapsed": time.time() - t0, "error": "trang bao loi"}
    return tra, {"gui": cach, "elapsed": time.time() - t0}


# ------------------------------------------------------------------ gom lo

# Tien to la ma ChatGPT tra ve khi bo boc noi dung an nham vao nut/thanh cong cu cua trang
# thay vi lay chu cua cau tra loi. Loi E35: ca 5/5 ban ghi dau tien ve kem tien to "Edit" ma
# script van exit 0 va bao "0 loi, 5/5 ban ghi". Hong im lang la loai nguy hiem nhat, nen day
# la mot cua CHAN, khong phai canh bao.
TIEN_TO_LA = ("Edit", "Copy", "Sao chép", "Chỉnh sửa", "Retry", "Thử lại", "Share", "Chia sẻ")


def kiem_ban_ghi(prompt, reply):
    """Tra ve chuoi mo ta loi, hoac None neu ban ghi trong sach.

    Hai phep kiem, deu tu loi that:
    - tien to la: bo boc DOM an nham chu cua nut (E35).
    - do dai lech qua xa: cau tra loi bi cat giua dong, hoac chi la mot manh giao dien.
      Nguong +-60% du rong de ChatGPT viet lai dai/ngan hon that su, du chat de bat manh vun.
    """
    r = (reply or "").strip()
    if not r:
        return "reply rong"
    for t in TIEN_TO_LA:
        if r.startswith(t):
            return "reply bat dau bang tien to la %r -- nghi boc nham chu cua nut (E35)" % t
    n_p, n_r = len(prompt.strip()), len(r)
    if n_p >= 60 and (n_r < n_p * 0.4 or n_r > n_p * 2.5):
        return "do dai lech qua xa: prompt %d ky tu, reply %d ky tu" % (n_p, n_r)
    return None


#: Khoi phieu ma moi mau trong `mau-prompt/` bat ChatGPT in ra (v3, 2026-10-02).
PHIEU_RE = re.compile(r"#{3}\s*PHIEU\s*#{3}(.*?)#{3}\s*HET\s*#{3}", re.S | re.I)


def kiem_phieu(reply):
    """Kieu `cham`: phieu cham KHONG cung do dai voi doan goc, nen bo phep kiem do dai
    cua `kiem_ban_ghi`. Thay bang: phai co khoi ###PHIEU### ... ###HET### va khong bat dau
    bang tien to la (E35 van ap dung)."""
    r = (reply or "").strip()
    if not r:
        return "reply rong"
    for t in TIEN_TO_LA:
        if r.startswith(t):
            return "reply bat dau bang tien to la %r -- nghi boc nham chu cua nut (E35)" % t
    if not PHIEU_RE.search(r):
        return "khong co khoi ###PHIEU### ... ###HET### -- bi cat giua chung hoac sai khuon"
    return None


def kiem_hoi(reply):
    """Kieu `hoi`: khong co khuon nao de kiem, chi chan reply rong va tien to la (E35)."""
    r = (reply or "").strip()
    if not r:
        return "reply rong"
    for t in TIEN_TO_LA:
        if r.startswith(t):
            return "reply bat dau bang tien to la %r -- nghi boc nham chu cua nut (E35)" % t
    return None


def gom_lo(muc, brief):
    """Ghep nhieu doan thanh MOT prompt, danh so bang dau phan cach cung."""
    p = [brief.strip(), "",
         "Duoi day la %d doan. Voi MOI doan, in ra dung khuon sau, khong them loi dan," % len(muc),
         "khong giai thich, khong danh gia:", "",
         TACH % "<so>", "<ban viet lai>", "", "----", ""]
    for i, m in enumerate(muc, 1):
        p += [TACH % i, m["prompt"].strip(), ""]
    return "\n".join(p)


def tach_lo(tra, so_muc):
    """Boc tra loi lo -> {so: van ban}. Tra None neu thieu/lech so muc (bo ca lo)."""
    phan = TACH_RE.split(tra)
    if len(phan) < 3:
        return None
    ra = {}
    for i in range(1, len(phan) - 1, 2):
        khoa = phan[i].strip()
        noi = phan[i + 1].strip()
        noi = re.sub(r"^-{3,}\s*", "", noi).strip()
        if noi:
            ra[khoa] = noi
    if len(ra) != so_muc or any(str(i) not in ra for i in range(1, so_muc + 1)):
        return None
    return ra


# --------------------------------------------------------------------- chay

def da_xong(duong_dan):
    xong = set()
    if os.path.exists(duong_dan):
        with io.open(duong_dan, encoding="utf-8") as f:
            for ln in f:
                ln = ln.strip()
                if not ln:
                    continue
                try:
                    d = json.loads(ln)
                except ValueError:
                    continue                    # dong hong do bi kill giua chung
                if d.get("reply") and not d.get("error"):
                    xong.add(d["id"])
    return xong


def doc_vao(duong_dan):
    muc = []
    with io.open(duong_dan, encoding="utf-8-sig") as f:
        for i, ln in enumerate(f, 1):
            ln = ln.strip()
            if not ln:
                continue
            try:
                d = json.loads(ln)
            except ValueError as e:
                raise RuntimeError("dong %d khong phai JSON hop le: %s" % (i, e))
            if not d.get("id") or not d.get("prompt"):
                raise RuntimeError("dong %d thieu 'id' hoac 'prompt'" % i)
            muc.append(d)
    return muc


_ghi_khoa = threading.Lock()
_nghi = threading.Event()          # dat = dang bi rate limit, ca nhom cung nghi


def _ghi(f, ban_ghi):
    with _ghi_khoa:
        f.write(json.dumps(ban_ghi, ensure_ascii=False) + "\n")
        f.flush()
        os.fsync(f.fileno())


def chay(vao, ra, port=9222, lo=1, workers=1, brief_path=None, timeout=300.0,
         on_dinh=4, cho=3.0, lui=300.0, tam_thoi=True, kieu="viet-lai"):
    if kieu in ("cham", "hoi") and lo > 1:
        raise RuntimeError("kieu %s chi chay 1 muc/luot (--lo 1)" % kieu)
    kiem = {"cham": kiem_phieu, "hoi": kiem_hoi}.get(kieu)
    muc = [m for m in doc_vao(vao) if m["id"] not in da_xong(ra)]
    if not muc:
        sys.stderr.write("khong con muc nao chua xong.\n")
        return 0
    brief = io.open(brief_path, encoding="utf-8").read() if brief_path else ""

    khoi = [muc[i:i + lo] for i in range(0, len(muc), lo)] if lo > 1 else [[m] for m in muc]
    hang = list(enumerate(khoi))
    hang_khoa = threading.Lock()
    f = io.open(ra, "a", encoding="utf-8")
    tabs = tim_tab(port)
    while len(tabs) < workers:
        tabs.append(mo_tab_moi(port))
    loi = [0]

    def worker(wid, target):
        ws = cdp_connect(target)
        kiem_dang_nhap(ws)
        n_da = 0
        while True:
            if _nghi.is_set():
                _nghi.wait()
            with hang_khoa:
                if not hang:
                    break
                idx, nhom = hang.pop(0)
            manh = (n_da > 0 and n_da % 25 == 0)
            n_da += 1
            prompt = gom_lo(nhom, brief) if len(nhom) > 1 else (
                (brief + "\n\n" + nhom[0]["prompt"]) if brief else nhom[0]["prompt"])
            for lan in range(3):
                try:
                    tra, meta = hoi(ws, prompt, timeout, on_dinh, tam_thoi, manh)
                    break
                except RateLimited as e:
                    sys.stderr.write("[w%d] GIOI HAN (lan %d/3): %s\n  nghi %.0fs...\n"
                                     % (wid, lan + 1, e, lui))
                    _nghi.set()
                    time.sleep(lui)
                    _nghi.clear()
                    if lan == 2:
                        loi[0] += 1
                        for m in nhom:
                            _ghi(f, {"id": m["id"], "reply": "", "error": "rate limit 3 lan"})
                        with hang_khoa:
                            hang.clear()          # dung ca luot, khong sinh 200 ban ghi rac
                        return
                except RuntimeError as e:
                    loi[0] += 1
                    for m in nhom:
                        _ghi(f, {"id": m["id"], "reply": "", "error": str(e)[:300]})
                    sys.stderr.write("[w%d] LOI: %s\n" % (wid, e))
                    tra = None
                    break
            if tra is None:
                continue

            if len(nhom) > 1:
                phan = tach_lo(tra, len(nhom))
                if phan is None:
                    # bo CA LO, ha ve 1 doan/luot cho rieng lo nay. Khong khop lech thu tu.
                    sys.stderr.write("[w%d] lo %d boc khong du %d muc -> ha ve 1 doan/luot\n"
                                     % (wid, idx, len(nhom)))
                    with hang_khoa:
                        for m in nhom:
                            hang.insert(0, (idx, [m]))
                    continue
                for i, m in enumerate(nhom, 1):
                    bg = {"id": m["id"], "reply": phan[str(i)],
                          "elapsed": round(meta["elapsed"] / len(nhom), 1)}
                    xau = kiem_ban_ghi(m["prompt"], phan[str(i)])
                    if xau:
                        bg["error"] = xau
                        loi[0] += 1
                        sys.stderr.write("[w%d] %-10s BAN GHI XAU: %s\n" % (wid, m["id"], xau))
                    _ghi(f, bg)
                sys.stderr.write("[w%d] lo %d: %d muc, %.1fs\n" % (wid, idx, len(nhom), meta["elapsed"]))
            else:
                bg = {"id": nhom[0]["id"], "reply": tra, "elapsed": round(meta["elapsed"], 1)}
                if meta.get("error"):
                    bg["error"] = meta["error"]
                    loi[0] += 1
                else:
                    xau = kiem(tra) if kiem else kiem_ban_ghi(nhom[0]["prompt"], tra)
                    if xau:
                        bg["error"] = xau
                        loi[0] += 1
                        sys.stderr.write("[w%d] %-10s BAN GHI XAU: %s\n"
                                         % (wid, nhom[0]["id"], xau))
                _ghi(f, bg)
                sys.stderr.write("[w%d] %-10s %s %.1fs %d ky tu\n"
                                 % (wid, nhom[0]["id"], meta.get("gui", "-"), meta["elapsed"], len(tra)))
            time.sleep(cho + random.uniform(0, 1.5))

    luong = [threading.Thread(target=worker, args=(i, tabs[i]), daemon=True) for i in range(workers)]
    for t in luong:
        t.start()
    for t in luong:
        t.join()
    f.close()
    sys.stderr.write("xong. %d loi/ban ghi xau.\n" % loi[0])
    if loi[0]:
        sys.stderr.write("CHUA DUOC DUNG FILE NAY cho cong A/B: mo cac ban ghi co truong "
                         "`error` ra doc truoc.\n")
    else:
        sys.stderr.write("Nhac: van phai mo 3 ban ghi ra doc bang mat truoc khi dung "
                         "(loop-review.md muc 8.4).\n")
    return 1 if loi[0] else 0


# ------------------------------------------------------------------ selftest

def selftest(port, live=False, tam_thoi=True):
    hong = 0

    def dong(ten, tt, ghi=""):
        print("[selftest] %-14s %-8s %s" % (ten, tt, ghi))

    tabs = tim_tab(port)
    dong("CDP %d" % port, "OK", "%d tab chatgpt.com" % len(tabs))
    ws = cdp_connect(tabs[0])
    dong("websocket", "OK", "suppress_origin=True")
    try:
        kiem_dang_nhap(ws)
        dong("dang nhap", "OK")
    except RuntimeError as e:
        dong("dang nhap", "HONG", str(e))
        return 1

    # Nut send CHI ton tai khi composer co chu (da kiem chung 2026-09-04) -> phai go thu vao
    # truoc khi kiem, khong thi selftest bao HONG oan.
    # Dung CHINH `go_prompt` chu khong goi Input.insertText tho: go_prompt co buoc selectAll +
    # Backspace de reset ProseMirror. Thieu buoc do, composer hien chu nhung React van coi la
    # rong va giu nut Voice o cho nut Send -- da do that 2026-09-04.
    try:
        go_prompt(ws, "x")
        time.sleep(0.6)
    except RuntimeError as e:
        dong("go thu", "HONG", str(e))
        hong += 1

    for ten in ("composer", "send", "assistant", "newchat", "stop", "markdown"):
        thay = None
        for i, s in enumerate(SEL[ten], 1):
            try:
                if cdp_eval(ws, "!!document.querySelector(%s)" % json.dumps(s)):
                    thay = (i, s)
                    break
            except RuntimeError:
                pass
        if thay and ten == "stop":
            # Nhom `stop` khop khi KHONG dang stream la LOI, khong phai "can cap
            # nhat". Dieu kien ket thuc cua vong doi la `not p["stop"]`, nen mot
            # phan tu THUONG TRU khop nhom nay lam cau noi khong bao gio thay luot
            # tra loi ket thuc. Da xay ra that: ung vien `button[aria-label*="Dung"]`
            # khop hai muc lich su chat o thanh ben ("Bien tap noi dung"), va mot
            # cau tra loi mot chu mat 121 giay.
            hong += 1
            dong(ten, "HONG",
                 "%s khop trong khi KHONG stream -> selector quet ra ngoai khung"
                 " soan. Rang buoc vao `form` roi chay lai." % thay[1])
        elif thay:
            ghi = "<- %s (ung vien %d/%d)%s" % (thay[1], thay[0], len(SEL[ten]),
                                                "  ** CAP NHAT SEL **" if thay[0] > 1 else "")
            dong(ten, "OK", ghi)
        elif ten in ("stop", "markdown", "assistant"):
            dong(ten, "BO QUA", "chi ton tai khi dang stream / da co luot tra loi")
        else:
            hong += 1
            dong(ten, "HONG", "khong ung vien nao khop -> Inspect roi them vao SEL[%r]" % ten)

    # don sach composer sau khi go thu
    try:
        cdp_eval(ws, JS_FOCUS % Q("composer"))
        cdp_eval(ws, "document.execCommand('selectAll'), 1")
        cdp_key(ws, "Backspace", 8, "Backspace")
    except RuntimeError:
        pass

    if live:
        print("[selftest] --live: chay mot vong that...")
        try:
            tra, meta = hoi(ws, "Tra loi dung mot tu: OK", timeout=120, tam_thoi=tam_thoi)
            ok = "OK" in (tra or "").upper()
            dong("vong that", "OK" if ok else "HONG",
                 "%.1fs, gui bang %s, nhan: %r" % (meta["elapsed"], meta.get("gui"), (tra or "")[:60]))
            hong += 0 if ok else 1
            # Chieu nguoc cua loi selector `stop`: phai THAY no it nhat mot lan
            # trong luc stream. Khong thay nghia la dieu kien ket thuc dung ngay
            # tu vong dau, va ban tra ve co the bi CUT ma khong bao gi (E35).
            if DA_THAY_STOP:
                dong("stop luc stream", "OK", "co thay nut dung -> vong doi that")
            else:
                hong += 1
                dong("stop luc stream", "HONG",
                     "KHONG thay nut dung lan nao trong luc stream -> dieu kien"
                     " ket thuc dung ngay tu dau, ban tra ve co the bi cut")
        except RuntimeError as e:
            hong += 1
            dong("vong that", "HONG", str(e))

    print("KET QUA: %s" % ("tat ca dat." if not hong else "%d HONG -> exit 1" % hong))
    return hong


def _utf8_console():
    """Console Windows mac dinh cp1252 -> in tieng Viet la UnicodeEncodeError.
    Da lam hong `--tu-kiem` mot lan (2026-09-04)."""
    for luong in (sys.stdout, sys.stderr):
        try:
            luong.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


if __name__ == "__main__":
    _utf8_console()
    ap = argparse.ArgumentParser(description="Cau noi ChatGPT web qua Chrome DevTools Protocol")
    ap.add_argument("vao", nargs="?")
    ap.add_argument("ra", nargs="?")
    ap.add_argument("--port", type=int, default=9222)
    ap.add_argument("--lo", type=int, default=1, help="so doan gop vao mot luot chat")
    ap.add_argument("--workers", type=int, default=1, help="so tab chay song song")
    ap.add_argument("--brief", help="file style brief gui kem moi request (KHONG mac dinh)")
    ap.add_argument("--tran", action="store_true",
                    help="gui dung doan van, bo qua ca --brief neu co (mac dinh la tran)")
    ap.add_argument("--timeout", type=float, default=300.0)
    ap.add_argument("--on-dinh", type=int, default=4, help="so lan poll noi dung phai dung yen")
    ap.add_argument("--cho", type=float, default=3.0, help="giay nghi giua hai luot")
    ap.add_argument("--lui", type=float, default=300.0, help="giay nghi khi bi rate limit")
    ap.add_argument("--khong-tam-thoi", action="store_true", help="khong dung Temporary Chat")
    ap.add_argument("--kieu", choices=("viet-lai", "cham", "hoi"), default="viet-lai",
                    help="cham: reply la phieu cham, bat buoc khoi ###PHIEU###, bo kiem do dai; "
                         "hoi: hoi dap tu do, chi kiem reply khong rong")
    ap.add_argument("--selftest", action="store_true")
    ap.add_argument("--live", action="store_true", help="selftest chay them mot vong that")
    a = ap.parse_args()
    try:
        if a.selftest:
            sys.exit(1 if selftest(a.port, a.live, not a.khong_tam_thoi) else 0)
        if not a.vao or not a.ra:
            ap.error("thieu <vao.jsonl> <ra.jsonl> (hoac dung --selftest)")
        sys.exit(chay(a.vao, a.ra, a.port, a.lo, a.workers,
                      None if a.tran else a.brief, a.timeout,
                      a.on_dinh, a.cho, a.lui, not a.khong_tam_thoi, a.kieu))
    except RuntimeError as e:
        sys.stderr.write("LOI: %s\n" % e)
        sys.exit(1)
