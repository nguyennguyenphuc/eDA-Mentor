# Xuat report dang mo trong Power BI Desktop ra PDF (File > Export > Export to PDF).
#
# DUNG XUAT NGAY SAU KHI MO TEP. Do that 01/09/2026: xuat lien sau khi mo
# mot .pbix thi trang 3 ra 5.08 / 2.11 / (Blank) / 28.8%, trong khi hoi
# thang model duoc 5.13 / 12.51 / 25.8% / 25.8%. Ban xuat ngay 31/08 (mo lau roi
# moi xuat) khop model tuyet doi. Nghia la Desktop tra ve PDF truoc khi cac
# visual tinh xong, va KHONG bao gi ca -- PDF van du 4 trang, van dep.
# Cho toi khi bam thu mot slicer thay so doi ngay, roi hay xuat.
# Phuong phap: click chuot that theo toa do cua so (keyboard injection khong tac dung voi PBI),
# xac nhan trang thai bang mau pixel tu PrintWindow, Save As dialog xu ly qua UIA.
# KHONG BAO GIO gui Ctrl+S. Chay UNSANDBOXED (can quyen input injection).
param(
    [string]$OutPdf = (Join-Path (Get-Location) "review\render.pdf"),
    [int]$TimeoutSec = 360,
    [string]$DebugPng = (Join-Path (Get-Location) "review\shots\pbi_debug.png"),
    # Chon dung cua so khi may dang mo nhieu tep, vd -TitleLike '*TenDuAn*'.
    # Bo trong thi giu hanh vi cu, nhung se BAO LOI neu co tu 2 cua so tro len.
    [string]$TitleLike = ""
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class W {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint x, uint y, uint d, UIntPtr e);
    [DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, uint f, UIntPtr e);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint f);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT p);
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint flags);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, IntPtr pid);
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool f);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    public struct RECT { public int L, T, R, B; }
    public struct POINT { public int X, Y; }
    public static RECT GetRect(IntPtr h) { RECT r; GetWindowRect(h, out r); return r; }
    public static IntPtr RootWindowAt(int x, int y) {
        POINT p; p.X = x; p.Y = y;
        return GetAncestor(WindowFromPoint(p), 2); // GA_ROOT
    }
}
"@

function Log($m) { Write-Host ("[{0}] {1}" -f (Get-Date -Format HH:mm:ss), $m) }

function SnapBmp($hwnd) {
    $r = [W]::GetRect($hwnd)
    $bmp = New-Object System.Drawing.Bitmap([int]($r.R - $r.L), [int]($r.B - $r.T))
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $dc = $g.GetHdc(); [W]::PrintWindow($hwnd, $dc, 2) | Out-Null; $g.ReleaseHdc($dc); $g.Dispose()
    return $bmp
}
function SaveSnap($hwnd, $path) {
    $bmp = SnapBmp $hwnd
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
    Log "Da chup: $path"
}
function Click($hwnd, $wx, $wy) {
    # click theo toa do TRONG cua so
    $r = [W]::GetRect($hwnd)
    $sx = [int]($r.L + $wx); $sy = [int]($r.T + $wy)
    [W]::SetCursorPos($sx, $sy) | Out-Null
    Start-Sleep -Milliseconds 200
    $under = [W]::RootWindowAt($sx, $sy)
    if ($under -ne $hwnd) {
        Log "CANH BAO: duoi con tro ($sx,$sy) la hwnd=$under, khong phai PBI ($hwnd) -> SetForeground lai"
        [W]::SetForegroundWindow($hwnd) | Out-Null
        Start-Sleep -Milliseconds 500
    }
    [W]::SetCursorPos($sx, $sy) | Out-Null
    [W]::mouse_event(2, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 80
    [W]::SetCursorPos($sx, $sy) | Out-Null
    [W]::mouse_event(4, 0, 0, 0, [UIntPtr]::Zero)
}
function IsBackstage($hwnd) {
    # (300,90): view thuong = ribbon toi (~65), backstage = nen sang (~250). Da kiem chung bang anh chup.
    $bmp = SnapBmp $hwnd
    $px = $bmp.GetPixel(300, 90); $bmp.Dispose()
    return ($px.R -gt 150 -and $px.G -gt 150 -and $px.B -gt 150)
}
function IsExportPage($hwnd) {
    # trang Export: icon PDF vien do trong vung x 245-285, y 228-250 -> dem pixel bao hoa
    $bmp = SnapBmp $hwnd
    $count = 0
    for ($x = 245; $x -le 285; $x++) { for ($y = 228; $y -le 250; $y++) {
        $px = $bmp.GetPixel($x, $y)
        $sat = ([Math]::Max([Math]::Max($px.R, $px.G), $px.B)) - ([Math]::Min([Math]::Min($px.R, $px.G), $px.B))
        if ($sat -gt 40) { $count++ }
    } }
    $bmp.Dispose()
    return ($count -ge 10)
}

# --- 0. Man hinh co dang khoa khong ---
# Export PDF phai click chuot that nen KHONG chay duoc khi phien Windows bi khoa:
# GetForegroundWindow() tra ve 0 va moi SetForegroundWindow deu that bai. Bao dung ly do,
# dung de nguoi doc tuong la loi Power BI.
#
# PHAI loc theo SESSION. Ban dau cho nay chi hoi `Get-Process LogonUI`, tuc hoi
# CA MAY. Khi nguoi dung noi vao bang RDP thi Windows khoa man hinh console vat
# ly, LogonUI chay o session do, va phep kiem bao "MAN HINH DANG KHOA" trong khi
# phien RDP van hoat dong binh thuong -- export khong bao gio chay duoc qua RDP.
# Da gap that ngay 31/08/2026: session 1 (rdp-tcp#0) Active, LogonUI o session 5
# (console) Conn.
#
# Kem theo phep do TRUC TIEP dieu kien that ma chu thich tren mo ta, thay vi chi
# suy tu mot dau hieu gian tiep.
$mySession = (Get-Process -Id $PID).SessionId
$lockUI = Get-Process LogonUI -ErrorAction SilentlyContinue |
          Where-Object { $_.SessionId -eq $mySession }
if ($lockUI -or [W]::GetForegroundWindow() -eq [IntPtr]::Zero) {
    throw "MAN HINH DANG KHOA - mo khoa Windows roi chay lai (export PDF can dieu khien chuot that)."
}

# --- 1. Focus PBI ---
# `Select-Object -First 1` la du khi chi co MOT cua so PBIDesktop. Mo hai tep cung
# luc thi thu tu process khong xac dinh, va script se xuat NHAM bao cao ma khong
# bao gi. Da gap that 01/09/2026: mot cua so ten "MoHinh1" mo san.
# `-TitleLike` la tuy chon, khong truyen thi hanh vi giu nguyen nhu cu.
$dsTatCa = Get-Process PBIDesktop -ErrorAction Stop | Where-Object { $_.MainWindowHandle -ne 0 }
if ($TitleLike) {
    $proc = $dsTatCa | Where-Object { $_.MainWindowTitle -like $TitleLike } | Select-Object -First 1
    if (-not $proc) {
        $co = ($dsTatCa | ForEach-Object { $_.MainWindowTitle }) -join " | "
        throw "Khong co cua so PBIDesktop nao khop '$TitleLike'. Dang mo: $co"
    }
} else {
    if ($dsTatCa.Count -gt 1) {
        $co = ($dsTatCa | ForEach-Object { $_.MainWindowTitle }) -join " | "
        throw "Co $($dsTatCa.Count) cua so PBIDesktop, phai chi ro -TitleLike. Dang mo: $co"
    }
    $proc = $dsTatCa | Select-Object -First 1
}
if (-not $proc) { throw "Khong tim thay cua so PBIDesktop" }
Log ("Cua so dich: " + $proc.MainWindowTitle)
$hwnd = $proc.MainWindowHandle
if ([W]::IsIconic($hwnd)) { [W]::ShowWindow($hwnd, 9) | Out-Null }
for ($i = 0; $i -lt 5; $i++) {
    [W]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero)
    [W]::SetForegroundWindow($hwnd) | Out-Null
    [W]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 600
    if ([W]::GetForegroundWindow() -eq $hwnd) { break }
}
if ([W]::GetForegroundWindow() -ne $hwnd) {
    # Windows chan SetForegroundWindow khi tien trinh goi khong giu foreground. Cach vuot:
    # gan input queue cua minh vao thread dang giu foreground roi moi doi. Da can den that
    # sau khi loop dong/mo Desktop lien tuc - khong co doan nay thi export chet han.
    Log "SetForegroundWindow bi tu choi, thu AttachThreadInput..."
    $fg = [W]::GetForegroundWindow()
    $tFg = [W]::GetWindowThreadProcessId($fg, [IntPtr]::Zero)
    $tMe = [W]::GetCurrentThreadId()
    [W]::AttachThreadInput($tMe, $tFg, $true) | Out-Null
    [W]::ShowWindow($hwnd, 9) | Out-Null
    [W]::BringWindowToTop($hwnd) | Out-Null
    [W]::SetForegroundWindow($hwnd) | Out-Null
    [W]::AttachThreadInput($tMe, $tFg, $false) | Out-Null
    Start-Sleep -Milliseconds 800
}
if ([W]::GetForegroundWindow() -ne $hwnd) {
    # Doan cuoi: thu nho roi phong to lai buoc Windows trao foreground.
    [W]::ShowWindow($hwnd, 6) | Out-Null; Start-Sleep -Milliseconds 400
    [W]::ShowWindow($hwnd, 3) | Out-Null; Start-Sleep -Milliseconds 800
    [W]::SetForegroundWindow($hwnd) | Out-Null; Start-Sleep -Milliseconds 400
}
if ([W]::GetForegroundWindow() -ne $hwnd) { throw "Khong dua duoc PBI len foreground" }
Log "PBI foreground OK (PID $($proc.Id))"

# --- 1b. PHONG TO cua so ---
# Moi toa do click ben duoi deu do tu cua so PHONG TO. Neu cua so dang o kich
# thuoc khac, nut File khong nam o (48,66) va ca trinh tu truot ngay tu buoc
# dau. Da dinh dung loi nay: vong truoc mo Desktop khong phong to -> "Khong mo
# duoc File backstage".
[W]::ShowWindow($hwnd, 3) | Out-Null      # SW_MAXIMIZE
Start-Sleep -Milliseconds 1200
$r = [W]::GetRect($hwnd)
Log ("Cua so: {0}x{1}" -f ($r.R - $r.L), ($r.B - $r.T))

if (Test-Path $OutPdf) { Remove-Item $OutPdf -Force }

# --- 2. Dam bao trang thai binh thuong (dong backstage neu con mo) ---
$tries = 0
while ((IsBackstage $hwnd) -and $tries -lt 3) {
    Log "Backstage dang mo -> bam nut back"
    Click $hwnd 46 74
    Start-Sleep -Milliseconds 1500
    $tries++
}
if (IsBackstage $hwnd) { SaveSnap $hwnd $DebugPng; throw "Khong dong duoc backstage cu" }

# --- 3. Mo File backstage (xac minh bang pixel) ---
$ok = $false
for ($i = 0; $i -lt 3 -and -not $ok; $i++) {
    Click $hwnd 48 66
    Start-Sleep -Milliseconds 1800
    $ok = IsBackstage $hwnd
    if (-not $ok) { Log "Backstage chua mo (lan $($i+1)), thu lai..." }
}
if (-not $ok) { SaveSnap $hwnd $DebugPng; throw "Khong mo duoc File backstage" }
Log "Backstage da mo"

# --- 4. Vao trang Export (xac minh icon PDF hien ra truoc khi click) ---
$ok = $false
for ($i = 0; $i -lt 3 -and -not $ok; $i++) {
    Click $hwnd 88 418      # muc 'Export' o menu trai
    Start-Sleep -Milliseconds 1500
    $ok = IsExportPage $hwnd
    if (-not $ok) { Log "Trang Export chua hien (lan $($i+1)), thu lai..." }
}
if (-not $ok) { SaveSnap $hwnd $DebugPng; throw "Khong vao duoc trang Export" }
Log "Trang Export da hien (thay icon PDF)"

$startTime = Get-Date
Click $hwnd 341 234     # 'Export to PDF'
Log "Da bam Export to PDF. PBI se render va tu luu vao %TEMP%\Power BI Desktop\print-job-*\ ..."

# --- 5. Doi PDF moi xuat hien trong print-job folder (PBI khong hoi noi luu) ---
$printDir = Join-Path $env:LOCALAPPDATA "Temp\Power BI Desktop"
$deadline = (Get-Date).AddSeconds($TimeoutSec)
$pdf = $null
while ((Get-Date) -lt $deadline -and -not $pdf) {
    $cand = Get-ChildItem $printDir -Filter "*.pdf" -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -gt $startTime -and $_.Name -notlike 'section-*' } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($cand) { $pdf = $cand } else { Start-Sleep -Milliseconds 1500 }
}
if (-not $pdf) { SaveSnap $hwnd $DebugPng; throw "Khong thay PDF moi trong $printDir sau $TimeoutSec s" }
Log "Thay PDF: $($pdf.FullName)"

# --- 6. Doi file on dinh roi copy ve dich ---
$lastSize = -1; $stable = 0
$deadline = (Get-Date).AddSeconds(120)
while ((Get-Date) -lt $deadline -and $stable -lt 3) {
    $sz = (Get-Item $pdf.FullName).Length
    if ($sz -gt 0 -and $sz -eq $lastSize) { $stable++ } else { $stable = 0 }
    $lastSize = $sz
    Start-Sleep -Milliseconds 1000
}
Copy-Item $pdf.FullName $OutPdf -Force
Log "XONG: $OutPdf ($([math]::Round((Get-Item $OutPdf).Length/1KB)) KB)"
exit 0
