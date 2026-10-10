# Tu phat hien loi cua Power BI Desktop va lay NGUYEN VAN thong bao loi.
#
# Van de: khi mo mot .pbip co TMDL/PBIR sai, Desktop bat mot dialog modal ("Issues were found",
# "Something went wrong"...) va VO HIEU HOA cua so chinh. Script/loop khong biet gi ca - moi thao
# tac tiep theo deu chet im lang. Script nay:
#   1. tim cua so PBIDesktop va moi dialog modal dang chan,
#   2. chup dialog ra PNG,
#   3. bam "Copy details to clipboard" (neu co) va IN nguyen van chi tiet loi,
#   4. voi -Close thi dong dialog de tra lai cua so chinh.
#
# Ma thoat: 0 = sach, 2 = co dialog loi (chi tiet da in ra stdout), 3 = khong thay PBIDesktop.
#
# Dung:
#   powershell -File pbi_health.ps1                     # chi bao cao
#   powershell -File pbi_health.ps1 -Close              # bao cao roi dong dialog
param(
    [string]$OutDir = (Join-Path (Get-Location) "review"),
    [switch]$Close
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class H {
  public delegate bool EnumProc(IntPtr h, IntPtr p);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint f);
  public struct RECT { public int L, T, R, B; }
  public static List<IntPtr> TopLevel(uint pid) {
    var list = new List<IntPtr>();
    EnumWindows((h, p) => { uint w; GetWindowThreadProcessId(h, out w); if (w == pid) list.Add(h); return true; }, IntPtr.Zero);
    return list;
  }
  public static string Title(IntPtr h) { var sb = new StringBuilder(512); GetWindowTextW(h, sb, 512); return sb.ToString(); }
  public static string Cls(IntPtr h) { var sb = new StringBuilder(512); GetClassNameW(h, sb, 512); return sb.ToString(); }
}
"@

function Snap([IntPtr]$h, [string]$path) {
    $r = New-Object H+RECT
    [H]::GetWindowRect($h, [ref]$r) | Out-Null
    $w = $r.R - $r.L; $ht = $r.B - $r.T
    if ($w -le 0 -or $ht -le 0) { return $false }
    $bmp = New-Object System.Drawing.Bitmap($w, $ht)
    $g = [System.Drawing.Graphics]::FromImage($bmp); $dc = $g.GetHdc()
    [H]::PrintWindow($h, $dc, 2) | Out-Null; $g.ReleaseHdc($dc); $g.Dispose()
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
    return $true
}

function InvokeByName([IntPtr]$hwnd, [string[]]$names) {
    $root = [System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
    if ($null -eq $root) { return $false }
    $all = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                         [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($e in $all) {
        $n = $e.Current.Name
        foreach ($want in $names) {
            if ($n -and $n.ToLower().Contains($want.ToLower())) {
                try {
                    $p = $e.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                    $p.Invoke(); Start-Sleep -Milliseconds 400; return $true
                } catch { }
            }
        }
    }
    return $false
}

function DialogText([IntPtr]$hwnd) {
    $root = [System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
    if ($null -eq $root) { return "" }
    $all = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                         [System.Windows.Automation.Condition]::TrueCondition)
    $parts = @()
    foreach ($e in $all) { if ($e.Current.Name) { $parts += $e.Current.Name } }
    return ($parts | Select-Object -Unique) -join " | "
}

$procs = @(Get-Process PBIDesktop -ErrorAction SilentlyContinue)
if ($procs.Count -eq 0) { Write-Host "PBI_HEALTH: khong thay PBIDesktop dang chay."; exit 3 }

New-Item -ItemType Directory -Force $OutDir | Out-Null
$found = 0

foreach ($p in $procs) {
    $wins = [H]::TopLevel([uint32]$p.Id)
    $main = $null; $dialogs = @()
    # Cua so chinh nhan dien bang CLASS, khong bang tieu de: khi da mo mot project, tieu de la
    # ten project ("LastMile") chu KHONG con hau to " - Power BI Desktop" - da bao nham mot lan.
    # Gom MOI cua so con dang hien, KE CA cua so KHONG CO TIEU DE. Ban truoc chi gom cua so
    # co tieu de, nen hop thoai dang nhap PostgreSQL (web view, tieu de rong) lot luoi: script
    # bao "sach" trong khi main_enabled=False va moi lenh phia sau treo im - da mat kha nhieu
    # thoi gian vi cho nay.
    foreach ($h in $wins) {
        if (-not [H]::IsWindowVisible($h)) { continue }
        if ([H]::Cls($h) -like "WindowsForms10.Window.8.app.*") { $main = $h; continue }
        # Bo cua so kich thuoc 0 - la cua so ky thuat an, khong phai hop thoai.
        $rr = New-Object H+RECT; [H]::GetWindowRect($h, [ref]$rr) | Out-Null
        if (($rr.R - $rr.L) -le 0 -or ($rr.B - $rr.T) -le 0) { continue }
        $dialogs += $h
    }
    $enabled = $true
    if ($main) {
        $enabled = [H]::IsWindowEnabled($main)
        Write-Host ("PBI_HEALTH: pid={0} title='{1}' main_enabled={2}" -f $p.Id, [H]::Title($main), $enabled)
    }

    # Cua so chinh CON dung duoc thi chi bao dong khi tieu de nam trong danh sach loi da biet -
    # neu khong, cac cua so web view cua chinh report se bi bao nham. Cua so chinh BI VO HIEU
    # HOA thi moi cua so con deu la thu dang chan, ke ca cai khong ten.
    $KNOWN = @("issues were found", "something went wrong", "couldn't load", "unable to",
               "error", "loi", "cannot")
    if ($enabled) {
        $dialogs = @($dialogs | Where-Object {
            $t = [H]::Title($_).ToLower()
            $t -and ($KNOWN | Where-Object { $t.Contains($_) }).Count -gt 0
        })
    }

    foreach ($d in $dialogs) {
        $found++
        $title = [H]::Title($d)
        $r0 = New-Object H+RECT; [H]::GetWindowRect($d, [ref]$r0) | Out-Null
        $kich = "{0}x{1}" -f ($r0.R - $r0.L), ($r0.B - $r0.T)
        if (-not $title) {
            # Hop thoai dang nhap nguon du lieu khong co tieu de; nhan dang bang kich thuoc de
            # bao dung benh thay vi de nguoi doc doan.
            $title = if ($kich -eq "702x311") { "(dang nhap nguon du lieu)" } else { "(khong ten)" }
        }
        Write-Host ("PBI_HEALTH: DIALOG '{0}' {1} enabled={2}" -f $title, $kich, [H]::IsWindowEnabled($d))
        $png = Join-Path $OutDir ("dialog_" + ($title -replace '[^A-Za-z0-9]', '_') + ".png")
        if (Snap $d $png) { Write-Host ("PBI_HEALTH: anh -> {0}" -f $png) }

        Write-Host ("PBI_HEALTH: text  | " + (DialogText $d))

        # nut "Copy details to clipboard" cho nguyen van loi - thu tin nhat
        [System.Windows.Forms.Clipboard]::Clear()
        if (InvokeByName $d @("Copy details", "Sao chep chi tiet")) {
            Start-Sleep -Milliseconds 500
            $txt = ""
            try { $txt = [System.Windows.Forms.Clipboard]::GetText() } catch { }
            if ($txt) {
                Write-Host "PBI_HEALTH: ---- CHI TIET ----"
                Write-Host $txt
                Write-Host "PBI_HEALTH: ---- HET ----"
                $txtFile = Join-Path $OutDir "last_error.txt"
                [IO.File]::WriteAllText($txtFile, $txt, [Text.Encoding]::UTF8)
                Write-Host ("PBI_HEALTH: da luu -> {0}" -f $txtFile)
            }
        }

        if ($Close) {
            if (InvokeByName $d @("Close", "OK", "Dong")) { Write-Host "PBI_HEALTH: da dong dialog." }
            else { Write-Host "PBI_HEALTH: KHONG dong duoc dialog - can dong tay." }
        }
    }
}

if ($found -gt 0) { exit 2 }
Write-Host "PBI_HEALTH: sach - khong co dialog nao dang chan."
exit 0
