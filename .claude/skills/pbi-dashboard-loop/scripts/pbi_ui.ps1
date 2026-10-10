# Điều khiển Power BI Desktop qua UIA: liệt kê element, bấm nút theo TÊN, chụp cửa sổ.
#
# Vì sao có file này: cửa sổ report của PBI là web view, nhưng UIA VẪN nhìn thấy toàn bộ
# (~1.500 element) — kể cả "Refresh now", "See details", "Fix this". Nhờ vậy loop tự phát hiện
# và tự xử lý được thay vì click mù theo toạ độ như export_pdf.ps1 phải làm.
#
# Dùng:
#   pbi_ui.ps1 -List                       # liệt kê tên element (lọc bằng -Match)
#   pbi_ui.ps1 -List -Match "Refresh"
#   pbi_ui.ps1 -Click "Refresh now"        # bấm nút/hyperlink đầu tiên khớp tên
#   pbi_ui.ps1 -Click "See details" -Index 0
#   pbi_ui.ps1 -Shot review\state.png      # chụp cửa sổ chính
#
# Mã thoát: 0 OK · 1 không tìm thấy element · 3 không thấy PBIDesktop.
param(
    [switch]$List,
    [string]$Match = "",
    [string]$Click = "",
    [int]$Index = 0,
    [string]$Shot = "",
    [switch]$Exact
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class U {
  public delegate bool EnumProc(IntPtr h, IntPtr p);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint f);
  public struct RECT { public int L, T, R, B; }
  public static string Cls(IntPtr h) { var sb = new StringBuilder(512); GetClassNameW(h, sb, 512); return sb.ToString(); }
  public static IntPtr MainOf(uint pid) {
    IntPtr found = IntPtr.Zero;
    EnumWindows((h, p) => {
      uint w; GetWindowThreadProcessId(h, out w);
      if (w == pid && IsWindowVisible(h) && Cls(h).StartsWith("WindowsForms10.Window.8.app.")) { found = h; return false; }
      return true;
    }, IntPtr.Zero);
    return found;
  }
}
"@

$proc = Get-Process PBIDesktop -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $proc) { Write-Host "PBI_UI: khong thay PBIDesktop."; exit 3 }
$hwnd = [U]::MainOf([uint32]$proc.Id)
if ($hwnd -eq [IntPtr]::Zero) { Write-Host "PBI_UI: khong thay cua so chinh."; exit 3 }

if ($Shot) {
    [U]::SetForegroundWindow($hwnd) | Out-Null
    Start-Sleep -Milliseconds 1200
    $r = New-Object U+RECT; [U]::GetWindowRect($hwnd, [ref]$r) | Out-Null
    $bmp = New-Object System.Drawing.Bitmap(($r.R - $r.L), ($r.B - $r.T))
    $g = [System.Drawing.Graphics]::FromImage($bmp); $dc = $g.GetHdc()
    [U]::PrintWindow($hwnd, $dc, 2) | Out-Null; $g.ReleaseHdc($dc); $g.Dispose()
    New-Item -ItemType Directory -Force (Split-Path $Shot) | Out-Null
    $bmp.Save($Shot, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
    Write-Host "PBI_UI: anh -> $Shot"
}

if ($List -or $Click) {
    $root = [System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
    $all = $root.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                         [System.Windows.Automation.Condition]::TrueCondition)
    if ($List) {
        $seen = @{}
        foreach ($e in $all) {
            $n = $e.Current.Name
            if (-not $n) { continue }
            if ($Match -and $n -notmatch $Match) { continue }
            $key = "$n|$($e.Current.ControlType.ProgrammaticName)"
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = 1
            Write-Host ("PBI_UI: {0}  [{1}]" -f $n, $e.Current.ControlType.ProgrammaticName)
        }
    }
    if ($Click) {
        $hits = @()
        foreach ($e in $all) {
            $n = $e.Current.Name
            if (-not $n) { continue }
            $ok = if ($Exact) { $n -eq $Click } else { $n.ToLower().Contains($Click.ToLower()) }
            if ($ok) { $hits += $e }
        }
        if ($hits.Count -le $Index) {
            Write-Host ("PBI_UI: KHONG thay element '{0}' (tim thay {1})" -f $Click, $hits.Count); exit 1
        }
        $e = $hits[$Index]
        # Khong phai element nao cung co InvokePattern. Tab trang cua Power BI la
        # SelectionItem, nut ribbon co the chi co LegacyIAccessible. Thu lan luot ba kieu,
        # neu khong thi bao ro - dung de loop tuong da bam duoc.
        $ok = $false
        foreach ($kieu in @("Invoke", "SelectionItem", "Legacy")) {
            try {
                switch ($kieu) {
                    "Invoke" {
                        $e.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke() }
                    "SelectionItem" {
                        $e.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select() }
                    "Legacy" {
                        $e.GetCurrentPattern(
                            [System.Windows.Automation.AutomationPattern]::LookupById(
                                "{1FEBE7D6-4F1D-4A47-9D6B-4C2D0F2B4E30}")).DoDefaultAction() }
                }
                Write-Host ("PBI_UI: da bam '{0}' ({1}/{2}) bang {3}" -f `
                            $e.Current.Name, ($Index + 1), $hits.Count, $kieu)
                $ok = $true; break
            } catch { }
        }
        if (-not $ok) {
            Write-Host ("PBI_UI: '{0}' khong bam duoc bang Invoke/SelectionItem/Legacy" -f $e.Current.Name)
            exit 1
        }
    }
}
exit 0
