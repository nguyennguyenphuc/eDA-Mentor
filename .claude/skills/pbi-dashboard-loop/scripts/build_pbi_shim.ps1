# Bien dich pbi_shim.cs -> pbi.exe (cung thu muc)
#
# Khong can cai gi: csc.exe di kem .NET Framework co san tren moi may Windows,
# con Microsoft.PowerBI.AdomdClient.dll do chinh Power BI Desktop dong goi.
#
#   powershell -ExecutionPolicy Bypass -File <skill>/scripts/build_pbi_shim.ps1

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$dll  = 'C:\Program Files\Microsoft Power BI Desktop\bin\Microsoft.PowerBI.AdomdClient.dll'
$csc  = Join-Path ([Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) 'csc.exe'

if (-not (Test-Path $dll)) { throw "Khong thay $dll -- Power BI Desktop da cai chua?" }
if (-not (Test-Path $csc)) { throw "Khong thay csc.exe tai $csc" }

$out = Join-Path $here 'pbi.exe'
if (Test-Path $out) { Remove-Item $out -Force }

& $csc /nologo /target:exe /out:"$out" `
    /reference:"$dll" /reference:"System.Data.dll" /reference:"System.Core.dll" `
    (Join-Path $here 'pbi_shim.cs')

if (-not (Test-Path $out)) { throw "Bien dich that bai" }
Write-Output "-> $out ($((Get-Item $out).Length) bytes)"
Write-Output ""
Write-Output "Dat thu muc nay len dau PATH neu may khong co pbi-cli (cac script goi pbi --json dax execute):"
Write-Output "  `$env:PATH = `"$here;`$env:PATH`""
