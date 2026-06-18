# render.ps1 — build the deck, convert to PDF, then rasterize every slide to JPG for QA.
# Usage:  powershell -ExecutionPolicy Bypass -File render.ps1
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

$soffice  = "C:\Program Files\LibreOffice\program\soffice.exe"
$pdftoppm = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\poppler-25.07.0\Library\bin\pdftoppm.exe"

Write-Host "1/3  building .pptx ..." -ForegroundColor Cyan
node build.js
if (-not $?) { throw "build.js failed" }

Write-Host "2/3  pptx -> pdf (LibreOffice headless) ..." -ForegroundColor Cyan
# kill any stale soffice that would block headless conversion
Get-Process soffice* -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
# soffice prints a harmless "Could not find platform independent libraries" warning to
# stderr; don't let it abort the script. The PDF is still produced.
$ErrorActionPreference = "Continue"
& $soffice --headless --convert-to pdf --outdir $here "Agentic-AI-Explained.pptx" 2>$null | Out-Null
$ErrorActionPreference = "Stop"
if (-not (Test-Path "$here\Agentic-AI-Explained.pdf")) { throw "PDF was not produced" }

Write-Host "3/3  pdf -> jpg (poppler, 120 dpi) ..." -ForegroundColor Cyan
Get-ChildItem "$here\slide-*.jpg" -ErrorAction SilentlyContinue | Remove-Item -Force
& $pdftoppm -jpeg -r 120 "Agentic-AI-Explained.pdf" "slide" | Out-Null

$imgs = Get-ChildItem "$here\slide-*.jpg" | Sort-Object Name
Write-Host ("done. {0} slides rendered:" -f $imgs.Count) -ForegroundColor Green
$imgs | ForEach-Object { Write-Host ("  " + $_.FullName) }
