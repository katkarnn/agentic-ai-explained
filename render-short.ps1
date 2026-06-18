# render-short.ps1 — regenerate + render the SHORT 20-slide LinkedIn cut.
# Outputs Agentic-AI-Explained-Short.pdf and short-slide-01..20.jpg (for a
# multi-image post — LinkedIn caps multi-image at 20). Master deck untouched.
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

$soffice  = "C:\Program Files\LibreOffice\program\soffice.exe"
$pdftoppm = "$env:LOCALAPPDATA\Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\poppler-25.07.0\Library\bin\pdftoppm.exe"

Write-Host "0/3  deriving build-short.js from build.js ..." -ForegroundColor Cyan
node make-short.js
if (-not $?) { throw "make-short.js failed" }

Write-Host "1/3  building short .pptx ..." -ForegroundColor Cyan
node build-short.js
if (-not $?) { throw "build-short.js failed" }

Write-Host "2/3  pptx -> pdf (LibreOffice headless) ..." -ForegroundColor Cyan
Get-Process soffice* -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
$ErrorActionPreference = "Continue"
& $soffice --headless --convert-to pdf --outdir $here "Agentic-AI-Explained-Short.pptx" 2>$null | Out-Null
$ErrorActionPreference = "Stop"
if (-not (Test-Path "$here\Agentic-AI-Explained-Short.pdf")) { throw "PDF was not produced" }

Write-Host "3/3  pdf -> jpg (poppler, 120 dpi) ..." -ForegroundColor Cyan
Get-ChildItem "$here\short-slide-*.jpg" -ErrorAction SilentlyContinue | Remove-Item -Force
& $pdftoppm -jpeg -r 120 "Agentic-AI-Explained-Short.pdf" "short-slide" 2>$null | Out-Null

$imgs = Get-ChildItem "$here\short-slide-*.jpg" | Sort-Object Name
Write-Host ("done. {0} slides rendered." -f $imgs.Count) -ForegroundColor Green
