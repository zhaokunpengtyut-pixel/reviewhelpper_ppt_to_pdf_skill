# Convert legacy .ppt / .pptx courseware to .pptx and export every slide as PNG.
# Requires Microsoft PowerPoint (COM). ASCII-only on purpose: PowerShell 5.1 reads
# .ps1 as ANSI unless the file has a BOM, so never put non-ASCII text in here.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File ppt2png.ps1 -InputDir <dir> [-PptxDir <dir>] [-SlideDir <dir>] [-Width 1600] [-Height 1200]
#
# Output layout:
#   <PptxDir>/ch01.pptx, ch02.pptx, ...      (one per source deck, in name order)
#   <PptxDir>/ch01.name.txt                  (original file name of the deck)
#   <SlideDir>/ch01/001.png ...              (1600x1200 by default, 4:3 decks)

param(
    [Parameter(Mandatory = $true)][string]$InputDir,
    [string]$PptxDir,
    [string]$SlideDir,
    [int]$Width = 1600,
    [int]$Height = 1200
)

$ErrorActionPreference = "Stop"
$InputDir = (Resolve-Path $InputDir).Path
if (-not $PptxDir)  { $PptxDir  = Join-Path $InputDir "_pptx" }
if (-not $SlideDir) { $SlideDir = Join-Path $InputDir "_slides" }
New-Item -ItemType Directory -Force -Path $PptxDir, $SlideDir | Out-Null

$decks = @(Get-ChildItem -Path $InputDir -File | Where-Object { $_.Extension -in ".ppt", ".pptx" } | Sort-Object Name)
if ($decks.Count -eq 0) { throw "no .ppt/.pptx found in $InputDir" }
Write-Host ("decks: " + $decks.Count)

$ppt = New-Object -ComObject PowerPoint.Application
try { $ppt.Visible = -1 } catch { }

$i = 0
foreach ($f in $decks) {
    $i++
    $tag = "ch{0:D2}" -f $i
    $pptxPath = Join-Path $PptxDir ($tag + ".pptx")
    if (Test-Path $pptxPath) { Remove-Item $pptxPath -Force }

    if ($f.Extension -eq ".pptx") {
        Copy-Item $f.FullName $pptxPath -Force
        $pres = $ppt.Presentations.Open($pptxPath, $true, $false, $false)
    } else {
        Write-Host ("convert: " + $f.Name)
        $pres = $ppt.Presentations.Open($f.FullName, $true, $false, $false)
        $pres.SaveAs($pptxPath, 24)   # 24 = ppSaveAsOpenXMLPresentation
    }

    $n = $pres.Slides.Count
    Write-Host ("  slides: " + $n)
    Set-Content -Path (Join-Path $PptxDir ($tag + ".name.txt")) -Value $f.BaseName -Encoding UTF8

    $dir = Join-Path $SlideDir $tag
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    for ($k = 1; $k -le $n; $k++) {
        $png = Join-Path $dir ("{0:D3}.png" -f $k)
        if (-not (Test-Path $png)) {
            $pres.Slides.Item($k).Export($png, "PNG", $Width, $Height)
        }
    }
    $pres.Close()
    Write-Host ("  exported -> " + $dir)
}

$ppt.Quit()
Write-Host "DONE"
