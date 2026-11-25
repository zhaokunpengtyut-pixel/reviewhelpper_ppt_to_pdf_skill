# 用 Windows.Data.Pdf (WinRT) 把 PDF 逐页栅格化成 PNG，供 read_image 目视校验。
#
# 用法：
#   & pdf2png.ps1 <pdfPath> <outDir> [maxPage=3]
#
# 说明：
#   - 仅 Windows + PowerShell 5.1/7 可用；无需安装 poppler/ghostscript/magick。
#   - 生成的 pdfpage_01.png ... 可直接用 read_image 打开逐页检查排版。
#   - 想校验全书时传一个较大的 maxPage（如 200），再按文件大小筛“近空页”。

param(
    [Parameter(Mandatory = $true)][string]$PdfPath,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [int]$MaxPage = 3
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Runtime.WindowsRuntime | Out-Null

$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]
$asTaskAction = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'
})[0]

function Await($op, $type) {
    $m = $asTaskGeneric.MakeGenericMethod($type)
    $t = $m.Invoke($null, @($op))
    $t.Wait(-1) | Out-Null
    $t.Result
}

[Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime] | Out-Null

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }

$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync((Resolve-Path $PdfPath).Path)) ([Windows.Storage.StorageFile])
$doc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
Write-Host ("pages: " + $doc.PageCount)

$n = [Math]::Min($MaxPage, $doc.PageCount)
for ($i = 0; $i -lt $n; $i++) {
    $page = $doc.GetPage($i)
    $ms = New-Object System.IO.MemoryStream
    $ras = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($ms)
    $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions
    $opt.DestinationWidth = 1200
    $act = $page.RenderToStreamAsync($ras, $opt)
    $asTaskAction.Invoke($null, @($act)).Wait(-1) | Out-Null
    $bytes = $ms.ToArray()
    $out = Join-Path $OutDir ("pdfpage_{0:D3}.png" -f ($i + 1))
    [System.IO.File]::WriteAllBytes($out, $bytes)
    Write-Host ("wrote " + $out + " " + $bytes.Length)
    $page.Dispose()
}
