# Verify loop helper: build the deck, export selected slides as PNG, optionally export the PDF.
#
#   .\render-slides.ps1 -Script .\build-deck.ps1 -Slides 16,24 -Pdf
#
# Then LOOK at the PNGs in shots\. Never trust an edit you have not seen rendered.
[CmdletBinding()]
param(
    [string]   $Script  = '.\build-deck.ps1',
    [int[]]    $Slides  = @(),
    [switch]   $Pdf,
    [switch]   $SkipBuild,
    [string]   $OutDir  = 'shots',
    [int]      $Width   = 1400,
    [int]      $Height  = 788
)

$ErrorActionPreference = 'Stop'

$scriptPath = (Resolve-Path $Script).Path
$root = Split-Path -Parent $scriptPath

$busy = Get-Process POWERPNT -ErrorAction SilentlyContinue
if ($busy) {
    # Never kill it - it may hold the user's unsaved work in another file.
    throw 'PowerPoint is running. Close it before building, then run this again.'
}

if (-not $SkipBuild) {
    Write-Host 'Building...' -ForegroundColor Cyan
    # Windows PowerShell 5.1 for COM stability. The script must be pure ASCII or
    # written with a UTF-8 BOM - see references/PITFALLS.md #1.
    & powershell -NoProfile -ExecutionPolicy Bypass -File $scriptPath
    if ($LASTEXITCODE -ne 0) { throw "build failed (exit $LASTEXITCODE)" }
}

$pptx = Get-ChildItem -LiteralPath $root -Filter *.pptx |
        Where-Object { $_.Name -notlike '~$*' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
if (-not $pptx) { throw "no .pptx found in $root" }

$shots = Join-Path $root $OutDir
if (-not (Test-Path $shots)) { New-Item -ItemType Directory -Path $shots | Out-Null }

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$pres = $app.Presentations.Open($pptx.FullName, $true, $false, $false)   # read-only
try {
    $wanted = if ($Slides.Count) { $Slides } else { 1..$pres.Slides.Count }
    foreach ($n in $wanted) {
        if ($n -lt 1 -or $n -gt $pres.Slides.Count) {
            Write-Warning "slide $n out of range (1..$($pres.Slides.Count))"
            continue
        }
        $png = Join-Path $shots ('slide-{0:D2}.png' -f $n)
        $pres.Slides.Item($n).Export($png, 'PNG', $Width, $Height)
        Write-Host "  $png" -ForegroundColor DarkGray
    }

    if ($Pdf) {
        $pdf = [IO.Path]::ChangeExtension($pptx.FullName, '.pdf')
        if (Test-Path $pdf) { Remove-Item -LiteralPath $pdf -Force }
        $pres.SaveAs($pdf, 32)
        Write-Host "PDF: $pdf" -ForegroundColor Green
    }
} finally {
    if ($pres) { $pres.Close() }
    $app.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}

Write-Host "`nNow open the PNGs in $shots and look at them." -ForegroundColor Yellow
Write-Host 'Delete the folder when the slides are verified.' -ForegroundColor DarkGray
