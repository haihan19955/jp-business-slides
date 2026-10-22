# Export every slide to PNG using real PowerPoint (Windows + PowerPoint required).
# Usage: powershell -ExecutionPolicy Bypass -File render_slides.ps1 -Pptx deck.pptx -OutDir renders
# Rendering with PowerPoint itself shows Japanese line wrapping and overflow exactly as the audience will see it.
# (Kept ASCII-only: Windows PowerShell 5.1 misreads UTF-8 files without a BOM.)
param(
  [Parameter(Mandatory = $true)][string]$Pptx,
  [string]$OutDir = "renders",
  [int]$Width = 1600
)
$ErrorActionPreference = "Stop"
$pptxPath = (Resolve-Path $Pptx).Path
New-Item -ItemType Directory -Force $OutDir | Out-Null
$outPath = (Resolve-Path $OutDir).Path
Get-ChildItem $outPath -Filter "slide-*.png" | Remove-Item -Force

$app = New-Object -ComObject PowerPoint.Application
try {
  # ReadOnly=1, Untitled=0, WithWindow=0
  $pres = $app.Presentations.Open($pptxPath, 1, 0, 0)
  $h = [int]($Width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
  foreach ($s in $pres.Slides) {
    $file = Join-Path $outPath ("slide-{0:D2}.png" -f $s.SlideIndex)
    $s.Export($file, "PNG", $Width, $h)
    Write-Output $file
  }
  $pres.Close()
} finally {
  $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
