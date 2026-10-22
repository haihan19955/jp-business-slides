# Dump how PowerPoint actually lays out every text frame: rendered lines per paragraph,
# text bound height vs. box height (overflow), font info. Output: JSON for jp_lint.py --layout.
# Usage: powershell -ExecutionPolicy Bypass -File export_layout.ps1 -Pptx deck.pptx -Out layout.json
# Windows + PowerPoint only. ASCII-only on purpose (Windows PowerShell 5.1 misreads BOM-less UTF-8).
param(
  [Parameter(Mandatory = $true)][string]$Pptx,
  [string]$Out = "layout.json"
)
$ErrorActionPreference = "Stop"
$pptxPath = (Resolve-Path $Pptx).Path
$items = New-Object System.Collections.ArrayList

function Add-Frame($slideIdx, $shapeId, $name, $kind, $tf, $boxW, $boxH) {
  if (-not $tf.HasText) { return }
  $tr = $tf.TextRange
  $paras = @()
  $pc = $tr.Paragraphs().Count
  for ($p = 1; $p -le $pc; $p++) {
    $para = $tr.Paragraphs($p, 1)
    $lines = @()
    for ($l = 1; $l -le 60; $l++) {
      $ln = $para.Lines($l, 1)
      if ($ln.Length -eq 0) { break }
      $lines += , ($ln.Text -replace "[\r\n\v]+$", "")
    }
    $paras += , @{ index = $p - 1; text = ($para.Text -replace "[\r\n\v]+$", ""); lines = $lines; size = $para.Font.Size; font = $para.Font.Name; fontEA = $para.Font.NameFarEast }
  }
  $innerH = $boxH - $tf.MarginTop - $tf.MarginBottom
  [void]$items.Add(@{
    slide = $slideIdx; id = $shapeId; shape = $name; kind = $kind
    boxW = [math]::Round($boxW / 72, 3); boxH = [math]::Round($boxH / 72, 3)
    textH = [math]::Round($tr.BoundHeight / 72, 3); innerH = [math]::Round($innerH / 72, 3)
    autoSize = $tf.AutoSize; wordWrap = $tf.WordWrap
    paragraphs = $paras
  })
}

function Walk-Shape($slideIdx, $shp) {
  if ($shp.Type -eq 6) { foreach ($c in $shp.GroupItems) { Walk-Shape $slideIdx $c }; return }
  if ($shp.HasTable) {
    $t = $shp.Table
    for ($r = 1; $r -le $t.Rows.Count; $r++) {
      for ($c = 1; $c -le $t.Columns.Count; $c++) {
        $cell = $t.Cell($r, $c).Shape
        Add-Frame $slideIdx $shp.Id ("{0}[r{1}c{2}]" -f $shp.Name, $r, $c) "table" $cell.TextFrame $t.Columns($c).Width $t.Rows($r).Height
      }
    }
    return
  }
  if ($shp.HasTextFrame) { Add-Frame $slideIdx $shp.Id $shp.Name "text" $shp.TextFrame $shp.Width $shp.Height }
}

$app = New-Object -ComObject PowerPoint.Application
try {
  $pres = $app.Presentations.Open($pptxPath, 1, 0, 0)
  foreach ($s in $pres.Slides) { foreach ($shp in $s.Shapes) { Walk-Shape $s.SlideIndex $shp } }
  $pres.Close()
} finally {
  $app.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
}
$json = ConvertTo-Json -InputObject @($items) -Depth 8
$outFile = if ([System.IO.Path]::IsPathRooted($Out)) { $Out } else { Join-Path (Get-Location) $Out }
[System.IO.File]::WriteAllText($outFile, $json, (New-Object System.Text.UTF8Encoding($false)))
Write-Output ("wrote {0} text frames to {1}" -f $items.Count, $Out)
