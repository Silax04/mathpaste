# Generates the add-in PNG icons (assets/*.png) with System.Drawing.
Add-Type -AssemblyName System.Drawing
$root = Split-Path -Parent $PSScriptRoot
$assets = Join-Path $root 'assets'
New-Item -ItemType Directory -Force $assets | Out-Null

function New-Icon([int]$size, [string]$text, [string]$path, [string]$bg) {
  $bmp = New-Object System.Drawing.Bitmap $size, $size
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAliasGridFit'
  $g.Clear([System.Drawing.Color]::Transparent)
  $r = [Math]::Max(2, [int]($size * 0.2))
  $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
  $d = $r * 2; $w = $size - 1
  $gp.AddArc(0, 0, $d, $d, 180, 90); $gp.AddArc($w - $d, 0, $d, $d, 270, 90)
  $gp.AddArc($w - $d, $w - $d, $d, $d, 0, 90); $gp.AddArc(0, $w - $d, $d, $d, 90, 90); $gp.CloseFigure()
  $brush = New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml($bg))
  $g.FillPath($brush, $gp)
  $font = New-Object System.Drawing.Font 'Cambria Math', ([float]($size * 0.5)), ([System.Drawing.FontStyle]::Regular), ([System.Drawing.GraphicsUnit]::Pixel)
  $sf = New-Object System.Drawing.StringFormat
  $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
  $rect = New-Object System.Drawing.RectangleF 0, ([float](-$size * 0.04)), $size, $size
  $g.DrawString($text, $font, [System.Drawing.Brushes]::White, $rect, $sf)
  $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $bmp.Dispose()
}
foreach ($s in 16, 32, 64, 80, 128) {
  New-Icon $s ([string][char]0x221A + 'x') (Join-Path $assets "icon-$s.png") '#185ABD'
}
foreach ($s in 16, 32, 80) {
  New-Icon $s ([string][char]0x2211) (Join-Path $assets "convert-$s.png") '#107C41'
}
Write-Host "Icons written to $assets"
