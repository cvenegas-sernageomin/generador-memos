# make-icons.ps1 - genera los iconos PNG de la PWA (documento + matraz, tema verde SERNAGEOMIN)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
$root = $PSScriptRoot
$iconDir = Join-Path $root "pwa\icons"
if (-not (Test-Path $iconDir)) { New-Item -ItemType Directory -Path $iconDir | Out-Null }

function New-Icon([int]$size, [string]$path) {
  $bmp = New-Object System.Drawing.Bitmap($size, $size)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

  # Fondo verde a sangre
  $g.Clear([System.Drawing.Color]::FromArgb(31, 92, 77))

  # Hoja de memo (blanca)
  $m = $size * 0.20
  $w = $size - 2 * $m
  $h = $w * 1.18
  $x = $m; $y = ($size - $h) / 2.0
  $paper = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(248, 246, 240))
  $g.FillRectangle($paper, [single]$x, [single]$y, [single]$w, [single]$h)

  # Lineas de texto (grises)
  $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(180, 120, 110, 95)), ([single]($size * 0.018))
  $lx = $x + $w * 0.16; $lw = $w * 0.68
  for ($i = 0; $i -lt 4; $i++) {
    $ly = $y + $h * (0.22 + $i * 0.14)
    $ww = if ($i -eq 3) { $lw * 0.55 } else { $lw }
    $g.DrawLine($pen, [single]$lx, [single]$ly, [single]($lx + $ww), [single]$ly)
  }

  # Matraz (verde) sobre la esquina inferior derecha
  $fx = $x + $w * 0.60; $fy = $y + $h * 0.60
  $fw = $w * 0.34; $fh = $h * 0.30
  $flask = New-Object System.Drawing.Drawing2D.GraphicsPath
  $flask.AddPolygon(@(
    (New-Object System.Drawing.PointF([single]($fx + $fw*0.38), [single]$fy)),
    (New-Object System.Drawing.PointF([single]($fx + $fw*0.62), [single]$fy)),
    (New-Object System.Drawing.PointF([single]($fx + $fw*0.62), [single]($fy + $fh*0.38))),
    (New-Object System.Drawing.PointF([single]($fx + $fw), [single]($fy + $fh))),
    (New-Object System.Drawing.PointF([single]$fx, [single]($fy + $fh))),
    (New-Object System.Drawing.PointF([single]($fx + $fw*0.38), [single]($fy + $fh*0.38)))
  ))
  $fbr = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(43, 125, 104))
  $g.FillPath($fbr, $flask)

  $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
  $paper.Dispose(); $pen.Dispose(); $fbr.Dispose(); $g.Dispose(); $bmp.Dispose()
}

New-Icon 192 (Join-Path $iconDir "icon-192.png")
New-Icon 512 (Join-Path $iconDir "icon-512.png")
New-Icon 512 (Join-Path $iconDir "icon-maskable-512.png")
Write-Host "Iconos generados en $iconDir" -ForegroundColor Green
