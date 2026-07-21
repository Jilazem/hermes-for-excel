# HermesExcel ikonlarini uretir (mor yuvarlak kare + beyaz 'H').
Add-Type -AssemblyName System.Drawing
$sizes = 16,32,64,80
$dir = $PSScriptRoot
foreach ($s in $sizes) {
  $bmp = New-Object System.Drawing.Bitmap($s, $s)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = 'AntiAlias'
  $g.Clear([System.Drawing.Color]::Transparent)
  $accent = [System.Drawing.Color]::FromArgb(255, 124, 58, 237)  # #7c3aed
  $brush = New-Object System.Drawing.SolidBrush($accent)
  $r = [Math]::Max(2, [int]($s * 0.18))
  # yuvarlak kare
  $path = New-Object System.Drawing.Drawing2D.GraphicsPath
  $d = $r * 2
  $path.AddArc(0, 0, $d, $d, 180, 90)
  $path.AddArc($s - $d, 0, $d, $d, 270, 90)
  $path.AddArc($s - $d, $s - $d, $d, $d, 0, 90)
  $path.AddArc(0, $s - $d, $d, $d, 90, 90)
  $path.CloseFigure()
  $g.FillPath($brush, $path)
  # 'H'
  $fontSize = [int]($s * 0.62)
  $font = New-Object System.Drawing.Font("Segoe UI", $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
  $sf = New-Object System.Drawing.StringFormat
  $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
  $white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
  $rect = New-Object System.Drawing.RectangleF(0, -1, $s, $s)
  $g.DrawString("H", $font, $white, $rect, $sf)
  $g.Dispose()
  $out = Join-Path $dir "icon-$s.png"
  $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
  $bmp.Dispose()
  Write-Host "yazildi: $out"
}
