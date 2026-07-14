param(
  [Parameter(Mandatory=$true)][string]$MediaDir,
  [Parameter(Mandatory=$true)][string]$OutputDir
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$media = [System.IO.Path]::GetFullPath($MediaDir)
$out = [System.IO.Path]::GetFullPath($OutputDir)
$workspace = [System.IO.Path]::GetFullPath((Get-Location).Path)
if (-not $media.StartsWith($workspace, [System.StringComparison]::OrdinalIgnoreCase) -or
    -not $out.StartsWith($workspace, [System.StringComparison]::OrdinalIgnoreCase)) {
  throw 'MediaDir and OutputDir must stay inside the current workspace.'
}
if (-not (Test-Path -LiteralPath $out)) { New-Item -ItemType Directory -Path $out | Out-Null }

$files = Get-ChildItem -LiteralPath $media -File | Sort-Object {
  if ($_.BaseName -match '(\d+)$') { [int]$Matches[1] } else { 0 }
}
$cols=4; $rows=4; $cellW=400; $cellH=330; $labelH=28
$perSheet=$cols*$rows
$font = New-Object System.Drawing.Font('Arial',12,[System.Drawing.FontStyle]::Bold)
$brush = [System.Drawing.Brushes]::Black
$bg = [System.Drawing.Color]::White

for($s=0; $s -lt [Math]::Ceiling($files.Count/$perSheet); $s++) {
  $bmp = [System.Drawing.Bitmap]::new($cols*$cellW,$rows*$cellH)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.Clear($bg)
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  for($j=0; $j -lt $perSheet; $j++) {
    $idx=$s*$perSheet+$j; if($idx -ge $files.Count){break}
    $f=$files[$idx]; $col=$j%$cols; $row=[Math]::Floor($j/$cols)
    $x=$col*$cellW; $y=$row*$cellH
    try {
      $img=[System.Drawing.Image]::FromFile($f.FullName)
      $scale=[Math]::Min(($cellW-12)/$img.Width,($cellH-$labelH-12)/$img.Height)
      $w=[int]($img.Width*$scale); $h=[int]($img.Height*$scale)
      $dx=$x+[int](($cellW-$w)/2); $dy=$y+$labelH+[int](($cellH-$labelH-$h)/2)
      $g.DrawImage($img,$dx,$dy,$w,$h)
      $img.Dispose()
    } catch {
      $g.DrawString('unrenderable',$font,[System.Drawing.Brushes]::Red,$x+8,$y+50)
    }
    $g.DrawString($f.Name,$font,$brush,$x+6,$y+4)
    $g.DrawRectangle([System.Drawing.Pens]::LightGray,$x,$y,$cellW-1,$cellH-1)
  }
  $path=Join-Path $out ('sheet_{0:D2}.png' -f ($s+1))
  $bmp.Save($path,[System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $bmp.Dispose()
}
$font.Dispose()
Write-Output ("sheets={0}; files={1}" -f [Math]::Ceiling($files.Count/$perSheet),$files.Count)
