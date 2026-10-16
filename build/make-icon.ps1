# Renders the app mark (same proportions as src/MailIcon.cs) into a multi-size .ico.
param([Parameter(Mandatory)][string]$Path)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function New-RoundedRect([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = 2 * $r
    $p.AddArc($x, $y, $d, $d, 180, 90)
    $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
    $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    $p
}

function New-Mark([int]$s) {
    $bmp = New-Object System.Drawing.Bitmap($s, $s, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.PixelOffsetMode = 'HighQuality'
    $blue = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 0x14, 0x50, 0xF5))
    $white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 0xFA, 0xF9, 0xF6))

    $g.FillPath($blue, (New-RoundedRect 0 0 $s $s (0.23 * $s)))
    $g.FillPath($white, (New-RoundedRect (0.21 * $s) (0.31 * $s) (0.576 * $s) (0.408 * $s) (0.075 * $s)))

    $pen = New-Object System.Drawing.Pen($blue.Color, [Math]::Max(0.048 * $s, 1.25))
    $pen.StartCap = 'Round'; $pen.EndCap = 'Round'; $pen.LineJoin = 'Round'
    $pts = [System.Drawing.PointF[]]@(
        (New-Object System.Drawing.PointF((0.19 * $s), (0.365 * $s))),
        (New-Object System.Drawing.PointF((0.5 * $s), (0.56 * $s))),
        (New-Object System.Drawing.PointF((0.81 * $s), (0.365 * $s))))
    $g.DrawLines($pen, $pts)

    $cx = 0.765 * $s; $cy = 0.314 * $s
    $r1 = 0.1 * $s; $r2 = 0.068 * $s
    $g.FillEllipse($blue, $cx - $r1, $cy - $r1, 2 * $r1, 2 * $r1)
    $g.FillEllipse($white, $cx - $r2, $cy - $r2, 2 * $r2, 2 * $r2)
    $g.Dispose()
    $bmp
}

# Classic DIB entry (BITMAPINFOHEADER + bottom-up BGRA + AND mask), readable everywhere.
function Get-DibBytes([System.Drawing.Bitmap]$bmp) {
    $s = $bmp.Width
    $ms = New-Object System.IO.MemoryStream
    $w = New-Object System.IO.BinaryWriter($ms)
    $w.Write([int]40); $w.Write([int]$s); $w.Write([int]($s * 2)); $w.Write([int16]1); $w.Write([int16]32)
    $w.Write([int]0); $w.Write([int]0); $w.Write([int]0); $w.Write([int]0); $w.Write([int]0); $w.Write([int]0)
    for ($y = $s - 1; $y -ge 0; $y--) {
        for ($x = 0; $x -lt $s; $x++) {
            $c = $bmp.GetPixel($x, $y)
            $w.Write([byte]$c.B); $w.Write([byte]$c.G); $w.Write([byte]$c.R); $w.Write([byte]$c.A)
        }
    }
    $maskRow = [int]([Math]::Ceiling($s / 32.0) * 4)
    $w.Write((New-Object byte[] ($maskRow * $s)))
    $w.Flush()
    $ms.ToArray()
}

$sizes = 16, 20, 24, 32, 40, 48, 64, 256
$images = foreach ($s in $sizes) {
    $bmp = New-Mark $s
    if ($s -eq 256) {
        $ms = New-Object System.IO.MemoryStream
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        , $ms.ToArray()
    } else {
        , (Get-DibBytes $bmp)
    }
    $bmp.Dispose()
}

$out = New-Object System.IO.MemoryStream
$w = New-Object System.IO.BinaryWriter($out)
$w.Write([int16]0); $w.Write([int16]1); $w.Write([int16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $s = $sizes[$i]; $data = $images[$i]
    $dim = if ($s -ge 256) { 0 } else { $s }
    $w.Write([byte]$dim); $w.Write([byte]$dim); $w.Write([byte]0); $w.Write([byte]0)
    $w.Write([int16]1); $w.Write([int16]32); $w.Write([int]$data.Length); $w.Write([int]$offset)
    $offset += $data.Length
}
foreach ($data in $images) { $w.Write([byte[]]$data) }
$w.Flush()
[System.IO.File]::WriteAllBytes($Path, $out.ToArray())

# Also keep a PNG preview of the mark next to the icon.
$preview = New-Mark 256
$preview.Save([System.IO.Path]::ChangeExtension($Path, '.png'), [System.Drawing.Imaging.ImageFormat]::Png)
$preview.Dispose()
