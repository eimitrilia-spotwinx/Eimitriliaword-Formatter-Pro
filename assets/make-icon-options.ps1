Add-Type -AssemblyName System.Drawing

$outDir = "$PSScriptRoot\icon-options"
New-Item -ItemType Directory $outDir -Force | Out-Null

function New-RoundedPath([int]$S, [double]$ratio) {
    $r = [int]($S * $ratio)
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $p.AddArc(0, 0, $r*2, $r*2, 180, 90)
    $p.AddArc($S-$r*2, 0, $r*2, $r*2, 270, 90)
    $p.AddArc($S-$r*2, $S-$r*2, $r*2, $r*2, 0, 90)
    $p.AddArc(0, $S-$r*2, $r*2, $r*2, 90, 90)
    $p.CloseFigure()
    return $p
}

function Draw-Bars($g, [int]$S, [double]$x, [double]$y, [double]$w, [double]$h, [double]$gap, $color, [double[]]$widths) {
    $brush = New-Object System.Drawing.SolidBrush $color
    for ($i = 0; $i -lt $widths.Count; $i++) {
        $yy = $y + $i * ($h + $gap)
        $g.FillRectangle($brush, $x, $yy, [double]($w * $widths[$i]), $h)
    }
    $brush.Dispose()
}

# ---------- 各方案绘制 ----------
function Draw-1-Align([int]$S) {
    $bmp = New-Object System.Drawing.Bitmap $S, $S, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $path = New-RoundedPath $S 0.22
    $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)), $path)
    $pen = New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml("#E2E5EA")), ([float]([Math]::Max(1,$S*0.012)))
    $g.DrawPath($pen, $path)
    $pad = $S * 0.18
    $bw = $S - 2*$pad
    $h  = $S * 0.062
    $gap = $S * 0.052
    # 上半：参差
    Draw-Bars $g $S $pad ($S*0.20) $bw $h $gap ([System.Drawing.ColorTranslator]::FromHtml("#C9CED6")) @(0.92,0.62,0.80)
    # 下半：对齐
    Draw-Bars $g $S $pad ($S*0.56) $bw $h $gap ([System.Drawing.ColorTranslator]::FromHtml("#2F6BFF")) @(1.0,1.0,1.0)
    $g.Dispose(); $pen.Dispose()
    return $bmp
}

function Draw-2-Doc([int]$S) {
    $bmp = New-Object System.Drawing.Bitmap $S, $S, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $bg = New-RoundedPath $S 0.22
    $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#F2F5FF"))), $bg)
    # 纸张
    $px = $S*0.24; $py = $S*0.15; $pw = $S*0.52; $ph = $S*0.70
    $fold = $S*0.16
    $paper = New-Object System.Drawing.Drawing2D.GraphicsPath
    $paper.AddPolygon(@(
        (New-Object System.Drawing.PointF $px, $py),
        (New-Object System.Drawing.PointF ($px+$pw-$fold), $py),
        (New-Object System.Drawing.PointF ($px+$pw), ($py+$fold)),
        (New-Object System.Drawing.PointF ($px+$pw), ($py+$ph)),
        (New-Object System.Drawing.PointF $px, ($py+$ph))
    ))
    $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)), $paper)
    $g.DrawPath((New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml("#C9D6F5")), [float]([Math]::Max(1,$S*0.012))), $paper)
    # 折角
    $g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml("#C9D6F5")), [float]([Math]::Max(1,$S*0.012))),
        [float]($px+$pw-$fold), [float]$py, [float]($px+$pw-$fold), [float]($py+$fold))
    $g.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml("#C9D6F5")), [float]([Math]::Max(1,$S*0.012))),
        [float]($px+$pw-$fold), [float]($py+$fold), [float]($px+$pw), [float]($py+$fold))
    # 对齐的三行
    Draw-Bars $g $S ($px+$pw*0.16) ($py+$ph*0.42) ($pw*0.68) ($S*0.045) ($S*0.055) ([System.Drawing.ColorTranslator]::FromHtml("#2F6BFF")) @(1.0,1.0,1.0)
    $g.Dispose()
    return $bmp
}

function Draw-3-W([int]$S) {
    $bmp = New-Object System.Drawing.Bitmap $S, $S, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.TextRenderingHint = 'AntiAliasGridFit'
    $g.Clear([System.Drawing.Color]::Transparent)
    $path = New-RoundedPath $S 0.227
    $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#2F6BFF"))), $path)
    $font = New-Object System.Drawing.Font "Segoe UI", ([float]($S*0.56)), ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
    $rect = New-Object System.Drawing.RectangleF 0, ($S*0.02), $S, $S
    # 左半深白，右半淡蓝
    $g.SetClip((New-Object System.Drawing.RectangleF 0,0,($S*0.52),$S))
    $g.DrawString("W", $font, [System.Drawing.Brushes]::White, $rect, $sf)
    $g.SetClip((New-Object System.Drawing.RectangleF ($S*0.52),0,($S*0.48),$S))
    $g.DrawString("W", $font, (New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#BFD4FF"))), $rect, $sf)
    $g.ResetClip()
    $g.Dispose(); $font.Dispose()
    return $bmp
}

function Draw-4-Seal([int]$S) {
    $bmp = New-Object System.Drawing.Bitmap $S, $S, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $path = New-RoundedPath $S 0.18
    $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#FBF8F1"))), $path)
    $g.DrawPath((New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml("#2F6BFF")), [float]([Math]::Max(1,$S*0.022))), $path)
    # 文字行
    Draw-Bars $g $S ($S*0.16) ($S*0.24) ($S*0.52) ($S*0.05) ($S*0.075) ([System.Drawing.ColorTranslator]::FromHtml("#B9BEC7")) @(1.0,0.86,0.92)
    # 红章
    $cx = $S*0.66; $cy = $S*0.66; $rad = $S*0.19
    $pen = New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml("#D64541")), [float]([Math]::Max(1,$S*0.028))
    $g.DrawEllipse($pen, [float]($cx-$rad), [float]($cy-$rad), [float]($rad*2), [float]($rad*2))
    # 五角星（简化：一个小圆点阵）
    $g.FillEllipse((New-Object System.Drawing.SolidBrush ([System.Drawing.ColorTranslator]::FromHtml("#D64541"))),
        [float]($cx-$rad*0.42), [float]($cy-$rad*0.42), [float]($rad*0.84), [float]($rad*0.84))
    $g.Dispose()
    return $bmp
}

function Draw-5-Mono([int]$S) {
    $bmp = New-Object System.Drawing.Bitmap $S, $S, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.Clear([System.Drawing.Color]::Transparent)
    $ink = [System.Drawing.ColorTranslator]::FromHtml("#3A3F47")
    $w = [float]([Math]::Max(1.2, $S*0.035))
    $pen = New-Object System.Drawing.Pen $ink, $w
    # 文档外框
    $x = $S*0.22; $y = $S*0.14; $ww = $S*0.50; $hh = $S*0.72
    $fold = $S*0.15
    $g.DrawLine($pen, [float]$x, [float]$y, [float]($x+$ww-$fold), [float]$y)
    $g.DrawLine($pen, [float]($x+$ww-$fold), [float]$y, [float]($x+$ww), [float]($y+$fold))
    $g.DrawLine($pen, [float]($x+$ww), [float]($y+$fold), [float]($x+$ww), [float]($y+$hh))
    $g.DrawLine($pen, [float]($x+$ww), [float]($y+$hh), [float]$x, [float]($y+$hh))
    $g.DrawLine($pen, [float]$x, [float]($y+$hh), [float]$x, [float]$y)
    # 三行
    for($i=0; $i -lt 3; $i++){
        $ly = $y + $hh*(0.30 + $i*0.18)
        $g.DrawLine($pen, [float]($x+$ww*0.14), [float]$ly, [float]($x+$ww*0.80), [float]$ly)
    }
    # 左侧标尺刻度
    $rx = $S*0.13
    for($i=0; $i -lt 4; $i++){
        $ty = $y + $hh*(0.25 + $i*0.17)
        $len = if($i % 2 -eq 0){ $S*0.055 } else { $S*0.035 }
        $g.DrawLine($pen, [float]$rx, [float]$ty, [float]($rx+$len), [float]$ty)
    }
    $g.Dispose(); $pen.Dispose()
    return $bmp
}

# ---------- 生成对照表 ----------
$designs = @(
    @{ id=1; name='1 段落对齐型'; fn=${function:Draw-1-Align} },
    @{ id=2; name='2 文档型';     fn=${function:Draw-2-Doc} },
    @{ id=3; name='3 W 字母双色'; fn=${function:Draw-3-W} },
    @{ id=4; name='4 公文印章型'; fn=${function:Draw-4-Seal} },
    @{ id=5; name='5 单色线框型'; fn=${function:Draw-5-Mono} }
)

$rowH = 210; $sheetW = 780
$sheet = New-Object System.Drawing.Bitmap $sheetW, ($rowH * $designs.Count + 20)
$sg = [System.Drawing.Graphics]::FromImage($sheet)
$sg.SmoothingMode = 'AntiAlias'
$sg.TextRenderingHint = 'AntiAliasGridFit'
$sg.Clear([System.Drawing.Color]::FromArgb(255,250,250,251))

$labelFont = New-Object System.Drawing.Font "Microsoft YaHei", 13, ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
$smallFont = New-Object System.Drawing.Font "Microsoft YaHei", 10, ([System.Drawing.FontStyle]::Regular), ([System.Drawing.GraphicsUnit]::Pixel)

$y = 10
foreach($d in $designs){
    # 行分隔线
    $sg.DrawLine((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(255,226,229,234)), 1), 0, ($y+$rowH-8), $sheetW, ($y+$rowH-8))
    # 标签
    $sg.DrawString($d.name, $labelFont, [System.Drawing.Brushes]::Black, 16, ($y+90))

    # 128 主图
    $b128 = & $d.fn 128
    $sg.DrawImage($b128, 210, ($y+40), 128, 128)
    $b128.Dispose()

    # 小尺寸
    $x = 370
    foreach($sz in @(48,32,24,16)){
        $b = & $d.fn $sz
        $sg.DrawImage($b, $x, ($y+80), $sz, $sz)
        $sg.DrawString("$sz", $smallFont, [System.Drawing.Brushes]::Gray, $x, ($y+132))
        $b.Dispose()
        $x += ($sz + 26)
    }
    $y += $rowH
}
$sg.Dispose()

$sheetPath = Join-Path $outDir "all-options.png"
$sheet.Save($sheetPath, [System.Drawing.Imaging.ImageFormat]::Png)
$sheet.Dispose()

# 单独导出 256 版本
foreach($d in $designs){
    $b = & $d.fn 256
    $b.Save((Join-Path $outDir ("opt{0}.png" -f $d.id)), [System.Drawing.Imaging.ImageFormat]::Png)
    $b.Dispose()
}
"已生成: $sheetPath"
Get-ChildItem $outDir | Select-Object Name,@{n='KB';e={[math]::Round($_.Length/1KB,1)}} | Format-Table -AutoSize
