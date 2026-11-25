Add-Type -AssemblyName System.Drawing

$dir = "$PSScriptRoot\..\assets\icon-options"
New-Item -ItemType Directory $dir -Force | Out-Null

function New-RoundedPath([int]$S, [double]$ratio) {
    $r = [int]($S * $ratio)
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $p.AddArc(0,0,$r*2,$r*2,180,90); $p.AddArc($S-$r*2,0,$r*2,$r*2,270,90)
    $p.AddArc($S-$r*2,$S-$r*2,$r*2,$r*2,0,90); $p.AddArc(0,$S-$r*2,$r*2,$r*2,90,90)
    $p.CloseFigure()
    return $p
}

function Draw-Align([int]$S, [switch]$Simple) {
    $bmp = New-Object System.Drawing.Bitmap $S, $S, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.PixelOffsetMode = 'HighQuality'
    $g.Clear([System.Drawing.Color]::Transparent)

    $path = New-RoundedPath $S 0.22
    $white = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
    $g.FillPath($white, $path)

    $inkGray  = [System.Drawing.ColorTranslator]::FromHtml("#C2C8D2")
    $inkBlue  = [System.Drawing.ColorTranslator]::FromHtml("#2F6BFF")

    if ($Simple) {
        # 小尺寸：2 灰参差 + 2 蓝对齐，条更粗
        $borderW = [float]([Math]::Max(1, $S*0.035))
        $g.DrawPath((New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml("#DDE1E8")), $borderW), $path)
        $pad = $S * 0.16
        $bw  = $S - 2*$pad
        $h   = $S * 0.13
        $gap = $S * 0.09
        $bGray = New-Object System.Drawing.SolidBrush $inkGray
        $bBlue = New-Object System.Drawing.SolidBrush $inkBlue
        $g.FillRectangle($bGray, [float]$pad, [float]($S*0.20), [float]($bw*0.85), [float]$h)
        $g.FillRectangle($bGray, [float]$pad, [float]($S*0.20 + $h + $gap), [float]($bw*0.55), [float]$h)
        $g.FillRectangle($bBlue, [float]$pad, [float]($S*0.55), [float]$bw, [float]$h)
        $g.FillRectangle($bBlue, [float]$pad, [float]($S*0.55 + $h + $gap), [float]$bw, [float]$h)
        $bGray.Dispose(); $bBlue.Dispose()
    } else {
        $borderW = [float]([Math]::Max(1, $S*0.012))
        $g.DrawPath((New-Object System.Drawing.Pen ([System.Drawing.ColorTranslator]::FromHtml("#E2E5EA")), $borderW), $path)
        $pad = $S * 0.18
        $bw  = $S - 2*$pad
        $h   = $S * 0.062
        $gap = $S * 0.052
        $bGray = New-Object System.Drawing.SolidBrush $inkGray
        $bBlue = New-Object System.Drawing.SolidBrush $inkBlue
        $grayW = @(0.92, 0.62, 0.80)
        for ($i=0; $i -lt 3; $i++) {
            $yy = $S*0.20 + $i*($h+$gap)
            $g.FillRectangle($bGray, [float]$pad, [float]$yy, [float]($bw*$grayW[$i]), [float]$h)
        }
        for ($i=0; $i -lt 3; $i++) {
            $yy = $S*0.56 + $i*($h+$gap)
            $g.FillRectangle($bBlue, [float]$pad, [float]$yy, [float]$bw, [float]$h)
        }
        $bGray.Dispose(); $bBlue.Dispose()
    }
    $g.Dispose()
    return $bmp
}

# ---------- 生成 ICO ----------
function To-Dib([System.Drawing.Bitmap]$bmp) {
    $S=$bmp.Width
    $rect=New-Object System.Drawing.Rectangle 0,0,$S,$S
    $d=$bmp.LockBits($rect,[System.Drawing.Imaging.ImageLockMode]::ReadOnly,[System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $raw=New-Object byte[] ($d.Stride*$S)
    [Runtime.InteropServices.Marshal]::Copy($d.Scan0,$raw,0,$raw.Length)
    $stride=$d.Stride; $bmp.UnlockBits($d)
    $xor=$S*$S*4; $andRow=[int]([math]::Ceiling($S/32.0)*4); $and=$andRow*$S
    $ms=New-Object IO.MemoryStream; $w=New-Object IO.BinaryWriter($ms)
    $w.Write([UInt32]40); $w.Write([Int32]$S); $w.Write([Int32]($S*2))
    $w.Write([UInt16]1); $w.Write([UInt16]32); $w.Write([UInt32]0); $w.Write([UInt32]$xor)
    $w.Write([Int32]0);$w.Write([Int32]0);$w.Write([UInt32]0);$w.Write([UInt32]0)
    for($y=$S-1;$y -ge 0;$y--){ $w.Write($raw,$y*$stride,$S*4) }
    $w.Write((New-Object byte[] $and),0,$and)
    $w.Flush(); $b=$ms.ToArray(); $w.Dispose(); $ms.Dispose(); return $b
}
function To-Png([System.Drawing.Bitmap]$bmp) {
    $ms=New-Object IO.MemoryStream
    $bmp.Save($ms,[System.Drawing.Imaging.ImageFormat]::Png)
    $b=$ms.ToArray(); $ms.Dispose(); return $b
}

$plan = @(
  @{s=256;png=$true}, @{s=128;png=$false}, @{s=64;png=$false}, @{s=48;png=$false},
  @{s=32;png=$false;simple=$true}, @{s=24;png=$false;simple=$true}, @{s=16;png=$false;simple=$true}
)
$imgs = @()
foreach($p in $plan){
    $bmp = if($p.simple){ Draw-Align $p.s -Simple } else { Draw-Align $p.s }
    $data = if($p.png){ To-Png $bmp } else { To-Dib $bmp }
    $bmp.Dispose()
    $imgs += [pscustomobject]@{ size=$p.s; data=$data }
}

$appDir = "$env:LOCALAPPDATA\WordFormatter"
New-Item -ItemType Directory $appDir -Force | Out-Null
$icoPath = Join-Path $appDir "wfp-align.ico"
$offset = 6 + 16*$imgs.Count
$fs=[IO.File]::Create($icoPath); $bw=New-Object IO.BinaryWriter($fs)
$bw.Write([UInt16]0);$bw.Write([UInt16]1);$bw.Write([UInt16]$imgs.Count)
foreach($i in $imgs){
    $dim = if($i.size -ge 256){0}else{$i.size}
    $bw.Write([Byte]$dim);$bw.Write([Byte]$dim);$bw.Write([Byte]0);$bw.Write([Byte]0)
    $bw.Write([UInt16]1);$bw.Write([UInt16]32)
    $bw.Write([UInt32]$i.data.Length);$bw.Write([UInt32]$offset)
    $offset += $i.data.Length
}
foreach($i in $imgs){ $bw.Write($i.data,0,$i.data.Length) }
$bw.Flush();$bw.Dispose();$fs.Close()
"ICO: $icoPath  ($([math]::Round((Get-Item $icoPath).Length/1KB,1)) KB, $($imgs.Count) 档)"

# ---------- 预览图 ----------
$sheet = New-Object System.Drawing.Bitmap 920, 300
$sg = [System.Drawing.Graphics]::FromImage($sheet)
$sg.SmoothingMode='AntiAlias'; $sg.TextRenderingHint='AntiAliasGridFit'
$sg.Clear([System.Drawing.Color]::FromArgb(255,247,247,248))
$font = New-Object System.Drawing.Font "Microsoft YaHei", 12, ([System.Drawing.FontStyle]::Regular), ([System.Drawing.GraphicsUnit]::Pixel)
$b256 = Draw-Align 256
$sg.DrawImage($b256, 20, 22, 256, 256)
$b256.Dispose()
$x = 310
foreach($sz in @(64,48,32,24,16)){
    $b = if($sz -le 32){ Draw-Align $sz -Simple } else { Draw-Align $sz }
    $sg.DrawImage($b, $x, 30, $sz, $sz)
    $sg.DrawString("$sz px", $font, [System.Drawing.Brushes]::Gray, $x, (30+$sz+6))
    $b.Dispose()
    $x += ($sz + 50)
}
$sg.Dispose()
$pv = Join-Path $dir "option1-final-preview.png"
$sheet.Save($pv, [System.Drawing.Imaging.ImageFormat]::Png)
$sheet.Dispose()
"预览: $pv"