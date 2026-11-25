# 一键重建 Eimitriliaword Formatter Pro 的 exe（onedir 形态）
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File .\build.ps1
#   powershell -ExecutionPolicy Bypass -File .\build.ps1 -KeepData   # 保留运行时数据目录
#
# 做的事：
#   1. 停掉正在运行的实例（否则 dist 被占用）
#   2. 首次运行时创建 Python 3.11 构建环境（之后复用，约 40 秒）
#   3. PyInstaller 构建 onedir（约 12 秒）
#   4. 把源码 webui 同步到 dist 并校验哈希
#
# 说明：只有改了后端 wfp_web.py 才需要跑本脚本；
#       只改前端时直接把 webui\index.html 覆盖到
#       dist\WordFormatter\_internal\webui\ 即可，不必重新打包。

param(
    [switch]$KeepData,          # 保留 dist 下已有的运行时数据（out/tmp/window-profile）
    [string]$PythonLauncher = 'py -3.11'
)

$ErrorActionPreference = 'Stop'
$Root      = $PSScriptRoot
$BuildDir  = Join-Path $Root '_build'
$Venv      = Join-Path $BuildDir '.venv'
$Py        = Join-Path $Venv 'Scripts\python.exe'
$DistRoot  = Join-Path $Root 'dist'
$DistDir   = Join-Path $DistRoot 'Eimitriliaword Formatter Pro'
$Exe       = Join-Path $DistDir 'Eimitriliaword Formatter Pro.exe'
$Mirror    = 'https://pypi.tuna.tsinghua.edu.cn/simple'
$Watch     = [System.Diagnostics.Stopwatch]::StartNew()

# 临时目录隔离：本机 Python 写系统 TEMP 会失败（pywin32 的 gen_py 需要可写目录）
$BuildTmp = Join-Path $BuildDir 'tmp'
New-Item -ItemType Directory $BuildTmp -Force | Out-Null
$env:TMP = $BuildTmp; $env:TEMP = $BuildTmp

function Step([string]$msg) { Write-Host "`n==> $msg" -ForegroundColor Cyan }
function Ok([string]$msg)   { Write-Host "    $msg" -ForegroundColor Green }
function Warn([string]$msg) { Write-Host "    $msg" -ForegroundColor Yellow }

# ---------- 1. 停掉运行中的实例 ----------
Step '检查运行中的实例'
$procs = @(Get-Process -Name 'Eimitriliaword Formatter Pro' -ErrorAction SilentlyContinue)
if ($procs.Count -gt 0) {
    $procs | Stop-Process -Force -ErrorAction SilentlyContinue
    Ok "已停止 $($procs.Count) 个 Eimitriliaword Formatter Pro 进程"
    Start-Sleep -Milliseconds 800
} else {
    Ok '没有运行中的实例'
}
# 顺带关掉它开的应用窗口（Edge app 模式）
$wins = @(Get-Process msedge -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like 'Eimitriliaword*' })
if ($wins.Count -gt 0) {
    $wins | Stop-Process -Force -ErrorAction SilentlyContinue
    Ok '已关闭应用窗口'
}

# ---------- 2. 构建环境 ----------
Step '准备构建环境'
if (-not (Test-Path $Py)) {
    Warn '首次运行，正在创建 Python 3.11 环境（之后会复用）'
    New-Item -ItemType Directory $BuildDir -Force | Out-Null
    $tmp = Join-Path $BuildDir 'tmp'
    New-Item -ItemType Directory $tmp -Force | Out-Null
    $env:TMP = $tmp; $env:TEMP = $tmp      # 遵守全局规则：pip 临时目录隔离

    $parts = $PythonLauncher -split '\s+'
    & $parts[0] @($parts[1..($parts.Count-1)]) -m venv --system-site-packages $Venv
    if (-not (Test-Path $Py)) { throw "构建环境创建失败：找不到 $Py" }

    & $Py -m pip install --disable-pip-version-check --only-binary=:all: -q -i $Mirror pywin32
    if ($LASTEXITCODE -ne 0) { throw 'pywin32 安装失败' }
    Ok '构建环境就绪'
} else {
    Ok '复用已有构建环境'
}

$check = & $Py -c "import docx, lxml, fontTools, win32com.client, PyInstaller; print('ok')" 2>&1
if ($check -notmatch 'ok') { throw "构建环境依赖缺失：`n$check" }

# ---------- 3. 清理旧产物 ----------
Step '清理旧产物'
if (Test-Path $DistDir) {
    if ($KeepData) {
        # 只删可执行部分，保留 WordFormatterData
        Get-ChildItem $DistDir -Force | Where-Object { $_.Name -ne 'WordFormatterData' } |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        Ok '已清理构建产物（保留 WordFormatterData）'
    } else {
        Remove-Item $DistDir -Recurse -Force
        Ok '已删除旧 dist\WordFormatter'
    }
}
Remove-Item (Join-Path $Root 'build') -Recurse -Force -ErrorAction SilentlyContinue

# ---------- 4. PyInstaller 构建 ----------
Step '生成版本信息'
$verFile = Join-Path $Root 'version_info.txt'
& $Py (Join-Path $Root 'packaging\make_version_info.py') | Out-Host
if (-not (Test-Path $verFile)) { throw '版本信息生成失败' }
Ok 'version_info.txt 已生成'

Step 'PyInstaller 构建'
$buildTmp = Join-Path $BuildDir 'tmp'
New-Item -ItemType Directory $buildTmp -Force | Out-Null
$env:TMP = $buildTmp; $env:TEMP = $buildTmp

Push-Location $Root
try {
    & $Py -m PyInstaller --noconfirm --clean --onedir --noconsole `
        --name 'Eimitriliaword Formatter Pro' `
        --icon "assets\wfp-app.ico" `
        --version-file "$verFile" `
        --add-data "webui;webui" `
        --hidden-import win32com --hidden-import win32com.client `
        --hidden-import pythoncom --hidden-import pywintypes `
        wfp_web.py | Out-Host
} finally {
    Pop-Location
}
if (-not (Test-Path $Exe)) { throw '构建失败：没有生成 exe' }
Ok '构建完成'

# ---------- 5. 同步前端 ----------
Step '同步前端页面'
$srcHtml = Join-Path $Root 'webui\index.html'
$dstHtml = Join-Path $DistDir '_internal\webui\index.html'
if (-not (Test-Path $dstHtml)) { throw "找不到打包后的 webui：$dstHtml" }
Copy-Item $srcHtml $dstHtml -Force
$h1 = (Get-FileHash $srcHtml -Algorithm SHA256).Hash
$h2 = (Get-FileHash $dstHtml -Algorithm SHA256).Hash
if ($h1 -ne $h2) { throw '前端同步失败：哈希不一致' }
Ok '前端已同步（哈希一致）'

# ---------- 6. 汇报 ----------
$Watch.Stop()
$sum = (Get-ChildItem $DistDir -Recurse -File | Measure-Object Length -Sum).Sum
Step '完成'
Write-Host ("    exe      : {0:N1} MB" -f ((Get-Item $Exe).Length / 1MB))
Write-Host ("    整包     : {0:N1} MB / {1} 个文件" -f ($sum / 1MB), (Get-ChildItem $DistDir -Recurse -File).Count)
Write-Host ("    exe 时间 : {0}" -f (Get-Item $Exe).LastWriteTime)
Write-Host ("    耗时     : {0:N1} 秒" -f $Watch.Elapsed.TotalSeconds)
Write-Host ''
Write-Host '    成品: dist\Eimitriliaword Formatter Pro\Eimitriliaword Formatter Pro.exe' -ForegroundColor Green
