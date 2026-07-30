<#
.SYNOPSIS
  使用 PyInstaller + UPX 将 PDF2Image 打包为单个 exe
#>

param([switch]$OneDir, [switch]$NoUPX)

$ErrorActionPreference = "Stop"
$ScriptDir = $PSScriptRoot
$OutputDir = Join-Path $ScriptDir "dist"

Write-Host ""
Write-Host "=== PDF2Image - 打包脚本 ===" -ForegroundColor Cyan
Write-Host ""

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "[错误] 未找到 Python" -ForegroundColor Red; exit 1
}

try { python -c "import PyInstaller" 2>$null | Out-Null }
catch { Write-Host "[信息] 安装 PyInstaller..."; pip install pyinstaller }

# 查找 UPX
$upxPath = $null
if (-not $NoUPX) {
    $candidates = @(
        Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Packages\UPX.UPX_Microsoft.Winget.Source_8wekyb3d8bbwe\upx-5.2.0-win64\upx.exe"
        "C:\Program Files\UPX\upx.exe"
        "C:\Program Files (x86)\UPX\upx.exe"
    )
    foreach ($c in $candidates) { if (Test-Path $c) { $upxPath = $c; break } }
    if (-not $upxPath) { $fromPath = (Get-Command "upx" -ErrorAction SilentlyContinue); if ($fromPath) { $upxPath = $fromPath.Source } }
    if ($upxPath) { Write-Host "[信息] UPX: $(& $upxPath --version 2>&1 | Select-Object -First 1)" -ForegroundColor Green }
    else { Write-Host "[提醒] 未找到 UPX" -ForegroundColor Yellow }
}

# 清理
foreach ($d in @("build", $OutputDir)) { if (Test-Path $d) { Remove-Item -Recurse -Force $d -ErrorAction SilentlyContinue } }

# 打包
$EntryPoint = Join-Path $ScriptDir "main.py"
$PyArgs = @(
    "--clean",
    "--name", "PDF2Image",
    "--distpath", $OutputDir,
    "--hidden-import", "pythoncom",
    "--hidden-import", "pywintypes",
    "--hidden-import", "win32com.client"
)

if ($OneDir) { $PyArgs += "--onedir" }
else { $PyArgs += "--onefile"; $PyArgs += "--noconsole" }

if ($upxPath) { $PyArgs += "--upx-dir"; $PyArgs += (Split-Path $upxPath -Parent) }

python -m PyInstaller @PyArgs $EntryPoint

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n=== 打包成功 ===" -ForegroundColor Green
    $outFile = Join-Path $OutputDir "PDF2Image.exe"
    if (Test-Path $outFile) {
        $size = [math]::Round((Get-Item $outFile).Length / 1MB, 1)
        Write-Host "输出: $outFile" -ForegroundColor Green
        Write-Host "大小: $size MB" -ForegroundColor Green
    }
} else {
    Write-Host "[错误] 打包失败" -ForegroundColor Red; exit 1
}

# 清理临时文件
Remove-Item (Join-Path $ScriptDir "build") -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $ScriptDir "PDF2Image.spec") -ErrorAction SilentlyContinue
