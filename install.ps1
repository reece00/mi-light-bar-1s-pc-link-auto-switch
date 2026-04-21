#Requires -Version 7.0
$ErrorActionPreference = "Stop"

$ProjectRoot = $PSScriptRoot

Write-Host "=== 米家挂灯控制器 Lite 安装脚本 ===" -ForegroundColor Cyan

# 检查 Python 版本
Write-Host "`n[1/5] 检查运行环境 (Python & uv)..." -ForegroundColor Yellow
$pythonVersion = python --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "错误: 未找到 Python，请先安装 Python 3.14+" -ForegroundColor Red
    exit 1
}

# 检查 uv
if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    Write-Host "未找到 uv，正在尝试自动安装..." -ForegroundColor Yellow
    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
    $env:Path += ";$env:USERPROFILE\.cargo\bin" # 临时添加到当前会话路径
    if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
        Write-Host "错误: uv 安装失败，请手动安装后重试。" -ForegroundColor Red
        exit 1
    }
}
if ($pythonVersion -notmatch "Python 3\.(1[4-9]|[2-9][0-9])") {
    Write-Host "警告: 当前 Python 版本可能过低，建议 3.14+" -ForegroundColor Yellow
}

# 创建虚拟环境
Write-Host "`n[2/5] 创建虚拟环境..." -ForegroundColor Yellow
if (Test-Path "$ProjectRoot\.venv") {
    Write-Host "  发现已有虚拟环境，跳过" -ForegroundColor Gray
} else {
    uv venv "$ProjectRoot\.venv" --python python
    if ($LASTEXITCODE -ne 0) { exit 1 }
}

# 安装依赖
Write-Host "`n[3/5] 安装依赖..." -ForegroundColor Yellow
uv pip install --python "$ProjectRoot\.venv\Scripts\python.exe"
if ($LASTEXITCODE -ne 0) { exit 1 }

# pywin32 后处理
Write-Host "`n[4/5] 配置 pywin32..." -ForegroundColor Yellow
$postInstall = "$ProjectRoot\.venv\Scripts\pywin32_postinstall.py"
if (Test-Path $postInstall) {
    & "$ProjectRoot\.venv\Scripts\python.exe" $postInstall -install
    if ($LASTEXITCODE -ne 0) {
        Write-Host "警告: pywin32 配置可能有问题，但可以尝试运行" -ForegroundColor Yellow
    }
} else {
    Write-Host "  pywin32 postinstall 脚本未找到，跳过" -ForegroundColor Gray
}

# 创建配置文件模板
Write-Host "`n[5/5] 检查配置文件..." -ForegroundColor Yellow
if (-not (Test-Path "$ProjectRoot\config.json")) {
    @"
{
    "ip": "",
    "token": "",
    "device_id": null,
    "auto_start": false,
    "on_boot": true,
    "on_sleep": true,
    "on_shutdown": true
}
"@ | Out-File -FilePath "$ProjectRoot\config.json" -Encoding UTF8
    Write-Host "  已创建 config.json 模板，请编辑填写 ip 和 token" -ForegroundColor Green
} else {
    Write-Host "  config.json 已存在" -ForegroundColor Gray
}

Write-Host "`n=== 安装完成 ===" -ForegroundColor Green
Write-Host "请编辑 config.json 填写设备 ip 和 token 后运行:" -ForegroundColor White
Write-Host "  uv run python main.py" -ForegroundColor Cyan