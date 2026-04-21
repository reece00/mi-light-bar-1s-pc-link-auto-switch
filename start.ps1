# 设置工作目录为脚本所在目录
Set-Location -Path $PSScriptRoot

Write-Host "正在启动米家挂灯控制器 Lite..." -ForegroundColor Cyan

if (Get-Command uv -ErrorAction SilentlyContinue) {
    uv run main.py
} else {
    Write-Error "错误: 未找到 uv，请确保已安装 uv 并添加到环境变量。"
    Read-Host "按回车键退出..."
}
