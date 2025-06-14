# 学年鉴定表自动化处理工具 - 虚拟环境激活脚本（PowerShell版本）

Write-Host "================================================" -ForegroundColor Green
Write-Host "          学年鉴定表自动化处理工具" -ForegroundColor Green
Write-Host "           开发环境初始化脚本" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""

# 检查虚拟环境是否存在
if (-not (Test-Path "venv\Scripts\Activate.ps1")) {
    Write-Host "错误：虚拟环境不存在！" -ForegroundColor Red
    Write-Host "请先运行: python -m venv venv" -ForegroundColor Yellow
    Read-Host "按任意键退出"
    exit 1
}

Write-Host "正在激活虚拟环境..." -ForegroundColor Yellow
& ".\venv\Scripts\Activate.ps1"

Write-Host ""
Write-Host "虚拟环境已激活！" -ForegroundColor Green
Write-Host "项目路径: $PWD" -ForegroundColor Cyan
Write-Host "Python版本:" -ForegroundColor Cyan
python --version

Write-Host ""
Write-Host "您现在可以：" -ForegroundColor Cyan
Write-Host "1. 运行 .\install_deps.bat 来安装依赖包" -ForegroundColor White
Write-Host "2. 运行 python src\main.py 来启动程序" -ForegroundColor White
Write-Host "3. 使用 deactivate 命令退出虚拟环境" -ForegroundColor White
Write-Host ""
