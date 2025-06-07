@echo off
chcp 65001 >nul 2>&1
echo ======================================
echo   学年鉴定表自动化处理工具 v2.0
echo ======================================
echo.

:: 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python 3.7+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 显示Python版本
echo 检测到的Python版本:
python --version

:: 检查依赖项
echo.
echo 正在检查依赖项...
python -c "import sys; sys.path.append('src'); from src.utils.dependency_manager import check_dependencies; exit(0 if check_dependencies() else 1)"
if errorlevel 1 (
    echo.
    echo 依赖项检查失败，请先安装依赖项:
    echo pip install -r requirements.txt
    pause
    exit /b 1
)

:: 运行主程序
echo.
echo 启动程序...
echo ======================================
python src/main.py

echo.
echo 程序已退出。
pause
