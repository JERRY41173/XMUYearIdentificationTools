@echo off
chcp 65001 >nul
echo ================================================
echo           安装项目依赖包
echo ================================================
echo.

echo 检查虚拟环境状态...
if not exist "venv\Scripts\python.exe" (
    echo 错误：虚拟环境不存在！
    echo 请先运行 python -m venv venv 创建虚拟环境
    pause
    exit /b 1
)

echo 激活虚拟环境...
call venv\Scripts\activate.bat

echo.
echo 升级pip到最新版本...
python -m pip install --upgrade pip

echo.
echo 安装项目依赖包...
pip install -r requirements.txt

echo.
echo ================================================
echo           依赖包安装完成！
echo ================================================
echo.
echo 已安装的包列表：
pip list
echo.
echo 现在可以运行程序了：python src\main.py
echo.
pause
