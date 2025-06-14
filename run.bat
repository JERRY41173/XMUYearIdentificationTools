@echo off
chcp 65001 >nul
echo ================================================
echo           学年鉴定表自动化处理工具
echo ================================================
echo.

echo 正在启动程序...
echo.

REM 激活虚拟环境并运行程序
call venv\Scripts\activate.bat && python src\main.py

echo.
echo 程序已退出。
pause
