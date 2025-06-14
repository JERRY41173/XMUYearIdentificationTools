@echo off
chcp 65001 >nul
echo ================================================
echo           学年鉴定表自动化处理工具
echo            开发环境初始化脚本
echo ================================================
echo.

echo 正在激活虚拟环境...
call venv\Scripts\activate.bat

echo.
echo 虚拟环境已激活！
echo 项目路径: %cd%
echo Python版本: 
python --version
echo.

echo 您现在可以：
echo 1. 运行 install_deps.bat 来安装依赖包
echo 2. 运行 python src\main.py 来启动程序
echo 3. 使用 deactivate 命令退出虚拟环境
echo.

cmd /k
