@echo off
chcp 65001 >nul
echo ================================================
echo          开发环境设置脚本
echo ================================================
echo.

echo 正在检查Python环境...
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo 错误：未找到Python！请先安装Python 3.8+
    pause
    exit /b 1
)

echo 当前Python版本：
python --version

echo.
echo 正在创建虚拟环境...
if exist "venv" (
    echo 虚拟环境已存在，是否重新创建？[Y/N]
    set /p choice=
    if /i "%choice%"=="Y" (
        rmdir /s /q venv
        python -m venv venv
    )
) else (
    python -m venv venv
)

echo.
echo 正在激活虚拟环境...
call venv\Scripts\activate.bat

echo.
echo 正在升级pip...
python -m pip install --upgrade pip

echo.
echo 正在安装开发依赖...
pip install -r requirements-dev.txt

echo.
echo ================================================
echo          开发环境设置完成！
echo ================================================
echo.
echo 开发环境已准备就绪，您现在可以：
echo 1. 运行程序：python src\main.py
echo 2. 运行测试：pytest tests/
echo 3. 代码格式化：black src/ tests/
echo 4. 代码检查：flake8 src/ tests/
echo 5. 类型检查：mypy src/
echo 6. 打包程序：python tools\build.py
echo.
echo 使用 deactivate 命令退出虚拟环境
echo.
pause
