@echo off
chcp 65001 >nul 2>&1
echo =============================================
echo 清理重构前的旧文件
echo =============================================
echo.

echo 以下文件将被删除或移动:
echo.
echo [旧的主入口文件]
echo   - main_exe.py
echo   - main_automation.py  
echo   - run_automation.py
echo.
echo [旧的功能文件]
echo   - rename_files.py
echo   - fill_evaluation_forms_strict.py
echo   - docx_to_pdf_converter.py
echo.
echo [旧的配置文件]
echo   - config.py
echo.
echo [旧的构建文件]
echo   - build_exe.spec
echo   - 打包EXE文件.bat
echo.
echo [临时目录]
echo   - __pycache__/
echo   - build/
echo.

set /p confirm="确认删除这些文件? (y/n): "
if /i not "%confirm%"=="y" (
    echo 取消清理操作
    pause
    exit /b 0
)

echo.
echo 开始清理...

:: 删除旧的主入口文件
if exist "main_exe.py" (
    del "main_exe.py"
    echo ✓ 删除 main_exe.py
)
if exist "main_automation.py" (
    del "main_automation.py"
    echo ✓ 删除 main_automation.py
)
if exist "run_automation.py" (
    del "run_automation.py"
    echo ✓ 删除 run_automation.py
)

:: 删除旧的功能文件
if exist "rename_files.py" (
    del "rename_files.py"
    echo ✓ 删除 rename_files.py
)
if exist "fill_evaluation_forms_strict.py" (
    del "fill_evaluation_forms_strict.py"
    echo ✓ 删除 fill_evaluation_forms_strict.py
)
if exist "docx_to_pdf_converter.py" (
    del "docx_to_pdf_converter.py"
    echo ✓ 删除 docx_to_pdf_converter.py
)

:: 删除旧的配置文件
if exist "config.py" (
    del "config.py"
    echo ✓ 删除 config.py
)

:: 删除旧的构建文件
if exist "build_exe.spec" (
    del "build_exe.spec"
    echo ✓ 删除 build_exe.spec
)
if exist "打包EXE文件.bat" (
    del "打包EXE文件.bat"
    echo ✓ 删除 打包EXE文件.bat
)

:: 删除缓存目录
if exist "__pycache__" (
    rmdir /s /q "__pycache__"
    echo ✓ 删除 __pycache__ 目录
)

if exist "build" (
    rmdir /s /q "build"
    echo ✓ 删除 build 目录
)

echo.
echo =============================================
echo 清理完成!
echo =============================================
echo.
echo 现在的项目结构更加清晰:
echo   - src/           源代码目录
echo   - config/        配置文件目录  
echo   - data/          数据文件目录
echo   - output/        输出文件目录
echo   - build_tools/   构建工具目录
echo   - docs/          文档目录
echo   - scripts/       脚本目录
echo.
pause
