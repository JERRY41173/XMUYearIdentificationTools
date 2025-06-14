@echo off
setlocal enabledelayedexpansion
chcp 65001 > nul
echo ===============================================
echo     学年鉴定表自动化处理工具 - 打包脚本
echo ===============================================
echo.

set PROJECT_ROOT=%~dp0
cd /d "%PROJECT_ROOT%"

echo [1/5] 检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到Python！请确保Python已安装并添加到PATH。
    pause
    exit /b 1
)
python --version

echo [2/5] 安装依赖包...
pip install -r requirements.txt
if errorlevel 1 (
    echo 错误：依赖包安装失败！
    pause
    exit /b 1
)

echo [3/5] 清理之前的构建...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "release" rmdir /s /q "release"

echo [4/5] 开始打包可执行文件...
echo 这可能需要几分钟时间，请耐心等待...

REM 使用spec文件打包
if exist "student_evaluation.spec" (
    echo 使用spec文件打包...
    pyinstaller student_evaluation.spec
) else (
    echo 使用直接命令打包...
    pyinstaller --onefile --console --name="学年鉴定表自动化处理工具" --icon="docs/zhu.ico" --add-data="config;config" --add-data="data;data" --hidden-import=pandas --hidden-import=openpyxl --hidden-import=docx --hidden-import=comtypes --hidden-import=colorama "src/main.py"
)

if errorlevel 1 (
    echo 错误：打包失败！
    pause
    exit /b 1
)

echo [5/5] 创建发布包...
mkdir "release" 2>nul

REM 复制可执行文件
if exist "dist\学年鉴定表自动化处理工具.exe" (
    copy "dist\学年鉴定表自动化处理工具.exe" "release\" >nul 2>&1
    echo 已复制可执行文件
) else (
    echo 错误：找不到生成的可执行文件！
    dir "dist\" /b
    pause
    exit /b 1
)

REM 复制配置和数据文件
if exist "config" (
    xcopy "config" "release\config\" /E /I /Y >nul 2>&1
    echo 已复制配置文件
)
if exist "data" (
    xcopy "data" "release\data\" /E /I /Y >nul 2>&1
    echo 已复制数据文件
)
if exist "README.md" copy "README.md" "release\" >nul 2>&1
if exist "LICENSE" copy "LICENSE" "release\" >nul 2>&1

echo 创建使用说明...
(
echo 学年鉴定表自动化处理工具 v2.0
echo ================================================
echo.
echo 使用方法：
echo 1. 双击运行"学年鉴定表自动化处理工具.exe"
echo 2. 按照程序提示输入Excel文件路径和源文件夹路径
echo 3. 程序将自动完成文件重命名、评语填写等操作
echo.
echo 注意事项：
echo - 确保已安装WPS Office或Microsoft Office
echo - 处理前请备份原始文件
echo - 如遇问题请参考README.md文档
echo.
echo 版本信息：
echo - 版本：v2.0
echo - 兼容系统：Windows 10/11
echo.
echo 技术支持：
echo 如有问题，请查看项目文档或联系开发者。
) > "release\使用说明.txt"

echo.
echo ===============================================
echo                  打包完成！
echo ===============================================
echo 发布包位置: %PROJECT_ROOT%release\
echo 可执行文件: 学年鉴定表自动化处理工具.exe
echo.

if exist "release\学年鉴定表自动化处理工具.exe" (
    for %%A in ("release\学年鉴定表自动化处理工具.exe") do (
        set /a size=%%~zA/1024/1024
        echo 文件大小: !size! MB
    )
)

echo.
echo 现在可以将release文件夹分发给用户使用！
echo 用户只需要双击exe文件即可运行程序。
echo.
pause
