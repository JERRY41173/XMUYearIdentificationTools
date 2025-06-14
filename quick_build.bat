@echo off
chcp 65001 > nul
echo 快速打包学年鉴定表自动化处理工具...
echo.

cd /d "%~dp0"

echo 安装PyInstaller...
pip install pyinstaller

echo 开始打包...
pyinstaller --onefile --console --name="学年鉴定表自动化处理工具" --icon="docs/zhu.ico" --add-data="config;config" --add-data="data;data" --hidden-import=pandas --hidden-import=openpyxl --hidden-import=docx --hidden-import=comtypes --hidden-import=colorama "src/main.py"

echo 打包完成！可执行文件位于 dist 文件夹中。
pause
