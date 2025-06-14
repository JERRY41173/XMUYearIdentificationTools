#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
项目打包构建脚本
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

# 项目根目录
PROJECT_ROOT = Path(__file__).parent.parent
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
RELEASE_DIR = PROJECT_ROOT / "release"

def run_command(cmd, cwd=None):
    """运行命令并返回结果"""
    try:
        result = subprocess.run(
            cmd, 
            shell=True, 
            cwd=cwd or PROJECT_ROOT, 
            capture_output=True, 
            text=True,
            encoding='utf-8'
        )
        if result.returncode != 0:
            print(f"命令执行失败: {cmd}")
            print(f"错误输出: {result.stderr}")
            return False
        return True
    except Exception as e:
        print(f"命令执行异常: {e}")
        return False

def clean_build():
    """清理构建目录"""
    print("清理构建目录...")
    for dir_path in [DIST_DIR, BUILD_DIR, RELEASE_DIR]:
        if dir_path.exists():
            shutil.rmtree(dir_path)
            print(f"已清理: {dir_path}")

def check_environment():
    """检查构建环境"""
    print("检查构建环境...")
    
    # 检查虚拟环境
    venv_python = PROJECT_ROOT / "venv" / "Scripts" / "python.exe"
    if not venv_python.exists():
        print("错误：虚拟环境不存在！请先运行 tools/setup-dev.bat")
        return False
    
    # 检查PyInstaller
    result = subprocess.run(
        [str(venv_python), "-m", "pyinstaller", "--version"],
        capture_output=True,
        text=True
    )
    if result.returncode != 0:
        print("错误：PyInstaller未安装！请安装开发依赖")
        return False
    
    print("构建环境检查通过")
    return True

def build_executable():
    """构建可执行文件"""
    print("开始构建可执行文件...")
    
    venv_python = PROJECT_ROOT / "venv" / "Scripts" / "python.exe"
    
    # PyInstaller命令
    cmd = [
        str(venv_python), "-m", "pyinstaller",
        "--onefile",
        "--console",
        "--name=学年鉴定表自动化处理工具",
        f"--icon={PROJECT_ROOT}/docs/zhu.ico",
        f"--add-data={PROJECT_ROOT}/config;config",
        f"--add-data={PROJECT_ROOT}/data;data",
        "--hidden-import=pandas",
        "--hidden-import=openpyxl", 
        "--hidden-import=docx",
        "--hidden-import=comtypes",
        "--hidden-import=colorama",
        f"{PROJECT_ROOT}/src/main.py"
    ]
    
    result = subprocess.run(cmd, cwd=PROJECT_ROOT)
    
    if result.returncode == 0:
        print("可执行文件构建成功！")
        return True
    else:
        print("可执行文件构建失败！")
        return False

def create_release_package():
    """创建发布包"""
    print("创建发布包...")
    
    # 创建发布目录
    RELEASE_DIR.mkdir(exist_ok=True)
    
    # 复制可执行文件
    exe_file = DIST_DIR / "学年鉴定表自动化处理工具.exe"
    if exe_file.exists():
        shutil.copy2(exe_file, RELEASE_DIR)
        print(f"已复制: {exe_file.name}")
    else:
        print("错误：可执行文件不存在！")
        return False
    
    # 复制必要文件
    files_to_copy = [
        "README.md",
        "LICENSE", 
        "CHANGELOG.md",
        "config",
        "data"
    ]
    
    for item in files_to_copy:
        src = PROJECT_ROOT / item
        dst = RELEASE_DIR / item
        
        if src.is_file():
            shutil.copy2(src, dst)
            print(f"已复制文件: {item}")
        elif src.is_dir():
            if dst.exists():
                shutil.rmtree(dst)
            shutil.copytree(src, dst)
            print(f"已复制目录: {item}")
    
    # 创建使用说明
    usage_txt = RELEASE_DIR / "使用说明.txt"
    with open(usage_txt, 'w', encoding='utf-8') as f:
        f.write("""学年鉴定表自动化处理工具 v2.0
================================================

使用方法：
1. 双击运行"学年鉴定表自动化处理工具.exe"
2. 按照程序提示输入Excel文件路径和源文件夹路径
3. 程序将自动完成文件重命名、评语填写等操作

注意事项：
- 确保已安装WPS Office或Microsoft Office
- 处理前请备份原始文件
- 如遇问题请参考README.md文档

版本信息：
- 版本：v2.0
- 兼容系统：Windows 10/11

技术支持：
如有问题，请查看项目文档或联系开发者。
""")
    
    print(f"发布包创建完成：{RELEASE_DIR}")
    return True

def main():
    """主函数"""
    print("=" * 60)
    print("        学年鉴定表自动化处理工具 - 构建脚本")
    print("=" * 60)
    
    # 检查环境
    if not check_environment():
        sys.exit(1)
    
    # 清理构建目录
    clean_build()
    
    # 构建可执行文件
    if not build_executable():
        sys.exit(1)
    
    # 创建发布包
    if not create_release_package():
        sys.exit(1)
    
    print("\n" + "=" * 60)
    print("                构建完成！")
    print("=" * 60)
    print(f"发布包位置: {RELEASE_DIR}")
    
    # 显示文件大小
    exe_file = RELEASE_DIR / "学年鉴定表自动化处理工具.exe"
    if exe_file.exists():
        size_mb = exe_file.stat().st_size / (1024 * 1024)
        print(f"可执行文件大小: {size_mb:.1f} MB")
    
    print("\n现在可以将release文件夹分发给用户使用！")

if __name__ == "__main__":
    main()
