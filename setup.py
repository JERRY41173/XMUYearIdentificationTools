#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
学年鉴定表自动化处理工具 - 安装脚本
"""

from setuptools import setup, find_packages
from pathlib import Path

# 读取README文件
README_PATH = Path(__file__).parent / "README.md"
try:
    with open(README_PATH, "r", encoding="utf-8") as f:
        long_description = f.read()
except FileNotFoundError:
    long_description = ""

# 读取版本信息
VERSION = "2.0.0"

setup(
    name="student-evaluation-automation",
    version=VERSION,
    description="学年鉴定表自动化处理工具",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="开发者",
    author_email="developer@example.com",
    url="https://github.com/yourusername/student-evaluation-automation",
    project_urls={
        "Bug Reports": "https://github.com/yourusername/student-evaluation-automation/issues",
        "Source": "https://github.com/yourusername/student-evaluation-automation",
    },
    license="MIT",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    include_package_data=True,
    package_data={
        "": ["config/*.json", "data/**/*"],
    },
    python_requires=">=3.8",
    install_requires=[
        "pandas>=1.5.0",
        "openpyxl>=3.0.10",
        "python-docx>=0.8.11",
        "comtypes>=1.1.14",
        "colorama>=0.4.6",
    ],
    extras_require={
        "dev": [
            "pyinstaller>=6.0.0",
            "pytest>=7.0.0",
            "black>=22.0.0",
            "flake8>=4.0.0",
            "mypy>=0.900",
        ],
    },
    entry_points={
        "console_scripts": [
            "student-evaluation=main:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Education",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
    ],
    keywords="automation, education, document processing, excel, word",
    zip_safe=False,
)
