# -*- coding: utf-8 -*-
"""
核心模块包
包含主要的业务逻辑类
"""

from .file_renamer import FileRenamer
from .evaluation_filler import EvaluationFiller
from .pdf_converter import PDFConverter

__version__ = "2.0.0"
__author__ = "学年鉴定表自动化处理工具"

__all__ = [
    'FileRenamer',
    'EvaluationFiller', 
    'PDFConverter'
]
