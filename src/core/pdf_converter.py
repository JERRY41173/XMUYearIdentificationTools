# -*- coding: utf-8 -*-
"""
PDF转换模块
使用WPS Office的COM接口进行Word文档到PDF的批量转换
"""

import os
import sys
import time
from pathlib import Path
from typing import List, Tuple, Optional
import traceback


class PDFConverter:
    """PDF转换器类，负责Word文档到PDF的转换"""
    
    def __init__(self):
        """初始化PDF转换器"""
        self.wps_app = None
        self.is_initialized = False
    
    def initialize_wps(self) -> bool:
        """
        初始化WPS Office应用程序
        
        Returns:
            bool: 是否成功初始化
        """
        if self.is_initialized:
            return True
            
        try:
            import comtypes.client
            
            print("正在启动WPS Writer...")
            
            # 尝试不同的WPS应用程序标识
            wps_identifiers = ['Kwps.Application', 'wps.Application']
            
            for identifier in wps_identifiers:
                try:
                    self.wps_app = comtypes.client.CreateObject(identifier)
                    break
                except Exception:
                    continue
            
            if self.wps_app is None:
                raise Exception("无法找到WPS Office应用程序")
            
            self.wps_app.Visible = False
            self.is_initialized = True
            
            print("✓ WPS Writer启动成功")
            return True
            
        except ImportError:
            print("错误: 缺少comtypes库")
            print("请运行: pip install comtypes")
            return False
        except Exception as e:
            print(f"错误: 无法启动WPS Office - {str(e)}")
            print("请确保已安装WPS Office")
            return False
    
    def convert_single_file(self, input_path: str, output_path: str) -> bool:
        """
        转换单个文件
        
        Args:
            input_path: 输入文件路径
            output_path: 输出文件路径
            
        Returns:
            bool: 是否转换成功
        """
        if not self.is_initialized:
            if not self.initialize_wps():
                return False
        
        doc = None
        try:
            # 打开文档
            doc = self.wps_app.Documents.Open(os.path.abspath(input_path))
            time.sleep(0.5)  # 等待文档完全加载
            
            # 导出为PDF
            doc.ExportAsFixedFormat(
                OutputFileName=os.path.abspath(output_path),
                ExportFormat=17,  # PDF格式
                OpenAfterExport=False,
                OptimizeFor=0,
                BitmapMissingFonts=True,
                DocStructureTags=True,
                CreateBookmarks=False
            )
            
            # 关闭文档
            doc.Close(SaveChanges=False)
            time.sleep(0.2)  # 等待文档关闭
            
            # 验证输出文件是否存在
            if os.path.exists(output_path):
                return True
            else:
                print(f"警告: 输出文件未生成 - {output_path}")
                return False
                
        except Exception as e:
            print(f"转换文件时出错: {str(e)}")
            try:
                if doc:
                    doc.Close(SaveChanges=False)
            except:
                pass
            return False
    
    def convert_batch(self, source_folder: str, output_folder: str, 
                     file_extension: str = '.docx') -> Tuple[int, int, List[str]]:
        """
        批量转换文件
        
        Args:
            source_folder: 源文件夹路径
            output_folder: 输出文件夹路径
            file_extension: 要转换的文件扩展名
            
        Returns:
            Tuple[int, int, List[str]]: (成功数量, 失败数量, 失败文件列表)
        """
        if not os.path.exists(source_folder):
            raise FileNotFoundError(f"源文件夹不存在: {source_folder}")
        
        # 创建输出文件夹
        os.makedirs(output_folder, exist_ok=True)
        print(f"输出文件夹: {output_folder}")
        
        # 获取所有待转换文件
        files_to_convert = []
        for file in os.listdir(source_folder):
            if file.endswith(file_extension) and not file.startswith('~'):
                files_to_convert.append(file)
        
        if not files_to_convert:
            print(f"未找到任何{file_extension}文件")
            return 0, 0, []
        
        print(f"找到 {len(files_to_convert)} 个{file_extension}文件")
        
        # 初始化WPS
        if not self.initialize_wps():
            return 0, len(files_to_convert), files_to_convert
        
        print("开始批量转换...")
        print("=" * 60)
        
        success_count = 0
        error_count = 0
        failed_files = []
        
        for i, filename in enumerate(files_to_convert, 1):
            try:
                input_path = os.path.join(source_folder, filename)
                output_filename = filename.replace(file_extension, '.pdf')
                output_path = os.path.join(output_folder, output_filename)
                
                print(f"[{i}/{len(files_to_convert)}] 正在转换: {filename}")
                
                if self.convert_single_file(input_path, output_path):
                    success_count += 1
                    print(f"✓ 转换成功: {output_filename}")
                else:
                    error_count += 1
                    failed_files.append(filename)
                    print(f"✗ 转换失败: {filename}")
                
            except Exception as e:
                error_count += 1
                failed_files.append(filename)
                print(f"✗ 转换失败: {filename}")
                print(f"  错误详情: {str(e)}")
        
        print("=" * 60)
        print(f"转换完成! 成功: {success_count}, 失败: {error_count}")
        
        return success_count, error_count, failed_files
    
    def cleanup(self):
        """清理资源"""
        try:
            if self.wps_app:
                self.wps_app.Quit()
                self.wps_app = None
            self.is_initialized = False
            print("✓ WPS应用程序已关闭")
        except Exception as e:
            print(f"清理WPS应用程序时出错: {str(e)}")
    
    def __del__(self):
        """析构函数，确保资源被清理"""
        self.cleanup()


def convert_folder_to_pdf(source_folder: str, output_folder: str = None) -> bool:
    """
    便捷函数：将文件夹中的所有Word文档转换为PDF
    
    Args:
        source_folder: 源文件夹路径
        output_folder: 输出文件夹路径，如果为None则自动生成
        
    Returns:
        bool: 是否有文件转换成功
    """
    if output_folder is None:
        output_folder = source_folder + "_PDF"
    
    converter = PDFConverter()
    try:
        success_count, error_count, failed_files = converter.convert_batch(
            source_folder, output_folder
        )
        
        if failed_files:
            print(f"\n失败的文件列表:")
            for filename in failed_files:
                print(f"  - {filename}")
        
        return success_count > 0
        
    finally:
        converter.cleanup()


if __name__ == "__main__":
    # 测试用例
    if len(sys.argv) >= 2:
        source_folder = sys.argv[1]
        output_folder = sys.argv[2] if len(sys.argv) >= 3 else None
        convert_folder_to_pdf(source_folder, output_folder)
    else:
        print("用法: python pdf_converter.py <源文件夹> [输出文件夹]")
