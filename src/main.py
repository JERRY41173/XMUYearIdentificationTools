# -*- coding: utf-8 -*-
"""
学年鉴定表自动化处理工具 - 主程序
整合所有功能模块，提供统一的用户界面
"""

import sys
import os
from pathlib import Path
from typing import Optional, List

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from src.core.file_renamer import FileRenamer
from src.core.evaluation_filler import EvaluationFiller
from src.core.pdf_converter import PDFConverter
from src.utils.config_handler import get_config
from src.utils.dependency_manager import check_dependencies


class AutomationApp:
    """学年鉴定表自动化处理应用程序主类"""
    
    def __init__(self):
        """初始化应用程序"""
        self.config = get_config()
        self.file_renamer = FileRenamer()
        self.evaluation_filler = EvaluationFiller()
        self.pdf_converter = PDFConverter()
        
        print("=" * 70)
        print("          学年鉴定表自动化处理工具 v2.0")
        print("=" * 70)
    
    def check_dependencies(self) -> bool:
        """检查依赖项"""
        print("\n检查系统依赖项...")
        if not check_dependencies():
            print("依赖项检查失败，程序无法正常运行。")
            return False
        return True
    
    def get_user_paths(self) -> tuple:
        """
        获取用户输入的路径或使用配置文件中的路径
        
        Returns:
            tuple: (excel_file, source_dir, output_dir)
        """
        valid_paths = self.config.get_valid_paths()
        auto_mode = self.config.is_auto_mode()
        
        # Excel文件路径
        excel_file = None
        if valid_paths['excel_file'] and auto_mode:
            excel_file = valid_paths['excel_file']
            print(f"使用配置的Excel文件: {excel_file}")
        else:
            while True:
                if valid_paths['excel_file']:
                    choice = input(f"使用默认的Excel文件路径 '{valid_paths['excel_file']}'? (y/n): ").lower()
                    if choice in ['y', 'yes', '是']:
                        excel_file = valid_paths['excel_file']
                        break
                
                excel_file = input("请输入Excel名单文件的完整路径: ").strip().replace('"', '')
                if os.path.exists(excel_file):
                    break
                else:
                    print("文件不存在，请重新输入。")
        
        # 源文件夹路径
        source_dir = None
        if valid_paths['source_dir'] and auto_mode:
            source_dir = valid_paths['source_dir']
            print(f"使用默认的源文件夹: {source_dir}")
        else:
            while True:
                if valid_paths['source_dir']:
                    choice = input(f"使用默认的源文件夹 '{valid_paths['source_dir']}'? (y/n): ").lower()
                    if choice in ['y', 'yes', '是']:
                        source_dir = valid_paths['source_dir']
                        break
                
                source_dir = input("请输入存放学年鉴定word文档的文件夹路径: ").strip().replace('"', '')
                if os.path.exists(source_dir):
                    break
                else:
                    print("文件夹不存在，请重新输入。")
        
        # 输出文件夹路径
        output_dir = valid_paths['output_dir']
        if not auto_mode:
            new_output = input(f"填写输出文件夹文件路径 (直接回车则默认输出在本程序所在目录{output_dir}): ").strip().replace('"', '')
            if new_output:
                output_dir = new_output
        
        return excel_file, source_dir, output_dir
    
    def select_classes(self, available_classes: List[str]) -> List[str]:
        """
        让用户选择要处理的班级
        
        Args:
            available_classes: 可用的班级列表
            
        Returns:
            List[str]: 选择的班级列表
        """
        if self.config.is_auto_process_all_classes():
            print("自动模式: 处理所有班级")
            return available_classes
        
        print(f"\n找到 {len(available_classes)} 个班级:")
        for i, class_name in enumerate(available_classes, 1):
            print(f"  {i}. {class_name}")
        
        while True:
            choice = input("\n请选择处理方式:\n1. 处理所有班级\n2. 选择特定班级\n请输入选择 (1/2): ").strip()
            
            if choice == '1':
                return available_classes
            elif choice == '2':
                selected_classes = []
                print("请输入要处理的班级编号 (用逗号分隔，如: 1,3,5):")
                try:
                    indices = input().strip().split(',')
                    for idx in indices:
                        idx = int(idx.strip()) - 1
                        if 0 <= idx < len(available_classes):
                            selected_classes.append(available_classes[idx])
                        else:
                            print(f"警告: 编号 {idx + 1} 超出范围")
                    
                    if selected_classes:
                        return selected_classes
                    else:
                        print("未选择任何班级，请重新选择。")
                        continue
                        
                except ValueError:
                    print("输入格式错误，请重新输入。")
                    continue
            else:
                print("无效选择，请输入 1 或 2。")
    
    def process_workflow(self):
        """执行完整的处理工作流程"""
        try:
            # 1. 获取路径
            print("\n步骤 1: 获取文件路径")
            excel_file, source_dir, output_dir = self.get_user_paths()
            
            # 2. 文件重命名
            print("\n步骤 2: 文件重命名")
            print("正在读取Excel文件并重命名文件...")
            
            success, available_classes = self.file_renamer.process_files(
                excel_file, source_dir, output_dir
            )
            
            if not success or not available_classes:
                print("文件重命名失败，程序终止。")
                return False
            
            # 3. 选择要处理的班级
            print("\n步骤 3: 选择填写评语的班级")
            selected_classes = self.select_classes(available_classes)
            
            if not selected_classes:
                print("未选择任何班级，程序终止。")
                return False
            
            # 4. 填写评语
            print("\n步骤 4: 填写评语")
            for class_name in selected_classes:
                print(f"\n正在处理班级: {class_name}")
                class_dir = os.path.join(output_dir, class_name)
                
                if os.path.exists(class_dir):
                    self.evaluation_filler.process_class_files(class_dir)
                else:
                    print(f"警告: 班级文件夹不存在 - {class_dir}")
            
            # 5. PDF转换
            if self.config.is_pdf_conversion_enabled():
                print("\n步骤 5: PDF转换")
                choice = input("是否要转换为PDF格式? (y/n): ").lower()
                
                if choice in ['y', 'yes', '是']:
                    for class_name in selected_classes:
                        class_dir = os.path.join(output_dir, class_name)
                        pdf_dir = class_dir + "_PDF"
                        
                        if os.path.exists(class_dir):
                            print(f"\n转换班级 {class_name} 的文件...")
                            success_count, error_count, failed_files = self.pdf_converter.convert_batch(
                                class_dir, pdf_dir
                            )
                            
                            if failed_files:
                                print(f"班级 {class_name} 中有 {error_count} 个文件转换失败")
                        else:
                            print(f"跳过不存在的班级文件夹: {class_dir}")
            
            print("\n" + "=" * 70)
            print("                    处理完成!")
            print("=" * 70)
            return True
            
        except KeyboardInterrupt:
            print("\n\n用户中断程序执行。")
            return False
        except Exception as e:
            print(f"\n处理过程中出错: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            # 清理资源
            self.pdf_converter.cleanup()
    
    def run(self):
        """运行应用程序"""
        try:
            # 检查依赖项
            # if not self.check_dependencies():
            #     input("\n按回车键退出...")
            #     return
            
            print("\n程序已启动，请按照提示操作。\n")
            
            # 执行主要工作流程
            success = self.process_workflow()
            
            if success:
                print("\n所有操作已完成！")
            else:
                print("\n程序执行过程中遇到问题。")
            
        except Exception as e:
            print(f"\n程序运行时出现错误: {str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            input("\n按回车键退出...")


def main():
    """主函数"""
    app = AutomationApp()
    app.run()


if __name__ == "__main__":
    main()
