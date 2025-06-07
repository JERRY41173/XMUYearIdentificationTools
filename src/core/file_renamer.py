# -*- coding: utf-8 -*-
"""
文件重命名核心模块
功能：从Excel文件中读取学号和姓名，将学年鉴定表文件重命名并按班级分类
"""

import pandas as pd
import os
import shutil
from pathlib import Path
from typing import List, Dict, Tuple


class FileRenamer:
    """文件重命名器"""
    
    def __init__(self):
        self.possible_id_columns = ['学号', '学生编号', 'ID', 'id', '编号']
        self.possible_name_columns = ['姓名', '名字', 'Name', 'name', '学生姓名']
    
    def get_class_list(self, excel_file):
        """获取Excel文件中的所有工作表（班级）名称"""
        try:
            xl_file = pd.ExcelFile(excel_file)
            return xl_file.sheet_names
        except Exception as e:
            raise Exception(f"读取Excel文件失败: {e}")
    
    def read_student_data(self, excel_file, class_name):
        """从指定班级的工作表中读取学号和姓名"""
        try:
            df = pd.read_excel(excel_file, sheet_name=class_name)
            
            # 查找学号和姓名列
            id_column = self._find_column(df, self.possible_id_columns)
            name_column = self._find_column(df, self.possible_name_columns)
            
            if not id_column or not name_column:
                raise Exception(f"在班级 '{class_name}' 中找不到学号或姓名列。可用列名: {list(df.columns)}")
            
            # 创建学号到姓名的映射
            student_dict = {}
            for _, row in df.iterrows():
                student_id = self._process_student_id(row[id_column])
                student_name = str(row[name_column]).strip()
                
                if student_id and student_name and student_id != 'nan' and student_name != 'nan':
                    student_dict[student_id] = student_name
            
            return student_dict
            
        except Exception as e:
            raise Exception(f"读取班级 '{class_name}' 数据失败: {e}")
    
    def rename_and_organize_files(self, student_dict, source_dir, output_dir, class_name):
        """重命名文件并按班级组织"""
        
        # 创建班级输出目录
        class_output_dir = os.path.join(output_dir, class_name)
        os.makedirs(class_output_dir, exist_ok=True)
        
        renamed_count = 0
        not_found_count = 0
        
        # 遍历源目录中的所有文件
        for filename in os.listdir(source_dir):
            if filename.endswith('.docx'):
                # 提取学号（去掉扩展名）
                student_id = filename.replace('.docx', '')
                
                if student_id in student_dict:
                    student_name = student_dict[student_id]
                    # 新文件名格式：姓名-学号.docx
                    new_filename = f"{student_name}-{student_id}.docx"
                    
                    source_path = os.path.join(source_dir, filename)
                    target_path = os.path.join(class_output_dir, new_filename)
                    
                    try:
                        # 复制并重命名文件
                        shutil.copy2(source_path, target_path)
                        renamed_count += 1
                    except Exception as e:
                        not_found_count += 1
                        raise Exception(f"重命名文件 {filename} 失败: {e}")
                else:
                    pass
        return renamed_count, not_found_count
    
    def _find_column(self, df, possible_columns):
        """查找匹配的列名"""
        for col in possible_columns:
            if col in df.columns:
                return col
        return None
    def _process_student_id(self, student_id_raw):
        """处理学号：统一转换为字符串格式"""
        if pd.isna(student_id_raw):
            return None
        if isinstance(student_id_raw, (int, float)):
            return str(int(student_id_raw)).strip()
        else:
            return str(student_id_raw).strip()
    
    def load_excel_data(self, excel_file: str) -> List[Dict]:
        """
        从Excel文件中加载所有学生数据
        
        Args:
            excel_file: Excel文件路径
            
        Returns:
            List[Dict]: 学生数据列表，每个字典包含学号、姓名、班级等信息
        """
        try:
            student_data = []
            
            # 获取所有工作表（班级）
            class_names = self.get_class_list(excel_file)
            
            for class_name in class_names:
                try:
                    # 读取每个班级的数据
                    student_dict = self.read_student_data(excel_file, class_name)
                    
                    # 将字典转换为统一格式的列表
                    for student_id, student_name in student_dict.items():
                        student_info = {
                            '学号': student_id,
                            '姓名': student_name,
                            '班级': class_name
                        }
                        student_data.append(student_info)
                        
                except Exception as e:
                    print(f"警告：读取班级 '{class_name}' 时出错: {e}")
                    continue
            return student_data
            
        except Exception as e:
            print(f"❌ 加载Excel数据失败: {e}")
            return []
    
    def rename_files_for_class(self, class_students: List[Dict], source_dir: str, output_dir: str, class_name: str) -> Tuple[int, int]:
        """
        为特定班级重命名文件
        
        Args:
            class_students: 班级学生数据列表
            source_dir: 源文件目录
            output_dir: 输出目录
            class_name: 班级名称
            
        Returns:
            Tuple[int, int]: (成功数量, 失败数量)
        """
        try:
            # 创建学号到姓名的映射字典
            student_dict = {}
            for student in class_students:
                student_dict[student['学号']] = student['姓名']
            
            # 使用现有的重命名方法
            return self.rename_and_organize_files(student_dict, source_dir, output_dir, class_name)
            
        except Exception as e:
            print(f"❌ 处理班级 '{class_name}' 时出错: {e}")
            return 0, len(class_students)
    
    def process_files(self, excel_file: str, source_dir: str, output_dir: str) -> Tuple[bool, List[str]]:
        """
        处理文件重命名的主要方法
        
        Args:
            excel_file: Excel名单文件路径
            source_dir: 源文件目录
            output_dir: 输出目录
            
        Returns:
            Tuple[bool, List[str]]: (是否成功, 可用班级列表)
        """
        try:
            # 1. 加载Excel数据
            print("正在读取Excel文件...")
            student_data = self.load_excel_data(excel_file)
            if not student_data:
                print("❌ Excel文件读取失败或为空")
                return False, []
            
            print(f"✓ 成功读取 {len(student_data)} 条学生记录")
            
            # 2. 获取可用班级
            available_classes = self.get_available_classes(student_data)
            if not available_classes:
                print("❌ 未找到任何班级信息")
                return False, []
            
            print(f"✓ 发现 {len(available_classes)} 个班级: {', '.join(available_classes)}")
            
            # 3. 创建输出目录
            os.makedirs(output_dir, exist_ok=True)
            
            # 4. 重命名文件
            print("\n开始文件重命名...")
            total_success = 0
            total_failed = 0
            
            for class_name in available_classes:
                class_students = [s for s in student_data if s['班级'] == class_name]
                success_count, failed_count = self.rename_files_for_class(
                    class_students, source_dir, output_dir, class_name
                )
                total_success += success_count
                total_failed += failed_count
            
            print(f"\n文件重命名完成!")
            print(f"成功: {total_success} 个文件")
            print(f"失败: {total_failed} 个文件")
            
            return total_success > 0, available_classes
            
        except Exception as e:
            print(f"❌ 处理过程中出错: {str(e)}")
            import traceback
            traceback.print_exc()
            return False, []
    
    def get_available_classes(self, student_data: List[Dict]) -> List[str]:
        """
        获取所有可用的班级列表
        
        Args:
            student_data: 学生数据列表
            
        Returns:
            List[str]: 班级名称列表
        """
        classes = set()
        for student in student_data:
            if '班级' in student and student['班级']:
                classes.add(str(student['班级']).strip())
        
        return sorted(list(classes))
