# -*- coding: utf-8 -*-
"""
项目测试脚本
用于验证重构后的代码是否正常工作
"""

import sys
import os
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from src.core.file_renamer import FileRenamer
from src.core.evaluation_filler import EvaluationFiller
from src.core.pdf_converter import PDFConverter
from src.utils.config_handler import get_config
from src.utils.dependency_manager import check_dependencies


def test_config():
    """测试配置系统"""
    print("=" * 50)
    print("测试配置系统")
    print("=" * 50)
    
    config = get_config()
    print("当前配置:")
    print(config)
    
    print("\n路径验证:")
    validation = config.validate_paths()
    for path_name, is_valid in validation.items():
        status = "✓" if is_valid else "✗"
        print(f"  {status} {path_name}: {config.get(f'paths.{path_name}')}")


def test_file_renamer():
    """测试文件重命名功能"""
    print("\n" + "=" * 50)
    print("测试文件重命名功能")
    print("=" * 50)
    
    config = get_config()
    renamer = FileRenamer()
    
    excel_file = config.get('paths.excel_file')
    source_dir = config.get('paths.source_dir') 
    output_dir = config.get('paths.output_dir')
    
    # 检查文件是否存在
    if not os.path.exists(excel_file):
        print(f"❌ Excel文件不存在: {excel_file}")
        return False
    
    if not os.path.exists(source_dir):
        print(f"❌ 源目录不存在: {source_dir}")
        return False
    
    print(f"Excel文件: {excel_file}")
    print(f"源目录: {source_dir}")
    print(f"输出目录: {output_dir}")
    
    # 测试Excel数据读取
    student_data = renamer.load_excel_data(excel_file)
    if student_data:
        print(f"✓ 成功读取 {len(student_data)} 条学生记录")
        
        # 显示前几条记录
        for i, student in enumerate(student_data[:3]):
            print(f"  学生 {i+1}: {student}")
    else:
        print("❌ Excel数据读取失败")
        return False
    
    return True


def test_evaluation_filler():
    """测试评语填写功能"""
    print("\n" + "=" * 50)
    print("测试评语填写功能")
    print("=" * 50)
    
    filler = EvaluationFiller()
    print("✓ EvaluationFiller 初始化成功")
    
    # 测试评语生成
    opinion = filler.get_random_academic_year_opinion(0)
    print(f"示例学年意见: {opinion[:50]}...")
    
    evaluation = filler.get_random_class_organization_evaluation()
    print(f"示例班团鉴定: {evaluation[:50]}...")
    
    return True


def test_pdf_converter():
    """测试PDF转换功能"""
    print("\n" + "=" * 50)
    print("测试PDF转换功能")
    print("=" * 50)
    
    converter = PDFConverter()
    print("✓ PDFConverter 初始化成功")
    
    # 尝试初始化WPS (不会实际启动)
    try:
        # 只是测试导入
        import comtypes.client
        print("✓ comtypes 库可用")
    except ImportError:
        print("❌ comtypes 库不可用")
        return False
    
    converter.cleanup()
    return True


def run_all_tests():
    """运行所有测试"""
    print("学年鉴定表自动化处理工具 - 系统测试")
    print("=" * 70)
    
    # 1. 依赖检查
    print("1. 检查依赖项...")
    if not check_dependencies():
        print("❌ 依赖项检查失败")
        return False
    
    # 2. 配置测试
    test_config()
    
    # 3. 功能模块测试
    tests = [
        ("文件重命名模块", test_file_renamer),
        ("评语填写模块", test_evaluation_filler),
        ("PDF转换模块", test_pdf_converter)
    ]
    
    for test_name, test_func in tests:
        try:
            if test_func():
                print(f"✓ {test_name} 测试通过")
            else:
                print(f"❌ {test_name} 测试失败")
                return False
        except Exception as e:
            print(f"❌ {test_name} 测试出错: {str(e)}")
            return False
    
    print("\n" + "=" * 70)
    print("🎉 所有测试通过！系统重构成功！")
    print("=" * 70)
    
    return True


if __name__ == "__main__":
    success = run_all_tests()
    if not success:
        print("\n❌ 测试失败，请检查相关问题")
        sys.exit(1)
    else:
        print("\n✅ 系统测试完成，可以正常使用")
        sys.exit(0)
