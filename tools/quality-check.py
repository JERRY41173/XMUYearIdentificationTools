#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
代码质量检查脚本
"""

import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent.parent
SRC_DIR = PROJECT_ROOT / "src"
TESTS_DIR = PROJECT_ROOT / "tests"

def run_command(cmd, description):
    """运行命令并显示结果"""
    print(f"\n{'='*50}")
    print(f"正在执行: {description}")
    print(f"{'='*50}")
    
    try:
        result = subprocess.run(
            cmd,
            shell=True,
            cwd=PROJECT_ROOT,
            text=True,
            encoding='utf-8'
        )
        
        if result.returncode == 0:
            print(f"✅ {description} 通过")
            return True
        else:
            print(f"❌ {description} 失败")
            return False
            
    except Exception as e:
        print(f"❌ {description} 执行异常: {e}")
        return False

def main():
    """主函数"""
    print("学年鉴定表自动化处理工具 - 代码质量检查")
    print("="*60)
    
    checks = [
        ("python -m black --check src/ tests/", "代码格式检查"),
        ("python -m flake8 src/ tests/", "代码风格检查"),
        ("python -m mypy src/", "类型检查"),
        ("python -m pytest tests/ -v", "单元测试"),
    ]
    
    results = []
    for cmd, desc in checks:
        result = run_command(cmd, desc)
        results.append((desc, result))
    
    # 显示总结
    print(f"\n{'='*60}")
    print("检查结果总结:")
    print(f"{'='*60}")
    
    passed = 0
    for desc, result in results:
        status = "✅ 通过" if result else "❌ 失败"
        print(f"{desc}: {status}")
        if result:
            passed += 1
    
    print(f"\n总计: {passed}/{len(results)} 项检查通过")
    
    if passed == len(results):
        print("🎉 所有检查都通过！代码质量良好。")
        sys.exit(0)
    else:
        print("⚠️  有检查项未通过，请修复后重试。")
        sys.exit(1)

if __name__ == "__main__":
    main()
