# -*- coding: utf-8 -*-
"""
依赖管理器模块
负责检查和安装项目所需的依赖项
"""

import sys
import subprocess
import importlib
from typing import List, Dict, Tuple, Optional


class DependencyManager:
    """依赖管理器类"""
    
    def __init__(self):
        """初始化依赖管理器"""
        self.required_packages = {
            'pandas': {
                'name': 'pandas',
                'description': 'Excel文件读取和数据处理',
                'install_name': 'pandas',
                'import_name': 'pandas'
            },
            'openpyxl': {
                'name': 'openpyxl',
                'description': 'Excel文件操作库',
                'install_name': 'openpyxl',
                'import_name': 'openpyxl'
            },
            'python-docx': {
                'name': 'python-docx',
                'description': 'Word文档操作库',
                'install_name': 'python-docx',
                'import_name': 'docx'
            },
            'comtypes': {
                'name': 'comtypes',
                'description': 'COM接口库（用于WPS/Office自动化）',
                'install_name': 'comtypes',
                'import_name': 'comtypes'
            }
        }
        
        self.optional_packages = {
            'colorama': {
                'name': 'colorama',
                'description': '彩色终端输出（可选）',
                'install_name': 'colorama',
                'import_name': 'colorama'
            }
        }
    
    def check_package(self, package_info: dict) -> bool:
        """
        检查单个包是否已安装
        
        Args:
            package_info: 包信息字典
            
        Returns:
            bool: 是否已安装
        """
        try:
            importlib.import_module(package_info['import_name'])
            return True
        except ImportError:
            return False
    
    def check_all_dependencies(self) -> Tuple[List[str], List[str]]:
        """
        检查所有依赖项
        
        Returns:
            Tuple[List[str], List[str]]: (已安装的包列表, 缺失的包列表)
        """
        installed = []
        missing = []
        
        print("正在检查依赖项...")
        print("=" * 50)
        
        # 检查必需的包
        for package_key, package_info in self.required_packages.items():
            if self.check_package(package_info):
                installed.append(package_key)
                print(f"✓ {package_info['name']} - 已安装")
            else:
                missing.append(package_key)
                print(f"✗ {package_info['name']} - 缺失")
        
        # 检查可选的包
        print("\n可选依赖项:")
        for package_key, package_info in self.optional_packages.items():
            if self.check_package(package_info):
                print(f"✓ {package_info['name']} - 已安装")
            else:
                print(f"- {package_info['name']} - 未安装（可选）")
        
        print("=" * 50)
        
        return installed, missing
    
    def install_package(self, package_key: str) -> bool:
        """
        安装单个包
        
        Args:
            package_key: 包的键名
            
        Returns:
            bool: 是否安装成功
        """
        if package_key not in self.required_packages:
            print(f"错误: 未知的包 '{package_key}'")
            return False
        
        package_info = self.required_packages[package_key]
        install_name = package_info['install_name']
        
        print(f"正在安装 {package_info['name']}...")
        
        try:
            # 使用pip安装包
            result = subprocess.run([
                sys.executable, '-m', 'pip', 'install', install_name
            ], capture_output=True, text=True, timeout=300)
            
            if result.returncode == 0:
                print(f"✓ {package_info['name']} 安装成功")
                return True
            else:
                print(f"✗ {package_info['name']} 安装失败:")
                print(result.stderr)
                return False
                
        except subprocess.TimeoutExpired:
            print(f"✗ {package_info['name']} 安装超时")
            return False
        except Exception as e:
            print(f"✗ {package_info['name']} 安装时出错: {str(e)}")
            return False
    
    def install_missing_dependencies(self, missing_packages: List[str]) -> bool:
        """
        安装所有缺失的依赖项
        
        Args:
            missing_packages: 缺失的包列表
            
        Returns:
            bool: 是否全部安装成功
        """
        if not missing_packages:
            print("所有必需的依赖项都已安装!")
            return True
        
        print(f"\n需要安装 {len(missing_packages)} 个缺失的依赖项:")
        for package_key in missing_packages:
            package_info = self.required_packages[package_key]
            print(f"  - {package_info['name']}: {package_info['description']}")
        
        # 询问用户是否要安装
        while True:
            choice = input("\n是否要自动安装这些依赖项? (y/n): ").lower().strip()
            if choice in ['y', 'yes', '是', '确定']:
                break
            elif choice in ['n', 'no', '否', '取消']:
                print("取消安装。请手动安装依赖项后再运行程序。")
                return False
            else:
                print("请输入 y 或 n")
        
        # 批量安装
        success_count = 0
        for package_key in missing_packages:
            if self.install_package(package_key):
                success_count += 1
        
        if success_count == len(missing_packages):
            print(f"\n✓ 所有依赖项安装完成!")
            return True
        else:
            print(f"\n⚠ 部分依赖项安装失败 ({success_count}/{len(missing_packages)})")
            return False
    
    def generate_requirements_txt(self, output_path: str = "requirements.txt"):
        """
        生成requirements.txt文件
        
        Args:
            output_path: 输出文件路径
        """
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write("# 学年鉴定表自动化处理工具依赖项\n")
                f.write("# 安装命令: pip install -r requirements.txt\n\n")
                
                f.write("# 必需依赖项\n")
                for package_info in self.required_packages.values():
                    f.write(f"{package_info['install_name']}\n")
                
                f.write("\n# 可选依赖项\n")
                for package_info in self.optional_packages.values():
                    f.write(f"# {package_info['install_name']}  # {package_info['description']}\n")
            
            print(f"✓ requirements.txt 已生成: {output_path}")
            
        except Exception as e:
            print(f"✗ 生成 requirements.txt 失败: {str(e)}")
    
    def get_system_info(self) -> Dict[str, str]:
        """
        获取系统信息
        
        Returns:
            Dict[str, str]: 系统信息字典
        """
        import platform
        
        return {
            'Python版本': sys.version,
            '平台': platform.platform(),
            '架构': platform.architecture()[0],
            'pip版本': self._get_pip_version()
        }
    
    def _get_pip_version(self) -> str:
        """获取pip版本"""
        try:
            result = subprocess.run([
                sys.executable, '-m', 'pip', '--version'
            ], capture_output=True, text=True, timeout=10)
            
            if result.returncode == 0:
                return result.stdout.strip()
            else:
                return "未知"
        except:
            return "未知"
    
    def print_system_info(self):
        """打印系统信息"""
        print("系统信息:")
        print("=" * 30)
        
        info = self.get_system_info()
        for key, value in info.items():
            print(f"{key}: {value}")
        
        print("=" * 30)
    
    def run_dependency_check(self) -> bool:
        """
        运行完整的依赖项检查流程
        
        Returns:
            bool: 是否所有依赖项都可用
        """
        print("学年鉴定表自动化处理工具 - 依赖项检查")
        print("=" * 60)
        
        # 打印系统信息
        self.print_system_info()
        print()
        
        # 检查依赖项
        installed, missing = self.check_all_dependencies()
        
        if not missing:
            print("✓ 所有必需的依赖项都已安装，可以正常运行程序!")
            return True
        else:
            return self.install_missing_dependencies(missing)


def check_dependencies() -> bool:
    """
    便捷函数：检查并安装依赖项
    
    Returns:
        bool: 是否所有依赖项都可用
    """
    manager = DependencyManager()
    return manager.run_dependency_check()


if __name__ == "__main__":
    # 直接运行依赖项检查
    check_dependencies()
