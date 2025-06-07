# -*- coding: utf-8 -*-
"""
配置处理器模块
负责处理应用程序的配置加载、验证和管理
"""

import os
import json
from pathlib import Path
from typing import Dict, Any, Optional


class ConfigHandler:
    """配置处理器类"""
    
    def __init__(self, config_file: str = None):
        """
        初始化配置处理器
        
        Args:
            config_file: 配置文件路径，默认为config/settings.json
        """
        if config_file is None:
            config_file = os.path.join("config", "settings.json")
        
        self.config_file = config_file
        self.config = {}
        self.default_config = {
            "paths": {
                "excel_file": "./data/samples/名单样例.xlsx",
                "source_dir": "./data/templates/四年制学年鉴定表",
                "output_dir": "./output/重命名后的文件"
            },
            "automation": {
                "auto_mode": False,
                "auto_process_all_classes": False,
                "backup_original_files": True
            },
            "pdf_conversion": {
                "enabled": True,
                "wps_timeout": 30,
                "retry_count": 3
            },
            "file_operations": {
                "allowed_extensions": [".docx", ".doc"],
                "skip_temp_files": True,
                "preserve_timestamps": False
            }
        }
        
        self.load_config()
    
    def load_config(self):
        """加载配置文件"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    # 合并默认配置和加载的配置
                    self.config = self._merge_configs(self.default_config, loaded_config)
            else:
                # 如果配置文件不存在，使用默认配置并保存
                self.config = self.default_config.copy()
                self.save_config()
                
        except Exception as e:
            print(f"警告: 加载配置文件失败 - {str(e)}")
            print("使用默认配置")
            self.config = self.default_config.copy()
    
    def save_config(self):
        """保存配置到文件"""
        try:
            # 确保配置目录存在
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            print(f"警告: 保存配置文件失败 - {str(e)}")
    
    def get(self, key_path: str, default=None) -> Any:
        """
        获取配置值
        
        Args:
            key_path: 配置键路径，用点分隔，如 'paths.excel_file'
            default: 默认值
            
        Returns:
            配置值或默认值
        """
        keys = key_path.split('.')
        value = self.config
        
        try:
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            return default
    
    def set(self, key_path: str, value: Any):
        """
        设置配置值
        
        Args:
            key_path: 配置键路径，用点分隔
            value: 要设置的值
        """
        keys = key_path.split('.')
        config_ref = self.config
        
        # 导航到正确的嵌套位置
        for key in keys[:-1]:
            if key not in config_ref:
                config_ref[key] = {}
            config_ref = config_ref[key]
        
        # 设置最终值
        config_ref[keys[-1]] = value
    
    def validate_paths(self) -> Dict[str, bool]:
        """
        验证配置中的路径是否存在
        
        Returns:
            Dict[str, bool]: 路径名到是否存在的映射
        """
        path_validation = {}
        
        # 验证文件路径
        excel_file = self.get('paths.excel_file')
        if excel_file:
            path_validation['excel_file'] = os.path.exists(excel_file)
        
        # 验证目录路径
        source_dir = self.get('paths.source_dir')
        if source_dir:
            path_validation['source_dir'] = os.path.exists(source_dir)
        
        output_dir = self.get('paths.output_dir')
        if output_dir:
            # 输出目录不存在是可以的，会自动创建
            path_validation['output_dir'] = True
        
        return path_validation
    
    def get_valid_paths(self) -> Dict[str, Optional[str]]:
        """
        获取有效的路径配置
        
        Returns:
            Dict[str, Optional[str]]: 有效路径的映射
        """
        valid_paths = {}
        path_validation = self.validate_paths()
        
        excel_file = self.get('paths.excel_file')
        if excel_file and path_validation.get('excel_file', False):
            valid_paths['excel_file'] = excel_file
        else:
            valid_paths['excel_file'] = None
        
        source_dir = self.get('paths.source_dir')
        if source_dir and path_validation.get('source_dir', False):
            valid_paths['source_dir'] = source_dir
        else:
            valid_paths['source_dir'] = None
        
        valid_paths['output_dir'] = self.get('paths.output_dir')
        
        return valid_paths
    
    def is_auto_mode(self) -> bool:
        """检查是否启用自动模式"""
        return self.get('automation.auto_mode', False)
    
    def is_auto_process_all_classes(self) -> bool:
        """检查是否自动处理所有班级"""
        return self.get('automation.auto_process_all_classes', False)
    
    def should_backup_files(self) -> bool:
        """检查是否应该备份原始文件"""
        return self.get('automation.backup_original_files', True)
    
    def is_pdf_conversion_enabled(self) -> bool:
        """检查是否启用PDF转换"""
        return self.get('pdf_conversion.enabled', True)
    
    def get_allowed_extensions(self) -> list:
        """获取允许的文件扩展名列表"""
        return self.get('file_operations.allowed_extensions', ['.docx', '.doc'])
    
    def _merge_configs(self, default: dict, loaded: dict) -> dict:
        """
        递归合并配置字典
        
        Args:
            default: 默认配置
            loaded: 加载的配置
            
        Returns:
            合并后的配置
        """
        result = default.copy()
        
        for key, value in loaded.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = self._merge_configs(result[key], value)
            else:
                result[key] = value
        
        return result
    
    def reset_to_defaults(self):
        """重置配置为默认值"""
        self.config = self.default_config.copy()
        self.save_config()
    
    def __str__(self) -> str:
        """返回配置的字符串表示"""
        return json.dumps(self.config, ensure_ascii=False, indent=2)


# 全局配置实例
config_handler = ConfigHandler()


def get_config() -> ConfigHandler:
    """获取全局配置处理器实例"""
    return config_handler


if __name__ == "__main__":
    # 测试配置处理器
    config = ConfigHandler()
    print("当前配置:")
    print(config)
    
    print("\n路径验证结果:")
    validation = config.validate_paths()
    for path_name, is_valid in validation.items():
        status = "✓" if is_valid else "✗"
        print(f"  {status} {path_name}: {config.get(f'paths.{path_name}')}")
