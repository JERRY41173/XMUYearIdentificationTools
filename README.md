# 学年鉴定表自动化处理工具 v2.0

> 🎓 专为厦门大学学年鉴定表批量处理设计的自动化工具

## ✨ 功能特性

### 核心功能

- **📝 文件重命名**：根据Excel名单自动重命名学年鉴定表文件,自动按班级分类组织文件
- **✍️ 自动评语填写**：批量填写标准化的学年意见和班团组织鉴定
- **📄 PDF批量转换**：将处理完的Word文档转换为PDF格式

### 技术特性

- **🔧 模块化设计**：清晰的代码结构，便于维护和扩展
- **⚙️ 配置化管理**：支持自定义路径和处理参数
- **🛡️ 错误处理**：完善的异常处理和用户友好的错误提示
- **📈 依赖检查**：自动检测和安装必要的依赖项

## 🚀 快速开始

从GitHub下载release版本：default.exe运行，根据程序提示执行即可

## 📁 项目结构

```bash
学年鉴定表自动化处理工具/
├── src/                          # 源代码目录
│   ├── core/                     # 核心功能模块
│   │   ├── __init__.py          # 模块初始化
│   │   ├── file_renamer.py      # 文件重命名模块
│   │   ├── evaluation_filler.py # 评语填写模块
│   │   └── pdf_converter.py     # PDF转换模块
│   ├── utils/                    # 工具模块
│   │   ├── __init__.py          # 工具模块初始化
│   │   ├── config_handler.py    # 配置管理器
│   │   └── dependency_manager.py # 依赖管理器
│   └── main.py                   # 主程序入口
├── tests/                        # 测试目录
│   ├── __init__.py              # 测试模块初始化
│   ├── test_config_handler.py   # 配置处理器测试
│   └── test_file_renamer.py     # 文件重命名器测试
├── tools/                        # 开发工具目录
│   ├── setup-dev.bat           # 开发环境设置脚本
│   ├── build.py                # 项目构建脚本
│   └── quality-check.py        # 代码质量检查脚本
├── config/                       # 配置文件目录
│   └── settings.json            # 主配置文件
├── data/                         # 数据目录
│   ├── samples/                 # 样例数据
│   └── templates/               # 模板文件
│       └── 四年制学年鉴定表/     # 学年鉴定表模板
├── docs/                         # 文档和资源目录
│   ├── zhu.ico                  # 应用程序图标
│   └── zhu.svg                  # 矢量图标
├── output/                       # 输出目录（运行时生成）
├── dist/                         # 打包输出目录（构建时生成）
├── release/                      # 发布包目录（构建时生成）
├── requirements.txt              # 运行时依赖列表
├── requirements-dev.txt          # 开发依赖列表
├── pyproject.toml               # 项目配置文件
├── setup.py                     # 安装配置脚本
├── MANIFEST.in                  # 打包清单文件
├── LICENSE                      # 开源许可证
├── CHANGELOG.md                 # 更新日志
├── .gitignore                   # Git忽略文件列表
├── run.bat                      # 快速启动脚本
└── README.md                    # 项目说明文档
```

## 🚀 从源代码开始

### 系统要求

- Windows 10/11
- Python 3.7+  
- WPS Office

### 安装步骤

#### 方法1：使用自动化脚本（推荐）

1. **下载项目**

   ```bash
   git clone <项目地址>
   cd 学年鉴定
   ```

2. **一键设置环境**
   - 双击运行 `activate_env.bat` 激活虚拟环境
   - 双击运行 `install_deps.bat` 安装依赖
   - 双击运行 `run.bat` 启动程序

#### 方法2：手动设置虚拟环境

1. **创建虚拟环境**

   ```bash
   python -m venv venv
   ```

2. **激活虚拟环境**

   ```bash
   # Windows PowerShell
   .\venv\Scripts\Activate.ps1
   
   # Windows Command Prompt
   venv\Scripts\activate.bat
   ```

3. **安装依赖**

   ```bash
   pip install -r requirements.txt
   ```

4. **运行程序**

   ```bash
   python src/main.py
   ```

### 使用方法

#### 1. 准备数据

- 将学生名单Excel文件放入 `data/samples/` 目录
- 将学年鉴定表Word模板放入 `data/templates/` 目录

#### 2. 配置设置

修改 `config/settings.json` 文件中的路径配置：

```json
{
  "paths": {
    "excel_file": "./data/samples/名单样例.xlsx",
    "source_dir": "./data/templates/四年制学年鉴定表",
    "output_dir": "./output/重命名后的文件"
  }
}
```

#### 3. 运行处理

程序将按以下步骤自动处理：

1. **依赖检查** - 验证系统环境
2. **路径配置** - 确认文件路径
3. **文件重命名** - 根据名单重命名文件
4. **班级选择** - 选择要处理的班级
5. **评语填写** - 自动填写标准评语
6. **PDF转换** - 转换为PDF格式

## ⚙️ 配置说明

### 主要配置项

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| `paths.excel_file` | Excel名单文件路径 | `./data/samples/名单样例.xlsx` |
| `paths.source_dir` | 源文件目录 | `./data/templates/四年制学年鉴定表` |
| `paths.output_dir` | 输出目录 | `./output/重命名后的文件` |
| `automation.auto_mode` | 自动模式开关 | `false` |
| `pdf_conversion.enabled` | 启用PDF转换 | `true` |

## 🔧 开发者指南

### 代码结构

- **核心模块** (`src/core/`): 包含主要业务逻辑
- **工具模块** (`src/utils/`): 包含配置管理、依赖检查等工具
- **主程序** (`src/main.py`): 应用程序入口和流程控制

### 扩展功能

要添加新的处理功能，可以：

1. 在 `src/core/` 中创建新的处理模块
2. 在 `src/main.py` 中集成新功能
3. 更新配置文件添加相关设置

### 构建可执行文件

```bash
cd build_tools
build.bat
```

## 📋 依赖项

### 必需依赖

- `pandas`: Excel文件读取和数据处理
- `openpyxl`: Excel文件操作
- `python-docx`: Word文档操作
- `comtypes`: Windows COM接口（用于Office自动化）

## 🐛 故障排除

### 常见问题

1. **ImportError: No module named 'xxx'**

   ```bash
   pip install -r requirements.txt
   ```

2. **WPS/Office无法启动**
   - 确保已安装WPS Office或Microsoft Office
   - 检查COM组件是否正确注册

3. **文件路径错误**
   - 检查配置文件中的路径设置
   - 确保文件和目录存在

4. **编码问题**
   - 确保Excel文件使用UTF-8编码
   - 检查Word模板文件格式

## 📖 更新日志

### v2.0.0 (2025-06-07)

- 🔄 完全重构代码架构
- 📁 重新组织项目目录结构
- ⚙️ 新增配置管理系统
- 🛠️ 改进依赖管理和错误处理
- 🎨 优化用户界面和交互体验

### v1.0.0

- 📝 基础文件重命名功能
- ✍️ 评语填写功能
- 📄 PDF转换功能

> 💡 **提示**: 建议在首次使用前仔细阅读本文档，确保正确配置所有必要的参数。
