# Office 文件批量转 Markdown 工具

> 将 Word、Excel、PowerPoint、PDF 文件批量转换为 Markdown 格式，适合知识库迁移、文档归档、AI 数据预处理等场景。

**[English](README.md)**

## 功能特性

- **批量转换**：支持 `.docx` `.xlsx` `.pptx` `.pdf` 直接转换，`.doc` `.xls` `.ppt` 自动升级后转换
- **Excel 智能转换**：保留合并单元格、百分比、货币、千分位、日期等数字格式
- **Word 标题识别**：支持中英文标题样式映射（标题 1-6 / Heading 1-6）
- **PPT 幻灯片编号**：自动转换为 Markdown 二级标题
- **YAML 元数据**：每个输出文件头部包含标题、来源、格式、转换时间等元信息
- **旧格式自动升级**：`.doc/.xls/.ppt` 通过 COM 接口自动转为新格式（需安装 Microsoft Office）
- **子目录结构保持**：输入目录中的子文件夹结构会在输出目录中自动保留

## 运行环境

| 项目 | 要求 |
|------|------|
| 操作系统 | Windows（旧格式转换依赖 COM，新格式转换跨平台） |
| Python | 3.8 或更高版本 |
| Microsoft Office | 仅在需要处理 `.doc/.xls/.ppt` 旧格式时必需 |

## 快速开始

### 1. 安装依赖

```bash
# 方式一：使用安装脚本（推荐，含 pywin32 自动配置）
python install_deps.py

# 方式二：使用 pip 直接安装
pip install -r requirements.txt
```

### 2. 放入文件

将需要转换的 Office 文件放入 `input` 文件夹：

```
input/
├── 报告.docx
├── 数据表.xlsx
├── 演示.pptx
└── 合同.pdf
```

支持的文件格式：

| 格式 | 扩展名 | 说明 |
|------|--------|------|
| Word | `.docx` | 直接转换 |
| Word | `.doc` | 自动升级为 `.docx` 后转换（需 Microsoft Office） |
| Excel | `.xlsx` | 直接转换（支持合并单元格、数字格式） |
| Excel | `.xls` | 自动升级为 `.xlsx` 后转换（需 Microsoft Office） |
| PowerPoint | `.pptx` | 直接转换 |
| PowerPoint | `.ppt` | 自动升级为 `.pptx` 后转换（需 Microsoft Office） |
| PDF | `.pdf` | 直接转换 |

### 3. 运行转换

```bash
python batch_convert_v2.py
```

或直接双击 `batch_convert_v2.py` 文件运行。

程序会自动完成：
1. 扫描 `input` 文件夹中的所有文件
2. 如有旧格式文件（`.doc/.xls/.ppt`），先调用 Office 升级为新格式
3. 将所有文件转换为 Markdown，输出到 `output` 文件夹
4. 生成转换日志文件

### 4. 查看结果

```
output/
├── 报告.md
├── 数据表.md
├── 演示.md
└── 合同.md
```

## 目录结构

```
office-file-convert/
├── input/                ← 把要转换的文件放这里
├── output/               ← 转换后的 Markdown 文件在这里
├── upgraded/             ← 旧格式升级后的中间文件（自动生成）
├── docs/                 ← 文档和截图
├── batch_convert_v2.py   ← 主程序
├── install_deps.py       ← 依赖安装脚本
├── fix_old_files.py      ← 独立的旧格式批量升级工具
├── requirements.txt      ← Python 依赖列表
├── LICENSE.txt           ← MIT 许可证
└── README.md             ← 英文主文档
```

## 配置说明

`batch_convert_v2.py` 中的配置区域（`batch_convert()` 函数开头）可修改以下参数：

```python
base_dir = os.path.dirname(os.path.abspath(__file__))  # 默认：脚本所在目录

input_dir = os.path.join(base_dir, "input")     # 输入目录
output_dir = os.path.join(base_dir, "output")   # 输出目录
upgraded_dir = os.path.join(base_dir, "upgraded")  # 旧格式升级目录
```

## 常见问题

### 旧格式文件（.doc/.xls/.ppt）转换失败

旧格式转换依赖 Microsoft Office 的 COM 接口，请确保：
- 已安装 Microsoft Office（WPS 不支持）
- Office 没有弹出激活提示或更新弹窗（会阻塞转换）
- 转换期间不要手动打开 Office 文件

### pywin32 安装后仍报错

手动执行 post-install 脚本：

```bash
python -m pywin32_postinstall -install
```

### 仅使用新格式（无需旧格式升级）

如果只处理 `.docx/.xlsx/.pptx/.pdf`，可以跳过 `pywin32` 依赖：

```bash
pip install markitdown openpyxl
```

此场景下可在任意操作系统上运行。

## 许可证

[MIT License](LICENSE.txt)
