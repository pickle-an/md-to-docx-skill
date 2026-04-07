# md-to-docx-skill 项目文档

## 1. 项目概述

md-to-docx-skill 是一个强大的 Markdown 转 Word 文档转换器 Agent Skill，能够将 Markdown 文件自动转换为专业格式的 Word 文档（.docx）。

### 核心功能
- 自动版本管理：智能识别文件名中的版本号并自动递增
- Markdown 格式规范化：自动修复常见的 Markdown 格式问题
- 完整的 Markdown 支持：支持标题、段落、粗体、斜体、列表、表格、代码块、链接、图片等元素
- 专业文档格式：符合中文文档规范的字体和字号设置，智能段落排版
- 模板支持：支持自定义 Word 模板

## 2. 目录结构

```
md-to-docx-skill/
├── skill/
│   ├── SKILL.md              # Skill 详细说明文档
│   ├── md_to_docx.py         # 主转换脚本
│   ├── markdown_normalizer.py    # Markdown 格式规范化
│   ├── version_manager.py        # 自动版本编号
│   ├── create_template.py        # 模板生成脚本
│   ├── create_preview.py         # 预览生成脚本
│   ├── template.docx             # 默认 Word 模板
│   └── template_preview.docx     # 格式规范预览文档
├── test_md/
│   ├── test_markdown1.md         # 测试 Markdown 文件
│   ├── test_markdown_V1.docx     # 测试输出 Word 文档
│   └── test_markdown_V1_normalized.md  # 测试输出规范化 Markdown 文件
└── README.md                    # 项目说明文档
```

## 3. 系统架构

### 处理流程

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   输入 MD 文件   │ ──▶ │  版本号管理处理  │ ──▶ │  格式规范化处理  │ ──▶ │   解析 MD 元素   │ ──▶ │  生成 Word 文档  │
└─────────────────┘     └─────────────────┘     └─────────────────┘     └─────────────────┘     └─────────────────┘
                                                      │
                                            ┌─────────┴─────────┐
                                            ▼                   ▼
                                      ┌───────────┐       ┌───────────┐
                                      │ 保存规范化 │       │ 应用模板   │
                                      │ 后的文件   │       │ 样式       │
                                      └───────────┘       └───────────┘
```

### 模块关系

| 模块 | 职责 | 依赖 |
|------|------|------|
| md_to_docx.py | 主转换脚本，负责整体流程控制 | markdown_normalizer.py, version_manager.py, python-docx |
| markdown_normalizer.py | Markdown 格式规范化 | 无 |
| version_manager.py | 自动版本编号 | 无 |
| create_template.py | 模板生成 | python-docx |
| create_preview.py | 预览生成 | python-docx |

## 4. 核心模块

### 4.1 主转换模块 (md_to_docx.py)

主转换模块是整个项目的核心，负责协调各个子模块的工作，完成从 Markdown 到 Word 文档的转换。

主要功能：
- 解析 Markdown 内容，识别各种元素
- 应用格式规范化
- 生成 Word 文档
- 管理版本号

### 4.2 格式规范化模块 (markdown_normalizer.py)

格式规范化模块负责修复 Markdown 中的常见格式问题，确保生成的文档格式正确。

主要功能：
- 修复未闭合的代码块
- 修复标题格式
- 修复列表格式
- 修复表格格式
- 修复内联格式（粗体、斜体等）
- 修复空行和间距

### 4.3 版本管理模块 (version_manager.py)

版本管理模块负责智能识别和管理文件版本号，确保生成的文档有正确的版本标识。

主要功能：
- 提取文件名中的版本号
- 自动递增版本号
- 生成带版本号的文件名
- 扫描目录中已有的版本文件

## 5. 关键类与函数

### 5.1 MarkdownParser 类

**功能**：解析 Markdown 内容，识别各种元素并转换为内部表示。

**主要方法**：
- `parse(content: str) -> List[Dict[str, Any]]`：解析 Markdown 内容，返回元素列表
- `_parse_image(line: str) -> List[Dict[str, Any]]`：解析包含图片的行
- `_parse_cell_content(cell_text: str) -> List[Dict[str, Any]]`：解析表格单元格内容
- `_flush_table()`：处理表格数据

### 5.2 TextFormatter 类

**功能**：处理 Markdown 中的内联格式，如粗体、斜体、代码等。

**主要方法**：
- `parse_inline(text: str) -> List[Dict[str, Any]]`：解析内联格式
- `parse_link(text: str) -> Optional[Dict[str, Any]]`：解析链接

### 5.3 DocxGenerator 类

**功能**：生成 Word 文档，应用样式和格式。

**主要方法**：
- `create_document(output_path: str)`：创建文档
- `generate(elements: List[Dict[str, Any]], version: str = '', date: str = '', enable_code_blocks: bool = False)`：生成文档内容
- `add_title(text: str, is_first: bool = False, version: str = '', date: str = '')`：添加标题
- `add_heading(level: int, text: str)`：添加标题
- `add_paragraph(text: str)`：添加段落
- `add_bullet(text: str, level: int = 1)`：添加无序列表项
- `add_ordered_item(text: str, number: int, level: int = 1)`：添加有序列表项
- `add_table(headers: List[List[Dict[str, Any]]], data: List[List[List[Dict[str, Any]]]])`：添加表格
- `add_code_block(code: str, language: str = '')`：添加代码块
- `add_image(img_path: str, alt_text: str = '', max_width: float = 15.0)`：添加图片
- `save(output_path: str)`：保存文档

### 5.4 MarkdownNormalizer 类

**功能**：规范化 Markdown 格式，修复常见问题。

**主要方法**：
- `normalize(content: str) -> str`：规范化 Markdown 内容
- `_fix_code_blocks(lines: List[str]) -> List[str]`：修复代码块
- `_fix_tables(lines: List[str]) -> List[str]`：修复表格
- `_fix_headings(lines: List[str]) -> List[str]`：修复标题
- `_fix_lists(lines: List[str]) -> List[str]`：修复列表
- `_fix_inline_formatting(lines: List[str]) -> List[str]`：修复内联格式
- `get_fixes() -> List[str]`：获取应用的修复

### 5.5 VersionManager 类

**功能**：管理文件版本号，生成带版本号的文件名。

**主要方法**：
- `extract_version(filename: str) -> Tuple[str, Optional[int]]`：提取版本号
- `get_next_version(filename: str) -> int`：获取下一个版本号
- `generate_versioned_filename(base_path: str, extension: str, version: Optional[int] = None, suffix: str = '') -> str`：生成带版本号的文件名
- `find_latest_version_file(base_path: str, extension: str) -> Tuple[Optional[str], int]`：查找最新版本文件

### 5.6 核心函数

**convert_markdown_to_docx**
- **功能**：将 Markdown 内容转换为 Word 文档
- **参数**：
  - `markdown_content: str`：Markdown 内容
  - `output_path: str`：输出文件路径
  - `template_path: str = None`：模板路径
  - `version: str = ''`：版本号
  - `date: str = ''`：日期
  - `normalize: bool = True`：是否启用格式规范化
  - `save_normalized: bool = True`：是否保存规范化后的文件
  - `normalized_output_path: Optional[str] = None`：规范化文件输出路径
  - `enable_code_blocks: bool = True`：是否启用代码块
  - `md_file_path: str = None`：Markdown 文件路径
  - `verbose: bool = False`：是否启用详细输出
- **返回值**：`Tuple[str, List[str]]`：输出路径和应用的修复

**convert_markdown_file**
- **功能**：将 Markdown 文件转换为 Word 文档
- **参数**：
  - `input_path: str`：输入文件路径
  - `template_path: str = None`：模板路径
  - `output_path: Optional[str] = None`：输出文件路径
  - `version: str = ''`：版本号
  - `date: str = ''`：日期
  - `normalize: bool = True`：是否启用格式规范化
  - `save_normalized: bool = True`：是否保存规范化后的文件
  - `enable_code_blocks: bool = True`：是否启用代码块
  - `verbose: bool = False`：是否启用详细输出
  - `use_versioning: bool = True`：是否启用版本管理
- **返回值**：`Tuple[str, str, List[str]]`：输出路径、规范化文件路径和应用的修复

## 6. 依赖关系

| 依赖 | 版本 | 用途 | 必需性 |
|------|------|------|--------|
| python-docx | >= 0.8.11 | 生成 Word 文档 | 必需 |
| requests | - | 下载网络图片 | 可选 |
| Pillow | - | 处理图片 | 可选 |

### 安装依赖

```bash
pip install python-docx

# 可选依赖
pip install requests Pillow
```

## 7. 运行方式

### 7.1 命令行运行

```bash
# 基本用法
python skill/md_to_docx.py <input.md> [template.docx] [output.docx] [version] [date]

# 选项
--no-normalize       跳过 Markdown 规范化
--no-save-norm       不保存规范化的 Markdown 文件
--enable-code-blocks 启用代码块渲染
--no-versioning      禁用自动版本编号
--verbose            显示详细处理信息
```

### 7.2 作为模块使用

```python
from skill.md_to_docx import convert_markdown_file

# 基本转换
output_path, normalized_path, fixes = convert_markdown_file('document.md')
print(f"文档已创建: {output_path}")
if normalized_path:
    print(f"规范化 MD: {normalized_path}")
print(f"应用的修复: {len(fixes)}")

# 带参数的转换
output_path, normalized_path, fixes = convert_markdown_file(
    'document.md',
    template_path='template.docx',
    output_path='output.docx',
    version='V1',
    date='2024年01月01日',
    normalize=True,
    save_normalized=True,
    enable_code_blocks=True,
    verbose=True,
    use_versioning=True
)
```

## 8. 部署流程

### 8.1 环境要求
- Python 3.6+
- python-docx 库
- 可选：requests、Pillow 库

### 8.2 安装步骤

1. 克隆项目仓库
   ```bash
   git clone <repository-url>
   cd md-to-docx-skill
   ```

2. 安装依赖
   ```bash
   pip install python-docx
   # 可选依赖
   pip install requests Pillow
   ```

3. 测试运行
   ```bash
   python skill/md_to_docx.py test_md/test_markdown1.md
   ```

## 9. 常见问题与解决方案

| 问题 | 原因 | 解决方案 |
|------|------|----------|
| 图片无法加载 | 图片路径错误或网络问题 | 检查图片路径是否正确，确保网络连接正常 |
| 表格格式不正确 | Markdown 表格格式不规范 | 启用格式规范化，或手动修复表格格式 |
| 代码块不显示 | 代码块语法错误 | 确保代码块有正确的开始和结束标记 |
| 版本号不递增 | 文件名格式不正确 | 确保文件名符合版本号格式，如 `document_V1.md` |
| 模板样式不应用 | 模板文件不存在或格式错误 | 确保模板文件存在且格式正确 |

## 10. 示例用法

### 10.1 基本转换

```bash
# 转换 Markdown 文件为 Word 文档
python skill/md_to_docx.py document.md

# 输出:
document_V1.docx
document_V1_normalized.md
```

### 10.2 使用自定义模板

```bash
# 使用自定义模板转换
python skill/md_to_docx.py document.md custom_template.docx
```

### 10.3 禁用版本管理

```bash
# 禁用自动版本编号
python skill/md_to_docx.py document.md --no-versioning

# 输出:
document.docx
document_normalized.md
```

### 10.4 禁用格式规范化

```bash
# 跳过 Markdown 规范化
python skill/md_to_docx.py document.md --no-normalize
```

## 11. 格式规范

### 11.1 字体规范

| 元素类型 | 中文字体 | 英文字体 | 字号 | 说明 |
|---------|---------|---------|------|------|
| 正文 | 宋体 | Times New Roman | 12pt（小四） | 标准正文字号 |
| 一级标题 | 宋体 | Times New Roman | 22pt（二号） | 大标题 |
| 二级标题 | 宋体 | Times New Roman | 16pt（三号） | 章节标题 |
| 三级标题 | 宋体 | Times New Roman | 15pt（小三） | 小节标题 |
| 四级标题 | 宋体 | Times New Roman | 14pt（四号） | 条目标题 |
| 五级标题 | 宋体 | Times New Roman | 14pt（四号） | 子条目标题 |
| 代码块 | Consolas | Consolas | 9pt（小五） | 略小于正文 |
| 行内代码 | Consolas | Consolas | 12pt（小四） | 与正文同字号 |

### 11.2 段落规范

| 属性 | 设置值 | 说明 |
|------|-------|------|
| 首行缩进 | 0.74cm | 约两个汉字宽度 |
| 行间距 | 1.5 倍 | 提升阅读舒适度 |
| 段前间距 | 0pt | 保持紧凑排版 |
| 段后间距 | 0pt | 保持紧凑排版 |

### 11.3 表格规范

| 属性 | 设置值 | 说明 |
|------|-------|------|
| 表格样式 | Table Grid | 带边框的标准表格 |
| 对齐方式 | 居中 | 表格整体居中显示 |
| 列宽 | 自动计算 | 根据内容智能分配 |
| 表头背景 | #D9D9D9 | 浅灰色背景突出表头 |
| 表头对齐 | 居中 | 表头文字居中对齐 |
| 单元格对齐 | 左对齐 | 数据内容左对齐 |

### 11.4 代码块规范

| 属性 | 设置值 | 说明 |
|------|-------|------|
| 字体 | Consolas | 等宽字体，代码清晰 |
| 字号 | 9pt | 略小于正文 |
| 背景色 | #F5F5F5 | 浅灰色背景区分代码 |
| 左缩进 | 0.5cm | 突出代码块层次 |
| 语言标签 | 斜体显示 | 如`[python]` |

### 11.5 引用块规范

| 属性 | 设置值 | 说明 |
|------|-------|------|
| 左边框 | #6366F1 | 紫色竖线标识 |
| 边框宽度 | 1.5pt | 清晰可见 |
| 左右缩进 | 1cm | 突出引用内容 |
| 字体样式 | 斜体 | 区分引用文字 |

## 12. 输出文件

转换后，生成以下文件：

| 文件 | 描述 |
|------|------|
| `document_V{n}.docx` | 带版本号的最终 Word 文档 |
| `document_V{n}_normalized.md` | 带版本号的规范化 Markdown |

**版本号示例**：
- 首次转换：`document_V1.docx`、`document_V1_normalized.md`
- 第二次转换：`document_V2.docx`、`document_V2_normalized.md`
- 带版本输入：`document_V3.md` → `document_V4.docx`

## 13. 技术栈

| 技术 | 用途 |
|------|------|
| Python 3.6+ | 主要开发语言 |
| python-docx | 生成 Word 文档 |
| re (正则表达式) | 解析和处理文本 |
| os (操作系统接口) | 文件系统操作 |
| requests (可选) | 下载网络图片 |
| Pillow (可选) | 处理图片 |

## 14. 总结

md-to-docx-skill 是一个功能强大的 Markdown 转 Word 文档工具，具有以下特点：

1. **自动化处理**：自动处理版本号、格式规范化等繁琐任务
2. **专业格式**：生成符合中文文档规范的 Word 文档
3. **完整支持**：支持几乎所有 Markdown 元素
4. **灵活配置**：提供多种参数和选项，适应不同需求
5. **易于集成**：可以作为模块集成到其他项目中

该工具适用于技术文档转换、项目报告生成、会议纪要格式化、知识库文档标准化等场景，为用户提供了一种便捷的方式来将 Markdown 文档转换为专业的 Word 文档。