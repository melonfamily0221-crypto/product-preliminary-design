---
name: functional-design-merge-to-word
description: 将批量生成的功能设计 Markdown 模块合并为符合模板的单个文档，通过 Pandoc 将其转换为 Word (docx) 格式，并使用 python-docx 为表格注入边框。
---

# 功能设计合并至 Word (Functional Design Merge to Word)

该技能自动化处理由 `functional-design-generator` 技能生成的各个独立模块的 Markdown 文件。它将这些文件合并为一个连贯、层级分明的功能设计文档，并导出为可用于直接演示的 Word (.docx) 格式。

## 何时使用此技能
当用户要求“合并文件”、“生成汇总设计文档”、“转换为 Word 文档”或“合并并导出 word”时，使用此技能。

## 前置条件
1. **Pandoc**: 系统必须已全局安装 `pandoc`，以便将 Markdown 转换为 Word 格式。
2. **python-docx**: Python 环境中必须已安装 `python-docx`（可通过 `pip install python-docx` 安装），以便为导出的 Word 表格应用标准网格边框。

## 工作流 (Workflow)

### 1. 使用辅助脚本进行合并和转换
该技能包含一个专用的辅助脚本，路径为 `.agents/skills/functional-design-merge-to-word/scripts/merge_to_word.py`。该脚本旨在稳健地处理端到端的转换过程，并避免 Windows 字符串编码问题。

### 2. 执行命令
使用以下命令结构执行该脚本。请提供源 Markdown 文件夹的目录、目标输出 Markdown 文件路径、目标输出 Word 文档路径以及总模块的名称。

```powershell
python .agents\skills\functional-design-merge-to-word\scripts\merge_to_word.py `
  --input_dir "docs\pdd\functional-design\系统管理" `
  --output_md "docs\pdd\achievement\2_功能设计-系统管理.md" `
  --output_docx "docs\pdd\achievement\2_功能设计-系统管理.docx" `
  --module_name "系统管理"
```

### 3. 脚本自动执行的操作
- **层级降级 (Hierarchical Shifting)**: 将所有内部标题统一下调 4 个级别（`#` → `#####`），并封顶为 6 级（`######`），防止超出 Markdown 规范导致渲染异常。
- **动态功能清单 (Dynamic Table of Features)**: 自动提取模块名称，并在合并文档的顶部生成"功能清单"汇总表格。
- **Pandoc Markdown 兼容性处理**: 使用正则表达式全量清除所有 HTML 注释（`<!-- ... -->`，含行尾注释），确保 Pandoc 能够正确识别表格和列表前的空行，防止渲染出现错位或中断。
- **资源路径修正 (Asset Path Correction)**: 显式将源文件中的 `../../assets/` 替换为 `../assets/`，确保在 `achievement` 目录中合并后的文件能正确找到 `assets` 目录中的图片。
- **Pandoc 转换**: 将合并后的 Markdown 文件编译为包含嵌入图片的 `.docx` Word 文档。
- **表格边框注入 (Table Border Injection)**: 使用 `python-docx` 遍历生成文档的底层 OXML 元素，并为所有表格强制应用标准的实线边框（`Table Grid`）。
