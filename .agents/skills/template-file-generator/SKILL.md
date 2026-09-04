---
name: template-file-generator
description: >-
  当用户希望从 Word (.docx) 格式的产品需求文档（PRD）或系统设计文档（如概要设计说明书.docx）中提取骨架 Markdown 模板时，使用此技能。它会自动解析文档的标题层级，并为 UI 页面、业务流程、数据库和 API 等章节注入标准占位符模板。
---

# 模板文件生成器 (Template File Generator)

该技能从 Microsoft Word 文档（.docx）中提取层级标题结构，并生成简洁的骨架 Markdown 模板。它会剥离所有正文段落、表格和原始数据，仅保留文档的结构性标题，然后针对特定功能领域智能注入标准化占位符模板（如空表格和 HTML 注释），涵盖：

- **UI 设计（页面原型、页面字段等）**
- **业务流程（业务流转、Mermaid 流程图）**
- **数据库设计（数据库表清单、ER 图）**
- **接口设计（接口参数、接收数据清单）**

## 执行步骤

1. **定位输入文档**：确认用户 `.docx` 文件的路径（例如：`docs/pdd/original-document/概要设计说明书.docx`）。
2. **确定输出路径**：确认生成的 Markdown 模板文件的保存位置（例如：`docs/pdd/概要设计模板.md`）。
3. **运行生成脚本**：携带 `--input` 和 `--output` 参数执行辅助 Python 脚本 `generate_template.py`。

### 示例命令

```powershell
python .agents/skills/template-file-generator/scripts/generate_template.py --input "docs\pdd\original-document\概要设计说明书.docx" --output "docs\pdd\概要设计模板.md"
```

4. **验证输出**：脚本执行完毕后，确认骨架模板文件已被正确生成，并告知用户。
