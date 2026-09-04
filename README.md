# 产品初步设计 (Product Preliminary Design)

这是一个通过 AI 辅助驱动的产品初步设计与文档生成项目。本项目结合了自定义大模型技能 (Skills) 与自动化文档处理流程，旨在从原始需求/设计文件快速生成标准化功能设计文档，并最终将其整合输出为交付级别的 Word 文档。

## 📁 目录结构

项目的主要目录结构及说明如下：

```
20260903-product-preliminary-design/
├── .agents/                    # AI 助手自定义配置与技能目录
│   └── skills/                 # 本项目特有的自动化处理技能
│       ├── template-file-generator         # 技能：从原始 Word 文档提取骨架 Markdown 模板
│       ├── functional-design-generator     # 技能：结合 UI 原型生成标准的 MD 功能设计文档
│       └── functional-design-merge-to-word # 技能：将多个 MD 文件合并并转换为格式化的 Word 文档
├── docs/                       # 项目核心文档目录
│   ├── pdd/                    # 产品设计文档 (Product Design Document) 工作区
│   │   ├── original-document/  # 原始输入文件，如《概要设计说明书.docx》
│   │   ├── template-file/      # 通过 AI 解析原始文档生成的 Markdown 骨架模板
│   │   ├── functional-design/  # 针对各模块生成的具体功能设计文档 (Markdown 格式)
│   │   ├── achievement/        # 最终合并和导出的交付成果 (如合并后的 Word 文档)
│   │   └── assets/             # 设计文档中关联的静态资源和图片
│   └── prompt-sample/          # AI 交互提示词样例，例如批量导出等快捷命令
└── README.md                   # 项目说明文档
```

## 🛠️ 核心工作流 (Workflow)

本项目定义了一套标准的 AI 辅助文档产出流程，主要包含以下几个阶段：

1. **模板提取 (Template Generation)**
   - **输入**: 将上游提供的文档（如 `docs/pdd/original-document/概要设计说明书.docx`）放入系统。
   - **操作**: 调用 `template-file-generator` 技能。
   - **输出**: 在 `docs/pdd/template-file/` 目录下生成各业务模块的 Markdown 骨架模板。

2. **功能设计生成 (Functional Design Generation)**
   - **输入**: 前期生成的 Markdown 模板 + UI 原型信息。
   - **操作**: 调用 `functional-design-generator` 技能。
   - **输出**: 在 `docs/pdd/functional-design/` 目录下生成带有详细页面布局、控件规则、交互规则的标准化功能说明 MD 文档。

3. **文档合并与导出 (Document Merge & Export)**
   - **输入**: 散落在 `docs/pdd/functional-design/` 中的 Markdown 文件。
   - **操作**: 调用 `functional-design-merge-to-word` 技能（参考 `docs/prompt-sample/prompt.md` 中的指令）。
   - **输出**: 在 `docs/pdd/achievement/` 目录下生成整合后的 Word 文档，系统将自动使用 Pandoc 转换并调整表格边框等排版格式，以符合企业交付规范。

## 💡 使用指南

您可以通过与大语言模型对话，结合系统内预置的 Skills 来推进设计工作。

例如，您可以对 AI 助手下达如下指令：
- *“帮我基于 docs/pdd/original-document/概要设计说明书.docx 提取功能设计模板。”*
- *“结合原型截图，完善 系统管理 模块的功能设计。”*
- *“把 docs/pdd/系统管理 模块合并并导出 word。”* (详见 `prompt-sample`)

## 📝 备注

请确保原始文档符合一定的结构规范，以便 AI 能够更精准地解析其层级，从而保障后续自动化步骤的顺畅进行。
