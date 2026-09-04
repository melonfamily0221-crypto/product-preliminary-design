import os
import re
import argparse
import docx
from docx.document import Document
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell
from docx.text.paragraph import Paragraph

PAGE_TEMPLATE = """**1. 页面原型**
<!-- 插入该页面的原型图或UI设计图 -->

**2. 功能说明**
<!-- 详细描述页面上的各个功能按钮、操作逻辑及数据流转规则 -->
- 查询：...
- 详情：...
- 编辑：...

**3. 页面控件**
<!-- 详细列出页面上的控件元素及其交互事件 -->
| 名称 | 类型 | 位置 | 事件 | 说明 |
| :--- | :--- | :--- | :--- | :--- |
|  |  |  |  |  |

**4. 页面字段**
<!-- 详细列出页面或表单中涉及的核心数据字段及其规则 -->
| 名称 | 数据类型 | 长度 | 允许操作 | 是否必输 | 录入方式 | 备注 |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
|  |  |  |  |  |  |  |
"""

FLOW_TEMPLATE = """<!-- 插入流程图 -->
<!-- 流程节点说明表格 -->
| 节点 | 主导部门 | 程序及事项 | 活动的具体要求 | 时限要求 | 产生文件、记录 |
| :--- | :--- | :--- | :--- | :--- | :--- |
|  |  |  |  |  |  |
"""

def iter_block_items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)

def build_functional_design_exemplar(first_h4, first_h5, first_h7, other_modules):
    h4_name = first_h4 or "设备台账"
    h5_name = first_h5 or "在用设备"
    h7_name = first_h7 or "在用设备"
    
    sample_others = "、".join(other_modules[:4]) if other_modules else "特种设备、点巡检管理、保养计划、维修管理"
    note_suffix = f"（如{sample_others}等）" if sample_others else ""

    return f"""#### [一级模块名称，例如：{h4_name}]

##### [二级模块/页面名称，例如：{h5_name}]

###### 功能描述
<!-- 描述该模块或页面的核心功能和业务目标 -->

###### 业务流程
<!-- 描述该模块内的具体业务流转过程，可包含流程图和流转说明 -->

```mermaid
graph TD
    A[开始节点] --> B{{判断节点}}
    B -- 是 --> C[处理流程1]
    B -- 否 --> D[处理流程2]
    C --> E[结束节点]
    D --> E
```

流程说明：
<!-- 详细描述业务流程的各个流转步骤、流转条件及业务规则要求 -->

1) 步骤一说明...
2) 步骤二说明...
3) 步骤三说明...
4) 步骤四说明...

###### 功能设计

####### [页面/弹窗名称，例如：{h7_name}列表]

{PAGE_TEMPLATE}
####### [页面/弹窗名称，例如：新增{h7_name}]
<!-- 如上述结构重复，分别描述对应的原型、说明、控件与字段 -->

---
*注：上述为单个模块及页面的标准设计模板示例，其余各一级功能模块{note_suffix}、二级功能页面及操作弹窗均参照此标准层级推导展开。*
"""

def generate_template(docx_path, output_file, split=True):
    doc = docx.Document(docx_path)
    
    headings = []
    for block in iter_block_items(doc):
        style_name = block.style.name if block.style else 'Normal'
        text = block.text.strip()
        if not text or not style_name.startswith('Heading '):
            continue
        try:
            level = int(style_name.replace('Heading ', ''))
            headings.append((level, text))
        except ValueError:
            continue

    # 预扫描功能设计章节，提取第一个典型模块/页面及后续模块名称用于生成示例
    in_fd = False
    first_h4 = None
    first_h5 = None
    first_h7 = None
    other_modules = []

    for level, text in headings:
        if level == 3 and "功能设计" in text:
            in_fd = True
            continue
        if in_fd:
            if level <= 3:
                break
            if level == 4 and text != "功能清单":
                if first_h4 is None:
                    first_h4 = text
                else:
                    other_modules.append(text)
            elif level == 5 and first_h5 is None:
                first_h5 = text
            elif level == 7 and first_h7 is None:
                first_h7 = text

    fd_exemplar = build_functional_design_exemplar(first_h4, first_h5, first_h7, other_modules)

    output_lines = []
    in_functional_design = False
    exemplar_injected = False
    last_h4_text = ""

    for level, text in headings:
        if level == 3 and "功能设计" in text:
            in_functional_design = True
            output_lines.append(f"### {text}\n")
            continue

        if in_functional_design:
            if level <= 3:
                in_functional_design = False
            else:
                if level == 4 and text == "功能清单":
                    output_lines.append("#### 功能清单\n")
                    output_lines.append("<!-- 列出该系统的功能架构清单，包括一级功能、二级功能等 -->\n| 系统编码 | 系统名称 | 一级功能编码 | 一级功能点 | 二级功能编码 | 二级功能点 |\n| :--- | :--- | :--- | :--- | :--- | :--- |\n|  |  |  |  |  |  |\n")
                elif not exemplar_injected:
                    output_lines.append(fd_exemplar)
                    exemplar_injected = True
                continue

        if level < 8:
            output_lines.append(f"{'#' * level} {text}\n")
            if level == 4:
                last_h4_text = text

            if text == "系统流程清单":
                output_lines.append("<!-- 列出该系统涉及的核心业务流程清单表格 -->\n| 序号 | 流程代码 | 流程名称 | 备注 |\n| :--- | :--- | :--- | :--- |\n| 1 |  |  |  |\n")
            elif text == "系统流程说明":
                output_lines.append("<!-- 针对上述流程清单中的各个流程进行详细说明，可包含流程图和节点说明 -->\n")
            elif level == 5 and last_h4_text == "系统流程说明":
                output_lines.append(FLOW_TEMPLATE)
            elif text == "数据库表关系":
                output_lines.append("<!-- 插入系统的核心数据库表ER图或实体关系图 -->\n")
            elif text == "数据库表清单":
                output_lines.append("<!-- 列出该系统涉及的所有核心数据库表 -->\n| 序号 | 表名 | 代码 | 所属模块 | 保存时长 | 备份周期 |\n| :--- | :--- | :--- | :--- | :--- | :--- |\n| 1 |  |  |  |  |  |\n")
            elif text == "通讯协议":
                output_lines.append("<!-- 描述系统接口的通讯协议，如HTTP, HTTPS, MQTT, TCP等 -->\n")
            elif text == "通讯策略":
                output_lines.append("<!-- 描述通讯的方式和策略，如同步/异步，心跳机制，重试机制等 -->\n")
            elif text == "报文格式":
                output_lines.append("<!-- 描述接口请求和响应的基本报文数据格式（如JSON, XML）及公共结构 -->\n")
            elif text == "接收数据清单":
                output_lines.append("<!-- 列出系统作为服务端接收的数据接口清单 -->\n| 序号 | 源系统 | 目标系统 | 接口名称 | 接口内容描述 | 频度 | 对应业务流程 |\n| :--- | :--- | :--- | :--- | :--- | :--- | :--- |\n| 1 |  |  |  |  |  |  |\n")
            elif text == "发送数据清单":
                output_lines.append("<!-- 列出系统作为客户端发送的数据接口清单 -->\n| 序号 | 源系统 | 目标系统 | 接口名称 | 接口内容描述 | 频度 | 对应业务流程 |\n| :--- | :--- | :--- | :--- | :--- | :--- | :--- |\n| 1 |  |  |  |  |  |  |\n")
            elif level == 6 and "接口设计" in str(output_lines) and "接口名称" in text:
                output_lines.append("**1. 接口说明**\n<!-- 详细描述该接口的功能和业务场景 -->\n\n**2. 接口参数**\n<!-- 列出请求和响应参数详情 -->\n| 序号 | 编码 | 类型 | 单位 | 为空 | 描述 |\n| :--- | :--- | :--- | :--- | :--- | :--- |\n| 1 |  |  |  |  |  |\n")

    output_dir = os.path.dirname(os.path.abspath(output_file))
    os.makedirs(output_dir, exist_ok=True)
    
    full_content = "\n".join(output_lines)
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(full_content)

    if split:
        lines = full_content.splitlines(keepends=True)
        idx_gn = next((i for i, l in enumerate(lines) if l.strip() == '### 功能设计'), -1)
        idx_sjk = next((i for i, l in enumerate(lines) if l.strip() == '### 数据库设计'), -1)
        idx_jk = next((i for i, l in enumerate(lines) if l.strip() == '### 接口设计'), -1)

        if idx_gn != -1 and idx_sjk != -1 and idx_jk != -1:
            part1 = lines[:idx_gn]
            part2 = lines[idx_gn:idx_sjk]
            part3 = lines[idx_sjk:idx_jk]
            part4 = lines[idx_jk:]

            with open(os.path.join(output_dir, "1_业务概况.md"), "w", encoding="utf-8") as f:
                f.writelines(part1)
            with open(os.path.join(output_dir, "2_功能设计.md"), "w", encoding="utf-8") as f:
                f.writelines(part2)
            with open(os.path.join(output_dir, "3_数据库设计.md"), "w", encoding="utf-8") as f:
                f.writelines(part3)
            with open(os.path.join(output_dir, "4_接口设计.md"), "w", encoding="utf-8") as f:
                f.writelines(part4)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract skeleton template from DOCX")
    parser.add_argument("--input", required=True, help="Input DOCX file path")
    parser.add_argument("--output", required=True, help="Output Markdown file path")
    parser.add_argument("--no-split", action="store_true", help="Do not split into modular template files")
    args = parser.parse_args()
    
    generate_template(args.input, args.output, split=not args.no_split)
    print(f"Template successfully extracted to {args.output}")
