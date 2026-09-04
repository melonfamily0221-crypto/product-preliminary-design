import os
import re
import argparse
import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

PAGE_TEMPLATE = """**1. 页面原型**
<!-- 插入该页面的原型图或UI设计图 -->

**2. 功能说明**
<!-- 详细描述页面上的各个功能按钮、操作逻辑及数据流转规则 -->

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

def generate_template(docx_path, output_file):
    doc = docx.Document(docx_path)
    output_lines = []
    
    last_h4_text = ""
    has_injected_page_template = False
    
    for block in iter_block_items(doc):
        style_name = block.style.name if block.style else 'Normal'
        text = block.text.strip()
        
        if not text or not style_name.startswith('Heading '):
            continue
            
        try:
            level = int(style_name.replace('Heading ', ''))
        except ValueError:
            continue
            
        if level < 8:
            has_injected_page_template = False
            
        if level == 4:
            last_h4_text = text
            
        if level < 8:
            output_lines.append(f"{'#' * level} {text}\n")
            
            if text == "功能描述":
                output_lines.append("<!-- 描述该模块或页面的核心功能和业务目标 -->\n")
            elif text == "业务流程":
                output_lines.append("<!-- 描述该模块内的具体业务流转过程，可包含流程图和流转说明 -->\n\n```mermaid\ngraph TD\n    A[开始节点] --> B{判断节点}\n    B -- 是 --> C[处理流程1]\n    B -- 否 --> D[处理流程2]\n    C --> E[结束节点]\n    D --> E\n```\n")
            elif text == "功能清单":
                output_lines.append("<!-- 列出该系统的功能架构清单，包括一级功能、二级功能等 -->\n| 系统编码 | 系统名称 | 一级功能编码 | 一级功能点 | 二级功能编码 | 二级功能点 |\n| :--- | :--- | :--- | :--- | :--- | :--- |\n|  |  |  |  |  |  |\n")
            elif text == "系统流程清单":
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
                
        elif level == 8:
            if not has_injected_page_template:
                output_lines.append(PAGE_TEMPLATE)
                has_injected_page_template = True
                
    os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        f.write("\n".join(output_lines))

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract skeleton template from DOCX")
    parser.add_argument("--input", required=True, help="Input DOCX file path")
    parser.add_argument("--output", required=True, help="Output Markdown file path")
    args = parser.parse_args()
    
    generate_template(args.input, args.output)
    print(f"Template successfully extracted to {args.output}")
