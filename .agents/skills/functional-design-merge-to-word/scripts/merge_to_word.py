import os
import glob
import re
import argparse
import subprocess
import sys
import base64
import urllib.request
import urllib.parse

def parse_args():
    parser = argparse.ArgumentParser(description="Merge Functional Design Markdown files and convert to Word with borders.")
    parser.add_argument("--input_dir", required=True, help="Directory containing the module markdown files (e.g. docs/functional-design/系统管理)")
    parser.add_argument("--output_md", required=True, help="Path for the merged Markdown output")
    parser.add_argument("--output_docx", required=True, help="Path for the final Word document")
    parser.add_argument("--module_name", required=True, help="Name of the top-level module (e.g. 系统管理)")
    parser.add_argument("--sys_code", default="SYS01", help="System Code for the feature list")
    parser.add_argument("--sys_name", default="后台管理系统", help="System Name for the feature list")
    parser.add_argument("--order", help="Comma-separated list of module names to specify display order (e.g. 用户管理,角色管理,菜单管理...)")
    return parser.parse_args()

def render_mermaid_diagram(mermaid_code, output_image_path):
    """将 Mermaid 代码渲染为 PNG 图片保存到指定路径。若已存在则直接复用。"""
    if os.path.exists(output_image_path) and os.path.getsize(output_image_path) > 0:
        return True
    try:
        b64 = base64.b64encode(mermaid_code.strip().encode('utf-8')).decode('ascii')
        quoted = urllib.parse.quote(b64, safe='')
        url = f'https://mermaid.ink/img/{quoted}'
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = resp.read()
            os.makedirs(os.path.dirname(os.path.abspath(output_image_path)), exist_ok=True)
            with open(output_image_path, 'wb') as f:
                f.write(data)
            print(f"Rendered mermaid diagram to {output_image_path}")
            return True
    except Exception as e:
        print(f"Warning: Failed to render mermaid diagram via mermaid.ink: {e}")
        return False

def apply_table_borders(docx_path):
    try:
        from docx import Document
        from docx.oxml.shared import OxmlElement
        from docx.oxml.ns import qn
    except ImportError:
        print("Warning: python-docx not installed. Table borders will not be applied.")
        print("Please install via: pip install python-docx")
        return

    def set_table_borders(table):
        tbl = table._element
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
            
        for e in tblPr.findall(qn('w:tblBorders')):
            tblPr.remove(e)
            
        tblBorders = OxmlElement('w:tblBorders')
        
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')
            tblBorders.append(border)
            
        tblPr.append(tblBorders)

    doc = Document(docx_path)
    count = 0
    for table in doc.tables:
        set_table_borders(table)
        count += 1
    doc.save(docx_path)
    print(f"Applied table borders to {count} tables in {docx_path}.")

def main():
    args = parse_args()
    
    if not os.path.isdir(args.input_dir):
        print(f"Error: Input directory {args.input_dir} does not exist.")
        sys.exit(1)
        
    md_files = glob.glob(os.path.join(args.input_dir, "*.md"))
    if not md_files:
        print(f"Warning: No markdown files found in {args.input_dir}.")
        
    if args.order:
        order_list = [x.strip() for x in args.order.split(',') if x.strip()]
        def get_sort_key(fpath):
            name = os.path.splitext(os.path.basename(fpath))[0]
            if name in order_list:
                return (0, order_list.index(name))
            return (1, name)
        md_files.sort(key=get_sort_key)
    else:
        md_files.sort()
        
    modules = [os.path.splitext(os.path.basename(f))[0] for f in md_files]
    
    out_lines = []
    out_lines.append("### 功能设计\n")
    out_lines.append("#### 功能清单\n")
    out_lines.append("| 系统编码 | 系统名称 | 一级功能编码 | 一级功能点 | 二级功能编码 | 二级功能点 |")
    out_lines.append("| :--- | :--- | :--- | :--- | :--- | :--- |")

    for i, mod in enumerate(modules):
        sec_code = f"SYS_M_{i+1:02d}"
        out_lines.append(f"| {args.sys_code} | {args.sys_name} | SYS_M | {args.module_name} | {sec_code} | {mod} |")

    out_lines.append(f"\n#### {args.module_name}\n")

    assets_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(args.output_md)), "../assets"))
    os.makedirs(assets_dir, exist_ok=True)

    for filepath in md_files:
        mod_name = os.path.splitext(os.path.basename(filepath))[0]
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            content = f.read()

        # 统一行尾符
        content = content.replace('\r', '')

        # 修正图片相对路径：源文件在 functional-design/{模块}/（../../assets/） -> achievement/（../assets/）
        content = content.replace('../../assets/', '../assets/')

        # 将 Mermaid 流程图代码块转换为渲染图，使 Word 能够直接显示图形流程图
        def replace_mermaid(match):
            code = match.group(1).strip()
            img_filename = f"flow_{mod_name}.png"
            img_abs_path = os.path.join(assets_dir, img_filename)
            success = render_mermaid_diagram(code, img_abs_path)
            if success:
                return f"\n\n![{mod_name}业务流程图](../assets/{img_filename})\n\n"
            else:
                return match.group(0)

        content = re.sub(r'```mermaid(.*?)```', replace_mermaid, content, flags=re.DOTALL)

        # 确保“流程说明：”前后均有空行，保证 Pandoc 正确识别列表并逐行换行
        content = re.sub(r'\n*(流程说明[：:])\s*\n+', r'\n\n\1\n\n', content)

        processed_lines = []
        for line in content.split('\n'):
            # 清除行内 HTML 注释（含行尾注释）
            line = re.sub(r'<!--.*?-->', '', line)
            # 若整行清除后只剩空白，则置为空行
            if not line.strip():
                processed_lines.append("")
                continue

            # 标题降级封顶为 6 级，防止超出 Markdown 规范
            match = re.match(r'^(#+)\s(.*)', line)
            if match:
                hashes = match.group(1)
                text = match.group(2)
                new_level = min(len(hashes) + 4, 6)
                processed_lines.append(f"{'#' * new_level} {text}")
            else:
                processed_lines.append(line)

        out_lines.append("\n".join(processed_lines))
        out_lines.append("\n")

    os.makedirs(os.path.dirname(os.path.abspath(args.output_md)), exist_ok=True)
    with open(args.output_md, 'w', encoding='utf-8') as f:
        f.write("\n".join(out_lines))

    print(f"Generated Merged Markdown: {args.output_md}")

    os.makedirs(os.path.dirname(os.path.abspath(args.output_docx)), exist_ok=True)
    
    work_dir = os.path.dirname(os.path.abspath(args.output_md))
    md_filename = os.path.basename(args.output_md)
    
    print("Calling pandoc...")
    result = subprocess.run(
        ["pandoc", md_filename, "-o", os.path.abspath(args.output_docx)],
        cwd=work_dir,
        capture_output=True,
        text=True
    )

    if result.returncode == 0:
        print(f"Successfully converted to {args.output_docx}")
        apply_table_borders(args.output_docx)
    else:
        print(f"Pandoc failed: {result.stderr}")
        sys.exit(1)

if __name__ == "__main__":
    main()
