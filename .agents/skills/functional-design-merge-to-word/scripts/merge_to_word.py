import os
import glob
import re
import argparse
import subprocess
import sys

def parse_args():
    parser = argparse.ArgumentParser(description="Merge Functional Design Markdown files and convert to Word with borders.")
    parser.add_argument("--input_dir", required=True, help="Directory containing the module markdown files (e.g. docs/pdd/系统管理)")
    parser.add_argument("--output_md", required=True, help="Path for the merged Markdown output")
    parser.add_argument("--output_docx", required=True, help="Path for the final Word document")
    parser.add_argument("--module_name", required=True, help="Name of the top-level module (e.g. 系统管理)")
    parser.add_argument("--sys_code", default="SYS01", help="System Code for the feature list")
    parser.add_argument("--sys_name", default="后台管理系统", help="System Name for the feature list")
    return parser.parse_args()

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

    for filepath in md_files:
        with open(filepath, 'r', encoding='utf-8-sig') as f:
            content = f.read()

        # 统一行尾符
        content = content.replace('\r', '')

        # Bug 3 修复：显式修正图片相对路径
        # 源文件在 functional-design/{模块}/ 层级（../../assets/）
        # 合并后文件在 achievement/ 层级（../assets/）
        content = content.replace('../../assets/', '../assets/')

        processed_lines = []
        for line in content.split('\n'):
            # Bug 2 修复：用正则全量清除行内 HTML 注释（含行尾注释）
            line = re.sub(r'<!--.*?-->', '', line)
            # 若整行清除后只剩空白，则置为空行
            if not line.strip():
                processed_lines.append("")
                continue

            # Bug 1 修复：标题降级封顶为 6 级，防止超出 Markdown 规范
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
