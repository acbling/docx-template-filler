import openpyxl
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import tempfile
import pkgutil  # 用于读取内嵌资源

# ========== 样式辅助函数 ==========
def format_excel_date(excel_date):
    if not excel_date:
        return ""
    if isinstance(excel_date, datetime):
        return excel_date.strftime('%Y年%m月%d日')
    try:
        return datetime.fromordinal(int(excel_date) + 693594).strftime('%Y年%m月%d日')
    except:
        return str(excel_date)

def set_font_fangsong(run):
    run.font.name = '仿宋_GB2312'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
    run.font.size = Pt(10.5)

def safe_fill_cell(cell, text):
    text = str(text) if text else ""
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.clear()
    p = cell.paragraphs[0]
    run = p.add_run(text)
    set_font_fangsong(run)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

def safe_fill_multiline(cell, text, last_line_right_align=False, first_line_indent=False):
    if not text:
        text = ""
    for p in cell.paragraphs:
        p._element.getparent().remove(p._element)
    paragraphs = str(text).split('br') if 'br' in str(text) else [str(text)]
    for i, para_text in enumerate(paragraphs):
        para_text = para_text.strip()
        if not para_text:
            continue
        p = cell.add_paragraph()
        run = p.add_run(para_text)
        set_font_fangsong(run)
        if i == len(paragraphs) - 1 and last_line_right_align:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if i == 0 and first_line_indent:
            p.paragraph_format.first_line_indent = Pt(21)

def center_align_table_rows(table, row_indices):
    for row_idx in row_indices:
        for cell in table.rows[row_idx].cells:
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

# ========== 生成文档 ==========
def fill_template_preserve_formatting(excel_path, output_folder):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    os.makedirs(output_folder, exist_ok=True)

    # 加载内置模板文件
    template_data = pkgutil.get_data(__name__, "template.docx")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(template_data)
        tmp_path = tmp.name

    for row_idx in range(4, ws.max_row + 1):
        if not ws.cell(row=row_idx, column=2).value:
            continue

        data = {
            '来文单位': ws.cell(row=row_idx, column=2).value,
            '发文日期': format_excel_date(ws.cell(row=row_idx, column=3).value),
            '收文日期': format_excel_date(ws.cell(row=row_idx, column=4).value),
            '文件编号': ws.cell(row=row_idx, column=5).value,
            '文件份数': ws.cell(row=row_idx, column=6).value,
            '文件页数': ws.cell(row=row_idx, column=7).value,
            '来文类型': ws.cell(row=row_idx, column=8).value,
            '公开属性': ws.cell(row=row_idx, column=9).value,
            '缓急程度': ws.cell(row=row_idx, column=10).value,
            '来文文号': ws.cell(row=row_idx, column=11).value,
            '文件标题': ws.cell(row=row_idx, column=12).value,
            '拟办意见': ws.cell(row=row_idx, column=13).value,
            '批示意见': ws.cell(row=row_idx, column=14).value,
            '传阅意见': ws.cell(row=row_idx, column=15).value,
            '办理情况': ws.cell(row=row_idx, column=16).value,
            '督办时间': format_excel_date(ws.cell(row=row_idx, column=17).value)
        }

        doc = Document(tmp_path)
        table = doc.tables[0]

        safe_fill_cell(table.cell(1, 1), data['来文单位'])
        safe_fill_cell(table.cell(1, 3), data['发文日期'])
        safe_fill_cell(table.cell(1, 5), data['收文日期'])
        safe_fill_cell(table.cell(2, 1), data['文件编号'])
        safe_fill_cell(table.cell(2, 3), data['文件份数'])
        safe_fill_cell(table.cell(2, 5), data['文件页数'])
        safe_fill_cell(table.cell(3, 1), data['来文类型'])
        safe_fill_cell(table.cell(3, 3), data['公开属性'])
        safe_fill_cell(table.cell(3, 5), data['缓急程度'])
        safe_fill_cell(table.cell(4, 1), data['来文文号'])
        safe_fill_cell(table.cell(4, 5), data['督办时间'])
        safe_fill_cell(table.cell(5, 1), data['文件标题'])
        safe_fill_multiline(table.cell(6, 1), data['拟办意见'])

        center_align_table_rows(table, [0, 1, 2, 3])

        收文日期 = ws.cell(row=row_idx, column=4).value
        if isinstance(收文日期, datetime):
            收文日期_str = 收文日期.strftime('%Y%m%d')
        else:
            try:
                收文日期_str = datetime.fromordinal(int(收文日期) + 693594).strftime('%Y%m%d')
            except:
                收文日期_str = "日期未知"

        文件标题 = str(data['文件标题']) if data['文件标题'] else "无标题"
        文件标题短 = 文件标题[:30] + ('…' if len(文件标题) > 30 else '')
        output_filename = f"{收文日期_str}党委组织部（党校）收文处理笺（{文件标题短}）.docx"
        doc.save(os.path.join(output_folder, output_filename))

# ========== GUI 主程序 ==========
def main():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("收文处理笺生成器", "请选择 Excel 文件")
    excel_file = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
    if not excel_file:
        return
    output_dir = filedialog.askdirectory(title="选择输出目录")
    if not output_dir:
        return

    try:
        fill_template_preserve_formatting(excel_file, output_dir)
        messagebox.showinfo("完成", f"✅ 已生成 Word 文件到：{output_dir}")
    except Exception as e:
        messagebox.showerror("出错", str(e))

if __name__ == "__main__":
    main()
