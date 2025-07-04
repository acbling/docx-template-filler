import openpyxl
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def center_align_table_rows(table, row_indices):
    """将指定行的所有单元格内容水平居中"""
    for row_idx in row_indices:
        row = table.rows[row_idx]
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

def format_excel_date(excel_date):
    """Excel日期转换为中文格式"""
    if not excel_date:
        return ""
    if isinstance(excel_date, datetime):
        return excel_date.strftime('%Y年%m月%d日')
    try:
        # 处理Excel的序列号
        return datetime.fromordinal(int(excel_date) + 693594).strftime('%Y年%m月%d日')
    except:
        return str(excel_date)

def set_font_fangsong(run):
    """设置为仿宋_GB2312字体，5号（10.5磅）"""
    run.font.name = '仿宋_GB2312'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')
    run.font.size = Pt(10.5)

def safe_fill_cell(cell, text):
    """写入文本并设置字体"""
    if not text:
        text = ""
    # 清空已有内容
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.clear()
    p = cell.paragraphs[0]
    run = p.add_run(str(text))
    set_font_fangsong(run)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

def safe_fill_multiline(cell, text, last_line_right_align=False, first_line_indent=False):
    """多行文本，最后一行右对齐，首段缩进2字符"""
    if not text:
        text = ""
    # 清空内容
    for paragraph in cell.paragraphs:
        p = paragraph._element
        p.getparent().remove(p)

    paragraphs = str(text).split('br') if 'br' in str(text) else [str(text)]
    n = len(paragraphs)

    for i, para_text in enumerate(paragraphs):
        para_text = para_text.strip()
        if not para_text:
            continue
        p = cell.add_paragraph()
        run = p.add_run(para_text)
        set_font_fangsong(run)

        if i == n - 1 and last_line_right_align:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        if i == 0 and first_line_indent:
            p.paragraph_format.first_line_indent = Pt(21)

def fill_template_preserve_formatting(excel_path, template_path, output_folder):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    os.makedirs(output_folder, exist_ok=True)

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

        doc = Document(template_path)
        table = doc.tables[0]

        # 基本信息填充
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

        # 多段文本填充
        safe_fill_multiline(table.cell(6, 1), data['拟办意见'], last_line_right_align=False, first_line_indent=False)
        # 你可以根据需要启用下面几行
        # safe_fill_multiline(table.cell(7, 1), data['批示意见'], last_line_right_align=True)
        # safe_fill_multiline(table.cell(8, 1), data['传阅意见'])
        # safe_fill_multiline(table.cell(9, 1), data['办理情况'])

        center_align_table_rows(table, [0, 1, 2, 3])

        # 文件名生成
        收文日期 = ws.cell(row=row_idx, column=4).value
        if isinstance(收文日期, datetime):
            收文日期_str = 收文日期.strftime('%Y%m%d')
        else:
            try:
                收文日期_str = datetime.fromordinal(int(收文日期) + 693594).strftime('%Y%m%d')
            except:
                收文日期_str = "日期未知"

        文件标题 = str(data['文件标题']) if data['文件标题'] else "无标题"
        文件标题 = 文件标题.replace('/', '-').replace('\\', '-')
        文件标题短 = 文件标题[:30] + ('…' if len(文件标题) > 30 else '')

        output_filename = f"{收文日期_str}党委组织部（党校）收文处理笺（{文件标题短}）.docx"
        output_path = os.path.join(output_folder, output_filename)

        doc.save(output_path)
        print(f"✅ 已生成 {output_path}")

def select_excel():
    path = filedialog.askopenfilename(title="选择 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls")])
    if path:
        excel_path_var.set(path)

def select_output_dir():
    path = filedialog.askdirectory(title="选择输出文件夹")
    if path:
        output_dir_var.set(path)

def run_fill():
    excel_path = excel_path_var.get()
    output_folder = output_dir_var.get()

    if not excel_path or not os.path.exists(excel_path):
        messagebox.showerror("错误", "请选择有效的 Excel 文件！")
        return
    if not output_folder or not os.path.isdir(output_folder):
        messagebox.showerror("错误", "请选择有效的保存目录！")
        return

    # 自动寻找模板文件：当前脚本同目录下的模板.docx
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, "template.docx")
    if not os.path.exists(template_path):
        messagebox.showerror("错误", f"找不到模板文件：{template_path}")
        return

    try:
        fill_template_preserve_formatting(excel_path, template_path, output_folder)
        messagebox.showinfo("完成", "所有收文处理笺已成功生成！")
    except Exception as e:
        messagebox.showerror("错误", f"生成过程中出现错误:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("收文处理笺生成工具")

    excel_path_var = tk.StringVar()
    output_dir_var = tk.StringVar()

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack()

    # Excel 文件选择
    tk.Label(frame, text="Excel 文件：").grid(row=0, column=0, sticky="e")
    tk.Entry(frame, textvariable=excel_path_var, width=50).grid(row=0, column=1, padx=5)
    tk.Button(frame, text="选择", command=select_excel).grid(row=0, column=2)

    # 输出文件夹选择
    tk.Label(frame, text="保存目录：").grid(row=1, column=0, sticky="e")
    tk.Entry(frame, textvariable=output_dir_var, width=50).grid(row=1, column=1, padx=5)
    tk.Button(frame, text="选择", command=select_output_dir).grid(row=1, column=2)

    # 生成按钮
    tk.Button(frame, text="生成收文处理笺", command=run_fill, bg="#4CAF50", fg="white", width=20).grid(row=2, column=0, columnspan=3, pady=15)

    root.mainloop()
