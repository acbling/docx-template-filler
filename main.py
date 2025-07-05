import openpyxl
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.font as tkFont

# ========================== Word 填充辅助函数 ==============================
def center_align_table_rows(table, row_indices):
    for row_idx in row_indices:
        row = table.rows[row_idx]
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

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
    if not text:
        text = ""
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.clear()
    p = cell.paragraphs[0]
    run = p.add_run(str(text))
    set_font_fangsong(run)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

def safe_fill_multiline(cell, text, last_line_right_align=False, first_line_indent=False):
    if not text:
        text = ""
    for paragraph in cell.paragraphs[:]:
        p = paragraph._element
        p.getparent().remove(p)
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

# ========================== Excel 数据读取函数 ==============================
def extract_entries(excel_path):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    entries = []
    for row_idx in range(5, ws.max_row + 1):
        title = str(ws.cell(row=row_idx, column=12).value or "").strip()
        if not title:
            continue
        date_cell = ws.cell(row=row_idx, column=4).value
        try:
            if isinstance(date_cell, datetime):
                date_str = date_cell.strftime("%Y-%m-%d")
            else:
                date_str = datetime.fromordinal(int(date_cell) + 693594).strftime("%Y-%m-%d")
        except:
            date_str = "日期未知"
        label = f"{date_str} - {title[:30]}"
        entries.append((row_idx, label))
    return entries

# ========================== Word 文档生成函数 ===============================
def fill_template_preserve_formatting(excel_path, template_path, output_folder, selected_rows):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    os.makedirs(output_folder, exist_ok=True)

    for row_idx in selected_rows:
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
        文件标题 = 文件标题.replace('/', '-').replace('\\', '-')
        文件标题短 = 文件标题[:30] + ('…' if len(文件标题) > 30 else '')
        output_filename = f"{收文日期_str}党委组织部（党校）收文处理笺（{文件标题短}）.docx"
        output_path = os.path.join(output_folder, output_filename)
        doc.save(output_path)
        print(f"✅ 已生成 {output_path}")

# ========================== GUI 部分 ========================================
root = tk.Tk()
root.title("收文处理笺生成工具")

excel_path_var = tk.StringVar()
output_dir_var = tk.StringVar()
entry_vars = []
entries = []

frame = tk.Frame(root, padx=10, pady=10)
frame.pack()

tk.Label(frame, text="Excel 文件：").grid(row=0, column=0, sticky="e")
tk.Entry(frame, textvariable=excel_path_var, width=50).grid(row=0, column=1)
tk.Button(frame, text="选择", command=lambda: select_excel()).grid(row=0, column=2)

tk.Label(frame, text="保存目录：").grid(row=1, column=0, sticky="e")
tk.Entry(frame, textvariable=output_dir_var, width=50).grid(row=1, column=1)
tk.Button(frame, text="选择", command=lambda: select_output_dir()).grid(row=1, column=2)

tk.Label(frame, text="选择要生成的条目：").grid(row=2, column=0, sticky="ne", pady=10)

entries_frame = tk.Frame(frame)
entries_frame.grid(row=2, column=1, pady=10, sticky="w")

canvas = tk.Canvas(entries_frame, width=460, height=220, bg="#ffffff")
scrollbar = tk.Scrollbar(entries_frame, orient="vertical", command=canvas.yview)
font = tkFont.Font(family="微软雅黑", size=10)
scrollable_frame = tk.Frame(canvas, bg="#f9f9f9", bd=1, relief="solid")

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

def select_excel():
    path = filedialog.askopenfilename(title="选择 Excel 文件", filetypes=[("Excel 文件", "*.xlsx *.xls")])
    if path:
        excel_path_var.set(path)
        for widget in scrollable_frame.winfo_children():
            widget.destroy()
        entry_vars.clear()
        entries.clear()
        for row_idx, label in extract_entries(path):
            var = tk.BooleanVar()
            chk = tk.Checkbutton(scrollable_frame, text=label, variable=var, anchor="w", justify="left",
                                 padx=10, font=font, width=60, bg="#f9f9f9", relief="flat",
                                 highlightthickness=0, bd=0)
            chk.pack(fill="x", anchor="w", pady=1)
            entry_vars.append(var)
            entries.append((row_idx, label))

def select_output_dir():
    path = filedialog.askdirectory(title="选择输出文件夹")
    if path:
        output_dir_var.set(path)

def run_fill():
    excel_path = excel_path_var.get()
    output_folder = output_dir_var.get()
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "template.docx")

    if not excel_path or not os.path.exists(excel_path):
        messagebox.showerror("错误", "请选择有效的 Excel 文件！")
        return
    if not output_folder or not os.path.isdir(output_folder):
        messagebox.showerror("错误", "请选择有效的保存目录！")
        return
    if not os.path.exists(template_path):
        messagebox.showerror("错误", "模板文件未找到！")
        return

    selected_rows = [entries[i][0] for i, var in enumerate(entry_vars) if var.get()]
    if not selected_rows:
        messagebox.showwarning("提示", "请选择要生成的条目")
        return

    try:
        fill_template_preserve_formatting(excel_path, template_path, output_folder, selected_rows)
        messagebox.showinfo("完成", "所选条目的收文处理笺已生成！")
    except Exception as e:
        messagebox.showerror("错误", str(e))

tk.Button(frame, text="生成选中条目", command=run_fill, bg="#4CAF50", fg="white", width=20).grid(row=3, column=0, columnspan=3, pady=15)

root.mainloop()
