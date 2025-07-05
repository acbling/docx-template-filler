import openpyxl
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# --------- 工具函数 ---------
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

# --------- 主要函数 ---------
def fill_template_preserve_formatting(excel_path, template_path, output_folder, selected_rows=None):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    os.makedirs(output_folder, exist_ok=True)

    for row_idx in range(5, ws.max_row + 1):
        if selected_rows and row_idx not in selected_rows:
            continue

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
        文件标题短 = 文件标题[:30] + ("…" if len(文件标题) > 30 else "")
        output_filename = f"{收文日期_str}党委组织部（党校）收文处理笺（{文件标题短}）.docx"
        output_path = os.path.join(output_folder, output_filename)

        doc.save(output_path)
        print(f"✅ 已生成 {output_path}")

# --------- 图形界面 ---------
class EntrySelectorApp:
    def __init__(self, master):
        self.master = master
        self.master.title("选择需要生成的收文处理条目")
        self.entries = []
        self.checkbox_vars = []

        self.excel_path = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel 文件", "*.xlsx")])
        if not self.excel_path:
            messagebox.showerror("错误", "未选择 Excel 文件")
            master.destroy()
            return

        self.template_path = filedialog.askopenfilename(title="选择模板 Word 文件", filetypes=[("Word 文件", "*.docx")])
        if not self.template_path:
            messagebox.showerror("错误", "未选择模板文件")
            master.destroy()
            return

        self.output_dir = filedialog.askdirectory(title="选择输出目录")

        self.extract_entries()

        self.scroll_canvas = tk.Canvas(master)
        self.frame = tk.Frame(self.scroll_canvas)
        self.scrollbar = tk.Scrollbar(master, orient="vertical", command=self.scroll_canvas.yview)
        self.scroll_canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side="right", fill="y")
        self.scroll_canvas.pack(side="left", fill="both", expand=True)
        self.scroll_canvas.create_window((0,0), window=self.frame, anchor="nw")
        self.frame.bind("<Configure>", lambda e: self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all")))

        for idx, label in enumerate(self.entries):
            var = tk.BooleanVar()
            cb = tk.Checkbutton(self.frame, text=label["label"], variable=var)
            cb.pack(anchor="w")
            self.checkbox_vars.append(var)

        tk.Button(master, text="生成 Word", command=self.generate_selected).pack(pady=10)

    def extract_entries(self):
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb.active
        for row_idx in range(5, ws.max_row + 1):
            title = str(ws.cell(row=row_idx, column=12).value or "无标题")
            date = ws.cell(row=row_idx, column=4).value
            date_str = format_excel_date(date)
            label = f"{date_str} | {title}"
            self.entries.append({
                "label": label,
                "row_idx": row_idx
            })

    def generate_selected(self):
        selected_indices = [e["row_idx"] for i, e in enumerate(self.entries) if self.checkbox_vars[i].get()]
        if not selected_indices:
            messagebox.showwarning("未选择", "请至少选择一条记录")
            return
        fill_template_preserve_formatting(
            self.excel_path, self.template_path, self.output_dir,
            selected_rows=selected_indices
        )
        messagebox.showinfo("完成", "已生成所选 Word 文件")

# --------- 启动程序 ---------
if __name__ == "__main__":
    root = tk.Tk()
    app = EntrySelectorApp(root)
    root.mainloop()
