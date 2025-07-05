import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os


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


def center_align_table_rows(table, row_indices):
    for row_idx in row_indices:
        row = table.rows[row_idx]
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("收文处理笺生成器")

        self.excel_path = ""
        self.template_path = os.path.abspath("template.docx")
        self.entries = []
        self.checkbox_vars = []

        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="选择 Excel 文件：").grid(row=0, column=0, sticky=tk.W)
        self.excel_entry = ttk.Entry(frame, width=40)
        self.excel_entry.grid(row=0, column=1)
        ttk.Button(frame, text="浏览", command=self.select_excel).grid(row=0, column=2)

        self.check_frame = ttk.LabelFrame(frame, text="选择要生成的条目：")
        self.check_frame.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=10)

        self.check_canvas = tk.Canvas(self.check_frame, height=200)
        self.check_scrollbar = ttk.Scrollbar(self.check_frame, orient="vertical", command=self.check_canvas.yview)
        self.inner_frame = ttk.Frame(self.check_canvas)

        self.inner_frame.bind("<Configure>", lambda e: self.check_canvas.configure(scrollregion=self.check_canvas.bbox("all")))
        self.check_canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        self.check_canvas.configure(yscrollcommand=self.check_scrollbar.set)

        self.check_canvas.pack(side="left", fill="both", expand=True)
        self.check_scrollbar.pack(side="right", fill="y")

        ttk.Button(frame, text="生成 Word 文件", command=self.generate_docs).grid(row=2, column=1, pady=10)

    def select_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.excel_path = path
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, path)
            self.load_entries()

    def load_entries(self):
        self.entries.clear()
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
        self.checkbox_vars.clear()

        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb.active
        for row_idx in range(5, ws.max_row + 1):
            title = str(ws.cell(row=row_idx, column=12).value or "").strip()
            date_cell = ws.cell(row=row_idx, column=4).value

            if not title and not date_cell:
                continue  # 跳过空行

            try:
                if isinstance(date_cell, datetime):
                    date_str = date_cell.strftime("%Y-%m-%d")
                else:
                    date_str = datetime.fromordinal(int(date_cell) + 693594).strftime("%Y-%m-%d")
            except:
                date_str = "日期未知"

            label = f"{date_str} - {title[:30]}"
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(self.inner_frame, text=label, variable=var)
            chk.pack(anchor="w", pady=2)
            self.entries.append(row_idx)
            self.checkbox_vars.append(var)

    def generate_docs(self):
        selected_rows = [row for row, var in zip(self.entries, self.checkbox_vars) if var.get()]
        if not selected_rows:
            messagebox.showwarning("提示", "请至少选择一条要生成的记录")
            return
        output_dir = filedialog.askdirectory()
        if not output_dir:
            return
        try:
            self.generate_docs_from_rows(selected_rows, output_dir)
            messagebox.showinfo("完成", "Word 文件已生成！")
        except Exception as e:
            messagebox.showerror("错误", f"生成失败：{e}")

    def generate_docs_from_rows(self, row_indices, output_folder):
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb.active
        os.makedirs(output_folder, exist_ok=True)

        for row_idx in row_indices:
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

            doc = Document(self.template_path)
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
            try:
                if isinstance(收文日期, datetime):
                    收文日期_str = 收文日期.strftime('%Y%m%d')
                else:
                    收文日期_str = datetime.fromordinal(int(收文日期) + 693594).strftime('%Y%m%d')
            except:
                收文日期_str = "日期未知"

            文件标题 = str(data['文件标题']) if data['文件标题'] else "无标题"
            文件标题短 = 文件标题[:30] + ("…" if len(文件标题) > 30 else "")

            filename = f"{收文日期_str}党委组织部收文处理笺（{文件标题短}）.docx"
            doc.save(os.path.join(output_folder, filename))


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
