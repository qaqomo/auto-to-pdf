import os
import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread

def convert_excel_to_pdf(folder_path, sheet_index, output_folder, update_callback):
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    supported_exts = (".xlsx", ".xls", ".xlsm", ".xlsb")
    files = [f for f in os.listdir(folder_path) if f.lower().endswith(supported_exts)]

    for filename in files:
        filepath = os.path.join(folder_path, filename)
        try:
            wb = excel.Workbooks.Open(filepath)
            if sheet_index > wb.Sheets.Count:
                wb.Close(False)
                continue

            sheet = wb.Sheets(sheet_index)
            sheet.Select()

            file_basename = os.path.splitext(filename)[0]
            pdf_path = os.path.join(output_folder, f"{file_basename}.pdf")
            sheet.ExportAsFixedFormat(0, pdf_path)
            wb.Close(False)
        except Exception as e:
            print(f"Excel转换失败: {filename} - {e}")
        update_callback()

    excel.Quit()

def convert_word_to_pdf(folder_path, output_folder, update_callback):
    word = win32com.client.gencache.EnsureDispatch('Word.Application')
    word.Visible = False

    supported_exts = (".doc", ".docx")
    files = [f for f in os.listdir(folder_path) if f.lower().endswith(supported_exts)]

    for filename in files:
        # 拼接绝对路径（不要改动文件名）
        filepath = os.path.join(folder_path, filename)
        try:
            doc = word.Documents.Open(filepath)  # 关键是这里用的 filepath 是原始路径
            file_basename = os.path.splitext(filename)[0]
            output_path = os.path.join(output_folder, file_basename + '.pdf')
            doc.SaveAs(output_path, FileFormat=17)  # 17 = wdFormatPDF
            doc.Close()
        except Exception as e:
            print(f"Word转换失败: {filename} - {e}")
        update_callback()

    word.Quit()


def start_conversion_thread():
    thread = Thread(target=start_conversion)
    thread.start()

def start_conversion():
    folder = folder_var.get()
    output = output_var.get() or folder
    try:
        sheet = int(sheet_var.get())
    except ValueError:
        messagebox.showerror("错误", "工作表索引必须是整数")
        return

    if not os.path.isdir(folder):
        messagebox.showerror("错误", "请选择有效的源文件夹")
        return

    os.makedirs(output, exist_ok=True)

    excel_files = [f for f in os.listdir(folder) if f.lower().endswith((".xlsx", ".xls", ".xlsm", ".xlsb"))]
    word_files = [f for f in os.listdir(folder) if f.lower().endswith((".doc", ".docx"))]
    total_files = len(excel_files) + len(word_files)

    if total_files == 0:
        messagebox.showinfo("提示", "文件夹中没有可转换的Word或Excel文件")
        return

    progress_bar["value"] = 0
    progress_bar.pack(pady=10)
    check_canvas.pack_forget()
    count = {"done": 0}

    def update_progress():
        count["done"] += 1
        percent = int((count["done"] / total_files) * 100)
        progress_bar["value"] = percent
        app.update_idletasks()
        if count["done"] == total_files:
            progress_bar.pack_forget()
            check_canvas.pack(pady=10)

    convert_excel_to_pdf(folder, sheet, output, update_progress)
    convert_word_to_pdf(folder, output, update_progress)

# ===== GUI 界面设计 =====
app = tk.Tk()
app.title("Word / Excel 批量转 PDF 工具")
app.geometry("500x340")

folder_var = tk.StringVar()
output_var = tk.StringVar()
sheet_var = tk.StringVar(value="1")

tk.Label(app, text="选择源文件夹:").pack(anchor="w", padx=10, pady=5)
tk.Entry(app, textvariable=folder_var, width=60).pack(padx=10)
tk.Button(app, text="浏览", command=lambda: folder_var.set(filedialog.askdirectory())).pack(pady=5)

tk.Label(app, text="输出PDF文件夹 (可选，默认源文件夹):").pack(anchor="w", padx=10)
tk.Entry(app, textvariable=output_var, width=60).pack(padx=10)
tk.Button(app, text="浏览", command=lambda: output_var.set(filedialog.askdirectory())).pack(pady=5)

tk.Label(app, text="Excel工作表索引 (默认1):").pack(anchor="w", padx=10, pady=5)
tk.Entry(app, textvariable=sheet_var, width=10).pack(padx=10)

progress_bar = ttk.Progressbar(app, length=400, mode='determinate')

check_canvas = tk.Canvas(app, width=40, height=40, bg='white', highlightthickness=0)
check_canvas.create_oval(5, 5, 35, 35, fill='#d4edda', outline='#c3e6cb')
check_canvas.create_line(12, 22, 18, 28, fill='green', width=3)
check_canvas.create_line(18, 28, 30, 14, fill='green', width=3)

tk.Button(app, text="开始转换", command=start_conversion_thread).pack(pady=10)

app.mainloop()
