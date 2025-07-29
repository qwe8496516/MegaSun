import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import main
import os
import threading
import sys

if sys.platform == 'win32':
    from pathlib import Path
    DOWNLOADS = str(Path.home() / 'Downloads')
else:
    DOWNLOADS = os.path.expanduser('~/Downloads')

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title('ERP 成本計算工具')
        self.root.geometry('800x400')
        self.root.resizable(False, False)

        self.input_path = tb.StringVar()
        self.output_dir = tb.StringVar(value=DOWNLOADS)
        self.output_file = 'ERP成本計算結果.xlsx'
        self.status_text = tb.StringVar(value='')

        tb.Label(root, text='選擇輸入 Excel 檔案：').pack(pady=(18, 0), anchor='w', padx=30)
        input_frame = tb.Frame(root)
        input_frame.pack(pady=5, padx=30, fill=X)
        tb.Entry(input_frame, textvariable=self.input_path).pack(side=LEFT, fill=X, expand=True)
        tb.Button(input_frame, text='瀏覽', bootstyle=SECONDARY, command=self.browse_input).pack(side=LEFT, padx=5)

        tb.Label(root, text='選擇輸出資料夾：').pack(pady=(12, 0), anchor='w', padx=30)
        output_frame = tb.Frame(root)
        output_frame.pack(pady=5, padx=30, fill=X)
        tb.Entry(output_frame, textvariable=self.output_dir).pack(side=LEFT, fill=X, expand=True)
        tb.Button(output_frame, text='瀏覽', bootstyle=SECONDARY, command=self.browse_output_dir).pack(side=LEFT, padx=5)
        tb.Label(root, text=f'輸出檔案名稱：{self.output_file}').pack(pady=(5, 0), anchor='w', padx=30)

        self.progress = tb.Progressbar(root, mode='indeterminate', length=320, bootstyle=INFO)

        self.status_label = tb.Label(root, textvariable=self.status_text, bootstyle=SECONDARY)
        self.status_label.pack(pady=(2, 0), padx=30, fill=X)

        self.run_btn = tb.Button(root, text='執行', width=15, bootstyle=SUCCESS, command=self.run_process_thread)
        self.run_btn.pack(pady=28)

    def browse_input(self):
        file_path = filedialog.askopenfilename(
            filetypes=[('Excel Files', '*.xlsx')],
            title='選擇輸入 Excel 檔案'
        )
        if file_path:
            self.input_path.set(file_path)

    def browse_output_dir(self):
        dir_path = filedialog.askdirectory(
            initialdir=DOWNLOADS,
            title='選擇輸出資料夾'
        )
        if dir_path:
            self.output_dir.set(dir_path)

    def run_process_thread(self):
        t = threading.Thread(target=self.run_process)
        t.start()

    def run_process(self):
        input_file = self.input_path.get()
        output_dir = self.output_dir.get()
        output_file = os.path.join(output_dir, self.output_file)
        if not input_file or not os.path.isfile(input_file):
            messagebox.showerror('錯誤', '請選擇正確的輸入 Excel 檔案！')
            return
        if not output_dir or not os.path.isdir(output_dir):
            messagebox.showerror('錯誤', '請選擇正確的輸出資料夾！')
            return

        if os.path.exists(output_file):
            ok = messagebox.askyesno('檔案已存在', f'檔案 {output_file} 已存在，是否要覆蓋？')
            if not ok:
                self.status_text.set('已取消執行')
                return
        self.status_text.set('處理中，請稍候...')
        self.progress.pack(pady=(18, 0), padx=30, fill=X)
        self.progress.start()
        self.run_btn.config(state='disabled')
        try:
            main.main(input_file, output_file)
            self.progress.stop()
            self.progress.pack_forget()
            self.status_text.set('處理完成！')
            self.run_btn.config(state='normal')
            messagebox.showinfo('完成', f'處理完成！\n輸出檔案：{output_file}')
        except Exception as e:
            self.progress.stop()
            self.progress.pack_forget()
            self.status_text.set('發生錯誤')
            self.run_btn.config(state='normal')
            messagebox.showerror('執行錯誤', str(e))

if __name__ == '__main__':
    app = tb.Window(themename='flatly')
    ExcelApp(app)
    app.mainloop() 