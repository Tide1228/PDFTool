import tkinter as tk
from tkinter import filedialog, messagebox
from pdf2docx import Converter
import ctypes
from PyPDF2 import PdfMerger


class PDFTool:
    def __init__(self, root):
        # 解决字体模糊问题
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # 启用DPI感知

        self.root = root
        self.root.title("PDF Tool")
        self.root.geometry("600x400")
        self.root.resizable(False, False)# 禁止调整窗口大小
        self.center_window()  # 窗口居中

        # 创建界面元素
        self.label = tk.Label(root, text="PDF Tool", font=("Arial", 20))
        self.label.grid(row=0, column=0, columnspan=2, pady=10)

        # 创建一个Frame来放置按钮
        self.button_frame = tk.Frame(root)
        self.button_frame.grid(row=1, column=0, padx=30, pady=10, sticky="n")

        self.select_pdf_button = tk.Button(
            self.button_frame,
            text="Select PDF",
            font=("Arial", 12),
            command=self.select_pdf,
            width=15,
            height=1
        )
        self.select_pdf_button.pack(pady=10)

        self.convert_button = tk.Button(
            self.button_frame,
            text="Convert to Word",
            font=("Arial", 12),
            command=self.convert_pdf_to_word,
            state=tk.DISABLED,
            width=15,
            height=1
        )
        self.convert_button.pack(pady=10)

        self.merge_button = tk.Button(
            self.button_frame,
            text="Merge PDFs",
            font=("Arial", 12),
            command=self.merge_pdfs,
            state=tk.DISABLED,
            width=15,
            height=1
        )
        self.merge_button.pack(pady=10)

        # 创建一个Listbox来显示已选择的文件
        self.file_listbox = tk.Listbox(root, width=30, height=12, font=("Arial", 16))
        self.file_listbox.grid(row=1, column=1, padx=20, pady=10, sticky="n")

        # 绑定鼠标事件
        self.file_listbox.bind("<ButtonPress-1>", self.on_start_drag)
        self.file_listbox.bind("<B1-Motion>", self.on_drag)
        self.file_listbox.bind("<ButtonRelease-1>", self.on_drop)
        # 绑定右键点击事件
        self.file_listbox.bind("<Button-3>", self.show_context_menu)

        # 创建上下文菜单
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Delete", command=self.delete_selected_file)

        self.pdf_path = []
        self.dragged_index = None

    def center_window(self):
        # 强制更新窗口的几何信息
        self.root.update_idletasks()

        # 获取屏幕的宽度和高度
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 获取窗口的宽度和高度
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()

        # 计算居中位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # 设置窗口位置
        self.root.geometry(f"+{x + 200}+{y + 50}")

    def select_pdf(self):
        self.pdf_path = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        self.pdf_path = list(self.pdf_path)
        if self.pdf_path:
            self.convert_button.config(state=tk.NORMAL)
            self.merge_button.config(state=tk.NORMAL)
            self.update_file_listbox()  # 更新文件列表显示
            messagebox.showinfo("File Selected", f"Selected PDF: {self.pdf_path}")


    def update_file_listbox(self):
        self.file_listbox.delete(0, tk.END)  # 清空当前列表
        for file_path in self.pdf_path:
            self.file_listbox.insert(tk.END, file_path.rsplit('/', 1)[-1])  # 插入新的文件路径

    def convert_pdf_to_word(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "No PDF file selected!")
            return

        for pdf_file in self.pdf_path:
            try:
                cv = Converter(pdf_file)
                cv.convert(start=0, end=None)
                cv.close()
                messagebox.showinfo("Success", f"{pdf_file} converted to Word successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")

    def merge_pdfs(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "No PDF files selected!")
            return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_path:
            try:
                # 创建一个PdfMerger对象
                merger = PdfMerger()

                # 遍历所有PDF文件并添加到merger中
                for pdf_file in self.pdf_path:
                    merger.append(pdf_file)

                # 合并所有PDF文件并保存到输出路径
                merger.write(output_path)
                merger.close()

                messagebox.showinfo("Success", f"PDF files have been successfully merged to {output_path}")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")

    def on_start_drag(self, event):
        index = self.file_listbox.nearest(event.y)
        self.dragged_index = index
        self.file_listbox.selection_clear(0, tk.END)
        self.file_listbox.selection_set(index)

    def on_drag(self, event):
        if self.dragged_index is not None:
            target_index = self.file_listbox.nearest(event.y)
            if target_index != self.dragged_index:
                # 交换列表项的位置
                item = self.pdf_path.pop(self.dragged_index)
                self.pdf_path.insert(target_index, item)
                self.update_file_listbox()
                self.dragged_index = target_index
                self.file_listbox.selection_clear(0, tk.END)
                self.file_listbox.selection_set(target_index)

    def on_drop(self, event):
        self.dragged_index = None

    def show_context_menu(self, event):
        index = self.file_listbox.nearest(event.y)
        self.file_listbox.selection_clear(0, tk.END)
        self.file_listbox.selection_set(index)
        self.context_menu.post(event.x_root, event.y_root)

    def delete_selected_file(self):
        selected_index = self.file_listbox.curselection()
        if selected_index:
            index = selected_index[0]
            del self.pdf_path[index]
            self.update_file_listbox()
            if not self.pdf_path:
                self.convert_button.config(state=tk.DISABLED)
                self.merge_button.config(state=tk.DISABLED)


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFTool(root)
    root.mainloop()
