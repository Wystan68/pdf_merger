import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Scrollbar
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
import threading
from PIL import Image
from PyPDF2 import PdfMerger
from docx2pdf import convert
import tempfile
import re
import shutil

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文件转PDF合并工具 (支持图片/Word/PDF)")
        self.root.geometry("600x500")

        # 存储文件路径的列表
        self.file_list = []

        # UI 布局
        self.create_widgets()

    def create_widgets(self):
        # 顶部说明
        lbl_instruction = tk.Label(self.root, text="请将文件 (PDF, Word, 图片) 拖入下方列表，\n文件将按列表顺序合并。", pady=10)
        lbl_instruction.pack()

        # 中间列表框框架
        frame_list = tk.Frame(self.root)
        frame_list.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 滚动条
        scrollbar = Scrollbar(frame_list)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 列表框 (用于显示拖入的文件)
        self.listbox = Listbox(frame_list, selectmode=tk.SINGLE, yscrollcommand=scrollbar.set, bg="#f0f0f0", font=("Arial", 10))
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)

        # 启用拖拽功能
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self.drop_files)

        # 底部按钮区域
        frame_btns = tk.Frame(self.root)
        frame_btns.pack(fill=tk.X, padx=10, pady=10)

        btn_clear = tk.Button(frame_btns, text="清空列表", command=self.clear_list, width=15)
        btn_clear.pack(side=tk.LEFT)

        btn_remove = tk.Button(frame_btns, text="移除选中项", command=self.remove_selected, width=15)
        btn_remove.pack(side=tk.LEFT, padx=10)

        self.btn_convert = tk.Button(frame_btns, text="开始转换并合并", command=self.start_conversion_thread, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
        self.btn_convert.pack(side=tk.RIGHT)

        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        lbl_status = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        lbl_status.pack(side=tk.BOTTOM, fill=tk.X)

    def parse_drop_files(self, event_data):
        """
        解析拖拽进来的文件路径。
        Windows下路径如果有空格，会被大括号包围，例如 {C:/My Files/test.pdf}
        """
        # 正则表达式匹配：要么在大括号内的内容，要么是非空白字符序列
        pattern = r'\{.*?\راع|\S+'
        files = re.findall(pattern, event_data)
        cleaned_files = []
        for f in files:
            # 去除大括号
            f = f.strip('{}')
            if os.path.isfile(f):
                cleaned_files.append(f)
        return cleaned_files

    def drop_files(self, event):
        files = self.parse_drop_files(event.data)
        for f in files:
            if f not in self.file_list:
                self.file_list.append(f)
                self.listbox.insert(tk.END, f)
                
    def clear_list(self):
        self.file_list = []
        self.listbox.delete(0, tk.END)
        self.status_var.set("列表已清空")

    def remove_selected(self):
        selected_indices = self.listbox.curselection()
        if not selected_indices:
            return
        
        # 从后往前删，避免索引错位
        for index in reversed(selected_indices):
            self.listbox.delete(index)
            del self.file_list[index]

    def start_conversion_thread(self):
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加文件！")
            return

        # 获取保存路径
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not output_path:
            return

        # 禁用按钮防止重复点击
        self.btn_convert.config(state=tk.DISABLED)
        self.status_var.set("正在处理中，请稍候...")
        
        # 开启新线程处理，避免界面卡死
        thread = threading.Thread(target=self.process_files, args=(output_path,))
        thread.start()

    def process_files(self, output_path):
        temp_dir = tempfile.mkdtemp() # 创建临时文件夹
        merger = PdfMerger()
        temp_pdf_files = [] # 记录生成的临时PDF，以便最后合并和清理

        try:
            for idx, file_path in enumerate(self.file_list):
                ext = os.path.splitext(file_path)[1].lower()
                filename = os.path.basename(file_path)
                self.status_var.set(f"正在处理 ({idx+1}/{len(self.file_list)}): {filename}")
                
                temp_pdf_path = os.path.join(temp_dir, f"temp_{idx}.pdf")

                if ext in ['.pdf']:
                    # 如果是PDF，直接用于合并
                    merger.append(file_path)
                
                elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']:
                    # 图片转PDF
                    try:
                        image = Image.open(file_path)
                        # 转换模式，除去透明度（RGBA转RGB），否则无法保存为PDF
                        if image.mode in ("RGBA", "P"): 
                            image = image.convert("RGB")
                        image.save(temp_pdf_path, "PDF", resolution=100.0)
                        merger.append(temp_pdf_path)
                        temp_pdf_files.append(temp_pdf_path)
                    except Exception as e:
                        print(f"图片转换错误: {e}")
                
                elif ext in ['.docx', '.doc']:
                    # Word转PDF
                    try:
                        # docx2pdf 需要完整路径
                        convert(file_path, temp_pdf_path)
                        merger.append(temp_pdf_path)
                        temp_pdf_files.append(temp_pdf_path)
                    except Exception as e:
                        print(f"Word转换错误 (请确保已安装MS Word): {e}")
                        messagebox.showerror("错误", f"无法转换 Word 文件: {filename}\n请确保已安装 Microsoft Word。")
                        return # 停止处理

            # 合并所有文件
            self.status_var.set("正在合并最终文件...")
            merger.write(output_path)
            merger.close()
            
            self.status_var.set("完成！")
            messagebox.showinfo("成功", f"文件已成功合并保存至:\n{output_path}")

        except Exception as e:
            self.status_var.set("发生错误")
            messagebox.showerror("错误", f"处理过程中发生错误:\n{str(e)}")
        
        finally:
            # 清理工作
            self.btn_convert.config(state=tk.NORMAL)
            # 删除临时文件夹及其中内容
            try:
                shutil.rmtree(temp_dir)
            except:
                pass

if __name__ == "__main__":
    # 使用 TkinterDnD.Tk 而不是标准的 tk.Tk
    root = TkinterDnD.Tk()
    app = PDFMergerApp(root)
    root.mainloop()