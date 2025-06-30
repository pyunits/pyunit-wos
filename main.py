import os
import tkinter as tk
from threading import Thread
from tkinter import messagebox, ttk

from wos import WOS, Status


class WOSApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WOS 数据下载工具")
        self.root.geometry("350x250")
        self.root.resizable(False, False)
        self.wos = None

        # 创建输入框和标签
        tk.Label(root, text="权限SID:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        self.sid_entry = tk.Entry(root, width=30)
        self.sid_entry.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(root, text="URL QID:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        self.qid_entry = tk.Entry(root, width=30)
        self.qid_entry.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(root, text="数据大小:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        self.size_var = tk.IntVar(value=0)
        self.size_spinbox = tk.Spinbox(root, from_=0, to=100_0000, textvariable=self.size_var, width=28)
        self.size_spinbox.grid(row=2, column=1, padx=10, pady=10)

        self.progress_label = tk.Label(root, text="下载进度:")
        self.progress_label.grid(row=3, column=0, padx=10, pady=10, sticky=tk.W)
        self.progress = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=215,
                                        style="default.Horizontal.TProgressbar")
        self.progress.grid(row=3, column=1)

        # 下载按钮
        self.button = tk.Button(root, text="开始下载", command=self.onclick, width=20, height=2)
        self.button.grid(row=4, column=0, columnspan=2, pady=10)

    def __download__(self):
        """获取用户输入并开始下载"""
        sid = self.sid_entry.get().strip()
        qid = self.qid_entry.get().strip()
        size = self.size_var.get().real

        if not sid or not qid or not size:
            messagebox.showerror("错误", "所有字段都必须填写！")
            self.button.config(state="normal")
            return
        try:
            self.progress["maximum"] = size
            dirs = os.path.abspath(".")
            file = os.path.join(dirs, f"{qid}.xlsx")
            self.wos = WOS(sid=sid, qid=qid, savefile=file)
            self.wos.run(size)
            self.wos.save()
            messagebox.showinfo("成功", f"文件已保存到：{file}")
        except Exception as e:
            messagebox.showerror("错误", f"下载失败：{str(e)}")
        finally:
            self.button.config(state="normal")

    def onclick(self):
        self.button.config(state="disabled")
        thread = Thread(target=self.__download__, daemon=True)
        thread.start()
        self.root.after(200, self.progress_bar, thread)

    def progress_bar(self, thread):
        if thread.is_alive():
            if self.wos is None:
                return

            code = self.wos.code
            if code == Status.ERROR:
                messagebox.showerror("错误", "输入错误或网络错误")
                self.button.config(state="normal")
                return

            if code == Status.LIMIT:
                self.progress_label.config(text=f"刷新网站~")
                self.progress.config(style="red.Horizontal.TProgressbar")
                self.root.after(8000, self.progress_bar, thread)
                return

            count = self.wos.size
            self.progress["value"] = count
            process = round(count / self.size_var.get().real * 100)
            self.progress_label.config(text=f"下载中: {process}%")
            self.progress.config(style="default.Horizontal.TProgressbar")
            self.root.after(3000, self.progress_bar, thread)


if __name__ == '__main__':
    ui = tk.Tk()
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("red.Horizontal.TProgressbar", background="red")
    style.configure("default.Horizontal.TProgressbar", background="blue")
    WOSApp(ui)
    ui.mainloop()
