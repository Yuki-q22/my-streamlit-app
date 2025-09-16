import tkinter as tk
from tkinter import messagebox, scrolledtext
import subprocess
import os

class GitPushApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Git 一键更新工具")
        self.root.geometry("650x500")

        tk.Label(root, text="请选择要更新的文件", font=("微软雅黑", 16, "bold")).pack(pady=10)

        # 文件复选框区域
        self.files_frame = tk.Frame(root)
        self.files_frame.pack(pady=5, fill="x")
        self.file_vars = {}

        # 提交信息输入
        tk.Label(root, text="提交信息：", font=("微软雅黑", 12)).pack(pady=5)
        self.entry_msg = tk.Entry(root, width=60)
        self.entry_msg.insert(0, "update")
        self.entry_msg.pack(pady=5)

        # 按钮
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="刷新文件列表", command=self.refresh_file_list, bg="blue", fg="white", width=15).pack(side="left", padx=10)
        tk.Button(btn_frame, text="一键更新", command=self.do_git_ops, bg="green", fg="white", width=15).pack(side="left", padx=10)

        # 日志输出
        self.log_box = scrolledtext.ScrolledText(root, width=80, height=15)
        self.log_box.pack(pady=10)

        self.target_files = ["requirements.txt", "school_data.xlsx", "wangye.py", "招生专业.xlsx"]
        self.refresh_file_list()

    def run_git_command(self, cmd, cwd=None):
        """运行 git 命令，隐藏黑框"""
        try:
            result = subprocess.run(
                cmd,
                shell=True,
                cwd=cwd,
                text=True,
                capture_output=True,
                creationflags=subprocess.CREATE_NO_WINDOW  # 隐藏黑框
            )
            if result.stdout:
                self.log_box.insert("end", result.stdout + "\n")
            if result.stderr:
                self.log_box.insert("end", result.stderr + "\n")
            self.log_box.see("end")
            return result.returncode == 0
        except Exception as e:
            self.log_box.insert("end", f"执行出错：{e}\n")
            return False

    def refresh_file_list(self):
        """显示目标文件复选框，不论是否修改"""
        for widget in self.files_frame.winfo_children():
            widget.destroy()
        self.file_vars.clear()

        for f in self.target_files:
            var = tk.BooleanVar()
            chk = tk.Checkbutton(self.files_frame, text=f, variable=var)
            chk.pack(anchor="w", padx=50)
            self.file_vars[f] = var

    def do_git_ops(self):
        files_to_add = [f for f, var in self.file_vars.items() if var.get()]
        if not files_to_add:
            messagebox.showerror("错误", "请至少选择一个文件进行更新！")
            return

        commit_msg = self.entry_msg.get() or "update"
        repo_dir = os.getcwd()
        success = True

        # git add
        for f in files_to_add:
            success &= self.run_git_command(f'git add "{f}"', cwd=repo_dir)
        # git commit
        success &= self.run_git_command(f'git commit -m "{commit_msg}"', cwd=repo_dir)

        # git push，第一次自动设置 upstream
        push_success = self.run_git_command("git push", cwd=repo_dir)
        if not push_success:
            self.log_box.insert("end", "尝试第一次推送，设置 upstream...\n")
            push_success = self.run_git_command("git push -u origin main", cwd=repo_dir)

        success &= push_success

        if success:
            messagebox.showinfo("完成", "一键更新完成！")
            self.refresh_file_list()
        else:
            messagebox.showerror("失败", "操作过程中出错，请查看日志。")

if __name__ == "__main__":
    root = tk.Tk()
    app = GitPushApp(root)
    root.mainloop()
