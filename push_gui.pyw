import tkinter as tk
from tkinter import messagebox, scrolledtext
import subprocess
import os
import socket

class GitPushApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Git 更新工具")
        self.root.geometry("720x720")
        self.root.configure(bg="#f0f0f0")  # 背景色

        # 顶部标题
        tk.Label(root, text="请选择要更新的文件", font=("微软雅黑", 16, "bold"), bg="#f0f0f0").pack(pady=10)

        # 文件复选框区域，带边框
        self.files_frame = tk.LabelFrame(root, text="文件列表", padx=20, pady=10, bg="#ffffff", font=("微软雅黑", 12))
        self.files_frame.pack(padx=20, pady=10, fill="x")
        self.file_vars = {}

        # 全选复选框
        self.select_all_var = tk.BooleanVar()
        select_all_chk = tk.Checkbutton(
            self.files_frame,
            text="全选",
            variable=self.select_all_var,
            command=self.toggle_all,
            bg="#ffffff",
            font=("微软雅黑", 14, "bold")
        )
        select_all_chk.pack(anchor="w", pady=5)

        # 目标文件复选框
        self.target_files = ["requirements.txt", "school_data.xlsx", "wangye.py", "招生专业.xlsx"]

        for f in self.target_files:
            var = tk.BooleanVar()
            chk = tk.Checkbutton(
                self.files_frame,
                text=f,
                variable=var,
                bg="#ffffff",
                font=("微软雅黑", 13)
            )
            chk.pack(anchor="w", padx=20, pady=3)
            self.file_vars[f] = var

        # 提交信息输入
        tk.Label(root, text="提交信息：", font=("微软雅黑", 12), bg="#f0f0f0").pack(pady=5)
        self.entry_msg = tk.Entry(root, width=50, font=("微软雅黑", 12))
        self.entry_msg.insert(0, "update")
        self.entry_msg.pack(pady=5)

        # 按钮
        btn_frame = tk.Frame(root, bg="#f0f0f0")
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="刷新文件列表", command=self.refresh_file_list, bg="#4a90e2", fg="white",
                  width=15, font=("微软雅黑", 12)).pack(side="left", padx=10)
        tk.Button(btn_frame, text="更新", command=self.do_git_ops, bg="#50c878", fg="white",
                  width=15, font=("微软雅黑", 12)).pack(side="left", padx=10)

        # 日志输出
        tk.Label(root, text="更新日志：", font=("微软雅黑", 12), bg="#f0f0f0").pack()
        self.log_box = scrolledtext.ScrolledText(root, width=85, height=18, font=("Consolas", 11), bg="#f7f7f7")
        self.log_box.pack(pady=5, padx=20)

    def log(self, text):
        """日志打印"""
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.root.update()

    def toggle_all(self):
        state = self.select_all_var.get()
        for var in self.file_vars.values():
            var.set(state)

    def run_git_command(self, cmd, cwd=None):
        try:
            result = subprocess.run(
                cmd,
                shell=True,
                cwd=cwd,
                text=True,
                capture_output=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            output = ""
            if result.stdout:
                output += result.stdout
            if result.stderr:
                output += result.stderr
            if output:
                self.log(output.strip())
            return result.returncode == 0
        except Exception as e:
            self.log(f"执行出错：{e}")
            return False

    def refresh_file_list(self):
        for f, var in self.file_vars.items():
            var.set(False)
        self.select_all_var.set(False)

    # ============ 新增：检查代理并自动取消 ============
    def check_proxy(self):
        proxy = subprocess.getoutput("git config --global --get https.proxy")
        if proxy and proxy.startswith("http://127.0.0.1"):
            host, port = proxy.replace("http://", "").split(":")
            self.log(f"检测到本地代理: {proxy}，检查端口是否可用...")
            if not self.test_port(host, int(port)):
                self.log("代理端口不可用，自动取消代理设置")
                self.run_git_command("git config --global --unset http.proxy")
                self.run_git_command("git config --global --unset https.proxy")
        else:
            self.log("未检测到本地代理或代理正常")

    def test_port(self, host, port):
        try:
            with socket.create_connection((host, port), timeout=2):
                return True
        except Exception:
            return False

    # ============ 新增：自动提交未暂存修改 ============
    def commit_unstaged_changes(self):
        status = subprocess.getoutput("git status --porcelain")
        if status.strip():
            self.log("检测到未暂存的更改，自动提交...")
            self.run_git_command("git add .")
            self.run_git_command('git commit -m "auto-commit before pull"')
        else:
            self.log("没有未暂存的更改")

    def do_git_ops(self):
        files_to_add = [f for f, var in self.file_vars.items() if var.get()]
        if not files_to_add:
            messagebox.showerror("错误", "请至少选择一个文件进行更新！")
            return

        commit_msg = self.entry_msg.get() or "update"
        repo_dir = os.getcwd()
        self.log(f"准备更新文件：{', '.join(files_to_add)}")

        success = True

        # 0. 检查代理
        self.check_proxy()

        # 0.5 自动提交未暂存修改，避免 pull 失败
        self.commit_unstaged_changes()

        # 1. 拉取远程最新，避免冲突
        self.log("拉取远程最新...")
        pull_ok = self.run_git_command("git pull --rebase origin main", cwd=repo_dir)
        if not pull_ok:
            self.log("拉取远程失败，请检查网络或冲突。")

        # 2. 添加选中文件
        for f in files_to_add:
            success &= self.run_git_command(f'git add "{f}"', cwd=repo_dir)

        # 3. 检查是否有实际改动
        result = subprocess.run(
            "git diff --cached --name-only",
            shell=True,
            cwd=repo_dir,
            text=True,
            capture_output=True
        )
        if not result.stdout.strip():
            self.log("没有新改动需要提交。")
        else:
            # 4. 提交改动
            success &= self.run_git_command(f'git commit -m "{commit_msg}"', cwd=repo_dir)

            # 5. 推送到远程
            push_success = self.run_git_command("git push", cwd=repo_dir)
            if not push_success:
                self.log("尝试第一次推送，设置 upstream...")
                push_success = self.run_git_command("git push -u origin main", cwd=repo_dir)
            success &= push_success

        if success:
            self.log("更新完成！\n---------------------")
            messagebox.showinfo("完成", "更新完成！")
            self.refresh_file_list()
        else:
            self.log("操作失败，请查看日志")
            messagebox.showerror("失败", "操作过程中出错，请查看日志。")

if __name__ == "__main__":
    root = tk.Tk()
    app = GitPushApp(root)
    root.mainloop()
