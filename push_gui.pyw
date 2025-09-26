import tkinter as tk
from tkinter import messagebox, scrolledtext
import subprocess
import os


class GitPushApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Git 更新工具")
        self.root.geometry("800x600")
        self.root.configure(bg="#f0f0f0")

        # 顶部标题
        tk.Label(
            root, text="一键推送工具", font=("微软雅黑", 16, "bold"),
            bg="#f0f0f0", fg="#333"
        ).pack(pady=10)

        # 文件选择区域
        frame_files = tk.LabelFrame(root, text="请选择要更新的文件", font=("微软雅黑", 12),
                                    bg="#f0f0f0", padx=10, pady=10)
        frame_files.pack(fill="x", padx=20, pady=10)

        self.file_vars = {}
        self.files = ["requirements.txt", "school_data.xlsx", "wangye.py", "招生专业.xlsx"]

        for f in self.files:
            var = tk.BooleanVar()
            chk = tk.Checkbutton(frame_files, text=f, variable=var,
                                 font=("微软雅黑", 11), bg="#f0f0f0")
            chk.pack(anchor="w", pady=2)
            self.file_vars[f] = var

        # 全选按钮
        self.select_all_var = tk.BooleanVar()
        chk_all = tk.Checkbutton(frame_files, text="全选", variable=self.select_all_var,
                                 command=self.toggle_select_all,
                                 font=("微软雅黑", 11, "bold"), bg="#f0f0f0", fg="#007acc")
        chk_all.pack(anchor="w", pady=5)

        # 操作按钮
        frame_buttons = tk.Frame(root, bg="#f0f0f0")
        frame_buttons.pack(pady=10)

        tk.Button(frame_buttons, text="更新", command=self.do_git_ops,
                  font=("微软雅黑", 12), width=12, bg="#4caf50", fg="white").grid(row=0, column=0, padx=10)

        tk.Button(frame_buttons, text="退出", command=root.quit,
                  font=("微软雅黑", 12), width=12, bg="#f44336", fg="white").grid(row=0, column=1, padx=10)

        # 日志输出
        self.log_box = scrolledtext.ScrolledText(root, width=100, height=20,
                                                 font=("Consolas", 10))
        self.log_box.pack(padx=20, pady=10)

    def log(self, msg):
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")

    def toggle_select_all(self):
        state = self.select_all_var.get()
        for var in self.file_vars.values():
            var.set(state)

    def run_cmd(self, cmd, cwd=None):
        result = subprocess.run(cmd, shell=True, cwd=cwd,
                                text=True, capture_output=True)
        if result.stdout:
            self.log(result.stdout.strip())
        if result.stderr:
            self.log(result.stderr.strip())
        return result

    def ensure_ssh_remote(self, repo_dir):
        """确保远程仓库使用 SSH"""
        try:
            result = subprocess.run(
                "git remote get-url origin",
                shell=True, cwd=repo_dir, text=True, capture_output=True
            )
            current_url = result.stdout.strip()
            ssh_url = "git@github.com:Yuki-q22/my-streamlit-app.git"
            if current_url != ssh_url:
                self.log(f"检测到远程仓库为 {current_url}，切换为 SSH...")
                subprocess.run(
                    f"git remote set-url origin {ssh_url}",
                    shell=True, cwd=repo_dir
                )
                self.log(f"已切换到 SSH: {ssh_url}")
        except Exception as e:
            self.log(f"检测/修改远程仓库失败：{e}")

    def handle_unstaged_changes(self, repo_dir):
        """处理未提交的更改"""
        status = subprocess.run("git status --porcelain", shell=True,
                                cwd=repo_dir, text=True, capture_output=True)
        if status.stdout.strip():
            choice = messagebox.askquestion(
                "检测到未提交更改",
                "本地有未提交的修改，选择操作：\n\n"
                "是 = 提交\n"
                "否 = 暂存\n"
                "取消 = 丢弃\n\n"
                "如果要取消推送，请关闭弹窗。"
            )
            if choice == "yes":  # 提交
                self.run_cmd("git add .", cwd=repo_dir)
                self.run_cmd('git commit -m "auto-commit before pull"', cwd=repo_dir)
                return True
            elif choice == "no":  # 暂存
                self.run_cmd("git stash", cwd=repo_dir)
                return "stash"
            else:  # 丢弃
                self.run_cmd("git reset --hard", cwd=repo_dir)
                return True
        return True

    def do_git_ops(self):
        repo_dir = os.getcwd()

        # 确保远程仓库走 SSH
        self.ensure_ssh_remote(repo_dir)

        # 检查文件选择
        files_to_add = [f for f, var in self.file_vars.items() if var.get()]
        if not files_to_add:
            messagebox.showerror("错误", "请至少选择一个文件进行更新！")
            return

        # 处理未提交改动
        result = self.handle_unstaged_changes(repo_dir)
        if not result:
            self.log("操作取消。")
            return

        self.log("准备更新文件：" + "、".join(files_to_add))

        # 添加文件
        for f in files_to_add:
            self.run_cmd(f"git add {f}", cwd=repo_dir)

        # 提交
        self.run_cmd('git commit -m "update"', cwd=repo_dir)

        # 拉取远程
        self.log("拉取远程最新...")
        pull_result = self.run_cmd("git pull --rebase origin main", cwd=repo_dir)

        if "error:" in pull_result.stderr:
            self.log("拉取远程失败，请检查网络或冲突。")
            return

        # 推送
        self.log("推送到远程...")
        push_result = self.run_cmd("git push origin main", cwd=repo_dir)
        if "fatal:" in push_result.stderr:
            self.log("推送失败，请检查 SSH 设置。")
        else:
            self.log("推送完成！")


if __name__ == "__main__":
    root = tk.Tk()
    app = GitPushApp(root)
    root.mainloop()
