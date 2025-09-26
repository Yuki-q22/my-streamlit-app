import tkinter as tk
from tkinter import messagebox, scrolledtext, simpledialog
import subprocess
import os


class GitPushApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Git 更新工具")
        self.root.geometry("720x550")
        self.root.configure(bg="#f0f0f0")  # 背景色

        # 顶部标题
        tk.Label(root, text="Git 更新工具", font=("Arial", 18, "bold"), bg="#f0f0f0").pack(pady=10)

        # 输出框
        self.output = scrolledtext.ScrolledText(root, width=80, height=25, wrap=tk.WORD, state="disabled")
        self.output.pack(padx=10, pady=10)

        # 按钮区域
        btn_frame = tk.Frame(root, bg="#f0f0f0")
        btn_frame.pack()

        tk.Button(btn_frame, text="更新 (Pull)", command=self.update_repo, width=15).grid(row=0, column=0, padx=10)
        tk.Button(btn_frame, text="推送 (Push)", command=self.push_repo, width=15).grid(row=0, column=1, padx=10)

    def run_command(self, cmd):
        """运行命令并返回输出"""
        try:
            result = subprocess.run(cmd, shell=True, text=True, capture_output=True, encoding="utf-8")
            return result.stdout + result.stderr
        except Exception as e:
            return str(e)

    def append_output(self, text):
        """在输出框中追加文本"""
        self.output.configure(state="normal")
        self.output.insert(tk.END, text + "\n")
        self.output.see(tk.END)
        self.output.configure(state="disabled")

    def check_repo_mode(self):
        """检查是否使用 SSH，如果是 HTTPS 就切换"""
        remote_url = self.run_command("git remote get-url origin").strip()
        if remote_url.startswith("https://"):
            self.append_output(f"检测到远程仓库为 {remote_url}，切换为 SSH...")
            ssh_url = remote_url.replace("https://github.com/", "git@github.com:")
            self.run_command(f"git remote set-url origin {ssh_url}")
            self.append_output(f"已切换到 SSH: {ssh_url}")

    def handle_local_changes(self):
        """检测并处理本地未提交的改动"""
        status = self.run_command("git status --porcelain").strip()
        if status:  # 有改动
            choice = messagebox.askquestion(
                "检测到未提交改动",
                "检测到本地有未提交的修改。\n\n是否要处理？\n\n"
                "是 = 选择操作方式\n"
                "否 = 取消 Pull 操作"
            )
            if choice == "no":
                return False

            action = simpledialog.askstring(
                "选择操作",
                "请输入操作方式：\n"
                "1 = 提交改动\n"
                "2 = 暂存改动\n"
                "3 = 丢弃改动\n"
                "其他 = 取消"
            )

            if action == "1":
                msg = simpledialog.askstring("提交信息", "请输入提交信息：", initialvalue="本地修改")
                self.append_output("提交改动中...")
                self.run_command("git add .")
                self.run_command(f'git commit -m "{msg}"')
            elif action == "2":
                self.append_output("暂存改动中...")
                self.run_command("git stash")
            elif action == "3":
                self.append_output("丢弃改动中...")
                self.run_command("git reset --hard")
            else:
                self.append_output("取消操作。")
                return False
        return True

    def update_repo(self):
        """拉取远程最新代码"""
        self.check_repo_mode()

        if not self.handle_local_changes():
            return

        self.append_output("拉取远程最新...")
        result = self.run_command("git pull --rebase")
        self.append_output(result if result else "拉取完成！")

        # 如果用 stash，需要恢复
        if "stash" in result:
            self.append_output("恢复暂存的改动...")
            self.append_output(self.run_command("git stash pop"))

        self.append_output("---------------------")

    def push_repo(self):
        """推送本地代码"""
        self.check_repo_mode()

        self.append_output("准备推送...")
        result = self.run_command("git push origin main")
        self.append_output(result if result else "推送完成！")
        self.append_output("---------------------")


if __name__ == "__main__":
    root = tk.Tk()
    app = GitPushApp(root)
    root.mainloop()
