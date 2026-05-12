import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import threading
import time
import re
import pyperclip
import keyboard
import pandas as pd
from datetime import datetime
from pathlib import Path


URL_PATTERN = re.compile(
    r'https?://[^\s<>\\"{}|\\^`\[\]]+', re.IGNORECASE)


class DouyinLinkCollector:
    def __init__(self, root):
        self.root = root
        self.root.title("抖音视频链接收集器")
        self.root.geometry("700x600")

        self.is_running = False
        self.is_paused = False
        self.worker_thread = None
        self.stop_event = threading.Event()

        self.links = []
        self.last_link = ""
        self.duplicate_count = 0
        self.recorded_links = set()

        self.interval = tk.DoubleVar(value=1.5)

        self.excel_path = Path("douyin_links.xlsx")

        self.setup_ui()
        self.load_existing_links()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        control_frame = ttk.LabelFrame(main_frame, text="控制", padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 10))

        self.start_btn = ttk.Button(control_frame, text="开始收集", command=self.start_collection)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.stop_btn = ttk.Button(control_frame, text="停止", command=self.stop_collection, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        self.pause_btn = ttk.Button(control_frame, text="暂停", command=self.toggle_pause, state=tk.DISABLED)
        self.pause_btn.pack(side=tk.LEFT, padx=5)

        self.clear_btn = ttk.Button(control_frame, text="清空", command=self.clear_links)
        self.clear_btn.pack(side=tk.LEFT, padx=5)

        self.export_btn = ttk.Button(control_frame, text="导出", command=self.export_links)
        self.export_btn.pack(side=tk.LEFT, padx=5)

        config_frame = ttk.Frame(control_frame)
        config_frame.pack(side=tk.RIGHT)

        ttk.Label(config_frame, text="操作间隔(秒):").pack(side=tk.LEFT)
        ttk.Spinbox(config_frame, from_=0.5, to=10, textvariable=self.interval, width=5).pack(side=tk.LEFT, padx=5)

        status_frame = ttk.LabelFrame(main_frame, text="状态", padding="10")
        status_frame.pack(fill=tk.X, pady=(0, 10))

        self.status_label = ttk.Label(status_frame, text="等待开始", foreground="gray")
        self.status_label.pack(side=tk.LEFT)

        self.count_label = ttk.Label(status_frame, text=f"已收集: 0 条")
        self.count_label.pack(side=tk.RIGHT)

        links_frame = ttk.LabelFrame(main_frame, text="已收集的链接", padding="10")
        links_frame.pack(fill=tk.BOTH, expand=True)

        self.links_text = scrolledtext.ScrolledText(links_frame, wrap=tk.WORD, height=15)
        self.links_text.pack(fill=tk.BOTH, expand=True)

    def load_existing_links(self):
        if self.excel_path.exists():
            try:
                df = pd.read_excel(self.excel_path)
                self.links = df.to_dict("records")
                self.recorded_links = {r.get("链接", r.get("link", "")) for r in self.links}
                self.update_links_display()
            except Exception as e:
                print(f"加载历史记录失败: {e}")

    def extract_links(self, text):
        matches = URL_PATTERN.findall(text)
        return [m.rstrip('.,;!?)\"\'') for m in matches]

    def save_links(self):
        if not self.links:
            return

        try:
            df = pd.DataFrame(self.links)
            df.columns = ["时间", "链接"]
            df.to_excel(self.excel_path, index=False, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")

    def update_links_display(self):
        self.links_text.delete(1.0, tk.END)
        for link_data in reversed(self.links):
            time_str = link_data.get("时间", link_data.get("time", ""))
            link_str = link_data.get("链接", link_data.get("link", ""))
            self.links_text.insert(tk.END, f"{time_str} - {link_str}\n\n")
        self.count_label.config(text=f"已收集: {len(self.links)} 条")

    def add_link(self, url):
        if url not in self.recorded_links:
            self.recorded_links.add(url)
            self.links.append({
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "link": url
            })
            self.save_links()
            self.root.after(0, self.update_links_display)
            self.duplicate_count = 0
        else:
            self.duplicate_count += 1

    def set_status(self, text, color="black"):
        self.status_label.config(text=text, foreground=color)

    def worker(self):
        pyperclip.copy("")
        self.last_link = ""
        self.duplicate_count = 0

        while not self.stop_event.is_set():
            if self.is_paused:
                time.sleep(0.1)
                continue

            try:
                self.set_status("正在复制链接...", "blue")
                keyboard.send("alt+v")
                time.sleep(self.interval.get() * 0.3)

                self.set_status("读取链接...", "blue")
                content = pyperclip.paste()

                links = self.extract_links(content)
                if links:
                    current_link = links[0]
                    self.add_link(current_link)

                    if self.duplicate_count >= 2:
                        self.set_status("检测到连续重复链接，停止收集", "red")
                        self.stop_collection()
                        break

                    self.last_link = current_link
                    self.set_status("切换下一个视频...", "green")
                    keyboard.send("down")
                else:
                    self.set_status("未找到链接，重试...", "orange")

            except Exception as e:
                self.set_status(f"错误: {str(e)[:30]}", "red")

            time.sleep(self.interval.get() * 0.7)

    def start_collection(self):
        self.is_running = True
        self.is_paused = False
        self.stop_event.clear()

        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.NORMAL)
        self.pause_btn.config(text="暂停")

        self.worker_thread = threading.Thread(target=self.worker, daemon=True)
        self.worker_thread.start()
        self.set_status("正在收集...", "green")

    def stop_collection(self):
        self.is_running = False
        self.stop_event.set()

        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.DISABLED)

        self.set_status("已停止", "gray")

    def toggle_pause(self):
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_btn.config(text="继续")
            self.set_status("已暂停", "orange")
        else:
            self.pause_btn.config(text="暂停")
            self.set_status("正在收集...", "green")

    def clear_links(self):
        if messagebox.askyesno("确认", "确定要清空所有链接吗？"):
            self.links = []
            self.recorded_links.clear()
            self.update_links_display()
            self.save_links()

    def export_links(self):
        if not self.links:
            messagebox.showwarning("警告", "没有可导出的链接")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel 文件", "*.xlsx"),
                ("CSV 文件", "*.csv"),
                ("所有文件", "*.*")
            ],
            initialfile=f"douyin_links_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )

        if file_path:
            try:
                df = pd.DataFrame(self.links)
                df.columns = ["时间", "链接"]

                if file_path.endswith(".csv"):
                    df.to_csv(file_path, index=False, encoding="utf-8-sig")
                else:
                    df.to_excel(file_path, index=False, engine="openpyxl")

                messagebox.showinfo("成功", "导出成功！")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {e}")


def main():
    root = tk.Tk()
    app = DouyinLinkCollector(root)
    root.mainloop()


if __name__ == "__main__":
    main()
