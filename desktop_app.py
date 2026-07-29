"""Windows 和 macOS 共用的桌面生成器界面。"""

from __future__ import annotations

import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from scripts.excel_builder import OUTPUT_DIR, build_workbooks


def open_output_folder() -> None:
    """使用当前系统的文件管理器打开输出目录。"""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if sys.platform == "win32":
        subprocess.Popen(["explorer", str(OUTPUT_DIR)])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(OUTPUT_DIR)])
    else:
        subprocess.Popen(["xdg-open", str(OUTPUT_DIR)])


class VocabularyApp(tk.Tk):
    """提供一键生成和输出目录入口。"""

    def __init__(self) -> None:
        super().__init__()
        self.title("大家的日语词库生成器")
        self.geometry("520x300")
        self.resizable(False, False)
        self.status = tk.StringVar(value="修改 data/vocabulary.json 后，点击下方按钮。")
        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=28)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="大家的日语词库生成器", font=("Arial", 20, "bold")).pack(pady=(0, 15))
        ttk.Label(
            frame,
            text="数据源：data/vocabulary.json\n生成结果：output 文件夹",
            justify="center",
        ).pack(pady=(0, 18))
        self.generate_button = ttk.Button(frame, text="重新生成 Excel", command=self.generate)
        self.generate_button.pack(ipadx=28, ipady=8)
        ttk.Button(frame, text="打开 output 文件夹", command=open_output_folder).pack(pady=12)
        ttk.Separator(frame).pack(fill="x", pady=8)
        ttk.Label(frame, textvariable=self.status, wraplength=450, justify="center").pack()

    def generate(self) -> None:
        self.generate_button.configure(state="disabled")
        self.status.set("正在生成，请稍候……")
        threading.Thread(target=self._generate_in_background, daemon=True).start()

    def _generate_in_background(self) -> None:
        try:
            build_workbooks()
        except Exception as error:
            self.after(0, self._finish_error, str(error))
        else:
            self.after(0, self._finish_success)

    def _finish_success(self) -> None:
        self.generate_button.configure(state="normal")
        self.status.set("生成成功：大家的日语Ⅰ词库.xlsx、大家的日语Ⅰ默写.xlsx")
        messagebox.showinfo("生成完成", "Excel 已生成到 output 文件夹。")

    def _finish_error(self, error: str) -> None:
        self.generate_button.configure(state="normal")
        self.status.set(f"生成失败：{error}")
        messagebox.showerror("生成失败", error)


def main() -> None:
    if "--generate-and-exit" in sys.argv:
        build_workbooks()
        return
    VocabularyApp().mainloop()


if __name__ == "__main__":
    main()
