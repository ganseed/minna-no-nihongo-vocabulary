"""Windows 和 macOS 共用的日语词库桌面工具。"""

from __future__ import annotations

import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from scripts.excel_builder import OUTPUT_DIR, build_workbooks
from scripts.export_word import export_word


def default_data_file() -> Path:
    """返回发布包内置词库路径。"""
    return OUTPUT_DIR.parent / "data" / "vocabulary.json"


def open_folder(path: Path) -> None:
    """使用当前系统的文件管理器打开目录。"""
    path.mkdir(parents=True, exist_ok=True)
    if sys.platform == "win32":
        subprocess.Popen(["explorer", str(path)])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class VocabularyApp(tk.Tk):
    """提供数据选择和分类导出功能。"""

    def __init__(self) -> None:
        super().__init__()
        self.title("大家的日语词库工具")
        self.geometry("680x420")
        self.minsize(620, 390)
        self.data_path = tk.StringVar(value=str(default_data_file()))
        self.status = tk.StringVar(value="请选择词库 JSON，然后选择需要生成的文件。")
        self.action_buttons: list[ttk.Button] = []
        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=28)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="大家的日语词库工具", font=("Arial", 21, "bold")).pack()
        ttk.Label(
            container,
            text="从 JSON 生成可编辑的 Word 词库、Excel 词库和默写模板",
            foreground="#52606D",
        ).pack(pady=(6, 24))

        source_frame = ttk.LabelFrame(container, text="词库数据", padding=14)
        source_frame.pack(fill="x")
        source_frame.columnconfigure(0, weight=1)
        ttk.Entry(source_frame, textvariable=self.data_path).grid(row=0, column=0, sticky="ew", padx=(0, 10))
        ttk.Button(source_frame, text="选择 JSON 文件", command=self.choose_data_file).grid(row=0, column=1)
        ttk.Label(
            source_frame,
            text="默认使用发布包 data 文件夹中的 vocabulary.json，也可以选择其他兼容词库。",
            foreground="#6B7280",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(9, 0))

        action_frame = ttk.Frame(container)
        action_frame.pack(fill="x", pady=24)
        action_frame.columnconfigure((0, 1), weight=1)
        excel_button = ttk.Button(
            action_frame,
            text="生成 Excel 词库 + 默写模板",
            command=lambda: self.start_generation("excel"),
        )
        excel_button.grid(row=0, column=0, sticky="ew", padx=(0, 8), ipady=10)
        word_button = ttk.Button(
            action_frame,
            text="生成 Word 词库",
            command=lambda: self.start_generation("word"),
        )
        word_button.grid(row=0, column=1, sticky="ew", padx=(8, 0), ipady=10)
        self.action_buttons.extend([excel_button, word_button])

        ttk.Button(
            container,
            text="打开输出文件夹",
            command=lambda: open_folder(OUTPUT_DIR),
        ).pack()
        ttk.Separator(container).pack(fill="x", pady=20)
        ttk.Label(container, textvariable=self.status, wraplength=610, justify="center").pack()

    def choose_data_file(self) -> None:
        """让用户选择其他兼容的 JSON 词库。"""
        selected = filedialog.askopenfilename(
            title="选择词库 JSON",
            initialdir=str(default_data_file().parent),
            filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")],
        )
        if selected:
            self.data_path.set(selected)
            self.status.set(f"已选择：{Path(selected).name}")

    def start_generation(self, kind: str) -> None:
        """检查数据文件并在后台执行生成任务。"""
        source = Path(self.data_path.get().strip()).expanduser()
        if not source.is_file():
            messagebox.showerror("找不到词库", "请选择有效的 JSON 文件。")
            return
        for button in self.action_buttons:
            button.configure(state="disabled")
        label = "Excel 词库和默写模板" if kind == "excel" else "Word 词库"
        self.status.set(f"正在生成{label}，请稍候……")
        threading.Thread(target=self._generate, args=(kind, source), daemon=True).start()

    def _generate(self, kind: str, source: Path) -> None:
        try:
            if kind == "excel":
                paths = build_workbooks(source, OUTPUT_DIR)
            else:
                paths = (export_word(source, OUTPUT_DIR / "大家的日语Ⅰ词库.docx"),)
        except Exception as error:
            self.after(0, self._finish_error, str(error))
        else:
            self.after(0, self._finish_success, paths)

    def _finish_success(self, paths: tuple[Path, ...]) -> None:
        for button in self.action_buttons:
            button.configure(state="normal")
        names = "、".join(path.name for path in paths)
        self.status.set(f"生成成功：{names}")
        messagebox.showinfo("生成完成", f"文件已保存到 output 文件夹：\n{names}")

    def _finish_error(self, error: str) -> None:
        for button in self.action_buttons:
            button.configure(state="normal")
        self.status.set(f"生成失败：{error}")
        messagebox.showerror("生成失败", error)


def main() -> None:
    if "--generate-excel-and-exit" in sys.argv:
        build_workbooks(default_data_file(), OUTPUT_DIR)
        return
    if "--generate-word-and-exit" in sys.argv:
        export_word(default_data_file(), OUTPUT_DIR / "大家的日语Ⅰ词库.docx")
        return
    VocabularyApp().mainloop()


if __name__ == "__main__":
    main()
