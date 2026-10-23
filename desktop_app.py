"""Windows 和 macOS 共用的词库与默写工具。"""
from __future__ import annotations
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from scripts.excel_builder import OUTPUT_DIR, build_workbooks
from scripts.export_word import export_word
from scripts.random_exam import DIRECTIONS, filter_records, load_dataset, parse_count, parse_lessons


def default_data_file() -> Path:
    return OUTPUT_DIR.parent / "data" / "vocabulary.json"


def open_folder(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    command = "explorer" if sys.platform == "win32" else "open" if sys.platform == "darwin" else "xdg-open"
    subprocess.Popen([command, str(path)])


class VocabularyApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("大家的日语词库工具")
        self.geometry("800x650")
        self.minsize(760, 620)
        builtin = default_data_file()
        self.data_path = tk.StringVar(value=str(builtin) if builtin.is_file() else "")
        self.output_path = tk.StringVar(value=str(OUTPUT_DIR))
        self.lessons = tk.StringVar(value="全部")
        self.amount = tk.StringVar(value="20")
        self.direction = tk.StringVar(value="中译日")
        self.preview = tk.StringVar()
        self.available_units = tk.StringVar()
        self.unit_label = tk.StringVar(value="课程")
        self.status = tk.StringVar(value="每次生成都会重新随机。宏版可在 Excel 内手动反复随机，答案在作答栏右侧。")
        self.controls = []
        self.busy = False
        self.results = queue.Queue()
        self.preview_job = None
        self._build_ui()
        for var in (self.data_path, self.lessons, self.amount, self.direction):
            var.trace_add("write", self._schedule_preview)
        self._update_preview()
        self.after(100, self._poll_results)
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _control(self, widget):
        self.controls.append((widget, str(widget.cget("state"))))
        return widget

    def _build_ui(self):
        frame = ttk.Frame(self, padding=24)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="大家的日语词库工具", font=("Arial", 21, "bold")).pack(anchor="w")
        ttk.Label(frame, text="软件设置用于初次生成；打开宏版后，可在 Excel 中更改设置并反复随机。", foreground="#52606D").pack(anchor="w", pady=(8, 18))
        files = ttk.LabelFrame(frame, text="文件位置", padding=12)
        files.pack(fill="x")
        files.columnconfigure(1, weight=1)
        for row, (label, var, caption, action) in enumerate([
            ("词库", self.data_path, "选择 JSON", self.choose_data_file),
            ("保存到", self.output_path, "选择文件夹", self.choose_output_folder),
        ]):
            ttk.Label(files, text=label).grid(row=row, column=0, padx=(0, 10), pady=6)
            self._control(ttk.Entry(files, textvariable=var)).grid(row=row, column=1, sticky="ew", padx=(0, 10), pady=6)
            self._control(ttk.Button(files, text=caption, command=action)).grid(row=row, column=2)
        options = ttk.LabelFrame(frame, text="本次练习", padding=14)
        options.pack(fill="x", pady=(16, 0))
        options.columnconfigure(1, weight=1)
        ttk.Label(options, textvariable=self.unit_label).grid(row=0, column=0, padx=(0, 16), sticky="w")
        self.unit_picker = self._control(ttk.Combobox(options, textvariable=self.lessons, values=["全部"]))
        self.unit_picker.grid(row=0, column=1, columnspan=3, sticky="ew")
        ttk.Label(options, textvariable=self.available_units, foreground="#52606D", wraplength=600).grid(row=1, column=1, columnspan=3, sticky="w", pady=(5, 12))
        ttk.Label(options, text="方向").grid(row=2, column=0, sticky="w")
        self._control(ttk.Combobox(options, textvariable=self.direction, values=list(DIRECTIONS), state="readonly", width=14)).grid(row=2, column=1, sticky="w")
        ttk.Label(options, text="题数").grid(row=2, column=2, padx=(16, 10))
        self._control(ttk.Combobox(options, textvariable=self.amount, values=[20, 50, 100, 200, "全部"], width=12)).grid(row=2, column=3, sticky="e")
        ttk.Label(options, textvariable=self.preview, wraplength=680, foreground="#375A7F").grid(row=3, column=0, columnspan=4, sticky="w", pady=(14, 0))
        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(18, 12))
        actions.columnconfigure((0, 1), weight=1)
        self._control(ttk.Button(actions, text="生成 Excel（含手动随机宏版）", command=lambda: self.start_generation("excel"))).grid(row=0, column=0, sticky="ew", padx=(0, 8), ipady=10)
        self._control(ttk.Button(actions, text="生成所选课程 Word 词库", command=lambda: self.start_generation("word"))).grid(row=0, column=1, sticky="ew", padx=(8, 0), ipady=10)
        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 10))
        ttk.Button(frame, text="打开输出文件夹", command=self._open_output).pack(anchor="w")
        ttk.Separator(frame).pack(fill="x", pady=14)
        ttk.Label(frame, textvariable=self.status, wraplength=720, justify="left").pack(fill="x")

    def choose_data_file(self):
        selected = filedialog.askopenfilename(title="选择词库 JSON", initialdir=str(default_data_file().parent), filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")])
        if selected:
            self.lessons.set("全部")
            self.data_path.set(selected)

    def choose_output_folder(self):
        selected = filedialog.askdirectory(title="选择保存文件夹", initialdir=self.output_path.get())
        if selected:
            self.output_path.set(selected)

    def _open_output(self):
        try:
            if not self.output_path.get().strip():
                raise ValueError("请先选择保存文件夹。")
            open_folder(Path(self.output_path.get()).expanduser())
        except (OSError, ValueError) as error:
            messagebox.showerror("无法打开文件夹", str(error))

    def _schedule_preview(self, *_):
        if self.preview_job is not None:
            self.after_cancel(self.preview_job)
        self.preview_job = self.after(200, self._update_preview)

    def _read_selection(self):
        source = Path(self.data_path.get().strip()).expanduser()
        if not source.is_file():
            raise ValueError("请选择有效的词库 JSON 文件。")
        records, _ = load_dataset(source)
        lessons = parse_lessons(self.lessons.get(), (r["lesson"] for r in records))
        selected = filter_records(records, lessons)
        return source, records, selected, lessons

    def _update_preview(self):
        self.preview_job = None
        try:
            source = Path(self.data_path.get().strip()).expanduser()
            records, metadata = load_dataset(source)
            units = sorted({r["lesson"] for r in records})
            self.unit_picker.configure(values=["全部", *units])
            self.unit_label.set(metadata["unit_label"])
            display = f"{units[0]}～{units[-1]}" if units == list(range(units[0], units[-1]+1)) else ",".join(map(str, units))
            self.available_units.set(f"词库可选{metadata['unit_label']}：{display}；也可输入多个编号或范围。")
            lessons = parse_lessons(self.lessons.get(), units)
            selected = filter_records(records, lessons)
            count = parse_count(self.amount.get())
            available = sum(bool(r["chinese"].strip()) for r in selected)
            actual = available if count is None else min(count, available)
            courses = len({r["lesson"] for r in selected})
            self.preview.set(f"{metadata['title']}：已选 {courses} 个{metadata['unit_label']}，共 {len(selected)} 条词条，可用 {available} 题；本次将生成 {actual} 题。" + (" 题数不足时按实际数量生成。" if count is not None and count > available else ""))
        except (OSError, UnicodeError, ValueError) as error:
            self.preview.set(str(error))

    def start_generation(self, kind):
        if self.busy:
            return
        try:
            source, _, selected, lessons = self._read_selection()
            count = parse_count(self.amount.get()) if kind == "excel" else 20
            if kind == "excel" and not any(r["chinese"].strip() for r in selected):
                raise ValueError("所选课程没有带中文释义的可用题目。")
            if not self.output_path.get().strip():
                raise ValueError("请选择保存文件夹。")
            destination = Path(self.output_path.get().strip()).expanduser()
            direction = DIRECTIONS[self.direction.get()]
        except (OSError, UnicodeError, ValueError, KeyError) as error:
            messagebox.showerror("请检查设置", str(error))
            return
        self.busy = True
        for control, _ in self.controls:
            control.configure(state="disabled")
        self.progress.start(12)
        self.status.set("正在生成，请稍候……")
        threading.Thread(target=self._generate, args=(kind, source, destination, lessons, count, direction), daemon=True).start()

    def _generate(self, kind, source, destination, lessons, count, direction):
        # 后台线程只处理文件和消息队列，不调用任何 Tk 方法。
        try:
            if kind == "excel":
                paths = build_workbooks(source, destination, lessons=lessons, count=count, direction=direction)
            else:
                _, metadata = load_dataset(source)
                paths = (export_word(source, destination / f"{metadata['title']}词库.docx", lessons=lessons),)
            self.results.put((True, (paths, destination)))
        except Exception as error:
            self.results.put((False, str(error)))

    def _poll_results(self):
        try:
            success, payload = self.results.get_nowait()
        except queue.Empty:
            pass
        else:
            self.busy = False
            self.progress.stop()
            for control, state in self.controls:
                control.configure(state=state)
            if success:
                paths, destination = payload
                self.status.set(f"已生成 {len(paths)} 个文件，保存到：{destination}")
                messagebox.showinfo("生成完成", f"保存位置：{destination}\n\n" + "\n".join(p.name for p in paths))
            else:
                self.status.set(f"生成失败：{payload}")
                messagebox.showerror("生成失败", payload)
        self.after(100, self._poll_results)

    def _close(self):
        if self.busy:
            messagebox.showinfo("正在生成", "文件正在生成，请完成后再关闭窗口。")
        else:
            self.destroy()


def main():
    if "--generate-excel-and-exit" in sys.argv:
        build_workbooks(default_data_file(), OUTPUT_DIR)
    elif "--generate-word-and-exit" in sys.argv:
        export_word(default_data_file(), OUTPUT_DIR / "大家的日语Ⅰ词库.docx")
    else:
        VocabularyApp().mainloop()


if __name__ == "__main__":
    main()
