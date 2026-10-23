"""Windows 和 macOS 共用的词库与默写工具。"""
from __future__ import annotations
import argparse
import errno
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

from scripts.excel_builder import DATA_FILE, OUTPUT_DIR, build_workbooks
from scripts.export_word import export_word
from scripts.random_exam import DIRECTIONS, filter_records, load_dataset, parse_count, parse_lessons


def enable_high_dpi() -> None:
    """必须在创建 Tk 窗口前设置，避免 Windows 对文字进行位图拉伸。"""
    if sys.platform != "win32":
        return
    import ctypes
    try:
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


def default_data_file() -> Path:
    return DATA_FILE


def open_folder(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    command = "explorer" if sys.platform == "win32" else "open" if sys.platform == "darwin" else "xdg-open"
    subprocess.Popen([command, str(path)])


class VocabularyApp(tk.Tk):
    def __init__(self) -> None:
        enable_high_dpi()
        super().__init__()
        self.withdraw()
        self._configure_appearance()
        self.title("大家的日语词库工具")

        builtin = default_data_file()
        self.data_path = tk.StringVar(value=str(builtin) if builtin.is_file() else "")
        self.output_path = tk.StringVar(value=str(OUTPUT_DIR))
        self.lessons = tk.StringVar(value="全部")
        self.amount = tk.StringVar(value="20")
        self.direction = tk.StringVar(value="中译日")
        self.preview = tk.StringVar()
        self.question_total = tk.StringVar(value="—")
        self.library_title = tk.StringVar(value="当前词库")
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
        self.update_idletasks()
        width = min(max(self.winfo_reqwidth(), self.px(820)), self.winfo_screenwidth() - self.px(56))
        height = min(self.winfo_reqheight(), self.winfo_screenheight() - self.px(100))
        self.geometry(f"{width}x{height}+{max(0, (self.winfo_screenwidth()-width)//2)}+{max(0, (self.winfo_screenheight()-height)//2-self.px(20))}")
        self.minsize(min(width, self.px(760)), min(height, self.px(620)))
        self.deiconify()

    def _control(self, widget):
        self.controls.append((widget, str(widget.cget("state"))))
        return widget

    def _configure_appearance(self):
        self.ui_scale = max(1.0, self.winfo_fpixels("1i") / 96.0)
        candidates = ("Microsoft YaHei UI", "Microsoft YaHei") if sys.platform == "win32" else ("PingFang SC", "Hiragino Sans GB", "Noto Sans CJK SC")
        available = set(tkfont.families(self))
        self.font_family = next((name for name in candidates if name in available), "TkDefaultFont")
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont", "TkTooltipFont"):
            tkfont.nametofont(name).configure(family=self.font_family, size=10)
        self.option_add("*TCombobox*Listbox.font", (self.font_family, 10))
        self.option_add("*TCombobox*Listbox.background", "#FFFFFF")
        self.option_add("*TCombobox*Listbox.selectBackground", "#4F46E5")
        self.configure(background="#F3F5FA")
        # 用原生像素图标替换 Tk 默认图标，不引入图片依赖。
        self.app_icon = tk.PhotoImage(width=32, height=32)
        self.app_icon.put("#4F46E5", to=(0, 0, 32, 32))
        self.app_icon.put("#FFFFFF", to=(7, 6, 25, 26))
        self.app_icon.put("#4F46E5", to=(10, 9, 22, 14))
        self.app_icon.put("#4F46E5", to=(10, 17, 22, 23))
        self.iconphoto(True, self.app_icon)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=(self.font_family, 10), background="#F3F5FA", foreground="#1E293B")
        style.configure("TFrame", background="#F3F5FA")
        style.configure("Card.TFrame", background="#FFFFFF")
        style.configure("TLabel", background="#F3F5FA")
        style.configure("Title.TLabel", font=(self.font_family, 24, "bold"), foreground="#17223B")
        style.configure("Muted.TLabel", foreground="#64748B")
        style.configure("Card.TLabel", background="#FFFFFF")
        style.configure("CardTitle.TLabel", background="#FFFFFF", font=(self.font_family, 12, "bold"))
        style.configure("Hint.TLabel", background="#FFFFFF", foreground="#64748B", font=(self.font_family, 9))
        style.configure("Metric.TLabel", background="#EEF2FF", foreground="#4338CA", font=(self.font_family, 27, "bold"))
        style.configure("Summary.TFrame", background="#EEF2FF")
        style.configure("Summary.TLabel", background="#EEF2FF", foreground="#4338CA")
        style.configure("TEntry", padding=self.px(9), fieldbackground="#F8FAFC", bordercolor="#DCE2EC", lightcolor="#DCE2EC", darkcolor="#DCE2EC")
        style.configure("TCombobox", padding=self.px(9), fieldbackground="#F8FAFC", background="#F8FAFC", bordercolor="#DCE2EC", arrowsize=self.px(14))
        for control in ("TEntry", "TCombobox"):
            style.map(control, bordercolor=[("focus", "#6366F1")], fieldbackground=[("disabled", "#F1F5F9"), ("readonly", "#F8FAFC")], foreground=[("disabled", "#94A3B8"), ("readonly", "#1E293B")])
        style.configure("TButton", padding=(self.px(14), self.px(9)), background="#EDF0F7", foreground="#334155", borderwidth=0, focusthickness=2, focuscolor="#6366F1")
        style.map("TButton", background=[("disabled", "#E8ECF3"), ("pressed", "#D9DFED"), ("active", "#E2E7F2")], foreground=[("disabled", "#94A3B8")])
        style.configure("Primary.TButton", background="#4F46E5", foreground="#FFFFFF", font=(self.font_family, 11, "bold"), padding=(self.px(18), self.px(12)))
        style.map("Primary.TButton", background=[("disabled", "#CDD2EC"), ("pressed", "#3730A3"), ("active", "#4338CA")], foreground=[("disabled", "#F8FAFC")])
        style.configure("Secondary.TButton", padding=(self.px(18), self.px(12)), font=(self.font_family, 11))
        style.configure("Horizontal.TProgressbar", troughcolor="#E8ECF4", background="#6366F1", borderwidth=0, thickness=self.px(3))

    def px(self, value):
        return round(value * self.ui_scale)

    def _build_ui(self):
        # 内容可滚动：高缩放和小屏幕下仍能访问底部操作。
        shell = ttk.Frame(self)
        shell.pack(fill="both", expand=True)
        canvas = tk.Canvas(shell, background="#F3F5FA", highlightthickness=0, width=self.px(820))
        scrollbar = ttk.Scrollbar(shell, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        content = ttk.Frame(canvas, padding=self.px(28))
        window_id = canvas.create_window(0, 0, window=content, anchor="nw")
        def resize_content(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            if content.winfo_reqheight() > canvas.winfo_height():
                scrollbar.pack(side="right", fill="y")
            else:
                scrollbar.pack_forget()
        def resize_canvas(event):
            canvas.itemconfigure(window_id, width=event.width)
            resize_content()
        content.bind("<Configure>", resize_content)
        canvas.bind("<Configure>", resize_canvas)
        self.bind("<MouseWheel>", lambda event: canvas.yview_scroll(-int(event.delta / 120) if sys.platform == "win32" else -event.delta, "units") if content.winfo_reqheight() > canvas.winfo_height() and not isinstance(event.widget, ttk.Combobox) else None)
        ttk.Label(content, text="日语练习", style="Title.TLabel").pack(anchor="w")
        ttk.Label(content, text="选好课程，开始今天的默写。", style="Muted.TLabel").pack(anchor="w", pady=(self.px(5), self.px(22)))
        files = ttk.Frame(content, style="Card.TFrame", padding=self.px(18))
        files.pack(fill="x")
        files.columnconfigure(1, weight=1)
        ttk.Label(files, text="词库与保存位置", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, self.px(12)))
        for row, (label, var, caption, action) in enumerate([
            ("词库文件", self.data_path, "更换词库", self.choose_data_file),
            ("保存位置", self.output_path, "选择位置", self.choose_output_folder),
        ], 1):
            ttk.Label(files, text=label, style="Card.TLabel").grid(row=row, column=0, padx=(0, self.px(14)), pady=self.px(5))
            self._control(ttk.Entry(files, textvariable=var, width=24)).grid(row=row, column=1, sticky="ew", padx=(0, self.px(10)), pady=self.px(5))
            self._control(ttk.Button(files, text=caption, command=action)).grid(row=row, column=2)
        options = ttk.Frame(content, style="Card.TFrame", padding=self.px(18))
        options.pack(fill="x", pady=(self.px(14), 0))
        options.columnconfigure((0, 1, 2), weight=1, uniform="selection")
        ttk.Label(options, text="本次练习", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, self.px(14)))
        for col, text in enumerate((None, "默写方向", "题目数量")):
            ttk.Label(options, text=text, textvariable=self.unit_label if col == 0 else None, style="Card.TLabel").grid(row=1, column=col, sticky="w", pady=(0, self.px(6)))
        self.unit_picker = self._control(ttk.Combobox(options, textvariable=self.lessons, values=["全部"], width=12))
        self.unit_picker.grid(row=2, column=0, sticky="ew", padx=(0, self.px(14)))
        self._control(ttk.Combobox(options, textvariable=self.direction, values=list(DIRECTIONS), state="readonly", width=12)).grid(row=2, column=1, sticky="ew", padx=(0, self.px(14)))
        self._control(ttk.Combobox(options, textvariable=self.amount, values=[20, 50, 100, 200, "全部"], width=10)).grid(row=2, column=2, sticky="ew")
        self.unit_hint = ttk.Label(options, textvariable=self.available_units, style="Hint.TLabel", wraplength=self.px(670))
        self.unit_hint.grid(row=3, column=0, columnspan=3, sticky="w", pady=(self.px(10), 0))
        summary = ttk.Frame(content, style="Summary.TFrame", padding=(self.px(18), self.px(14)))
        summary.pack(fill="x", pady=(self.px(14), 0))
        ttk.Label(summary, textvariable=self.question_total, style="Metric.TLabel").pack(side="left", padx=(0, self.px(12)))
        ttk.Label(summary, text="题", style="Summary.TLabel").pack(side="left", padx=(0, self.px(20)))
        detail = ttk.Frame(summary, style="Summary.TFrame")
        detail.pack(side="left", fill="both", expand=True)
        ttk.Label(detail, textvariable=self.library_title, style="Summary.TLabel", font=(self.font_family, 11, "bold")).pack(anchor="w")
        self.preview_label = ttk.Label(detail, textvariable=self.preview, style="Summary.TLabel", wraplength=self.px(530))
        self.preview_label.pack(anchor="w", pady=(self.px(4), 0))
        actions = ttk.Frame(content)
        actions.pack(fill="x", pady=(self.px(18), self.px(10)))
        actions.columnconfigure((0, 1), weight=1)
        self._control(ttk.Button(actions, text="生成 Excel · 含随机宏版", style="Primary.TButton", command=lambda: self.start_generation("excel"))).grid(row=0, column=0, sticky="ew", padx=(0, self.px(10)))
        self._control(ttk.Button(actions, text="导出 Word 词库", style="Secondary.TButton", command=lambda: self.start_generation("word"))).grid(row=0, column=1, sticky="ew")
        self.progress = ttk.Progressbar(content, mode="indeterminate")
        # 仅生成过程中显示进度条。
        footer = ttk.Frame(content)
        self.footer = footer
        footer.pack(fill="x")
        self.status_label = ttk.Label(footer, textvariable=self.status, style="Muted.TLabel", wraplength=self.px(540), justify="left")
        self.status_label.pack(side="left", fill="x", expand=True, padx=(0, self.px(12)))
        ttk.Button(footer, text="打开文件夹", command=self._open_output).pack(side="right")
        content.bind("<Configure>", lambda event: (resize_content(), self.unit_hint.configure(wraplength=max(self.px(300), event.width-self.px(92))), self.preview_label.configure(wraplength=max(self.px(220), event.width-self.px(230))), self.status_label.configure(wraplength=max(self.px(220), event.width-self.px(210)))))
        self.content_canvas = canvas
        self.content_frame = content
        # Canvas 的自然高度不随内部 frame 改变，初始尺寸按完整内容计算。
        self.update_idletasks()
        canvas.configure(height=content.winfo_reqheight())

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
            self.question_total.set(str(actual))
            self.library_title.set(metadata["title"])
            self.preview.set(f"已选 {courses} 个{metadata['unit_label']} · 可用 {available} 题 · 每次重新随机" + ("；按实际题数生成" if count is not None and count > available else ""))
        except (OSError, UnicodeError, ValueError) as error:
            self.question_total.set("—")
            self.library_title.set("请检查练习设置")
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
        self.progress.pack(fill="x", pady=(0, self.px(10)), before=self.footer)
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
            message = str(error)
            if isinstance(error, OSError) and error.errno in {errno.EROFS, errno.EACCES, errno.EPERM}:
                message = "保存位置无法写入，请选择“文稿”或“下载”等可写文件夹后重试。\n\n" + message
            self.results.put((False, message))

    def _poll_results(self):
        try:
            success, payload = self.results.get_nowait()
        except queue.Empty:
            pass
        else:
            self.busy = False
            self.progress.stop()
            self.progress.pack_forget()
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
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--output", type=Path, default=OUTPUT_DIR)
    args, _ = parser.parse_known_args()
    if "--generate-excel-and-exit" in sys.argv:
        build_workbooks(default_data_file(), args.output)
    elif "--generate-word-and-exit" in sys.argv:
        export_word(default_data_file(), args.output / "大家的日语Ⅰ词库.docx")
    else:
        VocabularyApp().mainloop()


if __name__ == "__main__":
    main()
