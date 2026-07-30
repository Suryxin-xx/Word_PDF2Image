"""
图形界面 — PDF / Word 导出为图片

功能：
  - 选择 PDF / Word 文件 + 输出目录
  - 选择图片格式（PNG/JPEG/TIFF/BMP/WEBP）
  - 调节 DPI 和质量
  - 选择页面范围（全部 / 自定义）
  - 实时进度条
  - 完成后打开输出目录
"""

import os
import threading
import tkinter as tk
import ctypes
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from src.converter import (
    document_to_images,
    SUPPORTED_FORMATS,
    SUPPORTED_INPUT_EXTENSIONS,
    FORMAT_KEYS,
    DEFAULT_FORMAT,
)


class PDF2ImageApp:
    """PDF / Word 导出为图片 — 主窗口"""

    def __init__(self):
        self._enable_dpi_awareness()
        self.root = tk.Tk()
        self.root.title("PDF/Word导出为图片")
        self.root.resizable(True, True)
        self._setup_style()

        self.input_path = tk.StringVar()
        self.out_dir = tk.StringVar()
        self.fmt_var = tk.StringVar(value=DEFAULT_FORMAT)
        self.dpi_var = tk.IntVar(value=200)
        self.quality_var = tk.IntVar(value=90)
        self.quality_input_var = tk.StringVar(value="90")
        self.input_type_var = tk.StringVar(value="PDF / WORD")
        self.input_status_var = tk.StringVar(value="等待选择文档")
        self.output_type_var = tk.StringVar(value=DEFAULT_FORMAT)
        self.enhance_option_text = tk.StringVar(
            value="优化扫描件\n锐化 · 去黄 · 对比度"
        )

        # 页面范围
        self.page_range_var = tk.StringVar(value="all")  # "all" 或 "custom"
        self.custom_pages_var = tk.StringVar()

        # 扫描件增强
        self.enhance_var = tk.BooleanVar(value=False)
        self.enhance_sharpness = tk.IntVar(value=80)
        self.enhance_cutoff = tk.IntVar(value=2)
        self.enhance_contrast = tk.DoubleVar(value=1.15)
        self.input_path.trace_add("write", lambda *_: self._on_input_path_change())
        self.dpi_var.trace_add("write", lambda *_: self._refresh_summary())
        self.quality_var.trace_add("write", lambda *_: self._on_quality_change())
        self.page_range_var.trace_add("write", lambda *_: self._on_page_mode_change())
        self.custom_pages_var.trace_add("write", lambda *_: self._refresh_summary())

        self._build_ui()
        self._apply_initial_window_size()
        self._running = False

    @staticmethod
    def _enable_dpi_awareness():
        """尽量避免 Windows 对 Tk 窗口做模糊缩放。"""
        if os.name != "nt":
            return
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

    # ------------------------------------------------------------------
    # 主题样式
    # ------------------------------------------------------------------
    def _setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")

        self.colors = {
            "workspace": "#F3F6FA",
            "card": "#FFFFFF",
            "ink": "#172033",
            "muted": "#667085",
            "subtle": "#98A2B3",
            "accent": "#2F6FED",
            "accent_hover": "#245CC7",
            "accent_soft": "#E9F0FF",
            "border": "#DDE3EC",
            "field": "#F8FAFD",
            "success": "#12845B",
        }
        bg = self.colors["workspace"]
        card_bg = self.colors["card"]
        fg = self.colors["ink"]
        sec_fg = self.colors["muted"]
        accent = self.colors["accent"]
        border = self.colors["border"]

        self.root.configure(bg=bg)

        font = ("Microsoft YaHei UI", 10)
        font_sm = ("Microsoft YaHei UI", 9)
        font_btn = ("Microsoft YaHei UI", 10, "bold")
        font_bold = ("Microsoft YaHei UI", 10, "bold")

        style.configure(".", background=bg, foreground=fg, font=font)
        style.configure("Workspace.TFrame", background=bg)
        style.configure("Card.TFrame", background=card_bg)
        style.configure("TLabel", background=card_bg, foreground=fg)
        style.configure(
            "Kicker.TLabel", background=card_bg, foreground=accent,
            font=("Segoe UI", 9, "bold"),
        )
        style.configure(
            "HeroTitle.TLabel", background=card_bg, foreground=fg,
            font=("Microsoft YaHei UI", 22, "bold"),
        )
        style.configure("HeroSub.TLabel", background=card_bg, foreground=sec_fg,
                        font=("Microsoft YaHei UI", 10))
        style.configure(
            "CardTitle.TLabel", background=card_bg, foreground=fg,
            font=("Microsoft YaHei UI", 13, "bold"),
        )
        style.configure(
            "FieldLabel.TLabel", background=card_bg, foreground=fg,
            font=font_bold,
        )
        style.configure(
            "Secondary.TLabel", background=card_bg, foreground=sec_fg,
            font=font_sm,
        )
        style.configure(
            "Micro.TLabel", background=card_bg, foreground=self.colors["subtle"],
            font=("Segoe UI", 8, "bold"),
        )
        style.configure(
            "InputBadge.TLabel", background=self.colors["accent_soft"],
            foreground=accent, font=("Segoe UI", 10, "bold"), padding=(12, 7),
        )
        style.configure(
            "OutputBadge.TLabel", background=accent, foreground="#FFFFFF",
            font=("Segoe UI", 10, "bold"), padding=(12, 7),
        )
        style.configure(
            "Arrow.TLabel", background=card_bg, foreground=self.colors["subtle"],
            font=("Segoe UI", 15),
        )
        style.configure(
            "Status.TLabel", background=card_bg, foreground=fg,
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.configure(
            "Summary.TLabel", background=card_bg, foreground=sec_fg,
            font=font_sm,
        )
        style.configure(
            "Callout.TFrame", background=self.colors["accent_soft"],
        )
        style.configure(
            "CalloutTitle.TLabel", background=self.colors["accent_soft"],
            foreground=accent, font=font_bold,
        )
        style.configure(
            "CalloutText.TLabel", background=self.colors["accent_soft"],
            foreground=sec_fg, font=font_sm,
        )

        style.configure(
            "Action.TButton", background=accent, foreground="white",
            borderwidth=0, focuscolor=accent, font=font_btn, padding=(28, 12),
        )
        style.map(
            "Action.TButton",
            background=[
                ("disabled", "#AFC4F3"),
                ("active", self.colors["accent_hover"]),
                ("pressed", "#1E4FAE"),
            ],
            foreground=[("disabled", "#F7F9FC")],
        )
        style.configure("Browse.TButton", background=card_bg, foreground=fg,
                        borderwidth=1, bordercolor=border, focuscolor=accent,
                        font=font, padding=(12, 7))
        style.map(
            "Browse.TButton",
            background=[
                ("disabled", "#F3F5F8"),
                ("active", self.colors["accent_soft"]),
                ("pressed", "#DCE8FF"),
            ],
            foreground=[
                ("disabled", self.colors["subtle"]),
                ("active", accent),
            ],
        )

        style.configure(
            "TEntry", fieldbackground=self.colors["field"], borderwidth=1,
            bordercolor=border, lightcolor=border, darkcolor=border,
            padding=(10, 8),
        )
        style.map(
            "TEntry",
            bordercolor=[("focus", accent)],
            lightcolor=[("focus", accent)],
            darkcolor=[("focus", accent)],
        )
        style.configure(
            "TCombobox", fieldbackground=self.colors["field"], borderwidth=1,
            bordercolor=border, lightcolor=border, darkcolor=border,
            padding=(9, 7), arrowcolor=sec_fg,
        )
        style.configure("TRadiobutton", background=card_bg, foreground=fg,
                        indicatorforeground=accent)
        style.configure("TCheckbutton", background=card_bg, foreground=fg,
                        indicatorforeground=accent)
        style.map("TRadiobutton", background=[("active", card_bg)])
        style.map("TCheckbutton", background=[("active", card_bg)])

        style.configure("Horizontal.TScale", background=card_bg, troughcolor="#E7ECF3",
                        slidercolor=accent, sliderlength=20, borderwidth=0)

        style.configure(
            "TProgressbar", background=accent, troughcolor="#E7ECF3",
            borderwidth=0, thickness=8,
        )
        style.configure("TSeparator", background=self.colors["border"])

    # ------------------------------------------------------------------
    # 布局
    # ------------------------------------------------------------------
    def _create_card(self, parent, padding=(18, 16)):
        outer = tk.Frame(
            parent,
            bg=self.colors["card"],
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["border"],
            highlightthickness=1,
            bd=0,
        )
        inner = ttk.Frame(outer, style="Card.TFrame", padding=padding)
        inner.pack(fill=tk.BOTH, expand=True)
        return outer, inner

    @staticmethod
    def _section_heading(parent, number, title, subtitle):
        ttk.Label(parent, text=number, style="Kicker.TLabel").grid(
            row=0, column=0, sticky=tk.W
        )
        ttk.Label(parent, text=title, style="CardTitle.TLabel").grid(
            row=1, column=0, sticky=tk.W, pady=(3, 0)
        )
        if subtitle:
            ttk.Label(parent, text=subtitle, style="Secondary.TLabel").grid(
                row=2, column=0, sticky=tk.W, pady=(5, 0)
            )

    def _choice_radio(self, parent, text, variable, value, command=None):
        """创建大尺寸、带明确选中态的块状单选按钮。"""
        return tk.Radiobutton(
            parent,
            text=text,
            variable=variable,
            value=value,
            command=command,
            indicatoron=False,
            bg=self.colors["field"],
            fg=self.colors["ink"],
            selectcolor=self.colors["accent_soft"],
            activebackground=self.colors["accent_soft"],
            activeforeground=self.colors["accent"],
            disabledforeground=self.colors["subtle"],
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"],
            font=("Segoe UI", 10, "bold"),
            padx=10,
            pady=10,
            cursor="hand2",
            takefocus=True,
        )

    def _numeric_entry(self, parent, text_var, start, end, increment, command):
        """创建可输入、可用箭头微调的数值框。"""
        entry = tk.Spinbox(
            parent,
            from_=start,
            to=end,
            increment=increment,
            textvariable=text_var,
            command=command,
            width=6,
            justify=tk.CENTER,
            font=("Segoe UI", 11, "bold"),
            bg=self.colors["field"],
            fg=self.colors["ink"],
            buttonbackground=self.colors["accent_soft"],
            relief=tk.SOLID,
            bd=1,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"],
        )
        entry.bind("<Return>", lambda _event: command())
        entry.bind("<FocusOut>", lambda _event: command())
        entry.bind(
            "<FocusIn>",
            lambda _event: entry.after_idle(lambda: entry.selection_range(0, tk.END)),
        )
        entry.commit_numeric = command
        return entry

    def _build_ui(self):
        main = ttk.Frame(self.root, style="Workspace.TFrame", padding=20)
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure(0, weight=11, uniform="content")
        main.columnconfigure(1, weight=9, uniform="content")
        main.rowconfigure(1, weight=1)

        hero_outer, hero = self._create_card(main, padding=(22, 18))
        hero_outer.grid(row=0, column=0, columnspan=2, sticky=tk.EW, pady=(0, 14))
        hero.columnconfigure(0, weight=1)
        ttk.Label(
            hero, text="DOCUMENT RASTERIZER", style="Kicker.TLabel"
        ).grid(row=0, column=0, sticky=tk.W)
        ttk.Label(hero, text="文档逐页导出", style="HeroTitle.TLabel").grid(
            row=1, column=0, sticky=tk.W, pady=(2, 0)
        )
        ttk.Label(
            hero,
            text="把 PDF 或 Word 变成清晰、可分享的逐页图片。",
            style="HeroSub.TLabel",
        ).grid(row=2, column=0, sticky=tk.W, pady=(7, 0))

        pipeline = ttk.Frame(hero, style="Card.TFrame")
        pipeline.grid(row=0, column=1, rowspan=3, sticky=tk.E)
        ttk.Label(pipeline, text="INPUT", style="Micro.TLabel").grid(
            row=0, column=0, sticky=tk.W, pady=(0, 5)
        )
        ttk.Label(pipeline, text="OUTPUT", style="Micro.TLabel").grid(
            row=0, column=2, sticky=tk.W, pady=(0, 5)
        )
        ttk.Label(
            pipeline, textvariable=self.input_type_var, style="InputBadge.TLabel"
        ).grid(row=1, column=0)
        ttk.Label(pipeline, text="→", style="Arrow.TLabel").grid(
            row=1, column=1, padx=12
        )
        ttk.Label(
            pipeline, textvariable=self.output_type_var, style="OutputBadge.TLabel"
        ).grid(row=1, column=2)

        # 左栏：输入与输出位置
        file_outer, file_card = self._create_card(main)
        file_outer.grid(row=1, column=0, sticky=tk.NSEW, padx=(0, 7))
        file_card.columnconfigure(0, weight=1)
        self._section_heading(
            file_card, "01 · DOCUMENT", "选择文档", "来源文件与图片保存位置"
        )

        input_header = ttk.Frame(file_card, style="Card.TFrame")
        input_header.grid(row=3, column=0, sticky=tk.EW, pady=(22, 7))
        input_header.columnconfigure(0, weight=1)
        ttk.Label(input_header, text="来源文件", style="FieldLabel.TLabel").grid(
            row=0, column=0, sticky=tk.W
        )
        ttk.Label(
            input_header, textvariable=self.input_status_var, style="Secondary.TLabel"
        ).grid(row=0, column=1, sticky=tk.E)
        pdf_row = ttk.Frame(file_card, style="Card.TFrame")
        pdf_row.grid(row=4, column=0, sticky=tk.EW)
        pdf_row.columnconfigure(0, weight=1)
        ttk.Entry(pdf_row, textvariable=self.input_path).grid(
            row=0, column=0, sticky=tk.EW, padx=(0, 8)
        )
        ttk.Button(pdf_row, text="浏览…", command=self._browse_input,
                   style="Browse.TButton").grid(row=0, column=1)
        ttk.Label(
            file_card, text="支持 PDF · DOC · DOCX",
            style="Secondary.TLabel",
        ).grid(row=5, column=0, sticky=tk.W, pady=(6, 0))

        ttk.Separator(file_card).grid(
            row=6, column=0, sticky=tk.EW, pady=(22, 20)
        )
        ttk.Label(file_card, text="保存到", style="FieldLabel.TLabel").grid(
            row=7, column=0, sticky=tk.W, pady=(0, 7)
        )
        out_row = ttk.Frame(file_card, style="Card.TFrame")
        out_row.grid(row=8, column=0, sticky=tk.EW)
        out_row.columnconfigure(0, weight=1)
        ttk.Entry(out_row, textvariable=self.out_dir).grid(
            row=0, column=0, sticky=tk.EW, padx=(0, 8)
        )
        ttk.Button(out_row, text="浏览…", command=self._browse_output,
                   style="Browse.TButton").grid(row=0, column=1)
        ttk.Label(
            file_card, text="选择来源文件后会自动生成默认文件夹。",
            style="Secondary.TLabel",
        ).grid(row=9, column=0, sticky=tk.W, pady=(6, 0))

        ttk.Separator(file_card).grid(
            row=10, column=0, sticky=tk.EW, pady=(22, 14)
        )
        row_enh = ttk.Frame(file_card, style="Card.TFrame")
        row_enh.grid(row=11, column=0, sticky=tk.EW)
        row_enh.columnconfigure(0, weight=1)
        self.enhance_cb = tk.Checkbutton(
            row_enh,
            textvariable=self.enhance_option_text,
            variable=self.enhance_var,
            command=self._on_enhance_toggle,
            indicatoron=False,
            anchor=tk.W,
            justify=tk.LEFT,
            bg=self.colors["field"],
            fg=self.colors["ink"],
            selectcolor=self.colors["accent_soft"],
            activebackground=self.colors["accent_soft"],
            activeforeground=self.colors["accent"],
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors["border"],
            highlightcolor=self.colors["accent"],
            font=("Microsoft YaHei UI", 10, "bold"),
            padx=14,
            pady=11,
            cursor="hand2",
            takefocus=True,
        )
        self.enhance_cb.grid(row=0, column=0, sticky=tk.EW)
        self.enhance_settings_btn = ttk.Button(
            row_enh, text="调整参数…", style="Browse.TButton",
            command=self._open_enhance_settings, state=tk.DISABLED,
        )
        self.enhance_settings_btn.grid(
            row=0, column=1, sticky=tk.NS, padx=(8, 0), ipadx=8
        )

        file_card.rowconfigure(12, weight=1)
        word_note = ttk.Frame(file_card, style="Callout.TFrame", padding=(12, 10))
        word_note.grid(row=13, column=0, sticky=tk.EW, pady=(22, 0))
        ttk.Label(
            word_note, text="WORD 转换", style="CalloutTitle.TLabel"
        ).grid(row=0, column=0, sticky=tk.W)
        ttk.Label(
            word_note,
            text="只读打开原文件；中间 PDF 会自动清理。需要本机安装 Microsoft Word。",
            style="CalloutText.TLabel",
            wraplength=500,
            justify=tk.LEFT,
        ).grid(row=1, column=0, sticky=tk.W, pady=(4, 0))

        # 右栏：图像规格与页面
        opt_outer, opt_card = self._create_card(main)
        opt_outer.grid(row=1, column=1, sticky=tk.NSEW, padx=(7, 0))
        opt_card.columnconfigure(0, weight=1)
        self._section_heading(
            opt_card, "02 · IMAGE", "设置输出", ""
        )

        ttk.Label(opt_card, text="图片格式", style="FieldLabel.TLabel").grid(
            row=3, column=0, sticky=tk.W, pady=(16, 6)
        )
        row_fmt = ttk.Frame(opt_card, style="Card.TFrame")
        row_fmt.grid(row=4, column=0, sticky=tk.EW)
        self.format_buttons = []
        for column, fmt in enumerate(FORMAT_KEYS):
            row_fmt.columnconfigure(column, weight=1, uniform="format")
            button = self._choice_radio(
                row_fmt, fmt, self.fmt_var, fmt, self._on_fmt_change
            )
            button.grid(
                row=0, column=column, sticky=tk.EW,
                padx=(0 if column == 0 else 3, 0 if column == 4 else 3),
            )
            self.format_buttons.append((fmt, button))
        self.fmt_desc = ttk.Label(row_fmt, text="", style="Secondary.TLabel")
        self.fmt_desc.grid(
            row=1, column=0, columnspan=5, sticky=tk.W, pady=(7, 0)
        )
        self._update_fmt_desc()
        self.fmt_desc.grid_remove()

        ttk.Separator(opt_card).grid(
            row=5, column=0, sticky=tk.EW, pady=(14, 12)
        )
        ttk.Label(opt_card, text="清晰度", style="FieldLabel.TLabel").grid(
            row=6, column=0, sticky=tk.W, pady=(0, 6)
        )
        row_dpi = ttk.Frame(opt_card, style="Card.TFrame")
        row_dpi.grid(row=7, column=0, sticky=tk.EW)
        self.dpi_buttons = []
        for column, val in enumerate((150, 200, 300, 400)):
            row_dpi.columnconfigure(column, weight=1, uniform="dpi")
            button = self._choice_radio(
                row_dpi, f"{val} DPI", self.dpi_var, val, self._on_dpi_change
            )
            button.grid(
                row=0, column=column, sticky=tk.EW,
                padx=(0 if column == 0 else 3, 0 if column == 3 else 3),
            )
            self.dpi_buttons.append((val, button))

        self.quality_frame = ttk.Frame(opt_card, style="Card.TFrame")
        self.quality_frame.grid(row=8, column=0, sticky=tk.EW, pady=(8, 0))
        self.quality_frame.grid_remove()
        self.quality_frame.columnconfigure(0, weight=1)
        ttk.Label(
            self.quality_frame, text="压缩质量", style="Secondary.TLabel"
        ).grid(row=0, column=0, sticky=tk.W)
        self.quality_scale = tk.Scale(
            self.quality_frame,
            from_=10,
            to=100,
            variable=self.quality_var,
            orient=tk.HORIZONTAL,
            resolution=1,
            showvalue=False,
            bg=self.colors["card"],
            troughcolor="#DCE5F2",
            activebackground=self.colors["accent"],
            sliderrelief=tk.FLAT,
            sliderlength=32,
            width=16,
            highlightthickness=0,
        )
        self.quality_scale.grid(
            row=1, column=0, columnspan=2, sticky=tk.EW, pady=(6, 0)
        )
        self.quality_entry = self._numeric_entry(
            self.quality_frame,
            self.quality_input_var,
            start=10,
            end=100,
            increment=1,
            command=self._commit_quality_input,
        )
        self.quality_entry.grid(row=0, column=1, sticky=tk.E)

        ttk.Separator(opt_card).grid(
            row=9, column=0, sticky=tk.EW, pady=(14, 12)
        )
        ttk.Label(opt_card, text="页面范围", style="FieldLabel.TLabel").grid(
            row=10, column=0, sticky=tk.W, pady=(0, 6)
        )
        row_page = ttk.Frame(opt_card, style="Card.TFrame")
        row_page.grid(row=11, column=0, sticky=tk.EW)
        row_page.columnconfigure(0, weight=1, uniform="page")
        row_page.columnconfigure(1, weight=1, uniform="page")
        self.page_buttons = []
        self.all_radio = self._choice_radio(
            row_page, "全部页面", self.page_range_var, "all"
        )
        self.all_radio.grid(row=0, column=0, sticky=tk.EW, padx=(0, 4))
        self.page_buttons.append(("all", self.all_radio))
        self.custom_radio = self._choice_radio(
            row_page, "指定页面", self.page_range_var, "custom"
        )
        self.custom_radio.grid(row=0, column=1, sticky=tk.EW, padx=(4, 0))
        self.page_buttons.append(("custom", self.custom_radio))

        custom_row = ttk.Frame(row_page, style="Card.TFrame")
        custom_row.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(8, 0))
        custom_row.columnconfigure(1, weight=1)
        ttk.Label(
            custom_row, text="页码", style="Secondary.TLabel"
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        self.custom_entry = ttk.Entry(
            custom_row, textvariable=self.custom_pages_var
        )
        self.custom_entry.grid(row=0, column=1, sticky=tk.EW)
        ttk.Label(
            custom_row, text="例如 1,3,5-10", style="Secondary.TLabel"
        ).grid(row=0, column=2, sticky=tk.E, padx=(8, 0))

        # 底部：状态与主操作
        action_outer, action_card = self._create_card(main, padding=(18, 14))
        action_outer.grid(
            row=2, column=0, columnspan=2, sticky=tk.EW, pady=(14, 0)
        )
        action_card.columnconfigure(0, weight=1)
        self.summary_var = tk.StringVar(value="PNG · 200 DPI · 全部页面")
        self.status_var = tk.StringVar(value="等待选择文档")
        ttk.Label(
            action_card, textvariable=self.status_var, style="Status.TLabel"
        ).grid(row=0, column=0, sticky=tk.W)
        ttk.Label(
            action_card, textvariable=self.summary_var, style="Summary.TLabel"
        ).grid(row=1, column=0, sticky=tk.W, pady=(4, 0))
        self.btn_convert = ttk.Button(
            action_card, text="开始导出", command=self._start_convert,
            style="Action.TButton",
        )
        self.btn_convert.grid(row=0, column=1, rowspan=2, sticky=tk.E, padx=(18, 0))
        self.progress = ttk.Progressbar(action_card, mode="determinate")
        self.progress.grid(
            row=2, column=0, columnspan=2, sticky=tk.EW, pady=(13, 0)
        )

        self._on_page_mode_change()
        self._refresh_summary()
        self._refresh_choice_styles()
        self._on_input_path_change()

    def _apply_initial_window_size(self):
        self.root.update_idletasks()
        req_w = self.root.winfo_reqwidth()
        req_h = self.root.winfo_reqheight()
        width = max(1040, req_w + 24)
        height = max(690, req_h + 24)
        self.root.minsize(920, 640)
        self._center_window(width, height)

    def _center_window(self, w: int, h: int):
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w = min(w, sw - 80)
        h = min(h, sh - 80)
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    # ------------------------------------------------------------------
    # 格式联动
    # ------------------------------------------------------------------
    def _refresh_choice_styles(self):
        """刷新块状选项的选中边框、底色和文字颜色。"""
        groups = (
            (getattr(self, "format_buttons", []), self.fmt_var.get()),
            (getattr(self, "dpi_buttons", []), self.dpi_var.get()),
            (getattr(self, "page_buttons", []), self.page_range_var.get()),
        )
        for buttons, selected in groups:
            for value, button in buttons:
                is_selected = value == selected
                button.config(
                    bg=(
                        self.colors["accent_soft"]
                        if is_selected else self.colors["field"]
                    ),
                    fg=(
                        self.colors["accent"]
                        if is_selected else self.colors["ink"]
                    ),
                    highlightbackground=(
                        self.colors["accent"]
                        if is_selected else self.colors["border"]
                    ),
                )

        if hasattr(self, "enhance_cb"):
            enabled = self.enhance_var.get()
            self.enhance_cb.config(
                bg=(
                    self.colors["accent_soft"]
                    if enabled else self.colors["field"]
                ),
                fg=(
                    self.colors["accent"]
                    if enabled else self.colors["ink"]
                ),
                highlightbackground=(
                    self.colors["accent"]
                    if enabled else self.colors["border"]
                ),
            )

    def _on_fmt_change(self, event=None):
        self.output_type_var.set(self.fmt_var.get())
        self._update_fmt_desc()
        self._update_fmt_quality_visibility()
        self._refresh_choice_styles()
        self._refresh_summary()

    def _on_dpi_change(self):
        self._refresh_choice_styles()
        self._refresh_summary()

    def _on_quality_change(self):
        self.quality_input_var.set(str(self.quality_var.get()))
        if hasattr(self, "summary_var"):
            self._refresh_summary()

    def _commit_quality_input(self):
        """提交质量输入框，纠正非数字或超出可用范围的值。"""
        try:
            quality = round(float(self.quality_input_var.get().strip()))
        except ValueError:
            quality = self.quality_var.get()
        quality = max(10, min(100, quality))
        self.quality_var.set(quality)
        self.quality_input_var.set(str(quality))

    def _update_fmt_desc(self):
        fmt = self.fmt_var.get()
        if fmt in SUPPORTED_FORMATS:
            self.fmt_desc.config(text=SUPPORTED_FORMATS[fmt][2])

    def _update_fmt_quality_visibility(self):
        """JPEG/WEBP 显示质量滑动条，其他格式隐藏"""
        fmt = self.fmt_var.get()
        if fmt in ("JPEG", "WEBP"):
            self.quality_frame.grid()
        else:
            self.quality_frame.grid_remove()
        self._refresh_summary()

    # ------------------------------------------------------------------
    # 事件
    # ------------------------------------------------------------------
    def _on_input_path_change(self):
        """根据路径即时更新输入类型与就绪状态。"""
        raw = self.input_path.get().strip()
        if not raw:
            self.input_type_var.set("PDF / WORD")
            self.input_status_var.set("等待选择文档")
            if hasattr(self, "btn_convert") and not getattr(self, "_running", False):
                self.btn_convert.config(state=tk.DISABLED)
            if hasattr(self, "status_var") and not getattr(self, "_running", False):
                self.status_var.set("等待选择文档")
            return

        path = Path(raw)
        ext = path.suffix.lower()
        if ext == ".pdf":
            doc_type = "PDF"
        elif ext in {".doc", ".docx"}:
            doc_type = "WORD"
        else:
            doc_type = "未知格式"

        self.input_type_var.set(doc_type)
        if ext not in SUPPORTED_INPUT_EXTENSIONS:
            self.input_status_var.set("格式不受支持")
            if hasattr(self, "btn_convert") and not getattr(self, "_running", False):
                self.btn_convert.config(state=tk.DISABLED)
            if hasattr(self, "status_var") and not getattr(self, "_running", False):
                self.status_var.set("请选择 PDF、DOC 或 DOCX")
        elif path.is_file():
            self.input_status_var.set(f"{doc_type} · 已选择")
            if hasattr(self, "btn_convert") and not getattr(self, "_running", False):
                self.btn_convert.config(state=tk.NORMAL)
            if hasattr(self, "status_var") and not getattr(self, "_running", False):
                self.status_var.set("准备就绪")
        else:
            self.input_status_var.set("路径待确认")
            if hasattr(self, "btn_convert") and not getattr(self, "_running", False):
                self.btn_convert.config(state=tk.DISABLED)
            if hasattr(self, "status_var") and not getattr(self, "_running", False):
                self.status_var.set("请确认来源文件路径")

    def _on_page_mode_change(self):
        """仅在自定义页面模式下启用页码输入框。"""
        if hasattr(self, "custom_entry"):
            state = tk.NORMAL if self.page_range_var.get() == "custom" else tk.DISABLED
            self.custom_entry.config(state=state)
        self._refresh_choice_styles()
        if hasattr(self, "summary_var"):
            self._refresh_summary()

    def _on_enhance_toggle(self):
        """启用扫描件优化，并联动高级参数按钮。"""
        enabled = self.enhance_var.get()
        state = tk.NORMAL if enabled else tk.DISABLED
        self.enhance_settings_btn.config(state=state)
        self.enhance_option_text.set(
            "优化扫描件 · 已开启\n锐化 · 去黄 · 对比度"
            if enabled
            else "优化扫描件\n锐化 · 去黄 · 对比度"
        )
        self._refresh_choice_styles()
        self._refresh_summary()

    def _open_enhance_settings(self):
        """在独立窗口中调整低频使用的扫描件参数。"""
        dialog = tk.Toplevel(self.root)
        dialog.title("扫描件优化参数")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.configure(bg=self.colors["workspace"])

        outer, card = self._create_card(dialog, padding=(20, 18))
        outer.pack(fill=tk.BOTH, expand=True, padx=14, pady=14)
        card.columnconfigure(1, weight=1)
        ttk.Label(card, text="SCAN ENHANCEMENT", style="Kicker.TLabel").grid(
            row=0, column=0, columnspan=3, sticky=tk.W
        )
        ttk.Label(card, text="优化扫描件", style="CardTitle.TLabel").grid(
            row=1, column=0, columnspan=3, sticky=tk.W, pady=(3, 0)
        )
        ttk.Label(
            card,
            text="适合发黄、文字边缘模糊或对比度不足的扫描页面。",
            style="Secondary.TLabel",
        ).grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(5, 16))

        controls = (
            ("锐化", "增强文字与线条边缘", self.enhance_sharpness, 0, 200),
            ("去黄", "自动拉伸色阶，减轻纸张底色", self.enhance_cutoff, 0, 10),
            ("对比度", "拉开文字与背景的明暗差异", self.enhance_contrast, 1.0, 2.0),
        )
        trace_handles = []
        for index, (label, hint, variable, start, end) in enumerate(controls):
            row = 3 + index * 2
            ttk.Label(card, text=label, style="FieldLabel.TLabel").grid(
                row=row, column=0, sticky=tk.W
            )
            ttk.Label(card, text=hint, style="Secondary.TLabel").grid(
                row=row, column=1, sticky=tk.W, padx=(10, 0)
            )
            is_float = isinstance(variable, tk.DoubleVar)
            input_var = tk.StringVar(
                value=f"{variable.get():.2f}" if is_float else str(variable.get())
            )

            def commit_value(
                var=variable,
                text_var=input_var,
                minimum=start,
                maximum=end,
                use_float=is_float,
            ):
                try:
                    parsed = float(text_var.get().strip())
                except ValueError:
                    parsed = var.get()
                parsed = max(minimum, min(maximum, parsed))
                if use_float:
                    var.set(round(parsed, 2))
                    text_var.set(f"{var.get():.2f}")
                else:
                    var.set(round(parsed))
                    text_var.set(str(var.get()))

            entry = self._numeric_entry(
                card,
                input_var,
                start=start,
                end=end,
                increment=0.01 if is_float else 1,
                command=commit_value,
            )
            entry.grid(row=row, column=2, sticky=tk.E)
            tk.Scale(
                card,
                from_=start,
                to=end,
                variable=variable,
                orient=tk.HORIZONTAL,
                resolution=0.01 if is_float else 1,
                showvalue=False,
                bg=self.colors["card"],
                troughcolor="#DCE5F2",
                activebackground=self.colors["accent"],
                sliderrelief=tk.FLAT,
                sliderlength=34,
                width=16,
                highlightthickness=0,
            ).grid(
                row=row + 1, column=0, columnspan=3, sticky=tk.EW,
                pady=(7, 16),
            )

            def sync_input(*_, var=variable, text_var=input_var, use_float=is_float):
                text_var.set(f"{var.get():.2f}" if use_float else str(var.get()))

            trace_handles.append(
                (variable, variable.trace_add("write", sync_input))
            )

        def close_dialog():
            for variable, trace_id in trace_handles:
                variable.trace_remove("write", trace_id)
            dialog.destroy()

        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=9, column=0, columnspan=3, sticky=tk.EW, pady=(2, 0))
        ttk.Button(
            actions, text="恢复默认", style="Browse.TButton",
            command=self._enhance_reset_default,
        ).pack(side=tk.LEFT)
        ttk.Button(
            actions, text="完成", style="Action.TButton", command=close_dialog,
        ).pack(side=tk.RIGHT)

        dialog.update_idletasks()
        width = max(520, dialog.winfo_reqwidth())
        height = dialog.winfo_reqheight()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - width) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - height) // 2
        dialog.geometry(f"{width}x{height}+{max(0, x)}+{max(0, y)}")
        dialog.protocol("WM_DELETE_WINDOW", close_dialog)
        dialog.grab_set()
        dialog.focus_set()

    def _enhance_reset_default(self):
        """恢复增强参数为默认值"""
        self.enhance_sharpness.set(80)
        self.enhance_cutoff.set(2)
        self.enhance_contrast.set(1.15)
        self._refresh_summary()

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="选择 PDF 或 Word 文件",
            filetypes=[
                ("支持的文档", "*.pdf *.doc *.docx"),
                ("PDF 文件", "*.pdf"),
                ("Word 文件", "*.doc *.docx"),
                ("所有文件", "*.*"),
            ],
        )
        if not path:
            return
        self.input_path.set(path)

        # 自动设置输出目录
        p = Path(path)
        default_out = p.parent / f"{p.stem}_图片"
        self.out_dir.set(str(default_out))
        self._refresh_summary()

    def _browse_output(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.out_dir.set(path)
            self._refresh_summary()

    def _refresh_summary(self):
        if self.page_range_var.get() == "all":
            page_text = "全部页面"
        else:
            custom = self.custom_pages_var.get().strip()
            page_text = f"页面 {custom}" if custom else "自定义页面"
        parts = [self.fmt_var.get(), f"{self.dpi_var.get()} DPI", page_text]
        if self.fmt_var.get() in ("JPEG", "WEBP"):
            parts.append(f"质量 {self.quality_var.get()}")
        if self.enhance_var.get():
            parts.append("增强已开启")
        self.output_type_var.set(self.fmt_var.get())
        self.summary_var.set(" · ".join(parts))

    def _start_convert(self):
        if self._running:
            return

        input_path = self.input_path.get().strip()
        out = self.out_dir.get().strip()

        if not input_path:
            messagebox.showwarning("提示", "请先选择 PDF 或 Word 文件")
            return
        if not os.path.isfile(input_path):
            messagebox.showerror("错误", "输入文件不存在")
            return
        if Path(input_path).suffix.lower() not in SUPPORTED_INPUT_EXTENSIONS:
            messagebox.showerror("错误", "仅支持 PDF、DOC、DOCX 文件")
            return
        if not out:
            messagebox.showwarning("提示", "请选择输出目录")
            return

        page_range = None
        if self.page_range_var.get() == "custom":
            page_range = self.custom_pages_var.get().strip()
        if self.page_range_var.get() == "custom" and not page_range:
            messagebox.showwarning("提示", "请填写页面范围，或选择「全部」")
            return

        self._running = True
        self.btn_convert.config(state=tk.DISABLED, text="导出中...")
        self.progress["value"] = 0
        self.status_var.set("准备中...")
        self.root.update()

        fmt = self.fmt_var.get()
        dpi = self.dpi_var.get()
        quality = self.quality_var.get()
        enhance = self.enhance_var.get()
        sharpness = self.enhance_sharpness.get()
        cutoff = self.enhance_cutoff.get()
        contrast = self.enhance_contrast.get()

        t = threading.Thread(
            target=self._do_convert,
            args=(input_path, out, fmt, dpi, quality, page_range, enhance,
                  sharpness, cutoff, contrast),
            daemon=True,
        )
        t.start()

    def _do_convert(self, input_path, out, fmt, dpi, quality, page_range,
                    enhance, sharpness, cutoff, contrast):
        try:
            generated = document_to_images(
                input_path, out, fmt=fmt, dpi=dpi, quality=quality,
                page_range=page_range, progress_cb=self._on_progress,
                image_enhance=enhance,
                enhance_sharpness=sharpness,
                enhance_cutoff=cutoff,
                enhance_contrast=contrast,
            )
            self.root.after(0, self._on_success, out, generated)
        except Exception as e:
            self.root.after(0, self._on_error, str(e))

    def _on_progress(self, current, total, stage):
        self.root.after(0, self._update_ui, current, total, stage)

    def _update_ui(self, current, total, stage):
        if stage:
            self.status_var.set(stage)
        if total > 0:
            pct = int(current / total * 100)
            self.progress["value"] = pct
            self.progress["maximum"] = 100
            self.status_var.set(f"正在导出第 {current}/{total} 张...")
        self.root.update()

    def _on_success(self, out_dir, generated):
        count = len(generated)
        # 计算总大小
        total_size = sum(os.path.getsize(f) for f in generated)
        size_str = self._fmt_size(total_size)

        self.status_var.set(f"✓ 导出完成！共 {count} 张图片 ({size_str})")
        self.progress["value"] = 100
        self._reset_btn()

        if messagebox.askyesno("完成", f"导出成功！\n共 {count} 张图片\n目录: {out_dir}\n总大小: {size_str}\n\n是否打开输出文件夹？"):
            try:
                os.startfile(out_dir)
            except Exception:
                pass

    def _on_error(self, msg):
        friendly = msg
        if "No such file" in msg or "not found" in msg.lower():
            friendly = f"文件未找到。\n{msg}"
        elif "Permission denied" in msg or "拒绝访问" in msg:
            friendly = f"访问被拒绝，请检查文件是否被占用。\n{msg}"

        self.status_var.set("✗ 导出失败")
        self.progress["value"] = 0
        self._reset_btn()
        messagebox.showerror("导出失败", friendly)

    def _reset_btn(self):
        self._running = False
        input_path = Path(self.input_path.get().strip())
        state = (
            tk.NORMAL
            if input_path.is_file()
            and input_path.suffix.lower() in SUPPORTED_INPUT_EXTENSIONS
            else tk.DISABLED
        )
        self.btn_convert.config(state=state, text="开始导出")

    @staticmethod
    def _fmt_size(byte: int) -> str:
        for unit in ("B", "KB", "MB", "GB"):
            if byte < 1024:
                return f"{byte:.1f} {unit}"
            byte /= 1024
        return f"{byte:.1f} TB"

    # ------------------------------------------------------------------
    # 启动
    # ------------------------------------------------------------------
    def run(self):
        self.root.mainloop()
