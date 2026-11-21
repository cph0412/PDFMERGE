"""PDF 工具箱 — 主界面（v2 浅色现代风，基于 sv-ttk）。"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import re
from pathlib import Path

from tkinterdnd2 import DND_FILES

try:
    import sv_ttk  # 随打包内置；缺失时回退到原生主题
    _HAS_SV_TTK = True
except Exception:
    _HAS_SV_TTK = False

from pdf_ops import (
    merge_pdfs,
    split_pdf,
    extract_pages,
    rotate_pages,
    merge_2up,
    merge_4up,
    get_pdf_info,
    PDFError,
)
from excel_ops import (
    merge_to_sheets,
    merge_to_one_sheet,
    merge_copy_sheets,
    ExcelError,
    ExcelCancelled,
)
from preview import PreviewWindow

# ── 字体 ──
FONT_UI = ("Microsoft YaHei UI", 10)
FONT_SMALL = ("Microsoft YaHei UI", 9)
FONT_CARD_TITLE = ("Microsoft YaHei UI", 11, "bold")
FONT_TITLE = ("Microsoft YaHei UI", 19, "bold")
FONT_SUB = ("Microsoft YaHei UI", 9)

# ── 窗口 ──
WIN_WIDTH = 1040
WIN_HEIGHT = 720

# ── 浅色配色（sv-ttk light 同源；用于 tk 原生控件）──
C_WIN = "#FAFAFA"
C_TEXT = "#1E222A"
C_MUTED = "#6E7681"
C_BORDER = "#E4E7EB"
C_SELECT_BG = "#005FB8"
C_SELECT_FG = "#FFFFFF"

# 标签页索引 → 底部执行按钮文本
TAB_LABELS = ["开始合并", "开始拆分", "开始提取", "开始旋转", "开始合并 Excel"]


class PDFToolApp:
    """PDF 工具箱主应用。"""

    def __init__(self, root):
        self.root = root
        self.root.title("PDF 工具箱")
        self.root.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}")
        self.root.minsize(900, 620)
        self.root.configure(bg=C_WIN)

        # ── 主题：sv-ttk 浅色，失败则退回 clam ──
        if _HAS_SV_TTK:
            try:
                sv_ttk.set_theme("light")
            except Exception:
                pass
        style = ttk.Style()
        if style.theme_use().startswith("sun-valley"):
            pass  # sv-ttk 已生效
        else:
            try:
                style.theme_use("clam")
            except Exception:
                pass
        self._setup_styles()

        # ── 顶部品牌栏 ──
        self._build_header()

        # ── 底部操作坞（先打包到底部，避免被后打包的内容区挤出可视区）──
        self._bottom_frame = ttk.Frame(root, padding=(4, 0, 4, 4))
        self._bottom_frame.pack(side="bottom", fill="x", padx=14, pady=(0, 12))

        ttk.Separator(self._bottom_frame, orient="horizontal").pack(
            fill="x", pady=(0, 10)
        )

        # 进度条（任务运行时为不确定模式）
        self._progress = ttk.Progressbar(self._bottom_frame, mode="indeterminate")
        self._progress.pack(fill="x", pady=(0, 8))

        dock_row = ttk.Frame(self._bottom_frame)
        dock_row.pack(fill="x")

        self._status_var = tk.StringVar(value="就绪")
        ttk.Label(
            dock_row, textvariable=self._status_var,
            style="App.Hint.TLabel",
        ).pack(side="left")

        # 取消按钮（支持取消的操作运行期间可用）
        self._btn_cancel = ttk.Button(
            dock_row, text="取消", style="App.Subtle.TButton",
            command=self._on_cancel_click, state="disabled",
        )
        self._btn_cancel.pack(side="right", padx=(8, 0))

        # 主执行按钮
        self._btn_execute = ttk.Button(
            dock_row, text="开始合并", style="App.Accent.TButton",
            command=self._run_operation,
        )
        self._btn_execute.pack(side="right")

        # ── 内容区（Notebook）最后打包，填充底栏与顶栏之间的空间 ──
        self._notebook = ttk.Notebook(root)
        self._notebook.pack(fill="both", expand=True, padx=16, pady=(8, 0))

        # ── 构建标签页 ──
        self._build_merge_tab()
        self._build_split_tab()
        self._build_extract_tab()
        self._build_rotate_tab()
        self._build_excel_tab()

        self._notebook.bind("<<NotebookTabChanged>>", self._on_tab_change)

        # 状态
        self._file_list: list[str] = []
        self._excel_file_list: list[str] = []
        self._preview_win: PreviewWindow | None = None
        # 取消支持（Excel 合并运行期间置位）
        self._cancel_event: threading.Event | None = None
        self._cancel_armed = False

    # ═══════════════════════════════════════════════════════
    # 视觉基础：样式 / 头部 / 通用构件
    # ═══════════════════════════════════════════════════════

    def _setup_styles(self) -> None:
        style = ttk.Style()
        style.configure("TNotebook", background=C_WIN, borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            font=FONT_UI, padding=(18, 9),
            borderwidth=0,
        )
        # 标签页文字/选中色交给主题；这里统一一下尺寸
        style.configure("App.TLabel", font=FONT_UI)
        style.configure("App.CardTitle.TLabel", font=FONT_CARD_TITLE)
        style.configure("App.Hint.TLabel", font=FONT_SMALL, foreground=C_MUTED)
        style.configure("App.TButton", font=FONT_UI)
        style.configure("App.Accent.TButton", font=FONT_UI)
        style.configure("App.Subtle.TButton", font=FONT_UI)
        style.configure("App.TEntry", font=FONT_UI)
        style.configure("App.TSpinbox", font=FONT_UI)
        style.configure("App.TCombobox", font=FONT_UI)

    def _build_header(self) -> None:
        """顶部品牌栏：标题 + 用途说明 + 作者信息。"""
        hdr = tk.Frame(self.root, bg=C_WIN)
        hdr.pack(fill="x", padx=22, pady=(14, 2))

        brand = tk.Frame(hdr, bg=C_WIN)
        brand.pack(side="left")
        tk.Label(
            brand, text="PDF 工具箱", font=FONT_TITLE,
            fg=C_TEXT, bg=C_WIN,
        ).pack(anchor="w")
        tk.Label(
            brand,
            text="文件合并 · 拆分 · 提取 · 旋转      Excel 合并 / 整表复制（保留格式，便于打印）",
            font=FONT_SUB, fg=C_MUTED, bg=C_WIN,
        ).pack(anchor="w", pady=(3, 0))

        # 右侧：创作者信息 + 版本
        info = tk.Frame(hdr, bg=C_WIN)
        info.pack(side="right", anchor="e", padx=(18, 0))
        tk.Label(
            info, text="程鹏昊 · 603381208@qq.com", font=FONT_SMALL,
            fg=C_MUTED, bg=C_WIN,
        ).pack(anchor="e")
        tk.Label(
            info, text="v2.2", font=FONT_SMALL,
            fg="#B0B6BF", bg=C_WIN,
        ).pack(anchor="e", pady=(2, 0))

    @staticmethod
    def _page(tab) -> ttk.Frame:
        """标签页内容容器（统一内边距）。"""
        page = ttk.Frame(tab, padding=(18, 14, 18, 8))
        page.pack(fill="both", expand=True)
        return page

    def _card(self, parent, title: str | None = None,
              expand: bool = False) -> ttk.Frame:
        """卡片：圆角浅底容器，返回卡片内部的内容容器。

        expand=True 时卡片在父容器里纵向铺满（用于文件列表大卡片）。
        """
        card = ttk.Frame(parent, style="Card.TFrame", padding=(16, 12))
        if title:
            ttk.Label(
                card, text=title, style="App.CardTitle.TLabel",
            ).pack(anchor="w", pady=(0, 2))
        body = ttk.Frame(card)
        body.pack(fill="both", expand=True)
        card.pack(fill="both", expand=expand, pady=(0, 12))
        return body

    @staticmethod
    def _path_row(parent, var, browse_cmd, placeholder: str = "浏览…") -> None:
        """一行「只读路径输入框 + 浏览按钮」。"""
        row = ttk.Frame(parent)
        row.pack(fill="x")
        entry = ttk.Entry(row, textvariable=var, style="App.TEntry",
                          state="readonly")
        entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(row, text=placeholder, style="App.TButton",
                   command=browse_cmd).pack(side="right")

    def _page_hint(self, parent, text: str) -> None:
        ttk.Label(
            parent, text=text, style="App.Hint.TLabel",
        ).pack(side="left", padx=(2, 0))

    # ═══════════════════════════════════════════════════════
    # 合并标签页
    # ═══════════════════════════════════════════════════════

    def _build_merge_tab(self) -> None:
        tab = ttk.Frame(self._notebook)
        self._notebook.add(tab, text="合并 PDF")
        page = self._page(tab)

        body = self._card(page, "文件列表（按合并顺序排列，可直接拖入 PDF）",
                          expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        # 工具行
        toolbar = ttk.Frame(body)
        toolbar.grid(row=0, column=0, sticky="ew", pady=(4, 10))

        for txt, cmd, st in (
            ("＋ 添加文件", self._add_files, "App.TButton"),
            ("✕ 移除所选", self._remove_selected, "App.TButton"),
            ("↑ 上移", self._move_up, "App.Subtle.TButton"),
            ("↓ 下移", self._move_down, "App.Subtle.TButton"),
        ):
            ttk.Button(toolbar, text=txt, style=st, command=cmd).pack(
                side="left", padx=(0, 6)
            )

        # 右侧：预览 + 拼版选项
        self._merge_2up_var = tk.BooleanVar(value=True)
        self._merge_4up_var = tk.BooleanVar(value=False)
        right = ttk.Frame(toolbar)
        right.pack(side="right")
        ttk.Checkbutton(
            right, text="二合一", variable=self._merge_2up_var,
            command=self._on_2up_toggle,
        ).pack(side="left", padx=(0, 12))
        ttk.Checkbutton(
            right, text="四合一", variable=self._merge_4up_var,
            command=self._on_4up_toggle,
        ).pack(side="left", padx=(0, 12))
        ttk.Button(right, text="预览", style="App.Subtle.TButton",
                   command=self._open_preview).pack(side="left")
        ttk.Label(
            right, text="将多页拼到一张 A4 上",
            style="App.Hint.TLabel",
        ).pack(side="left", padx=(10, 0))

        # 列表区
        listwrap = ttk.Frame(body)
        listwrap.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(listwrap, orient="vertical")
        scrollbar.pack(side="right", fill="y")

        self._merge_listbox = tk.Listbox(
            listwrap,
            yscrollcommand=scrollbar.set,
            font=FONT_UI,
            selectmode="extended",
            bg="#FFFFFF", fg=C_TEXT,
            selectbackground=C_SELECT_BG, selectforeground=C_SELECT_FG,
            relief="flat", borderwidth=0, activestyle="none",
            highlightthickness=1,
            highlightbackground=C_BORDER, highlightcolor=C_SELECT_BG,
            exportselection=False,
        )
        self._merge_listbox.pack(fill="both", expand=True, side="left")
        self._merge_listbox.bind("<<ListboxSelect>>", self._on_merge_select)
        self._merge_listbox.drop_target_register(DND_FILES)
        self._merge_listbox.dnd_bind("<<Drop>>", self._on_drop)
        scrollbar.configure(command=self._merge_listbox.yview)

    def _on_2up_toggle(self) -> None:
        if self._merge_2up_var.get():
            self._merge_4up_var.set(False)

    def _on_4up_toggle(self) -> None:
        if self._merge_4up_var.get():
            self._merge_2up_var.set(False)

    def _on_drop(self, event) -> None:
        """处理拖拽文件事件。"""
        # tkinterdnd2 传回格式: {path1} {path2} ...
        raw = event.data
        paths = re.findall(r'\{([^}]+)\}', raw)
        # 也处理不带花括号的路径
        if not paths:
            paths = raw.split()
        for p in paths:
            p = p.strip()
            if p.lower().endswith('.pdf') and p not in self._file_list:
                try:
                    info = get_pdf_info(p)
                    self._file_list.append(p)
                    self._merge_listbox.insert(
                        "end", f"{info['name']}  ({info['pages']} 页)"
                    )
                except PDFError as e:
                    messagebox.showerror("错误", str(e))

    def _add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        for p in paths:
            if p not in self._file_list:
                try:
                    info = get_pdf_info(p)
                    self._file_list.append(p)
                    self._merge_listbox.insert(
                        "end", f"{info['name']}  ({info['pages']} 页)"
                    )
                except PDFError as e:
                    messagebox.showerror("错误", str(e))

    def _remove_selected(self) -> None:
        sel = self._merge_listbox.curselection()
        for i in reversed(sel):
            self._merge_listbox.delete(i)
            del self._file_list[i]

    def _move_up(self) -> None:
        sel = self._merge_listbox.curselection()
        if not sel:
            return
        i = sel[0]
        if i == 0:
            return
        self._file_list[i], self._file_list[i - 1] = (
            self._file_list[i - 1],
            self._file_list[i],
        )
        self._refresh_merge_list()
        self._merge_listbox.selection_set(i - 1)

    def _move_down(self) -> None:
        sel = self._merge_listbox.curselection()
        if not sel:
            return
        i = sel[0]
        if i >= len(self._file_list) - 1:
            return
        self._file_list[i], self._file_list[i + 1] = (
            self._file_list[i + 1],
            self._file_list[i],
        )
        self._refresh_merge_list()
        self._merge_listbox.selection_set(i + 1)

    def _refresh_merge_list(self) -> None:
        self._merge_listbox.delete(0, "end")
        for p in self._file_list:
            info = get_pdf_info(p)
            self._merge_listbox.insert("end", f"{info['name']}  ({info['pages']} 页)")

    def _on_merge_select(self, event) -> None:
        # 预览已改为独立弹窗，选中仅保留高亮
        pass

    # ═══════════════════════════════════════════════════════
    # 通用「单文件操作页」基础
    # ═══════════════════════════════════════════════════════

    def _build_op_page(self, tab_title: str) -> tuple[ttk.Frame, ttk.Frame]:
        tab = ttk.Frame(self._notebook)
        self._notebook.add(tab, text=tab_title)
        page = self._page(tab)
        return tab, page

    # ═══════════════════════════════════════════════════════
    # 拆分标签页
    # ═══════════════════════════════════════════════════════

    def _build_split_tab(self) -> None:
        tab, page = self._build_op_page("拆分 PDF")

        body = self._card(page, "源文件")
        self._split_path_var = tk.StringVar()
        self._path_row(body, self._split_path_var, self._choose_split_file)

        body = self._card(page, "拆分设置")
        row = ttk.Frame(body)
        row.pack(fill="x")
        ttk.Label(row, text="每", style="App.TLabel").pack(side="left")
        self._split_pages_var = tk.StringVar(value="1")
        ttk.Spinbox(
            row, from_=1, to=999, width=6,
            textvariable=self._split_pages_var, style="App.TSpinbox",
        ).pack(side="left", padx=(6, 6))
        ttk.Label(row, text="页拆分为一个文件", style="App.TLabel").pack(side="left")

        body = self._card(page, "输出目录")
        self._split_out_var = tk.StringVar()
        self._path_row(body, self._split_out_var, self._choose_split_out)

        footer = ttk.Frame(page)
        footer.pack(fill="x")
        ttk.Button(footer, text="预览", style="App.Subtle.TButton",
                   command=self._open_preview).pack(side="left")
        self._page_hint(footer, "调整好选项后，点击右下角「开始拆分」执行")

    def _choose_split_file(self) -> None:
        p = filedialog.askopenfilename(
            title="选择要拆分的 PDF",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if p:
            self._split_path_var.set(p)

    def _choose_split_out(self) -> None:
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self._split_out_var.set(d)

    # ═══════════════════════════════════════════════════════
    # 提取标签页
    # ═══════════════════════════════════════════════════════

    def _build_extract_tab(self) -> None:
        tab, page = self._build_op_page("提取页面")

        body = self._card(page, "源文件")
        self._extract_path_var = tk.StringVar()
        self._path_row(body, self._extract_path_var, self._choose_extract_file)

        body = self._card(page, "页码范围")
        row = ttk.Frame(body)
        row.pack(fill="x")
        self._extract_range_var = tk.StringVar()
        self._extract_range_entry = ttk.Entry(
            row, textvariable=self._extract_range_var, style="App.TEntry",
        )
        self._extract_range_entry.pack(
            side="left", fill="x", expand=True, padx=(0, 8)
        )
        ttk.Label(row, text="例: 1-3,5,7-9", style="App.Hint.TLabel").pack(
            side="left"
        )

        body = self._card(page, "保存到")
        self._extract_out_var = tk.StringVar()
        self._path_row(body, self._extract_out_var, self._choose_extract_out)

        footer = ttk.Frame(page)
        footer.pack(fill="x")
        ttk.Button(footer, text="预览", style="App.Subtle.TButton",
                   command=self._open_preview).pack(side="left")
        self._page_hint(footer, "调整好选项后，点击右下角「开始提取」执行")

    def _choose_extract_file(self) -> None:
        p = filedialog.askopenfilename(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if p:
            self._extract_path_var.set(p)

    def _choose_extract_out(self) -> None:
        p = filedialog.asksaveasfilename(
            title="保存提取的 PDF",
            defaultextension=".pdf",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if p:
            self._extract_out_var.set(p)

    # ═══════════════════════════════════════════════════════
    # 旋转标签页
    # ═══════════════════════════════════════════════════════

    def _build_rotate_tab(self) -> None:
        tab, page = self._build_op_page("旋转页面")

        body = self._card(page, "源文件")
        self._rotate_path_var = tk.StringVar()
        self._path_row(body, self._rotate_path_var, self._choose_rotate_file)

        body = self._card(page, "旋转设置")
        grid = ttk.Frame(body)
        grid.pack(fill="x")
        grid.columnconfigure(1, weight=1)

        ttk.Label(grid, text="页码（逗号分隔）", style="App.TLabel").grid(
            row=0, column=0, sticky="w", pady=(2, 8), padx=(0, 10)
        )
        self._rotate_pages_var = tk.StringVar()
        ttk.Entry(grid, textvariable=self._rotate_pages_var,
                  style="App.TEntry").grid(
            row=0, column=1, sticky="ew", padx=(0, 8), pady=(2, 8)
        )
        ttk.Label(grid, text="留空 = 全部页面", style="App.Hint.TLabel").grid(
            row=0, column=2, sticky="w", pady=(2, 8)
        )

        ttk.Label(grid, text="旋转角度", style="App.TLabel").grid(
            row=1, column=0, sticky="w", pady=(2, 0), padx=(0, 10)
        )
        self._rotate_angle_var = tk.StringVar(value="90")
        ttk.Combobox(
            grid, textvariable=self._rotate_angle_var,
            values=["90", "180", "270"], state="readonly",
            width=8, style="App.TCombobox",
        ).grid(row=1, column=1, sticky="w", pady=(2, 0))
        ttk.Label(grid, text="顺时针", style="App.Hint.TLabel").grid(
            row=1, column=2, sticky="w", pady=(2, 0)
        )

        body = self._card(page, "保存到")
        self._rotate_out_var = tk.StringVar()
        self._path_row(body, self._rotate_out_var, self._choose_rotate_out)

        footer = ttk.Frame(page)
        footer.pack(fill="x")
        ttk.Button(footer, text="预览", style="App.Subtle.TButton",
                   command=self._open_preview).pack(side="left")
        self._page_hint(footer, "调整好选项后，点击右下角「开始旋转」执行")

    def _choose_rotate_file(self) -> None:
        p = filedialog.askopenfilename(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if p:
            self._rotate_path_var.set(p)

    def _choose_rotate_out(self) -> None:
        p = filedialog.asksaveasfilename(
            title="保存旋转后的 PDF",
            defaultextension=".pdf",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if p:
            self._rotate_out_var.set(p)

    # ═══════════════════════════════════════════════════════
    # 预览
    # ═══════════════════════════════════════════════════════

    def _current_tab_source(self) -> str:
        """非合并标签页当前选中的源 PDF 路径。"""
        idx = self._notebook.index("current")
        if idx == 1:
            return self._split_path_var.get()
        if idx == 2:
            return self._extract_path_var.get()
        if idx == 3:
            return self._rotate_path_var.get()
        return ""

    def _open_preview(self) -> None:
        """打开/刷新独立预览弹窗。"""
        if self._preview_win is None or self._preview_win.closed:
            self._preview_win = PreviewWindow(self.root)

        idx = self._notebook.index("current")
        tab_name = self._notebook.tab(idx, "text")

        if idx == 0:  # 合并：源页 + 排版结果
            self._preview_win.show_merge(
                self._file_list, self._merge_2up_var.get(),
                self._merge_4up_var.get(),
            )
        else:
            src = self._current_tab_source()
            if src:
                self._preview_win.show_sources([src], header=f"预览 — {tab_name}")
            else:
                self._preview_win.show_message(f"预览 — {tab_name}", "请先选择文件")

        self._preview_win.bring_front()

    # ═══════════════════════════════════════════════════════
    # 公共方法
    # ═══════════════════════════════════════════════════════

    def _on_tab_change(self, event) -> None:
        """切换标签页时更新底部执行按钮文本。"""
        self._status_var.set("就绪")
        idx = self._notebook.index("current")
        self._btn_execute.configure(text=TAB_LABELS[idx])

    def _set_ui_state(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self._btn_execute.configure(state=state)
        # 取消按钮仅在"任务运行中且该任务支持取消"时可用
        cancel_state = "normal" if (not enabled and self._cancel_armed) else "disabled"
        self._btn_cancel.configure(state=cancel_state)

    def _run_operation(self) -> None:
        tab_idx = self._notebook.index("current")
        if tab_idx == 0:
            self._do_merge()
        elif tab_idx == 1:
            self._do_split()
        elif tab_idx == 2:
            self._do_extract()
        elif tab_idx == 3:
            self._do_rotate()
        elif tab_idx == 4:
            self._do_excel_merge()

    def _run_in_thread(
        self, func, cancel_event: threading.Event | None = None
    ) -> None:
        self._cancel_event = cancel_event
        self._cancel_armed = cancel_event is not None
        self._set_ui_state(False)
        self._progress.start()
        self._status_var.set("处理中...")

        def wrapper():
            try:
                func()
                self.root.after(0, lambda: self._on_done(None))
            except ExcelCancelled:
                self.root.after(0, self._on_cancelled)
            except (PDFError, ExcelError) as e:
                self.root.after(0, lambda: self._on_done(str(e)))
            except Exception as e:
                self.root.after(0, lambda: self._on_done(f"未知错误: {str(e)}"))

        threading.Thread(target=wrapper, daemon=True).start()

    def _on_cancel_click(self) -> None:
        """用户点击取消：置位事件，合并循环在下一个检查点中止。"""
        if self._cancel_event is not None:
            self._cancel_event.set()
            self._btn_cancel.configure(state="disabled")
            self._status_var.set("正在取消，请稍候...")

    def _run_with_save(self, title, ext, ftypes, func) -> None:
        """先在主线程弹保存对话框，再开子线程执行 func(output_path)。"""
        out = filedialog.asksaveasfilename(
            title=title, defaultextension=ext, filetypes=ftypes,
        )
        if not out:
            return

        def _do():
            func(out)
            self.root.after(0, lambda: messagebox.showinfo("完成", f"已保存到:\n{out}"))

        self._run_in_thread(_do)

    def _on_done(self, error: str | None) -> None:
        self._cancel_event = None
        self._cancel_armed = False
        self._progress.stop()
        self._set_ui_state(True)
        if error:
            self._status_var.set(f"失败: {error}")
            messagebox.showerror("操作失败", error)
        else:
            self._status_var.set("完成")

    def _on_cancelled(self) -> None:
        """用户主动取消：静默恢复界面，不弹错误框。"""
        self._cancel_event = None
        self._cancel_armed = False
        self._progress.stop()
        self._set_ui_state(True)
        self._status_var.set("已取消")

    # ── 具体操作 ──

    def _do_merge(self) -> None:
        if len(self._file_list) < 1:
            messagebox.showwarning("提示", "请先添加至少一个 PDF 文件")
            return

        def _run(out):
            if self._merge_4up_var.get():
                merge_4up(self._file_list, out)
            elif self._merge_2up_var.get():
                merge_2up(self._file_list, out)
            else:
                merge_pdfs(self._file_list, out)

        self._run_with_save("保存合并后的 PDF", ".pdf",
                            [("PDF 文件", "*.pdf")], _run)

    def _do_split(self) -> None:
        src = self._split_path_var.get()
        if not src:
            messagebox.showwarning("提示", "请先选择要拆分的 PDF 文件")
            return
        try:
            per_file = int(self._split_pages_var.get())
        except ValueError:
            messagebox.showwarning("提示", "请输入有效的页数")
            return
        out_dir = self._split_out_var.get()
        if not out_dir:
            messagebox.showwarning("提示", "请先选择输出目录")
            return

        def _run():
            results = split_pdf(src, out_dir, per_file)
            msg = (f"已生成 {len(results)} 个文件:\n" + "\n".join(results[:5])
                   + ("\n... 等" if len(results) > 5 else ""))
            self.root.after(0, lambda: messagebox.showinfo("完成", msg))

        self._run_in_thread(_run)

    def _do_extract(self) -> None:
        src = self._extract_path_var.get()
        if not src:
            messagebox.showwarning("提示", "请先选择 PDF 文件")
            return
        raw = self._extract_range_var.get().strip()
        if not raw:
            messagebox.showwarning("提示", "请输入页码范围")
            return
        ranges = [r.strip() for r in raw.split(",") if r.strip()]
        out = self._extract_out_var.get()
        if not out:
            messagebox.showwarning("提示", "请先选择保存位置")
            return

        def _run():
            extract_pages(src, out, ranges)
            self.root.after(0, lambda: messagebox.showinfo("完成", f"已保存到:\n{out}"))

        self._run_in_thread(_run)

    def _do_rotate(self) -> None:
        src = self._rotate_path_var.get()
        if not src:
            messagebox.showwarning("提示", "请先选择 PDF 文件")
            return
        pages_raw = self._rotate_pages_var.get().strip()
        try:
            angle = int(self._rotate_angle_var.get())
        except ValueError:
            messagebox.showwarning("提示", "请选择旋转角度")
            return
        out = self._rotate_out_var.get()
        if not out:
            messagebox.showwarning("提示", "请先选择保存位置")
            return
        if pages_raw:
            page_nums = [int(p.strip()) for p in pages_raw.split(",") if p.strip()]
        else:
            page_nums = list(range(1, get_pdf_info(src)["pages"] + 1))

        def _run():
            rotate_pages(src, out, page_nums, angle)
            self.root.after(0, lambda: messagebox.showinfo("完成", f"已保存到:\n{out}"))

        self._run_in_thread(_run)

    # ═══════════════════════════════════════════════════════
    # 合并 Excel 标签页
    # ═══════════════════════════════════════════════════════

    def _build_excel_tab(self) -> None:
        tab = ttk.Frame(self._notebook)
        self._notebook.add(tab, text="合并 Excel")
        page = self._page(tab)

        body = self._card(page, "Excel 文件列表（可直接拖入 .xlsx / .xls）",
                          expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(body)
        toolbar.grid(row=0, column=0, sticky="ew", pady=(4, 10))
        for txt, cmd in (
            ("＋ 添加文件", self._add_excel_files),
            ("✕ 移除所选", self._remove_excel_selected),
            ("清空列表", self._clear_excel_list),
        ):
            ttk.Button(toolbar, text=txt, style="App.TButton",
                       command=cmd).pack(side="left", padx=(0, 6))

        listwrap = ttk.Frame(body)
        listwrap.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(listwrap, orient="vertical")
        scrollbar.pack(side="right", fill="y")

        self._excel_listbox = tk.Listbox(
            listwrap,
            yscrollcommand=scrollbar.set,
            font=FONT_UI,
            selectmode="extended",
            bg="#FFFFFF", fg=C_TEXT,
            selectbackground=C_SELECT_BG, selectforeground=C_SELECT_FG,
            relief="flat", borderwidth=0, activestyle="none",
            highlightthickness=1,
            highlightbackground=C_BORDER, highlightcolor=C_SELECT_BG,
            exportselection=False,
        )
        self._excel_listbox.pack(fill="both", expand=True, side="left")
        self._excel_listbox.drop_target_register(DND_FILES)
        self._excel_listbox.dnd_bind("<<Drop>>", self._on_excel_drop)
        scrollbar.configure(command=self._excel_listbox.yview)

        # 合并方式
        self._excel_merge_mode = tk.StringVar(value="sheets")
        body = self._card(page, "合并方式")

        ttk.Radiobutton(
            body,
            text="每文件一个 Sheet（文件名作 Sheet 名）— 适合把各报告并排存放",
            variable=self._excel_merge_mode, value="sheets",
        ).pack(anchor="w", pady=3)
        ttk.Radiobutton(
            body,
            text="追加到一个 Sheet（按表头列名对齐）— 适合结构相同的明细表",
            variable=self._excel_merge_mode, value="append",
        ).pack(anchor="w", pady=3)
        ttk.Radiobutton(
            body,
            text="整表复制到一个新文件（保留原格式）— 打印用，需本机安装 Excel",
            variable=self._excel_merge_mode, value="copy",
        ).pack(anchor="w", pady=3)
        self._excel_merge_hint = ttk.Label(
            body, style="App.Hint.TLabel",
            text="提示：要打印多份报告时选「整表复制」，格式 / 列宽 / 分页设置会原样保留，"
                 "合并后一次打印整个工作簿即可。",
        )
        self._excel_merge_hint.pack(anchor="w", pady=(8, 2))

    def _on_excel_drop(self, event) -> None:
        raw = event.data
        paths = re.findall(r'\{([^}]+)\}', raw)
        if not paths:
            paths = raw.split()
        for p in paths:
            p = p.strip()
            if p.lower().endswith(('.xlsx', '.xls')) and p not in self._excel_file_list:
                self._excel_file_list.append(p)
                self._excel_listbox.insert("end", Path(p).name)

    def _add_excel_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls")],
        )
        for p in paths:
            if p not in self._excel_file_list:
                self._excel_file_list.append(p)
                self._excel_listbox.insert("end", Path(p).name)

    def _remove_excel_selected(self) -> None:
        sel = self._excel_listbox.curselection()
        for i in reversed(sel):
            self._excel_listbox.delete(i)
            del self._excel_file_list[i]

    def _clear_excel_list(self) -> None:
        self._excel_listbox.delete(0, "end")
        self._excel_file_list.clear()

    def _do_excel_merge(self) -> None:
        if not self._excel_file_list:
            messagebox.showwarning("提示", "请先添加 Excel 文件")
            return

        mode = self._excel_merge_mode.get()

        out = filedialog.asksaveasfilename(
            title="保存合并后的 Excel", defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not out:
            return

        cancel_event = threading.Event()

        def progress(fi: int, total: int, fname: str, rows: int) -> None:
            # 由工作线程回调；通过 after 切回主线程更新状态栏
            if mode == "copy":
                txt = (f"整表复制 {fi}/{total}: {fname} — "
                       f"已复制 {rows} 个Sheet")
            else:
                txt = f"合并中 {fi}/{total}: {fname} — 已处理 {rows:,} 行"
            self.root.after(0, lambda: self._status_var.set(txt))

        def _run():
            if mode == "copy":
                merge_copy_sheets(self._excel_file_list, out,
                                  progress=progress, cancel_event=cancel_event)
            elif mode == "sheets":
                merge_to_sheets(self._excel_file_list, out,
                                progress=progress, cancel_event=cancel_event)
            else:
                merge_to_one_sheet(self._excel_file_list, out,
                                   progress=progress, cancel_event=cancel_event)
            self.root.after(0, lambda: messagebox.showinfo("完成", f"已保存到:\n{out}"))

        self._run_in_thread(_run, cancel_event=cancel_event)
