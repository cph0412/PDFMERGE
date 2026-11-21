# PDF 合并工具 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建一个 Windows 桌面 PDF 工具箱，支持合并、拆分、提取页面和旋转功能，带页面预览，打包为单个 .exe。

**Architecture:** 三层结构 — pdf_ops.py（纯逻辑层，无 GUI 依赖）→ preview.py（Canvas 渲染组件）→ gui.py（tkinter 界面，编排逻辑和 UI）。main.py 仅负责创建窗口和启动。

**Tech Stack:** Python 3.11+, tkinter, pymupdf (fitz), Pillow, PyInstaller

## Global Constraints

- 平台: Windows 10/11
- Python >= 3.11
- GUI: tkinter（Python 自带，零依赖）
- PDF 引擎: pymupdf >= 1.23.0
- 预览: Pillow >= 10.0.0
- 打包: PyInstaller --onefile --windowed

## File Structure

```
pdf/
├── main.py          # 入口，创建 root 窗口
├── gui.py           # 主界面（标签页、列表、按钮、编排）
├── pdf_ops.py       # 纯 PDF 操作函数
├── preview.py       # 页面预览组件
├── requirements.txt
└── build.bat        # 打包脚本
```

---

### Task 1: 项目脚手架

**Files:**
- Create: `requirements.txt`
- Create: `build.bat`

**Interfaces:**
- Produces: 项目依赖声明和打包脚本，后续所有任务依赖此环境

- [ ] **Step 1: 创建 requirements.txt**

```txt
pymupdf>=1.23.0
Pillow>=10.0.0
pyinstaller>=6.0.0
```

- [ ] **Step 2: 创建 build.bat**

```bat
@echo off
chcp 65001 >nul
echo ========================================
echo   PDF 工具箱 — 打包脚本
echo ========================================

if not exist "dist" mkdir dist

pyinstaller --onefile --windowed --name "PDF工具箱" --clean main.py

echo.
echo 打包完成: dist\PDF工具箱.exe
pause
```

- [ ] **Step 3: 安装依赖**

```bash
pip install -r requirements.txt
```

Expected: 三个包安装成功

- [ ] **Step 4: 验证环境**

```bash
python -c "import tkinter; import fitz; from PIL import Image; print('环境就绪')"
```

Expected: 输出 "环境就绪"

- [ ] **Step 5: Commit**

```bash
git add requirements.txt build.bat
git commit -m "chore: add project scaffold with dependencies and build script"
```

---

### Task 2: PDF 操作核心模块

**Files:**
- Create: `pdf_ops.py`
- Test: `test_pdf_ops.py`（手动验证用临时脚本，本任务是纯逻辑层，不测 GUI）

**Interfaces:**
- Produces:
  - `merge_pdfs(file_paths: list[str], output_path: str) -> None`
  - `split_pdf(file_path: str, output_dir: str, pages_per_file: int) -> list[str]`
  - `extract_pages(file_path: str, output_path: str, page_ranges: list[str]) -> None`
  - `rotate_pages(file_path: str, output_path: str, page_nums: list[int], angle: int) -> None`
  - `get_page_count(file_path: str) -> int`
  - `get_pdf_info(file_path: str) -> dict`

- [ ] **Step 1: 编写 pdf_ops.py 完整代码**

```python
"""PDF 核心操作模块 — 纯逻辑层，不依赖 GUI。"""

import fitz
from pathlib import Path


class PDFError(Exception):
    """PDF 操作异常。"""
    pass


def get_page_count(file_path: str) -> int:
    """返回 PDF 的页数。"""
    try:
        doc = fitz.open(file_path)
        if doc.is_encrypted:
            doc.close()
            raise PDFError(f"文件已加密，无法处理: {file_path}")
        count = doc.page_count
        doc.close()
        return count
    except PDFError:
        raise
    except Exception as e:
        raise PDFError(f"无法打开文件: {file_path}\n{str(e)}")


def get_pdf_info(file_path: str) -> dict:
    """返回 PDF 基本信息：文件名、页数。"""
    return {
        "path": file_path,
        "name": Path(file_path).name,
        "pages": get_page_count(file_path),
    }


def merge_pdfs(file_paths: list[str], output_path: str) -> None:
    """合并多个 PDF 文件为一个。

    Args:
        file_paths: PDF 文件路径列表，按顺序合并
        output_path: 输出文件路径
    """
    if not file_paths:
        raise PDFError("没有选择任何文件")

    result = fitz.open()

    try:
        for path in file_paths:
            doc = fitz.open(path)
            if doc.is_encrypted:
                doc.close()
                result.close()
                raise PDFError(f"文件已加密，无法合并: {path}")
            result.insert_pdf(doc)
            doc.close()

        result.save(output_path, garbage=4, deflate=True)
    except PDFError:
        raise
    except Exception as e:
        raise PDFError(f"合并失败: {str(e)}")
    finally:
        result.close()


def split_pdf(file_path: str, output_dir: str, pages_per_file: int) -> list[str]:
    """按每 N 页拆分 PDF。

    Args:
        file_path: 源 PDF 路径
        output_dir: 输出目录
        pages_per_file: 每个拆分文件的页数

    Returns:
        生成的 PDF 文件路径列表
    """
    if pages_per_file < 1:
        raise PDFError("每文件页数必须大于 0")

    doc = fitz.open(file_path)
    if doc.is_encrypted:
        doc.close()
        raise PDFError(f"文件已加密，无法拆分: {file_path}")

    total = doc.page_count
    base_name = Path(file_path).stem
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    outputs = []

    try:
        for start in range(0, total, pages_per_file):
            end = min(start + pages_per_file, total)
            new_doc = fitz.open()
            new_doc.insert_pdf(doc, from_page=start, to_page=end - 1)

            out_name = f"{base_name}_p{start+1}-{end}.pdf"
            out_path = out_dir / out_name
            new_doc.save(str(out_path), garbage=4, deflate=True)
            new_doc.close()
            outputs.append(str(out_path))

        doc.close()
        return outputs
    except Exception as e:
        raise PDFError(f"拆分失败: {str(e)}")


def extract_pages(file_path: str, output_path: str, page_ranges: list[str]) -> None:
    """提取指定页码范围为一个新 PDF。

    Args:
        file_path: 源 PDF 路径
        output_path: 输出文件路径
        page_ranges: 页码范围列表，如 ["1-3", "5", "7-9"]
                     页码从 1 开始计数
    """
    doc = fitz.open(file_path)
    if doc.is_encrypted:
        doc.close()
        raise PDFError(f"文件已加密，无法提取: {file_path}")

    total = doc.page_count
    new_doc = fitz.open()

    try:
        for r in page_ranges:
            r = r.strip()
            if '-' in r:
                parts = r.split('-', 1)
                start = int(parts[0].strip()) - 1
                end = int(parts[1].strip())
            else:
                start = int(r) - 1
                end = start + 1

            if start < 0 or end > total or start >= end:
                new_doc.close()
                doc.close()
                raise PDFError(f"页码范围无效: {r}（总页数: {total}）")

            new_doc.insert_pdf(doc, from_page=start, to_page=end - 1)

        new_doc.save(output_path, garbage=4, deflate=True)
    except PDFError:
        raise
    except ValueError:
        raise PDFError(f"页码范围格式错误: {page_ranges}")
    except Exception as e:
        raise PDFError(f"提取页面失败: {str(e)}")
    finally:
        new_doc.close()
        doc.close()


def rotate_pages(file_path: str, output_path: str, page_nums: list[int], angle: int) -> None:
    """旋转指定页面。

    Args:
        file_path: 源 PDF 路径
        output_path: 输出文件路径
        page_nums: 要旋转的页码列表，从 1 开始
        angle: 旋转角度，支持 90, 180, 270
    """
    if angle not in (90, 180, 270):
        raise PDFError(f"不支持的旋转角度: {angle}（仅支持 90, 180, 270）")

    doc = fitz.open(file_path)
    if doc.is_encrypted:
        doc.close()
        raise PDFError(f"文件已加密，无法旋转: {file_path}")

    total = doc.page_count

    try:
        for page_num in page_nums:
            if page_num < 1 or page_num > total:
                raise PDFError(f"页码无效: {page_num}（总页数: {total}）")

            page = doc[page_num - 1]
            page.set_rotation((page.rotation + angle) % 360)

        doc.save(output_path, garbage=4, deflate=True)
    finally:
        doc.close()
```

- [ ] **Step 2: 编写手动验证脚本 test_pdf_ops.py**

```python
"""pdf_ops 手动验证 — 使用前需准备测试 PDF"""
from pdf_ops import merge_pdfs, split_pdf, extract_pages, rotate_pages, get_page_count, get_pdf_info

# 请替换为实际的测试 PDF 路径
TEST_DIR = r"D:\CPH\APP\pdf\test_data"

if __name__ == "__main__":
    import os
    if not os.path.isdir(TEST_DIR):
        print(f"请创建测试目录并放入 PDF 文件: {TEST_DIR}")
        exit(1)

    # 测试 get_page_count
    files = [f for f in os.listdir(TEST_DIR) if f.endswith('.pdf')]
    if files:
        f = os.path.join(TEST_DIR, files[0])
        print("get_page_count:", get_page_count(f))
        print("get_pdf_info:", get_pdf_info(f))

    print("所有基础测试通过 ✓")
```

- [ ] **Step 3: 用实际 PDF 文件验证核心功能**

```bash
python test_pdf_ops.py
```

Expected: 无报错，输出页数和文件信息

- [ ] **Step 4: Commit**

```bash
git add pdf_ops.py
git commit -m "feat: add PDF core operations module"
```

---

### Task 3: 页面预览组件

**Files:**
- Create: `preview.py`

**Interfaces:**
- Consumes: `pymupdf (fitz)`, `PIL.Image`, `PIL.ImageTk`, `tkinter.Canvas`
- Produces: `class PDFPreview` — 方法:
  - `__init__(self, canvas: tk.Canvas)`
  - `load(self, pdf_path: str) -> None`
  - `show_page(self, page_num: int) -> None`
  - `next_page(self) -> None`
  - `prev_page(self) -> None`
  - `unload(self) -> None`
  - `page_count(self) -> int` (property)
  - `current_page(self) -> int` (property)

- [ ] **Step 1: 编写 preview.py 完整代码**

```python
"""PDF 页面预览组件 — 基于 tkinter Canvas + pymupdf 渲染。"""

import fitz
from PIL import Image, ImageTk
import tkinter as tk


class PDFPreview:
    """在 tkinter Canvas 上渲染 PDF 页面预览。

    用法:
        canvas = tk.Canvas(root, width=400, height=500)
        preview = PDFPreview(canvas)
        preview.load("example.pdf")
        preview.next_page()
        preview.unload()
    """

    def __init__(self, canvas: tk.Canvas):
        self._canvas = canvas
        self._doc = None
        self._page = 0
        self._pages = 0
        self._image_ref = None  # 持有引用防止 GC

    @property
    def page_count(self) -> int:
        return self._pages

    @property
    def current_page(self) -> int:
        """当前页码，从 1 开始。"""
        return self._page + 1

    def load(self, pdf_path: str) -> None:
        """加载 PDF 并显示第一页。"""
        self.unload()
        self._doc = fitz.open(pdf_path)
        self._pages = self._doc.page_count
        self._page = 0
        if self._pages > 0:
            self._render()

    def show_page(self, page_num: int) -> None:
        """显示指定页码（从 1 开始）。"""
        if not self._doc:
            return
        if 1 <= page_num <= self._pages:
            self._page = page_num - 1
            self._render()

    def next_page(self) -> None:
        if self._doc and self._page < self._pages - 1:
            self._page += 1
            self._render()

    def prev_page(self) -> None:
        if self._doc and self._page > 0:
            self._page -= 1
            self._render()

    def unload(self) -> None:
        if self._doc:
            self._doc.close()
            self._doc = None
        self._pages = 0
        self._page = 0
        self._image_ref = None
        self._canvas.delete("all")
        self._draw_placeholder("未加载文件")

    def _render(self) -> None:
        """渲染当前页到 Canvas。"""
        zoom = 2.0
        mat = fitz.Matrix(zoom, zoom)
        pix = self._doc[self._page].get_pixmap(matrix=mat)

        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        cw = self._canvas.winfo_width() or 350
        ch = self._canvas.winfo_height() or 480

        if cw > 20 and ch > 20:
            img.thumbnail((cw - 24, ch - 24), Image.LANCZOS)

        self._image_ref = ImageTk.PhotoImage(img)
        self._canvas.delete("all")
        self._canvas.create_image(
            cw // 2, ch // 2,
            anchor="center",
            image=self._image_ref,
        )

    def _draw_placeholder(self, text: str) -> None:
        cw = self._canvas.winfo_width() or 350
        ch = self._canvas.winfo_height() or 480
        self._canvas.create_text(
            cw // 2, ch // 2,
            text=text,
            fill="#888888",
            font=("Microsoft YaHei UI", 11),
        )
```

- [ ] **Step 2: Commit**

```bash
git add preview.py
git commit -m "feat: add PDF page preview component"
```

---

### Task 4: 主界面 GUI

**Files:**
- Create: `gui.py`

**Interfaces:**
- Consumes: `pdf_ops` (全部函数), `preview.PDFPreview`, `tkinter`, `tkinter.ttk`, `tkinter.filedialog`, `tkinter.messagebox`, `threading`
- Produces: `class PDFToolApp` — 方法:
  - `__init__(self, root: tk.Tk) -> None`

- [ ] **Step 1: 编写 gui.py — 导入和常量**

```python
"""PDF 工具箱 — 主界面。"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os

from pdf_ops import (
    merge_pdfs,
    split_pdf,
    extract_pages,
    rotate_pages,
    get_pdf_info,
    PDFError,
)
from preview import PDFPreview

# ── 界面常量 ──────────────────────────────────────────────
FONT_UI = ("Microsoft YaHei UI", 10)
FONT_BTN = ("Microsoft YaHei UI", 10)
FONT_TITLE = ("Microsoft YaHei UI", 11, "bold")
WIN_WIDTH = 960
WIN_HEIGHT = 640
PREVIEW_W = 380
```

- [ ] **Step 2: 编写 PDFToolApp.__init__ 骨架和主布局**

```python
class PDFToolApp:
    """PDF 工具箱主应用。"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF 工具箱")
        self.root.geometry(f"{WIN_WIDTH}x{WIN_HEIGHT}")
        self.root.minsize(800, 550)

        # 样式
        style = ttk.Style()
        style.theme_use("clam")

        # ── 主布局 ──
        # 顶部标签栏
        self._notebook = ttk.Notebook(root)
        self._notebook.pack(fill="x", padx=8, pady=(8, 0))

        # 中间区域：左（操作面板）+ 右（预览）
        self._main_frame = ttk.Frame(root)
        self._main_frame.pack(fill="both", expand=True, padx=8, pady=4)

        # 左侧操作面板
        self._left_panel = ttk.Frame(self._main_frame, width=WIN_WIDTH - PREVIEW_W - 24)
        self._left_panel.pack(side="left", fill="both", expand=True)
        self._left_panel.pack_propagate(False)

        # 右侧预览面板
        self._right_panel = ttk.LabelFrame(self._main_frame, text="预览")
        self._right_panel.pack(side="right", fill="both", expand=False, padx=(4, 0))
        self._right_panel.configure(width=PREVIEW_W)

        self._preview_canvas = tk.Canvas(
            self._right_panel,
            width=PREVIEW_W - 16,
            bg="#f0f0f0",
            highlightthickness=0,
        )
        self._preview_canvas.pack(fill="both", expand=True, padx=4, pady=4)
        self._preview = PDFPreview(self._preview_canvas)

        # 预览翻页按钮
        self._preview_ctrl = ttk.Frame(self._right_panel)
        self._preview_ctrl.pack(fill="x", padx=4, pady=(0, 4))
        self._btn_prev = ttk.Button(
            self._preview_ctrl, text="◀ 上一页", command=self._preview_prev
        )
        self._btn_prev.pack(side="left")
        self._lbl_page = ttk.Label(self._preview_ctrl, text="0 / 0")
        self._lbl_page.pack(side="left", expand=True)
        self._btn_next = ttk.Button(
            self._preview_ctrl, text="下一页 ▶", command=self._preview_next
        )
        self._btn_next.pack(side="right")

        # 底部按钮区
        self._bottom_frame = ttk.Frame(root)
        self._bottom_frame.pack(fill="x", padx=8, pady=(0, 8))

        self._progress = ttk.Progressbar(self._bottom_frame, mode="indeterminate")
        self._progress.pack(fill="x", pady=(0, 4))

        self._btn_execute = ttk.Button(
            self._bottom_frame, text="执行", command=self._run_operation
        )
        self._btn_execute.pack(side="right")

        self._status_var = tk.StringVar(value="就绪")
        ttk.Label(self._bottom_frame, textvariable=self._status_var,
                  foreground="#666").pack(side="left")

        # ── 构建标签页 ──
        self._build_merge_tab()
        self._build_split_tab()
        self._build_extract_tab()
        self._build_rotate_tab()

        self._notebook.bind("<<NotebookTabChanged>>", self._on_tab_change)

        # 预制状态
        self._current_pdf: str | None = None
        self._file_list: list[str] = []
```

- [ ] **Step 3: 编写合并标签页**

```python
    def _build_merge_tab(self) -> None:
        tab = ttk.Frame(self._notebook)
        self._notebook.add(tab, text="合并 PDF")

        # 文件列表
        list_frame = ttk.LabelFrame(tab, text="文件列表（按合并顺序排列）")
        list_frame.pack(fill="both", expand=True, padx=4, pady=4)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self._merge_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=FONT_UI,
            selectmode="extended",
        )
        self._merge_listbox.pack(fill="both", expand=True, padx=4, pady=4)
        self._merge_listbox.bind("<<ListboxSelect>>", self._on_merge_select)
        scrollbar.configure(command=self._merge_listbox.yview)

        # 按钮
        btn_frame = ttk.Frame(tab)
        btn_frame.pack(fill="x", padx=4, pady=(0, 4))

        ttk.Button(btn_frame, text="添加文件", command=self._add_files).pack(
            side="left", padx=2
        )
        ttk.Button(btn_frame, text="移除选中", command=self._remove_selected).pack(
            side="left", padx=2
        )
        ttk.Button(btn_frame, text="↑ 上移", command=self._move_up).pack(
            side="left", padx=2
        )
        ttk.Button(btn_frame, text="↓ 下移", command=self._move_down).pack(
            side="left", padx=2
        )

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
        sel = self._merge_listbox.curselection()
        if sel:
            self._preview.load(self._file_list[sel[0]])
            self._update_preview_label()
        else:
            self._preview.unload()
            self._update_preview_label()
```

- [ ] **Step 4: 编写拆分标签页**

```python
    def _build_split_tab(self) -> None:
        tab = ttk.Frame(self._notebook)
        self._notebook.add(tab, text="拆分 PDF")

        pad = {"padx": 8, "pady": 6}

        # 文件选择
        file_frame = ttk.LabelFrame(tab, text="选择文件")
        file_frame.pack(fill="x", **pad)

        self._split_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self._split_path_var,
                  state="readonly", font=FONT_UI).pack(
            side="left", fill="x", expand=True, padx=4, pady=4
        )
        ttk.Button(file_frame, text="浏览", command=self._choose_split_file).pack(
            side="right", padx=4, pady=4
        )

        # 参数设置
        param_frame = ttk.LabelFrame(tab, text="拆分设置")
        param_frame.pack(fill="x", **pad)

        ttk.Label(param_frame, text="每").pack(side="left")
        self._split_pages_var = tk.StringVar(value="1")
        ttk.Spinbox(
            param_frame, from_=1, to=999, width=5,
            textvariable=self._split_pages_var,
        ).pack(side="left", padx=4, pady=8)
        ttk.Label(param_frame, text="页拆分为一个文件").pack(side="left")

        # 输出目录
        out_frame = ttk.LabelFrame(tab, text="输出目录")
        out_frame.pack(fill="x", **pad)

        self._split_out_var = tk.StringVar()
        ttk.Entry(out_frame, textvariable=self._split_out_var,
                  state="readonly", font=FONT_UI).pack(
            side="left", fill="x", expand=True, padx=4, pady=4
        )
        ttk.Button(out_frame, text="浏览", command=self._choose_split_out).pack(
            side="right", padx=4, pady=4
        )

    def _choose_split_file(self) -> None:
        p = filedialog.askopenfilename(
            title="选择要拆分的 PDF",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if p:
            self._split_path_var.set(p)
            self._current_pdf = p
            self._preview.load(p)
            self._update_preview_label()

    def _choose_split_out(self) -> None:
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self._split_out_var.set(d)
```

- [ ] **Step 5: 编写提取标签页**

```python
    def _build_extract_tab(self) -> None:
        tab = ttk.Frame(self._notebook)
        self._notebook.add(tab, text="提取页面")

        pad = {"padx": 8, "pady": 6}

        # 文件选择
        file_frame = ttk.LabelFrame(tab, text="选择文件")
        file_frame.pack(fill="x", **pad)

        self._extract_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self._extract_path_var,
                  state="readonly", font=FONT_UI).pack(
            side="left", fill="x", expand=True, padx=4, pady=4
        )
        ttk.Button(file_frame, text="浏览", command=self._choose_extract_file).pack(
            side="right", padx=4, pady=4
        )

        # 页码范围
        range_frame = ttk.LabelFrame(tab, text="页码范围")
        range_frame.pack(fill="x", **pad)

        self._extract_range_var = tk.StringVar()
        self._extract_range_entry = ttk.Entry(
            range_frame, textvariable=self._extract_range_var, font=FONT_UI,
        )
        self._extract_range_entry.pack(
            side="left", fill="x", expand=True, padx=4, pady=4
        )
        ttk.Label(range_frame, text="例: 1-3,5,7-9").pack(side="left", padx=4)

        # 输出文件
        out_frame = ttk.LabelFrame(tab, text="保存到")
        out_frame.pack(fill="x", **pad)

        self._extract_out_var = tk.StringVar()
        ttk.Entry(out_frame, textvariable=self._extract_out_var,
                  state="readonly", font=FONT_UI).pack(
            side="left", fill="x", expand=True, padx=4, pady=4
        )
        ttk.Button(out_frame, text="浏览", command=self._choose_extract_out).pack(
            side="right", padx=4, pady=4
        )

    def _choose_extract_file(self) -> None:
        p = filedialog.askopenfilename(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if p:
            self._extract_path_var.set(p)
            self._current_pdf = p
            self._preview.load(p)
            self._update_preview_label()

    def _choose_extract_out(self) -> None:
        p = filedialog.asksaveasfilename(
            title="保存提取的 PDF",
            defaultextension=".pdf",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if p:
            self._extract_out_var.set(p)
```

- [ ] **Step 6: 编写旋转标签页**

```python
    def _build_rotate_tab(self) -> None:
        tab = ttk.Frame(self._notebook)
        self._notebook.add(tab, text="旋转页面")

        pad = {"padx": 8, "pady": 6}

        # 文件选择
        file_frame = ttk.LabelFrame(tab, text="选择文件")
        file_frame.pack(fill="x", **pad)

        self._rotate_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self._rotate_path_var,
                  state="readonly", font=FONT_UI).pack(
            side="left", fill="x", expand=True, padx=4, pady=4
        )
        ttk.Button(file_frame, text="浏览", command=self._choose_rotate_file).pack(
            side="right", padx=4, pady=4
        )

        # 参数
        param_frame = ttk.LabelFrame(tab, text="旋转设置")
        param_frame.pack(fill="x", **pad)

        ttk.Label(param_frame, text="页码（逗号分隔）：").grid(
            row=0, column=0, sticky="w", padx=4, pady=4
        )
        self._rotate_pages_var = tk.StringVar()
        ttk.Entry(param_frame, textvariable=self._rotate_pages_var,
                  font=FONT_UI, width=30).grid(
            row=0, column=1, sticky="ew", padx=4, pady=4
        )
        ttk.Label(param_frame, text="例: 1,3,5 或留空表示全部").grid(
            row=0, column=2, sticky="w", padx=4
        )

        ttk.Label(param_frame, text="旋转角度：").grid(
            row=1, column=0, sticky="w", padx=4, pady=4
        )
        self._rotate_angle_var = tk.StringVar(value="90")
        ttk.Combobox(
            param_frame, textvariable=self._rotate_angle_var,
            values=["90", "180", "270"], state="readonly", width=8,
        ).grid(row=1, column=1, sticky="w", padx=4, pady=4)

        param_frame.columnconfigure(1, weight=1)

        # 输出文件
        out_frame = ttk.LabelFrame(tab, text="保存到")
        out_frame.pack(fill="x", **pad)

        self._rotate_out_var = tk.StringVar()
        ttk.Entry(out_frame, textvariable=self._rotate_out_var,
                  state="readonly", font=FONT_UI).pack(
            side="left", fill="x", expand=True, padx=4, pady=4
        )
        ttk.Button(out_frame, text="浏览", command=self._choose_rotate_out).pack(
            side="right", padx=4, pady=4
        )

    def _choose_rotate_file(self) -> None:
        p = filedialog.askopenfilename(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if p:
            self._rotate_path_var.set(p)
            self._current_pdf = p
            self._preview.load(p)
            self._update_preview_label()

    def _choose_rotate_out(self) -> None:
        p = filedialog.asksaveasfilename(
            title="保存旋转后的 PDF",
            defaultextension=".pdf",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if p:
            self._rotate_out_var.set(p)
```

- [ ] **Step 7: 编写执行逻辑和辅助方法**

```python
    def _on_tab_change(self, event) -> None:
        """切换标签页时清除预览。"""
        self._preview.unload()
        self._update_preview_label()
        self._status_var.set("就绪")

    def _preview_prev(self) -> None:
        self._preview.prev_page()
        self._update_preview_label()

    def _preview_next(self) -> None:
        self._preview.next_page()
        self._update_preview_label()

    def _update_preview_label(self) -> None:
        if self._preview.page_count > 0:
            self._lbl_page.configure(
                text=f"{self._preview.current_page} / {self._preview.page_count}"
            )
        else:
            self._lbl_page.configure(text="0 / 0")

    def _set_ui_state(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self._btn_execute.configure(state=state)

    def _run_operation(self) -> None:
        """根据当前标签页执行对应操作（在后台线程中执行）。"""
        tab_idx = self._notebook.index("current")

        if tab_idx == 0:  # 合并
            self._run_in_thread(self._do_merge)
        elif tab_idx == 1:  # 拆分
            self._run_in_thread(self._do_split)
        elif tab_idx == 2:  # 提取
            self._run_in_thread(self._do_extract)
        elif tab_idx == 3:  # 旋转
            self._run_in_thread(self._do_rotate)

    def _run_in_thread(self, func) -> None:
        """在后台线程中执行操作，显示进度条。"""
        self._set_ui_state(False)
        self._progress.start()
        self._status_var.set("处理中...")

        def wrapper():
            try:
                func()
                self.root.after(0, lambda: self._on_done(None))
            except PDFError as e:
                self.root.after(0, lambda: self._on_done(str(e)))
            except Exception as e:
                self.root.after(0, lambda: self._on_done(f"未知错误: {str(e)}"))

        threading.Thread(target=wrapper, daemon=True).start()

    def _on_done(self, error: str | None) -> None:
        self._progress.stop()
        self._set_ui_state(True)
        if error:
            self._status_var.set(f"失败: {error}")
            messagebox.showerror("操作失败", error)
        else:
            self._status_var.set("完成")

    # ── 具体操作 ──

    def _do_merge(self) -> None:
        if len(self._file_list) < 1:
            raise PDFError("请先添加至少一个 PDF 文件")
        out = filedialog.asksaveasfilename(
            title="保存合并后的 PDF",
            defaultextension=".pdf",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if not out:
            return
        merge_pdfs(self._file_list, out)
        messagebox.showinfo("完成", f"已保存到:\n{out}")

    def _do_split(self) -> None:
        src = self._split_path_var.get()
        if not src:
            raise PDFError("请先选择要拆分的 PDF 文件")

        try:
            per_file = int(self._split_pages_var.get())
        except ValueError:
            raise PDFError("请输入有效的页数")

        out_dir = self._split_out_var.get()
        if not out_dir:
            raise PDFError("请先选择输出目录")

        results = split_pdf(src, out_dir, per_file)
        messagebox.showinfo("完成",
            f"已生成 {len(results)} 个文件:\n{chr(10).join(results[:5])}"
            + (f"\n... 等" if len(results) > 5 else ""))

    def _do_extract(self) -> None:
        src = self._extract_path_var.get()
        if not src:
            raise PDFError("请先选择 PDF 文件")

        raw = self._extract_range_var.get().strip()
        if not raw:
            raise PDFError("请输入页码范围")
        ranges = [r.strip() for r in raw.split(",") if r.strip()]

        out = self._extract_out_var.get()
        if not out:
            raise PDFError("请先选择保存位置")

        extract_pages(src, out, ranges)
        messagebox.showinfo("完成", f"已保存到:\n{out}")

    def _do_rotate(self) -> None:
        src = self._rotate_path_var.get()
        if not src:
            raise PDFError("请先选择 PDF 文件")

        pages_raw = self._rotate_pages_var.get().strip()
        try:
            angle = int(self._rotate_angle_var.get())
        except ValueError:
            raise PDFError("请选择旋转角度")

        out = self._rotate_out_var.get()
        if not out:
            raise PDFError("请先选择保存位置")

        if pages_raw:
            page_nums = [int(p.strip()) for p in pages_raw.split(",") if p.strip()]
        else:
            # 留空 = 全部页面
            page_nums = list(range(1, get_pdf_info(src)["pages"] + 1))

        rotate_pages(src, out, page_nums, angle)
        messagebox.showinfo("完成", f"已保存到:\n{out}")
```

- [ ] **Step 8: Commit**

```bash
git add gui.py
git commit -m "feat: add main GUI with merge/split/extract/rotate tabs"
```

---

### Task 5: 入口文件

**Files:**
- Create: `main.py`

**Interfaces:**
- Consumes: `gui.PDFToolApp`, `tkinter`
- Produces: 可运行的应用程序入口

- [ ] **Step 1: 编写 main.py**

```python
"""PDF 工具箱 — 入口。"""

import tkinter as tk
from gui import PDFToolApp


def main():
    root = tk.Tk()
    # 设置窗口图标（可选）
    try:
        root.iconbitmap(default="icon.ico")
    except Exception:
        pass

    app = PDFToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 手动启动验证**

```bash
python main.py
```

Expected: 窗口打开，四个标签页可切换，文件对话框可打开

- [ ] **Step 3: Commit**

```bash
git add main.py
git commit -m "feat: add application entry point"
```

---

### Task 6: 打包与最终验证

**Files:**
- Modify: `build.bat`（已创建，无需修改）
- Create: `dist/PDF工具箱.exe`（打包产物）

**Interfaces:**
- Consumes: 所有源文件
- Produces: 独立的 Windows .exe 文件

- [ ] **Step 1: 运行打包脚本**

```bash
cd D:\CPH\APP\pdf
build.bat
```

Expected: 在 `dist/` 目录生成 `PDF工具箱.exe`

- [ ] **Step 2: 验证 .exe 独立运行**

启动 `dist/PDF工具箱.exe`，验证：
- 窗口正常打开（无控制台黑窗）
- 四个标签页可切换
- 添加 PDF 文件正常显示
- 预览面板可渲染页面
- 合并功能正常
- 拆分功能正常

- [ ] **Step 3: 提交打包**

提交最终打包产物。

---

## 验证清单

完成全部任务后，执行以下验证：

1. **合并**: 选择 3 个 PDF → 调整顺序 → 合并 → 用 PDF 阅读器打开验证
2. **拆分**: 选一个 10 页 PDF → 每 3 页拆分 → 检查输出文件数量
3. **提取**: 选一个 PDF → 输入 `1-3,5` → 提取 → 验证输出为 4 页
4. **旋转**: 选一个 PDF → 第 1 页旋转 90° → 验证页面方向
5. **预览**: 在各标签页选择 PDF 后确认右侧显示缩略图
6. **错误处理**: 尝试打开非 PDF 文件、输入非法页码范围，确认有提示
7. **打包验证**: 将 `PDF工具箱.exe` 复制到没有 Python 环境的机器上运行
