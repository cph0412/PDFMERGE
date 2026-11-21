"""PDF 预览窗口 — 独立弹窗，可滚动缩略图展示源页面与排版结果。"""

import fitz
from pathlib import Path
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk

from pdf_ops import build_2up_doc, build_4up_doc, PDFError

THUMB_W = 280
THUMB_H = 380
PER_ROW = 3
ZOOM = 1.5
FONT = ("Microsoft YaHei UI", 10)


class PreviewWindow:
    """独立预览弹窗：用可滚动 Canvas 展示 PDF 页面缩略图。"""

    def __init__(self, parent):
        self.win = tk.Toplevel(parent)
        self.win.title("预览")
        self.win.geometry("960x720")
        self.win.minsize(640, 480)

        self._images = []  # 持有 PhotoImage 引用防 GC
        self._closed = False

        # 顶部标题
        self._header = ttk.Label(self.win, text="预览", anchor="w",
                                 font=("Microsoft YaHei UI", 10, "bold"))
        self._header.pack(fill="x", padx=8, pady=(8, 0))

        # 可滚动区域
        wrap = ttk.Frame(self.win)
        wrap.pack(fill="both", expand=True, padx=8, pady=8)
        sb = ttk.Scrollbar(wrap, orient="vertical")
        sb.pack(side="right", fill="y")
        self._canvas = tk.Canvas(wrap, bg="#f5f5f5", highlightthickness=0,
                                 yscrollcommand=sb.set)
        self._canvas.pack(side="left", fill="both", expand=True)
        sb.configure(command=self._canvas.yview)

        self._inner = ttk.Frame(self._canvas)
        self._inner.bind("<Configure>", self._on_inner_configure)
        self._inner_id = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._canvas.bind("<Configure>", self._on_canvas_configure)
        self._canvas.bind_all("<MouseWheel>", self._on_wheel)
        self.win.protocol("WM_DELETE_WINDOW", self._on_close)

    @property
    def closed(self) -> bool:
        return self._closed

    def _on_inner_configure(self, _):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_configure(self, e):
        self._canvas.itemconfig(self._inner_id, width=e.width)

    def _on_wheel(self, e):
        self._canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

    def _on_close(self):
        self._closed = True
        try:
            self._canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
        self.win.destroy()

    def bring_front(self):
        if self._closed:
            return
        self.win.deiconify()
        self.win.lift()
        self.win.focus_force()

    def set_header(self, text):
        self._header.configure(text=text)

    def clear(self):
        for w in self._inner.winfo_children():
            w.destroy()
        self._images.clear()

    def _add_section(self, title):
        ttk.Label(self._inner, text=title,
                  font=("Microsoft YaHei UI", 11, "bold")).pack(
            anchor="w", padx=4, pady=(14, 4))

    def _add_thumbnails(self, pages):
        """渲染缩略图。pages: list[(doc, pno, caption)]，渲染期间 doc 必须保持打开。"""
        row = None
        for idx, (doc, pno, caption) in enumerate(pages):
            if idx % PER_ROW == 0:
                row = ttk.Frame(self._inner)
                row.pack(fill="x", padx=4, pady=2)
            cell = ttk.Frame(row)
            cell.pack(side="left", padx=6, pady=4)
            try:
                pix = doc[pno].get_pixmap(matrix=fitz.Matrix(ZOOM, ZOOM))
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                img.thumbnail((THUMB_W, THUMB_H), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                self._images.append(photo)
                ttk.Label(cell, image=photo, background="#ffffff").pack()
                ttk.Label(cell, text=caption, font=FONT).pack()
            except Exception:
                ttk.Label(cell, text=f"[渲染失败] {caption}", foreground="#c00").pack()

    def _collect_source_pages(self, file_paths):
        """打开各源 PDF，返回 (pages, docs)。调用方负责关闭 docs。"""
        docs = []
        pages = []
        for path in file_paths:
            doc = fitz.open(path)
            docs.append(doc)
            stem = Path(path).stem
            for pno in range(doc.page_count):
                pages.append((doc, pno, f"{stem}  P{pno + 1}"))
        return pages, docs

    def show_message(self, title, msg):
        self.clear()
        self.set_header(title)
        self._add_section(msg)

    def show_sources(self, file_paths, header="预览"):
        """展示多个源 PDF 的全部页面。"""
        self.clear()
        self.set_header(header)
        if not file_paths:
            self._add_section("未选择文件")
            return
        self._add_section("源文件页面")
        pages, docs = self._collect_source_pages(file_paths)
        try:
            self._add_thumbnails(pages)
        finally:
            for d in docs:
                d.close()

    def show_merge(self, file_paths, use_2up, use_4up=False):
        """合并预览：源文件全部页面 + 排版结果。"""
        self.clear()
        self.set_header("预览 — 合并")
        if not file_paths:
            self._add_section("未添加文件")
            return

        # 源页
        self._add_section("源文件页面")
        pages, docs = self._collect_source_pages(file_paths)
        try:
            self._add_thumbnails(pages)
        finally:
            for d in docs:
                d.close()

        # 四合一结果
        if use_4up:
            self._add_section("四合一结果（四页排列在一张 A4）")
            out = None
            try:
                out = build_4up_doc(file_paths)
                result_pages = [(out, pno, f"结果 P{pno + 1}")
                                for pno in range(out.page_count)]
                self._add_thumbnails(result_pages)
            except PDFError as e:
                ttk.Label(self._inner, text=f"[四合一失败] {e}",
                          foreground="#c00", font=FONT).pack(anchor="w", padx=8)
            finally:
                if out is not None:
                    out.close()
        # 二合一结果
        elif use_2up:
            self._add_section("二合一结果（每两页上下排列在一张 A4）")
            out = None
            try:
                out = build_2up_doc(file_paths)
                result_pages = [(out, pno, f"结果 P{pno + 1}")
                                for pno in range(out.page_count)]
                self._add_thumbnails(result_pages)
            except PDFError as e:
                ttk.Label(self._inner, text=f"[二合一失败] {e}",
                          foreground="#c00", font=FONT).pack(anchor="w", padx=8)
            finally:
                if out is not None:
                    out.close()
