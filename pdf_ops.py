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
    """合并多个 PDF 文件为一个。"""
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
    """按每 N 页拆分 PDF。返回生成的 PDF 文件路径列表。"""
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
    """提取指定页码范围为一个新 PDF。页码从 1 开始。"""
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
    """旋转指定页面。页码从 1 开始，角度支持 90/180/270。"""
    if angle not in (90, 180, 270):
        raise PDFError(f"不支持的旋转角度: {angle}（仅支持 90, 180, 270）")

    doc = fitz.open(file_path)
    if doc.is_encrypted:
        doc.close()
        raise PDFError(f"文件已加密，无法旋转: {file_path}")

    try:
        for page_num in page_nums:
            page = doc[page_num - 1]
            page.set_rotation((page.rotation + angle) % 360)

        doc.save(output_path, garbage=4, deflate=True)
    finally:
        doc.close()


def build_2up_doc(file_paths: list[str]) -> fitz.Document:
    """在内存中构建二合一排版的 PDF 文档（不落盘）。

    每两页上下排列在一张 A4 竖版上，保持原页比例缩放并居中；非 A4 比例
    （横向页、信纸等）的页面不会被拉伸变形。调用方负责关闭返回的 doc。
    """
    if not file_paths:
        raise PDFError("没有选择任何文件")

    A4_W = 595.28
    A4_H = 841.89

    pages: list[tuple[fitz.Document, int]] = []
    docs: list[fitz.Document] = []

    try:
        for path in file_paths:
            doc = fitz.open(path)
            if doc.is_encrypted:
                doc.close()
                raise PDFError(f"文件已加密，无法处理: {path}")
            docs.append(doc)
            for pno in range(doc.page_count):
                pages.append((doc, pno))

        if not pages:
            raise PDFError("所有文件都无页面")

        out = fitz.open()
        success = False

        try:
            i = 0
            while i < len(pages):
                new_page = out.new_page(width=A4_W, height=A4_H)

                # 上半部 — 保持比例缩放并居中
                src_doc, src_pno = pages[i]
                rect_top = fitz.Rect(0, 0, A4_W, A4_H / 2)
                new_page.show_pdf_page(rect_top, src_doc, src_pno, keep_proportion=True)

                # 下半部
                if i + 1 < len(pages):
                    src_doc2, src_pno2 = pages[i + 1]
                    rect_bottom = fitz.Rect(0, A4_H / 2, A4_W, A4_H)
                    new_page.show_pdf_page(rect_bottom, src_doc2, src_pno2, keep_proportion=True)

                i += 2

            success = True
            return out
        finally:
            if not success:
                out.close()
    except PDFError:
        raise
    except Exception as e:
        raise PDFError(f"二合一排版失败: {str(e)}")
    finally:
        # 源 doc 在 show_pdf_page 时内容已复制进 out，可安全关闭
        for d in docs:
            try:
                d.close()
            except Exception:
                pass


def merge_2up(file_paths: list[str], output_path: str) -> None:
    """二合一排版并保存到文件。"""
    out = build_2up_doc(file_paths)  # 可能抛出 PDFError
    try:
        out.save(output_path, garbage=4, deflate=True)
    except PDFError:
        raise
    except Exception as e:
        raise PDFError(f"二合一排版失败: {str(e)}")
    finally:
        out.close()


def build_4up_doc(file_paths: list[str]) -> fitz.Document:
    """在内存中构建四合一排版的 PDF 文档（不落盘）。

    每四页按 2×2 网格排列在一张 A4 竖版上：
    ┌──────┬──────┐
    │ 第1页 │ 第2页 │
    ├──────┼──────┤
    │ 第3页 │ 第4页 │
    └──────┴──────┘
    保持原页比例缩放并居中。调用方负责关闭返回的 doc。
    """
    if not file_paths:
        raise PDFError("没有选择任何文件")

    A4_W = 595.28
    A4_H = 841.89
    HALF_W = A4_W / 2
    HALF_H = A4_H / 2

    # 四个格子的 Rect：左上、右上、左下、右下
    CELLS = [
        fitz.Rect(0, 0, HALF_W, HALF_H),           # 左上
        fitz.Rect(HALF_W, 0, A4_W, HALF_H),        # 右上
        fitz.Rect(0, HALF_H, HALF_W, A4_H),        # 左下
        fitz.Rect(HALF_W, HALF_H, A4_W, A4_H),     # 右下
    ]

    pages: list[tuple[fitz.Document, int]] = []
    docs: list[fitz.Document] = []

    try:
        for path in file_paths:
            doc = fitz.open(path)
            if doc.is_encrypted:
                doc.close()
                raise PDFError(f"文件已加密，无法处理: {path}")
            docs.append(doc)
            for pno in range(doc.page_count):
                pages.append((doc, pno))

        if not pages:
            raise PDFError("所有文件都无页面")

        out = fitz.open()
        success = False

        try:
            i = 0
            while i < len(pages):
                new_page = out.new_page(width=A4_W, height=A4_H)

                for cell_idx, rect in enumerate(CELLS):
                    if i + cell_idx < len(pages):
                        src_doc, src_pno = pages[i + cell_idx]
                        new_page.show_pdf_page(rect, src_doc, src_pno,
                                               keep_proportion=True)

                i += 4

            success = True
            return out
        finally:
            if not success:
                out.close()
    except PDFError:
        raise
    except Exception as e:
        raise PDFError(f"四合一排版失败: {str(e)}")
    finally:
        for d in docs:
            try:
                d.close()
            except Exception:
                pass


def merge_4up(file_paths: list[str], output_path: str) -> None:
    """四合一排版并保存到文件。"""
    out = build_4up_doc(file_paths)  # 可能抛出 PDFError
    try:
        out.save(output_path, garbage=4, deflate=True)
    except PDFError:
        raise
    except Exception as e:
        raise PDFError(f"四合一排版失败: {str(e)}")
    finally:
        out.close()
