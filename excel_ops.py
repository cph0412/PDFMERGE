"""Excel 合并操作模块 — 纯逻辑层，不依赖 GUI。

支持 .xlsx/.xlsm（openpyxl）与旧版 .xls（xlrd）。

性能说明：openpyxl 本身较慢（读/写约 1 万行/秒量级），大文件合并的
物理耗时无法完全消除；本模块通过
  * read_only 读 + write_only 写（流式，内存恒定，不随总行数增长）
  * 每处理一批行回调一次进度（progress）
  * 支持中途取消（cancel_event）
来保证界面有反馈、可中止，避免"看起来卡死"。
"""

from pathlib import Path
from threading import Event
from typing import Callable, Iterator, Optional

from openpyxl import Workbook, load_workbook

try:
    import xlrd
except ImportError:  # 只在真正需要读 .xls 时才需要
    xlrd = None

# 每处理这么多行回调一次进度（避免回调过密拖慢合并）
_PROGRESS_STEP = 5000


class ExcelError(Exception):
    """Excel 操作异常。"""
    pass


class ExcelCancelled(Exception):
    """用户取消合并（单独基类，GUI 需静默处理，不弹错误框）。"""
    pass


ProgressFn = Optional[Callable[[int, int, str, int], None]]
# progress(file_index_1based, file_count, 当前文件名, 当前文件已处理行数)


def _check_cancel(cancel_event: Optional[Event]) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise ExcelCancelled("操作已取消")


def _report(progress: ProgressFn, idx: int, total: int, name: str, rows: int) -> None:
    if progress is not None:
        try:
            progress(idx, total, name, rows)
        except Exception:
            pass  # 进度回调异常不应中断合并


def _ensure_xls_reader(file_paths: list[str]) -> None:
    """若包含旧版 .xls 且未安装 xlrd，提前明确报错。"""
    if any(str(p).lower().endswith(".xls") for p in file_paths) and xlrd is None:
        raise ExcelError(
            "检测到旧版 .xls 文件，但未安装 xlrd，请先执行 "
            "pip install xlrd 或重新打包程序"
        )


def _safe_sheet_name(name: str, existing: set[str]) -> str:
    """生成不重复的 Sheet 名称（Excel 限制 31 字符）。"""
    base = name[:28]  # 留余量给后缀
    candidate = base
    counter = 1
    while candidate in existing:
        candidate = f"{base[:25]}_{counter}"
        counter += 1
    return candidate


# ─────────────────────────── 统一行数据源 ───────────────────────────

def _is_xls(path: str) -> bool:
    return str(path).lower().endswith(".xls")


def _open_book(path: str):
    """返回 (kind, book)。kind: 'xls' 或 'xlsx'。book 需调用方释放。"""
    if _is_xls(path):
        return "xls", xlrd.open_workbook(str(path), on_demand=True)
    return "xlsx", load_workbook(path, read_only=True, data_only=True)


def _close_book(kind: str, book) -> None:
    try:
        if kind == "xls":
            book.release_resources()
        else:
            book.close()
    except Exception:
        pass


def _sheet_names(kind: str, book) -> list[str]:
    return book.sheet_names() if kind == "xls" else book.sheetnames


def _xls_cell_value(sh, r: int, c: int):
    """把 xlrd 单元格转成与 openpyxl 尽量一致的值。"""
    ct = sh.cell_type(r, c)
    v = sh.cell_value(r, c)
    if ct == xlrd.XL_CELL_DATE:
        try:
            return xlrd.xldate_as_datetime(v, sh.book.datemode)
        except Exception:
            return v
    if ct == xlrd.XL_CELL_NUMBER:
        # 整数值去掉尾缀 .0，便于后续比对/显示
        if isinstance(v, float) and v.is_integer() and abs(v) < 2 ** 53:
            return int(v)
        return v
    if ct in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_ERROR):
        return None
    if ct == xlrd.XL_CELL_BOOLEAN:
        return bool(v)
    return v


def _rows_of_sheet(kind: str, book, sheet_name: str) -> Iterator[list]:
    """按行产出某 Sheet 的全部行（list 形式），逐行流式。"""
    if kind == "xls":
        sh = book.sheet_by_name(sheet_name)
        ncols = sh.ncols
        for r in range(sh.nrows):
            yield [_xls_cell_value(sh, r, c) for c in range(ncols)]
    else:
        ws = book[sheet_name]
        for row in ws.iter_rows(values_only=True):
            yield list(row)


# ───────────────────────────── 合并入口 ─────────────────────────────

def merge_to_sheets(
    file_paths: list[str],
    output_path: str,
    progress: ProgressFn = None,
    cancel_event: Optional[Event] = None,
) -> None:
    """每个文件 → 一个 Sheet，Sheet 名取自文件名（去扩展名）。

    流式实现：read_only 读源文件 + write_only 写目标，内存占用恒定。
    """
    if not file_paths:
        raise ExcelError("没有选择任何文件")
    _ensure_xls_reader(file_paths)

    total = len(file_paths)
    wb = Workbook(write_only=True)
    used_names: set[str] = set()
    current_name = ""

    try:
        for idx, path in enumerate(file_paths):
            _check_cancel(cancel_event)
            current_name = Path(path).name
            try:
                kind, book = _open_book(path)
            except Exception as e:
                raise ExcelError(f"无法打开文件 {current_name}: {e}")

            try:
                sheets = _sheet_names(kind, book)
                stem = Path(path).stem
                for sheet in sheets:
                    _check_cancel(cancel_event)
                    # Sheet 名取自文件名
                    name = stem if len(sheets) == 1 else f"{stem}_{sheet}"
                    name = _safe_sheet_name(name, used_names)
                    used_names.add(name)

                    ws_new = wb.create_sheet(title=name)
                    rows_done = 0
                    for row in _rows_of_sheet(kind, book, sheet):
                        ws_new.append(row)
                        rows_done += 1
                        if rows_done % _PROGRESS_STEP == 0:
                            _check_cancel(cancel_event)
                            _report(progress, idx + 1, total, current_name, rows_done)
                    _report(progress, idx + 1, total, current_name, rows_done)
            finally:
                _close_book(kind, book)

        _check_cancel(cancel_event)
        wb.save(output_path)
    except ExcelCancelled:
        raise
    except ExcelError:
        raise
    except Exception as e:
        raise ExcelError(f"合并失败（文件 {current_name}）: {str(e)}")
    finally:
        wb.close()


def merge_to_one_sheet(
    file_paths: list[str],
    output_path: str,
    progress: ProgressFn = None,
    cancel_event: Optional[Event] = None,
) -> None:
    """所有文件行追加到同一 Sheet，保留第一个非空文件的表头，按列名对齐。"""
    if not file_paths:
        raise ExcelError("没有选择任何文件")
    _ensure_xls_reader(file_paths)

    total = len(file_paths)
    wb = Workbook(write_only=True)
    # write_only 模式没有 active sheet，必须显式创建
    ws = wb.create_sheet(title="合并数据")

    header_written = False
    header_cols: list[str] = []  # 第一个文件的表头列名
    current_name = ""

    try:
        for idx, path in enumerate(file_paths):
            _check_cancel(cancel_event)
            current_name = Path(path).name
            try:
                kind, book = _open_book(path)
            except Exception as e:
                raise ExcelError(f"无法打开文件 {current_name}: {e}")

            try:
                sheets = _sheet_names(kind, book)
                if not sheets:
                    continue
                rows_iter = _rows_of_sheet(kind, book, sheets[0])

                if not header_written:
                    # 第一个非空文件：写入表头 + 全部数据行
                    rows_done = 0
                    for ri, row in enumerate(rows_iter):
                        if ri == 0:
                            header_cols = [
                                str(c) if c is not None else "" for c in row
                            ]
                        ws.append(row)
                        rows_done += 1
                        if rows_done % _PROGRESS_STEP == 0:
                            _check_cancel(cancel_event)
                            _report(progress, idx + 1, total, current_name, rows_done)
                    if rows_done:
                        header_written = True
                    _report(progress, idx + 1, total, current_name, rows_done)
                else:
                    # 后续文件：跳过表头，按列名对齐写入
                    col_map: dict[int, int] = {}
                    rows_done = 0
                    for ri, row in enumerate(rows_iter):
                        if ri == 0:
                            src_header = [
                                str(c) if c is not None else "" for c in row
                            ]
                            col_map = {
                                si: header_cols.index(cn)
                                for si, cn in enumerate(src_header)
                                if cn in header_cols
                            }
                            continue  # 表头行不写入
                        aligned = [""] * len(header_cols)
                        for src_i, val in enumerate(row):
                            if src_i in col_map:
                                aligned[col_map[src_i]] = val
                        ws.append(aligned)
                        rows_done += 1
                        if rows_done % _PROGRESS_STEP == 0:
                            _check_cancel(cancel_event)
                            _report(progress, idx + 1, total, current_name, rows_done)
                    _report(progress, idx + 1, total, current_name, rows_done)
            finally:
                _close_book(kind, book)

        if not header_written:
            raise ExcelError("所有文件都为空")

        _check_cancel(cancel_event)
        wb.save(output_path)
    except ExcelCancelled:
        raise
    except ExcelError:
        raise
    except Exception as e:
        raise ExcelError(f"合并失败（文件 {current_name}）: {str(e)}")
    finally:
        wb.close()


def merge_copy_sheets(
    file_paths: list[str],
    output_path: str,
    progress: ProgressFn = None,
    cancel_event: Optional[Event] = None,
) -> None:
    """整表复制合并（保留全部格式）——通过本机 Excel COM 实现。

    把每个源文件的**全部 Sheet** 按顺序原样复制进一个新工作簿：
    字体/边框/列宽/行高/数字格式/合并单元格/打印页面设置等全部保留。
    要求：本机装有 Microsoft Excel，且已安装 pywin32。

    注意：本模式 progress 的 rows 参数表示"当前文件已复制的 Sheet 数"；
    取消检查在每个文件/每个 Sheet 之间进行。
    """
    if not file_paths:
        raise ExcelError("没有选择任何文件")

    try:
        import win32com.client
    except ImportError:
        raise ExcelError(
            "整表复制需要 pywin32，请先执行 pip install pywin32 "
            "（或重新打包程序）"
        )

    total = len(file_paths)
    app = None
    dst = None
    src = None
    current_name = ""
    copied_any = False

    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        try:
            app.AskToUpdateLinks = False
        except Exception:
            pass
        try:
            app.AutomationSecurity = 3  # 禁用源文件宏
        except Exception:
            pass

        dst = app.Workbooks.Add()
        # 记住新建工作簿自带的空白表（稍后删除），用名称+序号定位
        blank_names = [str(ws.Name) for ws in dst.Worksheets]

        for idx, path in enumerate(file_paths):
            _check_cancel(cancel_event)
            current_name = Path(path).name
            try:
                src = app.Workbooks.Open(
                    str(path), ReadOnly=True, UpdateLinks=0
                )
            except Exception as e:
                raise ExcelError(f"无法用 Excel 打开文件 {current_name}: {e}")

            try:
                n_sheets = src.Worksheets.Count
                for j in range(1, n_sheets + 1):
                    _check_cancel(cancel_event)
                    # 复制到目标工作簿末尾（注意：Excel 2007 后绑定时
                    # Copy(After=...) 命名参数不生效，必须用位置参数）
                    src.Worksheets(j).Copy(
                        None, dst.Worksheets(dst.Worksheets.Count)
                    )
                    copied_any = True
                    _report(progress, idx + 1, total, current_name, j)
            finally:
                try:
                    src.Close(SaveChanges=False)
                except Exception:
                    pass
                src = None

        if not copied_any:
            raise ExcelError("没有可复制的 Sheet（文件都为空？）")

        # 删除目标工作簿自带的空白 Sheet（Excel 2007 默认 3 个）
        for k in range(dst.Worksheets.Count, 0, -1):
            if str(dst.Worksheets(k).Name) in blank_names:
                dst.Worksheets(k).Delete()

        _check_cancel(cancel_event)
        dst.SaveAs(str(output_path), FileFormat=51)  # 51 = .xlsx
    except ExcelCancelled:
        raise
    except ExcelError:
        raise
    except Exception as e:
        raise ExcelError(f"整表复制失败（文件 {current_name}）: {str(e)}")
    finally:
        try:
            if src is not None:
                src.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if dst is not None:
                dst.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
