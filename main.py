"""PDF 工具箱 — 入口。"""

from tkinterdnd2 import TkinterDnD
from gui import PDFToolApp


def main():
    root = TkinterDnD.Tk()
    try:
        root.iconbitmap(default="icon.ico")
    except Exception:
        pass

    PDFToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
