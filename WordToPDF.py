# -*- coding: utf-8 -*-
"""
Word → PDF 图形界面。

- 原样导出 Word 文档为 PDF，页面内容不做任何改动
- 优先按文档自带目录（toc 1/2/3 样式）建立 PDF 书签，无目录时回退标题样式
- 支持自定义输出目录，默认输出到源文件所在目录
- 支持把 docx 直接拖入源文件框
- 转换在后台线程执行；进度条以平滑动画连续推进，不按阶段跳变
"""
import os
import re
import sys
import shutil
import tempfile
import queue
import threading
import traceback
from tkinter import filedialog, messagebox

import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_FILES


# ----------------------------------------------------------------------------
# 核心转换逻辑（与命令行版一致，增加输出目录与进度回调）
# ----------------------------------------------------------------------------

def _norm(text):
    return re.sub(r"[\s\u3000·.．、，,：:（）()【】\[\]…—\-]", "", text)


def _extract_toc(docx_path):
    from docx import Document

    doc = Document(docx_path)
    entries = []
    for para in doc.paragraphs:
        style = para.style.name.strip() if para.style else ""
        match = re.fullmatch(r"toc\s*([1-9])", style, re.I)
        if not match:
            continue
        text = para.text.strip()
        if "\t" not in text:
            continue
        title, page_str = text.rsplit("\t", 1)
        if not page_str.strip().isdigit():
            continue
        entries.append([int(match.group(1)), title.strip(), int(page_str.strip())])
    return entries


def _verify_pages(pdf, entries):
    page_count = len(pdf)
    texts = [_norm(page.get_text()) for page in pdf]
    verified = []
    adjusted = []
    for level, title, page in entries:
        key = re.sub(r"^[\d.．\s]+", "", _norm(title))
        probe = key[:8] if len(key) > 8 else key
        if not probe:
            verified.append([level, title, page])
            continue
        if 1 <= page <= page_count and probe in texts[page - 1]:
            verified.append([level, title, page])
            continue
        found = None
        for delta in (-1, 1, -2, 2):
            candidate = page + delta
            if 1 <= candidate <= page_count and probe in texts[candidate - 1]:
                found = candidate
                break
        if found is None:
            raise RuntimeError(
                '页码校验失败："%s" 标注第 %d 页，但附近页面找不到该标题'
                % (title, page)
            )
        adjusted.append((title, page, found))
        verified.append([level, title, found])
    return verified, adjusted


def convert(src, output_dir=None, progress=None):
    """执行转换。

    src: 源 docx 路径
    output_dir: 输出目录，为空则使用源文件所在目录
    progress: 回调 fn(percent:int, message:str)，在工作线程中调用
    返回 (输出路径, 书签数, 页数, 页码微调列表)
    """
    def report(pct, msg):
        if progress:
            progress(pct, msg)

    src = os.path.abspath(src)
    if not os.path.isfile(src):
        raise RuntimeError("文件不存在：" + src)
    if not src.lower().endswith(".docx"):
        raise RuntimeError("仅支持 .docx 文件：" + src)

    target_dir = os.path.abspath(output_dir) if output_dir else os.path.dirname(src)
    if not os.path.isdir(target_dir):
        raise RuntimeError("输出目录不存在：" + target_dir)

    tmpdir = tempfile.mkdtemp(prefix="word2pdf_")
    try:
        report(8, "正在准备临时文件…")
        ascii_docx = os.path.join(tmpdir, "input.docx")
        shutil.copy2(src, ascii_docx)

        report(25, "正在用 Word 导出 PDF…")
        import pythoncom
        import win32com.client as win32

        pythoncom.CoInitialize()
        try:
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            try:
                doc = word.Documents.Open(ascii_docx, ReadOnly=True)
                base_pdf = os.path.join(tmpdir, "base.pdf")
                doc.ExportAsFixedFormat(
                    OutputFileName=base_pdf,
                    ExportFormat=17,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    Range=0,
                    Item=0,
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=1,
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=False,
                )
                doc.Close(False)
            finally:
                word.Quit()
        finally:
            pythoncom.CoUninitialize()

        report(55, "正在读取文档目录…")
        import pymupdf

        pdf = pymupdf.open(base_pdf)
        page_count = len(pdf)
        entries = _extract_toc(ascii_docx)
        adjusted = []
        if entries:
            report(68, "正在校验书签页码…")
            entries, adjusted = _verify_pages(pdf, entries)
        pdf.close()

        report(80, "正在写入书签…")
        name = os.path.splitext(os.path.basename(src))[0] + ".pdf"
        out_pdf = os.path.join(target_dir, name)
        shutil.copy2(base_pdf, out_pdf)

        if entries:
            target = pymupdf.open(out_pdf)
            before_pages = len(target)
            before_text = [page.get_text() for page in target]
            target.set_toc(entries)
            try:
                target.saveIncr()
            except AttributeError:
                target.save(
                    out_pdf, incremental=True, encryption=pymupdf.PDF_ENCRYPT_KEEP
                )
            target.close()
            verify = pymupdf.open(out_pdf)
            assert len(verify) == before_pages, "页数发生变化"
            assert [p.get_text() for p in verify] == before_text, "页面文本发生变化"
            bookmark_count = len(verify.get_toc())
            verify.close()
        else:
            verify = pymupdf.open(out_pdf)
            bookmark_count = len(verify.get_toc())
            verify.close()

        report(100, "转换完成")
        return out_pdf, bookmark_count, page_count, adjusted
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ----------------------------------------------------------------------------
# 图形界面
# ----------------------------------------------------------------------------

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


def _resource_path(rel):
    """定位打包后的资源文件。"""
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, rel)


def _parse_drop_data(data):
    """解析拖放路径：带空格的路径被 tkinterdnd2 用花括号包裹。"""
    paths = []
    for match in re.findall(r"\{([^}]*)\}|(\S+)", data):
        path = match[0] or match[1]
        if path:
            paths.append(path)
    return paths


class App(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)
        self.title("Word → PDF")
        self.geometry("520x440")
        self.minsize(440, 400)
        self.resizable(True, True)
        try:
            self.iconbitmap(_resource_path("app_icon.ico"))
        except Exception:
            pass

        self.queue = queue.Queue()
        self.running = False
        self._display = 0.0      # 进度条当前显示值 0..1
        self._finishing = False  # 完成后冲满阶段

        self._build_ui()
        self.after(60, self._poll_queue)
        self.after(40, self._animate_progress)

    def _build_ui(self):
        card = ctk.CTkFrame(self, corner_radius=16)
        card.pack(fill="both", expand=True, padx=18, pady=18)

        title = ctk.CTkLabel(
            card,
            text="Word → PDF",
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        title.pack(pady=(22, 2))
        subtitle = ctk.CTkLabel(
            card,
            text="原样导出 · 自动生成书签 · 支持拖入文件",
            font=ctk.CTkFont(size=12),
            text_color=("gray60", "gray50"),
        )
        subtitle.pack(pady=(0, 22))

        # 源文件（支持拖入）
        ctk.CTkLabel(card, text="源文件（可拖入）", anchor="w").pack(fill="x", padx=24)
        src_row = ctk.CTkFrame(card, fg_color="transparent")
        src_row.pack(fill="x", padx=24, pady=(6, 14))
        self.src_entry = ctk.CTkEntry(
            src_row, placeholder_text="选择或拖入一个 .docx 文件", height=36
        )
        self.src_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(
            src_row, text="选择文件", width=90, height=36, command=self._choose_file
        ).pack(side="right")
        self.src_entry.drop_target_register(DND_FILES)
        self.src_entry.dnd_bind("<<Drop>>", self._on_drop)

        # 输出目录
        ctk.CTkLabel(card, text="输出目录（留空则保存到原目录）", anchor="w").pack(
            fill="x", padx=24
        )
        out_row = ctk.CTkFrame(card, fg_color="transparent")
        out_row.pack(fill="x", padx=24, pady=(6, 20))
        self.out_entry = ctk.CTkEntry(
            out_row, placeholder_text="默认：源文件所在目录", height=36
        )
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(
            out_row, text="选择目录", width=90, height=36, command=self._choose_dir
        ).pack(side="right")

        # 进度：连续平滑动画
        prog_row = ctk.CTkFrame(card, fg_color="transparent")
        prog_row.pack(fill="x", padx=24)
        self.progress = ctk.CTkProgressBar(prog_row, height=10)
        self.progress.pack(side="left", fill="x", expand=True)
        self.progress.set(0)
        self.pct_label = ctk.CTkLabel(prog_row, text="0%", width=44)
        self.pct_label.pack(side="right", padx=(10, 0))

        self.status = ctk.CTkLabel(
            card,
            text="就绪",
            font=ctk.CTkFont(size=12),
            text_color=("gray60", "gray50"),
        )
        self.status.pack(fill="x", padx=24, pady=(8, 18))

        self.start_btn = ctk.CTkButton(
            card,
            text="开始转换",
            height=42,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._start,
        )
        self.start_btn.pack(fill="x", padx=24, pady=(0, 22))

    def _set_source(self, path):
        """统一设置源文件：每次新文件都刷新输出目录为该文件所在目录。"""
        path = os.path.abspath(path)
        self.src_entry.delete(0, "end")
        self.src_entry.insert(0, path)
        self.out_entry.delete(0, "end")
        self.out_entry.insert(0, os.path.dirname(path))
        self.status.configure(text="已选择源文件")

    def _choose_file(self):
        path = filedialog.askopenfilename(
            title="选择 Word 文档",
            filetypes=[("Word 文档 (*.docx)", "*.docx"), ("所有文件", "*.*")],
        )
        if path:
            self._set_source(path)

    def _on_drop(self, event):
        paths = _parse_drop_data(event.data)
        if not paths:
            return
        path = paths[0]
        if not path.lower().endswith(".docx"):
            messagebox.showwarning("提示", "请拖入 .docx 文件")
            return
        self._set_source(path)

    def _choose_dir(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, path)

    def _start(self):
        if self.running:
            return
        src = self.src_entry.get().strip()
        out_dir = self.out_entry.get().strip()
        if not src:
            messagebox.showwarning("提示", "请先选择源文件")
            return
        if not os.path.isfile(src):
            messagebox.showerror("错误", "源文件不存在：" + src)
            return
        if out_dir and not os.path.isdir(out_dir):
            messagebox.showerror("错误", "输出目录不存在：" + out_dir)
            return

        self.running = True
        self._finishing = False
        self._display = 0.0
        self.start_btn.configure(state="disabled", text="转换中…")
        self.progress.set(0)
        self.pct_label.configure(text="0%")
        self.status.configure(text="正在启动…")

        threading.Thread(
            target=self._work, args=(src, out_dir), daemon=True
        ).start()

    def _work(self, src, out_dir):
        def report(pct, msg):
            # 真实阶段只更新状态文字；百分比由连续动画负责
            self.queue.put(("status", msg))

        try:
            result = convert(src, out_dir or None, report)
            self.queue.put(("done", result))
        except Exception as exc:
            self.queue.put(("error", "%s\n\n%s" % (exc, traceback.format_exc(limit=3))))

    def _animate_progress(self):
        """连续进度条：转换中渐近 90%，完成后快速冲满。"""
        if self.running:
            gap = 0.90 - self._display
            self._display += max(0.0025, gap * 0.035)
            if self._display > 0.90:
                self._display = 0.90
        elif self._finishing:
            self._display += max(0.02, (1.0 - self._display) * 0.28)
            if self._display >= 0.999:
                self._display = 1.0
                self._finishing = False
        else:
            self.after(40, self._animate_progress)
            return

        self.progress.set(self._display)
        self.pct_label.configure(text="%d%%" % int(self._display * 100))
        self.after(40, self._animate_progress)

    def _poll_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                kind = item[0]
                if kind == "status":
                    self.status.configure(text=item[1])
                elif kind == "done":
                    _, (out_pdf, count, pages, adjusted) = item
                    self.running = False
                    self._finishing = True
                    self.start_btn.configure(state="normal", text="开始转换")
                    self.status.configure(text="转换完成")
                    detail = "导出完成：%s\n\n页数：%d    书签：%d 条" % (
                        out_pdf, pages, count
                    )
                    if adjusted:
                        detail += "\n\n页码微调 %d 处" % len(adjusted)
                    if count == 0:
                        detail += "\n\n该文档没有目录段落，也没有内置标题样式，PDF 未带书签。"
                    messagebox.showinfo("完成", detail)
                elif kind == "error":
                    _, text = item
                    self.running = False
                    self.start_btn.configure(state="normal", text="开始转换")
                    self.status.configure(text="转换失败")
                    messagebox.showerror("导出失败", text)
        except queue.Empty:
            pass
        self.after(60, self._poll_queue)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
