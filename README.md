# Word → PDF

把 Word 文档（`.docx`）**原样导出**为带书签的 PDF。页面内容、排版、字体一律不改，只在 PDF 里挂一棵可点击的书签导航树。

## 功能

- 原样导出：调用本机 Microsoft Word 引擎生成 PDF，页面内容字节级不变
- 自动书签：优先读取文档自带的目录（`toc 1/2/3` 样式段落）生成三级书签；文档没有目录时，回退到 Word 内置「标题 1/2/3」样式的书签
- 页码校验：逐条核对书签标题与目标页，页码对不上会在邻近页微调，找不到则中止，避免错误书签
- 拖入文件：把 `.docx` 直接拖进「源文件」输入框
- 输出目录：默认保存到源文件所在目录，可自行修改；换文件时自动刷新为该文件所在目录
- 连续进度条：转换期间平滑推进，不按阶段跳变
- 可缩放窗口：支持自由调整大小

## 使用方法

### 运行 exe（推荐）

双击 `WordToPDF.exe`（位于上一级目录），在界面里选文件或拖入，点「开始转换」即可。exe 自带全部运行时，无需安装 Python。

前置条件：**电脑必须安装 Microsoft Word**（导出依赖 Word 引擎）。

### 运行源码

需要 Python 3.13 及以下依赖：

```bash
pip install customtkinter tkinterdnd2 pymupdf python-docx pywin32 pillow
python WordToPDF.py
```

## 文件说明

| 文件 | 用途 |
|---|---|
| `WordToPDF.py` | 主程序：GUI + 转换核心逻辑 |
| `app_icon.ico` | 应用图标（打包时嵌入 exe） |
| `app_icon.png` | 图标预览 |
| `docx2pdf_bookmarked.py` | 早期命令行版工具，无 GUI，可留作备用 |

## 打包

在项目目录执行：

```bash
pyinstaller --onefile --windowed --name WordToPDF \
  --collect-data customtkinter \
  --add-data "app_icon.ico;." \
  --exclude-module pandas --exclude-module scipy \
  --exclude-module matplotlib --exclude-module PyQt6 \
  --exclude-module numpy --exclude-module PIL \
  WordToPDF.py
```

注意：打包前务必确认这里的 `WordToPDF.py` 是最新版本，否则会打包进旧代码。

## 实现要点

- **书签来源**：Word 的 `ExportAsFixedFormat` 只认内置标题样式或可见 Word 书签，对自动目录生成的 `_Toc` 隐藏书签无能为力（`CreateBookmarks=1/2` 都导出 0 条）。因此改为从 `toc N` 段落直接提取「层级 / 标题 / 页码」，再用 PyMuPDF 的 `set_toc()` 增量写入 PDF，页面内容不受影响。
- **中文路径**：Word COM 打不开中文长路径下的 docx，转换时先复制到纯 ASCII 临时目录再处理。
- **拖放**：用 `tkinterdnd2`，窗口类同时继承 `ctk.CTk` 与 `TkinterDnD.DnDWrapper`，源文件框注册 `DND_FILES` 拖放目标。
- **进度条**：`_animate_progress()` 每 40 ms 更新一次，转换中渐近到 90%，完成后快速冲满；真实阶段回调只更新状态文字。
- **输出校验**：写入书签后比对页数与每页文本，确保与导出前完全一致。
