#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MarkITDown GUI - 微软 MarkITDown 工具的图形界面版本
将多种文件格式转换为 Markdown 文本

用法：python markitdown_gui.py
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ─── 检查 MarkITDown ─────────────────────────────────────────────────────────
try:
    from markitdown import MarkItDown
except ImportError:
    print("错误：未安装 markitdown，请先运行：pip install markitdown")
    sys.exit(1)


# ─── 转换类别定义 ─────────────────────────────────────────────────────────────
# 格式: (类别名, 文件扩展名列表, 是否支持URL)
FORMAT_CATEGORIES = {
    "📄 文档 (PDF / Word / EPUB)": {
        "extensions": [".pdf", ".docx", ".epub"],
        "accept_url": False,
    },
    "📊 表格 (Excel / CSV)": {
        "extensions": [".xlsx", ".csv"],
        "accept_url": False,
    },
    "📽 演示 (PowerPoint)": {
        "extensions": [".pptx"],
        "accept_url": False,
    },
    "💻 代码 (Jupyter Notebook)": {
        "extensions": [".ipynb", ".py", ".js", ".ts", ".java", ".c", ".cpp",
                       ".h", ".json", ".yaml", ".yml", ".toml", ".xml",
                       ".sh", ".bat", ".sql", ".md", ".html", ".css"],
        "accept_url": False,
    },
    "🌐 网页 (HTML / RSS)": {
        "extensions": [".html", ".htm", ".xml", ".rss"],
        "accept_url": True,
    },
    "🔗 在线链接 (Wikipedia / YouTube)": {
        "extensions": [],
        "accept_url": True,
    },
    "🖼 图片 (OCR / AI 描述)": {
        "extensions": [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp",
                       ".tiff", ".tif"],
        "accept_url": False,
    },
    "🎵 音频 (语音转文字)": {
        "extensions": [".mp3", ".wav", ".flac", ".m4a", ".ogg", ".wma"],
        "accept_url": False,
    },
    "📦 压缩包 (ZIP)": {
        "extensions": [".zip"],
        "accept_url": False,
    },
    "📧 邮件 (Outlook)": {
        "extensions": [".msg"],
        "accept_url": False,
    },
    "📝 纯文本": {
        "extensions": [".txt", ".log", ".ini", ".cfg", ".conf", ".gitignore"],
        "accept_url": False,
    },
}

# 当前选中类别对应的扩展名
CURRENT_EXTENSIONS = []


def get_all_extensions():
    """获取所有支持的扩展名（用于文件对话框过滤）"""
    exts = set()
    for cat in FORMAT_CATEGORIES.values():
        exts.update(cat["extensions"])
    return sorted(exts)


def get_extensions_for_category(category_name):
    """获取指定类别对应的扩展名"""
    return FORMAT_CATEGORIES.get(category_name, {}).get("extensions", [])


def accepts_url(category_name):
    """指定类别是否支持 URL 输入"""
    return FORMAT_CATEGORIES.get(category_name, {}).get("accept_url", False)

# ─── 转换逻辑（后台线程） ──────────────────────────────────────────────────────


def _convert_doc(file_path: str) -> tuple[bool, str]:
    """已移除对 .doc 格式的支持"""
    return False, "当前版本不支持 .doc 格式转换，请转换为 .docx 后重试。"


def convert_file(file_path: str, category: str) -> tuple[bool, str]:
    """
    调用 MarkITDown 转换文件；.doc 格式由 antiword 处理
    返回 (成功与否, 结果文本 或 错误信息)
    """
    # .doc 旧版 Word → 走 antiword
    if file_path.lower().endswith(".doc"):
        return _convert_doc(file_path)

    # 其余格式走 MarkITDown
    try:
        md = MarkItDown()
        result = md.convert(file_path)

        if result and result.text_content:
            title = result.title or os.path.basename(file_path)
            header = f"# {title}\n\n"
            return True, header + result.text_content
        else:
            return False, "转换结果为空，该文件可能不包含可提取的文本内容。"

    except Exception as e:
        return False, f"转换失败：{e}"


def convert_url(url: str) -> tuple[bool, str]:
    """转换在线链接"""
    try:
        md = MarkItDown()
        result = md.convert(url)

        if result and result.text_content:
            title = result.title or url
            header = f"# {title}\n\n"
            return True, header + result.text_content
        else:
            return False, "转换结果为空。"

    except Exception as e:
        return False, f"转换失败：{e}"


# ─── GUI 主界面 ────────────────────────────────────────────────────────────────
class MarkITDownGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MarkITDown GUI - 文件格式转换工具")
        self.root.geometry("900x700")
        self.root.minsize(650, 500)

        # 样式
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self._build_ui()

    def _build_ui(self):
        """构建界面"""
        # === 顶部：标题 ===
        header = ttk.Frame(self.root, padding="12 8 12 4")
        header.pack(fill=tk.X)

        ttk.Label(
            header,
            text="MarkITDown GUI",
            font=("Microsoft YaHei UI", 18, "bold"),
        ).pack(side=tk.LEFT)

        ttk.Label(
            header,
            text="微软 MarkITDown 图形化前端  ·  多种格式 → Markdown",
            font=("Microsoft YaHei UI", 9),
            foreground="#666666",
        ).pack(side=tk.LEFT, padx=(12, 0))

        # === 第一行：转换类别选择 ===
        row1 = ttk.Frame(self.root, padding="12 4 12 8")
        row1.pack(fill=tk.X)

        ttk.Label(row1, text="转换类别：", width=10).pack(side=tk.LEFT)

        self.category_var = tk.StringVar(value="📄 文档 (PDF / Word / EPUB / .doc)")
        self.category_combo = ttk.Combobox(
            row1,
            textvariable=self.category_var,
            values=list(FORMAT_CATEGORIES.keys()),
            state="readonly",
            width=38,
            font=("Microsoft YaHei UI", 10),
        )
        self.category_combo.pack(side=tk.LEFT, padx=(0, 12))
        self.category_combo.bind("<<ComboboxSelected>>", self._on_category_change)

        # 格式提示标签
        self.format_label = ttk.Label(row1, text="", foreground="#888888")
        self.format_label.pack(side=tk.LEFT)

        # === 第二行：文件选择 / URL 输入 ===
        row2 = ttk.Frame(self.root, padding="12 4 12 8")
        row2.pack(fill=tk.X)

        # 文件路径
        ttk.Label(row2, text="文件路径：", width=10).pack(side=tk.LEFT)

        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(row2, textvariable=self.path_var, width=50)
        self.path_entry.pack(side=tk.LEFT, padx=(0, 8))
        self.path_entry.bind("<Return>", lambda e: self._start_convert())

        self.browse_btn = ttk.Button(row2, text="浏览...", command=self._browse_file)
        self.browse_btn.pack(side=tk.LEFT, padx=(0, 4))

        self.url_btn = ttk.Button(row2, text="粘贴链接", command=self._paste_url)
        self.url_btn.pack(side=tk.LEFT, padx=(0, 4))

        # === 第三行：操作按钮 ===
        row3 = ttk.Frame(self.root, padding="12 4 12 8")
        row3.pack(fill=tk.X)

        ttk.Label(row3, text="", width=10).pack(side=tk.LEFT)  # 对齐占位

        self.convert_btn = ttk.Button(
            row3,
            text="▶  开始转换",
            command=self._start_convert,
            width=18,
        )
        self.convert_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.clear_btn = ttk.Button(
            row3,
            text="清空结果",
            command=self._clear_result,
            width=10,
        )
        self.clear_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.save_btn = ttk.Button(
            row3,
            text="💾 保存 .md",
            command=self._save_result,
            width=12,
            state=tk.DISABLED,
        )
        self.save_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.copy_btn = ttk.Button(
            row3,
            text="📋 复制",
            command=self._copy_result,
            width=10,
            state=tk.DISABLED,
        )
        self.copy_btn.pack(side=tk.LEFT)

        # === 进度条 ===
        self.progress = ttk.Progressbar(
            self.root,
            mode="indeterminate",
            length=400,
        )
        self.progress.pack_forget()

        # === 结果文本区 ===
        result_frame = ttk.Frame(self.root, padding="12 4 12 12")
        result_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(result_frame, text="转换结果：", font=("Microsoft YaHei UI", 10, "bold")).pack(
            anchor=tk.W, pady=(0, 4)
        )

        # 文本区域 + 滚动条
        text_container = ttk.Frame(result_frame)
        text_container.pack(fill=tk.BOTH, expand=True)

        self.text = tk.Text(
            text_container,
            wrap=tk.WORD,
            font=("Consolas", 10),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="#ffffff",
            padx=10,
            pady=10,
            relief=tk.FLAT,
            borderwidth=0,
        )
        scrollbar = ttk.Scrollbar(text_container, orient=tk.VERTICAL, command=self.text.yview)
        self.text.configure(yscrollcommand=scrollbar.set)

        self.text.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        text_container.grid_rowconfigure(0, weight=1)
        text_container.grid_columnconfigure(0, weight=1)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪 — 请选择类别并添加文件或链接")
        status_bar = ttk.Label(
            self.root,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            padding="4 2",
            font=("Microsoft YaHei UI", 9),
        )
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        # 初始化
        self._on_category_change(None)

    # ─── 事件处理 ────────────────────────────────────────────────────────────

    def _on_category_change(self, event):
        """切换转换类别"""
        category = self.category_var.get()
        exts = get_extensions_for_category(category)
        url_ok = accepts_url(category)

        # 显示类别对应的格式列表（仅供参考，浏览时不过滤）
        if exts:
            ext_str = ", ".join(exts[:8])
            if len(exts) > 8:
                ext_str += f" ... (共 {len(exts)} 种)"
            self.format_label.config(text=f"支持格式：{ext_str}")
        else:
            self.format_label.config(text="需输入在线链接")

        global CURRENT_EXTENSIONS
        CURRENT_EXTENSIONS = exts

        if url_ok:
            self.status_var.set("请选择文件，或点击「粘贴链接」输入 URL")
        else:
            self.status_var.set("请选择文件")

    def _browse_file(self):
        """浏览选择文件（不限格式）"""
        path = filedialog.askopenfilename(
            title="选择要转换的文件",
            filetypes=[("所有文件", "*.*")],
        )
        if path:
            self.path_var.set(path)
            self.status_var.set(f"已选择：{os.path.basename(path)}")

    def _paste_url(self):
        """弹窗输入 URL"""
        category = self.category_var.get()
        if not accepts_url(category):
            messagebox.showinfo("提示", "当前类别不支持 URL 输入，请选择「在线链接」类别。")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("输入链接")
        dialog.geometry("500x120")
        dialog.resizable(True, False)
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="请输入 URL：").pack(anchor=tk.W, padx=12, pady=(12, 4))

        url_var = tk.StringVar()
        url_entry = ttk.Entry(dialog, textvariable=url_var, width=60)
        url_entry.pack(fill=tk.X, padx=12, pady=(0, 12))
        url_entry.focus_set()

        def confirm():
            if url_var.get().strip():
                self.path_var.set(url_var.get().strip())
                self.status_var.set(f"已输入链接：{url_var.get().strip()[:60]}")
                dialog.destroy()

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=12, pady=(0, 12))
        ttk.Button(btn_frame, text="确定", command=confirm).pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.RIGHT, padx=(0, 6))

        url_entry.bind("<Return>", lambda e: confirm())

    def _start_convert(self):
        """开始转换（在后台线程中执行）"""
        path = self.path_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择文件或输入链接。")
            return

        category = self.category_var.get()

        # 禁用按钮，显示进度
        self.convert_btn.config(state=tk.DISABLED)
        self.browse_btn.config(state=tk.DISABLED)
        self.url_btn.config(state=tk.DISABLED)
        self.progress.pack(fill=tk.X, padx=12, pady=(0, 4))
        self.progress.start(12)
        self.status_var.set("正在转换中，请稍候...")
        self.text.delete("1.0", tk.END)
        self.save_btn.config(state=tk.DISABLED)
        self.copy_btn.config(state=tk.DISABLED)

        def _do_convert():
            if path.startswith(("http://", "https://")):
                success, result = convert_url(path)
            else:
                success, result = convert_file(path, category)

            # 在主线程更新 UI
            self.root.after(0, lambda: self._convert_done(success, result))

        threading.Thread(target=_do_convert, daemon=True).start()

    def _convert_done(self, success: bool, result: str):
        """转换完成回调（在主线程中）"""
        self.progress.stop()
        self.progress.pack_forget()
        self.convert_btn.config(state=tk.NORMAL)
        self.browse_btn.config(state=tk.NORMAL)
        self.url_btn.config(state=tk.NORMAL)

        if success:
            self.text.delete("1.0", tk.END)
            self.text.insert("1.0", result)
            self.save_btn.config(state=tk.NORMAL)
            self.copy_btn.config(state=tk.NORMAL)

            # 统计字符数
            char_count = len(result)
            line_count = result.count("\n")
            self.status_var.set(
                f"转换成功！共 {char_count} 字符 / {line_count} 行"
            )
        else:
            self.text.delete("1.0", tk.END)
            self.text.insert("1.0", f"❌ {result}\n")
            self.status_var.set("转换失败")

    def _clear_result(self):
        """清空结果"""
        self.text.delete("1.0", tk.END)
        self.save_btn.config(state=tk.DISABLED)
        self.copy_btn.config(state=tk.DISABLED)
        self.status_var.set("就绪")

    def _save_result(self):
        """保存结果为 .md 文件"""
        content = self.text.get("1.0", tk.END).strip()
        if not content:
            return

        # 尝试从内容提取标题作为文件名
        first_line = content.split("\n")[0].lstrip("# ").strip()
        default_name = first_line if first_line else "converted"
        # 清理非法字符
        for ch in r'\/:*?"<>|':
            default_name = default_name.replace(ch, "")
        if not default_name:
            default_name = "converted"
        default_name += ".md"

        path = filedialog.asksaveasfilename(
            title="保存 Markdown 文件",
            defaultextension=".md",
            initialfile=default_name,
            filetypes=[
                ("Markdown 文件", "*.md"),
                ("纯文本文件", "*.txt"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
                self.status_var.set(f"已保存至：{os.path.basename(path)}")
                messagebox.showinfo("保存成功", f"文件已保存至：\n{path}")
            except Exception as e:
                messagebox.showerror("保存失败", str(e))

    def _copy_result(self):
        """复制结果到剪贴板"""
        content = self.text.get("1.0", tk.END).strip()
        if not content:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.status_var.set("已复制到剪贴板")


# ─── 主入口 ───────────────────────────────────────────────────────────────────
def main():
    root = tk.Tk()

    # 设置图标（如果有的话）
    try:
        # Windows 下设置任务栏图标
        if sys.platform == "win32":
            root.iconbitmap(default="")
    except Exception:
        pass

    app = MarkITDownGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
