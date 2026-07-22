import sys, os
_lib = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_lib')
if os.path.isdir(_lib) and _lib not in sys.path: sys.path.insert(0, _lib)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import subprocess, sys, os, threading, json
from pathlib import Path


def resource_path():
    """获取脚本所在目录"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


class ExcelMatcherGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Matcher")
        self.root.geometry("800x720")
        self.root.minsize(700, 650)

        # 用 ttk 主题
        style = ttk.Style()
        style.theme_use("vista" if "vista" in style.theme_names() else "clam")

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10,5))

        # ── Tab 1: 主界面 ──
        self.tab_main = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_main, text=" 匹配设置 ")

        self._build_file_frame()
        self._build_mapping_frame()
        self._build_options_frame()
        self._build_action_frame()
        self._build_log_frame()

        # ── Tab 2: 使用说明 ──
        self.tab_help = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_help, text=" 使用说明 ")
        self._build_help_tab()

        # 状态变量
        self.proc = None
        self.running = False

    # ════════════════════════════════════════════════
    #  文件选择区域
    # ════════════════════════════════════════════════
    def _build_file_frame(self):
        f = ttk.LabelFrame(self.tab_main, text=" 文件选择 ", padding=8)
        f.pack(fill=tk.X, pady=(0,8))

        # 源表
        r1 = ttk.Frame(f)
        r1.pack(fill=tk.X, pady=2)
        ttk.Label(r1, width=7, text="源表:").pack(side=tk.LEFT)
        self.src_path = tk.StringVar()
        ttk.Entry(r1, textvariable=self.src_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,4))
        ttk.Button(r1, text="浏览...", command=lambda: self._browse_file(self.src_path, self.src_sheet_cb)).pack(side=tk.RIGHT)
        ttk.Label(r1, text="Sheet:").pack(side=tk.LEFT, padx=(10,2))
        self.src_sheet_cb = ttk.Combobox(r1, width=12, state="readonly")
        self.src_sheet_cb.pack(side=tk.LEFT)
        ttk.Label(r1, text="表头行:").pack(side=tk.LEFT, padx=(10,2))
        self.src_header = tk.IntVar(value=0)
        ttk.Spinbox(r1, from_=0, to=99, textvariable=self.src_header, width=4).pack(side=tk.LEFT)

        # 目标表
        r2 = ttk.Frame(f)
        r2.pack(fill=tk.X, pady=2)
        ttk.Label(r2, width=7, text="目标表:").pack(side=tk.LEFT)
        self.tgt_path = tk.StringVar()
        ttk.Entry(r2, textvariable=self.tgt_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,4))
        ttk.Button(r2, text="浏览...", command=lambda: self._browse_file(self.tgt_path, self.tgt_sheet_cb)).pack(side=tk.RIGHT)
        ttk.Label(r2, text="Sheet:").pack(side=tk.LEFT, padx=(10,2))
        self.tgt_sheet_cb = ttk.Combobox(r2, width=12, state="readonly")
        self.tgt_sheet_cb.pack(side=tk.LEFT)
        ttk.Label(r2, text="表头行:").pack(side=tk.LEFT, padx=(10,2))
        self.tgt_header = tk.IntVar(value=0)
        ttk.Spinbox(r2, from_=0, to=99, textvariable=self.tgt_header, width=4).pack(side=tk.LEFT)

        # 输出
        r3 = ttk.Frame(f)
        r3.pack(fill=tk.X, pady=2)
        ttk.Label(r3, width=7, text="输出:").pack(side=tk.LEFT)
        self.out_path = tk.StringVar()
        ttk.Entry(r3, textvariable=self.out_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,4))
        ttk.Button(r3, text="浏览...", command=self._browse_output).pack(side=tk.RIGHT)

    def _browse_file(self, var, cb):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if not path:
            return
        var.set(path)
        self._detect_sheets(path, cb)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="保存输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if path:
            self.out_path.set(path)

    def _detect_sheets(self, path, cb):
        """读取文件中的 sheet 名称"""
        cb["values"] = []
        try:
            import pandas as pd
            xls = pd.ExcelFile(path)
            sheets = xls.sheet_names
            cb["values"] = sheets
            if sheets:
                cb.set(sheets[0])
            cb.configure(state="readonly")
        except Exception as e:
            pass

    # ════════════════════════════════════════════════
    #  匹配映射区域
    # ════════════════════════════════════════════════
    def _build_mapping_frame(self):
        f = ttk.LabelFrame(self.tab_main, text=" 匹配设置 ", padding=8)
        f.pack(fill=tk.BOTH, expand=True, pady=(0,8))

        # 双列布局
        panes = ttk.PanedWindow(f, orient=tk.HORIZONTAL)
        panes.pack(fill=tk.BOTH, expand=True)

        # 左：索引列
        left = ttk.Frame(panes)
        panes.add(left, weight=1)
        tk.Label(left, text="索引列（用于匹配）", font=("", 10, "bold"),
                 fg="#1a5276").pack(anchor=tk.W)
        tk.Label(left, text='格式：COL 或 源:目标', fg="gray",
                 font=("", 8)).pack(anchor=tk.W)
        self.key_frame = ttk.Frame(left)
        self.key_frame.pack(fill=tk.BOTH, expand=True, pady=(4,0))
        self.key_rows = []
        self._add_key_row()

        btn_k = ttk.Button(left, text="+ 添加索引列", command=self._add_key_row)
        btn_k.pack(anchor=tk.W, pady=(4,0))

        # 右：填充列
        right = ttk.Frame(panes)
        panes.add(right, weight=1)
        tk.Label(right, text="填充列（要复制的内容）", font=("", 10, "bold"),
                 fg="#1a5276").pack(anchor=tk.W)
        tk.Label(right, text='格式：COL 或 源:目标', fg="gray",
                 font=("", 8)).pack(anchor=tk.W)
        self.fill_frame = ttk.Frame(right)
        self.fill_frame.pack(fill=tk.BOTH, expand=True, pady=(4,0))
        self.fill_rows = []
        self._add_fill_row()

        btn_f = ttk.Button(right, text="+ 添加填充列", command=self._add_fill_row)
        btn_f.pack(anchor=tk.W, pady=(4,0))

    def _add_key_row(self):
        self._add_col_row(self.key_frame, self.key_rows)

    def _add_fill_row(self):
        self._add_col_row(self.fill_frame, self.fill_rows)

    def _add_col_row(self, parent_frame, rows_list):
        idx = len(rows_list)
        f = ttk.Frame(parent_frame)
        f.pack(fill=tk.X, pady=1)
        var = tk.StringVar()
        if idx == 0:
            # 默认示例值
            if rows_list is self.key_rows:
                var.set("姓名")
            elif rows_list is self.fill_rows:
                var.set("学号")
        e = ttk.Entry(f, textvariable=var)
        e.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,4))
        btn = ttk.Button(f, text="✕", width=3, command=lambda: self._remove_col_row(f, rows_list))
        btn.pack(side=tk.RIGHT)
        rows_list.append((f, var))

    def _remove_col_row(self, frame, rows_list):
        for item in rows_list:
            if item[0] == frame:
                rows_list.remove(item)
                break
        frame.destroy()

    def _get_col_values(self, rows_list):
        """获取列值列表"""
        vals = []
        for _, var in rows_list:
            v = var.get().strip()
            if v:
                vals.append(v)
        return vals

    # ════════════════════════════════════════════════
    #  高级选项
    # ════════════════════════════════════════════════
    def _build_options_frame(self):
        f = ttk.LabelFrame(self.tab_main, text=" 高级选项 ", padding=8)
        f.pack(fill=tk.X, pady=(0,8))

        r1 = ttk.Frame(f)
        r1.pack(fill=tk.X)
        self.fuzzy_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(r1, text="启用模糊匹配", variable=self.fuzzy_var,
                        command=self._toggle_fuzzy).pack(side=tk.LEFT)

        ttk.Label(r1, text="相似度阈值:").pack(side=tk.LEFT, padx=(15,4))
        self.threshold_var = tk.DoubleVar(value=0.85)
        self.threshold_slider = ttk.Scale(
            r1, from_=0.5, to=1.0, variable=self.threshold_var,
            orient=tk.HORIZONTAL, length=150
        )
        self.threshold_slider.pack(side=tk.LEFT)
        self.threshold_label = ttk.Label(r1, width=5, text="0.85")
        self.threshold_label.pack(side=tk.LEFT, padx=(4,0))

        def update_threshold_label(*_):
            self.threshold_label.config(text=f"{self.threshold_var.get():.2f}")
        self.threshold_var.trace_add("write", update_threshold_label)

        self.backup_var = tk.BooleanVar(value=True)
        self.append_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(r1, text="自动备份目标文件", variable=self.backup_var).pack(side=tk.LEFT, padx=(20,0))
        ttk.Checkbutton(r1, text="追加缺失记录", variable=self.append_var).pack(side=tk.LEFT, padx=(10,0))

    def _toggle_fuzzy(self):
        state = tk.NORMAL if self.fuzzy_var.get() else tk.DISABLED
        self.threshold_slider.configure(state=state)
        self.threshold_label.configure(state=state)

    # ════════════════════════════════════════════════
    #  操作按钮
    # ════════════════════════════════════════════════
    def _build_action_frame(self):
        f = ttk.Frame(self.tab_main)
        f.pack(fill=tk.X, pady=(0,8))
        self.run_btn = ttk.Button(f, text="▶  开始匹配", command=self._run_match)
        self.run_btn.pack(side=tk.LEFT)
        ttk.Button(f, text="清空日志", command=lambda: self.log_text.delete("1.0", tk.END)).pack(side=tk.RIGHT)

    # ════════════════════════════════════════════════
    #  日志输出
    # ════════════════════════════════════════════════
    def _build_log_frame(self):
        f = ttk.LabelFrame(self.tab_main, text=" 运行日志 ", padding=4)
        f.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(f, height=8, wrap=tk.WORD,
                                                    font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    # ════════════════════════════════════════════════
    #  使用说明
    # ════════════════════════════════════════════════
    def _build_help_tab(self):
        text = scrolledtext.ScrolledText(self.tab_help, wrap=tk.WORD,
                                          font=("Microsoft YaHei", 10), padx=10, pady=10)
        text.pack(fill=tk.BOTH, expand=True)
        help_content = """Excel 数据匹配填充工具 - 使用说明
════════════════════════════════════════════════════════════


一、基本用法
────────────────────────────────────────
  1. 在"源表"中选择有完整数据的 Excel 文件
  2. 在"目标表"中选择待填充的 Excel 文件
  3. 设置索引列（用于匹配的列名）
  4. 设置填充列（要从源表复制到目标表的内容）
  5. 点击"开始匹配"


二、索引列格式
────────────────────────────────────────
  直接输入列名: 姓名
    源表和目标表都用相同的列名"姓名"进行匹配

  列名不同时用冒号分隔: Name:姓名
    源表列名叫"Name"，目标表列名叫"姓名"


三、填充列格式
────────────────────────────────────────
  和索引列格式相同。
  如填充列在目标表中不存在，会自动创建新列。


四、表头行设置
────────────────────────────────────────
  默认表头在第一行（行号 0）。
  如果第一行是标题（如"2001级"），表头实际在第二行，
  请将"表头行"设为 1。


五、模糊匹配
────────────────────────────────────────
  处理姓名有细微差异的情况（如"候瑞"vs"侯瑞"）。
  仅适用于单列索引。
  阈值越接近 1.0，匹配越严格。
"""
        text.insert("1.0", help_content)
        text.configure(state=tk.DISABLED)

    # ════════════════════════════════════════════════
    #  执行匹配
    # ════════════════════════════════════════════════
    def _run_match(self):
        if self.running:
            return

        # 验证输入
        src = self.src_path.get()
        tgt = self.tgt_path.get()
        if not src or not Path(src).exists():
            messagebox.showerror("错误", "请选择有效的源文件")
            return
        if not tgt or not Path(tgt).exists():
            messagebox.showerror("错误", "请选择有效的目标文件")
            return

        keys = self._get_col_values(self.key_rows)
        fills = self._get_col_values(self.fill_rows)
        if not keys:
            messagebox.showerror("错误", "请至少添加一个索引列")
            return
        if not fills:
            messagebox.showerror("错误", "请至少添加一个填充列")
            return

        out = self.out_path.get().strip()
        if not out:
            out = str(Path(tgt).parent / f"已填充_{Path(tgt).name}")
            self.out_path.set(out)

        # 构造命令行参数
        script = resource_path() / "match_excel.py"
        if not script.exists():
            messagebox.showerror("错误", f"找不到核心脚本: {script}")
            return

        cmd = [sys.executable, str(script), "-s", src, "-t", tgt, "-o", out]
        for k in keys:
            cmd += ["-k", k]
        for f_val in fills:
            cmd += ["-f", f_val]

        if self.fuzzy_var.get():
            cmd += ["--fuzzy", "--threshold", f"{self.threshold_var.get():.2f}"]

        if self.src_sheet_cb.get():
            cmd += ["--source-sheet", self.src_sheet_cb.get()]
        if self.tgt_sheet_cb.get():
            cmd += ["--target-sheet", self.tgt_sheet_cb.get()]

        cmd += ["--source-header-row", str(self.src_header.get())]
        cmd += ["--target-header-row", str(self.tgt_header.get())]

        if not self.backup_var.get():
            cmd += ["--no-backup"]
        if self.append_var.get():
            cmd += ["--append-missing"]

        self.log(f"> 启动匹配: {' '.join(cmd)}")
        self.log(f"  索引: {', '.join(keys)}")
        self.log(f"  填充: {', '.join(fills)}")
        self.log("─" * 50)

        self.running = True
        self.run_btn.configure(state=tk.DISABLED, text="⏳ 运行中...")

        thread = threading.Thread(target=self._run_process, args=(cmd,), daemon=True)
        thread.start()

    def _run_process(self, cmd):
        try:
            self.proc = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace",
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            )
            for line in self.proc.stdout:
                self.root.after(0, self.log, line.rstrip())
            self.proc.wait()
            self.root.after(0, self._on_finish, self.proc.returncode)
        except Exception as e:
            self.root.after(0, self.log, f"[错误] {e}")
            self.root.after(0, self._on_finish, -1)

    def _on_finish(self, code):
        self.running = False
        self.run_btn.configure(state=tk.NORMAL, text="▶  开始匹配")
        if code == 0:
            self.log("=" * 50)
            self.log("✅ 匹配完成！")
        else:
            self.log("=" * 50)
            self.log("❌ 匹配出错，请检查日志")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMatcherGUI(root)
    root.mainloop()



