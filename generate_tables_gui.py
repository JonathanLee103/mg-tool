"""MG 订单表 & 库存表 生成工具 — GUI 版"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

from generate_tables import generate, detect_files, FILES_EXPECTED


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("MG 订单表 & 库存表 生成工具")
        self.geometry("680x620")
        self.minsize(600, 520)
        self.resizable(True, True)

        # DPI awareness for Windows
        if sys.platform == "win32":
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass

        self._work_dir = tk.StringVar()
        self._file_labels = {}
        self._running = False

        self._build_ui()
        self._refresh_file_status()

    # ── UI 构建 ──────────────────────────────────────────────

    def _build_ui(self):
        # 主容器
        main = ttk.Frame(self, padding="16 12 16 12")
        main.pack(fill="both", expand=True)

        # ── 标题 ──
        ttk.Label(
            main, text="MG 订单表 & 库存表 生成工具",
            font=("Microsoft YaHei", 16, "bold") if sys.platform == "win32"
            else ("PingFang SC", 16, "bold"),
        ).pack(anchor="center", pady=(0, 12))

        # ── 工作目录选择 ──
        dir_frame = ttk.LabelFrame(main, text="工作目录", padding="8 6 8 6")
        dir_frame.pack(fill="x", pady=(0, 8))

        dir_row = ttk.Frame(dir_frame)
        dir_row.pack(fill="x")
        ttk.Entry(dir_row, textvariable=self._work_dir).pack(
            side="left", fill="x", expand=True, padx=(0, 6)
        )
        ttk.Button(dir_row, text="浏览...", command=self._on_select_dir, width=10).pack(
            side="right"
        )

        # ── 源文件检测 ──
        files_frame = ttk.LabelFrame(main, text="源文件检测", padding="8 6 8 6")
        files_frame.pack(fill="x", pady=(0, 8))

        # 期待文件名 → 显示名
        self._file_display = {
            "系统导出_采购订单数据.xlsx": "采购订单数据",
            "系统导出_订单汇总.xls": "订单汇总",
            "系统导出_业务库存数据.xlsx": "业务库存数据",
            "系统导出_已发货库存明细.xlsx": "已发货库存明细",
        }

        for fname, display in self._file_display.items():
            row = ttk.Frame(files_frame)
            row.pack(fill="x", pady=1)
            ttk.Label(row, text=f"{display}:", width=18, anchor="e").pack(side="left")
            lbl = ttk.Label(row, text="等待选择目录", foreground="gray")
            lbl.pack(side="left", padx=(6, 0))
            self._file_labels[fname] = lbl

        # ── 生成按钮 ──
        self._btn_generate = ttk.Button(
            main, text="开始生成", command=self._on_generate, state="disabled"
        )
        self._btn_generate.pack(pady=(6, 8))
        if sys.platform == "win32":
            self._btn_generate.configure(width=24)

        # ── 进度条 + 状态 ──
        self._progress = ttk.Progressbar(main, mode="determinate", maximum=100)
        self._progress.pack(fill="x", pady=(0, 2))

        self._status_label = ttk.Label(main, text="就绪", foreground="gray")
        self._status_label.pack(anchor="center")

        # ── 日志区域 ──
        log_frame = ttk.LabelFrame(main, text="日志", padding="4 4 4 4")
        log_frame.pack(fill="both", expand=True, pady=(8, 6))

        self._log_text = tk.Text(
            log_frame,
            height=12,
            wrap="word",
            font=("Consolas", 10) if sys.platform == "win32" else ("Menlo", 10),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="white",
            relief="flat",
            borderwidth=0,
        )
        self._log_text.pack(fill="both", expand=True)

        # log scrollbar
        sb = ttk.Scrollbar(log_frame, command=self._log_text.yview)
        sb.pack(side="right", fill="y")
        self._log_text.configure(yscrollcommand=sb.set)

        # ── 底部操作栏 ──
        bottom = ttk.Frame(main)
        bottom.pack(fill="x", pady=(4, 0))
        ttk.Button(
            bottom, text="打开输出文件夹", command=self._on_open_folder
        ).pack(side="left")
        ttk.Button(bottom, text="清空日志", command=self._on_clear_log).pack(
            side="right"
        )

    # ── 事件处理 ──────────────────────────────────────────────

    def _on_select_dir(self):
        path = filedialog.askdirectory(title="选择工作目录（包含 4 个源文件的目录）")
        if not path:
            return
        self._work_dir.set(path)
        self._refresh_file_status()

    def _refresh_file_status(self):
        work_dir = self._work_dir.get()
        if not work_dir:
            for lbl in self._file_labels.values():
                lbl.configure(text="等待选择目录", foreground="gray")
            self._btn_generate.configure(state="disabled")
            return

        status = detect_files(work_dir)
        all_found = True
        for fname, lbl in self._file_labels.items():
            if status.get(fname, False):
                lbl.configure(text="✓ 已找到", foreground="green")
            else:
                lbl.configure(text="✗ 未找到", foreground="red")
                all_found = False

        self._btn_generate.configure(state="normal" if all_found else "disabled")
        if all_found:
            self._status_label.configure(text="全部源文件已就绪，可以生成", foreground="green")
        else:
            self._status_label.configure(text="有文件缺失，请检查工作目录", foreground="red")

    def _on_generate(self):
        if self._running:
            return
        self._running = True
        self._btn_generate.configure(state="disabled", text="生成中...")
        self._progress["value"] = 0
        self._status_label.configure(text="正在执行...", foreground="black")
        self._log_text.delete("1.0", "end")

        work_dir = self._work_dir.get()

        def run():
            try:
                result = generate(
                    work_dir=work_dir,
                    log_cb=self._log,
                    progress_cb=self._progress_cb,
                )
                self._on_done(result)
            except Exception as e:
                self._on_error(str(e))

        threading.Thread(target=run, daemon=True).start()

    def _on_done(self, result):
        self._running = False
        self._btn_generate.configure(state="normal", text="开始生成")
        self._progress["value"] = 100
        self._status_label.configure(
            text=f"完成! 订单表 {result['order_rows']} 行, 库存表 {result['inventory_rows']} 行",
            foreground="green",
        )
        self.after(100, lambda: messagebox.showinfo(
            "生成完成",
            f"订单表 {result['order_rows']} 行\n"
            f"库存表 {result['inventory_rows']} 行\n\n"
            f"输出文件:\n{result['output_path']}",
        ))

    def _on_error(self, error_msg):
        self._running = False
        self._btn_generate.configure(state="normal", text="开始生成")
        self._status_label.configure(text="生成失败", foreground="red")
        self._log(f"\n[错误] {error_msg}")
        self.after(100, lambda: messagebox.showerror("错误", f"生成过程中出现错误:\n\n{error_msg}"))

    def _on_open_folder(self):
        work_dir = self._work_dir.get()
        target = work_dir if work_dir and os.path.isdir(work_dir) else os.getcwd()
        if sys.platform == "win32":
            os.startfile(target)
        elif sys.platform == "darwin":
            os.system(f'open "{target}"')
        else:
            os.system(f'xdg-open "{target}"')

    def _on_clear_log(self):
        self._log_text.delete("1.0", "end")

    # ── 回调（从工作线程调用，线程安全） ────────────────────────

    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        # tkinter 不是线程安全的，需要用 after 调度到主线程
        self.after(0, lambda: self._append_log(f"[{ts}] {msg}\n"))

    def _append_log(self, text):
        self._log_text.insert("end", text)
        self._log_text.see("end")

    def _progress_cb(self, pct, status):
        self.after(0, lambda: self._update_progress(pct, status))

    def _update_progress(self, pct, status):
        self._progress["value"] = pct
        self._status_label.configure(text=status)


if __name__ == "__main__":
    app = App()
    app.mainloop()
