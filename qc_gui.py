#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 侧弯支具出厂质检表 自动生成工具（桌面图形版）
================================================================================
现代化桌面界面，功能：
  读 Excel → 生成质检表 → 重命名 → 归档 → 打印（可选）

所有路径（工作跟进表 / 质检表模板 / 归档目录）都可以在界面上手动选择，
选择结果自动保存到个人目录配置文件中（~/.qc_tool_config.json），
下次打开自动载入 —— 换人换电脑都能用，无需改代码。

用法：
  python qc_gui.py          直接启动图形界面
  打包：pyinstaller --onefile --windowed --name 质检表工具 --collect-all customtkinter qc_gui.py
       （--windowed 使双击直接显示界面，不弹黑色命令行窗口）
依赖：customtkinter, python-docx, openpyxl
================================================================================
"""
import contextlib
import io
import json
import os
import queue
import re
import sys
import threading
from tkinter import filedialog

import customtkinter as ctk

import qc_pipeline as qc  # 复用命令行版的全部核心逻辑

# 主题外观
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

COLOR_BG = "#0f172a"        # 主背景 深蓝黑
COLOR_CARD = "#1e293b"      # 卡片背景
COLOR_BORDER = "#334155"    # 卡片边框
COLOR_TEXT = "#e2e8f0"      # 主文字
COLOR_SUB = "#94a3b8"       # 次要文字
COLOR_ACCENT = "#2563eb"    # 主色 蓝
COLOR_OK = "#22c55e"        # 成功绿
COLOR_ERR = "#ef4444"       # 错误红

CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".qc_tool_config.json")


class QueueWriter(io.StringIO):
    """把 print 输出重定向到队列，供 UI 线程轮询显示。"""

    def __init__(self, q):
        super().__init__()
        self._q = q

    def write(self, s):
        if s:
            self._q.put(s)
        return super().write(s)


class OrderSelectDialog(ctk.CTkToplevel):
    """同名患者有多条订单时的人工选择弹窗。选择结果写入 result dict 并关闭。"""

    def __init__(self, master, name, candidates, result):
        super().__init__(master)
        self.title("选择订单 - %s" % name)
        self.geometry("720x460")
        self.minsize(640, 360)
        self.configure(fg_color=COLOR_BG)
        self.transient(master)
        self.grab_set()

        self.candidates = candidates
        self.result = result

        ctk.CTkLabel(
            self,
            text="患者「%s」在 Excel 中有 %d 条订单，请选择要生成质检表的那一条："
                 % (name, len(candidates)),
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=COLOR_TEXT, wraplength=650).pack(
            anchor="w", padx=20, pady=(18, 10))

        scroll = ctk.CTkScrollableFrame(
            self, fg_color=COLOR_CARD, corner_radius=10,
            border_width=1, border_color=COLOR_BORDER)
        scroll.pack(fill="both", expand=True, padx=20, pady=6)

        for i, o in enumerate(candidates, 1):
            label = "%d) 订单 %s | 客户 %s | 批次 %s | 设备 %s | 理论 %sg | 实际 %sg%s" % (
                i, o["order_no"], o["customer"], o["batch"], o["machine"],
                o["theory"], o["actual"], " | 售后件" if o["is_after_sale"] else "")
            btn = ctk.CTkButton(
                scroll, text=label, height=40, corner_radius=8,
                anchor="w", font=ctk.CTkFont(size=13),
                fg_color="#0b1220", hover_color="#1e3a8a",
                text_color=COLOR_TEXT,
                command=lambda idx=i: self._pick(idx))
            btn.pack(fill="x", padx=10, pady=3)

        ctk.CTkButton(
            self, text="取消", height=36, corner_radius=8,
            fg_color="#334155", hover_color="#475569",
            font=ctk.CTkFont(size=13), command=self.destroy).pack(
            anchor="e", padx=20, pady=(8, 16))

    def _pick(self, idx):
        self.result["chosen"] = self.candidates[idx - 1]
        self.destroy()


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("侧弯支具出厂质检表 · 自动生成工具")
        self.geometry("980x800")
        self.minsize(860, 700)
        self.configure(fg_color=COLOR_BG)

        self.log_q = queue.Queue()
        self.running = False

        # 路径配置：优先读个人配置文件，没有则用本机默认值
        self.config = self._load_config()
        self.var_excel = ctk.StringVar(
            value=self.config.get("excel", qc.DEFAULT_EXCEL))
        self.var_template = ctk.StringVar(
            value=self.config.get("template", qc.DEFAULT_TEMPLATE))
        self.var_base_dir = ctk.StringVar(
            value=self.config.get("base_dir", qc.DEFAULT_BASE_DIR))
        self.var_copies = ctk.StringVar(
            value=str(self.config.get("copies", qc.DEFAULT_COPIES)))
        self.var_printer = ctk.StringVar(
            value=self.config.get("printer", qc.DEFAULT_PRINTER))

        self._build_header()
        self._build_path_card()
        self._build_config_card()
        self._build_log_card()
        self._build_statusbar()
        self._load_printers()

        # 定时轮询日志队列
        self.after(100, self._poll_logs)

    # ------------------------------------------------------------------ 配置读写
    def _load_config(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as fp:
                    return json.load(fp)
        except Exception:
            pass
        return {}

    def _save_config(self):
        try:
            cfg = {
                "excel": self.var_excel.get(),
                "template": self.var_template.get(),
                "base_dir": self.var_base_dir.get(),
                "printer": self.var_printer.get(),
                "copies": int(self.var_copies.get() or 2),
            }
            with open(CONFIG_FILE, "w", encoding="utf-8") as fp:
                json.dump(cfg, fp, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ------------------------------------------------------------------ UI 构建
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(20, 8))
        ctk.CTkLabel(
            header, text="3D打印脊柱侧弯矫正器 出厂质检报告生成器",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=COLOR_TEXT).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="读取工作跟进表 → 自动填写质检表 → 归档到订单文件夹 → 打印（可跳过）",
            font=ctk.CTkFont(size=13), text_color=COLOR_SUB).pack(
            anchor="w", pady=(2, 0))

    def _build_path_card(self):
        """数据源卡片：Excel / 模板 / 归档目录，均可手动选择。"""
        card = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=14,
                            border_width=1, border_color=COLOR_BORDER)
        card.pack(fill="x", padx=24, pady=10)

        ctk.CTkLabel(card, text="数据源（可手动选择，选择后自动记住）", anchor="w",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=COLOR_SUB).pack(fill="x", padx=18, pady=(14, 6))

        self._path_row(card, "工作跟进表", self.var_excel, "file",
                       [("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")])
        self._path_row(card, "质检表模板", self.var_template, "file",
                       [("Word 文件", "*.docx"), ("所有文件", "*.*")])
        self._path_row(card, "归档目录", self.var_base_dir, "dir",
                       None)
        # 占位底部间距
        ctk.CTkFrame(card, fg_color="transparent", height=6).pack(fill="x")

    def _path_row(self, parent, label, var, kind, filetypes):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=4)
        ctk.CTkLabel(row, text=label, width=90, anchor="w",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=COLOR_TEXT).pack(side="left")
        entry = ctk.CTkEntry(row, textvariable=var,
                             font=ctk.CTkFont(size=12), height=34,
                             fg_color="#0b1220", border_color=COLOR_BORDER)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        btn = ctk.CTkButton(row, text="选择...", width=76, height=32,
                            corner_radius=8, fg_color="#334155",
                            hover_color="#475569",
                            font=ctk.CTkFont(size=12),
                            command=lambda: self._pick_path(var, kind, filetypes))
        btn.pack(side="left")

    def _pick_path(self, var, kind, filetypes):
        if kind == "file":
            p = filedialog.askopenfilename(filetypes=filetypes,
                                           initialdir=os.path.dirname(var.get())
                                           if os.path.exists(os.path.dirname(var.get()))
                                           else os.path.expanduser("~"))
        else:
            p = filedialog.askdirectory(initialdir=var.get()
                                        if os.path.isdir(var.get())
                                        else os.path.expanduser("~"))
        if p:
            var.set(p)
            self._save_config()
            self._append_log("已选择：%s" % p, "info")

    def _build_config_card(self):
        card = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=14,
                            border_width=1, border_color=COLOR_BORDER)
        card.pack(fill="x", padx=24, pady=10)

        # 行1：模型名称
        row1 = ctk.CTkFrame(card, fg_color="transparent")
        row1.pack(fill="x", padx=18, pady=(16, 10))
        ctk.CTkLabel(row1, text="模型名称", width=90, anchor="w",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=COLOR_TEXT).pack(side="left")
        self.entry_names = ctk.CTkEntry(
            row1,
            placeholder_text="多个名称用空格 / 逗号分隔，如：聂梓瑜 马睿杰 方馨悦 张梓菡",
            font=ctk.CTkFont(size=13), height=38,
            fg_color="#0b1220", border_color=COLOR_BORDER)
        self.entry_names.pack(side="left", fill="x", expand=True)

        # 行2：打印机 + 份数
        row2 = ctk.CTkFrame(card, fg_color="transparent")
        row2.pack(fill="x", padx=18, pady=6)
        ctk.CTkLabel(row2, text="打印机", width=90, anchor="w",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=COLOR_TEXT).pack(side="left")
        self.menu_printer = ctk.CTkOptionMenu(
            row2, variable=self.var_printer, values=["正在加载..."],
            width=360, height=36, fg_color="#0b1220",
            button_color=COLOR_ACCENT, button_hover_color="#1d4ed8",
            text_color=COLOR_TEXT, command=lambda _: self._save_config())
        self.menu_printer.pack(side="left")
        ctk.CTkLabel(row2, text="每份张数", width=80, anchor="e",
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color=COLOR_TEXT).pack(side="left", padx=(16, 0))
        self.entry_copies = ctk.CTkEntry(
            row2, textvariable=self.var_copies, width=70, height=36,
            fg_color="#0b1220", border_color=COLOR_BORDER)
        self.entry_copies.pack(side="left")

        # 行3：打印开关 + 执行按钮
        row3 = ctk.CTkFrame(card, fg_color="transparent")
        row3.pack(fill="x", padx=18, pady=(10, 16))
        self.switch_print = ctk.CTkSwitch(
            row3, text="打印", font=ctk.CTkFont(size=14),
            progress_color=COLOR_ACCENT, text_color=COLOR_TEXT)
        self.switch_print.select()
        self.switch_print.pack(side="left")
        self.btn_run = ctk.CTkButton(
            row3, text="开始执行", height=40, corner_radius=10,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color=COLOR_ACCENT, hover_color="#1d4ed8",
            command=self._on_run)
        self.btn_run.pack(side="right")

    def _build_log_card(self):
        card = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=14,
                            border_width=1, border_color=COLOR_BORDER)
        card.pack(fill="both", expand=True, padx=24, pady=10)
        head = ctk.CTkFrame(card, fg_color="transparent")
        head.pack(fill="x", padx=14, pady=(12, 4))
        ctk.CTkLabel(head, text="运行日志", anchor="w",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=COLOR_SUB).pack(side="left")
        self.lbl_status = ctk.CTkLabel(head, text="● 就绪", anchor="e",
                                       font=ctk.CTkFont(size=13, weight="bold"),
                                       text_color=COLOR_SUB)
        self.lbl_status.pack(side="right")

        self.txt_log = ctk.CTkTextbox(
            card, font=ctk.CTkFont(family="Consolas", size=13),
            fg_color="#0b1220", corner_radius=10, wrap="word",
            text_color="#cbd5e1", border_width=1, border_color=COLOR_BORDER)
        self.txt_log.pack(fill="both", expand=True, padx=14, pady=(4, 14))
        self.txt_log.configure(state="disabled")
        self.txt_log.tag_config("ok", foreground=COLOR_OK)
        self.txt_log.tag_config("err", foreground=COLOR_ERR)
        self.txt_log.tag_config("info", foreground=COLOR_TEXT)

    def _build_statusbar(self):
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", padx=24, pady=(2, 14))
        self.lbl_result = ctk.CTkLabel(
            bar, text="", anchor="w", font=ctk.CTkFont(size=12),
            text_color=COLOR_SUB)
        self.lbl_result.pack(side="left")

    # ------------------------------------------------------------------ 逻辑
    def _load_printers(self):
        def work():
            try:
                printers = qc.list_printers()
                self.after(0, lambda: self._set_printers(printers))
            except Exception:
                self.after(0, lambda: self._set_printers([]))

        threading.Thread(target=work, daemon=True).start()

    def _set_printers(self, printers):
        if printers:
            cur = self.var_printer.get()
            if cur not in printers:
                cur = qc.DEFAULT_PRINTER if qc.DEFAULT_PRINTER in printers \
                    else printers[0]
                self.var_printer.set(cur)
            self.menu_printer.configure(values=printers)
            self.menu_printer.set(cur)
        else:
            self.menu_printer.configure(values=["（未找到打印机）"])
            self.menu_printer.set("（未找到打印机）")

    def _on_run(self):
        if self.running:
            return
        names_text = self.entry_names.get().strip()
        names = [n for n in re.split(r"[,，、\s]+", names_text) if n]
        if not names:
            self._append_log("请输入模型名称（多个用空格或逗号分隔）", "err")
            self._set_status("请先输入模型名称", COLOR_ERR)
            return
        try:
            copies = int(self.var_copies.get().strip() or "2")
            if copies < 1:
                raise ValueError
        except ValueError:
            self._append_log("每份张数必须是正整数", "err")
            return
        excel = self.var_excel.get().strip()
        template = self.var_template.get().strip()
        base_dir = self.var_base_dir.get().strip()
        printer = self.menu_printer.get()
        do_print = bool(self.switch_print.get())
        if not excel or not os.path.exists(excel):
            self._append_log("工作跟进表路径无效，请点「选择...」重新选择", "err")
            return
        if not template or not os.path.exists(template):
            self._append_log("质检表模板路径无效，请点「选择...」重新选择", "err")
            return
        if not base_dir:
            self._append_log("归档目录为空，请点「选择...」选择目录", "err")
            return
        if do_print and printer.startswith("（未找到"):
            self._append_log("系统中未找到打印机，无法打印；可关闭打印开关仅生成归档", "err")
            return

        self._save_config()
        self.running = True
        self.btn_run.configure(state="disabled", text="执行中...")
        self.lbl_result.configure(text="")
        self._append_log("═══ 任务开始 ═══", "info")
        t = threading.Thread(
            target=self._work,
            args=(names, excel, template, base_dir, printer, copies, do_print),
            daemon=True)
        t.start()

    def _work(self, names, excel, template, base_dir, printer, copies, do_print):
        q = queue.Queue()
        try:
            with contextlib.redirect_stdout(QueueWriter(q)):
                orders = qc.load_orders(excel, names, chooser=self._choose_order)
                print("[1/4] 生成质检表 ...")
                temp_files = []
                for o in orders:
                    t = qc.fill_report(template, o)
                    temp_files.append((t, o))
                    print("  已生成: %s" % os.path.basename(t))
                print("[2/4] 归档到订单文件夹 ...")
                finals = []
                for t, o in temp_files:
                    f = qc.archive_report(t, base_dir, o)
                    finals.append(f)
                    print("  已归档: %s" % f)
                if do_print:
                    print("[3/4] 检查打印机 ...")
                    printers = qc.list_printers()
                    if printer not in printers:
                        raise RuntimeError("打印机 '%s' 不在系统列表" % printer)
                    print("[4/4] 开始打印（每份 %d 张）..." % copies)
                    qc.print_reports(finals, printer, copies, auto_yes=True)
                else:
                    print("[3/4] 已跳过打印")
                print("\n═══ 全部完成 ═══")
                self.after(0, lambda: self._on_done(finals))
        except Exception as e:
            self.after(0, lambda: self._on_fail(str(e)))
        finally:
            while True:
                try:
                    self.log_q.put(q.get_nowait())
                except queue.Empty:
                    break

    def _poll_logs(self):
        try:
            while True:
                s = self.log_q.get_nowait()
                self._append_log(s.rstrip("\n"), "info")
        except queue.Empty:
            pass
        if self.running:
            self._set_status("● 运行中", COLOR_ACCENT)
        self.after(100, self._poll_logs)

    def _append_log(self, text, tag="info"):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", text + "\n", tag)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _set_status(self, text, color):
        self.lbl_status.configure(text=text, text_color=color)

    def _choose_order(self, name, candidates):
        """工作线程中调用：调度主线程弹窗选择，阻塞等待用户选择结果。"""
        evt = threading.Event()
        result = {}

        def ask():
            dialog = OrderSelectDialog(self, name, candidates, result)
            self.wait_window(dialog)
            evt.set()

        self.after(0, ask)  # 弹窗必须在主线程创建
        evt.wait()
        if "chosen" not in result:
            raise RuntimeError("未选择「%s」的订单，任务取消" % name)
        return result["chosen"]

    def _on_done(self, finals):
        self.running = False
        self.btn_run.configure(state="normal", text="开始执行")
        self._set_status("● 完成", COLOR_OK)
        self.lbl_result.configure(
            text="已归档 %d 个文件到 %s" % (len(finals), self.var_base_dir.get()))

    def _on_fail(self, msg):
        self.running = False
        self.btn_run.configure(state="normal", text="开始执行")
        self._set_status("● 失败", COLOR_ERR)
        self._append_log("错误: " + msg, "err")


if __name__ == "__main__":
    App().mainloop()
