"""
Unified Dashboard - Final Version
整合回測 + 交易操作建議
三大功能: 更新資料 | 回測 | 當前操作建議
"""
import sys, os
if getattr(sys, 'frozen', False):
    _internal = os.path.join(os.path.dirname(sys.executable), '_internal')
    os.environ.setdefault('TCL_LIBRARY', os.path.join(_internal, '_tcl_data'))
    os.environ.setdefault('TK_LIBRARY', os.path.join(_internal, '_tk_data'))

import matplotlib
matplotlib.use('TkAgg')
matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'Microsoft YaHei',
                                           'SimHei', 'Segoe UI']
matplotlib.rcParams['axes.unicode_minus'] = False

import tkinter as tk
from tkinter import ttk, messagebox
import re
import os
import json
import subprocess
import webbrowser
import threading
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import config_final as config
from selection import SelectionEngine

import sys as _sys
_BASE = (os.path.dirname(_sys.executable)
         if getattr(_sys, 'frozen', False)
         else os.path.dirname(os.path.abspath(__file__)))
SAVE_FILE = os.path.join(_BASE, "trading_dashboard_state.json")


class UnifiedDashboard:
    # ──────────── 配色（大地色系）────────────
    BG       = '#f5f0eb'    # 暖米色底
    CARD     = '#fdfcfa'    # 卡片暖白
    FG       = '#3d2e1e'    # 深棕文字
    FG2      = '#8b7355'    # 次要咖啡文字
    ACCENT   = '#8b6914'    # 琥珀色強調
    ACCENT_H = '#7a5c10'    # 按鈕 hover
    ACCENT_P = '#6b4f0c'    # 按鈕 pressed
    GREEN    = '#5d8c3e'    # 橄欖綠
    RED      = '#b5503a'    # 赤陶紅
    BORDER   = '#ddd5c8'    # 暖灰邊線
    INPUT_BG = '#faf7f2'    # 輸入框底色

    def __init__(self, root):
        self.root = root
        self.root.title("全天候SP500動能策略")
        self.root.configure(bg=self.BG)

        # 視窗最大化
        self.root.state('zoomed')

        self.selector = None
        self.entries = {}

        self.params = [
            ("起始日期",              "START_DATE",           str),
            ("初始資金 ($)",          "INITIAL_CASH",         int),
            ("目標持股數",            "TARGET_HOLDINGS",      int),
            ("再平衡星期",            "REBALANCE_WEEKDAY",    int),
            ("停損比例 %",            "STOP_LOSS_PCT",        float),
            ("賣出排名門檻",          "SELL_RANK_THRESHOLD",  int),
            ("動能計算天數",           "LOOKBACK",             int),
            ("出場 EMA 週期",         "EXIT_EMA",             int),
            ("ATR 計算天數",          "ATR_PERIOD",           int),
            ("殘差相關過濾",          "CORR_FILTER_ENABLED",  bool),
            ("相關係數門檻",          "CORR_THRESHOLD",       float),
            ("殘差回看天數",          "CORR_LOOKBACK",        int),
            ("候選股池大小",          "CORR_CANDIDATE_COUNT", int),
        ]
        # 隱藏參數（仍由 config_final.py 控制，不在儀表板顯示）
        # COMMISSION, REBALANCE_WEEKS, REBALANCE_THRESHOLD
        self.weekday_map = {0: "週一", 1: "週二", 2: "週三", 3: "週四", 4: "週五"}
        self.weekday_inv_map = {v: k for k, v in self.weekday_map.items()}
        self.bool_map = {True: "啟用", False: "停用"}
        self.bool_inv_map = {v: k for k, v in self.bool_map.items()}

        self._setup_validation()
        self._setup_styles()
        self._build_ui()
        self._load_config()
        self._load_trading_state()

    # ──────────── Validation ────────────
    def _setup_validation(self):
        def validate_float(v):
            if v == "":
                return True
            return bool(re.match(r'^-?(\d+\.?\d*|\.\d+)$', v))
        self.vcmd = (self.root.register(validate_float), '%P')

    # ──────────── Styles ────────────
    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use('clam')

        # ── 按鈕 ──
        s.configure('Action.TButton',
                     font=('Segoe UI', 10, 'bold'),
                     padding=(20, 12),
                     borderwidth=0,
                     relief='flat')
        s.map('Action.TButton',
              background=[('active', self.ACCENT_H),
                          ('pressed', self.ACCENT_P),
                          ('disabled', '#c4b49a'),
                          ('!disabled', self.ACCENT)],
              foreground=[('disabled', '#a89878'),
                          ('!disabled', '#ffffff')])

        # 淡色次要按鈕（執行回測用）
        s.configure('Secondary.TButton',
                     font=('Segoe UI', 10, 'bold'),
                     padding=(20, 12),
                     borderwidth=0,
                     relief='flat')
        s.map('Secondary.TButton',
              background=[('active', '#ede4d4'),
                          ('pressed', '#e0d5c0'),
                          ('disabled', '#f5f0eb'),
                          ('!disabled', '#f0e8d8')],
              foreground=[('disabled', '#b8a88e'),
                          ('!disabled', self.ACCENT)])

        s.configure('Small.TButton',
                     font=('Segoe UI', 9),
                     padding=(10, 6),
                     borderwidth=1,
                     relief='solid')
        s.map('Small.TButton',
              background=[('active', '#f0e8d8'),
                          ('pressed', '#e0d5c0'),
                          ('disabled', '#faf7f2'),
                          ('!disabled', self.CARD)],
              foreground=[('disabled', '#c4b49a'),
                          ('!disabled', self.FG)],
              bordercolor=[('active', self.ACCENT),
                           ('!disabled', self.BORDER)])

        # ── Entry / Combobox ──
        s.configure('In.TEntry',
                     fieldbackground=self.INPUT_BG,
                     foreground=self.FG,
                     borderwidth=1, relief='solid',
                     padding=6)
        s.map('In.TEntry',
              bordercolor=[('focus', self.ACCENT),
                           ('!focus', self.BORDER)])

        s.configure('In.TCombobox',
                     fieldbackground=self.INPUT_BG,
                     foreground=self.FG,
                     borderwidth=1, relief='solid',
                     padding=6)
        s.map('In.TCombobox',
              bordercolor=[('focus', self.ACCENT),
                           ('!focus', self.BORDER)],
              fieldbackground=[('readonly', self.INPUT_BG)],
              arrowcolor=[('!disabled', self.ACCENT)])

        # ── Progressbar ──
        try:
            s.layout('Bar.TProgressbar',
                     s.layout('Horizontal.TProgressbar'))
            s.configure('Bar.TProgressbar',
                         background=self.ACCENT,
                         troughcolor=self.BORDER,
                         borderwidth=0)
        except Exception:
            pass

    # ──────────── 輔助：畫分隔線 ────────────
    def _sep(self, parent, bg=None, **pack_kw):
        f = tk.Frame(parent, height=1, bg=bg or self.BORDER)
        f.pack(fill=tk.X, **pack_kw)
        return f

    # ──────────── 輔助：面板標題 ────────────
    def _panel_title(self, parent, text):
        # 藍色左邊條 + 標題
        title_row = tk.Frame(parent, bg=self.CARD)
        title_row.pack(fill=tk.X, padx=16, pady=(16, 0))
        accent_bar = tk.Frame(title_row, width=4, bg=self.ACCENT)
        accent_bar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        lbl = tk.Label(title_row, text=text,
                       font=('Segoe UI', 11, 'bold'),
                       fg=self.FG, bg=self.CARD, anchor='w')
        lbl.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=2)
        self._sep(parent, padx=16, pady=(10, 10))

    # ══════════════════════════════════════════════
    #  Build UI
    # ══════════════════════════════════════════════
    def _build_ui(self):
        # ── Header（白色頂欄）──
        hdr_bg = tk.Frame(self.root, bg=self.CARD)
        hdr_bg.pack(fill=tk.X)

        hdr = tk.Frame(hdr_bg, bg=self.CARD)
        hdr.pack(fill=tk.X, padx=28, pady=(18, 0))

        tk.Label(hdr, text="全天候SP500動能策略",
                 font=('Segoe UI', 22, 'bold'),
                 fg=self.FG, bg=self.CARD).pack(anchor=tk.W)
        tk.Label(hdr, text="回測分析  ·  交易操作建議",
                 font=('Segoe UI', 10),
                 fg=self.FG2, bg=self.CARD).pack(anchor=tk.W, pady=(4, 0))

        # ── Buttons ──
        btn_bar = tk.Frame(hdr_bg, bg=self.CARD)
        btn_bar.pack(fill=tk.X, padx=28, pady=(16, 0))

        for text, cmd, sty in [
            ("更新資料", self.update_data_process, 'Action.TButton'),
            ("操作建議", self.calculate_trades,     'Action.TButton'),
            ("執行回測", self.run_backtest,          'Secondary.TButton'),
        ]:
            b = ttk.Button(btn_bar, text=text, command=cmd,
                           style=sty, cursor='hand2')
            b.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=4)

        # ── Progress ──
        prog = tk.Frame(hdr_bg, bg=self.CARD)
        prog.pack(fill=tk.X, padx=28, pady=(14, 0))

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(prog, variable=self.progress_var,
                                            maximum=100, style='Bar.TProgressbar')
        self.progress_bar.pack(fill=tk.X, ipady=1)

        self.status_lbl = tk.Label(prog, text="準備就緒",
                                   font=('Segoe UI', 9),
                                   fg=self.FG2, bg=self.CARD, anchor='w')
        self.status_lbl.pack(fill=tk.X, pady=(6, 14))

        # ── Main 3‑column area（灰底，卡片凸顯）──
        main = tk.Frame(self.root, bg=self.BG)
        main.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        # 內層加 padding
        inner_main = tk.Frame(main, bg=self.BG)
        inner_main.pack(fill=tk.BOTH, expand=True, padx=20, pady=(16, 20))
        inner_main.columnconfigure(0, weight=0, minsize=310)
        inner_main.columnconfigure(1, weight=0, minsize=180)
        inner_main.columnconfigure(2, weight=1)
        inner_main.rowconfigure(0, weight=1)

        self._build_config_panel(inner_main)
        self._build_input_panel(inner_main)
        self._build_results_panel(inner_main)

        # 快速引用按鈕
        self.update_btn   = btn_bar.winfo_children()[0]
        self.trade_btn    = btn_bar.winfo_children()[1]
        self.backtest_btn = btn_bar.winfo_children()[2]

    # ──────────── 左欄：參數設定（固定，無捲動）────────────
    def _build_config_panel(self, parent):
        col = tk.Frame(parent, bg=self.CARD, bd=0, relief='flat',
                       highlightbackground=self.BORDER,
                       highlightthickness=1)
        col.grid(row=0, column=0, sticky='nsew', padx=(0, 10))

        self._panel_title(col, "參數設定")

        for label_text, var_name, dtype in self.params:
            row = tk.Frame(col, bg=self.CARD)
            row.pack(fill=tk.X, padx=16, pady=4)

            tk.Label(row, text=label_text,
                     font=('Segoe UI', 9), fg=self.FG, bg=self.CARD,
                     anchor='w', width=22).pack(side=tk.LEFT)

            if var_name == "REBALANCE_WEEKDAY":
                w = ttk.Combobox(row, values=list(self.weekday_map.values()),
                                 state='readonly', style='In.TCombobox', width=10)
            elif var_name == "CORR_FILTER_ENABLED":
                w = ttk.Combobox(row, values=list(self.bool_map.values()),
                                 state='readonly', style='In.TCombobox', width=10)
            elif dtype == float:
                w = ttk.Entry(row, validate='key', validatecommand=self.vcmd,
                              style='In.TEntry', width=12,
                              font=('Segoe UI', 9))
            else:
                w = ttk.Entry(row, style='In.TEntry', width=12,
                              font=('Segoe UI', 9))
            w.pack(side=tk.RIGHT, fill=tk.X, expand=True)
            self.entries[var_name] = w

    # ──────────── 中欄：當前持股 ────────────
    def _build_input_panel(self, parent):
        col = tk.Frame(parent, bg=self.CARD, bd=0, relief='flat',
                       highlightbackground=self.BORDER,
                       highlightthickness=1, width=300)
        col.grid(row=0, column=1, sticky='ns', padx=5)
        col.grid_propagate(False)
        col.pack_propagate(False)

        self._panel_title(col, "當前狀態")

        body = tk.Frame(col, bg=self.CARD)
        body.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 10))

        # 總權益
        r0 = tk.Frame(body, bg=self.CARD)
        r0.pack(fill=tk.X, pady=(0, 12))
        tk.Label(r0, text="總權益 ($)",
                 font=('Segoe UI', 9), fg=self.FG2, bg=self.CARD).pack(anchor=tk.W)
        self.equity_entry = ttk.Entry(r0, style='In.TEntry',
                                      font=('Segoe UI', 11))
        self.equity_entry.pack(fill=tk.X, pady=(4, 0))
        self.equity_entry.insert(0, "1434200")

        # 持股表頭
        self._sep(body)
        tk.Label(body, text="持股明細",
                 font=('Segoe UI', 9, 'bold'),
                 fg=self.FG, bg=self.CARD, anchor='w').pack(fill=tk.X, pady=(10, 6))

        hdr = tk.Frame(body, bg=self.CARD)
        hdr.pack(fill=tk.X)
        hdr.columnconfigure(0, weight=1, uniform='col')
        hdr.columnconfigure(1, weight=1, uniform='col')
        tk.Label(hdr, text="股票代號", font=('Segoe UI', 8, 'bold'),
                 fg=self.FG2, bg=self.CARD, anchor='w').grid(
                 row=0, column=0, sticky='w', padx=(2, 4))
        tk.Label(hdr, text="股數", font=('Segoe UI', 8, 'bold'),
                 fg=self.FG2, bg=self.CARD, anchor='w').grid(
                 row=0, column=1, sticky='w', padx=(2, 0))

        self.holdings_frame = tk.Frame(body, bg=self.CARD)
        self.holdings_frame.pack(fill=tk.X, pady=(4, 8))
        self.holding_rows = []

        for t, q in [("WDC", "1700"), ("ALB", "2114"),
                      ("DG", "2000"), ("WBD", "10219")]:
            self._add_holding_row(t, q)

        bf = tk.Frame(body, bg=self.CARD)
        bf.pack(fill=tk.X)
        ttk.Button(bf, text="＋ 新增", width=10,
                   command=lambda: self._add_holding_row(),
                   style='Small.TButton', cursor='hand2').pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(bf, text="－ 移除", width=10,
                   command=self._remove_holding_row,
                   style='Small.TButton', cursor='hand2').pack(side=tk.LEFT)

    # ──────────── 右欄：單一可滾動頁面 ────────────
    def _build_results_panel(self, parent):
        outer = tk.Frame(parent, bg=self.CARD, bd=0, relief='flat',
                         highlightbackground=self.BORDER,
                         highlightthickness=1)
        outer.grid(row=0, column=2, sticky='nsew', padx=(10, 0))

        # 可滾動 Canvas
        canvas = tk.Canvas(outer, bg=self.CARD, highlightthickness=0, bd=0)
        sb = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=canvas.yview)
        self._results_inner = tk.Frame(canvas, bg=self.CARD)

        self._results_inner.bind('<Configure>',
            lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        win_id = canvas.create_window((0, 0), window=self._results_inner,
                                      anchor='nw')
        canvas.bind('<Configure>',
            lambda e: canvas.itemconfig(win_id, width=e.width))
        canvas.configure(yscrollcommand=sb.set)

        # 滑鼠滾輪綁定到此 canvas
        def _on_mousewheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')
        canvas.bind('<Enter>',
            lambda e: canvas.bind_all('<MouseWheel>', _on_mousewheel))
        canvas.bind('<Leave>',
            lambda e: canvas.unbind_all('<MouseWheel>'))

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        inner = self._results_inner

        # ── Section 1+2: 操作建議 (左) + 目標持股比例 (右) ──
        top_row = tk.Frame(inner, bg=self.CARD)
        top_row.pack(fill=tk.X, padx=0, pady=0)
        top_row.columnconfigure(0, weight=1, uniform='half')
        top_row.columnconfigure(1, weight=0)          # divider
        top_row.columnconfigure(2, weight=1, uniform='half')

        # 左半 — 操作建議
        left_frame = tk.Frame(top_row, bg=self.CARD)
        left_frame.grid(row=0, column=0, sticky='new')
        self._panel_title(left_frame, "操作建議")
        self.results_text = self._make_rich_text(left_frame)

        # 中間分隔線
        div = tk.Frame(top_row, bg=self.BORDER, width=1)
        div.grid(row=0, column=1, sticky='ns', pady=8)

        # 右半 — 目標持股比例（不擴展，由 figure 決定高度）
        right_frame = tk.Frame(top_row, bg=self.CARD)
        right_frame.grid(row=0, column=2, sticky='new')
        self._panel_title(right_frame, "目標持股比例")
        self.weight_fig = Figure(figsize=(5, 1.6), dpi=100)
        self.weight_fig.patch.set_facecolor(self.CARD)
        self.weight_canvas = FigureCanvasTkAgg(self.weight_fig,
                                               master=right_frame)
        self.weight_canvas.get_tk_widget().pack(
            fill=tk.X, padx=8, pady=(0, 0))

        # 水平分隔線
        tk.Frame(inner, bg=self.BORDER, height=1).pack(
            fill=tk.X, padx=16, pady=(2, 2))

        # ── Section 3: 市場排名 ──
        self._panel_title(inner, "市場動能排名 — 前 20 名")
        self.slope_fig = Figure(figsize=(5, 5.5), dpi=100)
        self.slope_fig.patch.set_facecolor(self.CARD)
        self.slope_canvas = FigureCanvasTkAgg(self.slope_fig, master=inner)
        self.slope_canvas.get_tk_widget().pack(
            fill=tk.X, padx=16, pady=(0, 20))

    def _make_rich_text(self, parent):
        """建立操作建議用的 rich text widget"""
        wrapper = tk.Frame(parent, bg=self.CARD)
        wrapper.pack(fill=tk.X, padx=1, pady=(0, 4))

        txt = tk.Text(wrapper,
                      font=('Segoe UI', 10),
                      wrap=tk.WORD, state=tk.DISABLED,
                      bg=self.CARD, fg=self.FG,
                      selectbackground=self.ACCENT,
                      selectforeground='#ffffff',
                      relief='flat', bd=0,
                      padx=18, pady=10,
                      height=10,
                      spacing1=2, spacing3=2)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        vs = ttk.Scrollbar(wrapper, orient=tk.VERTICAL, command=txt.yview)
        vs.pack(side=tk.RIGHT, fill=tk.Y)
        txt.config(yscrollcommand=vs.set)

        # Apple-style tags
        txt.tag_configure('date',
                          font=('Segoe UI', 9), foreground=self.FG2,
                          spacing3=6)
        txt.tag_configure('section',
                          font=('Segoe UI', 11, 'bold'), foreground=self.FG,
                          spacing1=10, spacing3=4)
        txt.tag_configure('sell',
                          font=('Consolas', 10, 'bold'), foreground=self.RED,
                          lmargin1=4, lmargin2=4, spacing1=3, spacing3=3,
                          tabs=('60p', '130p', '210p'))
        txt.tag_configure('buy',
                          font=('Consolas', 10, 'bold'), foreground=self.GREEN,
                          lmargin1=4, lmargin2=4, spacing1=3, spacing3=3,
                          tabs=('60p', '130p', '210p'))
        txt.tag_configure('ok',
                          font=('Segoe UI', 10), foreground=self.GREEN,
                          lmargin1=4, spacing1=6, spacing3=6)
        txt.tag_configure('dim',
                          font=('Segoe UI', 9), foreground=self.FG2,
                          lmargin1=4, lmargin2=4)
        txt.tag_configure('divider',
                          font=('Segoe UI', 2), foreground=self.BORDER,
                          spacing1=4, spacing3=4)
        txt.tag_configure('subsection',
                          font=('Segoe UI', 10, 'bold'), foreground=self.FG,
                          spacing1=10, spacing3=3)
        txt.tag_configure('warn',
                          font=('Segoe UI', 9), foreground='#ea580c',
                          lmargin1=4, lmargin2=4)
        txt.tag_configure('mono',
                          font=('Consolas', 9), foreground=self.FG,
                          lmargin1=4, lmargin2=4, spacing1=1, spacing3=1)
        return txt

    # ──────────── Holding rows ────────────
    def _add_holding_row(self, ticker="", qty=""):
        row = tk.Frame(self.holdings_frame, bg=self.CARD)
        row.pack(fill=tk.X, pady=3)
        row.columnconfigure(0, weight=1, uniform='col')
        row.columnconfigure(1, weight=1, uniform='col')
        te = ttk.Entry(row, style='In.TEntry', font=('Segoe UI', 9))
        te.grid(row=0, column=0, sticky='ew', padx=(0, 4))
        te.insert(0, ticker)
        qe = ttk.Entry(row, style='In.TEntry', font=('Segoe UI', 9))
        qe.grid(row=0, column=1, sticky='ew', padx=(0, 0))
        qe.insert(0, qty)
        self.holding_rows.append((te, qe, row))

    def _remove_holding_row(self):
        if self.holding_rows:
            _, _, f = self.holding_rows.pop()
            f.destroy()

    def parse_holdings(self):
        holdings = {}
        for te, qe, _ in self.holding_rows:
            t = te.get().strip().upper()
            qs = qe.get().strip()
            if t and qs:
                try:
                    q = int(qs)
                    if q > 0:
                        holdings[t] = q
                except ValueError:
                    pass
        return holdings

    # ══════════════════════════════════════════════
    #  Config load / save
    # ══════════════════════════════════════════════
    def _get_base_dir(self):
        """取得程式所在目錄（支援 PyInstaller 打包後的路徑）"""
        import sys
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def _get_config_path(self):
        return os.path.join(self._get_base_dir(), "config_final.py")

    def _reload_config_from_file(self):
        """從磁碟上的 config_final.py 強制重新載入，繞過 PyInstaller 的打包快取"""
        global config
        import importlib.util
        path = self._get_config_path()
        spec = importlib.util.spec_from_file_location("config_final", path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        _sys.modules['config_final'] = mod
        config = mod

    def _load_config(self):
        path = self._get_config_path()
        try:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
            for _, var, _ in self.params:
                m = re.search(rf'^{var}\s*=\s*(.+?)(?:\s*#|$)',
                              content, re.MULTILINE)
                if m:
                    val = m.group(1).strip().strip("'\"")
                    if var == "REBALANCE_WEEKDAY":
                        try:
                            self.entries[var].set(
                                self.weekday_map.get(int(val), "週三"))
                        except Exception:
                            self.entries[var].set("週三")
                    elif var == "CORR_FILTER_ENABLED":
                        bool_val = val.strip() == 'True'
                        self.entries[var].set(
                            self.bool_map.get(bool_val, "啟用"))
                    else:
                        self.entries[var].delete(0, tk.END)
                        self.entries[var].insert(0, val)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load config: {e}")

    def _save_config(self):
        path = self._get_config_path()
        try:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
            new = content
            for _, var, dtype in self.params:
                val = self.entries[var].get()
                if var == "REBALANCE_WEEKDAY":
                    fmt = str(self.weekday_inv_map.get(val, 2))
                elif var == "CORR_FILTER_ENABLED":
                    fmt = str(self.bool_inv_map.get(val, True))
                elif dtype == str:
                    fmt = f"'{val}'"
                else:
                    fmt = val
                pat = rf'^({var}\s*=\s*)(.+?)(?=\s*#|$)'
                if re.search(pat, new, re.MULTILINE):
                    new = re.sub(pat, rf'\g<1>{fmt}', new, flags=re.MULTILINE)
            with open(path, 'w', encoding='utf-8') as f:
                f.write(new)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save config: {e}")
            return False

    # ══════════════════════════════════════════════
    #  Trading state save / load
    # ══════════════════════════════════════════════
    def _save_trading_state(self):
        state = {
            "equity": self.equity_entry.get().strip(),
            "holdings": [(t.get().strip(), q.get().strip())
                         for t, q, _ in self.holding_rows if t.get().strip()]
        }
        try:
            with open(SAVE_FILE, 'w', encoding='utf-8') as f:
                json.dump(state, f, ensure_ascii=False)
        except Exception:
            pass

    def _load_trading_state(self):
        if not os.path.exists(SAVE_FILE):
            return
        try:
            with open(SAVE_FILE, 'r', encoding='utf-8') as f:
                state = json.load(f)
            eq = state.get("equity", "")
            if eq:
                self.equity_entry.delete(0, tk.END)
                self.equity_entry.insert(0, eq)
            saved = state.get("holdings", [])
            if saved:
                for _, _, frame in self.holding_rows:
                    frame.destroy()
                self.holding_rows.clear()
                for t, q in saved:
                    self._add_holding_row(t, q)
        except Exception:
            pass

    # ══════════════════════════════════════════════
    #  Button 1 : Update Data
    # ══════════════════════════════════════════════
    def update_data_process(self):
        self._set_buttons(False)
        self.status_lbl.config(text="正在從 Yahoo Finance 更新數據...",
                               fg=self.ACCENT)
        self.progress_var.set(0)
        threading.Thread(target=self._execute_update_data, daemon=True).start()

    def _execute_update_data(self):
        try:
            import sys
            base = self._get_base_dir()
            if base not in sys.path:
                sys.path.insert(0, base)
            import update_data
            import importlib
            importlib.reload(update_data)

            self.root.after(0, lambda: self.status_lbl.config(
                text="正在下載最新數據..."))
            update_data.main()

            self.selector = None
            self.root.after(0, lambda: self.status_lbl.config(
                text="數據更新成功！", fg=self.GREEN))
            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: messagebox.showinfo("成功", "數據更新完成"))
        except Exception as e:
            self.root.after(0, lambda: self.status_lbl.config(
                text=f"更新失敗: {e}", fg=self.RED))
            self.root.after(0, lambda: messagebox.showerror("錯誤", str(e)))
        finally:
            self.root.after(0, lambda: self._set_buttons(True))

    # ══════════════════════════════════════════════
    #  Button 2 : Backtest
    # ══════════════════════════════════════════════
    def run_backtest(self):
        if not self._save_config():
            return
        self._set_buttons(False)
        self.status_lbl.config(text="正在執行回測...", fg=self.ACCENT)
        self.progress_var.set(0)
        threading.Thread(target=self._execute_backtest, daemon=True).start()

    def _execute_backtest(self):
        try:
            import sys
            base = self._get_base_dir()
            if base not in sys.path:
                sys.path.insert(0, base)

            self._reload_config_from_file()

            self.root.after(0, lambda: self.status_lbl.config(
                text="正在執行回測..."))

            import importlib
            import selection
            import portfolio_backtester_final
            import run_strategy_final
            importlib.reload(selection)
            importlib.reload(portfolio_backtester_final)
            importlib.reload(run_strategy_final)
            run_strategy_final.main()

            # 開啟報告
            report = os.path.join(base, "strategy_report_final.html")
            if os.path.exists(report):
                webbrowser.open(report)

            self.root.after(0, lambda: self.status_lbl.config(
                text="回測完成！報告已開啟", fg=self.GREEN))
            self.root.after(0, lambda: self.progress_var.set(100))
        except Exception as e:
            self.root.after(0, lambda: self.status_lbl.config(
                text=f"回測失敗: {e}", fg=self.RED))
            self.root.after(0, lambda:
                messagebox.showerror("回測失敗", str(e)))
        finally:
            self.root.after(0, lambda: self._set_buttons(True))

    # ══════════════════════════════════════════════
    #  Button 3 : Trading Recommendations
    # ══════════════════════════════════════════════
    def calculate_trades(self):
        if not self._save_config():
            return
        self._reload_config_from_file()
        self.selector = None

        try:
            total_equity = float(self.equity_entry.get())
            holdings = self.parse_holdings()
            if total_equity <= 0:
                messagebox.showerror("錯誤", "總權益必須大於 0")
                return
            self._save_trading_state()
            self._set_buttons(False)
            self.status_lbl.config(text="正在計算操作建議...", fg=self.ACCENT)
            threading.Thread(target=self._calculate_trades,
                             args=(total_equity, holdings), daemon=True).start()
        except ValueError:
            messagebox.showerror("錯誤", "請輸入有效的數字")

    def _calculate_trades(self, total_equity, holdings):
        try:
            if self.selector is None:
                self.selector = SelectionEngine()
                self.selector.preload_all_data()

            spy_data = self.selector._get_ticker_data('SPY')
            if spy_data is None or spy_data.empty:
                self.root.after(0, lambda: messagebox.showerror(
                    "錯誤", "無法讀取 SPY 數據，請先更新"))
                return

            latest_date = spy_data.index[-1]

            entry_scan = self.selector.scan_market(latest_date,
                                                    lookback=config.LOOKBACK)
            exit_scan = self.selector.scan_market(latest_date,
                                                   lookback=config.LOOKBACK)

            # ── 結構化輸出（tag, text）──
            out = []  # list of (tag, text)

            def get_price_ema(ticker):
                td = self.selector._get_ticker_data(ticker)
                if td is None or latest_date not in td.index:
                    return None, None, None
                price = float(td.loc[latest_date, 'Close'])
                ema_col = f'_EMA{config.EXIT_EMA}'
                if ema_col not in td.columns:
                    pc = 'Adj Close' if 'Adj Close' in td.columns else 'Close'
                    td[ema_col] = td[pc].ewm(span=config.EXIT_EMA,
                                             adjust=False).mean()
                ema = float(td.loc[latest_date, ema_col])
                return td, price, ema

            def get_rank(ticker, scan_list):
                for idx, m in enumerate(scan_list, 1):
                    if m.get('ticker') == ticker:
                        return idx
                return None

            def calc_atr_weights(ticker_list):
                items = []
                for t in ticker_list:
                    m = self.selector.calculate_metrics(t, latest_date,
                                                        config.LOOKBACK)
                    if m and m.get('atr_pct', 0) > 0:
                        items.append(m)
                if not items:
                    return {}
                inv_sum = sum(1 / c['atr_pct'] for c in items)
                return {c['ticker']: (1 / c['atr_pct']) / inv_sum for c in items}

            # ── Step 1: 輪動賣出判斷 ──
            rotation_sells = []
            keep_holdings = {}

            for ticker, qty in sorted(holdings.items()):
                td, price, ema = get_price_ema(ticker)
                if price is None:
                    continue
                rank_exit = get_rank(ticker, exit_scan)
                below_ema = price < ema
                not_in_top = ((rank_exit is None) or
                              (rank_exit > config.SELL_RANK_THRESHOLD))
                reasons = []
                if not_in_top:
                    rank_str = (f"#{rank_exit}" if rank_exit
                                else "未上榜")
                    reasons.append(f"排名{rank_str}")
                if below_ema:
                    reasons.append(f"跌破 EMA{config.EXIT_EMA}")
                if reasons:
                    rotation_sells.append((ticker, qty, price, reasons))
                else:
                    keep_holdings[ticker] = qty

            # ── Step 2: 組合 + ATR 權重 ──
            rebalance_threshold = max(config.REBALANCE_THRESHOLD, 0.01)
            total_sell_value = sum(p * q for _, q, p, _ in rotation_sells)

            current_positions = len(keep_holdings)
            needed = config.TARGET_HOLDINGS - current_positions

            max_adj_slope = getattr(config, 'MAX_ADJ_SLOPE', None)
            skip_gap = getattr(config, 'SKIP_MAX_GAP_PCT', 0.20)
            filtered_entry = entry_scan[:]
            if max_adj_slope is not None:
                filtered_entry = [x for x in filtered_entry
                                  if x.get('adj_slope', 999) < max_adj_slope]
            filtered_entry = [x for x in filtered_entry
                              if x.get('max_gap', 0) < skip_gap]
            entry_tickers = [x['ticker'] for x in filtered_entry]
            buy_candidates = [t for t in entry_tickers
                              if t not in keep_holdings]

            if getattr(config, 'CORR_FILTER_ENABLED', False) and needed > 0:
                import utils
                spy_path = os.path.join(config.DATA_DIR, 'SPY.csv')
                spy_df = utils.load_benchmark_data(spy_path)
                candidate_metrics = [x for x in filtered_entry
                                     if x['ticker'] in buy_candidates]
                to_buy = self.selector.filter_by_residual_correlation(
                    ranked_candidates=candidate_metrics,
                    date=latest_date,
                    spy_df=spy_df,
                    threshold=config.CORR_THRESHOLD,
                    lookback=config.CORR_LOOKBACK,
                    max_candidates=config.CORR_CANDIDATE_COUNT,
                    needed=needed,
                    existing_tickers=list(keep_holdings.keys())
                )
            else:
                to_buy = buy_candidates[:needed] if needed > 0 else []

            all_tickers = list(keep_holdings.keys()) + to_buy
            full_weights = calc_atr_weights(all_tickers)

            # ── Step 3: 超重賣出 ──
            overweight_sells = []
            if full_weights:
                for t, q in list(keep_holdings.items()):
                    target_w = full_weights.get(t, 0)
                    _, price, _ = get_price_ema(t)
                    if not price or not target_w:
                        continue
                    current_value = price * q
                    current_w = current_value / total_equity
                    deviation = current_w - target_w
                    if deviation >= rebalance_threshold:
                        target_value = total_equity * target_w
                        diff_value = current_value - target_value
                        if diff_value >= price:
                            qty_to_sell = min(int(diff_value / price), q)
                            if qty_to_sell > 0:
                                reason = (f"{current_w*100:.1f}% → "
                                          f"{target_w*100:.1f}%")
                                overweight_sells.append(
                                    (t, qty_to_sell, price, reason,
                                     current_w*100))
                                total_sell_value += qty_to_sell * price
                                keep_holdings[t] -= qty_to_sell
                                if keep_holdings[t] <= 0:
                                    del keep_holdings[t]

            # ── Step 4: 買入新股票 ──
            holdings_value = 0
            for t, q in holdings.items():
                _, p, _ = get_price_ema(t)
                if p:
                    holdings_value += p * q
            current_cash = total_equity - holdings_value
            available_cash = current_cash + total_sell_value

            new_buys = []
            newly_bought = set()
            current_atr_weights = full_weights

            if needed > 0 and to_buy:
                for ticker in to_buy:
                    w = full_weights.get(ticker, 0)
                    if w <= 0:
                        continue
                    alloc = total_equity * w
                    _, price, _ = get_price_ema(ticker)
                    if not price or price <= 0:
                        continue
                    buy_amount = min(alloc, available_cash)
                    qty = int(buy_amount / (price * (1 + config.COMMISSION)))
                    if qty > 0:
                        cost = qty * price * (1 + config.COMMISSION)
                        new_buys.append((ticker, qty, price, w))
                        newly_bought.add(ticker)
                        available_cash -= cost

            # ── Step 5: 低配補足 ──
            min_buy_pct = getattr(config, 'MIN_BUY_AMOUNT_PCT', 0.03)
            min_buy_amount = total_equity * min_buy_pct
            rebalance_buys = []

            if available_cash > 0 and current_atr_weights:
                underweight = []
                for t, q in keep_holdings.items():
                    if t in newly_bought:
                        continue
                    target_w = current_atr_weights.get(t, 0)
                    if target_w <= 0:
                        continue
                    _, price, _ = get_price_ema(t)
                    if not price:
                        continue
                    current_w = (price * q) / total_equity
                    shortfall = target_w - current_w
                    if shortfall >= rebalance_threshold:
                        underweight.append({
                            'ticker': t, 'price': price,
                            'shortfall': shortfall,
                            'target_w': target_w,
                            'current_w': current_w})
                if underweight:
                    total_shortfall = sum(s['shortfall'] for s in underweight)
                    alloc_cash = available_cash * 0.99
                    for s in underweight:
                        ratio = s['shortfall'] / total_shortfall
                        alloc = alloc_cash * ratio
                        qty = int(alloc / (s['price'] *
                                           (1 + config.COMMISSION)))
                        buy_amt = s['price'] * qty
                        if buy_amt >= min_buy_amount and qty > 0:
                            rebalance_buys.append(
                                (s['ticker'], qty, s['price'],
                                 s['shortfall']))

            # ════════════════════════════════════════
            #  組裝結構化輸出
            # ════════════════════════════════════════
            has_action = (rotation_sells or overweight_sells
                          or new_buys or rebalance_buys)

            # ── 文字摘要（等寬對齊）──
            if has_action:
                out.append(('section', "操作摘要"))

                # 蒐集所有操作行，統一格式化
                actions = []  # (tag, side, ticker, qty, reason)

                if rotation_sells:
                    for t, q, p, reasons in rotation_sells:
                        # 簡化原因
                        simple = []
                        for r in reasons:
                            if 'EMA' in r:
                                simple.append('跌破EMA')
                            elif '排名' in r or '未上榜' in r:
                                simple.append('動能排名下降')
                            else:
                                simple.append(r)
                        actions.append(('sell', 'SELL', t, q,
                                        ', '.join(simple)))
                if overweight_sells:
                    for t, q, p, reason, _ in overweight_sells:
                        actions.append(('sell', 'SELL', t, q,
                                        '再平衡調整'))
                if new_buys:
                    for t, q, p, w in new_buys:
                        actions.append(('buy', 'BUY ', t, q, '新倉'))
                if rebalance_buys:
                    for t, q, p, sf in rebalance_buys:
                        actions.append(('buy', 'BUY ', t, q,
                                        '再平衡調整'))

                # Tab 對齊輸出
                for tag, side, t, q, reason in actions:
                    out.append((tag,
                        f"{side}\t{t}\t{q} 股\t{reason}"))
            else:
                out.append(('section', "操作摘要"))
                out.append(('ok', "持倉均衡，無需操作"))

            # 組合概覽
            out.append(('divider', "─" * 40))
            out.append(('dim',
                f"保留 {current_positions} 檔  "
                f"新增 {max(needed, 0)} 檔  "
                f"目標 {config.TARGET_HOLDINGS} 檔"))

            # ── ATR 權重圖資料 ──
            sell_tickers = set(t for t, *_ in rotation_sells)
            sell_tickers.update(t for t, *_ in overweight_sells)
            buy_tickers = set(t for t, *_ in new_buys)
            buy_tickers.update(t for t, *_ in rebalance_buys)

            w_tickers = []
            w_target = []
            w_current = []
            w_colors = []
            for t, tw in sorted(full_weights.items(), key=lambda x: -x[1]):
                w_tickers.append(t)
                w_target.append(tw * 100)
                # 當前權重
                cw = 0.0
                if t in holdings:
                    _, p, _ = get_price_ema(t)
                    if p:
                        cw = (p * holdings[t]) / total_equity * 100
                w_current.append(cw)
                # 顏色
                if t in sell_tickers:
                    w_colors.append(self.RED)
                elif t in buy_tickers:
                    w_colors.append(self.GREEN)
                else:
                    w_colors.append(self.ACCENT)

            weight_data = {
                'tickers': w_tickers,
                'target_weights': w_target,
                'current_weights': w_current,
                'colors': w_colors,
            }

            # ── Slope 圖資料 ──
            # 調整完的持倉 = keep_holdings（保留）+ newly_bought（新買入）
            final_holdings = set(keep_holdings.keys()) | newly_bought
            s_tickers = []
            s_slopes = []
            s_colors = []
            s_labels = []
            for i, m in enumerate(entry_scan[:20], 1):
                t = m['ticker']
                slope = m.get('adj_slope', 0)
                s_tickers.append(t)
                s_slopes.append(slope)

                if t in final_holdings:
                    s_colors.append(self.GREEN)
                    if t in newly_bought:
                        s_labels.append('  新買入')
                    else:
                        s_labels.append('  持倉')
                else:
                    s_colors.append('#d5ccbe')
                    s_labels.append('')

            slope_data = {
                'tickers': s_tickers,
                'slopes': s_slopes,
                'colors': s_colors,
                'labels': s_labels,
            }

            display_data = {
                'text_lines': out,
                'weight_data': weight_data,
                'slope_data': slope_data,
            }

            self.root.after(0, lambda d=display_data:
                self._display_results(d))
            self.root.after(0, lambda: self.status_lbl.config(
                text="計算完成", fg=self.GREEN))
        except Exception as e:
            import traceback
            msg = f"計算錯誤:\n{e}\n\n{traceback.format_exc()}"
            self.root.after(0, lambda: messagebox.showerror("錯誤", msg))
            self.root.after(0, lambda: self.status_lbl.config(
                text="計算失敗", fg=self.RED))
        finally:
            self.root.after(0, lambda: self._set_buttons(True))

    # ══════════════════════════════════════════════
    #  Helpers
    # ══════════════════════════════════════════════
    def _set_buttons(self, enabled):
        state = tk.NORMAL if enabled else tk.DISABLED
        self.update_btn.config(state=state)
        self.backtest_btn.config(state=state)
        self.trade_btn.config(state=state)

    # ══════════════════════════════════════════════
    #  Chart Drawing (Apple minimal style)
    # ══════════════════════════════════════════════
    # 中文字體（全域）
    CJK_FONT = 'Microsoft JhengHei'

    def _apple_style_ax(self, ax):
        """Apply Apple-like minimal style to a matplotlib axes."""
        ax.set_facecolor(self.CARD)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.spines['bottom'].set_color(self.BORDER)
        ax.tick_params(axis='both', which='both',
                       colors=self.FG2, labelsize=8, length=0)
        ax.xaxis.label.set_color(self.FG2)
        ax.yaxis.label.set_color(self.FG2)
        for label in ax.get_yticklabels():
            label.set_fontfamily(self.CJK_FONT)
        for label in ax.get_xticklabels():
            label.set_fontfamily(self.CJK_FONT)

    def _draw_weight_chart(self, weight_data):
        """Draw ATR weight horizontal bar chart."""
        self.weight_fig.clear()

        if not weight_data or not weight_data.get('tickers'):
            ax = self.weight_fig.add_subplot(111)
            self._apple_style_ax(ax)
            ax.text(0.5, 0.5, '尚無資料', ha='center', va='center',
                    fontsize=11, color=self.FG2, fontfamily=self.CJK_FONT,
                    transform=ax.transAxes)
            ax.set_xlim(0, 1)
            ax.set_ylim(0, 1)
            ax.set_xticks([])
            ax.set_yticks([])
            self.weight_fig.tight_layout(pad=0.5)
            self.weight_canvas.draw_idle()
            return

        tickers = weight_data['tickers']
        target_w = weight_data['target_weights']
        current_w = weight_data['current_weights']
        colors = weight_data['colors']

        n = len(tickers)
        ax = self.weight_fig.add_subplot(111)
        self._apple_style_ax(ax)

        step = 0.6                    # 每組間距（原本 1.0）
        y_pos = [i * step for i in range(n)]
        bar_h = 0.22

        # 當前權重（灰色）
        ax.barh([y + bar_h / 2 for y in y_pos], current_w,
                height=bar_h, color='#d5ccbe', label='當前',
                edgecolor='none', zorder=2)
        # 目標權重（彩色）
        ax.barh([y - bar_h / 2 for y in y_pos], target_w,
                height=bar_h, color=colors, label='目標',
                edgecolor='none', zorder=2)

        # 數值標籤（粗體）
        for i, (tw, cw) in enumerate(zip(target_w, current_w)):
            y = y_pos[i]
            ax.text(tw + 0.3, y - bar_h / 2, f'{tw:.1f}%',
                    va='center', ha='left', fontsize=7,
                    fontweight='bold',
                    color=self.FG2, fontfamily=self.CJK_FONT)
            if cw > 0:
                ax.text(cw + 0.3, y + bar_h / 2, f'{cw:.1f}%',
                        va='center', ha='left', fontsize=7,
                        fontweight='bold',
                        color='#b8a88e', fontfamily=self.CJK_FONT)

        ax.set_yticks(y_pos)
        ax.set_yticklabels(tickers, fontsize=8, fontweight='bold',
                           fontfamily=self.CJK_FONT)
        ax.invert_yaxis()
        ax.set_xlabel('權重 %', fontsize=7, fontfamily=self.CJK_FONT)
        ax.xaxis.grid(True, color=self.BORDER, linewidth=0.5, alpha=0.7)
        ax.set_axisbelow(True)
        ax.set_ylim(y_pos[-1] + step * 0.5, -step * 0.5)

        # 圖例
        from matplotlib.patches import Patch
        legend_elements = [
            Patch(facecolor=self.GREEN, label='目標'),
            Patch(facecolor='#d5ccbe', label='現有'),
        ]
        ax.legend(handles=legend_elements, loc='lower right',
                  fontsize=7, frameon=False,
                  prop={'family': self.CJK_FONT})

        self.weight_fig.subplots_adjust(left=0.08, right=0.92,
                                        top=0.97, bottom=0.15)
        self.weight_canvas.draw_idle()

    def _draw_slope_chart(self, slope_data):
        """Draw horizontal bar chart for market ranking (top 20)."""
        self.slope_fig.clear()

        if not slope_data or not slope_data.get('tickers'):
            ax = self.slope_fig.add_subplot(111)
            self._apple_style_ax(ax)
            ax.text(0.5, 0.5, '尚無資料', ha='center', va='center',
                    fontsize=11, color=self.FG2, fontfamily=self.CJK_FONT,
                    transform=ax.transAxes)
            ax.set_xlim(0, 1)
            ax.set_ylim(0, 1)
            ax.set_xticks([])
            ax.set_yticks([])
            self.slope_fig.tight_layout(pad=0.5)
            self.slope_canvas.draw_idle()
            return

        tickers = slope_data['tickers']
        slopes = slope_data['slopes']
        colors = slope_data['colors']
        labels = slope_data.get('labels', [''] * len(tickers))

        n = len(tickers)
        ax = self.slope_fig.add_subplot(111)
        self._apple_style_ax(ax)

        y_pos = range(n)
        bar_h = 0.55
        bars = ax.barh(y_pos, slopes, height=bar_h,
                       color=colors, edgecolor='none', zorder=2)

        # Y 軸標籤：純 ticker（不用 #1 #2）
        ax.set_yticks(y_pos)
        ax.set_yticklabels(tickers, fontsize=9, fontweight='bold',
                           fontfamily=self.CJK_FONT)
        ax.invert_yaxis()

        # 數值標籤（粗體）+ 狀態標記
        max_slope = max(slopes) if slopes else 1
        for i, s in enumerate(slopes):
            # 數值
            ax.text(s + max_slope * 0.015, i, f'{s:.2f}',
                    va='center', ha='left', fontsize=8,
                    fontweight='bold',
                    color=self.FG2, fontfamily=self.CJK_FONT)
            # 持有/新買 標記
            if labels[i]:
                ax.text(s + max_slope * 0.09, i, labels[i],
                        va='center', ha='left', fontsize=7.5,
                        color=colors[i], fontfamily=self.CJK_FONT,
                        fontweight='bold')

        ax.set_xlabel('股票動能強度', fontsize=9, fontfamily=self.CJK_FONT)
        ax.xaxis.grid(True, color=self.BORDER, linewidth=0.5, alpha=0.5)
        ax.set_axisbelow(True)

        # 增加右邊空間給標籤
        ax.set_xlim(left=0, right=max_slope * 1.2)

        self.slope_fig.tight_layout(pad=1.0)
        self.slope_canvas.draw_idle()

    # ══════════════════════════════════════════════
    #  Display Results
    # ══════════════════════════════════════════════
    def _display_results(self, data):
        """
        data: dict with keys:
          - text_lines: [(tag, text), ...]
          - weight_data: dict for ATR weight chart
          - slope_data: dict for Slope chart
        """
        # 文字摘要
        tagged = data.get('text_lines', [])
        self.results_text.config(state=tk.NORMAL)
        self.results_text.delete("1.0", tk.END)
        for i, (tag, text) in enumerate(tagged):
            if i > 0:
                self.results_text.insert(tk.END, "\n")
            self.results_text.insert(tk.END, text, tag)
        self.results_text.config(state=tk.DISABLED)

        # ATR 權重圖
        self._draw_weight_chart(data.get('weight_data'))

        # Slope 圖
        self._draw_slope_chart(data.get('slope_data'))


if __name__ == "__main__":
    # 啟用 Windows 高 DPI 感知，避免模糊縮放
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)   # Per-Monitor V2
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    root = tk.Tk()
    # 讓 Tk 也知道當前 DPI
    try:
        root.tk.call('tk', 'scaling', root.winfo_fpixels('1i') / 72)
    except Exception:
        pass

    app = UnifiedDashboard(root)
    root.mainloop()
