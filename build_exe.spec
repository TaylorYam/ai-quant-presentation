# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec - 全天候SP500動能策略 Dashboard
所有模組直接 import，不用 subprocess
"""
import os

block_cipher = None
ROOT = SPECPATH

import importlib
BACKTESTING_DIR = os.path.dirname(importlib.import_module('backtesting').__file__)

a = Analysis(
    [os.path.join(ROOT, 'dashboard_final.py')],
    pathex=[ROOT],
    binaries=[],
    datas=[
        # config 需要被讀寫（使用者可調參數）
        (os.path.join(ROOT, 'config_final.py'), '.'),
        (os.path.join(ROOT, 'config.py'), '.'),
        # backtesting 套件的靜態資源
        (os.path.join(BACKTESTING_DIR, 'autoscale_cb.js'), 'backtesting'),
    ],
    hiddenimports=[
        'scipy.stats',
        'numpy',
        'pandas',
        'matplotlib',
        'matplotlib.backends.backend_tkagg',
        'yfinance',
        'backtesting',
        # 直接 import 的模組
        'update_data',
        'run_strategy_final',
        'portfolio_backtester_final',
        'report_generator_final',
        'market_regime',
        'selection',
        'utils',
        'data_updater',
        'config_final',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SP500_Strategy',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='SP500_Strategy',
)
