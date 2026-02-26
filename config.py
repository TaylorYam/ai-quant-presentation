import os
import sys

# 設定資料來源路徑（支援 PyInstaller 打包）
if getattr(sys, 'frozen', False):
    _BASE = os.path.dirname(sys.executable)
else:
    _BASE = os.path.dirname(__file__)
DATA_DIR = os.path.join(_BASE, 'data')

# 回測設定
INITIAL_CASH = 100000  # 初始資金 10萬
COMMISSION = 0.005    # 交易手續費 0.3%

# 策略設定
TARGET_HOLDINGS = 4      # 目標持股數
REBALANCE_WEEKS = 1      # 調倉頻率 (週為單位)
REBALANCE_WEEKDAY = 2    # 調倉日 (0=週一, 1=週二, ..., 3=週四, 4=週五)
LOOKBACK_ENTRY = 90      # 進場排名計算週期
LOOKBACK_EXIT = 90       # 出場排名計算週期
EXIT_EMA = 50            # 出場EMA週期 (跌破該EMA賣出)
# Strategy Settings
SKIP_MAX_GAP_PCT = 0.20 # 跳空缺口 > 15% 不買
GAP_EXIT_PCT = 0.5      # 持倉跳空缺口 > 15% 隔日開盤出場
SELL_RANK_THRESHOLD = 20 # 排名掉出前 20 賣出
STOP_LOSS_PCT = 0.2     # 10% 停損

# 基準與抄底標的
BENCHMARK_TICKER = 'SPY'
DIP_BUY_TICKER = 'SSO'

# SP500 成分股歷史檔案
CONST_FILE = 'sp500_constituents_daily_2015_2026.xlsx'

# 預設回測時間範圍 (None 表示全部)
START_DATE = '2020-01-01'
END_DATE = '2025-12-31'
