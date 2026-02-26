import os
import sys

# 設定資料來源路徑（支援 PyInstaller 打包）
if getattr(sys, 'frozen', False):
    _BASE = os.path.dirname(sys.executable)
else:
    _BASE = os.path.dirname(__file__)
DATA_DIR = os.path.join(_BASE, 'data')

# 回測設定
INITIAL_CASH = 1000000  # 初始資金 10萬
COMMISSION = 0.01    # 交易手續費 0.5%

# 策略設定
TARGET_HOLDINGS = 4      # 目標持股數
REBALANCE_WEEKS = 1      # 調倉頻率 (週為單位)
REBALANCE_WEEKDAY = 2    # 調倉日 (0=週一, 1=週二, ..., 3=週四, 4=週五)
LOOKBACK = 90             # 排名計算週期（進場/出場共用）
LOOKBACK_ENTRY = LOOKBACK # 向下相容
LOOKBACK_EXIT  = LOOKBACK # 向下相容
EXIT_EMA = 50            # 出場EMA週期 (跌破該EMA賣出)

# Strategy Settings
SKIP_MAX_GAP_PCT = 0.20  # 跳空缺口 > 20% 不買
GAP_EXIT_PCT = 0.5       # 持倉跳空缺口 > 50% 隔日開盤出場
SELL_RANK_THRESHOLD = 20 # 排名掉出前 20 賣出
STOP_LOSS_PCT = 0.2      # 20% 停損
MAX_ADJ_SLOPE = 1.5      # [NEW] Adj Slope 上限 (過熱回檔保護)

# 基準與抄底標的
BENCHMARK_TICKER = 'SPY'
DIP_BUY_TICKER = 'SSO'

# SP500 成分股歷史檔案
CONST_FILE = 'sp500_constituents_daily_2015_2026.xlsx'

# ======================================
# Final Version: 即時持股模擬模式
# ======================================
# START_DATE: 回測起始日
# END_DATE: None = 自動偵測最新可用日期
START_DATE = '2020-01-01'
END_DATE = None  # 自動偵測最新日期

# LIVE_MODE: 結束時不清倉，保留持股狀態
LIVE_MODE = True

# ATR 風險再平衡設定
ATR_PERIOD = 20               # ATR 計算週期
REBALANCE_THRESHOLD = 0.03    # 風險再平衡閾值 (超重/低配 3% 以上觸發)
MIN_BUY_AMOUNT_PCT = 0.03     # 最小買入金額 (總權益的 3%) - 低於此金額不執行買入

# 殘差相關性過濾設定 (Residual Correlation Filter)
CORR_FILTER_ENABLED = True      # 是否啟用殘差相關性過濾
CORR_THRESHOLD = 0.6            # 殘差相關係數門檻（超過此值視為同質）
CORR_LOOKBACK = 60              # 計算殘差的回看天數
CORR_CANDIDATE_COUNT = 20       # 進入相關性過濾的候選股數量

# ======================================
# BLACKLIST: 特定日期後不持有的股票
# Format: [('YYYY-MM-DD', 'TICKER'), ...]
# 在該日期後，不買入該股票，若已持有則強制賣出
# ======================================
BLACKLIST = [
    ('2025-12-31', 'WBD'),  # 2025/12/31 後不持有 WBD
]

