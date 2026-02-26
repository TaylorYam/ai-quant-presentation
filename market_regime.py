import pandas as pd
import numpy as np

class MarketRegime:
    def __init__(self, spy_df, sso_df):
        self.spy = spy_df.sort_index()
        self.sso = sso_df.sort_index()
        
        # 計算 SPY 指標
        self._calculate_indicators()
    
    def _calculate_indicators(self):
        # 1. 200MA
        self.spy['MA200'] = self.spy['Close'].rolling(window=200).mean()
        
        # 2. ATH (All Time High)
        # 注意：這裡使用擴展窗口最大值 (accumulated max)
        self.spy['ATH'] = self.spy['Close'].cummax()
        
        # 計算目前價格距離 ATH 的跌幅 (Drawdown from ATH)
        self.spy['DD_ATH'] = (self.spy['Close'] - self.spy['ATH']) / self.spy['ATH']

    def get_state(self, date):
        """
        獲取指定日期的市場狀態
        """
        # 找到該日期或之前的最近一個交易日
        if date not in self.spy.index:
            try:
                # 使用 asof 尋找最近的日期 (index 必須是 sorted)
                idx = self.spy.index.get_indexer([date], method='pad')[0]
                if idx == -1: return None # date is before start
                row = self.spy.iloc[idx]
            except:
                return None
        else:
            row = self.spy.loc[date]
            
        return {
            'SPY_Close': row['Close'],
            'SPY_MA200': row['MA200'],
            'SPY_ATH': row['ATH'],
            'SPY_DD': row['DD_ATH']
        }

    def is_bull_market(self, date):
        """
        判斷是否為牛市：
        條件：本周 (date) > 200MA  AND  上一周 (date - 7) > 200MA
        """
        state_now = self.get_state(date)
        if not state_now: return False
        bull_now = state_now['SPY_Close'] > state_now['SPY_MA200']
        
        # Check 7 days ago
        prev_date = date - pd.Timedelta(days=7)
        state_prev = self.get_state(prev_date)
        if not state_prev: return False # 如果數據不足，保守起見回傳 False
        bull_prev = state_prev['SPY_Close'] > state_prev['SPY_MA200']
        
        return bull_now and bull_prev
