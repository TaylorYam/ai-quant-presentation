import pandas as pd
import numpy as np
import scipy.stats as stats
import config
import os
import utils


class SelectionEngine:
    def __init__(self, data_cache=None):
        self.data_cache = data_cache if data_cache else {}
        self.scan_cache = {}  # Cache for scan_market results {(date, lookback): sorted_list}
        self.metrics_cache = {}  # Cache for calculation results {(ticker, date, lookback): stats_dict}
        self.constituents_df = self._load_constituents()
        self._all_tickers_loaded = False
        
    def _load_constituents(self):
        """讀取 S&P 500 成分股歷史資料 (Excel)"""
        path = os.path.join(config.DATA_DIR, config.CONST_FILE)
        print(f"Loading constituents from {path}...")
        try:
            # 使用 openpyxl engine
            xl = pd.ExcelFile(path, engine='openpyxl')
            sheets = xl.sheet_names
            valid_sheets = [s for s in sheets if s.isdigit()]
            all_data = []
            for sheet in valid_sheets:
                df = pd.read_excel(xl, sheet_name=sheet)
                if 'Date' in df.columns:
                    df['Date'] = pd.to_datetime(df['Date'])
                    df.set_index('Date', inplace=True)
                    all_data.append(df)
            if not all_data:
                print("No valid daily constituent data found in Excel.")
                return pd.DataFrame()
            full_df = pd.concat(all_data)
            full_df.sort_index(inplace=True)
            return full_df
        except Exception as e:
            print(f"Error loading constituents: {e}")
            return pd.DataFrame()

    def preload_all_data(self, start_date=None, end_date=None):
        """
        預載入所有股票資料到記憶體 (大幅減少 I/O)
        應在回測開始前呼叫
        """
        if self._all_tickers_loaded:
            return  # 已經載入過
            
        if self.constituents_df.empty:
            return
            
        # 收集所有可能的 ticker
        all_tickers = set()
        
        # 如果有指定時間範圍，只收集該範圍內的 ticker
        if start_date and end_date:
            start_dt = pd.to_datetime(start_date)
            end_dt = pd.to_datetime(end_date)
            mask = (self.constituents_df.index >= start_dt) & (self.constituents_df.index <= end_dt)
            subset = self.constituents_df[mask]
        else:
            subset = self.constituents_df
            
        for idx in subset.index:
            row = subset.loc[idx]
            tickers = [t for t in row.values if isinstance(t, str)]
            all_tickers.update(t.replace('.txt', '').strip() for t in tickers)
        
        print(f"Preloading {len(all_tickers)} tickers into memory...")
        loaded = 0
        for ticker in all_tickers:
            if self._get_ticker_data(ticker) is not None:
                loaded += 1
        print(f"Loaded {loaded} tickers successfully.")
        self._all_tickers_loaded = True

    def get_constituents(self, date):
        """獲取特定日期的成分股列表 (自動過濾黑名單)"""
        if self.constituents_df.empty:
            return []
        if date not in self.constituents_df.index:
            try:
                loc = self.constituents_df.index.get_indexer([date], method='pad')[0]
                if loc == -1:
                    return []
                target_date = self.constituents_df.index[loc]
            except:
                return []
        else:
            target_date = date
        row = self.constituents_df.loc[target_date]
        tickers = [t for t in row.values if isinstance(t, str)]
        tickers = [t.replace('.txt', '').strip() for t in tickers]
        
        # Filter out blacklisted stocks (check config_final first, then config)
        blacklist = getattr(config, 'BLACKLIST', None)
        if blacklist is None:
            try:
                import config_final
                blacklist = getattr(config_final, 'BLACKLIST', [])
            except ImportError:
                blacklist = []
        
        if blacklist:
            for bl_date_str, bl_ticker in blacklist:
                bl_date = pd.to_datetime(bl_date_str)
                bl_ticker = bl_ticker.upper()
                if date > bl_date and bl_ticker in tickers:
                    tickers.remove(bl_ticker)
        
        return tickers

    # 預計算的 EMA 週期列表 (涵蓋優化常用範圍 20-60)
    PRECOMPUTED_EMA_PERIODS = [20, 30, 40, 50, 60]
    
    def _get_ticker_data(self, ticker):
        """從 cache 或 disk 讀取股票資料，並預計算多個常用 EMA 週期"""
        if ticker in self.data_cache:
            return self.data_cache[ticker]
        
        p = os.path.join(config.DATA_DIR, f"{ticker}.txt")
        if not os.path.exists(p):
            p = os.path.join(config.DATA_DIR, f"{ticker}.csv")
            
        if os.path.exists(p):
            df = utils.load_data(p)
            if df is not None and not df.empty:
                # === 預計算多個常用 EMA 週期 (20, 30, 40, 50, 60) ===
                price_col = 'Adj Close' if 'Adj Close' in df.columns else 'Close'
                for period in self.PRECOMPUTED_EMA_PERIODS:
                    df[f'_EMA{period}'] = df[price_col].ewm(span=period, adjust=False).mean()
                self.data_cache[ticker] = df
                return df
        return None

    def calculate_metrics(self, ticker, current_date, lookback=60):
        """
        計算股票指標 (精簡版 - 只計算實際使用的指標)
        實際使用的指標: adj_slope, max_gap, price, ema50
        """
        # 0. Check Metrics Cache (包含 EXIT_EMA 和 ATR_PERIOD 支援優化)
        exit_ema_period = getattr(config, 'EXIT_EMA', 50)
        
        # 讀取 ATR_PERIOD (需與後續計算一致)
        atr_period = getattr(config, 'ATR_PERIOD', None)
        if atr_period is None:
            try:
                import config_final
                atr_period = getattr(config_final, 'ATR_PERIOD', 14)
            except:
                atr_period = 14
        
        cache_key = (ticker, current_date, lookback, exit_ema_period, atr_period)
        if cache_key in self.metrics_cache:
            return self.metrics_cache[cache_key]

        # 1. 獲取資料
        df = self._get_ticker_data(ticker)
        if df is None or df.empty:
            return None
        
        mask = df.index <= current_date
        hist_data = df[mask].tail(lookback + exit_ema_period)  # 動態計算所需額外天數
        
        if len(hist_data) < lookback:
            return None
            
        target_chunk = hist_data.iloc[-lookback:]
        
        # 優先使用 'Adj Close' 計算動能
        price_col_for_trend = 'Adj Close' if 'Adj Close' in hist_data.columns else 'Close'
        trend_series = target_chunk[price_col_for_trend].values
        closes = target_chunk['Close'].values
        
        # 2. Adjusted Slope (排序用)
        try:
            y_log = np.log(trend_series)
            x_axis = np.arange(len(y_log))
            slope, intercept, r_value, p_value, std_err = stats.linregress(x_axis, y_log)
            adj_slope = ((1 + slope) ** lookback - 1) * (r_value ** 2)
        except:
            adj_slope = -999

        # 3. Max Gap (過濾用 - 向量化計算)
        opens = target_chunk['Open'].values
        prev_closes = np.roll(closes, 1)
        prev_closes[0] = opens[0]
        gaps = np.abs((opens - prev_closes) / prev_closes)
        max_gap = np.max(gaps[1:])
        
        # 4. EXIT_EMA - 優先使用預計算值，否則動態計算
        ema_col = f'_EMA{exit_ema_period}'
        if ema_col in hist_data.columns:
            # 使用預計算的值 (O(1) 查詢)
            exit_ema = hist_data[ema_col].iloc[-1] if len(hist_data) > 0 else 0.0
        else:
            # 非常用週期，動態計算
            price_col = 'Adj Close' if 'Adj Close' in hist_data.columns else 'Close'
            ema_series = hist_data[price_col].ewm(span=exit_ema_period, adjust=False).mean()
            exit_ema = ema_series.iloc[-1] if len(ema_series) > 0 else 0.0
            
        current_price = closes[-1]
        
        # 5. ATR 計算 (Average True Range) - V3 風險評估用
        high = target_chunk['High'].values
        low = target_chunk['Low'].values
        prev_close = np.roll(closes, 1)
        prev_close[0] = closes[0]  # 處理第一個元素
        
        # True Range = max(High-Low, |High-PrevClose|, |Low-PrevClose|)
        tr = np.maximum(high - low, 
                        np.maximum(np.abs(high - prev_close), 
                                   np.abs(low - prev_close)))
        
        # 使用已計算的 atr_period (from cache_key logic)
        atr = np.mean(tr[-atr_period:]) if len(tr) >= atr_period else np.mean(tr)
        atr_pct = (atr / current_price) if current_price > 0 else 0  # 標準化為百分比
        
        # 精簡的結果 - 只包含實際使用的欄位
        result_dict = {
            'ticker': ticker,
            'adj_slope': adj_slope,
            'max_gap': max_gap,
            'price': current_price,
            'exit_ema': exit_ema,
            'atr': atr,           # V3: 原始 ATR 值
            'atr_pct': atr_pct,   # V3: 標準化 ATR (用於跨股票比較)
        }
        
        # Save to Cache
        self.metrics_cache[cache_key] = result_dict
        
        return result_dict

    def _compute_residuals(self, tickers, date, spy_df, lookback=60):
        """
        計算每檔股票相對於 SPY 的迴歸殘差
        回傳 {ticker: pd.Series(residuals)} 字典
        """
        # SPY 日報酬率
        spy_close = spy_df['Close']
        spy_mask = spy_close.index <= date
        spy_hist = spy_close[spy_mask].tail(lookback + 1)
        spy_returns = spy_hist.pct_change().dropna()
        
        if len(spy_returns) < int(lookback * 0.8):
            return {}
        
        residuals_dict = {}
        min_data = int(lookback * 0.8)
        
        for ticker in tickers:
            df = self._get_ticker_data(ticker)
            if df is None or df.empty:
                continue
            
            price_col = 'Adj Close' if 'Adj Close' in df.columns else 'Close'
            stock_mask = df.index <= date
            stock_hist = df[stock_mask][price_col].tail(lookback + 1)
            stock_returns = stock_hist.pct_change().dropna()
            
            # 取共同交易日
            common_idx = stock_returns.index.intersection(spy_returns.index)
            if len(common_idx) < min_data:
                continue
            
            sr = stock_returns.loc[common_idx].values
            mr = spy_returns.loc[common_idx].values
            
            # 迴歸: R_stock = alpha + beta * R_spy + epsilon
            slope, intercept, _, _, _ = stats.linregress(mr, sr)
            residuals = sr - (intercept + slope * mr)
            residuals_dict[ticker] = pd.Series(residuals, index=common_idx)
        
        return residuals_dict

    def filter_by_residual_correlation(self, ranked_candidates, date, spy_df,
                                        threshold, lookback, max_candidates,
                                        needed, existing_tickers=None):
        """
        殘差相關性過濾 - 包廂邏輯 (Box Seat Logic)
        ranked_candidates: 已排序的候選股清單 [{ticker, adj_slope, ...}, ...]
        existing_tickers: 現有持股 ticker 列表，需一併參與相關性檢查
        回傳: 篩選後的新買入 ticker 列表
        """
        if existing_tickers is None:
            existing_tickers = []
        
        # 取排名前 max_candidates 檔候選
        candidates = ranked_candidates[:max_candidates]
        candidate_tickers = [x['ticker'] for x in candidates]
        
        # 所有需要計算殘差的 ticker (候選 + 已持有)
        all_tickers = list(set(candidate_tickers + list(existing_tickers)))
        residuals_dict = self._compute_residuals(all_tickers, date, spy_df, lookback)
        
        # 已選集合 (先放入已持有的股票)
        selected = [t for t in existing_tickers if t in residuals_dict]
        result = []  # 只記錄新買入的
        
        for item in candidates:
            ticker = item['ticker']
            if len(result) >= needed:
                break
            
            if ticker not in residuals_dict:
                continue
            
            # 檢查與所有已選股的殘差相關係數
            passed = True
            for sel_ticker in selected:
                if sel_ticker not in residuals_dict:
                    continue
                res_a = residuals_dict[ticker]
                res_b = residuals_dict[sel_ticker]
                common = res_a.index.intersection(res_b.index)
                if len(common) < 20:
                    continue
                corr = np.corrcoef(res_a.loc[common].values, res_b.loc[common].values)[0, 1]
                if abs(corr) >= threshold:
                    passed = False
                    break
            
            if passed:
                selected.append(ticker)
                result.append(ticker)
        
        return result

    def scan_market(self, date, lookback=None):
        """掃描市場並排名 (使用快取)"""
        # Use provided lookback or default to LOOKBACK_ENTRY
        lb = lookback if lookback is not None else config.LOOKBACK
        
        # Optimization: Check cache (key includes lookback now)
        cache_key = (date, lb)
        if cache_key in self.scan_cache:
            return self.scan_cache[cache_key]

        tickers = self.get_constituents(date)
        if not tickers:
            return []
            
        results = []
        for t in tickers:
            metrics = self.calculate_metrics(t, date, lb)
            if metrics:
                results.append(metrics)
        
        # Filter Logic: 跳空缺口 > threshold 的股票不納入排名
        filtered = [r for r in results if r['max_gap'] <= config.SKIP_MAX_GAP_PCT]
        
        # Sort by Adjusted Slope (High to Low)
        sorted_list = sorted(filtered, key=lambda x: x['adj_slope'], reverse=True)
        
        # Save to cache
        self.scan_cache[cache_key] = sorted_list
        
        return sorted_list
