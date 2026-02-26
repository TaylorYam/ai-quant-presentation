"""
Portfolio Backtester V3
帶有 ATR 風險加權倉位和每週風險再平衡功能
"""
from backtesting import Strategy
import pandas as pd
import config_final as config
from selection import SelectionEngine
from market_regime import MarketRegime
import utils
import os

class PortfolioBacktesterFinal:
    def __init__(self, start_date, end_date, initial_capital=100000, compounding=False, report_suffix="", selector=None, spy_df=None, sso_df=None, write_reports=True):
        self.initial_capital = initial_capital
        self.cash = initial_capital
        self.holdings = {}      # {ticker: quantity}
        self.avg_costs = {}     # {ticker: avg_cost_per_share}
        self.target_weights = {}  # V3: {ticker: target_weight} 目標權重追蹤
        self.history = []       # 記錄每日權益
        self.trades = []        # 記錄交易
        
        self.compounding = compounding
        self.report_suffix = report_suffix
        self.selector_instance = selector
        
        # Optimization Flags
        self.preloaded_spy = spy_df
        self.preloaded_sso = sso_df
        self.write_reports = write_reports
        
        # Dip Buying State
        self.dip_state = {0.15: False, 0.20: False, 0.25: False}
        
        # [NEW] Bull Market Confirmation - 需連續兩周 SPY > 200MA
        self.bull_weeks_counter = 0  # 連續牛市週數計數器
        
        # [NEW] Rebalance Day Snapshots - 記錄每個調倉日的持倉和排名
        self.rebalance_snapshots = []

        
        # 時間範圍 (需要在 prep() 之前設定，因為 preload 會用到)
        self._start_date_raw = start_date
        self._end_date_raw = end_date
        
        # 初始化模組
        self.prep()
        
    def prep(self):
        # 1. 載入市場 Data
        if self.preloaded_spy is not None:
            self.spy_df = self.preloaded_spy
        else:
            spy_path = os.path.join(config.DATA_DIR, 'SPY.csv')
            self.spy_df = utils.load_benchmark_data(spy_path)
            
        if self.preloaded_sso is not None:
            self.sso_df = self.preloaded_sso
        else:
            sso_path = os.path.join(config.DATA_DIR, 'SSO.csv')
            self.sso_df = utils.load_benchmark_data(sso_path)
        
        self.market_regime = MarketRegime(self.spy_df, self.sso_df)
        
        # 設定時間範圍 (需要在 market_regime 載入後才能取得 spy index)
        self.start_date = pd.to_datetime(self._start_date_raw) if self._start_date_raw else self.market_regime.spy.index[200]
        self.end_date = pd.to_datetime(self._end_date_raw) if self._end_date_raw else self.market_regime.spy.index[-1]
        
        if self.selector_instance:
            self.selector = self.selector_instance
        else:
            self.selector = SelectionEngine()
        
        # === 預載入所有股票資料到記憶體 (Performance Optimization) ===
        if self.write_reports:
            print("Preloading stock data...")
        self.selector.preload_all_data(self.start_date, self.end_date)
        
        # 建立全局交易日曆 (以 SPY 為準)
        self.calendar = self.spy_df.index
        
    def run(self):
        mode_str = "Compound" if self.compounding else "Simple"
        if self.write_reports:
            print(f"Starting {mode_str} Portfolio Backtest Final from {self.start_date.date()} to {self.end_date.date()}")
            print(f"Strategy: ATR-Weighted Position Sizing + Risk Rebalancing")
        
        # 建立交易日曆
        trading_days = self.calendar[(self.calendar >= self.start_date) & (self.calendar <= self.end_date)]
        
        # 追蹤 SSO 觸發狀態
        self.dip_state = {0.15: False, 0.20: False, 0.25: False}
        
        total_days = len(trading_days)
        for i, date in enumerate(trading_days):
            # Progress Reporting
            if self.write_reports and (i % 500 == 0 or i == total_days - 1):
                progress = int((i + 1) / total_days * 100)
                print(f"[PROGRESS] {progress}", flush=True)
            
            # 1. Stop Loss Check (Prior to updating equity)
            if i > 0:
                prev_date = trading_days[i-1]
                self._check_stop_loss(date, prev_date)
                self._check_gap_exit(date, prev_date)

            # 每日更新淨值
            self._update_equity(date)
            
            # --- Daily Checks (Bear Flow & SSO) ---
            self._check_bear_sso_logic(date)
            
            # --- 統一換股/再平衡日 (使用 REBALANCE_WEEKDAY) ---
            iso_week = date.isocalendar()[1]
            is_rebalance_day = (date.weekday() == config.REBALANCE_WEEKDAY)
            is_rotation_week = (iso_week % config.REBALANCE_WEEKS == 0)
            
            if is_rebalance_day:
                prev_idx = i - 1
                signal_date = trading_days[prev_idx] if prev_idx >= 0 else date
                
                # === V3 核心：統一先賣後買流程 ===
                self._unified_rebalance(date, signal_date, is_rotation_week)
            
        # End of Backtest: LIVE_MODE keeps holdings, otherwise close all
        live_mode = getattr(config, 'LIVE_MODE', False)
        if not live_mode:
            self._force_close_all(self.end_date)
        else:
            self._update_equity(self.end_date)
            if self.write_reports:
                print("[LIVE MODE] Keeping holdings open for tracking")
        
        if self.write_reports:
            self._generate_report()

    def _calculate_atr_weights(self, candidates_with_atr):
        """
        V3: 計算基於 ATR 的反比例權重
        candidates_with_atr: [{'ticker': ..., 'atr_pct': ...}, ...]
        
        Weight_i = (1 / ATR_i) / Σ(1 / ATR_j)
        ATR 越低的股票，權重越高（波動小的股票配置更多資金）
        """
        # 過濾掉 atr_pct <= 0 的股票
        valid_candidates = [c for c in candidates_with_atr if c.get('atr_pct', 0) > 0]
        
        if not valid_candidates:
            return {}
        
        # 計算反比例 ATR 總和
        inv_atr_sum = sum(1 / c['atr_pct'] for c in valid_candidates)
        
        weights = {}
        for c in valid_candidates:
            weights[c['ticker']] = (1 / c['atr_pct']) / inv_atr_sum
        
        return weights

    def _unified_rebalance(self, date, signal_date, is_rotation_week):
        """
        統一再平衡流程 (已修正 - 用完整組合計算 ATR 權重):
        1. 輪動賣出 - 賣出不符合條件的股票 (排名掉出/跌破EMA)
        2. 確定完整組合 - 找出買入候選，計算完整 ATR 權重
        3. 超重賣出 - 用完整組合權重判斷 (當前權重 - 目標權重 >= 3%)
        4. 買入新股票補足至 TARGET_HOLDINGS
        5. 剩餘現金按比例分配給所有低配股票
        
        [CRITICAL] 只有在牛市確認 (連續兩周 SPY > 200MA) 後才執行個股交易
        """
        bull_confirmed = (self.bull_weeks_counter >= 2)
        
        if not bull_confirmed:
            current_stocks = [t for t in self.holdings if t != config.DIP_BUY_TICKER]
            if current_stocks:
                self._close_positions(date, target_type='STOCK')
                self.target_weights.clear()
            return
        
        # === Step 1: 輪動賣出 (僅在輪動週執行) ===
        rotation_sell_tickers = set()
        if is_rotation_week:
            rotation_sells = self._get_rotation_sells(date, signal_date)
            for ticker, qty, price, reason in rotation_sells:
                rotation_sell_tickers.add(ticker)
                total_equity = self._get_total_equity(date)
                current_qty = self.holdings.get(ticker, 0)
                weight_before = (price * current_qty / total_equity * 100) if total_equity > 0 else 0
                if ticker in self.holdings and self.holdings[ticker] >= qty:
                    self._sell(ticker, date, price, qty, reason, weight_before=weight_before)
                    if ticker in self.target_weights:
                        del self.target_weights[ticker]
        
        # === Step 2: 確定完整組合 + 計算完整 ATR 權重 ===
        # 先找到買入候選，再用 (保留持股 + 候選) 一起算 ATR 權重
        buy_candidates_info = []
        full_weights = {}
        if is_rotation_week:
            buy_candidates_info, full_weights = self._get_buy_candidates(date, signal_date)
        
        if not full_weights:
            # 非輪動週或無候選：用當前持股算權重
            current_stocks = [t for t in self.holdings if t != config.DIP_BUY_TICKER and t != 'SPY'
                              and t not in rotation_sell_tickers]
            if current_stocks:
                holdings_with_atr = []
                for ticker in current_stocks:
                    metrics = self.selector.calculate_metrics(ticker, date, config.LOOKBACK)
                    if metrics and metrics.get('atr_pct', 0) > 0:
                        holdings_with_atr.append(metrics)
                if holdings_with_atr:
                    full_weights = self._calculate_atr_weights(holdings_with_atr)
        
        # Update target weights with full portfolio weights
        if full_weights:
            self.target_weights = full_weights.copy()
        
        # === Step 3: 超重賣出 (每週執行，使用完整組合權重) ===
        overweight_sells = self._get_overweight_sells(
            date, exclude_tickers=rotation_sell_tickers,
            full_portfolio_weights=full_weights if full_weights else None)
        for ticker, qty, price, reason, weight_before in overweight_sells:
            if ticker in self.holdings and self.holdings[ticker] >= qty:
                self._sell(ticker, date, price, qty, reason, weight_before=weight_before)
        
        # === Step 4: 買入新股票 (僅在輪動週執行) ===
        newly_bought_tickers = set()
        if is_rotation_week and buy_candidates_info:
            if self.compounding:
                curr_equity = self._get_total_equity(date)
            else:
                curr_equity = self.initial_capital
            
            for ticker, exec_price in buy_candidates_info:
                full_weight = full_weights.get(ticker, 0)
                if full_weight <= 0 or exec_price <= 0:
                    continue
                alloc_amount = curr_equity * full_weight
                buy_amount = min(alloc_amount, self.cash)
                qty = int(buy_amount / (exec_price * (1 + config.COMMISSION)))
                if qty > 0 and self.cash >= exec_price * qty * (1 + config.COMMISSION):
                    reason = f"V3 ATR Buy (W:{full_weight*100:.1f}%)"
                    self._buy(ticker, date, exec_price, qty, reason, target_weight=full_weight)
                    newly_bought_tickers.add(ticker)
        
        # === Step 5: 剩餘現金按比例分配給所有低配股票 ===
        self._distribute_remaining_cash(date, exclude_tickers=newly_bought_tickers)
        
        # 記錄調倉日快照（僅在輪動週）
        if is_rotation_week:
            self._record_rebalance_snapshot(date, signal_date)

    
    def _get_overweight_sells(self, date, exclude_tickers=None, full_portfolio_weights=None):
        """
        獲取超重需要賣出的股票 (當前權重 - ATR目標權重 >= 3%)
        full_portfolio_weights: 如提供，使用此完整組合權重；否則用當前持股自行計算
        返回: [(ticker, qty, price, reason, weight_before)]
        """
        sells = []
        
        if exclude_tickers is None:
            exclude_tickers = set()
        
        current_stocks = [t for t in self.holdings if t != config.DIP_BUY_TICKER and t != 'SPY']
        current_stocks = [t for t in current_stocks if t not in exclude_tickers]
        
        if not current_stocks:
            return sells
        
        # 檢查數據完整性
        for ticker in current_stocks:
            check_price = self._get_price(ticker, date, use_open=True)
            if check_price <= 0:
                return sells
        
        total_equity = self._get_total_equity(date)
        if total_equity <= 0:
            return sells
        
        # 使用傳入的完整組合權重，或自行計算
        if full_portfolio_weights:
            dynamic_target_weights = full_portfolio_weights
        else:
            holdings_with_atr = []
            for ticker in current_stocks:
                metrics = self.selector.calculate_metrics(ticker, date, config.LOOKBACK)
                if metrics and metrics.get('atr_pct', 0) > 0:
                    holdings_with_atr.append(metrics)
            if not holdings_with_atr:
                return sells
            dynamic_target_weights = self._calculate_atr_weights(holdings_with_atr)
            self.target_weights = dynamic_target_weights.copy()
        
        # 檢查超重並賣出
        rebalance_threshold = config.REBALANCE_THRESHOLD
        if rebalance_threshold < 0.01:
            rebalance_threshold = 0.03
        
        for ticker in current_stocks:
            if ticker not in dynamic_target_weights:
                continue
            
            target_w = dynamic_target_weights[ticker]
            current_qty = self.holdings.get(ticker, 0)
            price = self._get_price(ticker, date, use_open=True)
            if price <= 0:
                continue
            
            current_value = price * current_qty
            current_w = current_value / total_equity if total_equity > 0 else 0
            deviation = current_w - target_w
            
            # 超重 >= 3% 才賣出
            if deviation >= rebalance_threshold:
                target_value = total_equity * target_w
                diff_value = current_value - target_value
                
                if diff_value >= price:
                    qty_to_sell = min(int(diff_value / price), current_qty)
                    if qty_to_sell > 0:
                        reason = f"V3 超重賣出 ({current_w*100:.1f}% → {target_w*100:.1f}%)"
                        weight_before = current_w * 100
                        sells.append((ticker, qty_to_sell, price, reason, weight_before))
        
        return sells
    
    def _distribute_remaining_cash(self, date, exclude_tickers=None):
        """
        將剩餘現金按比例分配給所有低配股票
        不設門檻，有多少現金就按 ATR 權重比例分配
        
        [FIX] exclude_tickers: 排除剛買入的股票，避免同一天重複購買
        """
        if exclude_tickers is None:
            exclude_tickers = set()
        
        current_stocks = [t for t in self.holdings if t != config.DIP_BUY_TICKER and t != 'SPY']
        
        if not current_stocks or self.cash <= 0:
            return
        
        total_equity = self._get_total_equity(date)
        if total_equity <= 0:
            return
        
        # 計算每檔股票的低配程度
        underweight_stocks = []
        for ticker in current_stocks:
            # [FIX] 跳過剛買入的股票，避免重複購買
            if ticker in exclude_tickers:
                continue
            
            if ticker not in self.target_weights:
                continue
            
            target_w = self.target_weights[ticker]
            current_qty = self.holdings.get(ticker, 0)
            price = self._get_price(ticker, date, use_open=True)
            if price <= 0:
                continue
            
            current_value = price * current_qty
            current_w = current_value / total_equity if total_equity > 0 else 0
            shortfall = target_w - current_w  # 低配程度 (正值表示低配)
            
            # 只有低配超過門檻 (3%) 才補足
            rebalance_threshold = config.REBALANCE_THRESHOLD
            if rebalance_threshold < 0.01:
                rebalance_threshold = 0.03
            
            if shortfall >= rebalance_threshold:
                underweight_stocks.append({
                    'ticker': ticker,
                    'price': price,
                    'shortfall': shortfall,
                    'target_w': target_w
                })
        
        if not underweight_stocks:
            return
        
        # 按低配程度比例分配剩餘現金
        total_shortfall = sum(s['shortfall'] for s in underweight_stocks)
        if total_shortfall <= 0:
            return
        
        available_cash = self.cash * 0.99  # 保留 1% buffer
        
        # 最小買入金額門檻 (總權益的百分比，例如 3%)
        min_buy_pct = getattr(config, 'MIN_BUY_AMOUNT_PCT', 0.03)
        min_buy_amount = total_equity * min_buy_pct
        
        for stock in underweight_stocks:
            # 按比例分配現金
            alloc_ratio = stock['shortfall'] / total_shortfall
            alloc_amount = available_cash * alloc_ratio
            
            qty = int(alloc_amount / (stock['price'] * (1 + config.COMMISSION)))
            buy_amount = stock['price'] * qty
            
            # 檢查：買入金額需 >= 總權益的 MIN_BUY_AMOUNT_PCT，且有足夠現金
            if buy_amount >= min_buy_amount and self.cash >= buy_amount * (1 + config.COMMISSION):
                reason = f"V3 低配補足 (目標:{stock['target_w']*100:.1f}%)"
                self._buy(stock['ticker'], date, stock['price'], qty, reason, target_weight=stock['target_w'])

    def _get_rotation_sells(self, date, signal_date):
        """
        獲取股票輪動需要賣出的股票
        返回: [(ticker, qty, price, reason)]
        """
        sells = []
        
        exit_ranked_list = self.selector.scan_market(signal_date, lookback=config.LOOKBACK)
        exit_candidate_tickers = [x['ticker'] for x in exit_ranked_list]
        
        top_n_threshold = config.SELL_RANK_THRESHOLD
        top_for_exit = exit_candidate_tickers[:top_n_threshold]
        
        current_stocks = [t for t in self.holdings if t != config.DIP_BUY_TICKER and t != 'SPY']
        
        for ticker in current_stocks:
            should_sell = False
            reason = ""
            
            if ticker not in top_for_exit:
                should_sell = True
                reason = f"排名跌出前{top_n_threshold}名"
                
            if not should_sell:
                metrics = next((x for x in exit_ranked_list if x['ticker'] == ticker), None)
                if metrics:
                    price = metrics['price']
                    exit_ema = metrics['exit_ema']
                    if price < exit_ema:
                        should_sell = True
                        reason = f"股價跌破EMA{config.EXIT_EMA}"
            
            if should_sell:
                exec_price = self._get_price(ticker, date, use_open=True)
                if exec_price > 0:
                    qty = self.holdings.get(ticker, 0)
                    sells.append((ticker, qty, exec_price, reason))
        
        return sells

    def _get_buy_candidates(self, date, signal_date):
        """
        確定買入候選並計算完整組合 ATR 權重
        返回: (buy_candidates_info, full_weights)
            buy_candidates_info: [(ticker, exec_price), ...]
            full_weights: {ticker: weight, ...} 包含保留持股 + 新候選的完整權重
        """
        entry_ranked_list = self.selector.scan_market(signal_date, lookback=config.LOOKBACK)
        initial_count = len(entry_ranked_list)
        
        # Filter by MAX_ADJ_SLOPE (過熱保護機制)
        max_adj_slope = getattr(config, 'MAX_ADJ_SLOPE', None)
        if max_adj_slope is not None:
            entry_ranked_list = [x for x in entry_ranked_list if x.get('adj_slope', 999) < max_adj_slope]
        after_slope_filter = len(entry_ranked_list)
        
        # Filter by SKIP_MAX_GAP_PCT (跳空缺口過濾)
        skip_max_gap = getattr(config, 'SKIP_MAX_GAP_PCT', 0.20)
        entry_ranked_list = [x for x in entry_ranked_list if x.get('max_gap', 0) < skip_max_gap]
        after_gap_filter = len(entry_ranked_list)
        
        entry_candidate_tickers = [x['ticker'] for x in entry_ranked_list]
        
        current_stocks = [t for t in self.holdings if t != config.DIP_BUY_TICKER and t != 'SPY']
        target_count = config.TARGET_HOLDINGS
        needed = target_count - len(current_stocks)
        
        if self.write_reports and needed > 0:
            print(f"  [ROTATION] Initial candidates: {initial_count}")
            print(f"  [ROTATION] After adj_slope<{max_adj_slope}: {after_slope_filter} (filtered {initial_count - after_slope_filter})")
            print(f"  [ROTATION] After max_gap<{skip_max_gap}: {after_gap_filter} (filtered {after_slope_filter - after_gap_filter})")
            print(f"  [ROTATION] Current holdings: {len(current_stocks)}, Need to buy: {needed}")
        
        buy_candidates_info = []  # [(ticker, exec_price)]
        full_weights = {}
        
        if needed > 0:
            buy_candidates = [t for t in entry_candidate_tickers if t not in current_stocks]

            if getattr(config, 'CORR_FILTER_ENABLED', False) and needed > 0:
                # 殘差相關性過濾
                candidate_metrics = [x for x in entry_ranked_list if x['ticker'] in buy_candidates]
                to_buy_tickers = self.selector.filter_by_residual_correlation(
                    ranked_candidates=candidate_metrics,
                    date=signal_date,
                    spy_df=self.spy_df,
                    threshold=config.CORR_THRESHOLD,
                    lookback=config.CORR_LOOKBACK,
                    max_candidates=config.CORR_CANDIDATE_COUNT,
                    needed=needed,
                    existing_tickers=current_stocks
                )
            else:
                to_buy_tickers = buy_candidates[:needed]
            
            if self.write_reports:
                print(f"  [ROTATION] Buy candidates available: {len(buy_candidates)}")
                if getattr(config, 'CORR_FILTER_ENABLED', False):
                    print(f"  [CORR] Residual correlation filter: threshold={config.CORR_THRESHOLD}, lookback={config.CORR_LOOKBACK}")
                print(f"  [ROTATION] Selected to buy: {to_buy_tickers}")
            
            # 獲取候選股票的執行價格
            for ticker in to_buy_tickers:
                exec_price = self._get_price(ticker, date, use_open=True)
                if exec_price > 0:
                    buy_candidates_info.append((ticker, exec_price))
            
            # 計算完整 ATR 權重 (保留持股 + 新候選)
            all_tickers = current_stocks + [t for t, _ in buy_candidates_info]
            all_with_atr = []
            for ticker in all_tickers:
                metrics = next((x for x in entry_ranked_list if x['ticker'] == ticker), None)
                if metrics is None:
                    metrics = self.selector.calculate_metrics(ticker, date, config.LOOKBACK)
                if metrics and metrics.get('atr_pct', 0) > 0:
                    all_with_atr.append(metrics)
            
            full_weights = self._calculate_atr_weights(all_with_atr)
        else:
            # 不需新增，用當前持股算權重
            all_with_atr = []
            for ticker in current_stocks:
                metrics = next((x for x in entry_ranked_list if x['ticker'] == ticker), None)
                if metrics is None:
                    metrics = self.selector.calculate_metrics(ticker, date, config.LOOKBACK)
                if metrics and metrics.get('atr_pct', 0) > 0:
                    all_with_atr.append(metrics)
            full_weights = self._calculate_atr_weights(all_with_atr)
        
        return buy_candidates_info, full_weights



    def _check_stop_loss(self, date, prev_date):
        """
        個股停損檢查：
        若前一日收盤價 (prev_date Close) < 平均成本 * 0.9 (10% Loss)
        則於今日開盤 (date Open)賣出
        SSO 不執行停損
        """
        current_holdings = list(self.holdings.keys())
        
        for ticker in current_holdings:
            if ticker == config.DIP_BUY_TICKER:
                continue
                
            price_close_prev = self._get_price(ticker, prev_date, use_open=False)
            if price_close_prev <= 0: continue
            
            avg_cost = self.avg_costs.get(ticker, 0)
            if avg_cost > 0:
                threshold = avg_cost * (1 - config.STOP_LOSS_PCT)
                
                if price_close_prev < threshold:
                    price_open_curr = self._get_price(ticker, date, use_open=True)
                    if price_open_curr > 0:
                        reason = f"Stop Loss (Prev Close {price_close_prev:.2f} < Cost {avg_cost:.2f})"
                        self._sell(ticker, date, price_open_curr, self.holdings[ticker], reason)
                        # V3: 清除該股票的目標權重
                        if ticker in self.target_weights:
                            del self.target_weights[ticker]

    def _check_gap_exit(self, date, prev_date):
        """
        跳空缺口出場檢查：
        若前一日出現 15% 以上跳空缺口 (Gap Up 或 Gap Down)
        則於今日開盤出場
        """
        current_holdings = list(self.holdings.keys())
        gap_threshold = getattr(config, 'GAP_EXIT_PCT', 0.15)
        
        for ticker in current_holdings:
            if ticker == config.DIP_BUY_TICKER:
                continue
            
            price_open_prev = self._get_price(ticker, prev_date, use_open=True)
            
            try:
                prev_idx = self.calendar.get_loc(prev_date)
                if prev_idx > 0:
                    day_before_prev = self.calendar[prev_idx - 1]
                    price_close_before = self._get_price(ticker, day_before_prev, use_open=False)
                    
                    if price_close_before > 0 and price_open_prev > 0:
                        gap = (price_open_prev - price_close_before) / price_close_before
                        
                        if abs(gap) >= gap_threshold:
                            price_open_curr = self._get_price(ticker, date, use_open=True)
                            if price_open_curr > 0:
                                direction = "Up" if gap > 0 else "Down"
                                reason = f"Gap Exit ({gap*100:+.1f}% {direction} on {prev_date.date()})"
                                self._sell(ticker, date, price_open_curr, self.holdings[ticker], reason)
                                # V3: 清除該股票的目標權重
                                if ticker in self.target_weights:
                                    del self.target_weights[ticker]
            except (KeyError, IndexError):
                continue

    def _check_bear_sso_logic(self, date):
        """
        熊市流程 (V3 改進版 - 連續兩周確認):
        1. 確認 close of spy < spy 200 ma (Bear) → 清倉個股，抄底 SSO
        2. [NEW] SPY > 200MA 連續兩周 → 清倉 SSO，啟動牛市策略
        3. SSO 分批抄底: -15%, -20%, -25%
        """
        state = self.market_regime.get_state(date)
        if not state: return
        
        spy_close = state['SPY_Close']
        spy_ma200 = state['SPY_MA200']
        spy_dd = abs(state['SPY_DD'])
        
        is_rebalance_day = (date.weekday() == config.REBALANCE_WEEKDAY)
        
        # === 市場狀態判斷（每周更新） ===
        if is_rebalance_day:
            if spy_close > spy_ma200:
                # SPY > 200MA: 增加牛市計數器
                self.bull_weeks_counter += 1
            else:
                # SPY < 200MA: 重置牛市計數器
                self.bull_weeks_counter = 0
        
        # === 牛市確認：連續兩周 SPY > 200MA ===
        bull_confirmed = (self.bull_weeks_counter >= 2)
        
        if bull_confirmed:
            # [牛市確認] 清倉 SSO，允許個股交易
            if is_rebalance_day and config.DIP_BUY_TICKER in self.holdings:
                self._close_positions(date, target_type='DIP')
                for k in self.dip_state: 
                    self.dip_state[k] = False
            # 注意：不重置 bull_weeks_counter，保持牛市狀態
            return

        # === 熊市邏輯：SPY < 200MA ===
        if spy_close < spy_ma200:
            # 清倉所有個股（如果有的話）
            current_stocks = [t for t in self.holdings if t != config.DIP_BUY_TICKER]
            if current_stocks:
                self._close_positions(date, target_type='STOCK')
                # V3: 清除所有目標權重
                self.target_weights.clear()
            
            # SSO 抄底：-25% 買入 40%，其他 30%
            dip_allocations = {0.15: 0.30, 0.20: 0.30, 0.25: 0.40}
            
            for level in [0.15, 0.20, 0.25]:
                if spy_dd >= level and not self.dip_state[level]:
                    alloc_pct = dip_allocations[level]
                    if self.compounding:
                        curr_equity = self._get_total_equity(date)
                        target_amt = curr_equity * alloc_pct
                    else:
                        target_amt = self.initial_capital * alloc_pct
                    
                    buy_amt = min(target_amt, self.cash)
                    price = self._get_price(config.DIP_BUY_TICKER, date)
                    
                    if price > 0 and buy_amt > 0:
                        qty = int(buy_amt / (price * (1 + config.COMMISSION)))
                        if qty > 0:
                            self._buy(config.DIP_BUY_TICKER, date, price, qty, f"Bear Dip Buy -{level*100:.0f}%")
                            self.dip_state[level] = True




    def _get_price(self, ticker, date, use_open=False):
        if ticker == config.DIP_BUY_TICKER:
            df = self.sso_df
        elif ticker == 'SPY':
            df = self.spy_df
        else:
            df = self.selector._get_ticker_data(ticker)
            
        if df is not None and not df.empty:
            if date in df.index:
                row = df.loc[date]
                if use_open and 'Open' not in df.columns:
                    val = row['Close']
                else:
                    val = row['Open'] if use_open else row['Close']
                return float(val)
        return 0.0

    def _close_positions(self, date, target_type='ALL'):
        holdings_list = list(self.holdings.keys())
        for ticker in holdings_list:
            is_dip = (ticker == config.DIP_BUY_TICKER)
            
            should_sell = False
            if target_type == 'ALL': should_sell = True
            elif target_type == 'DIP' and is_dip: should_sell = True
            elif target_type == 'STOCK' and not is_dip and ticker != 'SPY': should_sell = True
            
            if should_sell:
                use_open = True
                if target_type == 'DIP': use_open = False
                
                price = self._get_price(ticker, date, use_open=use_open)
                if price > 0:
                    self._sell(ticker, date, price, self.holdings[ticker], f"Clear {target_type}")

    def _force_close_all(self, date):
        self._close_positions(date, target_type='ALL')

    def _buy(self, ticker, date, price, qty, reason, target_weight=None):
        # V3: 計算買入前該股的權重
        total_equity_before = self._get_total_equity(date)
        old_qty = self.holdings.get(ticker, 0)
        weight_before = (price * old_qty / total_equity_before * 100) if total_equity_before > 0 and old_qty > 0 else 0
        
        cost = price * qty * (1 + config.COMMISSION)
        self.cash -= cost
        
        if ticker not in self.holdings:
            self.holdings[ticker] = qty
            self.avg_costs[ticker] = price
        else:
            total_val = (self.avg_costs[ticker] * old_qty) + (price * qty)
            self.holdings[ticker] += qty
            self.avg_costs[ticker] = total_val / self.holdings[ticker]
        
        # V3: 計算買入後該股的權重
        total_equity_after = self._get_total_equity(date)
        position_value = price * self.holdings[ticker]
        weight_after = (position_value / total_equity_after * 100) if total_equity_after > 0 else 0
        
        # V3: 目標權重 (如果有傳入)
        target_w_pct = target_weight * 100 if target_weight is not None else None
            
        if self.write_reports:
            target_str = f" | Target: {target_w_pct:.1f}%" if target_w_pct else ""
            print(f"  BUY {ticker}: {qty} @ {price:.2f} ({reason}) | Weight: {weight_before:.1f}% -> {weight_after:.1f}%{target_str}")
        holdings_snapshot = self._get_holdings_snapshot(date)  # [NEW]
        self.trades.append({
            'Date': date,
            'Ticker': ticker,
            'Action': 'BUY',
            'Price': price,
            'Quantity': qty,
            'Cost': cost,
            'Reason': reason,
            'Revenue': '',
            'PnL': '',
            'PnL_Pct': '',
            'Weight_Before': weight_before,
            'Weight_After': weight_after,
            'Target_Weight': target_w_pct,  # V3: ATR 目標權重 (只在再平衡時有值)
            'Total_Equity': total_equity_after,
            'Holdings_After': holdings_snapshot  # [NEW] 交易後完整持倉
        })
        
    def _sell(self, ticker, date, price, qty, reason, target_weight=None, weight_before=None):
        # V3: 使用傳入的 weight_before，若無則計算
        if weight_before is None:
            total_equity_before = self._get_total_equity(date)
            current_qty = self.holdings.get(ticker, 0)
            weight_before = (price * current_qty / total_equity_before * 100) if total_equity_before > 0 else 0
        
        current_qty = self.holdings.get(ticker, 0)
        remaining_qty = current_qty - qty
        
        revenue = price * qty * (1 - config.COMMISSION)
        self.cash += revenue
        
        cost_basis = self.avg_costs[ticker] * qty * (1 + config.COMMISSION)
        pnl = revenue - cost_basis
        pnl_pct = (pnl / cost_basis) * 100 if cost_basis > 0 else 0
        
        if ticker in self.holdings:
            self.holdings[ticker] -= qty
            if self.holdings[ticker] <= 0:
                del self.holdings[ticker]
                del self.avg_costs[ticker]
        
        # V3: 賣出後計算權重
        total_equity_after = self._get_total_equity(date)
        weight_after = 0
        if remaining_qty > 0 and total_equity_after > 0:
            weight_after = (price * remaining_qty / total_equity_after * 100)
        
        # V3: 目標權重 (如果有傳入)
        target_w_pct = target_weight * 100 if target_weight is not None else None
        
        if self.write_reports:
            target_str = f" | Target: {target_w_pct:.1f}%" if target_w_pct else ""
            print(f"  SELL {ticker}: {qty} @ {price:.2f} ({reason}) | Weight: {weight_before:.1f}% -> {weight_after:.1f}%{target_str} | PnL: {pnl:.2f}")
        holdings_snapshot = self._get_holdings_snapshot(date)  # [NEW]
        self.trades.append({
            'Date': date,
            'Ticker': ticker,
            'Action': 'SELL',
            'Price': price,
            'Quantity': qty,
            'Cost': '',
            'Reason': reason,
            'Revenue': revenue,
            'PnL': pnl,
            'PnL_Pct': pnl_pct,
            'Weight_Before': weight_before,
            'Weight_After': weight_after,
            'Target_Weight': target_w_pct,  # V3: ATR 目標權重 (只在再平衡時有值)
            'Total_Equity': total_equity_after,
            'Holdings_After': holdings_snapshot  # [NEW] 交易後完整持倉
        })

    def _update_equity(self, date):
        total_equity = self._get_total_equity(date)
        self.history.append({'Date': date, 'Equity': total_equity, 'Cash': self.cash})

    def _get_total_equity(self, date):
        val = 0.0
        for t, q in self.holdings.items():
            price = self._get_price(t, date, use_open=False)
            val += float(price) * q
        return self.cash + val

    def _get_holdings_snapshot(self, date):
        """
        生成當前持倉快照字串（用於交易記錄）
        格式: AAPL:25%, MSFT:30%, CASH:45%
        """
        total_equity = self._get_total_equity(date)
        if total_equity <= 0:
            return "CASH:100%"
        
        holdings_parts = []
        for ticker, qty in sorted(self.holdings.items()):
            if qty <= 0:
                continue
            price = self._get_price(ticker, date, use_open=False)
            value = price * qty
            weight = (value / total_equity * 100)
            holdings_parts.append(f"{ticker}:{weight:.1f}%")
        
        # 添加現金比例
        cash_weight = (self.cash / total_equity * 100) if total_equity > 0 else 100
        if cash_weight > 0.1:  # 只顯示 > 0.1% 的現金
            holdings_parts.append(f"CASH:{cash_weight:.1f}%")
        
        return ", ".join(holdings_parts) if holdings_parts else "CASH:100%"



    def get_current_holdings(self, date=None):
        """Get current holdings status for live tracking"""
        if date is None:
            date = self.end_date
        
        total_equity = self._get_total_equity(date)
        holdings_list = []
        
        for ticker, qty in self.holdings.items():
            if qty <= 0:
                continue
            
            current_price = self._get_price(ticker, date, use_open=False)
            avg_cost = self.avg_costs.get(ticker, 0)
            value = current_price * qty
            cost_basis = avg_cost * qty
            pnl = value - cost_basis
            pnl_pct = (pnl / cost_basis * 100) if cost_basis > 0 else 0
            weight = (value / total_equity * 100) if total_equity > 0 else 0
            target_w = self.target_weights.get(ticker, 0) * 100
            
            holdings_list.append({
                'ticker': ticker,
                'qty': qty,
                'avg_cost': avg_cost,
                'current_price': current_price,
                'value': value,
                'pnl': pnl,
                'pnl_pct': pnl_pct,
                'weight': weight,
                'target_weight': target_w
            })
        
        holdings_list.sort(key=lambda x: x['weight'], reverse=True)
        
        return {
            'date': date,
            'holdings': holdings_list,
            'cash': self.cash,
            'total_equity': total_equity,
            'target_weights': {k: v*100 for k, v in self.target_weights.items()}
        }

    def _generate_report(self):
        if self.write_reports:
            print(f"\n--- Backtest Final Complete ({self.report_suffix.strip('_') if self.report_suffix else 'Default'}) ---")
            final_eq = self._get_total_equity(self.end_date)
            print(f"Final Cash: {self.cash:.2f}")
            print(f"Final Equity: {final_eq:.2f}")
        
        suffix = self.report_suffix if self.report_suffix else ""
        if not suffix.startswith("_final"):
            suffix = "_final" + suffix
        
        if self.trades:
            trades_df = pd.DataFrame(self.trades)
            trades_df.to_csv(f'backtest_trades{suffix}.csv')
            
        if self.history:
            history_df = pd.DataFrame(self.history)
            history_df.set_index('Date', inplace=True)
            history_df.to_csv(f'equity_curve{suffix}.csv')
        
        # LIVE_MODE: Export current holdings to JSON
        if getattr(config, 'LIVE_MODE', False):
            import json
            holdings_info = self.get_current_holdings()
            holdings_info['date'] = str(holdings_info['date'].date())
            with open(f'current_holdings{suffix}.json', 'w', encoding='utf-8') as f:
                json.dump(holdings_info, f, indent=2, ensure_ascii=False)
            if self.write_reports:
                print(f"[Current Holdings]")
                for h in holdings_info['holdings']:
                    pnl_sign = '+' if h['pnl'] >= 0 else ''
                    print(f"  {h['ticker']}: {h['qty']} shares @ ${h['current_price']:.2f} | PnL: {pnl_sign}${h['pnl']:.2f} ({h['pnl_pct']:+.1f}%) | Weight: {h['weight']:.1f}%")

    def _record_rebalance_snapshot(self, date, signal_date):
        """閮?隤踹??????敹怎"""
        # ?脣??嗅???
        total_equity = self._get_total_equity(date)
        holdings_snapshot = {}
        for ticker, qty in self.holdings.items():
            if qty > 0:
                price = self._get_price(ticker, date, use_open=False)
                value = price * qty
                weight = (value / total_equity * 100) if total_equity > 0 else 0
                holdings_snapshot[ticker] = weight
        
        # ?脣? Top 20 ??
        entry_ranked_list = self.selector.scan_market(signal_date, lookback=config.LOOKBACK)
        
        # ??蕪璇辣嚗? _get_rotation_buys ?詨?嚗?
        max_adj_slope = getattr(config, 'MAX_ADJ_SLOPE', None)
        if max_adj_slope is not None:
            entry_ranked_list = [x for x in entry_ranked_list if x.get('adj_slope', 999) < max_adj_slope]
        
        skip_max_gap = getattr(config, 'SKIP_MAX_GAP_PCT', 0.20)
        entry_ranked_list = [x for x in entry_ranked_list if x.get('max_gap', 0) < skip_max_gap]
        
        top20_tickers = [x['ticker'] for x in entry_ranked_list[:20]]
        
        # 靽?敹怎
        self.rebalance_snapshots.append({
            'date': date,
            'holdings': holdings_snapshot,
            'top20': top20_tickers
        })
    
    def export_rebalance_excel(self, filename='rebalance_holdings.xlsx'):
        """撠隤踹??????Excel"""
        import pandas as pd
        
        if not self.rebalance_snapshots:
            print("No rebalance snapshots to export.")
            return
        
        # Sheet 1: Holdings (璈怠??澆?)
        # Date, Ticker1, Weight1, Ticker2, Weight2, Ticker3, Weight3, Ticker4, Weight4
        holdings_rows = []
        for snapshot in self.rebalance_snapshots:
            row = {'Date': snapshot['date'].strftime('%Y-%m-%d')}
            holdings = snapshot['holdings']
            
            # ????摨???
            sorted_holdings = sorted(holdings.items(), key=lambda x: x[1], reverse=True)
            
            for i, (ticker, weight) in enumerate(sorted_holdings, 1):
                row[f'Ticker{i}'] = ticker
                row[f'Weight{i}'] = f"{weight:.2f}%"
            
            holdings_rows.append(row)
        
        df_holdings = pd.DataFrame(holdings_rows)
        
        # Sheet 2: Top20 Rankings (璈怠??澆?)
        # Date, Rank1, Rank2, ..., Rank20
        rankings_rows = []
        for snapshot in self.rebalance_snapshots:
            row = {'Date': snapshot['date'].strftime('%Y-%m-%d')}
            top20 = snapshot['top20']
            
            for i, ticker in enumerate(top20, 1):
                row[f'Rank{i}'] = ticker
            
            rankings_rows.append(row)
        
        df_rankings = pd.DataFrame(rankings_rows)
        
        # 撖怠 Excel
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df_holdings.to_excel(writer, sheet_name='Holdings', index=False)
            df_rankings.to_excel(writer, sheet_name='Top20 Rankings', index=False)
        
        print(f"Rebalance snapshots exported to: {filename}")
