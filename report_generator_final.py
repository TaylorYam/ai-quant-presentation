"""
Report Generator Final - Redesigned UI
即時持股模擬系統報告生成器 - 新增當前持股狀況顯示
"""
import pandas as pd
import numpy as np
import os
import json
import config_final as config

def calculate_period_returns(df_eq):
    """計算年度和月度報酬率分析"""
    if df_eq.empty:
        return {}, {}
    
    df_eq = df_eq.copy()
    
    # Monthly Returns
    monthly = df_eq['Equity'].resample('ME').last()
    monthly_returns = monthly.pct_change() * 100
    
    if len(monthly) > 0:
        first_month_end = monthly.iloc[0]
        first_day_val = df_eq['Equity'].iloc[0]
        first_month_ret = ((first_month_end - first_day_val) / first_day_val) * 100
        
        if pd.isna(monthly_returns.iloc[0]):
            monthly_returns.iloc[0] = first_month_ret
            
    # Remove NaN at start only if it's truly empty (not first month calc)
    monthly_returns = monthly_returns.dropna()
    
    # Annual Returns
    yearly = df_eq['Equity'].resample('YE').last()
    yearly_returns = yearly.pct_change() * 100
    
    if len(yearly) > 0:
        first_year_end = yearly.iloc[0]
        first_day_val = df_eq['Equity'].iloc[0]
        first_year_ret = ((first_year_end - first_day_val) / first_day_val) * 100
        
        if pd.isna(yearly_returns.iloc[0]):
             yearly_returns.iloc[0] = first_year_ret
             
    yearly_returns = yearly_returns.dropna()
    
    return monthly_returns, yearly_returns

def calculate_drawdown_periods(df_eq):
    """
    計算歷史最大下跌風險的詳細信息
    返回每個下跌期間的：開始時間、結束時間（最低點）、恢復時間、虧損程度、持續天數
    """
    if df_eq.empty:
        return []
    
    df_eq = df_eq.copy()
    equity = df_eq['Equity']
    
    # 計算累積最大值和下跌幅度
    rolling_max = equity.cummax()
    drawdown = (equity - rolling_max) / rolling_max * 100
    
    # 找出所有高點（新的累積最大值）
    is_peak = equity >= rolling_max
    peak_indices = df_eq[is_peak].index.tolist()
    peak_values = equity[is_peak].tolist()
    
    if len(peak_indices) < 2:
        return []
    
    drawdown_periods = []
    
    # 遍歷每個高點，找出從該高點開始的下跌期間
    for i in range(len(peak_indices) - 1):
        peak_date = peak_indices[i]
        peak_value = peak_values[i]
        next_peak_date = peak_indices[i + 1]
        
        # 找出這個高點到下一個高點之間的所有數據
        period_data = df_eq[(df_eq.index >= peak_date) & (df_eq.index <= next_peak_date)]
        
        if len(period_data) < 2:
            continue
        
        # 找出最低點（最大下跌）
        period_equity = period_data['Equity']
        min_idx = period_equity.idxmin()
        min_value = float(period_equity.loc[min_idx].iloc[0] if hasattr(period_equity.loc[min_idx], 'iloc') else period_equity.loc[min_idx])
        min_date = min_idx
        
        # 計算下跌幅度
        drawdown_pct = ((min_value - peak_value) / peak_value) * 100
        
        # 如果下跌幅度小於 0.5%，忽略（避免太多小波動）
        if abs(drawdown_pct) < 0.5:
            continue
        
        # 計算持續天數（從高點到最低點）
        duration_days = (min_date - peak_date).days
        
        # 找出恢復時間（回到或超過原高點的時間）
        recovery_date = None
        recovery_days = None
        
        # 在下一個高點之前尋找恢復點
        recovery_data = df_eq[(df_eq.index > min_date) & (df_eq.index <= next_peak_date)]
        if not recovery_data.empty:
            # 找出第一個超過或等於原高點的日期
            recovery_mask = recovery_data['Equity'] >= peak_value
            if recovery_mask.any():
                recovery_date = recovery_data[recovery_mask].index[0]
                recovery_days = (recovery_date - peak_date).days
            else:
                # 如果在下一個高點之前沒有恢復，使用下一個高點作為恢復時間
                recovery_date = next_peak_date
                recovery_days = (recovery_date - peak_date).days
        
        drawdown_periods.append({
            'start_date': peak_date,
            'end_date': min_date,  # 最低點
            'recovery_date': recovery_date,
            'start_value': peak_value,
            'end_value': min_value,
            'drawdown_pct': drawdown_pct,
            'duration_days': duration_days,
            'recovery_days': recovery_days
        })
    
    # 按下跌幅度排序（從大到小，最嚴重的在前）
    # drawdown_pct 是負數，所以排序後最負的（最嚴重的）在前面
    drawdown_periods.sort(key=lambda x: x['drawdown_pct'])
    
    return drawdown_periods

def load_current_holdings(suffix):
    """載入當前持股 JSON"""
    json_file = f'current_holdings{suffix}.json'
    if os.path.exists(json_file):
        with open(json_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None

def generate_comparison_report():
    print("\n--- Generating Final Version Report (Redesigned UI - Compound Only) ---")
    
    # Only using Compound mode now
    mode = 'Compound'
    info = {'suffix': '_final_compound', 'color': '#FF7F0E'}
    
    data = {}
    
    eq_file = f'equity_curve{info["suffix"]}.csv'
    tr_file = f'backtest_trades{info["suffix"]}.csv'
    
    if os.path.exists(eq_file):
        df_eq = pd.read_csv(eq_file)
        df_eq['Date'] = pd.to_datetime(df_eq['Date'])
        df_eq.set_index('Date', inplace=True)
        
        if os.path.exists(tr_file):
            df_trades = pd.read_csv(tr_file)
        else:
            df_trades = pd.DataFrame()
            
        # Calculate Metrics
        initial_cap = df_eq['Equity'].iloc[0] if not df_eq.empty else 0
        final_cap = df_eq['Equity'].iloc[-1] if not df_eq.empty else 0
        total_return = ((final_cap - initial_cap) / initial_cap) * 100 if initial_cap > 0 else 0
        
        # CAGR
        cagr = 0
        if not df_eq.empty:
            days = (df_eq.index[-1] - df_eq.index[0]).days
            if days > 0:
                years = days / 365.25
                if initial_cap > 0:
                    cagr = ((final_cap / initial_cap) ** (1/years) - 1) * 100
        
        # Drawdown
        rolling_max = df_eq['Equity'].cummax()
        drawdown = (df_eq['Equity'] - rolling_max) / rolling_max * 100
        max_drawdown = drawdown.min()
        
        is_peak = df_eq['Equity'] >= rolling_max
        peaks_idx = df_eq[is_peak].index.strftime('%Y-%m-%d').tolist()
        peaks_val = df_eq[is_peak]['Equity'].tolist()
        
        # Sharpe
        sharpe = 0
        daily_sharpe = 0
        if not df_eq.empty:
            daily_returns = df_eq['Equity'].pct_change().dropna()
            mean_ret = daily_returns.mean()
            std_ret = daily_returns.std()
            if std_ret != 0:
                daily_sharpe = mean_ret / std_ret
                sharpe = daily_sharpe * (252**0.5)
        
        # Win Ratio
        win_ratio = 0
        win_count = 0
        total_closed_trades = 0
        if not df_trades.empty and 'PnL' in df_trades.columns:
            closed_trades = df_trades[df_trades['Action'] == 'SELL']
            total_closed_trades = len(closed_trades)
            if total_closed_trades > 0:
                win_count = len(closed_trades[closed_trades['PnL'] > 0])
                win_ratio = (win_count / total_closed_trades) * 100

        monthly_rets, yearly_rets = calculate_period_returns(df_eq)
        
        # Calculate drawdown periods
        drawdown_periods = calculate_drawdown_periods(df_eq)
        
        # Load current holdings
        holdings_data = load_current_holdings(info['suffix'])
        
        data = {
            'equity': df_eq,
            'trades': df_trades,
            'drawdown': drawdown,
            'peaks_idx': peaks_idx,
            'peaks_val': peaks_val,
            'monthly_returns': monthly_rets,
            'yearly_returns': yearly_rets,
            'drawdown_periods': drawdown_periods,
            'holdings': holdings_data,
            'metrics': {
                'initial': initial_cap,
                'final': final_cap,
                'return': total_return,
                'cagr': cagr,
                'mdd': max_drawdown,
                'sharpe': sharpe,
                'daily_sharpe': daily_sharpe,
                'count': len(df_trades),
                'win_ratio': win_ratio,
                'win_count': win_count,
                'total_trades': total_closed_trades
            },
            'color': info['color']
        }
    else:
        print(f"Warning: {eq_file} not found.")
        return # Cannot generate report without main data
    
    # Load benchmark data (SPY and QQQ)
    benchmarks = {}
    for bm_ticker in ['SPY', 'QQQ']:
        bm_file = os.path.join(config.DATA_DIR, f'{bm_ticker}.csv')
        if os.path.exists(bm_file):
            try:
                bm_df = pd.read_csv(bm_file)
                bm_df['Date'] = pd.to_datetime(bm_df['Date'])
                bm_df.set_index('Date', inplace=True)
                benchmarks[bm_ticker] = bm_df
                print(f"Loaded benchmark: {bm_ticker}")
            except Exception as e:
                print(f"Warning: Could not load {bm_ticker}: {e}")
        else:
            print(f"Warning: {bm_file} not found")

    # HTML Template Construction
    html_content = """
    <!DOCTYPE html>
    <html lang="zh-Hant">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>策略回測報告</title>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">
        <style>
            :root {
                /* FinLab 風格配色 */
                --primary-color: #1a237e; /* 深藍色主色調 */
                --primary-light: #3f51b5;
                --accent-color: #4caf50; /* 綠色強調 */
                --positive-color: #4caf50; /* 正收益綠色 */
                --positive-bg: #e8f5e9;
                --negative-color: #f44336; /* 負收益紅色 */
                --negative-bg: #ffebee;
                --text-primary: #212121; /* 主文字深灰 */
                --text-secondary: #757575; /* 次要文字中灰 */
                --text-light: #9e9e9e; /* 淺灰文字 */
                --bg-primary: #ffffff;
                --bg-secondary: #f5f5f5;
                --border-color: #e0e0e0;
                --shadow-sm: 0 2px 4px rgba(0,0,0,0.05);
                --shadow-md: 0 2px 8px rgba(0,0,0,0.1);
                --shadow-lg: 0 4px 16px rgba(0,0,0,0.12);
            }
            
            * {
                box-sizing: border-box;
            }
            
            body { 
                font-family: 'Roboto', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; 
                margin: 0; 
                padding: 0;
                background: var(--bg-secondary); 
                color: var(--text-primary);
                line-height: 1.6;
                animation: fadeIn 0.5s ease-in;
            }
            
            @keyframes fadeIn {
                from { opacity: 0; }
                to { opacity: 1; }
            }
            
            .container {
                max-width: 1400px;
                margin: 0 auto;
                background: var(--bg-primary);
                box-shadow: var(--shadow-lg);
                border-radius: 8px;
                padding: 48px;
                margin-top: 24px;
                margin-bottom: 24px;
                animation: slideUp 0.5s ease-out;
            }
            
            @keyframes slideUp {
                from { 
                    opacity: 0;
                    transform: translateY(20px);
                }
                to { 
                    opacity: 1;
                    transform: translateY(0);
                }
            }
            
            h1 { 
                font-weight: 400; 
                margin-bottom: 8px; 
                color: var(--text-primary);
                font-size: 32px;
                letter-spacing: -0.5px;
            }
            
            .header-meta {
                font-size: 13px;
                color: var(--text-secondary);
                margin-top: 4px;
            }
            
            /* Metric Cards - FinLab 風格 */
            .metrics-container {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
                gap: 24px;
                margin-bottom: 48px;
                margin-top: 32px;
            }
            
            .metric-card {
                background: var(--bg-primary);
                padding: 24px;
                border-radius: 8px;
                box-shadow: var(--shadow-md);
                border-left: 4px solid var(--primary-color);
                transition: all 0.3s ease;
                position: relative;
                overflow: hidden;
            }
            
            .metric-card::before {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                width: 4px;
                height: 100%;
                background: var(--primary-color);
                transition: width 0.3s ease;
            }
            
            .metric-card:hover {
                box-shadow: var(--shadow-lg);
                transform: translateY(-2px);
            }
            
            .metric-card:hover::before {
                width: 100%;
                opacity: 0.05;
            }
            
            .metric-title {
                font-size: 14px;
                color: var(--text-secondary);
                margin-bottom: 8px;
                font-weight: 500;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                font-size: 12px;
            }
            
            .metric-subtitle {
                font-size: 11px;
                color: var(--text-light);
                margin-bottom: 12px;
                text-transform: lowercase;
                letter-spacing: 0.3px;
            }
            
            .metric-value {
                font-size: 42px;
                font-weight: 300;
                color: var(--text-primary);
                line-height: 1.1;
                margin-top: 8px;
                letter-spacing: -1px;
            }
            
            .value-pos { 
                color: var(--positive-color);
                font-weight: 400;
            }
            
            .value-neg { 
                color: var(--negative-color);
                font-weight: 400;
            }
            
            /* 響應式指標卡片 */
            @media (max-width: 768px) {
                .metrics-container {
                    grid-template-columns: 1fr;
                    gap: 16px;
                }
                .metric-value {
                    font-size: 36px;
                }
            }

            /* Main Chart - FinLab 風格 */
            .chart-section {
                margin-bottom: 48px;
                height: 500px;
                background: var(--bg-primary);
                border-radius: 8px;
                box-shadow: var(--shadow-sm);
                padding: 16px;
                transition: box-shadow 0.3s ease;
            }
            
            .chart-section:hover {
                box-shadow: var(--shadow-md);
            }
            
            @media (max-width: 768px) {
                .chart-section {
                    height: 400px;
                    padding: 8px;
                }
            }
            
            /* Analysis Tabs - FinLab 風格 */
            .analysis-tabs {
                display: flex;
                border-bottom: 2px solid var(--border-color);
                margin-bottom: 32px;
                gap: 0;
                background: var(--bg-primary);
            }
            
            .analysis-tab {
                padding: 16px 24px;
                cursor: pointer;
                font-size: 15px;
                background: none;
                border: none;
                color: var(--text-secondary);
                position: relative;
                transition: all 0.3s ease;
                font-weight: 500;
                border-bottom: 3px solid transparent;
                margin-bottom: -2px;
            }
            
            .analysis-tab:hover { 
                color: var(--primary-color);
                background: rgba(26, 35, 126, 0.04);
            }
            
            .analysis-tab.active {
                color: var(--primary-color);
                font-weight: 600;
                border-bottom-color: var(--primary-color);
            }
            
            .analysis-tab.active::before {
                content: '';
                position: absolute;
                bottom: -2px;
                left: 0;
                width: 100%;
                height: 3px;
                background: var(--primary-color);
                animation: slideIn 0.3s ease;
            }
            
            @keyframes slideIn {
                from {
                    width: 0;
                }
                to {
                    width: 100%;
                }
            }
            
            /* Tab Content - 動畫效果 */
            .tab-content { 
                display: none; 
                animation: fadeInTab 0.4s ease-out;
            }
            
            .tab-content.active { 
                display: block; 
            }
            
            @keyframes fadeInTab { 
                from { 
                    opacity: 0; 
                    transform: translateY(10px);
                } 
                to { 
                    opacity: 1; 
                    transform: translateY(0);
                } 
            }
            
            .tab-content h3 {
                color: var(--text-primary);
                font-size: 18px;
                font-weight: 500;
                margin-top: 0;
                margin-bottom: 20px;
                padding-bottom: 12px;
                border-bottom: 2px solid var(--border-color);
            }
            
            /* Tables - FinLab 風格 */
            .custom-table {
                width: 100%;
                border-collapse: collapse;
                font-size: 14px;
                color: var(--text-primary);
                background: var(--bg-primary);
                border-radius: 8px;
                overflow: hidden;
                box-shadow: var(--shadow-sm);
            }
            
            .custom-table th {
                text-align: left;
                padding: 16px;
                color: var(--text-primary);
                font-weight: 600;
                border-bottom: 2px solid var(--border-color);
                background: var(--bg-secondary);
                text-transform: uppercase;
                font-size: 12px;
                letter-spacing: 0.5px;
            }
            
            .custom-table td {
                padding: 16px;
                border-bottom: 1px solid var(--border-color);
                transition: background 0.2s ease;
            }
            
            .custom-table tbody tr {
                transition: all 0.2s ease;
            }
            
            .custom-table tbody tr:hover {
                background: rgba(26, 35, 126, 0.04);
                transform: scale(1.01);
            }
            
            .custom-table tbody tr:last-child td {
                border-bottom: none;
            }
            
            /* Heatmap Table - FinLab 風格 */
            .heatmap-table {
                width: 100%;
                border-collapse: separate;
                border-spacing: 3px;
                font-size: 12px;
                margin-top: 24px;
                background: var(--bg-secondary);
                padding: 8px;
                border-radius: 8px;
            }
            
            .heatmap-table th, .heatmap-table td {
                padding: 12px 8px;
                text-align: center;
                border-radius: 4px;
                transition: transform 0.2s ease, box-shadow 0.2s ease;
            }
            
            .heatmap-table th { 
                background: var(--bg-primary); 
                color: var(--text-secondary); 
                font-weight: 600;
                font-size: 11px;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }
            
            .heatmap-cell {
                color: var(--text-primary);
                font-weight: 600;
                cursor: default;
            }
            
            .heatmap-cell:hover {
                transform: scale(1.1);
                box-shadow: var(--shadow-md);
                z-index: 1;
                position: relative;
            }
            
            /* FinLab 風格配色：綠色漸層 (正收益) */
            .cell-pos-1 { background-color: #c8e6c9; color: #1b5e20; }
            .cell-pos-2 { background-color: #a5d6a7; color: #1b5e20; }
            .cell-pos-3 { background-color: #81c784; color: #ffffff; }
            .cell-pos-4 { background-color: #4caf50; color: #ffffff; }
            
            /* FinLab 風格配色：紅色漸層 (負收益) */
            .cell-neg-1 { background-color: #ffcdd2; color: #b71c1c; }
            .cell-neg-2 { background-color: #ef9a9a; color: #b71c1c; }
            .cell-neg-3 { background-color: #e57373; color: #ffffff; }
            .cell-neg-4 { background-color: #f44336; color: #ffffff; }

            /* Holdings Table - FinLab 風格 */
            .holdings-table {
                width: 100%;
                border-collapse: collapse;
                margin-top: 24px;
                background: var(--bg-primary);
                border-radius: 8px;
                overflow: hidden;
                box-shadow: var(--shadow-sm);
            }
            
            .holdings-table th { 
                background-color: var(--bg-secondary);
                font-weight: 600;
                text-align: left;
                padding: 16px;
                border-bottom: 2px solid var(--border-color);
                color: var(--text-primary);
                text-transform: uppercase;
                font-size: 12px;
                letter-spacing: 0.5px;
            }
            
            .holdings-table td {
                padding: 16px;
                border-bottom: 1px solid var(--border-color);
                transition: background 0.2s ease;
            }
            
            .holdings-table tbody tr {
                transition: all 0.2s ease;
            }
            
            .holdings-table tbody tr:hover {
                background: rgba(26, 35, 126, 0.04);
            }
            
            .holdings-table tbody tr:last-child td {
                border-bottom: none;
            }
            
            /* 持倉卡片樣式 */
            .holdings-summary-card {
                background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-light) 100%);
                color: white;
                padding: 24px;
                border-radius: 8px;
                margin-bottom: 24px;
                box-shadow: var(--shadow-md);
            }
            
            .holdings-summary-card .summary-label {
                font-size: 13px;
                opacity: 0.9;
                margin-bottom: 8px;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }
            
            .holdings-summary-card .summary-value {
                font-size: 32px;
                font-weight: 300;
                margin-bottom: 4px;
            }
            
            .holdings-summary-card .summary-sub {
                font-size: 14px;
                opacity: 0.8;
            }
            
            /* 響應式設計 */
            @media (max-width: 1024px) {
                .container {
                    padding: 32px;
                    margin: 16px;
                }
                
                .metrics-container {
                    grid-template-columns: repeat(2, 1fr);
                }
            }
            
            @media (max-width: 768px) {
                .container {
                    padding: 24px;
                    margin: 12px;
                    border-radius: 0;
                }
                
                h1 {
                    font-size: 24px;
                }
                
                .metrics-container {
                    grid-template-columns: 1fr;
                    gap: 16px;
                }
                
                .metric-value {
                    font-size: 32px;
                }
                
                .analysis-tabs {
                    flex-wrap: wrap;
                }
                
                .analysis-tab {
                    padding: 12px 16px;
                    font-size: 14px;
                }
                
                .chart-section {
                    height: 350px;
                    padding: 8px;
                }
                
                .custom-table, .holdings-table {
                    font-size: 12px;
                }
                
                .custom-table th, .custom-table td,
                .holdings-table th, .holdings-table td {
                    padding: 12px 8px;
                }
                
                .heatmap-table {
                    font-size: 10px;
                    padding: 4px;
                }
                
                .heatmap-table th, .heatmap-table td {
                    padding: 8px 4px;
                }
            }
            
            @media (max-width: 480px) {
                .container {
                    padding: 16px;
                    margin: 8px;
                }
                
                h1 {
                    font-size: 20px;
                }
                
                .header-meta {
                    font-size: 11px;
                }
                
                .metric-card {
                    padding: 16px;
                }
                
                .metric-value {
                    font-size: 28px;
                }
                
                .chart-section {
                    height: 300px;
                }
            }
            
            /* 打印樣式 */
            @media print {
                body {
                    background: white;
                    padding: 0;
                }
                
                .container {
                    box-shadow: none;
                    margin: 0;
                    padding: 20px;
                }
                
                .analysis-tabs {
                    display: none;
                }
                
                .tab-content {
                    display: block !important;
                }
                
                .chart-section {
                    height: 400px;
                    page-break-inside: avoid;
                }
            }
            
        </style>
        <script>
            // 頁面載入動畫
            document.addEventListener('DOMContentLoaded', function() {
                // 為指標卡片添加延遲動畫
                const cards = document.querySelectorAll('.metric-card');
                cards.forEach((card, index) => {
                    card.style.opacity = '0';
                    card.style.transform = 'translateY(20px)';
                    setTimeout(() => {
                        card.style.transition = 'all 0.5s ease-out';
                        card.style.opacity = '1';
                        card.style.transform = 'translateY(0)';
                    }, index * 100);
                });
                
                // 圖表載入後觸發 resize
                setTimeout(() => {
                    window.dispatchEvent(new Event('resize'));
                }, 500);
            });
            
            function switchAnalysis(tabName) {
                // 獲取所有 tab 內容和按鈕
                const tabContents = document.querySelectorAll('.tab-content');
                const tabButtons = document.querySelectorAll('.analysis-tab');
                
                // 隱藏所有 tab 內容（帶淡出動畫）
                tabContents.forEach(el => {
                    el.style.opacity = '0';
                    el.style.transform = 'translateY(10px)';
                    setTimeout(() => {
                        el.style.display = 'none';
                    }, 200);
                });
                
                // 移除所有按鈕的 active 狀態
                tabButtons.forEach(el => el.classList.remove('active'));
                
                // 顯示目標 tab 內容（帶淡入動畫）
                setTimeout(() => {
                    const targetTab = document.getElementById('tab-' + tabName);
                    targetTab.style.display = 'block';
                    setTimeout(() => {
                        targetTab.style.opacity = '1';
                        targetTab.style.transform = 'translateY(0)';
                    }, 10);
                    
                    // 添加 active 狀態到對應按鈕
                    document.querySelector('.analysis-tab[data-tab="' + tabName + '"]').classList.add('active');
                    
                    // 觸發圖表 resize
                    window.dispatchEvent(new Event('resize'));
                }, 200);
            }
            
            // 為表格行添加懸停效果增強
            document.addEventListener('DOMContentLoaded', function() {
                const tableRows = document.querySelectorAll('.custom-table tbody tr, .holdings-table tbody tr');
                tableRows.forEach(row => {
                    row.addEventListener('mouseenter', function() {
                        this.style.transition = 'all 0.2s ease';
                    });
                });
            });
        </script>
    </head>
    <body>
        <div class="container">
            <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom: 8px;">
                <div>
                    <h1>策略績效報告</h1>
                    <div class="header-meta">複利模式 • Generated: """ + pd.Timestamp.now().strftime('%Y-%m-%d %H:%M') + """</div>
                </div>
            </div>
    """

    # Unpack data
    d = data
    metrics = d['metrics']
    df_eq = d['equity']
    dates = df_eq.index.strftime('%Y-%m-%d').tolist()
    equity_vals = df_eq['Equity'].tolist()
    
    # Color logic
    cagr_color = "value-pos" if metrics['cagr'] >= 0 else "value-neg"
    mdd_style = "value-neg" # Drawdown is usually negative, we color it teal
    win_style = "value-pos" # Win ratio is positive property, color it red
    
    # Benchmarks logic
    start_date = df_eq.index[0]
    end_date = df_eq.index[-1]
    initial_val = df_eq['Equity'].iloc[0]
    
    bm_traces = []
    for ticker, bm_df in benchmarks.items():
        bm_filtered = bm_df[(bm_df.index >= start_date) & (bm_df.index <= end_date)]
        if not bm_filtered.empty:
            bm_start = bm_filtered['Close'].iloc[0]
            bm_norm = (bm_filtered['Close'] / bm_start) * initial_val
            bm_traces.append({
                'x': bm_filtered.index.strftime('%Y-%m-%d').tolist(),
                'y': bm_norm.tolist(),
                'name': ticker
            })
    
    # Charts Data
    yearly_rets = d['yearly_returns']
    years = yearly_rets.index.year.tolist()
    y_ret_vals = yearly_rets.values.tolist()
    # FinLab 風格配色：綠色 (#4caf50) 用於正收益，紅色 (#f44336) 用於負收益
    y_colors = ['#4caf50' if v >= 0 else '#f44336' for v in y_ret_vals]
    # 格式化年報酬文字標籤
    y_text_labels = [f'{v:+.2f}%' for v in y_ret_vals]

    html_content += f"""
        <!-- Metrics Header -->
        <div class="metrics-container">
            <div class="metric-card">
                <div class="metric-title">年化報酬率</div>
                <div class="metric-subtitle">cagr</div>
                <div class="metric-value {cagr_color}">{metrics['cagr']:+.1f}%</div>
            </div>
            <div class="metric-card">
                <div class="metric-title">年化夏普值</div>
                <div class="metric-subtitle">daily sharpe * √252</div>
                <div class="metric-value">{metrics['sharpe']:.2f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-title">最大下跌</div>
                <div class="metric-subtitle">max drawdown</div>
                <div class="metric-value {mdd_style}">{metrics['mdd']:.1f}%</div>
            </div>
            <div class="metric-card">
                <div class="metric-title">勝率</div>
                <div class="metric-subtitle">win ratio ({metrics['win_count']}/{metrics['total_trades']})</div>
                <div class="metric-value {win_style}">{metrics['win_ratio']:.1f}%</div>
            </div>
        </div>

        <!-- Main Chart -->
        <div class="chart-section" id="main-chart"></div>
        <script>
            var trace_eq = {{
                x: {dates},
                y: {equity_vals},
                type: 'scatter',
                mode: 'lines',
                name: '策略',
                line: {{color: '#1a237e', width: 3}},
                fill: 'tozeroy',
                fillcolor: 'rgba(26, 35, 126, 0.08)'
            }};
            
            var traces = [trace_eq];
            
            // Benchmarks
            var bm_data = {json.dumps(bm_traces)};
            bm_data.forEach(function(bm) {{
                traces.push({{
                    x: bm.x,
                    y: bm.y,
                    type: 'scatter',
                    mode: 'lines',
                    name: bm.name,
                    line: {{width: 2, color: '#9e9e9e', dash: 'dash'}}
                }});
            }});

            var layout = {{
                title: {{
                    text: '',
                    font: {{size: 18, color: '#212121'}}
                }},
                autosize: true,
                margin: {{l: 60, r: 30, t: 20, b: 50}},
                xaxis: {{
                    showgrid: true,
                    gridcolor: '#f5f5f5',
                    zeroline: false,
                    tickformat: '%Y',
                    dtick: 'M24',
                    showline: true,
                    linecolor: '#e0e0e0',
                    tickfont: {{size: 11, color: '#757575'}},
                    titlefont: {{size: 12, color: '#757575'}}
                }},
                yaxis: {{
                    showgrid: true,
                    gridcolor: '#f5f5f5',
                    zeroline: false,
                    showline: true,
                    linecolor: '#e0e0e0',
                    tickfont: {{size: 11, color: '#757575'}},
                    titlefont: {{size: 12, color: '#757575'}}
                }},
                showlegend: true,
                legend: {{
                    orientation: 'h',
                    x: 0,
                    y: 1.02,
                    xanchor: 'left',
                    yanchor: 'bottom',
                    font: {{size: 12, color: '#757575'}},
                    bgcolor: 'rgba(255,255,255,0.8)'
                }},
                plot_bgcolor: '#ffffff',
                paper_bgcolor: '#ffffff',
                font: {{family: 'Roboto, sans-serif', color: '#212121'}},
                hovermode: 'x unified',
                hoverlabel: {{
                    bgcolor: 'rgba(26, 35, 126, 0.9)',
                    font: {{color: 'white', size: 12}}
                }}
            }};
            
            Plotly.newPlot('main-chart', traces, layout, {{responsive: true, displayModeBar: false}});
        </script>

        <!-- Analysis Tabs -->
        <div class="analysis-tabs">
            <button class="analysis-tab active" data-tab="return" onclick="switchAnalysis('return')">報酬分析</button>
            <button class="analysis-tab" data-tab="risk" onclick="switchAnalysis('risk')">風險分析</button>
            <button class="analysis-tab" data-tab="list" onclick="switchAnalysis('list')">選股清單</button>
        </div>

        <!-- Tab: Return Analysis -->
        <div id="tab-return" class="tab-content active">
            <h3>年報酬</h3>
            <div id="year-chart" style="height: 300px;"></div>
            <script>
                var trace_year = {{
                    x: {years},
                    y: {y_ret_vals},
                    type: 'bar',
                    marker: {{color: {json.dumps(y_colors)}}},
                    text: {json.dumps(y_text_labels)},
                    textposition: 'outside',
                    textangle: 0,
                    textfont: {{
                        size: 12,
                        color: '#212121',
                        family: 'Roboto, sans-serif'
                    }},
                    cliponaxis: false
                }};
                var layout_year = {{
                    autosize: true,
                    margin: {{l: 50, r: 20, t: 80, b: 40}},
                    xaxis: {{
                        tickmode: 'linear',
                        type: 'category',
                        showgrid: true,
                        gridcolor: '#f5f5f5',
                        showline: true,
                        linecolor: '#e0e0e0',
                        tickfont: {{size: 11, color: '#757575'}}
                    }},
                    yaxis: {{
                        showgrid: true,
                        gridcolor: '#f5f5f5',
                        showline: true,
                        linecolor: '#e0e0e0',
                        tickfont: {{size: 11, color: '#757575'}},
                        range: [null, null],
                        autorange: true,
                        fixedrange: false
                    }},
                    plot_bgcolor: '#ffffff',
                    paper_bgcolor: '#ffffff',
                    font: {{family: 'Roboto, sans-serif', size: 12, color: '#212121'}},
                    hoverlabel: {{
                        bgcolor: 'rgba(26, 35, 126, 0.9)',
                        font: {{color: 'white', size: 11}}
                    }}
                }};
                Plotly.newPlot('year-chart', [trace_year], layout_year, {{responsive: true, displayModeBar: false}});
            </script>
            
            <h3 style="margin-top:32px;">月報酬</h3>
            <style>
                /* Heatmap Color Utils */
                .hm-val {{ font-family: monospace; }}
            </style>
            """
    
    # Generate Monthly Heatmap Table
    monthly_html = '<table class="heatmap-table"><tr><th></th>'
    for m in range(1, 13): monthly_html += f'<th>{m}</th>'
    monthly_html += '<th>YTD</th></tr>'
    
    if not d['monthly_returns'].empty:
            monthly_df = d['monthly_returns'].to_frame('Return')
            monthly_df['Year'] = monthly_df.index.year
            monthly_df['Month'] = monthly_df.index.month
            pivot = monthly_df.pivot_table(values='Return', index='Year', columns='Month', aggfunc='sum')
            
            yearly_raw = d['yearly_returns']
            yearly_dict = {dt.year: val for dt, val in yearly_raw.items()}

            for year in sorted(pivot.index, reverse=True):
                monthly_html += f'<tr><td style="background:#fff; color:#666; font-weight:500;"><b>{year}</b></td>'
                for m in range(1, 13):
                    val = pivot.loc[year, m] if m in pivot.columns and not pd.isna(pivot.loc[year, m]) else None
                    if val is not None:
                        # Assign color class based on magnitude
                        cls = ""
                        if val > 0:
                            if val > 10: cls = "cell-pos-4"
                            elif val > 5: cls = "cell-pos-3"
                            elif val > 2: cls = "cell-pos-2"
                            else: cls = "cell-pos-1"
                        else:
                            if val < -10: cls = "cell-neg-4"
                            elif val < -5: cls = "cell-neg-3"
                            elif val < -2: cls = "cell-neg-2"
                            else: cls = "cell-neg-1"
                        monthly_html += f'<td class="heatmap-cell {cls}">{val:+.1f}%</td>'
                    else:
                        monthly_html += '<td style="color:#f0f0f0">-</td>'
                
                # YTD
                y_val = yearly_dict.get(year, 0)
                y_cls = "value-pos" if y_val >= 0 else "value-neg"
                monthly_html += f'<td style="font-weight:bold" class="{y_cls}">{y_val:+.1f}%</td></tr>'
                
    monthly_html += '</table>'
    html_content += monthly_html + "</div>" # End Return Tab

    # Tab: Risk Analysis
    dd_dates = d['drawdown'].index.strftime('%Y-%m-%d').tolist()
    dd_vals = d['drawdown'].tolist()
    
    html_content += f"""
        <div id="tab-risk" class="tab-content">
            <h3>下跌幅度 (Drawdown)</h3>
            <div id="dd-chart" style="height: 350px;"></div>
            <script>
                var trace_dd = {{
                    x: {dd_dates},
                    y: {dd_vals},
                    type: 'scatter',
                    mode: 'lines',
                    fill: 'tozeroy',
                    line: {{color: '#f44336', width: 2}},
                    fillcolor: 'rgba(244, 67, 54, 0.1)'
                }};
                var layout_dd = {{
                    autosize: true,
                    margin: {{l: 60, r: 30, t: 20, b: 50}},
                    xaxis: {{
                        gridcolor: '#f5f5f5',
                        tickformat: '%Y',
                        dtick: 'M24',
                        showline: true,
                        linecolor: '#e0e0e0',
                        tickfont: {{size: 11, color: '#757575'}}
                    }},
                    yaxis: {{
                        gridcolor: '#f5f5f5',
                        title: {{
                            text: '下跌幅度 (%)',
                            font: {{size: 12, color: '#757575'}}
                        }},
                        showline: true,
                        linecolor: '#e0e0e0',
                        tickfont: {{size: 11, color: '#757575'}}
                    }},
                    plot_bgcolor: '#ffffff',
                    paper_bgcolor: '#ffffff',
                    font: {{family: 'Roboto, sans-serif', color: '#212121'}},
                    hoverlabel: {{
                        bgcolor: 'rgba(26, 35, 126, 0.9)',
                        font: {{color: 'white', size: 11}}
                    }}
                }};
                Plotly.newPlot('dd-chart', [trace_dd], layout_dd, {{responsive: true, displayModeBar: false}});
            </script>
            
            <h3 style="margin-top:32px;">歷史最大下跌風險</h3>
            <table class="custom-table" style="margin-top:10px;">
                <thead>
                    <tr>
                        <th>開始時間</th><th>最低點時間</th><th>恢復時間</th><th>虧損程度</th><th>持續天數</th><th>恢復天數</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    # Generate drawdown periods table rows
    drawdown_periods = d.get('drawdown_periods', [])
    if drawdown_periods and len(drawdown_periods) > 0:
        # 顯示前 10 個最大的下跌期間（已經按下跌幅度排序）
        top_drawdowns = drawdown_periods[:10]
        for period in top_drawdowns:
            start_date_str = period['start_date'].strftime('%Y-%m-%d')
            end_date_str = period['end_date'].strftime('%Y-%m-%d')
            
            # 處理恢復時間
            if period['recovery_date'] is not None:
                recovery_str = period['recovery_date'].strftime('%Y-%m-%d')
                recovery_days_str = f"{period['recovery_days']} 天"
            else:
                recovery_str = '未恢復'
                recovery_days_str = '-'
            
            drawdown_pct = period['drawdown_pct']
            duration_days = period['duration_days']
            
            # 根據下跌幅度決定顏色和樣式
            if abs(drawdown_pct) >= 20:
                drawdown_class = "value-neg"
                drawdown_style = "font-weight:700; font-size:14px;"
            elif abs(drawdown_pct) >= 10:
                drawdown_class = "value-neg"
                drawdown_style = "font-weight:600;"
            else:
                drawdown_class = ""
                drawdown_style = "font-weight:500;"
            
            html_content += f"""
                    <tr>
                        <td>{start_date_str}</td>
                        <td>{end_date_str}</td>
                        <td>{recovery_str}</td>
                        <td class="{drawdown_class}" style="{drawdown_style}">{drawdown_pct:.2f}%</td>
                        <td>{duration_days} 天</td>
                        <td>{recovery_days_str}</td>
                    </tr>
            """
    else:
        html_content += f"""
                    <tr><td colspan="6" style="text-align:center; color:var(--text-secondary); padding:40px;">
                        無足夠數據計算下跌期間<br/>
                        <span style="font-size:12px; margin-top:8px; display:block;">目前最大回撤: <span class="value-neg" style="font-weight:600;">{metrics['mdd']:.1f}%</span></span>
                    </td></tr>
        """
    
    html_content += """
                </tbody>
            </table>
        </div>
    """

    # Tab: Stock List (Current Holdings)
    holdings_html = ""
    holdings_data = d['holdings']
    if holdings_data and 'holdings' in holdings_data:
        holdings_html += f"""
        <div class="holdings-summary-card">
            <div class="summary-label">總資產</div>
            <div class="summary-value">${holdings_data.get('total_equity',0):,.0f}</div>
            <div class="summary-sub">現金: ${holdings_data.get('cash',0):,.0f}</div>
        </div>
        <table class="holdings-table">
            <thead>
                <tr>
                    <th>代碼</th><th>股數</th><th>成本</th><th>現價</th><th>市值</th><th>損益 $</th><th>損益 %</th><th>權重</th>
                </tr>
            </thead>
            <tbody>
        """
        for h in holdings_data['holdings']:
            pnl = h.get('pnl', 0)
            pnl_class = "value-pos" if pnl >= 0 else "value-neg"
            holdings_html += f"""
            <tr>
                <td><b>{h['ticker']}</b></td>
                <td>{h['qty']}</td>
                <td>${h['avg_cost']:.2f}</td>
                <td>${h['current_price']:.2f}</td>
                <td>${h['value']:,.0f}</td>
                <td class="{pnl_class}">${h['pnl']:,.0f}</td>
                <td class="{pnl_class}">{h['pnl_pct']:+.1f}%</td>
                <td>{h['weight']:.1f}%</td>
            </tr>
            """
        holdings_html += "</tbody></table>"
    else:
        holdings_html = "<p>無目前持倉數據</p>"

    html_content += f"""
        <div id="tab-list" class="tab-content">
            <h3>當前持倉 (選股清單)</h3>
            {holdings_html}
        </div>
    </div>
    </body>
    </html>
    """
    
    with open('strategy_report_final.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("Report generated: strategy_report_final.html")

def generate_report():
    generate_comparison_report()

if __name__ == "__main__":
    generate_report()
