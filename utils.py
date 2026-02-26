import pandas as pd
import numpy as np
import os

def load_data(file_path):
    """
    通用資料讀取函數
    支援:
    1. CSV with header
    2. CSV without header (names provided)
    3. Ticker column removal
    4. Mapping 'o', 'c' etc to 'Open', 'Close'
    """
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return None
        
    try:
        # 預讀前幾行判斷格式
        with open(file_path, 'r') as f:
            first_line = f.readline()
            
        common_headers = ['date', 'open', 'high', 'low', 'close', 'volume', 'Date', 'Open']
        has_header = any(h in first_line.lower() for h in common_headers)
        
        if has_header:
            df = pd.read_csv(file_path)
        else:
            # 假設無 header，且格式為 Ticker, Date, Open, High, Low, Close, Volume
            df = pd.read_csv(file_path, header=None, names=['Ticker', 'Date', 'Open', 'High', 'Low', 'Close', 'Volume'])
            
        # 標準化欄位名稱
        # 移除前後空白
        df.rename(columns=lambda x: x.strip(), inplace=True)
        
        # 確保欄位名稱符合 backtesting 要求 (首字母大寫)
        rename_map = {
            'open': 'Open', 'high': 'High', 'low': 'Low', 'close': 'Close', 'volume': 'Volume',
            'Open': 'Open', 'High': 'High', 'Low': 'Low', 'Close': 'Close', 'Volume': 'Volume',
            'o': 'Open', 'h': 'High', 'l': 'Low', 'c': 'Close', 'vol': 'Volume', 'adj_c': 'Adj Close'
        }
        df.rename(columns=rename_map, inplace=True)
        
        # 移除不需要的 Ticker 欄位 (如果有)
        if 'Ticker' in df.columns:
            df.drop(columns=['Ticker'], inplace=True)
        if 'ticker' in df.columns:
            df.drop(columns=['ticker'], inplace=True)

        # 簡單檢查必要欄位
        required = ['Open', 'High', 'Low', 'Close']
        if not all(col in df.columns for col in required):
            raise ValueError(f"Missing required columns: {set(required) - set(df.columns)}")
            
        # 日期處理
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'])
            df.set_index('Date', inplace=True)
        elif 'date' in df.columns:
            df['date'] = pd.to_datetime(df['date'])
            df.set_index('date', inplace=True)
            df.index.name = 'Date'
        else:
            # 嘗試使用 index
            try:
                df.index = pd.to_datetime(df.index)
                df.index.name = 'Date'
            except:
                pass
                
        # 排序
        df.sort_index(inplace=True)
        
        # 檢查是否為空
        if df.empty:
            return None
            
        return df.dropna()
        
    except Exception as e:
        print(f"Error loading {file_path}: {e}")
        return None

def load_benchmark_data(file_path):
    """
    專門讀取 Benchmark (SPY, SSO) 資料，只需要 Date 和 Close
    """
    try:
        df = pd.read_csv(file_path)
        
        # 處理欄位名稱 (標準化為 Close)
        df.rename(columns=lambda x: x.strip().title(), inplace=True) # e.g. CLOSE -> Close
        
        # 尋找日期
        if 'Date' in df.columns:
            # 支援 YYYY/MM/DD 格式
            df['Date'] = pd.to_datetime(df['Date'], format='mixed') 
            df.set_index('Date', inplace=True)
            df.sort_index(inplace=True)
            
            if not df.empty:
                print(f"Loaded {os.path.basename(file_path)}: {df.index[0].date()} to {df.index[-1].date()}")
        else:
            raise ValueError("Date column missing in benchmark file")
            
        if 'Close' not in df.columns:
            raise ValueError("Close column missing in benchmark file")
            
        return df[['Close']] # 只回傳 Close
    except Exception as e:
        print(f"Error loading benchmark {file_path}: {e}")
        return None

def get_data_files(directory=None):
    if directory is None:
        import config
        directory = config.DATA_DIR
    
    files = []
    for f in os.listdir(directory):
        if f.endswith(".csv") or f.endswith(".txt"):
            files.append(os.path.join(directory, f))
    return files
