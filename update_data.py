"""
Update All Data
Standalone script to fetch latest data from Yahoo Finance for all stocks
Handles both .csv and .txt formats
"""
import os
import pandas as pd
from datetime import datetime, timedelta

# Use the base config to get DATA_DIR
try:
    import config_final as config
    DATA_DIR = config.DATA_DIR
except:
    DATA_DIR = 'data'

try:
    import yfinance as yf
    YF_AVAILABLE = True
except ImportError:
    YF_AVAILABLE = False
    print("ERROR: yfinance not installed. Run: pip install yfinance")
    exit(1)

def load_existing_data(filepath):
    """Load existing data file, handling both .csv and .txt formats"""
    if not os.path.exists(filepath):
        return None
    
    try:
        df = pd.read_csv(filepath)
        
        # Detect format and standardize
        if 'date' in df.columns:  # .txt format (lowercase)
            col_map = {
                'date': 'Date',
                'o': 'Open',
                'h': 'High', 
                'l': 'Low',
                'c': 'Close',
                'adj_c': 'Adj Close',
                'vol': 'Volume'
            }
            df = df.rename(columns=col_map)
            if 'ticker' in [c.lower() for c in df.columns]:
                df = df.drop(columns=['ticker'], errors='ignore')
        
        # Legacy uppercase columns
        col_map2 = {'CLOSE': 'Close', 'OPEN': 'Open', 'HIGH': 'High', 'LOW': 'Low', 'VOLUME': 'Volume'}
        df = df.rename(columns=col_map2)
        
        # Remove duplicate columns
        df = df.loc[:, ~df.columns.duplicated()]
        
        # Parse date
        df['Date'] = pd.to_datetime(df['Date'])
        
        return df
    except Exception as e:
        print(f"  Error loading {filepath}: {e}")
        return None

def get_last_date(filepath):
    """Get last date in file"""
    df = load_existing_data(filepath)
    if df is not None and 'Date' in df.columns:
        return df['Date'].max()
    return None

def save_data(df, filepath):
    """Save data in appropriate format"""
    ext = os.path.splitext(filepath)[1].lower()
    
    if ext == '.txt':
        # Convert to .txt format
        ticker = os.path.splitext(os.path.basename(filepath))[0]
        out_df = pd.DataFrame()
        out_df['ticker'] = ticker
        out_df['date'] = df['Date'].dt.strftime('%Y/%m/%d')
        out_df['o'] = df['Open']
        out_df['h'] = df['High']
        out_df['l'] = df['Low']
        out_df['c'] = df['Close']
        out_df['adj_c'] = df.get('Adj Close', df['Close'])
        out_df['vol'] = df['Volume'].astype(int)
        out_df.to_csv(filepath, index=False)
    else:
        # Standard CSV format
        standard_cols = ['Date', 'Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume']
        out_cols = [c for c in standard_cols if c in df.columns]
        df[out_cols].to_csv(filepath, index=False)

def update_ticker(ticker, filepath):
    """Update a single ticker from Yahoo Finance"""
    last_date = get_last_date(filepath)
    
    if last_date is None:
        if os.path.exists(filepath):
            print(f"  {ticker}: Could not read existing file")
            return False
        start_date = "2015-01-01"
        print(f"  {ticker}: No existing data, fetching from {start_date}...")
    else:
        # Check if already up to date
        if last_date.date() >= (datetime.now() - timedelta(days=1)).date():
            # Already up to date, skip
            return True
        start_date = (last_date + timedelta(days=1)).strftime('%Y-%m-%d')
    
    end_date = datetime.now().strftime('%Y-%m-%d')
    
    try:
        stock = yf.Ticker(ticker)
        new_data = stock.history(start=start_date, end=end_date, auto_adjust=False)
        
        if new_data.empty:
            return True  # No new data, but not an error
        
        # Process new data
        new_data = new_data.reset_index()
        new_data['Date'] = pd.to_datetime(new_data['Date']).dt.tz_localize(None)
        
        # Keep standard columns
        standard_cols = ['Date', 'Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume']
        new_data = new_data[[c for c in standard_cols if c in new_data.columns]]
        
        if os.path.exists(filepath):
            existing_df = load_existing_data(filepath)
            if existing_df is not None:
                combined = pd.concat([existing_df, new_data], ignore_index=True)
                combined = combined.drop_duplicates(subset=['Date'], keep='last')
                combined = combined.sort_values('Date')
                save_data(combined, filepath)
                print(f"  {ticker}: +{len(new_data)} rows (total: {len(combined)})")
            else:
                return False
        else:
            save_data(new_data, filepath)
            print(f"  {ticker}: Created with {len(new_data)} rows")
        
        return True
        
    except Exception as e:
        error_msg = str(e)
        if "404" not in error_msg and "delisted" not in error_msg:
            print(f"  {ticker}: ERROR - {error_msg[:50]}")
        return False

def main():
    print("\n" + "="*60)
    print("DATA UPDATER - Fetch Latest from Yahoo Finance")
    print("="*60)
    print(f"\nData directory: {DATA_DIR}")
    
    # 1. Update core benchmarks
    print("\n[1/3] Updating core tickers (SPY, SSO, QQQ)...")
    for ticker in ['SPY', 'SSO', 'QQQ']:
        filepath = os.path.join(DATA_DIR, f'{ticker}.csv')
        last = get_last_date(filepath)
        if last and last.date() >= (datetime.now() - timedelta(days=1)).date():
            print(f"  {ticker}: Already up to date ({last.date()})")
        else:
            update_ticker(ticker, filepath)
    
    # 2. Update all .txt stock files
    print("\n[2/3] Updating individual stocks (.txt files)...")
    if os.path.exists(DATA_DIR):
        txt_files = [f for f in os.listdir(DATA_DIR) if f.endswith('.txt')]
        
        total = len(txt_files)
        updated = 0
        skipped = 0
        failed = 0
        
        for i, filename in enumerate(txt_files):
            ticker = filename.replace('.txt', '')
            filepath = os.path.join(DATA_DIR, filename)
            
            last = get_last_date(filepath)
            if last and last.date() >= (datetime.now() - timedelta(days=1)).date():
                skipped += 1
                continue
            
            print(f"  [{i+1}/{total}] {ticker}: Fetching...")
            if update_ticker(ticker, filepath):
                updated += 1
            else:
                failed += 1
        
        print(f"\n  Summary: {updated} updated, {skipped} skipped (up-to-date), {failed} failed")
    
    # 3. Verify
    print("\n[3/3] Verification...")
    spy_last = get_last_date(os.path.join(DATA_DIR, 'SPY.csv'))
    print(f"  SPY last date: {spy_last.date() if spy_last else 'N/A'}")
    
    # Check a sample stock
    sample_txt = os.path.join(DATA_DIR, 'AAPL.txt')
    if os.path.exists(sample_txt):
        aapl_last = get_last_date(sample_txt)
        print(f"  AAPL last date: {aapl_last.date() if aapl_last else 'N/A'}")
    
    print("\n" + "="*60)
    print("UPDATE COMPLETE!")
    print("="*60)
    print("\nYou can now run the backtest with:")
    print("  python run_strategy_final.py")

if __name__ == "__main__":
    main()
