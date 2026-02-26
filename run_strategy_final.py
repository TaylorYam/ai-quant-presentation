"""
Run Strategy Final - Live Portfolio Simulation
Uses locally stored data (run update_data.py first to fetch latest)
"""
import config_final as config
from portfolio_backtester_final import PortfolioBacktesterFinal
from report_generator_final import generate_report

def main():
    print("=" * 60)
    print("FINAL VERSION - Live Portfolio Simulation")
    print("=" * 60)
    
    # Get date range
    start_date = config.START_DATE
    end_date = config.END_DATE  # None = auto-detect latest from local data
    
    print(f"\nStart Date: {start_date}")
    print(f"End Date: {'Auto-detect (latest in data)' if end_date is None else end_date}")
    print(f"LIVE_MODE: {config.LIVE_MODE}")
    print("-" * 60)
    
    # [OPTIMIZATION] Preload SPY and SSO data once (避免重複載入)
    print("\n[0/3] Preloading benchmark data...")
    import os
    import utils
    spy_path = os.path.join(config.DATA_DIR, 'SPY.csv')
    sso_path = os.path.join(config.DATA_DIR, 'SSO.csv')
    spy_df = utils.load_benchmark_data(spy_path)
    sso_df = utils.load_benchmark_data(sso_path)
    
    # [OPTIMIZATION] Create shared SelectionEngine (stock data 只載入一次)
    from selection import SelectionEngine
    shared_selector = SelectionEngine()
    
    # Run Compound Interest Backtest (Only)
    print("\n[1/2] Running Compound Interest Backtest...")
    bt_compound = PortfolioBacktesterFinal(
        start_date=start_date,
        end_date=end_date,
        initial_capital=config.INITIAL_CASH,
        compounding=True,
        report_suffix="_compound",
        selector=shared_selector,  # Share selector (reuse cache!)
        spy_df=spy_df,            # Preloaded
        sso_df=sso_df             # Preloaded
    )
    bt_compound.run()
    
    # Export Rebalance Snapshots to Excel
    print("\n[Bonus] Exporting rebalance day snapshots to Excel...")
    bt_compound.export_rebalance_excel('rebalance_holdings_final.xlsx')
    
    # Generate Report
    print("\n[2/2] Generating HTML Report...")
    generate_report()
    
    print("\n" + "=" * 60)
    print("COMPLETE! Open strategy_report_final.html to view results.")
    print("=" * 60)

if __name__ == "__main__":
    main()
