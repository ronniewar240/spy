import argparse
from pathlib import Path

from dotenv import load_dotenv

from spy_option_backtester.backtest import BacktestConfig, run_long_option_backtest
from spy_option_backtester.massive_client import MassiveClient
from spy_option_backtester.plotting import plot_equity_curve
from spy_option_backtester.tickers import build_option_ticker


def parse_args():
    parser = argparse.ArgumentParser(description="Backtest a SPY option contract using Massive historical candles.")
    parser.add_argument("--underlying", default="SPY")
    parser.add_argument("--expiry", required=True, help="YYYY-MM-DD")
    parser.add_argument("--right", choices=["C", "P"], required=True)
    parser.add_argument("--strike", type=float, required=True)
    parser.add_argument("--from-date", required=True, help="YYYY-MM-DD")
    parser.add_argument("--to-date", required=True, help="YYYY-MM-DD")
    parser.add_argument("--interval", type=int, default=1)
    parser.add_argument("--timespan", default="minute", choices=["minute", "hour", "day", "week", "month", "quarter", "year"])
    parser.add_argument("--entry-time", default="09:35")
    parser.add_argument("--exit-time", default="15:55")
    parser.add_argument("--stop-loss", type=float, default=0.35)
    parser.add_argument("--take-profit", type=float, default=0.80)
    parser.add_argument("--contracts", type=int, default=1)
    parser.add_argument("--starting-cash", type=float, default=10000)
    return parser.parse_args()


def main():
    load_dotenv()
    args = parse_args()

    ticker = build_option_ticker(args.underlying, args.expiry, args.right, args.strike)
    print(f"Fetching {ticker}...")

    client = MassiveClient()
    candles = client.get_aggregates(
        ticker=ticker,
        multiplier=args.interval,
        timespan=args.timespan,
        from_date=args.from_date,
        to_date=args.to_date,
    )

    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)
    candles.to_csv(output_dir / "candles.csv", index=False)

    config = BacktestConfig(
        contracts=args.contracts,
        entry_time=args.entry_time,
        exit_time=args.exit_time,
        stop_loss=args.stop_loss,
        take_profit=args.take_profit,
        starting_cash=args.starting_cash,
    )
    trades, equity = run_long_option_backtest(candles, config)
    trades.to_csv(output_dir / "trades.csv", index=False)
    equity.to_csv(output_dir / "equity_curve.csv", index=False)
    plot_equity_curve(equity, output_dir / "equity_curve.png")

    print("Done.")
    print(f"Candles: {len(candles)}")
    print(f"Trades: {len(trades)}")
    if not trades.empty:
        print(trades[["date", "entry_price", "exit_price", "net_pnl", "exit_reason", "cash_after_trade"]])
        print(f"Final equity: ${trades['cash_after_trade'].iloc[-1]:,.2f}")


if __name__ == "__main__":
    main()
