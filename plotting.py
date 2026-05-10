from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


def plot_equity_curve(equity: pd.DataFrame, output_path: str | Path) -> None:
    if equity.empty:
        return
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    plt.figure(figsize=(10, 5))
    plt.plot(equity["timestamp"], equity["equity"])
    plt.title("SPY Option Backtest Equity Curve")
    plt.xlabel("Date")
    plt.ylabel("Account Equity")
    plt.tight_layout()
    plt.savefig(output_path, dpi=150)
    plt.close()
