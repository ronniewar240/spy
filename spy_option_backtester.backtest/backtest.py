from dataclasses import dataclass
from datetime import time

import pandas as pd


@dataclass
class BacktestConfig:
    contracts: int = 1
    entry_time: str = "09:35"
    exit_time: str = "15:55"
    stop_loss: float = 0.35       # 35% premium loss
    take_profit: float = 0.80     # 80% premium gain
    starting_cash: float = 10_000
    commission_per_contract: float = 0.65
    entry_limit_price: float | None = None  # If set, buy only when option low <= this price


def _parse_hhmm(value: str) -> time:
    hh, mm = value.split(":")
    return time(int(hh), int(mm))


def _price_at_or_before(candles: pd.DataFrame, ts):
    """Return underlying close at or before timestamp, or None if unavailable."""
    if candles is None or candles.empty:
        return None
    prior = candles[candles["timestamp"] <= ts]
    if prior.empty:
        return None
    return float(prior.iloc[-1]["close"])


def _find_entry(day: pd.DataFrame, entry_t: time, exit_t: time, entry_limit_price: float | None):
    """Find entry candle.

    Market mode: first candle at/after entry_time, fill at close.
    Limit mode: first candle at/after entry_time and before exit_time where low <= limit.
    Fill price is the limit price for conservative/simple backtesting.
    """
    candidates = day[(day["clock"] >= entry_t) & (day["clock"] <= exit_t)]
    if candidates.empty:
        return None, None, "no_entry_window"

    if entry_limit_price is None:
        entry = candidates.iloc[0]
        return entry, float(entry["close"]), "market"

    fills = candidates[candidates["low"].astype(float) <= float(entry_limit_price)]
    if fills.empty:
        return None, None, "limit_not_filled"

    entry = fills.iloc[0]
    return entry, float(entry_limit_price), "limit"


def run_long_option_backtest(candles: pd.DataFrame, config: BacktestConfig, underlying_candles: pd.DataFrame | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    """One trade per day: buy option near entry_time or at limit, sell on SL/TP/exit_time.

    If entry_limit_price is provided, the trade only enters when a candle's low touches the limit price.
    If underlying_candles is provided, entry/exit SPY prices are added to the trade log.
    """
    if candles.empty:
        return pd.DataFrame(), pd.DataFrame()

    df = candles.copy()
    df["date"] = df["timestamp"].dt.date
    df["clock"] = df["timestamp"].dt.time

    entry_t = _parse_hhmm(config.entry_time)
    exit_t = _parse_hhmm(config.exit_time)

    cash = config.starting_cash
    trades = []
    equity_rows = []

    for session_date, day in df.groupby("date"):
        day = day.sort_values("timestamp").reset_index(drop=True)

        entry, entry_price, entry_type = _find_entry(day, entry_t, exit_t, config.entry_limit_price)
        if entry is None:
            trades.append({
                "date": session_date,
                "entry_type": entry_type,
                "entry_time_requested": config.entry_time,
                "entry_time": None,
                "fill_time": None,
                "minutes_to_fill": None,
                "entry_price": config.entry_limit_price,
                "exit_time": None,
                "exit_price": None,
                "underlying_entry_price": None,
                "underlying_exit_price": None,
                "contracts": config.contracts,
                "gross_pnl": 0.0,
                "fees": 0.0,
                "net_pnl": 0.0,
                "exit_reason": entry_type,
                "cash_after_trade": cash,
            })
            continue

        stop_price = entry_price * (1 - config.stop_loss)
        target_price = entry_price * (1 + config.take_profit)
        exit_row = None
        reason = "time_exit"

        for _, row in day[day["timestamp"] > entry["timestamp"]].iterrows():
            if row["clock"] > exit_t:
                break
            if float(row["low"]) <= stop_price:
                exit_row = row.copy()
                exit_row["close"] = stop_price
                reason = "stop_loss"
                break
            if float(row["high"]) >= target_price:
                exit_row = row.copy()
                exit_row["close"] = target_price
                reason = "take_profit"
                break
            if row["clock"] >= exit_t:
                exit_row = row
                break

        if exit_row is None:
            after_entry = day[(day["timestamp"] >= entry["timestamp"]) & (day["clock"] <= exit_t)]
            exit_row = after_entry.iloc[-1] if not after_entry.empty else day.iloc[-1]
            reason = "time_exit" if not after_entry.empty else "end_of_data"

        exit_price = float(exit_row["close"])
        gross_pnl = (exit_price - entry_price) * 100 * config.contracts
        fees = config.commission_per_contract * config.contracts * 2
        net_pnl = gross_pnl - fees
        cash += net_pnl

        underlying_entry_price = _price_at_or_before(underlying_candles, entry["timestamp"])
        underlying_exit_price = _price_at_or_before(underlying_candles, exit_row["timestamp"])
        minutes_to_fill = (entry["timestamp"] - pd.Timestamp(f"{session_date} {config.entry_time}", tz=entry["timestamp"].tz)).total_seconds() / 60

        trades.append({
            "date": session_date,
            "entry_type": entry_type,
            "entry_time_requested": config.entry_time,
            "entry_time": entry["timestamp"],
            "fill_time": entry["timestamp"],
            "minutes_to_fill": round(minutes_to_fill, 2),
            "entry_price": entry_price,
            "exit_time": exit_row["timestamp"],
            "exit_price": exit_price,
            "underlying_entry_price": underlying_entry_price,
            "underlying_exit_price": underlying_exit_price,
            "contracts": config.contracts,
            "gross_pnl": gross_pnl,
            "fees": fees,
            "net_pnl": net_pnl,
            "exit_reason": reason,
            "cash_after_trade": cash,
        })
        equity_rows.append({"timestamp": exit_row["timestamp"], "equity": cash})

    return pd.DataFrame(trades), pd.DataFrame(equity_rows)
