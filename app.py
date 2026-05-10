from __future__ import annotations

from pathlib import Path
from datetime import date

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from spy_option_backtester.backtest import BacktestConfig, run_long_option_backtest
from spy_option_backtester.massive_client import MassiveClient
from spy_option_backtester.plotting import plot_equity_curve
from spy_option_backtester.tickers import build_option_ticker
from ibkr_live import get_live_quotes

load_dotenv()

st.set_page_config(page_title="SPY Options Backtester + IBKR Live", page_icon="📈", layout="wide")

st.title("SPY Options Backtester + IBKR Live P&L")
st.caption("One project: Massive historical backtesting + native IBKR live P&L tracking. Read-only. No trades are placed.")

backtest_tab, live_tab, setup_tab = st.tabs(["Historical Backtest", "IBKR Live P&L", "Setup Notes"])


def fmt_money(value):
    return "N/A" if value is None else f"${float(value):,.2f}"


with backtest_tab:
    st.subheader("Massive Historical Options Backtest")
    st.caption("Backtest single-leg options using Massive historical aggregate candles. Includes SPY price at entry/exit, limit-fill time, and SL/TP buttons.")

    with st.sidebar:
        st.header("Backtest Settings")
        st.subheader("Contract")
        bt_underlying = st.text_input("Underlying", value="SPY", key="bt_underlying").upper().strip()
        bt_expiry = st.date_input("Expiry", value=date(2026, 5, 8), key="bt_expiry")
        bt_right_label = st.radio("Option Type", ["Call", "Put"], horizontal=True, key="bt_right_label")
        bt_right = "C" if bt_right_label == "Call" else "P"
        bt_strike = st.number_input("Strike", min_value=0.0, value=735.0, step=1.0, key="bt_strike")

        st.subheader("Historical Range")
        bt_from_date = st.date_input("From date", value=date(2026, 5, 1), key="bt_from")
        bt_to_date = st.date_input("To date", value=date(2026, 5, 8), key="bt_to")
        bt_interval = st.number_input("Interval", min_value=1, value=1, step=1, key="bt_interval")
        bt_timespan = st.selectbox("Timespan", ["minute", "hour", "day", "week", "month", "quarter", "year"], index=0, key="bt_timespan")

        st.subheader("Strategy")
        bt_entry_time = st.text_input("Entry time", value="09:35", key="bt_entry_time")
        bt_use_limit_entry = st.checkbox("Use limit entry price", value=False, key="bt_use_limit")
        bt_entry_limit_price = st.number_input("Entry limit price", min_value=0.0, value=2.50, step=0.01, format="%.2f", disabled=not bt_use_limit_entry, key="bt_limit_price")
        bt_exit_time = st.text_input("Exit time", value="15:55", key="bt_exit_time")

        st.subheader("Risk Controls")
        if "bt_stop_loss_pct" not in st.session_state:
            st.session_state.bt_stop_loss_pct = 35.0
        if "bt_take_profit_pct" not in st.session_state:
            st.session_state.bt_take_profit_pct = 80.0

        bt_step_pct = st.selectbox("Button step", [1.0, 2.5, 5.0, 10.0], index=0, format_func=lambda x: f"{x:g}%", key="bt_step_pct")

        st.caption("Stop Loss")
        sl_minus, sl_value, sl_plus = st.columns([1, 2, 1])
        with sl_minus:
            if st.button("−", key="bt_sl_down", use_container_width=True):
                st.session_state.bt_stop_loss_pct = max(0.0, st.session_state.bt_stop_loss_pct - bt_step_pct)
        with sl_value:
            bt_stop_loss_pct = st.number_input("Stop loss %", min_value=0.0, max_value=100.0, step=bt_step_pct, key="bt_stop_loss_pct")
        with sl_plus:
            if st.button("+", key="bt_sl_up", use_container_width=True):
                st.session_state.bt_stop_loss_pct = min(100.0, st.session_state.bt_stop_loss_pct + bt_step_pct)

        st.caption("Take Profit")
        tp_minus, tp_value, tp_plus = st.columns([1, 2, 1])
        with tp_minus:
            if st.button("−", key="bt_tp_down", use_container_width=True):
                st.session_state.bt_take_profit_pct = max(0.0, st.session_state.bt_take_profit_pct - bt_step_pct)
        with tp_value:
            bt_take_profit_pct = st.number_input("Take profit %", min_value=0.0, step=bt_step_pct, key="bt_take_profit_pct")
        with tp_plus:
            if st.button("+", key="bt_tp_up", use_container_width=True):
                st.session_state.bt_take_profit_pct = st.session_state.bt_take_profit_pct + bt_step_pct

        bt_contracts = st.number_input("Contracts", min_value=1, value=1, step=1, key="bt_contracts")
        bt_starting_cash = st.number_input("Starting cash", min_value=0.0, value=10000.0, step=500.0, key="bt_cash")
        bt_commission = st.number_input("Commission / contract", min_value=0.0, value=0.65, step=0.05, key="bt_commission")

    option_ticker = build_option_ticker(bt_underlying, str(bt_expiry), bt_right, bt_strike)
    st.info(f"Option ticker: `{option_ticker}`")

    if st.button("Run Historical Backtest", type="primary", key="run_bt"):
        try:
            client = MassiveClient()
            with st.spinner("Fetching option candles from Massive..."):
                option_candles = client.get_aggregates(
                    ticker=option_ticker,
                    multiplier=int(bt_interval),
                    timespan=bt_timespan,
                    from_date=str(bt_from_date),
                    to_date=str(bt_to_date),
                )
            with st.spinner(f"Fetching {bt_underlying} candles from Massive..."):
                underlying_candles = client.get_aggregates(
                    ticker=bt_underlying,
                    multiplier=int(bt_interval),
                    timespan=bt_timespan,
                    from_date=str(bt_from_date),
                    to_date=str(bt_to_date),
                )

            output_dir = Path("output")
            output_dir.mkdir(exist_ok=True)
            option_candles.to_csv(output_dir / "option_candles.csv", index=False)
            underlying_candles.to_csv(output_dir / "underlying_candles.csv", index=False)

            config = BacktestConfig(
                contracts=int(bt_contracts),
                entry_time=bt_entry_time,
                exit_time=bt_exit_time,
                stop_loss=float(bt_stop_loss_pct) / 100,
                take_profit=float(bt_take_profit_pct) / 100,
                starting_cash=float(bt_starting_cash),
                commission_per_contract=float(bt_commission),
                entry_limit_price=float(bt_entry_limit_price) if bt_use_limit_entry else None,
            )
            trades, equity = run_long_option_backtest(option_candles, config, underlying_candles=underlying_candles)
            trades.to_csv(output_dir / "trades.csv", index=False)
            equity.to_csv(output_dir / "equity_curve.csv", index=False)
            plot_equity_curve(equity, output_dir / "equity_curve.png")

            st.success("Backtest complete. Files saved in the output folder.")

            if not trades.empty:
                total_pnl = float(trades["net_pnl"].sum())
                winning = trades[trades["net_pnl"] > 0]
                win_rate = len(winning) / len(trades) * 100 if len(trades) else 0
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total net P&L", f"${total_pnl:,.2f}")
                m2.metric("Trades", str(len(trades)))
                m3.metric("Win rate", f"{win_rate:.1f}%")
                m4.metric("Final cash", f"${float(trades['cash_after_trade'].iloc[-1]):,.2f}")

                st.markdown("#### Trades")
                st.dataframe(trades, use_container_width=True)

                csv = trades.to_csv(index=False).encode("utf-8")
                st.download_button("Download trades.csv", csv, "trades.csv", "text/csv")

            if not equity.empty:
                st.markdown("#### Equity Curve")
                st.line_chart(equity.set_index("timestamp")["equity"])

            with st.expander("Raw candle samples"):
                st.markdown("##### Option candles")
                st.dataframe(option_candles.head(50), use_container_width=True)
                st.markdown("##### Underlying candles")
                st.dataframe(underlying_candles.head(50), use_container_width=True)
        except Exception as e:
            st.error(f"Backtest failed: {e}")


with live_tab:
    st.subheader("IBKR Live P&L Tracker")
    st.caption("Native IBKR API. No ib_insync or asyncio. Shows SPY + option bid/ask/last/close and the exact price used for P&L.")

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        host = st.text_input("Host", value="127.0.0.1", key="live_host")
        port = st.selectbox("Port", options=[7497, 7496, 4002, 4001], index=0, key="live_port",
                            help="7497=TWS paper, 7496=TWS live, 4002=Gateway paper, 4001=Gateway live")
        use_delayed = st.checkbox("Use delayed data", value=False, key="live_delayed")
    with col_b:
        live_underlying = st.text_input("Underlying", value="SPY", key="live_underlying").upper().strip()
        live_expiry = st.date_input("Option expiry", value=date(2026, 5, 8), key="live_expiry")
        live_right = st.selectbox("Call / Put", options=["C", "P"], index=0, key="live_right")
    with col_c:
        live_strike = st.number_input("Strike", value=735.0, step=1.0, key="live_strike")
        wait_seconds = st.slider("Quote wait seconds", 2.0, 12.0, 4.0, 0.5, key="live_wait")
        client_id = st.number_input("Client ID", value=0, step=1, key="live_client_id", help="Use 0 for random client ID")

    st.divider()

    pcol1, pcol2, pcol3, pcol4 = st.columns(4)
    with pcol1:
        live_entry_price = st.number_input("Your option entry price", value=0.0, min_value=0.0, step=0.01, format="%.2f", key="live_entry_price")
    with pcol2:
        live_contracts = st.number_input("Contracts", value=1, min_value=1, step=1, key="live_contracts")
    with pcol3:
        live_stop_loss_pct = st.number_input("Stop loss %", value=35.0, min_value=0.0, step=1.0, key="live_stop")
    with pcol4:
        live_take_profit_pct = st.number_input("Take profit %", value=70.0, min_value=0.0, step=1.0, key="live_tp")

    refresh = st.button("Refresh IBKR Live Quote", type="primary", key="refresh_ibkr")

    if refresh:
        with st.spinner("Connecting to IBKR and requesting quotes..."):
            result = get_live_quotes(
                host=host,
                port=int(port),
                client_id=None if int(client_id) == 0 else int(client_id),
                underlying=live_underlying,
                expiry=str(live_expiry),
                right=live_right,
                strike=float(live_strike),
                wait_seconds=float(wait_seconds),
                use_delayed_data=use_delayed,
            )

        if not result.get("ok"):
            st.error(result.get("error", "IBKR quote request failed."))
            for msg in result.get("messages", []):
                st.warning(msg)
        else:
            uq = result["underlying_quote"]
            oq = result["option_quote"]
            option_price = oq.get("market_price")
            spy_price = uq.get("market_price")

            st.caption(
                f"Data requested: {result.get('market_data_type_requested', 'unknown')} | "
                f"Snapshot UTC: {result.get('snapshot_time_utc', 'N/A')} | "
                f"Client ID: {result.get('client_id', 'N/A')}"
            )

            st.markdown("#### Price used for P&L")
            metric_cols = st.columns(4)
            metric_cols[0].metric("SPY used", fmt_money(spy_price), uq.get("market_price_source") or "")
            metric_cols[1].metric("Option used", fmt_money(option_price), oq.get("market_price_source") or "")
            metric_cols[2].metric("Option midpoint", fmt_money(oq.get("midpoint")))
            metric_cols[3].metric("Option last", fmt_money(oq.get("last")))

            st.markdown("#### SPY quote breakdown")
            spy_cols = st.columns(5)
            spy_cols[0].metric("SPY bid", fmt_money(uq.get("bid")))
            spy_cols[1].metric("SPY ask", fmt_money(uq.get("ask")))
            spy_cols[2].metric("SPY midpoint", fmt_money(uq.get("midpoint")))
            spy_cols[3].metric("SPY last", fmt_money(uq.get("last")))
            spy_cols[4].metric("SPY close", fmt_money(uq.get("close")))

            st.markdown("#### Option quote breakdown")
            opt_cols = st.columns(5)
            opt_cols[0].metric("Option bid", fmt_money(oq.get("bid")))
            opt_cols[1].metric("Option ask", fmt_money(oq.get("ask")))
            opt_cols[2].metric("Option midpoint", fmt_money(oq.get("midpoint")))
            opt_cols[3].metric("Option last", fmt_money(oq.get("last")))
            opt_cols[4].metric("Option close", fmt_money(oq.get("close")))

            if option_price is not None and option_price == oq.get("close") and oq.get("bid") is None and oq.get("ask") is None:
                st.warning("Live option bid/ask was not returned. P&L is using the option close price, so it may not reflect the current live value. Try delayed data, increase quote wait seconds, or check OPRA/options data permissions.")

            if option_price is not None and live_entry_price > 0:
                pnl = (float(option_price) - float(live_entry_price)) * 100 * int(live_contracts)
                pnl_pct = ((float(option_price) - float(live_entry_price)) / float(live_entry_price)) * 100
                stop_price = float(live_entry_price) * (1 - float(live_stop_loss_pct) / 100)
                take_profit_price = float(live_entry_price) * (1 + float(live_take_profit_pct) / 100)

                st.divider()
                q1, q2, q3, q4 = st.columns(4)
                q1.metric("Unrealized P&L", f"${pnl:,.2f}", f"{pnl_pct:.2f}%")
                q2.metric("Stop price", f"${stop_price:.2f}")
                q3.metric("Take-profit price", f"${take_profit_price:.2f}")
                q4.metric("Contracts", str(int(live_contracts)))

                if option_price <= stop_price:
                    st.error("Stop loss level is hit or below current option price.")
                elif option_price >= take_profit_price:
                    st.success("Take-profit level is hit or above current option price.")
                else:
                    st.info("No stop-loss or take-profit trigger right now.")
            elif option_price is None:
                st.warning("IBKR connected, but option price was not returned. Try increasing quote wait seconds, enabling delayed data, or checking your market data subscription.")
            else:
                st.info("Enter your option entry price to calculate live P&L.")

            quote_df = pd.DataFrame([
                {"instrument": "Underlying", **uq},
                {"instrument": "Option", **oq},
            ])
            display_cols = [
                "instrument", "symbol", "bid", "ask", "midpoint", "last", "close",
                "market_price", "market_price_source", "snapshot_time_utc",
                "bid_time_utc", "ask_time_utc", "last_time_utc", "close_time_utc",
            ]
            st.markdown("#### Full quote table with timestamps")
            st.dataframe(quote_df[[c for c in display_cols if c in quote_df.columns]], use_container_width=True)

            messages = result.get("messages", [])
            if messages:
                with st.expander("IBKR messages"):
                    for msg in messages:
                        st.write(msg)


with setup_tab:
    st.markdown(
        """
## Setup

### 1. Install packages

Open PowerShell inside this folder and run:

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Or double-click:

```text
run_ui.bat
```

### 2. Massive setup for Historical Backtest

Create a file named `.env` in this project folder:

```text
MASSIVE_API_KEY=your_api_key_here
```

The backtest tab uses Massive historical aggregate candles for both the option contract and the underlying SPY candles.

### 3. IBKR setup for Live P&L

Open TWS or IB Gateway first, then enable API access:

```text
File → Global Configuration → API → Settings → Enable ActiveX and Socket Clients
```

Common ports:

| Platform | Paper | Live |
|---|---:|---:|
| TWS | 7497 | 7496 |
| IB Gateway | 4002 | 4001 |

If you do not have live OPRA/options data permissions, check **Use delayed data** in the live tab.

### Notes

- This app is read-only. It does not place, modify, or cancel trades.
- For live P&L, the app uses option midpoint first, then last, then close as fallback.
- For historical backtests, output files are saved into the `output/` folder.
"""
    )
