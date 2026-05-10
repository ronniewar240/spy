# SPY Options Backtester + IBKR Live P&L

One Streamlit project with both tools merged:

1. **Historical Backtest** using Massive historical aggregate candles
2. **IBKR Live P&L Tracker** using IBKR's native `ibapi` package

This app is read-only. It does **not** place trades.

## Install

Open PowerShell in this folder:

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Or double-click:

```text
run_ui.bat
```

## Massive setup

Create a `.env` file in this folder:

```text
MASSIVE_API_KEY=your_api_key_here
```

The Historical Backtest tab uses this key.

## IBKR setup

Open TWS or IB Gateway and enable:

```text
File → Global Configuration → API → Settings → Enable ActiveX and Socket Clients
```

Ports:

| Platform | Paper | Live |
|---|---:|---:|
| TWS | 7497 | 7496 |
| IB Gateway | 4002 | 4001 |

If options bid/ask do not show, try **Use delayed data** and increase **Quote wait seconds**.

## Output files

Backtest results are saved in `output/`:

- `option_candles.csv`
- `underlying_candles.csv`
- `trades.csv`
- `equity_curve.csv`
- `equity_curve.png`
