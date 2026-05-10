# SPY Options Backtester UI

This is the UI version of the SPY options backtester.

## 1. Install packages

Open PowerShell inside this folder and run:

```powershell
pip install -r requirements.txt
```

## 2. Add your Massive API key

Create a file named `.env` in this folder:

```text
MASSIVE_API_KEY=your_api_key_here
```

## 3. Start the UI

PowerShell:

```powershell
python -m streamlit run app.py
```

Or double-click:

```text
run_ui.bat
```

## 4. Example settings

For SPY $735 Call expiring May 8, 2026:

- Underlying: SPY
- Expiry: 2026-05-08
- Option Type: Call
- Strike: 735
- From date: 2026-05-01
- To date: 2026-05-08
- Entry time: 09:35
- Exit time: 15:55
- Stop loss: 35%
- Take profit: 80%

The UI creates these output files after each run:

- `output/candles.csv`
- `output/trades.csv`
- `output/equity_curve.csv`
- `output/equity_curve.png`

## Limit Entry / Fill Time

The UI now has a **Use limit entry price** checkbox.

- Off: enters at the first option candle close at or after Entry time.
- On: waits until the option candle low is less than or equal to your Entry limit price.
- `fill_time` shows the first candle timestamp that touched your limit.
- `minutes_to_fill` shows how long it took from your requested entry time.

For example, if Entry time is `09:35` and Entry limit price is `2.50`, the backtest scans candles from 09:35 onward and fills the trade at the first candle where `low <= 2.50`.


## Stop Loss / Take Profit Buttons

The UI now has plus and minus buttons for Stop Loss % and Take Profit %.

- Choose the button step: 1%, 2.5%, 5%, or 10%.
- Press `+` to increase.
- Press `−` to decrease.
- You can still type the exact percentage manually.
