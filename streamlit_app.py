import calendar
import hashlib
import importlib.util
import json
import os
import sqlite3
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from dropbox_import import download_new_csvs

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "trades_streamlit.db"
UPLOAD_DIR = APP_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

# Reuse your Flask app's battle-tested broker parsers without running Flask.
spec = importlib.util.spec_from_file_location("legacy_flask_app", APP_DIR / "legacy_flask_app.py")
legacy = importlib.util.module_from_spec(spec)
spec.loader.exec_module(legacy)

st.set_page_config(page_title="Trade Journal", page_icon="📈", layout="wide")

CUSTOM_CSS = """
<style>
.block-container { padding-top: 1.2rem; padding-bottom: 3rem; }
.metric-card {
    background: linear-gradient(180deg, rgba(18,26,47,.95), rgba(15,21,40,.95));
    border: 1px solid rgba(148,163,184,.16);
    border-radius: 18px;
    padding: 16px;
    min-height: 96px;
}
.metric-label { color: #94a3b8; font-size: 13px; }
.metric-value { font-size: 28px; font-weight: 900; margin-top: 4px; }
.pos { color: #22c55e; }
.neg { color: #ef4444; }
.muted { color: #94a3b8; font-size: 13px; }
.calendar-cell {
    border: 1px solid rgba(148,163,184,.16);
    border-radius: 12px;
    padding: 8px;
    min-height: 92px;
    overflow: hidden;
}
.calendar-day { font-weight: 900; }
.calendar-pnl { font-weight: 900; font-size: 14px; }
.small-platform { font-size: 10px; line-height: 1.15; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def connect():
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout = 30000")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = NORMAL")
    return conn


def init_db():
    conn = connect()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL UNIQUE,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS portfolios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(user_id, name)
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS trades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            portfolio_id INTEGER NOT NULL,
            broker TEXT NOT NULL,
            asset_category TEXT,
            currency TEXT,
            symbol TEXT NOT NULL,
            trade_datetime TEXT,
            quantity REAL,
            side TEXT,
            trade_price REAL,
            close_price REAL,
            buy_price REAL,
            sell_price REAL,
            proceeds REAL,
            commission REAL,
            basis REAL,
            realized_pl REAL,
            mtm_pl REAL,
            risk_amount REAL,
            r_multiple REAL,
            code TEXT,
            import_file TEXT,
            batch_id TEXT,
            notes TEXT,
            trade_key TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS trade_journal (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trade_id INTEGER NOT NULL UNIQUE,
            setup_tag TEXT,
            mistake_tag TEXT,
            note TEXT,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS imported_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            source TEXT NOT NULL,
            file_name TEXT,
            file_hash TEXT NOT NULL,
            batch_id TEXT,
            parsed_count INTEGER DEFAULT 0,
            inserted_count INTEGER DEFAULT 0,
            skipped_count INTEGER DEFAULT 0,
            status TEXT,
            imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(user_id, source, file_hash)
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_trade_scope ON trades(user_id, portfolio_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_trade_key ON trades(user_id, portfolio_id, trade_key)")
    conn.commit()
    conn.close()


def get_or_create_user(email="streamlit@local"):
    conn = connect()
    row = conn.execute("SELECT id FROM users WHERE email = ?", (email,)).fetchone()
    if row:
        uid = row["id"]
    else:
        cur = conn.execute("INSERT INTO users(email) VALUES (?)", (email,))
        uid = cur.lastrowid
        conn.commit()
    conn.close()
    return uid


def get_portfolios(user_id):
    conn = connect()
    rows = conn.execute("SELECT id, name FROM portfolios WHERE user_id = ? ORDER BY name", (user_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_or_create_portfolio(user_id, name="Main Portfolio"):
    conn = connect()
    row = conn.execute("SELECT id FROM portfolios WHERE user_id = ? AND name = ?", (user_id, name)).fetchone()
    if row:
        pid = row["id"]
    else:
        cur = conn.execute("INSERT OR IGNORE INTO portfolios(user_id, name) VALUES (?, ?)", (user_id, name))
        conn.commit()
        row = conn.execute("SELECT id FROM portfolios WHERE user_id = ? AND name = ?", (user_id, name)).fetchone()
        pid = row["id"]
    conn.close()
    return pid


def file_hash(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def to_float(v, default=0.0):
    try:
        if v is None or str(v).strip() == "":
            return default
        return float(v)
    except Exception:
        return default


def normalize_key_value(value, places=8):
    if value is None:
        return ""
    text = str(value).strip()
    if text == "":
        return ""
    try:
        return f"{float(text):.{places}f}"
    except Exception:
        return text.upper()


def make_trade_key(user_id, portfolio_id, trade):
    if hasattr(legacy, "make_trade_key"):
        try:
            return legacy.make_trade_key(user_id, portfolio_id, trade)
        except Exception:
            pass
    parts = [
        str(user_id), str(portfolio_id),
        str(trade.get("broker") or "").upper().strip(),
        str(trade.get("symbol") or "").upper().strip(),
        str(trade.get("trade_datetime") or "").strip(),
        normalize_key_value(trade.get("quantity")),
        str(trade.get("side") or "").upper().strip(),
        normalize_key_value(trade.get("trade_price")),
        normalize_key_value(trade.get("proceeds")),
        normalize_key_value(trade.get("commission")),
        str(trade.get("code") or "").upper().strip(),
    ]
    return hashlib.sha256("|".join(parts).encode("utf-8")).hexdigest()


def derive_buy_sell_prices(trade):
    if hasattr(legacy, "derive_buy_sell_prices"):
        try:
            return legacy.derive_buy_sell_prices(trade)
        except Exception:
            pass
    buy = to_float(trade.get("buy_price"))
    sell = to_float(trade.get("sell_price"))
    if buy or sell:
        return buy, sell
    side = str(trade.get("side") or "").upper()
    qty = to_float(trade.get("quantity"))
    trade_price = to_float(trade.get("trade_price"))
    close_price = to_float(trade.get("close_price"))
    if side == "BUY" or qty > 0:
        return trade_price, close_price
    if side == "SELL" or qty < 0:
        return close_price, trade_price
    return 0.0, 0.0


def parse_uploaded_file(path):
    import_type = legacy.detect_import_type(str(path))
    if import_type in {"ibkr", "ibkr_trades"}:
        return import_type, legacy.parse_ibkr_activity_csv(str(path))
    if import_type == "ibkr_summary":
        return import_type, legacy.parse_ibkr_summary_csv(str(path))
    if import_type == "wealthsimple":
        return import_type, legacy.parse_wealthsimple_csv(str(path))
    if import_type == "performance":
        return import_type, legacy.parse_performance_csv(str(path))
    raise ValueError(f"Unsupported or unrecognized CSV format: {import_type}")


def insert_trades(user_id, portfolio_id, trades, import_file=None, batch_id=None):
    conn = connect()
    cur = conn.cursor()
    inserted = 0
    skipped = 0
    seen = set()
    for t in trades:
        t = dict(t)
        t.setdefault("broker", "Unknown")
        t.setdefault("asset_category", "")
        t.setdefault("currency", "")
        t.setdefault("symbol", "Unknown")
        t.setdefault("trade_datetime", None)
        t.setdefault("quantity", 0)
        t.setdefault("side", "")
        t.setdefault("trade_price", 0)
        t.setdefault("close_price", 0)
        t.setdefault("proceeds", 0)
        t.setdefault("commission", 0)
        t.setdefault("basis", 0)
        t.setdefault("realized_pl", 0)
        t.setdefault("mtm_pl", 0)
        t.setdefault("risk_amount", 0)
        t.setdefault("r_multiple", 0)
        t.setdefault("code", "")
        t.setdefault("notes", "")
        buy_price, sell_price = derive_buy_sell_prices(t)
        t["buy_price"] = buy_price
        t["sell_price"] = sell_price
        key = make_trade_key(user_id, portfolio_id, t)
        if key in seen:
            skipped += 1
            continue
        seen.add(key)
        existing = cur.execute(
            "SELECT id FROM trades WHERE user_id=? AND portfolio_id=? AND trade_key=? LIMIT 1",
            (user_id, portfolio_id, key),
        ).fetchone()
        if existing:
            skipped += 1
            continue
        cur.execute("""
            INSERT INTO trades (
                user_id, portfolio_id, broker, asset_category, currency, symbol,
                trade_datetime, quantity, side, trade_price, close_price, buy_price, sell_price,
                proceeds, commission, basis, realized_pl, mtm_pl, risk_amount, r_multiple,
                code, import_file, batch_id, notes, trade_key
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            user_id, portfolio_id, t["broker"], t["asset_category"], t["currency"], t["symbol"],
            t["trade_datetime"], to_float(t["quantity"]), t["side"], to_float(t["trade_price"]),
            to_float(t["close_price"]), to_float(t["buy_price"]), to_float(t["sell_price"]),
            to_float(t["proceeds"]), to_float(t["commission"]), to_float(t["basis"]),
            to_float(t["realized_pl"]), to_float(t["mtm_pl"]), to_float(t["risk_amount"]),
            to_float(t["r_multiple"]), t["code"], import_file, batch_id, t["notes"], key,
        ))
        inserted += 1
    conn.commit()
    conn.close()
    return inserted, skipped


def load_trades_df(user_id, portfolio_id):
    conn = connect()
    df = pd.read_sql_query("""
        SELECT t.*, j.setup_tag, j.mistake_tag, j.note AS journal_note
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        WHERE t.user_id = ? AND t.portfolio_id = ?
        ORDER BY COALESCE(t.trade_datetime, '') DESC, t.id DESC
    """, conn, params=(user_id, portfolio_id))
    conn.close()
    if not df.empty:
        df["date"] = pd.to_datetime(df["trade_datetime"], errors="coerce").dt.date
        df["month"] = pd.to_datetime(df["trade_datetime"], errors="coerce").dt.to_period("M").astype(str)
    return df


def broker_platform_name(broker):
    b = str(broker or "").lower()
    if "wealthsimple" in b:
        return "Wealthsimple"
    if "ninja" in b or "performance" in b:
        return "NinjaTrader"
    if "ibkr" in b or "interactive" in b:
        return "IBKR"
    return "Other"


def render_metric(label, value, klass=""):
    st.markdown(
        f'<div class="metric-card"><div class="metric-label">{label}</div><div class="metric-value {klass}">{value}</div></div>',
        unsafe_allow_html=True,
    )


def money(x):
    x = to_float(x)
    sign = "+" if x > 0 else "" 
    return f"{sign}{x:,.2f}"


def generate_tomorrow_plan(df):
    if df.empty:
        return {"mode": "Collect data", "best_setup": "No data", "avoid_setup": "No data", "best_time": "No data", "confidence": 0, "reason": "Import trades first."}
    work = df.copy()
    work["setup"] = work.get("setup_tag", pd.Series([None] * len(work))).fillna("Unlabeled")
    work["pnl"] = pd.to_numeric(work["realized_pl"], errors="coerce").fillna(0)
    setup = work.groupby("setup")["pnl"].agg(["sum", "count"]).sort_values("sum", ascending=False)
    best = setup.index[0] if len(setup) else "No data"
    worst = setup.index[-1] if len(setup) else "No data"
    work["dt"] = pd.to_datetime(work["trade_datetime"], errors="coerce")
    work["bucket"] = work["dt"].dt.hour.map(lambda h: "Morning" if pd.notna(h) and h < 11 else "Midday" if pd.notna(h) and h < 14 else "Afternoon" if pd.notna(h) else "Unknown")
    time_edge = work.groupby("bucket")["pnl"].sum().sort_values(ascending=False)
    best_time = time_edge.index[0] if len(time_edge) else "No data"
    recent = work.head(10)["pnl"].sum()
    confidence = min(95, int(abs(recent) / 50) + min(40, int(setup.iloc[0]["count"] * 5)) if len(setup) else 0)
    return {"mode": "Selective", "best_setup": best, "avoid_setup": worst, "best_time": best_time, "confidence": confidence, "reason": f"{best} has the strongest recent/overall P&L in your journal."}



def get_secret_value(*names, default=""):
    """Read a value from Streamlit secrets using a few possible key names."""
    for name in names:
        try:
            if "." in name:
                head, tail = name.split(".", 1)
                value = st.secrets.get(head, {}).get(tail, "")
            else:
                value = st.secrets.get(name, "")
            if value:
                return str(value)
        except Exception:
            continue
    return default


def has_imported_source_hash(user_id, source, file_hash):
    conn = connect()
    row = conn.execute(
        "SELECT id FROM imported_files WHERE user_id=? AND source=? AND file_hash=? LIMIT 1",
        (user_id, source, file_hash),
    ).fetchone()
    conn.close()
    return row is not None


def record_imported_source_file(user_id, source, file_name, file_hash, batch_id, parsed_count, inserted_count, skipped_count, status):
    conn = connect()
    conn.execute(
        """
        INSERT OR IGNORE INTO imported_files(
            user_id, source, file_name, file_hash, batch_id,
            parsed_count, inserted_count, skipped_count, status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (user_id, source, file_name, file_hash, batch_id, parsed_count, inserted_count, skipped_count, status),
    )
    conn.commit()
    conn.close()


def import_one_csv_path(user_id, portfolio_id, path, source, file_hash_value=None):
    path = Path(path)
    digest = file_hash_value or file_hash(path)
    if has_imported_source_hash(user_id, source, digest):
        return {
            "file": path.name,
            "status": "skipped_file_hash",
            "parsed": 0,
            "inserted": 0,
            "skipped": 0,
            "message": "Skipped already imported file.",
        }

    import_type, trades = parse_uploaded_file(path)
    batch_id = f"{source.upper()}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}"
    inserted, skipped = insert_trades(user_id, portfolio_id, trades, path.name, batch_id)
    record_imported_source_file(
        user_id,
        source,
        path.name,
        digest,
        batch_id,
        len(trades),
        inserted,
        skipped,
        "OK",
    )
    return {
        "file": path.name,
        "status": "imported",
        "type": import_type,
        "batch_id": batch_id,
        "parsed": len(trades),
        "inserted": inserted,
        "skipped": skipped,
        "message": f"{import_type}: parsed {len(trades)}, inserted {inserted}, skipped {skipped}.",
    }


def render_dropbox_import_page(user_id, portfolio_id):
    st.title("☁️ Dropbox folder auto-import")
    st.caption("Cloud-friendly replacement for local folder watching. Put broker CSV exports in a Dropbox folder, then scan/import them from Streamlit.")

    token = get_secret_value("dropbox.access_token", "DROPBOX_ACCESS_TOKEN")
    default_folder = get_secret_value("dropbox.folder", "DROPBOX_FOLDER", default="/TradeJournalExports")

    with st.expander("How to configure Streamlit secrets", expanded=not bool(token)):
        st.markdown(
            """
Add this in **Streamlit Cloud → App settings → Secrets**:

```toml
[dropbox]
access_token = "YOUR_DROPBOX_ACCESS_TOKEN"
folder = "/TradeJournalExports"
```

Then export/download your IBKR, Wealthsimple, or NinjaTrader CSVs into that Dropbox folder.
            """
        )

    folder = st.text_input("Dropbox folder path", value=default_folder or "/TradeJournalExports")
    st.write("Status:", "✅ Dropbox token found" if token else "❌ Missing Dropbox token in secrets")

    if st.button("Scan Dropbox Now", type="primary", disabled=not bool(token)):
        with st.spinner("Scanning Dropbox and importing new CSV files..."):
            try:
                download_dir = UPLOAD_DIR / "dropbox"
                downloaded = download_new_csvs(token, folder, download_dir)
                if not downloaded:
                    st.info("No CSV files found in that Dropbox folder.")
                    return

                results = []
                for item in downloaded:
                    if has_imported_source_hash(user_id, "dropbox", item.content_hash):
                        results.append({
                            "file": item.name,
                            "status": "skipped_file_hash",
                            "parsed": 0,
                            "inserted": 0,
                            "skipped": 0,
                            "message": "Skipped already imported Dropbox file.",
                        })
                        continue
                    try:
                        results.append(import_one_csv_path(user_id, portfolio_id, item.local_path, "dropbox", item.content_hash))
                    except Exception as e:
                        results.append({
                            "file": item.name,
                            "status": "error",
                            "parsed": 0,
                            "inserted": 0,
                            "skipped": 0,
                            "message": str(e),
                        })

                st.subheader("Scan results")
                st.dataframe(pd.DataFrame(results), use_container_width=True, hide_index=True)
                st.success(f"Checked {len(downloaded)} file(s). Inserted {sum(r.get('inserted', 0) for r in results)} trade(s).")
            except Exception as e:
                st.error(f"Dropbox scan failed: {e}")

    st.info("Tip: this simple version scans when you click the button. On Streamlit Cloud, true background jobs are not reliable; for scheduled scans, deploy the same logic as a cron job on Render/Railway/GitHub Actions, or add Dropbox webhooks later.")

def render_monthly_calendar(df):
    st.subheader("📅 Monthly P&L Calendar")
    if df.empty or "date" not in df:
        st.info("No dated trades yet.")
        return
    dated = df.dropna(subset=["date"]).copy()
    if dated.empty:
        st.info("No trades with valid trade dates yet.")
        return
    dated["date"] = pd.to_datetime(dated["date"])
    dated["platform"] = dated["broker"].map(broker_platform_name)
    months = sorted(dated["date"].dt.to_period("M").astype(str).unique(), reverse=True)
    selected = st.selectbox("Calendar month", months, index=0)
    month_df = dated[dated["date"].dt.to_period("M").astype(str) == selected]
    total = month_df["realized_pl"].sum()
    wins = month_df.groupby(month_df["date"].dt.date)["realized_pl"].sum()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Monthly Total P&L", money(total))
    c2.metric("Winning-Day P&L", money(wins[wins > 0].sum()))
    c3.metric("Losing-Day P&L", money(wins[wins < 0].sum()))
    c4.metric("Trades This Month", int(len(month_df)))
    p1, p2, p3 = st.columns(3)
    platform = month_df.groupby("platform").agg(pnl=("realized_pl", "sum"), trades=("id", "count"))
    for col, name in zip([p1, p2, p3], ["IBKR", "NinjaTrader", "Wealthsimple"]):
        pnl = platform.loc[name, "pnl"] if name in platform.index else 0
        trades = platform.loc[name, "trades"] if name in platform.index else 0
        col.metric(f"{name} Monthly P&L", money(pnl), f"{int(trades)} trades")

    year, mon = map(int, selected.split("-"))
    cal = calendar.Calendar(firstweekday=0)
    daily = month_df.groupby(month_df["date"].dt.date).agg(pnl=("realized_pl", "sum"), trades=("id", "count"))
    daily_platform = month_df.groupby([month_df["date"].dt.date, "platform"])["realized_pl"].sum()
    st.caption("Calendar uses trade date, not settlement date.")
    cols = st.columns(7)
    for col, name in zip(cols, ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]):
        col.markdown(f"**{name}**")
    for week in cal.monthdatescalendar(year, mon):
        cols = st.columns(7)
        for col, day in zip(cols, week):
            if day.month != mon:
                col.markdown('<div class="calendar-cell" style="opacity:.25"></div>', unsafe_allow_html=True)
                continue
            pnl = daily.loc[day, "pnl"] if day in daily.index else 0
            trades = daily.loc[day, "trades"] if day in daily.index else 0
            bg = "rgba(34,197,94,.22)" if pnl > 0 else "rgba(239,68,68,.22)" if pnl < 0 else "rgba(148,163,184,.08)"
            platforms = []
            for name, short in [("IBKR", "IBKR"), ("NinjaTrader", "NT"), ("Wealthsimple", "WS")]:
                try:
                    val = daily_platform.loc[(day, name)]
                    if val:
                        platforms.append(f"<div class='small-platform'>{short}: {money(val)}</div>")
                except KeyError:
                    pass
            col.markdown(
                f"""
                <div class="calendar-cell" style="background:{bg}">
                  <div class="calendar-day">{day.day}</div>
                  <div class="muted">{int(trades)} trades</div>
                  <div class="calendar-pnl {'pos' if pnl > 0 else 'neg' if pnl < 0 else ''}">{money(pnl) if trades else '—'}</div>
                  {''.join(platforms)}
                </div>
                """,
                unsafe_allow_html=True,
            )


def main():
    init_db()
    st.sidebar.title("Trade Journal")
    user_email = st.sidebar.text_input("User email", value=st.session_state.get("email", "streamlit@local"))
    st.session_state["email"] = user_email
    user_id = get_or_create_user(user_email)
    portfolios = get_portfolios(user_id)
    if not portfolios:
        get_or_create_portfolio(user_id)
        portfolios = get_portfolios(user_id)
    names = [p["name"] for p in portfolios]
    selected_name = st.sidebar.selectbox("Current portfolio", names)
    portfolio_id = next(p["id"] for p in portfolios if p["name"] == selected_name)
    new_portfolio = st.sidebar.text_input("Create portfolio")
    if st.sidebar.button("Add portfolio") and new_portfolio.strip():
        get_or_create_portfolio(user_id, new_portfolio.strip())
        st.rerun()
    st.sidebar.caption(f"Portfolio: {selected_name}")
    page = st.sidebar.radio("Page", ["Dashboard", "Import", "Dropbox auto-import", "Trades", "Monthly", "Export", "Folder import notes"])
    df = load_trades_df(user_id, portfolio_id)

    if page == "Dashboard":
        st.title("📈 Trade Journal Dashboard")
        st.caption(f"Current portfolio: {selected_name}")
        pnl = df["realized_pl"].sum() if not df.empty else 0
        wins = (df["realized_pl"] > 0).sum() if not df.empty else 0
        trade_count = len(df)
        win_rate = (wins / trade_count * 100) if trade_count else 0
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Trades", trade_count)
        c2.metric("Realized P&L", money(pnl))
        c3.metric("Win Rate", f"{win_rate:.1f}%")
        c4.metric("Platforms", df["broker"].nunique() if not df.empty else 0)
        render_monthly_calendar(df)
        st.subheader("📅 Tomorrow Trading Plan")
        plan = generate_tomorrow_plan(df)
        p1, p2, p3, p4 = st.columns(4)
        p1.metric("Mode", plan["mode"])
        p2.metric("Best Setup", plan["best_setup"])
        p3.metric("Avoid", plan["avoid_setup"])
        p4.metric("Confidence", f"{plan['confidence']}%")
        st.info(plan["reason"])
        if not df.empty:
            chart_df = df.dropna(subset=["date"]).sort_values("date").copy()
            chart_df["equity"] = chart_df["realized_pl"].cumsum()
            st.plotly_chart(px.line(chart_df, x="date", y="equity", title="Equity Curve"), use_container_width=True)

    elif page == "Import":
        st.title("📤 Import trades")
        files = st.file_uploader("Upload CSV files", type=["csv"], accept_multiple_files=True)
        if st.button("Import uploaded files", type="primary"):
            if not files:
                st.warning("Upload at least one CSV.")
            else:
                for f in files:
                    path = UPLOAD_DIR / f.name
                    path.write_bytes(f.getbuffer())
                    digest = file_hash(path)
                    conn = connect()
                    exists = conn.execute("SELECT id FROM imported_files WHERE user_id=? AND source=? AND file_hash=?", (user_id, "streamlit_upload", digest)).fetchone()
                    conn.close()
                    if exists:
                        st.info(f"Skipped already imported file: {f.name}")
                        continue
                    try:
                        import_type, trades = parse_uploaded_file(path)
                        batch_id = f"ST_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}"
                        inserted, skipped = insert_trades(user_id, portfolio_id, trades, f.name, batch_id)
                        conn = connect()
                        conn.execute("INSERT OR IGNORE INTO imported_files(user_id, source, file_name, file_hash, batch_id, parsed_count, inserted_count, skipped_count, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", (user_id, "streamlit_upload", f.name, digest, batch_id, len(trades), inserted, skipped, "OK"))
                        conn.commit(); conn.close()
                        st.success(f"{f.name}: parsed {len(trades)}, inserted {inserted}, skipped {skipped} ({import_type})")
                    except Exception as e:
                        st.error(f"{f.name}: {e}")

    elif page == "Dropbox auto-import":
        render_dropbox_import_page(user_id, portfolio_id)

    elif page == "Trades":
        st.title("📋 Trades")
        if df.empty:
            st.info("No trades yet.")
        else:
            cols = ["id", "broker", "symbol", "trade_datetime", "quantity", "side", "buy_price", "sell_price", "realized_pl", "setup_tag", "journal_note"]
            st.dataframe(df[[c for c in cols if c in df.columns]], use_container_width=True, hide_index=True)

    elif page == "Monthly":
        st.title("📊 Monthly analysis")
        if df.empty:
            st.info("No trades yet.")
        else:
            monthly = df.dropna(subset=["month"]).groupby("month").agg(realized_pl=("realized_pl", "sum"), fees=("commission", "sum"), trades=("id", "count")).reset_index()
            st.dataframe(monthly, use_container_width=True, hide_index=True)
            st.plotly_chart(px.bar(monthly, x="month", y="realized_pl", title="Monthly P&L"), use_container_width=True)

    elif page == "Export":
        st.title("⬇️ Export")
        if df.empty:
            st.info("Nothing to export.")
        else:
            csv = df.to_csv(index=False).encode("utf-8")
            st.download_button("Download trades CSV", csv, "trades_export.csv", "text/csv")

    else:
        st.title("📁 Folder import notes")
        st.info("Streamlit Cloud cannot watch folders on your local computer. For deployed use, upload CSVs on the Import page. Folder auto-import only works when running Streamlit locally on the same machine as the export folders.")
        folder = st.text_input("Local folder to scan (local-only)")
        if st.button("Scan local folder") and folder:
            folder_path = Path(folder).expanduser()
            if not folder_path.exists():
                st.error("Folder not found.")
            else:
                files = list(folder_path.glob("*.csv"))
                st.write(f"Found {len(files)} CSV file(s). Use the Import page for cloud deployment.")


if __name__ == "__main__":
    main()
