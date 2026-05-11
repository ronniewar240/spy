# -*- coding: utf-8 -*-
import os
import re
import csv
import io
import json
import sqlite3
import hashlib
import secrets
import time
import threading
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from collections import defaultdict, deque
from datetime import datetime
from decimal import Decimal, InvalidOperation

from flask import (
    Flask, request, jsonify, Response, redirect, url_for,
    send_from_directory, session, render_template
)
from openpyxl import Workbook
from openpyxl.styles import Font

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "change-this-secret-in-production")

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
SCREENSHOT_FOLDER = os.path.join(BASE_DIR, "screenshots")
DB_PATH = os.path.join(BASE_DIR, "trades_v2.db")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(SCREENSHOT_FOLDER, exist_ok=True)

OPTION_RE = re.compile(
    r"^(?P<underlying>[A-Z\.]+)\s+(?P<exp>\d{2}[A-Z]{3}\d{2})\s+(?P<strike>\d+(\.\d+)?)\s+(?P<cp>[CP])$"
)


# =========================================================
# DB
# =========================================================
def get_db_connection():
    # SQLite can report "database is locked" during CSV imports if another request
    # is reading/writing at the same time. A timeout + WAL mode makes local Flask
    # usage much more tolerant, especially on Windows.
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout = 30000")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = NORMAL")
    return conn


def table_columns(conn, table_name):
    return [r["name"] for r in conn.execute(f"PRAGMA table_info({table_name})").fetchall()]


def column_exists(conn, table_name, column_name):
    return column_name in table_columns(conn, table_name)


def ensure_column(conn, table_name, column_name, column_type):
    if not column_exists(conn, table_name, column_name):
        conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}")


def init_db():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS portfolios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(user_id, name),
            FOREIGN KEY(user_id) REFERENCES users(id)
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
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(user_id) REFERENCES users(id),
            FOREIGN KEY(portfolio_id) REFERENCES portfolios(id)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS trade_journal (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trade_id INTEGER NOT NULL UNIQUE,
            setup_tag TEXT,
            mistake_tag TEXT,
            note TEXT,
            screenshot_filename TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(trade_id) REFERENCES trades(id)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS ibkr_settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL UNIQUE,
            flex_token TEXT,
            query_id TEXT,
            account_id TEXT,
            report_format TEXT DEFAULT 'xml',
            auto_import_enabled INTEGER DEFAULT 0,
            auto_import_hour INTEGER DEFAULT 6,
            last_import_at TEXT,
            last_auto_import_date TEXT,
            last_auto_import_status TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(user_id) REFERENCES users(id)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS ninjatrader_settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL UNIQUE,
            folder_path TEXT,
            auto_import_enabled INTEGER DEFAULT 0,
            scan_interval_minutes INTEGER DEFAULT 10,
            last_import_at TEXT,
            last_scan_at TEXT,
            last_scan_status TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(user_id) REFERENCES users(id)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS wealthsimple_settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL UNIQUE,
            folder_path TEXT,
            auto_import_enabled INTEGER DEFAULT 0,
            scan_interval_minutes INTEGER DEFAULT 10,
            last_import_at TEXT,
            last_scan_at TEXT,
            last_scan_status TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(user_id) REFERENCES users(id)
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS imported_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            source TEXT NOT NULL,
            file_path TEXT,
            file_name TEXT,
            file_hash TEXT NOT NULL,
            file_size INTEGER,
            file_mtime REAL,
            batch_id TEXT,
            parsed_count INTEGER DEFAULT 0,
            inserted_count INTEGER DEFAULT 0,
            skipped_count INTEGER DEFAULT 0,
            status TEXT,
            imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(user_id, source, file_hash),
            FOREIGN KEY(user_id) REFERENCES users(id)
        )
    """)

    ensure_column(conn, "trades", "user_id", "INTEGER")
    ensure_column(conn, "trades", "portfolio_id", "INTEGER")
    ensure_column(conn, "trades", "broker", "TEXT")
    ensure_column(conn, "trades", "asset_category", "TEXT")
    ensure_column(conn, "trades", "currency", "TEXT")
    ensure_column(conn, "trades", "symbol", "TEXT")
    ensure_column(conn, "trades", "trade_datetime", "TEXT")
    ensure_column(conn, "trades", "quantity", "REAL")
    ensure_column(conn, "trades", "side", "TEXT")
    ensure_column(conn, "trades", "trade_price", "REAL")
    ensure_column(conn, "trades", "close_price", "REAL")
    ensure_column(conn, "trades", "buy_price", "REAL")
    ensure_column(conn, "trades", "sell_price", "REAL")
    ensure_column(conn, "trades", "proceeds", "REAL")
    ensure_column(conn, "trades", "commission", "REAL")
    ensure_column(conn, "trades", "basis", "REAL")
    ensure_column(conn, "trades", "realized_pl", "REAL")
    ensure_column(conn, "trades", "mtm_pl", "REAL")
    ensure_column(conn, "trades", "code", "TEXT")
    ensure_column(conn, "trades", "import_file", "TEXT")
    ensure_column(conn, "trades", "batch_id", "TEXT")
    ensure_column(conn, "trades", "notes", "TEXT")
    ensure_column(conn, "trades", "trade_key", "TEXT")
    ensure_column(conn, "trades", "created_at", "TEXT")
    ensure_column(conn, "trades", "risk_amount", "REAL")
    ensure_column(conn, "trades", "r_multiple", "REAL")
    ensure_column(conn, "trade_journal", "mistake_tag", "TEXT")
    ensure_column(conn, "ibkr_settings", "flex_token", "TEXT")
    ensure_column(conn, "ibkr_settings", "query_id", "TEXT")
    ensure_column(conn, "ibkr_settings", "account_id", "TEXT")
    ensure_column(conn, "ibkr_settings", "report_format", "TEXT DEFAULT 'xml'")
    ensure_column(conn, "ibkr_settings", "auto_import_enabled", "INTEGER DEFAULT 0")
    ensure_column(conn, "ibkr_settings", "auto_import_hour", "INTEGER DEFAULT 6")
    ensure_column(conn, "ibkr_settings", "last_import_at", "TEXT")
    ensure_column(conn, "ibkr_settings", "last_auto_import_date", "TEXT")
    ensure_column(conn, "ibkr_settings", "last_auto_import_status", "TEXT")
    ensure_column(conn, "ibkr_settings", "updated_at", "TEXT")

    ensure_column(conn, "ninjatrader_settings", "folder_path", "TEXT")
    ensure_column(conn, "ninjatrader_settings", "auto_import_enabled", "INTEGER DEFAULT 0")
    ensure_column(conn, "ninjatrader_settings", "scan_interval_minutes", "INTEGER DEFAULT 10")
    ensure_column(conn, "ninjatrader_settings", "last_import_at", "TEXT")
    ensure_column(conn, "ninjatrader_settings", "last_scan_at", "TEXT")
    ensure_column(conn, "ninjatrader_settings", "last_scan_status", "TEXT")
    ensure_column(conn, "ninjatrader_settings", "updated_at", "TEXT")

    ensure_column(conn, "wealthsimple_settings", "folder_path", "TEXT")
    ensure_column(conn, "wealthsimple_settings", "auto_import_enabled", "INTEGER DEFAULT 0")
    ensure_column(conn, "wealthsimple_settings", "scan_interval_minutes", "INTEGER DEFAULT 10")
    ensure_column(conn, "wealthsimple_settings", "last_import_at", "TEXT")
    ensure_column(conn, "wealthsimple_settings", "last_scan_at", "TEXT")
    ensure_column(conn, "wealthsimple_settings", "last_scan_status", "TEXT")
    ensure_column(conn, "wealthsimple_settings", "updated_at", "TEXT")
    ensure_column(conn, "imported_files", "file_path", "TEXT")
    ensure_column(conn, "imported_files", "file_name", "TEXT")
    ensure_column(conn, "imported_files", "file_hash", "TEXT")
    ensure_column(conn, "imported_files", "file_size", "INTEGER")
    ensure_column(conn, "imported_files", "file_mtime", "REAL")
    ensure_column(conn, "imported_files", "batch_id", "TEXT")
    ensure_column(conn, "imported_files", "parsed_count", "INTEGER DEFAULT 0")
    ensure_column(conn, "imported_files", "inserted_count", "INTEGER DEFAULT 0")
    ensure_column(conn, "imported_files", "skipped_count", "INTEGER DEFAULT 0")
    ensure_column(conn, "imported_files", "status", "TEXT")

    # Duplicate protection uses a stable fingerprint generated from trade details.
    # We intentionally do NOT include batch_id/import_file because those change on every upload.
    # Keep this as a normal index because old databases may already contain duplicates,
    # and insert_trades() performs the actual duplicate check before inserting.
    try:
        cur.execute("DROP INDEX IF EXISTS idx_trade_unique")
        cur.execute("""
            CREATE INDEX IF NOT EXISTS idx_trade_key_lookup
            ON trades (user_id, portfolio_id, trade_key)
        """)
    except sqlite3.OperationalError:
        pass

    try:
        cur.execute("""
            CREATE INDEX IF NOT EXISTS idx_imported_files_lookup
            ON imported_files (user_id, source, file_hash)
        """)
    except sqlite3.OperationalError:
        pass

    backfill_buy_sell_prices(conn)

    conn.commit()
    conn.close()


# =========================================================
# Auth / session
# =========================================================
def hash_password(password, salt):
    return hashlib.sha256(f"{salt}:{password}".encode("utf-8")).hexdigest()


def create_user(email, password):
    conn = get_db_connection()
    cur = conn.cursor()
    salt = secrets.token_hex(16)
    password_hash = hash_password(password, salt)
    cur.execute(
        "INSERT INTO users (email, password_hash, salt) VALUES (?, ?, ?)",
        (email.strip().lower(), password_hash, salt)
    )
    user_id = cur.lastrowid
    cur.execute(
        "INSERT INTO portfolios (user_id, name) VALUES (?, ?)",
        (user_id, "Main Portfolio")
    )
    conn.commit()
    conn.close()
    return user_id


def authenticate_user(email, password):
    conn = get_db_connection()
    user = conn.execute(
        "SELECT * FROM users WHERE email = ?",
        (email.strip().lower(),)
    ).fetchone()
    conn.close()

    if not user:
        return None

    if hash_password(password, user["salt"]) == user["password_hash"]:
        return user["id"]
    return None


def current_user_id():
    return session.get("user_id")


def require_login():
    return current_user_id() is not None


def ensure_default_portfolio(user_id):
    """Return a valid portfolio id for user_id, creating Main Portfolio if missing.

    This prevents imports from failing with NOT NULL constraint failed: trades.portfolio_id
    for users created before portfolio creation existed, or accounts whose only portfolio
    was deleted/corrupted.
    """
    if not user_id:
        return None

    conn = get_db_connection()
    cur = conn.cursor()
    row = cur.execute(
        """
        SELECT id
        FROM portfolios
        WHERE user_id = ?
        ORDER BY id ASC
        LIMIT 1
        """,
        (user_id,)
    ).fetchone()

    if row:
        portfolio_id = row["id"]
    else:
        cur.execute(
            "INSERT INTO portfolios (user_id, name) VALUES (?, ?)",
            (user_id, "Main Portfolio")
        )
        portfolio_id = cur.lastrowid
        conn.commit()

    conn.close()
    session["portfolio_id"] = portfolio_id
    return portfolio_id


def current_portfolio_id():
    """Return the active portfolio for the logged-in user only.

    If the logged-in user has no portfolio yet, create a default Main Portfolio.
    This keeps imports from inserting NULL portfolio_id values.
    """
    user_id = current_user_id()
    if not user_id:
        return None

    pid = session.get("portfolio_id")
    if pid:
        try:
            pid_int = int(pid)
        except (TypeError, ValueError):
            pid_int = None

        if pid_int and ensure_portfolio_access(user_id, pid_int):
            return pid_int

        # Safety reset: never keep a portfolio id from another account.
        session.pop("portfolio_id", None)

    return ensure_default_portfolio(user_id)


def get_user_portfolios(user_id):
    conn = get_db_connection()
    rows = conn.execute("""
        SELECT id, name, created_at
        FROM portfolios
        WHERE user_id = ?
        ORDER BY name
    """, (user_id,)).fetchall()
    conn.close()
    return rows


def get_current_portfolio():
    """Return the active portfolio row for the logged-in user, creating/selecting a default if needed."""
    user_id = current_user_id()
    if not user_id:
        return None

    portfolio_id = current_portfolio_id()
    if not portfolio_id:
        return None

    conn = get_db_connection()
    row = conn.execute(
        """
        SELECT id, name, created_at
        FROM portfolios
        WHERE id = ? AND user_id = ?
        """,
        (portfolio_id, user_id),
    ).fetchone()
    conn.close()
    return row


def current_portfolio_name():
    portfolio = get_current_portfolio()
    return portfolio["name"] if portfolio else "No Portfolio"


@app.context_processor
def inject_layout_context():
    """Keep sidebar/topbar portfolio data consistent on every template-rendered page."""
    if not require_login():
        return {
            "portfolios": [],
            "active_portfolio_id": None,
            "active_portfolio_name": "No Portfolio",
        }

    active_id = current_portfolio_id()
    active = get_current_portfolio()
    return {
        "portfolios": get_user_portfolios(current_user_id()),
        "active_portfolio_id": active_id,
        "active_portfolio_name": active["name"] if active else "No Portfolio",
    }


def ensure_portfolio_access(user_id, portfolio_id):
    conn = get_db_connection()
    row = conn.execute("""
        SELECT id FROM portfolios
        WHERE id = ? AND user_id = ?
    """, (portfolio_id, user_id)).fetchone()
    conn.close()
    return row is not None


# =========================================================
# Helpers
# =========================================================
def to_decimal(value, default="0"):
    try:
        if value is None or str(value).strip() == "":
            return Decimal(default)
        raw = str(value).replace(",", "").replace("$", "").strip()
        negative = raw.startswith("(") and raw.endswith(")")
        if negative:
            raw = raw[1:-1].strip()
        out = Decimal(raw)
        return -out if negative else out
    except (InvalidOperation, ValueError):
        return Decimal(default)


def parse_dt_any(dt_raw):
    if not dt_raw:
        return None
    dt_raw = str(dt_raw).strip()
    if not dt_raw:
        return None

    formats = [
        "%Y-%m-%d, %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%Y-%m-%dT%H:%M:%S",
        "%m/%d/%Y",
        "%m/%d/%Y %H:%M:%S",
        "%Y/%m/%d",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(dt_raw, fmt)
        except ValueError:
            continue
    return None


def fmt_num(v):
    if v is None:
        return ""
    try:
        return f"{float(v):,.2f}"
    except Exception:
        return str(v)


def fmt_pct(v):
    try:
        return f"{float(v):.1f}%"
    except Exception:
        return "0.0%"


def fmt_dt(v):
    if not v:
        return ""
    try:
        return datetime.fromisoformat(v).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return v


def safe_like(v):
    return f"%{v.strip()}%"


def new_batch_id(prefix):
    return f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}"


def normalize_trade_key_value(value, places=8):
    if value is None:
        return ""
    text = str(value).strip()
    if text == "":
        return ""
    try:
        return f"{float(text):.{places}f}"
    except (TypeError, ValueError):
        return text.upper()


def make_trade_key(user_id, portfolio_id, trade):
    """Create a stable duplicate-detection key for one trade row.

    Excludes import_file, batch_id, created_at, notes, and journal fields.
    Those can change between uploads and should not make the same execution unique.
    """
    parts = [
        str(user_id or ""),
        str(portfolio_id or ""),
        str(trade.get("broker") or "").strip().upper(),
        str(trade.get("symbol") or "").strip().upper(),
        str(trade.get("trade_datetime") or "").strip(),
        normalize_trade_key_value(trade.get("quantity")),
        str(trade.get("side") or "").strip().upper(),
        normalize_trade_key_value(trade.get("trade_price")),
        normalize_trade_key_value(trade.get("proceeds")),
        normalize_trade_key_value(trade.get("commission")),
        str(trade.get("code") or "").strip().upper(),
    ]
    raw = "|".join(parts)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def derive_buy_sell_prices(trade):
    """Return explicit buy/sell prices for display/export.

    For closed NinjaTrader Performance rows, parsers provide exact buy_price and sell_price.
    For single execution imports, this derives the side price from trade_price.
    For short closed trades, trade_price is the entry sell and close_price is the cover buy.
    """
    def as_float(value):
        try:
            if value is None or str(value).strip() == "":
                return 0.0
            return float(value)
        except (TypeError, ValueError):
            return 0.0

    explicit_buy = as_float(trade.get("buy_price"))
    explicit_sell = as_float(trade.get("sell_price"))
    if abs(explicit_buy) > 1e-12 or abs(explicit_sell) > 1e-12:
        return explicit_buy, explicit_sell

    side = str(trade.get("side") or "").strip().upper()
    qty = as_float(trade.get("quantity"))
    trade_price = as_float(trade.get("trade_price"))
    close_price = as_float(trade.get("close_price"))

    buy_price = 0.0
    sell_price = 0.0

    if side == "BUY" or qty > 0:
        buy_price = trade_price
        if abs(close_price) > 1e-12:
            sell_price = close_price
    elif side == "SELL" or qty < 0:
        sell_price = trade_price
        if abs(close_price) > 1e-12:
            buy_price = close_price

    return buy_price, sell_price


def backfill_buy_sell_prices(conn):
    """Populate buy_price/sell_price for older rows after the migration adds columns."""
    rows = conn.execute("""
        SELECT id, side, quantity, trade_price, close_price, buy_price, sell_price
        FROM trades
        WHERE IFNULL(buy_price, 0) = 0 OR IFNULL(sell_price, 0) = 0
    """).fetchall()
    for r in rows:
        buy_price, sell_price = derive_buy_sell_prices(dict(r))
        conn.execute(
            "UPDATE trades SET buy_price = ?, sell_price = ? WHERE id = ?",
            (buy_price, sell_price, r["id"]),
        )


def parse_option_symbol(symbol):
    """Parse common option symbols/descriptions into grouping fields.

    Supports the original IBKR CSV format:
      AAPL 17JAN25 200 C

    Also supports many Flex/OCC-style labels:
      AAPL 17JAN25 200 CALL
      AAPL 250117C00200000
      AAPL   250117C00200000
      SPXW 06MAY2026 5700 PUT

    Returns None when the label is not an option contract.
    """
    if not symbol:
        return None

    raw = str(symbol).strip().upper()
    raw = re.sub(r"\s+", " ", raw)

    # Original app format: UNDERLYING 17JAN25 200 C
    m = OPTION_RE.match(raw)
    if m:
        return {
            "underlying": m.group("underlying"),
            "expiration": m.group("exp"),
            "strike": float(m.group("strike")),
            "cp": m.group("cp"),
        }

    # Human/Flex descriptions: UNDERLYING 17JAN25 200 CALL/PUT
    m = re.search(
        r"^(?P<underlying>[A-Z0-9.\-\/]+)\s+"
        r"(?P<exp>\d{1,2}[A-Z]{3}\d{2,4})\s+"
        r"(?P<strike>\d+(?:\.\d+)?)\s+"
        r"(?P<cp>C|P|CALL|PUT)$",
        raw,
    )
    if m:
        cp_raw = m.group("cp")
        return {
            "underlying": m.group("underlying"),
            "expiration": m.group("exp"),
            "strike": float(m.group("strike")),
            "cp": "C" if cp_raw.startswith("C") else "P",
        }

    # OCC compact format: AAPL250117C00200000 or AAPL 250117C00200000
    compact = raw.replace(" ", "")
    m = re.match(
        r"^(?P<underlying>[A-Z0-9.\-]+)(?P<yy>\d{2})(?P<mm>\d{2})(?P<dd>\d{2})(?P<cp>[CP])(?P<strike>\d{8})$",
        compact,
    )
    if m:
        yy = m.group("yy")
        mm = int(m.group("mm"))
        dd = int(m.group("dd"))
        months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
        exp = f"{dd:02d}{months[mm - 1]}{yy}" if 1 <= mm <= 12 else f"{yy}{mm:02d}{dd:02d}"
        return {
            "underlying": m.group("underlying"),
            "expiration": exp,
            "strike": int(m.group("strike")) / 1000.0,
            "cp": m.group("cp"),
        }

    return None


def allowed_file(filename, allowed_exts):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in allowed_exts


def base_where_and_params():
    user_id = current_user_id()
    portfolio_id = current_portfolio_id()
    if not user_id:
        where = ["1 = 0"]
        params = []
    else:
        if not portfolio_id:
            portfolio_id = ensure_default_portfolio(user_id)
        where = ["t.user_id = ?", "t.portfolio_id = ?"]
        params = [user_id, portfolio_id]
    return where, params


def build_filters():
    broker = request.args.get("broker", "").strip()
    symbol = request.args.get("symbol", "").strip()
    setup_tag = request.args.get("setup_tag", "").strip()
    trade_date = request.args.get("date", "").strip()

    where, params = base_where_and_params()

    if broker:
        where.append("t.broker = ?")
        params.append(broker)

    if symbol:
        where.append("t.symbol LIKE ?")
        params.append(safe_like(symbol))

    if setup_tag:
        where.append("IFNULL(j.setup_tag, '') = ?")
        params.append(setup_tag)

    if trade_date:
        where.append("substr(t.trade_datetime, 1, 10) = ?")
        params.append(trade_date)

    where_sql = "WHERE " + " AND ".join(where)
    return broker, symbol, setup_tag, where_sql, params


def scoped_where_sql(where_sql="", params=None, alias="t"):
    """Guarantee trade queries are scoped to the current user and portfolio."""
    params = list(params or [])
    if where_sql and params:
        return where_sql, params

    user_id = current_user_id()
    portfolio_id = current_portfolio_id()
    if alias:
        return f"WHERE {alias}.user_id = ? AND {alias}.portfolio_id = ?", [user_id, portfolio_id]
    return "WHERE user_id = ? AND portfolio_id = ?", [user_id, portfolio_id]

def broker_platform_name(broker):
    """Normalize broker/source names into the three main platforms shown in calendar totals."""
    b = (broker or "").strip().lower()
    if "wealthsimple" in b:
        return "Wealthsimple"
    if "ninja" in b or "performance" in b:
        return "NinjaTrader"
    if "ibkr" in b or "interactive" in b:
        return "IBKR"
    return "Other"


def get_calendar_heatmap():
    user_id = current_user_id()
    portfolio_id = current_portfolio_id()

    conn = get_db_connection()
    rows = conn.execute("""
        SELECT
            substr(trade_datetime, 1, 10) as day,
            broker,
            SUM(realized_pl) as pnl,
            COUNT(*) as trades
        FROM trades
        WHERE user_id = ? AND portfolio_id = ?
        GROUP BY day, broker
    """, (user_id, portfolio_id)).fetchall()
    conn.close()

    data = {}
    for r in rows:
        # Some imported broker rows can have a blank/unparseable trade_datetime.
        # Those rows are valid for tables/P&L, but they cannot be placed on a date calendar.
        day = r["day"]
        if not day:
            continue

        platform = broker_platform_name(r["broker"])
        if day not in data:
            data[day] = {
                "pnl": 0.0,
                "trades": 0,
                "platforms": {
                    "IBKR": {"pnl": 0.0, "trades": 0},
                    "NinjaTrader": {"pnl": 0.0, "trades": 0},
                    "Wealthsimple": {"pnl": 0.0, "trades": 0},
                    "Other": {"pnl": 0.0, "trades": 0},
                }
            }

        pnl = float(r["pnl"] or 0)
        trades = int(r["trades"] or 0)
        data[day]["pnl"] += pnl
        data[day]["trades"] += trades
        data[day]["platforms"][platform]["pnl"] += pnl
        data[day]["platforms"][platform]["trades"] += trades

    # Round P&L values for cleaner JSON/display.
    for info in data.values():
        info["pnl"] = round(float(info.get("pnl") or 0), 2)
        for platform_info in info.get("platforms", {}).values():
            platform_info["pnl"] = round(float(platform_info.get("pnl") or 0), 2)

    return data

# =========================================================
# Parsers
# =========================================================
def parse_ibkr_activity_csv(file_path):
    trades = []
    with open(file_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)

        for row in reader:
            if len(row) < 16:
                continue
            if row[0] != "Trades" or row[1] != "Data" or row[2] != "Order":
                continue

            asset_category = row[3].strip()
            currency = row[4].strip()
            symbol = row[5].strip()
            dt_raw = row[6].strip()

            quantity = to_decimal(row[7])
            trade_price = to_decimal(row[8])
            close_price = to_decimal(row[9])
            proceeds = to_decimal(row[10])
            commission = to_decimal(row[11])
            basis = to_decimal(row[12])
            realized_pl = to_decimal(row[13])
            mtm_pl = to_decimal(row[14])
            code = row[15].strip()

            trade_dt = parse_dt_any(dt_raw)
            side = "BUY" if quantity > 0 else "SELL"

            trades.append({
                "broker": "IBKR",
                "asset_category": asset_category,
                "currency": currency,
                "symbol": symbol,
                "trade_datetime": trade_dt.isoformat() if trade_dt else None,
                "quantity": float(quantity),
                "side": side,
                "trade_price": float(trade_price),
                "close_price": float(close_price),
                "buy_price": float(trade_price) if side == "BUY" else 0.0,
                "sell_price": float(trade_price) if side == "SELL" else 0.0,
                "proceeds": float(proceeds),
                "commission": float(commission),
                "basis": float(basis),
                "realized_pl": float(realized_pl),
                "mtm_pl": float(mtm_pl),
                "code": code,
                "notes": None,
                "risk_amount": 0.0,
                "r_multiple": 0.0,
            })
    return trades



def parse_ibkr_summary_money(value):
    raw = str(value or "").strip().replace(",", "")
    if raw in {"", "--"}:
        return 0.0

    negative = False
    if raw.startswith("(") and raw.endswith(")"):
        negative = True
        raw = raw[1:-1]

    raw = raw.replace("$", "").strip()
    try:
        amount = float(raw)
    except ValueError:
        amount = 0.0

    return -amount if negative else amount


def extract_ibkr_statement_date(file_path):
    with open(file_path, "r", encoding="utf-8-sig", errors="ignore", newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) >= 4 and row[0].strip() == "Statement" and row[1].strip() == "Data" and row[2].strip() == "Period":
                raw = row[3].strip()
                dt = parse_dt_any(raw)
                if dt:
                    return dt.isoformat()
                # IBKR may use "April 20, 2026"
                for fmt in ("%B %d, %Y", "%b %d, %Y"):
                    try:
                        return datetime.strptime(raw, fmt).isoformat()
                    except ValueError:
                        pass
    return datetime.now().isoformat()


def parse_ibkr_summary_csv(file_path):
    """
    Parses IBKR Activity Statement files that do not include raw Trades rows,
    but do include Realized & Unrealized Performance Summary rows.
    These are imported as synthetic closed P&L rows so dashboard/calendar/AI can work.
    """
    trades = []
    statement_dt = extract_ibkr_statement_date(file_path)

    with open(file_path, "r", encoding="utf-8-sig", errors="ignore", newline="") as f:
        reader = csv.reader(f)
        headers = None

        for row in reader:
            if not row:
                continue

            section = row[0].strip() if len(row) > 0 else ""
            row_type = row[1].strip() if len(row) > 1 else ""

            if section == "Realized & Unrealized Performance Summary" and row_type == "Header":
                headers = row
                continue

            if section == "Realized & Unrealized Performance Summary" and row_type == "Data" and headers:
                item = {headers[i]: row[i] if i < len(row) else "" for i in range(len(headers))}

                asset_category = item.get("Asset Category", "").strip()
                symbol = item.get("Symbol", "").strip()
                if not symbol or symbol.lower().startswith("total") or asset_category.lower() in {"total", "forex"}:
                    continue
                if "total" in symbol.lower():
                    continue

                realized_pl = parse_ibkr_summary_money(item.get("Realized Total", 0))
                if abs(realized_pl) < 1e-12:
                    continue

                code = item.get("Code", "").strip()
                trade_asset = "OPT" if "option" in asset_category.lower() else asset_category or "SUMMARY"

                trades.append({
                    "broker": "IBKR Summary",
                    "asset_category": trade_asset,
                    "currency": "CAD",
                    "symbol": symbol,
                    "trade_datetime": statement_dt,
                    "quantity": 0.0,
                    "side": "SUMMARY",
                    "trade_price": 0.0,
                    "close_price": 0.0,
                    "buy_price": 0.0,
                    "sell_price": 0.0,
                    "proceeds": 0.0,
                    "commission": 0.0,
                    "basis": 0.0,
                    "realized_pl": float(realized_pl),
                    "mtm_pl": float(realized_pl),
                    "code": code or "IBKR_SUMMARY",
                    "notes": "Imported from IBKR Realized & Unrealized Performance Summary. This is a synthetic P&L row, not a raw execution.",
                    "risk_amount": 0.0,
                    "r_multiple": 0.0,
                })

    return trades


def parse_wealthsimple_csv(file_path):
    trades = []

    with open(file_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        raw_fields = reader.fieldnames or []
        fields = {k.strip().lower(): k for k in raw_fields}

        def get(row, *names):
            for name in names:
                key = name.lower()
                if key in fields:
                    return row.get(fields[key], "")
            return ""

        for row in reader:
            activity_type = str(get(row, "activity_type", "type")).strip()
            if activity_type.upper() != "TRADE":
                continue

            symbol = str(get(row, "symbol", "ticker", "security")).strip()
            if not symbol:
                continue

            action = str(get(
                row, "activity_sub_type", "action", "transaction type", "type"
            )).strip().upper()

            qty = to_decimal(get(row, "quantity", "shares", "units"))
            price = to_decimal(get(row, "unit_price", "price", "fill price", "avg price"))
            fees = to_decimal(get(row, "commission", "fees"))
            proceeds = to_decimal(get(row, "net_cash_amount", "amount", "net amount", "proceeds", "value"))
            currency = str(get(row, "currency")).strip() or "CAD"

            # Wealthsimple exports can include both a trading date and a settlement/transaction date.
            # For P&L calendar/monthly analytics, use the actual trading/fill date first.
            # Settlement date is only a last-resort fallback so option trades do not land on the wrong day.
            trade_dt_raw = get(
                row,
                "trade_date", "trade date", "traded_at", "traded at",
                "execution_date", "execution date", "executed_at", "executed at",
                "fill_date", "fill date", "filled_at", "filled at",
                "order filled at", "order_filled_at"
            )
            fallback_dt_raw = get(
                row,
                "activity_date", "activity date", "transaction_date", "transaction date",
                "date", "datetime"
            )
            settlement_dt_raw = get(row, "settlement_date", "settlement date", "settled_at", "settled at")
            dt_raw = trade_dt_raw or fallback_dt_raw or settlement_dt_raw
            dt = parse_dt_any(dt_raw)

            side = "BUY"
            if "SELL" in action:
                side = "SELL"
            elif "BUY" in action:
                side = "BUY"
            elif qty < 0:
                side = "SELL"

            qty_abs = abs(float(qty))
            signed_qty = qty_abs if side == "BUY" else -qty_abs

            asset_category = "STK"
            if re.search(r"\b[CP]\b$", symbol.upper()):
                asset_category = "OPT"

            trades.append({
                "broker": "Wealthsimple",
                "asset_category": asset_category,
                "currency": currency,
                "symbol": symbol,
                "trade_datetime": dt.isoformat() if dt else None,
                "quantity": signed_qty,
                "side": side,
                "trade_price": float(price),
                "close_price": 0.0,
                "buy_price": float(price) if side == "BUY" else 0.0,
                "sell_price": float(price) if side == "SELL" else 0.0,
                "proceeds": float(proceeds),
                "commission": float(fees),
                "basis": 0.0,
                "realized_pl": 0.0,
                "mtm_pl": 0.0,
                "code": action,
                "notes": str(get(row, "name")).strip() or None,
                "risk_amount": 0.0,
                "r_multiple": 0.0,
            })
    return trades


def parse_performance_csv(file_path):
    """Parse NinjaTrader-style Performance CSV closed trades.

    Expected columns include:
    symbol, qty, buyPrice, sellPrice, pnl, boughtTimestamp, soldTimestamp, duration
    """
    trades = []

    with open(file_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        raw_fields = reader.fieldnames or []
        fields = {k.strip().lower(): k for k in raw_fields}

        def get(row, *names):
            for name in names:
                key = name.lower()
                if key in fields:
                    return row.get(fields[key], "")
            return ""

        for row in reader:
            symbol = str(get(row, "symbol")).strip()
            if not symbol:
                continue

            qty = abs(float(to_decimal(get(row, "qty", "quantity"))))
            buy_price = float(to_decimal(get(row, "buyPrice", "buy price")))
            sell_price = float(to_decimal(get(row, "sellPrice", "sell price")))
            realized_pl = float(to_decimal(get(row, "pnl", "P&L", "profit")))

            bought_raw = get(row, "boughtTimestamp", "bought timestamp", "buyTime", "buy time")
            sold_raw = get(row, "soldTimestamp", "sold timestamp", "sellTime", "sell time")
            bought_dt = parse_dt_any(bought_raw)
            sold_dt = parse_dt_any(sold_raw)

            # If buy happened first, this was a long trade. If sell happened first, short trade.
            is_long = True
            if bought_dt and sold_dt:
                is_long = bought_dt <= sold_dt

            trade_dt = sold_dt or bought_dt
            side = "BUY" if is_long else "SELL"
            signed_qty = qty if is_long else -qty
            entry_price = buy_price if is_long else sell_price
            close_price = sell_price if is_long else buy_price

            trades.append({
                "broker": "NinjaTrader Performance",
                "asset_category": "FUT",
                "currency": "USD",
                "symbol": symbol,
                "trade_datetime": trade_dt.isoformat() if trade_dt else None,
                "quantity": signed_qty,
                "side": side,
                "trade_price": entry_price,
                "close_price": close_price,
                "buy_price": buy_price,
                "sell_price": sell_price,
                "proceeds": realized_pl,
                "commission": 0.0,
                "basis": 0.0,
                "realized_pl": realized_pl,
                "mtm_pl": 0.0,
                "code": f"PERFORMANCE:{get(row, 'buyFillId')}:{get(row, 'sellFillId')}",
                "notes": f"Duration: {get(row, 'duration')}",
                "risk_amount": 0.0,
                "r_multiple": 0.0,
            })

    return trades



def detect_import_type(file_path):
    with open(file_path, "r", encoding="utf-8-sig", errors="ignore", newline="") as f:
        sample = f.read(120000)

    lowered = sample.lower().replace(" ", "")

    if "trades,data,order" in lowered or "trades,header,assetcategory" in lowered:
        return "ibkr_trades"

    if "realized&unrealizedperformancesummary" in lowered:
        return "ibkr_summary"

    if "activity_type" in lowered or "activitysubtype" in lowered or "activity_sub_type" in lowered:
        return "wealthsimple"

    if "buyprice" in lowered and "sellprice" in lowered and "pnl" in lowered:
        return "performance"

    return "unknown"



def import_success_page(title, broker_name, batch_id, parsed_count, inserted, skipped):
    body = f"""
<div class="glass-card">
    <h2 class="section-title">{title}</h2>
    <div class="grid-kpi" style="grid-template-columns: repeat(4, minmax(0, 1fr)); margin-top:14px;">
        <div class="kpi"><div class="kpi-label">Batch</div><div class="kpi-value mono" style="font-size:18px;">{batch_id}</div></div>
        <div class="kpi"><div class="kpi-label">Parsed Rows</div><div class="kpi-value">{parsed_count}</div></div>
        <div class="kpi"><div class="kpi-label">Inserted</div><div class="kpi-value pos">{inserted}</div></div>
        <div class="kpi"><div class="kpi-label">Duplicates</div><div class="kpi-value warn">{skipped}</div></div>
    </div>
    <div class="section-note" style="margin-top:12px;">Detected format: {broker_name}</div>
    <div style="margin-top:16px; display:flex; gap:10px;">
        <a class="btn" href="/dashboard">Back to Dashboard</a>
        <a class="btn secondary" href="/imports">View Imports</a>
    </div>
</div>
"""
    return page_shell("Import Complete", body)


# =========================================================
# FIFO Wealthsimple
# =========================================================
def get_existing_wealthsimple_trades():
    user_id = current_user_id()
    portfolio_id = current_portfolio_id()
    conn = get_db_connection()
    rows = conn.execute("""
        SELECT
            id, broker, asset_category, currency, symbol, trade_datetime, quantity,
            side, trade_price, close_price, buy_price, sell_price, proceeds, commission, basis, realized_pl,
            mtm_pl, risk_amount, r_multiple, code, import_file, batch_id, notes
        FROM trades
        WHERE user_id = ? AND portfolio_id = ? AND broker = 'Wealthsimple'
        ORDER BY trade_datetime ASC, id ASC
    """, (user_id, portfolio_id)).fetchall()
    conn.close()

    existing = []
    for r in rows:
        existing.append({
            "db_id": r["id"],
            "is_new": False,
            "broker": r["broker"],
            "asset_category": r["asset_category"],
            "currency": r["currency"],
            "symbol": r["symbol"],
            "trade_datetime": r["trade_datetime"],
            "quantity": float(r["quantity"] or 0),
            "side": r["side"],
            "trade_price": float(r["trade_price"] or 0),
            "close_price": float(r["close_price"] or 0),
            "buy_price": float(r["buy_price"] or 0),
            "sell_price": float(r["sell_price"] or 0),
            "proceeds": float(r["proceeds"] or 0),
            "commission": float(r["commission"] or 0),
            "basis": float(r["basis"] or 0),
            "realized_pl": float(r["realized_pl"] or 0),
            "mtm_pl": float(r["mtm_pl"] or 0),
            "risk_amount": float(r["risk_amount"] or 0),
            "r_multiple": float(r["r_multiple"] or 0),
            "code": r["code"],
            "import_file": r["import_file"],
            "batch_id": r["batch_id"],
            "notes": r["notes"],
        })
    return existing



def get_existing_wealthsimple_trades_for_user_portfolio(user_id, portfolio_id):
    conn = get_db_connection()
    rows = conn.execute("""
        SELECT
            id, broker, asset_category, currency, symbol, trade_datetime, quantity,
            side, trade_price, close_price, buy_price, sell_price, proceeds, commission, basis, realized_pl,
            mtm_pl, risk_amount, r_multiple, code, import_file, batch_id, notes
        FROM trades
        WHERE user_id = ? AND portfolio_id = ? AND broker = 'Wealthsimple'
        ORDER BY trade_datetime ASC, id ASC
    """, (user_id, portfolio_id)).fetchall()
    conn.close()

    existing = []
    for r in rows:
        existing.append({
            "db_id": r["id"],
            "is_new": False,
            "broker": r["broker"],
            "asset_category": r["asset_category"],
            "currency": r["currency"],
            "symbol": r["symbol"],
            "trade_datetime": r["trade_datetime"],
            "quantity": float(r["quantity"] or 0),
            "side": r["side"],
            "trade_price": float(r["trade_price"] or 0),
            "close_price": float(r["close_price"] or 0),
            "buy_price": float(r["buy_price"] or 0),
            "sell_price": float(r["sell_price"] or 0),
            "proceeds": float(r["proceeds"] or 0),
            "commission": float(r["commission"] or 0),
            "basis": float(r["basis"] or 0),
            "realized_pl": float(r["realized_pl"] or 0),
            "mtm_pl": float(r["mtm_pl"] or 0),
            "risk_amount": float(r["risk_amount"] or 0),
            "r_multiple": float(r["r_multiple"] or 0),
            "code": r["code"],
            "import_file": r["import_file"],
            "batch_id": r["batch_id"],
            "notes": r["notes"],
        })
    return existing


def recompute_fifo_for_wealthsimple(existing_trades, new_trades):
    combined = []

    for t in existing_trades:
        item = dict(t)
        item["is_new"] = False
        combined.append(item)

    for t in new_trades:
        item = dict(t)
        item["is_new"] = True
        item["db_id"] = None
        combined.append(item)

    grouped = defaultdict(list)
    for t in combined:
        grouped[(t["broker"], t["symbol"], t["currency"])].append(t)

    updates_for_existing = []

    for _, items in grouped.items():
        items.sort(key=lambda x: (
            x["trade_datetime"] or "",
            0 if (x["quantity"] or 0) > 0 else 1,
            x["db_id"] or 0
        ))

        buy_lots = deque()

        for trade in items:
            qty = float(trade["quantity"] or 0)
            price = float(trade["trade_price"] or 0)
            commission = float(trade["commission"] or 0)

            if qty > 0:
                total_cost = qty * price + commission
                unit_cost = total_cost / qty if qty else 0.0
                buy_lots.append({"qty_remaining": qty, "unit_cost": unit_cost})
                trade["basis"] = round(total_cost, 2)
                trade["realized_pl"] = 0.0

            elif qty < 0:
                sell_qty = abs(qty)
                gross_sale = sell_qty * price
                net_sale = gross_sale - commission

                qty_remaining_to_match = sell_qty
                total_basis = 0.0

                while qty_remaining_to_match > 1e-12 and buy_lots:
                    lot = buy_lots[0]
                    matched_qty = min(qty_remaining_to_match, lot["qty_remaining"])
                    total_basis += matched_qty * lot["unit_cost"]
                    lot["qty_remaining"] -= matched_qty
                    qty_remaining_to_match -= matched_qty
                    if lot["qty_remaining"] <= 1e-12:
                        buy_lots.popleft()

                trade["basis"] = round(total_basis, 2)
                trade["realized_pl"] = round(net_sale - total_basis, 2)

            if not trade["is_new"]:
                updates_for_existing.append({
                    "db_id": trade["db_id"],
                    "basis": float(trade["basis"] or 0),
                    "realized_pl": float(trade["realized_pl"] or 0),
                })

    new_trades_with_fifo = [t for t in combined if t["is_new"]]
    return updates_for_existing, new_trades_with_fifo


def update_existing_trade_fifo_values(updates):
    if not updates:
        return
    conn = get_db_connection()
    cur = conn.cursor()
    for u in updates:
        cur.execute(
            "UPDATE trades SET basis = ?, realized_pl = ? WHERE id = ?",
            (u["basis"], u["realized_pl"], u["db_id"])
        )
    conn.commit()
    conn.close()


# =========================================================
# Journal helpers
# =========================================================
def upsert_trade_journal(trade_id, setup_tag="", mistake_tag="", note="", screenshot_filename=None):
    conn = get_db_connection()
    cur = conn.cursor()

    existing = cur.execute(
        "SELECT id FROM trade_journal WHERE trade_id = ?",
        (trade_id,)
    ).fetchone()

    if existing:
        if screenshot_filename is None:
            cur.execute("""
                UPDATE trade_journal
                SET setup_tag = ?, mistake_tag = ?, note = ?, updated_at = CURRENT_TIMESTAMP
                WHERE trade_id = ?
            """, (setup_tag, mistake_tag, note, trade_id))
        else:
            cur.execute("""
                UPDATE trade_journal
                SET setup_tag = ?, mistake_tag = ?, note = ?, screenshot_filename = ?, updated_at = CURRENT_TIMESTAMP
                WHERE trade_id = ?
            """, (setup_tag, mistake_tag, note, screenshot_filename, trade_id))
    else:
        cur.execute("""
            INSERT INTO trade_journal (trade_id, setup_tag, mistake_tag, note, screenshot_filename)
            VALUES (?, ?, ?, ?, ?)
        """, (trade_id, setup_tag, mistake_tag, note, screenshot_filename))

    conn.commit()
    conn.close()


def get_setup_tags():
    user_id = current_user_id()
    portfolio_id = current_portfolio_id()
    conn = get_db_connection()
    rows = conn.execute("""
        SELECT DISTINCT j.setup_tag
        FROM trade_journal j
        JOIN trades t ON t.id = j.trade_id
        WHERE t.user_id = ? AND t.portfolio_id = ?
          AND j.setup_tag IS NOT NULL
          AND TRIM(j.setup_tag) != ''
        ORDER BY j.setup_tag
    """, (user_id, portfolio_id)).fetchall()
    conn.close()
    return [r["setup_tag"] for r in rows]


# =========================================================
# Trade insert/delete
# =========================================================
def insert_trades(trades, import_file=None, batch_id=None):
    user_id = current_user_id()
    portfolio_id = current_portfolio_id()

    if not user_id:
        raise ValueError("You must be logged in before importing trades.")
    if not portfolio_id:
        portfolio_id = ensure_default_portfolio(user_id)
    if not portfolio_id:
        raise ValueError("Could not create or select a portfolio for this user.")

    conn = get_db_connection()
    cur = conn.cursor()

    inserted = 0
    skipped = 0
    seen_in_this_upload = set()

    for t in trades:
        trade_key = make_trade_key(user_id, portfolio_id, t)

        # Skip duplicate rows inside the same CSV before touching the database.
        if trade_key in seen_in_this_upload:
            skipped += 1
            continue
        seen_in_this_upload.add(trade_key)

        # Primary duplicate check for rows created by the fixed app.
        existing = cur.execute("""
            SELECT id
            FROM trades
            WHERE user_id = ? AND portfolio_id = ? AND trade_key = ?
            LIMIT 1
        """, (user_id, portfolio_id, trade_key)).fetchone()

        # Fallback duplicate check for rows that existed before trade_key was added.
        # ROUND makes float formatting differences less likely to create false-new trades.
        if existing is None:
            existing = cur.execute("""
                SELECT id
                FROM trades
                WHERE user_id = ?
                  AND portfolio_id = ?
                  AND broker = ?
                  AND symbol = ?
                  AND IFNULL(trade_datetime, '') = IFNULL(?, '')
                  AND ROUND(IFNULL(quantity, 0), 8) = ROUND(IFNULL(?, 0), 8)
                  AND IFNULL(side, '') = IFNULL(?, '')
                  AND ROUND(IFNULL(trade_price, 0), 8) = ROUND(IFNULL(?, 0), 8)
                  AND ROUND(IFNULL(proceeds, 0), 8) = ROUND(IFNULL(?, 0), 8)
                  AND ROUND(IFNULL(commission, 0), 8) = ROUND(IFNULL(?, 0), 8)
                  AND IFNULL(code, '') = IFNULL(?, '')
                LIMIT 1
            """, (
                user_id,
                portfolio_id,
                t["broker"],
                t["symbol"],
                t["trade_datetime"],
                t["quantity"],
                t.get("side"),
                t["trade_price"],
                t["proceeds"],
                t["commission"],
                t["code"],
            )).fetchone()

        if existing is not None:
            skipped += 1
            continue

        buy_price, sell_price = derive_buy_sell_prices(t)

        cur.execute("""
            INSERT INTO trades (
                user_id, portfolio_id, broker, asset_category, currency, symbol,
                trade_datetime, quantity, side, trade_price, close_price, buy_price, sell_price, proceeds,
                commission, basis, realized_pl, mtm_pl, risk_amount, r_multiple,
                code, import_file, batch_id, notes, trade_key
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            user_id, portfolio_id, t["broker"], t["asset_category"], t["currency"],
            t["symbol"], t["trade_datetime"], t["quantity"], t["side"],
            t["trade_price"], t["close_price"], buy_price, sell_price, t["proceeds"], t["commission"],
            t["basis"], t["realized_pl"], t["mtm_pl"], t.get("risk_amount", 0),
            t.get("r_multiple", 0), t["code"], import_file, batch_id,
            t.get("notes"), trade_key
        ))
        inserted += 1

    conn.commit()
    conn.close()
    return inserted, skipped

def delete_batch(batch_id):
    user_id = current_user_id()
    portfolio_id = current_portfolio_id()

    conn = get_db_connection()
    cur = conn.cursor()

    trade_ids = cur.execute("""
        SELECT id
        FROM trades
        WHERE user_id = ? AND portfolio_id = ? AND batch_id = ?
    """, (user_id, portfolio_id, batch_id)).fetchall()

    for row in trade_ids:
        j = cur.execute(
            "SELECT screenshot_filename FROM trade_journal WHERE trade_id = ?",
            (row["id"],)
        ).fetchone()
        if j and j["screenshot_filename"]:
            path = os.path.join(SCREENSHOT_FOLDER, j["screenshot_filename"])
            if os.path.exists(path):
                try:
                    os.remove(path)
                except OSError:
                    pass

    cur.execute("""
        DELETE FROM trade_journal
        WHERE trade_id IN (
            SELECT id FROM trades
            WHERE user_id = ? AND portfolio_id = ? AND batch_id = ?
        )
    """, (user_id, portfolio_id, batch_id))

    cur.execute("""
        DELETE FROM trades
        WHERE user_id = ? AND portfolio_id = ? AND batch_id = ?
    """, (user_id, portfolio_id, batch_id))
    deleted = cur.rowcount

    # If this batch came from an automated folder import, remove the file-hash
    # tracking row too. Otherwise re-importing the corrected CSV after deleting
    # the batch will be skipped as "already imported" even though its trades
    # were deleted.
    try:
        cur.execute("""
            DELETE FROM imported_files
            WHERE user_id = ? AND batch_id = ?
        """, (user_id, batch_id))
    except sqlite3.OperationalError:
        # Older databases may not have imported_files yet.
        pass

    conn.commit()
    conn.close()
    return deleted


# =========================================================
# Queries / analytics
# =========================================================
def get_brokers():
    user_id = current_user_id()
    portfolio_id = current_portfolio_id()
    conn = get_db_connection()
    rows = conn.execute("""
        SELECT DISTINCT broker
        FROM trades
        WHERE user_id = ? AND portfolio_id = ? AND broker IS NOT NULL AND broker != ''
        ORDER BY broker
    """, (user_id, portfolio_id)).fetchall()
    conn.close()
    return [r["broker"] for r in rows]


def get_dashboard_totals(where_sql="", params=None):
    params = params or []
    where_sql, params = scoped_where_sql(where_sql, params, alias="t")
    conn = get_db_connection()
    row = conn.execute(f"""
        SELECT
            COUNT(*) AS trade_count,
            COALESCE(SUM(t.realized_pl), 0) AS realized_pl,
            COALESCE(SUM(t.commission), 0) AS total_fees,
            COALESCE(SUM(CASE WHEN t.proceeds < 0 THEN ABS(t.proceeds) ELSE 0 END), 0) AS gross_buys,
            COALESCE(SUM(CASE WHEN t.proceeds > 0 THEN t.proceeds ELSE 0 END), 0) AS gross_sells
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        {where_sql}
    """, params).fetchone()
    conn.close()
    return row


def get_recent_trades(where_sql="", params=None, limit=500):
    params = params or []
    where_sql, params = scoped_where_sql(where_sql, params, alias="t")
    conn = get_db_connection()
    rows = conn.execute(f"""
        SELECT
            t.*,
            j.setup_tag,
            j.mistake_tag,
            j.note AS journal_note,
            j.screenshot_filename
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        {where_sql}
        ORDER BY t.trade_datetime DESC, t.id DESC
        LIMIT ?
    """, params + [limit]).fetchall()
    conn.close()
    return rows


def get_trade_by_id(trade_id):
    user_id = current_user_id()
    portfolio_id = current_portfolio_id()
    conn = get_db_connection()
    row = conn.execute("""
        SELECT
            t.*,
            j.setup_tag,
            j.mistake_tag,
            j.note AS journal_note,
            j.screenshot_filename
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        WHERE t.id = ? AND t.user_id = ? AND t.portfolio_id = ?
    """, (trade_id, user_id, portfolio_id)).fetchone()
    conn.close()
    return row


def get_open_positions(where_sql="", params=None):
    params = params or []
    where_sql, params = scoped_where_sql(where_sql, params, alias="t")
    conn = get_db_connection()
    rows = conn.execute(f"""
        SELECT
            t.broker, t.symbol, t.asset_category, t.currency,
            ROUND(SUM(t.quantity), 4) AS net_qty,
            ROUND(SUM(CASE WHEN t.quantity > 0 THEN t.quantity * t.trade_price ELSE 0 END), 4) AS weighted_cost_value,
            COUNT(*) AS trade_rows
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        {where_sql}
        GROUP BY t.broker, t.symbol, t.asset_category, t.currency
        HAVING ABS(SUM(t.quantity)) > 0.000001
        ORDER BY t.broker, t.symbol
    """, params).fetchall()
    conn.close()

    out = []
    for r in rows:
        net_qty = float(r["net_qty"] or 0)
        weighted_cost_value = float(r["weighted_cost_value"] or 0)
        avg_cost = weighted_cost_value / net_qty if abs(net_qty) > 1e-9 else 0
        out.append({
            "broker": r["broker"],
            "symbol": r["symbol"],
            "asset_category": r["asset_category"],
            "currency": r["currency"],
            "net_qty": net_qty,
            "avg_cost": avg_cost,
            "trade_rows": r["trade_rows"],
        })
    return out


def get_monthly_pnl(where_sql="", params=None):
    params = params or []
    where_sql, params = scoped_where_sql(where_sql, params, alias="t")
    conn = get_db_connection()
    rows = conn.execute(f"""
        SELECT
            substr(t.trade_datetime, 1, 7) AS month,
            ROUND(COALESCE(SUM(t.realized_pl), 0), 2) AS realized_pl,
            ROUND(COALESCE(SUM(t.commission), 0), 2) AS fees,
            COUNT(*) AS trade_count
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        {where_sql}
        GROUP BY substr(t.trade_datetime, 1, 7)
        ORDER BY month
    """, params).fetchall()
    conn.close()
    return rows


def get_daily_pnl(where_sql="", params=None):
    params = params or []
    where_sql, params = scoped_where_sql(where_sql, params, alias="t")
    conn = get_db_connection()
    rows = conn.execute(f"""
        SELECT
            substr(t.trade_datetime, 1, 10) AS day,
            ROUND(COALESCE(SUM(t.realized_pl), 0), 2) AS realized_pl
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        {where_sql}
        GROUP BY substr(t.trade_datetime, 1, 10)
        ORDER BY day
    """, params).fetchall()
    conn.close()
    return rows


def get_weekday_stats(where_sql="", params=None):
    params = params or []
    where_sql, params = scoped_where_sql(where_sql, params, alias="t")
    conn = get_db_connection()
    rows = conn.execute(f"""
        SELECT
            CASE strftime('%w', substr(t.trade_datetime, 1, 10))
                WHEN '0' THEN 'Sunday'
                WHEN '1' THEN 'Monday'
                WHEN '2' THEN 'Tuesday'
                WHEN '3' THEN 'Wednesday'
                WHEN '4' THEN 'Thursday'
                WHEN '5' THEN 'Friday'
                WHEN '6' THEN 'Saturday'
            END AS weekday,
            ROUND(COALESCE(SUM(t.realized_pl), 0), 2) AS pnl,
            COUNT(*) AS trade_count
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        {where_sql}
        GROUP BY strftime('%w', substr(t.trade_datetime, 1, 10))
        ORDER BY strftime('%w', substr(t.trade_datetime, 1, 10))
    """, params).fetchall()
    conn.close()
    return rows


def get_broker_equity_curves(where_sql="", params=None):
    params = params or []
    where_sql, params = scoped_where_sql(where_sql, params, alias="t")
    conn = get_db_connection()
    rows = conn.execute(f"""
        SELECT
            t.broker,
            substr(t.trade_datetime, 1, 10) AS day,
            ROUND(COALESCE(SUM(t.realized_pl), 0), 2) AS pnl
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        {where_sql}
        GROUP BY t.broker, substr(t.trade_datetime, 1, 10)
        ORDER BY t.broker, day
    """, params).fetchall()
    conn.close()

    grouped = {}
    for r in rows:
        broker = r["broker"] or "Unknown"
        grouped.setdefault(broker, [])
        grouped[broker].append((r["day"], float(r["pnl"] or 0)))

    out = {}
    for broker, series in grouped.items():
        running = 0.0
        labels = []
        values = []
        for day, pnl in series:
            running += pnl
            labels.append(day)
            values.append(round(running, 2))
        out[broker] = {"labels": labels, "values": values}
    return out


def get_import_history():
    user_id = current_user_id()
    portfolio_id = current_portfolio_id()
    conn = get_db_connection()
    rows = conn.execute("""
        SELECT
            batch_id,
            broker,
            import_file,
            MIN(created_at) AS imported_at,
            COUNT(*) AS row_count,
            ROUND(COALESCE(SUM(realized_pl), 0), 2) AS realized_pl,
            ROUND(COALESCE(SUM(commission), 0), 2) AS fees
        FROM trades
        WHERE user_id = ? AND portfolio_id = ?
          AND batch_id IS NOT NULL AND batch_id != ''
        GROUP BY batch_id, broker, import_file
        ORDER BY imported_at DESC
    """, (user_id, portfolio_id)).fetchall()
    conn.close()
    return rows


def get_option_strategy_groups(where_sql="", params=None):
    params = params or []
    where_sql, params = scoped_where_sql(where_sql, params, alias="t")
    conn = get_db_connection()
    rows = conn.execute(f"""
        SELECT
            t.*,
            j.setup_tag,
            j.mistake_tag,
            j.note AS journal_note,
            j.screenshot_filename
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        {where_sql}
        ORDER BY t.trade_datetime DESC, t.id DESC
    """, params).fetchall()
    conn.close()

    grouped = defaultdict(list)

    for r in rows:
        asset = (r["asset_category"] or "").upper()
        if "OPT" not in asset and "OPTION" not in asset:
            continue

        parsed = parse_option_symbol(r["symbol"])
        if not parsed:
            continue

        dt_key = (r["trade_datetime"] or "")[:10]
        key = (r["broker"], parsed["underlying"], parsed["expiration"], parsed["cp"], dt_key)
        grouped[key].append(r)

    results = []
    for key, legs in grouped.items():
        broker, underlying, expiration, cp, dt_key = key
        parsed_legs = [parse_option_symbol(l["symbol"]) for l in legs]
        strikes = sorted(set(p["strike"] for p in parsed_legs if p))
        leg_count = len(legs)
        total_proceeds = sum(float(l["proceeds"] or 0) for l in legs)
        total_fees = sum(float(l["commission"] or 0) for l in legs)
        total_realized = sum(float(l["realized_pl"] or 0) for l in legs)

        strategy_type = "Custom"
        unique_strikes = len(strikes)
        if leg_count == 1:
            strategy_type = "Single Option"
        elif leg_count == 2 and unique_strikes == 2:
            strategy_type = "Vertical Spread"
        elif leg_count == 4 and unique_strikes == 4:
            strategy_type = "Iron Condor / 4-Leg"
        elif leg_count == 2 and unique_strikes == 1:
            strategy_type = "Same Strike Pair"

        results.append({
            "broker": broker,
            "underlying": underlying,
            "expiration": expiration,
            "option_type": "CALL" if cp == "C" else "PUT",
            "trade_date": dt_key,
            "strategy_type": strategy_type,
            "leg_count": leg_count,
            "strikes": ", ".join(str(int(s) if float(s).is_integer() else s) for s in strikes),
            "net_proceeds": total_proceeds,
            "fees": total_fees,
            "realized_pl": total_realized,
            "legs": legs,
        })

    results.sort(key=lambda x: (x["trade_date"], x["underlying"]), reverse=True)
    return results

def get_strategy_performance(where_sql="", params=None):
    params = params or []
    where_sql, params = scoped_where_sql(where_sql, params, alias="t")
    conn = get_db_connection()
    rows = conn.execute(f"""
        SELECT
            COALESCE(NULLIF(TRIM(j.setup_tag), ''), 'Unlabeled') AS setup_tag,
            COUNT(*) AS trades,
            ROUND(COALESCE(SUM(t.realized_pl), 0), 2) AS pnl,
            ROUND(COALESCE(AVG(t.realized_pl), 0), 2) AS avg_pnl,
            ROUND(COALESCE(AVG(CASE WHEN ABS(COALESCE(t.r_multiple, 0)) > 0 THEN t.r_multiple END), 0), 2) AS avg_r
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        {where_sql}
        GROUP BY COALESCE(NULLIF(TRIM(j.setup_tag), ''), 'Unlabeled')
        ORDER BY pnl DESC, trades DESC
    """, params).fetchall()
    conn.close()
    return rows


def get_gallery_rows():
    user_id = current_user_id()
    portfolio_id = current_portfolio_id()
    conn = get_db_connection()
    rows = conn.execute("""
        SELECT
            t.id,
            t.symbol,
            t.trade_datetime,
            t.broker,
            t.realized_pl,
            j.setup_tag,
            j.screenshot_filename
        FROM trades t
        JOIN trade_journal j ON j.trade_id = t.id
        WHERE t.user_id = ? AND t.portfolio_id = ?
          AND j.screenshot_filename IS NOT NULL
          AND TRIM(j.screenshot_filename) != ''
        ORDER BY t.trade_datetime DESC, t.id DESC
    """, (user_id, portfolio_id)).fetchall()
    conn.close()
    return rows


def get_analytics(where_sql="", params=None):
    params = params or []
    where_sql, params = scoped_where_sql(where_sql, params, alias="t")
    conn = get_db_connection()

    rows = conn.execute(f"""
        SELECT
            t.id, t.symbol, t.trade_datetime, t.realized_pl, t.commission, t.broker,
            t.r_multiple, j.setup_tag, j.mistake_tag
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        {where_sql}
        ORDER BY t.trade_datetime ASC, t.id ASC
    """, params).fetchall()
    conn.close()

    realized_rows = [r for r in rows if abs(float(r["realized_pl"] or 0)) > 1e-12]

    total_trades = len(rows)
    closed_trades = len(realized_rows)

    wins = [float(r["realized_pl"]) for r in realized_rows if float(r["realized_pl"]) > 0]
    losses = [float(r["realized_pl"]) for r in realized_rows if float(r["realized_pl"]) < 0]

    win_count = len(wins)
    loss_count = len(losses)

    gross_profit = sum(wins)
    gross_loss_abs = abs(sum(losses))
    net_profit = sum(float(r["realized_pl"] or 0) for r in rows)

    win_rate = (win_count / closed_trades * 100) if closed_trades else 0
    avg_win = (gross_profit / win_count) if win_count else 0
    avg_loss = (sum(losses) / loss_count) if loss_count else 0
    profit_factor = (gross_profit / gross_loss_abs) if gross_loss_abs > 0 else 0
    expectancy = (net_profit / closed_trades) if closed_trades else 0

    best_trade = max(realized_rows, key=lambda r: float(r["realized_pl"]), default=None)
    worst_trade = min(realized_rows, key=lambda r: float(r["realized_pl"]), default=None)

    daily_map = defaultdict(float)
    for r in rows:
        day = (r["trade_datetime"] or "")[:10]
        if day:
            daily_map[day] += float(r["realized_pl"] or 0)

    equity_labels = sorted(daily_map.keys())
    equity_curve = []
    running = 0.0
    peak = 0.0
    max_drawdown = 0.0

    for d in equity_labels:
        running += daily_map[d]
        equity_curve.append(round(running, 2))
        peak = max(peak, running)
        dd = running - peak
        if dd < max_drawdown:
            max_drawdown = dd

    setup_rows = defaultdict(lambda: {"count": 0, "pnl": 0.0})
    for r in rows:
        tag = (r["setup_tag"] or "Unlabeled").strip() or "Unlabeled"
        setup_rows[tag]["count"] += 1
        setup_rows[tag]["pnl"] += float(r["realized_pl"] or 0)

    setup_r_rows = defaultdict(lambda: {"count": 0, "pnl": 0.0, "r_sum": 0.0, "r_count": 0})
    for r in rows:
        tag = (r["setup_tag"] or "Unlabeled").strip() or "Unlabeled"
        setup_r_rows[tag]["count"] += 1
        setup_r_rows[tag]["pnl"] += float(r["realized_pl"] or 0)
        r_mult = float(r["r_multiple"] or 0)
        if abs(r_mult) > 1e-12:
            setup_r_rows[tag]["r_sum"] += r_mult
            setup_r_rows[tag]["r_count"] += 1

    setup_stats = sorted(
        [{
            "setup_tag": k,
            "count": v["count"],
            "pnl": round(v["pnl"], 2),
            "avg_r": round((v["r_sum"] / v["r_count"]), 2) if v["r_count"] else 0.0
        } for k, v in setup_r_rows.items()],
        key=lambda x: x["pnl"],
        reverse=True
    )

    mistake_rows = defaultdict(lambda: {"count": 0, "pnl": 0.0})
    for r in rows:
        mistake = (r["mistake_tag"] or "No Mistake").strip() or "No Mistake"
        mistake_rows[mistake]["count"] += 1
        mistake_rows[mistake]["pnl"] += float(r["realized_pl"] or 0)

    mistake_stats = sorted(
        [{"mistake_tag": k, "count": v["count"], "pnl": round(v["pnl"], 2)} for k, v in mistake_rows.items()],
        key=lambda x: x["count"],
        reverse=True
    )

    return {
        "total_trades": total_trades,
        "closed_trades": closed_trades,
        "win_count": win_count,
        "loss_count": loss_count,
        "win_rate": win_rate,
        "gross_profit": gross_profit,
        "gross_loss_abs": gross_loss_abs,
        "net_profit": net_profit,
        "avg_win": avg_win,
        "avg_loss": avg_loss,
        "profit_factor": profit_factor,
        "expectancy": expectancy,
        "best_trade": best_trade,
        "worst_trade": worst_trade,
        "equity_labels": equity_labels,
        "equity_curve": equity_curve,
        "max_drawdown": max_drawdown,
        "setup_stats": setup_stats,
        "mistake_stats": mistake_stats,
    }


def generate_ai_review_from_rows(rows):
    """Rule-based AI-style trading coach.

    This does not call an external AI API. It turns your journal data into
    coach-style diagnostics, scores, warnings, and next-step suggestions.
    """
    if not rows:
        return {
            "summary": "No trades available yet.",
            "insights": ["Start journaling trades to unlock insights."],
            "warnings": [],
            "next_focus": ["Import trades", "Tag setups", "Add risk amounts"],
            "playbook": [],
            "setup_scores": [],
            "risk_flags": [],
            "ai_score": 0,
            "grade": "N/A",
            "edge_label": "No data yet",
            "metrics": {"total_trades": 0, "win_rate": 0, "avg_win": 0, "avg_loss": 0, "net_pnl": 0, "max_loss_streak": 0}
        }

    total = len(rows)
    pnl_list = [float(r.get("realized_pl") or 0) for r in rows]
    wins = [p for p in pnl_list if p > 0]
    losses = [p for p in pnl_list if p < 0]
    closed_count = len(wins) + len(losses)
    win_rate = (len(wins) / closed_count * 100) if closed_count else 0
    avg_win = sum(wins) / len(wins) if wins else 0
    avg_loss = sum(losses) / len(losses) if losses else 0
    net_pnl = sum(pnl_list)
    gross_profit = sum(wins)
    gross_loss_abs = abs(sum(losses))
    profit_factor = (gross_profit / gross_loss_abs) if gross_loss_abs else (999 if gross_profit > 0 else 0)
    expectancy = (net_pnl / closed_count) if closed_count else 0

    insights = []
    warnings = []
    next_focus = []
    playbook = []
    risk_flags = []

    if win_rate < 40:
        warnings.append(f"Win rate is low at {win_rate:.1f}%. Be more selective and reduce marginal entries.")
        next_focus.append("Tighten entry checklist before taking new trades")
    elif win_rate > 60:
        insights.append(f"Strong win rate at {win_rate:.1f}%. Keep execution consistent and avoid overconfidence.")

    if losses and abs(avg_loss) > max(avg_win, 1e-9):
        warnings.append(f"Average loss (${abs(avg_loss):.0f}) is larger than average win (${avg_win:.0f}). Risk/reward needs attention.")
        next_focus.append("Cut losers faster or only take setups with better reward-to-risk")

    if profit_factor < 1 and closed_count >= 10:
        warnings.append(f"Profit factor is below 1.00 ({profit_factor:.2f}). Your winners are not covering losses yet.")
    elif profit_factor >= 1.5 and closed_count >= 10:
        insights.append(f"Profit factor is healthy at {profit_factor:.2f}. Your edge is showing up in the sample.")

    streak = 0
    max_loss_streak = 0
    max_win_streak = 0
    win_streak = 0
    for p in pnl_list:
        if p < 0:
            streak += 1
            win_streak = 0
            max_loss_streak = max(max_loss_streak, streak)
        elif p > 0:
            win_streak += 1
            streak = 0
            max_win_streak = max(max_win_streak, win_streak)
        else:
            streak = 0
            win_streak = 0

    if max_loss_streak >= 3:
        warnings.append(f"Longest losing streak was {max_loss_streak} trades. Add a cooldown rule after 2 losses.")
        next_focus.append("Use a mandatory cooldown after two consecutive losses")

    daily_counts = defaultdict(int)
    daily_pnl = defaultdict(float)
    by_day = defaultdict(list)
    for r in rows:
        day = (r.get("trade_datetime") or "")[:10]
        if not day:
            continue
        pnl = float(r.get("realized_pl") or 0)
        daily_counts[day] += 1
        daily_pnl[day] += pnl
        by_day[day].append(r)

    overtrade_days = [d for d, count in daily_counts.items() if count >= 5]
    if overtrade_days:
        overtrade_pnl = sum(daily_pnl[d] for d in overtrade_days)
        msg = f"Overtrading detected: {len(overtrade_days)} day(s) had 5+ trades, totaling ${overtrade_pnl:.0f}."
        (warnings if overtrade_pnl < 0 else insights).append(msg)
        next_focus.append("Set a max-trades-per-day rule")

    revenge_days = [d for d in daily_counts if daily_counts[d] >= 4 and daily_pnl[d] < 0]
    if revenge_days:
        warnings.append(f"Possible revenge-trading pattern: {len(revenge_days)} losing day(s) had 4+ trades.")
        next_focus.append("Stop trading for the day after your rule break or second loss")

    bad_days = [d for d in daily_pnl if daily_pnl[d] < 0]
    avg_losing_day = sum(daily_pnl[d] for d in bad_days) / len(bad_days) if bad_days else 0
    if bad_days:
        insights.append(f"Average losing day is ${avg_losing_day:.0f}. Use this to set a realistic daily stop.")

    setup_map = defaultdict(float)
    setup_counts = defaultdict(int)
    setup_wins = defaultdict(int)
    setup_losses = defaultdict(int)
    setup_r_sum = defaultdict(float)
    setup_r_count = defaultdict(int)
    for r in rows:
        tag = (r.get("setup_tag") or "Unlabeled").strip() or "Unlabeled"
        pnl = float(r.get("realized_pl") or 0)
        setup_map[tag] += pnl
        setup_counts[tag] += 1
        if pnl > 0:
            setup_wins[tag] += 1
        elif pnl < 0:
            setup_losses[tag] += 1
        r_mult = float(r.get("r_multiple") or 0)
        if abs(r_mult) > 1e-12:
            setup_r_sum[tag] += r_mult
            setup_r_count[tag] += 1

    setup_scores = []
    for tag, pnl in setup_map.items():
        count = setup_counts[tag]
        win_pct = (setup_wins[tag] / max(setup_wins[tag] + setup_losses[tag], 1)) * 100
        avg_r = setup_r_sum[tag] / setup_r_count[tag] if setup_r_count[tag] else 0
        score = 50
        score += 20 if pnl > 0 else -20 if pnl < 0 else 0
        score += 10 if win_pct >= 55 else -10 if win_pct < 40 and count >= 5 else 0
        score += 10 if avg_r > 0.5 else -10 if avg_r < -0.25 and setup_r_count[tag] else 0
        score += 5 if count >= 5 else -5
        score = max(0, min(100, int(score)))
        setup_scores.append({"setup": tag, "score": score, "trades": count, "pnl": round(pnl, 2), "win_rate": round(win_pct, 1), "avg_r": round(avg_r, 2)})
    setup_scores.sort(key=lambda x: (x["score"], x["pnl"]), reverse=True)

    if setup_scores:
        best = setup_scores[0]
        worst = setup_scores[-1]
        insights.append(f"Best setup to study: {best['setup']} scored {best['score']}/100 with ${best['pnl']:.0f} P&L.")
        if worst["pnl"] < 0 and worst["trades"] >= 2:
            warnings.append(f"Weakest setup: {worst['setup']} scored {worst['score']}/100 with ${worst['pnl']:.0f} P&L.")
            next_focus.append(f"Review or pause setup: {worst['setup']}")

    mistake_map = defaultdict(float)
    mistake_counts = defaultdict(int)
    for r in rows:
        tag = (r.get("mistake_tag") or "No Mistake").strip() or "No Mistake"
        if tag == "No Mistake":
            continue
        mistake_map[tag] += float(r.get("realized_pl") or 0)
        mistake_counts[tag] += 1
    if mistake_map:
        worst_mistake = min(mistake_map, key=mistake_map.get)
        warnings.append(f"Most expensive mistake: {worst_mistake} (${mistake_map[worst_mistake]:.0f} across {mistake_counts[worst_mistake]} trade(s)).")
        next_focus.append(f"Create a pre-trade block for: {worst_mistake}")

    weekday_map = defaultdict(float)
    weekday_counts = defaultdict(int)
    for r in rows:
        dt = (r.get("trade_datetime") or "")[:10]
        if not dt:
            continue
        try:
            weekday = datetime.fromisoformat(dt).strftime("%A")
            weekday_map[weekday] += float(r.get("realized_pl") or 0)
            weekday_counts[weekday] += 1
        except Exception:
            pass
    if weekday_map:
        best_day = max(weekday_map, key=weekday_map.get)
        worst_day = min(weekday_map, key=weekday_map.get)
        insights.append(f"Best weekday: {best_day} (${weekday_map[best_day]:.0f} across {weekday_counts[best_day]} trade(s)).")
        if weekday_map[worst_day] < 0:
            warnings.append(f"Worst weekday: {worst_day} (${weekday_map[worst_day]:.0f} across {weekday_counts[worst_day]} trade(s)).")

    actual_total = 0.0
    simulated_total = 0.0
    for _, day_rows in by_day.items():
        day_rows = sorted(day_rows, key=lambda r: (r.get("trade_datetime") or ""))
        losses_seen = 0
        for r in day_rows:
            pnl = float(r.get("realized_pl") or 0)
            actual_total += pnl
            if losses_seen < 2:
                simulated_total += pnl
            if pnl < 0:
                losses_seen += 1
    improvement = simulated_total - actual_total
    if improvement > 0:
        insights.append(f"What-if: stopping after 2 losing trades per day could have improved results by about ${improvement:.0f}.")

    risk_values = [float(r.get("risk_amount") or 0) for r in rows if abs(float(r.get("risk_amount") or 0)) > 1e-12]
    r_values = [float(r.get("r_multiple") or 0) for r in rows if abs(float(r.get("r_multiple") or 0)) > 1e-12]
    if len(risk_values) < max(3, total * 0.25):
        risk_flags.append("Most trades do not have risk amount logged yet. Add risk to unlock better R-multiple coaching.")
        next_focus.append("Log risk amount on every trade")
    if risk_values:
        avg_risk = sum(risk_values) / len(risk_values)
        max_risk = max(risk_values)
        if avg_risk > 0 and max_risk > avg_risk * 2.5:
            risk_flags.append(f"Position sizing is inconsistent: max risk (${max_risk:.0f}) is over 2.5x average risk (${avg_risk:.0f}).")
            next_focus.append("Normalize position sizing before increasing size")
    if r_values:
        avg_r_multiple = sum(r_values) / len(r_values)
        if avg_r_multiple < 0:
            risk_flags.append(f"Average R is negative ({avg_r_multiple:.2f}R). The process is not paying for risk yet.")
        elif avg_r_multiple >= 0.5:
            insights.append(f"Average R is strong at {avg_r_multiple:.2f}R across logged-risk trades.")

    if setup_scores:
        top_setups = [s for s in setup_scores if s["score"] >= 65 and s["trades"] >= 3]
        if top_setups:
            playbook.append(f"Prioritize: {', '.join(s['setup'] for s in top_setups[:3])}.")
    if warnings:
        playbook.append("Before each trade: confirm setup, stop level, target, and max daily loss.")
    if overtrade_days or revenge_days:
        playbook.append("Daily rule: stop after 2 losses or 5 total trades, whichever comes first.")
    if not playbook:
        playbook.append("Keep collecting data; your next edge comes from better tags and risk notes.")

    score = 50
    score += 20 if net_pnl > 0 else -20 if net_pnl < 0 else 0
    score += 15 if profit_factor >= 1.5 else 5 if profit_factor >= 1.1 else -15 if closed_count >= 10 else 0
    score += 10 if win_rate >= 50 else -10 if win_rate < 40 and closed_count >= 10 else 0
    score += 10 if expectancy > 0 else -10 if expectancy < 0 and closed_count >= 10 else 0
    score -= min(15, max_loss_streak * 3)
    score -= min(10, len(overtrade_days) * 2)
    score = max(0, min(100, int(score)))
    grade = "A" if score >= 85 else "B" if score >= 70 else "C" if score >= 55 else "D" if score >= 40 else "F"
    edge_label = "Strong edge" if score >= 80 else "Developing edge" if score >= 60 else "Needs tightening" if score >= 40 else "High risk / low edge"

    if not insights:
        insights.append("No major issues detected yet. Keep logging more trades for stronger pattern detection.")

    def unique(items):
        seen = set()
        out = []
        for item in items:
            if item and item not in seen:
                seen.add(item)
                out.append(item)
        return out

    return {
        "summary": f"{total} trades analyzed | Win rate: {win_rate:.1f}% | Net P&L: ${net_pnl:.0f} | Coach score: {score}/100",
        "insights": unique(insights),
        "warnings": unique(warnings),
        "next_focus": unique(next_focus)[:6],
        "playbook": unique(playbook)[:5],
        "setup_scores": setup_scores[:8],
        "risk_flags": unique(risk_flags),
        "ai_score": score,
        "grade": grade,
        "edge_label": edge_label,
        "metrics": {"total_trades": total, "closed_trades": closed_count, "win_rate": win_rate, "avg_win": avg_win, "avg_loss": avg_loss, "net_pnl": net_pnl, "profit_factor": profit_factor, "expectancy": expectancy, "max_loss_streak": max_loss_streak, "max_win_streak": max_win_streak, "avg_losing_day": avg_losing_day}
    }


def generate_deep_ai_review(rows, base_review=None):
    """Second-layer coach analytics: patterns, profiles, weekly review, and action plan."""
    base_review = base_review or {}
    if not rows:
        return {
            "ai_confidence": 0,
            "discipline_profile": "No data yet",
            "pattern_cards": [],
            "symbol_scores": [],
            "weekly_review": [],
            "avoid_list": [],
            "action_plan": [
                {"priority": "High", "action": "Import trades", "reason": "The coach needs a trade sample before it can find patterns."},
                {"priority": "High", "action": "Tag setups", "reason": "Setup labels unlock edge detection."},
                {"priority": "Medium", "action": "Log risk per trade", "reason": "Risk values unlock R-multiple coaching."},
            ],
            "best_conditions": [],
            "worst_conditions": [],
            "coach_questions": [],
            "execution_summary": "Import trades to generate deeper coaching."
        }

    parsed = []
    for r in rows:
        pnl = float(r.get("realized_pl") or 0)
        dt_raw = r.get("trade_datetime") or ""
        dt = None
        try:
            dt = datetime.fromisoformat(dt_raw[:19]) if dt_raw else None
        except Exception:
            dt = None
        setup = (r.get("setup_tag") or "Unlabeled").strip() or "Unlabeled"
        mistake = (r.get("mistake_tag") or "No Mistake").strip() or "No Mistake"
        symbol = (r.get("symbol") or "Unknown").strip() or "Unknown"
        broker = (r.get("broker") or "Unknown").strip() or "Unknown"
        risk = float(r.get("risk_amount") or 0)
        r_mult = float(r.get("r_multiple") or 0)
        parsed.append({**dict(r), "pnl": pnl, "dt": dt, "setup": setup, "mistake": mistake, "symbol": symbol, "broker": broker, "risk": risk, "r_mult": r_mult})

    total = len(parsed)
    tagged = [r for r in parsed if r["setup"] != "Unlabeled"]
    risk_logged = [r for r in parsed if abs(r["risk"]) > 1e-12]
    closed = [r for r in parsed if abs(r["pnl"]) > 1e-12]
    tag_coverage = len(tagged) / total if total else 0
    risk_coverage = len(risk_logged) / total if total else 0
    sample_score = min(1.0, total / 50)
    ai_confidence = int(round((sample_score * 0.45 + tag_coverage * 0.30 + risk_coverage * 0.25) * 100))

    def group_stats(key_fn):
        buckets = defaultdict(lambda: {"trades": 0, "pnl": 0.0, "wins": 0, "losses": 0})
        for r in parsed:
            key = key_fn(r)
            if not key:
                continue
            b = buckets[key]
            b["trades"] += 1
            b["pnl"] += r["pnl"]
            if r["pnl"] > 0:
                b["wins"] += 1
            elif r["pnl"] < 0:
                b["losses"] += 1
        out = []
        for key, b in buckets.items():
            decided = b["wins"] + b["losses"]
            win_rate = (b["wins"] / decided * 100) if decided else 0
            avg_pnl = b["pnl"] / b["trades"] if b["trades"] else 0
            out.append({"name": key, "trades": b["trades"], "pnl": round(b["pnl"], 2), "avg_pnl": round(avg_pnl, 2), "win_rate": round(win_rate, 1)})
        return out

    symbol_scores = group_stats(lambda r: r["symbol"])
    symbol_scores.sort(key=lambda x: (x["pnl"], x["avg_pnl"]), reverse=True)

    weekday_stats = group_stats(lambda r: r["dt"].strftime("%A") if r["dt"] else None)
    time_stats = group_stats(lambda r: (
        "Morning" if r["dt"] and r["dt"].hour < 12 else
        "Midday" if r["dt"] and r["dt"].hour < 15 else
        "Late Day" if r["dt"] else None
    ))
    broker_stats = group_stats(lambda r: r["broker"])
    mistake_stats = [x for x in group_stats(lambda r: None if r["mistake"] == "No Mistake" else r["mistake"]) if x["trades"]]
    mistake_stats.sort(key=lambda x: x["pnl"])

    weekly = defaultdict(lambda: {"trades": 0, "pnl": 0.0, "wins": 0, "losses": 0})
    for r in parsed:
        if not r["dt"]:
            continue
        iso = r["dt"].isocalendar()
        label = f"{iso.year}-W{iso.week:02d}"
        weekly[label]["trades"] += 1
        weekly[label]["pnl"] += r["pnl"]
        if r["pnl"] > 0:
            weekly[label]["wins"] += 1
        elif r["pnl"] < 0:
            weekly[label]["losses"] += 1
    weekly_review = []
    for label, b in sorted(weekly.items())[-8:]:
        decided = b["wins"] + b["losses"]
        wr = (b["wins"] / decided * 100) if decided else 0
        mood = "Strong" if b["pnl"] > 0 and wr >= 50 else "Mixed" if b["pnl"] >= 0 else "Needs review"
        weekly_review.append({"week": label, "trades": b["trades"], "pnl": round(b["pnl"], 2), "win_rate": round(wr, 1), "label": mood})

    best_conditions = []
    worst_conditions = []
    for title, stats in [("Weekday", weekday_stats), ("Time window", time_stats), ("Broker", broker_stats)]:
        usable = [s for s in stats if s["trades"] >= 2]
        if usable:
            best = max(usable, key=lambda x: x["pnl"])
            worst = min(usable, key=lambda x: x["pnl"])
            if best["pnl"] > 0:
                best_conditions.append(f"{title}: {best['name']} (+${best['pnl']:.0f}, {best['trades']} trades)")
            if worst["pnl"] < 0:
                worst_conditions.append(f"{title}: {worst['name']} (${worst['pnl']:.0f}, {worst['trades']} trades)")

    avoid_list = []
    action_plan = []
    pattern_cards = []
    coach_questions = []

    overtrading_days = int(base_review.get("overtrading_days") or 0)
    max_loss_streak = int(base_review.get("max_losing_streak") or base_review.get("metrics", {}).get("max_loss_streak") or 0)
    net_pnl = sum(r["pnl"] for r in parsed)
    warnings_count = len(base_review.get("warnings", []))

    if overtrading_days:
        pattern_cards.append({"title": "Overtrading fingerprint", "value": str(overtrading_days), "detail": "Days crossed the 5+ trade threshold.", "tone": "warn"})
        avoid_list.append("Avoid adding trades after your daily trade limit is hit.")
    if max_loss_streak >= 3:
        pattern_cards.append({"title": "Loss-streak pressure", "value": str(max_loss_streak), "detail": "Longest run of losses. Use this as your cooldown trigger.", "tone": "neg"})
        avoid_list.append("Avoid taking another trade immediately after two losses.")
    if risk_coverage < 0.5:
        pattern_cards.append({"title": "Risk data gap", "value": f"{risk_coverage*100:.0f}%", "detail": "Trades with risk logged. More risk data = better coaching.", "tone": "warn"})
        action_plan.append({"priority": "High", "action": "Log risk on every new trade", "reason": "Without risk values, the coach cannot judge position sizing quality."})
    if tag_coverage < 0.6:
        action_plan.append({"priority": "High", "action": "Tag at least 80% of trades", "reason": "Unlabeled trades hide your real edge."})
    if mistake_stats:
        worst_m = mistake_stats[0]
        if worst_m["pnl"] < 0:
            avoid_list.append(f"Avoid trades showing this mistake: {worst_m['name']}.")
            coach_questions.append(f"What rule would have blocked your {worst_m['name']} trades?")

    worst_symbols = [s for s in sorted(symbol_scores, key=lambda x: x["pnl"]) if s["pnl"] < 0 and s["trades"] >= 2]
    if worst_symbols:
        avoid_list.append(f"Reduce or pause weakest symbol: {worst_symbols[0]['name']} (${worst_symbols[0]['pnl']:.0f}).")
        coach_questions.append(f"Why are you continuing to trade {worst_symbols[0]['name']} if it is dragging results?")
    best_symbols = [s for s in symbol_scores if s["pnl"] > 0 and s["trades"] >= 2]
    if best_symbols:
        action_plan.append({"priority": "Medium", "action": f"Study your best symbol: {best_symbols[0]['name']}", "reason": f"It produced about ${best_symbols[0]['pnl']:.0f} across {best_symbols[0]['trades']} trades."})

    if net_pnl > 0 and warnings_count <= 1 and risk_coverage >= 0.6:
        discipline_profile = "Structured winner"
    elif overtrading_days >= 2:
        discipline_profile = "Overtrading-prone"
    elif risk_coverage < 0.35:
        discipline_profile = "Under-measured risk"
    elif max_loss_streak >= 4:
        discipline_profile = "Cooldown needed"
    elif net_pnl > 0:
        discipline_profile = "Developing edge"
    else:
        discipline_profile = "Needs tighter filters"

    if best_conditions:
        pattern_cards.append({"title": "Best condition", "value": best_conditions[0].split(":", 1)[-1].strip(), "detail": best_conditions[0], "tone": "pos"})
    if worst_conditions:
        pattern_cards.append({"title": "Worst condition", "value": worst_conditions[0].split(":", 1)[-1].strip(), "detail": worst_conditions[0], "tone": "neg"})

    action_plan.append({"priority": "Medium", "action": "Write one lesson after each red day", "reason": "The fastest improvement usually comes from preventing repeat mistakes."})
    if not avoid_list:
        avoid_list.append("No clear avoid pattern yet. Keep tagging setups, mistakes, and risk.")
    if not coach_questions:
        coach_questions = [
            "Which setup would you still take if you could only take one trade per day?",
            "Which mistake is easiest to prevent with a checklist?",
            "What is your max daily loss before you must stop?",
        ]

    execution_summary = f"Profile: {discipline_profile}. AI confidence {ai_confidence}/100 based on {total} trades, {tag_coverage*100:.0f}% setup coverage, and {risk_coverage*100:.0f}% risk coverage."

    def unique_dicts(items):
        seen = set()
        out = []
        for item in items:
            key = tuple(sorted(item.items())) if isinstance(item, dict) else item
            if key not in seen:
                seen.add(key)
                out.append(item)
        return out

    return {
        "ai_confidence": ai_confidence,
        "discipline_profile": discipline_profile,
        "pattern_cards": pattern_cards[:6],
        "symbol_scores": symbol_scores[:8],
        "weekly_review": weekly_review,
        "avoid_list": list(dict.fromkeys(avoid_list))[:6],
        "action_plan": unique_dicts(action_plan)[:6],
        "best_conditions": best_conditions[:4],
        "worst_conditions": worst_conditions[:4],
        "coach_questions": coach_questions[:5],
        "execution_summary": execution_summary,
    }

def get_ai_review(where_sql="", params=None):
    params = params or []
    where_sql, params = scoped_where_sql(where_sql, params, alias="t")
    conn = get_db_connection()
    rows = conn.execute(f"""
        SELECT
            t.id,
            t.symbol,
            t.trade_datetime,
            t.realized_pl,
            t.r_multiple,
            t.broker,
            j.setup_tag,
            j.mistake_tag
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        {where_sql}
        ORDER BY t.trade_datetime ASC, t.id ASC
    """, params).fetchall()
    conn.close()
    return generate_ai_review_from_rows([dict(r) for r in rows])



def get_filtered_trade_dicts(where_sql="", params=None):
    params = params or []
    where_sql, params = scoped_where_sql(where_sql, params, alias="t")
    conn = get_db_connection()
    rows = conn.execute(f"""
        SELECT t.id, t.symbol, t.trade_datetime, t.realized_pl, t.quantity, t.side,
               t.trade_price, t.risk_amount, t.r_multiple, t.broker, j.setup_tag, j.mistake_tag, j.note AS journal_note
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        {where_sql}
        ORDER BY t.trade_datetime ASC, t.id ASC
    """, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def generate_what_if_analysis(rows):
    if not rows:
        return [{"title": "No data yet", "detail": "Import trades to unlock simulations.", "impact": 0.0}]

    scenarios = []
    by_day = defaultdict(list)
    for r in rows:
        day = (r.get("trade_datetime") or "")[:10]
        if day:
            by_day[day].append(r)

    actual_total = 0.0
    simulated_total = 0.0
    for day_rows in by_day.values():
        day_rows = sorted(day_rows, key=lambda r: r.get("trade_datetime") or "")
        losses_seen = 0
        for r in day_rows:
            pnl = float(r.get("realized_pl") or 0)
            actual_total += pnl
            if losses_seen < 2:
                simulated_total += pnl
            if pnl < 0:
                losses_seen += 1

    scenarios.append({
        "title": "Stop after 2 losses per day",
        "detail": "Simulates quitting for the day after your second losing trade.",
        "impact": round(simulated_total - actual_total, 2),
    })

    setup_map = defaultdict(float)
    setup_count = defaultdict(int)
    for r in rows:
        setup = (r.get("setup_tag") or "Unlabeled").strip() or "Unlabeled"
        pnl = float(r.get("realized_pl") or 0)
        setup_map[setup] += pnl
        setup_count[setup] += 1
    if setup_map:
        worst = min(setup_map, key=setup_map.get)
        scenarios.append({
            "title": f"Remove worst setup: {worst}",
            "detail": f"Removes {setup_count[worst]} trade(s) from this setup.",
            "impact": round(-setup_map[worst] if setup_map[worst] < 0 else 0.0, 2),
        })

    weekday_map = defaultdict(float)
    weekday_count = defaultdict(int)
    for r in rows:
        dt = (r.get("trade_datetime") or "")[:10]
        if not dt:
            continue
        try:
            weekday = datetime.fromisoformat(dt).strftime("%A")
            pnl = float(r.get("realized_pl") or 0)
            weekday_map[weekday] += pnl
            weekday_count[weekday] += 1
        except Exception:
            pass
    if weekday_map:
        worst_day = min(weekday_map, key=weekday_map.get)
        scenarios.append({
            "title": f"Avoid worst weekday: {worst_day}",
            "detail": f"Removes {weekday_count[worst_day]} trade(s) from this weekday.",
            "impact": round(-weekday_map[worst_day] if weekday_map[worst_day] < 0 else 0.0, 2),
        })

    mistake_map = defaultdict(float)
    mistake_count = defaultdict(int)
    for r in rows:
        mistake = (r.get("mistake_tag") or "No Mistake").strip() or "No Mistake"
        if mistake == "No Mistake":
            continue
        pnl = float(r.get("realized_pl") or 0)
        mistake_map[mistake] += pnl
        mistake_count[mistake] += 1
    if mistake_map:
        worst_mistake = min(mistake_map, key=mistake_map.get)
        scenarios.append({
            "title": f"Remove biggest mistake: {worst_mistake}",
            "detail": f"Removes {mistake_count[worst_mistake]} trade(s) tagged with this mistake.",
            "impact": round(-mistake_map[worst_mistake] if mistake_map[worst_mistake] < 0 else 0.0, 2),
        })

    return scenarios


# =========================================================
# UI shell
# =========================================================
def auth_page_shell(title, body):
    return f"""
    <!doctype html>
    <html>
    <head>
        <title>{title}</title>
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <link rel="preconnect" href="https://fonts.googleapis.com">
        <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
        <style>
            body {{
                margin: 0;
                font-family: Inter, sans-serif;
                background:
                    radial-gradient(circle at top left, rgba(91,140,255,0.18), transparent 28%),
                    radial-gradient(circle at top right, rgba(124,92,255,0.16), transparent 22%),
                    linear-gradient(180deg, #09101d 0%, #0b1020 100%);
                color: #e5ecf6;
                min-height: 100vh;
                display: grid;
                place-items: center;
            }}
            .card {{
                width: min(460px, 92vw);
                background: rgba(18, 26, 47, 0.92);
                border: 1px solid rgba(148, 163, 184, 0.14);
                border-radius: 22px;
                padding: 24px;
            }}
            h1 {{ margin-top: 0; }}
            input {{
                width: 100%;
                box-sizing: border-box;
                margin-bottom: 12px;
                padding: 12px 14px;
                border-radius: 14px;
                border: 1px solid rgba(148, 163, 184, 0.14);
                background: rgba(8, 13, 26, 0.72);
                color: white;
            }}
            button {{
                width: 100%;
                padding: 12px 16px;
                border: none;
                border-radius: 14px;
                color: white;
                font-weight: 700;
                cursor: pointer;
                background: linear-gradient(135deg, #5b8cff, #7c5cff);
            }}
            .muted {{ color: #94a3b8; font-size: 14px; }}
            a {{ color: #bfd4ff; text-decoration: none; }}
            .error {{
                background: rgba(239,68,68,0.12);
                border: 1px solid rgba(239,68,68,0.25);
                border-radius: 12px;
                padding: 10px 12px;
                margin-bottom: 12px;
                color: #fecaca;
            }}
        </style>
    </head>
    <body>
        <div class="card">{body}</div>
    </body>
    </html>
    """


def draggable_table_assets():
    return r'''
<style>
    th.column-draggable {
        cursor: grab;
        user-select: none;
        position: relative;
    }
    th.column-draggable:active { cursor: grabbing; }
    th.column-draggable::after {
        content: "↔";
        opacity: .32;
        font-size: 11px;
        margin-left: 8px;
        padding-right: 10px;
    }
    th.column-drag-over {
        outline: 2px dashed var(--primary, #5b8cff);
        outline-offset: -4px;
        background: rgba(91, 140, 255, .14) !important;
    }
    th.column-resizing,
    table.column-resizing * {
        cursor: col-resize !important;
        user-select: none !important;
    }
    .column-resize-handle {
        position: absolute;
        top: 0;
        right: 0;
        width: 12px;
        height: 100%;
        cursor: col-resize;
        z-index: 5;
        touch-action: none;
    }
    .column-resize-handle::after {
        content: "";
        position: absolute;
        top: 22%;
        bottom: 22%;
        right: 2px;
        width: 2px;
        border-radius: 99px;
        background: rgba(148,163,184,.34);
    }
    .column-resize-handle:hover::after {
        background: var(--primary, #5b8cff);
    }
    .table-wrap { position: relative; }
    .column-tools {
        display: flex;
        justify-content: flex-end;
        gap: 8px;
        padding: 8px 10px;
        border-bottom: 1px solid rgba(148,163,184,.10);
        background: rgba(7,12,23,.35);
    }
    .column-reset-btn {
        min-height: 30px !important;
        padding: 6px 10px !important;
        border-radius: 10px !important;
        font-size: 12px !important;
        font-weight: 800 !important;
        background: rgba(255,255,255,.05) !important;
        border: 1px solid var(--border, rgba(148,163,184,.16)) !important;
        color: var(--text, #e5ecf6) !important;
        width: auto !important;
    }
</style>
<script>
(function () {
    function hashText(text) {
        let hash = 0;
        for (let i = 0; i < text.length; i++) {
            hash = ((hash << 5) - hash) + text.charCodeAt(i);
            hash |= 0;
        }
        return Math.abs(hash).toString(36);
    }

    function textOf(el) {
        return (el.textContent || "").replace(/↔/g, "").replace(/\s+/g, " ").trim();
    }

    function getHeaderRow(table) {
        if (table.tHead && table.tHead.rows.length) return table.tHead.rows[0];
        return table.querySelector("tr");
    }

    function getColumnOrder(table) {
        const headerRow = getHeaderRow(table);
        if (!headerRow) return [];
        return Array.from(headerRow.children).map(cell => cell.dataset.colKey).filter(Boolean);
    }

    function tableBaseKey(table, tableIndex) {
        const headerRow = getHeaderRow(table);
        const headers = headerRow ? Array.from(headerRow.children).map(textOf).join("|") : "table";
        const tableHint = table.dataset.columnTableId || table.id || String(tableIndex);
        return location.pathname + ":" + tableHint + ":" + hashText(headers);
    }

    function orderStorageKey(table, tableIndex) {
        return "tradeJournal.columnOrder.v2:" + tableBaseKey(table, tableIndex);
    }

    function widthStorageKey(table, tableIndex) {
        return "tradeJournal.columnWidths.v2:" + tableBaseKey(table, tableIndex);
    }

    function assignColumnKeys(table) {
        const headerRow = getHeaderRow(table);
        if (!headerRow) return [];
        const headers = Array.from(headerRow.children);
        const keys = headers.map((cell, index) => {
            const existing = cell.dataset.colKey;
            if (existing) return existing;
            const label = textOf(cell) || "Column";
            const key = (label.toLowerCase().replace(/[^a-z0-9]+/g, "-") || "column") + "-" + index;
            cell.dataset.colKey = key;
            return key;
        });

        Array.from(table.rows).forEach(row => {
            const cells = Array.from(row.children);
            if (cells.length !== keys.length) return;
            cells.forEach((cell, index) => {
                cell.dataset.colKey = keys[index];
            });
        });
        return keys;
    }

    function applyOrder(table, order) {
        if (!order || !order.length) return;
        Array.from(table.rows).forEach(row => {
            const cells = Array.from(row.children);
            const byKey = new Map(cells.map(cell => [cell.dataset.colKey, cell]));
            const orderedCells = [];
            order.forEach(key => {
                if (byKey.has(key)) orderedCells.push(byKey.get(key));
            });
            cells.forEach(cell => {
                if (!orderedCells.includes(cell)) orderedCells.push(cell);
            });
            orderedCells.forEach(cell => row.appendChild(cell));
        });
    }

    function moveColumn(table, fromKey, toKey) {
        if (!fromKey || !toKey || fromKey === toKey) return;
        Array.from(table.rows).forEach(row => {
            const cells = Array.from(row.children);
            const from = cells.find(cell => cell.dataset.colKey === fromKey);
            const to = cells.find(cell => cell.dataset.colKey === toKey);
            if (!from || !to || from === to) return;
            const fromIndex = cells.indexOf(from);
            const toIndex = cells.indexOf(to);
            row.insertBefore(from, fromIndex < toIndex ? to.nextSibling : to);
        });
    }

    function numericPx(value) {
        const parsed = parseFloat(String(value || "").replace("px", ""));
        return Number.isFinite(parsed) ? parsed : 0;
    }

    function getVisibleHeaders(table) {
        const headerRow = getHeaderRow(table);
        if (!headerRow) return [];
        return Array.from(headerRow.children).filter(cell => cell.tagName.toLowerCase() === "th");
    }

    function syncTablePixelWidth(table) {
        if (!table) return;
        table.classList.add("column-enhanced");

        const headers = getVisibleHeaders(table);
        if (!headers.length) return;

        let total = 0;
        headers.forEach(th => {
            const styleWidth = numericPx(th.style.width) || numericPx(th.style.minWidth);
            const renderedWidth = th.getBoundingClientRect ? th.getBoundingClientRect().width : 0;
            const width = Math.max(64, Math.round(styleWidth || renderedWidth || 120));
            total += width;
        });

        const wrap = table.closest(".table-wrap");
        const wrapWidth = wrap ? Math.max(0, wrap.clientWidth - 2) : 0;
        const finalWidth = Math.max(total, wrapWidth, 720);

        table.style.width = finalWidth + "px";
        table.style.minWidth = finalWidth + "px";
    }

    function setColumnWidth(table, colKey, widthPx) {
        const width = Math.max(64, Math.round(widthPx));
        Array.from(table.rows).forEach(row => {
            Array.from(row.children).forEach(cell => {
                if (cell.dataset.colKey === colKey) {
                    cell.style.width = width + "px";
                    cell.style.minWidth = width + "px";
                    cell.style.maxWidth = width + "px";
                    cell.style.overflow = "hidden";
                    cell.style.textOverflow = "ellipsis";
                }
            });
        });
        syncTablePixelWidth(table);
    }

    function getSavedWidths(widthKey) {
        try { return JSON.parse(localStorage.getItem(widthKey) || "{}"); }
        catch (e) { return {}; }
    }

    function saveColumnWidth(widthKey, colKey, widthPx) {
        const widths = getSavedWidths(widthKey);
        widths[colKey] = Math.max(48, Math.round(widthPx));
        localStorage.setItem(widthKey, JSON.stringify(widths));
    }

    function applySavedWidths(table, widthKey) {
        const widths = getSavedWidths(widthKey);
        Object.keys(widths).forEach(colKey => setColumnWidth(table, colKey, widths[colKey]));
        syncTablePixelWidth(table);
    }

    function addResetButton(table, storageKey, widthKey) {
        const wrap = table.closest(".table-wrap");
        if (!wrap || wrap.querySelector(":scope > .column-tools")) return;
        const tools = document.createElement("div");
        tools.className = "column-tools";
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "column-reset-btn";
        btn.textContent = "Reset columns";
        btn.title = "Reset column order and column widths";
        btn.addEventListener("click", () => {
            localStorage.removeItem(storageKey);
            localStorage.removeItem(widthKey);
            location.reload();
        });
        tools.appendChild(btn);
        wrap.insertBefore(tools, wrap.firstChild);
    }

    function addResizeHandles(table, widthKey) {
        const headerRow = getHeaderRow(table);
        if (!headerRow) return;
        const headers = Array.from(headerRow.children).filter(cell => cell.tagName.toLowerCase() === "th");

        headers.forEach(th => {
            if (th.querySelector(":scope > .column-resize-handle")) return;
            th.style.position = th.style.position || "relative";
            const handle = document.createElement("span");
            handle.className = "column-resize-handle";
            handle.title = "Drag to resize column";

            handle.addEventListener("mousedown", event => {
                event.preventDefault();
                event.stopPropagation();
                table.classList.add("column-resizing");
                th.classList.add("column-resizing");
                const colKey = th.dataset.colKey;
                const startX = event.clientX;
                const startWidth = th.getBoundingClientRect().width;

                function onMove(moveEvent) {
                    const nextWidth = startWidth + (moveEvent.clientX - startX);
                    setColumnWidth(table, colKey, nextWidth);
                }

                function onUp(upEvent) {
                    document.removeEventListener("mousemove", onMove);
                    document.removeEventListener("mouseup", onUp);
                    table.classList.remove("column-resizing");
                    th.classList.remove("column-resizing");
                    const finalWidth = Math.max(48, startWidth + (upEvent.clientX - startX));
                    setColumnWidth(table, colKey, finalWidth);
                    saveColumnWidth(widthKey, colKey, finalWidth);
                }

                document.addEventListener("mousemove", onMove);
                document.addEventListener("mouseup", onUp);
            });

            handle.addEventListener("dragstart", event => {
                event.preventDefault();
                event.stopPropagation();
            });

            th.appendChild(handle);
        });
    }

    function prepareTable(table, tableIndex) {
        if (!table || table.dataset.noColumnReorder === "1") return;
        const headerRow = getHeaderRow(table);
        if (!headerRow) return;
        const headers = Array.from(headerRow.children).filter(cell => cell.tagName.toLowerCase() === "th");
        if (headers.length < 2) return;

        table.classList.add("column-enhanced");
        headers.forEach(th => {
            if (!th.style.width) {
                const w = Math.max(96, Math.round(th.getBoundingClientRect ? th.getBoundingClientRect().width : 120));
                th.style.width = w + "px";
                th.style.minWidth = w + "px";
                th.style.maxWidth = w + "px";
            }
        });

        assignColumnKeys(table);
        const orderKey = orderStorageKey(table, tableIndex);
        const widthKey = widthStorageKey(table, tableIndex);

        if (table.dataset.columnOrderApplied !== "1") {
            let savedOrder = [];
            try { savedOrder = JSON.parse(localStorage.getItem(orderKey) || "[]"); } catch (e) { savedOrder = []; }
            if (savedOrder.length) { applyOrder(table, savedOrder); syncTablePixelWidth(table); }
            table.dataset.columnOrderApplied = "1";
        }

        applySavedWidths(table, widthKey);

        const freshHeaders = Array.from(getHeaderRow(table).children).filter(cell => cell.tagName.toLowerCase() === "th");
        freshHeaders.forEach(th => {
            if (th.dataset.columnDragReady !== "1") {
                th.draggable = true;
                th.classList.add("column-draggable");
                th.title = "Drag to reorder columns. Drag the right edge to resize.";

                th.addEventListener("dragstart", event => {
                    if (event.target && event.target.classList && event.target.classList.contains("column-resize-handle")) {
                        event.preventDefault();
                        return;
                    }
                    event.dataTransfer.effectAllowed = "move";
                    event.dataTransfer.setData("text/plain", th.dataset.colKey || "");
                    table.dataset.dragColumnKey = th.dataset.colKey || "";
                });

                th.addEventListener("dragover", event => {
                    event.preventDefault();
                    th.classList.add("column-drag-over");
                });

                th.addEventListener("dragleave", () => th.classList.remove("column-drag-over"));

                th.addEventListener("drop", event => {
                    event.preventDefault();
                    th.classList.remove("column-drag-over");
                    const fromKey = event.dataTransfer.getData("text/plain") || table.dataset.dragColumnKey;
                    const toKey = th.dataset.colKey;
                    moveColumn(table, fromKey, toKey);
                    syncTablePixelWidth(table);
                    localStorage.setItem(orderKey, JSON.stringify(getColumnOrder(table)));
                    applySavedWidths(table, widthKey);
                });

                th.dataset.columnDragReady = "1";
            }
        });

        addResizeHandles(table, widthKey);
        addResetButton(table, orderKey, widthKey);
        syncTablePixelWidth(table);
        table.dataset.columnReorderReady = "1";
    }

    function initColumnControls() {
        document.querySelectorAll("table").forEach((table, index) => prepareTable(table, index));
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", initColumnControls);
    } else {
        initColumnControls();
    }

    window.addEventListener("resize", () => document.querySelectorAll("table.column-enhanced").forEach(syncTablePixelWidth));

    const observer = new MutationObserver(() => initColumnControls());
    observer.observe(document.documentElement, { childList: true, subtree: true });
})();
</script>

'''


def page_shell(title, body, extra_head=""):
    portfolios = get_user_portfolios(current_user_id()) if require_login() else []
    active_portfolio_id = current_portfolio_id()
    active_portfolio_name = current_portfolio_name() if require_login() else "No Portfolio"

    portfolio_links = "".join([
        f'<option value="{p["id"]}" {"selected" if p["id"] == active_portfolio_id else ""}>{p["name"]}</option>'
        for p in portfolios
    ])

    return f"""
    <!doctype html>
    <html lang="en">
    <head>
        <title>{title}</title>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <link rel="preconnect" href="https://fonts.googleapis.com">
        <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

        <style>
            :root {{
                --card: rgba(18, 26, 47, 0.88);
                --card2: rgba(24, 35, 63, 0.88);
                --border: rgba(148, 163, 184, 0.14);
                --text: #e5ecf6;
                --muted: #94a3b8;
                --primary: #5b8cff;
                --secondary: #7c5cff;
                --green: #22c55e;
                --red: #ef4444;
                --yellow: #f59e0b;
            }}
            * {{ box-sizing: border-box; }}
            html, body {{
                margin: 0;
                font-family: Inter, sans-serif;
                color: var(--text);
                background:
                    radial-gradient(circle at top left, rgba(91,140,255,0.18), transparent 28%),
                    radial-gradient(circle at top right, rgba(124,92,255,0.16), transparent 22%),
                    linear-gradient(180deg, #09101d 0%, #0b1020 100%);
            }}
            a {{ color: inherit; text-decoration: none; }}
            .app-shell {{
                display: grid;
                grid-template-columns: 260px 1fr;
                min-height: 100vh;
            }}
            .sidebar {{
                padding: 22px 18px;
                border-right: 1px solid var(--border);
                background: linear-gradient(180deg, rgba(10,15,30,0.94), rgba(12,18,34,0.98));
            }}
            .brand {{
                display: flex;
                align-items: center;
                gap: 12px;
                margin-bottom: 18px;
            }}
            .brand-badge {{
                width: 42px;
                height: 42px;
                border-radius: 14px;
                display: grid;
                place-items: center;
                font-weight: 800;
                background: linear-gradient(135deg, var(--primary), var(--secondary));
            }}
            .brand-sub {{ color: var(--muted); font-size: 12px; }}
            .nav-list {{
                display: flex;
                flex-direction: column;
                gap: 6px;
                margin-bottom: 18px;
            }}
            .nav-item {{
                display: flex;
                align-items: center;
                gap: 10px;
                padding: 12px 14px;
                border-radius: 14px;
                border: 1px solid transparent;
                color: #c7d2e3;
            }}
            .nav-item.active {{
                background: linear-gradient(135deg, rgba(91,140,255,0.16), rgba(124,92,255,0.14));
                border-color: rgba(91,140,255,0.22);
                color: white;
            }}
            .nav-dot {{
                width: 10px;
                height: 10px;
                border-radius: 999px;
                background: linear-gradient(135deg, var(--primary), var(--secondary));
            }}
            .sidebar-card {{
                margin-top: 20px;
                padding: 16px;
                border-radius: 18px;
                background: linear-gradient(180deg, rgba(91,140,255,0.10), rgba(124,92,255,0.08));
                border: 1px solid rgba(91,140,255,0.16);
            }}
            .main {{
                padding: 22px;
                min-width: 0;
            }}
            .topbar {{
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 16px;
                margin-bottom: 18px;
                flex-wrap: wrap;
            }}
            .page-title {{
                font-size: 28px;
                font-weight: 800;
                margin: 0;
            }}
            .page-subtitle {{
                color: var(--muted);
                font-size: 14px;
                margin-top: 6px;
            }}
            .pill {{
                display: inline-flex;
                align-items: center;
                gap: 8px;
                padding: 10px 14px;
                border-radius: 999px;
                background: rgba(255,255,255,0.04);
                border: 1px solid var(--border);
                color: var(--muted);
                font-size: 13px;
            }}
            .glass-card {{
                background: linear-gradient(180deg, var(--card), rgba(15, 21, 40, 0.92));
                border: 1px solid var(--border);
                border-radius: 20px;
                padding: 18px;
                margin-bottom: 18px;
            }}
            .soft {{
                background: linear-gradient(180deg, var(--card2), rgba(17, 24, 39, 0.96));
            }}
            .section-head {{
                display: flex;
                justify-content: space-between;
                gap: 12px;
                align-items: center;
                margin-bottom: 14px;
            }}
            .section-title {{
                margin: 0;
                font-size: 18px;
                font-weight: 700;
            }}
            .section-note {{
                font-size: 13px;
                color: var(--muted);
            }}
            .filters {{
                display: flex;
                flex-wrap: wrap;
                gap: 10px;
            }}
            input, select, textarea {{
                border: 1px solid var(--border);
                background: rgba(8, 13, 26, 0.72);
                color: var(--text);
                border-radius: 14px;
                padding: 12px 14px;
                font-size: 14px;
                outline: none;
            }}
            textarea {{
                width: 100%;
                min-height: 140px;
                resize: vertical;
            }}
            .btn {{
                display: inline-flex;
                align-items: center;
                justify-content: center;
                border: 1px solid transparent;
                border-radius: 14px;
                padding: 12px 16px;
                min-height: 46px;
                font-size: 14px;
                font-weight: 700;
                cursor: pointer;
                color: white;
                background: linear-gradient(135deg, var(--primary), var(--secondary));
            }}
            .btn.secondary {{
                background: rgba(255,255,255,0.04);
                border-color: var(--border);
                color: var(--text);
            }}
            .btn.danger {{
                background: linear-gradient(135deg, #ef4444, #dc2626);
            }}
            .grid-kpi {{
                display: grid;
                grid-template-columns: repeat(5, minmax(0, 1fr));
                gap: 14px;
            }}
            .kpi {{
                padding: 18px;
                border-radius: 20px;
                background: linear-gradient(180deg, rgba(255,255,255,0.05), rgba(255,255,255,0.02));
                border: 1px solid var(--border);
                min-height: 120px;
            }}
            .kpi-label {{
                font-size: 12px;
                color: var(--muted);
                margin-bottom: 10px;
                text-transform: uppercase;
                letter-spacing: .10em;
            }}
            .kpi-value {{
                font-size: 28px;
                font-weight: 800;
                line-height: 1.05;
            }}
            .kpi-foot {{
                margin-top: 10px;
                font-size: 12px;
                color: var(--muted);
            }}
            .split-2 {{
                display: grid;
                grid-template-columns: 1.35fr 1fr;
                gap: 18px;
            }}
            .split-3 {{
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                gap: 18px;
            }}
            .chart-wrap {{
                height: 320px;
                position: relative;
            }}
            .table-wrap {{
                overflow: auto;
                border-radius: 18px;
                border: 1px solid var(--border);
                background: rgba(7, 12, 23, 0.55);
            }}
            .compact-table-wrap {{
                max-height: 430px;
                overflow-y: auto;
            }}
            .compact-table-wrap table {{
                min-width: 0;
            }}
            .compact-table-wrap thead th {{
                position: sticky;
                top: 0;
                z-index: 2;
            }}
            .compact-table-wrap tbody td {{
                padding: 10px 12px;
                font-size: 13px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                min-width: 860px;
            }}
            thead th {{
                background: rgba(14, 21, 38, 0.96);
                color: #9fb0c9;
                font-size: 12px;
                text-transform: uppercase;
                letter-spacing: 0.08em;
                font-weight: 700;
                padding: 14px;
                text-align: left;
                border-bottom: 1px solid var(--border);
            }}
            tbody td {{
                padding: 14px;
                border-bottom: 1px solid rgba(148,163,184,0.08);
                font-size: 14px;
                color: #d6e0ef;
                white-space: nowrap;
            }}
            .tag {{
                display: inline-flex;
                align-items: center;
                gap: 8px;
                padding: 7px 10px;
                border-radius: 999px;
                font-size: 12px;
                font-weight: 700;
                border: 1px solid var(--border);
                background: rgba(255,255,255,0.04);
            }}
            .tag.buy {{
                color: #9af0b3;
                background: rgba(34,197,94,0.10);
                border-color: rgba(34,197,94,0.18);
            }}
            .tag.sell {{
                color: #ffb4b4;
                background: rgba(239,68,68,0.10);
                border-color: rgba(239,68,68,0.18);
            }}
            .tag.broker {{
                color: #bfd4ff;
                background: rgba(91,140,255,0.10);
                border-color: rgba(91,140,255,0.18);
            }}
            .tag.asset {{
                color: #d9c7ff;
                background: rgba(124,92,255,0.10);
                border-color: rgba(124,92,255,0.18);
            }}
            .mono {{
                font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
            }}
            .empty {{
                padding: 26px;
                text-align: center;
                color: var(--muted);
            }}
            .upload-grid {{
                display: grid;
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 14px;
            }}
            .thumb {{
                max-width: 220px;
                border-radius: 12px;
                border: 1px solid var(--border);
            }}
            .pos {{ color: var(--green); }}
            .neg {{ color: var(--red); }}
            .warn {{ color: var(--yellow); }}

            @media (max-width: 1200px) {{
                .grid-kpi {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
                .split-2, .split-3 {{ grid-template-columns: 1fr; }}
            }}
        
    .modal-backdrop {{
        display: none;
        position: fixed;
        inset: 0;
        background: rgba(2, 6, 23, 0.72);
        z-index: 999;
        padding: 24px;
        overflow: auto;
    }}
    .modal-panel {{
        width: min(960px, 96vw);
        margin: 5vh auto;
        background: linear-gradient(180deg, rgba(18, 26, 47, 0.98), rgba(15, 21, 40, 0.98));
        border: 1px solid var(--border);
        border-radius: 22px;
        padding: 18px;
        box-shadow: 0 24px 80px rgba(0,0,0,0.45);
    }}
    .modal-head {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        margin-bottom: 14px;
    }}

    @media (max-width: 980px) {{
                .app-shell {{ grid-template-columns: 1fr; }}
            }}
        </style>
        {draggable_table_assets()}
        {extra_head}
    </head>
    <body>
        <div class="app-shell">
            <aside class="sidebar">
                <div class="brand">
                    <div class="brand-badge">TJ</div>
                    <div>
                        <div style="font-weight:800;">Trade Journal</div>
                        <div class="brand-sub">Multi-portfolio workspace</div>
                    </div>
                </div>

                <div class="nav-list">
                    <a class="nav-item {'active' if title == 'Dashboard' else ''}" href="/dashboard"><span class="nav-dot"></span>Dashboard</a>
                    <a class="nav-item {'active' if title == 'Trades' else ''}" href="/trades"><span class="nav-dot"></span>Trades</a>
                    <a class="nav-item {'active' if title == 'Option Strategies' else ''}" href="/options"><span class="nav-dot"></span>Strategies</a>
                    <a class="nav-item {'active' if title == 'Imports' else ''}" href="/imports"><span class="nav-dot"></span>Imports</a>
                    <a class="nav-item {'active' if title == 'Gallery' else ''}" href="/gallery"><span class="nav-dot"></span>Gallery</a>
                    <a class="nav-item {'active' if title == 'IBKR Settings' else ''}" href="/settings/ibkr"><span class="nav-dot"></span>IBKR Settings</a>
                    <a class="nav-item {'active' if title == 'NinjaTrader Settings' else ''}" href="/settings/ninjatrader"><span class="nav-dot"></span>NinjaTrader Settings</a>
                    <a class="nav-item {'active' if title == 'Wealthsimple Settings' else ''}" href="/settings/wealthsimple"><span class="nav-dot"></span>Wealthsimple Settings</a>
                    <a class="nav-item" href="/export/trades"><span class="nav-dot"></span>Export CSV</a>
                    <a class="nav-item" href="/export/trades.xlsx"><span class="nav-dot"></span>Export XLSX</a>
                </div>

                <div class="sidebar-card">
                    <div style="font-weight:700; margin-bottom:4px;">Portfolio</div>
                    <div class="section-note" style="margin-bottom:10px;">Current: <strong>{active_portfolio_name}</strong></div>
                    <form method="post" action="/portfolio/switch">
                        <select name="portfolio_id" style="width:100%; margin-bottom:10px;">
                            {portfolio_links}
                        </select>
                        <button class="btn secondary" type="submit" style="width:100%;">Switch</button>
                    </form>
                    <form method="post" action="/portfolio/create" style="margin-top:10px;">
                        <input name="portfolio_name" placeholder="New portfolio name" style="width:100%; margin-bottom:10px;">
                        <button class="btn" type="submit" style="width:100%;">Create Portfolio</button>
                    </form>
                </div>
                <div class="sidebar-card">
                    <div style="font-weight:700; margin-bottom:8px;">Session</div>
                    <a class="btn secondary" href="/logout" style="width:100%;">Log Out</a>
                </div>
            </aside>

            <main class="main">
                <div class="topbar">
                    <div>
                        <h1 class="page-title">{title}</h1>
                        <div class="page-subtitle">Trading journal, analytics, and portfolio review.</div>
                    </div>
                    <div class="pill">Portfolio: {active_portfolio_name}</div>
                </div>
                {body}
            </main>
        </div>
    </body>
    </html>
    """


# =========================================================
# =========================================================
# Public landing page
# =========================================================
def landing_page_shell():
    return """
    <!doctype html>
    <html lang="en">
    <head>
        <title>Trade Journal</title>
        <meta charset="utf-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <link rel="preconnect" href="https://fonts.googleapis.com">
        <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
        <style>
            :root { --card: rgba(18, 26, 47, 0.86); --border: rgba(148,163,184,0.16); --text: #e5ecf6; --muted: #94a3b8; --primary: #5b8cff; --secondary: #7c5cff; --green: #22c55e; }
            * { box-sizing: border-box; }
            body { margin: 0; font-family: Inter, sans-serif; color: var(--text); background: radial-gradient(circle at top left, rgba(91,140,255,0.22), transparent 28%), radial-gradient(circle at top right, rgba(124,92,255,0.22), transparent 24%), linear-gradient(180deg, #09101d 0%, #0b1020 100%); min-height: 100vh; }
            a { color: inherit; text-decoration: none; }
            .container { width: min(1120px, calc(100% - 36px)); margin: 0 auto; }
            .nav { display: flex; align-items: center; justify-content: space-between; padding: 22px 0; }
            .brand { display: flex; align-items: center; gap: 12px; font-weight: 900; }
            .brand-badge { width: 42px; height: 42px; display: grid; place-items: center; border-radius: 14px; background: linear-gradient(135deg, var(--primary), var(--secondary)); box-shadow: 0 16px 60px rgba(91,140,255,0.28); }
            .nav-actions { display: flex; align-items: center; gap: 10px; }
            .btn { display: inline-flex; align-items: center; justify-content: center; min-height: 46px; padding: 12px 16px; border-radius: 14px; font-weight: 800; background: linear-gradient(135deg, var(--primary), var(--secondary)); color: white; border: 1px solid transparent; }
            .btn.secondary { background: rgba(255,255,255,0.05); border-color: var(--border); }
            .hero { display: grid; grid-template-columns: 1.1fr 0.9fr; gap: 34px; align-items: center; padding: 80px 0 64px; }
            .eyebrow { display: inline-flex; padding: 8px 12px; border-radius: 999px; border: 1px solid rgba(91,140,255,0.26); background: rgba(91,140,255,0.10); color: #bfd4ff; font-size: 13px; font-weight: 800; margin-bottom: 18px; }
            h1 { font-size: clamp(42px, 7vw, 76px); line-height: 0.95; letter-spacing: -0.06em; margin: 0 0 20px; }
            .subhead { color: var(--muted); font-size: 19px; line-height: 1.65; max-width: 650px; margin: 0 0 26px; }
            .hero-actions { display: flex; gap: 12px; flex-wrap: wrap; }
            .mock-card { border-radius: 28px; border: 1px solid var(--border); background: linear-gradient(180deg, var(--card), rgba(15,21,40,0.96)); padding: 20px; box-shadow: 0 30px 100px rgba(0,0,0,0.34); }
            .metric-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; margin-bottom: 14px; }
            .metric { border: 1px solid var(--border); border-radius: 18px; padding: 14px; background: rgba(255,255,255,0.04); }
            .metric-label { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: 0.09em; }
            .metric-value { font-size: 27px; font-weight: 900; margin-top: 8px; }
            .pos { color: var(--green); }
            .chart-placeholder { height: 210px; border-radius: 20px; background: linear-gradient(180deg, rgba(34,197,94,0.12), rgba(34,197,94,0)), repeating-linear-gradient(90deg, rgba(148,163,184,0.08) 0 1px, transparent 1px 42px), repeating-linear-gradient(0deg, rgba(148,163,184,0.08) 0 1px, transparent 1px 42px); border: 1px solid var(--border); position: relative; overflow: hidden; }
            .chart-line { position: absolute; inset: 46px 20px 36px 20px; border-bottom: 5px solid rgba(34,197,94,0.72); border-radius: 55% 35% 8% 8%; transform: rotate(-7deg); }
            .features { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; padding-bottom: 70px; }
            .feature { border-radius: 22px; border: 1px solid var(--border); background: rgba(18,26,47,0.72); padding: 20px; }
            .feature h3 { margin: 0 0 10px; }
            .feature p { color: var(--muted); line-height: 1.55; margin: 0; }
            @media (max-width: 900px) { .hero, .features { grid-template-columns: 1fr; } .hero { padding-top: 44px; } .nav { align-items: flex-start; gap: 12px; } .nav-actions { flex-wrap: wrap; justify-content: flex-end; } }
        </style>
    </head>
    <body>
        <div class="container">
            <nav class="nav">
                <a class="brand" href="/"><span class="brand-badge">TJ</span><span>Trade Journal</span></a>
                <div class="nav-actions"><a class="btn secondary" href="/login">Log in</a><a class="btn" href="/register">Get started free</a></div>
            </nav>
            <section class="hero">
                <div>
                    <div class="eyebrow">Trading analytics without spreadsheet chaos</div>
                    <h1>Track better. Review faster. Trade smarter.</h1>
                    <p class="subhead">Import your trades, journal your setups, review your mistakes, and spot patterns with dashboards built for serious traders.</p>
                    <div class="hero-actions"><a class="btn" href="/register">Start free</a><a class="btn secondary" href="/login">I already have an account</a></div>
                </div>
                <div class="mock-card">
                    <div class="metric-grid">
                        <div class="metric"><div class="metric-label">Realized P&amp;L</div><div class="metric-value pos">+$8,420</div></div>
                        <div class="metric"><div class="metric-label">Win Rate</div><div class="metric-value">58.4%</div></div>
                        <div class="metric"><div class="metric-label">Profit Factor</div><div class="metric-value">1.92</div></div>
                        <div class="metric"><div class="metric-label">AI Insights</div><div class="metric-value">12</div></div>
                    </div>
                    <div class="chart-placeholder"><div class="chart-line"></div></div>
                </div>
            </section>
            <section class="features">
                <div class="feature"><h3>📊 Performance dashboard</h3><p>Track P&amp;L, drawdown, win rate, expectancy, broker performance, and trade count.</p></div>
                <div class="feature"><h3>🧠 Built-in trade coach</h3><p>Find overtrading, loss streaks, weak setups, expensive mistakes, and rule ideas.</p></div>
                <div class="feature"><h3>📁 Broker imports</h3><p>Upload IBKR, Wealthsimple, IBKR summary, and NinjaTrader-style performance CSVs.</p></div>
            </section>
        </div>
    </body>
    </html>
    """


@app.route("/")
def home():
    if require_login():
        return redirect(url_for("dashboard"))
    return landing_page_shell()



# Auth routes
# =========================================================
@app.route("/register", methods=["GET", "POST"])
def register():
    if require_login():
        return redirect(url_for("dashboard"))

    error = ""
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "").strip()

        if not email or not password:
            error = "Email and password are required."
        else:
            try:
                user_id = create_user(email, password)
                session["user_id"] = user_id
                return redirect(url_for("dashboard"))
            except sqlite3.IntegrityError:
                error = "That email is already registered."

    body = f"""
<h1>Create account</h1>
<p class="muted">Start your journal workspace.</p>
{'<div class="error">' + error + '</div>' if error else ''}
<form method="post">
    <input name="email" type="email" placeholder="Email" required>
    <input name="password" type="password" placeholder="Password" required>
    <button type="submit">Create account</button>
</form>
<p class="muted" style="margin-top:12px;">Already have an account? <a href="/login">Log in</a></p>
"""
    return auth_page_shell("Register", body)


@app.route("/login", methods=["GET", "POST"])
def login():
    if require_login():
        return redirect(url_for("dashboard"))

    error = ""
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "").strip()
        user_id = authenticate_user(email, password)
        if user_id:
            session["user_id"] = user_id
            session.pop("portfolio_id", None)
            return redirect(url_for("dashboard"))
        error = "Invalid email or password."

    body = f"""
<h1>Log in</h1>
<p class="muted">Access your journal workspace.</p>
{'<div class="error">' + error + '</div>' if error else ''}
<form method="post">
    <input name="email" type="email" placeholder="Email" required>
    <input name="password" type="password" placeholder="Password" required>
    <button type="submit">Log in</button>
</form>
<p class="muted" style="margin-top:12px;">No account yet? <a href="/register">Create one</a></p>
"""
    return auth_page_shell("Login", body)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))



# =========================================================
# IBKR Flex settings
# =========================================================
def get_ibkr_settings(user_id=None):
    user_id = user_id or current_user_id()
    if not user_id:
        return None
    conn = get_db_connection()
    row = conn.execute("""
        SELECT user_id, flex_token, query_id, account_id, report_format,
               auto_import_enabled, auto_import_hour, last_import_at,
               last_auto_import_date, last_auto_import_status, updated_at
        FROM ibkr_settings
        WHERE user_id = ?
    """, (user_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def mask_secret(value, visible=4):
    value = str(value or "")
    if not value:
        return ""
    if len(value) <= visible:
        return "•" * len(value)
    return "•" * max(0, len(value) - visible) + value[-visible:]


def save_ibkr_settings(user_id, flex_token, query_id, account_id="", report_format="xml", auto_import_enabled=False, auto_import_hour=6):
    flex_token = str(flex_token or "").strip()
    query_id = str(query_id or "").strip()
    account_id = str(account_id or "").strip()
    report_format = (str(report_format or "xml").strip().lower() or "xml")
    if report_format not in {"xml", "csv"}:
        report_format = "xml"
    try:
        auto_import_hour = int(auto_import_hour)
    except (TypeError, ValueError):
        auto_import_hour = 6
    auto_import_hour = max(0, min(23, auto_import_hour))

    conn = get_db_connection()
    conn.execute("""
        INSERT INTO ibkr_settings (
            user_id, flex_token, query_id, account_id, report_format,
            auto_import_enabled, auto_import_hour, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(user_id) DO UPDATE SET
            flex_token = excluded.flex_token,
            query_id = excluded.query_id,
            account_id = excluded.account_id,
            report_format = excluded.report_format,
            auto_import_enabled = excluded.auto_import_enabled,
            auto_import_hour = excluded.auto_import_hour,
            updated_at = CURRENT_TIMESTAMP
    """, (
        user_id, flex_token, query_id, account_id, report_format,
        1 if auto_import_enabled else 0,
        auto_import_hour,
    ))
    conn.commit()
    conn.close()


def clear_ibkr_settings(user_id):
    conn = get_db_connection()
    conn.execute("DELETE FROM ibkr_settings WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()

# =========================================================
# Portfolio routes
# =========================================================
@app.route("/portfolio/create", methods=["POST"])
def portfolio_create():
    if not require_login():
        return redirect(url_for("login"))

    name = request.form.get("portfolio_name", "").strip()
    if name:
        conn = get_db_connection()
        try:
            conn.execute(
                "INSERT INTO portfolios (user_id, name) VALUES (?, ?)",
                (current_user_id(), name)
            )
            conn.commit()
        except sqlite3.IntegrityError:
            pass
        finally:
            conn.close()
    return redirect(request.referrer or url_for("dashboard"))


@app.route("/portfolio/switch", methods=["POST"])
def portfolio_switch():
    if not require_login():
        return redirect(url_for("login"))

    try:
        portfolio_id = int(request.form.get("portfolio_id", "0"))
    except ValueError:
        portfolio_id = 0

    if portfolio_id and ensure_portfolio_access(current_user_id(), portfolio_id):
        session["portfolio_id"] = portfolio_id

    return redirect(request.referrer or url_for("dashboard"))



def generate_tomorrow_plan(rows):
    """Adaptive, time-aware next-session trading plan.

    Local/rule-based AI from your own journal history. It weights recent trades
    more heavily, looks for time-of-day edge, detects declining recent results,
    and keeps the existing dashboard fields such as mode, trade_rule, watchlist,
    and avoid_reasons.
    """
    if not rows:
        return {
            "best_setup": "No data yet",
            "avoid_setup": "No data yet",
            "best_time": "No data yet",
            "confidence": 0,
            "trend_warning": None,
            "reason": "Import trades and tag setups to unlock a smart trading plan.",
            "mode": "Collect data",
            "trade_rule": "Log setup, mistake, and risk for every trade.",
            "watchlist": [],
            "avoid_reasons": [],
        }

    parsed = []
    for r in rows:
        pnl = float(r.get("realized_pl") or r.get("pnl") or 0)
        setup = (r.get("setup_tag") or r.get("setup") or "Unlabeled").strip() or "Unlabeled"
        symbol = (r.get("symbol") or "Unknown").strip() or "Unknown"
        dt_raw = r.get("trade_datetime") or ""
        dt = None
        try:
            dt = datetime.fromisoformat(str(dt_raw)[:19]) if dt_raw else None
        except Exception:
            dt = None
        parsed.append({"pnl": pnl, "setup": setup, "symbol": symbol, "dt": dt})

    parsed.sort(key=lambda x: x["dt"] or datetime.min)
    total_trades = len(parsed)

    setup_stats = defaultdict(lambda: {
        "trades": 0,
        "wins": 0,
        "losses": 0,
        "pnl": 0.0,
        "weighted_pnl": 0.0,
        "recent_weighted_pnl": 0.0,
        "pnl_values": [],
        "symbols": defaultdict(float),
    })
    time_stats = defaultdict(lambda: {"weighted_pnl": 0.0, "trades": 0})

    for idx, r in enumerate(parsed):
        # Stronger recency weighting: latest trades matter up to 3x more.
        recency_weight = 1 + (idx / max(total_trades, 1)) * 2
        s = setup_stats[r["setup"]]
        s["trades"] += 1
        s["pnl"] += r["pnl"]
        s["weighted_pnl"] += r["pnl"] * recency_weight
        s["pnl_values"].append(r["pnl"])
        s["symbols"][r["symbol"]] += r["pnl"] * recency_weight
        if idx >= max(0, total_trades - 10):
            s["recent_weighted_pnl"] += r["pnl"] * recency_weight
        if r["pnl"] > 0:
            s["wins"] += 1
        elif r["pnl"] < 0:
            s["losses"] += 1

        if r["dt"]:
            hour = r["dt"].hour
            if hour < 11:
                bucket = "Morning"
            elif hour < 14:
                bucket = "Midday"
            else:
                bucket = "Afternoon"
            time_stats[bucket]["weighted_pnl"] += r["pnl"] * recency_weight
            time_stats[bucket]["trades"] += 1

    scored = []
    for setup, s in setup_stats.items():
        decided = s["wins"] + s["losses"]
        win_rate = s["wins"] / decided if decided else 0
        avg_pnl = s["pnl"] / s["trades"] if s["trades"] else 0
        sample_factor = min(1.0, s["trades"] / 8)
        losses = [abs(x) for x in s["pnl_values"] if x < 0]
        wins = [x for x in s["pnl_values"] if x > 0]
        avg_loss = sum(losses) / len(losses) if losses else 0
        avg_win = sum(wins) / len(wins) if wins else 0
        risk_penalty = max(0.0, (avg_loss - avg_win) * 0.20) if avg_loss and avg_win else 0.0

        score = (
            (s["weighted_pnl"] * 0.45)
            + (s["recent_weighted_pnl"] * 0.35)
            + (avg_pnl * 8)
            + ((win_rate - 0.5) * 120)
            + (sample_factor * 35)
            - risk_penalty
        )

        best_symbols = sorted(s["symbols"].items(), key=lambda x: x[1], reverse=True)[:3]
        scored.append({
            "setup": setup,
            "score": score,
            "trades": s["trades"],
            "pnl": round(s["pnl"], 2),
            "avg_pnl": round(avg_pnl, 2),
            "win_rate": round(win_rate * 100, 1),
            "best_symbols": [sym for sym, pnl in best_symbols if pnl > 0],
        })

    scored.sort(key=lambda x: x["score"], reverse=True)
    best = scored[0] if scored else None
    worst = min(scored, key=lambda x: x["score"]) if scored else None

    best_time = "No clear time edge"
    if time_stats:
        best_time_key, best_time_data = max(time_stats.items(), key=lambda x: x[1]["weighted_pnl"])
        if best_time_data["weighted_pnl"] > 0:
            best_time = best_time_key

    recent_window = parsed[-10:]
    older_window = parsed[:-10]
    recent_pnl = sum(float(r["pnl"] or 0) for r in recent_window)
    older_avg = (sum(float(r["pnl"] or 0) for r in older_window) / len(older_window)) if older_window else 0
    recent_avg = (recent_pnl / len(recent_window)) if recent_window else 0

    trend_warning = None
    if older_window and recent_avg < older_avg:
        trend_warning = "Performance is declining recently. Reduce size or demand cleaner A+ setups."

    if not best:
        return {
            "best_setup": "No clear setup yet",
            "avoid_setup": "No clear avoid yet",
            "best_time": best_time,
            "confidence": 0,
            "trend_warning": trend_warning,
            "reason": "Add more tagged trades to build a reliable smart plan.",
            "mode": "Collect data",
            "trade_rule": "Tag every trade before relying on setup recommendations.",
            "watchlist": [],
            "avoid_reasons": [],
        }

    tagged_count = sum(1 for r in parsed if r["setup"] != "Unlabeled")
    tag_quality = tagged_count / max(total_trades, 1)
    separation = abs((best["score"] or 0) - (worst["score"] or 0)) if worst else 0
    recent_strength = min(abs(recent_pnl) / 500, 1) * 15
    confidence = int(min(
        95,
        max(
            10,
            (min(best["trades"], 15) / 15) * 35
            + tag_quality * 25
            + min(separation / 150, 1) * 25
            + recent_strength,
        ),
    ))

    if trend_warning:
        confidence = max(5, confidence - 15)

    if best["setup"] == "Unlabeled":
        mode = "Journal first"
        trade_rule = "Do not increase size until your best trades have setup tags. The current edge is unlabeled."
    elif best["pnl"] > 0 and best["win_rate"] >= 50 and not trend_warning:
        mode = "Attack selectively"
        trade_rule = f"Prioritize {best['setup']} during your strongest session: {best_time}. Skip anything that does not match the plan."
    elif best["pnl"] > 0:
        mode = "Small size only"
        trade_rule = f"{best['setup']} is currently your best edge, but use reduced size until recent results stabilize."
    else:
        mode = "Defense day"
        trade_rule = "No setup has a strong positive edge right now. Focus on A+ trades only and reduce size."

    avoid_reasons = []
    if worst:
        avoid_reasons.append(
            f"{worst['setup']} has the weakest adaptive score: ${worst['pnl']:.0f} P&L across {worst['trades']} trade(s)."
        )
        if worst["win_rate"] < 40:
            avoid_reasons.append(f"Win rate is only {worst['win_rate']:.1f}% for that setup.")
    if trend_warning:
        avoid_reasons.append("Recent performance is weaker than older performance, so avoid forcing trades after losses.")

    reason = (
        f"{best['setup']} ranks highest after weighting recent trades, P&L (${best['pnl']:.0f}), "
        f"win rate ({best['win_rate']:.1f}%), and sample size ({best['trades']} trade(s))."
    )

    return {
        "best_setup": best["setup"],
        "avoid_setup": worst["setup"] if worst else "No clear avoid yet",
        "best_time": best_time,
        "confidence": confidence,
        "trend_warning": trend_warning,
        "reason": reason,
        "mode": mode,
        "trade_rule": trade_rule,
        "watchlist": best.get("best_symbols", []),
        "avoid_reasons": avoid_reasons,
    }



def enhance_tomorrow_trading_plan(plan, rows):
    """Add a concrete next-session execution plan on top of the adaptive setup recommendation."""
    plan = dict(plan or {})
    rows = rows or []

    pnl_values = [float(r.get("realized_pl") or r.get("pnl") or 0) for r in rows]
    recent = pnl_values[-10:]
    recent_losses = sum(1 for p in recent if p < 0)
    recent_net = sum(recent)

    if not rows:
        risk_mode = "Data collection"
        max_trades = 3
        daily_stop_hint = "Stop after 2 losses while you build history."
    elif plan.get("trend_warning") or recent_net < 0 or recent_losses >= 5:
        risk_mode = "Defense"
        max_trades = 2
        daily_stop_hint = "Stop after 1-2 losses or any rule break."
    elif int(plan.get("confidence") or 0) >= 70:
        risk_mode = "Selective attack"
        max_trades = 4
        daily_stop_hint = "Stop after 2 losses, or after giving back more than half of open profits."
    else:
        risk_mode = "Balanced"
        max_trades = 3
        daily_stop_hint = "Stop after 2 losses or if you feel tempted to force a setup."

    best_setup = plan.get("best_setup") or "your highest-quality tagged setup"
    avoid_setup = plan.get("avoid_setup") or "your weakest recent setup"
    best_time = plan.get("best_time") or "your strongest session"

    plan["risk_mode"] = risk_mode
    plan["max_trades"] = max_trades
    plan["daily_stop_hint"] = daily_stop_hint
    plan["pre_market_checklist"] = [
        f"Only trade {best_setup} if it appears cleanly.",
        f"Prefer {best_time}; avoid taking random trades outside your best window.",
        "Write entry, stop, target, and invalidation before entering.",
        "Size the trade from risk first, not from excitement.",
    ]
    plan["avoid_rules"] = [
        f"Avoid {avoid_setup} unless there is a major reason to override the data.",
        "No revenge trades after a loss.",
        "No adding risk after breaking your plan.",
        "No new trade if you cannot explain the setup in one sentence.",
    ]
    plan["tomorrow_script"] = (
        f"Tomorrow is a {risk_mode.lower()} day. Focus on {best_setup}, ideally during {best_time}. "
        f"Cap yourself at {max_trades} trade(s). {daily_stop_hint}"
    )
    return plan




# =========================================================
# IBKR settings routes
# =========================================================
@app.route("/settings/ibkr", methods=["GET", "POST"])
def ibkr_settings_page():
    if not require_login():
        return redirect(url_for("login"))

    message = ""
    error = ""

    if request.method == "POST":
        action = request.form.get("action", "save")
        if action == "clear":
            clear_ibkr_settings(current_user_id())
            message = "IBKR Flex settings cleared."
        else:
            flex_token = request.form.get("flex_token", "").strip()
            query_id = request.form.get("query_id", "").strip()
            account_id = request.form.get("account_id", "").strip()
            report_format = request.form.get("report_format", "xml").strip().lower()
            auto_import_enabled = request.form.get("auto_import_enabled") == "on"
            auto_import_hour = request.form.get("auto_import_hour", "6")

            if not flex_token or not query_id:
                error = "Flex Token and Query ID are required."
            else:
                save_ibkr_settings(
                    current_user_id(),
                    flex_token=flex_token,
                    query_id=query_id,
                    account_id=account_id,
                    report_format=report_format,
                    auto_import_enabled=auto_import_enabled,
                    auto_import_hour=auto_import_hour,
                )
                message = "IBKR Flex settings saved."

    settings = get_ibkr_settings(current_user_id()) or {}
    settings["masked_flex_token"] = mask_secret(settings.get("flex_token"))

    return render_template(
        "ibkr_settings.html",
        title="IBKR Settings",
        portfolios=get_user_portfolios(current_user_id()),
        active_portfolio_id=current_portfolio_id(),
        settings=settings,
        message=message,
        error=error,
    )



# =========================================================
# IBKR Flex Web Service fetcher
# =========================================================
IBKR_FLEX_BASE = "https://ndcdyn.interactivebrokers.com/AccountManagement/FlexWebService"


def xml_text(root, tag_name, default=""):
    """Return XML child text ignoring case."""
    if root is None:
        return default
    target = tag_name.lower()
    for child in list(root):
        if child.tag.split("}")[-1].lower() == target:
            return (child.text or default).strip()
    return default


def flex_request(endpoint, params, timeout=45):
    """Small stdlib HTTP client for IBKR Flex Web Service."""
    query = urllib.parse.urlencode(params)
    url = f"{IBKR_FLEX_BASE}/{endpoint}?{query}"
    req = urllib.request.Request(
        url,
        headers={
            # IBKR requires a User-Agent header for Flex Web Service requests.
            "User-Agent": "Python/TradeJournalFlexImporter",
            "Accept": "application/xml,text/xml,text/csv,text/plain,*/*",
        },
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        charset = resp.headers.get_content_charset() or "utf-8"
        return resp.read().decode(charset, errors="replace")


def ibkr_send_flex_request(token, query_id, from_date=None, to_date=None, period=None):
    params = {"t": token, "q": query_id, "v": "3"}
    if from_date and to_date:
        params["fd"] = from_date
        params["td"] = to_date
    elif period:
        params["p"] = period

    raw = flex_request("SendRequest", params)
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        raise RuntimeError(f"IBKR returned a non-XML SendRequest response: {raw[:250]}") from exc

    status = xml_text(root, "Status")
    if status.lower() != "success":
        code = xml_text(root, "ErrorCode", "UNKNOWN")
        msg = xml_text(root, "ErrorMessage", "IBKR Flex request failed.")
        raise RuntimeError(f"IBKR Flex request failed ({code}): {msg}")

    reference_code = xml_text(root, "ReferenceCode")
    if not reference_code:
        raise RuntimeError("IBKR did not return a ReferenceCode.")
    return reference_code


def ibkr_get_flex_statement(token, reference_code, max_attempts=6, wait_seconds=3):
    """Poll GetStatement until IBKR returns the generated report."""
    last_error = None
    for attempt in range(1, max_attempts + 1):
        raw = flex_request("GetStatement", {"t": token, "q": reference_code, "v": "3"})
        stripped = raw.lstrip()

        # Successful statements are usually FlexQueryResponse XML, but CSV/Text
        # output is also possible depending on the user's Flex Query settings.
        if stripped.startswith("<"):
            try:
                root = ET.fromstring(raw)
                status = xml_text(root, "Status")
                if status.lower() == "fail":
                    code = xml_text(root, "ErrorCode", "UNKNOWN")
                    msg = xml_text(root, "ErrorMessage", "IBKR statement not ready or failed.")
                    last_error = f"IBKR GetStatement failed ({code}): {msg}"
                    # Some IBKR responses mean the report is not ready yet; retry.
                    if attempt < max_attempts:
                        time.sleep(wait_seconds)
                        continue
                    raise RuntimeError(last_error)
            except ET.ParseError:
                # Non-standard XML-ish body; save and let parser/report handler deal with it.
                pass
        return raw

    raise RuntimeError(last_error or "IBKR Flex statement was not ready. Try again in a minute.")


def xml_local_name(tag):
    """Return an XML tag/attribute name without namespace."""
    return str(tag or "").split("}")[-1]


def flex_name_key(name):
    """Normalize IBKR Flex field names for forgiving matching."""
    return re.sub(r"[^a-z0-9]", "", str(name or "").lower())


def attr_or_child(elem, *names, default=""):
    """Read IBKR Flex values from XML attributes or direct child nodes.

    IBKR Flex reports usually store Trade data as attributes, but some query
    outputs/tools may produce child nodes. This helper matches common variants
    case-insensitively and ignores punctuation/underscores.
    """
    if elem is None:
        return default

    wanted = {flex_name_key(n) for n in names}

    for key, value in elem.attrib.items():
        if flex_name_key(key) in wanted and value not in (None, ""):
            return str(value).strip()

    for child in list(elem):
        if flex_name_key(xml_local_name(child.tag)) in wanted and child.text not in (None, ""):
            return str(child.text).strip()

    return default


def first_flex_value(elem, name_groups, default=""):
    for group in name_groups:
        value = attr_or_child(elem, *group, default="")
        if str(value).strip() != "":
            return value
    return default


def parse_ibkr_flex_datetime(raw):
    """Parse the date/time formats commonly returned by IBKR Flex XML."""
    raw = str(raw or "").strip()
    if not raw:
        return None

    raw = raw.replace(";", " ").replace("T", " ").strip()
    raw = re.sub(r"\s+", " ", raw)

    # Remove trailing timezone labels/offsets that strptime may not parse.
    candidates = [raw]
    candidates.append(re.sub(r"\s+[A-Z]{2,6}$", "", raw).strip())
    candidates.append(re.sub(r"\s+[+-]\d{2}:?\d{2}$", "", raw).strip())

    formats = (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%Y%m%d %H:%M:%S",
        "%Y%m%d %H:%M",
        "%Y%m%d",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
    )

    for candidate in candidates:
        if not candidate:
            continue
        for fmt in formats:
            try:
                return datetime.strptime(candidate, fmt)
            except ValueError:
                continue

    return parse_dt_any(raw)


def clean_ibkr_symbol(asset_category, symbol, description):
    symbol = str(symbol or "").strip()
    description = str(description or "").strip()
    asset = str(asset_category or "").upper()

    # For options, the Flex description often carries the useful contract label.
    if ("OPT" in asset or "OPTION" in asset) and description:
        return description
    return symbol or description


def parse_ibkr_flex_trade_element(elem):
    """Convert one IBKR Flex <Trade> style element into the app trade dict."""
    level = attr_or_child(elem, "levelOfDetail", "level", default="").strip().upper()
    if level and level not in {"EXECUTION", "TRADE", "ORDER", ""}:
        # Skip subtotal/summary rows if the Flex Query includes them.
        if "SUMMARY" in level or "TOTAL" in level:
            return None

    asset_category = first_flex_value(elem, [
        ("assetCategory", "assetClass", "securityType", "secType"),
    ], default="")
    description = attr_or_child(elem, "description", "securityDescription", default="")
    symbol_raw = first_flex_value(elem, [
        ("symbol", "ibSymbol", "underlyingSymbol", "contractSymbol"),
        ("description",),
    ], default="")
    symbol = clean_ibkr_symbol(asset_category, symbol_raw, description)
    if not symbol or "total" in symbol.lower():
        return None

    currency = first_flex_value(elem, [("currency", "tradeCurrency", "fxCurrency")], default="") or "USD"

    dt_raw = first_flex_value(elem, [
        ("dateTime", "tradeDateTime", "tradeTime", "orderTime"),
        ("tradeDate", "date", "reportDate"),
    ], default="")
    trade_dt = parse_ibkr_flex_datetime(dt_raw)

    quantity = to_decimal(first_flex_value(elem, [("quantity", "qty", "shares", "units")], default="0"))
    trade_price = to_decimal(first_flex_value(elem, [("tradePrice", "price", "executionPrice", "avgPrice")], default="0"))
    close_price = to_decimal(first_flex_value(elem, [("closePrice", "markPrice")], default="0"))
    proceeds = to_decimal(first_flex_value(elem, [("proceeds", "netCash", "cash", "amount")], default="0"))
    commission = to_decimal(first_flex_value(elem, [("ibCommission", "commission", "brokerageCommission", "fees")], default="0"))
    basis = to_decimal(first_flex_value(elem, [("costBasis", "basis", "cost")], default="0"))
    realized_pl = to_decimal(first_flex_value(elem, [("fifoPnlRealized", "realizedPnl", "realizedPL", "realizedPnL", "pnl")], default="0"))
    mtm_pl = to_decimal(first_flex_value(elem, [("mtmPnl", "mtmPL", "mtmPnL", "markToMarketPnl")], default="0"))

    buy_sell = first_flex_value(elem, [("buySell", "side", "action", "transactionType")], default="").upper()
    if buy_sell in {"BUY", "BOT", "B", "BOUGHT"}:
        side = "BUY"
        if quantity < 0:
            quantity = -quantity
    elif buy_sell in {"SELL", "SLD", "S", "SOLD"}:
        side = "SELL"
        if quantity > 0:
            quantity = -quantity
    else:
        side = "BUY" if quantity > 0 else "SELL"

    code = first_flex_value(elem, [
        ("ibExecID", "ibExecId", "execID", "execId", "executionID", "executionId"),
        ("tradeID", "tradeId", "transactionID", "transactionId"),
        ("ibOrderID", "ibOrderId", "orderID", "orderId"),
        ("code",),
    ], default="") or "IBKR_FLEX"

    notes_parts = []
    account_id = attr_or_child(elem, "accountId", "account", default="")
    if account_id:
        notes_parts.append(f"Account: {account_id}")
    if description and description != symbol:
        notes_parts.append(description)
    notes_parts.append("Imported automatically from IBKR Flex Web Service.")

    return {
        "broker": "IBKR Flex",
        "asset_category": asset_category or "UNKNOWN",
        "currency": currency,
        "symbol": symbol.strip(),
        "trade_datetime": trade_dt.isoformat() if trade_dt else None,
        "quantity": float(quantity),
        "side": side,
        "trade_price": float(trade_price),
        "close_price": float(close_price),
        "buy_price": float(trade_price) if side == "BUY" else 0.0,
        "sell_price": float(trade_price) if side == "SELL" else 0.0,
        "proceeds": float(proceeds),
        "commission": float(commission),
        "basis": float(basis),
        "realized_pl": float(realized_pl),
        "mtm_pl": float(mtm_pl),
        "code": str(code).strip(),
        "notes": " | ".join(notes_parts),
        "risk_amount": 0.0,
        "r_multiple": 0.0,
    }


def parse_ibkr_flex_xml(file_path):
    """Parse IBKR Flex XML into normalized app trade rows.

    Supports common IBKR Flex Activity XML shapes such as:
    <FlexQueryResponse><FlexStatements><FlexStatement><Trades><Trade ... /></Trades>

    The parser is intentionally forgiving because field names vary depending on
    which columns were selected in the Flex Query.
    """
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
    except ET.ParseError as exc:
        raise RuntimeError("Could not parse IBKR Flex XML. Make sure the Flex Query output format is XML.") from exc

    # Surface IBKR error responses clearly.
    status = xml_text(root, "Status", "")
    if status.lower() == "fail":
        code = xml_text(root, "ErrorCode", "UNKNOWN")
        msg = xml_text(root, "ErrorMessage", "IBKR Flex XML response failed.")
        raise RuntimeError(f"IBKR Flex XML error ({code}): {msg}")

    trades = []
    seen_keys = set()
    trade_tags = {"trade", "tradeconfirm", "tradesummary"}

    for elem in root.iter():
        tag = xml_local_name(elem.tag).lower()
        if tag not in trade_tags:
            continue

        trade = parse_ibkr_flex_trade_element(elem)
        if not trade:
            continue

        # Avoid double-counting duplicate XML nodes inside one report.
        key = (
            trade.get("broker"),
            trade.get("symbol"),
            trade.get("trade_datetime"),
            round(float(trade.get("quantity") or 0), 8),
            round(float(trade.get("trade_price") or 0), 8),
            round(float(trade.get("proceeds") or 0), 8),
            round(float(trade.get("commission") or 0), 8),
            trade.get("code"),
        )
        if key in seen_keys:
            continue
        seen_keys.add(key)
        trades.append(trade)

    if not trades:
        raise RuntimeError(
            "No IBKR Flex trades found in the XML. Edit your IBKR Activity Flex Query to include the Trades section and fields such as Symbol, Date/Time, Quantity, Trade Price, Proceeds, Commission, and Realized P&L."
        )

    return trades


def save_flex_report(raw_report, report_format="xml"):
    ext = "csv" if report_format == "csv" or not raw_report.lstrip().startswith("<") else "xml"
    filename = f"ibkr_flex_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.{ext}"
    path = os.path.join(UPLOAD_FOLDER, filename)
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write(raw_report)
    return filename, path, ext


def update_ibkr_last_import(user_id):
    conn = get_db_connection()
    conn.execute(
        "UPDATE ibkr_settings SET last_import_at = CURRENT_TIMESTAMP, updated_at = CURRENT_TIMESTAMP WHERE user_id = ?",
        (user_id,),
    )
    conn.commit()
    conn.close()


@app.route("/import/ibkr/flex", methods=["POST"])
def import_ibkr_flex():
    if not require_login():
        return redirect(url_for("login"))

    settings = get_ibkr_settings(current_user_id())
    if not settings or not settings.get("flex_token") or not settings.get("query_id"):
        return redirect(url_for("ibkr_settings_page"))

    try:
        reference_code = ibkr_send_flex_request(
            settings["flex_token"],
            settings["query_id"],
        )
        raw_report = ibkr_get_flex_statement(settings["flex_token"], reference_code)
        filename, save_path, ext = save_flex_report(raw_report, settings.get("report_format") or "xml")

        if ext == "xml":
            trades = parse_ibkr_flex_xml(save_path)
        else:
            # If user configured the Flex Query as CSV, first try the existing
            # IBKR Activity parser/detector. XML is still recommended.
            import_type = detect_import_type(save_path)
            if import_type in {"ibkr", "ibkr_trades"}:
                trades = parse_ibkr_activity_csv(save_path)
            elif import_type == "ibkr_summary":
                trades = parse_ibkr_summary_csv(save_path)
            else:
                raise RuntimeError("Downloaded IBKR Flex report is CSV/Text but does not match the existing IBKR CSV parser. Set the Flex Query format to XML or include the standard Trades fields.")

        batch_id = new_batch_id("IBKRFLEX")
        inserted, skipped = insert_trades(trades, import_file=filename, batch_id=batch_id)
        update_ibkr_last_import(current_user_id())

        return import_success_page(
            "IBKR Flex Import Complete",
            "IBKR Flex Web Service",
            batch_id,
            len(trades),
            inserted,
            skipped,
        )

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500



# =========================================================
# IBKR scheduled daily auto-import
# =========================================================
def get_or_create_default_portfolio_id_for_user(user_id):
    """Background-safe default portfolio lookup/creation.

    Do not use current_portfolio_id() here because scheduled jobs do not have a
    Flask request/session context.
    """
    if not user_id:
        return None

    conn = get_db_connection()
    cur = conn.cursor()
    row = cur.execute(
        """
        SELECT id
        FROM portfolios
        WHERE user_id = ?
        ORDER BY id ASC
        LIMIT 1
        """,
        (user_id,),
    ).fetchone()

    if row:
        portfolio_id = row["id"]
    else:
        cur.execute(
            "INSERT INTO portfolios (user_id, name) VALUES (?, ?)",
            (user_id, "Main Portfolio"),
        )
        portfolio_id = cur.lastrowid
        conn.commit()

    conn.close()
    return portfolio_id


def insert_trades_for_user_portfolio(user_id, portfolio_id, trades, import_file=None, batch_id=None):
    """Same duplicate-safe insert as insert_trades(), but usable by background jobs."""
    if not user_id:
        raise ValueError("Missing user_id for scheduled import.")
    if not portfolio_id:
        raise ValueError("Missing portfolio_id for scheduled import.")

    conn = get_db_connection()
    cur = conn.cursor()
    inserted = 0
    skipped = 0
    seen_in_this_import = set()

    for t in trades:
        trade_key = make_trade_key(user_id, portfolio_id, t)

        if trade_key in seen_in_this_import:
            skipped += 1
            continue
        seen_in_this_import.add(trade_key)

        existing = cur.execute(
            """
            SELECT id
            FROM trades
            WHERE user_id = ? AND portfolio_id = ? AND trade_key = ?
            LIMIT 1
            """,
            (user_id, portfolio_id, trade_key),
        ).fetchone()

        if existing is None:
            existing = cur.execute(
                """
                SELECT id
                FROM trades
                WHERE user_id = ?
                  AND portfolio_id = ?
                  AND broker = ?
                  AND symbol = ?
                  AND IFNULL(trade_datetime, '') = IFNULL(?, '')
                  AND ROUND(IFNULL(quantity, 0), 8) = ROUND(IFNULL(?, 0), 8)
                  AND IFNULL(side, '') = IFNULL(?, '')
                  AND ROUND(IFNULL(trade_price, 0), 8) = ROUND(IFNULL(?, 0), 8)
                  AND ROUND(IFNULL(proceeds, 0), 8) = ROUND(IFNULL(?, 0), 8)
                  AND ROUND(IFNULL(commission, 0), 8) = ROUND(IFNULL(?, 0), 8)
                  AND IFNULL(code, '') = IFNULL(?, '')
                LIMIT 1
                """,
                (
                    user_id,
                    portfolio_id,
                    t["broker"],
                    t["symbol"],
                    t["trade_datetime"],
                    t["quantity"],
                    t.get("side"),
                    t["trade_price"],
                    t["proceeds"],
                    t["commission"],
                    t["code"],
                ),
            ).fetchone()

        if existing is not None:
            skipped += 1
            continue

        buy_price, sell_price = derive_buy_sell_prices(t)

        cur.execute(
            """
            INSERT INTO trades (
                user_id, portfolio_id, broker, asset_category, currency, symbol,
                trade_datetime, quantity, side, trade_price, close_price, buy_price, sell_price, proceeds,
                commission, basis, realized_pl, mtm_pl, risk_amount, r_multiple,
                code, import_file, batch_id, notes, trade_key
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                user_id,
                portfolio_id,
                t["broker"],
                t["asset_category"],
                t["currency"],
                t["symbol"],
                t["trade_datetime"],
                t["quantity"],
                t["side"],
                t["trade_price"],
                t["close_price"],
                buy_price,
                sell_price,
                t["proceeds"],
                t["commission"],
                t["basis"],
                t["realized_pl"],
                t["mtm_pl"],
                t.get("risk_amount", 0),
                t.get("r_multiple", 0),
                t["code"],
                import_file,
                batch_id,
                t.get("notes"),
                trade_key,
            ),
        )
        inserted += 1

    conn.commit()
    conn.close()
    return inserted, skipped


def get_ibkr_auto_import_users():
    conn = get_db_connection()
    rows = conn.execute(
        """
        SELECT user_id, flex_token, query_id, account_id, report_format,
               auto_import_enabled, auto_import_hour, last_auto_import_date
        FROM ibkr_settings
        WHERE auto_import_enabled = 1
          AND TRIM(IFNULL(flex_token, '')) != ''
          AND TRIM(IFNULL(query_id, '')) != ''
        """
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_ibkr_auto_import_status(user_id, status, import_date=None):
    import_date = import_date or datetime.now().strftime("%Y-%m-%d")
    conn = get_db_connection()
    conn.execute(
        """
        UPDATE ibkr_settings
        SET last_auto_import_date = ?,
            last_auto_import_status = ?,
            updated_at = CURRENT_TIMESTAMP
        WHERE user_id = ?
        """,
        (import_date, str(status)[:500], user_id),
    )
    conn.commit()
    conn.close()


def run_ibkr_flex_import_for_user(user_id, settings):
    """Run one IBKR Flex import for one user without needing a request session."""
    portfolio_id = get_or_create_default_portfolio_id_for_user(user_id)
    if not portfolio_id:
        raise RuntimeError("Could not find or create a portfolio for this user.")

    reference_code = ibkr_send_flex_request(settings["flex_token"], settings["query_id"])
    raw_report = ibkr_get_flex_statement(settings["flex_token"], reference_code)
    filename, save_path, ext = save_flex_report(raw_report, settings.get("report_format") or "xml")

    if ext == "xml":
        trades = parse_ibkr_flex_xml(save_path)
    else:
        import_type = detect_import_type(save_path)
        if import_type in {"ibkr", "ibkr_trades"}:
            trades = parse_ibkr_activity_csv(save_path)
        elif import_type == "ibkr_summary":
            trades = parse_ibkr_summary_csv(save_path)
        else:
            raise RuntimeError(
                "Downloaded IBKR Flex report is not compatible with the CSV parser. Set the Flex Query format to XML."
            )

    batch_id = new_batch_id("IBKRAUTO")
    inserted, skipped = insert_trades_for_user_portfolio(
        user_id,
        portfolio_id,
        trades,
        import_file=filename,
        batch_id=batch_id,
    )

    conn = get_db_connection()
    conn.execute(
        """
        UPDATE ibkr_settings
        SET last_import_at = CURRENT_TIMESTAMP,
            updated_at = CURRENT_TIMESTAMP
        WHERE user_id = ?
        """,
        (user_id,),
    )
    conn.commit()
    conn.close()

    return {
        "batch_id": batch_id,
        "parsed": len(trades),
        "inserted": inserted,
        "skipped": skipped,
    }


_ibkr_scheduler_started = False


def ibkr_auto_import_scheduler_loop():
    """Simple daily scheduler.

    It checks periodically and imports once per calendar day for each user whose
    IBKR settings have auto-import enabled. The schedule uses the machine/server
    local time. Default hour is 6 AM unless changed on the IBKR Settings page.
    """
    # Give Flask/SQLite a moment to finish startup before first check.
    time.sleep(5)

    while True:
        try:
            now = datetime.now()
            today = now.strftime("%Y-%m-%d")
            users = get_ibkr_auto_import_users()

            for settings in users:
                user_id = settings["user_id"]
                try:
                    hour = int(settings.get("auto_import_hour") or 6)
                except (TypeError, ValueError):
                    hour = 6
                hour = max(0, min(23, hour))

                already_ran_today = settings.get("last_auto_import_date") == today
                if already_ran_today or now.hour < hour:
                    continue

                try:
                    result = run_ibkr_flex_import_for_user(user_id, settings)
                    update_ibkr_auto_import_status(
                        user_id,
                        f"OK: inserted {result['inserted']}, skipped {result['skipped']}, batch {result['batch_id']}",
                        today,
                    )
                except Exception as exc:
                    update_ibkr_auto_import_status(user_id, f"ERROR: {exc}", today)

        except Exception as exc:
            print(f"IBKR auto-import scheduler error: {exc}")

        # Check every 10 minutes. It still only runs once per day per user.
        time.sleep(600)


def start_ibkr_auto_import_scheduler():
    global _ibkr_scheduler_started
    if _ibkr_scheduler_started:
        return
    if os.environ.get("IBKR_AUTO_IMPORT_DISABLED") == "1":
        return
    _ibkr_scheduler_started = True
    thread = threading.Thread(target=ibkr_auto_import_scheduler_loop, daemon=True)
    thread.start()



# =========================================================
# NinjaTrader folder auto-import
# =========================================================
def get_ninjatrader_settings(user_id=None):
    user_id = user_id or current_user_id()
    if not user_id:
        return None
    conn = get_db_connection()
    row = conn.execute(
        """
        SELECT user_id, folder_path, auto_import_enabled, scan_interval_minutes,
               last_import_at, last_scan_at, last_scan_status, updated_at
        FROM ninjatrader_settings
        WHERE user_id = ?
        """,
        (user_id,),
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def save_ninjatrader_settings(user_id, folder_path, auto_import_enabled=False, scan_interval_minutes=10):
    folder_path = os.path.abspath(os.path.expanduser(str(folder_path or "").strip())) if str(folder_path or "").strip() else ""
    try:
        scan_interval_minutes = int(scan_interval_minutes)
    except (TypeError, ValueError):
        scan_interval_minutes = 10
    scan_interval_minutes = max(1, min(1440, scan_interval_minutes))

    conn = get_db_connection()
    conn.execute(
        """
        INSERT INTO ninjatrader_settings (
            user_id, folder_path, auto_import_enabled, scan_interval_minutes, updated_at
        )
        VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(user_id) DO UPDATE SET
            folder_path = excluded.folder_path,
            auto_import_enabled = excluded.auto_import_enabled,
            scan_interval_minutes = excluded.scan_interval_minutes,
            updated_at = CURRENT_TIMESTAMP
        """,
        (user_id, folder_path, 1 if auto_import_enabled else 0, scan_interval_minutes),
    )
    conn.commit()
    conn.close()


def clear_ninjatrader_settings(user_id):
    conn = get_db_connection()
    conn.execute("DELETE FROM ninjatrader_settings WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()


def file_sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def has_imported_file(user_id, source, file_hash):
    conn = get_db_connection()
    row = conn.execute(
        """
        SELECT id, batch_id, inserted_count FROM imported_files
        WHERE user_id = ? AND source = ? AND file_hash = ?
        LIMIT 1
        """,
        (user_id, source, file_hash),
    ).fetchone()

    if not row:
        conn.close()
        return False

    batch_id = row["batch_id"]

    # Important repair path: if the user deleted the import batch but the
    # imported_files hash row was left behind by an older app version, do not
    # block re-import. Clean the stale hash row and allow the file through.
    if batch_id:
        trade_row = conn.execute(
            """
            SELECT id FROM trades
            WHERE user_id = ? AND batch_id = ?
            LIMIT 1
            """,
            (user_id, batch_id),
        ).fetchone()
        if not trade_row:
            conn.execute("DELETE FROM imported_files WHERE id = ?", (row["id"],))
            conn.commit()
            conn.close()
            return False

    conn.close()
    return True


def record_imported_file(user_id, source, file_path, file_hash, batch_id, parsed_count, inserted_count, skipped_count, status):
    try:
        stat = os.stat(file_path)
        file_size = int(stat.st_size)
        file_mtime = float(stat.st_mtime)
    except OSError:
        file_size = 0
        file_mtime = 0.0

    conn = get_db_connection()
    conn.execute(
        """
        INSERT OR IGNORE INTO imported_files (
            user_id, source, file_path, file_name, file_hash, file_size, file_mtime,
            batch_id, parsed_count, inserted_count, skipped_count, status
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            user_id,
            source,
            file_path,
            os.path.basename(file_path),
            file_hash,
            file_size,
            file_mtime,
            batch_id,
            int(parsed_count or 0),
            int(inserted_count or 0),
            int(skipped_count or 0),
            str(status or "")[:500],
        ),
    )
    conn.commit()
    conn.close()


def update_ninjatrader_scan_status(user_id, status, imported=False):
    conn = get_db_connection()
    if imported:
        conn.execute(
            """
            UPDATE ninjatrader_settings
            SET last_scan_at = CURRENT_TIMESTAMP,
                last_import_at = CURRENT_TIMESTAMP,
                last_scan_status = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE user_id = ?
            """,
            (str(status)[:500], user_id),
        )
    else:
        conn.execute(
            """
            UPDATE ninjatrader_settings
            SET last_scan_at = CURRENT_TIMESTAMP,
                last_scan_status = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE user_id = ?
            """,
            (str(status)[:500], user_id),
        )
    conn.commit()
    conn.close()


def list_ninjatrader_csv_files(folder_path):
    if not folder_path or not os.path.isdir(folder_path):
        return []
    out = []
    for name in os.listdir(folder_path):
        path = os.path.join(folder_path, name)
        if os.path.isfile(path) and name.lower().endswith(".csv"):
            out.append(path)
    out.sort(key=lambda p: (os.path.getmtime(p), p))
    return out


def run_ninjatrader_folder_import_for_user(user_id, settings=None, force=False):
    settings = settings or get_ninjatrader_settings(user_id)
    if not settings:
        raise RuntimeError("NinjaTrader folder settings are not configured.")

    folder_path = os.path.abspath(os.path.expanduser(str(settings.get("folder_path") or "").strip()))
    if not folder_path:
        raise RuntimeError("NinjaTrader export folder is blank.")
    if not os.path.isdir(folder_path):
        raise RuntimeError(f"NinjaTrader folder does not exist: {folder_path}")

    portfolio_id = get_or_create_default_portfolio_id_for_user(user_id)
    if not portfolio_id:
        raise RuntimeError("Could not find or create a portfolio for this user.")

    files = list_ninjatrader_csv_files(folder_path)
    if not files:
        update_ninjatrader_scan_status(user_id, "No CSV files found.", imported=False)
        return {"files_seen": 0, "files_imported": 0, "parsed": 0, "inserted": 0, "skipped": 0, "messages": ["No CSV files found."]}

    totals = {"files_seen": len(files), "files_imported": 0, "parsed": 0, "inserted": 0, "skipped": 0, "messages": []}

    for path in files:
        try:
            digest = file_sha256(path)
            if not force and has_imported_file(user_id, "ninjatrader_folder", digest):
                totals["messages"].append(f"Skipped already imported file: {os.path.basename(path)}")
                continue

            import_type = detect_import_type(path)
            if import_type != "performance":
                status = "Skipped: not a NinjaTrader Performance CSV"
                record_imported_file(user_id, "ninjatrader_folder", path, digest, None, 0, 0, 0, status)
                totals["messages"].append(f"{status}: {os.path.basename(path)}")
                continue

            trades = parse_performance_csv(path)
            batch_id = new_batch_id("NTAUTO")
            inserted, skipped = insert_trades_for_user_portfolio(
                user_id,
                portfolio_id,
                trades,
                import_file=os.path.basename(path),
                batch_id=batch_id,
            )
            record_imported_file(
                user_id,
                "ninjatrader_folder",
                path,
                digest,
                batch_id,
                len(trades),
                inserted,
                skipped,
                "OK",
            )
            totals["files_imported"] += 1
            totals["parsed"] += len(trades)
            totals["inserted"] += inserted
            totals["skipped"] += skipped
            totals["messages"].append(f"Imported {os.path.basename(path)}: inserted {inserted}, skipped {skipped}, batch {batch_id}")
        except Exception as exc:
            totals["messages"].append(f"ERROR importing {os.path.basename(path)}: {exc}")

    status = f"Scanned {totals['files_seen']} file(s). Imported {totals['files_imported']}. Inserted {totals['inserted']}, skipped {totals['skipped']}."
    update_ninjatrader_scan_status(user_id, status, imported=totals["inserted"] > 0 or totals["files_imported"] > 0)
    return totals


@app.route("/settings/ninjatrader", methods=["GET", "POST"])
def ninjatrader_settings_page():
    if not require_login():
        return redirect(url_for("login"))

    message = ""
    error = ""
    result = None

    if request.method == "POST":
        action = request.form.get("action", "save")
        try:
            if action == "clear":
                clear_ninjatrader_settings(current_user_id())
                message = "NinjaTrader settings cleared."
            elif action == "import_now":
                result = run_ninjatrader_folder_import_for_user(current_user_id())
                message = f"Folder scan complete: inserted {result['inserted']}, skipped {result['skipped']}."
            else:
                save_ninjatrader_settings(
                    current_user_id(),
                    request.form.get("folder_path", ""),
                    auto_import_enabled=bool(request.form.get("auto_import_enabled")),
                    scan_interval_minutes=request.form.get("scan_interval_minutes", 10),
                )
                message = "NinjaTrader settings saved."
        except Exception as exc:
            error = str(exc)

    settings = get_ninjatrader_settings(current_user_id()) or {}
    settings.setdefault("folder_path", "")
    settings.setdefault("auto_import_enabled", 0)
    settings.setdefault("scan_interval_minutes", 10)

    return render_template(
        "ninjatrader_settings.html",
        title="NinjaTrader Settings",
        portfolios=get_user_portfolios(current_user_id()),
        active_portfolio_id=current_portfolio_id(),
        settings=settings,
        message=message,
        error=error,
        result=result,
    )


def get_ninjatrader_auto_import_users():
    conn = get_db_connection()
    rows = conn.execute(
        """
        SELECT user_id, folder_path, auto_import_enabled, scan_interval_minutes,
               last_scan_at, last_scan_status
        FROM ninjatrader_settings
        WHERE auto_import_enabled = 1
          AND TRIM(IFNULL(folder_path, '')) != ''
        """
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def should_run_ninjatrader_scan(settings, now=None):
    now = now or datetime.now()
    try:
        minutes = int(settings.get("scan_interval_minutes") or 10)
    except (TypeError, ValueError):
        minutes = 10
    minutes = max(1, min(1440, minutes))

    last_scan_at = settings.get("last_scan_at")
    if not last_scan_at:
        return True
    try:
        last_dt = datetime.fromisoformat(str(last_scan_at).replace("Z", ""))
    except Exception:
        try:
            last_dt = datetime.strptime(str(last_scan_at)[:19], "%Y-%m-%d %H:%M:%S")
        except Exception:
            return True
    return (now - last_dt).total_seconds() >= minutes * 60


_ninjatrader_scheduler_started = False


def ninjatrader_folder_scheduler_loop():
    time.sleep(8)
    while True:
        try:
            now = datetime.now()
            users = get_ninjatrader_auto_import_users()
            for settings in users:
                if not should_run_ninjatrader_scan(settings, now):
                    continue
                user_id = settings["user_id"]
                try:
                    result = run_ninjatrader_folder_import_for_user(user_id, settings=settings)
                    update_ninjatrader_scan_status(
                        user_id,
                        f"OK: scanned {result['files_seen']} file(s), imported {result['files_imported']}, inserted {result['inserted']}, skipped {result['skipped']}.",
                        imported=result.get("files_imported", 0) > 0,
                    )
                except Exception as exc:
                    update_ninjatrader_scan_status(user_id, f"ERROR: {exc}", imported=False)
        except Exception as exc:
            print(f"NinjaTrader folder scheduler error: {exc}")

        time.sleep(60)


def start_ninjatrader_folder_scheduler():
    global _ninjatrader_scheduler_started
    if _ninjatrader_scheduler_started:
        return
    if os.environ.get("NINJATRADER_AUTO_IMPORT_DISABLED") == "1":
        return
    _ninjatrader_scheduler_started = True
    thread = threading.Thread(target=ninjatrader_folder_scheduler_loop, daemon=True)
    thread.start()



# =========================================================
# Wealthsimple folder auto-import
# =========================================================
def get_wealthsimple_settings(user_id=None):
    user_id = user_id or current_user_id()
    if not user_id:
        return None
    conn = get_db_connection()
    row = conn.execute(
        """
        SELECT user_id, folder_path, auto_import_enabled, scan_interval_minutes,
               last_import_at, last_scan_at, last_scan_status, updated_at
        FROM wealthsimple_settings
        WHERE user_id = ?
        """,
        (user_id,),
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def save_wealthsimple_settings(user_id, folder_path, auto_import_enabled=False, scan_interval_minutes=10):
    folder_path = os.path.abspath(os.path.expanduser(str(folder_path or "").strip())) if str(folder_path or "").strip() else ""
    try:
        scan_interval_minutes = int(scan_interval_minutes)
    except (TypeError, ValueError):
        scan_interval_minutes = 10
    scan_interval_minutes = max(1, min(1440, scan_interval_minutes))

    conn = get_db_connection()
    conn.execute(
        """
        INSERT INTO wealthsimple_settings (
            user_id, folder_path, auto_import_enabled, scan_interval_minutes, updated_at
        )
        VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(user_id) DO UPDATE SET
            folder_path = excluded.folder_path,
            auto_import_enabled = excluded.auto_import_enabled,
            scan_interval_minutes = excluded.scan_interval_minutes,
            updated_at = CURRENT_TIMESTAMP
        """,
        (user_id, folder_path, 1 if auto_import_enabled else 0, scan_interval_minutes),
    )
    conn.commit()
    conn.close()


def clear_wealthsimple_settings(user_id):
    conn = get_db_connection()
    conn.execute("DELETE FROM wealthsimple_settings WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()


def update_wealthsimple_scan_status(user_id, status, imported=False):
    conn = get_db_connection()
    if imported:
        conn.execute(
            """
            UPDATE wealthsimple_settings
            SET last_scan_at = CURRENT_TIMESTAMP,
                last_import_at = CURRENT_TIMESTAMP,
                last_scan_status = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE user_id = ?
            """,
            (str(status)[:500], user_id),
        )
    else:
        conn.execute(
            """
            UPDATE wealthsimple_settings
            SET last_scan_at = CURRENT_TIMESTAMP,
                last_scan_status = ?,
                updated_at = CURRENT_TIMESTAMP
            WHERE user_id = ?
            """,
            (str(status)[:500], user_id),
        )
    conn.commit()
    conn.close()


def list_wealthsimple_csv_files(folder_path):
    if not folder_path or not os.path.isdir(folder_path):
        return []
    out = []
    for name in os.listdir(folder_path):
        path = os.path.join(folder_path, name)
        if os.path.isfile(path) and name.lower().endswith(".csv"):
            out.append(path)
    out.sort(key=lambda p: (os.path.getmtime(p), p))
    return out


def run_wealthsimple_folder_import_for_user(user_id, settings=None, force=False):
    settings = settings or get_wealthsimple_settings(user_id)
    if not settings:
        raise RuntimeError("Wealthsimple folder settings are not configured.")

    folder_path = os.path.abspath(os.path.expanduser(str(settings.get("folder_path") or "").strip()))
    if not folder_path:
        raise RuntimeError("Wealthsimple export folder is blank.")
    if not os.path.isdir(folder_path):
        raise RuntimeError(f"Wealthsimple folder does not exist: {folder_path}")

    portfolio_id = get_or_create_default_portfolio_id_for_user(user_id)
    if not portfolio_id:
        raise RuntimeError("Could not find or create a portfolio for this user.")

    files = list_wealthsimple_csv_files(folder_path)
    if not files:
        update_wealthsimple_scan_status(user_id, "No CSV files found.", imported=False)
        return {"files_seen": 0, "files_imported": 0, "parsed": 0, "inserted": 0, "skipped": 0, "messages": ["No CSV files found."]}

    totals = {"files_seen": len(files), "files_imported": 0, "parsed": 0, "inserted": 0, "skipped": 0, "messages": []}

    for path in files:
        try:
            digest = file_sha256(path)
            if not force and has_imported_file(user_id, "wealthsimple_folder", digest):
                totals["messages"].append(f"Skipped already imported file: {os.path.basename(path)}")
                continue

            import_type = detect_import_type(path)
            if import_type != "wealthsimple":
                # Some Wealthsimple exports vary slightly. Try the parser before rejecting the file.
                try:
                    parsed_preview = parse_wealthsimple_csv(path)
                except Exception:
                    parsed_preview = []
                if not parsed_preview:
                    status = "Skipped: not a Wealthsimple activity CSV"
                    record_imported_file(user_id, "wealthsimple_folder", path, digest, None, 0, 0, 0, status)
                    totals["messages"].append(f"{status}: {os.path.basename(path)}")
                    continue
                trades = parsed_preview
            else:
                trades = parse_wealthsimple_csv(path)

            existing_ws_trades = get_existing_wealthsimple_trades_for_user_portfolio(user_id, portfolio_id)
            updates_for_existing, new_trades_with_fifo = recompute_fifo_for_wealthsimple(existing_ws_trades, trades)
            update_existing_trade_fifo_values(updates_for_existing)

            batch_id = new_batch_id("WSAUTO")
            inserted, skipped = insert_trades_for_user_portfolio(
                user_id,
                portfolio_id,
                new_trades_with_fifo,
                import_file=os.path.basename(path),
                batch_id=batch_id,
            )
            record_imported_file(
                user_id,
                "wealthsimple_folder",
                path,
                digest,
                batch_id,
                len(trades),
                inserted,
                skipped,
                "OK",
            )
            totals["files_imported"] += 1
            totals["parsed"] += len(trades)
            totals["inserted"] += inserted
            totals["skipped"] += skipped
            totals["messages"].append(f"Imported {os.path.basename(path)}: inserted {inserted}, skipped {skipped}, batch {batch_id}")
        except Exception as exc:
            totals["messages"].append(f"ERROR importing {os.path.basename(path)}: {exc}")

    status = f"Scanned {totals['files_seen']} file(s). Imported {totals['files_imported']}. Inserted {totals['inserted']}, skipped {totals['skipped']}."
    update_wealthsimple_scan_status(user_id, status, imported=totals["inserted"] > 0 or totals["files_imported"] > 0)
    return totals


@app.route("/settings/wealthsimple", methods=["GET", "POST"])
def wealthsimple_settings_page():
    if not require_login():
        return redirect(url_for("login"))

    message = ""
    error = ""
    result = None

    if request.method == "POST":
        action = request.form.get("action", "save")
        try:
            if action == "clear":
                clear_wealthsimple_settings(current_user_id())
                message = "Wealthsimple settings cleared."
            elif action == "import_now":
                result = run_wealthsimple_folder_import_for_user(current_user_id())
                message = f"Folder scan complete: inserted {result['inserted']}, skipped {result['skipped']}."
            else:
                save_wealthsimple_settings(
                    current_user_id(),
                    request.form.get("folder_path", ""),
                    auto_import_enabled=bool(request.form.get("auto_import_enabled")),
                    scan_interval_minutes=request.form.get("scan_interval_minutes", 10),
                )
                message = "Wealthsimple settings saved."
        except Exception as exc:
            error = str(exc)

    settings = get_wealthsimple_settings(current_user_id()) or {}
    settings.setdefault("folder_path", "")
    settings.setdefault("auto_import_enabled", 0)
    settings.setdefault("scan_interval_minutes", 10)

    return render_template(
        "wealthsimple_settings.html",
        title="Wealthsimple Settings",
        portfolios=get_user_portfolios(current_user_id()),
        active_portfolio_id=current_portfolio_id(),
        settings=settings,
        message=message,
        error=error,
        result=result,
    )


def get_wealthsimple_auto_import_users():
    conn = get_db_connection()
    rows = conn.execute(
        """
        SELECT user_id, folder_path, auto_import_enabled, scan_interval_minutes,
               last_scan_at, last_scan_status
        FROM wealthsimple_settings
        WHERE auto_import_enabled = 1
          AND TRIM(IFNULL(folder_path, '')) != ''
        """
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def should_run_wealthsimple_scan(settings, now=None):
    now = now or datetime.now()
    try:
        minutes = int(settings.get("scan_interval_minutes") or 10)
    except (TypeError, ValueError):
        minutes = 10
    minutes = max(1, min(1440, minutes))

    last_scan_at = settings.get("last_scan_at")
    if not last_scan_at:
        return True
    try:
        last_dt = datetime.fromisoformat(str(last_scan_at).replace("Z", ""))
    except Exception:
        try:
            last_dt = datetime.strptime(str(last_scan_at)[:19], "%Y-%m-%d %H:%M:%S")
        except Exception:
            return True
    return (now - last_dt).total_seconds() >= minutes * 60


_wealthsimple_scheduler_started = False


def wealthsimple_folder_scheduler_loop():
    time.sleep(11)
    while True:
        try:
            now = datetime.now()
            users = get_wealthsimple_auto_import_users()
            for settings in users:
                if not should_run_wealthsimple_scan(settings, now):
                    continue
                user_id = settings["user_id"]
                try:
                    result = run_wealthsimple_folder_import_for_user(user_id, settings=settings)
                    update_wealthsimple_scan_status(
                        user_id,
                        f"OK: scanned {result['files_seen']} file(s), imported {result['files_imported']}, inserted {result['inserted']}, skipped {result['skipped']}.",
                        imported=result.get("files_imported", 0) > 0,
                    )
                except Exception as exc:
                    update_wealthsimple_scan_status(user_id, f"ERROR: {exc}", imported=False)
        except Exception as exc:
            print(f"Wealthsimple folder scheduler error: {exc}")

        time.sleep(60)


def start_wealthsimple_folder_scheduler():
    global _wealthsimple_scheduler_started
    if _wealthsimple_scheduler_started:
        return
    if os.environ.get("WEALTHSIMPLE_AUTO_IMPORT_DISABLED") == "1":
        return
    _wealthsimple_scheduler_started = True
    thread = threading.Thread(target=wealthsimple_folder_scheduler_loop, daemon=True)
    thread.start()


# =========================================================
# Main app routes
# =========================================================
@app.route("/dashboard")
def dashboard():
    if not require_login():
        return redirect(url_for("login"))

    broker, symbol, setup_tag, where_sql, params = build_filters()
    totals = get_dashboard_totals(where_sql, params)
    monthly = get_monthly_pnl(where_sql, params)
    analytics = get_analytics(where_sql, params)
    brokers = get_brokers()
    setup_tags = get_setup_tags()
    ai_review = get_ai_review(where_sql, params)
    filtered_trade_rows = get_filtered_trade_dicts(where_sql, params)
    ai_review = dict(ai_review or {})
    ai_review.update(generate_deep_ai_review(filtered_trade_rows, ai_review))
    what_if_scenarios = generate_what_if_analysis(filtered_trade_rows)
    tomorrow_plan = enhance_tomorrow_trading_plan(generate_tomorrow_plan(filtered_trade_rows), filtered_trade_rows)
    heatmap_data = get_calendar_heatmap()

    # Monthly P&L chart/table data
    months = [r["month"] or "Undated" for r in monthly]
    monthly_pnl_values = [float(r["realized_pl"] or 0) for r in monthly]
    monthly_fee_values = [float(r["fees"] or 0) for r in monthly]
    monthly_trade_counts = [int(r["trade_count"] or 0) for r in monthly]
    monthly_rows = [
        {
            "month": r["month"] or "Undated",
            "realized_pl": float(r["realized_pl"] or 0),
            "fees": float(r["fees"] or 0),
            "trade_count": int(r["trade_count"] or 0),
        }
        for r in monthly
    ]

    # Pro dashboard chart data
    equity_labels = analytics.get("equity_labels", [])
    equity_values = analytics.get("equity_curve", [])

    setup_stats = analytics.get("setup_stats", [])
    setup_labels = [s.get("setup_tag", "Unlabeled") for s in setup_stats]
    setup_values = [float(s.get("pnl") or 0) for s in setup_stats]

    risk_rows = [r for r in filtered_trade_rows if abs(float(r.get("risk_amount") or 0)) > 1e-12]
    risk_labels = [(r.get("trade_datetime") or "")[:10] or str(r.get("id", "")) for r in risk_rows]
    risk_values = [float(r.get("risk_amount") or 0) for r in risk_rows]

    # Compact calendar data for the new card grid
    calendar_data = [
        {
            "date": day[-5:],
            "full_date": day,
            "pnl": round(float(info.get("pnl") or 0), 2),
            "trades": int(info.get("trades") or 0),
        }
        for day, info in sorted(
            [(day, info) for day, info in heatmap_data.items() if day],
            key=lambda x: x[0]
        )[-35:]
    ]

    # Behavior analytics derived from filtered trade rows
    daily_counts = defaultdict(int)
    daily_pnl = defaultdict(float)
    pnl_list = []
    for r in filtered_trade_rows:
        pnl = float(r.get("realized_pl") or 0)
        pnl_list.append(pnl)
        day = (r.get("trade_datetime") or "")[:10]
        if day:
            daily_counts[day] += 1
            daily_pnl[day] += pnl

    overtrading_days = sum(1 for count in daily_counts.values() if count >= 5)
    losing_days = [pnl for pnl in daily_pnl.values() if pnl < 0]
    avg_losing_day = round(sum(losing_days) / len(losing_days), 2) if losing_days else 0.0

    streak = 0
    max_losing_streak = 0
    for pnl in pnl_list:
        if pnl < 0:
            streak += 1
            max_losing_streak = max(max_losing_streak, streak)
        else:
            streak = 0

    ai_review["overtrading_days"] = overtrading_days
    ai_review["max_losing_streak"] = max_losing_streak
    ai_review["avg_losing_day"] = fmt_num(avg_losing_day)

    return render_template(
        "dashboard.html",
        title="Dashboard",
        broker=broker,
        symbol=symbol,
        setup_tag=setup_tag,
        brokers=brokers,
        setup_tags=setup_tags,
        portfolios=get_user_portfolios(current_user_id()),
        active_portfolio_id=current_portfolio_id(),
        totals=totals,
        monthly=monthly_rows,
        monthly_labels=json.dumps(months),
        monthly_pnl_values=json.dumps(monthly_pnl_values),
        monthly_fee_values=json.dumps(monthly_fee_values),
        monthly_trade_counts=json.dumps(monthly_trade_counts),
        analytics=analytics,
        ai_review=ai_review,
        scenarios=what_if_scenarios,
        tomorrow_plan=tomorrow_plan,
        calendar=calendar_data,
        calendar_heatmap_json=json.dumps(heatmap_data),
        equity_labels=json.dumps(equity_labels),
        equity_values=json.dumps(equity_values),
        setup_labels=json.dumps(setup_labels),
        setup_values=json.dumps(setup_values),
        risk_labels=json.dumps(risk_labels),
        risk_values=json.dumps(risk_values),
    )


@app.route("/import", methods=["POST"])
def import_auto_detect():
    if not require_login():
        return redirect(url_for("login"))

    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file uploaded"}), 400

    file = request.files["file"]
    if file.filename == "" or not file.filename.lower().endswith(".csv"):
        return jsonify({"ok": False, "error": "Valid CSV required"}), 400

    save_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(save_path)

    import_type = detect_import_type(save_path)

    try:
        if import_type in {"ibkr", "ibkr_trades"}:
            batch_id = new_batch_id("IBKR")
            trades = parse_ibkr_activity_csv(save_path)
            inserted, skipped = insert_trades(trades, import_file=file.filename, batch_id=batch_id)
            return import_success_page("IBKR Import Complete", "IBKR Activity Statement", batch_id, len(trades), inserted, skipped)

        if import_type == "ibkr_summary":
            batch_id = new_batch_id("IBKRSUM")
            trades = parse_ibkr_summary_csv(save_path)
            inserted, skipped = insert_trades(trades, import_file=file.filename, batch_id=batch_id)
            return import_success_page("IBKR Summary Import Complete", "IBKR Realized & Unrealized Summary", batch_id, len(trades), inserted, skipped)

        if import_type == "wealthsimple":
            batch_id = new_batch_id("WS")
            new_trades = parse_wealthsimple_csv(save_path)
            existing_ws_trades = get_existing_wealthsimple_trades()
            updates_for_existing, new_trades_with_fifo = recompute_fifo_for_wealthsimple(existing_ws_trades, new_trades)
            update_existing_trade_fifo_values(updates_for_existing)
            inserted, skipped = insert_trades(new_trades_with_fifo, import_file=file.filename, batch_id=batch_id)
            return import_success_page("Wealthsimple Import Complete", "Wealthsimple Activity Export", batch_id, len(new_trades), inserted, skipped)

        if import_type == "performance":
            batch_id = new_batch_id("PERF")
            trades = parse_performance_csv(save_path)
            inserted, skipped = insert_trades(trades, import_file=file.filename, batch_id=batch_id)
            return import_success_page("Performance CSV Import Complete", "Performance CSV", batch_id, len(trades), inserted, skipped)

        return jsonify({
            "ok": False,
            "error": "Could not auto-detect CSV type. Supported formats: IBKR Trades, IBKR Summary, Wealthsimple, Performance CSV."
        }), 400

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/import/ibkr", methods=["POST"])
def import_ibkr():
    if not require_login():
        return redirect(url_for("login"))

    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file uploaded"}), 400

    file = request.files["file"]
    if file.filename == "" or not file.filename.lower().endswith(".csv"):
        return jsonify({"ok": False, "error": "Valid CSV required"}), 400

    save_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(save_path)
    batch_id = new_batch_id("IBKR")

    try:
        trades = parse_ibkr_activity_csv(save_path)
        inserted, skipped = insert_trades(trades, import_file=file.filename, batch_id=batch_id)

        body = f"""
<div class="glass-card">
    <h2 class="section-title">IBKR Import Complete</h2>
    <div class="grid-kpi" style="grid-template-columns: repeat(4, minmax(0, 1fr)); margin-top:14px;">
        <div class="kpi"><div class="kpi-label">Batch</div><div class="kpi-value mono" style="font-size:18px;">{batch_id}</div></div>
        <div class="kpi"><div class="kpi-label">Parsed Rows</div><div class="kpi-value">{len(trades)}</div></div>
        <div class="kpi"><div class="kpi-label">Inserted</div><div class="kpi-value pos">{inserted}</div></div>
        <div class="kpi"><div class="kpi-label">Duplicates</div><div class="kpi-value warn">{skipped}</div></div>
    </div>
    <div style="margin-top:16px; display:flex; gap:10px;">
        <a class="btn" href="/dashboard">Back to Dashboard</a>
        <a class="btn secondary" href="/imports">View Imports</a>
    </div>
</div>
"""
        return page_shell("IBKR Import", body)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/import/wealthsimple", methods=["POST"])
def import_wealthsimple():
    if not require_login():
        return redirect(url_for("login"))

    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file uploaded"}), 400

    file = request.files["file"]
    if file.filename == "" or not file.filename.lower().endswith(".csv"):
        return jsonify({"ok": False, "error": "Valid CSV required"}), 400

    save_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(save_path)
    batch_id = new_batch_id("WS")

    try:
        new_trades = parse_wealthsimple_csv(save_path)
        existing_ws_trades = get_existing_wealthsimple_trades()
        updates_for_existing, new_trades_with_fifo = recompute_fifo_for_wealthsimple(existing_ws_trades, new_trades)

        update_existing_trade_fifo_values(updates_for_existing)
        inserted, skipped = insert_trades(new_trades_with_fifo, import_file=file.filename, batch_id=batch_id)

        body = f"""
<div class="glass-card">
    <h2 class="section-title">Wealthsimple Import Complete</h2>
    <div class="grid-kpi" style="grid-template-columns: repeat(4, minmax(0, 1fr)); margin-top:14px;">
        <div class="kpi"><div class="kpi-label">Batch</div><div class="kpi-value mono" style="font-size:18px;">{batch_id}</div></div>
        <div class="kpi"><div class="kpi-label">Parsed Rows</div><div class="kpi-value">{len(new_trades)}</div></div>
        <div class="kpi"><div class="kpi-label">Inserted</div><div class="kpi-value pos">{inserted}</div></div>
        <div class="kpi"><div class="kpi-label">Duplicates</div><div class="kpi-value warn">{skipped}</div></div>
    </div>
    <div style="margin-top:16px; display:flex; gap:10px;">
        <a class="btn" href="/dashboard">Back to Dashboard</a>
        <a class="btn secondary" href="/imports">View Imports</a>
    </div>
</div>
"""
        return page_shell("Wealthsimple Import", body)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/import/performance", methods=["POST"])
def import_performance():
    if not require_login():
        return redirect(url_for("login"))

    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No file uploaded"}), 400

    file = request.files["file"]
    if file.filename == "" or not file.filename.lower().endswith(".csv"):
        return jsonify({"ok": False, "error": "Valid CSV required"}), 400

    save_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(save_path)
    batch_id = new_batch_id("PERF")

    try:
        trades = parse_performance_csv(save_path)
        inserted, skipped = insert_trades(trades, import_file=file.filename, batch_id=batch_id)

        body = f"""
<div class="glass-card">
    <h2 class="section-title">Performance Import Complete</h2>
    <div class="grid-kpi" style="grid-template-columns: repeat(4, minmax(0, 1fr)); margin-top:14px;">
        <div class="kpi"><div class="kpi-label">Batch</div><div class="kpi-value mono" style="font-size:18px;">{batch_id}</div></div>
        <div class="kpi"><div class="kpi-label">Parsed Rows</div><div class="kpi-value">{len(trades)}</div></div>
        <div class="kpi"><div class="kpi-label">Inserted</div><div class="kpi-value pos">{inserted}</div></div>
        <div class="kpi"><div class="kpi-label">Duplicates</div><div class="kpi-value warn">{skipped}</div></div>
    </div>
    <div style="margin-top:16px; display:flex; gap:10px;">
        <a class="btn" href="/dashboard">Back to Dashboard</a>
        <a class="btn secondary" href="/imports">View Imports</a>
    </div>
</div>
"""
        return page_shell("Performance Import", body)
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/imports")
def imports_page():
    if not require_login():
        return redirect(url_for("login"))

    rows = get_import_history()

    body = f"""
<div class="glass-card">
    <div class="section-head">
        <div>
            <h2 class="section-title">Import History</h2>
            <div class="section-note">Track all CSV batches and remove bad loads safely.</div>
        </div>
    </div>

    <div class="table-wrap">
        <table>
            <thead>
                <tr>
                    <th>Imported At</th>
                    <th>Batch ID</th>
                    <th>Broker</th>
                    <th>File</th>
                    <th>Rows</th>
                    <th>Realized P&amp;L</th>
                    <th>Fees</th>
                    <th>Action</th>
                </tr>
            </thead>
            <tbody>
                {''.join([
                    f'''
                    <tr>
                        <td>{r["imported_at"] or ""}</td>
                        <td class="mono">{r["batch_id"]}</td>
                        <td><span class="tag broker">{r["broker"]}</span></td>
                        <td>{r["import_file"] or ""}</td>
                        <td>{r["row_count"]}</td>
                        <td class="{"pos" if float(r["realized_pl"] or 0) >= 0 else "neg"}">{fmt_num(r["realized_pl"])}</td>
                        <td>{fmt_num(r["fees"])}</td>
                        <td>
                            <form method="post" action="/imports/delete/{r["batch_id"]}" onsubmit="return confirm('Delete this imported batch?');">
                                <button class="btn danger" type="submit">Delete Batch</button>
                            </form>
                        </td>
                    </tr>
                    '''
                    for r in rows
                ]) or '<tr><td colspan="8"><div class="empty">No imports yet.</div></td></tr>'}
            </tbody>
        </table>
    </div>
</div>
"""
    return page_shell("Imports", body)


@app.route("/imports/delete/<batch_id>", methods=["POST"])
def imports_delete(batch_id):
    if not require_login():
        return redirect(url_for("login"))

    deleted = delete_batch(batch_id)
    body = f"""
<div class="glass-card">
    <h2 class="section-title">Batch Deleted</h2>
    <div class="grid-kpi" style="grid-template-columns: repeat(2, minmax(0, 1fr)); margin-top:14px;">
        <div class="kpi"><div class="kpi-label">Batch ID</div><div class="kpi-value mono" style="font-size:18px;">{batch_id}</div></div>
        <div class="kpi"><div class="kpi-label">Rows Deleted</div><div class="kpi-value">{deleted}</div></div>
    </div>
    <div style="margin-top:16px;">
        <a class="btn" href="/imports">Back to Imports</a>
    </div>
</div>
"""
    return page_shell("Imports", body)


@app.route("/trades")
def trades_page():
    if not require_login():
        return redirect(url_for("login"))

    broker, symbol, setup_tag, where_sql, params = build_filters()
    brokers = get_brokers()
    setup_tags = get_setup_tags()
    rows = get_recent_trades(where_sql, params, limit=1000)

    body = f"""
<div class="glass-card">
    <div class="section-head">
        <div>
            <h2 class="section-title">Trade Explorer</h2>
            <div class="section-note">Search, filter, and review execution-level history.</div>
        </div>
    </div>

    <form method="get" action="/trades">
        <div class="filters">
            <select name="broker">
                <option value="">All Brokers</option>
                {''.join([f'<option value="{b}" {"selected" if b == broker else ""}>{b}</option>' for b in brokers])}
            </select>

            <input type="text" name="symbol" placeholder="Search symbol" value="{symbol}">

            <select name="setup_tag">
                <option value="">All Setup Tags</option>
                {''.join([f'<option value="{t}" {"selected" if t == setup_tag else ""}>{t}</option>' for t in setup_tags])}
            </select>

            <button class="btn" type="submit">Apply Filters</button>
            <a class="btn secondary" href="/trades">Clear</a>
        </div>
    </form>
</div>

<div class="glass-card">
    <div class="table-wrap trades-fit-wrap">
        <table class="trades-fit-table" data-no-column-reorder="1">
            <thead>
                <tr>
                    <th>Broker</th>
                    <th>Symbol</th>
                    <th>Date</th>
                    <th>Qty</th>
                    <th>Side</th>
                    <th>Buy</th>
                    <th>Sell</th>
                    <th>P&amp;L</th>
                    <th>Setup</th>
                    <th>Note</th>
                    <th></th>
                </tr>
            </thead>
            <tbody>
                {''.join([
                    f'''
                    <tr>
                        <td><span class="tag broker">{r["broker"]}</span></td>
                        <td class="mono symbol-cell">{r["symbol"]}</td>
                        <td class="date-cell">{fmt_dt(r["trade_datetime"])}</td>
                        <td class="num-cell">{fmt_num(r["quantity"])}</td>
                        <td><span class="tag {"buy" if (r["side"] or "") == "BUY" else "sell"}">{r["side"] or ""}</span></td>
                        <td class="num-cell">{fmt_num(r["buy_price"])}</td>
                        <td class="num-cell">{fmt_num(r["sell_price"])}</td>
                        <td class="num-cell {"pos" if float(r["realized_pl"] or 0) >= 0 else "neg"}">{fmt_num(r["realized_pl"])}</td>
                        <td class="small-text-cell">{r["setup_tag"] or ""}</td>
                        <td class="note-cell"><input type="text" value="{(r["journal_note"] or "").replace('"', '&quot;')}" onchange="saveNote({r["id"]}, this.value)" title="{(r["journal_note"] or "").replace('"', '&quot;')}" ></td>
                        <td class="open-cell"><a class="btn secondary compact-open" href="/trade/{r["id"]}">Open</a></td>
                    </tr>
                    '''
                    for r in rows
                ]) or '<tr><td colspan="11"><div class="empty">No trades found.</div></td></tr>'}
            </tbody>
        </table>
    </div>
</div>
"""
    extra_head = """
<style>
/* Compact no-side-scroll Trades table */
.trades-fit-wrap {
    overflow-x: hidden !important;
}
.trades-fit-table {
    width: 100% !important;
    min-width: 0 !important;
    max-width: 100% !important;
    table-layout: fixed !important;
}
.trades-fit-table th,
.trades-fit-table td {
    padding: 8px 7px !important;
    font-size: 12px !important;
    line-height: 1.25 !important;
    white-space: normal !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
    vertical-align: middle !important;
}
.trades-fit-table th:nth-child(1), .trades-fit-table td:nth-child(1) { width: 11%; }
.trades-fit-table th:nth-child(2), .trades-fit-table td:nth-child(2) { width: 11%; }
.trades-fit-table th:nth-child(3), .trades-fit-table td:nth-child(3) { width: 13%; }
.trades-fit-table th:nth-child(4), .trades-fit-table td:nth-child(4) { width: 7%; }
.trades-fit-table th:nth-child(5), .trades-fit-table td:nth-child(5) { width: 7%; }
.trades-fit-table th:nth-child(6), .trades-fit-table td:nth-child(6) { width: 8%; }
.trades-fit-table th:nth-child(7), .trades-fit-table td:nth-child(7) { width: 8%; }
.trades-fit-table th:nth-child(8), .trades-fit-table td:nth-child(8) { width: 9%; }
.trades-fit-table th:nth-child(9), .trades-fit-table td:nth-child(9) { width: 10%; }
.trades-fit-table th:nth-child(10), .trades-fit-table td:nth-child(10) { width: 11%; }
.trades-fit-table th:nth-child(11), .trades-fit-table td:nth-child(11) { width: 5%; }
.trades-fit-table .tag {
    max-width: 100%;
    padding: 5px 7px !important;
    font-size: 11px !important;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}
.trades-fit-table .date-cell {
    white-space: normal !important;
    overflow-wrap: anywhere;
}
.trades-fit-table .symbol-cell,
.trades-fit-table .small-text-cell {
    overflow-wrap: anywhere;
}
.trades-fit-table .num-cell {
    text-align: right;
    white-space: nowrap !important;
}
.trades-fit-table .note-cell input {
    width: 100%;
    min-width: 0;
    padding: 7px 8px;
    font-size: 12px;
    border-radius: 10px;
}
.trades-fit-table .compact-open {
    min-height: 30px !important;
    padding: 6px 8px !important;
    border-radius: 10px !important;
    font-size: 12px !important;
}
@media (max-width: 900px) {
    .trades-fit-table th,
    .trades-fit-table td { padding: 6px 5px !important; font-size: 11px !important; }
    .trades-fit-table .tag { font-size: 10px !important; padding: 4px 5px !important; }
    .trades-fit-table th:nth-child(1), .trades-fit-table td:nth-child(1) { display: none; }
    .trades-fit-table th:nth-child(9), .trades-fit-table td:nth-child(9) { display: none; }
}
</style>
<script>
async function saveNote(id, note) {
    try {
        await fetch('/api/save_note', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({id, note})
        });
    } catch (err) {
        console.error(err);
    }
}
</script>
"""
    return page_shell("Trades", body, extra_head=extra_head)


@app.route("/trade/<int:trade_id>")
def trade_detail(trade_id):
    if not require_login():
        return redirect(url_for("login"))

    trade = get_trade_by_id(trade_id)
    if not trade:
        return page_shell("Trade Detail", '<div class="glass-card"><div class="empty">Trade not found.</div></div>')

    screenshot_html = ""
    if trade["screenshot_filename"]:
        screenshot_html = f"""
        <div style="margin-top:12px;">
            <div class="section-note" style="margin-bottom:8px;">Screenshot</div>
            <img class="thumb" src="/screenshots/{trade["screenshot_filename"]}" alt="screenshot">
        </div>
        """

    body = f"""
<div class="glass-card">
    <div class="section-head">
        <div>
            <h2 class="section-title">Trade #{trade["id"]} · {trade["symbol"]}</h2>
            <div class="section-note">{fmt_dt(trade["trade_datetime"])} · {trade["broker"]}</div>
        </div>
        <a class="btn secondary" href="/trades">Back to Trades</a>
    </div>

    <div class="grid-kpi" style="grid-template-columns: repeat(5, minmax(0, 1fr));">
        <div class="kpi"><div class="kpi-label">Qty</div><div class="kpi-value">{fmt_num(trade["quantity"])}</div></div>
        <div class="kpi"><div class="kpi-label">Side</div><div class="kpi-value">{trade["side"] or ""}</div></div>
        <div class="kpi"><div class="kpi-label">Trade Price</div><div class="kpi-value">{fmt_num(trade["trade_price"])}</div></div>
        <div class="kpi"><div class="kpi-label">Buy Price</div><div class="kpi-value">{fmt_num(trade["buy_price"])}</div></div>
        <div class="kpi"><div class="kpi-label">Sell Price</div><div class="kpi-value">{fmt_num(trade["sell_price"])}</div></div>
        <div class="kpi"><div class="kpi-label">Basis</div><div class="kpi-value">{fmt_num(trade["basis"])}</div></div>
        <div class="kpi"><div class="kpi-label">Proceeds</div><div class="kpi-value">{fmt_num(trade["proceeds"])}</div></div>
        <div class="kpi"><div class="kpi-label">P&amp;L</div><div class="kpi-value {'pos' if float(trade["realized_pl"] or 0) >= 0 else 'neg'}">{fmt_num(trade["realized_pl"])}</div></div>
        <div class="kpi"><div class="kpi-label">Risk</div><div class="kpi-value">{fmt_num(trade["risk_amount"])}</div></div>
        <div class="kpi"><div class="kpi-label">R Multiple</div><div class="kpi-value {'pos' if float(trade["r_multiple"] or 0) >= 0 else 'neg'}">{fmt_num(trade["r_multiple"])}</div></div>
    </div>
</div>

<div class="split-2">
    <div class="glass-card">
        <div class="section-head">
            <h2 class="section-title">Journal Entry</h2>
            <div class="section-note">Tag, note, and screenshot</div>
        </div>

        <form method="post" action="/trade/{trade["id"]}/journal" enctype="multipart/form-data">
            <div class="filters" style="margin-bottom:12px;">
                <input type="text" name="setup_tag" placeholder="Setup tag" value="{trade["setup_tag"] or ""}">
                <input type="number" step="0.01" name="risk_amount" placeholder="Risk ($)" value="{trade["risk_amount"] or 0}">
            </div>

            <div class="filters" style="margin-bottom:12px;">
                <select name="mistake_tag">
                    <option value="" {"selected" if not (trade["mistake_tag"] or "").strip() else ""}>No Mistake</option>
                    <option {"selected" if (trade["mistake_tag"] or "") == "FOMO" else ""}>FOMO</option>
                    <option {"selected" if (trade["mistake_tag"] or "") == "Overtrading" else ""}>Overtrading</option>
                    <option {"selected" if (trade["mistake_tag"] or "") == "Late Entry" else ""}>Late Entry</option>
                    <option {"selected" if (trade["mistake_tag"] or "") == "No Stop Loss" else ""}>No Stop Loss</option>
                    <option {"selected" if (trade["mistake_tag"] or "") == "Revenge Trade" else ""}>Revenge Trade</option>
                    <option {"selected" if (trade["mistake_tag"] or "") == "Oversized" else ""}>Oversized</option>
                </select>
            </div>

            <textarea name="note" placeholder="Trade notes...">{trade["journal_note"] or ""}</textarea>

            <div class="filters" style="margin-top:12px;">
                <input type="file" name="screenshot" accept=".png,.jpg,.jpeg,.webp">
                <button class="btn" type="submit">Save Journal</button>
            </div>
        </form>

        {screenshot_html}
    </div>

    <div class="glass-card">
        <div class="section-head">
            <h2 class="section-title">Trade Meta</h2>
            <div class="section-note">Imported fields</div>
        </div>

        <div class="table-wrap">
            <table>
                <tbody>
                    <tr><td>Broker</td><td>{trade["broker"]}</td></tr>
                    <tr><td>Symbol</td><td class="mono">{trade["symbol"]}</td></tr>
                    <tr><td>Asset</td><td>{trade["asset_category"] or ""}</td></tr>
                    <tr><td>Currency</td><td>{trade["currency"] or ""}</td></tr>
                    <tr><td>Batch</td><td class="mono">{trade["batch_id"] or ""}</td></tr>
                    <tr><td>Mistake Tag</td><td>{trade["mistake_tag"] or "No Mistake"}</td></tr>
                    <tr><td>Code</td><td>{trade["code"] or ""}</td></tr>
                    <tr><td>Import File</td><td>{trade["import_file"] or ""}</td></tr>
                </tbody>
            </table>
        </div>
    </div>
</div>
"""
    return page_shell("Trade Detail", body)


@app.route("/trade/<int:trade_id>/journal", methods=["POST"])
def trade_journal_save(trade_id):
    if not require_login():
        return redirect(url_for("login"))

    trade = get_trade_by_id(trade_id)
    if not trade:
        return redirect(url_for("trades_page"))

    setup_tag = request.form.get("setup_tag", "").strip()
    mistake_tag = request.form.get("mistake_tag", "").strip()
    note = request.form.get("note", "").strip()
    risk_amount = float(request.form.get("risk_amount") or 0)
    realized_pl = float(trade["realized_pl"] or 0)
    r_multiple = round((realized_pl / risk_amount), 2) if risk_amount > 0 else 0.0
    screenshot_filename = None

    if "screenshot" in request.files:
        file = request.files["screenshot"]
        if file and file.filename and allowed_file(file.filename, {"png", "jpg", "jpeg", "webp"}):
            ext = file.filename.rsplit(".", 1)[1].lower()
            screenshot_filename = f"{trade['symbol']}_{trade_id}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.{ext}"
            file.save(os.path.join(SCREENSHOT_FOLDER, screenshot_filename))

    upsert_trade_journal(
        trade_id,
        setup_tag=setup_tag,
        mistake_tag=mistake_tag,
        note=note,
        screenshot_filename=screenshot_filename
    )

    conn = get_db_connection()
    conn.execute(
        "UPDATE trades SET risk_amount = ?, r_multiple = ? WHERE id = ? AND user_id = ? AND portfolio_id = ?",
        (risk_amount, r_multiple, trade_id, current_user_id(), current_portfolio_id())
    )
    conn.commit()
    conn.close()

    return redirect(url_for("trade_detail", trade_id=trade_id))



@app.route("/api/day-trades")
def api_day_trades():
    if not require_login():
        return jsonify({"ok": False, "error": "Unauthorized"}), 401

    date_key = request.args.get("date", "").strip()
    if not re.match(r"^\d{4}-\d{2}-\d{2}$", date_key):
        return jsonify({"ok": False, "error": "Valid date required"}), 400

    conn = get_db_connection()
    rows = conn.execute("""
        SELECT
            t.id,
            t.broker,
            t.symbol,
            t.trade_datetime,
            t.quantity,
            t.side,
            t.trade_price,
            t.buy_price,
            t.sell_price,
            t.realized_pl,
            j.setup_tag,
            j.mistake_tag,
            j.note AS journal_note
        FROM trades t
        LEFT JOIN trade_journal j ON j.trade_id = t.id
        WHERE t.user_id = ?
          AND t.portfolio_id = ?
          AND substr(t.trade_datetime, 1, 10) = ?
        ORDER BY t.trade_datetime ASC, t.id ASC
    """, (current_user_id(), current_portfolio_id(), date_key)).fetchall()
    conn.close()

    trades = []
    total_pnl = 0.0
    for r in rows:
        pnl = float(r["realized_pl"] or 0)
        total_pnl += pnl
        trades.append({
            "id": r["id"],
            "broker": r["broker"],
            "symbol": r["symbol"],
            "time": fmt_dt(r["trade_datetime"]),
            "quantity": float(r["quantity"] or 0),
            "side": r["side"] or "",
            "trade_price": float(r["trade_price"] or 0),
            "buy_price": float(r["buy_price"] or 0),
            "sell_price": float(r["sell_price"] or 0),
            "realized_pl": pnl,
            "setup_tag": r["setup_tag"] or "",
            "mistake_tag": r["mistake_tag"] or "",
            "journal_note": r["journal_note"] or "",
        })

    return jsonify({
        "ok": True,
        "date": date_key,
        "trade_count": len(trades),
        "realized_pl": round(total_pnl, 2),
        "trades": trades,
    })

@app.route("/api/save_note", methods=["POST"])
def api_save_note():
    if not require_login():
        return jsonify({"ok": False, "error": "Unauthorized"}), 401

    payload = request.get_json(silent=True) or {}
    trade_id = int(payload.get("id") or 0)
    note = str(payload.get("note") or "").strip()

    trade = get_trade_by_id(trade_id)
    if not trade:
        return jsonify({"ok": False, "error": "Trade not found"}), 404

    upsert_trade_journal(
        trade_id,
        setup_tag=trade["setup_tag"] or "",
        mistake_tag=trade["mistake_tag"] or "",
        note=note,
        screenshot_filename=None
    )
    return jsonify({"ok": True})


@app.route("/gallery")
def gallery():
    if not require_login():
        return redirect(url_for("login"))

    rows = get_gallery_rows()

    cards = "".join([
        f'''
        <div class="glass-card soft">
            <div class="section-head">
                <div>
                    <h2 class="section-title">{r["symbol"]}</h2>
                    <div class="section-note">{fmt_dt(r["trade_datetime"])} · {r["broker"]}</div>
                </div>
                <div class="tag {"buy" if float(r["realized_pl"] or 0) >= 0 else "sell"}">{fmt_num(r["realized_pl"])}</div>
            </div>
            <img class="thumb" src="/screenshots/{r["screenshot_filename"]}" alt="screenshot" style="max-width:100%; width:100%;">
            <div style="margin-top:12px;">
                <div class="section-note">Setup: {r["setup_tag"] or "Unlabeled"}</div>
                <div style="margin-top:10px;">
                    <a class="btn secondary" href="/trade/{r["id"]}">Open Trade</a>
                </div>
            </div>
        </div>
        '''
        for r in rows
    ])

    body = f"""
<div class="glass-card">
    <div class="section-head">
        <div>
            <h2 class="section-title">Screenshot Gallery</h2>
            <div class="section-note">All saved trade screenshots in one place.</div>
        </div>
    </div>
</div>

<div class="split-3">
    {cards or '<div class="empty">No screenshots uploaded yet.</div>'}
</div>
"""
    return page_shell("Gallery", body)


@app.route("/screenshots/<path:filename>")
def screenshots(filename):
    return send_from_directory(SCREENSHOT_FOLDER, filename)


@app.route("/options")
def options_page():
    if not require_login():
        return redirect(url_for("login"))

    broker, symbol, setup_tag, where_sql, params = build_filters()
    brokers = get_brokers()
    setup_tags = get_setup_tags()
    groups = get_option_strategy_groups(where_sql, params)

    content = ""
    for g in groups:
        legs_html = "".join([
            f"""
            <tr>
                <td class="mono">{leg["symbol"]}</td>
                <td>{fmt_dt(leg["trade_datetime"])}</td>
                <td>{fmt_num(leg["quantity"])}</td>
                <td><span class="tag {"buy" if (leg["side"] or "") == "BUY" else "sell"}">{leg["side"] or ""}</span></td>
                <td>{fmt_num(leg["buy_price"])}</td>
                <td>{fmt_num(leg["sell_price"])}</td>
                <td>{fmt_num(leg["proceeds"])}</td>
                <td>{fmt_num(leg["commission"])}</td>
                <td class="{"pos" if float(leg["realized_pl"] or 0) >= 0 else "neg"}">{fmt_num(leg["realized_pl"])}</td>
            </tr>
            """
            for leg in g["legs"]
        ])

        content += f"""
<div class="glass-card">
    <div class="section-head">
        <div>
            <h2 class="section-title">{g["underlying"]} · {g["strategy_type"]}</h2>
            <div class="section-note">{g["trade_date"]} · {g["expiration"]} · {g["option_type"]} · Strikes: {g["strikes"]}</div>
        </div>
        <div class="tag broker">{g["broker"]}</div>
    </div>

    <div class="grid-kpi" style="grid-template-columns: repeat(4, minmax(0, 1fr));">
        <div class="kpi"><div class="kpi-label">Legs</div><div class="kpi-value">{g["leg_count"]}</div></div>
        <div class="kpi"><div class="kpi-label">Net Proceeds</div><div class="kpi-value">{fmt_num(g["net_proceeds"])}</div></div>
        <div class="kpi"><div class="kpi-label">Fees</div><div class="kpi-value warn">{fmt_num(g["fees"])}</div></div>
        <div class="kpi"><div class="kpi-label">Realized P&amp;L</div><div class="kpi-value {'pos' if float(g["realized_pl"] or 0) >= 0 else 'neg'}">{fmt_num(g["realized_pl"])}</div></div>
    </div>

    <div class="table-wrap" style="margin-top:14px;">
        <table>
            <thead>
                <tr>
                    <th>Symbol</th>
                    <th>Date/Time</th>
                    <th>Qty</th>
                    <th>Side</th>
                    <th>Buy Price</th>
                    <th>Sell Price</th>
                    <th>Proceeds</th>
                    <th>Commission</th>
                    <th>Realized P&amp;L</th>
                </tr>
            </thead>
            <tbody>{legs_html}</tbody>
        </table>
    </div>
</div>
"""

    body = f"""
<div class="glass-card">
    <div class="section-head">
        <div>
            <h2 class="section-title">Strategy Review</h2>
            <div class="section-note">Approximate grouping for option structures by day and expiry.</div>
        </div>
    </div>

    <form method="get" action="/options">
        <div class="filters">
            <select name="broker">
                <option value="">All Brokers</option>
                {''.join([f'<option value="{b}" {"selected" if b == broker else ""}>{b}</option>' for b in brokers])}
            </select>

            <input type="text" name="symbol" placeholder="Search symbol / underlying" value="{symbol}">

            <select name="setup_tag">
                <option value="">All Setup Tags</option>
                {''.join([f'<option value="{t}" {"selected" if t == setup_tag else ""}>{t}</option>' for t in setup_tags])}
            </select>

            <button class="btn" type="submit">Apply Filters</button>
            <a class="btn secondary" href="/options">Clear</a>
        </div>
    </form>
</div>

{content or '<div class="glass-card"><div class="empty">No option groups found.</div></div>'}
"""
    return page_shell("Option Strategies", body)


@app.route("/export/trades")
def export_trades():
    if not require_login():
        return redirect(url_for("login"))

    broker, symbol, setup_tag, where_sql, params = build_filters()
    rows = get_recent_trades(where_sql, params, limit=100000)

    output = io.StringIO()
    writer = csv.writer(output)

    writer.writerow([
        "id", "broker", "asset_category", "currency", "symbol", "trade_datetime",
        "quantity", "side", "trade_price", "close_price", "buy_price", "sell_price", "proceeds", "commission",
        "basis", "realized_pl", "mtm_pl", "risk_amount", "r_multiple", "code", "import_file", "batch_id",
        "setup_tag", "mistake_tag", "journal_note", "screenshot_filename", "created_at"
    ])

    for r in rows:
        writer.writerow([
            r["id"], r["broker"], r["asset_category"], r["currency"], r["symbol"], r["trade_datetime"],
            r["quantity"], r["side"], r["trade_price"], r["close_price"], r["buy_price"], r["sell_price"], r["proceeds"], r["commission"],
            r["basis"], r["realized_pl"], r["mtm_pl"], r["risk_amount"], r["r_multiple"], r["code"], r["import_file"], r["batch_id"],
            r["setup_tag"], r["mistake_tag"], r["journal_note"], r["screenshot_filename"], r["created_at"]
        ])

    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment; filename=trades_export.csv"}
    )


@app.route("/export/trades.xlsx")
def export_trades_xlsx():
    if not require_login():
        return redirect(url_for("login"))

    broker, symbol, setup_tag, where_sql, params = build_filters()
    rows = get_recent_trades(where_sql, params, limit=100000)

    wb = Workbook()
    ws = wb.active
    ws.title = "Trades"

    headers = [
        "ID", "Broker", "Asset", "Currency", "Symbol", "Trade Datetime",
        "Quantity", "Side", "Trade Price", "Close Price", "Buy Price", "Sell Price", "Proceeds",
        "Commission", "Basis", "Realized P&L", "MTM P&L", "Risk Amount", "R Multiple", "Code",
        "Import File", "Batch ID", "Setup Tag", "Mistake Tag", "Journal Note"
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for r in rows:
        ws.append([
            r["id"], r["broker"], r["asset_category"], r["currency"], r["symbol"],
            r["trade_datetime"], r["quantity"], r["side"], r["trade_price"],
            r["close_price"], r["buy_price"], r["sell_price"], r["proceeds"], r["commission"], r["basis"],
            r["realized_pl"], r["mtm_pl"], r["risk_amount"], r["r_multiple"], r["code"], r["import_file"],
            r["batch_id"], r["setup_tag"], r["mistake_tag"], r["journal_note"]
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return Response(
        output.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=trades_export.xlsx"}
    )


# =========================================================
# App bootstrap
# =========================================================
@app.before_request
def enforce_auth():
    public_paths = {"/", "/login", "/register"}
    if request.path.startswith("/screenshots/"):
        return
    if request.path in public_paths:
        return
    if request.path.startswith("/static/"):
        return
    if not require_login():
        return redirect(url_for("login"))

    # Make sure every authenticated session has a valid, user-owned portfolio.
    current_portfolio_id()


if __name__ == "__main__":
    init_db()
    start_ibkr_auto_import_scheduler()
    start_ninjatrader_folder_scheduler()
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 10000)),
        debug=os.environ.get("FLASK_DEBUG") == "1",
        use_reloader=False,
        threaded=False,
    )
