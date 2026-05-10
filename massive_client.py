import os
import time
from typing import Optional

import pandas as pd
import requests


class MassiveClient:
    def __init__(self, api_key: Optional[str] = None, base_url: str = "https://api.massive.com"):
        self.api_key = api_key or os.getenv("MASSIVE_API_KEY")
        if not self.api_key:
            raise ValueError("Missing MASSIVE_API_KEY. Set it in your .env file or environment.")
        self.base_url = base_url.rstrip("/")

    def get_aggregates(
        self,
        ticker: str,
        multiplier: int,
        timespan: str,
        from_date: str,
        to_date: str,
        adjusted: bool = True,
        sort: str = "asc",
        limit: int = 50000,
    ) -> pd.DataFrame:
        """Fetch aggregate OHLCV bars for an option ticker."""
        url = f"{self.base_url}/v2/aggs/ticker/{ticker}/range/{multiplier}/{timespan}/{from_date}/{to_date}"
        params = {
            "adjusted": str(adjusted).lower(),
            "sort": sort,
            "limit": limit,
            "apiKey": self.api_key,
        }

        rows = []
        next_url = url
        while next_url:
            response = requests.get(next_url, params=params if next_url == url else {"apiKey": self.api_key}, timeout=30)
            response.raise_for_status()
            payload = response.json()
            rows.extend(payload.get("results", []))
            next_url = payload.get("next_url")
            params = None
            if next_url:
                time.sleep(0.25)

        if not rows:
            return pd.DataFrame(columns=["timestamp", "open", "high", "low", "close", "volume", "vwap", "transactions"])

        df = pd.DataFrame(rows)
        df = df.rename(columns={"t": "timestamp", "o": "open", "h": "high", "l": "low", "c": "close", "v": "volume", "vw": "vwap", "n": "transactions"})
        df["timestamp"] = pd.to_datetime(df["timestamp"], unit="ms", utc=True).dt.tz_convert("America/New_York")
        keep_cols = [col for col in ["timestamp", "open", "high", "low", "close", "volume", "vwap", "transactions"] if col in df.columns]
        return df[keep_cols].sort_values("timestamp").reset_index(drop=True)
