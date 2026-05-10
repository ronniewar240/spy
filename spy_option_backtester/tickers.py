from datetime import datetime


def build_option_ticker(underlying: str, expiry: str, right: str, strike: float) -> str:
    """Build OCC-style option ticker used by Massive/Polygon.

    Example: SPY, 2026-05-08, C, 735 -> O:SPY260508C00735000
    """
    underlying = underlying.upper().strip()
    right = right.upper().strip()
    if right not in {"C", "P"}:
        raise ValueError("right must be 'C' for call or 'P' for put")

    exp = datetime.strptime(expiry, "%Y-%m-%d")
    yymmdd = exp.strftime("%y%m%d")
    strike_int = int(round(float(strike) * 1000))
    strike_code = f"{strike_int:08d}"
    return f"O:{underlying}{yymmdd}{right}{strike_code}"
