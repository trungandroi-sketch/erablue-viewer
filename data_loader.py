"""
data_loader.py – Google Sheets connector for Erablue Viewer
Reads directly from the public Google Sheet using the gviz CSV endpoint.
No API key required. Data is cached for 5 minutes (configurable).
"""
import io
import ssl
import urllib.parse
import urllib.request

import pandas as pd
import streamlit as st

# ─── Configuration ────────────────────────────────────────────────────────────
SHEET_ID = "17FvQ8YaVVV4U158yhGM8hWFNDfI6czuL1wPCUDMteOU"
CACHE_TTL = 300  # seconds (5 minutes)

SHEET_TABS = [
    "Erablue Existing",
    "Sellout Fixture",
    "Banner",
    "2 lantai",
    "Reklame store",
    "Fixture principle",
]

# SSL context – bypasses macOS cert-bundle issue
_SSL = ssl.create_default_context()
_SSL.check_hostname = False
_SSL.verify_mode = ssl.CERT_NONE


# ─── Internal helpers ─────────────────────────────────────────────────────────
def _gviz_url(sheet_name: str) -> str:
    enc = urllib.parse.quote(sheet_name)
    return (
        f"https://docs.google.com/spreadsheets/d/{SHEET_ID}"
        f"/gviz/tq?tqx=out:csv&sheet={enc}"
    )


def _fetch(sheet_name: str) -> str:
    url = _gviz_url(sheet_name)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=15, context=_SSL) as r:
        return r.read().decode("utf-8")


# ─── Public API ───────────────────────────────────────────────────────────────
@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def load_sheet(sheet_name: str) -> pd.DataFrame:
    """Load one sheet; cached for CACHE_TTL seconds. Normalizes column names to be unique."""
    raw = _fetch(sheet_name)
    df = pd.read_csv(io.StringIO(raw), header=0, dtype=str)
    df = df.dropna(how="all")
    # Normalize column names: collapse spaces and ensure uniqueness
    seen = {}
    new_cols = []
    for c in df.columns:
        normalized = " ".join(str(c).split())
        if normalized in seen:
            seen[normalized] += 1
            new_cols.append(f"{normalized}.{seen[normalized]}")
        else:
            seen[normalized] = 0
            new_cols.append(normalized)
    df.columns = new_cols
    return df


@st.cache_data(ttl=CACHE_TTL, show_spinner=False)
def load_erablue() -> pd.DataFrame:
    """
    Load 'Erablue Existing' and coerce numeric columns.
    Drops trailing summary rows (rows without a valid ID).
    """
    df = load_sheet("Erablue Existing")
    # Drop rows with no ID (summary rows at bottom)
    id_col = next((c for c in df.columns if "ID" in c and "Cửa hàng" in c), None)
    if id_col:
        df = df[df[id_col].notna() & (df[id_col].astype(str).str.strip() != "")]
    # Drop fully unnamed/empty columns
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
    df = df.loc[:, df.columns.str.strip() != "."]
    # Normalize location column values casing (e.g. Jawa Barat vs Jawa barat)
    for loc_col in ["Khu vực", "Tỉnh/Thành phố", "Tỉnh/Thành phố (Rút gọn)"]:
        if loc_col in df.columns:
            df[loc_col] = df[loc_col].apply(lambda x: str(x).strip().title() if pd.notna(x) and str(x).strip() != "" else x)
            
    # Coerce numeric columns (>30% parseable as float → convert)
    for col in df.columns:
        series = df[col]
        if hasattr(series, "dtype") and pd.api.types.is_object_dtype(series):
            converted = pd.to_numeric(series, errors="coerce")
            if converted.notna().sum() > len(df) * 0.3:
                df[col] = converted
    return df


def refresh():
    """Clear all cached data – next access will re-fetch from Google Sheets."""
    st.cache_data.clear()


def last_refresh_label() -> str:
    import datetime
    now = datetime.datetime.now()
    return now.strftime("Cập nhật lúc %H:%M ngày %d/%m/%Y")
