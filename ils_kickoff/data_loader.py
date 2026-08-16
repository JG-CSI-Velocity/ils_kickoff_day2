"""CSV/Excel data loading, cleaning, and validation."""

import logging
import warnings

import pandas as pd

from ils_kickoff.column_map import REQUIRED_COLUMNS, resolve_columns
from ils_kickoff.exceptions import ColumnMismatchError, DataLoadError
from ils_kickoff.settings import Settings

logger = logging.getLogger(__name__)


def load_data(settings: Settings) -> pd.DataFrame:
    """Load, clean, and validate client data from a CSV or Excel file.

    Returns a validated DataFrame with canonical column names.
    """
    path = settings.data_file
    logger.info("Loading data from %s", path)

    try:
        if path.suffix.lower() == ".csv":
            df = pd.read_csv(path, encoding="utf-8-sig", low_memory=False)
        else:
            df = pd.read_excel(path)
    except Exception as e:
        raise DataLoadError(f"Could not read '{path.name}': {e}") from e

    logger.info("Loaded %d rows x %d columns", len(df), len(df.columns))

    # Resolve column aliases to canonical names
    df = resolve_columns(df)

    # Validate required columns
    available = set(df.columns)
    missing = REQUIRED_COLUMNS - available
    if missing:
        raise ColumnMismatchError(missing=missing, available=available)

    # Clean data
    df = _clean_deposit_count(df)
    df = _clean_numeric_columns(df)
    df = _parse_dates(df)

    logger.info("Data ready: %d rows, columns: %s", len(df), list(df.columns))
    return df


def _clean_deposit_count(df: pd.DataFrame) -> pd.DataFrame:
    """Extract first value from tab-embedded DepositCount field."""
    raw = df["Deposit Count"].astype(str)
    if raw.str.contains("\t").any():
        logger.warning("DepositCount contains embedded tabs; extracting first value.")
        df = df.copy()
        df.loc[:, "Deposit Count"] = (
            raw.str.split("\t").str[0].pipe(pd.to_numeric, errors="coerce").fillna(0).astype(int)
        )
    else:
        df = df.copy()
        df.loc[:, "Deposit Count"] = pd.to_numeric(
            df["Deposit Count"], errors="coerce"
        ).fillna(0).astype(int)
    return df


def _clean_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Coerce numeric columns and fill NaN with 0."""
    numeric_cols = [
        "Total Items", "Paid Items", "Returned Items",
        "OD Limit", "Avg Bal", "Deposit Amount", "Swipes", "Spend",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df.loc[:, col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    # Returned Items should be int
    df.loc[:, "Returned Items"] = df["Returned Items"].astype(int)
    return df


def _parse_dates(df: pd.DataFrame) -> pd.DataFrame:
    """Parse Open Date and derive Year Opened."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        parsed = pd.to_datetime(df["Open Date"], errors="coerce")
    df = df.assign(**{"Open Date": parsed, "Year Opened": parsed.dt.year})
    return df
