"""Shared test fixtures for the ILS Kickoff test suite."""

from pathlib import Path

import pandas as pd
import pytest

from ils_kickoff.column_map import resolve_columns
from ils_kickoff.settings import Settings


SAMPLE_CSV = Path(__file__).parent / "data" / "sample.csv"


@pytest.fixture
def raw_df():
    """Load raw sample CSV without column resolution."""
    return pd.read_csv(SAMPLE_CSV, encoding="utf-8-sig")


@pytest.fixture
def sample_df(raw_df):
    """Load and resolve sample CSV to canonical column names."""
    df = resolve_columns(raw_df)
    # Mimic data_loader cleaning
    df["Deposit Count"] = pd.to_numeric(df["Deposit Count"], errors="coerce").fillna(0).astype(int)
    for col in ["Total Items", "Paid Items", "Returned Items",
                 "OD Limit", "Avg Bal", "Deposit Amount", "Swipes", "Spend"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["Returned Items"] = df["Returned Items"].astype(int)
    parsed = pd.to_datetime(df["Open Date"], errors="coerce")
    df = df.assign(**{"Open Date": parsed, "Year Opened": parsed.dt.year})
    return df


@pytest.fixture
def personal_open(sample_df):
    """Open personal accounts."""
    mask = (sample_df["Account Status"] == "O") & (sample_df["Business Flag"] == "P")
    return sample_df[mask]


@pytest.fixture
def business_open(sample_df):
    """Open business accounts."""
    mask = (sample_df["Account Status"] == "O") & (sample_df["Business Flag"] == "B")
    return sample_df[mask]


@pytest.fixture
def sample_settings(tmp_path):
    """Create a Settings object pointing to the sample CSV."""
    return Settings(data_file=SAMPLE_CSV, output_dir=tmp_path / "output")


@pytest.fixture
def sample_config_yaml(tmp_path):
    """Create a temporary config.yaml for testing."""
    config = tmp_path / "config.yaml"
    config.write_text(
        f'data_file: "{SAMPLE_CSV}"\n'
        f'output_dir: "{tmp_path / "output"}"\n'
        "outputs:\n"
        "  excel: true\n"
        "  powerpoint: true\n"
    )
    return config
