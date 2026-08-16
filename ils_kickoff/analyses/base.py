"""Base analysis infrastructure: AnalysisResult dataclass and helpers."""

from dataclasses import dataclass, field

import numpy as np
import pandas as pd


@dataclass
class AnalysisResult:
    """Result of a single analysis."""

    name: str
    title: str
    df: pd.DataFrame
    error: str | None = None
    sheet_name: str = ""

    def __post_init__(self):
        if not self.sheet_name:
            # Auto-generate sheet name from analysis name (max 31 chars for Excel)
            self.sheet_name = self.name.replace(" ", "_")[:31]


def safe_percentage(numerator: float, denominator: float) -> float:
    """Compute percentage with zero-division guard."""
    if denominator == 0 or pd.isna(denominator):
        return 0.0
    return round(numerator / denominator * 100, 2)


def safe_ratio(numerator: float, denominator: float, decimals: int = 2) -> float:
    """Compute ratio with zero-division guard."""
    if denominator == 0 or pd.isna(denominator):
        return 0.0
    return round(numerator / denominator, decimals)


def add_grand_total(
    df: pd.DataFrame,
    label_col: str,
    label: str = "Grand Total",
    source_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """Append a Grand Total row to a summary DataFrame.

    For sum columns, uses the sum of the summary df.
    For percentage columns, sets to 100.0.
    For ratio/average columns, recomputes from source_df if provided,
    otherwise uses overall mean.
    """
    totals = {}
    for col in df.columns:
        if col == label_col:
            totals[col] = label
        elif "%" in col:
            totals[col] = 100.0
        elif "Ratio" in col:
            # Will be recomputed by caller if needed
            totals[col] = np.nan
        elif "Avg" in col or "Average" in col or "Med" in col:
            if source_df is not None and col.split()[-1] in source_df.columns:
                # Recompute mean from raw data
                raw_col = col.split()[-1]
                totals[col] = source_df[raw_col].mean()
            else:
                totals[col] = df[col].mean()
        else:
            totals[col] = df[col].sum()

    total_row = pd.DataFrame([totals])
    return pd.concat([df, total_row], ignore_index=True)
