"""Shared value formatting used by both PPTX and Excel reports.

Column-name-based type inference extracted from the original format_ppt_table.
"""

import numpy as np
import pandas as pd


def format_value(val, col_name: str) -> str:
    """Format a value for display based on its column name."""
    if pd.isna(val) or val == "":
        return ""

    if isinstance(val, (int, np.integer)):
        return f"{val:,}"

    if isinstance(val, (float, np.floating)):
        if "%" in col_name:
            return f"{val:.1f}%"
        if "Ratio" in col_name:
            return f"{val:.2f}"
        if is_currency_column(col_name):
            return f"${val:,.0f}" if abs(val) >= 100 else f"${val:.2f}"
        if is_average_column(col_name):
            return f"{val:.1f}"
        return f"{val:,.0f}" if abs(val) >= 10 else f"{val:.2f}"

    return str(val)


def is_currency_column(col_name: str) -> bool:
    """Check if a column name indicates currency values."""
    return "$$" in col_name or "Limit" in col_name or "Dep" in col_name.split("/")[0]


def is_average_column(col_name: str) -> bool:
    """Check if a column name indicates average/median values."""
    return "Avg" in col_name or "Average" in col_name or "Med" in col_name


def is_percentage_column(col_name: str) -> bool:
    """Check if a column name indicates percentage values."""
    return "%" in col_name


def is_ratio_column(col_name: str) -> bool:
    """Check if a column name indicates ratio values."""
    return "Ratio" in col_name.lower()


def excel_number_format(col_name: str) -> str:
    """Return the Excel number format string for a column."""
    if is_percentage_column(col_name):
        return "0.0%"
    if is_ratio_column(col_name):
        return "0.00"
    if is_currency_column(col_name):
        return "$#,##0.00"
    if is_average_column(col_name):
        return "0.0"
    return "#,##0"


def is_grand_total_row(row) -> bool:
    """Check if a row is a Grand Total row by inspecting its first value."""
    first_val = str(row.iloc[0]).lower()
    return "total" in first_val or "grand" in first_val
