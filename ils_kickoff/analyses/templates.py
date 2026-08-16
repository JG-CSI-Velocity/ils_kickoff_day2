"""Parameterized analysis templates that replace 15 copy-pasted analysis blocks.

Two templates handle all 15 analyses:
- grouped_summary: for analyses that group by an existing column
- binned_summary: for analyses that bin a numeric column first, then group
"""

import numpy as np
import pandas as pd

from ils_kickoff.analyses.base import safe_percentage, safe_ratio


def grouped_summary(
    df: pd.DataFrame,
    group_col: str,
    agg_specs: dict[str, tuple[str, str]],
    pct_of: list[str] | None = None,
    pay_ratio_cols: tuple[str, str] | None = None,
    pay_ratio_name: str = "Pay Ratio",
    pay_ratio_pct: bool = False,
    label_map: dict | None = None,
) -> pd.DataFrame:
    """Run a grouped aggregation with optional percentages and pay ratio.

    Args:
        df: Input DataFrame.
        group_col: Column to group by.
        agg_specs: {output_col: (source_col, agg_func)} for pandas named agg.
        pct_of: Column names to compute "% of {col}" as share of total.
        pay_ratio_cols: (numerator_col, denominator_col) for pay ratio.
        pay_ratio_name: Name for the pay ratio column.
        pay_ratio_pct: If True, express ratio as percentage (0-100) instead of decimal.
        label_map: Optional mapping to rename group labels (e.g. {"P": "Personal"}).
    """
    result = df.groupby(group_col).agg(**agg_specs).reset_index()

    if label_map:
        result[group_col] = result[group_col].map(label_map).fillna(result[group_col])

    # Percentage columns
    if pct_of:
        for col in pct_of:
            total = result[col].sum()
            result[f"% of {col}"] = result[col].apply(lambda x: safe_percentage(x, total))

    # Pay ratio
    if pay_ratio_cols:
        num_col, denom_col = pay_ratio_cols
        if pay_ratio_pct:
            result[pay_ratio_name] = np.where(
                result[denom_col] > 0,
                (result[num_col] / result[denom_col] * 100).round(1),
                0,
            )
        else:
            result[pay_ratio_name] = np.where(
                result[denom_col] > 0,
                (result[num_col] / result[denom_col]).round(2),
                0,
            )

    return result


def binned_summary(
    df: pd.DataFrame,
    value_col: str,
    bins: list,
    labels: list[str],
    bin_name: str,
    agg_specs: dict[str, tuple[str, str]],
    pct_of: list[str] | None = None,
    pay_ratio_cols: tuple[str, str] | None = None,
    pay_ratio_name: str = "Pay Ratio",
    pay_ratio_pct: bool = False,
) -> pd.DataFrame:
    """Bin a numeric column, then run grouped_summary on the bins.

    Args:
        df: Input DataFrame (will be copied before mutation).
        value_col: Column to bin.
        bins: Bin edges for pd.cut.
        labels: Labels for each bin.
        bin_name: Name for the bin column in output.
        (remaining args same as grouped_summary)
    """
    df = df.copy()
    df[bin_name] = pd.cut(df[value_col], bins=bins, labels=labels)

    result = df.groupby(bin_name, observed=True).agg(**agg_specs).reset_index()
    result[bin_name] = result[bin_name].astype(str)

    # Percentage columns
    if pct_of:
        for col in pct_of:
            total = result[col].sum()
            result[f"% of {col}"] = result[col].apply(lambda x: safe_percentage(x, total))

    # Pay ratio
    if pay_ratio_cols:
        num_col, denom_col = pay_ratio_cols
        if pay_ratio_pct:
            result[pay_ratio_name] = np.where(
                result[denom_col] > 0,
                (result[num_col] / result[denom_col] * 100).round(1),
                0,
            )
        else:
            result[pay_ratio_name] = np.where(
                result[denom_col] > 0,
                (result[num_col] / result[denom_col]).round(4),
                0,
            )

    return result


def append_grand_total_row(
    summary_df: pd.DataFrame,
    label_col: str,
    source_df: pd.DataFrame | None = None,
    pay_ratio_cols: tuple[str, str] | None = None,
    pay_ratio_name: str = "Pay Ratio",
    pay_ratio_pct: bool = False,
    avg_cols_source: dict[str, str] | None = None,
) -> pd.DataFrame:
    """Append a Grand Total row to a summary DataFrame.

    Args:
        summary_df: The summary to append to.
        label_col: Column for the "Grand Total" label.
        source_df: Raw data for recomputing averages.
        pay_ratio_cols: (numerator_col, denominator_col) to recompute ratio.
        pay_ratio_name: Name of ratio column.
        pay_ratio_pct: If True, ratio is percentage.
        avg_cols_source: {summary_col: source_col} for recomputing averages from source_df.
    """
    totals = {}
    for col in summary_df.columns:
        if col == label_col:
            totals[col] = "Grand Total"
        elif "%" in col:
            totals[col] = 100.0
        elif col == pay_ratio_name and pay_ratio_cols:
            num_col, denom_col = pay_ratio_cols
            num_total = summary_df[num_col].sum()
            denom_total = summary_df[denom_col].sum()
            if pay_ratio_pct:
                totals[col] = round(num_total / denom_total * 100, 1) if denom_total > 0 else 0
            else:
                totals[col] = round(num_total / denom_total, 4) if denom_total > 0 else 0
        elif avg_cols_source and col in avg_cols_source and source_df is not None:
            totals[col] = source_df[avg_cols_source[col]].mean()
        elif "Avg" in col or "Average" in col:
            if source_df is not None:
                # Try to find matching source column
                totals[col] = summary_df[col].mean()
            else:
                totals[col] = summary_df[col].mean()
        else:
            totals[col] = summary_df[col].sum()

    total_row = pd.DataFrame([totals])
    return pd.concat([summary_df, total_row], ignore_index=True)
