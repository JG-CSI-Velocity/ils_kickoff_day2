"""Analyses 13, 15: Reg E Summary and Historical Reg E by Year Opened."""

from datetime import datetime

import numpy as np
import pandas as pd

from ils_kickoff.analyses.base import AnalysisResult, safe_percentage
from ils_kickoff.analyses.templates import grouped_summary, append_grand_total_row
from ils_kickoff.settings import Settings


def analyze_reg_e_summary(
    df: pd.DataFrame,
    personal_open: pd.DataFrame,
    business_open: pd.DataFrame,
    settings: Settings,
) -> AnalysisResult:
    """Analysis 13: Reg E Distribution (Personal Open Accounts)."""
    summary = grouped_summary(
        df=personal_open,
        group_col="Reg E Flag",
        agg_specs={"# of Accounts": ("AcctNo", "count")},
        pct_of=["# of Accounts"],
    )

    summary = summary.rename(columns={"% of # of Accounts": "% of Accounts"})

    summary = append_grand_total_row(summary, label_col="Reg E Flag")

    return AnalysisResult(
        name="Reg E Summary",
        title="Reg E Distribution - Personal Open Accounts",
        df=summary,
        sheet_name="Reg_E_Summary",
    )


def analyze_reg_e_historical(
    df: pd.DataFrame,
    personal_open: pd.DataFrame,
    business_open: pd.DataFrame,
    settings: Settings,
) -> AnalysisResult:
    """Analysis 15: Historical Reg E by Year Opened."""
    p_hist = personal_open[personal_open["Open Date"].notna()].copy()

    current_year = datetime.now().year

    def assign_year_bin(year):
        if pd.isna(year):
            return "Unknown"
        y = int(year)
        if y < 2010:
            return "<2010"
        if y <= current_year:
            return str(y)
        return f"{current_year}+"

    p_hist["Year Bin"] = p_hist["Year Opened"].apply(assign_year_bin)

    # Pivot to get Reg E flags as columns
    pivot_raw = (
        p_hist.groupby(["Year Bin", "Reg E Flag"], dropna=False)
        .agg(**{"# of Accounts": ("AcctNo", "count")})
        .reset_index()
    )
    pivot_table = pivot_raw.pivot(
        index="Year Bin", columns="Reg E Flag", values="# of Accounts"
    ).fillna(0)

    reg_e_flags = sorted([c for c in pivot_table.columns if c is not None])
    pivot_table = pivot_table.reindex(columns=reg_e_flags)
    pivot_table["# of Accounts"] = pivot_table.sum(axis=1)

    # Compute opt-in percentage
    opt_in_flag = "Y" if "Y" in reg_e_flags else (reg_e_flags[0] if reg_e_flags else None)
    if opt_in_flag:
        denom = pivot_table["# of Accounts"].replace({0: pd.NA})
        pivot_table["Opt In %"] = (pivot_table[opt_in_flag] / denom * 100).fillna(0)
    else:
        pivot_table["Opt In %"] = 0.0

    # Grand Total
    grand_total_hist = pd.DataFrame(pivot_table.sum(numeric_only=True)).T
    grand_total_hist.index = ["Grand Total"]
    if opt_in_flag:
        total_accts = grand_total_hist["# of Accounts"].iloc[0]
        opt_in_count = grand_total_hist[opt_in_flag].iloc[0]
        grand_total_hist["Opt In %"] = (
            round(opt_in_count / total_accts * 100, 1) if total_accts > 0 else 0
        )
    pivot_table = pd.concat([pivot_table, grand_total_hist])

    # Clean types
    int_cols = [c for c in pivot_table.columns if c in reg_e_flags or "# of" in c]
    for c in int_cols:
        pivot_table[c] = pivot_table[c].round(0).astype(int)
    pivot_table["Opt In %"] = pivot_table["Opt In %"].round(1)

    # Reset index and rename
    pivot_table = pivot_table.reset_index()
    if "Year Bin" in pivot_table.columns:
        pivot_table = pivot_table.rename(columns={"Year Bin": "Year Opened"})
    elif "index" in pivot_table.columns:
        pivot_table = pivot_table.rename(columns={"index": "Year Opened"})

    # Sort by year
    sort_order = (
        ["<2010"]
        + [str(y) for y in range(2010, current_year + 1)]
        + [f"{current_year}+", "Unknown", "Grand Total"]
    )
    pivot_table["Year Opened"] = pd.Categorical(
        pivot_table["Year Opened"], categories=sort_order, ordered=True
    )
    pivot_table = pivot_table.sort_values("Year Opened")

    return AnalysisResult(
        name="Historical Reg E by Year",
        title="Historical Reg E Opt-In by Year Opened",
        df=pivot_table,
        sheet_name="Historical_Reg_E",
    )
