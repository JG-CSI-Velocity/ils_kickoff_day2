"""Analyses 1-2: Account Status Summary and Account Type."""

import pandas as pd

from ils_kickoff.analyses.base import AnalysisResult
from ils_kickoff.analyses.templates import grouped_summary, append_grand_total_row
from ils_kickoff.settings import Settings


def analyze_account_status(
    df: pd.DataFrame,
    personal_open: pd.DataFrame,
    business_open: pd.DataFrame,
    settings: Settings,
) -> AnalysisResult:
    """Analysis 1: Account Status Summary (All Accounts)."""
    summary = grouped_summary(
        df=df,
        group_col="Account Status",
        agg_specs={
            "# of Accounts": ("AcctNo", "count"),
            "# of Items": ("Total Items", "sum"),
            "# of Items Paid": ("Paid Items", "sum"),
        },
        pct_of=["# of Accounts", "# of Items"],
        pay_ratio_cols=("# of Items Paid", "# of Items"),
    )

    summary = append_grand_total_row(
        summary,
        label_col="Account Status",
        pay_ratio_cols=("# of Items Paid", "# of Items"),
    )

    # Reorder columns to match original
    col_order = [
        "Account Status", "# of Accounts", "% of # of Accounts",
        "# of Items", "% of # of Items", "# of Items Paid", "Pay Ratio",
    ]
    summary = summary[[c for c in col_order if c in summary.columns]]

    return AnalysisResult(
        name="Account Status Summary",
        title="Account Status Analysis - All Accounts",
        df=summary,
        sheet_name="Stat_Code_Analysis",
    )


def analyze_account_type(
    df: pd.DataFrame,
    personal_open: pd.DataFrame,
    business_open: pd.DataFrame,
    settings: Settings,
) -> AnalysisResult:
    """Analysis 2: Account Type (Open Accounts Only)."""
    df_open = df[df["Account Status"] == "O"]

    summary = grouped_summary(
        df=df_open,
        group_col="Business Flag",
        agg_specs={
            "# of Accounts": ("AcctNo", "count"),
            "# of Items": ("Total Items", "sum"),
            "# of Items Paid": ("Paid Items", "sum"),
        },
        pct_of=["# of Accounts", "# of Items"],
        pay_ratio_cols=("# of Items Paid", "# of Items"),
        label_map={"P": "Personal", "B": "Business"},
    )

    summary = append_grand_total_row(
        summary,
        label_col="Business Flag",
        pay_ratio_cols=("# of Items Paid", "# of Items"),
    )

    return AnalysisResult(
        name="Account Type (Open)",
        title="Account Type Analysis - Open Accounts Only",
        df=summary,
        sheet_name="Account_Type",
    )
