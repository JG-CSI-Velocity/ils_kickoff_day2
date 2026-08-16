"""Analyses 11-12: OD Status Stratification (Personal and Business)."""

import pandas as pd

from ils_kickoff.analyses.base import AnalysisResult
from ils_kickoff.analyses.templates import grouped_summary, append_grand_total_row
from ils_kickoff.settings import Settings


def _od_status_analysis(
    source_df: pd.DataFrame,
    name: str,
    title: str,
    sheet_name: str,
) -> AnalysisResult:
    """Shared logic for personal/business OD status analysis."""
    summary = grouped_summary(
        df=source_df,
        group_col="OD Status",
        agg_specs={
            "# of Accounts": ("AcctNo", "count"),
            "Total OD/NSF Items": ("Total Items", "sum"),
            "# of Items Paid": ("Paid Items", "sum"),
        },
        pct_of=["# of Accounts", "Total OD/NSF Items"],
        pay_ratio_cols=("# of Items Paid", "Total OD/NSF Items"),
    )

    # Cast OD Status to string for display
    summary["OD Status"] = summary["OD Status"].astype(str)

    summary = summary.rename(columns={
        "% of # of Accounts": "% of Accounts",
        "% of Total OD/NSF Items": "% of Items Presented",
    })

    summary = append_grand_total_row(
        summary,
        label_col="OD Status",
        pay_ratio_cols=("# of Items Paid", "Total OD/NSF Items"),
    )

    return AnalysisResult(name=name, title=title, df=summary, sheet_name=sheet_name)


def analyze_od_status_personal(
    df: pd.DataFrame,
    personal_open: pd.DataFrame,
    business_open: pd.DataFrame,
    settings: Settings,
) -> AnalysisResult:
    """Analysis 11: Personal OD Status Stratification."""
    return _od_status_analysis(
        personal_open,
        "Personal OD Status",
        "Personal OD Status Code Stratification",
        "OD_Status_Personal",
    )


def analyze_od_status_business(
    df: pd.DataFrame,
    personal_open: pd.DataFrame,
    business_open: pd.DataFrame,
    settings: Settings,
) -> AnalysisResult:
    """Analysis 12: Business OD Status Stratification."""
    return _od_status_analysis(
        business_open,
        "Business OD Status",
        "Business OD Status Code Stratification",
        "OD_Status_Business",
    )
