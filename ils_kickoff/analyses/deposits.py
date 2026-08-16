"""Analyses 3-4: Deposit Distribution (Personal and Business)."""

import pandas as pd

from ils_kickoff.analyses.base import AnalysisResult
from ils_kickoff.analyses.templates import binned_summary, append_grand_total_row
from ils_kickoff.settings import Settings


def _deposit_analysis(
    source_df: pd.DataFrame,
    settings: Settings,
    name: str,
    title: str,
    sheet_name: str,
) -> AnalysisResult:
    """Shared logic for personal/business deposit distribution."""
    summary = binned_summary(
        df=source_df,
        value_col="Deposit Count",
        bins=settings.deposit_bins,
        labels=settings.deposit_labels,
        bin_name="Deposit Bin",
        agg_specs={
            "Accounts": ("AcctNo", "count"),
            "Avg $$ Deposits": ("Deposit Amount", "mean"),
            "Avg # Deposits": ("Deposit Count", "mean"),
        },
        pct_of=["Accounts"],
    )

    summary = append_grand_total_row(
        summary,
        label_col="Deposit Bin",
        source_df=source_df,
        avg_cols_source={
            "Avg $$ Deposits": "Deposit Amount",
            "Avg # Deposits": "Deposit Count",
        },
    )

    return AnalysisResult(name=name, title=title, df=summary, sheet_name=sheet_name)


def analyze_personal_deposits(
    df: pd.DataFrame,
    personal_open: pd.DataFrame,
    business_open: pd.DataFrame,
    settings: Settings,
) -> AnalysisResult:
    """Analysis 3: Personal Deposit Distribution."""
    return _deposit_analysis(
        source_df=personal_open,
        settings=settings,
        name="Personal Deposit Distribution",
        title="Personal Account Deposit Analysis",
        sheet_name="Personal_Deposits",
    )


def analyze_business_deposits(
    df: pd.DataFrame,
    personal_open: pd.DataFrame,
    business_open: pd.DataFrame,
    settings: Settings,
) -> AnalysisResult:
    """Analysis 4: Business Deposit Distribution."""
    return _deposit_analysis(
        source_df=business_open,
        settings=settings,
        name="Business Deposit Distribution",
        title="Business Account Deposit Analysis",
        sheet_name="Business_Deposits",
    )
