"""Analyses 5-10: NSF/OD Stratification (Volume, Pay Ratio, Full Metrics)."""

import pandas as pd

from ils_kickoff.analyses.base import AnalysisResult
from ils_kickoff.analyses.templates import binned_summary, append_grand_total_row
from ils_kickoff.settings import Settings


def _nsf_volume(
    source_df: pd.DataFrame,
    settings: Settings,
    name: str,
    title: str,
    sheet_name: str,
) -> AnalysisResult:
    """NSF volume stratification (Analyses 5-6)."""
    summary = binned_summary(
        df=source_df,
        value_col="Total Items",
        bins=settings.nsf_bins,
        labels=settings.nsf_labels,
        bin_name="NSF Bin",
        agg_specs={
            "# of Accounts": ("AcctNo", "count"),
            "Total OD/NSF Items": ("Total Items", "sum"),
        },
        pct_of=["# of Accounts", "Total OD/NSF Items"],
    )

    # Rename pct columns for clarity
    summary = summary.rename(columns={
        "% of # of Accounts": "% of Accounts",
        "% of Total OD/NSF Items": "% of Items Presented",
    })

    summary = append_grand_total_row(summary, label_col="NSF Bin")

    col_order = ["NSF Bin", "# of Accounts", "% of Accounts",
                 "Total OD/NSF Items", "% of Items Presented"]
    summary = summary[[c for c in col_order if c in summary.columns]]

    return AnalysisResult(name=name, title=title, df=summary, sheet_name=sheet_name)


def _nsf_pay_ratio(
    source_df: pd.DataFrame,
    settings: Settings,
    name: str,
    title: str,
    sheet_name: str,
) -> AnalysisResult:
    """NSF pay ratio stratification (Analyses 7-8)."""
    summary = binned_summary(
        df=source_df,
        value_col="Total Items",
        bins=settings.nsf_bins,
        labels=settings.nsf_labels,
        bin_name="NSF Bin",
        agg_specs={
            "# of Accounts": ("AcctNo", "count"),
            "Total OD/NSF Items": ("Total Items", "sum"),
            "# of Items Paid": ("Paid Items", "sum"),
        },
        pct_of=["# of Accounts", "Total OD/NSF Items"],
        pay_ratio_cols=("# of Items Paid", "Total OD/NSF Items"),
        pay_ratio_name="% Pay Rate",
        pay_ratio_pct=True,
    )

    summary = summary.rename(columns={
        "% of # of Accounts": "% of Accounts",
        "% of Total OD/NSF Items": "% of Items Presented",
    })

    summary = append_grand_total_row(
        summary,
        label_col="NSF Bin",
        pay_ratio_cols=("# of Items Paid", "Total OD/NSF Items"),
        pay_ratio_name="% Pay Rate",
        pay_ratio_pct=True,
    )

    return AnalysisResult(name=name, title=title, df=summary, sheet_name=sheet_name)


def _nsf_full_metrics(
    source_df: pd.DataFrame,
    settings: Settings,
    name: str,
    title: str,
    sheet_name: str,
) -> AnalysisResult:
    """NSF full behavioral metrics (Analyses 9-10)."""
    summary = binned_summary(
        df=source_df,
        value_col="Total Items",
        bins=settings.nsf_bins,
        labels=settings.nsf_labels,
        bin_name="NSF Bin",
        agg_specs={
            "# of Accounts": ("AcctNo", "count"),
            "Total OD/NSF Items": ("Total Items", "sum"),
            "# of Items Paid": ("Paid Items", "sum"),
            "Avg # Dep/Month": ("Deposit Count", "mean"),
            "Avg $$ Dep/Month": ("Deposit Amount", "mean"),
            "Average of OD Limit": ("OD Limit", "mean"),
            "Average of Swipes": ("Swipes", "mean"),
        },
        pct_of=["# of Accounts", "Total OD/NSF Items"],
        pay_ratio_cols=("# of Items Paid", "Total OD/NSF Items"),
        pay_ratio_name="% Pay Rate",
        pay_ratio_pct=True,
    )

    summary = summary.rename(columns={
        "% of # of Accounts": "% of Accounts",
        "% of Total OD/NSF Items": "% of Items Presented",
    })

    summary = append_grand_total_row(
        summary,
        label_col="NSF Bin",
        source_df=source_df,
        pay_ratio_cols=("# of Items Paid", "Total OD/NSF Items"),
        pay_ratio_name="% Pay Rate",
        pay_ratio_pct=True,
        avg_cols_source={
            "Avg # Dep/Month": "Deposit Count",
            "Avg $$ Dep/Month": "Deposit Amount",
            "Average of OD Limit": "OD Limit",
            "Average of Swipes": "Swipes",
        },
    )

    # Round average columns (coerce to numeric first since grand total row may introduce object dtype)
    for col in ["Avg # Dep/Month", "Avg $$ Dep/Month", "Average of OD Limit", "Average of Swipes"]:
        if col in summary.columns:
            summary[col] = pd.to_numeric(summary[col], errors="coerce").round(2)

    return AnalysisResult(name=name, title=title, df=summary, sheet_name=sheet_name)


# Public API

def analyze_nsf_volume_personal(
    df: pd.DataFrame, personal_open: pd.DataFrame,
    business_open: pd.DataFrame, settings: Settings,
) -> AnalysisResult:
    """Analysis 5: Personal NSF Stratification (Volume)."""
    return _nsf_volume(
        personal_open, settings,
        "Personal NSF Stratification", "Personal NSF/OD Stratification - Volume",
        "NSF_Strat_Personal",
    )


def analyze_nsf_volume_business(
    df: pd.DataFrame, personal_open: pd.DataFrame,
    business_open: pd.DataFrame, settings: Settings,
) -> AnalysisResult:
    """Analysis 6: Business NSF Stratification (Volume)."""
    return _nsf_volume(
        business_open, settings,
        "Business NSF Stratification", "Business NSF/OD Stratification - Volume",
        "NSF_Strat_Business",
    )


def analyze_nsf_pay_ratio_personal(
    df: pd.DataFrame, personal_open: pd.DataFrame,
    business_open: pd.DataFrame, settings: Settings,
) -> AnalysisResult:
    """Analysis 7: Personal NSF + Pay Ratio."""
    return _nsf_pay_ratio(
        personal_open, settings,
        "Personal NSF Pay Ratio", "Personal NSF/OD Stratification - Pay Ratio",
        "NSF_PayRatio_Personal",
    )


def analyze_nsf_pay_ratio_business(
    df: pd.DataFrame, personal_open: pd.DataFrame,
    business_open: pd.DataFrame, settings: Settings,
) -> AnalysisResult:
    """Analysis 8: Business NSF + Pay Ratio."""
    return _nsf_pay_ratio(
        business_open, settings,
        "Business NSF Pay Ratio", "Business NSF/OD Stratification - Pay Ratio",
        "NSF_PayRatio_Business",
    )


def analyze_nsf_full_personal(
    df: pd.DataFrame, personal_open: pd.DataFrame,
    business_open: pd.DataFrame, settings: Settings,
) -> AnalysisResult:
    """Analysis 9: Personal NSF + Deposits + Swipes."""
    return _nsf_full_metrics(
        personal_open, settings,
        "Personal NSF Full Metrics", "Personal NSF/OD - Full Behavioral Metrics",
        "NSF_Full_Personal",
    )


def analyze_nsf_full_business(
    df: pd.DataFrame, personal_open: pd.DataFrame,
    business_open: pd.DataFrame, settings: Settings,
) -> AnalysisResult:
    """Analysis 10: Business NSF + Deposits + Swipes."""
    return _nsf_full_metrics(
        business_open, settings,
        "Business NSF Full Metrics", "Business NSF/OD - Full Behavioral Metrics",
        "NSF_Full_Business",
    )
