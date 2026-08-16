"""Analysis 14: OD Limit Stratification (Personal)."""

import pandas as pd

from ils_kickoff.analyses.base import AnalysisResult
from ils_kickoff.analyses.templates import grouped_summary, append_grand_total_row
from ils_kickoff.settings import Settings


def analyze_od_limit_stratification(
    df: pd.DataFrame,
    personal_open: pd.DataFrame,
    business_open: pd.DataFrame,
    settings: Settings,
) -> AnalysisResult:
    """Analysis 14: Personal OD Limit Stratification."""
    summary = grouped_summary(
        df=personal_open,
        group_col="OD Limit",
        agg_specs={
            "# of Accounts": ("AcctNo", "count"),
            "Total OD/NSF Items": ("Total Items", "sum"),
            "# of Items Paid": ("Paid Items", "sum"),
            "Avg # Dep/Month": ("Deposit Count", "mean"),
            "Avg $$ Dep/Month": ("Deposit Amount", "mean"),
            "Average of Swipes": ("Swipes", "mean"),
        },
        pct_of=["# of Accounts", "Total OD/NSF Items"],
        pay_ratio_cols=("# of Items Paid", "Total OD/NSF Items"),
    )

    summary["OD Limit"] = summary["OD Limit"].astype(int).astype(str)

    summary = summary.rename(columns={
        "% of # of Accounts": "% of Accounts",
        "% of Total OD/NSF Items": "% of Items",
    })

    summary = append_grand_total_row(
        summary,
        label_col="OD Limit",
        source_df=personal_open,
        pay_ratio_cols=("# of Items Paid", "Total OD/NSF Items"),
        avg_cols_source={
            "Avg # Dep/Month": "Deposit Count",
            "Avg $$ Dep/Month": "Deposit Amount",
            "Average of Swipes": "Swipes",
        },
    )

    col_order = [
        "OD Limit", "# of Accounts", "% of Accounts",
        "Total OD/NSF Items", "% of Items", "# of Items Paid", "Pay Ratio",
        "Avg # Dep/Month", "Avg $$ Dep/Month", "Average of Swipes",
    ]
    summary = summary[[c for c in col_order if c in summary.columns]]

    return AnalysisResult(
        name="OD Limit Stratification",
        title="Personal OD Limit Stratification",
        df=summary,
        sheet_name="OD_Limit_Strat",
    )
