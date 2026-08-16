"""Analysis orchestration and registry."""

import logging
from typing import Callable

import pandas as pd

from ils_kickoff.analyses.base import AnalysisResult
from ils_kickoff.analyses.account_status import analyze_account_status, analyze_account_type
from ils_kickoff.analyses.deposits import (
    analyze_personal_deposits,
    analyze_business_deposits,
)
from ils_kickoff.analyses.nsf_stratification import (
    analyze_nsf_volume_personal,
    analyze_nsf_volume_business,
    analyze_nsf_pay_ratio_personal,
    analyze_nsf_pay_ratio_business,
    analyze_nsf_full_personal,
    analyze_nsf_full_business,
)
from ils_kickoff.analyses.od_status import (
    analyze_od_status_personal,
    analyze_od_status_business,
)
from ils_kickoff.analyses.reg_e import analyze_reg_e_summary, analyze_reg_e_historical
from ils_kickoff.analyses.od_limits import analyze_od_limit_stratification
from ils_kickoff.settings import Settings

logger = logging.getLogger(__name__)

# Ordered list of all analysis functions with their display names
ANALYSIS_REGISTRY: list[tuple[str, Callable]] = [
    ("Account Status Summary", analyze_account_status),
    ("Account Type (Open)", analyze_account_type),
    ("Personal Deposit Distribution", analyze_personal_deposits),
    ("Business Deposit Distribution", analyze_business_deposits),
    ("Personal NSF Stratification", analyze_nsf_volume_personal),
    ("Business NSF Stratification", analyze_nsf_volume_business),
    ("Personal NSF Pay Ratio", analyze_nsf_pay_ratio_personal),
    ("Business NSF Pay Ratio", analyze_nsf_pay_ratio_business),
    ("Personal NSF Full Metrics", analyze_nsf_full_personal),
    ("Business NSF Full Metrics", analyze_nsf_full_business),
    ("Personal OD Status", analyze_od_status_personal),
    ("Business OD Status", analyze_od_status_business),
    ("Reg E Summary", analyze_reg_e_summary),
    ("OD Limit Stratification", analyze_od_limit_stratification),
    ("Historical Reg E by Year", analyze_reg_e_historical),
]


def run_all_analyses(
    df: pd.DataFrame,
    settings: Settings,
    on_progress: Callable | None = None,
) -> list[AnalysisResult]:
    """Run all 15 analyses and return results.

    Failed analyses are logged and skipped (returns partial results).
    """
    # Pre-compute common filtered DataFrames
    open_mask = df["Account Status"] == "O"
    personal_open = df[open_mask & (df["Business Flag"] == "P")]
    business_open = df[open_mask & (df["Business Flag"] == "B")]

    results: list[AnalysisResult] = []
    total = len(ANALYSIS_REGISTRY)

    for i, (name, func) in enumerate(ANALYSIS_REGISTRY):
        logger.info("Running analysis %d/%d: %s", i + 1, total, name)
        if on_progress:
            on_progress(i, total, name)

        try:
            result = func(
                df=df,
                personal_open=personal_open,
                business_open=business_open,
                settings=settings,
            )
            results.append(result)
            logger.info("  -> %d rows", len(result.df))
        except Exception as e:
            logger.error("Analysis '%s' failed: %s", name, e, exc_info=True)
            results.append(AnalysisResult(name=name, title=name, df=pd.DataFrame(), error=str(e)))

    return results
