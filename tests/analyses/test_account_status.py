"""Tests for account status analyses (Analyses 1-2)."""

import pandas as pd
import pytest

from ils_kickoff.analyses.account_status import analyze_account_status, analyze_account_type


class TestAccountStatus:
    """Analysis 1: Account Status Summary."""

    def test_returns_analysis_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_account_status(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.name == "Account Status Summary"
        assert result.error is None
        assert len(result.df) > 0

    def test_has_expected_columns(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_account_status(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert "Account Status" in result.df.columns
        assert "# of Accounts" in result.df.columns

    def test_has_grand_total_row(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_account_status(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        last_row = result.df.iloc[-1]
        assert "Grand Total" in str(last_row.iloc[0]) or "Total" in str(last_row.iloc[0])

    def test_percentages_sum_to_100(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_account_status(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        pct_cols = [c for c in result.df.columns if "%" in c]
        if pct_cols:
            # Exclude Grand Total row for percentage sum check
            data_rows = result.df.iloc[:-1]
            for col in pct_cols:
                total = pd.to_numeric(data_rows[col], errors="coerce").sum()
                assert abs(total - 100.0) < 0.5, f"{col} sums to {total}, expected ~100"


class TestAccountType:
    """Analysis 2: Account Type Breakdown."""

    def test_returns_analysis_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_account_type(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.name == "Account Type (Open)"
        assert result.error is None
        assert len(result.df) > 0

    def test_has_business_flag_column(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_account_type(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert "Business Flag" in result.df.columns
