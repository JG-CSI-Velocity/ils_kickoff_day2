"""Tests for deposit analyses (Analyses 3-4)."""

import pytest

from ils_kickoff.analyses.deposits import analyze_personal_deposits, analyze_business_deposits


class TestPersonalDeposits:
    """Analysis 3: Personal Deposit Distribution."""

    def test_returns_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_personal_deposits(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.error is None
        assert len(result.df) > 0

    def test_has_deposit_bin_column(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_personal_deposits(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert "Deposit Bin" in result.df.columns


class TestBusinessDeposits:
    """Analysis 4: Business Deposit Distribution."""

    def test_returns_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_business_deposits(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.error is None
        assert len(result.df) > 0
