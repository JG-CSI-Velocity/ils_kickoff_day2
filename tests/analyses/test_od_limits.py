"""Tests for OD limit stratification (Analysis 14)."""

import pytest

from ils_kickoff.analyses.od_limits import analyze_od_limit_stratification


class TestODLimitStratification:
    """Analysis 14: OD Limit Stratification."""

    def test_returns_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_od_limit_stratification(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.error is None
        assert len(result.df) > 0

    def test_has_od_limit_bin(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_od_limit_stratification(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert "OD Limit" in result.df.columns

    def test_has_grand_total(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_od_limit_stratification(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        last_val = str(result.df.iloc[-1].iloc[0])
        assert "Total" in last_val or "Grand" in last_val
