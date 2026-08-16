"""Tests for Reg E analyses (Analyses 13, 15)."""

import pytest

from ils_kickoff.analyses.reg_e import analyze_reg_e_summary, analyze_reg_e_historical


class TestRegESummary:
    """Analysis 13: Reg E Summary."""

    def test_returns_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_reg_e_summary(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.error is None
        assert len(result.df) > 0

    def test_has_reg_e_column(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_reg_e_summary(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert "Reg E Flag" in result.df.columns


class TestRegEHistorical:
    """Analysis 15: Historical Reg E by Year."""

    def test_returns_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_reg_e_historical(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.error is None
        assert len(result.df) > 0

    def test_has_year_column(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_reg_e_historical(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert "Year Opened" in result.df.columns
