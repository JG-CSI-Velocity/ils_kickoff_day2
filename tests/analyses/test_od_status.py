"""Tests for OD status analyses (Analyses 11-12)."""

import pytest

from ils_kickoff.analyses.od_status import analyze_od_status_personal, analyze_od_status_business


class TestODStatusPersonal:
    """Analysis 11: Personal OD Status."""

    def test_returns_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_od_status_personal(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.error is None
        assert len(result.df) > 0

    def test_has_od_status_column(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_od_status_personal(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert "OD Status" in result.df.columns


class TestODStatusBusiness:
    """Analysis 12: Business OD Status."""

    def test_returns_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_od_status_business(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.error is None
        assert len(result.df) > 0
