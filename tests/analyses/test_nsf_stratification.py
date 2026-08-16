"""Tests for NSF stratification analyses (Analyses 5-10)."""

import pandas as pd
import pytest

from ils_kickoff.analyses.nsf_stratification import (
    analyze_nsf_volume_personal,
    analyze_nsf_volume_business,
    analyze_nsf_pay_ratio_personal,
    analyze_nsf_pay_ratio_business,
    analyze_nsf_full_personal,
    analyze_nsf_full_business,
)


class TestNSFVolumePersonal:
    """Analysis 5: Personal NSF Volume Stratification."""

    def test_returns_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_nsf_volume_personal(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.error is None
        assert len(result.df) > 0

    def test_has_nsf_bin_column(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_nsf_volume_personal(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert "NSF Bin" in result.df.columns

    def test_has_grand_total(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_nsf_volume_personal(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        last_val = str(result.df.iloc[-1]["NSF Bin"])
        assert "Total" in last_val or "Grand" in last_val


class TestNSFVolumeBusiness:
    """Analysis 6: Business NSF Volume Stratification."""

    def test_returns_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_nsf_volume_business(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.error is None
        assert len(result.df) > 0


class TestNSFPayRatioPersonal:
    """Analysis 7: Personal NSF Pay Ratio."""

    def test_returns_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_nsf_pay_ratio_personal(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.error is None

    def test_has_pay_rate_column(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_nsf_pay_ratio_personal(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert "% Pay Rate" in result.df.columns


class TestNSFPayRatioBusiness:
    """Analysis 8: Business NSF Pay Ratio."""

    def test_returns_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_nsf_pay_ratio_business(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.error is None


class TestNSFFullPersonal:
    """Analysis 9: Personal NSF Full Behavioral Metrics."""

    def test_returns_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_nsf_full_personal(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.error is None

    def test_has_behavioral_columns(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_nsf_full_personal(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        for col in ["Avg # Dep/Month", "Average of Swipes"]:
            assert col in result.df.columns, f"Missing column: {col}"

    def test_average_columns_are_numeric(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_nsf_full_personal(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        for col in ["Avg # Dep/Month", "Average of Swipes"]:
            if col in result.df.columns:
                vals = pd.to_numeric(result.df[col], errors="coerce")
                assert vals.notna().all(), f"{col} has non-numeric values"


class TestNSFFullBusiness:
    """Analysis 10: Business NSF Full Behavioral Metrics."""

    def test_returns_result(self, sample_df, personal_open, business_open, sample_settings):
        result = analyze_nsf_full_business(
            df=sample_df, personal_open=personal_open,
            business_open=business_open, settings=sample_settings,
        )
        assert result.error is None
