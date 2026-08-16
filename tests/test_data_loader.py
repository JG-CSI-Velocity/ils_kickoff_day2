"""Tests for data loading and validation."""

from pathlib import Path

import pandas as pd
import pytest

from ils_kickoff.column_map import REQUIRED_COLUMNS, resolve_columns
from ils_kickoff.data_loader import load_data
from ils_kickoff.exceptions import ColumnMismatchError, DataLoadError
from ils_kickoff.settings import Settings


class TestResolveColumns:
    """Test column alias resolution."""

    def test_renames_script_aliases(self):
        df = pd.DataFrame({"TOTALITEMS": [1], "PaidItems": [1]})
        result = resolve_columns(df)
        assert "Total Items" in result.columns
        assert "Paid Items" in result.columns

    def test_renames_notebook_aliases(self):
        df = pd.DataFrame({"TotalItems": [1], "Paid_Items": [1]})
        result = resolve_columns(df)
        assert "Total Items" in result.columns
        assert "Paid Items" in result.columns

    def test_preserves_canonical_names(self):
        df = pd.DataFrame({"Total Items": [1], "Paid Items": [1]})
        result = resolve_columns(df)
        assert "Total Items" in result.columns
        assert "Paid Items" in result.columns

    def test_does_not_mutate_input(self):
        df = pd.DataFrame({"TOTALITEMS": [1]})
        result = resolve_columns(df)
        assert "TOTALITEMS" in df.columns
        assert "Total Items" in result.columns

    def test_normalizes_paid_items_variant(self):
        df = pd.DataFrame({"# of Paid Items": [5]})
        result = resolve_columns(df)
        assert "# of Items Paid" in result.columns


class TestLoadData:
    """Test the full data loading pipeline."""

    def test_loads_sample_csv(self, sample_settings):
        df = load_data(sample_settings)
        assert len(df) == 100
        # Check all required columns are present
        missing = REQUIRED_COLUMNS - set(df.columns)
        assert missing == set(), f"Missing columns: {missing}"

    def test_deposit_count_is_numeric(self, sample_settings):
        df = load_data(sample_settings)
        assert pd.api.types.is_integer_dtype(df["Deposit Count"])

    def test_numeric_columns_coerced(self, sample_settings):
        df = load_data(sample_settings)
        for col in ["Total Items", "Paid Items", "OD Limit", "Avg Bal"]:
            assert pd.api.types.is_numeric_dtype(df[col]), f"{col} should be numeric"

    def test_returned_items_is_int(self, sample_settings):
        df = load_data(sample_settings)
        assert pd.api.types.is_integer_dtype(df["Returned Items"])

    def test_open_date_parsed(self, sample_settings):
        df = load_data(sample_settings)
        assert pd.api.types.is_datetime64_any_dtype(df["Open Date"])

    def test_year_opened_derived(self, sample_settings):
        df = load_data(sample_settings)
        assert "Year Opened" in df.columns
        assert df["Year Opened"].notna().all()

    def test_missing_columns_raises(self, tmp_path):
        csv = tmp_path / "bad.csv"
        csv.write_text("ColA,ColB\n1,2\n")
        settings = Settings.from_args(data_file=csv)
        with pytest.raises(ColumnMismatchError) as exc_info:
            load_data(settings)
        assert exc_info.value.missing
        assert exc_info.value.available


class TestLoadDataEdgeCases:
    """Edge case handling in data loading."""

    def test_tab_embedded_deposit_count(self, tmp_path):
        """DepositCount with tab-separated trailing values."""
        csv = tmp_path / "tabs.csv"
        header = "AcctNo,TOTALITEMS,PaidItems,ReturnedItems,ODLimit,ODStatus,ProdCode,BusinessFlag,AccountStatus,RegEValue,OpenDate,AvgColBal,DepositAmount,DepositCount,swipes,spend"
        row = "A001,5,4,1,500,Active OD,100,P,O,Y,01/15/2020,1200.50,3500.00,8\t0.00\t0.00\t0.00,25,450.00"
        csv.write_text(f"{header}\n{row}\n")
        settings = Settings.from_args(data_file=csv)
        df = load_data(settings)
        assert df["Deposit Count"].iloc[0] == 8
