"""Tests for the Settings configuration system."""

from pathlib import Path

import pytest

from ils_kickoff.exceptions import ConfigError
from ils_kickoff.settings import Settings


class TestSettingsFromArgs:
    """Test Settings.from_args() direct construction."""

    def test_valid_data_file(self, tmp_path):
        csv = tmp_path / "test.csv"
        csv.write_text("col1,col2\n1,2\n")
        s = Settings.from_args(data_file=csv)
        assert s.data_file == csv

    def test_missing_data_file_raises(self, tmp_path):
        with pytest.raises(ConfigError):
            Settings.from_args(data_file=tmp_path / "nonexistent.csv")

    def test_unsupported_extension_raises(self, tmp_path):
        txt_file = tmp_path / "data.txt"
        txt_file.write_text("data")
        with pytest.raises(ConfigError):
            Settings.from_args(data_file=txt_file)

    def test_xlsx_extension_accepted(self, tmp_path):
        xlsx = tmp_path / "data.xlsx"
        xlsx.write_bytes(b"\x00")  # Minimal file
        s = Settings.from_args(data_file=xlsx)
        assert s.data_file.suffix == ".xlsx"

    def test_client_id_auto_extracted(self, tmp_path):
        csv = tmp_path / "1774_INB_OD_data.csv"
        csv.write_text("a,b\n1,2\n")
        s = Settings.from_args(data_file=csv)
        assert s.client_id == "1774"

    def test_client_id_fallback_unknown(self, tmp_path):
        csv = tmp_path / "data_without_number.csv"
        csv.write_text("a,b\n1,2\n")
        s = Settings.from_args(data_file=csv)
        assert s.client_id == "unknown"

    def test_client_id_explicit_overrides_auto(self, tmp_path):
        csv = tmp_path / "1774_data.csv"
        csv.write_text("a,b\n1,2\n")
        s = Settings.from_args(data_file=csv, client_id="9999")
        assert s.client_id == "9999"

    def test_client_name_derives_from_id(self, tmp_path):
        csv = tmp_path / "1774_data.csv"
        csv.write_text("a,b\n1,2\n")
        s = Settings.from_args(data_file=csv)
        assert s.client_name == "Client 1774"

    def test_client_name_explicit(self, tmp_path):
        csv = tmp_path / "data.csv"
        csv.write_text("a,b\n1,2\n")
        s = Settings.from_args(data_file=csv, client_name="Test Bank")
        assert s.client_name == "Test Bank"

    def test_default_output_dir(self, tmp_path):
        csv = tmp_path / "data.csv"
        csv.write_text("a,b\n1,2\n")
        s = Settings.from_args(data_file=csv)
        assert s.output_dir == Path("output/")

    def test_default_outputs(self, tmp_path):
        csv = tmp_path / "data.csv"
        csv.write_text("a,b\n1,2\n")
        s = Settings.from_args(data_file=csv)
        assert s.outputs.excel is True
        assert s.outputs.powerpoint is True
        assert s.outputs.html_charts is False


class TestSettingsFromYaml:
    """Test Settings.from_yaml() YAML loading."""

    def test_load_from_yaml(self, sample_config_yaml, tmp_path):
        s = Settings.from_yaml(config_path=sample_config_yaml)
        assert s.data_file.exists()
        assert s.outputs.excel is True

    def test_cli_overrides_yaml(self, sample_config_yaml, tmp_path):
        s = Settings.from_yaml(
            config_path=sample_config_yaml,
            client_id="OVERRIDE",
        )
        assert s.client_id == "OVERRIDE"

    def test_missing_config_uses_defaults(self, tmp_path):
        csv = tmp_path / "data.csv"
        csv.write_text("a,b\n1,2\n")
        s = Settings.from_yaml(
            config_path=tmp_path / "nonexistent.yaml",
            data_file=csv,
        )
        assert s.data_file == csv

    def test_invalid_yaml_raises_config_error(self, tmp_path):
        yaml_file = tmp_path / "bad.yaml"
        yaml_file.write_text("data_file: /nonexistent/file.csv\n")
        with pytest.raises(ConfigError):
            Settings.from_yaml(config_path=yaml_file)


class TestSettingsDefaults:
    """Test default values for bins and chart config."""

    def test_default_nsf_bins(self, tmp_path):
        csv = tmp_path / "data.csv"
        csv.write_text("a,b\n1,2\n")
        s = Settings.from_args(data_file=csv)
        assert len(s.nsf_bins) == 8
        assert len(s.nsf_labels) == 7

    def test_default_chart_config(self, tmp_path):
        csv = tmp_path / "data.csv"
        csv.write_text("a,b\n1,2\n")
        s = Settings.from_args(data_file=csv)
        assert s.charts.theme == "plotly_white"
        assert s.charts.width == 900
        assert s.charts.height == 500
        assert len(s.charts.colors) == 6
