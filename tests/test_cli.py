"""Tests for the CLI entry point."""

from pathlib import Path

import pytest
from typer.testing import CliRunner

from ils_kickoff.cli import app

runner = CliRunner()


class TestCLIHelp:
    """Test --help output."""

    def test_help_exits_zero(self):
        result = runner.invoke(app, ["--help"])
        assert result.exit_code == 0

    def test_help_shows_usage(self):
        result = runner.invoke(app, ["--help"])
        assert "data_file" in result.output.lower() or "DATA_FILE" in result.output


class TestCLIExecution:
    """Test CLI execution paths."""

    def test_missing_config_and_no_data_file(self, tmp_path):
        result = runner.invoke(app, ["--config", str(tmp_path / "nonexistent.yaml")])
        assert result.exit_code != 0

    def test_run_with_sample_data(self, sample_config_yaml, tmp_path):
        result = runner.invoke(app, [
            "--config", str(sample_config_yaml),
        ])
        assert result.exit_code == 0

    def test_verbose_flag(self, sample_config_yaml):
        result = runner.invoke(app, [
            "--config", str(sample_config_yaml),
            "--verbose",
        ])
        assert result.exit_code == 0

    def test_custom_output_dir(self, sample_config_yaml, tmp_path):
        output = tmp_path / "custom_output"
        result = runner.invoke(app, [
            "--config", str(sample_config_yaml),
            "--output", str(output),
        ])
        assert result.exit_code == 0
        assert output.exists()

    def test_client_id_override(self, sample_config_yaml):
        result = runner.invoke(app, [
            "--config", str(sample_config_yaml),
            "--client-id", "TEST99",
        ])
        assert result.exit_code == 0
