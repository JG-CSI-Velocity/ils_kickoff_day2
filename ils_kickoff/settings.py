"""Configuration system using Pydantic BaseModel + yaml.safe_load."""

import logging
import re
from pathlib import Path

import yaml
from pydantic import BaseModel, field_validator, model_validator

from ils_kickoff.exceptions import ConfigError

logger = logging.getLogger(__name__)

DEFAULT_CONFIG_PATH = Path("config.yaml")

DEFAULT_NSF_BINS = [-1, 0, 6, 12, 24, 36, 48, float("inf")]
DEFAULT_NSF_LABELS = ["0", "1-6", "7-12", "13-24", "25-36", "37-48", "49+"]

DEFAULT_DEP_BINS = [-1, 0, 1, 2, 5, 10, float("inf")]
DEFAULT_DEP_LABELS = ["0", "1", "2", "3-5", "6-10", "10+"]

BRAND_COLORS = ["#1B365D", "#4A90D9", "#7BC67E", "#F5A623", "#D0021B", "#8B572A"]


class ChartConfig(BaseModel):
    """Chart rendering settings."""

    theme: str = "plotly_white"
    colors: list[str] = BRAND_COLORS.copy()
    width: int = 900
    height: int = 500
    scale: int = 3


class OutputConfig(BaseModel):
    """Which output formats to generate."""

    excel: bool = True
    powerpoint: bool = True
    html_charts: bool = False


class Settings(BaseModel):
    """Main application configuration."""

    data_file: Path
    client_id: str | None = None
    client_name: str | None = None
    output_dir: Path = Path("output/")
    outputs: OutputConfig = OutputConfig()
    charts: ChartConfig = ChartConfig()
    pptx_template: Path | None = None
    nsf_bins: list[float] = DEFAULT_NSF_BINS.copy()
    nsf_labels: list[str] = DEFAULT_NSF_LABELS.copy()
    deposit_bins: list[float] = DEFAULT_DEP_BINS.copy()
    deposit_labels: list[str] = DEFAULT_DEP_LABELS.copy()

    @field_validator("data_file")
    @classmethod
    def validate_data_file(cls, v: Path) -> Path:
        if not v.exists():
            raise ValueError(
                f"Data file not found: {v}\n"
                "Please check the path in config.yaml or pass the correct path as an argument."
            )
        suffix = v.suffix.lower()
        if suffix not in (".csv", ".xlsx", ".xls"):
            raise ValueError(
                f"Unsupported file type: {suffix}\n"
                "Supported formats: .csv, .xlsx, .xls"
            )
        return v

    @model_validator(mode="after")
    def derive_client_fields(self):
        if self.client_id is None:
            match = re.match(r"^(\d+)", self.data_file.stem)
            if match:
                self.client_id = match.group(1)
            else:
                self.client_id = "unknown"
                logger.warning(
                    "Could not extract client ID from filename '%s'. "
                    "Set client_id in config.yaml.",
                    self.data_file.name,
                )
        if self.client_name is None:
            self.client_name = f"Client {self.client_id}"
        return self

    @classmethod
    def from_yaml(cls, config_path: Path = DEFAULT_CONFIG_PATH, **cli_overrides) -> "Settings":
        """Load settings from YAML file with CLI overrides.

        Priority: CLI args > config.yaml > defaults.
        """
        raw = {}
        if config_path.exists():
            with open(config_path) as f:
                raw = yaml.safe_load(f) or {}
            logger.info("Loaded config from %s", config_path)

        # Apply CLI overrides (non-None values only)
        for key, value in cli_overrides.items():
            if value is not None:
                raw[key] = str(value) if isinstance(value, Path) else value

        try:
            return cls(**raw)
        except Exception as e:
            raise ConfigError(
                f"Configuration error: {e}\n\n"
                "Fix your config.yaml or pass the correct options on the command line.\n"
                "Example: python -m ils_kickoff data/your_file.csv"
            ) from e

    @classmethod
    def from_args(cls, data_file: Path, **kwargs) -> "Settings":
        """Create settings directly from arguments (no YAML needed)."""
        try:
            return cls(data_file=data_file, **kwargs)
        except Exception as e:
            raise ConfigError(f"Configuration error: {e}") from e
