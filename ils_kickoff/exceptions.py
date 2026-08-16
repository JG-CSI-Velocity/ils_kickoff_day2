"""Custom exception hierarchy for user-friendly error reporting."""


class ILSError(Exception):
    """Base exception for all ILS Kickoff errors."""


class ConfigError(ILSError):
    """Invalid or missing configuration."""


class DataLoadError(ILSError):
    """Failed to load or parse the data file."""


class ColumnMismatchError(DataLoadError):
    """Required columns missing from the dataset."""

    def __init__(self, missing: set[str], available: set[str]):
        self.missing = missing
        self.available = available
        missing_str = ", ".join(sorted(missing))
        super().__init__(
            f"Missing required columns: {missing_str}\n"
            f"Available columns: {', '.join(sorted(available))}"
        )


class AnalysisError(ILSError):
    """An individual analysis failed."""

    def __init__(self, analysis_name: str, cause: Exception):
        self.analysis_name = analysis_name
        self.cause = cause
        super().__init__(f"Analysis '{analysis_name}' failed: {cause}")
