"""Typer CLI entry point for ILS Kickoff."""

import logging
import sys
from pathlib import Path

import typer

from ils_kickoff.exceptions import ILSError
from ils_kickoff.settings import Settings, DEFAULT_CONFIG_PATH

app = typer.Typer(
    name="ils-kickoff",
    help="ILS Kickoff banking data analysis and report generation tool.",
    add_completion=False,
)


def setup_logging(verbose: bool) -> None:
    """Configure logging for CLI output."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(message)s" if not verbose else "%(levelname)s %(name)s: %(message)s",
        stream=sys.stderr,
    )


@app.command()
def main(
    data_file: Path = typer.Argument(
        None,
        help="Path to the client CSV/Excel file. Overrides config.yaml.",
        exists=False,
    ),
    config: Path = typer.Option(
        DEFAULT_CONFIG_PATH,
        "--config", "-c",
        help="Path to config.yaml file.",
    ),
    output_dir: Path = typer.Option(
        None,
        "--output", "-o",
        help="Output directory for reports.",
    ),
    client_id: str = typer.Option(
        None,
        "--client-id",
        help="Client identifier (auto-extracted from filename if omitted).",
    ),
    client_name: str = typer.Option(
        None,
        "--client-name",
        help="Client name for report titles.",
    ),
    verbose: bool = typer.Option(
        False,
        "--verbose", "-v",
        help="Enable debug logging.",
    ),
) -> None:
    """Run ILS Kickoff analysis and generate reports.

    Usage:
        python -m ils_kickoff data/client_file.csv
        python -m ils_kickoff --config config.yaml
        python -m ils_kickoff data/client.csv --output ./reports/ --client-id 1774
    """
    setup_logging(verbose)
    logger = logging.getLogger(__name__)

    try:
        # Build overrides from CLI args
        overrides = {}
        if data_file is not None:
            overrides["data_file"] = data_file
        if output_dir is not None:
            overrides["output_dir"] = output_dir
        if client_id is not None:
            overrides["client_id"] = client_id
        if client_name is not None:
            overrides["client_name"] = client_name

        settings = Settings.from_yaml(config_path=config, **overrides)
        logger.info("Client: %s (%s)", settings.client_name, settings.client_id)
        logger.info("Data: %s", settings.data_file)
        logger.info("Output: %s", settings.output_dir)

        from ils_kickoff.pipeline import run_pipeline, export_outputs

        result = run_pipeline(settings)

        successful = [a for a in result.analyses if a.error is None]
        logger.info("Analyses completed: %d/15", len(successful))

        generated = export_outputs(result)

        if generated:
            logger.info("")
            logger.info("Generated reports:")
            for path in generated:
                logger.info("  %s", path)
        else:
            logger.warning("No reports were generated.")

    except ILSError as e:
        logger.error(str(e))
        raise typer.Exit(code=1) from None
    except Exception as e:
        logger.error("Unexpected error: %s", e)
        if verbose:
            import traceback
            traceback.print_exc()
        raise typer.Exit(code=1) from None


if __name__ == "__main__":
    app()
