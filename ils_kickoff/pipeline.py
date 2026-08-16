"""Pipeline orchestrator shared by CLI and Streamlit."""

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

import pandas as pd
import plotly.graph_objects as go

from ils_kickoff.analyses import run_all_analyses
from ils_kickoff.analyses.base import AnalysisResult
from ils_kickoff.data_loader import load_data
from ils_kickoff.settings import Settings

logger = logging.getLogger(__name__)


@dataclass
class PipelineResult:
    """Container for all pipeline outputs."""

    settings: Settings
    df: pd.DataFrame
    analyses: list[AnalysisResult] = field(default_factory=list)
    charts: dict[str, go.Figure] = field(default_factory=dict)


def run_pipeline(
    settings: Settings,
    on_progress: Optional[Callable[[int, int, str], None]] = None,
) -> PipelineResult:
    """Execute the full analysis pipeline: load -> analyze -> chart.

    Args:
        settings: Application configuration.
        on_progress: Optional callback(step, total, message) for UI progress.
    """
    # Step 1: Load data
    if on_progress:
        on_progress(0, 3, "Loading data...")
    df = load_data(settings)

    # Step 2: Run analyses
    if on_progress:
        on_progress(1, 3, "Running analyses...")
    analyses = run_all_analyses(df, settings, on_progress=None)
    successful = [a for a in analyses if a.error is None]
    failed = [a for a in analyses if a.error is not None]
    if failed:
        for a in failed:
            logger.warning("Skipped: %s (%s)", a.name, a.error)
    logger.info("%d/%d analyses completed", len(successful), len(analyses))

    # Step 3: Build charts
    if on_progress:
        on_progress(2, 3, "Building charts...")
    charts = {}
    try:
        from ils_kickoff.charts import create_charts
        charts = create_charts(analyses, settings)
        logger.info("Built %d charts", len(charts))
    except ImportError:
        logger.warning("Charts module not available; skipping chart generation.")
    except Exception as e:
        logger.error("Chart generation failed: %s", e, exc_info=True)

    return PipelineResult(settings=settings, df=df, analyses=analyses, charts=charts)


def export_outputs(result: PipelineResult) -> list[Path]:
    """Export pipeline results to configured output formats.

    Returns list of generated file paths.
    """
    settings = result.settings
    settings.output_dir.mkdir(parents=True, exist_ok=True)

    generated: list[Path] = []
    from datetime import datetime

    date_str = datetime.now().strftime("%Y%m%d")
    client_id = settings.client_id or "unknown"

    if settings.outputs.excel:
        try:
            from ils_kickoff.excel_report import write_excel_report

            path = settings.output_dir / f"{client_id}_ILS_Kickoff_Report_{date_str}.xlsx"
            write_excel_report(result, path)
            generated.append(path)
            logger.info("Excel report: %s", path)
        except Exception as e:
            logger.error("Excel report failed: %s", e, exc_info=True)

    if settings.outputs.powerpoint:
        try:
            from ils_kickoff.pptx_report import write_pptx_report

            path = settings.output_dir / f"{client_id}_ILS_Kickoff_Presentation_{date_str}.pptx"
            write_pptx_report(result, path)
            generated.append(path)
            logger.info("PowerPoint report: %s", path)
        except Exception as e:
            logger.error("PowerPoint report failed: %s", e, exc_info=True)

    return generated
