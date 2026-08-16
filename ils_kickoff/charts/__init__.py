"""Chart creation registry and dispatcher."""

import logging
from io import BytesIO

import plotly.graph_objects as go

from ils_kickoff.analyses.base import AnalysisResult
from ils_kickoff.settings import Settings, ChartConfig

from ils_kickoff.charts.account_status import (
    chart_account_status,
    chart_account_type,
)
from ils_kickoff.charts.deposits import chart_deposit_distribution
from ils_kickoff.charts.nsf import (
    chart_nsf_volume,
    chart_nsf_pay_ratio,
    chart_nsf_full_metrics,
)
from ils_kickoff.charts.od_status import chart_od_status
from ils_kickoff.charts.reg_e import chart_reg_e_summary, chart_reg_e_historical

logger = logging.getLogger(__name__)

# Maps analysis name to chart builder function.
# Each function signature: (df: pd.DataFrame, config: ChartConfig) -> go.Figure
CHART_REGISTRY: dict[str, callable] = {
    "Account Status Summary": chart_account_status,
    "Account Type (Open)": chart_account_type,
    "Personal Deposit Distribution": chart_deposit_distribution,
    "Business Deposit Distribution": chart_deposit_distribution,
    "Personal NSF Stratification": chart_nsf_volume,
    "Business NSF Stratification": chart_nsf_volume,
    "Personal NSF Pay Ratio": chart_nsf_pay_ratio,
    "Business NSF Pay Ratio": chart_nsf_pay_ratio,
    "Personal NSF Full Metrics": chart_nsf_full_metrics,
    "Business NSF Full Metrics": chart_nsf_full_metrics,
    "Personal OD Status": chart_od_status,
    "Business OD Status": chart_od_status,
    "Reg E Summary": chart_reg_e_summary,
    "OD Limit Stratification": chart_od_status,
    "Historical Reg E by Year": chart_reg_e_historical,
}


def create_charts(
    analyses: list[AnalysisResult],
    settings: Settings,
) -> dict[str, go.Figure]:
    """Build Plotly figures for all successful analyses."""
    charts = {}
    config = settings.charts

    for analysis in analyses:
        if analysis.error is not None or analysis.df.empty:
            continue

        builder = CHART_REGISTRY.get(analysis.name)
        if builder is None:
            logger.debug("No chart builder for '%s'", analysis.name)
            continue

        try:
            fig = builder(analysis.df, config)
            fig.update_layout(title_text=analysis.title)
            charts[analysis.name] = fig
        except Exception as e:
            logger.warning("Chart for '%s' failed: %s", analysis.name, e)

    return charts


def render_chart_png(fig: go.Figure, config: ChartConfig) -> bytes:
    """Render a Plotly figure to PNG bytes (no temp files)."""
    return fig.to_image(
        format="png",
        width=config.width,
        height=config.height,
        scale=config.scale,
    )
