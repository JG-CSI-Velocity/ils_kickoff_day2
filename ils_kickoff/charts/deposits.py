"""Charts for Analyses 3-4: Deposit Distribution."""

import plotly.graph_objects as go

from ils_kickoff.settings import ChartConfig


def chart_deposit_distribution(df, config: ChartConfig) -> go.Figure:
    """Horizontal bar chart of deposit distribution by bin."""
    data = df[df.iloc[:, 0] != "Grand Total"].copy()
    colors = config.colors

    fig = go.Figure()

    fig.add_trace(go.Bar(
        y=data.iloc[:, 0],
        x=data["Accounts"],
        name="Accounts",
        orientation="h",
        marker_color=colors[0],
    ))

    fig.update_layout(
        template=config.theme,
        xaxis_title="Number of Accounts",
        yaxis_title="Deposit Count Bin",
        yaxis=dict(autorange="reversed"),
        margin=dict(t=60, b=40, l=80),
    )

    return fig
