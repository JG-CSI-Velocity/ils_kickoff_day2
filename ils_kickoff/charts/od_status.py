"""Charts for Analyses 11-12, 14: OD Status and OD Limit Stratification."""

import plotly.graph_objects as go

from ils_kickoff.settings import ChartConfig


def chart_od_status(df, config: ChartConfig) -> go.Figure:
    """Bar chart of accounts and pay ratio by OD Status (or OD Limit)."""
    data = df[df.iloc[:, 0] != "Grand Total"].copy()
    colors = config.colors
    group_col = data.columns[0]

    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=data[group_col],
        y=data["# of Accounts"],
        name="# of Accounts",
        marker_color=colors[0],
        yaxis="y",
    ))

    ratio_col = "Pay Ratio" if "Pay Ratio" in data.columns else None
    if ratio_col:
        fig.add_trace(go.Scatter(
            x=data[group_col],
            y=data[ratio_col],
            name="Pay Ratio",
            mode="lines+markers",
            marker=dict(color=colors[1], size=8),
            line=dict(color=colors[1], width=2),
            yaxis="y2",
        ))

    fig.update_layout(
        template=config.theme,
        xaxis_title=group_col,
        yaxis=dict(title="# of Accounts", side="left"),
        yaxis2=dict(title="Pay Ratio", side="right", overlaying="y", range=[0, 1.1]),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=40),
    )

    return fig
